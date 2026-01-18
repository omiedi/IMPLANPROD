Attribute VB_Name = "REMEVEN"
Public CARGAMEN%
'
Dim CORRESP$, ABREMES$, DES%, HAS%, DESFE$, HASFE$
Dim TIS$(10)
Dim COPRO$(128)
Dim CUEV%(8), IMXX#(8)
Dim CUEIVA1%, CUEADU1%, CUERET1%, CUEDEUV1%, YACUE%
Dim CUEIVA$, CUEADU$, CUERET$
Dim GRUCOD%(1024), GRUCLI%(1024), CUECON%(1024), CARUCO%
Dim GRUCOVE%(256), CAGRUCOVE%
Dim CUECOD%(32768)
Dim GRUCOC%(256), CAGRUCOCLI%
Dim GRCOCLI%(32768), CUEMA%(32768), RACLI$(32768), EXICLI%(32768)
Dim SALINI#(32767), SUMFAC#(32767), SUMCOB#(32767)
Dim SUMAFAC#(32)
'
Dim NCV%(8192), TII%(8192), NROV&(8192), FFII%(8192) ' dimensionamiento
Dim PRIMOSTO&(8192), ULTMOSTO&(8192)
Dim IFAC#(8192), IDET#(8192), SIFAC%(8192)
Dim NRECI&(8192), REGFA$(8192)
Dim X0 As String * 96
'
Dim PORCO(256), TOCOMI#(256), TOCOB#(256)
Dim SUPED#(256), COMPED#(256), SUCOB#(256), COMCOB#(256)
'
Public YALIQUI%
Public MODOCLAVE%     ' MODALIDAD DE CLASIFICACION DE VENTAS
Public MODOCALVE%     ' MODALIDAD CALCULO DE VENTAS
'

Sub CALREVEN()
     '
     Call COPYSTRU("MOVISTO", "MOVISVIE")
     EPAC$ = TRIM$(UCase$(EMPREAC$))
     '
     If RECLEN("SUIVAVE") <> 256 Then
       Call MENSERR(24, "Longitud de Registro Incorrecta - Archivo SUIVAVE.")
       Exit Sub
     End If
     '
     EPAC$ = TRIM$(UCase$(EMPREAC$))
     MODOSABO% = 0: CUEPIBRU% = 0
     If InStr(EPAC$, "SABO") > 0 Or InStr(EPAC$, "DOSIVA") > 0 Or InStr(EPAC$, "FERVI") > 0 Then
       MODOSABO% = 1
       CUEPER$ = CONTROL$("PERIBRU", "CUECONTA")
       CUEPIBRU% = CUECOINT%(CUEPER$)
     End If
     '
     Dim TAIVA(2)
     TAIVA(0) = 21: TAIVA(1) = 10.5
     For U& = 1 To 2
       If U& <= ULTREG&("GRUPIVA") Then
         X$ = REGLEIDO$("GRUPIVA", U&)
         TAIVA(U& - 1) = CVS(Mid$(X$, 13, 4))
       End If
     Next U&
     Call CIERRARCH("GRUPIVA")
     '
     'REMEVENT.LABEL2(20).Caption = "Tasa " + TRIM$(Str$(TAIVA(0))) + " %"
     'REMEVENT.LABEL2(11).Caption = "Tasa " + TRIM$(Str$(TAIVA(1))) + " %"
     '
     For XI% = 0 To 9: REMEVENT.Check1(XI%).Value = 0: Next XI%
     '
10   REMEVENT.Check1(0).Value = 2
     Call VALICUEIVA
     Screen.MousePointer = 11
     If YACUE% <> 1 Then
        OPCLAVEN.Show 1
        Screen.MousePointer = 1
        GoTo 999
     End If
     REMEVENT.Check1(0).Value = 1
     '
20   Call ARMAMES
     '
     Call BACKUPMEN(1, HAYERRO%)
     If HAYERRO% > 0 Then
        Screen.MousePointer = 1
        GoTo 999
     End If
     '
30   REMEVENT.Check1(1).Value = 2
     Call CARCUECOD
     If YACUE% <> 1 Then
        Screen.MousePointer = 1
        GoTo 999
     End If
     REMEVENT.Check1(1).Value = 1
     '
40   REMEVENT.Check1(2).Value = 2
     Call CARCUECLI
     If YACUE% <> 1 Then
        Screen.MousePointer = 1
        GoTo 999
     End If
     REMEVENT.Check1(2).Value = 1
     '
     Call BLOQARCH("CONCIVEN")
     Call BLANARCH("CONCIVEN")
     For XI% = 0 To 32: SUMAFAC#(XI%) = 0: Next XI%
     REGDETACC& = 0
     RESCUDEU.LINDICE.Clear
     '
     For IL& = 0 To 32767
        SALINI#(IL&) = 0
        SUMFAC#(IL&) = 0
        SUMCOB#(IL&) = 0
     Next IL&
     '
     For XII% = 0 To 7
        REMEVENT.Text1(XII%).TEXT = ""
        REMEVENT.Text1(XII%).Refresh
     Next XII%
     For XII% = 11 To 25
        REMEVENT.Text1(XII%).TEXT = ""
        REMEVENT.Text1(XII%).Refresh
     Next XII%
     '
     GoSub 10000   ' arrastre de saldos segun cobranza
     GoSub 11000   ' seleccionar facturas
     GoSub 11100   ' cargar facturas a vector busqueda
     GoSub 12000   ' lectura archivo MOVISTO
     GoSub 12200   ' grabar COEF# a CONCIVEN
     GoSub 11200   ' verificar con facturacion
     GoSub 13000   ' imputacion automatica
     GoSub 13900   ' resumenes mensuales por pieza y cliente
     GoSub 14000   ' GENERACION DE LISTADORES
     '
     Call CIERRARCH("*.*")
     Call BACKUPMEN(2, HAYERRO%)
     If HAYERRO% = 0 Then
        GoSub 15000   ' GRABAR RESUMEN MENSUAL
     End If
     '
     Close #N2%
     Call CIERRARCH("*.*")
     Screen.MousePointer = 1
     '
999  Call CIERRARCH("*.*")
     Call LIBARCH("*.*")
     Exit Sub



        '
10000 ' Revision del Archivo de Cobranza ... **********************
        '
        REMEVENT.Check1(3).Value = 2
        FIN& = ULTREG&("COBRAN")
        CAFAC% = 0
        REGCOBCLI& = 0
        '
        For IL& = 1 To FIN&
           Call TRACE(20, IL&, FIN&)
           X$ = REGLEIDO$("COBRAN", IL&)
           NC% = CVI(Left$(X$, 2))
           If NC% > 0 And NC% <= 32767 Then
              FF% = CVI(Mid$(X$, 7, 2))
              TI% = 4: SIG% = (-1)
              NRO& = CVS(Mid$(X$, 3, 4))
              TOT# = (Int(100 * (CVD(Mid$(X$, 15, 8))) + 0.5)) / 100
              '
' AA1 = CVI(Left$(x$, 2))
' AA2 = CVS(Mid$(x$, 3, 4))
' AA3 = CVI(Mid$(x$, 7, 2))
' AA5 = Asc(Mid$(x$, 9, 1))
' AA6 = Asc(Mid$(x$, 10, 1))
' AA4 = CVD(Mid$(x$, 15, 8))
' If IL& = 4987 Or IL& = 4988 Then
'   AAA = BBB
' End If
              If TOT# >= 0.005 Then
                 If EXICLI%(NC%) < 1 Then
                    Call MENSERR(24, "Cobranza Registrada sobre Cuenta '" + TRIM$(Str$(NC%)) + "' Inexistente.")
                 End If
              End If
              '
              If FF% < DES% Then
                  SUMAFAC#(0) = SUMAFAC#(0) + TOT# * SIG%
                  SUMAFAC#(7) = SUMAFAC#(7) + TOT# * SIG%
                  SALINI#(NC%) = SALINI#(NC%) + TOT# * SIG%
                ElseIf FF% <= HAS% Then
                  SUMAFAC#(5) = SUMAFAC#(5) + TOT#
                  SUMAFAC#(7) = SUMAFAC#(7) + TOT# * SIG%
                  SUMCOB#(NC%) = SUMCOB#(NC%) + TOT#
                  '
                  Y$ = REGBLAN$("DETACCD")
                  Call REPLA(Y$, Chr$(TI%), 1, 1)
                  Call REPLA(Y$, MKI$(FF%), 3, 2)
                  Call REPLA(Y$, MKS$(NRO&), 5, 4)
                  Call REPLA(Y$, MKI$(NC%), 9, 2)
                  Call REPLA(Y$, MKD$(TOT#), 11, 8)
                  Call REPLA(Y$, MKD$(TOT#), 51, 8)
                  '
                  REGDETACC& = REGDETACC& + 1
                  Call GRAREG("DETACCD", Y$, REGDETACC&)
                  FF% = CVI(Mid$(Y$, 3, 2))
                  Z$ = CSTRING$(MKI$(FF%), 1, 6, 0, 2) + CSTRING$(MKS$(REGDETACC&), 4, 8, 0, 2)
                  RESCUDEU.LINDICE.AddItem Z$
                  '
                  Y$ = REGBLAN$("RECOBCLI")
                  Call REPLA(Y$, MKI$(NC%), 1, 2)
                  Call REPLA(Y$, MKI$(FF%), 3, 2)
                  Call REPLA(Y$, MKS$(NRO&), 5, 4)
                  Call REPLA(Y$, MKD$(TOT#), 25, 8)
                  REGCOBCLI& = REGCOBCLI& + 1
                  Call GRAREG("RECOBCLI", Y$, REGCOBCLI&)
                  '
              End If
           End If
10049   Next IL&
        '
        Call SETULTREG("RECOBCLI", REGCOBCLI&)
        Call CIERRARCH("RECOBCLI")
        REMEVENT.Check1(3).Value = 1
        Return
        '

11000 ' Revision del Archivo de Facturacion ... **********************
        '
        REMEVENT.Check1(4).Value = 2
        FIN& = ULTREG&("FACTUR")
        CAFAC% = 0
        CARESUVE& = 0
        '
        FFANT% = 0
        For IL& = 1 To FIN&
           Call TRACE(20, IL&, FIN&)
           X$ = REGLEIDO$("FACTUR", IL&)
           NC% = CVI(Left$(X$, 2))
           FF% = CVI(Mid$(X$, 7, 2))
           If FF% = 0 Then
              FF% = FFANT%
              Call REPLA(X$, MKI$(FFANT%), 7, 2)
              Call GRAREG("FACTUR", X$, IL&)
           End If
           FFANT% = FF%
           '
           If NC% < 1 Or NC% > 32767 Then GoTo 11049
           '
           TI% = Asc(Mid$(X$, 13, 1)): If TI% < 0 Or TI% > 7 Then TI% = 2 ' default factura
           VA% = Asc(Mid$(X$, 14, 1)): If VA% < 1 Or VA% > 3 Then VA% = 1 ' default factura
           '
11002      NETO1# = (Int(100 * (CVD(Mid$(X$, 15, 8))) + 0.5)) / 100
           IVAD1# = (Int(100 * CVD(Mid$(X$, 23, 8)) + 0.5)) / 100
           RNI# = (Int(100 * CVD(Mid$(X$, 31, 8)) + 0.5)) / 100
           NETO2# = (Int(100 * CVD(Mid$(X$, 49, 8)) + 0.5)) / 100
           IVAD2# = (Int(100 * CVD(Mid$(X$, 57, 8)) + 0.5)) / 100
           IPER = 0
           '
           If MODOSABO% = 1 Then
             IPER = IVAD2#
             IVAD2# = 0
           End If
           '
           If Abs(NETO2#) > 0.1 Then
             If Abs(100 * IVAD2# / NETO2# - TAIVA(0)) < 0.2 Then
               NETO1# = NETO1# + NETO2#
               Call REPLA(X$, MKD$(NETO1#), 15, 8)
               NETO2# = 0
               Call REPLA(X$, MKD$(NETO2#), 49, 8)
               IVAD1# = IVAD1# + IVAD2#
               Call REPLA(X$, MKD$(IVAD1#), 23, 8)
               IVAD2# = 0
               Call REPLA(X$, MKD$(IVAD2#), 57, 8)
               Call GRAREG("FACTUR", X$, IL&)
             End If
           End If
           '
           NETO# = NETO1# + NETO2#
           IVAD# = IVAD1# + IVAD2#
           TOT# = NETO# + IVAD# + RNI# + CDbl(IPER) + IRET#
           '
           SIG% = 1
           If TI% > 3 Then
               NETO1# = Abs(NETO1#)
               NETO2# = Abs(NETO2#)
               NETO# = Abs(NETO#)
               IVAD1# = Abs(IVAD1#)
               IVAD2# = Abs(IVAD2#)
               IVAD# = Abs(IVAD#)
               RNI# = Abs(RNI#)
               IPER = Abs(IPER)
               TOT# = Abs(TOT#)
               SIG% = (-1)
           End If
           '
           If FF% < DES% Then
               SUMAFAC#(0) = SUMAFAC#(0) + TOT# * SIG%
               SUMAFAC#(7) = SUMAFAC#(7) + TOT# * SIG%
               If NC% > 0 Then SALINI#(NC%) = SALINI#(NC%) + TOT# * SIG%
           End If
           '
           If FF% >= DES% Then
             If FF% <= HAS% Then
                NCVI% = NC%
                If NCVI% <= 0 Then NCVI% = 0
                '
                FEMI% = FF%
                If NCVI% = 0 Then
                    RRCC$ = Space$(46) + Chr$(0)
                    Call REPLA(RRCC$, "Clientes de Mostrador", 1, 30)
                    Call REPLA(RRCC$, "NO POSEE", 31, 16)
                    Call REPLA(RRCC$, Chr$(0), 47, 1)
                    RACLI$(NCVI%) = RRCC$
                    GoTo 11009
                End If
                '
                If RACLI$(NCVI%) = "" Then
                    RRCC$ = Space$(46) + Chr$(0)
                    Call REPLA(RRCC$, RASOCLI$(NCVI%), 1, 30)
                    Call REPLA(RRCC$, CUITCLI$(NCVI%), 31, 16)
                    Call REPLA(RRCC$, Chr$(PIVACLI%(NCVI%)), 47, 1)
                    RACLI$(NCVI%) = RRCC$
                End If
                '
11009           If GRCOCLI%(NCVI%) < 1 Then
                  If REGICLI&(NCVI%) > 0 Then
11010               Screen.MousePointer = 1
                    If VDEF$ = "" Then
                        VDEF$ = MKI$(NCVI%) + MKI$(0)
                      Else
                        Call REPLA(VDEF$, MKI$(NCVI%), 1, 2)
                    End If
                    OBJE$ = OBJPLA$("CARCUECOCLI", VDEF$)
                    If OBJE$ = "" Then GoTo 999
                    If CVI(Mid$(OBJE$, 3, 2)) < 1 Then GoTo 11010
                    Screen.MousePointer = 11
                    cueco$ = Mid$(OBJE$, 3, 2)
                    If CVI(cueco$) < 1 Then GoTo 11010
                    For U& = 1 To ULTREG&("GRUCOCLI")
                        XK$ = REGLEIDO$("GRUCOCLI", U&)
                        If Left$(XK$, 2) = cueco$ Then
                            CUE1% = CVI(Left$(XK$, 2))
                            GoTo 11020
                        End If
                    Next U&
                    GoTo 11010
                    '
11020               If CUE1% < 1 Then GoTo 11010
                    GRCOCLI%(NCVI%) = CUE1%
                    VDEF$ = OBJE$
                    REGCLI& = REGICLI&(NCVI%)
                    RECLI2$ = REGLEIDO$("DEUDOR2", REGCLI&)
                    Call REPLA(RECLI2$, MKI$(CUE1%), 125, 2)
                    Call GRAREG("DEUDOR2", RECLI2$, REGCLI&)
                    '
                  End If
                End If
                '
                If CUEMA%(NCVI%) < 1 Then
                  If REGICLI&(NCVI%) > 0 Then
11030               Screen.MousePointer = 1
                    If VDEF1$ = "" Then
                        VDEF1$ = MKI$(NCVI%) + Space$(12)
                      Else
                        Call REPLA(VDEF1$, MKI$(NCVI%), 1, 2)
                    End If
                    OBJE$ = OBJPLA$("CARCUEMACLI", VDEF1$)
                    If OBJE$ = "" Then GoTo 999
                    Screen.MousePointer = 11
                    cueco$ = Mid$(OBJE$, 3, 12)
                    If TRIM$(cueco$) = "" Then GoTo 11010
                    For U& = 1 To ULTREG&("ECUECON")
                        XJ$ = REGLEIDO$("ECUECON", U&)
                        If Left$(XJ$, 12) = cueco$ Then
                            CUE1% = CVI(Mid$(XJ$, 41, 2))
                            GoTo 11040
                        End If
                    Next U&
                    GoTo 11030
                    '
11040               If CUE1% < 1 Then GoTo 11030
                    CUEMA%(NCVI%) = CUE1%
                    VDEF1$ = OBJE$
                    REGCLI& = REGICLI&(NCVI%)
                    RECLI2$ = REGLEIDO$("DEUDOR2", REGCLI&)
                    Call REPLA(RECLI2$, MKI$(CUE1%), 127, 2)
                    Call GRAREG("DEUDOR2", RECLI2$, REGCLI&)
                    '
                  End If
                End If
                '
11045           TIXX$ = TIS$(TI%)
                NRO& = CVS(Mid$(X$, 3, 4))
If NRO& = 2048 Then
AAA = BBB
End If
                NUSUCUVE% = CVI(Mid$(X$, 47, 2))
                If TI% = 0 Then
                    If NRO& = 0 Then
                        NRO& = IL&
                        Call REPLA(X$, MKS$(NRO&), 3, 4)
                        Call GRAREG("FACTUR", X$, IL&)
                    End If
                End If
                SUTT# = SUTT# + TOT# * SIG%
                COEF# = 1
                '
                Y$ = String$(256, 0)
                Call REPLA(Y$, TIXX$, 1, 2)
                Call REPLA(Y$, MKS$(NRO&), 3, 4)
                Call REPLA(Y$, MKI$(NC%), 7, 2)
                Call REPLA(Y$, MKI$(FEMI%), 9, 2)
                Call REPLA(Y$, MKI$(NUSUCUVE%), 11, 2)
                Call REPLA(Y$, String$(80, 0), 17, 80)
                Call REPLA(Y$, Chr$(TI%) + Chr$(VA%), 97, 2)
                Call REPLA(Y$, MKD$(NETO1# * SIG%), 99, 8)
                Call REPLA(Y$, MKD$(IVAD1# * SIG%), 107, 8)
                Call REPLA(Y$, MKD$(RNI# * SIG%), 115, 8)
                Call REPLA(Y$, MKD$(NETO2# * SIG%), 123, 8)
                Call REPLA(Y$, MKD$(IVAD2# * SIG%), 131, 8)
                Call REPLA(Y$, MKD$(TOT# * SIG%), 139, 8)
                Call REPLA(Y$, MKD$(COEF#), 147, 8)
                Call REPLA(Y$, MKI$(0), 155, 2)
' AQUI AGREGUE SIGNO
                Call REPLA(Y$, MKS$(IPER * SIG%), 157, 4)
                Call REGAPP("CONCIVEN", Y$)
                '
                If TI% > 0 Then
                   If TI% <= 6 Then
                      If NC% >= 0 Then
                            CLI$ = RASOCLI$(NC%)
                        ElseIf NC% < 0 Then
                            RECUI& = Abs(NC%)
                            If RECUI& <= ULTREG&("CUITS") Then
                               RCUX$ = REGLEIDO$("CUITS", RECUI&)
                               CLI$ = Left$(RCUX$, 30)
                            End If
                      End If
                      '
                      YR$ = REGBLAN$("RESUVEN")
                      Call REPLA$(YR$, MKI$(NC%), 1, 2)
                      Call REPLA$(YR$, CLI$, 3, 30)
                      Call REPLA(YR$, TIXX$ + AJUSTD$(Str$(NRO&), 6), 33, 8)
                      Call REPLA(YR$, MKI$(FEMI%), 41, 2)
                      Call REPLA(YR$, MKD$(NETO# * SIG%), 43, 8)
                      CARESUVE& = CARESUVE& + 1
                      Call GRAREG("RESUVEN", YR$, CARESUVE&)
                   End If
                End If
                '
                CAFAC% = CAFAC% + 1
                '
                Select Case TI%
                  Case 0
                    SUMAFAC#(6) = SUMAFAC#(6) + TOT# * SIG%
                  Case 1
                    SUMAFAC#(1) = SUMAFAC#(1) + TOT# * SIG%
                  Case 2
                    SUMAFAC#(1) = SUMAFAC#(1) + TOT# * SIG%
                  Case 3
                    SUMAFAC#(2) = SUMAFAC#(2) + TOT# * SIG%
                  Case 4
                    SUMAFAC#(5) = SUMAFAC#(5) + Abs(TOT#)
                  Case 5
                    SUMAFAC#(3) = SUMAFAC#(3) + Abs(TOT#)
                  Case 6
                    SUMAFAC#(3) = SUMAFAC#(3) + Abs(TOT#)
                  Case 7
                    SUMAFAC#(6) = SUMAFAC#(6) + TOT# * SIG%
                End Select
                SUMAFAC#(7) = SUMAFAC#(7) + TOT# * SIG%
                SUMAFAC#(4) = SUMAFAC#(1) + SUMAFAC#(2) - SUMAFAC#(3)
                '
                NCII% = NC%: If NCII% < 0 Then NCII% = 0
                SUMFAC#(NCII%) = SUMFAC#(NCII%) + TOT# * SIG%
                If NCII% <= 0 Then
                    SUMCOB#(NCII%) = SUMCOB#(NCII%) + TOT# * SIG%
                End If
                '
                Y$ = REGBLAN$("DETACCD")
                Call REPLA(Y$, Chr$(TI%), 1, 1)
                Call REPLA(Y$, MKI$(FEMI%), 3, 2)
                Call REPLA(Y$, MKS$(NRO&), 5, 4)
                Call REPLA(Y$, MKI$(NC%), 9, 2)
                Call REPLA(Y$, MKD$(NETO#), 11, 8)
                Call REPLA(Y$, MKD$(IVAD#), 19, 8)
                Call REPLA(Y$, MKD$(RNI#), 27, 8)
                Call REPLA(Y$, MKD$(CDbl(IPER)), 35, 8)
                Call REPLA(Y$, MKD$(IRET#), 43, 8)
                Call REPLA(Y$, MKD$(TOT#), 51, 8)
                '
                REGDETACC& = REGDETACC& + 1
                Call GRAREG("DETACCD", Y$, REGDETACC&)
                FEMI% = CVI(Mid$(Y$, 3, 2))
                Z$ = CSTRING$(MKI$(FEMI%), 1, 6, 0, 2) + CSTRING$(MKS$(REGDETACC&), 4, 8, 0, 2)
                RESCUDEU.LINDICE.AddItem Z$
                '
                For XII% = 0 To 7
                   REMEVENT.Text1(XII%).TEXT = CSTRING$(MKD$(SUMAFAC#(XII%)), 4, 11, 2, 2)
                   REMEVENT.Text1(XII%).Refresh
                Next XII%
                '
             End If
             '
           End If
           '
11049   Next IL&
        '
        For XII% = 0 To 7
           REMEVENT.Text1(XII%).TEXT = CSTRING$(MKD$(SUMAFAC#(XII%)), 4, 11, 2, 2)
           REMEVENT.Text1(XII%).Refresh
        Next XII%
        '
        Call SETULTREG("DETACCD", REGDETACC&)
        Call SETULTREG("RESUVEN", CARESUVE&)
        '
        REGFAC& = 0
        REGDEB& = 0
        REGCRE& = 0
        REGCOB& = 0
        REGOTR& = 0
        REGRES& = 0
        REGPRV& = 0
        REGPRX& = 0
        KPX% = 0
        '
        FIN& = REGDETACC&
        '
        FIN& = RESCUDEU.LINDICE.ListCount
        For U& = 1 To FIN&
            Call TRACE(20, U&, FIN&)
            RESCUDEU.LINDICE.ListIndex = U& - 1
            REG& = Val(Mid$(RESCUDEU.LINDICE.TEXT, 7, 8))
            Y$ = REGLEIDO$("DETACCD", REG&)
            TI% = Asc(Left$(Y$, 1))
            NC% = CVI(Mid$(Y$, 9, 2))
            '
            Select Case TI%
               Case 0
                  REGOTR& = REGOTR& + 1
                  Call GRAREG("DETOTRV", Y$, REGOTR&)
               Case 1
                  REGFAC& = REGFAC& + 1
                  Call GRAREG("DETFACV", Y$, REGFAC&)
               Case 2
                  REGFAC& = REGFAC& + 1
                  Call GRAREG("DETFACV", Y$, REGFAC&)
               Case 3
                  REGDEB& = REGDEB& + 1
                  Call GRAREG("DETDEBV", Y$, REGDEB&)
               Case 4
                  REGCOB& = REGCOB& + 1
                  Call GRAREG("DETCOBV", Y$, REGCOB&)
               Case 5
                  REGCRE& = REGCRE& + 1
                  Call GRAREG("DETCREV", Y$, REGCRE&)
               Case 6
                  REGCRE& = REGCRE& + 1
                  Call GRAREG("DETCREV", Y$, REGCRE&)
               Case 7
                  REGOTR& = REGOTR& + 1
                  Call GRAREG("DETOTRV", Y$, REGOTR&)
            End Select
            '
            If TI% > 0 Then
               If TI% < 7 Then
                  If TI% <> 4 Then
                    COPROVI$ = AJUSTI$(PROVCLI$(NC%), 4)
                    For KPP% = 1 To KPX%
                       If Left$(COPRO$(KPP%), 4) = COPROVI$ Then GoTo 11050
                    Next KPP%
                    KPX% = KPX% + 1
                    COPRO$(KPX%) = COPROVI$
                    KPP% = KPX%
                    '
11050               COPRO$(KPP%) = COPRO$(KPP%) + MKS$(REG&)
                    '
                  End If
               End If
            End If
            '
        Next U&
        '
        For KP1% = 1 To KPX%
            For KP2% = KP1% + 1 To KPX%
               If COPRO$(KP1%) > COPRO$(KP2%) Then
                  COPRA$ = COPRO$(KP1%)
                  COPRO$(KP1%) = COPRO$(KP2%)
                  COPRO$(KP2%) = COPRA$
               End If
            Next KP2%
            '
            COPROVI$ = Left$(COPRO$(KP1%), 1)
            For U& = 5 To Len(COPRO$(KP1%)) Step 4
               '
               REG& = CVS(Mid$(COPRO$(KP1%), U&, 4))
               Y$ = REGLEIDO$("DETACCD", REG&)
               TI% = Asc(Left$(Y$, 1))
               FF% = CVI(Mid$(Y$, 3, 2))
               NRO& = CVS(Mid$(Y$, 5, 4))
               NC% = CVI(Mid$(Y$, 9, 2))
               INET# = Abs(CVD(Mid$(Y$, 11, 8)))
               IIVA# = Abs(CVD(Mid$(Y$, 19, 8)))
               SIG% = 1: If TI% > 3 Then SIG% = (-1)
               '
               If Abs(INET#) >= 0.005 Then
                  If TI% > 0 Then
                    If TI% < 7 Then
                      If TI% <> 4 Then
                        TIX$ = TIS$(TI%)
                        NF& = NRO&: While NF& > 999999: NF& = NF& - 100000: Wend
                        NFX$ = TRIM$(Str$(NF&))
                        While Len(NFX$) < 8
                            NFX$ = " " + NFX$
                        Wend
                        '
                        Z$ = REGBLAN$("VEPROVIN")
                        Call REPLA(Z$, COPROVI$, 1, 1)
                        Call REPLA(Z$, MKI$(FF%), 2, 2)
                        Call REPLA(Z$, TIX$ + NFX$, 4, 10)
                        Call REPLA(Z$, MKI$(NC%), 14, 2)
                        Call REPLA(Z$, MKD$(INET# * SIG%), 16, 8)
                        '
                        If Abs(IIVA#) >= 0.005 Then
                            REGPRV& = REGPRV& + 1
                            Call GRAREG("VEPROVIN", Z$, REGPRV&)
                          Else
                            REGPRX& = REGPRX& + 1
                            Call GRAREG("VEPROVIX", Z$, REGPRX&)
                        End If
                        '
                      End If
                    End If
                  End If
               End If
               '
            Next U&
            '
        Next KP1%
        '
        Call SETULTREG("DETFACV", REGFAC&)
        Call SETULTREG("DETDEBV", REGDEB&)
        Call SETULTREG("DETCREV", REGCRE&)
        Call SETULTREG("DETCOBV", REGCOB&)
        Call SETULTREG("DETOTRV", REGOTR&)
        Call SETULTREG("VEPROVIN", REGPRV&)
        Call SETULTREG("VEPROVIX", REGPRX&)
        '
        For NC% = 1 To 32767
          If Abs(SALINI#(NC%)) >= 0.005 Or Abs(SUMFAC#(NC%)) >= 0.005 Or Abs(SUMCOB#(NC%)) >= 0.005 Then
            If ECOARCH$("DEUDOR1", MKI$(NC%)) = "" Then
              Call MENSERR(24, "Movimiento Sobre Cliente" + Str$(NC%) + " Inexistente.")
            End If
          End If
          If NC% >= 32767 Then GoTo 11060
        Next NC%
        '
11060   For U& = 1 To ULTREG&("DEUDOR1")
            NC% = CVI(Left$(REGLEIDO$("DEUDOR1", U&), 2))
            If NC% > 0 Then
               If Abs(SALINI#(NC%)) >= 0.005 Or Abs(SUMFAC#(NC%)) >= 0.005 Or Abs(SUMCOB#(NC%)) >= 0.005 Then
                   X$ = REGBLAN$("RESUCCD")
                   Call REPLA(X$, MKI$(NC%), 1, 2)
                   Call REPLA(X$, MKD$(SALINI#(NC%)), 3, 8)
                   Call REPLA(X$, MKD$(SUMFAC#(NC%)), 11, 8)
                   Call REPLA(X$, MKD$(SUMCOB#(NC%)), 19, 8)
                   Call REPLA(X$, MKD$(SALINI#(NC%) + SUMFAC#(NC%) - SUMCOB#(NC%)), 27, 8)
                   Call REPLA(X$, MKD$(SUMFAC#(NC%) - SUMCOB#(NC%)), 35, 8)
                   '
                   REGRES& = REGRES& + 1
                   Call GRAREG("RESUCCD", X$, REGRES&)
                   '
                End If
            End If
        Next U&
        Call SETULTREG("RESUCCD", REGRES&)
        '
        Call CIERRARCH("*.*")
        REMEVENT.Check1(4).Value = 1
        '
        If CAFAC% < 1 Then
           Screen.MousePointer = 1
           Call COMUNI("ADVERTENCIA:\No hay Facturas Registradas\en el Periodo Elegido.")
        End If
        '
        Return
        '
'*****************************************************************************
        '
11100   FIN& = ULTREG&("CONCIVEN")
        REMEVENT.Check1(5).Value = 2
        '
        For IL& = 1 To FIN&
           Call TRACE(20, IL&, FIN&)
           X$ = REGLEIDO$("CONCIVEN", IL&)
           '
           JI% = IL&
           NCV%(JI%) = CVI(Mid$(X$, 7, 2))
           FFII%(JI%) = CVI(Mid$(X$, 9, 2))
           TII%(JI%) = Asc(Mid$(X$, 97, 1))
           NROV&(JI%) = CVS(Mid$(X$, 3, 4))
           IFAC#(JI%) = CVD(Mid$(X$, 99, 8)) + CVD(Mid$(X$, 123, 8))
           PRIMOSTO&(JI%) = 0
           ULTMOSTO&(JI%) = 0
           '
        Next IL&
        REMEVENT.Check1(5).Value = 1
        '
        Return
        '
'*****************************************************************************
        '
11200   LU$ = LUDAT$
        Archi$ = "ICV" + ABREMES$ + ".DAT"
        ARCHIALT$ = "ICV" + Left$(ABREMES$, 3) + ".DAT"
        If EXISTE%(LU$ + Archi$) < 1 Then
            If EXISTE%(LU$ + ARCHIALT$) > 0 Then
                RESULTA% = COPYFILE%(LU$ + ARCHIALT$, LU$ + Archi$)
            End If
        End If
        '
        N2% = FILEOPEN(LU$ + Archi$, 0, 96)
        REGMAX% = 0
        '
        FIN& = LOF(N2%) / 96
        REMEVENT.Check1(7).Value = 2
        '
        For JI% = 1 To FIN&
           '
           U& = JI%
           Call TRACE(20, U&, FIN&)
           '
           Get #N2%, JI%, X0$
           X1$ = Left$(X0$, 2)
           X2$ = Mid$(X0$, 3, 4)
           X3$ = Mid$(X0$, 7, 2)
           X4$ = Mid$(X0$, 9, 2)
           X5$ = Mid$(X0$, 11, 6)
           X6$ = Mid$(X0$, 17, 80)
           '
           TIXX$ = X1$
           NRO& = CVS(X2$)
           NC% = CVI(X3$)
           FEMI% = CVI(X4$)
           SUICC# = 0: For KK% = 1 To 8
                       If CVI(Mid$(X6$, 10 * KK% - 9, 2)) > 0 Then
                           SUICC# = SUICC# + Abs(CVD(Mid$(X6$, 10 * KK% - 7, 8)))
                       End If
                    Next KK%
                    '
           For KI% = 1 To CAFAC%
             If NCV%(KI%) = NC% Then
               If FFII%(KI%) = FEMI% Then
                 If NROV&(KI%) = NRO& Or (NRO& = 0 And TIS$(TII%(KI%)) = "AS") Then
                   If TIS$(TII%(KI%)) = TIXX$ Then
                     If Abs(Abs(IFAC#(KI%)) - SUICC#) < 0.005 Then
                       REGCON& = KI%
                       X$ = REGLEIDO$("CONCIVEN", REGCON&)
                       Call REPLA(X$, X6$, 17, 80)
                       Call REPLA(X$, MKI$(1), 155, 2)
                       Call GRAREG("CONCIVEN", X$, REGCON&)
                       GoTo 11290
                     End If
                   End If
                 End If
               End If
             End If
           Next KI%
           '
           LSet X0$ = String$(96, 0)        ' tacha movimiento en ICCmmaa.DAT
           Put #N2%, JI%, X0$
           '
11290   Next JI%
        REMEVENT.Check1(7).Value = 1
        '
        Return
        '
'*****************************************************************************
        '
11300   PRINOCON& = 0
        FIN& = ULTREG&("CONCIVEN")
        '
        For IL& = 1 To FIN&
            Call TRACE(20, IL&, FIN&)
            X$ = REGLEIDO$("CONCIVEN", IL&)
            If Abs(CVD(Mid$(X$, 139, 8))) >= 0.005 Then
                If CVI(Mid$(X$, 155, 2)) <> 1 Then
                    TIII$ = Left$(X$, 2)
                    If TIII$ = "VC" Or TIII$ = "FC" Or TIII$ = "ND" Or TIII$ = "NC" Then
                       If PRINOCON& = 0 Then
                          PRINOCON& = IL&
                          GoTo 11390
                       End If
                    End If
                End If
            End If
        Next IL&
        '
11390   'REMEVENT.Check4.Value = 1
        Return
        '
'*****************************************************************************
        '
12000   FIN& = ULTREG&("MOVISTO")
        PRIRE& = FIN&: ULTRE& = 0
        REMEVENT.Check1(6).Value = 2
        '
        NMS% = NUMAS%
        UFACA% = 1
        VDEF$ = ""
        For IL& = 1 To FIN&
           Call TRACE(20, IL&, FIN&)
           X$ = REGLEIDO$("MOVISTO", IL&)
           CODIMAT% = CVI(Mid$(X$, 3, 2))
             If CODIMAT% > NMS% Then CODIMAT% = 0
           FF% = CVI(Left$(X$, 2))
           If FF% = 0 Then
              FF% = FFANT%
              Call REPLA(X$, MKI$(FFANT%), 7, 2)
              Call GRAREG("MOVISTO", X$, IL&)
           End If
           FFANT% = FF%
           If FF% >= DES% Then
               If PRIMOVISTO& < 1 Then PRIMOVISTO& = IL&
               If FF% <= HAS% Then
                   ULTMOVISTO& = IL&
                   COMO$ = UCase$(Mid$(X$, 33, 2))
                   If COMO$ = "FC" Or COMO$ = "NC" Or COMO$ = "ND" Or COMO$ = "RT" Then
                       NRO1& = Val(Mid$(X$, 36, 7))
                       CAN# = CVD(Mid$(X$, 96, 8)) - CVD(Mid$(X$, 88, 8))
                       PRE# = CVD(Mid$(X$, 73, 8))
                       TOT# = Abs(PRE# * CAN#)
                       If IL& < PRIRE& Then PRIRE& = IL&
                       If IL& > ULTRE& Then ULTRE& = IL&
                       '
                       CI% = CVI(Mid$(X$, 3, 2))
                       If CI% > 0 Then
                         If MODOCALVE% = 0 Then
                           If CUECOD%(CI%) < 1 Then
12010                        Screen.MousePointer = 1
                             If VDEF$ = "" Then
                                  VDEF$ = MKI$(CI%) + MKI$(0)
                                Else
                                  Call REPLA(VDEF$, MKI$(CI%), 1, 2)
                             End If
                             OBJE$ = OBJPLA$("CARGRUCOD", VDEF$)
                             If OBJE$ = "" Then GoTo 999
                             Screen.MousePointer = 11
                             CUE1% = CVI(Mid$(OBJE$, 3, 2))
                             If CUE1% = 0 Then GoTo 12010
                             For U& = 1 To ULTREG&("GRUCOVEN")
                                 X$ = REGLEIDO$("GRUCOVEN", U&)
                                 If CVI(Left$(X$, 2)) = CUE1% Then GoTo 12020
                             Next U&
                             GoTo 12010
                             '
12020                        If CUE1% < 1 Then GoTo 12010
                             CUECOD%(CI%) = CUE1%
                             VDEF$ = OBJE$
                             REGCLI& = CI% + 1
                             If REGCLI& > 1 Then
                               RECLI2$ = REGLEIDO$("MASTER", REGCLI&)
                               Call REPLA(RECLI2$, MKI$(CUE1%), 125, 2)
                               Call GRAREG("MASTER", RECLI2$, REGCLI&)
                             End If
                             '
                           End If
                         End If
                       End If
                       '
12030                  For JI% = UFACA% To CAFAC%
                         If FFII%(JI%) = FF% Then
                           If NROV&(JI%) = NRO1& Then
                             IDET#(JI%) = IDET#(JI%) + TOT#
                               If CODIMAT% > 0 Then
                                 If PRIMOSTO&(JI%) = 0 Then
                                    PRIMOSTO&(JI%) = IL&
                                 End If
                                 If IL& > ULTMOSTO&(JI%) Then
                                    ULTMOSTO&(JI%) = IL&
                                 End If
                                 UFACA% = JI% + 1
                               End If
                               GoTo 12090
                             End If
                          End If
                       Next JI%
                       If UFACA% > 1 Then
                           UFACA% = 1
                           GoTo 12030
                       End If
                       '
                   End If
               End If
           End If
12090   Next IL&
        '
        '
12100   FIN& = ULTREG&("MOVISVIE")
        PRIRE1& = FIN&: ULTRE1& = 0
        REMEVENT.Check1(6).Value = 2
        '
        UFACA% = 1
        VDEF$ = ""
        For IL& = 1 To FIN&
           Call TRACE(20, IL&, FIN&)
           X$ = REGLEIDO$("MOVISVIE", IL&)
           FF% = CVI(Left$(X$, 2))
           If FF% >= DES% Then
               If PRIMATVEN& < 1 Then PRIMATVEN& = IL&
               If FF% <= HAS% Then
                  ULTMATVEN& = IL&
                  COMO$ = UCase$(Mid$(X$, 33, 2))
                  If COMO$ = "FC" Or COMO$ = "NC" Or COMO$ = "ND" Or COMO$ = "RT" Then
                       NRO1& = Val(Mid$(X$, 36, 7))
                       CAN# = CVD(Mid$(X$, 96, 8)) - CVD(Mid$(X$, 88, 8))
                       PRE# = CVD(Mid$(X$, 73, 8))
                       TOT# = Abs(PRE# * CAN#)
                       If IL& < PRIRE1& Then PRIRE1& = IL&
                       If IL& > ULTRE1& Then ULTRE1& = IL&
                       '
                       CI% = CVI(Mid$(X$, 3, 2))
                       If CI% > 0 Then
                         If MODOCALVE% = 0 Then
                           If CUECOD%(CI%) < 1 Then
12110                        Screen.MousePointer = 1
                             If VDEF$ = "" Then
                                  VDEF$ = MKI$(CI%) + MKI$(0)
                                Else
                                  Call REPLA(VDEF$, MKI$(CI%), 1, 2)
                             End If
                             OBJE$ = OBJPLA$("CARGRUCOD", VDEF$)
                             If OBJE$ = "" Then GoTo 999
                             Screen.MousePointer = 11
                             CUE1% = CVI(Mid$(OBJE$, 3, 2))
                             If CUE1% = 0 Then GoTo 12110
                             For U& = 1 To ULTREG&("GRUCOVEN")
                                 X$ = REGLEIDO$("GRUCOVEN", U&)
                                 If CVI(Left$(X$, 2)) = CUE1% Then GoTo 12120
                             Next U&
                             GoTo 12110
                             '
12120                        If CUE1% < 1 Then GoTo 12110
                             CUECOD%(CI%) = CUE1%
                             VDEF$ = OBJE$
                             REGCLI& = CI% + 1
                             If REGCLI& > 1 Then
                                RECLI2$ = REGLEIDO$("MASTER", REGCLI&)
                                Call REPLA(RECLI2$, MKI$(CUE1%), 125, 2)
                                Call GRAREG("MASTER", RECLI2$, REGCLI&)
                             End If
                             '
                           End If
                         End If
                       End If
                       '
12130                  For JI% = UFACA% To CAFAC%
                          If FFII%(JI%) = FF% Then
                             If NROV&(JI%) = NRO1& Then
                                IDET#(JI%) = IDET#(JI%) + TOT#
                                If PRIMOSTO&(JI%) = 0 Then
                                    PRIMOSTO&(JI%) = (-1) * IL&
                                End If
                                If IL& > ULTMOSTO&(JI%) Then
                                    ULTMOSTO&(JI%) = IL&
                                End If
                                UFACA% = JI% + 1
                                GoTo 12190
                             End If
                          End If
                       Next JI%
                       If UFACA% > 1 Then
                           UFACA% = 1
                           GoTo 12130
                       End If
                       '
                   End If
               End If
           End If
12190   Next IL&
        '
        Return
        '
'*****************************************************************************
        '
12200   For JI% = 1 To CAFAC%
           REGCON& = JI%
           X$ = REGLEIDO$("CONCIVEN", REGCON&)
           COEF# = 1
           If IDET#(JI%) > 0 Then
               COEF# = IFAC#(JI%) / IDET#(JI%)
           End If
           X$ = REPLACE$(X$, MKD$(Abs(COEF#)), 147, 8)
           Call GRAREG("CONCIVEN", X$, REGCON&)
        Next JI%
        '
        REMEVENT.Check1(6).Value = 1
12290   Return
        '

'*****************************************************************************
        '
13000   FIN& = ULTREG&("CONCIVEN")
        '
        REMEVENT.Check1(8).Value = 2
        SUMATO# = 0
        INCOMPLE% = 0
        '
        For IL& = 1 To FIN&
           Call TRACE(20, IL&, FIN&)
           X$ = REGLEIDO$("CONCIVEN", IL&)
           '
' AQUI SE MODIFICO
           INET# = CVD(Mid$(X$, 99, 8)) + CVD(Mid$(X$, 123, 8))
           SUIPU# = 0
           SUMATO# = SUMATO# + INET#
           SUMIVA# = SUMIVA# + CVD(Mid$(X$, 107, 8)) + CVD(Mid$(X$, 131, 8))
           SUMRNI# = SUMRNI# + CVD(Mid$(X$, 115, 8))
           '
           IMPUTA$ = Mid$(X$, 17, 80)
           AMO$ = "MOVISTO": PRIMO& = PRIMOSTO&(IL&): ULTMO& = ULTMOSTO&(IL&)
           If PRIMO& < 0 Then
               AMO$ = "MOVISVIE": PRIMO& = Abs(PRIMO&): ULTMO& = Abs(ULTMO&)
           End If
           '
           If MODOCALVE% > 0 Then
             NCLIE% = CVI(Mid$(X$, 7, 2))
             PCLI% = PIVACLI%(NCLIE%)
             If PCLI% < 0 Or PCLI% > 7 Then PCLI% = 1
             '
             TI% = Asc(Mid$(X$, 97, 1))
             CACO% = 0
             For JI% = 1 To 8
               CUEV%(JI%) = 0: IMXX#(JI%) = 0
             Next JI%
             '
             If (TI% > 0 And TI% <= 2) Or (TI% = 6 And PRIMO& > 0) Then
               CACO% = 1
               CUEV%(1) = COCUEVEN%(PCLI%)
               IMXX#(1) = INET#
               IMPUTA$ = MKI$(CUEV%(1)) + MKD$(Abs(IMXX#(1))) + String$(70, 0)
               SUMAIMPU# = INET#
               GoTo 13555
             End If
             '
             If TI% = 3 Then
             '  CACO% = 1
             '  CUEV%(1) = COCUEVEN%(8)
             '  IMXX#(1) = INET#
             '  IMPUTA$ = MKI$(CUEV%(1)) + MKD$(Abs(IMXX#(1))) + String$(70, 0)
             '  SUMAIMPU# = INET#
               GoTo 13560
             End If
             '
             If TI% = 6 Then
             '  CACO% = 1
             '  CUEV%(1) = COCUEVEN%(9)
             '  IMXX#(1) = INET#
             '  IMPUTA$ = MKI$(CUEV%(1)) + MKD$(Abs(IMXX#(1))) + String$(70, 0)
             '  SUMAIMPU# = INET#
               GoTo 13560
             End If
             '
             CACO% = 1
             CUEV%(1) = COCUEVEN%(10)
             IMXX#(1) = INET#
             IMPUTA$ = MKI$(CUEV%(1)) + MKD$(Abs(IMXX#(1))) + String$(70, 0)
             SUMAIMPU# = INET#
             GoTo 13555
             '
           End If
           '
           If PRIMO& > 0 Then
              If ULTMO& >= PRIMO& Then
                 '
                 CACO% = 0
                 For JI% = 1 To 8
                    CUEV%(JI%) = 0: IMXX#(JI%) = 0
                 Next JI%
                 '
                 For IK& = PRIMO& To ULTMO&
                     If IK& < 1 Then GoTo 13540
                     MS$ = REGLEIDO$(AMO$, IK&)
                     FF% = CVI(Left$(MS$, 2))
                     'If AMO$ = "MATVEN" Then FF% = CVI(Mid$(MS$, 7, 2))
                     '
                     If FF% = FFII%(IL&) Then
'                        If AMO$ = "MATVEN" Then
'                            COMO$ = "FC"
'                            NC% = CVI(Left$(MS$, 2))
'                            NRO1& = CVS(Mid$(MS$, 3, 4))
'                            CI% = CVI(Mid$(MS$, 9, 2))
'                            CAN# = CVS(Mid$(MS$, 19, 4))
'                            PRE# = CVD(Mid$(MS$, 11, 8))
'                          Else
                            COMO$ = UCase$(Mid$(MS$, 33, 2))
                            NC% = CVI(Mid$(MS$, 84, 2))
                            NRO1& = Val(Mid$(MS$, 36, 7))
                            CI% = CVI(Mid$(MS$, 3, 2))
                            CAN# = CVD(Mid$(MS$, 96, 8)) - CVD(Mid$(MS$, 88, 8))
                            PRE# = CVD(Mid$(MS$, 73, 8))
                        'End If
                        '
                        If COMO$ = "FC" Or COMO$ = "NC" Or COMO$ = "ND" Then
                           If NC% = NCV%(IL&) Then
                              If NRO1& = NROV&(IL&) Then
                                  If CI% > 0 Then
                                     TOT# = Abs(Int(100 * PRE# * CAN# + 0.5)) / 100
                                     If SUIPU# + TOT# > Abs(INET#) Then TOT# = Abs(INET#) - SUIPU#
                                     SUIPU# = SUIPU# + TOT#
                                     '
                                     ' GoSub 13700
                                     '
                                     GRUCODI% = CUECOD%(CI%)
                                     GRUCLII% = GRCOCLI%(NC%)
                                     '
                                     CUECONI% = 0
                                     For KK% = 1 To CARUCO%
                                        If GRUCOD%(KK%) = GRUCODI% Then
                                           If GRUCLI%(KK%) = GRUCLII% Then
                                              CUECONI% = CUECON%(KK%)
                                              GoTo 13010
                                           End If
                                        End If
                                     Next KK%
                                     '
13010                                If CUECONI% > 0 Then
                                        For JI% = 1 To CACO%
                                           If CUECONI% = CUEV%(JI%) Then
                                               IMXX#(JI%) = IMXX#(JI%) + TOT#
                                               GoTo 13540
                                           End If
                                        Next JI%
                                        '
                                        If CACO% >= 6 Then
                                            Call COMUNI("Documento " + COMO$ + Str$(NRO1&) + "\no puede Contabilizarse Autom ticamente")
                                            GoTo 13550
                                        End If
                                        '
                                        CACO% = CACO% + 1
                                        CUEV%(CACO%) = CUECONI%
                                        IMXX#(CACO%) = TOT#
                                        '
                                     End If
                                  End If
                              End If
                           End If
                        End If
                     End If
13540            Next IK&
                 '
13550            SUMAIMPU# = 0: MAXIM# = 0: IMAXI% = 0
                 For JI% = 1 To CACO%
                    IMXX#(JI%) = (Int(100 * IMXX#(JI%) + 0.5)) / 100
                    SUMAIMPU# = SUMAIMPU# + Abs(IMXX#(JI%))
                    If IMXX#(JI%) > MAXIM# Then
                        MAXIM# = IMXX#(JI%)
                        IMAXI% = JI%
                    End If
                 Next JI%
                 '
                 DIFE# = Abs(INET#) - SUMAIMPU#
                 If Abs(DIFE#) <= 1 Then
                    If IMAXI% > 0 Then
                       IMXX#(IMAXI%) = IMXX#(IMAXI%) + DIFE#
                    End If
                 End If
                 '
                 IMPUTA$ = "": SUMAIMPU# = 0
                 For JI% = 1 To 8
                    If JI% <= CACO% Then
                         IMPUTA$ = IMPUTA$ + MKI$(CUEV%(JI%)) + MKD$(Abs(IMXX#(JI%)))
                         SUMAIMPU# = SUMAIMPU# + Abs(IMXX#(JI%))
                       Else
                         IMPUTA$ = IMPUTA$ + String$(10, 0)
                    End If
                 Next JI%
                 '
              End If
           End If
           '
13555      If Abs(INET#) < 0.005 Then
               IMPUTA$ = String$(80, 0)
               GoTo 13570
           End If
           '
           vuelta% = 0
13560      SUMAIMPU# = 0: PRESE% = 0
           For JI% = 1 To 80 Step 10
              If CVI(Mid$(IMPUTA$, JI%, 2)) > 0 Then
                If ECOARCH("CUECON", Mid$(IMPUTA$, JI%, 2)) = "" Then
                  If vuelta% > 0 Then
                    Call MENSERR(24, "Cuenta Inexistente")
                  End If
                  Call REPLA(IMPUTA$, String$(80, 0), JI%, 81 - JI%)
                End If
              End If
              If CVI(Mid$(IMPUTA$, JI%, 2)) > 0 Then
                    IPAR1# = (Int(100 * CVD(Mid$(IMPUTA$, JI% + 2, 8)) + 0.5)) / 100
                    If IPAR1# < 0.005 Then IPAR1# = Abs(INET#) - SUMAIMPU#: PRESE% = 1
                  Else
                    IPAR1# = Abs(INET#) - SUMAIMPU#
              End If
              Call REPLA(IMPUTA$, MKD$(IPAR1#), JI% + 2, 8)
              SUMAIMPU# = SUMAIMPU# + Abs(IPAR1#)
              If CVI(Mid$(IMPUTA$, JI%, 2)) < 1 And Abs(IPAR1#) >= 0.005 Then
                 AAA = CVI(Mid$(IMPUTA$, JI%, 2))
                 AAI1% = AAA
                 'BBB$ = CUECONTA$(AAI1%)
                 PRESE% = 3
              End If
              If CVI(Mid$(IMPUTA$, JI%, 2)) > 0 And Abs(IPAR1#) < 0.005 Then PRESE% = 4
           Next JI%
           '
           If PRESE% > 0 Then
               If vuelta% = 0 Then
                   X$ = REGLEIDO$("CONCIVEN", IL&)
                   IMPUTA$ = Mid$(X$, 17, 80)
                   vuelta% = 1
                   GoTo 13560
               End If
           End If
           '
           If PRESE% > 0 Or Abs(Abs(INET#) - SUMAIMPU#) >= 0.005 Then
               Screen.MousePointer = 1
               VDEF$ = Left$(X$, 16) + IMPUTA$ + String$(24, 0)
               Call REPLA(VDEF$, MKD$(Abs(INET#)), 97, 8)
               Call REPLA(VDEF$, MKD$(SUMAIMPU#), 113, 8)
               Call REPLA(VDEF$, MKD$(Abs(INET#) - SUMAIMPU#), 105, 8)
               XD$ = OBJPLA$("ICOVENTA", VDEF$)
               If XD$ = "" Then
                   INCOMPLE% = 1
                   GoTo 13600
               End If
               IMPUTA$ = Mid$(XD$, 17, 80)
               Screen.MousePointer = 11
               vuelta% = 1
               GoTo 13560
           End If
           '
13570      For JI% = 1 To 80 Step 10
             CUEIX1% = CVI(Mid$(IMPUTA$, JI%, 2))
             If CUEIX1% > 0 Then
                If CUEIX1% = CUEIVA1% Or CUEIX1% = CUEADU1% Or CUEIX1% = CUERET1% Or CUEIX1% = CUEDEUV1% Then
                  TINUX$ = Left$(X$, 2) + "." + TRIM$(Str$(CVS(Mid$(X$, 3, 4))))
                  Call MENSERR(24, "Imputación Contable no Válida\para Comprobante " + TINUX$ + ".")
                End If
             End If
           Next JI%
           Call REPLA(X$, IMPUTA$, 17, 80)
           Call GRAREG("CONCIVEN", X$, IL&)
        Next IL&
        '
13600   REMEVENT.Check1(8).Value = 1
        Screen.MousePointer = 11
        '
        For IL& = 1 To ULTREG&("CONCIVEN")
           X$ = Left$(REGLEIDO$("CONCIVEN", IL&), 96)
           LSet X0$ = X$
           Put #N2%, IL&, X0$
        Next IL&
        '
        For J& = IL& + 1 To LOF(N2%) / 96
           LSet X0$ = String$(96, 0)
           Put #N2%, J&, X0$
        Next J&
        '
        Close #N2%
        Screen.MousePointer = 1
        Return
        '
'*****************************************************************************
        '
13700   TOT# = Abs(Int(100 * PRE# * CAN# + 0.5)) / 100
        If NC% <> 0 Then
           If NC% < 0 Or NC% > 32767 Then NC% = 32767
           '
           If CI% > 0 Then
               BUSX$ = MKI$(NC%) + MKI$(CI%)
               For K& = 1 To (Len(IREVE$)) / 4  '  To ULTREG&("REVECLI")
                   If Mid$(IREVE$, 4 * K& - 3, 4) = BUSX$ Then
                       RVCL$ = REGLEIDO$("REVECLI", K&)
                       CARVCL = CAN# + CVS(Mid$(RVCL$, 5, 4))
                       VARVCL# = TOT# + CVD(Mid$(RVCL$, 9, 8))
                       Call REPLA(RVCL$, MKS$(CARVCL) + MKD$(VARVCL#), 5, 12)
                       Call GRAREG("REVECLI", RVCL$, K&)
                       '
                       PVCL$ = REGLEIDO$("REVEPIE", K&)
                       Call REPLA(PVCL$, MKS$(CARVCL), 5, 4)
                       Call GRAREG("REVEPIE", PVCL$, K&)
                       '
                       GoTo 13799
                   End If
               Next K&
               '
               CARVCL = CAN#
               RVCL$ = MKI$(NC%) + MKI$(CI%) + MKS$(CARVCL) + MKD$(TOT#) + MKD$(0)
               Call REGAPP("REVECLI", RVCL$)
               '
               PVCL$ = MKI$(CI%) + MKI$(NC%) + MKS$(CARVCL) + MKD$(0) + String$(4, 0)
               Call REGAPP("REVEPIE", PVCL$)
               '
               IREVE$ = IREVE$ + MKI$(NC%) + MKI$(CI%)
               '
           End If
        End If
        '
13799   Return
        '
'*****************************************************************************
        '
13900   Screen.MousePointer = 11
        Call BLANARCH("REVECLI")
        Call BLANARCH("REVEPIE")
        IREVE$ = ""
        '
        REMEVENT.Check1(8).Value = 2
        '
        SUMATO# = 0
        INCOMPLE% = 0
        '
        If PRIMOVISTO& > 0 Then
           FIN& = ULTMOVISTO& - PRIMOVISTO& + 1
           For IL& = PRIMOVISTO& To ULTMOVISTO&
              U& = IL& - PRIMOVISTO& + 1
              Call TRACE(20, U&, FIN&)
              MS$ = REGLEIDO$("MOVISTO", IL&)
              FF% = CVI(Left$(MS$, 2))
              If FF% >= DES% And FF% <= HAS% Then
                 COMO$ = UCase$(Mid$(MS$, 33, 2))
                 If TRIM$(COMO$) = "" Then
                    COMO$ = UCase$(Mid$(MS$, 23, 2))
                 End If
                 If COMO$ = "RT" Or COMO$ = "FC" Or COMO$ = "NC" Or COMO$ = "ND" Then
                    NC% = CVI(Mid$(MS$, 84, 2))
                    NRO1& = Val(Mid$(MS$, 36, 7))
                    If NRO1& < 1 Then
                        NRO1& = Val(Mid$(MS$, 26, 7))
                    End If
                    CI% = CVI(Mid$(MS$, 3, 2))
                    CAN# = CVD(Mid$(MS$, 96, 8)) - CVD(Mid$(MS$, 88, 8))
                    PRE# = CVD(Mid$(MS$, 73, 8))
                    GoSub 13700
                 End If
              End If
           Next IL&
        End If
        '
        If PRIMATVEN& > 0 Then
           For IL& = PRIMATVEN To ULTMATVEN&
              U& = IL& - PRIMOVISTO& + 1
              Call TRACE(20, U&, FIN&)
              MS$ = REGLEIDO$("MOVISVIE", IL&)
              FF% = CVI(Left$(MS$, 2))
              If FF% >= DES% And FF% <= HAS% Then
                 COMO$ = UCase$(Mid$(MS$, 33, 2))
                 If TRIM$(COMO$) = "" Then
                    COMO$ = UCase$(Mid$(MS$, 23, 2))
                 End If
                 If COMO$ = "RT" Or COMO$ = "FC" Or COMO$ = "NC" Or COMO$ = "ND" Then
                    NC% = CVI(Mid$(MS$, 84, 2))
                    NRO1& = Val(Mid$(MS$, 36, 7))
                    If NRO1& < 1 Then
                        NRO1& = Val(Mid$(MS$, 26, 7))
                    End If
                    CI% = CVI(Mid$(MS$, 3, 2))
                    CAN# = CVD(Mid$(MS$, 96, 8)) - CVD(Mid$(MS$, 88, 8))
                    PRE# = CVD(Mid$(MS$, 73, 8))
                    GoSub 13700
                 End If
              End If
           Next IL&
        End If
        '
        REMEVENT.Check1(8).Value = 1
        '
13999   Return
        '
'*****************************************************************************
        '
14000   REMEVENT.Check1(9).Value = 2
        Call COPYSTRU("IFACOMP", "IFACVEN")
        Call BLANARCH("SUDIVEN")
        Call BLANARCH("COSUDIV")
        Call BLANARCH("SUIVAVE")
        Call BLANARCH("IFACVEN")
        If MODOSABO% = 1 Then Call BLANARCH("SUPERIB")
        '
        TTXXY$ = Chr$(1) + Chr$(1) + MKI$(HAS%)
        J& = 0
        For IL& = 1 To ULTREG&("ASIDIAR")
           X$ = REGLEIDO$("ASIDIAR", IL&)
           If Left$(X$, 4) <> TTXXY$ Then
              J& = J& + 1
              Call GRAREG("ASIDIAR", X$, J&)
           End If
        Next IL&
        J0& = J&
        '
        For K& = J& + 1 To ULTREG&("ASIDIAR")
           Call GRAREG("ASIDIAR", String$(48, 0), K&)
        Next K&
        '
        FIN& = ULTREG&("CONCIVEN")
        '
        For IL& = 1 To FIN&
           Call TRACE(20, IL&, FIN&)
           X$ = REGLEIDO$("CONCIVEN", IL&)
           Y$ = String$(128, 0)
           '
           FF% = CVI(Mid$(X$, 9, 2))
           Call REPLA(Y$, MKI$(FF%), 1, 2)
           '
           NC% = CVI(Mid$(X$, 7, 2))
           TI% = Asc(Mid$(X$, 97, 1))
           VA% = Asc(Mid$(X$, 98, 1))
           '
           If TI% = 2 Then
              If NC% > 0 Then
                 If Mid$(RACLI$(NC%), 47, 1) = Chr$(0) Then
                    VA% = 2
                 End If
              End If
           End If
           '
           If VA% < 1 Or VA% > 3 Then VA% = 1
           NRO& = CVS(Mid$(X$, 3, 4))
If NRO& = 2048 Then
AAA = BBB
End If
           NUSUCUVE% = CVI(Mid$(X$, 11, 2))
           NDX$ = FORMDOC$(NRO&, TI%, VA%, NUSUCUVE%)
           
           PVCLX1% = PIVACLI%(NC%)
           If PVCLX1% = 5 Or PVCLX1% = 7 Then
              Call REPLA(NDX$, PREFEXPO$, 4, 4)
              Call REPLA(NDX$, LETREXPO$, 18, 1)
           End If
           '
           Call REPLA(Y$, NDX$, 3, 18)
           '
           CLI$ = "Clientes de Contado"
           iva$ = "No Posee"
           If NC% > 0 Then
                 CLI$ = Left$(RACLI$(NC%), 30)
                 iva$ = Mid$(RACLI$(NC%), 31)
               ElseIf NC% < 0 Then
                 RECUI& = Abs(NC%)
                 If RECUI& <= ULTREG&("CUITS") Then
                    RCUX$ = REGLEIDO$("CUITS", RECUI&)
                    CLI$ = Left$(RCUX$, 30)
                    iva$ = Mid$(RCUX$, 31, 16)
                 End If
           End If
           '
           Call REPLA(Y$, CLI$, 21, 30)
           Call REPLA(Y$, iva$, 51, 16)
           '
           RELIB$ = Y$
           SV$ = Left$(Y$, 64) + String$(56, 0) + Space$(8)
           '
           SIG% = Sgn(CVD(Mid$(X$, 139, 8)))
           NEFAC1# = CVD(Mid$(X$, 99, 8))
           NEFAC2# = CVD(Mid$(X$, 123, 8))
           NEFAC# = NEFAC1# + NEFAC2#
           IVAD1# = CVD(Mid$(X$, 107, 8))
           IVAD2# = CVD(Mid$(X$, 131, 8))
           IVAD# = IVAD1# + IVAD2#
           RNI# = CVD(Mid$(X$, 115, 8))
           IPER = CVS(Mid$(X$, 157, 4))
           TOFAC1# = CVD(Mid$(X$, 139, 8))
           '
           Call REPLA(Y$, MKD$(TOFAC1#), 67, 8)
           '
           NCVI% = NC%
           If NCVI% < 0 Then NCVI% = 0
           CUE% = CUEMA%(NCVI%)
           IMPO# = TOFAC1#
           GoSub 14600
           '
           YY$ = Y$: CUES$ = Space$(12)
               RGI& = CUE%
               If RGI& <= ULTREG&("CUENTAS") Then
                    CUES$ = Left$(REGLEIDO$("CUENTAS", RGI&), 12)
               End If
               Call REPLA(YY$, CUES$, 75, 12)
               If TOFAC1# > 0 Then
                     Call REPLA(YY$, MKD$(TOFAC1#), 87, 8)
                   Else
                     Call REPLA(YY$, MKD$(Abs(TOFAC1#)), 95, 8)
               End If
               '
               Call REGAPP("COSUDIV", YY$)
               '
           XX6$ = Mid$(X$, 17, 80)
           SUMA# = 0: IRET# = 0
           '
           For KI% = 1 To 8
              CUE% = CVI(Mid$(XX6$, 10 * KI% - 9, 2))
              IMM# = Abs(CVD(Mid$(XX6$, 10 * KI% - 7, 8))) * SIG%
              IMPO# = (-1) * IMM#
              GoSub 14600
              IMX# = (-1) * IMM#
              '
              If CUE% > 0 Then
                 If Abs(IMM#) > 0 Then
                    CUES$ = Space$(12)
                    RGI& = CUE%
                    If RGI& <= ULTREG&("CUENTAS") Then
                       CUES$ = Left$(REGLEIDO$("CUENTAS", RGI&), 12)
                    End If
                    Call REPLA(Y$, CUES$, 75, 12)
                    If IMX# > 0 Then
                        Call REPLA(Y$, MKD$(IMX#), 87, 8)
                      Else
                        Call REPLA(Y$, MKD$(Abs(IMX#)), 95, 8)
                    End If
                    '
                    If UCase$(Left$(NDX$, 2)) <> "AS" Then Call REGAPP("SUDIVEN", Y$)
                    Call REGAPP("COSUDIV", Y$)
                    '
                    If CUE% = CUERET1% Then
                       IRET# = IRET# + IMM#
                    End If
                    '
                    SUMA# = SUMA# + IMM#
                    Y$ = RELIB$
                 End If
              End If
           Next KI%
           '
           NEGRA1# = NEFAC1#: NEXEN# = 0
           If Abs(IVAD1#) < 0.005 Then
              NEXEN# = NEGRA1#
              NEGRA1# = 0
           End If
           '
           If InStr(EPAC$, "DOSIVA") > 0 Then
             If Abs(NEGRA1#) > 0.5 Then
               ALITEOR = 100 * IVAD1# / NEGRA1#
               If Abs(ALITEOR - ALIVA) > 2.5 Then
                 NEGRAX# = 100 * IVAD1# / ALIVA
                 NEXEN# = NEGRA1# - NEGRAX#
                 NEGRA1# = NEGRAX#
               End If
             End If
           End If
           '
           NEGRA2# = NEFAC2#
           If Abs(IVAD2#) < 0.005 Then
              NEXEN# = NEXEN# + NEGRA2#
              NEGRA2# = 0
           End If
          '
           ' AGREGAR IVA
           If Abs(IVAD#) >= 0.005 Then
              CUES$ = CUEIVA$
              Call REPLA(Y$, CUES$, 75, 12)
              '
              CUE% = CUEIVA1%
              IMPO# = (-1) * IVAD#
              GoSub 14600
              '
              IMM# = (-1) * IVAD#
              If IMM# > 0 Then
                   Call REPLA(Y$, MKD$(IMM#), 87, 8)
                 Else
                   Call REPLA(Y$, MKD$(Abs(IMM#)), 95, 8)
              End If
              '
              If UCase$(Left$(NDX$, 2)) <> "AS" Then Call REGAPP("SUDIVEN", Y$)
              Call REGAPP("COSUDIV", Y$)
              '
              Y$ = RELIB$
           End If
           '
           ' AGREGAR IVA RNI
           If Abs(RNI#) >= 0.005 Then
              CUES$ = CUEADU$
              Call REPLA(Y$, CUES$, 75, 12)
              '
              CUE% = CUEADU1%
              IMPO# = (-1) * RNI#
              GoSub 14600
              '
              IMM# = (-1) * RNI#
              If IMM# > 0 Then
                   Call REPLA(Y$, MKD$(IMM#), 87, 8)
                 Else
                   Call REPLA(Y$, MKD$(Abs(IMM#)), 95, 8)
              End If
              '
              If UCase$(Left$(NDX$, 2)) <> "AS" Then Call REGAPP("SUDIVEN", Y$)
              Call REGAPP("COSUDIV", Y$)
              '
              Y$ = RELIB$
           End If
           '
            ' AGREGAR PERCEPCION INGRESOS BRUTOS
           If Abs(IPER) >= 0.005 Then
              CUES$ = CUEPER$
              Call REPLA(Y$, CUES$, 75, 12)
              '
              CUE% = CUEPIBRU%
              IMPO# = (-1) * IPER
              GoSub 14600
              '
              IMM# = (-1) * IPER
              If IMM# > 0 Then
                   Call REPLA(Y$, MKD$(IMM#), 87, 8)
                 Else
                   Call REPLA(Y$, MKD$(Abs(IMM#)), 95, 8)
              End If
              '
              If UCase$(Left$(NDX$, 2)) <> "AS" Then Call REGAPP("SUDIVEN", Y$)
              Call REGAPP("COSUDIV", Y$)
              '
              Y$ = RELIB$
           End If
           '
           If Abs(IRET#) >= 0.005 Then
             If Abs(IRET# - TOFAC1#) < 0.005 Then
               NEXEN# = 0
               NEGRA1# = 0
               NEGRA2# = 0
             End If
           End If
           '
           Call REPLA(SV$, MKD$(NEXEN#), 65, 8)
           Call REPLA(SV$, MKD$(NEGRA1#), 73, 8)
           Call REPLA(SV$, MKD$(IVAD1#), 81, 8)
           Call REPLA(SV$, MKD$(NEGRA2#), 89, 8)
           Call REPLA(SV$, MKD$(IVAD2#), 97, 8)
           Call REPLA(SV$, MKD$(NEGRA3#), 105, 8)
           Call REPLA(SV$, MKD$(IVAD3# + CDbl(IPER)), 113, 8)
           Call REPLA(SV$, MKD$(RNI#), 121, 8)
           Call REPLA(SV$, MKD$(IRET#), 129, 8)
           '

           Call REPLA(SV$, MKD$(TOFAC1#), 137, 8)
           '
           If Abs(TOFAC1#) < 0.005 Then
              Call REPLA(SV$, "ANULADA", 145, 8)
           End If
           '
           If UCase$(Left$(NDX$, 2)) <> "AS" Then
              '
              Call REGAPP("SUIVAVE", SV$)
              '
              If MODOSABO% = 1 Then
                If Abs(IPER) >= 0.005 Then
                  SVIB$ = Left$(SV$, 64) + Space$(14) + String$(64, 0)
                  Call REPLA(SVIB$, NIBRCLI$(NC%), 65, 14)
                  Call REPLA(SVIB$, MKD$(NEXEN# + NEGRA1# + NEGRA2# + NEGRA3#), 79, 8)
                  SIGPER% = Sgn(NEXEN# + NEGRA1# + NEGRA2# + NEGRA3#)
                  IPERSI# = CDbl(Abs(IPER) * SIGPER%)
                  If Abs(IPERSI#) < 0.01 Then  'porque signo esta en cero
                    IPERSI# = IPER
                   End If
                  Call REPLA(SVIB$, MKD$(IPERSI#), 87, 8)
                  Call REGAPP("SUPERIB", SVIB$)
                End If
              End If
              '
              SUMAFAC#(11) = SUMAFAC#(11) + NEGRA1#
              SUMAFAC#(12) = SUMAFAC#(12) + NEXEN#
              SUMAFAC#(14) = SUMAFAC#(14) + IVAD1#
              SUMAFAC#(22) = SUMAFAC#(22) + NEGRA2#
              SUMAFAC#(23) = SUMAFAC#(23) + IVAD2#
              SUMAFAC#(15) = SUMAFAC#(15) + RNI#
              SUMAFAC#(16) = SUMAFAC#(16) + IRET#
              SUMAFAC#(17) = SUMAFAC#(17) + OTRO# + IPER
              '
              SUMAFAC#(13) = SUMAFAC#(11) + SUMAFAC#(12) + SUMAFAC#(22)
              SUMAFAC#(18) = SUMAFAC#(14) + SUMAFAC#(15) + SUMAFAC#(16) + SUMAFAC#(17) + SUMAFAC#(23) + SUMAFAC#(24)
              SUMAFAC#(19) = SUMAFAC#(11) + SUMAFAC#(14) + SUMAFAC#(15) + SUMAFAC#(16) + SUMAFAC#(17)
              SUMAFAC#(25) = SUMAFAC#(22) + SUMAFAC#(23) + SUMAFAC#(24)
              SUMAFAC#(20) = SUMAFAC#(12)
              SUMAFAC#(21) = SUMAFAC#(19) + SUMAFAC#(20) + SUMAFAC#(25)
              '
              For XII% = 11 To 25
                  REMEVENT.Text1(XII%).TEXT = CSTRING$(MKD$(SUMAFAC#(XII%)), 4, 10, 2, 2)
                  REMEVENT.Text1(XII%).Refresh
              Next XII%
              '
           End If
           TTXYZ$ = String$(64, 0)
           Call REPLA(TTXYZ$, Mid$(SV$, 3, 18), 1, 18)
           Call REPLA(TTXYZ$, Mid$(SV$, 21, 30), 21, 30)
           Call REPLA(TTXYZ$, MKD$(NEGRA1# + NEXEN# + NEGRA2#), 51, 8)
           Call REPLA(TTXYZ$, MKI$(NC%), 59, 2)
           Call REPLA(TTXYZ$, MKS$(NRO), 61, 4)
           Call REGAPP("IFACVEN", TTXYZ$)
           '
        Next IL&
        '
        Call CIERRARCH("*.*")
        Call GRARGEN("CONCIVEN", TRIM$(CORRESP$))
        '
        Call HEADERS("SUDIVEN", "")
        Call HEADERS("SUDIVEN", "Corresponde a: " + CORRESP$)
        '
        Call HEADERS("SUIVAVE", "")
        Call HEADERS("SUIVAVE", "Corresponde a: " + CORRESP$)
        '
        Call CIERRARCH("*.*")
        '
        Return


14600   BUS$ = TTXXY$ + MKI$(CUE%)
        IMPOX# = IMPO#
        '
        For K& = J0& + 1 To J&
            Z$ = REGLEIDO$("ASIDIAR", K&)
            If Left$(Z$, 6) = BUS$ Then
               IMPO# = (CVD(Mid$(Z$, 39, 8))) + ((Int(100 * IMPO# + 0.5)) / 100)
               Z$ = REPLACE$(Z$, MKD$(IMPO#), 39, 8)
               Call GRAREG("ASIDIAR", Z$, K&)
               GoTo 14690
            End If
        Next K&
        Z$ = REGBLAN$("ASIDIAR")
        Z$ = REPLACE$(Z$, BUS$, 1, 6)
        Z$ = REPLACE$(Z$, MKD$((Int(100 * IMPO# + 0.5)) / 100), 39, 8)
        J& = J& + 1: K& = J&
        Call GRAREG("ASIDIAR", Z$, J&)
        '
14690   SUTOTA# = SUTOTA# + IMPOX#
        Return
        '


15000   CORRESP$ = AJUSTI$(REMEVENT.MESAN.TEXT, 12)
        FGENE% = HOY(HO$): If INCOMPLE% > 0 Then FGENE% = 0
        HGENE% = HORANUM(Time$): If INCOMPLE% > 0 Then HGENE% = 0
        '
        REBLA$ = String$(300, 0)
        Call REPLA(REBLA$, CORRESP$, 1, 12)
        Call REPLA(REBLA$, MKI$(FGENE%), 13, 2)
        Call REPLA(REBLA$, MKI$(HGENE%), 15, 2)
        For KI% = 0 To 31
            Call REPLA(REBLA$, MKD$(SUMAFAC#(KI%)), 45 + 8 * KI%, 8)
        Next KI%
        '
        NX% = FILEOPEN%(LU$ + "IRESMEN.DAT", 0, 300)
        Dim REBLB As String * 300
        '
        For KI% = 1 To LOF(NX%) / 300
            Get #NX%, KI%, REBLB$
            If Left$(REBLB$, 12) = CORRESP$ Then
                REBLB$ = REBLA$
                Put #NX%, KI%, REBLB$
                GoTo 15090
            End If
        Next KI%
        '
        KI% = 1 + LOF(NX%) / 300
        REBLB$ = REBLA$
        Put #NX%, KI%, REBLB$
        '
15090   Close #NX%
        '
        FEX = CVI(Mid$(REBLA$, 13, 2))
        REMEVENT.Text4(0).TEXT = FECHATEX$(FEX)
        If FEX < 1 Then REMEVENT.Text4(0).TEXT = "00/00/00"
        REMEVENT.Text4(0).Refresh
        HOX% = CVI(Mid$(REBLA$, 15, 2))
        REMEVENT.Text4(1) = HORATEX$(HOX%)
        REMEVENT.Text4(1).Refresh
        '
        Return

End Sub
'

Sub CARGARESMEN()
        '
        REBLA$ = Space$(12) + String$(288, 0)
        CORRESP$ = AJUSTI$(REMEVENT.MESAN.TEXT, 12)
        NX% = FILEOPEN%(LUDAT$ + "IRESMEN.DAT", 0, 300)
        Dim REBLB As String * 300
        '
        For KI% = 1 To LOF(NX%) / 300
            Get #NX%, KI%, REBLB$
            If Left$(REBLB$, 12) = CORRESP$ Then
                REBLA$ = REBLB$
            End If
        Next KI%
        '
        FEX = CVI(Mid$(REBLA$, 13, 2))
        REMEVENT.Text4(0).TEXT = FECHATEX$(FEX)
        If FEX < 1 Then REMEVENT.Text4(0).TEXT = "00/00/00"
        REMEVENT.Text4(0).Refresh
        HOX% = CVI(Mid$(REBLA$, 15, 2))
        REMEVENT.Text4(1) = HORATEX$(HOX%)
        REMEVENT.Text4(1).Refresh
        '
        For KI% = 0 To 7
            SUMAFAC#(KI%) = CVD(Mid$(REBLA$, 45 + 8 * KI%, 8))
            REMEVENT.Text1(KI%).TEXT = CSTRING$(MKD$(SUMAFAC#(KI%)), 4, 11, 2, 2)
            REMEVENT.Text1(KI%).Refresh
        Next KI%
        '
        For KI% = 11 To 25
            SUMAFAC#(KI%) = CVD(Mid$(REBLA$, 45 + 8 * KI%, 8))
            REMEVENT.Text1(KI%).TEXT = CSTRING$(MKD$(SUMAFAC#(KI%)), 4, 10, 2, 2)
            REMEVENT.Text1(KI%).Refresh
        Next KI%
        '
        Close #NX%

End Sub


Sub VALICUEIVA()
     '
     Screen.MousePointer = 11
     YACUE% = 0
     '
     LU$ = LUDAT$
     TDAT$ = "AS,VC,FC,ND,RB,DC,NC,AS"
     For I% = 0 To 7
        TIS$(I%) = Mid$(TDAT$, 3 * I% + 1, 2)
     Next I%
     '
     CUEIVA$ = AJUSTI$(CONTROL$("SVD451S", "CUEIVA"), 12)
     CUEADU$ = AJUSTI$(CONTROL$("SVD451S", "CUEADU"), 12)
     CUERET$ = AJUSTI$(CONTROL$("SVD451S", "CUERET"), 12)
     CUEDEU$ = AJUSTI$(CONTROL$("AMACLI", "ECUECON"), 12)
     '
     CUEIVA1% = 0: CUEADU1% = 0: CUERET1% = 0: CUEDEUV1% = 0
     For U& = 1 To ULTREG&("ECUECON")
        X$ = REGLEIDO$("ECUECON", U&)
        If Left$(X$, 12) = CUEIVA$ Then CUEIVA1% = CVI(Mid$(X$, 41, 2)): DECIVA$ = Mid$(X$, 13, 28)
        If Left$(X$, 12) = CUEADU$ Then CUEADU1% = CVI(Mid$(X$, 41, 2)): DECADU$ = Mid$(X$, 13, 28)
        If Left$(X$, 12) = CUERET$ Then CUERET1% = CVI(Mid$(X$, 41, 2)): DECRET$ = Mid$(X$, 13, 28)
        If Left$(X$, 12) = CUEDEU$ Then CUEDEUV1% = CVI(Mid$(X$, 41, 2)): DECRET$ = Mid$(X$, 13, 28)
     Next U&
     '
     If CUEIVA1% < 1 Or CUEADU1% < 1 Or CUERET1% < 1 Then
         Call MENSERR(24, "Faltan Cuentas Contables Correspondientes\a I.V.A. Responsables Inscriptos y No Inscriptos.")
         Screen.MousePointer = 1
         Exit Sub
     End If
     '
     If CUEDEUV1% < 1 Then
         Call MENSERR(24, "Faltan Cuentas Contables Correspondientes\a Deudores por Ventas (omisión).")
         Screen.MousePointer = 1
         Exit Sub
     End If
     '
     '
     MODOCALVE% = 0
     If CONTROL$("REMEVENT", "MODCALVE") = "PIVACLI" Then
       MODOCALVE% = 1
     End If
     '
     If MODOCALVE% > 0 Then
        '
        For KKI% = 0 To 11
          CUEVENTA$(KKI%) = AJUSTI$(CONTROL$("SVD451S", "CUEVE-" + TRIM$(Str$(KKI%))), 12)
        Next KKI%
        '
        For U& = 1 To ULTREG&("ECUECON")
          X$ = REGLEIDO$("ECUECON", U&)
          For KKI% = 0 To 11
            If Left$(X$, 12) = CUEVENTA$(KKI%) Then COCUEVEN%(KKI%) = CVI(Mid$(X$, 41, 2)): DECIVA$ = Mid$(X$, 13, 28)
          Next KKI%
        Next U&
        '
        For KKI% = 0 To 11
          If COCUEVEN%(KKI%) < 1 Then
            Call MENSERR(24, "Faltan Cuentas Contables de Venta.")
            Screen.MousePointer = 1
            Exit Sub
          End If
        Next KKI%
        '
     End If
     '
     ' armar archivo de validacion CUECOVEN **************
     '
     Call COPYSTRU("CUECOVEN", "CUECOVEX")
     Call BLANARCH("CUECOVEX")
     '
     CAGRUCOVE% = 0: CAGRUCOCLI% = 0: vuelta% = 0
     If ULTREG&("GRUCOVEN") < 1 Then
         OPX% = COMALTER%("Falta Definir Grupos Contables\de Venta por Producto.\   \En Caso de Definirlos Ahora,\el Sistema Ofrecerá Valores por Omisión.\\Definir Ahora\Abandonar")
         If OPX% = 1 Then
            X$ = MKI$(1) + "Productos de Propia Produccion": Call REGAPP("GRUCOVEN", X$)
            X$ = MKI$(2) + "Material de Reventa": Call REGAPP("GRUCOVEN", X$)
            Call CIERRARCH("GRUCOVEN")
            GoTo 10
         End If
         Call FINAL("")
     End If
     '
10   Call BLANARCH("!GRUCOVEN")
     For II& = 1 To ULTREG&("GRUCOVEN")
         X$ = REGLEIDO$("GRUCOVEN", II&)
         If CVI(Left$(X$, 2)) > 0 Then
             Call REGAPP("!GRUCOVEN", X$)
         End If
     Next II&
     Call CIERRARCH("GRUCOVEN")
     '
     VTB% = VENTABU%("GRUCOVEN")
     If VTB% < 0 Then Exit Sub
     JI& = 0
     For II& = 1 To ULTREG&("!GRUCOVEN")
         X$ = REGLEIDO$("!GRUCOVEN", II&)
         If CVI(Left$(X$, 2)) > 0 Then
             JI = JI& + 1
             Call GRAREG("GRUCOVEN", X$, JI&)
         End If
     Next II&
     Call SETULTREG("GRUCOVEN", JI&)
     Call CIERRARCH("GRUCOVEN")
     '
     If ULTREG&("GRUCOCLI") < 1 Then
         OPX% = COMALTER%("Falta Definir Grupos de Clientes.\   \En Caso de Definirlos Ahora,\el Sistema Ofrecerá Valores por Omisión.\\Definir Ahora\Abandonar")
         If OPX% = 1 Then
            X$ = MKI$(1) + "Clientes en Cuenta Corriente": Call REGAPP("GRUCOCLI", X$)
            X$ = MKI$(2) + "Clientes de Contado": Call REGAPP("GRUCOCLI", X$)
            X$ = MKI$(9) + "Cuentas Suspendidas": Call REGAPP("GRUCOCLI", X$)
            Call CIERRARCH("GRUCOCLI")
            GoTo 20
         End If
         Call FINAL("")
     End If
     '
20   Call COPYSTRU("GRUCOCLI", "CLAVECLI")
     If ULTREG&("GRUCOCLI") <= 8 Then
          Call CIERRARCH("GRUCOCLI")
          Call CIERRARCH("CLAVECLI")
          FFXX% = COPYFILE%(LUDAT$ + "GRUCOCLI.DAT", LUDAT$ + "CLAVECLI.DAT")
          MODOCLAVE% = 1
        Else
          X$ = MKI$(1) + "Clientes Locales": Call GRAREG("CLAVECLI", X$, 1)
          X$ = MKI$(2) + "Clientes Exentos y Promovidos": Call GRAREG("CLAVECLI", X$, 2)
          X$ = MKI$(3) + "Clientes del Exterior": Call GRAREG("CLAVECLI", X$, 3)
          Call SETULTREG("CLAVECLI", 3)
          Call CIERRARCH("CLAVECLI")
          MODOCLAVE% = 2
     End If
     '
     If MODOCLAVE% <> 2 Then
         Call BLANARCH("!GRUCOCLI")
         For II& = 1 To ULTREG&("GRUCOCLI")
             X$ = REGLEIDO$("GRUCOCLI", II&)
             If CVI(Left$(X$, 2)) > 0 Then
                 Call REGAPP("!GRUCOCLI", X$)
             End If
         Next II&
         Call CIERRARCH("!GRUCOCLI")
         Call CIERRARCH("GRUCOCLI")
         '
         VTB% = VENTABU%("GRUCOCLI")
         If VTB% < 0 Then Exit Sub
         JI& = 0
         For II& = 1 To ULTREG&("!GRUCOCLI")
             X$ = REGLEIDO$("!GRUCOCLI", II&)
             If CVI(Left$(X$, 2)) > 0 Then
                 JI = JI& + 1
                 Call GRAREG("GRUCOCLI", X$, JI&)
             End If
         Next II&
         Call SETULTREG("GRUCOCLI", JI&)
         Call CIERRARCH("GRUCOCLI")
     End If
     '
     FINI& = ULTREG&("GRUCOVEN")
     Screen.MousePointer = 11
     '
     For IL& = 1 To ULTREG&("GRUCOVEN")
        Call TRACE(20, IL&, FINI&)
        X$ = REGLEIDO$("GRUCOVEN", IL&)
        XI1% = CVI(X$)
        '
        If XI1% > 0 Then
           If CAGRUCOVE% < 256 Then
               CAGRUCOVE% = CAGRUCOVE% + 1
               GRUCOVE%(CAGRUCOVE%) = XI1%
           End If
           '
           vuelta% = vuelta% + 1
           For J& = 1 To ULTREG&("CLAVECLI")
              Y$ = REGLEIDO$("CLAVECLI", J&)
              Y1% = CVI(Y$)
              If Y1% > 0 Then
                 If vuelta% = 1 Then
                    If CAGRUCOCLI% < 256 Then
                       CAGRUCOCLI% = CAGRUCOCLI% + 1
                       GRUCOC%(CAGRUCOCLI%) = Y1%
                    End If
                 End If
                 '
                 Z$ = MKI$(XI1%) + MKI$(Y1%) + String$(4, 0)
                 For K& = 1 To ULTREG&("CUECOVEN")
                    t$ = REGLEIDO$("CUECOVEN", K&)
                    If Left$(t$, 4) = Left$(Z$, 4) Then
                       Z$ = t$
                    End If
                 Next K&
                 Call REGAPP("CUECOVEX", Z$)
              End If
           Next J&
        End If
     Next IL&
     '
     Call CIERRARCH("CUECOVEN")
     Call CIERRARCH("CUECOVEX")
     Call FileCopy(LU$ + "CUECOVEX.DAT", LU$ + "CUECOVEN.DAT")
     '
     For IL& = 1 To ULTREG&("CUECOVEN")
        X$ = REGLEIDO$("CUECOVEN", IL&)
        If CVI(Left$(X$, 2)) > 0 Then
           If CVI(Mid$(X$, 3, 2)) > 0 Then
              CUES$ = Mid$(X$, 5, 2)
              If TRIM$(ECOARCH$("CUECON", CUES$)) = "" Then
                 Call REPLA(X$, CUES1$, 5, 2)
                 X$ = OBJPLA$("CARCUECOVEN", X$)
                 CUES1$ = Mid$(X$, 5, 2)
                 If X$ = "" Then
                    Exit Sub
                 End If
                 Call GRAREG("CUECOVEN", X$, IL&)
              End If
           End If
        End If
     Next IL&
     '
     CARUCO% = ULTREG&("CUECOVEN")
     '
'*****************************************************************************
     '
     FIN& = CARUCO%                     ' cargar vector
     For IL& = 1 To FIN&
        X$ = REGLEIDO$("CUECOVEN", IL&)
        JJ% = IL&
        GRUCOD%(JJ%) = CVI(Left$(X$, 2))
        GRUCLI%(JJ%) = CVI(Mid$(X$, 3, 2))
        CUECON%(JJ%) = CVI(Mid$(X$, 5, 2))
     Next IL&
     '
     Call CIERRARCH("*.*")
     YACUE% = 1
     '
End Sub
'

Sub ARMAMES()
     '
     CORRESP$ = TRIM$(REMEVENT.MESAN.TEXT)
     CORRESQ$ = MESANO$(CORRESP$, DES%, HAS%)
     FEX = DES%: DESFE$ = FECHATEX$(FEX)
     FEX = HAS%: HASFE$ = FECHATEX$(FEX)
     '
     CORRESQ$ = TRIM$(CORRESP$)
     ABREMES$ = Left$(CORRESQ$, 3) + Mid$(DESFE$, 7, 2)
     '
End Sub
'

Sub CARCUECOD()
    '
    YACUE% = 0
    FIN& = NUMAS%
    If FIN& > 32768 Then
        Call COMUNI("Base de Datos Saturada en Maestro de Items de Stock\Consulte a su Soporte de Sistemas.")
        Exit Sub
    End If
    '
    For IL& = 1 To FIN&
        Call TRACE(20, IL&, FIN&)
        RM$ = REGLEIDO$("MASTER", IL& + 1)
        GRUCODI% = CVI(Mid$(RM$, 125, 2))
        '                               validar GRUCODI%
        For GCI% = 1 To CAGRUCOVE%
            If GRUCODI% = GRUCOVE%(GCI%) Then GoTo 49
        Next GCI%
        GRUCODI% = 0
        '
49      CUECOD%(IL&) = GRUCODI%
    Next IL&
    '
    YACUE% = 1
    '
End Sub
'

Sub CARCUECLI()
    '
10  YACUE% = 0
    FIN& = ULTREG&("DEUDOR1")
    '
    CANCUEN% = ULTREG&("CUENTAS")
    YACLICON% = 0
    For JJL& = 1 To FIN&
        Call TRACE(20, JJL&, FIN&)
        CLX$ = REGLEIDO$("DEUDOR1", JJL&)
        NC% = CVI(Left$(CLX$, 2))
        If NC% = 0 Then
            Call REPLA(CLX$, "Clientes de Contado", 3, 30)
            Call GRAREG("DEUDOR1", CLX$, JJL&)
            YACLICON% = 1
        End If
        '
        If NC% >= 0 Then
            If NC% <= 32768 Then
                EXICLI%(NC%) = 1
                If MODOCLAVE% = 1 Then
                    GRUCLII% = CVI(Mid$(REGLEIDO$("DEUDOR2", JJL&), 125, 2))
                  Else
                    GRUCLII% = 1
                    PIVCLII% = Asc(Mid$(REGLEIDO$("DEUDOR2", JJL&), 104, 1))
                    If PIVCLII% = 7 Then GRUCLII% = 2
                    If PIVCLII% = 5 Then GRUCLII% = 3
                End If
If GRUCLII% <> 1 Then
AAA = BBB
End If
                '                       validar GRUCLII%
                For GCI% = 1 To CAGRUCOCLI%
                    If GRUCLII% = GRUCOC%(GCI%) Then GoTo 49
                Next GCI%
                GRUCLII% = 0
                '
49              GRCOCLI%(NC%) = GRUCLII%
                '
                CUEMA%(NC%) = CVI(Mid$(REGLEIDO$("DEUDOR2", JJL&), 127, 2))
                If CUEMA%(NC%) < 0 Or CUEMA%(NC%) > CANCUEN% Then
                    CUEMA%(NC%) = 0
                End If
                '
            End If
        End If
    Next JJL&
    YACUE% = 1
    '
End Sub
'

Sub CONASIDIA()
    '
    MESA$ = TRIM$(REMEVENT.MESAN.TEXT)
    If MESA$ = "" Then
        MESAN.SetFocus
        Exit Sub
    End If
    '
    MESB$ = MESANO$(MESA$, DES%, HAS%)
    If MESB$ <> MESA$ Then
        REMEVENT.MESAN.SetFocus
        Exit Sub
    End If
    '
    Screen.MousePointer = 11
    J& = 0
    Call BLANARCH("!ASIDIAFI")
    TTXXY$ = Chr$(1) + Chr$(1) + MKI$(HAS%)
    For IL& = 1 To ULTREG&("ASIDIAR")
        X$ = REGLEIDO$("ASIDIAR", IL&)
        If Left$(X$, 4) = TTXXY$ Then
            CUENU& = CVI(Mid$(X$, 5, 2))
            IDEB# = CVD(Mid$(X$, 39, 8))
            If Abs(IDEB#) >= 0.005 Then
                IHAB# = 0
                If IDEB# < 0 Then
                    IHAB# = Abs(IDEB#)
                    IDEB# = 0
                End If
                CUES$ = Space$(12)
                If CUENU& > 0 Then
                    If CUENU& <= ULTREG&("CUENTAS") Then
                        CUES$ = Left$(REGLEIDO$("CUENTAS", CUENU&), 12)
                    End If
                End If
                DENOCUES$ = ECOARCH$("ECUECON", CUES$)
                '
                Z$ = REGBLAN$("ASIDIAFI")
                Call REPLA(Z$, CUES$, 1, 12)
                Call REPLA(Z$, DENOCUES$, 13, 30)
                Call REPLA(Z$, MKD$(IDEB#), 43, 8)
                Call REPLA(Z$, MKD$(IHAB#), 51, 8)
                '
                J& = J& + 1
                Call GRAREG("!ASIDIAFI", Z$, J&)
            End If
        End If
    Next IL&
    Call SETULTREG("!ASIDIAFI", J&)
    '
    For K1& = 1 To J&
        For K2& = K1& + 1 To J&
            X1$ = REGLEIDO$("!ASIDIAFI", K1&)
            X2$ = REGLEIDO$("!ASIDIAFI", K2&)
            If Left$(X1$, 12) > Left$(X2$, 12) Then
                X0$ = X1$
                X1$ = X2$
                X2$ = X0$
                Call GRAREG("!ASIDIAFI", X1$, K1&)
                Call GRAREG("!ASIDIAFI", X2$, K2&)
            End If
        Next K2&
    Next K1&
    Call CIERRARCH("!ASIDIAFI")
    '
    Screen.MousePointer = 1
    VTB% = VENTABU%("ASIDIAFI")
    '
End Sub
'

Sub BACKUPMEN(MODO%, HAYERRO%)
    '
    ' MODO% = 1 ... existencia y acceso a directorio
    '
    Screen.MousePointer = 11
    Dim X As String * 2
    Dim ARESPA$(24)
    CARESPA% = 19
    '
    HAYERRO% = 0
    Call ARMAMES
    If Len(ABREMES$) <> 5 Then
        Call COMUNI("Período '" + ABREMES$ + "' no Procesable.\  \Consulte a su Soporte de Sistemas.")
        HAYERRO% = 1
        Exit Sub
    End If
    '
    ULOG$ = Left$(CurDir, 2)
    LUBAK$ = LUDAT$ + ABREMES$
    If Mid$(LUBAK$, 2, 1) <> ":" Then
       LUBAK$ = ULOG$ + LUBAK$
    End If
    EXTE$ = Mid$(DESFE$, 8, 1) + Mid$(DESFE$, 4, 2)
    '
    ARESPA$(0) = "ICV" + ABREMES$
    ARESPA$(1) = "DETACCD"
    ARESPA$(2) = "RESUCCD"
    ARESPA$(3) = "DETFACV"
    ARESPA$(4) = "DETDEBV"
    ARESPA$(5) = "DETCREV"
    ARESPA$(6) = "DETCOBV"
    ARESPA$(7) = "DETOTRV"
    ARESPA$(8) = "VEPROVIN"
    ARESPA$(9) = "VEPROVIX"
    ARESPA$(10) = "SUDIVEN"
    ARESPA$(11) = "SUIVAVE"
    ARESPA$(12) = "CONCIVEN"
    ARESPA$(13) = "RESUVEN"
    ARESPA$(14) = "REVECLI"
    ARESPA$(15) = "REVEPIE"
    ARESPA$(16) = "RECOBCLI"
    ARESPA$(17) = "COSUDIV"
    ARESPA$(18) = "IFACVEN"
    ARESPA$(19) = "SUPERIB"
    '
40  vuelta% = 1
41  ARXX$ = LUBAK$
    BRXX$ = CurDir
    On Error GoTo 43
    ChDir ARXX$
    ChDir BRXX$
    On Error GoTo 0
    GoTo 50
    '
43  ERNU$ = TRIM$(Str$(Err.Number))
    Resume 44
44  On Error GoTo 0
    If vuelta% = 1 Then
        On Error Resume Next
        MkDir ARXX$
        On Error GoTo 0
        vuelta% = 2
        GoTo 41
    End If
    '
    CRXX$ = REPLACAR(ARXX$, "\", "/")
    MENSA$ = "Se ha Producido un Error (" + ERNU$ + ") en la Apertura\"
    MENSA$ = MENSA$ + "para RESPALDO de la Siguiente Carpeta:\  \"
    MENSA$ = MENSA$ + CRXX$ + "\  \"
    '
    MENSA$ = MENSA$ + "Investigue las Siguientes Causas Posibles:\  \"
    MENSA$ = MENSA$ + "   1.- Problemas de Red - No hay Acceso a U.L. '" + Left$(ARXX$, 2) + "'.\"
    MENSA$ = MENSA$ + "   2.- Usuario de Red sin Derechos\"
    MENSA$ = MENSA$ + "       de Creación de Directorios.\"
    MENSA$ = MENSA$ + "   3.- Espacio en disco Insuficiente.\  \"
    MENSA$ = MENSA$ + "   \Consulte a su Soporte de Sistemas.\"
    '
    C$ = FORMESS$(MENSA$)
    RESPU% = MsgBox(C$, 53, "Error de Apertura de Carpeta")
    If RESPU% = 4 Then
        GoTo 40
    End If
    HAYERRO% = 1
    Exit Sub
    '
50  For I% = 0 To CARESPA%
        If TRIM$(ARESPA$(I%)) <> "" Then
            NX% = FILEOPEN(LUBAK$ + "\" + ARESPA$(I%) + "." + EXTE$, 0, 2)
            If NX% < 1 Then
                HAYERRO% = 1
                Exit Sub
            End If
            Get #NX%, 1, X$
            Put #NX%, 1, X$
            Close #NX%
        End If
    Next I%
    '
    '
    If MODO% = 2 Then
        LUDA$ = LUDAT$
        If Mid$(LUDA$, 2, 1) <> ":" Then
            LUDA$ = Left$(CurDir, 2) + LUDA$
        End If
        '
        Call CIERRARCH("*.*")
        For I% = 0 To CARESPA%
            If TRIM$(ARESPA$(I%)) <> "" Then
                '
                ORIGI$ = LUDA$ + ARESPA$(I%) + ".DAT"
                DESTI$ = LUBAK$ + "\" + ARESPA$(I%) + "." + EXTE$
                '
                If COPYFILE%(ORIGI$, DESTI$) <> 1 Then
                  HAYERRO% = 1
                  Exit Sub
                End If
                '
            End If
        Next I%
    End If
    '
    '
    If MODO% = 3 Then
        LUDA$ = LUDAT$
        If Mid$(LUDA$, 2, 1) <> ":" Then
            LUDA$ = Left$(CurDir, 2) + LUDA$
        End If
        C1$ = TRIM$(REMEVENT.MESAN.TEXT)
        C2$ = TRIM$(REMEVENT.Text4(0).TEXT)
        C3$ = TRIM$(REMEVENT.Text4(1).TEXT)
        '
        Call CIERRARCH("*.*")
        If TRIM$(CORRAR$("CONCIVEN")) = TRIM$(C1$) Then
            GoTo 80
        End If
        '
        For I% = 0 To CARESPA%
            If TRIM$(ARESPA$(I%)) <> "" Then
                '
                ORIGI$ = LUBAK$ + "\" + ARESPA$(I%) + "." + EXTE$
                DESTI$ = LUDA$ + ARESPA$(I%) + ".DAT"
                '
                If COPYFILE%(ORIGI$, DESTI$) <> 1 Then
                  HAYERRO% = 1
                  Exit Sub
                End If
                '
            End If
        Next I%
        '
        Call HEADERS("DETACCX", "")
        Call HEADERS("DETACCX", "Corresponde a: " + TRIM$(C1$) + " - Generado " + C2$ + " - " + C3$ + " hs.")
        Call HEADERS("RESUCCX", "")
        Call HEADERS("RESUCCX", "Corresponde a: " + TRIM$(C1$) + " - Generado " + C2$ + " - " + C3$ + " hs.")
        Call GRARGEN("CONCIVEN", TRIM$(C1$))
        '
80      For I% = 0 To CARESPA%
            If TRIM$(ARESPA$(I%)) <> "" Then
                '
                Call HEADERS(ARESPA$(I%), "")
                Call HEADERS(ARESPA$(I%), "Corresponde a: " + TRIM$(C1$) + " - Generado " + C2$ + " - " + C3$ + " hs.")
                '
            End If
        Next I%
        '
        Call HEADERS("SUDIVEN", "")
        Call HEADERS("SUDIVEN", "Corresponde a: " + TRIM$(C1$))
        '
        Call HEADERS("SUIVAVE", "")
        Call HEADERS("SUIVAVE", "Corresponde a: " + TRIM$(C1$))
        '
    End If
    '
99  Screen.MousePointer = 1
    '
End Sub
'

Sub CALCOMIS()
    '
    For U& = 1 To ULTREG&("VENDEDOR")
        X$ = REGLEIDO$("VENDEDOR", U&)
        VENDE% = Asc(Left$(X$, 1))
        If VENDE% > 0 Then
            If VENDE% <= 256 Then
                If PORCO(VENDE%) < 0.005 Then
                    PORCO(VENDE%) = CVS(Mid$(X$, 44, 4)) + CVS(Mid$(X$, 48, 4))
                End If
            End If
        End If
    Next U&
    Call CIERRARCH("VENDEDOR")
    '
    Call DESHASFECHA(DES%, HAS%, PERIO$)
    If PERIO$ = "" Then Exit Sub
    FORLICOMIS.MESAN = PERIO$
    '
    For U& = 1 To ULTREG&("DETCOBV")
        X$ = REGLEIDO$("DETCOBV", U&)
        NRECI&(U&) = CVS(Mid$(X$, 5, 4))
        REGFA$(U&) = ""
    Next U&
    '
    CARECI& = U&
    '
    FIN& = ULTREG&("APLICOB")
    FORLICOMIS.Check1.Value = 2
    '
    FIN& = ULTREG("APLICOB")
    For U& = 1 To FIN&
        Call TRACE(20, U&, FIN&)
        X$ = REGLEIDO$("APLICOB", U&)
        FEDOC% = CVI(Left$(X$, 2))
        If FEDOC% >= DES% Then
           If FEDOC% <= HAS% Then
              NRO& = Val(Mid$(X$, 15, 8))
              If NRO& > 0 Then
                 If Mid$(X$, 55, 2) = "FC" Then
                    For UU& = 1 To CARECI&
                       If NRECI&(UU&) = NRO& Then
                          REGFA$(UU&) = REGFA$(UU&) + MKS$(U&)
                          GoTo 19
                       End If
                    Next UU&
                 End If
              End If
           End If
        End If
19  Next U&
    FORLICOMIS.Check1.Value = 1
    '
    FIN& = CARECI&
    FORLICOMIS.Check2.Value = 2
    '
    CAFAC% = 0
    Call BLANARCH("LICOMIN")
    For U& = 1 To CARECI&
        Call TRACE(20, U&, CARECI&)
        X$ = REGLEIDO$("DETCOBV", U&)
        NRO& = CVS(Mid$(X$, 5, 4))
        FEDOC% = CVI(Mid$(X$, 3, 2))
        NC% = CVI(Mid$(X$, 9, 2))
        ICOB# = CVD(Mid$(X$, 11, 8))
        INOAPLI# = ICOB#
        '
        For V& = 1 To Len(REGFA$(U&)) Step 4
            W& = CVS(Mid$(REGFA$(U&), V&, 4))
            If W& > 0 Then
               If W& <= ULTREG&("APLICOB") Then
                  Y$ = REGLEIDO$("APLICOB", W&)
                  IAPLI# = CVD(Mid$(Y$, 23, 8))
                  If IAPLI# > 0 Then
                     NROFAC& = Val(Mid$(Y$, 63, 8))
                     INOAPLI# = INOAPLI# - IAPLI#
                     '
                     Z$ = REGBLAN$("LICOMIN")
                     Call REPLA(Z$, MKI$(NC%), 1, 2)
                     Call REPLA(Z$, MKS$(NRO&), 3, 4)
                     Call REPLA(Z$, MKI$(FEDOC%), 7, 2)
                     Call REPLA(Z$, MKD$(ICOB#), 9, 8)
                     Call REPLA(Z$, MKI$(NC%), 17, 2)
                     Call REPLA(Z$, Chr$(VENDCLI%(NC%)), 19, 2)
                     '
                     Call REPLA(Z$, MKS$(NROFAC&), 21, 4)
                     Call REPLA(Z$, MKD$(IAPLI#), 25, 8)
                     Call REGAPP("LICOMIN", Z$)
                     '
                     CAFAC% = CAFAC% + 1
                     NCV%(CAFAC%) = NC%
                     NROV&(CAFAC%) = NROFAC&
                     IDET#(CAFAC%) = 0
                     IFAC#(CAFAC%) = 0
                  End If
               End If
            End If
        Next V&
        '
        If Abs(INOAPLI#) >= 0.005 Then
            Z$ = REGBLAN$("LICOMIN")
            Call REPLA(Z$, MKI$(NC%), 1, 2)
            Call REPLA(Z$, MKS$(NRO&), 3, 4)
            Call REPLA(Z$, MKI$(FEDOC%), 7, 2)
            Call REPLA(Z$, MKD$(ICOB#), 9, 8)
            Call REPLA(Z$, MKI$(NC%), 17, 2)
            Call REPLA(Z$, Chr$(VENDCLI%(NC%)), 19, 2)
            '
            Call REPLA(Z$, MKS$(0), 21, 4)
            Call REPLA(Z$, MKD$(INOAPLI#), 25, 8)
            Call REGAPP("LICOMIN", Z$)
            '
            CAFAC% = CAFAC% + 1
            NCV%(CAFAC%) = NC%
            NROV&(CAFAC%) = 0
            IDET#(CAFAC%) = 0
            IFAC#(CAFAC%) = 0
            '
        End If
        '
    Next U&
    FORLICOMIS.Check2.Value = 1
    '
    ARMOVI$ = "MOVISTO"
    '
113 FIN& = ULTREG&(ARMOVI$)
    FORLICOMIS.Check3.Value = 2
    '
    For IL& = 1 To FIN&
        Call TRACE(20, IL&, FIN&)
        X$ = REGLEIDO$(ARMOVI$, IL&)
        COMO$ = UCase$(Mid$(X$, 33, 2))
        If COMO$ = "FC" Then
          NRO1& = Val(Mid$(X$, 36, 7))
          If NRO1& > 0 Then
             NC1% = CVI(Mid$(X$, 84, 2))
             For KI% = 1 To CAFAC%
                If NROV&(KI%) = NRO1& Then
                   If NCV%(KI%) = NC1% Then
                      CAN# = CVD(Mid$(X$, 96, 8)) - CVD(Mid$(X$, 88, 8))
                      PRE# = CVD(Mid$(X$, 73, 8))
                      TOT# = PRE# * CAN#
                      CI% = CVI(Mid$(X$, 3, 2))
                      FACAPRE% = 100
                      '
                      If CI% > 0 Then
                         If CI% <= NUMAS% Then
                            REGMA& = CI% + 1
                            Y$ = REGLEIDO$("MASTER", REGMA&)
                            FACAPRE% = CVI(Mid$(Y$, 127, 2))
                            If FACAPRE% < 1 Then
30                             VDEF$ = MKI$(CI%) + MKI$(100)
                               If FACAPRE% >= 1 Then
                                   Call REPLA(VDEF$, MKI$(FACAPRE%), 3, 2)
                               End If
                               OBX$ = OBJPLA$("FACAPRE", VDEF$)
                               If OBX$ = "" Then
                                  Call CIERRARCH("*.*")
                                  Exit Sub
                               End If
                               FACAPRE1% = CVI(Mid$(OBX$, 3, 2))
                               If FACAPRE1% < 1 Or FACAPRE1% > 32767 Then GoTo 30
                               If REGMA& > 1 Then
                                  Y$ = REGLEIDO$("MASTER", REGMA&)
                                  Call REPLA(Y$, MKI$(FACAPRE1%), 127, 2)
                                  Call GRAREG("MASTER", Y$, REGMA&)
                               End If
                               FACAPRE% = FACAPRE1%
                            End If
                         End If
                      End If
                      '
                      TOPONDE# = TOT# * FACAPRE% / 100
                      '
                      IDET#(KI%) = IDET#(KI%) + TOT#
                      IFAC#(KI%) = IFAC#(KI%) + TOPONDE#
                      '
                   End If
               End If
             Next KI%
          End If
          '
        End If
    Next IL&
    '
    If ARMOVI$ = "MOVISTO" Then
      If EXISTE%(LUDAT$ + "MOVISVIE.DAT") > 0 Then
        XXX$ = REGLEIDO$("MOVISTO", 1)
        FF% = CVI(Left$(XXX$, 2))
        If FF% >= DES% Then
          ARMOVI$ = "MOVISVIE"
          GoTo 113
        End If
      End If
    End If
    '
    FORLICOMIS.Check3.Value = 1
    '
    FIN& = CAFAC%
    FORLICOMIS.Check4.Value = 2
    '
    FIN& = CAFAC%
    For IL& = 1 To CAFAC%
        Call TRACE(20, IL&, FIN&)
        X$ = REGLEIDO$("LICOMIN", IL&)
        COEF = 1
        If IDET#(IL&) > 0 Then
            COEF = IFAC#(IL&) / IDET#(IL&)
        End If
        Call REPLA(X$, MKS$(COEF), 41, 4)
        '
        VENDE% = Asc(Mid$(X$, 19, 1))
        If VENDE% >= 0 Then
           If VENDE% <= 256 Then
              Call REPLA(X$, MKS$(PORCO(VENDE%)), 45, 4)
              ICOMI# = (Int(CVD(Mid$(X$, 25, 8)) * COEF * PORCO(VENDE%) + 0.5)) / 100
              Call REPLA(X$, MKD$(ICOMI#), 49, 8)
              TOCOMI#(VENDE%) = TOCOMI#(VENDE%) + ICOMI#
              TOCOB#(VENDE%) = TOCOB#(VENDE%) + CVD(Mid$(X$, 25, 8))
           End If
        End If
        '
        Call GRAREG("LICOMIN", X$, IL&)
        '
    Next IL&
    '
    Call CIERRARCH("LICOMIN")
    FORLICOMIS.List1.Clear
    For VENDE% = 0 To 255
      If TOCOB#(VENDE%) > 0 Then
        If VENDE% = 0 Then
             NOVEN$ = "DIRECCION"
           Else
             NOVEN$ = ECOARCH$("VENDEDOR", Chr$(VENDE%))
        End If
        TXT$ = Space$(48)
        Call REPLA(TXT$, CSTRING$(MKI$(VENDE%), 1, 3, 0, 2), 1, 4)
        Call REPLA(TXT$, NOVEN$, 6, 17)
        Call REPLA(TXT$, CSTRING$(MKD$(TOCOMI#(VENDE%)), 4, 10, 2, 2), 23, 10)
        FORLICOMIS.List1.AddItem TXT$
      End If
    Next VENDE%
    FORLICOMIS.Check4.Value = 1
    '
    If FORLICOMIS.List1.ListCount > 0 Then
      FORLICOMIS.List1.ListIndex = 0
    End If
    FORLICOMIS.List1.Refresh
    Command3.Enabled = True
    '
End Sub
'

Sub LICOMIS()
    '
    Screen.MousePointer = 11
    '
    YALIQUI% = 0: CUECOMI% = 0
    Dim CVX%(128), COVEN(128), COCOB(128)
    Dim PORCO(256), TOCOMI#(256), TOCOB#(256)
    Dim REGICLI&(32767), VENDECLI%(32767)
    '
    If EXISTE%(LUDAT$ + "CUECOMIS.RFS") Then
      CUECOMI% = 1
      Screen.MousePointer = 11
      Call BLOQARCH("CUECOMIS")
      FIN& = ULTREG&("CUECOMIS")
      JJ& = 0
      For U& = 1 To FIN&
        XX$ = REGLEIDO$("CUECOMIS", U&)
        FECO% = CVI(Mid$(XX$, 2, 2))
        TIMX% = Asc(Mid$(XX$, 4, 1))
        If FECO% < DES% Or FECO% > HAS% Or TIMX% = 2 Then
          JJ& = JJ& + 1
          Call GRAREG("CUECOMIS", XX$, JJ&)
        End If
      Next U&
      Call SETULTREG("CUECOMIS", JJ&)
      Call CIERRARCH("CUECOMIS")
      Screen.MousePointer = 1
    End If
    '
    For I& = 1 To 256
       SUPED#(I&) = 0
       COMPED#(I&) = 0
       SUCOB#(I&) = 0
       COMCOB#(I&) = 0
    Next I&
    '
    NUVE% = 0
    YAGERE% = 0
    For I& = 1 To ULTREG&("VENDEDOR")
        X$ = REGLEIDO$("VENDEDOR", I&)
        CV% = Asc(Left$(X$, 1))
        If CV% > 0 Then
            NUVE% = NUVE% + 1
            CVX%(NUVE%) = CV%
            COVEN(NUVE%) = CVS(Mid$(X$, 44, 4))
            COCOB(NUVE%) = CVS(Mid$(X$, 48, 4))
            If CV% = 99 Then YAGERE% = 1
        End If
    Next I&
    '
    If YAGERE% = 0 Then
        X$ = REGBLAN$("VENDEDOR")
        Call REPLA(X$, Chr$(99), 1, 1)
        Call REPLA(X$, "GERENCIA", 2, 30)
        Call REPLA(X$, "GERENCIA", 32, 12)
        Call REPLA(X$, MKS$(0), 44, 4)
        Call REPLA(X$, MKS$(0), 48, 4)
        Call REGAPP("VENDEDOR", X$)
        '
        NUVE% = NUVE% + 1
        CVX%(NUVE%) = 99
        COVEN(NUVE%) = CVS(Mid$(X$, 44, 4))
        COCOB(NUVE%) = CVS(Mid$(X$, 48, 4))
    End If
    Call CIERRARCH("VENDEDOR")
    '
50  Call DESHASFECHA(DES%, HAS%, PERIO$)
    If PERIO$ = "" Then Exit Sub
    FORLICOMIS.MESAN = PERIO$
    '
    Call HEADERS("LICOMIX", "")
    Call HEADERS("LICOMIX", "Corresponde a: " + CORRESP$)
    '
    FIN& = ULTREG&("DEUDOR1")
    URE2& = ULTREG&("DEUDOR2")
    For U& = 1 To FIN&
      Call TRACE(20, U&, FIN&)
      X$ = REGLEIDO$("DEUDOR1", U&)
      If U& <= URE2& Then
        X$ = X$ + REGLEIDO$("DEUDOR2", U&)
        NC% = CVI(Left$(X$, 2))
        If NC% > 0 Then
          If NC% <= 32767 Then
            If REGICLI&(NC%) < 1 Then
              REGICLI&(NC%) = U&
              VENDECLI%(NC%) = Asc(Mid$(X$, 152, 1))
            End If
          End If
        End If
      End If
    Next U&
    '
    'For U& = 1 To ULTREG&("HISTOVEN")
    '  X$ = REGLEIDO$("HISTOVEN", U&)
    '  NC% = CVI(Left$(X$, 2))
    '  VENDECLI%(NC%) = CVI(Mid$(X$, 3, 2))
    'Next U&
    'Call CIERRARCH("HISTOVEN")
    '
    MODO% = 0
    If FORLICOMIS.Option3.Value = True Then MODO% = 1
    If FORLICOMIS.Option4.Value = True Then MODO% = 2
    If FORLICOMIS.Option5.Value = True Then MODO% = 3
    '
    J& = 0: JLIBRO& = 0
    On MODO% GoTo 100, 200, 300
    GoTo 999
    '
100 Call HEADERS("LICOMIX", "Por: Pedidos")
    '
    FIN& = ULTREG&("PEDDETA")
    '
    INIPE& = 0
    For I& = 1 To FIN&
        Call TRACE(20, I& - INIPE&, FIN& - INIPE&)
        X$ = REGLEIDO$("PEDDETA", I&)
        FF% = CVI(Mid$(X$, 5, 2))
        If FF% >= DES% Then
            If FF% <= HAS% Then
                NRO& = CVS(Left$(X$, 4))
                If NRO& > 0 Then
If NRO& >= 44357 And NRO& <= 44370 Then GoTo 129
If NRO& >= 44428 And NRO& <= 44435 Then GoTo 129
If NRO& = 44440 Then GoTo 129
                    NC% = CVI(Mid$(X$, 7, 2))
                    If NC% > 0 Then
                       VENDE% = 0 ' VENDE%=ASC(MID$(X$,30,1))
                       If VENDE% < 1 Then VENDE% = VENDECLI%(NC%)
                       If VENDE% < 1 Then VENDE% = 99
                       '
                       CI% = CVI(Mid$(X$, 33, 2))
                       CAN = CVS(Mid$(X$, 79, 4))
                       If REGICLI&(NC%) > 0 Then
                          TTTT$ = REGLEIDO$("DEUDOR1", REGICLI&(NC%))
                       End If
                       DESCU = PORDESCLI(NC%)
                       PREUNI = CVS(Mid$(X$, 83, 4)) * (1 - DESCU / 100)
                       ITOT# = (Int(100 * CDbl(PREUNI) * CAN + 0.5)) / 100
                       TIDOC% = 8
                       '
                       For K% = 1 To NUVE%
                          If CVX%(K%) = VENDE% Then
                             PORCOM = COVEN(K%)
' puesto 3/8/00 para reiba por pedido de sabo - provisorio
                             If NC% = 1271 Then PORCOM = 1
                             GoTo 120
                          End If
                       Next K%
                       PORCOM = 0
                       '
120                    ICOM# = (Int(ITOT# * PORCOM + 0.5)) / 100
                       NUSUCUVE% = 0
                       GoSub 600 '
                       If INIPE& = 0 Then INIPE& = I&
                       '
                    End If
                End If
            End If
        End If
129 Next I&
    '
    GoTo 300                 ' PARA INCLUIR COBRANZA
    GoSub 610
    Call FILESORT("LICOMIS", "", 3, 5, "ASC")
    YALIQUI% = 1
    GoTo 999
    '
    '
200 Call HEADERS("LICOMIX", "Por: Facturacion")
    FIN& = ULTREG&("FACTUR")
    '
    For I& = 1 To FIN&
        Call TRACE(20, I&, FIN&)
        X$ = REGLEIDO$("FACTUR", I&)
        FF% = CVI(Mid$(X$, 7, 2))
        If FF% >= DES% Then
            If FF% <= HAS% Then
                NRO& = CVS(Mid$(X$, 3, 4))
                If NRO& > 0 Then
                    NC% = CVI(Mid$(X$, 1, 2))
                    If NC% > 0 Then
                       VENDE% = VENDECLI%(NC%)
                       If VENDE% < 1 Then VENDE% = 99
                       '
                       CI% = 0
                       CAN = 0
                       DESCU = 0
                       PREUNI = 0
                       ITOT# = CVD(Mid$(X$, 15, 8)) + CVD(Mid$(X$, 49, 8))
                       TIDOC% = Asc(Mid$(X$, 13, 1))
                       VADOC% = Asc(Mid$(X$, 14, 1))
                       If TIDOC% >= 4 Then
                           ITOT# = (-1) * Abs(ITOT#)
                       End If
                       '
                       NUSUCUVE% = CVI(Mid$(X$, 47, 2))
                       If TIDOC% > 0 Then
                          For K% = 1 To NUVE%
                             If CVX%(K%) = VENDE% Then
                                PORCOM = COVEN(K%)
                                GoTo 220
                             End If
                          Next K%
                          PORCOM = 0
                          '
220                       ICOM# = (Int(ITOT# * PORCOM + 0.5)) / 100
                          GoSub 600
                          '
                       End If
                    End If
                End If
            End If
        End If
    Next I&
    '
    GoSub 610
    Call FILESORT("LICOMIS", "", 3, 5, "ASC")
    YALIQUI% = 1
    GoTo 999
    '
    '
300 Call HEADERS("LICOMIX", "Por: Cobranza")
    FIN& = ULTREG&("COBRAN")
    '
    For I& = 1 To FIN&
        Call TRACE(20, I&, FIN&)
        X$ = REGLEIDO$("COBRAN", I&)
        FF% = CVI(Mid$(X$, 7, 2))
        If FF% >= DES% Then
            If FF% <= HAS% Then
                NRO& = CVS(Mid$(X$, 3, 4))
                If NRO& > 0 Then
                    NC% = CVI(Mid$(X$, 1, 2))
                    If NC% > 0 Then
                       VENDE% = VENDECLI%(NC%)
                       If VENDE% < 1 Then VENDE% = 99
                       '
                       CI% = 0
                       CAN = 0
                       DESCU = 0
                       PREUNI = 0
                       ITOT# = CVD(Mid$(X$, 15, 8))
                       TIDOC% = 4
                       '
                       NUSUCUVE% = 0
                       If TIDOC% > 0 Then
                          For K% = 1 To NUVE%
                             If CVX%(K%) = VENDE% Then
                                PORCOM = COCOB(K%)
                                GoTo 320
                             End If
                          Next K%
                          PORCOM = 0
                          '
320                       ' NO CONSIDERAR EL IVA
                          PVCL% = PIVACLI%(NC%)
                          DIVI = 1 + ALIVA / 100
                          If PVCL% = 2 Then DIVI = 1 + 1.5 * ALIVA / 100
                          If PVCL% = 5 Or PVCL% = 7 Then DIVI = 1
                          ICOM# = (Int(ITOT# * PORCOM / DIVI + 0.5)) / 100
                          GoSub 600
                          '
                       End If
                    End If
                End If
            End If
        End If
    Next I&
    '
    GoSub 610
    Call FILESORT("LICOMIS", "", 3, 5, "ASC")
    YALIQUI% = 1
    GoTo 999
    '
    '
999 If YALIQUI% = 1 Then
        FORLICOMIS.List1.Clear
        For VENDE% = 0 To 255
            If VENDE% = 0 Then
                NOVEN$ = "DIRECCION"
              Else
                NOVEN$ = ECOARCH$("VENDEDOR", Chr$(VENDE%))
            End If
            If TRIM$(NOVEN$) <> "" Then
                TXT$ = Space$(48)
                Call REPLA(TXT$, CSTRING$(MKI$(VENDE%), 1, 3, 0, 2), 1, 4)
                Call REPLA(TXT$, NOVEN$, 6, 17)
                Call REPLA(TXT$, CSTRING$(MKD$(TOCOMI#(VENDE%)), 4, 10, 2, 2), 23, 10)
                FORLICOMIS.List1.AddItem TXT$
            End If
        Next VENDE%
        FORLICOMIS.Check4.Value = 1
        '
        If FORLICOMIS.List1.ListCount > 0 Then
            FORLICOMIS.List1.ListIndex = 0
        End If
        FORLICOMIS.List1.Refresh
        '
    End If
    '
    Screen.MousePointer = 1
    Exit Sub
    '
    '
600 Y$ = REGBLAN$("LICOMIS")
    Call REPLA(Y$, Chr$(VENDE%), 1, 1)
    Call REPLA(Y$, MKI$(NC%), 2, 2)
    Call REPLA(Y$, MKI$(FF%), 4, 2)
    Call REPLA(Y$, Chr$(TIDOC%), 6, 1)
    Call REPLA(Y$, MKS$(NRO&), 7, 4)
    Call REPLA(Y$, MKI$(CI%), 11, 2)
    Call REPLA(Y$, MKS$(PREUNI), 13, 4)
    Call REPLA(Y$, MKS$(CAN), 17, 4)
    Call REPLA(Y$, MKD$(ITOT#), 21, 8)
    Call REPLA(Y$, MKS$(PORCOM), 29, 4)
    Call REPLA(Y$, MKD$(ICOM#), 33, 8)
    '
'If NRO& = 12562 Then
'AAA = BBB
'End If
    J& = J& + 1
    Call GRAREG("LICOMIS", Y$, J&)
    TOCOMI#(VENDE%) = TOCOMI#(VENDE%) + ICOM#
    '
    If CUECOMI% = 1 Then
      YCM$ = REGBLAN$("CUECOMIS")
      Call REPLA(YCM$, Chr$(VENDE%), 1, 1)
      Call REPLA(YCM$, MKI$(FF%), 2, 2)
      Call REPLA(YCM$, Chr$(1), 4, 1)
      Call REPLA(YCM$, MKS$(NRO&), 5, 4)
      Call REPLA(YCM$, RASOCLI$(NC%), 9, 32)
      Call REPLA(YCM$, MKD$(ICOM#), 41, 8)
      Call REGAPP("CUECOMIS", YCM$)
    End If
    '
    If TIDOC% <> TIDOCA% Or NRO& <> NROA& Or NC% <> NCAA% Then
       MOVENTA# = 0
       ICOMIS# = 0
       JLIBRO& = JLIBRO& + 1
       TIDOCA% = TIDOC%: NROA& = NRO&: NCAA% = NC%
    End If
    '
    Y$ = REGBLAN$("LIBROVIA")
    FEX = FF%: FEY$ = FECHATEX$(FEX)
    ANIO$ = Mid$(FEY$, 7, 2)
    MESI$ = Mid$(FEY$, 4, 2)
    DIAI$ = Left$(FEY$, 2)
    If Val(ANIO$) > 55 Then
         ANIO$ = "19" + ANIO$
       Else
         ANIO$ = "20" + ANIO$
    End If
    Call REPLA(Y$, ANIO$, 1, 4)
    Call REPLA(Y$, MESI$, 5, 2)
    Call REPLA(Y$, DIAI$, 7, 2)
    '
    If TIDOC% = 8 Then
         DOCU$ = "NP-" + TRIM$(Str$(NRO&))
         SUPED#(VENDE%) = SUPED#(VENDE%) + ITOT#
         COMPED#(VENDE%) = COMPED#(VENDE%) + ICOM#
       ElseIf TIDOC% = 4 Then
         DOCU$ = "RC-" + TRIM$(Str$(NRO&))
         SUCOB#(VENDE%) = SUCOB#(VENDE%) + ITOT#
         COMCOB#(VENDE%) = COMCOB#(VENDE%) + ICOM#
       Else
         DOCU$ = FORMDOC$(NRO&, TIDOC%, VADOC%, NUSUCUVE%)
    End If
    Call REPLA(Y$, DOCU$, 9, 18)
    '
    If REGICLI&(NC%) > 0 Then
       TTTT$ = REGLEIDO$("DEUDOR1", REGICLI&(NC%))
    End If
    Call REPLA(Y$, RASOCLI$(NC%), 27, 30)
    Call REPLA(Y$, "Acep.", 57, 12)
    Call REPLA(Y$, MKI$(FF%), 69, 2)
    '
    MOVENTA# = MOVENTA# + ITOT#
    ICOMIS# = ICOMIS# + ICOM#
    PORCOMIS = PORCOM
    If Abs(MOVENTA#) > 0.05 Then
       If PORCOMIS < 0.1 Then
          PORCOMIS = 100 * ICOMIS# / MOVENTA#
       End If
    End If
    '
    Call REPLA(Y$, MKS$(PORCOMIS), 71, 4)
    Call REPLA(Y$, MKD$(MOVENTA#), 75, 8)
    Call REPLA(Y$, MKD$(ICOMIS#), 83, 8)
    '
    OBSER$ = Space$(38)
    Call REPLA(Y$, OBSER$, 91, 36)
    Call REPLA(Y$, MKI$(VENDE%), 127, 2)
    '
    Call GRAREG("LIBROVIA", Y$, JLIBRO&)
    '
Return
    '
    '
610 For I& = J& + 1 To ULTREG&("LICOMIS")
        Call GRAREG("LICOMIS", String$(64, 0), I&)
    Next I&
    Call CIERRARCH("LICOMIS")
    If CUECOMI% = 1 Then
      Call CIERRARCH("CUECOMIS")
    End If
    '
    Call SETULTREG("LIBROVIA", JLIBRO&)
    Call CIERRARCH("LIBROVIA")
    '
Return
'
'
End Sub
'
Sub PRILICOMIS()
    '
    MODOCAL% = 0
    If FORLICOMIS.Option2.Value = True Then MODOCAL% = 1
    If FORLICOMIS.Option3.Value = True Then MODOCAL% = 2: MODOCA$ = "Por Pedidos Netos de Descuento"
    If FORLICOMIS.Option4.Value = True Then MODOCAL% = 3: MODOCA$ = "Por Facturacion"
    If FORLICOMIS.Option5.Value = True Then MODOCAL% = 4: MODOCA$ = "Por Cobranza Neta de IVA"
    If MODOCAL% < 1 Then
        Exit Sub
    End If
    '
    If FORLICOMIS.List1.ListIndex >= 0 Then
        TXT$ = FORLICOMIS.List1.TEXT
        VENDE% = Val(Left$(TXT$, 3))
        NOVEN$ = "Clientes de Direccion"
        If VENDE% > 0 Then
            NOVEN$ = ECOARCH$("VENDEDOR", Chr$(VENDE%))
        End If
        '
        CORRESP$ = FORLICOMIS.MESAN.TEXT
        '
        If MODOCAL% = 1 Then
            Call COPYSTRU("LICOMIN", "LICOMIV")
            Call BLANARCH("LICOMIV")
            Call HEADERS("LICOMIV", "")
            Call HEADERS("LICOMIV", "Corresponde a: " + CORRESP$)
            Call HEADERS("LICOMIV", "Vendedor: " + NOVEN$)
            '
            FIN& = ULTREG&("LICOMIN")
            '
            For IL& = 1 To FIN&
                Call TRACE(20, IL&, FIN&)
                X$ = REGLEIDO$("LICOMIN", IL&)
                If Asc(Mid$(X$, 19, 1)) = VENDE% Then
                    Call REGAPP("LICOMIV", X$)
                End If
            Next IL&
            '
            Call CIERRARCH("*.*")
            Call CONECRUN("*LICOMIV", "Listado de Comisiones Liquidadas")
            Exit Sub
            '
          Else
            J& = 0
            Call HEADERS("LICOMIX", "")
            
            Call HEADERS("LICOMIX", "Corresponde a: " + CORRESP$)
            Call HEADERS("LICOMIX", "Modalidad: " + MODOCA$)
            Call HEADERS("LICOMIX", "Vendedor: " + NOVEN$)
            '
            FIN& = ULTREG&("LICOMIS")
            '
            For IL& = 1 To FIN&
                Call TRACE(20, IL&, FIN&)
                X$ = REGLEIDO$("LICOMIS", IL&)
                If Asc(Mid$(X$, 1, 1)) = VENDE% Then
                    J& = J& + 1
                    Call GRAREG("LICOMIX", X$, J&)
                End If
            Next IL&
            '
            Call SETULTREG("LICOMIX", J&)
            Call CIERRARCH("*.*")
            Call CONECRUN("*LISLICO", "Listado de Comisiones Liquidadas")
            Exit Sub
        End If
        '
      Else
        '
        Call MENSERR(24, "Falta Seleccionar Vendedor")
        '
    End If
    '
End Sub
'
Sub PRILIVIAJAN()
    '
    MODX% = COMALTER%("Libro de Viajantes - Opciones Disponibles:\\Imprimir Libro\Pre-Numerar Hojas\Setup Formulario")
    If MODX% < 1 Or MODX% > 3 Then Exit Sub
    '
    If MODX% = 3 Then
      OPLIVIAJ.Show 1
      Exit Sub
    End If
    '
    If MODX% = 2 Then
      FORIPRE$ = TRIM$(CONTROL$("", "FORLIVIA"))
      If FORIPRE$ <> "" Then
        PRIHOJA$ = InputBox$("Numera Desde Hoja Numero", "Numeracion Libro de Viajantes", "1")
        CANHOJA$ = InputBox$("Cantidad de Hojas a Emitir", "Pre-Numeracion Libro de Viajantes", "100")
        '
        PHX& = Val(PRIHOJA$)
        CHX& = Val(CANHOJA$)
        If PHX& > 0 Then
          If CHX& > 0 Then
            For NHX& = PHX& To PHX& + CHX& - 1
              NUMEDOC$ = TRIM$(Str$(NHX&))
              TIPODOC$ = "Pre-Numeracion Libro de Viajantes"
              '
              CANCELPRINT% = 0
              NOCONFIRM% = 1
              '
              Call PRINTDOC("FORLIVIA")   ' Remito de Facturacion
            Next NHX&
          End If
        End If
      End If
      Exit Sub
    End If
    '
    Dim HDRX As String * 256
    If FORLICOMIS.List1.ListIndex >= 0 Then
        TXT$ = FORLICOMIS.List1.TEXT
        If Abs(Val(Mid$(TXT$, 23, 10))) < 0.005 Then
           Call MENSERR(24, "Imposible Emitir Libro de Viajantes\para el Vendedor Elegido.")
           Exit Sub
        End If
        VENDE% = Val(Left$(TXT$, 3))
        NOVEN$ = "Clientes de Direccion"
        If VENDE% > 0 Then
            NOVEN$ = ECOARCH$("VENDEDOR", Chr$(VENDE%))
        End If
        '
        CORRESP$ = AJUSTI$(FORLICOMIS.MESAN.TEXT, 12)
        '
        J& = 0
        Call BLANARCH("!DAVENDE")
        For U& = 1 To ULTREG&("DAVENDE")
           X$ = REGLEIDO$("DAVENDE", U&)
           If CVI(Left$(X$, 2)) = VENDE% Then
              Call REGAPP("!DAVENDE", X$)
              J& = J& + 1
           End If
        Next U&
        If J& < 1 Then
           X$ = MKI$(VENDE%) + "Nombre: ............................................... Fecha Ingreso: ..........................................."
           Call REPLA(X$, TRIM$(NOVEN$) + ".......................", 11, 36)
           Call REGAPP("!DAVENDE", X$)
           X$ = MKI$(VENDE%) + "Zona en que actua: ..............................................................................................."
           Call REGAPP("!DAVENDE", X$)
           X$ = MKI$(VENDE%) + ".................................................................................................................."
           Call REGAPP("!DAVENDE", X$)
           X$ = MKI$(VENDE%) + "Sueldo:..........................................................................................................."
           Call REGAPP("!DAVENDE", X$)
           X$ = MKI$(VENDE%) + "Viatico: ........................................................................................................."
           Call REGAPP("!DAVENDE", X$)
           X$ = MKI$(VENDE%) + "Comisiones %: ...................................................................................................."
           Call REGAPP("!DAVENDE", X$)
           X$ = MKI$(VENDE%) + "Especificaciones de mercaderias que no puede vender: ..............................................................................................."
           Call REGAPP("!DAVENDE", X$)
           X$ = MKI$(VENDE%) + "..............................................................................................."
           Call REGAPP("!DAVENDE", X$)
            X$ = MKI$(VENDE%) + "Comisiones %: ..............................................................................................."
           Call REGAPP("!DAVENDE", X$)
           X$ = MKI$(VENDE%) + "..............................................................................................."
           Call REGAPP("!DAVENDE", X$)
        End If
        Call CIERRARCH("DAVENDE")
        Call CIERRARCH("!DAVENDE")
        '
        VTB% = VENTABU%("DAVENDE")
        If VTB% < 0 Then Exit Sub
        '
        On Error Resume Next
        Kill LUCOM$ + "LIVIALIS.HDR"
        On Error GoTo 0
        NX% = FILEOPEN%(LUCOM$ + "LIVIALIS.HDR", 0, 256)
        J& = 0: K& = 0
        For U& = 1 To ULTREG&("DAVENDE")
           X$ = REGLEIDO$("DAVENDE", U&)
           If CVI(Left$(X$, 2)) <> VENDE% Then
              J& = J& + 1
              Call GRAREG("DAVENDE", X$, J&)
           End If
        Next U&
        For U& = 1 To ULTREG&("!DAVENDE")
           X$ = REGLEIDO$("!DAVENDE", U&)
           If CVI(Left$(X$, 2)) = VENDE% Then
              J& = J& + 1
              Call GRAREG("DAVENDE", X$, J&)
              'If U& = 1 Then
              '    Call REPLA(X$, "Hoja Nro:#####", 126, 14)
              'End If
              If TRIM$(Mid$(X$, 3)) <> "" Then
                   LSet HDRX$ = TRIM$(Mid$(X$, 3)) + String$(256, ".")
                 Else
                   LSet HDRX$ = TRIM$(Mid$(X$, 3)) + Space$(256)
              End If
              K& = K& + 1
              Put #NX%, K&, HDRX$
           End If
        Next U&
        Call SETULTREG("DAVENDE", J&)
        Call CIERRARCH("DAVENDE")
        Close #NX%
        '
        On Error Resume Next
        Kill LUCOM$ + "LIVIALIS.SBT"
        On Error GoTo 0
        NX% = FILEOPEN%(LUCOM$ + "LIVIALIS.SBT", 0, 256)
        TTXX$ = "¦    ¦  ¦  ¦SUBTOTAL PEDIDOS ¦                               ¦         ¦        ¦        ¦            ¦            ¦                    ¦"
        Call REPLA(TTXX$, NUMSTRI$(SUPED#(VENDE%), 12, 2, 0), 91, 12)
        If InStr(UCase$(EMPREAC$), "SABO") < 1 Then
          Call REPLA(TTXX$, NUMSTRI$(COMPED#(VENDE%), 12, 2, 0), 104, 12)
        End If
        LSet HDRX$ = TTXX$
        Put #NX%, 1, HDRX$
        TTXX$ = "¦    ¦  ¦  ¦SUBTOTAL COBRANZA¦                               ¦         ¦        ¦        ¦            ¦            ¦                    ¦"
        Call REPLA(TTXX$, NUMSTRI$(SUCOB#(VENDE%), 12, 2, 0), 91, 12)
        If InStr(UCase$(EMPREAC$), "SABO") < 1 Then
          Call REPLA(TTXX$, NUMSTRI$(COMCOB#(VENDE%), 12, 2, 0), 104, 12)
        End If
        LSet HDRX$ = TTXX$
        Put #NX%, 2, HDRX$
        Close #NX%
        '
        J& = 0
        Call COPYSTRU("LIBROVIA", "LIVIALIS")
        FIN& = ULTREG&("LIBROVIA")
        For U& = 1 To FIN&
           Call TRACE(20, U&, FIN&)
           Y$ = REGLEIDO$("LIBROVIA", U&)
           If CVI(Mid$(Y$, 127, 2)) = VENDE% Then
              If InStr(UCase$(EMPREAC$), "SABO") > 0 Then
                Call REPLA(Y$, MKS$(0), 71, 4)
                Call REPLA(Y$, MKD$(0), 83, 8)
              End If
              J& = J& + 1
              Call GRAREG("LIVIALIS", Y$, J&)
           End If
        Next U&
        '
        Call SETULTREG("LIVIALIS", J&)
        Call CIERRARCH("LIVIALIS")
        '
        'PRIHOJA$ = InputBox$("Numera Desde Hoja Numero", "Numeracion Libro de Viajantes")
        'Call CONECRUN("*LILIVIA/PRIHOJA=" + PRIHOJA$, "Libro de Viajantes")
        Call CONECRUN("*LILIVIA", "Libro de Viajantes")
        '
      Else
        '
        Call MENSERR(24, "Falta Seleccionar Vendedor")
        '
    End If
    '
End Sub
'
Sub CALREVECLI(DES%, HAS%)
        '
        Dim TOTVENTA#(32767)
        '
        Call BLANARCH("REVECLI")
        Call BLANARCH("REVEPIE")
        IREVE$ = ""
        '
        Screen.MousePointer = 11
        '
        SUMATO# = 0
        ARMOVI$ = "MOVISTO"
        '
113     FIN& = ULTREG&(ARMOVI$)
        '
        For IL& = 1 To FIN&
              Call TRACE(20, IL&, FIN&)
              MS$ = REGLEIDO$(ARMOVI$, IL&)
              FF% = CVI(Left$(MS$, 2))
              If FF% >= DES% And FF% <= HAS% Then
                 COMO$ = UCase$(Mid$(MS$, 33, 2))
                 If TRIM$(COMO$) = "" Then
                    COMO$ = UCase$(Mid$(MS$, 23, 2))
                 End If
                 If COMO$ = "RT" Or COMO$ = "FC" Or COMO$ = "NC" Or COMO$ = "ND" Then
                    NC% = CVI(Mid$(MS$, 84, 2))
                    NRO1& = Val(Mid$(MS$, 36, 7))
                    If NRO1& < 1 Then
                        NRO1& = Val(Mid$(MS$, 26, 7))
                    End If
                    CI% = CVI(Mid$(MS$, 3, 2))
                    CAN# = CVD(Mid$(MS$, 96, 8)) - CVD(Mid$(MS$, 88, 8))
                    PRE# = CVD(Mid$(MS$, 73, 8))
                    GoSub 13700
                 End If
              End If
        Next IL&
        '
        If ARMOVI$ = "MOVISTO" Then
          If EXISTE%(LUDAT$ + "MOVISVIE.DAT") > 0 Then
            XXX$ = REGLEIDO$("MOVISTO", 1)
            FF% = CVI(Left$(XXX$, 2))
            If FF% >= DES% Then
              ARMOVI$ = "MOVISVIE"
              GoTo 113
            End If
          End If
        End If
        '
'        For IL& = 1 To ULTREG&("MATVEN")
'              MS$ = REGLEIDO$("MATVEN", IL&)
'              FF% = CVI(Mid$(MS$, 7, 2))
'              If FF% >= DES% And FF% <= HAS% Then
'                 COMO$ = "FC"
'                 NC% = CVI(Left$(MS$, 2))
'                 NRO1& = CVS(Mid$(MS$, 3, 4))
'                 CI% = CVI(Mid$(MS$, 9, 2))
'                 CAN# = CVS(Mid$(MS$, 19, 4))
'                 PRE# = CVD(Mid$(MS$, 11, 8))
'                 GoSub 13700
'              End If
'        Next IL&
'        '
        Call CIERRARCH("REVECLI")
        Call CIERRARCH("REVEPIE")
        '
        FEX = DES%: PERIO$ = FECHATEX$(FEX)
        FEX = HAS%: PERIO$ = PERIO$ + " - " + FECHATEX$(FEX)
        Call HEADERS("REVECLI", "")
        Call HEADERS("REVECLI", "Periodo: " + PERIO$)
        Call HEADERS("REVEPIE", "")
        Call HEADERS("REVEPIE", "Periodo: " + PERIO$)
        '
        Screen.MousePointer = 1
        MOX% = ALTERNA%("Listado General por Cliente\Listado por Codigo de Producto\Detalle por Vendedor y Linea\Resumen por Vendedor y Linea")
        On MOX% GoTo 201, 202, 203, 203
        Exit Sub
        '
201     Call FINAL("*REVECLI"): Exit Sub
202     Call FINAL("*REVEPIE"): Exit Sub
        '
203     NCMIN% = 32767: NCMAX% = 0
        IXXX$ = LOADFILE$(LUDAT$ + "DEUDOR1.X01")
        '
        For KKL& = 1 To Len(IXXX$) Step 2
          NCMX% = CVI(Mid$(IXXX$, KKL&, 2))
          If NCMX% > 0 Then
            If NCMX% < NCMIN% Then NCMIN% = NCMX%
            If NCMX% > NCMAX% Then NCMAX% = NCMX%
          End If
        Next KKL&
        '
210     VDEF$ = MKI$(NCMIN%) + MKI$(NCMAX%) + Chr$(0) + MKI$(0)
212     OBX$ = OBJPLA$("DHCLIVENGRU", VDEF$)
        If OBX$ = "" Then
          NCL1% = 0: NCL2% = 0
          Exit Sub
        End If
        '
        NCL1% = CVI(Left$(OBX$, 2))
        NCL2% = CVI(Mid$(OBX$, 3, 2))
        If NCL1% < 1 Or NCL2% < 1 Or NCL2% < NCL1% Then
          Call MENSERR(24, "Rango no Válido")
          GoTo 210
        End If
        VDEF$ = OBX$
        TTXX$ = PERIO$
        VENDSEL% = Asc(Mid$(OBX$, 5, 1))
        If VENDSEL% > 0 Then
          If ECOARCH$("VENDEDOR", Chr$(VENDSEL%)) = "" Then
            Call MENSERR(24, "Vendedor no Valido")
            GoTo 212
          End If
        End If
        If VENDSEL% > 0 Then TTXX$ = TTXX$ + " - Vendedor: " + TRIM$(ECOARCH$("VENDEDOR", Chr$(VENDSEL%)))
        '
        GRUPSEL% = CVI(Mid$(OBX$, 6, 2))
        If GRUPSEL% > 0 Then
          If ECOARCH$("GRUCOVEN", MKI$(GRUPSEL%)) = "" Then
            Call MENSERR(24, "Linea no Valida")
            GoTo 212
          End If
        End If
        If GRUPSEL% > 0 Then TTXX$ = TTXX$ + " - Linea: " + TRIM$(ECOARCH$("GRUCOVEN", MKI$(GRUPSEL%)))
        CORRESP$ = "Corresponde a: " + TTXX$
        '
        If MOX% = 3 Then
             JJ& = 0
             Call COPYSTRU("REVECLI", "RANKVDR")
             FIN& = ULTREG&("REVECLI")
             For U& = 1 To FIN&
               Call TRACE(20, U&, FIN&)
               XX$ = REGLEIDO$("REVECLI", U&)
               NC% = CVI(Left$(XX$, 2))
               '\\\OT-7-0032-FG-29/01/07///
               If NC% > 0 And (NC% >= NCL1% And NC% <= NCL2%) Then
                 If VENDSEL% = 0 Or VENDSEL% = VENDCLI%(NC%) Then
                   CIXI% = CVI(Mid$(XX$, 3, 2))
                   If GRUPSEL% = 0 Or GRUPSEL% = GRUCOVEI%(CIXI%) Then
                     JJ& = JJ& + 1
                     Call GRAREG("RANKVDR", XX$, JJ&)
                   End If
                 End If
               End If
               '\\\OT-7-0032-FG-FIN///
             Next U&
             Call SETULTREG("RANKVDR", JJ&)
             Call FILESORT("RANKVDR", "", 1, 3, "ASC")
             Call HEADERS("RANKVDR", "")
             Call HEADERS("RANKVDR", CORRESP$)
             '
             Call FINAL("*RANKVDR")
             GoTo 212
          ElseIf MOX% = 4 Then
             JJ& = 0
             Screen.MousePointer = 11
             J& = 0: TOTAGEN# = 0
             For U& = 0 To 32767
               TOTVENTA#(U&) = 0
             Next U&
             '
             TOTAGEN# = 0
             FIN& = ULTREG&("REVECLI")
             For U& = 1 To FIN&
               Call TRACE(20, U&, FIN&)
               XX$ = REGLEIDO$("REVECLI", U&)
               NC% = CVI(Left$(XX$, 2))
               If NC% > 0 Then
                 If VENDSEL% = 0 Or VENDSEL% = VENDCLI%(NC%) Then
                   CIXI% = CVI(Mid$(XX$, 3, 2))
                   If GRUPSEL% = 0 Or GRUPSEL% = GRUCOVEI%(CIXI%) Then
                     TOTVENTA#(NC%) = TOTVENTA#(NC%) + CVD(Mid$(XX$, 9, 8))
                     TOTAGEN# = TOTAGEN# + CVD(Mid$(XX$, 9, 8))
                   End If
                 End If
               End If
             Next U&
             '
             For U& = 1 To 32767
               If Abs(TOTVENTA#(U&)) >= 0.005 Then
                 PORPAR = 100 * TOTVENTA#(U&) / TOTAGEN#
                 JJ& = JJ& + 1
                 Z$ = REGBLAN$("RANKCLI")
                 J& = J& + 1: JI% = J&
                 Call REPLA(Z$, MKI$(JI%), 1, 2)
                 NC% = U&
                 Call REPLA(Z$, MKI$(NC%), 3, 2)
                 Call REPLA(Z$, MKD$(TOTVENTA#(U&)), 5, 8)
                 Call REPLA(Z$, MKS$(PORPAR), 13, 4)
                 Call REPLA(Z$, MKS$(PORACU), 17, 4)
                 Call GRAREG("RANKCLI", Z$, J&)
               End If
             Next U&
             '
             Call SETULTREG("RANKCLI", J&)
             Call CIERRARCH("RANKCLI")
             Call HEADERS("RANKCLI", "")
             Call HEADERS("RANKCLI", CORRESP$)
             '
            Call FINAL("*LRANCLI")
        End If
        GoTo 212
        '
13700   TOT# = (Abs(Int(100 * PRE# * CAN# + 0.5)) / 100) * Sgn(CAN#)
        If NC% <> 0 Then
           If NC% < 0 Or NC% > 32767 Then NC% = 32767
           '
           If CI% > 0 Then
               BUSX$ = MKI$(NC%) + MKI$(CI%)
               For K& = 1 To (Len(IREVE$)) / 4  '  To ULTREG&("REVECLI")
                   If Mid$(IREVE$, 4 * K& - 3, 4) = BUSX$ Then
                       RVCL$ = REGLEIDO$("REVECLI", K&)
                       CARVCL = CAN# + CVS(Mid$(RVCL$, 5, 4))
                       VARVCL# = TOT# + CVD(Mid$(RVCL$, 9, 8))
                       Call REPLA(RVCL$, MKS$(CARVCL) + MKD$(VARVCL#), 5, 12)
                       Call GRAREG("REVECLI", RVCL$, K&)
                       '
                       PVCL$ = REGLEIDO$("REVEPIE", K&)
                       Call REPLA(PVCL$, MKS$(CARVCL), 5, 4)
                       Call GRAREG("REVEPIE", PVCL$, K&)
                       '
                       GoTo 13799
                   End If
               Next K&
               '
               CARVCL = CAN#
               RVCL$ = MKI$(NC%) + MKI$(CI%) + MKS$(CARVCL) + MKD$(TOT#) + MKD$(0)
               Call REGAPP("REVECLI", RVCL$)
               '
               PVCL$ = MKI$(CI%) + MKI$(NC%) + MKS$(CARVCL) + MKD$(0) + String$(4, 0)
               Call REGAPP("REVEPIE", PVCL$)
               '
               IREVE$ = IREVE$ + MKI$(NC%) + MKI$(CI%)
               '
           End If
        End If
        '
13799   Return
        '
End Sub
'

Sub DETARTIC()       ' detalle de entregas por articulo
    '
    Dim PUNMO$(32767)
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
        Exit Sub
    End If
    '
    Call BLOQARCH("DETARTIC")
    '
    DES% = CVI(Left$(OBX$, 2))
    HAS% = CVI(Mid$(OBX$, 3, 2))
    If DES% > HAS% Then
        GoTo 10
    End If
    FEX = DES%: PERIO$ = FECHATEX$(FEX)
    FEX = HAS%: PERIO$ = PERIO$ + " - " + FECHATEX$(FEX)
    '
    J& = 0
    '
    ARMOVI$ = "MOVISTO"
    '
113 Screen.MousePointer = 11
    FIN& = ULTREG&(ARMOVI$)
    '
    For U& = 1 To FIN&
       Call TRACE(20, U&, FIN&)
       X$ = REGLEIDO$(ARMOVI$, U&)
       FF% = CVI(Left$(X$, 2))
       If FF% >= DES% Then
          If FF% <= HAS% Then
             CI% = CVI(Mid$(X$, 3, 2))
             If CI% > 0 Then
                If CI% <= 32767 Then
                   TIMO1$ = UCase$(Mid$(X$, 23, 2))
                   TIMO2$ = UCase$(Mid$(X$, 33, 2))
                   If TIMO1$ = "RT" Or TIM2$ = "RT" Then
                      PUNMO$(CI%) = PUNMO$(CI%) + MKS$(U&)
                   End If
                End If
             End If
          End If
       End If
    Next U&
    '
    UREMO& = ULTREG&(ARMOVI$)
    FIN& = NUINV%
    For U& = 1 To FIN&
       Call TRACE(20, U&, FIN&)
       Y$ = REGLEIDO$("INVERT", U& + 1)
       CI% = CVI(Mid$(Y$, 17, 2))
       If CI% > 0 Then
          If CI% <= NUMAS% Then
             For KKI% = 1 To Len(PUNMO$(CI%)) Step 4
                REG& = CVS(Mid$(PUNMO$(CI%), KKI%, 4))
                If REG& > 0 Then
                   If REG& <= UREMO& Then
                      X$ = REGLEIDO$(ARMOVI$, REG&)
                      CLI% = CVI(Mid$(X$, 84, 2))
                      FF% = CVI(Left$(X$, 2))
                      CII% = CVI(Mid$(X$, 3, 2))
                      REMIFAC$ = Mid$(X$, 23, 10)
                      CANTU# = CVD(Mid$(X$, 96, 8)) - CVD(Mid$(X$, 88, 8))
                      '
                      Z$ = REGBLAN$("DETARTIC")
                      Call REPLA(Z$, MKI$(CII%), 1, 2)
                      Call REPLA(Z$, MKI$(FF%), 3, 2)
                      Call REPLA(Z$, REMIFAC$, 5, 10)
                      Call REPLA(Z$, MKI$(CLI%), 15, 2)
                      Call REPLA(Z$, MKD$(CANTU#), 17, 8)
                      J& = J& + 1
                      Call GRAREG("DETARTIC", Z$, J&)
                   End If
                End If
             Next KKI%
          End If
       End If
    Next U&
    '
    If ARMOVI$ = "MOVISTO" Then
      If EXISTE%(LUDAT$ + "MOVISVIE.DAT") > 0 Then
        XXX$ = REGLEIDO$("MOVISTO", 1)
        FF% = CVI(Left$(XXX$, 2))
        If FF% >= DES% Then
          ARMOVI$ = "MOVISVIE"
          GoTo 113
        End If
      End If
    End If
    '
    Call SETULTREG("DETARTIC", J&)
    Call CIERRARCH("DETARTIC")
    Call LIBARCH("DETARTIC")
    Call FINAL("*DETARTI")
    '
    Call CIERRARCH("*.*")
    Screen.MousePointer = 1
    '
End Sub
