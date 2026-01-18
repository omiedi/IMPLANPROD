Attribute VB_Name = "FAC_FUNCTIONS"
Dim DATCLI$(32)
'
Dim UCOBRA%, UNUREC&
Public TICADOL

Sub CAREMPRE(LU$, A$())
      '
      LDT$ = LUDAT$
      N% = FILEOPEN%(LDT$ + "SVD147.EXE", 0, 41)
      Dim V1 As String * 41
      If LOF(N%) < 192 Then
          Close #N%: N% = 0
          LDT1$ = REPLACAR$(LDT$, "\", "/")
          Call MENSERR(24, "Falta Archivo de Datos de la Empresa\en Directorio '" + LDT1$ + "'.")
          Exit Sub
      End If
      '
      For KK% = 1 To 8
          Get #N%, KK%, V1$
          A$(KK%) = TRIM$(DESENCRI$(V1$))
      Next KK%
      Close #N%: N% = 0
      '
End Sub
'

Sub GRAFAC98(LU$, ARCHIFAC$, OPCI%, RRR%, FLIM%, TIPO%, VA%, CAFOR%, NRO, NC%, CLI$, iva$, FF%, FV%, LPRE%, VENDE%, TOTAFAC#())
      '
      LU$ = LUDAT$
      Static UREG&
      Static NROMAX(3)
      Static CTX%
      Static TIPOVANT%
      Static SUCUVENA%
      Static SUCLOCAL%
      '
      SUCUVEN% = 0
      On Error Resume Next
      SUCUVEN% = Val(FORFACNEU.Label12.Caption)
      SUCUVEN% = Val(FORFACLI.Label12.Caption)
      SUCUVEN% = Val(FACTUR00.Label12.Caption)
      On Error GoTo 0
        '
        If OPCI% < 1 Or OPCI% > 4 Then
             Call MENSERR(24, "Invocacion de rutina GRAFAC (opcion) - Consulte Estudio C.I.M.E.")
             GoTo 40000
        End If
        '
        If VA% < 0 Or VA% > 3 Then
             Call MENSERR(24, "Invocacion rutina GRAFAC (variante) - Consulte Estudio C.I.M.E.")
             GoTo 40000
        End If

        TIPOVA% = 100 * TIPO% + VA%
        COMPRASI% = 0

        If CTX% = 0 Then
           CTX% = 1
           SUCLOCAL% = 0
           If ULTREG&("SUCREMOT") > 0 Then
              X$ = REGLEIDO$("SUCREMOT", 1&)
              SUCLOCAL = CVI(Left$(X$, 2))
           End If
           Call CIERRARCH("SUCREMOT")
        End If


110   ' determina NUFAC&

        NUFAC& = ULTREG&("FACTUR")
111     If NUFAC& < 1 Then NUFAC& = 0: GoTo 120
        F$ = REGLEIDO$("FACTUR", NUFAC&): GoSub 10000

120   ' determina NROMAX(VA%)

        If NROMAX(VA%) > 0 Then
           If TIPOVA% = TIPOVANT% Then
              If SUCUVEN% = SUCUVENA% Then
                 GoTo 200
              End If
           End If
        End If
        TIPOVANT% = TIPOVA%
        SUCUVENA% = SUCUVEN%
        NROMAX(VA%) = 0

        For MXC& = ULTREG&("FACTUR") To 1 Step -1
                F$ = REGLEIDO$("FACTUR", MXC&): GoSub 10000
                If CVI(F3$) < FLIM% Then GoTo 180
                If CVI(F11$) <> SUCUVEN% Then GoTo 180
                TIX% = Asc(F6$): VAX% = Asc(Mid$(F6$, 2))
                              If VAX% < 0 Or VAX% > 3 Then VAX% = 1
                              If VAX% <> VA% Then GoTo 180
                If TIPO% > 6 Then GoTo 150

                If TIX% < 1 Or TIX% = 4 Or TIX% = 5 Or TIX% > 6 Then GoTo 180 ' descarta otros documentos
                If CAFOR% > 1 Then GoTo 150 ' numeracion SEPARADA para facturas,debitos y creditos
                If CVS(F2$) > NROMAX(VAX%) Then NROMAX(VAX%) = CVS(F2$)
                GoTo 180

150             'If (TIPO% = 1 Or TIPO% = 2) Then If (TIX% = 1 Or TIX% = 2) Then TIX% = TIPO%
                If TIX% <> TIPO% Then GoTo 180
                If CVS(F2$) > NROMAX(VAX%) Then NROMAX(VAX%) = CVS(F2$)

180             If NROMAX(VA%) > 0 Then
                    If MXC& < ULTREG&("FACTUR") - 100 Then
                        GoTo 200
                    End If
                End If

        Next MXC&

200     If COMPRASI% = 1 Then CAFOR% = 2
        If OPCI% <> 1 Then GoTo 300
        NRO = NROMAX(VA%) + 1
        GoTo 40000

300     If OPCI% <> 2 Then GoTo 400
        If COMPRASI% = 0 Then If NRO >= NROMAX(VA%) + 1 Then GoTo 40000 ' acepta valor

        For MXC& = ULTREG&("FACTUR") To 1 Step -1
                F$ = REGLEIDO$("FACTUR", MXC&): GoSub 10000
                If CVI(F3$) < FLIM% Then GoTo 380
                If CVI(F11$) <> SUCUVEN% Then GoTo 380
                TIX% = Asc(F6$): VAX% = Asc(Mid$(F6$, 2)): If VAX% < 0 Or VAX% > 3 Then VAX% = 1
                If VAX% <> VA% Then GoTo 380
                If TIPO% > 6 Then GoTo 350

                If TIX% < 1 Or TIX% = 4 Or TIX% = 5 Or TIX% > 6 Then GoTo 380 ' descarta otros documentos

                If CAFOR% > 1 Then GoTo 350 ' numeracion SEPARADA para facturas,debitos y creditos
                If CVS(F2$) = NRO Then NRO = NROMAX(VA%) + 1: GoTo 40000
                GoTo 380

350             'If (TIPO% = 1 Or TIPO% = 2) Then If (TIX% = 1 Or TIX% = 2) Then TIX% = TIPO%
                If TIX% <> TIPO% Then GoTo 380
                If COMPRASI% = 1 Then If CVI(F1$) <> NC% Then GoTo 380
                If CVS(F2$) = NRO Then NRO = NROMAX(VA%) + 1: GoTo 40000

380     Next MXC&
390     GoTo 40000

' vuelve a determinar NUFAC& y NROMAX(1,2,3)

400     If OPCI% <> 3 Then GoTo 800     ' opci%=3 - reserva registro y numero

        If COMPRASI% = 1 Then GoTo 600

        If UREG& = 0 Then
            UREG& = ULTREG&("FACTUR") - 128
            If UREG& < 0 Then UREG& = 0
        End If

        For MXC& = UREG& + 1 To ULTREG&("FACTUR")
                F$ = REGLEIDO$("FACTUR", MXC&): GoSub 10000
                If CVI(F1$) = 0 Then GoTo 490
                If CVI(F11$) <> SUCUVEN% Then GoTo 480
                NUFAC& = MXC&: If CVI(F3$) < FLIM% Then GoTo 480
                TIX% = Asc(F6$): VAX% = Asc(Mid$(F6$, 2)): If VAX% < 0 Or VAX% > 3 Then VAX% = 1
                If TIPO% > 6 Then GoTo 450

                If TIX% < 1 Or TIX% = 4 Or TIX% = 5 Or TIX% > 6 Then GoTo 480 ' descarta otros documentos

                If CAFOR% > 1 Then GoTo 450 ' numeracion SEPARADA para facturas,debitos y creditos
                If CVS(F2$) > NROMAX(VAX%) Then NROMAX(VAX%) = CVS(F2$)
                GoTo 480

450             'If (TIPO% = 1 Or TIPO% = 2) Then If (TIX% = 1 Or TIX% = 2) Then TIX% = TIPO%
                If TIX% <> TIPO% Then GoTo 480
                If CVS(F2$) > NROMAX(VAX%) Then NROMAX(VAX%) = CVS(F2$)

480     Next MXC&

490     ' CONTINUE

500     If NRO >= NROMAX(VA%) + 1 Then GoTo 600 ' acepta valor

        For MXC& = NUFAC& To 1 Step -1
                F$ = REGLEIDO$("FACTUR", MXC&): GoSub 10000
                If CVI(F3$) < FLIM% Then GoTo 580
                If CVI(F11$) <> SUCUVEN% Then GoTo 580
                TIX% = Asc(F6$): VAX% = Asc(Mid$(F6$, 2)): If VAX% < 0 Or VAX% > 3 Then VAX% = 1
                If VAX% <> VA% Then GoTo 580
                If TIPO% > 6 Then GoTo 550

                If TIX% < 1 Or TIX% = 4 Or TIX% = 5 Or TIX% > 6 Then GoTo 580 ' descarta otros documentos

                If CAFOR% > 1 Then GoTo 550 ' numeracion SEPARADA para facturas,debitos y creditos
                If CVS(F2$) = NRO Then Call MENSERR(24, "Documento" + Str$(NRO) + " ya Utilizado - Use Numero" + Str$(NROMAX(VA%) + 1)): NRO = NROMAX(VA%) + 1: GoTo 600
                GoTo 580

550             'If (TIPO% = 1 Or TIPO% = 2) Then If (TIX% = 1 Or TIX% = 2) Then TIX% = TIPO%
                If TIX% <> TIPO% Then GoTo 580
                If CVS(F2$) = NRO Then Call MENSERR(24, "Documento" + Str$(NRO) + " ya Utilizado - Use Numero" + Str$(NROMAX(VA%) + 1)): NRO = NROMAX(VA%) + 1: GoTo 600

580     Next MXC&

590     ' CONTINUE

600     CLI$ = Left$(CLI$ + Space$(30), 30)
        If NC% = 0 And CLI$ <> Space$(30) Then GoSub 1000
        If NC% = 0 Then NC% = 10001

610     F$ = ""
        F$ = F$ + MKI$(NC%)
        F$ = F$ + MKS$(NRO)
        F$ = F$ + MKI$(FF%)
        F$ = F$ + MKI$(FV%)
        F$ = F$ + Chr$(LPRE%)
        F$ = F$ + Chr$(MONEMI%)
        F$ = F$ + Chr$(TIPO%) + Chr$(VA%)
        F$ = F$ + MKD$(0)
        F$ = F$ + MKD$(0)
        F$ = F$ + MKD$(0)
        F$ = F$ + MKD$(0)
        F$ = F$ + MKI$(SUCUVEN%)
        F$ = F$ + MKD$(0)
        F$ = F$ + MKD$(0)
        '
        Call REGAPP("FACTUR", F$)
        UREG& = ULTREG&("FACTUR")
        Call CIERRARCH("FACTUR")

        If NRO > NROMAX(VA%) Then NROMAX(VA%) = NRO
        GoTo 40000

800     ' OPCI%=4 - grabacion definitiva
        NEFA# = TOTAFAC#(1, 5)
        NIVA# = TOTAFAC#(1, 6)
        NNIN# = TOTAFAC#(1, 7)
        TOTA# = TOTAFAC#(1, 8)
        IINT# = TOTAFAC#(1, 9)
        PERC# = TOTAFAC#(1, 10)
        '
        F$ = REGLEIDO$("FACTUR", UREG&)
        Call REPLA(F$, MKD$(NEFA#), 15, 8)
        Call REPLA(F$, MKD$(NIVA#), 23, 8)
        Call REPLA(F$, MKD$(NNIN#), 31, 8)
        Call REPLA(F$, MKD$(TOTA#), 39, 8)
        Call REPLA(F$, MKI$(SUCUVEN%), 47, 2)
        Call REPLA(F$, MKD$(IINT#), 49, 8)
        Call REPLA(F$, MKD$(PERC#), 57, 8)
        '
        Call GRAREG("FACTUR", F$, UREG&)
        Call CIERRARCH("FACTUR")
        '
        GoTo 40000

1000    CU$ = REGBLAN$("CUITS")
        CU$ = REPLACE$(CU$, CLI$, 1, 30)
        CU$ = REPLACE$(CU$, iva$, 31, 15)
        Call REGAPP("CUITS", CU$)
        NC% = (-1) * ULTREG&("CUITS")
        Call CIERRARCH("CUITS")
        '
        Return

10000   F1$ = Left$(F$, 2)
        F2$ = Mid$(F$, 3, 4)
        F3$ = Mid$(F$, 7, 2)
        F4$ = Mid$(F$, 9, 2)
        F5$ = Mid$(F$, 11, 2)
        F6$ = Mid$(F$, 13, 2)
        F7$ = Mid$(F$, 15, 8)
        F8$ = Mid$(F$, 23, 8)
        F9$ = Mid$(F$, 31, 8)
        F10$ = Mid$(F$, 39, 8)
        F11$ = Mid$(F$, 47, 2)
        F12$ = Mid$(F$, 49, 8)
        F13$ = Mid$(F$, 57, 8)
        If CVI(F11$) < 1 Then F11$ = MKI$(SUCLOCAL%)
        Return


40000  End Sub
'

Sub GRAM2000(NC%, NRO, REMI, FF%)
        '
        SUCUVEN% = 0
        On Error Resume Next
        SUCUVEN% = Val(FORFACNEU.Label12.Caption)
        SUCUVEN% = Val(FACTUR00.Label12.Caption)
        On Error GoTo 0
        '
        DEPOSALI% = 1
        If SUCUVEN% > 1 Then
           For IU& = 1 To ULTREG&("SUCREMOT")
              X$ = REGLEIDO$("SUCREMOT", IU&)
              If CVI(Left$(X$, 2)) = SUCUVEN% Then
                 DEPOSALI% = Asc(Mid$(X$, 33, 1))
                 If ECOARCH$("TADEPOSI", Chr$(DEPOSALI%)) = "" Then DEPOSALI% = 1
              End If
           Next IU&
        End If
        Call CIERRARCH("SUCREMOT")
        '
        NC1% = NC%: If NC% < 1 Then NC1% = 10001
        REFE$ = ECOARCH$("DEUDOR1", MKI$(NC%))
        FIN& = ULTREG&("!CARDEFAC")
        '
        For I& = 1 To FIN&
             Call TRACE(20, I&, FIN&)
             XQ$ = REGLEIDO$("!CARDEFAC", I&)
             CANTAA = CVS(Left$(XQ$, 4))
             If CANTAA > 0 Then
               CI% = CVI(Mid$(XQ$, 5, 2))
               'If CI% > 0 Then 'DESHABILITADO PARA RESERVAR NUMERO
                                'EN FACTURACION DE TEXTO LIBRE
                 COD$ = CODEXT$(CI%)
                 TICAX = 1
                 If MONEMI% = 2 Then
                    TICAX = VH(MONEMI%)
                    If TICADOL > 0.005 Then TICAX = TICADOL
                 End If
                 PRE# = CDbl(CVS(Mid$(XQ$, 429, 4))) * TICAX
                 CMOV$ = "RT"
                 If Mid$(XQ$, 461, 1) = "R" Then CMOV$ = "*R" ' REPARACION
                 DOPRI$ = "RT." + TRIM$(Str$(REMI))
                 DOSEC$ = "FC." + TRIM$(Str$(NRO))
                 '
                 REPED& = CVS(Mid$(XQ$, 509, 4))
                 If REPED& < 0 Then
                     REMIRX = Val(Mid$(XQ$, 387, 16))
                     If REMIRX > 0.5 Then
                        DOPRI$ = "RT." + TRIM$(Str$(REMIRX))
                        ENCURE% = 0
                        GoSub 2000
                        If ENCURE% > 0 Then GoTo 1995
                     End If
                 End If
                 '
                 X$ = ""
                 SIG% = (-1)
                 '
                 If NRO < 0 Then
                    CMOV$ = "RD"
                    DOPRI$ = "DV." + TRIM$(Str$(Abs(NRO)))
                    DOSEC$ = "NC." + TRIM$(Str$(Abs(NRO)))
                    SIG% = 1
                 End If
                 '
                 If REMI < 0 Then
                    CMOV$ = "ND"
                    DOPRI$ = "ND." + TRIM$(Str$(Abs(REMI)))
                    DOSEC$ = "ND." + TRIM$(Str$(Abs(NRO)))
                    SIG% = (-1)
                 End If
                 '
                 If InStr(EMPREAC$, "NEUMANN") > 0 Then
                   If REMI = 0 Then
                     DOPRI$ = DOSEC$
                   End If
                 End If
                 '
                 ' correccion falla sabo
                 If InStr(EMPREAC$, "SABO") > 0 Then
                    If Val(Mid$(DOPRI$, 4)) < 1 Then
                       If Left$(DOSEC$, 3) = "FC." Then
                          DOPRI$ = "RT." + TRIM$(Str$(NRO))
                       End If
                    End If
                 End If
                 '
                 Call MOVISTO(FF%, CI%, LOTE$, CMOV$, DOPRI$, DOSEC$, REFE$, PRE#, "$", NC%, DEPOSALI%, SIG% * CANTAA)
         EPAC$ = TRIM$(UCase$(EMPREAC$))
         If InStr(EPAC$, "DOSIVA") > 0 Then
           If Left$(DOPRI$, 2) = "RT" Then
             DOPRIQ$ = "PL." + TRIM$(Mid$(XQ$, 371, 16))
             If TRIM$(DOPRIQ$) = "PL." Then   ' REMITO DIRECTO SIN O/VENTA
               NRXLX& = 900000 + Val(Text10)
               DOPRIQ$ = DOPRIQ$ + TRIM$(Str$(NRXLX&))
             End If
             REFEQ$ = "Packing-List " + REFE$
             If CMOV$ = "*R" Then
                  CMOVQ$ = "*D"
               Else
                  CMOVQ$ = "RP"
             End If
             Call MOVISTO(FF%, CI%, LOTE$, CMOVQ$, DOPRIQ$, DOSEC$, REFEQ$, 0, "$", 0, DEPOSALI%, (-1) * SIG% * CANTAA)
           End If
         End If
         '
                 '
1995             If CI% > 0 Then
                    If CI% <= NUMAS% Then
                        RXX$ = RECFILT$("DESPAIMP", 1, MKI$(CI%))
                        CANTU = CANTAA * SIG% * (-1)
                        '
                        For VV% = 1 To Len(RXX$) Step 4
                            V& = CVS(Mid$(RXX$, VV%, 4))
                            If V& > 0 Then
                              Y$ = REGLEIDO$("DESPAIMP", V&)
                              If CVI(Left$(Y$, 2)) = CI% Then
                                CAPEN# = CVD(Mid$(Y$, 41, 8))
                                If CAPEN# >= CANTU Then
                                    CAPEN# = CAPEN# - CDbl(CANTU)
                                    Call REPLA(Y$, MKD$(CAPEN#), 41, 8)
                                    Call GRAREG("DESPAIMP", Y$, V&)
                                    CANTU = 0
                                    GoTo 1998
                                  Else
                                    CANTU = CANTU - CAPEN#
                                    CAPEN# = 0
                                    Call REPLA(Y$, MKD$(0), 41, 8)
                                    Call GRAREG("DESPAIMP", Y$, V&)
                                End If
                              End If
                            End If
                        Next VV%
                    End If
                 End If
               'End If
             End If
             '
1998    Next I&
        Call CIERRARCH("MOVISTO")
        GoTo 3000
        '
2000    ' graba por remitos pendientes
        '
        For ILKK& = ULTREG("MOVISTO") To 1 Step -1
            X$ = REGLEIDO$("MOVISTO", ILKK&)
            If CVI(Mid$(X$, 84, 2)) = NC% Then
                If TRIM$(Mid$(X$, 23, 10)) = TRIM$(Left$(DOPRI$, 10)) Then
                    If CVI(Mid$(X$, 3, 2)) = CI% Then
                        If Abs(CVD(Mid$(X$, 96, 8)) - CANTAA) < 0.05 Then
                            ENCURE% = 1
                            Call REPLA(X$, DOSEC$, 33, 10)
                            Call REPLA(X$, MKD$(PRE#), 73, 8)
                            Call GRAREG("MOVISTO", X$, ILKK&)
                            GoTo 2090
                        End If
                    End If
                End If
            End If
        Next ILKK&
        '
2090    Call CIERRARCH("MOVISTO")
        Return
        '
3000    End Sub
'
Sub GRADECONSI(RECONSI$)
    
    For I& = 1 To ULTREG&("!CARDEFAC")
       X$ = REGLEIDO$("!CARDEFAC", I&)
       CANTB# = CVS(Left$(X$, 4))
       If CANTB# > 0 Then
          CI% = CVI(Mid$(X$, 5, 2))
          If CI% > 0 Then
            If CI% <= NUMAS% Then
              If CANTB# > 0.005 Then
                 ZZ$ = RECONSI$
                 Call REPLA(ZZ$, MKI$(CI%), 35, 2)
                 Call REPLA(ZZ$, MKD$(CANTB#), 45, 8)
                 Call REPLA(ZZ$, MEDIDAS$, 109, 12)
                 Call REGAPP("CONSICLI", ZZ$)
              End If
            End If
          End If
       End If
    Next I&
    '
    Call CIERRARCH("CONSICLI")
    '
End Sub
'
Sub GRACLI(LU$, Archi$, NC%, TOTFAC#, DATCLI$(), OPCI%)
        '
        If OPCI% <> 2 Then
            Call MENSERR(24, "Error de Invocacion de Rutina 'GRACLI'\Consulte a su Soporte de Sistemas.")
            Call FINAL("")
        End If
        '
        ARCHXX$ = "DEUDOR": If NC% < 0 Then ARCHXX$ = "PROVED"
        RGXX& = REGICLI&(NC%)
        '
        If RGXX& > 0 Then
            If RGXX& <= ULTREG&(ARCHXX$ + "1") Then
                TTXS$ = REGLEIDO$(ARCHXX$ + "2", RGXX&)
                SALNUE# = (Int(100 * TOTFAC# + 0.5)) / 100 + (Int(100 * CVD(Mid$(TTXS$, 109, 8)) + 0.5)) / 100
                Call REPLA(TTXS$, MKD$(SALNUE#), 109, 8)
                Call GRAREG(ARCHXX$ + "2", TTXS$, RGXX&)
            End If
        End If
        '
        Call CIERRARCH(ARCHXX$ + "2")
        '
End Sub
'


Sub DESCUEN(PORDES, DESC$)
    '
    Dim DESC1(10)
    PORDES = 0: KKK% = 0
    DESCX$ = REPLACAR$(DESC$, "%", " ") + "+"
    '
10  XDC% = InStr(DESCX$, "+")
    KKK% = KKK% + 1
    DESC1(KKK%) = Val(Left$(DESCX$, XDC% - 1))
    DESCX$ = TRIM$(Mid$(DESCX$, XDC% + 1))
    If DESCX$ <> "" Then GoTo 10
    '
    COEFIDE = 1
    For KKJ% = 1 To KKK%
       COEFIDE = COEFIDE * (1 - DESC1(KKJ%) / 100)
    Next KKJ%
    PORDES = 100 * (1 - COEFIDE)
    '
End Sub
'

Function LETREXPO$()
    '
    Static LTX$
    If LTX$ = "" Then
      LTX$ = "A"
      If ULTREG&("FORMFAC") >= 2 Then
        FFS$ = REGLEIDO$("FORMFAC", 2&)
        LTX1$ = Mid$(FFS$, 94, 1)
        If Asc(LTX1$) >= 65 Then
            LTX$ = UCase$(LTX1$)
        End If
      End If
    End If
    '
    LETREXPO$ = LTX$
    '
End Function
'
Function PREFEXPO$()
    '
    Static PFX$
    If PFX$ = "" Then
      PFX$ = "0001"
      If ULTREG&("FORMFAC") >= 2 Then
        FFS$ = REGLEIDO$("FORMFAC", 2&)
        PFX1$ = Mid$(FFS$, 82, 4)
        If TRIM$(PFX1$) = PFX1$ Then
            PFX$ = PFX1$
        End If
      End If
    End If
    '
    PREFEXPO$ = PFX$
    '
End Function
'
Sub APLIPAG(FE%, NC%, TI%, VA%, PRF%, NROF&, IMAPLI#, REFE$, DOCORI$, ITOTORI#, IVAORI#)
    '
    ARAPLI$ = "APLICOB": If NC% < 0 Then ARAPLI$ = "APLIPAG"
    PRFS$ = TRIM$(Str$(PRF%)): While Len(PRFS$) < 4: PRFS$ = "0" + PRFS$: Wend
    NROS$ = TRIM$(Str$(NROF&)): While Len(NROS$) < 8: NROS$ = "0" + NROS$: Wend
    '
    TIX% = TI% Mod 10
    TINUVA$ = "FC"                                               ' por default
    If TIX% = 0 Then TINUVA$ = "AS"
    If TIX% = 1 Then TINUVA$ = "VC"
    If TIX% = 3 Then TINUVA$ = "ND"
        If TIX% = 4 Then
        TINUVA$ = "RB"
        If NC% < 0 Then
            TINUVA$ = "OP"
            If TI% = 14 Then TINUVA$ = "FF"
        End If
    End If
    '
    If TIX% = 5 Then TINUVA$ = "DC"
    If TIX% = 6 Then TINUVA$ = "NC"
    If TIX% = 7 Then TINUVA$ = "DP"
    '
    VAD$ = Chr$(64 + VA%): If VAD$ < "A" Or VAD$ > "Z" Then VAD$ = "X"
    TINUVA$ = TINUVA$ + "-" + VAD$ + "-" + PRFS$ + "-" + NROS$
    '
    RECORD$ = MKI$(FE%)
    RECORD$ = RECORD$ + MKI$(Abs(NC%))
    '
    RECORD$ = RECORD$ + TINUVA$
    RECORD$ = RECORD$ + MKD$(IMAPLI#)
    '
    REFE$ = Left$(REFE$ + Space$(24), 24)
    RECORD$ = RECORD$ + REFE$
    '
    RECORD$ = RECORD$ + DOCORI$ + MKD$(ITOTORI#) + MKD$(IVAORI#)
    '
    RECORD$ = RECORD$ + String$(8, 0)
    '
    Call REGAPP(ARAPLI$, RECORD$)
    'Call CIERRARCH(ARAPLI$)
    '
End Sub
'

Sub PRIRECISTA()
    '
' En la impresion del recibo se introducen las siguientes opciones:
' 1) El CUIT del librador puede imprimirse en el mismo renglon que
'    el cheque recibido, para lo cual el formato A16 va con el
'    parámetro de impresion CUIT-TRA. Verificar en este caso que la
'    longitud de registro de PRINRECI.RFS sea al menos 96
' 2) El CUIT del librador va en el renglon siguiente, en la misma
'    columna que el nombre del banco, con lo que el formato CUIT-TRA
'    debe ir en blanco.
'
    Dim ORCOP$(6)
    ORCOP$(1) = "ORIGINAL"
    ORCOP$(2) = "DUPLICADO"
    ORCOP$(3) = "TRIPLICADO"
    ORCOP$(4) = "CUADRUPLICADO"
    ORCOP$(5) = "QUINTUPLICADO"
    ORCOP$(6) = "SEXTUPLICADO"
    '
    FORIPRE$ = TRIM$(CONTROL$("", "FORECIST"))
    'If FORIPRE$ = "" Then
    '    Call COMUNI("Error de Configuración:\  \Falta Definir Formulario\de Impresión de Recibo.")
    '    Exit Sub
    'End If
    '
    If FORIPRE$ <> "" Then
        LU$ = LUDAT$
        If EXISTE%(LU$ + FORIPRE$ + ".FRM") <> 1 Then
            RMCC$ = REPLACAR$(LU$ + FORIPRE$ + ".FRM", "\", "/")
            Call COMUNI("Error de Configuración:\  \No Existe Formulario de Impresión\de Recibo de Cobranza:\  \'" + RMCC$ + "'.")
            Exit Sub
        End If
    End If
    '
    CONTA% = ULTREG&("!CARMECO")
    If CONTA% < 1 Then
        Call COMUNI("No se ha Ingresado Detalle de Valores")
        Exit Sub
    End If
    '
    AVAHOJA$ = TRIM$(CONTROL$("RECISTA", "AVAHOJA"))
    CANCELPRINT% = 0
    If CONTROL$("RECISTA", "SUPRINT") = "SI" Then CANCELPRINT% = 1
    '
    Screen.MousePointer = 11
    Call BLANARCH("!PRINRECI")
    For U& = 1 To ULTREG&("!CARMECO")
        X$ = REGLEIDO$("!CARMECO", U&)
        If X$ = REBLA$ Then GoTo 19
        MECO% = CVI(Left$(X$, 2))
        BANCO$ = TRIM$(ECOARCH$("MECOBRA", MKI$(MECO%)))
        NUVA$ = Mid$(X$, 19, 12)
        IMVA# = CVD(Mid$(X$, 31, 8))
        FEMII% = CVI(Mid$(X$, 39, 2))
        FEVEI% = CVI(Mid$(X$, 41, 2))
        PROTER$ = UCase$(TRIM$(Mid$(X$, 43, 2)))
        CLEA% = CVI(Mid$(X$, 45, 2))
        CUITLIB$ = TRIM$(Mid$(X$, 47, 16))
        '
        TIPOSS$ = Space$(4)
        If MECO% > 0 Then
            For V& = 1 To ULTREG&("MECOBRA")
                MCB$ = REGLEIDO$("MECOBRA", V&)
                If CVI(Left$(MCB$, 2)) = MECO% Then
                    TIPOSS$ = Mid$(MCB$, 57, 4)
                    GoTo 10
                End If
            Next V&
        End If
        '
10      Z$ = REGBLAN$("!PRINRECI")
        Call REPLA(Z$, TIPOSS$, 1, 4)
        Call REPLA(Z$, NUVA$, 5, 12)
        Call REPLA(Z$, BANCO$, 17, 16)
        Call REPLA(Z$, MKD$(IMVA#), 33, 8)
        Call REPLA(Z$, MKI$(0), 51, 2)
        FEX = FEVEI%
        If FEX > 0 Then
            Call REPLA(Z$, FECHATEX$(FEX), 41, 8)
        End If
        Call REGAPP("!PRINRECI", Z$)
        '
        If CUITLIB$ <> "" Then
          FORPRICU$ = TRIM$(UCase$(ATRIFORM$(FORIPRE$, "CUIT-TRA/FORMATO")))
          If FORPRICU$ <> "" Then
              URE& = ULTREG&("!PRINRECI")
              Z$ = REGLEIDO$("!PRINRECI", URE&)
              Call REPLA(Z$, CUITLIB$, 79, 16)
              Call GRAREG("!PRINRECI", Z$, URE&)
            Else
              Z$ = REGBLAN$("!PRINRECI")
              Call REPLA(Z$, ".CUIT Librad", 5, 12)
              Call REPLA(Z$, CUITLIB$, 17, 16)
              Call REGAPP("!PRINRECI", Z$)
          End If
        End If
        '
19  Next U&
    '
    If Val(FORECISTA.DESCUEN(0)) >= 0.005 Then
        Z$ = REGBLAN$("!PRINRECI")
        Call REPLA(Z$, "N/Cr", 1, 4)
        Call REPLA(Z$, "", 5, 12)   ' NUMERO DE N/CREDITO
        Call REPLA(Z$, "Desc.Caja", 17, 16)
        Call REPLA(Z$, MKD$(Val(FORECISTA.DESCUEN(0))), 33, 8)
        Call REPLA(Z$, MKI$(0), 51, 2)
        FEX = HOY(HOS$)
        Call REPLA(Z$, FECHATEX$(FEX), 41, 8)
        Call REGAPP("!PRINRECI", Z$)
    End If
    '
    V0& = ULTREG&("!PRINRECI")
    '
    VACON% = 1
    FORECISTA.List2.Visible = False
    ILISTA& = FORECISTA.List2.ListIndex
    TOLISTA& = FORECISTA.List2.TopIndex
    '
    V& = 0
    For U& = 1 To FORECISTA.List2.ListCount
        FORECISTA.List2.ListIndex = U& - 1
        X$ = FORECISTA.List2.TEXT
        DOCU$ = Mid$(X$, 97, 18)
        CANCE# = CDbl(Val(Mid$(X$, 70, 14)))
        If Abs(CANCE#) >= 0.005 Then
            V& = V& + 1
            If V& <= V0& Then
                  Z$ = REGLEIDO$("!PRINRECI", V&)
               Else
                  Z$ = REGBLAN$("!PRINRECI")
            End If
            FEMIFA% = FECHANUM(Mid$(X$, 23, 8))
            If TRIM$(Mid$(X$, 23, 8)) = "" Then FEMIFA% = HOY(HOS$)
            Call REPLA(Z$, MKI$(FEMIFA%), 51, 2)
            Call REPLA(Z$, DOCU$, 53, 18)
            Call REPLA(Z$, MKD$(CANCE#), 71, 8)
            Call GRAREG("!PRINRECI", Z$, V&)
        End If
    Next U&
    VACON% = 0
    '
    V0& = ULTREG&("!PRINRECI")
    '
    TOTAVAL# = CDbl(Val(FORECISTA.Label14(0).Caption))
    TOTACAN# = CDbl(Val(FORECISTA.Label18(0).Caption)) + CDbl(Val(FORECISTA.Label20(0).Caption))
    ITEXTO$ = NUMTEX$(TOTAVAL#)
    LOTEX% = Val(ATRIFORM$(FORIPRE$, "LOIMPLET")): If LOTEX% < 20 Then LOTEX% = 20
    Call FRATEX(ITEXTO$, NRX1$, NRX2$, NRX3$, LOTEX%)
    IMPLET1$ = TRIM$(NRX1$) + " " + "***************************************************************************************************************"
    IMPLET2$ = TRIM$(NRX2$) + " " + TRIM$(NRX3$) + " ctvs. *****************************************************************************************"
    '
    ANTI# = CDbl(Val(FORECISTA.Label20(0).Caption))
    '
    If Abs(ANTI#) >= 0.005 Then
        '
        V& = V& + 1
        If V& <= V0& Then
             Z$ = REGLEIDO$("!PRINRECI", V&)
           Else
             Z$ = REGBLAN$("!PRINRECI")
        End If
        Call REPLA(Z$, Space$(18), 53, 18)
        Call REPLA(Z$, MKD$(0), 71, 8)
        Call GRAREG("!PRINRECI", Z$, V&)
        '
        V0& = ULTREG&("!PRINRECI")
        '
        V& = V& + 1
        If V& <= V0& Then
             Z$ = REGLEIDO$("!PRINRECI", V&)
           Else
             Z$ = REGBLAN$("!PRINRECI")
        End If
        Call REPLA(Z$, "Anticipo N/Aplic.", 53, 18)
        'If ANTI# < 0 Then
        '    Call REPLA(Z$, "Descuento de Caja", 53, 18)
        'End If
        Call REPLA(Z$, MKD$(ANTI#), 71, 8)
        Call GRAREG("!PRINRECI", Z$, V&)
        '
    End If
    '
    Call CIERRARCH("!PRINRECI")
    '
30  If TRIM$(FORIPRE$) <> "" Then
        '
        CODFOR$ = TRIM$(UCase$(FORIPRE$))
        MARIZW = Val(ATRIFORM$(CODFOR$, "MARIZQMM"))
        MARSUW = Val(ATRIFORM$(CODFOR$, "MARSUPMM"))
        '
        INTERLI = Val(ATRIFORM$(CODFOR$, "INTERLIN"))
        If INTERLI < 1 Then
            Call COMUNI("Error de Configuración de Interlinea\en Formulario '" + CODFOR$ + ".FRM'.")
            INTERLI = 1
        End If
        '
        DESMM = Val(ATRIFORM$(CODFOR$, "DETDESMM"))
        HASMM = Val(ATRIFORM$(CODFOR$, "DETHASMM"))
        '
        CACOPI% = Val(ATRIFORM$(CODFOR$, "CAN-COP"))
        If CACOPI% < 1 Or CACOPI% > 6 Then
            Call COMUNI("Error de Configuración de Cantidad de Copias\en Formulario '" + CODFOR$ + ".FRM'.")
            CACOPI% = 1
        End If
        '
        ORIENTA% = 1
        PORI$ = TRIM$(UCase$(ATRIFORM$(CODFOR$, "ORIENTA")))
        If PORI$ = "PAISAJE" Or PORI$ = "2" Then
            ORIENTA% = 2
        End If
        '
        PORIAN = Printer.Orientation  ' retiene orientacion
        '
        NRORECI$ = TRIM$(Mid$(FORECISTA.Text11.TEXT, 4))
        '
        Screen.MousePointer = 11
        If AVAHOJA$ <> "NEWPAGE" Then
            'Printer.Orientation = 1
            Call SETSPOOL("ORIENTATION", "1")
            If ORIENTA% = 2 Then
                'Printer.Orientation = 2
                Call SETSPOOL("ORIENTATION", "2")
            End If
        End If
        '
        For COP% = 1 To CACOPI%
            '
            If AVAHOJA$ = "NEWPAGE" Then
               If PRINTERCONTROL% = 1 Then
                  'Printer.NewPage
                  Call SETSPOOL("NEWPAGE", "")
               End If
            End If
            PRINTERCONTROL% = 1
            '
            SUMAVA# = 0: SUMACAN# = 0
            Call CARFORWIN
            '
            Call PRINTFORWIN(TRIM$(Str$(NC%)), "CUEN-CLI", 0, 0)
            Call PRINTFORWIN(TRIM$(RASOCLI$(NC%)), "RASO-CLI", 0, 0)
            Call PRINTFORWIN(DOMICLI$(NC%), "DOMI-CLI", 0, 0)
            Call PRINTFORWIN(CPOSCLI$(NC%), "CPOS-CLI", 0, 0)
            Call PRINTFORWIN(LOCACLI$(NC%), "LOCA-CLI", 0, 0)
            Call PRINTFORWIN(CUITCLI$(NC%), "CUIT-CLI", 0, 0)
            Call PRINTFORWIN(POSIVAL$(NC%), "PIVA-CLI", 0, 0)
            FF% = FECHANUM(TRIM$(FORECISTA.Text9.TEXT))
            Call PRINTFORWIN(FETEXTO$(FF%), "FECH-EMI", 0, 0)
            Call PRINTFORWIN(NRORECI$, "NUME-COM", 0, 0)
            Call PRINTFORWIN(ORCOP$(COP%), "ORD-COPI", 0, 0)
            '
            YAC = DESMM
            '
            For YX& = 1 To ULTREG&("!PRINRECI")
               '
               YAC = YAC + INTERLI
               '
               If YAC > HASMM - 2 * INTERLI Then
                   YAC = YAC + INTERLI
                   Call PRINTFORWIN("Transporte", "BANC-CHE", 0, YAC)
                   Call PRINTFORWIN(CSTRING$(MKD$(SUMAVA#), 4, 12, 2, 2), "IMPO-CHE", 0, YAC)
                   Call PRINTFORWIN("Transporte", "DOCU-APL", 0, YAC)
                   Call PRINTFORWIN(CSTRING$(MKD$(SUMACAN#), 4, 12, 2, 2), "IMPO-CAN", 0, YAC)
                   '
                   Call PRINTFORWIN(String$(20, "*"), "TOTA-VAL", 0, 0)
                   Call PRINTFORWIN(String$(20, "*"), "TOTA-CAN", 0, 0)
                   Call PRINTFORWIN(String$(96, "*"), "IMP-LET1", 0, 0)
                   Call PRINTFORWIN(String$(96, "*"), "IMP-LET2", 0, 0)
                   '
                   If AVAHOJA$ = "NEWPAGE" Then
                        'Printer.NewPage
                        Call SETSPOOL("NEWPAGE", "")
                      Else
                        'Printer.EndDoc
                        Call SETSPOOL("ENDDOC", "")
                   End If
                   '
                   Call CARFORWIN
                   Call PRINTFORWIN(TRIM$(Str$(NC%)), "CUEN-CLI", 0, 0)
                   Call PRINTFORWIN(TRIM$(RASOCLI$(NC%)), "RASO-CLI", 0, 0)
                   Call PRINTFORWIN(DOMICLI$(NC%), "DOMI-CLI", 0, 0)
                   Call PRINTFORWIN(CPOSCLI$(NC%), "CPOS-CLI", 0, 0)
                   Call PRINTFORWIN(LOCACLI$(NC%), "LOCA-CLI", 0, 0)
                   Call PRINTFORWIN(CUITCLI$(NC%), "CUIT-CLI", 0, 0)
                   Call PRINTFORWIN(POSIVAL$(NC%), "PIVA-CLI", 0, 0)
                   FEX = FF%
                   Call PRINTFORWIN(FETEXTO$(FF%), "FECH-EMI", 0, 0)
                   Call PRINTFORWIN(NRORECI$, "NUME-COM", 0, 0)
                   '
                   YAC = DESMM + INTERLI
                   Call PRINTFORWIN("Transporte", "BANC-CHE", 0, YAC)
                   Call PRINTFORWIN(CSTRING$(MKD$(SUMAVA#), 4, 12, 2, 2), "IMPO-CHE", 0, YAC)
                   Call PRINTFORWIN("Transporte", "DOCU-APL", 0, YAC)
                   Call PRINTFORWIN(CSTRING$(MKD$(SUMACAN#), 4, 12, 2, 2), "IMPO-CAN", 0, YAC)
                   YAC = YAC + 2 * INTERLI
                   '
                End If
                '
                YY$ = REGLEIDO$("!PRINRECI", YX&)
                Call PRINTFORWIN(Left$(YY$, 4), "TIPO-VAL", 0, YAC)
                Call PRINTFORWIN(Mid$(YY$, 5, 12), "NUME-CHE", 0, YAC)
                Call PRINTFORWIN(Mid$(YY$, 17, 16), "BANC-CHE", 0, YAC)
                Call PRINTFORWIN(CSTRING$(Mid$(YY$, 33, 8), 4, 12, 2, 2), "IMPO-CHE", 0, YAC)
                SUMAVA# = SUMAVA# + CVD(Mid$(YY$, 33, 8))
                Call PRINTFORWIN(CSTRING$(Mid$(YY$, 41, 8), 6, 8, 0, 1), "FECH-CHE", 0, YAC)
                  FEXUU = CVI(Mid$(YY$, 51, 2))
                If FEXUU > 9000 Then Call PRINTFORWIN(FECHATEX$(FEXUU), "FECH-LAN", 0, YAC)
                Call PRINTFORWIN(Mid$(YY$, 53, 18), "DOCU-APL", 0, YAC)
                Call PRINTFORWIN(CSTRING$(Mid$(YY$, 71, 8), 4, 12, 2, 2), "IMPO-CAN", 0, YAC)
                  SUMACAN# = SUMACAN# + CVD(Mid$(YY$, 71, 8))
                Call PRINTFORWIN(TRIM$(Mid$(YY$, 79, 16)), "CUIT-TRA", 0, YAC)
                '
            Next YX&
            '
            Call PRINTFORWIN(CSTRING$(MKD$(TOTAVAL#), 4, 12, 2, 2), "TOTA-VAL", 0, 0)
            Call PRINTFORWIN(CSTRING$(MKD$(TOTACAN#), 4, 12, 2, 2), "TOTA-CAN", 0, 0)
            Call PRINTFORWIN(IMPLET1$, "IMP-LET1", 0, 0)
            Call PRINTFORWIN(IMPLET2$, "IMP-LET2", 0, 0)
                   '
            If AVAHOJA$ <> "NEWPAGE" Then
                 'Printer.EndDoc
                 Call SETSPOOL("ENDDOC", "")
            End If
            '
        Next COP%
        '
        Call CIERRARCH("*.*")
        DESCRIDOC$ = "Recibo Nro. " + NRORECI$ + " - " + TRIM$(RASOCLI$(NC%))
        FEX = FF%
        DESCRIDOC$ = DESCRIDOC$ + " - " + FECHATEX$(FEX)
        '
        Call SAVESPOOL(DESCRIDOC$, OKX%)
        Screen.MousePointer = 1
        If COMALTER%("Recibo Numero\'" + NRORECI$ + "'\Emitido y Registrado\\Repetir Impresion\Otro Recibo") = 1 Then GoTo 30
        '
        If AVAHOJA$ <> "NEWPAGE" Then
            'Printer.Orientation = PORIAN
            Call SETSPOOL("ORIENTATION", Str$(PORIAN))
        End If
        '
    End If
    Call CIERRARCH("*.*")
    '
End Sub

Sub ANURECI()
    '
    ' 1.- Ingresar Numero de Cliente
    '     **************************
    '
    Call SELECHO("DEUDOR1")
    NC% = Val(VALACT1$("DEUDOR1"))
    If NC% < 1 Then
        GoTo 999
    End If
    '
    FECHALIM% = Val(CONTROL$("", "CIEMEVEN"))
    '
    Screen.MousePointer = 11
    FENUE$ = CONTROL$("", "FENURECI")
    FLIXX% = FECHANUM(FENUE$)
    If FLIXX% <= FECHALIM% Then FLIXX% = FECHALIM% + 1
    If FLIXX% > HOY(HO$) Then FLIXX% = HOY(HO$)
    Call BLOQARCH("!SELCUERE")
    Call BLANARCH("!SELCUERE")
    '
301 For U& = 1 To ULTREG&("COBRAN")
        X$ = REGLEIDO("COBRAN", U&)
        If CVI(Left$(X$, 2)) = NC% Then
            NROLL& = CVS(Mid$(X$, 3, 4))
            IMPO# = CVD(Mid$(X$, 15, 8))
            FC% = CVI(Mid$(X$, 7, 2))
            If FC% >= FLIXX% Then
                If NROLL& > 0 Then
                    IMPOS$ = NUMSTRI$(IMPO#, 13, 2, 1)
                    NNXX# = CDbl(NROLL&)
                    DOCU$ = "Recibo No. " + TRIM$(Str$(NROLL&))
                    Y$ = REGBLAN$("!SELCUERE")
                    Call REPLA(Y$, MKS$(U&), 1, 4)
                    Call REPLA(Y$, DOCU$, 5, 18)
                    Call REPLA(Y$, IMPOS$, 23, 13)
                    Call REGAPP("!SELCUERE", Y$)
                End If
            End If
        End If
    Next U&
    Call CIERRARCH("!SELCUERE")
    Call FRESHECHO("!SELCUERE")
    '
    Screen.MousePointer = 1
    Call SELECHO("!SELCUERE")
    REGCUEC& = Val(VALACT1$("!SELCUERE"))
    If Val(REGCUEC&) < 1 Then
        GoTo 999
    End If
    '
    XX0$ = REGLEIDO$("COBRAN", REGCUEC&)
    FERECIBO% = CVI(Mid$(XX0$, 7, 2))
    NUTALONA% = CVI(Mid$(XX0$, 23, 2))
    IMPO# = Abs(CVD(Mid$(XX0$, 15, 8)))
    IMPOS$ = Left$("$ " + TRIM$(NUMSTRI$(IMPO#, 14, 2, 1)) + Space$(18), 18)
    NROLL& = CVS(Mid$(XX0$, 3, 4)): NNXX# = CDbl(NROLL&)
    X$ = "Recibo Nro" + NUMSTRI$(NNXX#, 8, 0, 1)
    Y$ = IMPOS$
    '
    If COMALTER%("Documento: " + X$ + "\  Importe: " + Y$ + "\\Abandona\Confirma Anulacion") <> 2 Then GoTo 999
    '
    ' anulacion propiamente dicha
    ' ***************************
    '
311 Screen.MousePointer = 11
    Call BLOQARCH("COBRAN")
    Call BLOQARCH("SVD202")
    Call BLOQARCH("CHECAR")
    Call BLOQARCH("APLICOB")
    '
    Call REPLA(XX0$, MKS$(0), 3, 4)
    Call REPLA(XX0$, MKI$(1), 7, 2)
    Call REPLA(XX0$, MKD$(0), 15, 8)
    Call GRAREG("COBRAN", XX0$, REGCUEC&)
    Call CIERRARCH("COBRAN")
    '
    NC1% = NC%
    RC& = REGICLI&(NC1%)
    DX$ = REGLEIDO$("DEUDOR2", RC&)
    SALDO# = CVD(Mid$(DX$, 109, 8)) + IMPO#
    Call REPLA(DX$, MKD$(SALDO#), 109, 8)
    Call GRAREG("DEUDOR2", DX$, RC&)
    Call CIERRARCH("DEUDOR2")
    '
    J& = 0
    For U& = 1 To ULTREG&("CHECAR")
      Z$ = REGLEIDO$("CHECAR", U&)
      If CVS(Mid$(Z$, 47, 4)) = NROLL& Then
        If CVI(Mid$(Z$, 45, 2)) = FERECIBO% Then
          If CVI(Mid$(Z$, 3, 2)) = NC% Then
            Call REPLA(Z$, MKD$(0), 33, 8)
            Call GRAREG("CHECAR", Z$, U&)
            ZZ1$ = Mid$(Z$, 51, 2)
          End If
        End If
      End If
312 Next U&
    Call CIERRARCH("CHECAR")
    '
    REBLA$ = REGBLAN$("BADACHE")
    For U& = 1 To ULTREG&("BADACHE")
      Z$ = REGLEIDO$("BADACHE", U&)
      If Val(Mid$(Z$, 66, 7)) = NROLL& Then
        If CVI(Mid$(Z$, 61, 2)) = FERECIBO% Then
          Call GRAREG("BADACHE", REBLA$, U&)
        End If
      End If
    Next U&
    Call CIERRARCH("BADACHE")
    '
    J& = 0
    IAJU# = IMPO#
    For U& = ULTREG&("APLICOB") To 1 Step -1
        Z$ = REGLEIDO$("APLICOB", U&)
        If Mid$(Z$, 5, 2) = "RB" Then
          If Val(Mid$(Z$, 15, 8)) = NROLL& Then
            FFA% = CVI(Left$(Z$, 2))
            NCA% = CVI(Mid$(Z$, 3, 2))
            ICA# = CVD(Mid$(Z$, 23, 8))
            dca$ = Mid$(Z$, 55, 18)
            TIA% = 2
            TIAS$ = Left$(dca$, 2)
            If TIAS$ = "ND" Then TIA% = 3
            If TIAS$ = "AS" Or TIAS$ = "AJ" Then TIA% = 0
            NRA& = Val(Mid$(dca$, 9, 8))
            VAA% = 1
            VAAS$ = Mid$(dca$, 18, 1)
            If VAAS$ = "B" Then VAA% = 2
            If VAAS$ = "C" Then VAA% = 3
            ido# = CVD(Mid$(Z$, 73, 8))
            '
            Call REPLA(Z$, MKD$(0), 23, 8)
            Call GRAREG("APLICOB", Z$, U&)
            '
            If ICA# > IAJU# Then ICA# = IAJU#
            
            ZZ$ = REGBLAN$("SVD202")
            Call REPLA(ZZ$, MKI$(NCA%), 1, 2)
            Call REPLA(ZZ$, Chr$(TIA%), 3, 1)
            Call REPLA(ZZ$, Chr$(VAA%), 4, 1)
            Call REPLA(ZZ$, MKS$(NRA&), 5, 4)
            Call REPLA(ZZ$, MKI$(FFA%), 9, 2)
            Call REPLA(ZZ$, MKD$(ICA#), 11, 8)
            Call REPLA(ZZ$, MKI$(FFA%), 19, 2)
            Call REPLA(ZZ$, MKD$(ido#), 21, 8)
            Call REPLA(ZZ$, String$(2, 255), 31, 2)
            Call REGAPP("SVD202", ZZ$)
            IAJU# = IAJU# - ICA#
            If IAJU# < 0.005 Then GoTo 315
          End If
        End If
    Next U&
    '
315 If Abs(IAJU#) >= 0.005 Then
        Call COMUNI("No se han Encontrado Todos los Documentos\Cancelados por este Recibo " + TRIM$(Str$(NROLL&)) + ".\  \La Diferencia Quedó Registrada\como Ajuste de Saldo.")
        '
        Z$ = REGBLAN$("SVD202")
        Call REPLA(ZZ$, MKI$(NC%), 1, 2)
        Call REPLA(ZZ$, Chr$(0), 3, 1)
        Call REPLA(ZZ$, Chr$(0), 4, 1)
        Call REPLA(ZZ$, MKS$(NROLL&), 5, 4)
        Call REPLA(ZZ$, MKI$(FERECIBO%), 9, 2)
        Call REPLA(ZZ$, MKD$(IAJU#), 11, 8)
        Call REPLA(ZZ$, MKI$(FERECIBO%), 19, 2)
        Call REPLA(ZZ$, MKD$(IMPO#), 21, 8)
        Call REPLA(ZZ$, String$(2, 255), 31, 2)
        Call REGAPP("SVD202", ZZ$)
        '
    End If
    '
    Call CIERRARCH("SVD202")
    Call CIERRARCH("APLICOB")
    '
    Call CIERRARCH("*.*")
    Call LIBARCH("COBRAN")
    Call LIBARCH("SVD202")
    Call LIBARCH("CHECAR")
    Call LIBARCH("APLICOB")
    Call COMUNI("Anulación Recibo " + TRIM$(Str$(NROLL&)) + " Completa")
    RESCUDEU.Text1.TEXT = TRIM$(Str$(NC%))
    RECICOMP.Hide
    Call FRESHALL
    PROLLAMA$ = "RECICOMP"
    RESCUDEU.Show
    Call FRESHALL
    '
    '
999 Call CIERRARCH("*.*")
    Call LIBARCH("!SELCUERE")
    Screen.MousePointer = 1
    '
End Sub
'

Sub REPIRECI()
    '
    ' 1.- Ingresar Numero de Cliente
    '     **************************
    '
10  Screen.MousePointer = 1
    Call SELECHO("DEUDOR1")
    NC% = Val(VALACT1$("DEUDOR1"))
    If NC% < 1 Then
        GoTo 999
    End If
    '
    Screen.MousePointer = 11
    FENUE$ = CONTROL$("", "FENURECI")
    FLIXX% = FECHANUM(FENUE$)
    If FLIXX% > HOY(HO$) Then FLIXX% = HOY(HO$)
    Call BLOQARCH("!SELCUERE")
    Call BLANARCH("!SELCUERE")
    '
301 RFF$ = RECFILT$("COBRAN", 1, MKI$(NC%))
    If Len(RFF$) < 4 Then
      Call MENSERR(24, "Sin Movimientos")
      GoTo 10
    End If
    '
    JJ& = 0
    For UL& = 1 To Len(RFF$) Step 4
      U& = CVS(Mid$(RFF$, UL&, 4))
      X$ = REGLEIDO("COBRAN", U&)
      NROLL& = CVS(Mid$(X$, 3, 4))
      IMPO# = CVD(Mid$(X$, 15, 8))
      FC% = CVI(Mid$(X$, 7, 2))
      If FC% >= FLIXX% Then
        If NROLL& > 0 Then
          IMPOS$ = NUMSTRI$(IMPO#, 13, 2, 1)
          NNXX# = CDbl(NROLL&)
          DOCU$ = "Recibo No. " + TRIM$(Str$(NROLL&))
          Y$ = REGBLAN$("!SELCUERE")
          Call REPLA(Y$, MKS$(U&), 1, 4)
          Call REPLA(Y$, DOCU$, 5, 18)
          Call REPLA(Y$, IMPOS$, 23, 13)
          JJ& = JJ& + 1
          Call GRAREG("!SELCUERE", Y$, JJ&)
        End If
      End If
    Next UL&
    Call CIERRARCH("!SELCUERE")
    Call FRESHECHO("!SELCUERE")
    Screen.MousePointer = 1
    '
    If JJ& < 1 Then
      Call MENSERR(24, "Sin Recibos Posteriores al " + FENUE$)
      GoTo 10
    End If
    '
350 Call SELECHO("!SELCUERE")
    REGCUEC& = Val(VALACT1$("!SELCUERE"))
    If Val(REGCUEC&) < 1 Then
        GoTo 999
    End If
    '
    XX0$ = REGLEIDO$("COBRAN", REGCUEC&)
    NROLL& = CVS(Mid$(XX0$, 3, 4))
    NCLIE% = CVI(Left$(XX0$, 2))
    '
    TIDOCUM$ = "Recibo Nro."
    NUDOCU$ = " " + TRIM$(Str$(NROLL&))
    RACLI$ = TRIM$(Left$(RASOCLI$(NCLIE%), 8))
    '
    Screen.MousePointer = 11
    For UI& = ULTREG&("PDOCLST") To 1 Step -1
      XX$ = Left$(REGLEIDO$("PDOCLST", UI&), 64)
      If InStr(XX$, TIDOCUM$) > 4 Then
        If InStr(XX$, NUDOCU$) > 4 Then
          If InStr(XX$, RACLI$) > 16 Then
            DOCUNU& = CVS(Left$(XX$, 4))
            GoTo 5
          End If
        End If
      End If
    Next UI&
    Screen.MousePointer = 1
    Call COMUNI("Imposible Abrir Documento Desde esta Aplicación")
    GoTo 10
    '
5   Call MUESTRADOC(DOCUNU&) 'PERMITE VER LOS DOCUMENTOS ANTIGUOS POR PANTALLA
    GoTo 350
    '
999 Call CIERRARCH("*.*")
    Call LIBARCH("*.*")
    '
    Screen.MousePointer = 1
    '
End Sub


Sub AJUSALD()
    '
    HOII% = HOY(HO$)
    FAJUI% = HOII%
    '
5   VDEF$ = MKI$(NC%) + MKI$(FAJUI%)
    CLSS$ = OBJPLA$("SELCLIAJU", VDEF$)
    If CLSS$ = "" Or CLSS$ = "<ESC>" Then GoTo 99
    NC% = CVI(CLSS$)
    If NC% < 1 Then GoTo 5
    CLI$ = RASOCLI$(NC%)
    FAJUI% = CVI(Mid$(CLSS$, 3, 2))
    If FAJUI% < 1 Or FAJUI% > HOII% Then GoTo 5
    '
    FECHALIM% = Val(CONTROL$("", "CIEMEVEN"))
    If FAJUI% <= FECHALIM% Then
        Call MENSERR(24, "Fecha no Válida.\  \Ese Período Mensual Está Cerrado.")
        GoTo 5
    End If
    '
    Screen.MousePointer = 11
    '
    SARRASFE# = 0: SALFINAL# = 0: SALDEUTO# = 0
    '
    NUSAL& = ULTREG&("FACTUR")
    For I& = 1 To NUSAL&
       Call TRACE(20, I&, NUSAL&)
       F0$ = REGLEIDO$("FACTUR", I&)
       F1$ = Left$(F0$, 2)
       '
       If CVI(F1$) = NC% Then
          '
          F3$ = Mid$(F0$, 7, 2)
          F6$ = Mid$(F0$, 13, 2)
          F7$ = Mid$(F0$, 15, 8)
          F8$ = Mid$(F0$, 23, 8)
          F9$ = Mid$(F0$, 31, 8)
          F12$ = Mid$(F0$, 49, 8)
          F13$ = Mid$(F0$, 57, 8)
          '
          FF% = CVI(F3$)
          '
          IMPO# = CVD(F7$) + CVD(F8$) + CVD(F9$) + CVD(F12$) + CVD(F13$)
          SIG% = 1: If IMPO# <= (-0.005) Then SIG% = (-1)
          IMPO# = Abs((Int(100 * IMPO# + 0.5)) / 100)
          TI% = Asc(F6$): If TI% < 0 Or TI% > 6 Then TI% = 2
          If TI% >= 4 Then If TI% <= 6 Then SIG% = -1
          '
          SALFINAL# = SALFINAL# + IMPO# * SIG%
          If FF% <= FAJUI% Then
              SARRASFE# = SARRASFE# + IMPO# * SIG%
          End If
          '
       End If
    Next I&
    '
    NUSAL& = ULTREG&("COBRAN")
    For I& = 1 To NUSAL&
       Call TRACE(20, I&, NUSAL&)
       C0$ = REGLEIDO$("COBRAN", I&)
       C1$ = Left$(C0$, 2)
       '
       If CVI(C1$) = NC% Then
          '
          C3$ = Mid$(C0$, 7, 2)
          C5$ = Mid$(C0$, 15, 8)
          '
          FF% = CVI(C3$)
          IMPO# = (Int(100 * CVD(C5$) + 0.5)) / 100
          '
          SALFINAL# = SALFINAL# - IMPO#
          If FF% <= FAJUI% Then
            SARRASFE# = SARRASFE# - IMPO#
          End If
          '
       End If
    Next I&
    '
    NUSAL& = ULTREG&("SVD202")
    For I& = 1 To NUSAL&
       Call TRACE(20, I&, NUSAL&)
       SS0$ = REGLEIDO$("SVD202", I&)
       SS1$ = Left$(SS0$, 2)
       If CVI(SS1$) = NC% Then
          '
          SS2$ = Mid$(SS0$, 3, 2)
          SS4$ = Mid$(SS0$, 9, 2)
          SS5$ = Mid$(SS0$, 11, 8)
          '
          FVIJ% = CVI(SS4$)
          If FVIJ% > 0 Then
             TI% = Asc(Mid$(SS2$, 1, 1))
             If TI% > 6 Then TI% = 2
             SALDEU# = CVD(SS5$)
             If TI% > 3 Then SALDEU# = (-1) * Abs(SALDEU#)
             '
             If Abs(SALDEU#) >= 0.005 Then
                SALDEUTO# = SALDEUTO# + SALDEU#
             End If
          End If
          '
       End If
    Next I&
    '
100 SALAJU# = SARRASFE#
    VDEF$ = CLSS$ + MKD$(SARRASFE#) + MKD$(SALAJU#) + MKI$(FAJUI%) + Space$(256)
    '
    Screen.MousePointer = 1
    '
    CLTT$ = OBJPLA$("CARAJUCLI", VDEF$)
    If CLTT$ = "" Or CLTT$ = "<ESC>" Then GoTo 5
    SALAJU# = CVD(Mid$(CLTT$, 13, 8))
    FEVEN% = CVI(Mid$(CLTT$, 21, 2))
    If FEVEN% < FAJUI% Then GoTo 100
    '
    MOTIVO$ = Mid$(CLTT$, 23)
    If TRIM$(MOTIVO$) = "" Then
       Call MENSERR(24, "Debe Indicar un Motivo del Ajuste")
       GoTo 100
    End If
    '
    DIFESAL# = SALAJU# - SARRASFE#
    DIFEDEU# = SALFINAL# - SALDEUTO# + DIFESAL#
    '
    If DIFESAL# < 0 Then
       FEVEN% = FAJUI%
    End If
    '
    Call CIERRARCH("FACTUR")
    Call CIERRARCH("SVD202")
    Call CIERRARCH("DEUDOR2")
    '
    Call BLOQARCH("FACTUR")
    Call BLOQARCH("SVD202")
    Call BLOQARCH("DEUDOR2")
    '
    SALFINAL# = SALFINAL# + DIFESAL#
    RGX& = REGICLI&(NC%)
    If RGX& > 0 Then
        If RGX& <= ULTREG&("DEUDOR2") Then
           DX$ = REGLEIDO$("DEUDOR2", RGX&)
           DX$ = REPLACE$(DX$, MKD$(SALFINAL#), 109, 8)
           Call GRAREG("DEUDOR2", DX$, RGX&)
        End If
    End If
    '
    Call CIERRARCH("DEUDOR2")
    '
    HAYAJU% = 0
    If Abs(DIFESAL#) >= 0.005 Or Abs(DIFEDEU#) >= 0.005 Then
       '
       Screen.MousePointer = 11
       '
       FF0$ = REGBLAN$("FACTUR")
       NUAJU& = 1 + ULTREG&("FACTUR")
       Call REPLA(FF0$, MKI$(NC%), 1, 2)
       Call REPLA(FF0$, MKS$(NUAJU&), 3, 4)
       Call REPLA(FF0$, MKI$(FAJUI%), 7, 2)
       Call REPLA(FF0$, MKI$(FAJUI%), 9, 2)
       Call REPLA(FF0$, String$(4, 0), 11, 4)
       Call REPLA(FF0$, MKD$(DIFESAL#), 15, 8)
       Call REPLA(FF0$, MKD$(DIFESAL#), 39, 8)
       Call REGAPP("FACTUR", FF0$)
       '
       SS0$ = REGBLAN$("SVD202")
       Call REPLA(SS0$, MKI$(NC%), 1, 2)
       Call REPLA(SS0$, String$(2, 0), 3, 2)
       Call REPLA(SS0$, MKS$(NUAJU&), 5, 4)
       Call REPLA(SS0$, MKI$(FEVEN%), 9, 2)
       Call REPLA(SS0$, MKD$(DIFEDEU#), 11, 8)
       Call REPLA(SS0$, MKI$(FAJUI%), 19, 2)
       Call REPLA(SS0$, MKD$(DIFESAL#), 21, 8)
       Call REGAPP("SVD202", SS0$)
       '
       Call CIERRARCH("FACTUR")
       Call CIERRARCH("SVD202")
       '
       HAYAJU% = 1
       '
    End If
    '
    Call CIERRARCH("*.*")
    Call LIBARCH("FACTUR")
    Call LIBARCH("SVD202")
    '
    If HAYAJU% > 0 Then
      Call GRAREG("!OBSERVOF", Left$(MOTIVO$, 64), 1)
      Call GRAREG("!OBSERVOF", Mid$(MOTIVO$, 65, 64), 2)
      Call GRAREG("!OBSERVOF", Mid$(MOTIVO$, 129, 64), 3)
      Call GRAREG("!OBSERVOF", Mid$(MOTIVO$, 193, 64), 4)
      Call SETULTREG("!OBSERVOF", 4)
      Call CIERRARCH("!OBSERVOF")
      '
      FORIPRE$ = TRIM$(CONTROL$("", "FORAJUSA"))
      If TRIM$(FORIPRE$) <> "" Then
         '
         FEX = HOY(HOS$)
         FECHDOC$ = FECHATEX$(FEX)
         NUMEDOC$ = TRIM$(Str$(NUAJU&))
         TIPODOC$ = "Ajuste Saldo Cliente"
         DESTIDOC% = NC%
         '
         FEX = FAJUI%
         CODPARAM$(1) = "FECH-VEN": DOCPARAM$(1) = FECHATEX$(FEX)
         CODPARAM$(2) = "IMPO-PEN": DOCPARAM$(2) = CSTRING$(MKD$(SARRASFE#), 4, 12, 2, 2)
         CODPARAM$(3) = "IMPO-CAN": DOCPARAM$(3) = CSTRING$(MKD$(DIFESAL#), 4, 12, 2, 2)
         CODPARAM$(4) = "IMP-TOTA": DOCPARAM$(4) = CSTRING$(MKD$(SALAJU#), 4, 12, 2, 2)
         CAPARDOC% = 4
         CACOLTAB1% = 0
         CAITAB1% = 0
         '
         Call PRINTDOC("FORAJUSA")   ' Remito de Facturacion
         '
      End If
      Call SETULTREG("!OBSERVOF", 0)
      Call CIERRARCH("!OBSERVOF")
      '
      Screen.MousePointer = 1
      Call COMUNI("Ajuste Completo")
      '
    End If
    '
    GoTo 5
    '
99  Call CIERRARCH("*.*")
    '
End Sub
'
Sub ANUFAC()
    '
    If DERACCE%("ANU-FAC", "Anulacion Facturas-Debitos-Creditos") < 1 Then
        GoTo 999
    End If
    '
    Call CIERRARCH("*.*")
    Call FRESHALL
    '
    Dim TIS$(10)
    TDAT$ = "AS,VC,FC,ND,RB,DC,NC,AS"
    For II% = 0 To 7
        TIS$(II%) = Mid$(TDAT$, 3 * II% + 1, 2)
    Next II%
    '
    Call CARSELMESA
5   Call SELECHO("SELMESA")
    MESASEL$ = VALACT1$("SELMESA")
    If TRIM$(MESASEL$) = "" Then
        GoTo 999
    End If
    '
    MMXX0$ = TRIM$(MESASEL$)
    MMXX1$ = MESANO$(MMXX0$, DES%, HAS%)
    '
    FECHALIM% = Val(CONTROL$("", "CIEMEVEN"))
    If HAS% <= FECHALIM% Then
       Call MENSERR(24, "Período no Válido.\  \Ese Período Mensual Está Cerrado.")
       GoTo 5
    End If
    '
10  Screen.MousePointer = 11
    Call FRESHECHO("!FACSEL1")
    Call BLANARCH("!FACSEL1")
    FIN& = ULTREG&("FACTUR")
    For I& = 1 To FIN&
        Call TRACE(20, I&, FIN&)
        AX$ = REGLEIDO$("FACTUR", I&)
        If CVI(Mid$(AX$, 7, 2)) >= DES% Then
            If CVI(Mid$(AX$, 7, 2)) <= HAS% Then
                NC% = CVI(Left$(AX$, 2))
                CLI$ = AJUSTI$(RASOCLI$(NC%), 30)
                NROL& = CVS(Mid$(AX$, 3, 4))
                FEX = CVI(Mid$(AX$, 7, 2))
                TI% = Asc(Mid$(AX$, 13, 1)): If TI% < 0 Or TI% > 7 Then TI% = 2
                TIX$ = TIS$(TI%)
                VA% = Asc(Mid$(AX$, 14, 1))
                VAX$ = "A": If VA% < 1 Or VA% > 2 Then VAX$ = "X"
                            If VA% = 2 Then VAX$ = "B"
                IMPO# = CVD(Mid$(AX$, 15, 8))
                IMPO# = IMPO# + CVD(Mid$(AX$, 23, 8))
                IMPO# = IMPO# + CVD(Mid$(AX$, 31, 8))
                IMPO# = IMPO# + CVD(Mid$(AX$, 49, 8))
                IMPO# = IMPO# + CVD(Mid$(AX$, 57, 8))
                SIG% = 1: If TI% > 3 Then SIG% = (-1)
                '
                If Abs(IMPO#) >= 0.005 Then
                    X$ = MKS$(I&) + TIX$ + "." + TRIM$(Str$(NROL&)) + "-" + VAX$
                    X$ = X$ + " - " + FECHATEX$(FEX)
                    X$ = X$ + " - " + TRIM$(CSTRING$(MKD$(Abs(IMPO#)), 3, 12, 2, 2))
                    X$ = X$ + " - " + CLI$
                    Call REGAPP("!FACSEL1", X$)
                End If
                '
            End If
        End If
    Next I&
    Call CIERRARCH("*.*")
    Screen.MousePointer = 1
    '
20  Call SELECHO("!FACSEL1")
    X$ = VALACT1$("!FACSEL1")
    If X$ = "" Then
        GoTo 5
    End If
    REFA& = Val(X$)
    If REFA& < 1 Or REFA& > ULTREG&("FACTUR") Then
        GoTo 999
    End If
    '
    AX$ = REGLEIDO$("FACTUR", REFA&)
    NC% = CVI(Left$(AX$, 2))
    NROL& = CVS(Mid$(AX$, 3, 4))
    FF% = CVI(Mid$(AX$, 7, 2))
    TI% = Asc(Mid$(AX$, 13, 1)): If TI% < 0 Or TI% > 7 Then TI% = 2
    TIX$ = TIS$(TI%)
    VA% = Asc(Mid$(AX$, 14, 1))
    VAX$ = "A": If VA% < 1 Or VA% > 2 Then VAX$ = "X"
                If VA% = 2 Then VAX$ = "B"
    NETO# = CVD(Mid$(AX$, 15, 8))
    IMPO# = CVD(Mid$(AX$, 15, 8))
    IMPO# = IMPO# + CVD(Mid$(AX$, 23, 8))
    IMPO# = IMPO# + CVD(Mid$(AX$, 31, 8))
    IMPO# = IMPO# + CVD(Mid$(AX$, 49, 8))
    IMPO# = IMPO# + CVD(Mid$(AX$, 57, 8))
    SIG% = 1: If TI% > 3 Then SIG% = (-1)
    NETO# = NETO# * SIG%: IMPO# = IMPO# * SIG%
    '
    VDEF$ = String$(40, 0)
    Call REPLA(VDEF$, TIX$ + "." + TRIM$(Str$(NROL&)) + "-" + VAX$, 1, 18)
    Call REPLA(VDEF$, MKI$(NC%), 19, 2)
    Call REPLA(VDEF$, MKI$(FF%), 21, 2)
    Call REPLA(VDEF$, MKD$(NETO#), 23, 8)
    Call REPLA(VDEF$, MKD$(IMPO#), 31, 8)
    '
    OBX$ = OBJPLA$("MUESTRAFAC", VDEF$)
    If OBX$ = "" Then GoTo 20
    '
    Screen.MousePointer = 11
    DOCANU$ = TIX$ + "." + TRIM$(Str$(NROL&))
    FIN& = ULTREG&("SVD202")
    '
    SUMADEU# = 0
    For I& = 1 To FIN&
        Call TRACE(20, I&, FIN&)
        X$ = REGLEIDO$("SVD202", I&)
        If CVI(Left$(X$, 2)) = NC% Then
            If CVS(Mid$(X$, 5, 4)) = NROL& Then
                If Asc(Mid$(X$, 3, 1)) = TI% Then
                    SG% = 1: If TI% >= 4 Then If TI% <= 6 Then SG% = -1
                    SALDO# = CVD(Mid$(X$, 11, 8))
                    IORIG# = CVD(Mid$(X$, 21, 8))
                    If TI% <> 0 Then SALDO# = Abs(SALDO#)
                    SALDO# = SG% * SALDO#
                    SUMADEU# = SUMADEU# + SALDO#
                End If
            End If
        End If
    Next I&
    '
    If Abs(IMPO# - SUMADEU#) > 0.005 Then
        Call COMUNI("La Anulacion de este Documento\no es Posible por Registrar\Cancelaciones Totales o Parciales.\  \Emita Notas de Debito, Credito o\Ajustes de Saldo según Corresponda.")
        GoTo 20
    End If
    '
    AX$ = REGLEIDO$("FACTUR", REFA&)
    SUCUVEN% = CVI(Mid$(AX$, 47, 2)) '\\\OT-06-0075-WAR-02/03/06///
    Call REPLA(AX$, String$(50, 0), 15, 50)
    Call REPLA(AX$, MKI$(SUCUVEN%), 47, 2) '\\\OT-06-0075-WAR-02/03/06///
    Call GRAREG("FACTUR", AX$, REFA&)
    Call CIERRARCH("FACTUR")
    '\\\OT-05-0524-WAR-25/07/05///
    If EXISTE%(LUDAT$ + "FACBONIF.DAT") > 0 Then
      Call ANUFABON(NC%, NROL&, FF%, TI%, VA%)
    End If
    '\\\OT-05-0524-WAR-FIN///
    For I& = 1 To FIN&
        Call TRACE(20, I&, FIN&)
        X$ = REGLEIDO$("SVD202", I&)
        If CVI(Left$(X$, 2)) = NC% Then
            If CVS(Mid$(X$, 5, 4)) = NROL& Then
                If Asc(Mid$(X$, 3, 1)) = TI% Then
                    Call REPLA(X$, MKD$(0), 11, 8)
                    Call REPLA(X$, MKD$(0), 21, 8)
                    Call GRAREG("SVD202", X$, I&)
                End If
            End If
        End If
    Next I&
    '
    RGXX& = REGICLI&(NC%)
    If RGXX& > 0 Then
        If RGXX& <= ULTREG&("DEUDOR2") Then
            TTXS$ = REGLEIDO$("DEUDOR2", RGXX&)
            SALNUE# = (Int(100 * CVD(Mid$(TTXS$, 109, 8)) + 0.5)) / 100 - (Int(100 * IMPO# + 0.5)) / 100
            Call REPLA(TTXS$, MKD$(SALNUE#), 109, 8)
            Call GRAREG("DEUDOR2", TTXS$, RGXX&)
        End If
    End If
    Call CIERRARCH("DEUDOR2")
    '
    ARMOVI$ = "MOVISTO"
    '
113 FIN& = ULTREG&(ARMOVI$)
    '
    For I& = 1 To FIN&
        Call TRACE(20, I&, FIN&)
        X$ = REGLEIDO$(ARMOVI$, I&)
        If TRIM$(Mid$(X$, 33, 10)) = DOCANU$ Then
            If CVI(Left$(X$, 2)) = FF% Then
                If CVI(Mid$(X$, 84, 2)) = NC% Then
                    CI% = CVI(Mid$(X$, 3, 2))
                    LOTE$ = Mid$(X$, 5, 16)
                    CMOV$ = Mid$(X$, 21, 2)
                    DOPRI$ = Mid$(X$, 23, 10)
                    DOSEC$ = DOCANU$
                    REFE$ = "Anulacion " + DOCANU$
                    VUNI# = CVD(Mid$(X$, 73, 8))
                    MONE$ = Mid$(X$, 81, 2)
                    DEPO% = Asc(Mid$(X$, 83, 1))
                    CANTI = CVD(Mid$(X$, 96, 8)) - CVD(Mid$(X$, 88, 8))
                    Call MOVISTO(FF%, CI%, LOTE$, CMOV$, DOPRI$, DOSEC$, REFE$, VUNI#, MONE$, NC%, DEPO%, CANTI)
                End If
            End If
        End If
    Next I&
    '
    Call CIERRARCH(ARMOVI$) '
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

    Call CIERRARCH("*.*")
    '
    Call COMUNI("Anulación Completa")
    GoTo 10
    '
999 Call CIERRARCH("*.*")
    If UCase$(App.EXEName) = "FACTUMAI" Then
         OPFACMAI.Option1.Value = True
         OPFACMAI.Option1.Refresh
      ElseIf UCase$(App.EXEName) = "FACTUDIP" Then
         OPFACDIP.Option1.Value = True
         OPFACDIP.Option1.Refresh
      ElseIf Left$(UCase$(App.EXEName), 6) = "FACTUR" Then
         Exit Sub
         'OPFACTURA.Option1.Value = True
         'OPFACTURA.Option1.Refresh
    End If
    '
End Sub
'
Sub ANUFORM()
   ANUFORMU.Show 1
End Sub
'\\\OT-05-0524-WAR-25/07/05///
Sub ANUFABON(NC2%, NROL2&, FF2%, TI2%, VA2%)
  'ANULA DOCUMENTO EN FACBONIF
  For J& = ULTREG("FACBONIF") To 1 Step -1
    X$ = REGLEIDO$("FACBONIF", J&)
    NC1% = CVI(Mid$(X$, 29, 2))
    NROL1& = CVS(Mid$(X$, 3, 4))
    FF1% = CVI(Mid$(X$, 7, 2))
    TI1% = Asc(Mid$(X$, 9, 1))
    VA1% = Asc(Mid$(X$, 10, 1))
    '
    If NC1% = NC2% Then
      If NROL1& = NROL2& Then
        If FF1% = FF2% Then
          If TI1% = TI2% Then
            If VA1% = VA2% Then
              'ELIMINO REGISTRO
              AX$ = REGLEIDO("FACBONIF", J&)
              Call REPLA(AX$, MKD$(0), 33, 8)
              Call REPLA(AX$, " ", 41, 16)
              Call REPLA(AX$, MKD$(0), 57, 8)
              Call REPLA(AX$, " ", 65, 16)
              Call REPLA(AX$, MKD$(0), 81, 8)
              Call REPLA(AX$, MKD$(0), 89, 8)
              Call REPLA(AX$, MKD$(0), 97, 8)
              Call GRAREG("FACBONIF", AX$, J&)
              Call CIERRARCH("FACBONIF")
              Exit Sub
            End If
          End If
        End If
      End If
    End If
  Next J&
  '
End Sub
'\\\OT-05-0524-WAR-FIN///
