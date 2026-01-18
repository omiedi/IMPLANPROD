Attribute VB_Name = "FACFUNSABO"
Public MODORECI%
Public FECHAVENCI(10), IMPOVENCI#(10), CAPLAVENCI%
'

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

Sub GRAFAC98(LU$, ARCHIFAC$, OPCI%, RRR%, FLIM%, TIPO%, VA%, CAFOR%, NRO, NC%, CLI$, IVA$, FF%, FV%, LPRE%, VENDE%, TOTAFAC#())
      '
      LU$ = LUDAT$
      Static UREG&
      Static NROMAX(3)
      Static CTX%
      Static TIPOVANT%
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
        COMPRAS% = 0

        If CTX% = 0 Then
              CTX% = 1
        End If


110   ' determina NUFAC&

        NUFAC& = ULTREG&("FACTUR")
111     If NUFAC& < 1 Then NUFAC& = 0: GoTo 120
        F$ = REGLEIDO$("FACTUR", NUFAC&): GoSub 10000

120   ' determina NROMAX(VA%)

        If NROMAX(VA%) > 0 Then
            If TIPOVA% = TIPOVANT% Then
                GoTo 200
            End If
        End If
        TIPOVANT% = TIPOVA%
        NROMAX(VA%) = 0

        For MXC& = ULTREG&("FACTUR") To 1 Step -1
                F$ = REGLEIDO$("FACTUR", MXC&): GoSub 10000
                If CVI(F3$) < FLIM% Then GoTo 180
                TIX% = Asc(F6$): VAX% = Asc(Mid$(F6$, 2))
                              If VAX% < 0 Or VAX% > 3 Then VAX% = 1
                              If VAX% <> VA% Then GoTo 180
                If TIPO% > 6 Then GoTo 150

                If TIX% < 1 Or TIX% = 4 Or TIX% = 5 Or TIX% > 6 Then GoTo 180 ' descarta otros documentos
                If CAFOR% > 1 Then GoTo 150 ' numeracion SEPARADA para facturas,debitos y creditos
                If CVS(F2$) > NROMAX(VAX%) Then NROMAX(VAX%) = CVS(F2$)
                GoTo 180

150             If (TIPO% = 1 Or TIPO% = 2) Then If (TIX% = 1 Or TIX% = 2) Then TIX% = TIPO%
                If TIX% <> TIPO% Then GoTo 180
                If CVS(F2$) > NROMAX(VAX%) Then NROMAX(VAX%) = CVS(F2$)

180             If NROMAX(VA%) > 0 Then
                    If MXC& < ULTREG&("FACTUR") - 100 Then
                        GoTo 200
                    End If
                End If

        Next MXC&

200     If COMPRAS% = 1 Then CAFOR% = 2
        If OPCI% <> 1 Then GoTo 300
        NRO = NROMAX(VA%) + 1
        GoTo 40000

300     If OPCI% <> 2 Then GoTo 400
        If COMPRAS% = 0 Then If NRO >= NROMAX(VA%) + 1 Then GoTo 40000 ' acepta valor

        For MXC& = ULTREG&("FACTUR") To 1 Step -1
                F$ = REGLEIDO$("FACTUR", MXC&): GoSub 10000
                If CVI(F3$) < FLIM% Then GoTo 380
                TIX% = Asc(F6$): VAX% = Asc(Mid$(F6$, 2)): If VAX% < 0 Or VAX% > 3 Then VAX% = 1
                If VAX% <> VA% Then GoTo 380
                If TIPO% > 6 Then GoTo 350

                If TIX% < 1 Or TIX% = 4 Or TIX% = 5 Or TIX% > 6 Then GoTo 380 ' descarta otros documentos

                If CAFOR% > 1 Then GoTo 350 ' numeracion SEPARADA para facturas,debitos y creditos
                If CVS(F2$) = NRO Then NRO = NROMAX(VA%) + 1: GoTo 40000
                GoTo 380

350             If (TIPO% = 1 Or TIPO% = 2) Then If (TIX% = 1 Or TIX% = 2) Then TIX% = TIPO%
                If TIX% <> TIPO% Then GoTo 380
                If COMPRAS% = 1 Then If CVI(F1$) <> NC% Then GoTo 380
                If CVS(F2$) = NRO Then NRO = NROMAX(VA%) + 1: GoTo 40000

380     Next MXC&
390     GoTo 40000

' vuelve a determinar NUFAC& y NROMAX(1,2,3)

400     If OPCI% <> 3 Then GoTo 800     ' opci%=3 - reserva registro y numero

        If COMPRAS% = 1 Then GoTo 600

        If UREG& = 0 Then
            UREG& = ULTREG&("FACTUR") - 128
            If UREG& < 0 Then UREG& = 0
        End If

        For MXC& = UREG& + 1 To ULTREG&("FACTUR")
                F$ = REGLEIDO$("FACTUR", MXC&): GoSub 10000
                If CVI(F1$) = 0 Then GoTo 490
                NUFAC& = MXC&: If CVI(F3$) < FLIM% Then GoTo 480
                TIX% = Asc(F6$): VAX% = Asc(Mid$(F6$, 2)): If VAX% < 0 Or VAX% > 3 Then VAX% = 1
                If TIPO% > 6 Then GoTo 450

                If TIX% < 1 Or TIX% = 4 Or TIX% = 5 Or TIX% > 6 Then GoTo 480 ' descarta otros documentos

                If CAFOR% > 1 Then GoTo 450 ' numeracion SEPARADA para facturas,debitos y creditos
                If CVS(F2$) > NROMAX(VAX%) Then NROMAX(VAX%) = CVS(F2$)
                GoTo 480

450             If (TIPO% = 1 Or TIPO% = 2) Then If (TIX% = 1 Or TIX% = 2) Then TIX% = TIPO%
                If TIX% <> TIPO% Then GoTo 480
                If CVS(F2$) > NROMAX(VAX%) Then NROMAX(VAX%) = CVS(F2$)

480     Next MXC&

490     ' CONTINUE

500     If NRO >= NROMAX(VA%) + 1 Then GoTo 600 ' acepta valor

        For MXC& = NUFAC& To 1 Step -1
                F$ = REGLEIDO$("FACTUR", MXC&): GoSub 10000
                If CVI(F3$) < FLIM% Then GoTo 580
                TIX% = Asc(F6$): VAX% = Asc(Mid$(F6$, 2)): If VAX% < 0 Or VAX% > 3 Then VAX% = 1
                If VAX% <> VA% Then GoTo 580
                If TIPO% > 6 Then GoTo 550

                If TIX% < 1 Or TIX% = 4 Or TIX% = 5 Or TIX% > 6 Then GoTo 580 ' descarta otros documentos

                If CAFOR% > 1 Then GoTo 550 ' numeracion SEPARADA para facturas,debitos y creditos
                If CVS(F2$) = NRO Then Call MENSERR(24, "Documento" + Str$(NRO) + " ya Utilizado - Use Numero" + Str$(NROMAX(VA%) + 1)): NRO = NROMAX(VA%) + 1: GoTo 600
                GoTo 580

550             If (TIPO% = 1 Or TIPO% = 2) Then If (TIX% = 1 Or TIX% = 2) Then TIX% = TIPO%
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
        F$ = F$ + MKI$(LPRE%)
        F$ = F$ + Chr$(TIPO%) + Chr$(VA%)
        F$ = F$ + MKD$(0)
        F$ = F$ + MKD$(0)
        F$ = F$ + MKD$(0)
        F$ = F$ + MKD$(0)
        F$ = F$ + MKI$(VENDE%)
        F$ = F$ + MKD$(0)
        F$ = F$ + MKD$(0)

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
        If VA% = 2 Then
                TOTA# = NEFA# + NIVA# + NNIN#
                NNIN# = 0
                TOTB# = TOTA# - IINT#
                ALIC = 1: If TOTA# > 0 Then ALIC = NEFA# / TOTA#
                NEFA# = TOTB# * ALIC
                ALIC = 1: If TOTA# > 0 Then ALIC = NIVA# / TOTA#
                NIVA# = TOTB# * ALIC
                TOTA# = TOTA# + PERC#
                TOTAFAC#(1, 8) = TOTA#
        End If

        F$ = REGLEIDO$("FACTUR", UREG&)
        F$ = REPLACE$(F$, MKD$(NEFA#), 15, 8)
        F$ = REPLACE$(F$, MKD$(NIVA#), 23, 8)
        F$ = REPLACE$(F$, MKD$(NNIN#), 31, 8)
        F$ = REPLACE$(F$, MKD$(TOTA#), 39, 8)
        F$ = REPLACE$(F$, MKI$(VENDE%), 47, 8)
        F$ = REPLACE$(F$, MKD$(IINT#), 49, 8)
        F$ = REPLACE$(F$, MKD$(PERC#), 57, 8)
        Call GRAREG("FACTUR", F$, UREG&)
        Call CIERRARCH("FACTUR")
        '
        GoTo 40000

1000    CU$ = REGBLAN$("CUITS")
        CU$ = REPLACE$(CU$, CLI$, 1, 30)
        CU$ = REPLACE$(CU$, IVA$, 31, 15)
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
        Return


40000  End Sub
'

Sub GRAM2000(NC%, NRO, REMI, FF%)
        '
        NC1% = NC%: If NC% < 1 Then NC1% = 10001
        REFE$ = ECOARCH$("DEUDOR1", MKI$(NC%))
        FIN& = ULTREG&("!CARDEFAC")
        '
        For I& = 1 To FIN&
             Call TRACE(20, I&, FIN&)
             X$ = REGLEIDO$("!CARDEFAC", I&)
             CANTAA = CVS(Left$(X$, 4))
             If CANTAA > 0 Then
                 CI% = CVI(Mid$(X$, 5, 2))
                 '
                 COD$ = CODEXT$(CI%)
                 PRE# = CDbl(CVS(Mid$(X$, 429, 4)))
                 CMOV$ = "RT"
                 DOPRI$ = "RT." + TRIM$(Str$(REMI))
                 DOSEC$ = "FC." + TRIM$(Str$(NRO))
                 '
                 REPED& = CVS(Mid$(X$, 509, 4))
                 If REPED& < 0 Then
                     REMI = Val(Mid$(X$, 131, 16))
                     DOPRI$ = "RT." + TRIM$(Str$(REMI))
                     ENCURE% = 0
                     GoSub 2000
                     If ENCURE% > 0 Then GoTo 1995
                 End If
                 '
                 X$ = ""
                 SIG% = (-1)
                 '
                 If NRO < 0 Then
                    CMOV$ = "RD"
                    DOPRI$ = "DV." + TRIM$(Str$(Abs(REMI)))
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
                 ' correccion falla sabo
                 If InStr(EMPRELI$, "SABO") > 0 Then
                    If Val(Mid$(DOPRI$, 4)) < 1 Then
                       If Left$(DOSEC$, 3) = "FC." Then
                          DOPRI$ = "RT." + TRIM$(Str$(NRO))
                       End If
                    End If
                 End If
                 '
                 Call MOVISTO(FF%, CI%, LOTE$, CMOV$, DOPRI$, DOSEC$, REFE$, PRE#, "$", NC%, 1, SIG% * CANTAA)
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
             End If
             '
1998    Next I&
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
2090    Return
        '
3000    End Sub
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

Sub GRADEU(LU$, ARCHIDEU$, RDX%, TIPO%, VA%, NRO, NC%, FFAC%, COPAGI%, TOTFAC#)
        '
        Dim FVI%(6), PLAII%(6), IM#(6), FEX(3)
        '
        ARDEUX$ = ARCHIDEU$
        PPXX% = InStr(ARDEUX$, "/")
        If PPXX% > 0 Then
          ARDEUX$ = Left$(ARDEUX$, PPXX% - 1)
        End If
        '
        If TIPO% > 3 Then
            NUPAG% = 1
            FVI%(1) = FFAC%
            IM#(1) = TOTFAC#
            GoTo 410
        End If


      ' busca condicion de pago segun COPAGI%

200     Call BUCONPA(LU$, COPAGI%, COPAG$, FREF%, PLA$)



      ' determina vencimientos segun total factura y condicion de pago

300     FF = FFAC%: If FF < 1 Then FF = FECHANUM("01/01/80")
        FFS$ = FECHATEX$(FF)
        PLA$ = Left$(PLA$ + String$(6, 0), 6)

        If (FREF% < 1 Or FREF% > 4) Then
                MENSERR 24, "Mal Fecha Referencia en Cond.Pago Nro." + Str$(COPAGI%)
                FREF% = 1
        End If

        If FREF% = 2 Then FREF% = 1   ' AUN NO FUNCIONA SEMANA DE FACTURA
        If FREF% <> 3 Then GoTo 305
        If Val(Left$(FFS$, 2)) <= 15 Then FEBAS$ = "15" + Mid$(FFS$, 3)
        If Val(Left$(FFS$, 2)) > 15 Then FEBAS$ = "31" + Mid$(FFS$, 3)
305     If FREF% = 1 Then FEBAS$ = FFS$
        If FREF% = 4 Then FEBAS$ = "31" + Mid$(FFS$, 3)

        For I% = 1 To 6: FVI%(I%) = 0: Next I%

        For I% = 1 To 6
                PLAI% = Asc(Mid$(PLA$, I%, 1)): PLAII%(I%) = PLAI%
                If I% > 1 Then If (PLAII%(I%) < PLAII%(I% - 1) Or PLAII%(I%) <= 0) Then GoTo 310

                        FE$ = FEBAS$: GoSub 30500
                        FEX(1) = FEX(1) + PLAI%
306                     If FEX(1) > 31 Then FEX(2) = FEX(2) + 1: FEX(1) = FEX(1) - 31: GoTo 306
307                     If FEX(2) > 12 Then FEX(3) = FEX(3) + 1: FEX(2) = FEX(2) - 12: GoTo 307
                        GoSub 30580
                FVI%(I%) = FEXU: NUPAG% = I%
        Next I%

310     FVII% = FVI%(NUPAG%): If NUPAG% < 2 Then FVII% = FFAC%

      ' determina importes segun numero de pagos

400     If NUPAG% < 1 Then NUPAG% = 1
        IM#(1) = (Int(100 * TOTFAC# / NUPAG% + 0.5)) / 100

        For H% = 2 To NUPAG% - 1
                IM#(H%) = IM#(1)
        Next H%
        IM#(NUPAG%) = TOTFAC#

        For H% = 1 To NUPAG% - 1
                IM#(NUPAG%) = IM#(NUPAG%) - IM#(H%)
        Next H%

410     For I% = 1 To NUPAG%
                SS$ = MKI$(NC%) + Chr$(TIPO%) + Chr$(VA%) + MKS$(NRO)
                SS$ = SS$ + MKI$(FVI%(I%)) + MKD$(IM#(I%))
                SS$ = SS$ + MKI$(FFAC%) + MKD$(TOTFAC#) + String$(2, 0) + String$(2, 255)
                Call REGAPP(ARDEUX$, SS$)
                FECHAVENCI(I%) = FVI%(I%)
                IMPOVENCI#(I%) = IM#(I%)
        Next I%
        CAPLAVENCI% = NUPAG%
        '
        Call CIERRARCH(ARDEUX$)
        '
        Exit Sub
        '
30500   AKII% = 0
        For II% = 1 To 3
            FEX(II%) = Val(Mid$(FE$, 3 * II% - 2, 2))
            If II% = 3 Then
                If FEX(II%) < 60 Then
                    FEX(II%) = FEX(II%) + 100
                End If
            End If
            If FEX(II%) < 1 Then AKII% = 2: GoTo 30595
        Next II%
        '
        If ((FEX(1) > 31) Or (FEX(2) > 12)) Then AKII% = 2: GoTo 30595
30580   FEXU = FEX(1) + 32 * FEX(2) + 32 * 13 * (FEX(3) - 80)
30595   Return


40000   End Sub
'

Sub BUCONPA(LU$, COPAGI%, COPAG$, FREF%, PLA$)
        '
        If ULTREG&("CONPAG") < 1 Then
                Call MENSERR(24, "No Existe Archivo de Cond. Pago en U.Logica '" + LU$ + "'")
                COPAGI% = 0: COPAG$ = "CONTADO" + Space$(28): FREF% = 1: PLA$ = String$(6, 0)
                Exit Sub
        End If

200     If COPAGI% < 0 Or COPAGI% > 255 Then
                Call MENSERR(24, "Numero de Cond. Pago Fuera de Rango")
                COPAGI% = 0: COPAG$ = "CONTADO" + Space$(28): FREF% = 1: PLA$ = String$(6, 0)
                GoTo 40000
        End If

300     If COPAGI% = 0 Then
                COPAGI% = 0: COPAG$ = "CONTADO" + Space$(28): FREF% = 1: PLA$ = String$(6, 0)
                GoTo 40000
        End If

400     For I& = 1 To ULTREG("CONPAG")
                LL$ = REGLEIDO$("CONPAG", I&)
                L1$ = Left$(LL$, 1)
                If Asc(L1$) = 0 Then
                        Call MENSERR(24, "Condicion de Pago Numero" + Str$(COPAGI%) + " no Figura en Tabla")
                        COPAGI% = 0: COPAG$ = "CONTADO" + Space$(28): FREF% = 1: PLA$ = String$(6, 0)
                        Exit Sub
                End If
                '
                If Asc(L1$) = COPAGI% Then
                    COPAG$ = Mid$(LL$, 2, 35)
                    FREF% = Asc(Mid$(LL$, 37, 1))
                    PLA$ = Mid$(LL$, 38, 6)
                    Exit Sub
                End If
        Next I&

        Call MENSERR(24, "Condicion de Pago Numero" + Str$(COPAGI%) + " no Figura en Tabla")
        COPAGI% = 0
        COPAG$ = "CONTADO" + Space$(28)
        FREF% = 1: PLA$ = String$(6, 0)
        '
40000   End Sub
   

Sub BLAPEFAC(REMI, FF%)  '  blanqueo pedidos pendientes - asociado a FAC2000
        '
        Screen.MousePointer = 11
        '
        FIN& = ULTREG&("!CARDEFAC")
        For U& = 1 To FIN&
           X$ = REGLEIDO$("!CARDEFAC", U&)
           CIENT% = CVI(Mid$(X$, 5, 2))
           CAENTRE = CVS(Left$(X$, 4))
           If CAENTRE > 0 Then
              REGPEDPE& = CVS(Mid$(X$, 509, 4))
              If REGPEDPE& > 0 Then
                If REGPEDPE& <= ULTREG&("PEDDETP") Then
                   Y$ = REGLEIDO$("PEDDETP", REGPEDPE&)
                   If CVI(Mid$(Y$, 33, 2)) = CIENT% Then ' debe coincidir codigo
                      TOENTRE = CAENTRE + CVS(Mid$(Y$, 95, 4))
                      SALPEN = CVS(Mid$(Y$, 79, 4)) - TOENTRE: If SALPEN < 0 Then SALPEN = 0
                      '
                      CASIG = 0: If Mid$(Y$, 75, 4) <> "    " Then CASIG = CVS(Mid$(Y$, 75, 4))
                      CASIG = CASIG - CAENTRE: If CASIG < 0 Then CASIG = 0
                      Call REPLA(Y$, MKS$(CASIG), 75, 4)
                      '
                      Call REPLA(Y$, MKS$(TOENTRE), 95, 4)
                      Call REPLA(Y$, MKS$(SALPEN), 99, 4)
                      Call REPLA(Y$, MKI$(FF%), 103, 2)
                      Call REPLA(Y$, MKS$(REMI), 105, 4)
                      Call GRAREG("PEDDETP", Y$, REGPEDPE&)
                      '
                      REGPEDA& = CVS(Mid$(Y$, 109, 4))
                      If REGPEDA& > 0 Then
                          If REGPEDA& <= ULTREG&("PEDDETA") Then
                              Z$ = REGLEIDO$("PEDDETA", REGPEDA&)
                              If Left$(Z$, 30) = Left$(Y$, 30) Then
                                If Mid$(Z$, 33, 42) = Mid$(Y$, 33, 42) Then
                                  Call REPLA(Z$, MKS$(CASIG), 75, 4)
                                  Call REPLA(Z$, MKS$(TOENTRE), 95, 4)
                                  Call REPLA(Z$, MKS$(SALPEN), 99, 4)
                                  Call REPLA(Z$, MKI$(FF%), 103, 2)
                                  Call REPLA(Z$, MKS$(REMI), 105, 4)
                                  Call GRAREG("PEDDETA", Z$, REGPEDA&)
                                End If
                              End If
                          End If
                      End If
                      '
                   End If
                End If
              End If
              '
              REGODES& = CVS(Mid$(X$, 505, 4))
              If REGODES& > 0 Then
                If REGODES& <= ULTREG&("ORDESPA") Then
                   Y$ = REGLEIDO$("ORDESPA", REGODES&)
                   If CVI(Mid$(Y$, 37, 2)) = CIENT% Then ' debe coincidir codigo
                      TOENTRE = CAENTRE + CVS(Mid$(Y$, 83, 4))
                      SALPEN = CVS(Mid$(Y$, 79, 4)) - TOENTRE
                      If SALPEN < 0 Then SALPEN = 0
                      '
                      Call REPLA(Y$, MKS$(TOENTRE), 83, 4)
                      Call REPLA(Y$, MKS$(SALPEN), 87, 4)
                      Call REPLA(Y$, MKI$(FF%), 91, 2)
                      Call REPLA(Y$, MKS$(REMI), 99, 4)
                      Call REPLA(Y$, MKS$(REMI), 103, 4)
                      Call GRAREG("ORDESPA", Y$, REGODES&)
                      '
                   End If
                End If
              End If
           End If
        Next U&
        '
        For U& = 1 To FIN&
           X$ = REGLEIDO$("!CARDEFAC", U&)
           CIENT% = CVI(Mid$(X$, 5, 2))
           CAENTRE = CVS(Left$(X$, 4))
           If CAENTRE > 0 Then
              REGPEDPE& = CVS(Mid$(X$, 509, 4))
              If REGPEDPE& < 0 Then
                REGREMPE& = Abs(REGPEDPE&)
                If REGREMPE& <= ULTREG&("REMPEN") Then
                   Y$ = REGLEIDO$("REMPEN", REGREMPE&)
                   '
                   Call REPLA(Y$, String$(96, 0), 1, 96)
                   Call GRAREG("REMPEN", Y$, REGREMPE&)
                   '
                End If
              End If
           End If
        Next U&
        '
        Z$ = ""
        Call CIERRARCH("*.*")
        Screen.MousePointer = 1
        '
End Sub
'

Sub BLANPEDP()
    '
    Call BLOQARCH("REMPEN")
    Call BLOQARCH("PEDDETP")
    Call BLOQARCH("ORDESPA")
    Call BLOQARCH("ORDESPB")
    '
    Screen.MousePointer = 11
        '
        Dim NPED&(16384)
        PEDANT& = 0
        CAPEDP% = 0
        '
        Call BLOQARCH("PEDDETP")
        FIN& = ULTREG&("PEDDETP")
        FIN1& = FIN&: If FIN1& < 1 Then FIN1& = 1
        '
        PEDANT& = 0: J& = 0: QQ% = 0
        For I& = 1 To FIN&
          Call TRACE(20, I&, FIN&)
          X$ = REGLEIDO$("PEDDETP", I&)
           If CVS(Mid$(X$, 79, 4)) - CVS(Mid$(X$, 95, 4)) > 0.05 Then
               J& = J& + 1
               Call GRAREG("PEDDETP", X$, J&)
               NPEDU& = CVS(Left$(X$, 4))
               If NPEDU& <> PEDANT& Then
                  QQ% = QQ% + 1
                  NPED&(QQ%) = NPEDU&
                  CAPEDP% = CAPEDP% + 1
                  PEDANT& = NPEDU&
               End If
           End If
        Next I&
        '
        If J& < I& Then
           For K& = J& + 1 To FIN&
              Call GRAREG("PEDDETP", String$(128, 0), K&)
           Next K&
           '
           FIN& = ULTREG&("PEDENCA")
           FIN1& = FIN&: If FIN1& < 1 Then FIN1& = 1
           NPA% = FILENBR%("PEDENCA")
           '
           For I& = 1 To FIN&
              Call TRACE(20, I&, FIN&)
              X$ = REGLEIDO$("PEDENCA", I&)
              STAT% = Asc(Mid$(X$, 267, 1))
              If STAT% <= 1 Then
                 NPEDU& = CVS(Left$(X$, 4))
                 For QR% = 1 To QQ%
                    If NPED&(QR%) = NPEDU& Then GoTo 999
                 Next QR%
                 Call REPLA(X$, Chr$(3), 267, 1)
                 Call GRAREG("PEDENCA", X$, I&)
              End If
999        Next I&
           '
        End If
        '
        Call BLOQARCH("REMPEN")
        FIN& = ULTREG&("REMPEN")
        FIN1& = FIN&: If FIN1& < 1 Then FIN1& = 1
        '
        J& = 0
        For I& = 1 To FIN&
          Call TRACE(20, I&, FIN&)
          X$ = REGLEIDO$("REMPEN", I&)
          If X$ <> String$(96, 0) Then
               J& = J& + 1
               Call GRAREG("REMPEN", X$, J&)
           End If
        Next I&
        '
        Call SETULTREG&("REMPEN", J&)
        '
        Call BLOQARCH("ORDESPA")
        FIN& = ULTREG&("ORDESPA")
        FIN1& = FIN&: If FIN1& < 1 Then FIN1& = 1
        '
        J& = 0
        For I& = 1 To FIN&
          Call TRACE(20, I&, FIN&)
          X$ = REGLEIDO$("ORDESPA", I&)
          If I& < FIN& And CVS(Mid$(X$, 83, 4)) >= 0.05 Then
               Call REGAPP("ORDESPB", X$)
            Else
               J& = J& + 1
               Call GRAREG("ORDESPA", X$, J&)
           End If
        Next I&
        '
        Call SETULTREG&("ORDESPA", J&)
        Call CIERRARCH("ORDESPA")
        '
    Screen.MousePointer = 1
    Call CIERRARCH("*.*")
    '
End Sub
'

Sub DESCUEN(PORDES, DESC$)
    '
    PORDES = 0
    DESCX$ = REPLACAR$(DESC$, "%", " ")
    '
    If InStr(DESCX$, "+") < 1 Then
        PORDES = Val(DESCX$)
        Exit Sub
    End If
    '
    XDC% = InStr(DESCX$, "+")
    DESC1 = Val(Left$(DESCX$, XDC% - 1))
    DESC2 = Val(Mid$(DESCX$, XDC% + 1))
    PORDES = 100 * (1 - (1 - DESC1 / 100) * (1 - DESC2 / 100))
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
Sub NURECIBO(NROPRO&)
    '
    Screen.MousePointer = 11
    '
    FENUE$ = CONTROL$("", "FENURECI")
    FLIM% = FECHANUM(FENUE$)
    If FLIM% > HOY(HO$) Then FLIM% = HOY(HO$)
    '
    CAMOLE% = 0
    NROMA& = 0
    For U& = ULTREG&("COBRAN") To 1 Step -1
        XX$ = REGLEIDO$("COBRAN", U&)
        CAMOLE% = CAMOLE% + 1
        If CAMOLE% > 512 Then GoTo 90
        '
        FEMO% = CVI(Mid$(XX$, 7, 2))
        If FEMO% >= FLIM% Then
           NROLE& = CVS(Mid$(XX$, 3, 4))
           If NROLE& = NROPRO& Then
               NROPRO& = 0
           End If
           If NROLE& > NROMA& Then NROMA& = NROLE&
        End If
    Next U&
    '
90  If NROPRO& < 1 Then NROPRO& = 1 + NROMA&
    Screen.MousePointer = 1
    '
End Sub
'

Sub SIGUERECI1()
    '
    If MODORECI% = 0 Then
        MODORECI% = 1       ' por emision
        If CONTROL$("RECISTA", "MODORDE") = "VENCIM" Then
           MODORECI% = 2
        End If
    End If
    '
    NCXI = Val(FORECISTA.Text1.TEXT)
    If NCXI < 1 Or NCXI > 32767 Then
        Call MENSERR(24, "Numero de Cuenta no Válido")
        FORECISTA.Text1.SetFocus
        Exit Sub
    End If
    '
    NC% = NCXI
    If ECOARCH$("DEUDOR1", MKI$(NC%)) = "" Then
        Call MENSERR(24, "Cuenta no Figura\en Base de Datos de Clientes")
        FORECISTA.Text1.SetFocus
        Exit Sub
    End If
    '
    FE$ = TRIM$(FORECISTA.Text9.TEXT)
    FF% = FECHANUM(FE$)
    HOI% = HOY(HO$)
    If FF% < 1 Or FF% > HOI% + 1 Then
        Call MENSERR(24, "Fecha no Válida")
        FORECISTA.Text9.SetFocus
        Exit Sub
    End If
    '
    If DIENTRE%(FF%, HOI%) > 30 Then
        Call MENSERR(24, "Fecha no Válida")
        FORECISTA.Text9.SetFocus
        Exit Sub
    End If
   '
    Call CIERRARCH("*.*")
    '
    Screen.MousePointer = 11
    NC% = Val(FORECISTA.Text1.TEXT)
    If NC% > 0 Then
        FORECISTA.Text2.TEXT = RASOCLI$(NC%)
        FORECISTA.Text3.TEXT = DOMICLI$(NC%)
        FORECISTA.Text4.TEXT = CPOSCLI$(NC%)
        FORECISTA.Text5.TEXT = LOCACLI$(NC%)
        'FORECISTA.Text7.TEXT = TEL1CLI$(NC%)
        'FORECISTA.Text10.TEXT = CSTRING$(MKD$(SALDCLI#(NC%)), 4, 13, 2, 2)
        FORECISTA.Refresh
    End If
    '
100 FORECISTA.List1.Clear
    FORECISTA.List1.Refresh
    FORECISTA.List2.Clear
    FORECISTA.List2.Refresh
    '
    MAXDEU& = ULTREG&("SVD202")
    RFD$ = RECFILT$("SVD202", 1, MKI$(NC%))
    '
    RESCUDEU.LINDICE.Clear
    For U& = 1 To Len(RFD$) Step 4
        REG& = CVS(Mid$(RFD$, U&, 4))
        X$ = REGLEIDO$("SVD202", REG&)
        If CVI(Left$(X$, 2)) <> NC% Then
            X$ = REGLEIDO$("SVD202", 1)
            Call GRAREG("SVD202", X$, 1)
            Call CIERRARCH("SVD202")
            GoTo 100
        End If
        IMPO# = CVD(Mid$(X$, 21, 8))
        SALDO# = CVD(Mid$(X$, 11, 8))
        If Abs(SALDO#) >= 0.005 Then
          If Sgn(IMPO#) <> Sgn(SALDO#) Or Abs(SALDO#) > Abs(IMPO#) Then
            ABUS$ = Left$(X$, 8)
            For UU& = 1 To Len(RFD$) Step 4
                REG& = CVS(Mid$(RFD$, UU&, 4))
                X$ = REGLEIDO$("SVD202", REG&)
                If Left$(X$, 8) = ABUS$ Then
                    Call GRAREG("SVD202", String$(32, 0), REG&)
                End If
            Next UU&
            GoTo 100
          End If
        End If
        '
        If MODORECI% = 1 Then
             FF% = CVI(Mid$(X$, 19, 2)) ' orden emision
           Else
             FF% = CVI(Mid$(X$, 9, 2)) ' orden vencimiento
        End If
        Z$ = CSTRING$(MKI$(FF%), 1, 6, 0, 2) + CSTRING$(MKS$(REG&), 4, 8, 0, 2)
        RESCUDEU.LINDICE.AddItem Z$
    Next U&
    '
    SUMADEU# = 0: SUMAFAC# = 0
    For U& = 1 To RESCUDEU.LINDICE.ListCount
        RESCUDEU.LINDICE.ListIndex = U& - 1
        REG& = Val(Mid$(RESCUDEU.LINDICE.TEXT, 7, 8))
        If REG& > 0 Then
            If REG& <= MAXDEU& Then
                X$ = REGLEIDO$("SVD202", REG&)
                TI% = Asc(Mid$(X$, 3, 1))
                VA% = Asc(Mid$(X$, 4, 1))
                TIPOS$ = "Factura": If TI% = 0 Then TIPOS$ = "Ajuste Saldo"
                                   If TI% = 1 Then TIPOS$ = "Venta Contado"
                                   If TI% = 3 Then TIPOS$ = "Nota Debito"
                                   If TI% = 4 Then TIPOS$ = "Anticipo"
                                   If TI% = 5 Then TIPOS$ = "Desc. Caja"
                                   If TI% = 6 Then TIPOS$ = "Nota Credito"
                                   If TI% = 7 Then TIPOS$ = "Ajuste Saldo"
                                   '
                SG% = 1: If TI% >= 4 Then If TI% <= 6 Then SG% = -1
                IMPO# = CVD(Mid$(X$, 21, 8))
                If TI% <> 0 Then IMPO# = Abs(IMPO#)
                IMPO# = SG% * IMPO#
                '
                SALDO# = CVD(Mid$(X$, 11, 8))
                If TI% <> 0 Then SALDO# = Abs(SALDO#)
                SALDO# = SG% * SALDO#
                '
                If SALDO# >= 0.005 Then
                    SUMAFAC# = SUMAFAC# + IMPO#
                    SUMADEU# = SUMADEU# + SALDO#
                    NROL& = CVS(Mid$(X$, 5, 4))
                    FEMI = CVI(Mid$(X$, 19, 2))
                    FEVE = CVI(Mid$(X$, 9, 2))
                    '
                    IMPTEX$ = CSTRING$(MKD$(IMPO#), 3, 14, 2, 2)
                    SALTEX$ = CSTRING$(MKD$(SALDO#), 3, 14, 2, 2)
                    DOCU$ = FORMDOC$(NROL&, TI%, VA%, 1)
                    '
                    Z$ = Space$(128)
                    Call REPLA(Z$, TIPOS$, 1, 14)
                    Call REPLA(Z$, CSTRING$(MKS$(NROL&), 3, 8, 0, 2), 13, 8)
                    Call REPLA(Z$, FECHATEX$(FEMI), 23, 8)
                    Call REPLA(Z$, IMPTEX$, 31, 14)
                    Call REPLA(Z$, FECHATEX$(FEVE), 48, 8)
                    Call REPLA(Z$, SALTEX$, 56, 14)
                    Call REPLA(Z$, DOCU$, 97, 18)
                    Call REPLA(Z$, TRIM$(Str$(REG&)), 121, 6)
                    '
                    FORECISTA.List2.AddItem Z$
                End If
                '
            End If
        End If
        '
    Next U&
    '
    FORECISTA.List2.ListIndex = -1
    FORECISTA.List2.Refresh
    'FORECISTA.Label14.Caption = CSTRING$(MKD$(SUMAFAC#), 4, 14, 2, 2)
    'FORECISTA.Label14.Refresh
    FORECISTA.Label8.Caption = CSTRING$(MKD$(SUMADEU#), 4, 14, 2, 2)
    FORECISTA.Label8.Refresh
    FORECISTA.Label10.Caption = ""
    FORECISTA.Label10.Refresh
    FORECISTA.List2.Visible = True
    FORECISTA.List2.SetFocus
    If FORECISTA.List2.ListCount > 0 Then
        VACON% = 1
        FORECISTA.List2.ListIndex = 0
        VACON% = 0
    End If
    '
    FORECISTA.BOTNEXT.Enabled = False
    Call BLANARCH("!CARMECO")
    Screen.MousePointer = 1
    '
    Call SIGUERECI2
    '
End Sub
'

Sub SIGUERECI2()
    '
    HOI% = HOY(HO$)
    '
200 ATRI$ = "/NC"
    ATRI$ = ATRI$ + "/TITUPAN=Ingreso de Medios de Pago Codificados"
    ATRI$ = ATRI$ + "/BORDESUP=W" + TRIM$(Str$(FORECISTA.Top + FORECISTA.Picture1.Top + 100))
    ATRI$ = ATRI$ + "/BORDEIZQ=W" + TRIM$(Str$(FORECISTA.Left + FORECISTA.Picture1.Left))
    '
    VTB% = VENTABU%("CARMECOW" + ATRI$)
    Call CIERRARCH("!CARMECO")
    If VTB% < 0 Then
        FORECISTA.BOTNEXT.Enabled = True
        Exit Sub
    End If
    '
    
    FORECISTA.List1.Clear
    SUMAVA# = 0
    FORECISTA.Label14 = ""
    For U& = 1 To ULTREG&("!CARMECO")
        X$ = REGLEIDO$("!CARMECO", U&)
        If X$ = REBLA$ Then GoTo 290
        MECO% = CVI(Left$(X$, 2))
        BANCO$ = TRIM$(ECOARCH$("MECOBRA", MKI$(MECO%)))
        '
        If BANCO$ = "" Then
            Call MENSERR(24, "Medio de Pago no Válido")
            GoTo 200
        End If
        '
        NUVA$ = Mid$(X$, 19, 12)
        IMVA# = CVD(Mid$(X$, 31, 8))
        If IMVA# < 0.005 Then
            Call MENSERR(24, "Falta Importe")
            GoTo 200
        End If
        '
        FEMI = CVI(Mid$(X$, 39, 2))
        F1 = FEMI: F2 = HOI%
        If FEMI < HOI% - 32 * 13 Or FEMI > HOI% + 30 Or DIENTRE%(F1, F2) > 365 Then
            Call MENSERR(24, "Fecha de Emisión no Válida")
            GoTo 200
        End If
        '
        FEVE = CVI(Mid$(X$, 41, 2))
        F1 = FEVE: F2 = HOI%
        If FEVE < HOI% - 64 Or DIENTRE%(F1, F2) > 30 Then
            Call MENSERR(24, "Fecha de Vencimiento no Válida")
            GoTo 200
        End If
        '
        PROTER$ = UCase$(TRIM$(Mid$(X$, 43, 2)))
        If PROTER$ = "" Or Left$(PROTER$, 1) = "P" Then
             PROTER$ = "  "
           Else
             PROTER$ = "3 "
        End If
        Call REPLA(X$, PROTER$, 43, 2)
        Call GRAREG("!CARMECO", X$, U&)
        '
        CLEA% = CVI(Mid$(X$, 45, 2))
        If CLEA% < 24 Then CLEA% = 24
        If (CLEA% Mod 24) <> 0 Then
            Call MENSERR(24, "Clearing no Válido")
            GoTo 200
        End If
        Call REPLA(X$, MKI$(CLEA%), 45, 2)
        Call GRAREG("!CARMECO", X$, U&)
        '
        Z$ = Space$(96)
        SUMAVA# = SUMAVA# + IMVA#
        IMPTEX$ = CSTRING$(MKD$(IMVA#), 3, 14, 2, 2)
        Call REPLA(Z$, CSTRING$(MKI$(MECO%), 1, 4, 0, 2), 1, 4)
        Call REPLA(Z$, BANCO$, 7, 14)
        Call REPLA(Z$, NUVA$, 23, 12)
        Call REPLA(Z$, PROTER$, 38, 2)
        Call REPLA(Z$, CSTRING$(MKI$(CLEA%), 1, 4, 0, 2), 42, 4)
        Call REPLA(Z$, FECHATEX$(FEMI), 50, 8)
        Call REPLA(Z$, FECHATEX$(FEVE), 62, 8)
        Call REPLA(Z$, IMPTEX$, 70, 14)
        '
        FORECISTA.List1.AddItem Z$
        '
290 Next U&
    '
    If FORECISTA.List1.ListCount > 0 Then
        FORECISTA.List1.ListIndex = 0
    End If
    '
    FORECISTA.Label14.Caption = TRIM$(CSTRING$(MKD$(SUMAVA#), 4, 14, 2, 2))
    FORECISTA.Label14.Refresh
    ANTI# = SUMAVA# - CDbl(Val(FORECISTA.Label18.Caption))
    If ANTI# < 0 Then
          FORECISTA.Label19.Caption = "Descuentos:"
        Else
          FORECISTA.Label19.Caption = "Anticipos:"
    End If
    FORECISTA.Label20.Caption = TRIM$(CSTRING$(MKD$(ANTI#), 4, 14, 2, 2))
    FORECISTA.Label20.Refresh
    '
End Sub
'

Sub SIGUERECI3()
    '
    If FORECISTA.List2.ListCount < 1 Then
        Exit Sub
    End If
    '
    Call BLANARCH("!APLIRECI")
    For U& = 1 To FORECISTA.List2.ListCount
        FORECISTA.List2.ListIndex = U& - 1
        Z$ = FORECISTA.List2.TEXT
        TIPOS$ = Left$(Z$, 12)
        NROL& = Val(Mid$(Z$, 13, 8))
        FEMII% = FECHANUM(Mid$(Z$, 23, 8))
        IPED# = CDbl(Val(Mid$(Z$, 31, 14)))
        FEVEI% = FECHANUM(Mid$(Z$, 48, 8))
        SALD# = CDbl(Val(Mid$(Z$, 56, 14)))
        ACRE# = CDbl(Val(Mid$(Z$, 70, 14)))
        DOCU$ = Mid$(Z$, 97, 18)
        REG2& = Val(Mid$(Z$, 121, 6))
        '
        Y$ = REGBLAN$("!APLIRECI")
        Call REPLA(Y$, TIPOS$, 1, 16)
        Call REPLA(Y$, MKS$(NROL&), 17, 4)
        Call REPLA(Y$, MKI$(FEMII%), 21, 2)
        Call REPLA(Y$, MKD$(IPED#), 23, 8)
        Call REPLA(Y$, MKI$(FEVEI%), 31, 2)
        Call REPLA(Y$, MKD$(SALD#), 33, 8)
        Call REPLA(Y$, MKD$(ACRE#), 41, 8)
        Call REPLA(Y$, DOCU$, 49, 18)
        Call REPLA(Y$, MKS$(REG2&), 67, 4)
        '
        Call REGAPP("!APLIRECI", Y$)
    Next U&
    Call CIERRARCH("!APLIRECI")
    '
200 ATRI$ = "/NC"
    ATRI$ = ATRI$ + "/TITUPAN=Aplicación de Cobranza"
    ATRI$ = ATRI$ + "/BORDESUP=W" + TRIM$(Str$(FORECISTA.Top + FORECISTA.Picture2.Top))
    ATRI$ = ATRI$ + "/BORDEIZQ=W" + TRIM$(Str$(FORECISTA.Left + FORECISTA.Picture2.Left))
    '
    VTB% = VENTABU%("APLIRECI" + ATRI$)
    '
    VACON% = 1
    If VTB% >= 0 Then
        FORECISTA.List2.Clear
        For U& = 1 To ULTREG&("!APLIRECI")
            Y$ = REGLEIDO$("!APLIRECI", U&)
            TIPOS$ = Left$(Y$, 16)
            NROL& = CVS(Mid$(Y$, 17, 4))
            FEMI = CVI(Mid$(Y$, 21, 2))
            IPED# = CVD(Mid$(Y$, 23, 8))
            IMPTEX$ = CSTRING$(MKD$(IPED#), 4, 14, 2, 2)
            FEVE = CVI(Mid$(Y$, 31, 8))
            SALD# = CVD(Mid$(Y$, 33, 8))
            SALTEX$ = CSTRING$(MKD$(SALD#), 4, 14, 2, 2)
            ACRE# = CVD(Mid$(Y$, 41, 8))
            '
            If ACRE# < 0 Then
                Call MENSERR(24, "No se Permite Aplicación Negativa")
                GoTo 200
            End If
            If ACRE# > SALD# Then
                Call MENSERR(24, "El Importe Acreditado Debe Ser\Inferior al Saldo del Documento")
                GoTo 200
            End If
            ACRTEX$ = CSTRING$(MKD$(ACRE#), 4, 14, 2, 2)
            DOCU$ = Mid$(Y$, 49, 18)
            REG2& = CVS(Mid$(Y$, 67, 4))
            '
            Z$ = Space$(128)
            Call REPLA(Z$, TIPOS$, 1, 14)
            Call REPLA(Z$, CSTRING$(MKS$(NROL&), 3, 8, 0, 2), 13, 8)
            Call REPLA(Z$, FECHATEX$(FEMI), 23, 8)
            Call REPLA(Z$, IMPTEX$, 31, 14)
            Call REPLA(Z$, FECHATEX$(FEVE), 48, 8)
            Call REPLA(Z$, SALTEX$, 56, 14)
            Call REPLA(Z$, ACRTEX$, 70, 14)
            Call REPLA(Z$, DOCU$, 97, 18)
            Call REPLA(Z$, TRIM$(Str$(REG2&)), 121, 6)
            '
            FORECISTA.List2.AddItem Z$
        Next U&
    End If
    '
    SUMACRE# = 0
    For U& = 1 To FORECISTA.List2.ListCount
        FORECISTA.List2.ListIndex = U& - 1
        Z$ = FORECISTA.List2.TEXT
        SUMACRE# = SUMACRE# + CDbl(Val(Mid$(Z$, 70, 14)))
    Next U&
    '
    FORECISTA.Label10.Caption = CSTRING$(MKD$(SUMACRE#), 4, 14, 2, 2)
    FORECISTA.Label10.Refresh
    FORECISTA.Label18.Caption = TRIM$(CSTRING$(MKD$(SUMACRE#), 4, 14, 2, 2))
    FORECISTA.Label18.Refresh
    ANTI# = CDbl(Val(FORECISTA.Label14.Caption)) - SUMACRE#
    If ANTI# < 0 Then
          FORECISTA.Label19.Caption = "Descuentos:"
        Else
          FORECISTA.Label19.Caption = "Anticipos:"
    End If
    FORECISTA.Label20.Caption = TRIM$(CSTRING$(MKD$(ANTI#), 4, 14, 2, 2))
    FORECISTA.Label20.Refresh
    '
    FORECISTA.List2.Refresh
    VACON% = 1
    '
End Sub
'

Sub GRARECISTA(OKEY%)
    '
700 OKEY% = 0
    '
    NC% = Val(FORECISTA.Text1.TEXT)
    NROO& = Val(Mid$(TRIM$(FORECISTA.Text11.TEXT), 4))
    FC% = FECHANUM(TRIM$(FORECISTA.Text9.TEXT))
    TOTAVAL# = CDbl(Val(FORECISTA.Label14.Caption))
    ANTI# = CDbl(Val(FORECISTA.Label20.Caption))
    URE202& = ULTREG&("SVD202")
    '
    Call BLOQARCH("COBRAN")
    Call BLOQARCH("CHECAR")
    Call BLOQARCH("APLICOB")
    '
    Screen.MousePointer = 11   ' Grabando Recibo - Espere un Momento, por Favor"
    '
    URE202& = ULTREG&("SVD202")
    '
    If ANTI# >= 0.005 Then
        '
        Z$ = REGBLAN$("SVD202")
        Call REPLA(Z$, MKI$(NC%), 1, 2)
        Call REPLA(Z$, Chr$(4), 3, 1)
        Call REPLA(Z$, Chr$(0), 4, 1)
        Call REPLA(Z$, MKS$(NROO&), 5, 4)
        Call REPLA(Z$, MKI$(FC%), 9, 2)
        Call REPLA(Z$, MKD$(ANTI#), 11, 8)
        Call REPLA(Z$, MKI$(FC%), 19, 2)
        Call REPLA(Z$, MKD$(TOTAVAL#), 21, 8)
        Call REPLA(Z$, String$(2, 255), 31, 2)
        '
        Call BLOQARCH("SVD202")
        Call REGAPP("SVD202", Z$)
        Call CIERRARCH("SVD202")
        '
    End If
    OKEY% = 1
    '
800 ' *********************************** aqui blanquear facturas canceladas
    '
    XX% = 0
    VACON% = 1
    FORECISTA.List2.Visible = False
    ILISTA& = FORECISTA.List2.ListIndex
    TOLISTA& = FORECISTA.List2.TopIndex
    '
    For U& = 1 To FORECISTA.List2.ListCount
        'FORECISTA.List2.ListIndex = U& - 1
        X$ = FORECISTA.List2.List(U& - 1)
        CANCE# = CDbl(Val(Mid$(X$, 70, 14)))
        If CANCE# >= 0.005 Then
            NREG& = Val(Mid$(X$, 121, 6))
            If NREG& > 0 Then
                If NREG& <= ULTREG&("SVD202") Then
                    '
                    SS$ = REGLEIDO$("SVD202", NREG&)
                    SALDEU# = CVD(Mid$(SS$, 11, 8)) - CANCE#
                    Call REPLA(SS$, MKD$(SALDEU#), 11, 8)
                    Call GRAREG("SVD202", SS$, NREG&)
                    '
                End If
            End If
        End If
    Next U&
    Call CIERRARCH("SVD202")
    OKEY% = 2
    '
750 RGX& = REGICLI&(NC%)
    If RGX& > 0 Then
        If RGX& <= ULTREG&("DEUDOR2") Then
           DX$ = REGLEIDO$("DEUDOR2", RGX&)
           SALDO# = CVD(Mid$(DX$, 109, 8)) - TOTAVAL#
           Call REPLA(DX$, MKD$(SALDO#), 109, 8)
           Call GRAREG("DEUDOR2", DX$, RGX&)
        End If
    End If
    Call CIERRARCH("DEUDOR2")
    OKEY% = 3
    '
    Z$ = REGBLAN$("COBRAN")
    Call REPLA(Z$, MKI$(NC%), 1, 2)
    Call REPLA(Z$, MKS$(NROO&), 3, 4)
    Call REPLA(Z$, MKI$(FC%), 7, 2)
    Call REPLA(Z$, String$(6, 0), 9, 6)
    Call REPLA(Z$, MKD$(TOTAVAL#), 15, 8)
    '
    Call REGAPP("COBRAN", Z$)
    Call CIERRARCH("COBRAN")
    OKEY% = 4
    '
    DESCAJ# = 0
    ANTI# = CDbl(Val(FORECISTA.Label20.Caption))
    If ANTI# < 0 Then
        DESCAJ# = Abs(ANTI#)
        ANTI# = 0
    End If
    '
    If DESCAJ# >= 0.005 Then
'aun no se permiten descuentos de caja
    TIPO% = 5
     ' grabar como tipo 5 en 'FACTUR.DAT'
    End If
    '
    OKEY% = 5

830 ' *************************************** aqui grabar archivo APLICOB.DAT
    '
    For U& = 1 To FORECISTA.List2.ListCount
        'FORECISTA.List2.ListIndex = U& - 1
        X$ = FORECISTA.List2.List(U& - 1)
        CANCE# = CDbl(Val(Mid$(X$, 70, 14)))
        If CANCE# >= 0.005 Then
            NREG& = Val(Mid$(X$, 121, 6))
            If NREG& > 0 Then
                If NREG& <= URE202& Then
                    SS$ = REGLEIDO$("SVD202", NREG&)
                    DOCORI$ = FORMDOC$(CVS(Mid$(SS$, 5, 4)), Asc(Mid$(SS$, 3, 1)), Asc(Mid$(SS$, 4, 1)), 1)
                    ITOTORI# = CVD(Mid$(SS$, 21, 8))
                    IVAORI# = (Int(ALIVA * ITOTORI# / (1 + ALIVA / 100) + 0.5)) / 100
                    '
                    REFAPLI$ = "Recibo " + TRIM$(Str$(NROO&))
                    Call APLIPAG(FC%, NC%, 4, 1, SUC%, NROO&, CANCE#, REFAPLI$, DOCORI$, ITOTORI#, IVAORI#)
                    '
                End If
            End If
        End If
    Next U&
    VACON% = 0
    Call CIERRARCH("APLICOB")
    OKEY% = 6
    '
840 ' ************************************** aqui grabar archivo CHECAR.DAT
    '
    For U& = 1 To ULTREG&("!CARMECO")
        X$ = REGLEIDO$("!CARMECO", U&)
        TIPO% = CVI(Left$(X$, 2))
        BANCO$ = Mid$(X$, 3, 16)
        NUCHE$ = Mid$(X$, 19, 12)
        IMCHE# = CVD(Mid$(X$, 31, 8))
        FECHE% = CVI(Mid$(X$, 39, 2)): If FECHE% < 1 Then FECHE% = HO%
        FEVEN% = CVI(Mid$(X$, 41, 2)): If FEVEN% < FECHE% Then FEVEN% = FECHE%
        '
        If IMCHE# > 0 Then
            '
            Z$ = REGBLAN$("CHECAR")
            Call REPLA(Z$, MKI$(TIPO%), 1, 2)
            Call REPLA(Z$, MKI$(NC%), 3, 2)
            Call REPLA(Z$, NUCHE$, 5, 12)
            Call REPLA(Z$, BANCO$, 17, 16)
            Call REPLA(Z$, MKD$(IMCHE#), 33, 8)
            Call REPLA(Z$, MKI$(FEVEN%), 41, 2)
            Call REPLA(Z$, MKI$(FECHE%), 43, 2)
            Call REPLA(Z$, MKI$(FC%), 45, 2)
            Call REPLA(Z$, MKS$(NROO&), 47, 4)
            Call REPLA(Z$, String$(14, 0), 51, 14)
            '
            Call REGAPP("CHECAR", Z$)
            '
        End If
    Next U&
    Call CIERRARCH("CHECAR")
    OKEY% = 6
    '
890 Screen.MousePointer = 1
    Call CIERRARCH("*.*")
    '
End Sub
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
    FORIPRE$ = TRIM$(CONTROL$("", "FORECIST"))
    If FORIPRE$ = "" Then
        Call COMUNI("Error de Configuración:\  \Falta Definir Formulario\de Impresión de Recibo.")
        Exit Sub
    End If
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
        FEX = FEVEI%
        If FEX > 0 Then
            Call REPLA(Z$, FECHATEX$(FEX), 41, 8)
        End If
        Call REGAPP("!PRINRECI", Z$)
        '
19  Next U&
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
            Call REPLA(Z$, DOCU$, 53, 18)
            Call REPLA(Z$, MKD$(CANCE#), 71, 8)
            Call GRAREG("!PRINRECI", Z$, V&)
        End If
    Next U&
    VACON% = 0
    '
    V0& = ULTREG&("!PRINRECI")
    '
    TOTAVAL# = CDbl(Val(FORECISTA.Label14.Caption))
    TOTACAN# = CDbl(Val(FORECISTA.Label18.Caption)) + CDbl(Val(FORECISTA.Label20.Caption))
    ITEXTO$ = NUMTEX$(TOTAVAL#)
    LOTEX% = Val(ATRIFORM$("", "LOIMPLET")): If LOTEX% < 20 Then LOTEX% = 20
    Call FRATEX(ITEXTO$, NRX1$, NRX2$, NRX3$, LOTEX%)
    IMPLET1$ = TRIM$(NRX1$) + " " + "***************************************************************************************************************"
    IMPLET2$ = TRIM$(NRX2$) + " " + TRIM$(NRX3$) + " ctvs. *****************************************************************************************"
    '
    ANTI# = CDbl(Val(FORECISTA.Label20.Caption))
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
        If ANTI# < 0 Then
            Call REPLA(Z$, "Descuento de Caja", 53, 18)
        End If
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
            Printer.Orientation = 1
            If ORIENTA% = 2 Then
                Printer.Orientation = 2
            End If
        End If
        '
        For COP% = 1 To CACOPI%
            '
            If AVAHOJA$ = "NEWPAGE" Then
               If PRINTERCONTROL% = 1 Then
                  Printer.NewPage
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
                        Printer.NewPage
                      Else
                       Printer.EndDoc
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
                Call PRINTFORWIN(Mid$(YY$, 53, 18), "DOCU-APL", 0, YAC)
                Call PRINTFORWIN(CSTRING$(Mid$(YY$, 71, 8), 4, 12, 2, 2), "IMPO-CAN", 0, YAC)
                SUMACAN# = SUMACAN# + CVD(Mid$(YY$, 71, 8))
                '
            Next YX&
            '
            Call PRINTFORWIN(CSTRING$(MKD$(TOTAVAL#), 4, 12, 2, 2), "TOTA-VAL", 0, 0)
            Call PRINTFORWIN(CSTRING$(MKD$(TOTACAN#), 4, 12, 2, 2), "TOTA-CAN", 0, 0)
            Call PRINTFORWIN(IMPLET1$, "IMP-LET1", 0, 0)
            Call PRINTFORWIN(IMPLET2$, "IMP-LET2", 0, 0)
                   '
            If AVAHOJA$ <> "NEWPAGE" Then
                 Printer.EndDoc
            End If
            '
        Next COP%
        '
        Call CIERRARCH("*.*")
        Screen.MousePointer = 1
        If COMALTER%("Recibo Numero\'" + NRORECI$ + "'\Emitido y Registrado\\Repetir Impresion\Otro Recibo") = 1 Then GoTo 30
        '
        If AVAHOJA$ <> "NEWPAGE" Then
            Printer.Orientation = PORIAN
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
        If CVS(Mid$(Z$, 47, 4)) <> NROLL& Then
            J& = J& + 1
            Call GRAREG("CHECAR", Z$, J&)
        End If
    Next U&
    Call SETULTREG("CHECAR", J&)
    Call CIERRARCH("CHECAR")
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
    PROLLAMA$ = "RECICOMP"
    RESCUDEU.Show
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
    IMPO# = Abs(CVD(Mid$(XX0$, 15, 8)))
    IMPOS$ = Left$("$ " + TRIM$(NUMSTRI$(IMPO#, 14, 2, 1)) + Space$(18), 18)
    NROLL& = CVS(Mid$(XX0$, 3, 4)): NNXX# = CDbl(NROLL&)
    X$ = "Recibo Nro" + NUMSTRI$(NNXX#, 8, 0, 1)
    Y$ = IMPOS$
    FEX = CVI(Mid$(XX0$, 7, 2))
    '
    If COMALTER%("Documento: " + X$ + "\  Importe: " + Y$ + "\\Repetir Impresion\Cancelar") <> 1 Then GoTo 999
    FORECISTA.Text9.TEXT = FECHATEX$(FEX)
    FORECISTA.Text11.TEXT = "RB." + TRIM$(Str$(NROLL&))
    '

    ' carga detalle de recibo
    ' ***********************
    '
311 Screen.MousePointer = 11
    Call BLANARCH("!CARMECO")
    '
    J& = 0
    TOTAVAL# = 0
    For U& = 1 To ULTREG&("CHECAR")
        X$ = REGLEIDO$("CHECAR", U&)
        If CVS(Mid$(X$, 47, 4)) = NROLL& Then
            '
            TIPO% = CVI(Left$(X$, 2))
            BANCO$ = Mid$(X$, 17, 16)
            NUCHE$ = Mid$(X$, 5, 12)
            IMCHE# = CVD(Mid$(X$, 33, 8))
            FECHE% = CVI(Mid$(X$, 43, 2))
            FEVEN% = CVI(Mid$(X$, 41, 2))
            '
            Z$ = REGBLAN$("!CARMECO")
            Call REPLA(Z$, MKI$(TIPO%), 1, 2)
            Call REPLA(Z$, BANCO$, 3, 16)
            Call REPLA(Z$, NUCHE$, 19, 12)
            Call REPLA(Z$, MKD$(IMCHE#), 31, 8)
            Call REPLA(Z$, MKI$(FECHE%), 39, 2)
            Call REPLA(Z$, MKI$(FEVEN%), 41, 2)
            Call REGAPP("!CARMECO", Z$)
            '
            TOTAVAL# = TOTAVAL# + IMCHE#
            '
        End If
    Next U&
    Call CIERRARCH("CHECAR")
    '
    FORECISTA.List2.Clear
    TOTACAN# = 0
    For U& = 1 To ULTREG&("APLICOB")
        Z$ = REGLEIDO$("APLICOB", U&)
        If Mid$(Z$, 5, 2) = "RB" Then
            If Val(Mid$(Z$, 15, 8)) = NROLL& Then
                FEMI = CVI(Left$(Z$, 2))
                NCA% = CVI(Mid$(Z$, 3, 2))
                SALTEX$ = CSTRING$(Mid$(Z$, 23, 8), 4, 14, 2, 2)
                TOTACAN# = TOTACAN# + CDbl(Val(SALTEX$))
                DOCU$ = Mid$(Z$, 55, 18)
                IMPTEX$ = CSTRING$(Mid$(Z$, 73, 8), 4, 14, 2, 2)
                '
                ZZ$ = Space$(128)
                Call REPLA(ZZ$, DOCU$, 1, 18)
                Call REPLA(ZZ$, FECHATEX$(FEMI), 23, 8)
                Call REPLA(ZZ$, IMPTEX$, 31, 14)
                Call REPLA(ZZ$, FECHATEX$(FEMI), 48, 8)
                Call REPLA(ZZ$, SALTEX$, 56, 14)
                Call REPLA(ZZ$, SALTEX$, 70, 14)
                Call REPLA(ZZ$, DOCU$, 97, 18)
                '
                FORECISTA.List2.AddItem ZZ$
            End If
        End If
    Next U&
    '
    FORECISTA.Label14.Caption = CSTRING$(MKD$(TOTAVAL#), 4, 12, 2, 2)
    FORECISTA.Label18.Caption = CSTRING$(MKD$(TOTACAN#), 4, 12, 2, 2)
    FORECISTA.Label20.Caption = CSTRING$(MKD$(TOTAVAL# - TOTACAN#), 4, 12, 2, 2)
    '
    Call CIERRARCH("*.*")
    Call PRIRECISTA
    '
999 Call CIERRARCH("*.*")
    Call LIBARCH("!SELCUERE")
    '
    Call CIERRARCH("!CARMECO")
    Call BLANARCH("!CARMECO")
    Call CIERRARCH("!APLIRECI")
    Call BLANARCH("!APLIRECI")
    FORECISTA.List1.Clear
    FORECISTA.List2.Clear
    FORECISTA.Label4.Caption = ""
    FORECISTA.Label8.Caption = ""
    FORECISTA.Label10.Caption = ""
    FORECISTA.Label14.Caption = ""
    FORECISTA.Label18.Caption = ""
    FORECISTA.Label20.Caption = ""
    FORECISTA.BOTNEXT.Enabled = True
    FORECISTA.Command5.Enabled = False
    ' FORECISTA.Command12.Enabled = False
    '
    Screen.MousePointer = 1
    '
End Sub


Sub IMPUCRE1()
    '
    FORECISTA.Hide
    FIMPUCRE.Show
    '
    ' 1.- Ingresar Numero de Cliente
    '     **************************
    '
10  FIMPUCRE.List2.Clear
    FIMPUCRE.Label1.Caption = ""
    FIMPUCRE.Label2.Caption = ""
    FIMPUCRE.Label3.Caption = ""
    FIMPUCRE.Label6.Caption = ""
    FIMPUCRE.Label12.Caption = ""
    FIMPUCRE.Label15.Caption = ""
    FIMPUCRE.Label22.Caption = ""
    FIMPUCRE.Label23.Caption = ""
    FIMPUCRE.Label24.Caption = ""
    FIMPUCRE.Label25.Caption = ""
    FIMPUCRE.Label8.Caption = ""
    FIMPUCRE.Label10.Caption = ""
    FIMPUCRE.Label14.Caption = ""
    FIMPUCRE.Label18.Caption = ""
    FIMPUCRE.Label20.Caption = ""
    FIMPUCRE.Command5.Enabled = False
    FIMPUCRE.Refresh
    Call FRESHECHO("!SELDOCAP")
    '
    If MODORECI% = 0 Then
        MODORECI% = 1       ' por emision
        If CONTROL$("RECISTA", "MODORDE") = "VENCIM" Then
           MODORECI% = 2
        End If
    End If
    '
    Call SELECHO("DEUDOR1")
    NC% = Val(VALACT1$("DEUDOR1"))
    If NC% < 1 Then
        GoTo 999
    End If
    '
    FIMPUCRE.Label1.Caption = Str$(NC%)
    FIMPUCRE.Label2.Caption = RASOCLI$(NC%)
    FIMPUCRE.Label3.Caption = DOMICLI$(NC%)
    FIMPUCRE.Label6.Caption = CPOSCLI$(NC%)
    FIMPUCRE.Label12.Caption = LOCACLI$(NC%)
    FIMPUCRE.Refresh
    '
    Screen.MousePointer = 11
    '
100 FIMPUCRE.List2.Clear
    FIMPUCRE.List2.Refresh
    Call BLANARCH("!SELDOCAP")
    '
    MAXDEU& = ULTREG&("SVD202")
    RFD$ = RECFILT$("SVD202", 1, MKI$(NC%))
    '
    RESCUDEU.LINDICE.Clear
    For U& = 1 To Len(RFD$) Step 4
        REG& = CVS(Mid$(RFD$, U&, 4))
        X$ = REGLEIDO$("SVD202", REG&)
        If CVI(Left$(X$, 2)) <> NC% Then
            X$ = REGLEIDO$("SVD202", 1)
            Call GRAREG("SVD202", X$, 1)
            Call CIERRARCH("SVD202")
            GoTo 100
        End If
        IMPO# = CVD(Mid$(X$, 21, 8))
        SALDO# = CVD(Mid$(X$, 11, 8))
        If Sgn(IMPO#) <> Sgn(SALDO#) Or Abs(SALDO#) > Abs(IMPO#) Then
          If Abs(SALDO#) >= 0.005 Then
            ABUS$ = Left$(X$, 8)
            For UU& = 1 To Len(RFD$) Step 4
                REG& = CVS(Mid$(RFD$, UU&, 4))
                X$ = REGLEIDO$("SVD202", REG&)
                If Left$(X$, 8) = ABUS$ Then
                    Call GRAREG("SVD202", String$(32, 0), REG&)
                End If
            Next UU&
            GoTo 100
          End If
        End If
        '
        If MODORECI% = 1 Then
             FF% = CVI(Mid$(X$, 19, 2)) ' orden emision
           Else
             FF% = CVI(Mid$(X$, 9, 2)) ' orden vencimiento
        End If
        Z$ = CSTRING$(MKI$(FF%), 1, 6, 0, 2) + CSTRING$(MKS$(REG&), 4, 8, 0, 2)
        RESCUDEU.LINDICE.AddItem Z$
    Next U&
    '
    SUMADEU# = 0: SUMAFAC# = 0: KQ& = 0
    For U& = 1 To RESCUDEU.LINDICE.ListCount
        RESCUDEU.LINDICE.ListIndex = U& - 1
        REG& = Val(Mid$(RESCUDEU.LINDICE.TEXT, 7, 8))
        If REG& > 0 Then
            If REG& <= MAXDEU& Then
                X$ = REGLEIDO$("SVD202", REG&)
                TI% = Asc(Mid$(X$, 3, 1))
                VA% = Asc(Mid$(X$, 4, 1))
                TIPOS$ = "Factura": If TI% = 0 Then TIPOS$ = "Ajuste Saldo"
                                   If TI% = 1 Then TIPOS$ = "Venta Contado"
                                   If TI% = 3 Then TIPOS$ = "Nota Debito"
                                   If TI% = 4 Then TIPOS$ = "Anticipo"
                                   If TI% = 5 Then TIPOS$ = "Desc. Caja"
                                   If TI% = 6 Then TIPOS$ = "Nota Credito"
                                   If TI% = 7 Then TIPOS$ = "Ajuste Saldo"
                                   '
                SG% = 1: If TI% >= 4 Then If TI% <= 6 Then SG% = -1
                IMPO# = CVD(Mid$(X$, 21, 8))
                If TI% <> 0 Then IMPO# = Abs(IMPO#)
                IMPO# = SG% * IMPO#
                '
                SALDO# = CVD(Mid$(X$, 11, 8))
                If TI% <> 0 Then SALDO# = Abs(SALDO#)
                SALDO# = SG% * SALDO#
                '
                If SALDO# >= 0.005 Then
                    SUMAFAC# = SUMAFAC# + IMPO#
                    SUMADEU# = SUMADEU# + SALDO#
                    NROL& = CVS(Mid$(X$, 5, 4))
                    FEMI = CVI(Mid$(X$, 19, 2))
                    FEVE = CVI(Mid$(X$, 9, 2))
                    '
                    IMPTEX$ = CSTRING$(MKD$(IMPO#), 3, 14, 2, 2)
                    SALTEX$ = CSTRING$(MKD$(SALDO#), 3, 14, 2, 2)
                    DOCU$ = FORMDOC$(NROL&, TI%, VA%, 1)
                    '
                    Z$ = Space$(128)
                    Call REPLA(Z$, TIPOS$, 1, 14)
                    Call REPLA(Z$, CSTRING$(MKS$(NROL&), 3, 8, 0, 2), 13, 8)
                    Call REPLA(Z$, FECHATEX$(FEMI), 23, 8)
                    Call REPLA(Z$, IMPTEX$, 31, 14)
                    Call REPLA(Z$, FECHATEX$(FEVE), 48, 8)
                    Call REPLA(Z$, SALTEX$, 56, 14)
                    Call REPLA(Z$, DOCU$, 97, 18)
                    Call REPLA(Z$, TRIM$(Str$(REG&)), 121, 6)
                    '
                    FIMPUCRE.List2.AddItem Z$
                    '
                  ElseIf SALDO# <= (-0.005) Then
                    '
                    NROL& = CVS(Mid$(X$, 5, 4))
                    FEMI = CVI(Mid$(X$, 19, 2))
                    FEVE = CVI(Mid$(X$, 9, 2))
                    '
                    IMPTEX$ = CSTRING$(MKD$(Abs(IMPO#)), 3, 14, 2, 2)
                    SALTEX$ = CSTRING$(MKD$(Abs(SALDO#)), 3, 14, 2, 2)
                    DOCU$ = FORMDOC$(NROL&, TI%, VA%, 1)
                    '
                    Z$ = Space$(60) + MKS$(0)
                    Call REPLA(Z$, DOCU$, 1, 18)
                    Call REPLA(Z$, FECHATEX$(FEMI), 19, 8)
                    Call REPLA(Z$, IMPTEX$, 27, 14)
                    Call REPLA(Z$, SALTEX$, 41, 14)
                    Call REPLA(Z$, MKS$(REG&), 61, 4)
                    '
                    Call REGAPP("!SELDOCAP", Z$)
                    KQ& = KQ& + 1
                    
                End If
                '
            End If
        End If
        '
    Next U&
    Call CIERRARCH("!SELDOCAP")
    '
    FIMPUCRE.List2.ListIndex = -1
    FIMPUCRE.List2.Refresh
    FIMPUCRE.Label8.Caption = CSTRING$(MKD$(SUMADEU#), 4, 14, 2, 2)
    FIMPUCRE.Label8.Refresh
    FIMPUCRE.Label10.Caption = ""
    FIMPUCRE.Label10.Refresh
    FIMPUCRE.List2.Visible = True
    FIMPUCRE.List2.SetFocus
    If FIMPUCRE.List2.ListCount > 0 Then
        VACON% = 1
        FIMPUCRE.List2.ListIndex = 0
        VACON% = 0
    End If
    '
    Screen.MousePointer = 1
    If KQ& < 1 Then
        Call COMUNI("No Aparecen Créditos no Imputados\para la Cuenta Elegida.")
        GoTo 10
    End If
    '
    RGSEL$ = REGSEL$("!SELDOCAP")
    If Len(RGSEL$) <= 4 Then
        GoTo 10
    End If
    '
    FIMPUCRE.Label25.Caption = Left$(RGSEL$, 18)
    FIMPUCRE.Label24.Caption = Mid$(RGSEL$, 19, 8)
    FIMPUCRE.Label23.Caption = Mid$(RGSEL$, 27, 14)
    FIMPUCRE.Label22.Caption = Mid$(RGSEL$, 41, 14)
    FIMPUCRE.Label15.Caption = CVS(Mid$(RGSEL$, 61, 4))
    '
    FIMPUCRE.Label14.Caption = Mid$(RGSEL$, 41, 14)
    FIMPUCRE.Label20.Caption = Mid$(RGSEL$, 41, 14)
    FIMPUCRE.Command5.Enabled = True
    '
    Exit Sub
    '
999 Call CIERRARCH("*.*")
    FORECISTA.Show
    '
End Sub
'
Sub impucre2()
    '
    CANCE# = CDbl(Val(FIMPUCRE.Label18.Caption))
    If CANCE# < 0.005 Then
        Exit Sub
    End If
    '
    REG202& = Val(FIMPUCRE.Label15.Caption)
    If REG202& > 0 Then
        If REG202& <= ULTREG&("SVD202") Then
            X$ = REGLEIDO$("SVD202", REG202&)
            SALANT1# = CVD(Mid$(X$, 11, 8))
            SGG% = 1: If CVD(Mid$(X$, 21, 8)) < 0 Then SGG% = (-1)
            SALPE# = (Abs(CVD(Mid$(X$, 11, 8))) - CANCE#) * SGG%
            Call REPLA(X$, MKD$(SALPE#), 11, 8)
            Call GRAREG("SVD202", X$, REG202&)
        End If
    End If
    '
    VACON% = 1
    For U& = 1 To FIMPUCRE.List2.ListCount
        FIMPUCRE.List2.ListIndex = U& - 1
        X$ = FIMPUCRE.List2.TEXT
        CANCE# = CDbl(Val(Mid$(X$, 70, 14)))
        If CANCE# >= 0.005 Then
            NREG& = Val(Mid$(X$, 121, 6))
            If NREG& > 0 Then
                If NREG& <= ULTREG&("SVD202") Then
                    SS$ = REGLEIDO$("SVD202", NREG&)
                    DOCORI$ = FORMDOC$(CVS(Mid$(SS$, 5, 4)), Asc(Mid$(SS$, 3, 1)), Asc(Mid$(SS$, 4, 1)), 1)
                    ITOTORI# = CVD(Mid$(SS$, 21, 8))
                    IVAORI# = (Int(ALIVA * ITOTORI# / (1 + ALIVA / 100) + 0.5)) / 100
                    '
                    SALANT# = CVD(Mid$(SS$, 11, 8))
                    SALDEU# = CVD(Mid$(SS$, 11, 8)) - CANCE#
                    Call REPLA(SS$, MKD$(SALDEU#), 11, 8)
                    Call GRAREG("SVD202", SS$, NREG&)
                    '
                    NCI% = Val(FIMPUCRE.Label1.Caption)
                    HOII% = HOY(HO$)
                    REFAPLI$ = "Aplicacion Manual " + HO$
                    FEDOAPLI% = FECHANUM(TRIM$(FIMPUCRE.Label24.Caption))
                    DOCAPLI$ = TRIM$(FIMPUCRE.Label25.Caption)
                    
                    Z$ = REGBLAN$("APLICOB")
                    Call REPLA(Z$, MKI$(FEDOAPLI%), 1, 2)
                    Call REPLA(Z$, MKI$(NCI%), 3, 2)
                    Call REPLA(Z$, DOCAPLI$, 5, 18)
                    Call REPLA(Z$, MKD$(CANCE#), 23, 8)
                    Call REPLA(Z$, REFAPLI$, 31, 24)
                    Call REPLA(Z$, DOCORI$, 55, 18)
                    Call REPLA(Z$, MKD$(ITOTORI#), 73, 8)
                    Call REPLA(Z$, MKD$(IVAORI#), 81, 8)
                    '
                    Call REGAPP("APLICOB", Z$)
                    '
                End If
            End If
        End If
    Next U&
    VACON% = 0
    '
    Call CIERRARCH("*.*")
    '
End Sub
'

Sub IMPUCRE3()
    '
    If FIMPUCRE.List2.ListCount < 1 Then
        Exit Sub
    End If
    '
    Call BLANARCH("!APLIRECI")
    For U& = 1 To FIMPUCRE.List2.ListCount
        FIMPUCRE.List2.ListIndex = U& - 1
        Z$ = FIMPUCRE.List2.TEXT
        TIPOS$ = Left$(Z$, 12)
        NROL& = Val(Mid$(Z$, 13, 8))
        FEMII% = FECHANUM(Mid$(Z$, 23, 8))
        IPED# = CDbl(Val(Mid$(Z$, 31, 14)))
        FEVEI% = FECHANUM(Mid$(Z$, 48, 8))
        SALD# = CDbl(Val(Mid$(Z$, 56, 14)))
        ACRE# = CDbl(Val(Mid$(Z$, 70, 14)))
        DOCU$ = Mid$(Z$, 97, 18)
        REG2& = Val(Mid$(Z$, 121, 6))
        '
        Y$ = REGBLAN$("!APLIRECI")
        Call REPLA(Y$, TIPOS$, 1, 16)
        Call REPLA(Y$, MKS$(NROL&), 17, 4)
        Call REPLA(Y$, MKI$(FEMII%), 21, 2)
        Call REPLA(Y$, MKD$(IPED#), 23, 8)
        Call REPLA(Y$, MKI$(FEVEI%), 31, 2)
        Call REPLA(Y$, MKD$(SALD#), 33, 8)
        Call REPLA(Y$, MKD$(ACRE#), 41, 8)
        Call REPLA(Y$, DOCU$, 49, 18)
        Call REPLA(Y$, MKS$(REG2&), 67, 4)
        '
        Call REGAPP("!APLIRECI", Y$)
    Next U&
    Call CIERRARCH("!APLIRECI")
    '
200 ATRI$ = "/NC"
    ATRI$ = ATRI$ + "/TITUPAN=Aplicación de Cobranza"
    ATRI$ = ATRI$ + "/BORDESUP=W" + TRIM$(Str$(FIMPUCRE.Top + FIMPUCRE.Picture2.Top + FIMPUCRE.Frame1.Top))
    ATRI$ = ATRI$ + "/BORDEIZQ=W" + TRIM$(Str$(FIMPUCRE.Left + FIMPUCRE.Picture2.Left))
    '
    VTB% = VENTABU%("APLIRECI" + ATRI$)
    '
    VACON% = 1
    If VTB% >= 0 Then
        FIMPUCRE.List2.Clear
        For U& = 1 To ULTREG&("!APLIRECI")
            Y$ = REGLEIDO$("!APLIRECI", U&)
            TIPOS$ = Left$(Y$, 16)
            NROL& = CVS(Mid$(Y$, 17, 4))
            FEMI = CVI(Mid$(Y$, 21, 2))
            IPED# = CVD(Mid$(Y$, 23, 8))
            IMPTEX$ = CSTRING$(MKD$(IPED#), 4, 14, 2, 2)
            FEVE = CVI(Mid$(Y$, 31, 8))
            SALD# = CVD(Mid$(Y$, 33, 8))
            SALTEX$ = CSTRING$(MKD$(SALD#), 4, 14, 2, 2)
            ACRE# = CVD(Mid$(Y$, 41, 8))
            '
            If ACRE# < 0 Then
                Call MENSERR(24, "No se Permite Aplicación Negativa")
                GoTo 200
            End If
            If ACRE# > SALD# Then
                Call MENSERR(24, "El Importe Acreditado Debe Ser\Inferior al Saldo del Documento")
                GoTo 200
            End If
            ACRTEX$ = CSTRING$(MKD$(ACRE#), 4, 14, 2, 2)
            DOCU$ = Mid$(Y$, 49, 18)
            REG2& = CVS(Mid$(Y$, 67, 4))
            '
            Z$ = Space$(128)
            Call REPLA(Z$, TIPOS$, 1, 14)
            Call REPLA(Z$, CSTRING$(MKS$(NROL&), 3, 8, 0, 2), 13, 8)
            Call REPLA(Z$, FECHATEX$(FEMI), 23, 8)
            Call REPLA(Z$, IMPTEX$, 31, 14)
            Call REPLA(Z$, FECHATEX$(FEVE), 48, 8)
            Call REPLA(Z$, SALTEX$, 56, 14)
            Call REPLA(Z$, ACRTEX$, 70, 14)
            Call REPLA(Z$, DOCU$, 97, 18)
            Call REPLA(Z$, TRIM$(Str$(REG2&)), 121, 6)
            '
            FIMPUCRE.List2.AddItem Z$
        Next U&
    End If
    '
    SUMACRE# = 0
    For U& = 1 To FIMPUCRE.List2.ListCount
        FIMPUCRE.List2.ListIndex = U& - 1
        Z$ = FIMPUCRE.List2.TEXT
        SUMACRE# = SUMACRE# + CDbl(Val(Mid$(Z$, 70, 14)))
    Next U&
    '
    FIMPUCRE.Label10.Caption = CSTRING$(MKD$(SUMACRE#), 4, 14, 2, 2)
    FIMPUCRE.Label10.Refresh
    FIMPUCRE.Label18.Caption = TRIM$(CSTRING$(MKD$(SUMACRE#), 4, 14, 2, 2))
    FIMPUCRE.Label18.Refresh
    ANTI# = CDbl(Val(FIMPUCRE.Label14.Caption)) - SUMACRE#
    If ANTI# < 0 Then
          FIMPUCRE.Label19.Caption = "Descuentos:"
        Else
          FIMPUCRE.Label19.Caption = "Pendiente:"
    End If
    FIMPUCRE.Label20.Caption = TRIM$(CSTRING$(MKD$(ANTI#), 4, 14, 2, 2))
    FIMPUCRE.Label20.Refresh
    '
    FIMPUCRE.List2.Refresh
    VACON% = 1
    '
End Sub
'

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
          IMPO# = Abs(IMPO#)
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
          IMPO# = CVD(C5$)
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
          FVII% = CVI(SS4$)
          If FVII% > 0 Then
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
    VDEF$ = CLSS$ + MKD$(SARRASFE#) + MKD$(SALAJU#) + MKI$(FAJUI%)
    '
    Screen.MousePointer = 1
    '
    CLSS$ = OBJPLA$("CARAJUCLI", VDEF$)
    If CLSS$ = "" Or CLSS$ = "<ESC>" Then GoTo 5
    SALAJU# = CVD(Mid$(CLSS$, 13, 8))
    FEVEN% = CVI(Mid$(CLSS$, 21, 2))
    If FEVEN% < FAJUI% Then GoTo 100
    '
    DIFESAL# = SALAJU# - SARRASFE#
    DIFEDEU# = SALFINAL# - SALDEUTO# + DIFESAL#
    '
    If DIFESAL# < 0 Then
       FEVEN% = FAJUI%
    End If
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
       Call REPLA(FF0$, MKS$(NUAJU&), 5, 4)
       Call REPLA(SS0$, MKI$(FEVEN%), 9, 2)
       Call REPLA(SS0$, MKD$(DIFEDEU#), 11, 8)
       Call REPLA(SS0$, MKI$(FAJUI%), 19, 2)
       Call REPLA(SS0$, MKD$(DIFESAL#), 21, 8)
       Call REGAPP("SVD202", SS0$)
       '
       Call CIERRARCH("FACTUR")
       Call CIERRARCH("SVD202")
       '
       Screen.MousePointer = 1
       Call COMUNI("Ajuste Completo")
       '
    End If
    '
    Call CIERRARCH("*.*")
    Call LIBARCH("FACTUR")
    Call LIBARCH("SVD202")
    '
    GoTo 5
    '
99  Call CIERRARCH("*.*")
    '
End Sub
'

Sub REAPLICRE(NC%)
    '
    ATENTI% = 1
    Dim DOCU$(8192), IMPO#(8192), REGDO&(8192), FFDOC%(8192)
    '
100 Screen.MousePointer = 11
    '
    MAXDIAC& = ULTREG&("COBRAN")
    MAXE& = ULTREG&("FACTUR")
    '
    RFF$ = RECFILT$("FACTUR", 1, MKI$(NC%))
    RFC$ = RECFILT$("COBRAN", 1, MKI$(NC%))
    '
    NUDOCU% = 0
    FIN& = Len(RFF$)
    For U& = 1 To Len(RFF$) Step 4
        Call TRACE(20, U&, FIN&)
        REG& = CVS(Mid$(RFF$, U&, 4))
        X$ = REGLEIDO$("FACTUR", REG&)
        If CVI(Left$(X$, 2)) <> NC% Then
            X$ = REGLEIDO$("FACTUR", 1)
            Call GRAREG("FACTUR", X$, 1)
            Call CIERRARCH("FACTUR")
            GoTo 100
        End If
110     NUDO& = CVS(Mid$(X$, 3, 4))
        TI% = Asc(Mid$(X$, 13, 1))
        If NUDO& < 1 Then
           If TI% = 0 Or TI% = 7 Then
              Call REPLA(X$, MKS$(REG&), 3, 4)
              Call GRAREG("FACTUR", X$, REG&)
              GoTo 110
           End If
        End If
        VA% = Asc(Mid$(X$, 14, 1))
        SG% = 1: If TI% >= 4 Then If TI% <= 6 Then SG% = -1
        NUDOCU% = NUDOCU% + 1
        If NUDOCU% > 8192 Then
           Call MENSERR(24, "Imposible Continuar Proceso.\Demasiados Movimientos en la Cuenta Elegida.\Consulte a su Soporte de Sistemas FLEXOFT.")
           Call CIERRARCH("*.*")
           Call FINAL("")
        End If
        IMPO#(NUDOCU%) = CVD(Mid$(X$, 15, 8))
        IMPO#(NUDOCU%) = IMPO#(NUDOCU%) + CVD(Mid$(X$, 23, 8))
        IMPO#(NUDOCU%) = IMPO#(NUDOCU%) + CVD(Mid$(X$, 31, 8))
        IMPO#(NUDOCU%) = IMPO#(NUDOCU%) + CVD(Mid$(X$, 49, 8))
        IMPO#(NUDOCU%) = IMPO#(NUDOCU%) + CVD(Mid$(X$, 57, 8))
        IMPO#(NUDOCU%) = SG% * IMPO#(NUDOCU%)
        DOCU$(NUDOCU%) = Left$(FORMDOC$(NUDO&, TI%, VA%, 1), 2) + "-" + TRIM$(Str$(NUDO&))
        REGDO&(NUDOCU%) = REG&
        FFDOC%(NUDOCU%) = CVI(Mid$(X$, 7, 2))
    Next U&
    '
    FIN& = Len(RFC$)
    For U& = 1 To Len(RFC$) Step 4
        Call TRACE(20, U&, FIN&)
        REG& = CVS(Mid$(RFC$, U&, 4))
        X$ = REGLEIDO$("COBRAN", REG&)
        If CVI(Left$(X$, 2)) <> NC% Then
            X$ = REGLEIDO$("COBRAN", 1)
            Call GRAREG("COBRAN", X$, 1)
            Call CIERRARCH("COBRAN")
            GoTo 100
        End If
        NUDO& = CVS(Mid$(X$, 3, 4))
        TI% = 4
        VA% = 1
        SG% = (-1)
        NUDOCU% = NUDOCU% + 1
        If NUDOCU% > 8192 Then
           Call MENSERR(24, "Imposible Continuar Proceso.\Demasiados Movimientos en la Cuenta Elegida.\Consulte a su Soporte de Sistemas FLEXOFT.")
           Call CIERRARCH("*.*")
           Call FINAL("")
        End If
        IMPO#(NUDOCU%) = CVD(Mid$(X$, 15, 8))
        IMPO#(NUDOCU%) = SG% * IMPO#(NUDOCU%)
        DOCU$(NUDOCU%) = "RB-" + TRIM$(Str$(NUDO&))
        REGDO&(NUDOCU%) = (-1) * REG&
        FFDOC%(NUDOCU%) = CVI(Mid$(X$, 7, 2))
    Next U&
    '
    FIN& = NUDOCU%
    For KKI% = 1 To NUDOCU%
       U& = KKI%
       Call TRACE(20, U&, FIN&)
       For KKJ% = KKI% + 1 To NUDOCU%
         If FFDOC%(KKI%) > FFDOC%(KKJ%) Then
           DOCX$ = DOCU$(KKI%)
           FFX% = FFDOC%(KKI%)
           IMX# = IMPO#(KKI%)
           RGX& = REGDO&(KKI%)
           DOCU$(KKI%) = DOCU$(KKJ%)
           FFDOC%(KKI%) = FFDOC%(KKJ%)
           IMPO#(KKI%) = IMPO#(KKJ%)
           REGDO&(KKI%) = REGDO&(KKJ%)
           DOCU$(KKJ%) = DOCX$
           FFDOC%(KKJ%) = FFX%
           IMPO#(KKJ%) = IMX#
           REGDO&(KKJ%) = RGX&
         End If
       Next KKJ%
    Next KKI%
    '
    SUMAPEN# = 0
    For KKI% = 1 To NUDOCU%
       SUMAPEN# = SUMAPEN# + IMPO#(KKI%)
       If Abs(SUMAPEN#) < 0.005 Then
          For KKJ% = 1 To KKI%
             IMPO#(KKJ%) = 0
          Next KKJ%
       End If
    Next KKI%
    '
    If NUDOCU% >= 1 Then
       If Left$(DOCU$(1), 2) = "AS" Then
          FIN& = ULTREG&("SALINID")
          For U& = 1 To FIN&
             X$ = REGLEIDO$("SALINID", U&)
             If CVI(Left$(X$, 2)) = NC% Then
               TI% = Asc(Mid$(X$, 3, 1))
               VA% = Asc(Mid$(X$, 4, 1))
               SG% = 1: If TI% >= 4 Then If TI% <= 6 Then SG% = -1
               IPEN# = SG% * CVD(Mid$(X$, 11, 8))
               If Sgn(IPEN#) = Sgn(IMPO#(1)) Then
                 If Abs(IPEN#) <= Abs(IMPO#(1)) Then
                   IMPO#(1) = IMPO#(1) - IPEN#
                   NUDO& = CVS(Mid$(X$, 5, 4))
                   NUDOCU% = NUDOCU% + 1
                   If NUDOCU% > 8192 Then
                      Call MENSERR(24, "Imposible Continuar Proceso.\Demasiados Movimientos en la Cuenta Elegida.\Consulte a su Soporte de Sistemas FLEXOFT.")
                      Call CIERRARCH("*.*")
                      Call FINAL("")
                   End If
                   IMPO#(NUDOCU%) = IPEN#
                   DOCU$(NUDOCU%) = Left$(FORMDOC$(NUDO&, TI%, VA%, 1), 2) + "-" + TRIM$(Str$(NUDO&))
                   REGDO&(NUDOCU%) = 900000 + U&
                   FFDOC%(NUDOCU%) = CVI(Mid$(X$, 19, 2))
                 End If
               End If
             End If
          Next U&
       End If
    End If
    '
    FIN& = NUDOCU%
    For KKI% = 1 To NUDOCU%
       U& = KKI%
       Call TRACE(20, U&, FIN&)
       For KKJ% = KKI% + 1 To NUDOCU%
         If FFDOC%(KKI%) > FFDOC%(KKJ%) Then
           DOCX$ = DOCU$(KKI%)
           FFX% = FFDOC%(KKI%)
           IMX# = IMPO#(KKI%)
           RGX& = REGDO&(KKI%)
           DOCU$(KKI%) = DOCU$(KKJ%)
           FFDOC%(KKI%) = FFDOC%(KKJ%)
           IMPO#(KKI%) = IMPO#(KKJ%)
           REGDO&(KKI%) = REGDO&(KKJ%)
           DOCU$(KKJ%) = DOCX$
           FFDOC%(KKJ%) = FFX%
           IMPO#(KKJ%) = IMX#
           REGDO&(KKJ%) = RGX&
         End If
       Next KKJ%
    Next KKI%
    '
' HASTA AQUI O.K
    KKI% = 0
    For KKJ% = 1 To NUDOCU%
       If Abs(IMPO#(KKJ%)) >= 0.005 Then
          KKI% = KKI% + 1
          DOCU$(KKI%) = DOCU$(KKJ%)
          FFDOC%(KKI%) = FFDOC%(KKJ%)
          IMPO#(KKI%) = IMPO#(KKJ%)
          REGDO&(KKI%) = REGDO&(KKJ%)
       End If
    Next KKJ%
    NUDOCU% = KKI%
    '
    FIN& = ULTREG&("APLICOB")
    For U& = 1 To FIN&
       Call TRACE(20, U&, FIN&)
       X$ = REGLEIDO$("APLICOB", U&)
       If CVI(Mid$(X$, 3, 2)) = NC% Then
          RECI$ = Mid$(X$, 5, 18)
          RECI$ = Left$(RECI$, 2) + "-" + TRIM$(Val(Mid$(RECI$, 11)))
          FACU$ = Mid$(X$, 55, 18)
          PINI% = 11: If Mid$(FACU$, 17, 1) = "-" Then PINI% = 9
          FACU$ = Left$(FACU$, 2) + "-" + TRIM$(Val(Mid$(FACU$, PINI%, 8)))
          '
210       IAPLI# = CVD(Mid$(X$, 23, 8))
          IRECIMA# = 0: IFACMAX# = 0
          For KKI% = 1 To NUDOCU%
             If DOCU$(KKI%) = RECI$ Or (RECI$ = "AS-0" And Left$(DOCU$(KKI%), 2) = "AS") Then
               If IMPO#(KKI%) < 0 Then
                 IRECIMA# = IRECIMA# - IMPO#(KKI%)
               End If
             End If
             If DOCU$(KKI%) = FACU$ Then
               If IMPO#(KKI%) > 0 Then
                 IFACMAX# = IFACMAX# + IMPO#(KKI%)
               End If
             End If
          Next KKI%
          '
          If IAPLI# > IRECIMA# Then IAPLI# = IRECIMA#
          If IAPLI# > IFACMAX# Then IAPLI# = IFACMAX#
          '
          IRECIMA# = IAPLI#
          IFACMAX# = IAPLI#
          '
          If IAPLI# > 0 Then
             For KKI% = 1 To NUDOCU%
                If DOCU$(KKI%) = RECI$ Or (RECI$ = "AS-0" And Left$(DOCU$(KKI%), 2) = "AS") Then
                   If IMPO#(KKI%) < 0 Then
                      IAP1# = IRECIMA#
                      If IAP1# > Abs(IMPO#(KKI%)) Then IAP1# = Abs(IMPO#(KKI%))
                      IMPO#(KKI%) = IMPO#(KKI%) + IAP1#
                      IRECIMA# = IRECIMA# - IAP1#
                   End If
                End If
                If DOCU$(KKI%) = FACU$ Then
                   If IMPO#(KKI%) > 0 Then
                      IAP2# = IFACMAX#
                      If IAP2# > Abs(IMPO#(KKI%)) Then IAP2# = Abs(IMPO#(KKI%))
                      IMPO#(KKI%) = IMPO#(KKI%) - IAP2#
                      IFACMAX# = IFACMAX# - IAP2#
                   End If
                End If
             Next KKI%
          End If
       End If
    Next U&
    '
    FIN& = NUDOCU%
    For KKI% = 1 To NUDOCU%
       U& = KKI%
       Call TRACE(20, U&, FIN&)
       For KKJ% = KKI% + 1 To NUDOCU%
         If FFDOC%(KKI%) > FFDOC%(KKJ%) Then
           DOCX$ = DOCU$(KKI%)
           FFX% = FFDOC%(KKI%)
           IMX# = IMPO#(KKI%)
           RGX& = REGDO&(KKI%)
           DOCU$(KKI%) = DOCU$(KKJ%)
           FFDOC%(KKI%) = FFDOC%(KKJ%)
           IMPO#(KKI%) = IMPO#(KKJ%)
           REGDO&(KKI%) = REGDO&(KKJ%)
           DOCU$(KKJ%) = DOCX$
           FFDOC%(KKJ%) = FFX%
           IMPO#(KKJ%) = IMX#
           REGDO&(KKJ%) = RGX&
         End If
       Next KKJ%
    Next KKI%
    '
    SUMAPEN# = 0
    For KKI% = 1 To NUDOCU%
       SUMAPEN# = SUMAPEN# + IMPO#(KKI%)
       If Abs(SUMAPEN#) < 0.005 Then
          For KKJ% = 1 To KKI%
             IMPO#(KKJ%) = 0
          Next KKJ%
       End If
    Next KKI%
    '
    KKI% = 0
    For KKJ% = 1 To NUDOCU%
       If Abs(IMPO#(KKJ%)) >= 0.005 Then
          KKI% = KKI% + 1
          DOCU$(KKI%) = DOCU$(KKJ%)
          FFDOC%(KKI%) = FFDOC%(KKJ%)
          IMPO#(KKI%) = IMPO#(KKJ%)
          REGDO&(KKI%) = REGDO&(KKJ%)
       End If
    Next KKJ%
    NUDOCU% = KKI%
    '
    For KKI% = 1 To NUDOCU%
       For KKJ% = KKI% + 1 To NUDOCU%
          If Abs(IMPO#(KKI%) + IMPO#(KKJ%)) < 0.005 Then
             IMPO#(KKI%) = 0
             IMPO#(KKJ%) = 0
          End If
       Next KKJ%
    Next KKI%
    '
    KKI% = 0
    For KKJ% = 1 To NUDOCU%
       If Abs(IMPO#(KKJ%)) >= 0.005 Then
          KKI% = KKI% + 1
          DOCU$(KKI%) = DOCU$(KKJ%)
          FFDOC%(KKI%) = FFDOC%(KKJ%)
          IMPO#(KKI%) = IMPO#(KKJ%)
          REGDO&(KKI%) = REGDO&(KKJ%)
       End If
    Next KKJ%
    NUDOCU% = KKI%
    '
    Screen.MousePointer = 1
    If COMALTER%("Se ha Completado la Reconstrucción\de la Composición de la Deuda de\   \'" + TRIM$(RASOCLI$(NC%)) + "'\   \Desea Grabar la Nueva Composición de la Deuda ?\\Cancelar\Grabar") <> 2 Then
       GoTo 999
    End If
    '
    Screen.MousePointer = 11
    FIN& = ULTREG&("SVD202")
    For U& = 1 To FIN&
       Call TRACE(20, U&, FIN&)
       X$ = REGLEIDO$("SVD202", U&)
       If CVI(Left$(X$, 2)) = NC% Then
          Call REPLA(X$, String$(30, 0), 3, 30)
          Call GRAREG("SVD202", X$, U&)
       End If
    Next U&
    '
    LU$ = LUDAT$
    ARCHIDEU$ = "SVD202"
    RDX% = 1
    FIN& = NUDOCU%
    YABLO% = 0
    For KKI% = 1 To NUDOCU%
       U& = KKI%
       Call TRACE(20, U&, FIN&)
       If Abs(IMPO#(KKI%)) >= 0.005 Then
          If REGDO&(KKI%) > 0 Then
             If REGDO&(KKI%) > 900000 Then
                X$ = REGLEIDO$("SALINID", REGDO&(KKI%) - 900000)
                NRO = CVS(Mid$(X$, 5, 4))
                TIPO% = Asc(Mid$(X$, 3, 1))
                VA% = Asc(Mid$(X$, 4, 1))
                FFAC% = CVI(Mid$(X$, 9, 2))
                COPAGI% = 0
                GoTo 290
             End If
             X$ = REGLEIDO$("FACTUR", REGDO&(KKI%))
             NRO = CVS(Mid$(X$, 3, 4))
             TIPO% = Asc(Mid$(X$, 13, 1))
             VA% = Asc(Mid$(X$, 14, 1))
             FFAC% = CVI(Mid$(X$, 7, 2))
             COPAGI% = CPAGCLI%(NC%)
             If TIPO% < 1 Or TIPO% > 3 Then COPAGI% = 0
           Else
             X$ = REGLEIDO$("COBRAN", (-1) * REGDO&(KKI%))
             NRO = CVS(Mid$(X$, 3, 4))
             TIPO% = 4
             VA% = 0
             FFAC% = CVI(Mid$(X$, 7, 2))
             COPAGI% = 0
          End If
290       TOTFAC# = IMPO#(KKI%)
          If YABLO% = 0 Then
             YABLO% = 1
             Call BLOQARCH(ARCHIDEU$)
          End If
          Call GRADEU(LU$, ARCHIDEU$, RDX%, TIPO%, VA%, NRO, NC%, FFAC%, COPAGI%, TOTFAC#)
       End If
    Next KKI%
    '
999 Call CIERRARCH(ARCHIDEU$)
    Call CIERRARCH("*.*")
    Screen.MousePointer = 1
    ATENTI% = 0
End Sub

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
    Call REPLA(AX$, String$(50, 0), 15, 50)
    Call GRAREG("FACTUR", AX$, REFA&)
    Call CIERRARCH("FACTUR")
    '
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
    FIN& = ULTREG&("MOVISTO")
    '
    For I& = 1 To FIN&
        Call TRACE(20, I&, FIN&)
        X$ = REGLEIDO$("MOVISTO", I&)
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
    Call CIERRARCH("MOVISTO")
    Call CIERRARCH("*.*")
    '
    Call COMUNI("Anulación Completa")
    GoTo 10
    '
999 Call CIERRARCH("*.*")
    If UCase$(App.EXEName) = "FACTUMAI" Then
       OPFACMAI.Option1.Value = True
       OPFACMAI.Option1.Refresh
    End If
    '
End Sub
'

