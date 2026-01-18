Attribute VB_Name = "MABRUVE1"
Sub CALMABRU()
    '
3 ' MABRUVE.QBS      ' generacion archivo margen bruto de ventas

     Dim CANT(16384), IMPO#(16384)
     Dim REG&(16384), NRO(16384), INET#(16384)
     Dim LE$(10)

53  LU$ = LUDAT$

    TICO% = 1
    TICOS$ = "Valores Informados"
    '
600 CORRESP$ = TRIM$(MABRUVE.MESAN.Text)
    MMXX0$ = TRIM$(MABRUVE.MESAN.Text)
    MMXX1$ = MESANO$(MMXX0$, DES%, HAS%)
    If TRIM$(MMXX1$) = "" Then
        Call MENSERR(24, "Falta Seleccionar Período Mensual")
        Exit Sub
    End If
    '
700 FEX = DES%: DESS$ = FECHATEX$(FEX)
710 FEX = HAS%: HASS$ = FECHATEX$(FEX)
    '
    Call HEADERS("MABRUVE", "")
    Call HEADERS("MABRUVE", "Periodo Considerado: " + CORRESP$ + " - Desde: " + DESS$ + "      hasta: " + HASS$)
    Call HEADERS("MABRUVE", "Costos: " + TICOS$)
    '
    Screen.MousePointer = 11

1000 GoSub 5200                         ' abre archivo factur en #N13%

     Dim F0 As String * 64
     FIN& = LOF(N13%) / 64
     MABRUVE.ProgressBar1.Min = 0
     If FIN& > 0 Then
        MABRUVE.ProgressBar1.Max = FIN&
     End If
     '
     I& = 0: CAFAC& = 0: SUFAC# = 0
     For I& = 1 To FIN&                ' lee secuencialmente archivo FACTUR
        MABRUVE.ProgressBar1.Value = I&
        Get #N13%, I&, F0$

        TIPO% = Asc(Mid$(F0$, 13, 1))

        If (TIPO% = 1 Or TIPO% = 2) Then
            FEVEN% = CVI(Mid$(F0$, 7, 2))

            If (FEVEN% >= DES% And FEVEN% <= HAS%) Then
                '
                CAFAC& = CAFAC& + 1
                REG&(CAFAC&) = I&
                NRO(CAFAC&) = CVS(Mid$(F0$, 3, 4))
                INET#(CAFAC&) = CVD(Mid$(F0$, 15, 8))
                SUFAC# = SUFAC# + (Int(100 * CVD(Mid$(F0$, 15, 8)) + 0.5)) / 100
                '
            End If
        End If
1009 Next I&
     Close #N13%: N13% = 0

     FIN& = ULTREG&("MOVISTO")
     MABRUVE.ProgressBar1.Min = 0
     If FIN& > 0 Then
        MABRUVE.ProgressBar1.Max = FIN&
     End If
     '
     For I& = 1 To ULTREG&("MOVISTO")
        '
        MABRUVE.ProgressBar1.Value = I&
        X$ = REGLEIDO$("MOVISTO", I&)
        '
        FEVEN% = CVI(Left$(X$, 2))
        If FEVEN% >= DES% Then
            If FEVEN% <= HAS% Then
                TIMO0$ = UCase$(Mid$(X$, 21, 2))
                TIMO1$ = UCase$(Mid$(X$, 23, 2))
                TIMO2$ = UCase$(Mid$(X$, 33, 2))
                If (TIMO0$ = "RT" Or TIMO0$ = "FC" Or TIMO1$ = "RT" Or TIMO1$ = "FC" Or TIMO2$ = "RT" Or TIMO2$ = "FC") Then
                    CI% = CVI(Mid$(X$, 3, 2))
                    If CI% < 1 Or CI% > NUMAS% Then ' MATERIAL LIBRE
                        CI% = NUMAS% + 1
                    End If
                    CAN# = CVD(Mid$(X$, 96, 8)) - CVD(Mid$(X$, 88, 8))
                    IMPOD# = CVD(Mid$(X$, 73, 8)) * CAN#
                    '
                    CANT(CI%) = CANT(CI%) + CAN#
                    IMPO#(CI%) = IMPO#(CI%) + IMPOD#
                    '
                End If
            End If
        End If
        '
     Next I&
     '
1300 GoSub 5200
     FIN& = LOF(N13%) / 64
     MABRUVE.ProgressBar1.Min = 0
     If FIN& > 0 Then
        MABRUVE.ProgressBar1.Max = FIN&
     End If
     '
1305 I& = 0: CAFAC& = 0
1310 For I& = 1 To LOF(N13%) / 64
        MABRUVE.ProgressBar1.Value = I&
        Get #N13%, I&, F0$
        F1$ = Left$(F0$, 2)
        F3$ = Mid$(F0$, 7, 2)
        F6$ = Mid$(F0$, 13, 1)
        F7$ = Mid$(F0$, 15, 8)
        '
        If CVI(F1$) > 0 Then
                TIPO% = Asc(F6$)
                If TIPO% <> 3 Then If TIPO% <> 5 Then If TIPO% <> 6 Then GoTo 1390 ' NO ES DEBITO NI CREDITO
                FEVEN% = CVI(F3$)
                If FEVEN% < DES% Or FEVEN% > HAS% Then GoTo 1390

                If TIPO% = 3 Then
                        CANT(NUMAS% + 2) = CANT(NUMAS% + 2) + 1
                        IMPO#(NUMAS% + 2) = IMPO#(NUMAS% + 2) + CVD(F7$)
                        GoTo 1390
                End If

                If (TIPO% = 5 Or TIPO% = 6) Then
                        CANT(NUMAS% + 3) = CANT(NUMAS% + 3) + 1
                        IMPO#(NUMAS% + 3) = IMPO#(NUMAS% + 3) - CVD(F7$)
                        GoTo 1390
                End If
        End If
1390 Next I&
     Close #N13%
     '
     LE$(1) = "MATERIAL SIN CODIGO"
     LE$(2) = "DEBITOS VARIOS"
     LE$(3) = "CREDITOS VARIOS"

     For J& = 1 To 3
        While Len(LE$(J&)) < 48
                LE$(J&) = LE$(J&) + " "
        Wend
     Next J&

     J& = 0
     NM& = NUMAS%
     For I& = 1 To NM& + 3
        CMOV$ = "0"
        If I& > NM& Then
                CMOV$ = Chr$(48 + I& - NM&)
        End If
        REGSS$ = CMOV$

        If Abs(CANT(I&)) > 0 Or Abs(IMPO#(I&)) >= 0.005 Then
                CI% = I&

                COD$ = Space$(15)
                If CI% <= NUMAS% Then
                        COD$ = Left$(CODEXT$(CI%), 15)
                End If
                REGSS$ = REGSS$ + COD$

                If CI% <= NUMAS% Then
                        DE$ = DESCRIT0$(CI%)
                        REGSS$ = REGSS$ + AJUSTI$(DE$, 48)
                        GoTo 1410
                End If

                If (CI% > NUMAS% And CI% < NUMAS% + 4) Then
                        REGSS$ = REGSS$ + LE$(CI% - NUMAS%)
                End If

1410            REGSS$ = REGSS$ + MKS$(CANT(I&))

                PREPROM# = IMPO#(I&) / CANT(I&)
1420            'If Abs(PREPROM#) < 0.01 Then
                '    Call COMUNI("&9.25&ATENCION: Falta Precio Promedio\del Codigo " + TRIM$(CODEXT$(CI%)) + "\" + TRIM$(DESCRIT$(CI%)) + ".")
                '    PREPROM# = Val(InputBox$("&15.55&Precio Promedio de Venta", 10))
                '    GoTo 1420
                'End If
                REGSS$ = REGSS$ + MKD$(PREPROM#)

                COU# = 0
                If CI% <= NUMAS% Then
                        COU# = COSUNI(CI%)
                        'If TICO% = 2 Then
                        '   GET #N9%,CI%
                        '   COU# = CVSMBF(Q43$)
                        'End If
                End If

                REGSS$ = REGSS$ + MKD$(COU#)
                If COU# < 0.00005 Then
                    TTXX$ = "Codigo: " + TRIM$(CODEXT$(CI%))
                    TTXX$ = TTXX$ + "\" + DESCRIT0$(CI%)
                    TTXX$ = TTXX$ + "\   \Falta Informar Costo de Reposicion."
                    Call COMUNI(TTXX$)
                End If
                '
                MABRUNI# = PREPROM# - COU#
                If CMOV$ > "0" Then MABRUNI# = 0
                REGSS$ = REGSS$ + MKD$(MABRUNI#)

                MABRUTO# = MABRUNI# * CANT(I&)
                If CMOV$ > "0" Then MABRUTO# = 0
                REGSS$ = REGSS$ + MKD$(MABRUTO#)

                PORCOU = 99999999
                If COU# > 0.01 Then PORCOU = 100 * MABRUNI# / COU#
                REGSS$ = REGSS$ + MKS$(PORCOU)

                PORPRE = 99999999
                If PREPROM# > 0.01 Then PORPRE = 100 * MABRUNI# / PREPROM#
                REGSS$ = REGSS$ + MKS$(PORPRE)


                While Len(REGSS$) < 128
                        REGSS$ = REGSS$ + Chr$(0)
                Wend

                J& = J& + 1
                Call GRAREG("MABRUVE", REGSS$, J&)
        End If
     Next I&
     Call SETULTREG("MABRUVE", J&)
     Call CIERRARCH("*.*")
     Screen.MousePointer = 1
     '
3900 If MABRUVE.Option5.Value = True Then
           Call FINAL("*SMABRUVS")
        Else
           Call FINAL("*SMABRUV")
     End If
     '
Exit Sub
     

5200 N13% = FILEOPEN%(LU$ + "FACTUR.DAT", 0, 64)
     Return
     '
     '
End Sub


Sub CARCOSUN()
    '
    Call BLANARCH("!CARCOREP")
    Screen.MousePointer = 11
    FIN& = NUINV%
    MABRUVE.ProgressBar1.Min = 0
    If FIN& > 0 Then
        MABRUVE.ProgressBar1.Max = FIN&
    End If
    '
    For U& = 1 To NUINV%
        MABRUVE.ProgressBar1.Value = U&
        X$ = REGLEIDO$("INVERT", U& + 1)
        If Mid$(X$, 16, 1) <> "`" Then
           CI% = CVI(Mid$(X$, 17, 2))
           If CI% > 0 Then
             If CI% <= NUMAS% Then
                COU = COSUNI(CI%)
                FECO% = FECOSTO%(CI%)
                Y$ = REGBLAN$("!CARCOREP")
                Call REPLA(Y$, MKI$(CI%), 1, 2)
                Call REPLA(Y$, MKS$(COU), 3, 4)
                Call REPLA(Y$, MKI$(FECO%), 7, 2)
                Call REGAPP("!CARCOREP", Y$)
             End If
           End If
        End If
    Next U&
    Call CIERRARCH("!CARCOREP")
    Screen.MousePointer = 1
    '
    VTB% = VENTABU%("CARCOREP")
    '
    Screen.MousePointer = 11
    FIN& = ULTREG&("!CARCOREP")
    MABRUVE.ProgressBar1.Min = 0
    If FIN& > 0 Then
        MABRUVE.ProgressBar1.Max = FIN&
    End If
    '
    If VTB% >= 0 Then
        For U& = 1 To ULTREG&("!CARCOREP")
            MABRUVE.ProgressBar1.Value = U&
            Y$ = REGLEIDO$("!CARCOREP", U&)
            CI% = CVI(Left$(Y$, 2))
            If CI% > 0 Then
                If CI% <= NUMAS% Then
                    REGMA& = CI% + 1
                    COU = CVS(Mid$(Y$, 3, 4))
                    FECO% = CVI(Mid$(Y$, 7, 2))
                    X$ = REGLEIDO$("MASTER", REGMA&)
                    Call REPLA(X$, MKS$(COU), 111, 4)
                    Call REPLA(X$, MKI$(FECO%), 115, 2)
                    Call GRAREG("MASTER", X$, REGMA&)
                End If
            End If
        Next U&
    End If
    '
    Call CIERRARCH("*.*")
    Screen.MousePointer = 1
    Call COMUNI("Actualización Completa")
    '
End Sub
'

Sub LICOSUN()
    '
    Screen.MousePointer = 11
    FIN& = NUINV%
    MABRUVE.ProgressBar1.Min = 0
    If FIN& > 0 Then
        MABRUVE.ProgressBar1.Max = FIN&
    End If
    '
    J& = 0
    For U& = 1 To NUINV%
        MABRUVE.ProgressBar1.Value = U&
        X$ = REGLEIDO$("INVERT", U& + 1)
        If Mid$(X$, 16, 1) <> "`" Then
           CI% = CVI(Mid$(X$, 17, 2))
           If CI% > 0 Then
             If CI% <= NUMAS% Then
                COU = COSUNI(CI%)
                FECO% = FECOSTO%(CI%)
                Y$ = REGBLAN$("CARCOREP")
                Call REPLA(Y$, MKI$(CI%), 1, 2)
                Call REPLA(Y$, MKS$(COU), 3, 4)
                Call REPLA(Y$, MKI$(FECO%), 7, 2)
                J& = J& + 1
                Call GRAREG("CARCOREP", Y$, J&)
             End If
           End If
        End If
    Next U&
    Call SETULTREG("CARCOREP", J&)
    Call CIERRARCH("CARCOREP")
    Screen.MousePointer = 1
    '
    Call CONECRUN("*LICOREP", "Listado de Costos de Reposición")
    '
End Sub
'

Sub EXCOSUN()
    '
    ARIMPO$ = TRIM$(MABRUVE.Text1.Text)
    If ARIMPO$ = "" Then
        Call MENSERR(24, "Falta Seleccionar Archivo de Exportación")
        Exit Sub
    End If
    '
    ASAL% = FILEOPEN%(ARIMPO$, 2, LGG%)
    '
    Screen.MousePointer = 11
    FIN& = NUINV%
    MABRUVE.ProgressBar1.Min = 0
    If FIN& > 0 Then
        MABRUVE.ProgressBar1.Max = FIN&
    End If
    '
    J& = 0
    For U& = 1 To NUINV%
        MABRUVE.ProgressBar1.Value = U&
        X$ = REGLEIDO$("INVERT", U& + 1)
        If Mid$(X$, 16, 1) <> "`" Then
           CI% = CVI(Mid$(X$, 17, 2))
           If CI% > 0 Then
             If CI% <= NUMAS% Then
                COU = COSUNI(CI%)
                FECO% = FECOSTO%(CI%)
                Y$ = Space$(108)
                Call REPLA(Y$, CODEXT$(CI%), 1, 16)
                Call REPLA(Y$, DESCRIT0$(CI%), 17, 64)
                Call REPLA(Y$, UNIMED$(CI%), 81, 4)
                Call REPLA(Y$, CSTRING$(MKS$(COU), 4, 12, 4, 2), 85, 12)
                FEX = FECO%
                Call REPLA(Y$, FECHATEX$(FEX), 101, 8)
                Print #ASAL%, Y$
             End If
           End If
        End If
    Next U&
    Close #ASAL%
    Screen.MousePointer = 1
    '
    Call COMUNI("Exportación Completa")
    '
End Sub
'

Sub CALMANOVO()
   MABRUNOV.List1.Clear
   '
600 CORRESP$ = TRIM$(MABRUNOV.MESAN.Text)
    MMXX0$ = TRIM$(MABRUNOV.MESAN.Text)
    MMXX1$ = MESANO$(MMXX0$, DES%, HAS%)
    If TRIM$(MMXX1$) = "" Then
        Call MENSERR(24, "Falta Seleccionar Período Mensual")
        Exit Sub
    End If
    '
700 FEX = DES%: DESS$ = FECHATEX$(FEX)
710 FEX = HAS%: HASS$ = FECHATEX$(FEX)
    '
    FIN& = ULTREG&("PEDDETA")
    For U& = 1 To FIN&
       Call TRACE(20, U&, FIN&)
       X$ = REGLEIDO$("PEDDETA", U&)
       FF% = CVI(Mid$(X$, 103, 2))
       If FF% >= DES% Then
          If FF% <= HAS% Then
             Z$ = Space$(160)
             NUREMIT& = CVS(Mid$(X$, 105, 4))
             FEREMITO = FF%
             CI% = CVI(Mid$(X$, 33, 2))
             LARGO = CVS(Mid$(X$, 59, 4))
             ANCHO = CVS(Mid$(X$, 63, 4))
             ESPES = CVS(Mid$(X$, 67, 4))
             COEF = 1
             If LARGO >= 1 Then
               If ANCHO >= 1 Then
                 If ESPES >= 1 Then
                   COEF = (LARGO / 1000) * (ANCHO / 1000) * (ESPES / 1000)
                 End If
               End If
             End If
             PRELIS = CVS(Mid$(X$, 71, 4))
             PREUNI = PRELIS * COEF
             '
             Call REPLA(Z$, CSTRING$(MKI$(FF%), 1, 4, 0, 2), 1, 4)
             Call REPLA(Z$, CSTRING$(MKS$(NUREMIT&), 3, 6, 0, 2), 9, 6)
             Call REPLA(Z$, "A " + CSTRING$(MKS$(U&), 3, 6, 0, 2), 17, 8)
             Call REPLA(Z$, CODEXT$(CI%), 27, 16)
             Call REPLA(Z$, CSTRING$(MKS$(LARGO), 3, 6, 0, 2), 43, 6)
             Call REPLA(Z$, CSTRING$(MKS$(ANCHO), 3, 6, 0, 2), 49, 6)
             Call REPLA(Z$, CSTRING$(MKS$(ESPES), 3, 6, 0, 2), 55, 6)
             Call REPLA(Z$, CSTRING$(MKS$(COEF), 3, 10, 5, 2), 61, 10)
             Call REPLA(Z$, CSTRING$(MKS$(PRELIS), 3, 10, 2, 2), 71, 10)
             Call REPLA(Z$, CSTRING$(MKS$(PREUNI), 3, 10, 5, 2), 81, 10)
             '
             MABRUNOV.List1.AddItem Z$
          End If
       End If
    Next U&
    '
    Dim COCALMP(16384), COCALMO(16384), COCALST(16384)
    For U& = 1 To ULTREG&("COCALCU")
      X$ = REGLEIDO$("COCALCU", U&)
      Q9$ = REGLEIDO$("MODIREC", U&)
      CI% = CVI(Left$(X$, 2))
      COCALMP(CI%) = CVS(Mid$(X$, 15, 4))
      COCALMO(CI%) = CVS(Mid$(Q9$, 3, 4)) + CVS(Mid$(Q9$, 7, 4))
      COCALST(CI%) = CVS(Mid$(X$, 39, 4))
    Next U&
    '
    MABRUNOV.List2.Clear
    FIN& = ULTREG&("MOVISTO")
    For U& = 1 To FIN&
       Call TRACE(20, U&, FIN&)
       X$ = REGLEIDO$("MOVISTO", U&)
       FF% = CVI(Mid$(X$, 1, 2))
       If FF% >= DES% Then
          If FF% <= HAS% Then
            If Mid$(X$, 23, 2) = "FC" Or Mid$(X$, 33, 2) = "FC" Then
              NC% = CVI(Mid$(X$, 84, 2))
              NUFA& = Val(Mid$(X$, 26, 7))
              If NUFA& < 1 Then
                NUFA& = Val(Mid$(X$, 36, 7))
              End If
              PREUNI = CVD(Mid$(X$, 73, 8))
              CAENT = CVD(Mid$(X$, 96, 8)) - CVD(Mid$(X$, 88, 8))
              For T& = 1 To MABRUNOV.List1.ListCount
                Z$ = MABRUNOV.List1.List(T& - 1)
                If Val(Mid$(Z$, 9, 6)) = NUFA& Then
                  If Val(Mid$(Z$, 91, 10)) < 0.0001 Then
                    COEF = Val(Mid$(Z$, 61, 10))
                    CI% = CODINT%(Mid$(Z$, 27, 16))
                    COMP = COCALMP(CI%) * COEF
                    COMO = COCALMO(CI%)
                    COST = COCALST(CI%)
                    COTT = COMP + COMO + COST
                    Call REPLA(Z$, CSTRING$(MKS$(PREUNI), 3, 10, 5, 2), 91, 10)
                    Call REPLA(Z$, CSTRING$(MKS$(CAENT), 3, 9, 1, 2), 101, 10)
                    Call REPLA(Z$, CSTRING$(MKS$(COMP), 3, 9, 5, 2), 111, 10)
                    Call REPLA(Z$, CSTRING$(MKS$(COMO), 3, 9, 5, 2), 121, 10)
                    Call REPLA(Z$, CSTRING$(MKS$(COST), 3, 9, 5, 2), 131, 10)
                    Call REPLA(Z$, CSTRING$(MKS$(COTT), 3, 9, 5, 2), 141, 10)
                    Call REPLA(Z$, CSTRING$(MKI$(NC%), 1, 6, 0, 2), 151, 6)
                    '
                    MABRUNOV.List1.RemoveItem T& - 1
                    MABRUNOV.List2.AddItem Z$
                    GoTo 19
                  End If
                End If
              Next T&
             Z$ = Space$(160)
             NUREMIT& = NUFA&
             FEREMITO = FF%
             CI% = CVI(Mid$(X$, 3, 2))
             PRELIS = 0
             '
             Call REPLA(Z$, CSTRING$(MKI$(FF%), 1, 4, 0, 2), 1, 4)
             Call REPLA(Z$, CSTRING$(MKS$(NUREMIT&), 3, 6, 0, 2), 9, 6)
             Call REPLA(Z$, "A " + CSTRING$(MKS$(U&), 3, 6, 0, 2), 17, 8)
             Call REPLA(Z$, CODEXT$(CI%), 27, 16)
             Call REPLA(Z$, CSTRING$(MKS$(PRELIS), 3, 10, 2, 2), 71, 10)
             '
             COEF = 1
             PRELIS = 0
             PRELI = PRELISTA("LIPRE001", CI%)
             If PRELI > 0 Then
                COEF = PREUNI / PRELI
             End If
             COMP = COCALMP(CI%) * COEF
             COMO = COCALMO(CI%)
             COST = COCALST(CI%)
             COTT = COMP + COMO + COST
             Call REPLA(Z$, CSTRING$(MKS$(COEF), 3, 10, 5, 2), 61, 10)
             Call REPLA(Z$, CSTRING$(MKS$(PREUNI), 3, 10, 5, 2), 91, 10)
             Call REPLA(Z$, CSTRING$(MKS$(CAENT), 3, 9, 1, 2), 101, 10)
             Call REPLA(Z$, CSTRING$(MKS$(COMP), 3, 9, 5, 2), 111, 10)
             Call REPLA(Z$, CSTRING$(MKS$(COMO), 3, 9, 5, 2), 121, 10)
             Call REPLA(Z$, CSTRING$(MKS$(COST), 3, 9, 5, 2), 131, 10)
             Call REPLA(Z$, CSTRING$(MKS$(COTT), 3, 9, 5, 2), 141, 10)
             Call REPLA(Z$, CSTRING$(MKI$(NC%), 1, 6, 0, 2), 151, 6)
             MABRUNOV.List2.AddItem Z$
             '
            End If
          End If
       End If
19  Next U&
    '
    J& = 0
    For U& = 1 To MABRUNOV.List2.ListCount
       Z$ = MABRUNOV.List2.List(U& - 1)
       NUFA& = Val(Mid$(Z$, 9, 6))
       FEFA% = Val(Mid$(Z$, 1, 4))
       NC% = Val(Mid$(Z$, 151, 6))
       CI% = CODINT%(Mid$(Z$, 27, 16))
       MEDI$ = TRIM$(Mid$(Z$, 43, 18))
       If MEDI$ <> "" Then
            MEDI$ = TRIM$(Mid$(Z$, 43, 6)) + "x" + TRIM$(Mid$(Z$, 49, 6)) + "x" + TRIM$(Mid$(Z$, 55, 6))
          Else
            If Val(Mid$(Z$, 61, 10)) < 0.999 Then
               MEDI$ = TRIM$(Mid$(Z$, 61, 10)) + " m3)"
            End If
       End If
       CAN = Val(Mid$(Z$, 101, 10))
       PREU = Val(Mid$(Z$, 91, 10))
       IFAC# = CDbl(CAN * PREU)
       COMP = CAN * Val(Mid$(Z$, 111, 10))
       COMO = CAN * Val(Mid$(Z$, 121, 10))
       COST = CAN * Val(Mid$(Z$, 131, 10))
       COTT = CAN * Val(Mid$(Z$, 141, 10))
       MBRU# = IFAC# - COTT
       PORMB = 999
       If Abs(IFAC#) > 1 Then PORMB = 100 * MBRU / IFAC#
       '
       X$ = REGBLAN$("AMAVEN")
       Call REPLA(X$, MKS$(NUFA&), 1, 4)
       Call REPLA(X$, MKI$(FEFA%), 5, 2)
       Call REPLA(X$, MKI$(NC%), 7, 2)
       Call REPLA(X$, MKI$(CI%), 9, 2)
       Call REPLA(X$, MEDI$, 11, 24)
       Call REPLA(X$, MKS$(CAN), 35, 4)
       Call REPLA(X$, MKS$(PREU), 39, 4)
       Call REPLA(X$, MKD$(IFAC#), 43, 8)
       Call REPLA(X$, MKS$(COMP), 65, 4)
       Call REPLA(X$, MKS$(COMO), 69, 4)
       Call REPLA(X$, MKS$(COST), 73, 4)
       Call REPLA(X$, MKS$(COTT), 93, 4)
       Call REPLA(X$, MKD$(MBRU#), 97, 8)
       Call REPLA(X$, MKS$(PORMB), 105, 4)
       '
       J& = J& + 1
       Call GRAREG("AMAVEN", X$, J&)
    Next U&
    '
    Call SETULTREG("AMAVEN", J&)
    Call CIERRARCH("AMAVEN")
    Call GRARGEN("AMAVEN", MMXX1$)
    '
    Call DATARGEN("AMAVEN", FEGE%, HOGE%, USGE%, TEGE%, CORRE$)
    FEX = FEGE%
    MABRUNOV.Label1.Caption = FECHATEX$(FEX)
    MABRUNOV.Label3.Caption = HORATEX$(HOGE%)
    MABRUNOV.MESAN.Text = MMXX1$
    '
    MABRUNOV.Command1.Enabled = True
    MABRUNOV.Command4.Enabled = True
    MABRUNOV.Command2.Enabled = True
    '
    Screen.MousePointer = 1
End Sub
'

Sub LIANACOS()
    '
    MODO% = COMALTER%("Costos de Materiales por Producto\\Pantalla\Impresión\Cancelar")
    If MODO% = 1 Then
       Call CIERRARCH("*.*")
       Call CONECRUN("*ANACOS/" + COD$, "Consulta de Analisis de Costos por Pieza")
       Exit Sub
    End If
    '
    If MODO% <> 2 Then Exit Sub
    '
    XX$ = Left$(REGLEIDO$("ECOMAST", 1), 16): CI1% = Abs(CODINT%(XX$))
    XX$ = Left$(REGLEIDO$("ECOMAST", ULTREG&("ECOMAST")), 16): CI2% = Abs(CODINT%(XX$))
    VDEF$ = MKI$(CI1%) + MKI$(CI2%)
10  RACOD$ = OBJPLA$("DESHASCOD1", VDEF$)
    If RACOD$ = "" Then
        Exit Sub
    End If
    CI1% = CVI(Left$(RACOD$, 2)): cod1$ = CODEXT$(CI1%)
    CI2% = CVI(Mid$(RACOD$, 3, 2)): cod2$ = CODEXT$(CI2%)
    TMT% = Asc(Mid$(RACOD$, 5, 1))
    TIMAT$ = ECOARCH$("TATIMAT", Chr$(TMT%))
    If TIMAT$ = "" Then
        TMT% = 0
        Call REPLA(RACOD$, Chr$(0), 5, 1)
        TIMAT$ = "Sin Clasificación"
    End If
    '
    If cod1$ = "" Or cod2$ = "" Or cod2$ < cod1$ Then
        VDEF$ = RACOD$
        GoTo 10
    End If
    '
    Screen.MousePointer = 11
    Dim PUNMAS$(32767)
    '
    FIN& = ULTREG&("ANACOS")
    For U& = 1 To FIN&
       Call TRACE(20, U&, FIN&)
       X$ = REGLEIDO$("ANACOS", U&)
       CIXI% = CVI(Left$(X$, 2))
       If CIXI% > 0 Then
         If CIXI% <= 32767 Then
           PUNMAS$(CIXI%) = PUNMAS$(CIXI%) + MKS$(U&)
         End If
       End If
    Next U&
    '
    If COMALTER%("Saltea o Incluye Subconjuntos\\Saltear\Incluir") <> 2 Then
       FIN& = ULTREG&("ESTRUNUE")
       FINDOS& = ULTREG&("ESTRUDOS")
       For REG& = 1 To FIN&
         Call TRACE(20, REG&, FIN&)
         X$ = REGLEIDO$("ESTRUNUE", REG&)
         DPS$ = Mid$(X$, 3, 120)
         If REG& <= FINDOS& Then
            DPS$ = DPS$ + REGLEIDO$("ESTRUDOS", REG&)
         End If
         '
         For K% = 1 To Len(DPS$) Step 8
            DPSX$ = Mid$(DPS$, K%, 8)
            CIKK% = CVI(Left$(DPSX$, 2))
            If CIKK% > 0 Then
              If CIKK% <= 32767 Then
                If Abs(CVS(Mid$(DPSX$, 3))) >= 0.000005 Then
                  PUNMAS$(CIKK%) = ""
                End If
              End If
            End If
         Next K%
         '
       Next REG&
    End If
    '
    FIN& = NUINV%
    REGQ& = 0
    For U& = 1 To FIN&
        Call TRACE(20, U&, FIN&)
        X$ = REGLEIDO$("INVERT", U& + 1)
        CODYY$ = Left$(X$, 16)
        If CODYY$ >= cod1$ Then
          If CODYY$ <= cod2$ Then
            If Mid$(CODYY$, 16, 1) <> Chr$(96) Then
              CIXI% = CVI(Mid$(X$, 17, 2))
              If CIXI% > 0 Then
                If CIXI% <= 32767 Then
                  If TMT% = 0 Or TIMATE%(CIXI%) = TMT% Then
                    If Len(PUNMAS$(CIXI%)) > 0 Then
                      '
                      If REGQ& > 0 Then
                        REGQ& = REGQ& + 1
                        Call GRAREG("LIANACOS", String$(128, "="), REGQ&)
                      End If
                      '
                      XX$ = REGBLAN$("LIANACOS")
                      REGQ& = REGQ& + 1
                      Call GRAREG("LIANACOS", XX$, REGQ&)
                      '
                      XX$ = REGBLAN$("LIMULTI")
                      Call REPLA(XX$, CODEXT$(CIXI%), 3, 16)
                      Call REPLA(XX$, DESCRIT0$(CIXI%), 19, 48)
                      REGQ& = REGQ& + 1
                      Call GRAREG("LIANACOS", XX$, REGQ&)
                      '
                      XX$ = REGBLAN$("LIANACOS")
                      REGQ& = REGQ& + 1
                      Call GRAREG("LIANACOS", XX$, REGQ&)
                      '
                      For J& = 1 To Len(PUNMAS$(CIXI%)) Step 4
                        JJ& = CVS(Mid$(PUNMAS$(CIXI%), J&, 4))
                        XX$ = REGLEIDO$("ANACOS", JJ&)
                        REGQ& = REGQ& + 1
                        Call GRAREG("LIANACOS", XX$, REGQ&)
                      Next J&
                      '
                    End If
                  End If
                End If
              End If
            End If
          End If
        End If
        '
    Next U&
    '
    Call SETULTREG("LIANACOS", REGQ&)
    Call CIERRARCH("LIANACOS")
    Screen.MousePointer = 1
    '
    Call CONECRUN("*LIANACO", "Listado Analítico de Costos Calculados")
    '
End Sub
