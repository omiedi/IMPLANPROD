Attribute VB_Name = "NEGASTO1"
Dim IAT&
'

Sub EXISTO2()

3   ' EXISTO.QBS    DETERMINACION EXISTENCIAS DE STOCK

    CI1% = PRICODACT%
    CI2% = ULTCODACT%
    TMTX% = 0
    If InStr(EMPREAC$, "DOSIVAC") > 0 Then
       TMTX% = 5
       If HOY(HOS$) = FECHANUM("20/12/01") Then
         Call CLASIAUTDOSI
       End If
    End If
    VDEF$ = MKI$(CI1%) + MKI$(CI2%) + Chr$(TMTX%)
    
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
       TIMAT$ = "Material No Clasificado"
    End If
    '
    If cod1$ = "" Or cod2$ = "" Or cod2$ < cod1$ Then
       VDEF$ = RACOD$
       GoTo 10
    End If
    '
    Call HEADERS("NEGASTO", "")
    Call HEADERS("LIMASEL", "Tipo de Material: " + TIMAT$)
    '
    J& = 0             'Call BLANARCH("NEGASTO")
    Screen.MousePointer = 11
    NVII& = NUINV%
    For I& = 1 To NVII&
      Call TRACE(20, I&, NVII&)
      II$ = REGLEIDO$("INVERT", I& + 1)
      CIXI% = CVI(Mid$(II$, 17, 2))
      If Mid$(II$, 16, 1) <> "`" Then
        CDDX$ = CODEXT$(CIXI%)
        If CDDX$ >= cod1$ Then
          If CDDX$ <= cod2$ Then
            TMT1% = TIMATE%(CIXI%)
            If TMT% < 1 Or TMT1% = TMT% Then
              CII& = CIXI%
              SST$ = REGLEIDO$("SALSTOCK", CII& + 1) + Mid$(REGLEIDO$("SALDEPOS", CII& + 1), 41)
              '
              SSTT = 0
              For K% = 1 To 248 Step 4
                SSTT = SSTT + CVS(Mid$(SST$, K%, 4))
              Next K%
              '
              If Abs(SSTT) > 0.05 Then
                TTN$ = MKI$(CIXI%) + MKS$(SSTT) + String$(4, 0)
                J& = J& + 1
                Call GRAREG("NEGASTO", TTN$, J&)
              End If
            End If
          End If
        End If
      End If
    Next I&
    '
    Call SETULTREG("NEGASTO", J&)
    Call CIERRARCH("*.*")
    Screen.MousePointer = 1
    '
    If J& < 1 Then
        Call COMUNI("No se Detectaron Existencias Físicas.\Pulse cualquier tecla para volver al menu.")
      Else
        If CONSTOK.Option5.Value = True Then
            Call CONECRUN("*EXISTS", "")
          ElseIf CONSTOK.Option4.Value = True Then
            Call CONECRUN("*EXISTL", "")
        End If
    End If
    '
End Sub
'

Sub NEGASTO2()

3   ' NEGASTO.QBS                     DETERMINACION ITEMS CON STOCK NEGATIVO

    IUI% = -1
    For IU% = 0 To Forms.Count - 1
        XIU$ = UCase$(Forms(IU%).Name)
        If XIU$ = "COBASTO" Or XIU$ = "CONSTOK" Then
            IUI% = IU%
        End If
    Next IU%
    '
    J& = 0             'Call BLANARCH("NEGASTO")
    '
    Screen.MousePointer = 11
    NVII& = NUINV%
    SST$ = String$(256, 0)
    For I& = 1 To NVII&
        Call TRACE(20, I&, NVII&)
        II$ = REGLEIDO$("INVERT", I& + 1)
        CI% = CVI(Mid$(II$, 17, 2))
        If Mid$(II$, 16, 1) <> "`" Then
          If CI% > 0 Then
             If CI% <= NUMAS% Then
                CII& = CI%
                SST$ = REGLEIDO$("SALSTOCK", CII& + 1) + Mid$(REGLEIDO$("SALDEPOS", CII& + 1), 41)
                '
                SSTT = 0
                For K% = 1 To 248 Step 4
                    SSTT = SSTT + CVS(Mid$(SST$, K%, 4))
                Next K%
                '
                If SSTT < 0 Then
                    TTN$ = MKI$(CI%) + MKS$(SSTT) + String$(4, 0)
                    J& = J& + 1
                    Call GRAREG("NEGASTO", TTN$, J&)
                End If
             End If
          End If
        End If
     Next I&
     '
     Call SETULTREG("NEGASTO", J&)
     Call CIERRARCH("*.*")
     Screen.MousePointer = 1
     '
     If J& < 1 Then
         Call COMUNI("No se detectaron items con stock negativo.\Pulse cualquier tecla para volver al menu.")
       Else
         Call COMUNI("Se Detectaron " + TRIM$(Str$(J&)) + " Casos\de Stocks Negativos.")
     End If
     '
End Sub
'

Sub CRITSTO2()

3   ' CRITSTO.QBS                     DETERMINACION ITEMS CON STOCK NEGATIVO

    IUI% = 0
    For IU% = 0 To Forms.Count - 1
        XIU$ = UCase$(Forms(IU%).Name)
        If XIU$ = "COBASTO" Or XIU$ = "CONSTOK" Then
            IUI% = IU%
        End If
    Next IU%
    '
    J& = 0             'Call BLANARCH("NEGASTO")
    '
    Screen.MousePointer = 11
    NVII& = NUINV%
    For I& = 1 To NVII&
        Call TRACE(20, I&, NVII&)
        II$ = REGLEIDO$("INVERT", I& + 1)
        CI% = CVI(Mid$(II$, 17, 2))
        If Mid$(II$, 16, 1) <> "`" Then
           If CI% > 0 Then
              If CI% <= NUMAS% Then
                CII& = CI%
                SST$ = REGLEIDO$("SALSTOCK", CII& + 1) + Mid$(REGLEIDO$("SALDEPOS", CII& + 1), 41)
                '
                SSTT = 0
                For K% = 1 To 248 Step 4
                    SSTT = SSTT + CVS(Mid$(SST$, K%, 4))
                Next K%
                '
                STMIN = STOMIN(CI%)
                If SSTT < STMIN Then
                    '
                    COU = COSUNI(CI%)
                    REPOS = STMIN - SSTT
                    If REPOS < LOTESTAN(CI%) Then REPOS = LOTESTAN(CI%)
                    '
                    TTN$ = REGBLAN$("CRITSTO")
                    Call REPLA(TTN$, MKI$(CI%), 1, 2)
                    Call REPLA(TTN$, UMED$, 3, 2)
                    Call REPLA(TTN$, MKS$(SSTT), 5, 4)
                    Call REPLA(TTN$, MKS$(STMIN), 9, 4)
                    Call REPLA(TTN$, MKS$(STMIN - SSTT), 13, 4)
                    Call REPLA(TTN$, MKS$(REPOS), 17, 4)
                    Call REPLA(TTN$, MKS$(COU), 21, 4)
                    Call REPLA(TTN$, MKD$(COU * REPOS), 25, 8)
                    J& = J& + 1
                    Call GRAREG("CRITSTO", TTN$, J&)
                    '
                End If
              End If
           End If
        End If
     Next I&
     '
     Call SETULTREG("CRITSTO", J&)
     Call CIERRARCH("*.*")
     Screen.MousePointer = 1
     '
     If J& < 1 Then
         Call COMUNI("No se detectaron items con stock sub-mínimo.\Pulse cualquier tecla para volver al menu.")
       Else
         Call COMUNI("Se Detectaron " + TRIM$(Str$(J&)) + " Casos\de Stocks Sub-Mínimo.")
     End If
     '
End Sub

Sub STODISP()
    '
    HOII% = HOY(HOS$)
    IAT& = 0            ' INDICE ARCHIVO TRANSFERENCIA
    '
    VECOPEN$ = PUNOCPEN$
    '
    Call CARDEPOS
    '
    TMT% = Val(DISPOSTO.Text3.TEXT)
    TIMAT$ = ECOARCH$("TATIMAT", Chr$(TMT%))
    If TIMAT$ = "" Then
        TMT% = 0
    End If
    '
    CI1% = CODINT%(DISPOSTO.Text1.TEXT)
    cod1$ = CODEXT$(CI1%)
    If TRIM$(cod1$) = "" Then Exit Sub
    CI2% = CODINT%(DISPOSTO.Text2.TEXT)
    cod2$ = CODEXT$(CI2%)
    If TRIM$(cod2$) = "" Then Exit Sub
    '
    NIVEFIL% = 1
    If DISPOSTO.Option2.Value = True Then NIVEFIL% = 2
    If DISPOSTO.Option1.Value = True Then NIVEFIL% = 3
    If DISPOSTO.Option4.Value = True Then NIVEFIL% = 4
    '
    Screen.MousePointer = 11
    Dim EXIS(16384, 6)
    Dim PUNOF$(16384)
    Dim PROTER%(16384)      ' MARCA PRODUCTO TERMINADO
    Call CAROEMI((-32767), 0, 0) ' PARA INICIALIZAR
    '
    If DISPOSTO.Check2.Value = 1 Then
       '
       UREMAS& = ULTREG&("MASTER")
       USER% = USNBR% Mod 100
       If USER% < 1 Then USER% = 1
       '
       While ULTREG&("SALSTOCK") < UREMAS&
         X$ = String$(39, 0) + Chr$(1)
         Call REGAPP("SALSTOCK", X$)
       Wend
       '
       While ULTREG&("SALDEPOS") < UREMAS&
         Y$ = REGLEIDO$("SALSTOCK", 1 + ULTREG&("SALDEPOS"))
         X$ = String$(256, 0)
         Call REPLA(X$, Y$, 1, 40)
         Call REPLA(X$, MKI$(HOII%), 253, 2)
         Call REPLA(X$, MKI$(HORANUM%(Time$)), 255, 2)
         Call REGAPP("SALDEPOS", X$)
       Wend
       '
       Call CIERRARCH("SALSTOCK")   ' para liberar bloqueos
       Call CIERRARCH("SALDEPOS")   ' para liberar bloqueos
       '
       For U& = 1 To UREMAS&
         Call TRACE(20, U&, UREMAS&)
         X$ = String$(256, 0)
         Call REPLA(X$, MKI$(HOII%), 253, 2)
         Call REPLA(X$, MKI$(HORANUM%(Time$)), 255, 2)
         Call GRAREG("!SALDEPOS", X$, U&)
       Next U&
       '
       FIN& = ULTREG&("MOVISTO")
       '
       For I& = 1 To FIN&
         Call TRACE(20, I&, FIN&)
         MVS$ = REGLEIDO$("MOVISTO", I&)
         '
         CI% = CVI(Mid$(MVS$, 3, 2))
         DEP% = Asc(Mid$(MVS$, 83, 1))
         SUCX% = IDEPOS%(DEP%)
         CANTI = CVD(Mid$(MVS$, 88, 8)) - CVD(Mid$(MVS$, 96, 8))
         '
         If CI% > 0 Then
           If CI% <= NUMAS% Then
             RGG& = CI% + 1: PX% = 4 * SUCX% - 3
             If PX% > 0 Then
               If PX% <= 245 Then
                 STRI$ = REGLEIDO$("!SALDEPOS", RGG&)
                 VALO# = CDbl(CANTI)
                 Call RECUPD(STRI$, PX%, 4, VALO#)
                 Call GRAREG("!SALDEPOS", STRI$, RGG&)
               End If
             End If
           End If
         End If
       Next I&
       '
       For I& = 1 To UREMAS&
         Call TRACE(20, I&, UREMAS&)
         X$ = REGLEIDO$("!SALDEPOS", I&)
         Call GRAREG("SALSTOCK", Left$(X$, 40), I&)
         Call GRAREG("SALDEPOS", X$, I&)
       Next I&
       '
       Call CIERRARCH("SALSTOCK")
       Call CIERRARCH("SALDEPOS")
       '
    End If
    '
    FIN& = NUMAS%
    For I& = 1 To FIN&
       Call TRACE(20, I&, FIN&)
       SST$ = REGLEIDO$("SALSTOCK", I& + 1) + Mid$(REGLEIDO$("SALDEPOS", I& + 1), 41)
       SSTT = 0
       For KI% = 1 To NUSU%
          K% = 4 * KI% - 3
          If DISTO%(KI%) >= 0 Then
            SSTT = SSTT + CVS(Mid$(SST$, K%, 4))
          End If
       Next KI%
       EXIS(I&, 1) = SSTT
       CIIX% = I&
       Call CARTRAST(CIIX%, "ST", "Stock al " + HOS$ + " - " + Time$, "", SSTT, HOII%)
       EXIS(I&, 4) = STOMIN(CIIX%)
    Next I&
    '
    FIN& = ULTREG&("OCOMPRA")
    For U& = 1 To FIN&
       Call TRACE(20, U&, FIN&)
       X$ = REGLEIDO$("OCOMPRA", U&)
       STAT% = Asc(Mid$(X$, 62, 1))
       If STAT% <= 0 Then
          CATOT# = CVS(Mid$(X$, 13, 4))
          CAPRA# = CVS(Mid$(X$, 58, 4))
          CAPEN = CATOT# - CAPRA#
          If CAPEN > 0 Then
             CIIX% = CVI(Mid$(X$, 11, 2))
             NC% = CVI(Mid$(X$, 7, 2))
             COMPROB$ = "OC." + TRIM$(Str$(CVS(Left$(X$, 4))))
             FENTRESOL% = CVI(Mid$(X$, 9, 2))
             EXIS(CIIX%, 2) = EXIS(CIIX%, 2) + CAPEN
             Call CARTRAST(CIIX%, "OC", RASOCLI$((-1) * NC%), COMPROB$, CAPEN, FENTRESOL%)
          End If
       End If
    Next U&
    '
    FIN& = ULTREG&("OCOMDETA")
    For U& = 1 To FIN&
      Call TRACE(20, U&, FIN&)
      X$ = REGLEIDO$("OCOMDETA", U&)
      STAT% = Asc(Mid$(X$, 124, 1))
      If STAT% < 9 Then
        NUORSE& = CVS(Mid$(X$, 117, 4))
        If NUORSE& < 1 Then        ' no es orden de servicio
          COEFCAN = CVS(Mid$(X$, 63, 4))
          If COEFCAN < 0.005 Then COEFCAN = 1
          CATOT# = CVS(Mid$(X$, 55, 4)) * COEFCAN
          CAPRA# = CVS(Mid$(X$, 99, 4)) * COEFCAN
          CAPEN = CATOT# - CAPRA#
          If CAPEN > 0 Then
             CIIX% = CVI(Mid$(X$, 9, 2))
             NC% = CVI(Mid$(X$, 7, 2))
             COMPROB$ = "OC." + TRIM$(Str$(CVS(Left$(X$, 4))))
             FENTRESOL% = CVI(Mid$(X$, 95, 2))
             EXIS(CIIX%, 2) = EXIS(CIIX%, 2) + CAPEN
             Call CARTRAST(CIIX%, "OC", RASOCLI$((-1) * NC%), COMPROB$, CAPEN, FENTRESOL%)
          End If
        End If
      End If
    Next U&
    '
    FIN& = ULTREG&("BOLREPEN")
    For U& = 1 To FIN&
       Call TRACE(20, U&, FIN&)
       X$ = REGLEIDO$("BOLREPEN", U&)
       CAPEN = CVD(Mid$(X$, 89, 8))
       If CAPEN > 0 Then
          CIIX% = CVI(Mid$(X$, 83, 2))
          NC% = CVI(Mid$(X$, 105, 2))
          COMPROB$ = "BR." + TRIM$(Str$(CVS(Left$(X$, 4))))
          FENTRESOL% = CVI(Mid$(X$, 81, 2))
          EXIS(CIIX%, 2) = EXIS(CIIX%, 2) + CAPEN
          Call CARTRAST(CIIX%, "BR", "RT." + TRIM$(Mid$(X$, 119, 16)) + " - " + RASOCLI$((-1) * NC%), COMPROB$, CAPEN, FENTRESOL%)
       End If
    Next U&
    '
    FIN& = ULTREG&("ORDEFABR")
    For U& = 1 To FIN&
       Call TRACE(20, U&, FIN&)
       X$ = REGLEIDO$("ORDEFABR", U&)
       STAT% = Asc(Mid$(X$, 108, 1))
       If STAT% > 0 Then
          If STAT% < 3 Then
             CATOT# = CVD(Mid$(X$, 89, 8))
             CAPRA# = CVD(Mid$(X$, 119, 8))
             CAPEN = CATOT# - CAPRA#
             If CAPEN > 0 Then
                CIIX% = CVI(Mid$(X$, 83, 2))
                REFE$ = Mid$(X$, 5, 64)
                COMPROB$ = "OF." + TRIM$(Str$(CVS(Left$(X$, 4))))
                FENTRESOL% = CVI(Mid$(X$, 105, 2))
                EXIS(CIIX%, 3) = EXIS(CIIX%, 3) + CAPEN
                Call CARTRAST(CIIX%, "OF", REFE$, COMPROB$, CAPEN, FENTRESOL%)
                PUNOF$(CIIX%) = PUNOF$(CIIX%) + MKS$(U&)
             End If
          End If
       End If
    Next U&
    '
    APEDPE$ = "PEDDETA": MODOMAI% = 0
    If CONTROL$("ADMIPED", "FORMACAR") = "ITEMIZADO" Then
       MODOMAI% = 1: APEDPE$ = "PEDCLIEN"
    End If
    FIN& = ULTREG&(APEDPE$)
    '
    For U& = 1 To FIN&
       Call TRACE(20, U&, FIN&)
       X$ = REGLEIDO$(APEDPE$, U&)
       COMPROB$ = "PC." + TRIM$(Str$(CVS(Left$(X$, 4))))
       NC% = CVI(Mid$(X$, 7, 2))
       '
       CAPEN = 0
       If MODOMAI% <> 1 Then
            CIIX% = CVI(Mid$(X$, 33, 2))
            CAPEN = CVS(Mid$(X$, 79, 4)) - CVS(Mid$(X$, 95, 4))
            FENTRESOL% = CVI(Mid$(X$, 25, 2))
            If Asc(Mid$(X$, 27, 1)) > 1 Then CAPEN = 0 ' Cumplido o Anulado
          Else
            CIIX% = CVI(Mid$(X$, 640 + 3, 2))
            STATUS% = Asc(Mid$(X$, 300, 1))
            If STATUS% <= 1 Then
               CACOPED = CVS(Mid$(X$, 513, 4))
               CAITPED = CVS(Mid$(X$, 645, 4))
               CACOENT = CVS(Mid$(X$, 512 + 77, 4))
               CAPEN = (CACOPED - CACOENT) * CAITPED
               FENTRESOL% = CVI(Mid$(X$, 512 + 63, 2))
            End If
       End If
       '
       If CAPEN > 0 Then
          If Abs(EXIS(CIIX%, 0)) <> 1 Then
             EXIS(CIIX%, 0) = DESCOMPO%(CIIX%)
             If EXIS(CIIX%, 0) <> 1 Then EXIS(CIIX%, 0) = (-1)
          End If
          '
          EXIS(CIIX%, 5) = EXIS(CIIX%, 5) + CAPEN
          Call CARTRAST(CIIX%, "PC", RASOCLI$(NC%), COMPROB$, (-1) * CAPEN, FENTRESOL%)
          '
          If EXIS(CIIX%, 0) = 1 Then ' items descompuestos
               EXIS(CIIX%, 3) = EXIS(CIIX%, 3) + CAPEN
               Call LEEEXPLO(CODEXT$(CIIX%), 1)
               For KKI% = 1 To CAIT%
                  EXIS(CIJ%(KKI%), 6) = EXIS(CIJ%(KKI%), 6) + CAPEN * USOI(KKI%)
               Next KKI%
          End If
       End If
    Next U&
    '
    FIN& = ULTREG&("RESERVA")
    For I& = 1 To FIN&
       Call TRACE(20, I&, FIN&)
       XX$ = REGLEIDO$("RESERVA", I&)
       CIIX% = CVI(Left$(XX$, 2))
       TORE# = CVD(Mid$(XX$, 19, 8))
       CACO# = CVD(Mid$(XX$, 27, 8))
       '
       If UCase$(UNIMED$(CIIX%)) = "GK" Then
          TORE# = TORE# / 1000
          CACO# = CACO# / 1000
       End If
       '
       SIGNO% = Sgn(CVS(Mid$(XX$, 75, 4)))
       SARE = TORE# - CACO#
       If SARE * SIGNO% < 0 Then
          SARE = 0
       End If
       EXIS(CIIX%, 6) = EXIS(CIIX%, 6) + SARE
       If SARE >= 0.05 Then
          COMPROB$ = Mid$(XX$, 3, 12)
          REFE$ = "Reserva Fabric." + COMPROB$
          FENTRESOL% = CVI(Mid$(XX$, 17, 2))
          Call CARTRAST(CIIX%, "RF", REFE$, COMPROB$, (-1) * Abs(SARE), FENTRESOL%)
       End If
    Next I&
    '
    JJJ& = 0
    FIN& = ULTREG&("MACONSIG")
    For I& = 1 To FIN&
       Call TRACE(20, I&, FIN&)
       XX$ = REGLEIDO$("MACONSIG", I&)
       CIIX% = CVI(Mid$(XX$, 33, 2))
       TORE# = CVD(Mid$(XX$, 55, 8))
       CACO# = CVD(Mid$(XX$, 63, 8))
       '
       If CACO# < 0.005 Then
          NUOCO& = Val(Mid$(XX$, 24, 6))
          For TYU& = 1 To Len(VECOPEN$) Step 4
            If Abs(CVS(Mid$(VECOPEN$, TYU&, 4)) - NUOCO&) < 0.1 Then GoTo 42
          Next TYU&
          GoTo 49  ' FILTRA PARA O/C's PENDIENTES
       End If
       '
42     JJJ& = JJJ& + 1
       Call GRAREG("MACONSIG", XX$, JJJ&)
       '
       If UCase$(UNIMED$(CIIX%)) = "GK" Then
          TORE# = TORE# / 1000
          CACO# = CACO# / 1000
       End If
       '
       SARE = TORE# - CACO#
       If SARE < 0 Then
          SARE = 0
       End If
       If SARE >= 0.05 Then
          COMPROB$ = Mid$(XX$, 21, 12)
          NUOCO& = Val(Mid$(XX$, 24, 6))
          For TYU& = 1 To Len(VECOPEN$) Step 4
            If Abs(CVS(Mid$(VECOPEN$, TYU&, 4)) - NUOCO&) < 0.1 Then GoTo 45
          Next TYU&
          GoTo 49  ' FILTRA PARA O/C's PENDIENTES
          '
45        EXIS(CIIX%, 6) = EXIS(CIIX%, 6) + SARE
          REFE$ = "Consig." + RASOCLI$((-1) * CVI(Mid$(XX$, 3, 2)))
          FENTRESOL% = CVI(Mid$(XX$, 1, 2))
          Call CARTRAST(CIIX%, "RF", REFE$, COMPROB$, (-1) * Abs(SARE), FENTRESOL%)
       End If
49  Next I&
    Call SETULTREG("MACONSIG", JJJ&)
    Call CIERRARCH("MACONSIG")
    '
    FIN& = ULTREG&("STOREMOT")
    For I& = 1 To FIN&
       Call TRACE(20, I&, FIN&)
       XX$ = REGLEIDO$("STOREMOT", I&)
       CIIX% = CVI(Left$(XX$, 2))
       If CIIX% > 0 Then
         If CIIX% <= NUMAS% Then
           TIMO$ = UCase$(Mid$(XX$, 15, 2))
           TORE# = CVD(Mid$(XX$, 67, 8))
           CACO# = CVD(Mid$(XX$, 75, 8))
           SIGNO% = 1: COLUST% = 1
           SARE = TORE# - CACO#
           If Abs(SARE) >= 0.05 Then
             If TIMO$ = "OC" Then COLUST% = 2
             If TIMO$ = "BR" Then COLUST% = 2
             If TIMO$ = "OF" Then COLUST% = 3
             If TIMO$ = "PC" Then COLUST% = 5: SIGNO% = (-1)
             If TIMO$ = "RF" Then COLUST% = 6: SIGNO% = (-1)
             EXIS(CIIX%, COLUST%) = EXIS(CIIX%, COLUST%) + SARE * SIGNO%
           End If
         End If
       End If
    Next I&
    Call CIERRARCH("STOREMOT")
    '
    If NIVEFIL% = 4 Then GoTo 100 ' situacion productos terminados
    '
    J& = 0
    NVII& = NUINV%
    For I& = 1 To NVII&
        Call TRACE(20, I&, NVII&)
        II$ = REGLEIDO$("INVERT", I& + 1)
        CIIX% = CVI(Mid$(II$, 17, 2))
        If Mid$(II$, 16, 1) <> "`" Then
          If CIIX% > 0 Then
            If CIIX% <= NUMAS% Then
              CODE$ = CODEXT$(CIIX%)
              If CODE$ >= cod1$ Then
               If CODE$ <= cod2$ Then
                If TMT% = 0 Or TIMATE%(CIIX%) = TMT% Then
                  X$ = REGBLAN$("STODISP")
                  Call REPLA(X$, MKI$(CIIX%), 1, 2)
                  Call REPLA(X$, Chr$(TIMATE%(CIIX%)), 3, 1)
                  SUMA = 0: HAYMO% = 0
                  For K% = 1 To 6
                    SIG% = 1: If K% >= 4 Then SIG% = (-1)
                    SUMA = SUMA + SIG% * EXIS(CIIX%, K%)
                    If Abs(EXIS(CIIX%, K%)) >= 0.05 Then HAYMO% = 1
                    Call REPLA(X$, MKS$(EXIS(CIIX%, K%)), 4 * K% + 1, 4)
                  Next K%
                  If SUMA > 0.05 Then
                      Call REPLA(X$, MKS$(SUMA), 29, 4)
                    ElseIf SUMA < (-0.05) Then
                      Call REPLA(X$, MKS$((-1) * SUMA), 33, 4)
                      Call CAROEMI(CIIX%, Abs(SUMA), HOII%)
                  End If
                  '
                  If (NIVEFIL% = 1) Or (NIVEFIL% = 2 And HAYMO% = 1) Or (NIVEFIL% = 3 And SUMA <= (-0.05)) Then
                    Call REPLA(X$, MKS$(LOTESTAN(CIIX%)), 37, 4)
                    J& = J& + 1
                    Call GRAREG("STODISP", X$, J&)
                  End If
                End If
               End If
              End If
            End If
          End If
        End If
    Next I&
    '
    Call SETULTREG("STODISP", J&)
    Call CIERRARCH("STODISP")
    '
    Call SETULTREG("ATRASTOK", IAT&)
    XX$ = REGLEIDO$("ATRASTOK", 1)
    Call REPLA(XX$, HOS$, 17, 8)
    Call REPLA(XX$, Time$, 25, 8)
    Call GRAREG("ATRASTOK", XX$, 1)
    Call CIERRARCH("ATRASTOK")
    '
    Call CIERRARCH("*.*")
    '
    Call CARUTRA
    AAI% = 1
    If TRIM$(IDENSEDE$) <> "" Then
       If TRIM$(DESTITRA$) <> "" Then
          AAI% = COPYFILE%(LUDAT$ + "ATRASTOK.DAT", DESTITRA$)
       End If
    End If
    Screen.MousePointer = 1
    If AAI% <> 1 Then Exit Sub
    '
    If DISPOSTO.Option11.Value = True Then
         Call CONECRUN("*STODISC", "")
       ElseIf DISPOSTO.Option12.Value = True Then
         Call CONECRUN("*STODISL", "")
    End If
    Exit Sub
    '
100 FIN& = ULTREG&("ESTRUNUE")
    FINDOS& = ULTREG&("ESTRUDOS")
    For REG& = 1 To FIN&
       Call TRACE(20, REG&, FIN&)
       X$ = REGLEIDO$("ESTRUNUE", REG&)
       CIIX% = CVI(Left$(X$, 2))
       '
       If CIIX% > 0 Then
          If CIIX% <= NUMAS% Then
             If PROTER%(CIIX%) = 0 Then PROTER%(CIIX%) = 1
             '
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
                      If PROTER%(CIIX%) = 0 Then PROTER%(CIIX%) = 1
                      PROTER%(CIKK%) = (-1)
                    End If
                  End If
                End If
             Next K%
          End If
       End If
       '
    Next REG&
    '
    J& = 0
    NVII& = NUINV%
    For I& = 1 To NVII&
        Call TRACE(20, I&, NVII&)
        II$ = REGLEIDO$("INVERT", I& + 1)
        CIIX% = CVI(Mid$(II$, 17, 2))
        If Mid$(II$, 16, 1) <> "`" Then
          If CIIX% > 0 Then
            If CIIX% <= NUMAS% Then
              If PROTER%(CIIX%) = 1 Then
                X0$ = REGBLAN$("SIPROTER")
                Call REPLA(X0$, MKI$(CIIX%), 1, 2)
                Call REPLA(X0$, MKS$(EXIS(CIIX%, 1)), 3, 4)
                If Len(PUNOF$(CIIX%)) < 4 Then
                     J& = J& + 1
                     Call GRAREG("SIPROTER", X0$, J&)
                   Else
                     For KK% = 1 To Len(PUNOF$(CIIX%)) Step 4
                       U& = CVS(Mid$(PUNOF$(CIIX%), KK%, 4))
                       X$ = REGLEIDO$("ORDEFABR", U&)
                       CATOF = CVD(Mid$(X$, 89, 8))
                       CAPEN = CATOF - CVD(Mid$(X$, 119, 8))
                       If CAPEN < 0 Then CAPEN = 0
                       COMPROB$ = "OF." + TRIM$(Str$(CVS(Left$(X$, 4))))
                       X1$ = X0$
                       Call REPLA(X1$, COMPROB$, 7, 12)
                       Call REPLA(X1$, MKS$(CATOF), 19, 4)
                       Call REPLA(X1$, MKS$(CAPEN), 23, 4)
                       J& = J& + 1
                       Call GRAREG("SIPROTER", X1$, J&)
                     Next KK%
                End If
              End If
            End If
          End If
        End If
    Next I&
    '
    Call SETULTREG("SIPROTER", J&)
    Call CIERRARCH("SIPROTER")
    '
    Call CONECRUN("*SIPROTE", "Situacion de Stocks Terminados")
    '
End Sub
'

Sub CARTRAST(CIIY%, COMO$, REMO$, NUCO$, CATRA, FEPRO%)
   '
   If IAT& < 1 Then
      If RECLEN%("ATRASTOK") < 256 Then
         Call COMUNI("Error en Controlador de Archivo 'ATRASTOK.RFS'.\  \Consulte a su Soporte de Sistemas.")
         Call FINAL("")
      End If
      HOII% = HOY(HOS$)
      Call CARUTRA
      XX$ = Space$(128) + String$(128, 0)
      Call REPLA(XX$, IDENSEDE$, 1, 8)
      Call REPLA(XX$, "HEADER", 9, 8)
      Call REPLA(XX$, HOS$, 17, 8)
      Call REPLA(XX$, Time$, 25, 8)
      Call REPLA(XX$, USERNAME$, 33, 32)
      Call REPLA(XX$, NOMTER$, 65, 32)
      Call REPLA(XX$, Space$(32), 97, 32)
      Call GRAREG("ATRASTOK", XX$, 1)
      IAT& = 1
   End If
   '
   If Abs(CATRA) >= 0.05 Then
      '
      XX$ = REGBLAN$("ATRASTOK")
      Call REPLA(XX$, IDENSEDE$, 1, 8)
      Call REPLA(XX$, "DISPOREM", 9, 8)
      Call REPLA(XX$, CODEXT$(CIIY%), 17, 16)
      Call REPLA(XX$, COMO$, 33, 2)
      Call REPLA(XX$, REMO$, 35, 50)
      Call REPLA(XX$, NUCO$, 85, 12)
      If CATRA > 0 Then
           Call REPLA(XX$, MKD$(CDbl(CATRA)), 97, 8)
         Else
           Call REPLA(XX$, MKD$(CDbl(Abs(CATRA))), 105, 8)
      End If
      Call REPLA(XX$, MKI$(FEPRO%), 113, 2)
      '
      IAT& = IAT& + 1
      Call GRAREG("ATRASTOK", XX$, IAT&)
      '
   End If
   '
End Sub
'

Sub CAROEMI(CIII%, CANOR, FEOR%)
   '
   Static CTX%, JCOM&, JFAB&
   Static ACON%(32767)
   '
   If CTX% = 0 Or CIII% = (-32767) Then
     FIN& = ULTREG&("ESTRUNUE")
     For REG& = 1 To FIN&
       Call TRACE(20, REG&, FIN&)
       X$ = REGLEIDO$("ESTRUNUE", REG&)
       CIIQ% = CVI(Left$(X$, 2))
       If CIIQ% > 0 Then
         If CIIQ% <= 32767 Then
           If ACON%(CIIQ%) = 0 Then
             DPSX$ = Mid$(X$, 3, 8)
             CIKK% = CVI(Left$(DPSX$, 2))
             If CIKK% > 0 Then
               If CIKK% <= 32767 Then
                 If Abs(CVS(Mid$(DPSX$, 3))) >= 0.000005 Then
                   ACON%(CIIQ%) = 1
                 End If
               End If
             End If
           End If
         End If
       End If
     Next REG&
     CTX% = 1
     JCON& = 0: Call SETULTREG("OCOMEMI", JCON&)
     JFAB& = 0: Call SETULTREG("ORFAEMI", JFAB&)
   End If
   '
   If CIII% < 1 Then Exit Sub
   '
   FF1$ = REGBLAN$("ORFAEMI")
   Call REPLA(FF1$, MKI$(CIII%), 1, 2)
   Call REPLA(FF1$, DESCRIT$(CIII%), 3, 26)
   Call REPLA(FF1$, CSTRING$(MKS$(CANOR), 4, 10, 1, 2), 29, 10)
   FEX = FEOR%
   Call REPLA(FF1$, FECHATEX$(FEX), 41, 8)
   Call REPLA(FF1$, MKD$(CDbl(CANOR)), 49, 8)
   Call REPLA(FF1$, MKI$(FEOR%), 57, 2)
   Call REPLA(FF1$, MKI$(CIII%), 59, 2)
   '
   If ACON(CIII%) = 1 Then
         JFAB& = JFAB& + 1
         Call GRAREG("ORFAEMI", FF1$, JFAB&)
      Else
         JCOM& = JCOM& + 1
         Call GRAREG("OCOMEMI", FF1$, JCOM&)
   End If
   '
End Sub
