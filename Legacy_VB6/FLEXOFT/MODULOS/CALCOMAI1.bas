Attribute VB_Name = "CALCOMAI1"
Sub CALCOMCHE()
    '
    Dim CHETEX$(4096), PUFAC$(4096), TAPLI#(4096), TOFAC#(4096), TOBAS#(4096), SUAPLI#(4096)
    Dim BABAN$(32767)
    Dim TIS$(7)
    '
    Dim CVX%(256), COVEN(256), COCOB(256)
    Dim PORCO(256), TOCOMI#(256), TOCOB#(256)
    Dim SUCOB#(256), COMCOB#(256)
    Dim SUVENE#(256), SUVABA#(256)
    '
' carga vector de tipos de movimiento
    TDAT$ = "AS,VC,FC,ND,RB,DC,NC,AS"
    For I& = 0 To 7
        TIS$(I&) = Mid$(TDAT$, 3 * I& + 1, 2)
    Next I&
    '
'carga vector de vendedores
    For U& = 1 To 256
       CVX%(U&) = 0
       SUVENE#(U&) = 0
       SUVABA#(U&) = 0
       SUCOB#(U&) = 0
       COMCOB#(U&) = 0
    Next U&
    '
    NUVE% = 0: YAGERE% = 0: PUNVENDE$ = ""
    For I& = 1 To ULTREG&("VENDEDOR")
        X$ = REGLEIDO$("VENDEDOR", I&)
        CV% = Asc(Left$(X$, 1))
        If CV% > 0 Then
          If CV% < 256 Then
            If CVX%(CV%) = 0 Then
              NUVE% = NUVE% + 1
              CVX%(CV%) = CV%
              COVEN(CV%) = CVS(Mid$(X$, 44, 4))
              COCOB(CV%) = CVS(Mid$(X$, 48, 4))
              If CV% = 99 Then YAGERE% = 1
              PUNVENDE$ = PUNVENDE$ + Chr$(CV%)
            End If
          End If
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
        CVX%(99) = 99
        COVEN(99) = CVS(Mid$(X$, 44, 4))
        COCOB(99) = CVS(Mid$(X$, 48, 4))
        PUNVENDE$ = PUNVENDE$ + Chr$(99)
    End If
    Call CIERRARCH("VENDEDOR")
    '
' carga vector de bancos
    For I& = 1 To ULTREG&("MECOBRA")
      F0$ = REGLEIDO$("MECOBRA", I&)
      COMECO% = CVI(Left$(F0$, 2))
      If COMECO% > 0 Then
        If COMECO% <= 32767 Then
          If BABAN$(COMECO%) = "" Then
            BABAN$(COMECO%) = TRIM$(Mid$(F0$, 15, 16))
            If BABAN$(COMECO%) = "" Then BABAN$(COMECO%) = TRIM$(Mid$(F0$, 3, 12))
          End If
        End If
      End If
    Next I&
    '
' selecciona mes y año de cálculo
    MMXX0$ = TRIM$(FORLICOMAI.Label3.Caption)
    MMXX1$ = MESANO$(MMXX0$, DES%, HAS%)
    If TRIM$(MMXX1$) = "" Then
        FORSELMEC.Show 1
        Exit Sub
    End If
    CORRESP$ = MMXX1$
    '
' aparea archivo de valorizacion de facturas
    FIN& = ULTREG&("FACTUR")
    UREVA& = ULTREG&("VAFACTUR")
    For U& = 1 To FIN&
       Call TRACE(20, U&, FIN&)
       X$ = REGLEIDO$("FACTUR", U&)
       If U& <= UREVA& Then
         Y$ = REGLEIDO$("VAFACTUR", U&)
         If Mid$(Y$, 5, 64) = X$ Then
           If CVI(Mid$(Y$, 115, 2)) > 0 Then
             If Abs(CVD(Mid$(Y$, 99, 8)) + CVD(Mid$(Y$, 107, 8))) >= 0.005 Then
                VENDE% = CVI(Mid$(Y$, 97, 2))
                If VENDE% > 0 Then
                  If VENDE% < 256 Then
                    If InStr(PUNVENDE$, Chr$(VENDE%)) > 0 Then
                      '
                      NC% = CVI(Mid$(X$, 1, 2))
                      NRO& = CVS(Mid$(X$, 3, 4))
                      TI% = Asc(Mid$(X$, 13, 1))
                      If TI% < 0 Or TI% > 7 Then TI% = 2
                      TDOC$ = AJUSTI$(TIS$(TI%) + "." + TRIM$(Str$(NRO&)), 10)
                      KDOC2$ = MKI$(NC%) + TDOC$
                      IDOC2% = 0: For K% = 1 To 12
                                   IDOC2% = IDOC2% + K% * Asc(Mid$(KDOC2$, K%, 1))
                                 Next K%
                      Call REPLA(Y$, MKI$(IDOC2%), 3, 2)
                      Call GRAREG("VAFACTUR", Y$, U&)
                      GoTo 19
                      '
                    End If
                  End If
                End If
             End If
           End If
         End If
       End If
       '
       NC% = CVI(Mid$(X$, 1, 2))
       NRO& = CVS(Mid$(X$, 3, 4))
       FF% = CVI(Mid$(X$, 7, 2))
       TI% = Asc(Mid$(X$, 13, 1))
       VA% = Asc(Mid$(X$, 14, 1))
       If TI% < 0 Or TI% > 7 Then TI% = 2
       TDOC$ = AJUSTI$(TIS$(TI%) + "." + TRIM$(Str$(NRO&)), 10)
       KDOC$ = MKI$(FF%) + MKI$(NC%) + TDOC$
       IDOC% = 0: For K% = 1 To 14
                    IDOC% = IDOC% + K% * Asc(Mid$(KDOC$, K%, 1))
                  Next K%
       KDOC2$ = MKI$(NC%) + TDOC$
       IDOC2% = 0: For K% = 1 To 12
                    IDOC2% = IDOC2% + K% * Asc(Mid$(KDOC2$, K%, 1))
                  Next K%
                  '
       VENDE% = 0
       If NRO& > 0 Then
          If NC% > 0 Then
             VENDE% = VENDCLI%(NC%)
          End If
       End If
       If VENDE% < 1 Then VENDE% = 99
       '
       Y$ = REGBLAN$("VAFACTUR")
       Call REPLA(Y$, MKI$(IDOC%), 1, 2)
       Call REPLA(Y$, MKI$(IDOC2%), 3, 2)
       Call REPLA(Y$, X$, 5, 64)
       Call REPLA(Y$, String$(30, 0), 69, 28)
       Call REPLA(Y$, MKI$(VENDE%), 97, 2)
       Call REPLA(Y$, MKD$(0), 99, 8)   ' VALOR BASE CALCULADO
       Call REPLA(Y$, MKD$(0), 107, 8)  ' VALOR BASE A PROMEDIO
       Call REPLA(Y$, MKI$(0), 115, 2)
       '
       Call GRAREG("VAFACTUR", Y$, U&)
       VEDOCPEN$ = VEDOCPEN$ + TDOC$
       '
19  Next U&
    Call CIERRARCH("VAFACTUR")
    UREVA& = ULTREG&("VAFACTUR")
    '
' valoriza facturas a precio base
    FIN& = ULTREG&("MOVISTO")
    For IL& = 1 To FIN&
        Call TRACE(20, IL&, FIN&)
        XMV$ = REGLEIDO$("MOVISTO", IL&)
        TDOC$ = Mid$(XMV$, 33, 10)
        If InStr(VEDOCPEN$, TDOC$) > 0 Then
           FF% = CVI(Left$(XMV$, 2))
           NC% = CVI(Mid$(XMV$, 84, 2))
           NRO& = Val(Mid$(TDOC$, 4))
           KDOC$ = MKI$(FF%) + MKI$(NC%) + TDOC$
           IDOC% = 0: For K% = 1 To 14
                        IDOC% = IDOC% + K% * Asc(Mid$(KDOC$, K%, 1))
                      Next K%
           RFF$ = RECFILT$("VAFACTUR", 1, MKI$(IDOC%))
           CAPAR# = CVD(Mid$(XMV$, 96, 8)) - CVD(Mid$(XMV$, 88, 8))
           CI% = CVI(Mid$(XMV$, 3, 2))
           VALOBA# = CAPAR# * PREBASE(CI%)
           '
           For K% = 1 To Len(RFF$) Step 4
             RENOTA& = CVS(Mid$(RFF$, K%, 4))
             If RENOTA& > 0 Then
               If RENOTA& <= UREVA& Then
                 Y$ = REGLEIDO$("VAFACTUR", RENOTA&)
                 If CVI(Mid$(Y$, 5, 2)) = NC% Then
                   If CVS(Mid$(Y$, 7, 4)) = NRO& Then
                     If CVI(Mid$(Y$, 11, 2)) = FF% Then
                       If CVI(Mid$(Y$, 115, 2)) < 1 Then
                         VALOBA# = VALOBA# + CVD(Mid$(Y$, 99, 8))
                         Call REPLA(Y$, MKD$(VALOBA#), 99, 8)
                         Call GRAREG("VAFACTUR", Y$, RENOTA&)
                       End If
                       GoTo 29
                     End If
                   End If
                 End If
               End If
             End If
           Next K%
        End If
29  Next IL&
    Call CIERRARCH("VAFACTUR")
    '
    FIN& = ULTREG&("VAFACTUR")
    For U& = 1 To FIN&
      Call TRACE(20, U&, FIN&)
      Y$ = REGLEIDO$("VAFACTUR", U&)
      If Abs(CVD(Mid$(Y$, 99, 8))) > 0.005 Then
         VENDE% = CVI(Mid$(Y$, 97, 2))
         VANEFAC# = CVD(Mid$(Y$, 19, 8))
         PREBAFA# = CVD(Mid$(Y$, 99, 8))
         SIG% = 1: If Asc(Mid$(Y$, 17, 1)) > 3 Then SIG% = (-1)
         If VENDE% > 0 And VENDE% < 256 Then
            SUVENE#(VENDE%) = SUVENE#(VENDE%) + VANEFAC# * SIG%
            SUVABA#(VENDE%) = SUVABA#(VENDE%) + PREBAFA# * SIG%
            GoTo 39
         End If
      End If
39  Next U&
    '
    For U& = 1 To FIN&
      Call TRACE(20, U&, FIN&)
      Y$ = REGLEIDO$("VAFACTUR", U&)
      If Abs(CVD(Mid$(Y$, 99, 8))) < 0.005 Then
        If Abs(CVD(Mid$(Y$, 107, 8))) < 0.005 Then
          VANEFAC# = CVD(Mid$(Y$, 19, 8))
          VENDE% = CVI(Mid$(Y$, 97, 2))
          If VENDE% > 0 And VENDE% < 256 Then
             If Abs(SUVENE#(VENDE%)) >= 1 Then
                PREBAFA# = VANEFAC# * SUVABA#(VENDE%) / SUVENE#(VENDE%)
                Call REPLA(Y$, MKD$(PREBAFA#), 107, 8)
                Call REPLA(Y$, MKI$(1), 115, 1)
                Call GRAREG("VAFACTUR", Y$, U&)
              End If
          End If
        End If
      End If
49  Next U&
    '
Call CIERRARCH("VAFACTUR")

' carga vector de cheques acreditados
    CACHE% = 0: PUNRECI$ = ""
    FIN& = ULTREG&("CHECAR")
    For U& = 1 To FIN&
       Call TRACE(20, U&, FIN&)
       X$ = REGLEIDO$("CHECAR", U&)
       FEACRECHE% = CVI(Mid$(X$, 41, 2))
       If FEACRECHE% >= DES% Then
          If FEACRECHE% <= HAS% Then
             COMECO% = CVI(Left$(X$, 2))
             NC% = CVI(Mid$(X$, 3, 2))
             '
             COEFIVA = (100 + ALIVA) / 100
             If PIVACLI%(NC%) = 2 Then
                  COEFIVA = (100 + ALIVA + ALIVARNI) / 100
                ElseIf PIVACLI%(NC%) = 5 Then
                  COEFIVA = 1
             End If
             '
             NUCHE$ = Mid$(X$, 5, 12)
             BACHE$ = TRIM$(Mid$(X$, 17, 16))
             If BACHE$ = "" Then BACHE$ = BABAN$(COMECO%)
             BANU$ = TRIM$(BACHE$) + " - " + NUCHE$
             IMCHE# = CVD(Mid$(X$, 33, 8))
             FEX = CVI(Mid$(X$, 41, 2)): FEACRE$ = FECHATEX$(FEX)
             FEX = CVI(Mid$(X$, 43, 2)): FEMISI$ = FECHATEX$(FEX)
             FEX = CVI(Mid$(X$, 45, 2)): FECOBR$ = FECHATEX$(FEX)
             NURECI& = CVS(Mid$(X$, 47, 4))
             '
             TTXX$ = Space$(256)
             Call REPLA(TTXX$, BANU$, 1, 32)
             Call REPLA(TTXX$, CSTRING$(MKI$(NC%), 1, 6, 0, 2), 33, 6)
             Call REPLA(TTXX$, CSTRING$(MKD$(IMCHE# / COEFIVA), 3, 12, 2, 2), 39, 12)
             Call REPLA(TTXX$, FEMISI$, 53, 8)
             Call REPLA(TTXX$, FEACRE$, 63, 8)
             Call REPLA(TTXX$, FECOBR$, 73, 8)
             Call REPLA(TTXX$, CSTRING$(MKS$(NURECI&), 3, 8, 0, 2), 83, 8)
             '
             CACHE% = CACHE% + 1
             CHETEX$(CACHE%) = TTXX$
             '
             PRB$ = "RB." + AJUSTI$(Str$(NURECI&), 7)
             If InStr(PUNRECI$, PRB$) < 1 Then
                PUNRECI$ = PUNRECI$ + PRB$
             End If
             '
          End If
       End If
    Next U&
    '
    FIN& = ULTREG&("APLICOB")
    For U& = 1 To FIN&
       Call TRACE(20, U&, FIN&)
       X$ = REGLEIDO$("APLICOB", U&)
       NURECIBO$ = Mid$(X$, 5, 18)
       If Left$(NURECIBO$, 2) = "RB" Then
         PPXX% = InStr(6, NURECIBO$, "-")
         If PPXX% > 0 Then
           NURE& = Val(Mid$(NURECIBO$, PPXX% + 1, 8))
           PRB$ = "RB." + AJUSTI$(Str$(NURE&), 7)
           If InStr(PUNRECI$, PRB$) > 0 Then
             NC% = CVI(Mid$(X$, 3, 2))
             DOCUORIG$ = Mid$(X$, 55, 18)
             IAPLI$ = CSTRING$(Mid$(X$, 23, 8), 4, 14, 2, 2)
             For K% = 1 To CACHE%
               If Val(Mid$(CHETEX$(K%), 33, 6)) = NC% Then
                 If Val(Mid$(CHETEX$(K%), 83, 8)) = NURE& Then
                   PUFAC$(K%) = PUFAC$(K%) + DOCUORIG$ + IAPLI$
                   TAPLI#(K%) = TAPLI#(K%) + CVD(Mid$(X$, 23, 8))
                 End If
               End If
             Next K%
           End If
         End If
       End If
    Next U&
    '
    FIN& = CACHE%
    UREVA& = ULTREG&("VAFACTUR")
    For U& = 1 To FIN&
       Call TRACE(20, U&, FIN&)
       NC% = Val(Mid$(CHETEX$(U&), 33, 6))
       For K% = 1 To Len(PUFAC$(U&)) Step 32
         NUFA& = Val(Mid$(PUFAC$(U&), K% + 8, 8))
         TDOC$ = AJUSTI$(Mid$(PUFAC$(U&), K%, 2) + "." + TRIM$(Str$(NUFA&)), 10)
         KDOC2$ = MKI$(NC%) + TDOC$
         IDOC2% = 0: For K1% = 1 To 12
                      IDOC2% = IDOC2% + K1% * Asc(Mid$(KDOC2$, K1%, 1))
                    Next K1%
         RFF$ = RECFILT$("VAFACTUR", 2, MKI$(IDOC2%))
         For K1% = 1 To Len(RFF$) Step 4
           RENOTA& = CVS(Mid$(RFF$, K1%, 4))
           If RENOTA& > 0 Then
             If RENOTA& <= UREVA& Then
               Y$ = REGLEIDO$("VAFACTUR", RENOTA&)
                 If CVS(Mid$(Y$, 7, 4)) = NUFA& Then
                   TOFAC#(U&) = TOFAC#(U&) + CVD(Mid$(Y$, 19, 8))
                   TOBAS#(U&) = TOBAS#(U&) + CVD(Mid$(Y$, 99, 8)) + CVD(Mid$(Y$, 107, 8))
                 End If
             End If
           End If
         Next K1%
         SUAPLI#(U&) = SUAPLI#(U&) + Val(Mid$(PUFAC$(U&), 18 + K%, 14))
       Next K%
       VACHE# = Val(Mid$(CHETEX$(U&), 39, 12))
       COEF = 1: If SUAPLI#(U&) > 0.05 Then COEF = VACHE# / SUAPLI#(U&)
       For K% = 1 To Len(PUFAC$(U&)) Step 32
         VAPLI# = Val(Mid$(PUFAC$(U&), 18 + K%, 14)) * COEF
         VPXX$ = CSTRING$(MKD$(VAPLI#), 4, 14, 2, 2)
         Mid$(PUFAC$(U&), 18 + K%, 14) = VPXX$
       Next K%
    Next U&
    '
    JJ& = 0
    FIN& = CACHE%
    For U& = 1 To FIN&
       Call TRACE(20, U&, FIN&)
       VACHE# = Val(Mid$(CHETEX$(U&), 39, 12))
       NC% = Val(Mid$(CHETEX$(U&), 33, 6))
       FERECIBO% = FECHANUM(Mid$(CHETEX$(U&), 73, 8))
       NURECIBO = Val(Mid$(CHETEX$(U&), 83, 8))
       CHENUBAN$ = Left$(CHETEX$(U&), 32)
       FECHEQUE% = FECHANUM(Mid$(CHETEX$(U&), 63, 8))
       ' TTXX$ = Left$(CHETEX$(U&), 90)
       TTXX$ = REGBLAN$("LICOMCHE")
       VENDE% = VENDCLI%(NC%) Mod 256
       Call REPLA(TTXX$, Chr$(VENDE%), 1, 1)
       Call REPLA(TTXX$, MKI$(NC%), 3, 2)
       Call REPLA(TTXX$, MKI$(FERECIBO%), 5, 2)
       Call REPLA(TTXX$, MKS$(NURECIBO), 7, 4)
       Call REPLA(TTXX$, CHENUBAN$, 11, 32)
       Call REPLA(TTXX$, MKI$(FECHEQUE%), 43, 2)
       TTX0$ = TTXX$
       '
       PUFA$ = Left$(PUFAC$(U&), 32)
       If TRIM$(PUFA$) = "" Then
          PUFA$ = "Anticipo no Imput.  " + Mid$(CHETEX$(U&), 39, 12)
       End If
       Call REPLA(TTXX$, AJUSTI$(PUFA$, 32), 45, 32)
       '
       If Abs(TOFAC#(U&)) < 0.005 Then
          TOFAC#(U&) = VACHE#
       End If
       '
       If Abs(TOBAS#(U&)) < 0.005 Then
          If VENDE% > 0 And VENDE% < 256 Then
             If Abs(SUVENE#(VENDE%)) >= 1 Then
                TOBAS#(U&) = TOFAC#(U&) * SUVABA#(VENDE%) / SUVENE#(VENDE%)
             End If
          End If
       End If
       '
59     If TOFAC#(U&) > 0.005 Then
          Call REPLA(TTXX$, MKD$(TOFAC#(U&) * VACHE# / TOFAC#(U&)), 77, 8)
          Call REPLA(TTXX$, MKD$(TOBAS#(U&) * VACHE# / TOFAC#(U&)), 85, 8)
          Call REPLA(TTXX$, MKD$((TOFAC#(U&) - TOBAS#(U&)) * VACHE# / TOFAC#(U&)), 93, 8)
       End If
       '
       PORCOM = 0
       If VENDE% > 0 And VENDE% < 256 Then
          PORCOM = COCOB(VENDE%)
       End If
       '
69     Call REPLA(TTXX$, MKS$(PORCOM), 101, 4)
       Call REPLA(TTXX$, MKD$(PORCOM * CVD(Mid$(TTXX$, 93, 8)) / 100), 105, 8)
       '
       JJ& = JJ& + 1
       Call GRAREG("LICOMCHE", TTXX$, JJ&)
       '
       If VENDE% > 0 And VENDE% < 256 Then
          SUCOB#(VENDE%) = SUCOB#(VENDE%) + CVD(Mid$(TTXX$, 93, 8))
          COMCOB#(VENDE%) = COMCOB#(VENDE%) + PORCOM * CVD(Mid$(TTXX$, 93, 8)) / 100
       End If
       '
       For K% = 33 To Len(PUFAC$(U&)) Step 32
          TTXX$ = TTX0$
          Call REPLA(TTXX$, Mid$(PUFAC$(U&), K%, 32), 45, 32)
          JJ& = JJ& + 1
          Call GRAREG("LICOMCHE", TTXX$, JJ&)
       Next K%
       '
    Next U&
    Call SETULTREG("LICOMCHE", JJ&)
    Call CIERRARCH("LICOMCHE")
    '
    FORLICOMAI.List1.Clear
    For VENDE% = 0 To 255
       If Abs(COMCOB#(VENDE%)) >= 0.005 Then
          If VENDE% = 0 Then
              NOVEN$ = "DIRECCION"
            Else
              NOVEN$ = ECOARCH$("VENDEDOR", Chr$(VENDE%))
          End If
          TXT$ = Space$(48)
          Call REPLA(TXT$, CSTRING$(MKI$(VENDE%), 1, 3, 0, 2), 1, 4)
          Call REPLA(TXT$, NOVEN$, 6, 17)
          Call REPLA(TXT$, CSTRING$(MKD$(COMCOB#(VENDE%)), 4, 10, 2, 2), 23, 10)
          FORLICOMAI.List1.AddItem TXT$
       End If
    Next VENDE%
    '
    If FORLICOMAI.List1.ListCount > 0 Then
       FORLICOMAI.List1.ListIndex = 0
    End If
    FORLICOMAI.List1.Refresh
    FORLICOMAI.Check1.Refresh
    FORLICOMAI.Check2.Refresh
    FORLICOMAI.Check3.Refresh
    FORLICOMAI.Check4.Refresh
    FORLICOMAI.MARCOBARRA.Refresh
    FORLICOMAI.Frame3.Refresh
    FORLICOMAI.Refresh
    '
    HOII% = HOY(HOS$)
    FORLICOMAI.Label4.Caption = HOS$
    FORLICOMAI.Label6.Caption = Left$(Time$, 5)
    '
    FORLICOMAI.List1.Refresh
    FORLICOMAI.Command3.Enabled = True
    FORLICOMAI.Hide
    FORLICOMAI.Show
    Screen.MousePointer = 1
    '
    Call CIERRARCH("*.*")
    '
End Sub
'
Sub CALCOMVEN()
    '
    Screen.MousePointer = 11
    '
    YALIQUI% = 0
    Dim CVX%(256), COVEN(256), COCOB(256)
    Dim PORCO(256), TOCOMI#(256), TOCOB#(256)
    Dim SUPED#(256), COMPED#(256), SUCOB#(256), COMCOB#(256)
    '
    Dim NFAC$(32767)
    Dim REGFA$(32767)
    '
    Dim TIS$(7)
    '
    TDAT$ = "AS,VC,FC,ND,RB,DC,NC,AS"
    For I& = 0 To 7
        TIS$(I&) = Mid$(TDAT$, 3 * I& + 1, 2)
    Next I&
    '
    For U& = 1 To 256
       SUPED#(U&) = 0
       COMPED#(U&) = 0
       SUCOB#(U&) = 0
       COMCOB#(U&) = 0
    Next U&
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
50  MMXX0$ = TRIM$(FORLICOMAI.Label3.Caption)
    MMXX1$ = MESANO$(MMXX0$, DES%, HAS%)
    If TRIM$(MMXX1$) = "" Then
        FORSELMEC.Show 1
        Exit Sub
    End If
    CORRESP$ = MMXX1$
    '
    Call HEADERS("LICOMIX", "")
    Call HEADERS("LICOMIX", "Corresponde a: " + CORRESP$)
    '
200 Call HEADERS("LICOMIX", "Por: Facturacion")
    FIN& = ULTREG&("FACTUR")
    '
    JJ& = 0
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
                       VENDE% = VENDCLI%(NC%)
                       If VENDE% < 1 Then VENDE% = 99
                       '
                       TI% = Asc(Mid$(X$, 13, 1))
                       If TI% < 1 Or TI% > 7 Then GoTo 209
                       '
                       TDOC$ = AJUSTI$(TIS$(TI%) + "." + TRIM$(Str$(NRO&)), 10)
                       JJ& = JJ& + 1
                       NFAC$(JJ&) = TDOC$
                       REGFA$(JJ&) = MKS$(I&)
                    End If
                End If
            End If
        End If
209 Next I&
    CAFAC% = JJ&
    '
    FIN& = ULTREG&("MOVISTO")
    For IL& = 1 To FIN&
        Call TRACE(20, IL&, FIN&)
        XMV$ = REGLEIDO$("MOVISTO", IL&)
        FEMO% = CVI(Left$(XMV$, 2))
        If FEMO% = 0 Then GoTo 214
        If FEMO% >= DES% Then
          If FEMO% <= HAS% Then
214         COMO$ = Mid$(XMV$, 33, 10)
            For KI% = 1 To CAFAC%
              If NFAC$(KI%) = COMO$ Then
                REGFA$(KI%) = REGFA$(KI%) + MKS$(IL&)
                If FEMO% = 0 Then
                   U& = CVS(Left$(REGFA$(KI%), 4))
                   XXX$ = REGLEIDO$("FACTUR", U&)
                   FEMO% = CVI(Mid$(XXX$, 7, 2))
                   Call REPLA(XMV$, MKI$(FEMO%), 1, 2)
                   Call GRAREG("MOVISTO", XMV$, IL&)
                End If
                GoTo 219
              End If
            Next KI%
          End If
        End If
219 Next IL&
    '
    FIN& = CAFAC%
    J& = 0
    For IL& = 1 To FIN&
        Call TRACE(20, IL&, FIN&)
        U& = CVS(Left$(REGFA$(IL&), 4))
        If U& > 0 Then
          If U& <= ULTREG&("FACTUR") Then
            X$ = REGLEIDO$("FACTUR", U&)
            FF% = CVI(Mid$(X$, 7, 2))
            NRO& = CVS(Mid$(X$, 3, 4))
            NC% = CVI(Mid$(X$, 1, 2))
            VENDE% = VENDCLI%(NC%)
            TI% = Asc(Mid$(X$, 13, 1))
            '
            ITOT# = CVD(Mid$(X$, 15, 8))
            TIDOC% = Asc(Mid$(X$, 13, 1))
            VADOC% = Asc(Mid$(X$, 14, 1))
            SIG% = 1
            If TIDOC% >= 4 Then
               ITOT# = (-1) * Abs(ITOT#)
               SIG% = (-1)
            End If
            '
            If Abs(ITOT#) >= 0.005 Then
               For IJ& = 5 To Len(REGFA$(IL&)) Step 4
                  U2& = CVS(Mid$(REGFA$(IL&), IJ&, 4))
                  If U2& > 0 Then
                    If U2& <= ULTREG&("MOVISTO") Then
                      Z$ = REGLEIDO$("MOVISTO", U2&)
                      CAN = CVD(Mid$(Z$, 96, 8)) - CVD(Mid$(Z$, 88, 8))
                      PREUNI = Abs(CVD(Mid$(Z$, 73, 8)))
                      TOT# = CDbl(PREUNI * CAN)
                      CI% = CVI(Mid$(Z$, 3, 2))
                      ITOT# = ITOT# - TOT# * SIG%
                      GoSub 600
                    End If
                  End If
               Next IJ&
               If Abs(ITOT#) >= 0.005 Then
                  CI% = 0
                  PREUNI = ITOT#
                  CAN = 1
                  TOT# = ITOT#
                  GoSub 600
               End If
            End If
            '
          End If
        End If
    Next IL&
    '
    Call SETULTREG("LICOMAIV", J&)
    Call CIERRARCH("LICOMAIV")
    Call FINAL("*SORCOMA")
    YALIQUI% = 1
    '
999 If YALIQUI% = 1 Then
        FORLICOMAI.List1.Clear
        For VENDE% = 0 To 255
            If Abs(TOCOMI#(VENDE%)) >= 0.005 Then
                If VENDE% = 0 Then
                    NOVEN$ = "DIRECCION"
                  Else
                    NOVEN$ = ECOARCH$("VENDEDOR", Chr$(VENDE%))
                End If
                TXT$ = Space$(48)
                Call REPLA(TXT$, CSTRING$(MKI$(VENDE%), 1, 3, 0, 2), 1, 4)
                Call REPLA(TXT$, NOVEN$, 6, 17)
                Call REPLA(TXT$, CSTRING$(MKD$(TOCOMI#(VENDE%)), 4, 10, 2, 2), 23, 10)
                FORLICOMAI.List1.AddItem TXT$
            End If
        Next VENDE%
        FORLICOMAI.Check4.Value = 1
        '
        If FORLICOMAI.List1.ListCount > 0 Then
            FORLICOMAI.List1.ListIndex = 0
        End If
        FORLICOMAI.List1.Refresh
        '
        FORLICOMAI.Label4.Caption = HOS$
        FORLICOMAI.Label6.Caption = Left$(Time$, 5)
        '
        FORLICOMAI.List1.Refresh
        FORLICOMAI.Command3.Enabled = True
        '
    End If
    '
    Screen.MousePointer = 1
    Exit Sub
    '
600 Y$ = REGBLAN$("LICOMAIV")
    Call REPLA(Y$, Chr$(VENDE%), 1, 1)
    Call REPLA(Y$, MKI$(NC%), 2, 2)
    Call REPLA(Y$, MKI$(FF%), 4, 2)
    Call REPLA(Y$, Chr$(TI%), 6, 1)
    Call REPLA(Y$, MKS$(NRO&), 7, 4)
    Call REPLA(Y$, MKI$(CI%), 11, 2)
    Call REPLA(Y$, MKS$(PREUNI), 13, 4)
    Call REPLA(Y$, MKS$(CAN), 17, 4)
    Call REPLA(Y$, MKD$(TOT#), 21, 8)
    PREBAS = PREBASE(CI%)
    If PREBAS < 0.01 Then PREBAS = PREUNI / 1.92
    Call REPLA(Y$, MKS$(PREBAS), 37, 4)
    Call REPLA(Y$, MKS$(PREUNI - PREBAS), 41, 4)
    Call REPLA(Y$, MKD$(CDbl((PREUNI - PREBAS) * CAN)), 45, 8)
    PORCOM = COVEN(VENDE%)
    ICOM# = PORCOM * (CDbl((PREUNI - PREBAS) * CAN)) / 100
    Call REPLA(Y$, MKS$(PORCOM), 53, 4)
    Call REPLA(Y$, MKD$(ICOM#), 57, 8)
    '
    J& = J& + 1
    Call GRAREG("LICOMAIV", Y$, J&)
    TOCOMI#(VENDE%) = TOCOMI#(VENDE%) + ICOM#
    '
Return
'
End Sub
'
Sub CALCOMAI()
    '
    Dim PORCO(256), TOCOMI#(256), TOCOB#(256)
    Dim NRECI&(32767)
    Dim REGFA$(32767)
    Dim NCV%(32767)
    Dim NROV&(32767)
    Dim IDET#(32767)
    Dim IFAC#(32767)
    '
MsgBox "NO USAR - UTILIZAR CALCOMVEN"
Exit Sub

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
    MMXX0$ = TRIM$(FORLICOMAI.Label3.Caption)
    MMXX1$ = MESANO$(MMXX0$, DES%, HAS%)
    If TRIM$(MMXX1$) = "" Then
        FORSELMEC.Show 1
        Exit Sub
    End If
    
    For U& = 1 To ULTREG&("DETCOBV")
        X$ = REGLEIDO$("DETCOBV", U&)
        NRECI&(U&) = CVS(Mid$(X$, 5, 4))
        REGFA$(U&) = ""
    Next U&
    '
    CARECI& = U&
    '
    FIN& = ULTREG&("APLICOB")
    FORLICOMAI.Check1.Value = 2
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
    FORLICOMAI.Check1.Value = 1
    '
    FIN& = CARECI&
    FORLICOMAI.Check2.Value = 2
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
    FORLICOMAI.Check2.Value = 1
    '
    FIN& = ULTREG&("MOVISTO")
    FORLICOMAI.Check3.Value = 2
    '
    For IL& = 1 To FIN&
        Call TRACE(20, IL&, FIN&)
        X$ = REGLEIDO$("MOVISTO", IL&)
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
    FORLICOMAI.Check3.Value = 1
    '
    FIN& = CAFAC%
    FORLICOMAI.Check4.Value = 2
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
    FORLICOMAI.List1.Clear
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
        FORLICOMAI.List1.AddItem TXT$
      End If
    Next VENDE%
    FORLICOMAI.Check4.Value = 1
    '
    If FORLICOMAI.List1.ListCount > 0 Then
        FORLICOMAI.List1.ListIndex = 0
    End If
    FORLICOMAI.List1.Refresh
    FORLICOMAI.Command3.Enabled = True
    FORLICOMAI.Command4.Enabled = True
    '
End Sub
'

Sub PRILICOMAI(MOCALCOM%)
    '
    MODOCAL% = MOCALCOM%
    If MODOCAL% = 1 Then MODOCA$ = "Por Cobranza y Precio Base"
    If MODOCAL% = 2 Then MODOCA$ = "Por Pedidos Netos de Descuento"
    If MODOCAL% = 3 Then MODOCA$ = "Por Facturacion"
    If MODOCAL% = 4 Then MODOCA$ = "Por Cobranza Neta de IVA"
    If MODOCAL% < 1 Then
        Exit Sub
    End If
    '
    If FORLICOMAI.List1.ListIndex >= 0 Then
        TXT$ = FORLICOMAI.List1.Text
        VENDE% = Val(Left$(TXT$, 3))
        NOVEN$ = "Clientes de Direccion"
        If VENDE% > 0 Then
            NOVEN$ = ECOARCH$("VENDEDOR", Chr$(VENDE%))
        End If
        '
        CORRESP$ = AJUSTI$(FORLICOMAI.Label3.Caption, 12)
        '
        If MODOCAL% = 3 Then
            Call COPYSTRU("LICOMAIV", "LICOMAIX")
            Call BLANARCH("LICOMAIX")
            Call HEADERS("LICOMAIX", "")
            Call HEADERS("LICOMAIX", "Corresponde a: " + CORRESP$)
            Call HEADERS("LICOMAIX", "Vendedor: " + NOVEN$)
            '
            FIN& = ULTREG&("LICOMAIV")
            '
            For IL& = 1 To FIN&
                Call TRACE(20, IL&, FIN&)
                X$ = REGLEIDO$("LICOMAIV", IL&)
                If Asc(Mid$(X$, 1, 1)) = VENDE% Then
                    Call REGAPP("LICOMAIX", X$)
                End If
            Next IL&
            '
            Call CIERRARCH("*.*")
            Call CONECRUN("*LICOMAV", "Listado de Comisiones Liquidadas")
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
