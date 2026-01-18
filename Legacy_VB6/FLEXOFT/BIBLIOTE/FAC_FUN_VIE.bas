Attribute VB_Name = "FAC_FUN_VIE"
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


