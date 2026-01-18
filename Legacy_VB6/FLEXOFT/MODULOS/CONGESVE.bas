Attribute VB_Name = "CONGESVE"

Sub GELISALCLI()
    '
    HOXX = HOY(HO$)
    Dim NC%(16384), DEUDA#(16384), SALDO#(16384)
    '
    Screen.MousePointer = 11
    Call HEADERS("LISACLI", "")
    Call HEADERS("LISACLI", "Generado el: " + HO$ + " - " + Left$(Time$, 5) + " hs.")
    '
    cx& = ULTREG&("DEUDOR1")
    '
    For IL& = 1 To cx&
        Call TRACE(20, IL&, cx&)
        X$ = REGLEIDO$("DEUDOR1", IL&)
        NC%(IL&) = CVI(Left$(X$, 2))
        SALDO#(IL&) = SALDCLI#(NC%(IL&))
        DEUDA#(IL&) = 0
    Next IL&
    '
    FIN& = ULTREG&("SVD202")
    For IL& = 1 To FIN&
        ss$ = REGLEIDO$("SVD202", IL&)
        NCI% = CVI(Left$(ss$, 2))
        TI% = Asc(Mid$(ss$, 3, 1))
        If TI% > 7 Then TI% = 2
        SALDEU# = CVD(Mid$(ss$, 11, 8))
        If TI% > 3 Then SALDEU# = (-1) * Abs(SALDEU#)
        For JL& = 1 To cx&
            If NC%(JL&) = NCI% Then
                DEUDA#(JL&) = DEUDA#(JL&) + SALDEU#
                GoTo 809
            End If
        Next JL&
809 Next IL&

900 RELISA& = 0
    '
    For IL& = 1 To cx&
       If NC%(IL&) > 0 Then
         If Abs(SALDO#(IL&)) >= 0.005 Or Abs(DEUDA#(IL&)) >= 0.005 Then
            Y$ = REGBLAN$("LISACLI")
            Call REPLA(Y$, MKI$(NC%(IL&)), 1, 2)
            Call REPLA(Y$, MKD$(SALDO#(IL&)), 3, 8)
            Call REPLA(Y$, MKD$(DEUDA#(IL&)), 11, 8)
            If Abs(SALDO#(IL&) - DEUDA#(IL&)) > 0.005 Then
               Call REPLA(Y$, "ATENCION: Saldo <> Deuda", 19, 30)
            End If
            '
            RELISA& = RELISA& + 1
            Call GRAREG("LISACLI", Y$, RELISA&)
         End If
       End If
    Next IL&
    '
    Call SETULTREG("LISACLI", RELISA&)
    Call CIERRARCH("*.*")
    Screen.MousePointer = 1
    '
    If COGEVENT.Option1.Value = True Then
          Call FINAL("*LISACLI")
        Else
          Call FINAL("*LIMPCLI")
    End If
    '
End Sub
'

Sub RANKCLI()
    '
    Dim TOTVENTA#(32767)
    '
10  FEX = HOY(HO$)
    HAS% = FEX
    HASS$ = FECHATEX$(FEX)
    DESS$ = "01" + Mid$(HASS$, 3)
    DES% = FECHANUM(DESS$)
    VDEF$ = MKI$(DES%) + MKI$(HAS%)
    '
20  OBX$ = OBJPLA$("DESHASFECHA", VDEF$)
    If OBX$ = "" Then
        GoTo 99
    End If
    '
    DES% = CVI(Left$(OBX$, 2))
    HAS% = CVI(Mid$(OBX$, 3, 2))
    If DES% > HAS% Then
        GoTo 10
    End If
    FEX = DES%: PERIO$ = FECHATEX$(FEX)
    FEX = HAS%: PERIO$ = PERIO$ + " - " + FECHATEX$(FEX)
    VDEF$ = OBX$
    '
    VENDSEL% = 0
    RSX$ = REGSEL$("VENDEDOR")
    If Len(RSX$) >= 4 Then
      VENDSEL% = Asc(Left$(RSX$, 1))
      PERIO$ = PERIO$ + " - Vendedor: " + TRIM$(Mid$(RSX$, 2, 30))
    End If
    '
    Screen.MousePointer = 11
    J& = 0: TOTAGEN# = 0
    For U& = 0 To 32767
      TOTVENTA#(U&) = 0
    Next U&
    '
    FIN& = ULTREG&("FACTUR")
    TOTAGEN# = 0
    For U& = 1 To FIN&
      Call TRACE(20, U&, FIN&)
      X$ = REGLEIDO$("FACTUR", U&)
      FF% = CVI(Mid$(X$, 7, 2))
      If FF% >= DES% Then
        If FF% <= HAS% Then
          NC% = CVI(Left$(X$, 2))
          If NC% > 0 Then
            If NC% <= 32767 Then
              If VENDSEL% = 0 Or VENDCLI%(NC%) = VENDSEL% Then
                    TI% = Asc(Mid$(X$, 13, 1))
                    If TI% < 0 Or TI% > 7 Then TI% = 2 ' default factura
                    If TI% > 0 Then
                       If TI% < 7 Then
                          NETO# = (Int(100 * (CVD(Mid$(X$, 15, 8)) + CVD(Mid$(X$, 49, 8))) + 0.5)) / 100
                          If TI% > 3 Then
                             NETO# = (-1) * Abs(NETO#)
                          End If
                          TOTVENTA#(NC%) = TOTVENTA#(NC%) + NETO#
                          TOTAGEN# = TOTAGEN# + NETO#
                       End If
                    End If
              End If
            End If
          End If
        End If
      End If
    Next U&
    '
    For NCI% = 1 To 32766
       If Abs(TOTVENTA#(NCI%)) > 0.005 Then
          RELIB$ = "A" + CSTRING$(MKD$(TOTAGEN# - TOTVENTA#(NCI%)), 4, 12, 2, 2)
          PORPAR = 100 * TOTVENTA#(NCI%) / TOTAGEN#
          RELIB$ = RELIB$ + CSTRING$(MKI$(NCI%), 1, 6, 0, 2)
          RELIB$ = RELIB$ + CSTRING$(MKS$(PORPAR), 4, 8, 2, 2)
          RELIB$ = RELIB$ + CSTRING$(MKD$(TOTVENTA#(NCI%)), 4, 12, 2, 2)
          RESUESTA.List1.AddItem RELIB$
       End If
    Next NCI%
    '
    J& = 0: PORACU = 0
    FIN& = RESUESTA.List1.ListCount
    '
    For U& = 1 To FIN&
        RESUESTA.List1.ListIndex = U& - 1
        NC% = XVALO(Mid$(RESUESTA.List1.TEXT, 14, 6))
        TONEVEN# = XVALO(Mid$(RESUESTA.List1.TEXT, 28, 12))
        PORPAR = XVALO(Mid$(RESUESTA.List1.TEXT, 20, 8))
        PORACU = PORACU + PORPAR
        '
        Z$ = REGBLAN$("RANKCLI")
        J& = J& + 1: JI% = J&
        Call REPLA(Z$, MKI$(JI%), 1, 2)
        Call REPLA(Z$, MKI$(NC%), 3, 2)
        Call REPLA(Z$, MKD$(TONEVEN#), 5, 8)
        Call REPLA(Z$, MKS$(PORPAR), 13, 4)
        Call REPLA(Z$, MKS$(PORACU), 17, 4)
        Call GRAREG("RANKCLI", Z$, J&)
        '
    Next U&
    '
    Call SETULTREG("RANKCLI", J&)
    Call CIERRARCH("RANKCLI")
    Call HEADERS("RANKCLI", "")
    Call HEADERS("RANKCLI", "Periodo: " + PERIO$)
    Screen.MousePointer = 1
    RESUESTA.List1.Clear
    '
    If RESUESTA.Option1(4).Value = True Then
         Call FINAL("*RANKCLI")
       Else
         Call FINAL("*LRANCLI")
    End If
    GoTo 20
    '
99  Screen.MousePointer = 1
    Call CIERRARCH("*.*")
    '
End Sub
'

Sub RANKPIE()
    '
    Dim TOTVENTA#(32767)
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
        GoTo 99
    End If
    '
    DES% = CVI(Left$(OBX$, 2))
    HAS% = CVI(Mid$(OBX$, 3, 2))
    If DES% > HAS% Then
        GoTo 10
    End If
    FEX = DES%: PERIO$ = FECHATEX$(FEX)
    FEX = HAS%: PERIO$ = PERIO$ + " - " + FECHATEX$(FEX)
    '
    Screen.MousePointer = 11
    J& = 0
    '
    ARMOVI$ = "MOVISTO"
15  FIN& = ULTREG&(ARMOVI$)
    TOTAGEN# = 0
    For U& = 1 To FIN&
        Call TRACE(20, U&, FIN&)
        X$ = REGLEIDO$(ARMOVI$, U&)
        FF% = CVI(Left$(X$, 2))
        If FF% >= DES% Then
           If FF% <= HAS% Then
              Ci% = CVI(Mid$(X$, 3, 2))
              If Ci% > 0 Then
                 If Ci% <= 32767 Then
                    COMO$ = UCase$(Mid$(X$, 33, 2))
                    If TRIM$(COMO$) = "" Then
                       COMO$ = UCase$(Mid$(X$, 23, 2))
                    End If
                    If COMO$ = "RT" Or COMO$ = "FC" Or COMO$ = "NC" Or COMO$ = "ND" Then
                       NETO# = (CVD(Mid$(X$, 96, 8)) - CVD(Mid$(X$, 88, 8))) * CVD(Mid$(X$, 73, 8))
                       TOTVENTA#(Ci%) = TOTVENTA#(Ci%) + NETO#
                       TOTAGEN# = TOTAGEN# + NETO#
                    End If
                 End If
              End If
           End If
        End If
    Next U&
    If ARMOVI$ = "MOVISTO" Then
      Call COPYSTRU("MOVISTO", "MOVISVIE")
      ARMOVI$ = "MOVISVIE"
      GoTo 15
    End If
    '
    For NCI% = 1 To NUMAS%
       If Abs(TOTVENTA#(NCI%)) > 0.005 Then
          RELIB$ = "A" + CSTRING$(MKD$(TOTAGEN# - TOTVENTA#(NCI%)), 4, 12, 2, 2)
          PORPAR = 100 * TOTVENTA#(NCI%) / TOTAGEN#
          RELIB$ = RELIB$ + CSTRING$(MKI$(NCI%), 1, 6, 0, 2)
          RELIB$ = RELIB$ + CSTRING$(MKS$(PORPAR), 4, 8, 2, 2)
          RELIB$ = RELIB$ + CSTRING$(MKD$(TOTVENTA#(NCI%)), 4, 12, 2, 2)
          RESUESTA.List1.AddItem RELIB$
       End If
    Next NCI%
    '
    J& = 0
    FIN& = RESUESTA.List1.ListCount
    '
    For U& = 1 To FIN&
        RESUESTA.List1.ListIndex = U& - 1
        Ci% = XVALO(Mid$(RESUESTA.List1.TEXT, 14, 6))
        TONEVEN# = XVALO(Mid$(RESUESTA.List1.TEXT, 28, 12))
        PORPAR = XVALO(Mid$(RESUESTA.List1.TEXT, 20, 8))
        PORACU = PORACU + PORPAR
        '
        Z$ = REGBLAN$("RANKPIE")
        J& = J& + 1: JI% = J&
        Call REPLA(Z$, MKI$(JI%), 1, 2)
        Call REPLA(Z$, MKI$(Ci%), 3, 2)
        Call REPLA(Z$, MKD$(TONEVEN#), 5, 8)
        Call REPLA(Z$, MKS$(PORPAR), 13, 4)
        Call REPLA(Z$, MKS$(PORACU), 17, 4)
        Call GRAREG("RANKPIE", Z$, J&)
        '
    Next U&
    '
    Call SETULTREG("RANKPIE", J&)
    Call CIERRARCH("RANKPIE")
    Call HEADERS("RANKPIE", "")
    Call HEADERS("RANKPIE", "Periodo: " + PERIO$)
    Screen.MousePointer = 1
    RESUESTA.List1.Clear
    '
    If RESUESTA.Option1(4).Value = True Then
         Call FINAL("*RANKPIE")
       Else
         Call FINAL("*LRANPIE")
    End If
    '
99  Screen.MousePointer = 1
    Call CIERRARCH("*.*")
    '
End Sub
'

Sub REVEPRO()
    '
    Dim TOTVENTA#(32767), TOTCAN(32767)
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
        GoTo 99
    End If
    '
    DES% = CVI(Left$(OBX$, 2))
    HAS% = CVI(Mid$(OBX$, 3, 2))
    If DES% > HAS% Then
        GoTo 10
    End If
    FEX = DES%: PERIO$ = FECHATEX$(FEX)
    FEX = HAS%: PERIO$ = PERIO$ + " - " + FECHATEX$(FEX)
    '
    Screen.MousePointer = 11
    J& = 0
    '
    ARMOVI$ = "MOVISTO"
15  FIN& = ULTREG&(ARMOVI$)
    Screen.MousePointer = 11
    TOTAGEN# = 0
    For U& = 1 To FIN&
      Call TRACE(20, U&, FIN&)
      X$ = REGLEIDO$(ARMOVI$, U&)
      If CVI(Mid$(X$, 84, 2)) > 0 Then ' codigo de cliente
        FF% = CVI(Left$(X$, 2))
        If FF% >= DES% Then
          If FF% <= HAS% Then
            Ci% = CVI(Mid$(X$, 3, 2))
            If Ci% > 0 Then
              If Ci% <= 32767 Then
                COMO$ = UCase$(Mid$(X$, 33, 2))
                If TRIM$(COMO$) = "" Then
                  COMO$ = UCase$(Mid$(X$, 23, 2))
                End If
                If COMO$ = "RT" Or COMO$ = "FC" Or COMO$ = "NC" Or COMO$ = "ND" Then
                  CANTI = CVD(Mid$(X$, 96, 8)) - CVD(Mid$(X$, 88, 8))
                  NETO# = CVD(Mid$(X$, 73, 8)) * CANTI
                  TOTCAN(Ci%) = TOTCAN(Ci%) + CANTI
                  TOTVENTA#(Ci%) = TOTVENTA#(Ci%) + NETO#
                End If
              End If
            End If
          End If
        End If
      End If
    Next U&
    If ARMOVI$ = "MOVISTO" Then
      Call COPYSTRU("MOVISTO", "MOVISVIE")
      ARMOVI$ = "MOVISVIE"
      GoTo 15
    End If
    '
    J& = 0
    For U& = 1 To NUINV%
       XXX$ = REGLEIDO$("INVERT", U& + 1)
       NCI% = CVI(Mid$(XXX$, 17, 2))
       If NCI% > 0 Then
          If NCI% <= NUMAS% Then
             If Abs(TOTCAN(NCI%)) > 0.005 Then
                RELIB$ = REGBLAN$("REVEPRO")
                Call REPLA(RELIB$, MKI$(NCI%), 1, 2)
                Call REPLA(RELIB$, MKS$(TOTVENTA#(NCI%) / TOTCAN(NCI%)), 3, 4)
                Call REPLA(RELIB$, MKS$(TOTCAN(NCI%)), 7, 4)
                Call REPLA(RELIB$, MKD$(TOTVENTA#(NCI%)), 11, 8)
                J& = J& + 1
                Call GRAREG("REVEPRO", RELIB$, J&)
             End If
          End If
       End If
    Next U&
    '
    Call SETULTREG("REVEPRO", J&)
    Call CIERRARCH("REVEPRO")
    '
    Call HEADERS("REVEPRO", "")
    Call HEADERS("REVEPRO", "Periodo: " + PERIO$)
    Screen.MousePointer = 1
    '
    Call FINAL("*REVEPRO")
    '
99  Screen.MousePointer = 1
    Call CIERRARCH("*.*")
    '
End Sub
'

Sub FICHENCLI()
    '
10  FEX = HOY(HO$)
    HAS% = FEX
    HASS$ = FECHATEX$(FEX)
    DESS$ = "01" + Mid$(HASS$, 3)
    DES% = FECHANUM(DESS$)
    '
20  VDEF$ = MKI$(DES%) + MKI$(HAS%)
    '
    OBX$ = OBJPLA$("DESHASFECHA", VDEF$)
    If OBX$ = "" Then
        GoTo 99
    End If
    '
    DES% = CVI(Left$(OBX$, 2))
    HAS% = CVI(Mid$(OBX$, 3, 2))
    If DES% > HAS% Then
        GoTo 20
    End If
    FEX = DES%: PERIO$ = FECHATEX$(FEX)
    FEX = HAS%: PERIO$ = PERIO$ + " - " + FECHATEX$(FEX)
    '
    Screen.MousePointer = 11
    J& = 0
    '
    ARMOVI$ = "MOVISTO"
15  FIN& = ULTREG&(ARMOVI$)
    Screen.MousePointer = 11
    TOTAGEN# = 0
    For U& = 1 To FIN&
        '
        Call TRACE(20, U&, FIN&)
        X$ = REGLEIDO$(ARMOVI$, U&)
        CLI% = CVI(Mid$(X$, 84, 2))
        If CLI% > 0 Then
           FF% = CVI(Left$(X$, 2))
           If FF% >= DES% Then
              If FF% <= HAS% Then
                 CII% = CVI(Mid$(X$, 3, 2))
                 If CII% > 0 Then
                    If CII% <= NUMAS% Then
                       If rasocli$(CLI%) <> "" Then
                          '
                          REMIFAC$ = Mid$(X$, 23, 20)
                          CANTU# = CVD(Mid$(X$, 96, 8)) - CVD(Mid$(X$, 88, 8))
                          '
                          Z$ = REGBLAN$("FICHENT")
                          Call REPLA(Z$, MKI$(CLI%), 1, 2)
                          Call REPLA(Z$, PERIO$, 3, 20)
                          Call REPLA(Z$, MKI$(FF%), 23, 2)
                          Call REPLA(Z$, MKI$(CII%), 25, 2)
                          Call REPLA(Z$, MKD$(CANTU#), 27, 8)
                          Call REPLA(Z$, REMIFAC$, 35, 20)
                          '
                          J& = J& + 1
                          Call GRAREG("FICHENT", Z$, J&)
                          '
                       End If
                    End If
                 End If
              End If
           End If
        End If
    Next U&
    If ARMOVI$ = "MOVISTO" Then
      Call COPYSTRU("MOVISTO", "MOVISVIE")
      ARMOVI$ = "MOVISVIE"
      GoTo 15
    End If
    '
    Call SETULTREG("FICHENT", J&)
    Call CIERRARCH("FICHENT")
    Screen.MousePointer = 1
    '
    Call FINAL("*FICHENT")
    '
99  Call CIERRARCH("*.*")
    '
End Sub
'

Sub RESUENCLI()
    '
    Dim TOCAEN$(32767)
    '
10  FEX = HOY(HO$)
    HAS% = FEX
    HASS$ = FECHATEX$(FEX)
    DESS$ = "01" + Mid$(HASS$, 3)
    DES% = FECHANUM(DESS$)
    '
20  VDEF$ = MKI$(DES%) + MKI$(HAS%)
    '
    OBX$ = OBJPLA$("DESHASFECHA", VDEF$)
    If OBX$ = "" Then
        GoTo 99
    End If
    '
    DES% = CVI(Left$(OBX$, 2))
    HAS% = CVI(Mid$(OBX$, 3, 2))
    If DES% > HAS% Then
        GoTo 20
    End If
    FEX = DES%: PERIO$ = FECHATEX$(FEX)
    FEX = HAS%: PERIO$ = PERIO$ + " - " + FECHATEX$(FEX)
    '
    If RESUESTA.Option1(4).Value = True Then
22     OBX$ = OBJPLA$("SELECLI", MKI$(NC%))
       If OBX$ = "" Then GoTo 20
       '
       NC% = CVI(Left$(OBX$, 2))
       Screen.MousePointer = 11
25     TTXX$ = ""
       ARMOVI$ = "MOVISTO"
26     RFF$ = RECFILT$(ARMOVI$, 12, MKI$(NC%))
       FIN& = Len(RFF$)
       For UK& = 1 To FIN& Step 4
          Call TRACE(20, UK&, FIN&)
          U& = CVS(Mid$(RFF$, UK&, 4))
          X$ = REGLEIDO$(ARMOVI$, U&)
          CLI% = CVI(Mid$(X$, 84, 2))
          If CLI% <> NC% Then
             X$ = REGLEIDO$(ARMOVI$, 1)
             Call GRAREG(ARMOVI$, X$, 1)
             Call CIERRARCH(ARMOVI$)
             GoTo 25
          End If
          FF% = CVI(Left$(X$, 2))
          If FF% >= DES% Then
            If FF% <= HAS% Then
              If CLI% > 0 Then
                If CLI% <= 32767 Then
                  CII% = CVI(Mid$(X$, 3, 2))
                  If CII% > 0 Then
                    If CII% <= NUMAS% Then
                      CANTU = CVD(Mid$(X$, 96, 8)) - CVD(Mid$(X$, 88, 8))
                      '
                      For JI% = 1 To Len(TTXX$) Step 6
                        If CVI(Mid$(TTXX$, JI%, 2)) = CII% Then
                          CANTU = CANTU + CVS(Mid$(TTXX$, JI% + 2, 4))
                          Call REPLA(TTXX$, MKS$(CANTU), JI% + 2, 4)
                          GoTo 29
                        End If
                      Next JI%
                      TTXX$ = TTXX$ + MKI$(CII%) + MKS$(CANTU)
                      '
29                  End If
                  End If
                End If
              End If
            End If
          End If
       Next UK&
       If ARMOVI$ = "MOVISTO" Then
         Call COPYSTRU("MOVISTO", "MOVISVIE")
         ARMOVI$ = "MOVISVIE"
         GoTo 26
       End If
       '
       If Len(TTXX$) < 1 Then
          Screen.MousePointer = 1
          Call MENSERR(24, "Cliente sin Movimientos")
          GoTo 22
       End If
       '
       RESUESTA.List1.Clear
       For JI% = 1 To Len(TTXX$) Step 6
          CII% = CVI(Mid$(TTXX$, JI%, 2))
          CANTU = CVS(Mid$(TTXX$, JI% + 2, 4))
          If CII% > 0 Then
            If CII% <= NUMAS% Then
              If Abs(CANTU) >= 0.05 Then
                Z$ = Space$(24)
                Call REPLA(Z$, "A" + CODEXT$(CII%), 1, 17)
                Call REPLA(Z$, CSTRING$(MKI$(JI%), 1, 7, 0, 2), 18, 7)
                RESUESTA.List1.AddItem Z$
              End If
            End If
          End If
       Next JI%
       '
       jj& = 0
       For JK% = 1 To RESUESTA.List1.ListCount
          Z$ = RESUESTA.List1.List(JK% - 1)
          CII% = CODINT%(Mid$(Z$, 2, 16))
          JI% = XVALO(Mid$(Z$, 18, 7))
          CANTU = CVS(Mid$(TTXX$, JI% + 2, 4))
          '
          X$ = REGBLAN$("CORENCLI")
          Call REPLA(X$, MKI$(NC%), 1, 2)
          Call REPLA(X$, MKI$(CII%), 3, 2)
          Call REPLA(X$, MKS$(CANTU), 5, 4)
          Call REPLA(X$, PERIO$, 25, 20)
          jj& = jj& + 1
          Call GRAREG("CORENCLI", X$, jj&)
       Next JK%
       Call SETULTREG("CORENCLI", jj&)
       Call CIERRARCH("CORENCLI")
       '
       Screen.MousePointer = 1
       Call FINAL("*CORENCL/" + TRIM$(Str$(NC%)) + "    ")
       GoTo 22
       '
    End If
    '
    Screen.MousePointer = 11
    J& = 0
    '
    ARMOVI$ = "MOVISTO"
15  FIN& = ULTREG&(ARMOVI$)
    Screen.MousePointer = 11
    TOTAGEN# = 0
    For U& = 1 To FIN&
        '
        Call TRACE(20, U&, FIN&)
        X$ = REGLEIDO$(ARMOVI$, U&)
        CLI% = CVI(Mid$(X$, 84, 2))
        FF% = CVI(Left$(X$, 2))
        If FF% >= DES% Then
          If FF% <= HAS% Then
            If CLI% > 0 Then
              If CLI% <= 32767 Then
                CII% = CVI(Mid$(X$, 3, 2))
                If CII% > 0 Then
                  If CII% <= NUMAS% Then
                    CANTU = CVD(Mid$(X$, 96, 8)) - CVD(Mid$(X$, 88, 8))
                    TTXX$ = TOCAEN$(CLI%)
                    '
                    For JI% = 1 To Len(TTXX$) Step 6
                      If CVI(Mid$(TTXX$, JI%, 2)) = CII% Then
                        CANTU = CANTU + CVS(Mid$(TTXX$, JI% + 2, 4))
                        Call REPLA(TTXX$, MKS$(CANTU), JI% + 2, 4)
                        GoTo 19
                      End If
                    Next JI%
                    TTXX$ = TTXX$ + MKI$(CII%) + MKS$(CANTU)
                    '
19                  TOCAEN$(CLI%) = TTXX$
                    '
                  End If
                End If
              End If
            End If
          End If
        End If
    Next U&
    If ARMOVI$ = "MOVISTO" Then
      Call COPYSTRU("MOVISTO", "MOVISVIE")
      ARMOVI$ = "MOVISVIE"
      GoTo 15
    End If
    '
    jj& = 0
    FIN& = ULTREG&("DEUDOR1")
    For U& = 1 To FIN&
      Call TRACE(20, U&, FIN&)
      Y$ = REGLEIDO$("DEUDOR1", U&)
      i% = CVI(Left$(Y$, 2))
      If i% > 0 Then
       If i% <= 32767 Then
         TTXX$ = TOCAEN$(i%)
         If Len(TTXX$) > 0 Then
           RESUESTA.List1.Clear
           For JI% = 1 To Len(TTXX$) Step 6
             CII% = CVI(Mid$(TTXX$, JI%, 2))
             CANTU = CVS(Mid$(TTXX$, JI% + 2, 4))
             If CII% > 0 Then
               If CII% <= NUMAS% Then
                 If Abs(CANTU) >= 0.05 Then
                    Z$ = Space$(24)
                    Call REPLA(Z$, "A" + CODEXT$(CII%), 1, 17)
                    Call REPLA(Z$, CSTRING$(MKI$(JI%), 1, 7, 0, 2), 18, 7)
                    RESUESTA.List1.AddItem Z$
                 End If
               End If
             End If
           Next JI%
           '
           For JK% = 1 To RESUESTA.List1.ListCount
             Z$ = RESUESTA.List1.List(JK% - 1)
             CII% = CODINT%(Mid$(Z$, 2, 16))
             JI% = XVALO(Mid$(Z$, 18, 7))
             CANTU = CVS(Mid$(TTXX$, JI% + 2, 4))
             '
             X$ = REGBLAN$("RESENCLI")
             Call REPLA(X$, MKI$(i%), 1, 2)
             Call REPLA(X$, MKI$(CII%), 3, 2)
             Call REPLA(X$, MKS$(CANTU), 5, 4)
             Call REPLA(X$, PERIO$, 25, 20)
             jj& = jj& + 1
             Call GRAREG("RESENCLI", X$, jj&)
           Next JK%
         End If
       End If
      End If
    Next U&
    '
    Call SETULTREG("RESENCLI", jj&)
    Call CIERRARCH("RESENCLI")
    Call HEADERS("RESENCLI", "")
    Call HEADERS("RESENCLI", "Corresponde a: " + PERIO$)
    Screen.MousePointer = 1
    '
    Call FINAL("*LIRENCL")
    '
99  Call CIERRARCH("*.*")
    '
End Sub

Sub LILIMCRE()
   '
   Screen.MousePointer = 11
   '
   Dim DEUCUECO#(16384), CHEPEAC#(16384)
   '
   FIN& = ULTREG&("SVD202")
   For IL& = 1 To FIN&
      Call TRACE(20, IL&, FIN&)
      ss$ = REGLEIDO$("SVD202", IL&)
      NCI% = CVI(Left$(ss$, 2))
      If NCI% > 0 Then
         If NCI% <= 16384 Then
            TI% = Asc(Mid$(ss$, 3, 1))
            If TI% > 7 Then TI% = 2
            SALDEU# = CVD(Mid$(ss$, 11, 8))
            If TI% > 3 Then SALDEU# = (-1) * Abs(SALDEU#)
            DEUCUECO#(NCI%) = DEUCUECO#(NCI%) + SALDEU#
            GoTo 19
         End If
      End If
19 Next IL&
   '
   HOII% = HOY(HOS$)
   FIN& = ULTREG&("CHECAR")
   For IL& = 1 To FIN&
      Call TRACE(20, IL&, FIN&)
      XX$ = REGLEIDO$("CHECAR", IL&)
      NCI% = CVI(Mid$(XX$, 3, 2))
      If NCI% > 0 Then
         If NCI% <= 16384 Then
            If CVI(Mid$(XX$, 41, 2)) >= HOII% Then
               CHEPEAC#(NCI%) = CHEPEAC#(NCI%) + CVD(Mid$(XX$, 33, 8))
            End If
         End If
      End If
   Next IL&
   '
   FIN& = ULTREG&("DEUDOR1")
   jj& = 0
   For IL& = 1 To FIN&
      Call TRACE(20, IL&, FIN&)
      XX$ = REGLEIDO$("DEUDOR1", IL&)
      NCI% = CVI(Left$(XX$, 2))
      If NCI% > 0 Then
         If NCI% <= 16384 Then
            If Abs(DEUCUECO#(NCI%)) + Abs(CHEPEAC#(NCI%)) >= 0.005 Then
               YY$ = REGLEIDO$("DEUDOR2", IL&)
               CRELIM# = CVS(Mid$(YY$, 121, 4))
               DISPO# = CRELIM# - DEUCUECO#(NCI%) - CHEPEAC#(NCI%)
               EXCES# = 0
               If DISPO# < 0 Then
                  EXCES# = Abs(DISPO#)
                  DISPO# = 0
               End If
               ZZ$ = REGBLAN$("LILICRE")
               Call REPLA(ZZ$, MKI$(NCI%), 1, 2)
               Call REPLA(ZZ$, MKD$(CRELIM#), 3, 8)
               Call REPLA(ZZ$, MKD$(DEUCUECO#(NCI%)), 11, 8)
               Call REPLA(ZZ$, MKD$(CHEPEAC#(NCI%)), 19, 8)
               Call REPLA(ZZ$, MKD$(DISPO#), 27, 8)
               Call REPLA(ZZ$, MKD$(EXCES#), 35, 8)
               jj& = jj& + 1
               Call GRAREG("LILICRE", ZZ$, jj&)
            End If
         End If
      End If
   Next IL&
   '
   Call SETULTREG("LILICRE", jj&)
   Call CIERRARCH("LILICRE")
   '
   OPSALI$ = "IMPRE"
   On Error GoTo 40
   If AMACLI1.Option1(0).Value = True Then OPSALI$ = "SCREEN"
   GoTo 41
   '
40 Resume 41
41 On Error GoTo 50
   If COGEVENT.Option1.Value = True Then OPSALI$ = "SCREEN"
   GoTo 51
   '
50 Resume 51
51 On Error GoTo 0
   If OPSALI$ = "SCREEN" Then
        Call CONECRUN("*COLICRE", "Consulta de Límites de Crédito")
      Else
        Call CONECRUN("*LILICRE", "Listado de Límites de Crédito")
   End If
   Screen.MousePointer = 1
   '
End Sub


