Attribute VB_Name = "CONGESCO"
Dim TICUE%(32767)

Sub FAPEPAG1()
    '
    Screen.MousePointer = 11
    '
    NCMIN% = 32767: NCMAX% = 0
    For U& = 1 To 32767: TICUE%(U&) = 1: Next U&
    FIN& = ULTREG&("ACREDOR")
    '
    For U& = 1 To FIN&
        Call TRACE(20, U&, FIN&)
        NCX% = CVI(Left$(REGLEIDO$("ACREDOR", U&), 2))
        If NCX% > 0 Then
            If NCX% > NCMAX% Then NCMAX% = NCX%
            If NCX% < NCMIN% Then NCMIN% = NCX%
            TICUE%(NCX%) = 0
        End If
    Next U&
    FAPEPAGO.Text10.TEXT = TRIM$(Str$(NCMIN%))
    FAPEPAGO.Text9.TEXT = RASOCLI$((-1) * NCMIN%)
    FAPEPAGO.Text8.TEXT = TRIM$(Str$(NCMAX%))
    FAPEPAGO.Text7.TEXT = RASOCLI$((-1) * NCMAX%)
    '
    FLIM = HOY(HO$)
    FIN& = ULTREG&("CUECORR")
    RFF$ = ""
    '
    For U& = 1 To FIN&
      Call TRACE(20, U&, FIN&)
      X$ = REGLEIDO$("CUECORR", U&)
      NCX% = CVI(Mid$(X$, 5, 2))
      If NCX% > 0 Then
        If TICUE%(NCX%) = 0 Then
          SALDEU$ = Mid$(X$, 139, 60)
          ATX% = 0
          For KKI% = 1 To 6
            FEVE = CVI(Mid$(SALDEU$, 10 * KKI% - 9, 2))
            SALDO# = CVD(Mid$(SALDEU$, 10 * KKI% - 7, 8))
            If Abs(SALDO#) > 0 Then
              ATX% = 1
              If FEVE > FLIM Then FLIM = FEVE
            End If
          Next KKI%
          If ATX% > 0 Then
            RFF$ = RFF$ + MKS$(U&)
          End If
        End If
      End If
    Next U&
    FAPEPAGO.Text6.TEXT = FECHATEX$(FLIM)
    FAPEPAGO.Text6.Refresh
    '
    UJ& = 0
    FIN& = Len(RFF$)
    For U& = 1 To Len(RFF$) Step 4
        Call TRACE(20, U&, FIN&)
        REG& = CVS(Mid$(RFF$, U&, 4))
        X$ = REGLEIDO$("CUECORR", REG&)
        SALDEU$ = Mid$(X$, 139, 60)
        For KKI% = 1 To 6
           FVI% = CVI(Mid$(SALDEU$, 10 * KKI% - 9, 2))
           SALDO# = CVD(Mid$(SALDEU$, 10 * KKI% - 7, 8))
           If FVI% > 0 Then
              If Abs(SALDO#) >= 0.005 Then
                 FF% = CVI(Left$(X$, 2))
                 NC% = CVI(Mid$(X$, 5, 2))
                 NFS$ = Mid$(X$, 7, 18)
                 IMORIG# = Abs(CVD(Mid$(X$, 115, 8)) - CVD(Mid$(X$, 123, 8)))
                 ES$ = "    "
                 If SALDO# < 0 Then
                      ES$ = " CR "
                    Else
                      If FVI% <= HOO% + 8 Then ES$ = " ** "
                      If FVI% <= HOO% Then ES$ = "!!!!"
                 End If
                 '
                 Z$ = REGBLAN$("FAPEPAG")
                 Call REPLA(Z$, MKI$(NC%), 1, 2)
                 Call REPLA(Z$, NFS$, 3, 18)
                 Call REPLA(Z$, MKI$(FF%), 21, 2)
                 Call REPLA(Z$, MKI$(FVI%), 23, 2)
                 Call REPLA(Z$, MKD$(SALDO#), 25, 8)
                 Call REPLA(Z$, MKD$(IMORIG#), 33, 8)
                 Call REPLA(Z$, ES$, 41, 4)
                 '
                 UJ& = UJ& + 1
                 Call GRAREG("FAPEPAG", Z$, UJ&)
                 '
              End If
           End If
        Next KKI%
    Next U&
    '
    Call SETULTREG("FAPEPAG", UJ&)
    Screen.MousePointer = 1
    '
End Sub
'

Sub FAPEPAG2()
    '
    Screen.MousePointer = 11
    FLIM = FECHANUM(FAPEPAGO.Text6.TEXT): If FLIM < HOY(HOS$) Then FLIM = HOY(HOS$)
    NCMIN% = Val(FAPEPAGO.Text10.TEXT)
    NCMAX% = Val(FAPEPAGO.Text8.TEXT)
    '
    MODORDE% = 1
    If FAPEPAGO.Option1(2) = True Then MODORDE% = 2
    If FAPEPAGO.Option1(3) = True Then MODORDE% = 3
    '
    FIN& = ULTREG&("FAPEPAG")
    UJ& = 0
    FAPEPAGO.List1.Clear
    '
    For U& = 1 To FIN&
        '
        Call TRACE(20, U&, FIN&)
        X$ = REGLEIDO$("FAPEPAG", U&)
        NC% = CVI(Left$(X$, 2))
        '
        If NC% >= NCMIN% Then
          If NC% <= NCMAX% Then
            FVI% = CVI(Mid$(X$, 23, 2))
            If FVI% <= FLIM Then
               '
               TLISTA$ = Space$(36)
               If MODORDE% = 1 Then
                    Call REPLA(TLISTA$, CSTRING$(MKI$(NC%), 1, 6, 0, 2), 1, 6)
                 ElseIf MODORDE% = 2 Then
                    Call REPLA(TLISTA$, RASOCLI$((-1) * NC%), 1, 30)
                 ElseIf MODORDE% = 3 Then
                    Call REPLA(TLISTA$, CSTRING$(MKI$(FVI%), 1, 6, 0, 2), 1, 6)
               End If
               Call REPLA(TLISTA$, CSTRING$(MKS$(U&), 3, 6, 0, 2), 31, 6)
               FAPEPAGO.List1.AddItem TLISTA$
               '
            End If
          End If
        End If
        '
    Next U&
    '
    FIN& = FAPEPAGO.List1.ListCount
    '
    For U& = 1 To FIN&
        '
        Call TRACE(20, U&, FIN&)
        FAPEPAGO.List1.ListIndex = U& - 1
        REG& = Val(Mid$(FAPEPAGO.List1.TEXT, 31, 8))
        X$ = REGLEIDO$("FAPEPAG", REG&)
        '
        NC% = CVI(Left$(X$, 2))
        NFS$ = Mid$(X$, 3, 18)
        FF% = CVI(Mid$(X$, 21, 2))
        FVI% = CVI(Mid$(X$, 23, 2))
        SALDEU# = CVD(Mid$(X$, 25, 8))
        IMORIG# = CVD(Mid$(X$, 33, 8))
        ES$ = Mid$(X$, 41, 4)
        '
        Y$ = REGBLAN$("FAPEPAGX")
        Call REPLA(Y$, MKI$(NC%), 1, 2)
        Call REPLA(Y$, RASOCLI$((-1) * NC%), 3, 30)
        Call REPLA(Y$, TEL1CLI$((-1) * NC%), 33, 16)
        Call REPLA(Y$, NFS$, 49, 18)
        Call REPLA(Y$, MKI$(FF%), 67, 2)
        Call REPLA(Y$, MKD$(IMORIG#), 69, 8)
        Call REPLA(Y$, MKI$(FVI%), 77, 2)
        Call REPLA(Y$, MKD$(SALDEU#), 79, 8)
        Call REPLA(Y$, ES$, 95, 4)
        '
        UJ& = UJ& + 1
        Call GRAREG("FAPEPAGX", Y$, UJ&)
        '          '
    Next U&
    '
    Call SETULTREG("FAPEPAGX", UJ&)
    Call CIERRARCH("FAPEPAGX")
    '
    Call HEADERS("FAPEPAGX", "")
    Call HEADERS("FAPEPAGX", "Vencimiento Hasta: " + FECHATEX$(FLIM))
    Call HEADERS("FAPEPAGX", "Rango de Cuentas: " + TRIM$(Str$(NCMIN%)) + " - " + TRIM$(Str$(NCMAX%)))
    '
    Screen.MousePointer = 1
    '
End Sub
'

Sub GELISALPRO()
    '
    HOXX = HOY(HO$)
    Dim DEUDA#(32767), SALDO#(32767)
    '
    Screen.MousePointer = 11
    Call HEADERS("LISAPRO", "")
    Call HEADERS("LISAPRO", "Generado el: " + HO$ + " - " + Left$(Time$, 5) + " hs.")
    '
    FIN& = ULTREG&("PROVED1")
    For U& = 1 To FIN&
       Call TRACE(20, U&, FIN&)
       PP1$ = REGLEIDO$("PROVED1", U&)
       NC% = CVI(Left$(PP1$, 2))
       If NC% > 0 Then
          PP2$ = REGLEIDO$("PROVED2", U&)
          TICUE%(NC%) = Asc(Mid$(PP2$, 119, 1))
          If TICUE%(NC%) < 0 Or TICUE%(NC%) > 1 Then
             Call MENSERR(24, "Tipo de Cuenta no Válido\para Cuenta " + TRIM$(Str$(NC%)) + " " + TRIM$(Mid$(PP2$, 3, 30)) + ".")
             TICUE%(NC%) = 0
          End If
       End If
    Next U&
    '
    Screen.MousePointer = 11
    FIN& = ULTREG&("CUECORR")
    '
    For IX& = 1 To FIN&
       Call TRACE(20, IX&, FIN&)
       SS$ = REGLEIDO$("CUECORR", IX&)
       SS4$ = Mid$(SS$, 5, 2)
       SD1$ = Mid$(SS$, 115, 8)
       SH1$ = Mid$(SS$, 123, 8)
       SALDEU$ = Mid$(SS$, 139, 60)
       '
       NC% = CVI(Mid$(SS$, 5, 2))
       If NC% > 0 Then
          If TICUE%(NC%) = 0 Then   ' saltea cuentas de gastos
             TOT# = CVD(SH1$) - CVD(SD1$)
             SALDO#(NC%) = SALDO#(NC%) + TOT#
             '
             For KKI% = 1 To 6
                FVI% = CVI(Mid$(SALDEU$, 10 * KKI% - 9, 2))
                SALPAR# = CVD(Mid$(SALDEU$, 10 * KKI% - 7, 8))
                If FVI% > 0 Then
                  If Abs(SALPAR#) >= 0.005 Then
                    DEUDA#(NC%) = DEUDA#(NC%) + SALPAR#
                  End If
                End If
             Next KKI%
          End If
       End If
    Next IX&
    '
    CX& = ULTREG&("PROVED1")
    RELISA& = 0
    '
    For IL& = 1 To CX&
       Call TRACE(20, IL&, CX&)
       X$ = REGLEIDO$("PROVED1", IL&)
       NC% = CVI(Left$(X$, 2))
       If NC% > 0 Then
          ' ajustar saldo de la cuenta
          If Abs(SALDO#(NC%)) >= 0.005 Or Abs(DEUDA#(NC%)) >= 0.005 Then
             Y$ = REGBLAN$("LISAPRO")
             Call REPLA(Y$, MKI$(NC%), 1, 2)
             Call REPLA(Y$, MKD$(SALDO#(NC%)), 3, 8)
             Call REPLA(Y$, MKD$(DEUDA#(NC%)), 11, 8)
             If Abs(SALDO#(NC%) - DEUDA#(NC%)) > 0.005 Then
                Call REPLA(Y$, "ATENCION: Saldo <> Deuda", 19, 30)
             End If
             '
             RELISA& = RELISA& + 1
             Call GRAREG("LISAPRO", Y$, RELISA&)
          End If
       End If
    Next IL&
    '
    Call SETULTREG("LISAPRO", RELISA&)
    Call CIERRARCH("*.*")
    Screen.MousePointer = 1
    '
    If COGECOMP.Option1.Value = True Then
          Call FINAL("*LISAPRO")
        Else
          Call FINAL("*LIMPPRO")
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
                    TI% = Asc(Mid$(X$, 13, 1))
                    If TI% < 0 Or TI% > 7 Then TI% = 2 ' default factura
                    If TI% > 0 Then
                       If TI% < 7 Then
                          NETO# = (Int(100 * (CVD(Mid$(X$, 15, 8)) + CVD(Mid$(X$, 57, 8))) + 0.5)) / 100
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
    J& = 0
    FIN& = RESUESTA.List1.ListCount
    '
    For U& = 1 To FIN&
        RESUESTA.List1.ListIndex = U& - 1
        NC% = Val(Mid$(RESUESTA.List1.TEXT, 14, 6))
        TONEVEN# = Val(Mid$(RESUESTA.List1.TEXT, 28, 12))
        PORPAR = Val(Mid$(RESUESTA.List1.TEXT, 20, 8))
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
    Call FINAL("*RANKCLI")
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
    FIN& = ULTREG&("MOVISTO")
    TOTAGEN# = 0
    For U& = 1 To FIN&
        Call TRACE(20, U&, FIN&)
        X$ = REGLEIDO$("MOVISTO", U&)
        FF% = CVI(Left$(X$, 2))
        If FF% >= DES% Then
           If FF% <= HAS% Then
              CI% = CVI(Mid$(X$, 3, 2))
              If CI% > 0 Then
                 If CI% <= 32767 Then
                    TIMOS$ = UCase$(Mid$(X$, 33, 2))
                    If TIMOS$ = "FC" Or TIMO$ = "NC" Then
                       NETO# = (CVD(Mid$(X$, 96, 8)) - CVD(Mid$(X$, 88, 8))) * CVD(Mid$(X$, 73, 8))
                       TOTVENTA#(CI%) = TOTVENTA#(CI%) + NETO#
                       TOTAGEN# = TOTAGEN# + NETO#
                    End If
                 End If
              End If
           End If
        End If
    Next U&
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
        CI% = Val(Mid$(RESUESTA.List1.TEXT, 14, 6))
        TONEVEN# = Val(Mid$(RESUESTA.List1.TEXT, 28, 12))
        PORPAR = Val(Mid$(RESUESTA.List1.TEXT, 20, 8))
        PORACU = PORACU + PORPAR
        '
        Z$ = REGBLAN$("RANKPIE")
        J& = J& + 1: JI% = J&
        Call REPLA(Z$, MKI$(JI%), 1, 2)
        Call REPLA(Z$, MKI$(CI%), 3, 2)
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
    Call FINAL("*RANKPIE")
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
    FIN& = ULTREG&("MOVISTO")
    Screen.MousePointer = 11
    TOTAGEN# = 0
    For U& = 1 To FIN&
        Call TRACE(20, U&, FIN&)
        X$ = REGLEIDO$("MOVISTO", U&)
        FF% = CVI(Left$(X$, 2))
        If FF% >= DES% Then
           If FF% <= HAS% Then
              CI% = CVI(Mid$(X$, 3, 2))
              If CI% > 0 Then
                 If CI% <= 32767 Then
                    TIMOS$ = UCase$(Mid$(X$, 33, 2))
                    If TIMOS$ = "FC" Or TIMO$ = "NC" Then
                       CANTI = CVD(Mid$(X$, 96, 8)) - CVD(Mid$(X$, 88, 8))
                       NETO# = CVD(Mid$(X$, 73, 8)) * CANTI
                       TOTCAN(CI%) = TOTCAN(CI%) + CANTI
                       TOTVENTA#(CI%) = TOTVENTA#(CI%) + NETO#
                    End If
                 End If
              End If
           End If
        End If
    Next U&
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

Sub DESCACU()
    '
    Screen.MousePointer = 11
    FLIM% = FECHANUM(DESCACUECO.Text1.TEXT)
    FIN& = ULTREG&("CUECORR")
    Dim CCX2 As String * 300
    Dim CCX3 As String * 300
    LU$ = LUDAT$
    Dim SARRAS#(32767)
    '
    FIN& = ULTREG&("PROVED1")
    For U& = 1 To FIN&
       Call TRACE(20, U&, FIN&)
       PP1$ = REGLEIDO$("PROVED1", U&)
       NC% = CVI(Left$(PP1$, 2))
       If NC% > 0 Then
          PP2$ = REGLEIDO$("PROVED2", U&)
          TICUE%(NC%) = Asc(Mid$(PP2$, 119, 1))
          If TICUE%(NC%) < 0 Or TICUE%(NC%) > 1 Then
             Call MENSERR(24, "Tipo de Cuenta no Válido\para Cuenta " + TRIM$(Str$(NC%)) + " " + TRIM$(Mid$(PP2$, 3, 30)) + ".")
             TICUE%(NC%) = 0
          End If
       End If
    Next U&
    '
    N1% = FILENBR%("CUECORR")
    N2% = FILEOPEN%(LU$ + "CUECORR.OLD", 0, 300)
    N3% = FILEOPEN%(LU$ + "TRANSIT.$$$", 0, 300)
    '
    On Error GoTo 10
    Lock #N1%
    Lock #N2%
    Lock #N3%
    On Error GoTo 0
    GoTo 20
    '
10  Resume 11
11  On Error GoTo 0
    Screen.MousePointer = 1
    Call MENSERR(24, "Proceso no Realizable.\Archivos de Datos en Uso.\  \Cierre todas las Aplicaciones\y Re-Inicie el Proceso.")
    Call FINAL("")
    Exit Sub
    '
20  I2& = LOF(N2%) / 300
    I3& = 0
    '
    FIN& = ULTREG&("CUECORR")
    For U& = 1 To FIN&
      Call TRACE(20, U&, FIN&)
      X$ = REGLEIDO$("CUECORR", U&)
      FF% = CVI(Left$(X$, 2))
      NCX% = CVI(Mid$(X$, 5, 2))
      If NCX% > 0 Then
         IMPO# = CVD(Mid$(X$, 115, 8)) - CVD(Mid$(X$, 123, 8))
         If TICUE%(NCX%) <> 0 Then
            IMPO# = 0
         End If
         If FF% < FLIM% Then
             I2& = I2& + 1
             LSet CCX2$ = X$
             Put #N2%, I2&, CCX2$
           Else
             SARRAS#(NCX%) = SARRAS(NCX%) - IMPO#
             I3& = I3& + 1
             LSet CCX3$ = X$
             Put #N3%, I3&, CCX3$
         End If
      End If
    Next U&
    '
    Unlock #N2%: Close #N2%
    Unlock #N1%: Call CIERRARCH("CUECORR")
    '
    Call BLANARCH("CUECORR")
    '
    For U& = 1 To I3&
      Call TRACE(20, U&, I3&)
      Get #N3%, U&, CCX3$
      X$ = CCX3$
      Call REGAPP("CUECORR", X$)
    Next U&
    '
    Call CIERRARCH("CUECORR")
    Unlock #N3%
    Close #N3%
    '
    CX& = ULTREG&("PROVED1")
    For IL& = 1 To CX&
       Call TRACE(20, IL&, CX&)
       X$ = REGLEIDO$("PROVED1", IL&)
       NC% = CVI(Left$(X$, 2))
       If NC% > 0 Then
          Y$ = REGLEIDO$("PROVED2", IL&)
          Call REPLA(Y$, MKD$(SARRAS#(NC%)), 109, 8)
          Call GRAREG("PROVED2", Y$, IL&)
       End If
    Next IL&
    '
    Call CIERRARCH("*.*")
    Call FINAL("")
    Exit Sub
    '
End Sub
'

