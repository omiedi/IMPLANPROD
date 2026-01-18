Attribute VB_Name = "RESCUPRO_RUTS"
Sub CARESCUEPRO()
    '
    If ATENTI% = 1 Then
       Exit Sub
    End If
    '
    Screen.MousePointer = 11
    NC% = XVALO(RESCUPRO.Text1.TEXT)
    If NC% > 0 Then
        NC1% = (-1) * NC%
        RESCUPRO.Text2.TEXT = RASOCLI$(NC1%)
        RESCUPRO.Text3.TEXT = DOMICLI$(NC1%)
        RESCUPRO.Text4.TEXT = CPOSCLI$(NC1%)
        RESCUPRO.Text5.TEXT = LOCACLI$(NC1%)
        RESCUPRO.Text7.TEXT = TEL1CLI$(NC1%)
        RESCUPRO.Text10.TEXT = CSTRING$(MKD$(SALDCLI#(NC1%)), 4, 13, 2, 2)
        TCUE% = TICUEPRO%(NC1%)
        If TCUE% = 1 Then
            RESCUPRO.Text6.TEXT = "GASTOS"
            RESCUPRO.Picture1.Height = 5340
            RESCUPRO.List1.Height = 5005
          Else
            RESCUPRO.Text6.TEXT = "CORRIENTE"
            RESCUPRO.Picture1.Height = 2430
            RESCUPRO.List1.Height = 2085
        End If
        RESCUPRO.Refresh
    End If
    '
100 RESCUPRO.List1.Clear
    RESCUPRO.List1.Refresh
    RESCUPRO.List2.Clear
    RESCUPRO.List2.Refresh
    '
    MAXDIAL& = ULTREG&("CUECORR")
    '
    RFF$ = RECFILT$("CUECORR", 4, MKI$(NC%))
    '
    RESCUPRO.LINDICE.Clear
    For U& = 1 To Len(RFF$) Step 4
        REG& = CVS(Mid$(RFF$, U&, 4))
        X$ = REGLEIDO$("CUECORR", REG&)
        If CVI(Mid$(X$, 5, 2)) <> NC% Then
            X$ = REGLEIDO$("CUECORR", 1)
            Call GRAREG("CUECORR", X$, 1)
            Call CIERRARCH("CUECORR")
            GoTo 100
        End If
        FF% = CVI(Left$(X$, 2))
        Z$ = CSTRING$(MKI$(FF%), 1, 6, 0, 2) + CSTRING$(MKS$(REG&), 4, 8, 0, 2)
        RESCUPRO.LINDICE.AddItem Z$
    Next U&
    '
    SALDO# = 0: SARRAS# = 0
    For U& = 1 To RESCUPRO.LINDICE.ListCount
        RESCUPRO.LINDICE.ListIndex = U& - 1
        REG& = XVALO(Mid$(RESCUPRO.LINDICE.TEXT, 7, 8))
        If REG& > 0 Then
            If REG& > MAXDIAL& Then GoTo 990
            X$ = REGLEIDO$("CUECORR", REG&)
            TIX$ = Mid$(X$, 7, 2)
            SG% = (-1)
            TIPOS$ = "Factura": If TIX$ = "AS" Then TIPOS$ = "Ajuste de Saldo"
                                If TIX$ = "TF" Then TIPOS$ = "Ticket-Factura"
                                If TIX$ = "RH" Then TIPOS$ = "Rec.Honorarios"
                                If TIX$ = "ND" Then TIPOS$ = "Nota de Debito"
                                If TIX$ = "CP" Then TIPOS$ = "Crédito Propio"
                                If TIX$ = "CO" Then TIPOS$ = "Comprobante"
                                If TIX$ = "OP" Then TIPOS$ = "Orden de Pago": SG% = 1
                                If TIX$ = "FF" Then TIPOS$ = "O/P Fondo Fijo": SG% = 1
                                If TIX$ = "NC" Then TIPOS$ = "Nota de Credito": SG% = 1
                                If TIX$ = "DP" Then TIPOS$ = "Debito Propio": SG% = 1
                                If TIX$ = "CC" Then TIPOS$ = "Comp.Ctas.": SG% = 1
                                '
            IMPO# = CVD(Mid$(X$, 115, 8)) - CVD(Mid$(X$, 123, 8))
            GoTo 890
        End If
        GoTo 990
        '
890     SALDO# = SALDO# - IMPO#
        FEX = CVI(Mid$(X$, 1, 2))
        NROX$ = Mid$(X$, 12, 13)
        Z$ = Space$(96)
        '
        Call REPLA(Z$, FECHATEX$(FEX), 1, 8)
        Call REPLA(Z$, TIPOS$, 11, 16)
        Call REPLA(Z$, NROX$, 27, 13)
        IMPTEX$ = CSTRING$(MKD$(Abs(IMPO#)), 3, 11, 2, 2)
        SALTEX$ = CSTRING$(MKD$(SALDO#), 3, 13, 2, 2)
        If IMPO# > 0 Then
            Call REPLA(Z$, IMPTEX$, 40, 11)
          ElseIf IMPO# < 0 Then
            Call REPLA(Z$, IMPTEX$, 53, 11)
        End If
        Call REPLA(Z$, SALTEX$, 64, 13)
        Call REPLA(Z$, TRIM$(Str$(REG&)), 81, 8)
        RESCUPRO.List1.AddItem Z$
        '
990 Next U&
    '
991 SARRAS# = SALDO#
    RESCUPRO.List1.ListIndex = RESCUPRO.List1.ListCount - 1
    RESCUPRO.List1.Refresh
    '
    If TCUE% <> 0 Then GoTo 999
    '
    RESCUPRO.LINDICE.Clear
    For U& = 1 To Len(RFF$) Step 4
        REG& = CVS(Mid$(RFF$, U&, 4))
        X$ = REGLEIDO$("CUECORR", REG&)
        SALDEU$ = Mid$(X$, 139, 60)
        For KKI% = 1 To 6
           FF% = CVI(Mid$(SALDEU$, 10 * KKI% - 9, 2))
           SALDO# = CVD(Mid$(SALDEU$, 10 * KKI% - 7, 8))
           If FF% > 0 Then
              If Abs(SALDO#) >= 0.005 Then
                 Z$ = CSTRING$(MKI$(FF%), 1, 6, 0, 2) + CSTRING$(MKS$(REG&), 4, 8, 0, 2) + CSTRING$(MKI$(KKI%), 1, 4, 0, 2)
                 RESCUPRO.LINDICE.AddItem Z$
              End If
           End If
        Next KKI%
    Next U&
    '
    SUMADEU# = 0: SUMAFAC# = 0
    For U& = 1 To RESCUPRO.LINDICE.ListCount
        RESCUPRO.LINDICE.ListIndex = U& - 1
        REG& = XVALO(Mid$(RESCUPRO.LINDICE.TEXT, 7, 8))
        KKI% = XVALO(Mid$(RESCUPRO.LINDICE.TEXT, 15, 4))
        If REG& > 0 Then
          If REG& <= MAXDIAL& Then
            If KKI% > 0 Then
              If KKI% <= 6 Then
                X$ = REGLEIDO$("CUECORR", REG&)
                TIX$ = Mid$(X$, 7, 2)
                SG% = (-1)
                TIPOS$ = "Factura": If TIX$ = "AS" Then TIPOS$ = "Ajuste de Saldo"
                                    If TIX$ = "TF" Then TIPOS$ = "Ticket-Factura"
                                    If TIX$ = "RH" Then TIPOS$ = "Rec.Honorarios"
                                    If TIX$ = "ND" Then TIPOS$ = "Nota de Debito"
                                    If TIX$ = "CP" Then TIPOS$ = "Crédito Propio"
                                    If TIX$ = "CO" Then TIPOS$ = "Comprobante"
                                    If TIX$ = "OP" Then TIPOS$ = "Orden de Pago": SG% = 1
                                    If TIX$ = "FF" Then TIPOS$ = "O/P Fondo Fijo": SG% = 1
                                    If TIX$ = "NC" Then TIPOS$ = "Nota de Credito": SG% = 1
                                    If TIX$ = "DP" Then TIPOS$ = "Debito Propio": SG% = 1
                                    If TIX$ = "CC" Then TIPOS$ = "Comp.Ctas.": SG% = 1
                                   '
                IMPO# = CVD(Mid$(X$, 115, 8)) - CVD(Mid$(X$, 123, 8))
                SUMAFAC# = SUMAFAC# + IMPO#
                IMPO# = Abs(IMPO#)
                '
                SALDO# = CVD(Mid$(X$, 139 + 10 * KKI% - 8, 8))
                SUMADEU# = SUMADEU# + SALDO#
                '
                NROX$ = Mid$(X$, 12, 13)
                FEMI = CVI(Mid$(X$, 1, 2))
                FEVE = CVI(Mid$(X$, 139 + 10 * KKI% - 10, 2))
                '
                IMPTEX$ = CSTRING$(MKD$(IMPO#), 3, 12, 2, 2)
                SALTEX$ = CSTRING$(MKD$(SALDO#), 3, 12, 2, 2)
                '
                Z$ = Space$(80)
                Call REPLA(Z$, TIPOS$, 1, 16)
                Call REPLA(Z$, NROX$, 17, 13)
                Call REPLA(Z$, FECHATEX$(FEMI), 32, 8)
                Call REPLA(Z$, IMPTEX$, 40, 12)
                Call REPLA(Z$, FECHATEX$(FEVE), 56, 8)
                Call REPLA(Z$, SALTEX$, 65, 12)
                '
                RESCUPRO.List2.AddItem Z$
                '
              End If
            End If
          End If
        End If
        '
    Next U&
    '
    If Abs(SARRAS# - SUMADEU#) >= 0.005 Then
       ATENTI% = 1
       TTXX$ = "Se ha Detectado una Diferencia\en la Composición de la Deuda\de la Cuenta Elegida.\   \Puede Intentarse la Reconstucción\Mediante un Proceso de Re-Aplicación.\  \Este Proceso Reconstruye\la Composición de la Deuda de la Cuenta Elegida\en Base a la Aplicación de Pagos y Créditos\Registrada en el Sistema.\  \Desea Cancelar o Continuar el Proceso ?\\Cancelar\Continuar"
       If COMALTER%(TTXX$) = 2 Then
           Call REAPLICREPRO(NC%)
           GoTo 100
         Else
           ATENTI% = 0
           Call CIERRARCH("*.*")
           Call FINAL("")
       End If
    End If
    '
    RESCUPRO.List2.ListIndex = RESCUPRO.List2.ListCount - 1
    RESCUPRO.List2.Refresh
    RESCUPRO.Label5.Caption = CSTRING$(MKD$(SUMADEU#), 4, 14, 2, 2)
    RESCUPRO.Label5.Refresh
    RESCUPRO.Label6.Caption = CSTRING$(MKD$(SUMADEU#), 4, 14, 2, 2)
    RESCUPRO.Label6.Refresh
    '
999 Screen.MousePointer = 1
    '
End Sub
'
Sub GERESCUEPRO()
    '
    Screen.MousePointer = 11
    '
    DES% = FECHANUM(RESCUEPRO.Text6.TEXT)
    HAS% = FECHANUM(RESCUEPRO.Text5.TEXT)
    NCMIN% = XVALO(RESCUEPRO.Text10.TEXT)
    NCMAX% = XVALO(RESCUEPRO.Text8.TEXT)
    '
    MODORDE% = 1
    If RESCUEPRO.Option1(2) = True Then MODORDE% = 2
    '
    FIN& = ULTREG&("PROVED1")
    UJ& = 0
    RESCUEPRO.List1.Clear
    Call BLANARCH("RESCUPRO")
    '
    FIN& = ULTREG&("PROVED1")
    For U& = 1 To FIN&
        '
        Call TRACE(20, U&, FIN&)
        X$ = REGLEIDO$("PROVED1", U&)
        NC% = CVI(Left$(X$, 2))
        '
        If NC% >= NCMIN% Then
          If NC% <= NCMAX% Then
            '
            TLISTA$ = Space$(36)
            If MODORDE% = 1 Then
                 Call REPLA(TLISTA$, CSTRING$(MKI$(NC%), 1, 6, 0, 2), 1, 6)
               ElseIf MODORDE% = 2 Then
                 Call REPLA(TLISTA$, RASOCLI$((-1) * NC%), 1, 30)
            End If
            Call REPLA(TLISTA$, CSTRING$(MKI$(NC%), 1, 6, 0, 2), 31, 6)
            RESCUEPRO.List1.AddItem TLISTA$
            '
          End If
        End If
        '
    Next U&
    '
    FIN& = RESCUEPRO.List1.ListCount
    '
    For U& = 1 To FIN&
        '
        Call TRACE(20, U&, FIN&)
        RESCUEPRO.List1.ListIndex = U& - 1
        NC% = XVALO(Mid$(RESCUEPRO.List1.TEXT, 31, 6))
        RESCUPRO.Text1.TEXT = TRIM$(Str$(NC%))
        Call CARESCUEPRO
        Call FRESHALL
        '
        REBLA$ = REGBLAN$("RESCUPRO")
        CLIE$ = TRIM$(RESCUPRO.Text2.TEXT) + " (" + TRIM$(Str$(NC%)) + ")"
        DOMIE$ = TRIM$(RESCUPRO.Text3.TEXT)
        LOCAE$ = TRIM$(RESCUPRO.Text5.TEXT)
        If TRIM$(RESCUPRO.Text4.TEXT) <> "" Then
            LOCAE$ = TRIM$(RESCUPRO.Text4.TEXT) + " - " + LOCAE$
        End If
        TELEE$ = "Telefax: " + TRIM$(RESCUPRO.Text7.TEXT)
        '
        Call REPLA(REBLA$, MKI$(NC%), 1, 2)
        Call REPLA(REBLA$, "Sres.", 97, 6)
        Call REPLA(REBLA$, CLIE$, 103, 48)
        Call REPLA(REBLA$, DOMIE$, 151, 48)
        Call REPLA(REBLA$, LOCAE$, 199, 32)
        Call REPLA(REBLA$, TELEE$, 231, 24)
        '
        YATRANSPO% = 0: ITRANSPO# = 0
        For V& = 1 To RESCUPRO.List1.ListCount
            RESCUPRO.List1.ListIndex = V& - 1
            REGI$ = RESCUPRO.List1.TEXT
            FECHA$ = Left$(REGI$, 8)
            FF% = FECHANUM(FECHA$)
            If FF% < DES% Then
                ITRANSPO# = XVALO(Mid$(REGI$, 64, 13))
            End If
            If FF% >= DES% Then
                If FF% <= HAS% Then
                    If YATRANSPO% = 0 Then
                        REGJ$ = REBLA$
                        FEX = DES%
                        Call REPLA(REGJ$, FECHATEX$(FEX), 3, 8)
                        Call REPLA(REGJ$, "Saldo Anterior", 11, 32)
                        Call REPLA(REGJ$, MKD$(ITRANSPO#), 63, 8)
                        Call REGAPP("RESCUPRO", REGJ$)
                        Call REGAPP("RESCUPRO", REBLA$)
                        YATRANSPO% = 1
                    End If
                    '
                    REGJ$ = REBLA$
                    DOCU$ = Mid$(REGI$, 11, 16)
                    NUDO$ = Mid$(REGI$, 27, 13)
                    DEBE# = XVALO(Mid$(REGI$, 40, 11))
                    HABER# = XVALO(Mid$(REGI$, 53, 11))
                    SALDO# = XVALO(Mid$(REGI$, 64, 13))
                    '
                    Call REPLA(REGJ$, FECHA$, 3, 8)
                    Call REPLA(REGJ$, DOCU$, 11, 18)
                    Call REPLA(REGJ$, NUDO$, 29, 18)
                    Call REPLA(REGJ$, MKD$(DEBE#), 47, 8)
                    Call REPLA(REGJ$, MKD$(HABER#), 55, 8)
                    Call REPLA(REGJ$, MKD$(SALDO#), 63, 8)
                    Call REGAPP("RESCUPRO", REGJ$)
                    '
                End If
            End If
        Next V&
        '
        If YATRANSPO% = 1 Then
            Call REGAPP("RESCUPRO", String$(256, "="))
        End If
        '
    Next U&
    '
    Call CONECRUN("*RESCPRO", "Listado Resumen de Cuenta Proveedores")
    '
End Sub
'



