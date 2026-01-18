Attribute VB_Name = "FAC_FUNC"

Sub GRADEU(LU$, ARDEX$, RDX%, Tipo%, VA%, NRO, NC%, FFAC%, COPAGI%, TOTFAC#)
        '
        Dim FVI%(6), PLAII%(6), IM#(6), FEX(3)
        '
        IORIG# = TOTFAC#
        ARCHIDEU$ = ARDEX$
        PPXX1% = InStr(ARDEX$, "/")
        If PPXX1% > 0 Then
          ARCHIDEU$ = Left$(ARDEX$, PPXX1% - 1)
          IORIG# = XVALO(Mid$(ARDEX$, PPXX1% + 1))
        End If
        '
        If Tipo% = 0 Then
          If IORIG# < 0 Then
            Tipo% = 7
          End If
        End If
        '
        If Tipo% > 3 Then
            NUPAG% = 1
            FVI%(1) = FFAC%
            IM#(1) = TOTFAC#     ' IORIG#
            GoTo 410
        End If


      ' busca condicion de pago segun COPAGI%

200     Call BUCONPA(LU$, COPAGI%, CoPag$, FREF%, PLA$)



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
        If XVALO(Left$(FFS$, 2)) <= 15 Then FEBAS$ = "15" + Mid$(FFS$, 3)
        If XVALO(Left$(FFS$, 2)) > 15 Then FEBAS$ = "31" + Mid$(FFS$, 3)
305     If FREF% = 1 Then FEBAS$ = FFS$
        If FREF% = 4 Then FEBAS$ = "31" + Mid$(FFS$, 3)

        For i% = 1 To 6: FVI%(i%) = 0: Next i%

        For i% = 1 To 6
                PLAI% = Asc(Mid$(PLA$, i%, 1)): PLAII%(i%) = PLAI%
                If i% > 1 Then If (PLAII%(i%) < PLAII%(i% - 1) Or PLAII%(i%) <= 0) Then GoTo 310

                        FE$ = FEBAS$: GoSub 30500
                        FEX(1) = FEX(1) + PLAI%
306                     If FEX(1) > 31 Then FEX(2) = FEX(2) + 1: FEX(1) = FEX(1) - 31: GoTo 306
307                     If FEX(2) > 12 Then FEX(3) = FEX(3) + 1: FEX(2) = FEX(2) - 12: GoTo 307
                        GoSub 30580
                        '
1234                    FEVEX$ = FECHATEX$(FEXU)
                        If XVALO(Mid$(FEVEX$, 4, 2)) = 2 Then
                          If XVALO(Mid$(FEVEX$, 1, 2)) > 28 Then
                            FEXU = DIPOST(FEXU)
                            GoTo 1234
                          End If
                        End If
                        '
                FVI%(i%) = FEXU: NUPAG% = i%
        Next i%

310     FVIJ% = FVI%(NUPAG%): If NUPAG% < 2 Then FVIJ% = FFAC%

      ' determina importes segun numero de pagos

400     If NUPAG% < 1 Then NUPAG% = 1
        'IM#(1) = (Int(100 * TOTFAC# / NUPAG% + 0.5)) / 100
        IM#(1) = (Int(100 * IORIG# / NUPAG% + 0.5)) / 100

        For H% = 2 To NUPAG% - 1
                IM#(H%) = IM#(1)
        Next H%
        IM#(NUPAG%) = IORIG#

        For H% = 1 To NUPAG% - 1
                IM#(NUPAG%) = IM#(NUPAG%) - IM#(H%)
        Next H%
        '
        ICANCE# = IORIG# - TOTFAC#
        For H% = 1 To NUPAG%
          If ICANCE# > 0 Then
            If ICANCE# >= IM#(H%) Then
                ICANCE# = ICANCE# - IM#(H%)
                IM#(H%) = 0
              Else
                IM#(H%) = IM#(H%) - ICANCE#
                ICANCE# = 0
            End If
          End If
        Next H%
        '
410     For i% = 1 To NUPAG%
          If Abs(IM#(i%)) >= 0.005 Then
            TPXX% = Tipo%: If Tipo% = 7 Then TPXX% = 0
            ss$ = MKI$(NC%) + Chr$(TPXX%) + Chr$(VA%) + MKS$(NRO)
            ss$ = ss$ + MKI$(FVI%(i%)) + MKD$(IM#(i%))
            ss$ = ss$ + MKI$(FFAC%) + MKD$(IORIG#) + String$(2, 0) + String$(2, 255)
            Call REPLA(ss$, Chr$(MONEMI%), 31, 1)
            Call REPLA(ss$, Chr$(0), 32, 1) ' FILLER
            Call REGAPP(ARCHIDEU$, ss$)
            FECHAVENCI(i%) = FVI%(i%)
            IMPOVENCI#(i%) = IM#(i%)
          End If
        Next i%
        CAPLAVENCI% = NUPAG%
        '
        Call CIERRARCH(ARCHIDEU$)
        '
        If NUPAG% >= 1 Then
          On Error Resume Next
          FEXU = FVI%(1)
          FORFACLI.FVENCIM = FECHATEX$(FEXU)
          
          On Error GoTo 0
        End If
        '
        Exit Sub
        '
30500   A% = 0
        For II% = 1 To 3
            FEX(II%) = XVALO(Mid$(FE$, 3 * II% - 2, 2))
            If II% = 3 Then
                If FEX(II%) < 60 Then
                    FEX(II%) = FEX(II%) + 100
                End If
            End If
            If FEX(II%) < 1 Then A% = 2: GoTo 30595
        Next II%
        '
        If ((FEX(1) > 31) Or (FEX(2) > 12)) Then A% = 2: GoTo 30595
30580   FEXU = FEX(1) + 32 * FEX(2) + 32 * 13 * (FEX(3) - 80)
30595   Return


40000   End Sub
'
Sub BUCONPA(LU$, COPAGI%, CoPag$, FREF%, PLA$)
        '
        If ULTREG&("CONPAG") < 1 Then
                Call MENSERR(24, "No Existe Archivo de Cond. Pago en U.Logica '" + LU$ + "'")
                COPAGI% = 0: CoPag$ = "CONTADO" + Space$(28): FREF% = 1: PLA$ = String$(6, 0)
                Exit Sub
        End If

200     If COPAGI% < 0 Or COPAGI% > 255 Then
                Call MENSERR(24, "Numero de Cond. Pago Fuera de Rango")
                COPAGI% = 0: CoPag$ = "CONTADO" + Space$(28): FREF% = 1: PLA$ = String$(6, 0)
                GoTo 40000
        End If

300     If COPAGI% = 0 Then
                COPAGI% = 0: CoPag$ = "CONTADO" + Space$(28): FREF% = 1: PLA$ = String$(6, 0)
                GoTo 40000
        End If

400     For i& = 1 To ULTREG("CONPAG")
                LL$ = REGLEIDO$("CONPAG", i&)
                L1$ = Left$(LL$, 1)
                If Asc(L1$) = 0 Then
                        Call MENSERR(24, "Condicion de Pago Numero" + Str$(COPAGI%) + " no Figura en Tabla")
                        COPAGI% = 0: CoPag$ = "CONTADO" + Space$(28): FREF% = 1: PLA$ = String$(6, 0)
                        Exit Sub
                End If
                '
                If Asc(L1$) = COPAGI% Then
                    CoPag$ = Mid$(LL$, 2, 35)
                    FREF% = Asc(Mid$(LL$, 37, 1))
                    PLA$ = Mid$(LL$, 38, 6)
                    Exit Sub
                End If
        Next i&

        Call MENSERR(24, "Condicion de Pago Numero" + Str$(COPAGI%) + " no Figura en Tabla")
        COPAGI% = 0
        CoPag$ = "CONTADO" + Space$(28)
        FREF% = 1: PLA$ = String$(6, 0)
        '
40000   End Sub
   
Sub IMPUCRE1()
    '
    On Error Resume Next
    For KKI% = Forms.Count To 1 Step -1
      If Forms(KKI% - 1).Visible = True Then
        If Forms(KKI% - 1).BorderStyle = 3 Then
          Forms(KKI% - 1).Hide
        End If
      End If
    Next KKI%
    'FORECISTA.Hide
    On Error GoTo 0
    '
    Call FRESHALL
    FIMPUCRE.Show
    Call FRESHALL
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
    FIMPUCRE.LabeL15.Caption = ""
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
    NC% = XVALO(VALACT1$("DEUDOR1"))
    If NC% < 1 Then
        GoTo 999
    End If
    '
    FIMPUCRE.Label1.Caption = Str$(NC%)
    FIMPUCRE.Label2.Caption = rasocli$(NC%)
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
    On Error Resume Next
    RESCUDEU.LINDICE.Clear
    RECUEC00.LINDICE.Clear
    On Error GoTo 0
    '
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
        If Sgn(IMPO#) <> Sgn(SALDO#) Or Abs(SALDO#) > Abs(IMPO#) + 0.005 Then
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
        On Error Resume Next
        RESCUDEU.LINDICE.AddItem Z$
        RECUEC00.LINDICE.AddItem Z$
        On Error GoTo 0
    Next U&
    '
    SUMADEU# = 0: SUMAFAC# = 0: KQ& = 0
    On Error Resume Next
    FIN& = RECUEC00.LINDICE.ListCount
    FIN& = RESCUDEU.LINDICE.ListCount
    On Error GoTo 0
    '
    HAYVIEJOS% = 0: FELIVIE% = 0
    If InStr(UCase$(EMPREAC$), "DOSIVA") > 0 Then
      FELIVIE% = FECHANUM("31/08/01")
    End If
    '
    For U& = 1 To FIN&
        On Error Resume Next
        REG& = XVALO(Mid$(RESCUDEU.LINDICE.List(U& - 1), 7, 8))
        REG& = XVALO(Mid$(RECUEC00.LINDICE.List(U& - 1), 7, 8))
        On Error GoTo 0
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
                    NUSUCUVE% = CVI(Mid$(X$, 29, 2))
                    '
                    IMPTEX$ = CSTRING$(MKD$(IMPO#), 3, 14, 2, 2)
                    SALTEX$ = CSTRING$(MKD$(SALDO#), 3, 14, 2, 2)
                    DOCU$ = FORMDOC$(NROL&, TI%, VA%, NUSUCUVE%)
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
                    NUSUCUVE% = CVI(Mid$(X$, 29, 2))
                    '
                    IMPTEX$ = CSTRING$(MKD$(Abs(IMPO#)), 3, 14, 2, 2)
                    SALTEX$ = CSTRING$(MKD$(Abs(SALDO#)), 3, 14, 2, 2)
                    DOCU$ = FORMDOC$(NROL&, TI%, VA%, NUSUCUVE%)
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
                    '
                    If FEMI <= FELIVIE% Then
                      HAYVIEJOS% = 1
                    End If
                    '
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
    FIMPUCRE.Command1.Visible = False
    If HAYVIEJOS% > 0 Then
      FIMPUCRE.Command1.Visible = True
    End If
    '
    FIMPUCRE.Label25.Caption = Left$(RGSEL$, 18)
    FIMPUCRE.Label24.Caption = Mid$(RGSEL$, 19, 8)
    FIMPUCRE.Label23.Caption = Mid$(RGSEL$, 27, 14)
    FIMPUCRE.Label22.Caption = Mid$(RGSEL$, 41, 14)
    FIMPUCRE.LabeL15.Caption = CVS(Mid$(RGSEL$, 61, 4))
    '
    FIMPUCRE.Label14.Caption = Mid$(RGSEL$, 41, 14)
    FIMPUCRE.Label20.Caption = Mid$(RGSEL$, 41, 14)
    FIMPUCRE.Command5.Enabled = True
    '
    Exit Sub
    '
999 Call CIERRARCH("*.*")
    On Error Resume Next
    FORECISTA.Show
    On Error GoTo 0
    FIMPUCRE.Hide
    Call FRESHALL
    For KU% = 1 To Forms.Count
      On Error Resume Next
      If Forms(KU% - 1).Visible = True Then
        On Error GoTo 0
        Exit Sub
      End If
    Next KU%
    On Error GoTo 0
    Call FINAL("")
    '
End Sub
'
Sub IMPUCRE2()
    '
    CANCE# = CDbl(XVALO(FIMPUCRE.Label18.Caption))
    If CANCE# < 0.005 Then
        Exit Sub
    End If
    '
    REG202& = XVALO(FIMPUCRE.LabeL15.Caption)
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
      CANCE# = CDbl(XVALO(Mid$(X$, 70, 14)))
      If CANCE# >= 0.005 Then
        NREG& = XVALO(Mid$(X$, 121, 6))
        If NREG& > 0 Then
          If NREG& <= ULTREG&("SVD202") Then
            ss$ = REGLEIDO$("SVD202", NREG&)
            NUSUCUVE% = CVI(Mid$(ss$, 29, 2))
            DOCORI$ = FORMDOC$(CVS(Mid$(ss$, 5, 4)), Asc(Mid$(ss$, 3, 1)), Asc(Mid$(ss$, 4, 1)), NUSUCUVE%)
            ITOTORI# = CVD(Mid$(ss$, 21, 8))
            IVAORI# = (Int(ALIVA * ITOTORI# / (1 + ALIVA / 100) + 0.5)) / 100
            '
            SALANT# = CVD(Mid$(ss$, 11, 8))
            SALDEU# = CVD(Mid$(ss$, 11, 8)) - CANCE#
            Call REPLA(ss$, MKD$(SALDEU#), 11, 8)
            Call GRAREG("SVD202", ss$, NREG&)
            '
            NCI% = XVALO(FIMPUCRE.Label1.Caption)
            HOII% = HOY(HO$)
            RefApli$ = "Aplicacion Manual " + HO$
            FEDOAPLI% = FECHANUM(TRIM$(FIMPUCRE.Label24.Caption))
            DOCAPLI$ = TRIM$(FIMPUCRE.Label25.Caption)
            '
            Z$ = REGBLAN$("APLICOB")
            Call REPLA(Z$, MKI$(FEDOAPLI%), 1, 2)
            Call REPLA(Z$, MKI$(NCI%), 3, 2)
            Call REPLA(Z$, DOCAPLI$, 5, 18)
            Call REPLA(Z$, MKD$(CANCE#), 23, 8)
            Call REPLA(Z$, RefApli$, 31, 24)
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
    Call CIERRARCH("APLICOB")
    Call CIERRARCH("SVD202")
    VACON% = 0
    '
    RFF$ = RECFILT$("CARBIMON", 2, MKI$(NCI%))
    If Len(RFF$) > 0 Then
      TICAMBX# = XVALO(FIMPUCRE.TICAMBIS.TEXT)
      If TICAMBX# < 1 Then TICAMBX# = 1
      For U& = 1 To FIMPUCRE.List2.ListCount
        X$ = FIMPUCRE.List2.List(U& - 1)
        CANCE# = CDbl(XVALO(Mid$(X$, 70, 14)))
        If CANCE# >= 0.005 Then
          NREG& = XVALO(Mid$(X$, 121, 6))
          If NREG& > 0 Then
            If NREG& <= ULTREG&("SVD202") Then
              ss$ = REGLEIDO$("SVD202", NREG&)
              ABUSQUE$ = Mid$(ss$, 3, 6)
              TIDOAP% = 4
              If Left$(DOCAPLI$, 2) = "NC" Then TIDOAP% = 6
              If Left$(DOCAPLI$, 2) = "AJ" Then TIDOAP% = 7
              VADOAP% = Asc(Mid$(DOCAPLI$, 18, 1)) - 64
              If VADOAP% < 1 Or VADOAP% > 2 Then VADOAP% = 0
              NROO& = XVALO(Mid$(DOCAPLI$, 9, 8))
              SUC% = XVALO(Mid$(DOCAPLI$, 4, 4))
              '
              For JKLX& = 1 To Len(RFF$) Step 4
                RECABI& = CVS(Mid$(RFF$, JKLX&, 4))
                YCB$ = REGLEIDO$("CARBIMON", RECABI&)
                If Mid$(YCB$, 5, 6) = ABUSQUE$ Or Mid$(YCB$, 37, 6) = ABUSQUE$ Then
                  Call REPLA(YCB$, MKI$(FEDOAPLI%), 33, 2)
                  Call REPLA(YCB$, MKI$(NCI%), 35, 2)
                  Call REPLA(YCB$, Chr$(TIDOAP%), 37, 1)
                  Call REPLA(YCB$, Chr$(VADOAP%), 38, 1)
                  Call REPLA(YCB$, MKS$(NROO&), 39, 4)
                  Call REPLA(YCB$, MKI$(SUC%), 43, 2)
                  Call REPLA(YCB$, FORMDOC$(NROO&, 4, 1, SUC%), 45, 18)
                  Call REPLA(YCB$, MKI$(1), 63, 2) ' RECIBO SIEMPRE EN PESOS
                  Call REPLA(YCB$, MKD$(0), 65, 8)
                  Call REPLA(YCB$, MKD$(CANCE#), 73, 8)
                  Call REPLA(YCB$, MKD$(0), 81, 8)
                  Call REPLA(YCB$, MKD$(TICAMBX#), 89, 8)
                  Call REPLA(YCB$, MKD$(0), 97, 8)
                  Call REPLA(YCB$, MKD$(CANCE# / TICAMBX#), 105, 8)
                  Call REPLA(YCB$, MKD$(0), 113, 8)
                  '
                  If Mid$(YCB$, 5, 6) <> ABUSQUE$ Then ' APLICADO AL DEBITO POR D.C.
                    Call REPLA(YCB$, MKD$(0), 105, 8)  ' NO AJUSTA SALDO EN DOLARES
                  End If
                  '
                  Call REGAPP("CARBIMON", YCB$)
                  GoTo 835
                End If
              Next JKLX&
            End If
          End If
        End If
835   Next U&
    End If
    Call CIERRARCH("CARBIMON")
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
        NROL& = XVALO(Mid$(Z$, 13, 8))
        FEMII% = FECHANUM(Mid$(Z$, 23, 8))
        IPED# = CDbl(XVALO(Mid$(Z$, 31, 14)))
        FEVEI% = FECHANUM(Mid$(Z$, 48, 8))
        SALD# = CDbl(XVALO(Mid$(Z$, 56, 14)))
        ACRE# = CDbl(XVALO(Mid$(Z$, 70, 14)))
        DOCU$ = Mid$(Z$, 97, 18)
        REG2& = XVALO(Mid$(Z$, 121, 6))
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
        SUMACRE# = SUMACRE# + CDbl(XVALO(Mid$(Z$, 70, 14)))
    Next U&
    '
    FIMPUCRE.Label10.Caption = CSTRING$(MKD$(SUMACRE#), 4, 14, 2, 2)
    FIMPUCRE.Label10.Refresh
    FIMPUCRE.Label18.Caption = TRIM$(CSTRING$(MKD$(SUMACRE#), 4, 14, 2, 2))
    FIMPUCRE.Label18.Refresh
    ANTI# = CDbl(XVALO(FIMPUCRE.Label14.Caption)) - SUMACRE#
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
Sub CAREMPRE(LU$, A$())
      '
      LDT$ = LUDAT$
      n% = FILEOPEN%(LDT$ + "SVD147.EXE", 0, 41)
      Dim V1 As String * 41
      If LOF(n%) < 192 Then
          Close #n%: n% = 0
          LDT1$ = REPLACAR$(LDT$, "\", "/")
          Call MENSERR(24, "Falta Archivo de Datos de la Empresa\en Directorio '" + LDT1$ + "'.")
          Exit Sub
      End If
      '
      For KK% = 1 To 8
          Get #n%, KK%, V1$
          A$(KK%) = TRIM$(DESENCRI$(V1$))
      Next KK%
      Close #n%: n% = 0
      '
End Sub
