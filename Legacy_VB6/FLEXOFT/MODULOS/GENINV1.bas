Attribute VB_Name = "GENINV1"
Sub GENINV2()
    '
    FECHINVE$ = GENINVE.Text1.TEXT
    If FECHANUM(FECHINVE$) < 1 Then
        GENINVE.Text1.TEXT = HOS$
        GENINVE.Text1.SetFocus
        Exit Sub
    End If
    '
    Call GRACONTROL("INVENTAR", "FECHAINV", FECHINVE$)
    GENINVE.Text1.TEXT = FECHINVE$
    GENINVE.Text1.Refresh
    Call GRACONTROL("INVENTAR", "FEGENINV", "Generado el " + HOS$ + " - " + Left$(Time$, 5) + " hs.")
    FF% = FECHANUM(FECHINVE$)
    '
    Screen.MousePointer = 11
    '
    Dim COSU%(62), SUCU$(62), DISTKK%(62)
    NUSU% = 1
    For i& = 1 To ULTREG&("SUCURSAL")
      x$ = REGLEIDO$("SUCURSAL", i&)
      If i& <= 62 Then
        If Asc(Left$(x$, 1)) > 32 Then
          NUSU% = i&
          COSU%(i&) = Asc(Mid$(x$, 64, 1))
          SUCU$(i&) = Left$(x$, 20)
          DISTKK%(i&) = 1
          If GENINVE.Check5.Value <> 1 Then
            If UCase$(Mid$(x$, 62, 2) = "NO") Then
              DISTKK%(i&) = 0
            End If
          End If
        End If
      End If
    Next i&
    Call CIERRARCH("SUCURSAL")
    '
100 GENINVE.Caption = "Generacion de Inventarios al " + FECHINVE$
    '
    Call COPYSTRU("SALDEPOS", "SALDEPO")
    '
    VALORI% = 1
    TICO$ = "Costos: Informados"
    If GENINVE.Option3 = True Then
        VALORI% = 2: TICOS$ = "Costos: Calculados"
      ElseIf GENINVE.Option4 = True Then
        VALORI% = 3: TICOS$ = "Costos: S/Precios de Venta"
        Dim APRELI$(128), MONELI%(128)
        CALIS% = ULTREG&("LISTAS"): If CALIS% > 64 Then CALIS% = 64
        For KKI% = 1 To CALIS%
          ULK& = KKI%
          XX$ = REGLEIDO$("LISTAS", ULK&)
          MONELI%(KKI%) = Asc(Mid$(XX$, 50, 1))
          APRELI$(KKI%) = TRIM$(Str$(KKI%))
          While Len(APRELI$(KKI%)) < 3: APRELI$(KKI%) = "0" + APRELI$(KKI%): Wend
          APRELI$(KKI%) = "LIPRE" + APRELI$(KKI%)
        Next KKI%
      ElseIf GENINVE.Option5 = True Then
        VALORI% = 4: TICOS$ = "Costos: Costo de Suministros"
    End If
    '
    
' ********************* INICIALIZANDO ARCHIVO DE TRANSFERENCIA
    '
    GENINVE.Check1.Value = 2
    FIN& = NUMAS% + 1
    '
    'Call BLANARCH("SALDEPO")
    'N2% = FILEOPEN%("SALDEPO", 0, 256)
    'Dim SSQ As String * 256
    '
    FIN& = NUMAS% + 1
    For i& = 1 To FIN&
        Call TRACE(20, i&, FIN&)
        SSQ$ = String$(255, 0) + Chr$(1)
        Call GRAREG("SALDEPO", SSQ$, i&)
    Next i&
    Call SETULTREG("SALDEPO", FIN&)
    GENINVE.Check1.Value = 1
    '
' ******************************* RE-CALCULANDO STOCK AL  FFS$
    '
    GENINVE.Check2.Value = 2
    FIN& = ULTREG&("MOVISTO")
    '
    'N4% = FILEOPEN%("MOVISTO")
    'Dim MVS As String * 128
    '
    For i& = 1 To FIN&
      Call TRACE(20, i&, FIN&)
      MVS$ = REGLEIDO$("MOVISTO", i&)
        'Get #N4%, I&, MVS$
      If CVI(Left$(MVS$, 2)) <= FF% Then
        CI% = CVI(Mid$(MVS$, 3, 2))
        DEP% = Asc(Mid$(MVS$, 83, 1))
        LOT$ = Mid$(MVS$, 5, 16)
105     SUCX% = 1
        For U% = 1 To NUSU%
          If COSU%(U%) = DEP% Then SUCX% = U%: GoTo 110
        Next U%
        '
110     CANTI = CVD(Mid$(MVS$, 88, 8)) - CVD(Mid$(MVS$, 96, 8))
        CANTI = CANTI * DISTKK%(SUCX%)
        '
        If CI% > 0 Then
          If CI% <= NUMAS% Then
                  'Get #N2%, CI% + 1, SSQ$
            RESAL& = CI% + 1
            SSQ$ = REGLEIDO$("SALDEPO", RESAL&)
            CANUE = CANTI + (Int(10 * CVS(Mid$(SSQ$, 4 * SUCX% - 3, 4)) + 0.5)) / 10
            Mid$(SSQ$, 4 * SUCX% - 3, 4) = MKS$(CANUE)
                    'Put #N2%, CI% + 1, SSQ$
            Call GRAREG("SALDEPO", SSQ$, RESAL&)
          End If
        End If
      End If
119 Next i&
    GENINVE.Check2.Value = 1
    Call CIERRARCH("MOVISTO")
    Call CIERRARCH("SALDEPO")
    '
' ********************** INICIALIZANDO ARCHIVOS DE INVENTARIOS

    GENINVE.Check3.Value = 2
    '
    Call HEADERS("INVEDEQ", "")
    Call HEADERS("INVEDEQ", "Fecha de Inventario: " + FECHINVE$)
    Call HEADERS("INVEDEQ", TICOS$)
    If TRIM$(LOTAUT$) <> "" Then Call HEADERS("INVEDEQ", "Lote / Partida: " + LOTAUT$)
    '
    Call HEADERS("INVEGEN", "")
    Call HEADERS("INVEGEN", "Fecha de Inventario: " + FECHINVE$)
    Call HEADERS("INVEGEN", TICOS$)
    If TRIM$(LOTAUT$) <> "" Then Call HEADERS("INVEGEN", "Lote / Partida: " + LOTAUT$)
    '
    Call HEADERS("INVANAL", "")
    Call HEADERS("INVANAL", "Fecha de Inventario: " + FECHINVE$)
    Call HEADERS("INVANAL", TICOS$)
    If TRIM$(LOTAUT$) <> "" Then Call HEADERS("INVANAL", "Lote / Partida: " + LOTAUT$)

120 Call BLANARCH("INVEDEP")
130 Call BLANARCH("INVEGEN")
140 Call BLANARCH("INVANAL")
   '
    GENINVE.Check3.Value = 1
    '
' ******************** GENERANDO INVENTARIO ANALITICO Y GLOBAL
    '
    GENINVE.Check4.Value = 2
    '
    FIN& = NUINV%
    '
    J& = 0: TTI% = 0
    URE& = ULTREG&("SALDEPO")
    '
    For i& = 1 To FIN&
      Call TRACE(20, i&, FIN&)
      I1$ = REGLEIDO$("INVERT", i& + 1)
      If Mid$(I1$, 16, 1) = Chr$(96) Then GoTo 199          ' items de baja
      '
      CI% = CVI(Mid$(I1$, 17, 2))
      If CI% > 0 Then
        If CI% <= NUMAS% Then
          '
          REG& = CI% + 1
          SSQ$ = REGLEIDO$("SALDEPO", REG&)
          '
          SALTOTU = 0
          For U% = 1 To NUSU%
            SALTOTU = SALTOTU + Abs(CVS(Mid$(SSQ$, 4 * U% - 3, 4)))
          Next U%
          '
          If Abs(SALTOTU) >= 0.05 Then
            TOTAL = 0
            COU = COSUNI(CI%) * TICAMONE(MONECOSTO%(CI%))
            If VALORI% = 2 Then
                RFG$ = RECFILT$("RECOSCAL", 1, MKI$(CI%))
                If Len(RFG$) >= 4 Then
                  REGB& = CVS(RFG$)
                  If REGB& > 0 Then
                    Q$ = REGLEIDO$("RECOSCAL", REGB&)
                    COU = CVD(Mid$(Q$, 89, 8))
                  End If
                End If
              ElseIf VALORI% = 3 Then
                COU = 0
                For KKI% = 1 To CALIS%
                  COU = PRELISTA#(APRELI$(KKI%), CI%)
                  If COU >= 0.00005 Then
                    COU = COU * TICAMONE(MONECOSTO%(MONELI%(KKI%)))
                    Exit For
                  End If
                Next KKI%
              ElseIf VALORI% = 4 Then
                RFG$ = RECFILT$("RECOSCAL", 1, MKI$(CI%))
                If Len(RFG$) >= 4 Then
                  REGB& = CVS(RFG$)
                  If REGB& > 0 Then
                    Q$ = REGLEIDO$("RECOSCAL", REGB&)
                    COU = CVD(Mid$(Q$, 33, 8))
                  End If
                End If
                
            End If
            '
            TMT% = TIMATE%(CI%)
            '
            For U% = 1 To NUSU%
              SALDE = CVS(Mid$(SSQ$, 4 * U% - 3, 4))
              If Abs(SALDE) >= 0.05 Then
                '
                TOTAL = TOTAL + SALDE
                '
                Z$ = REGBLAN$("INVANAL")
                Call REPLA(Z$, MKI$(CI%), 1, 2)
                Call REPLA(Z$, Chr$(COSU%(U%)), 3, 1)
                Call REPLA(Z$, Chr$(TMT%), 4, 1)
                Call REPLA(Z$, MKS$(SALDE), 5, 4)
                Call REPLA(Z$, MKS$(0), 9, 4)
                Call REPLA(Z$, MKS$(0), 13, 4)
                Call REPLA(Z$, MKD$(0), 17, 8)
                Call REGAPP("INVANAL", Z$)
                REGA& = ULTREG&("INVANAL")
                '
                x$ = Left$(Z$, 8) + String$(12, 0)
                Call REPLA(x$, Chr$(TMT%), 4, 1)
                Call REPLA(x$, MKS$(COU), 9, 4)
                Call REPLA(x$, MKD$(CDbl(SALDE * COU)), 13, 8)
                Call REGAPP("INVEDEP", x$)
                '
              End If
            Next U%
            '
            If Abs(TOTAL) >= 0.05 Then
              Call REPLA(Z$, MKS$(TOTAL), 9, 4)
              Call REPLA(Z$, MKS$(COU), 13, 4)
              Call REPLA(Z$, MKD$(COU * TOTAL), 17, 8)
              Call GRAREG("INVANAL", Z$, REGA&)
              '
              Y$ = REGBLAN$("INVEGEN")
              Call REPLA(Y$, MKI$(CI%), 1, 2)
              Call REPLA(Y$, Chr$(0), 3, 1)
              Call REPLA(Y$, Chr$(TMT%), 4, 1)
              Call REPLA(Y$, MKS$(TOTAL), 5, 4)
              Call REPLA(Y$, MKS$(COU), 9, 4)
              Call REPLA(Y$, MKD$(CDbl(TOTAL * COU)), 13, 8)
              Call REGAPP("INVEGEN", Y$)
              '
            End If
          End If
        End If
      End If
199 Next i&
    '
    GENINVE.Check4.Value = 1
    '
999 Call CIERRARCH("*.*")
    Screen.MousePointer = 1
    GENINVE.Hide
    Unload GENINVE
    FORMINVEN.Text1.TEXT = FECHINVE$
    FORMINVEN.Text2.TEXT = TRIM$(CONTROL$("INVENTAR", "FEGENINV"))
    '
End Sub
'

Sub GENINV3()
    '
10  XX1$ = Left$(REGLEIDO$("ECOMAST", 1), 16): CI1% = Abs(CODINT%(XX1$))
    XX2$ = Left$(REGLEIDO$("ECOMAST", ULTREG&("ECOMAST")), 16): CI2% = Abs(CODINT%(XX2$))
    VDEF$ = MKI$(CI1%) + MKI$(CI2%) + Chr$(0) + MKI$(0)
    '
    OBX$ = OBJPLA$("DESHASCOD3", VDEF$)
    If Len(OBX$) < 5 Then
        Exit Sub
    End If
    COD1$ = CODEXT$(CVI(Left$(OBX$, 2)))
    COD2$ = CODEXT$(CVI(Mid$(OBX$, 3, 2)))
    TMT% = Asc(Mid$(OBX$, 5, 1))
    TIMAT$ = ECOARCH$("TATIMAT", Chr$(TMT%))
    If TIMAT$ = "" Then
        TMT% = 0
        Call REPLA(OBX$, Chr$(0), 5, 1)
        TIMAT$ = "Sin Clasificación por Tipo"
    End If
    '
    GCV% = CVI(Mid$(OBX$, 6, 2))
    GRUCOS$ = ECOARCH$("GRUCOVEN", MKI$(GCV%))
    If GRUCOS$ = "" Then
        GCV% = 0
        Call REPLA(OBX$, MKI$(0), 6, 2)
        GRUCOS$ = "Sin Clasificación Contable"
    End If
    '
    If TRIM$(COD1$) = "" Or TRIM$(COD2$) = "" Or COD2$ < COD1$ Then
        GoTo 10
    End If
    '
    Call COPYSTRU("INVEGEN", "INVEGENF")
    Call BLANARCH("INVEGENF")
    '
    Screen.MousePointer = 11
    FIN& = ULTREG&("INVEGEN")
    FIN1& = FIN&: If FIN1& < 1 Then FIN1& = 1
    '
    For i& = 1 To FIN&
        Call TRACE(20, i&, FIN&)
        x$ = REGLEIDO$("INVEGEN", i&)
        CI% = CVI(Left$(x$, 2))
        If CI% > 0 Then
           If CI% <= NUMAS% Then
              COD$ = CODEXT$(CI%)
              If COD$ >= COD1$ Then
                 If COD$ <= COD2$ Then
                   If TMT% = 0 Or TIMATE%(CI%) = TMT% Then
                     If GCV% = 0 Or GRUCOVE%(CI%) = GCV% Then
                       Call REGAPP("INVEGENF", x$)
                     End If
                   End If
                 End If
              End If
           End If
        End If
    Next i&
    '
    Call CIERRARCH("INVEGENF")
    Screen.MousePointer = 1
    '
    If FORMINVEN.Option6.Value = True Then
         Call FINAL("*LIINGEN1")
       ElseIf FORMINVEN.Option5.Value = True Then
         Call FINAL("*LIINGEN2")
    End If
    '
End Sub
'

Sub GENINV4()
    '
    XX1$ = Left$(REGLEIDO$("ECOMAST", 1), 16): CI1% = Abs(CODINT%(XX1$))
    XX2$ = Left$(REGLEIDO$("ECOMAST", ULTREG&("ECOMAST")), 16): CI2% = Abs(CODINT%(XX2$))
    VDEF$ = MKI$(CI1%) + MKI$(CI2%) + Chr$(0)
    '
10  OBX$ = OBJPLA$("DESHASCODEP", VDEF$)
    If Len(OBX$) < 5 Then
        Exit Sub
    End If
    COD1$ = CODEXT$(CVI(Left$(OBX$, 2)))
    COD2$ = CODEXT$(CVI(Mid$(OBX$, 3, 2)))
    If TRIM$(COD1$) = "" Or TRIM$(COD2$) = "" Or COD2$ < COD1$ Then
        GoTo 10
    End If
    SUCUI% = Asc(Mid$(OBX$, 5, 1))
    If ECOARCH("ECODEP", Chr$(SUCUI%)) = "" Then
        GoTo 10
    End If
    '
    Call COPYSTRU("INVEDEP", "INVEDEQ")
    Call BLANARCH("INVEDEQ")
    '
    '\\\OT-07-0004-RG-05-01-07///
    DEPNOM$ = ECOARCH$("ECODEP", Chr$(SUCUI%))
    Call HEADERS("INVEDEQ", "")
    Call HEADERS("INVEDEQ", "Depósito: " + TRIM$(Str$((SUCUI%))) + " - " + DEPNOM$)
    '\\\OT-07-0004-RG-FIN///
    '
    Screen.MousePointer = 11
    FIN& = ULTREG&("INVEDEP")
    FIN1& = FIN&: If FIN1& < 1 Then FIN1& = 1
    '
    For i& = 1 To FIN&
        Call TRACE(20, i&, FIN&)
        x$ = REGLEIDO$("INVEDEP", i&)
        If Asc(Mid$(x$, 3, 1)) = SUCUI% Then
           CI% = CVI(Left$(x$, 2))
           If CI% > 0 Then
              If CI% <= NUMAS% Then
                 COD$ = CODEXT$(CI%)
                 If COD$ >= COD1$ Then
                    If COD$ <= COD2$ Then
                       Call REGAPP("INVEDEQ", x$)
                    End If
                 End If
              End If
           End If
        End If
    Next i&
    '
    Call CIERRARCH("INVEDEQ")
    Screen.MousePointer = 1
    '
    If FORMINVEN.Option6.Value = True Then
         Call FINAL("*LIINDEP1")
       ElseIf FORMINVEN.Option5.Value = True Then
         Call FINAL("*LIINDEP2")
    End If
    '
End Sub
'

Sub GENINV5()
    '
    XX1$ = Left$(REGLEIDO$("ECOMAST", 1), 16): CI1% = Abs(CODINT%(XX1$))
    XX2$ = Left$(REGLEIDO$("ECOMAST", ULTREG&("ECOMAST")), 16): CI2% = Abs(CODINT%(XX2$))
    VDEF$ = MKI$(CI1%) + MKI$(CI2%) + Chr$(0)
    '
10  OBX$ = OBJPLA$("DESHASCOD1", VDEF$)
    If Len(OBX$) < 5 Then
        Exit Sub
    End If
    COD1$ = CODEXT$(CVI(Left$(OBX$, 2)))
    COD2$ = CODEXT$(CVI(Mid$(OBX$, 3, 2)))
    TMT% = Asc(Mid$(OBX$, 5, 1))
    TIMAT$ = ECOARCH$("TATIMAT", Chr$(TMT%))
    If TIMAT$ = "" Then
        TMT% = 0
        Call REPLA(OBX$, Chr$(0), 5, 1)
        TIMAT$ = "Sin Clasificación"
    End If
    '
    If TRIM$(COD1$) = "" Or TRIM$(COD2$) = "" Or COD2$ < COD1$ Then
        GoTo 10
    End If
    '
    Call COPYSTRU("INVANAL", "INVANALF")
    Call BLANARCH("INVANALF")
    '
    Screen.MousePointer = 11
    FIN& = ULTREG&("INVANAL")
    FIN1& = FIN&: If FIN1& < 1 Then FIN1& = 1
    '
    For i& = 1 To FIN&
        Call TRACE(20, i&, FIN&)
        x$ = REGLEIDO$("INVANAL", i&)
        CI% = CVI(Left$(x$, 2))
        If CI% > 0 Then
           If CI% <= NUMAS% Then
              COD$ = CODEXT$(CI%)
              If COD$ >= COD1$ Then
                 If COD$ <= COD2$ Then
                    If TMT% = 0 Or TIMATE%(CI%) = TMT% Then
                        Call REGAPP("INVANALF", x$)
                    End If
                 End If
              End If
           End If
        End If
    Next i&
    '
    Call CIERRARCH("INVANALF")
    Screen.MousePointer = 1
    '
    If FORMINVEN.Option6.Value = True Then
         Call FINAL("*LIINANA1")
       ElseIf FORMINVEN.Option5.Value = True Then
         Call FINAL("*LIINANA2")
    End If
    '
End Sub
