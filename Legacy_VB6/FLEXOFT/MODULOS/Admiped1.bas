Attribute VB_Name = "ADMIPED1"
Public FORMAPED%
Public FORCAPED%    ' 1.- MANUAL    2.- ITEMIZADO
Public VERLICRE%
Public YAINDI%
Public REVIPEPE%
Public ARCHIPRE$
'
Dim ENCAPED$
Dim CONPA$
Dim VENDE$
Dim NOCOM$
'

Sub PEDCLINUE()
    '
    Call PEDIDONUE(1)
    '
End Sub
'

Sub PEDINTNUE()
    '
    Call PEDIDONUE(2)
    '
End Sub
'

Sub PEDIDONUE(TIPED%)
     '
     HOII% = HOY(HO$)
     DIANTIDES% = Val(CONTROL$("", "ANTIDESP"))
     If DIANTIDES% < 1 Then DIANTIDES% = 3
     FORMAPED% = Val(CONTROL$("", "FORMAPED"))
     If FORMAPED% < 1 Or FORMAPED% > 2 Then FORMAPED% = 1
     If CONTROL$("ADMIPED", "VERLICRE") = "SI" Then VERLICRE% = 1
     '
1000 NROPE& = 1
     If ULTREG&("PEDENCA") > 0 Then
        X$ = REGLEIDO$("PEDENCA", ULTREG&("PEDENCA"))
        If Len(X$) >= 4 Then NROPE& = 1 + CVS(Left$(X$, 4))
     End If
     '
1100 If TIPED% = 2 Then
        PEDX$ = MKS$(NROPE&) + MKI$(HOII%) + MKI$(9999)
        Y$ = Space$(344)
        Call REPLA(Y$, Left$(PEDX$, 8), 1, 8)
        Call REPLA(Y$, "INTERNO", 9, 32)
        '
        Y$ = Y$ + MKI$(HOII%) + Chr$(0) + Chr$(0) + Chr$(0) + Chr$(0) + String$(20, 0)
        Call REPLA(Y$, MKI$(0), 345, 2)
        PEDY$ = Y$
        GoTo 1200
     End If
     '
     Screen.MousePointer = 1
     If PEDX$ = "" Then PEDX$ = MKS$(NROPE&) + MKI$(HOII%) + MKI$(NC%)
     PEDX$ = OBJPLA$("PEDIDO-NUE", PEDX$)
     If Len(PEDX$) < 24 Then
        Call CIERRARCH("*.*")
        Exit Sub
     End If
     '
     NROPE& = CVS(Left$(PEDX$, 4))
     NC% = CVI(Mid$(PEDX$, 7, 2))
     If NC% < 1 Then
        Beep
        GoTo 1000
     End If
     '
     If TRIM$(DOMICLI$(NC%)) = "" Then
        Call MENSERR(24, "Imposible Cargar Pedido a Esta Cuenta.\  \Falta Domicilio Legal.")
        GoTo 1000
     End If
     '
     Call NOTACLI(NC%, 1)
     '
     If VERLICRE% = 1 Then
       LMCR# = LIMICRE(NC%)
       SDEU# = SALDCLI#(NC%)
       CHPA# = CHEPEACRE#(NC%)
       If SDEU# + CHPA# > LMCR# Then
         Call COMUNI("Este Pedido no Podrá Facturarse.\  \Límite de Crédito Asignado ya está Superado.\  \Informe al Supervisor de Ventas.")
       End If
     End If
     '
     LENTREG$ = Space$(96)
     RFEX$ = RECFILT$("LENTRECL", 1, MKI$(NC%))
     If Len(RFEX$) >= 4 Then
       RELENT& = CVS(Left$(RFEX$, 4))
       If RELENT& > 0 Then
         If RELENT& <= ULTREG&("LENTRECL") Then
           LENTREG$ = Mid$(REGLEIDO$("LENTRECL", RELENT&), 3)
         End If
       End If
     End If
     '
     Y$ = Left$(PEDY$, 344): If Y$ = "" Then Y$ = Space$(344)
     Call REPLA(Y$, Left$(PEDX$, 8), 1, 8)
     Call REPLA(Y$, Mid$(PEDX$, 9, 25), 137, 16)
     Call REPLA(Y$, LENTREG$, 153, 96)
     Call REPLA(Y$, RASOCLI$(NC%), 9, 32)
     Call REPLA(Y$, DOMICLI$(NC%), 41, 32)
     CLOCA$ = TRIM$(CPOSCLI$(NC%)): If CLOCA$ <> "" Then CLOCA$ = CLOCA$ + " "
     Call REPLA(Y$, CLOCA$ + LOCACLI$(NC%), 73, 32)
     Call REPLA(Y$, TRIM$(CUITCLI$(NC%)) + " " + TRIM$(POSIVAC$(NC%)), 105, 32)
     '
     Y$ = Y$ + MKI$(HOII%) + Chr$(0) + Chr$(LIPRCLI%(NC%)) + Chr$(CPAGCLI%(NC%)) + Chr$(VENDCLI%(NC%)) + String$(20, 0)
     If App.EXEName = "PEDBASI" Then
        PEDY$ = Y$ + String$(4, 0)
        ARCHIPRE$ = "MASTER"
        GoTo 1200
     End If
     '
     PEDY$ = OBJPLA$("PEDIDO-ENCA", Y$)
     If PEDY$ = "" Then GoTo 1100
     Call REPLA(PEDY$, MKI$(0), 345, 2)
1110 PEDY$ = OBJPLA$("PEDIDO-ENCB", PEDY$)
     If PEDY$ = "" Then GoTo 1100
     '
     LENTREG$ = Mid$(PEDY$, 153, 96)
     If RELENT& > 0 Then
       If RELENT& <= ULTREG&("LENTRECL") Then
          Call GRAREG("LENTRECL", MKI$(NC%) + LENTREG$, RELENT&)
          GoTo 1120
       End If
     End If
     Call REGAPP("LENTRECL", MKI$(NC%) + LENTREG$)
1120 Call CIERRARCH("LENTRECL")
     '
     LPRE% = Asc(Mid$(PEDY$, 348, 1))
     If LPRE% < 1 Then
        Beep
        GoTo 1110
     End If
     '
     CPAG% = Asc(Mid$(PEDY$, 349, 1)): CONPA$ = DECONPA$(CPAG%)
     VEND% = Asc(Mid$(PEDY$, 350, 1)): VENDE$ = DEVENDE$(VEND%)
     '
     ARCHIPRE$ = TRIM$(Str$(LPRE%))
     While Len(ARCHIPRE$) < 3: ARCHIPRE$ = "0" + ARCHIPRE$: Wend
     ARCHIPRE$ = "LIPRE" + ARCHIPRE$
     '
1200 Screen.MousePointer = 1
     PRF$ = "CARDETP": If FORMAPED% = 2 Then PRF$ = "CARDETQ"
     EPAC$ = UCase$(EMPREAC$)
     If InStr(EPAC$, "NEUMANN") > 0 Then
       PRF$ = "CARPENEU"
     End If
     If TIPED% = 2 Then ARCHIPRE$ = "MASTER"
     XXX% = VENTABU%(PRF$ + "/NC/ARCHECO(1)=" + ARCHIPRE$)
     If XXX% < 0 Then
        GoTo 1900
     End If
     '
1300 If ULTREG&("!CARDETPE") < 1 Then
        GoTo 1200
     End If
     '
     Screen.MousePointer = 11   ' Registrando Pedido ...
     '
     EPAC$ = TRIM$(UCase$(EMPREAC$))
     FIN& = ULTREG&("!CARDETPE")
     For U& = 1 To FIN&
        X$ = REGLEIDO$("!CARDETPE", U&)
        CI% = CVI(Left$(X$, 2))
        If CI% < 1 Or CI% > NUMAS% Then
          Call MENSERR(24, "Codigo no Válido")
          GoTo 1200
        End If
        If TIPED% <> 2 Then
          If InStr(EPAC$, "SABO") > 0 Then
            If PRELISTA#(ARCHIPRE$, CI%) < 0.005 Then
              Call MENSERR(24, "Código '" + TRIM$(CODEXT$(CI%)) + "' no figura\en Lista de Precios " + ARCHIPRE$)
              GoTo 1200
            End If
          End If
        End If
        '
        For W& = 1 To U& - 1
           X$ = REGLEIDO$("!CARDETPE", W&)
           If CVI(Left$(X$, 2)) = CI% Then
              Screen.MousePointer = 1
              Call MENSERR(24, "Código " + TRIM$(CODEXT$(CI%)) + " Duplicado en Pedido.")
              GoTo 1200
           End If
        Next W&
     Next U&
     '
     HOII% = HOY(HOS$)
     ATX% = 0: CIANT% = 0: SUTOTPED# = 0
     For U& = 1 To ULTREG&("!CARDETPE")
       X$ = REGLEIDO$("!CARDETPE", U&)
       CI% = CVI(Left$(X$, 2))
       '
       If InStr(EPAC$, "SABO") > 0 Then
         PRELICARGA = PRELISTA#(ARCHIPRE$, CI%)
         Call REPLA(X$, MKS$(PRELICARGA), 39, 4)
         Call REPLA(X$, MKS$(PRELICARGA), 51, 4)
         Call REPLA(X$, MKS$(PRELICARGA), 65, 4)
         Call REPLA(X$, MKS$(0), 65, 4)
         Call GRAREG("!CARDETPE", X$, U&)
       End If
       '
       CAN = CVS(Mid$(X$, 47, 4))
       PRELICARGA = CVS(Mid$(X$, 39, 4))
       If PRELICARGA < 0.005 Then PRELICARGA = CVS(Mid$(X$, 51, 4))
       PRECARGADO = CVS(Mid$(X$, 65, 4))
       DESCARGADO = CVS(Mid$(X$, 69, 4))
       If PRECARGADO < 0.005 Then
         If PRELICARGA >= 0.005 Then
           PRECARGADO = PRELICARGA
         End If
       End If
       PRELICARGA = PRELISTA#(ARCHIPRE$, CI%)
       If PRECARGADO < 0.005 Then PRECARGADO = PRELICARGA
       PORDESCUIN = CVS(Mid$(X$, 69, 4))
       If PORDESCUEIN < 0 Or PORDESCUEIN > 99 Then PORDESCUEIN = 0
       PREUNITAR = PRECARGADO * (100 - PORDESCUIN) / 100
       TOTANT# = CVD(Mid$(X$, 77, 8))
       '
1310   TOT# = CDbl(CAN * PREUNITAR)
       SUTOTPED# = SUTOTPED# + TOT#
       If Abs(TOT# - TOTANT#) > 0.005 Then ATX% = 1
       '
       FENTRESOL% = CVI(Mid$(X$, 63, 2))
       If FENTRESOL% < HOII% Then
          FENTRESOL% = HOII%
          Call REPLA(X$, MKI$(FENTRESOL%), 63, 2)
          Call GRAREG("!CARDETPE", X$, U&)
          ATX% = 1
       End If
       '
       Call REPLA(X$, MKS$(PRELICARGA), 39, 4)
       Call REPLA(X$, MKS$(PRELICARGA), 51, 4)
       Call REPLA(X$, MKS$(PRECARGADO), 65, 4)
       Call REPLA(X$, MKS$(PORDESCUIN), 69, 4)
       Call REPLA(X$, MKS$(PREUNITAR), 73, 4)
       Call REPLA(X$, MKD$(TOT#), 77, 8)
       Call GRAREG("!CARDETPE", X$, U&)
       '
     Next U&
     '
     If ATX% > 0 Then
        Screen.MousePointer = 1
        GoTo 1200
     End If
     '
1320 VDEF$ = String$(41, 0)
     Call REPLA(VDEF$, MKD$(SUTOTPED#), 1, 8)
     Call REPLA(VDEF$, DESCUCLI$(NC%), 9, 16)
     Call REPLA(VDEF$, MKD$(SUTOTPED# * PORDESCLI(NC%) / 100), 25, 8)
     Call REPLA(VDEF$, MKD$(SUTOTPED# * (100 - PORDESCLI(NC%)) / 100), 33, 8)
     Call REPLA(VDEF$, Chr$(CPAG%), 41, 1)
     Screen.MousePointer = 1
     '
     EPAC$ = TRIM$(UCase$(EMPREAC$))
     If InStr(EPAC$, "NEUMAN") > 0 Then
        OBX$ = OBJPLA$("PEDIDO-CARCONPA", VDEF$)
        If OBX$ = "" Then GoTo 1200
        CPAG% = Asc(Mid$(OBX$, 41, 1))
        CONPA$ = TRIM$(DECONPA$(CPAG%))
        If CONPA$ = "" Then
           Call MENSERR(24, "Condicion de Venta no Válida")
           GoTo 1320
        End If
        CONPA$ = Left$(CONPA$, 1) + LCase$(Mid$(CONPA$, 2))
     End If
     '
     If VERLICRE% = 1 Then
       LMCR# = LIMICRE(NC%)
       SDEU# = SALDCLI#(NC%)
       CHPA# = CHEPEACRE#(NC%)
       TPED# = SUTOTPED# * (100 - PORDESCLI(NC%)) / 100
       If SDEU# + CHPA# + SUTOTPED# * (100 - PORDESCLI(NC%)) / 100 > LMCR# Then
         Call COMUNI("Este Pedido no Podrá Facturarse.\  \Supera el Límite de Crédito Asignado\para esta Cuenta de $ " + TRIM$(CSTRING$(MKD$(LMCR#), 3, 12, 2, 2)) + ".\  \Informe al Supervisor de Ventas.")
       End If
     End If
     '
     Screen.MousePointer = 11
     Call BLOQARCH("PEDENCA")
     Call BLOQARCH("PEDDETA")
     Call BLOQARCH("PEDDETP")
     Call BLOQARCH("ODESEMI")
     '
     NROPED& = CVS(Left$(PEDY$, 4))
     NROPRO& = 0
     Call NUPEDIDO(NROPRO&)
     EPAC$ = EMPREAC$
     If InStr(EPAC$, "NEUMANN") < 1 Then
        If NROPRO& <> NROPED& Then
           Call COMUNI("FLEXOFT ha Re-Numerado el Pedido\como Pedido Número '" + TRIM$(Str$(NROPRO&)) + "'.")
           NROPED& = NROPRO&
           Call REPLA(PEDY$, MKS$(NROPED&), 1, 4)
        End If
     End If
     '
     REFE$ = " - " + TRIM$(Mid$(PEDY$, 137, 16))
     If REFE$ = " - " Then REFE$ = ""
     FEX = CVI(Mid$(PEDY$, 5, 2))
     REFE$ = REFE$ + " - " + FECHATEX$(FEX)
     LOMA% = 48 - Len(REFE$)
     REFE$ = Left$(TRIM$(ECOARCH$("DEUDOR1", MKI$(NC%))), LOMA%) + REFE$
     '
     PDC$ = REGBLAN$("PEDENCA")
     Call REPLA(PDC$, Left$(PEDY$, 4), 1, 4)
     Call REPLA(PDC$, REFE$, 5, 48)
     Call REPLA(PDC$, Mid$(PEDY$, 5, 4), 53, 4)
     Call REPLA(PDC$, Mid$(PEDY$, 137, 214), 57, 214)
     Call REPLA(PDC$, MKI$(0), 265, 2)
     Call REPLA(PDC$, Chr$(1), 267, 1)
     Call REPLA(PDC$, Chr$(CPAG%), 269, 1)
     Call REPLA(PDC$, MKS$(1 + ULTREG&("PEDDETA")), 271, 4)
     Call REPLA(PDC$, String$(14, 0), 275, 14)
     Call REPLA(PDC$, CAVERI$, 289, 14)
     '
     Call REGAPP("PEDENCA", PDC$)
     REGENCA& = ULTREG&("PEDENCA")
     Call CIERRARCH("PEDENCA")
     '
     MODODESPA% = 0
     For U& = 1 To ULTREG&("!CARDETPE")
       X$ = REGLEIDO$("!CARDETPE", U&)
       CAN = CVS(Mid$(X$, 47, 4))
       If CAN > 0.05 Then
           '
           FENTRESOL% = CVI(Mid$(X$, 63, 2))
           If FENTRESOL% < HOII% Then
              FENTRESOL% = HOII%
              Call REPLA(X$, MKI$(FENTRESOL%), 63, 2)
              Call GRAREG("!CARDETPE", X$, U&)
           End If
           '
           PRECARGADO = CVS(Mid$(X$, 65, 4))
           PREUNITAR = CVS(Mid$(X$, 73, 4))
           TOTITEM# = CVD(Mid$(X$, 77, 8))
           PDC$ = REGBLAN$("PEDDETA")
           Call REPLA(PDC$, Left$(PEDY$, 8), 1, 8)
           Call REPLA(PDC$, Mid$(PEDY$, 137, 16), 9, 16)
           Call REPLA(PDC$, MKI$(FENTRESOL%), 25, 2)
           Call REPLA(PDC$, Chr$(1), 27, 1)
           Call REPLA(PDC$, Mid$(PEDY$, 348, 3), 28, 3)
           Call REPLA(PDC$, Chr$(CPAG%), 29, 1)
           Call REPLA(PDC$, MKI$(0), 31, 2)
           Call REPLA(PDC$, Left$(X$, 62), 33, 62)
           Call REPLA(PDC$, MKS$(PRECARGADO), 71, 4)
           Call REPLA(PDC$, MKS$(PREUNITAR), 83, 4)
           Call REPLA(PDC$, MKD$(TOTITEM#), 87, 8)
           Call REPLA(PDC$, String$(18, 0), 95, 18)
           Call REPLA(PDC$, MKS$(CAN), 99, 4)
           Call REPLA(PDC$, MKS$(REGENCA&), 109, 4)  ' puntero a ENCABEZADO
           Call REPLA(PDC$, CAVERI$, 113, 16)
           '
           Call REGAPP("PEDDETA", PDC$)
           '
           RGX& = ULTREG&("PEDDETA")
           Call REPLA(PDC$, MKS$(RGX&), 109, 4)  ' puntero a PEDDETA
           Call REGAPP("PEDDETP", PDC$)
           '
           If App.EXEName = "PEDBASI" Then GoTo 1390
           If InStr(EMPREAC$, "SABO") < 1 Then
              If InStr(EMPREAC$, "NEUMANN") < 1 Then
                 GoTo 1390
              End If
           End If
           '
           FE1 = HOII%: FE2 = FENTRESOL%
           If DIENTRE%(FE1, FE2) <= DIANTIDES% Then
               CI% = CVI(Mid$(PDC$, 33, 2))
               CASODES = CAN
               STDN = STOCKDISNEVEN(CI%)
               If CASODES > STDN Then
                   CASODES = STDN
               End If
               '
               NPED& = CVS(Left$(PDC$, 4))
               FEPE% = CVI(Mid$(PDC$, 5, 2))
               NC% = CVI(Mid$(PDC$, 7, 2))
               '
               Y$ = REGBLAN$("ODESEMI")
               Call REPLA(Y$, MKS$(NPED&), 1, 4)
               Call REPLA(Y$, MKI$(FEPE%), 5, 2)
               Call REPLA(Y$, MKI$(NC%), 7, 2)
               Call REPLA(Y$, MKI$(FENTRESOL%), 9, 2)
               Call REPLA(Y$, MKI$(CI%), 11, 2)
               Call REPLA(Y$, MKS$(CAN), 13, 4)
               Call REPLA(Y$, MKS$(CASODES), 17, 4)
               Call REPLA(Y$, MKS$(RGX&), 21, 4)
               '
               Call REGAPP("ODESEMI", Y$)
               '
               If MODODESPA% = 0 Then
                  If CASODES >= 0.05 Then
                     MODODESPA% = 2
                  End If
               End If
               '
               If MODODESPA% > 0 Then
                  If CASODES < CAN Then
                     MODODESPA% = 1
                  End If
               End If
               '
           End If
        End If
1390 Next U& '
     Call CIERRARCH("PEDDETA")
     Call CIERRARCH("PEDDETP")
     Call CIERRARCH("ODESEMI")
     '
     Call CIERRARCH("*.*")
     '
     If App.EXEName = "PEDBASI" Then GoTo 1900
     '
     '****************************************************
     '
     FORIPRE$ = CONTROL$("", "FORPECLI")
     If FORIPRE$ = "" Then GoTo 1500
     '
     Screen.MousePointer = 11
     HOII% = HOY(HOS$)
     FECHDOC$ = HOS$
     NUMEDOC$ = TRIM$(Str$(NROPE&))
     While Len(NUMEDOC$) < 6
        NUMEDOC$ = "0" + NUMEDOC$
     Wend
     TIPODOC$ = "Pedido de Cliente"
     If NC% < 1 Then
        TIPODOC$ = "Pedido Interno"
     End If
     '
     CODTABU1$(1) = "COD-MAT"
     CODTABU1$(2) = "DES-MAT"
     CODTABU1$(3) = "CANTIDAD"
     CODTABU1$(4) = "PRE-LIST"
     CODTABU1$(5) = "PORDESCU"
     CODTABU1$(6) = "PRE-NETO"
     CODTABU1$(7) = "NET-ITEM"
     CODTABU1$(8) = "FECH-REC"
     CACOLTAB1% = 8
     '
     DESTIDOC% = NC%
     CAITAB1% = 0
     CAITCRU% = 0
     '
     SUMATOTA# = 0
     For U& = 1 To ULTREG&("!CARDETPE")
        XX$ = REGLEIDO$("!CARDETPE", U&)
        CIPRE% = CVI(Left$(XX$, 2))
        CAPEX$ = Str$(CVS(Mid$(XX$, 47, 4)))
        '
        If CIPRE% > 0 Then
           If CIPRE% <= NUMAS% Then
              If Val(CAPEX$) > 0 Then
                 '
                 PREDELISTA = PRELISTA#(ARCHIPRE$, CIPRE%)
                 If PREDELISTA < 0.005 Then PREDELISTA = CVS(Mid$(XX$, 51, 4))
                 PRELI$ = Str$(PREDELISTA)
                 PREUX$ = Str$(CVS(Mid$(XX$, 73, 4)))
                 DESCX$ = ""
                 If Val(PRELI$) >= 0.005 Then DESCX$ = Str$(100 * (Val(PRELI$) - Val(PREUX$)) / Val(PRELI$))
                 ITOTX$ = Str$((Val(CAPEX$)) * (Val(PREUX$)))
                 SUMATOTA# = SUMATOTA# + CDbl(Val(ITOTX$))
                 FEX = CVI(Mid$(XX$, 63, 2))
                 FENTX$ = FECHATEX$(FEX)
                 '
                 CAITAB1% = CAITAB1% + 1
                 TETABU1$(1, CAITAB1%) = CODEXT$(CIPRE%)
                 TETABU1$(2, CAITAB1%) = DESCRIT0$(CIPRE%)
                 TETABU1$(3, CAITAB1%) = CAPEX$
                 TETABU1$(4, CAITAB1%) = PRELI$
                 TETABU1$(5, CAITAB1%) = DESCX$
                 TETABU1$(6, CAITAB1%) = PREUX$
                 TETABU1$(7, CAITAB1%) = ITOTX$
                 TETABU1$(8, CAITAB1%) = FENTX$
                 '
              End If
           End If
        End If
109  Next U&
     '
     IDESCUEN# = SUMATOTA# * PORDESCLI(NC%) / 100
     CODPARAM$(1) = "IMP-NETO"
     DOCPARAM$(1) = Str$(SUMATOTA#)
     CODPARAM$(2) = "VENDEDOR"
     DOCPARAM$(2) = VENDE$
     CODPARAM$(3) = "COND-PAG"
     DOCPARAM$(3) = CONPA$
     CODPARAM$(4) = "NRO-OCOM"
     DOCPARAM$(4) = Mid$(PEDY$, 137, 16)
     CODPARAM$(5) = "DOMI-TRA"
     DOCPARAM$(5) = Mid$(PEDY$, 153, 32)
     CODPARAM$(6) = "DESTINO"
     DOCPARAM$(6) = Mid$(PEDY$, 185, 32)
     CODPARAM$(7) = "TRANSPOR"
     DOCPARAM$(7) = Mid$(PEDY$, 217, 32)
     CODPARAM$(8) = "IMP-DESC"
     DOCPARAM$(8) = Str$(IDESCUEN#)
     CODPARAM$(9) = "IMPNEDES"
     DOCPARAM$(9) = Str$(SUMATOTA# - IDESCUEN#)
     PORCENIVA = ALIVA
     If NC% > 0 Then
        If PIVACLI(NC%) = 2 Then
           PORCENIVA = ALIVA + ALIVARNI
        End If
     End If
     CODPARAM$(10) = "IMP-IVA"
     DOCPARAM$(10) = Str$((SUMATOTA# - IDESCUEN#) * PORCENIVA / 100)
     CODPARAM$(11) = "IMP-TOTA"
     DOCPARAM$(11) = Str$((SUMATOTA# - IDESCUEN#) * (100 + PORCENIVA) / 100)
     CAPARDOC% = 11
     '
     Call GRAREG("!OBSERVOF", Mid$(PEDY$, 249, 32), 1)
     Call GRAREG("!OBSERVOF", Mid$(PEDY$, 281, 32), 2)
     Call GRAREG("!OBSERVOF", Mid$(PEDY$, 313, 32), 3)
     Call SETULTREG("!OBSERVOF", 3)
     Call CIERRARCH("!OBSERVOF")
     '
     FORIPRE$ = TRIM$(CONTROL$("", "FORPECLI"))
     If FORIPRE$ <> "" Then
        Call PRINTDOC("FORPECLI")   ' PEDIDO DE CLIENTE
     End If
     '
     '*****************************************************
     '
1500 If InStr(EMPREAC$, "SABO") < 1 Then
        If InStr(EMPREAC$, "NEUMANN") < 1 Then
           GoTo 1900
        End If
     End If
     ' EMITIR ORDEN DE DESPACHO
     If MODODESPA% > 0 Then
        TTXX$ = "Puede Ahora Emitir la Orden de Despacho\Correspondiente a Este Pedido"
        If MODODESPA% = 1 Then
            TTXX$ = TTXX$ + "\en Forma PARCIAL."
        End If
        Call BLANARCH("!OBSERVOF")
        If COMALTER%(TTXX$ + "\\Emitir O.Despacho\Continuar") = 1 Then
            '
20          Call GENODESPA(NPED&, NODES&)
            If NODES& > 0 Then
                Call PRIODESPA(NODES&)
            End If
            If NPED& < 0 Then
               NPED& = Abs(NPED&)
               GoTo 20
            End If
        End If
     End If
     '
1900 Call CIERRARCH("*.*")
     '
     PEDX$ = ""
     Call BLANARCH("!CARDETPE")
     '
     If TIPED% = 1 Then GoTo 1000
     '
End Sub
'

Sub CAMBIAPED()
     '
     If BLOQRESULT%("PEDIFAC") <> 1 Then
        TMENSA$ = "Asignación de Pedidos"
        If TRIM$(UCase$(App.EXEName)) = "ADMIPED" Then
           TMENSA$ = "Facturación"
        End If
        Call MENSERR(24, "Proceso Imposible en este Momento.\Proceso de " + TMENSA$ + " Abierto.\  \Resuelva y Vuelva a Intentarlo.")
        Exit Sub
     End If
     '
     FORMAPED% = Val(CONTROL$("", "FORMAPED"))
     If FORMAPED% < 1 Or FORMAPED% > 2 Then FORMAPED% = 1
     '
2100 If YAINDI% = 0 Then
        '
        Screen.MousePointer = 11
        JJ& = 0
        FIN& = ULTREG&("PEDENCA")
        For I& = 1 To FIN&
           X$ = REGLEIDO$("PEDENCA", I&)
           If Asc(Mid$(X$, 267, 1)) <= 1 Then
               Call VERIREFPED(X$, I&)
               JJ& = JJ& + 1
               Call GRAREG("INDIPEP", Left$(X$, 52) + MKS$(I&), JJ&)
           End If
        Next I&
        Call SETULTREG("INDIPEP", JJ&)
        Call CIERRARCH("INDIPEP")
        '
        Screen.MousePointer = 1
        YAINDI% = 1
        '
     End If
     '
2112 If ULTREG&("INDIPEP") < 1 Then
        Call COMUNI("No hay Pedidos Abiertos\que Puedan Modificarse")
        GoTo 9999
     End If
     '
2110 X$ = REGSEL$("INDIPEP/Pedidos Pendientes de Entrega")
     If Len(X$) <= 4 Then
        GoTo 9999
     End If
     '
     NROPE& = CVS(Left$(X$, 4))
     REGENCA& = CVS(Mid$(X$, 53, 4))
     '
     'VERIFICAR CONSISTENCIA
     '
     If REGENCA& < 1 Or REGENCA& > ULTREG&("PEDENCA") Then
        GoTo 9999
     End If
     '
     RGL% = Int(REGENCA&)
     '
     Call BLANARCH("!CAMBPED")
     Call BLOQARCH("PEDDETP")
     '
     Screen.MousePointer = 11
     For I& = 1 To ULTREG&("PEDDETP")
        X$ = REGLEIDO$("PEDDETP", I&)
        If CVS(Left$(X$, 4)) = NROPE& Then
            If VERIPEDPE%(X$) <> 1 Then GoTo 9999
            Y$ = MKS$(I&) + Mid$(X$, 5)
            Call REGAPP("!CAMBPED", Y$)
        End If
     Next I&
     '
2120 PRF$ = "CAMBPED"
     If FORMAPED% = 2 Then PRF$ = "CAMBPEQ"
     EPAC$ = UCase$(EMPREAC$)
     If InStr(EPAC$, "NEUMANN") > 0 Then
       PRF$ = "CAMPENEU"
     End If
     XXX% = VENTABU%(PRF$ + "/NC")
     '
     If XXX% >= 0 Then
        Call BLOQARCH("PEDDETP")
        Call BLOQARCH("PEDDETA")
        For I& = 1 To ULTREG&("!CAMBPED")
            X$ = REGLEIDO$("!CAMBPED", I&)
            CAN = CVS(Mid$(X$, 79, 4))
            ENTRE = CVS(Mid$(X$, 95, 4))
            PENDE = CAN - ENTRE: If PENDE < 0 Then PENDE = 0
            Call REPLA(X$, MKS$(PENDE), 99, 4)
            ITOTA# = CDbl(CAN * CVS(Mid$(X$, 83, 4)))
            '
            REG& = CVS(Left$(X$, 4))
            Y$ = MKS$(NROPE&) + Mid$(X$, 5)
            Call GRAREG("PEDDETP", Y$, REG&)
            '
            REG& = CVS(Mid$(Y$, 109, 4))
            Z$ = Mid$(REGLEIDO$("PEDDETA", REG&), 109, 4)
            Call REPLA(Y$, Z$, 109, 4)
            If REG& > 0 Then
                If REG& <= ULTREG&("PEDDETA") Then
                    Call GRAREG("PEDDETA", Y$, REG&)
                End If
            End If
            '
        Next I&
     End If
     '
     Call CIERRARCH("PEDENCA")
     Call CIERRARCH("PEDDETP")
     Call CIERRARCH("PEDDETA")
     Call CIERRARCH("!CAMBPED")
     '
     GoTo 2112
     '
9999 Call CIERRARCH("*.*")
     Screen.MousePointer = 1
     '
End Sub
'

Sub ANULAPED()
     '
     If BLOQRESULT%("PEDIFAC") <> 1 Then
        TMENSA$ = "Asignación de Pedidos"
        If TRIM$(UCase$(App.EXEName)) = "ADMIPED" Then
           TMENSA$ = "Facturación"
        End If
        Call MENSERR(24, "Proceso Imposible en este Momento.\Proceso de " + TMENSA$ + " Abierto.\  \Resuelva y Vuelva a Intentarlo.")
        Exit Sub
     End If
     '
     YAINDI% = 0
     '
2000 JJ& = 0
     Call FRESHECHO("INDIPEP")
     FIN& = ULTREG&("PEDENCA")
     Screen.MousePointer = 11
     For I& = 1 To FIN&
       Call TRACE(20, I&, FIN&)
       X$ = REGLEIDO$("PEDENCA", I&)
       If Asc(Mid$(X$, 267, 1)) <= 1 Then
          Call VERIREFPED(X$, I&)
          JJ& = JJ& + 1
          Call GRAREG("INDIPEP", Left$(X$, 52) + MKS$(I&), JJ&)
       End If
     Next I&
     Call SETULTREG("INDIPEP", JJ&)
     Call CIERRARCH("INDIPEP")
     Screen.MousePointer = 1
     '
     If ULTREG&("INDIPEP") < 1 Then
        Call COMUNI("No hay Pedidos Abiertos\que Puedan Anularse")
        Call CIERRARCH("*.*")
        Exit Sub
     End If
     '
2210 X$ = REGSEL$("INDIPEP/Pedidos Pendientes para Anulación")
     If Len(X$) <= 4 Then GoTo 9999
     '
     NROPE& = CVS(Left$(X$, 4))
     REGENCA& = CVS(Mid$(X$, 53, 4))
     If REGENCA& < 1 Then GoTo 9999
     If REGENCA& > ULTREG&("PEDENCA") Then GoTo 9999
     '
     RPENCA$ = REGLEIDO$("PEDENCA", REGENCA&)
     If CVS(Left$(RPENCA$, 4)) <> NROPE& Then GoTo 9999
     '
     RGL% = Int(REGENCA&)
     '
     Screen.MousePointer = 11
     Call BLOQARCH("PEDDETP")
     Call BLOQARCH("ORDESPA")
     Call BLOQARCH("ODESEMI")
     J& = 0
     For I& = 1 To ULTREG&("PEDDETP")
        X$ = REGLEIDO$("PEDDETP", I&)
        If CVS(Left$(X$, 4)) <> NROPE& Then
              J& = J& + 1
              Call GRAREG("PEDDETP", X$, J&)
            Else
              If VERIPEDPE%(X$) <> 1 Then GoTo 9999
        End If
     Next I&
     '
     Call SETULTREG&("PEDDETP", J&)
     Call CIERRARCH("PEDDETP")
     '
     X$ = REGLEIDO$("PEDENCA", REGENCA&)
     Call REPLA(X$, Chr$(9), 267, 1)
     Call GRAREG("PEDENCA", X$, REGENCA&)
     '
     REGDETA& = CVS(Mid$(X$, 271, 4))
     If REGDETA& > 0 Then
        If REGDETA& <= ULTREG&("PEDDETA") Then
            Y$ = REGLEIDO$("PEDDETA", REGDETA&)
            If CVS(Left$(Y$, 4)) = NROPE& Then GoTo 2220
        End If
     End If
     REGDETA& = 1
     '
2220 FIN& = ULTREG&("PEDDETA")
     For I& = REGDETA& To FIN&
        Call TRACE(20, I&, FIN&)
        Y$ = REGLEIDO$("PEDDETA", I&)
        If CVS(Left$(Y$, 4)) = NROPE& Then
            Call REPLA(Y$, Chr$(9), 27, 1)
            Call GRAREG("PEDDETA", Y$, I&)
        End If
     Next I&
     '
     J& = 0
     NODEA& = 0
     FIN& = ULTREG&("ORDESPA")
     For I& = 1 To FIN&
        Call TRACE(20, I&, FIN&)
        Y$ = REGLEIDO$("ORDESPA", I&)
        If CVS(Mid$(Y$, 7, 4)) <> NROPE& Then
             J& = J& + 1
             Call GRAREG("ORDESPA", Y$, J&)
           Else
             NODE& = CVS(Left$(Y$, 4))
             If NODE& <> NODEA& Then
                NODEA& = NODE&
                Call COMUNI("Se ha Anulado la Orden de Despacho Número" + Str$(NODE&))
             End If
        End If
     Next I&
     Call SETULTREG("ORDESPA", J&)
     Call CIERRARCH("ORDESPA")
     '
     J& = 0
     FIN& = ULTREG&("ODESEMI")
     For I& = 1 To FIN&
        Call TRACE(20, I&, FIN&)
        Y$ = REGLEIDO$("ODESEMI", I&)
        If CVS(Mid$(Y$, 1, 4)) <> NROPE& Then
             J& = J& + 1
             Call GRAREG("ODESEMI", Y$, J&)
        End If
     Next I&
     Call SETULTREG("ODESEMI", J&)
     Call CIERRARCH("ODESEMI")
     '
     Call CIERRARCH("*.*")
     Screen.MousePointer = 1
     GoTo 2000
     '
9999 Call CIERRARCH("*.*")
     Screen.MousePointer = 1
     '
End Sub
'

Sub DESANUPED()
     '
     YAINDI% = 0
     '
2300 Screen.MousePointer = 11
     JJ& = 0
     Call FRESHECHO("INDIPEP")
     FIN& = ULTREG&("PEDENCA")
     For I& = 1 To FIN&
       Call TRACE(20, I&, FIN&)
       X$ = REGLEIDO$("PEDENCA", I&)
       If Asc(Mid$(X$, 267, 1)) = 9 Then
          Call VERIREFPED(X$, I&)
          JJ& = JJ& + 1
          Call GRAREG("INDIPEP", Left$(X$, 52) + MKS$(I&), JJ&)
       End If
     Next I&
     Call SETULTREG("INDIPEP", JJ&)
     Call CIERRARCH("INDIPEP")
     Screen.MousePointer = 1
     '
     If ULTREG&("INDIPEP") < 1 Then
        Call COMUNI("No hay Pedidos Anulados\que Puedan Re-Abrirse")
        Call CIERRARCH("*.*")
        Exit Sub
     End If
     '
2310 X$ = REGSEL$("INDIPEP/Pedidos Anulados - Revertir Anulación")
     If Len(X$) <= 4 Then
        Call CIERRARCH("*.*")
        Exit Sub
     End If
     '
     NROPE& = CVS(Left$(X$, 4))
     REGENCA& = CVS(Mid$(X$, 53, 4))
     If REGENCA& < 1 Or REGENCA& > ULTREG&("PEDENCA") Then
        Call CIERRARCH("*.*")
        Exit Sub
     End If
     '
     RGL% = Int(REGENCA&)
'     Call FICHA("*COPEDNU/NC", RGL%)
'     If ALTERNA%("&16.60&-Confirma Re-Apertura") <> 1 Then
'        Call BLIMESS("-")
'        GoTo 2000
'    End If
     '
     Call BLOQARCH("PEDDETP")
     '
     Screen.MousePointer = 11
     X$ = REGLEIDO$("PEDENCA", REGENCA&)
     Call REPLA(X$, Chr$(1), 267, 1)
     Call GRAREG("PEDENCA", X$, REGENCA&)
     '
     J& = 0
     FIN& = ULTREG&("PEDDETA")
     For I& = 1 To FIN&
        Call TRACE(20, I&, FIN&)
        X$ = REGLEIDO$("PEDDETA", I&)
        PENDE = CVS(Mid$(X$, 79, 4)) - CVS(Mid$(X$, 95, 4))
        If PENDE < 0 Then PENDE = 0
        Call REPLA(X$, MKS$(PENDE), 99, 4)
        If CVS(Left$(X$, 4)) = NROPE& Then
        Call REPLA(X$, Chr$(1), 27, 1)
        Call GRAREG("PEDDETA", X$, I&)
        '
        Call REPLA(X$, MKS$(I&), 109, 4)       ' puntero a PEDDETA
        Call REGAPP("PEDDETP", X$)
        End If
     Next I&
     '
     Call CIERRARCH("PEDDETP")
     Call CIERRARCH("*.*")
     '
     Screen.MousePointer = 1
     '
     GoTo 2300
     '
End Sub
'

Function VERIPEDPE%(REGPE$)
    '
    Static CTX&, CTY&
    If CTX& = 0 Then
        CTX& = ULTREG&("PEDENCA")
        '
        If BLOQRESULT%("PEDIFAC") <> 1 Then
           TMENSA$ = "Asignación de Pedidos"
           If TRIM$(UCase$(App.EXEName)) = "ADMIPED" Then
              TMENSA$ = "Facturación"
           End If
           Call MENSERR(24, "Proceso Imposible en este Momento.\Proceso de " + TMENSA$ + " Abierto.\  \Resuelva y Vuelva a Intentarlo.")
           Call FINAL("")
        End If
        '
    End If
    '
    If CTY& = 0 Then
        CTY& = ULTREG&("PEDDETA")
    End If
    '
    VPP% = 0
    '
    X1$ = REGPE$
    NPED0& = CVS(Left$(X1$, 4))
    REGA& = CVS(Mid$(X1$, 109, 4))
    Call REPLA(X1$, String$(4, 0), 109, 4)
    Call REPLA(X1$, String$(2, 0), 31, 2)
10  If REGA& > 0 Then
       If REGA& <= CTY& Then
          X2$ = REGLEIDO$("PEDDETA", REGA&)
          REGENCA& = CVS(Mid$(X2$, 109, 4))
          Call REPLA(X2$, String$(4, 0), 109, 4)
          Call REPLA(X2$, String$(2, 0), 31, 2)
          If X1$ = X2$ Then
             NPEDI& = CVS(Left$(X2$, 4))
             If REGENCA& > 0 Then
                If REGENCA& <= CTX& Then
                   X3$ = REGLEIDO$("PEDENCA", REGENCA&)
                   NPEDJ& = CVS(Left$(X3$, 4))
                   If NPEDJ& = NPEDI& Then
                      VPP% = 1
                      GoTo 999
                   End If
                   MRX$ = "(1)"
                End If
                MRX$ = "(2)"
             End If
             MRX$ = "(3)"
          End If
          MRX$ = "(4)"
       End If
       MRX$ = "(5)"
    End If
    MRX$ = "(6)"
    '
    Screen.MousePointer = 1
    MENSA$ = "FLEXOFT Debe Ahora Reorganizar\"
    MENSA$ = MENSA$ + "la Base de Datos de Pedidos Pendientes.\"
    MENSA$ = MENSA$ + "Error " + MRX$ + " en Pedido " + TRIM$(Str$(NPED0&)) + ".\  \"
    MENSA$ = MENSA$ + "Pinche sobre 'Aceptar Mensaje'\"
    MENSA$ = MENSA$ + "para Iniciar el Proceso,\"
    MENSA$ = MENSA$ + "y Espere Hasta su Finalización."
    Call COMUNI(MENSA$)
    '
    Screen.MousePointer = 11
    '
    Dim NPI&(65535)
    If CTZ& = 0 Then
       CTZ& = ULTREG&("PEDENCA")
       For U& = 1 To CTZ&
           Call TRACE(20, U&, CTZ&)
           X3$ = REGLEIDO$("PEDENCA", U&)
           NPI&(U&) = CVS(Left$(X3$, 4))
       Next U&
    End If
    '
    Call BLOQARCH("PEDDETP")
    J& = 0
    FIN& = ULTREG&("PEDDETA")
    For U& = 1 To FIN&
       Call TRACE(20, U&, FIN&)
       X1$ = REGLEIDO$("PEDDETA", U&)
       NPEDI& = CVS(Left$(X1$, 4))
       If NPEDI& > 0 Then
          REGENCA& = CVS(Mid$(X1$, 109, 4))
          If REGENCA& > 0 Then
             If REGENCA& <= CTZ& Then
                X3$ = REGLEIDO$("PEDENCA", REGENCA&)
                If CVS(Left$(X3$, 4)) = NPEDI& Then
                   GoTo 100
                End If
             End If
          End If
          '
          For UI& = 1 To CTZ&
             If NPI&(UI&) = NPEDI& Then
                REGENCA& = UI&
                Call REPLA(X1$, MKS$(REGENCA&), 109, 4)
                Call GRAREG("PEDDETA", X1$, U&)
                GoTo 100
             End If
          Next UI&
          '
          If NPEDI& <> NPEDIA& Then
             Call COMUNI("No se Encontró el Encabezado\del Pedido Número" + Str$(NPEDI&) + ".\Este Pedido Queda Como 'Anulado'\   \Consulte a su Soporte de Sistemas FLEXOFT.")
             NPEDIA& = NPEDI&
          End If
          Call REPLA(X1$, Chr$(9), 27, 1)
          Call GRAREG("PEDDETA", X1$, U&)
       End If
       '
100    CAPED = CVS(Mid$(X1$, 79, 4))
       CAENT = CVS(Mid$(X1$, 95, 4))
       CAPEN = CAPED - CAENT: If CAPEN < 0 Then CAPEN = 0
       If CVS(Mid$(X1$, 99, 4)) <> CAPEN Then
          Call REPLA(X1$, MKS$(CAPEN), 99, 4)
          Call GRAREG("PEDDETA", X1$, U&)
       End If
       '
       If Asc(Mid$(X1$, 27, 1)) <= 1 Then
            If CAPEN >= 0.05 Then
                Call REPLA(X1$, MKS$(U&), 109, 4)
                J& = J& + 1
                Call GRAREG("PEDDETP", X1$, J&)
            End If
       End If
       '
    Next U&
    '
    Call SETULTREG("PEDDETP", J&)
    Call CIERRARCH("PEDDETP")
    Screen.MousePointer = 1
    '
999 VERIPEDPE% = VPP%
    '
End Function

Sub VERIREFPED(RPENCA$, REGENCA&)
    '
    REFE$ = TRIM$(Mid$(RPENCA$, 5, 48))
    If Left$(REFE$, 1) = "-" Then
        NCI% = CVI(Mid$(RPENCA$, 55, 2))
        RASO$ = TRIM$(RASOCLI$(NCI%))
        If InStr(REFE$, RASO$) < 1 Then
            REFE$ = RASO$ + " " + Mid$(RPENCA$, 5, 48)
            Call REPLA(RPENCA$, REFE$, 5, 48)
            Call GRAREG("PEDENCA", RPENCA$, REGENCA&)
        End If
    End If
    '
End Sub
'

Sub CALDISPOST()
     '
     HOI% = HOY(HOS$)
     CAIT% = NUMAS%
     Dim cix%(16384)
     '
     FLIM% = 32767
     '
     J1& = 0
     J2& = 0
     '
     For II% = 1 To CAIT%
       cix%(II%) = 0
       XX$ = MKI$(II%) + String$(30, 0)
       REG& = II%
       Call GRAREG("TODISPO", XX$, REG&)
     Next II%
     '
     FIN& = NUMAS%
     For I& = 1 To FIN&
        Call TRACE(20, I&, FIN&)
        ST$ = REGLEIDO$("SALSTOCK", I& + 1) + String$(40, 0)
        STOCK = 0: For KK% = 1 To 40 Step 4
                      STOCK = STOCK + CVS(Mid$(ST$, KK%, 4))
                   Next KK%
        If Abs(STOCK) >= 0.05 Then
            Z$ = REGBLAN$("DISPOST")
            CI% = I&
            Call REPLA(Z$, MKI$(CI%), 1, 2)
            Call REPLA(Z$, MKI$(11), 3, 2)
            Call REPLA(Z$, MKI$(HOI%), 5, 2)
            Call REPLA(Z$, "         STOCK AL " + HOS$, 7, 30)
            Call REPLA(Z$, MKS$(STOCK), 37, 4)
            J2& = J2& + 1
            Call GRAREG("DISPOST", Z$, J2&)
            '
            XX$ = REGLEIDO$("TODISPO", I&)
            STOCK = STOCK + CVS(Mid$(XX$, 5, 4))
            Call REPLA(XX$, MKS$(STOCK), 5, 4)
            Call REPLA(XX$, MKI$(1), 3, 2)
            Call GRAREG("TODISPO", XX$, I&)
         End If
           '
     Next I&
     '
     NMM% = NUMAS%
     FIN& = ULTREG&("PEDDETP")
     '
     For U& = 1 To FIN&
        Call TRACE(20, U&, FIN&)
        X$ = REGLEIDO$("PEDDETP", U&)
        If Asc(Mid$(X$, 27, 1)) <= 1 Then                ' status activo
            TOPEN = CVS(Mid$(X$, 79, 4)) - CVS(Mid$(X$, 95, 4))
            If TOPEN < 0 Then TOPEN = 0
            ASIG = 0: If Mid$(X$, 75, 4) <> Space$(4) Then ASIG = CVS(Mid$(X$, 75, 4))
            If ASIG > TOPEN Then ASIG = TOPEN
            NOASI = TOPEN - ASIG
            If TOPEN > 0.05 Then
                 NROPE& = CVS(Left$(X$, 4))
                 NROS$ = AJUSTD$(TRIM$(Str$(NROPE&)), 6)
                 '
                 CI% = CVI(Mid$(X$, 33, 2))
                 If CI% < 1 Or CI% > NMM% Then GoTo 3300
                 If cix%(CI%) > 0 Then GoTo 3235
                 '
                 cix%(CI%) = CI%
                 '
                 Z$ = REGBLAN$("DISPOEN")
                 Call REPLA(Z$, MKI$(CI%), 1, 2)
                 Call REPLA(Z$, MKI$(11), 3, 2)
                 Call REPLA(Z$, MKI$(HOI%), 5, 2)
                 Call REPLA(Z$, "         STOCK AL " + HOS$, 7, 30)
                 '
                 JJ& = CI%
                 ST$ = REGLEIDO$("SALSTOCK", JJ& + 1) + String$(40, 0)
                 STOCK = 0: For KK% = 1 To 40 Step 4
                               STOCK = STOCK + CVS(Mid$(ST$, KK%, 4))
                            Next KK%
                 '
                 Call REPLA(Z$, MKS$(STOCK), 37, 4)
                 If CVI(Mid$(X$, 25, 2)) <= FLIM% Then
                     J1& = J1& + 1
                     Call GRAREG("DISPOEN", Z$, J1&)
                 End If
                 '
3235             Z$ = REGBLAN$("DISPOEN")
                 Call REPLA(Z$, MKI$(CI%), 1, 2)
                 Call REPLA(Z$, MKI$(21), 3, 2)
                 NC% = CVI(Mid$(X$, 7, 2))
                 Call REPLA(Z$, Mid$(X$, 25, 2), 5, 2)
                 Call REPLA(Z$, NROS$ + "   " + ECOARCH$("DEUDOR1", MKI$(NC%)), 7, 30)
                 Call REPLA(Z$, MKS$(TOPEN), 53, 4)
                 If CVI(Mid$(X$, 25, 2)) <= FLIM% Then
                     J1& = J1& + 1
                     Call GRAREG("DISPOEN", Z$, J1&)
                 End If
                 '
                 Call REPLA(Z$, MKS$(TOPEN), 53, 4)
                 Call REPLA(Z$, MKS$(ASIG), 57, 4)
                 Call REPLA(Z$, MKS$(NOASI), 61, 4)
                 J2& = J2& + 1
                 Call GRAREG("DISPOST", Z$, J2&)
                 '
                 I& = CI%
                 XX$ = REGLEIDO$("TODISPO", I&)
                 TOPEN = TOPEN + CVS(Mid$(XX$, 21, 4))
                 ASIG = ASIG + CVS(Mid$(XX$, 25, 4))
                 NOASI = NOASI + CVS(Mid$(XX$, 29, 4))
                 Call REPLA(XX$, MKS$(TOPEN), 21, 4)
                 Call REPLA(XX$, MKS$(ASIG), 25, 4)
                 Call REPLA(XX$, MKS$(NOASI), 29, 4)
                 DINETA = CVS(Mid$(XX$, 5, 4)) - ASIG - NOASI
                 Call REPLA(XX$, MKS$(DINETA), 37, 4)
                 Call REPLA(XX$, MKI$(1), 3, 2)
                 Call GRAREG("TODISPO", XX$, I&)
            End If
        End If
3300 Next U&
     '
     Call CIERRARCH("DISPOEN")
     Call CIERRARCH("DISPOST")
     Call GRARGEN("DISPOEN", "")
     '
     FIN& = ULTREG&("TODISPO")
     J& = 0: K& = 0
     For U& = 1 To FIN&
         Call TRACE(20, U&, FIN&)
         XX$ = REGLEIDO$("TODISPO", U&)
         If CVI(Mid$(XX$, 3, 2)) > 0 Then
                J& = J& + 1
                Call GRAREG("TODISPO", XX$, J&)
                TTT = Abs(CVS(Mid$(XX$, 21, 4))) + Abs(CVS(Mid$(XX$, 25, 4))) + Abs(CVS(Mid$(XX$, 29, 4)))
                If TTT > 0.05 Then
                    K& = K& + 1
                    Call GRAREG("PEDISPO", XX$, K&)
                End If
         End If
     Next U&
     '
     Call SETULTREG("TODISPO", J&)
     Call SETULTREG("PEDISPO", K&)
     Call CIERRARCH("TODISPO")
     Call CIERRARCH("PEDISPO")
     Call FINAL("*DISPOSO")
     '
Exit Sub

300 'Call CLEARVEN
    MODO% = ALTERNA%("Consulta por Codigo\Listado General\Listado Material Pedido\Listado de Detalle\Disponibilidad de Entregas\-\Cancelar")
    On MODO% GoTo 310, 320, 325, 330, 340
    GoTo 3900
    '
310 Call FINAL("*CODISTP")
320 Call FINAL("*LITODIS")
325 Call FINAL("*LIPEDIS")
330 Call FINAL("*LIDEDIS")
340 Call FINAL("*DISPOLI")
        '
3900    Call FINAL("")
    
    '
End Sub
'

Sub CALASIAUT()          ' revision de situacion de stocks
    '                      para asignacion de pedidos
    '
    If BLOQRESULT%("PEDIFAC") <> 1 Then
       TMENSA$ = "Asignación de Pedidos"
       If TRIM$(UCase$(App.EXEName)) = "ADMIPED" Then
           TMENSA$ = "Facturación"
       End If
       Call MENSERR(24, "Proceso Imposible en este Momento.\Proceso de " + TMENSA$ + " Abierto.\  \Resuelva y Vuelva a Intentarlo.")
       Exit Sub
    End If
    '
    Screen.MousePointer = 11
    Dim PUNPED$(16384)
    Dim STOLE%(16384)
    Dim STTOT(16384)
    Dim STASI(16384)
    '
    REVIPEPE% = 0
    NPEDA& = NPED&
    NODEA& = NODE&
    '
    HOI% = HOY(HOS$)
    DIANTIDES% = Val(CONTROL$("", "ANTIDESP"))
    If DIANTIDES% < 1 Then DIANTIDES% = 3
    '
    If ULTREG&("ORDESPA") < 1 Then
        FIN& = ULTREG&("PEDDETP")
        For U& = 1 To FIN&
            Call TRACE(20, U&, FIN&)
            '
            X$ = REGLEIDO$("PEDDETP", U&)
            If Mid$(X$, 75, 4) = Space$(4) Then
                Call REPLA(X$, MKS$(0), 75, 4)
            End If
            '
            CASIG = CVS(Mid$(X$, 75, 4))
            If CASIG > 0 Then
                Y$ = REGBLAN$("ORDESPA")
                Call REPLA(Y$, Left$(X$, 6), 1, 6)
                Call REPLA(Y$, Left$(X$, 30), 7, 30)
                Call REPLA(Y$, Mid$(X$, 33, 48), 37, 48)
                Call REPLA(Y$, Mid$(X$, 109, 4), 111, 4)
                Call REGAPP("ORDESPA", Y$)
            End If
            '
        Next U&
        Call CIERRARCH("ORDESPA")
        '
    End If
    '
10  'NMS& = NUMAS%
    'Call BLIMESS("Revisando Existencias Actuales")
    'For I& = 1 To NMS&
    '    Call TRACE(20, I&, NMS&)
    '    ST$ = REGLEIDO$("SALSTOCK", I& + 1) + String$(40, 0)
    '    STOCK = 0: For KK% = 1 To 40 Step 4
    '                  STOCK = STOCK + CVS(Mid$(ST$, KK%, 4))
    '               Next KK%
    '    STTOT(I&) = STOCK
    'Next I&
    '
    FIN& = ULTREG&("PEDDETP")
    Call BLOQARCH("PEDDETP")
    '
    ' revisar y depurar pedidos pendientes
    '
    V1& = 0: YABLOPEP% = 0
    For U& = 1 To FIN&
        Call TRACE(20, U&, FIN&)
        X$ = REGLEIDO$("PEDDETP", U&)
        If VERIPEDPE%(X$) <> 1 Then
            Call CIERRARCH("*.*")
            For V% = 1 To NUMAS%
                PUNPED$(V%) = ""
            Next V%
            GoTo 10
        End If
        '
        CI% = CVI(Mid$(X$, 33, 2))
        NPED& = CVS(Left$(X$, 4))
        '
        REGPEDA& = CVS(Mid$(X$, 109, 4))
        Call REPLA(X$, MKS$(0), 75, 4)
        '
        Y$ = REGLEIDO$("PEDDETA", REGPEDA&)
        Call REPLA(Y$, MKS$(0), 75, 4)
        '
        REGENCA& = CVS(Mid$(Y$, 109, 4))
        Z$ = REGLEIDO$("PEDENCA", REGENCA&)
        STATU% = 1
        If Asc(Mid$(Z$, 267, 1)) > 1 Then
            Call REPLA(Y$, Chr$(9), 27, 1)
            STATU% = 0
        End If
        
        Call GRAREG("PEDDETA", Y$, REGPEDA&)
        If STATU% = 1 Then
           V1& = V1& + 1
           If V1& <> U& Then
              If YABLOPEP% = 0 Then
                 Call BLOQARCH("PEDDETP")
                 YABLOPEP% = 1
              End If
           End If
           Call GRAREG("PEDDETP", X$, V1&)
           If NPED& > 0 Then
              If CI% > 0 Then
                 If CI% <= NUMAS% Then
                    PUNPED$(CI%) = PUNPED$(CI%) + MKS$(V1&)
                 End If
              End If
           End If
        End If
        '
90  Next U&
    Call SETULTREG("PEDDETP", V1&)
    Call CIERRARCH("PEDDETP")
    '
    ' depurar segun ordenes de despacho en curso
    '
    UREPEP& = ULTREG&("PEDDETP")
    FIN& = ULTREG&("ORDESPA")
    For U& = 1 To FIN&
        Call TRACE(20, U&, FIN&)
        X$ = REGLEIDO$("ORDESPA", U&)
        '
        NODE& = CVS(Left$(X$, 4))
        CI% = CVI(Mid$(X$, 37, 2))
        NPED& = CVS(Mid$(X$, 7, 4))
        CAPEN = CVS(Mid$(X$, 79, 4)) - CVS(Mid$(X$, 83, 4))
        '
        If NPED& > 0 Then
          If CI% > 0 Then
            If CI% <= NUMAS% Then
              For II% = 1 To Len(PUNPED$(CI%)) Step 4
                REGPEDP& = CVS(Mid$(PUNPED$(CI%), II%, 4))
                If REGPEDP& > 0 Then
                  If REGPEDP& <= UREPEP& Then
                    Y$ = REGLEIDO$("PEDDETP", REGPEDP&)
                    If CVS(Left$(Y$, 4)) = NPED& Then
                      '
                      CAPENPE = CVS(Mid$(Y$, 79, 4)) - CVS(Mid$(Y$, 95, 4))
                      CASIG = CAPEN
                      If CASIG > CAPENPE Then
                        CASIG = CAPENPE
                        '
                        If NPED& <> NPEDA& Or NODE& <> NODEA& Then
                           MENSA$ = "En la Orden de Despacho Número " + TRIM$(Str$(NODE&))
                           MENSA$ = MENSA$ + "\Correspondiente al Pedido Número " + TRIM$(Str$(NPED&))
                           MENSA$ = MENSA$ + "\la Cantidad Asignada Supera la Pendiente de Entrega."
                           MENSA$ = MENSA$ + "\   \Revise o Anule la Orden de Despacho Número " + TRIM$(Str$(NODE&)) + "."
                           Call COMUNI(MENSA$)
                           NPEDA& = NPED&
                           NODEA& = NODE&
                        End If
                        '
                      End If
                      CAPEN = CAPEN - CASIG
                      TOASI = CVS(Mid$(Y$, 75, 4)) + CASIG
                      Call REPLA(Y$, MKS$(TOASI), 75, 4)
                      Call GRAREG("PEDDETP", Y$, REGPEDP&)
                      '
                      REGPEDA& = CVS(Mid$(Y$, 109, 4))
                      Y$ = REGLEIDO$("PEDDETA", REGPEDA&)
                      Call REPLA(Y$, MKS$(TOASI), 75, 4)
                      Call GRAREG("PEDDETA", Y$, REGPEDA&)
                      '
                      If STOLE%(CI%) = 0 Then
                         STOLE%(CI%) = 1
                         STTOT(CI%) = STODEPOS(CI%, DEPOEXPE%)
                      End If
                      STTOT(CI%) = STTOT(CI%) - CASIG
                      '
                      GoTo 99
                      '
                    End If
                  End If
                End If
              Next II%
              '
              If NPED& <> NPEDA& Or NODE& <> NODEA& Then
                 MENSA$ = "El Pedido Número " + TRIM$(Str$(NPED&))
                 MENSA$ = MENSA$ + "\Correspondiente a la Orden de Despacho Número " + TRIM$(Str$(NODE&))
                 MENSA$ = MENSA$ + "\no Figura Actualmente Como Pendiente de Entrega."
                 MENSA$ = MENSA$ + "\   \Revise o Anule la Orden de Despacho Número " + TRIM$(Str$(NODE&)) + "."
                 Call COMUNI(MENSA$)
                 NPEDA& = NPED&
                 NODEA& = NODE&
              End If
              '
            End If
          End If
        End If
99  Next U&
    '
    FIN& = ULTREG&("PEDDETP")
    JJ& = 0: KJ& = 0
    For U& = 1 To FIN&
        Call TRACE(20, U&, FIN&)
        X$ = REGLEIDO$("PEDDETP", U&)
        CAPENPE = CVS(Mid$(X$, 79, 4)) - CVS(Mid$(X$, 75, 4)) - CVS(Mid$(X$, 95, 4))
        If CAPENPE > 0.05 Then
          CI% = CVI(Mid$(X$, 33, 2))
          NPED& = CVS(Left$(X$, 4))
          FEPE% = CVI(Mid$(X$, 5, 2))
          NC% = CVI(Mid$(X$, 7, 2))
          FENTRESOL% = CVI(Mid$(X$, 25, 2))
          REGPEDA& = CVS(Mid$(X$, 109, 4))
          '
          FE1 = HOI%: FE2 = FENTRESOL%
          If DIENTRE%(FE1, FE2) <= DIANTIDES% Then
          'If FENTRESOL% <= FLIM% Then
            If CI% > 0 Then
              If CI% <= NUMAS% Then
                '
                If STOLE%(CI%) = 0 Then
                   STOLE%(CI%) = 1
                   STTOT(CI%) = STODEPOS(CI%, DEPOEXPE%)
                End If
                '
                If STTOT(CI%) >= 0.05 Then
                  Y$ = REGBLAN$("ODESEMI")
                  Call REPLA(Y$, MKS$(NPED&), 1, 4)
                  Call REPLA(Y$, MKI$(FEPE%), 5, 2)
                  Call REPLA(Y$, MKI$(NC%), 7, 2)
                  Call REPLA(Y$, MKI$(FENTRESOL%), 9, 2)
                  Call REPLA(Y$, MKI$(CI%), 11, 2)
                  Call REPLA(Y$, MKS$(CAPENPE), 13, 4)
                  '
                  CANASI = CAPENPE
                  If CANASI > STTOT(CI%) - STASI(CI%) Then
                      CANASI = STTOT(CI%) - STASI(CI%)
                      CODD$ = TRIM$(CODEXT$(CI%))
                      For UVX& = 1 To KJ&
                         ZZ$ = REGLEIDO$("CODISTRI", UVX&)
                         If TRIM$(Left$(ZZ$, 16)) = CODD$ Then GoTo 100
                      Next UVX&
                      ZZ$ = REGBLAN$("CODISTRI")
                      KJ& = KJ& + 1
                      Call REPLA(ZZ$, CODD$, 1, 16)
                      Call REPLA(ZZ$, DESCRIT$(CI%), 17, 48)
                      Call GRAREG("CODISTRI", ZZ$, KJ&)
                  End If
                  '
100               Call REPLA(Y$, MKS$(CANASI), 17, 4)
                  Call REPLA(Y$, MKS$(REGPEDA&), 21, 4)
                  '
                  JJ& = JJ& + 1
                  Call GRAREG("ODESEMI", Y$, JJ&)
                  '
                  STASI(CI%) = STASI(CI%) + CANASI
                  '
                End If
              End If
            End If
          End If
        End If
    Next U&
    '
    Call SETULTREG("ODESEMI", JJ&)
    Call CIERRARCH("ODESEMI")
    Call SETULTREG("CODISTRI", KJ&)
    Call CIERRARCH("CODISTRI")
    Call CIERRARCH("*.*")
    Screen.MousePointer = 1
    '
    TTXX$ = "Proceso de Revisión Completo.\  \Puede Ahora Emitir Listado\de Pedidos a Despachar."
    If COMALTER%(TTXX$ + "\\Emitir Listado\Cancelar") = 1 Then
        Call CONECRUN("*ODESEMI", "")
    End If
    '
End Sub
'

Sub MUEDISTRI(CI%)
    '
    If CI% < 1 Or CI% > NUMAS% Then
        Exit Sub
    End If
    '
    Screen.MousePointer = 11
    '
    Static NUSU%, COSU%(10), SUCU$(10)
    If NUSU% = 0 Then
        For U& = 1 To ULTREG&("SUCURSAL")
           If U& <= 10 Then
              SS$ = REGLEIDO$("SUCURSAL", U&)
              If Asc(Left$(SS$, 1)) > 32 Then
                 NUSU% = U&
                 COSU%(U&) = Asc(Mid$(SS$, 64, 1))
                 SUCU$(U&) = Left$(SS$, 20)
              End If
           End If
        Next U&
     End If
     '
     STOTOT = 0
     '
     REG& = CI% + 1
     If REG& > 0 Then
        If REG& <= NUMAS% Then
           SALCA$ = REGLEIDO$("SALSTOCK", REG&)
           For U& = 1 To NUSU%
               XXZ$ = Mid$(SALCA$, 4 * U& - 3, 4)
               XXX$ = SUCU$(U&) + CSTRING$(XXZ$, 4, 12, 1, 2)
               STOTOT = STOTOT + CVS(XXZ$)
           Next U&
           DISTRIMAT.Label4.Caption = TRIM$(CSTRING$(MKS$(STOTOT), 4, 12, 1, 2))
       End If
     End If
     '
10   DISTRIMAT.List2.Clear
     HOII% = HOY(HO$)
     SALSTO = STOTOT
     '
     RGXX$ = RECFILT$("PEDDETP", 12, MKI$(CI%))
     PEDI = 0: ASII = 0: TOASI = 0
     For I& = 1 To Len(RGXX$) Step 4
        REG& = CVS(Mid$(RGXX$, I&, 4))
        X$ = REGLEIDO$("PEDDETP", REG&)
        If Asc(Mid$(X$, 27, 1)) <= 1 Then
            CIP% = CVI(Mid$(X$, 33, 2))
            If CIP% <> CI% Then        ' fuerza reindexacion
                X$ = REGLEIDO$("PEDDETP", 1)
                Call GRAREG("PEDDETP", X$, 1)
                Call CIERRARCH("PEDDETP")
                GoTo 10
            End If
            '
            NROPED& = CVS(Left$(X$, 4))
            NCI% = CVI(Mid$(X$, 7, 2))
            CASIG = 0
            If Mid$(X$, 75, 4) <> Space$(4) Then
                CASIG = CVS(Mid$(X$, 75, 4))
            End If
            CAPEN = CVS(Mid$(X$, 79, 4)) - CVS(Mid$(X$, 95, 4))
            If CAPEN < 0 Then CAPEN = 0
            '
            If ASII + CASIG > STOTOT Then
                CASIG = STOTOT - ASII
            End If
            '
            If CAPEN > 0 Or CASIG > 0 Then
                PEDI = PEDI + CAPEN
                ASII = ASII + CASIG
                If ASII > PEDI Then ASII = PEDI
            End If
            '
            SALSTO = SALSTO - CAPEN
            '
        End If
     Next I&
     If ASII > STOTOT Then
        ASII = STOTOT
     End If
     '
     DISTRIMAT.Label5.Caption = CSTRING$(MKS$(PEDI), 4, 12, 2, 2)
     DISTRIMAT.Label7.Caption = CSTRING$(MKS$(ASII), 4, 12, 2, 2)
     DISTRIMAT.Label14.Caption = CSTRING$(MKS$(STOTOT - PEDI), 4, 12, 2, 2)
     DISTRIMAT.Label10.Caption = CSTRING$(MKS$(STOTOT - ASII), 4, 12, 2, 2)
     '
     Call BLANARCH("!DISTRIMA")
     For I& = 1 To ULTREG&("ODESEMI")
        X$ = REGLEIDO$("ODESEMI", I&)
        If CVI(Mid$(X$, 11, 2)) = CI% Then
            '
            NROPED& = CVS(Left$(X$, 4))
            NC% = CVI(Mid$(X$, 7, 2))
            CANPE = CVS(Mid$(X$, 13, 4))
            CANAS = CVS(Mid$(X$, 17, 4))
            '
            Z$ = REGBLAN$("DISTRIMA")
            Call REPLA(Z$, MKS$(NROPED&), 1, 4)
            Call REPLA(Z$, RASOCLI$(NC%), 5, 30)
            Call REPLA(Z$, MKS$(CANPE), 35, 4)
            Call REPLA(Z$, MKS$(CANAS), 39, 4)
            Call REPLA(Z$, MKS$(I&), 43, 4)
            Call REGAPP("!DISTRIMA", Z$)
            '
        End If
     Next I&
     '
     Screen.MousePointer = 1
     '
     ATRI$ = "/NC"
     ATRI$ = ATRI$ + "/TITUPAN=Asignacion de Pedidos - Redistribucion"
     ATRI$ = ATRI$ + "/BORDESUP=W" + TRIM$(Str$(DISTRIMAT.Top + DISTRIMAT.Label15.Top))
     ATRI$ = ATRI$ + "/BORDEIZQ=W" + TRIM$(Str$(DISTRIMAT.Left + DISTRIMAT.Label15.Left))
     '
20   VTB% = VENTABU%("DISTRIMA" + ATRI$)
     If VTB% >= 0 Then
        SUMASI = 0
        For I& = 1 To ULTREG&("!DISTRIMA")
            X$ = REGLEIDO$("!DISTRIMA", I&)
            If CVS(Mid$(X$, 39, 4)) > CVS(Mid$(X$, 35, 4)) Then
                Call MENSERR(24, "La Cantidad Asignada\Supera la Pedida.")
                GoTo 20
            End If
            SUMASI = SUMASI + CVS(Mid$(X$, 39, 4))
        Next I&
        '
        If SUMASI > Val(DISTRIMAT.Label10.Caption) Then
            Call MENSERR(24, "Las Cantidades Distribuidas\Superan el Stock Disponible")
            GoTo 20
        End If
        '
        For I& = 1 To ULTREG&("!DISTRIMA")
            X$ = REGLEIDO$("!DISTRIMA", I&)
            CANAS = CVS(Mid$(X$, 39, 4))
            REGODES& = CVS(Mid$(X$, 43, 4))
            Z$ = REGLEIDO$("ODESEMI", REGODES&)
            Call REPLA(Z$, MKS$(CANAS), 17, 4)
            Call GRAREG("ODESEMI", Z$, REGODES&)
        Next I&
        '
        Call CIERRARCH("ODESEMI")
        '
     End If
     '
     Call CIERRARCH("!DISTRIMA")
     Call BLANARCH("!DISTRIMA")
     '
End Sub
'

Sub EMIODES()
    '
    Screen.MousePointer = 11
    FODESEMI.LIPEDES.Clear
    '
    NPEDA& = 0
    For U& = 1 To ULTREG&("ODESEMI")
        X$ = REGLEIDO$("ODESEMI", U&)
        NPED& = CVS(Left$(X$, 4))
        If NPED& > 0 Then
           If NPED& <> NPEDA& Then
              If CVS(Mid$(X$, 17, 4)) >= 0.05 Then
                 NPEDA& = NPED&
                 NC% = CVI(Mid$(X$, 7, 2))
                 Z$ = CSTRING$(MKS$(NPED&), 3, 7, 0, 2) + "  " + RASOCLI$(NC%)
                 FODESEMI.LIPEDES.AddItem Z$
              End If
           End If
        End If
    Next U&
    '
    Call CIERRARCH("*.*")
    Screen.MousePointer = 1
    If FODESEMI.LIPEDES.ListCount < 1 Then
        Call COMUNI("No Hay Pedidos Pendientes\que Puedan Asignarse.")
        Exit Sub
    End If
    '
    FODESEMI.Show 1
    Call FRESHALL
    '
    Screen.MousePointer = 11
    For U& = 1 To FODESEMI.LIPEDES.ListCount
        If FODESEMI.LIPEDES.Selected(U& - 1) = True Then
            FODESEMI.LIPEDES.ListIndex = U& - 1
            CAPESEL% = CAPESEL% + 1
            TTXX$ = FODESEMI.LIPEDES.TEXT
            '
            NPED& = Val(Left$(TTXX$, 7))
20          Call GENODESPA(NPED&, NODES&)
            If NODES& > 0 Then
                Call PRIODESPA(NODES&)
            End If
            If NPED& < 0 Then
               NPED& = Abs(NPED&)
               GoTo 20
            End If
        End If
    Next U&
    '
    Call CIERRARCH("*.*")
    '
    Screen.MousePointer = 1
    '
End Sub
'

Sub GENODESPA(NPED&, NODES&)
    '
    Dim REODES&(1024)
    KKK% = 0
    '
    Call BLOQARCH("ORDESPA")
    '
    FORIPRE$ = TRIM$(CONTROL$("", "FORDESPA"))
    If FORIPRE$ = "" Then
        FORIPRE$ = "ODES-SAB"
    End If
    CODFOR$ = TRIM$(UCase$(FORIPRE$))
    MAXIT% = Val(ATRIFORM$(CODFOR$, "MAX-ITEM"))
    If MAXIT% < 1 Or MAXIT% > 256 Then MAXIT% = 10
    '
    NODES& = 0
    For U& = 1 To ULTREG&("ODESEMI")
        X$ = REGLEIDO$("ODESEMI", U&)
        If CVS(Left$(X$, 4)) = NPED& Then
            CI% = CVI(Mid$(X$, 11, 2))
            Call VESARRAS(CI%)
            CANASI = CVS(Mid$(X$, 17, 4))
            STAXX = STODEPOS(CI%, DEPOEXPE%)
            If STAXX < 0 Then STAXX = 0
            If CANASI > STAXX Then CANASI = STAXX
            If CANASI >= 0.05 Then
               REGPEDA& = CVS(Mid$(X$, 21, 4))
               If REGPEDA& > 0 Then
                  If REGPEDA& <= ULTREG&("PEDDETA") Then
                     Y$ = REGLEIDO$("PEDDETA", REGPEDA&)
                     If CVS(Left$(Y$, 4)) = NPED& Then
                        If CVI(Mid$(Y$, 33, 2)) = CI% Then
                           KKK% = KKK% + 1
                           REODES&(KKK%) = U&
                           GoTo 10
                        End If
                     End If
                  End If
               End If
               '
               Call MENSERR(24, "Imposible Emitir Orden de Despacho\Correspondiente a Pedido " + TRIM$(Str$(NPED&)) + ".")
               GoTo 99
               '
            End If
        End If
10  Next U&
    '
    FIN& = ULTREG&("ORDESPA")
    NODES1& = 1
    If FIN& > 0 Then
        Z$ = REGLEIDO$("ORDESPA", FIN&)
        NODES1& = 1 + CVS(Left$(Z$, 4))
    End If
    FF% = HOY(HOS$)
    '
    For KK1% = 1 To KKK%
        If KK1% > MAXIT% Then
            NPED& = (-1) * Abs(NPED&)
            GoTo 99
        End If
        Y$ = REGLEIDO$("ODESEMI", REODES&(KK1%))
        CI% = CVI(Mid$(Y$, 11, 2))
        CANASI = CVS(Mid$(Y$, 17, 4))
        CANASIANT = CANASI
        REGPEDA& = CVS(Mid$(Y$, 21, 4))
        
        X$ = REGLEIDO$("PEDDETA", REGPEDA&)
        If Mid$(X$, 75, 4) = Space$(4) Then
            Call REPLA(X$, MKS$(0), 75, 4)
        End If
        '
15      VDEF$ = Y$
        Call REPLA(VDEF$, MKS$(CANASI), 29, 4)
        EPAC$ = UCase$(EMPREAC$)
        If InStr(EPAC$, "SABO") > 0 Then
          OBX$ = VDEF$
          GoTo 18
        End If
        '
        OBX$ = OBJPLA$("ASIGPEDI", VDEF$)
        If OBX$ = "" Then GoTo 98
        If CVS(Mid$(OBX$, 29, 4)) > CANASI + 0.05 Then
          If COMALTER%("ATENCION: Está Asignando Más que el Disponible !\\Corregir\Continuar") <> 2 Then GoTo 15
        End If
        '
18      CANASI = CVS(Mid$(OBX$, 29, 4))
        SALNOASI = CANASIANT - CANASI
        If SALNOASI < 0 Then SALNOASI = 0
        '
        CASIG = CVS(Mid$(X$, 75, 4)) + CANASI
        Call REPLA(X$, MKS$(CASIG), 75, 4)
        Call GRAREG("PEDDETA", X$, REGPEDA&)
        '
        XX$ = Left$(X$, 74)
        For W& = 1 To ULTREG&("PEDDETP")
            Z$ = REGLEIDO$("PEDDETP", W&)
            If Left$(Z$, 30) = Left$(XX$, 30) Then
               If Mid$(Z$, 33, 42) = Mid$(XX$, 33, 42) Then
                  Z$ = X$
                  Call REPLA(Z$, MKS$(REGPEDA&), 109, 4)
                  Call GRAREG("PEDDETP", Z$, W&)
                  GoTo 20
               End If
            End If
        Next W&
        '
        MENSA$ = "Se ha Detectado un Error de Consistencia\"
        MENSA$ = MENSA$ + "en la Base de Datos de Pedidos Pendientes.\"
        MENSA$ = MENSA$ + "Correspondiente al Pedido Número " + TRIM$(Str$(Abs(NPED&))) + ".\  \"
        MENSA$ = MENSA$ + "FLEXOFT Reparará este Error oportunamente."
        Call MENSERR(24, MENSA$)
        '
20      Call REPLA(Y$, MKS$(SALNOASI), 17, 4)
        Call GRAREG("ODESEMI", Y$, REODES&(KK1%))
        '
        Z$ = REGBLAN$("ORDESPA")
        Call REPLA(Z$, MKS$(NODES1&), 1, 4)
        Call REPLA(Z$, MKI$(FF%), 5, 2)
        Call REPLA(Z$, Left$(X$, 30), 7, 30)
        Call REPLA(Z$, Mid$(X$, 33, 42), 37, 42)
        Call REPLA(Z$, MKS$(CANASI), 79, 4)
        Call REPLA(Z$, MKS$(REGPEDA&), 111, 4)
        Call REGAPP("ORDESPA", Z$)
        '
        NODES& = NODES1&
        '
98  Next KK1%
    '
99  Call CIERRARCH("ORDESPA")
    Call CIERRARCH("*.*")
    '
End Sub
'

Sub PRIODESPA(NODES&)
    '
    If NODES& < 1 Then
        Call MENSERR(24, "Error de Invocacion de Rutina\de Impresion de Orden de Despacho")
        Exit Sub
    End If
    '
    FORIPRE$ = TRIM$(CONTROL$("", "FORDESPA"))
    If FORIPRE$ = "" Then
        FORIPRE$ = "ODES-SAB"
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
    AVAHOJA$ = TRIM$(CONTROL$("FORDESPA", "AVAHOJA"))
    Screen.MousePointer = 11
    '
30  CODFOR$ = TRIM$(UCase$(FORIPRE$))
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
    Screen.MousePointer = 11
    'Printer.Orientation = 1
    Call SETSPOOL("ORIENTATION", "1")
    If ORIENTA% = 2 Then
         'Printer.Orientation = 2
         Call SETSPOOL("ORIENTATION", "2")
    End If
    COP% = 0
    '
5   YAENCA% = 0
    NUHOJ% = 0
    For U& = ULTREG&("ORDESPA") To 1 Step -1
        X$ = REGLEIDO$("ORDESPA", U&)
        If CVS(Left$(X$, 4)) < NODES& Then GoTo 10
    Next U&
    U& = 0
    '
10  FINUBI& = ULTREG&("UBISTOCK")
    FEVEN% = 32767
    For V& = U& + 1 To ULTREG&("ORDESPA")
        X$ = REGLEIDO$("ORDESPA", V&)
        If CVS(Left$(X$, 4)) = NODES& Then
            If YAENCA% = 0 Then
                NPED& = CVS(Mid$(X$, 7, 4))
                NC% = CVI(Mid$(X$, 13, 2))
                GoSub 257
                YAENCA% = 1
            End If
            '
            CI% = CVI(Mid$(X$, 37, 2))
            ASIG = CVS(Mid$(X$, 79, 4))
            If ASIG > 0 Then
               '
               YAC = YAC + INTERLI
               If YAC > HASMM Then
                   If FEVEN% < HOY(HOS$) Then FEVEN% = HOY(HOS$)
                   FEX = FEVEN%
                   Call PRINTFORWIN(FECHATEX$(FEX), "FECH-VEN", 0, 0)
                   'Printer.EndDoc
                   Call SETSPOOL("ENDDOC", "")
                   GoSub 257
               End If
               '
               FENTRESOL% = CVI(Mid$(X$, 31, 2))
               If FENTRESOL% > 0 Then
                  If FENTRESOL% < FEVEN% Then
                     FEVEN% = FENTRESOL%
                  End If
               End If
               FEX = FENTRESOL%
               '
               Call PRINTFORWIN(CODEXT$(CI%), "COD-MAT", 0, YAC)
               Call PRINTFORWIN(DESCRIT$(CI%), "DES-MAT", 0, YAC)
               Call PRINTFORWIN(TRIM$(Str$(ASIG)), "CANTIDAD", 0, YAC)
               Call PRINTFORWIN(FECHATEX$(FEX), "FECH-LAN", 0, YAC)
               Call PRINTFORWIN("..........", "REF-MAT1", 0, YAC)
               '
20             KKK% = 0
               RGXX$ = RECFILT$("UBISTOCK", 1, MKI$(CI%))
               For W& = 1 To Len(RGXX$) Step 4
                   REG& = CVS(Mid$(RGXX$, W&, 4))
                   UB$ = REGLEIDO$("UBISTOCK", REG&)
                   UB1$ = Left$(UB$, 2)
                   If CVI(UB1$) <> CI% Then
                      UB$ = REGLEIDO$("UBISTOCK", 1)
                      Call GRAREG("UBISTOCK", UB$, 1)
                      Call CIERRARCH("UBISTOCK")
                      GoTo 20
                   End If
                   UB2$ = Mid$(UB$, 3, 24)
                   If TRIM$(UB2$) <> "" Then
                      If KKK% > 0 Then
                         'Printer.Print Space$(76);
                         Call SETSPOOL("PRINT", Space$(76))
                      End If
                      KKK% = KKK% + 1
                      If KKK% > 1 Then
                          YAC = YAC + INTERLI
                      End If
                      Call PRINTFORWIN(Left$(UB2$, 16), "REF-MAT2", 0, YAC)
                   End If
               Next W&
               '
            End If
        End If
    Next V&
    '
    If FEVEN% < HOY(HOS$) Then FEVEN% = HOY(HOS$)
    FEX = FEVEN%
    Call PRINTFORWIN(FECHATEX$(FEX), "FECH-VEN", 0, 0)
    'Printer.EndDoc
    Call SETSPOOL("ENDDOC", "")
    '
    COP% = COP% + 1
    If COP% < CACOPI% Then GoTo 5
    '
    DESCRIDOC$ = "Orden de Despacho Nro. " + TRIM$(Str$(NODES&))
    DESCRIDOC$ = DESCRIDOC$ + " - " + TRIM$(RASOCLI$(NC%))
    FEX = FF%
    DESCRIDOC$ = DESCRIDOC$ + " - " + FECHATEX$(FEX)
    Call SAVESPOOL(DESCRIDOC$, OKX%)
    '
    Screen.MousePointer = 1
    Printer.Orientation = PORIAN
    '
Exit Sub
    '
    '
257 Call CARFORWIN
    '
    NUHOJ% = NUHOJ% + 1
    Call PRINTFORWIN(TRIM$(RASOCLI$(NC%)) + " (" + TRIM$(Str$(NC%)) + ")", "RASO-CLI", 0, 0)
    Call PRINTFORWIN(DOMICLI$(NC%), "DOMI-CLI", 0, 0)
    Call PRINTFORWIN(TRIM$(CPOSCLI$(NC%) + " " + LOCACLI$(NC%)), "LOCA-CLI", 0, 0)
    FF% = HOY(HOS$)
    Call PRINTFORWIN(HOS$, "FECH-EMI", 0, 0)
    Call PRINTFORWIN(TRIM$(Str$(NODES&)) + " / " + TRIM$(Str$(NUHOJ%)), "NUME-COM", 0, 0)
    Call PRINTFORWIN(TRIM$(Str$(NPED&)), "NRO-OCOM", 0, 0)
    '
    For RPE& = ULTREG&("PEDENCA") To 1 Step -1
       PENCX$ = REGLEIDO$("PEDENCA", RPE&)
       If (Abs(CVS(Left$(PENCX$, 4))) - NPED&) < 0.05 Then
         Call PRINTFORWIN(Mid$(PENCX$, 57, 16), "N-OCOMPR", 0, 0)
         Call PRINTFORWIN(Mid$(PENCX$, 73, 32), "DOMI-TRA", 0, 0)
         Call PRINTFORWIN(Mid$(PENCX$, 105, 32), "DESTINO", 0, 0)
         Call PRINTFORWIN(Mid$(PENCX$, 137, 32), "TRANSPOR", 0, 0)
         CPAGO% = Asc(Mid$(PENCX$, 269, 1))
         CONDIPAG$ = DECONPA$(CPAGO%)
         Call PRINTFORWIN(CONDIPAG$, "COND-PAG", 0, 0)
         GoTo 259
       End If
    Next RPE&
    '
259 YAC = DESMM
    '
    Return
    '
End Sub
'

Sub ANULODESPA(NODES&)
   '
   Call BLOQARCH("ORDESPA")
   Screen.MousePointer = 11
   J& = 0
   For I& = 1 To ULTREG&("ORDESPA")
        Y$ = REGLEIDO$("ORDESPA", I&)
        If CVS(Left$(Y$, 4)) <> NODES& Then
           J& = J& + 1
           Call GRAREG("ORDESPA", Y$, J&)
        End If
   Next I&
   Call SETULTREG("ORDESPA", J&)
   '
   Call CIERRARCH("*.*")
   Screen.MousePointer = 1
   '
End Sub
'

Sub CALPESTOCE()
   '
   Screen.MousePointer = 11
   JJ& = 0
   FIN& = ULTREG&("PEDDETP")
   '
   Dim CAPE(16384), VAPE#(16384)
   '
   MODO% = COMALTER%("Clasificación por Fecha de Entrega Solicitada:\\Mes en Curso\Totales Pendientes")
   Call HEADERS("PESTOCE", "")
   If MODO% = 1 Then
         Call HEADERS("PESTOCE", "Alcance: Mes en Curso")
      ElseIf MODO% = 2 Then
         Call HEADERS("PESTOCE", "Alcance: Total de Pedidos Pendientes")
      Else
         Exit Sub
   End If
   '
   FLI1% = -1: FLI2% = FINMESACT%
   If MODO% = 2 Then
      FLI1% = -1: FLI2% = 32767
   End If
   '
   OPX% = ALTERNA%("Solo Cantidades\Cantidades e Importes")
   '
   For U& = 1 To FIN&
      Call TRACE(20, U&, FIN&)
      X$ = REGLEIDO$("PEDDETP", U&)
      FENTRESOL% = CVI(Mid$(X$, 25, 2))
      If FENTRESOL% >= FLI1% Then
         If FENTRESOL% <= FLI2% Then
            CAPEN = CVS(Mid$(X$, 79, 4)) - CVS(Mid$(X$, 95, 4))
            If CAPEN < 0 Then CAPEN = 0
            If CAPEN > 0 Then
               CI% = CVI(Mid$(X$, 33, 2))
               If CI% > 0 Then
                  If CI% <= 16384 Then
                     NC% = CVI(Mid$(X$, 7, 2))
                     PREUNI = CVS(Mid$(X$, 83, 4)) * (100 - PORDESCLI(NC%)) / 100
                     If OPX% <> 2 Then PREUNI = 0
                     ITOTA# = CDbl(CAPEN * PREUNI)
                     '
                     CAPE(CI%) = CAPE(CI%) + CAPEN
                     VAPE#(CI%) = VAPE#(CI%) + ITOTA#
                  End If
               End If
            End If
         End If
      End If
   Next U&
   '
   NMS& = NUMAS%
   For U& = 1 To NMS&         '
      Call TRACE(20, U&, NMS&)
      CI% = U&
      If CAPE(CI%) >= 0.05 Then
         STACT = STODEPOS(CI%, DEPOEXPE%)
         If STACT < CAPE(CI%) Then
            FALTA = CAPE(CI%) - STACT
            Y$ = REGBLAN$("PESTOCE")
            Call REPLA(Y$, MKI$(CI%), 1, 2)
            Call REPLA(Y$, MKS$(STACT), 3, 4)
            Call REPLA(Y$, MKS$(CAPE(CI%)), 7, 4)
            Call REPLA(Y$, MKD$(VAPE#(CI%)), 11, 8)
            Call REPLA(Y$, MKS$(FALTA), 19, 4)
            Call REPLA(Y$, MKD$(VAPE#(CI%) * FALTA / CAPE(CI%)), 23, 8)
            JJ& = JJ& + 1
            Call GRAREG("PESTOCE", Y$, JJ&)
         End If
      End If
   Next U&
   '
   Call SETULTREG("PESTOCE", JJ&)
   Call CIERRARCH("PESTOCE")
   Screen.MousePointer = 1
   '
End Sub
'

Sub REVIPEDKIT()
   '
   Screen.MousePointer = 11
   '
   Dim KIT%(16384)
   Dim CAPE(16384)
   Dim VAPE#(16384)
   Dim TOCO(16384)
   Dim RESER(16384)
   '
   FIN& = NUMAS%
   For U& = 1 To FIN&
      Call TRACE(20, U&, FIN&)
      CIII% = U&
      If InStr(UCase$(DESCRIT$(CIII%)), "KIT") = 1 Then
         KIT%(CIII%) = 1
      End If
   Next U&
   '
   FIN& = ULTREG&("PEDDETP")
   '
   MODO% = 2
   If MODO% = 2 Then
      FLI1% = -1: FLI2% = 32767
   End If
   '
   For U& = 1 To FIN&
      Call TRACE(20, U&, FIN&)
      X$ = REGLEIDO$("PEDDETP", U&)
      CISX% = CVI(Mid$(X$, 33, 2))
      If CISX% > 0 Then
         If CISX% <= 16384 Then
            If KIT%(CISX%) = 1 Then
               FENTRESOL% = CVI(Mid$(X$, 25, 2))
               If FENTRESOL% >= FLI1% Then
                  If FENTRESOL% <= FLI2% Then
                     CAPEN = CVS(Mid$(X$, 79, 4)) - CVS(Mid$(X$, 95, 4))
                     If CAPEN < 0 Then CAPEN = 0
                     If CAPEN > 0 Then
                        NC% = CVI(Mid$(X$, 7, 2))
                        PREUNI = CVS(Mid$(X$, 83, 4)) * (100 - PORDESCLI(NC%)) / 100
                        ITOTA# = CDbl(CAPEN * PREUNI)
                        '
                        CAPE(CISX%) = CAPE(CISX%) + CAPEN
                        VAPE#(CISX%) = VAPE#(CISX%) + ITOTA#
                     End If
                  End If
               End If
            End If
         End If
      End If
   Next U&
   '
   NMS& = NUMAS%
   JJ& = 0: KK& = 0: LL& = 0
   For U& = 1 To NMS&         '
      Call TRACE(20, U&, NMS&)
      CISX% = U&
      If KIT%(CISX%) = 1 Then
         If CAPE(CISX%) >= 0.05 Then
            STACT = STODEPOS(CISX%, DEPOEXPE%)
            FALTA = 0
            If STACT < CAPE(CISX%) Then
               FALTA = CAPE(CISX%) - STACT
            End If
            Y$ = REGBLAN$("PEDIKIT")
            Call REPLA(Y$, MKI$(CISX%), 1, 2)
            Call REPLA(Y$, MKS$(STACT), 3, 4)
            Call REPLA(Y$, MKS$(CAPE(CISX%)), 7, 4)
            Call REPLA(Y$, MKD$(VAPE#(CISX%)), 11, 8)
            Call REPLA(Y$, MKS$(FALTA), 19, 4)
            Call REPLA(Y$, MKD$(VAPE#(CISX%) * FALTA / CAPE(CISX%)), 23, 8)
            JJ& = JJ& + 1
            Call GRAREG("PEDIKIT", Y$, JJ&)
            '
            If FALTA >= 0.05 Then
               CODKI$ = CODEXT$(CISX%)
               Call LEEEXPLOMUL(CODKI$, 1, 2)
               If CAIT% < 1 Then
                  Call MENSERR(24, "Kit " + TRIM$(CODKI$) + "\" + DESCRIT$(CISX%) + "\   \Falta Fórmula de Composición de este Kit.")
               End If
               '
               Z$ = REGBLAN$("NECOMKIT")
               Call REPLA(Z$, MKI$(CISX%), 1, 2)
               Call REPLA(Z$, DESCRIT$(CISX%), 3, 64)
               Call REPLA(Z$, MKD$(CDbl(FALTA)), 67, 8)
               KK& = KK& + 1
               Call GRAREG("NECOMKIT", Z$, KK&)
               '
               Call REPLA(Z$, Space$(64), 3, 64)
               Call REPLA(Z$, MKD$(0), 67, 8)
               KK& = KK& + 1
               Call GRAREG("NECOMKIT", Z$, KK&)
               '
               MAXAR# = 99999999
               For ITI% = 1 To CAIT%
                  NECOM# = FALTA * USOI(ITI%)
                  STOAC# = STODEPOS(CIJ%(ITI%), DEPOEXPE%)
                  STASI# = STOCKASI(CIJ%(ITI%))
                  STDIS# = STOAC# - STASI#: If STDIS# < 0 Then STDIS# = 0
                  FALCO# = NECOM# - STDIS#: If FALCO# < 0 Then FALCO# = 0
                  TOCO(CIJ%(ITI%)) = TOCO(CIJ%(ITI%)) + FALTA * USOI(ITI%)
                  '
                  Call REPLA(Z$, CODEXT$(CIJ%(ITI%)), 3, 16)
                  Call REPLA(Z$, DESCRIT$(CIJ%(ITI%)), 19, 48)
                  Call REPLA(Z$, MKD$(NECOM#), 67, 8)
                  Call REPLA(Z$, MKD$(STOAC#), 75, 8)
                  Call REPLA(Z$, MKD$(STASI#), 83, 8)
                  Call REPLA(Z$, MKD$(STDIS#), 91, 8)
                  Call REPLA(Z$, MKD$(FALCO#), 99, 8)
                  '
                  KK& = KK& + 1
                  Call GRAREG("NECOMKIT", Z$, KK&)
                  '
                  MAXAS# = (STDIS# - RESER(CIJ%(ITI%))) / USOI(ITI%)
                  If MAXAS# < MAXAR# Then
                     MAXAR# = MAXAS#
                  End If
               Next ITI%
               '
               MAXAR# = Int(MAXAR#)
               ARMAR# = FALTA
               If MAXAR# < ARMAR# Then
                  ARMAR# = MAXAR#
               End If
               ARMAR# = Int(ARMAR#)
               '
               T$ = REGBLAN$("ARMAKIT")
               Call REPLA(T$, MKI$(CISX%), 1, 2)
               Call REPLA(T$, MKD$(CDbl(FALTA)), 3, 8)
               Call REPLA(T$, MKD$(MAXAR#), 11, 8)
               Call REPLA(T$, MKD$(ARMAR#), 19, 8)
               LL& = LL& + 1
               Call GRAREG("ARMAKIT", T$, LL&)
               '
               For ITI% = 1 To CAIT%
                  RESER(CIJ%(ITI%)) = RESER(CIJ%(ITI%)) + ARMAR# * USOI(ITI%)
               Next ITI%
               '
            End If
         End If
      End If
   Next U&
   '
   Call SETULTREG("PEDIKIT", JJ&)
   Call CIERRARCH("PEDIKIT")
   Call SETULTREG("NECOMKIT", KK&)
   Call CIERRARCH("NECOMKIT")
   Call SETULTREG("ARMAKIT", LL&)
   Call CIERRARCH("ARMAKIT")
   '
   NMS& = NUMAS%
   JJ& = 0: KK& = 0
   For U& = 1 To NMS&         '
      Call TRACE(20, U&, NMS&)
      CISX% = U&
      If TOCO(CISX%) >= 0.05 Then
         NECOM# = TOCO(CISX%)
         STOAC# = STODEPOS(CISX%, DEPOEXPE%)
         STASI# = STOCKASI(CISX%)
         STDIS# = STOAC# - STASI#: If STDIS# < 0 Then STDIS# = 0
         FALCO# = NECOM# - STDIS#: If FALCO# < 0 Then FALCO# = 0
         '
         Z$ = REGBLAN$("TOCOMKIT")
         Call REPLA(Z$, MKI$(CISX%), 1, 2)
         Call REPLA(Z$, MKD$(NECOM#), 3, 8)
         Call REPLA(Z$, MKD$(STOAC#), 11, 8)
         Call REPLA(Z$, MKD$(STASI#), 19, 8)
         Call REPLA(Z$, MKD$(STDIS#), 27, 8)
         Call REPLA(Z$, MKD$(FALCO#), 35, 8)
         '
         KK& = KK& + 1
         Call GRAREG("TOCOMKIT", Z$, KK&)
      End If
   Next U&
   '
   Call SETULTREG("TOCOMKIT", KK&)
   Call CIERRARCH("TOCOMKIT")
   '
   For II% = 1 To NUMAS%
      CAPE(II%) = 0
      TOCO(II%) = 0
      RESER(II%) = 0
   Next II%
   '
   CAMESES = 6
   HOII = HOY(HO$)
   M0% = Val(Mid$(HO$, 4, 2)): A0% = Val(Mid$(HO$, 7, 2))
   A2% = A0%: M2% = M0% - 1: If M2% < 1 Then M2% = M2% + 12: A2% = A2% - 1
   A1% = A2%: M1% = M2% - CAMESES + 1: If M1% < 1 Then M1% = M1% + 12: A1% = A1% - 1
   If A2% < 0 Then A2% = A2% + 100
   If A1% < 0 Then A1% = A1% + 100
   '
   MS1$ = TRIM$(Str$(M1%)): While Len(MS1$) < 2: MS1$ = "0" + MS1$: Wend
   AS1$ = TRIM$(Str$(A1%)): While Len(AS1$) < 2: AS1$ = "0" + AS1$: Wend
   DXS$ = "01/" + MS1$ + "/" + AS1$: DES% = FECHANUM(DXS$)
   '
   MS2$ = TRIM$(Str$(M2%)): While Len(MS2$) < 2: MS2$ = "0" + MS2$: Wend
   AS2$ = TRIM$(Str$(A2%)): While Len(AS2$) < 2: AS2$ = "0" + AS2$: Wend
   HXS$ = "31/" + MS2$ + "/" + AS2$: HAS% = FECHANUM(HXS$)
   '
   FIN& = ULTREG&("MOVISTO")
   For I& = 1 To FIN&
      Call TRACE(20, I&, FIN&)
      MVX$ = REGLEIDO$("MOVISTO", I&)
      X$ = MVX$
      '
      FF% = CVI(Left$(X$, 2))
      If FF% >= DES% Then
        If FF% <= HAS% Then
          CISX% = CVI(Mid$(X$, 3, 2))
          If CISX% > 0 Then
            If CISX% <= NUMAS% Then
              If KIT%(CISX%) = 1 Then
                COMO$ = UCase$(Mid$(X$, 33, 2))
                If COMO$ = "FC" Or COMO$ = "NC" Or COMO$ = "ND" Then
                  CAPE(CISX%) = CAPE(CISX%) + CVD(Mid$(X$, 96, 8)) - CVD(Mid$(X$, 88, 8))
                End If
              End If
            End If
          End If
        End If
      End If
   Next I&
   '
   NMS& = NUMAS%
   For U& = 1 To NMS&         '
      Call TRACE(20, U&, NMS&)
      CISX% = U&
      If KIT%(CISX%) = 1 Then
        If CAPE(CISX%) >= 0.05 Then
          '
          FALTA = CAPE(CISX%) / CAMESES
          CODKI$ = CODEXT$(CISX%)
          Call LEEEXPLOMUL(CODKI$, 1, 2)
          If CAIT% < 1 Then
             Call MENSERR(24, "Kit " + TRIM$(CODKI$) + "\" + DESCRIT$(CISX%) + "\   \Falta Fórmula de Composición de este Kit.")
          End If
          '
          For ITI% = 1 To CAIT%
            TOCO(CIJ%(ITI%)) = TOCO(CIJ%(ITI%)) + FALTA * USOI(ITI%)
'If TRIM$(CODEXT$(CIJ%(ITI%))) = "S02544.00V" Then
'FFF = FALTA * USOI(ITI%)
'End If
          Next ITI%
          '
        End If
      End If
   Next U&
   '
   KK& = 0
   For U& = 1 To NMS&         '
      Call TRACE(20, U&, NMS&)
      CISX% = U&
      If TOCO(CISX%) >= 0.05 Then
        NECOM# = TOCO(CISX%)
        STOAC# = STODEPOS(CISX%, DEPOEXPE%)
        STASI# = STOCKASI(CISX%)
        STDIS# = STOAC# - STASI#: If STDIS# < 0 Then STDIS# = 0
        '
        Z$ = REGBLAN$("NEMENPIE")
        Call REPLA(Z$, MKI$(CISX%), 1, 2)
        Call REPLA(Z$, MKD$(NECOM#), 3, 8)
        Call REPLA(Z$, MKD$(STOAC#), 11, 8)
        Call REPLA(Z$, MKD$(STASI#), 19, 8)
        Call REPLA(Z$, MKD$(STDIS#), 27, 8)
        COBE = STDIS# / NECOM#: If COBE > 999 Then COBE = 999
        Call REPLA(Z$, MKS$(COBE), 35, 4)
        '
        KK& = KK& + 1
        Call GRAREG("NEMENPIE", Z$, KK&)
        '
      End If
   Next U&
   '
   Call SETULTREG("NEMENPIE", KK&)
   Call CIERRARCH("NEMENPIE")
   Call CIERRARCH("*.*")
   Call FINAL("*SORKIT1")
   '
   Call GRACONTROL("", "FEREKITS", HO$)
   Call GRACONTROL("", "HOREKITS", Left$(Time$, 5))
   FANAKITS.Label21.Caption = HO$
   FANAKITS.Label22.Caption = Left$(Time$, 5)
   
   Screen.MousePointer = 1
   '
End Sub
'

Sub COPEFAC()
   '
   JJ& = 0
   FIN& = ULTREG("PEDDETA")
   For U& = 1 To FIN&
      Call TRACE(20, U&, FIN&)
      XX$ = REGLEIDO$("PEDDETA", U&)
      NC% = CVI(Mid$(XX$, 7, 2))
      CI% = CVI(Mid$(XX$, 33, 2))
      FF% = CVI(Mid$(XX$, 5, 2))
      NF = CVS(Left$(XX$, 4))
      DOCU$ = "PC-" + TRIM$(Str$(NF))
      CAPED = CVS(Mid$(XX$, 79, 4))
      SALPE = CAPED - CVS(Mid$(XX$, 95, 4))
      '
      YY$ = REGBLAN$("COPEFAC")
      Call REPLA(YY$, MKI$(NC%), 1, 2)
      Call REPLA(YY$, MKI$(CI%), 3, 2)
      Call REPLA(YY$, MKI$(FF%), 5, 2)
      Call REPLA(YY$, DOCU$, 7, 12)
      Call REPLA(YY$, MKS$(SALPE), 19, 4)
      Call REPLA(YY$, MKS$(CAPED), 27, 4)
      '
      JJ& = JJ& + 1
      Call GRAREG("COPEFAC", YY$, JJ&)
      '
   Next U&
   '
   FIN& = ULTREG("MOVISTO")
   For U& = 1 To FIN&
      Call TRACE(20, U&, FIN&)
      XX$ = REGLEIDO$("MOVISTO", U&)
      DOPRI$ = Mid$(XX$, 23, 10): DOSEC$ = Mid$(XX$, 33, 10)
      If Left$(DOPRI$, 2) = "FC" Or Left$(DOSEC$, 2) = "FC" Then
         NC% = CVI(Mid$(XX$, 84, 2))
         CI% = CVI(Mid$(XX$, 3, 2))
         FF% = CVI(Mid$(XX$, 1, 2))
         DOCU$ = DOPRI$
         CAENT = CVD(Mid$(XX$, 96, 8))
         '
         YY$ = REGBLAN$("COPEFAC")
         Call REPLA(YY$, MKI$(NC%), 1, 2)
         Call REPLA(YY$, MKI$(CI%), 3, 2)
         Call REPLA(YY$, MKI$(FF%), 5, 2)
         Call REPLA(YY$, DOCU$, 7, 12)
         Call REPLA(YY$, MKS$(CAENT), 31, 4)
         '
         JJ& = JJ& + 1
         Call GRAREG("COPEFAC", YY$, JJ&)
         '
      End If
   Next U&
   '
   Call SETULTREG("COPEFAC", JJ&)
   Call CIERRARCH("COPEFAC")
   Call CIERRARCH("*.*")
   Call FINAL("*SOPEFAC")
   '
End Sub
'

Sub COPEDPE()
   '
   Screen.MousePointer = 11
   Dim NPED&(16384)
   CAPEP% = 0
   FIN& = ULTREG&("PEDENCA")
   For I& = 1 To FIN&
      Call TRACE(20, I&, FIN&)
      X$ = REGLEIDO$("PEDENCA", I&)
      If Asc(Mid$(X$, 267, 1)) <= 1 Then
         CAPEP% = CAPEP% + 1
         NPED&(CAPEP) = CVS(Left$(X$, 4))
      End If
   Next I&
   '
   NPXA& = 0
   FIN& = ULTREG("PEDDETP")
   For U& = 1 To FIN&
      Call TRACE(20, U&, FIN&)
      XX$ = REGLEIDO$("PEDDETP", U&)
      NPX& = CVS(Left$(XX$, 4))
      If NPX& = NPXA& Then GoTo 19
      '
      NPXA& = NPX&
      For K% = 1 To CAPEP%
         If NPED&(K%) = NPX& Then GoTo 19
      Next K%
      '
      RECU% = COMALTER%("Pedido" + Str$(NPX&) + " Error de Cabecera.\\Restaurar\Continuar Control\Cancelar")
      Screen.MousePointer = 11
      If RECU% = 1 Then
          For L& = 1 To ULTREG&("PEDENCA")
             YY$ = REGLEIDO$("PEDENCA", L&)
             If CVS(Left$(YY$, 4)) = NPX& Then
               Call REPLA(YY$, Chr$(1), 267, 1)
               Call GRAREG("PEDENCA", YY$, L&)
               GoTo 19
             End If
          Next L&
        ElseIf RECU% = 3 Then
          GoTo 99
      End If
19 Next U&
   '
99 Call CIERRARCH("*.*")
   Screen.MousePointer = 1
   Call COMUNI("Control Completo")
   '
End Sub
'
Sub NUPEDIDO(NROPRO&)
    '
    NROPRO& = 1
    If ULTREG&("PEDENCA") > 0 Then
        X$ = REGLEIDO$("PEDENCA", ULTREG&("PEDENCA"))
        If Len(X$) >= 4 Then NROPRO& = 1 + CVS(Left$(X$, 4))
    End If
    '
End Sub
'
Sub CARDATPEDMAI(OKEY%)
   '
   HOII% = HOY(HOS$)
   
   OKEY% = 1
   If Val(FORCARPED.Text1.TEXT) > 0 Then
      If Val(FORCARPED.Text1.TEXT) <= 32767 Then
         NC% = Val(FORCARPED.Text1.TEXT)
         If ECOARCH$("DEUDOR1", MKI$(NC%)) <> "" Then
            '
15          PPIVA% = PIVACLI%(NC%)
            If PPIVA% > 0 And PPIVA% <> 5 Then
               CUITT$ = TRIM$(CUITCLI$(NC%))
               If CUITT$ = "" Or VALICUIT%(CUITT$) <> 1 Then
                  Call VALDATCLI(NC%, OK2%)
                  If OK2% <> 1 Then Call FINAL("")
                  GoTo 15
               End If
            End If
            '
            FORCARPED.Text2.TEXT = RASOCLI$(NC%)
            FORCARPED.Text3.TEXT = DOMICLI$(NC%)
            FORCARPED.Text4.TEXT = CPOSCLI$(NC%)
            FORCARPED.Text5.TEXT = LOCACLI$(NC%)
            '
            PEDX$ = String$(480, 0)
            NROPE& = Val(Mid$(FORCARPED.Text11.TEXT, 4))
            Call REPLA(PEDX$, MKS$(NROPE&), 1, 4)
            Call REPLA(PEDX$, MKI$(HOII%), 5, 2)
            Call REPLA(PEDX$, MKI$(NC%), 7, 2)
            Call REPLA(PEDX$, RASOCLI$(NC%), 9, 32)
            Call REPLA(PEDX$, DOMICLI$(NC%), 41, 32)
            CLOCA$ = TRIM$(CPOSCLI$(NC%)): If CLOCA$ <> "" Then CLOCA$ = CLOCA$ + " "
            Call REPLA(PEDX$, CLOCA$ + LOCACLI$(NC%), 73, 32)
            Call REPLA(PEDX$, TRIM$(CUITCLI$(NC%)) + " " + TRIM$(POSIVAC$(NC%)), 105, 32)
            Call REPLA(PEDX$, Space$(160), 137, 160) ' Nro O/Compra
            '
            Call REPLA(PEDX$, Chr$(LIPRCLI%(NC%)), 297, 1)
            Call REPLA(PEDX$, Chr$(CPAGCLI%(NC%)), 298, 1)
            Call REPLA(PEDX$, Chr$(VENDCLI%(NC%)), 299, 1)
            '
100         ENCAPED$ = OBJPLA$("PEDIDO-ENCMAI", PEDX$)
            If ENCAPED$ = "" Then
               OKEY% = 0
               Exit Sub
            End If
            PEDX$ = ENCAPED$
            '
            LPRE% = Asc(Mid$(ENCAPED$, 297, 1))
            If LPRE% < 1 Or ECOARCH$("LISTAS", Chr$(LPRE%)) = "" Then
               Call MENSERR(24, "Lista de Precios no Válida")
               GoTo 100
            End If
            '
            CPAG% = Asc(Mid$(ENCAPED$, 298, 1))
            If CPAG% < 1 Or ECOARCH$("CONPAG", Chr$(CPAG%)) = "" Then
               Call MENSERR(24, "Condición de Pago no Válida")
               GoTo 100
            End If
            CONPA$ = DECONPA$(CPAG%)
            '
            VEND% = Asc(Mid$(ENCAPED$, 299, 1))
            If VEND% > 0 Then
               If ECOARCH$("VENDEDOR", Chr$(VEND%)) = "" Then
                  Call MENSERR(24, "Vendedor no Figura en Tablas")
                  GoTo 100
               End If
            End If
            VENDE$ = DEVENDE$(VEND%)
            '
            ARCHIPRE$ = TRIM$(Str$(LPRE%))
            While Len(ARCHIPRE$) < 3: ARCHIPRE$ = "0" + ARCHIPRE$: Wend
            ARCHIPRE$ = "LIPRE" + ARCHIPRE$
            '
            NOCOM$ = Mid$(ENCAPED$, 137, 18)
            Call BLANARCH("!CARCONPE")
            Call BLANARCH("!CARITPED")
            Call CARCONJPED
            OKEY% = 1
            Exit Sub
         End If
      End If
   End If
   '
   Call MENSERR(24, "Numero de Cliente no Válido\o Inexistente")
   FORCARPED.Text1.TEXT = ""
   '
End Sub

Sub CARCONJPED()
    '
200 ATRI$ = "/NC"
    ATRI$ = ATRI$ + "/TITUPAN=Ingreso Conjuntos a Itemizar"
    ATRI$ = ATRI$ + "/BORDESUP=W" + TRIM$(Str$(FORCARPED.Top + FORCARPED.Picture1.Top + 100))
    ATRI$ = ATRI$ + "/BORDEIZQ=W" + TRIM$(Str$(FORCARPED.Left + FORCARPED.Picture1.Left))
    '
    VTB% = VENTABU%("CARCONPE" + ATRI$)
    Call CIERRARCH("!CARCONPE")
    If VTB% < 0 Then
       '
       For J& = 1 To FORCARPED.List1.ListCount
          Z$ = FORCARPED.List1.List(J& - 1)
          CACON = Val(Mid$(Z$, 4, 7))
          RECON$ = Mid$(Z$, 13, 42)
          FFI% = FECHANUM(Mid$(Z$, 57, 8))
          PREUCO = Val(Mid$(Z$, 65, 9))
          ITOTCO# = Val(Mid$(Z$, 74, 10))
          ITID% = Val(Left$(Z$, 3))
          '
          X$ = REGBLAN$("CARCONPE")
          Call REPLA(X$, MKS$(CACON), 1, 4)
          Call REPLA(X$, RECON$, 5, 58)
          Call REPLA(X$, MKI$(FFI%), 63, 2)
          Call REPLA(X$, MKS$(PREUCO), 65, 4)
          Call REPLA(X$, MKD$(ITOTCO#), 69, 8)
          Call REPLA(X$, MKI$(ITID%), 127, 2)
          Call GRAREG("!CARCONPE", X$, J&)
       Next J&
       Call SETULTREG("!CARCONPE", FORCARPED.List1.ListCount)
       Exit Sub
       '
    End If
    '
    REBLA$ = REGBLAN$("CARCONPE")
    FORCARPED.List1.Clear
    TEIT$ = ""
    For U& = 1 To ULTREG&("!CARCONPE")
        X$ = REGLEIDO$("!CARCONPE", U&)
        ITID% = CVI(Mid$(X$, 127, 2))
        While Len(TEIT$) < ITID%
           TEIT$ = TEIT$ + Chr$(0)
        Wend
        If ITID% > 0 Then Mid$(TEIT$, ITID%) = Chr$(1)
    Next U&
    '
    J& = 0: K& = 0
    For U& = 1 To ULTREG&("!CARCONPE")
        X$ = REGLEIDO$("!CARCONPE", U&)
        ITID% = CVI(Mid$(X$, 127, 2))
205     If ITID% < 1 Then
           TEIT$ = TEIT$ + Chr$(0)
           ITID% = InStr(TEIT$, Chr$(0))
           Mid$(TEIT$, ITID%) = Chr$(1)
           Call REPLA(X$, MKI$(ITID%), 127, 2)
           Call GRAREG("!CARCONPE", X$, U&)
        End If
        '
        CACO = CVS(Left$(X$, 4))
        If CACO < 0.1 Then GoTo 240
        '
        ITOTCO# = 0
        For T& = 1 To ULTREG&("!CARITPED")
           Y$ = REGLEIDO$("!CARITPED", T&)
           If CVI(Left$(Y$, 2)) = ITID% Then
              CAIT = CVS(Mid$(Y$, 5, 4))
              PRECO = CVS(Mid$(Y$, 17, 4))
              IMPO# = CACO * CAIT * PRECO
              Call REPLA(Y$, MKD$(IMPO#), 21, 8)
              J& = J& + 1
              Call GRAREG("!CARITPEQ", Y$, J&)
              ITOTCO# = ITOTCO# + IMPO#
           End If
        Next T&
        '
        PREUCO = ITOTCO# / CACO
        Call REPLA(X$, MKS$(PREUCO), 65, 4)
        Call REPLA(X$, MKD$(ITOTCO#), 69, 8)
        K& = K& + 1
        Call GRAREG("!CARCONPE", X$, K&)
        '
240 Next U&
    Call SETULTREG("!CARCONPE", K&)
    '
    For T& = 1 To J&
        Y$ = REGLEIDO$("!CARITPEQ", T&)
        Call GRAREG("!CARITPED", Y$, T&)
    Next T&
    Call SETULTREG("!CARITPED", J&)
    '
    TOPEDIDO# = 0
    For U& = 1 To ULTREG&("!CARCONPE")
        X$ = REGLEIDO$("!CARCONPE", U&)
        If X$ = REBLA$ Then GoTo 290
        CACON = CVS(Left$(X$, 4))
        If CACON < 1 Then GoTo 290
        '
        RECON$ = TRIM$(Mid$(X$, 5, 58))
        ITID% = CVI(Mid$(X$, 127, 2))
        If RECON$ = "" Then
           RECON$ = "Item Nro." + Str$(ITID%)
           Call REPLA(X$, RECON$, 5, 58)
           Call GRAREG("!CARCONPE", X$, U&)
        End If
        FEX = CVI(Mid$(X$, 63, 2))
        If FEX < 1 Then
           FEX = HOY(HO$)
           FFI% = FEX
           Call REPLA(X$, MKI$(FFI%), 63, 2)
           Call GRAREG("!CARCONPE", X$, U&)
        End If
        '
        PREUCO = CVS(Mid$(X$, 65, 4))
        ITOTCO# = CVD(Mid$(X$, 69, 8))
        '
        Z$ = Space$(96)
        Call REPLA(Z$, CSTRING$(MKI$(ITID%), 1, 3, 0, 2), 1, 3)
        Call REPLA(Z$, CSTRING$(MKS$(CACON), 3, 7, 0, 2), 4, 7)
        Call REPLA(Z$, RECON$, 13, 42)
        Call REPLA(Z$, FECHATEX$(FEX), 57, 8)
        Call REPLA(Z$, CSTRING$(MKS$(PREUCO), 3, 9, 2, 2), 65, 9)
        Call REPLA(Z$, CSTRING$(MKD$(ITOTCO#), 3, 10, 2, 2), 74, 10)
        FORCARPED.List1.AddItem Z$
        TOPEDIDO# = TOPEDIDO# + ITOTCO#
        '
290 Next U&
    '
    If FORCARPED.List1.ListCount > 0 Then
        FORCARPED.List1.ListIndex = 0
        FORCARPED.List1.SetFocus
    End If
    '
    ALI = ALIVA
    If PIVACLI%(NC%) = 2 Then
       ALI = ALIVA + ALIVARNI
    End If
    FORCARPED.Label14.Caption = CSTRING$(MKD$(TOPEDIDO#), 3, 11, 2, 2)
    FORCARPED.Label18.Caption = CSTRING$(MKD$(TOPEDIDO# * ALI / 100), 3, 11, 2, 2)
    FORCARPED.Label20.Caption = CSTRING$(MKD$(TOPEDIDO# * (1 + ALI / 100)), 3, 11, 2, 2)
End Sub
'

Sub CARITPED1()
   '
   If FORCARPED.List1.ListCount < 1 Then
      Exit Sub
   End If
   '
   ILIS1& = FORCARPED.List1.ListIndex
   If ILIS1& < 0 Or ILIS1& > FORCARPED.List1.ListCount - 1 Then
      FORCARPED.List1.ListIndex = 0
   End If
   '
   U& = FORCARPED.List1.ListIndex + 1
   X$ = REGLEIDO$("!CARCONPE", U&)
   ITID% = CVI(Mid$(X$, 127, 2))
   CACO = CVS(Left$(X$, 4))
   If ITID% < 1 Then
      Call MENSERR(24, "Imposible Seguir Proceso - Consulte")
      Call FINAL("")
   End If
   '
   FORCARPED.List2.Clear
   ITOTCO# = 0
   For T& = 1 To ULTREG&("!CARITPED")
      Y$ = REGLEIDO$("!CARITPED", T&)
      If CVI(Left$(Y$, 2)) = ITID% Then
         CAIT = CVS(Mid$(Y$, 5, 4))
         CIIT% = CVI(Mid$(Y$, 3, 2))
         PRELI = CVS(Mid$(Y$, 9, 4))
         PREBA = CVS(Mid$(Y$, 13, 4))
         PRECO = CVS(Mid$(Y$, 17, 4))
         IMPO# = CACO * CAIT * PRECO
         '
         Z$ = Space$(96)
         Call REPLA(Z$, CSTRING$(MKS$(CAIT), 3, 4, 0, 2), 1, 4)
         Call REPLA(Z$, CODEXT$(CIIT%), 7, 15)
         Call REPLA(Z$, DESCRIT0$(CIIT%), 23, 33)
         Call REPLA(Z$, CSTRING$(MKS$(PRELI), 3, 9, 2, 2), 56, 9)
         Call REPLA(Z$, CSTRING$(MKS$(PRECO), 3, 9, 2, 2), 65, 9)
         Call REPLA(Z$, CSTRING$(MKD$(IMPO#), 3, 10, 2, 2), 74, 10)
         ITOTCO# = ITOTCO# + IMPO#
         FORCARPED.List2.AddItem Z$
         '
      End If
   Next T&
   FORCARPED.Label10.Caption = CSTRING$(MKD$(ITOTCO#), 3, 12, 2, 2)
   '
End Sub
'

Sub PRIPEDCLI()
    '
    TTXX$ = Mid$(ENCAPED$, 301, 60): Call GRAREG("!OBSERVOF", TTXX$, 1)
    TTXX$ = Mid$(ENCAPED$, 361, 60): Call GRAREG("!OBSERVOF", TTXX$, 2)
    TTXX$ = Mid$(ENCAPED$, 421, 60): Call GRAREG("!OBSERVOF", TTXX$, 3)
    Call SETULTREG("!OBSERVOF", 3)
    Call CIERRARCH("!OBSERVOF")
    '
    FORIPRE$ = CONTROL$("", "FORPECLI")
    If FORIPRE$ = "" Then Exit Sub
    '
    Screen.MousePointer = 11
    HOII% = HOY(HOS$)
    FECHDOC$ = FORCARPED.Text9.TEXT
    NROPE& = Val(Mid$(FORCARPED.Text11.TEXT, 4))
    NUMEDOC$ = TRIM$(Str$(NROPE&))
    While Len(NUMEDOC$) < 8
      NUMEDOC$ = "0" + NUMEDOC$
    Wend
    NUMEDOC$ = "0001-" + NUMEDOC$
    TIPODOC$ = "Pedido de Cliente"
    '
    CODTABU1$(1) = "COD-MAT"
    CODTABU1$(2) = "DES-MAT"
    CODTABU1$(3) = "CANTIDAD"
    CODTABU1$(4) = "PRE-NETO"
    CODTABU1$(5) = "NET-ITEM"
    CODTABU1$(6) = "FECH-REC"
    CACOLTAB1% = 6
    '
    DESTIDOC% = NC%
    CAITAB1% = 0
    CAITCRU% = 0
    '
    SUMATOTA# = 0: J& = 0
    For U& = 1 To ULTREG&("!CARCONPE")
       X$ = REGLEIDO$("!CARCONPE", U&)
       ITID% = CVI(Mid$(X$, 127, 2))
       CACO = CVS(Left$(X$, 4))
       If CACO < 0.1 Then GoTo 240
       '
       DECO$ = Mid$(X$, 5, 58)
       FEX = CVI(Mid$(X$, 63, 2))
       PREU = CVS(Mid$(X$, 65, 4))
       ITOT# = CVD(Mid$(X$, 69, 8))
       '
       FENTX$ = FECHATEX$(FEX)
       CAPEX$ = Str$(CACO)
       PREUX$ = Str$(PREU)
       ITOTX$ = Str$(ITOT#)
       SUMATOTA# = SUMATOTA# + CDbl(Val(ITOTX$))
       '
       J& = J& + 1: JJX$ = TRIM$(Str$(J&)): While Len(JJX$) < 2: JJX$ = "0" + JJX$: Wend
       CAITAB1% = CAITAB1% + 1
       TETABU1$(1, CAITAB1%) = JJX$
       TETABU1$(2, CAITAB1%) = DECO$
       TETABU1$(3, CAITAB1%) = CAPEX$
       TETABU1$(4, CAITAB1%) = PREUX$
       TETABU1$(5, CAITAB1%) = ITOTX$
       TETABU1$(6, CAITAB1%) = FENTX$
       '
       ITOTCO# = 0
       For T& = 1 To ULTREG&("!CARITPED")
           Y$ = REGLEIDO$("!CARITPED", T&)
           If CVI(Left$(Y$, 2)) = ITID% Then
              CIIT% = CVI(Mid$(Y$, 3, 2))
              CAIT = CVS(Mid$(Y$, 5, 4))
              Z$ = Space$(80)
              Call REPLA(Z$, ".", 1, 1)
              Call REPLA(Z$, CSTRING$(MKS$(CAIT), 3, 6, 1, 2), 2, 6)
              Call REPLA(Z$, CODEXT$(CIIT%), 10, 15)
              Call REPLA(Z$, DESCRIT0$(CIIT%), 26, 55)
              '
              CAITAB1% = CAITAB1% + 1
              TETABU1$(1, CAITAB1%) = ""
              TETABU1$(2, CAITAB1%) = Z$
              TETABU1$(3, CAITAB1%) = ""
              TETABU1$(4, CAITAB1%) = ""
              TETABU1$(5, CAITAB1%) = ""
              TETABU1$(6, CAITAB1%) = ""
              '
           End If
        Next T&
        '
        CAITAB1% = CAITAB1% + 1
        TETABU1$(1, CAITAB1%) = ""
        TETABU1$(2, CAITAB1%) = String$(80, "-")
        TETABU1$(3, CAITAB1%) = ""
        TETABU1$(4, CAITAB1%) = ""
        TETABU1$(5, CAITAB1%) = ""
        TETABU1$(6, CAITAB1%) = ""
        '
240 Next U&
     '
     CODPARAM$(1) = "IMP-NETO"
     DOCPARAM$(1) = Str$(SUMATOTA#)
     CODPARAM$(2) = "VENDEDOR"
     DOCPARAM$(2) = VENDE$
     CODPARAM$(3) = "COND-PAG"
     DOCPARAM$(3) = CONPA$
     CODPARAM$(4) = "NRO-OCOM"
     DOCPARAM$(4) = NOCOM$
     CAPARDOC% = 4
     '
     Call PRINTDOC("FORPECLI")   ' PEDIDO DE CLIENTE
     '
End Sub

Sub GRAPEDCLI(OKEY%)
    '
    HOII% = HOY(HO$)
    '
    OKEY% = 0
    If ULTREG&("!CARCONPE") < 1 Then
       Call MENSERR(24, "Pedido Vacío")
       Exit Sub
    End If
    '
    For U& = 1 To ULTREG&("!CARCONPE")
       X$ = REGLEIDO$("!CARCONPE", U&)
       CACO = CVS(Left$(X$, 4))
       ITID% = CVI(Mid$(X$, 127, 2))
       For T& = 1 To ULTREG&("!CARITPED")
         Y$ = REGLEIDO$("!CARITPED", T&)
         If CVI(Left$(Y$, 2)) = ITID% Then
           CAN = CACO * CVS(Mid$(Y$, 5, 4))
           If CAN > 0.05 Then
             CIIT% = CVI(Mid$(Y$, 3, 2))
             If CODEXT$(CIIT%) <> "" Then
               PREU = CVS(Mid$(Y$, 17, 4))
               If PREU >= 0.005 Then
                 GoTo 19
               End If
             End If
           End If
         End If
       Next T&
       '
       REFCON$ = TRIM$(Mid$(X$, 5, 56))
       Call MENSERR(24, "Faltan Datos en Item\" + REFCON$)
       Exit Sub
       '
19  Next U&
    '
    NC% = CVI(Mid$(ENCAPED$, 7, 2))
    LMCR# = LIMICRE(NC%)
    SDEU# = SALDCLI#(NC%)
    CHPA# = CHEPEACRE#(NC%)
    TPED# = Val(FORCARPED.Label20.Caption)
    If SDEU# + CHPA# + TPED# > LMCR# Then
       Call COMUNI("Este Pedido no Podrá Facturarse.\  \Supera el Límite de Crédito Asignado\para esta Cuenta de $ " + TRIM$(CSTRING$(MKD$(LMCR#), 3, 12, 2, 2)) + ".\  \Informe al Supervisor de Ventas.")
    End If
    '
    Call CIERRARCH("PEDCLIEN")
    Call CIERRARCH("PEDENCA")
    Call CIERRARCH("PEDDETA")
    Call CIERRARCH("PEDDETP")
    '
    Call BLOQARCH("PEDCLIEN")
    Call BLOQARCH("PEDENCA")
    Call BLOQARCH("PEDDETA")
    Call BLOQARCH("PEDDETP")
    '
    NROPED& = CVS(Left$(ENCAPED$, 4))
    NROPRO& = 0
    Call NUPEDIDO(NROPRO&)
    If NROPRO& <> NROPED& Then
       Call COMUNI("FLEXOFT ha Re-Numerado el Pedido\como Pedido Número '" + TRIM$(Str$(NROPRO&)) + "'.")
       NROPED& = NROPRO&
       Call REPLA(ENCAPED$, MKS$(NROPED&), 1, 4)
       FORCARPED.Text11.TEXT = "PC." + TRIM$(Str$(NROPRO&))
    End If
    '
    FEPED% = CVI(Mid$(ENCAPED$, 5, 2))
    NC% = CVI(Mid$(ENCAPED$, 7, 2))
    REFE$ = "Pedido Nro." + Str$(NROPED&) + " - " + RASOCLI$(NC%)
    NOCOM$ = Mid$(ENCAPED$, 137, 16)
    
    PDC$ = REGBLAN$("PEDENCA")
    Call REPLA(PDC$, MKS$(NROPED&), 1, 4)
    Call REPLA(PDC$, REFE$, 5, 48)
    Call REPLA(PDC$, MKI$(FEPED%), 53, 2)
    Call REPLA(PDC$, MKI$(NC%), 55, 2)
    Call REPLA(PDC$, Mid$(ENCAPED$, 137, 112), 57, 112)
    Call REPLA(PDC$, MKI$(0), 265, 2)
    Call REPLA(PDC$, Chr$(1), 267, 1)
    Call REPLA(PDC$, MKS$(1 + ULTREG&("PEDDETA")), 271, 4)
    Call REPLA(PDC$, String$(14, 0), 275, 14)
    Call REPLA(PDC$, CAVERI$, 289, 14)
    '
    Call REGAPP("PEDENCA", PDC$)
    REGENCA& = ULTREG&("PEDENCA")
    '
    For U& = 1 To ULTREG&("!CARCONPE")
       X$ = REGLEIDO$("!CARCONPE", U&)
       '
       FENTRESOL% = CVI(Mid$(X$, 63, 2))
       If FENTRESOL% < HOII% Then FENTRESOL% = HOII%
       CACO = CVS(Left$(X$, 4))
       ITID% = CVI(Mid$(X$, 127, 2))
       '
       For T& = 1 To ULTREG&("!CARITPED")
         Y$ = REGLEIDO$("!CARITPED", T&)
         If CVI(Left$(Y$, 2)) = ITID% Then
           Z$ = String$(1024, 0)
           Call REPLA(Z$, ENCAPED$, 1, 512)
           Call REPLA(Z$, X$, 513, 128)
           Call REPLA(Z$, Y$, 641, 64)
           Call REGAPP("PEDCLIEN", Z$)
           '
           CAN = CACO * CVS(Mid$(Y$, 5, 4))
           If CAN > 0.05 Then
              CIIT% = CVI(Mid$(Y$, 3, 2))
              PREU = CVS(Mid$(Y$, 17, 4))
              ITOT# = CVD(Mid$(Y$, 21, 8))
              PDC$ = REGBLAN$("PEDDETA")
              Call REPLA(PDC$, MKS$(NROPED&), 1, 4)
              Call REPLA(PDC$, MKI$(FEPED%), 5, 2)
              Call REPLA(PDC$, MKI$(NC%), 7, 2)
              Call REPLA(PDC$, NOCOM$, 9, 16)
              Call REPLA(PDC$, MKI$(FENTRESOL%), 25, 2)
              Call REPLA(PDC$, Chr$(1), 27, 1)
              Call REPLA(PDC$, Mid$(ENCAPED$, 297, 3), 28, 3)
              Call REPLA(PDC$, MKI$(0), 31, 2)
              Call REPLA(PDC$, MKI$(CIIT%), 33, 2)
              Call REPLA(PDC$, COLOR$(CIIT%), 35, 16)
              Call REPLA(PDC$, MKS$(CAN), 79, 4)
              Call REPLA(PDC$, MKS$(PREU), 83, 4)
              Call REPLA(PDC$, MKD$(ITOT#), 87, 8)
              Call REPLA(PDC$, String$(18, 0), 95, 18)
              Call REPLA(PDC$, MKS$(CAN), 99, 4)
              Call REPLA(PDC$, MKS$(REGENCA&), 109, 4)  ' puntero a ENCABEZADO
              Call REPLA(PDC$, CAVERI$, 113, 16)
              '
              Call REGAPP("PEDDETA", PDC$)
              '
              RGX& = ULTREG&("PEDDETA")
              Call REPLA(PDC$, MKS$(RGX&), 109, 4)  ' puntero a PEDDETA
              Call REGAPP("PEDDETP", PDC$)
              '
           End If
         End If
       Next T&
       '
240 Next U&
    '
    Call CIERRARCH("PEDCLIEN")
    Call CIERRARCH("PEDENCA")
    Call CIERRARCH("PEDDETA")
    Call CIERRARCH("PEDDETP")
    '
    Call CIERRARCH("*.*")
    OKEY% = 1
    '
End Sub
'

Sub REPIMPED()  ' repetir impresion de pedido
   '
10 Call SELECHO("DEUDOR1")
   NC% = Val(VALACT1$("DEUDOR1"))
   If NC% < 1 Then
      Exit Sub
   End If
   '
15 RFF$ = RECFILT$("PEDCLIEN", 3, MKI$(NC%))
   If Len(RFF$) < 4 Then
      Call MENSERR(24, "No hay Pedidos Registrados\para '" + TRIM$(RASOCLI$(NC%)) + "'")
      GoTo 10
   End If
   '
   Call BLANARCH("!INDIPEP")
   Call FRESHECHO("!INDIPEP")
   REFANT$ = "": NROPEDA& = 0: REFPED$ = ""
   FIN& = ULTREG&("PEDCLIEN")
   '
   Screen.MousePointer = 11
   For UK% = 1 To Len(RFF$) Step 4
      U& = CVS(Mid$(RFF$, UK%, 4))
      If U& > 0 Then
        If U& <= ULTREG&("PEDCLIEN") Then
          Call TRACE(20, I&, FIN&)
          X$ = REGLEIDO$("PEDCLIEN", U&)
          If CVI(Mid$(X$, 7, 2)) <> NC% Then
             X$ = REGLEIDO$("PEDCLIEN", 1)
             Call GRAREG("PEDCLIEN", X$, 1)
             Call CIERRARCH("PEDCLIEN")
             GoTo 15
          End If
          NROPED& = CVS(Left$(X$, 4))
20        If NROPED& > 0 Then
            If NROPED& <> NROPEDA& Then
               FEX = CVI(Mid$(X$, 5, 2))
               NC% = CVI(Mid$(X$, 7, 2))
               REFE$ = MKS$(NROPED&) + FECHATEX$(FEX) + "   " + RASOCLI$(NC%) + String$(56, 0)
               Call REPLA(REFE$, MKS$(U&), 53, 4)
               Call REGAPP("!INDIPEP", REFE$)
               NROPEDA& = NROPED&
            End If
          End If
        End If
      End If
   Next UK%
   Screen.MousePointer = 1
   '
30 Z$ = REGSEL$("!INDIPEP")
   If Len(Z$) <= 4 Then
      GoTo 10
   End If
   '
   NROPED& = CVS(Left$(Z$, 4))
   REGPED& = CVS(Mid$(Z$, 53, 4))
   '
   X$ = REGLEIDO$("PEDCLIEN", REGPED&)
   ENCAPED$ = Left$(X$, 512)
   '
   NC% = CVI(Mid$(X$, 7, 2))
   FEX = CVI(Mid$(X$, 5, 2))
   FORCARPED.Text9.TEXT = FECHATEX$(FEX)
   FORCARPED.Text11.TEXT = "PC." + TRIM$(Str$(NROPED&))
   '
   LPRE% = Asc(Mid$(ENCAPED$, 297, 1))
   CPAG% = Asc(Mid$(ENCAPED$, 298, 1))
   CONPA$ = DECONPA$(CPAG%)
   VEND% = Asc(Mid$(ENCAPED$, 299, 1))
   VENDE$ = DEVENDE$(VEND%)
   NOCOM$ = Mid$(ENCAPED$, 137, 18)
   '
   Call BLANARCH("!CARCONPE")
   Call BLANARCH("!CARITPED")
   '
   ITIDA% = (-1)
   For UU& = REGPED& To FIN&
      X$ = REGLEIDO$("PEDCLIEN", UU&)
      If CVS(Left$(X$, 4)) <> NROPED& Then GoTo 50
      '
      COPED$ = Mid$(X$, 513, 128)
      ITPED$ = Mid$(X$, 641, 64)
      '
      REFE$ = Mid$(COPED$, 5, 58)
      ITID% = CVI(Mid$(COPED$, 127, 2))
      If ITID% <> ITIDA% Then
         Call REGAPP("!CARCONPE", COPED$)
         ITIDA% = ITID%
      End If
      Call REGAPP("!CARITPED", ITPED$)
   Next UU&
   '
50 Call PRIPEDCLI
   GoTo 30
   '
End Sub
   
Sub ANULAPEDMAI()
   '
10 Call BLANARCH("!INDIPEP")
   '
   REFANT$ = "": NROPEDA& = 0: REFPED$ = ""
   FIN& = ULTREG&("PEDCLIEN")
   '
   Screen.MousePointer = 11
   For U& = 1 To FIN&
      Call TRACE(20, U&, FIN&)
      X$ = REGLEIDO$("PEDCLIEN", U&)
      STATUS% = Asc(Mid$(X$, 300, 1))
      If STATUS% <= 1 Then
         '
         CACOPED = CVS(Mid$(X$, 513, 4))
         CAITPED = CVS(Mid$(X$, 645, 4))
         CACOENT = CVS(Mid$(X$, 512 + 77, 4))
         CANPEN = (CACOPED - CACOENT) * CAITPED
         '
         If CANPEN > 0.05 Then
           NROPED& = CVS(Left$(X$, 4))
20         If NROPED& > 0 Then
             If NROPED& <> NROPEDA& Then
               FEX = CVI(Mid$(X$, 5, 2))
               NC% = CVI(Mid$(X$, 7, 2))
               REFE$ = MKS$(NROPED&) + FECHATEX$(FEX) + "   " + RASOCLI$(NC%) + String$(56, 0)
               Call REPLA(REFE$, MKS$(U&), 53, 4)
               Call REGAPP("!INDIPEP", REFE$)
               NROPEDA& = NROPED&
             End If
           End If
         End If
      End If
   Next U&
   Screen.MousePointer = 1
   '
30 Z$ = REGSEL$("!INDIPEP")
   If Len(Z$) <= 4 Then
      Exit Sub
   End If
   '
   NROPED& = CVS(Left$(Z$, 4))
   REGPED& = CVS(Mid$(Z$, 53, 4))
   '
   If COMALTER%("Confirma Anulación de Pedido Número '" + TRIM$(Str$(NROPED&)) + "'.\\Cancelar\Anular") <> 2 Then GoTo 30
   '
   For U& = REGPED& To FIN&
      X$ = REGLEIDO$("PEDCLIEN", U&)
      If CVS(Left$(X$, 4)) <> NROPED& Then GoTo 99
      Call REPLA(X$, Chr$(9), 300, 1)
      Call GRAREG("PEDCLIEN", X$, U&)
   Next U&
   '
99 Call CIERRARCH("PEDCLIEN")
   Call FRESHECHO("!INDIPEP")
   GoTo 10
   '
End Sub

Sub LIPEDIT(MODO%)
   '
   Screen.MousePointer = 11
   JJ& = 0
   FIN& = ULTREG&("PEDCLIEN")
   '
   For U& = 1 To FIN&
      Call TRACE(20, U&, FIN&)
      X$ = REGLEIDO$("PEDCLIEN", U&)
      '
      STATUS% = Asc(Mid$(X$, 300, 1))
      If STATUS% <= 1 Then
         '
         CACOPED = CVS(Mid$(X$, 513, 4))
         CAITPED = CVS(Mid$(X$, 645, 4))
         CACOENT = CVS(Mid$(X$, 512 + 77, 4))
         CANPEN = (CACOPED - CACOENT) * CAITPED
         '
         If CANPEN > 0.05 Then
            ENCAPED$ = Left$(X$, 512)
            COPED$ = Mid$(X$, 513, 128)
            ITPED$ = Mid$(X$, 641, 64)
            '
            NROPED& = CVS(Mid$(X$, 1, 4))
            FF% = CVI(Mid$(X$, 5, 2))
            NC% = CVI(Mid$(X$, 7, 2))
            LPRE% = Asc(Mid$(ENCAPED$, 297, 1))
            CPAG% = Asc(Mid$(ENCAPED$, 298, 1))
            VEND% = Asc(Mid$(ENCAPED$, 299, 1))
            NOCOM$ = Mid$(ENCAPED$, 137, 18)
            '
            REFE$ = Mid$(COPED$, 5, 58)
            FENTRESOL% = CVI(Mid$(COPED$, 63, 2))
            '
            CI% = CVI(Mid$(ITPED$, 3, 2))
            PRELI = CVS(Mid$(ITPED$, 9, 4))
            PREUNI = CVS(Mid$(ITPED$, 17, 4))
            '
            Y$ = REGBLAN$("LIPEDPE")
            Call REPLA(Y$, MKI$(NC%), 1, 2)
            Call REPLA(Y$, MKI$(CI%), 3, 2)
            Call REPLA(Y$, MKS$(PRELI), 41, 4)
            Call REPLA(Y$, MKI$(NC%), 45, 2)
            Call REPLA(Y$, MKS$(NROPED&), 47, 4)
            Call REPLA(Y$, MKI$(FF%), 51, 2)
            Call REPLA(Y$, NOCOM$, 53, 16)
            Call REPLA(Y$, MKI$(FENTRESOL%), 69, 2)
            Call REPLA(Y$, Chr$(STATUS%), 71, 1)
            Call REPLA(Y$, Chr$(LPRE%), 72, 1)
            Call REPLA(Y$, Chr$(CPAG%), 73, 1)
            Call REPLA(Y$, Chr$(VEND%), 74, 1)
            Call REPLA(Y$, UNIMED$(CI%), 77, 4)
            Call REPLA(Y$, MKS$(CACOPED * CAITPED), 81, 4)
            Call REPLA(Y$, MKS$(PREUNI), 85, 4)
            Call REPLA(Y$, MKD$(CDbl(CACOPED * CAITPED * PREUNI)), 89, 8)
            Call REPLA(Y$, MKS$(CACOENT * CAITPED), 97, 4)
            Call REPLA(Y$, MKS$(CANPEN), 101, 4)
            Call REPLA(Y$, MKS$(CANPEN), 111, 4)
            Call REPLA(Y$, MKD$(CDbl(CANPEN * PREUNI)), 119, 8)
            Call REPLA(Y$, Space$(34), 127, 34)
            '
            JJ& = JJ& + 1
            Call GRAREG("LIPEDPE", Y$, JJ&)
         End If
      End If
   Next U&
   '
   Call SETULTREG("LIPEDPE", JJ&)
   Call CIERRARCH("LIPEDPE")
   Screen.MousePointer = 1
   '
   If MODO% = 2 Then
         Call FINAL("*SOPEDNU")
      ElseIf MODO% = 3 Then
         Call FINAL("*SOPENCL")
      ElseIf MODO% = 4 Then
         Call FINAL("*SOPENMA")
   End If
   '
End Sub
