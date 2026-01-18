Attribute VB_Name = "FACTUNUE"
Public PIVAFAC%                     ' posicion de IVA de factura
Dim LETRAFAC$                    ' letra de la factura
Dim FACIVA                       ' coeficiente precios por IVA
Dim VARIFAC%                     ' variante de facturación
Dim ALITOTA                      ' alicuota total de IVA
Public RENOIN%                      ' marca responsable no inscripto
'
Public FEFAC%                       ' fecha de factura
Dim TIDOC%                       ' tipo de documento
'
Dim ENCAPED$
Public ARCHIPRE$
'
Dim NUPEDSEL&(2048)
Dim RPEDX$
Dim CAPEDPE%                     ' cantidad de pedidos pendientes
Dim CAPESEL%                     ' cantidad de pedidos seleccionados
Dim REFPED$
'
Public NUMECORR%
Public MONUMERA$
Dim MODOCARCON%
'
Public MESABONO$, DESABONO%, HASABONO%
Dim PORDESCUENTO
'

Sub FACITEMI1(OKEY%)
   '
   OKEY% = 0
   '
   NC% = Val(FORFACTURA.Text1.TEXT)
   If NC% < 1 Then
      Call MENSERR(24, "Número de Cuenta no Válido")
      Exit Sub
   End If
   '
   If ECOARCH$("DEUDOR1", MKI$(NC%)) = "" Then
      Call MENSERR(24, "Cuenta Inexistente")
      Exit Sub
   End If
   '
   FORFACTURA.Text1.TEXT = Str$(NC%)
   FORFACTURA.Text2.TEXT = RASOCLI$(NC%)
   FORFACTURA.Text3.TEXT = DOMICLI$(NC%)
   FORFACTURA.Text4.TEXT = CPOSCLI$(NC%)
   FORFACTURA.Text5.TEXT = LOCACLI$(NC%)
   FORFACTURA.Text7.TEXT = TRIM$(CUITCLI$(NC%)) + " - " + POSIVAC$(NC%)
   '
   PIVAFAC% = PIVACLI%(NC%)
   If PIVAFAC% <> 0 Then
      If PIVAFAC% <> 5 Then
         If TRIM$(CUITCLI$(NC%)) = "" Or VALICUIT%(CUITCLI$(NC%)) <> 1 Then
            Call MENSERR(24, "Imposible Facturar a Esta Cuenta.\Número de C.U.I.T. no Válido.")
            Exit Sub
         End If
      End If
   End If
   '
   If LIMICRE(NC%) < 10 Then
      Call MENSERR(24, "Imposible Facturar a Esta Cuenta.\Falta Asignar Límite de Crédito.")
      Exit Sub
   End If
   '
   If LIMICRE(NC%) < SALDCLI#(NC%) + CHEPEACRE#(NC%) Then
      Call MENSERR(24, "Imposible Facturar a Esta Cuenta.\Ha Excedido Límite de Crédito.")
      Exit Sub
   End If
   '
   Call CARLETFAC
   FORFACTURA.Text6.TEXT = LETRAFAC$
   TIDOC% = 2                 ' tipo de documento = FACTURA
   '
   NROPRO& = 0
   If Left$(FORFACTURA.Text11.TEXT, 3) = "FC." Then NROPRO& = Val(Mid$(FORFACTURA.Text11.TEXT, 4))
   Call NUFACTURA(NROPRO&)
   FORFACTURA.Text11.TEXT = "FC." + TRIM$(Str$(NROPRO&))
   '
   Call SELPEDPEN             ' seleccion pedidos pendientes
   'If CAPESEL% < 1 Then Exit Sub
   '
   PEDX$ = String$(512, 0)
   Call REPLA(PEDX$, MKS$(NROPRO&), 1, 4)
   FECHDOC$ = FORFACTURA.Text9.TEXT
   FEFAC% = FECHANUM(FECHDOC$)
   If FEFAC% < 8600 Then FEFAC% = HOY(HO$)
   Call REPLA(PEDX$, MKI$(FEFAC%), 5, 2)
   Call REPLA(PEDX$, MKI$(NC%), 7, 2)
   Call REPLA(PEDX$, RASOCLI$(NC%), 9, 32)
   Call REPLA(PEDX$, DOMICLI$(NC%), 41, 32)
   CPOSDOMI$ = ""
   If CPOSCLI$(NC%) <> "" Then CPOSDOMI$ = CPOSCLI$(NC%) + " -"
   CPOSDOMI$ = CPOSDOMI$ + DOMICLI$(NC%)
   Call REPLA(PEDX$, CPOSDOMI$, 73, 32)
   Call REPLA(PEDX$, CUITCLI$(NC%), 105, 32)
   Call REPLA(PEDX$, Space$(376), 137, 376)
   Call REPLA(PEDX$, Chr$(LIPRCLI%(NC%)), 297, 1)
   Call REPLA(PEDX$, Chr$(CPAGCLI%(NC%)), 298, 1)
   Call REPLA(PEDX$, Chr$(VENDCLI%(NC%)), 299, 1)
   '
10 ENCAPED$ = OBJPLA$("FACTURA-ENCDIP", PEDX$)
   If ENCAPED$ = "" Then
      OKEY% = 0
      Exit Sub
   End If
   '
   PEDX$ = ENCAPED$
   '
   LPRE% = Asc(Mid$(ENCAPED$, 297, 1))
   If LPRE% < 1 Or ECOARCH$("LISTAS", Chr$(LPRE%)) = "" Then
      Call MENSERR(24, "Lista de Precios no Válida")
      GoTo 10
   End If
   '
   CPAG% = Asc(Mid$(ENCAPED$, 298, 1))
   If CPAG% < 1 Or ECOARCH$("CONPAG", Chr$(CPAG%)) = "" Then
      Call MENSERR(24, "Condición de Pago no Válida")
      GoTo 10
   End If
   CONPA$ = DECONPA$(CPAG%)
   '
   VEND% = Asc(Mid$(ENCAPED$, 299, 1))
   If VEND% > 0 Then
      If ECOARCH$("VENDEDOR", Chr$(VEND%)) = "" Then
         Call MENSERR(24, "Vendedor no Figura en Tablas")
         GoTo 10
      End If
   End If
   VENDE$ = DEVENDE$(VEND%)
   '
   ARCHIPRE$ = TRIM$(Str$(LPRE%))
   While Len(ARCHIPRE$) < 3: ARCHIPRE$ = "0" + ARCHIPRE$: Wend
   ARCHIPRE$ = "LIPRE" + ARCHIPRE$
   '
   NOCOM$ = Mid$(ENCAPED$, 137, 18)
   '
   MODOCARCON% = 1
   Call CARPEDSEL             ' carga de pedidos seleccionados
   MODOCARCON% = 0
   '
   OKEY% = 1
   '
End Sub
'

Sub CARLETFAC()
    '
81  FACIVA = 1
    LETRAFAC$ = "A"
    VARIFAC% = 1
    ALITOTA = ALIVA
    RENOIN% = 0
    '
    If PIVAFAC% = 0 Or PIVAFAC% = 3 Or PIVAFAC% = 4 Or PIVAFAC% = 6 Then
          FACIVA = 1 + ALIVA / 100
          LETRAFAC$ = "B"
          VARIFAC% = 2
        ElseIf PIVAFAC% = 2 Then
          ALITOTA = ALIVA + ALIVARNI
          RENOIN% = 1
        ElseIf PIVAFAC% = 5 Then
          RENOIN% = 0
          ALITOTA = 0
          LETRAFAC$ = "E"
          VARIFAC% = 3
          If LETREXPO$ = "A" Then
              VARIFAC% = 1
              LETRAFAC$ = "A"
          End If
    End If
    '
End Sub
'

Sub NUFACTURA(NROPRO&)
   '
   Static FLINUMERA%
   If FLINUMERA% = 0 Then
      FLINUMERA% = FECHANUM("01/01/99")
      If ULTREG&("FORMFAC") > 0 Then
         FFS$ = REGLEIDO$("FORMFAC", 2&)
         FLINUMERA% = CVI(Mid$(FFS$, 19, 2))
      End If
   End If
   '
   If FEFAC% < FLINUMERA% Then
      Call MENSERR(24, "Imposible Facturar en la Fecha Elegida.\  \La Fecha de Factura Debe Ser POSTERIOR\a la Fecha de Nueva Numeración.")
      Call CIERRARCH("*.*")
      Call FINAL("")
   End If
   '
   If NUMECORR% = 0 Then
      NUMECORR% = 1
      If CONTROL$("NUMEFACT", "NUMECORR") = "INDEPEN" Then
         NUMECORR% = (-1)
      End If
   End If
   '
   If MONUMERA$ = "" Then
      MONUMERA$ = "AUTO"
      If TRIM$(CONTROL$("", "NUMFACTU")) = "MANUAL" Then
         MONUMERA$ = "MANUAL"
      End If
   End If
   '
   If MONUMERA$ = "MANUAL" Then
     If NROPRO& >= 1 Then
       '
       For MXC& = ULTREG&("FACTUR") To 1 Step -1
         F$ = REGLEIDO$("FACTUR", MXC&)
         TIX% = Asc(Mid$(F$, 13, 1)): VAX% = Asc(Mid$(F$, 14, 1))
         If TIX% = TIDOC% Or NUMECORR% = 1 Then
           If VAX% = VARIFAC% Then
             If CVI(Mid$(F$, 7, 2)) >= FLINUMERA% Then
               If CVS(Mid$(F$, 3, 4)) = NROPRO& Then
                 Call MENSERR(24, "Número " + TRIM$(Str$(NROPRO&)) + " ya Utilizado")
                 GoTo 20
               End If
             End If
           End If
         End If
       Next MXC&
       Exit Sub
       '
     End If
   End If
   '
20 NUMAXI& = 0: KKK% = 0
   For MXC& = ULTREG&("FACTUR") To 1 Step -1
       F$ = REGLEIDO$("FACTUR", MXC&)
       TIX% = Asc(Mid$(F$, 13, 1)): VAX% = Asc(Mid$(F$, 14, 1))
       If TIX% = TIDOC% Or NUMECORR% = 1 Then
          If VAX% = VARIFAC% Then
             If CVI(Mid$(F$, 7, 2)) >= FLINUMERA% Then
                If CVS(Mid$(F$, 3, 4)) > NUMAXI& Then NUMAXI& = CVS(Mid$(F$, 3, 4))
                NROPRO& = 1 + NUMAXI&
                KKK% = KKK% + 1
                If KKK% > 128 Then Exit Sub
             End If
          End If
       End If
   Next MXC&
   If NROPRO& > 0 Then Exit Sub
   '
10 NROPRO& = Val(InputBox$("Numero Primer Comprobante"))
   If NROPRO& < 1 Then GoTo 10
   '
End Sub
'

Sub SELPEDPEN()
   '
10 RPEDX$ = RECFILT$("REMPEN", 1, MKI$(NC%))
   NROPEDA& = 0
   CAPADPE% = 0
   FORSELPED.LIPEDPE.Clear
   FIN& = Len(RPEDX$)
   '
   For UU& = 1 To Len(RPEDX$) Step 4
      '
      If FIN& > 0 Then
         Call TRACE(20, UU&, FIN&)
      End If
      '
      U& = CVS(Mid$(RPEDX$, UU&, 4))
      X$ = REGLEIDO$("REMPEN", U&)
      If CVI(Mid$(X$, 1, 2)) <> NC% Then
         X$ = REGLEIDO$("REMPEN", 1)
         Call GRAREG("REMPEN", X$, 1)
         Call CIERRARCH("REMPEN")
         GoTo 10
      End If
      '
      CACOPED = 1
      CAN = CVS(Mid$(X$, 27, 4))
      '
      If CAN > 0.05 Then
         NROPED = CVS(Mid$(X$, 3, 4))
         If NROPED <> NROPEDA Then
            TTXX$ = "RT-" + TRIM$(CSTRING$(MKS$(NROPED), 3, 8, 0, 2))
            TTXX$ = AJUSTI$(TTXX$, 11)
            FEX = CVI(Mid$(X$, 5, 2))
            TTXX$ = TTXX$ + "  " + FECHATEX$(FEX)
            FORSELPED.LIPEDPE.AddItem TTXX$
         End If
         NROPEDA = NROPED
         CAPEDPE% = CAPEDPE% + 1
         FENTRE% = CVI(Mid$(X$, 7, 2))
         If FENTRE% < FINF Then FINF = FENTRE%
         If FENTRE% > FSUP Then FSUP = FENTRE%
      End If
19 Next UU&
   '
20 Screen.MousePointer = 1
   '
   If CAPEDPE% < 1 Then
      Call COMUNI("No hay Remitos Pendientes\para este Cliente")
      GoTo 99
   End If
   '
   FORSELPED.Show 1
   '
   CAPESEL% = 0
   For U& = 1 To FORSELPED.LIPEDPE.ListCount
      If FORSELPED.LIPEDPE.Selected(U& - 1) = True Then
         FORSELPED.LIPEDPE.ListIndex = U& - 1
         CAPESEL% = CAPESEL% + 1
         TTXX$ = FORSELPED.LIPEDPE.TEXT
         NUPEDSEL&(CAPESEL%) = Val(Mid$(TTXX$, 4, 8))
      End If
   Next U&
   '
99 End Sub
'

Sub CARPEDSEL()
   '
10 FIN& = Len(RPEDX$)
   Call BLANARCH("!CARCONPE")
   Call BLANARCH("!CARITPED")
   REFANT$ = "":   REFPED$ = ""
   NROPEDA& = 0: ITIDX% = 0: ITIDA% = 0
   '
   For UU& = 1 To Len(RPEDX$) Step 4
      '
      If FIN& > 0 Then
         Call TRACE(20, UU&, FIN&)
      End If
      '
      U& = CVS(Mid$(RPEDX$, UU&, 4))
      X$ = REGLEIDO$("REMPEN", U&)
      If CVI(Mid$(X$, 1, 2)) <> NC% Then
         Call MENSERR(24, "Error de Indices - Consulte")
         Call FINAL("")
      End If
      '
      NROPED& = CVS(Mid$(X$, 3, 4))
      For JI% = 1 To CAPESEL%
         If NUPEDSEL&(JI%) = NROPED& Then GoTo 20
      Next JI%
      GoTo 29
      '
20    NROPEDA& = NROPED&
      CACOPED = 1
      CIIT% = CVI(Mid$(X$, 9, 2))
      PREBA = PREBASE(CIIT%)
      PRELI = PRELISTA#(ARCHIPRE$, CIIT%)
      PRECO = PRELI
      CAIT = CVS(Mid$(X$, 27, 4))
      IMPO# = CDbl(CAIT * PRECO)
      '
      If CAIT > 0.05 Then
         ITID% = 1
         If REFE$ <> REFANT$ Or ITID% <> ITIDA% Then
            ITIDX% = ITIDX% + 1
            Call REPLA(COPED$, MKS$(1), 1, 4)
            Call REPLA(COPED$, MKI$(ITIDX%), 127, 2)
            Call REGAPP("!CARCONPE", COPED$)
         End If
         REFANT$ = REFE$
         ITIDA% = ITID%
         '
         For KKI& = 1 To ULTREG&("!CARITPED")
           OBX$ = REGLEIDO$("!CARITPED", KKI&)
           If CVI(Mid$(OBX$, 3, 2)) = CIIT% Then
             If CVS(Mid$(OBX$, 33, 4)) < 0.5 Then
               TOIT = CVS(Mid$(OBX$, 5, 4)) + CAIT
               Call REPLA(OBX$, MKS$(TOIT), 5, 4)
               Call REPLA(OBX$, MKS$(TOIT), 29, 4)
               IMTOT# = CVD(Mid$(OBX$, 21, 8)) + IMPO#
               Call REPLA(OBX$, MKD$(IMTOT#), 21, 8)
               VEREMI$ = Mid$(OBX$, 37, 28) + MKS$(U&)
               Call REPLA(OBX$, VEREMI$, 33, 32)
               Call GRAREG("!CARITPED", OBX$, KKI&)
               GoTo 29
             End If
           End If
         Next KKI&
         '
         OBX$ = REGBLAN$("!CARITPED")
         Call REPLA(OBX$, MKI$(1), 1, 2)
         Call REPLA(OBX$, MKI$(CIIT%), 3, 2)
         Call REPLA(OBX$, MKS$(CAIT), 5, 4)
         Call REPLA(OBX$, MKS$(PRELI), 9, 4)
         Call REPLA(OBX$, MKS$(PREBA), 13, 4)
         Call REPLA(OBX$, MKS$(PRECO), 17, 4)
         Call REPLA(OBX$, MKD$(IMPO#), 21, 8)
         Call REPLA(OBX$, MKS$(CAIT), 29, 4)
         Call REPLA(OBX$, String$(28, 0), 33, 28)
         Call REPLA(OBX$, MKS$(U&), 61, 4)
22       If IMPO# < 0.05 Then
            VDEF$ = OBX$ + "/ARCHECO(3)"
            CAANT = CAIT: PREANT = PRECO
25          Y$ = OBJPLA$("CARITPED2", VDEF$)
            If Y$ = "" Then Exit Sub
            If CVI(Mid$(Y$, 3, 2)) <> CIIT% Then
               Call MENSERR(24, "Cambio de Artículo no Válido\en esta Pantalla.")
               GoTo 10
            End If
            '
            '
            VDEF$ = Y$
            CIIT% = CVI(Mid$(Y$, 3, 2))
            CAIT = CVS(Mid$(Y$, 5, 4))
            PRELI = CVS(Mid$(Y$, 9, 4))
            PREBA = PREBASE(CIIT%)
            PRECO = CVS(Mid$(Y$, 17, 4))
            IMPO# = CDbl(CAIT * PRECO)
            GoTo 22
         End If
         '
         ITID% = 1
         CACO = 1
         '
         Call REPLA(OBX$, MKI$(ITID%), 1, 2)
         Call REPLA(OBX$, MKD$(IMPO#), 21, 8)
         Call REGAPP("!CARITPED", OBX$)
         '
      End If
29 Next UU&
   '
   For UU& = 1 To ULTREG&("!CARITPED")
      '
      Z$ = Space$(96)
      OBX$ = REGLEIDO$("!CARITPED", UU&)
      CAIT = CVS(Mid$(OBX$, 5, 4))
      CIIT% = CVI(Mid$(OBX$, 3, 2))
      PRELI = CVS(Mid$(OBX$, 9, 4))
      PRECO = CVS(Mid$(OBX$, 17, 4))
      IMPO# = CVD(Mid$(OBX$, 21, 8))
      U& = CVS(Mid$(OBX$, 61, 4))
      '
      Call REPLA(Z$, CSTRING$(MKS$(CAIT), 3, 4, 0, 2), 1, 4)
      Call REPLA(Z$, CODEXT$(CIIT%), 7, 15)
      Call REPLA(Z$, DESCRIT0$(CIIT%), 23, 33)
      Call REPLA(Z$, CSTRING$(MKS$(PRELI), 3, 9, 2, 2), 56, 9)
      Call REPLA(Z$, CSTRING$(MKS$(PRECO), 3, 9, 2, 2), 65, 9)
      Call REPLA(Z$, CSTRING$(MKD$(IMPO#), 3, 10, 2, 2), 74, 10)
      Call REPLA(Z$, CSTRING$(MKS$(U&), 3, 6, 0, 2), 91, 6)
      FORFACDIP.List2.AddItem Z$
      FORFACDIP.List2.ListIndex = FORFACDIP.List2.ListCount - 1
      '
39 Next UU&
   '
   ITOTCO# = 0
   For U& = 1 To ULTREG&("!CARITPED")
      Y$ = REGLEIDO$("!CARITPED", U&)
      If CVI(Left$(Y$, 2)) = ITID% Then
         CAIT = CVS(Mid$(Y$, 5, 4))
         PRECO = CVS(Mid$(Y$, 17, 4))
         IMPO# = CAIT * PRECO
         Call REPLA(Y$, MKD$(IMPO#), 21, 8)
         ITOTCO# = ITOTCO# + IMPO#
      End If
   Next U&
   If CACO < 1 Then CACO = 1
   PREUCO = ITOTCO# / CACO
   FORFACTURA.Label10.Caption = CSTRING$(MKD$(ITOTCO#), 3, 12, 2, 2)
   '
   ILISTA& = 0
   X$ = REGLEIDO$("!CARCONPE", ILISTA& + 1)
   Call REPLA(X$, MKS$(PREUCO), 65, 4)
   Call REPLA(X$, MKD$(ITOTCO#), 69, 8)
   Call GRAREG("!CARCONPE", X$, ILISTA& + 1)
   '
   TOPEDIDO# = 0
   For U& = 1 To ULTREG&("!CARCONPE")
       X$ = REGLEIDO$("!CARCONPE", U&)
       TOPEDIDO# = TOPEDIDO# + CVD(Mid$(X$, 69, 8))
   Next U&
   '
   ALITOTA = 0
   If PIVAFAC% <> 5 Then ALITOTA = ALIVA
   FORFACTURA.Label14.Caption = CSTRING$(MKD$(TOPEDIDO#), 3, 11, 2, 2)
   If PIVAFAC% <> 5 Then FORFACTURA.Label18.Caption = CSTRING$(MKD$(TOPEDIDO# * ALIVA / 100), 3, 11, 2, 2)
   If RENOIN% = 1 Then
      FORFACTURA.Label21.Caption = CSTRING$(MKD$(TOPEDIDO# * ALIVARNI / 100), 3, 11, 2, 2)
      ALITOTA = ALIVA + ALIVARNI
   End If
   FORFACTURA.Label20.Caption = CSTRING$(MKD$(TOPEDIDO# * (1 + ALITOTA / 100)), 3, 11, 2, 2)
   '
End Sub
'

Sub CARCONFAC()
    '
200 If MODOCARCON% = 1 Then GoTo 210
    '
    ATRI$ = "/NC"
    ATRI$ = ATRI$ + "/TITUPAN=Ingreso Conjuntos a Itemizar"
    ATRI$ = ATRI$ + "/BORDESUP=W" + TRIM$(Str$(FORFACDIP.Top + 100))
    ATRI$ = ATRI$ + "/BORDEIZQ=W" + TRIM$(Str$(FORFACDIP.Left))
    '
    ITID% = 1: J& = 1
    '
          X$ = REGBLAN$("CARCONPE")
          Call REPLA(X$, MKS$(1), 1, 4)
          Call REPLA(X$, RECON$, 5, 58)
          Call REPLA(X$, MKI$(FFI%), 63, 2)
          Call REPLA(X$, MKS$(PREUCO), 65, 4)
          Call REPLA(X$, MKD$(ITOTCO#), 69, 8)
          Call REPLA(X$, MKI$(ITID%), 127, 2)
          Call GRAREG("!CARCONPE", X$, J&)
       Call SETULTREG("!CARCONPE", 1)
       Exit Sub
       '
    '
210 REBLA$ = REGBLAN$("CARCONPE")
    'FORFACDIP.List1.Clear
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
        If ITID% < 1 Then
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
        'FORFACDIP.List1.AddItem Z$
        TOPEDIDO# = TOPEDIDO# + ITOTCO#
        '
290 Next U&
    '
    ALITOTA = 0
    If VARIFAC% <> 3 Then ALITOTA = ALIVA
    FORFACDIP.Label14.Caption = CSTRING$(MKD$(TOPEDIDO#), 3, 11, 2, 2)
    If PIVAFAC% <> 5 Then FORFACDIP.Label18.Caption = CSTRING$(MKD$(TOPEDIDO# * ALIVA / 100), 3, 11, 2, 2)
    If RENOIN% = 1 Then
       FORFACDIP.Label21.Caption = CSTRING$(MKD$(TOPEDIDO# * ALIVARNI / 100), 3, 11, 2, 2)
       ALITOTA = ALIVA + ALIVARNI
    End If
    FORFACDIP.Label20.Caption = CSTRING$(MKD$(TOPEDIDO# * (1 + ALITOTA / 100)), 3, 11, 2, 2)
    '
End Sub
'

Sub GRAFACDIP(OKEY%)
   '
   HOII% = HOY(HO$)
   '
   OKEY% = 0
   If ULTREG&("!CARITPED") < 1 Then
       Call MENSERR(24, "Factura Vacía")
       Exit Sub
   End If
   '
   FECHDOC$ = FORFACDIP.Text9.TEXT
   FEFAC% = FECHANUM(FECHDOC$)
   If FEFAC% < 8600 Then
      Call MENSERR(24, "Fecha de Factura no Válida")
      Exit Sub
   End If
   FEX = FEFAC%
   FORFACDIP.Text9.TEXT = FECHATEX$(FEX)
   '
   CACO = 1
   ITID% = 1
   For T& = 1 To ULTREG&("!CARITPED")
     Y$ = REGLEIDO$("!CARITPED", T&)
     If CVI(Left$(Y$, 2)) = ITID% Then
       CAN = CACO * CVS(Mid$(Y$, 5, 4))
       If CAN > 0.05 Then
         CIIT% = CVI(Mid$(Y$, 3, 2))
         If CODEXT$(CIIT%) <> "" Then
           PREU = CVS(Mid$(Y$, 17, 4))
           If PREU >= 0.005 Then
             GoTo 14
           End If
         End If
       End If
     End If
14 Next T&
   '
   FORIPRE$ = CONTROL$("", "FORFAC-" + LETRAFAC$)
   MAXITS$ = ATRIFORM$(FORIPRE$, "MAX-ITEM")
   If MAXITS$ = "HOJACONT" Then GoTo 20
   MAXIT% = Val(MAXITS$)
   
   CAITAB1% = 0
   ITID% = 1
   CACO = 1
   VECREMI$ = ""
   For T& = 1 To ULTREG&("!CARITPED")
     Y$ = REGLEIDO$("!CARITPED", T&)
     If CVI(Left$(Y$, 2)) = ITID% Then
       CAITAB1% = CAITAB1% + 1
       For KKI% = 33 To 64 Step 4
         REIT& = CVS(Mid$(Y$, KKI%, 4))
         If REIT& > 0 Then
           If REIT& <= ULTREG&("REMPEN") Then
             Z$ = REGLEIDO$("REMPEN", REIT&)
             NUREMI& = CVS(Mid$(Z$, 3, 4))
             For KKJ% = 1 To Len(VECREMI$) Step 4
               If CVS(Mid$(VECREMI$, KKJ%, 4)) = NUREMI& Then GoTo 19
             Next KKJ%
             VECREMI$ = VECREMI$ + MKS$(NUREMI&)
           End If
         End If
19     Next KKI%
     End If
   Next T&
   '
   REMITOS$ = ""
   For KKJ% = 1 To Len(VECREMI$) Step 4
     If REMITOS$ <> "" Then REMITOS$ = REMITOS$ + ","
     REMITOS$ = REMITOS$ + TRIM$(Str$(CVS(Mid$(VECREMI$, KKJ%, 4))))
   Next KKJ%
   FORFACDIP.Label2.Caption = REMITOS$
   '
   If CAITAB1% > MAXIT% Then
      Call MENSERR(24, "Imposible Registrar esta Factura.\Excede Máximo de Renglones Admisible\en el Formulario de Impresión.")
      Exit Sub
   End If
   '
20 TODESACO# = 0
   PORDESCUENTO = 0
   NEFA# = Val(FORFACDIP.Label14.Caption)
   OBX$ = OBJPLA$("CARDESCUENTO", MKD$(NEFA#) + MKS$(0))
   If OBX$ = "" Then Exit Sub
   '
   PORDESCUENTO = CVS(Mid$(OBX$, 9, 4))
   If PORDESCUENTO < 0 Or PORDESCUENTO > 90 Then GoTo 20
   NEFA# = Val(FORFACDIP.Label14.Caption) * (1 - PORDESCUENTO / 100)
   '
30 Call BLOQARCH("FACTUR")
   Call BLOQARCH("SVD202")
   Call BLOQARCH("DEUDOR2")
   Call BLOQARCH("MOVISTO")
   Call BLOQARCH("BKMOVIST")
   Call BLOQARCH("REMPEN")
   Call BLOQARCH("FACOPIOS")
   '
   NROPRO& = 0
   If Left$(FORFACDIP.Text11.TEXT, 3) = "FC." Then NROPRO& = Val(Mid$(FORFACDIP.Text11.TEXT, 4))
   Call NUFACTURA(NROPRO&)
   FORFACDIP.Text11.TEXT = "FC." + TRIM$(Str$(NROPRO&))
   '
   X$ = REGBLAN$("FACTUR")
   Call REPLA(X$, MKI$(NC%), 1, 2)
   Call REPLA(X$, MKS$(NROPRO&), 3, 4)
   Call REPLA(X$, MKI$(FEFAC%), 7, 2)
   Call REPLA(X$, MKI$(FVI%), 9, 2)
   Call REPLA(X$, MKI$(LPRE%), 11, 2)
   Call REPLA(X$, Chr$(TIDOC%), 13, 1)
   Call REPLA(X$, Chr$(VARIFAC%), 14, 1)
   '
   NEFA# = Val(FORFACDIP.Label14.Caption) * (1 - PORDESCUENTO / 100)
   NIVA# = Val(FORFACDIP.Label18.Caption) * (1 - PORDESCUENTO / 100)
   NNIN# = Val(FORFACDIP.Label21.Caption) * (1 - PORDESCUENTO / 100)
   TOTA# = Val(FORFACDIP.Label20.Caption) * (1 - PORDESCUENTO / 100)
   IINT# = 0
   PERC# = 0
   '
   TOTA1# = 0
   If TOTA# - TOTA1# + SALDCLI#(NC%) + CHEPEACRE#(NC%) > LIMICRE(NC%) Then
      Call MENSERR(24, "Imposible Registrar esta Factura.\  \Supera el Límite de Crédito Asignado\para este Cliente.")
      Call CIERRARCH("FACTUR")
      Call CIERRARCH("SVD202")
      Call CIERRARCH("DEUDOR2")
      Call CIERRARCH("MOVISTO")
      Call CIERRARCH("BKMOVIST")
      Call CIERRARCH("REMPEN")
      Exit Sub
   End If
   Call REPLA(X$, MKD$(NEFA#), 15, 8)
   Call REPLA(X$, MKD$(NIVA#), 23, 8)
   Call REPLA(X$, MKD$(NNIN#), 31, 8)
   Call REPLA(X$, MKD$(TOTA#), 39, 8)
   Call REPLA(X$, MKI$(VENDE%), 47, 8)
   Call REPLA(X$, MKD$(IINT#), 49, 8)
   Call REPLA(X$, MKD$(PERC#), 57, 8)
   '
   Call REGAPP("FACTUR", X$)
   '
   TOTA1# = 0
   '
   Call CIERRARCH("FACTUR")
   '
   CPAG% = Asc(Mid$(ENCAPED$, 298, 1))
   RSV1& = ULTREG&("SVD202")
   Call GRADEU(LUDAT$, "SVD202", RDX%, TIDOC%, VARIFAC%, NROPRO&, NC%, FEFAC%, CPAG%, TOTA# - TOTA1#)
   RSV2& = ULTREG&("SVD202")
   For U& = RSV1& + 1 To RSV2&
      X$ = REGLEIDO$("SVD202", U&)
      Call REPLA(X$, MKD$(TOTA#), 21, 8)
      Call GRAREG("SVD202", X$, U&)
   Next U&
   '
   Call CIERRARCH("SVD202")
   '
   RGX& = REGICLI&(NC%)
   If RGX& > 0 Then
      If RGX& <= ULTREG&("DEUDOR2") Then
         DX$ = REGLEIDO$("DEUDOR2", RGX&)
         SALDO# = CVD(Mid$(DX$, 109, 8)) + TOTA# - TOTA1#
         Call REPLA(DX$, MKD$(SALDO#), 109, 8)
         Call GRAREG("DEUDOR2", DX$, RGX&)
      End If
   End If
   Call CIERRARCH("DEUDOR2")
   '
   CACO = 1
   ITID% = 1
   '
   For T& = 1 To ULTREG&("!CARITPED")
      Y$ = REGLEIDO$("!CARITPED", T&)
      If CVI(Left$(Y$, 2)) = ITID% Then
         CAN = CACO * CVS(Mid$(Y$, 5, 4))
         If CAN > 0.05 Then
            CIIT% = CVI(Mid$(Y$, 3, 2))
            PREU = CVS(Mid$(Y$, 17, 4)) * (1 - PORDESCUENTO / 100)
            ITOT# = CVD(Mid$(Y$, 21, 8)) * (1 - PORDESCUENTO / 100)
            REIT& = CVS(Mid$(Y$, 61, 4))
            If REIT& < 1 Or REIT& > ULTREG&("REMPEN") Then
               Call MOVISTO(FEFAC%, CIIT%, "", "FC", "RT." + TRIM$(Str$(NROPRO&)), "FC." + TRIM$(Str$(NROPRO&)), RASOCLI$(NC%), CDbl(PREU), "$", NC%, 1, (-1) * CAN)
            End If
         End If
      End If
   Next T&
   '
   Call CIERRARCH("MOVISTO")
   Call CIERRARCH("BKMOVIST")
   '
   CACO = 1
   ITID% = 1
   '
   For T& = 1 To ULTREG&("!CARITPED")
      Y$ = REGLEIDO$("!CARITPED", T&)
      If CVI(Left$(Y$, 2)) = ITID% Then
         If CACO > 0.05 Then
            CIIT% = CVI(Mid$(Y$, 3, 2))
            PREU = CVS(Mid$(Y$, 17, 4)) * (1 - PORDESCUENTO / 100)
            ITOT# = CVD(Mid$(Y$, 21, 8)) * (1 - PORDESCUENTO / 100)
            For KKI% = 33 To 64 Step 4
              REIT& = CVS(Mid$(Y$, KKI%, 4))
              If REIT& > 0 Then
                If REIT& <= ULTREG&("REMPEN") Then
                  Z$ = String$(96, 0)
                  Call GRAREG("REMPEN", Z$, REIT&)
                End If
              End If
            Next KKI%
         End If
      End If
   Next T&
   '
   Call CIERRARCH("REMPEN")
   '
   OKEY% = 1
   '
End Sub
'

Sub GRADEBILIB(OKEY%)
   '
   HOII% = HOY(HO$)
   OKEY% = 0
   '
   NC% = Val(FORDELIDIP.Text1.TEXT)
   If NC% < 1 Or ECOARCH$("DEUDOR1", MKI$(NC%)) = "" Then
       Call MENSERR(24, "Numero de Cuenta no Válido o Inexistente")
       FORDELIDIP.Text1.SetFocus
       Exit Sub
   End If
   '
   If FORDELIDIP.Text8.TEXT = "" Then
       Call MENSERR(24, "Falta Ingresar Detalle de Nota de Debito")
       FORDELIDIP.Text8.SetFocus
       Exit Sub
   End If
   '
   FECHDOC$ = FORDELIDIP.Text9.TEXT
   FEFAC% = FECHANUM(FECHDOC$)
   If FEFAC% < 8600 Then
      Call MENSERR(24, "Fecha de N.Débito no Válida")
      FORDELIDIP.Text9.SetFocus
      Exit Sub
   End If
   FEX = FEFAC%
   FORDELIDIP.Text9.TEXT = FECHATEX$(FEX)
   '
   CPAG% = Val(FORDELIDIP.TEXT21(1).TEXT)
   If ECOARCH$("CONPAG", MKI$(CPAG%)) = "" Then
      Call MENSERR(24, "Condición de Pago no Válida")
      FORDELIDIP.TEXT21(1).SetFocus
      Exit Sub
   End If
   CONPAG$ = DECONPA$(CPAG%)
   
   VEND% = Val(FORDELIDIP.TEXT21(1).TEXT)
   VENDE$ = DEVENDE$(VEND%)
   
   Call BLOQARCH("FACTUR")
   Call BLOQARCH("SVD202")
   Call BLOQARCH("DEUDOR2")
   '
   NROPRO& = 0
   If Left$(FORDELIDIP.Text11.TEXT, 3) = "ND." Then NROPRO& = Val(Mid$(FORDELIDIP.Text11.TEXT, 4))
   Call NUFACTURA(NROPRO&)
   FORDELIDIP.Text11.TEXT = "ND." + TRIM$(Str$(NROPRO&))
   '
   X$ = REGBLAN$("FACTUR")
   Call REPLA(X$, MKI$(NC%), 1, 2)
   Call REPLA(X$, MKS$(NROPRO&), 3, 4)
   Call REPLA(X$, MKI$(FEFAC%), 7, 2)
   Call REPLA(X$, MKI$(FVI%), 9, 2)
   Call REPLA(X$, MKI$(0), 11, 2)
   Call REPLA(X$, Chr$(TIDOC%), 13, 1)
   Call REPLA(X$, Chr$(VARIFAC%), 14, 1)
   '
   NEFA# = Val(FORDELIDIP.Text10.TEXT)
   NIVA# = Val(FORDELIDIP.Label18.Caption)
   NNIN# = Val(FORDELIDIP.Label23.Caption)
   TOTA# = Val(FORDELIDIP.Label20.Caption)
   IINT# = 0
   PERC# = 0
   '
   If NEFA# < 0.01 Then
      Call MENSERR(24, "Faltan Importes")
      FORDELIDIP.Text10.SetFocus
      Exit Sub
   End If
   '
   Call REPLA(X$, MKD$(NEFA#), 15, 8)
   Call REPLA(X$, MKD$(NIVA#), 23, 8)
   Call REPLA(X$, MKD$(NNIN#), 31, 8)
   Call REPLA(X$, MKD$(TOTA#), 39, 8)
   Call REPLA(X$, MKI$(VEND%), 47, 8)
   Call REPLA(X$, MKD$(IINT#), 49, 8)
   Call REPLA(X$, MKD$(PERC#), 57, 8)
   '
   Call REGAPP("FACTUR", X$)
   Call CIERRARCH("FACTUR")
   '
   Call GRADEU(LUDAT$, "SVD202", RDX%, TIDOC%, VARIFAC%, NROPRO&, NC%, FEFAC%, CPAG%, TOTA#)
   Call CIERRARCH("SVD202")
   '
   RGX& = REGICLI&(NC%)
   If RGX& > 0 Then
      If RGX& <= ULTREG&("DEUDOR2") Then
         DX$ = REGLEIDO$("DEUDOR2", RGX&)
         SALDO# = CVD(Mid$(DX$, 109, 8)) + TOTA#
         Call REPLA(DX$, MKD$(SALDO#), 109, 8)
         Call GRAREG("DEUDOR2", DX$, RGX&)
      End If
   End If
   Call CIERRARCH("DEUDOR2")
   '
   '
   OKEY% = 1
   '
End Sub
'

Sub GRACRELIB(OKEY%)
   '
   HOII% = HOY(HO$)
   OKEY% = 0
   '
   NC% = Val(FORCRELIDIP.Text1.TEXT)
   If NC% < 1 Or ECOARCH$("DEUDOR1", MKI$(NC%)) = "" Then
       Call MENSERR(24, "Numero de Cuenta no Válido o Inexistente")
       FORCRELIDIP.Text1.SetFocus
       Exit Sub
   End If
   '
   If FORCRELIDIP.Text8.TEXT = "" Then
       Call MENSERR(24, "Falta Ingresar Detalle de Nota de Crédito")
       FORCRELIDIP.Text8.SetFocus
       Exit Sub
   End If
   '
   FECHDOC$ = FORCRELIDIP.Text9.TEXT
   FEFAC% = FECHANUM(FECHDOC$)
   If FEFAC% < 8600 Then
      Call MENSERR(24, "Fecha de N.Crédito no Válida")
      FORCRELIDIP.Text9.SetFocus
      Exit Sub
   End If
   FEX = FEFAC%
   FORCRELIDIP.Text9.TEXT = FECHATEX$(FEX)
   '
   CPAG% = 0
   CONPAG$ = DECONPA$(CPAG%)
   
   VEND% = Val(FORCRELIDIP.TEXT21(1).TEXT)
   VENDE$ = DEVENDE$(VEND%)
   
   Call BLOQARCH("FACTUR")
   Call BLOQARCH("SVD202")
   Call BLOQARCH("DEUDOR2")
   '
   NROPRO& = 0
   If Left$(FORCRELIDIP.Text11.TEXT, 3) = "NC." Then NROPRO& = Val(Mid$(FORCRELIDIP.Text11.TEXT, 4))
   Call NUFACTURA(NROPRO&)
   FORCRELIDIP.Text11.TEXT = "NC." + TRIM$(Str$(NROPRO&))
   '
   X$ = REGBLAN$("FACTUR")
   Call REPLA(X$, MKI$(NC%), 1, 2)
   Call REPLA(X$, MKS$(NROPRO&), 3, 4)
   Call REPLA(X$, MKI$(FEFAC%), 7, 2)
   Call REPLA(X$, MKI$(FVI%), 9, 2)
   Call REPLA(X$, MKI$(0), 11, 2)
   Call REPLA(X$, Chr$(TIDOC%), 13, 1)
   Call REPLA(X$, Chr$(VARIFAC%), 14, 1)
   '
   NEFA# = Val(FORCRELIDIP.Text10.TEXT)
   NIVA# = Val(FORCRELIDIP.Label18.Caption)
   NNIN# = Val(FORCRELIDIP.Label23.Caption)
   TOTA# = Val(FORCRELIDIP.Label20.Caption)
   IINT# = 0
   PERC# = 0
   '
   If NEFA# < 0.01 Then
      Call MENSERR(24, "Faltan Importes")
      FORCRELIDIP.Text10.SetFocus
      Exit Sub
   End If
   '
   Call REPLA(X$, MKD$(NEFA#), 15, 8)
   Call REPLA(X$, MKD$(NIVA#), 23, 8)
   Call REPLA(X$, MKD$(NNIN#), 31, 8)
   Call REPLA(X$, MKD$(TOTA#), 39, 8)
   Call REPLA(X$, MKI$(VEND%), 47, 8)
   Call REPLA(X$, MKD$(IINT#), 49, 8)
   Call REPLA(X$, MKD$(PERC#), 57, 8)
   '
   Call REGAPP("FACTUR", X$)
   Call CIERRARCH("FACTUR")
   '
   Call GRADEU(LUDAT$, "SVD202", RDX%, TIDOC%, VARIFAC%, NROPRO&, NC%, FEFAC%, CPAG%, TOTA#)
   Call CIERRARCH("SVD202")
   '
   RGX& = REGICLI&(NC%)
   If RGX& > 0 Then
      If RGX& <= ULTREG&("DEUDOR2") Then
         DX$ = REGLEIDO$("DEUDOR2", RGX&)
         SALDO# = CVD(Mid$(DX$, 109, 8)) - TOTA#
         Call REPLA(DX$, MKD$(SALDO#), 109, 8)
         Call GRAREG("DEUDOR2", DX$, RGX&)
      End If
   End If
   Call CIERRARCH("DEUDOR2")
   '
   OKEY% = 1
   '
End Sub
'

Sub GRACREDEV(OKEY%)
   '
   HOII% = HOY(HO$)
   '
   OKEY% = 0
   If FORCREDEVDIP.List2.ListCount < 1 Then
       Call MENSERR(24, "Nota de Crédito Vacía")
       Exit Sub
   End If
   '
   FECHDOC$ = FORCREDEVDIP.Text9.TEXT
   FEFAC% = FECHANUM(FECHDOC$)
   If FEFAC% < 8600 Then
      Call MENSERR(24, "Fecha de N.Crédito no Válida")
      Exit Sub
   End If
   FEX = FEFAC%
   FORCREDEVDIP.Text9.TEXT = FECHATEX$(FEX)
   '
   Call BLOQARCH("FACTUR")
   Call BLOQARCH("SVD202")
   Call BLOQARCH("DEUDOR2")
   Call BLOQARCH("MOVISTO")
   Call BLOQARCH("BKMOVIST")
   '
   NROPRO& = 0
   If Left$(FORCREDEVDIP.Text11.TEXT, 3) = "NC." Then NROPRO& = Val(Mid$(FORCREDEVDIP.Text11.TEXT, 4))
   Call NUFACTURA(NROPRO&)
   FORCREDEVDIP.Text11.TEXT = "NC." + TRIM$(Str$(NROPRO&))
   '
   X$ = REGBLAN$("FACTUR")
   Call REPLA(X$, MKI$(NC%), 1, 2)
   Call REPLA(X$, MKS$(NROPRO&), 3, 4)
   Call REPLA(X$, MKI$(FEFAC%), 7, 2)
   Call REPLA(X$, MKI$(FVI%), 9, 2)
   Call REPLA(X$, MKI$(LPRE%), 11, 2)
   Call REPLA(X$, Chr$(TIDOC%), 13, 1)
   Call REPLA(X$, Chr$(VARIFAC%), 14, 1)
   '
   NEFA# = Val(FORCREDEVDIP.Label14.Caption)
   NIVA# = Val(FORCREDEVDIP.Label18.Caption)
   NNIN# = Val(FORCREDEVDIP.Label23.Caption)
   TOTA# = Val(FORCREDEVDIP.Label20.Caption)
   IINT# = 0
   PERC# = 0
   '
   Call REPLA(X$, MKD$(NEFA#), 15, 8)
   Call REPLA(X$, MKD$(NIVA#), 23, 8)
   Call REPLA(X$, MKD$(NNIN#), 31, 8)
   Call REPLA(X$, MKD$(TOTA#), 39, 8)
   Call REPLA(X$, MKI$(VENDE%), 47, 8)
   Call REPLA(X$, MKD$(IINT#), 49, 8)
   Call REPLA(X$, MKD$(PERC#), 57, 8)
   '
   Call REGAPP("FACTUR", X$)
   Call CIERRARCH("FACTUR")
   '
   '
   CPAG% = 0
   Call GRADEU(LUDAT$, "SVD202", RDX%, TIDOC%, VARIFAC%, NROPRO&, NC%, FEFAC%, CPAG%, TOTA#)
   Call CIERRARCH("SVD202")
   '
   '
   RGX& = REGICLI&(NC%)
   If RGX& > 0 Then
      If RGX& <= ULTREG&("DEUDOR2") Then
         DX$ = REGLEIDO$("DEUDOR2", RGX&)
         SALDO# = CVD(Mid$(DX$, 109, 8)) - TOTA#
         Call REPLA(DX$, MKD$(SALDO#), 109, 8)
         Call GRAREG("DEUDOR2", DX$, RGX&)
      End If
   End If
   Call CIERRARCH("DEUDOR2")
   '
   '
   For U& = 1 To FORCREDEVDIP.List2.ListCount
      Z$ = FORCREDEVDIP.List2.List(U& - 1)
      CAN = Val(Left$(Z$, 4))
      If CAN >= 0.05 Then
         CODX$ = Mid$(Z$, 7, 15)
         CIIT% = CODINT%(CODX$)
         PREU = Val(Mid$(Z$, 56, 9))
         ITOT# = Val(Mid$(Z$, 74, 10))
         Call MOVISTO(FEFAC%, CIIT%, "", "DV", "DV." + TRIM$(Str$(NROPRO&)), "NC." + TRIM$(Str$(NROPRO&)), RASOCLI$(NC%), CDbl(PREU), "$", NC%, 1, CAN)
      End If
   Next U&
   Call CIERRARCH("MOVISTO")
   Call CIERRARCH("BKMOVIST")
   '
   OKEY% = 1
   '
End Sub
'
Sub GRAFACABO(OKEY%)
   '
   HOII% = HOY(HO$)
   OKEY% = 0
   '
   NC% = Val(FORFACABO.Text1.TEXT)
   If NC% < 1 Or ECOARCH$("DEUDOR1", MKI$(NC%)) = "" Then
       Call MENSERR(24, "Numero de Cuenta no Válido o Inexistente")
       FORFACABO.Text1.SetFocus
       Exit Sub
   End If
   '
   If FORFACABO.Text8.TEXT = "" Then
       Call MENSERR(24, "Falta Ingresar Detalle de Factura de Abono")
       FORFACABO.Text8.SetFocus
       Exit Sub
   End If
   '
   FECHDOC$ = FORFACABO.Text9.TEXT
   FEFAC% = FECHANUM(FECHDOC$)
   If FEFAC% < 8600 Then
      Call MENSERR(24, "Fecha de Factura de Abono no Válida")
      FORFACABO.Text9.SetFocus
      Exit Sub
   End If
   FEX = FEFAC%
   FORFACABO.Text9.TEXT = FECHATEX$(FEX)
   '
   CPAG% = Val(FORFACABO.TEXT21(1).TEXT)
   If ECOARCH$("CONPAG", MKI$(CPAG%)) = "" Then
      Call MENSERR(24, "Condición de Pago no Válida")
      FORFACABO.TEXT21(1).SetFocus
      Exit Sub
   End If
   CONPAG$ = DECONPA$(CPAG%)
   
   VEND% = Val(FORFACABO.TEXT21(1).TEXT)
   VENDE$ = DEVENDE$(VEND%)
   
   Call BLOQARCH("FACTUR")
   Call BLOQARCH("SVD202")
   Call BLOQARCH("DEUDOR2")
   '
   NROPRO& = 0
   If Left$(FORFACABO.Text11.TEXT, 3) = "FC." Then NROPRO& = Val(Mid$(FORFACABO.Text11.TEXT, 4))
   Call NUFACTURA(NROPRO&)
   FORFACABO.Text11.TEXT = "FC." + TRIM$(Str$(NROPRO&))
   '
   X$ = REGBLAN$("FACTUR")
   Call REPLA(X$, MKI$(NC%), 1, 2)
   Call REPLA(X$, MKS$(NROPRO&), 3, 4)
   Call REPLA(X$, MKI$(FEFAC%), 7, 2)
   Call REPLA(X$, MKI$(FVI%), 9, 2)
   Call REPLA(X$, MKI$(0), 11, 2)
   Call REPLA(X$, Chr$(TIDOC%), 13, 1)
   Call REPLA(X$, Chr$(VARIFAC%), 14, 1)
   '
   NEFA# = Val(FORFACABO.Label10.Caption)
   NIVA# = Val(FORFACABO.Label18.Caption)
   NNIN# = Val(FORFACABO.Label23.Caption)
   TOTA# = Val(FORFACABO.Label20.Caption)
   IINT# = 0
   PERC# = 0
   '
   If NEFA# < 0.01 Then
      Call MENSERR(24, "Faltan Importes")
      Exit Sub
   End If
   '
   Call REPLA(X$, MKD$(NEFA#), 15, 8)
   Call REPLA(X$, MKD$(NIVA#), 23, 8)
   Call REPLA(X$, MKD$(NNIN#), 31, 8)
   Call REPLA(X$, MKD$(TOTA#), 39, 8)
   Call REPLA(X$, MKI$(VEND%), 47, 8)
   Call REPLA(X$, MKD$(IINT#), 49, 8)
   Call REPLA(X$, MKD$(PERC#), 57, 8)
   '
   Call REGAPP("FACTUR", X$)
   Call CIERRARCH("FACTUR")
   '
   Call GRADEU(LUDAT$, "SVD202", RDX%, TIDOC%, VARIFAC%, NROPRO&, NC%, FEFAC%, CPAG%, TOTA#)
   Call CIERRARCH("SVD202")
   '
   RGX& = REGICLI&(NC%)
   If RGX& > 0 Then
      If RGX& <= ULTREG&("DEUDOR2") Then
         DX$ = REGLEIDO$("DEUDOR2", RGX&)
         SALDO# = CVD(Mid$(DX$, 109, 8)) + TOTA#
         Call REPLA(DX$, MKD$(SALDO#), 109, 8)
         Call GRAREG("DEUDOR2", DX$, RGX&)
      End If
   End If
   Call CIERRARCH("DEUDOR2")
   '
   OKEY% = 1
   '
End Sub
'
Sub PRIFACDIP()
    '
    TTXX$ = Mid$(ENCAPED$, 301, 60): Call GRAREG("!OBSERVOF", TTXX$, 1)
    TTXX$ = Mid$(ENCAPED$, 361, 60): Call GRAREG("!OBSERVOF", TTXX$, 2)
    TTXX$ = Mid$(ENCAPED$, 421, 60): Call GRAREG("!OBSERVOF", TTXX$, 3)
    Call SETULTREG("!OBSERVOF", 3)
    Call CIERRARCH("!OBSERVOF")
    '
    FORIPRE$ = CONTROL$("", "FORFAC-" + LETRAFAC$)
    If FORIPRE$ = "" Then
       Call MENSERR(24, "Falta Definir Formulario de Impresión\para Facturación Tipo '" + LETRAFAC$ + "'.")
       Exit Sub
    End If
    '
    Screen.MousePointer = 11
    HOII% = HOY(HOS$)
    FECHDOC$ = FORFACDIP.Text9.TEXT
    NROPE& = Val(Mid$(FORFACDIP.Text11.TEXT, 4))
    NUMEDOC$ = TRIM$(Str$(NROPE&))
    While Len(NUMEDOC$) < 8
      NUMEDOC$ = "0" + NUMEDOC$
    Wend
    '
    PREFINUM$ = ATRIFORM$(FORIPRE$, "PREFIFAC")
    While Len(PREFINUM$) < 4
      PREFINUM$ = "0" + PREFINUM$
    Wend
    NUMEDOC$ = PREFINUM$ + "-" + NUMEDOC$
    NUREMITO$ = NUMEDOC$
    NUFACORI$ = NUMEDOC$
    TIPODOC$ = "Factura"
    If LETRAFAC$ = "E" Then
       TIPODOC$ = "Factura de Exportación"
    End If
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
    ITOTCO# = 0
    '
    AAA = ALIVA
    BBB$ = LETRAFAC$
    For T& = 1 To ULTREG&("!CARITPED")
       Y$ = REGLEIDO$("!CARITPED", T&)
       CIIT% = CVI(Mid$(Y$, 3, 2))
       CAIT = CVS(Mid$(Y$, 5, 4))
       PREUNITA = CVS(Mid$(Y$, 17, 4)): If LETRAFAC$ = "B" Then PREUNITA = PREUNITA * (100 + ALIVA) / 100
       PRECO$ = TRIM$(CSTRING$(MKS$(PREUNITA), 3, 12, 4, 2))
       TOTAITEM# = CVD(Mid$(Y$, 21, 8)): If LETRAFAC$ = "B" Then TOTAITEM# = TOTAITEM# * (100 + ALIVA) / 100
       TOTEX$ = TRIM$(CSTRING$(MKD$(TOTAITEM#), 3, 12, 2, 2))
       ENVASE$ = TRIM$(Mid$(Y$, 33, 28))
       '
       CAITAB1% = CAITAB1% + 1
       TETABU1$(1, CAITAB1%) = CODEXT$(CIIT%)
       TETABU1$(2, CAITAB1%) = TRIM$(ENVASE$ + " " + DESCRIT0$(CIIT%))
       TETABU1$(3, CAITAB1%) = Str$(CAIT)
       TETABU1$(4, CAITAB1%) = PRECO$
       TETABU1$(5, CAITAB1%) = TOTEX$
       TETABU1$(6, CAITAB1%) = ""
       '
    Next T&
    '
    CPAG% = Asc(Mid$(ENCAPED$, 298, 1))
    VEND% = Asc(Mid$(ENCAPED$, 299, 1))
    CONPA$ = DECONPA$(CPAG%)
    NOCOM$ = Mid$(ENCAPED$, 137, 18)
    '
    CODPARAM$(1) = "CIVA-CLI": DOCPARAM$(1) = LETRAFAC$
    CODPARAM$(2) = "NRO-OCOM": DOCPARAM$(2) = NOCOM$
    CODPARAM$(3) = "COND-PAG": DOCPARAM$(3) = CONPA$
    CODPARAM$(4) = "IMP-NETO": DOCPARAM$(4) = FORFACDIP.Label14.Caption
    CODPARAM$(5) = "POR-DESC": DOCPARAM$(5) = CSTRING$(MKS$(PORDESCUENTO), 3, 5, 1, 2)
    IDESCUENTO = Val(FORFACDIP.Label14.Caption) * (PORDESCUENTO / 100)
    If UCase$(LETRAFAC$) = "B" Then IDESCUENTO = IDESCUENTO * (1 + ALIVA / 100)
    CODPARAM$(6) = "IMP-DESC": DOCPARAM$(6) = CSTRING$(MKS$(IDESCUENTO), 3, 12, 2, 2)
    CODPARAM$(7) = "IMPNEDES": DOCPARAM$(7) = CSTRING$(MKS$(Val(FORFACDIP.Label14.Caption) * (1 - PORDESCUENTO / 100)), 3, 12, 2, 2)
    CODPARAM$(8) = "IMP-IVA": DOCPARAM$(8) = CSTRING$(MKS$(Val(FORFACDIP.Label18.Caption) * (1 - PORDESCUENTO / 100)), 3, 12, 2, 2)
    CODPARAM$(9) = "IMP-RNI": DOCPARAM$(9) = CSTRING$(MKS$(Val(FORFACDIP.Label21.Caption) * (1 - PORDESCUENTO / 100)), 3, 12, 2, 2)
    CODPARAM$(10) = "IMP-TOTA": DOCPARAM$(10) = CSTRING$(MKS$(Val(FORFACDIP.Label20.Caption) * (1 - PORDESCUENTO / 100)), 3, 12, 2, 2)
    '
    ITEXTO$ = NUMTEX$((Val(FORFACDIP.Label20.Caption)) * (1 - PORDESCUENTO / 100)) + "."
    LOTEX% = Val(ATRIFORM$("", "LOIMPLET")): If LOTEX% < 20 Then LOTEX% = 20
    Call FRATEX(ITEXTO$, NRX1$, NRX2$, NRX3$, LOTEX%)
    IMPLET1$ = TRIM$(NRX1$) + " " + "***************************************************************************************************************"
    IMPLET2$ = TRIM$(NRX2$) + " " + "***************************************************************************************************************"
    '
    CODPARAM$(11) = "IMP-LET1": DOCPARAM$(11) = IMPLET1$
    CODPARAM$(12) = "IMP-LET2": DOCPARAM$(12) = IMPLET2$
    '
    CODPARAM$(13) = "FECH-EMI": DOCPARAM$(13) = FORFACDIP.Text9.TEXT
    CODPARAM$(14) = "NRO-REMI": DOCPARAM$(14) = TRIM$(FORFACDIP.Label2.Caption)
    CODPARAM$(15) = "NRO-AENT": DOCPARAM$(15) = REFPED$
    CODPARAM$(16) = "TIPO-COM": DOCPARAM$(16) = UCase$(TIPODOC$)
    CODPARAM$(17) = "ITO-IVA": DOCPARAM$(17) = Str$(Val(DOCPARAM$(5)) + Val(DOCPARAM$(6)))
    CODPARAM$(18) = "VENDEDOR": DOCPARAM$(18) = TRIM$(Str$(VEND%))
    '
    CAPARDOC% = 18
    '
    FEVENCIM$ = ""
    If CAPLAVENCI% > 0 Then FEVENCIM$ = FECHATEX$(FECHAVENCI(1))
    If CAPLAVENCI% > 1 Then FEVENCIM$ = FEVENCIM$ + " y ss"
    CAPARDOC% = CAPARDOC% + 1
    CODPARAM$(CAPARDOC%) = "FECH-VEN": DOCPARAM$(CAPARDOC%) = FEVENCIM$
    '
    Call PRINTDOC("FORFAC-" + LETRAFAC$) ' FACTURACION
    '
End Sub
'

Sub CLEARFORFACDIP()
   '
   Call CIERRARCH("!CARCONPE")
   Call CIERRARCH("!CARITPED")
   Call BLANARCH("!CARCONPE")
   Call BLANARCH("!CARITPED")
   '
   FORFACTURA.List2.Clear
   FORFACTURA.Text1.TEXT = ""
   FORFACTURA.Text2.TEXT = ""
   FORFACTURA.Text3.TEXT = ""
   FORFACTURA.Text4.TEXT = ""
   FORFACTURA.Text5.TEXT = ""
   FORFACTURA.Text6.TEXT = "X"
   FORFACTURA.Text7.TEXT = ""
   '
   FEFAC% = HOY(HO$)
   FORFACTURA.Text9.TEXT = HO$
    '
   FORFACTURA.Text11.TEXT = ""
   '
   FORFACTURA.Label14.Caption = ""
   FORFACTURA.Label18.Caption = ""
   FORFACTURA.Label20.Caption = ""
   '
   FORFACTURA.Text1.SetFocus
   '
End Sub
'
Sub CLEARDEBILIB()
   '
   FORDELIDIP.Text1.TEXT = ""
   FORDELIDIP.Text2.TEXT = ""
   FORDELIDIP.Text3.TEXT = ""
   FORDELIDIP.Text4.TEXT = ""
   FORDELIDIP.Text5.TEXT = ""
   FORDELIDIP.Text6.TEXT = "X"
   FORDELIDIP.Text7.TEXT = ""
   '
   FEFAC% = HOY(HO$)
   FORDELIDIP.Text9.TEXT = HO$
    '
   FORDELIDIP.Text11.TEXT = ""
   '
   FORDELIDIP.Text10.TEXT = ""
   FORDELIDIP.Label18.Caption = ""
   FORDELIDIP.Label20.Caption = ""
   FORDELIDIP.Label23.Caption = ""
   '
   FORDELIDIP.Text1.SetFocus
   '

End Sub
'

Sub CLEARCRELIB()
   '
   FORCRELIDIP.Text1.TEXT = ""
   FORCRELIDIP.Text2.TEXT = ""
   FORCRELIDIP.Text3.TEXT = ""
   FORCRELIDIP.Text4.TEXT = ""
   FORCRELIDIP.Text5.TEXT = ""
   FORCRELIDIP.Text6.TEXT = "X"
   FORCRELIDIP.Text7.TEXT = ""
   '
   FEFAC% = HOY(HO$)
   FORCRELIDIP.Text9.TEXT = HO$
    '
   FORCRELIDIP.Text11.TEXT = ""
   '
   FORCRELIDIP.Text10.TEXT = ""
   FORCRELIDIP.Label18.Caption = ""
   FORCRELIDIP.Label20.Caption = ""
   FORCRELIDIP.Label23.Caption = ""
   '
   FORCRELIDIP.Text1.SetFocus
   '

End Sub
'

Sub CLEARFORCREDEV()
   '
   Call CIERRARCH("!CARITPED")
   Call BLANARCH("!CARITPED")
   '
   FORCREDEVDIP.List2.Clear
   FORCREDEVDIP.Text1.TEXT = ""
   FORCREDEVDIP.Text2.TEXT = ""
   FORCREDEVDIP.Text3.TEXT = ""
   FORCREDEVDIP.Text4.TEXT = ""
   FORCREDEVDIP.Text5.TEXT = ""
   FORCREDEVDIP.Text6.TEXT = "X"
   FORCREDEVDIP.Text7.TEXT = ""
   '
   FEFAC% = HOY(HO$)
   FORCREDEVDIP.Text9.TEXT = HO$
   '
   FORCREDEVDIP.Text11.TEXT = ""
   '
   FORCREDEVDIP.Label14.Caption = ""
   FORCREDEVDIP.Label18.Caption = ""
   FORCREDEVDIP.Label20.Caption = ""
   '
   FORCREDEVDIP.Text1.SetFocus
   '
End Sub
'

Sub CLEARFACABO()
   '
   FORFACABO.Text1.TEXT = ""
   FORFACABO.Text2.TEXT = ""
   FORFACABO.Text3.TEXT = ""
   FORFACABO.Text4.TEXT = ""
   FORFACABO.Text5.TEXT = ""
   FORFACABO.Text6.TEXT = "X"
   FORFACABO.Text7.TEXT = ""
   '
   FEFAC% = HOY(HO$)
   FORFACABO.Text9.TEXT = HO$
    '
   FORFACABO.Text11.TEXT = ""
   '
   FORFACABO.Label10.Caption = ""
   FORFACABO.Label18.Caption = ""
   FORFACABO.Label20.Caption = ""
   FORFACABO.Label23.Caption = ""
   '
   FORFACABO.Text1.SetFocus
   '

End Sub

Sub CLEARCONFABO()
   '
   FORCONFABO.Text1.TEXT = ""
   FORCONFABO.Text2.TEXT = ""
   FORCONFABO.Text3.TEXT = ""
   FORCONFABO.Text4.TEXT = ""
   FORCONFABO.Text5.TEXT = ""
   FORCONFABO.Text6.TEXT = "X"
   FORCONFABO.Text7.TEXT = ""
   '
   FORCONFABO.Text10.TEXT = ""
   FORCONFABO.Label18.Caption = ""
   FORCONFABO.Label20.Caption = ""
   FORCONFABO.Label23.Caption = ""
   '
   FORCONFABO.Text1.SetFocus
   '

End Sub
'
Sub DEBILIB1(OKEY%)
   '
   OKEY% = 0
   '
   NC% = Val(FORDELIDIP.Text1.TEXT)
   If NC% < 1 Then
      Call MENSERR(24, "Número de Cuenta no Válido")
      Exit Sub
   End If
   '
   If ECOARCH$("DEUDOR1", MKI$(NC%)) = "" Then
      Call MENSERR(24, "Cuenta Inexistente")
      Exit Sub
   End If
   '
   FORDELIDIP.Text1.TEXT = Str$(NC%)
   FORDELIDIP.Text2.TEXT = RASOCLI$(NC%)
   FORDELIDIP.Text3.TEXT = DOMICLI$(NC%)
   FORDELIDIP.Text4.TEXT = CPOSCLI$(NC%)
   FORDELIDIP.Text5.TEXT = LOCACLI$(NC%)
   FORDELIDIP.Text7.TEXT = TRIM$(CUITCLI$(NC%)) + " - " + POSIVAC$(NC%)
   '
   FORDELIDIP.TEXT21(1).TEXT = Str$(CPAGCLI%(NC%))
   FORDELIDIP.TEXT21(2).TEXT = Str$(VENDCLI%(NC%))
   If Val(FORDELIDIP.TEXT21(2).TEXT) < 1 Then
      FORDELIDIP.Eco21(2).TEXT = "GERENCIA"
   End If
   '
   PIVAFAC% = PIVACLI%(NC%)
   If PIVAFAC% <> 0 Then
      If PIVAFAC% <> 5 Then
         If TRIM$(CUITCLI$(NC%)) = "" Or VALICUIT%(CUITCLI$(NC%)) <> 1 Then
            Call MENSERR(24, "Imposible Facturar a Esta Cuenta.\Número de C.U.I.T. no Válido.")
            Exit Sub
         End If
      End If
   End If
   '
   Call CARLETFAC
   FORDELIDIP.Text6.TEXT = LETRAFAC$
   FORDELIDIP.Text12.TEXT = TRIM$(Str$(ALIVA))
   If PIVAFAC% = 5 Then
      Text12.TEXT = "0.0"
   End If
   TIDOC% = 3                 ' tipo de documento = NOTA DE DEBITO
   '
   NROPRO& = 0
   If Left$(FORDELIDIP.Text11.TEXT, 3) = "ND." Then NROPRO& = Val(Mid$(FORDELIDIP.Text11.TEXT, 4))
   Call NUFACTURA(NROPRO&)
   FORDELIDIP.Text11.TEXT = "ND." + TRIM$(Str$(NROPRO&))
   '
   OKEY% = 1
   '
End Sub

Sub CREDILIB1(OKEY%)
   '
   OKEY% = 0
   '
   NC% = Val(FORCRELIDIP.Text1.TEXT)
   If NC% < 1 Then
      Call MENSERR(24, "Número de Cuenta no Válido")
      Exit Sub
   End If
   '
   If ECOARCH$("DEUDOR1", MKI$(NC%)) = "" Then
      Call MENSERR(24, "Cuenta Inexistente")
      Exit Sub
   End If
   '
   FORCRELIDIP.Text1.TEXT = Str$(NC%)
   FORCRELIDIP.Text2.TEXT = RASOCLI$(NC%)
   FORCRELIDIP.Text3.TEXT = DOMICLI$(NC%)
   FORCRELIDIP.Text4.TEXT = CPOSCLI$(NC%)
   FORCRELIDIP.Text5.TEXT = LOCACLI$(NC%)
   FORCRELIDIP.Text7.TEXT = TRIM$(CUITCLI$(NC%)) + " - " + POSIVAC$(NC%)
   '
   FORCRELIDIP.TEXT21(1).TEXT = "0"
   FORCRELIDIP.Eco21(1) = "CONTADO"
   FORCRELIDIP.TEXT21(2).TEXT = Str$(VENDCLI%(NC%))
   If Val(FORCRELIDIP.TEXT21(2).TEXT) < 1 Then
      FORCRELIDIP.Eco21(2).TEXT = "GERENCIA"
   End If
   '
   PIVAFAC% = PIVACLI%(NC%)
   If PIVAFAC% <> 0 Then
      If PIVAFAC% <> 5 Then
         If TRIM$(CUITCLI$(NC%)) = "" Or VALICUIT%(CUITCLI$(NC%)) <> 1 Then
            Call MENSERR(24, "Imposible Facturar a Esta Cuenta.\Número de C.U.I.T. no Válido.")
            Exit Sub
         End If
      End If
   End If
   '
   Call CARLETFAC
   FORCRELIDIP.Text6.TEXT = LETRAFAC$
   TIDOC% = 6                 ' tipo de documento = NOTA DE CREDITO
   '
   NROPRO& = 0
   If Left$(FORCRELIDIP.Text11.TEXT, 3) = "NC." Then NROPRO& = Val(Mid$(FORCRELIDIP.Text11.TEXT, 4))
   Call NUFACTURA(NROPRO&)
   FORCRELIDIP.Text11.TEXT = "NC." + TRIM$(Str$(NROPRO&))
   '
   OKEY% = 1
   '
End Sub

Sub CREDEVOL1(OKEY%)
   '
   OKEY% = 0
   '
   NC% = Val(FORCREDEVDIP.Text1.TEXT)
   If NC% < 1 Then
      Call MENSERR(24, "Número de Cuenta no Válido")
      Exit Sub
   End If
   '
   If ECOARCH$("DEUDOR1", MKI$(NC%)) = "" Then
      Call MENSERR(24, "Cuenta Inexistente")
      Exit Sub
   End If
   '
   FORCREDEVDIP.Text1.TEXT = Str$(NC%)
   FORCREDEVDIP.Text2.TEXT = RASOCLI$(NC%)
   FORCREDEVDIP.Text3.TEXT = DOMICLI$(NC%)
   FORCREDEVDIP.Text4.TEXT = CPOSCLI$(NC%)
   FORCREDEVDIP.Text5.TEXT = LOCACLI$(NC%)
   FORCREDEVDIP.Text7.TEXT = TRIM$(CUITCLI$(NC%)) + " - " + POSIVAC$(NC%)
   '
   FORCREDEVDIP.TEXT21(1).TEXT = Str$(LIPRCLI%(NC%))
   If Val(FORCREDEVDIP.TEXT21(1)) < 1 Then FORCREDEVDIP.TEXT21(1) = "1"
   FORCREDEVDIP.Eco21(1).TEXT = DELIPRE$(Val(FORCREDEVDIP.TEXT21(1).TEXT))
   If FORCREDEVDIP.Eco21(1).TEXT = "" Then
      Call MENSERR(24, "Lista de precios no Válida")
      FORCREDEVDIP.TEXT21(1).SetFocus
      Exit Sub
   End If
   LPRE% = Val(FORCREDEVDIP.TEXT21(1).TEXT)
   ARCHIPRE$ = TRIM$(Str$(LPRE%))
   While Len(ARCHIPRE$) < 3: ARCHIPRE$ = "0" + ARCHIPRE$: Wend
   ARCHIPRE$ = "LIPRE" + ARCHIPRE$
   '
   FORCREDEVDIP.TEXT21(2).TEXT = Str$(VENDCLI%(NC%))
   FORCREDEVDIP.Eco21(2).TEXT = DEVENDE$(Val(FORCREDEVDIP.TEXT21(2).TEXT))
   If Val(FORCREDEVDIP.TEXT21(2).TEXT) < 1 Then
      FORCREDEVDIP.Eco21(2).TEXT = "GERENCIA"
   End If
   '
   PIVAFAC% = PIVACLI%(NC%)
   If PIVAFAC% <> 0 Then
      If PIVAFAC% <> 5 Then
         If TRIM$(CUITCLI$(NC%)) = "" Or VALICUIT%(CUITCLI$(NC%)) <> 1 Then
            Call MENSERR(24, "Imposible Facturar a Esta Cuenta.\Número de C.U.I.T. no Válido.")
            Exit Sub
         End If
      End If
   End If
   '
   Call CARLETFAC
   FORCREDEVDIP.Text6.TEXT = LETRAFAC$
   TIDOC% = 6                 ' tipo de documento = NOTA DE CREBITO
   '
   NROPRO& = 0
   If Left$(FORCREDEVDIP.Text11.TEXT, 3) = "NC." Then NROPRO& = Val(Mid$(FORCREDEVDIP.Text11.TEXT, 4))
   Call NUFACTURA(NROPRO&)
   FORCREDEVDIP.Text11.TEXT = "NC." + TRIM$(Str$(NROPRO&))
   '
   OKEY% = 1
   '
End Sub
'

Sub FACABO1(OKEY%)
   '
   OKEY% = 0
   '
   NC% = Val(FORFACABO.Text1.TEXT)
   If NC% < 1 Then
      Call MENSERR(24, "Número de Cuenta no Válido")
      Exit Sub
   End If
   '
   If ECOARCH$("DEUDOR1", MKI$(NC%)) = "" Then
      Call MENSERR(24, "Cuenta Inexistente")
      Exit Sub
   End If
   '
   FORFACABO.Text1.TEXT = Str$(NC%)
   FORFACABO.Text2.TEXT = RASOCLI$(NC%)
   FORFACABO.Text3.TEXT = DOMICLI$(NC%)
   FORFACABO.Text4.TEXT = CPOSCLI$(NC%)
   FORFACABO.Text5.TEXT = LOCACLI$(NC%)
   FORFACABO.Text7.TEXT = TRIM$(CUITCLI$(NC%)) + " - " + POSIVAC$(NC%)
   '
   FORFACABO.TEXT21(1).TEXT = Str$(CPAGCLI%(NC%))
   FORFACABO.TEXT21(2).TEXT = Str$(VENDCLI%(NC%))
   If Val(FORFACABO.TEXT21(2).TEXT) < 1 Then
      FORFACABO.Eco21(2).TEXT = "GERENCIA"
   End If
   '
   PIVAFAC% = PIVACLI%(NC%)
   If PIVAFAC% <> 0 Then
      If PIVAFAC% <> 5 Then
         If TRIM$(CUITCLI$(NC%)) = "" Or VALICUIT%(CUITCLI$(NC%)) <> 1 Then
            Call MENSERR(24, "Imposible Facturar a Esta Cuenta.\Número de C.U.I.T. no Válido.")
            Exit Sub
         End If
      End If
   End If
   '
   Call CARLETFAC
   FORFACABO.Text6.TEXT = LETRAFAC$
   TIDOC% = 2                ' tipo de documento = NOTA DE DEBITO
   '
   NROPRO& = 0
   If Left$(FORFACABO.Text11.TEXT, 3) = "FC." Then NROPRO& = Val(Mid$(FORFACABO.Text11.TEXT, 4))
   Call NUFACTURA(NROPRO&)
   FORFACABO.Text11.TEXT = "FC." + TRIM$(Str$(NROPRO&))
   '
   RFF$ = RECFILT$("ABONOS", 1, MKI$(NC%))
   If Len(RFF$) < 4 Then
      Call MENSERR(24, "No Hay Abonos Definidos\para la Cuenta Elegida.")
      Call CLEARFACABO
      FORFACABO.Text1.SetFocus
      Exit Sub
   End If
   '
   If Len(RFF$) >= 4 Then
      REABO& = CVS(Left$(RFF$, 4))
      If REABO& > 0 Then
         If REABO& <= ULTREG&("ABONOS") Then
            TTXX$ = REGLEIDO$("ABONOS", REABO&)
            CPAG% = Asc(Mid$(TTXX$, 3, 1))
            IABO# = CVD(Mid$(TTXX$, 5, 8))
            If IABO# < 0.5 Then
               Call MENSERR(24, "No Hay Abonos Definidos\para la Cuenta Elegida.")
               Call CLEARFACABO
               FORFACABO.Text1.SetFocus
               Exit Sub
            End If
            If CPAG% > 0 Then
               FORFACABO.TEXT21(1).TEXT = Str$(CPAG%)
            End If
            VEND% = Asc(Mid$(TTXX$, 4, 1))
            If VEND% > 0 Then
               FORFACABO.TEXT21(2).TEXT = Str$(VEND%)
            End If
            FORFACABO.Label10.Caption = CSTRING$(Mid$(TTXX$, 5, 8), 3, 10, 2, 2)
            TEXTOABO$ = TRIM$(Mid$(TTXX$, 25, 1000))
            '
25          PPXX% = InStr(TEXTOABO$, "#")
            If PPXX% > 0 Then
               TEXTOABO$ = Left$(TEXTOABO$, PPXX% - 1) + MESABONO$ + Mid$(TEXTOABO$, PPXX% + 1)
               GoTo 25
            End If
            FORFACABO.Text8.TEXT = TEXTOABO$
         End If
      End If
   End If
   '
   OKEY% = 1
   '
End Sub
'
Sub CONFABO1(OKEY%)
   '
   OKEY% = 0
   '
   NC% = Val(FORCONFABO.Text1.TEXT)
   If NC% < 1 Then
      Call MENSERR(24, "Número de Cuenta no Válido")
      Exit Sub
   End If
   '
   If ECOARCH$("DEUDOR1", MKI$(NC%)) = "" Then
      Call MENSERR(24, "Cuenta Inexistente")
      Exit Sub
   End If
   '
   FORCONFABO.Text1.TEXT = Str$(NC%)
   FORCONFABO.Text2.TEXT = RASOCLI$(NC%)
   FORCONFABO.Text3.TEXT = DOMICLI$(NC%)
   FORCONFABO.Text4.TEXT = CPOSCLI$(NC%)
   FORCONFABO.Text5.TEXT = LOCACLI$(NC%)
   FORCONFABO.Text7.TEXT = TRIM$(CUITCLI$(NC%)) + " - " + POSIVAC$(NC%)
   '
   FORCONFABO.TEXT21(1).TEXT = Str$(CPAGCLI%(NC%))
   FORCONFABO.TEXT21(2).TEXT = Str$(VENDCLI%(NC%))
   If Val(FORCONFABO.TEXT21(2).TEXT) < 1 Then
      FORCONFABO.Eco21(2).TEXT = "GERENCIA"
   End If
   '
   PIVAFAC% = PIVACLI%(NC%)
   If PIVAFAC% <> 0 Then
      If PIVAFAC% <> 5 Then
         If TRIM$(CUITCLI$(NC%)) = "" Or VALICUIT%(CUITCLI$(NC%)) <> 1 Then
            Call MENSERR(24, "Imposible Facturar a Esta Cuenta.\Número de C.U.I.T. no Válido.")
            Exit Sub
         End If
      End If
   End If
   '
   Call CARLETFAC
   FORCONFABO.Text6.TEXT = LETRAFAC$
   FORCONFABO.Text12.TEXT = TRIM$(Str$(ALIVA))
   If PIVAFAC% = 5 Then
      Text12.TEXT = "0.0"
   End If
   '
   RFF$ = RECFILT$("ABONOS", 1, MKI$(NC%))
   If Len(RFF$) >= 4 Then
      REABO& = CVS(Left$(RFF$, 4))
      If REABO& > 0 Then
         If REABO& <= ULTREG&("ABONOS") Then
            TTXX$ = REGLEIDO$("ABONOS", REABO&)
            CPAG% = Asc(Mid$(TTXX$, 3, 1))
            If CPAG% > 0 Then
               FORCONFABO.TEXT21(1).TEXT = Str$(CPAG%)
            End If
            VEND% = Asc(Mid$(TTXX$, 4, 1))
            If VEND% > 0 Then
               FORCONFABO.TEXT21(2).TEXT = Str$(VEND%)
            End If
            FORCONFABO.Text10.TEXT = CSTRING$(Mid$(TTXX$, 5, 8), 3, 10, 2, 2)
            FORCONFABO.Text8.TEXT = TRIM$(Mid$(TTXX$, 25, 1000))
         End If
      End If
   End If
   '
   OKEY% = 1
   '
End Sub


Sub PRIDEBILIB()
    '
    FORIPRE$ = CONTROL$("", "FORFAC-" + LETRAFAC$)
    If FORIPRE$ = "" Then
       Call MENSERR(24, "Falta Definir Formulario de Impresión\para Facturación Tipo '" + LETRAFAC$ + "'.")
       Exit Sub
    End If
    '
    Screen.MousePointer = 11
    HOII% = HOY(HOS$)
    FECHDOC$ = FORDELIDIP.Text9.TEXT
    NROPE& = Val(Mid$(FORDELIDIP.Text11.TEXT, 4))
    NUMEDOC$ = TRIM$(Str$(NROPE&))
    While Len(NUMEDOC$) < 8
      NUMEDOC$ = "0" + NUMEDOC$
    Wend
    '
    PREFINUM$ = ATRIFORM$(FORIPRE$, "PREFIFAC")
    While Len(PREFINUM$) < 4
      PREFINUM$ = "0" + PREFINUM$
    Wend
    NUMEDOC$ = PREFINUM$ + "-" + NUMEDOC$
    NUREMITO$ = NUMEDOC$
    TIPODOC$ = "Nota de Débito"
    '
    CODTABU1$(1) = "DES-MAT"
    CODTABU1$(2) = "NET-ITEM"
    CACOLTAB1% = 2
    '
    DESTIDOC% = NC%
    CAITAB1% = 0
    CAITCRU% = 0
    '
    SUMATOTA# = 0: J& = 0
    ITOT# = Val(FORDELIDIP.Text10.TEXT)
    If LETRAFAC$ = "B" Then
       ITOT# = (Int(ITOT# * (100 + ALIVA) + 0.5)) / 100
    End If
    '
    TTXX$ = FORDELIDIP.Text8.TEXT
    Call FRATEXTO(TTXX$, 48)
    '
    For KK% = 1 To CARETEX%
       CAITAB1% = CAITAB1% + 1
       If TRIM$(RETEX$(KK%)) = "" Then RETEX$(KK%) = "."
       TETABU1$(1, CAITAB1%) = RETEX$(KK%)
       TETABU1$(2, CAITAB1%) = Space$(10)
    Next KK%
    '
    ITOTX$ = Str$(ITOT#)
    TETABU1$(2, CAITAB1%) = ITOTX$
    '
    CPAG% = Val(FORDELIDIP.TEXT21(1).TEXT)
    CONPA$ = DECONPA$(CPAG%)
    NOCOM$ = ""
    '
    CODPARAM$(1) = "CIVA-CLI": DOCPARAM$(1) = LETRAFAC$
    CODPARAM$(2) = "NRO-OCOM": DOCPARAM$(2) = ""
    CODPARAM$(3) = "COND-PAG": DOCPARAM$(3) = CONPA$
    CODPARAM$(4) = "IMP-NETO": DOCPARAM$(4) = FORDELIDIP.Text10.TEXT
    CODPARAM$(5) = "IMP-IVA": DOCPARAM$(5) = FORDELIDIP.Label18.Caption
    CODPARAM$(6) = "IMP-RNI": DOCPARAM$(6) = FORDELIDIP.Label23.Caption
    CODPARAM$(7) = "IMP-TOTA": DOCPARAM$(7) = FORDELIDIP.Label20.Caption
    '
    ITEXTO$ = NUMTEX$(Val(FORDELIDIP.Label20.Caption)) + "."
    LOTEX% = Val(ATRIFORM$("", "LOIMPLET")): If LOTEX% < 20 Then LOTEX% = 20
    Call FRATEX(ITEXTO$, NRX1$, NRX2$, NRX3$, LOTEX%)
    IMPLET1$ = TRIM$(NRX1$) + " " + "***************************************************************************************************************"
    IMPLET2$ = TRIM$(NRX2$) + " " + "***************************************************************************************************************"
    '
    CODPARAM$(8) = "IMP-LET1": DOCPARAM$(8) = IMPLET1$
    CODPARAM$(9) = "IMP-LET2": DOCPARAM$(9) = IMPLET2$
    '
    CODPARAM$(10) = "FECH-EMI": DOCPARAM$(10) = FORDELIDIP.Text9.TEXT
    CODPARAM$(11) = "NRO-REMI": DOCPARAM$(11) = ""
    CODPARAM$(12) = "NRO-AENT": DOCPARAM$(12) = ""
    CODPARAM$(13) = "TIPO-COM": DOCPARAM$(13) = UCase$(TIPODOC$)
    CODPARAM$(14) = "ITO-IVA": DOCPARAM$(14) = Str$(Val(DOCPARAM$(5)) + Val(DOCPARAM$(6)))
    CAPARDOC% = 14
    '
    FEVENCIM$ = ""
    If CAPLAVENCI% > 0 Then FEVENCIM$ = FECHATEX$(FECHAVENCI(1))
    If CAPLAVENCI% > 1 Then FEVENCIM$ = FEVENCIM$ + " y ss"
    CAPARDOC% = CAPARDOC% + 1
    CODPARAM$(CAPARDOC%) = "FECH-VEN": DOCPARAM$(CAPARDOC%) = FEVENCIM$
    '
    Call PRINTDOC("FORFAC-" + LETRAFAC$) ' FACTURACION
    '
End Sub
'

Sub PRICRELIB()
    '
    FORIPRE$ = CONTROL$("", "FORFAC-" + LETRAFAC$)
    If FORIPRE$ = "" Then
       Call MENSERR(24, "Falta Definir Formulario de Impresión\para Facturación Tipo '" + LETRAFAC$ + "'.")
       Exit Sub
    End If
    '
    Screen.MousePointer = 11
    HOII% = HOY(HOS$)
    FECHDOC$ = FORCRELIDIP.Text9.TEXT
    NROPE& = Val(Mid$(FORCRELIDIP.Text11.TEXT, 4))
    NUMEDOC$ = TRIM$(Str$(NROPE&))
    While Len(NUMEDOC$) < 8
      NUMEDOC$ = "0" + NUMEDOC$
    Wend
    '
    PREFINUM$ = ATRIFORM$(FORIPRE$, "PREFIFAC")
    While Len(PREFINUM$) < 4
      PREFINUM$ = "0" + PREFINUM$
    Wend
    NUMEDOC$ = PREFINUM$ + "-" + NUMEDOC$
    NUREMITO$ = NUMEDOC$
    TIPODOC$ = "Nota de Crédito"
    '
    CODTABU1$(1) = "DES-MAT"
    CODTABU1$(2) = "NET-ITEM"
    CACOLTAB1% = 2
    '
    DESTIDOC% = NC%
    CAITAB1% = 0
    CAITCRU% = 0
    '
    SUMATOTA# = 0: J& = 0
    ITOT# = Val(FORCRELIDIP.Text10.TEXT)
    If LETRAFAC$ = "B" Then
       ITOT# = (Int(ITOT# * (100 + ALIVA) + 0.5)) / 100
    End If
    '
    TTXX$ = FORCRELIDIP.Text8.TEXT
    Call FRATEXTO(TTXX$, 48)
    '
    For KK% = 1 To CARETEX%
       CAITAB1% = CAITAB1% + 1
       If TRIM$(RETEX$(KK%)) = "" Then RETEX$(KK%) = "."
       TETABU1$(1, CAITAB1%) = RETEX$(KK%)
       TETABU1$(2, CAITAB1%) = Space$(10)
    Next KK%
    '
    ITOTX$ = Str$(ITOT#)
    TETABU1$(2, CAITAB1%) = ITOTX$
    '
    CPAG% = Val(FORCRELIDIP.TEXT21(1).TEXT)
    CONPA$ = DECONPA$(CPAG%)
    NOCOM$ = ""
    '
    CODPARAM$(1) = "CIVA-CLI": DOCPARAM$(1) = LETRAFAC$
    CODPARAM$(2) = "NRO-OCOM": DOCPARAM$(2) = NOCOM$
    CODPARAM$(3) = "COND-PAG": DOCPARAM$(3) = CONPA$
    CODPARAM$(4) = "IMP-NETO": DOCPARAM$(4) = FORCRELIDIP.Text10.TEXT
    CODPARAM$(5) = "IMP-IVA": DOCPARAM$(5) = FORCRELIDIP.Label18.Caption
    CODPARAM$(6) = "IMP-RNI": DOCPARAM$(6) = FORCRELIDIP.Label23.Caption
    CODPARAM$(7) = "IMP-TOTA": DOCPARAM$(7) = FORCRELIDIP.Label20.Caption
    '
    ITEXTO$ = NUMTEX$(Val(FORCRELIDIP.Label20.Caption)) + "."
    LOTEX% = Val(ATRIFORM$("", "LOIMPLET")): If LOTEX% < 20 Then LOTEX% = 20
    Call FRATEX(ITEXTO$, NRX1$, NRX2$, NRX3$, LOTEX%)
    IMPLET1$ = TRIM$(NRX1$) + " " + "***************************************************************************************************************"
    IMPLET2$ = TRIM$(NRX2$) + " " + "***************************************************************************************************************"
    '
    CODPARAM$(8) = "IMP-LET1": DOCPARAM$(8) = IMPLET1$
    CODPARAM$(9) = "IMP-LET2": DOCPARAM$(9) = IMPLET2$
    '
    CODPARAM$(10) = "FECH-EMI": DOCPARAM$(10) = FORCRELIDIP.Text9.TEXT
    CODPARAM$(11) = "NRO-REMI": DOCPARAM$(11) = ""
    CODPARAM$(12) = "NRO-AENT": DOCPARAM$(12) = ""
    CODPARAM$(13) = "TIPO-COM": DOCPARAM$(13) = UCase$(TIPODOC$)
    CODPARAM$(14) = "ITO-IVA": DOCPARAM$(14) = Str$(Val(DOCPARAM$(5)) + Val(DOCPARAM$(6)))
    CAPARDOC% = 14
    '
    FEVENCIM$ = ""
    CAPARDOC% = CAPARDOC% + 1
    CODPARAM$(CAPARDOC%) = "FECH-VEN": DOCPARAM$(CAPARDOC%) = FEVENCIM$
    '
    Call PRINTDOC("FORFAC-" + LETRAFAC$) ' FACTURACION
    '
End Sub
'

Sub PRICREDEV()
    '
    FORIPRE$ = CONTROL$("", "FORFAC-" + LETRAFAC$)
    If FORIPRE$ = "" Then
       Call MENSERR(24, "Falta Definir Formulario de Impresión\para Facturación Tipo '" + LETRAFAC$ + "'.")
       Exit Sub
    End If
    '
    Screen.MousePointer = 11
    HOII% = HOY(HOS$)
    FECHDOC$ = FORCREDEVDIP.Text9.TEXT
    NROPE& = Val(Mid$(FORCREDEVDIP.Text11.TEXT, 4))
    NUMEDOC$ = TRIM$(Str$(NROPE&))
    While Len(NUMEDOC$) < 8
      NUMEDOC$ = "0" + NUMEDOC$
    Wend
    '
    PREFINUM$ = ATRIFORM$(FORIPRE$, "PREFIFAC")
    While Len(PREFINUM$) < 4
      PREFINUM$ = "0" + PREFINUM$
    Wend
    NUMEDOC$ = PREFINUM$ + "-" + NUMEDOC$
    NUREMITO$ = NUMEDOC$
    TIPODOC$ = "Nota de Crédito"
    '
    CODTABU1$(1) = "COD-MAT"
    CODTABU1$(2) = "DES-MAT"
    CODTABU1$(3) = "CANTIDAD"
    CODTABU1$(4) = "PRE-NETO"
    CODTABU1$(5) = "NET-ITEM"
    CACOLTAB1% = 5
    '
    DESTIDOC% = NC%
    CAITAB1% = 0
    CAITCRU% = 0
    '
    SUMATOTA# = 0: J& = 0
    For U& = 1 To FORCREDEVDIP.List2.ListCount
       Z$ = FORCREDEVDIP.List2.List(U& - 1)
       CACO = Val(Left$(Z$, 4))
       If CACO >= 0.05 Then
          CODX$ = Mid$(Z$, 7, 15)
          CIIX% = CODINT%(CODX$)
          DECO$ = DESCRIT0$(CIIX%)
          PREU = Val(Mid$(Z$, 56, 9))
          ITOT# = Val(Mid$(Z$, 74, 10))
          '
          If LETRAFAC$ = "B" Then
             PREU = (Int(PREU * (100 + ALIVA) + 0.5)) / 100
             ITOT# = (Int(ITOT# * (100 + ALIVA) + 0.5)) / 100
          End If
          '
          CAPEX$ = Str$(CACO)
          PREUX$ = Str$(PREU)
          ITOTX$ = Str$(ITOT#)
          SUMATOTA# = SUMATOTA# + CDbl(Val(ITOTX$))
          '
          CAITAB1% = CAITAB1% + 1
          TETABU1$(1, CAITAB1%) = CODX$
          TETABU1$(2, CAITAB1%) = DECO$
          TETABU1$(3, CAITAB1%) = CAPEX$
          TETABU1$(4, CAITAB1%) = PREUX$
          TETABU1$(5, CAITAB1%) = ITOTX$
          '
       End If
    Next U&
    '
    CPAG% = 0
    CONPA$ = DECONPA$(CPAG%)
    NOCOM$ = ""
    '
    CODPARAM$(1) = "CIVA-CLI": DOCPARAM$(1) = LETRAFAC$
    CODPARAM$(2) = "NRO-OCOM": DOCPARAM$(2) = NOCOM$
    CODPARAM$(3) = "COND-PAG": DOCPARAM$(3) = CONPA$
    CODPARAM$(4) = "IMP-NETO": DOCPARAM$(4) = FORCREDEVDIP.Label14.Caption
    CODPARAM$(5) = "IMP-IVA": DOCPARAM$(5) = FORCREDEVDIP.Label18.Caption
    CODPARAM$(6) = "IMP-RNI": DOCPARAM$(6) = FORCREDEVDIP.Label23.Caption
    CODPARAM$(7) = "IMP-TOTA": DOCPARAM$(7) = FORCREDEVDIP.Label20.Caption
    '
    ITEXTO$ = NUMTEX$(Val(FORCREDEVDIP.Label20.Caption)) + "."
    LOTEX% = Val(ATRIFORM$("", "LOIMPLET")): If LOTEX% < 20 Then LOTEX% = 20
    Call FRATEX(ITEXTO$, NRX1$, NRX2$, NRX3$, LOTEX%)
    IMPLET1$ = TRIM$(NRX1$) + " " + "***************************************************************************************************************"
    IMPLET2$ = TRIM$(NRX2$) + " " + "***************************************************************************************************************"
    '
    CODPARAM$(8) = "IMP-LET1": DOCPARAM$(8) = IMPLET1$
    CODPARAM$(9) = "IMP-LET2": DOCPARAM$(9) = IMPLET2$
    '
    CODPARAM$(10) = "FECH-EMI": DOCPARAM$(10) = FORDELIDIP.Text9.TEXT
    CODPARAM$(11) = "NRO-REMI": DOCPARAM$(11) = ""
    CODPARAM$(12) = "NRO-AENT": DOCPARAM$(12) = ""
    CODPARAM$(13) = "TIPO-COM": DOCPARAM$(13) = UCase$(TIPODOC$)
    CODPARAM$(14) = "ITO-IVA": DOCPARAM$(14) = Str$(Val(DOCPARAM$(5)) + Val(DOCPARAM$(6)))
    CAPARDOC% = 14
    '
    FEVENCIM$ = ""
    CAPARDOC% = CAPARDOC% + 1
    CODPARAM$(CAPARDOC%) = "FECH-VEN": DOCPARAM$(CAPARDOC%) = FEVENCIM$
    '
    Call PRINTDOC("FORFAC-" + LETRAFAC$) ' FACTURACION
    '
End Sub
'

Sub PRIFACABO()
    '
    FORIPRE$ = CONTROL$("", "FORFAC-" + LETRAFAC$)
    If FORIPRE$ = "" Then
       Call MENSERR(24, "Falta Definir Formulario de Impresión\para Facturación Tipo '" + LETRAFAC$ + "'.")
       Exit Sub
    End If
    '
    Screen.MousePointer = 11
    HOII% = HOY(HOS$)
    FECHDOC$ = FORFACABO.Text9.TEXT
    NROPE& = Val(Mid$(FORFACABO.Text11.TEXT, 4))
    NUMEDOC$ = TRIM$(Str$(NROPE&))
    While Len(NUMEDOC$) < 8
      NUMEDOC$ = "0" + NUMEDOC$
    Wend
    '
    PREFINUM$ = ATRIFORM$(FORIPRE$, "PREFIFAC")
    While Len(PREFINUM$) < 4
      PREFINUM$ = "0" + PREFINUM$
    Wend
    NUMEDOC$ = PREFINUM$ + "-" + NUMEDOC$
    NUREMITO$ = NUMEDOC$
    TIPODOC$ = "Factura"
    '
    CODTABU1$(1) = "DES-MAT"
    CODTABU1$(2) = "NET-ITEM"
    CACOLTAB1% = 2
    '
    DESTIDOC% = NC%
    CAITAB1% = 0
    CAITCRU% = 0
    '
    SUMATOTA# = 0: J& = 0
    ITOT# = Val(FORFACABO.Label10.Caption)
    If LETRAFAC$ = "B" Then
       ITOT# = (Int(ITOT# * (100 + ALIVA) + 0.5)) / 100
    End If
    '
    TTXX$ = FORFACABO.Text8.TEXT
    Call FRATEXTO(TTXX$, 48)
    '
    For KK% = 1 To CARETEX%
       CAITAB1% = CAITAB1% + 1
       If TRIM$(RETEX$(KK%)) = "" Then RETEX$(KK%) = "."
       TETABU1$(1, CAITAB1%) = RETEX$(KK%)
       TETABU1$(2, CAITAB1%) = Space$(10)
    Next KK%
    '
    ITOTX$ = Str$(ITOT#)
    TETABU1$(2, CAITAB1%) = ITOTX$
    '
    CPAG% = Val(FORFACABO.TEXT21(1).TEXT)
    VEND% = Val(FORFACABO.TEXT21(2).TEXT)
    CONPA$ = DECONPA$(CPAG%)
    NOCOM$ = ""
    '
    REMITOS$ = "": VECREMI$ = ""
    For U& = 1 To ULTREG&("MOVISTO")
      MVS$ = REGLEIDO$("MOVISTO", U&)
      If CVI(Mid$(MVS$, 84, 2)) = NC% Then
        FEMO% = CVI(Mid$(MVS$, 1, 2))
        If FEMO% >= DESABONO% Then
          If FEMO% <= HASABONO% Then
            NUREMI& = Val(Mid$(MVS$, 26, 7))
            For KKJ% = 1 To Len(VECREMI$) Step 4
              If CVS(Mid$(VECREMI$, KKJ%, 4)) = NUREMI& Then GoTo 19
            Next KKJ%
            VECREMI$ = VECREMI$ + MKS$(NUREMI&)
          End If
        End If
      End If
19  Next U&
    '
    For KKJ% = 1 To Len(VECREMI$) Step 4
      If REMITOS$ <> "" Then REMITOS$ = REMITOS$ + ","
      REMITOS$ = REMITOS$ + TRIM$(Str$(CVS(Mid$(VECREMI$, KKJ%, 4))))
    Next KKJ%
    '
    CODPARAM$(1) = "CIVA-CLI": DOCPARAM$(1) = LETRAFAC$
    CODPARAM$(2) = "NRO-OCOM": DOCPARAM$(2) = ""
    CODPARAM$(3) = "COND-PAG": DOCPARAM$(3) = CONPA$
    CODPARAM$(4) = "IMP-NETO": DOCPARAM$(4) = FORFACABO.Label10.Caption
    CODPARAM$(5) = "IMP-IVA": DOCPARAM$(5) = FORFACABO.Label18.Caption
    CODPARAM$(6) = "IMP-RNI": DOCPARAM$(6) = FORFACABO.Label23.Caption
    CODPARAM$(7) = "IMP-TOTA": DOCPARAM$(7) = FORFACABO.Label20.Caption
    '
    ITEXTO$ = NUMTEX$(Val(FORFACABO.Label20.Caption)) + "."
    LOTEX% = Val(ATRIFORM$("", "LOIMPLET")): If LOTEX% < 20 Then LOTEX% = 20
    Call FRATEX(ITEXTO$, NRX1$, NRX2$, NRX3$, LOTEX%)
    IMPLET1$ = TRIM$(NRX1$) + " " + "***************************************************************************************************************"
    IMPLET2$ = TRIM$(NRX2$) + " " + "***************************************************************************************************************"
    '
    CODPARAM$(8) = "IMP-LET1": DOCPARAM$(8) = IMPLET1$
    CODPARAM$(9) = "IMP-LET2": DOCPARAM$(9) = IMPLET2$
    '
    CODPARAM$(10) = "FECH-EMI": DOCPARAM$(10) = FORFACABO.Text9.TEXT
    CODPARAM$(11) = "NRO-REMI": DOCPARAM$(11) = REMITOS$
    CODPARAM$(12) = "NRO-AENT": DOCPARAM$(12) = ""
    CODPARAM$(13) = "TIPO-COM": DOCPARAM$(13) = UCase$(TIPODOC$)
    CODPARAM$(14) = "ITO-IVA": DOCPARAM$(14) = Str$(Val(DOCPARAM$(5)) + Val(DOCPARAM$(6)))
    CODPARAM$(15) = "VENDEDOR": DOCPARAM$(15) = TRIM$(Str$(VEND%))
    CAPARDOC% = 15
    '
    FEVENCIM$ = ""
    If CAPLAVENCI% > 0 Then FEVENCIM$ = FECHATEX$(FECHAVENCI(1))
    If CAPLAVENCI% > 1 Then FEVENCIM$ = FEVENCIM$ + " y ss"
    CAPARDOC% = CAPARDOC% + 1
    CODPARAM$(CAPARDOC%) = "FECH-VEN": DOCPARAM$(CAPARDOC%) = FEVENCIM$
    '
    Call PRINTDOC("FORFAC-" + LETRAFAC$) ' FACTURACION
    '
End Sub
'

