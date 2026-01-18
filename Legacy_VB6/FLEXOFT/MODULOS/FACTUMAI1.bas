Attribute VB_Name = "FACTUMAI1"
Public PIVAFAC%                     ' posicion de IVA de factura
Dim LETRAFAC$                    ' letra de la factura
Dim FACIVA                       ' coeficiente precios por IVA
Dim VARIFAC%                     ' variante de facturación
Dim ALITOTA                      ' alicuota total de IVA
Public RENOIN%                      ' marca responsable no inscripto
'
Public FEFAC%                       ' fecha de factura
Dim TIDOC%                       ' tipo de documento
Public NUCREDESA&                ' Número N/Credito por Desacopio
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
Dim MODOCARCON%
'

Sub FACITEMI1(OKEY%)
   '
   OKEY% = 0
   '
   NC% = Val(FORFACMAI.Text1.TEXT)
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
   FORFACMAI.Text1.TEXT = Str$(NC%)
   FORFACMAI.Text2.TEXT = RASOCLI$(NC%)
   FORFACMAI.Text3.TEXT = DOMICLI$(NC%)
   FORFACMAI.Text4.TEXT = CPOSCLI$(NC%)
   FORFACMAI.Text5.TEXT = LOCACLI$(NC%)
   FORFACMAI.Text7.TEXT = TRIM$(CUITCLI$(NC%)) + " - " + POSIVAC$(NC%)
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
   FORFACMAI.Text6.TEXT = LETRAFAC$
   TIDOC% = 2                 ' tipo de documento = FACTURA
   '
   NROPRO& = 0
   Call NUFACTURA(NROPRO&)
   FORFACMAI.Text11.TEXT = "FC." + TRIM$(Str$(NROPRO&))
   '
   Call SELPEDPEN             ' seleccion pedidos pendientes
   If CAPESEL% < 1 Then Exit Sub
   '
   Call CARPEDSEL             ' carga de pedidos seleccionados
   PEDX$ = ENCAPED$
   Call REPLA(PEDX$, MKS$(NROPRO&), 1, 4)
   FECHDOC$ = FORFACMAI.Text9.TEXT
   FEFAC% = FECHANUM(FECHDOC$)
   If FEFAC% < 8600 Then FEFAC% = HOY(HO$)
   Call REPLA(PEDX$, MKI$(FEFAC%), 5, 2)
   '
10 ENCAPED$ = OBJPLA$("FACTURA-ENCMAI", PEDX$)
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
   Call CARCONFAC             ' carga de listas de facturacion
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
      If Control$("NUMEFACT", "NUMECORR") = "INDEPEN" Then
         NUMECORR% = (-1)
      End If
   End If
   '
   For MXC& = ULTREG&("FACTUR") To 1 Step -1
       F$ = REGLEIDO$("FACTUR", MXC&)
       TIX% = Asc(Mid$(F$, 13, 1)): VAX% = Asc(Mid$(F$, 14, 1))
       If TIX% = TIDOC% Or NUMECORR% = 1 Then
          If VAX% = VARIFAC% Then
             If CVI(Mid$(F$, 7, 2)) >= FLINUMERA% Then
                NROPRO& = 1 + CVS(Mid$(F$, 3, 4))
                Exit Sub
             End If
          End If
       End If
   Next MXC&
   '
10 NROPRO& = Val(InputBox$("Numero Primer Comprobante"))
   If NROPRO& < 1 Then GoTo 10
   '
End Sub
'

Sub SELPEDPEN()
   '
10 RPEDX$ = RECFILT$("PEDCLIEN", 3, MKI$(NC%))
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
      X$ = REGLEIDO$("PEDCLIEN", U&)
      If CVI(Mid$(X$, 7, 2)) <> NC% Then
         X$ = REGLEIDO$("PEDCLIEN", 1)
         Call GRAREG("PEDCLIEN", X$, 1)
         Call CIERRARCH("PEDCLIEN")
         GoTo 10
      End If
      '
      STATUS% = Asc(Mid$(X$, 300, 1))
      If STATUS% = 8 Then
         If CVI(Mid$(X$, 5, 2)) <= FECHANUM("28/08/00") Then
            STATUS% = 0
            Call REPLA(X$, Chr$(0), 300, 1)
            Call GRAREG("PEDCLIEN", X$, U&)
         End If
      End If
      '
      If STATUS% <= 1 Then
         '
         CACOPED = CVS(Mid$(X$, 513, 4))
         CAITPED = CVS(Mid$(X$, 645, 4))
         CACOENT = CVS(Mid$(X$, 512 + 77, 4))
         CAN = (CACOPED - CACOENT) * CAITPED
         '
         If CAN > 0.05 Then
            If TIDOC% <> 3 And VERISTO% = 1 Then
               CI% = CVI(Mid$(X$, 640 + 3, 2))
               If CI% > 0 Then
                  STACT = STOCKACT(CI%)
                  STDIS = STACT - CAN
                  If STACT < 0.05 Then
                     GoTo 19
                  End If
               End If
            End If
            '
            NROPED = CVS(Left$(X$, 4))
            If NROPED <> NROPEDA Then
               TTXX$ = "PC-" + TRIM$(CSTRING$(MKS$(NROPED), 3, 8, 0, 2))
               TTXX$ = AJUSTI$(TTXX$, 11)
               FEX = CVI(Mid$(X$, 5, 2))
               TTXX$ = TTXX$ + "  " + FECHATEX$(FEX)
               FORSELPED.LIPEDPE.AddItem TTXX$
            End If
            NROPEDA = NROPED
            CAPEDPE% = CAPEDPE% + 1
            FENTRE% = CVI(Mid$(X$, 512 + 63, 2))
            If FENTRE% < FINF Then FINF = FENTRE%
            If FENTRE% > FSUP Then FSUP = FENTRE%
         End If
      End If
19 Next UU&
   '
20 Screen.MousePointer = 1
   '
   If CAPEDPE% < 1 Then
      Call COMUNI("No hay Pedidos Pendientes\para este Cliente")
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
   REFANT$ = "": NROPEDA& = 0: ITIDX% = 0
   REFPED$ = ""
   '
   For UU& = 1 To Len(RPEDX$) Step 4
      '
      If FIN& > 0 Then
         Call TRACE(20, UU&, FIN&)
      End If
      '
      U& = CVS(Mid$(RPEDX$, UU&, 4))
      X$ = REGLEIDO$("PEDCLIEN", U&)
      If CVI(Mid$(X$, 7, 2)) <> NC% Then
         Call MENSERR(24, "Error de Indices - Consulte")
         Call FINAL("")
      End If
      '
      NROPED& = CVS(Left$(X$, 4))
      '
      For JI% = 1 To CAPESEL%
         If NUPEDSEL&(JI%) = NROPED& Then GoTo 20
      Next JI%
      GoTo 29
      '
20    If NROPED& <> NROPEDA& Then
         REFANT$ = ""
         ITANT% = 0
         If REFPED$ <> "" Then REFPED$ = REFPED$ + " - "
         REFPED$ = REFPED$ + TRIM$(Str$(NROPED&))
      End If
      NROPEDA& = NROPED&
      '
      ENCAPED$ = Left$(X$, 512)
      COPED$ = Mid$(X$, 513, 128)
      ITPED$ = Mid$(X$, 641, 64)
      '
      CACOPED = CVS(Mid$(X$, 513, 4))
      CAITPED = CVS(Mid$(X$, 645, 4))
      CACOENT = CVS(Mid$(X$, 512 + 77, 4))
      CAN = (CACOPED - CACOENT) * CAITPED
      '
      If CAN > 0.05 Then
         REFE$ = Mid$(COPED$, 5, 58)
         ITID% = CVI(Mid$(COPED$, 127, 2))
         If REFE$ <> REFANT$ Or ITID% <> ITIDA% Then
            ITIDX% = ITIDX% + 1
            Call REPLA(COPED$, MKS$(CACOPED - CACOENT), 1, 4)
            Call REPLA(COPED$, MKI$(ITIDX%), 127, 2)
            Call REGAPP("!CARCONPE", COPED$)
         End If
         REFANT$ = REFE$
         ITIDA% = ITID%
         '
         Call REPLA(ITPED$, MKI$(ITIDX%), 1, 2)
         Call REPLA(ITPED$, MKS$(U&), 61, 4)
         Call REGAPP("!CARITPED", ITPED$)
         '
      End If
      '
29 Next UU&
   '
End Sub
'

Sub CARCONFAC()
    '
200 If MODOCARCON% = 1 Then GoTo 210
    '
    ATRI$ = "/NC"
    ATRI$ = ATRI$ + "/TITUPAN=Ingreso Conjuntos a Itemizar"
    ATRI$ = ATRI$ + "/BORDESUP=W" + TRIM$(Str$(FORFACMAI.Top + FORFACMAI.Picture1.Top + 100))
    ATRI$ = ATRI$ + "/BORDEIZQ=W" + TRIM$(Str$(FORFACMAI.Left + FORFACMAI.Picture1.Left))
    '
    VTB% = VENTABU%("CARCONPE" + ATRI$)
    Call CIERRARCH("!CARCONPE")
    If VTB% < 0 Then
       '
       For J& = 1 To FORFACMAI.List1.ListCount
          Z$ = FORFACMAI.List1.List(J& - 1)
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
       Call SETULTREG("!CARCONPE", FORFACMAI.List1.ListCount)
       Exit Sub
       '
    End If
    '
210 REBLA$ = REGBLAN$("CARCONPE")
    FORFACMAI.List1.Clear
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
        FORFACMAI.List1.AddItem Z$
        TOPEDIDO# = TOPEDIDO# + ITOTCO#
        '
290 Next U&
    '
    If FORFACMAI.List1.ListCount > 0 Then
        FORFACMAI.List1.ListIndex = 0
        FORFACMAI.List1.SetFocus
    End If
    '
    ALITOTA = 0
    If VARIFAC% <> 3 Then ALITOTA = ALIVA
    FORFACMAI.Label14.Caption = CSTRING$(MKD$(TOPEDIDO#), 3, 11, 2, 2)
    If PIVAFAC% <> 5 Then FORFACMAI.Label18.Caption = CSTRING$(MKD$(TOPEDIDO# * ALIVA / 100), 3, 11, 2, 2)
    If RENOIN% = 1 Then
       FORFACMAI.Label21.Caption = CSTRING$(MKD$(TOPEDIDO# * ALIVARNI / 100), 3, 11, 2, 2)
       ALITOTA = ALIVA + ALIVARNI
    End If
    FORFACMAI.Label20.Caption = CSTRING$(MKD$(TOPEDIDO# * (1 + ALITOTA / 100)), 3, 11, 2, 2)
    '
End Sub
'
Sub CARITFAC1()
   '
   If FORFACMAI.List1.ListCount < 1 Then
      Exit Sub
   End If
   '
   ILIS1& = FORFACMAI.List1.ListIndex
   If ILIS1& < 0 Or ILIS1& > FORFACMAI.List1.ListCount - 1 Then
      FORFACMAI.List1.ListIndex = 0
   End If
   '
   U& = FORFACMAI.List1.ListIndex + 1
   X$ = REGLEIDO$("!CARCONPE", U&)
   ITID% = CVI(Mid$(X$, 127, 2))
   CACO = CVS(Left$(X$, 4))
   If ITID% < 1 Then
      Call MENSERR(24, "Imposible Seguir Proceso - Consulte")
      Call FINAL("")
   End If
   '
   FORFACMAI.List2.Clear
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
         FORFACMAI.List2.AddItem Z$
         '
      End If
   Next T&
   FORFACMAI.Label10.Caption = CSTRING$(MKD$(ITOTCO#), 3, 12, 2, 2)
   '
End Sub
'

Sub GRAFACMAI(OKEY%)
   '
   HOII% = HOY(HO$)
   '
   OKEY% = 0
   If ULTREG&("!CARCONPE") < 1 Then
       Call MENSERR(24, "Factura Vacía")
       Exit Sub
   End If
   '
   FECHDOC$ = FORFACMAI.Text9.TEXT
   FEFAC% = FECHANUM(FECHDOC$)
   If FEFAC% < 8600 Then
      Call MENSERR(24, "Fecha de Factura no Válida")
      Exit Sub
   End If
   FEX = FEFAC%
   FORFACMAI.Text9.TEXT = FECHATEX$(FEX)
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
19 Next U&
   '
   FORIPRE$ = Control$("", "FORFAC-" + LETRAFAC$)
   MAXITS$ = ATRIFORM$(FORIPRE$, "MAX-ITEM")
   If MAXITS$ = "HOJACONT" Then GoTo 20
   MAXIT% = Val(MAXITS$)
   
   CAITAB1% = 0
   For U& = 1 To ULTREG&("!CARCONPE")
      X$ = REGLEIDO$("!CARCONPE", U&)
      ITID% = CVI(Mid$(X$, 127, 2))
      CACO = CVS(Left$(X$, 4))
      If CACO >= 0.1 Then
         CAITAB1% = CAITAB1% + 2
         For T& = 1 To ULTREG&("!CARITPED")
            Y$ = REGLEIDO$("!CARITPED", T&)
            If CVI(Left$(Y$, 2)) = ITID% Then
               CAITAB1% = CAITAB1% + 1
            End If
         Next T&
      End If
   Next U&
   '
   If CAITAB1% > MAXIT% Then
      Call MENSERR(24, "Imposible Registrar esta Factura.\Excede Máximo de Renglones Admisible\en el Formulario de Impresión.")
      Exit Sub
   End If
   '
20 Call BLOQARCH("!SELACOP")
   Call BLANARCH("!SELACOP")
   TODESACO# = 0
   '
   RFF$ = RECFILT$("FACOPIOS", 1, MKI$(NC%))
   If Len(RFF$) < 4 Then
      Call MENSERR(24, "No aparecen Acopios Pendientes\para la Cuenta Elegida.")
      GoTo 30
   End If
   '
   NEFA# = Val(FORFACMAI.Label14.Caption)
   For UP& = 1 To Len(RFF$) Step 4
      U& = CVS(Mid$(RFF$, UP&, 4))
      RACOPI$ = REGLEIDO$("FACOPIOS", U&)
      If CVI(Left$(RACOPI$, 2)) <> NC% Then
         RACOPI$ = REGLEIDO$("FACOPIOS", 1)
         Call GRAREG("FACOPIOS", RACOPI$, 1)
         Call CIERRARCH("FACOPIOS")
         GoTo 20
      End If
      '
      If CVD(Mid$(RACOPI$, 73, 8)) < 0.005 Then
         NUFACOP& = CVS(Mid$(RACOPI$, 3, 4))
         FEFACOP% = CVI(Mid$(RACOPI$, 7, 2))
         INEFACO# = CVD(Mid$(RACOPI$, 15, 8))
         IPEFACO# = CVD(Mid$(RACOPI$, 65, 8))
         IADESAC# = IPEFACO#
         If TODESACO# + IADESAC# > 0.3 * NEFA# Then
            IADESAC# = 0.3 * NEFA# - TODESACO#
         End If
         TODESACO# = TODESACO# + IADESAC#
         RESEL$ = REGBLAN$("SELACOP")
         Call REPLA(RESEL$, MKS$(NUFACOP&), 1, 4)
         Call REPLA(RESEL$, MKI$(FEFACOP%), 5, 2)
         Call REPLA(RESEL$, MKD$(INEFACO#), 7, 8)
         Call REPLA(RESEL$, MKD$(IPEFACO#), 15, 8)
         Call REPLA(RESEL$, MKD$(IADESAC#), 23, 8)
         Call REPLA(RESEL$, MKS$(U&), 31, 4)
         Call REGAPP("!SELACOP", RESEL$)
      End If
      '
   Next UP&
   Call CIERRARCH("!SELACOP")
   If ULTREG&("!SELACOP") < 1 Then
      Call MENSERR(24, "No aparecen Acopios Pendientes\para la Cuenta Elegida.")
      GoTo 30
   End If
   '
   VTB% = VENTABU%("SELACOP")
   If VTB% < 0 Then
      Call MENSERR(24, "Factura Cancelada")
      Exit Sub
   End If
   '
   TODESACO# = 0
   For U& = 1 To ULTREG&("!SELACOP")
      RESEL$ = REGLEIDO$("!SELACOP", U&)
      If CVS(Mid$(RESEL$, 31, 4)) < 1 Then
         Call MENSERR(24, "Ingreso no Válido")
         GoTo 20
      End If
      If CVD(Mid$(RESEL$, 23, 8)) < 0 Then
         Call MENSERR(24, "Ingreso no Válido")
         GoTo 20
      End If
      If CVD(Mid$(RESEL$, 23, 8)) > CVD(Mid$(RESEL$, 15, 8)) Then
         Call MENSERR(24, "Ingreso no Válido")
         GoTo 20
      End If
      IADESAC# = CVD(Mid$(RESEL$, 23, 8))
      TODESACO# = TODESACO# + IADESAC#
   Next U&
   If TODESACO# > NEFA# Then
      Call MENSERR(24, "Los Desacopios Superan\el Importe Neto de Factura.")
      GoTo 20
   End If
   Call CIERRARCH("!SELACOP")
   '
30 Call CIERRARCH("FACTUR")
   Call CIERRARCH("SVD202")
   Call CIERRARCH("MOVISTO")
   Call CIERRARCH("DEUDOR2")
   '
   Call BLOQARCH("FACTUR")
   Call BLOQARCH("SVD202")
   Call BLOQARCH("DEUDOR2")
   Call BLOQARCH("MOVISTO")
   Call BLOQARCH("BKMOVIST")
   Call BLOQARCH("PEDCLIEN")
   Call BLOQARCH("FACOPIOS")
   '
   NROPRO& = 0
   Call NUFACTURA(NROPRO&)
   If Val(Mid$(FORFACMAI.Text11.TEXT, 4)) <> NROPRO& Then
      Call COMUNI("Este Documento Ha Sido Re-Numerado\Como FC." + TRIM$(Str$(NROPRO&)) + ".\  \Atención al Colocar Formulario de Impresión.")
   End If
   FORFACMAI.Text11.TEXT = "FC." + TRIM$(Str$(NROPRO&))
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
   NEFA# = Val(FORFACMAI.Label14.Caption)
   NIVA# = Val(FORFACMAI.Label18.Caption)
   NNIN# = Val(FORFACMAI.Label21.Caption)
   TOTA# = Val(FORFACMAI.Label20.Caption)
   IINT# = 0
   PERC# = 0
   '
   TOTA1# = 0
   If TODESACO# >= 0.005 Then
      NEFA1# = TODESACO#
      NIVA1# = 0
      If PIVAFAC% <> 5 Then NIVA1# = Val(CSTRING$(MKD$(TODESACO# * ALIVA / 100), 3, 11, 2, 2))
      NNIN1# = 0
      If RENOIN% = 1 Then
         NNIN1# = Val(CSTRING$(MKD$(TODESACO# * ALIVARNI / 100), 3, 11, 2, 2))
      End If
      TOTA1# = NEFA1# + NIVA1# + NNIN1#
   End If
   If TOTA# - TOTA1# + SALDCLI#(NC%) + CHEPEACRE#(NC%) > LIMICRE(NC%) Then
      Call MENSERR(24, "Imposible Registrar esta Factura.\  \Supera el Límite de Crédito Asignado\para este Cliente.")
      Call CIERRARCH("FACTUR")
      Call CIERRARCH("SVD202")
      Call CIERRARCH("DEUDOR2")
      Call CIERRARCH("MOVISTO")
      Call CIERRARCH("BKMOVIST")
      Call CIERRARCH("PEDCLIEN")
      Call CIERRARCH("FACOPIOS")
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
   NUCREDESA& = 0
   TOTA1# = 0
   If TODESACO# >= 0.005 Then
      '
      Call BLOQARCH("FACTUR")
      TIDOCA% = TIDOC%
      NROPROA& = NROPRO&
      NROPRO& = 0: TIDOC% = 6
      Call NUFACTURA(NROPRO&)
      NUCREDESA& = NROPRO&
      '
      X$ = REGBLAN$("FACTUR")
      Call REPLA(X$, MKI$(NC%), 1, 2)
      Call REPLA(X$, MKS$(NUCREDESA&), 3, 4)
      Call REPLA(X$, MKI$(FEFAC%), 7, 2)
      Call REPLA(X$, MKI$(FEFAC%), 9, 2)
      Call REPLA(X$, MKI$(0), 11, 2)
      Call REPLA(X$, Chr$(6), 13, 1)
      Call REPLA(X$, Chr$(VARIFAC%), 14, 1)
      '
      NEFA1# = TODESACO#
      NIVA1# = 0
      If PIVAFAC% <> 5 Then NIVA1# = Val(CSTRING$(MKD$(TODESACO# * ALIVA / 100), 3, 11, 2, 2))
      NNIN1# = 0
      If RENOIN% = 1 Then
         NNIN1# = Val(CSTRING$(MKD$(TODESACO# * ALIVARNI / 100), 3, 11, 2, 2))
      End If
      TOTA1# = NEFA1# + NIVA1# + NNIN1#
      IINT1# = 0
      PERC1# = 0
      '
      Call REPLA(X$, MKD$(NEFA1#), 15, 8)
      Call REPLA(X$, MKD$(NIVA1#), 23, 8)
      Call REPLA(X$, MKD$(NNIN1#), 31, 8)
      Call REPLA(X$, MKD$(TOTA1#), 39, 8)
      Call REPLA(X$, MKI$(VENDE%), 47, 8)
      Call REPLA(X$, MKD$(IINT1#), 49, 8)
      Call REPLA(X$, MKD$(PERC1#), 57, 8)
      '
      Call REGAPP("FACTUR", X$)
      '
      TIDOC% = TIDOCA%
      NROPRO& = NROPROA&
      '
   End If
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
   For U& = 1 To ULTREG&("!CARCONPE")
      X$ = REGLEIDO$("!CARCONPE", U&)
      '
      CACO = CVS(Left$(X$, 4))
      ITID% = CVI(Mid$(X$, 127, 2))
      '
      For T& = 1 To ULTREG&("!CARITPED")
         Y$ = REGLEIDO$("!CARITPED", T&)
         If CVI(Left$(Y$, 2)) = ITID% Then
            CAN = CACO * CVS(Mid$(Y$, 5, 4))
            If CAN > 0.05 Then
               CIIT% = CVI(Mid$(Y$, 3, 2))
               PREU = CVS(Mid$(Y$, 17, 4))
               ITOT# = CVD(Mid$(Y$, 21, 8))
               Call MOVISTO(FEFAC%, CIIT%, "", "FC", "RT." + TRIM$(Str$(NROPRO&)), "FC." + TRIM$(Str$(NROPRO&)), RASOCLI$(NC%), CDbl(PREU), "$", NC%, 1, (-1) * CAN)
            End If
         End If
      Next T&
   Next U&
   Call CIERRARCH("MOVISTO")
   Call CIERRARCH("BKMOVIST")
   '
   For U& = 1 To ULTREG&("!CARCONPE")
      X$ = REGLEIDO$("!CARCONPE", U&)
      '
      CACO = CVS(Left$(X$, 4))
      ITID% = CVI(Mid$(X$, 127, 2))
      '
      For T& = 1 To ULTREG&("!CARITPED")
         Y$ = REGLEIDO$("!CARITPED", T&)
         If CVI(Left$(Y$, 2)) = ITID% Then
            If CACO > 0.05 Then
               CIIT% = CVI(Mid$(Y$, 3, 2))
               PREU = CVS(Mid$(Y$, 17, 4))
               ITOT# = CVD(Mid$(Y$, 21, 8))
               REIT& = CVS(Mid$(Y$, 61, 4))
               If REIT& > 0 Then
                  If REIT& <= ULTREG&("PEDCLIEN") Then
                     Z$ = REGLEIDO$("PEDCLIEN", REIT&)
                     CACOENT = CVS(Mid$(Z$, 512 + 77, 4)) + CACO
                     Call REPLA(Z$, MKS$(CACOENT), 512 + 77, 4)
                     Call GRAREG("PEDCLIEN", Z$, REIT&)
                  End If
               End If
            End If
         End If
      Next T&
   Next U&
   Call CIERRARCH("PEDCLIEN")
   '
   If TODESACO# > 0 Then
      For U& = 1 To ULTREG&("!SELACOP")
         RESEL$ = REGLEIDO$("!SELACOP", U&)
         IDESACO# = CVD(Mid$(RESEL$, 23, 8))
         If IDESACO# >= 0.005 Then
            REFAC& = CVS(Mid$(RESEL$, 31, 4))
            REFACO$ = REGLEIDO$("FACOPIOS", REFAC&)
            IPEDESA# = CVD(Mid$(REFACO$, 65, 8))
            Call REPLA(REFACO$, MKD$(IDESACO#), 73, 8)
            Call REPLA(REFACO$, MKI$(FEFAC%), 81, 2)
            Call REPLA(REFACO$, MKS$(NROPRO&), 83, 4)
            Call REPLA(REFACO$, MKS$(NROPRO& + 1), 87, 4)
            Call GRAREG("FACOPIOS", REFACO$, REFAC&)
            '
            If IDESACO# < IPEDESA# Then
               Call REPLA(REFACO$, MKD$(IPEDESA# - IDESACO#), 65, 8)
               Call REPLA(REFACO$, String$(56, 0), 73, 56)
               Call REGAPP("FACOPIOS", REFACO$)
            End If
            '
         End If
      Next U&
   End If
   Call CIERRARCH("FACOPIOS")
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
   NC% = Val(FORDELIB.Text1.TEXT)
   If NC% < 1 Or ECOARCH$("DEUDOR1", MKI$(NC%)) = "" Then
       Call MENSERR(24, "Numero de Cuenta no Válido o Inexistente")
       FORDELIB.Text1.SetFocus
       Exit Sub
   End If
   '
   If FORDELIB.Text8.TEXT = "" Then
       Call MENSERR(24, "Falta Ingresar Detalle de Nota de Debito")
       FORDELIB.Text8.SetFocus
       Exit Sub
   End If
   '
   FECHDOC$ = FORDELIB.Text9.TEXT
   FEFAC% = FECHANUM(FECHDOC$)
   If FEFAC% < 8600 Then
      Call MENSERR(24, "Fecha de N.Débito no Válida")
      FORDELIB.Text9.SetFocus
      Exit Sub
   End If
   FEX = FEFAC%
   FORDELIB.Text9.TEXT = FECHATEX$(FEX)
   '
   CPAG% = Val(FORDELIB.TEXT21(1).TEXT)
   If ECOARCH$("CONPAG", MKI$(CPAG%)) = "" Then
      Call MENSERR(24, "Condición de Pago no Válida")
      FORDELIB.TEXT21(1).SetFocus
      Exit Sub
   End If
   CONPAG$ = DECONPA$(CPAG%)
   
   VEND% = Val(FORDELIB.TEXT21(1).TEXT)
   VENDE$ = DEVENDE$(VEND%)
   
   Call CIERRARCH("FACTUR")
   Call CIERRARCH("SVD202")
   Call CIERRARCH("DEUDOR2")
   '
   Call BLOQARCH("FACTUR")
   Call BLOQARCH("SVD202")
   Call BLOQARCH("DEUDOR2")
   '
   NROPRO& = 0
   Call NUFACTURA(NROPRO&)
   If Val(Mid$(FORDELIB.Text11.TEXT, 4)) <> NROPRO& Then
      Call COMUNI("Este Documento Ha Sido Re-Numerado\Como ND." + TRIM$(Str$(NROPRO&)) + ".\  \Atención al Colocar Formulario de Impresión.")
   End If
   FORDELIB.Text11.TEXT = "ND." + TRIM$(Str$(NROPRO&))
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
   NEFA# = Val(FORDELIB.Text10.TEXT)
   NIVA# = Val(FORDELIB.Label18.Caption)
   NNIN# = Val(FORDELIB.Label23.Caption)
   TOTA# = Val(FORDELIB.Label20.Caption)
   IINT# = 0
   PERC# = 0
   '
   If NEFA# < 0.01 Then
      Call MENSERR(24, "Faltan Importes")
      FORDELIB.Text10.SetFocus
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
   NC% = Val(FORCRELIB.Text1.TEXT)
   If NC% < 1 Or ECOARCH$("DEUDOR1", MKI$(NC%)) = "" Then
       Call MENSERR(24, "Numero de Cuenta no Válido o Inexistente")
       FORCRELIB.Text1.SetFocus
       Exit Sub
   End If
   '
   If FORCRELIB.Text8.TEXT = "" Then
       Call MENSERR(24, "Falta Ingresar Detalle de Nota de Crédito")
       FORCRELIB.Text8.SetFocus
       Exit Sub
   End If
   '
   FECHDOC$ = FORCRELIB.Text9.TEXT
   FEFAC% = FECHANUM(FECHDOC$)
   If FEFAC% < 8600 Then
      Call MENSERR(24, "Fecha de N.Crédito no Válida")
      FORCRELIB.Text9.SetFocus
      Exit Sub
   End If
   FEX = FEFAC%
   FORCRELIB.Text9.TEXT = FECHATEX$(FEX)
   '
   CPAG% = 0
   CONPAG$ = DECONPA$(CPAG%)
   
   VEND% = Val(FORCRELIB.TEXT21(1).TEXT)
   VENDE$ = DEVENDE$(VEND%)
   
   Call CIERRARCH("FACTUR")
   Call CIERRARCH("SVD202")
   Call CIERRARCH("DEUDOR2")
   '
   Call BLOQARCH("FACTUR")
   Call BLOQARCH("SVD202")
   Call BLOQARCH("DEUDOR2")
   '
   NROPRO& = 0
   Call NUFACTURA(NROPRO&)
   If Val(Mid$(FORCRELIB.Text11.TEXT, 4)) <> NROPRO& Then
      Call COMUNI("Este Documento Ha Sido Re-Numerado\Como NC." + TRIM$(Str$(NROPRO&)) + ".\  \Atención al Colocar Formulario de Impresión.")
   End If
   FORCRELIB.Text11.TEXT = "NC." + TRIM$(Str$(NROPRO&))
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
   NEFA# = Val(FORCRELIB.Text10.TEXT)
   NIVA# = Val(FORCRELIB.Label18.Caption)
   NNIN# = Val(FORCRELIB.Label23.Caption)
   TOTA# = Val(FORCRELIB.Label20.Caption)
   IINT# = 0
   PERC# = 0
   '
   If NEFA# < 0.01 Then
      Call MENSERR(24, "Faltan Importes")
      FORCRELIB.Text10.SetFocus
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
   If FORCREDEV.List2.ListCount < 1 Then
       Call MENSERR(24, "Nota de Crédito Vacía")
       Exit Sub
   End If
   '
   FECHDOC$ = FORCREDEV.Text9.TEXT
   FEFAC% = FECHANUM(FECHDOC$)
   If FEFAC% < 8600 Then
      Call MENSERR(24, "Fecha de N.Crédito no Válida")
      Exit Sub
   End If
   FEX = FEFAC%
   FORCREDEV.Text9.TEXT = FECHATEX$(FEX)
   '
   Call CIERRARCH("FACTUR")
   Call CIERRARCH("SVD202")
   Call CIERRARCH("MOVISTO")
   Call CIERRARCH("DEUDOR2")
   '
   Call BLOQARCH("FACTUR")
   Call BLOQARCH("SVD202")
   Call BLOQARCH("DEUDOR2")
   Call BLOQARCH("MOVISTO")
   Call BLOQARCH("BKMOVIST")
   '
   NROPRO& = 0
   Call NUFACTURA(NROPRO&)
   If Val(Mid$(FORCREDEV.Text11.TEXT, 4)) <> NROPRO& Then
      Call COMUNI("Este Documento Ha Sido Re-Numerado\Como NC." + TRIM$(Str$(NROPRO&)) + ".\  \Atención al Colocar Formulario de Impresión.")
   End If
   FORCREDEV.Text11.TEXT = "NC." + TRIM$(Str$(NROPRO&))
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
   NEFA# = Val(FORCREDEV.Label14.Caption)
   NIVA# = Val(FORCREDEV.Label18.Caption)
   NNIN# = Val(FORCREDEV.Label23.Caption)
   TOTA# = Val(FORCREDEV.Label20.Caption)
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
   For U& = 1 To FORCREDEV.List2.ListCount
      Z$ = FORCREDEV.List2.List(U& - 1)
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
Sub GRAFACOPIO(OKEY%)
   '
   HOII% = HOY(HO$)
   OKEY% = 0
   '
   NC% = Val(FORFACOPIO.Text1.TEXT)
   If NC% < 1 Or ECOARCH$("DEUDOR1", MKI$(NC%)) = "" Then
       Call MENSERR(24, "Numero de Cuenta no Válido o Inexistente")
       FORFACOPIO.Text1.SetFocus
       Exit Sub
   End If
   '
   If FORFACOPIO.Text8.TEXT = "" Then
       Call MENSERR(24, "Falta Ingresar Detalle de Factura de Acopio")
       FORFACOPIO.Text8.SetFocus
       Exit Sub
   End If
   '
   FECHDOC$ = FORFACOPIO.Text9.TEXT
   FEFAC% = FECHANUM(FECHDOC$)
   If FEFAC% < 8600 Then
      Call MENSERR(24, "Fecha de Factura de Acopio no Válida")
      FORFACOPIO.Text9.SetFocus
      Exit Sub
   End If
   FEX = FEFAC%
   FORFACOPIO.Text9.TEXT = FECHATEX$(FEX)
   '
   CPAG% = Val(FORFACOPIO.TEXT21(1).TEXT)
   If ECOARCH$("CONPAG", MKI$(CPAG%)) = "" Then
      Call MENSERR(24, "Condición de Pago no Válida")
      FORFACOPIO.TEXT21(1).SetFocus
      Exit Sub
   End If
   CONPAG$ = DECONPA$(CPAG%)
   
   VEND% = Val(FORFACOPIO.TEXT21(1).TEXT)
   VENDE$ = DEVENDE$(VEND%)
   
   Call CIERRARCH("FACTUR")
   Call CIERRARCH("SVD202")
   Call CIERRARCH("DEUDOR2")
   '
   Call BLOQARCH("FACTUR")
   Call BLOQARCH("SVD202")
   Call BLOQARCH("DEUDOR2")
   '
   NROPRO& = 0
   Call NUFACTURA(NROPRO&)
   If Val(Mid$(FORFACOPIO.Text11.TEXT, 4)) <> NROPRO& Then
      Call COMUNI("Este Documento Ha Sido Re-Numerado\Como FC." + TRIM$(Str$(NROPRO&)) + ".\  \Atención al Colocar Formulario de Impresión.")
   End If
   FORFACOPIO.Text11.TEXT = "FC." + TRIM$(Str$(NROPRO&))
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
   NEFA# = Val(FORFACOPIO.Text10.TEXT)
   NIVA# = Val(FORFACOPIO.Label18.Caption)
   NNIN# = Val(FORFACOPIO.Label23.Caption)
   TOTA# = Val(FORFACOPIO.Label20.Caption)
   IINT# = 0
   PERC# = 0
   '
   If NEFA# < 0.01 Then
      Call MENSERR(24, "Faltan Importes")
      FORFACOPIO.Text10.SetFocus
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
   RECOPIO$ = String$(128, 0)
   Call REPLA(RECOPIO$, X$, 1, 64)
   Call REPLA(RECOPIO$, MKD$(NEFA#), 65, 8)
   '
   Call REGAPP("FACTUR", X$)
   Call CIERRARCH("FACTUR")
   '
   Call REGAPP("FACOPIOS", RECOPIO$)
   Call CIERRARCH("FACOPIOS")
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
Sub PRIFACMAI()
    '
    TTXX$ = Mid$(ENCAPED$, 301, 60): Call GRAREG("!OBSERVOF", TTXX$, 1)
    TTXX$ = Mid$(ENCAPED$, 361, 60): Call GRAREG("!OBSERVOF", TTXX$, 2)
    TTXX$ = Mid$(ENCAPED$, 421, 60): Call GRAREG("!OBSERVOF", TTXX$, 3)
    Call SETULTREG("!OBSERVOF", 3)
    Call CIERRARCH("!OBSERVOF")
    '
    FORIPRE$ = Control$("", "FORFAC-" + LETRAFAC$)
    If FORIPRE$ = "" Then
       Call MENSERR(24, "Falta Definir Formulario de Impresión\para Facturación Tipo '" + LETRAFAC$ + "'.")
       Exit Sub
    End If
    '
    Screen.MousePointer = 11
    HOII% = HOY(HOS$)
    FECHDOC$ = FORFACMAI.Text9.TEXT
    NROPE& = Val(Mid$(FORFACMAI.Text11.TEXT, 4))
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
       If LETRAFAC$ = "B" Then
          PREU = (Int(PREU * (100 + ALIVA) + 0.5)) / 100
          ITOT# = (Int(ITOT# * (100 + ALIVA) + 0.5)) / 100
       End If
       '
       FENTX$ = FECHATEX$(FEX)
       CAPEX$ = TRIM$(CSTRING$(MKS$(CACO), 3, 10, 1, 2))
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
              CAPEX$ = "(" + TRIM$(CSTRING$(MKS$(CACO * CAIT), 3, 10, 0, 2)) + ")"
              '
              CAITAB1% = CAITAB1% + 1
              TETABU1$(1, CAITAB1%) = ""
              TETABU1$(2, CAITAB1%) = Z$
              TETABU1$(3, CAITAB1%) = CAPEX$
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
     CPAG% = Asc(Mid$(ENCAPED$, 298, 1))
     CONPA$ = DECONPA$(CPAG%)
     NOCOM$ = Mid$(ENCAPED$, 137, 18)
     '
     CODPARAM$(1) = "CIVA-CLI": DOCPARAM$(1) = LETRAFAC$
     CODPARAM$(2) = "NRO-OCOM": DOCPARAM$(2) = NOCOM$
     CODPARAM$(3) = "COND-PAG": DOCPARAM$(3) = CONPA$
     CODPARAM$(4) = "IMP-NETO": DOCPARAM$(4) = FORFACMAI.Label14.Caption
     CODPARAM$(5) = "IMP-IVA": DOCPARAM$(5) = FORFACMAI.Label18.Caption
     CODPARAM$(6) = "IMP-RNI": DOCPARAM$(6) = FORFACMAI.Label21.Caption
     CODPARAM$(7) = "IMP-TOTA": DOCPARAM$(7) = FORFACMAI.Label20.Caption
     '
     ITEXTO$ = NUMTEX$(Val(FORFACMAI.Label20.Caption)) + "."
     LOTEX% = Val(ATRIFORM$("", "LOIMPLET")): If LOTEX% < 20 Then LOTEX% = 20
     Call FRATEX(ITEXTO$, NRX1$, NRX2$, NRX3$, LOTEX%)
     IMPLET1$ = TRIM$(NRX1$) + " " + "***************************************************************************************************************"
     IMPLET2$ = TRIM$(NRX2$) + " " + "***************************************************************************************************************"
     '
     CODPARAM$(8) = "IMP-LET1": DOCPARAM$(8) = IMPLET1$
     CODPARAM$(9) = "IMP-LET2": DOCPARAM$(9) = IMPLET2$
     '
     CODPARAM$(10) = "FECH-EMI": DOCPARAM$(10) = FORFACMAI.Text9.TEXT
     CODPARAM$(11) = "NRO-REMI": DOCPARAM$(11) = NUREMITO$
     CODPARAM$(12) = "NRO-AENT": DOCPARAM$(12) = REFPED$
     CODPARAM$(13) = "TIPO-COM": DOCPARAM$(13) = UCase$(TIPODOC$)
     CODPARAM$(14) = "ITO-IVA": DOCPARAM$(14) = Str$(Val(DOCPARAM$(5)) + Val(DOCPARAM$(6)))
     '
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
     TIPODOC$ = "Remito"
     FORIPRE$ = Control$("", "FOREMFAC")
     If FORIPRE$ = "" Then
       Call MENSERR(24, "Falta Definir Formulario\de Impresión de Remito.")
       Exit Sub
     End If
     '
     CODPARAM$(1) = "CIVA-CLI": DOCPARAM$(1) = "X"
     Call PRINTDOC("FOREMFAC")   ' REMITO
     '
     'IMPRIMIR NOTA DE CREDITO POR DESACOPIO
     TODESACO# = 0
     For U& = 1 To ULTREG&("!SELACOP")
        RESEL$ = REGLEIDO$("!SELACOP", U&)
        IADESAC# = CVD(Mid$(RESEL$, 23, 8))
        TODESACO# = TODESACO# + IADESAC#
     Next U&
     '
     If TODESACO# >= 0.005 Then
        FECHDOC$ = FORFACMAI.Text9.TEXT
        NROPE& = NUCREDESA&
        NUMEDOC$ = TRIM$(Str$(NROPE&))
        While Len(NUMEDOC$) < 8
           NUMEDOC$ = "0" + NUMEDOC$
        Wend
        '
        NUMEDOC$ = PREFINUM$ + "-" + NUMEDOC$
        NUREMITO$ = ""
        TIPODOC$ = "Nota de Credito"
        '
        CAITAB1% = 0
        CAITCRU% = 0
        '
        SUMATOTA# = 0: J& = 0
        CAITAB1% = CAITAB1% + 1
        TETABU1$(1, CAITAB1%) = ""
        TETABU1$(2, CAITAB1%) = "Por Desacopio Factura " + NUFACORI$ + " del " + FORFACMAI.Text9.TEXT
        TETABU1$(3, CAITAB1%) = ""
        TETABU1$(4, CAITAB1%) = ""
        TETABU1$(5, CAITAB1%) = ""
        TETABU1$(6, CAITAB1%) = ""
        '
        CAITAB1% = CAITAB1% + 1
        TETABU1$(1, CAITAB1%) = ""
        TETABU1$(2, CAITAB1%) = "según el siguiente detalle:"
        TETABU1$(3, CAITAB1%) = ""
        TETABU1$(4, CAITAB1%) = ""
        TETABU1$(5, CAITAB1%) = ""
        TETABU1$(6, CAITAB1%) = ""
        '
        CAITAB1% = CAITAB1% + 1
        TETABU1$(1, CAITAB1%) = ""
        TETABU1$(2, CAITAB1%) = "."
        TETABU1$(3, CAITAB1%) = ""
        TETABU1$(4, CAITAB1%) = ""
        TETABU1$(5, CAITAB1%) = ""
        TETABU1$(6, CAITAB1%) = ""
        '
        TODESACO# = 0
        For U& = 1 To ULTREG&("!SELACOP")
           RESEL$ = REGLEIDO$("!SELACOP", U&)
           IADESAC# = CVD(Mid$(RESEL$, 23, 8))
           If IADESAC# > 0 Then
              Z$ = Space$(80)
              NUFACO& = CVS(Mid$(RESEL$, 1, 4))
              FEX = CVI(Mid$(RESEL$, 5, 2))
              Z$ = "- " + "Factura Nro." + Str$(NUFACO&) + " del " + FECHATEX$(FEX)
              PREU = 0
              ITOT# = IADESAC#
              TODESACO# = TODESACO# + IADESAC#
              '
              If LETRAFAC$ = "B" Then
                 PREU = (Int(PREU * (100 + ALIVA) + 0.5)) / 100
                 ITOT# = (Int(ITOT# * (100 + ALIVA) + 0.5)) / 100
              End If
              '
              FENTX$ = ""
              CAPEX$ = ""
              PREUX$ = Str$(PREU)
              ITOTX$ = Str$(ITOT#)
              SUMATOTA# = SUMATOTA# + CDbl(Val(ITOTX$))
              '
              CAITAB1% = CAITAB1% + 1
              TETABU1$(1, CAITAB1%) = ""
              TETABU1$(2, CAITAB1%) = Z$
              TETABU1$(3, CAITAB1%) = CAPEX$
              TETABU1$(4, CAITAB1%) = PREUX$
              TETABU1$(5, CAITAB1%) = ITOTX$
              TETABU1$(6, CAITAB1%) = FENTX$
              '
           End If
           '
        Next U&
        '
        CPAG% = 0
        CONPA$ = "CONTADO"
        NOCOM$ = ""
        '
        NEFA1# = TODESACO#
        NIVA1# = 0
        If PIVAFAC% <> 5 Then NIVA1# = Val(CSTRING$(MKD$(TODESACO# * ALIVA / 100), 3, 11, 2, 2))
        NNIN1# = 0
        If RENOIN% = 1 Then
           NNIN1# = Val(CSTRING$(MKD$(TODESACO# * ALIVARNI / 100), 3, 11, 2, 2))
        End If
        TOTA1# = NEFA1# + NIVA1# + NNIN1#
        IINT1# = 0
        PERC1# = 0
        '
        CODPARAM$(1) = "CIVA-CLI": DOCPARAM$(1) = LETRAFAC$
        CODPARAM$(2) = "NRO-OCOM": DOCPARAM$(2) = ""
        CODPARAM$(3) = "COND-PAG": DOCPARAM$(3) = CONPA$
        CODPARAM$(4) = "IMP-NETO": DOCPARAM$(4) = Str$(NEFA1#)
        CODPARAM$(5) = "IMP-IVA": DOCPARAM$(5) = Str$(NIVA1#)
        CODPARAM$(6) = "IMP-RNI": DOCPARAM$(6) = Str$(NNIN1#)
        CODPARAM$(7) = "IMP-TOTA": DOCPARAM$(7) = Str$(TOTA1#)
        '
        ITEXTO$ = NUMTEX$(TOTA1#) + "."
        'LOTEX% = Val(ATRIFORM$("", "LOIMPLET")): If LOTEX% < 20 Then LOTEX% = 20
        Call FRATEX(ITEXTO$, NRX1$, NRX2$, NRX3$, LOTEX%)
        IMPLET1$ = TRIM$(NRX1$) + " " + "***************************************************************************************************************"
        IMPLET2$ = TRIM$(NRX2$) + " " + "***************************************************************************************************************"
        '
        CODPARAM$(8) = "IMP-LET1": DOCPARAM$(8) = IMPLET1$
        CODPARAM$(9) = "IMP-LET2": DOCPARAM$(9) = IMPLET2$
        '
        CODPARAM$(10) = "FECH-EMI": DOCPARAM$(10) = FORFACMAI.Text9.TEXT
        CODPARAM$(11) = "NRO-REMI": DOCPARAM$(11) = ""
        CODPARAM$(12) = "NRO-AENT": DOCPARAM$(12) = ""
        CODPARAM$(13) = "TIPO-COM": DOCPARAM$(13) = UCase$(TIPODOC$)
        CODPARAM$(14) = "ITO-IVA": DOCPARAM$(14) = Str$(Val(DOCPARAM$(5)) + Val(DOCPARAM$(6)))
        '
        CAPARDOC% = 14
        '
        FEVENCIM$ = ""
        CAPARDOC% = CAPARDOC% + 1
        CODPARAM$(CAPARDOC%) = "FECH-VEN": DOCPARAM$(CAPARDOC%) = FEVENCIM$
        '
        Call PRINTDOC("FORFAC-" + LETRAFAC$) ' FACTURACION
        '
     End If
     '
End Sub
'

Sub CLEARFORFACMAI()
   '
   Call CIERRARCH("!CARCONPE")
   Call CIERRARCH("!CARITPED")
   Call BLANARCH("!CARCONPE")
   Call BLANARCH("!CARITPED")
   '
   FORFACMAI.List1.Clear
   FORFACMAI.List2.Clear
   FORFACMAI.Text1.TEXT = ""
   FORFACMAI.Text2.TEXT = ""
   FORFACMAI.Text3.TEXT = ""
   FORFACMAI.Text4.TEXT = ""
   FORFACMAI.Text5.TEXT = ""
   FORFACMAI.Text6.TEXT = "X"
   FORFACMAI.Text7.TEXT = ""
   '
   FEFAC% = HOY(HO$)
   FORFACMAI.Text9.TEXT = HO$
    '
   FORFACMAI.Text11.TEXT = ""
   '
   FORFACMAI.Label14.Caption = ""
   FORFACMAI.Label18.Caption = ""
   FORFACMAI.Label20.Caption = ""
   '
   FORFACMAI.Text1.SetFocus
   '
End Sub
'
Sub CLEARDEBILIB()
   '
   FORDELIB.Text1.TEXT = ""
   FORDELIB.Text2.TEXT = ""
   FORDELIB.Text3.TEXT = ""
   FORDELIB.Text4.TEXT = ""
   FORDELIB.Text5.TEXT = ""
   FORDELIB.Text6.TEXT = "X"
   FORDELIB.Text7.TEXT = ""
   '
   FEFAC% = HOY(HO$)
   FORDELIB.Text9.TEXT = HO$
    '
   FORDELIB.Text11.TEXT = ""
   '
   FORDELIB.Text10.TEXT = ""
   FORDELIB.Label18.Caption = ""
   FORDELIB.Label20.Caption = ""
   FORDELIB.Label23.Caption = ""
   '
   FORDELIB.Text1.SetFocus
   '

End Sub
'

Sub CLEARCRELIB()
   '
   FORCRELIB.Text1.TEXT = ""
   FORCRELIB.Text2.TEXT = ""
   FORCRELIB.Text3.TEXT = ""
   FORCRELIB.Text4.TEXT = ""
   FORCRELIB.Text5.TEXT = ""
   FORCRELIB.Text6.TEXT = "X"
   FORCRELIB.Text7.TEXT = ""
   '
   FEFAC% = HOY(HO$)
   FORCRELIB.Text9.TEXT = HO$
    '
   FORCRELIB.Text11.TEXT = ""
   '
   FORCRELIB.Text10.TEXT = ""
   FORCRELIB.Label18.Caption = ""
   FORCRELIB.Label20.Caption = ""
   FORCRELIB.Label23.Caption = ""
   '
   FORCRELIB.Text1.SetFocus
   '

End Sub
'

Sub CLEARFORCREDEV()
   '
   Call CIERRARCH("!CARITPED")
   Call BLANARCH("!CARITPED")
   '
   FORCREDEV.List2.Clear
   FORCREDEV.Text1.TEXT = ""
   FORCREDEV.Text2.TEXT = ""
   FORCREDEV.Text3.TEXT = ""
   FORCREDEV.Text4.TEXT = ""
   FORCREDEV.Text5.TEXT = ""
   FORCREDEV.Text6.TEXT = "X"
   FORCREDEV.Text7.TEXT = ""
   '
   FEFAC% = HOY(HO$)
   FORCREDEV.Text9.TEXT = HO$
   '
   FORCREDEV.Text11.TEXT = ""
   '
   FORCREDEV.Label14.Caption = ""
   FORCREDEV.Label18.Caption = ""
   FORCREDEV.Label20.Caption = ""
   '
   FORCREDEV.Text1.SetFocus
   '
End Sub
'

Sub CLEARFACOPIO()
   '
   FORFACOPIO.Text1.TEXT = ""
   FORFACOPIO.Text2.TEXT = ""
   FORFACOPIO.Text3.TEXT = ""
   FORFACOPIO.Text4.TEXT = ""
   FORFACOPIO.Text5.TEXT = ""
   FORFACOPIO.Text6.TEXT = "X"
   FORFACOPIO.Text7.TEXT = ""
   '
   FEFAC% = HOY(HO$)
   FORFACOPIO.Text9.TEXT = HO$
    '
   FORFACOPIO.Text11.TEXT = ""
   '
   FORFACOPIO.Text10.TEXT = ""
   FORFACOPIO.Label18.Caption = ""
   FORFACOPIO.Label20.Caption = ""
   FORFACOPIO.Label23.Caption = ""
   '
   FORFACOPIO.Text1.SetFocus
   '

End Sub

Sub DEBILIB1(OKEY%)
   '
   OKEY% = 0
   '
   NC% = Val(FORDELIB.Text1.TEXT)
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
   FORDELIB.Text1.TEXT = Str$(NC%)
   FORDELIB.Text2.TEXT = RASOCLI$(NC%)
   FORDELIB.Text3.TEXT = DOMICLI$(NC%)
   FORDELIB.Text4.TEXT = CPOSCLI$(NC%)
   FORDELIB.Text5.TEXT = LOCACLI$(NC%)
   FORDELIB.Text7.TEXT = TRIM$(CUITCLI$(NC%)) + " - " + POSIVAC$(NC%)
   '
   FORDELIB.TEXT21(1).TEXT = Str$(CPAGCLI%(NC%))
   FORDELIB.TEXT21(2).TEXT = Str$(VENDCLI%(NC%))
   If Val(FORDELIB.TEXT21(2).TEXT) < 1 Then
      FORDELIB.Eco21(2).TEXT = "GERENCIA"
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
   FORDELIB.Text6.TEXT = LETRAFAC$
   FORDELIB.Text12.TEXT = TRIM$(Str$(ALIVA))
   If PIVAFAC% = 5 Then
      Text12.TEXT = "0.0"
   End If
   TIDOC% = 3                 ' tipo de documento = NOTA DE DEBITO
   '
   NROPRO& = 0
   Call NUFACTURA(NROPRO&)
   FORDELIB.Text11.TEXT = "ND." + TRIM$(Str$(NROPRO&))
   '
   OKEY% = 1
   '
End Sub

Sub CREDILIB1(OKEY%)
   '
   OKEY% = 0
   '
   NC% = Val(FORCRELIB.Text1.TEXT)
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
   FORCRELIB.Text1.TEXT = Str$(NC%)
   FORCRELIB.Text2.TEXT = RASOCLI$(NC%)
   FORCRELIB.Text3.TEXT = DOMICLI$(NC%)
   FORCRELIB.Text4.TEXT = CPOSCLI$(NC%)
   FORCRELIB.Text5.TEXT = LOCACLI$(NC%)
   FORCRELIB.Text7.TEXT = TRIM$(CUITCLI$(NC%)) + " - " + POSIVAC$(NC%)
   '
   FORCRELIB.TEXT21(1).TEXT = "0"
   FORCRELIB.Eco21(1) = "CONTADO"
   FORCRELIB.TEXT21(2).TEXT = Str$(VENDCLI%(NC%))
   If Val(FORCRELIB.TEXT21(2).TEXT) < 1 Then
      FORCRELIB.Eco21(2).TEXT = "GERENCIA"
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
   FORCRELIB.Text6.TEXT = LETRAFAC$
   TIDOC% = 6                 ' tipo de documento = NOTA DE CREDITO
   '
   NROPRO& = 0
   Call NUFACTURA(NROPRO&)
   FORCRELIB.Text11.TEXT = "NC." + TRIM$(Str$(NROPRO&))
   '
   OKEY% = 1
   '
End Sub

Sub CREDEVOL1(OKEY%)
   '
   OKEY% = 0
   '
   NC% = Val(FORCREDEV.Text1.TEXT)
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
   FORCREDEV.Text1.TEXT = Str$(NC%)
   FORCREDEV.Text2.TEXT = RASOCLI$(NC%)
   FORCREDEV.Text3.TEXT = DOMICLI$(NC%)
   FORCREDEV.Text4.TEXT = CPOSCLI$(NC%)
   FORCREDEV.Text5.TEXT = LOCACLI$(NC%)
   FORCREDEV.Text7.TEXT = TRIM$(CUITCLI$(NC%)) + " - " + POSIVAC$(NC%)
   '
   FORCREDEV.TEXT21(1).TEXT = Str$(LIPRCLI%(NC%))
   If Val(FORCREDEV.TEXT21(1)) < 1 Then FORCREDEV.TEXT21(1) = "1"
   FORCREDEV.Eco21(1).TEXT = DELIPRE$(Val(FORCREDEV.TEXT21(1).TEXT))
   If FORCREDEV.Eco21(1).TEXT = "" Then
      Call MENSERR(24, "Lista de precios no Válida")
      FORCREDEV.TEXT21(1).SetFocus
      Exit Sub
   End If
   LPRE% = Val(FORCREDEV.TEXT21(1).TEXT)
   ARCHIPRE$ = TRIM$(Str$(LPRE%))
   While Len(ARCHIPRE$) < 3: ARCHIPRE$ = "0" + ARCHIPRE$: Wend
   ARCHIPRE$ = "LIPRE" + ARCHIPRE$
   '
   FORCREDEV.TEXT21(2).TEXT = Str$(VENDCLI%(NC%))
   FORCREDEV.Eco21(2).TEXT = DEVENDE$(Val(FORCREDEV.TEXT21(2).TEXT))
   If Val(FORCREDEV.TEXT21(2).TEXT) < 1 Then
      FORCREDEV.Eco21(2).TEXT = "GERENCIA"
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
   FORCREDEV.Text6.TEXT = LETRAFAC$
   TIDOC% = 6                 ' tipo de documento = NOTA DE CREBITO
   '
   NROPRO& = 0
   Call NUFACTURA(NROPRO&)
   FORCREDEV.Text11.TEXT = "NC." + TRIM$(Str$(NROPRO&))
   '
   OKEY% = 1
   '
End Sub
'

Sub FACOPIO1(OKEY%)
   '
   OKEY% = 0
   '
   NC% = Val(FORFACOPIO.Text1.TEXT)
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
   FORFACOPIO.Text1.TEXT = Str$(NC%)
   FORFACOPIO.Text2.TEXT = RASOCLI$(NC%)
   FORFACOPIO.Text3.TEXT = DOMICLI$(NC%)
   FORFACOPIO.Text4.TEXT = CPOSCLI$(NC%)
   FORFACOPIO.Text5.TEXT = LOCACLI$(NC%)
   FORFACOPIO.Text7.TEXT = TRIM$(CUITCLI$(NC%)) + " - " + POSIVAC$(NC%)
   '
   FORFACOPIO.TEXT21(1).TEXT = Str$(CPAGCLI%(NC%))
   FORFACOPIO.TEXT21(2).TEXT = Str$(VENDCLI%(NC%))
   If Val(FORFACOPIO.TEXT21(2).TEXT) < 1 Then
      FORFACOPIO.Eco21(2).TEXT = "GERENCIA"
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
   FORFACOPIO.Text6.TEXT = LETRAFAC$
   TIDOC% = 2                ' tipo de documento = NOTA DE DEBITO
   '
   NROPRO& = 0
   Call NUFACTURA(NROPRO&)
   FORFACOPIO.Text11.TEXT = "FC." + TRIM$(Str$(NROPRO&))
   '
   OKEY% = 1
   '
End Sub
'

Sub PRIDEBILIB()
    '
    FORIPRE$ = Control$("", "FORFAC-" + LETRAFAC$)
    If FORIPRE$ = "" Then
       Call MENSERR(24, "Falta Definir Formulario de Impresión\para Facturación Tipo '" + LETRAFAC$ + "'.")
       Exit Sub
    End If
    '
    Screen.MousePointer = 11
    HOII% = HOY(HOS$)
    FECHDOC$ = FORDELIB.Text9.TEXT
    NROPE& = Val(Mid$(FORDELIB.Text11.TEXT, 4))
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
    ITOT# = Val(FORDELIB.Text10.TEXT)
    If LETRAFAC$ = "B" Then
       ITOT# = (Int(ITOT# * (100 + ALIVA) + 0.5)) / 100
    End If
    '
    TTXX$ = FORDELIB.Text8.TEXT
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
    CPAG% = Val(FORDELIB.TEXT21(1).TEXT)
    CONPA$ = DECONPA$(CPAG%)
    NOCOM$ = ""
    '
    CODPARAM$(1) = "CIVA-CLI": DOCPARAM$(1) = LETRAFAC$
    CODPARAM$(2) = "NRO-OCOM": DOCPARAM$(2) = ""
    CODPARAM$(3) = "COND-PAG": DOCPARAM$(3) = CONPA$
    CODPARAM$(4) = "IMP-NETO": DOCPARAM$(4) = FORDELIB.Text10.TEXT
    CODPARAM$(5) = "IMP-IVA": DOCPARAM$(5) = FORDELIB.Label18.Caption
    CODPARAM$(6) = "IMP-RNI": DOCPARAM$(6) = FORDELIB.Label23.Caption
    CODPARAM$(7) = "IMP-TOTA": DOCPARAM$(7) = FORDELIB.Label20.Caption
    '
    ITEXTO$ = NUMTEX$(Val(FORDELIB.Label20.Caption)) + "."
    LOTEX% = Val(ATRIFORM$("", "LOIMPLET")): If LOTEX% < 20 Then LOTEX% = 20
    Call FRATEX(ITEXTO$, NRX1$, NRX2$, NRX3$, LOTEX%)
    IMPLET1$ = TRIM$(NRX1$) + " " + "***************************************************************************************************************"
    IMPLET2$ = TRIM$(NRX2$) + " " + "***************************************************************************************************************"
    '
    CODPARAM$(8) = "IMP-LET1": DOCPARAM$(8) = IMPLET1$
    CODPARAM$(9) = "IMP-LET2": DOCPARAM$(9) = IMPLET2$
    '
    CODPARAM$(10) = "FECH-EMI": DOCPARAM$(10) = FORDELIB.Text9.TEXT
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
    FORIPRE$ = Control$("", "FORFAC-" + LETRAFAC$)
    If FORIPRE$ = "" Then
       Call MENSERR(24, "Falta Definir Formulario de Impresión\para Facturación Tipo '" + LETRAFAC$ + "'.")
       Exit Sub
    End If
    '
    Screen.MousePointer = 11
    HOII% = HOY(HOS$)
    FECHDOC$ = FORCRELIB.Text9.TEXT
    NROPE& = Val(Mid$(FORCRELIB.Text11.TEXT, 4))
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
    ITOT# = Val(FORCRELIB.Text10.TEXT)
    If LETRAFAC$ = "B" Then
       ITOT# = (Int(ITOT# * (100 + ALIVA) + 0.5)) / 100
    End If
    '
    TTXX$ = FORCRELIB.Text8.TEXT
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
    CPAG% = Val(FORCRELIB.TEXT21(1).TEXT)
    CONPA$ = DECONPA$(CPAG%)
    NOCOM$ = ""
    '
    CODPARAM$(1) = "CIVA-CLI": DOCPARAM$(1) = LETRAFAC$
    CODPARAM$(2) = "NRO-OCOM": DOCPARAM$(2) = NOCOM$
    CODPARAM$(3) = "COND-PAG": DOCPARAM$(3) = CONPA$
    CODPARAM$(4) = "IMP-NETO": DOCPARAM$(4) = FORCRELIB.Text10.TEXT
    CODPARAM$(5) = "IMP-IVA": DOCPARAM$(5) = FORCRELIB.Label18.Caption
    CODPARAM$(6) = "IMP-RNI": DOCPARAM$(6) = FORCRELIB.Label23.Caption
    CODPARAM$(7) = "IMP-TOTA": DOCPARAM$(7) = FORCRELIB.Label20.Caption
    '
    ITEXTO$ = NUMTEX$(Val(FORCRELIB.Label20.Caption)) + "."
    LOTEX% = Val(ATRIFORM$("", "LOIMPLET")): If LOTEX% < 20 Then LOTEX% = 20
    Call FRATEX(ITEXTO$, NRX1$, NRX2$, NRX3$, LOTEX%)
    IMPLET1$ = TRIM$(NRX1$) + " " + "***************************************************************************************************************"
    IMPLET2$ = TRIM$(NRX2$) + " " + "***************************************************************************************************************"
    '
    CODPARAM$(8) = "IMP-LET1": DOCPARAM$(8) = IMPLET1$
    CODPARAM$(9) = "IMP-LET2": DOCPARAM$(9) = IMPLET2$
    '
    CODPARAM$(10) = "FECH-EMI": DOCPARAM$(10) = FORCRELIB.Text9.TEXT
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
    FORIPRE$ = Control$("", "FORFAC-" + LETRAFAC$)
    If FORIPRE$ = "" Then
       Call MENSERR(24, "Falta Definir Formulario de Impresión\para Facturación Tipo '" + LETRAFAC$ + "'.")
       Exit Sub
    End If
    '
    Screen.MousePointer = 11
    HOII% = HOY(HOS$)
    FECHDOC$ = FORCREDEV.Text9.TEXT
    NROPE& = Val(Mid$(FORCREDEV.Text11.TEXT, 4))
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
    For U& = 1 To FORCREDEV.List2.ListCount
       Z$ = FORCREDEV.List2.List(U& - 1)
       CACO = Val(Left$(Z$, 4))
       If CACO >= 0.05 Then
          CODX$ = Mid$(Z$, 7, 15)
          CIIX% = CODINT%(CODX$)
          DECO$ = AJUSTI$(CODX$, 16) + DESCRIT0$(CIIX%)
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
    CODPARAM$(4) = "IMP-NETO": DOCPARAM$(4) = FORCREDEV.Label14.Caption
    CODPARAM$(5) = "IMP-IVA": DOCPARAM$(5) = FORCREDEV.Label18.Caption
    CODPARAM$(6) = "IMP-RNI": DOCPARAM$(6) = FORCREDEV.Label23.Caption
    CODPARAM$(7) = "IMP-TOTA": DOCPARAM$(7) = FORCREDEV.Label20.Caption
    '
    ITEXTO$ = NUMTEX$(Val(FORCREDEV.Label20.Caption)) + "."
    LOTEX% = Val(ATRIFORM$("", "LOIMPLET")): If LOTEX% < 20 Then LOTEX% = 20
    Call FRATEX(ITEXTO$, NRX1$, NRX2$, NRX3$, LOTEX%)
    IMPLET1$ = TRIM$(NRX1$) + " " + "***************************************************************************************************************"
    IMPLET2$ = TRIM$(NRX2$) + " " + "***************************************************************************************************************"
    '
    CODPARAM$(8) = "IMP-LET1": DOCPARAM$(8) = IMPLET1$
    CODPARAM$(9) = "IMP-LET2": DOCPARAM$(9) = IMPLET2$
    '
    CODPARAM$(10) = "FECH-EMI": DOCPARAM$(10) = FORDELIB.Text9.TEXT
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

Sub PRIFACOPIO()
    '
    FORIPRE$ = Control$("", "FORFAC-" + LETRAFAC$)
    If FORIPRE$ = "" Then
       Call MENSERR(24, "Falta Definir Formulario de Impresión\para Facturación Tipo '" + LETRAFAC$ + "'.")
       Exit Sub
    End If
    '
    Screen.MousePointer = 11
    HOII% = HOY(HOS$)
    FECHDOC$ = FORFACOPIO.Text9.TEXT
    NROPE& = Val(Mid$(FORFACOPIO.Text11.TEXT, 4))
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
    TIPODOC$ = "Factura de Acopio"
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
    ITOT# = Val(FORFACOPIO.Text10.TEXT)
    If LETRAFAC$ = "B" Then
       ITOT# = (Int(ITOT# * (100 + ALIVA) + 0.5)) / 100
    End If
    '
    TTXX$ = FORFACOPIO.Text8.TEXT
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
    CPAG% = Val(FORFACOPIO.TEXT21(1).TEXT)
    CONPA$ = DECONPA$(CPAG%)
    NOCOM$ = ""
    '
    CODPARAM$(1) = "CIVA-CLI": DOCPARAM$(1) = LETRAFAC$
    CODPARAM$(2) = "NRO-OCOM": DOCPARAM$(2) = ""
    CODPARAM$(3) = "COND-PAG": DOCPARAM$(3) = CONPA$
    CODPARAM$(4) = "IMP-NETO": DOCPARAM$(4) = FORFACOPIO.Text10.TEXT
    CODPARAM$(5) = "IMP-IVA": DOCPARAM$(5) = FORFACOPIO.Label18.Caption
    CODPARAM$(6) = "IMP-RNI": DOCPARAM$(6) = FORFACOPIO.Label23.Caption
    CODPARAM$(7) = "IMP-TOTA": DOCPARAM$(7) = FORFACOPIO.Label20.Caption
    '
    ITEXTO$ = NUMTEX$(Val(FORFACOPIO.Label20.Caption)) + "."
    LOTEX% = Val(ATRIFORM$("", "LOIMPLET")): If LOTEX% < 20 Then LOTEX% = 20
    Call FRATEX(ITEXTO$, NRX1$, NRX2$, NRX3$, LOTEX%)
    IMPLET1$ = TRIM$(NRX1$) + " " + "***************************************************************************************************************"
    IMPLET2$ = TRIM$(NRX2$) + " " + "***************************************************************************************************************"
    '
    CODPARAM$(8) = "IMP-LET1": DOCPARAM$(8) = IMPLET1$
    CODPARAM$(9) = "IMP-LET2": DOCPARAM$(9) = IMPLET2$
    '
    CODPARAM$(10) = "FECH-EMI": DOCPARAM$(10) = FORFACOPIO.Text9.TEXT
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

