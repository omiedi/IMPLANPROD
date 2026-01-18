Attribute VB_Name = "ADMIOCOM"
Sub OCOMNUE()                           ' NUEVA ORDEN DE COMPRA
    '
    HO% = HOY(HOS$)
    LU$ = LUDAT$
    '
    FORIPRE$ = TRIM$(CONTROL$("", "FORMOCOM"))
    If FORIPRE$ <> "" Then
        LU$ = LUDAT$
        If EXISTE%(LU$ + FORIPRE$ + ".FRM") <> 1 Then
            RMCC$ = REPLACAR$(LU$ + FORIPRE$ + ".FRM", "\", "/")
            Call COMUNI("Error de Configuración:\  \No Existe Formulario de Impresión\de Orden de Compra:\  \'" + RMCC$ + "'.")
            Exit Sub
        End If
    End If
    '
    MONUME$ = "AUTO"
    If CONTROL$("", "NUMEOCOM") = "MANUAL" Then
        MONUME$ = "MANUAL"
    End If
    '
    PREDESCU% = 1
    If CONTROL$("", "MODOOCOM") = "PRENETO" Then
      PREDESCU% = 0
    End If
    '
    MODO% = COMALTER%("Puede Generar Ordenes de Compra\en Forma Directa ó sobre la Base\del Cálculo MRP\\Directa\MRP\Cancelar")
    FORMOCOM.Refresh
    If MODO% < 1 Or MODO% > 2 Then
        Call CIERRARCH("*.*")
        Exit Sub
    End If
    '
    AVALIDA$ = "MASTER"
    If MODO% = 2 Then AVALIDA$ = "OCOMEMI"
    '
100 Call CIERRARCH("*.*")
    REGXX& = ULTREG&("OCOMPRA")
    NRO& = 1
    If REGXX& > 0 Then
        NRO& = 1 + CVS(Left$(REGLEIDO$("OCOMPRA", REGXX&), 4))
    End If
    If NRO& <= REGXX& Then
        Screen.MousePointer = 11
        For U& = 1 To REGXX&
            X0$ = REGLEIDO$("OCOMPRA", U&)
            If CVS(Left$(X0$, 4)) >= NRO& Then
                NRO& = 1 + CVS(Left$(X0$, 4))
            End If
        Next U&
        Screen.MousePointer = 1
    End If
    '
101 Call SELECHO("PROVED1")
    NPRO% = Val(VALACT1$("PROVED1"))
    If NPRO% < 1 Or ECOARCH$("PROVED1", MKI$(NPRO%)) = "" Then
        Call CIERRARCH("*.*")
        Call BAJALDISCO
        Exit Sub
    End If
    '
    X0$ = String$(10, 0) + Space$(160)
    Call REPLA(X0$, MKS$(NRO&), 1, 4)
    Call REPLA(X0$, MKI$(HO%), 5, 2)
    Call REPLA(X0$, MKI$(HO%), 7, 2)
    Call REPLA(X0$, MKI$(NPRO%), 9, 2)
    '
    DOMIPRO$ = TRIM$(CPOSCLI$((-1) * NPRO%))
    If DOMIPRO$ <> "" Then DOMIPRO$ = DOMIPRO$ + " - "
    Call REPLA(X0$, RASOCLI$((-1) * NPRO%), 11, 40)
    Call REPLA(X0$, DOMICLI$((-1) * NPRO%), 51, 40)
    Call REPLA(X0$, DOMIPRO$ + LOCACLI$((-1) * NPRO%), 91, 40)
    '
    NC1% = (-1) * NPRO%
    CPAG% = CPAGCLI%(NC1%)
    FORPAG$ = TRIM$(ECOARCH$("CONPAG", Chr$(CPAG%)))
    If FORPAG$ = "" Then FORPAG$ = "Contado Contra Entrega"
    '
11  X1$ = OBJPLA$("OCOMPRA-ENCA", X0$)
    If X1$ = "" Then GoTo 101
    '
    FECHEMI% = CVI(Mid$(X1$, 5, 2))
    If FECHEMI% > HO% Then
       Call MENSERR(24, "Fecha de Emision no Válida")
       GoTo 11
    End If
    '
    FECHENT% = CVI(Mid$(X1$, 7, 2))
    If FECHENT% < FECHEMI% Or FECHENT% < HO% Then
       Call MENSERR(24, "Fecha de Entrega no Válida")
       GoTo 11
    End If
    '
    ATEN$ = Mid$(X1$, 131, 40)
    '
    Call SETULTREG("!OCOMDETA", 0)
    '
102 TITU$ = TRIM$(Str$(NRO&)) + " - " + TRIM$(RASOCLI$(NC1%))
    VTB% = VENTABU%("CAROCO1/ARCHECO(1)=" + AVALIDA$ + "/TITUPAN=Nota de Pedido Nro. " + TITU$)
    If VTB% < 0 Then GoTo 101
    '
    If ULTREG&("!OCOMDETA") < 1 Then
       Call MENSERR(24, "Falta Ingresar Detalle Ordenado")
       GoTo 102
    End If
    '
    IDENOC$ = MKS$(NRO&) + MKI$(FECHEMI%) + MKI$(NPRO%)
    For J& = 1 To ULTREG&("!OCOMDETA")
       X$ = REGLEIDO$("!OCOMDETA", J&)
       Call REPLA(X$, IDENOC$, 1, 8)
       '
       CI0% = CVI(Mid$(X$, 9, 2))
       If CODEXT$(CI0%) = "" Then
          Call MENSERR(24, "Código de Artículo no Válido\en Item Número " + TRIM$(Str$(J&)) + ".")
          GoTo 102
       End If
       UNIMAS$ = UNIMED$(CI0%)
       Call REPLA(X$, UNIMAS$, 11, 4)
       If TRIM$(Mid$(X$, 59, 4)) = "" Then
          Call REPLA(X$, UNIMED$(CI0%), 59, 4)
          Call REPLA(X$, MKS$(1), 63, 4)
       End If
       '
       CANTU = CVS(Mid$(X$, 55, 4))
       If CANTU < 0.05 Then
          Call MENSERR(24, "Falta Indicar Cantidad en Código '" + TRIM$(CODEXT$(CI0%)) + "'.")
          GoTo 102
       End If
       '
       FENTREIT% = CVI(Mid$(X$, 95, 2))
       If FENTREIT% < HO% Then
          Call REPLA(X$, MKI$(FECHENT%), 95, 2)
       End If
       '
       PREUNID# = CVD(Mid$(X$, 67, 8))
       If PREUNID# < 0.00005 Then
          UPREDES$ = ""
          Call BUSCAULTIPRE(NPRO%, CI0%, UNIMAS$, UPREDES$)
          If UPREDES$ <> "" Then
            Call REPLA(X$, UPREDES$, 59, 20)
          End If
       End If
       '
       Call LEEEXPLO(CODEXT$(CI0%), 1)
       If CAIT% > 0 Then
          Call REPLA(X$, " *", 69, 2)
       End If
       '
       Call GRAREG("!OCOMDETA", X$, J&)
       '
    Next J&
    '
103 VTB% = VENTABU%("CAROCO2/TITUPAN=Nota de Pedido Nro. " + TITU$)
    If VTB% < 0 Then GoTo 102
    '
    TOTANE# = 0
    For J& = 1 To ULTREG&("!OCOMDETA")
       X$ = REGLEIDO$("!OCOMDETA", J&)
       CI0% = CVI(Mid$(X$, 9, 2))
       UNIMAS$ = Mid$(X$, 11, 4)
       CANTU = CVS(Mid$(X$, 55, 4))
       PREUNID# = CVD(Mid$(X$, 67, 8))
       PORDE = CVS(Mid$(X$, 75, 4))
       FENTRE% = CVI(Mid$(X$, 95, 2))
       '
       If CODEXT$(CI0%) = "" Then
          Call MENSERR(24, "Código de Artículo no Válido\en Item Número " + TRIM$(Str$(J&)) + ".")
          GoTo 102
       End If
       If TRIM$(UNIMAS$) = "" Then
           Call REPLA(X$, UNIMED$(CI0%), 11, 4)
       End If
       If TRIM$(Mid$(X$, 59, 4)) = "" Then
          Call REPLA(X$, UNIMED$(CI0%), 59, 4)
          Call REPLA(X$, MKS$(1), 63, 4)
       End If
       '
       If CANTU < 0.05 Then
          Call MENSERR(24, "Falta Indicar Cantidad en Código '" + TRIM$(CODEXT$(CI0%)) + "'.")
          GoTo 103
       End If
       '
       If PORDE < 0 Or PORDE > 99.99 Then
          Call MENSERR(24, "Porcentaje de Descuento no Válido\en Código '" + TRIM$(CODEXT$(CI0%)) + "'.")
          GoTo 103
       End If
       '
       If FENTRE% < FECHEMI% Then
          Call MENSERR(24, "Fecha de Entrega Incoherente\en Código '" + TRIM$(CODEXT$(CI0%)) + "'.")
          Call REPLA(X0$, MKI$(FECHEMI%), 9, 2)
          GoTo 103
       End If
       '
       PREUNINE# = (Int(PREUNID# * 100 * (100 - PORDE) + 0.5)) / 10000
       INETITEM# = CDbl((Int(100 * PREUNINE# * CANTU + 0.5)) / 100)
       TOTANE# = TOTANE# + INETITEM#
       '
       Call REPLA(X$, MKD$(PREUNINE#), 79, 8)
       Call REPLA(X$, MKD$(INETITEM#), 87, 8)
       Call GRAREG("!OCOMDETA", X$, J&)
       '
    Next J&
    '
    Call CIERRARCH("!OCOMDETA")
    '
    TOTIVA# = 0
    PVCL% = PIVACLI%(NC1%)
    If PVCL% <> 2 Then
      If PVCL% <> 4 Then
        TOTIVA# = (Int(ALIVA * TOTANE# + 0.5)) / 100
      End If
    End If
    TOTATO# = TOTANE# + TOTIVA#
    '
' presentar registro de totales y condicion de pago
    '
    Call BLOQARCH("OCOMENCA")
    Call BLOQARCH("OCOMDETA")
    Call BLOQARCH("MACONSIG")
    '
'VERIFICAR NUMERO LIBRE
    '
    For J& = 1 To ULTREG&("!OCOMDETA")
       X$ = REGLEIDO$("!CARDETOC", J&)
       CI0% = CVI(Left$(X$, 2))
       CANTU = CVS(Mid$(X$, 47, 4))
       PREUNID# = CVD(Mid$(X$, 51, 8))
       'PORDE = CVS(Mid$(X$, 55, 4))
       FENTRE% = CVI(Mid$(X$, 67, 2))
       CONSIMAT% = 0: If TRIM$(Mid$(X$, 69, 2)) <> "" Then CONSIMAT% = 1
       '
       X2$ = Left$(X1$, 64)
       Call REPLA(X2$, MKI$(FENTRE%), 9, 2)
       Call REPLA(X2$, MKI$(CI0%), 11, 2)
       Call REPLA(X2$, MKS$(CANTU), 13, 4)
       PREUNI = PREUNID#
       Call REPLA(X2$, "@$" + MKS$(PREUNI) + MKS$(PORDE), 47, 10)
       Call REPLA(X2$, MKS$(0), 58, 4)
       Call REPLA(X2$, Chr$(1 + CONSIMAT%), 57, 1) ' normal
       Call REPLA(X1$, Chr$(0), 62, 1)          ' abierta
       '
       Call REGAPP("OCOMPRA", X2$)
       '
       EMICONSI% = 0
       If CONSIMAT% = 1 Then
          '
          NOCOM$ = TRIM$(Str$(NRO&))
          While Len(NOCOM$) < 6: NOCOM$ = "0" + NOCOM$: Wend
          ITNU$ = TRIM$(Str$(J&))
          While Len(ITNU$) < 2: ITNU$ = "0" + ITNU$: Wend
          NOCOM$ = "OC-" + Left$(NOCOM$, 6) + "-" + ITNU$
          '
          Call LEEEXPLO(CODEXT$(CI0%), CANTU)
          For I% = 1 To CAIT%
             YY$ = REGBLAN$("MACONSIG")
             Call REPLA(YY$, MKI$(FEMI%), 1, 2)
             Call REPLA(YY$, NPX$, 3, 2)
             Call REPLA(YY$, NOCOM$, 21, 12)
             '
             CII% = CIJ%(I%)
             CANTI# = CDbl(CANTU * USOI(I%))
             '
             If TRIM$(CODEXT$(CII%)) <> "" Then
               If CANTI# > 0.05 Then
                 Call REPLA(YY$, MKI$(CII%), 33, 2)
                 Call REPLA(YY$, UNIMED$(CII%), 51, 4)
                 Call REPLA(YY$, MKD$(CANTI#), 55, 8)
                 Call REGAPP("MACONSIG", YY$)
                 EMICONSI% = 1
               End If
             End If
          Next I%
       End If
       '
    Next J&
    '
    Call CIERRARCH("OCOMPRA")
    Call CIERRARCH("MACONSIG")
    '
    PRIPLANOI% = 0
    If CONTROL$("OCOMPRA", "PRIPLANO") = "SI" Then PRIPLANOI% = 1
    PRIESPEC% = 0
    If CONTROL$("OCOMPRA", "PRIESPEC") = "SI" Then PRIESPEC% = 1
    PRIHRINS% = 0
    If CONTROL$("OCOMPRA", "PRIHRINS") = "SI" Then PRIHRINS% = 1
    '
' preguntar si imprime si/no
    '
    If TRIM$(FORIPRE$) <> "" Then
      '
      Screen.MousePointer = 11
      VTB% = VENTABU%("OBSERVOF/TITUPAN=Pedido de Suministro Nro: " + TRIM$(Str$(NUOCO&)))
      '
      Screen.MousePointer = 11
      FECHDOC$ = FETEXTO$(FEMI%)
      NUMEDOC$ = TRIM$(Str$(NRO&))
      While Len(NUMEDOC$) < 6
        NUMEDOC$ = "0" + NUMEDOC$
      Wend
      TIPODOC$ = "Nota de Pedido"
      '
      CODPARAM$(1) = "COND-PAG"
      CODPARAM$(2) = "OBSERVA"
      CODPARAM$(3) = "IMP-NETO"
      CODPARAM$(4) = "POR-IVA"
      CODPARAM$(5) = "IMP-IVA"
      CODPARAM$(6) = "IMP-TOTA"
      CODPARAM$(7) = "VENDEDOR"
      CODPARAM$(8) = "LIST-PRE"
      CODPARAM$(9) = "ZONA-VEN"
      CODPARAM$(10) = "FECH-LAN"
      CODPARAM$(11) = "IMPNEDES"
      CAPARDOC% = 11
      '
      DOCPARAM$(1) = FORPAG$
      DOCPARAM$(2) = OBSE$
      DOCPARAM$(3) = Str$(TOTANE#)
      DOCPARAM$(4) = Str$(ALIVA)
      DOCPARAM$(5) = Str$(TOTIVA#)
      DOCPARAM$(6) = Str$(TOTATO#)
      DOCPARAM$(7) = Mid$(X1$, 147, 20)
      DOCPARAM$(8) = Mid$(X1$, 67, 40)
      DOCPARAM$(9) = Mid$(X1$, 167, 40)
      DOCPARAM$(10) = Mid$(X1$, 107, 40)
      DOCPARAM$(11) = Str$(TOTANE#)
      '
      CODTABU1$(1) = "COD-MAT"
      CODTABU1$(2) = "DES-MAT"
      CODTABU1$(3) = "UNIMED"
      CODTABU1$(4) = "CANTIDAD"
      CODTABU1$(5) = "FECH-VEN"
      CODTABU1$(6) = "REF-MAT1"
      CODTABU1$(7) = "PRE-LIST"
      CODTABU1$(8) = "PORDESCU"
      CODTABU1$(9) = "PRE-NETO"
      CODTABU1$(10) = "NET-ITEM"
      CODTABU1$(11) = "CAKILOS"
      CODTABU1$(12) = "F-PLANO"
      CODTABU1$(13) = "ORD-ITEM"
      CACOLTAB1% = 13
      '
      DESTIDOC% = NC1%
      CAITAB1% = 0
      CAITCRU% = 0
      '
      For J& = 1 To ULTREG&("!CARDETOC")
         X$ = REGLEIDO$("!CARDETOC", J&)
         CI0% = CVI(Left$(X$, 2))
         CANTU = CVS(Mid$(X$, 47, 4))
         PREUNID# = CVD(Mid$(X$, 51, 8))
         'PORDE = CVS(Mid$(X$, 55, 4))
         FEX = CVI(Mid$(X$, 67, 2))
         CONSIMAT% = 0: If TRIM$(Mid$(X$, 69, 2)) <> "" Then CONSIMAT% = 1
         '
         CAITAB1% = CAITAB1% + 1
         TETABU1$(1, CAITAB1%) = CODEXT$(CI0%)
         TETABU1$(2, CAITAB1%) = DESCRIT0$(CI0%)
         TETABU1$(3, CAITAB1%) = Mid$(X$, 3, 4)
         TETABU1$(4, CAITAB1%) = TRIM$(Str$(CANTU))
         TETABU1$(5, CAITAB1%) = FECHATEX$(FEX)
         TETABU1$(6, CAITAB1%) = Mid$(X$, 69, 2)
         TETABU1$(7, CAITAB1%) = Str$(PREUNID#)
         TETABU1$(8, CAITAB1%) = Str$(PORDE)
         TETABU1$(9, CAITAB1%) = Str$(PREUNID# * (1 - PORDE / 100))
         TETABU1$(10, CAITAB1%) = Str$(CANTU * PREUNID# * (1 - PORDE / 100))
         TETABU1$(11, CAITAB1%) = Str$(CANTU * PESUNI(CI0%))
         TETABU1$(12, CAITAB1%) = FEPLANO$(CI0%)
6         TETABU1$(13, CAITAB1%) = TRIM$(Str$(J&))
         '
109   Next J&
      '
      Call PRINTDOC("FORMOCOM")   ' Orden de Compra
      '
      DOCUORIG$ = "N/Ped-" + NUMEDOC$
      NUMEDOC$ = ""
      For J& = 1 To ULTREG&("!CARDETOC")
         X$ = REGLEIDO$("!CARDETOC", J&)
         CI0% = CVI(Left$(X$, 2))
         If CI0% > 0 Then
           If CI0% <= NUMAS% Then
             If PRIPLANOI% = 1 Then Call PRIPLANO(CI0%)
             If PRIESPEC% = 1 Then Call PRIHOJESP(CI0%)
             If PRIHRINS% = 1 Then Call PRIHOJRINSP(CI0%)
           End If
         End If
      Next J&
      '
    End If
    Call CIERRARCH("*.*")
    '
    If EMICONSI% = 1 Then
       If COMALTER%("Puede Ahora Generar\el Remito de Consignacion de Materiales\\Consignar Materiales\Continuar") = 1 Then
          Call ENTRECONSI(NPRO%)
       End If
    End If
    Call CIERRARCH("*.*")
    '
    GoTo 100
    '
299 Call CIERRARCH("*.*")
    '
End Sub
'


Sub OCOMANU()
    '
    HO% = HOY(HOS$)
    LU$ = LUDAT$
    '
    TITU$ = "PROCESO DE ANULACION DE ORDENES DE COMPRA"

100  NROOF$ = ""
     '
1400 If INDI% = 0 Then
       '
       Screen.MousePointer = 11
       Call COPYSTRU("OCOMPRA", "OCOMANU")
       J& = 0: K& = 0: OCOA& = 0
       '
       For I& = 1 To ULTREG&("OCOMPRA")
         XX$ = REGLEIDO$("OCOMPRA", I&)
         '
         If CVS(Left(XX$, 4)) >= 0.5 Then    ' depura O/C's
             K& = K& + 1                     ' sin numero
             Call GRAREG("OCOMPRA", XX$, K&)
         End If
         '
         ST% = Asc(Mid$(XX$, 62, 1))
         '
         If ST% < 1 Then     '    excluye cerradas
            If CVS(Left$(XX$, 4)) <> OCOA& Then
               J& = J& + 1
               Call GRAREG("OCOMANU", XX$, J&)
               OCOA& = CVS(Left$(XX$, 4))
            End If
         End If
       Next I&
       '
       Call SETULTREG("OCOMPRA", K&)
       Call SETULTREG("OCOMANU", J&)
       Call CIERRARCH("OCOMANU")
       '
       If J& < 1 Then
          Call MENSERR(24, "No hay Ordenes de Compra\Abiertas que Puedan Anularse.")
          Screen.MousePointer = 1
          Exit Sub
       End If
       '
       INDI% = 1
       Screen.MousePointer = 1
       '
     End If
     
1405 Call FRESHECHO("OCOMANU")
     If ULTREG&("OCOMANU") < 1 Then INDI% = 0: Exit Sub
     Call SELECHO("OCOMANU/Ordenes de Compra para Anulación")
     '
1410 NROOFL& = Val(VALACT1$("OCOMANU"))
     If NROOFL& < 1 Then
         INDI% = 0
         Exit Sub
     End If
     '
     For I& = ULTREG&("OCOMPRA") To 1 Step -1
       XX$ = REGLEIDO$("OCOMPRA", I&)
       If CVS(Left$(XX$, 4)) = NROOFL& Then GoTo 1420
     Next I&
     '
     Call COMUNI("Inconsistencia en Base de Datos\de Ordenes de Compra.\  \Consulte al Soporte de Sistemas.")
     INDI% = 0
     Exit Sub
     '
1420 YY$ = OBJPLA$("OCOMANU-1/Anulación de Ordenes de Compra", XX$)
     If YY$ = "" Then GoTo 1405
     '
     Screen.MousePointer = 11
     For J& = I& To 1 Step -1
        XX$ = REGLEIDO$("OCOMPRA", J&)
        If CVS(Left$(XX$, 4)) <> NROOFL& Then GoTo 1425
        Call REPLA(XX$, Chr$(9), 62, 1)
        Call GRAREG("OCOMPRA", XX$, J&)
     Next J&
     '
1425 J& = 0
     For I& = 1 To ULTREG&("OCOMANU")
        XX$ = REGLEIDO$("OCOMANU", I&)
        If CVS(Left$(XX$, 4)) <> NROOFL& Then
           J& = J& + 1
           Call GRAREG("OCOMANU", XX$, J&)
        End If
     Next I&
     Call SETULTREG("OCOMANU", J&)
     '
     ' CANCELAR RESERVAS
     '
     JJ& = 0
     For J& = 1 To ULTREG&("MACONSIG")
         XX$ = REGLEIDO$("MACONSIG", J&)
         If Val(Mid$(XX$, 24, 6)) <> NROOFL& Then
                JJ& = JJ& + 1
                Call GRAREG("MACONSIG", XX$, JJ&)
             Else
                CAENTRE# = CVD(Mid$(XX$, 63, 8))
                CADEVUE# = CVD(Mid$(XX$, 111, 8))
                If Abs(CAENTRE# - CADEVUE#) >= 0.05 Then
                   JJ& = JJ& + 1
                   Call GRAREG("MACONSIG", XX$, JJ&)
                End If
         End If
     Next J&
     '
     Call SETULTREG("MACONSIG", JJ&)
     Screen.MousePointer = 1
     Call CIERRARCH("*.*")
     GoTo 1405
     '
End Sub
'
Sub BUSCAULTIPRE(NPRO%, CIIX%, UNIMAS$, UPREDES$)
   '
   RFF$ = RECFILT$("OCOMDETA", 3, MKI$(NPRO%))
   If Len(RFF$) >= 4 Then
     For U& = 1 To Len(RFF$) Step 4
       REGO& = CVS(Mid$(RFF$, U&, 4))
       YX$ = REGLEIDO$("OCOMDETA", REGO&)
       If CVI(Mid$(YX$, 7, 2)) = NPRO% Then
         If CVI(Mid$(YX$, 9, 2)) = CIIX% Then
           If Mid$(YX$, 11, 4) = UNIMAS$ Then
             UPREDES$ = Mid$(YX$, 59, 20)
           End If
         End If
       End If
     Next U&
   End If
   '
End Sub

Sub CONVERCOM()
   ' conversion de formato de ordenes de compra
   '
   If ULTREG&("OCOMPRA") > 0 Then
     '
     Call BLOQARCH("OCOMPRA")
     Call BLOQARCH("OCOMENCA")
     Call BLOQARCH("OCOMDETA")
     '
     Screen.MousePointer = 11
     OCOA& = 0: NPROA% = 0: ENCAO$ = ""
     FIN& = ULTREG&("OCOMPRA")
     '
     For U& = 1 To FIN&
       Call TRACE(20, U&, FIN&)
       XX$ = REGLEIDO$("OCOMPRA", U&)
       NUOCO& = CVS(Left$(XX$, 4))
       FEMI% = CVI(Mid$(XX$, 5, 2))
       NPRO% = CVI(Mid$(XX$, 7, 2))
       FENTRE% = CVI(Mid$(XX$, 9, 2))
       CI0% = CVI(Mid$(XX$, 11, 2))
       CAOR = CVS(Mid$(XX$, 13, 4))
       OBSER$ = Mid$(XX$, 17, 30)
       PREUNID# = CVS(Mid$(XX$, 49, 4))
       DESCUNI = CVS(Mid$(XX$, 53, 4))
       TISU% = Asc(Mid$(XX$, 57, 1))
       CARE = CVS(Mid$(XX$, 58, 4))
       STAT% = Asc(Mid$(XX$, 62, 1))
       '
       If NUOCO& <> OCOA& Or NPRO% <> NPROA% Then
         If OCOA& <> 0 Then
           TOTIVA# = 0
           PVCL% = PIVACLI%((-1) * NPRO%)
           If PVCL% <> 2 Then
             If PVCL% <> 4 Then
               TOTIVA# = (Int(ALIVA * TOTANE# + 0.5)) / 100
             End If
           End If
           Call REPLA(ENCAO$, MKD$(TOTANE#), 89, 8)
           Call REPLA(ENCAO$, MKD$(TOTIVA#), 121, 8)
           Call REPLA(ENCAO$, MKD$(TOTANE# + TOTIVA#), 137, 8)
           Call REGAPP("OCOMENCA", ENCAO$)
         End If
         '
         OCOA& = NUOCO&: NPROA% = NPRO%
         ENCAO$ = REGBLAN$("OCOMENCA")
         '
         Call REPLA(ENCAO$, MKS$(NUOCO&), 1, 4)
         FEX = FEMI%
         REFE$ = FECHATEX$(FEX) + " - " + RASOCLI$((-1) * NPRO%)
         Call REPLA(ENCAO$, REFE$, 5, 44)
         Call REPLA(ENCAO$, MKI$(FEMI%), 49, 2)
         Call REPLA(ENCAO$, MKI$(NPRO%), 51, 2)
         Call REPLA(ENCAO$, Chr$(TISU%), 53, 1)
         Call REPLA(ENCAO$, Chr$(STAT%), 54, 1)
         '
         HO% = HOY(HOS$)
         HP = 100 * Val(Left$(Time$, 2)) + Val(Mid$(Time$, 4, 2))
         Call REPLA(ENCAO$, MKI$(HO%), 55, 2)
         Call REPLA(ENCAO$, MKI$(Int(HP)), 57, 2)
         Call REPLA(ENCAO$, MKI$(USNBR%), 59, 2)
         '
         If STAT% = 9 Then
           Call REPLA(ENCAO$, MKI$(HO%), 67, 2)
           Call REPLA(ENCAO$, MKI$(Int(HP)), 69, 2)
           Call REPLA(ENCAO$, MKI$(USNBR%), 71, 2)
           Call REPLA(ENCAO$, "Conversion BD", 73, 16)
         End If
         '
         CPAG% = CPAGCLI%((-1) * NPRO%)
         Call REPLA(ENCAO$, MKI$(CPAG%), 145, 2)
         REDETA& = 1 + ULTREG&("OCOMDETA")
         Call REPLA(ENCAO$, MKS$(REDETA&), 147, 4)
         Call REPLA(ENCAO$, MKS$(REDETA&), 151, 4)
         '
         Call REPLA(ENCAO$, OBSER$, 193, 64)
         '
       End If
       '
       REDET$ = REGBLAN$("OCOMDETA")
       Call REPLA(REDET$, MKS$(NUOCO&), 1, 4)
       Call REPLA(REDET$, MKI$(FEMI%), 5, 2)
       Call REPLA(REDET$, MKI$(NPRO%), 7, 2)
       Call REPLA(REDET$, MKI$(CI0%), 9, 2)
       UNIMAS$ = UNIMED$(CI0%)
       Call REPLA(REDET$, UNIMAS$, 11, 4)
       Call REPLA(REDET$, MKS$(CAOR), 55, 4)
       Call REPLA(REDET$, UNIMED$(CI0%), 59, 4)
       Call REPLA(REDET$, MKS$(1), 63, 4)
       Call REPLA(REDET$, MKD$(PREUNID#), 67, 8)
       Call REPLA(REDET$, MKD$(PREUNID#), 79, 8)
       Call REPLA(REDET$, MKD$(PREUNID# * CAOR), 87, 8)
       Call REPLA(REDET$, MKI$(FENTRE%), 95, 2)
       If TISU% > 1 Then
         Call REPLA(REDET$, " *", 97, 2)
       End If
       Call REPLA(REDET$, MKS$(CARE), 99, 4)
       RENCA& = 1 + ULTREG&("OCOMENCA")
       Call REPLA(REDET$, MKS$(RENCA&), 125, 4)
       Call REGAPP("OCOMDETA", REDET$)
       '
       TOTANE# = TOTANE# + PREUNID# * CAOR
       REDETA& = ULTREG&("OCOMDETA")
       Call REPLA(ENCAO$, MKS$(REDETA&), 151, 4)
       '
     Next U&
     '
     If OCOA& <> 0 Then
       TOTIVA# = 0
       PVCL% = PIVACLI%((-1) * NPRO%)
       If PVCL% <> 2 Then
         If PVCL% <> 4 Then
           TOTIVA# = (Int(ALIVA * TOTANE# + 0.5)) / 100
         End If
       End If
       Call REPLA(ENCAO$, MKD$(TOTANE#), 89, 8)
       Call REPLA(ENCAO$, MKD$(TOTIVA#), 121, 8)
       Call REPLA(ENCAO$, MKD$(TOTANE# + TOTIVA#), 137, 8)
       Call REGAPP("OCOMENCA", ENCAO$)
     End If
     '
     Call CIERRARCH("*.*")
     Call BAJALDISCO
     Screen.MousePointer = 1
     '
   End If
   '
End Sub
