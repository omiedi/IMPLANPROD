Attribute VB_Name = "ODESPA_RUTS"
Sub GENODESPA(NPED&, NODES&, MODOEMI%, OKEY%)
    '
    ' MODOEMI% = 1 --> DESDE GENERACION DE PEDIDO
    '            2 --> DESDE ORDENES DE DESPACHO
    '
    DIANTIDES% = XVALO(CONTROL$("", "ANTIDESP"))
    If DIANTIDES% < 1 Then DIANTIDES% = 3
    KKK% = 0
    OKEY% = 0
    '
    FORIPRE$ = TRIM$(CONTROL$("", "FORDESPA"))
    If FORIPRE$ = "" Then
        FORIPRE$ = "ODES-SAB"
    End If
    CODFOR$ = TRIM$(UCase$(FORIPRE$))
    MAXIT% = XVALO(ATRIFORM$(CODFOR$, "MAX-ITEM"))
    If MAXIT% < 1 Or MAXIT% > 256 Then MAXIT% = 10
    '
    NODES& = 0
    NODESPA& = 1: UREODE& = ULTREG&("MAVENDID")
    If UREODE& > 0 Then
      For U& = UREODE& To 1 Step -1
        ZX$ = REGLEIDO$("MAVENDID", U&)
        NODESPA& = 1 + CVS(Mid$(ZX$, 87, 4))
        If NODESPA& > 1.1 Then GoTo 5
      Next U&
    End If
5   FEDES% = HOY(HOS$)
    '
    If MODOEMI% = 1 Then
      CLIPED% = 0
      FEPED% = HOY(HOS$)
      RFG$ = RECFILT$("PEDDETA", 1, MKS$(NPED&))
      If Len(RFG$) >= 4 Then
        PDXC$ = REGLEIDO$("PEDDETA", CVS(Left$(RFG$, 4)))
        FEPED% = CVI(Mid$(PDXC$, 5, 2))
        CLIPED% = CVI(Mid$(PDXC$, 7, 2))
      End If
      If CLIPED% < 1 Then
        Call MENSERR(24, "Imposible Identificar Cliente.\Emita O.Despacho por Revisión Situación de Stocks")
        GoTo 98
      End If
      GoTo 15
    End If
    '
    REBLA$ = REGBLAN$("CARODES"): jj& = 0
    ALECPED$ = "ODESEMI"
    If MODOEMI% = 1 Then
      ALECPED$ = "PEDDETP"
    End If
    '
    ' esto se agraga para ver si ataja el error de SABO
    Call CIERRARCH(ALECPED$)
    Call CIERRARCH("*.*")
    Call BAJALDISCO
    Call INDEXARCH(ALECPED$, 1)
    '
    RFF$ = RECFILT$(ALECPED$, 1, MKS$(NPED&))
    For UKK& = 1 To Len(RFF$) Step 4
      U& = CVS(Mid$(RFF$, UKK&, 4))
      X$ = REGLEIDO$(ALECPED$, U&)
      If CVS(Left$(X$, 4)) = NPED& Then
        If MODOEMI% = 1 Then     ' lee de PEDDETP
            FEPED% = CVI(Mid$(X$, 5, 2))
            CLIPED% = CVI(Mid$(X$, 7, 2))
            FENTRE% = CVI(Mid$(X$, 25, 2))
            CIPED% = CVI(Mid$(X$, 33, 2))
            CAPENPED = CVS(Mid$(X$, 79, 4)) '- CVS(Mid$(X$, 95, 4))
            CAPREASI = CAPENPED
            REGPEDA& = CVS(Mid$(X$, 109, 4))
            REGPEDP& = U&
            FE1 = HOY(HOS$): FE2 = FENTRE%
            If DIENTRE%(FE1, FE2) > DIANTIDES% Then GoTo 10
          Else                   ' LEE DE ODESEMI
            FEPED% = CVI(Mid$(X$, 5, 2))
            CLIPED% = CVI(Mid$(X$, 7, 2))
            FENTRE% = CVI(Mid$(X$, 9, 2))
            CIPED% = CVI(Mid$(X$, 11, 2))
            CAPENPED = CVS(Mid$(X$, 13, 4))
            CAPREASI = CVS(Mid$(X$, 17, 4))
            REGPEDA& = CVS(Mid$(X$, 21, 4))
            REGPEDP& = CVS(Mid$(X$, 29, 4))
        End If
        '
8       Call VESARRAS(CIPED%)
        STODISPO = STODEPOS(CIPED%, DEPOEXPE%)
        STODISPO = STODISPO - STOCKASI(CIPED%)
        For J& = 1 To jj&  ' netea de preasignacion anterior
          Z$ = REGLEIDO$("!CARODES", J&)
          If CVI(Left$(Z$, 2)) = CIPED% Then
            STODISPO = STODISPO - CVS(Mid$(Z$, 33, 4))
          End If
        Next J&
        If STODISPO < 0 Then STODISPO = 0
        '
        If REGPEDA& > 0 Then
          If REGPEDA& <= ULTREG&("PEDDETA") Then
            Y$ = REGLEIDO$("PEDDETA", REGPEDA&)
            If CVS(Left$(Y$, 4)) = NPED& Then
              If CVI(Mid$(Y$, 33, 2)) = CIPED% Then
                If CVI(Mid$(Y$, 7, 2)) = CLIPED% Then
                  '
                  CAPEDORI = CVS(Mid$(Y$, 79, 4))
                  CAPENPED = CAPEDORI - CVS(Mid$(Y$, 95, 4)) 'Q.entregada
                  CAPENPED = CAPENPED - CVS(Mid$(Y$, 75, 4)) 'Q.ya asignada
                  If CAPENPED < 0 Then CAPENPED = 0
                  '
                  If CAPENPED >= 0.05 Then
                     CanAsig = CAPREASI
                     If CanAsig > CAPENPED Then CanAsig = CAPENPED
                     If CanAsig > STODISPO Then
                       TTXX$ = "En Codigo " + TRIM$(CODEXT$(CIPED%)) + " el Stock\"
                       TTXX$ = TTXX$ + "Disponible (" + TRIM$(Str$(STODISPO)) + ") es Insuficiente\"
                       TTXX$ = TTXX$ + "para la Cantidad a Asignar (" + TRIM$(Str$(CanAsig)) + ")"
                       OPXY% = COMALTER%(TTXX$ + "\\Continuar\Búsqueda con Re-Indexación")
                       If OPXY% = 2 Then
                         Call INDEXARCH("MOVISTO", 1)
                         Call INDEXARCH("MAVENDID", 1)
                         Call CIERRARCH("*.*")
                         GoTo 8
                       End If
                       CanAsig = STODISPO
                     End If
                     '
                     FENTRESO% = CVI(Mid$(Y$, 25, 2))
                     '
                     Z$ = REBLA$
                     Call REPLA(Z$, MKI$(CIPED%), 1, 2)
                     Call REPLA(Z$, Space$(16), 3, 16)   ' lote / color
                     Call REPLA(Z$, MKI$(FENTRESO%), 19, 2)
                     Call REPLA(Z$, MKS$(CAPEDORI), 21, 4)
                     Call REPLA(Z$, MKS$(CAPENPED), 25, 4)
                     Call REPLA(Z$, MKS$(STODISPO), 29, 4)
                     Call REPLA(Z$, MKS$(CanAsig), 33, 4)
                     Call REPLA(Z$, MKS$(REGPEDA&), 37, 4)
                     Call REPLA(Z$, MKS$(REGPEDP&), 41, 4)
                     Call REPLA(Z$, MKS$(U&), 45, 4)
                     jj& = jj& + 1
                     Call GRAREG("!CARODES", Z$, jj&)
                  End If
                  '
                End If
              End If
            End If
          End If
        End If
      End If
10  Next UKK&
    '
    Call SETULTREG("!CARODES", jj&)
    Call CIERRARCH("!CARODES")
    '
    If jj& < 1 Then
      If MODOEMI% = 1 Then
          Call MENSERR(24, "No hay Material Disponible\para Asignar a este Pedido.")
        Else
          Call MENSERR(24, "Imposible Generar Orden de Despacho\Correspondiente al Pedido Número " + TRIM$(Str$(NPED&)) + ".\  \Repita Revisión de Situación de Stocks.")
      End If
      GoTo 99
    End If
    '
15  FEPEDID$ = REPLACAR$(FETEXCOR$(FEPED%), "/", " ")
    ATRIBU$ = "TITUPAN=Pedido Nro. " + TRIM$(Str$(NPED&)) + " - " + TRIM$(rasocli$(CLIPED%)) + " - " + FEPEDID$
    VTB% = VENTABU%("CARODES/" + ATRIBU$)
    If VTB% < 0 Then
       OKEY% = (-1)
       GoTo 99
    End If
    '
    ATXXX% = 0
    For KU& = 1 To ULTREG&("!CARODES")
      Z$ = REGLEIDO$("!CARODES", KU&)
      CanAsig = CVS(Mid$(Z$, 33, 4))
      If CanAsig > CVS(Mid$(Z$, 25, 4)) Then
        Call MENSERR(24, "No Válido Asignar más que Pendiente Pedido.")
        Call REPLA(Z$, Mid$(Z$, 25, 4), 33, 4)
        Call GRAREG("!CARODES", Z$, KU&)
        GoTo 15
      End If
      If CanAsig > CVS(Mid$(Z$, 29, 4)) Then
        Call MENSERR(24, "No Válido Asignar más que Stock Disponible.")
        Call REPLA(Z$, Mid$(Z$, 29, 4), 33, 4)
        Call GRAREG("!CARODES", Z$, KU&)
        GoTo 15
      End If
      '
      CIPED% = CVI(Mid$(Z$, 1, 2))
      STODISPO = STODEPOS(CIPED%, DEPOEXPE%) - STOCKASI(CIPED%)
      If STODISPO < 0 Then STODISPO = 0
      If STODISPO < CVS(Mid$(Z$, 29, 4)) - 0.1 Then
        ATXXX% = 1
        Call REPLA(Z$, MKS$(STODISPO), 29, 4)
        Call GRAREG("!CARODES", Z$, KU&)
      End If
      '
    Next KU&
    '
    If ATXXX% > 0 Then
      Call COMUNI("Por Cambios en Stock Disponible\Revise las Cantidades Asignadas.")
      GoTo 15
    End If
    '
    Screen.MousePointer = 11
    Call BLOQARCH("MAVENDID")
    Call BLOQARCH("PEDDETA")
    Call BLOQARCH("PEDDETP")
    Call BLOQARCH("ODESEMI")
    '
    For KU& = 1 To ULTREG&("!CARODES")
      Z$ = REGLEIDO$("!CARODES", KU&)
      CIPED% = CVI(Mid$(Z$, 1, 2))
      FENTRESO% = CVI(Mid$(Z$, 19, 2))
      CAPENPED = CVS(Mid$(Z$, 25, 4))
      REGPEDP& = CVS(Mid$(Z$, 41, 4))
      If REGPEDP& > 0 Then
        If REGPEDP& <= ULTREG&("PEDDETP") Then
          X$ = REGLEIDO$("PEDDETP", REGPEDP&)
          If FENTRESO% = CVI(Mid$(X$, 25, 2)) Then
            If CIPED% = CVI(Mid$(X$, 33, 2)) Then
              CAPENPEP = CVS(Mid$(X$, 79, 4)) - CVS(Mid$(X$, 95, 4)) - CVS(Mid$(X$, 75, 4))
              If Abs(CAPENPED - CAPENPEP) < 0.01 Then
                GoTo 19
              End If
            End If
          End If
        End If
      End If
      '
      Call CIERRARCH("MAVENDID")
      Call CIERRARCH("PEDDETA")
      Call CIERRARCH("PEDDETP")
      Call CIERRARCH("ODESEMI")
      Call COMUNI("Por Cambios en B.Datos de Pedidos Pendientes\Imposible Emitir Orden de Despacho.\Cierre y Re-Intente revisando Situación de Stocks.")
      GoTo 98
      '
19  Next KU&
    '
    NODESPA& = 1: UREODE& = ULTREG&("MAVENDID")
    If UREODE& > 0 Then
      For U& = UREODE& To 1 Step -1
        ZX$ = REGLEIDO$("MAVENDID", U&)
        NODESPA& = 1 + CVS(Mid$(ZX$, 87, 4))
        If NODESPA& > 1.1 Then GoTo 25
      Next U&
    End If
25  FEDES% = HOY(HOS$)
    '
    REBLA$ = REGBLAN$("MAVENDID")
    JKI% = 0
    For KU& = 1 To ULTREG&("!CARODES")
      Z$ = REGLEIDO$("!CARODES", KU&)
      CIPED% = CVI(Mid$(Z$, 1, 2))
      LOTEA$ = Mid$(Z$, 3, 16)  ' lote / color
      FENTRESO% = CVI(Mid$(Z$, 19, 2))
      CanAsig = CVS(Mid$(Z$, 33, 4))
      REGPEDA& = CVS(Mid$(Z$, 37, 4))
      REGPEDP& = CVS(Mid$(Z$, 41, 4))
      REGODES& = CVS(Mid$(Z$, 45, 4))
      '
      If CanAsig > 0.05 Then
         X$ = REGLEIDO$("PEDDETA", REGPEDA&)
         NOPED& = CVS(Left$(X$, 4))
         FEPED% = CVI(Mid$(X$, 5, 2))
         NCLIE% = CVI(Mid$(X$, 7, 2))
         NUOCO$ = AJUSTI$(Mid$(X$, 9, 16), 16)
         PRECOVEN$ = Mid$(X$, 28, 3) 'L.PRECIO-C.PAGO-VENDEDOR
         CIXI% = CIPED%
         PORCOVE = CVI(Mid$(X$, 31, 2)) / 10
         KCOLOR$ = Mid$(X$, 35, 16)
         KTALLE$ = Mid$(X$, 51, 8)
         KLARGO = CVS(Mid$(X$, 59, 4))
         KANCHO = CVS(Mid$(X$, 63, 4))
         KESPES = CVS(Mid$(X$, 67, 4))
         PREUNIBRU# = CVS(Mid$(X$, 71, 4))
         PREUNINET# = CVS(Mid$(X$, 83, 4))
         CAENTR = 0
         FENTRE% = 0
         NREMI& = 0
         '
         Y$ = REBLA$
         Call REPLA(Y$, MKI$(NCLIE%), 1, 2)
         Call REPLA(Y$, Chr$(0), 3, 1) ' codigo de empresa
         Call REPLA(Y$, Chr$(0), 4, 1) ' CODIGO DE SUCURSAL
         Call REPLA(Y$, MKI$(CIPED%), 5, 2)
         Call REPLA(Y$, AJUSTI$(KCOLOR$, 16), 7, 16)
         Call REPLA(Y$, AJUSTI$(KTALLE$, 8), 25, 8)
         Call REPLA(Y$, MKS$(KLARGO), 33, 4)
         Call REPLA(Y$, MKS$(KANCHO), 37, 4)
         Call REPLA(Y$, MKS$(KESPES), 41, 4)
         Call REPLA(Y$, String$(4, 0), 45, 4)  ' FILLER
         UNI$ = Unimed$(CIPED%)
         Call REPLA(Y$, UNI$, 49, 4)
         Call REPLA(Y$, UNI$, 53, 4)
         Call REPLA(Y$, MKS$(1), 57, 4)  ' COEFICIENTE CONVERSION UNIDADES
         Call REPLA(Y$, "PC", 61, 2)     ' MOVIMIENTO DE ORIGEN
         Call REPLA(Y$, Chr$(1), 63, 1)  ' STATUS pendiente
         Call REPLA(Y$, Chr$(0), 64, 1)  ' FILLER
         Call REPLA(Y$, MKI$(NCLIE%), 65, 2)  ' CLIENTE
         Call REPLA(Y$, MKI$(CIXI%), 67, 2)  ' CODIGO
         Call REPLA(Y$, String$(12, 0), 69, 12) ' FILLER
         '
         Call REPLA(Y$, MKI$(FEDES%), 81, 2) ' FECHA O/DESPACHO
         Call REPLA(Y$, MKS$(Int(HORANUM(Time$))), 83, 2) ' HORA-USUA-TERM
         Call REPLA(Y$, Chr$(USNBR% Mod 100), 85, 1)
         Call REPLA(Y$, Chr$(0), 86, 1)              ' terminal
         Call REPLA(Y$, MKS$(NODESPA&), 87, 4)  ' NRO O/DESP
         JKI% = JKI% + 1
         Call REPLA(Y$, MKI$(JKI%), 91, 2)       ' NRO DE ORDEN
         Call REPLA(Y$, MKS$(NOPED&), 93, 4)
         Call REPLA(Y$, Chr$(0), 97, 1) ' codigo de empresa
         Call REPLA(Y$, Chr$(0), 98, 1) ' CODIGO DE SUCURSAL
         Call REPLA(Y$, MKS$(REGPEDA&), 99, 4) ' REGISTRO PEDDETA
         Call REPLA(Y$, MKI$(FEPED%), 103, 2) ' FECHA DE PEDIDO
         Call REPLA(Y$, AJUSTI$(NUOCO$, 16), 105, 16)
         Call REPLA(Y$, PRECOVEN$ + Chr$(0), 121, 4)
         Call REPLA(Y$, MKI$(FENTRESO%), 125, 2) ' FECHA O/DESPACHO
         Call REPLA(Y$, MKD$(PREUNIBRU#), 127, 8) ' PRECIO UNITARIO BRUTO
         Call REPLA(Y$, MKS$(0), 135, 4) ' PORCIENTO DESCUENTO item
         Call REPLA(Y$, MKD$(PREUNINET#), 139, 8) ' PRECIO UNITARIO NETO
         Call REPLA(Y$, MKS$(0), 147, 4) ' PORCIENTO DESCUENTO general pedido
         Call REPLA(Y$, MKD$(PREUNINET#), 151, 8) ' PRECIO UNITARIO NETO
         Call REPLA(Y$, Chr$(0), 159, 1)       ' MONEDA = $
         Call REPLA(Y$, Chr$(0), 160, 1)       ' FILLER
         Call REPLA(Y$, MKD$(1), 161, 8)       ' TIPO DE CAMBIO
         Call REPLA(Y$, MKD$(PREUNINET#), 169, 8) ' valor unitario neto-neto EN PESOS
         Call REPLA(Y$, MKS$(CanAsig), 177, 4)  ' CANTIDAD ASIGNADA
         Call REPLA(Y$, MKS$(PORCOVE), 181, 4)  ' POCEN.COMIS.VENDEDOR
         Call REPLA(Y$, String$(2, 0), 185, 2) ' FILLER
         Call REPLA(Y$, String$(6, 0), 187, 6) ' ANULACION (F/H/U/T)
         '
         Call REGAPP("MAVENDID", Y$)
         '
         X$ = REGLEIDO$("PEDDETA", REGPEDA&)
         Call REPLA(X$, Chr$(2), 27, 1)
         Call REPLA(X$, MKS$(CanAsig + CVS(Mid$(X$, 75, 4))), 75, 4)
         Call GRAREG("PEDDETA", X$, REGPEDA&)
         '
         X$ = REGLEIDO$("PEDDETP", REGPEDP&)
         Call REPLA(X$, Chr$(2), 27, 1)
         Call REPLA(X$, MKS$(CanAsig + CVS(Mid$(X$, 75, 4))), 75, 4)
         Call GRAREG("PEDDETP", X$, REGPEDP&)
         '
         If MODOEMI% = 2 Then    ' blanquea ODESEMI
            X$ = REGLEIDO$("ODESEMI", REGODES&)
            CAPENPE = CVS(Mid$(X$, 13, 4)) - CanAsig: If CAPENPE < 0 Then CAPENPE = 0
            Call REPLA(X$, MKS$(CAPENPE), 13, 4)
            CAPENAS = CVS(Mid$(X$, 17, 4)) - CanAsig: If CAPENAS < 0 Then CAPENAS = 0
            Call REPLA(X$, MKS$(CAPENAS), 17, 4)
            Call GRAREG("ODESEMI", Y$, REGODES&)
         End If
         '
         OKEY% = 1
         '
      End If
      '
    Next KU&
    '
98  Call CIERRARCH("MAVENDID")
    Call CIERRARCH("PEDDETA")
    Call CIERRARCH("PEDDETP")
    Call CIERRARCH("ODESEMI")
    Call CIERRARCH("*.*")
    Call BAJALDISCO
    '
    NODES& = NODESPA&
    '
99  Call CIERRARCH("*.*")
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
    If CONTROL$("ORDESPA", "SUPRINT") = "SI" Then
      Call BLOQARCH("IPEDPRIN")
    End If
    Call BLOQARCH("PDOCSPL")
    Call BLOQARCH("PDOCLST")
    '
    AVAHOJA$ = TRIM$(CONTROL$("FORDESPA", "AVAHOJA"))
    Screen.MousePointer = 11
    '
30  CODFOR$ = TRIM$(UCase$(FORIPRE$))
    MARIZW = XVALO(ATRIFORM$(CODFOR$, "MARIZQMM"))
    MARSUW = XVALO(ATRIFORM$(CODFOR$, "MARSUPMM"))
    '
    INTERLI = XVALO(ATRIFORM$(CODFOR$, "INTERLIN"))
    If INTERLI < 1 Then
         Call COMUNI("Error de Configuración de Interlinea\en Formulario '" + CODFOR$ + ".FRM'.")
         INTERLI = 1
    End If
    '
    DESMM = XVALO(ATRIFORM$(CODFOR$, "DETDESMM"))
    HASMM = XVALO(ATRIFORM$(CODFOR$, "DETHASMM"))
    '
    CACOPI% = XVALO(ATRIFORM$(CODFOR$, "CAN-COP"))
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
    RFF$ = RECFILT$("MAVENDID", 24, MKS$(NODES&))
    '
10  FINUBI& = ULTREG&("UBISTOCK")
    FEVEN% = 32767
    For UI& = 1 To Len(RFF$) Step 4
       v& = CVS(Mid$(RFF$, UI&, 4))
       X$ = REGLEIDO$("MAVENDID", v&)
       If YAENCA% = 0 Then
          NPED& = CVS(Mid$(X$, 93, 4))
          NC% = CVI(Left$(X$, 2))
          GoSub 257
          YAENCA% = 1
       End If
       '
       Ci% = CVI(Mid$(X$, 5, 2))
       Asig = CVS(Mid$(X$, 177, 4))
       '
       If Asig > 0 Then
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
         FENTRESOL% = CVI(Mid$(X$, 125, 2))
         If FENTRESOL% > 0 Then
            If FENTRESOL% < FEVEN% Then
               FEVEN% = FENTRESOL%
            End If
         End If
         FEX = FENTRESOL%
         '
         Call PRINTFORWIN(CODEXT$(Ci%), "COD-MAT", 0, YAC)
         Call PRINTFORWIN(DESCRIT$(Ci%), "DES-MAT", 0, YAC)
         Call PRINTFORWIN(TRIM$(Str$(Asig)), "CANTIDAD", 0, YAC)
         Call PRINTFORWIN(FECHATEX$(FEX), "FECH-LAN", 0, YAC)
         Call PRINTFORWIN("..........", "REF-MAT1", 0, YAC)
         '
20       KKK% = 0
         RGXX$ = RECFILT$("UBISTOCK", 1, MKI$(Ci%))
         For W& = 1 To Len(RGXX$) Step 4
           REG& = CVS(Mid$(RGXX$, W&, 4))
           UB$ = REGLEIDO$("UBISTOCK", REG&)
           UB2$ = Mid$(UB$, 3, 24)
           If TRIM$(UB2$) <> "" Then
             If KKK% > 0 Then
               Call PRINTFORWIN(Space$(16), "COD-MAT", 0, YAC)
               Call PRINTFORWIN(Space$(64), "DES-MAT", 0, YAC)
               Call PRINTFORWIN(Space$(16), "CANTIDAD", 0, YAC)
               Call PRINTFORWIN(Space$(16), "FECH-LAN", 0, YAC)
               Call PRINTFORWIN(Space$(16), "REF-MAT1", 0, YAC)
               'Printer.Print Space$(76);
               'Call SETSPOOL("PRINT", Space$(76))
             End If
             KKK% = KKK% + 1
             If KKK% > 1 Then
               YAC = YAC + INTERLI
             End If
             Call PRINTFORWIN(Left$(UB2$, 16), "REF-MAT2", 0, YAC)
           End If
         Next W&
       End If
       '
    Next UI&
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
    DESCRIDOC$ = DESCRIDOC$ + " - " + TRIM$(rasocli$(NC%))
    FEX = FF%
    DESCRIDOC$ = DESCRIDOC$ + " - " + FECHATEX$(FEX)
    '
    CANCELPRINT% = 0
    If CONTROL$("ORDESPA", "SUPRINT") = "SI" Then
       CANCELPRINT% = 1
    End If
    Call SAVESPOOL(DESCRIDOC$, OKX%)
    Call CIERRARCH("PDOCSPL")
    Call CIERRARCH("PDOCLST")
    '
    If CANCELPRINT% > 0 Then
       NDDX$ = TRIM$(Str$(NODES&))
       While Len(NDDX) < 6: NDDX$ = " " + NDDX$: Wend
       NDDX$ = "OD" + NDDX$
       PENPRIN$ = Space$(64)
       Call REPLA(PENPRIN$, NDDX$, 1, 8)
       Call REPLA(PENPRIN$, rasocli$(NC%), 9, 48)
       Call REPLA(PENPRIN$, MKI$(0), 57, 2)
       Call REPLA(PENPRIN$, Space$(6), 59, 6)
       '
       Call REGAPP("IPEDPRIN", PENPRIN$)
       Call CIERRARCH("IPEDPRIN")
       Call CIERRARCH("*.*")
       Call BAJALDISCO
    End If
    '
    Call CIERRARCH("IPEDPRIN")
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
    Call PRINTFORWIN(TRIM$(rasocli$(NC%)) + " (" + TRIM$(Str$(NC%)) + ")", "RASO-CLI", 0, 0)
    Call PRINTFORWIN(DOMICLI$(NC%), "DOMI-CLI", 0, 0)
    Call PRINTFORWIN(TRIM$(CPOSCLI$(NC%) + " " + LOCACLI$(NC%)), "LOCA-CLI", 0, 0)
    FF% = HOY(HOS$)
    Call PRINTFORWIN(HOS$, "FECH-EMI", 0, 0)
    Call PRINTFORWIN(TRIM$(Str$(NODES&)) + " / " + TRIM$(Str$(NUHOJ%)), "NUME-COM", 0, 0)
    Call PRINTFORWIN(TRIM$(Str$(NPED&)), "NRO-OCOM", 0, 0)
    '
    Call PRINTFORWIN(ZONACLI$(NC%), "PROV-CLI", 0, 0)
    Call PRINTFORWIN(DOMIENT$(NC%), "DOMI-TRA", 0, 0)
    Call PRINTFORWIN(LOCAENT$(NC%), "DESTINO", 0, 0)
    Call PRINTFORWIN(TRANENT$(NC%), "TRANSPOR", 0, 0)
    Call PRINTFORWIN(HORAENT$(NC%), "TIE-TOTA", 0, 0)
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
Sub CONODESCLI()
   '
   Dim NUPEDSEL&(1024)
   '
10 Call SELECHO("DEUDOR1")
   NCXI% = XVALO(TRIM$(VALACT1$("DEUDOR1")))
   If NCXI% < 1 Then Exit Sub
   '
   FORSELPED.LIPEDPE.Clear
   RPEDX$ = RECFILT$("MAVENDID", 19, MKI$(NCXI%))
   NRODESA& = 0
   FIN& = Len(RPEDX$)
   For UU& = 1 To FIN& Step 4
     Call TRACE(20, UU&, FIN&)
     U& = CVS(Mid$(RPEDX$, UU&, 4))
     X$ = REGLEIDO$("MAVENDID", U&)
     If CVI(Mid$(X$, 1, 2)) = NCXI% Then   ' coincide cliente
       If Asc(Mid$(X$, 63, 1)) = 1 Then    ' STATUS = 1
         CAN = CVS(Mid$(X$, 177, 4))
         CADESPACH = CVS(Mid$(X$, 225, 4))
         '
         If CADESPACH >= 0.05 Then         ' MARCA COMO CUMPLIDA
           Call REPLA(X$, Chr$(2), 63, 1)
           Call REPLA(X$, MKI$(0), 65, 2)
           Call REPLA(X$, MKI$(0), 67, 2)
           Call GRAREG("MAVENDID", X$, U&)
           Call CIERRARCH("MAVENDID")
           GoTo 19
         End If
         '
         CAPENDE = CAN - CADESPACH
         If CAN >= 0.05 Then
           NRODES& = CVS(Mid$(X$, 87, 4))
           If NRODES& <> NRODESA& Then
             TTXX$ = AJUSTI$("OD-" + TRIM$(CSTRING$(MKS$(NRODES&), 3, 8, 0, 2)), 11)
             FEX = CVI(Mid$(X$, 125, 2))
             TTXX$ = TTXX$ + "   Sol." + FECHATEX$(FEX)
             FORSELPED.LIPEDPE.AddItem TTXX$
           End If
           NRODESA& = NRODES&
           CAPED% = CAPED% + 1
           FENTRE% = CVI(Mid$(X$, 125, 2))
           If FENTRE% < FINF Then FINF = FENTRE%
           If FENTRE% > FSUP Then FSUP = FENTRE%
         End If
       End If
     End If
19 Next UU&
   '
   If FORSELPED.LIPEDPE.ListCount < 1 Then
     Call COMUNI("No hay Ordenes de Despacho en Curso\para " + rasocli$(NCXI%) + ".")
     GoTo 10
   End If
   '
   FORSELPED.Show 1
   '
   CAPESEL% = 0
   PRESEFA.Label4 = ""
   For U& = 1 To FORSELPED.LIPEDPE.ListCount
     If FORSELPED.LIPEDPE.Selected(U& - 1) = True Then
       CAPESEL% = CAPESEL% + 1
       TTXX$ = FORSELPED.LIPEDPE.List(U& - 1)
       NUPEDSEL&(CAPESEL%) = XVALO(Left$(TTXX$, 10))
       If Left$(TTXX$, 3) = "OD-" Then
           NUPEDSEL&(CAPESEL%) = XVALO(Mid$(TTXX$, 4, 10))
           If PRESEFA.Label4.Caption <> "" Then
             PRESEFA.Label4.Caption = PRESEFA.Label4.Caption + ", "
           End If
           PRESEFA.Label4.Caption = TRIM$(PRESEFA.Label4.Caption + TRIM$(Mid$(TTXX$, 4, 7)))
         ElseIf Left$(TTXX$, 3) = "RT-" Then
           NUPEDSEL&(CAPESEL%) = (-1) * XVALO(Mid$(TTXX$, 4, 10))
       End If
     End If
   Next U&
   '
   If CAPESEL% < 1 Then GoTo 10
   '
   LPRE% = LIPRCLI%(NCXI%)
   If LPRE% < 1 Then LPRE% = 1
   ARCHIPRE$ = TRIM$(Str$(LPRE%)): LIPRACTI% = LPRE%
   While Len(ARCHIPRE$) < 3: ARCHIPRE$ = "0" + ARCHIPRE$: Wend
   ARCHIPRE$ = "LIPRE" + ARCHIPRE$
   '
   CAITCAR% = 0
   PRESEFA.List1.Clear
   PRESEFA.Label1.Caption = ""
   FIN& = Len(RPEDX$)
   SUMATO# = 0
   For UU& = 1 To Len(RPEDX$) Step 4
     U& = CVS(Mid$(RPEDX$, UU&, 4))
     X$ = REGLEIDO$("MAVENDID", U&)
     REGPEDA& = CVS(Mid$(X1$, 99, 4))
     NODES& = CVS(Mid$(X$, 87, 4))
     '
     For UI% = 1 To CAPESEL%
       If NUPEDSEL&(UI%) = NODES& Then GoTo 20
     Next UI%
     GoTo 39
     '
20   CAN = CVS(Mid$(X$, 177, 4)) - CVS(Mid$(X$, 225, 4))
     If CAN >= 0.05 Then
       Ci% = CVI(Mid$(X$, 5, 2))
       Z$ = Space$(128)
       Call REPLA(Z$, CSTRING$(MKS$(CAN), 3, 5, 0, 2), 1, 5)  ' cantidad
       Call REPLA(Z$, CODEXT$(Ci%), 8, 12)
       Call REPLA(Z$, DESCRIT$(Ci%), 21, 40)
         PRELI = CVD(Mid$(X$, 127, 8))
         If PRELI < 0.0001 Then PRELI = PRELISTA#(ARCHIPRE$, Ci%)
         PRECIO = CVD(Mid$(X$, 151, 8))
         If PRECIO < 0.00005 Then PRECIO = PRELI
         If PRELI > 0.005 And PRELI < PRECIO Then PRELI = PRECIO
       Call REPLA(Z$, CSTRING$(MKS$(PRECIO), 3, 10, 4, 2), 63, 10)
         IMPO# = CDbl(PRECIO) * CAN
       Call REPLA(Z$, CSTRING$(MKD$(IMPO#), 3, 12, 2, 2), 73, 12)
       SUMATO# = SUMATO# + IMPO#
       PRESEFA.List1.AddItem Z$
     End If
     '
39 Next UU&
   '
   PRESEFA.Caption = rasocli$(NCXI%)
   PRESEFA.Label1.Caption = CSTRING$(MKD$(SUMATO#), 3, 12, 2, 2)
   '
   PRESEFA.Show 1
   GoTo 10
   '
End Sub


