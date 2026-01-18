Attribute VB_Name = "FACTU2002"
Global A$(8)
' Public TOTAFAC#(3, 10)
Public CONTROFA$
Public FORMAIMPRE$
'
Public ARCHIFAC$, ARCHIDEU$, ARCHICLI$, ARCHIMAT$, ARCHISTO$
'
Public TDFAC$
Public TIDOC%
Public CAFOR%
Public NRO, NROCON&
Public TIPO%, VA%
Public PIVA%
Public ALICU, ALI1, ALI2

Dim NUPEDSEL&(1024)
Dim RF$(4)
Dim DATCLI$(32)
'
Dim POI%(10)
Dim T$(2)
Dim RG&(16)
'
Dim TPX%(8), TIPCOMS$(8), VTIPS$(8)
Dim TIFORM%
Dim FORVENTA$
Dim FORIMPRE$
Dim VENCADET$
Dim VEPIEDET$
Dim MAXITEMS%
Dim UREN%, RG0%
Dim RTO$, TRXP$
Dim CAPED%
'
Dim TIPCOM$
Dim TICA
Dim FFI%, FVI%
Dim PORDES, MODESCU%, VERISTO%
Dim FACIVA
Dim CATIVAS$
'
Dim SELORDESPA%
'
Dim EXPOR%, FINFACEX$
Dim SOLOFAC%


Sub INDIPEDI()
   '
   ' este modulo revisa y restablece los punteros
   ' entre los archivos "ORDESPA", "PEDDETA" Y "PEDDETP"
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
   Screen.MousePointer = 11
''Call FRESHALL
   FIN& = ULTREG&("PEDDETP")
   For I& = 1 To FIN&
       Call TRACE(20, I&, FIN&)
       X$ = REGLEIDO$("PEDDETP", I&)
       If VERIPEDPE%(X$) <> 1 Then
           Call FINAL("")
       End If
       REGPEDA& = CVS(Mid$(X$, 109, 4))
       RPP% = I&
       Y$ = REGLEIDO$("PEDDETA", REGPEDA&)
       Call REPLA(Y$, MKI$(RPP%), 31, 2)
       Call GRAREG("PEDDETA", Y$, REGPEDA&)
   Next I&
   Call CIERRARCH("PEDDETA")
   '
   UREPA& = ULTREG&("PEDDETA")
   J& = 0
   FIN& = ULTREG&("ORDESPA")
   For I& = 1 To FIN&
       Call TRACE(20, I&, FIN&)
       X$ = REGLEIDO$("ORDESPA", I&)
       If CVS(Left$(X$, 4)) > 0.5 Then
          If CVS(Mid$(X$, 7, 4)) > 0.5 Then
             J& = J& + 1
             Call GRAREG("ORDESPA", X$, J&)
             REGPEDA& = CVS(Mid$(X$, 111, 4))
             If REGPEDA& > 0 Then
                If REGPEDA& <= UREPA& Then
                   Y$ = REGLEIDO$("PEDDETA", REGPEDA&)
                   If Left$(Y$, 30) = Mid$(X$, 7, 30) Then
                      If Mid$(Y$, 33, 42) = Mid$(X$, 37, 42) Then
                         GoTo 99
                      End If
                   End If
                End If
             End If
             '
             FIN1& = ULTREG&("PEDDETA")
             For K& = FIN1& To 1 Step -1
                Call TRACE(20, FIN1& - K& + 1, FIN1&)
                Y$ = REGLEIDO$("PEDDETA", K&)
                If Left$(Y$, 30) = Mid$(X$, 7, 30) Then
                   If Mid$(Y$, 33, 42) = Mid$(X$, 37, 42) Then
                      REGPEDA& = K&
                      Call REPLA(X$, MKS$(REGPEDA&), 111, 4)
                      Call GRAREG("ORDESPA", X$, J&)
                      GoTo 99
                   End If
                End If
             Next K&
             '
             NODES& = CVS(Left$(X$, 4))
             Call COMUNI("No se ha Encontrado el Pedido Original\Correspondiente a la Orden de Despacho " + TRIM$(Str$(NODES&)) + ".\  \FLEXOFT Anulará Esta O.Despacho.")
             Call GRAREG("ORDESPA", String$(128, 0), J&)
             '
          End If
       End If
       '
99 Next I&
   '
   Call SETULTREG("ORDESPA", J&)
   Call CIERRARCH("ORDESPA")        ' fuerza indexacion
   Screen.MousePointer = 1
   '
End Sub
'



Sub CARDACLJ()

    NC% = Val(FORMRAC.Text1.TEXT)
    If NC% > 0 Then
        FORMRAC.Text2.TEXT = RASOCLI$(NC%)
        FORMRAC.Text3.TEXT = DOMICLI$(NC%)
        FORMRAC.Text4.TEXT = CPOSCLI$(NC%)
        FORMRAC.Text5.TEXT = LOCACLI$(NC%)
        FORMRAC.Text6.TEXT = PIVACLI%(NC%)
        FORMRAC.Text7.TEXT = POSIVAC$(NC%)
        FORMRAC.Text8.TEXT = CUITCLI$(NC%)
        '
        FORMRAC.TEXT21(0).TEXT = Str$(LIPRCLI%(NC%))
        FORMRAC.TEXT21(1).TEXT = Str$(CPAGCLI%(NC%))
        FORMRAC.TEXT21(2).TEXT = Str$(VENDCLI%(NC%))
        '
        If Val(FORMRAC.TEXT21(2).TEXT) < 1 Then
           FORMRAC.Eco21(2).TEXT = "GERENCIA"
        End If
        '
        PIVA% = Val(FORMRAC.Text6.TEXT)
        If PIVA% < 0 Or PIVA% > 7 Then
            FORMRAC.Text6.SetFocus
            Screen.MousePointer = 1
            Exit Sub
        End If
        '
        FORMRAC.PORDESCU.TEXT = DESCUCLI$(NC%)
        '
        GoSub 81                        ' determina posicion IVA y CATIVAS$
        FORMRAC.PORIVA.TEXT = ALICU
        '
1050    LU$ = LUDAT$
        CLI$ = RASOCLI$(NC%)
        '
    End If
90  Screen.MousePointer = 1
    Exit Sub
    '
    '
81  FACIVA = 1
    CATIVAS$ = "A"
    VA% = 1
    ALICU = ALI1
    RNI% = 0
    '
    If PIVA% = 0 Or PIVA% = 3 Or PIVA% = 4 Or PIVA% = 6 Then
          FACIVA = 1 + ALICU / 100
          CATIVAS$ = "B"
          VA% = 2
          '
        ElseIf PIVA% = 2 Then
          ALICU = ALI1 + ALI2
          RNI% = 1
          '
        ElseIf PIVA% = 5 Or PIVA% = 7 Then
          ALICU = 0
          CATIVAS$ = " "
          VA% = 0
          If LETREXPO$ = "A" Then
              VA% = 1
              CATIVAS$ = "A"
          End If
    End If
    '
Return
'
End Sub



Sub RECARFAC()
    '
    LU$ = LUDAT$
    NCC$ = TRIM$(FORMRAC.Text1.TEXT)
    If Val(NCC$) < 1 Then
        FORMRAC.Text1.SetFocus
        Exit Sub
    End If
    '
    FFI1 = FECHANUM(TRIM$(FORMRAC.Text9.TEXT))
    FFI% = FFI1
    '
    NC% = Val(NCC$)
502 PIVA% = Val(FORMRAC.Text6.TEXT)
    If ECOARCH$("CAPOSIV", Chr$(PIVA%)) = "" Then
        FORMRAC.Text1.TEXT = "1"
        FORMRAC.Text6.SetFocus
        Exit Sub
    End If
    '
    If PIVA% <> 0 Then
       If PIVA% <> 5 Then
          CUITT$ = TRIM$(FORMRAC.Text8.TEXT)
          If CUITT$ = "" Or VALICUIT%(CUITT$) <> 1 Then
             Call MENSERR(24, "Imposible Facturar a esta Cuenta.\  \Número de CUIT no es Válido.")
             FORMRAC.Text1.SetFocus
             Exit Sub
          End If
       End If
    End If
    '
    If TRIM$(DOMICLI$(NC%)) = "" Then
       Call MENSERR(24, "Imposible Facturar a esta Cuenta.\  \Falta Domicilio Legal.")
       FORMRAC.Text1.SetFocus
       Exit Sub
    End If
    '
    LPRE% = Val(FORMRAC.TEXT21(0).TEXT)
    If ECOARCH("LISTAS", MKI$(LPRE%)) = "" Then
        FORMRAC.TEXT21(0).SetFocus
        Exit Sub
    End If
    '
    CPAG% = Val(FORMRAC.TEXT21(1).TEXT)
    If ECOARCH("CONPAG", MKI$(CPAG%)) = "" Then
        FORMRAC.TEXT21(1).SetFocus
        Exit Sub
    End If
    '
    VENDE% = Val(FORMRAC.TEXT21(2).TEXT)
    If VENDE% > 0 Then
        If ECOARCH("VENDEDOR", Chr$(VENDE%)) = "" Then
            FORMRAC.TEXT21(2).SetFocus
            Exit Sub
        End If
    End If
    '
    PROV$ = NPROCLI$(NC%)          ' numero de proveedor
    '
    Call CLOSEFORMPRIN("")
    FORVENTA$ = Mid$(CONTROFA$, 81, 8)
    FORIMPRE$ = Mid$(CONTROFA$, 89, 8)
    If PIVA% = 1 Or PIVA% = 2 Then
              FORVENTA$ = Mid$(CONTROFA$, 65, 8)
              FORIMPRE$ = Mid$(CONTROFA$, 73, 8)
            ElseIf PIVA% = 5 Or PIVA% = 7 Then
              FORVENTA$ = Mid$(CONTROFA$, 97, 8)
              FORIMPRE$ = Mid$(CONTROFA$, 105, 8)
    End If
    VENCADET$ = TRIM$(Mid$(CONTROFA$, 113, 8))
    VEPIEDET$ = TRIM$(Mid$(CONTROFA$, 121, 8))
    VERISTO% = Asc(Mid$(CONTROFA$, 129, 1))
    MODESCU% = Asc(Mid$(CONTROFA$, 130, 1))
    '
    FORVENTA$ = TRIM$(FORVENTA$)
    FORIMPRE$ = TRIM$(FORIMPRE$)
    If FORVENTA$ = "" Then
         Call COMUNI("Falta Definir Controlador\de Ventana de Carga de Factura")
         Call FINAL("")
      ElseIf EXISTE%(FORVENTA$ + ".PRF") <> 1 Then
         Call COMUNI("Falta Definir Controlador\de Ventana de Carga de Factura")
         Call FINAL("")
    End If
    '
    If FORIMPRE$ = "" Then
         Call COMUNI("Falta Definir Controlador\de Impresion de Documento")
         Call FINAL("")
       ElseIf EXISTE%(LU$ + FORIMPRE$ + ".FRM") <> 1 Then
         Call COMUNI("Falta Definir Controlador\de Impresion de Documento")
         Call FINAL("")
    End If
    '
    MAXITEMS% = Val(ATRIFORM$(FORIMPRE$, "MAX-ITEM"))
    CONTIHOJA% = 0: If TRIM$(ATRIFORM$(FORIMPRE$, "MAX-ITEM")) = "HOJACONT" Then CONTIHOJA% = 1
    Call CLOSEFORMPRIN("")
    '
    FORMRAC.BOTNEXT.Enabled = False
    EXPOR% = 0: If PIVA% = 5 And TIPO% = 2 Then EXPOR% = 1

'*********************************************************** fecha de factura

1200  If TIDOC% <> 1 Then
          CAPED% = 0
      End If

'************************************** determina fecha de pedidos pendientes

      Screen.MousePointer = 11
      FINF = 32767: FSUP = 0: CAPED% = 0
      '
1210  Z$ = ""
      Call CIERRARCH("!CARDEFAC")
      Call BLANARCH("!CARDEFAC")
      '
1211  TOTAFAC#(1, 9) = 0
      '
      ARCHIPRE$ = TRIM$(Str$(LPRE%))
      While Len(ARCHIPRE$) < 3: ARCHIPRE$ = "0" + ARCHIPRE$: Wend
      ARCHIPRE$ = "LIPRE" + ARCHIPRE$
      '
1215  ATRI$ = "/COLORPAN=B-N/PS/TITUPAN=" + VARIFA$
      ATRI$ = ATRI$ + "/BORDESUP=7/BORDEINF=20/BORDEIZQ=4"
      If InStr(UCase$(FORVENTA$), "LIB") < 1 Then ATRI$ = ATRI$ + "/ARCHECO(2)=" + ARCHIPRE$
      ' lo anterior para permitir facturacion libre
      XI% = VENTABU%(FORVENTA$ + ATRI$)
      If XI% < 0 Then
          GoTo 3500
          FORMRAC.BOTNEXT.Enabled = True
          FORMRAC.Text1.SetFocus
          Exit Sub
      End If
      '
      FORMRAC.Picture1.Refresh
      Call CIERRARCH("*.*")
      '
      If ULTREG&("!CARDEFAC") < 1 Then GoTo 3500
      '
      Screen.MousePointer = 11
      ATX% = 0: CIANT% = 0: FIN& = ULTREG&("!CARDEFAC")
      '
      For U& = 1 To FIN&
         '
         Call TRACE(20, U&, FIN&)
         '
         X$ = REGLEIDO$("!CARDEFAC", U&)
         CI% = CVI(Mid$(X$, 5, 2))
         CAN = CVS(Left$(X$, 4))
         '
         If CVS(Mid$(X$, 509, 4)) < 0 Then
            REREPEN& = Abs(CVS(Mid$(X$, 509, 4)))
            Z$ = REGLEIDO$("REMPEN", REREPEN&)
            If CVS(Mid$(Z$, 27, 4)) <> CAN Then
                Call COMUNI("Operacion Incorrecta.\En Facturacion por Remito Pendiente\no Pueden Modificarse las Cantidades.")
                Call CIERRARCH("*.*")
                GoTo 1200
            End If
         End If
         '
         PRECIOD# = CDbl(CVS(Mid$(X$, 421, 4)))
         '
         If PRECIOD# < 0.00005 Then
           If CAN > 0.05 Then
             If CI% > 0 Then
               If CI% <= NUMAS% Then
                 PRECIOD# = PRELISTA#(ARCHIPRE$, CI%)
                 Call REPLA(X$, MKS$(PRECIOD#), 421, 4)
                 Call REPLA(X$, MKS$(PRECIOD#), 429, 4)
                 DE$ = DESCRIT$(CI%)
                 Call REPLA(X$, DE$, 7, 64)
               End If
             End If
           End If
         End If
         '
         PRELIXL = PRELISTA#(ARCHIPRE$, CI%)
         If PRELIXL > 0.005 Then
           If CVS(Mid$(X$, 429, 4)) > 2 * PRELIXL Then
             Call MENSERR(24, "Precio Unitario no Procesable")
             Call REPLA(X$, MKS$(PRELIXL), 429, 4)
             Call GRAREG("!CARDEFAC", X$, U&)
           End If
         End If
         '
'MsgBox ARCHIPRE$ & "   " & CI%
         NORI$ = NUORIG$(ARCHIPRE$, CI%)
'MsgBox ARCHIPRE$ & "   " & CI% & "   " & NORI$
         Call REPLA(X$, NORI$, 339, 16)
         '
         OCOM$ = NUOCOM$(NC%, CI%)
         Call REPLA(X$, OCOM$, 355, 16)
         '
1220     PREUNI# = CDbl(CVS(Mid$(X$, 429, 4)))
         LARGO = CVS(Mid$(X$, 303, 4))
         ANCHO = CVS(Mid$(X$, 307, 4))
         ALTO = CVS(Mid$(X$, 311, 4))
         '
         CX% = 0: RF$(1) = "": RF$(2) = "": RF$(3) = ""
         If LARGO * ANCHO * ALTO > 10 Then
               CX% = CX% + 1
               RF$(CX%) = TRIM$(Str$(LARGO)) + " mm x " + TRIM$(Str$(ANCHO)) + " mm x " + TRIM$(Str$(ALTO)) + " mm"
               PREUNI# = CDbl(PRECIOD# * (LARGO / 1000) * (ANCHO / 1000) * (ALTO / 1000))
         End If
         '
         Call REPLA(X$, MKS$(PREUNI#), 429, 4)
         TOT# = CDbl(CAN * PREUNI#)
         If TOT# > 0.005 Then Call REPLA(X$, MKD$(TOT#), 433, 8)
         '
         TXT$ = ""
         If NORI$ <> "" Then TXT$ = "Nro.Orig.: " + NORI$
         If OCOM$ <> "" Then TXT$ = TXT$ + "    O.Compra: " + OCOM$
         NORI$ = "": OCOM$ = ""
         TXT$ = TRIM$(TXT$)
         If TXT$ <> "" Then
             CX% = CX% + 1
             RF$(CX%) = TXT$
         End If
         '
         If CI% > 0 Then
             TXT$ = TRIM$(DESPACHO$(CI%))
             If TXT$ <> "" Then
                 CX% = CX% + 1
                 RF$(CX%) = "Desp: " + TXT$
             End If
         End If
         '
         For CY% = 1 To CX%
              Call REPLA(X$, RF$(CY%), 11 + 64 * CY%, 64)
         Next CY%
         '
         Call GRAREG("!CARDEFAC", X$, U&)
         '
         If CVI(Mid$(X$, 5, 2)) > 0 Then                   ' verificar precio
             If CVS(Mid$(X$, 429, 4)) < 0.00005 Then
                 PREINFO = CVS(Mid$(X$, 421, 4))
                 If PREINFO >= 0.005 Then
                     Call REPLA(X$, MKS$(PREINFO), 429, 4)
                     Call GRAREG("!CARDEFAC", X$, U&)
                 End If
                 If CVS(Mid$(X$, 429, 4)) < 0.00005 Then
                     ATX% = 1
                 End If
             End If
         End If
         '
      Next U&
      Screen.MousePointer = 1
      '
      X$ = "": DE$ = "": TXT$ = "": RF$(1) = "": RF$(2) = "": RF$(3) = ""
      If ATX% > 0 Then
          Call CIERRARCH("*.*")
          Call COMUNI("Factura no Procesable\hay Items sin Precio.")
          GoTo 1215
      End If
      '
1310  Call CIERRARCH("*.*")
      GoSub 8200                          ' calcular IVA de cada item
      '
      GoSub 8700                                        ' presentar totales
      TITU$ = "FORMATO NORMAL": If TIFORM% = 2 Then TITU$ = "FORMATO RESUMEN" Else If TIFORM% = 3 Then TITU$ = "FORMATO TEXTO LIBRE"
1330  '
      FACIVA = 1
      CATIVAS$ = "A"
      VA% = 1
      If PIVA% = 0 Or PIVA% = 3 Or PIVA% = 4 Or PIVA% = 6 Then
           FACIVA = 1 + ALICU / 100
           CATIVAS$ = "B"
           VA% = 2
           '
          ElseIf PIVA% = 2 Then
           ALICU = ALI1 + ALI2
           RNI% = 1
           '
          ElseIf PIVA% = 5 Or PIVA% = 7 Then
           ALICU = 0
           CATIVAS$ = " "
           VA% = 0
           If LETREXPO$ = "A" Then
              VA% = 1
              CATIVAS$ = "A"
           End If
      End If
      FOPRES$ = "PRESEFAC": If FACIVA <> 1 Then FOPRES$ = "PRESFACB"
      ATRI$ = "/BORDEIZQ=3/BORDESUP=5/BORDEINF=21"
      FORMRAC.Picture1.Visible = False
      XI% = VENTABU%(FOPRES$ + ATRI$ + "/COLORPAN=B-N/NC/TITUPAN=" + TITU$)
      FORMRAC.Picture1.Visible = True
      If XI% < 0 Then
          GoTo 3500
          FORMRAC.BOTNEXT.Enabled = True
          FORMRAC.Text1.SetFocus
          Exit Sub
      End If
      '
      ATX% = 0
      FIN& = ULTREG&("!CARDEFAC")
      '
      For U& = 1 To FIN&
         X$ = REGLEIDO$("!CARDEFAC", U&)
         CAN = CVS(Left$(X$, 4))
         PREUNI# = CDbl(CVS(Mid$(X$, 429, 4)))
         If CAN > 0 Then
           If PREUNI# > 0 Then
             TOT# = CDbl(CAN * PREUNI#)
             If TOT# <> CVD(Mid$(X$, 433, 8)) Then
               ATX% = 1
               Call REPLA(X$, MKD$(TOT#), 433, 8)
               Call GRAREG("!CARDEFAC", X$, U&)
             End If
           End If
         End If
      Next U&
      '
      If ATX% <> 0 Then GoTo 1310                ' volver a presentar factura
'
'****************************************************************************
'
2010  ' pedir porcentaje de descuento
        '
        If TIDOC% = 1 Or (TIDOC% = 3 And InStr(UCase$(TDFAC$), "DEVOLUC") > 0) Then                   ' solo facturas
            If DESCUPRE% = 0 Then
                DESC$ = TRIM$(FORMRAC.PORDESCU.TEXT)
                Call DESCUEN(PORDES, DESC$)
                '
2011            If PORDES > 0 Then
                    COEF = 1 - PORDES / 100
                    If MODESCU% = 1 Then COEF = 1
                    For U& = 1 To ULTREG&("!CARDEFAC")
                       X$ = REGLEIDO$("!CARDEFAC", U&)
                       CAN = CVS(Left$(X$, 4))
                       PREUNI# = (CDbl(CVS(Mid$(X$, 429, 4)))) * COEF
                       Call REPLA(X$, MKS$(PREUNI#), 429, 4)
                       TOT# = CDbl(CAN * PREUNI#)
                       If CVI(Mid$(X$, 5, 2)) < 1 Then
                           TOT# = CVD(Mid$(X$, 433, 8)) * COEF
                       End If
                       '
                       IVAITEM# = TOT# * (ALICU / 100)
                       BRUITEM# = TOT# + IVAITEM#
                       '
                       Call REPLA(X$, MKD$(TOT#), 433, 8)
                       Call REPLA(X$, MKS$(ALICU), 441, 4)
                       Call REPLA(X$, MKD$(IVAITEM#), 445, 8)
                       Call REPLA(X$, MKD$(BRUITEM#), 453, 8)
                       Call GRAREG("!CARDEFAC", X$, U&)
                    Next U&
                End If
                '
            End If
        End If
        GoSub 8700
'****************************************************************************
'
2020 '  calculo percepcion ingresos brutos

        VVX% = 0: TASA1 = 0: TASA2 = 0
        '
        For IL& = 1 To ULTREG&("PERCIBRU")
             If CVI(Left$(REGLEIDO$("PERCIBRU", IL&), 2)) = NC% Then
                 TASA1 = CVS(Mid$(REGLEIDO$("PERCIBRU", IL&), 3, 4))
                 TASA2 = CVS(Mid$(REGLEIDO$("PERCIBRU", IL&), 7, 4))
                 GoTo 2022
             End If
        Next IL&
        GoTo 2025
        '
2022    If (Int(TOTAFAC#(1, 9))) < 0.01 Then
            DENOMI = 100 - PORDES: If DESCUPRE% = 1 Then DENOMI = 100
            If TOTAFAC#(1, 5) >= VALIM1# Then
                If TASA1 > 0.005 Then
                    TOTAFAC#(1, 9) = (Int(TOTAFAC#(1, 5) * TASA1 + 0.5)) / DENOMI
                    If Abs(TOTAFAC#(1, 9)) > 0.005 Then VVX% = 1000 ' para senalizar ingresos brutos
                    GoSub 8700
                    GoTo 2025
                End If
            End If
            '
            If TOTAFAC#(1, 5) > VALIM2# Then
                If TASA2 > 0.005 Then
                    TOTAFAC#(1, 9) = (Int(TOTAFAC#(1, 5) * TASA2 + 0.5)) / DENOMI
                    If Abs(TOTAFAC#(1, 9)) > 0.005 Then VVX% = 1000 ' para senalizar ingresos brutos
                    GoSub 8700
                    GoTo 2025
                End If
            End If
            '
        End If
        '
2025    If TIDOC% > 1 Then                          ' debitos y creditos
            If FACIVA = 1 Then                      ' solo facturas 'A'
                IVA# = TOTAFAC#(1, 6)
                'Call ENTRADA(21, 1, 66, 13, 14, 3, 2, 2, VDI%, IVA#, VDS$, CT%, VDC%)
                If ALICU > 0.5 Then TOTAFAC#(1, 6) = (Int(100 * IVA# * ALI1 / ALICU + 0.5)) / 100
                TOTAFAC#(1, 7) = IVA# - TOTAFAC#(1, 6)
                TOTAFAC#(1, 8) = TOTAFAC#(1, 5) + TOTAFAC#(1, 6) + TOTAFAC#(1, 7) + TOTAFAC#(1, 9) + TOTAFAC#(1, 10)
                'Call PRESENTA(22, 0, 66, 13, 13, 3, 2, 2, VDI%, TOTAFAC#(1, 8), VDS$, CT%)
            End If
        End If
        '
2028    For I% = 1 To 10
            TOTAFAC#(2, I%) = TOTAFAC#(1, I%) / TICA
        Next I%
'
'****************************************************************************
'
2200    If TIDOC% > 1 Then GoTo 2900
        '
        TRXP$ = OBJPLA$("TRANSPOR", TRXP$)
        '
2300    If VENCADET$ <> "" Then
           VTBX% = VENTABU%(VENCADET$ + "/PS")
        End If
        '
        If VEPIEDET$ <> "" Then
           VTBX% = VENTABU%(VEPIEDET$ + "/PS/NC")
        End If
        '
2900  ' aprobacion de factura
        '
        XI% = COMALTER%("Fin de Carga de " + TIPCOM$ + " " + TRIM$(Str$(NRO)) + "\\Aprobar\Corregir\Anular")
        ''Call FRESHALL
        FORMRAC.Show
        '
        On XI% GoTo 2910, 2905, 3500
        Beep
        GoTo 2900
        '
2905    If SELORDESPA% = 1 Then
           Call COMUNI("Correccion Imposible.\Vuelva a Cargar Factura.")
           GoTo 3500
        End If
        '
        For UV& = 1 To ULTREG&("!CARDEFAC")
           Z$ = REGLEIDO$("!CARDEFAC", UV&)
           U& = CVS(Mid$(Z$, 509, 4))
           CAN = CVS(Left$(Z$, 4))
           PRECIO = 0
           If U& > 0 Then
              If U& <= ULTREG&("PEDDETP") Then
                 X$ = REGLEIDO$("PEDDETP", U&)
                 PRECIO = CVS(Mid$(X$, 71, 4))
                 If PRECIO < 0.0001 Then PRECIO = CVS(Mid$(X$, 83, 4))
              End If
           End If
           '
           Call REPLA(Z$, MKS$(PRECIO), 421, 4)
           Call REPLA(Z$, MKS$(PRECIO), 429, 4)
           IMPO# = CDbl(PRECIO) * CAN
           Call REPLA(Z$, MKD$(IMPO#), 433, 8)
           Call GRAREG("!CARDEFAC", Z$, UV&)
        Next UV&
        '
        GoTo 1211
        '
2910  ' reserva numero de documento
        '
        GRAICON% = 0
        If TIDOC% > 1 Then
          If TIDOC% <> 3 Or InStr(UCase$(TDFAC$), "DEVOLUC") < 1 Then
           VDEF$ = "ND": If TIDOC% = 3 Then VDEF$ = "NC"
           VDEF$ = VDEF$ + String$(94, 0)
           Call REPLA(VDEF$, MKS$(NRO), 3, 4)
           Call REPLA(VDEF$, MKI$(NC%), 7, 2)
           Call REPLA(VDEF$, MKI$(FFI%), 9, 2)
           Call REPLA(VDEF$, MKD$(TOTAFAC#(1, 1)), 19, 8)
2915       XD$ = OBJPLA$("ICOVEN", VDEF$)
           If XD$ = "" Then GoTo 2918
           '
           TOIMPU# = 0
           For QQ% = 17 To 96 Step 10
              CUE% = CVI(Mid$(XD$, QQ%, 2))
              IMPO# = (Int(100 * CVD(Mid$(XD$, QQ% + 2, 8)) + 0.5)) / 100
              If CUE% > 0 Then
                 If ECOARCH$("CUECON", MKI$(CUE%)) <> "" Then
                    TOIMPU# = TOIMPU# + IMPO#
                 End If
              End If
           Next QQ%
           '
           If Abs(TOIMPU# - TOTAFAC#(1, 1)) >= 0.005 Then
               VDEF$ = XD$
               GoTo 2915
           End If
           '
           GRAICON% = 1
           '
          End If
        End If
        '
2918    FINFACEX$ = ""
GoTo 2920
        If EXPOR% = 1 Then
          FINFACEX$ = String$(68, 0)
          Call REPLA(FINFACEX$, MKD$(TOTAFAC#(1, 8)), 1, 8)
          Call REPLA(FINFACEX$, "AEREO       ", 9, 12)
          Call REPLA(FINFACEX$, MKD$(0), 21, 8)
          Call REPLA(FINFACEX$, "CFR Bs.As.  ", 29, 12)
          Call REPLA(FINFACEX$, MKD$(TOTAFAC#(1, 8)), 41, 8)
          Call REPLA(FINFACEX$, "DOLARES     ", 49, 12)
          Call REPLA(FINFACEX$, MKD$(0), 61, 8)
          '
          VDEF$ = FINFACEX$
2919      OBX$ = OBJPLA$("FINFACEXPO", VDEF$)
          If OBX$ = "" Then GoTo 2918
          '
          VDEF$ = OBX$
          TICAMBX# = CVD(Mid$(OBX$, 61, 8))
          If TICAMBX# < 0.1 Then
            Call MENSERR(24, "Ingrese Tipo de Cambio")
            GoTo 2919
          End If
          '
          VAFLE# = CVD(Mid$(OBX$, 21, 8))
          VAFOB# = CVD(Mid$(OBX$, 1, 8)) + VAFLE#
          If Abs(VAFOB# - CVD(Mid$(OBX$, 41, 8))) > 0.005 Then
            Call REPLA(VDEF$, MKD$(VAFOB#), 41, 8)
            GoTo 2919
          End If
          '
          FINFACEX$ = OBX$
          TOTAFAC#(1, 9) = VAFLE#
          TOTAFAC#(1, 8) = VAFOB#
          '
        End If
        '
        If EXPOR% = 1 Then
          For KKI1% = 0 To 10
            TOTAFAC#(1, KKI1%) = TOTAFAC#(1, KKI1%) * TICAMBX#
          Next KKI1%
        End If
        '
        If EXPOR% = 1 Then
          For KKI1% = 0 To 10
            TOTAFAC#(1, KKI1%) = TOTAFAC#(1, KKI1%) / TICAMBX#
          Next KKI1%
        End If
        '
2920    Call CARIMPRE(CCI$())
        '
        Call CARSETUP("NEW")
        Call CARSETUP(CCI$(1))            ' POR SI IMPRESORA APAGADA
        If InStr(EMPRELI$, "NOVOESPU") > 0 Then
            Call CARSETUP(CCI$(4))
            Call CARSETUP(CCI$(10) + CCI$(11))
          Else
            Call CARSETUP(CCI$(3))
            Call CARSETUP(CCI$(10) + CCI$(12))
        End If
        '
        VENDEX% = VENDE% + VVX%
        '
        CPAGX% = CPAG% + VVX%
        TPC$ = TIPCOM$
        NPR$ = TRIM$(NPROCLI$(NC%))        ' Numero de proveedor
        '
        RTO$ = "": RTO2$ = ""
        If InStr(EMPREAC$, "SABO") > 0 Then GoTo 2930
        '
        For UV& = 1 To ULTREG&("!CARDEFAC")
           Z$ = REGLEIDO$("!CARDEFAC", UV&)
           RTO1$ = TRIM$(Mid$(Z$, 131, 16))
           If RTO1$ = RTO2$ Then RTO1$ = ""
           If RTO1$ <> "" Then
               RTO2$ = RTO1$
               If RTO$ <> "" Then RTO$ = RTO$ + "-"
               RTO$ = RTO$ + RTO1$
           End If
        Next UV&
        If Len(RTO$) > 16 Then RTO$ = "VARIOS"
        '
2930    If RTO$ = "" Then
           REMI = NRO                         ' para SABO el numero es el mismo
           '
           If InStr(EMPREAC$, "SABO") > 0 Then
              If EXPOR% = 1 Then
                'REMI = NRO - 300
                REMI = NRO
                GoTo 2931
              End If
              '
              If NRO <= 28000 Then
                  REMI = NRO - 12500
                Else
                  'REMI = NRO - 12200 ' 12350
                  REMI = NRO - 12050 '12350
              End If
              If VA% = 2 Then
                'REMI = NRO + 15500 - 750
                REMI = NRO + 35900
              End If
           End If
           '
2931       RTO$ = TRIM$(Str$(REMI))
           While Len(RTO$) < 8: RTO$ = "0" + RTO$: Wend
           PREFI$ = ATRIFORM$(FORIMPRE$, "PREFIFAC")
           RTO$ = PREFI$ + "-" + RTO$
        End If
        If TRIM$(FORMRAC.Text10.TEXT) <> "" Then
            RTO$ = TRIM$(FORMRAC.Text10.TEXT)
        End If
        '
        NROCON& = NRO
        VALOR# = TOTAFAC#(1, 8)
        '
        If FORMAIMPRE$ = "DOS" Then
              'Call PRIFADOS
           Else
              GoSub 6000                         ' imprimir
        End If
        '
        NRXX& = NRO + 1
        While NRXX& < NROCON& + 0.1
           NRT = NRXX&
           NRXX& = NRXX& + 1
        Wend
        '
2990    TPXS$ = "Factura"
        If TIPO% = 3 Then TPXS$ = "Nota de Débito"
        If TIPO% = 6 Then TPXS$ = "Nota de Crédito"
        AI% = COMALTER%("Revise la " + TPXS$ + " Impresa\\Esta O.K.\Abortar")
        ''Call FRESHALL
        FORMRAC.Show
        If AI% <> 1 Then
           '
           For U& = URE2021& + 1 To URE2022&
              X$ = REGBLAN$("SVD202")
              Call GRAREG("SVD202", X$, U&)
           Next U&
           '
           For U& = UREFAC1& + 1 To UREFAC2&
              X$ = REGLEIDO$("FACTUR", U&)
              Call REPLA(X$, String$(50, 0), 15, 50)
              Call GRAREG("FACTUR", X$, U&)
           Next U&
           '
           For U& = UREMOV1& + 1 To UREMOV2&
              X$ = REGBLAN$("MOVISTO")
              Call GRAREG("MOVISTO", X$, U&)
           Next U&
           '
          GoTo 3500
          '
        End If
        '
        If MODESCU% = 1 Then
            If PORDES > 0 Then
                COEF = 1 - PORDES / 100
                For U& = 1 To ULTREG&("!CARDEFAC")
                       X$ = REGLEIDO$("!CARDEFAC", U&)
                       CAN = CVS(Left$(X$, 4))
                       PREUNI# = (CDbl(CVS(Mid$(X$, 429, 4)))) * COEF
                       Call REPLA(X$, MKS$(PREUNI#), 429, 4)
                       TOT# = CDbl(CAN * PREUNI#)
                       If CVI(Mid$(X$, 5, 2)) < 1 Then
                           TOT# = CVD(Mid$(X$, 433, 8)) * COEF
                       End If
                       '
                       IVAITEM# = TOT# * (ALICU / 100)
                       BRUITEM# = TOT# + IVAITEM#
                       '
                       Call REPLA(X$, MKD$(TOT#), 433, 8)
                       Call REPLA(X$, MKS$(ALICU), 441, 4)
                       Call REPLA(X$, MKD$(IVAITEM#), 445, 8)
                       Call REPLA(X$, MKD$(BRUITEM#), 453, 8)
                       Call GRAREG("!CARDEFAC", X$, U&)
                Next U&
                Call CIERRARCH("!CARDEFAC")
            End If
        End If
        '

3000    Screen.MousePointer = 11
        '
        VENDEX% = VENDE% + VVX%
        Call CIERRARCH("*.*")
        '
        ATRI$ = ""
        '
        Z$ = ""

3500  ' aqui limpia pantalla y vectores ****************
        '
        Call CIERRARCH("*.*")
        '
        FORMRAC.BOTNEXT.Enabled = True
        NC% = 0
        Call BLANARCH("!CARDEFAC")
        '
        FORMRAC.Text1.TEXT = "": FORMRAC.Text1.Refresh
        FORMRAC.Text2.TEXT = "": FORMRAC.Text2.Refresh
        FORMRAC.Text3.TEXT = "": FORMRAC.Text3.Refresh
        FORMRAC.Text4.TEXT = "": FORMRAC.Text4.Refresh
        FORMRAC.Text5.TEXT = "": FORMRAC.Text5.Refresh
        FORMRAC.Text6.TEXT = "": FORMRAC.Text6.Refresh
        FORMRAC.Text7.TEXT = "": FORMRAC.Text7.Refresh
        FORMRAC.Text8.TEXT = "": FORMRAC.Text8.Refresh
        FORMRAC.Text9.TEXT = "": FORMRAC.Text9.Refresh
        FORMRAC.Text10.TEXT = "": FORMRAC.Text10.Refresh
        FORMRAC.Text11.TEXT = "": FORMRAC.Text11.Refresh
        '
        For KK% = 0 To 2
           FORMRAC.TEXT21(KK%).TEXT = ""
           FORMRAC.Eco21(KK%) = ""
        Next KK%
        '
        For KK% = 0 To 3
           FORMRAC.Total(KK%).TEXT = ""
        Next KK%
        '
        FORMRAC.PORDESCU.TEXT = ""
        FORMRAC.PORIVA.TEXT = ""
        FFI% = HOY(HO$)
        FORMRAC.Text9.TEXT = HO$
        'Call FRESHALL
        '
        FORMRAC.Text1.SetFocus
        Screen.MousePointer = 1
        ''Call FRESHALL
        'FORMRAC.Hide
        FORMRAC.Show
        '
        Exit Sub
        '
'****************************************************************************
'
4000    If X$ = "" Then GoTo 3500
        '
        Call BLANARCH("!FACSEL")
        NX% = FILENBR%("FACTUR")
        Dim AX As String * 64
        FIN& = LOF(NX%) / 64
        For IL& = 1 To FIN&
           'Call TRACE(20, IL&, FIN&)
           Get #NX%, IL&, AX$
           If CVI(Mid$(AX$, 7, 2)) >= DES% Then
              If CVI(Mid$(AX$, 7, 2)) <= HAS% Then
                 NC% = CVI(Left$(AX$, 2))
                 CLI$ = AJUSTI$(RASOCLI$(NC%), 30)
                 X$ = Mid$(AX$, 3, 4) + CLI$ + MKS$(IL&)
                 Call REGAPP("!FACSEL", X$)
              End If
           End If
        Next IL&
        Call CIERRARCH("*.*")
        '
        'X$ = REGSEL$("FACSEL")
        If Len(X$) <= 4 Then GoTo 3500
        REFA& = CVS(Mid$(X$, 35, 4))
        '
        If ALTERNA%("-Confirma Anulación") <> 1 Then GoTo 3500
        If REFA& > 0 Then
            If REFA& <= ULTREG&("FACTUR") Then
               X$ = REGLEIDO$("FACTUR", REFA&)
               Call REPLA(X$, String$(50, 0), 15, 50)
               Call GRAREG("FACTUR", X$, REFA&)
               Call CIERRARCH("*.*")
               Call FINAL("DIAGNOS")
            End If
        End If
        '
        GoTo 3500
'
'****************************************************************************
'
6000    PORIAN = Printer.Orientation  ' retiene orientacion
        CODFOR$ = FORIMPRE$
        GoSub 6800                            ' LEER ATRIBUTOS DEL FORMULARIO
        '
        Printer.Orientation = 1
        If ORIENTA% = 2 Then
            Printer.Orientation = 2
        End If
        '
        Call COMUNI("Verifique Encuadre y Pulse\<ENTER> para Continuar ...")
        PRIDET0& = 1: HOJACONT% = 0: TOTITEMS = 0
6005    TRANSPOR% = 0
        ORDECOP% = 0
           For COP% = 1 To CACOPI%
              PRIDET& = PRIDET0&
              Call CARFORWIN
              GoSub 7100                      ' DATOS GENERALES DE FACTURA
              YAC = DETDESMM
              GoSub 7300: UREDET% = UREN%     ' CARGAR DETALLE DE FACTURACION
              GoSub 7500                      ' CARGAR LEYENDAS
              '
              'If COP% >= PRICOFAC% And COP% <= ULTCOFAC% Then
                  If HOJACONT% = 0 Then        ' SOLO EN PRIMERA HOJA
                     If CONTIHOJA% = 0 Then
                        GoSub 7400                '                CARGAR TOTALES
                     End If
                  End If
              'End If
              '
              If CONTIHOJA% > 0 Then
                  If TRANSPOR% = 0 Then
                      GoSub 7400
                  End If
              End If
              '
              Printer.EndDoc
              '
        Next COP%
        '
        PRIDET0& = PRIDET&
        '
        If TRANSPOR% > 0 Then                 ' TRANSPORTA A HOJA SIGUIENTE
            HOJACONT% = HOJACONT% + 1
            If CONTIHOJA% = 0 Then
                NROCON& = NROCON& + 1
            End If
            '
            Printer.EndDoc
            '
            GoTo 6005
        End If
        '
        If VEPIEDET$ <> "" Then               ' LEYENDAS COMPLEMENTARIAS
           For COP% = 1 To CANCOP%
              Call CARFORWIN
              GoSub 7100                      ' DATOS GENERALES DE FACTURA
              '
              YAC = DETDESMM
              For U& = 1 To ULTREG&("!FACEXPO2")
                  XXX$ = REGLEIDO$("!FACEXPO2", U&)
                  TXT$ = XXX$
                  Call APPFORM(TXT$, "ESPE-DET")
                  YAC = YAC + INTERLI
              Next U&
              '
              Printer.EndDoc    ' imprime hoja
              '
           Next COP%
           If CONTIHOJA% = 0 Then
               NROCON& = NROCON& + 1
           End If
        End If
        '
6007    AI5% = COMALTER%("Puede Repetir Impresión\\Continuar\Repite Impresion")
        ''Call FRESHALL
        FORMRAC.Show
        If AI5% = 2 Then
            NROCON& = NRO                   ' EMPIEZA DESDE EL PRIMERO
            GoTo 6000
        End If
        Printer.Orientation = PORIAN
        '
        Call CIERRARCH("FORMUFAC")
        '
        Return
        '
'****************************************************************************
        '
6800    PRICOFAC% = Val(ATRIFORM$(CODFOR$, "IMP-DES"))
        ULTCOFAC% = Val(ATRIFORM$(CODFOR$, "IMP-HAS"))
        MONPRI$ = ATRIFORM$(CODFOR$, "MON-PRI")
        MONSEC$ = ATRIFORM$(CODFOR$, "MON-SEC")
        PREFI$ = ATRIFORM$(CODFOR$, "PREFIFAC")
        LOIMPLET% = Val(ATRIFORM$(CODFOR$, "LOIMPLET"))
        If LOIMPLET% < 16 Then LOIMPLET% = 32
        '
        MARIZW = Val(ATRIFORM$(CODFOR$, "MARIZQMM"))
        MARSUW = Val(ATRIFORM$(CODFOR$, "MARSUPMM"))
        DETDESMM = Val(ATRIFORM$(CODFOR$, "DETDESMM"))
        DETHASMM = Val(ATRIFORM$(CODFOR$, "DETHASMM"))
        INTERLI = Val(ATRIFORM$(CODFOR$, "INTERLIN"))
        If INTERLI < 1 Then
            Call COMUNI("Error de Configuración de Interlinea\en Formulario '" + CODFOR$ + ".FRM'.")
            INTERLI = 25.4 / 6
        End If
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
        Return
        '
'****************************************************************************
        '
7100    FEX = FFI%
        TCMP$ = UCase$(TIPCOM$): If COP% > ULTCOFAC% Then TCMP$ = "REMITO"
        Call PRINTFORWIN(TCMP$, "TIPO-COM", 0, 0)
        
        Call PRINTFORWIN(FECHATEX$(FEX), "FECH-EMI", 0, 0)
        Call PRINTFORWIN(FORMRAC.Text2.TEXT, "RASO-CLI", 0, 0)
        Call PRINTFORWIN(FORMRAC.Text8.TEXT, "CUIT-CLI", 0, 0)
        Call PRINTFORWIN(FORMRAC.Text3.TEXT, "DOMI-CLI", 0, 0)
        Call PRINTFORWIN(NIBRCLI$(NC%), "IBRU-CLI", 0, 0)
        Call PRINTFORWIN(FORMRAC.Text4.TEXT, "CPOS-CLI", 0, 0)
        Call PRINTFORWIN(FORMRAC.Text5.TEXT, "LOCA-CLI", 0, 0)
        Call PRINTFORWIN(NPROCLI$(NC%), "PROV-CLI", 0, 0)
        Call PRINTFORWIN(Str$(NC%), "CUEN-CLI", 0, 0)
        POSIVA$ = ECOARCH$("CAPOSIV", Chr$(PIVA%))
        Call PRINTFORWIN(POSIVA$, "PIVA-CLI", 0, 0)
        TCMP$ = UCase$(CATIVAS$): If COP% > ULTCOFAC% Then TCMP$ = "X"
        Call PRINTFORWIN(TCMP$, "CIVA-CLI", 0, 0)
        Call PRINTFORWIN(DECONPA$(CPAG%), "COND-PAG", 0, 0)
        '
        Call PRINTFORWIN(RTO$, "NRO-REMI", 0, 0)
        '
        Call PRINTFORWIN(Mid$(TRXP$, 1, 32), "TRANSPOR", 0, 0)
        Call PRINTFORWIN(Mid$(TRXP$, 33, 32), "DOMI-TRA", 0, 0)
        Call PRINTFORWIN(Mid$(TRXP$, 65, 16), "CUIT-TRA", 0, 0)
        Call PRINTFORWIN(Mid$(TRXP$, 81, 16), "PATE-TRA", 0, 0)
        Call PRINTFORWIN(Mid$(TRXP$, 97, 32), "COND-TRA", 0, 0)
        '
        NCN$ = TRIM$(Str$(NROCON&))
        While Len(NCN$) < 8: NCN$ = "0" + NCN$: Wend
        PREFI$ = ATRIFORM$(CODFOR$, "PREFIFAC")
        NCN$ = PREFI$ + "-" + NCN$
        Call PRINTFORWIN(NCN$, "NUME-COM", 0, 0)
        COPXX% = COP%: If COPXX% > ULTCOFAC% Then COPXX% = COP% - ULTCOFAC%
        NCN$ = ORCOP$(COPXX%)
        Call PRINTFORWIN(NCN$, "ORD-COPI", 0, 0)
         '
7190    Return
        '
'****************************************************************************
        '
7300    If PRIDET& = 1 Then
           If VENCADET$ <> "" Then
              For U& = 1 To ULTREG&("!FACEXPO1")
                  XXX$ = REGLEIDO$("!FACEXPO1", U&)
                  TXT$ = XXX$
                  Call APPFORM(TXT$, "ENCA-DET")
                  RG0% = RGII% + 1
              Next U&
              YAC = IAC + INTERLI
           End If
        End If
        '
        For U& = PRIDET& To ULTREG&("!CARDEFAC")
            '
            If YAC > DETHASMM - 4 * INTERLI Then                      ' SALTO DE HOJA
               PRIDET& = U&
               TRANSPOR% = 1
               GoTo 7390
            End If
            '
            XXX$ = REGLEIDO$("!CARDEFAC", U&)
            '
            TOTITEMS = TOTITEMS + CVS(Left$(XXX$, 4))
            TXT$ = Str$(CVS(Left$(XXX$, 4)))       ' cantidad
                Call APPFORM(TXT$, "CANTIDAD")
                '
            TXT$ = CODEXT$(CVI(Mid$(XXX$, 5, 2)))     ' codigo
                Call APPFORM(TXT$, "COD-MAT")
                '
            TXT$ = Mid$(XXX$, 7, 64)                  ' descripcion
                Call APPFORM(TXT$, "DES-MAT")
                '
            TXT$ = Mid$(XXX$, 71, 4)                  ' unidad de medida
                Call APPFORM(TXT$, "UNIMED")
                '
            TXT$ = Mid$(XXX$, 75, 64)                 ' referencia 1
                Call APPFORM(TXT$, "REF-MAT1")
                '
            TXT$ = Mid$(XXX$, 139, 64)                ' referencia 2
                Call APPFORM(TXT$, "REF-MAT2")
                '
            TXT$ = Mid$(XXX$, 203, 64)                ' referencia 3
                Call APPFORM(TXT$, "REF-MAT3")
                '
            TXT$ = Mid$(XXX$, 315, 24)                ' numero de plano
                Call APPFORM(TXT$, "N-PLANO")
                '
            TXT$ = Mid$(XXX$, 339, 16)                ' numero original
                Call APPFORM(TXT$, "N-ORIGIN")
                '
            TXT$ = Mid$(XXX$, 355, 16)                ' numero de o.compra
                Call APPFORM(TXT$, "N-OCOMPR")
                '
            TXT$ = Mid$(XXX$, 387, 16)                ' numero de remito
                Call APPFORM(TXT$, "N-REMITO")
                '
            TXT$ = Mid$(XXX$, 403, 16)                ' numero de doc-recepcion
                Call APPFORM(TXT$, "N-BOLREC")
                '
            FEX = CVI(Mid$(XXX$, 419, 2))             ' fecha de recepcion
                TXT$ = FECHATEX$(FEX)
                Call APPFORM(TXT$, "FECH-REC")
                '
          If COP% >= PRICOFAC% And COP% <= ULTCOFAC% Then
              TXT$ = Str$(FACIVA * CVS(Mid$(XXX$, 421, 4))) ' precio de lista
                Call APPFORM(TXT$, "PRE-LIST")
                '
              TXT$ = Str$(FACIVA * CVS(Mid$(XXX$, 429, 4))) ' precio unitario
                Call APPFORM(TXT$, "PRE-NETO")
                '
              TXT$ = Str$(CVD(Mid$(XXX$, 433, 8)))   ' total neto
                Call APPFORM(TXT$, "NET-ITEM")
                '
              TXT$ = Str$(CVS(Mid$(XXX$, 441, 8)))   ' alicuota iva
                Call APPFORM(TXT$, "ALICUIVA")
                '
              TXT$ = Str$(CVD(Mid$(XXX$, 445, 8)))   ' importe iva del item
                Call APPFORM(TXT$, "IVA-ITEM")
                '
              TXT$ = Str$(CVD(Mid$(XXX$, 453, 8)))   ' total bruto
                Call APPFORM(TXT$, "BRU-ITEM")
                '
            Else
               '
               TXT$ = "********************"
               Call APPFORM(TXT$, "PRE-LIST")
               Call APPFORM(TXT$, "PRE-NETO")
               Call APPFORM(TXT$, "NET-ITEM")
               Call APPFORM(TXT$, "ALICUIVA")
               Call APPFORM(TXT$, "IVA-ITEM")
               Call APPFORM(TXT$, "BRU-ITEM")
               '
          End If
          '
          YAC = YAC + INTERLI
          '
        Next U&
        '
        YAC = YAC + INTERLI
        Call APPFORM(String$(32, "="), "LIN-CIE")
        YAC = YAC + INTERLI
        Call APPFORM(Str$(TOTITEMS), "TOTA-ITS")
        Call APPFORM("Total de Items", "LETOT-IT")
        '
7390    Return
        '
'****************************************************************************
        '
7400    If COP% > ULTCOFAC% Then GoTo 7410
        '
        Call PRINTFORWIN(Str$(FACIVA * TOTAFAC#(1, 1)), "IMP-NETO", 0, 0)
        DESCUE$ = TRIM$(FORMRAC.PORDESCU.TEXT)
        Call PRINTFORWIN(DESCUE$, "POR-DESC", 0, 0)
        Call PRINTFORWIN(Str$(FACIVA * TOTAFAC#(1, 4)), "IMP-DESC", 0, 0)
        Call PRINTFORWIN(Str$(FACIVA * TOTAFAC#(1, 5)), "IMPNEDES", 0, 0)
        Call PRINTFORWIN(Str$(TOTAFAC#(1, 6)), "IMP-IVA", 0, 0)
        Call PRINTFORWIN(Str$(TOTAFAC#(1, 7)), "IMP-RNI", 0, 0)
        Call PRINTFORWIN(Str$(TOTAFAC#(1, 8)), "IMP-TOTA", 0, 0)
        Call PRINTFORWIN(Str$(TOTAFAC#(1, 9)), "IMP-VARS", 0, 0)
        Call PRINTFORWIN(Str$(TOTAFAC#(1, 10)), "PERC-IVA", 0, 0)
        '
        TTTT# = TOTAFAC#(1, 8)
        TOTFAC$ = NUMTEX$(TTTT#)
        Call FRATEX(TOTFAC$, NRX1$, NRX2$, NRX3$, LOIMPLET%)
        Call PRINTFORWIN(NRX1$, "IMP-LET1", 0, 0)
        Call PRINTFORWIN(NRX2$, "IMP-LET2", 0, 0)
        Call PRINTFORWIN(NRX3$, "IMP-LET3", 0, 0)
        '
        Return
        '
7410    Call PRINTFORWIN("******************", "IMP-NETO", 0, 0)
        Call PRINTFORWIN("****************", "POR-DESC", 0, 0)
        Call PRINTFORWIN("****************", "IMP-DESC", 0, 0)
        Call PRINTFORWIN("****************", "IMPNEDES", 0, 0)
        Call PRINTFORWIN("****************", "IMP-IVA", 0, 0)
        Call PRINTFORWIN("****************", "IMP-RNI", 0, 0)
        Call PRINTFORWIN("****************", "IMP-TOTA", 0, 0)
        Call PRINTFORWIN("****************", "IMP-VARS", 0, 0)
        Call PRINTFORWIN("****************", "PERC-IVA", 0, 0)
        '
        Call PRINTFORWIN(String$(80, "*"), "IMP-LET1", 0, 0)
        Call PRINTFORWIN(String$(80, "*"), "IMP-LET2", 0, 0)
        Call PRINTFORWIN(String$(80, "*"), "IMP-LET3", 0, 0)
        '
        Return
        '
'****************************************************************************
        '
7500    Call CARLEYEN("")
        '
        Return
        '
'****************************************************************************
        '
'****************************************************************************
'
8200    ' calcular IVA de cada item
        '
        For U& = 1 To ULTREG&("!CARDEFAC")
            X$ = REGLEIDO$("!CARDEFAC", U&)
            TOT# = CVD(Mid$(X$, 433, 8))
            IVAITEM# = TOT# * (ALICU / 100)
            BRUITEM# = TOT# + IVAITEM#
            '
            Call REPLA(X$, MKS$(ALICU), 441, 4)
            Call REPLA(X$, MKD$(IVAITEM#), 445, 8)
            Call REPLA(X$, MKD$(BRUITEM#), 453, 8)
            '
            Call GRAREG("!CARDEFAC", X$, U&)
        Next U&
        '
        Return
'
'****************************************************************************
'
8700 TOT1# = 0: TOT2# = 0: TOT3# = 0: TOT4# = 0
     RGXI% = RGII%: CLII% = 66: LCII% = 13: TV% = 3: LD% = 2: JTII% = 2
8702 For JJ& = 1 To ULTREG&("!CARDEFAC")
        TOT# = CVD(Mid$(REGLEIDO$("!CARDEFAC", JJ&), 433, 8) + String$(8, 0))
8715 TOT1# = (Int(100 * TOT1# + 0.5) / 100) + (Int(100 * TOT# + 0.5) / 100)
8720 Next JJ&
8722 '
8727 'por ahora los descuentos se aplican directamente sobre el pr.unit
     TOT2# = 0
     If PORDES > 0 Then
        If MODESCU% = 1 Then
8730        TOT2# = (Int(TOT1# * PORDES + 0.5)) / 100
8732        TOTAFAC#(1, 9) = (Int(TOTAFAC#(1, 9) * (100 - PORDES) + 0.5)) / 100
        End If
     End If
8735 '
8743 TOT3# = (Int((TOT1# - TOT2#) * ALICU + 0.5)) / 100:
     TOT4# = TOT1# - TOT2# + TOT3#
8748 '
8750 If FACIVA <> 1 Then GoTo 8760
8751 RGII% = 19: VD# = TOT1#
     FORMRAC.Total(0).TEXT = CSTRING$(MKD$(VD#), 4, 18, 2, 2)
        TOTAFAC#(1, 1) = TOT1#
        TOTAFAC#(1, 3) = TOT1#
8752 RGII% = 20: VD# = TOT2#
     FORMRAC.Total(1).TEXT = CSTRING$(MKD$(VD#), 4, 18, 2, 2)
     'Call PRESENTA(20, 0, CL%, LC%, LK%, TV%, LD%, JT%, VDI%, TOT2#, VDS$, CT%)
        TOTAFAC#(1, 4) = TOT2#
        TOTAFAC#(3, 4) = 0
                If DESCUPRE% = 0 Then
                        TOTAFAC#(3, 4) = CDbl(PORDES)
                End If
        TOTAFAC#(1, 5) = ((Int(100 * TOTAFAC#(1, 3) + 0.5)) - (Int(100 * TOTAFAC#(1, 4) + 0.5))) / 100
8753 RGII% = 21: VD# = TOT3#
     FORMRAC.Total(2).TEXT = CSTRING$(MKD$(VD#), 4, 18, 2, 2)
     'Call PRESENTA(21, 0, CL%, LC%, LK%, TV%, LD%, JT%, VDI%, TOT3# + TOTAFAC#(1, 10), VDS$, CT%)
        TOTAFAC#(1, 6) = (Int(TOTAFAC#(1, 5) * ALI1 + 0.5)) / 100
        TOTAFAC#(3, 6) = CDbl(ALI1)
        TOTAFAC#(1, 7) = (Int(TOTAFAC#(1, 5) * ALI2 + 0.5)) / 100
        TOTAFAC#(3, 7) = CDbl(ALI2)
        If ALICU = ALI1 Then TOTAFAC#(1, 7) = 0: TOTAFAC#(3, 7) = 0
        If ALICU < 0.01 Then TOTAFAC#(1, 6) = 0: TOTAFAC#(3, 6) = 0: TOTAFAC#(1, 7) = 0: TOTAFAC#(3, 7) = 0
        TOTAFAC#(1, 8) = TOTAFAC#(1, 5) + TOTAFAC#(1, 6) + TOTAFAC#(1, 7) + TOTAFAC#(1, 9) + TOTAFAC#(1, 10)
8754 RGII% = 22: VD# = TOT4#
     FORMRAC.Total(3).TEXT = CSTRING$(MKD$(VD#), 4, 18, 2, 2)
     'Call PRESENTA(22, 0, CL%, LC%, LK%, TV%, LD%, JT%, VDI%, TOT4# + TOTAFAC#(1, 9) + TOTAFAC#(1, 10), VDS$, CT%)
8756 GoTo 8768
8757 '
8760 RGII% = 19: VD# = TOT4#
     'Call PRESENTA(19, 0, CL%, LC%, LK%, TV%, LD%, JT%, VDI%, TOT4#, VDS$, CT%)
        TOTAFAC#(1, 1) = TOT1#
        TOTAFAC#(1, 3) = TOT1#
8762 RGII% = 20: VD# = TOT2#
     'Call PRESENTA(20, 0, CL%, LC%, LK%, TV%, LD%, JT%, VDI%, TOT2#, VDS$, CT%)
        TOTAFAC#(1, 4) = TOT2#
        TOTAFAC#(3, 4) = 0
                If DESCUPRE% = 0 Then
                        TOTAFAC#(3, 4) = CDbl(PORDES)
                End If
        TOTAFAC#(1, 5) = ((Int(100 * TOTAFAC#(1, 3) + 0.5)) - (Int(100 * TOTAFAC#(1, 4) + 0.5))) / 100
8763 RGII% = 21: VD# = TOT4#
     'Call PRESENTA(21, 0, CL%, LC%, LK%, TV%, LD%, JT%, VDI%, TOT4#, VDS$, CT%)
        TOTAFAC#(1, 8) = TOT4#
8764 RGII% = 22: VD# = TOT3#
     'Call PRESENTA(22, 0, CL%, LC%, LK%, TV%, LD%, JT%, VDI%, TOT3#, VDS$, CT%)
        TOTAFAC#(1, 6) = TOT3#
        TOTAFAC#(1, 7) = 0

8768 For I% = 0 To 3
        FORMRAC.Total(I%).Refresh
     Next I%
     RGII% = RGXI%

8800 For I% = 1 To 10
        TOTAFAC#(2, I%) = TOTAFAC#(1, I%) / TICA
     Next I%

     Return

End Sub
'****************************************************************************


