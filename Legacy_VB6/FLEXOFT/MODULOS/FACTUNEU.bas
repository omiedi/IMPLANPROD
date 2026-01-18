Attribute VB_Name = "FACTUNEU"
Dim A$(8)
Dim DATCLI$(32)
Dim POI%(10)
Dim T$(2)
Dim RG&(16)
Dim RF$(4)
'
Dim TPX%(8), TIPCOMS$(8), VTIPS$(8)
Dim TIFORM%
Public CONTROFA$
Dim FORVENTA$
Dim FORIMPRE$
Dim VENCADET$
Dim VEPIEDET$
Public FORMAIMPRE$
Dim MAXITEMS%
Dim UREN%, RG0%
Dim RTO$, TRXP$
'
Dim NUPEDSEL&(512)
'
Public ARCHIFAC$, ARCHIDEU$, ARCHICLI$, ARCHIMAT$, ARCHISTO$
'
Public TDFAC$
Public TIDOC%
Public CAFOR%
Public NRO, NROCON&
Dim TIPCOM$
Public TIPO%, VA%
Public PIVA%
Dim TICA
Dim FFI%, FVI%
Public ALICU, ALI1, ALI2
Dim PORDES, MODESCU%, VERISTO%
Dim FACIVA
Dim CATIVAS$
'
Dim SELORDESPA%
'

Sub Main()
    '
    Static CTX%
    '
    Call VERJOBID
    Call BLOQARCH("FACTURA")
    '
    FORFACNEU.Show
    FORFACNEU.Refresh
    '
    ' recuperacion de datos novoespuma
    LU$ = LUDAT$
    EPAC$ = EMPREAC$
    If InStr(EPAC$, "NOVOESPU") > 0 Then
       HOII% = HOY(HOS$)
       If HOII% = FECHANUM("21/05/01") Then
          If EXISTE%(LU$ + "FATRANS.DAT") = 1 Then
             Call COPYSTRU("FACTUR", "FATRANS")
             For U& = 1 To ULTREG&("FATRANS")
                XXX$ = REGLEIDO$("FATRANS", U&)
                Call REGAPP("FACTUR", XXX$)
             Next U&
             Call CIERRARCH("FACTUR")
             Call CIERRARCH("FATRANS")
             Call FILEKILL(LU$ + "FATRANS.DAT")
             Call COMUNI("Recuperacion de Datos Completa")
             Call FINAL("")
          End If
       End If
    End If
    '
'For U& = 1 To ULTREG&("FACTUR")
'XX$ = REGLEIDO$("FACTUR", U&)
'MsgBox Asc(Mid$(XX$, 47, 1)) & "   " & Asc(Mid$(XX$, 48, 1))
'Next U&
    FORMAIMPRE$ = "WIN"
    EPAC$ = EMPREAC$
    If InStr(UCase$(EPAC$), "SABO") > 0 Then FORMAIMPRE$ = "DOS"
    '
    UTI$ = LUBAS$
    If Mid$(UTI$, 2, 1) <> ":" Then
        UTI$ = Left$(CurDir, 2) + UTI$
    End If
    If EXISTE%(UTI$ + "PRINTFRM.EXE") < 1 Then
        Call COMUNI("Imposible Procesar Facturación.\  \Falta Utilitario 'PRINTFRM.EXE'\en Directorio '" + REPLACAR$(UTI$, "\", "/") + "'.")
        Call FINAL("")
    End If
    '
    ORCOP$(1) = "ORIGINAL"
    ORCOP$(2) = "DUPLICADO"
    ORCOP$(3) = "TRIPLICADO"
    ORCOP$(4) = "CUADRUPLICADO"
    ORCOP$(5) = "QUINTUPLICADO"
    ORCOP$(6) = "SEXTUPLICADO"
    
100 Call CIERRARCH("*.*")
    If CTX% = 0 Then
        IQ& = 0
        IQ& = IQ& + 1: Call GRAREG("CAPOSIV", Chr$(0) + AJUSTI$("Consumidor Final            ", 31) + AJUSTI$("Cons.Fin", 9) + String$(23, 0), IQ&)
        IQ& = IQ& + 1: Call GRAREG("CAPOSIV", Chr$(1) + AJUSTI$("IVA Responsable Inscripto   ", 31) + AJUSTI$("Resp.Insc", 9) + String$(23, 0), IQ&)
        IQ& = IQ& + 1: Call GRAREG("CAPOSIV", Chr$(2) + AJUSTI$("IVA Responsable no Inscripto", 31) + AJUSTI$("R.No Insc", 9) + String$(23, 0), IQ&)
        IQ& = IQ& + 1: Call GRAREG("CAPOSIV", Chr$(3) + AJUSTI$("IVA No Responsable          ", 31) + AJUSTI$("No Resp.", 9) + String$(23, 0), IQ&)
        IQ& = IQ& + 1: Call GRAREG("CAPOSIV", Chr$(4) + AJUSTI$("IVA Exento                  ", 31) + AJUSTI$("Exento  ", 9) + String$(23, 0), IQ&)
        IQ& = IQ& + 1: Call GRAREG("CAPOSIV", Chr$(5) + AJUSTI$("IVA Exportador Exento       ", 31) + AJUSTI$("Export. ", 9) + String$(23, 0), IQ&)
        IQ& = IQ& + 1: Call GRAREG("CAPOSIV", Chr$(6) + AJUSTI$("Contribuyente Monotributo   ", 31) + AJUSTI$("Monotrib", 9) + String$(23, 0), IQ&)
        Call SETULTREG&("CAPOSIV", IQ&)
        Call CIERRARCH("CAPOSIV")
        CTX% = 1
    End If
    '
102 Call CAREMPRE(LU$, A$())
    Call FORMPRIN(LU$, FLIM%, ALI1, ALI2, TC1, TLET$, INTER%, TFOR%, ANFO%, LARFO%, DIFER, CAFOR%, MARIZ%, MARINT%, REDON%, MPRI$, MCEN$, FACTO, CANCON%, COMA%, PARDEC%)
    TICA = TC1
    '
109 VTIP2$ = "Remito(s)"
    ARCHIFAC$ = "FACTUR"
    ARCHIDEU$ = "SVD202"         ' replantear estas variantes segun facstan
    ARCHICLI$ = "DEUDOR"
    ARCHIMAT$ = "MATVEN"
    ARCHISTO$ = "SALSTOCK"
    '
' ***************************************************************************

200 Screen.MousePointer = 1
    '
    Call VAPUBLI
    '
    Call INDIPEDI
    '
' ***************************************************************************

400 FFI% = HOY(HO$)
    FORFACNEU.Text9.TEXT = HO$
    '
    VARIFA$ = TRIM$(TIPCOM$) + " - " + TRIM$(Mid$(CONTROFA$, 33, 32))
    '
' ***************************************************************************
    '
    ' correccion falla sabo
    If InStr(EMPRELI$, "SABO") < 1 Then GoTo 500
    '
    Screen.MousePointer = 11
    UI& = ULTREG&("MOVISTO")
    For U& = ULTREG&("MOVISTO") To 64000 Step -1
       X$ = REGLEIDO$("MOVISTO", U&)
       If TRIM$(Mid$(X$, 23, 10)) = "RT.0" Then
          If Mid$(X$, 33, 3) = "FC." Then
             NUFA$ = Mid$(X$, 36, 7)
             Call REPLA(X$, NUFA$, 26, 7)
             Call GRAREG("MOVISTO", X$, U&)
             UI& = U&
          End If
       End If
       If U& < UI& - 512 Then GoTo 500
    Next U&
    '
500 Call CIERRARCH("MOVISTO")
    Screen.MousePointer = 1
    FORFACNEU.Text1.SetFocus
    '
End Sub
'

Sub VAPUBLI()
    '
    EPAC$ = TRIM$(UCase$(EMPREAC$))
    If CONTROFA$ = "" Then
       Call MENSERR(24, "No se ha Establecido\el Controlador de Facturación.\  \Consulte a su Soporte de Sistemas.")
       Call FINAL("")
    End If
    '
    TIDOC% = Asc(Left$(CONTROFA$, 1))
    VTIP2$ = Space$(9)
    '
    On TIDOC% + 1 GoTo 200, 210, 220, 230
    Call FINAL("")
    '
200 TIPO% = 1
    TIPCOM$ = "Contado"
    MONEDA% = 1: GoTo 300
    '
210 TIPO% = 2: VTIP2$ = "Remito(s)"
    TIPCOM$ = "Factura"
    MONEDA% = 1: GoTo 300
    '
220 TIPO% = 3
    TIPCOM$ = "N.Debito": If InStr(EPAC$, "NEUMANN") > 0 Then TIPCOM$ = "Nota de Debito"
    MONEDA% = 1: GoTo 300
    '
230 TIPO% = 6
    TIPCOM$ = "Credito": If InStr(EPAC$, "NEUMANN") > 0 Then TIPCOM$ = "Nota de Credito"
    MONEDA% = 1: GoTo 300
    '
300 FLIM% = FECHANUM("01/01/99")
    If ULTREG&("FORMFAC") > 0 Then
       FFS$ = REGLEIDO$("FORMFAC", 2&)
       FLIM% = CVI(Mid$(FFS$, 19, 2))
    End If
    '
    FORFACNEU.Text11.TEXT = ""
    If NC% > 0 Then
        Call GRAFAC98(LU$, ARCHIFAC$, 1, RD%, FLIM%, TIPO%, VA%, CAFOR%, NRO, NC%, CLI$, IVAS$, FFI%, FVI%, LPRE%, VENDE%, TOTAFAC#())
        NROPRO = NRO
        FORFACNEU.Text11.TEXT = NROPRO
    End If
    '
End Sub
'

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

Sub SELTIFAC()
    '
    TDFAC$ = REGSEL$("SELTIFAC")
    If Len(TDFAC$) <= 4 Then
        Call CIERRARCH("*.*")
        Call FINAL("")
    End If
    '
    RGTIFA& = CVS(Mid$(TDFAC$, 129, 4))
    CONTROFA$ = REGLEIDO$("TIDOCFAC", RGTIFA&)
    Call CIERRARCH("TIDOCFAC")
    '
    EPAC$ = TRIM$(EMPREAC$)
    If EPAC$ <> "" Then
          ENCABE$ = "  " + EPAC$ + "  -  "
        Else
          ENCABE$ = "  "
    End If
    ENCABE$ = ENCABE$ + TRIM$(Left$(TDFAC$, 56))
    FORFACNEU.Label10.Caption = ENCABE$

End Sub
'

Sub CARDACLI()
    '
    Screen.MousePointer = 11
    NC% = Val(FORFACNEU.Text1.TEXT)
    If NC% > 0 Then
        FORFACNEU.Text2.TEXT = RASOCLI$(NC%)
        FORFACNEU.Text3.TEXT = DOMICLI$(NC%)
        FORFACNEU.Text4.TEXT = CPOSCLI$(NC%)
        FORFACNEU.Text5.TEXT = LOCACLI$(NC%)
        FORFACNEU.Text6.TEXT = PIVACLI%(NC%)
        FORFACNEU.Text7.TEXT = POSIVAC$(NC%)
        FORFACNEU.Text8.TEXT = CUITCLI$(NC%)
        '
        FORFACNEU.TEXT21(0).TEXT = Str$(LIPRCLI%(NC%))
        FORFACNEU.TEXT21(1).TEXT = Str$(CPAGCLI%(NC%))
        FORFACNEU.TEXT21(2).TEXT = Str$(VENDCLI%(NC%))
        '
        If Val(FORFACNEU.TEXT21(2).TEXT) < 1 Then
           FORFACNEU.Eco21(2).TEXT = "GERENCIA"
        End If
        '
        PIVA% = Val(FORFACNEU.Text6.TEXT)
        If PIVA% < 0 Or PIVA% > 6 Then
            FORFACNEU.Text6.SetFocus
            Screen.MousePointer = 1
            Exit Sub
        End If
        '
        FORFACNEU.PORDESCU.TEXT = DESCUCLI$(NC%)
        '
        GoSub 81                        ' determina posicion IVA y CATIVAS$
        'FORFACNEU.PORIVA(0).TEXT = ALICU
        '
1050    LU$ = LUDAT$
        CLI$ = RASOCLI$(NC%)
        FLIM% = FECHANUM("01/01/99")
        If ULTREG&("FORMFAC") > 0 Then
            FFS$ = REGLEIDO$("FORMFAC", 2&)
            FLIM% = CVI(Mid$(FFS$, 19, 2))
        End If
        '
        Call GRAFAC98(LU$, ARCHIFAC$, 1, RD%, FLIM%, TIPO%, VA%, CAFOR%, NRO, NC%, CLI$, IVAS$, FFI%, FVI%, LPRE%, VENDE%, TOTAFAC#())
        '
        NROPRO = NRO
        FORFACNEU.Text11.TEXT = NROPRO
        '
        Call NOTACLI(NC%, 1)
        '
    End If
    Screen.MousePointer = 1
    Exit Sub
    '
    '
81  'FACIVA = 1
    CATIVAS$ = "A"
    VA% = 1
    'ALICU = ALI1
    RNI% = 0
    '
    If PIVA% = 0 Or PIVA% = 3 Or PIVA% = 4 Or PIVA% = 6 Then
          'FACIVA = 1 + ALICU / 100
          CATIVAS$ = "B"
          VA% = 2
          '
        ElseIf PIVA% = 2 Then
          'ALICU = ALI1 + ALI2
          RNI% = 1
          '
        ElseIf PIVA% = 5 Then
          'ALICU = 0
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
'

Sub SIGUEFAC1()
    '
    LU$ = LUDAT$
    NCC$ = TRIM$(FORFACNEU.Text1.TEXT)
    If Val(NCC$) < 1 Then
        FORFACNEU.Text1.SetFocus
        Exit Sub
    End If
    '
    HOI1 = HOY(HO$)
    FFI1 = FECHANUM(TRIM$(FORFACNEU.Text9.TEXT))
    If FFI1 < HOI1 - 30 Or DIENTRE%(FFI1, HOI1) > 5 Then
        Call MENSERR(24, "Fecha no Válida")
        FORFACNEU.Text9.TEXT = FECHATEX$(HOI1)
        Exit Sub
    End If
    If DIENTRE%(HOI1, FFI1) > 1 Then
        Call MENSERR(24, "Fecha no Válida")
        FORFACNEU.Text9.TEXT = FECHATEX$(HOI1)
        Exit Sub
    End If
    FFI% = FFI1
    '
    NC% = Val(NCC$)
502 PIVA% = Val(FORFACNEU.Text6.TEXT)
    If ECOARCH$("CAPOSIV", Chr$(PIVA%)) = "" Then
        FORFACNEU.Text1.TEXT = "1"
        FORFACNEU.Text6.SetFocus
        Exit Sub
    End If
    '
    If PIVA% <> 0 Then
       If PIVA% <> 5 Then
          CUITT$ = TRIM$(FORFACNEU.Text8.TEXT)
          If CUITT$ = "" Or VALICUIT%(CUITT$) <> 1 Then
             Call MENSERR(24, "Imposible Facturar a esta Cuenta.\  \Número de CUIT no es Válido.")
             FORFACNEU.Text1.SetFocus
             Exit Sub
          End If
       End If
    End If
    '
    If TRIM$(DOMICLI$(NC%)) = "" Then
       Call MENSERR(24, "Imposible Facturar a esta Cuenta.\  \Falta Domicilio Legal.")
       FORFACNEU.Text1.SetFocus
       Exit Sub
    End If
    '
    LPRE% = Val(FORFACNEU.TEXT21(0).TEXT)
    If ECOARCH("LISTAS", MKI$(LPRE%)) = "" Then
        FORFACNEU.TEXT21(0).SetFocus
        Exit Sub
    End If
    '
    CPAG% = Val(FORFACNEU.TEXT21(1).TEXT)
    If ECOARCH("CONPAG", MKI$(CPAG%)) = "" Then
        FORFACNEU.TEXT21(1).SetFocus
        Exit Sub
    End If
    '
    VENDE% = Val(FORFACNEU.TEXT21(2).TEXT)
    If VENDE% > 0 Then
        If ECOARCH("VENDEDOR", Chr$(VENDE%)) = "" Then
            FORFACNEU.TEXT21(2).SetFocus
            Exit Sub
        End If
    End If
    '
    If LIMICRE(NC%) < 10 Then
      Call MENSERR(24, "Imposible Facturar a Esta Cuenta.\Falta Asignar Límite de Crédito.")
      FORFACNEU.Text1.SetFocus
      Exit Sub
    End If
    '
    If LIMICRE(NC%) < SALDCLI#(NC%) + CHEPEACRE#(NC%) Then
      Call MENSERR(24, "Imposible Facturar a Esta Cuenta.\Ha Excedido Límite de Crédito.")
      FORFACNEU.Text1.SetFocus
      Exit Sub
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
            ElseIf PIVA% = 5 Then
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
    FORFACNEU.BOTNEXT.Enabled = False
    EXPOR% = 0: If PIVA% = 5 And TIPO% = 2 Then EXPOR% = 1

'*********************************************************** fecha de factura

1200  If TIDOC% <> 1 Then
          CAPED% = 0
          GoTo 1205
      End If

'************************************** determina fecha de pedidos pendientes

      Screen.MousePointer = 11
      FINF = 32767: FSUP = 0: CAPED% = 0
      '
      FORSELPED.LIPEDPE.Clear
      '
      SELORDESPA% = 0
      NUOCOMPRA$ = "": NUOCOPARA$ = ""
      NROPEDIDO$ = "": NROPEDIPARA$ = ""
      FIN& = ULTREG&("ORDESPA")
      If FIN& < 1 Then GoTo 12021  ' selecciona por pedidos
      '
      ARCHIPRE$ = TRIM$(Str$(LPRE%))
      While Len(ARCHIPRE$) < 3: ARCHIPRE$ = "0" + ARCHIPRE$: Wend
      ARCHIPRE$ = "LIPRE" + ARCHIPRE$
      If LPRE% < 1 Then ARCHIPRE$ = "LIPRE001"
      CONSICLI% = 0
      RREMX$ = RECFILT$("CONSICLI", 1, MKI$(NC%))
AAA = Len(RREMX$)
      REGODES& = 0
      Call SETULTREG("!CARDEFAC", 0)
      If Len(RREMX$) >= 4 Then
         For UU& = 1 To Len(RREMX$) Step 4
           U& = CVS(Mid$(RREMX$, UU&, 4))
           X$ = REGLEIDO$("CONSICLI", U&)
           If CVI(Left$(X$, 2)) = NC% Then
              CI% = CVI(Mid$(X$, 35, 2))
              CAN = CVD(Mid$(X$, 37, 8)) - CVD(Mid$(X$, 45, 8))
              For JJ& = 1 To ULTREG&("!CARDEFAC")
                Z$ = REGLEIDO$("!CARDEFAC", JJ&)
                If CVI(Mid$(Z$, 5, 2)) = CI% Then
                  CAN = CAN + CVS(Mid$(Z$, 1, 4))
                  Call REPLA(Z$, MKS$(CAN), 1, 4)
                  Call GRAREG("!CARDEFAC", Z$, JJ&)
                  GoTo 1201
                End If
              Next JJ&
              '
              Z$ = REGBLAN$("!CARDEFAC")
              Call REPLA(Z$, MKS$(CAN), 1, 4)      ' cantidad
                   COD$ = CODEXT$(CI%)
              Call REPLA(Z$, MKI$(CI%), 5, 2)
              Call REPLA(Z$, DESCRIT$(CI%), 7, 64)
              Call REPLA(Z$, Mid$(X$, 5, 18), 387, 16) ' numero de remito
              PRECIOD# = PRELISTA#(ARCHIPRE$, CI%)
              Call REPLA(Z$, MKS$(PRECIOD#), 421, 4)
              Call REPLA(Z$, MKS$(PRECIOD#), 429, 4)
              Call REPLA(Z$, MKS$(PORCEIVA(CI%)), 441, 4)
              Call REPLA(Z$, MKS$(0), 505, 4)
              Call REPLA(Z$, MKS$(0), 509, 4)
              '
              Call REGAPP("!CARDEFAC", Z$)
              CAITCAR% = CAITCAR% + 1
              '
           End If
1201     Next UU&
         '
         JJ& = 0
         For U& = 1 To ULTREG&("!CARDEFAC")
            Z$ = REGLEIDO$("!CARDEFAC", U&)
            If CVS(Mid$(Z$, 1, 4)) >= 0.5 Then
              JJ& = JJ& + 1
              Call GRAREG("!CARDEFAC", Z$, JJ&)
            End If
         Next U&
         Call SETULTREG("!CARDEFAC", JJ&)
         Call CIERRARCH("CARDEFAC")
         '
         If JJ& > 0 Then
            If COMALTER%("Hay Consignaciones Pendientes\\Facturar Consignaciones\Continuar") = 1 Then
              CONSICLI% = 1
              GoTo 1210
            End If
         End If
         '
      End If
      SELORDESPA% = 1
      RPEDX$ = RECFILT$("ORDESPA", 5, MKI$(NC%))
      NRODESA& = 0
      FIN& = Len(RPEDX$)
      For UU& = 1 To Len(RPEDX$) Step 4
        U& = CVS(Mid$(RPEDX$, UU&, 4))
        If FIN& > 0 Then
            Call TRACE(20, UU&, FIN&)
        End If
        X$ = REGLEIDO$("ORDESPA", U&)
        If CVI(Mid$(X$, 13, 2)) = NC% Then
          CAN = CVS(Mid$(X$, 79, 4))
          CADESPACH = CVS(Mid$(X$, 83, 4))
          If CADESPACH < 0.05 Then
            If CAN >= 0.05 Then
              '
              If TIDOC% <> 3 And VERISTO% = 1 Then
                CI% = CVI(Mid$(X$, 37, 2))
                If CI% > 0 Then
                    STACT = STOCKACT(CI%)
                    STDIS = STACT - CAN
                    If STACT < 0.05 Then
                        GoTo 1202
                    End If
                End If
              End If
              '
              NRODES& = CVS(Left$(X$, 4))
              If NRODES& <> NRODESA& Then
                  TTXX$ = AJUSTI$("OD-" + TRIM$(CSTRING$(MKS$(NRODES&), 3, 8, 0, 2)), 11)
                  FEX = CVI(Mid$(X$, 31, 2))
                  TTXX$ = TTXX$ + "   Sol." + FECHATEX$(FEX)
                  FORSELPED.LIPEDPE.AddItem TTXX$
              End If
              NRODESA& = NRODES&
              CAPED% = CAPED% + 1
              FENTRE% = CVI(Mid$(X$, 31, 2))
              If FENTRE% < FINF Then FINF = FENTRE%
              If FENTRE% > FSUP Then FSUP = FENTRE%
            End If
          End If
        End If
1202  Next UU&
      GoTo 1204
      '
12021 EPAC$ = EMPREAC$
      If InStr(EPAC$, "NOVOESPU") > 0 Then GoTo 12031
      '
      FIN& = ULTREG&("PEDDETP")
      RPEDX$ = RECFILT$("PEDDETP", 3, MKI$(NC%))
      NROPEDA& = 0
      FIN& = Len(RPEDX$)
      For UU& = 1 To Len(RPEDX$) Step 4
        U& = CVS(Mid$(RPEDX$, UU&, 4))
        If FIN& > 0 Then
            Call TRACE(20, UU&, FIN&)
        End If
        X$ = REGLEIDO$("PEDDETP", U&)
        If CVI(Mid$(X$, 7, 2)) = NC% Then
          STATUS% = Asc(Mid$(X$, 27, 1))
          If STATUS% <= 1 Then
            CAN = CVS(Mid$(X$, 79, 4)) - CVS(Mid$(X$, 95, 4))
            If CAN > 0.05 Then
              '
              If TIDOC% <> 3 And VERISTO% = 1 Then
                CI% = CVI(Mid$(X$, 33, 2))
                If CI% > 0 Then
                    STACT = STOCKACT(CI%)
                    STDIS = STACT - CAN
                    If STACT < 0.05 Then
                        GoTo 1203
                    End If
                End If
              End If
              '
              NROPED = CVS(Left$(X$, 4))
              If NROPED <> NROPEDA Then
                  TTXX$ = CSTRING$(MKS$(NROPED), 3, 8, 0, 2)
                  FEX = CVI(Mid$(X$, 5, 2))
                  TTXX$ = TTXX$ + "          " + FECHATEX$(FEX)
                  FORSELPED.LIPEDPE.AddItem TTXX$
              End If
              NROPEDA = NROPED
              CAPED% = CAPED% + 1
              FENTRE% = CVI(Mid$(X$, 25, 2))
              If FENTRE% < FINF Then FINF = FENTRE%
              If FENTRE% > FSUP Then FSUP = FENTRE%
            End If
          End If
        End If
1203  Next UU&
      '
12031 RREMX$ = RECFILT$("REMPEN", 1, MKI$(NC%))
      NROREMA& = 0
      For UU& = 1 To Len(RREMX$) Step 4
        U& = CVS(Mid$(RREMX$, UU&, 4))
        X$ = REGLEIDO$("REMPEN", U&)
        If CVI(Left$(X$, 2)) = NC% Then
           CAN = CVS(Mid$(X$, 27, 4))
           If CAN > 0.05 Then
              NROREM = CVS(Mid$(X$, 3, 4))
              If NROREM <> NROREMA Then
                 TTXX$ = AJUSTI$("RT-" + TRIM$(CSTRING$(MKS$(NROREM), 3, 8, 0, 2)), 11)
                 FEX = CVI(Mid$(X$, 7, 2))
                 TTXX$ = TTXX$ + "       " + FECHATEX$(FEX)
                 FORSELPED.LIPEDPE.AddItem TTXX$
              End If
              NROREMA = NROREM
              CAPED% = CAPED% + 1
           End If
        End If
      Next UU&
      '
1204  Screen.MousePointer = 1
      If CAPED% < 1 Then
          Call COMUNI("No hay Pedidos o Remitos Pendientes\para este Cliente")
          GoTo 1210
      End If
      '
      FORSELPED.Show 1
      '
      FORFACNEU.Picture1.Refresh
      CAPESEL% = 0
      For U& = 1 To FORSELPED.LIPEDPE.ListCount
          If FORSELPED.LIPEDPE.Selected(U& - 1) = True Then
              FORSELPED.LIPEDPE.ListIndex = U& - 1
              CAPESEL% = CAPESEL% + 1
              TTXX$ = FORSELPED.LIPEDPE.TEXT
              NUPEDSEL&(CAPESEL%) = Val(Left$(TTXX$, 10))
              If Left$(TTXX$, 3) = "OD-" Then
                  NUPEDSEL&(CAPESEL%) = Val(Mid$(TTXX$, 4, 10))
                ElseIf Left$(TTXX$, 3) = "RT-" Then
                  NUPEDSEL&(CAPESEL%) = (-1) * Val(Mid$(TTXX$, 4, 10))
              End If
          End If
      Next U&
      '
      If CAPESEL% < 1 Then GoTo 1210
      '
'********************************************* ingresa detalle de facturacion
      '
1205  TINI = Timer
      ARCHIPRE$ = TRIM$(Str$(LPRE%))
      While Len(ARCHIPRE$) < 3: ARCHIPRE$ = "0" + ARCHIPRE$: Wend
      ARCHIPRE$ = "LIPRE" + ARCHIPRE$
      If LPRE% < 1 Then ARCHIPRE$ = "LIPRE001"
      '
      Call CIERRARCH("!CARDEFAC")
      Call BLANARCH("!CARDEFAC")
      If TIDOC% <> 1 Then GoTo 1211
      '
      CAITCAR% = 0
      FIN& = Len(RPEDX$)
      For UU& = 1 To Len(RPEDX$) Step 4
        U& = CVS(Mid$(RPEDX$, UU&, 4))
        If MAXITEMS% > 0 Then
            If CAITCAR% >= MAXITEMS% Then
                GoTo 1210
            End If
        End If
        '
        If FIN& > 0 Then
            Call TRACE(20, UU&, FIN&)
        End If
        '
' revisar a partir de aqui
        '
        REGODES& = 0
        REGPEDP& = U&
        If SELORDESPA% = 1 Then
             REGODES& = U&
             X1$ = REGLEIDO$("ORDESPA", U&)
             REGPEDA& = CVS(Mid$(X1$, 111, 4))
             X2$ = REGLEIDO$("PEDDETA", REGPEDA&)
             REGPEDP& = CVI(Mid$(X2$, 31, 2))
        End If
        X$ = REGLEIDO$("PEDDETP", REGPEDP&)
        '
        If CVI(Mid$(X$, 7, 2)) = NC% Then
          STATUS% = Asc(Mid$(X$, 27, 1))
          If STATUS% <= 1 Or SELORDESPA% = 1 Then
            '
            If SELORDESPA% = 1 Then
               NODES& = CVS(Left$(X1$, 4))
               For UI% = 1 To CAPESEL%
                  If NUPEDSEL&(UI%) = NODES& Then GoTo 1207
               Next UI%
               GoTo 1209
            End If
            '
            NUPED& = CVS(Left$(X$, 4))
            For UI% = 1 To CAPESEL%
               If NUPEDSEL&(UI%) = NUPED& Then GoTo 1207
            Next UI%
            GoTo 1209
            '
1207        CAN = CVS(Mid$(X$, 79, 4)) - CVS(Mid$(X$, 95, 4))
            NUOCOPAR$ = TRIM$(Mid$(X$, 9, 16))
            NROPEDIPAR$ = TRIM$(Str$(CVS(Left$(X$, 4))))
            If SELORDESPA% = 1 Then
               CAN = CVS(Mid$(X1$, 79, 4))
               CADESPACH = CVS(Mid$(X1$, 83, 4))
               If CADESPACH >= 0.05 Then GoTo 1209
            End If
            If CAN >= 0.05 Then
              '
              CI% = CVI(Mid$(X$, 33, 2))
              If VERISTO% = 1 Then
                If CI% > 0 Then
                    STACT = STOCKACT(CI%)
                    STDIS = STACT - CAN
                    If STDIS < 0 Then
                      If TIDOC% <> 3 Then
                        CAN = STACT
                        If CAN < 0.05 Then GoTo 1209
                      End If
                    End If
                End If
              End If
              '
1208          Z$ = REGBLAN$("!CARDEFAC")
              Call REPLA(Z$, MKS$(CAN), 1, 4)      ' cantidad
                   COD$ = CODEXT$(CI%)
              Call REPLA(Z$, MKI$(CI%), 5, 2)
              Call REPLA(Z$, DESCRIT$(CI%), 7, 64)
                   TACOMED$ = Mid$(X$, 51, 8) + Mid$(X$, 35, 12) + Mid$(X$, 59, 12)
              Call REPLA(Z$, TACOMED$, 283, 32)
                   PRELI = CVS(Mid$(X$, 71, 4))
                   If PRELI < 0.0001 Then
                      PRELI = PRELISTA#(ARCHIPRE$, CI%)
                   End If
                   PRECIO = CVS(Mid$(X$, 83, 4))
                   If PRELI < PRECIO Then PRELI = PRECIO
                   PORDESCUEN = 0: If PRELI >= 0.0001 Then PORDESCUEN = 100 * (PRELI - PRECIO) / PRELI
              Call REPLA(Z$, MKS$(PRELI), 421, 4)
              Call REPLA(Z$, MKS$(PORDESCUEN), 425, 4)
              Call REPLA(Z$, MKS$(PRECIO), 429, 4)
                   IMPO# = CDbl(PRECIO) * CAN
              Call REPLA(Z$, MKD$(IMPO#), 433, 8)
              Call REPLA(Z$, MKS$(PORCEIVA(CI%)), 441, 4)
              Call REPLA(Z$, MKS$(REGODES&), 505, 4)
              Call REPLA(Z$, MKS$(REGPEDP&), 509, 4)
              '
              Call REGAPP("!CARDEFAC", Z$)
              CAITCAR% = CAITCAR% + 1
              '
              If NUOCOPAR$ <> "" Then
                 If NUOCOPAR$ <> NUOCOPARA$ Then
                    If NUOCOMPRA$ <> "" Then NUOCOMPRA$ = NUOCOMPRA$ + ", "
                    NUOCOMPRA$ = NUOCOMPRA$ + NUOCOPAR$
                    NUOCOPARA$ = NUOCOPAR$
                 End If
              End If
              '
              If NROPEDIPAR$ <> "" Then
                 If NROPEDIPAR$ <> NROPEDIPARA$ Then
                    If NROPEDIDO$ <> "" Then NROPEDIDO$ = NROPEDIDO$ + ", "
                    NROPEDIDO$ = NROPEDIDO$ + NROPEDIPAR$
                    NROPEDIPARA$ = NROPEDIPAR$
                 End If
              End If
              '
            End If
          End If
        End If
1209  Next UU&
      '
      REGODES& = 0
      For UU& = 1 To Len(RREMX$) Step 4
        U& = CVS(Mid$(RREMX$, UU&, 4))
        If MAXITEMS% > 0 Then
            If CAITCAR% >= MAXITEMS% Then
                GoTo 1210
            End If
        End If
        '
        X$ = REGLEIDO$("REMPEN", U&)
        If CVI(Left$(X$, 2)) = NC% Then
           NUREM& = CVS(Mid$(X$, 3, 4))
           For UI% = 1 To CAPESEL%
               If NUPEDSEL&(UI%) = (-1) * NUREM& Then GoTo 12071
           Next UI%
           GoTo 12091
           '
12071      CAN = CVS(Mid$(X$, 27, 4))
           If CAN > 0.05 Then
              '
              CI% = CODINT%(Mid$(X$, 11, 16))
              MEDIX$ = Mid$(X$, 85, 12)
              '
              Z$ = REGBLAN$("!CARDEFAC")
              Call REPLA(Z$, MKS$(CAN), 1, 4)      ' cantidad
                   COD$ = CODEXT$(CI%)
              Call REPLA(Z$, MKI$(CI%), 5, 2)
              Call REPLA(Z$, DESCRIT$(CI%), 7, 64)
              If CI% < 1 Then Call REPLA(Z$, Mid$(X$, 33, 64), 7, 64)
              Call REPLA(Z$, TRIM$(Str$(NUREM&)), 387, 16) ' numero de remito
              Call REPLA(Z$, MEDIX$, 303, 12)
              PRECIOD# = PRELISTA#(ARCHIPRE$, CI%)
              Call REPLA(Z$, MKS$(PRECIOD#), 421, 4)
              Call REPLA(Z$, MKS$(PRECIOD#), 429, 4)
              Call REPLA(Z$, MKS$(PORCEIVA(CI%)), 441, 4)
              Call REPLA(Z$, MKS$(REGODES&), 505, 4)
              Call REPLA(Z$, MKS$(U& * (-1)), 509, 4)
              '
              Call REGAPP("!CARDEFAC", Z$)
              CAITCAR% = CAITCAR% + 1
              '
           End If
        End If
12091  Next UU&
      '
1210  Z$ = ""
      Call CIERRARCH("!CARDEFAC")
      '
1211  TOTAFAC#(1, 9) = 0
      '
1215  ARCHIPRE$ = TRIM$(Str$(LPRE%))
      While Len(ARCHIPRE$) < 3: ARCHIPRE$ = "0" + ARCHIPRE$: Wend
      ARCHIPRE$ = "LIPRE" + ARCHIPRE$
      If LPRE% < 1 Then ARCHIPRE$ = "LIPRE001"
      '
      ATRI$ = "/COLORPAN=B-N/PS/TITUPAN=" + VARIFA$
      ATRI$ = ATRI$ + "/BORDESUP=5/BORDEINF=19/BORDEIZQ=1"
      If InStr(UCase$(FORVENTA$), "LIB") < 1 Then ATRI$ = ATRI$ + "/ARCHECO(2)=" + ARCHIPRE$
      ' lo anterior para permitir facturacion libre
      XI% = VENTABU%(FORVENTA$ + ATRI$)
      If XI% < 0 Then
          GoTo 3500
          FORFACNEU.BOTNEXT.Enabled = True
          FORFACNEU.Text1.SetFocus
          Exit Sub
      End If
      '
      FORFACNEU.Picture1.Refresh
      Call CIERRARCH("*.*")
      '
      If ULTREG&("!CARDEFAC") < 1 Then GoTo 3500
      '
      Screen.MousePointer = 11
      ATX% = 0: CIANT% = 0: FIN& = ULTREG&("!CARDEFAC")
      '
      SUCUVEN% = 0
      On Error Resume Next
      SUCUVEN% = Val(FORFACNEU.Label12.Caption)
      On Error GoTo 0
      '
      DEPOSALI% = 1
      If SUCUVEN% > 1 Then
         For IU& = 1 To ULTREG&("SUCREMOT")
            X$ = REGLEIDO$("SUCREMOT", IU&)
            If CVI(Left$(X$, 2)) = SUCUVEN% Then
               DEPOSALI% = Asc(Mid$(X$, 33, 1))
               If ECOARCH$("TADEPOSI", Chr$(DEPOSALI%)) = "" Then DEPOSALI% = 1
            End If
         Next IU&
      End If
      Call CIERRARCH("SUCREMOT")
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
         If VERISTO% = 1 Then
           If CVS(Mid$(X$, 509, 4)) >= 0 Then ' suprime remitos
             If CI% > 0 Then
               STDIS = STODEPOS(CI%, DEPOSALI%) - CAN
               If STDIS < 0 Then
                  If TIDOC% < 3 Then
                     Call COMUNI("Factura no Procesable.\Stock Negativo en Código " + TRIM$(CODEXT$(CI%)))
                     GoTo 1215
                  End If
               End If
             End If
           End If
         End If
         '
         PRECIOD# = CDbl(CVS(Mid$(X$, 421, 4)))
         PORDESCUEN = CVS(Mid$(X$, 425, 4))
         '
         If PRECIOD# < 0.00005 Then
            PRECIOD# = PRELISTA#(ARCHIPRE$, CI%)
         End If
         '
         If CAN > 0.05 Then
           If CI% > 0 Then
             If CI% <= NUMAS% Then
               Call REPLA(X$, MKS$(PRECIOD#), 421, 4)
               Call REPLA(X$, MKS$(PRECIOD# * (100 - PORDESCUEN) / 100), 429, 4)
               Call REPLA(X$, MKS$(PORCEIVA(CI%)), 441, 4)
               DE$ = DESCRIT$(CI%)
               Call REPLA(X$, DE$, 7, 64)
             End If
           End If
         End If
         '
         NORI$ = NUORIG$(ARCHIPRE$, CI%)
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
                     PORDESCUEN = CVS(Mid$(X$, 425, 4))
                     Call REPLA(X$, MKS$(PREINFO * (100 - PORDESCUEN) / 100), 429, 4)
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
'VOY A CALCULAR EL IVA
      GoSub 8200                          ' calcular IVA de cada item
      '
      GoSub 8700                                        ' presentar totales
      TITU$ = "FORMATO NORMAL": If TIFORM% = 2 Then TITU$ = "FORMATO RESUMEN" Else If TIFORM% = 3 Then TITU$ = "FORMATO TEXTO LIBRE"
1330  '
      FACIVA = 1
      CATIVAS$ = "A"
      VA% = 1
      If PIVA% = 0 Or PIVA% = 3 Or PIVA% = 4 Or PIVA% = 6 Then
           On Error Resume Next
           FACIVA = TOTAFAC#(1, 8) / (TOTAFAC(1, 5) + TOTAFAC#(1, 9))
           On Error GoTo 0
           CATIVAS$ = "B"
           VA% = 2
           '
          ElseIf PIVA% = 2 Then
           'ALICU = ALI1 + ALI2
           RNI% = 1
           '
          ElseIf PIVA% = 5 Then
           'ALICU = 0
           CATIVAS$ = " "
           VA% = 0
           If LETREXPO$ = "A" Then
              VA% = 1
              CATIVAS$ = "A"
           End If
      End If
      FOPRES$ = "PRESEFAC": If FACIVA <> 1 Then FOPRES$ = "PRESFACB"
      ATRI$ = "/BORDEIZQ=1/BORDESUP=5/BORDEINF=19"
      FORFACNEU.Picture1.Visible = False
      XI% = VENTABU%(FOPRES$ + ATRI$ + "/COLORPAN=B-N/NC/TITUPAN=" + TITU$)
      FORFACNEU.Picture1.Visible = True
      If XI% < 0 Then
          GoTo 3500
          FORFACNEU.BOTNEXT.Enabled = True
          FORFACNEU.Text1.SetFocus
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
        If TIDOC% = 0 Or TIDOC% = 1 Or (TIDOC% = 3 And InStr(UCase$(TDFAC$), "DEVOLUC") > 0) Then                 ' solo facturas
            If DESCUPRE% = 0 Then
                DESC$ = TRIM$(FORFACNEU.PORDESCU.TEXT)
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
                       ALIVITEM = CVS(Mid$(X$, 441, 4))
                       '
                       IVAITEM# = TOT# * (ALIVITEM / 100)
                       BRUITEM# = TOT# + IVAITEM#
                       '
                       Call REPLA(X$, MKD$(TOT#), 433, 8)
                       Call REPLA(X$, MKS$(ALIVITEM), 441, 4)
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
2025    If TIDOC% > 1 Or InStr(UCase$(FORFACNEU.Label10.Caption), "LIBRE") > 0 Then                      ' debitos y creditos
          If FACIVA = 1 Then                      ' solo facturas 'A'
2026        VDEF$ = MKD$(TOTAFAC#(1, 5)) + MKD$(TOTAFAC#(1, 9)) + MKD$(TOTAFAC#(1, 6)) + MKD$(TOTAFAC#(1, 10)) + MKD$(TOTAFAC#(1, 7)) + MKD$(TOTAFAC#(1, 8))
            OBX$ = OBJPLA$("ENTRAIVAFAC", VDEF$)
            If OBX$ = "" Then GoTo 3500
            TOTAFAC#(1, 6) = CVD(Mid$(OBX$, 17, 8))
            TOTAFAC#(1, 10) = CVD(Mid$(OBX$, 25, 8))
            TOTAFAC#(1, 7) = 0
            If PIVA% = 2 Then TOTAFAC#(1, 7) = 0.5 * (TOTAFAC#(1, 6) + TOTAFAC#(1, 10))
            TOTAFAC#(1, 8) = TOTAFAC#(1, 5) + TOTAFAC#(1, 9) + TOTAFAC#(1, 6) + TOTAFAC#(1, 10) + TOTAFAC#(1, 7)
            If Abs(TOTAFAC#(1, 6) * TOTAFAC#(1, 10)) >= 0.005 Then
               Call MENSERR(24, "Imposible Tasas de IVA Distintas\en Documentos de Texto Libre")
               GoTo 2026
            End If
            If Abs(TOTAFAC#(1, 10)) >= 0.005 Then
               If Abs(TOTAFAC#(1, 5)) >= 0.005 Then
                  TOTAFAC#(1, 9) = TOTAFAC#(1, 5)
                  TOTAFAC#(1, 5) = 0
                  GoTo 2026
               End If
            End If
            For KKII% = 1 To 48 Step 8
              If Abs(CVD(Mid$(OBX$, KKII%, 8)) - CVD(Mid$(VDEF$, KKII%, 8))) >= 0.005 Then
                GoTo 2026
              End If
            Next KKII%
          End If
        End If
        '
        VAFLETE# = 0
        If PIVA% = 5 Then
           VAFLETE# = Val(InputBox$("Importe del Flete", "Ingrese Importe de Flete"))
           'TOTAFAC#(1, 2) = TOTAFAC#(1, 2) + VAFLETE#
           TOTAFAC#(1, 5) = TOTAFAC#(1, 5) + VAFLETE#
           TOTAFAC#(1, 8) = TOTAFAC#(1, 8) + VAFLETE#
           'TOTAFAC#(1, 9) = TOTAFAC#(1, 9) + VAFLETE#
        End If
           
        
2028    For I% = 1 To 10
            TOTAFAC#(1, I%) = (Int(100 * TOTAFAC#(1, I%) + 0.5)) / 100
            TOTAFAC#(2, I%) = TOTAFAC#(1, I%) / TICA
        Next I%
'
'****************************************************************************
'
2200    If TIDOC% > 1 Then GoTo 2900
        '
        If TOTAFAC#(1, 8) + SALDCLI#(NC%) + CHEPEACRE#(NC%) > LIMICRE(NC%) Then
           Call MENSERR(24, "Imposible Registrar esta Factura.\  \Supera el Límite de Crédito Asignado\para este Cliente.")
           GoTo 3500
        End If
        '
1320    If CPAG% > 0 Then
          VDEF$ = String$(41, 0)
          Call REPLA(VDEF$, MKD$(TOTAFAC#(1, 8)), 1, 8)
          Call REPLA(VDEF$, MKD$(TOTAFAC#(1, 8)), 33, 8)
          Call REPLA(VDEF$, Chr$(CPAG%), 41, 1)
          Screen.MousePointer = 1
          OBX$ = OBJPLA$("PEDIDO-CARCONPA", VDEF$)
          If OBX$ = "" Then GoTo 1200
          CPAG% = Asc(Mid$(OBX$, 41, 1))
          CONPA$ = TRIM$(DECONPA$(CPAG%))
          If CONPA$ = "" Then
            Call MENSERR(24, "Condicion de Venta no Válida")
            GoTo 1320
          End If
          CONPA$ = Left$(CONPA$, 1) + LCase$(Mid$(CONPA$, 2))
          FORFACNEU.TEXT21(1).TEXT = Str$(CPAG%)
        End If
        '
        RETRAN& = 0: TRXP$ = Space$(256)
        RFF$ = RECFILT$("TRANSCLI", 1, MKI$(NC%))
        If Len(RFF$) >= 4 Then
           RETRAN& = CVS(RFF$)
           If RETRAN& < 1 Or RETRAN& > ULTREG&("TRANSCLI") Then RETRAN& = 0
           If RETRAN& > 0 Then
              TRXP$ = Mid$(REGLEIDO$("TRANSCLI", RETRAN&), 2)
              Call REPLA(TRXP$, Space$(6), 153, 6)
           End If
        End If
        TRXP$ = OBJPLA$("TRANSPOR", TRXP$) + Space$(256)
        If TRIM$(TRXP$) <> "" Then
           TRYP$ = MKI$(NC%) + TRXP$
           Call REPLA(TRYP$, Space$(6), 155, 6)
           If RETRAN& > 0 Then
               Call GRAREG("TRANSCLI", TRYP$, RETRAN&)
             Else
               Call REGAPP("TRANSCLI", TRYP$)
           End If
           Call CIERRARCH("TRANSCLI")
        End If
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
        TCMP$ = TRIM$(TIPCOM$)
        If InStr(UCase$(TIPCOM$), "FACTURA") > 0 Then
          If InStr(UCase$(TIPCOM$), "CONTADO") < 1 Then
            If CPAG% < 1 Or InStr(UCase$(DECONPA$(CPAG%)), "CONTADO") > 0 Then
              TCMP$ = TCMP$ + " CONTADO"
            End If
          End If
        End If
        TICOMPROBAN$ = TCMP$
        FINFACTURA.Label2.Caption = TCMP$
        FINFACTURA.Label1(0) = TCMP$ + ":"
        FINFACTURA.Label4.Caption = CATIVAS$
        '
        CODFOR$ = FORIMPRE$
        NROCON& = NRO
        NCN$ = TRIM$(Str$(NROCON&))
        While Len(NCN$) < 8: NCN$ = "0" + NCN$: Wend
        PREFI$ = ATRIFORM$(CODFOR$, "PREFIFAC")
        If SUCUVEN% > 0 Then
           If PIVA% <> 5 Then
              PREFI$ = FORFACNEU.Label12.Caption
           End If
        End If
        NCN$ = PREFI$ + "-" + NCN$
        NUCOMPROBAN$ = NCN$
        FINFACTURA.Label3.Caption = NCN$
        '
        COPIASFAC% = 2: COPIASREM% = 2
        If InStr(TCMP$, "CONTADO") > 0 Then
           COPIASFAC% = 3: COPIASREM% = 0
        End If
        If TIDOC% > 1 Then
           COPIASFAC% = 2: COPIASREM% = 0
        End If
        If InStr(EMPREAC$, "NEUMANN") > 0 Then
           SUCUVEN% = 0
           On Error Resume Next
           SUCUVEN% = Val(FORFACNEU.Label12.Caption)
           If SUCUVEN% > 1 Then
              COPIASREM% = 0
           End If
           If CONSICLI% > 0 Then
              COPIASREM% = 0
           End If
        End If
        '
        FINFACTURA.Text2(0) = TRIM$(Str$(COPIASFAC%))
        FINFACTURA.Text2(1) = TRIM$(Str$(COPIASREM%))
        FINFACTURA.Label5.Caption = ""
        For U& = 0 To 4: FINFACTURA.Text1(U&).TEXT = "": Next U&
        '
        FINFACTURA.Show 1
        '
        If FINFACTURA.Label5.Caption <> "APRO" Then
           If COMALTER%("Realmente Quiere CANCELAR este Documento\\No, Continuar\Si, Cancelar") <> 2 Then GoTo 2900
           GoTo 3500
        End If
        '
        COPIASFAC% = Val(FINFACTURA.Text2(0))
        If COPIASFAC% < 1 Or COPIASFAC% > 6 Then GoTo 2900
        COPIASREM% = Val(FINFACTURA.Text2(1))
        If COPIASREM% > 6 Then GoTo 2900
        GoTo 2910
        '
        'XI% = COMALTER%("Fin de Carga de " + TIPCOM$ + " " + TRIM$(Str$(NRO)) + "\\Aprobar\Corregir\Anular")
        ' FORFACNEU.Refresh
        'On XI% GoTo 2910, 2905, 3500
        'Beep
        'GoTo 2900
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
2918    Call BLOQARCH("SVD202")
        Call BLOQARCH("DEUDOR2")
        Call BLOQARCH("FACTUR")
        Call BLOQARCH("MOVISTO")
        Call BLOQARCH("SALSTOCK")
        '
        URE2021& = ULTREG&("SVD202")
        UREFAC1& = ULTREG&("FACTUR")
        UREMOV1& = ULTREG&("MOVISTO")
        '
        If (NC% > 0 And NC% < 10000) Then
             '
             'Call BLIMESS("Calculando Vencimientos ...")
             Call GRADEU(LU$, ARCHIDEU$, RD%, TIPO%, VA%, NRO, NC%, FFI%, CPAG%, TOTAFAC#(1, 8))
             URE2022& = ULTREG&("SVD202")
             Call CIERRARCH("SVD202")
             '
             SIGX% = 1: If TIPO% = 6 Then SIGX% = (-1)
             TTF# = TOTAFAC#(1, 8) * SIGX%
             '
             'Call BLIMESS("Actualizando Saldo de Cuenta ...")
             Call GRACLI(LU$, ARCHICLI$, NC%, TTF#, DATCLI$(), 2)
             Call CIERRARCH("DEUDOR2")
             '
        End If
        '
        FLIM% = FECHANUM("01/01/99")
        If ULTREG&("FORMFAC") > 0 Then
            FFS$ = REGLEIDO$("FORMFAC", 2&)
            FLIM% = CVI(Mid$(FFS$, 19, 2))
        End If
        '
        Call GRAFAC98(LU$, ARCHIFAC$, 3, RD%, FLIM%, TIPO%, VA%, CAFOR%, NRO, NC%, CLI$, IVAS$, FFI%, FVI%, LPRE%, VENDEX%, TOTAFAC#())
        Call GRAFAC98(LU$, ARCHIFAC$, 4, RD%, FLIM%, TIPO%, VA%, CAFOR%, NRO, NC%, CLI$, IVAS$, FFI%, FVI%, LPRE%, VENDEX%, TOTAFAC#())
        UREFAC2& = ULTREG&("FACTUR")
        '
        'Call BLIMESS("Grabando Movimientos de Material ...")
        NRO1 = NRO: If TIPO% = 6 Then NRO1 = (-1) * NRO
        '
        REMI = NRO: NROREMI$ = "": NUREMIFAC& = 0
        If InStr(EMPREAC$, "NEUMANN") > 0 Then
           SUCUVEN% = 0
           On Error Resume Next
           SUCUVEN% = Val(FORFACNEU.Label12.Caption)
           On Error GoTo 0
           '
           If TIDOC% <= 1 Then
             If SUCUVEN% > 1 Then
                 NROREMI$ = InputBox$("Número de Remito", "Ingresar Número de Remito de Sucursal")
               Else
                 If COPIASREM% > 0 Then
                   If InStr(UCase$(DECONPA$(CPAG%)), "CONTADO") < 1 Then
                     NROPRO& = 0
                     Call NUREMIFA(NROPRO&)
                     REMI = NROPRO&
                     NUREMIFAC& = REMI
                     PREFIREM$ = CONTROL$("", "PREFIREM")
                     If PREFIREM$ = "" Then PREFIREM$ = "0001"
                     NROREMI$ = TRIM$(Str$(NUREMIFAC&))
                     While Len(NROREMI$) < 8: NROREMI$ = "0" + NROREMI$: Wend
                     NROREMI$ = PREFIREM$ + "-" + NROREMI$
                   End If
                 End If
             End If
           End If
        End If
        '
        If InStr(EMPREAC$, "SABO") > 0 Then
           ' para SABO el numero es el mismo MENOS 12500
           If NRO > 12500 Then
              REMI = REMI - 12500
           End If
        End If
        '
        REMI1 = REMI: If TIPO% = 3 Then REMI1 = (-1) * REMI
        If InStr(EMPREAC$, "NEUMANN") > 0 Then
           REMI1 = NUREMIFAC&
        End If
        If CONSICLI% < 1 Then
             Call GRAM2000(NC%, NRO1, REMI1, FFI%)
           Else
            RECONSI$ = REGBLAN$("CONSICLI")
            Call REPLA(RECONSI$, MKI$(NC%), 1, 2)
            Call REPLA(RECONSI$, MKI$(FFI%), 3, 2)
            Call REPLA(RECONSI$, NCN$, 5, 18)
            Call REPLA(RECONSI$, String$(12, 0), 109, 12)
            Call REPLA(RECONSI$, CTRXX$, 122, 7)
            Call GRADECONSI(RECONSI$)
        End If
        UREMOV2& = ULTREG&("MOVISTO")
        Call CIERRARCH("MOVISTO")
        Call CIERRARCH("SALSTOCK")
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
        FLIM% = FECHANUM("01/01/99")
        If ULTREG&("FORMFAC") > 0 Then
            FFS$ = REGLEIDO$("FORMFAC", 2&)
            FLIM% = CVI(Mid$(FFS$, 19, 2))
        End If
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
           RTO1$ = TRIM$(Mid$(Z$, 387, 16))
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
              If NRO > 13000 Then
                 REMI = REMI - 12500
              End If
           End If
           '
           RTO$ = TRIM$(Str$(REMI))
           While Len(RTO$) < 8: RTO$ = "0" + RTO$: Wend
           PREFI$ = ATRIFORM$(FORIMPRE$, "PREFIFAC")
           RTO$ = PREFI$ + "-" + RTO$
        End If
        If TRIM$(FORFACNEU.Text10.TEXT) <> "" Then
            RTO$ = TRIM$(FORFACNEU.Text10.TEXT)
        End If
        '
        NROCON& = NRO
        VALOR# = TOTAFAC#(1, 8)
        '
        If FORMAIMPRE$ = "DOS" Then
              Call PRIFADOS
           Else
              GoSub 6000                         ' imprimir
              EPAC$ = TRIM$(UCase$(EMPREAC$))
              If InStr(EPAC$, "NEUMANN") > 0 Then
                If COPIASREM% > 0 Then
                   GoSub 9000       ' imprimir remito
                End If
              End If
        End If
        '
        FLIM% = FECHANUM("01/01/99")
        If ULTREG&("FORMFAC") > 0 Then
            FFS$ = REGLEIDO$("FORMFAC", 2&)
            FLIM% = CVI(Mid$(FFS$, 19, 2))
        End If
        '
        NRXX& = NRO + 1
        While NRXX& < NROCON& + 0.1
           NRT = NRXX&
           For I% = 0 To 3: For J% = 0 To 10: TOTAFAC#(I%, J%) = 0: Next J%: Next I%
           Call GRAFAC98(LU$, ARCHIFAC$, 3, RD%, FLIM%, TIPO%, VA%, CAFOR%, NRT, NC%, CLI$, IVAS$, FFI%, FVI%, LPRE%, VENDEX%, TOTAFAC#())
           NRXX& = NRXX& + 1
        Wend
        Call CIERRARCH("FACTUR")
        '
2990    TPXS$ = "Factura"
        If TIPO% = 3 Then TPXS$ = "Nota de Débito"
        If TIPO% = 6 Then TPXS$ = "Nota de Crédito"
        '
        GoTo 2995      ' no permite anular directamente
        '              ' por problemas NEUMANN
        '
        AI% = ALTERNA%("Revise la " + TPXS$ + " Impresa - Esta O.K.")
        If AI% <> 1 Then
           '
           For U& = URE2021& + 1 To URE2022&
              X$ = REGBLAN$("SVD202")
              Call GRAREG("SVD202", X$, U&)
           Next U&
           '
           For U& = UREFAC1& + 1 To UREFAC2&
              X$ = REGLEIDO$("FACTUR", U&)
              Call REPLA(X$, String$(32, 0), 15, 32)  ' para mantener sucursal
              Call REPLA(X$, String$(16, 0), 49, 16)  ' para mantener sucursal
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
2995    If MODESCU% = 1 Then
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
                       IVAITEM# = TOT# * (CVS(Mid$(X$, 441, 4))) / 100
                       BRUITEM# = TOT# + IVAITEM#
                       '
                       Call REPLA(X$, MKD$(TOT#), 433, 8)
                       'Call REPLA(X$, MKS$(ALICU), 441, 4)
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
        FLIM% = FECHANUM("01/01/99")
        If ULTREG&("FORMFAC") > 0 Then
            FFS$ = REGLEIDO$("FORMFAC", 2&)
            FLIM% = CVI(Mid$(FFS$, 19, 2))
        End If
        '
        Call BLOQARCH("ORDESPA")
        Call BLOQARCH("DESPAIMP")
        Call BLOQARCH("PEDDETP")
        Call BLOQARCH("PEDDETA")
        '
      ' blanquea pedidos pendientes ********************
        '
        If TIDOC% = 1 Then
          Call BLAPEFAC(REMI, FFI%)
          Call CIERRARCH("PEDDETA")
          Call CIERRARCH("PEDDETP")
          Call CIERRARCH("ORDESPA")
        End If
        '
      ' imputacion contable de ventas
        '
        If GRAICON% = 1 Then
           '
           XD1$ = REGLEIDO$("ICOVEN", 1)
           MMXX$ = "ENEFEBMARABRMAYJUNJULAGOSETOCTNOVDIC"
           FEX = FFI%: FEFA$ = FECHATEX$(FEX)
           MESI% = Val(Mid$(FEFA$, 4, 2))
           ARCHICO$ = "ICV" + Mid$(MMXX$, 3 * MESI% - 2, 3) + Mid$(FEFA$, 7)
           Call COPYSTRU("ICOVEN", ARCHICO$)
           Call REGAPP(ARCHICO$, XD$)
           '
        End If
        '
        VARIFA$ = "": ATRI$ = ""
        '
        Z$ = ""
        If TIDOC% = 0 Then  ' venta contado
           Call CONECRUN("RECISTA/" + TRIM$(Str$(NC%)), "Recibo de Cobranza/WAIT")
        End If
3500  ' aqui limpia pantalla y vectores ****************
        '
        Call CIERRARCH("*.*")
        '
        FORFACNEU.BOTNEXT.Enabled = True
        NC% = 0
        Call BLANARCH("!CARDEFAC")
        '
        FORFACNEU.Text1.TEXT = "": FORFACNEU.Text1.Refresh
        FORFACNEU.Text2.TEXT = "": FORFACNEU.Text2.Refresh
        FORFACNEU.Text3.TEXT = "": FORFACNEU.Text3.Refresh
        FORFACNEU.Text4.TEXT = "": FORFACNEU.Text4.Refresh
        FORFACNEU.Text5.TEXT = "": FORFACNEU.Text5.Refresh
        FORFACNEU.Text6.TEXT = "": FORFACNEU.Text6.Refresh
        FORFACNEU.Text7.TEXT = "": FORFACNEU.Text7.Refresh
        FORFACNEU.Text8.TEXT = "": FORFACNEU.Text8.Refresh
        FORFACNEU.Text9.TEXT = "": FORFACNEU.Text9.Refresh
        FORFACNEU.Text10.TEXT = "": FORFACNEU.Text10.Refresh
        FORFACNEU.Text11.TEXT = "": FORFACNEU.Text11.Refresh
        '
        For KK% = 0 To 2
           FORFACNEU.TEXT21(KK%).TEXT = ""
           FORFACNEU.Eco21(KK%) = ""
        Next KK%
        '
        For KK% = 0 To 5
           FORFACNEU.Total(KK%).TEXT = ""
        Next KK%
        '
        FORFACNEU.PORDESCU.TEXT = ""
        FORFACNEU.PORIVA(0).TEXT = ""
        FORFACNEU.PORIVA(1).TEXT = ""
        FORFACNEU.PORIVA(2).TEXT = ""
   For U& = 1 To 2
      X$ = REGLEIDO$("GRUPIVA", U&)
      FORFACNEU.PORIVA(U& - 1) = CSTRING$(MKS$(TAIVA(U& - 1)), 3, 6, 2, 2)
   Next U&
   Call CIERRARCH("GRUPIVA")
        FFI% = HOY(HO$)
        FORFACNEU.Text9.TEXT = HO$
        Call FRESHALL
        '
        FORFACNEU.Text1.SetFocus
        Screen.MousePointer = 1
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
        If ALTERNA%("&17.60&-Confirma Anulación") <> 1 Then GoTo 3500
        If REFA& > 0 Then
            If REFA& <= ULTREG&("FACTUR") Then
               X$ = REGLEIDO$("FACTUR", REFA&)
               Call REPLA(X$, String$(32, 0), 15, 32)  ' para mantener sucursal
               Call REPLA(X$, String$(16, 0), 49, 16)  ' para mantener sucursal
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
        If COPIASFAC% > 0 Then CACOPI% = COPIASFAC%
        '
        'Printer.Orientation = 1
        Call SETSPOOL("ORIENTATION", "1")
        If ORIENTA% = 2 Then
            'Printer.Orientation = 2
            Call SETSPOOL("ORIENTATION", "2")
        End If
        '
        Call COMUNI("Verifique Encuadre y Pulse\<ENTER> para Continuar ...")
        PRIDET0& = 1: HOJACONT% = 0: TOTITEMS = 0
        TRANSPORTE# = 0: HOJNUM% = 0
        '
6005    TRANSPOR% = 0
        ORDECOP% = 0
           For COP% = 1 To CACOPI%
              PRIDET& = PRIDET0&
              'Printer.Orientation = 1
              Call SETSPOOL("ORIENTATION", "1")
              If ORIENTA% = 2 Then
                 'Printer.Orientation = 2
                 Call SETSPOOL("ORIENTATION", "2")
              End If
              '
              Call CARFORWIN
              If COP% = 1 Then HOJNUM% = HOJNUM% + 1
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
              'Printer.EndDoc
              Call SETSPOOL("ENDDOC", "")
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
            'Printer.EndDoc
            Call SETSPOOL("ENDDOC", "")
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
              'Printer.EndDoc    ' imprime hoja
              Call SETSPOOL("ENDDOC", "")
              '
           Next COP%
           If CONTIHOJA% = 0 Then
               NROCON& = NROCON& + 1
           End If
        End If
        '
6007    DESCRIDOC$ = TICOMPROBAN$ + " " + NUCOMPROBAN$ + " - " + TRIM$(FORFACNEU.Text2.TEXT)
        FEX = FFI%
        DESCRIDOC$ = DESCRIDOC$ + " - " + FECHATEX$(FEX)
        '
        Call SAVESPOOL(DESCRIDOC$, OKX%)
        If ALTERNA%("-Repite Impresion") = 1 Then
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
        TCMP$ = TRIM$(UCase$(TIPCOM$)): If COP% > ULTCOFAC% Then TCMP$ = "REMITO"
        If InStr(UCase$(TIPCOM$), "FACTURA") > 0 Then
          If InStr(UCase$(TIPCOM$), "CONTADO") < 1 Then
            If CPAG% < 1 Or InStr(UCase$(DECONPA$(CPAG%)), "CONTADO") > 0 Then
              TCMP$ = TCMP$ + " CONTADO"
            End If
          End If
        End If
        Call PRINTFORWIN(TCMP$, "TIPO-COM", 0, 0)
        Call PRINTFORWIN(TRIM$(Str$(HOJNUM%)), "HOJA-NUM", 0, 0)
        '
        Call PRINTFORWIN(FECHATEX$(FEX), "FECH-EMI", 0, 0)
        Call PRINTFORWIN(FORFACNEU.Text2.TEXT, "RASO-CLI", 0, 0)
        CUTTT$ = TRIM$(FORFACNEU.Text8.TEXT): If CUTTT$ = "" Then CUTTT$ = "NO POSEE"
        Call PRINTFORWIN(CUTTT$, "CUIT-CLI", 0, 0)
        Call PRINTFORWIN(FORFACNEU.Text3.TEXT, "DOMI-CLI", 0, 0)
        Call PRINTFORWIN(NIBRCLI$(NC%), "IBRU-CLI", 0, 0)
        Call PRINTFORWIN(FORFACNEU.Text4.TEXT, "CPOS-CLI", 0, 0)
        LOCA$ = FORFACNEU.Text5.TEXT
        If PIVA% = 5 Then LOCA$ = AJUSTI$(LOCA$, 15) + " - " + PAISCLI$(NC%)
        Call PRINTFORWIN(LOCA$, "LOCA-CLI", 0, 0)
        Call PRINTFORWIN(NPROCLI$(NC%), "PROV-CLI", 0, 0)
        Call PRINTFORWIN(Str$(NC%), "CUEN-CLI", 0, 0)
        POSIVA$ = ECOARCH$("CAPOSIV", Chr$(PIVA%))
        Call PRINTFORWIN(POSIVA$, "PIVA-CLI", 0, 0)
        TCMP$ = UCase$(CATIVAS$): If COP% > ULTCOFAC% Then TCMP$ = "X"
        Call PRINTFORWIN(TCMP$, "CIVA-CLI", 0, 0)
        Call PRINTFORWIN(DECONPA$(CPAG%), "COND-PAG", 0, 0)
        '
        If NROREMI$ <> "" Then RTO$ = NROREMI$
        If TIDOC% <= 1 Then
          If VA% = 1 Then       ' SOLO FACTURAS "A"
            Call PRINTFORWIN(RTO$, "NRO-REMI", 0, 0)
          End If
        End If
        Call PRINTFORWIN(NUOCOMPRA$, "NRO-OCOM", 0, 0)
        Call PRINTFORWIN(NROPEDIDO$, "NRO-AENT", 0, 0)
        '
        Call PRINTFORWIN(Mid$(TRXP$, 1, 32), "TRANSPOR", 0, 0)
        Call PRINTFORWIN(Mid$(TRXP$, 33, 32), "DOMI-TRA", 0, 0)
        Call PRINTFORWIN(Mid$(TRXP$, 159, 32), "LOCA-TRA", 0, 0)
        Call PRINTFORWIN(Mid$(TRXP$, 65, 16), "CUIT-TRA", 0, 0)
        Call PRINTFORWIN(Mid$(TRXP$, 81, 16), "PATE-TRA", 0, 0)
        Call PRINTFORWIN(Mid$(TRXP$, 97, 32), "COND-TRA", 0, 0)
        '
        NCN$ = TRIM$(Str$(NROCON&))
        While Len(NCN$) < 8: NCN$ = "0" + NCN$: Wend
        PREFI$ = ATRIFORM$(CODFOR$, "PREFIFAC")
        If SUCUVEN% > 0 Then
           If PIVA% <> 5 Then
              PREFI$ = FORFACNEU.Label12.Caption
           End If
        End If
        NCN$ = PREFI$ + "-" + NCN$
        DOCUORIG$ = NCN$
        Call PRINTFORWIN(NCN$, "NUME-COM", 0, 0)
        COPXX% = COP%: If COPXX% > ULTCOFAC% Then COPXX% = COP% - ULTCOFAC%
        NCN$ = ORCOP$(COPXX%)
        Call PRINTFORWIN(NCN$, "ORD-COPI", 0, 0)
        '
        If InStr(EMPREAC$, "NEUMANN") > 0 Then
            YAC = 0
            For U& = 0 To 4
               TXOB$ = Space$(20)
               On Error Resume Next
               TXOB$ = FINFACTURA.Text1(U&).TEXT
               On Error GoTo 0
               Call GRAREG("!OBSERVOF", TXOB$, U& + 1)
               Call PRINTFORWIN(TXOB$, "OBSERVA", 0, YAC)
               YAC = YAC + INTERLI
            Next U&
        End If
        Call CIERRARCH("!OBSERVOF")
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
        If PRIDET& > 1 Then
          If TRANSPORTE# > 0 Then
            TXT$ = "TRANSPORTE ................................................"
              Call APPFORM(TXT$, "DES-MAT")
            TXT$ = Str$(TRANSPORTE#)   ' total neto
              Call APPFORM(TXT$, "NET-ITEM")
            YAC = YAC + 2 * INTERLI
          End If
        End If
        
        For U& = PRIDET& To ULTREG&("!CARDEFAC")
            '
            If YAC > DETHASMM - 4 * INTERLI Then                      ' SALTO DE HOJA
               YAC = YAC + INTERLI
               TXT$ = "TRANSPORTE ............................................."                 ' descripcion
                  Call APPFORM(TXT$, "DES-MAT")
               TXT$ = Str$(TRANSPORTE#)   ' total neto
                  Call APPFORM(TXT$, "NET-ITEM")
               PRIDET& = U&
               TRANSPOR% = 1
               GoTo 7390
            End If
            '
            XXX$ = REGLEIDO$("!CARDEFAC", U&)
            '
            TOTITEMS = TOTITEMS + CVS(Left$(XXX$, 4))
            TXT$ = Str$(U&)                        ' ITEM NUMERO
                Call APPFORM(TXT$, "ORD-ITEM")
                '
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
              If LPRE% > 0 Then
                 TXT$ = Str$(FACIVA * CVS(Mid$(XXX$, 421, 4))) ' precio de lista
                   Call APPFORM(TXT$, "PRE-LIST")
                   '
                 TXT$ = Str$(CVS(Mid$(XXX$, 425, 4))) ' PORCENTAJE DE DESCUENTO
                   Call APPFORM(TXT$, "PORDESCU")
              End If
              '
              TXT$ = Str$(FACIVA * CVS(Mid$(XXX$, 429, 4))) ' precio unitario
                Call APPFORM(TXT$, "PRE-NETO")
                '
              TXT$ = Str$(CVD(Mid$(XXX$, 433, 8)))   ' total neto
                If COP% = 1 Then
                   TRANSPORTE# = TRANSPORTE# + Val(TXT$)
                End If
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
        Call PRINTFORWIN(Str$(FACIVA * (TOTAFAC#(1, 1) + TOTAFAC#(1, 2))), "IMP-NETO", 0, 0)
        DESCUE$ = TRIM$(FORFACNEU.PORDESCU.TEXT)
        Call PRINTFORWIN(DESCUE$, "POR-DESC", 0, 0)
        Call PRINTFORWIN(Str$(FACIVA * TOTAFAC#(1, 4)), "IMP-DESC", 0, 0)
        Call PRINTFORWIN(Str$(FACIVA * (TOTAFAC#(1, 5) + TOTAFAC#(1, 9))), "IMPNEDES", 0, 0)
        Call PRINTFORWIN(Str$(TOTAFAC#(1, 6)), "IMP-IVA", 0, 0)
        Call PRINTFORWIN(Str$(TOTAFAC#(1, 7)), "IMP-RNI", 0, 0)
        Call PRINTFORWIN(Str$(TOTAFAC#(1, 8)), "IMP-TOTA", 0, 0)
        Call PRINTFORWIN(Str$(TOTAFAC#(1, 10)), "IMP-IVA1", 0, 0)
        Call PRINTFORWIN(Str$(VAFLETE#), "IMP-VARS", 0, 0)
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
            IVAITEM# = (TOT# * CVS(Mid$(X$, 441, 4)) + 0.5) / 100
            BRUITEM# = TOT# + IVAITEM#
            '
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
8700 For KK% = 1 To 10: TOTAFAC#(1, KK%) = 0: Next KK%
     TASA1 = Val(FORFACNEU.PORIVA(0).TEXT)
     TASA2 = Val(FORFACNEU.PORIVA(1).TEXT)
8702 For JJ& = 1 To ULTREG&("!CARDEFAC")
        X$ = REGLEIDO$("!CARDEFAC", JJ&)
        TOT# = CVD(Mid$(X$, 433, 8))
        TASARE = CVS(Mid$(X$, 441, 4))
        If Abs(TOT#) >= 0.005 Then
          If TASARE < 0.005 Then
            TASARE = TASA1
            Call REPLA(X$, MKS$(TASARE), 441, 4)
            Call GRAREG("!CARDEFAC", X$, JJ&)
          End If
        End If
        IVAITEM1# = (Int(TOT * TASARE + 0.5)) / 100
        Call REPLA(X$, MKD$(IVAITEM1#), 445, 8)
        Call GRAREG("!CARDEFAC", X$, JJ&)
        If Abs(TASARE - TASA2) < 0.01 Then
            TOTAFAC#(1, 2) = TOTAFAC#(1, 2) + TOT#
            TOTAFAC#(1, 9) = TOTAFAC#(1, 9) + TOT#
            TOTAFAC#(1, 10) = TOTAFAC#(1, 10) + CVD(Mid$(X$, 445, 8))
          Else
            TOTAFAC#(1, 1) = TOTAFAC#(1, 1) + TOT#
            TOTAFAC#(1, 5) = TOTAFAC#(1, 5) + TOT#
            TOTAFAC#(1, 6) = TOTAFAC#(1, 6) + CVD(Mid$(X$, 445, 8))
        End If
        TOTAFAC#(1, 3) = TOTAFAC#(1, 3) + TOT#
     Next JJ&
     '
     If PORDES > 0 Then
       If MODESCU% = 1 Then
         TOTAFAC#(1, 4) = (Int((TOTAFAC#(1, 1) + TOTAFAC#(1, 2)) * PORDES + 0.5)) / 100
         TOTAFAC#(1, 5) = (Int(TOTAFAC#(1, 5) * (100 - PORDES) + 0.5)) / 100
         TOTAFAC#(1, 6) = (Int(TOTAFAC#(1, 6) * (100 - PORDES) + 0.5)) / 100
         TOTAFAC#(1, 9) = (Int(TOTAFAC#(1, 9) * (100 - PORDES) + 0.5)) / 100
         TOTAFAC#(1, 10) = (Int(TOTAFAC#(1, 10) * (100 - PORDES) + 0.5)) / 100
        End If
     End If
     '
     If PIVA% = 2 Then
        TOTAFAC#(1, 7) = (Int(50 * (TOTAFAC#(1, 6) + TOTAFAC#(1, 10)) + 0.5)) / 100
     End If
     '
     If PIVA% = 5 Then
        TOTAFAC#(1, 6) = 0
        TOTAFAC#(1, 10) = 0
        TOTAFAC#(1, 7) = 0
     End If
     '
     TOTATOTA# = 0: TOTAFAC#(1, 8) = 0
     For KK% = 5 To 10
       TOTATOTA# = TOTATOTA# + TOTAFAC#(1, KK%)
     Next KK%
     TOTAFAC#(1, 8) = TOTATOTA#
     '
8800 For I% = 1 To 10
        TOTAFAC#(2, I%) = TOTAFAC#(1, I%) / TICA
     Next I%
     '
     FORFACNEU.Total(0).TEXT = CSTRING$(MKD$(TOTAFAC#(1, 3)), 4, 12, 2, 2)
     FORFACNEU.Total(1).TEXT = CSTRING$(MKD$(TOTAFAC#(1, 4)), 4, 12, 2, 2)
     FORFACNEU.Total(2).TEXT = CSTRING$(MKD$(TOTAFAC#(1, 6)), 4, 12, 2, 2)
     FORFACNEU.Total(3).TEXT = CSTRING$(MKD$(TOTAFAC#(1, 8)), 4, 12, 2, 2)
     FORFACNEU.Total(4).TEXT = CSTRING$(MKD$(TOTAFAC#(1, 10)), 4, 12, 2, 2)
     FORFACNEU.Total(5).TEXT = CSTRING$(MKD$(TOTAFAC#(1, 7)), 4, 12, 2, 2)
     Return

'********************************************************

9000 FORIPRE$ = TRIM$(CONTROL$("", "FOREMIFA"))
     If TRIM$(FORIPRE$) <> "" Then
       '
       PREFIREM$ = CONTROL$("", "PREFIREM")
       If PREFIREM$ = "" Then PREFIREM$ = "0001"
       NROREMI$ = TRIM$(Str$(NUREMIFAC&))
       While Len(NROREMI$) < 8: NROREMI$ = "0" + NROREMI$: Wend
       NROREMI$ = PREFIREM$ + "-" + NROREMI$
       '
       FEX = FFI%
       FECHDOC$ = FECHATEX$(FEX)
       NUMEDOC$ = NROREMI$
       TIPODOC$ = "Remito"
       '
       CODPARAM$(1) = "TRANSPOR": DOCPARAM$(1) = TRIM$(Mid$(TRXP$, 1, 32))
       CODPARAM$(2) = "NUBULTOS": DOCPARAM$(2) = TRIM$(Mid$(TRXP$, 153, 6))
       CODPARAM$(3) = "CUIT-TRA": DOCPARAM$(3) = TRIM$(Mid$(TRXP$, 65, 16))
       CODPARAM$(4) = "VALO-DEC": DOCPARAM$(4) = TRIM$(CSTRING$(MKD$(TOTAFAC#(1, 1)), 3, 12, 2, 2))
       CODPARAM$(5) = "COND-TRA": DOCPARAM$(5) = TRIM$(Mid$(TRXP$, 97, 32))
       CODPARAM$(6) = "DOCUCOND": DOCPARAM$(6) = TRIM$(Mid$(TRXP$, 129, 24))
       CODPARAM$(7) = "DEPOSAL ": 'DOCPARAM$(7) = TRIM$(DEPOSAL$)
       CODPARAM$(8) = "NRO-REMI": DOCPARAM$(8) = TRIM$(NROREMI$)
       CODPARAM$(9) = "NRO-OCOM":  DOCPARAM$(9) = NUOCOMPRA$
       CODPARAM$(14) = "NRO-AENT":  DOCPARAM$(14) = NROPEDIDO$
       CODPARAM$(10) = "PATE-TRA":  DOCPARAM$(10) = TRIM$(Mid$(TRXP$, 81, 16))
       CODPARAM$(11) = "DOMI-TRA":  DOCPARAM$(11) = TRIM$(Mid$(TRXP$, 33, 32))
       CODPARAM$(12) = "LOCA-TRA":  DOCPARAM$(12) = TRIM$(Mid$(TRXP$, 159, 32))
       CODPARAM$(13) = "COND-PAG": DOCPARAM$(13) = DECONPA$(CPAG%)
       CAPARDOC% = 14
       '
       CODTABU1$(1) = "COD-MAT"
       CODTABU1$(2) = "DES-MAT"
       CODTABU1$(3) = "CANTIDAD"
       CACOLTAB1% = 3
       '
       DESTIDOC% = NC%
       CAITAB1% = 0
       '
       For U& = 1 To ULTREG&("!CARDEFAC")
          XXX$ = REGLEIDO$("!CARDEFAC", U&)
          CANS$ = Str$(CVS(Left$(XXX$, 4)))       ' cantidad
          If Val(CANS$) >= 0.05 Then
             If TRIM$(Mid$(XXX$, 387, 16)) = "" Then
                CAITAB1% = CAITAB1% + 1
                CDXXS$ = CODEXT$(CVI(Mid$(XXX$, 5, 2)))     ' codigo
                DDTXT$ = Mid$(XXX$, 7, 64)                  ' descripcion
                '
                TETABU1$(1, CAITAB1%) = CDXXS$
                TETABU1$(2, CAITAB1%) = DDTXT$
                TETABU1$(3, CAITAB1%) = CANS$
                '
             End If
          End If
       Next U&
       '
       CACOPDOC% = COPIASREM%
       Call PRINTDOC("FOREMIFA")   ' Remito de Facturacion
       '
    End If
    '
Return
'
End Sub
'****************************************************************************

Sub FORMPRIN(LU$, FLIM%, ALI1, ALI2, TC1, TLET$, INTER%, TFOR%, ANFO%, LARFO%, DIFER, CAFOR%, MARIZ%, MARINT%, REDON%, MPRI$, MCEN$, FACTO, CANCON%, COMA%, PARDEC%)
      '
      If ULTREG&("FORMFAC") > 0 Then
            FFS$ = REGLEIDO$("FORMFAC", 2&)
            ALI1 = CVS(Mid$(FFS$, 2, 4))
            ALI2 = CVS(Mid$(FFS$, 7, 4))
            TC1 = CVS(Mid$(FFS$, 11, 4))
            '
            TFOR% = Asc(Mid$(FFS$, 6, 1))
            CAFOR% = 2: If TFOR% > 10 Then CAFOR% = (TFOR% Mod 10): TFOR% = Int(TFOR% / 10)
            MPRI$ = TRIM$(Mid$(FFS$, 24, 30))
            MCEN$ = TRIM$(Mid$(FFS$, 54, 30))
            '
            If TFOR% < 1 Or TFOR% > 2 Then TFOR% = 1
         Else
            Call MENSERR(24, "Falta Archivo de Formato de Impresion de Factura")
      End If
      '
End Sub
'


Sub PRIFADOS()

6000 Call CLEARFORMDOS(FORIMPRE$)
     GoSub 6800                            ' LEER ATRIBUTOS DEL FORMULARIO
     '
     PRIDET& = 1: HOJACONT% = 0: TOTITEMS = 0
6005 TRANSPOR% = 0
     For COP% = 1 To CANCOP%
         Call CLEARFORMDOS(FORIMPRE$)     ' GENERAR FORMULARIO EN BLANCO
         GoSub 7100                       ' DATOS GENERALES DE FACTURA
         GoSub 7300: UREDET% = UREN%      ' CARGAR DETALLE DE FACTURACION
         GoSub 7500                       ' CARGAR LEYENDAS
         '
         'If COP% >= PRICOFAC% And COP% <= ULTCOFAC% Then
             If HOJACONT% = 0 Then        ' SOLO EN PRIMERA HOJA
               GoSub 7400                 ' CARGAR TOTALES
             End If
         'End If
         '
         Call PRINTFORMDOS("")            ' IMPRIMIR
    Next COP%
    '
    If TRANSPOR% > 0 Then                 ' TRANSPORTA A HOJA SIGUIENTE
        HOJACONT% = HOJACONT% + 1
        NROCON& = NROCON& + 1
        GoTo 6005
    End If
    '
    If VEPIEDET$ <> "" Then               ' LEYENDAS COMPLEMENTARIAS
        NROCON& = NROCON& + 1
        For COP% = 1 To CANCOP%
           Call CLEARFORMDOS(FORIMPRE$)       ' GENERAR FORMULARIO EN BLANCO
           GoSub 7100                      ' DATOS GENERALES DE FACTURA
           '
           RG0% = PRIREDET%: UREN% = RG0% - 1
           For U& = 1 To ULTREG&("!FACEXPO2")
              XXX$ = REGLEIDO$("!FACEXPO2", U&)
              RGII% = RG0%
              TXT$ = XXX$
              Call APPFORMDOS("", RGII%, TXT$, "ESPE-DET")
              RG0% = RGII% + 1
           Next U&
           '
           Call PRINTFORMDOS("")              ' IMPRIMIR
           '
        Next COP%
    End If
    '
    If ALTERNA%("-Repite Impresion") = 1 Then
        NROCON& = NRO                   ' EMPIEZA DESDE EL PRIMERO
        GoTo 6000
    End If
    '
    Call CIERRARCH("FORMUFAC")
    '
    Exit Sub
    '
'****************************************************************************
    '
    '
6800 TILETRA$ = ATRIFORMDOS$("", "TI-LETRA")
     LINPULG% = Val(ATRIFORMDOS$("", "LIN-PULG"))
     CACOL% = Val(ATRIFORMDOS$("", "CAN-COL"))
     CAREN% = Val(ATRIFORMDOS$("", "CAN-REN"))
     MARSUP% = Val(ATRIFORMDOS$("", "MARG-SUP"))
     MARIZQ% = Val(ATRIFORMDOS$("", "MARG-IZQ"))
     CANCOP% = Val(ATRIFORMDOS$("", "CAN-COP"))
     PRICOFAC% = Val(ATRIFORMDOS$("", "IMP-DES"))
     ULTCOFAC% = Val(ATRIFORMDOS$("", "IMP-HAS"))
     PRIREDET% = Val(ATRIFORMDOS$("", "DET-DES"))
     ULTREDET% = Val(ATRIFORMDOS$("", "DET-HAS"))
     MONPRI$ = ATRIFORMDOS$("", "MON-PRI")
     MONSEC$ = ATRIFORMDOS$("", "MON-SEC")
     PREFI$ = ATRIFORMDOS$("", "PREFIFAC")
     '
     Return
     '
'****************************************************************************
     '
7100 FEX = FECHANUM(TRIM$(FORFACNEU.Text9.TEXT))
     CPAG% = Val(FORFACNEU.TEXT21(1).TEXT)
     
     Call SETFORMDOS("", TIPCOM$, "TIPO-COM")
     Call SETFORMDOS("", FECHATEX$(FEX), "FECH-EMI")
     Call SETFORMDOS("", RASOCLI$(NC%), "RASO-CLI")
     Call SETFORMDOS("", CUITCLI$(NC%), "CUIT-CLI")
     Call SETFORMDOS("", DOMICLI$(NC%), "DOMI-CLI")
     Call SETFORMDOS("", NIBRCLI$(NC%), "IBRU-CLI")
     Call SETFORMDOS("", CPOSCLI$(NC%), "CPOS-CLI")
     Call SETFORMDOS("", LOCACLI$(NC%), "LOCA-CLI")
     Call SETFORMDOS("", NPROCLI$(NC%), "PROV-CLI")
     Call SETFORMDOS("", Str$(NC%), "CUEN-CLI")
     Call SETFORMDOS("", POSIVAL$(NC%), "PIVA-CLI")
     Call SETFORMDOS("", DECONPA$(CPAG%), "COND-PAG")
     '
     Call SETFORMDOS("", RTO$, "NRO-REMI")
     '
     Call SETFORMDOS("", Mid$(TRXP$, 1, 32), "TRANSPOR")
     Call SETFORMDOS("", Mid$(TRXP$, 33, 32), "DOMI-TRA")
     Call SETFORMDOS("", Mid$(TRXP$, 65, 16), "CUIT-TRA")
     Call SETFORMDOS("", Mid$(TRXP$, 81, 16), "PATE-TRA")
     Call SETFORMDOS("", Mid$(TRXP$, 97, 32), "COND-TRA")
     '
     NCN$ = TRIM$(Str$(NROCON&))
     While Len(NCN$) < 8: NCN$ = "0" + NCN$: Wend
     PREFI$ = ATRIFORMDOS$("", "PREFIFAC")
     NCN$ = PREFI$ + "-" + NCN$
     Call SETFORMDOS("", NCN$, "NUME-COM")
     '
     TTTT# = TOTAFAC#(1, 8)
     TOTFAC$ = NUMTEX$(TTTT#)
     Call SETFORMDOS("", TOTFAC$, "IMP-LET1")
     '
7190 Return
     '
'****************************************************************************
     '
7300 RG0% = PRIREDET%: UREN% = RG0% - 1
     '
     If PRIDET& = 1 Then
       If VENCADET$ <> "" Then
          For U& = 1 To ULTREG&("!FACEXPO1")
          XXX$ = REGLEIDO$("!FACEXPO1", U&)
          RGII% = RG0%
          TXT$ = XXX$
          Call APPFORMDOS("", RGII%, TXT$, "ENCA-DET")
          RG0% = RGII% + 1
          Next U&
          RG0% = RG0% + 1
       End If
     End If
     '
     For U& = PRIDET& To ULTREG&("!CARDEFAC")
        '
        If RG0% > ULTREDET% Then                          ' SALTO DE HOJA
           PRIDET& = U&
           TRANSPOR% = 1
           GoTo 7390
        End If
        '
        XXX$ = REGLEIDO$("!CARDEFAC", U&)
        RGII% = RG0%
        '
        TOTITEMS = TOTITEMS + CVS(Left$(XXX$, 4))
        TXT$ = Str$(CVS(Left$(XXX$, 4)))       ' cantidad
        Call APPFORMDOS("", RGII%, TXT$, "CANTIDAD")
        If RGII% > UREN% Then UREN% = RGII%
        RGII% = RG0%
        '
        TXT$ = CODEXT$(CVI(Mid$(XXX$, 5, 2)))     ' codigo
        Call APPFORMDOS("", RGII%, TXT$, "COD-MAT")
        If RGII% > UREN% Then UREN% = RGII%
        RGII% = RG0%
        '
        TXT$ = Mid$(XXX$, 7, 64)                  ' descripcion
        Call APPFORMDOS("", RGII%, TXT$, "DES-MAT")
        If RGII% > UREN% Then UREN% = RGII%
        RGII% = RG0%
        '
        TXT$ = Mid$(XXX$, 71, 4)                  ' unidad de medida
        Call APPFORMDOS("", RGII%, TXT$, "UNIMED")
        If RGII% > UREN% Then UREN% = RGII%
        RGII% = RG0%
        '
        TXT$ = Mid$(XXX$, 75, 64)                 ' referencia 1
        Call APPFORMDOS("", RGII%, TXT$, "REF-MAT1")
        If RGII% > UREN% Then UREN% = RGII%
        RGII% = RG0%
        '
        TXT$ = Mid$(XXX$, 139, 64)                ' referencia 2
        Call APPFORMDOS("", RGII%, TXT$, "REF-MAT2")
        If RGII% > UREN% Then UREN% = RGII%
        RGII% = RG0%
        '
        TXT$ = Mid$(XXX$, 203, 64)                ' referencia 3
        Call APPFORMDOS("", RGII%, TXT$, "REF-MAT3")
        If RGII% > UREN% Then UREN% = RGII%
        RGII% = RG0%
        '
        TXT$ = Mid$(XXX$, 315, 24)                ' numero de plano
        Call APPFORMDOS("", RGII%, TXT$, "N-PLANO")
        If RGII% > UREN% Then UREN% = RGII%
        RGII% = RG0%
        '
        TXT$ = Mid$(XXX$, 339, 16)                ' numero original
        Call APPFORMDOS("", RGII%, TXT$, "N-ORIGIN")
        If RGII% > UREN% Then UREN% = RGII%
        RGII% = RG0%
        '
        TXT$ = Mid$(XXX$, 355, 16)                ' numero de o.compra
        Call APPFORMDOS("", RGII%, TXT$, "N-OCOMPR")
        If RGII% > UREN% Then UREN% = RGII%
        RGII% = RG0%
        '
        TXT$ = Mid$(XXX$, 387, 16)                ' numero de remito
        Call APPFORMDOS("", RGII%, TXT$, "N-REMITO")
        If RGII% > UREN% Then UREN% = RGII%
        RGII% = RG0%
        '
        TXT$ = Mid$(XXX$, 403, 16)                ' numero de doc-recepcion
        Call APPFORMDOS("", RGII%, TXT$, "N-BOLREC")
        If RGII% > UREN% Then UREN% = RGII%
        RGII% = RG0%
        '
        FEX = CVI(Mid$(XXX$, 419, 2))             ' fecha de recepcion
        TXT$ = FECHATEX$(FEX)
        Call APPFORMDOS("", RGII%, TXT$, "FECH-REC")
        If RGII% > UREN% Then UREN% = RGII%
        RGII% = RG0%
        '
      If COP% >= PRICOFAC% And COP% <= ULTCOFAC% Then
        TXT$ = Str$(FACIVA * CVS(Mid$(XXX$, 421, 4))) ' precio de lista
        Call APPFORMDOS("", RGII%, TXT$, "PRE-LIST")
        If RGII% > UREN% Then UREN% = RGII%
        RGII% = RG0%
        '
        TXT$ = Str$(FACIVA * CVS(Mid$(XXX$, 429, 4))) ' precio unitario
        Call APPFORMDOS("", RGII%, TXT$, "PRE-NETO")
        If RGII% > UREN% Then UREN% = RGII%
        RGII% = RG0%
        '
        TXT$ = Str$(CVD(Mid$(XXX$, 433, 8)))   ' total neto
        Call APPFORMDOS("", RGII%, TXT$, "NET-ITEM")
        If RGII% > UREN% Then UREN% = RGII%
        RGII% = RG0%
        '
        TXT$ = Str$(CVS(Mid$(XXX$, 441, 4)))   ' alicuota iva
        Call APPFORMDOS("", RGII%, TXT$, "ALICUIVA")
        If RGII% > UREN% Then UREN% = RGII%
        RGII% = RG0%
        '
        TXT$ = Str$(CVD(Mid$(XXX$, 445, 8)))   ' importe iva del item
        Call APPFORMDOS("", RGII%, TXT$, "IVA-ITEM")
        If RGII% > UREN% Then UREN% = RGII%
        RGII% = RG0%
        '
        TXT$ = Str$(CVD(Mid$(XXX$, 453, 8)))   ' total bruto
        Call APPFORMDOS("", RGII%, TXT$, "BRU-ITEM")
        If RGII% > UREN% Then UREN% = RGII%
        RGII% = RG0%
        '
      End If
      '
      RG0% = UREN% + 1
     Next U&
     '
     RGII% = RG0%
     Call APPFORMDOS("", RGII%, String$(32, "="), "LIN-CIE")
       If RGII% > UREN% Then UREN% = RGII%
     RGII% = RG0%
     Call APPFORMDOS("", RGII%, Str$(TOTITEMS), "TOTA-ITS")
       If RGII% > UREN% Then UREN% = RGII%
     RGII% = RG0%
     Call APPFORMDOS("", RGII%, "Total de Items", "LETOT-IT")
     If RGII% > UREN% Then UREN% = RGII%
     '
7390 Return
     '
'****************************************************************************
    '
7400 If COP% >= PRICOFAC% And COP% <= ULTCOFAC% Then
         Call SETFORMDOS("", Str$(FACIVA * TOTAFAC#(1, 1)), "IMP-NETO")
         Call SETFORMDOS("", Str$(PORDES), "POR-DESC")
         Call SETFORMDOS("", Str$(FACIVA * TOTAFAC#(1, 4)), "IMP-DESC")
         Call SETFORMDOS("", Str$(FACIVA * TOTAFAC#(1, 5)), "IMPNEDES")
         Call SETFORMDOS("", Str$(TOTAFAC#(1, 6)), "IMP-IVA")
         Call SETFORMDOS("", Str$(TOTAFAC#(1, 7)), "IMP-RNI")
         Call SETFORMDOS("", Str$(TOTAFAC#(1, 8)), "IMP-TOTA")
         Call SETFORMDOS("", Str$(TOTAFAC#(1, 9)), "IMP-VARS")
         Call SETFORMDOS("", Str$(TOTAFAC#(1, 10)), "PERC-IVA")
     End If
     '
     Return
     '
'****************************************************************************
    '
7500 Call CARLEYENDOS("")
     '
     Return
     '
End Sub


Sub FORMMGRDOS(Form$, RG%, CL%, VLX$, CMD$)
   '
   ' FORM$  =  Nombre de formulario
   ' RG%    =  Renglon
   ' CL%    =  Columna
   ' VLX$   =  Valor o texto a colocar
   ' CMD$   =  Comando a ejecutar
   '
   Static FORMACT$, NX%, NY%, ATOT%, CARE%
   Static FXX As String * 128
   Static FYY As String * 256
   '
   If Form$ <> "" Then
      If UCase$(TRIM$(Form$)) <> FORMACT$ Then
         FORMACT$ = UCase$(TRIM$(Form$))
         FORMUDRV$ = LUDAT$ + FORMACT$ + ".FRM"
         If NX% > 0 Then Close #NX%
         NX% = FILEOPEN%(FORMUDRV$, 0, 128)
         '
         ATOT% = 256: CARE% = 128       ' ancho total y cantidad de renglones
         For K% = 1 To LOF(NX%) / 128
            Get #NX%, K%, FXX$
            If Left$(FXX$, 8) = "CAN-COL " Then
                  ATOT% = Val(TRIM$(Mid$(FXX$, 89, 8)))
                ElseIf Left$(FXX$, 8) = "CAN-REN " Then
                  CARE% = Val(TRIM$(Mid$(FXX$, 89, 8)))
            End If
         Next K%
      End If
   End If
   '
   COMA$ = UCase$(TRIM$(CMD$))
   '
   If Left$(COMA$, 5) = "CLEAR" Then           ' genera formulario en blanco
      LUX$ = LUCOM$
      If NY% > 0 Then Close #NY%
      NY% = FILEOPEN%(LUX$ + "\" + FORMACT$, 0, 256)
      '
      For I% = 1 To CARE%
         LSet FYY$ = Space$(256)
         Put #NY%, I%, FYY$
      Next I%
      '
      Exit Sub
      '
   ElseIf Left$(COMA$, 5) = "PRINT" Then
      Screen.MousePointer = 11
      UTI$ = LUBAS$
      If Mid$(UTI$, 2, 1) <> ":" Then
          UTI$ = Left$(CurDir, 2) + UTI$
      End If
      LUX$ = LUCOM$
      ATRIBU$ = "/" + LUX$ + "\" + FORMACT$
      ATRIBU$ = ATRIBU$ + "/CAREN=" + TRIM$(Str$(CARE%))
      ATRIBU$ = ATRIBU$ + "/SETUP=" + SETUPSTRI$
      ATRIBU$ = ATRIBU$ + "/MARSU=" + TRIM$(Str$(MARSUP%))
      '
      NZ% = FILEOPEN%("C:\FLEXOFT\COMBAT.$$$", 2, 0)
      Print #NZ%, "FINBAT"
      Close #NZ%
      '
      On Error Resume Next
      Kill "C:\FLEXOFT\FINBAT.$$$"
      On Error GoTo 0
      '
      NZ% = FILEOPEN%("C:\FLEXOFT\DOSCOMM.BAT", 2, 0)
      Print #NZ%, UTI$ + "PRINTFRM.EXE " + ATRIBU$
      Close #NZ%
      '
      Shell "C:\FLEXOFT\DOSCOMM.PIF"
      TINI = Timer
      MENSA$ = "Se ha producido un error no recuperable" + Chr$(10) + Chr$(13)
      MENSA$ = MENSA$ + "en la impresion de este documento." + Chr$(10) + Chr$(13) + Chr$(10) + Chr$(13)
      MENSA$ = MENSA$ + "Consulte a su Soporte de Sistemas FLEXOFT"
      '
149   If Timer > TINI + 30 Then
          Screen.MousePointer = 1
          MsgBox MENSA$, 4096 + 48, "Control de Impresion"
          On Error Resume Next
          Kill "C:\FLEXOFT\COMBAT.$$$"
          Kill "C:\FLEXOFT\FINBAT.$$$"
          On Error GoTo 0
          ABOPRIN% = 1
          Exit Sub
      End If
      '
      If EXISTE%("C:\FLEXOFT\FINBAT.$$$") < 1 Then GoTo 149
      Screen.MousePointer = 1
      ABOPRIN% = 0
      Exit Sub
      '
   ElseIf Left$(COMA$, 4) = "FSET" Then
      If RG% > 0 Then
        If RG% <= CARE% Then
          If CL% > 0 Then
            If CL% <= ATOT% Then
              Get #NY%, RG%, FYY$
              TXTT$ = FYY$
              Mid$(TXTT$, CL%, Len(VLX$)) = VLX$
              LSet FYY$ = TXTT$
              Put #NY%, RG%, FYY$
            End If
          End If
        End If
      End If
      Exit Sub
      '
   ElseIf Left$(COMA$, 5) = "LEYEN" Then
      For K% = 1 To LOF(NX%) / 128
        Get #NX%, K%, FXX$
        If Left$(FXX$, 4) = "LEYE" Then
          VLX$ = TRIM$(Mid$(FXX$, 9, 80))
          If VLX$ <> "" Then
            For TTI% = 97 To 128 Step 8
              RG% = Asc(Mid$(FXX$, TTI%, 1))
              If RG% > 0 Then
                If RG% <= CARE% Then
                  CL% = Asc(Mid$(FXX$, TTI% + 1, 1))
                  If CL% > 0 Then
                    If CL% <= ATOT% Then
                      FORMATO$ = TRIM$(Mid$(FXX$, TTI% + 2, 6))
                      If FORMATO$ <> "" Then
                        TX$ = TEXTFORM$(VLX$, FORMATO$)
                        Get #NY%, RG%, FYY$
                        TXTT$ = FYY$
                        Mid$(TXTT$, CL%, Len(TX$)) = TX$
                        LSet FYY$ = TXTT$
                        Put #NY%, RG%, FYY$
                      End If
                    End If
                  End If
                End If
              End If
            Next TTI%
          End If
        End If
      Next K%
      '
      Exit Sub
      '
   ElseIf Left$(COMA$, 1) = "@" Then
      CD$ = AJUSTI$(Mid$(COMA$, 2), 8)
      For K% = 1 To LOF(NX%) / 128
        Get #NX%, K%, FXX$
        If Left$(FXX$, 8) = CD$ Then
          RG0% = RG%
          UREGUT% = 0
          For TTI% = 97 To 128 Step 8
            RG% = RG0%
            RGX% = Asc(Mid$(FXX$, TTI%, 1))
            If RG% + RGX% > 0 Then
              If RG% + RGX% <= CARE% Then
                CL% = Asc(Mid$(FXX$, TTI% + 1, 1))
                If CL% > 0 Then
                  If CL% <= ATOT% Then
                    FORMATO$ = TRIM$(Mid$(FXX$, TTI% + 2, 6))
                    If FORMATO$ <> "" Then
                      ILETRA% = 0
                      '
                      If CD$ = "IMP-LET1" Then
                        ILETRA% = 1
                        LOTEX% = Val(Mid$(FORMATO$, 2, 4))
                        Call FRATEX(VLX$, NRX1$, NRX2$, NRX3$, LOTEX%)
                        VLX$ = NRX1$ + String$(LOTEX%, "*")
                        RL% = 1
                      End If
                      '
1000                  TX$ = TEXTFORM$(VLX$, FORMATO$)
                      RG% = RG% + RGX%
                      Get #NY%, RG%, FYY$
                      TXTT$ = FYY$
                      Mid$(TXTT$, CL%, Len(TX$)) = TX$
                      LSet FYY$ = TXTT$
                      Put #NY%, RG%, FYY$
                      If RG% > UREGUT% Then UREGUT% = RG%
                      '
                      If ILETRA% = 1 Then
                        If RL% = 1 Then
                          RL% = 2
                          VLX$ = NRX2$ + String$(LOTEX%, "*")
                          RGX% = 1
                          GoTo 1000
                        ElseIf RL% = 2 Then
                          RL% = 3
                          VLX$ = NRX3$ + String$(LOTEX%, "*")
                          RGX% = 1
                          GoTo 1000
                        End If
                      End If
                    End If
                  End If
                End If
              End If
            End If
          Next TTI%
          If UREGUT% > RG% Then RG% = UREGUT%
          Exit Sub
        End If
      Next K%
      Exit Sub
      '
   ElseIf Left$(COMA$, 1) = "?" Then
      CD$ = AJUSTI$(Mid$(COMA$, 2), 8)
      If NX% > 0 Then
        For K% = 1 To LOF(NX%) / 128
          Get #NX%, K%, FXX$
          If Left$(FXX$, 8) = CD$ Then
            VLX$ = TRIM$(Mid$(FXX$, 89, 8))
            Exit Sub
          End If
        Next K%
      End If
      VLX$ = ""
      Exit Sub
      '
   Else
      '
      Call COMUNI("Comando '" + COMA$ + "' no Procesable\en Rutina 'FORMMGR'\para Formulario '" + FORMACT$ + "'.")
      '
   End If
   '
End Sub
   '
'*****************************************************************************
   '
Sub SETFORMDOS(Form$, VLX$, CTRX$)
   '
   CMD$ = "@" + TRIM$(CTRX$)
   RGII% = 0: CLII% = 0
   Call FORMMGRDOS(Form$, RGII%, CLII%, VLX$, CMD$)
   ' sobre RG% y CL% devuelve renglon y columna utilizado
   '
End Sub
'
'*****************************************************************************
'
Sub CARLEYENDOS(Form$)
   '
   CMD$ = "LEYEN"
   RGII% = 0: CLII% = 0
   Call FORMMGRDOS(Form$, RGII%, CLII%, VLX$, CMD$)
   '
End Sub
'
'*****************************************************************************
'
Sub APPFORMDOS(Form$, RGII%, VLX$, CTRX$)
   '
   CMD$ = "@" + TRIM$(CTRX$)
   If TRIM$(VLX$) <> "" Then
       Call FORMMGRDOS(Form$, RGII%, CLII%, VLX$, CMD$) ' sobre RG% y CL% devuelve
   End If                                    '    renglon y columna utilizado
   '
End Sub
'
'*****************************************************************************
'
Sub CLEARFORMDOS(Form$)
   '
   Call FORMMGRDOS(Form$, RGII%, CLII%, VLX$, "CLEAR")
   '
End Sub
'
'*****************************************************************************
'
Sub PRINTFORMDOS(Form$)
   '
   Call FORMMGRDOS(Form$, RGII%, CLII%, VLX$, "PRINT")
   '
End Sub
'
'*****************************************************************************
'
Function ATRIFORMDOS$(Form$, CTRX$)
   '
   CMD$ = "?" + TRIM$(CTRX$)
   Call FORMMGRDOS(Form$, RGII%, CLII%, VLX$, CMD$)
   ATRIFORMDOS$ = VLX$
   '
End Function
