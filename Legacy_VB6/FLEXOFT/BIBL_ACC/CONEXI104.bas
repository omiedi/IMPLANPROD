Attribute VB_Name = "CONEXI1"
Public ABUS$
Public YALIS%
Public ILISTA&
'
Public PREPA%
Public FORMAJU%
'
Dim REFE$
Dim FF%
Dim CIII%
Public MODENTRAJU%
'
Public MODOMAI%
Public APEDPE$
Public PUNPEDPE$(32767)
'

Sub MUEXIS(CI%)
    '
    Screen.MousePointer = 11
    '
    Call CARDEPOS
    '
    CONEXI.List1.Clear
    STOTOT = 0
    '
    Call VESARRAS(CI%)
    '
    REG& = CI% + 1
    If REG& > 0 Then
        If REG& <= NUMAS% + 1 Then
           SALCA$ = REGLEIDO$("SALSTOCK", REG&) + Mid$(REGLEIDO$("SALDEPOS", REG&), 41)
           For U& = 1 To NUSU%
               XXZ$ = Mid$(SALCA$, 4 * U& - 3, 4)
               XXX$ = SUCU$(U&) + CSTRING$(XXZ$, 4, 12, 1, 2)
               CONEXI.List1.AddItem XXX$
               STOTOT = STOTOT + CVS(XXZ$)
           Next U&
           CONEXI.Label4.Caption = TRIM$(CSTRING$(MKS$(STOTOT), 4, 12, 1, 2))
       End If
    End If
    '
10  CONEXI.List2.Clear
    HOII% = HOY(HO$)
    LI2TEX$ = Space$(60)
    Call REPLA(LI2TEX$, "Stock al " + HO$, 7, 20)
    Call REPLA(LI2TEX$, CSTRING$(MKS$(STOTOT), 4, 8, 1, 2), 48, 8)
    CONEXI.List2.AddItem LI2TEX$
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
            NroPed& = CVS(Left$(X$, 4))
            NCI% = CVI(Mid$(X$, 7, 2))
            CASIG = 0
            If Mid$(X$, 75, 4) <> Space$(4) Then
                CASIG = CVS(Mid$(X$, 75, 4))
            End If
            CAPEN = CVS(Mid$(X$, 79, 4)) - CVS(Mid$(X$, 95, 4))
            If CAPEN < 0 Then CAPEN = 0
            '
            'If ASII + CASIG > STOTOT Then
            '    CASIG = STOTOT - ASII
            'End If
            '
            If CAPEN > 0 Or CASIG > 0 Then
                PEDI = PEDI + CAPEN
                ASII = ASII + CASIG
                If ASII > PEDI Then ASII = PEDI
            End If
            '
            LITEX2$ = Space$(60)
            Call REPLA(LI2TEX$, CSTRING$(MKS$(NroPed&), 3, 5, 0, 2), 1, 5)
            Call REPLA(LI2TEX$, RASOCLI$(NCI%), 7, 30)
            Call REPLA(LI2TEX$, CSTRING$(MKS$(CASIG), 4, 8, 1, 2), 32, 8)
            Call REPLA(LI2TEX$, CSTRING$(MKS$(CAPEN - CASIG), 4, 8, 1, 2), 40, 8)
            SALSTO = SALSTO - CAPEN
            Call REPLA(LI2TEX$, CSTRING$(MKS$(SALSTO), 4, 8, 1, 2), 48, 8)
            CONEXI.List2.AddItem LI2TEX$
            '
        End If
    Next I&
    If ASII > STOTOT Then
        Call COMUNI("ATENCION:\Las Cantidades Asignadas\Superan el Stock Actual.\  \Revise las Asignaciones para este Item.")
     '   ASII = STOTOT
    End If
    '
    CONEXI.LABEL5.Caption = CSTRING$(MKS$(PEDI), 4, 12, 2, 2)
    CONEXI.Label7.Caption = CSTRING$(MKS$(ASII), 4, 12, 2, 2)
    CONEXI.Label14.Caption = CSTRING$(MKS$(STOTOT - PEDI), 4, 12, 2, 2)
    CONEXI.Label10.Caption = CSTRING$(MKS$(STOTOT - ASII), 4, 12, 2, 2)
    '
20  CONEXI.LIUBIC.Clear
    RGXX$ = RECFILT$("UBISTOCK", 1, MKI$(CI%))
    For U& = 1 To Len(RGXX$) Step 4
        REG& = CVS(Mid$(RGXX$, U&, 4))
        X$ = REGLEIDO$("UBISTOCK", REG&)
        If CVI(Left$(X$, 2)) <> CI% Then
           X$ = REGLEIDO$("UBISTOCK", 1)
           Call GRAREG("UBISTOCK", X$, 1)
           Call CIERRARCH("UBISTOCK")
           GoTo 20
        End If
        '
        If TRIM$(Mid$(X$, 3, 24)) <> "" Then
           Y$ = Mid$(X$, 3, 24)
           CONEXI.LIUBIC.AddItem Y$
        End If
    Next U&
    '
    Screen.MousePointer = 1
     '
End Sub
'

Sub CARPEDPE()
    '
    FIN& = NUMAS%
    For U& = 1 To FIN&
       If PUNPEDPE$(U&) <> "" Then
          Call TRACE(20, U&, FIN&)
          PUNPEDPE$(U&) = ""
       End If
    Next U&
    '
    APEDPE$ = "PEDDETP": MODOMAI% = 0
    If CONTROL$("ADMIPED", "FORMACAR") = "ITEMIZADO" Then
       MODOMAI% = 1: APEDPE$ = "PEDCLIEN"
       Dim BUFI As String * 1024
       NX% = FILEOPEN%(LUDAT$ + "PEDCLIEN.DAT", 0, 1024)
    End If
    '
    UINI& = XVALO(CONTROL$("ADMIPED", "PRIPEDPE"))
    If UINI& < 1 Then UINI& = 1
    PRIREPEN& = 0: FIN& = ULTREG&(APEDPE$)
    For U& = UINI& To FIN&
       Call TRACE(20, U&, FIN&)
       If MODOMAI% = 1 Then
           Get #NX%, U&, BUFI$
           X$ = BUFI$
         Else
           X$ = REGLEIDO$(APEDPE$, U&)
       End If
       CAPEN = 0
       If MODOMAI% <> 1 Then
            CIIX% = CVI(Mid$(X$, 33, 2))
            CAPEN = CVS(Mid$(X$, 79, 4)) - CVS(Mid$(X$, 95, 4))
          Else
            CIIX% = CVI(Mid$(X$, 640 + 3, 2))
            Status% = Asc(Mid$(X$, 300, 1))
            If Status% <= 1 Then
               CACOPED = CVS(Mid$(X$, 513, 4))
               CAITPED = CVS(Mid$(X$, 645, 4))
               CACOENT = CVS(Mid$(X$, 512 + 77, 4))
               CAPEN = (CACOPED - CACOENT) * CAITPED
            End If
       End If
       '
       If CAPEN > 0 Then
         If CIIX% > 0 Then
           If CIIX% <= NUMAS% Then
             PUNPEDPE$(CIIX%) = PUNPEDPE$(CIIX%) + MKS$(U&)
             If PRIREPEN& = 0 Then PRIREPEN& = U&
           End If
        End If
      End If
    Next U&
    Call GRACONTROL("ADMIPED", "PRIPEDPE", Str$(PRIREPEN&))
    If MODOMAI% = 1 Then Close #NX%
    '
End Sub
'
Sub RECDEF()
    '
    HO% = HOY(HOS$)
    LU$ = LUDAT$
    '
    TITU$ = "PROCESO DE APROBACION DE RECEPCIONES"
    '
10  Call FRESHECHO("BOLREPE")
    If ULTREG&("BOLREPE") < 1 Then
        Call MENSERR(24, "No hay Recepciones\Pendientes de Aprobación")
        Exit Sub
    End If
    '
    Call SELECHO("BOLREPE/Recepciones para Aprobación")
    '
20  BOLRE$ = VALACT1$("BOLREPE")
    If TRIM$(BOLRE$) = "" Then
        Exit Sub
    End If
    '
    REGG& = REGNUM("BOLREPE")
    XX$ = REGLEIDO$("BOLREPE", REGG&)
    ARCHORI$ = Mid$(XX$, 85, 8)
    REGI& = CVS(Mid$(XX$, 93, 4))
    If REGI& > 0 Then
        If REGI& <= ULTREG&(ARCHORI$) Then
            YY$ = REGLEIDO$(ARCHORI$, REGI&)
            LORI% = RECLEN%(ARCHORI$)
            If Mid$(XX$, 97, LORI%) = YY$ Then GoTo 30
        End If
    End If
    '
    Call MENSERR(24, "Se ha Detectado una Inconsistencia\en la Base de Datos de Recepciones\Pendientes.\  \Consulte a su Soporte de Sistemas.")
    Exit Sub
    '
30  VDEF$ = REGLEIDO$(ARCHORI$, REGI&)
    CANOMI# = CVD(Mid$(VDEF$, 58, 8))
    ACEPANT# = CVD(Mid$(VDEF$, 66, 8))
    RECHANT# = CVD(Mid$(VDEF$, 74, 8))
    FALTANT# = CVD(Mid$(VDEF$, 82, 8))
    '
    CACEP# = CANOMI# - ACEPANT# - RECHANT# - FALTANT#
    CARECH# = 0
    CAFALT# = 0
    Call REPLA(VDEF$, MKD$(CACEP#), 66, 8)
    Call REPLA(VDEF$, MKD$(CARECH#), 74, 8)
    Call REPLA(VDEF$, MKD$(CAFALT#), 82, 8)
    '
    If TRIM$(Mid$(VDEF$, 11, 12)) = "" Then
        Call REPLA(VDEF$, BOLRE$, 11, 12)
    End If
    '
    VNUE$ = OBJPLA$("RECDEF", VDEF$)
    If VNUE$ = "" Then GoTo 10
    '
    Screen.MousePointer = 11
    '
    FECHA% = CVI(Left$(VNUE$, 2))
    CI% = CVI(Mid$(VNUE$, 9, 2))
    LOTE$ = Mid$(VNUE$, 11, 12)
    CMOV$ = "BR"
    DOPRI$ = "BR-" + TRIM$(Str$(CVS(Mid$(VNUE$, 3, 4))))
    DOSEC$ = Mid$(VNUE$, 28, 10)
    REFE$ = ECOARCH$("PROVED1", Mid$(VNUE$, 7, 2))
    VUNI# = 0
    MONE$ = ""
    NC% = (-1) * CVI(Mid$(VNUE$, 7, 2))
    DEPO% = Asc(Mid$(VNUE$, 27, 1))
    '
    ACEP = CVD(Mid$(VNUE$, 66, 8))
    RECH = CVD(Mid$(VNUE$, 74, 8))
    FALT = CVD(Mid$(VNUE$, 82, 8))
    TOIN = ACEP + RECH + FALT
    '
    If TOIN > 0 Then
        'Call MOVISTOK(FECHA%, CI%, LOTE$, CMOV$, DOPRI$, DOSEC$, REFE$, VUNI#, MONE$, NC%, DEPO%, TOIN)
    End If
    '
    If RECH > 0 Then
        REFE$ = "Rechazo en Recepcion"
        'Call MOVISTOK(FECHA%, CI%, LOTE$, CMOV$, DOPRI$, DOSEC$, REFE$, VUNI#, MONE$, NC%, DEPO%, (-1) * RECH)
    End If
    '
    If Abs(FALT) > 0 Then
        REFE$ = "Diferencia en Recepcion"
        'Call MOVISTOK(FECHA%, CI%, LOTE$, CMOV$, DOPRI$, DOSEC$, REFE$, VUNI#, MONE$, NC%, DEPO%, (-1) * FALT)
    End If
    '
    VDEF$ = REGLEIDO$(ARCHORI$, REGI&)
    CANOMI# = CVD(Mid$(VDEF$, 58, 8))
    ACEPANT# = CVD(Mid$(VDEF$, 66, 8)) + ACEP
    RECHANT# = CVD(Mid$(VDEF$, 74, 8)) + RECH
    FALTANT# = CVD(Mid$(VDEF$, 82, 8)) + FALT
    Call REPLA(VDEF$, MKD$(ACEPANT#), 66, 8)
    Call REPLA(VDEF$, MKD$(RECHANT#), 74, 8)
    Call REPLA(VDEF$, MKD$(FALTANT#), 82, 8)
    Call GRAREG(ARCHORI$, VDEF$, REGI&)
    '
    XX$ = REGLEIDO$("BOLREPE", REGG&)
    Call REPLA(XX$, VDEF$, 97, LORI%)
    Call GRAREG("BOLREPE", XX$, REGG&)
    '
    If CANOMI# - ACEPANT# - RECHANT# - FALTANT# < 0.05 Then
        FIN& = ULTREG&("BOLREPE")
        For U& = REGG& + 1 To FIN&
            XX$ = REGLEIDO$("BOLREPE", U&)
            Call GRAREG("BOLREPE", XX$, U& - 1)
        Next U&
        XX$ = String$(512, 0)
        Call GRAREG("BOLREPE", XX$, FIN&)
    End If
    '
    Call CIERRARCH("*.*")
    Screen.MousePointer = 1
    GoTo 10
    '
End Sub
'
Sub RECPROP()
    '
    If PRECOCIE% = 0 Then
       PRECOCIE% = (-1)
       If CONTROL$("", "PRECOCIE") = "SI" Then
          PRECOCIE% = 1
       End If
    End If
    '
10  HOOI% = HOY(HO$)
    If ULTREG&("ORDEFABR") < 1 Then   ' por numero de O/F
        GoTo 100
    End If
    '
    MODO% = COMALTER%("Puede Declarar Entradas a Stock:\  \a) por Cierre Parcial o Total de Orden de Fabricación\o\b) en Forma Directa por Armado de Conjunto\\Cierre de O./Fabricación\Armado de Conjunto\Cancelar")
    On MODO% GoTo 20, 100
    Exit Sub
20  MsgBox "falta orfcie"
    'call ORFCIE
    Exit Sub
    '
100 Screen.MousePointer = 11
    For U& = ULTREG&("RECPROP") To 1 Step -1
        XX$ = REGLEIDO$("RECPROP", U&)
        If CVS(Mid$(XX$, 59, 4)) > 0 Then
            UPARTE& = CVS(Mid$(XX$, 59, 4))
            GoTo 105
        End If
    Next U&
    UPARTE& = 0
    '
105 Screen.MousePointer = 1
    Call BLANARCH("!CAREPRO")
    '
    VDEF$ = String$(88, 0)
    Call REPLA(VDEF$, MKI$(HOOI%), 1, 2)
    Call REPLA(VDEF$, MKS$(UPARTE& + 1), 85, 4)
    Call REPLA(VDEF$, Chr$(0), 37, 1)
    Call REPLA(VDEF$, Chr$(0), 38, 1)
    '
110 XX$ = OBJPLA$("RECPROP", VDEF$)
    If XX$ = "" Then Exit Sub
    '
160 VDEF$ = XX$
    FECHA% = CVI(Left$(XX$, 2))
    If FECHA% > HOOI% Or FECHA% < 1 Then
        Call MENSERR(24, "Fecha Incorrecta")
        GoTo 110
    End If
    '
    'CI% = CVI(Mid$(XX$, 3, 2))
    'If TRIM$(CODEXT$(CI%)) = "" Then
    '    Call MENSERR(24, "Codigo Inexistente.")
    '    GoTo 110
    'End If
    'CODCON$ = CODEXT$(CI%)
    '
    DP1% = Asc(Mid$(XX$, 37, 1))
    DP2% = Asc(Mid$(XX$, 38, 1))
    If TRIM$(ECOARCH$("ECODEP", Chr$(DP1%))) = "" Then
        Call MENSERR(24, "Falta Depósito de Entrada")
        GoTo 110
    End If
    '
    If TRIM$(ECOARCH$("ECODEP", Chr$(DP2%))) = "" Then
        Call MENSERR(24, "Falta Depósito de Consumo")
        GoTo 110
    End If
    '
120 VTB% = VENTABU%("CAREPRO")
    If VTB% < 0 Then GoTo 110
    '
    CARE& = ULTREG&("!CAREPRO")
    If CARE& < 1 Then GoTo 120
    '
    For U& = 1 To CARE&
       XY$ = REGLEIDO$("!CAREPRO", U&)
       CANTI = CVD(Mid$(XY$, 23, 8))
       SCRAP = 0
       CI% = CVI(Left$(XY$, 2))
       CODICONJ$ = CODEXT$(CI%)
       CICONJ% = CI%
       '
       If CANTI < 0 Or SCRAP < 0 Or CANTI + SCRAP < 0.05 Then
          Call MENSERR(24, "Cantidades Incorrectas en Código " + TRIM$(CODICONJ$))
          GoTo 120
       End If
       '
       NORFABRI$ = ""
       CACONJU# = CDbl(CANTI + SCRAP)
       VACONTROL% = 1
       '
       Call SETULTREG("!MATEROF", 0)
       Call CIERRARCH("!MATEROF")
       'Call CARCOMPO(CODICONJ$, NORFABRI$, CACONJU#, VACONTROL%, 2)
       If VACONTROL% < 0 Then GoTo 120           ' cancelacion
       '
       NRO& = CVS(Mid$(XX$, 85, 4))
       NROOF& = CVS(Mid$(XY$, 19, 4))
       If NROOF& < 1 Then NROOF& = NRO&
       DOPRI$ = "OF-" + TRIM$(Str$(NROOF&))
       DOSEC$ = "PF-" + TRIM$(Str$(NRO&))
       LOTE$ = Mid$(XY$, 3, 16)
       If LOTE$ = "" Then
          LOTE$ = DOPRI$
       End If
       REFE$ = Mid$(XX$, 55, 30)
       If TRIM$(REFE$) = "" Then
          REFE$ = "Informe de Produccion"
       End If
       '
       NC% = 0
       VUNI# = 0
       MONE$ = ""
       CMOV$ = "PF"
       '
       DEPX% = DP1%: If DP1% = 99 Then DEPX% = DEPOPRE%(CICONJ%)
       'Call MOVISTOK(FECHA%, CICONJ%, LOTE$, CMOV$, DOPRI$, DOSEC$, REFE$, VUNI#, MONE$, NC%, DEPX%, CANTI)
       '
       YY$ = REGBLAN$("RECPROP")
       Call REPLA(YY$, MKI$(FECHA%), 1, 2)
       Call REPLA(YY$, Chr$(DP1%), 3, 1)
       Call REPLA(YY$, Chr$(DP2%), 4, 1)
       Call REPLA(YY$, REFE$, 5, 32)
       Call REPLA(YY$, MKI$(CI%), 37, 2)
       Call REPLA(YY$, MKS$(CANTI), 39, 4)
       Call REPLA(YY$, LOTE$, 43, 16)
       Call REPLA(YY$, MKS$(NRO&), 59, 4)
       Call REPLA(YY$, MKI$(-1), 63, 2)
       Call REGAPP("RECPROP", YY$)
    
       COD$ = CODEXT$(CICONJ%)
       CODCON$ = COD$
       '
       For U2& = 1 To ULTREG&("!MATEROF")
          ZZ$ = REGLEIDO$("!MATEROF", U2&)
          CI2% = CVI(Left$(ZZ$, 2))
          CAN = (-1) * CVS(Mid$(ZZ$, 5, 4))
          '
          If TRIM$(UCase$(UNIMED$(CI2%))) = "GK" Then
             CAN = CAN / 1000
          End If
          '
          If CI2% > 0 Then
             If CI2% <= NUMAS% Then
                If Abs(CAN) >= 0.05 Then
                   LOTE$ = ""
                   CMOV$ = "CF"
                   DOPRI$ = "OF-" + TRIM$(Str$(NROOF&))
                   DOSEC$ = "PF-" + TRIM$(Str$(NRO&))
                   REFE$ = "CONSUMO S/DESP." + CODCON$
                   DEPX% = DP2%: If DP2% = 99 Then DEPX% = DEPOPRE%(CI2%)
                   '
                   If STODEPOS(CI2%, DEPX%) < (-1) * CAN Then
                     If InStr(EPAC$, "YALTRE") > 0 Then
                       Call CARDEPOS
                       For KKII% = 1 To NUSU%
                         If STODEPOS(CI2%, COSU%(KKII%)) >= (-1) * CAN Then
                           DEPX% = COSU%(KKII%)
                           GoTo 124
                         End If
                       Next KKII%
                     End If
                   End If
                   '
124                'Call MovistoK(FECHA%, CI2%, LOTE$, CMOV$, DOPRI$, DOSEC$, REFE$, VUNI#, MONE$, NC%, DEPX%, CAN)
                End If
             End If
          End If
       Next U2&
       Call CIERRARCH("!MATEROF")
       '
129 Next U&
    Call CIERRARCH("MOVISTO")
    Call CIERRARCH("RECPROP")
    '
180 Call CIERRARCH("*.*")
    Call BLANARCH("!CAREPRO")
    Screen.MousePointer = 1
    GoTo 100
    '
200 MsgBox "falta orfcie"
    'Call ORFCIE
    Exit Sub
    '
End Sub

Sub LIUBIEXI()
    '
    Static NUSU%
    '
    TXT$ = "El Listado de Artículos y Ubicaciones Existente\"
    FEX = FEGEAR%("UBIEXI")
    If FEX < 1 Then GoTo 310
    TXT$ = TXT$ + "Corresponde al " + FECHATEX$(FEX) + " a las " + HORATEX$(HOGEAR%("UBIEXI")) + " hs.\"
    TXT$ = TXT$ + "\Consultar\Re-Generar\Abandonar"
    X% = COMALTER%(TXT$)
    On X% GoTo 320, 310
    Exit Sub
    '
310 Screen.MousePointer = 11
    '
    If NUSU% = 0 Then
        NUSU% = 1
        For I& = 1 To ULTREG&("SUCURSAL")
            XXX$ = REGLEIDO$("SUCURSAL", I&)
            If I& <= 62 Then
                If Asc(Left$(XXX$, 1)) > 32 Then
                    NUSU% = I&
                End If
            End If
        Next I&
        Call CIERRARCH("SUCURSAL")
    End If
    '
    Dim PRIRE&(32767)
    FIN& = ULTREG&("UBISTOCK")
    J& = 0
    '
    For I& = 1 To FIN&
        Call TRACE(20, I&, FIN&)
        XX$ = REGLEIDO$("UBISTOCK", I&)
        CI% = CVI(Left$(XX$, 2))
        UBI$ = AJUSTI$(Mid$(XX$, 3, 24), 24) + AJUSTI$(Mid$(XX$, 31, 30), 30)
        If TRIM$(UBI$) <> "" Then
            If CI% > 0 Then
                If CI% <= NUMAS% Then
                    YY$ = REGBLAN$("UBIEXI")
                    Call REPLA(YY$, MKI$(CI%), 1, 2)
                    Call REPLA(YY$, UBI$, 7, 54)
                    J& = J& + 1
                    Call GRAREG("UBIEXI", YY$, J&)
                    If PRIRE&(CI%) < 1 Then
                        PRIRE&(CI%) = J&
                    End If
                End If
            End If
        End If
    Next I&
    '
311 FIN& = NUMAS%
    For I& = 1 To FIN&
        Call TRACE(20, I&, FIN&)
        CI% = I&
        CAN = 0
        SALCA$ = REGLEIDO$("SALDEPOS", I& + 1)
        For U& = 1 To NUSU%
            CAN = CAN + CVS(Mid$(SALCA$, 4 * U& - 3, 4))
        Next U&
        '
        If Abs(CAN) >= 0.05 Then
            If PRIRE&(CI%) > 0 Then
                YY$ = REGLEIDO$("UBIEXI", PRIRE&(CI%))
                Call REPLA(YY$, MKS$(CAN), 3, 4)
                Call GRAREG("UBIEXI", YY$, PRIRE&(CI%))
              Else
                YY$ = REGBLAN$("UBIEXI")
                Call REPLA(YY$, MKI$(CI%), 1, 2)
                Call REPLA(YY$, MKS$(CAN), 3, 4)
                J& = J& + 1
                Call GRAREG("UBIEXI", YY$, J&)
                PRIRE&(CI%) = J&
            End If
        End If
    Next I&
    '
    Call SETULTREG("UBIEXI", J&)
    Call CIERRARCH("UBIEXI")
    Call GRARGEN("UBIEXI", "")
    '
320 Screen.MousePointer = 1
    FIN& = ULTREG&("UBIEXI")
    If FIN& < 1 Then
        Call CIERRARCH("*.*")
        Exit Sub
    End If
    '
330 XX1$ = Left$(REGLEIDO$("ECOMAST", 1), 16): CI1% = Abs(CODINT%(XX1$))
    XX2$ = Left$(REGLEIDO$("ECOMAST", ULTREG&("ECOMAST")), 16): CI2% = Abs(CODINT%(XX2$))
    VDEF$ = MKI$(CI1%) + MKI$(CI2%) + Chr$(0)
    '
    OBX$ = OBJPLA$("DESHASCOD", VDEF$)
    If OBX$ <> "" Then
        '
        Screen.MousePointer = 11
        CI1% = CVI(Left$(OBX$, 2))
        cod1$ = CODEXT$(CI1%)
        CI2% = CVI(Mid$(OBX$, 3, 2))
        cod2$ = CODEXT$(CI2%)
        '
        J& = 0
        For U& = 1 To FIN&
            Call TRACE(20, U&, FIN&)
            YY$ = REGLEIDO$("UBIEXI", U&)
            CIII% = CVI(Left$(YY$, 2))
            CODX$ = CODEXT$(CIII%)
            If CODX$ >= cod1$ Then
                If CODX$ <= cod2$ Then
                    J& = J& + 1
                    Call GRAREG("LUBIEXI", YY$, J&)
                End If
            End If
        Next U&
        Call SETULTREG("LUBIEXI", J&)
        Call CIERRARCH("LUBIEXI")
        Call CIERRARCH("*.*")
        Screen.MousePointer = 1
        Call FINAL("*SUBIEXI")
    End If
    '
End Sub
'
Sub GENCONPEN(MODOEMI%)
   '
   Screen.MousePointer = 11
   Dim PUNRE$(32767)
   FIN& = ULTREG&("MACONSIG")
   '
   For U& = 1 To FIN&
      '
      Call TRACE(20, U&, FIN&)
      X$ = REGLEIDO$("MACONSIG", U&)
      CAPEN# = CVD(Mid$(X$, 63, 8)) - CVD(Mid$(X$, 111, 8))
      If CAPEN# < 0 Then CAPEN# = 0
      If CVD(Mid$(X$, 119, 8)) <> CAPEN# Then
         Call REPLA(X$, MKD$(CAPEN#), 119, 8)
         Call GRAREG("MACONSIG", X$, U&)
      End If
      '
      If CAPEN# > 0.05 Then
         IPUN% = CVI(Mid$(X$, 3, 2))
         If MODOEMI% = 3 Then IPUN% = CVI(Mid$(X$, 33, 2))
         If IPUN% > 0 Then
            If IPUN% <= 32767 Then
               PUNRE$(IPUN%) = PUNRE$(IPUN%) + MKS$(U&)
            End If
         End If
      End If
      '
   Next U&
   '
   J& = 0
   If MODOEMI% = 2 Then
      FIN& = ULTREG&("PROVED1")
      For U& = 1 To FIN&
         Y$ = REGLEIDO$("PROVED1", U&)
         IPUN% = CVI(Left$(Y$, 2))
         If IPUN% > 0 Then
            If IPUN% <= 32767 Then
               If Len(PUNRE$(IPUN%)) >= 4 Then
                  For K& = 1 To Len(PUNRE$(IPUN%)) Step 4
                     RECO& = CVS(Mid$(PUNRE$(IPUN%), K&, 4))
                     If RECO& > 0 Then
                        X$ = REGLEIDO$("MACONSIG", RECO&)
                        J& = J& + 1
                        Call GRAREG("LCONPEN", X$, J&)
                     End If
                  Next K&
               End If
            End If
         End If
      Next U&
      '
    ElseIf MODOEMI% = 3 Then
      '
      FIN& = ULTREG&("INVERT")
      For U& = 2 To FIN&
         Y$ = REGLEIDO$("INVERT", U&)
         IPUN% = CVI(Mid$(Y$, 17, 2))
         If IPUN% > 0 Then
            If IPUN% <= 32767 Then
               If Len(PUNRE$(IPUN%)) >= 4 Then
                  For K& = 1 To Len(PUNRE$(IPUN%)) Step 4
                     RECO& = CVS(Mid$(PUNRE$(IPUN%), K&, 4))
                     If RECO& > 0 Then
                        X$ = REGLEIDO$("MACONSIG", RECO&)
                        J& = J& + 1
                        Call GRAREG("LCONPEN", X$, J&)
                     End If
                  Next K&
               End If
            End If
         End If
      Next U&
      '
   End If
   '
   Call SETULTREG("LCONPEN", J&)
   Call CIERRARCH("LCONPEN")
   '
   If MODOEMI% = 2 Then
        Call FINAL("*LCONPE2")
      ElseIf MODOEMI% = 3 Then
        Call FINAL("*LCONPE3")
   End If
   Screen.MousePointer = 1
End Sub

Function PUNOCPEN$()
   '
   Screen.MousePointer = 11
   PCPX$ = "": NUOCOA& = 0
   FIN& = ULTREG&("OCOMPRA")
   For U& = 1 To FIN&
     Call TRACE(20, U&, FIN&)
     XXX$ = REGLEIDO$("OCOMPRA", U&)
     NUOCO& = CVS(Left$(XXX$, 4))
     If NUOCO& > 0 Then
       If NUOCO& <> NUOCOA& Then
         If Asc(Mid$(XXX$, 62, 1)) < 1 Then ' ordenes abiertas
           If CVS(Mid$(XXX$, 58, 4)) < CVS(Mid$(XXX$, 13, 4)) - 0.1 Then
             PCPX$ = PCPX$ + MKS$(NUOCO&)
           End If
         End If
       End If
       NUOCOA& = NUOCO&
     End If
   Next U&
   '
   NUOCOA& = 0
   FIN& = ULTREG&("OCOMDETA")
   For U& = 1 To FIN&
     Call TRACE(20, U&, FIN&)
     XXX$ = REGLEIDO$("OCOMDETA", U&)
     NUOCO& = CVS(Left$(XXX$, 4))
     If NUOCO& > 0 Then
       If NUOCO& <> NUOCOA& Then
         If Asc(Mid$(XXX$, 124, 1)) < 8 Then ' ordenes abiertas
           If CVS(Mid$(XXX$, 99, 4)) < CVS(Mid$(XXX$, 55, 4)) - 0.1 Then
             PCPX$ = PCPX$ + MKS$(NUOCO&)
           End If
         End If
       End If
       NUOCOA& = NUOCO&
     End If
   Next U&
   '
   PUNOCPEN$ = PCPX$
   Screen.MousePointer = 1

End Function
