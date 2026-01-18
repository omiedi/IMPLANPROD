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
    Call VERISAL(CI%)
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
            NROPED& = CVS(Left$(X$, 4))
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
            Call REPLA(LI2TEX$, CSTRING$(MKS$(NROPED&), 3, 5, 0, 2), 1, 5)
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
    CONEXI.Label5.Caption = CSTRING$(MKS$(PEDI), 4, 12, 2, 2)
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
    UINI& = Val(CONTROL$("ADMIPED", "PRIPEDPE"))
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
            STATUS% = Asc(Mid$(X$, 300, 1))
            If STATUS% <= 1 Then
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
        Call MOVISTO(FECHA%, CI%, LOTE$, CMOV$, DOPRI$, DOSEC$, REFE$, VUNI#, MONE$, NC%, DEPO%, TOIN)
    End If
    '
    If RECH > 0 Then
        REFE$ = "Rechazo en Recepcion"
        Call MOVISTO(FECHA%, CI%, LOTE$, CMOV$, DOPRI$, DOSEC$, REFE$, VUNI#, MONE$, NC%, DEPO%, (-1) * RECH)
    End If
    '
    If Abs(FALT) > 0 Then
        REFE$ = "Diferencia en Recepcion"
        Call MOVISTO(FECHA%, CI%, LOTE$, CMOV$, DOPRI$, DOSEC$, REFE$, VUNI#, MONE$, NC%, DEPO%, (-1) * FALT)
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
    Call CIERRARCH("MOVISTO")
    Call CIERRARCH("*.*")
    Screen.MousePointer = 1
    GoTo 10
    '
End Sub
'
Sub TRASTOCK()
    '
    HOOI% = HOY(HO$)
    FORIPRE$ = TRIM$(CONTROL$("", "FORTRAST"))
    '
100 REGXX& = ULTREG&("TRANSFE")
    '
    Screen.MousePointer = 11
    NRO& = 1
    For U& = ULTREG&("MOVISTO") To 1 Step -1
      XX$ = REGLEIDO$("MOVISTO", U&)
      If UCase$(Mid$(XX$, 21, 2)) = "TR" Then
        NRO& = 1 + Val(Mid$(XX$, 26, 7))
        GoTo 110
      End If
    Next U&
    '
110 Screen.MousePointer = 1
    VDEF$ = String$(6, 0) + MKI$(HOOI%) + Space$(16) + MKS$(0) + Space$(32)
    OPX% = COMALTER%("Opciones de Procesamiento:\\Item único\Lista Múltiple\Cancelar")
    If OPX% < 1 Or OPX% > 2 Then Exit Sub
    '
    If OPX% = 2 Then
      '\\\OT-5-0403-FG-15/06/05///
      FORIPRE$ = TRIM$(CONTROL$("", "FORTRSTM"))
      '\\\OT-5-0403-FG-FIN///
      '
115   X22$ = OBJPLA$("TRANSTOMUL", VDEF$)
      If X22$ = "" Then
        Exit Sub                        ' abandonar
      End If
      '
      VDEF$ = X22$
      DEPORIG% = Asc(Mid$(X22$, 3, 1)): DEPOSI$ = ECOARCH$("ECODEP", Chr$(DEPORIG%))
        If DEPOSI$ = "" Then
           Call MENSERR(24, "Depósito Origen Inexistente o No Válido")
           GoTo 115
        End If
      DEPODES% = Asc(Mid$(X22$, 6, 1)): DEPODE$ = ECOARCH$("ECODEP", Chr$(DEPODES%))
        If DEPODE$ = "" Then
           Call MENSERR(24, "Depósito Destino Inexistente o No Válido")
           GoTo 115
        End If
      FEMOVI% = CVI(Mid$(X22$, 7, 2))
      REFMO$ = TRIM$(Mid$(X22$, 29, 32))
      If REFMO$ = "" Then
        Call MENSERR(24, "Referencia o Motivo es Mandatoria")
        GoTo 115
      End If
      '
1112  Call SETULTREG("!BOLREDET", 0)
      Call CIERRARCH("*.*")
      '
      Screen.MousePointer = 11
      '
      FFF$ = REPLACAR$(FECHATEX$(FEMOVI%), "/", "-")
1113  VTB% = VENTABU%("BOLREDET/NC/TITUPAN=Transferencia de Stock - De " + TRIM$(DEPOSI$) + " a " + TRIM$(DEPODE$) + " - " + FFF$)
      If VTB% < 0 Or ULTREG&("!BOLREDET") < 1 Then
         GoTo 110
      End If
      '
      CONTA% = ULTREG&("!BOLREDET")
      '
1115  If CONTA% > 0 Then
        Screen.MousePointer = 11  'Call BLIMESS("Registrando Recepcion ...")
        '
        AINDI% = 0
        NUORIT% = (-1)
        For YX& = 1 To ULTREG&("!BOLREDET")
          YY$ = REGLEIDO$("!BOLREDET", YX&)
          CI% = CVI(Mid$(YY$, 83, 2))
          CAN# = CVD(Mid$(YY$, 89, 8))
          LOTE$ = Mid$(YY$, 69, 12)
            If TRIM$(LOTE$) = "" Then LOTE$ = Mid$(YY$, 107, 12)
          REMI$ = "TR." + TRIM$(Str$(NRO&))
          If CI% > 0 Then
            If CI% <= NUMAS% Then
              If CAN# > 0.005 Then
                DOSEC$ = ""
1120            CANTI = (-1) * CAN#
                Call MOVISTO(FEMOVI%, CI%, LOTE$, "TR", REMI$, DOSEC$, REFMO$, 0, "$", 0, DEPORIG%, CANTI)
                Call MOVISTO(FEMOVI%, CI%, LOTE$, "TR", REMI$, DOSEC$, REFMO$, 0, "$", 0, DEPODES%, (-1) * CANTI)
              End If
            End If
          End If
        Next YX&
        Call CIERRARCH("MOVISTO")
        '
        '\\\OT-5-0403-FG-15/06/05///
        'IMPRESION DEL VALE DE TRANSFERENCIA MULTIPLE
        '\\\OT-5-597-FG-25/08/05///
        If FORIPRE$ <> "" Then
          LU$ = LUDAT$
          If EXISTE%(LU$ + FORIPRE$ + ".FRM") = 1 Then
            CODPARAM$(1) = "NUME-COM": DOCPARAM$(1) = CStr(NRO&)
            CODPARAM$(2) = "FECH-EMI": DOCPARAM$(2) = FECHATEX(FEMOVI%)
            CODPARAM$(3) = "DEPOSITO": DOCPARAM$(3) = DEPOSI$
            CODPARAM$(4) = "DESTINO": DOCPARAM$(4) = DEPODE$
            CODPARAM$(5) = "REF-MAT1": DOCPARAM$(5) = REFMO$
            CAPARDOC% = 5
            '
            CODTABU1$(1) = "COD-MPRI"
            CODTABU1$(2) = "DES-MPRI"
            CODTABU1$(3) = "UNI-MPRI"
            CODTABU1$(4) = "CAN-MPRI"
            CODTABU1$(5) = "TIP-MPRI"
            CACOLTAB1% = 5
            '
            CAITAB1% = 0
            '
            For YX& = 1 To ULTREG&("!BOLREDET")
              YY$ = REGLEIDO$("!BOLREDET", YX&)
              CI% = CVI(Mid$(YY$, 83, 2))
              CAN# = CVD(Mid$(YY$, 89, 8))
              NOPED& = CVS(Mid$(YY$, 151, 4))
              '
              CAITAB1% = CAITAB1% + 1
                CODMAT$ = TRIM$(CODEXT(CI%))
              TETABU1$(1, CAITAB1%) = CODMAT$
                DESMAT$ = DESCRIT$(CI%)
              TETABU1$(2, CAITAB1%) = DESMAT$
                UNIMAT$ = UNIMED$(CI%)
              TETABU1$(3, CAITAB1%) = UNIMAT$
                CANTMAT$ = FORMATNUM(CAN#, "F12.2")
              TETABU1$(4, CAITAB1%) = CANTMAT$
                NOPEDIX$ = FORMATNUM$(NOPED&, "F6.0")
              TETABU1$(5, CAITAB1%) = NOPEDIX$
            Next YX&
            '
            Call PRINTDOC("FORTRSTM")
          End If
        End If
        '\\\OT-5-597-FG-FIN///
        '\\\OT-5-0403-FG-FIN///
        '
        Call COMUNI("Transferencia Procesada")
        '
      End If
      GoTo 110
    End If
    '
140 Screen.MousePointer = 1
    VDEF$ = String$(6, 0) + MKI$(HOOI%) + Space$(16) + MKS$(0) + Space$(32)
150 X22$ = OBJPLA$("TRANSFESTOCK", VDEF$)
    If X22$ = "" Then
        Exit Sub                        ' abandonar
    End If
    '
    VDEF$ = X22$
    TRANSFE$ = X22$
    '
    CI1% = CVI(Left$(TRANSFE$, 2))
    DP1% = Asc(Mid$(TRANSFE$, 3, 1))
    CI2% = CVI(Mid$(TRANSFE$, 4, 2))
    DP2% = Asc(Mid$(TRANSFE$, 6, 1))
    FECHA% = CVI(Mid$(TRANSFE$, 7, 2))
      If FECHA% < 1 Then FECHA% = HOY(HO$)
    LOTE$ = Mid$(TRANSFE$, 9, 16)
    CAN = CVS(Mid$(TRANSFE$, 25, 4))
    REFE0$ = TRIM$(Mid$(TRANSFE$, 29, 32))
    '
    DEPOSI$ = ECOARCH$("ECODEP", Chr$(DP1%))
    If DEPOSI$ = "" Then
       Call MENSERR(24, "Depósito Origen Inexistente o No Válido")
       GoTo 150
    End If
    DEPODE$ = ECOARCH$("ECODEP", Chr$(DP2%))
    If DEPODE$ = "" Then
       Call MENSERR(24, "Depósito Destino Inexistente o No Válido")
       GoTo 150
    End If
    '
    If CI1% > 0 Then
      If CI1% <= NUMAS% Then
        If DP1% > 0 Then
           If CI2% > 0 Then
              If CI2% <= NUMAS% Then
                If DP2% > 0 Then
                   If CAN >= 0.1 Then
                      GoTo 160
                   End If
                End If
              End If
           End If
        End If
      End If
    End If
    '
    Call MENSERR(24, "Datos Incompletos !!!\Transferencia Abortada.")
    GoTo 150
    '
160 NC% = 0
    VUNI# = 0
    MONE$ = ""
    CMOV$ = TRIM$(CONTROL$("", "CODTRAST"))
    If CMOV$ = "" Then CMOV$ = "TR"
    '
    DOPRI$ = "TR." + TRIM$(Str$(NRO&))
    DOSEC$ = "TR." + Mid$(HO$, 8, 1) + Mid$(HO$, 4, 2) + Left$(HO$, 2)
                              While Len(DOSEC$) < 10
                                DOSEC$ = DOSEC$ + " "
                              Wend
                              '
    COD$ = LTrim$(RTrim$(CODEXT$(CI2%)))
    REFE$ = TRIM$(REFE0$)
    If REFE$ = "" Then
       REFE$ = "TRANSF. A " + COD$ + " - D." + Str$(DP2%)
    End If
    '
    CANTI = (-1) * CAN
    Call BLOQARCH("MOVISTO")
    Call BLOQARCH("TRANSFE")
    '
    Call MOVISTO(FECHA%, CI1%, LOTE$, CMOV$, DOPRI$, DOSEC$, REFE$, VUNI#, MONE$, NC%, DP1%, CANTI)
    '
    COD$ = LTrim$(RTrim$(CODEXT$(CI1%)))
    REFE$ = TRIM$(REFE0$)
    If REFE$ = "" Then
       REFE$ = "TRANSF. DE " + COD$ + " - D." + Str$(DP1%)
    End If
    '
    CANTI = CAN
    Call MOVISTO(FECHA%, CI2%, LOTE$, CMOV$, DOPRI$, DOSEC$, REFE$, VUNI#, MONE$, NC%, DP2%, CANTI)
    Call CIERRARCH("MOVISTO")
    '
    XXX$ = REGBLAN$("TRANSFE")
    Call REPLA(XXX$, MKI$(FECHA%), 1, 2)
    Call REPLA(XXX$, Chr$(DP1%), 3, 1)
    Call REPLA(XXX$, Chr$(DP2%), 4, 1)
    Call REPLA(XXX$, MKI$(CI1%), 5, 2)
    Call REPLA(XXX$, MKI$(CI2%), 7, 2)
    Call REPLA(XXX$, LOTE$, 9, 16)
    Call REPLA(XXX$, LOTE$, 25, 16)
    Call REPLA(XXX$, MKS$(CAN), 41, 4)
    Call REPLA(XXX$, MKS$(NRO&), 45, 4)
    Call REGAPP("TRANSFE", XXX$)
    Call CIERRARCH("TRANSFE")
    '
    Call CIERRARCH("*.*")
    '
'IMPRIMIR FORMULARIO
237 If TRIM$(FORIPRE$) <> "" Then
             LU$ = LUDAT$
             CODFOR$ = TRIM$(UCase$(FORIPRE$))
             '
             YAC = Val(ATRIFORM$(CODFOR$, "DETDESMM"))
             MARIZW = Val(ATRIFORM$(CODFOR$, "MARIZQMM"))
             MARSUW = Val(ATRIFORM$(CODFOR$, "MARSUPMM"))
             '
             INTERLI = Val(ATRIFORM$(CODFOR$, "INTERLIN"))
             If INTERLI < 1 Then
                Call COMUNI("Error de Configuración de Interlinea\en Formulario '" + CODFOR$ + ".FRM'.")
                INTERLI = 1
             End If
             '
             CACOPI% = Val(ATRIFORM$(CODFOR$, "CAN-COP"))
             If CACOPI% < 1 Or CACOPI% > 6 Then
                Call COMUNI("Error de Configuración de Cantidad de Copias\en Formulario '" + CODFOR$ + ".FRM'.")
                CACOPI% = 1
             End If
             '
             For COP% = 1 To CACOPI%
               '
               Call CARFORWIN
               '
               FEX = FECHA%
               Call PRINTFORWIN(FECHATEX$(FEX), "FECH-EMI", 0, 0)
               Call PRINTFORWIN(Time$, "HORA-EMI", 0, 0)
               Call PRINTFORWIN(USERNAME$, "USUARIO", 0, 0)
               Call PRINTFORWIN(Str$(NRO&), "NUME-COM", 0, 0)
               Call PRINTFORWIN(TRIM$(CODEXT$(CI1%)), "COD-MAT", 0, YAC)
               Call PRINTFORWIN(TRIM$(CODEXT$(CI2%)), "COD-MAT", 0, YAC + INTERLI)
               Call PRINTFORWIN(DESCRIT$(CI1%), "DES-MAT", 0, YAC)
               Call PRINTFORWIN(DESCRIT$(CI2%), "DES-MAT", 0, YAC + INTERLI)
               Call PRINTFORWIN(UNIMED$(CI1%), "UNIMED", 0, YAC)
               Call PRINTFORWIN(UNIMED$(CI2%), "UNIMED", 0, YAC + INTERLI)
               Call PRINTFORWIN(ECOARCH$("ECODEP", Chr$(DP1%)), "DEPOSITO", 0, YAC)
               Call PRINTFORWIN(ECOARCH$("ECODEP", Chr$(DP2%)), "DEPOSITO", 0, YAC + INTERLI)
               Call PRINTFORWIN(LOTE$, "NRO-LOTE", 0, YAC)
               CANS$ = CSTRING$(MKS$(CAN), 3, 9, 1, 2)
               Call PRINTFORWIN(CANS$, "CANTIDAD", 0, YAC)
               Call PRINTFORWIN(REFE0$, "REF-MAT1", 0, YAC)
               '
               'Printer.EndDoc
               Call SETSPOOL("ENDDOC", "")
               '
             Next COP%
             Call SAVESPOOL("Transferencia de Stock" + Str$(NRO&), OKX%)
    End If
    '
299 Call COMUNI("Transferencia Completa\y Verificada !!!")
    '
    TRANSFE$ = ""
    GoTo 100
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
20  Call ORFCIE
    Exit Sub
    '
100 If DERACCE%("ARMACON", "Cierre por Armado de Conjunto") < 2 Then Exit Sub
    '
    Screen.MousePointer = 11
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
    '\\\OT-06-0093-WAR-15/03/06///
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
    Next U&
    '\\\OT-06-0093-WAR-FIN///
    For U& = 1 To CARE&
       XY$ = REGLEIDO$("!CAREPRO", U&)
       CANTI = CVD(Mid$(XY$, 23, 8))
       SCRAP = 0
       CI% = CVI(Left$(XY$, 2))
       CODICONJ$ = CODEXT$(CI%)
       CICONJ% = CI%
       '
       NORFABRI$ = ""
       CACONJU# = CDbl(CANTI + SCRAP)
       VACONTROL% = 1
       '
       Call SETULTREG("!MATEROF", 0)
       Call CIERRARCH("!MATEROF")
       Call CARCOMPO(CODICONJ$, NORFABRI$, CACONJU#, VACONTROL%, 2)
       If VACONTROL% < 0 Then            ' cancelacion
         Call COMUNI("No Admisible Cancelar\Transacción en Proceso.")
       End If
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
       'Call MOVISTO(FECHA%, CICONJ%, LOTE$, CMOV$, DOPRI$, DOSEC$, REFE$, VUNI#, MONE$, NC%, DEPX%, CANTI)
       Call ARMAVECMOVI(FECHA%, CICONJ%, LOTE$, CMOV$, DOPRI$, DOSEC$, REFE$, VUNI#, MONE$, NC%, DEPX%, CANTI)
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
124                Call ARMAVECMOVI(FECHA%, CI2%, LOTE$, CMOV$, DOPRI$, DOSEC$, REFE$, VUNI#, MONE$, NC%, DEPX%, CAN)
                   'Call MOVISTO(FECHA%, CI2%, LOTE$, CMOV$, DOPRI$, DOSEC$, REFE$, VUNI#, MONE$, NC%, DEPX%, CAN)
                End If
             End If
          End If
       Next U2&
       Call CIERRARCH("!MATEROF")
       '
129 Next U&
    Call ARMAVECMOVI(-1, 0, "", "", "", "", "", 0, "", 0, 0, 0)
    Call CIERRARCH("MOVISTO")
    Call CIERRARCH("RECPROP")
    '
180 Call CIERRARCH("*.*")
    Call BLANARCH("!CAREPRO")
    Screen.MousePointer = 1
    GoTo 100
    '
200 Call ORFCIE
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
Sub GENCONPEN(MODOEMI1%)
   '
   Screen.MousePointer = 11
   Dim PUNRE$(32767)
   FIN& = ULTREG&("MACONSIG")
   '
   MODOEMI% = MODOEMI1% Mod 10
   NOLIS% = 0: If MODOEMI1% > 10 Then NOLIS% = 1
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
   If NOLIS% = 0 Then    ' BANDERA PARA NO LISTAR
     If MODOEMI% = 2 Then
          Call FINAL("*LCONPE2")
        ElseIf MODOEMI% = 3 Then
          Call FINAL("*LCONPE3")
     End If
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
'\\\OT-06-0093-WAR-15/03/06///
Sub ARMAVECMOVI(FECHA%, CICONJ%, LOTE$, CMOV$, DOPRI$, DOSEC$, REFE$, VUNI#, MONE$, NC%, DEPX%, CANTI)
  '
  Static CONTA71%, FECHA71%(1024), CICONJ71%(1024), LOTE71$(1024), CMOV71$(1024), DOPRI71$(1024), DOSEC71$(1024), REFE71$(1024), VUNI71#(1024), MONE71$(1024), NC71%(1024), DEPX71%(1024), CANTI71(1024)
  'Si se llama para grabar el vector en el movisto
  If FECHA% < 0 Or CONTA71% = 1024 Then
    Call BLOQARCH("MOVISTO")
    For I% = 1 To CONTA71%
      Call MOVISTO(FECHA71%(I%), CICONJ71%(I%), LOTE71$(I%), CMOV71$(I%), DOPRI71$(I%), DOSEC71$(I%), REFE71$(I%), VUNI71#(I%), MONE71$(I%), NC71%(I%), DEPX71%(I%), CANTI71(I%))
    Next I%
    Call CIERRARCH("MOVISTO")
    For I% = 1 To CONTA71%
      FECHA71%(CONTA71%) = 0
      CICONJ71%(CONTA71%) = 0
      LOTE71$(CONTA71%) = ""
      CMOV71$(CONTA71%) = ""
      DOPRI71$(CONTA71%) = ""
      DOSEC71$(CONTA71%) = ""
      REFE71$(CONTA71%) = ""
      VUNI71#(CONTA71%) = 0
      MONE71$(CONTA71%) = ""
      NC71%(CONTA71%) = 0
      DEPX71%(CONTA71%) = 0
      CANTI71(CONTA71%) = 0
    Next I%
    CONTA71% = 0
    If FECHA% < 0 Then Exit Sub
  End If
  'Si se llama para armar vector
  CONTA71% = CONTA71% + 1
  FECHA71%(CONTA71%) = FECHA%
  CICONJ71%(CONTA71%) = CICONJ%
  LOTE71$(CONTA71%) = LOTE$
  CMOV71$(CONTA71%) = CMOV$
  DOPRI71$(CONTA71%) = DOPRI$
  DOSEC71$(CONTA71%) = DOSEC$
  REFE71$(CONTA71%) = REFE$
  VUNI71#(CONTA71%) = VUNI#
  MONE71$(CONTA71%) = MONE$
  NC71%(CONTA71%) = NC%
  DEPX71%(CONTA71%) = DEPX%
  CANTI71(CONTA71%) = CANTI
  '
End Sub
'\\\OT-06-0093-WAR-FIN///
