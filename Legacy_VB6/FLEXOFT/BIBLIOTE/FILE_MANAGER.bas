Attribute VB_Name = "FILE_MANAGER"
Public AREXTRA$
Dim ABLO$(1024), STABLO%(1024)
Dim CARBLO As Integer
'
' ATENCION CON EL LIMITE DE FECHA SETEADO PARA EL 27/08/2000
'

Sub FILEMGR(Archi$, FINBR%, LOREG%, CAREG&, REG&, REGS$, OPCI%)
'
'              ' en ecoarch falta detectar error de indexacion de archivo
'              ' en ecoarch pasar a busqueda secuencial si no lo encuentra
'
Static CTX%, CARCHI%, CABIER%, MX%, MAXABIER%
Static AX$(), FINUX%(), LRG%(), CRG&(), RGL&()
Static RGLS$(), RGB$(), PRIMOD&(), ULTACX()
Static IXC$(), LXC$()
Static PX%(), TX%()
'
' ************************************************ inicializacion
'
SMA = Screen.MousePointer
If CTX% = 0 Then
    '
'   If HOY(HO$) > 8603 Then
'       Close: End
'   End If
    '
    CARCHI% = 0             ' total de archivos
    CABIER% = 0             ' archivos abiertos
    MAXABIER% = 64           ' maximo de archivos abiertos
    MX% = 64                ' total maximo de archivos
    CTX% = 1                ' variable de control
    '
    ReDim AX$(MX%)
    ReDim FINUX%(MX%)
    ReDim LRG%(MX%)
    ReDim CRG&(MX%)
    ReDim RGL&(MX%)
    ReDim RGLS$(MX%)
    ReDim RGB$(MX%)
    ReDim IXC$(MX%)
    ReDim LXC$(MX%)
    ReDim PRIMOD&(MX%)
    ReDim ULTACX(MX%)
    '
    ReDim PX%(32)
    ReDim TX%(32)
    '
End If
'
If Archi$ = "" Then GoTo 10
'
If Left$(TRIM$(Archi$), 1) = "*" Then
    If OPCI% = 7 Then
        GoSub cierrato
        GoTo 999
      Else
        Call MENSERR(24, "Imposible Abrir Archivo '" + TRIM$(Archi$) + "' - Consulte")
        Call FINAL("")
    End If
End If
'
' GoSub 910  ' arma el nombre completo
MENSAI% = 0
If InStr(UCase$(Archi$), "MENSAJE") > 0 Then
    MENSAI% = 1
End If
'
GoSub 920         ' determinar numero de orden del archivo
'
If FINUX%(IARCH%) < 1 Or (OPCI% = 11 And OPCI% = 14) Then
    GoSub 940                        ' abre efectivamente el archivo
End If
'
If OPCI% = 1 Or OPCI% = 5 Or NUEVO% = 1 Or CRG&(IARCH%) < 1 Then
    GoSub 950                        ' re-confirma fin de archivo
    If NUEVO% = 1 Then
       CARETO& = LOF(FINUX%(IARCH%)) / LRG%(IARCH%)
       If CRG&(IARCH%) < CARETO& - 128 Then
           Call BLOQARCH(AX$(IARCH%))
           Screen.MousePointer = 11
           FINUY% = FILEOPEN%("C:\FLEXOFT\TRANSIT.$$$", 0, LRG%(IARCH%))
           UUZ& = 0
           For UUY& = 1 To CRG&(IARCH%)
              XXXY$ = LEEREG$(FINUX%(IARCH%), LRG%(IARCH%), UUY&)
              Call GRABAREG(FINUY%, LRG%(IARCH%), XXXY$, UUY&)
           Next UUY&
           Close #FINUX%(IARCH%): FINUX%(IARCH%) = 0
           '
           ARY1$ = AX$(IARCH%): GoSub 910
           LDDT$ = LUDAT$
           'FILEKILL ARY2$
           On Error GoTo 3
           Kill ARY2$
           On Error Resume Next
           Kill LDDT$ + AX$(IARCH%) + ".X*"
           On Error GoTo 0
           '
           ' Open ARY2$ For Random Access Read Write Shared As #FINUX%(IARCH%) Len = LRG%(IARCH%)
           FINUZ% = FILEOPEN%(ARY2$, 0, LRG%(IARCH%)) ' AQUI
           FINUX%(IARCH%) = FINUZ%
           RGL&(IARCH%) = 0
           PRIMOD&(IARCH%) = 0
           '
           For UUY& = 1 To CRG&(IARCH%)
              XXXY$ = LEEREG$(FINUY%, LRG%(IARCH%), UUY&)
              Call GRABAREG(FINUX%(IARCH%), LRG%(IARCH%), XXXY$, UUY&)
           Next UUY&
           '
           Close #FINUY%: FINUY% = 0
           Close #FINUX%(IARCH%): FINUX%(IARCH%) = 0
           Call LIBARCH(AX$(IARCH%))
           FINUZ% = FILEOPEN%(ARY2$, 0, LRG%(IARCH%)) ' AQUI
           FINUX%(IARCH%) = FINUZ%
           RGL&(IARCH%) = 0
           PRIMOD&(IARCH%) = 0
           '
           On Error Resume Next
           Kill "C:\FLEXOFT\TRANSIT.$$$"
           On Error GoTo 0
           '
           GoSub ixarch
           '
       End If
    End If
End If
On Error GoTo 0
GoTo 5
'
3   Resume 4
4   On Error GoTo 0
    Call LIBARCH(AX$(IARCH%))
    FINUZ% = FILEOPEN%(ARY2$, 0, LRG%(IARCH%)) ' AQUI
    FINUX%(IARCH%) = FINUZ%
    '
5   ULTACX(IARCH%) = Timer
    '
10  If OPCI% = 1 Then
        FINBR% = FINUX%(IARCH%)
        LOREG% = LRG%(IARCH%)
        CAREG& = CRG&(IARCH%)
        GoTo 999
   End If

20 If OPCI% = 2 Then
        REGS$ = RGB$(IARCH%)
        GoTo 999
   End If

30 If OPCI% = 3 Then
        If RGL&(IARCH%) < 1 Or RGL&(IARCH%) > CRG&(IARCH%) Then
            REGS$ = RGB$(IARCH%)
            REG& = RGL&(IARCH%)
            GoTo 999
        End If
        LOREG% = LRG%(IARCH%)
        RGLS$(IARCH%) = LEEREG$(FINUX%(IARCH%), LOREG%, RGL&(IARCH%))
'        Get #FINUX%(IARCH%), RGL&(IARCH%), RGLS$(IARCH%)
        REGS$ = RGLS$(IARCH%)
        REG& = RGL&(IARCH%)
        GoTo 999
   End If



40 If OPCI% = 4 Then
        RGL&(IARCH%) = RGL&(IARCH%) + 1
        If RGL&(IARCH%) > CRG&(IARCH%) Then
             RGL&(IARCH%) = CRG&(IARCH%) + 1
             REGS$ = ""
             GoTo 999
        End If
        LOREG% = LRG%(IARCH%)
        Get #FINUX%(IARCH%), RGL&(IARCH%), RGLS$(IARCH%)
        RGLS$(IARCH%) = XXX
        REGS$ = RGLS$(IARCH%)
        GoTo 999
   End If

50 If OPCI% = 5 Then
        RGL&(IARCH%) = CRG&(IARCH%) + 1
        REG& = RGL&(IARCH%)
        If REG& > 48 Then
           If VERDEMO% <> (-1) Then
              If MENSAI% <> 1 Then
                 OPX% = COMALTER%("Versión Limitada para Demostración.\  \Imposible Grabar a Archivos\\Aceptar Mensaje")
                 Call FUERASYS
              End If
            End If
        End If
        CRG&(IARCH%) = CRG&(IARCH%) + 1
        Call GRABAREG(FINUX%(IARCH%), LRG%(IARCH%), REGS$, RGL&(IARCH%))
        If RGL&(IARCH%) < PRIMOD&(IARCH%) Then PRIMOD&(IARCH%) = RGL&(IARCH%)
'For RGBLX& = CRG&(IARCH%) + 1 To CRG&(IARCH%) + 33
'Call GRABAREG(FINUX%(IARCH%), LRG%(IARCH%), String$(LRG%(IARCH%), 0), RGBLX&)
'Next RGBLX&
        GoTo 999
   End If

60 If OPCI% = 6 Then
        REGS$ = RGB$(IARCH%)
        Call GRABAREG(FINUX%(IARCH%), LRG%(IARCH%), REGS$, 1&)
        Close #FINUX%(IARCH%): FINUX%(IARCH%) = 0
        '
        ARY1$ = AX$(IARCH%): GoSub 910
        FILEKILL ARY2$
        On Error Resume Next
        Kill LUDAT$ + AX$(IARCH%) + ".X*"
        On Error GoTo 0
        '

        'Open ARY2$ For Random Access Read Write Shared As #FINUX%(IARCH%) Len = LRG%(IARCH%)
        FINUZ% = FILEOPEN%(ARY2$, 0, LRG%(IARCH%)) 'AQUI
        FINUX%(IARCH%) = FINUZ%
        RGL&(IARCH%) = 0
        CRG&(IARCH%) = 0
        PRIMOD&(IARCH%) = 0
        GoTo 999
   End If

70 If OPCI% = 7 Then
If Archi$ <> UASD$ Then
   UASD$ = Archi$
End If
        GoSub ixarch
        If CRG&(IARCH%) > 0 Then ' fuerza posicion al principio del archivo
            XXXY$ = LEEREG$(FINUX%(IARCH%), LRG%(IARCH%), 1)
        End If
        Close #FINUX%(IARCH%): FINUX%(IARCH%) = 0
                             RGB$(IARCH%) = ""
                             IXC$(IARCH%) = ""
                             LXC$(IARCH%) = ""
                             CRG&(IARCH%) = 0
                             PRIMOD&(IARCH%) = 0
                             CABIER% = CABIER% - 1
        GoTo 999
   End If


80 If OPCI% = 8 Then
        RGL&(IARCH%) = REG&
        If RGL&(IARCH%) > CRG&(IARCH%) Then
             REG& = CRG&(IARCH%) + 1
             RGL&(IARCH%) = CRG&(IARCH%) + 1
             REGS$ = ""
             GoTo 999
        ElseIf RGL&(IARCH%) < 1 Then
             REGS$ = RGB$(IARCH%)
             GoTo 999
        Else
             LOREG% = LRG%(IARCH%)
             RGLS$(IARCH%) = LEEREG$(FINUX%(IARCH%), LOREG%, RGL&(IARCH%))
             REGS$ = RGLS$(IARCH%)
             GoTo 999
        End If
   End If

90 If OPCI% = 9 Then
        RGL&(IARCH%) = REG&
        If RGL&(IARCH%) > CRG&(IARCH%) Or REG& < 1 Then
             REG& = CRG&(IARCH%) + 1
             RGL&(IARCH%) = CRG&(IARCH%) + 1
             REGS$ = ""
             GoTo 999
        End If
        '
        REGS$ = LEEREG$(FINUX%(IARCH%), LRG%(IARCH%), RGL&(IARCH%))
        GoTo 999
   End If

100 If OPCI% = 10 Then
        If REG& < 1 Then
           Call MENSERR(24, "Error - Intento de Grabar fuera de Limites de Archivo")
           GoTo 999
        End If
        '
        If REG& > 48 Then
           If VERDEMO% <> (-1) Then
              If MENSAI% <> 1 Then
                 OPX% = COMALTER%("Versión Limitada para Demostración.\  \Imposible Grabar a Archivos\\Aceptar Mensaje")
                 Call FUERASYS
              End If
           End If
        End If
        '
        For U& = 1 + CRG&(IARCH%) To REG&
           REGSTT$ = String$(LRG%(IARCH%), 0)
           Call GRABAREG(FINUX%(IARCH%), LRG%(IARCH%), REGSTT$, U&)
        Next U&
        '
        RGL&(IARCH%) = REG&
        If REG& > CRG&(IARCH%) Then CRG&(IARCH%) = REG&
        Call GRABAREG(FINUX%(IARCH%), LRG%(IARCH%), REGS$, REG&)
        If REG& < PRIMOD&(IARCH%) Then PRIMOD&(IARCH%) = REG&
        GoTo 999
   End If

110 If OPCI% = 11 Then
        PPX% = REG&: REG& = 0
        If PPX% > 0 Then
           If PPX% <= Len(LXC$(IARCH%)) Then
              REG& = Asc(Mid$(LXC$(IARCH%), PPX%, 1))
           End If
        End If
        GoTo 999
    End If

120 If OPCI% = 12 Then
        GoSub datarch
        If (REG& > 0 And REG& <= CAFI%) Then REG& = PX%(REG&) Else REG& = 0
        GoTo 999
    End If

130 If OPCI% = 13 Then
        GoSub datarch
        If (REG& > 0 And REG& <= CAFI%) Then REG& = TX%(REG&) Else REG& = 0
        GoTo 999
    End If

140 If OPCI% = 14 Then
        PPX% = REG&: REG& = 0
        If PPX% > 0 Then
           If PPX% <= Len(IXC$(IARCH%)) Then
              REG& = Asc(Mid$(IXC$(IARCH%), PPX%, 1))
           End If
        End If
        GoTo 999
    End If

150 If OPCI% = 15 Then
        GoSub datarch
        REGS$ = TITU$
        GoTo 999
    End If

240 If OPCI% = 24 Then
        GoSub 990
        GoTo 999
    End If

250 If OPCI% = 25 Then
        GoSub CIEMASVIE
        GoTo 999
    End If


cierrato:
    RESIARCH% = IARCH%
    For I% = 1 To CARCHI%
        If FINUX%(I%) > 0 Then
            IARCH% = I%
            GoSub ixarch
            Close #FINUX%(IARCH%): FINUX%(IARCH%) = 0
                                 RGB$(IARCH%) = ""
                                 IXC$(IARCH%) = ""
                                 LXC$(IARCH%) = ""
                                 CRG&(IARCH%) = 0
        End If
    Next I%
    IARCH% = RESIARCH%: RESIARCH% = 0
    CABIER% = 0
    'Call LIBARCH("*.*")
    Return

CIEMASVIE:
    If CABIER% < 1 Then GoTo 799
    MASVIE% = 0: ULTACC = 999999.99
    For J% = 1 To CARCHI%
        If FINUX%(J%) > 0 Then         ' esta abierto
            If ULTACX(J%) < ULTACC Then
                ULTACC = ULTACX(J%)
                MASVIE% = J%
            End If
        End If
    Next J%
    '
    If MASVIE% < 1 Then
        Call COMUNI("Proceso Interrumpido en Rutina 'FILEMGR'.\Imposible Determinar Archivo LRU.\Consulte a su Especialista en Software.")
        Call LIBARCH("*.*")
        Close: End
    End If
    '
    RESIARCH% = IARCH%
    IARCH% = MASVIE%
    GoSub ixarch
    Close #FINUX%(IARCH%): FINUX%(IARCH%) = 0
                         RGB$(IARCH%) = ""
                         IXC$(IARCH%) = ""
                         LXC$(IARCH%) = ""
    CABIER% = CABIER% - 1
    IARCH% = RESIARCH%: RESIARCH% = 0
    '
799 Return




datarch:
    '
    ARY$ = ARX$
    If Left$(ARY$, 1) = "!" Then
        ARY$ = Mid$(ARX$, 2)
    End If
    '
    vuelta% = 1
    '
805 FINU% = FILEOPEN%(LUDAT$ + ARY$ + ".RFS", 0, 128)
    Dim XYZ As String * 128
    If LOF(FINU%) <= 128 Then
        XYZ$ = String$(128, 0)
        Put #FINU%, 1, XYZ$
        Close #FINU%: FINU% = 0
        FILEKILL LUDAT$ + ARY$ + ".RFS"
        '
        If vuelta% = 1 Then
           Call EXTRACT(Archi$)
           vuelta% = 2
           GoTo 805
        End If
        '
        ARX1$ = TRIM$(REPLACAR$(ARX$, "\", "/"))
        Call COMUNI("Proceso Interrumpido en Rutina 'FILEMGR'.\  \Falta Controlador de Estructura del Archivo\'" + ARX1$ + ".DAT'.")
        FINU% = 0: LOR% = 0: REGBL$ = ""
        Call FINAL("")
        Close: End
        '
    End If
    '
    Get #FINU%, 1, XYZ$
    LOR% = Asc(Mid$(XYZ$, 49, 1))
    LOR1% = Asc(Mid$(XYZ$, 48, 1)): If LOR1% > 16 Then LOR1% = 0
    LOR% = 256 * LOR1% + LOR%
    If LOR% < 1 Then
        ARX1$ = TRIM$(REPLACAR$(ARX$, "\", "/"))
        Call COMUNI("Proceso Interrumpido en Rutina 'FILEMGR'.\  \Error en Longitud de Registro del Archivo\'" + ARX1$ + ".DAT'.")
        FINU% = 0: LOR% = 0: REGBL$ = ""
        Close: End
    End If
    '
    TITU$ = Left$(XYZ$, 47)
    RBL$ = String$(LOR%, 0)
    IXX$ = "": LXX$ = ""
    '
    For I% = 1 To 32
       If I% + 1 <= LOF(FINU%) / 128 Then
           Get #FINU%, I% + 1, XYZ$
           If XYZ$ <> String$(128, 0) Then
                CAFI% = I%
                TV% = Asc(Mid$(XYZ$, 33)): TX%(I%) = TV%
                PS% = Asc(Mid$(XYZ$, 35)): PX%(I%) = PS%
                LG% = Asc(Mid$(XYZ$, 36)): LXX$ = LXX$ + Chr$(LG%)
                XC% = Asc(Mid$(XYZ$, 37)): If XC% <> 1 Then XC% = 0
                If Left$(ARX$, 1) <> "!" Then
                    IXX$ = IXX$ + Chr$(XC%)
                End If
                If TV% = 6 Then
                   If PS% > 0 Then
                      If LG% > 0 Then
                         If PS% + LG% <= LOR% + 1 Then
                            Call REPLA(RBL$, Space$(LG%), PS%, LG%)
                         End If
                      End If
                   End If
                End If
           End If
       End If
    Next I%
    '
    Close #FINU%: FINU% = 0
    '
    Return
    '
ixarch:
    '
    If Left$(AX$(IARCH%), 1) = "!" Then GoTo 590
    '
    On Error Resume Next
    TENTRAIX = Timer
    On Error GoTo 0
    SCMPANT = Screen.MousePointer
    '
    For CAM% = 1 To Len(IXC$(IARCH%))
       If Asc(Mid$(IXC$(IARCH%), CAM%, 1)) = 1 Then
           ARX$ = AX$(IARCH%)
           ' FINU% = FreeFile
           'If FINU% > 32 Then
           '    GoSub cierrato
           '    Call MENSERR(24, "Proceso Interrumpido - Demasiados Archivos Abiertos")
           '    FINU% = 0: LOR% = 0: REGBL$ = ""
           '    Close: End
           'End If
           '
           ARY$ = ARX$
           If Left$(ARY$, 1) = "!" Then
              ARY$ = Mid$(ARX$, 2)
           End If
           '
           LDT$ = LUDAT$
           FINU% = FILEOPEN%(LDT$ + ARY$ + ".RFS", 0, 128)
           XYZ$ = LEEREG$(FINU%, 128, CAM% + 1)
           PS% = Asc(Mid$(XYZ$, 35))
           LG% = Asc(Mid$(XYZ$, 36))
           Close #FINU%: FINU% = 0
           '
           EXTX$ = TRIM$(Str$(CAM%)): If Len(EXTX$) < 2 Then EXTX$ = "0" + EXTX$
           EXTX$ = ".X" + EXTX$
           ARZ1$ = ARX$: GoSub 915
           '
505        'FINU% = FreeFile
           'Open ARZ2$ + EXTX$ For Random Access Read Write Shared As #FINU% Len = LG%
           FINU% = FILEOPEN%(ARZ2$ + EXTX$, 0, LG%) 'AQUI
           '
           INIII& = PRIMOD&(IARCH%)
           If INIII& > LOF(FINU%) / LG% Then INIII& = LOF(FINU%) / LG%
           If INIII& < 1 Then INIII& = 1
           '
           For U& = INIII& - 128 To INIII& - 1
              If U& > 0 Then
                 RGLS$(IARCH%) = LEEREG$(FINUX%(IARCH%), LRG%(IARCH%), U&)
                 IDX$ = LEEREG$(FINU%, LG%, U&)
                 If Mid$(RGLS$(IARCH%), PS%, LG%) <> IDX$ Then
                    INIII& = 1
                    GoTo 510
                 End If
              End If
           Next U&
           '
510        For III& = INIII& To CRG&(IARCH%)
                If Timer > TENTRAIX + 1 Then
                   Screen.MousePointer = 11
                   Call TRACE(20, III& - INIII& + 1, CRG&(IARCH%) - INIII& + 1)
                End If
                RGLS$(IARCH%) = LEEREG$(FINUX%(IARCH%), LRG%(IARCH%), III&)
                IDY$ = Mid$(RGLS$(IARCH%), PS%, LG%)
                Call GRABAREG(FINU%, LG%, IDY$, III&)

           Next III&
           For III& = CRG&(IARCH%) + 1 To LOF(FINU%) / LG%
                Call GRABAREG(FINU%, LG%, String$(LG%, 0), III&)
           Next III&
           Close #FINU%: FINU% = 0
           'PRIMOD&(IARCH%) = CRG&(IARCH%)
           
           '
       End If
    Next CAM%
    PRIMOD&(IARCH%) = CRG&(IARCH%)
    '
590 Screen.MousePointer = SCMPANT
    Return

910 '***************************************************  arma nombre completo
    '
    ARY2$ = TRIM$(ARY1$)
    If Left$(ARY1$, 1) = "!" Then
          ARY2$ = LUCOM$ + Mid$(ARY2$, 2) + ".WKF"
       Else
          If Left$(ARY2$, 1) <> "\" Then
              If Mid$(ARY2$, 2, 2) <> ":\" Then
                  ARY2$ = LUDAT$ + ARY2$ + ".DAT"
              End If
          End If
    End If
    '
    Return
    '
    '
915 ARZ2$ = TRIM$(ARZ1$)
    If Left$(ARZ1$, 1) = "!" Then
          ARZ2$ = LUCOM$ + Mid$(ARZ2$, 2)
       Else
          If Left$(ARZ2$, 1) <> "\" Then
              If Mid$(ARZ2$, 2, 2) <> ":\" Then
                  ARZ2$ = LUDAT$ + ARZ2$
              End If
          End If
    End If
    '
    Return
    '
920 '***********************************************  determina si el archivo
    '                                                ya estuvo abierto
    ARX$ = TRIM$(Archi$)
    NUEVO% = 0
    '
    For I% = 1 To CARCHI%
        If AX$(I%) = ARX$ Then
            IARCH% = I%
            GoTo 929
        End If
    Next I%
    NUEVO% = 1
    '
    If CARCHI% >= MX% Then
        Call COMUNI("Proceso Interrumpido en Rutina 'FILEMGR'.\  \Demasiados Archivos en Acceso Directo.\Consulte a su Especialista en Software.")
        Close: End
    End If
    '
    GoSub datarch                                  ' lee datos de archivo
    '
    CARCHI% = CARCHI% + 1: AX$(CARCHI%) = ARX$
                           FINUX%(CARCHI%) = 0           ' todavia no esta abierto
                           LRG%(CARCHI%) = LOR%
                           CRG&(CARCHI%) = 0                ' todavia no esta abierto
                           RGL&(CARCHI%) = 0
                           RGB$(CARCHI%) = RBL$

                           IXC$(CARCHI%) = IXX$
                           LXC$(CARCHI%) = LXX$
                           PRIMOD&(CARCHI%) = 0
                           ULTACX(CARCHI%) = Timer
                           '
    IARCH% = CARCHI%                               ' devuelve el numero de orden
    '
929 Return
    '
940 '***********************************************  abre efectivamente
    '                                                 el archivo
    If CABIER% >= MAXABIER% Then
        GoSub CIEMASVIE
    End If
    '
    If NUEVO% = 0 Then
         ARX$ = AX$(IARCH%)
         GoSub datarch                                  ' lee datos de archivo
         RGB$(IARCH%) = RBL$
         IXC$(IARCH%) = IXX$
         LXC$(IARCH%) = LXX$
    End If
    '
    'FINU% = FreeFile
    ARY1$ = AX$(IARCH%): GoSub 910
    FINU% = FILEOPEN%(ARY2$, 0, LRG%(IARCH%)) 'AQUI
    'Open ARY2$ For Random Access Read Write Shared As #FINU% Len = LRG%(IARCH%)
    '
    FINUX%(IARCH%) = FINU%
    CABIER% = CABIER% + 1
    '
    Return
    '
950 '***********************************************  reconfirma fin de archivo
    '
    ARY1$ = AX$(IARCH%): GoSub 910
    VUERABRE% = 0
    '
952 If FileLen(ARY2$) > LOF(FINUX%(IARCH%)) Then
       SMA = Screen.MousePointer
       Screen.MousePointer = 11
       If CRG&(IARCH%) > 0 Then ' fuerza posicion al principio del archivo
          XXXY$ = LEEREG$(FINUX%(IARCH%), LRG%(IARCH%), 1)
       End If
       Close #FINUX%(IARCH%)
       '
       If VUERABRE% > 3 Then
          FINU% = FILEOPEN%(ARY2$, 0, LRG%(IARCH%)) 'RE-ABRE EL ARCHIVO
          FINUX%(IARCH%) = FINU%
          LODISCO& = FileLen(ARY2$) / LRG%(IARCH%)
          LOBUFER& = LOF(FINUX%(IARCH%)) / LRG%(IARCH%)
          For UIUL& = LOBUFER& + 1 To LODISCO&
             XXXY$ = LEEREG$(FINUX%(IARCH%), LRG%(IARCH%), UIUL&)
             Call GRABAREG(FINUX%(IARCH%), LRG%(IARCH%), XXXY$, UIUL&)
          Next UIUL&
          GoTo 953
       End If
       '
       TINX = Timer: While Timer < TINX + 0.5: Wend
       FINU% = FILEOPEN%(ARY2$, 0, LRG%(IARCH%)) 'RE-ABRE EL ARCHIVO
       FINUX%(IARCH%) = FINU%
       VUERABRE% = VUERABRE% + 1
       GoTo 952
       '
    End If
    '
953 Screen.MousePointer = SMA
    'AAA = FileLen(ARY2$)
    'BBB = LOF(FINUX%(IARCH%))
    'If AAA <> BBB Then
    '      CCC = BBB
    'End If
    CARETO& = LOF(FINUX%(IARCH%)) / LRG%(IARCH%)
    CRG&(IARCH%) = CARETO&
    While CRG&(IARCH%) > 0
       RGLX$ = LEEREG$(FINUX%(IARCH%), LRG%(IARCH%), CRG&(IARCH%))
       If RGLX$ <> String$(LRG%(IARCH%), 0) Then GoTo 955
       CRG&(IARCH%) = CRG&(IARCH%) - 1
    Wend
    '
955 If RGL&(IARCH%) > 0 Then
       If RGL&(IARCH%) <= CRG&(IARCH%) Then
          RGLX$ = LEEREG$(FINUX%(IARCH%), LRG%(IARCH%), RGL&(IARCH%))
        End If
    End If
    '
    ULTACX(IARCH%) = Timer
    If PRIMOD&(IARCH%) = 0 Then PRIMOD&(IARCH%) = CRG&(IARCH%)
    '
959 Return
    '
    '
990 '***********************************************  imprime situacion general
    '                                                de los archivos
    For I% = 1 To CARCHI%
      ' LPRINT USING "##   ";I%;
      ' LPRINT LEFT$(AX$(I%)+SPACE$(30),30);
      ' LPRINT USING "######";LRG%(I%);
      ' LPRINT USING "########";CRG&(I%);
      ' LPRINT USING "########";RGL&(I%);
      ' LPRINT USING "######.##";ULTACX(I%);
      ' LPRINT USING "####";FINUX%(I%)
    Next I%
    'LPRINT String$(80, 45)
    'LPRINT
    '
    Return
    '
    '
999 'MsgBox "SALGO DE FILEMGR " + ARCHI$ + Str$(REG&) + Str$(IARCH%) + Str$(CRG&(IARCH%))
    End Sub
' ***************************************************************************


Function FILENBR%(Archi$)
'
Call FILEMGR(Archi$, FINBR%, LOREG%, CAREG&, REG&, REGS$, 1)
FILENBR% = FINBR%
End Function
' ***************************************************************************
Function TITUARCH$(Archi$)
'
Call FILEMGR(Archi$, FINBR%, LOREG%, CAREG&, REG&, REGS$, 15)
TITUARCH$ = REGS$
End Function

Function RECLEN%(Archi$)
'
Call FILEMGR(Archi$, FINBR%, LOREG%, CAREG&, REG&, REGS$, 1)
RECLEN% = LOREG%
End Function
' ***************************************************************************


Function ULTREG&(Archi$)
Call FILEMGR(Archi$, FINBR%, LOREG%, CAREG&, REG&, REGS$, 1)
ULTREG& = CAREG&
End Function
' ***************************************************************************


Function REGBLAN$(Archi$)
'
Call FILEMGR(Archi$, FINBR%, LOREG%, CAREG&, REG&, REGS$, 2)
REGBLAN$ = REGS$
End Function
' ***************************************************************************


Function REGACT$(Archi$)
'
Call FILEMGR(Archi$, FINBR%, LOREG%, CAREG&, REG&, REGS$, 3)
REGACT$ = REGS$
End Function
' ***************************************************************************


Function REGNUM&(Archi$)
'
Call FILEMGR(Archi$, FINBR%, LOREG%, CAREG&, REG&, REGS$, 3)
REGNUM& = REG&
End Function
' ***************************************************************************


Function NEXTREGS$(Archi$)
'
Call FILEMGR(Archi$, FINBR%, LOREG%, CAREG&, REG&, REGS$, 4)
NEXTREGS$ = REGS$
End Function
' ***************************************************************************


Sub REGAPP(Archi$, REGS$)
'
If Left$(Archi$, 1) <> "!" Then
   If DERACTU% < 1 Then
      If InStr(Archi$, "MENSAJES") < 1 Then
         TMENSA$ = "Usuario sin Derecho de Escritura"
         TMENSA$ = TMENSA$ + "\en la Presente Aplicación.\  "
         TMENSA$ = TMENSA$ + "\Usuario: " + USERNAME$ + "  -  Aplicación: " + TRIM$(App.EXEName) + "\  "
         TMENSA$ = TMENSA$ + "\Archivo: " + TRIM$(Archi$) + "\  "
         TMENSA$ = TMENSA$ + "\Solicite la Habilitación Correspondiente"
         TMENSA$ = TMENSA$ + "\al Supervisor del Sistema."
         Call COMUNI(TMENSA$)
         Call FINAL("")
      End If
   End If
End If
'
Call BLOQARCH(Archi$)
Call FILEMGR(Archi$, FINBR%, LOREG%, CAREG&, REG&, REGS$, 5)
End Sub
' ***************************************************************************


Sub BLANARCH(Archi$)
'
Call BLOQARCH(Archi$)
Call FILEMGR(Archi$, FINBR%, LOREG%, CAREG&, REG&, REGS$, 6)
End Sub
' ***************************************************************************


Sub CIERRARCH(Archi$)
'
SCMP = Screen.MousePointer
Screen.MousePointer = 11
Call FILEMGR(Archi$, FINBR%, LOREG%, CAREG&, REG&, REGS$, 7)
If Left$(Archi$, 1) <> "*" Then
   Call LIBARCH(Archi$)
End If
Screen.MousePointer = SCMP
End Sub
' ***************************************************************************


Sub SETREG(Archi$, REG&)
'
Call FILEMGR(Archi$, FINBR%, LOREG%, CAREG&, REG&, REGS$, 8)
End Sub
' ***************************************************************************


Function REGLEIDO$(Archi$, REG&)
'
Call FILEMGR(Archi$, FINBR%, LOREG%, CAREG&, REG&, REGS$, 9)
REGLEIDO$ = REGS$
End Function
' ***************************************************************************


Sub GRAREG(Archi$, REGS$, REG&)
'
If Left$(Archi$, 1) <> "!" Then
   If DERACTU% < 1 Then
      If InStr(Archi$, "MENSAJES") < 1 Then
         TMENSA$ = "Usuario sin Derecho de Escritura"
         TMENSA$ = TMENSA$ + "\en la Presente Aplicación.\  "
         TMENSA$ = TMENSA$ + "\Usuario: " + USERNAME$ + "  -  Aplicación: " + TRIM$(App.EXEName)
         TMENSA$ = TMENSA$ + "\Archivo: " + TRIM$(Archi$) + "\  "
         TMENSA$ = TMENSA$ + "\Solicite la Habilitación Correspondiente"
         TMENSA$ = TMENSA$ + "\al Supervisor del Sistema."
         Call COMUNI(TMENSA$)
         Call FINAL("")
      End If
   End If
End If
'
Call FILEMGR(Archi$, FINBR%, LOREG%, CAREG&, REG&, REGS$, 10)
End Sub
' ***************************************************************************

Sub ACRECORD(Archi$, REG&, COLU%, TXT$)
'
XXX$ = REGLEIDO$(Archi$, REG&)
Call REPLA(XXX$, TXT$, COLU%, Len(TXT$))
Call GRAREG(Archi$, XXX$, REG&)
End Sub
' ***************************************************************************

Sub SETULTREG(Archi$, REG&)
'
LOREG% = RECLEN%(Archi$)
REGB$ = String$(LOREG%, 0)
For I& = REG& + 1 To ULTREG&(Archi$)
    Call GRAREG(Archi$, REGB$, I&)
Next I&
Call CIERRARCH(Archi$)
'
End Sub
' ***************************************************************************

Function LENFIELD%(Archi$, X%)
Static LF%, ARCHANT$, XA%
'
If ARCHANT$ <> Archi$ Or XA% <> X% Then
    ARCHANT$ = Archi$: XA% = X%
    REG& = X%
    Call FILEMGR(Archi$, FINBR%, LOREG%, CAREG&, REG&, REGS$, 11)
    LF% = REG&: REG& = 0
End If
'
LENFIELD% = LF%
End Function

' ***************************************************************************

Function POSFIELD%(Archi$, X%)
Static PF%, ARCHANT$, XA%
'
If ARCHANT$ <> Archi$ Or XA% <> X% Then
    ARCHANT$ = Archi$: XA% = X%
    REG& = X%
    Call FILEMGR(Archi$, FINBR%, LOREG%, CAREG&, REG&, REGS$, 12)
    PF% = REG&: REG& = 0
End If
'
POSFIELD% = PF%
End Function

' ***************************************************************************

Function TIPFIELD%(Archi$, X%)
Static TF%, ARCHANT$, XA%
'
If ARCHANT$ <> Archi$ Or XA% <> X% Then
    ARCHANT$ = Archi$: XA% = X%
    REG& = X%
    Call FILEMGR(Archi$, FINBR%, LOREG%, CAREG&, REG&, REGS$, 13)
    TF% = REG&: REG& = 0
End If
'
TIPFIELD% = TF%
End Function

' ***************************************************************************

Function IDXFIELD%(Archi$, X%)
Static XF%, ARCHANT$, XA%
'
If ARCHANT$ <> Archi$ Or XA% <> X% Then
    ARCHANT$ = Archi$: XA% = X%
    REG& = X%
    Call FILEMGR(Archi$, FINBR%, LOREG%, CAREG&, REG&, REGS$, 14)
    XF% = REG&: REG& = 0
End If
'
IDXFIELD% = XF%
End Function

' ***************************************************************************

Function ECOARCH$(Archi$, ARGU$)
Static ARCHANT$, XXX$, FINEX%, LOFINEX&, IX%, LX%, LY%
'
If TRIM$(UCase$(Archi$)) = "MASTER" Then
    ECOARCH$ = DESCRIT$(CVI(ARGU$))
    Exit Function
End If
'
If Left$(UCase$(Archi$), 5) = "LIPRE" Then
   If UCase$(App.EXEName) <> "ADMIPED" Then
     ECOARCH$ = DESCRIT0$(CVI(ARGU$))
     Exit Function
   End If
End If
'
If Len(ARGU$) = 0 Then GoTo 99
'
ARCHJ$ = TRIM$(Archi$)
If Left$(ARCHJ$, 1) = "*" Then
      FRXX$ = Mid$(ARCHJ$, 2) + ".PRF"
      N% = FILEOPEN%(FRXX$, 0, 128)
      Dim YYY As String * 128
      Get #N%, 1, YYY$
      ARCHJ$ = TRIM$(Mid$(YYY$, 57, 8))
      Close #N%: N% = 0
    ElseIf Left$(ARCHJ$, 1) = "^" Then
      ARCHJ$ = Mid$(ARCHJ$, 2)
End If
'
OTRO% = 0
If ARCHJ$ <> ARCHANT$ Then
    OTRO% = 1
    ARCHANT$ = ARCHJ$
    IX% = IDXFIELD%(ARCHJ$, 1)
    LX% = LENFIELD%(ARCHJ$, 1)
    LY% = LENFIELD%(ARCHJ$, 2)
End If
'
AZ$ = Left$(ARGU$ + Space$(LX%), LX%)
'
If IX% = 1 Then
   If LX% = 2 Then        ' SOLO SI LONGITUD DEL FILTRO =2
     If OTRO% = 1 Then
        If FINEX% > 0 Then
          On Error Resume Next
          LOFINEY& = LOF(FINEX%)
          On Error GoTo 0
          If LOFINEY& <> LOFINEX& Then FINEX% = 0: OTRO% = 1
          If FINEX% > 0 Then Close #FINEX%: FINEX% = 0
        End If
        UREPRIN& = ULTREG&(ARCHANT$)
1004    LU$ = LUDAT$
1005    FINEX% = FILEOPEN%(LU$ + ARCHANT$ + ".X01", 0, 512)
        LOFINEX& = LOF(FINEX%)
        If LOF(FINEX%) / 2 < UREPRIN& Then
            Close #FINEX%: FINEX% = 0
            XYZT$ = REGLEIDO$(ARCHANT$, 1)
            Call GRAREG(ARCHANT$, XYZT$, 1)
            Call CIERRARCH(ARCHANT$)
            GoTo 1005
        End If
        '
     End If
     '
     Dim XXXX As String * 512
     On Error GoTo 1006
     LOFI& = Int(LOF(FINEX%) / 512)
     On Error GoTo 0
     GoTo 1009
     '
1006 Resume 1007
1007 On Error GoTo 0
     GoTo 1005
     '
1009 For RGX& = 1 To 1 + Int(LOF(FINEX%) / 512)
          On Error GoTo 1016
          Get #FINEX%, RGX&, XXXX$
          On Error GoTo 0
          PINI = 1
1010      POSI% = InStr(PINI, XXXX$, ARGU$)
          If POSI% < 1 Then GoTo 1015                           ' NEXT RGX&
          If (POSI% Mod 2) <> 1 Then PINI = POSI% + 1: GoTo 1010
          PINI = POSI% + 2
          RGL& = 256 * (RGX& - 1) + (POSI% + 1) / 2
          Call SETREG(Archi$, RGL&)
          GoTo 1020
1015 Next RGX&
     '
     Call SETREG(ARCHJ$, 0&)
     ECOARCH$ = ""
     Exit Function
   End If
End If
GoTo 1020
'
1016 Resume 1017
1017 On Error GoTo 0
     GoTo 1004
     '
1020 If Left$(REGACT$(ARCHJ$), LX%) = AZ$ Then
         ECOARCH$ = Mid$(REGACT$(ARCHJ$), 1 + LX%, LY%)
         Exit Function
     End If
     '
     For RGL& = ULTREG&(ARCHJ$) To 1 Step -1
        RGLXX$ = REGLEIDO$(ARCHJ$, RGL&)
        If Left$(RGLXX$, LX%) = AZ$ Then
           ECOARCH$ = Mid$(RGLXX$, 1 + LX%, LY%)
           Exit Function
        End If
     Next RGL&
     '
99   Call SETREG(ARCHJ$, 0&)
     ECOARCH$ = ""
     '
End Function
'
' ***************************************************************************

Function RECFILT$(Archi$, X%, ARGU$)
'Static ARCHA$, LX%, FINEX%, XA%
'
REG& = REGNUM&(Archi$)
RX$ = ""
'
ARCHJ$ = Archi$
IX% = IDXFIELD%(ARCHJ$, X%)
If IX% = 1 Then
    OTRO% = 0
    If ARCHJ$ <> ARCHANT$ Or X% <> XA% Then
        OTRO% = 1
        ARCHANT$ = ARCHJ$
        XA% = X%
        LX% = LENFIELD%(ARCHJ$, X%)
    End If
    '
    AZ$ = Left$(ARGU$ + Space$(LX%), LX%)
    '
    If LX% = 2 Then        ' SOLO SI LONGITUD DEL FILTRO =2
1003    If OTRO% = 1 Then
            If FINEX% > 0 Then
                Close #FINEX%: FINEX% = 0
            End If
            LU$ = LUDAT$
1005        FINEX% = FreeFile
            EXTX$ = TRIM$(Str$(X%))
            While Len(EXTX$) < 2: EXTX$ = "0" + EXTX$: Wend
            FINEX% = FILEOPEN%(LU$ + ARCHANT$ + ".X" + EXTX$, 0, 512)
            If LOF(FINEX%) / 2 < ULTREG&(Archi$) Then
                REAC1& = LOF(FINEX%) / 2 - 64: If REAC1& < 1 Then REAC1& = 1
                If REAC1& <= ULTREG&(Archi$) Then
                   Close #FINEX%: FINEX% = 0
                   XYZT$ = REGLEIDO$(Archi$, REAC1&)
                   Call GRAREG(Archi$, XYZT$, REAC1&)
                   Call CIERRARCH(Archi$)
                   GoTo 1005
                End If
            End If
            '
        End If
        '
        Dim XXXX As String * 512
        On Error GoTo 1020
        For RGX& = 1 To 1 + Int(LOF(FINEX%) / 512)
            Get #FINEX%, RGX&, XXXX$
            On Error GoTo 0
            PINI = 1
1010        POSI% = InStr(PINI, XXXX$, ARGU$)
            If POSI% < 1 Then GoTo 1015                           ' NEXT RGX&
            If (POSI% Mod 2) <> 1 Then PINI = POSI% + 1: GoTo 1010
            PINI = POSI% + 2
            RGL& = 256 * (RGX& - 1) + (POSI% + 1) / 2
            RX$ = RX$ + MKS$(RGL&)
            GoTo 1010
1015    Next RGX&
        On Error GoTo 0
        GoTo 1900
        '
1020    OTRO% = 1
        FINEX% = 0
        Resume 1003
        '
    End If
End If
'
PX% = POSFIELD%(Archi$, X%)
LX% = LENFIELD%(Archi$, X%)
AZ$ = Left$(ARGU$ + Space$(LX%), LX%)
'
If PX% > 0 Then
    If LX% > 0 Then
        For RGL& = ULTREG&(Archi$) To 1 Step -1
           If Mid$(REGLEIDO$(Archi$, RGL&), PX%, LX%) = AZ$ Then
               RX$ = MKS$(RGL&) + RX$
           End If
        Next RGL&
    End If
End If
'
1900 Call SETREG(Archi$, REG&)
     RECFILT$ = RX$
     '
If FINEX% > 0 Then
   Close #FINEX%
   FINEX% = 0
End If
'
End Function
' ***************************************************************************

Function REGBUS&(Archi$, X%, ARGU$)
'
PX% = POSFIELD%(Archi$, X%)
LX% = LENFIELD%(Archi$, X%)
AZ$ = Left$(ARGU$ + Space$(LX%), LX%)
REG& = 0
'
If X% = 1 Then
    ECO$ = ECOARCH$(Archi$, ARGU$)
    RGL& = REGNUM&(Archi$)
    '
    If Mid$(REGLEIDO$(Archi$, RGL&), PX%, LX%) = AZ$ Then
        REGBUS& = RGL&
        Exit Function
    End If
End If
'
If PX% > 0 Then
    If LX% > 0 Then
        For RGL& = 1 To ULTREG&(Archi$)
           If Mid$(REGLEIDO$(Archi$, RGL&), PX%, LX%) = AZ$ Then
               '
               REGBUS& = RGL&
               Exit Function
               '
           End If
        Next RGL&
    End If
End If
'
Call SETREG(Archi$, 0&)
REGBUS& = 0
'
End Function
' ***************************************************************************

Sub SITUARCH()
'
Call FILEMGR(Archi$, FINBR%, LOREG%, CAREG&, REG&, REGS$, 24)
End Sub
'

Sub CIEMASVIE()
'
Call FILEMGR(Archi$, FINBR%, LOREG%, CAREG&, REG&, REGS$, 25)
End Sub
'

Function LEEREG$(FINU%, LOREG%, RECNO&)
'
Static BU1 As String * 1
Static BU2 As String * 2
Static BU4 As String * 4
Static BU6 As String * 6
Static BU8 As String * 8
Static BU10 As String * 10
Static BU12 As String * 12
Static BU16 As String * 16
Static BU18 As String * 18
Static BU20 As String * 20
Static BU24 As String * 24
Static BU32 As String * 32
Static BU36 As String * 36
Static BU38 As String * 38
Static BU40 As String * 40
Static BU42 As String * 42
Static BU48 As String * 48
Static BU56 As String * 56
Static BU64 As String * 64
Static BU72 As String * 72
Static BU80 As String * 80
Static BU96 As String * 96
Static BU128 As String * 128
Static BU160 As String * 160
Static BU256 As String * 256
Static BU300 As String * 300
Static BU304 As String * 304
Static BU512 As String * 512
Static BU1024 As String * 1024
'
LEEREG$ = ""
On Error GoTo 198
'
Select Case LOREG%
Case 1
   Get #FINU%, RECNO&, BU1
   LEEREG$ = BU1
Case 2
   Get #FINU%, RECNO&, BU2
   LEEREG$ = BU2
Case 4
   Get #FINU%, RECNO&, BU4
   LEEREG$ = BU4
Case 6
   Get #FINU%, RECNO&, BU6
   LEEREG$ = BU6
Case 8
   Get #FINU%, RECNO&, BU8
   LEEREG$ = BU8
Case 10
   Get #FINU%, RECNO&, BU10
   LEEREG$ = BU10
Case 12
   Get #FINU%, RECNO&, BU12
   LEEREG$ = BU12
Case 16
   Get #FINU%, RECNO&, BU16
   LEEREG$ = BU16
Case 18
   Get #FINU%, RECNO&, BU18
   LEEREG$ = BU18
Case 20
   Get #FINU%, RECNO&, BU20
   LEEREG$ = BU20
Case 24
   Get #FINU%, RECNO&, BU24
   LEEREG$ = BU24
Case 32
   Get #FINU%, RECNO&, BU32
   LEEREG$ = BU32
Case 36
   Get #FINU%, RECNO&, BU36
   LEEREG$ = BU36
Case 38
   Get #FINU%, RECNO&, BU38
   LEEREG$ = BU38
Case 40
   Get #FINU%, RECNO&, BU40
   LEEREG$ = BU40
Case 42
   Get #FINU%, RECNO&, BU42
   LEEREG$ = BU42
Case 48
   Get #FINU%, RECNO&, BU48
   LEEREG$ = BU48
Case 56
   Get #FINU%, RECNO&, BU56
   LEEREG$ = BU56
Case 64
   Get #FINU%, RECNO&, BU64
   LEEREG$ = BU64
Case 72
   Get #FINU%, RECNO&, BU72
   LEEREG$ = BU72
Case 80
   Get #FINU%, RECNO&, BU80
   LEEREG$ = BU80
Case 96
   Get #FINU%, RECNO&, BU96
   LEEREG$ = BU96
Case 128
   Get #FINU%, RECNO&, BU128
   LEEREG$ = BU128
Case 160
   Get #FINU%, RECNO&, BU160
   LEEREG$ = BU160
Case 256
   Get #FINU%, RECNO&, BU256
   LEEREG$ = BU256
Case 300
   Get #FINU%, RECNO&, BU300
   LEEREG$ = BU300
Case 304
   Get #FINU%, RECNO&, BU304
   LEEREG$ = BU304
Case 512
   Get #FINU%, RECNO&, BU512
   LEEREG$ = BU512
Case 1024
   Get #FINU%, RECNO&, BU1024
   LEEREG$ = BU1024
Case Else
   Call MENSERR(24, "No  hay  Controlador  de  Acceso  a\Archivo para Longitud de Registro =" + Str$(LOREG%))
   Close: End
End Select
'
On Error GoTo 0
Exit Function
'
198 ERNU$ = TRIM$(Str$(Err.Number))
    MENSA$ = "Se ha Producido un Error (" + ERNU$ + ") en la LECTURA\"
    MENSA$ = MENSA$ + "ALEATORIA del Archivo Numero " + TRIM$(Str$(FINU%)) + ".\  \"
    '
    If FINU% < 1 Or FINU% > 512 Then
         MENSA$ = MENSA$ + "ERROR: Número Incorrecto de Apertura de Archivo."
       ElseIf RECNO& < 1 Then
         MENSA$ = MENSA$ + "ERROR: Número de Registro Leído Incorrecto."
       Else
         MENSA$ = MENSA$ + "Investigue las Siguientes Causas Posibles:\  \"
         MENSA$ = MENSA$ + "   1.- Archivo Número '" + TRIM$(Str$(FINU%)) + "' no está Abierto.\"
         MENSA$ = MENSA$ + "   2.- Problemas de Red - No hay Acceso al Servidor.\"
         MENSA$ = MENSA$ + "   3.- Ruta de Acceso Inexistente.\"
         MENSA$ = MENSA$ + "   4.- Archivo de Solo Lectura u Oculto.\"
         MENSA$ = MENSA$ + "   5.- Archivo Inexistente (LECTURA SECUENCIAL).\"
         MENSA$ = MENSA$ + "   6.- Archivo no Compartido en Uso por Otra Aplicación.\"
    End If
    '
    MENSA$ = MENSA$ + "   \Consulte a su Soporte de Sistemas.\"
    '
    C$ = FORMESS$(MENSA$)
    RESPU% = MsgBox(C$, 53, "Error de Lectura de Archivo")
    If RESPU% = 4 Then
        Resume
    End If
    Resume 199
199 LEEREG$ = String$(LOREG%, 0)
On Error GoTo 0
'
End Function

Sub GRABAREG(FINU%, LOREG%, REGX$, RECNO&)
'
Static BU1 As String * 1
Static BU2 As String * 2
Static BU4 As String * 4
Static BU6 As String * 6
Static BU8 As String * 8
Static BU10 As String * 10
Static BU12 As String * 12
Static BU16 As String * 16
Static BU18 As String * 18
Static BU20 As String * 20
Static BU24 As String * 24
Static BU32 As String * 32
Static BU36 As String * 36
Static BU38 As String * 38
Static BU40 As String * 40
Static BU42 As String * 42
Static BU48 As String * 48
Static BU56 As String * 56
Static BU64 As String * 64
Static BU72 As String * 72
Static BU80 As String * 80
Static BU96 As String * 96
Static BU128 As String * 128
Static BU160 As String * 160
Static BU256 As String * 256
Static BU300 As String * 300
Static BU304 As String * 304
Static BU512 As String * 512
Static BU1024 As String * 1024
'
On Error GoTo 198
'
Select Case LOREG%
Case 1
   BU1 = REGX$
   Put #FINU%, RECNO&, BU1
Case 2
   BU2 = REGX$
   Put #FINU%, RECNO&, BU2
Case 4
   BU4 = REGX$
   Put #FINU%, RECNO&, BU4
Case 6
   BU6 = REGX$
   Put #FINU%, RECNO&, BU6
Case 8
   BU8 = REGX$
   Put #FINU%, RECNO&, BU8
Case 10
   BU10 = REGX$
   Put #FINU%, RECNO&, BU10
Case 12
   BU12 = REGX$
   Put #FINU%, RECNO&, BU12
Case 16
   BU16 = REGX$
   Put #FINU%, RECNO&, BU16
Case 18
   BU18 = REGX$
   Put #FINU%, RECNO&, BU18
Case 20
   BU20 = REGX$
   Put #FINU%, RECNO&, BU20
Case 24
   BU24 = REGX$
   Put #FINU%, RECNO&, BU24
Case 32
   BU32 = REGX$
   Put #FINU%, RECNO&, BU32
Case 36
   BU36 = REGX$
   Put #FINU%, RECNO&, BU36
Case 38
   BU38 = REGX$
   Put #FINU%, RECNO&, BU38
Case 40
   BU40 = REGX$
   Put #FINU%, RECNO&, BU40
Case 42
   BU42 = REGX$
   Put #FINU%, RECNO&, BU42
Case 48
   BU48 = REGX$
   Put #FINU%, RECNO&, BU48
Case 56
   BU56 = REGX$
   Put #FINU%, RECNO&, BU56
Case 64
   BU64 = REGX$
   Put #FINU%, RECNO&, BU64
Case 72
   BU72 = REGX$
   Put #FINU%, RECNO&, BU72
Case 80
   BU80 = REGX$
   Put #FINU%, RECNO&, BU80
Case 96
   BU96 = REGX$
   Put #FINU%, RECNO&, BU96
Case 128
   BU128 = REGX$
   Put #FINU%, RECNO&, BU128
Case 160
   BU160 = REGX$
   Put #FINU%, RECNO&, BU160
Case 256
   BU256 = REGX$
   Put #FINU%, RECNO&, BU256
Case 300
   BU300 = REGX$
   Put #FINU%, RECNO&, BU300
Case 304
   BU304 = REGX$
   Put #FINU%, RECNO&, BU304
Case 512
   BU512 = REGX$
   Put #FINU%, RECNO&, BU512
Case 1024
   BU1024 = REGX$
   Put #FINU%, RECNO&, BU1024
Case Else
   Call MENSERR(24, "No  hay  Controlador  de  Acceso a\Archivo para Longitud de Registro =" + Str$(LOREG%))
   Close: End
End Select
'
On Error GoTo 0
Exit Sub
'
198 ERNU$ = TRIM$(Str$(Err.Number))
    MENSA$ = "Se ha Producido un Error (" + ERNU$ + ") en la ESCRITURA ALEATORIA\"
    MENSA$ = MENSA$ + MODACCE$ + " de un Archivo.\  \"
    '
    If FINU% < 1 Or FINU% > 512 Then
         MENSA$ = MENSA$ + "ERROR: Número Incorrecto de Apertura de Archivo."
       ElseIf RECNO& < 1 Then
         MENSA$ = MENSA$ + "ERROR: Número de Registro Leído Incorrecto."
       Else
         MENSA$ = MENSA$ + "Investigue las Siguientes Causas Posibles:\  \"
         MENSA$ = MENSA$ + "   1.- Archivo Número '" + TRIM$(Str$(FINU%)) + "' no está Abierto.\"
         MENSA$ = MENSA$ + "   2.- Problemas de Red - No hay Acceso al Servidor.\"
         MENSA$ = MENSA$ + "   3.- Ruta de Acceso Inexistente.\"
         MENSA$ = MENSA$ + "   4.- Archivo de Solo Lectura u Oculto.\"
         MENSA$ = MENSA$ + "   5.- Archivo Inexistente (LECTURA SECUENCIAL).\"
         MENSA$ = MENSA$ + "   6.- Archivo no Compartido en Uso por Otra Aplicación.\"
    End If
    '
    MENSA$ = MENSA$ + "   \Consulte a su Soporte de Sistemas.\"
    '
    C$ = FORMESS$(MENSA$)
    RESPU% = MsgBox(C$, 53, "Error de Escritura de Archivo")
    If RESPU% = 4 Then
        Resume
    End If
    Resume 199
    '
199 On Error GoTo 0
'
End Sub
'

Function FILEOPEN%(ARXX$, MODO%, LOREGG%)
    '
    On Error GoTo 99
    FILI% = FreeFile
    Select Case MODO%
      Case 0
        MODACCE$ = "ACCESO ALEATORIO"
        Open ARXX$ For Random Access Read Write Shared As #FILI% Len = LOREGG%
        GoTo 89
      Case 1
        MODACCE$ = "LECTURA SECUENCIAL"
        Open ARXX$ For Input Shared As #FILI%
        GoTo 89
      Case 2
        MODACCE$ = "ESCRITURA SECUENCIAL"
        Open ARXX$ For Output As #FILI%
        GoTo 89
      Case 3
        MODACCO$ = "SALIDA SECUENCIAL INCREMENTAL"
        Open ARXX$ For Append As #FILI%
        GoTo 89
      Case 4
        MODACCE$ = "ACCESO BINARIO"
        Open ARXX$ For Binary Access Read Write Shared As #FILI%
        GoTo 89
      Case 9
        MODACCE$ = "ACCESO ALEATORIO SOLO LECTURA"
        Open ARXX$ For Random Access Read Shared As #FILI% Len = LOREGG%
        GoTo 89
    End Select
    FILI% = 0
    '
89  If FILI% > 0 Then
        On Error GoTo 0
        FILEOPEN% = FILI%
        Exit Function
    End If
    GoTo 100
    '
99  ERNU$ = TRIM$(Str$(Err.Number))
    'Resume 100
100 BRXX$ = TRIM$(REPLACAR$(ARXX$, "\", "/"))
    If Mid$(BRXX$, 2, 1) <> ":" Then
        BRXX$ = Left$(CurDir, 2) + BRXX$
    End If
    MENSA$ = "Se ha Producido un Error (" + ERNU$ + ") en la Apertura para\"
    MENSA$ = MENSA$ + MODACCE$ + " del Siguiente Archivo:\  \"
    MENSA$ = MENSA$ + BRXX$ + "\  \"
    '
    MENSA$ = MENSA$ + "Investigue las Siguientes Causas Posibles:\  \"
    MENSA$ = MENSA$ + "   1.- Problemas de Red - No hay Acceso a U.L. '" + Left$(BRXX$, 2) + "'.\"
    MENSA$ = MENSA$ + "   2.- Ruta de Acceso Inexistente.\"
    MENSA$ = MENSA$ + "   3.- Archivo de Solo Lectura u Oculto.\"
    MENSA$ = MENSA$ + "   4.- Archivo Inexistente (LECTURA SECUENCIAL).\"
    MENSA$ = MENSA$ + "   5.- Archivo no Compartido en Uso por Otra Aplicación.\"
    MENSA$ = MENSA$ + "   \Consulte a su Soporte de Sistemas.\"
    '
    C$ = FORMESS$(MENSA$)
    RESPU% = MsgBox(C$, 53, "Error de Apertura de Archivo")
    If RESPU% = 4 Then
        Resume
    End If
    Close: End
    '
End Function
'

Sub FILEKILL(ARKI$)
    On Error GoTo 99
    Kill ARKI$
    On Error GoTo 0
    Exit Sub
    '
99  ERNU$ = TRIM$(Str$(Err.Number))
100 BRXX$ = TRIM$(REPLACAR$(ARKI$, "\", "/"))
    If Mid$(BRXX$, 2, 1) <> ":" Then
        BRXX$ = Left$(CurDir, 2) + BRXX$
    End If
    MENSA$ = "Se ha Producido un Error (" + ERNU$ + ") en la Apertura para\"
    MENSA$ = MENSA$ + "BORRADO del Siguiente Archivo:\  \"
    MENSA$ = MENSA$ + BRXX$ + "\  \"
    '
    MENSA$ = MENSA$ + "Investigue las Siguientes Causas Posibles:\  \"
    MENSA$ = MENSA$ + "   1.- Problemas de Red - No hay Acceso a U.L. '" + Left$(BRXX$, 2) + "'.\"
    MENSA$ = MENSA$ + "   2.- Ruta de Acceso Inexistente.\"
    MENSA$ = MENSA$ + "   3.- Archivo de Solo Lectura u Oculto.\"
    MENSA$ = MENSA$ + "   4.- Archivo Inexistente.\"
    MENSA$ = MENSA$ + "   5.- Archivo en Uso por Otra Aplicación.\"
    MENSA$ = MENSA$ + "   \Consulte a su Soporte de Sistemas.\"
    '
    C$ = FORMESS$(MENSA$)
    RESPU% = MsgBox(C$, 53, "Error de Borrado de Archivo")
    If RESPU% = 4 Then
        Resume
    End If
    Close: End
    '
End Sub
'

Sub COPYSTRU(A$, B$)
        '
        LU$ = LUDAT$
        Dim ORI As String * 128
        Dim NUE As String * 128
        '
        X$ = TRIM$(A$): Y$ = TRIM$(B$)
        If X$ = "" Or Y$ = "" Or Len(X$) > 8 Or Len(Y$) > 8 Then
            Call COMUNI("Error de Invocación de Rutina 'COPYSTRU'.\  \Consulte a su Soporte de Sistemas FLEXOFT.")
            Exit Sub
        End If
        '
        If EXISTE%(LU$ + X$ + ".RFS") < 1 Then
           Call EXTRACT(X$ + ".RFS")
        End If
        '
        N1% = FILEOPEN%(LU$ + X$ + ".RFS", 0, 128)
        N2% = FILEOPEN%(LU$ + Y$ + ".RFS", 0, 128)
        '
        J% = 0
        For I% = 1 To LOF(N1%) / 128
           Get #N1%, I%, ORI$
           NXX$ = ORI$
           NUE$ = NXX$
           Put #N2%, I%, NUE$
           J% = I%
        Next I%
        '
        For I% = J% + 1 To LOF(N2%) / 128
           NUE$ = String$(128, 0)
           Put #N2%, I%, NUE$
        Next I%
        '
        If N2% > 0 Then Close #N2%: N2% = 0
        If N1% > 0 Then Close #N1%: N1% = 0
        X$ = "": Y$ = ""
        '
End Sub
'

Function COPYFILE%(ORIGI$, DESTI$)
    '
    On Error GoTo 10
    FileCopy ORIGI$, DESTI$
    CPF% = 1
    On Error GoTo 0
    GoTo 99
    '
10  ERNU$ = TRIM$(Str$(Err.Number))
    '
    ORIGI1$ = TRIM$(REPLACAR$(ORIGI$, "\", "/"))
    DESTI1$ = TRIM$(REPLACAR$(DESTI$, "\", "/"))
    '
    MENSA$ = "Se ha Producido un Error (" + ERNU$ + ") en la Copia de los\"
    MENSA$ = MENSA$ + "Siguientes Archivos:\  \"
    MENSA$ = MENSA$ + "Origen:\  " + ORIGI1$ + "\  \"
    MENSA$ = MENSA$ + "Destino:\  " + DESTI1$ + "\  \"
    '
    MENSA$ = MENSA$ + "Investigue las Siguientes Causas Posibles:\  \"
    MENSA$ = MENSA$ + "   1.- Problema de Red - No hay Acceso a U.L.'s.\"
    MENSA$ = MENSA$ + "   2.- Rutas de Acceso Inexistentes.\"
    MENSA$ = MENSA$ + "   3.- Archivo Destino de Solo Lectura u Oculto.\"
    MENSA$ = MENSA$ + "   4.- Archivo Origen Inexistente.\"
    MENSA$ = MENSA$ + "   5.- Archivo Destino en Uso por Otra Aplicación.\"
    MENSA$ = MENSA$ + "   \Consulte a su Soporte de Sistemas.\"
    '
    C$ = FORMESS$(MENSA$)
    RESPU% = MsgBox(C$, 53, "Error de Copia de Archivos")
    If RESPU% = 4 Then
        Resume
    End If
    CPF% = 0
    '
99  COPYFILE = CPF%
    '
End Function
'

Sub HEADERS(ARCHIVO$, TEXTO$)
    '
    LU$ = LUDAT$
    ARCH$ = UCase$(AJUSTI$(ARCHIVO$, 8))
    '
200 If TEXTO$ = "" Then

            For I& = 1 To ULTREG&("HEADERS")
                XXX$ = REGLEIDO$("HEADERS", I&)
                Archi$ = Left$(XXX$, 8)
                If Archi$ <> ARCH$ Then
                  If Archi$ <> Space$(8) Then
                    If Archi$ <> String$(8, 0) Then
                        J& = J& + 1
                        Call GRAREG("HEADERS", XXX$, J&)
                    End If
                  End If
                End If
            Next I&

            For I& = J& + 1 To ULTREG&("HEADERS")
                Call GRAREG("HEADERS", String$(128, 0), I&)
            Next I&
            '
            Call CIERRARCH("HEADERS")
            Exit Sub
            '
        Else
            '
            For I& = 1 To ULTREG&("HEADERS")
                XXX$ = REGLEIDO$("HEADERS", I&)
                Archi$ = Left$(XXX$, 8)
                'If ARCHI$ <> ARCH$ Then
                  If Archi$ <> Space$(8) Then
                    If Archi$ <> String$(8, 0) Then
                      J& = J& + 1
                      Call GRAREG("HEADERS", XXX$, J&)
                    End If
                  End If
                'End If
            Next I&
            '
            XXX$ = ARCH$ + TEXTO$
            Call GRAREG("HEADERS", XXX$, J& + 1)
            '
            For I& = J& + 2 To ULTREG&("HEADERS")
                Call GRAREG("HEADERS", String$(128, 0), I&)
            Next I&
            '
            Call CIERRARCH("HEADERS")
    End If
    '
End Sub
'
Sub ENCACOL(ARCHIVO$, TEXTO$)
    '
    LU$ = LUDAT$
    ARCH$ = UCase$(AJUSTI$(ARCHIVO$, 8))
    '
    NXX% = FILEOPEN%(LUCOM$ + ARCHIVO$ + ".HDP", 0, 256)
    Dim TENCA As String * 256
    '
200 If TEXTO$ = "" Then
        Close #NXX%
        On Error Resume Next
        Kill LUCOM$ + ARCHIVO$ + ".HDP"
        On Error GoTo 0
       Else
        REG& = 1 + LOF(NXX%) / 256
        TENCA$ = Left$(TEXTO$ + Space$(256), 256)
        Put #NXX%, REG&, TENCA$
        Close #NXX%
    End If
    '
End Sub
'

Function CONTROL$(PROCE$, CLAVE$)
    '
    PROBUS$ = AJUSTI$(PROCE$, 8)
    If TRIM$(PROCE$) = "" Then PROBUS$ = ""
    CODBUS$ = AJUSTI$(CLAVE$, 8)
    '
    Call CIERRARCH("FLEXCRL")
    For I& = 1 To ULTREG&("FLEXCRL")
       XI$ = REGLEIDO$("FLEXCRL", I&)
       If PROBUS$ = "" Or PROBUS$ = Left$(XI$, 8) Then
           If Mid$(XI$, 9, 8) = CODBUS$ Then
               CONTROL$ = TRIM$(Mid$(XI$, 17, 40))
               Call CIERRARCH("FLEXCRL")
               Exit Function
           End If
       End If
    Next I&
    CONTROL$ = ""
    Call CIERRARCH("FLEXCRL")
    '
End Function
'
Sub GRACONTROL(PROCE$, CLAVE$, VALCON$)
    '
    PROBUS$ = AJUSTI$(PROCE$, 8)
    If TRIM$(PROCE$) = "" Then PROBUS$ = ""
    CODBUS$ = AJUSTI$(CLAVE$, 8)
    '
    For I& = 1 To ULTREG&("FLEXCRL")
       XI$ = REGLEIDO$("FLEXCRL", I&)
       If PROBUS$ = "" Or PROBUS$ = Left$(XI$, 8) Then
           If Mid$(XI$, 9, 8) = CODBUS$ Then
               Call REPLA(XI$, VALCON$, 17, 40)
               Call GRAREG("FLEXCRL", XI$, I&)
               Call CIERRARCH("FLEXCRL")
               Exit Sub
           End If
       End If
    Next I&
    '
    XI$ = REGBLAN$("FLEXCRL")
    Call REPLA(XI$, PROBUS$, 1, 8)
    Call REPLA(XI$, CODBUS$, 9, 8)
    Call REPLA(XI$, VALCON$, 17, 40)
    Call REGAPP("FLEXCRL", XI$)
    Call CIERRARCH("FLEXCRL")
    '
End Sub
'

Sub RECUPD(STRI$, PINI%, LOCO%, VALO#)
  
10 If PINI% + LOCO% > Len(STRI$) + 1 Then
      MENSERR 24, "Longitud Incorrecta en Invocacion Rutina RECUPD"
      GoTo 10000
      End If

30 If LOCO% = 1 Then
      SUMA% = (Asc(Mid$(STRI$, PINI%) + String$(1, 0)) + Int(VALO# + 0.5)) Mod 256
      If SUMA% < 0 Then
         MENSERR 24, "Suma Negativa en Rutina RECUPD"
         GoTo 10000
         End If
      Mid$(STRI$, PINI%, 1) = Chr$(SUMA%)
      GoTo 950
      End If

50 If LOCO% = 2 Then
      SUMA% = CVI(Mid$(STRI$, PINI%) + String$(2, 0))
      If Abs(VALO# + SUMA%) > 32767 Then
         MENSERR 24, "Suma Fuera de Rango Entero en Rutina RECUPD"
         GoTo 10000
      End If
      SUMA% = SUMA% + Int(VALO# + 0.5)
      Mid$(STRI$, PINI%, 2) = MKI$(SUMA%)
      GoTo 950
      End If

70 If LOCO% = 4 Then
      SUMAS = CVS(Mid$(STRI$, PINI%) + String$(4, 0))
      SUMAS = SUMAS + VALO#
      Mid$(STRI$, PINI%, 4) = MKS$(SUMAS)
      GoTo 950
      End If

90 If LOCO% = 8 Then
      SUMAD# = CVD(Mid$(STRI$, PINI%) + String$(8, 0))
      SUMAD# = SUMAD# + VALO#
      Mid$(STRI$, PINI%, 8) = MKD$(SUMAD#)
      GoTo 950
      End If

      MENSERR 24, "Longitud Incorrecta en Invocacion de Rutina RECUPD"

950 GoTo 10000

10000 End Sub
'
Function EXISTE%(A$)
    '
    ARXX$ = TRIM$(A$)
    If InStr(ARXX$, ".") < 1 Then
       ARXX$ = ARXX$ + ".DAT"
    End If
    '
    EXI% = 0
    On Error GoTo 89
    If FileLen(ARXX$) > 0 Then
        EXI% = 1
    End If
    GoTo 99
    '
89  Resume 99
    '
99  On Error GoTo 0
    EXISTE% = EXI%
    '
End Function
'
Sub EXTRACT(FINAME$)   ' vigente desde 10/03/2002
   '
   AREX$ = TRIM$(UCase$(FINAME$))
   If InStr(AREX$, ".") < 2 Then
     AREX$ = AREX$ + ".RFS"
   End If
   '
   If Left$(AREX$, 1) = "!" Then
     AREX$ = Mid$(AREX$, 2)
   End If
   '
   If InStr(AREX$, ".PRF") > 0 Then
29   POSI% = InStr(AREX$, "\")
     If POSI% > 0 Then
       AREX$ = Mid$(AREX$, POSI% + 1)
       GoTo 29
     End If
   End If
   '
   If EXTRACTFORM.Visible = True Then
      EXTRACTFORM.Hide
      Call FRESHALL
   End If
   EXTRACTFORM.Label18.Caption = AREX$
   On Error Resume Next
   EXTRACTFORM.Show 1
   On Error GoTo 0
   Call FRESHALL
   '
End Sub
'
Sub EXTRACTVIE(ARXX$)
    '
    BRXX$ = TRIM$(UCase$(ARXX$))
    If InStr(BRXX$, ".PRF") > 0 Then GoTo 199
    '
    If Left$(BRXX$, 1) = "!" Then
        BRXX$ = Mid$(BRXX$, 2)
    End If
    '
    If EXISTE%(LUBAS$ + "FLEXRFS.ZIP") = 1 Then
       If EXISTE%(LUBAS$ + "PKUNZIP.EXE") = 1 Then
          On Error Resume Next
          Kill "C:\FLEXOFT\FINBAT.$$$"
          NX% = FILEOPEN%("C:\FLEXOFT\COMBAT.$$$", 2, 0)
          Print #NX%, "FINBAT"
          Close #NX%: NX% = 0
          '
          LDT$ = TRIM$(UCase$(LUDAT$))
          NX% = FILEOPEN%("C:\FLEXOFT\DOSCOMM.BAT", 2, 0)
          Print #NX%, LUBAS$ + "PKUNZIP -O " + LUBAS$ + "FLEXRFS " + TRIM$(BRXX$) + ".RFS"
          Print #NX%, "COPY C:\FLEXOFT\COMBAT.$$$ C:\FLEXOFT\FINBAT.$$$"
          Close #NX%: NX% = 0
          '
          Shell "C:\FLEXOFT\DOSCOMM.PIF"
          TINI = Timer
          MENSA$ = "No se Encontró el Controlador Adecuado" + Chr$(10) + Chr$(13)
          MENSA$ = MENSA$ + "para el Acceso a esta Base de Datos." + Chr$(10) + Chr$(13) + Chr$(10) + Chr$(13)
          MENSA$ = MENSA$ + "FLEXOFT Realizará Ahora la Extracción" + Chr$(10) + Chr$(13)
          MENSA$ = MENSA$ + "Correspondiente." + Chr$(10) + Chr$(13) + Chr$(10) + Chr$(13)
          MENSA$ = MENSA$ + "Controlador: " + TRIM$(BRXX$) + ".RFS"
          '
149       If Timer > TINI + 30 Then Exit Sub
          If EXISTE%("C:\FLEXOFT\FINBAT.$$$") < 1 Then GoTo 149
          MsgBox MENSA$, 4096 + 48, "Controladores de Proceso"
          Call FRESHALL
          FILEKILL "C:\FLEXOFT\COMBAT.$$$"
          FILEKILL "C:\FLEXOFT\FINBAT.$$$"
          If LDT$ <> "C:\FLEXOFT\" Then
              On Error Resume Next
              FileCopy "C:\FLEXOFT\" + TRIM$(BRXX$) + ".RFS ", LDT$ + TRIM$(BRXX$) + ".RFS"
              Kill "C:\FLEXOFT\" + TRIM$(BRXX$) + ".RFS"
          End If
          On Error GoTo 0
          '
       End If
    End If
    Exit Sub
    '
199 CRXX$ = BRXX$
299 POSI% = InStr(CRXX$, "\")
    If POSI% > 0 Then
        CRXX$ = Mid$(CRXX$, POSI% + 1)
        GoTo 299
    End If
    '
    If EXISTE%(LUBAS$ + "FLEXPRF.ZIP") = 1 Then
       If EXISTE%(LUBAS$ + "PKUNZIP.EXE") = 1 Then
          On Error Resume Next
          LDT$ = TRIM$(CurDir)
          Kill "C:\FLEXOFT\FINBAT.$$$"
          NX% = FILEOPEN%("C:\FLEXOFT\COMBAT.$$$", 2, 0)
          Print #NX%, "FINBAT.$$$"
          Close #NX%: NX% = 0
          '
          NX% = FILEOPEN%("C:\FLEXOFT\DOSCOMM.BAT", 2, 0)
          Print #NX%, LUBAS$ + "PKUNZIP -O " + LUBAS$ + "FLEXPRF " + CRXX$

          Print #NX%, "COPY C:\FLEXOFT\COMBAT.$$$ C:\FLEXOFT\FINBAT.$$$"
          Close #NX%: NX% = 0
          '
          Shell "C:\FLEXOFT\DOSCOMM.PIF"
          TINI = Timer
          '
          MENSA$ = "No se Encontró el Controlador Adecuado" + Chr$(10) + Chr$(13)
          MENSA$ = MENSA$ + "para el Acceso a este Procedimiento." + Chr$(10) + Chr$(13) + Chr$(10) + Chr$(13)
          MENSA$ = MENSA$ + "FLEXOFT Realizará Ahora la Extracción" + Chr$(10) + Chr$(13)
          MENSA$ = MENSA$ + "Correspondiente." + Chr$(10) + Chr$(13) + Chr$(10) + Chr$(13)
          MENSA$ = MENSA$ + "Controlador: " + TRIM$(CRXX$)
          '
399       If Timer > TINI + 30 Then Exit Sub
          If EXISTE%("C:\FLEXOFT\FINBAT.$$$") < 1 Then GoTo 399
          MsgBox MENSA$, 4096 + 48, "Controladores de Proceso"
          Call FRESHALL
          FILEKILL "C:\FLEXOFT\COMBAT.$$$"
          FILEKILL "C:\FLEXOFT\FINBAT.$$$"
          If TRIM$(UCase$(CurDir)) <> "C:\FLEXOFT" Then
              On Error Resume Next
              FileCopy "C:\FLEXOFT\" + CRXX$, BRXX$
              Kill "C:\FLEXOFT\" + CRXX$
          End If
          On Error GoTo 0
          '
       End If
    End If
    Exit Sub
    '
End Sub
'

Sub TRALOCAL(Archi$, ALOCA$)
    '
    ARCHX$ = Archi$
    ARCHY$ = TRIM$(ALOCA$)
    If ARCHY$ = "" Then ARCHY$ = TRIM$(ARCHX$)
    Call BLANARCH("!" + ARCHY$)
    Screen.MousePointer = 11
    FIN& = ULTREG&(ARCHX$)
    For I& = 1 To FIN&
       Call TRACE(20, I&, FIN&)
       X$ = REGLEIDO$(ARCHX$, I&)
       Call REGAPP("!" + ARCHY$, X$)
    Next I&
    Call CIERRARCH("!" + ARCHY$)
    Screen.MousePointer = 1
    '
End Sub

Sub TRACENTRAL(Archi$, ALOCA$)
    '
    ARCHX$ = Archi$
    Call BLOQARCH(ARCHX$)
    Screen.MousePointer = 11
    ARCHY$ = TRIM$(ALOCA$)
    If ARCHY$ = "" Then ARCHY$ = TRIM$(ARCHX$)
    J& = 0
    FIN& = ULTREG&("!" + ARCHY$)
    For I& = 1 To FIN&
       Call TRACE(20, I&, FIN&)
       X$ = REGLEIDO$("!" + ARCHY$, I&)
       J& = J& + 1
       Call GRAREG(ARCHX$, X$, J&)
    Next I&
    Call SETULTREG(ARCHX$, J&)
    Call CIERRARCH(ARCHX$)
    Screen.MousePointer = 1
    '
End Sub
'

Sub BLOQARCH(Archi$)
    '
    Call BLOQUEAR(Archi$, 1)
    '
End Sub
'

Sub LIBARCH(Archi$)
   '
   Call BLOQUEAR(Archi$, (-1))
   '
End Sub
'
Sub BLOQUEAR(ARCHIVO$, MODO%)
   '
   ' Nueva rutina de bloqueos a partir del 25 Enero 2002
   ' ***************************************************
   '
   ' MODO% =  0 --> Devuelve sobre MODO% resultado del intento
   ' MODO% =  1 --> Bloquea GRAREG y REGAPP
   ' MODO% =  2 --> Bloqueo Total del Archivo
   ' MODO% = -1 --> Libera Bloqueos
   '
   Static LU$, FECHAC%, HO$
   Static USX$, NTX$, IDX$
   Static CANARBLO%
   Static ARCHNAME$(2048), FINULOCK%(2048), REGLOCK&(2048)
   '
   Static AA As String * 128
   Static BB As String * 128
   Static CC As String * 128
   Static DD As String * 256
   '
   If Abs(MODO%) > 1 Then Stop 'Rutina de Bloqueos Incompleta
   '
   GoSub ABLONAME      ' arma nombre de archivo a bloquear
   GoSub INIBLOQ       ' inicializa variables estaticas
   GoSub LEEARBLO      ' lee vector - devuelve IABLO%
   '
   If MODO% = (-1) Then
      GoSub LIBERBLOQ  ' libera archivo
      Exit Sub
   End If
   '
10 If FINULOCK%(IABLO%) > 0 Then
      If MODO% = 0 Then MODO% = 1
      Exit Sub         ' ya estaba bloqueado
   End If
   '
   SCA = Screen.MousePointer
   Screen.MousePointer = 11
   GoSub ARMATEBLO     ' arma texto mensaje de bloqueo
   '
12 CAINTEN% = 0
14 GoSub GRABALOCK     ' abre archivo .LCK sobre FINULOCK%(IABLO%)
   If LOGRABLO% < 1 Then
     '
16   If CAINTEN% < 5 Then
       CAINTEN% = CAINTEN% + 1
       GoSub REBLOKANTE
       If REBLOKA% > 0 Then
            GoTo 14
          Else
            GoTo 16
       End If
     End If
     '                 ' mensaje archivo en uso exclusivo
     GoSub MENYABLOQ
     If REINTEN% = 1 Then GoTo 12
     '
   End If
   Screen.MousePointer = SCA
   '
Exit Sub
   '
'  **************************************************
   '
   '
ABLONAME:     ' arma nombre de archivo a bloquear
   '
   ARX$ = TRIM$(UCase$(ARCHIVO$))
   If InStr(ARX$, ".") > 1 Then
      PPXX% = InStr(ARX$, ".")
      ARX$ = Left$(ARX$, PPXX% - 1)
   End If
   '
   If Left$(ARX$, 1) = "!" Then Exit Sub
   '
   If ARX$ = "*" Then
      If MODO% = (-1) Then GoSub LIBERTODO
      Exit Sub
   End If
   '
   If Len(ARX$) < 1 Or Len(ARX$) > 8 Then
      Exit Sub
   End If
   '
19 Return
   '
'  **************************************************
   '
INIBLOQ:      ' inicializa variables estaticas
   '
   If FECHAC% > 0 Then
     If LU$ <> "" Then
       If USX$ <> "" Then
         If NTX$ <> "" Then
           If IDX$ <> "" Then
             GoTo 29
           End If
         End If
       End If
     End If
   End If
   '
   FECHAC% = HOY(HO$)
   LU$ = LUDAT$
   USX$ = USERNAME$
   NTX$ = NOMTER$
   IDX$ = IDENTER$
   '
29 Return
   '
'  **************************************************
   '
LEEARBLO:
   '
   For IA% = 1 To CANARBLO%
      If ARCHNAME$(IA%) = ARX$ Then
         IABLO% = IA%
         GoTo 39                     ' ya estaba en lista
      End If
   Next IA%
   '
   CANARBLO% = CANARBLO% + 1
   IABLO% = CANARBLO%
   ARCHNAME$(IABLO%) = ARX$
   FINULOCK%(IABLO%) = 0
   '
39 Return
   '
'  **************************************************
   '
LIBERBLOQ:
   '
   If IABLO% > 0 Then
     If FINULOCK%(IABLO%) > 0 Then
       '
       NK% = FINULOCK%(IABLO%)
       On Error Resume Next
       Get #NK%, 1, AA$
       Call REPLA(AA$, Space$(116), 13, 116)
       Put #NK%, 1, AA$
       Close #NK%, NK% = 0
       FINULOCK%(IABLO%) = 0
       '
       If REGLOCK&(IABLO%) > 0 Then
         IAX% = IABLO%
         GoSub GRASECU
         REGLOCK&(IABLO%) = 0
       End If
       '
      End If
   End If
   '
59 Return
   '
'  **************************************************
   '
   '
ARMATEBLO:
   '
   TEBLO$ = Space$(128)
   Call REPLA(TEBLO$, ARX$, 1, 12)
   Call REPLA(TEBLO$, USX$, 13, 30)
   Call REPLA(TEBLO$, NTX$, 43, 32)
   Call REPLA(TEBLO$, HO$, 75, 10)
   Call REPLA(TEBLO$, Time$, 85, 10)
   Call REPLA(TEBLO$, IDX$, 95, 6)
   Call REPLA(TEBLO$, UCase$(App.EXEName), 101, 8)
   Call REPLA(TEBLO$, PROPRIN$, 109, 8)
   '
49 Return
   '
'  **************************************************
LIBERTODO:
   '
   For IA% = 1 To CANARBLO%
      If FINULOCK%(IA%) > 0 Then
         '
         GoSub ABREBLOQ
         Get #N1%, IA%, BB$
         B1$ = TRIM$(UCase$(Left$(BB$, 12)))
         If InStr(B1$, ".") > 1 Then
            PPXX% = InStr(B1$, ".")
            B1$ = Left$(B1$, PPXX% - 1)
         End If
         '
         If B1$ = ARCHNAME$(IA%) Then
            Call REPLA(BB$, Space$(116), 13, 116)
            Put #N1%, IA%, BB$
         End If
         '
         NK% = FINULOCK%(IA%)
         Get #NK%, 1, AA$
         Call REPLA(AA$, Space$(116), 13, 116)
         Put #NK%, 1, AA$
         Close #NK%: NK% = 0
         FINULOCK%(IA%) = 0
         '
         IAX% = IA%: GoSub GRASECU
         '
         Close #N1%: N1% = 0
         '
      End If
   Next IA%
   '
69 Return
   '
'  **************************************************
   '
GRASECU:       ' graba en archivo de secuencia de bloqueos
   '
   N2% = FreeFile
   Open LU$ + "SECUBLOQ.DAT" For Random Access Read Write As #N2% Len = 128
   If REGLOCK&(IAX%) < 1 Then
        RGXX& = 1 + LOF(N2%) / 128
        Get #N2%, RGXX&, CC$
        REGLOCK&(IAX%) = RGXX&
     Else
        RGXX& = REGLOCK&(IAX%)
        Get #N2%, RGXX&, CC$
        TESECU$ = CC$
        Call REPLA(TESECU$, HO$, 45, 8)
        Call REPLA(TESECU$, Time$, 55, 8)
        REGLOCK&(IAX%) = 0
   End If
   Call REPLA(TESECU$, Chr$(10), 128, 1)
   LSet CC$ = TESECU$
   Put #N2%, RGXX&, CC$
   Close #N2%: N2% = 0
   '
Return
   '
'  **************************************************
   '
ABREBLOQ:
   '
   If N1% < 1 Then      ' BLOQAR.DAT no esta abierto
     '
     CAINBLO% = 0
91   TACC1 = Timer
92   On Error GoTo 95
     N1% = FreeFile
     Open LU$ + "BLOQAR.DAT" For Random Access Read Write Lock Write As #N1% Len = 128
     On Error GoTo 0
     GoTo 99
     '
95   ERNU$ = TRIM$(Str$(Err.Number))
     Resume 96
96   On Error GoTo 0
     Close #N1%: N1% = 0
     If (Timer >= TACC1 And Timer < TACC1 + 15) Then ' reintenta 15 segundos
       GoTo 92                   ' vuelve a intentar
     End If
     If CAINBLO% < 3 Then
       CAINBLO% = CAINBLO% + 1
       MsgBox "Espere un Momento y Pulse Aceptar ..."
       GoTo 91
     End If
     ABLOCAR$ = LU$ + "BLOQAR.DAT"
     MsgBox "Error (" + ERNU$ + ") en la Apertura del Semáforo " + ABLOCAR$ + Chr$(10) + Chr$(13) + "Debe Ahora Re-Iniciarse la Aplicación " + UCase(App.EXEName) + "."
     Close: N1% = 0: End
     '
   End If
   '
99 Return
   '
'  **************************************************
   '
GRABALOCK:
   '
   LOGRABLO% = 0
   '
70 GoSub ABREBLOQ      ' abre BLOQAR.DAT en exclusiva sobre N1%
72 On Error GoTo 75
   NK% = FreeFile
   Open LU$ + ARX$ + ".LCK" For Random Access Read Write Lock Write As #NK% Len = 128
   ATXX% = 0: If LOF(NK%) < 128 Then ATXX% = 1
   LSet AA$ = TEBLO$
   Put #NK%, 1, AA$
   If ATXX% > 0 Then
      Close #NK%: NK% = 0
      GoTo 72
   End If
   FINULOCK%(IABLO%) = NK%
   REGLOCK&(IABLO%) = 0
   On Error GoTo 0
   '
   TESECU$ = Space$(128)
   ARXX$ = TRIM$(ARX$): If InStr(ARXX$, ".") < 1 Then ARXX$ = ARXX$ + ".DAT"
   Call REPLA(TESECU$, TRIM$(ARXX$), 1, 12)
   Call REPLA(TESECU$, App.EXEName, 15, 8)
   Call REPLA(TESECU$, HO$, 25, 8)
   Call REPLA(TESECU$, Time$, 35, 8)
   Call REPLA(TESECU$, USX$, 65, 24)
   Call REPLA(TESECU$, NTX$, 90, 24)
   IAX% = IABLO%: GoSub GRASECU
   '
   If MODO% = 0 Then MODO% = 1
   LOGRABLO% = 1            ' bloqueo exitoso
   GoTo 79
   '
75 Resume 76
76 On Error GoTo 0
   Close #NK%: NK% = 0
   '
79 Close #N1%: N1% = 0      ' cierra y libera BLOQAR.DAT
   Return
   '
'  **************************************************
   '
REBLOKANTE:
   '
   REBLOKA% = 0
   For KKII% = 1 To CANARBLO%
     If FINULOCK%(KKII%) > 0 Then
       Close #FINULOCK%(KKII%)
       FINULOCK%(KKII%) = (-1)
     End If
   Next KKII%
   '
   TTT1 = Timer
   While Timer < TTT1 + 2: Wend ' espera 2 segundos
   '
   GoSub ABREBLOQ      ' abre BLOQAR.DAT en exclusiva sobre N1%
   On Error GoTo 82
   For KKII% = 1 To CANARBLO%
     If FINULOCK%(KKII%) = (-1) Then
       NK% = FreeFile
       ARTX$ = ARCHNAME$(KKII%)
       Open LU$ + ARTX$ + ".LCK" For Random Access Read Write Lock Write As #NK% Len = 128
       FINULOCK%(KKII%) = NK%
     End If
   Next KKII%
   On Error GoTo 0
   REBLOKA% = 1
   GoTo 89
   '
82 Resume 83
83 On Error GoTo 0
   '
89 Close #N1%: N1% = 0
   Return
   '
'  **************************************************
   '
MENYABLOQ:
   '
   REINTEN% = 0
   If N1% > 0 Then Close #N1%: N1% = 0   ' libera BLOQAR.DAT
   For KKII% = 1 To CANARBLO%
     If FINULOCK%(KKII%) > 0 Then
       Close #FINULOCK%(KKII%)
       FINULOCK%(KKII%) = (-1)
     End If
   Next KKII%
   '
   USZ$ = "Desconocido"
   TRZ$ = "Desconocida"
   FFZ1$ = "00-00-00": FFZ2$ = "00:00:00"
   '
   On Error Resume Next
   NKX% = FreeFile
   Open LU$ + ARX$ + ".LCK" For Random Access Read Shared As #NKX% Len = 128
   Get #NKX%, 1, AA$
   On Error GoTo 0
   '
   If TRIM$(AA$) <> "" Then
     USZ$ = TRIM$(Mid$(AA$, 13, 30))
     TRZ$ = TRIM$(Mid$(AA$, 43, 32))
     FFZ1$ = TRIM$(Mid$(AA$, 75, 10))
     FFZ2$ = TRIM$(Mid$(AA$, 85, 10))
     PRZ$ = TRIM$(Mid$(AA$, 101, 8)) + ".EXE"
   End If
   '
   LOMAX% = 0
   USY$ = "Usuario " + TRIM$(USZ$): If Len(USY$) > LOMAX% Then LOMAX% = Len(USY$)
   TRY$ = "En Terminal " + TRIM$(TRZ$): If Len(TRY$) > LOMAX% Then LOMAX% = Len(TRY$)
   FFY$ = "Desde el " + TRIM$(FFZ1$) + " a las " + TRIM$(FFZ2$) + " horas": If Len(FFY$) > LOMAX% Then LOMAX% = Len(FFY$)
   PRY$ = "Ejecutando Proceso " + TRIM$(PRZ$): If Len(PRY$) > LOMAX% Then LOMAX% = Len(PRY$)
   '
   COMU$ = "ADVERTENCIA: Archivo '" + TRIM$(ARX$) + "' en Uso Exclusivo !!!\-\" + USY$ + "\" + TRY$ + "\" + FFY$ + "\" + PRY$ + "\\Intentar de Nuevo\Abandonar Tarea"
   If COMALTER%(COMU$) = 1 Then
        REINTEN% = 1
      Else
        Call CIERRARCH("*.*")
        Call FINAL("")
   End If
   '
Return
   '
'  **************************************************
   '
   '
End Sub
'


Function BLOQRESULT%(ARCHIVO$)
   '
   MODO% = 0
   Call BLOQUEAR(ARCHIVO$, MODO%)
   BLOQRESULT% = MODO%
   '
End Function
'

Function LOADFILE$(FINA$)
   '
   Static BFF As String * 2048
   '
   LFX$ = ""
   SCMA = Screen.MousePointer
   Screen.MousePointer = 11
   '
   On Error GoTo 97
   LOPUCO& = FileLen(FINA$)
   CARELE& = 1 + Int(LOPUCO& / 2048)
   NX% = FILEOPEN%(FINA$, 0, 2048)
   LFX$ = String$(2048 * CARELE&, 0)
   For JL& = 1 To CARELE&
     Get #NX%, JL&, BFF$
     Mid$(LFX$, 1 + 2048 * (JL& - 1), 2048) = BFF$
   Next JL&
   LFX$ = Left$(LFX$, LOPUCO&)
   GoTo 99
   '
97 Resume 98
98 LFX$ = ""
   '
99 If NX% > 0 Then Close #NX%
   Screen.MousePointer = SCMA
   LOADFILE$ = LFX$
   '
End Function

Sub BAJALDISCO()
   '
   Call CIERRARCH("*.*")
   Call BLOQEXE("*")
   Call LIBARCH("*.*")
   Reset
   Call BLOQEXE(App.EXEName)
   '
End Sub

Sub VIEMONIBLOQ()              ' monitoreo de bloqueos
   '
   Dim AA As String * 128
   Dim BB As String * 128
   Dim ARCHNAME$(2048)
   '
10 Screen.MousePointer = 11
   LU$ = LUDAT$
   N1% = FILEOPEN%(LU$ + "BLOQAR.DAT", 0, 128)
   CANARBLO% = LOF(N1%) / 128
   LOBLOQ = LOF(N1%)
   If CANARBLO% > 2048 Then
      Call MENSERR(24, "Archivo de Bloqueos\Fuera de Rango.")
      CANARBLO% = 2048
   End If
   '
   For IABLO% = 1 To CANARBLO%
      Get #N1%, IABLO%, BB$
      B1$ = TRIM$(UCase$(Left$(BB$, 12)))
      If InStr(B1$, ".") > 1 Then
         PPXX% = InStr(B1$, ".")
         B1$ = Left$(B1$, PPXX% - 1)
      End If
      ARCHNAME$(IABLO%) = B1$
   Next IABLO%
   Close #N1%: N1% = 0
   '
12 Screen.MousePointer = 11
   Call BLANARCH("!BLOQAR")
   '
   N1% = FILEOPEN%(LU$ + "BLOQAR.DAT", 0, 128)
   If LOF(N1%) > LOBLOQ Then
      Close #N1%: N1% = 0
      GoTo 10
   End If
   '
   For IABLO% = 1 To CANARBLO%
     If TRIM$(ARCHNAME$(IABLO%)) <> "" Then
       For JJI% = 1 To IABLO% - 1
         If ARCHNAME$(JJI%) = ARCHNAME$(IABLO%) Then GoTo 19 ' evita repetidos
       Next JJI%
       If EXISTE%(LU$ + ARCHNAME$(IABLO%) + ".LCK") = 1 Then
         On Error GoTo 15
         NK% = FreeFile
         Open LU$ + ARCHNAME$(IABLO%) + ".LCK" For Random Access Read Write Lock Write As #NK% Len = 128
         Get #NK%, 1, AA$
         Put #NK%, 1, AA$
         Close #NK%: NK% = 0
         On Error GoTo 0
         GoTo 19
         '
15       Resume 16
16       On Error GoTo 0
         Close #NK%: NK% = 0
         'Get #N1%, IABLO%, BB$
         'Call REGAPP("!BLOQAR", BB$)
         Call REGAPP("!BLOQAR", AA$)
       End If
     End If
19 Next IABLO%
   Close #N1%: N1% = 0
   '
   Call CIERRARCH("!BLOQAR")
   Screen.MousePointer = 1
   VTB% = VENTABU%("BLOQAR")
   If VTB% >= 0 Then GoTo 12
   '
End Sub

Sub VIEBLOQUEAR(ARCHIVO$, MODO%)
   '
   ' MODO% =  0 --> Devuelve sobre MODO% resultado del intento
   ' MODO% =  1 --> Bloquea GRAREG y REGAPP
   ' MODO% =  2 --> Bloqueo Total del Archivo
   ' MODO% = -1 --> Libera Bloqueos
   '
   Static CANARBLO%, ARCHNAME$(2048), FINULOCK%(2048), REGLOCK&(2048)
   Static USX$, NTX$, IDX$, FECHAC%, HO$, LU$
   Static AA As String * 128
   Static BB As String * 128
   Static CC As String * 128
   Static DD As String * 256
   Static YAENTRO%
   '
If Abs(MODO%) > 1 Then
  Call MENSERR(24, "Rutina de Bloqueos Incompleta")
  Stop
End If
   '
   GoSub ABLONAME   ' arma nombre de archivo a bloquear
   If CANARBLO% < 1 Then GoSub INIBLOQ
   If USX$ = "" Then GoSub INIBLOQ
   If NTX$ = "" Then GoSub INIBLOQ
   If IDX$ = "" Then GoSub INIBLOQ
   If LU$ = "" Then GoSub INIBLOQ
   '
   GoSub LEEARBLO   ' devuelve IABLO%
   If MODO% = (-1) Then GoSub LIBERBLOQ: Exit Sub
   '
10 If FINULOCK%(IABLO%) > 0 Then
      If MODO% = 0 Then MODO% = 1
      Exit Sub              ' ya esta bloqueado
   End If
   '
   GoSub ARMATEBLO
   GoSub ABREBLOQ
   '
   TACCE = Timer
   SCA = Screen.MousePointer
   Screen.MousePointer = 11
   '
   CAINTENTOS% = 0
   '
12 GoSub ABREBLOQ
   On Error GoTo 15
13 NK% = FreeFile
   Open LU$ + ARX$ + ".LCK" For Random Access Read Write Lock Write As #NK% Len = 128
   ATXX% = 0: If LOF(NK%) < 128 Then ATXX% = 1
   LSet AA$ = TEBLO$
   Put #NK%, 1, AA$
   If ATXX% > 0 Then
      Close #NK%: NK% = 0
      GoTo 13
   End If
   FINULOCK%(IABLO%) = NK%
   REGLOCK&(IABLO%) = 0
   On Error GoTo 0
   '
   Get #N1%, IABLO%, BB$
   B1$ = TRIM$(UCase$(Left$(BB$, 12)))
   If InStr(B1$, ".") > 1 Then
      PPXX% = InStr(B1$, ".")
      B1$ = Left$(B1$, PPXX% - 1)
   End If
   '
   If B1$ <> ARX$ Then
      Screen.MousePointer = 1
      MsgBox ("Error de Semáforo de Bloqueos - Re-Inicie FLEXOFT.")
      Close: N1% = 0: End
   End If
   '
   Call REPLA(BB$, Mid$(TEBLO$, 13), 13, 116)
   Put #N1%, IABLO%, BB$
   '
   TESECU$ = Space$(128)
   Call REPLA(TESECU$, TRIM$(B1$) + ".DAT", 1, 12)
   Call REPLA(TESECU$, Mid$(BB$, 101, 8), 15, 8)
   Call REPLA(TESECU$, Mid$(BB$, 75, 20), 25, 20)
   Call REPLA(TESECU$, USX$, 65, 24)
   Call REPLA(TESECU$, NTX$, 90, 24)
   IAX% = IABLO%: GoSub GRASECU
   '
   Close #N1%: N1% = 0
   '
   Screen.MousePointer = SCA
   If MODO% = 0 Then MODO% = 1
   Exit Sub                     ' bloqueo exitoso
   '
'  **************************************************
   '
15 Resume 16
16 On Error GoTo 0
   Close #NK%: NK% = 0
   If (Timer >= TACCE And Timer < TACCE + 5) Then ' reintenta 5 segundos
      GoTo 12                   ' vuelve a intentar
   End If
   '
   Get #N1%, IABLO%, BB$
   B1$ = TRIM$(UCase$(Left$(BB$, 12)))
   If InStr(B1$, ".") > 1 Then
      PPXX% = InStr(B1$, ".")
      B1$ = Left$(B1$, PPXX% - 1)
   End If
   '
   If B1$ <> ARX$ Then
      MsgBox ("Error de Semáforo de Bloqueos - Re-Inicie FLEXOFT.")
      Close: N1% = 0: End
   End If
   '
   B2$ = Mid$(BB$, 13, 30)
   B3$ = Mid$(BB$, 43, 32)
   B4$ = Mid$(BB$, 75, 10)
   B5$ = Mid$(BB$, 85, 10)
   B6$ = Mid$(BB$, 95, 6)
   B7$ = Mid$(BB$, 101, 8)
   B8$ = Mid$(BB$, 109, 8)
   B9$ = Mid$(BB$, 117, 12)
   Close #N1%: N1% = 0
   '
   TEMENBLO$ = Space$(255) + Chr$(10)
   Call REPLA(TEMENBLO$, ARX$, 1, 12)
   Call REPLA(TEMENBLO$, UCase$(App.EXEName), 15, 8)
   Call REPLA(TEMENBLO$, USX$, 26, 24)
   Call REPLA(TEMENBLO$, NTX$, 51, 24)
   Call REPLA(TEMENBLO$, HO$, 76, 8)
   Call REPLA(TEMENBLO$, Time$, 86, 8)
   Call REPLA(TEMENBLO$, UCase$(B7$), 96, 8)
   Call REPLA(TEMENBLO$, B2$, 106, 24)
   Call REPLA(TEMENBLO$, B3$, 131, 24)
   Call REPLA(TEMENBLO$, B4$, 156, 8)
   Call REPLA(TEMENBLO$, B5$, 166, 8)
   '
   GoSub ABREBLOQ
   N3% = FreeFile
   Open LU$ + "MENSABLO.DAT" For Random Access Read Write Shared As #N3% Len = 256
   RGXX& = 1 + LOF(N3%) / 256
   LSet DD$ = TEMENBLO$
   Put #N3%, RGXX&, DD$
   Close #N3%: N3% = 0
   Close #N1%: N1% = 0
   '
   If MODO% = 0 Then
      Screen.MousePointer = 1
      MODO% = (-1)
      Exit Sub
   End If
   '
   If CAINTENTOS% < 5 Then
      CAINTENTOS% = CAINTENTOS% + 1
      MsgBox "Sistema Ocupado ... Pulse Aceptar"
      TACCE = Timer
      GoTo 12
   End If
   '
   LOMAX% = 0
   USY$ = "Usuario " + TRIM$(B2$): If Len(USY$) > LOMAX% Then LOMAX% = Len(USY$)
   TRY$ = "En Terminal " + TRIM$(B3$): If Len(TRY$) > LOMAX% Then LOMAX% = Len(TRY$)
   FFY$ = "Desde el " + TRIM$(B4$) + " a las " + TRIM$(B5$) + " horas": If Len(FFY$) > LOMAX% Then LOMAX% = Len(FFY$)
   '
   COMU$ = "ADVERTENCIA: Archivo '" + TRIM$(ARX$) + "' en Uso Exclusivo !!!\-\" + USY$ + "\" + TRY$ + "\" + FFY$ + "\\Intentar de Nuevo\Abandonar Tarea"
   If COMALTER%(COMU$) = 1 Then
      TACCE = Timer
      GoTo 12
   End If
   Call CIERRARCH("*.*")
   Call FINAL("")
   '
'  **************************************************
   '
ABLONAME:
   '
   ARX$ = TRIM$(UCase$(ARCHIVO$))
   If InStr(ARX$, ".") > 1 Then
      PPXX% = InStr(ARX$, ".")
      ARX$ = Left$(ARX$, PPXX% - 1)
   End If
   '
   If Left$(ARX$, 1) = "!" Then Exit Sub
   '
   If ARX$ = "*" Then
      If MODO% = (-1) Then GoSub LIBERTODO
      Exit Sub
   End If
   '
   If Len(ARX$) < 1 Or Len(ARX$) > 8 Then
      Exit Sub
   End If
   '
19 Return
   '
'  **************************************************
   '
INIBLOQ:
   '
   FECHAC% = HOY(HO$)
   LU$ = LUDAT$
   USX$ = USERNAME$
   NTX$ = NOMTER$
   IDX$ = IDENTER$
   '
29 Return
   '
'  **************************************************
   '
LEEARBLO:
   '
   IABLO% = 0
   For IA% = 1 To CANARBLO%
      If ARCHNAME$(IA%) = ARX$ Then
         IABLO% = IA%
         GoTo 39                     ' ya estaba en lista
      End If
   Next IA%
   '
   GoSub ABREBLOQ
   For IA% = 1 To LOF(N1%) / 128
      '
      Get #N1%, IA%, BB$
      B1$ = TRIM$(UCase$(Left$(BB$, 12)))
      If InStr(B1$, ".") > 1 Then
         PPXX% = InStr(B1$, ".")
         B1$ = Left$(B1$, PPXX% - 1)
      End If
      If IA% > CANARBLO% Then
         If IA% > 2048 Then
            MsgBox ("Control de Bloqueos Fuera de Rango - Re-Inicie FLEXOFT.")
            Close: N1% = 0: End
         End If
         ARCHNAME$(IA%) = B1$
         FINULOCK%(IA%) = 0
         REGLOCK&(IA%) = 0
         CANARBLO% = IA%
      End If
      '
      If IABLO% < 1 Then
         If ARCHNAME$(IA%) = ARX$ Then
            IABLO% = IA%
            GoTo 35
         End If
      End If
      '
      If FECHANUM(Mid$(BB$, 75, 8)) <> FECHAC% Then
         Call REPLA(BB$, Space$(116), 13, 116)
         Put #N1%, IA%, BB$
      End If
      '
   Next IA%
   '
   If IABLO% < 1 Then
      CANARBLO% = CANARBLO% + 1
      If CANARBLO% > 2048 Then
         MsgBox ("Control de Bloqueos Fuera de Rango - Re-Inicie FLEXOFT.")
         Close: N1% = 0: End
      End If
      CC$ = Space$(128)
      ARX1$ = TRIM$(ARX$) + ".DAT"
      Call REPLA(CC$, ARX1$, 1, 12)
      LSet BB$ = CC$
      Put #N1%, CANARBLO%, BB$
      '
      IABLO% = CANARBLO%
      ARCHNAME$(IABLO%) = ARX$
      FINULOCK%(IABLO%) = 0
      REGLOCK&(IABLO%) = 0
   End If
   '
35 Close #N1%: N1% = 0
   '
39 Return
   '
'  **************************************************
   '
ARMATEBLO:
   '
   TEBLO$ = Space$(128)
   Call REPLA(TEBLO$, ARX$, 1, 12)
   Call REPLA(TEBLO$, USX$, 13, 30)
   Call REPLA(TEBLO$, NTX$, 43, 32)
   Call REPLA(TEBLO$, HO$, 75, 10)
   Call REPLA(TEBLO$, Time$, 85, 10)
   Call REPLA(TEBLO$, IDX$, 95, 6)
   Call REPLA(TEBLO$, UCase$(App.EXEName), 101, 8)
   Call REPLA(TEBLO$, PROPRIN$, 109, 8)
   '
49 Return
   '
'  **************************************************
   '
LIBERBLOQ:
   '
   If IABLO% > 0 Then
      If FINULOCK%(IABLO%) > 0 Then
         '
         GoSub ABREBLOQ
         Get #N1%, IABLO%, BB$
         B1$ = TRIM$(UCase$(Left$(BB$, 12)))
         If InStr(B1$, ".") > 1 Then
            PPXX% = InStr(B1$, ".")
            B1$ = Left$(B1$, PPXX% - 1)
         End If
         '
         If B1$ = ARX$ Then
            Call REPLA(BB$, Space$(116), 13, 116)
            Put #N1%, IABLO%, BB$
         End If
         '
         NK% = FINULOCK%(IABLO%)
         Get #NK%, 1, AA$
         Call REPLA(AA$, Space$(116), 13, 116)
         Put #NK%, 1, AA$
         Close #NK%, NK% = 0
         FINULOCK%(IABLO%) = 0
         '
         IAX% = IABLO%: GoSub GRASECU
         '
         Close #N1%: N1% = 0
         '
      End If
   End If
   '
59 Return
   '
'  **************************************************
   '
LIBERTODO:
   '
   For IA% = 1 To CANARBLO%
      If FINULOCK%(IA%) > 0 Then
         '
         GoSub ABREBLOQ
         Get #N1%, IA%, BB$
         B1$ = TRIM$(UCase$(Left$(BB$, 12)))
         If InStr(B1$, ".") > 1 Then
            PPXX% = InStr(B1$, ".")
            B1$ = Left$(B1$, PPXX% - 1)
         End If
         '
         If B1$ = ARCHNAME$(IA%) Then
            Call REPLA(BB$, Space$(116), 13, 116)
            Put #N1%, IA%, BB$
         End If
         '
         NK% = FINULOCK%(IA%)
         Get #NK%, 1, AA$
         Call REPLA(AA$, Space$(116), 13, 116)
         Put #NK%, 1, AA$
         Close #NK%: NK% = 0
         FINULOCK%(IA%) = 0
         '
         IAX% = IA%: GoSub GRASECU
         '
         Close #N1%: N1% = 0
         '
      End If
   Next IA%
   '
69 Return
   '
'  **************************************************
   '
GRASECU:
   '
   N2% = FreeFile
   Open LU$ + "SECUBLOQ.DAT" For Random Access Read Write As #N2% Len = 128
   If REGLOCK&(IAX%) < 1 Then
        RGXX& = 1 + LOF(N2%) / 128
        Get #N2%, RGXX&, CC$
        REGLOCK&(IAX%) = RGXX&
     Else
        RGXX& = REGLOCK&(IAX%)
        Get #N2%, RGXX&, CC$
        TESECU$ = CC$
        Call REPLA(TESECU$, HO$, 45, 8)
        Call REPLA(TESECU$, Time$, 55, 8)
        REGLOCK&(IAX%) = 0
   End If
   Call REPLA(TESECU$, Chr$(10), 128, 1)
   LSet CC$ = TESECU$
   Put #N2%, RGXX&, CC$
   Close #N2%: N2% = 0
   '
79 Return
   '
'  **************************************************
   '
ABREBLOQ:
   '
   If N1% > 0 Then GoTo 99
   '
   CAINTEN% = 0
91 TACC1 = Timer
92 On Error GoTo 95
   N1% = FreeFile
   Open LU$ + "BLOQAR.DAT" For Random Access Read Write Lock Write As #N1% Len = 128
   On Error GoTo 0
   GoTo 97
   '
95 ERNU$ = TRIM$(Str$(Err.Number))
   Resume 96
96 On Error GoTo 0
   Close #N1%: N1% = 0
   If (Timer >= TACC1 And Timer < TACC1 + 15) Then ' reintenta 15 segundos
      GoTo 92                   ' vuelve a intentar
   End If
   If CAINTEN% < 3 Then
      CAINTEN% = CAINTEN% + 1
      MsgBox "Espere un Momento y Pulse Aceptar ..."
      GoTo 91
   End If
   ABLOCAR$ = LU$ + "BLOQAR.DAT"
   MsgBox "Error (" + ERNU$ + ") en la Apertura del Semáforo " + ABLOCAR$ + Chr$(10) + Chr$(13) + "Debe Ahora Re-Iniciarse la Aplicación " + UCase(App.EXEName) + "."
   Close: N1% = 0: End
   '
97 If UCase$(App.EXEName) = "FLEXOFT" Then
      If YAENTRO% = 0 Then
         YAENTRO% = 1
         If LOF(N1%) > 128& * 1024& Then
            Screen.MousePointer = 11
            CANARBLO% = LOF(N1%) / 128
            IC% = 0
971         N2% = FILEOPEN%(LUCOM$ + "TRANSIT.$$$", 0, 128)
            If LOF(N2%) > 0 Then
               Close #N2%: N2% = 0
               Kill LUCOM$ + "TRANSIT.$$$"
               GoTo 971
            End If
            For IA% = 1 To CANARBLO%
               Get #N1%, IA%, BB$
               B1$ = TRIM$(UCase$(Left$(BB$, 12)))
               For IB% = 1 To IC%
                  If ABLO$(IB%) = B1$ Then GoTo 98
               Next IB%
               If IC% < 1024 Then
                  IC% = IC% + 1
                  ABLO$(IC%) = B1$
                  Put #N2%, IB%, BB$
               End If
98          Next IA%
            Close #N2%: N2% = 0
            Close #N1%: N1% = 0
            On Error Resume Next
            FileCopy LUCOM$ + "TRANSIT.$$$", LU$ + "BLOQAR.DAT"
            Screen.MousePointer = 1
            MsgBox "Control de Bloqueos Reindexado - Re-Inicie Aplicación " + UCase$(App.EXEName) + "."
            Close: End
         End If
      End If
   End If
   '
99 Return
   '
'  **************************************************
   '
End Sub
'

Sub VIEXTRACT(ARXX$) ' DISCONTINUADA DESDE 10/03/2002
    '
    BRXX$ = TRIM$(UCase$(ARXX$))
    If InStr(BRXX$, ".PRF") > 0 Then GoTo 199
    '
    If Left$(BRXX$, 1) = "!" Then
        BRXX$ = Mid$(BRXX$, 2)
    End If
    '
    If EXISTE%(LUBAS$ + "FLEXRFS.ZIP") = 1 Then
       If EXISTE%(LUBAS$ + "PKUNZIP.EXE") = 1 Then
          On Error Resume Next
          Kill "C:\FLEXOFT\FINBAT.$$$"
          NX% = FILEOPEN%("C:\FLEXOFT\COMBAT.$$$", 2, 0)
          Print #NX%, "FINBAT"
          Close #NX%: NX% = 0
          '
          LDT$ = TRIM$(UCase$(LUDAT$))
          NX% = FILEOPEN%("C:\FLEXOFT\DOSCOMM.BAT", 2, 0)
          Print #NX%, LUBAS$ + "PKUNZIP -O " + LUBAS$ + "FLEXRFS " + TRIM$(BRXX$) + ".RFS"
          Print #NX%, "COPY C:\FLEXOFT\COMBAT.$$$ C:\FLEXOFT\FINBAT.$$$"
          Close #NX%: NX% = 0
          '
          Shell "C:\FLEXOFT\DOSCOMM.PIF"
          TINI = Timer
          MENSA$ = "No se Encontró el Controlador Adecuado" + Chr$(10) + Chr$(13)
          MENSA$ = MENSA$ + "para el Acceso a esta Base de Datos." + Chr$(10) + Chr$(13) + Chr$(10) + Chr$(13)
          MENSA$ = MENSA$ + "FLEXOFT Realizará Ahora la Extracción" + Chr$(10) + Chr$(13)
          MENSA$ = MENSA$ + "Correspondiente." + Chr$(10) + Chr$(13) + Chr$(10) + Chr$(13)
          MENSA$ = MENSA$ + "Controlador: " + TRIM$(BRXX$) + ".RFS"
          '
149       If Timer > TINI + 30 Then Exit Sub
          If EXISTE%("C:\FLEXOFT\FINBAT.$$$") < 1 Then GoTo 149
          MsgBox MENSA$, 4096 + 48, "Controladores de Proceso"
          Call FRESHALL
          FILEKILL "C:\FLEXOFT\COMBAT.$$$"
          FILEKILL "C:\FLEXOFT\FINBAT.$$$"
          If LDT$ <> "C:\FLEXOFT\" Then
              On Error Resume Next
              FileCopy "C:\FLEXOFT\" + TRIM$(BRXX$) + ".RFS ", LDT$ + TRIM$(BRXX$) + ".RFS"
              Kill "C:\FLEXOFT\" + TRIM$(BRXX$) + ".RFS"
          End If
          On Error GoTo 0
          '
       End If
    End If
    Exit Sub
    '
199 CRXX$ = BRXX$
299 POSI% = InStr(CRXX$, "\")
    If POSI% > 0 Then
        CRXX$ = Mid$(CRXX$, POSI% + 1)
        GoTo 299
    End If
    '
    If EXISTE%(LUBAS$ + "FLEXPRF.ZIP") = 1 Then
       If EXISTE%(LUBAS$ + "PKUNZIP.EXE") = 1 Then
          On Error Resume Next
          LDT$ = TRIM$(CurDir)
          Kill "C:\FLEXOFT\FINBAT.$$$"
          NX% = FILEOPEN%("C:\FLEXOFT\COMBAT.$$$", 2, 0)
          Print #NX%, "FINBAT.$$$"
          Close #NX%: NX% = 0
          '
          NX% = FILEOPEN%("C:\FLEXOFT\DOSCOMM.BAT", 2, 0)
          Print #NX%, LUBAS$ + "PKUNZIP -O " + LUBAS$ + "FLEXPRF " + CRXX$

          Print #NX%, "COPY C:\FLEXOFT\COMBAT.$$$ C:\FLEXOFT\FINBAT.$$$"
          Close #NX%: NX% = 0
          '
          Shell "C:\FLEXOFT\DOSCOMM.PIF"
          TINI = Timer
          '
          MENSA$ = "No se Encontró el Controlador Adecuado" + Chr$(10) + Chr$(13)
          MENSA$ = MENSA$ + "para el Acceso a este Procedimiento." + Chr$(10) + Chr$(13) + Chr$(10) + Chr$(13)
          MENSA$ = MENSA$ + "FLEXOFT Realizará Ahora la Extracción" + Chr$(10) + Chr$(13)
          MENSA$ = MENSA$ + "Correspondiente." + Chr$(10) + Chr$(13) + Chr$(10) + Chr$(13)
          MENSA$ = MENSA$ + "Controlador: " + TRIM$(CRXX$)
          '
399       If Timer > TINI + 30 Then Exit Sub
          If EXISTE%("C:\FLEXOFT\FINBAT.$$$") < 1 Then GoTo 399
          MsgBox MENSA$, 4096 + 48, "Controladores de Proceso"
          Call FRESHALL
          FILEKILL "C:\FLEXOFT\COMBAT.$$$"
          FILEKILL "C:\FLEXOFT\FINBAT.$$$"
          If TRIM$(UCase$(CurDir)) <> "C:\FLEXOFT" Then
              On Error Resume Next
              FileCopy "C:\FLEXOFT\" + CRXX$, BRXX$
              Kill "C:\FLEXOFT\" + CRXX$
          End If
          On Error GoTo 0
          '
       End If
    End If
    Exit Sub
    '
End Sub
'
Function FIDATIM#(ARCHX$)
   '
   ARCHY$ = ARCHX$
   If InStr(ARCHY$, "\") < 1 Then ARCHY$ = LUDAT$ + ARCHY$
   If InStr(ARCHY$, ".") < 1 Then ARCHY$ = ARCHY$ + ".DAT"
   '
   FDT# = 0
   On Error Resume Next
   FDT# = FileDateTime(ARCHY$)
   On Error GoTo 0
   '
   FIDATIM# = FDT#
   '
End Function

Sub SAVEFILE(FINA$, LFX$)
   '
   SCMA = Screen.MousePointer
   Screen.MousePointer = 11
   '
   On Error Resume Next
   NX% = FILEOPEN%(FINA$, 2, 1)
   Close #NX%
   On Error GoTo 0
   '
   LOPUCO& = 0
   If EXISTE%(FINA$) > 0 Then
      LOPUCO& = FileLen(FINA$)
   End If
   '
   LFY$ = LFX$
   If Len(LFY$) < LOPUCO& Then
     LFY$ = LFY$ + String(LOPUCO& - Len(LFY$), 0)
   End If
   '
   On Error GoTo 97
   NX% = FILEOPEN%(FINA$, 4, 2048)
   Put #NX%, 1, LFY$
   GoTo 99
   '
97 Resume 99
   '
99 If NX% > 0 Then Close #NX%
   On Error GoTo 0
   Screen.MousePointer = SCMA
   '
End Sub
'
Function SERIVOL$(ULOVOL$)
    '
    Dim AA$(32)
    SERIVOL$ = ""
    '
    NX% = FreeFile
    On Error GoTo 98
    Open "C:\FLEXOFT\DOSCOMM.BAT" For Output As #NX%
    Print #NX%, "VOL " + ULOVOL$ + " >C:\FLEXOFT\SERIVOL"
    Close #NX%: NX% = 0
    On Error GoTo 0
    '
    Shell "C:\FLEXOFT\DOSCOMM.PIF"
    TINI = Timer
    While Timer < TINI + 3: Wend
    '
    NX% = FreeFile
    Open "C:\FLEXOFT\SERIVOL" For Input As #NX%
    AX% = 0
    While Not EOF(NX%)
        Line Input #NX%, XX$
        If AX% < 32 Then
            AX% = AX% + 1
            AA$(AX%) = XX$
        End If
    Wend
    Close #NX%: NX% = 0
    '
    For AY% = 1 To AX%
      XX$ = TRIM$(AA$(AY%))
      POSI% = InStr(XX$, "-")
      If POSI% > 5 Then
        If POSI% <= Len(XX$) - 4 Then
          If Mid$(XX$, POSI% - 5, 1) = " " Then
            SERIVOL$ = Mid$(XX$, POSI% - 4, 9)
            If UCase$(ULOVOL$) = "C:" Then
              SEVOLX$ = LOADFILE$("C:\FLEXOFT\SERIVOL.DRV")
              If ENCRIPTA$(SERIVOL$) <> SEVOLX$ Then
                Call SAVEFILE("C:\FLEXOFT\SERIVOL.DRV", ENCRIPTA$(SERIVOL$))
              End If
            End If
            Exit Function
          End If
        End If
      End If
    Next AY%
    '
    SEVOLX$ = LOADFILE$(ULOVOL$ + "\FLEXOFT\SERIVOL.DRV")
    SERIVOL$ = TRIM$(DESENCRI$(SEVOLX$))
    If SERIVOL$ = "" Then
      GoTo 99
    End If
    Exit Function
    '
98  Resume 99
99  On Error GoTo 0
    MsgBox "Imposible Identificar Disco '" + ULOVOL$ + "'."
    Close: End
    '
End Function

Function FILTERGETRECORD$(Archi$, X%, ARGU$)
   RFF$ = RECFILT$(Archi$, X%, ARGU$)
   FGRC$ = ""
   If Len(RFF$) >= 4 Then
     REFI& = CVS(Left$(RFF$, 4))
     If REFI& > 0 Then
       If REFI& <= ULTREG&(Archi$) Then
         FGRC$ = REGLEIDO$(Archi$, REFI&)
       End If
     End If
   End If
   '
   If FGRC$ = "" Then FGRC$ = REGBLAN$(Archi$)
   FILTERGETRECORD$ = FGRC$
End Function

Sub FILTERPUTRECORD(Archi$, X%, ARGU$, RECORD$)
   RFF$ = RECFILT$(Archi$, X%, ARGU$)
   REFJ& = 0
   If Len(RFF$) >= 4 Then
     REFI& = CVS(Left$(RFF$, 4))
     If REFI& > 0 Then
       If REFI& <= ULTREG&(Archi$) Then
         REFJ& = REFI&
       End If
     End If
   End If
   If REFJ& < 1 Then REFJ& = 1 + ULTREG&(Archi$)
   Call GRAREG(Archi$, RECORD$, REFJ&)
End Sub


