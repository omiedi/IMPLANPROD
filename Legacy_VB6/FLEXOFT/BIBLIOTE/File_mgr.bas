Attribute VB_Name = "FILE_MGR"
Global CAMBIOVALGRILLA%
'
Global BLOCKFINU%
Global CANARCHIV%
Global ECOPRECIO%

Global BUFLOCKLIST As String * 128
Global BUFLOCKFILE As String * 128
'
Global FINAMSTRI$
Global FINAMSTRI2$
'
Global FINUMBER%(128)
Global FINULOCK%(128)
Global RELILOCK%(128)
Global REGLOCK&(128)
Global TEBLOFI$(128)
Global PRIREMOD&(128)
Global ULTACTEOF#(128)
'
Global VECREGIS$(), ARVEC$
'
Global XFIELD%(128)
Global XSTRIN$(128)
Global ULTACTIDX#(128)
'
Global FIFULLNAME$(128)
Global FISHORNAME$(128)
Global FILOCKNAME$(128)
Global FISTRUNAME$(128)
Global FILOCKTEXT$(128)
'
Global BLANREG$(128)
Global LOREGIS%(128)
Global FILETIT$(128)
Global TVASTRI$(128)
Global POSSTRI$(128)
Global LONSTRI$(128)
Global IDXSTRI$(128)
Global DECSTRI$(128)
'
Global UREGIST&(128)
'
Global UREGLEI&(128)
Global RELEIDO$(128)
'
Public TIMATEFIL%
'
Function REGSEL$(ARCHX$)
   '
   XXX$ = "": REGA& = 0
   Call SELECHO(ARCHX$)
   '
   ARCHX1$ = TRIM$(UCase$(ARCHX$))
   PPXX% = InStr(ARCHX1$, "/")
   If PPXX% < 2 Then PPXX% = 1 + Len(ARCHX1$)
   ARCHX1$ = TRIM$(UCase$(Left$(ARCHX1$, PPXX% - 1)))
   If ARCHX1$ = "MASTER" Then ARCHX1$ = "ECOMAST"
   '
   For KKI% = 0 To 15
     If FORSELEC.ATABLA(KKI%).Caption = ARCHX1$ Then
       REGA& = XVALO(FORSELEC.RENUAC(KKI%).Caption)
       If REGA& > 0 Then
         If REGA& <= ULTREG&(ARCHX1$) Then
           XXX$ = REGLEIDO$(ARCHX1$, REGA&)
           GoTo 90
         End If
       End If
     End If
   Next KKI%
   '
90 YYY$ = MKS$(REGA&)
   REGSEL$ = XXX$ + YYY$
   '
End Function
'
Function REGNUM&(ARCHX$)
   IARCH% = FILEINDEX%(ARCHX$)
   REGNUM& = UREGLEI&(IARCH%)
End Function
'
Sub BLANARCH(ARCHX$)
    '
    ARCHX1$ = ARCHX$: MODOX$ = ""
    PPXXX% = InStr(ARCHX1$, "/")
    If PPXXX% > 0 Then
      ARCHX1$ = Left$(ARCHX$, PPXXX% - 1)
      MODOX$ = UCase$(Mid$(ARCHX$, PPXXX% + 1))
    End If
    '
    IARCH% = FILEINDEX%(ARCHX1$)
    FINUX% = FILENBR%(ARCHX1$)
    REGS$ = BLANREG$(IARCH%)
    ARY1$ = FIFULLNAME(IARCH%)
    ARY2$ = FISHORNAME(IARCH%)
    '
    '
    If MODOX$ = "NOERR" Then
        Call CIERRARCH(ARCHX1$)
        On Error Resume Next
        Kill ARY1$
        On Error GoTo 0
      Else
        Call GRAREG(ARCHX1$, REGS$, 1)
        Call CIERRARCH(ARCHX1$)
        FILEKILL ARY1$
    End If
    '
    On Error Resume Next
    Kill LUDAT$ + ARY2$ + ".X*"
    On Error GoTo 0
    '
    PRIREMOD&(IARCH%) = 0
    UREGIST&(IARCH%) = 0
    UREGLEI&(IARCH%) = 0
    RELEIDO$(IARCH%) = String$(LOREGIS%(IARCH%), 0)
    '
End Sub
'
Sub SETREG(ARCHX$, REGG&)
    '
    IARCH% = FILEINDEX%(ARCHX$)
    FINUME% = FINUMBER%(IARCH%)
    UREGI& = UREGIST&(IARCH%)
    LOREGI% = LOREGIS%(IARCH%)
    '
    RGLL& = 0: REGS$ = BLANREG$(IARCH%)
    If REGG& > 0 Then
      If REGG& <= UREGI& Then
        RGLL& = REGG&
        RELE$ = LEEREGIS$(IARCH%, FINUME%, LOREGI%, REGG&)
      End If
    End If
    '
End Sub
'
Function FILENBR%(Archi$)
   '
   FINDEX% = FILEINDEX%(Archi$)
   finbr% = FINUMBER%(FINDEX%)
   '
   LOARX& = (-1)
   On Error Resume Next
   LOARX& = LOF(finbr%)
   On Error GoTo 0
   If LOARX& < 0 Then finbr% = 0
   '
   If finbr% < 1 Then
     ARCHX$ = FIFULLNAME$(FINDEX%)
     finbr% = FILEOPEN%(ARCHX$, 4, 1024)
     FINUMBER%(FINDEX%) = finbr%
     UREGLEI&(FINDEX%) = 0
   End If
   FILENBR% = finbr%
   '
End Function
'
Function RECLEN%(Archi$)
   '
   FINDEX% = FILEINDEX%(Archi$)
   RECLEN% = LOREGIS%(FINDEX%)
   '
End Function
'
Function REGBLAN$(Archi$)
   '
   FINDEX% = FILEINDEX%(Archi$)
   REGBLAN$ = BLANREG$(FINDEX%)
   If REGBLAN$ = "" Then
      Call LEESTRUCARCH(FINDEX%)
      REGBLAN$ = BLANREG$(FINDEX%)
   End If
   '
End Function
'
Function ULTREG&(Archi$)
   '
   If Left$(Archi$, 1) = "#" Then
     If ARVEC$ = "" Then ARVEC$ = UCase$(Archi$)
     If ARVEC$ <> UCase$(Archi$) Then
       Call MENSERR(24, "Imposible Hallar LOF de Archivo Virtual '" + Archi$ + "'.\Vector en Uso por Archivo '" + ARVEC$ + "'.\Consulte.")
       Call FINAL("")
     End If
     On Error Resume Next
     URE& = UBound(VECREGIS$)
     On Error GoTo 0
     ULTREG& = URE&
     Exit Function
   End If
   '
   '
   FINDEX% = FILEINDEX%(Archi$)
   If UREGIST&(FINDEX%) > 0 Then
     If FIDATIM#(Archi$) = ULTACTEOF#(FINDEX%) Then
       GoTo 99
     End If
   End If
   '
   FINUME% = FILENBR%(Archi$)
   Close #FINUME%: FINUMBER%(FINDEX%) = 0
   '
   FINUME% = FILENBR%(Archi$)
   LOREGI% = LOREGIS%(FINDEX%)
   ULTACTEOF#(FINDEX%) = FIDATIM#(Archi$)
   '
   If LOREGI% < 1 Then
     Call MENSERR(24, "Error en Controlador de Acceso\a Base de Datos '" + TRIM$(Archi$) + "'.\Longitud de Registro no Válida.\  \Re-Intente o Consulte a Soporte de Sistemas.")
     Call FINAL("")
   End If
   '
   CARETO& = LOF(FINUME%) / LOREGI%
   UREGIST&(FINDEX%) = CARETO&
   While UREGIST&(FINDEX%) > 0
     RGLX$ = REGLEIDO$(Archi$, UREGIST&(FINDEX%))
     If RGLX$ <> String$(LOREGI%, 0) Then GoTo 99
     UREGIST&(FINDEX%) = UREGIST&(FINDEX%) - 1
   Wend
   '
99 ULTREG& = UREGIST&(FINDEX%)
   '
End Function
'
Function REGLEIDO$(Archi$, REG&)
   '
   If Left$(Archi$, 1) = "#" Then
     If ARVEC$ = "" Then ARVEC$ = UCase$(Archi$)
     If ARVEC$ <> UCase$(Archi$) Then
       Call MENSERR(24, "Imposible Leer de Archivo Virtual '" + Archi$ + "'.\Vector en Uso por Archivo '" + ARVEC$ + "'.\Consulte.")
       Call FINAL("")
     End If
     On Error Resume Next
     URE& = UBound(VECREGIS$)
     On Error GoTo 0
     '
     If REG& < 1 Or REG& > URE& Then
       ARXX1$ = Mid$(Archi$, 2)
       REGLEIDO$ = REGBLAN$(ARXX1$)
       Exit Function
     End If
     '
     REGLEIDO$ = VECREGIS$(REG&)
     If REGLEIDO$ = "" Then
       ARXX1$ = Mid$(Archi$, 2)
       REGLEIDO$ = REGBLAN$(ARXX1$)
     End If
     Exit Function
   End If
   '
   '
   FINDEX% = FILEINDEX%(Archi$)
   FINUME% = FILENBR%(Archi$)
   LOREGI% = LOREGIS%(FINDEX%)
   '
   REGLEIDO$ = LEEREGIS$(FINDEX%, FINUME%, LOREGI%, REG&)
   '
End Function
'
Function REGACT$(ARCHX$)
   '
   FINDEX% = FILEINDEX%(ARCHX$)
   REGACT$ = RELEIDO$(FINDEX%)
   '
End Function
'
Sub GRAREG(Archi$, REGIS$, REGIGRA&)
   '
   If Left$(Archi$, 1) = "#" Then
     If ARVEC$ = "" Then ARVEC$ = UCase$(Archi$)
     If ARVEC$ <> UCase$(Archi$) Then
       Call MENSERR(24, "Imposible Grabar Archivo Virtual '" + Archi$ + "'.\Vector en Uso por Archivo '" + ARVEC$ + "'.\Consulte.")
       Call FINAL("")
     End If
     On Error Resume Next
     URE& = UBound(VECREGIS$)
     On Error GoTo 0
     '
     If REGIGRA& > URE& Then
       ReDim Preserve VECREGIS$(REGIGRA&)
     End If
     '
     VECREGIS$(REGIGRA&) = REGIS$
     Exit Sub
   End If
   '
   URE& = ULTREG&(Archi$)
   If REGIGRA& > URE& Then
      Call BLOQARCH(Archi$)
   End If
   '
   If REGIGRA& < 1 Or REGIGRA& > URE& + 1 Then
      Call MENSERR(24, "Intento de Grabar Fuera de Límites de Archivo.\Archivo = " + TRIM$(Archi$) + "  ..  Registro = " + TRIM$(Str$(REGIGRA&)) + "  ..  Lastrecord = " + TRIM$(Str$(URE&)))
      Exit Sub
   End If
   '
   FINDEX% = FILEINDEX%(Archi$)
   FINUME% = FILENBR%(Archi$)
   LOREGI% = LOREGIS%(FINDEX%)
   '
   POSGRABA& = (REGIGRA& - 1) * LOREGI% + 1
   BUFFGRABA$ = Left$(REGIS$ + String$(LOREGI%, 0), LOREGI%)
   '
   Put #FINUME%, POSGRABA&, BUFFGRABA$
   UREGLEI&(FINDEX%) = REGIGRA&
   RELEIDO$(FINDEX%) = BUFFGRABA
   If REGIGRA& > UREGIST&(FINDEX%) Then
     UREGIST&(FINDEX%) = REGIGRA&
   End If
   '
   If PRIREMOD&(FINDEX%) = 0 Or REGIGRA& < PRIREMOD&(FINDEX%) Then
     PRIREMOD&(FINDEX%) = REGIGRA&
   End If
   '
End Sub
'
Sub REGAPP(ARCHX$, REGIS$)
   '
   FINDEX% = FILEINDEX%(ARCHX$)
   ULTACTEOF#(FINDEX%) = 0
   '
   REG& = 1 + ULTREG&(ARCHX$)
   Call GRAREG(ARCHX$, REGIS$, REG&)
   '
End Sub
'
Function LEEREGIS$(FINDEX%, FINUME%, LOREGI%, REG&)
   '
   If REG& < 1 Then
     RELEIDO$(FINDEX%) = BLANREG$(FINDEX%)
     GoTo 90
   End If
   '
   POSLECTURA& = (REG& - 1) * LOREGI% + 1
   If REG& <> UREGLEI&(FINDEX%) Then
     '
     On Error Resume Next
15   Get #FINUME%, POSLECTURA&, RELEIDO$(FINDEX%)
     If Err = 59 Then
       On Error GoTo 0
       RELEIDO$(FINDEX%) = BLANREG$(FINDEX%)
       GoTo 15
     End If
     On Error GoTo 0
     UREGLEI&(FINDEX%) = REG&
     '
   End If
90 LEEREGIS$ = RELEIDO$(FINDEX%)
   '
End Function
'
Sub BLOQARCH(ARCHX$)
   '
   SMPANTE = Screen.MousePointer
   If Left$(ARCHX$, 1) = "!" Then Exit Sub
   '
   FINDEX% = FILEINDEX%(ARCHX$)
   If FINULOCK%(FINDEX%) > 0 Then
      Screen.MousePointer = SMPANTE
      Exit Sub                      ' ya esta bloqueado
   End If
   '
   Call ARMATEBLO(FINDEX%, TEBLO$)
   Call OPENBLOCKLIST               ' abre BLOQAR.DAT en uso exclusivo
   ILOCK% = RELILOCK%(FINDEX%)
   If ILOCK% < 1 Then Call LEERELOCK(FINDEX%)
   '
   Call LOCKFILE(FINDEX%, TEBLO$)   ' abre archivo .LCK
   Call CLOSEBLOCKLIST              ' cierra y libera BLOQAR.DAT
   Screen.MousePointer = SMPANTE
   '
End Sub
'
Function ISLOCKED%(ARCHX$)
   '
   If Left$(ARCHX$, 1) = "!" Then
      ISLOCKED% = 1
      Exit Function
   End If
   '
   REPITE% = 0
   '
20 ISLOCKED% = 0
   FINDEX% = FILEINDEX%(ARCHX$)
   If FINULOCK%(FINDEX%) > 0 Then
      ARX$ = FILOCKNAME$(FINDEX%)
      LOCKFILETEXT$ = LOADFILE$(ARX$)
      If LOCKFILETEXT$ = TEBLOFI$(FINDEX%) Then
         NXX1% = FreeFile
         On Error Resume Next
         Open ARX$ For Random Access Read Write Shared As #NXX1% Len = 128
         If Err = 70 Then
            On Error GoTo 0
            Close #NXX1%
            If REPITE% < 1 Then
               ' CONTROLA 3 SEGUNDOS MÁS TARDE A VER SI MANTIENE BLOQUEO
               REPITE% = 1
               TINI = Timer
               While Timer < TINI + 3: Openforms = DoEvents: Wend
               GoTo 20
            End If
            ISLOCKED% = 1
            Exit Function
         End If
         On Error GoTo 0
         Close #NXX1
      End If
   End If
   '
80
End Function
'
Sub LIBARCH(ARCHX$)
   '
   If Left$(ARCHX$, 1) = "*" Then
     If LUDAT$ <> "" Then
       Call OPENBLOCKLIST               ' abre BLOQAR.DAT en uso exclusivo
       CANARCHIV% = Len(FINAMSTRI$) / 16
       For KKII% = 1 To CANARCHIV%
         If FINULOCK%(KKII%) > 0 Then
           ILOCK% = RELILOCK%(KKII%)
           If ILOCK% < 1 Then Call LEERELOCK(KKII%)
           Call LIBERFILE(KKII%)   ' abre archivo .LCK
         End If
       Next KKII%
       Call CLOSEBLOCKLIST
     End If
     Exit Sub
   End If
   '
   FINDEX% = FILEINDEX%(ARCHX$)
   If FINULOCK%(FINDEX%) < 1 Then
      Exit Sub                      ' esta libre
   End If
   '
   Call OPENBLOCKLIST               ' abre BLOQAR.DAT en uso exclusivo
   ILOCK% = RELILOCK%(FINDEX%)
   If ILOCK% < 1 Then Call LEERELOCK(FINDEX%)
   Call LIBERFILE(FINDEX%)   ' abre archivo .LCK
   Call CLOSEBLOCKLIST              ' cierra y libera BLOQAR.DAT
   '
End Sub
'
Sub ARMATEBLO(FINDEX%, TEBLO$)
   '
   Static TEBLOBAS$
   '
   If TEBLOBAS$ = "" Then
      FECHAC% = HOY(HO$)
      LU$ = LUDAT$
      USX$ = USERNAME$
      NTX$ = NOMTER$
      IDX$ = IDENTER$
      '
      TEBLOBAS$ = Space$(128)
      Call REPLA(TEBLOBAS$, USX$, 13, 30)
      Call REPLA(TEBLOBAS$, NTX$, 43, 32)
      Call REPLA(TEBLOBAS$, HO$, 75, 10)
      Call REPLA(TEBLOBAS$, IDX$, 95, 6)
      Call REPLA(TEBLOBAS$, UCase$(App.EXEName), 101, 8)
   End If
   '
   TEBLO$ = TEBLOBAS$
   Call REPLA(TEBLO$, FILOCKTEXT$(FINDEX%), 1, 12)
   Call REPLA(TEBLO$, Time$, 85, 10)
   Call REPLA(TEBLO$, PROPRIN$, 109, 8)
   '
End Sub
'
Sub OPENBLOCKLISTVIE()
   ' deshabilitado a partir del 17/05/03
   If BLOCKFINU% < 1 Then
     '
     CAINTEN% = 0
91   TACC1 = Timer
92   On Error GoTo 95
     BLOCKFINU% = FreeFile
     Open LUDAT$ + "BLOQAR.DAT" For Random Access Read Write Lock Write As #BLOCKFINU% Len = 128
     GoTo 99
     '
95   Resume 96
96   On Error GoTo 0
     Close #BLOCKFINU%: BLOCKFINU% = 0
     If (Timer >= TACC1 And Timer < TACC1 + 15) Then ' reintenta 15 segundos
        GoTo 92                   ' vuelve a intentar
     End If
     '
'     If CAINTEN% < 3 Then
'        CAINTEN% = CAINTEN% + 1
'        MsgBox "Espere un Momento y Pulse Aceptar ..."
'        GoTo 91
'     End If
     If CAINTEN% < 4 Then
        CAINTEN% = CAINTEN% + 1
        TINI = Timer: While Timer < TINI + 1: Wend
        GoTo 91
     End If
     '
     MsgBox "Semáforo en Uso Exclusivo - Re-Inicie Aplicación " + UCase(App.EXEName) + "."
     Close: Reset: BLOCKFINU% = 0: End
     '
   End If
   '
99 On Error GoTo 0
   '
End Sub

Sub OPENBLOCKLIST()
   '
   If BLOCKFINU% < 1 Then
     '
     CAINTEN% = 0
91   TACC1 = Timer
92   On Error Resume Next
     BLOCKFINU% = FreeFile
     Open LUDAT$ + "BLOQAR.DAT" For Random Access Read Write Lock Write As #BLOCKFINU% Len = 128
     '
     If Err > 0 Then
        Close #BLOCKFINU%: BLOCKFINU% = 0
        If (Timer >= TACC1 And Timer < TACC1 + 15) Then ' reintenta 15 segundos
          TACC2 = Timer: While Timer < TACC2 + 0.5: Wend
          GoTo 92                   ' vuelve a intentar
        End If
        '
'        If CAINTEN% < 3 Then
'            CAINTEN% = CAINTEN% + 1
'            MsgBox "Espere un Momento y Pulse Aceptar ..."
'            GoTo 91
'          Else
'            MsgBox "Semáforo en Uso Exclusivo - Re-Inicie Aplicación " + UCase(App.EXEName) + "."
'            Close: Reset: BLOCKFINU% = 0: End
'        End If
        If CAINTEN% < 4 Then
            CAINTEN% = CAINTEN% + 1
            TINI = Timer: While Timer < TINI + 1: Wend
            GoTo 91
          Else
            MsgBox "Semáforo en Uso Exclusivo - Re-Inicie Aplicación " + UCase(App.EXEName) + "."
            Close: Reset: BLOCKFINU% = 0: End
        End If
     End If
     '
   End If
   '
99 On Error GoTo 0
   '
End Sub

Sub CLOSEBLOCKLIST()
   '
   If BLOCKFINU% > 0 Then
     Close #BLOCKFINU%
     BLOCKFINU% = 0
   End If
   '
End Sub
'
Sub LEERELOCK(FINDEX%)
   Static PLOCK$
   '
   ARX$ = FILOCKTEXT$(FINDEX%)
   '
   If ARX$ <> "" Then
     '
     If PLOCK$ = "" Then
       For IA% = 1 To LOF(BLOCKFINU%) / 128
         Get #BLOCKFINU%, IA%, BUFLOCKLIST$
         B1$ = TRIM$(UCase$(Left$(BUFLOCKLIST$, 12)))
         If InStr(B1$, ".") < 1 Then
           B1$ = AJUSTI$(B1$ + ".DAT", 12) + Mid$(BUFLOCKLIST$, 13)
           BUFLOCKLIST$ = B1$
           Put #BLOCKFINU%, IA%, BUFLOCKLIST$
           B1$ = TRIM$(UCase$(Left$(BUFLOCKLIST$, 12)))
         End If
         PLOCK$ = PLOCK$ + AJUSTI$(B1$, 12)
       Next IA%
     End If
     '
     IA% = (11 + InStr(PLOCK$, ARX$)) / 12
     If IA% > 0 Then
       Get #BLOCKFINU%, IA%, BUFLOCKLIST$
       B1$ = TRIM$(UCase$(Left$(BUFLOCKLIST$, 12)))
       If B1$ = ARX$ Then
         RELILOCK%(FINDEX%) = IA%
         GoTo 20
       End If
     End If
     '
10   For IA% = 1 To LOF(BLOCKFINU%) / 128
       Get #BLOCKFINU%, IA%, BUFLOCKLIST$
       B1$ = TRIM$(UCase$(Left$(BUFLOCKLIST$, 12)))
       If InStr(B1$, ".") < 1 Then
         B1$ = AJUSTI$(B1$ + ".DAT", 12) + Mid$(BUFLOCKLIST$, 13)
         BUFLOCKLIST$ = B1$
         Put #BLOCKFINU%, IA%, BUFLOCKLIST$
         B1$ = TRIM$(UCase$(Left$(BUFLOCKLIST$, 12)))
       End If
       '
       If B1$ = ARX$ Then
         RELILOCK%(FINDEX%) = IA%
         GoTo 20
       End If
     Next IA%
     '
     IA% = LOF(BLOCKFINU%) / 128 + 1
     B1$ = AJUSTI$(ARX$, 12) + Space$(116)
     PLOCK$ = PLOCK$ + AJUSTI$(B1$, 12)
     BUFLOCKLIST$ = B1$
     Put #BLOCKFINU%, IA%, BUFLOCKLIST$
     Close #BLOCKFINU%: BLOCKFINU% = 0
     Call OPENBLOCKLIST
     GoTo 10
     '
20   IA% = RELILOCK%(FINDEX%)
     If IA% < 1 Or IA% > LOF(BLOCKFINU%) / 128 Then
       Call MENSERR(24, "Imposible Continuar Ejecución.")
       Close: Reset: End
     End If
     '
     Get #BLOCKFINU%, IA%, BUFLOCKLIST$
     B1$ = TRIM$(UCase$(Left$(BUFLOCKLIST$, 12)))
     If B1$ <> ARX$ Then
       Call MENSERR(24, "Imposible Continuar Ejecución.\   \Error en Semáforo de Bloqueos.\Consulte a su Soporte de Sistemas.")
       Close: Reset: End
     End If
     '
   End If
   '
End Sub
'

Sub LOCKFILE(FINDEX%, TEBLO$)
   '
12 CAINTEN1% = 0: CAINTEN2% = 0
   SCA = Screen.MousePointer: If SCA < 1 Then SCA = 1
   Screen.MousePointer = 11
   On Error GoTo 15
   '
13 LOCKFINU% = FreeFile
   ARX$ = FILOCKNAME$(FINDEX%)
   On Error GoTo 15
   '
   ' abre y graba archivo '*.LCK' ****************************
   Open ARX$ For Random Access Read Write Lock Write As #LOCKFINU% Len = 128
   ATXX% = 0: If LOF(LOCKFINU%) < 128 Then ATXX% = 1
   LSet BUFLOCKFILE$ = TEBLO$
   Put #LOCKFINU%, 1, BUFLOCKFILE$
   If ATXX% > 0 Then
      Close #LOCKFINU%
      GoTo 13
   End If
   FINULOCK%(FINDEX%) = LOCKFINU%
   '
   ' graba mensaje bloqueo en 'BLOQAR.dat'
   Get #BLOCKFINU%, RELILOCK%(FINDEX%), BUFLOCKLIST$
   B1$ = TRIM$(UCase$(Left$(BUFLOCKLIST$, 12)))
   If InStr(B1$, ".") < 2 Then B1$ = B1$ + ".DAT"
   If B1$ <> TRIM$(FILOCKTEXT$(FINDEX%)) Then
      MsgBox ("Error de Semáforo de Bloqueos - Re-Inicie FLEXOFT.")
      Close: Reset: End
   End If
   LSet BUFLOCKLIST$ = TEBLO$
   Put #BLOCKFINU%, RELILOCK%(FINDEX%), BUFLOCKLIST$
   abc123 = BLOCKFINU%
   '
   On Error GoTo 0
   Call GRASECUBLOQ(FINDEX%, TEBLO$)
   TEBLOFI$(FINDEX%) = TEBLO$
   '
   Screen.MousePointer = SCA
   Exit Sub                     ' bloqueo exitoso
   '
'  **************************************************
   '
15 Resume 16
16 On Error GoTo 0
   Close #LOCKFINU%: LOCKFINU% = 0: FINULOCK%(FINDEX%) = 0
   '
18 Call LIBLOCKANT          ' libera bloqueos anteriores
   CAINTEN1% = CAINTEN1% + 1
   If CAINTEN1% > 3 Then GoTo 40
   TTT1 = Timer: While Timer < TTT1 + 1: Wend ' espera 1 segundos
   '
20 Call OPENBLOCKLIST
   On Error GoTo 30
   For KKII% = 1 To CANARCHIV%  ' re-bloquea anteriores
     If FINULOCK%(KKII%) = (-1) Then
       NK% = FreeFile
       ARTX$ = FILOCKNAME$(KKII%)
       Open ARTX$ For Random Access Read Write Lock Write As #NK% Len = 128
       FINULOCK%(KKII%) = NK%
     End If
   Next KKII%
   On Error GoTo 0
   GoTo 13
   '
30 Resume 31
31 On Error GoTo 0
   If CAINTEN2% < 5 Then
     CAINTEN2% = CAINTEN2% + 1
     GoTo 18
   End If
   '
40 Call MENYABLOQ(FINDEX%, REINTEN%)   ' mensaje en uso exclusivo
   If REINTEN% = 1 Then GoTo 12
   Call CIERRARCH("*.*")
   Close: End
   '
End Sub
'
Sub MENYABLOQ(FINDEX%, REINTEN%)
   '
   REINTEN% = 0
   Call CLOSEBLOCKLIST
   Call LIBLOCKANT
   '
   Dim AA As String * 128
   USZ$ = "Desconocido"
   TRZ$ = "Desconocida"
   PRZ$ = "Desconocido"
   FFZ1$ = "00-00-00": FFZ2$ = "00:00:00"
   '
   ARX$ = FILOCKNAME$(FINDEX%)
   On Error Resume Next
   NKX% = FreeFile
   Open ARX$ For Random Access Read Shared As #NKX% Len = 128
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
   ARY$ = REPLACAR$(FISHORNAME$(FINDEX%), "\", "/")
   '
   COMU$ = "ADVERTENCIA: Archivo '" + TRIM$(ARY$) + "' en Uso Exclusivo !!!\-\" + USY$ + "\" + TRY$ + "\" + FFY$ + "\" + PRY$ + "\\Intentar de Nuevo\Abandonar Tarea"
   If COMALTER%(COMU$) = 1 Then
        REINTEN% = 1
        Screen.MousePointer = 11
      Else
        Call CIERRARCH("*.*")
        Call FINAL("")
   End If
   '
End Sub
'
Sub LIBLOCKANT()    ' libera bloqueos anteriores
   '
   CANARCHIV% = Len(FINAMSTRI$) / 16
   For KKII% = 1 To CANARCHIV%
     If FINULOCK%(KKII%) > 0 Then
       If FINUMBER%(KKII%) > 0 Then
         Close #FINUMBER%(KKII%): FINUMBER%(KKII%) = 0
       End If
       Close #FINULOCK%(KKII%)
       FINULOCK%(KKII%) = (-1)
     End If
   Next KKII%
   '
   Call CLOSEBLOCKLIST
   '
End Sub
'
Sub LIBERFILE(FINDEX%)
   '
12 TACCE = Timer
   SCA = Screen.MousePointer
   If SCA < 1 Then SCA = 1
   Screen.MousePointer = 11
   '
13 LOCKFINU% = FINULOCK%(FINDEX%)
   If LOCKFINU% > 0 Then
     If FINUMBER%(FINDEX%) > 0 Then  ' CIERRA ARCHIVO ANTES DE LIBERAR
        Close #FINUMBER%(FINDEX%): FINUMBER%(FINDEX%) = 0
     End If
     On Error Resume Next
     Get #LOCKFINU%, 1, BUFLOCKFILE$
     If Err = 52 Then
          On Error GoTo 0
          GoTo 15
        ElseIf Err > 0 Then
          MsgBox "Error " & Err.Number & " en Rutina LIBERFILE"
          On Error GoTo 0
          Close: End
     End If
     '
     TEBLO$ = BUFLOCKFILE$
     Call REPLA(TEBLO$, Space$(116), 13, 116)
     BUFLOCKFILE$ = TEBLO$
     Put #LOCKFINU%, 1, BUFLOCKFILE$
     Close #LOCKFINU%
15   FINULOCK%(FINDEX%) = 0
     '
     Get #BLOCKFINU%, RELILOCK%(FINDEX%), BUFLOCKLIST$
     B1$ = TRIM$(UCase$(Left$(BUFLOCKLIST$, 12)))
     If InStr(B1$, ".") < 2 Then B1$ = B1$ + ".DAT"
     If B1$ <> TRIM$(FILOCKTEXT$(FINDEX%)) Then
        MsgBox ("Error de Semáforo de Bloqueos - Re-Inicie FLEXOFT.")
        Close: Reset: End
     End If
     LSet BUFLOCKLIST$ = TEBLO$
     Put #BLOCKFINU%, RELILOCK%(FINDEX%), BUFLOCKLIST$
     '
     On Error GoTo 0
     Call GRASECUBLOQ(FINDEX%, TEBLO$)
   End If
   '
   If SCA < 1 Then SCA = 1
   Screen.MousePointer = SCA
   '
End Sub
'
Sub GRASECUBLOQ(FINDEX%, TEBLO$)
   '
   Dim CC As String * 128
   TESECU$ = Space$(128)
   USX$ = USERNAME$
   NTX$ = NOMTER$
   HOII% = HOY(HO$)
   Call REPLA(TESECU$, Left$(TEBLO$, 12), 1, 12)
   Call REPLA(TESECU$, Mid$(TEBLO$, 101, 8), 15, 8)
   Call REPLA(TESECU$, Mid$(TEBLO$, 75, 20), 25, 20)
   Call REPLA(TESECU$, USX$, 65, 24)
   Call REPLA(TESECU$, NTX$, 90, 24)
   '
   N2% = FreeFile
   Open LUDAT$ + "SECUBLOQ.DAT" For Random Access Read Write As #N2% Len = 128
   If REGLOCK&(FINDEX%) < 1 Then
        RGXX& = 1 + LOF(N2%) / 128
        Get #N2%, RGXX&, CC$
        REGLOCK&(FINDEX%) = RGXX&
     Else
        RGXX& = REGLOCK&(FINDEX%)
        Get #N2%, RGXX&, CC$
        TESECU$ = CC$
        Call REPLA(TESECU$, HO$, 45, 8)
        Call REPLA(TESECU$, Time$, 55, 8)
        REGLOCK&(FINDEX%) = 0
   End If
   Call REPLA(TESECU$, Chr$(10), 128, 1)
   LSet CC$ = TESECU$
   Put #N2%, RGXX&, CC$
   Close #N2%: N2% = 0
   '
End Sub
'

Sub COPYSTRU(A$, B$)
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
    ORI$ = LOADFILE$(LUDAT$ + X$ + ".RFS")
    NUE$ = LOADFILE$(LUDAT$ + Y$ + ".RFS")
    If NUE$ <> ORI$ Then
       Call SAVEFILE(LUDAT$ + Y$ + ".RFS", ORI$)
    End If
    '
End Sub
'


Sub HEADERS(Archivo$, Texto$)
    '
    LU$ = LUDAT$
    ARCH$ = UCase$(AJUSTI$(Archivo$, 8))
    '
200 If Texto$ = "" Then

            For i& = 1 To ULTREG&("HEADERS")
                XXX$ = REGLEIDO$("HEADERS", i&)
                Archi$ = Left$(XXX$, 8)
                If Archi$ <> ARCH$ Then
                  If Archi$ <> Space$(8) Then
                    If Archi$ <> String$(8, 0) Then
                        J& = J& + 1
                        Call GRAREG("HEADERS", XXX$, J&)
                    End If
                  End If
                End If
            Next i&

            For i& = J& + 1 To ULTREG&("HEADERS")
                Call GRAREG("HEADERS", String$(128, 0), i&)
            Next i&
            '
            Call CIERRARCH("HEADERS")
            Exit Sub
            '
        Else
            '
            For i& = 1 To ULTREG&("HEADERS")
                XXX$ = REGLEIDO$("HEADERS", i&)
                Archi$ = Left$(XXX$, 8)
                'If ARCHI$ <> ARCH$ Then
                  If Archi$ <> Space$(8) Then
                    If Archi$ <> String$(8, 0) Then
                      J& = J& + 1
                      Call GRAREG("HEADERS", XXX$, J&)
                    End If
                  End If
                'End If
            Next i&
            '
            XXX$ = ARCH$ + Texto$
            Call GRAREG("HEADERS", XXX$, J& + 1)
            '
            For i& = J& + 2 To ULTREG&("HEADERS")
                Call GRAREG("HEADERS", String$(128, 0), i&)
            Next i&
            '
            Call CIERRARCH("HEADERS")
    End If
    '
End Sub
'
Sub ENCACOL(Archivo$, Texto$)
    '
    LU$ = LUDAT$
    ARCH$ = UCase$(AJUSTI$(Archivo$, 8))
    '
    NXX% = FILEOPEN%(LUCOM$ + Archivo$ + ".HDP", 0, 256)
    Dim TENCA As String * 256
    '
200 If Texto$ = "" Then
        Close #NXX%
        On Error Resume Next
        Kill LUCOM$ + Archivo$ + ".HDP"
        On Error GoTo 0
       Else
        REG& = 1 + LOF(NXX%) / 256
        TENCA$ = Left$(Texto$ + Space$(256), 256)
        Put #NXX%, REG&, TENCA$
        Close #NXX%
    End If
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
Sub EXTRACT(FINAME$)   ' vigente desde 10/03/2002
   '
   Dim BUFLE As String * 160
   '
   SCMA = Screen.MousePointer
   Screen.MousePointer = 11
   AREX$ = TRIM$(UCase$(FINAME$))
   If InStr(AREX$, ".") < 2 Then AREX$ = AREX$ + ".RFS"
   If Left$(AREX$, 1) = "!" Then AREX$ = Mid$(AREX$, 2)
   '
   If InStr(AREX$, ".PRF") > 0 Then
      If VERDEMO% = 1 Then Exit Sub
      While InStr(AREX$, "\") > 0
         AREX$ = Mid$(AREX$, 2)
      Wend
   End If
   '
   LUDESTI$ = LUDAT$
   If InStr(AREX$, ".PRF") > 0 Then LUDESTI$ = CurDir
   If Right$(LUDESTI$, 1) <> "\" Then LUDESTI$ = LUDESTI$ + "\"
   '
   NX% = FreeFile
   Open LUBAS$ + "FLEXPRFS.FLX" For Random Access Read Shared As #NX% Len = 160
   Fin& = LOF(NX%) / 160
   '
   LI& = 0: LS& = Fin& + 1
   INI& = 0
   ABRE$ = TRIM$(UCase$(AREX$))
51 E& = Int((LS& - LI&) / 2): If E& = 0 Then GoTo 58
   L& = LI& + E&: If L& < 1 Or L& > Fin& Then GoTo 58
   Get #NX%, L&, BUFLE$
   I1$ = TRIM$(UCase$(Left$(BUFLE$, 32)))
   If I1$ < ABRE$ Then LI& = L&: GoTo 51
   If I1$ > ABRE$ Then LS& = L&: GoTo 51
   INI& = L& - 128: If INI& < 1 Then INI& = 1
   '
58 If INI& < 1 Then
     Screen.MousePointer = 1
     Call MENSERR(24, "Imposible Extraer Controlador " + ABRE$)
     GoTo 90
   End If
   '
   BUFGRA$ = ""
   For II& = INI& To Fin&
     Get #NX%, II&, BUFLE$
     I1$ = TRIM$(UCase$(Left$(BUFLE$, 32)))
     If I1$ > AREX$ Then
         GoTo 80
       ElseIf I1$ = AREX$ Then
         BUFGRA$ = BUFGRA$ + Mid$(BUFLE$, 33)
     End If
   Next II&
   '
80 Call SAVEFILE(LUDESTI$ + AREX$, BUFGRA$)
   Close #NX%
   If Control("NEWASPEC", "CAMBIOAS") <> "SI" Then Call COMUNI("Controlador: " + AREX$ + "\Instalación O.K.")

   '
90 Screen.MousePointer = SCMA
   '
End Sub
'
Sub TRALOCAL(Archi$, ALOCA$)
    '
    ARCHX$ = Archi$
    ARCHY$ = TRIM$(ALOCA$)
    If ARCHY$ = "" Then ARCHY$ = TRIM$(ARCHX$)
    Call CIERRARCH("*.*")
    On Error GoTo 10
    FileCopy LUDAT$ + ARCHX$ + ".DAT", LUCOM$ + ARCHY$ + ".WKF"
    On Error GoTo 0
    Exit Sub
    '
10  Resume 11
11  On Error GoTo 0
    Call BLANARCH("!" + ARCHY$)
    Screen.MousePointer = 11
    Fin& = ULTREG&(ARCHX$)
    For i& = 1 To Fin&
       Call TRACE(20, i&, Fin&)
       X$ = REGLEIDO$(ARCHX$, i&)
       Call GRAREG("!" + ARCHY$, X$, i&)
    Next i&
    Call SETULTREG("!" + ARCHY$, Fin&)
    Call CIERRARCH("!" + ARCHY$)
    Screen.MousePointer = 1
    '
End Sub

Sub TRACENTRAL(Archi$, ALOCA$)
    '
    ARCHX$ = Archi$
    Screen.MousePointer = 11
    ARCHY$ = TRIM$(ALOCA$)
    If ARCHY$ = "" Then ARCHY$ = TRIM$(ARCHX$)
    '
    Call CIERRARCH("*.*")
    Call BLOQARCH(ARCHX$)
    On Error GoTo 10
    FileCopy LUCOM$ + ARCHY$ + ".WKF", LUDAT$ + ARCHX$ + ".DAT"
    On Error GoTo 0
    Call INDEXARCH(ARCHX$, 1)
    GoTo 90
    '
10  Resume 11
11  On Error GoTo 0
    J& = 0
    Fin& = ULTREG&("!" + ARCHY$)
    For i& = 1 To Fin&
       Call TRACE(20, i&, Fin&)
       X$ = REGLEIDO$("!" + ARCHY$, i&)
       J& = J& + 1
       Call GRAREG(ARCHX$, X$, J&)
    Next i&
    Call SETULTREG(ARCHX$, J&)
    '
90  Call CIERRARCH(ARCHX$)
    Screen.MousePointer = 1
    '
End Sub

Sub COMPACTATABLA(Archivo$)
    '
    'VALIDO LA EXISTENCIA DEL ARCHIVO
    If EXISTE%(LUDAT$ + Archivo$ + ".RFS") < 1 Then Exit Sub
    '
    'IF EXISTE EL ARCHIVO LOCAL LO ELIMINO
    On Error Resume Next
    Kill (LUCOM$ + Archivo$ + ".WKF")
    On Error GoTo 0
    '
    'VALIDO SI SE ELIMINO
    If EXISTE%(LUCOM$ + Archivo$ + ".WKF") > 0 Then
       Call COMUNI("Imposible Compactar Archivo (1): " & Archivo$ & " \Consulte a su Soporte de Sistema.")
       Exit Sub
    End If
    '
    'RECORRO EL ARCHIVO
    For i& = 1 To ULTREG&(Archivo$)
        TX$ = REGLEIDO$(Archivo$, i&)
        Call REGAPP("!" & Archivo$, TX$)
    Next i&
    Call CIERRARCH(Archivo$)
    Call CIERRARCH("!" & Archivo$)
    '
    'TRANSFIERO LA INOFRMACION
    ARCHILOC$ = LUCOM$ + Archivo$ + ".WKF"
    ARCHIDEST$ = LUDAT$ + Archivo$ + ".DAT"
    On Error Resume Next
    FileCopy ARCHILOC$, ARCHIDEST$
    If Err > 0 Then
       On Error GoTo 0
       Call COMUNI("Imposible Compactar Archivo (2): " & Archivo$ & " \Consulte a su Soporte de Sistema.")
       Exit Sub
    End If
    On Error GoTo 0
    '
    Call FRESHECHO(Archivo$)
    '
End Sub

'
Sub TRALOFIL(Archi$, X%, ARGU$)
    '
    SCMA = Screen.MousePointer
    Screen.MousePointer = 11
    ARCHX$ = Archi$
    ARCHY$ = "!" + ARCHX$
    JJ1& = 0
    RFF$ = RECFILT$(ARCHX$, X%, ARGU$)
    For ULU& = 1 To Len(RFF$) Step 4
      RULU& = CVS(Mid$(RFF$, ULU&, 4))
      XX1$ = REGLEIDO$(ARCHX$, RULU&)
      JJ1& = JJ1& + 1
      Call GRAREG(ARCHY$, XX1$, JJ1&)
    Next ULU&
    Call SETULTREG(ARCHY$, JJ1&)
    Call CIERRARCH(ARCHY$)
    Screen.MousePointer = SCMA
    '
End Sub

Sub TRACENFIL(Archi$, X%, ARGU$)
    '
    SCMA = Screen.MousePointer
    Screen.MousePointer = 11
    ARCHX$ = Archi$
    ARCHY$ = "!" + ARCHX$
    POFII% = POSFIELD%(ARCHX$, X%)
    If POFII% < 1 Then Exit Sub
    LOFII% = LENFIELD%(ARCHX$, X%)
    LOREI% = RECLEN%(ARCHX$)
    '
    URN& = ULTREG&(ARCHX$)
    RFF$ = RECFILT$(ARCHX$, X%, ARGU$)
    KK& = (-3)
    '
    For ULU& = 1 To ULTREG&(ARCHY$)
      XX$ = REGLEIDO$(ARCHY$, ULU&)
      Call REPLA(XX$, ARGU$, POFII%, LOFII%)
10    KK& = KK& + 4
      If KK& <= Len(RFF$) - 3 Then
           RENOTA& = CVS(Mid$(RFF$, KK&, 4))
           XY$ = REGLEIDO$(ARCHX$, RENOTA&)
           If Mid$(XY$, POFII%, LOFII%) <> ARGU$ Then GoTo 10
         Else
           URN& = URN& + 1
           RENOTA& = URN&
      End If
      Call GRAREG(ARCHX$, XX$, RENOTA&)
   Next ULU&
   '
20 While KK& < Len(RFF$) - 3
      KK& = KK& + 4
      RENOTA& = CVS(Mid$(RFF$, KK&, 4))
      Call GRAREG(ARCHX$, String$(LOREI%, 0), RENOTA&)
   Wend
   '
   Call CIERRARCH(ARCHX$)
   Screen.MousePointer = SCMA
   '
End Sub
'
Function BLOQRESULT%(Archivo$)
   '
   MODO% = 0
   'Call BLOQUEAR(ARCHIVO$, MODO%)
   BLOQRESULT% = MODO%
   '
End Function
'

Sub CIERRARCH(Archi$)
   '
   If TRIM$(Archi$) <> "" Then
     If Left$(Archi$, 1) <> "*" Then
         FINDEX% = FILEINDEX%(Archi$)
         Call INDEXARCH(Archi$, PRIREMOD&(FINDEX%))
         FINUME% = FILENBR%(Archi$)
         Close #FINUME%: FINUMBER%(FINDEX%) = 0
         PRIREMOD&(FINDEX%) = 0
         UREGIST&(FINDEX%) = 0
         UREGLEI&(FINDEX%) = 0
         RELEIDO$(FINDEX%) = String$(LOREGIS%(FINDEX%), 0)
         Call LIBARCH(Archi$)
       Else
         For KKI% = 1 To 128
           If FINUMBER%(KKI%) > 0 Then
             Call INDEXARCH(FISHORNAME$(KKI%), PRIREMOD&(KKI%))
             Close #FINUMBER%(KKI%): FINUMBER%(KKI%) = 0
             PRIREMOD&(KKI%) = 0
             UREGIST&(KKI%) = 0
             UREGLEI&(KKI%) = 0
             RELEIDO$(KKI%) = String$(LOREGIS%(KKI%), 0)
           End If
         Next KKI%
         ' Call LIBARCH("*.*")   SUPRIMIDO - si cierra todos los archivos
         '                       la liberacion debe ser explicita
     End If
     '
     ' RECUPERA SALDOS DE STOCK
     If TRIM$(UCase$(Archi$)) = "MOVISTO" Then
        SMCA = Screen.MousePointer
        Screen.MousePointer = 11
        NMSX& = NUMAS%
        For KKIL& = 1 To NMSX&
          Call TRACE(20, KKIL&, NMSX&)
          KKII% = KKIL&
          If FLAGMOVI%(KKII%) > 0 Then
             Call VESARRAS(KKII%)
             FLAGMOVI%(KKII%) = 0
          End If
        Next KKIL&
        Screen.MousePointer = SMCA
     End If
     '
   End If
   '
End Sub

Sub INDEXARCH(ARCHX$, PRIREG&)
   '
Call INDEXARCHNUE(ARCHX$, PRIREG&)
Exit Sub

   PRXXX& = PRIREG&
   ARCHJ$ = ARCHX$
   If Left$(ARCHX$, 1) = "!" Then Exit Sub
   '
   If PRXXX& > 0 Then
     UREAR& = ULTREG&(ARCHX$)
     If PRIREG& <= UREAR& Then
       FINDEX% = FILEINDEX%(ARCHJ$)
       ISTRI$ = IDXSTRI$(FINDEX%)
       '
       For KKI% = 1 To Len(ISTRI$)
         MARCX% = Asc(Mid$(ISTRI$, KKI%, 1))
         If MARCX% = 1 Then
           Lx% = LENFIELD%(ARCHX$, KKI%)
           PX% = POSFIELD%(ARCHX$, KKI%)
           If Lx% > 0 Then       ' SOLO SI LONGITUD DEL FILTRO =2
             '
             EXTX$ = TRIM$(Str$(KKI%))
             While Len(EXTX$) < 2: EXTX$ = "0" + EXTX$: Wend
             IFANAME$ = LUDAT$ + ARCHJ$ + ".X" + EXTX$
             '
             XXXX$ = LOADFILE$(IFANAME$)
             If PRXXX& > Int(Len(XXXX$) / Lx%) Then
               PRXXX& = 1 + Int(Len(XXXX$) / Lx%)
             End If
             XXXX$ = XXXX$ + String$(Lx% * UREAR&, 0)
             XXXX$ = Left$(XXXX$, Lx% * UREAR&)
             '
             FINUMX% = FILENBR%(ARCHX$)
             LOREGX% = LOREGIS%(FINDEX%)
             For REAC1& = PRXXX& To UREAR&
               'XYZT$ = REGLEIDO$(ARCHX$, REAC1&)
               XYZT$ = LEEREGIS$(FINDEX%, FINUMX%, LOREGX%, REAC1&)
               ARIN$ = Mid$(XYZT$, PX%, Lx%)
               Mid$(XXXX$, 1 + Lx% * (REAC1& - 1), Lx%) = ARIN$
             Next REAC1&
             '
             Call SAVEFILE(IFANAME$, XXXX$)
             '
           End If
         End If
       Next KKI%
       PRIREMOD&(FINDEX%) = 0
     End If
   End If
   '
End Sub
'
Sub INDEXARCHNUE(ARCHX$, PRIREG&)
   '
   Dim LONGY%(100), POSFY%(100)
   Dim IFANAME$(100), XXXX$(100)
   '
   PRXXX& = PRIREG&
   ARCHJ$ = ARCHX$
   If Left$(ARCHX$, 1) = "!" Then Exit Sub
   '
   If PRXXX& > 0 Then
     UREAR& = ULTREG&(ARCHX$)
     If PRXXX& <= UREAR& Then
       '
       ' cargar posicion y longitud campos indexados
       FINDEX% = FILEINDEX%(ARCHJ$)
       ISTRI$ = IDXSTRI$(FINDEX%)
       '
       CACAM% = Len(ISTRI$)
       If CACAM% > 99 Then CACAM% = 99
       '
       HAYINDEX% = 0
       For KKI% = 1 To CACAM%
         '
         MARCX% = Asc(Mid$(ISTRI$, KKI%, 1))
         LONGY%(KKI%) = 0
         POSFY%(KKI%) = 0
         '
         If MARCX% = 1 Then
           LONGY%(KKI%) = LENFIELD%(ARCHX$, KKI%)
           POSFY%(KKI%) = POSFIELD%(ARCHX$, KKI%)
           If POSFY%(KKI%) < 1 Then LONGY%(KKI%) = 0
           '
           If LONGY%(KKI%) > 0 Then       ' SOLO SI LONGITUD DEL FILTRO =2
             HAYINDEX% = 1
             EXTX$ = TRIM$(Str$(KKI%))
             While Len(EXTX$) < 2: EXTX$ = "0" + EXTX$: Wend
             IFANAME$(KKI%) = LUDAT$ + ARCHJ$ + ".X" + EXTX$
             '
             On Error Resume Next
             FILEN& = FileLen(IFANAME$(KKI%))
             On Error GoTo 0
             If PRXXX& > Int(FILEN& / LONGY%(KKI%)) Then
               PRXXX& = 1 + Int(FILEN& / LONGY%(KKI%))
             End If
             '
           End If
         End If
       Next KKI%
       '
       If HAYINDEX% < 1 Then Exit Sub
       '
       SCMA = Screen.MousePointer
       Screen.MousePointer = 11
       '
       ' revisar ultimos 128 registros
       FINUMX% = FILENBR%(ARCHJ$)
       LOREGX% = LOREGIS%(FINDEX%)
       VUELREV% = 0
       '
10     VUELREV% = VUELREV% + 1
       If VUELREV% >= 8 Then PRXXX& = 1
       If PRXXX& <= 1 Then GoTo 20
       '
       PRYYY& = PRXXX& - 128
       If PRYYY& < 1 Then PRYYY& = 1
       '
       For KKI% = 1 To CACAM%
         XXXX$(KKI%) = ""
         If LONGY%(KKI%) > 0 Then
           POSINLEC& = 1 + LONGY%(KKI%) * (PRYYY& - 1)
           XXXX$(KKI%) = LOADSUBFILE$(IFANAME$(KKI%), POSINLEC&)
         End If
       Next KKI%
       '
       For REAC1& = PRYYY& To PRXXX& - 1
         XYZT$ = LEEREGIS$(FINDEX%, FINUMX%, LOREGX%, REAC1&)
         '
         For KKI% = 1 To CACAM%
           If LONGY%(KKI%) > 0 Then
             ARIN$ = Mid$(XYZT$, POSFY%(KKI%), LONGY%(KKI%))
             If Mid$(XXXX$(KKI%), 1 + LONGY%(KKI%) * (REAC1& - PRYYY&), LONGY%(KKI%)) <> ARIN$ Then
               PRXXX& = REAC1&
               GoTo 10
             End If
           End If
         Next KKI%
         '
       Next REAC1&
       '
20     For KKI% = 1 To CACAM%
         XXXX$(KKI%) = ""
         If LONGY%(KKI%) > 0 Then
           XXXX$(KKI%) = String$((1 + UREAR& - PRXXX&) * LONGY%(KKI%), 0)
         End If
       Next KKI%
       '
       For REAC1& = PRXXX& To UREAR&
         XYZT$ = LEEREGIS$(FINDEX%, FINUMX%, LOREGX%, REAC1&)
         '
         For KKI% = 1 To CACAM%
           If LONGY%(KKI%) > 0 Then
             ARIN$ = Mid$(XYZT$, POSFY%(KKI%), LONGY%(KKI%))
             Mid$(XXXX$(KKI%), 1 + LONGY%(KKI%) * (REAC1& - PRXXX&), LONGY%(KKI%)) = ARIN$
           End If
         Next KKI%
         '
       Next REAC1&
       '
       For KKI% = 1 To CACAM%
         If LONGY%(KKI%) > 0 Then
           FINUIX% = FILEOPEN%(IFANAME$(KKI%), 4, 128)
           POSGRA& = 1 + (PRXXX& - 1) * LONGY%(KKI%)
           Put #FINUIX%, POSGRA&, XXXX$(KKI%)
           Close #FINUIX%
         End If
       Next KKI%
       '
       PRIREMOD&(FINDEX%) = 0
       Screen.MousePointer = SCMA
       '
     End If
   End If
   '
End Sub
'
Function ECOARCH$(ARCHX$, ARGU$)
   '
   ECA$ = ""
   ARGULO$ = ARGU$
   ARCHX$ = UCase$(ARCHX$)
   If TRIM$(ARCHX$) = "MASTER" Or TRIM$(ARCHX$) = "MASTER0" Then
       If Len(ARGULO$) <> 2 Then ARGULO$ = MKI$(CODINT%(ARGULO$))
       If COLTELIB% < 1 Then
           ECA$ = DESCRIT$(CVI(ARGULO$))
         Else
           ECA$ = DESCRIT0$(CVI(ARGULO$))
       End If
       If TRIM$(ARCHX$) = "MASTER0" Then
         If ARGULO$ = MKI$(0) Then
           ECA$ = ""
         End If
       End If
       GoTo 99
     ElseIf TRIM$(UCase$(ARCHX$)) = "DEUDOR1" Then
       ECA$ = rasocli$(CVI(ARGULO$))
       If TRIM$(ECA$) = "" Then
          If IsNumeric(ARGULO$) = True Then
             ECA$ = rasocli$(XVALO(ARGULO$))
          End If
       End If
       GoTo 99
     ElseIf TRIM$(UCase$(ARCHX$)) = "PROVED1" Or TRIM$(UCase$(ARCHX$)) = "ACREEDOR" Then
       ECA$ = rasocli$((-1) * CVI(ARGULO$))
       If TRIM$(ECA$) = "" Then
          If IsNumeric(ARGULO$) = True Then
             ECA$ = rasocli$((-1) * XVALO(ARGULO$))
          End If
       End If
       GoTo 99
   End If
   '
   ARCHJ$ = TRIM$(ARCHX$)
   If Left$(ARCHJ$, 1) = "^" Then
       ARCHJ$ = Mid$(ARCHJ$, 2)
     ElseIf Left$(ARCHJ$, 1) = "*" Then
       FRXX$ = Mid$(ARCHJ$, 2) + ".PRF"
       'If VERDEMO% = 1 Then
       '  FRXX$ = "\FX_PROFILS\" + FRXX$
       'End If
       ARCHJ$ = TRIM$(Mid$(LOADFILE$(FRXX$), 57, 8))
       If ARCHJ$ = "" Then GoTo 99
   End If
   '
   VUELTA% = 0
14 RFF$ = RECFILT$(ARCHJ$, 1, ARGULO$)
   If RFF$ = "" Then
      If VUELTA% = 0 Then
        If Len(ARGULO$) <> LENFIELD%(ARCHJ$, 1) Then
           ARGULO$ = CBIN$(ARGULO$, TIPFIELD%(ARCHJ$, 1), LENFIELD%(ARCHJ$, 1))
           VUELTA% = 1
           GoTo 14
        End If
      End If
   End If
   '
   If Len(RFF$) >= 4 Then
     REG& = CVS(Left$(RFF$, 4))
     If REG& > 0 Then
       If REG& <= ULTREG&(ARCHJ$) Then
         XYZT$ = REGLEIDO$(ARCHJ$, REG&)
         Lx% = LENFIELD%(ARCHJ$, 2)
         PX% = POSFIELD%(ARCHJ$, 2)
         ECA$ = Mid$(XYZT$, PX%, Lx%)
         Call SETREG(ARCHJ$, REG&)
       End If
     End If
   End If
   '
99 ECOARCH$ = ECA$
   '
End Function
'
Function RECFILT$(Archi$, X%, ARGU$)
   '
   Lx% = LENFIELD%(Archi$, X%)
   AZ$ = Left$(ARGU$ + Space$(Lx%), Lx%)
   '
   ARCHJ$ = Archi$
   If Left$(Archi$, 1) = "!" Then
     ARCHJ$ = Mid$(Archi$, 2)
     GoTo 50                    ' LECTURA SECUENCIAL
   End If
   '
10 IX% = IDXFIELD%(ARCHJ$, X%)
   If IX% = 1 Then
     If Lx% > 0 Then       ' SOLO SI LONGITUD DEL FILTRO =2
       '
       FINDEX% = FILEINDEX%(Archi$)
       If XFIELD%(FINDEX%) = X% Then
         If FIDATIM#(Archi$) = ULTACTIDX#(FINDEX%) Then
           UREAR& = ULTREG&(Archi$)
           GoTo 16
         End If
       End If
       '
15     UREAR& = ULTREG&(Archi$)
       LU$ = LUDAT$
       EXTX$ = TRIM$(Str$(X%)): While Len(EXTX$) < 2: EXTX$ = "0" + EXTX$: Wend
       XSTRIN$(FINDEX%) = Left$(LOADFILE$(LUDAT$ + ARCHJ$ + ".X" + EXTX$), Lx% * UREAR&)
       XFIELD%(FINDEX%) = X%
       ULTACTIDX#(FINDEX%) = FIDATIM#(Archi$)
       '
16     ' RGLIM& MARCA EL REGISTRO A PARTIR DEL CUAL SE HACE CONTROL CRUZADO
       ' ENTRE EL ÍNDICE Y EL PRINCIPAL - E.P.B. - 1/05/2007
       RGLIM& = 0.95 * UREAR&
       If RGLIM& > UREAR& - 2048 Then RGLIM& = UREAR& - 2048
       '
       If Len(XSTRIN$(FINDEX%)) / Lx% < UREAR& Then
         REAC1& = Len(XSTRIN$(FINDEX%)) / Lx% - 64: If REAC1& < 1 Then REAC1& = 1
         If REAC1& <= UREAR& Then
           Call INDEXARCH(Archi$, REAC1&)
           GoTo 15
         End If
       End If
       '
       PX% = POSFIELD%(Archi$, X%)
       FINUMX% = FILENBR%(Archi$)
       LOREGX% = LOREGIS%(FINDEX%)
18     RX$ = String$(4096, 0): PINI& = 1: POIX& = 1
20     POSI& = InStr(PINI&, XSTRIN$(FINDEX%), AZ$)
       If POSI& < 1 Then GoTo 25                           ' NEXT RGX&
       If Lx% > 1 Then
         If (POSI& Mod Lx%) <> 1 Then
           PINI& = POSI& + 1
           GoTo 20
         End If
       End If
       PINI& = POSI& + Lx%
       RGL& = 1 + (POSI& - 1) / Lx%
       '
       If RGL& >= RGLIM& Then
          ' SE HACE CONTROL CRUZADO CON EL PRINCIPAL
          XYZT$ = LEEREGIS$(FINDEX%, FINUMX%, LOREGX%, RGL&)
          If Mid$(XYZT$, PX%, Lx%) <> AZ$ Then
            REAC1& = RGL& - 128
            If REAC1& < 1 Then REAC1& = 1
            Call INDEXARCH(Archi$, REAC1&)
            GoTo 15
          End If
       End If
       '
       If POIX& > Len(RX$) Then RX$ = RX$ + String$(Len(RX$), 0)
       Mid$(RX$, POIX&, 4) = MKS$(RGL&)
       POIX& = POIX& + 4
       GoTo 20
       '
25     RX$ = Left$(RX$, POIX& - 1)
       GoTo 90
       '
     End If
   End If
   '
   ' LECTURA SECUENCIAL
50 PX% = POSFIELD%(Archi$, X%)
   Lx% = LENFIELD%(Archi$, X%)
   '
   FINDEX% = FILEINDEX%(Archi$)
   If Lx% > 0 Then
     If X% <= Len(IDXSTRI$(FINDEX%)) Then
       Mid$(IDXSTRI$(FINDEX%), X%) = Chr$(1)
       Call INDEXARCH(Archi$, 1)
       GoTo 10
     End If
   End If
   '
   FINUMX% = FILENBR%(Archi$)
   LOREGX% = LOREGIS%(FINDEX%)
   If PX% > 0 Then
     If Lx% > 0 Then
       For RGL& = ULTREG&(Archi$) To 1 Step -1
         XYZT$ = LEEREGIS$(FINDEX%, FINUMX%, LOREGX%, RGL&)
         If Mid$(XYZT$, PX%, Lx%) = AZ$ Then
           RX$ = MKS$(RGL&) + RX$
         End If
       Next RGL&
     End If
   End If
   '
90 Call SETREG(Archi$, REG&)
   RECFILT$ = RX$
   '
End Function
'
Function INDEXSTRI$(Archi$, X%)
   '
   IXS$ = ""
   Lx% = LENFIELD%(Archi$, X%)
   If Lx% < 1 Then GoTo 90     ' SALE DE FUNCION
   '
   ARCHJ$ = Archi$
   If Left$(Archi$, 1) = "!" Then
     ARCHJ$ = Mid$(Archi$, 2)
     GoTo 50                    ' LECTURA SECUENCIAL
   End If
   '
10 IX% = IDXFIELD%(ARCHJ$, X%)
   If IX% = 1 Then
     FINDEX% = FILEINDEX%(Archi$)
     If XFIELD%(FINDEX%) = X% Then
       If FIDATIM#(Archi$) = ULTACTIDX#(FINDEX%) Then
         GoTo 16
       End If
     End If
     '
15   UREAR& = ULTREG&(Archi$)
     LU$ = LUDAT$
     EXTX$ = TRIM$(Str$(X%)): While Len(EXTX$) < 2: EXTX$ = "0" + EXTX$: Wend
     XSTRIN$(FINDEX%) = Left$(LOADFILE$(LUDAT$ + ARCHJ$ + ".X" + EXTX$), Lx% * UREAR&)
     XFIELD%(FINDEX%) = X%
     ULTACTIDX#(FINDEX%) = FIDATIM#(Archi$)
16   If Len(XSTRIN$(FINDEX%)) / Lx% < UREAR& Then
       REAC1& = Len(XSTRIN$(FINDEX%)) / Lx% - 64
       If REAC1& < 1 Then REAC1& = 1
       If REAC1& <= UREAR& Then
         Call INDEXARCH(Archi$, REAC1&)
         GoTo 15
       End If
     End If
     '
     IXS$ = XSTRIN$(FINDEX%)
     GoTo 90
     '
   End If
   '
   ' LECTURA SECUENCIAL
50 PX% = POSFIELD%(Archi$, X%)
   Lx% = LENFIELD%(Archi$, X%)
   '
   FINDEX% = FILEINDEX%(Archi$)
   If Lx% > 0 Then
     If X% <= Len(IDXSTRI$(FINDEX%)) Then
       Mid$(IDXSTRI$(FINDEX%), X%) = Chr$(1)
       GoTo 10
     End If
   End If
   '
   FINUMX% = FILENBR%(Archi$)
   LOREGX% = LOREGIS%(FINDEX%)
   IXS$ = ""
   If PX% > 0 Then
     If Lx% > 0 Then
       For RGL& = 1 To ULTREG&(Archi$)
         XYZT$ = LEEREGIS$(FINDEX%, FINUMX%, LOREGX%, RGL&)
         IXS$ = IXS$ + Mid$(XYZT$, PX%, Lx%)
       Next RGL&
     End If
   End If
   '
90 INDEXSTRI$ = IXS$
   '
End Function
'
Function TIPFIELD%(Archi$, A%)
   '
   FINDEX% = FILEINDEX%(Archi$)
   VECTIP$ = TVASTRI$(FINDEX%)
   If VECTIP$ = "" Then
      Call LEESTRUCARCH(FINDEX%)
      VECTIP$ = TVASTRI$(FINDEX%)
   End If
   On Error Resume Next
   TIPFIELD% = Asc(Mid$(VECTIP$, A%))
   On Error GoTo 0
   '
End Function
'
Function LENFIELD%(Archi$, A%)
   '
   FINDEX% = FILEINDEX%(Archi$)
   VECTIP$ = LONSTRI$(FINDEX%)
   If VECTIP$ = "" Then
      Call LEESTRUCARCH(FINDEX%)
      VECTIP$ = LONSTRI$(FINDEX%)
   End If
   On Error Resume Next
   LENFIELD% = Asc(Mid$(VECTIP$, A%))
   On Error GoTo 0
   '
End Function
'
Sub SETLENFIELD(Archi$, A%, LONGI%)
   LOANTE% = LENFIELD%(Archi$, A%)
   If LONGI% = LOANTE% Then Exit Sub
   '
   DIFE% = 0: If LONGI% > LOANTE% Then DIFE% = LONGI% - LOANTE%
   AAXX$ = LOADFILE$(LUDAT$ + Archi$ + ".RFS")
   For i% = A% To ((Len(AAXX$) - 128) / 128)
      BBXX$ = Mid$(AAXX$, 128 * i% + 1, 128)
      POAX% = Asc(Mid$(BBXX$, 35, 1))
      LOAX% = Asc(Mid$(BBXX$, 36, 1))
      If i% = A% Then
          LOAX% = LONGI%
        Else
          POAX% = POAX% + DIFE%
      End If
      Call REPLA(BBXX$, Chr$(POAX%), 35, 1)
      Call REPLA(BBXX$, Chr$(LOAX%), 36, 1)
      Mid$(AAXX$, 128 * i% + 1, 128) = BBXX$
   Next i%
   BBXX$ = Mid$(AAXX$, 1, 128)
   LORE1% = Asc(Mid$(BBXX$, 48, 1))
   LORE2% = Asc(Mid$(BBXX$, 49, 1)) + DIFE%
   If LORE2% > 255 Then
      LORE1% = LORE1% + 1
      LORE2% = LORE2% - 256
   End If
   Call REPLA(BBXX$, Chr$(LORE1%), 48, 1)
   Call REPLA(BBXX$, Chr$(LORE2%), 49, 1)
   Mid$(AAXX$, 1, 128) = BBXX$
   Call SAVEFILE(LUDAT$ + Archi$ + ".RFS", AAXX$)
   Call FRESHECHO(Archi$)
End Sub
'
Function POSFIELD%(Archi$, A%)
   '
   
   FINDEX% = FILEINDEX%(Archi$)
   VECTIP$ = POSSTRI$(FINDEX%)
   If VECTIP$ = "" Then
      Call LEESTRUCARCH(FINDEX%)
      VECTIP$ = POSSTRI$(FINDEX%)
   End If
   POSFIELD% = CVI(Mid$(VECTIP$, 2 * A% - 1))
   '
End Function

Function DECFIELD%(Archi$, A%)
   '
   FINDEX% = FILEINDEX%(Archi$)
   VECTIP$ = DECSTRI$(FINDEX%)
   If VECTIP$ = "" Then
      Call LEESTRUCARCH(FINDEX%)
      VECTIP$ = DECSTRI$(FINDEX%)
   End If
   DECFIELD% = Asc(Mid$(VECTIP$, A%))
   '
End Function
'
Sub SETULTREG(Archi$, REG&)
   '
   If Left$(Archi$, 1) = "#" Then
     If ARVEC$ = "" Then ARVEC$ = UCase$(Archi$)
     If ARVEC$ <> UCase$(Archi$) Then
       Call MENSERR(24, "Imposible Grabar EOF de Archivo Virtual '" + Archi$ + "'.\Vector en Uso por Archivo '" + ARVEC$ + "'.\Consulte.")
       Call FINAL("")
     End If
     '
     ReDim Preserve VECREGIS$(REG&)
     Exit Sub
   End If
   '
   '
   LOREGI% = RECLEN%(Archi$)
   REGB$ = String$(LOREG%, 0)
   FINDEX% = FILEINDEX%(Archi$)
   FINUME% = FINUMBER%(FINDEX%)
   If FINUME% < 1 Then FINUME% = FILENBR%(Archi$)
   LOARCHI& = LOF(FINUME%)
   LOBLA& = LOARCHI& - REG& * LOREGI%
   '
   If LOBLA& > 0 Then
     LLR$ = String$(LOBLA&, 0)
     Put #FINUME%, 1 + (REG& * LOREGI%), LLR$
   End If
   Call CIERRARCH(Archi$)
   '
   If PRIREMOD&(FINDEX%) = 0 Or REG& < PRIREMOD&(FINDEX%) Then
     PRIREMOD&(FINDEX%) = REG&
   End If
   '
End Sub
' ***************************************************************************

Sub FILEDUMP(Archi$)
   '
   If Mid$(ARVEC$, 2) <> UCase$(Archi$) Then
     Call MENSERR(24, "Imposible Bajar a Disco Archivo Virtual '" + Archi$ + "'.\Vector en Uso por Archivo '" + ARVEC$ + "'.\Consulte.")
     Call FINAL("")
   End If
   '
   On Error Resume Next
   URE& = UBound(VECREGIS$)
   On Error GoTo 0
   LOREGI% = RECLEN%(Archi$)
   VECSTRI$ = String$(URE& * LOREGI%, 0)
   '
   For KKUL& = 1 To URE&
     Mid$(VECSTRI$, 1 + (KKUL& - 1) * LOREGI%, LOREGI%) = VECREGIS$(KKUL&)
   Next KKUL&
   Call CIERRARCH(Archi$)
   If Left$(Archi$, 1) <> "!" Then
        Call SAVEFILE(LUDAT$ + Archi$ + ".DAT", VECSTRI$)
      Else
        Call SAVEFILE(LUCOM$ + Mid$(Archi$, 2) + ".WKF", VECSTRI$)
   End If
   '
   ReDim VECREGIS$(0)
   ARVEC$ = ""
   '
End Sub
'
Sub FREEMEM(Archi$)
   '
   If Mid$(ARVEC$, 2) <> UCase$(Archi$) Then
     Call MENSERR(24, "Imposible Liberar Memoria de Archivo Virtual '" + Archi$ + "'.\Vector en Uso por Archivo '" + ARVEC$ + "'.\Consulte.")
     Call FINAL("")
   End If
   '
   ReDim VECREGIS$(0)
   ARVEC$ = ""
   '
End Sub
'
Sub BAJALDISCO()
   '
   Call CIERRARCH("*.*")
   Call BLOQEXE("*")
   Call LIBARCH("*.*")
   REFRESHFORM% = 1
   REFRESHFORMDOS% = 1
   Reset
   Call BLOQEXE(App.EXEName)
   '
End Sub

Function IDXFIELD%(ARCHX$, FIEX%)
   '
   XF% = 0
   If FIEX% > 0 Then
     FINDEX% = FILEINDEX%(ARCHX$)
     If FIEX% <= Len(IDXSTRI$(FINDEX%)) Then
       XF% = Asc(Mid$(IDXSTRI$(FINDEX%), FIEX%, 1))
     End If
   End If
   '
   IDXFIELD% = XF%
   '
End Function

Function FILEINDEX%(ARCHX$)
   '
   Static ARCHANT$, FINDEXA%, LUDATA$
   '
   If LUDAT$ <> LUDATA$ Then   ' agregado para posibilitar cambio de empresa en CONTAB04
     FINAMSTRI$ = ""
     FINAMSTRI2$ = ""
     ARCHANT$ = ""
     FINDEXA% = 0
     LUDATA$ = LUDAT$
   End If
   '
   If ARCHX$ = ARCHANT$ Then
     FILEINDEX% = FINDEXA%
     Exit Function
   End If
   '
   ARCHX1$ = TRIM$(ARCHX$)
   '
   If ARCHX1$ <> "" Then
      PPXX% = InStr(FINAMSTRI$, "@@" + ARCHX1$ + " ")
      If PPXX% Mod 16 = 1 Then
         FILEINDEX% = 1 + ((PPXX% - 1) / 16)
         ARCHANT$ = ARCHX$: FINDEXA% = FILEINDEX%
         Exit Function
      End If
      '
      PPXX% = InStr(FINAMSTRI$, Left$("@@" + ARCHX1$ + Space$(14), 14))
      If PPXX% Mod 16 = 1 Then
         FILEINDEX% = 1 + ((PPXX% - 1) / 16)
         ARCHANT$ = ARCHX$: FINDEXA% = FILEINDEX%
         Exit Function
      End If
      '
      If InStr(ARCHX1$, "@@") > 0 Then
         Call MENSERR(24, "Error de Invocación de Función 'FILEINDEX%'.\Nombre de Archivo '@@' no Procesable.\  \Consulte a su Soporte de Sistemas.")
      Close: End
      End If
      '
      PPXX% = InStr(FINAMSTRI$, Left$("@@" + UCase$(ARCHX1$) + Space$(14), 14))
      If PPXX% Mod 16 = 1 Then
         FILEINDEX% = 1 + ((PPXX% - 1) / 16)
         ARCHANT$ = ARCHX$: FINDEXA% = FILEINDEX%
         Exit Function
      End If
      '
      FILEFULLNAME$ = ARCHX1$
      If Left$(ARCHX1$, 1) = "!" Then
          FILEFULLNAME$ = LUCOM$ + Mid$(ARCHX1$, 2) + ".WKF"
        Else
          If Left$(ARCHX1$, 1) <> "\" Then
            If Mid$(ARCHX1$, 2, 2) <> ":\" Then
              FILEFULLNAME$ = LUDAT$ + ARCHX1$ + ".DAT"
            End If
          End If
      End If
      '
      FILOCKNA$ = ""
      PPXX% = InStr(UCase$(FILEFULLNAME$), ".DAT")
      If PPXX% > 1 Then FILOCKNA$ = Left$(FILEFULLNAME$, PPXX% - 1) + ".LCK"
      '
      ARCHX2$ = ARCHX1$
      If Left$(ARCHX2$, 1) = "!" Then
        ARCHX2$ = Mid$(ARCHX1$, 2)
      End If
      FISTRUNA$ = LUDAT$ + ARCHX2$ + ".RFS"
      '
      If Len(FILEFULLNAME$) > 126 Then
         ARCHX2$ = REPLACAR$(FILEFULLNAME$, "\", "/")
         Call MENSERR(24, "Error de Invocación de Función 'FILEINDEX%'.\Archivo '" + ARCHX2$ + "'\no Procesable por Longitud Excesiva.\  \Consulte a su Soporte de Sistemas.")
         Close: End
      End If
      '
      PPXX% = InStr(FINAMSTRI2$, Left$("@@" + UCase$(FILEFULLNAME$) + Space$(126), 126))
      If PPXX% Mod 128 = 1 Then
         FILEINDEX% = 1 + ((PPXX% - 1) / 128)
         ARCHANT$ = ARCHX$: FINDEXA% = FILEINDEX%
         Exit Function
      End If
      '
      ' ESTO ES PARA ATAJAR EL CASO DE QUE SE ABRAN MÁS ARCHIVOS QUE EL
      ' LIMITE DE 128 CORRESPONDIENTE AL DIMENSIONAMIENTO DEL VECTOR.
      ' IMPLEMENTADO POR ERROR DE J & M (ERROR 9) AL 3 / 11/ 08
      If Len(FINAMSTRI$) >= 16 * 120 Then
         Call CIERRARCH("*.*")
         Call BAJALDISCO
         FINAMSTRI$ = ""
         FINAMSTRI2$ = ""
      End If
      ' FIN CORRIGE ERROR DE J & M
      '
      FILESHORTNAME$ = ARCHX1$
      If Len(ARCHX1$) > 14 Then FILESHORTNAME = "**************"
      FINAMSTRI$ = FINAMSTRI$ + AJUSTI$("@@" + FILESHORTNAME$, 16)
      FILEINDEX% = Len(FINAMSTRI$) / 16
      FINAMSTRI2$ = FINAMSTRI2$ + AJUSTI$("@@" + FILEFULLNAME$, 128)
      '
      ' INICIALAR VECTORES
      FINUMBER%(FILEINDEX%) = 0
      FINULOCK%(FILEINDEX%) = 0
      RELILOCK%(FILEINDEX%) = 0
      PRIREMOD&(FILEINDEX%) = 0
      '
      FIFULLNAME$(FILEINDEX%) = FILEFULLNAME$
      FISHORNAME$(FILEINDEX%) = ARCHX1$
      FILOCKNAME$(FILEINDEX%) = FILOCKNA$
      FISTRUNAME$(FILEINDEX%) = FISTRUNA$
      '
      ARCHX1$ = TRIM$(UCase$(FISHORNAME$(FILEINDEX%)))
      PPXX% = InStr(ARCHX1$, ".")
      If PPXX% > 1 Then
         ARCHX1$ = Left$(ARCHX1$, PPXX% - 1)
      End If
      '
      If Left$(ARCHX1$, 1) = "!" Then ARCHX1$ = ""
      FILOCKTEXT$(FILEINDEX%) = Left$(ARCHX1$, 8)
      If FILOCKTEXT$(FILEINDEX%) <> "" Then FILOCKTEXT$(FILEINDEX%) = FILOCKTEXT(FILEINDEX%) + ".DAT"
      '
      FINDEX% = FILEINDEX%
      Call LEESTRUCARCH(FINDEX%)
      ARCHANT$ = ARCHX$: FINDEXA% = FILEINDEX%
      '
      If FILEINDEX% > 128 Then
        Call MENSERR(24, "Demasiados Archivos Abiertos.\Imposible Continuar Ejecución.\  \Consulte a su Soporte de Sistemas.")
        Call CIERRARCH("*.*")
        Call FINAL("")
      End If
      '
      Exit Function
      '
   End If
   '
   Call MENSERR(24, "Error de Invocación de Función 'FILEINDEX%'.\Nombre de Archivo '  ' no Procesable.\  \Consulte a su Soporte de Sistemas.")
   Close: End
   '
End Function

Sub LEESTRUCARCH(FINDEX%)
   '
   Call GRATABCONTRO(FISTRUNAME$(FINDEX%))
   '
If InStr(FISTRUNAME$(FINDEX%), "MONIEXEC") > 0 Then Exit Sub
   VUELTA% = 1
   '
85 FINU% = FILEOPEN%(FISTRUNAME$(FINDEX%), 0, 128)
   Dim XYZ As String * 128
   '
   AAA = LOF(FINU%)
   If LOF(FINU%) <= 128 Then
     XYZ$ = String$(128, 0)
     Put #FINU%, 1, XYZ$
     Close #FINU%: FINU% = 0
     FILEKILL FISTRUNAME$(FINDEX%)
     '
     If VUELTA% = 1 Then
        ARCHEX$ = FISTRUNAME$(FINDEX%)
86      PPXX% = InStr(ARCHEX$, "\")
        If PPXX% > 0 Then ARCHEX$ = Mid$(ARCHEX$, PPXX% + 1): GoTo 86
        PPXX% = InStr(UCase$(ARCHEX$), ".RFS")
        If PPXX% > 0 Then ARCHEX$ = Left$(ARCHEX$, PPXX% - 1)
        Call EXTRACT(ARCHEX$)
        VUELTA% = 2
        GoTo 85
     End If
     '
     ARX1$ = TRIM$(REPLACAR$(FIFULLNAME$(FINDEX%), "\", "/"))
     Call COMUNI("Proceso Interrumpido en Rutina 'FILEMGR'.\  \Falta Controlador de Estructura del Archivo\'" + ARX1$ + "'.")
     Close: Reset: End
   End If
   '
   Get #FINU%, 1, XYZ$
   LOR% = Asc(Mid$(XYZ$, 49, 1))
   LOR1% = Asc(Mid$(XYZ$, 48, 1)): If LOR1% > 16 Then LOR1% = 0
   LOR% = 256 * LOR1% + LOR%
   If LOR% < 1 Then
     ARX1$ = TRIM$(REPLACAR$(FIFULLNAME$(FINDEX%), "\", "/"))
     Call COMUNI("Proceso Interrumpido en Rutina 'FILEMGR'.\  \Error en Longitud de Registro del Archivo\'" + ARX1$ + "'.")
     Close: Reset: End
   End If
   '
   TITU$ = Left$(XYZ$, 47)
   RBL$ = String$(LOR%, 0)
   IXX$ = "": VXX$ = "": PXX$ = "": LXX$ = "": DXX$ = ""
   '
   ARX$ = FISHORNAME$(FINDEX%)
   PSSA% = 0
   For i% = 1 To 128
     If i% + 1 <= LOF(FINU%) / 128 Then
       Get #FINU%, i% + 1, XYZ$
       If XYZ$ <> String$(128, 0) Then
         CAFI% = i%
         TV% = Asc(Mid$(XYZ$, 33)): VXX$ = VXX$ + Chr$(TV%)
         DC% = Asc(Mid$(XYZ$, 34)): DXX$ = DXX$ + Chr$(DC%)
         PS% = Asc(Mid$(XYZ$, 35))
           While PS% < PSSA%: PS% = PS% + 256: Wend
           PXX$ = PXX$ + MKI$(PS%)
           PSSA% = PS%
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
   Next i%
   '
   Close #FINU%: FINU% = 0
   '
   BLANREG$(FINDEX%) = RBL$
   LOREGIS%(FINDEX%) = LOR%
   RELEIDO$(FINDEX%) = String$(LOR%, 0)
   FILETIT$(FINDEX%) = TITU$
   TVASTRI$(FINDEX%) = VXX$
   DECSTRI$(FINDEX%) = DXX$
   POSSTRI$(FINDEX%) = PXX$
   LONSTRI$(FINDEX%) = LXX$
   IDXSTRI$(FINDEX%) = IXX$
   '
End Sub
'
Function TITUARCH$(Archi$)
   '
   ARCHO$ = Archi$
   If Left$(Archi$, 1) = "!" Then ARCHO$ = Mid$(Archi$, 2)
   TITUARCH$ = "Archivo " + UCase$(ARCHO$)
   If EXISTE%(LUDAT$ + ARCHO$ + ".RFS") > 0 Then
      FINDEX% = FILEINDEX%(ARCHO$)
      TITUARCH$ = TRIM$(FILETIT$(FINDEX%))
      If TITUARCH$ = "" Then
         Call LEESTRUCARCH(FINDEX%)
         TITUARCH$ = FILETIT$(FINDEX%)
      End If
   End If
   '
End Function
'
Sub FILEKILL(ARCHX$)

    ARCHXX$ = TRIM$(ARCHX$)
    If Left$(ARCHXX$, 1) = "!" Then
       Call CIERRARCH(ARCHXX$)
       ARCHXX$ = Mid$(ARCHXX$, 2)
       PPXX% = InStr(ARCHXX$, ".")
       If PPXX% < 1 Then PPXX% = 1 + Len(ARCHXX$)
       ARCHXX$ = LUCOM$ + Left$(ARCHXX$, PPXX% - 1) + ".WKF"
    End If
    '
    If EXISTE%(ARCHXX$) > 0 Then
      On Error GoTo 99
      Kill ARCHXX$
      On Error GoTo 0
    End If
    Exit Sub
    '
99  ERNU$ = TRIM$(Str$(Err.Number))
100 BRXX$ = TRIM$(REPLACAR$(ARCHXX$, "\", "/"))
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
    c$ = FORMESS$(MENSA$)
    RESPU% = MsgBox(c$, 53, "Error de Borrado de Archivo")
    If RESPU% = 4 Then
        Resume
    End If
    Close: End
    '
End Sub
'
Function COPYFILE%(ORIGI$, DESTI$)
    '
    If EXISTE%(ORIGI$) < 1 Then
      Call SAVEFILE(ORIGI$, "")
    End If
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
    c$ = FORMESS$(MENSA$)
    RESPU% = MsgBox(c$, 53, "Error de Copia de Archivos")
    If RESPU% = 4 Then
        Resume
    End If
    CPF% = 0
    '
99  COPYFILE = CPF%
    '
End Function
'
Sub CommSel(ARSELEC$, CommButton, ShowBox)
    CommButton.Enabled = False
    Call SELECHO(ARSELEC$)
    VDEVU$ = VALACT1$(ARSELEC$)
    If VDEVU$ <> "" Then
       On Error Resume Next
       ShowBox.Caption = VDEVU$
       ShowBox.TEXT = VDEVU$
       On Error GoTo 0
    End If
    CommButton.Enabled = True
End Sub
'
Sub CliCombo(COMBO As comboBox, TEXTT As String)
    For U& = 1 To COMBO.ListCount
       If TRIM$(COMBO.List(U& - 1)) = TEXTT Then
          COMBO.ListIndex = U& - 1
          Exit Sub
       End If
    Next U&
End Sub
'
Sub SELECHO(ARCHX$)
   FORSELEC.ARTITUL = TRIM$(UCase$(ARCHX$))
   FORSELEC.Show 1
   FORSELEC.ABUSQUE = ""
   Call FRESHALL
End Sub

Sub FRESHECHO(ARCHX$, Optional MODOX$)
   '
   ARCHX1$ = TRIM$(UCase$(ARCHX$))
   If Left$(ARCHX1$, 1) = "^" Then
       ARCHX1$ = Mid$(ARCHX1$, 2)
   End If
   '
   PPXX% = InStr(ARCHX1$, "/")
   If PPXX% < 2 Then PPXX% = 1 + Len(ARCHX1$)
   ARCHX1$ = TRIM$(UCase$(Left$(ARCHX1$, PPXX% - 1)))
   If ARCHX1$ = "MASTER" Then ARCHX1$ = "ECOMAST"
   '
   For KKI% = 0 To 15
     If FORSELEC.ATABLA(KKI%).Caption = ARCHX1$ Then
       FORSELEC.TABLA(KKI%).Clear
       FORSELEC.FILTER(KKI%) = ""
       On Error Resume Next
       Kill LUDAT$ + ARCHX1$ + ".SEL"
       On Error GoTo 0
     End If
   Next KKI%
   '
   If FORSELEC.ARTITUL = TRIM$(UCase$(ARCHX$)) Then
     FORSELEC.ARTITUL = ""
     FORSELEC.Text1 = ""
   End If
   '
   If UCase(MODOX$) <> "NORECAR" Then
' suprimido el 12/10/09 - ver como se comporta
'      Call FORSELEC.RECARSQL(ARCHX1$)
   End If
   '
End Sub
'
Function VALACT1$(ARCHX$)
   '
   VC1$ = ""
   ARCHX1$ = TRIM$(UCase$(ARCHX$))
   If Left$(ARCHX1$, 1) = "^" Then
       ARCHX1$ = Mid$(ARCHX1$, 2)
   End If
   '
   PPXX% = InStr(ARCHX1$, "/")
   If PPXX% < 2 Then PPXX% = 1 + Len(ARCHX1$)
   ARCHX1$ = TRIM$(UCase$(Left$(ARCHX1$, PPXX% - 1)))
   If ARCHX1$ = "MASTER" Then ARCHX1$ = "ECOMAST"
   If ARCHX1$ = "FECHA" Then ARCHX1$ = "SELFECHA"
   If ARCHX1$ = "DEUDOR1" Then
      If EXISTE%(LUDAT$ + "DEUDORL1.RFS") > 0 Then
        ARCHX1$ = "DEUDORL1"
      End If
   End If
   '
   For KKI% = 0 To 15
     If FORSELEC.ATABLA(KKI%).Caption = ARCHX1$ Then
       VC1$ = FORSELEC.VD1(KKI%).Caption
       GoTo 90
     End If
   Next KKI%
   '
90 VALACT1$ = TRIM$(VC1$)
   '
End Function

Function VALACT2$(ARCHX$)
   '
   VC2$ = ""
   ARCHX1$ = TRIM$(UCase$(ARCHX$))
   If Left$(ARCHX1$, 1) = "^" Then
       ARCHX1$ = Mid$(ARCHX1$, 2)
   End If
   '
   PPXX% = InStr(ARCHX1$, "/")
   If PPXX% < 2 Then PPXX% = 1 + Len(ARCHX1$)
   ARCHX1$ = TRIM$(UCase$(Left$(ARCHX1$, PPXX% - 1)))
   If ARCHX1$ = "MASTER" Then ARCHX1$ = "ECOMAST"
   If ARCHX1$ = "DEUDOR1" Then
      If EXISTE%(LUDAT$ + "DEUDORL1.RFS") > 0 Then
        ARCHX1$ = "DEUDORL1"
      End If
   End If
   '
   For KKI% = 0 To 15
     If FORSELEC.ATABLA(KKI%).Caption = ARCHX1$ Then
       VC2$ = FORSELEC.VD2(KKI%).Caption
       GoTo 90
     End If
   Next KKI%
   '
90 VALACT2$ = TRIM$(VC2$)
   '
End Function

Function FIDATIM#(ARCHX$)
   '
   FINDEX% = FILEINDEX%(ARCHX$)
   ARCHY$ = FIFULLNAME(FINDEX%)
   '
   FDT# = 0
   On Error Resume Next
   FDT# = FileDateTime(ARCHY$)
   On Error GoTo 0
   '
   FIDATIM# = FDT#
   '
End Function

Function REGBUS&(ARCHX$, X%, ARGU$)
   '
   RFF$ = RECFILT$(ARCHX$, X%, ARGU$)
   If Len(RFF$) >= 4 Then
     REG& = CVS(Left$(RFF$, 4))
     If REG& > 0 Then
       If REG& <= ULTREG&(ARCHX$) Then
         RELEI$ = REGLEIDO$(ARCHX$, REG&)
         REGBUS& = REG&
         Exit Function
       End If
     End If
   End If
   '
   Call SETREG(ARCHX$, 0&)
   REGBUS& = 0
   '
End Function

Function VENTABU%(FORMA$)
   FORVENTB.PROFILE.Caption = FORMA$
   FORVENTB.Show 1
   VENTABU% = VTB%
End Function

Sub EXTRAPRF(FINAME$)
   '
   AREX$ = TRIM$(UCase$(FINAME$))
   If InStr(AREX$, ".PRF") < 1 Then
     AREX$ = AREX$ + ".PRF"
   End If
   '
   If VERDEMO% = 1 Then Exit Sub
   '
29 POSI% = InStr(AREX$, "\")
   If POSI% > 0 Then
     AREX$ = Mid$(AREX$, POSI% + 1)
     GoTo 29
   End If
   '
   Dim BUFGRA As String * 128
   Dim BUFLE As String * 160
   NX% = FreeFile
   Open LUBAS$ + "FLEXPRFS.FLX" For Random Access Read Shared As #NX% Len = 160
   Fin& = LOF(NX%) / 160
   '
   ARCHIN$ = AREX$
   LUDESTI$ = CurDir
   If Right$(LUDESTI$, 1) <> "\" Then LUDESTI$ = LUDESTI$ + "\"
   '
   N1% = FreeFile
   Open LUDESTI$ + ARCHIN$ For Random Access Read Write Shared As #N1% Len = 128
   JJ& = 0
   '
   LI& = 0: LS& = Fin& + 1
   INI& = 0
   ABRE$ = TRIM$(UCase$(ARCHIN$))
10 E& = Int((LS& - LI&) / 2): If E& = 0 Then GoTo 80
   L& = LI& + E&: If L& < 1 Or L& > Fin& Then GoTo 80
   Get #NX%, L&, BUFLE$
   I1$ = TRIM$(UCase$(Left$(BUFLE$, 32)))
   If I1$ < ABRE$ Then LI& = L&: GoTo 10
   If I1$ > ABRE$ Then LS& = L&: GoTo 10
   INI& = L& - 128: If INI& < 1 Then INI& = 1
   '
80 If INI& < 1 Then
      Call MENSERR(24, "Imposible Extraer Controlador " + ABRE$)
      GoTo 90
   End If
   '
   For II& = INI& To Fin&
     Call TRACE(20, II&, Fin&)
     DoEvents
     Get #NX%, II&, BUFLE$
     I1$ = TRIM$(UCase$(Left$(BUFLE$, 32)))
     If I1$ > ARCHIN$ Then GoTo 90
     If I1$ = ARCHIN$ Then
       BUFGRA$ = Mid$(BUFLE$, 33)
       JJ& = JJ& + 1
       Put #N1%, JJ&, BUFGRA$
     End If
   Next II&
   '
90 Close #N1%
   Close #NX%
   '
End Sub

Sub FILESORT(ARENTRA$, ASALIDA$, CALIMI%, CALIMS%, Optional MODOX$)
   '
      Screen.MousePointer = 11
   Dim PRX%(64), LFX%(64), TVX%(64), JTX%(64)
   Dim ACX%(64), LDX%(64), PIX%(64)
   Dim NEGAMAX#(64)
   Dim ARGUOR() As String
   '
   ARSALID$ = UCase$(TRIM$(ASALIDA$))
   If ARSALID$ = "" Then ARSALID$ = ARENTRA$
   '
   If ARSALID$ <> ARENTRA$ Then
     If EXISTE%(LUDAT$ + ARSALID$ + ".RFS") < 1 Then
       Call COPYSTRU(ARENTRA$, ARSALID$)
     End If
   End If
   '
   URETRA& = ULTREG&(ARENTRA$)
   If URETRA& < 1 Then
     If ARSALID$ <> ARENTRA$ Then
       Call SETULTREG(ARSALID$, 0)
     End If
     GoTo 99
   End If
   '
   ReDim ARGUOR$(URETRA&)
   '
   FINDEX% = FILEINDEX%(ARENTRA$)
   FINUME% = FILENBR%(ARENTRA$)
   LOREGI% = LOREGIS%(FINDEX%)
   '
   ASORLI$ = ""
   For J% = CALIMI% To CALIMS%
     J1% = J% - CALIMI%
     PRX%(J1%) = POSFIELD%(ARENTRA$, J%)
     LFX%(J1%) = LENFIELD%(ARENTRA$, J%)
     TVX%(J1%) = TIPFIELD%(ARENTRA$, J%)
     '
     If PRX%(J1%) < 1 Then
       TTXX$ = "Error Fatal. Imposible Ordenar por Campo Numero " + TRIM$(Str$(J%)) + "\en archivo '" + ARENTRA$ + "'."
       ARMALO$ = ARENTRA$: If Left$(UCase$(ARENTRA$), 4) = "TRAN" Then ARMALO$ = Mid$(ARCHILOC$, 2)
       TTXX$ = TTXX$ + "\   \Causa Probable: Error de Configuración de Controlador '" + ARMALO$ + ".RFS"
       TTXX$ = TTXX$ + "\   \Se Recomienda Refrescar este Controlador."
       Call MENSERR(24, TTXX$)
       Call FINAL("")
       Exit Sub
     End If
     '
     If TVX%(J1%) = 5 Or TVX%(J1%) = 8 Then TVX%(J1%) = 1
     ACX%(J1%) = LFX%(J1%)
     JTX%(J1%) = 1
     LDX%(J1%) = 5
     '
     If TVX%(J1%) = 1 Or TVX%(J1%) = 2 Then
         JTX%(J1%) = 2
         LDX%(J1%) = 0
         If LFX%(J1%) < 2 Then
             ACX%(J1%) = 3
           Else
             ACX%(J1%) = 6
         End If
       ElseIf TVX%(J1%) = 3 Or TVX%(J1%) = 4 Then
         JTX%(J1%) = 2
         If LFX%(J1%) < 8 Then
             ACX%(J1%) = 7
           Else
             ACX%(J1%) = 12
         End If
         If LDX%(J1%) > 0 Then
            ACX%(J1%) = ACX%(J1%) + 5
         End If
       ElseIf TVX%(J1%) = 7 Then
         ACX%(J1%) = 16
         LDX%(J1%) = 0
     End If
     '
     PIX%(J1%) = 1 + Len(ASORLI$)
     ASORLI$ = ASORLI$ + Space$(ACX%(J1%))
     '
   Next J%
   LARGUI% = 1 + Len(ASORLI$)
   ASORLI$ = ASORLI$ + Space$(8)
   '
   HAYNEGA% = 0
   For RECNO& = 1 To URETRA&
     '
     Call TRACE(20, RECNO&, URETRA&)
     XXX$ = LEEREGIS$(FINDEX%, FINUME%, LOREGI%, RECNO&)
     '
     ARGU$ = ASORLI$
     For J1% = 0 To CALIMS% - CALIMI%
       XXZ$ = Mid$(XXX$, PRX%(J1%), LFX%(J1%))
       Mid$(ARGU$, PIX%(J1%), ACX%(J1%)) = CSTRING$(XXZ$, TVX%(J1%), ACX%(J1%), LDX%(J1%), JTX%(J1%))
       '
       If TVX%(J1%) < 5 Then
         If XVALO(CSTRING$(XXZ$, TVX%(J1%), ACX%(J1%), LDX%(J1%), JTX%(J1%))) < 0 Then
           HAYNEGA% = 1
           VALONE# = XVALO(Mid$(ARGU$, PIX%(J1%), ACX%(J1%)))
           If VALONE# < NEGAMAX#(J1%) Then NEGAMAX#(J1%) = VALONE#
         End If
       End If
       '
     Next J1%
     '
     Mid$(ARGU$, LARGUI%, 8) = CSTRING$(MKS$(RECNO&), 3, 8, 0, 2)
     ARGUOR$(RECNO&) = ARGU$
     '
   Next RECNO&
   '
   If HAYNEGA% > 0 Then
     For RECNO& = 1 To URETRA&
       '
       Call TRACE(20, RECNO&, URETRA&)
       XXX$ = LEEREGIS$(FINDEX%, FINUME%, LOREGI%, RECNO&)
       '
       ARGU$ = ASORLI$
       For J1% = 0 To CALIMS% - CALIMI%
         XXZ$ = Mid$(XXX$, PRX%(J1%), LFX%(J1%))
         TVX1% = TVX%(J1%)
         If TVX%(J1%) < 5 Then
           VALOCOM# = XVALO(CSTRING$(XXZ$, TVX%(J1%), ACX%(J1%), LDX%(J1%), JTX%(J1%))) - NEGAMAX#(J1%)
           XXZ$ = MKD$(VALOCOM#)
           TVX1% = 4
         End If
         Mid$(ARGU$, PIX%(J1%), ACX%(J1%)) = CSTRING$(XXZ$, TVX1%, ACX%(J1%), LDX%(J1%), JTX%(J1%))
       Next J1%
       '
       Mid$(ARGU$, LARGUI%, 8) = CSTRING$(MKS$(RECNO&), 3, 8, 0, 2)
       ARGUOR$(RECNO&) = ARGU$
       '
     Next RECNO&
   End If
   '
   ' ORDENAMIENTO SHELL
   INCRE& = URETRA& \ 2
   Do Until INCRE& < 1
     For UI& = INCRE& + 1 To URETRA&
       ARTEMP$ = ARGUOR$(UI&)
       For UJ& = UI& - INCRE& To 1 Step -INCRE&
         If ARTEMP$ >= ARGUOR$(UJ&) Then Exit For
         ARGUOR$(UJ& + INCRE&) = ARGUOR$(UJ&)
       Next UJ&
       ARGUOR$(UJ& + INCRE&) = ARTEMP$
     Next UI&
     INCRE& = INCRE& \ 2
   Loop
   ' FIN ORDENAMIENTO SHELL
   '
   '
   ARCHINTER$ = ARSALID$
   If ARSALID$ = ARENTRA$ Then
     ARCHINTER$ = "!" + ARENTRA$
   End If
   '
   JJ& = 0
   FINDEXINT% = FILEINDEX%(ARCHINTER$)
   FINUINTER% = FILENBR%(ARCHINTER$)
   UX1& = 1: UX2& = URETRA&: SENTI% = 1
   If Left$(UCase$(MODOX$), 4) = "DESC" Then
      UX1& = URETRA&: UX2& = 1: SENTI% = (-1)
   End If
   '
   For U& = UX1& To UX2& Step SENTI%
     Call TRACE(20, U&, URETRA&)
     RECNO& = XVALO(Mid$(ARGUOR$(U&), LARGUI%))
     If RECNO& > 0 Then
       If RECNO& <= URETRA& Then
         XXX$ = LEEREGIS$(FINDEX%, FINUME%, LOREGI%, RECNO&)
         JJ& = JJ& + 1
         POSGRABA& = (JJ& - 1) * LOREGI% + 1
         Put #FINUINTER%, POSGRABA&, XXX$
       End If
     End If
   Next U&
   '
   FINPRO& = JJ&
   If ARSALID$ = ARENTRA$ Then
     JJ& = 0
     For U& = 1 To FINPRO&
       Call TRACE(20, U&, FINPRO&)
       XXX$ = LEEREGIS$(FINDEXINT%, FINUINTER%, LOREGI%, U&)
       JJ& = JJ& + 1
       POSGRABA& = (JJ& - 1) * LOREGI% + 1
       Put #FINUME%, POSGRABA&, XXX$
     Next U&
     Call SETULTREG(ARENTRA$, JJ&)
   End If
   Call SETULTREG(ARCHINTER$, FINPRO&)
   '
   Call CIERRARCH(ARCHINTER$)
   Call CIERRARCH(ARENTRA$)
   '
99 Call CIERRARCH("*.*")
   Call BAJALDISCO
   Screen.MousePointer = 1

'   If ARENTRA$ = "BOLRESEL" Then 'TIRA ERROR SI SIGUE, NO LO PUDE SOLUCIONAR.
'      Exit Sub
'   End If
'   Screen.MousePointer = 11
'   Dim PRX%(64), LFX%(64), TVX%(64), JTX%(64)
'   Dim ACX%(64), LDX%(64), PIX%(64)
'   Dim NEGAMAX#(64)
'   Dim ARGUOR() As String
'   '
'   ARSALID$ = UCase$(TRIM$(ASALIDA$))
'   If ARSALID$ = "" Then ARSALID$ = ARENTRA$
'   '
'   If ARSALID$ <> ARENTRA$ Then
'     If EXISTE%(LUDAT$ + ARSALID$ + ".RFS") < 1 Then
'       Call COPYSTRU(ARENTRA$, ARSALID$)
'     End If
'   End If
'   '
'   URETRA& = ULTREG&(ARENTRA$)
'   If URETRA& < 1 Then
'     If ARSALID$ <> ARENTRA$ Then
'       Call SETULTREG(ARSALID$, 0)
'     End If
'     GoTo 99
'   End If
'   '
'   ReDim ARGUOR$(URETRA&)
'   '
'   FINDEX% = FILEINDEX%(ARENTRA$)
'   FINUME% = FILENBR%(ARENTRA$)
'   LOREGI% = LOREGIS%(FINDEX%)
'   '
'   ASORLI$ = ""
'   For J% = CALIMI% To CALIMS%
'     J% = J% - CALIMI%
'     If J% > 0 Then
'        PRX%(J%) = POSFIELD%(ARENTRA$, J%)  ' el error lo da aca por que j% obtiene como resultado 0
'        LFX%(J%) = LENFIELD%(ARENTRA$, J%)
'        TVX%(J%) = TIPFIELD%(ARENTRA$, J%)
'     End If
'     '
'     If PRX%(J%) < 1 Then
'       TTXX$ = "Error Fatal. Imposible Ordenar por Campo Numero " + TRIM$(Str$(J%)) + "\en archivo '" + ARENTRA$ + "'."
'       ARMALO$ = ARENTRA$: If Left$(UCase$(ARENTRA$), 4) = "TRAN" Then ARMALO$ = Mid$(ARCHILOC$, 2)
'       TTXX$ = TTXX$ + "\   \Causa Probable: Error de Configuración de Controlador '" + ARMALO$ + ".RFS"
'       TTXX$ = TTXX$ + "\   \Se Recomienda Refrescar este Controlador."
'       Call MENSERR(24, TTXX$)
'       Call FINAL("")
'       Exit Sub
'     End If
'     '
'     If TVX%(J%) = 5 Or TVX%(J%) = 8 Then TVX%(J1%) = 1
'     ACX%(J%) = LFX%(J%)
'     JTX%(J%) = 1
'     LDX%(J%) = 5
'     '
'     If TVX%(J%) = 1 Or TVX%(J%) = 2 Then
'         JTX%(J%) = 2
'         LDX%(J%) = 0
'         If LFX%(J%) < 2 Then
'             ACX%(J%) = 3
'           Else
'             ACX%(J%) = 6
'         End If
'       ElseIf TVX%(J%) = 3 Or TVX%(J%) = 4 Then
'         JTX%(J%) = 2
'         If LFX%(J%) < 8 Then
'             ACX%(J%) = 7
'           Else
'             ACX%(J%) = 12
'         End If
'         If LDX%(J%) > 0 Then
'            ACX%(J%) = ACX%(J%) + 5
'         End If
'       ElseIf TVX%(J%) = 7 Then
'         ACX%(J%) = 16
'         LDX%(J%) = 0
'     End If
'     '
'     PIX%(J%) = 1 + Len(ASORLI$)
'     ASORLI$ = ASORLI$ + Space$(ACX%(J%))
'     '
'   Next J%
'   LARGUI% = 1 + Len(ASORLI$)
'   ASORLI$ = ASORLI$ + Space$(8)
'   '
'   HAYNEGA% = 0
'   For RECNO& = 1 To URETRA&
'     '
'     Call TRACE(20, RECNO&, URETRA&)
'     XXX$ = LEEREGIS$(FINDEX%, FINUME%, LOREGI%, RECNO&)
'     '
'     ARGU$ = ASORLI$
'     For J% = 0 To CALIMS% - CALIMI%
'       XXZ$ = Mid$(XXX$, PRX%(J%), LFX%(J%))
'       Mid$(ARGU$, PIX%(J%), ACX%(J%)) = CSTRING$(XXZ$, TVX%(J%), ACX%(J%), LDX%(J%), JTX%(J%))
'       '
'       If TVX%(J%) < 5 Then
'         If XVALO(CSTRING$(XXZ$, TVX%(J%), ACX%(J%), LDX%(J%), JTX%(J%))) < 0 Then
'           HAYNEGA% = 1
'           VALONE# = XVALO(Mid$(ARGU$, PIX%(J%), ACX%(J%)))
'           If VALONE# < NEGAMAX#(J%) Then NEGAMAX#(J%) = VALONE#
'         End If
'       End If
'       '
'     Next J%
'     '
'     Mid$(ARGU$, LARGUI%, 8) = CSTRING$(MKS$(RECNO&), 3, 8, 0, 2)
'     ARGUOR$(RECNO&) = ARGU$
'     '
'   Next RECNO&
'   '
'   If HAYNEGA% > 0 Then
'     For RECNO& = 1 To URETRA&
'       '
'       Call TRACE(20, RECNO&, URETRA&)
'       XXX$ = LEEREGIS$(FINDEX%, FINUME%, LOREGI%, RECNO&)
'       '
'       ARGU$ = ASORLI$
'       For J% = 0 To CALIMS% - CALIMI%
'         XXZ$ = Mid$(XXX$, PRX%(J%), LFX%(J%))
'         TVX1% = TVX%(J%)
'         If TVX%(J%) < 5 Then
'           VALOCOM# = XVALO(CSTRING$(XXZ$, TVX%(J%), ACX%(J%), LDX%(J%), JTX%(J%))) - NEGAMAX#(J%)
'           XXZ$ = MKD$(VALOCOM#)
'           TVX1% = 4
'         End If
'         Mid$(ARGU$, PIX%(J%), ACX%(J%)) = CSTRING$(XXZ$, TVX1%, ACX%(J%), LDX%(J%), JTX%(J%))
'       Next J%
'       '
'       Mid$(ARGU$, LARGUI%, 8) = CSTRING$(MKS$(RECNO&), 3, 8, 0, 2)
'       ARGUOR$(RECNO&) = ARGU$
'       '
'     Next RECNO&
'   End If
'   '
'   ' ORDENAMIENTO SHELL
'   INCRE& = URETRA& \ 2
'   Do Until INCRE& < 1
'     For UI& = INCRE& + 1 To URETRA&
'       ARTEMP$ = ARGUOR$(UI&)
'       For UJ& = UI& - INCRE& To 1 Step -INCRE&
'         If ARTEMP$ >= ARGUOR$(UJ&) Then Exit For
'         ARGUOR$(UJ& + INCRE&) = ARGUOR$(UJ&)
'       Next UJ&
'       ARGUOR$(UJ& + INCRE&) = ARTEMP$
'     Next UI&
'     INCRE& = INCRE& \ 2
'   Loop
'   ' FIN ORDENAMIENTO SHELL
'   '
'   '
'   ARCHINTER$ = ARSALID$
'   If ARSALID$ = ARENTRA$ Then
'     ARCHINTER$ = "!" + ARENTRA$
'   End If
'   '
'   JJ& = 0
'   FINDEXINT% = FILEINDEX%(ARCHINTER$)
'   FINUINTER% = FILENBR%(ARCHINTER$)
'   UX1& = 1: UX2& = URETRA&: SENTI% = 1
'   If Left$(UCase$(MODOX$), 4) = "DESC" Then
'      UX1& = URETRA&: UX2& = 1: SENTI% = (-1)
'   End If
'   '
'   For U& = UX1& To UX2& Step SENTI%
'     Call TRACE(20, U&, URETRA&)
'     RECNO& = XVALO(Mid$(ARGUOR$(U&), LARGUI%))
'     If RECNO& > 0 Then
'       If RECNO& <= URETRA& Then
'         XXX$ = LEEREGIS$(FINDEX%, FINUME%, LOREGI%, RECNO&)
'         JJ& = JJ& + 1
'         POSGRABA& = (JJ& - 1) * LOREGI% + 1
'         Put #FINUINTER%, POSGRABA&, XXX$
'       End If
'     End If
'   Next U&
'   '
'   FINPRO& = JJ&
'   If ARSALID$ = ARENTRA$ Then
'     JJ& = 0
'     For U& = 1 To FINPRO&
'       Call TRACE(20, U&, FINPRO&)
'       XXX$ = LEEREGIS$(FINDEXINT%, FINUINTER%, LOREGI%, U&)
'       JJ& = JJ& + 1
'       POSGRABA& = (JJ& - 1) * LOREGI% + 1
'       Put #FINUME%, POSGRABA&, XXX$
'     Next U&
'     Call SETULTREG(ARENTRA$, JJ&)
'   End If
'   Call SETULTREG(ARCHINTER$, FINPRO&)
'   '
'   Call CIERRARCH(ARCHINTER$)
'   Call CIERRARCH(ARENTRA$)
'   '
'99 Call CIERRARCH("*.*")
'   Call BAJALDISCO
'   Screen.MousePointer = 1
'   '
End Sub
'
Function FILESHORTNA$(FILOX$)
   FISHNA$ = FILOX$
   While InStr(FISHNA$, "\") > 0
     FISHNA$ = Mid$(FISHNA$, 2)
   Wend
   FILESHORTNA$ = FISHNA$
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

Sub FILTERSETRECORD(Archi$, X%, ARGU$, RECORD$)
   ' definir equivalencia de rutinas SET y PUT
   Call FILTERPUTRECORD(Archi$, X%, ARGU$, RECORD$)
End Sub

Sub FILTERDELETE(Archi$, X%, ARGU$)   ' Elimina los registros que matchean argu$
   '
   RFF$ = RECFILT$(Archi$, X%, ARGU$)
   If RFF$ = "" Then Exit Sub
   '
   RINI& = CVS(Left$(RFF$, 4))
   If RINI& < 1 Then Exit Sub
   '
   Fin& = ULTREG&(Archi$)
   PFX% = POSFIELD%(Archi$, X%)
   LFX% = LENFIELD%(Archi$, X%)
   JJ& = RINI& - 1
   '
   For U& = RINI& To Fin&
     XX$ = REGLEIDO$(Archi$, U&)
     If Mid$(XX$, PFX%, LFX%) <> ARGU$ Then
       JJ& = JJ& + 1
       Call GRAREG(Archi$, XX$, JJ&)
     End If
   Next U&
   Call SETULTREG(Archi$, JJ&)
   '
End Sub

Sub FILTERFILE(Archi$, ARCHOUT$, X%, ARGU$)
   '
   AREN$ = Archi$: ARSA$ = ARCHOUT$
   If ARSA$ = "" Then ARSA$ = "!" + Archi$
   If Left$(ARSA$, 1) <> "" Then
     AREN1$ = UCase$(AREN$): If Left$(AREN1$, 1) = "!" Then AREN1$ = Mid$(AREN1$, 2)
     ARSA1$ = UCase$(ARSA$): If Left$(ARSA1$, 1) = "!" Then ARSA1$ = Mid$(ARSA1$, 2)
     If AREN1$ <> ARSA1$ Then
       Call COPYSTRU(AREN1$, ARSA1$)
     End If
   End If
   '
   RFF$ = RECFILT$(Archi$, X%, ARGU$)
   JJ& = 0
   For U& = 1 To Len(RFF$) Step 4
     REFI& = CVS(Mid$(RFF$, U&, 4))
     XX$ = REGLEIDO$(Archi$, REFI&)
     JJ& = JJ& + 1
     Call GRAREG(ARSA$, XX$, JJ&)
   Next U&
   Call SETULTREG(ARSA$, JJ&)
   Call CIERRARCH(ARSA$)
   '
End Sub

Sub FILTERUPDATE(Archi$, ARCHIN$, X%, ARGU$)
   '
   ARIN$ = ARCHIN$
   If ARIN$ = "" Then ARIN$ = "!" + Archi$
   '
   PINIX% = POSFIELD%(Archi$, X%)
   LONIX% = LENFIELD%(Archi$, X%)
   '
   URN& = ULTREG&(Archi$)
   RFF$ = RECFILT$(Archi$, X%, ARGU$)
   KK& = (-3)
   '
   JJ& = 0
   For U& = 1 To ULTREG&(ARIN$)
     XX$ = REGLEIDO$(ARIN$, U&)
     Call REPLA(XX$, ARGU$, PINIX%, LONIX%)
     KK& = KK& + 4
     If KK& <= Len(RFF$) - 3 Then
         RENOTA& = CVS(Mid$(RFF$, KK&, 4))
       Else
         URN& = URN& + 1
         RENOTA& = URN&
     End If
     Call GRAREG(Archi$, XX$, RENOTA&)
   Next U&
   '
20 REBLA$ = REGBLAN$(Archi$)
   While KK& < Len(RFF$) - 3
     KK& = KK& + 4
     RENOTA& = CVS(Mid$(RFF$, KK&, 4))
     Call GRAREG(Archi$, REBLA$, RENOTA&)
   Wend
   '
   Call SETULTREG(Archi$, URN&)
   Call CIERRARCH(Archi$)
   Call INDEXARCH(Archi$, 1)
   '
End Sub

Sub FRESHGRID()
   On Error Resume Next
   'FORVENTB.GRID.Clear
   On Error GoTo 0
End Sub

Sub FILESORTVIE(ARENTRA$, ASALIDA$, CALIMI%, CALIMS%)
   '
   ' deshabilitada desde 15 Mayo 2004 por problema YALTRES - Capacidad solo 32767 registros
   '
   Screen.MousePointer = 11
   Dim PRX%(64), LFX%(64), TVX%(64)
   '
   URETRA& = ULTREG&(ARENTRA$)
   If URETRA& < 1 Then GoTo 99
   '
   ARSALID$ = UCase$(TRIM$(ASALIDA$))
   If ARSALID$ = "" Then ARSALID$ = ARENTRA$
   '
   If ARSALID$ <> ARENTRA$ Then
     If EXISTE%(LUDAT$ + ARSALID$ + ".RFS") < 1 Then
       Call COPYSTRU(ARENTRA$, ARSALID$)
     End If
   End If
   '
   FINDEX% = FILEINDEX%(ARENTRA$)
   FINUME% = FILENBR%(ARENTRA$)
   LOREGI% = LOREGIS%(FINDEX%)
   '
   For J% = CALIMI% To CALIMS%
     PRX%(J% - CALIMI%) = POSFIELD%(ARENTRA$, J%)
     LFX%(J% - CALIMI%) = LENFIELD%(ARENTRA$, J%)
     TVX%(J% - CALIMI%) = TIPFIELD%(ARENTRA$, J%)
   Next J%
   '
   FORMALTER.SORTLIST.Clear
   For RECNO& = 1 To URETRA&
     '
     Call TRACE(20, RECNO&, URETRA&)
     XXX$ = LEEREGIS$(FINDEX%, FINUME%, LOREGI%, RECNO&)
     '
     ARGU$ = "": LARGU% = Len(ARGU$)
     For J% = CALIMI% To CALIMS%
       XXZ$ = Mid$(XXX$, PRX%(J% - CALIMI%), LFX%(J% - CALIMI%))
       TV1% = TVX%(J% - CALIMI%): If TV1% = 5 Or TV1% = 8 Then TV1% = 1
       LFX%(J%) = LFX%(J% - CALIMI%)
       AC1% = LFX%(J% - CALIMI%)
       JT1% = 1: LD1% = 5
       '
       If TV1% = 1 Or TV1% = 2 Then
             JT1% = 2: LD1% = 0
             If LFX%(J%) < 2 Then AC1% = 3 Else AC1% = 5
           ElseIf TV1% = 3 Or TV1% = 4 Then
             JT1% = 2
             If LFX%(J%) < 8 Then AC1% = 7 Else AC1% = 12
             If LD1% > 0 Then AC1% = AC1% + 5
           ElseIf TV1% = 5 Then
             AC1% = 8: LD1% = 0
           ElseIf TV1% = 7 Then
             AC1% = 16: LD1% = 0
       End If
       '
       ARGX$ = CSTRING$(XXZ$, TV1%, AC1%, LD1%, JT1%)
       ARGU$ = ARGU$ + ARGX$
     Next J%
     '
     If LARGU% = 0 Then
       LARGU% = Len(ARGU$)
     End If
     '
     ARGU$ = ARGU$ + CSTRING$(MKS$(RECNO&), 3, 8, 0, 2)
     On Error Resume Next
     FORMALTER.SORTLIST.AddItem ARGU$
     If Err Then
       Call MENSERR(24, "Imposible Ordenar " + ARENTRA$ + " en Este Equipo\Se ha Excedido la Capacidad de " + Str$(RECNO&) + " Registros.")
       GoTo 99
     End If
     '
   Next RECNO&
   '
   PPX% = LARGU% + 1
   ARCHINTER$ = ARSALID$
   If ARSALID$ = ARENTRA$ Then
     ARCHINTER$ = "!" + ARENTRA$
   End If
   '
   JJ& = 0
   FINDEXINT% = FILEINDEX%(ARCHINTER$)
   FINUINTER% = FILENBR%(ARCHINTER$)
   FINPRO& = FORMALTER.SORTLIST.ListCount - 1
   For U& = 0 To FORMALTER.SORTLIST.ListCount - 1
     Call TRACE(20, U&, FINPRO&)
     RECNO& = XVALO(Mid$(FORMALTER.SORTLIST.List(U&), PPX% + 1))
     If RECNO& > 0 Then
       If RECNO& <= URETRA& Then
         XXX$ = LEEREGIS$(FINDEX%, FINUME%, LOREGI%, RECNO&)
         JJ& = JJ& + 1
         POSGRABA& = (JJ& - 1) * LOREGI% + 1
         Put #FINUINTER%, POSGRABA&, XXX$
       End If
     End If
   Next U&
   '
   FORMALTER.SORTLIST.Clear
   FINPRO& = JJ&
   If ARSALID$ = ARENTRA$ Then
     JJ& = 0
     For U& = 1 To FINPRO&
       Call TRACE(20, U&, FINPRO&)
       XXX$ = LEEREGIS$(FINDEXINT%, FINUINTER%, LOREGI%, U&)
       JJ& = JJ& + 1
       POSGRABA& = (JJ& - 1) * LOREGI% + 1
       Put #FINUME%, POSGRABA&, XXX$
     Next U&
     Call SETULTREG(ARENTRA$, JJ&)
   End If
   Call SETULTREG(ARCHINTER$, FINPRO&)
   '
   Call CIERRARCH(ARCHINTER$)
   Call CIERRARCH(ARENTRA$)
   '
99 Call CIERRARCH("*.*")
   Call BAJALDISCO
   FORMALTER.SORTLIST.Clear
   Screen.MousePointer = 1
   '
End Sub
'
Sub GENERASTRINGALEATORIO()
   ' se mantiene aqui como prueba
   '
   'FIN& = XVALO(Text1)
   'For U& = 1 To FIN&
   '  TEXTO$(U&) = ""
   '  For J% = 1 To 16
   '    NX% = Int(26 * Rnd(1))
   '    TEXTO$(U&) = TEXTO$(U&) + Chr$(NX% + 65)
   '  Next J%
   'Next U&
End Sub

Sub SORTSHELL()
   ' se mantiene aqui como prueba
   '
   'TINI = Timer
   'NUE& = XVALO(Text1)
   'INCRE& = NUE& \ 2
   'Do Until INCRE& < 1
   '  For UI& = INCRE& + 1 To NUE&
   '    TEMP$ = TEXTO$(UI&)
   '    For UJ& = UI& - INCRE& To 1 Step -INCRE&
   '      If TEMP$ >= TEXTO$(UJ&) Then Exit For
   '      TEXTO$(UJ& + INCRE&) = TEXTO$(UJ&)
   '    Next UJ&
   '    TEXTO$(UJ& + INCRE&) = TEMP$
   '  Next UI&
   'INCRE& = INCRE \ 2
   'Loop
   'Label1 = Timer - TINI
   'Label1.Refresh
   '
   'List1.Clear
   'For U& = 1 To NUE&
   '  List1.AddItem TEXTO$(U&)
   'Next U&
   'List1.Refresh
End Sub

Sub CANCELUCOM()
    '
    If Left$(JOBID$, 2) = "WF" Then
      APPDIR$ = App.Path
      LICARTRA$ = LOADFILE$(APPDIR$ + "\LICARTRA.FLX")
      For KU& = 1 To Len(LICARTRA$) / 128
        XX$ = Mid$(LICARTRA$, 128 * KU& - 127, 128)
        If Left$(XX$, 8) = JOBID$ Then
          Mid$(XX$, 9, 16) = Space$(16)
          Mid$(LICARTRA$, 128 * KU& - 127 + 8, 16) = Space$(16)
          Call SAVEFILE(APPDIR$ + "\LICARTRA.FLX", LICARTRA$)
          Exit For
        End If
      Next KU&
    End If
    '
End Sub

Sub VESARRAS(CIXI%)
   '
   If TIBASE = "DAT" Then
      Call VESARRDAT(CIXI%)
      Exit Sub
   End If
   '
   If TIBASE$ <> "MDB" Then Exit Sub
   '
   If CIXI% > 0 Then
     '
     SCMA = Screen.MousePointer
     Screen.MousePointer = 11
     '
     Call CARDEPOS
     For KI% = 1 To 64
       CATEO#(KI%) = 0
     Next KI%
     '
     Call APERBASDAT("STOCKS")
     SQLX$ = " SELECT * FROM MOVISTO "
     SQLX$ = SQLX$ + "WHERE Cod_Inte = " & CIXI% & " "
     '
     Set MoviTemp = Stocks.OpenRecordset(SQLX$, dbOpenSnapshot)
     CAMOVI& = 0
     With MoviTemp
       On Error Resume Next
       .MoveFirst
       If Err < 1 Then
         Do While Not .EOF
           CAMOVI& = CAMOVI& + 1
           CANTI# = XVALO(!CANTINGRE) - XVALO(!CANTSALID)
           DEP% = XVALO(!DEPOMOVI)
           '
           COLUX% = 1
           For KI% = 1 To NUSU%
             If DEP% = COSU%(KI%) Then
               COLUX% = KI%
               Exit For
             End If
           Next KI%
           '
           CATEO#(COLUX%) = CATEO#(COLUX%) + CANTI#
         .MoveNext
         Loop
       End If
     End With
     '
     VESAL$ = String$(256, 0)
     For KI% = 1 To NUSU%
       Mid$(VESAL$, 4 * KI% - 3, 4) = MKS$(CATEO#(KI%))
     Next KI%
     '
     With SalStock
       .FindFirst "CodiInte = " & CIXI%
       If .NoMatch = False Then
           .Edit
         Else
           .AddNew   ' YA CONVERTIDO
           On Error Resume Next
           !CODIEMPR = CodiEmp%
           On Error GoTo 0
           !CodiInte = CIXI%
       End If
       !VeSalSto = VESAL$
       .Update
     End With
     SalStock.Close
     '
   End If
   '
99 End Sub
'
Sub VESARRDAT(CIXI%)
   '
   Static HOII%, YAINDEX%, UREMAS&
   '
   SMP = Screen.MousePointer
   '
   If UREMAS& < 1 Then
      UREMAS& = ULTREG&("MASTER")
   End If
   '
   REGSTO& = CIXI% + 1
   If CIXI% < 1 Or REGSTO& > UREMAS& Then
      GoTo 99
   End If
   '
   If HOII% < 1 Then
      HOII% = HOY(HOS$)
   End If
   '
   User% = USNBR% Mod 100
   If User% < 1 Then User% = 1
   '
   Call CARDEPOS
   For U& = 1 To NUSU%: CAFIN#(U&) = 0: Next U&
   '
   If YAINDEX% = 0 Then
      RINI& = ULTREG&("MOVISTO") - 513
      If RINI& < 1 Then RINI& = 1
      MVS$ = REGLEIDO$("MOVISTO", RINI&)
      Call GRAREG("MOVISTO", MVS$, RINI&)
      Call CIERRARCH("*.*")
      YAINDEX% = 1
   End If
   '
   REG& = 0: CANACU = 0
   URE& = ULTREG&("MOVISTO")
   RGM$ = RECFILT$("MOVISTO", 2, MKI$(CIXI%))
   '
   For k& = 1 To Len(RGM$) Step 4
      REG& = CVS(Mid$(RGM$, k&, 4))
      If REG& <= URE& Then
         MVS$ = REGLEIDO$("MOVISTO", REG&)
         CANTI# = (Int(10 * (CVD(Mid$(MVS$, 88, 8)) - CVD(Mid$(MVS$, 96, 8))) + 0.5)) / 10
         DEP% = Asc(Mid$(MVS$, 83, 1))
         LOT$ = Mid$(MVS$, 5, 16)
         '
         COLUX% = 1
         For KI% = 1 To NUSU%
            If DEP% = COSU%(KI%) Then
               COLUX% = KI%
               GoTo 15
            End If
         Next KI%
         '
15       CAFIN#(COLUX%) = CAFIN#(COLUX%) + CANTI#
         CANACU = CANACU + CANTI#
      End If
   Next k&
   '
   If REG& > 0 Then
      MVS$ = REGLEIDO$("MOVISTO", REG&)
      If Abs(CVS(Mid$(MVS$, 118, 4)) - CANACU) >= 0.05 Then
         Call REPLA(MVS$, MKS$(CANACU), 118, 4)
         Call GRAREG("MOVISTO", MVS$, REG&)
      End If
   End If
   '
   SST$ = String$(40, 0)
   SDP$ = String$(256, 0)
   For KI% = 1 To NUSU%
      CADEP = CAFIN#(KI%)
      If Abs(CADEP) > 0.005 Then
         If KI% <= 10 Then
            Call REPLA(SST$, MKS$(CADEP), 4 * KI% - 3, 4)
         End If
         If KI% <= 62 Then
            Call REPLA(SDP$, MKS$(CADEP), 4 * KI% - 3, 4)
         End If
      End If
   Next KI%
   Call REPLA(SDP$, MKS$(REG&), 249, 4)
   Call REPLA(SDP$, MKI$(HOII%), 253, 2)
   Call REPLA(SDP$, MKI$(Int(HORANUM(Time$))), 255, 2)
   Fin& = NUMAS% + 1
   '
   CTSX% = 0
   For UKL& = ULTREG&("SALSTOCK") + 1 To Fin&
     Call GRAREG("SALSTOCK", String$(39, 0) + Chr$(1), UKL&)
     CTSX% = 1
   Next UKL&
   Call GRAREG("SALSTOCK", SST$, REGSTO&)
   For UKL& = ULTREG&("SALDEPOS") + 1 To Fin&
     Call GRAREG("SALDEPOS", String$(255, 0) + Chr$(1), UKL&)
     CTSX% = 1
   Next UKL&
   Call GRAREG("SALDEPOS", SDP$, REGSTO&)
   '
   If CTSX% > 0 Then        ' liberar bloqueos
     Call CIERRARCH("SALSTOCK")
     Call CIERRARCH("SALDEPOS")
   End If
   '
99 End Sub
'
Sub ARMALISTARCH(UL$, EXTE$, Optional FILSEC$)
   ULO$ = UL$: If Right$(ULO$, 1) <> "\" Then ULO$ = UL$ + "\"
   EXTO$ = EXTE$: PPXX% = InStr(EXTO$, "."): If PPXX% > 0 Then EXTO$ = TRIM$(Mid$(EXTO$, PPXX% + 1)) ' verificado EPB y corregido error de < por >
   ASALI$ = "DI" + EXTO$
   If ASALI$ <> "DIRFS" Then Call COPYSTRU("DIRFS", ASALI$)
   '
   On Error Resume Next
   FORMALTER.File1.Path = ULO$
   If Err > 0 Then GoTo 10
   FORMALTER.File1.Pattern = "*." + EXTO$
   If Err > 0 Then GoTo 10
   On Error GoTo 0
   FORMALTER.File1.Refresh
   '
   Fin& = FORMALTER.File1.ListCount
   REG& = 1
   For AX% = 0 To FORMALTER.File1.ListCount - 1
     U& = AX% + 1
     Call TRACE(20, U&, Fin&)
     PriChar$ = Left$(FORMALTER.File1.List(AX%), 1)
     If PriChar$ <> "!" Then
        TAM% = InStr(FORMALTER.File1.List(AX%), ".") - 1
        If TAM% < 0 Then TAM% = Len(FORMALTER.File1.List(AX%))
        XX$ = LOADFILE(ULO$ + "\" + TRIM$(FORMALTER.File1.List(AX%)))
        '
        If FILSEC$ <> "" Then
          If Mid$(XX$, 127, 2) <> FILSEC$ Then
            GoTo 19
          End If
        End If
        '
        TITU$ = TRIM$(Left$(XX$, 47))
        FICH$ = Mid$(FORMALTER.File1.List(AX%), 1, TAM%)
        '
        Call REPLA(YY$, UCase$(FICH$), 1, 9)
        Call REPLA(YY$, UCase$(TITU$), 10, 47)
        Call GRAREG(ASALI$, YY$, REG&)
        REG& = REG& + 1
     End If
19 Next AX%
   '
   Call SETULTREG(ASALI$, REG& - 1)
   Call FRESHECHO(ASALI$)
   '
   Exit Sub
   '
10 COMUNI ("Imposible Armar Lista de Archivos en U.L. " & UL$ & " y Extensión " & EXTE$)
   Call FINAL("")
End Sub


Function SELECCION_Archivo$(Extension$, NAMECOMMONDIALOG As CommonDialog)
'ABRE EL NAMECOMMONDIALOG
'SE LE DEBE PASAR LA EXENTSION DEL ARCHIVO SIN EL PUNTO
'SI NO SE LE PASA NADA TRAE *.*
   NAMECOMMONDIALOG.CancelError = True
   NAMECOMMONDIALOG.DialogTitle = "Selección de Archivo"
   ' cancelado, si no da error - NAMECOMMONDIALOG.Flags = &H2& + &H4& + &H8 + &H800 + &H2000 + &H8000 'cdlOFNCreatePrompt
   If TRIM$(Extension$) <> "" Then
    tiparch1$ = "*." & Extension$
    tiparch2$ = Extension$ & "|"
    NAMECOMMONDIALOG.FILTER = "Archivo de Texto(" & tiparch1$ & ")|*." & tiparch2$
   Else
    NAMECOMMONDIALOG.FILTER = "Todos los archivos (*.*)|*.*|Mapas de Bit" & _
         "(*.BMP)|*.BMP|Archivos Graficos JPEG (*.JPG)|*.JPG"
   End If
   NAMECOMMONDIALOG.FilterIndex = 1
   On Error GoTo 99
   NAMECOMMONDIALOG.ShowOpen
   On Error GoTo 0
   '
   DoEvents
   Screen.MousePointer = 11
   SELECCION_Archivo$ = UCase$(NAMECOMMONDIALOG.FileName)
99
End Function

Function CLAVEUSERDAT$(NUS%, RET_IDENTIFICADOR$)
'   LEO EL USER.DAT Y DEVUELVE LA CLAVE SEGUN USUARIO
    FUSER$ = LOADFILE$("USER.DAT")
    JJI& = 0: KKI% = 0
    CLAVEUSERDAT$ = "": RET_IDENTIFICADOR$ = ""
    For i% = 1 To Len(FUSER$) Step 64
       JJ% = JJ% + 1
       F0$ = Mid$(FUSER$, i%, 64)
       CLAV$ = Mid$(F0$, 43, 12): If Left$(CLAV$, 1) = "@" Then CLAV$ = DESENCRI$(Mid$(F0$, 44, 11)) + " "
       
       If NUS% = JJ% Then
          CLAVEUSERDAT$ = TRIM$(CLAV$)
          Ident$ = Mid$(F0$, 31, 12)
          If Left$(Ident$, 1) = "@" Then Ident$ = DESENCRI$(Mid$(Ident$, 2, 11)) + " "
          RET_IDENTIFICADOR$ = TRIM$(Ident$)
          Exit For
       End If
    Next i%

End Function
