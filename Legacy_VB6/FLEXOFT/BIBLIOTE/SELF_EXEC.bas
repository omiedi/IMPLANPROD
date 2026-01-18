Attribute VB_Name = "SELF_EXEC"
' RAM-FUNCTIONS ' ************ revisa estado de memoria RAM

' ****************************************************************************

'    sub  GRARAM
'                 graba directamente en memoria RAM
'    fun  LEERAM$
'                 lee directamente de memoria RAM
'    fun  USNBR%
'                 determina numero de usuario conectado
'    fun  IDENTER$
'                 determina identificador de terminal
'    fun LUBOOT$
'                 determina unidad de boot
' ****************************************************************************

Sub ACCRAM(TX$, PX%, LX%, MODO%)
    '
    Static RAMSTRI As String * 512
    Static CTX%
    '
    If CTX% = 0 Then
        'NX% = FILEOPEN%("C:\FLEXOFT\FLEXRAM.DRV", 0, 512)
        NX% = FILEOPEN%(LUCOM$ + "FLEXRAM.DRV", 0, 512)
        Get #NX%, 1, RAMSTRI$
        Close #NX%: NX% = 0
        '
        ATRA$ = CurDir
        On Error GoTo 100
        NX% = FreeFile
        Open "\ATRABA" For Input Shared As #NX%
        Line Input #NX%, AA$
        Close #NX%: NX% = 0
        ATRA$ = TRIM$(AA$)
        GoTo 110
        '
100     Resume 110
110     On Error GoTo 0
        CTX% = 1
    End If
    '
    If MODO% = 1 Then ' accede p/escritura de TX$ en pos. PX%
        Mid$(RAMSTRI$, PX%, Len(TX$)) = TX$
      ElseIf MODO% = 2 Then
        TX$ = Mid$(RAMSTRI$, PX%, LX%)
      ElseIf MODO% = 3 Then
        LDTXXX$ = Mid$(RAMSTRI$, 99, 46)
112     If Mid$(LDTXXX$, 2, 1) = ":" Then
           LDTXXX$ = Mid$(LDTXXX$, 3)
           GoTo 112
        End If
        If Mid$(LDTXXX$, 2, 1) = ":" Then LDTXXX$ = Mid$(LDTXXX$, 3)
        LDTYYY$ = ULODATOS$ + LDTXXX$
        Call REPLA(RAMSTRI$, LDTYYY$, 99, 46)
        '
        If UCase$(App.EXEName) = "SEGUIMIENTO COMERCIAL" Then
           Call REPLA(RAMSTRI$, LUDAT$ + Space$(46), 99, 46)
           Call SAVEFILE(LUCOM$ + "ULOGDAT.DRV", Left$(LUDAT$, 2))
        End If
        NX% = FILEOPEN%(LUCOM$ + "FLEXRAM.DRV", 0, 512)
        Put #NX%, 1, RAMSTRI$
        Close #NX%: NX% = 0
        '
        If VERCD% = 1 Then
          On Error Resume Next
          NN11% = FreeFile
          'Open "C:\FLEXOFT\ULOGDAT.DRV" For Output As #NN11%
          Open LUCOM$ + "ULOGDAT.DRV" For Output As #NN11%
          Print #NN11%, "C:"
          Close #NN11%
          On Error GoTo 0
        End If
      Else
        MsgBox "Error de Invocacion de Rutina 'ACCRAM' - Consulte."
    End If
    '
End Sub

Sub GRARAM(PX%, TX$) '************* Rutina de Grabacion Directa en Memoria RAM
    '
    LEYE$ = TX$
    POSI% = PX%
    CACA% = Len(TX$)
    Call ACCRAM(LEYE$, POSI%, CACA%, 1)
    '
End Sub
'

Function LEERAM$(PX%, LX%) '********** Rutina de Lectura Directa de Memoria RAM
    '
    POSI% = PX%
    CACA% = LX%
    Call ACCRAM(LEYE$, POSI%, CACA%, 2)
    LEERAM$ = LEYE$
    '
End Function
'

Sub BAJRAM()
    '
    Call ACCRAM(LEYE$, POSI%, CACA%, 3)
    '
End Sub

Sub AUTOEXEC()
    '
    LANGUA% = 0
    VERDEMO% = (-1)
    INTERPRE% = 0
    '
    IDENTER$ = "000000"
    NOMTER$ = "Estación Local"
    '
    ULOAC$ = Left$(CurDir, 2)
    LUBAS$ = CurDir: If Right$(CurDir$, 1) <> "\" Then LUBAS$ = LUBAS$ + "\"
    LUCOM$ = CurDir: If Right$(CurDir$, 1) <> "\" Then LUCOM$ = LUCOM$ + "\"
    LUDAT$ = CurDir: If Right$(CurDir$, 1) <> "\" Then LUDAT$ = LUDAT$ + "\"
    LUBOOT$ = Left$(CurDir, 1)
    '
    USNBR% = 101
    USERNAME$ = "Usuario General"
    MENUPRIN$ = ""
    MODINVO$ = ""
    DENOMOD$ = ""
    MODMENU$ = ""
    '
    EMPREAC$ = "Su Empresa - " + VERNUM$ + " - " + USERNAME$
    '
    On Error Resume Next
    MkDir "C:\FLEXOFT"
    On Error GoTo 0
    '
    If EXISTE%("C:\FLEXOFT\IDENTER.DRV") Then
      X1$ = Left$(LOADFILE$("C:\FLEXOFT\IDENTER.DRV"), 40)
      J1$ = DESENCRI$(Left$(X1$, 38))
      J2$ = TRIM$(Left$(J1$, 32))
      If J2$ <> "" Then NOMTER$ = J2$
      J2$ = TRIM$(Mid$(J1$, 33, 6))
      If J2$ <> "" Then IDENTER$ = J2$
    End If
    '
    Call VERILUCOM
    '
    Call CARDATRAM
    Call BAJRAM
    '
End Sub
'
Sub VERILUCOM()
    '
    LUCOM$ = "C:\FLEXOFT\"
    NX% = FreeFile
    On Error Resume Next
    Open LUCOM$ + "RCONTROL.DRV" For Output As #NX%
    If Err Then
       On Error GoTo 0
       Close #NX%
       GoTo 20
    End If
    Exit Sub
    '
20  Randomize Timer
24  JOBI$ = ""
25  PPX% = Int(Rnd * 90)
    If PPX% < 48 Then GoTo 25
    If PPX% > 57 And PPX% < 65 Then GoTo 25
    JOBI$ = JOBI$ + Chr$(PPX%)
    If Len(JOBI$) < 8 Then GoTo 25
    JOBID$ = JOBI$
    If Left$(JOBID$, 2) = "WF" Then GoTo 24
    '
    '
50  ' LLEGA AQUI SI NO ES UN DEMO Y NO PUDO ACCEDER A C:\FLEXOFT
    '
    APPDIR$ = App.Path
    LICARTRA$ = LOADFILE$(APPDIR$ + "\LICARTRA.FLX")
    For KU& = 1 To Len(LICARTRA$) / 128
      XX$ = Mid$(LICARTRA$, 128 * KU& - 127, 128)
      If TRIM$(Mid$(XX$, 9, 16)) = "" Then
        If Left$(XX$, 2) = "WF" Then
          JOBID$ = Left$(XX$, 8)
          LUCOM$ = APPDIR$ + "\" + JOBID$ + "\"
          If EXISTE%(LUCOM$ + ".") > 0 Then
            AHORAX$ = Format$(Now, "DD/MM/YY   HH:MM")
            Mid$(XX$, 9, 16) = AHORAX$
            Mid$(LICARTRA$, 128 * KU& - 127, 128) = XX$
            Call SAVEFILE(APPDIR$ + "\LICARTRA.FLX", LICARTRA$)
            Exit Sub
          End If
        End If
      End If
    Next KU&
    '
    Randomize Timer
    JOBI$ = ""
55  PPX% = Int(Rnd * 90)
    If PPX% < 48 Then GoTo 55
    If PPX% > 57 And PPX% < 65 Then GoTo 55
    JOBI$ = JOBI$ + Chr$(PPX%)
    If Len(JOBI$) < 6 Then GoTo 55
    JOBID$ = "WF" + JOBI$
    LUCOM$ = APPDIR$ + "\" + JOBID$ + "\"
    On Error Resume Next
    MkDir APPDIR$ + "\" + JOBID$
    If Err Then
      MsgBox "Imposible Crear Carpeta de Trabajo " + APPDIR$ + "\" + JOBID$
      Close: End: Exit Sub
    End If
    '
    While (Len(LICARTRA$) Mod 128) <> 0
      LICARTRA$ = LICARTRA$ + " "
    Wend
    XX$ = Space$(128)
    Mid$(XX$, 1, 8) = JOBID$
    AHORAX$ = Format$(Now, "DD/MM/YY   HH:MM")
    Mid$(XX$, 9, 16) = AHORAX$
    LICARTRA$ = LICARTRA$ + XX$
    Call SAVEFILE(APPDIR$ + "\LICARTRA.FLX", LICARTRA$)
    '

End Sub

Sub BLOQEXE(PROEXE$)
   '
   Static NK%, RECNOI&
   Dim AA As String * 128
   Dim BB As String * 128
   '
   SCMA = Screen.MousePointer
   If LUCOM$ = CurDir + "\" Then GoTo 99
   If UCase$(App.EXEName) = "ADMIFLEX" Then GoTo 99
   If UCase$(App.EXEName) = "CTRLOCK" Then GoTo 99
   LU$ = LUDAT$
   '
   HOII% = HOY(HO$)
   PRXX$ = TRIM$(UCase$(PROEXE$))
   If Left$(PRXX$, 1) = "*" Then       ' libera el proceso
      If RECNOI& > 0 Then
         Call BLOQARCH("MONIEXEC")
         N1% = FILEOPEN(LU$ + "MONIEXEC.DAT", 0, 128)
         Get #N1%, RECNOI&, BB$
         TEBLO$ = BB$
         Call REPLA(TEBLO$, HO$, 89, 8)
         Call REPLA(TEBLO$, Time$, 99, 8)
         Call REPLA(TEBLO$, Chr$(10), 128, 1)
         LSet BB$ = TEBLO$
         Put #N1%, RECNOI&, BB$
         Close #N1%: N1% = 0
         RECNOI& = 0
         Call LIBARCH("MONIEXEC")
      End If
      If NK% > 0 Then
         Close #NK%: NK% = 0
      End If
      Exit Sub
   End If
   '
   TEBLO$ = Space$(128)
   USX$ = USERNAME$
   ITX$ = IDENTER$
   TRX$ = NOMTER$
   '
   Call REPLA(TEBLO$, PRXX$ + ".EXE", 1, 16)
   Call REPLA(TEBLO$, USX$, 17, 24)
   Call REPLA(TEBLO$, TRX$, 45, 24)
   Call REPLA(TEBLO$, HO$, 69, 8)
   Call REPLA(TEBLO$, Time$, 79, 8)
   Call REPLA(TEBLO$, ITX$, 119, 8)
   '
   TINI = Timer: CAINTEN% = 0
   SCMA = Screen.MousePointer
11 Screen.MousePointer = 11
12 On Error GoTo 15
   If NK% < 1 Then
      NK% = FreeFile
      Open LUCOM$ + PRXX$ + ".EXK" For Random Access Read Write Lock Write As #NK% Len = 128
   End If
   LSet AA$ = TEBLO$
   Put #NK%, 1, AA$
   On Error GoTo 0
   '
   If LU$ <> "" Then
      Call BLOQARCH("MONIEXEC")
      N1% = FILEOPEN%(LU$ + "MONIEXEC.DAT", 0, 128)
      If RECNOI& < 1 Then
           RECNOI& = 1 + LOF(N1%) / 128
         Else
           Get #N1%, RECNOI&, BB$
           TEBLO$ = BB$
           Call REPLA(TEBLO$, USX$, 17, 24)
           Call REPLA(TEBLO$, TRX$, 45, 24)
           Call REPLA(TEBLO$, ITX$, 119, 8)
      End If
      Call REPLA(TEBLO$, Chr$(10), 128, 1)
      LSet BB$ = TEBLO$
      Put #N1%, RECNOI&, BB$
      Close #N1%: N1% = 0
      Call LIBARCH("MONIEXEC")
   End If
   '
99 Screen.MousePointer = SCMA
   Exit Sub
   '
   '
'  **************************************************
   '
15 Resume 16
16 On Error GoTo 0
   Close #NK%: NK% = 0
   If Timer < TINI + 8 Then GoTo 12
   '
   If UCase$(App.EXEName) = "WINTABU" Then
      If CAINTEN% < 3 Then
         CAINTEN% = CAINTEN% + 1
         TINI = Timer
         GoTo 11
      End If
   End If
   '
   If NK% < 1 Then NK% = FreeFile
   Open LUCOM$ + PRXX$ + ".EXK" For Random Access Read As #NK% Len = 128
   Get #NK%, 1, AA$
   B2$ = Mid$(AA$, 1, 16)
   B4$ = Mid$(AA$, 69, 8)
   B5$ = Mid$(AA$, 79, 8)
   B6$ = Mid$(AA$, 17, 24)
   Close #NK%: NK% = 0
   '
   LOMAX% = 0
   USY$ = "El Proceso '" + TRIM$(B2$) + "'"
   TRY$ = "Está Ejecutándose en Esta Terminal"
   FFY$ = "Desde el " + TRIM$(B4$) + " a las " + TRIM$(B5$) + " horas"
   If TRIM$(B6$) <> "" Then FFY$ = FFY$ + "\por el Usuario " + TRIM$(B6$)
   '
   If UCase$(TRIM$(PROEXE$)) = "WINTABU" Then
        COMU$ = "ATENCION: Dificultad de Acceso\a Utilitario de Consultas y Listados.\Cierre o Finalice Consulta Previa Abierta\en este Puesto de Trabajo."
      Else
        COMU$ = "ATENCION: Imposible Continuar !!!\  \" + USY$ + "\" + TRY$ + "\" + FFY$ + "\  \Cierre el Proceso Abierto\y Vuelva a Intentar el Ingreso."
   End If
   'Call COMUNI(COMU$)
   If COMALTER%(COMU$ + "\\Re-Intentar\Cancelar") = 1 Then
        GoTo 12
      Else
        Close: End
   End If
   'Call CIERRARCH("*.*")
   Call FINAL("")
   '
End Sub
'

Sub GRAJOBID()
    '
    Randomize Timer
    JOBI$ = ""
10  PPX% = Int(Rnd * 90)
    If PPX% < 48 Then GoTo 10
    If PPX% > 57 And PPX% < 65 Then GoTo 10
    JOBI$ = JOBI$ + Chr$(PPX%)
    If Len(JOBI$) < 8 Then GoTo 10
    JOBID$ = JOBI$
    JOBIX$ = ENCRIPTA$(JOBI$)
    Call GRARAM(223, JOBIX$)
    '
End Sub
'
Sub GRATICO()
    '
    HHH = HOY(HMX$)
    TMX = Timer
    TMX1$ = Left$(Str$(TMX) + Space$(8), 8)
    DVR = 0: For U% = 1 To 8
                U1% = 1 + (5 * U% Mod 8)
                U2% = 9 - U%
                DVR = DVR + 25.42 * Asc(Mid$(HMX$, U1%, 1)) - 17.83 * Asc(Mid$(TMX1$, U2%, 1))
             Next U%
             '
  '  Call GRARAM(1, MKS$(TMX))
  '  Call GRARAM(5, MKS$(DVR))
    '
End Sub


Sub REDIVERI() ' revisa timer y digito verificador **************
    '
    'TME = Timer: TMX = CVS(LEERAM$(1, 4))
    '             DVR = CVS(LEERAM$(5, 4))
    HMX$ = Mid$(Date$, 4, 2) + "/" + Left$(Date$, 2) + "/" + Right$(Date$, 2)
    TMX1$ = Left$(Str$(TMX) + Space$(8), 8)
    '
    DVX = 0: For U% = 1 To 8
                U1% = 1 + (5 * U% Mod 8): U2% = 9 - U%
                DVX = DVX + 25.42 * Asc(Mid$(HMX$, U1%, 1)) - 17.83 * Asc(Mid$(TMXS$, U2%, 1))
             Next U%
             '
    If DVX <> DVR Or TME < TMX Or TME > TMX + 20 Then
         MsgBox "Error de Conexion al Módulo - Dígito Verificador no Válido - Consulte."
         Close: End
    End If
    '
End Sub
'

Sub CARDATRAM() '******** carga datos basicos de memoria RAM
    '
    Static CMD$
    '
' numero y nombre de usuario ***************************
    '
    Call GRARAM(9, MKI$(USNBR%))
    Call GRARAM(231, AJUSTI$(USERNAME$, 30))
    '
' driver del menu principal activo *********************
    '
    Call GRARAM(13, AJUSTI$(MENUPRIN$, 8))
    '
' carpeta sistema basico LUBAS$
    '
    Call GRARAM(97, LUBOOT$)
    Call GRARAM(99, AJUSTI$(LUDAT$, 46))
    Call GRARAM(145, AJUSTI$(LUBAS$, 46))
    Call GRARAM(193, AJUSTI$(EMPREAC$, 30))
    '
' identificador y nombre de terminal
    '
    Call GRARAM(91, AJUSTI$(IDENTER$, 6))
    Call GRARAM(261, AJUSTI$(NOMTER$, 32))
    '
' tipo de borde de ventanas
    '
    If EXISTE%("SALVAPAN.DRV") > 0 Then
        Dim NXX As String * 16
        NX% = FILEOPEN%("SALVAPAN.DRV", 9, 16)
        Get #NX%, 2, NXX$
        TV% = Asc(Left$(NXX$, 1))
        Close #NX%: NK% = 0
        MCX% = 1               ' fuerza monitor color
    End If
    Call GRARAM(98, Chr$(10 * TV% + MCX%))
    '
    Call BLOQEXE(App.EXEName)
    '
End Sub
'

Sub SELEJER(EJUI%)
      '
5     LU$ = LUDAT$
      Dim COMEJ1%(2), FINEJ1%(2)
      EJU% = 1
      If VERCD% = 1 Then GoTo 90
      '
      Dim V As String * 32
      '
20    N1% = FILEOPEN%(LU$ + "VARIAB.DAT", 0, 32)
      Get #N1%, 2, V$
      EJACT1$ = TRIM$(Left$(V$, 28))
      COMEJ1%(1) = CVI(Mid$(V$, 29, 2))
      FINEJ1%(1) = CVI(Mid$(V$, 31, 2))
      '
      COMEPRO = COMEJ1%(1) + 32 * 13
      FINECAL% = DIANTE(COMEPRO)
      If FINEJ1%(1) <> FINECAL% Then
           V1$ = V$
           Mid$(V1$, 31, 2) = MKI$(FINECAL%)
           LSet V$ = V1$
           Put #N1%, 2, V$
           Close #N1%: N1% = 0
           GoTo 20
      End If
      '
      If COMEJ1%(1) = FECHANUM("01/08/97") Then
           V1$ = V$
           Mid$(V1$, 29, 2) = MKI$(FECHANUM("01/08/98"))
           Mid$(V1$, 31, 2) = MKI$(FECHANUM("31/07/99"))
           LSet V$ = V1$
           Put #N1%, 2, V$
           Close #N1%: N1% = 0
           GoTo 20
      End If
      '
      Close #N1%: N1% = 0
      '
30    N1% = FILEOPEN%(LU$ + "VARIAB.OLD", 0, 32)
      Get #N1%, 2, V$
      EJANT$ = TRIM$(Left$(V$, 28))
      COMEJ1%(2) = CVI(Mid$(V$, 29, 2))
      FINEJ1%(2) = CVI(Mid$(V$, 31, 2))
      '
      COMEPRO = COMEJ1%(2) + 32 * 13
      FINECAL% = DIANTE(COMEPRO)
      If FINEJ1%(2) <> FINECAL% Then
           V1$ = V$
           Mid$(V1$, 31, 2) = MKI$(FINECAL%)
           LSet V$ = V1$
           Put #N1%, 2, V$
           Close #N1%: N1% = 0
           GoTo 30
      End If
      Close #N1%: N1% = 0
      '
      If TRIM$(EJANT$) = "" Or TRIM$(EJACT1$) = "" Then
          EJU% = 1
          GoTo 90
      End If
      '
      MENSA$ = "SISTEMA DE CONTABILIDAD GENERAL"
      MENSA$ = MENSA$ + "\FLEXOFT - SCG - Version 10.1"
      MENSA$ = MENSA$ + "\Apertura Simultanea de 2 Ejercicios Contables"
      MENSA$ = MENSA$ + "\ \Seleccione Ejercicio Contable de Trabajo\\" + EJACT1$ + "\" + EJANT$ + "\Cancelar"
      '
      EJU% = COMALTER%(MENSA$)
      If EJU% < 1 Or EJU% > 2 Then EJU% = 0
      '
      Call GRARAM(87, MKI$(COMEJ1%(EJU%)))
      Call GRARAM(89, MKI$(FINEJ1%(EJU%)))
      '
90    If EJU% = 2 Then EJU% = 255
      Call GRARAM(69, Chr$(EJU%))
      If EJU% = 255 Then EJU% = (-1)
      EJUI% = EJU%
      '
End Sub
'

Sub GRACOMEJ()
      '
105   LU$ = LUDAT$
      '
      Dim V As String * 32
      '
120   N1% = FILEOPEN%(LU$ + "VARIAB.DAT", 0, 32)
      Get #N1%, 2, V$
      CMJ% = CVI(Mid$(V$, 29, 2))
      FMJ% = CVI(Mid$(V$, 31, 2))
      '
      COMEPRO = CMJ% + 32 * 13
      FINECAL% = DIANTE(COMEPRO)
      If FMJ% <> FINECAL% Then
           V1$ = V$
           Mid$(V1$, 31, 2) = MKI$(FINECAL%)
           LSet V$ = V1$
           Put #N1%, 2, V$
           Close #N1%: N1% = 0
           GoTo 120
      End If
      '
      If CMJ% = FECHANUM("01/08/97") Then
           V1$ = V$
           Mid$(V1$, 29, 2) = MKI$(FECHANUM("01/08/98"))
           Mid$(V1$, 31, 2) = MKI$(FECHANUM("31/07/99"))
           LSet V$ = V1$
           Put #N1%, 2, V$
           Close #N1%: N1% = 0
           GoTo 120
      End If
      '
      Close #N1%: N1% = 0
      Call GRARAM(87, MKI$(CMJ%))
      Call GRARAM(89, MKI$(FMJ%))
End Sub
      '

Function SUBEMP$()
      '
      SE$ = LEERAM$(70, 1)
      If SE$ < "0" Or SE$ > "9" Then SE$ = "0"
      SUBEMP$ = SE$
      '
End Function
'


Function EJACT$()
      '
      EXTE$ = "DAT": If EJER% < 0 Then EXTE$ = "OLD"
      LU$ = LUDAT$
      '
      Dim V As String * 32
      N1% = FILEOPEN%(LU$ + "VARIAB." + EXTE$, 9, 32)
      Get #N1%, 2, V$
      EJACT$ = TRIM$(Left$(V$, 28))
      Close #N1%: N1% = 0
      '
End Function
'
Sub TRANSLABELS(FORNAM$)
   '
   If LANGUA% = 0 Then Exit Sub
   '
   For IFORMU% = 1 To Forms.Count
     If Forms(IFORMU% - 1).Name = FORNAM$ Then
       '
       CAPTI$ = ""
       On Error Resume Next
       CAPTI$ = TRIM$(Forms(IFORMU% - 1).Caption)
       Forms(IFORMU% - 1).Caption = TRANSLATE$(CAPTI$)
       On Error GoTo 0
       '
       For KKI% = 1 To Forms(IFORMU% - 1).Controls.Count
         '
10       CAPTI$ = ""
         On Error GoTo 91
         CAPTI$ = TRIM$(Forms(IFORMU% - 1).Controls(KKI% - 1).Caption)
         Forms(IFORMU% - 1).Controls(KKI% - 1).Caption = TRANSLATE$(CAPTI$)
         For KKJ% = 1 To 99
           CAPTI$ = TRIM$(Forms(IFORMU% - 1).Controls(KKI% - 1).Caption(KKJ%))
           Forms(IFORMU% - 1).Controls(KKI% - 1).Caption(KKJ%) = TRANSLATE$(CAPTI$)
         Next KKJ%
         On Error GoTo 0
         '
20       TOTIP$ = ""
         On Error GoTo 92
         TOTIP$ = TRIM$(Forms(IFORMU% - 1).Controls(KKI% - 1).ToolTipText)
         Forms(IFORMU% - 1).Controls(KKI% - 1).ToolTipText = TRANSLATE$(TOTIP$)
         For KKJ% = 1 To 99
           TOTIP$ = TRIM$(Forms(IFORMU% - 1).Controls(KKI% - 1).ToolTipText(KKJ%))
           Forms(IFORMU% - 1).Controls(KKI% - 1).ToolTipText(KKJ%) = TRANSLATE$(TOTIP$)
         Next KKJ%
         On Error GoTo 0
         '
30       CAPTI$ = ""
         On Error GoTo 93
         For KKJ% = 0 To 99
           CAPTI$ = TRIM$(Forms(IFORMU% - 1).Controls(KKI% - 1).TabCaption(KKJ%))
           Forms(IFORMU% - 1).Controls(KKI% - 1).TabCaption(KKJ%) = TRANSLATE$(CAPTI$)
         Next KKJ%
         On Error GoTo 0
         '
80       On Error Resume Next
         Forms(IFORMU% - 1).Controls(KKI% - 1).Refresh
         On Error GoTo 0
         '
       Next KKI%
     End If
   Next IFORMU%
   Exit Sub
   '
91 Resume 20
92 Resume 30
93 Resume 80
   '
End Sub

Function TRANSLATE$(TESPA$)
   '
   Static FITRA%
   Static BUFTRA As String * 1024
   Static URETRA&
   Static XINDI$
   '
   TRANSLATE$ = TESPA$
   If LANGUA% < 1 Or LANGUA% > 9 Then Exit Function
   If TRIM$(TESPA$) = "" Then Exit Function
   If LUBAS$ = "" Then Exit Function
   '
   If FITRA% < 1 Then
     Screen.MousePointer = 11
     FITRA% = FreeFile
     Open LUBAS$ + "TRALABS.DRV" For Random Access Read Write Shared As #FITRA% Len = 1024
     URETRA& = LOF(FITRA%) / 1024
     XINDI$ = String$(2 * URETRA&, 0)
     For U& = 1 To URETRA&
       Get #FITRA%, U&, BUFTRA$
       Mid$(XINDI$, 2 * U& - 1, 2) = Left$(BUFTRA$, 2)
     Next U&
     Screen.MousePointer = 1
   End If
   '
   BVT% = BINVAL%(TESPA$)
50 BVT% = BINVAL%(TESPA$)
   AZ$ = MKI$(BVT%)
   RFF$ = "": PINI& = 1
60 POSI& = InStr(PINI&, XINDI$, AZ$)
   If POSI& < 1 Then GoTo 70                           ' NEXT RGX&
   If (POSI& Mod 2) <> 1 Then
      PINI& = POSI& + 1
      GoTo 60
   End If
   PINI& = POSI& + 2
   RGL& = (POSI& + 1) / 2
   RFF$ = RFF$ + MKS$(RGL&)
   GoTo 60
   '
70 For KKI% = 1 To Len(RFF$) Step 4
     REIN& = CVS(Mid$(RFF$, KKI%, 4))
     If REIN& > 0 Then
       If REIN& <= URETRA& Then
         Get #FITRA%, REIN&, BUFTRA$
         TINDI$ = TRIM$(Mid$(BUFTRA$, 3, 126))
         If TINDI$ = TESPA$ Then
           TRALA$ = TRIM$(Mid$(BUFTRA$, 128 * LANGUA% + 1, 128))
           If TRALA$ <> "" Then TRANSLATE$ = TRALA$
           Exit Function
         End If
       End If
     End If
   Next KKI%
   '
End Function



