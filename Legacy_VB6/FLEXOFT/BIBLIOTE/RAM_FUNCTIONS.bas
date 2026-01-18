Attribute VB_Name = "RAM_FUNCTIONS"

Dim VENCIAB%
Dim RAMSTRI$
Public CURDINI$

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

Sub ACCRAM(TX$, PX%, Lx%, MODO%)
    '
    Static CTX%
    '
    If RAMSTRI$ = "" Then
       RAMSTRI$ = Left$(LOADFILE(LUCOM$ + "FLEXRAM.DRV") + String$(512, 0), 512)
    End If
    '
    If CTX% = 0 Then
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
        TX$ = Mid$(RAMSTRI$, PX%, Lx%)
      ElseIf MODO% = 3 Then
        'LDTXXX$ = Mid$(RAMSTRI$, 99, 46)
112     'If Mid$(LDTXXX$, 2, 1) = ":" Then
        '  LDTXXX$ = Mid$(LDTXXX$, 3)
        '  GoTo 112
        'End If
'MsgBox LDTXXX$
        'If Mid$(LDTXXX$, 2, 1) = ":" Then LDTXXX$ = Mid$(LDTXXX$, 3)
        'LDTYYY$ = ULODATOS$ + LDTXXX$
        'Call REPLA(RAMSTRI$, LDTYYY$, 99, 46)
'MsgBox LDTYYY$
        '
        Call SAVEFILE(LUCOM$ + "FLEXRAM.DRV", RAMSTRI$)
        '
        'If VERCD% = 1 Then
        '  On Error Resume Next
        '  NN11% = FreeFile
        '  Open "C:\FLEXOFT\ULOGDAT.DRV" For Output As #NN11%
        '  Print #NN11%, "C:"
        '  Close #NN11%
        '  On Error GoTo 0
        'End If
'MsgBox "ESPERE"
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

Function LEERAM$(PX%, Lx%) '********** Rutina de Lectura Directa de Memoria RAM
   '
   If RAMSTRI$ = "" Then
      RAMSTRI$ = Left$(LOADFILE(LUCOM$ + "FLEXRAM.DRV") + String$(512, 0), 512)
   End If
   '
   LEERAM$ = ""
   If PX% > 0 Then
     LEERAM$ = Mid$(RAMSTRI$, PX%, Lx%)
   End If
   '
End Function
'

Sub BAJRAM()
    '
    Call ACCRAM(LEYE$, POSI%, CACA%, 3)
    '
End Sub

Sub MERRFINAL(TMESSA$)
   On Error GoTo 0
   MsgBox TMESSA$, 0, "Imposible Continuar Ejecución"
   Call FINAL("")
   Close
   End
End Sub
'
Sub VERJOBID(OKEY%)
    '
    LUCOM$ = "C:\FLEXOFT\"
    '
    CMDX$ = Command$
    PPX% = InStr(CMDX$, "#")
    If Mid$(CMDX$, PPX% + 9, 1) = "#" Then
      JOBID$ = Mid$(CMDX$, PPX% + 1, 8)
    End If
    '
    
    If Left$(JOBID$, 2) = "WF" Then
      LUCOM$ = App.Path + "\" + JOBID$ + "\"
    End If
    '
    If PPX% > 0 Then
      ATRIXX$ = TRIM$(Left$(CMDX$, PPX% - 1))
      If Len(ATRIXX$) > 0 Then
        PPYY% = InStr(ATRIXX$, "/")
        If PPYY% < 1 Then PPYY% = InStr(ATRIXX$, ";")
        If PPYY% > 0 Then
           PARAPLI$ = Mid$(ATRIXX$, PPYY% + 1)
           ATRIXX$ = Left$(ATRIXX$, PPYY% - 1)
        End If
        APLINVO$ = UCase$(ATRIXX$)
      End If
    End If
    '
    OKEY% = 0
    If TIBASE$ = "" Then TIBASE$ = "DAT"    ' para inicializarla
   
    '
' Set FEHOR = New FEHOR   ' ESTO ES PARA ENLAZAR LA BIBLIOTECA FECHORAS
    
    EJER% = Asc(LEERAM$(69, 1))
    COMEJ% = CVI(LEERAM$(87, 2))
    FINEJ% = CVI(LEERAM$(89, 2))
    LANGUA% = XVALO(LEERAM$(305, 1))
    CodiEmp% = Asc(LEERAM$(191, 1))
    CodiSuc% = Asc(LEERAM$(192, 1))
    
    '
'   verifica interprete o ejecutable
    If INTERPRE% = 1 Then GoTo 10
    '
    
    ' LUCOM$ = "C:\FLEXOFT\"
    If Command$ = "" Then GoTo 99
    '
    
    ' verifica copia legitima
    If LEGITIM% <> 1 Then Exit Sub
    '
    
    USNBR% = CVI(LEERAM$(9, 2))
    
      USNX% = USNBR% Mod 100
      USERID$ = ""
      If USNX% > 0 Then
         TXTTX$ = LOADFILE$("USER.DAT")
         TXTTY$ = Mid$(TXTTX$, 64 * USNX% - 63, 64)
         Ident$ = Mid$(TXTTY$, 31, 12)
         If Left$(Ident$, 1) = "@" Then Ident$ = DESENCRI$(Mid$(Ident$, 2, 11)) + " "
         USERID$ = TRIM$(Ident$)
      End If
      
    MENUPRIN$ = TRIM$(LEERAM$(13, 8))
    MODINVO$ = TRIM$(LEERAM$(21, 8))
    DENOMOD$ = TRIM$(LEERAM$(29, 32))
    MODMENU$ = TRIM$(LEERAM$(61, 8))
    EJER% = Asc(LEERAM$(69, 1))
    '
    
    COMEJ% = CVI(LEERAM$(87, 2))
    FINEJ% = CVI(LEERAM$(89, 2))
    IDENTER$ = TRIM$(LEERAM$(91, 6))
    LUBOOT$ = TRIM$(LEERAM$(97, 1))
    '
   
    LUDAT$ = TRIM$(LEERAM$(99, 46))
    LUBAS$ = TRIM$(LEERAM$(145, 46))
    LUACC$ = LUDAT$ + "BdAccess\"
      On Error Resume Next
      MkDir LUACC$
      On Error GoTo 0
    CURDINI$ = CurDir
      '
      
5   EMPREAC$ = TRIM$(LEERAM$(193, 30))
    EMPREAC$ = EMPREAC$ + " (" + TRIM$(str$(CodiEmp%)) + ")"
    '
    
    CMDZ$ = LEERAM$(223, 8)
    USERNAME$ = TRIM$(LEERAM$(231, 30))
    NOMTER$ = TRIM$(LEERAM$(261, 30))
    '
    
    CONVERDE$ = ENCRIPTA$(LEERAM$(293, 6))
    CLAVERDE$ = LEERAM$(299, 6)
    
    RSW% = XVALO(LEERAM$(306, 1))
    VENCIAB% = XVALO(LEERAM$(307, 1))
    '
    
    VERDEMO% = 1
    If CLAVERDE$ = CONVERDE$ Then VERDEMO% = (-1)
    '
    
    If JOBID$ = "" Then
        Call GRARAM(223, String$(8, 0))
        Call BAJRAM
        '
        CMDX$ = Command$
        PPX% = InStr(CMDX$, "#" + DESENCRI$(CMDZ$) + "#")
        If PPX% < 1 Then GoTo 99    ' conecta con logon
        JOBID$ = Mid$(CMDX$, PPX% + 1, 8)
        ATRIXX$ = TRIM$(Left$(CMDX$, PPX% - 1))
        If Len(ATRIXX$) > 0 Then
          APLINVO$ = UCase$(ATRIXX$)
        End If
    End If
    '
    
10  If Left$(UCase$(App.EXEName), 6) <> "TRAZAB" Then
      If Left$(UCase$(App.EXEName), 6) <> "FSETUP" Then
        Call MUESTRA_FORMINI
      End If
    End If
    '
    CONNSQL$ = TRIM$(LOADFILE$(LUDAT$ + "CONNSQLS.DRV"))
    '
    Screen.MousePointer = 11
    '
    ULOAC$ = Left$(CurDir, 2)
    If EMPREAC$ <> "" Then EMPREAC$ = EMPREAC$ + " - "
    EMPREAC$ = EMPREAC$ + VERNUM$ + " - " + USERNAME$
    '
    If EJER% <> 1 Then EJER% = (-1)
    '
    EJU$ = EJACT$
    SBE$ = SUBEMP$
    '
    
    USERTER$ = TRIM(Left$(USERNAME$, 8)) + "/" + NOMTER$
    '
    NUEWINTA% = 0: If RSW% = 1 Then NUEWINTA% = 1
    '
   
    If VENCIAB% > 0 Then
      If TRIM$(VERANTER$) <> "" Then
        Call CONECRUN(VERANTER$, "")
        Call FINAL("")
      End If
    End If
    '
   
    If UCase$(App.Title) <> UCase$(App.EXEName) Then
'       MsgBox UCase$(App.Title)
'       MsgBox "Imposible Ejecutar '" + App.EXEName + "' por Presunta Alteración."
    End If
    '
   
    If App.EXEName <> "PRUEBA" Then Call BLOQEXE(App.EXEName)
    '
    Screen.MousePointer = 1
    OKEY% = 1
    '
Exit Sub
    '
    '
99  ' conecta con FLOGON.EXE
    On Error Resume Next
    RETOR$ = TRIM$(UCase$(App.EXEName))
    Shell "FLOGON.EXE " + RETOR$, vbNormalFocus
    If Err Then
       Call MERRFINAL("Falta Utilitario 'FLOGON.EXE'.")
    End If
    Exit Sub
    '
    '
    JK$ = ""
    Randomize Timer
100 PPX% = Int(Rnd * 63)
    If PPX% < 48 Then GoTo 100
    If PPX% > 57 Then PPX% = PPX% + 7
    JK$ = JK$ + Chr$(PPX%)
    If Len(JK$) = 4 Then JK$ = JK$ + ":"
    If Len(JK$) < 9 Then GoTo 100
    MsgBox "IQL Interrupt at JOB_ID " + JK$, 4096 + 16
    Call FINAL("")
    '
End Sub
'
Function INTERPRE%()
    '
    Static INTERANT%
    '
    If INTERANT% <> 0 Then
      INTERPRE% = INTERANT%
      If INTERANT% < 0 Then INTERPRE% = 0
      Exit Function
    End If
    
    INTERPRE% = 0: INTERANT% = (-1)
    AA1$ = CurDir
    If LCase(Mid$(TRIM$(CurDir), 3, 21)) = "\archivos de programa" Or CurDir = "C:\Program Files\Microsoft Visual Studio\VB98" Or CurDir = "C:\Program Files (x86)\Microsoft Visual Studio\VB98" Or CurDir = "C:\Program Files\DevStudio\VB\FLEXOFT\PROY_SQL" Or CurDir = "C:\Program Files\Microsoft Visual Studio\VB98\FLEXOFT\PROY_ACC" Or CurDir = "C:\Program Files (x86)\Microsoft Visual Studio\VB98" Or CurDir = "C:\Program Files (x86)\Microsoft Visual Studio\VB98\FLEXOFT_0\PROY_SQL" Or CurDir = "C:\Windows\system32" Or CurDir = "C:\Program Files (x86)\Microsoft Visual Studio\VB98" Or CurDir = " & AA1$ Then  ' \DevStudio\VB" Then
      SVLC$ = SERIVOL$("C:")
      If SVLC$ = "851E-A480" Or SVLC$ = "445E-73A2" Or SVLC$ = "3E4C-0FFF" Or SVLC$ = "0E44-11CF" Or SVLC$ = "081D-0BD9" Or SVLC$ = "41D8-F7FA" Or SVLC$ = "881C-3B03" Or SVLC$ = "3023-1C01" Or SVLC$ = "0E62-07D3" Or SVLC$ = "60D9-89EE" Or SVLC$ = "BCF8-2846" Or SVLC$ = "7C4E-86D4" Or SVLC$ = "F02F-93D4" Or SVLC$ = "F42D-F2FE" Or SVLC$ = "DC0D-8CC7" Or SVLC$ = "E2A6-15A5" Or SVLC$ = "6036-F9FE" Or SVLC$ = "5693-4F90" Or SVLC$ = "A4A3-7C8A" Then   ' NOTEBOOK        LUCOM$ = "C:\FLEXOFT\"
        LUCOM$ = "C:\FLEXOFT\"
        LUDAT$ = TRIM$(LEERAM$(99, 46))
        LUACC$ = LUDAT$ + "BdAccess\"
             On Error Resume Next
            MkDir LUACC$
            On Error GoTo 0
        LUBAS$ = "F:\FLEXOFT\NOVEDAD\"  '  "\FLEXOWIN\"
        If EXISTE%(LUBAS$ + "FLEXPRFS.FLX") < 1 Then
          LUBAS$ = Left$(LUDAT$, 2) + "\FLEXOFT\WINDOWS\"
        End If
        USNBR% = 101
        USERNAME$ = TRIM$(LEERAM$(231, 30))
        EMPREAC$ = TRIM$(LEERAM$(193, 30))
        CMDZ$ = LEERAM$(223, 8)
        VERDEMO% = (-1)
        X1$ = Left$(LOADFILE$("C:\FLEXOFT\IDENTER.DRV"), 40)
        J1$ = DESENCRI$(Left$(X1$, 38))
        IDENTER$ = TRIM$(Mid$(J1$, 33, 6))
        NOMTER$ = TRIM$(LEERAM$(261, 30))
        INTERPRE% = 1: INTERANT% = 1
        MsgBox "Directorio de Datos: " + LUDAT$
      End If
    End If
    '
End Function
'

Function LEGITIM%()
   '
   Static CTVOL%
   '
   If CTVOL% < 1 Then
     IKXX$ = "@?"
     On Error Resume Next
     IKXX$ = FORSELEC.Label99.Caption
     On Error GoTo 0
     '
     If IKXX$ = "@?" Then
        ' NO HUBO ACCESO A LABEL99 DE FORSELEC
        CTVOL% = 1
        GoTo 90
     End If
     If Left$(IKXX$, 40) = "Microsoft Vísual Basic 5.O ............." Then
        ' NO ESTA BLOQUEADO
        CTVOL% = 1
        GoTo 90
     End If
     '
' provisoriamente queda todo con LEGITIM%
CTVOL% = 1
GoTo 90
     '
     CUDI$ = Left$(CurDir$, 2)
     If Mid$(CUDI$, 2, 1) <> ":" Then
       MsgBox "Imposible Ejecutar " + App.EXEName + " en la Carpeta " + CurDir$
       Close: End
     End If
     SEVOL$ = SERIVOL$(CUDI$)
     If Len(SEVOL$) < 9 Then
       MsgBox "Imposible Ejecutar " + App.EXEName + " en Volumen " + SEVOL$
       Close: End
     End If
     '
     While Len(SEVOL$) < 13
       KKII% = KKII% + 1
       SEVOL$ = SEVOL$ + Chr$(18 + Asc(Mid$(SEVOL$, KKII%, 1)))
     Wend
     SEVOL$ = ENCRIPTA$(Left$(SEVOL$, 13))
     '
     FINA$ = App.EXEName
     PPXX% = InStr(FINA$, ".")
     If PPXX% < 1 Then PPXX% = 1 + Len(FINA$)
     EXENA1$ = UCase$(Left$(FINA$, PPXX% - 1))
     EXENA$ = ENCRIPTA$(UCase$(Left$(FINA$, PPXX% - 1)))
     If Len(EXENA$) < 6 Then
       CTVOL% = 1
       GoTo 90
     End If
     '
     While Len(EXENA$) < 13
       CARAVA% = (Asc(Mid$(EXENA$, 2, 1)) + Asc(Mid$(EXENA$, 4, 1))) Mod 256
       EXENA$ = Chr$(CARAVA%) + EXENA$
       CARAVA% = (Asc(Mid$(EXENA$, Len(EXENA$) - 3, 1)) + Asc(Mid$(EXENA$, Len(EXENA$) - 1, 1))) Mod 256
       EXENA$ = EXENA$ + Chr$(CARAVA%)
     Wend
     EXENA$ = Left$(EXENA$, 13)
     '
     CSTRI$ = Chr$(64)
     For KKI% = 2 To 13
       CARAVA% = (KKI% - 1) * Asc(Mid$(EXENA$, KKI%, 1)) + (KKI% + 3) * Asc(Mid$(SEVOL$, KKI%, 1))
       While CARAVA% < 48: CARAVA% = CARAVA% + 43: Wend
       While CARAVA% > 90: CARAVA% = CARAVA% - 43: Wend
       CSTRI$ = CSTRI$ + Chr$(CARAVA%)
     Next KKI%
     '
     If Mid$(IKXX$, 28, 13) = CSTRI$ Then
       CTVOL% = 1
       GoTo 90
     End If
     '
     Screen.MousePointer = 1
     Call MENSERR(24, "Módulo '" + App.EXEName + "' no Habilitado\en la Presente Instalación.")
     Close: End
     '
   End If
   '
90 LEGITIM% = CTVOL%
   '
End Function
'
Function IDACTU$()
   '
   Static IDACT0$
   If IDACT0$ = "" Then
     IDX$ = IDENTER$
     '
     NTI% = 0
     If VERDEMO% <> 1 Then
       Dim IDS As String * 96
       n1% = FreeFile
       On Error Resume Next
       Open "IDENSYS.DRV" For Random Access Read Write Shared As #n1% Len = 96
       If Err Then GoTo 20
       On Error GoTo 0
       'N1% = FILEOPEN%("IDENSYS.DRV", 0, 96)
       For i% = 11 To LOF(n1%) / 96
         Get #n1%, i%, IDS$
         ABC$ = DESENCRI$(Mid$(IDS$, 38, 38))
         If Mid$(ABC$, 33, 6) = IDX$ Then
           NTI% = i% - 10
           GoTo 20
         End If
       Next i%
       '
20     Close #n1%
     End If
     '
     On Error GoTo 0
     IDACT0$ = String$(6, 0)
     Call REPLA(IDACT0$, Chr$(USNBR% Mod 100), 5, 1)
     Call REPLA(IDACT0$, Chr$(NTI% Mod 256), 6, 1)
   End If
   '
   HOII% = HOY(HOS$)
   TIII% = Int(HORANUM(Time))
   IDACTU$ = IDACT0$
   Call REPLA(IDACTU$, MKI$(HOII%), 1, 2)
   Call REPLA(IDACTU$, MKI$(TIII%), 3, 2)
   '
End Function
'
Sub RECUMOVI()
    '
    Screen.MousePointer = 11
    Dim BUGRA As String * 140
    NKX% = FreeFile
    Open LUCOM$ + "MOVISTO.LOC" For Random Access Read Write As #NKX% Len = 140
    '
    RECUPE% = 0
    For KKL& = 1 To LOF(NKX%) / 140
      Get #NKX%, KKL&, BUGRA$
      If BUGRA$ <> String$(140, 0) Then
        ARED$ = TRIM$(Left$(BUGRA$, 8))
        '
        If ARED$ <> "" Then
          If Left$(UCase$(ARED$), 4) = "MOVI" Then
           RECONO& = CVS(Mid$(BUGRA$, 9, 4))
           RERECU$ = Mid$(BUGRA$, 13)
           If RECONO& > 0 Then
              If RECONO& <= ULTREG&(ARED$) Then
                RERED$ = REGLEIDO$(ARED$, RECONO&)
                If Left$(RERED$, 32) <> Left$(RERECU$, 32) Then
                    RECUPE% = RECUPE% + 1
                    Call BLOQARCH(ARED$)
                    Call REGAPP(ARED$, RERECU$)
                    RECONO& = ULTREG&(ARED$)
                    Call CIERRARCH(ARED$)
                    TTXX$ = BUGRA$
                    Call REPLA(TTXX$, MKS$(RECONO&), 9, 4)
                    BUGRA$ = TTXX$
                    Put #NKX%, KKL&, BUGRA$
                    BUGRB$ = BUGRA$ + String$(116, 0)
                    Call REGAPP("MOVIRECU", BUGRB$)
                  Else
                    If RECONO& < ULTREG&(ARED$) - 256 Then
                      BUGRA$ = String$(140, 0)
                      Put #NKX%, KKL&, BUGRA$
                    End If
                End If
              End If
            End If
          End If
        End If
      End If
19  Next KKL&
    '
    NKY% = FreeFile
    Open LUCOM$ + "MOVISTO.LNW" For Random Access Read Write As #NKY% Len = 140
    JJ& = 0
    For KKL& = 1 To LOF(NKX%) / 140
       Get #NKX%, KKL&, BUGRA$
       If BUGRA$ <> String$(140, 0) Then
         ARED$ = TRIM$(Left$(BUGRA$, 8))
         If ARED$ <> "" Then
           If Left$(UCase$(ARED$), 4) = "MOVI" Then
             JJ& = JJ& + 1
             Put #NKY%, JJ&, BUGRA$
           End If
         End If
       End If
    Next KKL&
    '
    Close #NKX%
    Close #NKY%
    On Error Resume Next
    Kill LUCOM$ + "MOVISTO.LOC"
    Name LUCOM$ + "MOVISTO.LNW" As LUCOM$ + "MOVISTO.LOC"
    On Error GoTo 0
    '
    Call CIERRARCH("MOVIRECU")
    Call CIERRARCH("*.*")
    Call BAJALDISCO
    Screen.MousePointer = 1
    '
End Sub
'
Sub BLOQEXE(PROEXE$)
   '
   Static NK%, RECNOI&
   Dim AA As String * 128
   Dim BB As String * 128
   '
   If INTERPRE% > 0 Then GoTo 99
   If UCase$(App.EXEName) = "ADMIFLEX" Then GoTo 99
   If UCase$(App.EXEName) = "CTRLOCK" Then GoTo 99
   '
   ' LO QUE SIGUE ES PARA QUE EN TERMINAL SERVER AL TENER PROBLEMAS DE ACCEDER A
   ' C:\FLEXOFT AL CONECTARSE CON FLOGON SALGA ORDENADAMENTE - EPB - 7 MAYO 2005
   NN11% = FreeFile
   On Error Resume Next
   Open LUCOM$ + "IDENTER.DRV" For Random Access Read Write As #NN11% Len = 40
   If Err Then
     On Error GoTo 0
     GoTo 99
   End If
   On Error GoTo 0
   ' HASTA AQUI EPB - 7 MAYO 2005
   '
   'If InStr(UCase$(App.EXEName), "WINTAB") > 0 Then GoTo 99
   ULOX$ = ULODATOS$
   LU$ = LUDAT$
   '
   HOII% = HOY(HO$)
   PRXX$ = TRIM$(UCase$(PROEXE$))
   If Left$(PRXX$, 1) = "*" Then       ' libera el proceso
      If RECNOI& > 0 Then
         'Call BLOQARCH("MONIEXEC") deshabilitado por colgadas sabo
         n1% = FILEOPEN(LU$ + "MONIEXEC.DAT", 0, 128)
         Get #n1%, RECNOI&, BB$
         TEBLO$ = BB$
         Call REPLA(TEBLO$, HO$, 89, 8)
         Call REPLA(TEBLO$, Time$, 99, 8)
         Call REPLA(TEBLO$, Chr$(10), 128, 1)
         LSet BB$ = TEBLO$
         Put #n1%, RECNOI&, BB$
         Close #n1%: n1% = 0
         RECNOI& = 0
         'Call LIBARCH("MONIEXEC") deshabilitado por colgadas sabo
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
      'Call BLOQARCH("MONIEXEC")
      n1% = FILEOPEN%(LU$ + "MONIEXEC.DAT", 0, 128)
      If RECNOI& < 1 Then
           RECNOI& = 1 + LOF(n1%) / 128
         Else
           Get #n1%, RECNOI&, BB$
           TEBLO$ = BB$
           Call REPLA(TEBLO$, USX$, 17, 24)
           Call REPLA(TEBLO$, TRX$, 45, 24)
           Call REPLA(TEBLO$, ITX$, 119, 8)
      End If
      Call REPLA(TEBLO$, Chr$(10), 128, 1)
      LSet BB$ = TEBLO$
      Put #n1%, RECNOI&, BB$
      Close #n1%: n1% = 0
      'Call LIBARCH("MONIEXEC")
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
    If JOBID$ = "" Then
      Randomize Timer
      JOBI$ = ""
10    PPX% = Int(Rnd * 90)
      If PPX% < 48 Then GoTo 10
      If PPX% > 57 And PPX% < 65 Then GoTo 10
      JOBI$ = JOBI$ + Chr$(PPX%)
      If Len(JOBI$) < 8 Then GoTo 10
      JOBID$ = JOBI$
    End If
    JOBIX$ = ENCRIPTA$(JOBID$)
    Call GRARAM(223, JOBIX$)
    '
End Sub
'
Sub GRATICO()
    '
    HHH = HOY(HMX$)
    TMX = Timer
    TMX1$ = Left$(str$(TMX) + Space$(8), 8)
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
    TMX1$ = Left$(str$(TMX) + Space$(8), 8)
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
    Static cmd$
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
       PPXX% = InStr(EMPREAC$, "(")
       If PPXX% < 1 Then PPXX% = 1 + Len(TRIM$(EMPREAC$))
       EMPRAC1$ = Left$(EMPREAC$, PPXX% - 1)
    Call GRARAM(193, AJUSTI$(EMPRAC1$, 30))
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

Sub CARIMPRE(CCI$())
'
Static CTZ%
      '
      Dim CCP$(10)
      LUX$ = ""
      If EXISTE%(LUCOM$ + "SVD0200.EXE") = 1 Then
           LUX$ = LUCOM$
           If CTZ% = 0 Then
               CTZ% = 1
           End If
      End If
      '
      GoSub 32000
      '
      If Left$(CCI$(1), 4) = Chr$(27) + Chr$(64) + Chr$(27) + Chr$(120) Then
          CCI$(1) = Chr$(27) + Chr$(64) + Chr$(27) + Chr$(120) + Chr$(0)
      End If
      '
      If Left$(CCI$(15), 2) = Chr$(27) + Chr$(87) Then
          CCI$(15) = Chr$(27) + Chr$(87) + Chr$(0)
      End If
      '
      GoTo 40000


32000 n% = FreeFile
      GoSub 32026
      If CTX% = 1 Then GoTo 32020
      Dim PAR As String * 8
32002 If n% < 1 Then n% = FreeFile
      Open LUX$ + "SVD0200.EXE" For Random Access Read Write Shared As #n% Len = 8
32006 If LOF(n%) <= 128 Then GoSub 32038: GoTo 32020
32008 For II% = 1 To 32
32010 Get #n%, II%, PAR$
      PAR1$ = PAR$
32012 '
32014 JJ% = 9: If InStr(PAR1$, Chr$(255)) > 0 Then JJ% = InStr(PAR1$, Chr$(255))
32016 CCI$(II%) = Left$(PAR1$, JJ% - 1)
32018 Next II%
32020 Close #n%: n% = 0
32022 Return
32024 '
32026 CTX% = 0
      Dim X128 As String * 128
32028 n% = FreeFile
      Open LUX$ + "SVD0200.EXE" For Random Access Read Write Shared As #n% Len = 128
      If LOF(n%) < 16 Then GoTo 32036
32030 Close #n%: n% = 0
32032 GoTo 32048
32034 '
32036 Close #n%: n% = 0
32038 FILEKILL "SVD0200.EXE"
32040 CCI$(1) = "27-64-27-120-0-18"
      CCI$(2) = "12"
      CCI$(3) = "27-50"
      CCI$(4) = "27-48"
      CCI$(5) = "27-51"
      CCI$(6) = "27-65"
      CCI$(7) = "27-67"
      CCI$(8) = "27-118"
      CCI$(9) = "8"
      CCI$(10) = "27-80"
      CCI$(11) = "27-77"
      CCI$(12) = "27-15"
      CCI$(13) = "18"
      CCI$(14) = "27-87-1"
      CCI$(15) = "27-87-0"
      CCI$(16) = "27-71"
      CCI$(17) = "27-72"
      CCI$(18) = "124"
      CCI$(19) = "80"
      CCI$(20) = "150"
32044 For JJ% = 1 To 20
         GoSub 32052
         GoSub 32084
      Next JJ%
32046 CTX% = 1
32048 Return
32050 '
32052 For X2% = 1 To 8
32054 CCP$(X2%) = ""
32056 Next X2%
32058 X2% = 1
32060 '
32062 CON11$ = CCI$(JJ%)
32064 For X% = 1 To Len(CON11$)
32066 X1$ = Mid$(CON11$, X%, 1)
32068 If X1$ = " " Then GoTo 32078
32070 If X1$ = "-" Then X2% = X2% + 1: GoTo 32076
32072 If CCP$(X2%) = "" Then CCP$(X2%) = X1$: GoTo 32076
32074 CCP$(X2%) = CCP$(X2%) + X1$
32076 Next X%
32078 '
32080 Return
32082 '
32084 B$ = "": If JJ% = 20 Then B$ = MKI$(XVALO(CCP$(1))): GoTo 32094
32086 For XX% = 1 To 8
32088 BX$ = ""
      If Len(CCP$(XX%)) > 0 Then BX$ = Chr$(XVALO(CCP$(XX%)))
32090 B$ = B$ + BX$
32092 Next XX%
32094 CCI$(JJ%) = B$
32096 Return
      '
40000 End Sub
'

Sub SELEJER(EJUI%)
      '
5     LU$ = LUDAT$
      Dim COMEJ1%(2), FINEJ1%(2)
      EJU% = 1
      If VERCD% = 1 Then GoTo 90
      '
      Dim v As String * 32
      '
20    n1% = FILEOPEN%(LU$ + "VARIAB.DAT", 0, 32)
      Get #n1%, 2, v$
      EJACT1$ = TRIM$(Left$(v$, 28))
      COMEJ1%(1) = CVI(Mid$(v$, 29, 2))
      FINEJ1%(1) = CVI(Mid$(v$, 31, 2))
      '
      COMEPRO = COMEJ1%(1) + 32 * 13
      FINECAL% = DIANTE(COMEPRO)
      If FINEJ1%(1) <> FINECAL% Then
           VXX1$ = v$
           Mid$(VXX1$, 31, 2) = MKI$(FINECAL%)
           LSet v$ = VXX1$
           Put #n1%, 2, v$
           Close #n1%: n1% = 0
           GoTo 20
      End If
      '
      If COMEJ1%(1) = FECHANUM("01/08/97") Then
           VXX1$ = v$
           Mid$(VXX1$, 29, 2) = MKI$(FECHANUM("01/08/98"))
           Mid$(VXX1$, 31, 2) = MKI$(FECHANUM("31/07/99"))
           LSet v$ = VXX1$
           Put #n1%, 2, v$
           Close #n1%: n1% = 0
           GoTo 20
      End If
      '
      Close #n1%: n1% = 0
      '
30    n1% = FILEOPEN%(LU$ + "VARIAB.OLD", 0, 32)
      Get #n1%, 2, v$
      EJANT$ = TRIM$(Left$(v$, 28))
      COMEJ1%(2) = CVI(Mid$(v$, 29, 2))
      FINEJ1%(2) = CVI(Mid$(v$, 31, 2))
      '
      COMEPRO = COMEJ1%(2) + 32 * 13
      FINECAL% = DIANTE(COMEPRO)
      If FINEJ1%(2) <> FINECAL% Then
           VXX1$ = v$
           Mid$(VXX1$, 31, 2) = MKI$(FINECAL%)
           LSet v$ = VXX1$
           Put #n1%, 2, v$
           Close #n1%: n1% = 0
           GoTo 30
      End If
      Close #n1%: n1% = 0
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
      EJER% = Asc(LEERAM$(69, 1)): If EJER% <> 1 Then EJER% = (-1)
      COMEJ% = CVI(LEERAM$(87, 2))
      FINEJ% = CVI(LEERAM$(89, 2))
      EJUIX$ = EJACT$
      SE$ = "0"
      '
End Sub
'

Sub GRACOMEJ()
      '
105   LU$ = LUDAT$
      '
      Dim v As String * 32
      '
120   n1% = FILEOPEN%(LU$ + "VARIAB.DAT", 0, 32)
      Get #n1%, 2, v$
      CMJ% = CVI(Mid$(v$, 29, 2))
      FMJ% = CVI(Mid$(v$, 31, 2))
      '
      COMEPRO = CMJ% + 32 * 13
      FINECAL% = DIANTE(COMEPRO)
      If FMJ% <> FINECAL% Then
           VXX1$ = v$
           Mid$(VXX1$, 31, 2) = MKI$(FINECAL%)
           LSet v$ = VXX1$
           Put #n1%, 2, v$
           Close #n1%: n1% = 0
           GoTo 120
      End If
      '
      If CMJ% = FECHANUM("01/08/97") Then
           VXX1$ = v$
           Mid$(VXX1$, 29, 2) = MKI$(FECHANUM("01/08/98"))
           Mid$(VXX1$, 31, 2) = MKI$(FECHANUM("31/07/99"))
           LSet v$ = VXX1$
           Put #n1%, 2, v$
           Close #n1%: n1% = 0
           GoTo 120
      End If
      '
      Close #n1%: n1% = 0
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
      Dim v As String * 32
      n1% = FILEOPEN%(LU$ + "VARIAB." + EXTE$, 9, 32)
      Get #n1%, 2, v$
      EJACT$ = TRIM$(Left$(v$, 28))
      Close #n1%: n1% = 0
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
         For kkj% = 1 To 99
           CAPTI$ = TRIM$(Forms(IFORMU% - 1).Controls(KKI% - 1).Caption(kkj%))
           Forms(IFORMU% - 1).Controls(KKI% - 1).Caption(kkj%) = TRANSLATE$(CAPTI$)
         Next kkj%
         On Error GoTo 0
         '
20       TOTIP$ = ""
         On Error GoTo 92
         TOTIP$ = TRIM$(Forms(IFORMU% - 1).Controls(KKI% - 1).ToolTipText)
         Forms(IFORMU% - 1).Controls(KKI% - 1).ToolTipText = TRANSLATE$(TOTIP$)
         For kkj% = 1 To 99
           TOTIP$ = TRIM$(Forms(IFORMU% - 1).Controls(KKI% - 1).ToolTipText(kkj%))
           Forms(IFORMU% - 1).Controls(KKI% - 1).ToolTipText(kkj%) = TRANSLATE$(TOTIP$)
         Next kkj%
         On Error GoTo 0
         '
30       CAPTI$ = ""
         On Error GoTo 93
         For kkj% = 0 To 99
           CAPTI$ = TRIM$(Forms(IFORMU% - 1).Controls(KKI% - 1).TabCaption(kkj%))
           Forms(IFORMU% - 1).Controls(KKI% - 1).TabCaption(kkj%) = TRANSLATE$(CAPTI$)
         Next kkj%
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

Sub MUESTRA_FORMINI()
   '
   If Left$(UCase$(App.EXEName), 5) = "WINTA" Then Exit Sub
   If Left$(UCase$(App.EXEName), 7) = "FLEXO09" Then Exit Sub
   If Left$(UCase$(App.EXEName), 7) = "FLEXBRE" Then Exit Sub
   If Left$(UCase$(App.EXEName), 6) = "DSGFRM" Then Exit Sub
   '
   If App.EXEName <> "FLSETUP" Or Command$ = "" Then
     If VENCIAB% = 0 Then
       If Forms.Count >= 1 Then
         '
         Call TRANSLABELS(Forms(0).Name)
         '
         If Screen.Height > 9000 Then
           If Forms(0).StartUpPosition = 0 Then
             If Top = 0 Then
               Forms(0).Top = (Screen.Height - Forms(0).Height) / 2
               Forms(0).Left = (Screen.Width - Forms(0).Width) / 2
             End If
           End If
         End If
         '
         If Left$(TRIM$(Command$), 1) = "#" Then
            RUTXX$ = LUBAS$
            If INTERPRE% > 0 Then RUTXX$ = LUDAT$
            ARCHIVOSKIN$ = TRIM$(Left$(LOADFILE$(RUTXX$ + "FLEX_SKIN.CFG"), 512))
            If ARCHIVOSKIN = "" Then
               Forms(0).Show
               On Error Resume Next
               Forms(0).Refresh
               On Error GoTo 0
            End If
         End If
         '
       End If
     End If
   End If
   '
   
End Sub

Sub RECONECTA()
   '
   ' CONECTA CON EL MISMO MODULO
   Call CIERRARCH("*.*")
   Call BLOQEXE("*.*")
   Call BAJALDISCO
   If CURDINI$ = "" Then CURDINI$ = CurDir
   ChDrive Left$(CURDINI$, 2)
   ChDir CURDINI$
   Call CONECRUN(App.EXEName, "")
   Call FINAL("")
   Close: End
   '
End Sub

Function VALIDAUSER$()
    '
1   INICIO07.Show 1
    '
    If INICIO07.OK = False Then Exit Function
    '
    USUX$ = TRIM$(UCase(INICIO07.Text1))
    CONTX$ = TRIM$(UCase(INICIO07.Text2))
    '
    RUTAWIN$ = App.Path
    RUTAWIN$ = "X:\GB_GABAL\FLEXOFT\WINDOWS"
    n1% = FILEOPEN%(RUTAWIN$ + "\USER.DAT", 9, 64)
    Dim F1 As String * 64
    '
    For i% = 1 To LOF(n1%) / 64
       Get #n1%, i%, F1$
       If Asc(Left$(F1$, 1)) = 0 Then GoTo 99
       NOM$ = Left$(F1$, 30)
       If Left$(NOM$, 1) = "@" Then
           NOM$ = DESENCRI$(Mid$(F1$, 2, 29))
       End If
       '
       IDE$ = Mid$(F1$, 31, 12)
       If Left$(IDE$, 1) = "@" Then
           IDE$ = DESENCRI$(Mid$(F1$, 32, 11))
       End If
       CLA$ = Mid$(F1$, 43, 12)
       If Left$(CLA$, 1) = "@" Then
           CLA$ = DESENCRI$(Mid$(F1$, 44, 11))
       End If
       '
       If USUX$ = TRIM(IDE$) Then
          If CONTX$ = TRIM(CLA$) Then
             VALIDAUSER = NOM$
             Close #n1%
             Exit Function
          End If
       End If
       '
    Next i%
99  '
    Close #n1%
    VALIDAUSER$ = ""
    Call COMUNI("Nombre o Contraseña Incorrecta")
    GoTo 1
End Function


