Attribute VB_Name = "FILE_DRV"
Public SEVOLF$
Public VERDEFLOG%
'
Function EXISTE%(ARCHX$)
    '
    ARXX$ = TRIM$(ARCHX$)
    If InStr(ARXX$, ".") < 1 Then ARXX$ = ARXX$ + ".DAT"
    ESUNDIR% = 0
    If Right$(ARXX$, 2) = "\." Then
        ARXX$ = Left$(ARXX$, Len(ARXX$) - 2)   ' PARA DIRECTORIOS
        ESUNDIR% = 1
    End If
    '
    ARZZ$ = UCase$(ARXX$)
    If InStr(ARZZ$, ".RFS") > 0 Or InStr(ARZZ$, ".PRF") > 0 Then
      Call GRATABCONTRO(ARZZ$)
    End If
    '
    EXI% = 0
    If TRIM$(ARXX$) <> "" Then
      On Error Resume Next
      If ESUNDIR% = 1 Then
           EXISZ$ = TRIM$(Dir(ARXX$, vbDirectory))
         Else
           EXISZ$ = TRIM$(Dir(ARXX$))
      End If
      If Err < 1 Then
        If EXISZ$ <> "" Then
          EXI% = 1
          '
          If InStr(ARXX$, ".") > 0 Then
            If FileLen(ARXX$) < 1 Then
              EXI% = 0
            End If
          End If
          '
        End If
      End If
    End If
    '
    EXISTE% = EXI%
    On Error GoTo 0
    '
End Function
'
Function ExisteArchivo%(Ruta As String, NombreDeArchivo As String)
    
    Dim RutaA As String
    Dim Resultado As String
    
    RutaA = Ruta
    If Mid$(Ruta, Len(Ruta), 1) <> "\" Then RutaA = Ruta & "\"
    Resultado = Dir(RutaA & NombreDeArchivo)
    
    ExisteArchivo% = 0
    If Resultado <> "" Then ExisteArchivo% = 1
    
End Function

Function LOADFILE$(FINA$)
   '
   LFX$ = ""
   SCMA = Screen.MousePointer
   Screen.MousePointer = 11
   '
   On Error GoTo 97
   LOPUCO& = FileLen(FINA$)
   NX% = FILEOPEN%(FINA$, 41, 2048)
   LOPUCF& = LOF(NX%)
   If LOPUCF& > LOCUPO& Then LOPUCO& = LOPUCF&
   '
   LFX$ = String$(LOPUCO&, 0)
   Get #NX%, 1, LFX$
   GoTo 99
   '
97 Resume 98
98 LFX$ = ""
   '
99 If NX% > 0 Then Close #NX%
   On Error GoTo 0
   Screen.MousePointer = SCMA
   LOADFILE$ = LFX$
   '
End Function

Function LOADSUBFILE$(FINA$, POSINIC&)
   '
   LFX$ = ""
   SCMA = Screen.MousePointer
   Screen.MousePointer = 11
   PINIX& = POSINIC&: If PINIX& < 1 Then PINIX& = 1
   '
   On Error Resume Next
   LOPUCO& = FileLen(FINA$)
   On Error GoTo 0
   '
   If LOPUCO& > 0 Then
     If PINIX& <= LOPUCO& Then
       NX% = FILEOPEN%(FINA$, 41, 2048)
       LFX$ = String$(1 + LOPUCO& - PINIX&, 0)
       Get #NX%, PINIX&, LFX$
     End If
   End If
   '
99 If NX% > 0 Then Close #NX%
   On Error GoTo 0
   Screen.MousePointer = SCMA
   '
   LOADSUBFILE$ = LFX$
   '
End Function

Sub SAVEFILE(FINA$, LFX$)
   '
   SCMA = Screen.MousePointer
   Screen.MousePointer = 11
   '
   ' LO QUE SIGUE ES PARA QUE EN TERMINAL SERVER AL TENER PROBLEMAS DE ACCEDER A
   ' C:\FLEXOFT AL CONECTARSE CON FLOGON SALGA ORDENADAMENTE - EPB - 7 MAYO 2005
   If InStr(UCase$(FINA$), "RCONTROL.DRV") > 0 Then
     NX% = FreeFile
     On Error Resume Next
     Open FINA$ For Output As #NX%
     If Err Then
       On Error GoTo 0
       GoTo 99
     End If
     Close #NX%
     On Error GoTo 0
   End If
   ' HASTA AQUI EPB 7 DE MAYO 2005
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
Function FILEOPEN%(ARCHX$, MODO%, LOREGG%)
    '
    MODACCE$ = "ASIGNACION NUMERO DE ARCHIVO"
    On Error GoTo 99
    FILI% = FreeFile
    Select Case MODO%
      Case 0
        MODACCE$ = "ACCESO ALEATORIO"
        Open ARCHX$ For Random Access Read Write Shared As #FILI% Len = LOREGG%
        GoTo 89
      Case 10
        MODACCE$ = "ACCESO ALEATORIO SOLO LECTURA"
        Open ARCHX$ For Random Access Read Shared As #FILI% Len = LOREGG%
        GoTo 89
      Case 1
        MODACCE$ = "LECTURA SECUENCIAL"
        Open ARCHX$ For Input Shared As #FILI%
        GoTo 89
      Case 2
        MODACCE$ = "ESCRITURA SECUENCIAL"
        Open ARCHX$ For Output As #FILI%
        GoTo 89
      Case 3
        MODACCE$ = "SALIDA SECUENCIAL INCREMENTAL"
        Open ARCHX$ For Append As #FILI%
        GoTo 89
      Case 4
        MODACCE$ = "ACCESO BINARIO"
        Open ARCHX$ For Binary Access Read Write Shared As #FILI%
        GoTo 89
      Case 41
        MODACCE$ = "ACCESO BINARIO SOLO LECTURA" ' LOADFILE
        Open ARCHX$ For Binary Access Read Shared As #FILI%
        GoTo 89
     Case 9
        MODACCE$ = "ACCESO ALEATORIO SOLO LECTURA"
        Open ARCHX$ For Random Access Read Shared As #FILI% Len = LOREGG%
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
99  ERNU$ = TRIM$(str$(Err.Number))
    '
100 BRXX$ = REPLACAR$(ARCHX$, "\", "/")
    If Mid$(BRXX$, 2, 1) <> ":" Then
      If Left$(BRXX$, 2) <> "//" Then
        If Mid$(BRXX$, 2, 1) <> ":" Then
          BRXX$ = Left$(CurDir, 2) + BRXX$
        End If
      End If
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
    c$ = FORMESS$(MENSA$)
    RESPU% = MsgBox(c$, 53, "Error de Apertura de Archivo")
    If RESPU% = 4 Then
        Resume
    End If
    Close: End
    '
End Function

Function FILEDATE#(ARCHX$)
   '
   FDT# = 0
   On Error Resume Next
   FDT# = FileDateTime(ARCHX$)
   On Error GoTo 0
   '
   FILEDATE# = FDT#
   '
End Function

Sub DOSCOMMAND(CMDX$)
   '
'''   TTXX$ = CMDX$ + Chr$(10) + Chr$(13) + "EXIT"
'''   TTXXVIE$ = LOADFILE$(LUCOM$ + "DOSCOMM.BAT")
'''   If TTXX$ <> TTXXVIE$ Then
'''      Call SAVEFILE(LUCOM$ + "DOSCOMM.BAT", TTXX$)
'''   End If
'''   On Error Resume Next
'''   NNXC = Shell(LUCOM$ + "DOSCOMM.BAT", 0)
'''   On Error GoTo 0
'''   '
'''   ' esperar un segundo
'''   TINI = Timer
'''   OPENFORMS = DoEvents
'''   While Timer < TINI + 1
'''     DoEvents
'''   Wend
'''   '
   TTXX$ = CMDX$ + Chr$(10) + Chr$(13) + "EXIT"
   '
   For i% = 0 To 9
     IX$ = TRIM$(CStr(i%))
     ADOSK$ = "DOSCOMM" + IX$ + ".BAT"
     TTYY$ = LOADFILE$(LUCOM$ + ADOSK$)
     If TTYY$ = TTXX$ Then GoTo 31        ' SI COINCIDE USA EL MISMO
     '
     On Error Resume Next
     Kill LUCOM$ + ADOSK$
     On Error GoTo 0
     If EXISTE%(LUCOM$ + ADOSK$) < 1 Then GoTo 30
19 Next i%
   '
20 TTXX$ = "Error Fatal al Borrar la Serie de Archivos 'DOSCOMM*.BAT'"
   TTXX$ = TTXX$ + vbCrLf + "con '*' Variando de 0 (cero) a 9 (nueve) en la Carpeta"
   TTXX$ = TTXX$ + vbCrLf + LUCOM$ + "."
   
   TTXX$ = TTXX$ + vbCrLf + vbCrLf + "Estos Archivos Deben Borrarse Previo a la Ejecución de un"
   TTXX$ = TTXX$ + vbCrLf + "Comando DOS en la Aplicación '" & App.EXEName & ".EXE'."
   TTXX$ = TTXX$ + vbCrLf + vbCrLf + "Intente el Borrado Manual en la Carpeta Referida o Consulte"
   TTXX$ = TTXX$ + vbCrLf + "Con Mesa de Ayuda FLEXOFT."
   '
   MsgBox TTXX$
   '
   Close: End
   '
30 Call SAVEFILE(LUCOM$ + ADOSK$, TTXX$)
   '
31 On Error Resume Next
   NNXC = Shell(LUCOM$ + ADOSK$, 0)
   On Error GoTo 0
   '
   ' esperar un segundo
   TINI = Timer
   Openforms = DoEvents
   While Timer < TINI + 1
     DoEvents
   Wend
   '
   On Error Resume Next
   Kill LUCOM$ + ADOSK$
   On Error GoTo 0
   '
End Sub

Function SERIVOL$(ULOVOL$)
   '
   If VERCD% = 1 Or VERDEFLOG% > 0 Then
      SERIVOL$ = "0000-0000"
      Exit Function
   End If
   '
   SERIVOL$ = "": TIN00 = Timer
   TANTE# = FILEDATE#(LUCOM$ + "SERIVOL")
10 Call DOSCOMMAND("VOL " + ULOVOL$ + " > " + LUCOM$ + "SERIVOL")
   TINI = Timer
   While Timer < TINI + 0.2: DoEvents: Wend
   '
   If FILEDATE#(LUCOM$ + "SERIVOL") <> TANTE# Then
     SVLX$ = LOADFILE$(LUCOM$ + "SERIVOL")
     If Len(SVLX$) > 0 Then
       GoTo 20
     End If
   End If
   If Timer < TIN00 + 10 Then
     GoTo 10
   End If
   '
20 XX$ = LOADFILE$(LUCOM$ + "SERIVOL")
   Screen.MousePointer = SCMA
   POSI% = InStr(XX$, "-")
   If POSI% > 5 Then
     If POSI% <= Len(XX$) - 4 Then
       If Mid$(XX$, POSI% - 5, 1) = " " Then
         SERIVOL$ = Mid$(XX$, POSI% - 4, 9)
         If UCase$(ULOVOL$) = "C:" Or UCase$(ULOVOL$) = "D:" Then
           SEVOLX$ = LOADFILE$(ULOVOL$ + "\FLEXOFT\SERIVOL.DRV")
           If ENCRIPTA$(SERIVOL$) <> SEVOLX$ Then
             Call SAVEFILE(ULOVOL$ + "\FLEXOFT\SERIVOL.DRV", ENCRIPTA$(SERIVOL$))
           End If
         End If
         Exit Function
       End If
     End If
   End If
   '
   On Error Resume Next
   MkDir ULOVOL$ + "\FLEXOFT"
   On Error GoTo 0
   '
15 SEVOLX$ = LOADFILE$(ULOVOL$ + "\FLEXOFT\SERIVOL.DRV")
   SERIVOL$ = TRIM$(DESENCRI$(SEVOLX$))
   If SERIVOL$ <> "" Then
     Exit Function
   End If
   '
   CLAX$ = "": CLACO% = 0
30 PPX% = Int(Rnd * 63)
   If PPX% < 48 Or PPX% > 57 Then GoTo 30
   CLAX$ = CLAX$ + Chr$(PPX%)
   If Len(CLAX$) < 6 Then GoTo 30
   '
   RESPUES$ = CONTRACLA$(CLAX$)
   CONTRASE$ = InputBox$("Ingrese Contraseña", "Santo y Seña (" + CLAX$ + ")")
   If CONTRASE$ = RESPUES$ Then
     SEVOLX$ = Left$(CLAX$, 4) + "-" + Left$(CONTRASE$, 4)
     Call SAVEFILE(ULOVOL$ + "\FLEXOFT\SERIVOL.DRV", ENCRIPTA$(SEVOLX$))
     GoTo 15
   End If
   '
99 MsgBox "Imposible Identificar Disco '" + ULOVOL$ + "'."
   Close: End
   '
End Function
'
Function SERIVOLVIE$(ULOVOL$)
    '
    Static SVLEIDO$(16), CASVL%
    '
    If UCase$(ULOVOL$) >= "C" Then
      For UI% = 1 To CASVL%
        If TRIM$(Left$(SVLEIDO$(UI%), 4)) = ULOVOL$ Then
          SERIVOLVIE$ = TRIM$(Mid$(SVLEIDO$(UI%), 5))
          Exit Function
        End If
      Next UI%
    End If
    '
    Dim AA$(32)
    SERIVOLVIE$ = ""
    '
    AHORAA# = CDbl(Now)
    NX% = FreeFile
    On Error GoTo 98
    Open LUCOM$ + "DOSCOMM.BAT" For Output As #NX%
    Print #NX%, "VOL " + ULOVOL$ + " >C:\FLEXOFT\SERIVOL"
    Close #NX%: NX% = 0
    On Error GoTo 0
    '
    Shell LUCOM$ + "DOSCOMM.PIF"
    TINI = Timer
    While Timer < TINI + 3
      If FILEDATE#(LUCOM$ + "SERIVOL") > AHORAA# Then GoTo 20
    Wend
    '
20  NX% = FreeFile
    Open LUCOM$ + "SERIVOL" For Input As #NX%
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
    If VERCD% = 1 Or VERDEFLOG% > 0 Then
      SERIVOLVIE$ = "0000-0000"
      CASVL% = CASVL% + 1
      If CASVL% <= 16 Then SVLEIDO$(CASVL%) = Left$(ULOVOL$ + Space$(4), 4) + SERIVOLVIE$
      Exit Function
    End If
    '
    For AY% = 1 To AX%
      XX$ = TRIM$(AA$(AY%))
      POSI% = InStr(XX$, "-")
      If POSI% > 5 Then
        If POSI% <= Len(XX$) - 4 Then
          If Mid$(XX$, POSI% - 5, 1) = " " Then
            SERIVOLVIE$ = Mid$(XX$, POSI% - 4, 9)
            CASVL% = CASVL% + 1
            If CASVL% <= 16 Then SVLEIDO$(CASVL%) = Left$(ULOVOL$ + Space$(4), 4) + SERIVOLVIE$
            If UCase$(ULOVOL$) = "C:" Or UCase$(ULOVOL$) = "D:" Then
              SEVOLX$ = LOADFILE$(ULOVOL$ + "\FLEXOFT\SERIVOL.DRV")
              If ENCRIPTA$(SERIVOLVIE$) <> SEVOLX$ Then
                Call SAVEFILE(ULOVOL$ + "\FLEXOFT\SERIVOL.DRV", ENCRIPTA$(SERIVOLVIE$))
              End If
            End If
            Exit Function
          End If
        End If
      End If
    Next AY%
    '
15  SEVOLX$ = LOADFILE$(ULOVOL$ + "\FLEXOFT\SERIVOL.DRV")
    SERIVOLVIE$ = TRIM$(DESENCRI$(SEVOLX$))
    If SERIVOLVIE$ <> "" Then
      CASVL% = CASVL% + 1
      If CASVL% <= 16 Then SVLEIDO$(CASVL%) = Left$(ULOVOL$ + Space$(4), 4) + SERIVOLVIE$
      Exit Function
    End If
    '
    CLACO% = 0
    CLAX$ = ""
30  PPX% = Int(Rnd * 63)
    If PPX% < 48 Or PPX% > 57 Then GoTo 30
    CLAX$ = CLAX$ + Chr$(PPX%)
    If Len(CLAX$) < 6 Then GoTo 30
    '
    RESPUES$ = ""
    For l% = 1 To 6
      POSI% = (5 * l%) Mod 6
      While POSI% < 1: POSI% = POSI% + 6: Wend
      VORI% = XVALO(Mid$(CLAX$, POSI%, 1))
      VTRA% = (VORI% + 3 * l%) Mod 10
      RESPUES$ = RESPUES$ + Mid$(str$(VTRA%), 2)
    Next l%
    '
    CONTRASE$ = InputBox$("Ingrese Contraseña", "Santo y Seña (" + CLAX$ + ")")
    If CONTRASE$ = RESPUES$ Then
      SEVOLX$ = Left$(CLAX$, 4) + "-" + Left$(CONTRASE$, 4)
      Call SAVEFILE(ULOVOL$ + "\FLEXOFT\SERIVOL.DRV", ENCRIPTA$(SEVOLX$))
      GoTo 15
    End If
    '
    GoTo 99
    '
98  Resume 99
99  On Error GoTo 0
    MsgBox "Imposible Identificar Disco '" + ULOVOL$ + "'."
    Close: End
    '
End Function

Function OBJPLANUE$(OBX$, VDEF$)
   '
   If FOBJPLA.Visible = True Then
      PPXX% = InStr(OBX$, "/")
      If PPXX% < 2 Then PPXX% = 1 + Len(OBX$)
      OBY$ = Left$(OBX$, PPXX% - 1)
      MsgBox "Imposible Presentar Formulario '" + TRIM$(OBY$) + "'" + CRLF$ + "en el Contexto Actual del Sistema.", 48, "Informacion Critica"
      Exit Function
   End If
   '
   OBJECT_DEFVAL = VDEF$
   FOBJPLA.ObjectName = OBX$
   FOBJPLA.Show 1
   OBJPLANUE$ = OBJECT_RETVAL
   '
End Function

Function FILEFIND$(ULO$, ABUS$, MODO%)
   '
   SCMA = Screen.MousePointer
   '
   On Error Resume Next
   abc1$ = FileSearch.Label1
   If Err Then
     MsgBox "Imposible Ejecutar 'FileFind' - Consulte"
     FILEFIND$ = ""
     Exit Function
   End If
   '
   FFII$ = ""
   ESTEQ$ = ULO$
   Screen.MousePointer = 11
   '
   ABUSCA$ = ABUS$
   FileSearch.Dir8.Path = ESTEQ$ + "\"
   FileSearch.File8.Path = FileSearch.Dir8.Path
   FileSearch.File8.Pattern = ABUSCA$
   DIEXPLO$ = ESTEQ$ + "\"
   '
5  ' CARGA EN LIST1 LA LISTA DE LAS CARPETAS A EXPLORAR *************
   FileSearch.List8.Clear
   ILISTA& = 0
   FileSearch.List8.AddItem DIEXPLO$
   '
10 FileSearch.Dir8.Path = FileSearch.List8.List(ILISTA&)
   FileSearch.Dir8.Refresh
   DoEvents
   For KKI% = 1 To FileSearch.Dir8.ListCount
     TTXX$ = FileSearch.Dir8.List(KKI% - 1) + "\"
     FileSearch.List8.AddItem TTXX$
     FileSearch.List8.ListIndex = FileSearch.List8.ListCount - 1
     FileSearch.List8.Refresh
     DoEvents
   Next KKI%
   '
   If ILISTA& < FileSearch.List8.ListCount - 1 Then
     ILISTA& = ILISTA& + 1
     GoTo 10
   End If
   '
   For ILISTA& = 1 To FileSearch.List8.ListCount
     FileSearch.List8.ListIndex = ILISTA& - 1
     FileSearch.File8.Path = FileSearch.List8.List(ILISTA& - 1)
     DoEvents
     FileSearch.File8.Pattern = ABUSCA$
     '
     For KKI% = 1 To FileSearch.File8.ListCount
       FINAME$ = TRIM$(UCase$(FileSearch.File8.List(KKI% - 1)))
       If FINAME$ <> "" Then
         ACOMPLE$ = FileSearch.List8.List(ILISTA& - 1) + FINAME$
         FILEFIND$ = ACOMPLE$
         Screen.MousePointer = SCMA
         Exit Function
       End If
29   Next KKI%
   Next ILISTA&
   Screen.MousePointer = SCMA
   '
End Function

Sub RECIBARJ(ANAME$)
   '
   If EXISTE%(ANAME$) > 0 Then
     ULODAX$ = Left$(LUDAT$, 2)
     CUDI$ = CurDir
     TTXX$ = ULODAX$ + Chr$(10) + Chr$(13)
     TTXX$ = TTXX$ + "CD \FLEXOFT\NOVEDAD" + Chr$(10) + Chr$(13)
     TTXX$ = TTXX$ + "ARJ E " + ANAME$ + Chr$(10) + Chr$(13)
     TTXX$ = TTXX$ + Left$(CUDI$, 2) + Chr$(10) + Chr$(13)
     TTXX$ = TTXX$ + "CD " + Mid$(CUDI$, 3) + Chr$(10) + Chr$(13)
     TTXX$ = TTXX$ + "EXIT"
     Call SAVEFILE(LUCOM$ + "DOSCOMM.BAT", TTXX$)
     On Error Resume Next
     NNXC = Shell(LUCOM$ + "DOSCOMM.BAT", 0)
     On Error GoTo 0
     '
     ' esperar un segundo
     TINI = Timer
     Openforms = DoEvents
     While Timer < TINI + 1
       DoEvents
     Wend
   End If
   '
End Sub

Sub GRATABCONTRO(CONTRONAME$)
   '
   Static GRACO%, APPNAME$, VECON$(), CACON&
   '
   If GRACO% < 0 Then Exit Sub
   '
   HAYMALOS% = 0
   If GRACO% = 0 Then
     GRACO% = (-1)
     EXIX% = EXISTE%("F:\FLEXOFT\NOVEDAD\CTRLPRFS\IAPPPRFS.DRV")
     If EXIX% > 0 Then
       GRACO% = 1
     End If
     If GRACO% < 0 Then Exit Sub
     '
     APPNAME$ = App.EXEName
     '
     VECONTRO$ = LOADFILE$("F:\FLEXOFT\NOVEDAD\CTRLPRFS\" + APPNAME$ + ".TCR")
     CACON& = Len(VECONTRO$) / 32
     ReDim VECON$(CACON& + 128)
     U2& = 0
     For U& = 1 To CACON&
       XCON$ = TRIM$(UCase$(Mid$(VECONTRO$, 32 * U& - 31, 30)))
       If XCON$ <> "" Then
         If Left$(XCON$, 4) <> "TRAN" Then
           If InStr(XCON$, ".RFS") > 0 Or InStr(XCON$, ".PRF") > 0 Then
             U2& = U2& + 1
             VECON$(U2&) = XCON$
           End If
         End If
       End If
       HAYMALOS% = 1
19   Next U&
     CACON& = U2&
   End If
   '
   CONAME$ = TRIM$(UCase$(CONTRONAME$))
20 PPXX% = InStr(CONAME$, "\")
   If PPXX% > 0 Then
     CONAME$ = Mid$(CONAME$, PPXX% + 1)
     GoTo 20
   End If
   '
   If Left$(CONAME$, 4) = "TRAN" Then GoTo 90
   '
   For U& = 1 To CACON&
     If VECON$(U&) = CONAME$ Then GoTo 90
   Next U&
   '
   If CACON& >= UBound(VECON$) - 10 Then
     ReDim Preserve VECON$(CACON& + 128)
   End If
   '
   CACON& = CACON& + 1
   VECON$(CACON&) = CONAME$
   '
80 VECONTRO$ = ""
   For U& = 1 To CACON&
     VECONTRO$ = VECONTRO$ + AJUSTI$(VECON$(U&), 30) + CRLF$
   Next U&
   Call SAVEFILE("F:\FLEXOFT\NOVEDAD\CTRLPRFS\" + APPNAME$ + ".TCR", VECONTRO$)
   HAYMALOS% = 0
   '
90 If HAYMALOS% > 0 Then GoTo 80
   End Sub

Function PARENTFOLDER$(CARPETA$)
   '
   PFX$ = ""
   CARPETB$ = CARPETA$
   If Right$(CARPETA$, 1) = "\" Then CARPETB$ = Left$(CARPETA$, Len(CARPETA$) - 1)
   '
   For KKI% = Len(CARPETA$) To 1 Step -1
      If Mid$(CARPETA$, KKI%, 1) = "\" Then
         PFX$ = Left$(CARPETA$, KKI% - 1)
         GoTo 90
      End If
   Next KKI%
   '
90 PARENTFOLDER$ = PFX$
   '
End Function

