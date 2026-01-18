Attribute VB_Name = "FILE_DRV1"
Public SEVOLF$
Public VERDEFLOG%
'
Function EXISTE%(ARCHX$)
    '
    ARXX$ = Trim$(ARCHX$)
    If InStr(ARXX$, ".") < 1 Then ARXX$ = ARXX$ + ".DAT"
    ESUNDIR% = 0
    If Right$(ARXX$, 2) = "\." Then
        ARXX$ = Left$(ARXX$, Len(ARXX$) - 2)   ' PARA DIRECTORIOS
        ESUNDIR% = 1
    End If
    '
    EXI% = 0
    If Trim$(ARXX$) <> "" Then
      On Error Resume Next
      If ESUNDIR% = 1 Then
           EXISZ$ = Trim$(Dir(ARXX$, vbDirectory))
         Else
           EXISZ$ = Trim$(Dir(ARXX$))
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
        MODACCO$ = "SALIDA SECUENCIAL INCREMENTAL"
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
99  ERNU$ = Trim$(Str$(Err.Number))
    '
100 BRXX$ = REPLACAR$(ARCHX$, "\", "/")
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

Sub DOSCOMMAND_VIE(CMDX$)
   '
' MsgBox "ESTOY EN DOSCOMMAND CON LUCOM$ = " & LUCOM$
' MsgBox "EL COMANDO ES " & CMDX$
   ' lo que sigue se puso por CIVAL para dar la pista de cual es el problema del DOSCOMM
   If EXISTE%(LUCOM$ + "DOSCOMM.BAT") > 0 Then
      On Error Resume Next
      Kill LUCOM$ + "DOSCOMM.BAT"
      If Err Then
         On Error GoTo 0
         ABORRA$ = LUCOM$ + "DOSCOMM.BAT"
         If Mid$(ABORRA$, 3) <> "\" Then
            ABORRA$ = App.Path + "\DOSCOMM.BAT"
         End If
         MsgBox "No se Pudo Borrar '" + ABORRA$ + "'." + vbCrLf + vbCrLf + "Causa Probable que el Archivo Sea de Solo Lectura o Bloqueado por Otra Aplicación." + vbCrLf + vbCrLf + "Recomendamos Intentar Borrado Manual y Luego Repetir la Transaccion."
         Close
         End
      End If
      On Error GoTo 0
   End If
   '
   TTXX$ = CMDX$ + Chr$(10) + Chr$(13) + "EXIT"
   Call SAVEFILE(LUCOM$ + "DOSCOMM.BAT", TTXX$)
   On Error Resume Next
   NNXC = Shell(LUCOM$ + "DOSCOMM.BAT", 0)
   On Error GoTo 0
   '
   ' esperar un segundo
   TINI = Timer
   OPENFORMS = DoEvents
   While Timer < TINI + 1
     DoEvents
   Wend
   '
End Sub
'
Sub DOSCOMMAND(CMDX$)
   '
' MsgBox "ESTOY EN DOSCOMMAND CON LUCOM$ = " & LUCOM$
' MsgBox "EL COMANDO ES " & CMDX$
   '
   'SI LUCOM ESTA VACIO SE LE FUERZA LA RUTA
   LUCOM1$ = LUCOM$
   If LUCOM1$ = "" Then LUCOM1$ = App.Path + "\"
   '
   TTXX$ = CMDX$ + Chr$(10) + Chr$(13) + "EXIT"
   '
   For I% = 0 To 9
     IX$ = Trim$(CStr(I%))
     ADOSK$ = "DOSCOMM" + IX$ + ".BAT"
     TTYY$ = LOADFILE$(LUCOM1$ + ADOSK$)
     If TTYY$ = TTXX$ Then GoTo 31        ' SI COINCIDE USA EL MISMO
     '
     On Error Resume Next
     Kill LUCOM1$ + ADOSK$
     On Error GoTo 0
     If EXISTE%(LUCOM1$ + ADOSK$) < 1 Then GoTo 30
19 Next I%
   '
20 TTXX$ = "Error Fatal al Borrar la Serie de Archivos 'DOSCOMM*.BAT'"
   TTXX$ = TTXX$ + vbCrLf + "con '*' Variando de 0 (cero) a 9 (nueve) en la Carpeta"
   TTXX$ = TTXX$ + vbCrLf + LUCOM1$ + "."
   
   TTXX$ = TTXX$ + vbCrLf + vbCrLf + "Estos Archivos Deben Borrarse Previo a la Ejecución de un"
   TTXX$ = TTXX$ + vbCrLf + "Comando DOS en la Aplicación '" & App.EXEName & ".EXE'."
   TTXX$ = TTXX$ + vbCrLf + vbCrLf + "Intente el Borrado Manual en la Carpeta Referida o Consulte"
   TTXX$ = TTXX$ + vbCrLf + "Con Mesa de Ayuda FLEXOFT."
   '
   MsgBox TTXX$
   '
   Close: End
   '
30 Call SAVEFILE(LUCOM1$ + ADOSK$, TTXX$)
   '
31 On Error Resume Next
   NNXC = Shell(LUCOM1$ + ADOSK$, 0)
   On Error GoTo 0
   '
   ' esperar un segundo
   TINI = Timer
   OPENFORMS = DoEvents
   While Timer < TINI + 1
     DoEvents
   Wend
   '
   On Error Resume Next
   Kill LUCOM1$ + ADOSK$
   On Error GoTo 0
   '
End Sub

Function SERIVOL$(ULOVOL$)
   '
   Static SEVO$, ULOV$
   '
   If VERCD% = 1 Or VERDEFLOG% > 0 Then
      SERIVOL$ = "0000-0000"
      Exit Function
   End If
   '
   If ULOVOL$ = ULOV$ Then
      If SEVO$ <> "" Then
         SERIVOL$ = SEVO$
         Exit Function
      End If
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
         If Len(Trim$(SERIVOL$)) <> 9 Then GoTo 15 ' NO ES UN VALOR VÁLIDO
         If UCase$(ULOVOL$) >= "C:" Then
           SEVOLX$ = LOADFILE$(ULOVOL$ + "\FLEXOFT\SERIVOL.DRV")
           If ENCRIPTA$(SERIVOL$) <> SEVOLX$ Then
             Call SAVEFILE(ULOVOL$ + "\FLEXOFT\SERIVOL.DRV", ENCRIPTA$(SERIVOL$))
           End If
         End If
         SEVO$ = SERIVOL$
         ULOV$ = ULOVOL$
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
   SERIVOL$ = Trim$(DESENCRI$(SEVOLX$))
   If SERIVOL$ <> "" Then
      SEVO$ = SERIVOL$
      ULOV$ = ULOVOL$
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
Function OBJPLANUE$(OBX$, VDEF$)
   '
   If FOBJPLA.Visible = True Then
      PPXX% = InStr(OBX$, "/")
      If PPXX% < 2 Then PPXX% = 1 + Len(OBX$)
      OBY$ = Left$(OBX$, PPXX% - 1)
      crlf$ = Chr$(13) + Chr$(10)
      MsgBox "Imposible Presentar Formulario '" + Trim$(OBY$) + "'" + crlf$ + "en el Contexto Actual del Sistema.", 48, "Informacion Critica"
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
   ABC1$ = FILESEARCH.Label1
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
   FILESEARCH.Dir8.Path = ESTEQ$ + "\"
   FILESEARCH.File8.Path = FILESEARCH.Dir8.Path
   FILESEARCH.File8.Pattern = ABUSCA$
   DIEXPLO$ = ESTEQ$ + "\"
   '
5  ' CARGA EN LIST1 LA LISTA DE LAS CARPETAS A EXPLORAR *************
   FILESEARCH.List8.Clear
   ILISTA& = 0
   FILESEARCH.List8.AddItem DIEXPLO$
   '
10 FILESEARCH.Dir8.Path = FILESEARCH.List8.List(ILISTA&)
   FILESEARCH.Dir8.Refresh
   DoEvents
   For KKI% = 1 To FILESEARCH.Dir8.ListCount
     TTXX$ = FILESEARCH.Dir8.List(KKI% - 1) + "\"
     FILESEARCH.List8.AddItem TTXX$
     FILESEARCH.List8.ListIndex = FILESEARCH.List8.ListCount - 1
     FILESEARCH.List8.Refresh
     DoEvents
   Next KKI%
   '
   If ILISTA& < FILESEARCH.List8.ListCount - 1 Then
     ILISTA& = ILISTA& + 1
     GoTo 10
   End If
   '
   For ILISTA& = 1 To FILESEARCH.List8.ListCount
     FILESEARCH.List8.ListIndex = ILISTA& - 1
     FILESEARCH.File8.Path = FILESEARCH.List8.List(ILISTA& - 1)
     DoEvents
     FILESEARCH.File8.Pattern = ABUSCA$
     '
     For KKI% = 1 To FILESEARCH.File8.ListCount
       FINAME$ = Trim$(UCase$(FILESEARCH.File8.List(KKI% - 1)))
       If FINAME$ <> "" Then
         ACOMPLE$ = FILESEARCH.List8.List(ILISTA& - 1) + FINAME$
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
     TTXX$ = TTXX$ + "CD " + Mid$(CUDI$, 3)    ' + Chr$(10) + Chr$(13)
'     TTXX$ = TTXX$ + "EXIT"
'     Call SAVEFILE(LUCOM$ + "DOSCOMM.BAT", TTXX$)
'     On Error Resume Next
'     NNXC = Shell(LUCOM$ + "DOSCOMM.BAT", 0)
'     On Error GoTo 0
     Call DOSCOMMAND(TTXX$)
     '
     ' esperar un segundo
     TINI = Timer
     OPENFORMS = DoEvents
     While Timer < TINI + 1
       DoEvents
     Wend
   End If
   '
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
