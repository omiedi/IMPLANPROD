Attribute VB_Name = "WEB_FUNCTIONS"
Public Declare Function URLDownloadToFile Lib "urlmon" Alias _
"URLDownloadToFileA" (ByVal pCaller As Long, ByVal szURL As String, ByVal _
szFileName As String, ByVal dwReserved As Long, ByVal lpfnCB As Long) As Long
'
Public Declare Function FindFirstUrlCacheGroup Lib "wininet.dll" ( _
    ByVal dwFlags As Long, _
    ByVal dwFilter As Long, _
    ByRef lpSearchCondition As Long, _
    ByVal dwSearchCondition As Long, _
    ByRef lpGroupId As Date, _
    ByRef lpReserved As Long) As Long

Public Declare Function FindNextUrlCacheGroup Lib "wininet.dll" ( _
    ByVal hFind As Long, _
    ByRef lpGroupId As Date, _
    ByRef lpReserved As Long) As Long
    
Public Declare Function DeleteUrlCacheGroup Lib "wininet.dll" ( _
    ByVal sGroupID As Date, _
    ByVal dwFlags As Long, _
    ByRef lpReserved As Long) As Long
    
Public Declare Function FindFirstUrlCacheEntry Lib "wininet.dll" Alias "FindFirstUrlCacheEntryA" ( _
    ByVal lpszUrlSearchPattern As String, _
    ByRef lpFirstCacheEntryInfo As INTERNET_CACHE_ENTRY_INFO, _
    ByRef lpdwFirstCacheEntryInfoBufferSize As Long) As Long
    
Public Type INTERNET_CACHE_ENTRY_INFO
    dwStructSize As Long
    szRestOfData(1024) As Long
End Type

Public Declare Function DeleteUrlCacheEntry Lib "wininet.dll" Alias "DeleteUrlCacheEntryA" ( _
    ByVal lpszUrlName As Long) As Long

Public Declare Function FindNextUrlCacheEntry Lib "wininet.dll" Alias "FindNextUrlCacheEntryA" ( _
    ByVal hEnumHandle As Long, _
    ByRef lpNextCacheEntryInfo As INTERNET_CACHE_ENTRY_INFO, _
    ByRef lpdwNextCacheEntryInfoBufferSize As Long) As Long

Private Const CACHGROUP_SEARCH_ALL = &H0
Private Const ERROR_NO_MORE_FILES = 18
Private Const ERROR_NO_MORE_ITEMS = 259
Private Const CACHEGROUP_FLAG_FLUSHURL_ONDELETE = &H2
Private Const BUFFERSIZE = 4096
'
' ------------------------------------------------------------------------------
'
Public BADACON As Database
'
Public FIDATEX
'
Public NEPLI$      ' NOMBRE DE LA EMPRESA
Public IDEPLI$     ' IDENTIFICADOR - NUMERO DE SERIE DEL VOLUMEN
Public USERNU%     ' NUMERO DE USUARIO
Public USERNA$     ' NOMBRE DE USUARIO (COMPLETO)
Public USERID$     ' IDENTIFICACION DE USUARIO
'
Public CARPEDOWN$  ' CARPETA DEL DOWNLOAD
Public DIREPRO$    ' CARPETA DEL EJECUTABLE
Public DIRENEW$    ' CARPETA DE NOVEDADES
Public DIREVIE$    ' RESPALDO VERSIONES VIEJAS
'
Sub LEEDATSYS()
   '
'   FLEXLOG3.Label2 = 1711: DoEvents
   LFI$ = LOADFILE$(App.Path + "\IDENSYS.DRV")
'   FLEXLOG3.Label2 = 1712: DoEvents
   LFJ$ = DESENCRI$(Left$(LFI$, 96))
   NEPLI$ = TRIM$(Mid$(LFJ$, 13, 30))
   '
   UL$ = Left$(App.Path, 2)
   IDEPLI$ = SERIVOL$(UL$)
'   FLEXLOG3.Label2 = 1713: DoEvents
   '
   LFI$ = LOADFILE$("C:\FLEXOFT\FLEXRAM.DRV")
   USERNU% = CVI(Mid$(LFI$, 9, 2)) Mod 100
   If USERNU% < 1 Then USERNU% = 1
   '
'   FLEXLOG3.Label2 = 1714: DoEvents
   VUSER$ = LOADFILE$("USER.DAT")
'   FLEXLOG3.Label2 = 1715: DoEvents
   USERNA$ = "": USERID$ = ""
   F0$ = Mid$(VUSER$, 64 * USERNU% - 63, 64)
   USERID$ = TRIM$(Mid$(F0$, 31, 12))
      If Left$(USERID$, 1) = "@" Then USERID$ = TRIM$(DESENCRI$(Mid$(F0$, 32, 11)) + " ")
   USERNA$ = TRIM$(Left$(F0$, 30))
      If Left$(USERNA$, 1) = "@" Then USERNA$ = TRIM$(DESENCRI$(Mid$(F0$, 2, 29)) + " ")
   '
   ' DETERMINA LA CARPETA DEL DOWNLOAD
   CARPEDOWN$ = App.Path
   While Len(CARPEDOWN$) > 0 And Right$(CARPEDOWN$, 1) <> "\"
      CARPEDOWN$ = Left$(CARPEDOWN$, Len(CARPEDOWN$) - 1)
   Wend
   CARPEDOWN$ = CARPEDOWN$ + "DOWNLOAD\"
   ' CREA LA CARPETA
   On Error Resume Next
   MkDir Left$(CARPEDOWN$, Len(CARPEDOWN$) - 1)   ' PARA SUPRIMIR LA ULTIMA ANTIBARRA
   On Error GoTo 0
   '
   ' directorios EJECUTABLE - NOVEDADES - VIEJOS
   DIREPRO$ = App.Path
'   FLEXLOG3.Label2 = 1716: DoEvents
   If EXISTE%("F:\FLEXOFT\NOVEDAD\.") > 0 Then
       DIRENEW$ = "F:\FLEXOFT\NOVEDAD"
     ElseIf EXISTE%("C:\FLEXOFT\NOVEDAD\.") > 0 Then
       DIRENEW$ = "C:\FLEXOFT\NOVEDAD"
     Else
       DIRENEW$ = PARENTFOLDER$(DIREPRO$) + "\NOVEDAD"
   End If
   '
   ' CREA LA CARPETA \NOVEDAD
'   FLEXLOG3.Label2 = 1717: DoEvents
   On Error Resume Next
   MkDir DIRENEW$
   On Error GoTo 0
   '
   DIREVIE$ = DIREPRO$ + "\VIEJOS\"
   '
   ' CREA LA CARPETA \VIEJOS
'   FLEXLOG3.Label2 = 1718: DoEvents
   On Error Resume Next
   MkDir Left$(DIREVIE$, Len(DIREVIE$) - 1) ' PARA SUPRIMIR LA ULTIMA ANTIBARRA
   On Error GoTo 0
'   FLEXLOG3.Label2 = 1719: DoEvents
   '
End Sub

Sub CREABACONLEC()
   ' crea la base de control de lecturas
   Static CTX%
   If CTX% < 1 Then
      CTX% = 1
      '
      Call CREACAMPO_MDB("SQLCTRL", "CONTRO_DOWN", "File_Name", 10, 64)
      Call CREACAMPO_MDB("SQLCTRL", "CONTRO_DOWN", "File_DateTime", 10, 16)
      Call CREACAMPO_MDB("SQLCTRL", "CONTRO_DOWN", "User_Dest", 10, 16)
      Call CREACAMPO_MDB("SQLCTRL", "CONTRO_DOWN", "Num_Pack", 10, 8)
      Call CREACAMPO_MDB("SQLCTRL", "CONTRO_DOWN", "FeHor_Baja", 10, 16)
      Call CREACAMPO_MDB("SQLCTRL", "CONTRO_DOWN", "Usu_Baja", 10, 64)
      Call CREACAMPO_MDB("SQLCTRL", "CONTRO_DOWN", "FeHor_Lect", 10, 16)
      Call CREACAMPO_MDB("SQLCTRL", "CONTRO_DOWN", "Usu_Lect", 10, 64)
      '
   End If
   '
   On Error Resume Next
   Set BADACON = OpenDatabase(App.Path + "\SQLCTRL.MDB")
   On Error GoTo 0
   '
End Sub

Sub VERIACWEB(MODOI%)
   '
   ' ESTE PROCESO BAJA ARCHIVOS DE LA WEB FLEXOFT Y LOS COLOCA EN \FLEXOFT\DOWNLOAD
   ' ADEMÁS GENERA UN ARCHIVO DE PENDIENTES DE PROCESO EN \FLEXOFT\DOWNLOAD
   '
   ' MODOI%=0 ES EL MODO AUTOMATICO QUE VERIFICA HABILITACION DE
   '             LA TERMINAL PARA BAJAR ACTUALIZACIONES Y CANCELA
   '             SI NO HABILITADA
   ' MODOI%=1 ES QUE SE LLAMÓ DEL BOTÓN DE REFRESCAR Y FUERZA LA BAJADA
   '
   Dim JJ1 As Long
   Dim FIRECVIE$()
   Dim ARCHAC$(2048)
   '
   ' crea la carpeta de bajada de novedades CARPEDOWN
   CARPINS$ = Left$(CARPEDOWN$, Len(CARPEDOWN$) - 1)
   On Error Resume Next
   MkDir CARPINS$
   On Error GoTo 0
   If MODOI% = 1 Then MsgBox "Creada la Carpeta " + CARPEDOWN$
   '
   ' presenta el identificador de la empresa licenciataria
   If MODOI% = 1 Then MsgBox "Identificador: " + IDEPLI$
   '
   ' cancela si esta subrimido la bajada automatica
   If MODOI% <> 1 Then
      If PUEDEBAJARDELAWEB% < 1 Then GoTo 99
   End If
   '
   ' limpia la memoria cache
   Call BORRACACHE
   '
   ' baja la última version de IUPDATES.TXT
   JJ1& = BAJADELAWEB("IUPDATES.TXT")
   If MODOI% = 1 Then MsgBox "Termine Bajada de IUPDATES.TXT con JJ1& = " & JJ1&
   If JJ1& <> 0 Then Exit Sub
   '
   ' compara con el registrado en el directorio de trabajo
   FULTACT$ = LOADFILE$(App.Path + "\IUPDATES.TXT") ' LO LEE DEL DIRECTORIO DE TRABAJO
   FIDOWNL$ = LOADFILE$(CARPEDOWN$ + "IUPDATES.TXT")
   If FIDOWNL$ = FULTACT$ Then
     ' EL ARCHIVO BAJADO COINCIDE CON EL ANTERIOR - NO HAY NOVEDADES
     If MODOI% = 1 Then
          MsgBox "El IUPDATES.TXT  Bajado Coincide con el Anterior."
        Else
          GoTo 99
      End If
   End If
   '
   ' LEER LA TABLA DE NUEVAS ACTUALIZACIONES
   CANREGNUE& = Len(FIDOWNL$) / 128
   J2& = 0
   For U& = 1 To CANREGNUE&
      FIRECNU$ = Mid$(FIDOWNL$, 128 * U& - 127, 128)
      If TRIM$(Left$(FIRECNU$, 64)) <> "" Then
         EPLIDEST$ = TRIM$(Mid$(FIRECNU$, 87, 16))
         ' QUE EL NÚMERO DE SERIE DEL VOLUMEN COINCIDA O ESTÉ EN BLANCO
         If EPLIDEST$ <> "" Then
            If EPLIDEST$ <> IDEPLI$ Then
              GoTo 29    ' NUMERO DE SERIE DIFIERE
            End If
         End If
         '
         FINAME$ = TRIM$(Left$(FIRECNU$, 64))
         FEHO$ = TRIM$(Mid$(FIRECNU$, 65, 10))
         FIDATI$ = FETEX$(Val(Mid$(FEHO$, 1, 5))) + " " + HORTEX$(Val(Mid$(FEHO$, 6, 5)))
         UDEST1$ = TRIM$(Mid$(FIRECNU$, 75, 12))
         UDEST2$ = TRIM$(Mid$(FIRECNU$, 107, 12))
         NUCTRL$ = TRIM$(Mid$(FIRECNU$, 103, 4))
         'If MODOI% = 1 Then
         '   MsgBox "Antes de CargaLiBaja con " + FINAME$ + " " + FIDATI$
         'End If
         Call CARGALIBAJA(FINAME$, FIDATI$, UDEST1$, UDEST2$, NUCTRL$)
         'If MODOI% = 1 Then
         '   MsgBox "Despues de CargaLiBaja con " + FINAME$ + " " + FIDATI$
         'End If
         '
      End If
29 Next U&
   '
   ' APAREA EL ARCHIVO DE BAJADA
   Call SAVEFILE(App.Path + "\IUPDATES.TXT", FIDOWNL$)
   '
   ' VERIFICA SI PUEDE BAJAR
   If MODOI% <> 1 Then
      If PUEDEBAJARDELAWEB% < 1 Then GoTo 99
   End If
   '
   Screen.MousePointer = 11
   HUBOBAJA% = 0
   SQLX$ = "SELECT * FROM CONTRO_DOWN"
   SQLX$ = SQLX$ + " WHERE FeHor_Baja = ''"
   '
   Call CREABACONLEC
   On Error Resume Next
   Set MIRS = BADACON.OpenRecordset(SQLX$, dbOpenDynaset)
   MIRS.MoveFirst
   If Err Then
      BADACON.Close
      On Error GoTo 0
      GoTo 99                 ' no hay nada que bajar
   End If
   On Error GoTo 0
   '
   With MIRS
      Do While Not .EOF
        ARCHA$ = TRIM$(!File_Name)
        FIDATX$ = File_DateTime
        If MODOI% = 1 Then
           MsgBox "Voy a Bajar " + ARCHA$ + " " + FIDATX$
        End If
' MsgBox "*** Bajando " + ARCHA$
        JJ1& = BAJADELAWEB(ARCHA$)
        If JJ1& <> 0 Then
            If MODOI% = 1 Then MsgBox "Fue Imposible Bajar " & ARCHA$
          Else
            .Edit
            AHORA1$ = Format$(Now, "dd/mm/yy  hh:mm")
            !FeHor_Baja = AHORA1$
            !Usu_Baja = USERNA$
            If InStr(FINAME$, "BKSEVOL") > 0 Then
               ' MARCA DE LECTURA AUTOMATICA PARA ESTE CASO
               !FeHor_Lect = AHORA1$
               !Usu_Lect = USERNA$
            End If
            .Update
            HUBOBAJA% = 1
        End If
      .MoveNext
      Loop
   End With
   MIRS.Close
   BADACON.Close
   Screen.MousePointer = 1
   '
   If HUBOBAJA% > 0 Then
      On Error Resume Next
      FIDADIS = FileDateTime(App.Path + "\" + App.EXEName + ".EXE")
      On Error GoTo 0
      If FIDADIS <> FIDATEX Then
         ' re-inicia la aplicación
         XXX = Shell(App.Path + "\" + App.EXEName + ".EXE REFRESH")
         Close: End
         Exit Sub
      End If
      On Error Resume Next
      Call WEBROWSR.SHOWLINEWS
      If Err Then
         On Error GoTo 0
         MsgBox "Imposible Mostrar Lista de Novedades."
      End If
   End If
   '
99 Screen.MousePointer = 1
   '
End Sub

Sub CARGALIBAJA(FINAME$, FIDATI$, UDEST1$, UDEST2$, NUCTRL$)
    '
    ' LOS EXE SE BAJAN SOLO SI PRE-EXISTEN
    ' los exe se bajan para la organizacion, no para el destinatario
    ' VER CRITERIOS RESPECTO DE LOS RESTANTES TIPOS DE ARCHIVOS
    Call CREABACONLEC
    '
    UD1$ = UDEST1$: UD2$ = UDEST2$
    FINAM1$ = UCase$(TRIM$(Left$(FINAME$, 64)))
    NARCHI$ = FINAM1$
    If NARCHI$ = "FLXMAIL.EXE" Or NARCHI$ = "FLOGON.EXE" Or NARCHI$ = "FLX_BRWS.EXE" Or NARCHI$ = "BKSEVOL.DAT" Then GoTo 10
    '
    If InStr(FINAM1$, ".EXE") > 0 Then
      If EXISTE%(App.Path + "\" + FINAM1$) < 1 Then
' MsgBox "*** No existe " + App.Path + "\" + FINAM1$
         Exit Sub
      End If
      UD1$ = "": UD2$ = ""
    End If
    '
10  STRSQL = "SELECT * FROM CONTRO_DOWN"
    STRSQL = STRSQL + " WHERE File_Name = '" & FINAME$ & "'"
    STRSQL = STRSQL + " AND File_DateTime = '" & FIDATI$ & "'"
    STRSQL = STRSQL + " AND User_Dest = '" & UD1$ & "'"
    '
    Call CREABACONLEC
    Set MIRS = BADACON.OpenRecordset(STRSQL, dbOpenDynaset)
    On Error Resume Next
    MIRS.MoveFirst
    If Err Then
       On Error GoTo 0
       UDES$ = UD1$
       GoSub AGREGAREGISTRO
    End If
    On Error GoTo 0
    '
    If UD2$ <> "" Then
      If UD2$ <> UD1$ Then
        STRSQL = "SELECT * FROM CONTRO_DOWN"
        STRSQL = STRSQL + " WHERE File_Name = '" & FINAME$ & "'"
        STRSQL = STRSQL + " AND File_DateTime = '" & FIDATI$ & "'"
        STRSQL = STRSQL + " AND User_Dest = '" & UD2$ & "'"
        '
        Set MIRS = BADACON.OpenRecordset(STRSQL, dbOpenDynaset)
        On Error Resume Next
        MIRS.MoveFirst
        If Err Then
           On Error GoTo 0
           UDES$ = UD2$
           GoSub AGREGAREGISTRO
        End If
        On Error GoTo 0
      End If
    End If
    '
    On Error Resume Next
    MIRS.Close
    BADACON.Close
    On Error GoTo 0
    '
Exit Sub
    
AGREGAREGISTRO:
    With MIRS
      .AddNew   ' YA CONVERTIDO
      !File_Name = FINAME$
      !File_DateTime = FIDATI$
      !User_Dest = UDES$
      !Num_Pack = NUCTRL$
      !FeHor_Baja = ""
      !Usu_Baja = ""
      .Update
    End With
    MIRS.Close
' MsgBox "*** Agregue Registro " + FINAME$ + " " + FIDATI$ + " " + UDES$ + " " + NUCTRL$

Return

End Sub
'
Function BAJADELAWEB(NARCHI$)
   '
   If MODOX$ = "LOCAL" Then
        JJ1& = 1
        TTXX$ = LOADFILE$("F:\WEBUPDATES\NOVEDAD\" + NARCHI$)
        If TTXX$ <> "" Then
           Call SAVEFILE(CARPEDOWN$ + NARCHI$, TTXX$)
           JJ1& = 0
        End If
      Else
        DESTIBAJA$ = CARPEDOWN$
        If InStr(FINAME$, "BKSEVOL") > 0 Then
           DESTIBAJA$ = App.Path + "\"
           ' DE ESTA FORMA LO BAJA DIRECTAMENTE A LA CARPETA LOCAL
        End If
        JJ1& = URLDownloadToFile(0, "http://www.flexoft.com.ar/UPDATES/NOVETEST/" + NARCHI$, DESTIBAJA$ + NARCHI$, 0, 0)
        If JJ1& <> 0 Then   ' OCURRIO ERROR DE BAJADA
          PPXX% = InStr(NARCHI$, ".")
          If PPXX% > 1 Then
             'INTENTA DIVERSAS COMBINACIONES DE MINUSCULA Y MAYUSCULA
             ARMIN$ = LCase$(Left$(NARCHI$, PPXX% - 1))
             ARMAX$ = UCase$(ARMIN$)
             EXMIN$ = LCase$(Mid$(NARCHI$, PPXX% + 1))
             EXMAX$ = UCase$(Mid$(NARCHI$, PPXX% + 1))
             JJ1& = URLDownloadToFile(0, "http://www.flexoft.com.ar/UPDATES/NOVETEST/" + ARMIN$ + "." + EXMIN$, DESTIBAJA$ + NARCHI$, 0, 0)
             If JJ1& = 0 Then GoTo 90
             JJ1& = URLDownloadToFile(0, "http://www.flexoft.com.ar/UPDATES/NOVETEST/" + ARMAX$ + "." + EXMIN$, DESTIBAJA$ + NARCHI$, 0, 0)
             If JJ1& = 0 Then GoTo 90
             JJ1& = URLDownloadToFile(0, "http://www.flexoft.com.ar/UPDATES/NOVETEST/" + ARMIN$ + "." + EXMAX$, DESTIBAJA$ + NARCHI$, 0, 0)
             If JJ1& = 0 Then GoTo 90
             JJ1& = URLDownloadToFile(0, "http://www.flexoft.com.ar/UPDATES/NOVETEST/" + ARMAX$ + "." + EXMAX$, DESTIBAJA$ + NARCHI$, 0, 0)
             If JJ1& = 0 Then GoTo 90
          End If
        End If
   End If
   '
90 BAJADELAWEB = JJ1&
   '
End Function
'
Sub BORRACACHE()
    '
    ' ESTO SE CORTÓ Y PEGÓ TAL CUAL DE LA PAGINA DE MICROSOFT
    ' NO QUEDA MUY CLARO COMO FUNCIONA PERO LO HACE Y NO DEBE TOCARSE NI MODI-
    ' FICARSE BAJO NINGUN CONCEPTO
    '
    Dim sGroupID As Date
    Dim hGroup As Long
    Dim hFile As Long
    Dim sEntryInfo As INTERNET_CACHE_ENTRY_INFO
    Dim iSize As Long
    '
EPAC$ = TRIM$(UCase$(EMPREAC$))
If InStr(EPAC$, "FACIL") > 0 Then Exit Sub
    '
    Screen.MousePointer = 11
    VUELTA% = 0      ' LO HACE 3 VECES
    '
10  On Error Resume Next
    ' Delete the groups
    hGroup = FindFirstUrlCacheGroup(0, 0, 0, 0, sGroupID, 0)
    ' To avoid error using it with IE4 as FindFirstUrlCacheGroup is not implemented
    If Err Then
      If Err.Number <> 453 Then
         If (hGroup = 0) And (Err.LastDllError <> 2) Then
            Screen.MousePointer = 1
            If Left$(UCase$(LUCOM$), 2) = "C:" Then
               WEBSE$ = "DISABLED"
               Call SAVEFILE(LUCOM$ + "WEB_SER.TXT", WEBSE$)
            End If
            TTXX$ = "Se ha Producido Error " & Err.LastDllError & " al Enumerar "
            TTXX$ = TTXX$ + "los Grupos de Cache. No Podrá Obtener Actualizaciones "
            TTXX$ = TTXX$ + "Automáticas FLEXOFT en esta Estación de Trabajo.\  \"
            TTXX$ = TTXX$ + "Se Recomienda Ejecutar 'LIMPIA_CACHE.EXE' en esta Terminal.\   \"
            TTXX$ = TTXX$ + "FLEXOFT ha Deshabilitado WEB Services en Esta Estacion de Trabajo."
            MsgBox TTXX$
            Exit Sub
         End If
       Else
         Err.Clear
      End If
    End If
    '
    If (hGroup <> 0) Then
        'we succeeded in finding the first cache group.. enumerate and
        'delete
        Do
            If (0 = DeleteUrlCacheGroup(sGroupID, CACHEGROUP_FLAG_FLUSHURL_ONDELETE, 0)) Then
               
               ' To avoid error using it with IE4 as FindFirstUrlCacheGroup is not implemented
               If Err.Number <> 453 Then
                 MsgBox "Error deleting cache group " & Err.LastDllError
                 Exit Sub
               Else
                  Err.Clear
               End If
            End If
            iSize = BUFFERSIZE
            If (0 = FindNextUrlCacheGroup(hGroup, sGroupID, iSize)) And (Err.LastDllError <> 2) Then
                MsgBox "Error finding next url cache group! - " & Err.LastDllError
            End If
        Loop Until Err.LastDllError = 2
    End If
  ' Delete the files
    sEntryInfo.dwStructSize = 80
    iSize = BUFFERSIZE
    hFile = FindFirstUrlCacheEntry(0, sEntryInfo, iSize)
    If (hFile = 0) Then
        If (Err.LastDllError = ERROR_NO_MORE_ITEMS) Then
            GoTo done
        End If
        If (Err.LastDllError = 122) Then
            MsgBox "Error 122 - Tamaño de Buffer " & BUFFERSIZE & " Insuficiente. Consulte."
            Exit Sub
        End If
        MsgBox "ERROR: FindFirstUrlCacheEntry - " & Err.LastDllError
        Exit Sub
    End If
    TINIX = Timer
    Do
        If (0 = DeleteUrlCacheEntry(sEntryInfo.szRestOfData(0))) _
            And (Err.LastDllError <> 2) Then
            Err.Clear
        End If
        iSize = BUFFERSIZE
        If (0 = FindNextUrlCacheEntry(hFile, sEntryInfo, iSize)) And (Err.LastDllError <> ERROR_NO_MORE_ITEMS) Then
            MsgBox "Error:  Unable to find the next cache entry - " & Err.LastDllError
            Exit Sub
        End If
    Loop Until Err.LastDllError = ERROR_NO_MORE_ITEMS
done:
    If VUELTA% < 1 Then
       VUELTA% = VUELTA% + 1
       GoTo 10
    End If
    Screen.MousePointer = 1
    '
End Sub
  
Sub MANDALAWEB(LISTADEARCHIVOS$, Optional DESTINO$)   ' RUTAS Y NOMBRES SEPARADOS POR ;
   '
   Dim FINX$()
   ReDim FINX$(0)
   '
   STRID$ = String$(10000, 0)
   DESTX$ = TRIM$(DESTINO$)
   PPXX% = InStr(DESTX$, ";")
   If PPXX% < 1 Then PPXX% = 1 + Len(DESTX$)
   EPDESTI$ = Left$(DESTX$, PPXX% - 1)
   USDESTI$ = Mid$(DESTX$, PPXX% + 1)
   PPXX% = InStr(USDESTI$, "+")
   If PPXX% < 1 Then PPXX% = 1 + Len(USDESTI$)
   CONCOPI$ = TRIM$(Mid$(USDESTI$, PPXX% + 1))
   USDESTI$ = TRIM$(Left$(USDESTI$, PPXX% - 1))
   '
   If TRIM$(EPDESTI$) = "" Or (TRIM$(USDESTI$) = "" And InStr(LISTADEARCHIVOS$, ".REL") < 1) Then
      MsgBox "Faltan Datos de Empresa / Usuario Destinatario." + vbCrLf + vbCrLf + "Esta C-Tel debe enviarse por e-mail."
      Exit Sub
   End If
   '
   ARINDEX$ = "F:\WEBUPDATES\NOVEDAD\IUPDATES.TXT"
   If EXISTE%(ARINDEX$) > 0 Then
      FULTACT$ = LOADFILE$(ARINDEX$)
      CANREGACT& = Len(FULTACT$) / 128
      JJ& = 0
      ReDim FINX$(CANREGACT&)
      For U& = 1 To CANREGACT&
         FIRECVI$ = Mid$(FULTACT$, 128 * U& - 127, 128)
         If TRIM$(Left$(FIRECVI$, 64)) <> "" Then
           XBUS$ = TRIM$(UCase$(Left$(FIRECVI$, 64)))
           FBUS& = Val(Mid$(FIRECVI$, 65, 5))
           USEP$ = TRIM$(UCase$(Mid$(FIRECVI$, 75, 28)))
           '
           ' AQUI ELIMINA DEL INDICE LOS QUE TIENEN MÁS DE 10 DÍAS
           FECHAVI = FBUS&
           FECHAHO = HOY(HOS$)
           If DIENTRE%(FECHAVI, FECHAHO) > 15 Then
              GoTo 8
           End If
           '
           For KK& = 1 To JJ&
             If TRIM$(UCase$(Left$(FINX$(KK&), 64))) = XBUS$ Then
               If TRIM$(UCase$(Mid$(FIRECVI$, 75, 28))) = USEP$ Then
                 If Val(Mid$(FINX$(KK&), 65, 5)) < FBUS& Then
                   FINX$(KK&) = FIRECVI$
                 End If
                 GoTo 8
               End If
             End If
           Next KK&
           JJ& = JJ& + 1
           FINX$(JJ&) = FIRECVI$
           '
           ' DETERMINA LOS IDENTIFICADORES OCUPADOS
           IDXX& = Val(Mid$(FIRECVI$, 103, 4))
           If IDXX& >= 0 Then
             If IDXX& <= 9999 Then
               Mid$(STRID$, IDXX& + 1, 1) = Chr$(1)
             End If
           End If
         End If
8     Next U&
   End If
   '
   ' arma el identificador del mensaje mediane el que se unifican la CTel
   ' con todos sus adjuntos. Es un numero de 0000 a 9999 justado a 4 pos
   ' en el indice IUPDATES en posiciòn 103, 4.
   PRIDLI& = InStr(STRID$, Chr$(0)) - 1
   If PRIDLI& < 0 Or PRIDLI& > 9999 Then
      MsgBox "No Hay Identificadores Libres - Consulte.", vbCritical
   End If
   IDME$ = TRIM$(Str$(PRIDLI&))
   While Len(IDME$) < 4: IDME$ = "0" + IDME$: Wend
   '
   LDAX$ = LISTADEARCHIVOS$
   If Right$(LDAX$, 1) <> ";" Then LDAX$ = LDAX$ + ";"
   Do While Len(LDAX$) > 0
      PPYY% = InStr(LDAX$, ";")
      If PPYY% < 1 Then PPYY% = 1 + Len(LDAX$)
      '
      ACOMPLE$ = UCase$(TRIM$(Left$(LDAX$, PPYY% - 1)))
      LDAX$ = Mid$(LDAX$, PPYY% + 1)
      If ACOMPLE$ <> "" Then
         TTXX$ = ACOMPLE$
10       PPXX% = InStr(TTXX$, "\")
         If PPXX% > 0 Then
            TTXX$ = Mid$(TTXX$, PPXX% + 1)
            GoTo 10
         End If
         TXTNOMARC$ = UCase$(TTXX$)
         '
         If EXISTE%(ACOMPLE$) > 0 Then
            TTXX$ = Format$(FileDateTime(ACOMPLE$), "dd/mm/yy  hh:mm")  'Se obtiene la fecha de la ultima modificacion
            TXTFECHA% = FECHNUM(Left$(TTXX$, 8))           'Se obtiene la fecha de la ultima modificacion
            TXTHORA% = 100 * Val(Mid$(TTXX$, 11, 2)) + Val(Mid$(TTXX$, 14, 2))   'Se obtiene la hora de la ultima modificacion
            '
            XX$ = Space$(128)
            'Se colocan los los datos en el registro
            Call REPLA(XX$, TXTNOMARC$, 1, 64)
              TTTT$ = TRIM$(Str$(TXTFECHA%)): While Len(TTTT$) < 5: TTTT$ = " " + TTTT$: Wend
            Call REPLA(XX$, TTTT$, 65, 5)
              TTTT$ = TRIM$(Str$(TXTHORA%)): While Len(TTTT$) < 4: TTTT$ = " " + TTTT$: Wend
            Call REPLA(XX$, TTTT$, 71, 4)
            Call REPLA(XX$, USDESTI$, 75, 12)
            Call REPLA(XX$, EPDESTI$, 87, 16)
            Call REPLA(XX$, IDME$, 103, 4)
            ' libres 107, 22
            '
            For KK& = 1 To JJ&
              If TRIM$(Left$(FINX$(KK&), 64)) = TXTNOMARC$ Then  'archivo
                If TRIM$(Mid$(FINX$(KK&), 75, 12)) = USDESTI$ Then 'usuario
                  If TRIM$(Mid$(FINX$(KK&), 87, 16)) = EPDESTI$ Then 'empresa
                    FINX$(KK&) = XX$
                    GoTo 29
                  End If
                End If
              End If
            Next KK&
            If JJ& >= UBound(FINX$) Then ReDim Preserve FINX$(JJ& + 1)
            JJ& = JJ& + 1
            FINX$(JJ&) = XX$
            '
            If CONCOPI$ <> "" Then
               Call REPLA(XX$, AJUSTI$(CONCOPI$, 12), 75, 12)
               If JJ& >= UBound(FINX$) Then ReDim Preserve FINX$(JJ& + 1)
               JJ& = JJ& + 1
               FINX$(JJ&) = XX$
            End If
            '
29          On Error Resume Next
            Kill "F:\WEBUPDATES\NOVEDAD\" + UCase$(TXTNOMARC$)
            On Error GoTo 0
            On Error Resume Next
            FileCopy ACOMPLE$, "F:\WEBUPDATES\NOVEDAD\" + TXTNOMARC$
            If Err Then
               On Error GoTo 0
               MsgBox "Error al Copiar " + ACOMPLE$ + " hacia F:\WEBUPDATES\NOVEDAD\" + TXTNOMARC$
            End If
            On Error GoTo 0
            '
         End If
      End If
   Loop
   '
   ' ESCRIBE EN EL ARCHIVO
   FULTACT$ = ""
   For KK& = 1 To JJ&
       FULTACT$ = FULTACT$ + FINX$(KK&)
   Next KK&
   Call SAVEFILE("F:\WEBUPDATES\NOVEDAD\IUPDATES.TXT", FULTACT$)
   MsgBox "Actualización Enviada a WEBUPDATES"
   '
End Sub

Function PUEDEBAJARDELAWEB%()
    '
    WEBSE$ = "DISABLED": PUEBA% = 0
    TTXX$ = UCase$(LOADFILE$("C:\FLEXOFT\WEB_SER.TXT"))
    If TTXX$ = "ENABLED" Then
       WEBSE$ = TTXX$
       PUEBA% = 1
    End If
    '
    PUEDEBAJARDELAWEB% = PUEBA%
    '
End Function
'
Function HORTEX$(XL%)
    '
    HORA% = Int(XL% / 100)
    MINU% = XL% Mod 100
    If HORA% < 0 Or HORA% > 24 Or MINU% < 0 Or MINU% > 59 Or XL% > 2400 Then
        HORA% = 0: MINU% = 0
    End If
    '
    HORAS$ = Mid$(Str$(HORA%), 2)
    While Len(HORAS$) < 2: HORAS$ = " " + HORAS$: Wend
    MINUS$ = Mid$(Str$(MINU%), 2)
    While Len(MINUS$) < 2: MINUS$ = "0" + MINUS$: Wend
    '
    HORTEX$ = HORAS$ + ":" + MINUS$
    '
End Function
'
Function FECHNUM(FETX$)
    '
    Dim FE1(3)
    AX% = 0
    FETY$ = REPLACAR$(FETX$, ".", "/")
    FETX$ = REPLACAR$(FETY$, "-", "/")
    If Mid$(FETX$, 3, 1) <> "/" Then FETX$ = Left$(FETX$, 2) + "/" + Mid$(FETX$, 3)
    If Mid$(FETX$, 6, 1) <> "/" Then FETX$ = Left$(FETX$, 5) + "/" + Mid$(FETX$, 6)
    '
    If ((Len(FETX$) <> 8) Or (Mid$(FETX$, 3, 1) <> "/") Or (Mid$(FETX$, 6, 1) <> "/")) Then
        AX% = 1: GoTo 610
    End If
    '
    For II% = 1 To 3
        FE1(II%) = Val(Mid$(FETX$, 3 * II% - 2, 2))
        If II% <> 3 Then If FE1(II%) < 1 Then GoTo 600
    Next II%
    '
    If FE1(3) < 60 Then FE1(3) = 100 + FE1(3)                 ' rutina a¤o 2000
    '
    If ((FE1(1) > 31) Or (FE1(2) > 12)) Then GoTo 600
    FE = FE1(1) + 32 * FE1(2) + 32 * 13 * (FE1(3) - 80)
    GoTo 610

600 AX% = 2
610 If AX% > 0 Then FE = 0
    If FE < (-32767) Then FE = (-32767)
    If FE > 32767 Then FE = 32767
    '
    FECHNUM = FE
    '
End Function
'
Sub BAJADIREC()
   '
   Dim JJ1 As Long
   Dim FIRECVIE$()
   Dim ARCHAC$(2048)
   '
   ' limpia la memoria cache
   Call BORRACACHE
   '
   ' baja la última version de IUPDATES.TXT
   JJ1& = BAJADELAWEB("IUPDATES.TXT")
   If JJ1& <> 0 Then
      MsgBox "Error al Bajar 'IUPDATES.TXT'"
      Exit Sub
   End If
   FIDOWNL$ = LOADFILE$(CARPEDOWN$ + "IUPDATES.TXT")
   '
   ' LEER LA TABLA DE NUEVAS ACTUALIZACIONES
   CANREGNUE& = Len(FIDOWNL$) / 128
   J2& = 0:   WEBROWSR.List3.Clear
   For U& = 1 To CANREGNUE&
      FIRECNU$ = Mid$(FIDOWNL$, 128 * U& - 127, 128)
      If TRIM$(Left$(FIRECNU$, 64)) <> "" Then
         EPLIDEST$ = TRIM$(Mid$(FIRECNU$, 87, 16))
         ' QUE EL NÚMERO DE SERIE DEL VOLUMEN COINCIDA O ESTÉ EN BLANCO
         If EPLIDEST$ <> "" Then
            If EPLIDEST$ <> IDEPLI$ Then
              GoTo 29    ' NUMERO DE SERIE DIFIERE
            End If
         End If
         '
         FINAME$ = TRIM$(Left$(FIRECNU$, 64))
         FEHO$ = TRIM$(Mid$(FIRECNU$, 65, 10))
         FIDATI$ = FETEX$(Val(Mid$(FEHO$, 1, 5))) + " " + HORTEX$(Val(Mid$(FEHO$, 6, 5)))
         UDEST1$ = TRIM$(Mid$(FIRECNU$, 75, 12))
         UDEST2$ = TRIM$(Mid$(FIRECNU$, 107, 12))
         NUCTRL$ = TRIM$(Mid$(FIRECNU$, 103, 4))
         
         TTZZ$ = AJUSTI$(AJUSTI$(FINAME$, 32) + FIDATI$, 48)
         TTZZ$ = TTZZ$ + AJUSTI$(UDEST1$, 12)
         TTZZ$ = TTZZ$ + AJUSTI$(UDEST2$, 12)
         TTZZ$ = TTZZ$ + AJUSTI$(NUCTRL$, 4)
         ABAFORZA.List3.AddItem TTZZ$
      End If
29 Next U&
   '
   ABAFORZA.Show 1
   '
99 Screen.MousePointer = 1
   '
End Sub
