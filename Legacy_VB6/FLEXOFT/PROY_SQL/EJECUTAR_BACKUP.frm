VERSION 5.00
Begin VB.Form EJECUTAR_BACKUP 
   Caption         =   "Form1"
   ClientHeight    =   345
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   2025
   LinkTopic       =   "Form1"
   ScaleHeight     =   345
   ScaleWidth      =   2025
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "EJECUTAR_BACKUP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim FLEXCONN As ADODB.Connection
Dim CarpetaDefault As String

Private Sub Form_Load()
    CarpetaDefault = "C:\FLEXOFT"
    Call ABRECONEXION
    MsgBox "Inicio de Backup"
    
    FECHA$ = CStr(Now)
    BASEDA$ = "BackupFlexoft_" & Format(Now, "dd-mm-yyyy")
    BASEDA$ = Replace(BASEDA$, "-", "")
    RUTABK$ = "C:\FLEXOFT\BACKUPS"
    tmp_RUTABK$ = RUTABK$ & "\" & BASEDA$ & ".BAK"
    
    BASEDA$ = "TODOBUS"
    
    Screen.MousePointer = 11
    sqlx$ = "BACKUP DATABASE " & BASEDA$ & " TO DISK='" & tmp_RUTABK$ & "' WITH INIT"
    FLEXCONN.Execute sqlx$
    FLEXCONN.Close
    Set FLEXCONN = Nothing
    Screen.MousePointer = 1
    
    MsgBox "Finalizó el Backup"
    
    End
    Close
End Sub

Sub ABRECONEXION()
   '
   On Error Resume Next
   ESTADODECONEXION% = FLEXCONN.State
   If ESTADODECONEXION% = 1 Then
      Exit Sub
   End If
   '
   CONNSQL$ = Trim$(LOADFILE$(CarpetaDefault + "\CONNSQLS.DRV"))
'MsgBox CONNSQL$
   If CONNSQL$ <> "" Then
      ' VERIFICA LA CONEXION
      Set FLEXCONN = New ADODB.Connection
      FLEXCONN.ConnectionString = CONNSQL$   ' "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=True;Data Source=FLEXSERVER;Initial Catalog=" + conecta$ + "; Connect Timeout=999"
      On Error Resume Next
      FLEXCONN.Open
      If Err Then
         MsgBox "Imposible Abrir Conexión a Servidor SQL con " + CONNSQL$, vbCritical
         Close: End
      End If
      On Error GoTo 0
      '
      FLEXCONN.Execute ("SET DATEFORMAT dmy")
      FLEXCONN.CommandTimeout = 240 '120
      TIBASE$ = "SQL"
      TXT$ = CONNSQL$
      PINI% = InStr(UCase$(TXT$), "DATA SOURCE=") + 12
      TXT1$ = Mid$(TXT$, PINI%)
      PFIN% = InStr(TXT1$, ";")
      SERNA$ = Left$(TXT1$, PFIN% - 1)
      PINI% = InStr(UCase$(TXT$), "INITIAL CATALOG=") + 16
      TXT1$ = Mid$(TXT$, PINI%)
      PFIN% = InStr(TXT1$, ";")
      BINIC$ = Left$(TXT1$, PFIN% - 1)
      '
      BASEACTIVA$ = BINIC$
      TIBASE$ = "SQL"
      GoTo 50
   Else
      MsgBox "Archivo de Ruta de Conexión a Servidor Sql " & vbCrLf & "(CONNSQL.DRV) Inexistente o Vacio ", vbCritical
      Close: End
   End If
   '
   On Error Resume Next
   XX$ = CONN_SQLS.Label6.Caption
   If Err Then
      On Error GoTo 0
      MsgBox "No se han Establecido los Parámetros\de Conexión al Servidor SQL-Server."
   End If
   On Error GoTo 0
   CONN_SQLS.Show 1
   MsgBox "Re-Inicie Ahora la Aplicación."
   '
50
'   '
End Sub

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
99  ERNU$ = Trim$(Str$(Err.Number))
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
    c$ = (MENSA$)
    RESPU% = MsgBox(c$, 53, "Error de Apertura de Archivo")
    If RESPU% = 4 Then
        Resume
    End If
    Close: End
    '
End Function

Function REPLACAR$(A$, b$, c$)
    A1$ = A$
    '
    If Len(b$) > 0 Then
10        XL& = InStr(A1$, b$)   ' Reemplazado XL% por XL& para evitar desbordamiento - EPB 16/01/11
          If XL& > 0 Then
              A1$ = Left$(A1$, XL& - 1) + c$ + Mid$(A1$, XL& + Len(b$))
              GoTo 10
          End If
    End If
    '
    REPLACAR$ = A1$
    A1$ = ""
    '
End Function

