Attribute VB_Name = "ABREBASE_SQL"
Public FLEXCONN As ADODB.Connection
Public FLEXCONN_DESTINO As ADODB.Connection
Public BASEACTIVA$
Dim BASABIER$
'
Public Contable As Database
Public Tabejcon As Recordset
Public Plancuen As Recordset
Public Encasien As Recordset
Public Detasien As Recordset
Public Balasec As Recordset
Public TaProyec As Recordset '\\\OT-06-0244-WAR-15/06/06///
'
Public ARTICULOS As Database
Public Master As Recordset
Public Formulas As Recordset
Public Secoper As Recordset
Public TmusOper As Recordset '\\\OT-06-0252-22/06/06///
'
Public CueCorri As Database
Public CajaBanc As Recordset
Public Sacrepen As Recordset
Public SaDeuPen As Recordset
Public Aplipago As Recordset
Public Aplicobra As Recordset
Public EncMovCc As Recordset 'nueva tabla p/imporclieCtaCte
Public DetMovCc As Recordset 'nueva tabla p/imporclieCtaCte
'
Public Stocks As Database
Public Movisto As Recordset
Public SalStock As Recordset
Public Reserva As Recordset
Public Bolrecep As Recordset
Public MaConsig As Recordset
'
Public Ofabric As Database
Public OrdeFabr As Recordset
Public OperFab As Recordset
Public InfoFab As Recordset  'nueva tabla
Public InfoProy As Recordset '\\\OT-06-0252-WAR-26/06/06///

Public Compras As Database
Public OCOMENCA As Recordset
Public OCOMDETA As Recordset
Public EncoCabi As Recordset '\\\OT-06-0149-WAR-19/04/06///
Public DetoCabi As Recordset '\\\OT-06-0149-WAR-19/04/06///
Public AutoEnt As Recordset '\\\OT-06-0149-WAR-19/04/06///
'
Public Ventas As Database
Public PresuVen As Recordset
Public PedEnca As Recordset
Public PedDeta As Recordset
'
Public Controcen As Database
Public BloqArchA As Recordset
Public BloqArchB As Recordset
'
Public STOCKORIGI As Database
Public EncaStock As Recordset
Public DetaStock As Recordset
'
Public OTrabajo As Database
Public OrdeTrab As Recordset
Public PlanCob As Recordset
Public GrupOtra As Recordset
Public ListaMat As Recordset
'
Public StockVie As Database   'MOVIMIENTOS STOCK ANTIGUO.
Public Movisto1 As Recordset  '1 XQ MOVISTO RECORDSET YA EXISTE.
'
Public PresuCam As Database
Public EncaPresu As Recordset
Public DetaPresu As Recordset
Public Despecial As Recordset
Public DOCADJUN As Recordset
'
'OT-05-0856-WAR-28/11/05///
Public Maquinarias As Database
Public Equipos As Recordset
Public GaleImag As Recordset
Public RepuMaq As Recordset
Public Contadores As Recordset
Public SubConjuntos As Recordset
'
Public Omantin As Database
Public OrdeMant As Recordset
Public RecuProp As Recordset
'
Public ACTICARGA%
Dim HUBOERROR%
'
Sub EJEMPLOPARACREARINDICES()
   Call APERBASDAT("CABAN")
   AA1 = CueCorri.TableDefs("CAJABANC").Indexes.Count
   For Ab1 = 1 To 13
   MsgBox CueCorri.TableDefs("CAJABANC").Indexes(Ab1).Name
   Next Ab1
End Sub
'
Sub APERBASDAT(NBASE$)
   '
   RUTIACT$ = "ABREBASDAT"
   'On Error GoTo 90    ' BUSCA CAPTURAR ERROR 91
   '
   NBAS$ = UCase$(NBASE$)
   Select Case NBAS$
     Case "MASTER"
        Call ABREMASTER
     Case "STRUC"
        Call ABRESTRUC
     Case "CABAN"
        Call ABRECABAN
     Case "STOCKS"
        Call ABRESTOCKS
'     Case "STOCKVIE"
'        Call ABRESTOCKVIE
     Case "COMPRAS"
        Call ABRECOMPRAS
     Case "PLANCUEN"
        Call ABREPLANCUEN
     Case "RESERVA"
        Call ABRESERVA
     Case "ORFAB"
        Call ABREORFAB
     Case "PEDCLI"
        Call ABREPEDCLI
     Case "CLIEN"
        Call ABRECLIEN
     Case "PROVE"
        Call ABREPROVE
     Case Else
        MsgBox "Imposible Abrir B.Datos '" & NBAS$ & "'", vbExclamation
        Call FINAL("")
   End Select
   '
   RUTIACT$ = ""
   '
Exit Sub
'*******************************************************************************
90 TTXX$ = "Se ha Producido un Error no Recuperable (" & Err.Number & ") en la Rutina " & RUTIACT$ & " al Abrir " & NBAS$
   TTXX$ = TTXX$ + ".\Capture o Imprima esta Pantalla, Enviándola por Mail a asistec@flexoft.com.ar."
   TTXX$ = TTXX$ + "\   \Será Respondido a la Brevedad. Muchas Gracias por su Colaboración."
   Resume 91
   '
91 On Error GoTo 0
   RUTIACT$ = ""
   MsgBox FORMESS$(TTXX$), vbExclamation
   Call FINAL("")
End Sub
'
Sub ABREMASTER()
   '
   If ESTABIER%("ARTICULOS") < 1 Then
      Call MENSERR(24, "Imposible Abrir Base de Articulos.")
      Call FINAL("")
   End If
   '
End Sub
'
Sub ABRESTRUC()
   '
   If ESTABIER%("ARTICULOS") < 1 Then
      Call MENSERR(24, "Imposible Abrir Base de Articulos.")
      Call FINAL("")
   End If
   '
End Sub
'
Sub ABRESECOP()
   '
   If ESTABIER%("ARTICULOS") < 1 Then
      Call MENSERR(24, "Imposible Abrir Base de Articulos.")
      Call FINAL("")
   End If
   '
End Sub
'
Sub ABRECLIEN()
   '
   If ESTABIER%("BADACLIEN") < 1 Then
      Call MENSERR(24, "Imposible Abrir Base de Articulos.")
      Call FINAL("")
   End If
   '
End Sub

Sub ABRECONEXION()
   '
   On Error Resume Next
   ESTADODECONEXION% = FLEXCONN.State
   If ESTADODECONEXION% = 1 Then
      Exit Sub
   End If
   '
   CONNSQL$ = TRIM$(LOADFILE$(LUDAT$ + "CONNSQLS.DRV"))
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
      Call MENSERR(24, "No se han Establecido los Parámetros\de Conexión al Servidor SQL-Server.")
      Call FINAL("")
   End If
   On Error GoTo 0
   CONN_SQLS.Show 1
   Call COMUNI("Re-Inicie Ahora la Aplicación.")
   Call FINAL("")
   '
   '
50 ' verifica que exista el campo CodiEmpr en todas las tablas - si no lo crea
'   strsql = "SELECT SYSOBJECTS.NAME FROM SYSOBJECTS "
'   Dim MIRS As ADODB.Recordset
'   Set MIRS = New ADODB.Recordset
'   Set MIRS = FLEXCONN.Execute(strsql)
'   With MIRS
'      Do While Not .EOF
'         NTABLA$ = !Name
'         If UCase$(Left$(NTABLA$, 1)) = Left$(NTABLA$, 1) Then
'           If Left$(NTABLA$, 3) <> "DF_" Then
'             Call CREACAMPO("", NTABLA$, "CodiEmpr", 3, 2, 1)
'             QTR& = CantRegistros(NTABLA$, "CodiEmpr IS NULL")
'             If QTR& > 0 Then
'                SQLX$ = "UPDATE " + NTABLA$ + " SET CodiEmpr = '" & CodiEmp% & "' WHERE CodiEmpr IS NULL"
'                On Error Resume Next
'                 FLEXCONN.Execute (SQLX$)
'                If Err Then
'                   On Error GoTo 0
'                   Call PONE_CODIEMPR_EN_CERO(NTABLA$)
'                End If
'                On Error GoTo 0
'             End If
'           End If
'         End If
'      .MoveNext
'      Loop
'   End With
'   MIRS.Close
'   Set MIRS = Nothing
'   '
End Sub
'
Sub ABRECONEXION_DESTINO(BASE_DESTI$)
   '
   If TRIM$(BASE_DESTI$) = "" Then Exit Sub
   On Error Resume Next
   ESTADODECONEXION2% = FLEXCONN_DESTINO.State
   If ESTADODECONEXION2% = 1 Then
      Exit Sub
   End If
   '

   'LEE EL STRING DE CONEXION ORIGINAL
   CONNSQL$ = TRIM$(LOADFILE$(LUDAT$ + "CONNSQLS.DRV"))
   PFIN% = InStrRev(CONNSQL$, ";")
   PINI% = InStrRev(CONNSQL$, "=")
   TXXT$ = Left$(CONNSQL$, PINI% - 1)
   PINI% = InStrRev(TXXT$, "=")
   BINIC$ = Mid$(CONNSQL$, PINI% + 1, (PFIN%) - (PINI% + 1))
   '
   'REEMPLAZA EL STRING DE CONEXION CON LA NUEVA BASE DE DATOS DE DESTINO
   CONNSQL_DESTINO$ = REPLACAR(TRIM$(UCase$(CONNSQL$)), TRIM$(UCase$(BINIC$)), TRIM$(UCase$(BASE_DESTI$)))
   '
   'GENERA EL ARCHIVO DE CONEXION DESTINO
   NARCHI2% = FILEOPEN%(LUDAT$ + "CONNSQLS_DESTINO.DRV", 2, 256) ' ABRE EL ARCHIVO DE ESCRITURA
   Print #NARCHI2%, CONNSQL_DESTINO$ ' ESCRIBE EN EL ARCHIVO EL NUEVO STRING DE CONEXION
   Close #NARCHI2%
   '
'MsgBox CONNSQL$
   If CONNSQL_DESTINO$ <> "" Then
      ' VERIFICA LA CONEXION
      Set FLEXCONN_DESTINO = New ADODB.Connection
      FLEXCONN_DESTINO.ConnectionString = CONNSQL_DESTINO$   ' "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=True;Data Source=FLEXSERVER;Initial Catalog=" + conecta$ + "; Connect Timeout=999"
      On Error Resume Next
      FLEXCONN_DESTINO.Open
      If Err Then
         MsgBox "Imposible Abrir Conexión a Servidor SQL con " + CONNSQL_DESTINO$, vbCritical
         Close: End
      End If
      On Error GoTo 0
      '
      FLEXCONN_DESTINO.Execute ("SET DATEFORMAT dmy")
      FLEXCONN_DESTINO.CommandTimeout = 240 '120
      TIBASE$ = "SQL"
      TXT$ = CONNSQL_DESTINO$
      PINI% = InStr(UCase$(TXT$), "DATA SOURCE=") + 12
      TXT1$ = Mid$(TXT$, PINI%)
      PFIN% = InStr(TXT1$, ";")
      SERNA$ = Left$(TXT1$, PFIN% - 1)
      PINI% = InStr(UCase$(TXT$), "INITIAL CATALOG=") + 16
      TXT1$ = Mid$(TXT$, PINI%)
      PFIN% = InStr(TXT1$, ";")
      BINIC$ = Left$(TXT1$, PFIN% - 1)
      '
      BASEACTIVA_DESTI$ = BINIC$
      TIBASE$ = "SQL"
      MsgBox "STRING DE CONEXION: " & vbCrLf & CONNSQL_DESTINO$
      
''''      GoTo 50
   Else
      MsgBox "Archivo de Ruta de Conexión a Servidor Sql " & vbCrLf & "(CONNSQLS_DESTINO.DRV) Inexistente o Vacio ", vbCritical
''''      Close: End
   End If
   '
''''   On Error Resume Next
''''   XX$ = CONN_SQLS.Label6.Caption
''''   If Err Then
''''      On Error GoTo 0
''''      Call MENSERR(24, "No se han Establecido los Parámetros\de Conexión al Servidor SQL-Server.")
''''      Call FINAL("")
''''   End If
''''   On Error GoTo 0
'''''   CONN_SQLS.Show 1
''''   Call COMUNI("Re-Inicie Ahora la Aplicación.")
''''   Call FINAL("")
   '
   '
50 ' verifica que exista el campo CodiEmpr en todas las tablas - si no lo crea
'   strsql = "SELECT SYSOBJECTS.NAME FROM SYSOBJECTS "
'   Dim MIRS As ADODB.Recordset
'   Set MIRS = New ADODB.Recordset
'   Set MIRS = FLEXCONN_DESTINO.Execute(strsql)
'   With MIRS
'      Do While Not .EOF
'         NTABLA$ = !Name
'         If UCase$(Left$(NTABLA$, 1)) = Left$(NTABLA$, 1) Then
'           If Left$(NTABLA$, 3) <> "DF_" Then
'             Call CREACAMPO("", NTABLA$, "CodiEmpr", 3, 2, 1)
'             QTR& = CantRegistros(NTABLA$, "CodiEmpr IS NULL")
'             If QTR& > 0 Then
'                SQLX$ = "UPDATE " + NTABLA$ + " SET CodiEmpr = '" & CodiEmp% & "' WHERE CodiEmpr IS NULL"
'                On Error Resume Next
'                 FLEXCONN_DESTINO.Execute (SQLX$)
'                If Err Then
'                   On Error GoTo 0
'                   Call PONE_CODIEMPR_EN_CERO(NTABLA$)
'                End If
'                On Error GoTo 0
'             End If
'           End If
'         End If
'      .MoveNext
'      Loop
'   End With
'   MIRS.Close
'   Set MIRS = Nothing
'   '
End Sub
Sub PONE_CODIEMPR_EN_CERO(NTABLA$)
   '
   Call COMUNI("FLEXOFT Debe Ahora Actualizar Campo 'CodiEmpr'\en la Tabla '" + NTABLA$ + "'.\   \Esto Puede Demandar Varios Minutos.\   \Pulse 'Aceptar Mensaje' para Comenzar.")
   Screen.MousePointer = 11
   '
   CONDI$ = "CodiEmpr IS NULL"
   FIN& = CantRegistros(NTABLA$, CONDI$)
   STRSQL = "Select * from " + NTABLA$ + " WHERE " + CONDI$
   Dim RTABLA As ADODB.Recordset
   Set RTABLA = New ADODB.Recordset
25 On Error Resume Next
   RTABLA.Open (STRSQL), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText

   If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
      On Error GoTo 0
      Call CapturaErrorOpen
      GoTo 25
   End If
   On Error GoTo 0
   '
   U& = 0
   With RTABLA
      Do While Not .EOF
         U& = U& + 1
         Call TRACE(20, U&, FIN&)
         !CodiEmpr = CodiEmp%
         .Update
      .MoveNext
      Loop
   End With
   '
   RTABLA.Close
   Set RTABLA = Nothing
   Screen.MousePointer = 1
   Call COMUNI("Proceso Completo")
   '
End Sub
'
Sub ABRECONTROL()
   '
   If ESTABIER%("BCONTROL") < 1 Then
      Call MENSERR(24, "Imposible Abrir Base de Control.")
      Call FINAL("")
   End If
   '
End Sub
'
Sub ABREPROVE()
   '
   If ESTABIER%("BADAPROVE") < 1 Then
      Call MENSERR(24, "Imposible Abrir Base de Proveedores.")
      Call FINAL("")
   End If
   '
End Sub
'
Sub ABRECABAN()
   '
   If ESTABIER%("CUECORRI") < 1 Then
      Call MENSERR(24, "Imposible Abrir Base de Cuenta Corriente.")
      Call FINAL("")
   End If
   '
End Sub
'
Sub ABREPLANCUEN()
   '
   If ESTABIER%("CONTAGEN") < 1 Then
      Call MENSERR(24, "Imposible Abrir Base de Contabilidad.")
      Call FINAL("")
   End If
   '
End Sub

Sub ABRECOMPRAS()
   '
   If ESTABIER%("COMPRAS") < 1 Then
      Call MENSERR(24, "Imposible Abrir Base de Compras.")
      Call FINAL("")
   End If
   '
End Sub

Sub ABREPRESVEN()
   '
   If ESTABIER%("VENTAS") < 1 Then
      Call MENSERR(24, "Imposible Abrir Base de Ventas.")
      Call FINAL("")
   End If
   '
End Sub

Sub ABREPEDCLI()
   '
   If ESTABIER%("VENTAS") < 1 Then
      Call MENSERR(24, "Imposible Abrir Base de Ventas.")
      Call FINAL("")
   End If
   '
End Sub

Sub ABRESTOCKS()
   '
   If ESTABIER%("STOCKS") < 1 Then
      Call MENSERR(24, "Imposible Abrir Base de Stocks.")
      Call FINAL("")
   End If
   '
End Sub
'
Sub ABRESTOCKVIE()
   '
   If ESTABIER%("STOCKVIE") < 1 Then
      Call MENSERR(24, "Imposible Abrir Base de Stocks - Viejos.")
      Call FINAL("")
   End If
   '
End Sub

Sub ABRESERVA()
   '
   If ESTABIER%("STOCKS") < 1 Then
      Call MENSERR(24, "Imposible Abrir Base de Reservas.")
      Call FINAL("")
   End If
   '
End Sub
'
Sub ABREORFAB()
   '
   If ESTABIER%("OFABRIC") < 1 Then
      Call MENSERR(24, "Imposible Abrir Base de Ordenes de Fabricacion.")
      Call FINAL("")
   End If
   '
End Sub
'
Function ESTABIER%(NOMBASE$)
   '
   Static BASEANT$
   '
   If BASEACTIVA$ <> BASEANT$ Then
      BASABIER$ = ""   ' SI CAMBIA LA BASE ACTIVA DEVUELVE CERO
      BASEANT$ = BASEACTIVA$
   End If
   '
   If InStr(BASABIER$, UCase$(TRIM$(NOMBASE$) + ",")) > 0 Then
      ESTABIER% = 1
      Exit Function
   End If
   '
   SCMP = Screen.MousePointer
   Screen.MousePointer = 11
   Call ABRECONEXION
   '
   HUBOERROR% = 0
   Call GENERBASE(NOMBASE$)
   BASABIER$ = BASABIER$ + TRIM$(UCase$(NOMBASE$)) + ", "
   ESTABIER% = 1
   Screen.MousePointer = SCMP
   '
End Function

Sub ABREMOVSTK()
   '
   If ESTABIER%("STOCKORIGI") < 1 Then
      Call MENSERR(24, "Imposible Abrir Base de Movimientos de Stock.")
      Call FINAL("")
   End If
   '
End Sub

Sub ABREORTRA()
   '
   If ESTABIER%("OTRABAJO") < 1 Then
      Call MENSERR(24, "Imposible Abrir Base de Ordenes de Trabajo.")
      Call FINAL("")
   End If
   '
End Sub
'
Sub ABREPRECAM()
   '
   If ESTABIER%("PRESUCAM") < 1 Then
      Call MENSERR(24, "Imposible Abrir Base de Presupuestos de Cámara.")
      Call FINAL("")
   End If
   '
End Sub
'
Sub ABREMAQUI()
   '
   If ESTABIER%("MAQUINARIAS") < 1 Then
      Call MENSERR(24, "Imposible Abrir Base de Maquinarias.")
      Call FINAL("")
   End If
   '
End Sub

Sub ABREOMAN()
   '
   If ESTABIER%("OMANTIN") < 1 Then
      Call MENSERR(24, "Imposible Abrir Base de Ordenes de Mantenimiento.")
      Call FINAL("")
   End If
   '
End Sub

Sub APAREDAT(BADAT$, LUREP$, LUACC$, YAPAR%)
   YAPAR% = 1
   PPXX1% = InStr(UCase$(LUACC$), "\BDACCESS")
   PPXX2% = InStr(UCase$(LUREP$), "\BDACCESS")
   If PPXX1% > 1 Then
     If PPXX2% > 1 Then
       LDX1$ = Left$(LUACC$, PPXX1%)
       LDX2$ = Left$(LUREP$, PPXX2%)
       If LDX1$ <> LDX2$ Then
         BDX$ = TRIM$(UCase$(BADAT$))
         Select Case BDX$
            Case "ARTICULOS"
               ARAP$ = "MASTER": GoSub 60
               ARAP$ = "ECOMAST": GoSub 60
         End Select
       End If
     End If
   End If
   Exit Sub
   '
60 FDT1# = FILEDATE#(LDX1$ + ARAP$ + ".DAT")
   FDT2# = FILEDATE#(LDX2$ + ARAP$ + ".DAT")
   If Abs(FDT1# - FDT2#) > 0.00001 Then
      If EXISTE%(LDX2$ + ARAP$ + ".DAT") Then
         YAPAR% = 0
         Call CIERRARCH(ARAP$)
         On Error Resume Next
         Call FileCopy(LDX2$ + ARAP$ + ".DAT", LDX1$ + ARAP$ + ".DAT")
         On Error GoTo 0
      End If
   End If
Return
'
End Sub
'
Sub CreaTabla(NTABLA$, NCAMPO$, TCAMPO, LOCAMPO)
   '
   If BASEACTIVA$ = "" Then
      MsgBox "Falta Abrir Conexion y Base Inicial."
      Call FINAL("")
   End If
   '
10 FLEXCONN.Execute ("USE " + BASEACTIVA$)
   STRSQL = "If NOT EXISTS (Select name from sysobjects where name= '" & NTABLA$ & "')"
   STRSQL = STRSQL & " CREATE TABLE " & NTABLA$ & "("
   STRSQL = STRSQL & "[CodiEmpr] [smallint] NULL ,"
   STRSQL = STRSQL & "[CodiSucu] [smallint] NULL ,"
   STRSQL = STRSQL & "[NuClien] [smallint] NULL ,"
   STRSQL = STRSQL & "[CodiClie] [nvarchar] (16) COLLATE Modern_Spanish_CI_AS NULL ,"
   STRSQL = STRSQL & "[RasoClie] [nvarchar] (48) COLLATE Modern_Spanish_CI_AS NULL ,"
   STRSQL = STRSQL & "[FechEmisi] [smalldatetime] NULL ,"
   STRSQL = STRSQL & "[CoDocAp_Lar] [nvarchar] (18) COLLATE Modern_Spanish_CI_AS NULL ,"
   STRSQL = STRSQL & "[CoDocAp_Cor] [nvarchar] (12) COLLATE Modern_Spanish_CI_AS NULL ,"
   STRSQL = STRSQL & "[ImpoApli] [float] NULL ,"
   STRSQL = STRSQL & "[RefeApli] [nvarchar] (32) COLLATE Modern_Spanish_CI_AS NULL ,"
   STRSQL = STRSQL & "[CoDocOr_Lar] [nvarchar] (18) COLLATE Modern_Spanish_CI_AS NULL ,"
   STRSQL = STRSQL & "[CoDocOr_Cor] [nvarchar] (12) COLLATE Modern_Spanish_CI_AS NULL ,"
   STRSQL = STRSQL & "[ItotOrig] [float] NULL ,"
   STRSQL = STRSQL & "[IIvaOrig] [float] NULL ,"
   STRSQL = STRSQL & "[Mone_Apli] [smallint] NULL ,"
   STRSQL = STRSQL & "[ItotOri_Dol] [float] NULL ,"
   STRSQL = STRSQL & "[ImpoApli_Dol] [float] NULL"
   STRSQL = STRSQL & ") ON [PRIMARY]"
   
   FLEXCONN.Execute (STRSQL)
   
   'EJEMPLO DE AGREGAR COLUMNAS
   STRSQL = "IF EXISTS (Select name from sysobjects where name= '" & NTABLA$ & "')"
   STRSQL = STRSQL & " ALTER TABLE " & NTABLA$ & " ADD column_A INT NULL"
   STRSQL = STRSQL & " , column_B VARCHAR NULL"
   FLEXCONN.Execute (STRSQL)
   
   'EJEMPLO DE BORRAR COLUMNAS
   STRSQL = "IF EXISTS (Select name from sysobjects where name= '" & NTABLA$ & "')"
   STRSQL = STRSQL & " ALTER TABLE " & NTABLA$ & " DROP COLUMN column_b"
   
   'EJEMPLO DE MODIFICAR COLUMNAS
   STRSQL = "IF EXISTS (Select name from sysobjects where name= '" & NTABLA$ & "')"
   STRSQL = STRSQL & " ALTER TABLE " & NTABLA$ & " ALTER COLUMN column_A smalldatetime NULL DEFAULT getdate() WITH VALUES"
   '
   'EJEMPLO DE CREAR UNA BASE DE DATOS - pls JANDY VERIFICAR
   STRSQL = "IF EXISTS (Select name from sysdatabases where name ='" + BASEACTIVA$ + "')"
   STRSQL = STRSQL & " create database '" + BASEACTIVA$ + "'"
   '
90 End Sub

Sub EJECUTACOMANDO(STRQ)
    '
    On Error Resume Next
    FLEXCONN.Execute FILTSQL$(STRQ)
    If Err Then
      Call MENSERR(24, "Error Fatal " & Err.Number & " (" & Err.Description & ")\en EJECUTACOMANDO con Comando '" & FILTSQL$(STRQ))
      On Error GoTo 0
      Call FINAL("")
    End If
    On Error GoTo 0

End Sub

Function FechaCentral() As String
    '
    Dim MIRS As ADODB.Recordset
    STRSQL = "select right('0'+cast(datepart(dd,getdate()) as varchar(2)),2)+'/'"
    STRSQL = STRSQL & "+ right('0'+cast(datepart(mm,getdate()) as varchar(2)),2)+'/'"
    STRSQL = STRSQL & "+ right('0'+cast(datepart(yy,getdate()) as varchar(4)),2) as MiFecha"
    '
    Set MIRS = FLEXCONN.Execute(STRSQL)
    If Not (MIRS.EOF And MIRS.BOF) Then
         FechaCentral = MIRS!MiFecha
       Else
         HOII% = HOY(HOS$)
         FechaCentral = HOS$
    End If
    '
    MIRS.Close
    Set MIRS = Nothing
    '
End Function
'
Function HoraCentral() As String
    '
    Dim MIRS As ADODB.Recordset
    STRSQL = "select right('0'+cast(datepart(hh,getdate()) as varchar(2)),2)+':'"
    STRSQL = STRSQL & "+ right('0'+cast(datepart(mi,getdate()) as varchar(2)),2)+':'"
    STRSQL = STRSQL & "+ right('0'+cast(datepart(ss,getdate()) as varchar(4)),2) as MiHora"
    '
    Set MIRS = FLEXCONN.Execute(STRSQL)
    If Not (MIRS.EOF And MIRS.BOF) Then
         HoraCentral = MIRS!MiHora
       Else
         HoraCentral = Time$
    End If
    '
    MIRS.Close
    Set MIRS = Nothing
    '
End Function

Function EXISTELCAMPO%(NBASE$, NTABLA$, NCAMPO$, TCAMPO, LOCAMPO, Optional CONINDEX%)
    '
   ' ESTO SE AGREGO PARA NO CREAR LOS CAMPOS YA CREADOS
   Static BASACTI$, CANCAM%, VECTAB$(), IDTAB(), VECCAM$(), VECTIP%(), VECLON%(), VECIDX%()
   NTAB$ = UCase$(NTABLA$)
   NCAM$ = UCase$(NCAMPO$)
   
   If BASACTI$ <> FLEXCONN.DefaultDatabase Then
      '
      BASACTI$ = FLEXCONN.DefaultDatabase
      CANCAM% = 0: NUEDIM% = 64
      ReDim Preserve VECTAB$(NUEDIM%), IDTAB(NUEDIM%), VECCAM$(NUEDIM%), VECTIP%(NUEDIM%), VECLON%(NUEDIM%), VECIDX%(NUEDIM%)
      STRSQL = "SELECT SYSOBJECTS.NAME AS NOTAB,SYSOBJECTS.ID AS IDENTA, SYSCOLUMNS.NAME AS NOCAM, "
      STRSQL = STRSQL + " SYSCOLUMNS.XTYPE AS TYPE, SYSCOLUMNS.LENGTH AS LENG "
      STRSQL = STRSQL + " FROM SYSCOLUMNS WITH(NOLOCK) "
      STRSQL = STRSQL + " INNER JOIN sysobjects "
      STRSQL = STRSQL & " ON sysobjects.id = syscolumns.id "
      STRSQL = STRSQL + " WHERE Left(SYSOBJECTS.NAME, 3) <> 'sys' "
      STRSQL = STRSQL + " AND LEFT(SYSOBJECTS.NAME,3) <> 'DF_' "
      '
      Set MIRS = FLEXCONN.Execute(STRSQL)
      With MIRS
         Do While Not .EOF
           If Left$(!NOTAB, 1) = UCase$(Left$(!NOTAB, 1)) Then
             CANCAM% = CANCAM% + 1
             If CANCAM% >= UBound(VECTAB$) Then
                NUEDIM% = UBound(VECTAB$) + 64
                ReDim Preserve VECTAB$(NUEDIM%), IDTAB(NUEDIM%), VECCAM$(NUEDIM%), VECTIP%(NUEDIM%), VECLON%(NUEDIM%), VECIDX%(NUEDIM%)
             End If
             IDTAB(CANCAM%) = !IDENTA
             VECTAB$(CANCAM%) = UCase$(!NOTAB)
             VECCAM$(CANCAM%) = UCase$(!NOCAM)
             VECTIP(CANCAM%) = !Type      ' TICABA
             VECLON(CANCAM%) = !LENG      ' LOCABA
           End If
         .MoveNext
         Loop
      End With
      MIRS.Close
      '
      STRSQL = "Select * from sysindexes "
      ' STRSQL = "Select * from sysindexes WHERE id = OBJECT_ID('" & NTABLA$ & "') AND NAME = 'I" + NCAMPO$ + "' "
      Set MIRS = FLEXCONN.Execute(STRSQL)
      With MIRS
         Do While Not .EOF
            ITA = !ID
            NOCA$ = UCase$(Mid$(SAFETEXT$(!Name), 2))
            For KKI% = 1 To CANCAM%
              If IDTAB(KKI%) = ITA Then
                If VECCAM$(KKI%) = NOCA$ Then
                  VECIDX%(KKI%) = 1
                  GoTo 15
                End If
              End If
            Next KKI%
15       .MoveNext
         Loop
      End With
      '
   End If
   '
   '
   If TCAMPO = 11 Then TCAMPO = 12    ' en ambos casos es secuencia ascii
   Select Case TCAMPO
      Case 1: TSQL = 104
      Case 2: TSQL = 173
      Case 3: TSQL = 52
      Case 4: TSQL = 56
      Case 5: TSQL = 60
      Case 6: TSQL = 62
      Case 7: TSQL = 62
      Case 8: TSQL = 58
      Case 10: TSQL = 231
      Case 12: TSQL = 99 ' ojo como trabajarlo
   End Select

19 EXISTELCAMPO% = 0
   For KKI% = 1 To CANCAM%
     If VECTAB$(KKI%) = NTAB$ Then
       If VECCAM$(KKI%) = NCAM$ Then
         If VECIDX%(KKI%) = CONINDEX% Or CONINDEX% < 1 Then
            EXISTELCAMPO% = 3
            If TSQL < 1 Then
               TCAMPO = VECTIP(KKI%)
               LOCAMPO = VECLON(KKI%)
               Exit Function
            End If
            '
            Select Case VECTIP(KKI%)
               Case 56
                  If TSQL = 231 Then EXISTELCAMPO% = 2
                  If TSQL = 62 Then EXISTELCAMPO% = 2
               Case 104
                  If TSQL = 52 Or TSQL = 56 Or TSQL = 62 Then EXISTELCAMPO% = 2
               Case 52
                  If TSQL = 56 Or TSQL = 62 Then EXISTELCAMPO% = 2
               Case 173
                  If TSQL = VECTIP(KKI%) And LOCAMPO > VECLON(KKI%) Then EXISTELCAMPO% = 2
               Case 231
                  If TSQL = VECTIP(KKI%) And LOCAMPO > VECLON(KKI%) / 2 Then EXISTELCAMPO% = 2
            End Select
            Exit Function            ' EL CAMPO YA EXISTE
         End If
       End If
     End If
   Next KKI%
   '
   ' AQUI CAE SI EL CAMPO NO EXISTE O HAY QUE MODIFICAR ALGUN DATO
   '
End Function
Function EXISTELCAMPO_DESTINO%(NBASE$, NTABLA$, NCAMPO$, TCAMPO, LOCAMPO, Optional CONINDEX%)
    '
   ' ESTO SE AGREGO PARA NO CREAR LOS CAMPOS YA CREADOS
   Static BASACTI$, CANCAM%, VECTAB$(), IDTAB(), VECCAM$(), VECTIP%(), VECLON%(), VECIDX%()
   NTAB$ = UCase$(NTABLA$)
   NCAM$ = UCase$(NCAMPO$)
   If BASACTI$ <> FLEXCONN_DESTINO.DefaultDatabase Then
      '
      BASACTI$ = FLEXCONN_DESTINO.DefaultDatabase
      CANCAM% = 0: NUEDIM% = 64
      ReDim Preserve VECTAB$(NUEDIM%), IDTAB(NUEDIM%), VECCAM$(NUEDIM%), VECTIP%(NUEDIM%), VECLON%(NUEDIM%), VECIDX%(NUEDIM%)
      STRSQL = "SELECT SYSOBJECTS.NAME AS NOTAB,SYSOBJECTS.ID AS IDENTA, SYSCOLUMNS.NAME AS NOCAM, "
      STRSQL = STRSQL + " SYSCOLUMNS.XTYPE AS TYPE, SYSCOLUMNS.LENGTH AS LENG "
      STRSQL = STRSQL + " FROM SYSCOLUMNS WITH(NOLOCK) "
      STRSQL = STRSQL + " INNER JOIN sysobjects "
      STRSQL = STRSQL & " ON sysobjects.id = syscolumns.id "
      STRSQL = STRSQL + " WHERE Left(SYSOBJECTS.NAME, 3) <> 'sys' "
      STRSQL = STRSQL + " AND LEFT(SYSOBJECTS.NAME,3) <> 'DF_' "
      '
      Set MIRS = FLEXCONN_DESTINO.Execute(STRSQL)
      With MIRS
         Do While Not .EOF
           If Left$(!NOTAB, 1) = UCase$(Left$(!NOTAB, 1)) Then
             CANCAM% = CANCAM% + 1
             If CANCAM% >= UBound(VECTAB$) Then
                NUEDIM% = UBound(VECTAB$) + 64
                ReDim Preserve VECTAB$(NUEDIM%), IDTAB(NUEDIM%), VECCAM$(NUEDIM%), VECTIP%(NUEDIM%), VECLON%(NUEDIM%), VECIDX%(NUEDIM%)
             End If
             IDTAB(CANCAM%) = !IDENTA
             VECTAB$(CANCAM%) = UCase$(!NOTAB)
             VECCAM$(CANCAM%) = UCase$(!NOCAM)
             VECTIP(CANCAM%) = !Type      ' TICABA
             VECLON(CANCAM%) = !LENG      ' LOCABA
           End If
         .MoveNext
         Loop
      End With
      MIRS.Close
      '
      STRSQL = "Select * from sysindexes "
      ' STRSQL = "Select * from sysindexes WHERE id = OBJECT_ID('" & NTABLA$ & "') AND NAME = 'I" + NCAMPO$ + "' "
      Set MIRS = FLEXCONN_DESTINO.Execute(STRSQL)
      With MIRS
         Do While Not .EOF
            ITA = !ID
            NOCA$ = UCase$(Mid$(SAFETEXT$(!Name), 2))
            For KKI% = 1 To CANCAM%
              If IDTAB(KKI%) = ITA Then
                If VECCAM$(KKI%) = NOCA$ Then
                  VECIDX%(KKI%) = 1
                  GoTo 15
                End If
              End If
            Next KKI%
15       .MoveNext
         Loop
      End With
      '
   End If
   '
   '
   If TCAMPO = 11 Then TCAMPO = 12    ' en ambos casos es secuencia ascii
   Select Case TCAMPO
      Case 1: TSQL = 104
      Case 2: TSQL = 173
      Case 3: TSQL = 52
      Case 4: TSQL = 56
      Case 5: TSQL = 60
      Case 6: TSQL = 62
      Case 7: TSQL = 62
      Case 8: TSQL = 58
      Case 10: TSQL = 231
      Case 12: TSQL = 99 ' ojo como trabajarlo
   End Select

19 EXISTELCAMPO_DESTINO% = 0
   For KKI% = 1 To CANCAM%
     If VECTAB$(KKI%) = NTAB$ Then
       If VECCAM$(KKI%) = NCAM$ Then
         If VECIDX%(KKI%) = CONINDEX% Or CONINDEX% < 1 Then
            EXISTELCAMPO_DESTINO% = 3
            If TSQL < 1 Then
               TCAMPO = VECTIP(KKI%)
               LOCAMPO = VECLON(KKI%)
               Exit Function
            End If
            '
            Select Case VECTIP(KKI%)
               Case 56
                  If TSQL = 231 Then EXISTELCAMPO_DESTINO% = 2
                  If TSQL = 62 Then EXISTELCAMPO_DESTINO% = 2
               Case 104
                  If TSQL = 52 Or TSQL = 56 Or TSQL = 62 Then EXISTELCAMPO_DESTINO% = 2
               Case 52
                  If TSQL = 56 Or TSQL = 62 Then EXISTELCAMPO_DESTINO% = 2
               Case 173
                  If TSQL = VECTIP(KKI%) And LOCAMPO > VECLON(KKI%) Then EXISTELCAMPO_DESTINO% = 2
               Case 231
                  If TSQL = VECTIP(KKI%) And LOCAMPO > VECLON(KKI%) / 2 Then EXISTELCAMPO_DESTINO% = 2
            End Select
            Exit Function            ' EL CAMPO YA EXISTE
         End If
       End If
     End If
   Next KKI%
   '
   ' AQUI CAE SI EL CAMPO NO EXISTE O HAY QUE MODIFICAR ALGUN DATO
   '
End Function

'
Sub CREACAMPO_DESTINO(NBASE$, NTABLA$, NCAMPO$, TCAMPO, LOCAMPO, Optional CONINDICE%)
    '
    Static YACREADA%                     ' BANDERA VERIFICADA EXISTENCIA DE BASEACTIVA
    Static TABLASREVISADAS$
    BASEACTIVA_DESTI$ = NBASE$
    If YACREADA% > 0 Then GoTo 15
    '
    
    If BASEACTIVA_DESTI$ = "" Then
       Call MENSERR(24, "Error Fatal en CREACAMPO:\Falta Abrir Conexion y Base Inicial.")
       Call FINAL("")
    End If
    '
10  ' Verifica si existe BASEACTIVA
    FLEXCONN_DESTINO.Execute ("USE Master")
    STRSQL = "Select name from MASTER.DBO.sysdatabases where name = '" + BASEACTIVA_DESTI$ + "'"
    On Error Resume Next
    Set MIRS = FLEXCONN_DESTINO.Execute(STRSQL)
    If Err Then
       Call MENSERR(24, "Error No Recuperable en Rutina 'CREACAMPO'\Buscando B.Datos '" + BASEACTIVA_DESTI$ + "'.")
       Call FINAL("")
    End If
    On Error GoTo 0
    If (MIRS.EOF And MIRS.BOF) Then
       Call MENSERR(24, "Error Fatal en CREACAMPO:\No Existe Base Inicial '" + BASEACTIVA_DESTI$ + "'")
       Call FINAL("")
       ' Nota: Suprimida la opción de crearla aquí para poder definir nombre y ruta.
       '       Se crea con estos parámetros desde el FLOGON.EXE."
    End If
    YACREADA% = 1
    '
15  If YACREADA% < 2 Then
        FLEXCONN_DESTINO.Execute ("USE " + BASEACTIVA_DESTI$)
        YACREADA% = 2
    End If
    '
    EXICA% = EXISTELCAMPO_DESTINO%(NBASE$, NTABLA$, NCAMPO$, TCAMPO, LOCAMPO, CONINDICE%)
    If EXICA% = 3 Then
'       GoTo 85                  ' PARA REVISAR LAS TABLAS CON CODIEMPR%
       Exit Sub
'       If CONINDICE% = 1 Then
'           GoTo 80
'         Else
'           Exit Sub
'       End If
    End If
    ' LO ANTERIOR ES PARA PARA ACELERAR LA APERTURA DE LAS TABLAS.
    ' LA FUNCION EXISTELCAMPO MARCA SI EL CAMPO YA EXISTE EN LA TABLA
    ' Y SI SUS ATRIBUTOS COINCIDEN CON LO INVOCADO - 17 OCTUBRE 2008
    '
    TXCAMPO$ = "[" + NCAMPO$ + "] "
    If TCAMPO = 11 Then TCAMPO = 12    ' en ambos casos es secuencia ascii
    Select Case TCAMPO
      Case 1
        TXCAMPO$ = TXCAMPO$ + "[bit]": TSQL = 104
      Case 2
        TXCAMPO$ = TXCAMPO$ + "[binary] (" + TRIM$(Str$(LOCAMPO)) + ")": TSQL = 173
      Case 3
        TXCAMPO$ = TXCAMPO$ + "[smallint]": TSQL = 52
      Case 4
        TXCAMPO$ = TXCAMPO$ + "[int]": TSQL = 56
      Case 5
        TXCAMPO$ = TXCAMPO$ + "[money]": TSQL = 60
      Case 6
        TXCAMPO$ = TXCAMPO$ + "[float]": TSQL = 62
      Case 7
        TXCAMPO$ = TXCAMPO$ + "[float]": TSQL = 62
      Case 8
        TXCAMPO$ = TXCAMPO$ + "[smalldatetime]": TSQL = 58
      Case 10
        TXCAMPO$ = TXCAMPO$ + "[nvarchar] (" + TRIM$(Str$(LOCAMPO)) + ") COLLATE Modern_Spanish_CI_AS": TSQL = 231
      Case 12
        TXCAMPO$ = TXCAMPO$ + "[ntext] ": TSQL = 99 ' ojo como trabajarlo
      Case Else
        Call MENSERR(24, "Imposible Crear Campo '" + NCAMPO$ + "' con Tipo = " & TCAMPO)
        Call FINAL("")
    End Select
    TXCAMPO = TXCAMPO$ + " NULL"
    '
20  VUELTA% = 0
    FLEXCONN_DESTINO.Execute ("USE " + BASEACTIVA_DESTI$)
22  SQLX$ = "Select NAME from sysobjects where name = '" + NTABLA$ + "'"
    Set MIRS = FLEXCONN_DESTINO.Execute(SQLX$)
    If (MIRS.EOF And MIRS.BOF) Then
       If VUELTA% < 1 Then
           STRSQL = "CREATE TABLE " & NTABLA$ & " (" + TXCAMPO$ + ") ON [PRIMARY]"
           On Error Resume Next
           FLEXCONN_DESTINO.Execute (STRSQL)      ' CREO LA TABLA Y EL PRIMER CAMPO
           MIRS.Close
           Set MIRS = Nothing
           On Error GoTo 0
           VUELTA% = VUELTA% + 1
           GoTo 22                       ' PARA VERIFICAR
         Else
           Call MENSERR(24, "Imposible Crear la Tabla '" + NTABLA$ + "' en '" + BASEACTIVA_DESTI$ + "'")
       End If
    End If
    '
    On Error Resume Next
    FLEXCONN_DESTINO.Execute ("ALTER TABLE " & NTABLA$ & " ADD " & TXCAMPO$)
    On Error GoTo 0
    '
'    VIEJA SELECCION DE ALEJANDRA CIGNA
'    STRSQL = "Select systypes.name, systypes.xtype as xtype1 from sysobjects inner join syscolumns on "
'    STRSQL = STRSQL & "sysobjects.id = syscolumns.id inner join systypes "
'    STRSQL = STRSQL & "on syscolumns.xtype = systypes.xtype "
'    STRSQL = STRSQL & "where sysobjects.name = '" & NTABLA$ & "' "
'    STRSQL = STRSQL & "and syscolumns.name ='" & NCAMPO$ & "'"
    STRSQL = "SELECT SYSCOLUMNS.NAME AS NAME, SYSCOLUMNS.XTYPE AS XTYPE1, SYSCOLUMNS.LENGTH AS LENG FROM syscolumns "
    STRSQL = STRSQL + "INNER JOIN sysobjects "
    STRSQL = STRSQL & "ON sysobjects.id = syscolumns.id "
    STRSQL = STRSQL & "where sysobjects.name = '" & NTABLA$ & "' "
    STRSQL = STRSQL & "AND syscolumns.name = '" & UCase$(NCAMPO$) & "'"
    ' STRSQL = STRSQL & "AND syscolumns.name LIKE '" & UCase$(NCAMPO$) & "'"
    ' EL 1/4/2009 SE CAMBIO LIKE POR = YA QUE NO CREABA EL CAMPO DESC_ITEM POR
    ' CONFUNDIRLO CON EL DESCRITEM. QUEDA POR AVERIGUAR POR QUÉ SE PUSO ENTES EL
    ' LIKE EN VEZ DEL IGUAL. - epb
    Set MIRS = FLEXCONN_DESTINO.Execute(STRSQL)
    If Not (MIRS.EOF And MIRS.BOF) Then
        If MIRS!xtype1 <> TSQL Or ((TSQL = 173 Or TSQL = 231) And LOCAMPO > MIRS!LENG / 2) Then
           '
           ' PRESENTA MENSAJE DE ADVERTENCIA
           TECOMU$ = "FLEXOFT Debe Ahora Modificar el Formato del Campo\'" + NCAMPO$ + "' en la Tabla '" + NTABLA$ + "' de " & MIRS!xtype1
           If (TSQL = 173 Or TSQL = 231) Then TECOMU$ = TECOMU$ + "(" & MIRS!LENG / 2 & ")"
           TECOMU$ = TECOMU$ & " a " & TSQL
           If (TSQL = 173 Or TSQL = 231) Then TECOMU$ = TECOMU$ + "(" & LOCAMPO & ")"
           TECOMU$ = TECOMU$ & ".\   \Esto Puede Demandar Varios Minutos.\   \Pulse 'Aceptar' Para Continuar."
           Call COMUNI(TECOMU$)
           '
           ' verifica si existe UN INDICE
           Screen.MousePointer = 11
           STRSQL = "Select * from sysindexes WHERE id = OBJECT_ID('" & NTABLA$ & "') AND NAME = 'I" + NCAMPO$ + "' "
           Set MIRS1 = FLEXCONN_DESTINO.Execute(STRSQL)
           If Not (MIRS1.EOF And MIRS1.BOF) Then   ' EXISTE INDICE - LO BORRA
              On Error Resume Next
               FLEXCONN_DESTINO.Execute ("DROP INDEX " + NTABLA$ + ".I" + NCAMPO$)
              On Error GoTo 0
           End If
           On Error Resume Next          ' por si no se puede modificar
            FLEXCONN_DESTINO.Execute ("ALTER TABLE " & NTABLA$ & " ALTER COLUMN " & TXCAMPO$)
           Screen.MousePointer = 1
           If Err Then
              On Error GoTo 0
              Call MENSERR(24, "No Fue Posible Modificar Campo '" + NCAMPO$ + "'\en la Tabla '" + NTABLA$ + "' de " & MIRS!xtype1 & " a " & TSQL & ".\Resultados Impredecibles. Consulte.")
           End If
           On Error GoTo 0
        End If
    End If
    '
    ' AQUI GENERAR LOS INDICES
80  If CONINDICE% = 1 Then
       ' verifica si el indice existe
       STRSQL = "Select * from sysindexes WHERE id = OBJECT_ID('" & NTABLA$ & "') AND NAME = 'I" + NCAMPO$ + "' "
       Set MIRS = FLEXCONN_DESTINO.Execute(STRSQL)
       If Not (MIRS.EOF And MIRS.BOF) Then
          GoTo 90                              ' el indice ya existe
       End If
       '
       TIMUT = FLEXCONN_DESTINO.CommandTimeout
       FLEXCONN_DESTINO.CommandTimeout = 999
       CRINDEX$ = "CREATE INDEX I" + NCAMPO$ + " ON " + NTABLA$ + "([" + NCAMPO$ + "]);"
       On Error Resume Next   ' PARA ATAJAR QUE EL INDICE YA EXISTA
       FLEXCONN_DESTINO.Execute CRINDEX$
       If Err = False Then
           If MIGRATODO% <> 1 Then
             MsgBox "INDICE I" + NCAMPO$ + " CREADO EN TABLA " + NTABLA$
           End If
         Else
           MsgBox "Imposible Crear Indice I" + NCAMPO$ + " EN TABLA " + NTABLA$ + ". Consulte para Creación Manual del Indice Requerido."
       End If
       On Error GoTo 0
       FLEXCONN_DESTINO.CommandTimeout = TIMUT
    End If
   '
'85 ' VERIFICA CAMPO CODIEMPR EN LA TABLA QUE ESTA ABRIENDO
'   If InStr(TABLASREVISADAS$, "@@@" + UCase$(NTABLA$)) < 1 Then
'      TABLASREVISADAS$ = TABLASREVISADAS$ + "@@@" + UCase$(NTABLA$)
'      If UCase$(Left$(NTABLA$, 1)) = Left$(NTABLA$, 1) Then
'        If Left$(NTABLA$, 3) <> "DF_" Then
'          Call CREACAMPO_DESTINO("", NTABLA$, "CodiEmpr", 3, 2, 1)
'          QTR& = CantRegistros(NTABLA$, "CodiEmpr IS NULL")
'          If QTR& > 0 Then
'             SQLX$ = "UPDATE " + NTABLA$ + " SET CodiEmpr = '" & CodiEmp% & "' WHERE CodiEmpr IS NULL"
'             On Error Resume Next
'             FLEXCONN.Execute (SQLX$)
'             If Err Then
'                On Error GoTo 0
'                Call PONE_CODIEMPR_EN_CERO(NTABLA$)
'             End If
'             On Error GoTo 0
'          End If
'        End If
'      End If
'   End If
   '
90 End Sub
Sub CREACAMPO(NBASE$, NTABLA$, NCAMPO$, TCAMPO, LOCAMPO, Optional CONINDICE%)
    '
    Static YACREADA%                     ' BANDERA VERIFICADA EXISTENCIA DE BASEACTIVA
    Static TABLASREVISADAS$
    If YACREADA% > 0 Then GoTo 15
    '
    If BASEACTIVA$ = "" Then
       Call MENSERR(24, "Error Fatal en CREACAMPO:\Falta Abrir Conexion y Base Inicial.")
       Call FINAL("")
    End If
    '
10  ' Verifica si existe BASEACTIVA
    FLEXCONN.Execute ("USE Master")
    STRSQL = "Select name from MASTER.DBO.sysdatabases where name = '" + BASEACTIVA$ + "'"
    On Error Resume Next
    Set MIRS = FLEXCONN.Execute(STRSQL)
    If Err Then
       Call MENSERR(24, "Error No Recuperable en Rutina 'CREACAMPO'\Buscando B.Datos '" + BASEACTIVA$ + "'.")
       Call FINAL("")
    End If
    On Error GoTo 0
    If (MIRS.EOF And MIRS.BOF) Then
       Call MENSERR(24, "Error Fatal en CREACAMPO:\No Existe Base Inicial '" + BASEACTIVA + "'")
       Call FINAL("")
       ' Nota: Suprimida la opción de crearla aquí para poder definir nombre y ruta.
       '       Se crea con estos parámetros desde el FLOGON.EXE."
    End If
    YACREADA% = 1
    '
15  If YACREADA% < 2 Then
        FLEXCONN.Execute ("USE " + BASEACTIVA$)
        YACREADA% = 2
    End If
    '
    EXICA% = EXISTELCAMPO%(NBASE$, NTABLA$, NCAMPO$, TCAMPO, LOCAMPO, CONINDICE%)
    If EXICA% = 3 Then
       GoTo 85                  ' PARA REVISAR LAS TABLAS CON CODIEMPR%
       Exit Sub
'       If CONINDICE% = 1 Then
'           GoTo 80
'         Else
'           Exit Sub
'       End If
    End If
    ' LO ANTERIOR ES PARA PARA ACELERAR LA APERTURA DE LAS TABLAS.
    ' LA FUNCION EXISTELCAMPO MARCA SI EL CAMPO YA EXISTE EN LA TABLA
    ' Y SI SUS ATRIBUTOS COINCIDEN CON LO INVOCADO - 17 OCTUBRE 2008
    '
    TXCAMPO$ = "[" + NCAMPO$ + "] "
    If TCAMPO = 11 Then TCAMPO = 12    ' en ambos casos es secuencia ascii
    Select Case TCAMPO
      Case 1
        TXCAMPO$ = TXCAMPO$ + "[bit]": TSQL = 104
      Case 2
        TXCAMPO$ = TXCAMPO$ + "[binary] (" + TRIM$(Str$(LOCAMPO)) + ")": TSQL = 173
      Case 3
        TXCAMPO$ = TXCAMPO$ + "[smallint]": TSQL = 52
      Case 4
        TXCAMPO$ = TXCAMPO$ + "[int]": TSQL = 56
      Case 5
        TXCAMPO$ = TXCAMPO$ + "[money]": TSQL = 60
      Case 6
        TXCAMPO$ = TXCAMPO$ + "[float]": TSQL = 62
      Case 7
        TXCAMPO$ = TXCAMPO$ + "[float]": TSQL = 62
      Case 8
        TXCAMPO$ = TXCAMPO$ + "[smalldatetime]": TSQL = 58
      Case 10
        TXCAMPO$ = TXCAMPO$ + "[nvarchar] (" + TRIM$(Str$(LOCAMPO)) + ") COLLATE Modern_Spanish_CI_AS": TSQL = 231
      Case 12
        TXCAMPO$ = TXCAMPO$ + "[ntext] ": TSQL = 99 ' ojo como trabajarlo
      Case Else
        Call MENSERR(24, "Imposible Crear Campo '" + NCAMPO$ + "' con Tipo = " & TCAMPO)
        Call FINAL("")
    End Select
    TXCAMPO = TXCAMPO$ + " NULL"
    '
20  VUELTA% = 0
    FLEXCONN.Execute ("USE " + BASEACTIVA$)
22  SQLX$ = "Select NAME from sysobjects where name = '" + NTABLA$ + "'"
    Set MIRS = FLEXCONN.Execute(SQLX$)
    If (MIRS.EOF And MIRS.BOF) Then
       If VUELTA% < 1 Then
           STRSQL = "CREATE TABLE " & NTABLA$ & " (" + TXCAMPO$ + ") ON [PRIMARY]"
           On Error Resume Next
           FLEXCONN.Execute (STRSQL)      ' CREO LA TABLA Y EL PRIMER CAMPO
           MIRS.Close
           Set MIRS = Nothing
           On Error GoTo 0
           VUELTA% = VUELTA% + 1
           GoTo 22                       ' PARA VERIFICAR
         Else
           Call MENSERR(24, "Imposible Crear la Tabla '" + NTABLA$ + "' en '" + BASEACTIVA$ + "'")
       End If
    End If
    '
    On Error Resume Next
    FLEXCONN.Execute ("ALTER TABLE " & NTABLA$ & " ADD " & TXCAMPO$)
    On Error GoTo 0
    '
'    VIEJA SELECCION DE ALEJANDRA CIGNA
'    STRSQL = "Select systypes.name, systypes.xtype as xtype1 from sysobjects inner join syscolumns on "
'    STRSQL = STRSQL & "sysobjects.id = syscolumns.id inner join systypes "
'    STRSQL = STRSQL & "on syscolumns.xtype = systypes.xtype "
'    STRSQL = STRSQL & "where sysobjects.name = '" & NTABLA$ & "' "
'    STRSQL = STRSQL & "and syscolumns.name ='" & NCAMPO$ & "'"
    STRSQL = "SELECT SYSCOLUMNS.NAME AS NAME, SYSCOLUMNS.XTYPE AS XTYPE1, SYSCOLUMNS.LENGTH AS LENG FROM syscolumns "
    STRSQL = STRSQL + "INNER JOIN sysobjects "
    STRSQL = STRSQL & "ON sysobjects.id = syscolumns.id "
    STRSQL = STRSQL & "where sysobjects.name = '" & NTABLA$ & "' "
    STRSQL = STRSQL & "AND syscolumns.name = '" & UCase$(NCAMPO$) & "'"
    ' STRSQL = STRSQL & "AND syscolumns.name LIKE '" & UCase$(NCAMPO$) & "'"
    ' EL 1/4/2009 SE CAMBIO LIKE POR = YA QUE NO CREABA EL CAMPO DESC_ITEM POR
    ' CONFUNDIRLO CON EL DESCRITEM. QUEDA POR AVERIGUAR POR QUÉ SE PUSO ENTES EL
    ' LIKE EN VEZ DEL IGUAL. - epb
    Set MIRS = FLEXCONN.Execute(STRSQL)
    If Not (MIRS.EOF And MIRS.BOF) Then
        If MIRS!xtype1 <> TSQL Or ((TSQL = 173 Or TSQL = 231) And LOCAMPO > MIRS!LENG / 2) Then
           '
           ' PRESENTA MENSAJE DE ADVERTENCIA
           TECOMU$ = "FLEXOFT Debe Ahora Modificar el Formato del Campo\'" + NCAMPO$ + "' en la Tabla '" + NTABLA$ + "' de " & MIRS!xtype1
           If (TSQL = 173 Or TSQL = 231) Then TECOMU$ = TECOMU$ + "(" & MIRS!LENG / 2 & ")"
           TECOMU$ = TECOMU$ & " a " & TSQL
           If (TSQL = 173 Or TSQL = 231) Then TECOMU$ = TECOMU$ + "(" & LOCAMPO & ")"
           TECOMU$ = TECOMU$ & ".\   \Esto Puede Demandar Varios Minutos.\   \Pulse 'Aceptar' Para Continuar."
           Call COMUNI(TECOMU$)
           '
           ' verifica si existe UN INDICE
           Screen.MousePointer = 11
           STRSQL = "Select * from sysindexes WHERE id = OBJECT_ID('" & NTABLA$ & "') AND NAME = 'I" + NCAMPO$ + "' "
           Set MIRS1 = FLEXCONN.Execute(STRSQL)
           If Not (MIRS1.EOF And MIRS1.BOF) Then   ' EXISTE INDICE - LO BORRA
              On Error Resume Next
               FLEXCONN.Execute ("DROP INDEX " + NTABLA$ + ".I" + NCAMPO$)
              On Error GoTo 0
           End If
           On Error Resume Next          ' por si no se puede modificar
            FLEXCONN.Execute ("ALTER TABLE " & NTABLA$ & " ALTER COLUMN " & TXCAMPO$)
           Screen.MousePointer = 1
           If Err Then
              On Error GoTo 0
              Call MENSERR(24, "No Fue Posible Modificar Campo '" + NCAMPO$ + "'\en la Tabla '" + NTABLA$ + "' de " & MIRS!xtype1 & " a " & TSQL & ".\Resultados Impredecibles. Consulte.")
           End If
           On Error GoTo 0
        End If
    End If
    '
    ' AQUI GENERAR LOS INDICES
80  If CONINDICE% = 1 Then
       ' verifica si el indice existe
       STRSQL = "Select * from sysindexes WHERE id = OBJECT_ID('" & NTABLA$ & "') AND NAME = 'I" + NCAMPO$ + "' "
       Set MIRS = FLEXCONN.Execute(STRSQL)
       If Not (MIRS.EOF And MIRS.BOF) Then
          GoTo 90                              ' el indice ya existe
       End If
       '
       TIMUT = FLEXCONN.CommandTimeout
       FLEXCONN.CommandTimeout = 999
       CRINDEX$ = "CREATE INDEX I" + NCAMPO$ + " ON " + NTABLA$ + "([" + NCAMPO$ + "]);"
       On Error Resume Next   ' PARA ATAJAR QUE EL INDICE YA EXISTA
       FLEXCONN.Execute CRINDEX$
       If Err = False Then
           If MIGRATODO% <> 1 Then
             MsgBox "INDICE I" + NCAMPO$ + " CREADO EN TABLA " + NTABLA$
           End If
         Else
           MsgBox "Imposible Crear Indice I" + NCAMPO$ + " EN TABLA " + NTABLA$ + ". Consulte para Creación Manual del Indice Requerido."
       End If
       On Error GoTo 0
       FLEXCONN.CommandTimeout = TIMUT
    End If
   '
85 ' VERIFICA CAMPO CODIEMPR EN LA TABLA QUE ESTA ABRIENDO
   If InStr(TABLASREVISADAS$, "@@@" + UCase$(NTABLA$)) < 1 Then
      TABLASREVISADAS$ = TABLASREVISADAS$ + "@@@" + UCase$(NTABLA$)
      If UCase$(Left$(NTABLA$, 1)) = Left$(NTABLA$, 1) Then
        If Left$(NTABLA$, 3) <> "DF_" Then
          Call CREACAMPO("", NTABLA$, "CodiEmpr", 3, 2, 1)
          QTR& = CantRegistros(NTABLA$, "CodiEmpr IS NULL")
          If QTR& > 0 Then
             SQLX$ = "UPDATE " + NTABLA$ + " SET CodiEmpr = '" & CodiEmp% & "' WHERE CodiEmpr IS NULL"
             On Error Resume Next
             FLEXCONN.Execute (SQLX$)
             If Err Then
                On Error GoTo 0
                Call PONE_CODIEMPR_EN_CERO(NTABLA$)
             End If
             On Error GoTo 0
          End If
        End If
      End If
   End If
   '
90 End Sub


Sub COPIA_STRUC_MOVISTO()
    '
    Call ABRECONEXION
    Call ABRESTOCKS
    Dim MOVISEL As ADODB.Recordset
    Set MOVISEL = READSET("SELECT * FROM MOVISTO WITH(NOLOCK)")
    FIN& = MOVISEL.Fields.Count
    KU& = 0
    For KKI% = 1 To MOVISEL.Fields.Count
       KU& = KU& + 1
       Call TRACE(20, KU&, FIN&)
       NOCAM$ = MOVISEL.Fields(KKI% - 1).Name
       TCAMPO = TIPODECAMPO("MOVISTO", NOCAM$)
       LOCAM = LONCAMPO("MOVISTO", NOCAM$)
       CONIN% = INDICEDECAMPO%("MOVISTO", NOCAM$)
       '
       TICAM = 10
       Select Case TCAMPO
          Case 104
            TICAM = 1
          Case 173
            TICAM = 2
          Case 52
            TICAM = 3
          Case 56
            TICAM = 4
          Case 60
            TICAM = 5
          Case 62
            TICAM = 7
          Case 58
            TICAM = 8
          Case 231
            TICAM = 10
          Case 99
            TICAM = 12
          Case Else
            Call MENSERR(24, "Imposible Crear Campo '" + NCAMPO$ + "' con Tipo = " & TCAMPO)
            Call FINAL("")
       End Select
       '
       Call CREACAMPO("", "MOVIVIE", NOCAM$, TICAM, LOCAM, CONIN%)
       '
    Next KKI%
    '
End Sub

Sub GENERBASE(NOMBASE$)
   '
   If TRIM$(UCase$(NOMBASE$)) = "ARTICULOS" Then Call CREATABLA_ARTICULOS: Exit Sub
   If TRIM$(UCase$(NOMBASE$)) = "BADACLIEN" Then Call CREATABLA_BADACLIEN: Exit Sub
   If TRIM$(UCase$(NOMBASE$)) = "BADAPROVE" Then Call CREATABLA_BADAPROVE: Exit Sub
   If TRIM$(UCase$(NOMBASE$)) = "BCONTROL" Then Call CREATABLA_BCONTROL: Exit Sub
   If TRIM$(UCase$(NOMBASE$)) = "COMPRAS" Then Call CREATABLA_COMPRAS: Exit Sub
   If TRIM$(UCase$(NOMBASE$)) = "CONTAGEN" Then Call CREATABLA_CONTAGEN: Exit Sub
   If TRIM$(UCase$(NOMBASE$)) = "CUECORRI" Then Call CREATABLA_CUECORRI: Exit Sub
   If TRIM$(UCase$(NOMBASE$)) = "EMPLEADOS" Then Call CREATABLA_EMPLEADOS: Exit Sub
   If TRIM$(UCase$(NOMBASE$)) = "EQUIPOS" Then Call CREATABLA_EQUIPOS: Exit Sub
   If TRIM$(UCase$(NOMBASE$)) = "MAQUINARIAS" Then Call CREATABLA_MAQUINARIAS: Exit Sub
   If TRIM$(UCase$(NOMBASE$)) = "MENSAFLEX" Then Call CREATABLA_MENSAFLEX: Exit Sub
   If TRIM$(UCase$(NOMBASE$)) = "OFABRIC" Then Call CREATABLA_OFABRIC: Exit Sub
   If TRIM$(UCase$(NOMBASE$)) = "OMANTIN" Then Call CREATABLA_OMANTIN: Exit Sub
   If TRIM$(UCase$(NOMBASE$)) = "OTRABAJO" Then Call CREATABLA_OTRABAJO: Exit Sub
   If TRIM$(UCase$(NOMBASE$)) = "PRESUCAM" Then Call CREATABLA_PRESUCAM: Exit Sub
   If TRIM$(UCase$(NOMBASE$)) = "STOCKORIGI" Then Call CREATABLA_STOCKORIGI: Exit Sub
   If TRIM$(UCase$(NOMBASE$)) = "STOCKS" Then Call CREATABLA_STOCKS: Exit Sub
   If TRIM$(UCase$(NOMBASE$)) = "STOCKVIE" Then Call CREATABLA_STOCKVIE: Exit Sub
   If TRIM$(UCase$(NOMBASE$)) = "VENTAS" Then Call CREATABLA_VENTAS: Exit Sub
   '
End Sub
'
Sub CREATABLA_ARTICULOS()
   '
   Call CREACAMPO("ARTICULOS", "BADAMODE", "Cod_Mod", 10, 16)
   Call CREACAMPO("ARTICULOS", "BADAMODE", "Descripcion", 10, 64)
   Call CREACAMPO("ARTICULOS", "BADAMODE", "Ubic", 10, 32)
   Call CREACAMPO("ARTICULOS", "BADAMODE", "CantPiezas", 6, 0)
   Call CREACAMPO("ARTICULOS", "BADAMODE", "Frecuen", 10, 32)
   Call CREACAMPO("ARTICULOS", "BADAMODE", "FechUlt_Con", 8, 0)
   Call CREACAMPO("ARTICULOS", "BADAMODE", "FechPro_Con", 8, 0)
   Call CREACAMPO("ARTICULOS", "BADAMODE", "Material", 3, 0)
   Call CREACAMPO("ARTICULOS", "BADAMODE", "Desc_Mat", 10, 32)
   Call CREACAMPO("ARTICULOS", "BADAMODE", "Cod_Inte", 3, 0)
   Call CREACAMPO("ARTICULOS", "BADAMODE", "AnotaMod", 12, 0)
   Call CREACAMPO("ARTICULOS", "BADAMODE", "ObservaMod", 12, 0)
   Call CREACAMPO("ARTICULOS", "BADAMODE", "Foto1", 10, 128)
   Call CREACAMPO("ARTICULOS", "BADAMODE", "Foto2", 10, 128)
   Call CREACAMPO("ARTICULOS", "BADAMODE", "Foto3", 10, 128)
   Call CREACAMPO("ARTICULOS", "BADAMODE", "Foto4", 10, 128)
   '
   Call CREACAMPO("ARTICULOS", "DOCUREL", "Cod_Inte", 3, 0, 1)
   Call CREACAMPO("ARTICULOS", "DOCUREL", "Descripcion", 10, 24)
   Call CREACAMPO("ARTICULOS", "DOCUREL", "Ruta", 10, 512)
   Call CREACAMPO("ARTICULOS", "DOCUREL", "Marborra", 3, 0)
   '
   Call CREACAMPO("ARTICULOS", "Formulas", "CodIConj", 3, 0, 1)
   Call CREACAMPO("ARTICULOS", "Formulas", "CodXConj", 10, 16)
   Call CREACAMPO("ARTICULOS", "Formulas", "CodiElem", 3, 0, 1)
   Call CREACAMPO("ARTICULOS", "Formulas", "CodXElem", 10, 16)
   Call CREACAMPO("ARTICULOS", "Formulas", "NuorIt", 3, 2, 1)
   Call CREACAMPO("ARTICULOS", "Formulas", "UsoBruto", 6, 0)
   Call CREACAMPO("ARTICULOS", "Formulas", "Unis_Form", 6, 0)
   Call CREACAMPO("ARTICULOS", "Formulas", "Lar_Corte", 6, 0)
   Call CREACAMPO("ARTICULOS", "Formulas", "Pladispo", 3, 0)
   Call CREACAMPO("ARTICULOS", "Formulas", "RefeComp", 12, 0)
   Call CREACAMPO("ARTICULOS", "Formulas", "Marborra", 3, 0)
   '
   Call CREACAMPO("ARTICULOS", "FormuPat", "CodIConj", 3, 2, 1)
   Call CREACAMPO("ARTICULOS", "FormuPat", "CodxConj", 10, 16)
   Call CREACAMPO("ARTICULOS", "FormuPat", "CodiElem", 3, 2, 1)
   Call CREACAMPO("ARTICULOS", "FormuPat", "CodXElem", 10, 16)
   Call CREACAMPO("ARTICULOS", "FormuPat", "NuorIt", 3, 2)
   Call CREACAMPO("ARTICULOS", "FormuPat", "UsoBruto", 6, 4)
   Call CREACAMPO("ARTICULOS", "FormuPat", "Operacion", 3, 2)
   Call CREACAMPO("ARTICULOS", "FormuPat", "Marborra", 3, 2)
   Call CREACAMPO("ARTICULOS", "FormuPat", "Acepta_Min", 6, 0)
   Call CREACAMPO("ARTICULOS", "FormuPat", "Acepta_Max", 6, 0)
   Call CREACAMPO("ARTICULOS", "FormuPat", "Acepta_Prec", 6, 0)
   '
   Call CREACAMPO("ARTICULOS", "GRAFICOS", "Codigo", 3, 2)
   Call CREACAMPO("ARTICULOS", "GRAFICOS", "Nume_Ord", 3, 2)
   '
   Call CREACAMPO("ARTICULOS", "HERRAMEN", "Codigo", 10, 16, 1)
   Call CREACAMPO("ARTICULOS", "HERRAMEN", "Descripcion", 10, 64)
   Call CREACAMPO("ARTICULOS", "HERRAMEN", "Marborra", 3, 0)
   '
   Call CREACAMPO("ARTICULOS", "HERRAOPER", "Codint", 3, 0, 1)
   Call CREACAMPO("ARTICULOS", "HERRAOPER", "Cod_Oper", 3, 0)
   Call CREACAMPO("ARTICULOS", "HERRAOPER", "Cod_Herramental", 10, 16)
   Call CREACAMPO("ARTICULOS", "HERRAOPER", "Cantidad", 6, 0)
   Call CREACAMPO("ARTICULOS", "HERRAOPER", "Marborra", 3, 0)
   Call CREACAMPO("ARTICULOS", "HERRAOPER", "Tipo_Herr", 3, 0)
   '
   Call CREACAMPO("ARTICULOS", "Master", "Ferevis", 8, 0)
   Call CREACAMPO("ARTICULOS", "Master", "Observa", 12, 0)
   Call CREACAMPO("ARTICULOS", "Master", "Especificacion", 12, 0)
   Call CREACAMPO("ARTICULOS", "Master", "DesCompAut", 3, 0)
   Call CREACAMPO("ARTICULOS", "Master", "DescriComp", 10, 64)
   Call CREACAMPO("ARTICULOS", "Master", "DepoEntra", 3, 0)
   Call CREACAMPO("ARTICULOS", "Master", "DepoCon", 3, 0)
   Call CREACAMPO("ARTICULOS", "Master", "Codigo", 10, 24, 1)
   Call CREACAMPO("ARTICULOS", "Master", "Descripcion", 10, 64)
   Call CREACAMPO("ARTICULOS", "Master", "Descridos", 10, 128)
   Call CREACAMPO("ARTICULOS", "Master", "Status", 3, 0)
   Call CREACAMPO("ARTICULOS", "Master", "Codint", 4, 0, 1)
   Call CREACAMPO("ARTICULOS", "Master", "Alias", 10, 24, 1)
   Call CREACAMPO("ARTICULOS", "Master", "Umedida", 10, 4)
   Call CREACAMPO("ARTICULOS", "Master", "Timate", 3, 0, 1)
   Call CREACAMPO("ARTICULOS", "Master", "Critrep", 3, 0)
   Call CREACAMPO("ARTICULOS", "Master", "Monecos", 3, 0)
   Call CREACAMPO("ARTICULOS", "Master", "Fecosto", 8, 0, 1)
   Call CREACAMPO("ARTICULOS", "Master", "PorDescCos", 6, 0)
   Call CREACAMPO("ARTICULOS", "Master", "Cosuni", 6, 0)
   Call CREACAMPO("ARTICULOS", "Master", "Pornacio", 6, 0)
   Call CREACAMPO("ARTICULOS", "Master", "PorComVen", 6, 0)
   Call CREACAMPO("ARTICULOS", "Master", "Provhabi", 3, 0)
   Call CREACAMPO("ARTICULOS", "Master", "Stomin", 6, 0, 1)
   Call CREACAMPO("ARTICULOS", "Master", "Stomax", 6, 0)
   Call CREACAMPO("ARTICULOS", "Master", "Lorepos", 6, 0, 1)
   Call CREACAMPO("ARTICULOS", "Master", "Lindivis", 6, 0)
   Call CREACAMPO("ARTICULOS", "Master", "Pesuni", 6, 0)
   Call CREACAMPO("ARTICULOS", "Master", "PesMetro", 6, 0)
   Call CREACAMPO("ARTICULOS", "Master", "Color", 10, 32)
   Call CREACAMPO("ARTICULOS", "Master", "Talle", 10, 32)
   Call CREACAMPO("ARTICULOS", "Master", "Descompo", 3, 0)
   Call CREACAMPO("ARTICULOS", "Master", "Grupiva", 3, 0)
   Call CREACAMPO("ARTICULOS", "Master", "Grucoven", 3, 0)
   Call CREACAMPO("ARTICULOS", "Master", "CLinea", 10, 8)
   Call CREACAMPO("ARTICULOS", "Master", "Cfamilia", 10, 8)
   Call CREACAMPO("ARTICULOS", "Master", "NPlano", 10, 24)
   Call CREACAMPO("ARTICULOS", "Master", "Revision", 10, 16)
   Call CREACAMPO("ARTICULOS", "Master", "Trazable", 3, 0)
   Call CREACAMPO("ARTICULOS", "Master", "Revi_Formu", 8, 0)
   Call CREACAMPO("ARTICULOS", "Master", "FechApFor", 8, 0)
   Call CREACAMPO("ARTICULOS", "Master", "UsuApFor", 10, 24)
   Call CREACAMPO("ARTICULOS", "Master", "Cod_Interna", 10, 64)
   Call CREACAMPO("ARTICULOS", "Master", "Bien_Cap", 3, 0)
   Call CREACAMPO("ARTICULOS", "Master", "Codigo_Ncm", 10, 32)
   Call CREACAMPO("ARTICULOS", "Master", "Codigo_Sec", 10, 32)
   Call CREACAMPO("ARTICULOS", "Master", "Pro_Umed", 10, 32)
   Call CREACAMPO("ARTICULOS", "Master", "Iva_Id", 10, 8)
   Call CREACAMPO("ARTICULOS", "Master", "ArchiPdf", 12, 0)
   Call CREACAMPO("ARTICULOS", "Master", "Familia", 10, 4)
   Call CREACAMPO("ARTICULOS", "Master", "Lindiven", 6, 0)
   Call CREACAMPO("ARTICULOS", "Master", "CanDeci", 10, 2)
   Call CREACAMPO("ARTICULOS", "Master", "ConsuProm", 6, 0)
   Call CREACAMPO("ARTICULOS", "Master", "CosCal", 7, 0)
   Call CREACAMPO("ARTICULOS", "Master", "FeCosCal", 8, 0)
   Call CREACAMPO("ARTICULOS", "Master", "Stat_Movi", 3, 0)
   Call CREACAMPO("ARTICULOS", "Master", "Clas_ATC", 10, 64)
   Call CREACAMPO("ARTICULOS", "Master", "Vers_BC", 10, 32)
   Call CREACAMPO("ARTICULOS", "Master", "FeVig_BC", 8, 0)
   Call CREACAMPO("ARTICULOS", "Master", "FeRev_BC", 8, 0)
   Call CREACAMPO("ARTICULOS", "Master", "CantApro", 3, 0)
   Call CREACAMPO("ARTICULOS", "Master", "LeTiCom", 10, 16)
   Call CREACAMPO("ARTICULOS", "Master", "PiezasXCaja", 4, 0)
   Call CREACAMPO("ARTICULOS", "Master", "M3Caja", 6, 0)
   Call CREACAMPO("ARTICULOS", "Master", "KiloNetoCaja", 6, 0)
   Call CREACAMPO("ARTICULOS", "Master", "KiloBrutoCaja", 6, 0)
   Call CREACAMPO("ARTICULOS", "Master", "Autono_Proy", 10, 10)
   Call CREACAMPO("ARTICULOS", "Master", "UsuPrePes", 10, 24)
   Call CREACAMPO("ARTICULOS", "Master", "UsuPreDol", 10, 24)
   Call CREACAMPO("ARTICULOS", "Master", "MANO_OBRA_REAL", 6, 0)
   Call CREACAMPO("ARTICULOS", "Master", "Tamano_Etiqueta", 3, 0)
   '
   Call CREACAMPO("", "CODICLIENTE", "Codigo_Cliente", 10, 24, 1)
   Call CREACAMPO("", "CODICLIENTE", "Status", 3, 0)
   Call CREACAMPO("", "CODICLIENTE", "Codint", 4, 0, 1)

   '
   Call CREACAMPO("ARTICULOS", "SECOPER", "CodIConj", 3, 0, 1)
   Call CREACAMPO("ARTICULOS", "SECOPER", "CodXConj", 10, 24)
   Call CREACAMPO("ARTICULOS", "SECOPER", "NumeOper", 3, 0)
   Call CREACAMPO("ARTICULOS", "SECOPER", "AltOper", 3, 0)
   Call CREACAMPO("ARTICULOS", "SECOPER", "DescOper", 10, 64)
   Call CREACAMPO("ARTICULOS", "SECOPER", "StatOper", 3, 0)
   Call CREACAMPO("ARTICULOS", "SECOPER", "CodMaq0", 10, 6)
   Call CREACAMPO("ARTICULOS", "SECOPER", "CodMaq1", 10, 6)
   Call CREACAMPO("ARTICULOS", "SECOPER", "CodMaq2", 10, 6)
   Call CREACAMPO("ARTICULOS", "SECOPER", "CodMaq3", 10, 6)
   Call CREACAMPO("ARTICULOS", "SECOPER", "CodMaq4", 10, 6)
   Call CREACAMPO("ARTICULOS", "SECOPER", "CodMaq5", 10, 6)
   Call CREACAMPO("ARTICULOS", "SECOPER", "HorsPrep", 6, 0)
   Call CREACAMPO("ARTICULOS", "SECOPER", "HorsFijo", 6, 0)
   Call CREACAMPO("ARTICULOS", "SECOPER", "MinuStan", 6, 0)
   Call CREACAMPO("ARTICULOS", "SECOPER", "HorsMHer", 6, 0)
   Call CREACAMPO("ARTICULOS", "SECOPER", "HorsLimp", 6, 0)
   Call CREACAMPO("ARTICULOS", "SECOPER", "PieCiclo", 6, 0)
   Call CREACAMPO("ARTICULOS", "SECOPER", "PorDemor", 6, 0)
   Call CREACAMPO("ARTICULOS", "SECOPER", "Cadencia", 6, 0)
   Call CREACAMPO("ARTICULOS", "SECOPER", "CoefSolp", 6, 0)
   Call CREACAMPO("ARTICULOS", "SECOPER", "LoteStan", 6, 0)
   Call CREACAMPO("ARTICULOS", "SECOPER", "CantiOps", 6, 0)
   Call CREACAMPO("ARTICULOS", "SECOPER", "Cantiops2", 6, 0)
   Call CREACAMPO("ARTICULOS", "SECOPER", "CategOps", 3, 0)
   Call CREACAMPO("ARTICULOS", "SECOPER", "Categops2", 3, 0)
   Call CREACAMPO("ARTICULOS", "SECOPER", "ProgCNum", 10, 24)
   Call CREACAMPO("ARTICULOS", "SECOPER", "HoProFab", 10, 24)
   Call CREACAMPO("ARTICULOS", "SECOPER", "ObseOper", 12, 0)
   Call CREACAMPO("ARTICULOS", "SECOPER", "ProTerOp", 3, 0)
   Call CREACAMPO("ARTICULOS", "SECOPER", "ProvOper", 10, 12)
   Call CREACAMPO("ARTICULOS", "SECOPER", "PreUnid", 7, 0)
   Call CREACAMPO("ARTICULOS", "SECOPER", "MonePrec", 3, 0)
   Call CREACAMPO("ARTICULOS", "SECOPER", "FechPrec", 8, 0)
   Call CREACAMPO("ARTICULOS", "SECOPER", "Nuoper_Prep", 6, 4)
   Call CREACAMPO("ARTICULOS", "SECOPER", "Catoper_Prep", 3, 2)
   Call CREACAMPO("ARTICULOS", "SECOPER", "Fech_Revis", 8, 0)
   Call CREACAMPO("ARTICULOS", "SECOPER", "Usua_Revis", 10, 32)
   '
   Call CREACAMPO("ARTICULOS", "TMUSOPER", "Cod_Inte", 3, 0, 1)
   Call CREACAMPO("ARTICULOS", "TMUSOPER", "Cod_Oper", 3, 0)
   Call CREACAMPO("ARTICULOS", "TMUSOPER", "Cod_TMU", 10, 4)
   Call CREACAMPO("ARTICULOS", "TMUSOPER", "Nume_Ord", 3, 2)
   Call CREACAMPO("ARTICULOS", "TMUSOPER", "Cantidad", 7, 8)
   '
   'TABLA DE NOVEDADES DE CAMBIOS DE PRECIOS
   Call CREACAMPO("", "NOVEDADES", "Codint", 3, 0, 1)
   Call CREACAMPO("", "NOVEDADES", "Nume_Usuar", 3, 0)
   Call CREACAMPO("", "NOVEDADES", "Fech_Mod", 8, 0)
   Call CREACAMPO("", "NOVEDADES", "Prec_Ant", 7, 0)
   Call CREACAMPO("", "NOVEDADES", "Prec_New", 7, 0)
   Call CREACAMPO("", "NOVEDADES", "NumProv", 4, 0, 1)
   Call CREACAMPO("", "NOVEDADES", "Observacion", 10, 512)
   Call CREACAMPO("", "NOVEDADES", "Moneda", 3, 2, 1)
   Call CREACAMPO("", "NOVEDADES", "Lista_Repos", 10, 12)
   Call CREACAMPO("", "NOVEDADES", "COTIZA_DOLAR", 7, 0)
   
   'TABLA DE IMPRESION TEMPORARIA DE NOVEDADES PARA LOS REPORTES
   Call CREACAMPO("", "PRINTNOVEDADES", "Codint", 3, 0, 1)
   Call CREACAMPO("", "PRINTNOVEDADES", "Porcentaje", 7, 0)
   Call CREACAMPO("", "PRINTNOVEDADES", "NOM_USER", 10, 24)
   Call CREACAMPO("", "PRINTNOVEDADES", "Fech_Mod1", 8, 0)
   Call CREACAMPO("", "PRINTNOVEDADES", "Fech_Mod2", 8, 0)
   Call CREACAMPO("", "PRINTNOVEDADES", "Prec_Ant", 7, 0)
   Call CREACAMPO("", "PRINTNOVEDADES", "Prec_New", 7, 0)
   Call CREACAMPO("", "PRINTNOVEDADES", "NumProv", 4, 0, 1)
   Call CREACAMPO("", "PRINTNOVEDADES", "Observacion", 10, 512)
   Call CREACAMPO("", "PRINTNOVEDADES", "Moneda", 3, 2, 1)
   Call CREACAMPO("", "PRINTNOVEDADES", "IMP_ORG_DOLAR", 7, 0)
   Call CREACAMPO("", "PRINTNOVEDADES", "COTIZA_DOLAR", 7, 0)
   Call CREACAMPO("", "PRINTNOVEDADES", "Lista_Repos", 10, 12)
   '
   'TABLA DE IMPRESION TEMPORARIA DE NOVEDADES PARA LOS REPORTES
   Call CREACAMPO("", "HISTOCOSTOS", "Periodo", 10, 16, 1)
   Call CREACAMPO("", "HISTOCOSTOS", "IdPeriodo", 4, 0, 1)
   Call CREACAMPO("", "HISTOCOSTOS", "FeCosCal", 8, 0)
   Call CREACAMPO("", "HISTOCOSTOS", "Cod_Inte", 3, 0, 1)
   Call CREACAMPO("", "HISTOCOSTOS", "Timate", 3, 0)
   Call CREACAMPO("", "HISTOCOSTOS", "Umedida", 10, 2)
   Call CREACAMPO("", "HISTOCOSTOS", "MoneCalCo", 3, 0)
   Call CREACAMPO("", "HISTOCOSTOS", "CostoMP", 7, 0)
   Call CREACAMPO("", "HISTOCOSTOS", "Piezas", 7, 0)
   Call CREACAMPO("", "HISTOCOSTOS", "TTerceros", 7, 0)
   Call CREACAMPO("", "HISTOCOSTOS", "TotaSuminis", 7, 0)
   Call CREACAMPO("", "HISTOCOSTOS", "ManoObra", 7, 0)
   Call CREACAMPO("", "HISTOCOSTOS", "Equipos", 7, 0)
   Call CREACAMPO("", "HISTOCOSTOS", "Otros", 7, 0)
   Call CREACAMPO("", "HISTOCOSTOS", "TotElaboracion", 7, 0)
   Call CREACAMPO("", "HISTOCOSTOS", "TotCostoDirecto", 7, 0)
   Call CREACAMPO("", "HISTOCOSTOS", "CostoStandard", 7, 0)
   Call CREACAMPO("", "HISTOCOSTOS", "CostoStdTotal", 7, 0)
   Call CREACAMPO("", "HISTOCOSTOS", "PrecioVenta", 7, 0)
   Call CREACAMPO("", "HISTOCOSTOS", "PrecioLista", 7, 0)
   Call CREACAMPO("", "HISTOCOSTOS", "Margen", 7, 0)
   Call CREACAMPO("", "HISTOCOSTOS", "PorcentajeMBruto", 7, 0)
   Call CREACAMPO("", "HISTOCOSTOS", "IndiceCalidad", 7, 0)
   '
End Sub
   
'
Sub CREATABLA_BADACLIEN()
   '
   Call CREACAMPO("BADACLIEN", "BSEGCOM", "Codi_Cli", 10, 12, 1)
   Call CREACAMPO("BADACLIEN", "BSEGCOM", "Nume_Cli", 4, 0, 1)
   Call CREACAMPO("BADACLIEN", "BSEGCOM", "Raso_Cli", 10, 64)
   Call CREACAMPO("BADACLIEN", "BSEGCOM", "Mail_Cli", 10, 64)
   Call CREACAMPO("BADACLIEN", "BSEGCOM", "Domi_Cli", 10, 64)
   Call CREACAMPO("BADACLIEN", "BSEGCOM", "Cpos_Cli", 10, 16)
   Call CREACAMPO("BADACLIEN", "BSEGCOM", "Loca_Cli", 10, 48)
   Call CREACAMPO("BADACLIEN", "BSEGCOM", "Prov_Cli", 10, 2)
   Call CREACAMPO("BADACLIEN", "BSEGCOM", "Pais_Cli", 10, 48)
   Call CREACAMPO("BADACLIEN", "BSEGCOM", "Origen", 10, 64)
   Call CREACAMPO("BADACLIEN", "BSEGCOM", "Vend_Cli", 3, 0)
   Call CREACAMPO("BADACLIEN", "BSEGCOM", "RamoActi", 10, 64)
   Call CREACAMPO("BADACLIEN", "BSEGCOM", "ProduPrin", 10, 128)
   Call CREACAMPO("BADACLIEN", "BSEGCOM", "Empleados", 4, 0)
   Call CREACAMPO("BADACLIEN", "BSEGCOM", "Tele_Cli", 10, 32)
   Call CREACAMPO("BADACLIEN", "BSEGCOM", "Tele_Cli2", 10, 32)
   Call CREACAMPO("BADACLIEN", "BSEGCOM", "Fax_Cli", 10, 24)
   Call CREACAMPO("BADACLIEN", "BSEGCOM", "Ctac_Cli", 10, 48)
   Call CREACAMPO("BADACLIEN", "BSEGCOM", "Cargo1", 10, 32)
   Call CREACAMPO("BADACLIEN", "BSEGCOM", "Sector1", 10, 64)
   Call CREACAMPO("BADACLIEN", "BSEGCOM", "Interno1", 10, 32)
   Call CREACAMPO("BADACLIEN", "BSEGCOM", "Contacto2", 10, 48)
   Call CREACAMPO("BADACLIEN", "BSEGCOM", "Cargo2", 10, 32)
   Call CREACAMPO("BADACLIEN", "BSEGCOM", "Sector2", 10, 64)
   Call CREACAMPO("BADACLIEN", "BSEGCOM", "Interno2", 10, 32)
   Call CREACAMPO("BADACLIEN", "BSEGCOM", "Contacto3", 10, 48)
   Call CREACAMPO("BADACLIEN", "BSEGCOM", "Cargo3", 10, 32)
   Call CREACAMPO("BADACLIEN", "BSEGCOM", "Sector3", 10, 64)
   Call CREACAMPO("BADACLIEN", "BSEGCOM", "Interno3", 10, 32)
   Call CREACAMPO("BADACLIEN", "BSEGCOM", "PagWeb", 10, 48)
   Call CREACAMPO("BADACLIEN", "BSEGCOM", "FechBaja", 10, 8)
   Call CREACAMPO("BADACLIEN", "BSEGCOM", "Mot_Baja", 10, 64)
   Call CREACAMPO("BADACLIEN", "BSEGCOM", "FechUltEnvio", 8, 0)
   Call CREACAMPO("BADACLIEN", "BSEGCOM", "Grupos", 10, 30)
   Call CREACAMPO("BADACLIEN", "BSEGCOM", "FechProxCto", 8, 0)
   Call CREACAMPO("BADACLIEN", "BSEGCOM", "Nota_Cli", 12, 0)
   Call CREACAMPO("BADACLIEN", "BSEGCOM", "His_Contac", 12, 0)
   '
   Call CREACAMPO("BADACLIEN", "Deudor12", "Lentre0_Cli", 10, 48)
   Call CREACAMPO("BADACLIEN", "Deudor12", "Lentre1_Cli", 10, 48)
   Call CREACAMPO("BADACLIEN", "Deudor12", "NTrans_Cli", 3, 0)
   Call CREACAMPO("BADACLIEN", "Deudor12", "Lentre2_Cli", 10, 48)
   Call CREACAMPO("BADACLIEN", "Deudor12", "Hora_Ent", 10, 48)
   Call CREACAMPO("BADACLIEN", "Deudor12", "Zona_Cli", 10, 32)
   Call CREACAMPO("BADACLIEN", "Deudor12", "Zona_Rep", 10, 8)
   Call CREACAMPO("BADACLIEN", "Deudor12", "Gru_Div", 4, 0)
   Call CREACAMPO("BADACLIEN", "Deudor12", "Nume_Cli", 4, 0, 1)
   Call CREACAMPO("BADACLIEN", "Deudor12", "Codi_Cli", 10, 12, 1)
   Call CREACAMPO("BADACLIEN", "Deudor12", "Raso_Cli", 10, 64, 1)
   Call CREACAMPO("BADACLIEN", "Deudor12", "Fant_Cli", 10, 32)
   Call CREACAMPO("BADACLIEN", "Deudor12", "Stat_Cli", 3, 0, 1)
   Call CREACAMPO("BADACLIEN", "Deudor12", "Domi_Cli", 10, 64)
   Call CREACAMPO("BADACLIEN", "Deudor12", "Cpos_Cli", 10, 16)
   Call CREACAMPO("BADACLIEN", "Deudor12", "Loca_Cli", 10, 48)
   Call CREACAMPO("BADACLIEN", "Deudor12", "Pais_Cli", 10, 48)
   Call CREACAMPO("BADACLIEN", "Deudor12", "Prov_Cli", 10, 2)
   Call CREACAMPO("BADACLIEN", "Deudor12", "Tele_Cli", 10, 32)
   Call CREACAMPO("BADACLIEN", "Deudor12", "Faxi_Cli", 10, 32)
   Call CREACAMPO("BADACLIEN", "Deudor12", "Mail_Cli", 10, 64)
   Call CREACAMPO("BADACLIEN", "Deudor12", "PagWeb", 10, 128)
   Call CREACAMPO("BADACLIEN", "Deudor12", "Piva_Cli", 3, 0)
   Call CREACAMPO("BADACLIEN", "Deudor12", "Cuit_Cli", 10, 24, 1)
   Call CREACAMPO("BADACLIEN", "Deudor12", "Nibr_Cli", 10, 24)
   Call CREACAMPO("BADACLIEN", "Deudor12", "Prvd_Cli", 10, 24)
   Call CREACAMPO("BADACLIEN", "Deudor12", "Ctac_Cli", 10, 48)
   Call CREACAMPO("BADACLIEN", "Deudor12", "Hora_Cli", 10, 32)
   Call CREACAMPO("BADACLIEN", "Deudor12", "Cpag_Cli", 3, 0)
   Call CREACAMPO("BADACLIEN", "Deudor12", "Lpre_Cli", 3, 0)
   Call CREACAMPO("BADACLIEN", "Deudor12", "Vend_Cli", 3, 0)
   Call CREACAMPO("BADACLIEN", "Deudor12", "Tcpr_Cli", 3, 0, 1)
   Call CREACAMPO("BADACLIEN", "Deudor12", "CaIb_Cli", 10, 2)
   Call CREACAMPO("BADACLIEN", "Deudor12", "Dsta_Cli", 10, 32)
   Call CREACAMPO("BADACLIEN", "Deudor12", "Lcre_Cli", 6, 0)
   Call CREACAMPO("BADACLIEN", "Deudor12", "Sald_Cli", 7, 0)
   Call CREACAMPO("BADACLIEN", "Deudor12", "CuMa_Cli", 3, 0)
   Call CREACAMPO("BADACLIEN", "Deudor12", "GrCo_Cli", 3, 2)
   Call CREACAMPO("BADACLIEN", "Deudor12", "Nota_Cli", 12, 0)
   Call CREACAMPO("BADACLIEN", "Deudor12", "Obser_Entre", 12, 0)
   Call CREACAMPO("BADACLIEN", "Deudor12", "Recar_Embal", 6, 0)
   Call CREACAMPO("BADACLIEN", "Deudor12", "Mar_Ya_Fac", 3, 2)
   Call CREACAMPO("BADACLIEN", "Deudor12", "DePosIva", 10, 32)
   Call CREACAMPO("BADACLIEN", "Deudor12", "LCred_CCte", 6, 0)
   Call CREACAMPO("BADACLIEN", "Deudor12", "EnvMaiDeuda", 3, 0)
   Call CREACAMPO("BADACLIEN", "Deudor12", "Ctact_Cob", 10, 48)
   Call CREACAMPO("BADACLIEN", "Deudor12", "Tele_Cob", 10, 48)
   Call CREACAMPO("BADACLIEN", "Deudor12", "Mail_Cob", 10, 64)
   Call CREACAMPO("BADACLIEN", "Deudor12", "DiasRec_Cob", 10, 64)
   Call CREACAMPO("BADACLIEN", "Deudor12", "DiasPago_Cob", 10, 64)
   Call CREACAMPO("BADACLIEN", "Deudor12", "Segui_Cob", 12, 0)
   Call CREACAMPO("BADACLIEN", "Deudor12", "MailFactura_Cob", 10, 64)
   '
   Call CREACAMPO("BADACLIEN", "SUCENTRE", "Nume_Cli", 4, 0)
   Call CREACAMPO("BADACLIEN", "SUCENTRE", "Nume_Suc", 3, 0)
   Call CREACAMPO("BADACLIEN", "SUCENTRE", "Deno_Suc", 10, 128)
   Call CREACAMPO("BADACLIEN", "SUCENTRE", "Domi_Suc", 10, 128)
   Call CREACAMPO("BADACLIEN", "SUCENTRE", "Loca_Suc", 10, 128)
   Call CREACAMPO("BADACLIEN", "SUCENTRE", "Tele_Suc", 10, 64)
   Call CREACAMPO("BADACLIEN", "SUCENTRE", "Codi_Tra", 3, 0)
   '
   Call CREACAMPO("BADACLIEN", "TRANSPORTES", "Codi_Tra", 3, 0)
   Call CREACAMPO("BADACLIEN", "TRANSPORTES", "Deno_Tra", 10, 128)
   Call CREACAMPO("BADACLIEN", "TRANSPORTES", "Raso_Tra", 10, 128)
   Call CREACAMPO("BADACLIEN", "TRANSPORTES", "Domi_Tra", 10, 128)
   Call CREACAMPO("BADACLIEN", "TRANSPORTES", "Loca_Tra", 10, 128)
   Call CREACAMPO("BADACLIEN", "TRANSPORTES", "Cuit_Tra", 10, 32)
   Call CREACAMPO("BADACLIEN", "TRANSPORTES", "Piva_Tra", 3, 0)
   Call CREACAMPO("BADACLIEN", "TRANSPORTES", "Tele_Tra", 10, 64)
   Call CREACAMPO("BADACLIEN", "TRANSPORTES", "Faxi_Tra", 10, 64)
   Call CREACAMPO("BADACLIEN", "TRANSPORTES", "Mail_Tra", 10, 128)
   Call CREACAMPO("BADACLIEN", "TRANSPORTES", "Ctac_Tra", 10, 128)
   '
   Call CREACAMPO("BADACLIEN", "VENDEDOR", "Vend_Cli", 3, 0)
   Call CREACAMPO("BADACLIEN", "VENDEDOR", "RasoVend", 10, 48)
   Call CREACAMPO("BADACLIEN", "VENDEDOR", "DomiVend", 10, 48)
   Call CREACAMPO("BADACLIEN", "VENDEDOR", "LocaVend", 10, 48)
   Call CREACAMPO("BADACLIEN", "VENDEDOR", "TeleVend_1", 10, 32)
   Call CREACAMPO("BADACLIEN", "VENDEDOR", "TeleVend_2", 10, 32)
   Call CREACAMPO("BADACLIEN", "VENDEDOR", "EmailVend", 10, 48)
   Call CREACAMPO("BADACLIEN", "VENDEDOR", "Observaciones", 12, 0)
   '
End Sub
'
Sub CREATABLA_BADAPROVE()
   '
   Call CREACAMPO("BADAPROVE", "Proved12", "Lcre_Cli", 6, 0)
   Call CREACAMPO("BADAPROVE", "Proved12", "Sald_Cli", 7, 0)
   Call CREACAMPO("BADAPROVE", "Proved12", "CuMa_Cli", 3, 0)
   Call CREACAMPO("BADAPROVE", "Proved12", "GrCo_Cli", 3, 2)
   Call CREACAMPO("BADAPROVE", "Proved12", "Nota_Cli", 12, 0)
   Call CREACAMPO("BADAPROVE", "Proved12", "C_Ret", 10, 10)
   Call CREACAMPO("BADAPROVE", "Proved12", "Posivat_Cli", 10, 10)
   Call CREACAMPO("BADAPROVE", "Proved12", "Nume_Cli", 4, 0, 1)
   Call CREACAMPO("BADAPROVE", "Proved12", "Codi_Cli", 10, 12, 1)
   Call CREACAMPO("BADAPROVE", "Proved12", "Raso_Cli", 10, 64)
   Call CREACAMPO("BADAPROVE", "Proved12", "Fant_Cli", 10, 32)
   Call CREACAMPO("BADAPROVE", "Proved12", "Stat_Cli", 3, 0)
   Call CREACAMPO("BADAPROVE", "Proved12", "Domi_Cli", 10, 64)
   Call CREACAMPO("BADAPROVE", "Proved12", "Cpos_Cli", 10, 16)
   Call CREACAMPO("BADAPROVE", "Proved12", "Loca_Cli", 10, 48)
   Call CREACAMPO("BADAPROVE", "Proved12", "Pais_Cli", 10, 48)
   Call CREACAMPO("BADAPROVE", "Proved12", "Prov_Cli", 10, 2)
   Call CREACAMPO("BADAPROVE", "Proved12", "Tele_Cli", 10, 32)
   Call CREACAMPO("BADAPROVE", "Proved12", "Tele2_Cli", 10, 32)
   Call CREACAMPO("BADAPROVE", "Proved12", "Faxi_Cli", 10, 32)
   Call CREACAMPO("BADAPROVE", "Proved12", "Mail_Cli", 10, 64)
   Call CREACAMPO("BADAPROVE", "Proved12", "Piva_Cli", 3, 0)
   Call CREACAMPO("BADAPROVE", "Proved12", "Cuit_Cli", 10, 24)
   Call CREACAMPO("BADAPROVE", "Proved12", "Nibr_Cli", 10, 24)
   Call CREACAMPO("BADAPROVE", "Proved12", "Prvd_Cli", 10, 24)
   Call CREACAMPO("BADAPROVE", "Proved12", "Ctac_Cli", 10, 48)
   Call CREACAMPO("BADAPROVE", "Proved12", "Hora_Cli", 10, 32)
   Call CREACAMPO("BADAPROVE", "Proved12", "Ord_Cheq", 10, 48)
   Call CREACAMPO("BADAPROVE", "Proved12", "Cpag_Cli", 3, 0)
   Call CREACAMPO("BADAPROVE", "Proved12", "CCom_Cli", 12, 0)
   Call CREACAMPO("BADAPROVE", "Proved12", "Lpre_Cli", 3, 0)
   Call CREACAMPO("BADAPROVE", "Proved12", "Vend_Cli", 3, 0)
   Call CREACAMPO("BADAPROVE", "Proved12", "Tcpr_Cli", 3, 0, 1)
   Call CREACAMPO("BADAPROVE", "Proved12", "CaIb_Cli", 10, 2)
   Call CREACAMPO("BADAPROVE", "Proved12", "Dsta_Cli", 10, 32)
   Call CREACAMPO("BADAPROVE", "Proved12", "Calinteg", 6, 0)
   Call CREACAMPO("BADAPROVE", "Proved12", "ObsCompras", 12, 0)
   Call CREACAMPO("BADAPROVE", "Proved12", "LuEntCompras", 12, 0)
   '
End Sub
'
Sub CREATABLA_BCONTROL()
   '
   Call CREACAMPO("BCONTROL", "AUDITRA", "FeReTra", 8, 8)
   Call CREACAMPO("BCONTROL", "AUDITRA", "TipoTra", 10, 8)
   Call CREACAMPO("BCONTROL", "AUDITRA", "DescriTra", 10, 64)
   Call CREACAMPO("BCONTROL", "AUDITRA", "NuCompro", 10, 24)
   Call CREACAMPO("BCONTROL", "AUDITRA", "UserTra", 10, 16)
   Call CREACAMPO("BCONTROL", "AUDITRA", "IdTermi", 10, 24)
   '
   Call CREACAMPO("BCONTROL", "BloArchA", "IdenTabla", 10, 32)
   '
   Call CREACAMPO("BCONTROL", "BloArchB", "CtrlName", 10, 32)
   Call CREACAMPO("BCONTROL", "BloArchB", "ExecName", 10, 32)
   Call CREACAMPO("BCONTROL", "BloArchB", "TermName", 10, 32)
   Call CREACAMPO("BCONTROL", "BloArchB", "IdenTabla", 10, 32)
   Call CREACAMPO("BCONTROL", "BloArchB", "InitDate", 8, 0)
   Call CREACAMPO("BCONTROL", "BloArchB", "UserName", 10, 32)
   '
   Call CREACAMPO("BCONTROL", "MENUES", "CodPadre", 10, 32)
   Call CREACAMPO("BCONTROL", "MENUES", "Relacion", 4, 4)
   Call CREACAMPO("BCONTROL", "MENUES", "CodHijo", 10, 128)
   Call CREACAMPO("BCONTROL", "MENUES", "Descripcion", 10, 64)
   '
   Call CREACAMPO("BCONTROL", "SQL_LISELEC", "NomLiSele", 10, 32)
   Call CREACAMPO("BCONTROL", "SQL_LISELEC", "ACol_1", 3, 2)
   Call CREACAMPO("BCONTROL", "SQL_LISELEC", "ACol_2", 3, 2)
   Call CREACAMPO("BCONTROL", "SQL_LISELEC", "Tabla_Origen", 10, 256)
   Call CREACAMPO("BCONTROL", "SQL_LISELEC", "Lista_Campos", 12, 0)
   Call CREACAMPO("BCONTROL", "SQL_LISELEC", "Condi_Selec", 12, 0)
   Call CREACAMPO("BCONTROL", "SQL_LISELEC", "Fech_Actu", 8, 0)
   Call CREACAMPO("BCONTROL", "SQL_LISELEC", "LiSelec", 12, 0)
   '
   Call CREACAMPO("BCONTROL", "SQL_QUERIES", "Cod_Query", 10, 16)
   Call CREACAMPO("BCONTROL", "SQL_QUERIES", "Des_Query", 10, 256)
   Call CREACAMPO("BCONTROL", "SQL_QUERIES", "PAR_Query", 12, 0)
   Call CREACAMPO("BCONTROL", "SQL_QUERIES", "SQL_Query", 12, 0)
   '
   Call CREACAMPO("", "AUDITACT", "AuTabla", 10, 64)
   Call CREACAMPO("", "AUDITACT", "AuCampo", 10, 64)
   Call CREACAMPO("", "AUDITACT", "AuCondi", 12, 0)
   Call CREACAMPO("", "AUDITACT", "AuValAnt", 11, 0)
   Call CREACAMPO("", "AUDITACT", "AuNueVa", 11, 0)
   Call CREACAMPO("", "AUDITACT", "AuFeHor", 8, 0)
   Call CREACAMPO("", "AUDITACT", "AuUser", 3, 0, 1)
   Call CREACAMPO("", "AUDITACT", "AuUsuario", 10, 64)
   Call CREACAMPO("", "AUDITACT", "AuAplica", 10, 32)
   Call CREACAMPO("", "AUDITACT", "AuTerminal", 10, 64)

End Sub
'
Sub CREATABLA_COMPRAS()
   '
   Call CREACAMPO("COMPRAS", "AUTOENT", "Nume_Ocom", 4, 4)
   Call CREACAMPO("COMPRAS", "AUTOENT", "Nume_Aten", 4, 4)
   Call CREACAMPO("COMPRAS", "AUTOENT", "Femi_Aten", 8, 0)
   Call CREACAMPO("COMPRAS", "AUTOENT", "Npro_Aten", 3, 2)
   Call CREACAMPO("COMPRAS", "AUTOENT", "Raso_Aten", 10, 64)
   Call CREACAMPO("COMPRAS", "AUTOENT", "Cod_Inte", 3, 0)
   Call CREACAMPO("COMPRAS", "AUTOENT", "Cod_Exte", 10, 16)
   Call CREACAMPO("COMPRAS", "AUTOENT", "Descritem", 10, 64)
   Call CREACAMPO("COMPRAS", "AUTOENT", "Fentre_Sol", 8, 0)
   Call CREACAMPO("COMPRAS", "AUTOENT", "Cant_Aten", 6, 4)
   Call CREACAMPO("COMPRAS", "AUTOENT", "Acuen_Aten", 6, 4)
   Call CREACAMPO("COMPRAS", "AUTOENT", "Sald_Aten", 6, 4)
   Call CREACAMPO("COMPRAS", "AUTOENT", "NuOrd_Item", 3, 2)
   Call CREACAMPO("COMPRAS", "AUTOENT", "Stat_Aten", 3, 2)
   Call CREACAMPO("COMPRAS", "AUTOENT", "Observa_Aten", 10, 64)
   '
   Call CREACAMPO("COMPRAS", "COTIRECI", "Nume_Solint", 4, 4)
   Call CREACAMPO("COMPRAS", "COTIRECI", "NuOrd_item", 3, 2)
   Call CREACAMPO("COMPRAS", "COTIRECI", "Npro_Ocom", 4, 4)
   Call CREACAMPO("COMPRAS", "COTIRECI", "Cod_Inte", 3, 2)
   Call CREACAMPO("COMPRAS", "COTIRECI", "Cod_Exte", 10, 24)
   Call CREACAMPO("COMPRAS", "COTIRECI", "Descritem", 12, 0)
   Call CREACAMPO("COMPRAS", "COTIRECI", "Cant_Solint", 6, 4)
   Call CREACAMPO("COMPRAS", "COTIRECI", "Uni_Sol", 10, 4)
   Call CREACAMPO("COMPRAS", "COTIRECI", "Fech_Coti", 8, 0)
   Call CREACAMPO("COMPRAS", "COTIRECI", "Mone_Emis", 3, 4)
   Call CREACAMPO("COMPRAS", "COTIRECI", "Pre_Unit", 7, 8)
   Call CREACAMPO("COMPRAS", "COTIRECI", "INet_Item", 7, 8)
   Call CREACAMPO("COMPRAS", "COTIRECI", "Plaz_Entre", 3, 2)
   Call CREACAMPO("COMPRAS", "COTIRECI", "Venci_Ofer", 8, 0)
   Call CREACAMPO("COMPRAS", "COTIRECI", "Observa", 12, 0)
   Call CREACAMPO("COMPRAS", "COTIRECI", "Nume_Ocom", 4, 4)
   Call CREACAMPO("COMPRAS", "COTIRECI", "CondPag", 3, 0)
   Call CREACAMPO("COMPRAS", "COTIRECI", "Sect_Solint", 10, 12)
   Call CREACAMPO("COMPRAS", "COTIRECI", "Pre_Unit_Pes", 7, 8)
   '
   Call CREACAMPO("COMPRAS", "DETAIVA", "Nume_Ocom", 4, 4)
   Call CREACAMPO("COMPRAS", "DETAIVA", "Norit", 3, 2)
   Call CREACAMPO("COMPRAS", "DETAIVA", "Tasa_Iva", 10, 16)
   Call CREACAMPO("COMPRAS", "DETAIVA", "Base_Impo", 7, 8)
   Call CREACAMPO("COMPRAS", "DETAIVA", "Impo_Iva", 7, 8)
   '
   Call CREACAMPO("COMPRAS", "DETASOLINT", "Nume_Solint", 4, 4)
   Call CREACAMPO("COMPRAS", "DETASOLINT", "NuOrd_Item", 3, 2)
   Call CREACAMPO("COMPRAS", "DETASOLINT", "Cod_Inte", 3, 2)
   Call CREACAMPO("COMPRAS", "DETASOLINT", "Descritem", 12, 0)
   Call CREACAMPO("COMPRAS", "DETASOLINT", "Uni_Mas", 10, 4)
   Call CREACAMPO("COMPRAS", "DETASOLINT", "Uni_Sol", 10, 4)
   Call CREACAMPO("COMPRAS", "DETASOLINT", "Cant_Solint", 6, 4)
   Call CREACAMPO("COMPRAS", "DETASOLINT", "Fentre_Sol", 8, 0)
   Call CREACAMPO("COMPRAS", "DETASOLINT", "Coti_Solint", 4, 4)
   Call CREACAMPO("COMPRAS", "DETASOLINT", "Desti_Solint", 10, 255)
   Call CREACAMPO("COMPRAS", "DETASOLINT", "Iconta_Solint", 10, 32)
   Call CREACAMPO("COMPRAS", "DETASOLINT", "Status", 3, 2)
   '
   Call CREACAMPO("COMPRAS", "DETOCABI", "Nume_Ocom", 4, 4)
   Call CREACAMPO("COMPRAS", "DETOCABI", "NuOrd_Item", 3, 2)
   Call CREACAMPO("COMPRAS", "DETOCABI", "Femi_Ocom", 8, 0)
   Call CREACAMPO("COMPRAS", "DETOCABI", "Npro_Ocom", 3, 2)
   Call CREACAMPO("COMPRAS", "DETOCABI", "RaSo_Prov", 10, 64)
   Call CREACAMPO("COMPRAS", "DETOCABI", "Cod_Inte", 3, 0)
   Call CREACAMPO("COMPRAS", "DETOCABI", "Cod_Exte", 10, 16)
   Call CREACAMPO("COMPRAS", "DETOCABI", "Descritem", 10, 64)
   Call CREACAMPO("COMPRAS", "DETOCABI", "Delibre", 12, 0)
   Call CREACAMPO("COMPRAS", "DETOCABI", "IContab", 3, 2)
   Call CREACAMPO("COMPRAS", "DETOCABI", "Sect_Ocom", 10, 12)
   Call CREACAMPO("COMPRAS", "DETOCABI", "Uni_Mas", 10, 4)
   Call CREACAMPO("COMPRAS", "DETOCABI", "Colo_Ocom", 10, 16)
   Call CREACAMPO("COMPRAS", "DETOCABI", "Medi_Ocom", 10, 20)
   Call CREACAMPO("COMPRAS", "DETOCABI", "Valo_Unid", 10, 12)
   Call CREACAMPO("COMPRAS", "DETOCABI", "Unids_Ocom", 6, 4)
   Call CREACAMPO("COMPRAS", "DETOCABI", "Cant_Ocom", 6, 4)
   Call CREACAMPO("COMPRAS", "DETOCABI", "Uni_Com", 10, 4)
   Call CREACAMPO("COMPRAS", "DETOCABI", "Coef_Unids", 6, 4)
   Call CREACAMPO("COMPRAS", "DETOCABI", "Pre_Unit", 7, 8)
   Call CREACAMPO("COMPRAS", "DETOCABI", "Desc_Item", 6, 4)
   Call CREACAMPO("COMPRAS", "DETOCABI", "PrUn_Neto", 7, 8)
   Call CREACAMPO("COMPRAS", "DETOCABI", "INet_Item", 7, 8)
   Call CREACAMPO("COMPRAS", "DETOCABI", "FVig_Ocom", 8, 0)
   Call CREACAMPO("COMPRAS", "DETOCABI", "Consimat", 10, 2)
   Call CREACAMPO("COMPRAS", "DETOCABI", "Cant_Rec", 6, 4)
   Call CREACAMPO("COMPRAS", "DETOCABI", "Cod_Oper", 3, 2)
   Call CREACAMPO("COMPRAS", "DETOCABI", "Doc_Orig", 10, 12)
   Call CREACAMPO("COMPRAS", "DETOCABI", "N_OrServ", 6, 4)
   Call CREACAMPO("COMPRAS", "DETOCABI", "Mone_Cos", 3, 2)
   Call CREACAMPO("COMPRAS", "DETOCABI", "Mone_Emis", 3, 2)
   Call CREACAMPO("COMPRAS", "DETOCABI", "TSum_Ocom", 3, 2)
   Call CREACAMPO("COMPRAS", "DETOCABI", "Stat_Ocom", 3, 2)
   Call CREACAMPO("COMPRAS", "DETOCABI", "Nro_Otrab", 10, 16)
   Call CREACAMPO("COMPRAS", "DETOCABI", "Nro_Grup", 10, 10)
   '
   Call CREACAMPO("COMPRAS", "ENCASOLINT", "Nume_Solint", 4, 4)
   Call CREACAMPO("COMPRAS", "ENCASOLINT", "Femi_Solint", 8, 0)
   Call CREACAMPO("COMPRAS", "ENCASOLINT", "Usua_Gen", 3, 2)
   Call CREACAMPO("COMPRAS", "ENCASOLINT", "Auto_1", 3, 2)
   Call CREACAMPO("COMPRAS", "ENCASOLINT", "Auto_2", 3, 2)
   Call CREACAMPO("COMPRAS", "ENCASOLINT", "Auto_3", 3, 2)
   Call CREACAMPO("COMPRAS", "ENCASOLINT", "Auto_4", 3, 2)
   Call CREACAMPO("COMPRAS", "ENCASOLINT", "Obse_Solint", 12, 0)
   Call CREACAMPO("COMPRAS", "ENCASOLINT", "Sect_Solint", 10, 12)
   Call CREACAMPO("COMPRAS", "ENCASOLINT", "Autoriza_1", 8, 0)
   Call CREACAMPO("COMPRAS", "ENCASOLINT", "Autoriza_2", 8, 0)
   Call CREACAMPO("COMPRAS", "ENCASOLINT", "Autoriza_3", 8, 0)
   Call CREACAMPO("COMPRAS", "ENCASOLINT", "Autoriza_4", 8, 0)
   Call CREACAMPO("COMPRAS", "ENCASOLINT", "AnotaSolint", 12, 0)
   Call CREACAMPO("COMPRAS", "ENCASOLINT", "HayNews", 3, 2)
   Call CREACAMPO("COMPRAS", "ENCASOLINT", "Descri_Solint", 10, 64)
   Call CREACAMPO("COMPRAS", "ENCASOLINT", "Status", 3, 2)
   '
   Call CREACAMPO("COMPRAS", "ENCOCABI", "FVig_Ocom", 8, 2)
   Call CREACAMPO("COMPRAS", "ENCOCABI", "Nume_Ocom", 4, 4)
   Call CREACAMPO("COMPRAS", "ENCOCABI", "Refe_Ocom", 10, 64)
   Call CREACAMPO("COMPRAS", "ENCOCABI", "Femi_Ocom", 8, 0)
   Call CREACAMPO("COMPRAS", "ENCOCABI", "Npro_Ocom", 3, 2)
   Call CREACAMPO("COMPRAS", "ENCOCABI", "RaSo_Prov", 10, 48)
   Call CREACAMPO("COMPRAS", "ENCOCABI", "TSum_Ocom", 3, 2)
   Call CREACAMPO("COMPRAS", "ENCOCABI", "Stat_Ocom", 3, 2)
   Call CREACAMPO("COMPRAS", "ENCOCABI", "FGen_Ocom", 8, 0)
   Call CREACAMPO("COMPRAS", "ENCOCABI", "Usua_Gen", 3, 2)
   Call CREACAMPO("COMPRAS", "ENCOCABI", "FApro_Ocom", 8, 0)
   Call CREACAMPO("COMPRAS", "ENCOCABI", "Usua_Apro", 3, 2)
   Call CREACAMPO("COMPRAS", "ENCOCABI", "FMod_Ocom", 8, 0)
   Call CREACAMPO("COMPRAS", "ENCOCABI", "Usua_Mod", 3, 2)
   Call CREACAMPO("COMPRAS", "ENCOCABI", "FAnu_Ocom", 8, 0)
   Call CREACAMPO("COMPRAS", "ENCOCABI", "Usua_Anul", 3, 2)
   Call CREACAMPO("COMPRAS", "ENCOCABI", "Moti_Anul", 10, 16)
   Call CREACAMPO("COMPRAS", "ENCOCABI", "Mone_Cos", 3, 2)
   Call CREACAMPO("COMPRAS", "ENCOCABI", "Mone_Emis", 3, 2)
   Call CREACAMPO("COMPRAS", "ENCOCABI", "Tcam_Ocom", 6, 4)
   Call CREACAMPO("COMPRAS", "ENCOCABI", "ToBru_Ocom", 7, 8)
   Call CREACAMPO("COMPRAS", "ENCOCABI", "Pdes_Ocom", 10, 24)
   Call CREACAMPO("COMPRAS", "ENCOCABI", "Ides_Ocom", 7, 8)
   Call CREACAMPO("COMPRAS", "ENCOCABI", "Iiva_Ocom", 7, 8)
   Call CREACAMPO("COMPRAS", "ENCOCABI", "Iotr1_Ocom", 7, 8)
   Call CREACAMPO("COMPRAS", "ENCOCABI", "Iotr2_Ocom", 7, 8)
   Call CREACAMPO("COMPRAS", "ENCOCABI", "Total_Ocom", 7, 8)
   Call CREACAMPO("COMPRAS", "ENCOCABI", "Cpag_Ocom", 3, 2)
   Call CREACAMPO("COMPRAS", "ENCOCABI", "Embala_Ocom", 3, 2)
   Call CREACAMPO("COMPRAS", "ENCOCABI", "AtSr_Ocom", 10, 64)
   Call CREACAMPO("COMPRAS", "ENCOCABI", "Obse_Ocom", 12, 0)
   Call CREACAMPO("COMPRAS", "ENCOCABI", "Lentrega", 10, 64)
   Call CREACAMPO("COMPRAS", "ENCOCABI", "Anex_Ocom", 10, 128)
   Call CREACAMPO("COMPRAS", "ENCOCABI", "OcomRevis", 3, 2)
   '
   Call CREACAMPO("COMPRAS", "MACONSIG", "Femi_Ocom", 8, 0)
   Call CREACAMPO("COMPRAS", "MACONSIG", "Npro_Ocom", 3, 2)
   Call CREACAMPO("COMPRAS", "MACONSIG", "RaSo_Prov", 10, 64)
   Call CREACAMPO("COMPRAS", "MACONSIG", "Secc_Ejec", 6, 4)
   Call CREACAMPO("COMPRAS", "MACONSIG", "Deno_Secc", 10, 64)
   Call CREACAMPO("COMPRAS", "MACONSIG", "Nro_Orfa", 10, 12)
   Call CREACAMPO("COMPRAS", "MACONSIG", "Nro_Ocom", 10, 12)
   Call CREACAMPO("COMPRAS", "MACONSIG", "Cod_Inte", 3, 0)
   Call CREACAMPO("COMPRAS", "MACONSIG", "Cod_Exte", 10, 16)
   Call CREACAMPO("COMPRAS", "MACONSIG", "Descritem", 10, 64)
   Call CREACAMPO("COMPRAS", "MACONSIG", "Lote_Ocom", 10, 16)
   Call CREACAMPO("COMPRAS", "MACONSIG", "Unidad", 10, 4)
   Call CREACAMPO("COMPRAS", "MACONSIG", "Cant_Asig", 7, 8)
   Call CREACAMPO("COMPRAS", "MACONSIG", "Cant_Entreg", 7, 8)
   Call CREACAMPO("COMPRAS", "MACONSIG", "Fech_Entreg", 8, 0)
   Call CREACAMPO("COMPRAS", "MACONSIG", "NRemi_Consig", 10, 16)
   Call CREACAMPO("COMPRAS", "MACONSIG", "Valo_Unit", 7, 8)
   Call CREACAMPO("COMPRAS", "MACONSIG", "Moneda", 3, 2)
   Call CREACAMPO("COMPRAS", "MACONSIG", "Valo_Tot", 7, 8)
   Call CREACAMPO("COMPRAS", "MACONSIG", "Cant_Rendid", 7, 8)
   Call CREACAMPO("COMPRAS", "MACONSIG", "Saldo_Pend", 7, 8)
   Call CREACAMPO("COMPRAS", "MACONSIG", "Observa", 12, 0)
   Call CREACAMPO("COMPRAS", "MACONSIG", "NuOrdItem", 3, 0)
   '
   Call CREACAMPO("COMPRAS", "OCOEMI", "Cod_Inte", 3, 2)
   Call CREACAMPO("COMPRAS", "OCOEMI", "Cantidad", 6, 4)
   Call CREACAMPO("COMPRAS", "OCOEMI", "Fentre_Sol", 8, 0)
   Call CREACAMPO("COMPRAS", "OCOEMI", "Npro_Ocom", 3, 2)
   Call CREACAMPO("COMPRAS", "OCOEMI", "Uni_Mas", 10, 4)
   Call CREACAMPO("COMPRAS", "OCOEMI", "RaSo_Prov", 10, 48)
   Call CREACAMPO("COMPRAS", "OCOEMI", "Cant_Ocom", 6, 4)
   Call CREACAMPO("COMPRAS", "OCOEMI", "FechRecep", 8, 0)
   Call CREACAMPO("COMPRAS", "OCOEMI", "Nume_Ocom", 4, 2)
   Call CREACAMPO("COMPRAS", "OCOEMI", "Periodo", 10, 24)
   Call CREACAMPO("COMPRAS", "OCOEMI", "Cosuni", 6, 0)
   Call CREACAMPO("COMPRAS", "OCOEMI", "StokAct", 6, 0)
   Call CREACAMPO("COMPRAS", "OCOEMI", "Stomin", 6, 0)
   Call CREACAMPO("COMPRAS", "OCOEMI", "Lorepos", 6, 0)
   '
   Call CREACAMPO("COMPRAS", "OCOMDETA", "Nume_Ocom", 4, 4, 1)
   Call CREACAMPO("COMPRAS", "OCOMDETA", "Nume_Pedi", 4, 0)
   Call CREACAMPO("COMPRAS", "OCOMDETA", "NuOrd_Item", 3, 2)
   Call CREACAMPO("COMPRAS", "OCOMDETA", "Femi_Ocom", 8, 0)
   Call CREACAMPO("COMPRAS", "OCOMDETA", "Npro_Ocom", 3, 2)
   Call CREACAMPO("COMPRAS", "OCOMDETA", "Nume_Cli", 4, 0)
   Call CREACAMPO("COMPRAS", "OCOMDETA", "RaSo_Prov", 10, 64)
   Call CREACAMPO("COMPRAS", "OCOMDETA", "Cod_Inte", 3, 0)
   Call CREACAMPO("COMPRAS", "OCOMDETA", "Cod_Exte", 10, 16)
   Call CREACAMPO("COMPRAS", "OCOMDETA", "Descritem", 10, 64)
   Call CREACAMPO("COMPRAS", "OCOMDETA", "Delibre", 12, 0)
   Call CREACAMPO("COMPRAS", "OCOMDETA", "IContab", 3, 2)
   Call CREACAMPO("COMPRAS", "OCOMDETA", "Sect_Ocom", 10, 12)
   Call CREACAMPO("COMPRAS", "OCOMDETA", "Uni_Mas", 10, 4)
   Call CREACAMPO("COMPRAS", "OCOMDETA", "Colo_Ocom", 10, 16)
   Call CREACAMPO("COMPRAS", "OCOMDETA", "Medi_Ocom", 10, 20)
   Call CREACAMPO("COMPRAS", "OCOMDETA", "Valo_Unid", 10, 12)
   Call CREACAMPO("COMPRAS", "OCOMDETA", "Unids_Ocom", 6, 4)
   Call CREACAMPO("COMPRAS", "OCOMDETA", "Cant_Ocom", 6, 4)
   Call CREACAMPO("COMPRAS", "OCOMDETA", "Uni_Com", 10, 4)
   Call CREACAMPO("COMPRAS", "OCOMDETA", "Coef_Unids", 6, 4)
   Call CREACAMPO("COMPRAS", "OCOMDETA", "Pre_Unit", 7, 8)
   Call CREACAMPO("COMPRAS", "OCOMDETA", "Desc_Item", 6, 4)
   Call CREACAMPO("COMPRAS", "OCOMDETA", "PrUn_Neto", 7, 8)
   Call CREACAMPO("COMPRAS", "OCOMDETA", "INet_Item", 7, 8)
   Call CREACAMPO("COMPRAS", "OCOMDETA", "Fentre_Sol", 8, 0)
   Call CREACAMPO("COMPRAS", "OCOMDETA", "Consimat", 10, 2)
   Call CREACAMPO("COMPRAS", "OCOMDETA", "Cant_Rec", 6, 4)
   Call CREACAMPO("COMPRAS", "OCOMDETA", "Cod_Oper", 3, 2)
   Call CREACAMPO("COMPRAS", "OCOMDETA", "Doc_Orig", 10, 12)
   Call CREACAMPO("COMPRAS", "OCOMDETA", "N_OrServ", 6, 4)
   Call CREACAMPO("COMPRAS", "OCOMDETA", "Mone_Cos", 3, 2)
   Call CREACAMPO("COMPRAS", "OCOMDETA", "Mone_Emis", 3, 2)
   Call CREACAMPO("COMPRAS", "OCOMDETA", "TSum_Ocom", 3, 2)
   Call CREACAMPO("COMPRAS", "OCOMDETA", "Stat_Ocom", 3, 2)
   Call CREACAMPO("COMPRAS", "OCOMDETA", "Nro_Otrab", 10, 16)
   Call CREACAMPO("COMPRAS", "OCOMDETA", "Nro_Grup", 10, 10)
   Call CREACAMPO("COMPRAS", "OCOMDETA", "IdenReque", 10, 24)
   Call CREACAMPO("COMPRAS", "OCOMDETA", "PrUn_Neto_Pes", 7, 8)
   '
   Call CREACAMPO("COMPRAS", "OCOMENCA", "AtSr_Ocom", 10, 64)
   Call CREACAMPO("COMPRAS", "OCOMENCA", "Obse_Ocom", 12, 0)
   Call CREACAMPO("COMPRAS", "OCOMENCA", "Lentrega", 10, 64)
   Call CREACAMPO("COMPRAS", "OCOMENCA", "Anex_Ocom", 10, 128)
   Call CREACAMPO("COMPRAS", "OCOMENCA", "OcomRevis", 3, 2)
   Call CREACAMPO("COMPRAS", "OCOMENCA", "Nume_Ocom", 4, 4, 1)
   Call CREACAMPO("COMPRAS", "OCOMENCA", "Nume_Pedi", 4, 0)
   Call CREACAMPO("COMPRAS", "OCOMENCA", "Refe_Ocom", 10, 64)
   Call CREACAMPO("COMPRAS", "OCOMENCA", "Femi_Ocom", 8, 0)
   Call CREACAMPO("COMPRAS", "OCOMENCA", "Npro_Ocom", 3, 2)
   Call CREACAMPO("COMPRAS", "OCOMENCA", "Nume_Cli", 4, 0)
   Call CREACAMPO("COMPRAS", "OCOMENCA", "RaSo_Prov", 10, 48)
   Call CREACAMPO("COMPRAS", "OCOMENCA", "TSum_Ocom", 3, 2)
   Call CREACAMPO("COMPRAS", "OCOMENCA", "Stat_Ocom", 3, 2)
   Call CREACAMPO("COMPRAS", "OCOMENCA", "FGen_Ocom", 8, 0)
   Call CREACAMPO("COMPRAS", "OCOMENCA", "Usua_Gen", 3, 2)
   Call CREACAMPO("COMPRAS", "OCOMENCA", "FApro_Ocom", 8, 0)
   Call CREACAMPO("COMPRAS", "OCOMENCA", "Usua_Apro", 3, 2)
   Call CREACAMPO("COMPRAS", "OCOMENCA", "FMod_Ocom", 8, 0)
   Call CREACAMPO("COMPRAS", "OCOMENCA", "Usua_Mod", 3, 2)
   Call CREACAMPO("COMPRAS", "OCOMENCA", "FAnu_Ocom", 8, 0)
   Call CREACAMPO("COMPRAS", "OCOMENCA", "Usua_Anul", 3, 2)
   Call CREACAMPO("COMPRAS", "OCOMENCA", "Moti_Anul", 10, 16)
   Call CREACAMPO("COMPRAS", "OCOMENCA", "Mone_Cos", 3, 2)
   Call CREACAMPO("COMPRAS", "OCOMENCA", "Mone_Emis", 3, 2)
   Call CREACAMPO("COMPRAS", "OCOMENCA", "Tcam_Ocom", 6, 4)
   Call CREACAMPO("COMPRAS", "OCOMENCA", "ToBru_Ocom", 7, 8)
   Call CREACAMPO("COMPRAS", "OCOMENCA", "Pdes_Ocom", 10, 24)
   Call CREACAMPO("COMPRAS", "OCOMENCA", "Ides_Ocom", 7, 8)
   Call CREACAMPO("COMPRAS", "OCOMENCA", "Iiva_Ocom", 7, 8)
   Call CREACAMPO("COMPRAS", "OCOMENCA", "Iotr1_Ocom", 7, 8)
   Call CREACAMPO("COMPRAS", "OCOMENCA", "Iotr2_Ocom", 7, 8)
   Call CREACAMPO("COMPRAS", "OCOMENCA", "Total_Ocom", 7, 8)
   Call CREACAMPO("COMPRAS", "OCOMENCA", "Cpag_Ocom", 3, 2)
   Call CREACAMPO("COMPRAS", "OCOMENCA", "Embala_Ocom", 3, 2)
   Call CREACAMPO("COMPRAS", "OCOMENCA", "PorcIva", 10, 16)
   Call CREACAMPO("COMPRAS", "OCOMENCA", "Id_Comprob", 10, 16)
   Call CREACAMPO("COMPRAS", "OCOMENCA", "CodCarpImp", 10, 12)
   Call CREACAMPO("COMPRAS", "OCOMENCA", "Respon_Ocompra", 10, 64)
   Call CREACAMPO("COMPRAS", "OCOMENCA", "Via_Entrega", 3, 0)
   '
   Call CREACAMPO("COMPRAS", "PROVECOD", "Num_Prov", 4, 0, 1)
   Call CREACAMPO("COMPRAS", "PROVECOD", "Cod_inte", 4, 0, 1)
   Call CREACAMPO("COMPRAS", "PROVECOD", "Precio", 7, 0)
   Call CREACAMPO("COMPRAS", "PROVECOD", "Moneda", 3, 2)
   Call CREACAMPO("COMPRAS", "PROVECOD", "Fecha", 8, 0)
   Call CREACAMPO("COMPRAS", "PROVECOD", "MarBorra", 3, 0)
   '
End Sub
'
Sub CREATABLA_CONTAGEN()
   '
   Call CREACAMPO("CONTAGEN", "DetaAsien", "NuIdAsien", 10, 16, 1)
   Call CREACAMPO("CONTAGEN", "DetaAsien", "OrNuPase", 3, 0)
   Call CREACAMPO("CONTAGEN", "DetaAsien", "RefePase", 10, 64)
   Call CREACAMPO("CONTAGEN", "DetaAsien", "FechAsien", 8, 0, 1)
   Call CREACAMPO("CONTAGEN", "DetaAsien", "OrFeAsien", 3, 0)
   Call CREACAMPO("CONTAGEN", "DetaAsien", "SucuAsien", 3, 0)
   Call CREACAMPO("CONTAGEN", "DetaAsien", "NuDfAsien", 4, 0)
   Call CREACAMPO("CONTAGEN", "DetaAsien", "StatPase", 3, 0, 1)
   Call CREACAMPO("CONTAGEN", "DetaAsien", "CueCoInt", 3, 0, 1)
   Call CREACAMPO("CONTAGEN", "DetaAsien", "CueCoExt", 10, 16)
   Call CREACAMPO("CONTAGEN", "DetaAsien", "IDebePase", 5, 0)
   Call CREACAMPO("CONTAGEN", "DetaAsien", "IHabePase", 5, 0)
   Call CREACAMPO("CONTAGEN", "DetaAsien", "MarConsi", 10, 8)
   Call CREACAMPO("CONTAGEN", "DetaAsien", "CenCos", 10, 8)
   Call CREACAMPO("CONTAGEN", "DetaAsien", "Niv_Clasi", 3, 2, 1)
   '
   Call CREACAMPO("CONTAGEN", "EncabeAsien", "NuIdAsien", 10, 16)
   Call CREACAMPO("CONTAGEN", "EncabeAsien", "RefeAsien", 10, 64)
   Call CREACAMPO("CONTAGEN", "EncabeAsien", "FechAsien", 8, 0)
   Call CREACAMPO("CONTAGEN", "EncabeAsien", "OrFeAsien", 3, 0)
   Call CREACAMPO("CONTAGEN", "EncabeAsien", "SucuAsien", 3, 0)
   Call CREACAMPO("CONTAGEN", "EncabeAsien", "NuDfAsien", 4, 0)
   Call CREACAMPO("CONTAGEN", "EncabeAsien", "TipoAsien", 3, 0)
   Call CREACAMPO("CONTAGEN", "EncabeAsien", "StatAsien", 3, 0)
   Call CREACAMPO("CONTAGEN", "EncabeAsien", "EditAsien", 3, 0)
   Call CREACAMPO("CONTAGEN", "EncabeAsien", "QPasAsien", 3, 0)
   Call CREACAMPO("CONTAGEN", "EncabeAsien", "DebeAsien", 5, 0)
   Call CREACAMPO("CONTAGEN", "EncabeAsien", "Observa", 12, 0)
   Call CREACAMPO("CONTAGEN", "EncabeAsien", "Nume_Usuar", 3, 0)
   Call CREACAMPO("CONTAGEN", "EncabeAsien", "Niv_Clasi", 3, 2, 1)
   '
   Call CREACAMPO("CONTAGEN", "PLACOPRE", "CueContaI", 3, 0, 1)
   Call CREACAMPO("CONTAGEN", "PLACOPRE", "CenCosto", 10, 8)
   Call CREACAMPO("CONTAGEN", "PLACOPRE", "Fecha", 8, 0)
   Call CREACAMPO("CONTAGEN", "PLACOPRE", "Concepto", 10, 128)
   Call CREACAMPO("CONTAGEN", "PLACOPRE", "Ingreso", 7, 0)
   Call CREACAMPO("CONTAGEN", "PLACOPRE", "Gastos", 7, 0)
   Call CREACAMPO("CONTAGEN", "PLACOPRE", "Marborra", 3, 0)
   Call CREACAMPO("CONTAGEN", "PLACOPRE", "CodCuenta", 10, 16)
   Call CREACAMPO("CONTAGEN", "PLACOPRE", "MesAno", 10, 24)
   Call CREACAMPO("CONTAGEN", "PLACOPRE", "Tipo", 10, 8)
   Call CREACAMPO("CONTAGEN", "PLACOPRE", "Imp_Real", 7, 0)
   Call CREACAMPO("CONTAGEN", "PLACOPRE", "Ing_Gas", 10, 8)
   '
   Call CREACAMPO("CONTAGEN", "PLANCUEN", "CodExCue", 10, 16, 1)
   Call CREACAMPO("CONTAGEN", "PLANCUEN", "Denocue", 10, 120)
   Call CREACAMPO("CONTAGEN", "PLANCUEN", "CodInCue", 3, 0, 1)
   Call CREACAMPO("CONTAGEN", "PLANCUEN", "Statcue", 3, 0, 1)
   Call CREACAMPO("CONTAGEN", "PLANCUEN", "NivAgru", 3, 4)
   Call CREACAMPO("CONTAGEN", "PLANCUEN", "TipoCue", 10, 8)
   Call CREACAMPO("CONTAGEN", "PLANCUEN", "NuOrdIt", 3, 0)
   Call CREACAMPO("CONTAGEN", "PLANCUEN", "PosImpu", 3, 0)
   Call CREACAMPO("CONTAGEN", "PLANCUEN", "Observa", 12, 0)
   Call CREACAMPO("CONTAGEN", "PLANCUEN", "Salini", 5, 0)
   Call CREACAMPO("CONTAGEN", "PLANCUEN", "Sudebe", 5, 0)
   Call CREACAMPO("CONTAGEN", "PLANCUEN", "Suhabe", 5, 0)
   Call CREACAMPO("CONTAGEN", "PLANCUEN", "Result", 5, 0)
   Call CREACAMPO("CONTAGEN", "PLANCUEN", "Salfin", 5, 0)
   Call CREACAMPO("CONTAGEN", "PLANCUEN", "IndiAjus", 3, 0)
   Call CREACAMPO("CONTAGEN", "PLANCUEN", "MoneExtr", 3, 0)
   '
   Call CREACAMPO("CONTAGEN", "TABEJCON", "CodiEjer", 3, 0)
   Call CREACAMPO("CONTAGEN", "TABEJCON", "DenoEjCon", 10, 48)
   Call CREACAMPO("CONTAGEN", "TABEJCON", "FechAper", 8, 0)
   Call CREACAMPO("CONTAGEN", "TABEJCON", "FechCierr", 8, 0)
   Call CREACAMPO("CONTAGEN", "TABEJCON", "Status", 3, 0)
   Call CREACAMPO("CONTAGEN", "TABEJCON", "Diverif", 3, 0)
   Call CREACAMPO("CONTAGEN", "TABEJCON", "MarcActi", 3, 0)
   '
   Call CREACAMPO("CONTAGEN", "TAPROYEC", "Cod_Proy", 10, 8)
   Call CREACAMPO("CONTAGEN", "TAPROYEC", "Deno_Proy", 10, 48)
   Call CREACAMPO("CONTAGEN", "TAPROYEC", "Clien_Proy", 3, 2)
   Call CREACAMPO("CONTAGEN", "TAPROYEC", "FeAlt_Proy", 8, 2)
   Call CREACAMPO("CONTAGEN", "TAPROYEC", "FeCie_Proy", 8, 2)
   Call CREACAMPO("CONTAGEN", "TAPROYEC", "Stat_Proy", 3, 2)
   Call CREACAMPO("CONTAGEN", "TAPROYEC", "Descrip_Proy", 12, 0)
   '
End Sub
'
Sub CREATABLA_CUECORRI()
   '
   Call CREACAMPO("CUECORRI", "APLICOBRA", "CodiEmpr", 3, 2)
   Call CREACAMPO("CUECORRI", "APLICOBRA", "CodiSucu", 3, 2)
   Call CREACAMPO("CUECORRI", "APLICOBRA", "NuClien", 4, 2)
   Call CREACAMPO("CUECORRI", "APLICOBRA", "CodiClie", 10, 16)
   Call CREACAMPO("CUECORRI", "APLICOBRA", "RasoClie", 10, 48)
   Call CREACAMPO("CUECORRI", "APLICOBRA", "FechEmisi", 8, 0)
   Call CREACAMPO("CUECORRI", "APLICOBRA", "CoDocAp_Lar", 10, 18)
   Call CREACAMPO("CUECORRI", "APLICOBRA", "CoDocAp_Cor", 10, 12)
   Call CREACAMPO("CUECORRI", "APLICOBRA", "ImpoApli", 7, 8)
   Call CREACAMPO("CUECORRI", "APLICOBRA", "RefeApli", 10, 32)
   Call CREACAMPO("CUECORRI", "APLICOBRA", "CoDocOr_Lar", 10, 18)
   Call CREACAMPO("CUECORRI", "APLICOBRA", "CoDocOr_Cor", 10, 12)
   Call CREACAMPO("CUECORRI", "APLICOBRA", "ItotOrig", 7, 8)
   Call CREACAMPO("CUECORRI", "APLICOBRA", "IIvaOrig", 7, 8)
   Call CREACAMPO("CUECORRI", "APLICOBRA", "Mone_Apli", 3, 2)
   Call CREACAMPO("CUECORRI", "APLICOBRA", "ItotOri_Dol", 7, 15)
   Call CREACAMPO("CUECORRI", "APLICOBRA", "ImpoApli_Dol", 7, 15)
   '
   Call CREACAMPO("CUECORRI", "APLIPAGO", "CodiEmpr", 3, 2)
   Call CREACAMPO("CUECORRI", "APLIPAGO", "CodiSucu", 3, 2)
   Call CREACAMPO("CUECORRI", "APLIPAGO", "NuProve", 3, 2)
   Call CREACAMPO("CUECORRI", "APLIPAGO", "CodiProv", 10, 16)
   Call CREACAMPO("CUECORRI", "APLIPAGO", "RasoProv", 10, 48)
   Call CREACAMPO("CUECORRI", "APLIPAGO", "FechEmisi", 8, 0)
   Call CREACAMPO("CUECORRI", "APLIPAGO", "CoDocAp_Lar", 10, 18)
   Call CREACAMPO("CUECORRI", "APLIPAGO", "CoDocAp_Cor", 10, 12)
   Call CREACAMPO("CUECORRI", "APLIPAGO", "ImpoApli", 7, 8)
   Call CREACAMPO("CUECORRI", "APLIPAGO", "RefeApli", 10, 32)
   Call CREACAMPO("CUECORRI", "APLIPAGO", "CoDocOr_Lar", 10, 18)
   Call CREACAMPO("CUECORRI", "APLIPAGO", "CoDocOr_Cor", 10, 12)
   Call CREACAMPO("CUECORRI", "APLIPAGO", "ItotOrig", 7, 8)
   Call CREACAMPO("CUECORRI", "APLIPAGO", "IIvaOrig", 7, 8)
   Call CREACAMPO("CUECORRI", "APLIPAGO", "Monemi_Ori", 3, 0)
   Call CREACAMPO("CUECORRI", "APLIPAGO", "ImpoApli_Ori", 7, 8)
   '
   Call CREACAMPO("CUECORRI", "CAJABANC", "IdenAlta", 10, 8)
   Call CREACAMPO("CUECORRI", "CAJABANC", "IdenBaja", 10, 8)
   Call CREACAMPO("CUECORRI", "CAJABANC", "FechEntre", 8, 0)
   Call CREACAMPO("CUECORRI", "CAJABANC", "FeHoReal", 8, 0)
   Call CREACAMPO("CUECORRI", "CAJABANC", "NuAsiDiar", 10, 16, 1)
   Call CREACAMPO("CUECORRI", "CAJABANC", "FeDeposEnt", 8, 0)
   Call CREACAMPO("CUECORRI", "CAJABANC", "DocumEnt", 10, 24)
   Call CREACAMPO("CUECORRI", "CAJABANC", "DestiCheq", 10, 64)
   Call CREACAMPO("CUECORRI", "CAJABANC", "RefeSalChe", 10, 48)
   Call CREACAMPO("CUECORRI", "CAJABANC", "IdenSalChe", 10, 48)
   Call CREACAMPO("CUECORRI", "CAJABANC", "UsuaSalChe", 10, 48)
   Call CREACAMPO("CUECORRI", "CAJABANC", "PerLiComis", 10, 24)
   Call CREACAMPO("CUECORRI", "CAJABANC", "ImpoComis", 7, 0)
   Call CREACAMPO("CUECORRI", "CAJABANC", "VendeComis", 3, 0)
   Call CREACAMPO("CUECORRI", "CAJABANC", "CodiCom_Cor", 10, 12, 1)
   Call CREACAMPO("CUECORRI", "CAJABANC", "CodiCom_Lar", 10, 24, 1)
   Call CREACAMPO("CUECORRI", "CAJABANC", "RefeCom", 10, 48)
   Call CREACAMPO("CUECORRI", "CAJABANC", "FechConta", 8, 0, 1)
   Call CREACAMPO("CUECORRI", "CAJABANC", "FechAcred", 8, 0)
   Call CREACAMPO("CUECORRI", "CAJABANC", "MarConsi", 10, 4)
   Call CREACAMPO("CUECORRI", "CAJABANC", "CueContaI", 3, 0, 1)
   Call CREACAMPO("CUECORRI", "CAJABANC", "CenCosto", 10, 8)
   Call CREACAMPO("CUECORRI", "CAJABANC", "ItemConc", 10, 8)
   Call CREACAMPO("CUECORRI", "CAJABANC", "NuClien", 4, 2, 1)
   Call CREACAMPO("CUECORRI", "CAJABANC", "Sucu_Clie", 3, 2, 1)
   Call CREACAMPO("CUECORRI", "CAJABANC", "NuProve", 3, 3, 1)
   Call CREACAMPO("CUECORRI", "CAJABANC", "TipoMovim", 10, 8, 1)
   Call CREACAMPO("CUECORRI", "CAJABANC", "OpciMovim", 10, 8, 1)
   Call CREACAMPO("CUECORRI", "CAJABANC", "VariMovim", 10, 8, 1)
   Call CREACAMPO("CUECORRI", "CAJABANC", "DescriMov", 10, 48, 1)
   Call CREACAMPO("CUECORRI", "CAJABANC", "FechEmisi", 8, 0, 1)
   Call CREACAMPO("CUECORRI", "CAJABANC", "CodiEmpr", 3, 0, 1)
   Call CREACAMPO("CUECORRI", "CAJABANC", "CodiSucu", 3, 0)
   Call CREACAMPO("CUECORRI", "CAJABANC", "NumeTalo", 3, 0, 1)
   Call CREACAMPO("CUECORRI", "CAJABANC", "NumeComp", 4, 0, 1)
   Call CREACAMPO("CUECORRI", "CAJABANC", "Status", 3, 0, 1)
   Call CREACAMPO("CUECORRI", "CAJABANC", "Mone_Emis", 3, 2)
   Call CREACAMPO("CUECORRI", "CAJABANC", "Tipo_Cam", 7, 0)
   Call CREACAMPO("CUECORRI", "CAJABANC", "Impo_Dola", 7, 0)
   Call CREACAMPO("CUECORRI", "CAJABANC", "ValAbsComp", 7, 0, 1)
   Call CREACAMPO("CUECORRI", "CAJABANC", "IDebeComp", 7, 0, 1)
   Call CREACAMPO("CUECORRI", "CAJABANC", "IHabeComp", 7, 0, 1)
   Call CREACAMPO("CUECORRI", "CAJABANC", "BaseImpo", 7, 0)
   Call CREACAMPO("CUECORRI", "CAJABANC", "IncSubIva", 10, 2)
   Call CREACAMPO("CUECORRI", "CAJABANC", "NroCai", 10, 24)
   Call CREACAMPO("CUECORRI", "CAJABANC", "FeVtoCai", 8, 0)
   Call CREACAMPO("CUECORRI", "CAJABANC", "ObserGen", 10, 64)
   Call CREACAMPO("CUECORRI", "CAJABANC", "AtraCheque", 4, 0)
   Call CREACAMPO("CUECORRI", "CAJABANC", "AtraCance", 4, 0)
   Call CREACAMPO("CUECORRI", "CAJABANC", "CondPag", 3, 0)
   Call CREACAMPO("CUECORRI", "CAJABANC", "CodRetGan", 10, 3)
   Call CREACAMPO("CUECORRI", "CAJABANC", "CodPerGan", 10, 3)
   Call CREACAMPO("CUECORRI", "CAJABANC", "Niv_Clasi", 3, 2, 1)
   Call CREACAMPO("CUECORRI", "CAJABANC", "LetraComis", 10, 8)
   Call CREACAMPO("CUECORRI", "CAJABANC", "Vend_Cli", 3, 0)
   Call CREACAMPO("CUECORRI", "CAJABANC", "Cae", 10, 32)
   Call CREACAMPO("CUECORRI", "CAJABANC", "IdCae", 10, 32)
   Call CREACAMPO("CUECORRI", "CAJABANC", "CodCarpImp", 10, 12)
   Call CREACAMPO("CUECORRI", "CAJABANC", "TextoCae", 12, 0)
   Call CREACAMPO("CUECORRI", "CAJABANC", "Notas", 12, 0)
   Call CREACAMPO("CUECORRI", "CAJABANC", "Incoterms", 10, 12)
   Call CREACAMPO("CUECORRI", "CAJABANC", "OCompra", 12, 0)
   Call CREACAMPO("CUECORRI", "CAJABANC", "Pedidos", 12, 0)
   Call CREACAMPO("CUECORRI", "CAJABANC", "Cli_Real", 4, 2, 1)
   Call CREACAMPO("CUECORRI", "CAJABANC", "FechVtoCae", 10, 32)
   Call CREACAMPO("CUECORRI", "CAJABANC", "SimMone", 10, 16)
   Call CREACAMPO("CUECORRI", "CAJABANC", "Cod_Barra", 10, 64)
   Call CREACAMPO("CUECORRI", "CAJABANC", "Descu1", 10, 32)
   Call CREACAMPO("CUECORRI", "CAJABANC", "Descu2", 10, 32)
   Call CREACAMPO("CUECORRI", "CAJABANC", "ImpoDesc1", 7, 0)
   Call CREACAMPO("CUECORRI", "CAJABANC", "ImpoDesc2", 7, 0)
   Call CREACAMPO("CUECORRI", "CAJABANC", "IDCAENUM", 4, 0, 1)
   Call CREACAMPO("CUECORRI", "CAJABANC", "COTIZA_DOLAR", 7, 0)
   Call CREACAMPO("CUECORRI", "CAJABANC", "CodigoImpuesto", 3, 0)
   Call CREACAMPO("CUECORRI", "CAJABANC", "CodigoRegimen", 3, 0)
   '
   Call CREACAMPO("", "FACTUANTI", "FechEmisi", 8, 0)
   Call CREACAMPO("", "FACTUANTI", "DOCU_ORIG", 10, 18)
   Call CREACAMPO("", "FACTUANTI", "DOCU_APLI", 10, 18)
   Call CREACAMPO("", "FACTUANTI", "IMPODEBE", 7, 0)
   Call CREACAMPO("", "FACTUANTI", "IMPOHABE", 7, 0)
   Call CREACAMPO("", "FACTUANTI", "STATUS", 3, 0)
   Call CREACAMPO("", "FACTUANTI", "NuClien", 4, 0, 1)
   Call CREACAMPO("", "FACTUANTI", "Cli_Real", 4, 0, 1)
   Call CREACAMPO("", "FACTUANTI", "NumeCom", 3, 0)
   Call CREACAMPO("", "FACTUANTI", "TIPOMOVIM", 10, 8, 1)
   Call CREACAMPO("", "FACTUANTI", "VARIMOVIM", 10, 8)
   Call CREACAMPO("", "FACTUANTI", "RefeCom", 10, 128)
   Call CREACAMPO("", "FACTUANTI", "COTIZA_DOLAR", 7, 0)
   Call CREACAMPO("", "FACTUANTI", "BASEIMPODEBE", 7, 0)
   Call CREACAMPO("", "FACTUANTI", "BASEIMPOHABE", 7, 0)
   Call CREACAMPO("", "FACTUANTI", "OBSERVACIONES", 10, 1024)
   Call CREACAMPO("", "FACTUANTI", "AJUSTE_NUMERO", 4, 1)
   Call CREACAMPO("", "FACTUANTI", "USUARIO", 10, 18)
   Call CREACAMPO("", "FACTUANTI", "OBSERVACIONES", 10, 128)
   Call CREACAMPO("", "FACTUANTI", "Mone_Emis", 3, 4)
   '
   Call CREACAMPO("CUECORRI", "DETMOVCC", "N_Mov", 4, 0)
   Call CREACAMPO("CUECORRI", "DETMOVCC", "N_OrIt", 4, 0)
   Call CREACAMPO("CUECORRI", "DETMOVCC", "Fec_Mov", 8, 0)
   Call CREACAMPO("CUECORRI", "DETMOVCC", "N_CueImpu", 3, 0)
   Call CREACAMPO("CUECORRI", "DETMOVCC", "ImporMovI", 7, 0)
   Call CREACAMPO("CUECORRI", "DETMOVCC", "MoneMov", 3, 0)
   Call CREACAMPO("CUECORRI", "DETMOVCC", "ParMonMov", 7, 0)
   '
   Call CREACAMPO("CUECORRI", "DETRECCC", "N_Mov", 4, 0)
   Call CREACAMPO("CUECORRI", "DETRECCC", "T_Comp", 10, 25)
   Call CREACAMPO("CUECORRI", "DETRECCC", "N_Comp", 10, 25)
   Call CREACAMPO("CUECORRI", "DETRECCC", "N_Cli", 4, 0)
   Call CREACAMPO("CUECORRI", "DETRECCC", "Fec_Mov", 8, 0)
   Call CREACAMPO("CUECORRI", "DETRECCC", "N_CueImpu", 4, 0)
   Call CREACAMPO("CUECORRI", "DETRECCC", "ImporMovI", 7, 0)
   Call CREACAMPO("CUECORRI", "DETRECCC", "MoneMov", 3, 0)
   Call CREACAMPO("CUECORRI", "DETRECCC", "ParMonMov", 7, 0)
   '
   Call CREACAMPO("CUECORRI", "ENCMOVCC", "N_Mov", 4, 0)
   Call CREACAMPO("CUECORRI", "ENCMOVCC", "Compro", 10, 25)
   Call CREACAMPO("CUECORRI", "ENCMOVCC", "T_Comp", 10, 25)
   Call CREACAMPO("CUECORRI", "ENCMOVCC", "N_Comp", 4, 0)
   Call CREACAMPO("CUECORRI", "ENCMOVCC", "N_Cli", 4, 0)
   '
   Call CREACAMPO("CUECORRI", "PLAGASPEN", "Fecha", 8, 0)
   Call CREACAMPO("CUECORRI", "PLAGASPEN", "Referencia", 10, 64)
   Call CREACAMPO("CUECORRI", "PLAGASPEN", "Ingresos", 7, 0)
   Call CREACAMPO("CUECORRI", "PLAGASPEN", "Gastos", 7, 0)
   Call CREACAMPO("CUECORRI", "PLAGASPEN", "Tipo", 3, 0)
   Call CREACAMPO("CUECORRI", "PLAGASPEN", "MarBorra", 3, 0)
   '
   Call CREACAMPO("CUECORRI", "SACREPEN", "CodiEmpr", 3, 2, 1)
   Call CREACAMPO("CUECORRI", "SACREPEN", "CodiSucu", 3, 2)
   Call CREACAMPO("CUECORRI", "SACREPEN", "NuClien", 3, 2, 1)
   Call CREACAMPO("CUECORRI", "SACREPEN", "NuProve", 3, 2, 1)
   Call CREACAMPO("CUECORRI", "SACREPEN", "CodCliPro", 10, 16)
   Call CREACAMPO("CUECORRI", "SACREPEN", "RaSoCliPro", 10, 48)
   Call CREACAMPO("CUECORRI", "SACREPEN", "CodiCom_Cor", 10, 12)
   Call CREACAMPO("CUECORRI", "SACREPEN", "CodiCom_Lar", 10, 24)
   Call CREACAMPO("CUECORRI", "SACREPEN", "FechEmisi", 8, 0)
   Call CREACAMPO("CUECORRI", "SACREPEN", "Mone_Emis", 3, 0)
   Call CREACAMPO("CUECORRI", "SACREPEN", "Ibru_Dola", 7, 0)
   Call CREACAMPO("CUECORRI", "SACREPEN", "SaldDebe_Dola", 7, 0)
   Call CREACAMPO("CUECORRI", "SACREPEN", "SaldAcre_Dola", 7, 0)
   Call CREACAMPO("CUECORRI", "SACREPEN", "IBruComp", 7, 8)
   Call CREACAMPO("CUECORRI", "SACREPEN", "IIvaComp", 7, 8)
   Call CREACAMPO("CUECORRI", "SACREPEN", "OrVencim", 3, 2)
   Call CREACAMPO("CUECORRI", "SACREPEN", "FeVencim", 8, 0)
   Call CREACAMPO("CUECORRI", "SACREPEN", "SaldDebe", 7, 8)
   Call CREACAMPO("CUECORRI", "SACREPEN", "SaldAcre", 7, 8)
   Call CREACAMPO("CUECORRI", "SACREPEN", "Status", 3, 2)
   Call CREACAMPO("CUECORRI", "SACREPEN", "Niv_Clasi", 3, 2, 1)
   Call CREACAMPO("CUECORRI", "SACREPEN", "Observa", 12, 0)
   '
   Call CREACAMPO("CUECORRI", "SADEUPEN", "CodiEmpr", 3, 2, 1)
   Call CREACAMPO("CUECORRI", "SADEUPEN", "CodiSucu", 3, 2)
   Call CREACAMPO("CUECORRI", "SADEUPEN", "NuClien", 4, 2)
   Call CREACAMPO("CUECORRI", "SADEUPEN", "Sucu_Clie", 3, 2)
   Call CREACAMPO("CUECORRI", "SADEUPEN", "NuProve", 3, 2)
   Call CREACAMPO("CUECORRI", "SADEUPEN", "CodCliPro", 10, 16)
   Call CREACAMPO("CUECORRI", "SADEUPEN", "RaSoCliPro", 10, 48)
   Call CREACAMPO("CUECORRI", "SADEUPEN", "CodiCom_Cor", 10, 12)
   Call CREACAMPO("CUECORRI", "SADEUPEN", "CodiCom_Lar", 10, 24)
   Call CREACAMPO("CUECORRI", "SADEUPEN", "FechEmisi", 8, 0)
   Call CREACAMPO("CUECORRI", "SADEUPEN", "Mone_Emis", 3, 2)
   Call CREACAMPO("CUECORRI", "SADEUPEN", "Ibru_Dola", 7, 15)
   Call CREACAMPO("CUECORRI", "SADEUPEN", "SaldDebe_Dola", 7, 15)
   Call CREACAMPO("CUECORRI", "SADEUPEN", "SaldAcre_Dola", 7, 15)
   Call CREACAMPO("CUECORRI", "SADEUPEN", "IbruComp", 7, 8)
   Call CREACAMPO("CUECORRI", "SADEUPEN", "IIvaComp", 7, 8)
   Call CREACAMPO("CUECORRI", "SADEUPEN", "OrVencim", 3, 2)
   Call CREACAMPO("CUECORRI", "SADEUPEN", "FeVencim", 8, 0)
   Call CREACAMPO("CUECORRI", "SADEUPEN", "SaldDebe", 7, 8)
   Call CREACAMPO("CUECORRI", "SADEUPEN", "SaldAcre", 7, 8)
   Call CREACAMPO("CUECORRI", "SADEUPEN", "Status", 3, 2)
   Call CREACAMPO("CUECORRI", "SADEUPEN", "CondPag", 3, 0)
   '
End Sub
'
Sub CREATABLA_EMPLEADOS()
   '
   Call CREACAMPO("EMPLEADOS", "CompxEmple", "Legajo", 3, 2)
   Call CREACAMPO("EMPLEADOS", "CompxEmple", "Competencia", 10, 8)
   Call CREACAMPO("EMPLEADOS", "CompxEmple", "FechAlta", 8, 0)
   Call CREACAMPO("EMPLEADOS", "CompxEmple", "FechVenc", 8, 0)
   '
   Call CREACAMPO("EMPLEADOS", "PADRON", "Legajo", 3, 2, 1)
   Call CREACAMPO("EMPLEADOS", "PADRON", "ID_SYS", 10, 16, 1)
   Call CREACAMPO("EMPLEADOS", "PADRON", "Nombre", 10, 64)
   Call CREACAMPO("EMPLEADOS", "PADRON", "Grupo", 3, 2, 1)
   Call CREACAMPO("EMPLEADOS", "PADRON", "Status", 3, 2, 1)
   Call CREACAMPO("EMPLEADOS", "PADRON", "Documento", 10, 13)
   Call CREACAMPO("EMPLEADOS", "PADRON", "Sexo", 10, 1)
   Call CREACAMPO("EMPLEADOS", "PADRON", "FechaNacim", 8, 10)
   Call CREACAMPO("EMPLEADOS", "PADRON", "Nacionalidad", 10, 32)
   Call CREACAMPO("EMPLEADOS", "PADRON", "EstadoCivil", 10, 10)
   Call CREACAMPO("EMPLEADOS", "PADRON", "Domicilio", 10, 64)
   Call CREACAMPO("EMPLEADOS", "PADRON", "CodPostal", 10, 10)
   Call CREACAMPO("EMPLEADOS", "PADRON", "Localidad", 10, 64)
   Call CREACAMPO("EMPLEADOS", "PADRON", "Sector", 6, 4)
   Call CREACAMPO("EMPLEADOS", "PADRON", "Tarea", 10, 32)
   Call CREACAMPO("EMPLEADOS", "PADRON", "Categoria", 10, 9)
   Call CREACAMPO("EMPLEADOS", "PADRON", "Horario", 10, 32)
   Call CREACAMPO("EMPLEADOS", "PADRON", "FechaIng", 8, 10)
   Call CREACAMPO("EMPLEADOS", "PADRON", "CUIL", 10, 13)
   Call CREACAMPO("EMPLEADOS", "PADRON", "AFJP", 10, 32)
   Call CREACAMPO("EMPLEADOS", "PADRON", "CajaJub", 10, 32)
   Call CREACAMPO("EMPLEADOS", "PADRON", "FechaEgre", 8, 10)
   Call CREACAMPO("EMPLEADOS", "PADRON", "MotivoEgre", 10, 64)
   Call CREACAMPO("EMPLEADOS", "PADRON", "Foto", 10, 100)
   Call CREACAMPO("EMPLEADOS", "PADRON", "Anotaciones", 12, 100)
   Call CREACAMPO("EMPLEADOS", "PADRON", "FechaAntig", 8, 10)
   '
End Sub
'
Sub CREATABLA_EQUIPOS()
   '
   Call CREACAMPO("EQUIPOS", "CONTROLFIN", "NumeroSerie", 10, 32, 1)
   Call CREACAMPO("EQUIPOS", "CONTROLFIN", "Cod_Inte", 3, 1)
   Call CREACAMPO("EQUIPOS", "CONTROLFIN", "NumeroOrden", 3, 1)
   Call CREACAMPO("EQUIPOS", "CONTROLFIN", "Resultado", 10, 4)
   Call CREACAMPO("EQUIPOS", "CONTROLFIN", "MarBorra", 3, 1)
   Call CREACAMPO("EQUIPOS", "CONTROLFIN", "Fech_Ensa", 8, 0)
   Call CREACAMPO("EQUIPOS", "CONTROLFIN", "Revision", 3, 0)
   Call CREACAMPO("EQUIPOS", "CONTROLFIN", "Origen", 10, 16)
   Call CREACAMPO("EQUIPOS", "CONTROLFIN", "NumControl", 4, 0)
   Call CREACAMPO("EQUIPOS", "CONTROLFIN", "Tec_Ensa", 10, 32)
   Call CREACAMPO("EQUIPOS", "CONTROLFIN", "Aext_Ensa", 10, 32)
   Call CREACAMPO("EQUIPOS", "CONTROLFIN", "Aint_Ensa", 10, 32)
   Call CREACAMPO("EQUIPOS", "CONTROLFIN", "Obse_Ensa", 12, 0)
   Call CREACAMPO("EQUIPOS", "CONTROLFIN", "Dic_Ensa", 10, 32)
   '
   Call CREACAMPO("EQUIPOS", "ENSAR070", "NumeroSerie", 10, 32, 1)
   Call CREACAMPO("EQUIPOS", "ENSAR070", "Modelo", 3, 1)
   Call CREACAMPO("EQUIPOS", "ENSAR070", "Inspector", 10, 32)
   Call CREACAMPO("EQUIPOS", "ENSAR070", "Comentarios", 12, 0)
   Call CREACAMPO("EQUIPOS", "ENSAR070", "Corr_Ensa1", 6, 0)
   Call CREACAMPO("EQUIPOS", "ENSAR070", "CTension1", 6, 0)
   Call CREACAMPO("EQUIPOS", "ENSAR070", "Veredicto1", 10, 6)
   Call CREACAMPO("EQUIPOS", "ENSAR070", "Corr_Ensa2", 6, 0)
   Call CREACAMPO("EQUIPOS", "ENSAR070", "CTension2", 6, 0)
   Call CREACAMPO("EQUIPOS", "ENSAR070", "VerGeneral", 10, 6)
   Call CREACAMPO("EQUIPOS", "ENSAR070", "Veredicto2", 10, 6)
   Call CREACAMPO("EQUIPOS", "ENSAR070", "Bloqueado", 10, 2)
   Call CREACAMPO("EQUIPOS", "ENSAR070", "Fech_Ensa", 8, 0)
   Call CREACAMPO("EQUIPOS", "ENSAR070", "Tecnico", 10, 32)
   Call CREACAMPO("EQUIPOS", "ENSAR070", "Revision", 3, 0)
   Call CREACAMPO("EQUIPOS", "ENSAR070", "Corr_Ensa3", 6, 0)
   Call CREACAMPO("EQUIPOS", "ENSAR070", "Corr_Ensa4", 6, 0)
   Call CREACAMPO("EQUIPOS", "ENSAR070", "Corr_Ensa5", 6, 0)
   Call CREACAMPO("EQUIPOS", "ENSAR070", "Corr_Ensa6", 6, 0)
   Call CREACAMPO("EQUIPOS", "ENSAR070", "CTension3", 6, 0)
   Call CREACAMPO("EQUIPOS", "ENSAR070", "CTension4", 6, 0)
   Call CREACAMPO("EQUIPOS", "ENSAR070", "CTension5", 6, 0)
   Call CREACAMPO("EQUIPOS", "ENSAR070", "CTension6", 6, 0)
   Call CREACAMPO("EQUIPOS", "ENSAR070", "Veredicto3", 10, 6)
   Call CREACAMPO("EQUIPOS", "ENSAR070", "Veredicto4", 10, 6)
   Call CREACAMPO("EQUIPOS", "ENSAR070", "Veredicto5", 10, 6)
   Call CREACAMPO("EQUIPOS", "ENSAR070", "Veredicto6", 10, 6)
   Call CREACAMPO("EQUIPOS", "ENSAR070", "Instrumento", 10, 32)
   '
   Call CREACAMPO("EQUIPOS", "ENSAR072", "NumeroSerie", 10, 32, 1)
   Call CREACAMPO("EQUIPOS", "ENSAR072", "Modelo", 3, 1)
   Call CREACAMPO("EQUIPOS", "ENSAR072", "Inspector", 10, 32)
   Call CREACAMPO("EQUIPOS", "ENSAR072", "CorrFuga1", 6, 0)
   Call CREACAMPO("EQUIPOS", "ENSAR072", "Veredicto1", 10, 6)
   Call CREACAMPO("EQUIPOS", "ENSAR072", "CorrFuga2", 6, 0)
   Call CREACAMPO("EQUIPOS", "ENSAR072", "Veredicto2", 10, 6)
   Call CREACAMPO("EQUIPOS", "ENSAR072", "CorrFuga3", 6, 0)
   Call CREACAMPO("EQUIPOS", "ENSAR072", "Veredicto3", 10, 6)
   Call CREACAMPO("EQUIPOS", "ENSAR072", "CorrFuga4", 6, 0)
   Call CREACAMPO("EQUIPOS", "ENSAR072", "Veredicto4", 10, 6)
   Call CREACAMPO("EQUIPOS", "ENSAR072", "CorrFuga5", 6, 0)
   Call CREACAMPO("EQUIPOS", "ENSAR072", "Veredicto5", 10, 6)
   Call CREACAMPO("EQUIPOS", "ENSAR072", "CorrFuga6", 6, 0)
   Call CREACAMPO("EQUIPOS", "ENSAR072", "Veredicto6", 10, 6)
   Call CREACAMPO("EQUIPOS", "ENSAR072", "Corrfuga7", 6, 0)
   Call CREACAMPO("EQUIPOS", "ENSAR072", "Veredicto7", 10, 6)
   Call CREACAMPO("EQUIPOS", "ENSAR072", "CorrFuga8", 6, 0)
   Call CREACAMPO("EQUIPOS", "ENSAR072", "Veredicto8", 10, 6)
   Call CREACAMPO("EQUIPOS", "ENSAR072", "Bloqueado", 10, 2)
   Call CREACAMPO("EQUIPOS", "ENSAR072", "VerGeneral", 10, 6)
   Call CREACAMPO("EQUIPOS", "ENSAR072", "Comentarios", 12, 0)
   Call CREACAMPO("EQUIPOS", "ENSAR072", "Fech_Ensa", 8, 0)
   Call CREACAMPO("EQUIPOS", "ENSAR072", "Tecnico", 10, 32)
   Call CREACAMPO("EQUIPOS", "ENSAR072", "Revision", 3, 0)
   Call CREACAMPO("EQUIPOS", "ENSAR072", "Tension", 10, 32)
   Call CREACAMPO("EQUIPOS", "ENSAR072", "Instrumento", 10, 32)
   Call CREACAMPO("EQUIPOS", "ENSAR072", "RiDieInstru", 10, 32)
   Call CREACAMPO("EQUIPOS", "ENSAR072", "RiDieTension", 10, 32)
   Call CREACAMPO("EQUIPOS", "ENSAR072", "RiDieTiempo", 10, 16)
   Call CREACAMPO("EQUIPOS", "ENSAR072", "RiDieVeredicto", 10, 6)
   Call CREACAMPO("EQUIPOS", "ENSAR072", "RiDieObserva", 12, 0)
   '
   Call CREACAMPO("EQUIPOS", "ENSAR074", "NumeroSerie", 10, 32, 1)
   Call CREACAMPO("EQUIPOS", "ENSAR074", "Modelo", 3, 1)
   Call CREACAMPO("EQUIPOS", "ENSAR074", "Inspector", 10, 32)
   Call CREACAMPO("EQUIPOS", "ENSAR074", "Comentarios", 12, 0)
   Call CREACAMPO("EQUIPOS", "ENSAR074", "VerGeneral", 10, 6)
   Call CREACAMPO("EQUIPOS", "ENSAR074", "Inten_Med1", 6, 0)
   Call CREACAMPO("EQUIPOS", "ENSAR074", "Inten_Med2", 6, 0)
   Call CREACAMPO("EQUIPOS", "ENSAR074", "Inten_Med3", 6, 0)
   Call CREACAMPO("EQUIPOS", "ENSAR074", "Inten_Med4", 6, 0)
   Call CREACAMPO("EQUIPOS", "ENSAR074", "Inten_Ver1", 10, 6)
   Call CREACAMPO("EQUIPOS", "ENSAR074", "Inten_Ver2", 10, 6)
   Call CREACAMPO("EQUIPOS", "ENSAR074", "Inten_Ver3", 10, 6)
   Call CREACAMPO("EQUIPOS", "ENSAR074", "Inten_Ver4", 10, 6)
   Call CREACAMPO("EQUIPOS", "ENSAR074", "AnPul_Med1", 6, 0)
   Call CREACAMPO("EQUIPOS", "ENSAR074", "AnPul_Med2", 6, 0)
   Call CREACAMPO("EQUIPOS", "ENSAR074", "AnPul_Med3", 6, 0)
   Call CREACAMPO("EQUIPOS", "ENSAR074", "AnPul_Med4", 6, 0)
   Call CREACAMPO("EQUIPOS", "ENSAR074", "AnPul_Ver1", 10, 6)
   Call CREACAMPO("EQUIPOS", "ENSAR074", "AnPul_Ver2", 10, 6)
   Call CREACAMPO("EQUIPOS", "ENSAR074", "AnPul_Ver3", 10, 6)
   Call CREACAMPO("EQUIPOS", "ENSAR074", "AnPul_Ver4", 10, 6)
   Call CREACAMPO("EQUIPOS", "ENSAR074", "Frec_Med1", 6, 0)
   Call CREACAMPO("EQUIPOS", "ENSAR074", "Frec_Med2", 6, 0)
   Call CREACAMPO("EQUIPOS", "ENSAR074", "Frec_Med3", 6, 0)
   Call CREACAMPO("EQUIPOS", "ENSAR074", "Frec_Med4", 6, 0)
   Call CREACAMPO("EQUIPOS", "ENSAR074", "Frec_Ver1", 10, 6)
   Call CREACAMPO("EQUIPOS", "ENSAR074", "Frec_Ver2", 10, 6)
   Call CREACAMPO("EQUIPOS", "ENSAR074", "Frec_Ver3", 10, 6)
   Call CREACAMPO("EQUIPOS", "ENSAR074", "Frec_Ver4", 10, 6)
   Call CREACAMPO("EQUIPOS", "ENSAR074", "Fech_Ensa", 8, 0)
   Call CREACAMPO("EQUIPOS", "ENSAR074", "Bloqueado", 10, 2)
   Call CREACAMPO("EQUIPOS", "ENSAR074", "Tecnico", 10, 32)
   Call CREACAMPO("EQUIPOS", "ENSAR074", "Revision", 3, 0)
   Call CREACAMPO("EQUIPOS", "ENSAR074", "INSTRUMENTO1", 10, 32)
   Call CREACAMPO("EQUIPOS", "ENSAR074", "INSTRUMENTO2", 10, 32)
   '
   Call CREACAMPO("EQUIPOS", "MEDIFIN", "NumeroSerie", 10, 32, 1)
   Call CREACAMPO("EQUIPOS", "MEDIFIN", "Cod_Inte", 3, 1)
   Call CREACAMPO("EQUIPOS", "MEDIFIN", "CHN", 10, 6)
   Call CREACAMPO("EQUIPOS", "MEDIFIN", "MedA", 6, 0)
   Call CREACAMPO("EQUIPOS", "MEDIFIN", "MedB", 6, 0)
   Call CREACAMPO("EQUIPOS", "MEDIFIN", "MedC", 6, 0)
   Call CREACAMPO("EQUIPOS", "MEDIFIN", "MarBorra", 3, 1)
   Call CREACAMPO("EQUIPOS", "MEDIFIN", "Fech_Ensa", 8, 0)
   Call CREACAMPO("EQUIPOS", "MEDIFIN", "Revision", 3, 0)
   Call CREACAMPO("EQUIPOS", "MEDIFIN", "Origen", 10, 16)
   '
   Call CREACAMPO("EQUIPOS", "TANUMSE", "NumeroSerie", 10, 32, 1)
   Call CREACAMPO("EQUIPOS", "TANUMSE", "Mesanio", 10, 4)
   Call CREACAMPO("EQUIPOS", "TANUMSE", "NumeOrd", 3, 1)
   Call CREACAMPO("EQUIPOS", "TANUMSE", "Cod_Inte", 3, 1)
   Call CREACAMPO("EQUIPOS", "TANUMSE", "Cod_Exte", 10, 24)
   Call CREACAMPO("EQUIPOS", "TANUMSE", "Descrip", 10, 64)
   Call CREACAMPO("EQUIPOS", "TANUMSE", "IdSubOr", 10, 12, 1)
   Call CREACAMPO("EQUIPOS", "TANUMSE", "FechAlta", 8, 0)
   Call CREACAMPO("EQUIPOS", "TANUMSE", "FechEntrega", 8, 0)
   Call CREACAMPO("EQUIPOS", "TANUMSE", "NumRemito", 10, 18)
   Call CREACAMPO("EQUIPOS", "TANUMSE", "NumFactura", 10, 18)
   Call CREACAMPO("EQUIPOS", "TANUMSE", "NumeCli", 4, 1)
   Call CREACAMPO("EQUIPOS", "TANUMSE", "Paciente", 10, 18)
   Call CREACAMPO("EQUIPOS", "TANUMSE", "Ten_Ensa", 3, 0)
   Call CREACAMPO("EQUIPOS", "TANUMSE", "Dic_Ensa", 10, 4)
   Call CREACAMPO("EQUIPOS", "TANUMSE", "Tec_Ensa", 10, 32)
   Call CREACAMPO("EQUIPOS", "TANUMSE", "Aext_Ensa", 10, 32)
   Call CREACAMPO("EQUIPOS", "TANUMSE", "Obse_Ensa", 12, 0)
   Call CREACAMPO("EQUIPOS", "TANUMSE", "Fech_Ensa", 8, 0)
   Call CREACAMPO("EQUIPOS", "TANUMSE", "Aint_Ensa", 10, 32)
   Call CREACAMPO("EQUIPOS", "TANUMSE", "DepoNumse", 3, 0, 1)
   Call CREACAMPO("EQUIPOS", "TANUMSE", "Historial", 12, 0)
   Call CREACAMPO("EQUIPOS", "TANUMSE", "FeVenGar", 8, 0)
   Call CREACAMPO("EQUIPOS", "TANUMSE", "FeVenGaRep", 8, 0)
   Call CREACAMPO("EQUIPOS", "TANUMSE", "UNumRep", 4, 0)
   Call CREACAMPO("EQUIPOS", "TANUMSE", "Ubicacion", 10, 128)
   Call CREACAMPO("EQUIPOS", "TANUMSE", "INSTRUMENTO1", 10, 32)
   Call CREACAMPO("EQUIPOS", "TANUMSE", "INSTRUMENTO2", 10, 32)
   Call CREACAMPO("EQUIPOS", "TANUMSE", "INSTRUMENTO3", 10, 32)
   Call CREACAMPO("EQUIPOS", "TANUMSE", "INSTRUMENTO4", 10, 32)
   '
End Sub
'
Sub CREATABLA_MAQUINARIAS()
   '
   Call CREACAMPO("MAQUINARIAS", "CONTADORES", "Num_Con", 3, 2)
   Call CREACAMPO("MAQUINARIAS", "CONTADORES", "Referen_Con", 10, 32)
   Call CREACAMPO("MAQUINARIAS", "CONTADORES", "Freq_Con", 6, 4)
   Call CREACAMPO("MAQUINARIAS", "CONTADORES", "FUL_Con", 8, 2)
   Call CREACAMPO("MAQUINARIAS", "CONTADORES", "Val_Con", 7, 8)
   Call CREACAMPO("MAQUINARIAS", "CONTADORES", "FPL_Con", 8, 2)
   Call CREACAMPO("MAQUINARIAS", "CONTADORES", "Prom_Con", 7, 8)
   Call CREACAMPO("MAQUINARIAS", "CONTADORES", "Observa_Con", 10, 64)
   Call CREACAMPO("MAQUINARIAS", "CONTADORES", "Codigo_Equ", 10, 6)
   '
   Call CREACAMPO("MAQUINARIAS", "EQUIPOS", "Codigo_Equ", 10, 6)
   Call CREACAMPO("MAQUINARIAS", "EQUIPOS", "CodLar_Equ", 10, 16)
   Call CREACAMPO("MAQUINARIAS", "EQUIPOS", "Descri_Equ", 10, 64)
   Call CREACAMPO("MAQUINARIAS", "EQUIPOS", "Marca_Equ", 10, 48)
   Call CREACAMPO("MAQUINARIAS", "EQUIPOS", "Mode_Equ", 10, 48)
   Call CREACAMPO("MAQUINARIAS", "EQUIPOS", "NuSeRi_Equ", 10, 48)
   Call CREACAMPO("MAQUINARIAS", "EQUIPOS", "Fabri_Equ", 10, 48)
   Call CREACAMPO("MAQUINARIAS", "EQUIPOS", "Represe_Equ", 3, 2)
   Call CREACAMPO("MAQUINARIAS", "EQUIPOS", "Ubica_Equ", 6, 4)
   Call CREACAMPO("MAQUINARIAS", "EQUIPOS", "Alimenta_Equ", 10, 48)
   Call CREACAMPO("MAQUINARIAS", "EQUIPOS", "Potenci_Equ", 10, 48)
   Call CREACAMPO("MAQUINARIAS", "EQUIPOS", "Kilos_Equ", 6, 0)
   Call CREACAMPO("MAQUINARIAS", "EQUIPOS", "Largo_Equ", 6, 0)
   Call CREACAMPO("MAQUINARIAS", "EQUIPOS", "Ancho_Equ", 6, 0)
   Call CREACAMPO("MAQUINARIAS", "EQUIPOS", "Alto_Equ", 6, 0)
   Call CREACAMPO("MAQUINARIAS", "EQUIPOS", "Linea_Equ", 10, 48)
   Call CREACAMPO("MAQUINARIAS", "EQUIPOS", "Posi_Equ", 10, 12)
   Call CREACAMPO("MAQUINARIAS", "EQUIPOS", "Planta_Equ", 10, 4)
   Call CREACAMPO("MAQUINARIAS", "EQUIPOS", "Finst_Equ", 8, 0)
   Call CREACAMPO("MAQUINARIAS", "EQUIPOS", "Fgara_Equ", 8, 0)
   Call CREACAMPO("MAQUINARIAS", "EQUIPOS", "Observa_Equ", 12, 0)
   '
   Call CREACAMPO("MAQUINARIAS", "GALEIMAG", "Codigo_Equ", 10, 6)
   Call CREACAMPO("MAQUINARIAS", "GALEIMAG", "Norit_Imag", 3, 2)
   Call CREACAMPO("MAQUINARIAS", "GALEIMAG", "Ruta_Imag", 12, 0)
   Call CREACAMPO("MAQUINARIAS", "GALEIMAG", "Refe_Imag", 12, 0)
   '
   Call CREACAMPO("MAQUINARIAS", "PLAMANPREV", "Codigo_Equ", 10, 6)
   Call CREACAMPO("MAQUINARIAS", "PLAMANPREV", "CodLar_Equ", 10, 16)
   Call CREACAMPO("MAQUINARIAS", "PLAMANPREV", "Nro_Sub", 4, 4)
   Call CREACAMPO("MAQUINARIAS", "PLAMANPREV", "Orden_Tar", 3, 2)
   Call CREACAMPO("MAQUINARIAS", "PLAMANPREV", "Descri_Tar", 10, 128)
   Call CREACAMPO("MAQUINARIAS", "PLAMANPREV", "Frecu_Tar", 4, 4)
   Call CREACAMPO("MAQUINARIAS", "PLAMANPREV", "Uni_Frecu", 10, 16)
   Call CREACAMPO("MAQUINARIAS", "PLAMANPREV", "Obse_Tar", 10, 128)
   '
   Call CREACAMPO("MAQUINARIAS", "REPUMAQ", "Codigo_Equ", 10, 6)
   Call CREACAMPO("MAQUINARIAS", "REPUMAQ", "CodInte_Rep", 3, 2)
   Call CREACAMPO("MAQUINARIAS", "REPUMAQ", "CodExte_Rep", 10, 16)
   Call CREACAMPO("MAQUINARIAS", "REPUMAQ", "Descrip_Rep", 10, 64)
   Call CREACAMPO("MAQUINARIAS", "REPUMAQ", "UM_Rep", 10, 4)
   Call CREACAMPO("MAQUINARIAS", "REPUMAQ", "Cant_Rep", 6, 4)
   Call CREACAMPO("MAQUINARIAS", "REPUMAQ", "Observa_Rep", 10, 128)
   '
   Call CREACAMPO("MAQUINARIAS", "SUBCONJUNTOS", "Nro_Sub", 3, 2)
   Call CREACAMPO("MAQUINARIAS", "SUBCONJUNTOS", "Descri_Sub", 10, 32)
   Call CREACAMPO("MAQUINARIAS", "SUBCONJUNTOS", "Marca_Sub", 10, 24)
   Call CREACAMPO("MAQUINARIAS", "SUBCONJUNTOS", "Modelo_Sub", 10, 20)
   Call CREACAMPO("MAQUINARIAS", "SUBCONJUNTOS", "NSerie_Sub", 10, 20)
   Call CREACAMPO("MAQUINARIAS", "SUBCONJUNTOS", "Codigo_Equ", 10, 6)
   Call CREACAMPO("MAQUINARIAS", "SUBCONJUNTOS", "Criticidad", 10, 2)
   '
End Sub
'
Sub CREATABLA_MENSAFLEX()
   '
   Call CREACAMPO("MENSAFLEX", "MENSAJES", "NumMensa", 4, 0, 1)
   Call CREACAMPO("MENSAFLEX", "MENSAJES", "MensaViejo", 12, 0)
   Call CREACAMPO("MENSAFLEX", "MENSAJES", "MensaEsp", 12, 0)
   Call CREACAMPO("MENSAFLEX", "MENSAJES", "MensaPor", 12, 0)
   '
End Sub
'
Sub CREATABLA_OFABRIC()
   '
   Call CREACAMPO("OFABRIC", "DOCANRE", "NumRepa", 4, 0)
   Call CREACAMPO("OFABRIC", "DOCANRE", "NomAsignado", 10, 256)
   Call CREACAMPO("OFABRIC", "DOCANRE", "FILE_NAME", 10, 256)
   Call CREACAMPO("OFABRIC", "DOCANRE", "FILE_DATA", 12, 0)
   '
   Call CREACAMPO("OFABRIC", "GALIMREP", "NumRepa", 4, 0)
   Call CREACAMPO("OFABRIC", "GALIMREP", "Descripcion", 12, 0)
   Call CREACAMPO("OFABRIC", "GALIMREP", "Norit_Imag", 3, 2)
   Call CREACAMPO("OFABRIC", "GALIMREP", "Nom_Imag", 10, 256)
   Call CREACAMPO("OFABRIC", "GALIMREP", "Imagen", 12, 0)
   '
   Call CREACAMPO("OFABRIC", "INFOFAB", "IdSubOr", 10, 12)
   Call CREACAMPO("OFABRIC", "INFOFAB", "DescriOF", 10, 64)
   Call CREACAMPO("OFABRIC", "INFOFAB", "NuOrInf", 3, 2)
   Call CREACAMPO("OFABRIC", "INFOFAB", "Cod_Inte", 3, 2)
   Call CREACAMPO("OFABRIC", "INFOFAB", "Cod_Exte", 10, 16)
   Call CREACAMPO("OFABRIC", "INFOFAB", "Descrip", 10, 64)
   Call CREACAMPO("OFABRIC", "INFOFAB", "NumeOper", 3, 3)
   Call CREACAMPO("OFABRIC", "INFOFAB", "AltOper", 3, 0)
   Call CREACAMPO("OFABRIC", "INFOFAB", "DescOper", 10, 48)
   Call CREACAMPO("OFABRIC", "INFOFAB", "ReTraba", 3, 0)
   Call CREACAMPO("OFABRIC", "INFOFAB", "AtMulti", 3, 0)
   Call CREACAMPO("OFABRIC", "INFOFAB", "CanOrPro", 7, 8)
   Call CREACAMPO("OFABRIC", "INFOFAB", "CantReal", 7, 8)
   Call CREACAMPO("OFABRIC", "INFOFAB", "CaNoConf", 7, 8)
   Call CREACAMPO("OFABRIC", "INFOFAB", "Maquina", 10, 6)
   Call CREACAMPO("OFABRIC", "INFOFAB", "DescMaq", 10, 64)
   Call CREACAMPO("OFABRIC", "INFOFAB", "LegaOps", 11, 64)
   Call CREACAMPO("OFABRIC", "INFOFAB", "FeHorLib", 8, 0)
   Call CREACAMPO("OFABRIC", "INFOFAB", "Hora_Des", 3, 0)
   Call CREACAMPO("OFABRIC", "INFOFAB", "Hora_Has", 3, 0)
   Call CREACAMPO("OFABRIC", "INFOFAB", "Hora_Fri", 3, 0)
   Call CREACAMPO("OFABRIC", "INFOFAB", "MaqLiber", 10, 1)
   Call CREACAMPO("OFABRIC", "INFOFAB", "TiemEjec", 6, 4)
   Call CREACAMPO("OFABRIC", "INFOFAB", "TiemStan", 6, 4)
   Call CREACAMPO("OFABRIC", "INFOFAB", "Rendim", 6, 4)
   Call CREACAMPO("OFABRIC", "INFOFAB", "Observa", 10, 30)
   Call CREACAMPO("OFABRIC", "INFOFAB", "Referen", 12, 0)
   Call CREACAMPO("OFABRIC", "INFOFAB", "MotiRecha", 3, 0)
   Call CREACAMPO("OFABRIC", "INFOFAB", "Status", 3, 0)
   Call CREACAMPO("OFABRIC", "INFOFAB", "NroInfNC", 10, 32)
   Call CREACAMPO("OFABRIC", "INFOFAB", "RefeInf", 10, 64)
   Call CREACAMPO("OFABRIC", "INFOFAB", "UserControl", 10, 64)
   Call CREACAMPO("OFABRIC", "INFOFAB", "Operarios", 12, 0)
   Call CREACAMPO("OFABRIC", "INFOFAB", "IdInfofab", 10, 24, 1)
   '
   Call CREACAMPO("OFABRIC", "INFOPROY", "Nuor_Inf", 4, 4)
   Call CREACAMPO("OFABRIC", "INFOPROY", "Nume_Inf", 10, 12)
   Call CREACAMPO("OFABRIC", "INFOPROY", "Fecha_Inf", 8, 0)
   Call CREACAMPO("OFABRIC", "INFOPROY", "Refer_Inf", 10, 64)
   Call CREACAMPO("OFABRIC", "INFOPROY", "Cant_Inf", 6, 4)
   Call CREACAMPO("OFABRIC", "INFOPROY", "Proy_Inf", 10, 8)
   Call CREACAMPO("OFABRIC", "INFOPROY", "LegaOps_Inf", 11, 64)
   Call CREACAMPO("OFABRIC", "INFOPROY", "Total_Inf", 7, 8)
   '
   Call CREACAMPO("OFABRIC", "NEBRUNE", "Cod_Inte", 3, 2)
   Call CREACAMPO("OFABRIC", "NEBRUNE", "Cant_Reque", 7, 4)
   Call CREACAMPO("OFABRIC", "NEBRUNE", "Fech_Nec", 8, 0)
   Call CREACAMPO("OFABRIC", "NEBRUNE", "Referen", 10, 64)
   Call CREACAMPO("OFABRIC", "NEBRUNE", "Destino", 10, 32)
   '
   Call CREACAMPO("OFABRIC", "OPERFAB", "IdSubOr", 10, 12)
   Call CREACAMPO("OFABRIC", "OPERFAB", "Cod_Inte", 3, 0)
   Call CREACAMPO("OFABRIC", "OPERFAB", "Cod_Exte", 10, 16)
   Call CREACAMPO("OFABRIC", "OPERFAB", "Descrip", 10, 64)
   Call CREACAMPO("OFABRIC", "OPERFAB", "NumeOper", 3, 2)
   Call CREACAMPO("OFABRIC", "OPERFAB", "DescOper", 10, 48)
   Call CREACAMPO("OFABRIC", "OPERFAB", "Maquina", 10, 6)
   Call CREACAMPO("OFABRIC", "OPERFAB", "DescMaq", 10, 64)
   Call CREACAMPO("OFABRIC", "OPERFAB", "EquiAux", 10, 18)
   Call CREACAMPO("OFABRIC", "OPERFAB", "CantiOps", 6, 4)
   Call CREACAMPO("OFABRIC", "OPERFAB", "CantProg", 7, 8)
   Call CREACAMPO("OFABRIC", "OPERFAB", "CantReal", 7, 8)
   Call CREACAMPO("OFABRIC", "OPERFAB", "CantPend", 7, 8)
   Call CREACAMPO("OFABRIC", "OPERFAB", "TiemPrep", 6, 4)
   Call CREACAMPO("OFABRIC", "OPERFAB", "TiemFijo", 6, 4)
   Call CREACAMPO("OFABRIC", "OPERFAB", "TiemVari", 6, 4)
   Call CREACAMPO("OFABRIC", "OPERFAB", "TiemDemo", 6, 4)
   Call CREACAMPO("OFABRIC", "OPERFAB", "TiemMHer", 6, 4)
   Call CREACAMPO("OFABRIC", "OPERFAB", "TiemLimp", 6, 4)
   Call CREACAMPO("OFABRIC", "OPERFAB", "TiemOtro", 6, 4)
   Call CREACAMPO("OFABRIC", "OPERFAB", "TiemTota", 6, 4)
   Call CREACAMPO("OFABRIC", "OPERFAB", "TiemEjec", 6, 4)
   Call CREACAMPO("OFABRIC", "OPERFAB", "TiemPend", 6, 4)
   Call CREACAMPO("OFABRIC", "OPERFAB", "StatOper", 3, 2)
   Call CREACAMPO("OFABRIC", "OPERFAB", "ComiePro", 8, 0)
   Call CREACAMPO("OFABRIC", "OPERFAB", "TermiPro", 8, 0)
   Call CREACAMPO("OFABRIC", "OPERFAB", "TermiTemp", 8, 0)
   Call CREACAMPO("OFABRIC", "OPERFAB", "Observa", 10, 48)
   Call CREACAMPO("OFABRIC", "OPERFAB", "HoBatchRec", 11, 0)
   Call CREACAMPO("OFABRIC", "OPERFAB", "SecciPro", 6, 4)
   '
   Call CREACAMPO("OFABRIC", "ORDEFABR", "NumeOrFa", 4, 4)
   Call CREACAMPO("OFABRIC", "ORDEFABR", "Referen", 10, 64)
   Call CREACAMPO("OFABRIC", "ORDEFABR", "IdSubOr", 10, 12)
   Call CREACAMPO("OFABRIC", "ORDEFABR", "FechGen", 8, 0)
   Call CREACAMPO("OFABRIC", "ORDEFABR", "Cod_Inte", 3, 0)
   Call CREACAMPO("OFABRIC", "ORDEFABR", "Cod_Exte", 10, 16)
   Call CREACAMPO("OFABRIC", "ORDEFABR", "Descrip", 10, 64)
   Call CREACAMPO("OFABRIC", "ORDEFABR", "Unidad", 10, 32)
   Call CREACAMPO("OFABRIC", "ORDEFABR", "CanProy", 7, 8)
   Call CREACAMPO("OFABRIC", "ORDEFABR", "Destino", 10, 32)
   Call CREACAMPO("OFABRIC", "ORDEFABR", "FechEnSol", 8, 0)
   Call CREACAMPO("OFABRIC", "ORDEFABR", "Priorid", 3, 1)
   Call CREACAMPO("OFABRIC", "ORDEFABR", "StatuOrf", 3, 1)
   Call CREACAMPO("OFABRIC", "ORDEFABR", "NuOrProd", 10, 16)
   Call CREACAMPO("OFABRIC", "ORDEFABR", "FechLan", 8, 0)
   Call CREACAMPO("OFABRIC", "ORDEFABR", "SecciPro", 6, 4)
   Call CREACAMPO("OFABRIC", "ORDEFABR", "FechCierr", 8, 0)
   Call CREACAMPO("OFABRIC", "ORDEFABR", "CanApro", 7, 8)
   Call CREACAMPO("OFABRIC", "ORDEFABR", "CanRech", 7, 8)
   Call CREACAMPO("OFABRIC", "ORDEFABR", "IdenRetra", 10, 12)
   Call CREACAMPO("OFABRIC", "ORDEFABR", "FechProyEntre", 8, 0)
   Call CREACAMPO("OFABRIC", "ORDEFABR", "FechTempEntre", 8, 0)
   Call CREACAMPO("OFABRIC", "ORDEFABR", "FechRecal", 8, 0)
   Call CREACAMPO("OFABRIC", "ORDEFABR", "Observa", 12, 0)
   Call CREACAMPO("OFABRIC", "ORDEFABR", "Origen", 10, 64)
   Call CREACAMPO("OFABRIC", "ORDEFABR", "AnotaOrf", 12, 0)
   Call CREACAMPO("OFABRIC", "ORDEFABR", "IdenLote", 10, 32)
   Call CREACAMPO("OFABRIC", "ORDEFABR", "NumRem_Prov", 10, 16)
   Call CREACAMPO("OFABRIC", "ORDEFABR", "Nume_Ocom", 10, 16)
   Call CREACAMPO("OFABRIC", "ORDEFABR", "Nume_Cli", 4, 4)
   Call CREACAMPO("OFABRIC", "ORDEFABR", "Id_Orfa", 10, 32)
   Call CREACAMPO("OFABRIC", "ORDEFABR", "NroPed", 4, 0)
   Call CREACAMPO("OFABRIC", "ORDEFABR", "Armador ", 3, 0)
   Call CREACAMPO("OFABRIC", "ORDEFABR", "LoteIndivis", 7, 8)
   '
   Call CREACAMPO("OFABRIC", "RECONCILIA", "IdSubOr", 10, 16)
   Call CREACAMPO("OFABRIC", "RECONCILIA", "Cod_Inte", 3, 0)
   Call CREACAMPO("OFABRIC", "RECONCILIA", "IdenLote", 10, 16)
   Call CREACAMPO("OFABRIC", "RECONCILIA", "Consumido", 7, 8)
   Call CREACAMPO("OFABRIC", "RECONCILIA", "Mermas", 7, 8)
   Call CREACAMPO("OFABRIC", "RECONCILIA", "Devolucion", 7, 8)
   Call CREACAMPO("OFABRIC", "RECONCILIA", "ConsuNeto", 7, 8)
   Call CREACAMPO("OFABRIC", "RECONCILIA", "ConsuFormu", 7, 8)
   Call CREACAMPO("OFABRIC", "RECONCILIA", "Diferencia", 7, 8)
   Call CREACAMPO("OFABRIC", "RECONCILIA", "Autorizado", 10, 64)
   Call CREACAMPO("OFABRIC", "RECONCILIA", "Fecha", 8, 0)
   '
   Call CREACAMPO("OFABRIC", "REPARA", "NumRepa", 4, 0, 1)
   Call CREACAMPO("OFABRIC", "REPARA", "Referen", 10, 128)
   Call CREACAMPO("OFABRIC", "REPARA", "FechRecep", 8, 0)
   Call CREACAMPO("OFABRIC", "REPARA", "FechEnSol", 8, 0)
   Call CREACAMPO("OFABRIC", "REPARA", "Nume_Cli", 4, 0)
   Call CREACAMPO("OFABRIC", "REPARA", "Raso_Cli", 10, 64)
   Call CREACAMPO("OFABRIC", "REPARA", "Domi_Cli", 10, 64)
   Call CREACAMPO("OFABRIC", "REPARA", "Loca_Cli", 10, 64)
   Call CREACAMPO("OFABRIC", "REPARA", "Pais_Cli", 10, 64)
   Call CREACAMPO("OFABRIC", "REPARA", "Tele_Cli", 10, 24)
   Call CREACAMPO("OFABRIC", "REPARA", "Mail_Cli", 10, 48)
   Call CREACAMPO("OFABRIC", "REPARA", "Cod_Inte", 3, 0)
   Call CREACAMPO("OFABRIC", "REPARA", "Descripcion", 10, 128)
   Call CREACAMPO("OFABRIC", "REPARA", "NumeroSerie", 10, 32)
   Call CREACAMPO("OFABRIC", "REPARA", "Accesorios", 12, 0)
   Call CREACAMPO("OFABRIC", "REPARA", "Defec_Infor", 12, 0)
   Call CREACAMPO("OFABRIC", "REPARA", "Observa", 12, 0)
   Call CREACAMPO("OFABRIC", "REPARA", "Trab_Efect", 12, 0)
   Call CREACAMPO("OFABRIC", "REPARA", "PresuEsti", 7, 0)
   Call CREACAMPO("OFABRIC", "REPARA", "Status", 3, 0)
   Call CREACAMPO("OFABRIC", "REPARA", "Usuario_Repu", 10, 24)
   Call CREACAMPO("OFABRIC", "REPARA", "Estima_Horas", 7, 0)
   Call CREACAMPO("OFABRIC", "REPARA", "Or_Compra", 10, 24)
   Call CREACAMPO("OFABRIC", "REPARA", "FOr_Compra", 8, 0)
   Call CREACAMPO("OFABRIC", "REPARA", "NumRemito", 10, 18)
   Call CREACAMPO("OFABRIC", "REPARA", "FechEnt", 8, 0)
   Call CREACAMPO("OFABRIC", "REPARA", "Nom_Tra", 10, 48)
   Call CREACAMPO("OFABRIC", "REPARA", "Cuit_Cli", 10, 24)
   Call CREACAMPO("OFABRIC", "REPARA", "Piva_Cli", 10, 48)
   Call CREACAMPO("OFABRIC", "REPARA", "Conductor", 10, 48)
   Call CREACAMPO("OFABRIC", "REPARA", "Patente", 10, 16)
   Call CREACAMPO("OFABRIC", "REPARA", "Lentre_Cli", 10, 64)
   Call CREACAMPO("OFABRIC", "REPARA", "Status_Gar", 3, 0)
   Call CREACAMPO("OFABRIC", "REPARA", "CausaFalla", 3, 0)
   Call CREACAMPO("OFABRIC", "REPARA", "DeCauFalla", 10, 128)
   Call CREACAMPO("OFABRIC", "REPARA", "TipoFalla", 3, 0)
   Call CREACAMPO("OFABRIC", "REPARA", "DeTipFalla", 10, 128)
   Call CREACAMPO("OFABRIC", "REPARA", "Referen2", 10, 64)
   Call CREACAMPO("OFABRIC", "REPARA", "HoraMObra", 7, 12)
   Call CREACAMPO("OFABRIC", "REPARA", "Imp_Mobra", 7, 12)
   Call CREACAMPO("OFABRIC", "REPARA", "ValMObra", 7, 12)
   Call CREACAMPO("OFABRIC", "REPARA", "Imp_Repu", 7, 12)
   Call CREACAMPO("OFABRIC", "REPARA", "Imp_Tot", 7, 12)
   Call CREACAMPO("OFABRIC", "REPARA", "FechPres", 8, 0)
   Call CREACAMPO("OFABRIC", "REPARA", "FechApro", 8, 0)
   Call CREACAMPO("OFABRIC", "REPARA", "NroPed", 3, 12)
   Call CREACAMPO("OFABRIC", "REPARA", "AnotaRep", 12, 0)
   Call CREACAMPO("OFABRIC", "REPARA", "CodRepa", 10, 16)
   Call CREACAMPO("OFABRIC", "REPARA", "NuOrdRep", 4, 0, 1)
   Call CREACAMPO("OFABRIC", "REPARA", "RefOrdRep", 10, 64)
   Call CREACAMPO("OFABRIC", "REPARA", "FechOrdRep", 8, 0)
   Call CREACAMPO("OFABRIC", "REPARA", "Inspeccion", 12, 0)
   Call CREACAMPO("OFABRIC", "REPARA", "Reparacion", 12, 0)
   Call CREACAMPO("OFABRIC", "REPARA", "IdSubOr", 10, 12)
   Call CREACAMPO("OFABRIC", "REPARA", "Sitio", 12, 0)
   Call CREACAMPO("OFABRIC", "REPARA", "Factura", 10, 24)
   Call CREACAMPO("OFABRIC", "REPARA", "Importe", 10, 24)
   Call CREACAMPO("OFABRIC", "REPARA", "Tot_Repuestos", 7, 0)
   Call CREACAMPO("OFABRIC", "REPARA", "Tot_Mano_Obra", 7, 0)
   Call CREACAMPO("OFABRIC", "REPARA", "UsRealizaRepa", 10, 32)
   '
   Call CREACAMPO("OFABRIC", "REPUREPA", "NumRepa", 4, 0, 1)
   Call CREACAMPO("OFABRIC", "REPUREPA", "Cod_Inte", 3, 0)
   Call CREACAMPO("OFABRIC", "REPUREPA", "Cant_Sol", 7, 0)
   Call CREACAMPO("OFABRIC", "REPUREPA", "PrUn_Neto", 7, 0)
   Call CREACAMPO("OFABRIC", "REPUREPA", "MarBorra", 3, 0)
   Call CREACAMPO("OFABRIC", "REPUREPA", "Impo_Neto", 7, 0)
   Call CREACAMPO("OFABRIC", "REPUREPA", "Nuord_Item", 3, 0, 1)
   Call CREACAMPO("OFABRIC", "REPUREPA", "IdenLote", 10, 32)
   Call CREACAMPO("OFABRIC", "REPUREPA", "Origen", 3, 0)
   '
   Call CREACAMPO("OFABRIC", "RESERVASF", "Cod_Inte", 3, 2)
   '
End Sub
'
Sub CREATABLA_OMANTIN()
   '
   Call CREACAMPO("OMANTIN", "ORDEMANT", "NumeOrMa_OM", 4, 4)
   Call CREACAMPO("OMANTIN", "ORDEMANT", "Referen_OM", 10, 64)
   Call CREACAMPO("OMANTIN", "ORDEMANT", "FechGen_OM", 8, 0)
   Call CREACAMPO("OMANTIN", "ORDEMANT", "Codigo_Equ", 10, 6)
   Call CREACAMPO("OMANTIN", "ORDEMANT", "Descri_OM", 12, 0)
   Call CREACAMPO("OMANTIN", "ORDEMANT", "Status_OM", 3, 1)
   Call CREACAMPO("OMANTIN", "ORDEMANT", "FechCie_OM", 8, 0)
   Call CREACAMPO("OMANTIN", "ORDEMANT", "HoraCie_OM", 8, 0)
   Call CREACAMPO("OMANTIN", "ORDEMANT", "UsuGen_OM", 10, 25)
   Call CREACAMPO("OMANTIN", "ORDEMANT", "UsuCie_OM", 10, 25)
   Call CREACAMPO("OMANTIN", "ORDEMANT", "Proved_OM", 3, 2)
   Call CREACAMPO("OMANTIN", "ORDEMANT", "CostoTer_OM", 7, 8)
   Call CREACAMPO("OMANTIN", "ORDEMANT", "FeProCom_OM", 8, 0)
   Call CREACAMPO("OMANTIN", "ORDEMANT", "HoProCom_OM", 8, 0)
   Call CREACAMPO("OMANTIN", "ORDEMANT", "FeProFin_OM", 8, 0)
   Call CREACAMPO("OMANTIN", "ORDEMANT", "HoProFin_OM", 8, 0)
   '
   Call CREACAMPO("OMANTIN", "ORECORRI", "NumORec", 4, 0)
   Call CREACAMPO("OMANTIN", "ORECORRI", "DesORec", 10, 128)
   Call CREACAMPO("OMANTIN", "ORECORRI", "FechEmis", 8, 0)
   Call CREACAMPO("OMANTIN", "ORECORRI", "FechCierr", 8, 0)
   Call CREACAMPO("OMANTIN", "ORECORRI", "ObseEmis", 12, 0)
   Call CREACAMPO("OMANTIN", "ORECORRI", "ObseCierr", 12, 0)
   '
   Call CREACAMPO("OMANTIN", "RECUPROP", "NumeOrMa_OM", 4, 4)
   Call CREACAMPO("OMANTIN", "RECUPROP", "Descri_OM", 10, 64)
   Call CREACAMPO("OMANTIN", "RECUPROP", "LegaEmp_OM", 3, 2)
   Call CREACAMPO("OMANTIN", "RECUPROP", "Empleado_OM", 10, 64)
   Call CREACAMPO("OMANTIN", "RECUPROP", "FeCie_OM", 8, 0)
   Call CREACAMPO("OMANTIN", "RECUPROP", "CaHora_OM", 6, 4)
   '
   Call CREACAMPO("OMANTIN", "SETARECO", "Cod_Reco", 10, 8, 1)
   Call CREACAMPO("OMANTIN", "SETARECO", "Des_Reco", 10, 72)
   Call CREACAMPO("OMANTIN", "SETARECO", "Norit_Tar", 3, 0)
   Call CREACAMPO("OMANTIN", "SETARECO", "MaEquip", 10, 8)
   Call CREACAMPO("OMANTIN", "SETARECO", "Des_Larga", 12, 0)
   '
End Sub
'
Sub CREATABLA_OTRABAJO()
   '
   Call CREACAMPO("OTRABAJO", "GRUPOTRA", "Nro_Otrab", 10, 16)
   Call CREACAMPO("OTRABAJO", "GRUPOTRA", "Nro_Grup", 10, 10)
   Call CREACAMPO("OTRABAJO", "GRUPOTRA", "Prog_Grup", 10, 16)
   Call CREACAMPO("OTRABAJO", "GRUPOTRA", "Comp_Grup", 10, 16)
   Call CREACAMPO("OTRABAJO", "GRUPOTRA", "Deno_Grup", 10, 64)
   Call CREACAMPO("OTRABAJO", "GRUPOTRA", "Observa_Grup", 10, 64)
   Call CREACAMPO("OTRABAJO", "GRUPOTRA", "Pla05_Grup", 10, 10)
   Call CREACAMPO("OTRABAJO", "GRUPOTRA", "Pla07_Grup", 10, 10)
   '
   Call CREACAMPO("OTRABAJO", "LISTAMAT", "Nro_Otrab", 10, 16)
   Call CREACAMPO("OTRABAJO", "LISTAMAT", "Nro_Grup", 10, 10)
   Call CREACAMPO("OTRABAJO", "LISTAMAT", "NuOrden", 3, 0)
   Call CREACAMPO("OTRABAJO", "LISTAMAT", "Cantid", 7, 8)
   Call CREACAMPO("OTRABAJO", "LISTAMAT", "Descri_Cor", 10, 48)
   Call CREACAMPO("OTRABAJO", "LISTAMAT", "NPlano", 10, 24)
   Call CREACAMPO("OTRABAJO", "LISTAMAT", "Posic", 10, 16)
   Call CREACAMPO("OTRABAJO", "LISTAMAT", "Material", 10, 24)
   Call CREACAMPO("OTRABAJO", "LISTAMAT", "CodExte", 10, 24)
   Call CREACAMPO("OTRABAJO", "LISTAMAT", "Kilos", 6, 0)
   Call CREACAMPO("OTRABAJO", "LISTAMAT", "Clave", 10, 8)
   Call CREACAMPO("OTRABAJO", "LISTAMAT", "Descri_Lar", 12, 0)
   Call CREACAMPO("OTRABAJO", "LISTAMAT", "Nro_Ocom", 10, 16)
   Call CREACAMPO("OTRABAJO", "LISTAMAT", "Nro_BolRe", 10, 250)
   '
   Call CREACAMPO("OTRABAJO", "ORDETRAB", "Nro_Otrab", 10, 16)
   Call CREACAMPO("OTRABAJO", "ORDETRAB", "Refe_Otrab", 10, 64)
   Call CREACAMPO("OTRABAJO", "ORDETRAB", "Nclie_Otrab", 3, 2)
   Call CREACAMPO("OTRABAJO", "ORDETRAB", "Femi_Otrab", 8, 0)
   Call CREACAMPO("OTRABAJO", "ORDETRAB", "Fentre_Otrab", 8, 0)
   Call CREACAMPO("OTRABAJO", "ORDETRAB", "UsuConf_Otrab", 10, 32)
   Call CREACAMPO("OTRABAJO", "ORDETRAB", "UsuCont_Otrab", 10, 32)
   Call CREACAMPO("OTRABAJO", "ORDETRAB", "UsuApro_Otrab", 10, 32)
   Call CREACAMPO("OTRABAJO", "ORDETRAB", "NroOcom_Otrab", 10, 32)
   Call CREACAMPO("OTRABAJO", "ORDETRAB", "FeOcom_Otrab", 8, 0)
   Call CREACAMPO("OTRABAJO", "ORDETRAB", "NroOfer_Otrab", 10, 16)
   Call CREACAMPO("OTRABAJO", "ORDETRAB", "FeOfer_Otrab", 8, 0)
   Call CREACAMPO("OTRABAJO", "ORDETRAB", "NuOrItem_Otrab", 3, 2)
   Call CREACAMPO("OTRABAJO", "ORDETRAB", "CanPed_Otrab", 3, 2)
   Call CREACAMPO("OTRABAJO", "ORDETRAB", "Especi_Otrab", 12, 0)
   Call CREACAMPO("OTRABAJO", "ORDETRAB", "PreVen_Otrab", 7, 8)
   Call CREACAMPO("OTRABAJO", "ORDETRAB", "Mone_Otrab", 3, 2)
   Call CREACAMPO("OTRABAJO", "ORDETRAB", "CotiMone_Otrab", 7, 8)
   Call CREACAMPO("OTRABAJO", "ORDETRAB", "LEntrega_Otrab", 10, 255)
   Call CREACAMPO("OTRABAJO", "ORDETRAB", "Ajusta_Otrab", 3, 2)
   Call CREACAMPO("OTRABAJO", "ORDETRAB", "Vended_Otrab", 3, 2)
   Call CREACAMPO("OTRABAJO", "ORDETRAB", "ComiVend_Otrab", 7, 8)
   Call CREACAMPO("OTRABAJO", "ORDETRAB", "HsInge_Otrab", 6, 0)
   Call CREACAMPO("OTRABAJO", "ORDETRAB", "HsAvor_Otrab", 6, 0)
   Call CREACAMPO("OTRABAJO", "ORDETRAB", "HsCCal_Otrab", 6, 0)
   Call CREACAMPO("OTRABAJO", "ORDETRAB", "HsCald_Otrab", 6, 0)
   Call CREACAMPO("OTRABAJO", "ORDETRAB", "HsMeca_Otrab", 6, 0)
   Call CREACAMPO("OTRABAJO", "ORDETRAB", "HsMoInt_Otrab", 6, 0)
   Call CREACAMPO("OTRABAJO", "ORDETRAB", "HsMoExt_Otrab", 6, 0)
   Call CREACAMPO("OTRABAJO", "ORDETRAB", "HsIngeInf_Otrab", 6, 0)
   Call CREACAMPO("OTRABAJO", "ORDETRAB", "HsAvorInf_Otrab", 6, 0)
   Call CREACAMPO("OTRABAJO", "ORDETRAB", "HsCCalInf_Otrab", 6, 0)
   Call CREACAMPO("OTRABAJO", "ORDETRAB", "HsCaldInf_Otrab", 6, 0)
   Call CREACAMPO("OTRABAJO", "ORDETRAB", "HsMecaInf_Otrab", 6, 0)
   Call CREACAMPO("OTRABAJO", "ORDETRAB", "HsMoIntInf_Otrab", 6, 0)
   Call CREACAMPO("OTRABAJO", "ORDETRAB", "HsMoExtInf_Otrab", 6, 0)
   Call CREACAMPO("OTRABAJO", "ORDETRAB", "Observa_Otrab", 12, 0)
   Call CREACAMPO("OTRABAJO", "ORDETRAB", "FeRev_Otrab", 8, 0)
   Call CREACAMPO("OTRABAJO", "ORDETRAB", "NuRev_Otrab", 10, 6)
   Call CREACAMPO("OTRABAJO", "ORDETRAB", "IContab", 3, 2)
   '
   Call CREACAMPO("OTRABAJO", "PLACOB", "Nro_Otrab", 10, 16)
   Call CREACAMPO("OTRABAJO", "PLACOB", "NuOrItem", 3, 2)
   Call CREACAMPO("OTRABAJO", "PLACOB", "Refe_Pcob", 10, 64)
   Call CREACAMPO("OTRABAJO", "PLACOB", "Impo_Pcob", 7, 8)
   Call CREACAMPO("OTRABAJO", "PLACOB", "FeVen_Pcob", 8, 0)
   Call CREACAMPO("OTRABAJO", "PLACOB", "NroFac_Pcob", 10, 16)
   Call CREACAMPO("OTRABAJO", "PLACOB", "Moneda", 3, 2)
   Call CREACAMPO("OTRABAJO", "PLACOB", "Coti_Pcob", 7, 8)
   '
End Sub
'
Sub CREATABLA_PRESUCAM()
   '
   Call CREACAMPO("PRESUCAM", "DESPECIAL", "Nro_Presu", 10, 16)
   Call CREACAMPO("PRESUCAM", "DESPECIAL", "Num_Orden", 3, 2)
   Call CREACAMPO("PRESUCAM", "DESPECIAL", "Referen", 12, 0)
   Call CREACAMPO("PRESUCAM", "DESPECIAL", "Descrip", 12, 0)
   Call CREACAMPO("PRESUCAM", "DESPECIAL", "Importe", 7, 8)
   '
   Call CREACAMPO("PRESUCAM", "DETAPRESU", "Nro_Presu", 10, 16)
   Call CREACAMPO("PRESUCAM", "DETAPRESU", "Rubro_Det", 3, 2)
   Call CREACAMPO("PRESUCAM", "DETAPRESU", "SubRub_Det", 3, 2)
   Call CREACAMPO("PRESUCAM", "DETAPRESU", "NuorItem", 3, 2)
   Call CREACAMPO("PRESUCAM", "DETAPRESU", "CodInte_Det", 3, 2)
   Call CREACAMPO("PRESUCAM", "DETAPRESU", "CodExte_Det", 10, 16)
   Call CREACAMPO("PRESUCAM", "DETAPRESU", "Cantid_Det", 6, 0)
   Call CREACAMPO("PRESUCAM", "DETAPRESU", "PreUni_Det", 7, 8)
   Call CREACAMPO("PRESUCAM", "DETAPRESU", "Mone_Det", 3, 2)
   Call CREACAMPO("PRESUCAM", "DETAPRESU", "Med1_Det", 6, 0)
   Call CREACAMPO("PRESUCAM", "DETAPRESU", "Med2_Det", 6, 0)
   Call CREACAMPO("PRESUCAM", "DETAPRESU", "Med3_Det", 6, 0)
   Call CREACAMPO("PRESUCAM", "DETAPRESU", "Med4_Det", 6, 0)
   Call CREACAMPO("PRESUCAM", "DETAPRESU", "Impo_Det", 7, 8)
   Call CREACAMPO("PRESUCAM", "DETAPRESU", "Descri_Det", 10, 128)
   Call CREACAMPO("PRESUCAM", "DETAPRESU", "Hora_Det", 6, 0)
   Call CREACAMPO("PRESUCAM", "DETAPRESU", "Opera_Det", 3, 2)
   Call CREACAMPO("PRESUCAM", "DETAPRESU", "ValoHor_Det", 7, 8)
   '
   Call CREACAMPO("PRESUCAM", "DOCADJUN", "Nro_Presu", 10, 16)
   Call CREACAMPO("PRESUCAM", "DOCADJUN", "RutArch", 12, 0)
   '
   Call CREACAMPO("PRESUCAM", "ENCAPRESU", "Observa_Presu", 12, 0)
   Call CREACAMPO("PRESUCAM", "ENCAPRESU", "FeRev_Presu", 8, 0)
   Call CREACAMPO("PRESUCAM", "ENCAPRESU", "NuRev_Presu", 10, 6)
   Call CREACAMPO("PRESUCAM", "ENCAPRESU", "IContab", 3, 2)
   Call CREACAMPO("PRESUCAM", "ENCAPRESU", "Largo_Presu", 6, 0)
   Call CREACAMPO("PRESUCAM", "ENCAPRESU", "ForPag_Presu", 12, 0)
   Call CREACAMPO("PRESUCAM", "ENCAPRESU", "PlaEnt_Presu", 12, 0)
   Call CREACAMPO("PRESUCAM", "ENCAPRESU", "Precios_Presu", 12, 0)
   Call CREACAMPO("PRESUCAM", "ENCAPRESU", "ValOfer_Presu", 12, 0)
   Call CREACAMPO("PRESUCAM", "ENCAPRESU", "BoniVars_Presu", 7, 8)
   Call CREACAMPO("PRESUCAM", "ENCAPRESU", "FormElec_Presu", 6, 4)
   Call CREACAMPO("PRESUCAM", "ENCAPRESU", "HorasCons_Presu", 6, 4)
   Call CREACAMPO("PRESUCAM", "ENCAPRESU", "ImporCons_Presu", 7, 8)
   Call CREACAMPO("PRESUCAM", "ENCAPRESU", "Ancho_Presu", 6, 0)
   Call CREACAMPO("PRESUCAM", "ENCAPRESU", "Alto_Presu", 6, 0)
   Call CREACAMPO("PRESUCAM", "ENCAPRESU", "Espesor_Presu", 6, 0)
   Call CREACAMPO("PRESUCAM", "ENCAPRESU", "GGMat_Presu", 7, 8)
   Call CREACAMPO("PRESUCAM", "ENCAPRESU", "GGOba_Presu", 7, 8)
   Call CREACAMPO("PRESUCAM", "ENCAPRESU", "DesBon_Presu", 7, 8)
   Call CREACAMPO("PRESUCAM", "ENCAPRESU", "Nro_Presu", 10, 16)
   Call CREACAMPO("PRESUCAM", "ENCAPRESU", "Refe_Presu", 10, 64)
   Call CREACAMPO("PRESUCAM", "ENCAPRESU", "Nclie_Presu", 3, 2)
   Call CREACAMPO("PRESUCAM", "ENCAPRESU", "Femi_Presu", 8, 0)
   Call CREACAMPO("PRESUCAM", "ENCAPRESU", "FVali_Presu", 8, 0)
   Call CREACAMPO("PRESUCAM", "ENCAPRESU", "UsuConf_Presu", 10, 32)
   Call CREACAMPO("PRESUCAM", "ENCAPRESU", "UsuCont_Presu", 10, 32)
   Call CREACAMPO("PRESUCAM", "ENCAPRESU", "UsuApro_Presu", 10, 32)
   Call CREACAMPO("PRESUCAM", "ENCAPRESU", "NroOcom_Presu", 10, 32)
   Call CREACAMPO("PRESUCAM", "ENCAPRESU", "FeOcom_Presu", 8, 0)
   Call CREACAMPO("PRESUCAM", "ENCAPRESU", "PediCoti_Presu", 10, 16)
   Call CREACAMPO("PRESUCAM", "ENCAPRESU", "FePedCot_Presu", 8, 0)
   Call CREACAMPO("PRESUCAM", "ENCAPRESU", "NuOrItem_Presu", 3, 2)
   Call CREACAMPO("PRESUCAM", "ENCAPRESU", "CanPed_Presu", 3, 2)
   Call CREACAMPO("PRESUCAM", "ENCAPRESU", "Especi_Presu", 12, 0)
   Call CREACAMPO("PRESUCAM", "ENCAPRESU", "PreVen_Presu", 7, 8)
   Call CREACAMPO("PRESUCAM", "ENCAPRESU", "Mone_Presu", 3, 2)
   Call CREACAMPO("PRESUCAM", "ENCAPRESU", "CotiMone_Presu", 7, 8)
   Call CREACAMPO("PRESUCAM", "ENCAPRESU", "LEntrega_Presu", 10, 255)
   Call CREACAMPO("PRESUCAM", "ENCAPRESU", "Ajusta_Presu", 3, 2)
   Call CREACAMPO("PRESUCAM", "ENCAPRESU", "Vended_Presu", 3, 2)
   Call CREACAMPO("PRESUCAM", "ENCAPRESU", "ComiVend_Presu", 7, 8)
   Call CREACAMPO("PRESUCAM", "ENCAPRESU", "Anti_Presu", 6, 4)
   Call CREACAMPO("PRESUCAM", "ENCAPRESU", "Cuot_Presu", 3, 2)
   Call CREACAMPO("PRESUCAM", "ENCAPRESU", "Inter_Presu", 6, 4)
   Call CREACAMPO("PRESUCAM", "ENCAPRESU", "Plazo_Entre", 3, 2)
   Call CREACAMPO("PRESUCAM", "ENCAPRESU", "Plazo_Pago", 3, 2)
   '
End Sub
'
Sub CREATABLA_STOCKORIGI()
   '
   Call CREACAMPO("STOCKORIGI", "DETASTOCK", "C_Com", 10, 10)
   Call CREACAMPO("STOCKORIGI", "DETASTOCK", "N_Com", 4, 0)
   Call CREACAMPO("STOCKORIGI", "DETASTOCK", "C_Tmv", 10, 10)
   Call CREACAMPO("STOCKORIGI", "DETASTOCK", "Fecha_Mov", 8, 0)
   Call CREACAMPO("STOCKORIGI", "DETASTOCK", "C_Sig", 10, 1)
   Call CREACAMPO("STOCKORIGI", "DETASTOCK", "C_Dep", 10, 2)
   Call CREACAMPO("STOCKORIGI", "DETASTOCK", "C_Art", 10, 16)
   Call CREACAMPO("STOCKORIGI", "DETASTOCK", "Canti", 7, 8)
   Call CREACAMPO("STOCKORIGI", "DETASTOCK", "N_Trn", 4, 0)
   '
   Call CREACAMPO("STOCKORIGI", "ENCASTOCK", "C_Com", 10, 10)
   Call CREACAMPO("STOCKORIGI", "ENCASTOCK", "N_Com", 4, 0)
   Call CREACAMPO("STOCKORIGI", "ENCASTOCK", "C_Tmv", 10, 10)
   Call CREACAMPO("STOCKORIGI", "ENCASTOCK", "FechaMov", 8, 0)
   Call CREACAMPO("STOCKORIGI", "ENCASTOCK", "C_Ente", 10, 16)
   Call CREACAMPO("STOCKORIGI", "ENCASTOCK", "C_Dep", 10, 2)
   Call CREACAMPO("STOCKORIGI", "ENCASTOCK", "N_Trn", 4, 0)
   Call CREACAMPO("STOCKORIGI", "ENCASTOCK", "Descrip", 10, 64)
   '
End Sub
'
Sub CREATABLA_STOCKS()
   '
   Call CREACAMPO("STOCKS", "BALAREQMO", "CodiCompro", 10, 16)
   Call CREACAMPO("STOCKS", "BALAREQMO", "NumeCompro", 4, 4)
   Call CREACAMPO("STOCKS", "BALAREQMO", "CodSucu", 4, 4)
   Call CREACAMPO("STOCKS", "BALAREQMO", "NuOrItem", 4, 4)
   Call CREACAMPO("STOCKS", "BALAREQMO", "IdenReque", 10, 24)
   Call CREACAMPO("STOCKS", "BALAREQMO", "Cod_Inte", 3, 2)
   Call CREACAMPO("STOCKS", "BALAREQMO", "IdenLote", 10, 32)
   Call CREACAMPO("STOCKS", "BALAREQMO", "DoCompra", 10, 32)
   Call CREACAMPO("STOCKS", "BALAREQMO", "NuCert_Apro", 4, 4)
   Call CREACAMPO("STOCKS", "BALAREQMO", "Posicion", 10, 24)
   Call CREACAMPO("STOCKS", "BALAREQMO", "Cantidad", 7, 8)
   Call CREACAMPO("STOCKS", "BALAREQMO", "CodEnvase", 10, 64)
   Call CREACAMPO("STOCKS", "BALAREQMO", "Kilo_Reque", 7, 8)
   Call CREACAMPO("STOCKS", "BALAREQMO", "Kilo_Asig", 7, 8)
   Call CREACAMPO("STOCKS", "BALAREQMO", "Kilo_Saldo", 7, 8)
   Call CREACAMPO("STOCKS", "BALAREQMO", "NuLisPick", 4, 4)
   Call CREACAMPO("STOCKS", "BALAREQMO", "DeLisPick", 10, 64)
   Call CREACAMPO("STOCKS", "BALAREQMO", "Kilo_Consu", 7, 8)
   Call CREACAMPO("STOCKS", "BALAREQMO", "FeReMovi", 8, 8)
   Call CREACAMPO("STOCKS", "BALAREQMO", "FechMov", 8, 8)
   Call CREACAMPO("STOCKS", "BALAREQMO", "MarqAsi", 3, 2)
   Call CREACAMPO("STOCKS", "BALAREQMO", "CreadoPor", 10, 4)
   Call CREACAMPO("STOCKS", "BALAREQMO", "Destino", 10, 255)
   Call CREACAMPO("STOCKS", "BALAREQMO", "Ident_Asig", 10, 32)
   Call CREACAMPO("STOCKS", "BALAREQMO", "DepoMovi", 3, 1)
   Call CREACAMPO("STOCKS", "BALAREQMO", "Nume_Cli", 4, 4)
   Call CREACAMPO("STOCKS", "BALAREQMO", "Kilo_Aped", 7, 8)
   '
   Call CREACAMPO("STOCKS", "BOLRECEP", "NumeBoRe", 6, 4, 1)
   Call CREACAMPO("STOCKS", "BOLRECEP", "Descrip", 10, 64)
   Call CREACAMPO("STOCKS", "BOLRECEP", "IdenLote", 10, 12, 1)
   Call CREACAMPO("STOCKS", "BOLRECEP", "FechRecep", 8, 0)
   Call CREACAMPO("STOCKS", "BOLRECEP", "Cod_Inte", 3, 0, 1)
   Call CREACAMPO("STOCKS", "BOLRECEP", "Cod_Exte", 10, 24)
   Call CREACAMPO("STOCKS", "BOLRECEP", "DescripItem", 10, 64)
   Call CREACAMPO("STOCKS", "BOLRECEP", "Unidad", 10, 4)
   Call CREACAMPO("STOCKS", "BOLRECEP", "CantNom", 7, 8)
   Call CREACAMPO("STOCKS", "BOLRECEP", "M_Largo", 6, 0)
   Call CREACAMPO("STOCKS", "BOLRECEP", "M_Ancho", 6, 0)
   Call CREACAMPO("STOCKS", "BOLRECEP", "DepoEnt", 3, 1)
   Call CREACAMPO("STOCKS", "BOLRECEP", "UbicFis", 10, 32)
   Call CREACAMPO("STOCKS", "BOLRECEP", "NumProv", 3, 2, 1)
   Call CREACAMPO("STOCKS", "BOLRECEP", "RaSo_Prov", 10, 64)
   Call CREACAMPO("STOCKS", "BOLRECEP", "LoteCol_Prov", 10, 16)
   Call CREACAMPO("STOCKS", "BOLRECEP", "CerCal_Prov", 10, 24)
   Call CREACAMPO("STOCKS", "BOLRECEP", "NumRem_Prov", 10, 16)
   Call CREACAMPO("STOCKS", "BOLRECEP", "NumFac_Prov", 10, 16)
   Call CREACAMPO("STOCKS", "BOLRECEP", "NumConVig", 10, 24)
   Call CREACAMPO("STOCKS", "BOLRECEP", "NumOC", 4, 4, 1)
   Call CREACAMPO("STOCKS", "BOLRECEP", "Nume_Aten", 4, 4)
   Call CREACAMPO("STOCKS", "BOLRECEP", "IContab", 3, 2)
   Call CREACAMPO("STOCKS", "BOLRECEP", "MoneVaUni", 10, 4)
   Call CREACAMPO("STOCKS", "BOLRECEP", "Valo_Unit", 6, 4)
   Call CREACAMPO("STOCKS", "BOLRECEP", "CanToApro", 7, 8)
   Call CREACAMPO("STOCKS", "BOLRECEP", "CanToRech", 7, 8)
   Call CREACAMPO("STOCKS", "BOLRECEP", "CanToFalta", 7, 8)
   Call CREACAMPO("STOCKS", "BOLRECEP", "FechLibe", 8, 0)
   Call CREACAMPO("STOCKS", "BOLRECEP", "NumeOrdSer", 6, 0)
   Call CREACAMPO("STOCKS", "BOLRECEP", "Observa", 12, 0)
   Call CREACAMPO("STOCKS", "BOLRECEP", "Status", 3, 0)
   Call CREACAMPO("STOCKS", "BOLRECEP", "UsuApro", 10, 32)
   Call CREACAMPO("STOCKS", "BOLRECEP", "FeHoApro", 8, 0)
   Call CREACAMPO("STOCKS", "BOLRECEP", "Marca", 10, 32)
   Call CREACAMPO("STOCKS", "BOLRECEP", "NumeParte", 10, 32)
   Call CREACAMPO("STOCKS", "BOLRECEP", "FechFab", 8, 0)
   Call CREACAMPO("STOCKS", "BOLRECEP", "Plazo_Uso", 8, 8)
   Call CREACAMPO("STOCKS", "BOLRECEP", "Traz_Cli", 10, 32)
   Call CREACAMPO("STOCKS", "BOLRECEP", "Cod_Cli", 10, 32)
   Call CREACAMPO("STOCKS", "BOLRECEP", "Tipo", 10, 32)
   Call CREACAMPO("STOCKS", "BOLRECEP", "Carac", 10, 32)
   Call CREACAMPO("STOCKS", "BOLRECEP", "DepoUbic", 3, 0)
   Call CREACAMPO("STOCKS", "BOLRECEP", "UbicFis", 10, 32)
   Call CREACAMPO("STOCKS", "BOLRECEP", "AnotaOrf", 12, 0)
   Call CREACAMPO("STOCKS", "BOLRECEP", "MotiRecha", 3, 0)
   Call CREACAMPO("STOCKS", "BOLRECEP", "ObseRecha", 12, 0)
   Call CREACAMPO("STOCKS", "BOLRECEP", "Destirecha", 3, 0)
   Call CREACAMPO("STOCKS", "BOLRECEP", "Via_Entrega", 3, 0)

   '
   Call CREACAMPO("STOCKS", "DETACARTEM", "Num_Os", 10, 16)
   Call CREACAMPO("STOCKS", "DETACARTEM", "Num_Entre", 10, 16)
   Call CREACAMPO("STOCKS", "DETACARTEM", "FechEntre", 8, 0)
   Call CREACAMPO("STOCKS", "DETACARTEM", "NorItCa", 4, 0)
   Call CREACAMPO("STOCKS", "DETACARTEM", "Cod_Parte", 10, 32)
   Call CREACAMPO("STOCKS", "DETACARTEM", "Num_Inter", 4, 0)
   Call CREACAMPO("STOCKS", "DETACARTEM", "Pie_Medio", 4, 0)
   Call CREACAMPO("STOCKS", "DETACARTEM", "Pie_Pedid", 4, 0)
   Call CREACAMPO("STOCKS", "DETACARTEM", "Pie_Recu", 4, 0)
   Call CREACAMPO("STOCKS", "DETACARTEM", "Pie_Tota", 4, 0)
   Call CREACAMPO("STOCKS", "DETACARTEM", "Pie_Enviar", 4, 0)
   Call CREACAMPO("STOCKS", "DETACARTEM", "Cod_Exte", 10, 24)
   Call CREACAMPO("STOCKS", "DETACARTEM", "Cod_Inte", 3, 0)
   Call CREACAMPO("STOCKS", "DETACARTEM", "Cant_Entre", 4, 0)
   Call CREACAMPO("STOCKS", "DETACARTEM", "SubCod", 10, 8)
   Call CREACAMPO("STOCKS", "DETACARTEM", "Dopricor", 10, 12)
   Call CREACAMPO("STOCKS", "DETACARTEM", "Doprilar", 10, 24)
   Call CREACAMPO("STOCKS", "DETACARTEM", "NuorItem", 3, 2)
   Call CREACAMPO("STOCKS", "DETACARTEM", "FechEmis", 8, 0)
   Call CREACAMPO("STOCKS", "DETACARTEM", "IdCartem", 10, 48)
   '
   Call CREACAMPO("STOCKS", "DETAOSER", "NuOrdSe", 4, 4)
   Call CREACAMPO("STOCKS", "DETAOSER", "NuOrden", 3, 2)
   Call CREACAMPO("STOCKS", "DETAOSER", "Modelo", 3, 2)
   Call CREACAMPO("STOCKS", "DETAOSER", "NumeroSerie", 10, 32)
   Call CREACAMPO("STOCKS", "DETAOSER", "HorasFunc", 3, 2)
   '
   Call CREACAMPO("STOCKS", "ENCACARTEM", "FechEntre", 8, 0)
   Call CREACAMPO("STOCKS", "ENCACARTEM", "Num_Os", 10, 16)
   Call CREACAMPO("STOCKS", "ENCACARTEM", "Num_Entre", 10, 16)
   '
   Call CREACAMPO("STOCKS", "ENCAOSER", "NuOrdSe", 4, 4)
   Call CREACAMPO("STOCKS", "ENCAOSER", "FeOrdSe", 8, 0)
   Call CREACAMPO("STOCKS", "ENCAOSER", "Nume_Cli", 3, 2)
   Call CREACAMPO("STOCKS", "ENCAOSER", "TipoOrdSe", 3, 2)
   Call CREACAMPO("STOCKS", "ENCAOSER", "Observa", 12, 0)
   '
   Call CREACAMPO("STOCKS", "ENCAREMI", "DoSecLar", 10, 24)
   Call CREACAMPO("STOCKS", "ENCAREMI", "DoPriLar", 10, 24)
   Call CREACAMPO("STOCKS", "ENCAREMI", "Descripcion", 10, 64)
   Call CREACAMPO("STOCKS", "ENCAREMI", "FechMov", 8, 0)
   Call CREACAMPO("STOCKS", "ENCAREMI", "Propiedad", 10, 64)
   Call CREACAMPO("STOCKS", "ENCAREMI", "Concepto", 10, 64)
   Call CREACAMPO("STOCKS", "ENCAREMI", "Producto", 12, 0)
   Call CREACAMPO("STOCKS", "ENCAREMI", "Bultos", 3, 5)
   Call CREACAMPO("STOCKS", "ENCAREMI", "Origen", 3, 3)
   Call CREACAMPO("STOCKS", "ENCAREMI", "Deposito", 3, 3)
   Call CREACAMPO("STOCKS", "ENCAREMI", "Nume_Clie", 3, 4)
   Call CREACAMPO("STOCKS", "ENCAREMI", "Nume_Prov", 3, 4)
   Call CREACAMPO("STOCKS", "ENCAREMI", "CantPalet", 3, 4)
   '
   Call CREACAMPO("STOCKS", "LISEMPAQ", "NumPallet", 3, 0)
   Call CREACAMPO("STOCKS", "LISEMPAQ", "NuOrdPallet", 3, 0)
   Call CREACAMPO("STOCKS", "LISEMPAQ", "DesdeCaja", 3, 0)
   Call CREACAMPO("STOCKS", "LISEMPAQ", "HastaCaja", 3, 0)
   Call CREACAMPO("STOCKS", "LISEMPAQ", "CantiCajas", 3, 0)
   Call CREACAMPO("STOCKS", "LISEMPAQ", "Codint", 3, 0)
   Call CREACAMPO("STOCKS", "LISEMPAQ", "PiezasXCaja", 4, 0)
   Call CREACAMPO("STOCKS", "LISEMPAQ", "PiezasTotal", 4, 0)
   Call CREACAMPO("STOCKS", "LISEMPAQ", "M3Caja", 6, 0)
   Call CREACAMPO("STOCKS", "LISEMPAQ", "M3Total", 7, 0)
   Call CREACAMPO("STOCKS", "LISEMPAQ", "KiloNetoCaja", 6, 0)
   Call CREACAMPO("STOCKS", "LISEMPAQ", "KiloNetoTot", 7, 0)
   Call CREACAMPO("STOCKS", "LISEMPAQ", "KiloBrutoCaja", 6, 0)
   Call CREACAMPO("STOCKS", "LISEMPAQ", "KiloBrutoTot", 7, 0)
   Call CREACAMPO("STOCKS", "LISEMPAQ", "Marborra", 1, 0)
   Call CREACAMPO("STOCKS", "LISEMPAQ", "NumLtaEmp", 4, 0)
   Call CREACAMPO("STOCKS", "LISEMPAQ", "NroPed", 4, 0)
   '
   Call CREACAMPO("STOCKS", "LOTINGRE", "Cod_Inte", 3, 2, 1)
   Call CREACAMPO("STOCKS", "LOTINGRE", "IdenLote", 10, 32)
   Call CREACAMPO("STOCKS", "LOTINGRE", "DocuIngre", 10, 24)
   Call CREACAMPO("STOCKS", "LOTINGRE", "DocuOrig", 10, 24)
   Call CREACAMPO("STOCKS", "LOTINGRE", "FechAlta", 8, 0)
   Call CREACAMPO("STOCKS", "LOTINGRE", "Cantidad", 7, 0)
   Call CREACAMPO("STOCKS", "LOTINGRE", "CanAlta", 7, 0)
   Call CREACAMPO("STOCKS", "LOTINGRE", "CanConsu", 7, 0)
   Call CREACAMPO("STOCKS", "LOTINGRE", "CanReser", 7, 0)
   Call CREACAMPO("STOCKS", "LOTINGRE", "CanSaldo", 7, 0)
   Call CREACAMPO("STOCKS", "LOTINGRE", "SalNeto", 7, 0)
   Call CREACAMPO("STOCKS", "LOTINGRE", "KilBruto", 7, 0)
   Call CREACAMPO("STOCKS", "LOTINGRE", "CodEnvase", 10, 64)
   Call CREACAMPO("STOCKS", "LOTINGRE", "AnchoMili", 7, 0)
   Call CREACAMPO("STOCKS", "LOTINGRE", "MetroIni", 7, 0)
   Call CREACAMPO("STOCKS", "LOTINGRE", "SalMetro", 7, 0)
   Call CREACAMPO("STOCKS", "LOTINGRE", "FechFab", 8, 0)
   Call CREACAMPO("STOCKS", "LOTINGRE", "FechVenc", 8, 0)
   Call CREACAMPO("STOCKS", "LOTINGRE", "DepoUbic", 3, 0)
   Call CREACAMPO("STOCKS", "LOTINGRE", "UbicFis", 10, 32)
   Call CREACAMPO("STOCKS", "LOTINGRE", "Marca", 10, 32)
   Call CREACAMPO("STOCKS", "LOTINGRE", "Origen", 10, 255)
   Call CREACAMPO("STOCKS", "LOTINGRE", "LoteCol_Prov", 10, 16)
   Call CREACAMPO("STOCKS", "LOTINGRE", "Costo_Rep", 7, 8)
   Call CREACAMPO("STOCKS", "LOTINGRE", "Moneco_Rep", 3, 2)
   Call CREACAMPO("STOCKS", "LOTINGRE", "Fechco_Rep", 8, 8)
   Call CREACAMPO("STOCKS", "LOTINGRE", "ValoHis", 7, 8)
   Call CREACAMPO("STOCKS", "LOTINGRE", "ValoRep", 7, 8)
   Call CREACAMPO("STOCKS", "LOTINGRE", "Libera", 10, 16)
   Call CREACAMPO("STOCKS", "LOTINGRE", "TiMate", 3, 0, 1)
   Call CREACAMPO("STOCKS", "LOTINGRE", "Cod_Desti", 10, 16)
   Call CREACAMPO("STOCKS", "LOTINGRE", "MarInve", 3, 0)
   Call CREACAMPO("STOCKS", "LOTINGRE", "Potencia", 7, 0)
   Call CREACAMPO("STOCKS", "LOTINGRE", "Tara", 7, 0)
   Call CREACAMPO("STOCKS", "LOTINGRE", "Precinto", 10, 16)
   Call CREACAMPO("STOCKS", "LOTINGRE", "SalMeCua", 6, 0)
   Call CREACAMPO("STOCKS", "LOTINGRE", "Ref_Apro", 12, 0)
   Call CREACAMPO("STOCKS", "LOTINGRE", "NuCert_Apro", 4, 4)
   Call CREACAMPO("STOCKS", "LOTINGRE", "Usu_Apro", 10, 64)
   Call CREACAMPO("STOCKS", "LOTINGRE", "Fech_Apro", 8, 0)
   Call CREACAMPO("STOCKS", "LOTINGRE", "OperProd", 10, 32)
   Call CREACAMPO("STOCKS", "LOTINGRE", "TurnoProd", 10, 32)
   Call CREACAMPO("STOCKS", "LOTINGRE", "Nro_Analisis", 10, 32)
   Call CREACAMPO("STOCKS", "LOTINGRE", "EtiqxCaja", 4, 0)
   Call CREACAMPO("STOCKS", "LOTINGRE", "RolxCaja", 4, 0)
   Call CREACAMPO("STOCKS", "LOTINGRE", "EtiqxRollo", 4, 0)
   Call CREACAMPO("STOCKS", "LOTINGRE", "CodCarpImp", 10, 12)
   '
   Call CREACAMPO("STOCKS", "MACONSIG", "Femi_Ocom", 8, 0)
   Call CREACAMPO("STOCKS", "MACONSIG", "Npro_Ocom", 3, 2)
   Call CREACAMPO("STOCKS", "MACONSIG", "RaSo_Prov", 10, 64)
   Call CREACAMPO("STOCKS", "MACONSIG", "Secc_Ejec", 6, 4)
   Call CREACAMPO("STOCKS", "MACONSIG", "Deno_Secc", 10, 64)
   Call CREACAMPO("STOCKS", "MACONSIG", "Nro_Orfa", 10, 12)
   Call CREACAMPO("STOCKS", "MACONSIG", "Nro_Ocom", 10, 12)
   Call CREACAMPO("STOCKS", "MACONSIG", "Cod_Inte", 3, 0)
   Call CREACAMPO("STOCKS", "MACONSIG", "Cod_Exte", 10, 24)
   Call CREACAMPO("STOCKS", "MACONSIG", "Descritem", 10, 64)
   Call CREACAMPO("STOCKS", "MACONSIG", "Lote_Ocom", 10, 16)
   Call CREACAMPO("STOCKS", "MACONSIG", "Unidad", 10, 4)
   Call CREACAMPO("STOCKS", "MACONSIG", "Cant_Asig", 7, 8)
   Call CREACAMPO("STOCKS", "MACONSIG", "Cant_Entreg", 7, 8)
   Call CREACAMPO("STOCKS", "MACONSIG", "Fecha_Entreg", 8, 0)
   Call CREACAMPO("STOCKS", "MACONSIG", "NRemi_Consig", 10, 16)
   Call CREACAMPO("STOCKS", "MACONSIG", "Valo_Unit", 7, 8)
   Call CREACAMPO("STOCKS", "MACONSIG", "Moneda", 3, 2)
   Call CREACAMPO("STOCKS", "MACONSIG", "Valo_Tot", 7, 8)
   Call CREACAMPO("STOCKS", "MACONSIG", "Cant_Rendid", 7, 8)
   Call CREACAMPO("STOCKS", "MACONSIG", "Saldo_Pend", 7, 8)
   Call CREACAMPO("STOCKS", "MACONSIG", "Observa", 12, 0)
   Call CREACAMPO("STOCKS", "MACONSIG", "Cant_Devu", 7, 0)
   Call CREACAMPO("STOCKS", "MACONSIG", "MarBorra", 3, 0)
   '
   Call CREACAMPO("STOCKS", "MASUMCLI", "Nume_Cli", 3, 0)
   Call CREACAMPO("STOCKS", "MASUMCLI", "NumeIng", 4, 0, 1)
   Call CREACAMPO("STOCKS", "MASUMCLI", "Descripcion", 10, 64)
   Call CREACAMPO("STOCKS", "MASUMCLI", "Yacimiento", 10, 64)
   Call CREACAMPO("STOCKS", "MASUMCLI", "Referencia", 10, 64)
   Call CREACAMPO("STOCKS", "MASUMCLI", "TarAcordadas", 12, 0)
   Call CREACAMPO("STOCKS", "MASUMCLI", "Observaciones", 12, 0)
   Call CREACAMPO("STOCKS", "MASUMCLI", "DepInterviniente", 12, 0)
   Call CREACAMPO("STOCKS", "MASUMCLI", "FechRecep", 8, 0)
   Call CREACAMPO("STOCKS", "MASUMCLI", "UsuApFor_1", 10, 64)
   Call CREACAMPO("STOCKS", "MASUMCLI", "UsuApFor_2", 10, 64)
   Call CREACAMPO("STOCKS", "MASUMCLI", "UsuApFor_3", 10, 64)
   Call CREACAMPO("STOCKS", "MASUMCLI", "UsuApFor_4", 10, 64)
   Call CREACAMPO("STOCKS", "MASUMCLI", "EstadoRec", 12, 0)
   Call CREACAMPO("STOCKS", "MASUMCLI", "Ubicacion", 10, 64)
   Call CREACAMPO("STOCKS", "MASUMCLI", "FechVerif", 8, 0)
   Call CREACAMPO("STOCKS", "MASUMCLI", "Trab_Efect", 12, 0)
   Call CREACAMPO("STOCKS", "MASUMCLI", "FechTrabEfect", 8, 0)
   Call CREACAMPO("STOCKS", "MASUMCLI", "EstEntrega", 12, 0)
   Call CREACAMPO("STOCKS", "MASUMCLI", "Status", 3, 0)
   Call CREACAMPO("STOCKS", "MASUMCLI", "FechEnt", 8, 0)
   Call CREACAMPO("STOCKS", "MASUMCLI", "Lentre_Cli", 10, 64)
   Call CREACAMPO("STOCKS", "MASUMCLI", "Patente", 10, 16)
   Call CREACAMPO("STOCKS", "MASUMCLI", "Conductor", 10, 48)
   Call CREACAMPO("STOCKS", "MASUMCLI", "Piva_Cli", 10, 48)
   Call CREACAMPO("STOCKS", "MASUMCLI", "Cuit_Cli", 10, 24)
   Call CREACAMPO("STOCKS", "MASUMCLI", "Nom_Tra", 10, 48)
   Call CREACAMPO("STOCKS", "MASUMCLI", "Fech_Remito", 8, 0)
   Call CREACAMPO("STOCKS", "MASUMCLI", "NumRemito", 10, 18)
   '
   Call CREACAMPO("STOCKS", "MOVIREPA", "FeReMovi", 8, 0)
   Call CREACAMPO("STOCKS", "MOVIREPA", "NumUsuar", 3, 0)
   Call CREACAMPO("STOCKS", "MOVIREPA", "IdenMovi", 10, 32)
   Call CREACAMPO("STOCKS", "MOVIREPA", "DeLaRemi", 12, 0)
   Call CREACAMPO("STOCKS", "MOVIREPA", "Descri_Lar", 12, 0)
   Call CREACAMPO("STOCKS", "MOVIREPA", "FechMov", 8, 0)
   Call CREACAMPO("STOCKS", "MOVIREPA", "Cod_Inte", 3, 0, 1)
   Call CREACAMPO("STOCKS", "MOVIREPA", "Cod_Exte", 10, 16)
   Call CREACAMPO("STOCKS", "MOVIREPA", "Descrip", 10, 64)
   Call CREACAMPO("STOCKS", "MOVIREPA", "IdenLote", 10, 16)
   Call CREACAMPO("STOCKS", "MOVIREPA", "LotConsum", 12, 0)
   Call CREACAMPO("STOCKS", "MOVIREPA", "PNL_Des", 10, 32)
   Call CREACAMPO("STOCKS", "MOVIREPA", "PLC_Has", 10, 32)
   Call CREACAMPO("STOCKS", "MOVIREPA", "CodiMovi", 10, 2, 1)
   Call CREACAMPO("STOCKS", "MOVIREPA", "DoPriCor", 10, 12)
   Call CREACAMPO("STOCKS", "MOVIREPA", "DoPriLar", 10, 24)
   Call CREACAMPO("STOCKS", "MOVIREPA", "DoSecCor", 10, 12)
   Call CREACAMPO("STOCKS", "MOVIREPA", "DoSecLar", 10, 24)
   Call CREACAMPO("STOCKS", "MOVIREPA", "NumRem_Cli", 10, 16)
   Call CREACAMPO("STOCKS", "MOVIREPA", "NuorItem", 3, 2)
   Call CREACAMPO("STOCKS", "MOVIREPA", "NuPedCli", 4, 4, 1)
   Call CREACAMPO("STOCKS", "MOVIREPA", "NumeOrdPed", 3, 2)
   Call CREACAMPO("STOCKS", "MOVIREPA", "ReferCor", 10, 64)
   Call CREACAMPO("STOCKS", "MOVIREPA", "CostoRep", 7, 0)
   Call CREACAMPO("STOCKS", "MOVIREPA", "ValoUnit", 7, 0)
   Call CREACAMPO("STOCKS", "MOVIREPA", "MoneVal", 3, 2)
   Call CREACAMPO("STOCKS", "MOVIREPA", "TipoCam", 7, 8)
   Call CREACAMPO("STOCKS", "MOVIREPA", "DepoMovi", 3, 0, 1)
   Call CREACAMPO("STOCKS", "MOVIREPA", "Nume_Clie", 3, 0)
   Call CREACAMPO("STOCKS", "MOVIREPA", "Sucu_Clie", 3, 0)
   Call CREACAMPO("STOCKS", "MOVIREPA", "Nume_Prov", 3, 0)
   Call CREACAMPO("STOCKS", "MOVIREPA", "CantIngre", 7, 0)
   Call CREACAMPO("STOCKS", "MOVIREPA", "CantSalid", 7, 0)
   Call CREACAMPO("STOCKS", "MOVIREPA", "CantSaldo", 7, 0)
   Call CREACAMPO("STOCKS", "MOVIREPA", "DocuOrig", 10, 24)
   Call CREACAMPO("STOCKS", "MOVIREPA", "ValoUnitIva", 7, 0)
   Call CREACAMPO("STOCKS", "MOVIREPA", "ValTotIt", 7, 0)
   Call CREACAMPO("STOCKS", "MOVIREPA", "ValTotItIva", 7, 0)
   Call CREACAMPO("STOCKS", "MOVIREPA", "VerificadoPor", 3, 0)
   Call CREACAMPO("STOCKS", "MOVIREPA", "FHor_Verif", 8, 0)
   Call CREACAMPO("STOCKS", "MOVIREPA", "CodCarImp", 10, 12)
   Call CREACAMPO("STOCKS", "MOVIREPA", "User_Tran", 10, 32)
   Call CREACAMPO("STOCKS", "MOVIREPA", "Causa", 3, 0)
   Call CREACAMPO("STOCKS", "MOVIREPA", "Verificado", 3, 1)
   Call CREACAMPO("STOCKS", "MOVIREPA", "OCOMPRA_RTOABIERTO", 10, 64)
   '
   Call CREACAMPO("STOCKS", "MOVISTO", "FeReMovi", 8, 0)
   Call CREACAMPO("STOCKS", "MOVISTO", "NumUsuar", 3, 0, 1)
   Call CREACAMPO("STOCKS", "MOVISTO", "IdenMovi", 10, 32)
   Call CREACAMPO("STOCKS", "MOVISTO", "DeLaRemi", 12, 0)
   Call CREACAMPO("STOCKS", "MOVISTO", "Descri_Lar", 12, 0)
   Call CREACAMPO("STOCKS", "MOVISTO", "FechMov", 8, 0, 1)
   Call CREACAMPO("STOCKS", "MOVISTO", "Cod_Inte", 3, 0, 1)
   Call CREACAMPO("STOCKS", "MOVISTO", "Cod_Exte", 10, 24)
   Call CREACAMPO("STOCKS", "MOVISTO", "Descrip", 10, 64)
   Call CREACAMPO("STOCKS", "MOVISTO", "IdenLote", 10, 16, 1)
   Call CREACAMPO("STOCKS", "MOVISTO", "LotConsum", 12, 0)
   Call CREACAMPO("STOCKS", "MOVISTO", "PNL_Des", 10, 32)
   Call CREACAMPO("STOCKS", "MOVISTO", "PLC_Has", 10, 32)
   Call CREACAMPO("STOCKS", "MOVISTO", "CodiMovi", 10, 2, 1)
   Call CREACAMPO("STOCKS", "MOVISTO", "DoPriCor", 10, 12, 1)
   Call CREACAMPO("STOCKS", "MOVISTO", "DoPriLar", 10, 24, 1)
   Call CREACAMPO("STOCKS", "MOVISTO", "DoSecCor", 10, 12, 1)
   Call CREACAMPO("STOCKS", "MOVISTO", "DoSecLar", 10, 24, 1)
   Call CREACAMPO("STOCKS", "MOVISTO", "DocuOrig", 10, 24)
   Call CREACAMPO("STOCKS", "MOVISTO", "NumRem_Cli", 10, 16)
   Call CREACAMPO("STOCKS", "MOVISTO", "NuorItem", 3, 2)
   Call CREACAMPO("STOCKS", "MOVISTO", "NuPedCli", 4, 4, 1)
   Call CREACAMPO("STOCKS", "MOVISTO", "NumeOrdPed", 3, 2)
   Call CREACAMPO("STOCKS", "MOVISTO", "ReferCor", 10, 64)
   Call CREACAMPO("STOCKS", "MOVISTO", "CostoRep", 7, 0)
   Call CREACAMPO("STOCKS", "MOVISTO", "ValoUnit", 7, 0)
   Call CREACAMPO("STOCKS", "MOVISTO", "ValoUnitIva", 7, 0)
   Call CREACAMPO("STOCKS", "MOVISTO", "ValTotIt", 7, 0)
   Call CREACAMPO("STOCKS", "MOVISTO", "ValTotItIva", 7, 0)
   Call CREACAMPO("STOCKS", "MOVISTO", "MoneVal", 3, 2)
   Call CREACAMPO("STOCKS", "MOVISTO", "TipoCam", 7, 8)
   Call CREACAMPO("STOCKS", "MOVISTO", "DepoMovi", 3, 0, 1)
   Call CREACAMPO("STOCKS", "MOVISTO", "Nume_Clie", 4, 0, 1)
   Call CREACAMPO("STOCKS", "MOVISTO", "Sucu_Clie", 3, 0)
   Call CREACAMPO("STOCKS", "MOVISTO", "Nume_Prov", 3, 0)
   Call CREACAMPO("STOCKS", "MOVISTO", "CantIngre", 7, 0, 1)
   Call CREACAMPO("STOCKS", "MOVISTO", "CantSalid", 7, 0, 1)
   Call CREACAMPO("STOCKS", "MOVISTO", "CantSaldo", 7, 0)
   Call CREACAMPO("STOCKS", "MOVISTO", "Verificado", 3, 1)
   Call CREACAMPO("STOCKS", "MOVISTO", "Causa", 3, 0)
   Call CREACAMPO("STOCKS", "MOVISTO", "User_Tran", 10, 32)
   Call CREACAMPO("STOCKS", "MOVISTO", "CodCarpImp", 10, 12)
   Call CREACAMPO("STOCKS", "MOVISTO", "VerificadoPor", 3, 0)
   Call CREACAMPO("STOCKS", "MOVISTO", "FHor_Verif", 8, 0)
   Call CREACAMPO("STOCKS", "MOVISTO", "TipoTran", 3, 0)
   Call CREACAMPO("STOCKS", "MOVISTO", "OCOMPRA_RTOABIERTO", 10, 64)
   Call CREACAMPO("STOCKS", "MOVISTO", "NroEnvio", 4, 0)
   '
   Call CREACAMPO("STOCKS", "MOVIVIE", "FeReMovi", 8, 0)
   Call CREACAMPO("STOCKS", "MOVIVIE", "NumUsuar", 3, 0)
   Call CREACAMPO("STOCKS", "MOVIVIE", "IdenMovi", 10, 32)
   Call CREACAMPO("STOCKS", "MOVIVIE", "DeLaRemi", 12, 0)
   Call CREACAMPO("STOCKS", "MOVIVIE", "Descri_Lar", 12, 0)
   Call CREACAMPO("STOCKS", "MOVIVIE", "FechMov", 8, 0)
   Call CREACAMPO("STOCKS", "MOVIVIE", "Cod_Inte", 3, 0, 1)
   Call CREACAMPO("STOCKS", "MOVIVIE", "Cod_Exte", 10, 24)
   Call CREACAMPO("STOCKS", "MOVIVIE", "Descrip", 10, 64)
   Call CREACAMPO("STOCKS", "MOVIVIE", "IdenLote", 10, 16)
   Call CREACAMPO("STOCKS", "MOVIVIE", "LotConsum", 12, 0)
   Call CREACAMPO("STOCKS", "MOVIVIE", "PNL_Des", 10, 32)
   Call CREACAMPO("STOCKS", "MOVIVIE", "PLC_Has", 10, 32)
   Call CREACAMPO("STOCKS", "MOVIVIE", "CodiMovi", 10, 2, 1)
   Call CREACAMPO("STOCKS", "MOVIVIE", "DoPriCor", 10, 12)
   Call CREACAMPO("STOCKS", "MOVIVIE", "DoPriLar", 10, 24)
   Call CREACAMPO("STOCKS", "MOVIVIE", "DosecCor", 10, 12)
   Call CREACAMPO("STOCKS", "MOVIVIE", "DosecLar", 10, 24)
   Call CREACAMPO("STOCKS", "MOVIVIE", "NumRem_Cli", 10, 16)
   Call CREACAMPO("STOCKS", "MOVIVIE", "NuorItem", 3, 2)
   Call CREACAMPO("STOCKS", "MOVIVIE", "NuPedCli", 4, 4, 1)
   Call CREACAMPO("STOCKS", "MOVIVIE", "NumeOrdPed", 3, 2)
   Call CREACAMPO("STOCKS", "MOVIVIE", "ReferCor", 10, 64)
   Call CREACAMPO("STOCKS", "MOVIVIE", "CostoRep", 7, 0)
   Call CREACAMPO("STOCKS", "MOVIVIE", "ValoUnit", 7, 0)
   Call CREACAMPO("STOCKS", "MOVIVIE", "MoneVal", 3, 2)
   Call CREACAMPO("STOCKS", "MOVIVIE", "TipoCam", 7, 8)
   Call CREACAMPO("STOCKS", "MOVIVIE", "DepoMovi", 3, 0, 1)
   Call CREACAMPO("STOCKS", "MOVIVIE", "Nume_Clie", 3, 0)
   Call CREACAMPO("STOCKS", "MOVIVIE", "Sucu_Clie", 3, 0)
   Call CREACAMPO("STOCKS", "MOVIVIE", "Nume_Prov", 3, 0)
   Call CREACAMPO("STOCKS", "MOVIVIE", "CantIngre", 7, 0)
   Call CREACAMPO("STOCKS", "MOVIVIE", "CantSalid", 7, 0)
   Call CREACAMPO("STOCKS", "MOVIVIE", "CantSaldo", 7, 0)
   Call CREACAMPO("STOCKS", "MOVIVIE", "OCOMPRA_RTOABIERTO", 10, 64)
   '
   Call CREACAMPO("STOCKS", "RESERVA", "Cod_Inte", 3, 2, 1)
   Call CREACAMPO("STOCKS", "RESERVA", "Cod_Exte", 10, 16)
   Call CREACAMPO("STOCKS", "RESERVA", "Descrip", 10, 64)
   Call CREACAMPO("STOCKS", "RESERVA", "IdSubOr", 10, 12, 1)
   Call CREACAMPO("STOCKS", "RESERVA", "NuOrItem", 3, 0)
   Call CREACAMPO("STOCKS", "RESERVA", "Status", 3, 0)
   Call CREACAMPO("STOCKS", "RESERVA", "FechRes", 8, 0)
   Call CREACAMPO("STOCKS", "RESERVA", "CantRes", 7, 8, 1)
   Call CREACAMPO("STOCKS", "RESERVA", "CantCons", 7, 8, 1)
   Call CREACAMPO("STOCKS", "RESERVA", "Saldo_Entre", 7, 8)
   Call CREACAMPO("STOCKS", "RESERVA", "IdenLote", 10, 16, 1)
   Call CREACAMPO("STOCKS", "RESERVA", "Deposito", 3, 2)
   Call CREACAMPO("STOCKS", "RESERVA", "UbicaLot", 10, 32)
   Call CREACAMPO("STOCKS", "RESERVA", "UsoUnit", 7, 8)
   Call CREACAMPO("STOCKS", "RESERVA", "Observa", 10, 128)
   Call CREACAMPO("STOCKS", "RESERVA", "PNL_Des", 10, 24)
   Call CREACAMPO("STOCKS", "RESERVA", "PLC_Has", 10, 24)
   Call CREACAMPO("STOCKS", "RESERVA", "Nume_Solint", 4, 0)
   Call CREACAMPO("STOCKS", "RESERVA", "Nume_Ocom", 4, 0)
   Call CREACAMPO("STOCKS", "RESERVA", "Marborra", 3, 0)
   '
   Call CREACAMPO("STOCKS", "SALSTOCK", "CodiInte", 3, 2)
   Call CREACAMPO("STOCKS", "SALSTOCK", "VeSalSto", 11, 256)
   '
   Call CREACAMPO("STOCKS", "TRANMOVI", "CodiCompro", 10, 5)
   Call CREACAMPO("STOCKS", "TRANMOVI", "NumeCompro", 4, 4)
   Call CREACAMPO("STOCKS", "TRANMOVI", "CodSucu", 4, 4)
   Call CREACAMPO("STOCKS", "TRANMOVI", "NuOrItem", 4, 4)
   Call CREACAMPO("STOCKS", "TRANMOVI", "Cod_Exte", 10, 24)
   Call CREACAMPO("STOCKS", "TRANMOVI", "Descripcion", 10, 64)
   Call CREACAMPO("STOCKS", "TRANMOVI", "Descridos", 10, 255)
   Call CREACAMPO("STOCKS", "TRANMOVI", "IdenLote", 10, 32)
   Call CREACAMPO("STOCKS", "TRANMOVI", "Cantidad", 7, 8)
   Call CREACAMPO("STOCKS", "TRANMOVI", "CodEnvase", 10, 64)
   Call CREACAMPO("STOCKS", "TRANMOVI", "CanKilos", 7, 8)
   Call CREACAMPO("STOCKS", "TRANMOVI", "FeReMovi", 8, 8)
   Call CREACAMPO("STOCKS", "TRANMOVI", "FechMov", 8, 8)
   Call CREACAMPO("STOCKS", "TRANMOVI", "TipoTrans", 3, 2)
   Call CREACAMPO("STOCKS", "TRANMOVI", "Status", 3, 2)
   Call CREACAMPO("STOCKS", "TRANMOVI", "CreadoPor", 10, 4)
   Call CREACAMPO("STOCKS", "TRANMOVI", "Destino", 10, 255)
   Call CREACAMPO("STOCKS", "TRANMOVI", "Ingresado", 3, 0)
   Call CREACAMPO("STOCKS", "TRANMOVI", "Ident_Asig", 10, 32, 1)
   '
   Call CREACAMPO("STOCKS", "TRANRECEP", "CodiCompro", 10, 5)
   Call CREACAMPO("STOCKS", "TRANRECEP", "NumeCompro", 4, 4)
   Call CREACAMPO("STOCKS", "TRANRECEP", "NuOrItem", 4, 4)
   Call CREACAMPO("STOCKS", "TRANRECEP", "DocuIngre", 10, 24)
   Call CREACAMPO("STOCKS", "TRANRECEP", "Cod_Exte", 10, 24)
   Call CREACAMPO("STOCKS", "TRANRECEP", "IdenLote", 10, 32)
   Call CREACAMPO("STOCKS", "TRANRECEP", "Descripcion", 10, 64)
   Call CREACAMPO("STOCKS", "TRANRECEP", "Cantidad", 7, 8)
   Call CREACAMPO("STOCKS", "TRANRECEP", "CodEnvase", 10, 64)
   Call CREACAMPO("STOCKS", "TRANRECEP", "CanKilos", 7, 8)
   Call CREACAMPO("STOCKS", "TRANRECEP", "NumRemProv", 10, 24)
   Call CREACAMPO("STOCKS", "TRANRECEP", "FeReMovi", 8, 8)
   Call CREACAMPO("STOCKS", "TRANRECEP", "FechMov", 8, 8)
   Call CREACAMPO("STOCKS", "TRANRECEP", "FechVenc", 8, 8)
   Call CREACAMPO("STOCKS", "TRANRECEP", "FechFab", 8, 8)
   Call CREACAMPO("STOCKS", "TRANRECEP", "Origen", 10, 255)
   Call CREACAMPO("STOCKS", "TRANRECEP", "AproRech", 10, 1)
   Call CREACAMPO("STOCKS", "TRANRECEP", "Status", 3, 2)
   Call CREACAMPO("STOCKS", "TRANRECEP", "Ingresado", 3, 0)
   '
   Call CREACAMPO("STOCKS", "UBIPRE", "DepoUbic", 3, 0, 1)
   Call CREACAMPO("STOCKS", "UBIPRE", "Codiubic", 10, 16, 1)
   Call CREACAMPO("STOCKS", "UBIPRE", "Referencia", 10, 64)
   Call CREACAMPO("STOCKS", "UBIPRE", "Status", 3, 2)
   Call CREACAMPO("STOCKS", "UBIPRE", "MarBorra", 3, 0)
   '
   Call CREACAMPO("STOCKS", "TRAZARECE", "Cod_Inte", 3, 0, 1)
   Call CREACAMPO("STOCKS", "TRAZARECE", "Cod_Exte", 10, 24)
   Call CREACAMPO("STOCKS", "TRAZARECE", "Descrip", 10, 64)
   Call CREACAMPO("STOCKS", "TRAZARECE", "IdenLote", 10, 12)
   Call CREACAMPO("STOCKS", "TRAZARECE", "NumeroSerie", 3, 2)
   Call CREACAMPO("STOCKS", "TRAZARECE", "Cfamilia", 10, 24)
   Call CREACAMPO("STOCKS", "TRAZARECE", "Origen", 10, 24)
   Call CREACAMPO("STOCKS", "TRAZARECE", "Destino", 10, 24)
   Call CREACAMPO("STOCKS", "TRAZARECE", "DestinoConsumo", 10, 24)
   '
End Sub
'
Sub CREATABLA_STOCKVIE()
   '
   Call CREACAMPO("STOCKVIE", "MOVISTO", "FechMov", 8, 0)
   Call CREACAMPO("STOCKVIE", "MOVISTO", "Cod_Inte", 3, 0)
   Call CREACAMPO("STOCKVIE", "MOVISTO", "Cod_Exte", 10, 16)
   Call CREACAMPO("STOCKVIE", "MOVISTO", "Descrip", 10, 64)
   Call CREACAMPO("STOCKVIE", "MOVISTO", "IdenLote", 10, 16)
   Call CREACAMPO("STOCKVIE", "MOVISTO", "CodiMovi", 10, 2)
   Call CREACAMPO("STOCKVIE", "MOVISTO", "DoPriCor", 10, 12)
   Call CREACAMPO("STOCKVIE", "MOVISTO", "DoPriLar", 10, 24)
   Call CREACAMPO("STOCKVIE", "MOVISTO", "DoSecCor", 10, 12)
   Call CREACAMPO("STOCKVIE", "MOVISTO", "DoSecLar", 10, 24)
   Call CREACAMPO("STOCKVIE", "MOVISTO", "NuorItem", 3, 2)
   Call CREACAMPO("STOCKVIE", "MOVISTO", "NuPedCli", 4, 4)
   Call CREACAMPO("STOCKVIE", "MOVISTO", "ReferCor", 10, 64)
   Call CREACAMPO("STOCKVIE", "MOVISTO", "ValoUnit", 7, 0)
   Call CREACAMPO("STOCKVIE", "MOVISTO", "MoneVal", 3, 2)
   Call CREACAMPO("STOCKVIE", "MOVISTO", "DepoMovi", 3, 0)
   Call CREACAMPO("STOCKVIE", "MOVISTO", "Nume_Clie", 3, 0)
   Call CREACAMPO("STOCKVIE", "MOVISTO", "Nume_Prov", 3, 0)
   Call CREACAMPO("STOCKVIE", "MOVISTO", "CantIngre", 7, 0)
   Call CREACAMPO("STOCKVIE", "MOVISTO", "CantSalid", 7, 0)
   Call CREACAMPO("STOCKVIE", "MOVISTO", "CantSaldo", 7, 0)
   Call CREACAMPO("STOCKVIE", "MOVISTO", "FeReMovi", 8, 0)
   Call CREACAMPO("STOCKVIE", "MOVISTO", "NumUsuar", 3, 0)
   '
End Sub
'
Sub CREATABLA_VENTAS()
   '
   Call CREACAMPO("VENTAS", "DESCOMBI", "Nume_Cli", 4, 0)
   Call CREACAMPO("VENTAS", "DESCOMBI", "GrCo_Cli", 3, 2)
   Call CREACAMPO("VENTAS", "DESCOMBI", "Cod_Inte", 3, 2)
   Call CREACAMPO("VENTAS", "DESCOMBI", "GruCoVen", 3, 0)
   Call CREACAMPO("VENTAS", "DESCOMBI", "PorDescu", 10, 12)
   Call CREACAMPO("VENTAS", "DESCOMBI", "Modo_Desc", 3, 0)
   '
   Call CREACAMPO("VENTAS", "DETCOTIZA", "NroCoti", 10, 16, 1)
   Call CREACAMPO("VENTAS", "DETCOTIZA", "ClaValor", 10, 48, 1)
   Call CREACAMPO("VENTAS", "DETCOTIZA", "Valor1", 12, 0)
   Call CREACAMPO("VENTAS", "DETCOTIZA", "Valor2", 12, 0)
   Call CREACAMPO("VENTAS", "DETCOTIZA", "Valor3", 12, 0)
   '
   Call CREACAMPO("VENTAS", "LISTAPRE", "Cod_Lista", 3, 2, 1)
   Call CREACAMPO("VENTAS", "LISTAPRE", "Denom_Lista", 10, 64)
   Call CREACAMPO("VENTAS", "LISTAPRE", "Vigencia", 8, 2)
   Call CREACAMPO("VENTAS", "LISTAPRE", "Formato", 3, 2)
   Call CREACAMPO("VENTAS", "LISTAPRE", "Decimales", 3, 2)
   Call CREACAMPO("VENTAS", "LISTAPRE", "Moneda", 3, 2)
   Call CREACAMPO("VENTAS", "LISTAPRE", "Desc_Abrev", 10, 64)
   Call CREACAMPO("VENTAS", "LISTAPRE", "Descu_lista", 6, 32)
   Call CREACAMPO("VENTAS", "LISTAPRE", "IncluIva", 3, 0)
   '
   Call CREACAMPO("VENTAS", "PEDDETA", "PreList", 6, 4)
   Call CREACAMPO("VENTAS", "PEDDETA", "CanAsig", 6, 4)
   Call CREACAMPO("VENTAS", "PEDDETA", "CanPed", 6, 4, 1)
   Call CREACAMPO("VENTAS", "PEDDETA", "PreUnid", 6, 4)
   Call CREACAMPO("VENTAS", "PEDDETA", "ImpoTot", 7, 8)
   Call CREACAMPO("VENTAS", "PEDDETA", "ImpoAnti", 7, 8)
   Call CREACAMPO("VENTAS", "PEDDETA", "CantEnt", 6, 4, 1)
   Call CREACAMPO("VENTAS", "PEDDETA", "CanSaldo", 6, 4)
   Call CREACAMPO("VENTAS", "PEDDETA", "FeUltEnt", 8, 0)
   Call CREACAMPO("VENTAS", "PEDDETA", "FApro_Nope", 8, 0)
   Call CREACAMPO("VENTAS", "PEDDETA", "AcuPrepa", 6, 0)
   Call CREACAMPO("VENTAS", "PEDDETA", "AcuDespa", 6, 0)
   Call CREACAMPO("VENTAS", "PEDDETA", "NroRemi", 10, 24)
   Call CREACAMPO("VENTAS", "PEDDETA", "TPedItem", 10, 16)
   Call CREACAMPO("VENTAS", "PEDDETA", "NroPed", 4, 0, 1)
   Call CREACAMPO("VENTAS", "PEDDETA", "TipoPed", 10, 8, 1)
   Call CREACAMPO("VENTAS", "PEDDETA", "FechEmis", 8, 0)
   Call CREACAMPO("VENTAS", "PEDDETA", "NroClie", 4, 0, 1)
   Call CREACAMPO("VENTAS", "PEDDETA", "RaSoClie", 10, 64)
   Call CREACAMPO("VENTAS", "PEDDETA", "NroOcom", 10, 16)
   Call CREACAMPO("VENTAS", "PEDDETA", "FeSolEnt", 8, 0)
   Call CREACAMPO("VENTAS", "PEDDETA", "FinVigen", 8, 0)
   Call CREACAMPO("VENTAS", "PEDDETA", "Status", 3, 0, 1)
   Call CREACAMPO("VENTAS", "PEDDETA", "StaDespa", 3, 0)
   Call CREACAMPO("VENTAS", "PEDDETA", "ListPre", 3, 0)
   Call CREACAMPO("VENTAS", "PEDDETA", "CondPag", 3, 0)
   Call CREACAMPO("VENTAS", "PEDDETA", "Vended", 3, 0)
   Call CREACAMPO("VENTAS", "PEDDETA", "PorComVend", 6, 0)
   Call CREACAMPO("VENTAS", "PEDDETA", "NuOrItem", 3, 0, 1)
   Call CREACAMPO("VENTAS", "PEDDETA", "CodiInt", 3, 0, 1)
   Call CREACAMPO("VENTAS", "PEDDETA", "CodiExt", 10, 16)
   Call CREACAMPO("VENTAS", "PEDDETA", "Descrip", 10, 64)
   Call CREACAMPO("VENTAS", "PEDDETA", "SubCod", 10, 8)
   Call CREACAMPO("VENTAS", "PEDDETA", "TColor", 10, 16)
   Call CREACAMPO("VENTAS", "PEDDETA", "TTalle", 10, 16)
   Call CREACAMPO("VENTAS", "PEDDETA", "TLargo", 6, 4)
   Call CREACAMPO("VENTAS", "PEDDETA", "TAncho", 6, 4)
   Call CREACAMPO("VENTAS", "PEDDETA", "TEspesor", 6, 4)
   Call CREACAMPO("VENTAS", "PEDDETA", "CBase", 10, 16)
   Call CREACAMPO("VENTAS", "PEDDETA", "VenRep", 10, 8)
   Call CREACAMPO("VENTAS", "PEDDETA", "MoneEmis", 3, 2)
   Call CREACAMPO("VENTAS", "PEDDETA", "TipoCam", 6, 4)
   Call CREACAMPO("VENTAS", "PEDDETA", "DescLista", 6, 0)
   Call CREACAMPO("VENTAS", "PEDDETA", "DescCombi", 10, 64)
   Call CREACAMPO("VENTAS", "PEDDETA", "DescRubro", 6, 0)
   Call CREACAMPO("VENTAS", "PEDDETA", "DescTotal", 6, 0)
   Call CREACAMPO("VENTAS", "PEDDETA", "NumeroSerie", 10, 32)
   Call CREACAMPO("VENTAS", "PEDDETA", "IdenLote", 10, 32)
   Call CREACAMPO("VENTAS", "PEDDETA", "LISTA_DE_ITEM", 3, 0)
   '
   Call CREACAMPO("VENTAS", "PEDENCA", "NroPed", 4, 0, 1)
   Call CREACAMPO("VENTAS", "PEDENCA", "Referencia", 10, 48)
   Call CREACAMPO("VENTAS", "PEDENCA", "TipoPed", 10, 8, 1)
   Call CREACAMPO("VENTAS", "PEDENCA", "FechEmis", 8, 0)
   Call CREACAMPO("VENTAS", "PEDENCA", "NroClie", 4, 0, 1)
   Call CREACAMPO("VENTAS", "PEDENCA", "RasoClie", 10, 64)
   Call CREACAMPO("VENTAS", "PEDENCA", "NroOcom", 10, 16)
   Call CREACAMPO("VENTAS", "PEDENCA", "DomiEnt", 10, 48)
   Call CREACAMPO("VENTAS", "PEDENCA", "LocaEnt", 10, 48)
   Call CREACAMPO("VENTAS", "PEDENCA", "TranspEnt", 10, 48)
   Call CREACAMPO("VENTAS", "PEDENCA", "NTranspo", 3, 0)
   Call CREACAMPO("VENTAS", "PEDENCA", "NSucEnt", 3, 0)
   Call CREACAMPO("VENTAS", "PEDENCA", "Observa", 12, 0)
   Call CREACAMPO("VENTAS", "PEDENCA", "FeSolEnt", 8, 0)
   Call CREACAMPO("VENTAS", "PEDENCA", "FinVigen", 8, 0)
   Call CREACAMPO("VENTAS", "PEDENCA", "Status", 3, 0, 1)
   Call CREACAMPO("VENTAS", "PEDENCA", "VenRep", 10, 8, 1)
   Call CREACAMPO("VENTAS", "PEDENCA", "ListPre", 3, 0)
   Call CREACAMPO("VENTAS", "PEDENCA", "CondPag", 3, 0)
   Call CREACAMPO("VENTAS", "PEDENCA", "Vended", 3, 0)
   Call CREACAMPO("VENTAS", "PEDENCA", "PorComVend", 6, 0)
   Call CREACAMPO("VENTAS", "PEDENCA", "TDescApli", 10, 14)
   Call CREACAMPO("VENTAS", "PEDENCA", "AnotaPed", 12, 0)
   Call CREACAMPO("VENTAS", "PEDENCA", "FApro_Nope", 8, 0)
   Call CREACAMPO("VENTAS", "PEDENCA", "Usua_Apro ", 3, 2)
   Call CREACAMPO("VENTAS", "PEDENCA", "Rec_ConPa", 10, 14)
   Call CREACAMPO("VENTAS", "PEDENCA", "OV_Motivo", 3, 0)
   Call CREACAMPO("VENTAS", "PEDENCA", "OV_Destino", 10, 128)
   Call CREACAMPO("VENTAS", "PEDENCA", "OV_Vehiculo", 10, 128)
   Call CREACAMPO("VENTAS", "PEDENCA", "OV_Fesalida", 8, 0)
   Call CREACAMPO("VENTAS", "PEDENCA", "OV_FeRegreso", 8, 0)
   Call CREACAMPO("VENTAS", "PEDENCA", "OV_CondDespa", 12, 0)
   Call CREACAMPO("VENTAS", "PEDENCA", "OV_Autoriza1", 10, 64)
   Call CREACAMPO("VENTAS", "PEDENCA", "OV_Autoriza2", 10, 64)
   Call CREACAMPO("VENTAS", "PEDENCA", "OV_Observaciones", 12, 0)
   Call CREACAMPO("VENTAS", "PEDENCA", "OV_NSerieDev", 10, 24)
   Call CREACAMPO("VENTAS", "PEDENCA", "CodCarpImp", 10, 12)
   Call CREACAMPO("VENTAS", "PEDENCA", "TipoIngre", 3, 0)
   Call CREACAMPO("VENTAS", "PEDENCA", "Dto_Ret_Cte", 10, 14)
   Call CREACAMPO("VENTAS", "PEDENCA", "DtoxCant", 10, 14)
   '
   Call CREACAMPO("VENTAS", "PRESUVEN", "NroPres", 6, 4)
   Call CREACAMPO("VENTAS", "PRESUVEN", "Cod_Rubro", 10, 8)
   Call CREACAMPO("VENTAS", "PRESUVEN", "NumeCRubro", 4, 4)
   Call CREACAMPO("VENTAS", "PRESUVEN", "Cod_Anio", 10, 4)
   Call CREACAMPO("VENTAS", "PRESUVEN", "FechEmis", 8, 0)
   Call CREACAMPO("VENTAS", "PRESUVEN", "NroClie", 4, 0)
   Call CREACAMPO("VENTAS", "PRESUVEN", "Codi_Cli", 10, 12)
   Call CREACAMPO("VENTAS", "PRESUVEN", "RasoClie", 10, 64)
   Call CREACAMPO("VENTAS", "PRESUVEN", "Referen", 10, 16)
   Call CREACAMPO("VENTAS", "PRESUVEN", "AteSen", 10, 48)
   Call CREACAMPO("VENTAS", "PRESUVEN", "LisPre", 3, 0)
   Call CREACAMPO("VENTAS", "PRESUVEN", "ConPag", 3, 0)
   Call CREACAMPO("VENTAS", "PRESUVEN", "Vended", 3, 0)
   Call CREACAMPO("VENTAS", "PRESUVEN", "MonedEmi", 3, 2)
   Call CREACAMPO("VENTAS", "PRESUVEN", "NuOrdIt", 3, 0)
   Call CREACAMPO("VENTAS", "PRESUVEN", "CodiInt", 3, 0)
   Call CREACAMPO("VENTAS", "PRESUVEN", "CodiExt", 10, 16)
   Call CREACAMPO("VENTAS", "PRESUVEN", "Descrip", 10, 64)
   Call CREACAMPO("VENTAS", "PRESUVEN", "Especif", 12, 0)
   Call CREACAMPO("VENTAS", "PRESUVEN", "PreList", 7, 0)
   Call CREACAMPO("VENTAS", "PRESUVEN", "TDescuen", 10, 32)
   Call CREACAMPO("VENTAS", "PRESUVEN", "PorcDesc", 6, 4)
   Call CREACAMPO("VENTAS", "PRESUVEN", "ImpoDesc", 6, 4)
   Call CREACAMPO("VENTAS", "PRESUVEN", "PreUnit", 6, 4)
   Call CREACAMPO("VENTAS", "PRESUVEN", "Cantit", 6, 4)
   Call CREACAMPO("VENTAS", "PRESUVEN", "ImToNe", 7, 8)
   Call CREACAMPO("VENTAS", "PRESUVEN", "FechVent", 8, 0)
   Call CREACAMPO("VENTAS", "PRESUVEN", "NroPedi", 6, 0)
   Call CREACAMPO("VENTAS", "PRESUVEN", "NroFac_Cor", 10, 12)
   Call CREACAMPO("VENTAS", "PRESUVEN", "NroFac_Lar", 10, 24)
   Call CREACAMPO("VENTAS", "PRESUVEN", "MarNoVig", 10, 1)
   Call CREACAMPO("VENTAS", "PRESUVEN", "CBase", 10, 12)
   Call CREACAMPO("VENTAS", "PRESUVEN", "TTalle", 10, 12)
   Call CREACAMPO("VENTAS", "PRESUVEN", "TColor", 10, 16)
   Call CREACAMPO("VENTAS", "PRESUVEN", "NroCoti", 10, 16)
   Call CREACAMPO("VENTAS", "PRESUVEN", "Reventa", 10, 48)
   Call CREACAMPO("VENTAS", "PRESUVEN", "Antecedentes", 12, 0)
   Call CREACAMPO("VENTAS", "PRESUVEN", "Observa", 12, 0)
   Call CREACAMPO("VENTAS", "PRESUVEN", "Imagen", 11, 0)
   Call CREACAMPO("VENTAS", "PRESUVEN", "Status", 3, 0)
   Call CREACAMPO("VENTAS", "PRESUVEN", "Grco_Cli", 3, 2)
   Call CREACAMPO("VENTAS", "PRESUVEN", "CantBultos", 12, 0)
   Call CREACAMPO("VENTAS", "PRESUVEN", "PesoNeto", 10, 128)
   Call CREACAMPO("VENTAS", "PRESUVEN", "PesoTotal", 10, 128)
   Call CREACAMPO("VENTAS", "PRESUVEN", "NCM", 12, 0)
   Call CREACAMPO("VENTAS", "PRESUVEN", "CondVta", 10, 128)
   Call CREACAMPO("VENTAS", "PRESUVEN", "Vers_Presu", 3, 0, 1)
   Call CREACAMPO("VENTAS", "PRESUVEN", "Vers_Presu", 3, 0, 1)
   Call CREACAMPO("VENTAS", "PRESUVEN", "Num_Prov", 4, 0, 1)
   Call CREACAMPO("VENTAS", "PRESUVEN", "COSTO_UNIT", 6, 4)
   Call CREACAMPO("VENTAS", "PRESUVEN", "Via_Entrega", 3, 0, 1)
   Call CREACAMPO("VENTAS", "PRESUVEN", "MARBORRA", 3, 0, 1)
   Call CREACAMPO("VENTAS", "PRESUVEN", "COTIZAMONEDA", 7, 0)
   Call CREACAMPO("VENTAS", "PRESUVEN", "LOTES_Reservados", 10, 512)
   Call CREACAMPO("VENTAS", "PRESUVEN", "Canti_Lotes", 6, 0)
   Call CREACAMPO("VENTAS", "PRESUVEN", "TTIVA21", 7, 0)
   Call CREACAMPO("VENTAS", "PRESUVEN", "TTIVA105", 7, 0)
   Call CREACAMPO("VENTAS", "PRESUVEN", "TOTPER", 7, 0)
   
   '
   Call CREACAMPO("VENTAS", "PREVENTA", "Cod_Lista", 3, 2, 1)
   Call CREACAMPO("VENTAS", "PREVENTA", "Cod_Pieza", 3, 2, 1)
   Call CREACAMPO("VENTAS", "PREVENTA", "Num_Orden", 3, 2)
   Call CREACAMPO("VENTAS", "PREVENTA", "Desc_Pieza", 10, 64)
   Call CREACAMPO("VENTAS", "PREVENTA", "Denom_Pieza", 10, 64)
   Call CREACAMPO("VENTAS", "PREVENTA", "Carac_Control", 10, 64)
   Call CREACAMPO("VENTAS", "PREVENTA", "Num_Original", 10, 64)
   Call CREACAMPO("VENTAS", "PREVENTA", "Precio_Unitario", 7, 8)
   Call CREACAMPO("VENTAS", "PREVENTA", "Cant_Limite", 6, 8)
   Call CREACAMPO("VENTAS", "PREVENTA", "Porc_Limite", 7, 32)
   Call CREACAMPO("VENTAS", "PREVENTA", "MarBorra", 3, 0)
   '
   Call CREACAMPO("VENTAS", "TROQUELES", "CodTroquel", 10, 12)
   Call CREACAMPO("VENTAS", "TROQUELES", "DescriTroquel", 10, 64)
   Call CREACAMPO("VENTAS", "TROQUELES", "DenoTroquel", 10, 64)
   Call CREACAMPO("VENTAS", "TROQUELES", "FormTroquel", 10, 32)
   Call CREACAMPO("VENTAS", "TROQUELES", "AvanceTroquel", 6, 0)
   Call CREACAMPO("VENTAS", "TROQUELES", "AnchoTroquel", 6, 0)
   Call CREACAMPO("VENTAS", "TROQUELES", "DiamTroquel", 6, 0)
   Call CREACAMPO("VENTAS", "TROQUELES", "BocasTroquel", 3, 0)
   Call CREACAMPO("VENTAS", "TROQUELES", "MediTroquel", 6, 0)
   '
End Sub
Sub CREARCAMPOS_FACTURAYGASTOS()
       Call ABRECONEXION
       Call CREACAMPO("", "FACTURAYGASTOS", "Nume_Cli", 6, 0)
       Call CREACAMPO("", "FACTURAYGASTOS", "VAL_COMPRA", 7, 0)
       Call CREACAMPO("", "FACTURAYGASTOS", "FECHEMISI", 8, 0)
       Call CREACAMPO("", "FACTURAYGASTOS", "IMPORTE_FACTURA", 7, 0)
       Call CREACAMPO("", "FACTURAYGASTOS", "IMPO_GASTO", 7, 0)
       Call CREACAMPO("", "FACTURAYGASTOS", "FACTURAPROV_ASOC", 10, 24)
       Call CREACAMPO("", "FACTURAYGASTOS", "TIPO_FACTURA", 10, 8, 1)
       Call CREACAMPO("", "FACTURAYGASTOS", "FACTURAPROV", 10, 24)
       Call CREACAMPO("", "FACTURAYGASTOS", "NUMEBORE", 6, 4, 1)
       Call CREACAMPO("", "FACTURAYGASTOS", "MoneVal", 3, 2, 1)
       Call CREACAMPO("", "FACTURAYGASTOS", "TipoCam", 7, 8)
End Sub
Sub GENERA_TABLA_COMPRAS_DE_PRESUPUESTO()
       Call ABRECONEXION
       Call CREACAMPO("", "PRESUCOMPRA", "Nume_Cli", 4, 0, 1)
       Call CREACAMPO("", "PRESUCOMPRA", "Cantidad", 7, 0)
       Call CREACAMPO("", "PRESUCOMPRA", "FECHEMISI", 8, 0)
       Call CREACAMPO("", "PRESUCOMPRA", "NroPres", 6, 0, 1)
       Call CREACAMPO("", "PRESUCOMPRA", "Num_Prov", 6, 0)
       Call CREACAMPO("", "PRESUCOMPRA", "Codint", 3, 0)
       Call CREACAMPO("", "PRESUCOMPRA", "MARBORRA", 3, 0, 1)
       Call CREACAMPO("", "PRESUCOMPRA", "STATUS", 3, 0)
       Call CREACAMPO("", "PRESUCOMPRA", "NumSolCompra", 6, 0)
       Call CREACAMPO("", "PRESUCOMPRA", "COSTO", 7, 0)
       Call CREACAMPO("", "PRESUCOMPRA", "ORDEN", 3, 0, 1)
       Call CREACAMPO("", "PRESUCOMPRA", "Mone_Cos", 3, 0)
       
End Sub

