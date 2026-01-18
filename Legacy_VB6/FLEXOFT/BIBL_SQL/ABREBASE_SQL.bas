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
Public articulos As Database
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

Public compras As Database
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
Sub CreaTablaMovistoCTP()

   Call CREACAMPO("STOCKS", "MOVISTOCTP", "IdenMovi", 10, 32)
   If InStr(EPAC$, "CONTEPLAST") > 0 Or InStr(EPAC$, "CAMBACERES") > 0 Then
    If ValidaExisteCampo%("MOVISTOCTP_ID", "MOVISTOCTP") < 1 Then
      SQLX$ = "ALTER TABLE MOVISTOCTP ADD MOVISTOCTP_ID INT IDENTITY"
      FLEXCONN.Execute (SQLX$)
    End If
   End If
   
   Call CREACAMPO("STOCKS", "MOVISTOCTP", "FechMov", 8, 0, 1)
   Call CREACAMPO("STOCKS", "MOVISTOCTP", "CodiEmpr", 3, 2, 1)
   Call CREACAMPO("STOCKS", "MOVISTOCTP", "FeReMovi", 8, 0)
   Call CREACAMPO("STOCKS", "MOVISTOCTP", "Descri_Lar", 12, 0)
   Call CREACAMPO("STOCKS", "MOVISTOCTP", "IdenLote", 10, 16, 1)
   Call CREACAMPO("STOCKS", "MOVISTOCTP", "CodiMovi", 10, 2, 1)
   Call CREACAMPO("STOCKS", "MOVISTOCTP", "DoPriCor", 10, 12, 1)
   Call CREACAMPO("STOCKS", "MOVISTOCTP", "DoPriLar", 10, 24, 1)
   Call CREACAMPO("STOCKS", "MOVISTOCTP", "DoSecCor", 10, 12, 1)
   Call CREACAMPO("STOCKS", "MOVISTOCTP", "DoSecLar", 10, 24, 1)
   Call CREACAMPO("STOCKS", "MOVISTOCTP", "NuorItem", 3, 2)
   Call CREACAMPO("STOCKS", "MOVISTOCTP", "NuPedCli", 4, 4, 1)
   Call CREACAMPO("STOCKS", "MOVISTOCTP", "NumeOrdPed", 3, 2)
   Call CREACAMPO("STOCKS", "MOVISTOCTP", "ValoUnit", 7, 0)
   Call CREACAMPO("STOCKS", "MOVISTOCTP", "ValTotIt", 7, 0)
   Call CREACAMPO("STOCKS", "MOVISTOCTP", "MoneVal", 3, 2)
   Call CREACAMPO("STOCKS", "MOVISTOCTP", "TipoCam", 7, 8)
   Call CREACAMPO("STOCKS", "MOVISTOCTP", "DepoMovi", 3, 0, 1)
   Call CREACAMPO("STOCKS", "MOVISTOCTP", "Nume_Clie", 4, 0, 1)
   Call CREACAMPO("STOCKS", "MOVISTOCTP", "RASO_CLI", 10, 64)
   Call CREACAMPO("STOCKS", "MOVISTOCTP", "CantSalid", 7, 0)
   Call CREACAMPO("STOCKS", "MOVISTOCTP", "CantSaldo", 7, 0)
   Call CREACAMPO("STOCKS", "MOVISTOCTP", "NUMPVTA", 3, 0, 1)
   Call CREACAMPO("STOCKS", "MOVISTOCTP", "DOPRINUM", 4, 0, 1)
   Call CREACAMPO("STOCKS", "MOVISTOCTP", "MOVISTO_IDCTP", 6, 0)
   Call CREACAMPO("STOCKS", "MOVISTOCTP", "CantIngre", 7, 0)
   Call CREACAMPO("STOCKS", "MOVISTOCTP", "CondPag", 3, 0)
   Call CREACAMPO("STOCKS", "MOVISTOCTP", "Vended", 3, 0)
   Call CREACAMPO("STOCKS", "MOVISTOCTP", "NumUsuar", 3, 0)
   Call CREACAMPO("STOCKS", "MOVISTOCTP", "Cod_Inte", 3, 0, 1)
   Call CREACAMPO("STOCKS", "MOVISTOCTP", "Cod_Exte", 10, 24)
   Call CREACAMPO("STOCKS", "MOVISTOCTP", "Descrip", 10, 64)
   Call CREACAMPO("STOCKS", "MOVISTOCTP", "FeRemAWb", 8, 0) 'fecha remitido a waldbot
   Call CREACAMPO("STOCKS", "MOVISTOCTP", "PasoaWb", 3, 0, 1) 'marca de pasado a Waldbot
   Call CREACAMPO("STOCKS", "MOVISTOCTP", "MARBORRA", 3, 0)
   Call CREACAMPO("STOCKS", "MOVISTOCTP", "VerificadoPor", 3, 0)
   Call CREACAMPO("STOCKS", "MOVISTOCTP", "DomiEnt", 10, 64)
   Call CREACAMPO("STOCKS", "MOVISTOCTP", "NAMEVEND", 10, 64)
   Call CREACAMPO("STOCKS", "MOVISTOCTP", "ORDENCOMPRA", 10, 32)
   Call CREACAMPO("STOCKS", "MOVISTOCTP", "NroVend_Envap", 3, 0)
   Call CREACAMPO("STOCKS", "MOVISTOCTP", "ObservaItem", 10, 512)
   Call CREACAMPO("STOCKS", "MOVISTOCTP", "Descridos", 10, 128)
   Call CREACAMPO("STOCKS", "MOVISTOCTP", "NroModificado", 10, 32)
   Call CREACAMPO("STOCKS", "MOVISTOCTP", "AcuEntParcial", 7, 0) 'acumula las entregas parciales
   
   

End Sub

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
     Case "VENTAS"
        Call ABREVENTAS
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
   CONNSQL_DESTINO$ = REPLACAR((CONNSQL$), TRIM$(UCase$(BINIC$)), TRIM$(UCase$(BASE_DESTI$)))
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
   strSQL = "Select * from " + NTABLA$ + " WHERE " + CONDI$
   Dim RTABLA As ADODB.Recordset
   Set RTABLA = New ADODB.Recordset
25 On Error Resume Next
   RTABLA.Open (strSQL), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText

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
         !CODIEMPR = CodiEmp%
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
Sub ABREVENTAS()
   '
   If ESTABIER%("VENTAS") < 1 Then
      Call MENSERR(24, "Imposible Abrir Base de Pedidos de Clientes.")
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

Sub ABRECALIDAD()
   '
   If ESTABIER%("CALIDAD") < 1 Then
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
   strSQL = "If NOT EXISTS (Select name from sysobjects where name= '" & NTABLA$ & "')"
   strSQL = strSQL & " CREATE TABLE " & NTABLA$ & "("
   strSQL = strSQL & "[CodiEmpr] [smallint] NULL ,"
   strSQL = strSQL & "[CodiSucu] [smallint] NULL ,"
   strSQL = strSQL & "[NuClien] [smallint] NULL ,"
   strSQL = strSQL & "[CodiClie] [nvarchar] (16) COLLATE Modern_Spanish_CI_AS NULL ,"
   strSQL = strSQL & "[RasoClie] [nvarchar] (48) COLLATE Modern_Spanish_CI_AS NULL ,"
   strSQL = strSQL & "[FechEmisi] [smalldatetime] NULL ,"
   strSQL = strSQL & "[CoDocAp_Lar] [nvarchar] (18) COLLATE Modern_Spanish_CI_AS NULL ,"
   strSQL = strSQL & "[CoDocAp_Cor] [nvarchar] (12) COLLATE Modern_Spanish_CI_AS NULL ,"
   strSQL = strSQL & "[ImpoApli] [float] NULL ,"
   strSQL = strSQL & "[RefeApli] [nvarchar] (32) COLLATE Modern_Spanish_CI_AS NULL ,"
   strSQL = strSQL & "[CoDocOr_Lar] [nvarchar] (18) COLLATE Modern_Spanish_CI_AS NULL ,"
   strSQL = strSQL & "[CoDocOr_Cor] [nvarchar] (12) COLLATE Modern_Spanish_CI_AS NULL ,"
   strSQL = strSQL & "[ItotOrig] [float] NULL ,"
   strSQL = strSQL & "[IIvaOrig] [float] NULL ,"
   strSQL = strSQL & "[Mone_Apli] [smallint] NULL ,"
   strSQL = strSQL & "[ItotOri_Dol] [float] NULL ,"
   strSQL = strSQL & "[ImpoApli_Dol] [float] NULL"
   strSQL = strSQL & ") ON [PRIMARY]"
   
   FLEXCONN.Execute (strSQL)
   
   'EJEMPLO DE AGREGAR COLUMNAS
   strSQL = "IF EXISTS (Select name from sysobjects where name= '" & NTABLA$ & "')"
   strSQL = strSQL & " ALTER TABLE " & NTABLA$ & " ADD column_A INT NULL"
   strSQL = strSQL & " , column_B VARCHAR NULL"
   FLEXCONN.Execute (strSQL)
   
   'EJEMPLO DE BORRAR COLUMNAS
   strSQL = "IF EXISTS (Select name from sysobjects where name= '" & NTABLA$ & "')"
   strSQL = strSQL & " ALTER TABLE " & NTABLA$ & " DROP COLUMN column_b"
   
   'EJEMPLO DE MODIFICAR COLUMNAS
   strSQL = "IF EXISTS (Select name from sysobjects where name= '" & NTABLA$ & "')"
   strSQL = strSQL & " ALTER TABLE " & NTABLA$ & " ALTER COLUMN column_A smalldatetime NULL DEFAULT getdate() WITH VALUES"
   '
   'EJEMPLO DE CREAR UNA BASE DE DATOS - pls JANDY VERIFICAR
   strSQL = "IF EXISTS (Select name from sysdatabases where name ='" + BASEACTIVA$ + "')"
   strSQL = strSQL & " create database '" + BASEACTIVA$ + "'"
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
    strSQL = "select right('0'+cast(datepart(dd,getdate()) as varchar(2)),2)+'/'"
    strSQL = strSQL & "+ right('0'+cast(datepart(mm,getdate()) as varchar(2)),2)+'/'"
    strSQL = strSQL & "+ right('0'+cast(datepart(yy,getdate()) as varchar(4)),2) as MiFecha"
    '
    Set MIRS = FLEXCONN.Execute(strSQL)
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
    strSQL = "select right('0'+cast(datepart(hh,getdate()) as varchar(2)),2)+':'"
    strSQL = strSQL & "+ right('0'+cast(datepart(mi,getdate()) as varchar(2)),2)+':'"
    strSQL = strSQL & "+ right('0'+cast(datepart(ss,getdate()) as varchar(4)),2) as MiHora"
    '
    Set MIRS = FLEXCONN.Execute(strSQL)
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

Function EXISTE_CAMPO%(NTABLA$, NCAMPO$)
    'DEVUELVE 0 SI NO EXISE O 1 SI EXISTE
    EXISTE_CAMPO% = 0
    SQLX$ = "SELECT COUNT(*) From information_schema.Columns  "
    SQLX$ = SQLX$ + " WHERE  table_name= '" & NTABLA$ & "'"
    SQLX$ = SQLX$ + " and column_name='" & NCAMPO$ & "'"
    Set rs = READSET(SQLX$, , 1)
    EXISTE_CAMPO% = rs.Fields(0).Value
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
      strSQL = "SELECT SYSOBJECTS.NAME AS NOTAB,SYSOBJECTS.ID AS IDENTA, SYSCOLUMNS.NAME AS NOCAM, "
      strSQL = strSQL + " SYSCOLUMNS.XTYPE AS TYPE, SYSCOLUMNS.LENGTH AS LENG "
      strSQL = strSQL + " FROM SYSCOLUMNS WITH(NOLOCK) "
      strSQL = strSQL + " INNER JOIN sysobjects "
      strSQL = strSQL & " ON sysobjects.id = syscolumns.id "
      strSQL = strSQL + " WHERE Left(SYSOBJECTS.NAME, 3) <> 'sys' "
      strSQL = strSQL + " AND LEFT(SYSOBJECTS.NAME,3) <> 'DF_' "
      '
      Set MIRS = FLEXCONN.Execute(strSQL)
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
      strSQL = "Select * from sysindexes "
      ' STRSQL = "Select * from sysindexes WHERE id = OBJECT_ID('" & NTABLA$ & "') AND NAME = 'I" + NCAMPO$ + "' "
      Set MIRS = FLEXCONN.Execute(strSQL)
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
      strSQL = "SELECT SYSOBJECTS.NAME AS NOTAB,SYSOBJECTS.ID AS IDENTA, SYSCOLUMNS.NAME AS NOCAM, "
      strSQL = strSQL + " SYSCOLUMNS.XTYPE AS TYPE, SYSCOLUMNS.LENGTH AS LENG "
      strSQL = strSQL + " FROM SYSCOLUMNS WITH(NOLOCK) "
      strSQL = strSQL + " INNER JOIN sysobjects "
      strSQL = strSQL & " ON sysobjects.id = syscolumns.id "
      strSQL = strSQL + " WHERE Left(SYSOBJECTS.NAME, 3) <> 'sys' "
      strSQL = strSQL + " AND LEFT(SYSOBJECTS.NAME,3) <> 'DF_' "
      '
      Set MIRS = FLEXCONN_DESTINO.Execute(strSQL)
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
      strSQL = "Select * from sysindexes "
      ' STRSQL = "Select * from sysindexes WHERE id = OBJECT_ID('" & NTABLA$ & "') AND NAME = 'I" + NCAMPO$ + "' "
      Set MIRS = FLEXCONN_DESTINO.Execute(strSQL)
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
    strSQL = "Select name from MASTER.DBO.sysdatabases where name = '" + BASEACTIVA_DESTI$ + "'"
    On Error Resume Next
    Set MIRS = FLEXCONN_DESTINO.Execute(strSQL)
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
        TXCAMPO$ = TXCAMPO$ + "[binary] (" + TRIM$(str$(LOCAMPO)) + ")": TSQL = 173
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
        TXCAMPO$ = TXCAMPO$ + "[nvarchar] (" + TRIM$(str$(LOCAMPO)) + ") COLLATE Modern_Spanish_CI_AS": TSQL = 231
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
           strSQL = "CREATE TABLE " & NTABLA$ & " (" + TXCAMPO$ + ") ON [PRIMARY]"
           On Error Resume Next
           FLEXCONN_DESTINO.Execute (strSQL)      ' CREO LA TABLA Y EL PRIMER CAMPO
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
    strSQL = "SELECT SYSCOLUMNS.NAME AS NAME, SYSCOLUMNS.XTYPE AS XTYPE1, SYSCOLUMNS.LENGTH AS LENG FROM syscolumns "
    strSQL = strSQL + "INNER JOIN sysobjects "
    strSQL = strSQL & "ON sysobjects.id = syscolumns.id "
    strSQL = strSQL & "where sysobjects.name = '" & NTABLA$ & "' "
    strSQL = strSQL & "AND syscolumns.name = '" & UCase$(NCAMPO$) & "'"
    ' STRSQL = STRSQL & "AND syscolumns.name LIKE '" & UCase$(NCAMPO$) & "'"
    ' EL 1/4/2009 SE CAMBIO LIKE POR = YA QUE NO CREABA EL CAMPO DESC_ITEM POR
    ' CONFUNDIRLO CON EL DESCRITEM. QUEDA POR AVERIGUAR POR QUÉ SE PUSO ENTES EL
    ' LIKE EN VEZ DEL IGUAL. - epb
    Set MIRS = FLEXCONN_DESTINO.Execute(strSQL)
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
           strSQL = "Select * from sysindexes WHERE id = OBJECT_ID('" & NTABLA$ & "') AND NAME = 'I" + NCAMPO$ + "' "
           Set MIRS1 = FLEXCONN_DESTINO.Execute(strSQL)
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
       strSQL = "Select * from sysindexes WHERE id = OBJECT_ID('" & NTABLA$ & "') AND NAME = 'I" + NCAMPO$ + "' "
       Set MIRS = FLEXCONN_DESTINO.Execute(strSQL)
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
    strSQL = "Select name from MASTER.DBO.sysdatabases where name = '" + BASEACTIVA$ + "'"
    On Error Resume Next
    Set MIRS = FLEXCONN.Execute(strSQL)
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
        TXCAMPO$ = TXCAMPO$ + "[binary] (" + TRIM$(str$(LOCAMPO)) + ")": TSQL = 173
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
        TXCAMPO$ = TXCAMPO$ + "[nvarchar] (" + TRIM$(str$(LOCAMPO)) + ") COLLATE Modern_Spanish_CI_AS": TSQL = 231
      Case 12
        TXCAMPO$ = TXCAMPO$ + "[ntext] ": TSQL = 99 ' ojo como trabajarlo
      Case 13
        TXCAMPO$ = TXCAMPO$ + "[bigint]": TSQL = 127
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
           strSQL = "CREATE TABLE " & NTABLA$ & " (" + TXCAMPO$ + ") ON [PRIMARY]"
           On Error Resume Next
           FLEXCONN.Execute (strSQL)      ' CREO LA TABLA Y EL PRIMER CAMPO
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
    strSQL = "SELECT SYSCOLUMNS.NAME AS NAME, SYSCOLUMNS.XTYPE AS XTYPE1, SYSCOLUMNS.LENGTH AS LENG FROM syscolumns "
    strSQL = strSQL + "INNER JOIN sysobjects "
    strSQL = strSQL & "ON sysobjects.id = syscolumns.id "
    strSQL = strSQL & "where sysobjects.name = '" & NTABLA$ & "' "
    strSQL = strSQL & "AND syscolumns.name = '" & UCase$(NCAMPO$) & "'"
    ' STRSQL = STRSQL & "AND syscolumns.name LIKE '" & UCase$(NCAMPO$) & "'"
    ' EL 1/4/2009 SE CAMBIO LIKE POR = YA QUE NO CREABA EL CAMPO DESC_ITEM POR
    ' CONFUNDIRLO CON EL DESCRITEM. QUEDA POR AVERIGUAR POR QUÉ SE PUSO ENTES EL
    ' LIKE EN VEZ DEL IGUAL. - epb
    Set MIRS = FLEXCONN.Execute(strSQL)
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
           strSQL = "Select * from sysindexes WHERE id = OBJECT_ID('" & NTABLA$ & "') AND NAME = 'I" + NCAMPO$ + "' "
           Set MIRS1 = FLEXCONN.Execute(strSQL)
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
    ' *** CONINDICE% = 1 DEFINE INDICE NO AGRUPADO. CONINDICE% = 2 DEFINE INDICE AGRUPADO
80  If CONINDICE% > 0 Then
       ' verifica si el indice existe
       strSQL = "Select * from sysindexes WHERE id = OBJECT_ID('" & NTABLA$ & "') AND NAME = 'I" + NCAMPO$ + "' "
       Set MIRS = FLEXCONN.Execute(strSQL)
       If Not (MIRS.EOF And MIRS.BOF) Then
          ' *** VERIFICA SI EL INDICE CAMBIO DE NO AGRUPADO A AGRUPADO. EN CASO AFIRMATIVO,
          ' *** ELIMINA EL INDICE Y LO VUELVA A CREAR PERO AGRUPADO.
          ' *** status = 0 No agrupado | status = 16 Agrupado
          If MIRS!Status = 0 And CONINDICE% = 2 Then
            strSQL = "DROP INDEX " & NTABLA$ & ".I" + NCAMPO$
            Set MIRS = FLEXCONN.Execute(strSQL)
          Else
            GoTo 90                              ' el indice ya existe
          End If
       End If
       '
83     TIMUT = FLEXCONN.CommandTimeout
       FLEXCONN.CommandTimeout = 999
       CRINDEX$ = "CREATE INDEX I" + NCAMPO$ + " ON " + NTABLA$ + "([" + NCAMPO$ + "]);" ' INDICE NO AGRUPADO (NON CLUSTERED)
       If CONINDICE% = 2 Then ' INDICE AGRUPADO (CLUSTERED)
            CRINDEX$ = "CREATE CLUSTERED INDEX I" + NCAMPO$ + " ON " + NTABLA$ + "([" + NCAMPO$ + "]);"
       End If
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

Sub CreaIndiceCompuesto(TABLA As String, Campos As String, Optional TipoIndice As Integer)

    ' TipoIndice = 1 INDICE NO AGRUPADO (NON CLUSTERED)
    ' TipoIndice = 2 INDICE AGRUPADO (CLUSTERED)

    If TRIM$(TABLA) = "" Then Exit Sub
    If TRIM$(Campos$) = "" Then Exit Sub
    If TipoIndice < 1 Or TipoIndice > 2 Then TipoIndice = 1
    
    ' *** ARMA NOMBRE DEL INDICE Y LISTA DE CAMPOS
    CamposIndice$ = Campos
    PosInicial% = 1
    Do While CamposIndice$ <> ""
        PosFinal% = InStr(1, CamposIndice$, "+")
        If PosFinal% = 0 Then PosFinal% = Len(CamposIndice$) + 1
        NombreIndice$ = NombreIndice$ & "_" & TRIM$(REPLACAR$(REPLACAR$(Mid$(CamposIndice$, PosInicial%, PosFinal% - 1), " ASC", ""), " DESC", ""))
        ListaCamposIndice$ = ListaCamposIndice$ & "," & Mid$(CamposIndice$, PosInicial%, PosFinal% - 1)
        CamposIndice$ = TRIM$(Mid$(CamposIndice$, PosFinal% + 1))
    Loop
    
    NombreIndice$ = Mid$(NombreIndice$, 2)
    ListaCamposIndice$ = Mid$(ListaCamposIndice$, 2)

    ' *** VERIFICA SI EXISTE EL INDICE
    ConsultaSql$ = "Select * from sysindexes WHERE id = OBJECT_ID('" & TABLA & "') AND NAME = 'I" + NombreIndice$ + "' "
    Set RstIndices = FLEXCONN.Execute(ConsultaSql$)
    If Not (RstIndices.EOF And RstIndices.BOF) Then
      ' *** VERIFICA SI EL INDICE CAMBIO DE NO AGRUPADO A AGRUPADO. EN CASO AFIRMATIVO,
      ' *** ELIMINA EL INDICE Y LO VUELVA A CREAR PERO AGRUPADO.
      ' *** status = 0 No agrupado | status = 16 Agrupado
      If RstIndices!Status = 0 And TipoIndice = 2 Then
        ConsultaSql$ = "DROP INDEX " & TABLA & ".I" + NombreIndice$
        FLEXCONN.Execute (ConsultaSql$)
      Else
        Exit Sub ' EL INDICE EXISTE
      End If
   End If
   On Error Resume Next
   RstIndices.Close
   Set RstIndices = Nothing
   On Error GoTo 0
   
   '*** CREA EL NUEVO INDICE
   TIMUT = FLEXCONN.CommandTimeout
   FLEXCONN.CommandTimeout = 999
   CRINDEX$ = "CREATE INDEX I" + NombreIndice$ + " ON " + TABLA + "(" + ListaCamposIndice$ + ");" ' INDICE NO AGRUPADO (NON CLUSTERED)
   If TipoIndice = 2 Then ' INDICE AGRUPADO (CLUSTERED)
        CRINDEX$ = "CREATE CLUSTERED INDEX I" + NombreIndice$ + " ON " + TABLA + "([" + ListaCamposIndice$ + "]);"
   End If
   On Error Resume Next   ' PARA ATAJAR QUE EL INDICE YA EXISTA
   FLEXCONN.Execute CRINDEX$
   If Err = False Then
        Call COMUNI("INDICE I" + NombreIndice$ + " CREADO EN TABLA " + TABLA)
   Else
        Call COMUNI("Imposible Crear Indice I" + NombreIndice$ + " EN TABLA " + TABLA + ". Consulte para Creación Manual del Indice Requerido.")
   End If
   On Error GoTo 0
   FLEXCONN.CommandTimeout = TIMUT

End Sub

Sub COPIA_STRUC_MOVISTO(Optional TABLA$)
    '
    If TABLA$ <> "" Then
      If EXISTE_TABLA(TABLA$) > 0 Then Exit Sub
    End If
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
       If TABLA$ <> "" Then
          Call CREACAMPO("", TABLA$, NOCAM$, TICAM, LOCAM, CONIN%)
       Else 'por defecto movivie
          Call CREACAMPO("", "MOVIVIE", NOCAM$, TICAM, LOCAM, CONIN%)
       End If
       '
    Next KKI%
    '
End Sub

Sub COPIA_STRUC_MOVISTO_A_MOVIPRU()
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
       Call CREACAMPO("", "MOVIPRU", NOCAM$, TICAM, LOCAM, CONIN%)
       '
    Next KKI%
    Call CREACAMPO("", "MOVIPRU", "YAMIGRADO", 3, 0, 1)
    
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
   If TRIM$(UCase$(NOMBASE$)) = "CALIDAD" Then Call CREATABLA_CALIDAD: Exit Sub
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
   If ValidaExisteCampo%("formulas_id", "Formulas") < 1 Then
      SQLX$ = "ALTER TABLE Formulas ADD formulas_id INT IDENTITY"
      FLEXCONN.Execute (SQLX$)
    End If

   Call CREACAMPO("ARTICULOS", "Formulas", "CodIConj", 3, 0, 1)
   Call CREACAMPO("ARTICULOS", "Formulas", "CodXConj", 10, 24)
   Call CREACAMPO("ARTICULOS", "Formulas", "CodiElem", 3, 0, 1)
   Call CREACAMPO("ARTICULOS", "Formulas", "CodXElem", 10, 24)
   Call CREACAMPO("ARTICULOS", "Formulas", "NuorIt", 3, 2, 1)
   Call CREACAMPO("ARTICULOS", "Formulas", "UsoBruto", 6, 0)
   Call CREACAMPO("ARTICULOS", "Formulas", "Unis_Form", 6, 0)
   Call CREACAMPO("ARTICULOS", "Formulas", "Lar_Corte", 6, 0)
   Call CREACAMPO("ARTICULOS", "Formulas", "Pladispo", 3, 0)
   Call CREACAMPO("ARTICULOS", "Formulas", "RefeComp", 12, 0)
   Call CREACAMPO("ARTICULOS", "Formulas", "Marborra", 3, 0)
   Call CREACAMPO("ARTICULOS", "Formulas", "COMPONENTE", 10, 48) 'CONTIENE LA TAREA CONTEPLAST (EJ. TAPA O ESLINGAS)
   Call CREACAMPO("ARTICULOS", "Formulas", "MODUXBB", 3, 0, 1)
   Call CREACAMPO("ARTICULOS", "Formulas", "CORTE", 7, 0)
   Call CREACAMPO("ARTICULOS", "FORMULAS", "UMED_FORM", 10, 8)
   Call CREACAMPO("ARTICULOS", "FORMULAS", "VALOBADA_1", 10, 512)

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
   'los combios realizados en master igualar en Master2
   '***************************************************
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
   Call CREACAMPO("ARTICULOS", "Master", "Codint", 4, 0, 2)
   Call CREACAMPO("ARTICULOS", "Master", "Alias", 10, 24, 1)
   Call CREACAMPO("ARTICULOS", "Master", "Umedida", 10, 4)
   Call CREACAMPO("ARTICULOS", "Master", "Timate", 3, 0, 1)
   Call CREACAMPO("ARTICULOS", "Master", "Critrep", 3, 0)
   Call CREACAMPO("ARTICULOS", "Master", "Monecos", 3, 0)
   Call CREACAMPO("ARTICULOS", "Master", "Fecosto", 8, 0)
   Call CREACAMPO("ARTICULOS", "Master", "PorDescCos", 6, 0)
   Call CREACAMPO("ARTICULOS", "Master", "Cosuni", 6, 0)
   Call CREACAMPO("ARTICULOS", "Master", "Pornacio", 6, 0)
   Call CREACAMPO("ARTICULOS", "Master", "PorComVen", 6, 0)
   Call CREACAMPO("ARTICULOS", "Master", "Provhabi", 3, 0)
   Call CREACAMPO("ARTICULOS", "Master", "Stomin", 6, 0)
   Call CREACAMPO("ARTICULOS", "Master", "Stomax", 6, 0)
   Call CREACAMPO("ARTICULOS", "Master", "Lorepos", 6, 0)
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
   Call CREACAMPO("ARTICULOS", "Master", "Trazable", 3, 0, 1)
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
   Call CREACAMPO("ARTICULOS", "Master", "CueContaI", 3, 0, 1)
   Call CREACAMPO("ARTICULOS", "Master", "Molde", 3, 0, 1)
   Call CREACAMPO("ARTICULOS", "Master", "Consumo_X_Corte", 3, 0)
   Call CREACAMPO("ARTICULOS", "Master", "GRAMAJE", 6, 0)
   Call CREACAMPO("ARTICULOS", "Master", "MICRONES", 6, 0)
   Call CREACAMPO("ARTICULOS", "Master", "ANCHO", 6, 0)
   
   Call CREACAMPO("ARTICULOS", "Master", "GR_ACTU", 3, 0)
   Call CREACAMPO("ARTICULOS", "Master", "Cod_Arba", 10, 16)
   Call CREACAMPO("ARTICULOS", "Master", "VALOBADA_1", 10, 512)
   Call CREACAMPO("ARTICULOS", "Master", "VALOBADA_2", 10, 512)
   Call CREACAMPO("ARTICULOS", "Master", "VALOBADA_3", 10, 512)
   Call CREACAMPO("ARTICULOS", "Master", "VALOBADA_4", 10, 512)
   Call CREACAMPO("ARTICULOS", "Master", "VALOBADA_5", 10, 512)
   Call CREACAMPO("ARTICULOS", "Master", "VALOBADA_6", 10, 512)

   
   
   'los combios realizados en master igualar en Master2
   '***************************************************
'   EPAC$ = TRIM$(UCase$(EMPREAC$))
'   If InStr(EPAC$, "AHP") > 0 Then
        Call CREACAMPO("ARTICULOS", "Master2", "Ferevis", 8, 0)
        Call CREACAMPO("ARTICULOS", "Master2", "Observa", 12, 0)
        Call CREACAMPO("ARTICULOS", "Master2", "Especificacion", 12, 0)
        Call CREACAMPO("ARTICULOS", "Master2", "DesCompAut", 3, 0)
        Call CREACAMPO("ARTICULOS", "Master2", "DescriComp", 10, 64)
        Call CREACAMPO("ARTICULOS", "Master2", "DepoEntra", 3, 0)
        Call CREACAMPO("ARTICULOS", "Master2", "DepoCon", 3, 0)
        Call CREACAMPO("ARTICULOS", "Master2", "Codigo", 10, 24, 1)
        Call CREACAMPO("ARTICULOS", "Master2", "Descripcion", 10, 64)
        Call CREACAMPO("ARTICULOS", "Master2", "Descridos", 10, 128)
        Call CREACAMPO("ARTICULOS", "Master2", "Status", 3, 0)
        Call CREACAMPO("ARTICULOS", "Master2", "Codint", 4, 0, 1)
        Call CREACAMPO("ARTICULOS", "Master2", "Alias", 10, 24, 1)
        Call CREACAMPO("ARTICULOS", "Master2", "Umedida", 10, 4)
        Call CREACAMPO("ARTICULOS", "Master2", "Timate", 3, 0, 1)
        Call CREACAMPO("ARTICULOS", "Master2", "Critrep", 3, 0)
        Call CREACAMPO("ARTICULOS", "Master2", "Monecos", 3, 0)
        Call CREACAMPO("ARTICULOS", "Master2", "Fecosto", 8, 0)
        Call CREACAMPO("ARTICULOS", "Master2", "PorDescCos", 6, 0)
        Call CREACAMPO("ARTICULOS", "Master2", "Cosuni", 6, 0)
        Call CREACAMPO("ARTICULOS", "Master2", "Pornacio", 6, 0)
        Call CREACAMPO("ARTICULOS", "Master2", "PorComVen", 6, 0)
        Call CREACAMPO("ARTICULOS", "Master2", "Provhabi", 3, 0)
        Call CREACAMPO("ARTICULOS", "Master2", "Stomin", 6, 0)
        Call CREACAMPO("ARTICULOS", "Master2", "Stomax", 6, 0)
        Call CREACAMPO("ARTICULOS", "Master2", "Lorepos", 6, 0)
        Call CREACAMPO("ARTICULOS", "Master2", "Lindivis", 6, 0)
        Call CREACAMPO("ARTICULOS", "Master2", "Pesuni", 6, 0)
        Call CREACAMPO("ARTICULOS", "Master2", "PesMetro", 6, 0)
        Call CREACAMPO("ARTICULOS", "Master2", "Color", 10, 32)
        Call CREACAMPO("ARTICULOS", "Master2", "Talle", 10, 32)
        Call CREACAMPO("ARTICULOS", "Master2", "Descompo", 3, 0)
        Call CREACAMPO("ARTICULOS", "Master2", "Grupiva", 3, 0)
        Call CREACAMPO("ARTICULOS", "Master2", "Grucoven", 3, 0)
        Call CREACAMPO("ARTICULOS", "Master2", "CLinea", 10, 8)
        Call CREACAMPO("ARTICULOS", "Master2", "Cfamilia", 10, 8)
        Call CREACAMPO("ARTICULOS", "Master2", "NPlano", 10, 24)
        Call CREACAMPO("ARTICULOS", "Master2", "Revision", 10, 16)
        Call CREACAMPO("ARTICULOS", "Master2", "Trazable", 3, 0, 1)
        Call CREACAMPO("ARTICULOS", "Master2", "Revi_Formu", 8, 0)
        Call CREACAMPO("ARTICULOS", "Master2", "FechApFor", 8, 0)
        Call CREACAMPO("ARTICULOS", "Master2", "UsuApFor", 10, 24)
        Call CREACAMPO("ARTICULOS", "Master2", "Cod_Interna", 10, 64)
        Call CREACAMPO("ARTICULOS", "Master2", "Bien_Cap", 3, 0)
        Call CREACAMPO("ARTICULOS", "Master2", "Codigo_Ncm", 10, 32)
        Call CREACAMPO("ARTICULOS", "Master2", "Codigo_Sec", 10, 32)
        Call CREACAMPO("ARTICULOS", "Master2", "Pro_Umed", 10, 32)
        Call CREACAMPO("ARTICULOS", "Master2", "Iva_Id", 10, 8)
        Call CREACAMPO("ARTICULOS", "Master2", "ArchiPdf", 12, 0)
        Call CREACAMPO("ARTICULOS", "Master2", "Familia", 10, 4)
        Call CREACAMPO("ARTICULOS", "Master2", "Lindiven", 6, 0)
        Call CREACAMPO("ARTICULOS", "Master2", "CanDeci", 10, 2)
        Call CREACAMPO("ARTICULOS", "Master2", "ConsuProm", 6, 0)
        Call CREACAMPO("ARTICULOS", "Master2", "CosCal", 7, 0)
        Call CREACAMPO("ARTICULOS", "Master2", "FeCosCal", 8, 0)
        Call CREACAMPO("ARTICULOS", "Master2", "Stat_Movi", 3, 0)
        Call CREACAMPO("ARTICULOS", "Master2", "Clas_ATC", 10, 64)
        Call CREACAMPO("ARTICULOS", "Master2", "Vers_BC", 10, 32)
        Call CREACAMPO("ARTICULOS", "Master2", "FeVig_BC", 8, 0)
        Call CREACAMPO("ARTICULOS", "Master2", "FeRev_BC", 8, 0)
        Call CREACAMPO("ARTICULOS", "Master2", "CantApro", 3, 0)
        Call CREACAMPO("ARTICULOS", "Master2", "LeTiCom", 10, 16)
        Call CREACAMPO("ARTICULOS", "Master2", "PiezasXCaja", 4, 0)
        Call CREACAMPO("ARTICULOS", "Master2", "M3Caja", 6, 0)
        Call CREACAMPO("ARTICULOS", "Master2", "KiloNetoCaja", 6, 0)
        Call CREACAMPO("ARTICULOS", "Master2", "KiloBrutoCaja", 6, 0)
        Call CREACAMPO("ARTICULOS", "Master2", "Autono_Proy", 10, 10)
        Call CREACAMPO("ARTICULOS", "Master2", "UsuPrePes", 10, 24)
        Call CREACAMPO("ARTICULOS", "Master2", "UsuPreDol", 10, 24)
        Call CREACAMPO("ARTICULOS", "Master2", "MANO_OBRA_REAL", 6, 0)
        Call CREACAMPO("ARTICULOS", "Master2", "Tamano_Etiqueta", 3, 0)
        Call CREACAMPO("ARTICULOS", "Master2", "CueContaI", 3, 0, 1)
        Call CREACAMPO("ARTICULOS", "Master2", "Molde", 3, 0, 1)
        Call CREACAMPO("ARTICULOS", "Master2", "GR_ACTU", 3, 0)

'   End If
        'TABLA DE JYM CON CAMBIOS DE PRECIO MINORISTAS
        Call CREACAMPO("", "COSMINORISTA", "FECHINICIO", 8, 0)
        Call CREACAMPO("", "COSMINORISTA", "NumeroUsuario", 3, 0)
        Call CREACAMPO("", "COSMINORISTA", "Porcentaje", 7, 0)
        Call CREACAMPO("", "COSMINORISTA", "UltUpdate", 8, 0)
        Call CREACAMPO("", "COSMINORISTA", "Cod_Inte", 3, 0, 1)



   Call CREACAMPO("", "CODICLIENTE", "Codigo_Cliente", 10, 24, 1)
   Call CREACAMPO("", "CODICLIENTE", "Status", 3, 0)
   Call CREACAMPO("", "CODICLIENTE", "Codint", 4, 0, 1)

   '
   Call CREACAMPO("ARTICULOS", "SECOPER", "CodIConj", 3, 0, 1)
   Call CREACAMPO("ARTICULOS", "SECOPER", "CodXConj", 10, 24)
   Call CREACAMPO("ARTICULOS", "SECOPER", "NumeOper", 3, 0, 1)
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
   Call CREACAMPO("ARTICULOS", "SECOPER", "SecuenciaFabricacion", 3, 0)
   Call CREACAMPO("ARTICULOS", "SECOPER", "OperacionTerceros", 3, 0)
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
   'TABLA DE UBICACIONES POR PROVEEDOR (AHP)
   Call CREACAMPO("", "UBICASQL", "Cod_Inte", 4, 0, 1)
   Call CREACAMPO("", "UBICASQL", "CODIUBICA", 10, 32)
   Call CREACAMPO("", "UBICASQL", "CANTIDAD", 6, 4)
   Call CREACAMPO("", "UBICASQL", "OBSERVACION", 10, 64)
   Call CREACAMPO("", "UBICASQL", "Nume_Pro", 4, 0, 1)
   Call CREACAMPO("", "UBICASQL", "MARBORRA", 3, 0)
   '
   Call CREACAMPO("", "COMBINA_IVA", "DescripcionEq", 10, 64)
   If ValidaExisteCampo%("COMBINA_IVA_ID", "COMBINA_IVA") < 1 Then
    SQLX$ = "ALTER TABLE COMBINA_IVA ADD COMBINA_IVA_ID INT IDENTITY"
    FLEXCONN.Execute (SQLX$)
   End If
   '
   Call CREACAMPO("", "COMBINA_IVA", "CodigoEq", 10, 24, 1)
   Call CREACAMPO("", "COMBINA_IVA", "CodintEq", 4, 0, 1)
   Call CREACAMPO("", "COMBINA_IVA", "CodigoAcc", 10, 24, 1)
   Call CREACAMPO("", "COMBINA_IVA", "DescripcionAcc", 10, 64)
   Call CREACAMPO("", "COMBINA_IVA", "CodintAcc", 4, 0, 1)
   Call CREACAMPO("", "COMBINA_IVA", "Status", 3, 0)

   Call GENERATABLAIMPRE

   Call CREACAMPO("", "HISTOCOTIZA", "Moneda", 3, 2)
   Call CREACAMPO("", "HISTOCOTIZA", "FechaCotizacion", 8, 0)
   Call CREACAMPO("", "HISTOCOTIZA", "Cotizacion", 7, 0)
   
   Call CreaIndiceCompuesto("HISTOCOTIZA", "Moneda+FechaCotizacion", 1)
   
   Call CREACAMPO("", "INDICEINDEC", "Periodo", 4, 0, 1)
   Call CREACAMPO("", "INDICEINDEC", "IndiceInflacion", 7, 8)
   Call CREATABLA_ENCAPRF
   
   'REPLICA EL DAT SECPROD CADA VEZ QUE SE MODIFICA
   Call CREACAMPO("", "secprod_sql", "NumeOper", 10, 12)
   Call CREACAMPO("", "secprod_sql", "DescOper", 10, 64)
   Call CREACAMPO("", "secprod_sql", "CODIEMPR", 3, 2, 1)
   If ValidaExisteCampo%("secprod_sql_ID", "secprod_sql") < 1 Then
    SQLX$ = "ALTER TABLE secprod_sql ADD secprod_sql_ID INT IDENTITY"
    FLEXCONN.Execute (SQLX$)
   End If
   Call CREACAMPO("", "secprod_sql", "MARBORRA", 3, 0)


End Sub
   
'
Sub CREATABLA_BADACLIEN()
   '
   Call CREACAMPO("BADACLIEN", "BSEGCOM", "Codi_Cli", 10, 12, 1)
   Call CREACAMPO("BADACLIEN", "BSEGCOM", "Nume_Cli", 4, 0, 1)
   Call CREACAMPO("BADACLIEN", "BSEGCOM", "Raso_Cli", 10, 128)
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
   Call CREACAMPO("BADACLIEN", "Deudor12", "Lentre3_Cli", 10, 10)
   Call CREACAMPO("BADACLIEN", "Deudor12", "ProvEntre_Cli", 10, 2)
   Call CREACAMPO("BADACLIEN", "Deudor12", "Hora_Ent", 10, 48)
   Call CREACAMPO("BADACLIEN", "Deudor12", "Zona_Cli", 10, 32)
   Call CREACAMPO("BADACLIEN", "Deudor12", "Zona_Rep", 10, 8)
   Call CREACAMPO("BADACLIEN", "Deudor12", "Gru_Div", 4, 0)
   Call CREACAMPO("BADACLIEN", "Deudor12", "Nume_Cli", 4, 0, 2)
   Call CREACAMPO("BADACLIEN", "Deudor12", "Codi_Cli", 10, 12, 1)
   Call CREACAMPO("BADACLIEN", "Deudor12", "Raso_Cli", 10, 128, 1)
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
   Call CREACAMPO("BADACLIEN", "Deudor12", "RegimenFacturaCredito", 3, 0)
   Call CREACAMPO("BADACLIEN", "Deudor12", "Nibr_Cli", 10, 24)
   Call CREACAMPO("BADACLIEN", "Deudor12", "Prvd_Cli", 10, 24)
   Call CREACAMPO("BADACLIEN", "Deudor12", "Ctac_Cli", 10, 48)
   Call CREACAMPO("BADACLIEN", "Deudor12", "Hora_Cli", 10, 32)
   Call CREACAMPO("BADACLIEN", "Deudor12", "Cpag_Cli", 3, 0)
   Call CREACAMPO("BADACLIEN", "Deudor12", "Lpre_Cli", 3, 0)
   Call CREACAMPO("BADACLIEN", "Deudor12", "Vend_Cli", 3, 0)
   Call CREACAMPO("BADACLIEN", "Deudor12", "Tcpr_Cli", 3, 0, 1)
   Call CREACAMPO("BADACLIEN", "Deudor12", "CaIb_Cli", 10, 2)
   Call CREACAMPO("BADACLIEN", "Deudor12", "JurisdiccioIIBB_CLI", 10, 2)
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
   Call CREACAMPO("BADACLIEN", "Deudor12", "Mail_Cob", 10, 128)
   Call CREACAMPO("BADACLIEN", "Deudor12", "DiasRec_Cob", 10, 64)
   Call CREACAMPO("BADACLIEN", "Deudor12", "DiasPago_Cob", 10, 64)
   Call CREACAMPO("BADACLIEN", "Deudor12", "Segui_Cob", 12, 0)
   Call CREACAMPO("BADACLIEN", "Deudor12", "MailFactura_Cob", 10, 64)
   Call CREACAMPO("BADACLIEN", "Deudor12", "CELU_CLI", 10, 128)
   Call CREACAMPO("BADACLIEN", "Deudor12", "FechAlta", 8, 0)
   '
   Call CREACAMPO("BADACLIEN", "SUCENTRE", "Nume_Cli", 4, 0, 1)
   Call CREACAMPO("BADACLIEN", "SUCENTRE", "Nume_Suc", 3, 0)
   Call CREACAMPO("BADACLIEN", "SUCENTRE", "Deno_Suc", 10, 128)
   Call CREACAMPO("BADACLIEN", "SUCENTRE", "Domi_Suc", 10, 128)
   Call CREACAMPO("BADACLIEN", "SUCENTRE", "CodPostal_Suc", 10, 10)
   Call CREACAMPO("BADACLIEN", "SUCENTRE", "Loca_Suc", 10, 128)
   Call CREACAMPO("BADACLIEN", "SUCENTRE", "Tele_Suc", 10, 64)
   Call CREACAMPO("BADACLIEN", "SUCENTRE", "Codi_Tra", 3, 0)
   Call CREACAMPO("BADACLIEN", "SUCENTRE", "Prov_Cli", 10, 2)
   Call CREACAMPO("BADACLIEN", "SUCENTRE", "Hora_Ent", 10, 48)
   Call CREACAMPO("BADACLIEN", "SUCENTRE", "InstruccionSuc", 10, 512)
   Call CREACAMPO("BADACLIEN", "SUCENTRE", "MARBORRA", 3, 0)
   Call CREACAMPO("BADACLIEN", "SUCENTRE", "STATUS", 3, 0)
   
   
   '
   Call CREACAMPO("BADACLIEN", "TRANSPORTES", "Codi_Tra", 3, 0, 1)
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
   Call CREACAMPO("BADACLIEN", "TRANSPORTES", "Prov_Tra", 10, 4)
   Call CREACAMPO("BADACLIEN", "TRANSPORTES", "CodPostal_Tra", 10, 24)
   Call CREACAMPO("BADACLIEN", "TRANSPORTES", "HoraTransporte", 10, 48)
   
   Call CREACAMPO("BADACLIEN", "INSTRUC_ENTRE", "Instruccion", 10, 512)
   If ValidaExisteCampo%("INSTRUC_ENTRE_ID", "INSTRUC_ENTRE") < 1 Then
      SQLX$ = "ALTER TABLE INSTRUC_ENTRE ADD INSTRUC_ENTRE_ID INT IDENTITY"
      FLEXCONN.Execute (SQLX$)
   End If
   Call CREACAMPO("BADACLIEN", "INSTRUC_ENTRE", "Nume_Cli", 4, 0, 1)
   Call CREACAMPO("BADACLIEN", "INSTRUC_ENTRE", "Nume_Suc", 3, 0, 1)
   Call CREACAMPO("BADACLIEN", "INSTRUC_ENTRE", "Codi_Tra", 3, 0, 1)
   Call CREACAMPO("BADACLIEN", "INSTRUC_ENTRE", "MARBORRA", 3, 0)
   '
   Call CREACAMPO("BADACLIEN", "VENDEDOR", "Vend_Cli", 3, 0, 1)
   Call CREACAMPO("BADACLIEN", "VENDEDOR", "RasoVend", 10, 48)
   Call CREACAMPO("BADACLIEN", "VENDEDOR", "DomiVend", 10, 48)
   Call CREACAMPO("BADACLIEN", "VENDEDOR", "LocaVend", 10, 48)
   Call CREACAMPO("BADACLIEN", "VENDEDOR", "TeleVend_1", 10, 32)
   Call CREACAMPO("BADACLIEN", "VENDEDOR", "TeleVend_2", 10, 32)
   Call CREACAMPO("BADACLIEN", "VENDEDOR", "EmailVend", 10, 48)
   Call CREACAMPO("BADACLIEN", "VENDEDOR", "Observaciones", 12, 0)
   '
   
   Call CREACAMPO("BADACLIEN", "AgendaOriginal", "CodTrabajo", 3, 0, 1)
   Call CREACAMPO("BADACLIEN", "AgendaOriginal", "CodexTRab", 10, 24)
   Call CREACAMPO("BADACLIEN", "AgendaOriginal", "DescTrab", 10, 64)

   Call CREACAMPO("BADACLIEN", "AgendaOriginal", "TipoTrabajo", 10, 64, 1)
   Call CREACAMPO("BADACLIEN", "AgendaOriginal", "Dia", 10, 24)
   Call CREACAMPO("BADACLIEN", "AgendaOriginal", "OrdenSemana", 10, 16)
   Call CREACAMPO("BADACLIEN", "AgendaOriginal", "Horario", 10, 32)
   Call CREACAMPO("BADACLIEN", "AgendaOriginal", "Continuidad", 3, 2)
   Call CREACAMPO("BADACLIEN", "AgendaOriginal", "Cuotas", 3, 48)
   Call CREACAMPO("BADACLIEN", "AgendaOriginal", "Nume_Cli", 4, 0, 1)
   Call CREACAMPO("BADACLIEN", "AgendaOriginal", "MARBORRA", 3, 0)
   Call CREACAMPO("BADACLIEN", "AgendaOriginal", "ORDEN", 3, 0)
   Call CREACAMPO("BADACLIEN", "AgendaOriginal", "ID_DIA", 3, 0, 1)
   Call CREACAMPO("BADACLIEN", "AgendaOriginal", "APARTIRDE", 10, 16)
   If ValidaExisteCampo%("ID", "AgendaOriginal") < 1 Then
    SQLX$ = "ALTER TABLE AgendaOriginal ADD ID INT IDENTITY"
    FLEXCONN.Execute (SQLX$)
    Call CREACAMPO("BADACLIEN", "AgendaOriginal", "ID", 4, 0, 1)
   End If
   '
   Call CREACAMPO("BADACLIEN", "Agenda", "CodTrabajo", 3, 0, 1)
   Call CREACAMPO("BADACLIEN", "Agenda", "TipoTrabajo", 10, 64)
   Call CREACAMPO("BADACLIEN", "Agenda", "CodexTRab", 10, 24)
   Call CREACAMPO("BADACLIEN", "Agenda", "DescTrab", 10, 64)
   Call CREACAMPO("BADACLIEN", "Agenda", "Fecha_Original", 8, 0)
   Call CREACAMPO("BADACLIEN", "Agenda", "OrdenSemana", 10, 16)
   Call CREACAMPO("BADACLIEN", "Agenda", "Horario_Orig", 10, 32)
   Call CREACAMPO("BADACLIEN", "Agenda", "Continuidad", 3, 2)
   Call CREACAMPO("BADACLIEN", "Agenda", "Cuotas", 3, 48)
   Call CREACAMPO("BADACLIEN", "Agenda", "Nume_Cli", 4, 0, 1)
   Call CREACAMPO("BADACLIEN", "Agenda", "MARBORRA", 3, 0)
   Call CREACAMPO("BADACLIEN", "Agenda", "ORDEN", 3, 0)
   Call CREACAMPO("BADACLIEN", "Agenda", "Dia_Organizado", 8, 0)
   Call CREACAMPO("BADACLIEN", "Agenda", "Horario_Organizado", 10, 32)
   Call CREACAMPO("BADACLIEN", "Agenda", "Dia_Realizado", 8, 0)
   Call CREACAMPO("BADACLIEN", "Agenda", "FACTURADO", 10, 8)
   Call CREACAMPO("BADACLIEN", "Agenda", "MESANO_ORIG", 10, 32)
   Call CREACAMPO("BADACLIEN", "Agenda", "ID_ORIGINAL", 4, 0, 1)
   Call CREACAMPO("BADACLIEN", "Agenda", "REALIZADO", 3, 0)
   Call CREACAMPO("BADACLIEN", "Agenda", "OPER_REALIZO", 3, 0)
   Call CREACAMPO("BADACLIEN", "Agenda", "PEDIDO", 4, 0, 1)
   If ValidaExisteCampo%("ID", "Agenda") < 1 Then
    SQLX$ = "ALTER TABLE Agenda ADD ID INT IDENTITY"
    FLEXCONN.Execute (SQLX$)
    Call CREACAMPO("BADACLIEN", "Agenda", "ID", 4, 0, 1)
   End If
       
       
'   If EXISTEVISTA%("[SUCUTRANSP]") < 1 Then
'       SQLX$ = "CREATE VIEW SUCUTRANSP AS"
'    Else
'       SQLX$ = "ALTER VIEW SUCUTRANSP AS"
'    End If
'
'       SQLX$ = SQLX$ + " select s.nume_cli,t.Deno_Tra,T.Codi_Tra,S.Nume_Suc,S.Deno_Suc  "
'       SQLX$ = SQLX$ + " from sucentre s "
'       SQLX$ = SQLX$ + " inner join transportes t on s.Codi_Tra = t.Codi_Tra "
'       FLEXCONN.Execute (SQLX$)

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
   Call CREACAMPO("BADAPROVE", "Proved12", "Nume_Cli", 4, 0, 2)
   Call CREACAMPO("BADAPROVE", "Proved12", "Codi_Cli", 10, 12, 1)
   Call CREACAMPO("BADAPROVE", "Proved12", "Raso_Cli", 10, 128)
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
   Call CREACAMPO("BADAPROVE", "Proved12", "CategoriaIBCABA_Cli", 3, 0)
   Call CREACAMPO("BADAPROVE", "Proved12", "CategoriaIBTUCU_Cli", 3, 0)
   Call CREACAMPO("BADAPROVE", "Proved12", "CategoriaIBMISIONES_Cli", 3, 0)
   Call CREACAMPO("BADAPROVE", "Proved12", "ObsCompras", 12, 0)
   Call CREACAMPO("BADAPROVE", "Proved12", "LuEntCompras", 12, 0)
   '
   Call CREATABLA_COSTOXPROVE
End Sub
Sub CREATABLA_COSTOXPROVE()
    Call CREACAMPO("", "COSTOXPROVE", "Nume_Cli", 4, 0, 1)
    Call CREACAMPO("", "COSTOXPROVE", "CODINT", 4, 0, 1)
    Call CREACAMPO("", "COSTOXPROVE", "LISTA", 3, 0)
    Call CREACAMPO("", "COSTOXPROVE", "PRECIOVENTA", 7, 8)
    Call CREACAMPO("", "COSTOXPROVE", "Via_Entrega", 3, 0)
    Call CREACAMPO("", "COSTOXPROVE", "MARBORRA", 3, 0)
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
   Call CREACAMPO("BCONTROL", "SQL_QUERIES", "Tip_Query", 3, 2)
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
   Call CREACAMPO("COMPRAS", "AUTOENT", "Nume_Ocom", 4, 4, 1)
   Call CREACAMPO("COMPRAS", "AUTOENT", "Nume_Aten", 4, 4)
   Call CREACAMPO("COMPRAS", "AUTOENT", "Femi_Aten", 8, 0)
   Call CREACAMPO("COMPRAS", "AUTOENT", "Npro_Aten", 3, 2)
   Call CREACAMPO("COMPRAS", "AUTOENT", "Raso_Aten", 10, 64)
   Call CREACAMPO("COMPRAS", "AUTOENT", "Cod_Inte", 3, 0, 1)
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
   Call CREACAMPO("COMPRAS", "COTIRECI", "Nume_Solint", 4, 4, 1)
   Call CREACAMPO("COMPRAS", "COTIRECI", "NuOrd_item", 3, 2)
   Call CREACAMPO("COMPRAS", "COTIRECI", "Npro_Ocom", 4, 4)
   Call CREACAMPO("COMPRAS", "COTIRECI", "Cod_Inte", 3, 2, 1)
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
   Call CREACAMPO("COMPRAS", "COTIRECI", "Descuento", 10, 64)
   '
   Call CREACAMPO("COMPRAS", "DETAIVA", "Nume_Ocom", 4, 4, 1)
   Call CREACAMPO("COMPRAS", "DETAIVA", "Norit", 3, 2)
   Call CREACAMPO("COMPRAS", "DETAIVA", "Tasa_Iva", 10, 16)
   Call CREACAMPO("COMPRAS", "DETAIVA", "Base_Impo", 7, 8)
   Call CREACAMPO("COMPRAS", "DETAIVA", "Impo_Iva", 7, 8)
   '
   Call CREACAMPO("COMPRAS", "DETASOLINT", "Nume_Solint", 4, 4, 1)
   Call CREACAMPO("COMPRAS", "DETASOLINT", "NuOrd_Item", 3, 2)
   Call CREACAMPO("COMPRAS", "DETASOLINT", "Cod_Inte", 3, 2, 1)
   Call CREACAMPO("COMPRAS", "DETASOLINT", "Descritem", 12, 0)
   Call CREACAMPO("COMPRAS", "DETASOLINT", "Uni_Mas", 10, 4)
   Call CREACAMPO("COMPRAS", "DETASOLINT", "Uni_Sol", 10, 4)
   Call CREACAMPO("COMPRAS", "DETASOLINT", "Cant_Solint", 6, 4)
   Call CREACAMPO("COMPRAS", "DETASOLINT", "Fentre_Sol", 8, 0)
   Call CREACAMPO("COMPRAS", "DETASOLINT", "Coti_Solint", 4, 4, 1)
   Call CREACAMPO("COMPRAS", "DETASOLINT", "Desti_Solint", 10, 255)
   Call CREACAMPO("COMPRAS", "DETASOLINT", "Iconta_Solint", 10, 32)
   Call CREACAMPO("COMPRAS", "DETASOLINT", "Status", 3, 2, 1)
   '
   Call CREACAMPO("COMPRAS", "DETOCABI", "Nume_Ocom", 4, 4, 1)
   Call CREACAMPO("COMPRAS", "DETOCABI", "NuOrd_Item", 3, 2)
   Call CREACAMPO("COMPRAS", "DETOCABI", "Femi_Ocom", 8, 0)
   Call CREACAMPO("COMPRAS", "DETOCABI", "Npro_Ocom", 3, 2, 1)
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
   Call CREACAMPO("COMPRAS", "ENCASOLINT", "Nume_Solint", 4, 4, 1)
   Call CREACAMPO("COMPRAS", "ENCASOLINT", "Femi_Solint", 8, 0, 1)
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
   Call CREACAMPO("COMPRAS", "ENCOCABI", "Nume_Ocom", 4, 4, 1)
   Call CREACAMPO("COMPRAS", "ENCOCABI", "Refe_Ocom", 10, 64)
   Call CREACAMPO("COMPRAS", "ENCOCABI", "Femi_Ocom", 8, 0)
   Call CREACAMPO("COMPRAS", "ENCOCABI", "Npro_Ocom", 3, 2, 1)
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
   Call CREACAMPO("COMPRAS", "MACONSIG", "Nro_Ocom", 10, 14)
   Call CREACAMPO("COMPRAS", "MACONSIG", "Cod_Inte", 3, 0, 1)
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
   Call CreaIndiceCompuesto("MACONSIG", "NPRO_OCOM+NRO_OCOM")
   '
   Call CREACAMPO("COMPRAS", "OCOEMI", "Cod_Inte", 3, 2, 1)
   Call CREACAMPO("COMPRAS", "OCOEMI", "Cantidad", 6, 4)
   Call CREACAMPO("COMPRAS", "OCOEMI", "Fentre_Sol", 8, 0)
   Call CREACAMPO("COMPRAS", "OCOEMI", "Npro_Ocom", 3, 2, 1)
   Call CREACAMPO("COMPRAS", "OCOEMI", "Uni_Mas", 10, 4)
   Call CREACAMPO("COMPRAS", "OCOEMI", "RaSo_Prov", 10, 48)
   Call CREACAMPO("COMPRAS", "OCOEMI", "Cant_Ocom", 6, 4)
   Call CREACAMPO("COMPRAS", "OCOEMI", "FechRecep", 8, 0)
   Call CREACAMPO("COMPRAS", "OCOEMI", "Nume_Ocom", 4, 2, 1)
   Call CREACAMPO("COMPRAS", "OCOEMI", "Periodo", 10, 24)
   Call CREACAMPO("COMPRAS", "OCOEMI", "Cosuni", 6, 0)
   Call CREACAMPO("COMPRAS", "OCOEMI", "StokAct", 6, 0)
   Call CREACAMPO("COMPRAS", "OCOEMI", "Stomin", 6, 0)
   Call CREACAMPO("COMPRAS", "OCOEMI", "Lorepos", 6, 0)
   '
   Call CREACAMPO("COMPRAS", "OCOMDETA", "Nume_Ocom", 4, 4, 1)
   Call CREACAMPO("COMPRAS", "OCOMDETA", "Nume_Pedi", 4, 0)
   Call CREACAMPO("COMPRAS", "OCOMDETA", "NuOrd_Item", 3, 2)
   Call CREACAMPO("COMPRAS", "OCOMDETA", "Femi_Ocom", 8, 0, 1)
   Call CREACAMPO("COMPRAS", "OCOMDETA", "Npro_Ocom", 3, 2, 1)
   Call CREACAMPO("COMPRAS", "OCOMDETA", "Nume_Cli", 4, 0)
   Call CREACAMPO("COMPRAS", "OCOMDETA", "RaSo_Prov", 10, 64)
   Call CREACAMPO("COMPRAS", "OCOMDETA", "Cod_Inte", 3, 0)
   Call CREACAMPO("COMPRAS", "OCOMDETA", "Cod_Exte", 10, 24)
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
   Call CREACAMPO("COMPRAS", "OCOMDETA", "Fentre_Sol", 8, 0, 1)
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
   Call CreaIndiceCompuesto("OCOMDETA", "NPRO_OCOM+STAT_OCOM")
   Call CreaIndiceCompuesto("OCOMDETA", "NPRO_OCOM+COD_INTE+STAT_OCOM")
   Call CreaIndiceCompuesto("OCOMDETA", "NPRO_OCOM+COD_INTE+NUME_OCOM")
   Call CreaIndiceCompuesto("OCOMDETA", "NUME_OCOM+NUORD_ITEM+STAT_OCOM")
   '
   Call CREACAMPO("COMPRAS", "OCOMENCA", "AtSr_Ocom", 10, 64)
   If ValidaExisteCampo%("OCOMENCA_ID", "OCOMENCA") < 1 Then
      SQLX$ = "ALTER TABLE OCOMENCA ADD OCOMENCA_ID INT IDENTITY"
      FLEXCONN.Execute (SQLX$)
   End If
   Call CREACAMPO("COMPRAS", "OCOMENCA", "Obse_Ocom", 12, 0)
   Call CREACAMPO("COMPRAS", "OCOMENCA", "Lentrega", 10, 128)
   Call CREACAMPO("COMPRAS", "OCOMENCA", "Anex_Ocom", 10, 128)
   Call CREACAMPO("COMPRAS", "OCOMENCA", "OcomRevis", 3, 2)
   Call CREACAMPO("COMPRAS", "OCOMENCA", "Nume_Ocom", 4, 4, 1)
   Call CREACAMPO("COMPRAS", "OCOMENCA", "Nume_Pedi", 4, 0)
   Call CREACAMPO("COMPRAS", "OCOMENCA", "Refe_Ocom", 10, 64)
   Call CREACAMPO("COMPRAS", "OCOMENCA", "Femi_Ocom", 8, 0, 1)
   Call CREACAMPO("COMPRAS", "OCOMENCA", "Npro_Ocom", 3, 2, 1)
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
   Call CREACAMPO("COMPRAS", "OCOMENCA", "Respon_Ocompra", 10, 128)
   Call CREACAMPO("COMPRAS", "OCOMENCA", "Via_Entrega", 3, 0)
   Call CREACAMPO("COMPRAS", "OCOMENCA", "COTIZA_DOLAR", 7, 0)
   Call CREACAMPO("COMPRAS", "OCOMENCA", "OComCliente", 10, 16)
   Call CREACAMPO("COMPRAS", "OCOMENCA", "IdCompExt", 4, 0) 'ID DE REGISTRO DE REPLICA DE TABLA MINEPARTS
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
   Call CREACAMPO("CONTAGEN", "DetaAsien", "DatRetPer", 10, 128)
   
   '
   Call CREACAMPO("CONTAGEN", "EncabeAsien", "NuIdAsien", 10, 16, 1)
   Call CREACAMPO("CONTAGEN", "EncabeAsien", "RefeAsien", 10, 64)
   Call CREACAMPO("CONTAGEN", "EncabeAsien", "FechAsien", 8, 0, 1)
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
   Call CREACAMPO("CONTAGEN", "EncabeAsien", "FeReAsien", 8, 8)
   Call CREACAMPO("CONTAGEN", "EncabeAsien", "ORIGEN", 10, 64)
   '
   Call CREACAMPO("CONTAGEN", "PLACOPRE", "CueContaI", 3, 0, 1)
   Call CREACAMPO("CONTAGEN", "PLACOPRE", "CenCosto", 10, 8)
   Call CREACAMPO("CONTAGEN", "PLACOPRE", "Fecha", 8, 0, 1)
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
   Call CREACAMPO("CONTAGEN", "TABEJCON", "CodiEjer", 3, 0, 1)
   Call CREACAMPO("CONTAGEN", "TABEJCON", "DenoEjCon", 10, 48)
   Call CREACAMPO("CONTAGEN", "TABEJCON", "FechAper", 8, 0)
   Call CREACAMPO("CONTAGEN", "TABEJCON", "FechCierr", 8, 0)
   Call CREACAMPO("CONTAGEN", "TABEJCON", "Status", 3, 0)
   Call CREACAMPO("CONTAGEN", "TABEJCON", "Diverif", 3, 0)
   Call CREACAMPO("CONTAGEN", "TABEJCON", "MarcActi", 3, 0)
   '
   Call CREACAMPO("CONTAGEN", "TAPROYEC", "Cod_Proy", 10, 8, 1)
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
   Call CREACAMPO("CUECORRI", "APLICOBRA", "CodiEmpr", 3, 2, 1)
   Call CREACAMPO("CUECORRI", "APLICOBRA", "CodiSucu", 3, 2)
   Call CREACAMPO("CUECORRI", "APLICOBRA", "NuClien", 4, 2, 1)
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
   Call CREACAMPO("CUECORRI", "APLIPAGO", "CodiEmpr", 3, 2, 1)
   Call CREACAMPO("CUECORRI", "APLIPAGO", "CodiSucu", 3, 2)
   Call CREACAMPO("CUECORRI", "APLIPAGO", "NuProve", 3, 2, 1)
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
   Call CREACAMPO("CUECORRI", "CAJABANC", "CodiCom_Lar_ExSimple", 10, 24)
   Call CREACAMPO("CUECORRI", "CAJABANC", "RefeCom", 10, 48)
   Call CREACAMPO("CUECORRI", "CAJABANC", "FechConta", 8, 0, 1)
   Call CREACAMPO("CUECORRI", "CAJABANC", "FechAcred", 8, 0)
   Call CREACAMPO("CUECORRI", "CAJABANC", "MarConsi", 10, 4)
   Call CREACAMPO("CUECORRI", "CAJABANC", "CueContaI", 3, 0, 1)
   Call CREACAMPO("CUECORRI", "CAJABANC", "CenCosto", 10, 8)
   Call CREACAMPO("CUECORRI", "CAJABANC", "ItemConc", 10, 8)
   Call CREACAMPO("CUECORRI", "CAJABANC", "NuClien", 4, 2, 1)
   Call CREACAMPO("CUECORRI", "CAJABANC", "Sucu_Clie", 3, 2)
   Call CREACAMPO("CUECORRI", "CAJABANC", "NuProve", 3, 3, 1)
   Call CREACAMPO("CUECORRI", "CAJABANC", "TipoMovim", 10, 8, 1)
   Call CREACAMPO("CUECORRI", "CAJABANC", "OpciMovim", 10, 8, 1)
   Call CREACAMPO("CUECORRI", "CAJABANC", "ModoOpciMovim", 10, 8, 1)
   Call CREACAMPO("CUECORRI", "CAJABANC", "VariMovim", 10, 8, 1)
   Call CREACAMPO("CUECORRI", "CAJABANC", "CBUAlias", 10, 24)
   Call CREACAMPO("CUECORRI", "CAJABANC", "DescriMov", 10, 128)
   Call CREACAMPO("CUECORRI", "CAJABANC", "FechEmisi", 8, 0, 1)
   Call CREACAMPO("CUECORRI", "CAJABANC", "CodiEmpr", 3, 0, 1)
   Call CREACAMPO("CUECORRI", "CAJABANC", "CodiSucu", 3, 0)
   Call CREACAMPO("CUECORRI", "CAJABANC", "NumeTalo", 3, 0, 1)
   Call CREACAMPO("CUECORRI", "CAJABANC", "NumeComp", 4, 0, 1)
   Call CREACAMPO("CUECORRI", "CAJABANC", "Status", 3, 0, 1)
   Call CREACAMPO("CUECORRI", "CAJABANC", "Mone_Emis", 3, 2)
   Call CREACAMPO("CUECORRI", "CAJABANC", "Tipo_Cam", 7, 0)
   Call CREACAMPO("CUECORRI", "CAJABANC", "Impo_Dola", 7, 0)
   Call CREACAMPO("CUECORRI", "CAJABANC", "ValAbsComp", 7, 0)
   Call CREACAMPO("CUECORRI", "CAJABANC", "IDebeComp", 7, 0)
   Call CREACAMPO("CUECORRI", "CAJABANC", "IHabeComp", 7, 0)
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
   Call CREACAMPO("CUECORRI", "CAJABANC", "Niv_Clasi", 3, 2)
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
   Call CREACAMPO("CUECORRI", "CAJABANC", "MonedaCOTIZA_DOLAR", 3, 2)
   Call CREACAMPO("CUECORRI", "CAJABANC", "COTIZA_DOLAR", 7, 0)
   Call CREACAMPO("CUECORRI", "CAJABANC", "CodigoImpuesto", 3, 0)
   Call CREACAMPO("CUECORRI", "CAJABANC", "CodigoRegimen", 3, 0)
   Call CREACAMPO("CUECORRI", "CAJABANC", "CUIT", 10, 13)
   Call CREACAMPO("CUECORRI", "CAJABANC", "DespaImpo", 10, 24)
   Call CREACAMPO("CUECORRI", "CAJABANC", "Genera_Debito", 3, 0)
   Call CREACAMPO("CUECORRI", "CAJABANC", "POS_IVA", 3, 0)
   Call CREACAMPO("CUECORRI", "CAJABANC", "Cuit_Cli", 10, 24)
   Call CREACAMPO("CUECORRI", "CAJABANC", "NroDeRetencion", 4, 0)
   Call CREACAMPO("CUECORRI", "CAJABANC", "USERNUM", 3, 0)
   '
   Call CreaIndiceCompuesto("CAJABANC", "NUCLIEN+TIPOMOVIM+CODIEMPR", 1)
   Call CreaIndiceCompuesto("CAJABANC", "NUPROVE+TIPOMOVIM+CODIEMPR", 1)
   '
   Call CREACAMPO("", "FACTUANTI", "FechEmisi", 8, 0)
   Call CREACAMPO("", "FACTUANTI", "DOCU_ORIG", 10, 18)
   Call CREACAMPO("", "FACTUANTI", "DOCU_APLI", 10, 18)
   Call CREACAMPO("", "FACTUANTI", "IMPODEBE", 7, 0)
   Call CREACAMPO("", "FACTUANTI", "IMPOHABE", 7, 0)
   Call CREACAMPO("", "FACTUANTI", "STATUS", 3, 0)
   Call CREACAMPO("", "FACTUANTI", "NuClien", 4, 0, 1)
   Call CREACAMPO("", "FACTUANTI", "Cli_Real", 4, 0, 1)
   Call CREACAMPO("", "FACTUANTI", "NumeCom", 4, 0)
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
   Call CREACAMPO("CUECORRI", "DETMOVCC", "N_Mov", 4, 0, 1)
   Call CREACAMPO("CUECORRI", "DETMOVCC", "N_OrIt", 4, 0)
   Call CREACAMPO("CUECORRI", "DETMOVCC", "Fec_Mov", 8, 0)
   Call CREACAMPO("CUECORRI", "DETMOVCC", "N_CueImpu", 3, 0)
   Call CREACAMPO("CUECORRI", "DETMOVCC", "ImporMovI", 7, 0)
   Call CREACAMPO("CUECORRI", "DETMOVCC", "MoneMov", 3, 0)
   Call CREACAMPO("CUECORRI", "DETMOVCC", "ParMonMov", 7, 0)
   '
   Call CREACAMPO("CUECORRI", "DETRECCC", "N_Mov", 4, 0, 1)
   Call CREACAMPO("CUECORRI", "DETRECCC", "T_Comp", 10, 25)
   Call CREACAMPO("CUECORRI", "DETRECCC", "N_Comp", 10, 25)
   Call CREACAMPO("CUECORRI", "DETRECCC", "N_Cli", 4, 0)
   Call CREACAMPO("CUECORRI", "DETRECCC", "Fec_Mov", 8, 0)
   Call CREACAMPO("CUECORRI", "DETRECCC", "N_CueImpu", 4, 0)
   Call CREACAMPO("CUECORRI", "DETRECCC", "ImporMovI", 7, 0)
   Call CREACAMPO("CUECORRI", "DETRECCC", "MoneMov", 3, 0)
   Call CREACAMPO("CUECORRI", "DETRECCC", "ParMonMov", 7, 0)
   '
   Call CREACAMPO("CUECORRI", "ENCMOVCC", "N_Mov", 4, 0, 1)
   Call CREACAMPO("CUECORRI", "ENCMOVCC", "Compro", 10, 25)
   Call CREACAMPO("CUECORRI", "ENCMOVCC", "T_Comp", 10, 25)
   Call CREACAMPO("CUECORRI", "ENCMOVCC", "N_Comp", 4, 0)
   Call CREACAMPO("CUECORRI", "ENCMOVCC", "N_Cli", 4, 0)
   '
   Call CREACAMPO("CUECORRI", "PLAGASPEN", "Fecha", 8, 0, 1)
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
   Call CREACAMPO("CUECORRI", "SADEUPEN", "NuClien", 4, 2, 1)
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
   Call CREACAMPO("CUECORRI", "SIRCREB", "Periodo", 4, 0)
   Call CREACAMPO("CUECORRI", "SIRCREB", "CodigoJurisdiccion", 3, 0)
   Call CREACAMPO("CUECORRI", "SIRCREB", "CodigoProvincia", 10, 6)
   Call CREACAMPO("CUECORRI", "SIRCREB", "Provincia", 10, 64)
   Call CREACAMPO("CUECORRI", "SIRCREB", "CodigoContable", 3, 0)
   Call CREACAMPO("CUECORRI", "SIRCREB", "CuentaContable", 10, 24)
   Call CREACAMPO("CUECORRI", "SIRCREB", "Coeficiente", 7, 8)
   Call CREACAMPO("CUECORRI", "SIRCREB", "FechaUltimaModificacion", 8, 0)
   Call CREACAMPO("CUECORRI", "SIRCREB", "MarcaBorrado", 3, 0)
   
   Call CREACAMPO("CUECORRI", "PADRONIBCABA", "FechaActualizacion", 8, 0)
   Call CREACAMPO("CUECORRI", "PADRONIBCABA", "VigenciaDesde", 8, 0)
   Call CREACAMPO("CUECORRI", "PADRONIBCABA", "VigenciaHasta", 8, 0)
   Call CREACAMPO("CUECORRI", "PADRONIBCABA", "Cuit", 13, 8, 2)
   Call CREACAMPO("CUECORRI", "PADRONIBCABA", "Retencion", 7, 0)
   Call CREACAMPO("CUECORRI", "PADRONIBCABA", "Percepcion", 7, 0)
   Call CREACAMPO("CUECORRI", "PADRONIBCABA", "MarcaBorrado", 3, 0)
   
   Call CREACAMPO("", "GRUVECUECON", "Grucoven", 3, 1)
   Call CREACAMPO("", "GRUVECUECON", "CodInCue", 3, 0)
   
   EPAC$ = TRIM$(UCase$(EMPREAC$))
   If InStr(EPAC$, "MEDITEA") > 0 Then
        Call CREACAMPO("CUECORRI", "APLIREPA", "DOC_APLI", 10, 24)
        If ValidaExisteCampo%("APLIREPA_ID", "APLIREPA") < 1 Then
          SQLX$ = "ALTER TABLE APLIREPA ADD APLIREPA_ID INT IDENTITY"
          FLEXCONN.Execute (SQLX$)
        End If
        Call CREACAMPO("CUECORRI", "APLIREPA", "CodiEmpr", 3, 2, 1)
        Call CREACAMPO("CUECORRI", "APLIREPA", "NuProve", 3, 2, 1)
        Call CREACAMPO("CUECORRI", "APLIREPA", "TEC_EXT", 3, 0, 1)  'TECNICO EXTERNO
        Call CREACAMPO("CUECORRI", "APLIREPA", "NumRepa", 4, 0, 1)
        Call CREACAMPO("CUECORRI", "APLIREPA", "IMP_MOBRA", 7, 8)
        Call CREACAMPO("CUECORRI", "APLIREPA", "FECHAPLI", 8, 0)
        Call CREACAMPO("CUECORRI", "APLIREPA", "IMP_APLI", 7, 8)
        Call CREACAMPO("CUECORRI", "APLIREPA", "SALDO_OP", 7, 8)

   End If
   
   Call CREACAMPO("CUECORRI", "IVADIGCO", "CodiEmpr", 3, 2, 1)
   Call CREACAMPO("CUECORRI", "IVADIGCO", "FechEmisi", 8, 0)
   Call CREACAMPO("CUECORRI", "IVADIGCO", "Codicom_Lar", 10, 18)
   Call CREACAMPO("CUECORRI", "IVADIGCO", "Cuit", 10, 13)
   Call CREACAMPO("CUECORRI", "IVADIGCO", "RazonSocial", 10, 48)
   Call CREACAMPO("CUECORRI", "IVADIGCO", "Moneda", 10, 10)
   Call CREACAMPO("CUECORRI", "IVADIGCO", "TipoCambio", 7, 8)
   Call CREACAMPO("CUECORRI", "IVADIGCO", "ImpoNeto", 7, 8)
   Call CREACAMPO("CUECORRI", "IVADIGCO", "ImpoNoGravado", 7, 8)
   Call CREACAMPO("CUECORRI", "IVADIGCO", "ImpoExento", 7, 8)
   Call CREACAMPO("CUECORRI", "IVADIGCO", "ImpoIva", 7, 8)
   Call CREACAMPO("CUECORRI", "IVADIGCO", "ImpoTotal", 7, 8)

End Sub
'
Sub CREATABLA_EMPLEADOS()
   '
   Call CREACAMPO("EMPLEADOS", "CompxEmple", "Legajo", 3, 2, 1)
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
   EPAC$ = TRIM$(UCase$(EMPREAC$))


   Call CREACAMPO("EQUIPOS", "CONTROLFIN", "NumeroSerie", 10, 32, 2)
   Call CREACAMPO("EQUIPOS", "CONTROLFIN", "Cod_Inte", 3, 1)
   Call CREACAMPO("EQUIPOS", "CONTROLFIN", "NumeroOrden", 3, 1)
   Call CREACAMPO("EQUIPOS", "CONTROLFIN", "Resultado", 10, 4)
   Call CREACAMPO("EQUIPOS", "CONTROLFIN", "MarBorra", 3, 1)
   Call CREACAMPO("EQUIPOS", "CONTROLFIN", "Fech_Ensa", 8, 0)
   Call CREACAMPO("EQUIPOS", "CONTROLFIN", "Revision", 3, 0, 1)
   Call CREACAMPO("EQUIPOS", "CONTROLFIN", "Origen", 10, 16)
   Call CREACAMPO("EQUIPOS", "CONTROLFIN", "NumControl", 4, 0)
   Call CREACAMPO("EQUIPOS", "CONTROLFIN", "Tec_Ensa", 10, 32)
   Call CREACAMPO("EQUIPOS", "CONTROLFIN", "Aext_Ensa", 10, 32)
   Call CREACAMPO("EQUIPOS", "CONTROLFIN", "Aint_Ensa", 10, 32)
   Call CREACAMPO("EQUIPOS", "CONTROLFIN", "Obse_Ensa", 12, 0)
   Call CREACAMPO("EQUIPOS", "CONTROLFIN", "Dic_Ensa", 10, 32)
   '
   Call CreaIndiceCompuesto("CONTROLFIN", "NUMEROSERIE+ORIGEN+REVISION")
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
   Call CREACAMPO("EQUIPOS", "ENSAR070", "Revision", 3, 0, 1)
   Call CREACAMPO("EQUIPOS", "ENSAR070", "Corr_Ensa3", 6, 0)
   Call CREACAMPO("EQUIPOS", "ENSAR070", "Corr_Ensa4", 6, 0)
   Call CREACAMPO("EQUIPOS", "ENSAR070", "Corr_Ensa5", 6, 0)
   Call CREACAMPO("EQUIPOS", "ENSAR070", "Corr_Ensa6", 6, 0)
   Call CREACAMPO("EQUIPOS", "ENSAR070", "Corr_Ensa7", 6, 0)
   Call CREACAMPO("EQUIPOS", "ENSAR070", "Corr_Ensa8", 6, 0)
   Call CREACAMPO("EQUIPOS", "ENSAR070", "Corr_Ensa9", 6, 0)
   Call CREACAMPO("EQUIPOS", "ENSAR070", "Corr_Ensa10", 6, 0)
   Call CREACAMPO("EQUIPOS", "ENSAR070", "CTension3", 6, 0)
   Call CREACAMPO("EQUIPOS", "ENSAR070", "CTension4", 6, 0)
   Call CREACAMPO("EQUIPOS", "ENSAR070", "CTension5", 6, 0)
   Call CREACAMPO("EQUIPOS", "ENSAR070", "CTension6", 6, 0)
   Call CREACAMPO("EQUIPOS", "ENSAR070", "CTension7", 6, 0)
   Call CREACAMPO("EQUIPOS", "ENSAR070", "CTension8", 6, 0)
   Call CREACAMPO("EQUIPOS", "ENSAR070", "CTension9", 6, 0)
   Call CREACAMPO("EQUIPOS", "ENSAR070", "CTension10", 6, 0)
   Call CREACAMPO("EQUIPOS", "ENSAR070", "Veredicto3", 10, 6)
   Call CREACAMPO("EQUIPOS", "ENSAR070", "Veredicto4", 10, 6)
   Call CREACAMPO("EQUIPOS", "ENSAR070", "Veredicto5", 10, 6)
   Call CREACAMPO("EQUIPOS", "ENSAR070", "Veredicto6", 10, 6)
   Call CREACAMPO("EQUIPOS", "ENSAR070", "Veredicto7", 10, 6)
   Call CREACAMPO("EQUIPOS", "ENSAR070", "Veredicto8", 10, 6)
   Call CREACAMPO("EQUIPOS", "ENSAR070", "Veredicto9", 10, 6)
   Call CREACAMPO("EQUIPOS", "ENSAR070", "Veredicto10", 10, 6)
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
   Call CREACAMPO("EQUIPOS", "ENSAR072", "Revision", 3, 0, 1)
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
   Call CREACAMPO("EQUIPOS", "ENSAR074", "Revision", 3, 0, 1)
   Call CREACAMPO("EQUIPOS", "ENSAR074", "INSTRUMENTO1", 10, 32)
   Call CREACAMPO("EQUIPOS", "ENSAR074", "INSTRUMENTO2", 10, 32)
   '
   Call CREACAMPO("EQUIPOS", "MEDIFIN", "NumeroSerie", 10, 32, 2)
   Call CREACAMPO("EQUIPOS", "MEDIFIN", "Cod_Inte", 3, 1)
   Call CREACAMPO("EQUIPOS", "MEDIFIN", "CHN", 10, 6)
   Call CREACAMPO("EQUIPOS", "MEDIFIN", "MedA", 6, 0)
   Call CREACAMPO("EQUIPOS", "MEDIFIN", "MedB", 6, 0)
   Call CREACAMPO("EQUIPOS", "MEDIFIN", "MedC", 6, 0)
   Call CREACAMPO("EQUIPOS", "MEDIFIN", "MarBorra", 3, 1)
   Call CREACAMPO("EQUIPOS", "MEDIFIN", "Fech_Ensa", 8, 0)
   Call CREACAMPO("EQUIPOS", "MEDIFIN", "Revision", 3, 0, 1)
   Call CREACAMPO("EQUIPOS", "MEDIFIN", "Origen", 10, 16)
   '
   Call CreaIndiceCompuesto("MEDIFIN", "NUMEROSERIE+ORIGEN+REVISION")
   '
   Call CREACAMPO("EQUIPOS", "TANUMSE", "NumeroSerie", 10, 32, 2)
   If ValidaExisteCampo%("TANUMSE_ID", "TANUMSE") < 1 Then
    SQLX$ = "ALTER TABLE TANUMSE ADD TANUMSE_ID INT IDENTITY"
    FLEXCONN.Execute (SQLX$)
   End If
   Call CREACAMPO("EQUIPOS", "TANUMSE", "Mesanio", 10, 24)
   Call CREACAMPO("EQUIPOS", "TANUMSE", "NumeOrd", 3, 1)
   Call CREACAMPO("EQUIPOS", "TANUMSE", "Cod_Inte", 3, 1, 1)
   Call CREACAMPO("EQUIPOS", "TANUMSE", "Cod_Exte", 10, 24)
   Call CREACAMPO("EQUIPOS", "TANUMSE", "Descrip", 10, 64)
   Call CREACAMPO("EQUIPOS", "TANUMSE", "IdSubOr", 10, 12, 1)
   Call CREACAMPO("EQUIPOS", "TANUMSE", "FechAlta", 8, 0)
   Call CREACAMPO("EQUIPOS", "TANUMSE", "FechPrueba", 8, 0)
   Call CREACAMPO("EQUIPOS", "TANUMSE", "CodStatusPrueba", 3, 0)
   Call CREACAMPO("EQUIPOS", "TANUMSE", "FechPreparado", 8, 0)
   Call CREACAMPO("EQUIPOS", "TANUMSE", "FechInformado", 8, 0)
   Call CREACAMPO("EQUIPOS", "TANUMSE", "FechEntrega", 8, 0)
   Call CREACAMPO("EQUIPOS", "TANUMSE", "FechRetiro", 8, 0)
   Call CREACAMPO("EQUIPOS", "TANUMSE", "NumRemito", 10, 18)
   Call CREACAMPO("EQUIPOS", "TANUMSE", "NumFactura", 10, 18)
   Call CREACAMPO("EQUIPOS", "TANUMSE", "NumPedido", 4, 0)
   Call CREACAMPO("EQUIPOS", "TANUMSE", "NumOrdPedido", 3, 0)
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
   Call CREACAMPO("EQUIPOS", "TANUMSE", "NumeBoRe", 6, 4, 1)
   Call CREACAMPO("EQUIPOS", "TANUMSE", "NumeroArmado", 4, 0)
   
   '
   Call CreaIndiceCompuesto("TANUMSE", "COD_INTE+IDSUBOR+NUMEROSERIE")
   Call CreaIndiceCompuesto("TANUMSE", "COD_INTE+NUMEROSERIE")
   '
   Call CREACAMPO("EQUIPOS", "SEGELECT", "NumeroSerie", 10, 32, 1)
   Call CREACAMPO("EQUIPOS", "SEGELECT", "Modelo", 3, 1)
   Call CREACAMPO("EQUIPOS", "SEGELECT", "Inspector", 10, 32)
   Call CREACAMPO("EQUIPOS", "SEGELECT", "CorrFuga1", 6, 0)
   Call CREACAMPO("EQUIPOS", "SEGELECT", "Veredicto1", 10, 6)
   Call CREACAMPO("EQUIPOS", "SEGELECT", "CorrFuga2", 6, 0)
   Call CREACAMPO("EQUIPOS", "SEGELECT", "Veredicto2", 10, 6)
   Call CREACAMPO("EQUIPOS", "SEGELECT", "CorrFuga3", 6, 0)
   Call CREACAMPO("EQUIPOS", "SEGELECT", "Veredicto3", 10, 6)
   Call CREACAMPO("EQUIPOS", "SEGELECT", "CorrFuga4", 6, 0)
   Call CREACAMPO("EQUIPOS", "SEGELECT", "Veredicto4", 10, 6)
   Call CREACAMPO("EQUIPOS", "SEGELECT", "Bloqueado", 10, 2)
   Call CREACAMPO("EQUIPOS", "SEGELECT", "VerGeneral", 10, 6)
   Call CREACAMPO("EQUIPOS", "SEGELECT", "Comentarios", 12, 0)
   Call CREACAMPO("EQUIPOS", "SEGELECT", "Fech_Ensa", 8, 0)
   Call CREACAMPO("EQUIPOS", "SEGELECT", "Tecnico", 10, 32)
   Call CREACAMPO("EQUIPOS", "SEGELECT", "Revision", 3, 0, 1)
   Call CREACAMPO("EQUIPOS", "SEGELECT", "Tension", 10, 32)
   Call CREACAMPO("EQUIPOS", "SEGELECT", "Cod_Instrumento", 10, 32)
   Call CREACAMPO("EQUIPOS", "SEGELECT", "Instrumento", 3, 0)
   Call CREACAMPO("EQUIPOS", "SEGELECT", "Clase", 10, 32)
   Call CREACAMPO("EQUIPOS", "SEGELECT", "Frec_Alim", 10, 32)
   Call CREACAMPO("EQUIPOS", "SEGELECT", "CorrClase1", 6, 0)
   Call CREACAMPO("EQUIPOS", "SEGELECT", "CorrClase2", 6, 0)
   Call CREACAMPO("EQUIPOS", "SEGELECT", "CorrClase3", 6, 0)
   Call CREACAMPO("EQUIPOS", "SEGELECT", "CorrClase4", 6, 0)
   
   Call CREACAMPO("EQUIPOS", "REGIDATO", "Fech_Ensa", 8, 0)
   Call CREACAMPO("EQUIPOS", "REGIDATO", "CONTROLO", 10, 32)
   Call CREACAMPO("EQUIPOS", "REGIDATO", "Aprobo", 10, 32)
   Call CREACAMPO("EQUIPOS", "REGIDATO", "Revision", 3, 0, 1)
   Call CREACAMPO("EQUIPOS", "REGIDATO", "OP", 3, 0, 1)
   Call CREACAMPO("EQUIPOS", "REGIDATO", "Descripcion", 10, 128)
   Call CREACAMPO("EQUIPOS", "REGIDATO", "Rango_Acep_Min", 10, 16)
   Call CREACAMPO("EQUIPOS", "REGIDATO", "Rango_Acep_Max", 10, 16)
   Call CREACAMPO("EQUIPOS", "REGIDATO", "VM", 10, 32)
   Call CREACAMPO("EQUIPOS", "REGIDATO", "VALOR_RESUL", 10, 24) 'VALOR O RESULTADO
   Call CREACAMPO("EQUIPOS", "REGIDATO", "UNIMED", 10, 16)
   Call CREACAMPO("EQUIPOS", "REGIDATO", "NumeroSerie", 10, 32, 1)
   Call CREACAMPO("EQUIPOS", "REGIDATO", "Bloqueado", 10, 2)
   Call CREACAMPO("EQUIPOS", "REGIDATO", "Modelo", 3, 1)
   Call CREACAMPO("EQUIPOS", "REGIDATO", "Comentarios", 10, 512)
   Call CREACAMPO("EQUIPOS", "REGIDATO", "IdSubOr", 10, 16, 1)
   Call CREACAMPO("EQUIPOS", "REGIDATO", "DP006", 10, 8)
   Call CREACAMPO("EQUIPOS", "REGIDATO", "DP100", 10, 8)
   Call CREACAMPO("EQUIPOS", "REGIDATO", "RESULTADO", 10, 8)
   Call CREACAMPO("EQUIPOS", "REGIDATO", "DISPOMEDINUM", 3, 0)
   Call CREACAMPO("EQUIPOS", "REGIDATO", "DISPOMEDIX", 10, 24)
  
   Call CREACAMPO("EQUIPOS", "CONTFINEYM", "CONTROLO", 10, 32)
   If ValidaExisteCampo%("CONTFINEYM_ID", "CONTFINEYM") < 1 Then
    SQLX$ = "ALTER TABLE CONTFINEYM ADD CONTFINEYM_ID INT IDENTITY"
    FLEXCONN.Execute (SQLX$)
   End If

   Call CREACAMPO("EQUIPOS", "CONTFINEYM", "Fech_Ensa", 8, 0)
   Call CREACAMPO("EQUIPOS", "CONTFINEYM", "Aprobo", 10, 32)
   Call CREACAMPO("EQUIPOS", "CONTFINEYM", "Revision", 3, 0, 1)
   Call CREACAMPO("EQUIPOS", "CONTFINEYM", "OP", 3, 0, 1) 'ITEM
   Call CREACAMPO("EQUIPOS", "CONTFINEYM", "Descripcion", 10, 128)
   Call CREACAMPO("EQUIPOS", "CONTFINEYM", "Rango_Acep_Min", 10, 16)
   Call CREACAMPO("EQUIPOS", "CONTFINEYM", "Rango_Acep_Max", 10, 16)
   Call CREACAMPO("EQUIPOS", "CONTFINEYM", "VM", 10, 32)
   Call CREACAMPO("EQUIPOS", "CONTFINEYM", "VALOR_RESUL", 10, 24) 'VALOR O RESULTADO
   Call CREACAMPO("EQUIPOS", "CONTFINEYM", "UNIMED", 10, 16)
   Call CREACAMPO("EQUIPOS", "CONTFINEYM", "NumeroSerie", 10, 32, 1)
   Call CREACAMPO("EQUIPOS", "CONTFINEYM", "Bloqueado", 10, 2)
   Call CREACAMPO("EQUIPOS", "CONTFINEYM", "Modelo", 3, 1)
   Call CREACAMPO("EQUIPOS", "CONTFINEYM", "Comentarios", 10, 512)
   Call CREACAMPO("EQUIPOS", "CONTFINEYM", "IdSubOr", 10, 16, 1)
   Call CREACAMPO("EQUIPOS", "CONTFINEYM", "RESULTADO", 10, 8)
   Call CREACAMPO("EQUIPOS", "CONTFINEYM", "OK_O_MAL", 10, 8)
   Call CREACAMPO("EQUIPOS", "CONTFINEYM", "COD_EXTE", 10, 24)
   
   
   'TABLA DE RELACIONES DE N.SERIE CON EQUPIPOS SUPERIORES EYM
   Call CREACAMPO("EQUIPOS", "RELASERIE", "IDSUBOR", 10, 16)
   If ValidaExisteCampo%("RELASERIE_ID", "RELASERIE") < 1 Then
    SQLX$ = "ALTER TABLE RELASERIE ADD RELASERIE_ID INT IDENTITY"
    FLEXCONN.Execute (SQLX$)
   End If
   Call CREACAMPO("EQUIPOS", "RELASERIE", "NumeroSerie", 10, 32, 1)
   Call CREACAMPO("EQUIPOS", "RELASERIE", "ORD", 3, 0)
   Call CREACAMPO("EQUIPOS", "RELASERIE", "COD_INTE", 3, 1)
   Call CREACAMPO("EQUIPOS", "RELASERIE", "COD_EXTE", 10, 24)
   Call CREACAMPO("EQUIPOS", "RELASERIE", "DESCRIPCION", 10, 128)
   Call CREACAMPO("EQUIPOS", "RELASERIE", "MARBORRA", 3, 1)
   
   'TABLA DE TRAZABILIDAD DE COMPONENTES  EYM
   Call CREACAMPO("EQUIPOS", "TRAZAXSERIE", "COD_EXTE", 10, 24)
   If ValidaExisteCampo%("TRAZAXSERIE_ID", "TRAZAXSERIE") < 1 Then
    SQLX$ = "ALTER TABLE TRAZAXSERIE ADD TRAZAXSERIE_ID INT IDENTITY"
    FLEXCONN.Execute (SQLX$)
   End If
   Call CREACAMPO("EQUIPOS", "TRAZAXSERIE", "IDSUBOR", 10, 16)
   Call CREACAMPO("EQUIPOS", "TRAZAXSERIE", "OF_ARRASTRA", 10, 16)
   Call CREACAMPO("EQUIPOS", "TRAZAXSERIE", "IDENLOTE", 10, 32)
   Call CREACAMPO("EQUIPOS", "TRAZAXSERIE", "NumeroSerie", 10, 32, 1)
   Call CREACAMPO("EQUIPOS", "TRAZAXSERIE", "ORD", 3, 0)
   Call CREACAMPO("EQUIPOS", "TRAZAXSERIE", "TRAZABLE", 3, 0, 1)
   Call CREACAMPO("EQUIPOS", "TRAZAXSERIE", "COD_INTE", 3, 1)
   Call CREACAMPO("EQUIPOS", "TRAZAXSERIE", "DESCRIPCION", 10, 128)
   Call CREACAMPO("EQUIPOS", "TRAZAXSERIE", "Canti", 7, 8)
   Call CREACAMPO("EQUIPOS", "TRAZAXSERIE", "FechAlta", 8, 0)
   Call CREACAMPO("EQUIPOS", "TRAZAXSERIE", "MARBORRA", 3, 1)
   Call CREACAMPO("EQUIPOS", "TRAZAXSERIE", "StatuOrf", 3, 1, 1)
   Call CREACAMPO("EQUIPOS", "TRAZAXSERIE", "CONSUMIDO", 3, 1, 1)
   Call CREACAMPO("EQUIPOS", "TRAZAXSERIE", "DESV_TRANS", 10, 64)
   Call CREACAMPO("EQUIPOS", "TRAZAXSERIE", "PREPARADOPOR", 10, 64)
   
   Call CREACAMPO("EQUIPOS", "DATOSCPU", "CODIGO", 10, 24)
   If ValidaExisteCampo%("DATOSCPU_ID", "DATOSCPU") < 1 Then
    SQLX$ = "ALTER TABLE DATOSCPU ADD DATOSCPU_ID INT IDENTITY"
    FLEXCONN.Execute (SQLX$)
   End If
   Call CREACAMPO("EQUIPOS", "DATOSCPU", "COD_INTE", 3, 1)
   Call CREACAMPO("EQUIPOS", "DATOSCPU", "DESCRIPCION", 10, 24)
   Call CREACAMPO("EQUIPOS", "DATOSCPU", "IDENTIFICACION", 10, 24)
   Call CREACAMPO("EQUIPOS", "DATOSCPU", "DESV_TRANS", 10, 64)
   Call CREACAMPO("EQUIPOS", "DATOSCPU", "NumeroSerie", 10, 32, 1)
   Call CREACAMPO("EQUIPOS", "DATOSCPU", "IDSUBOR", 10, 16)
  
   '
   

   Call CREACAMPO("EQUIPOS", "CONTROEYM", "COD_EXTE", 10, 24)
   If ValidaExisteCampo%("CONTROEYM_ID", "CONTROEYM") < 1 Then
    SQLX$ = "ALTER TABLE CONTROEYM ADD CONTROEYM_ID INT IDENTITY"
    FLEXCONN.Execute (SQLX$)
   End If
   Call CREACAMPO("EQUIPOS", "CONTROEYM", "COD_INTE", 3, 1)
   Call CREACAMPO("EQUIPOS", "CONTROEYM", "DESCRIPCION", 10, 128)
   Call CREACAMPO("EQUIPOS", "CONTROEYM", "C_FRMCONTROL", 10, 32)
   Call CREACAMPO("EQUIPOS", "CONTROEYM", "REF_CONTROL", 10, 32)
   Call CREACAMPO("EQUIPOS", "CONTROEYM", "EXCEPCIONES", 10, 32)
   Call CREACAMPO("EQUIPOS", "CONTROEYM", "MARBORRA", 3, 1)
   
   
   Call CREACAMPO("EQUIPOS", "DES_TRANS", "DESCRIPCION", 10, 256)
   If ValidaExisteCampo%("DES_TRANS_ID", "DES_TRANS") < 1 Then
    SQLX$ = "ALTER TABLE DES_TRANS ADD DES_TRANS_ID INT IDENTITY"
    FLEXCONN.Execute (SQLX$)
   End If
   '
   Call CREACAMPO("EQUIPOS", "DES_TRANS", "DESV_TRANS", 4, 1)
   Call CREACAMPO("EQUIPOS", "DES_TRANS", "REF_DESVIO", 10, 256)
   Call CREACAMPO("EQUIPOS", "DES_TRANS", "IDENLOTE", 10, 32)
   Call CREACAMPO("EQUIPOS", "DES_TRANS", "NUMEROSERIE", 10, 32)
   Call CREACAMPO("EQUIPOS", "DES_TRANS", "COD_INTE", 3, 1)
   Call CREACAMPO("EQUIPOS", "DES_TRANS", "COD_EXTE", 10, 24)
   Call CREACAMPO("EQUIPOS", "DES_TRANS", "IDSUBOR", 10, 16)
   Call CREACAMPO("EQUIPOS", "DES_TRANS", "FecheEmis", 8, 0)
   Call CREACAMPO("EQUIPOS", "DES_TRANS", "Situacion", 10, 512)
   '
   If InStr(EPAC$, "EYM") > 0 Then
    Call CREACAMPO("EQUIPOS", "LOTEEYM", "NUMEROSERIE", 10, 24)
    Call CREACAMPO("EQUIPOS", "LOTEEYM", "IdSubOr", 10, 12, 1)
    Call CREACAMPO("EQUIPOS", "LOTEEYM", "CodiElem", 3, 0, 1)
    Call CREACAMPO("EQUIPOS", "LOTEEYM", "CodIConj", 3, 0, 1)
    Call CREACAMPO("EQUIPOS", "LOTEEYM", "IDENLOTE", 10, 32)
    Call CREACAMPO("EQUIPOS", "LOTEEYM", "FeReMovi", 8, 0)
     If ValidaExisteCampo%("LOTEEYM_ID", "LOTEEYM") < 1 Then
       SQLX$ = "ALTER TABLE LOTEEYM ADD LOTEEYM_ID INT IDENTITY"
       FLEXCONN.Execute (SQLX$)
     End If
  End If
   
End Sub
'
Sub CREATABLA_MAQUINARIAS()
   '
   Call CREACAMPO("MAQUINARIAS", "CONTADORES", "Num_Con", 3, 2, 1)
   Call CREACAMPO("MAQUINARIAS", "CONTADORES", "Referen_Con", 10, 32)
   Call CREACAMPO("MAQUINARIAS", "CONTADORES", "Freq_Con", 6, 4)
   Call CREACAMPO("MAQUINARIAS", "CONTADORES", "FUL_Con", 8, 2)
   Call CREACAMPO("MAQUINARIAS", "CONTADORES", "Val_Con", 7, 8)
   Call CREACAMPO("MAQUINARIAS", "CONTADORES", "FPL_Con", 8, 2)
   Call CREACAMPO("MAQUINARIAS", "CONTADORES", "Prom_Con", 7, 8)
   Call CREACAMPO("MAQUINARIAS", "CONTADORES", "Observa_Con", 10, 64)
   Call CREACAMPO("MAQUINARIAS", "CONTADORES", "Codigo_Equ", 10, 6)
   '
   Call CREACAMPO("MAQUINARIAS", "EQUIPOS", "Codigo_Equ", 10, 6, 1)
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
   Call CREACAMPO("MAQUINARIAS", "GALEIMAG", "Codigo_Equ", 10, 6, 1)
   Call CREACAMPO("MAQUINARIAS", "GALEIMAG", "Norit_Imag", 3, 2)
   Call CREACAMPO("MAQUINARIAS", "GALEIMAG", "Ruta_Imag", 12, 0)
   Call CREACAMPO("MAQUINARIAS", "GALEIMAG", "Refe_Imag", 12, 0)
   '
   Call CREACAMPO("MAQUINARIAS", "PLAMANPREV", "Codigo_Equ", 10, 6, 1)
   Call CREACAMPO("MAQUINARIAS", "PLAMANPREV", "CodLar_Equ", 10, 16)
   Call CREACAMPO("MAQUINARIAS", "PLAMANPREV", "Nro_Sub", 4, 4)
   Call CREACAMPO("MAQUINARIAS", "PLAMANPREV", "Orden_Tar", 3, 2)
   Call CREACAMPO("MAQUINARIAS", "PLAMANPREV", "Descri_Tar", 10, 128)
   Call CREACAMPO("MAQUINARIAS", "PLAMANPREV", "Frecu_Tar", 4, 4)
   Call CREACAMPO("MAQUINARIAS", "PLAMANPREV", "Uni_Frecu", 10, 16)
   Call CREACAMPO("MAQUINARIAS", "PLAMANPREV", "Obse_Tar", 10, 128)
   '
   Call CREACAMPO("MAQUINARIAS", "REPUMAQ", "Codigo_Equ", 10, 6, 1)
   Call CREACAMPO("MAQUINARIAS", "REPUMAQ", "CodInte_Rep", 3, 2)
   Call CREACAMPO("MAQUINARIAS", "REPUMAQ", "CodExte_Rep", 10, 16)
   Call CREACAMPO("MAQUINARIAS", "REPUMAQ", "Descrip_Rep", 10, 64)
   Call CREACAMPO("MAQUINARIAS", "REPUMAQ", "UM_Rep", 10, 4)
   Call CREACAMPO("MAQUINARIAS", "REPUMAQ", "Cant_Rep", 6, 4)
   Call CREACAMPO("MAQUINARIAS", "REPUMAQ", "Observa_Rep", 10, 128)
   '
   Call CREACAMPO("MAQUINARIAS", "SUBCONJUNTOS", "Nro_Sub", 3, 2, 1)
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
   Call CREACAMPO("OFABRIC", "DOCANRE", "NumRepa", 4, 0, 1)
   Call CREACAMPO("OFABRIC", "DOCANRE", "NomAsignado", 10, 256)
   Call CREACAMPO("OFABRIC", "DOCANRE", "FILE_NAME", 10, 256)
   Call CREACAMPO("OFABRIC", "DOCANRE", "FILE_DATA", 12, 0)
   '
   Call CREACAMPO("OFABRIC", "GALIMREP", "NumRepa", 4, 0, 1)
   Call CREACAMPO("OFABRIC", "GALIMREP", "Descripcion", 12, 0)
   Call CREACAMPO("OFABRIC", "GALIMREP", "Norit_Imag", 3, 2)
   Call CREACAMPO("OFABRIC", "GALIMREP", "Nom_Imag", 10, 256)
   Call CREACAMPO("OFABRIC", "GALIMREP", "Imagen", 12, 0)
   '
   Call CREACAMPO("OFABRIC", "INFOFAB", "IdSubOr", 10, 12, 1)
   EPAC$ = TRIM$(UCase$(EMPREAC$))
   If InStr(EPAC$, "CONTEPLAST") > 0 Or InStr(EPAC$, "PERGAMINO") > 0 Then
    If ValidaExisteCampo%("INFOFAB_ID", "INFOFAB") < 1 Then
      SQLX$ = "ALTER TABLE INFOFAB ADD INFOFAB_ID INT IDENTITY"
      FLEXCONN.Execute (SQLX$)
    End If
   End If
   Call CREACAMPO("OFABRIC", "INFOFAB", "DescriOF", 10, 64)
   Call CREACAMPO("OFABRIC", "INFOFAB", "NuOrInf", 3, 2)
   Call CREACAMPO("OFABRIC", "INFOFAB", "Cod_Inte", 3, 2, 1)
   Call CREACAMPO("OFABRIC", "INFOFAB", "Cod_Exte", 10, 24)
   Call CREACAMPO("OFABRIC", "INFOFAB", "Descrip", 10, 64)
   Call CREACAMPO("OFABRIC", "INFOFAB", "NumeOper", 3, 3)
   
   Call CREACAMPO("OFABRIC", "INFOFAB", "AltOper", 3, 0)
   Call CREACAMPO("OFABRIC", "INFOFAB", "DescOper", 10, 48)
   Call CREACAMPO("OFABRIC", "INFOFAB", "ReTraba", 3, 0)
   Call CREACAMPO("OFABRIC", "INFOFAB", "AtMulti", 3, 0)
   Call CREACAMPO("OFABRIC", "INFOFAB", "CanOrPro", 7, 8)
   Call CREACAMPO("OFABRIC", "INFOFAB", "CantReal", 7, 8)
   Call CREACAMPO("OFABRIC", "INFOFAB", "CantOrigCte", 7, 8) 'cantidad original del corte cuando se informa la primera vez
   Call CREACAMPO("OFABRIC", "INFOFAB", "CaNoConf", 7, 8)
   Call CREACAMPO("OFABRIC", "INFOFAB", "CORTE", 7, 8)
   Call CREACAMPO("OFABRIC", "INFOFAB", "Maquina", 10, 6)
   Call CREACAMPO("OFABRIC", "INFOFAB", "DescMaq", 10, 64)
   Call CREACAMPO("OFABRIC", "INFOFAB", "LegaOps", 11, 64)
   Call CREACAMPO("OFABRIC", "INFOFAB", "FeHorLib", 8, 0)
   Call CREACAMPO("OFABRIC", "INFOFAB", "FeHorInf", 8, 0)
   Call CREACAMPO("OFABRIC", "INFOFAB", "Hora_Des", 3, 0)
   Call CREACAMPO("OFABRIC", "INFOFAB", "Hora_Has", 3, 0)
   Call CREACAMPO("OFABRIC", "INFOFAB", "Hora_Fri", 3, 0)
   Call CREACAMPO("OFABRIC", "INFOFAB", "MaqLiber", 10, 1)
   Call CREACAMPO("OFABRIC", "INFOFAB", "TiemEjec", 6, 4)
   Call CREACAMPO("OFABRIC", "INFOFAB", "TiemStan", 6, 4)
   Call CREACAMPO("OFABRIC", "INFOFAB", "Rendim", 6, 4)
   Call CREACAMPO("OFABRIC", "INFOFAB", "Observa", 10, 64)
   Call CREACAMPO("OFABRIC", "INFOFAB", "Referen", 12, 0)
   Call CREACAMPO("OFABRIC", "INFOFAB", "MotiRecha", 3, 0)
   Call CREACAMPO("OFABRIC", "INFOFAB", "Status", 3, 0)
   Call CREACAMPO("OFABRIC", "INFOFAB", "NroInfNC", 10, 32)
   Call CREACAMPO("OFABRIC", "INFOFAB", "RefeInf", 10, 64)
   Call CREACAMPO("OFABRIC", "INFOFAB", "UserControl", 10, 64)
   Call CREACAMPO("OFABRIC", "INFOFAB", "NUMOPERARIO", 3, 0)
   Call CREACAMPO("OFABRIC", "INFOFAB", "NAMEOPERARIO", 10, 64)
   Call CREACAMPO("OFABRIC", "INFOFAB", "Operarios", 12, 0)
   Call CREACAMPO("OFABRIC", "INFOFAB", "IdInfofab", 10, 24, 1)
   Call CREACAMPO("OFABRIC", "INFOFAB", "IdenLote", 10, 32)
   Call CREACAMPO("OFABRIC", "INFOFAB", "CodiElem", 3, 0, 1)
   Call CREACAMPO("OFABRIC", "INFOFAB", "CodXElem", 10, 24)
   Call CREACAMPO("OFABRIC", "INFOFAB", "FACTORCORTE", 3, 0)
   Call CREACAMPO("OFABRIC", "INFOFAB", "COORDUBIC", 10, 64)
   Call CREACAMPO("OFABRIC", "INFOFAB", "IDCOORDUBIC", 4, 0, 1)
   Call CREACAMPO("OFABRIC", "INFOFAB", "OperFucionada", 4, 0)  ' INFORMA EN QUE CON QUE ID (FILA) SE FUSIONO - CONTEPLAST
   'Call CREACAMPO("OFABRIC", "INFOFAB", "Oper_Consume", 3, 0) ' O CALCULA  CONSUMO , -1 NO SE CONSIDEERA PARA CONSUMIR conteplast
   Call CREACAMPO("OFABRIC", "INFOFAB", "Peso", 7, 0)
   Call CREACAMPO("OFABRIC", "INFOFAB", "IdRegPadre", 6, 0)
   Call CREACAMPO("OFABRIC", "INFOFAB", "Procesado", 3, 0)
   Call CREACAMPO("OFABRIC", "INFOFAB", "FeTranLinea", 8, 0)
   Call CREACAMPO("OFABRIC", "INFOFAB", "CONSUMO", 7, 8)
   Call CREACAMPO("OFABRIC", "INFOFAB", "IDMOVISTO", 7, 0, 1) 'ANOTO EL MOVISTO_ID CUANDO SE INFORMA EL CORTE EN FRM INFOCORTE
   Call CREACAMPO("OFABRIC", "INFOFAB", "CantSalid", 7, 8)
   Call CREACAMPO("OFABRIC", "INFOFAB", "CodiMovi", 10, 2, 1)
   '
   Call CREACAMPO("OFABRIC", "INFOPROY", "Nuor_Inf", 4, 4, 1)
   Call CREACAMPO("OFABRIC", "INFOPROY", "Nume_Inf", 10, 12)
   Call CREACAMPO("OFABRIC", "INFOPROY", "Fecha_Inf", 8, 0)
   Call CREACAMPO("OFABRIC", "INFOPROY", "Refer_Inf", 10, 64)
   Call CREACAMPO("OFABRIC", "INFOPROY", "Cant_Inf", 6, 4)
   Call CREACAMPO("OFABRIC", "INFOPROY", "Proy_Inf", 10, 8)
   Call CREACAMPO("OFABRIC", "INFOPROY", "LegaOps_Inf", 11, 64)
   Call CREACAMPO("OFABRIC", "INFOPROY", "Total_Inf", 7, 8)
   '
   Call CREACAMPO("OFABRIC", "NEBRUNE", "Cod_Inte", 3, 2, 1)
   Call CREACAMPO("OFABRIC", "NEBRUNE", "Cant_Reque", 7, 4)
   Call CREACAMPO("OFABRIC", "NEBRUNE", "Fech_Nec", 8, 0)
   Call CREACAMPO("OFABRIC", "NEBRUNE", "Referen", 10, 64)
   Call CREACAMPO("OFABRIC", "NEBRUNE", "Destino", 10, 32)
   '
   Call CREACAMPO("OFABRIC", "OPERFAB", "IdSubOr", 10, 12, 1)
   Call CREACAMPO("OFABRIC", "OPERFAB", "Cod_Inte", 3, 0, 1)
   Call CREACAMPO("OFABRIC", "OPERFAB", "Cod_Exte", 10, 24)
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
   Call CREACAMPO("OFABRIC", "ORDEFABR", "NumeOrFa", 4, 4, 1)
   Call CREACAMPO("OFABRIC", "ORDEFABR", "Referen", 10, 128)
   Call CREACAMPO("OFABRIC", "ORDEFABR", "IdSubOr", 10, 12, 1)
   Call CREACAMPO("OFABRIC", "ORDEFABR", "FechGen", 8, 0)
   Call CREACAMPO("OFABRIC", "ORDEFABR", "Cod_Inte", 3, 0, 1)
   Call CREACAMPO("OFABRIC", "ORDEFABR", "Cod_Exte", 10, 24)
   Call CREACAMPO("OFABRIC", "ORDEFABR", "Descrip", 10, 128)
   Call CREACAMPO("OFABRIC", "ORDEFABR", "Unidad", 10, 32)
   Call CREACAMPO("OFABRIC", "ORDEFABR", "CanProy", 7, 8)
   Call CREACAMPO("OFABRIC", "ORDEFABR", "Destino", 10, 32)
   Call CREACAMPO("OFABRIC", "ORDEFABR", "FechEnSol", 8, 0)
   Call CREACAMPO("OFABRIC", "ORDEFABR", "Priorid", 3, 1)
   Call CREACAMPO("OFABRIC", "ORDEFABR", "StatuOrf", 3, 1, 1)
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
   Call CREACAMPO("OFABRIC", "ORDEFABR", "Armador ", 4, 0)
   Call CREACAMPO("OFABRIC", "ORDEFABR", "LoteIndivis", 7, 8)
   Call CREACAMPO("OFABRIC", "ORDEFABR", "inicio_informado", 8, 0)
   Call CREACAMPO("OFABRIC", "ORDEFABR", "FinDeProceso", 8, 0)
   Call CREACAMPO("OFABRIC", "ORDEFABR", "Canti_fin_proceso", 7, 8)
   Call CREACAMPO("OFABRIC", "ORDEFABR", "Status_Proceso", 3, 0)
   Call CREACAMPO("OFABRIC", "ORDEFABR", "Canti_Iniciada", 7, 8)
   Call CREACAMPO("OFABRIC", "ORDEFABR", "NumeroChasis", 10, 24) ' pin (regalia)
   Call CREACAMPO("OFABRIC", "ORDEFABR", "TipoChasis", 10, 24)
   Call CREACAMPO("OFABRIC", "ORDEFABR", "LineaInterno", 10, 8)
   Call CREACAMPO("OFABRIC", "ORDEFABR", "FechEntMat", 8, 0) 'Fecha de entrega del material para poder iniciar la of (regalia)
   Call CREACAMPO("OFABRIC", "ORDEFABR", "DespachoFabrica", 8, 0)
   Call CREACAMPO("OFABRIC", "ORDEFABR", "NuOrNroPed", 3, 0)
   Call CREACAMPO("OFABRIC", "ORDEFABR", "ProdEstandar", 4, 0) 'Producción Standar Conteplast
   Call CREACAMPO("OFABRIC", "ORDEFABR", "Anexos", 10, 128) 'Producción Standar Conteplast
   Call CREACAMPO("OFABRIC", "ORDEFABR", "CantHilos", 3, 0)
   Call CREACAMPO("OFABRIC", "ORDEFABR", "PesoBolsa", 7, 8)
   '
   Call CREACAMPO("OFABRIC", "RECONCILIA", "IdSubOr", 10, 16, 1)
   Call CREACAMPO("OFABRIC", "RECONCILIA", "Cod_Inte", 3, 0, 1)
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
   Call CREACAMPO("OFABRIC", "REPARA", "FechaCierreRepa", 8, 0)
   Call CREACAMPO("OFABRIC", "REPARA", "Nume_Cli", 4, 0, 1)
   Call CREACAMPO("OFABRIC", "REPARA", "Raso_Cli", 10, 128)
   Call CREACAMPO("OFABRIC", "REPARA", "Domi_Cli", 10, 64)
   Call CREACAMPO("OFABRIC", "REPARA", "Loca_Cli", 10, 64)
   Call CREACAMPO("OFABRIC", "REPARA", "Pais_Cli", 10, 64)
   Call CREACAMPO("OFABRIC", "REPARA", "Tele_Cli", 10, 24)
   Call CREACAMPO("OFABRIC", "REPARA", "Mail_Cli", 10, 48)
   Call CREACAMPO("OFABRIC", "REPARA", "Cod_Inte", 3, 0, 1)
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
   Call CREACAMPO("OFABRIC", "REPARA", "RefOrdRep", 10, 128)
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
   Call CREACAMPO("OFABRIC", "REPARA", "DESCTO1", 7, 0) 'DESCUENTO 1 DEL PRESUPUESTO
   Call CREACAMPO("OFABRIC", "REPARA", "DESCTO2", 7, 0) 'DESCUENTO 2 DEL PRESUPUESTO X MEDIO DE COBRANZA
   Call CREACAMPO("OFABRIC", "REPARA", "MED_COBRANZA", 3, 0) 'MEDIO DE COBRANZA TABLA MEDITEA
   Call CREACAMPO("OFABRIC", "REPARA", "NUIDASIEN", 10, 16)
   Call CREACAMPO("OFABRIC", "REPARA", "TEC_EXT", 3, 0, 1) 'TECNICO EXTERNO
   Call CREACAMPO("OFABRIC", "REPARA", "TALON_TEC", 10, 24)
   Call CREACAMPO("OFABRIC", "REPARA", "MOBRA_EN_GTIA", 7, 0) 'MANO DE OBRA EN GARANTIA CUANDO TIENE TECNICO EXTERNO
   Call CREACAMPO("OFABRIC", "REPARA", "ObsLibre", 10, 128)
   Call CREACAMPO("OFABRIC", "REPARA", "MoneEmis", 3, 2)
   Call CREACAMPO("OFABRIC", "REPARA", "TipoCam", 6, 4)
   Call CREACAMPO("OFABRIC", "REPARA", "STAT_APROBACION", 3, 0)
   Call CREACAMPO("OFABRIC", "REPARA", "CondPag", 3, 0)
   
  
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
   Call CREACAMPO("OFABRIC", "REPUREPA", "PORDESC", 7, 0)
   Call CREACAMPO("OFABRIC", "REPUREPA", "MarcaItemAgregado", 3, 0)
   
   '
   Call CREACAMPO("OFABRIC", "RESERVASF", "Cod_Inte", 3, 2)
   '
   Call CREACAMPO("OFABRIC", "SECCIFAB", "Seccion", 7, 8, 1)
   Call CREACAMPO("OFABRIC", "SECCIFAB", "FechaSeteo", 8, 0)
   Call CREACAMPO("OFABRIC", "SECCIFAB", "Periodo", 10, 16)
   Call CREACAMPO("OFABRIC", "SECCIFAB", "HorasAusencia", 7, 0)
   Call CREACAMPO("OFABRIC", "SECCIFAB", "MotivoAusencia", 10, 64)
   Call CREACAMPO("OFABRIC", "SECCIFAB", "Observaciones", 10, 64)
   Call CREACAMPO("OFABRIC", "SECCIFAB", "CantidadEmpleados", 4, 0)
   Call CREACAMPO("OFABRIC", "SECCIFAB", "DiasLaborables", 4, 0)
   Call CREACAMPO("OFABRIC", "SECCIFAB", "NumOrden", 4, 0)
   Call CREACAMPO("OFABRIC", "SECCIFAB", "MarBorra", 4, 0)
   '
   Call CREACAMPO("OFABRIC", "ARMADOPEDIDOS", "NumeroArmado", 4, 0, 1)
   Call CREACAMPO("OFABRIC", "ARMADOPEDIDOS", "FechaEmision", 8, 0)
   Call CREACAMPO("OFABRIC", "ARMADOPEDIDOS", "Armador", 3, 0)
   Call CREACAMPO("OFABRIC", "ARMADOPEDIDOS", "Usuario", 10, 24)
   '
   Call CREACAMPO("OFABRIC", "SERIESBOMBAS", "Serie", 10, 24)
   Call CREACAMPO("OFABRIC", "SERIESBOMBAS", "FechaAsignacion", 8, 0)
   Call CREACAMPO("OFABRIC", "SERIESBOMBAS", "Codigo", 3, 0, 1)
   Call CREACAMPO("OFABRIC", "SERIESBOMBAS", "TAG", 10, 24)
   Call CREACAMPO("OFABRIC", "SERIESBOMBAS", "Yacimiento", 10, 24)
   Call CREACAMPO("OFABRIC", "SERIESBOMBAS", "Cliente", 3, 0, 1)
   Call CREACAMPO("OFABRIC", "SERIESBOMBAS", "NotaDeVenta", 3, 0, 1)
   Call CREACAMPO("OFABRIC", "SERIESBOMBAS", "FechaNotaDeVenta", 8, 0)
   Call CREACAMPO("OFABRIC", "SERIESBOMBAS", "OrdenCompra", 10, 24)
   Call CREACAMPO("OFABRIC", "SERIESBOMBAS", "FechaOrdenCompra", 8, 0)
   Call CREACAMPO("OFABRIC", "SERIESBOMBAS", "CodigoCliente", 10, 24)
   Call CREACAMPO("OFABRIC", "SERIESBOMBAS", "Motriz", 10, 24)
   Call CREACAMPO("OFABRIC", "SERIESBOMBAS", "CuerpoHidraulico", 10, 24)
   Call CREACAMPO("OFABRIC", "SERIESBOMBAS", "Ciguenal", 10, 24)
   Call CREACAMPO("OFABRIC", "SERIESBOMBAS", "Corona", 10, 24)
   Call CREACAMPO("OFABRIC", "SERIESBOMBAS", "Pinon", 10, 24)
   Call CREACAMPO("OFABRIC", "SERIESBOMBAS", "Reductor", 10, 24)
   Call CREACAMPO("OFABRIC", "SERIESBOMBAS", "Valvula", 10, 24)
   Call CREACAMPO("OFABRIC", "SERIESBOMBAS", "Motor", 10, 24)
   Call CREACAMPO("OFABRIC", "SERIESBOMBAS", "AmortigSuccion", 10, 24)
   Call CREACAMPO("OFABRIC", "SERIESBOMBAS", "AmortigDescarga", 10, 24)
   Call CREACAMPO("OFABRIC", "SERIESBOMBAS", "CarretelSuccion", 10, 24)
   Call CREACAMPO("OFABRIC", "SERIESBOMBAS", "CarretelDescarga", 10, 24)
   Call CREACAMPO("OFABRIC", "SERIESBOMBAS", "PresostatoSuccion", 10, 24)
   Call CREACAMPO("OFABRIC", "SERIESBOMBAS", "PresostatoDescarga", 10, 24)
   Call CREACAMPO("OFABRIC", "SERIESBOMBAS", "TableroElectrico", 10, 24)
   Call CREACAMPO("OFABRIC", "SERIESBOMBAS", "BombaCentrifuga", 10, 24)
   Call CREACAMPO("OFABRIC", "SERIESBOMBAS", "MarBorra", 3, 0)
   
   Call CREACAMPO("OFABRIC", "BULTOSFABRIC", "Fecha", 8, 0)
   Call CREACAMPO("OFABRIC", "BULTOSFABRIC", "IdPallet", 4, 0)
   Call CREACAMPO("OFABRIC", "BULTOSFABRIC", "OrdenFabricacion", 10, 24)
   Call CREACAMPO("OFABRIC", "BULTOSFABRIC", "IdBulto", 4, 0)
   Call CREACAMPO("OFABRIC", "BULTOSFABRIC", "CantidadBulto", 7, 8)
   Call CREACAMPO("OFABRIC", "BULTOSFABRIC", "Pedido", 4, 0)
   Call CREACAMPO("OFABRIC", "BULTOSFABRIC", "Procesado", 3, 0)
   Call CREACAMPO("OFABRIC", "BULTOSFABRIC", "Remito", 10, 18)
   Call CREACAMPO("OFABRIC", "BULTOSFABRIC", "Status", 3, 0)
   
   
   Call CREA_TABLA_USUARIOSSQL
End Sub
Sub CREA_TABLA_USUARIOSSQL()
   
   Call CREACAMPO("", "USUARIOSSQL", "FECHA", 8, 0)
   Call CREACAMPO("", "USUARIOSSQL", "NUSER", 3, 0, 1)
   Call CREACAMPO("", "USUARIOSSQL", "NOMBRE", 10, 48)
   Call CREACAMPO("", "USUARIOSSQL", "CLAVE", 10, 48)
   Call CREACAMPO("", "USUARIOSSQL", "NIVELCONF", 3, 0)
   Call CREACAMPO("", "USUARIOSSQL", "IDENTIFICADOR", 10, 48)
   Call CREACAMPO("", "USUARIOSSQL", "MARBORRA", 3, 0)
   Call CREACAMPO("", "USUARIOSSQL", "IDMAQUINA", 10, 12)
   
   
   
   
End Sub
Sub CREATABLA_GRUCOVENSQL()
    Call CREACAMPO("", "GRUCOVENSQL", "DESCRIPCION", 10, 48)
    Call CREACAMPO("", "GRUCOVENSQL", "CODIGO", 3, 0, 1)
    
    If ValidaExisteCampo%("GRUCOVENSQL_ID", "USUARIOSSQL") < 1 Then
        SQLX$ = "ALTER TABLE USUARIOSSQL ADD GRUCOVENSQL_ID INT IDENTITY"
        FLEXCONN.Execute (SQLX$)
    End If

    

End Sub
'
Sub CREATABLA_OMANTIN()
   '
   Call CREACAMPO("OMANTIN", "ORDEMANT", "NumeOrMa_OM", 4, 4, 1)
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
   Call CREACAMPO("OMANTIN", "ORECORRI", "NumORec", 4, 0, 1)
   Call CREACAMPO("OMANTIN", "ORECORRI", "DesORec", 10, 128)
   Call CREACAMPO("OMANTIN", "ORECORRI", "FechEmis", 8, 0)
   Call CREACAMPO("OMANTIN", "ORECORRI", "FechCierr", 8, 0)
   Call CREACAMPO("OMANTIN", "ORECORRI", "ObseEmis", 12, 0)
   Call CREACAMPO("OMANTIN", "ORECORRI", "ObseCierr", 12, 0)
   '
   Call CREACAMPO("OMANTIN", "RECUPROP", "NumeOrMa_OM", 4, 4, 1)
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

Sub CREATABLA_CALIDAD()
   '
   Call CREACAMPO("CALIDAD", "INFOCALIDAD", "NumeroInforme", 4, 4, 1)
   Call CREACAMPO("CALIDAD", "INFOCALIDAD", "FechaInforme", 8, 0)
   Call CREACAMPO("CALIDAD", "INFOCALIDAD", "FechaReal", 8, 0)
   Call CREACAMPO("CALIDAD", "INFOCALIDAD", "NumeroUsuario", 3, 0)
   Call CREACAMPO("CALIDAD", "INFOCALIDAD", "DocumentoOrigen", 10, 18)
   Call CREACAMPO("CALIDAD", "INFOCALIDAD", "Cod_Inte", 3, 0)
   Call CREACAMPO("CALIDAD", "INFOCALIDAD", "NumeroDeControl", 10, 32)
   Call CREACAMPO("CALIDAD", "INFOCALIDAD", "CantidadAprobada", 6, 4)
   Call CREACAMPO("CALIDAD", "INFOCALIDAD", "CantidadRechazada", 6, 4)
   Call CREACAMPO("CALIDAD", "INFOCALIDAD", "Observaciones", 10, 128)
   Call CREACAMPO("CALIDAD", "INFOCALIDAD", "EstadoInforme", 3, 2)
   Call CREACAMPO("CALIDAD", "INFOCALIDAD", "Encapsulado", 10, 24)
   Call CREACAMPO("CALIDAD", "INFOCALIDAD", "Embalaje", 10, 24)
   Call CREACAMPO("CALIDAD", "INFOCALIDAD", "Certificado", 10, 24)
   Call CREACAMPO("CALIDAD", "INFOCALIDAD", "Status", 10, 12)
   Call CREACAMPO("CALIDAD", "INFOCALIDAD", "NombreFotoCalidad", 10, 32)
   Call CREACAMPO("CALIDAD", "INFOCALIDAD", "PATHFOTOCALIDAD", 10, 192)
   '
End Sub
'
Sub CREATABLA_ARCHIVOS_QR()
   '
   Call CREACAMPO("", "ARCHIVOS_QR", "Fecha", 8, 0)
   Call CREACAMPO("", "ARCHIVOS_QR", "FechaReal", 8, 0)
   Call CREACAMPO("", "ARCHIVOS_QR", "NumeroUsuario", 3, 0)
   Call CREACAMPO("", "ARCHIVOS_QR", "DocumentoOrigen", 10, 18)
   Call CREACAMPO("", "ARCHIVOS_QR", "PATH_IMAG", 10, 256)
   '
End Sub

Sub CREATABLA_OTRABAJO()
   '
   Call CREACAMPO("OTRABAJO", "GRUPOTRA", "Nro_Otrab", 10, 16, 1)
   Call CREACAMPO("OTRABAJO", "GRUPOTRA", "Nro_Grup", 10, 10)
   Call CREACAMPO("OTRABAJO", "GRUPOTRA", "Prog_Grup", 10, 16)
   Call CREACAMPO("OTRABAJO", "GRUPOTRA", "Comp_Grup", 10, 16)
   Call CREACAMPO("OTRABAJO", "GRUPOTRA", "Deno_Grup", 10, 64)
   Call CREACAMPO("OTRABAJO", "GRUPOTRA", "Observa_Grup", 10, 64)
   Call CREACAMPO("OTRABAJO", "GRUPOTRA", "Pla05_Grup", 10, 10)
   Call CREACAMPO("OTRABAJO", "GRUPOTRA", "Pla07_Grup", 10, 10)
   '
   Call CREACAMPO("OTRABAJO", "LISTAMAT", "Nro_Otrab", 10, 16, 1)
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
   Call CREACAMPO("OTRABAJO", "ORDETRAB", "Nro_Otrab", 10, 16, 1)
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
   Call CREACAMPO("OTRABAJO", "PLACOB", "Nro_Otrab", 10, 16, 1)
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
   Call CREACAMPO("PRESUCAM", "DESPECIAL", "Nro_Presu", 10, 16, 1)
   Call CREACAMPO("PRESUCAM", "DESPECIAL", "Num_Orden", 3, 2)
   Call CREACAMPO("PRESUCAM", "DESPECIAL", "Referen", 12, 0)
   Call CREACAMPO("PRESUCAM", "DESPECIAL", "Descrip", 12, 0)
   Call CREACAMPO("PRESUCAM", "DESPECIAL", "Importe", 7, 8)
   '
   Call CREACAMPO("PRESUCAM", "DETAPRESU", "Nro_Presu", 10, 16, 1)
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
   Call CREACAMPO("PRESUCAM", "DOCADJUN", "Nro_Presu", 10, 16, 1)
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
   Call CREACAMPO("PRESUCAM", "ENCAPRESU", "Nro_Presu", 10, 16#)
   Call CREACAMPO("PRESUCAM", "ENCAPRESU", "Refe_Presu", 10, 64)
   Call CREACAMPO("PRESUCAM", "ENCAPRESU", "Nclie_Presu", 3, 2, 1)
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
   Call CREACAMPO("STOCKORIGI", "DETASTOCK", "N_Com", 4, 0, 1)
   Call CREACAMPO("STOCKORIGI", "DETASTOCK", "C_Tmv", 10, 10)
   Call CREACAMPO("STOCKORIGI", "DETASTOCK", "Fecha_Mov", 8, 0)
   Call CREACAMPO("STOCKORIGI", "DETASTOCK", "C_Sig", 10, 1)
   Call CREACAMPO("STOCKORIGI", "DETASTOCK", "C_Dep", 10, 2)
   Call CREACAMPO("STOCKORIGI", "DETASTOCK", "C_Art", 10, 16)
   Call CREACAMPO("STOCKORIGI", "DETASTOCK", "Canti", 7, 8)
   Call CREACAMPO("STOCKORIGI", "DETASTOCK", "N_Trn", 4, 0)
   '
   Call CREACAMPO("STOCKORIGI", "ENCASTOCK", "C_Com", 10, 10)
   Call CREACAMPO("STOCKORIGI", "ENCASTOCK", "N_Com", 4, 0, 1)
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
   Call CREACAMPO("STOCKS", "BALAREQMO", "NumeCompro", 4, 4, 1)
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
   Call CREACAMPO("STOCKS", "BOLRECEP", "FEYHORecep", 8, 0)
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
   Call CREACAMPO("STOCKS", "BOLRECEP", "LoteCol_Prov", 10, 48)
   Call CREACAMPO("STOCKS", "BOLRECEP", "CerCal_Prov", 10, 24)
   Call CREACAMPO("STOCKS", "BOLRECEP", "NumRem_Prov", 10, 16)
   Call CREACAMPO("STOCKS", "BOLRECEP", "NumFac_Prov", 10, 16)
   Call CREACAMPO("STOCKS", "BOLRECEP", "PatenteTransporte", 10, 16)
   Call CREACAMPO("STOCKS", "BOLRECEP", "NumConVig", 10, 24)
   Call CREACAMPO("STOCKS", "BOLRECEP", "NumOC", 4, 4, 1)
   Call CREACAMPO("STOCKS", "BOLRECEP", "NroItemOC", 3, 0)
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
   Call CREACAMPO("STOCKS", "BOLRECEP", "Peso", 7, 0)
   Call CREACAMPO("STOCKS", "BOLRECEP", "ProvExterior", 4, 0, 1) 'proveedor real caso mineparts
   Call CREACAMPO("STOCKS", "BOLRECEP", "PjeNacionaliz", 7, 0) 'pjenacionalizacion del proveedor ext mineparts
   Call CREACAMPO("STOCKS", "BOLRECEP", "ViaExterior", 3, 0) 'via de entrega ahp del prov.exterior mineparts
   Call CREACAMPO("STOCKS", "BOLRECEP", "KgBalanzaPublica", 7, 0)
   Call CREACAMPO("STOCKS", "BOLRECEP", "CantidadBultos", 3, 0)
   '
   Call CREACAMPO("STOCKS", "BOLRETMP", "idRecepcion", 4, 0, 1)
   Call CREACAMPO("STOCKS", "BOLRETMP", "OrdenItem", 4, 0, 1)
   Call CREACAMPO("STOCKS", "BOLRETMP", "Cod_Inte", 4, 0, 0)
   Call CREACAMPO("STOCKS", "BOLRETMP", "OrdenCompra", 4, 0, 0)
   Call CREACAMPO("STOCKS", "BOLRETMP", "OrdenItemCompra", 4, 0, 0)
   Call CREACAMPO("STOCKS", "BOLRETMP", "FechaEntrega", 8, 0, 0)
   Call CREACAMPO("STOCKS", "BOLRETMP", "CantidadPendiente", 7, 8)
   Call CREACAMPO("STOCKS", "BOLRETMP", "Lote", 10, 32)
   Call CREACAMPO("STOCKS", "BOLRETMP", "Recibido", 7, 8)
   Call CREACAMPO("STOCKS", "BOLRETMP", "Cantidad_1", 7, 8)
   Call CREACAMPO("STOCKS", "BOLRETMP", "Cantidad_2", 7, 8)
   Call CREACAMPO("STOCKS", "BOLRETMP", "Status", 3, 0, 0)
   Call CREACAMPO("STOCKS", "BOLRETMP", "CantidadBultos", 3, 0)
   '
   Call CREACAMPO("STOCKS", "DETACARTEM", "Num_Os", 10, 16, 1)
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
   Call CREACAMPO("STOCKS", "DETAOSER", "NuOrdSe", 4, 4, 1)
   Call CREACAMPO("STOCKS", "DETAOSER", "NuOrden", 3, 2)
   Call CREACAMPO("STOCKS", "DETAOSER", "Modelo", 3, 2)
   Call CREACAMPO("STOCKS", "DETAOSER", "NumeroSerie", 10, 32)
   Call CREACAMPO("STOCKS", "DETAOSER", "HorasFunc", 3, 2)
   '
   Call CREACAMPO("STOCKS", "ENCACARTEM", "FechEntre", 8, 0)
   Call CREACAMPO("STOCKS", "ENCACARTEM", "Num_Os", 10, 16)
   Call CREACAMPO("STOCKS", "ENCACARTEM", "Num_Entre", 10, 16)
   '
   Call CREACAMPO("STOCKS", "ENCAOSER", "NuOrdSe", 4, 4, 1)
   Call CREACAMPO("STOCKS", "ENCAOSER", "FeOrdSe", 8, 0)
   Call CREACAMPO("STOCKS", "ENCAOSER", "Nume_Cli", 3, 2)
   Call CREACAMPO("STOCKS", "ENCAOSER", "TipoOrdSe", 3, 2)
   Call CREACAMPO("STOCKS", "ENCAOSER", "Observa", 12, 0)
   '
   Call CREACAMPO("STOCKS", "ENCAREMI", "DoSecLar", 10, 24, 1)
   Call CREACAMPO("STOCKS", "ENCAREMI", "DoPriLar", 10, 24, 1)
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
   Call CREACAMPO("STOCKS", "LISEMPAQ", "NumPallet", 3, 0, 1)
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
   Call CREACAMPO("STOCKS", "LISEMPAQ", "NroPed", 4, 0, 1)
   '
   Call CREACAMPO("STOCKS", "LOTINGRE", "Cod_Inte", 3, 2, 1)
   Call CREACAMPO("STOCKS", "LOTINGRE", "IdenLote", 10, 32, 1)
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
   Call CREACAMPO("STOCKS", "MACONSIG", "Npro_Ocom", 3, 2, 1)
   Call CREACAMPO("STOCKS", "MACONSIG", "RaSo_Prov", 10, 64)
   Call CREACAMPO("STOCKS", "MACONSIG", "Secc_Ejec", 6, 4)
   Call CREACAMPO("STOCKS", "MACONSIG", "Deno_Secc", 10, 64)
   Call CREACAMPO("STOCKS", "MACONSIG", "Nro_Orfa", 10, 12, 1)
   Call CREACAMPO("STOCKS", "MACONSIG", "Nro_Ocom", 10, 12, 1)
   Call CREACAMPO("STOCKS", "MACONSIG", "Cod_Inte", 3, 0, 1)
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
   Call CREACAMPO("STOCKS", "MASUMCLI", "Nume_Cli", 3, 0, 1)
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
   Call CREACAMPO("STOCKS", "MOVIREPA", "IdenMovi", 10, 32)
   EPAC$ = TRIM$(UCase$(EMPREAC$))
   If InStr(EPAC$, "CONTEPLAST") > 0 Then
    If ValidaExisteCampo%("MOVIREPA_ID", "MOVIREPA") < 1 Then
      SQLX$ = "ALTER TABLE MOVIREPA ADD MOVIREPA_ID INT IDENTITY"
      FLEXCONN.Execute (SQLX$)
    End If
   End If
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
   Call CREACAMPO("STOCKS", "MOVIREPA", "DoPriCor", 10, 12, 1)
   Call CREACAMPO("STOCKS", "MOVIREPA", "DoPriLar", 10, 24, 1)
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
   Call CREACAMPO("STOCKS", "MOVIREPA", "Nume_Clie", 3, 0, 1)
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
   Call CREACAMPO("STOCKS", "MOVIREPA", "NroEnvio", 4, 0)
   Call CREACAMPO("STOCKS", "MOVIREPA", "NUMPVTA", 3, 0, 1)
   Call CREACAMPO("STOCKS", "MOVIREPA", "DOPRINUM", 4, 0, 1)
   '
   Call CREACAMPO("STOCKS", "MOVISTO", "IdenMovi", 10, 32)
   If InStr(EPAC$, "CONTEPLAST") > 0 Or InStr(EPAC$, "CAMBACERES") > 0 Or InStr(EPAC$, "AHP") > 0 Or InStr(EPAC$, "EYM") > 0 Or InStr(EPAC$, "PERGAMINO") > 0 Then
    If ValidaExisteCampo%("MOVISTO_ID", "MOVISTO") < 1 Then
      SQLX$ = "ALTER TABLE MOVISTO ADD MOVISTO_ID INT IDENTITY"
      FLEXCONN.Execute (SQLX$)
    End If
   End If
 
   Call CREACAMPO("STOCKS", "MOVISTO", "PESOKG", 7, 0)
   Call CREACAMPO("STOCKS", "MOVISTO", "METROS", 7, 0)
   Call CREACAMPO("STOCKS", "MOVISTO", "FeReMovi", 8, 0)
   Call CREACAMPO("STOCKS", "MOVISTO", "NumUsuar", 3, 0)
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
   Call CREACAMPO("STOCKS", "MOVISTO", "CantIngre", 7, 0)
   Call CREACAMPO("STOCKS", "MOVISTO", "CantSalid", 7, 0)
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
   Call CREACAMPO("STOCKS", "MOVISTO", "NUMPVTA", 3, 0, 1)
   Call CREACAMPO("STOCKS", "MOVISTO", "DOPRINUM", 4, 0, 1)
   Call CREACAMPO("STOCKS", "MOVISTO", "IdCalidad", 4, 0)
   Call CREACAMPO("STOCKS", "MOVISTO", "ID_OPERACION", 3, 0)
   
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
   Call CREACAMPO("STOCKS", "MOVIVIE", "NUMPVTA", 3, 0)
   Call CREACAMPO("STOCKS", "MOVIVIE", "DOPRINUM", 4, 0, 1)
   '
   Call CREACAMPO("STOCKS", "RESERVA", "Cod_Exte", 10, 16)
   If ValidaExisteCampo%("RESERVA_ID", "RESERVA") < 1 Then
    SQLX$ = "ALTER TABLE RESERVA ADD RESERVA_ID INT IDENTITY"
    FLEXCONN.Execute (SQLX$)
   End If
   
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
   Call CREACAMPO("STOCKS", "RESERVA", "NUMEROSERIE", 10, 24)
   
   '
   Call CREACAMPO("STOCKS", "SALSTOCK", "CodiInte", 3, 2)
   Call CREACAMPO("STOCKS", "SALSTOCK", "VeSalSto", 11, 256)
   '
   Call CREACAMPO("STOCKS", "TRANMOVI", "CodiCompro", 10, 5)
   Call CREACAMPO("STOCKS", "TRANMOVI", "NumeCompro", 4, 4, 1)
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
   Call CREACAMPO("STOCKS", "TRANRECEP", "NumeCompro", 4, 4, 1)
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
   Call CREACAMPO("STOCKS", "TRAZARECE", "Orden", 3, 0)
   Call CREACAMPO("STOCKS", "TRAZARECE", "NumeroColada", 3, 0)
   Call CREACAMPO("STOCKS", "TRAZARECE", "Material", 10, 24)
   Call CREACAMPO("STOCKS", "TRAZARECE", "Observaciones", 10, 48)
   Call CREACAMPO("STOCKS", "TRAZARECE", "Modelo", 10, 24)
   Call CREACAMPO("STOCKS", "TRAZARECE", "Marca", 10, 24)
   Call CREACAMPO("STOCKS", "TRAZARECE", "NumeroSerieOrigen", 10, 24)
   Call CREACAMPO("STOCKS", "TRAZARECE", "ClienteDestino", 3, 0)
   Call CREACAMPO("STOCKS", "TRAZARECE", "Seteo", 10, 24)
   Call CREACAMPO("STOCKS", "TRAZARECE", "Alambre", 10, 24)
   Call CREACAMPO("STOCKS", "TRAZARECE", "Cfamilia", 10, 24)
   Call CREACAMPO("STOCKS", "TRAZARECE", "Origen", 10, 24, 1)
   Call CREACAMPO("STOCKS", "TRAZARECE", "Destino", 10, 24, 1)
   Call CREACAMPO("STOCKS", "TRAZARECE", "CodInte_DestinoConsumo", 3, 0, 1)
   Call CREACAMPO("STOCKS", "TRAZARECE", "CodExte_DestinoConsumo", 10, 24)
   Call CREACAMPO("STOCKS", "TRAZARECE", "DestinoConsumo", 10, 24, 1)
   Call CREACAMPO("STOCKS", "TRAZARECE", "Status", 3, 0)
   Call CREACAMPO("STOCKS", "TRAZARECE", "MotivoBaja", 10, 150)
   '
   Call CREACAMPO("STOCKS", "MOVITRANSTO", "FeReMovi", 8, 0)
   Call CREACAMPO("STOCKS", "MOVITRANSTO", "NumUsuar", 3, 0)
   Call CREACAMPO("STOCKS", "MOVITRANSTO", "IdenMovi", 10, 32)
   Call CREACAMPO("STOCKS", "MOVITRANSTO", "DeLaRemi", 12, 0)
   Call CREACAMPO("STOCKS", "MOVITRANSTO", "Descri_Lar", 12, 0)
   Call CREACAMPO("STOCKS", "MOVITRANSTO", "FechMov", 8, 0, 1)
   Call CREACAMPO("STOCKS", "MOVITRANSTO", "Cod_Inte", 3, 0, 1)
   Call CREACAMPO("STOCKS", "MOVITRANSTO", "Cod_Exte", 10, 24)
   Call CREACAMPO("STOCKS", "MOVITRANSTO", "Descrip", 10, 64)
   Call CREACAMPO("STOCKS", "MOVITRANSTO", "IdenLote", 10, 16, 1)
   Call CREACAMPO("STOCKS", "MOVITRANSTO", "LotConsum", 12, 0)
   Call CREACAMPO("STOCKS", "MOVITRANSTO", "PNL_Des", 10, 32)
   Call CREACAMPO("STOCKS", "MOVITRANSTO", "PLC_Has", 10, 32)
   Call CREACAMPO("STOCKS", "MOVITRANSTO", "CodiMovi", 10, 2, 1)
   Call CREACAMPO("STOCKS", "MOVITRANSTO", "DoPriCor", 10, 12, 1)
   Call CREACAMPO("STOCKS", "MOVITRANSTO", "DoPriLar", 10, 24, 1)
   Call CREACAMPO("STOCKS", "MOVITRANSTO", "DoSecCor", 10, 12, 1)
   Call CREACAMPO("STOCKS", "MOVITRANSTO", "DoSecLar", 10, 24, 1)
   Call CREACAMPO("STOCKS", "MOVITRANSTO", "DocuOrig", 10, 24)
   Call CREACAMPO("STOCKS", "MOVITRANSTO", "NumRem_Cli", 10, 16)
   Call CREACAMPO("STOCKS", "MOVITRANSTO", "NuorItem", 3, 2)
   Call CREACAMPO("STOCKS", "MOVITRANSTO", "NuPedCli", 4, 4, 1)
   Call CREACAMPO("STOCKS", "MOVITRANSTO", "NumeOrdPed", 3, 2)
   Call CREACAMPO("STOCKS", "MOVITRANSTO", "ReferCor", 10, 64)
   Call CREACAMPO("STOCKS", "MOVITRANSTO", "CostoRep", 7, 0)
   Call CREACAMPO("STOCKS", "MOVITRANSTO", "ValoUnit", 7, 0)
   Call CREACAMPO("STOCKS", "MOVITRANSTO", "ValoUnitIva", 7, 0)
   Call CREACAMPO("STOCKS", "MOVITRANSTO", "ValTotIt", 7, 0)
   Call CREACAMPO("STOCKS", "MOVITRANSTO", "ValTotItIva", 7, 0)
   Call CREACAMPO("STOCKS", "MOVITRANSTO", "MoneVal", 3, 2)
   Call CREACAMPO("STOCKS", "MOVITRANSTO", "TipoCam", 7, 8)
   Call CREACAMPO("STOCKS", "MOVITRANSTO", "DepoMovi", 3, 0, 1)
   Call CREACAMPO("STOCKS", "MOVITRANSTO", "Nume_Clie", 4, 0, 1)
   Call CREACAMPO("STOCKS", "MOVITRANSTO", "Sucu_Clie", 3, 0)
   Call CREACAMPO("STOCKS", "MOVITRANSTO", "Nume_Prov", 3, 0)
   Call CREACAMPO("STOCKS", "MOVITRANSTO", "CantIngre", 7, 0)
   Call CREACAMPO("STOCKS", "MOVITRANSTO", "CantSalid", 7, 0)
   Call CREACAMPO("STOCKS", "MOVITRANSTO", "CantSaldo", 7, 0)
   Call CREACAMPO("STOCKS", "MOVITRANSTO", "Verificado", 3, 1)
   Call CREACAMPO("STOCKS", "MOVITRANSTO", "Causa", 3, 0)
   Call CREACAMPO("STOCKS", "MOVITRANSTO", "User_Tran", 10, 32)
   Call CREACAMPO("STOCKS", "MOVITRANSTO", "CodCarpImp", 10, 12)
   Call CREACAMPO("STOCKS", "MOVITRANSTO", "VerificadoPor", 3, 0)
   Call CREACAMPO("STOCKS", "MOVITRANSTO", "FHor_Verif", 8, 0)
   Call CREACAMPO("STOCKS", "MOVITRANSTO", "TipoTran", 3, 0)
   Call CREACAMPO("STOCKS", "MOVITRANSTO", "OCOMPRA_RTOABIERTO", 10, 64)
   Call CREACAMPO("STOCKS", "MOVITRANSTO", "NroEnvio", 4, 0)
   Call CREACAMPO("STOCKS", "MOVITRANSTO", "NUMPVTA", 3, 0, 1)
   Call CREACAMPO("STOCKS", "MOVITRANSTO", "DOPRINUM", 4, 0, 1)
   Call CREACAMPO("STOCKS", "MOVITRANSTO", "IdCalidad", 4, 0)
   Call CREACAMPO("STOCKS", "MOVITRANSTO", "OrdenRegistroMovimiento", 3, 2)
   Call CREACAMPO("STOCKS", "MOVITRANSTO", "CantidadRechazada", 7, 0)
   '
   Call CREACAMPO("STOCKS", "MAXREMITO", "CodiMovi", 10, 2, 1)
   Call CREACAMPO("STOCKS", "MAXREMITO", "NUMPVTA", 3, 0, 1)
   Call CREACAMPO("STOCKS", "MAXREMITO", "CODIEMPR", 3, 0)
   Call CREACAMPO("STOCKS", "MAXREMITO", "FechMov", 8, 0)
   Call CREACAMPO("STOCKS", "MAXREMITO", "DOPRINUM", 4, 0)
   '
   Call CREACAMPO("STOCKS", "SALOTEDEP", "COD_INTE", 4, 0, 1)
   Call CREACAMPO("STOCKS", "SALOTEDEP", "CODIEMPR", 3, 0)
   Call CREACAMPO("STOCKS", "SALOTEDEP", "IdenLote", 10, 16)
   Call CREACAMPO("STOCKS", "SALOTEDEP", "SALDO", 7, 0)
   Call CREACAMPO("STOCKS", "SALOTEDEP", "DepoMovi", 3, 0, 1)
   
   Call CREACAMPO("STOCKS", "SALOTEDEP", "FechMov", 8, 0, 1)
   Call CREACAMPO("STOCKS", "SALOTEDEP", "Nume_Cli", 4, 0, 1)
   '
   'TABLA DE RESPALDO PARA MEDITEA QUE UTILZA EN LA GENERACION DEL REMITO OFICIAL
   '*****************************************************************************
   Call CREACAMPO("STOCKS", "MOVIOFI", "Cod_Exte", 10, 24)
   If ValidaExisteCampo%("MOVIOFI_ID", "MOVIOFI") < 1 Then
    SQLX$ = "ALTER TABLE MOVIOFI ADD MOVIOFI_ID INT IDENTITY"
    FLEXCONN.Execute (SQLX$)
   End If
   Call CREACAMPO("STOCKS", "MOVIOFI", "FeReMovi", 8, 0)
   Call CREACAMPO("STOCKS", "MOVIOFI", "FechMov", 8, 0, 1)
   Call CREACAMPO("STOCKS", "MOVIOFI", "Cod_Inte", 3, 0, 1)
   Call CREACAMPO("STOCKS", "MOVIOFI", "Descrip", 10, 64)
   Call CREACAMPO("STOCKS", "MOVIOFI", "IdenLote", 10, 16)
   Call CREACAMPO("STOCKS", "MOVIOFI", "DoPriLar", 10, 24, 1)
   Call CREACAMPO("STOCKS", "MOVIOFI", "NuPedCli", 4, 4, 1)
   Call CREACAMPO("STOCKS", "MOVIOFI", "Nume_Clie", 4, 0, 1)
   Call CREACAMPO("STOCKS", "MOVIOFI", "CantIngre", 7, 0)
   Call CREACAMPO("STOCKS", "MOVIOFI", "CantSalid", 7, 0)
   Call CREACAMPO("STOCKS", "MOVIOFI", "DOPRINUM", 4, 0, 1)
   Call CREACAMPO("STOCKS", "MOVIOFI", "User_Tran", 10, 32)
   Call CREACAMPO("STOCKS", "MOVIOFI", "REMIOFI", 10, 32, 1)
   Call CREACAMPO("STOCKS", "MOVIOFI", "REFERENCIA", 10, 32)
   Call CREACAMPO("STOCKS", "MOVIOFI", "STATUS", 3, 0)
   '
   'TABLA PARA CONTENER LOS CIERRES DE OF PENDIENTES DE APROBACION
   '**************************************************************
   Call CREACAMPO("STOCKS", "ORFAPEN", "FechMov", 8, 0, 1)
   Call CREACAMPO("STOCKS", "ORFAPEN", "Cod_Inte", 3, 0, 1)
   Call CREACAMPO("STOCKS", "ORFAPEN", "Cod_Exte", 10, 24)
   Call CREACAMPO("STOCKS", "ORFAPEN", "Descrip", 10, 64)
   Call CREACAMPO("STOCKS", "ORFAPEN", "IdenLote", 10, 16, 1)
   Call CREACAMPO("STOCKS", "ORFAPEN", "CodiMovi", 10, 2, 1)
   Call CREACAMPO("STOCKS", "ORFAPEN", "DoPriCor", 10, 12, 1)
   Call CREACAMPO("STOCKS", "ORFAPEN", "DoPriLar", 10, 24, 1)
   Call CREACAMPO("STOCKS", "ORFAPEN", "DoSecCor", 10, 12, 1)
   Call CREACAMPO("STOCKS", "ORFAPEN", "DoSecLar", 10, 24, 1)
   Call CREACAMPO("STOCKS", "ORFAPEN", "ReferCor", 10, 64)
   Call CREACAMPO("STOCKS", "ORFAPEN", "ValoUnit", 7, 0)
   Call CREACAMPO("STOCKS", "ORFAPEN", "MoneVal", 3, 2)
   Call CREACAMPO("STOCKS", "ORFAPEN", "DepoMovi", 3, 0, 1)
   Call CREACAMPO("STOCKS", "ORFAPEN", "DepoConsu", 3, 0, 1)
   
   Call CREACAMPO("STOCKS", "ORFAPEN", "Nume_Clie", 4, 0, 1)
   Call CREACAMPO("STOCKS", "ORFAPEN", "Nume_Prov", 3, 0)
   Call CREACAMPO("STOCKS", "ORFAPEN", "CantIngre", 7, 0)
   Call CREACAMPO("STOCKS", "ORFAPEN", "CantSalid", 7, 0)
   Call CREACAMPO("STOCKS", "ORFAPEN", "CantSaldo", 7, 0)
   Call CREACAMPO("STOCKS", "ORFAPEN", "FeReMovi", 8, 0)
   Call CREACAMPO("STOCKS", "ORFAPEN", "NumUsuar", 3, 0)
   Call CREACAMPO("STOCKS", "ORFAPEN", "STATUS", 3, 0)
   Call CREACAMPO("STOCKS", "ORFAPEN", "IDTEXT", 10, 24)
   Call CREACAMPO("STOCKS", "ORFAPEN", "IDSITUACION", 10, 16)
   '
   Call CREACAMPO("STOCKS", "ENCACOT", "NumUsuar", 3, 0)
   Call CREACAMPO("STOCKS", "ENCACOT", "FechMov", 8, 0, 1)
   Call CREACAMPO("STOCKS", "ENCACOT", "FeFtoCot", 10, 12)
   Call CREACAMPO("STOCKS", "ENCACOT", "FESALIDA", 8, 0)
   Call CREACAMPO("STOCKS", "ENCACOT", "COT", 10, 128)
   Call CREACAMPO("STOCKS", "ENCACOT", "TEXTCOT", 10, 12)
   Call CREACAMPO("STOCKS", "ENCACOT", "CODIEMPR", 3, 2, 1)
   Call CREACAMPO("STOCKS", "ENCACOT", "CODINTEGRIDAD", 10, 64)

   
   
   
   If ValidaExisteCampo%("ENCACOT_ID", "ENCACOT") < 1 Then
    SQLX$ = "ALTER TABLE ENCACOT ADD ENCACOT_ID INT IDENTITY"
    FLEXCONN.Execute (SQLX$)
   End If
   '
   Call CREACAMPO("STOCKS", "DETACOT", "FechMov", 8, 0, 1)
   Call CREACAMPO("STOCKS", "DETACOT", "FeFormCot", 10, 12)
   Call CREACAMPO("STOCKS", "DETACOT", "Cod_Inte", 3, 0, 1)
   Call CREACAMPO("STOCKS", "DETACOT", "Cod_Exte", 10, 24)
   Call CREACAMPO("STOCKS", "DETACOT", "Descrip", 10, 128)
   Call CREACAMPO("STOCKS", "DETACOT", "ValoUnit", 7, 0)
   Call CREACAMPO("STOCKS", "DETACOT", "Nume_Clie", 4, 0, 1)
   Call CREACAMPO("STOCKS", "DETACOT", "CantIngre", 7, 0)
   Call CREACAMPO("STOCKS", "DETACOT", "CantSalid", 7, 0)
   Call CREACAMPO("STOCKS", "DETACOT", "NUMPVTA", 3, 0, 1)
   Call CREACAMPO("STOCKS", "DETACOT", "DOPRINUM", 4, 0, 1)
   Call CREACAMPO("STOCKS", "DETACOT", "DoPriLar", 10, 24)
   Call CREACAMPO("STOCKS", "DETACOT", "ID_ENCA", 4, 0, 1)
   Call CREACAMPO("STOCKS", "DETACOT", "CODIEMPR", 3, 2, 1)
   Call CREACAMPO("STOCKS", "DETACOT", "FEREMITO", 8, 0)

    
   Call CREACAMPO("STOCKS", "DETACOT", "DOM_ENTRE", 10, 128)
   Call CREACAMPO("STOCKS", "DETACOT", "LOCALIDAD_ENTR", 10, 64)
   Call CREACAMPO("STOCKS", "DETACOT", "COPOSTAL", 10, 24)
   Call CREACAMPO("STOCKS", "DETACOT", "CUITTRANSP", 10, 24)
   Call CREACAMPO("STOCKS", "DETACOT", "PROVINCIA", 10, 24)
   Call CREACAMPO("STOCKS", "DETACOT", "DOMITRANSP", 10, 64)
   Call CREACAMPO("STOCKS", "DETACOT", "LOCALITRANSP", 10, 64)
   Call CREACAMPO("STOCKS", "DETACOT", "PATENTE", 10, 24)
   Call CREACAMPO("STOCKS", "DETACOT", "DNITRANS", 10, 24) ' AHORA AQUI SE GUARDA EL DNI DEL CLIENTE SI ES CONS.FINAL
   Call CREACAMPO("STOCKS", "DETACOT", "CODUNICO", 10, 24)
   Call CREACAMPO("STOCKS", "DETACOT", "TIPOCAM", 7, 8) '
   Call CREACAMPO("STOCKS", "DETACOT", "MONEVAL", 3, 2)
   
   If ValidaExisteCampo%("DETACOT_ID", "DETACOT") < 1 Then
    SQLX$ = "ALTER TABLE DETACOT ADD DETACOT_ID INT IDENTITY"
    FLEXCONN.Execute (SQLX$)
   End If
   '
   Call CREACAMPO("STOCKS", "FORMANSERIE", "TIP_DATO", 3, 0)
   Call CREACAMPO("STOCKS", "FORMANSERIE", "REFER_DATO", 10, 48)
   Call CREACAMPO("STOCKS", "FORMANSERIE", "CANTICARACT", 3, 0)
   Call CREACAMPO("STOCKS", "FORMANSERIE", "APARTIR", 3, 0)
   Call CREACAMPO("STOCKS", "FORMANSERIE", "EXTREMO", 10, 8)
   Call CREACAMPO("STOCKS", "FORMANSERIE", "COMPLETACON", 10, 8)
   Call CREACAMPO("STOCKS", "FORMANSERIE", "CORRESPONDE_A", 10, 48, 1)
   Call CREACAMPO("STOCKS", "FORMANSERIE", "MASCARA", 10, 24)
   Call CREACAMPO("STOCKS", "FORMANSERIE", "MEDIDAS", 7, 0)
   
   Call CREACAMPO("STOCKS", "FORMANSERIE", "MARBORRA", 3, 0)
   Call CREACAMPO("STOCKS", "FORMANSERIE", "ORDEN", 3, 0)
   '
   '
   '
   Call CREACAMPO("STOCKS", "DESPAIMPO", "IDENLOTE", 10, 24)
   If ValidaExisteCampo%("DESPAIMPO_ID", "DESPAIMPO") < 1 Then
    SQLX$ = "ALTER TABLE DESPAIMPO ADD DESPAIMPO_ID INT IDENTITY"
    FLEXCONN.Execute (SQLX$)
   End If
   '
   Call CREACAMPO("STOCKS", "DESPAIMPO", "IDENLOTE", 10, 24)
   Call CREACAMPO("STOCKS", "DESPAIMPO", "DESPACHO", 10, 128)
   Call CREACAMPO("STOCKS", "DESPAIMPO", "Nume_Ocom", 4, 4, 1)
   Call CREACAMPO("STOCKS", "DESPAIMPO", "FECHA", 8, 0)
   Call CREACAMPO("STOCKS", "DESPAIMPO", "ADUANA", 10, 128)
   Call CREACAMPO("STOCKS", "DESPAIMPO", "NUMEBORE", 6, 4, 1)
   Call CREACAMPO("STOCKS", "DESPAIMPO", "MARBORRA", 3, 0)

   If InStr(EPAC$, "CONTEPLAST") > 0 Then
    Call CREACAMPO("STOCKS", "CONFSCRAP", "DescSector", 10, 32)
    If ValidaExisteCampo%("CONFSCRAP_ID", "CONFSCRAP") < 1 Then
      SQLX$ = "ALTER TABLE CONFSCRAP ADD CONFSCRAP_ID INT IDENTITY"
      FLEXCONN.Execute (SQLX$)
    End If
    Call CREACAMPO("STOCKS", "CONFSCRAP", "CodSector", 3, 1)
    Call CREACAMPO("STOCKS", "CONFSCRAP", "CodCateg", 3, 1)
    Call CREACAMPO("STOCKS", "CONFSCRAP", "FECHA", 8, 0)
    Call CREACAMPO("STOCKS", "CONFSCRAP", "Marborra", 10, 64)
    Call CREACAMPO("STOCKS", "CONFSCRAP", "DescCateg", 10, 64)
    Call CREACAMPO("STOCKS", "CONFSCRAP", "CiArt", 6, 4, 1)
    Call CREACAMPO("STOCKS", "CONFSCRAP", "CodArt", 10, 24)
    Call CREACAMPO("STOCKS", "CONFSCRAP", "DescArt", 10, 64)
    Call CREACAMPO("STOCKS", "CONFSCRAP", "DepIngre", 3, 1)
    Call CREACAMPO("STOCKS", "CONFSCRAP", "CODIEMPR", 3, 2, 1)
    Call CREACAMPO("STOCKS", "CONFSCRAP", "MARBORRA", 3, 0, 1)

    '
    'TABLA DETASCRAP
    Call CREACAMPO("STOCKS", "DETASCRAP", "DescSector", 10, 32)
    If ValidaExisteCampo%("DETASCRAP_ID", "DETASCRAP") < 1 Then
      SQLX$ = "ALTER TABLE DETASCRAP ADD DETASCRAP_ID INT IDENTITY"
      FLEXCONN.Execute (SQLX$)
    End If
    Call CREACAMPO("STOCKS", "DETASCRAP", "Kilos", 7, 8)
    Call CREACAMPO("STOCKS", "DETASCRAP", "CodSector", 3, 1)
    Call CREACAMPO("STOCKS", "DETASCRAP", "CodCateg", 3, 1)
    Call CREACAMPO("STOCKS", "DETASCRAP", "FECHA", 8, 0)
    Call CREACAMPO("STOCKS", "DETASCRAP", "FEREMOV", 8, 0)
    Call CREACAMPO("STOCKS", "DETASCRAP", "DescCateg", 10, 64)
    Call CREACAMPO("STOCKS", "DETASCRAP", "CiArt", 6, 4, 1)
    Call CREACAMPO("STOCKS", "DETASCRAP", "CodArt", 10, 24)
    Call CREACAMPO("STOCKS", "DETASCRAP", "DescArt", 10, 64)
    Call CREACAMPO("STOCKS", "DETASCRAP", "DepIngre", 3, 1)
    Call CREACAMPO("STOCKS", "DETASCRAP", "FEREMOVI", 8, 0)
    Call CREACAMPO("STOCKS", "DETASCRAP", "CODIEMPR", 3, 2, 1)
    Call CREACAMPO("STOCKS", "DETASCRAP", "MARBORRA", 3, 0, 1)

   End If
   If InStr(EPAC$, "ARCON") > 0 Then
    Call CREACAMPO("STOCKS", "CONTROSERIE", "Cod_Inte", 3, 0, 1)
    Call CREACAMPO("STOCKS", "CONTROSERIE", "Cod_Exte", 10, 24)
    If ValidaExisteCampo%("CONTROSERIE_ID", "CONTROSERIE") < 1 Then
      SQLX$ = "ALTER TABLE CONTROSERIE ADD CONTROSERIE_ID INT IDENTITY"
      FLEXCONN.Execute (SQLX$)
    End If
    
    Call CREACAMPO("STOCKS", "CONTROSERIE", "Descrip", 10, 64)
    Call CREACAMPO("STOCKS", "CONTROSERIE", "IdenLote", 10, 12)
    Call CREACAMPO("STOCKS", "CONTROSERIE", "NumeroSerie", 10, 24)
    Call CREACAMPO("STOCKS", "CONTROSERIE", "Observaciones", 10, 48)
    Call CREACAMPO("STOCKS", "CONTROSERIE", "NumeroSerieOrigen", 10, 24)
    Call CREACAMPO("STOCKS", "CONTROSERIE", "CodInte_DestinoConsumo", 3, 0, 1)
    Call CREACAMPO("STOCKS", "CONTROSERIE", "CodExte_DestinoConsumo", 10, 24)
    Call CREACAMPO("STOCKS", "CONTROSERIE", "Status", 3, 0)
    Call CREACAMPO("STOCKS", "CONTROSERIE", "ORDEN", 3, 0, 1)
    Call CREACAMPO("STOCKS", "CONTROSERIE", "MARBORRA", 3, 0)
    Call CREACAMPO("STOCKS", "CONTROSERIE", "IdCodAsignado", 4, 0)
    Call CREACAMPO("STOCKS", "CONTROSERIE", "DepoMovi", 3, 1)
    Call CREACAMPO("STOCKS", "CONTROSERIE", "CantIngre", 7, 0)
    Call CREACAMPO("STOCKS", "CONTROSERIE", "CantSalid", 7, 0)
    Call CREACAMPO("STOCKS", "CONTROSERIE", "FEREMOVI", 8, 0)
    Call CREACAMPO("STOCKS", "CONTROSERIE", "IDSUBOR", 10, 12, 1)
    Call CREACAMPO("STOCKS", "CONTROSERIE", "IDSUBOR2", 10, 16)
    
    Call CREACAMPO("STOCKS", "CONTROSERIE", "UsoBruto", 6, 0)
    Call CREACAMPO("STOCKS", "CONTROSERIE", "EspIdentado", 3, 0)
    
   End If
   
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
   Call CREACAMPO("VENTAS", "DESCOMBI", "Nume_Cli", 4, 0, 1)
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
   Call CREACAMPO("VENTAS", "PEDDETA", "NroRemi", 10, 24)
   If ValidaExisteCampo%("PEDDETA_ID", "PEDDETA") < 1 Then
      SQLX$ = "ALTER TABLE PEDDETA ADD PEDDETA_ID INT IDENTITY"
      FLEXCONN.Execute (SQLX$)
    End If
   Call CREACAMPO("VENTAS", "PEDDETA", "PreList", 6, 4)
   Call CREACAMPO("VENTAS", "PEDDETA", "CanAsig", 6, 4)
   Call CREACAMPO("VENTAS", "PEDDETA", "CanPed", 6, 4)
   Call CREACAMPO("VENTAS", "PEDDETA", "PreUnid", 6, 4)
   Call CREACAMPO("VENTAS", "PEDDETA", "ImpoTot", 7, 8)
   Call CREACAMPO("VENTAS", "PEDDETA", "ImpoAnti", 7, 8)
   Call CREACAMPO("VENTAS", "PEDDETA", "CantEnt", 6, 4)
   Call CREACAMPO("VENTAS", "PEDDETA", "CanSaldo", 6, 4)
   Call CREACAMPO("VENTAS", "PEDDETA", "FeUltEnt", 8, 0)
   Call CREACAMPO("VENTAS", "PEDDETA", "FechPrepa", 8, 0)
   Call CREACAMPO("VENTAS", "PEDDETA", "FechInformado", 8, 0)
   Call CREACAMPO("VENTAS", "PEDDETA", "FechRetiro", 8, 0)
   Call CREACAMPO("VENTAS", "PEDDETA", "FApro_Nope", 8, 0)
   Call CREACAMPO("VENTAS", "PEDDETA", "AcuPrepa", 6, 0)
   Call CREACAMPO("VENTAS", "PEDDETA", "AcuDespa", 6, 0)
   Call CREACAMPO("VENTAS", "PEDDETA", "TPedItem", 10, 16)
   Call CREACAMPO("VENTAS", "PEDDETA", "NroPed", 4, 0, 2)
   Call CREACAMPO("VENTAS", "PEDDETA", "TipoPed", 10, 8, 1)
   Call CREACAMPO("VENTAS", "PEDDETA", "FechEmis", 8, 0, 1)
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
   Call CREACAMPO("VENTAS", "PEDDETA", "CodiExt", 10, 24)
   Call CREACAMPO("VENTAS", "PEDDETA", "Descrip", 10, 384)
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
   Call CREACAMPO("VENTAS", "PEDDETA", "NumeroSerie", 10, 32, 1)
   Call CREACAMPO("VENTAS", "PEDDETA", "IdenLote", 10, 32)
   Call CREACAMPO("VENTAS", "PEDDETA", "LISTA_DE_ITEM", 3, 0)
   Call CREACAMPO("VENTAS", "PEDDETA", "Num_Prov", 4, 0)
   Call CREACAMPO("VENTAS", "PEDDETA", "ObservaItem", 10, 512)
   Call CREACAMPO("VENTAS", "PEDDETA", "SIN_CARGO", 3, 0)
   Call CREACAMPO("VENTAS", "PEDDETA", "Armador", 3, 0)
   Call CREACAMPO("VENTAS", "PEDDETA", "NumeroArmado", 4, 0)
   Call CREACAMPO("VENTAS", "PEDDETA", "AnotaItem", 12, 0)
   Call CREACAMPO("VENTAS", "PEDDETA", "FechUltCambio", 8, 0)
   Call CREACAMPO("VENTAS", "PEDDETA", "MarBorra", 4, 0)
   Call CREACAMPO("VENTAS", "PEDDETA", "NROCOTI", 10, 16)
   Call CREACAMPO("VENTAS", "PEDDETA", "OrdenFabricacion", 10, 18)
   '
   Call CreaIndiceCompuesto("PEDDETA", "NROPED+STATUS")
   Call CreaIndiceCompuesto("PEDDETA", "CODIINT+NUORITEM+NUMEROSERIE")
   Call CreaIndiceCompuesto("PEDDETA", "NROPED+NUMEROSERIE")
   '
   Call CREACAMPO("VENTAS", "PEDENCA", "NroPed", 4, 0, 2)
   Call CREACAMPO("VENTAS", "PEDENCA", "Revision", 3, 2)
   Call CREACAMPO("VENTAS", "PEDENCA", "Referencia", 10, 128)
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
   Call CREACAMPO("VENTAS", "PEDENCA", "FLETE_TEX", 10, 64)
   Call CREACAMPO("VENTAS", "PEDENCA", "FechAnul", 8, 0)
   Call CREACAMPO("VENTAS", "PEDENCA", "FechUltCambio", 8, 0)
   Call CREACAMPO("VENTAS", "PEDENCA", "AtributoComercial", 3, 0)
   '
   'datos de entrega AHP
   Call CREACAMPO("VENTAS", "DatoEntrega", "Deno_Suc", 10, 128)
   If ValidaExisteCampo%("DatoEntrega_ID", "DatoEntrega") < 1 Then
    SQLX$ = "ALTER TABLE DatoEntrega ADD DatoEntrega_ID INT IDENTITY"
    FLEXCONN.Execute (SQLX$)
   End If

   Call CREACAMPO("VENTAS", "DatoEntrega", "NroPed", 4, 0, 2)
   Call CREACAMPO("VENTAS", "DatoEntrega", "NroClie", 4, 0, 1)
   Call CREACAMPO("VENTAS", "DatoEntrega", "Nume_Suc", 3, 0)
   Call CREACAMPO("VENTAS", "DatoEntrega", "Domi_Suc", 10, 128)
   Call CREACAMPO("VENTAS", "DatoEntrega", "CodPostal_Suc", 10, 16)
   Call CREACAMPO("VENTAS", "DatoEntrega", "Loca_Suc", 10, 128)
   Call CREACAMPO("VENTAS", "DatoEntrega", "Tele_Suc", 10, 64)
   Call CREACAMPO("VENTAS", "DatoEntrega", "Codi_Tra", 3, 0, 1)
   Call CREACAMPO("VENTAS", "DatoEntrega", "Hora_Ent_Suc", 10, 48)
   Call CREACAMPO("VENTAS", "DatoEntrega", "InstruccionSuc", 10, 512)
   Call CREACAMPO("VENTAS", "DatoEntrega", "ProvEntre_Cli", 10, 4)
   
   Call CREACAMPO("VENTAS", "DatoEntrega", "Deno_Tra", 10, 128)
'   Call CREACAMPO("VENTAS", "TRANSPORTES", "Raso_Tra", 10, 128)
   Call CREACAMPO("VENTAS", "DatoEntrega", "Domi_Tra", 10, 128)
   Call CREACAMPO("VENTAS", "DatoEntrega", "Loca_Tra", 10, 128)
   Call CREACAMPO("VENTAS", "DatoEntrega", "CodPost_Tra", 10, 16)
   Call CREACAMPO("VENTAS", "DatoEntrega", "Prov_Tra", 10, 4)
   Call CREACAMPO("VENTAS", "DatoEntrega", "Tele_Tra", 10, 64)
   Call CREACAMPO("VENTAS", "DatoEntrega", "HoraTransporte", 10, 64)
   Call CREACAMPO("VENTAS", "DatoEntrega", "InstrucTransporte", 10, 512)
   Call CREACAMPO("VENTAS", "DatoEntrega", "MARBORRA", 3, 0)
   Call CREACAMPO("VENTAS", "DatoEntrega", "EntEnDomi", 3, 0)
   Call CREACAMPO("VENTAS", "DatoEntrega", "CantBultos", 3, 0)
   Call CREACAMPO("VENTAS", "DatoEntrega", "PagoOrigen", 3, 0)
   Call CREACAMPO("VENTAS", "DatoEntrega", "PagoDestino", 3, 0)
  '_______________________
   
   Call CREACAMPO("VENTAS", "DetaBultos", "NumeroBulto", 3, 0)
   If ValidaExisteCampo%("DetaBultos_ID", "DetaBultos") < 1 Then
    SQLX$ = "ALTER TABLE DetaBultos ADD DetaBultos_ID INT IDENTITY"
    FLEXCONN.Execute (SQLX$)
   End If
   Call CREACAMPO("VENTAS", "DetaBultos", "pesokg", 7, 0)
   Call CREACAMPO("VENTAS", "DetaBultos", "NroPed", 6, 0, 1)
   Call CREACAMPO("VENTAS", "DetaBultos", "NroClie", 4, 0, 1)
   Call CREACAMPO("VENTAS", "DetaBultos", "MARBORRA", 3, 0)
   Call CREACAMPO("VENTAS", "DetaBultos", "STATUS", 3, 0) ' pendiente de entrega = 0,entregado = 1,devuelto = -1
   Call CREACAMPO("VENTAS", "DetaBultos", "NroDespa", 6, 0, 1)
   Call CREACAMPO("VENTAS", "DetaBultos", "FeReMovi", 8, 0)
   Call CREACAMPO("VENTAS", "DetaBultos", "NroGuia", 10, 48)
  '_______________________
  
   Call CREACAMPO("VENTAS", "PRESUVEN", "Cod_Rubro", 10, 8)
   If ValidaExisteCampo%("PRESUVEN_ID", "PRESUVEN") < 1 Then
    SQLX$ = "ALTER TABLE PRESUVEN ADD PRESUVEN_ID INT IDENTITY"
    FLEXCONN.Execute (SQLX$)
   End If

   Call CREACAMPO("VENTAS", "PRESUVEN", "NroPres", 6, 4)
   Call CREACAMPO("VENTAS", "PRESUVEN", "Cod_Rubro", 10, 8)
   Call CREACAMPO("VENTAS", "PRESUVEN", "NumeCRubro", 4, 4)
   Call CREACAMPO("VENTAS", "PRESUVEN", "Cod_Anio", 10, 4)
   Call CREACAMPO("VENTAS", "PRESUVEN", "FechEmis", 8, 0, 1)
   Call CREACAMPO("VENTAS", "PRESUVEN", "NroClie", 4, 0, 1)
   Call CREACAMPO("VENTAS", "PRESUVEN", "Codi_Cli", 10, 12)
   Call CREACAMPO("VENTAS", "PRESUVEN", "RasoClie", 10, 64)
   Call CREACAMPO("VENTAS", "PRESUVEN", "Referen", 10, 64)
   Call CREACAMPO("VENTAS", "PRESUVEN", "AteSen", 10, 48)
   Call CREACAMPO("VENTAS", "PRESUVEN", "LisPre", 3, 0)
   Call CREACAMPO("VENTAS", "PRESUVEN", "ConPag", 3, 0)
   Call CREACAMPO("VENTAS", "PRESUVEN", "Vended", 3, 0)
   Call CREACAMPO("VENTAS", "PRESUVEN", "MonedEmi", 3, 2)
   Call CREACAMPO("VENTAS", "PRESUVEN", "NuOrdIt", 3, 0)
   Call CREACAMPO("VENTAS", "PRESUVEN", "CodiInt", 3, 0)
   Call CREACAMPO("VENTAS", "PRESUVEN", "CodiExt", 10, 24)
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
   Call CREACAMPO("VENTAS", "PRESUVEN", "Mail_Cli", 10, 64)
   Call CREACAMPO("VENTAS", "PRESUVEN", "COSTOFOB", 6, 4)
   Call CREACAMPO("VENTAS", "PRESUVEN", "PresComReal", 7, 0)
   
   
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
   
   Call CREACAMPO("VENTAS", "PREVENTA_RESP", "Cod_Lista", 3, 2, 1)
   Call CREACAMPO("VENTAS", "PREVENTA_RESP", "Cod_Pieza", 3, 2, 1)
   Call CREACAMPO("VENTAS", "PREVENTA_RESP", "Num_Orden", 3, 2)
   Call CREACAMPO("VENTAS", "PREVENTA_RESP", "Desc_Pieza", 10, 64)
   Call CREACAMPO("VENTAS", "PREVENTA_RESP", "Denom_Pieza", 10, 64)
   Call CREACAMPO("VENTAS", "PREVENTA_RESP", "Carac_Control", 10, 64)
   Call CREACAMPO("VENTAS", "PREVENTA_RESP", "Num_Original", 10, 64)
   Call CREACAMPO("VENTAS", "PREVENTA_RESP", "Precio_Unitario", 7, 8)
   Call CREACAMPO("VENTAS", "PREVENTA_RESP", "Cant_Limite", 6, 8)
   Call CREACAMPO("VENTAS", "PREVENTA_RESP", "Porc_Limite", 7, 32)
   Call CREACAMPO("VENTAS", "PREVENTA_RESP", "MarBorra", 3, 0)
   Call CREACAMPO("VENTAS", "PREVENTA_RESP", "FECHMOD", 8, 0)
   Call CREACAMPO("VENTAS", "PREVENTA_RESP", "IDMODIFIC", 7, 0, 1)
   Call CREACAMPO("VENTAS", "PREVENTA_RESP", "HORA", 10, 24)
   
   If ValidaExisteCampo%("PREVENTA_RESP_ID", "PREVENTA_RESP") < 1 Then
    SQLX$ = "ALTER TABLE PREVENTA_RESP ADD PREVENTA_RESP_ID INT IDENTITY"
    FLEXCONN.Execute (SQLX$)
   End If
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
   
   Call CREAR_TABLA_HojaTrabajoDifCambio
   '
End Sub
Sub CREARCAMPOS_FACTURAYGASTOS()
       Call ABRECONEXION
       Call CREACAMPO("", "FACTURAYGASTOS", "Nume_Cli", 6, 0, 1)
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
       Call CREACAMPO("", "FACTURAYGASTOS", "Nume_Cli_Asoc", 6, 0, 1)
        Call CREACAMPO("", "FACTURAYGASTOS", "AJUSTE_VAL_COMPRA", 7, 0)
End Sub

Sub CREARCAMPOS_DETFACOMSQL()
       Call ABRECONEXION
       Call CREACAMPO("", "DETFACOMSQL", "Nume_Cli", 4, 0, 1)
       Call CREACAMPO("", "DETFACOMSQL", "COD_INTE", 3, 0, 1)
       Call CREACAMPO("", "DETFACOMSQL", "FACTURAPROV", 10, 24, 1)
       Call CREACAMPO("", "DETFACOMSQL", "FECHEMISI", 8, 0)
       Call CREACAMPO("", "DETFACOMSQL", "FECHCONTA", 8, 0)
       Call CREACAMPO("", "DETFACOMSQL", "NumeBoRe", 4, 0, 1)
       Call CREACAMPO("", "DETFACOMSQL", "Nume_Ocom", 4, 0)
       Call CREACAMPO("", "DETFACOMSQL", "IDENLOTE", 10, 16)
       Call CREACAMPO("", "DETFACOMSQL", "LOTECOL_PROV", 10, 16)
       Call CREACAMPO("", "DETFACOMSQL", "CantiSegunRemito", 7, 0)
       Call CREACAMPO("", "DETFACOMSQL", "CantiAprobada", 7, 0)
       Call CREACAMPO("", "DETFACOMSQL", "CantPendFact", 7, 0)
       Call CREACAMPO("", "DETFACOMSQL", "CantFacturada", 7, 0)
       Call CREACAMPO("", "DETFACOMSQL", "PreunitS_OC", 7, 8)
       Call CREACAMPO("", "DETFACOMSQL", "PreunitC_Bonif", 7, 8)
       Call CREACAMPO("", "DETFACOMSQL", "PreunitS_Factura", 7, 8)
       Call CREACAMPO("", "DETFACOMSQL", "IMPORTE", 7, 8)
       Call CREACAMPO("", "DETFACOMSQL", "DIFERENCIA", 7, 8)
       Call CREACAMPO("", "DETFACOMSQL", "CodInCue", 3, 0)
       Call CREACAMPO("", "DETFACOMSQL", "MARBORRA", 3, 0)
       Call CREACAMPO("", "DETFACOMSQL", "NordenBrItem", 3, 0)
       Call CREACAMPO("", "DETFACOMSQL", "Nroped", 4, 0, 1) 'el nro de pedido que se genera hacia ahp desde facturacion de acreedores

End Sub

Sub GENERA_TABLA_COMPRAS_DE_PRESUPUESTO()
       Call ABRECONEXION
       Call CREACAMPO("", "PRESUCOMPRA", "Nume_Cli", 3, 0, 1)
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
       Call CREACAMPO("", "PRESUCOMPRA", "REFERENCIA", 10, 64)
       Call CREACAMPO("", "PRESUCOMPRA", "Nume_Ocom", 4, 4, 1)
       Call CREACAMPO("", "PRESUCOMPRA", "CODICONJ", 4, 0, 1)
       Call CREACAMPO("", "PRESUCOMPRA", "CODXCONJ", 10, 64)
       Call CREACAMPO("", "PRESUCOMPRA", "DescripCodConj", 10, 128)
       
End Sub

Sub CREAR_TABLA_OC_DE_SUCURSAL()
    Call ABRECONEXION
    Call CREACAMPO("", "OC_SUCURSAL", "SUCURSAL", 4, 2, 1)
    Call CREACAMPO("", "OC_SUCURSAL", "FECHEMISI", 8, 0)
    Call CREACAMPO("", "OC_SUCURSAL", "Codint", 3, 0, 1)
    Call CREACAMPO("", "OC_SUCURSAL", "NROPED_SUC", 4, 0, 1)
    Call CREACAMPO("", "OC_SUCURSAL", "NCLIEN_PEDIDO", 4, 0, 1)
    Call CREACAMPO("", "OC_SUCURSAL", "RAZ_CTE_SUC", 10, 32)
    Call CREACAMPO("", "OC_SUCURSAL", "USUARIO", 3, 0)
    Call CREACAMPO("", "OC_SUCURSAL", "O_COMPRA_SUC", 4, 4, 1)
    Call CREACAMPO("", "OC_SUCURSAL", "CANTIDAD", 6, 4)
    Call CREACAMPO("", "OC_SUCURSAL", "NPED_GENERADO_EN_AVELLANEDA", 4, 0, 1)
    Call CREACAMPO("", "OC_SUCURSAL", "VERSION_COMPRA_SUC", 3, 0)
    Call CREACAMPO("", "OC_SUCURSAL", "MARBORRA", 3, 0)
    Call CREACAMPO("", "OC_SUCURSAL", "STATUS", 3, 0)
    Call CREACAMPO("", "OC_SUCURSAL", "FECHA_SOLICI_COMPRA", 8, 0)
    Call CREACAMPO("", "OC_SUCURSAL", "OComCliente", 10, 16)
    Call CREACAMPO("", "OC_SUCURSAL", "COD_EXTE", 10, 24)
    
End Sub

Sub CREAR_TABLA_MENSAJES_DE_SUCURSAL()
    Call ABRECONEXION
    Call CREACAMPO("", "MENSA_SUCURSAL", "O_COMPRA_SUC", 4, 4, 1)
    Call CREACAMPO("", "MENSA_SUCURSAL", "SUCURSAL", 4, 2, 1)
    Call CREACAMPO("", "MENSA_SUCURSAL", "USUARIO", 3, 0)
    Call CREACAMPO("", "MENSA_SUCURSAL", "STATUS", 3, 0)
    Call CREACAMPO("", "MENSA_SUCURSAL", "MENSAJE", 10, 1024)
End Sub
Sub CREAR_TABLA_HojaTrabajoDifCambio()
    Call ABRECONEXION
    Call CREACAMPO("", "HojaTrabajoDifCambio", "NCLIEN", 4, 0, 1)
    Call CREACAMPO("", "HojaTrabajoDifCambio", "RECIBO", 10, 24, 1)
    Call CREACAMPO("", "HojaTrabajoDifCambio", "FechaRecibo", 8, 0)
    Call CREACAMPO("", "HojaTrabajoDifCambio", "MedioPago", 10, 64)
    Call CREACAMPO("", "HojaTrabajoDifCambio", "FechaAcreditPago", 8, 0)
    Call CREACAMPO("", "HojaTrabajoDifCambio", "ImpoMedioPagoPesos", 7, 0)
    Call CREACAMPO("", "HojaTrabajoDifCambio", "TCambioFechaAcredit", 7, 0)
    '
    Call CREACAMPO("", "HojaTrabajoDifCambio", "DolaresCobrados", 7, 0)
    Call CREACAMPO("", "HojaTrabajoDifCambio", "FACTURA_APL", 10, 24)
    Call CREACAMPO("", "HojaTrabajoDifCambio", "FECHA_FACTURA_APL", 8, 0)
    Call CREACAMPO("", "HojaTrabajoDifCambio", "ImpoDolSobreFactura", 7, 0)
    Call CREACAMPO("", "HojaTrabajoDifCambio", "STATUS", 3, 0)
    Call CREACAMPO("", "HojaTrabajoDifCambio", "ORDEN", 3, 0)
    Call CREACAMPO("", "HojaTrabajoDifCambio", "MARBORRA", 3, 0)
    Call CREACAMPO("", "HojaTrabajoDifCambio", "NOTADEBITOGENERADA", 10, 24)
    Call CREACAMPO("", "HojaTrabajoDifCambio", "NumeComp", 4, 0, 1)
    Call CREACAMPO("", "HojaTrabajoDifCambio", "DifCambioEnPesos", 7, 0)
    Call CREACAMPO("", "HojaTrabajoDifCambio", "ImpoPesosAplicados", 7, 0)
    Call CREACAMPO("", "HojaTrabajoDifCambio", "DifCambioPesosBruto", 7, 0)
    '
    'IMPO_ORIG_FACTU_USS = IMPORTE ORIGINAL DE LA FACTURA EN DOLARES AL MOMENTO DE LA FACTURA
End Sub
Sub CREAR_TABLA_HISTOCAMBIO()
    Call ABRECONEXION
    Call CREACAMPO("", "HISTOCAMBIO", "ID", 3, 0, 1)
    Call CREACAMPO("", "HISTOCAMBIO", "DESCRIPCION", 10, 24)
    Call CREACAMPO("", "HISTOCAMBIO", "FECHA", 8, 0, 1)
    Call CREACAMPO("", "HISTOCAMBIO", "COTIZA_DOLAR", 7, 0)
End Sub
Sub CREAR_TABLA_ANOTADIFCAMBIO()
    Call ABRECONEXION
    Call CREACAMPO("", "ANOTADIFCAMBIO", "RECIBO", 10, 24, 1)
    Call CREACAMPO("", "ANOTADIFCAMBIO", "HISTORIAL", 12, 0)
    Call CREACAMPO("", "ANOTADIFCAMBIO", "NUMECOMP", 4, 0, 1)
    Call CREACAMPO("", "ANOTADIFCAMBIO", "NUCLIEN", 4, 2, 1)
End Sub

Sub CREAR_TABLA_DetaOVenta()
    Call ABRECONEXION
    Call CREACAMPO("", "DetaOVenta", "Nroped", 4, 0, 1)
    Call CREACAMPO("", "DetaOVenta", "Nro_Ocom", 4, 0, 1)
    Call CREACAMPO("", "DetaOVenta", "Codint", 4, 0, 1)
    Call CREACAMPO("", "DetaOVenta", "CodiExt", 10, 24)
    Call CREACAMPO("", "DetaOVenta", "Descrip", 10, 64)
    Call CREACAMPO("", "DetaOVenta", "CanPed", 6, 4)
    Call CREACAMPO("", "DetaOVenta", "CantEnt", 6, 4)
    Call CREACAMPO("", "DetaOVenta", "FechaModif", 8, 0)
    Call CREACAMPO("", "DetaOVenta", "Accion", 10, 24)
    Call CREACAMPO("", "DetaOVenta", "NCtePedido", 4, 0, 1)
    Call CREACAMPO("", "DetaOVenta", "MARBORRA", 3, 0)

End Sub
Sub CREAR_TABLA_RESERPED()
   Call ABRECONEXION
   Call CREACAMPO("", "RESERPED", "Cod_Inte", 3, 2, 1)
   Call CREACAMPO("", "RESERPED", "Cod_Exte", 10, 16)
   Call CREACAMPO("", "RESERPED", "Descrip", 10, 64)
   Call CREACAMPO("", "RESERPED", "NuOrItem", 3, 0)
   Call CREACAMPO("", "RESERPED", "FechRes", 8, 0)
   Call CREACAMPO("", "RESERPED", "FechCierre", 8, 0)
   Call CREACAMPO("", "RESERPED", "FechSecado", 8, 0)
   Call CREACAMPO("", "RESERPED", "CantRes", 7, 8)
   Call CREACAMPO("", "RESERPED", "CantCons", 7, 8)
   Call CREACAMPO("", "RESERPED", "Saldo_Entre", 7, 8)
   Call CREACAMPO("", "RESERPED", "IdenLote", 10, 16, 1)
   Call CREACAMPO("", "RESERPED", "NroPed", 4, 0, 1)
   Call CREACAMPO("", "RESERPED", "NroClie", 4, 0, 1)
   Call CREACAMPO("", "RESERPED", "STAT_PED", 3, 0)
   Call CREACAMPO("", "RESERPED", "Marborra", 3, 0)
   
End Sub
   

Function EXISTE_TABLA%(NombreTabla$)
   '0 NO EXIST
   '1 EXISTE
   
   Call ABRECONEXION
   SQLX$ = "SELECT COUNT(*) FROM INFORMATION_SCHEMA.TABLES WITH(NOLOCK) "
   SQLX$ = SQLX$ + " WHERE TABLE_NAME = '" & NombreTabla$ & "'"
   Set rs = READSET(SQLX$)
   EXISTE_TABLA% = rs.Fields(0)
End Function

Sub CREATABLA_EQUIPOAHP()

   Call CREACAMPO("", "EQUIPOAHP", "MARCA", 10, 64)
   Call CREACAMPO("", "EQUIPOAHP", "ID", 4, 0, 1)
   Call CREACAMPO("", "EQUIPOAHP", "T_EQUIPO", 10, 641)
   Call CREACAMPO("", "EQUIPOAHP", "MODELO", 10, 64)
   Call CREACAMPO("", "EQUIPOAHP", "MarBorra", 4, 1)
   Call CREACAMPO("", "EQUIPOAHP", "STATUS", 3, 0)
   '
   Call CREACAMPO("", "EQUICLI", "IDEQUIPO", 4, 0, 1)
   Call CREACAMPO("", "EQUICLI", "IDCOMBI", 4, 0, 1)
   Call CREACAMPO("", "EQUICLI", "INTERNO", 10, 64)
   Call CREACAMPO("", "EQUICLI", "MOTOR", 10, 64)
   Call CREACAMPO("", "EQUICLI", "SERIEMOTOR", 10, 64)
   Call CREACAMPO("", "EQUICLI", "NUCLIEN", 4, 0, 1)
   Call CREACAMPO("", "EQUICLI", "MARBORRA", 3, 0)
   Call CREACAMPO("", "EQUICLI", "STATUS", 3, 0)
   Call CREACAMPO("", "EQUICLI", "NUMEROSERIE", 10, 64)

End Sub
Sub CREATABLA_CONTMINI()
   Call CREACAMPO("", "CONTMINI", "ID", 4, 0, 1)
   Call CREACAMPO("", "CONTMINI", "COD_INTE", 3, 0, 1)
   Call CREACAMPO("", "CONTMINI", "COD_Exte", 10, 24)
   Call CREACAMPO("", "CONTMINI", "MiniSuger", 7, 0)
   Call CREACAMPO("", "CONTMINI", "StoFeProceso", 7, 0)
   Call CREACAMPO("", "CONTMINI", "FeProceso", 8, 8)
   Call CREACAMPO("", "CONTMINI", "ALERTA", 3, 0, 1)
   
   Call CREACAMPO("", "CONTMINI", "COMP_PEND", 7, 0)
   Call CREACAMPO("", "CONTMINI", "RECE_PEND", 7, 0)
 
End Sub
Sub CREATABLA_DETAMASU()
   
   Call CREACAMPO("", "DETAMASU", "ID", 4, 0, 1)
   Call CREACAMPO("", "DETAMASU", "CODIMOVI", 10, 2, 1)
   Call CREACAMPO("", "DETAMASU", "COD_INTE", 3, 0, 1)
   Call CREACAMPO("", "DETAMASU", "COD_Exte", 10, 64)
   Call CREACAMPO("", "DETAMASU", "TALON", 7, 0, 1)
   Call CREACAMPO("", "DETAMASU", "NReparación", 7, 0, 1)
   Call CREACAMPO("", "DETAMASU", "CANTINGRE", 7, 0)
   Call CREACAMPO("", "DETAMASU", "CANTSALID", 7, 0)
   Call CREACAMPO("", "DETAMASU", "FECHMOV", 8, 0)
   Call CREACAMPO("", "DETAMASU", "FEREMOVI", 8, 0)
   Call CREACAMPO("", "DETAMASU", "NUSER", 3, 0)
   Call CREACAMPO("", "DETAMASU", "Nume_Clie", 4, 0, 1)
   Call CREACAMPO("", "DETAMASU", "NumRemito", 10, 24)
   Call CREACAMPO("", "DETAMASU", "OBSERVACIONES", 10, 64)
   Call CREACAMPO("", "DETAMASU", "REFERENCIA", 10, 64)
   
End Sub
Sub CREATABLA_PLAFABR()

   Call CREACAMPO("", "PLAFABR", "ID", 4, 0, 1)
   Call CREACAMPO("", "PLAFABR", "FECALCULO", 8, 0)
   Call CREACAMPO("", "PLAFABR", "COD_INTE", 4, 0, 1)
   Call CREACAMPO("", "PLAFABR", "DESCRIPCION", 10, 128)
   Call CREACAMPO("", "PLAFABR", "FECHNECE", 8, 0)
   Call CREACAMPO("", "PLAFABR", "PERIO_PROGRAM", 10, 24)
   Call CREACAMPO("", "PLAFABR", "CANT_PROG", 7, 0)
   Call CREACAMPO("", "PLAFABR", "OF_PEND", 7, 0)
   Call CREACAMPO("", "PLAFABR", "OF_A_EMITIR", 7, 0)
   Call CREACAMPO("", "PLAFABR", "NROPROV", 3, 0)
   Call CREACAMPO("", "PLAFABR", "RAZOPROV", 10, 64)
   Call CREACAMPO("", "PLAFABR", "Pre_Unit", 7, 8)
   Call CREACAMPO("", "PLAFABR", "IMP_TOTAL", 7, 8)
   Call CREACAMPO("", "PLAFABR", "CODIMOVI", 10, 8, 1)
   Call CREACAMPO("", "PLAFABR", "NUSER", 3, 0)
End Sub
Sub CREATABLA_OCOEMIHIS()

   Call CREACAMPO("", "OCOEMIHIS", "ID", 4, 0, 1)
   Call CREACAMPO("", "OCOEMIHIS", "FECALCULO", 8, 0)
   Call CREACAMPO("", "OCOEMIHIS", "Cod_Inte", 3, 2, 1)
   Call CREACAMPO("", "OCOEMIHIS", "Cantidad", 6, 4)
   Call CREACAMPO("", "OCOEMIHIS", "Fentre_Sol", 8, 0)
   Call CREACAMPO("", "OCOEMIHIS", "Npro_Ocom", 3, 2, 1)
   Call CREACAMPO("", "OCOEMIHIS", "Uni_Mas", 10, 4)
   Call CREACAMPO("", "OCOEMIHIS", "RaSo_Prov", 10, 48)
   Call CREACAMPO("", "OCOEMIHIS", "Cant_Ocom", 6, 4)
   Call CREACAMPO("", "OCOEMIHIS", "FechRecep", 8, 0)
   Call CREACAMPO("", "OCOEMIHIS", "Nume_Ocom", 4, 2, 1)
   Call CREACAMPO("", "OCOEMIHIS", "Periodo", 10, 24)
   Call CREACAMPO("", "OCOEMIHIS", "Cosuni", 6, 0)
   Call CREACAMPO("", "OCOEMIHIS", "StokAct", 6, 0)
   Call CREACAMPO("", "OCOEMIHIS", "Stomin", 6, 0)
   Call CREACAMPO("", "OCOEMIHIS", "Lorepos", 6, 0)
   Call CREACAMPO("", "OCOEMIHIS", "CODIMOVI", 10, 8, 1)
   Call CREACAMPO("", "OCOEMIHIS", "NUSER", 3, 0)
   Call CREACAMPO("", "OCOEMIHIS", "DESCRIPCION", 10, 128)
End Sub

Sub CREATABLA_COBERFA()

   Call CREACAMPO("", "COBERFA", "ID", 4, 0, 1)
   Call CREACAMPO("", "COBERFA", "FECALCULO", 8, 0)
   Call CREACAMPO("", "COBERFA", "ORDENANOMES", 3, 0, 1)
   Call CREACAMPO("", "COBERFA", "Cod_Inte", 3, 2, 1)
   Call CREACAMPO("", "COBERFA", "Codigo", 10, 24)
   Call CREACAMPO("", "COBERFA", "Descripcion", 10, 64)
   
   Call CREACAMPO("", "COBERFA", "MESANO", 10, 24)
   Call CREACAMPO("", "COBERFA", "DETREFE", 10, 48)
   Call CREACAMPO("", "COBERFA", "TIPOMOVI", 3, 0, 1)
   Call CREACAMPO("", "COBERFA", "CantiDisp", 6, 4)
   Call CREACAMPO("", "COBERFA", "CantiComprom", 6, 4)
   Call CREACAMPO("", "COBERFA", "SaldoDisp", 6, 4)
   Call CREACAMPO("", "COBERFA", "NUSER", 3, 0)
   Call CREACAMPO("", "COBERFA", "ORDENPRES", 6, 4)
   Call CREACAMPO("", "COBERFA", "IDYFECAL", 10, 32)
   Call CREACAMPO("", "COBERFA", "FecNecesi", 8, 0)

End Sub

Sub CREATABLAS_SQLCTRL()
   Static CTX%
   If CTX% < 1 Then
      CTX% = 1
      Call CREACAMPO("SQLCTRL", "SQL_QUERIES", "Cod_Query", 10, 16)
      Call CREACAMPO("SQLCTRL", "SQL_QUERIES", "Des_Query", 10, 254)
      Call CREACAMPO("SQLCTRL", "SQL_QUERIES", "Tip_Query", 3, 2)
      Call CREACAMPO("SQLCTRL", "SQL_QUERIES", "Par_Query", 12, 0)
      Call CREACAMPO("SQLCTRL", "SQL_QUERIES", "SQL_Query", 12, 0)
      '
      Call CREACAMPO("SQLCTRL", "SQL_GRAFICS", "Cod_Query", 10, 16)
      Call CREACAMPO("SQLCTRL", "SQL_GRAFICS", "Cod_Graph", 10, 16)
      Call CREACAMPO("SQLCTRL", "SQL_GRAFICS", "Des_Graph", 10, 254)
      Call CREACAMPO("SQLCTRL", "SQL_GRAFICS", "Par_Graph", 12, 0)
      '
      Call CREACAMPO("SQLCTRL", "SQL_PROCEDS", "Cod_PROCEX", 10, 16)
      Call CREACAMPO("SQLCTRL", "SQL_PROCEDS", "Des_PROCEX", 10, 254)
      Call CREACAMPO("SQLCTRL", "SQL_PROCEDS", "Tip_PROCEX", 3, 2)
      Call CREACAMPO("SQLCTRL", "SQL_PROCEDS", "Par_PROCEX", 12, 0)
      Call CREACAMPO("SQLCTRL", "SQL_PROCEDS", "SQL_PROCEX", 12, 0)
      '
      Call CREACAMPO("SQLCTRL", "USER_PARMS", "Usuario", 3, 2, 1)
      Call CREACAMPO("SQLCTRL", "USER_PARMS", "Proceso", 10, 128, 1)
      Call CREACAMPO("SQLCTRL", "USER_PARMS", "Clave", 10, 128, 1)
      Call CREACAMPO("SQLCTRL", "USER_PARMS", "ValoCon", 12, 0)
      '
      Call CREACAMPO("SQLCTRL", "USER_MENUS", "Usuario", 3, 2, 1)
      Call CREACAMPO("SQLCTRL", "USER_MENUS", "CodigoPadre", 10, 32, 1)
      Call CREACAMPO("SQLCTRL", "USER_MENUS", "CodigoHijo", 10, 32, 1)
      Call CREACAMPO("SQLCTRL", "USER_MENUS", "Descripcion", 10, 64)
      Call CREACAMPO("SQLCTRL", "USER_MENUS", "CodigoImagen", 10, 32)
      Call CREACAMPO("SQLCTRL", "USER_MENUS", "NumeOrden", 3, 2)
      Call CREACAMPO("SQLCTRL", "USER_MENUS", "CantiAcce", 5, 4)
      '
      Call CREACAMPO("SQLCTRL", "SQL_MENUS", "Form_Name", 10, 32)
      Call CREACAMPO("SQLCTRL", "SQL_MENUS", "Cod_Query", 10, 32)
      Call CREACAMPO("SQLCTRL", "SQL_MENUS", "Num_Orden", 3, 2)
      '
      Call CREACAMPO("SQLCTRL", "SQL_ACCDIR", "Ejecutable", 10, 32)
      Call CREACAMPO("SQLCTRL", "SQL_ACCDIR", "LEYENDA", 10, 128)
      Call CREACAMPO("SQLCTRL", "SQL_ACCDIR", "CODIGO", 10, 32)
      Call CREACAMPO("SQLCTRL", "SQL_ACCDIR", "Status", 3, 0)
      
      '
      Call CREACAMPO("SQLCTRL", "INDI_APLI", "COD_APLI", 10, 32)
      Call CREACAMPO("SQLCTRL", "INDI_APLI", "DES_APLI", 10, 128)
      '
      Call CREACAMPO("SQLCTRL", "DER_ACCE", "NUM_USER", 3, 2, 1)
      Call CREACAMPO("SQLCTRL", "DER_ACCE", "COD_APLI", 10, 32, 1)
      Call CREACAMPO("SQLCTRL", "DER_ACCE", "DER_READ", 3, 2)
      Call CREACAMPO("SQLCTRL", "DER_ACCE", "DER_WRITE", 3, 2)
      '
   End If
End Sub
Sub CREATABLA_DEMORA()

'   Call CREACAMPO("", "COBERFA", "ID", 4, 0, 1)
'   Call CREACAMPO("", "COBERFA", "FECALCULO", 8, 0)
'   Call CREACAMPO("", "COBERFA", "ORDENANOMES", 3, 0, 1)
'   Call CREACAMPO("", "COBERFA", "Cod_Inte", 3, 2, 1)
'   Call CREACAMPO("", "COBERFA", "Codigo", 10, 24)
'   Call CREACAMPO("", "COBERFA", "Descripcion", 10, 64)
'
'   Call CREACAMPO("", "COBERFA", "MESANO", 10, 24)
'   Call CREACAMPO("", "COBERFA", "DETREFE", 10, 48)
'   Call CREACAMPO("", "COBERFA", "TIPOMOVI", 3, 0, 1)
'   Call CREACAMPO("", "COBERFA", "CantiDisp", 6, 4)
'   Call CREACAMPO("", "COBERFA", "CantiComprom", 6, 4)
'   Call CREACAMPO("", "COBERFA", "SaldoDisp", 6, 4)
'   Call CREACAMPO("", "COBERFA", "NUSER", 3, 0)
'   Call CREACAMPO("", "COBERFA", "ORDENPRES", 6, 4)
'   Call CREACAMPO("", "COBERFA", "IDYFECAL", 10, 32)
'   Call CREACAMPO("", "COBERFA", "FecNecesi", 8, 0)

End Sub

Sub GENERATABLA_ESPECIFIC()
   Call CREACAMPO("", "ESPECIFIC", "ID", 4, 0, 1)
   Call CREACAMPO("", "ESPECIFIC", "COD_INTE", 3, 2, 1)
   Call CREACAMPO("", "ESPECIFIC", "Codigo", 10, 24)
   Call CREACAMPO("", "ESPECIFIC", "Descripcion", 10, 64)
   Call CREACAMPO("", "ESPECIFIC", "ITEM", 10, 128)
   Call CREACAMPO("", "ESPECIFIC", "CODIDATO", 10, 128)
   Call CREACAMPO("", "ESPECIFIC", "PREPARO", 10, 24)
   Call CREACAMPO("", "ESPECIFIC", "REVISO", 10, 24)
   Call CREACAMPO("", "ESPECIFIC", "APROBO", 10, 24)
   Call CREACAMPO("", "ESPECIFIC", "FEPREPARO", 8, 0)
   Call CREACAMPO("", "ESPECIFIC", "FEREVISO", 8, 0)
   Call CREACAMPO("", "ESPECIFIC", "FEAPROBO", 8, 0)
   
   Call CREACAMPO("", "ESPECIFIC", "COMENTARIO", 10, 64)
   Call CREACAMPO("", "ESPECIFIC", "CODCLIENTE", 10, 64)
   Call CREACAMPO("", "ESPECIFIC", "OBSERVACION", 12, 0) 'OBSERVACION SE DEBE GUARDAR COMO OTRO ITEM
   Call CREACAMPO("", "ESPECIFIC", "MARBORRA", 3, 0, 1)
   Call CREACAMPO("", "ESPECIFIC", "INDIOBJ", 4, 0, 1)
   Call CREACAMPO("", "ESPECIFIC", "UsoBruto", 6, 0)
   Call CREACAMPO("", "ESPECIFIC", "CORTE", 7, 8)
   Call CREACAMPO("", "ESPECIFIC", "Codint_Comp", 3, 2, 1)
   Call CREACAMPO("", "ESPECIFIC", "FUELLE", 10, 12)
   If ValidaExisteCampo%("ESPECIFIC_ID", "ESPECIFIC") < 1 Then
    SQLX$ = "ALTER TABLE ESPECIFIC ADD ESPECIFIC_ID INT IDENTITY"
    FLEXCONN.Execute (SQLX$)
   End If

   
   
End Sub
Sub CREATABLA_ULTRECIBO()
   Call CREACAMPO("", "ULTRECIBO", "CODIEMPR", 3, 2, 1)
   Call CREACAMPO("", "ULTRECIBO", "NumeComp", 4, 0, 1)
   Call CREACAMPO("", "ULTRECIBO", "TIPOMOVIM", 10, 8, 1)
   Call CREACAMPO("", "ULTRECIBO", "FECHEMISI", 8, 0)
   Call CREACAMPO("", "ULTRECIBO", "NUMETALO", 3, 0, 1)
   Call CREACAMPO("", "ULTRECIBO", "Status", 3, 0)
End Sub

Sub CREATABLA_ENCAPRF()
   Call CREACAMPO("", "ENCAPRFUN", "TITURECORD", 10, 128)
   If ValidaExisteCampo%("ENCAPRFUN_ID", "ENCAPRFUN") < 1 Then
    SQLX$ = "ALTER TABLE ENCAPRFUN ADD ENCAPRFUN_ID INT IDENTITY"
    FLEXCONN.Execute (SQLX$)
   End If
   Call CREACAMPO("", "ENCAPRFUN", "DESCRIPCION", 10, 64)
   Call CREACAMPO("", "ENCAPRFUN", "CONSULTASQL", 10, 1024)
   Call CREACAMPO("", "ENCAPRFUN", "LISTA", 10, 64, 1)
   Call CREACAMPO("", "ENCAPRFUN", "BORRAR", 3, 0)
   Call CREACAMPO("", "ENCAPRFUN", "INSERTAR", 3, 0)
   Call CREACAMPO("", "ENCAPRFUN", "NOTILDE", 3, 0)
   Call CREACAMPO("", "ENCAPRFUN", "CANTICAMPOS", 3, 0)
   Call CREACAMPO("", "ENCAPRFUN", "CODIEMPR", 3, 2, 1)
   Call CREACAMPO("", "ENCAPRFUN", "MARBORRA", 3, 0)
   Call CREACAMPO("", "ENCAPRFUN", "ADDREGISTROS", 3, 0)
   Call CREACAMPO("", "ENCAPRFUN", "USAFORMULA", 3, 0)
   Call CREACAMPO("", "ENCAPRFUN", "ID_LISTA", 4, 0, 1)
   Call CREACAMPO("", "ENCAPRFUN", "USARECORD", 3, 0)
   Call CREACAMPO("", "ENCAPRFUN", "FECHEMISI", 8, 0)
   
   Call CREACAMPO("", "ENCAPRF", "FORMATO", 10, 24)
   If ValidaExisteCampo%("ENCAPRF_ID", "ENCAPRF") < 1 Then
    SQLX$ = "ALTER TABLE ENCAPRF ADD ENCAPRF_ID INT IDENTITY"
    FLEXCONN.Execute (SQLX$)
   End If
   
   Call CREACAMPO("", "ENCAPRF", "COLCAMECO", 10, 128)
   Call CREACAMPO("", "ENCAPRF", "ENCABEZADO", 10, 256)
   Call CREACAMPO("", "ENCAPRF", "FORMULA", 10, 256)
   Call CREACAMPO("", "ENCAPRF", "CAMPEDITAB", 10, 256)
   Call CREACAMPO("", "ENCAPRF", "TXTTOTAL", 10, 256)
   Call CREACAMPO("", "ENCAPRF", "TOTAL", 10, 256)
   Call CREACAMPO("", "ENCAPRF", "DEJARLIBRE", 10, 256)
   Call CREACAMPO("", "ENCAPRF", "MARBORRA", 3, 0)
   Call CREACAMPO("", "ENCAPRF", "ID_LISTA", 4, 0, 1)
   Call CREACAMPO("", "ENCAPRF", "CODIEMPR", 3, 2, 1)

   Call CREACAMPO("", "DATASQL", "CAMPO1", 10, 256)
   
   If ValidaExisteCampo%("DATASQL_ID", "DATASQL") < 1 Then
    SQLX$ = "ALTER TABLE DATASQL ADD DATASQL_ID INT IDENTITY"
    FLEXCONN.Execute (SQLX$)
   End If

   Call CREACAMPO("", "DATASQL", "CAMPO1", 10, 256)
   Call CREACAMPO("", "DATASQL", "CAMPO2", 10, 256)
   Call CREACAMPO("", "DATASQL", "CAMPO3", 10, 256)
   Call CREACAMPO("", "DATASQL", "CAMPO4", 10, 256)
   Call CREACAMPO("", "DATASQL", "CAMPO5", 10, 256)
   Call CREACAMPO("", "DATASQL", "CAMPO6", 10, 256)
   Call CREACAMPO("", "DATASQL", "CAMPO7", 10, 256)
   Call CREACAMPO("", "DATASQL", "CAMPO8", 10, 256)
   Call CREACAMPO("", "DATASQL", "CAMPO9", 10, 256)
   Call CREACAMPO("", "DATASQL", "CAMPO10", 10, 256)
   Call CREACAMPO("", "DATASQL", "CAMPO11", 10, 256)
   Call CREACAMPO("", "DATASQL", "CAMPO12", 10, 256)
   Call CREACAMPO("", "DATASQL", "CAMPO13", 10, 256)
   Call CREACAMPO("", "DATASQL", "CAMPO14", 10, 256)
   Call CREACAMPO("", "DATASQL", "CAMPOFE1", 8, 0)
   Call CREACAMPO("", "DATASQL", "CAMPOFE2", 8, 0)
   
   Call CREACAMPO("", "DATASQL", "MARBORRA", 3, 0, 1)
   Call CREACAMPO("", "DATASQL", "ID_LISTA", 4, 0, 1)
   Call CREACAMPO("", "DATASQL", "LISTA", 10, 64)
   
   Call CREACAMPO("", "DATASQL", "CAMPO15", 10, 256)
   Call CREACAMPO("", "DATASQL", "CAMPO16", 10, 256)
   Call CREACAMPO("", "DATASQL", "CAMPO17", 10, 256)
   Call CREACAMPO("", "DATASQL", "CAMPO18", 10, 256)
   Call CREACAMPO("", "DATASQL", "CAMPO19", 10, 256)
   Call CREACAMPO("", "DATASQL", "CAMPO20", 10, 256)


   
End Sub
Sub CREATABLA_PDIARPROD()
'   SQLX$ = "DROP TABLE PDIARPROD1 "
'   FLEXCONN.Execute (SQLX$)

    Call CREACAMPO("", "PDIARPROD1", "IDSUBOR", 10, 12, 1)
    Call CREACAMPO("", "PDIARPROD1", "INDISOLSUP", 3, 0)
    Call CREACAMPO("", "PDIARPROD1", "INDISOLINF", 3, 0)
    Call CREACAMPO("", "PDIARPROD1", "SOLAPASUP", 10, 48, 1)
    Call CREACAMPO("", "PDIARPROD1", "SOLAPAINF", 10, 48, 1)
    Call CREACAMPO("", "PDIARPROD1", "MARBORRA", 3, 0, 1)
    Call CREACAMPO("", "PDIARPROD1", "MEDIDAS", 10, 64)
    Call CREACAMPO("", "PDIARPROD1", "FECHA", 8, 0)
    Call CREACAMPO("", "PDIARPROD1", "FECHAYHORA", 8, 0)
    Call CREACAMPO("", "PDIARPROD1", "NUMMAQUI", 10, 24, 1)
    Call CREACAMPO("", "PDIARPROD1", "NAME_MAQUI", 10, 48)
    Call CREACAMPO("", "PDIARPROD1", "NUMUSUMAQ", 3, 0, 1)
    Call CREACAMPO("", "PDIARPROD1", "NAMEUSUMAQ", 10, 48)
    Call CREACAMPO("", "PDIARPROD1", "NUMAYUD1", 3, 0)
    Call CREACAMPO("", "PDIARPROD1", "NAMEAYUD1", 10, 48)
    Call CREACAMPO("", "PDIARPROD1", "NUMAYUD2", 3, 0)
    Call CREACAMPO("", "PDIARPROD1", "NAMEAYUD2", 10, 48)
    Call CREACAMPO("", "PDIARPROD1", "NUMAYUD3", 3, 0)
    Call CREACAMPO("", "PDIARPROD1", "NAMEAYUD3", 10, 48)
    Call CREACAMPO("", "PDIARPROD1", "NUMAYUD4", 3, 0)
    Call CREACAMPO("", "PDIARPROD1", "NAMEAYUD4", 10, 48)
    
    Call CREACAMPO("", "PDIARPROD1", "NUMAYUD5", 3, 0)
    Call CREACAMPO("", "PDIARPROD1", "NAMEAYUD5", 10, 48)
    Call CREACAMPO("", "PDIARPROD1", "NUMAYUD6", 3, 0)
    Call CREACAMPO("", "PDIARPROD1", "NAMEAYUD6", 10, 48)
    Call CREACAMPO("", "PDIARPROD1", "NUMAYUD7", 3, 0)
    Call CREACAMPO("", "PDIARPROD1", "NAMEAYUD7", 10, 48)
    Call CREACAMPO("", "PDIARPROD1", "NUMAYUD8", 3, 0)
    Call CREACAMPO("", "PDIARPROD1", "NAMEAYUD8", 10, 48)
    Call CREACAMPO("", "PDIARPROD1", "COD_P_DIAR", 3, 0, 1)
    Call CREACAMPO("", "PDIARPROD1", "DESC_P_DIAR", 10, 64, 1)
    Call CREACAMPO("", "PDIARPROD1", "H_DESDE", 10, 24)
    Call CREACAMPO("", "PDIARPROD1", "H_DESDE_FN", 4, 0)
    Call CREACAMPO("", "PDIARPROD1", "H_HASTA", 10, 24)
    Call CREACAMPO("", "PDIARPROD1", "H_TRABAJADAS", 7, 8)
    Call CREACAMPO("", "PDIARPROD1", "BOLSASXBTO", 4, 0)
    Call CREACAMPO("", "PDIARPROD1", "CANBULTOS", 4, 0) 'SE USA PARA EMBALAJE
    Call CREACAMPO("", "PDIARPROD1", "NROBOBINA", 4, 0) ' EQUIVALE A NRO DE PALLET
    Call CREACAMPO("", "PDIARPROD1", "CANTOK", 7, 8) ' EN CASO DE EMBALAJE EQUIVALE A BOLSAS EMBALADAS
    Call CREACAMPO("", "PDIARPROD1", "CodPallet", 3, 0)
    Call CREACAMPO("", "PDIARPROD1", "TipoPallet", 10, 64)
    Call CREACAMPO("", "PDIARPROD1", "COD_DESCAR", 3, 0, 1)
    Call CREACAMPO("", "PDIARPROD1", "DESCR_DESCAR", 10, 64)
    Call CREACAMPO("", "PDIARPROD1", "KG_DESCAR", 7, 8)
    Call CREACAMPO("", "PDIARPROD1", "OBSERVACIONES", 10, 512)
    Call CREACAMPO("", "PDIARPROD1", "NroClie", 4, 0, 1)
    Call CREACAMPO("", "PDIARPROD1", "NroPed", 4, 0, 1)
    Call CREACAMPO("", "PDIARPROD1", "NUSERLOG", 3, 0, 1)
    Call CREACAMPO("", "PDIARPROD1", "NUSERGRAB", 3, 0, 1)
    Call CREACAMPO("", "PDIARPROD1", "NULDESCRIPTAREA", 10, 64)
    Call CREACAMPO("", "PDIARPROD1", "NULCOEFICIENTE", 7, 8)
    Call CREACAMPO("", "PDIARPROD1", "NULIDOP1", 3, 0)
    Call CREACAMPO("", "PDIARPROD1", "NULNAMEOP1", 10, 48)
    Call CREACAMPO("", "PDIARPROD1", "NULIDOP2", 3, 0)
    Call CREACAMPO("", "PDIARPROD1", "NULNAMEOP2", 10, 48)
    If ValidaExisteCampo%("PDIARPROD1_ID", "PDIARPROD1") < 1 Then
     SQLX$ = "ALTER TABLE PDIARPROD1 ADD PDIARPROD1_ID INT IDENTITY"
     FLEXCONN.Execute (SQLX$)
    End If
    '
'   SQLX$ = "DROP TABLE PDIARPROD2 "
'   FLEXCONN.Execute (SQLX$)
    Call CREACAMPO("", "PDIARPROD2", "IDSUBOR", 10, 12, 1)
    Call CREACAMPO("", "PDIARPROD2", "FECHA", 8, 0)
    Call CREACAMPO("", "PDIARPROD2", "FECHAYHORA", 8, 0)
    Call CREACAMPO("", "PDIARPROD2", "NUMMAQUI", 10, 24, 1)
    Call CREACAMPO("", "PDIARPROD2", "NAME_MAQUI", 10, 48)
    Call CREACAMPO("", "PDIARPROD2", "IDOP1", 3, 0)
    Call CREACAMPO("", "PDIARPROD2", "NAMEOP1", 10, 48)
    Call CREACAMPO("", "PDIARPROD2", "IDOP2", 3, 0)
    Call CREACAMPO("", "PDIARPROD2", "NAMEOP2", 10, 48)
    Call CREACAMPO("", "PDIARPROD2", "CODTAREA", 4, 0, 1)
    Call CREACAMPO("", "PDIARPROD2", "DESCRIPTAREA", 10, 64)
    Call CREACAMPO("", "PDIARPROD2", "COEFICIENTE", 7, 8)
    Call CREACAMPO("", "PDIARPROD2", "CANTIDAD", 7, 8)
    Call CREACAMPO("", "PDIARPROD2", "H_DESDE", 10, 24)
    Call CREACAMPO("", "PDIARPROD2", "H_DESDE_FN", 4, 0)

    Call CREACAMPO("", "PDIARPROD2", "H_HASTA", 10, 24)
    Call CREACAMPO("", "PDIARPROD2", "H_TRABAJADAS", 7, 8)
    Call CREACAMPO("", "PDIARPROD2", "OBSERVACIONES", 10, 512)
    Call CREACAMPO("", "PDIARPROD2", "INDISOLSUP", 3, 0, 1)
    Call CREACAMPO("", "PDIARPROD2", "INDISOLINF", 3, 0, 1)
    Call CREACAMPO("", "PDIARPROD2", "SOLAPASUP", 10, 48, 1)
    Call CREACAMPO("", "PDIARPROD2", "SOLAPAINF", 10, 48, 1)
    Call CREACAMPO("", "PDIARPROD2", "MARBORRA", 3, 1)
    Call CREACAMPO("", "PDIARPROD2", "NroClie", 4, 0, 1)
    Call CREACAMPO("", "PDIARPROD2", "NroPed", 4, 0, 1)
    Call CREACAMPO("", "PDIARPROD2", "NUSERLOG", 3, 0)
    Call CREACAMPO("", "PDIARPROD2", "NUSERGRAB", 3, 0)
    Call CREACAMPO("", "PDIARPROD2", "MEDIDAS", 10, 64)
    Call CREACAMPO("", "PDIARPROD2", "NULNUMUSUMAQ", 3, 0)
    Call CREACAMPO("", "PDIARPROD2", "NULNAMEUSUMAQ", 10, 48)
    Call CREACAMPO("", "PDIARPROD2", "NULNUMAYUD1", 3, 0)
    Call CREACAMPO("", "PDIARPROD2", "NULNAMEAYUD1", 10, 48)
    Call CREACAMPO("", "PDIARPROD2", "NULNUMAYUD2", 3, 0)
    Call CREACAMPO("", "PDIARPROD2", "NULNAMEAYUD2", 10, 48)
    Call CREACAMPO("", "PDIARPROD2", "NULNUMAYUD3", 3, 0)
    Call CREACAMPO("", "PDIARPROD2", "NULNAMEAYUD3", 10, 48)
    Call CREACAMPO("", "PDIARPROD2", "NULNUMAYUD4", 3, 0)
    Call CREACAMPO("", "PDIARPROD2", "NULNAMEAYUD4", 10, 48)
    Call CREACAMPO("", "PDIARPROD2", "NULNUMAYUD5", 3, 0)
    Call CREACAMPO("", "PDIARPROD2", "NULNAMEAYUD5", 10, 48)
    Call CREACAMPO("", "PDIARPROD2", "NULNUMAYUD6", 3, 0)
    Call CREACAMPO("", "PDIARPROD2", "NULNAMEAYUD6", 10, 48)
    Call CREACAMPO("", "PDIARPROD2", "NULNUMAYUD7", 3, 0)
    Call CREACAMPO("", "PDIARPROD2", "NULNAMEAYUD7", 10, 48)
    Call CREACAMPO("", "PDIARPROD2", "NULNUMAYUD8", 3, 0)
    Call CREACAMPO("", "PDIARPROD2", "NULNAMEAYUD8", 10, 48)
    Call CREACAMPO("", "PDIARPROD2", "NULMEDIDAS", 10, 64)
    If ValidaExisteCampo%("PDIARPROD2_ID", "PDIARPROD2") < 1 Then
     SQLX$ = "ALTER TABLE PDIARPROD2 ADD PDIARPROD2_ID INT IDENTITY"
     FLEXCONN.Execute (SQLX$)
    End If
    '
    Call CREACAMPO("", "CONSUMO1", "IDSUBOR", 10, 12, 1)
    Call CREACAMPO("", "CONSUMO1", "IDMATERIAL", 3, 1)
    Call CREACAMPO("", "CONSUMO1", "FECHA", 8, 0)
    Call CREACAMPO("", "CONSUMO1", "FECHAYHORA", 8, 0)
    Call CREACAMPO("", "CONSUMO1", "DescriMPrima", 10, 48)
    Call CREACAMPO("", "CONSUMO1", "CODINT", 4, 0)
    Call CREACAMPO("", "CONSUMO1", "CODIGO", 10, 128)
    Call CREACAMPO("", "CONSUMO1", "DESCRIPCION", 10, 64)
    Call CREACAMPO("", "CONSUMO1", "Num_Prov", 4, 0, 1)
    Call CREACAMPO("", "CONSUMO1", "RazonSocial", 10, 64)
    Call CREACAMPO("", "CONSUMO1", "LOTE", 10, 24)
    Call CREACAMPO("", "CONSUMO1", "OBSERVACIONES", 10, 512)
    Call CREACAMPO("", "CONSUMO1", "INDISOLSUP", 3, 0)
    Call CREACAMPO("", "CONSUMO1", "INDISOLINF", 3, 0)
    Call CREACAMPO("", "CONSUMO1", "SOLAPASUP", 10, 48, 1)
    Call CREACAMPO("", "CONSUMO1", "SOLAPAINF", 10, 48, 1)
    Call CREACAMPO("", "CONSUMO1", "MARBORRA", 3, 1)
    Call CREACAMPO("", "CONSUMO1", "NroClie", 4, 0, 1)
    Call CREACAMPO("", "CONSUMO1", "NroPed", 4, 0, 1)
    Call CREACAMPO("", "CONSUMO1", "NUSERLOG", 3, 0)
    Call CREACAMPO("", "CONSUMO1", "NUSERGRAB", 3, 0)
    Call CREACAMPO("", "CONSUMO1", "CONSUMO", 7, 8)
    '
    If ValidaExisteCampo%("CONSUMO1_ID", "CONSUMO1") < 1 Then
     SQLX$ = "ALTER TABLE CONSUMO1 ADD CONSUMO1_ID INT IDENTITY"
     FLEXCONN.Execute (SQLX$)
    End If
    Call CREACAMPO("", "CONSUMO2", "IDMATERIAL", 3, 1)
    Call CREACAMPO("", "CONSUMO2", "DescriMPrima", 10, 48)
    Call CREACAMPO("", "CONSUMO2", "CODINT", 4, 0)
    Call CREACAMPO("", "CONSUMO2", "CODIGO", 10, 128)
    Call CREACAMPO("", "CONSUMO2", "DESCRIPCION", 10, 64)
    Call CREACAMPO("", "CONSUMO2", "IDSUBOR", 10, 12, 1)
    Call CREACAMPO("", "CONSUMO2", "INDISOLSUP", 3, 0)
    Call CREACAMPO("", "CONSUMO2", "INDISOLINF", 3, 0)
    Call CREACAMPO("", "CONSUMO2", "SOLAPASUP", 10, 48, 1)
    Call CREACAMPO("", "CONSUMO2", "SOLAPAINF", 10, 48, 1)
    Call CREACAMPO("", "CONSUMO2", "MARBORRA", 3, 1)
    Call CREACAMPO("", "CONSUMO2", "FECHA", 8, 0)
    Call CREACAMPO("", "CONSUMO2", "FECHAYHORA", 8, 0)
    Call CREACAMPO("", "CONSUMO2", "Num_Prov", 4, 0, 1)
    Call CREACAMPO("", "CONSUMO2", "RazonSocial", 10, 64)
    Call CREACAMPO("", "CONSUMO2", "BOBINALOTE", 10, 64)
    Call CREACAMPO("", "CONSUMO2", "ANCHO", 7, 8)
    Call CREACAMPO("", "CONSUMO2", "GRAMAJE", 7, 8)
    Call CREACAMPO("", "CONSUMO2", "COLOR", 10, 64)
    Call CREACAMPO("", "CONSUMO2", "DIAM_INICIAL", 7, 8)
    Call CREACAMPO("", "CONSUMO2", "DIAM_FINAL", 7, 8)
    Call CREACAMPO("", "CONSUMO2", "KG_INICIAL", 7, 8)
    Call CREACAMPO("", "CONSUMO2", "KG_FIINAL", 7, 8)
    Call CREACAMPO("", "CONSUMO2", "CONSUMO", 7, 8)
    Call CREACAMPO("", "CONSUMO2", "OBSERVACIONES", 10, 512)
    Call CREACAMPO("", "CONSUMO2", "NroClie", 4, 0, 1)
    Call CREACAMPO("", "CONSUMO2", "NroPed", 4, 0, 1)
    Call CREACAMPO("", "CONSUMO2", "NUSERLOG", 3, 0)
    Call CREACAMPO("", "CONSUMO2", "NUSERGRAB", 3, 0)
    If ValidaExisteCampo%("CONSUMO2_ID", "CONSUMO2") < 1 Then
     SQLX$ = "ALTER TABLE CONSUMO2 ADD CONSUMO2_ID INT IDENTITY"
     FLEXCONN.Execute (SQLX$)
    End If
    
    Call CREACAMPO("", "CONSUMO3", "IDMATERIAL", 3, 1)
    Call CREACAMPO("", "CONSUMO3", "DescriMPrima", 10, 48)
    Call CREACAMPO("", "CONSUMO3", "CODINT", 4, 0)
    Call CREACAMPO("", "CONSUMO3", "CODIGO", 10, 128)
    Call CREACAMPO("", "CONSUMO3", "DESCRIPCION", 10, 64)
    Call CREACAMPO("", "CONSUMO3", "IDSUBOR", 10, 12, 1)
    Call CREACAMPO("", "CONSUMO3", "INDISOLSUP", 3, 0)
    Call CREACAMPO("", "CONSUMO3", "INDISOLINF", 3, 0)
    Call CREACAMPO("", "CONSUMO3", "SOLAPASUP", 10, 48, 1)
    Call CREACAMPO("", "CONSUMO3", "SOLAPAINF", 10, 48, 1)
    Call CREACAMPO("", "CONSUMO3", "MARBORRA", 3, 1)
    Call CREACAMPO("", "CONSUMO3", "FECHA", 8, 0)
    Call CREACAMPO("", "CONSUMO3", "FECHAYHORA", 8, 0)
    Call CREACAMPO("", "CONSUMO3", "Num_Prov", 4, 0, 1)
    Call CREACAMPO("", "CONSUMO3", "RazonSocial", 10, 64)
    Call CREACAMPO("", "CONSUMO3", "BOBINALOTE", 10, 64)
    Call CREACAMPO("", "CONSUMO3", "ANCHO", 7, 8)
    Call CREACAMPO("", "CONSUMO3", "GRAMAJE", 7, 8)
    Call CREACAMPO("", "CONSUMO3", "COLOR", 10, 64)
    Call CREACAMPO("", "CONSUMO3", "METROS", 7, 8)
    Call CREACAMPO("", "CONSUMO3", "CONSUMO", 7, 8)
    Call CREACAMPO("", "CONSUMO3", "SALDO", 7, 8)
    Call CREACAMPO("", "CONSUMO3", "OBSERVACIONES", 10, 512)
    Call CREACAMPO("", "CONSUMO3", "NroClie", 4, 0, 1)
    Call CREACAMPO("", "CONSUMO3", "NroPed", 4, 0, 1)
    Call CREACAMPO("", "CONSUMO3", "NUSERLOG", 3, 0)
    Call CREACAMPO("", "CONSUMO3", "NUSERGRAB", 3, 0)
    If ValidaExisteCampo%("CONSUMO3_ID", "CONSUMO3") < 1 Then
     SQLX$ = "ALTER TABLE CONSUMO3 ADD CONSUMO3_ID INT IDENTITY"
     FLEXCONN.Execute (SQLX$)
    End If
    '
    Call CREACAMPO("", "CHECK1", "IDSUBOR", 10, 12, 1)
    Call CREACAMPO("", "CHECK1", "INDISOLSUP", 3, 0)
    Call CREACAMPO("", "CHECK1", "INDISOLINF", 3, 0)
    Call CREACAMPO("", "CHECK1", "SOLAPASUP", 10, 48, 1)
    Call CREACAMPO("", "CHECK1", "SOLAPAINF", 10, 48, 1)
    Call CREACAMPO("", "CHECK1", "MARBORRA", 3, 1)
    Call CREACAMPO("", "CHECK1", "FECHA", 8, 0)
    Call CREACAMPO("", "CHECK1", "FECHAYHORA", 8, 0)
    Call CREACAMPO("", "CHECK1", "DESCRIPCION", 10, 128)
    Call CREACAMPO("", "CHECK1", "ACCION", 10, 128)
    Call CREACAMPO("", "CHECK1", "CTRL", 3, 0)
    Call CREACAMPO("", "CHECK1", "OBSERVACIONES", 10, 512)
    Call CREACAMPO("", "CHECK1", "NroClie", 4, 0, 1)
    Call CREACAMPO("", "CHECK1", "NroPed", 4, 0, 1)
    Call CREACAMPO("", "CHECK1", "NUSERLOG", 3, 0)
    Call CREACAMPO("", "CHECK1", "NUSERGRAB", 3, 0)
    If ValidaExisteCampo%("CHECK1_ID", "CHECK1") < 1 Then
     SQLX$ = "ALTER TABLE CHECK1 ADD CHECK1_ID INT IDENTITY"
     FLEXCONN.Execute (SQLX$)
    End If
    
    Call CREACAMPO("", "CHECKPROCESO", "IDSUBOR", 10, 12, 1)
    Call CREACAMPO("", "CHECKPROCESO", "INDISOLSUP", 3, 0)
    Call CREACAMPO("", "CHECKPROCESO", "INDISOLINF", 3, 0)
    Call CREACAMPO("", "CHECKPROCESO", "SOLAPASUP", 10, 48, 1)
    Call CREACAMPO("", "CHECKPROCESO", "SOLAPAINF", 10, 48, 1)
    Call CREACAMPO("", "CHECKPROCESO", "MARBORRA", 3, 1)
    Call CREACAMPO("", "CHECKPROCESO", "FECHA", 8, 0)
    Call CREACAMPO("", "CHECKPROCESO", "FECHAYHORA", 8, 0)
    Call CREACAMPO("", "CHECKPROCESO", "DESCRIPCION", 10, 128)
    Call CREACAMPO("", "CHECKPROCESO", "ACCION", 10, 128)
    Call CREACAMPO("", "CHECKPROCESO", "CTRL", 3, 0)
    Call CREACAMPO("", "CHECKPROCESO", "OBSERVACIONES", 10, 512)
    Call CREACAMPO("", "CHECKPROCESO", "NroClie", 4, 0, 1)
    Call CREACAMPO("", "CHECKPROCESO", "NroPed", 4, 0, 1)
    Call CREACAMPO("", "CHECKPROCESO", "NUSERLOG", 3, 0)
    Call CREACAMPO("", "CHECKPROCESO", "NUSERGRAB", 3, 0)
    If ValidaExisteCampo%("CHECKPROCESO_ID", "CHECKPROCESO") < 1 Then
     SQLX$ = "ALTER TABLE CHECKPROCESO ADD CHECKPROCESO_ID INT IDENTITY"
     FLEXCONN.Execute (SQLX$)
    End If
    '
    Call CREACAMPO("", "DESCARCONS", "IDSUBOR", 10, 12, 1)
    Call CREACAMPO("", "DESCARCONS", "INDISOLSUP", 3, 0)
    Call CREACAMPO("", "DESCARCONS", "INDISOLINF", 3, 0)
    Call CREACAMPO("", "DESCARCONS", "SOLAPASUP", 10, 48, 1)
    Call CREACAMPO("", "DESCARCONS", "SOLAPAINF", 10, 48, 1)
    Call CREACAMPO("", "DESCARCONS", "MARBORRA", 3, 1)
    Call CREACAMPO("", "DESCARCONS", "FECHA", 8, 0)
    Call CREACAMPO("", "DESCARCONS", "FECHAYHORA", 8, 0)
    Call CREACAMPO("", "DESCARCONS", "CODIDESC", 3, 0)
    Call CREACAMPO("", "DESCARCONS", "DESCRIDESC", 10, 128)
    Call CREACAMPO("", "DESCARCONS", "CANT_DESC", 7, 8)
    Call CREACAMPO("", "DESCARCONS", "OBSERVACIONES", 10, 512)
    Call CREACAMPO("", "DESCARCONS", "NroClie", 4, 0, 1)
    Call CREACAMPO("", "DESCARCONS", "NroPed", 4, 0, 1)
    Call CREACAMPO("", "DESCARCONS", "NUSERLOG", 3, 0)
    Call CREACAMPO("", "DESCARCONS", "NUSERGRAB", 3, 0)
    Call CREACAMPO("", "DESCARCONS", "NUMMAQUI", 10, 24)
    Call CREACAMPO("", "DESCARCONS", "NAME_MAQUI", 10, 48)
    If ValidaExisteCampo%("DESCARCONS_ID", "DESCARCONS") < 1 Then
     SQLX$ = "ALTER TABLE DESCARCONS ADD DESCARCONS_ID INT IDENTITY"
     FLEXCONN.Execute (SQLX$)
    End If

    Call CREACAMPO("", "DOTACION", "IDSUBOR", 10, 12, 1)
    Call CREACAMPO("", "DOTACION", "INDISOLSUP", 3, 0)
    Call CREACAMPO("", "DOTACION", "INDISOLINF", 3, 0)
    Call CREACAMPO("", "DOTACION", "SOLAPASUP", 10, 48, 1)
    Call CREACAMPO("", "DOTACION", "SOLAPAINF", 10, 48, 1)
    Call CREACAMPO("", "DOTACION", "MARBORRA", 3, 1)
    Call CREACAMPO("", "DOTACION", "FECHA", 8, 0, 1)
    Call CREACAMPO("", "DOTACION", "FECHAYHORA", 8, 0)
    Call CREACAMPO("", "DOTACION", "CANT_DESC", 7, 8)
    Call CREACAMPO("", "DOTACION", "OBSERVACIONES", 10, 512)
    Call CREACAMPO("", "DOTACION", "NroClie", 4, 0, 1)
    Call CREACAMPO("", "DOTACION", "NUSER", 3, 0, 1)
    Call CREACAMPO("", "DOTACION", "NOMBRE", 10, 48)
    Call CREACAMPO("", "DOTACION", "H_DESDE", 10, 24, 1)
    Call CREACAMPO("", "DOTACION", "H_DESDE_FN", 4, 0)
    Call CREACAMPO("", "DOTACION", "H_HASTA", 10, 24, 1)
    Call CREACAMPO("", "DOTACION", "H_TRABAJADAS", 7, 8)
    Call CREACAMPO("", "DOTACION", "NUMMAQUI", 10, 24, 1)
    Call CREACAMPO("", "DOTACION", "NAME_MAQUI", 10, 48)
    If ValidaExisteCampo%("DOTACION_ID", "DOTACION") < 1 Then
     SQLX$ = "ALTER TABLE DOTACION ADD DOTACION_ID INT IDENTITY"
     FLEXCONN.Execute (SQLX$)
    End If
    '
End Sub

Sub GENERATABLAIMPRE()
   '
   Call ABRECONEXION
   Call CREACAMPO("", "TABLA_IMPRE", "CAMPO1", 10, 128)
   
   If ValidaExisteCampo%("TABLA_IMPRE_ID", "TABLA_IMPRE") < 1 Then
    SQLX$ = "ALTER TABLE TABLA_IMPRE ADD TABLA_IMPRE_ID INT IDENTITY"
    FLEXCONN.Execute (SQLX$)
   End If
   
   Call CREACAMPO("", "TABLA_IMPRE", "CAMPO2", 10, 128)
   Call CREACAMPO("", "TABLA_IMPRE", "CAMPO3", 10, 128)
   Call CREACAMPO("", "TABLA_IMPRE", "CAMPO4", 10, 128)
   Call CREACAMPO("", "TABLA_IMPRE", "CAMPO5", 10, 128)
   Call CREACAMPO("", "TABLA_IMPRE", "CAMPO6", 10, 128)
   Call CREACAMPO("", "TABLA_IMPRE", "CAMPO7", 10, 128)
   Call CREACAMPO("", "TABLA_IMPRE", "CAMPO8", 10, 128)
   Call CREACAMPO("", "TABLA_IMPRE", "CAMPO9", 10, 128)
   Call CREACAMPO("", "TABLA_IMPRE", "CAMPO10", 10, 128)
   Call CREACAMPO("", "TABLA_IMPRE", "CAMPO11", 10, 128)
   Call CREACAMPO("", "TABLA_IMPRE", "CAMPO12", 10, 128)
   Call CREACAMPO("", "TABLA_IMPRE", "CAMPO13", 10, 128)
   Call CREACAMPO("", "TABLA_IMPRE", "CAMPO14", 10, 128)
   Call CREACAMPO("", "TABLA_IMPRE", "CAMPO15", 10, 128)
   Call CREACAMPO("", "TABLA_IMPRE", "CAMPO16", 10, 128)
   Call CREACAMPO("", "TABLA_IMPRE", "CAMPO17", 10, 128)
   Call CREACAMPO("", "TABLA_IMPRE", "CAMPO18", 10, 128)
   Call CREACAMPO("", "TABLA_IMPRE", "CAMPO19", 10, 128)
   Call CREACAMPO("", "TABLA_IMPRE", "CAMPO20", 10, 128)
   Call CREACAMPO("", "TABLA_IMPRE", "CAMPO21", 10, 128)
   Call CREACAMPO("", "TABLA_IMPRE", "CAMPO22", 10, 128)
   Call CREACAMPO("", "TABLA_IMPRE", "CAMPO23", 10, 128)
   Call CREACAMPO("", "TABLA_IMPRE", "CAMPO24", 10, 128)
   Call CREACAMPO("", "TABLA_IMPRE", "CAMPO25", 10, 128)
   Call CREACAMPO("", "TABLA_IMPRE", "CAMPO26", 10, 128)
   Call CREACAMPO("", "TABLA_IMPRE", "CAMPO27", 10, 128)
   Call CREACAMPO("", "TABLA_IMPRE", "CAMPO28", 10, 128)
   Call CREACAMPO("", "TABLA_IMPRE", "CAMPO29", 10, 128)
   Call CREACAMPO("", "TABLA_IMPRE", "CAMPO30", 10, 128)
   Call CREACAMPO("", "TABLA_IMPRE", "CAMPO31", 10, 128)
   
   Call CREACAMPO("", "TABLA_IMPRE", "IDCONSUL", 4, 0, 1)
   Call CREACAMPO("", "TABLA_IMPRE", "FECHA", 8, 0)
   
   '
   '
   '
End Sub

Function EXISTEVISTA%(VISTA$)
   'DEVUELVE 0 SI NO EXISTE
   'DEVUELVE 1 SI EXISTE
   If VersionBD% < 2017 Then
     SQLX$ = "IF EXISTS (SELECT 1 FROM sysobjects  WHERE xtype='V' AND name='" & VISTA$ & "') SELECT '1'ELSE SELECT '0'"
   Else
     'SQLX$ = "IF EXISTS (SELECT * FROM sysobjects WHERE id = OBJECT_ID(N'[dbo].[" & VISTA$ & "]') )SELECT '1' ELSE SELECT '0'"
     SQLX$ = "IF EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[" & VISTA$ & "]') AND type IN (N'V', N'U')) "
     SQLX$ = SQLX$ + "SELECT '1' Else SELECT '0' ;"

   End If
   Set rs = READSET(SQLX$, , 1)
   EXISTEVISTA% = XVALO(rs.Fields(0))

   rs.Close
   Set rs = Nothing
End Function
Sub CREARVISTAPROD1Y2()
    If EXISTEVISTA%("PROD1Y2") < 1 Then
       SQLX$ = "CREATE VIEW PROD1Y2 AS"
    Else
       SQLX$ = "ALTER VIEW PROD1Y2 AS"
    End If
    
    If VersionBD% > 2016 Then
       SQLX$ = "CREATE OR ALTER VIEW PROD1Y2 AS"
    End If
    
       SQLX$ = SQLX$ + " SELECT  P.H_DESDE,P.H_HASTA,P.NUMAYUD1,P.NUMAYUD2,P.NUMAYUD3,P.NUMAYUD4,P.NUMAYUD5,P.NUMAYUD6,P.NUMAYUD7"
       SQLX$ = SQLX$ + " ,P.NUMAYUD8,P.FECHA,P.IDSUBOR,P.NULIDOP1,P.NULIDOP2,P.SOLAPASUP,P.PDIARPROD1_ID,P.INDISOLSUP"
       SQLX$ = SQLX$ + " ,P.NUMUSUMAQ,P.H_TRABAJADAS  ,P.NAMEUSUMAQ,P.NAMEAYUD1,P.NAMEAYUD2,P.NAMEAYUD3,P.NAMEAYUD4"
       SQLX$ = SQLX$ + " ,P.NAMEAYUD5,P.NAMEAYUD6,P.NAMEAYUD7,P.NAMEAYUD8,P.NULNAMEOP1,P.NULNAMEOP2,P.DESC_P_DIAR,P.NUMMAQUI,P.FECHAYHORA,P.NROCLIE,P.NROPED,P.CANTOK "
       SQLX$ = SQLX$ + " FROM (SELECT H_DESDE,H_HASTA,NUMAYUD1,NUMAYUD2 , NUMAYUD3, NUMAYUD4, NUMAYUD5 , NUMAYUD6, NUMAYUD7"
       SQLX$ = SQLX$ + " , NUMAYUD8, FECHA, IDSUBOR, NULIDOP1, NULIDOP2  ,SOLAPASUP, INDISOLSUP,PDIARPROD1_ID,NUMUSUMAQ"
       SQLX$ = SQLX$ + " ,H_TRABAJADAS,NAMEUSUMAQ,NAMEAYUD1,NAMEAYUD2,  NAMEAYUD3,NAMEAYUD4,NAMEAYUD5,NAMEAYUD6,NAMEAYUD7"
       SQLX$ = SQLX$ + " ,NAMEAYUD8,NULNAMEOP1,NULNAMEOP2,DESC_P_DIAR,NUMMAQUI,FECHAYHORA,NROCLIE,NROPED,CANTOK  FROM PDIARPROD1  "
       SQLX$ = SQLX$ + " union all  "
       SQLX$ = SQLX$ + " SELECT H_DESDE,H_HASTA,NULNUMAYUD1,NULNUMAYUD2,NULNUMAYUD3,NULNUMAYUD4,NULNUMAYUD5,NULNUMAYUD6 "
       SQLX$ = SQLX$ + " ,NULNUMAYUD7,NULNUMAYUD8,FECHA,IDSUBOR,IDOP1,IDOP2,SOLAPASUP ,INDISOLSUP,PDIARPROD2_ID,NULNUMUSUMAQ"
       SQLX$ = SQLX$ + " ,H_TRABAJADAS,NULNAMEUSUMAQ,NULNAMEAYUD1 ,NULNAMEAYUD2,NULNAMEAYUD3,NULNAMEAYUD4,NULNAMEAYUD5"
       SQLX$ = SQLX$ + " ,NULNAMEAYUD6,NULNAMEAYUD7,NULNAMEAYUD8 ,NAMEOP1 , NAMEOP2,DESCRIPTAREA,NUMMAQUI,FECHAYHORA,NROCLIE,NROPED,CANTIDAD FROM PDIARPROD2) "
       SQLX$ = SQLX$ + " AS P"
       FLEXCONN.Execute (SQLX$)
'    End If
End Sub
Sub CREATABLA_CONDIVENTA()
    Call CREACAMPO("", "LICONPAG", "CODICONPAG", 3, 0, 1)
    Call CREACAMPO("", "LICONPAG", "REFERENCIA", 10, 64)
    Call CREACAMPO("", "LICONPAG", "CtaAPartirde", 3, 0) 'EL CAMPO QUE ES 1)F.F., 2) 7 DIAS F.F., 3)15 DIAS F.F.,4)30 DIAS F.F.
    Call CREACAMPO("", "LICONPAG", "P1", 3, 0)
    Call CREACAMPO("", "LICONPAG", "P2", 3, 0)
    Call CREACAMPO("", "LICONPAG", "P3", 3, 0)
    Call CREACAMPO("", "LICONPAG", "P4", 3, 0)
    Call CREACAMPO("", "LICONPAG", "P5", 3, 0)
    Call CREACAMPO("", "LICONPAG", "P6", 3, 0)
    If ValidaExisteCampo%("LICONPAG_ID", "LICONPAG") < 1 Then
     SQLX$ = "ALTER TABLE LICONPAG ADD LICONPAG_ID INT IDENTITY"
     FLEXCONN.Execute (SQLX$)
    End If
End Sub
Sub CREATABLA_PROFORMTL()
    'HASTA ACA LLEGUE CON OR-19-0002 AMBEST
    'PROVORMA DE TEXTO LIBRE
    Call CREACAMPO("", "PROFORMTL", "REFERENCIA", 10, 512)
    If ValidaExisteCampo%("PROFORMTL_ID", "PROFORMTL") < 1 Then
     SQLX$ = "ALTER TABLE PROFORMTL ADD PROFORMTL_ID INT IDENTITY"
     FLEXCONN.Execute (SQLX$)
    End If
    
    Call CREACAMPO("", "PROFORMTL", "CODICONPAG", 3, 0, 1)
    Call CREACAMPO("", "PROFORMTL", "NUCLIEN", 4, 2, 1)
    Call CREACAMPO("", "PROFORMTL", "BASE1", 7, 0)
    Call CREACAMPO("", "PROFORMTL", "BASE2", 7, 0)
    Call CREACAMPO("", "PROFORMTL", "TASA1", 10, 8)
    Call CREACAMPO("", "PROFORMTL", "TASA2", 10, 8)
    Call CREACAMPO("", "PROFORMTL", "COTIZA_DOLAR", 7, 0)
    
    Call CREACAMPO("", "PROFORMTL", "NumeComp", 4, 0, 1)
    Call CREACAMPO("", "PROFORMTL", "Status", 3, 0, 1)
    Call CREACAMPO("", "PROFORMTL", "Mone_Emis", 3, 2)
    Call CREACAMPO("", "PROFORMTL", "Tipo_Cam", 7, 0)
    Call CREACAMPO("", "PROFORMTL", "FeHoReal", 8, 0)
    Call CREACAMPO("", "PROFORMTL", "Fechemisi", 8, 0)
    
    Call CREACAMPO("", "PROFORMTL", "TotNeto", 7, 0)
    Call CREACAMPO("", "PROFORMTL", "ImpoTot", 7, 0)
    Call CREACAMPO("", "PROFORMTL", "FacturaReal", 10, 24)
    Call CREACAMPO("", "PROFORMTL", "FechFactura", 8, 0)
    Call CREACAMPO("", "PROFORMTL", "Vend_Cli", 3, 0)
    Call CREACAMPO("", "PROFORMTL", "NUMETALO", 3, 0)
    Call CREACAMPO("", "PROFORMTL", "NROPED", 4, 0)
    Call CREACAMPO("", "PROFORMTL", "IMPOIVA1", 7, 0)
    Call CREACAMPO("", "PROFORMTL", "IMPOIVA2", 7, 0)
    Call CREACAMPO("", "PROFORMTL", "VARIMOVIM", 10, 8)
    Call CREACAMPO("", "PROFORMTL", "TIPOFACT", 10, 8) 'SI ES TEXTO LIBRE = TL O SI NO MT (MATERIAL CODIFICADO)
    
    
    'TABLA DETAPROFO (DETALLE DE PROFORMA)
    Call CREACAMPO("", "DETAPROFO", "RaSoClie", 10, 128)
    If ValidaExisteCampo%("DETAPROFO_ID", "DETAPROFO") < 1 Then
     SQLX$ = "ALTER TABLE DETAPROFO ADD DETAPROFO_ID INT IDENTITY"
     FLEXCONN.Execute (SQLX$)
    End If
    
    Call CREACAMPO("", "DETAPROFO", "NumeComp", 4, 0, 1)
    Call CREACAMPO("", "DETAPROFO", "CodiEmpr", 3, 0, 1)
    Call CREACAMPO("", "DETAPROFO", "PreList", 7, 0)
    Call CREACAMPO("", "DETAPROFO", "CanAsig", 7, 0)
    Call CREACAMPO("", "DETAPROFO", "CanPed", 7, 0)
    Call CREACAMPO("", "DETAPROFO", "PreUnid", 7, 0)
    Call CREACAMPO("", "DETAPROFO", "ImpoTot", 7, 0)
    Call CREACAMPO("", "DETAPROFO", "ImpoAnti", 7, 0)
    Call CREACAMPO("", "DETAPROFO", "CantEnt", 7, 0)
    Call CREACAMPO("", "DETAPROFO", "CanSaldo", 7, 0)
    Call CREACAMPO("", "DETAPROFO", "FeUltEnt", 8, 0)
    Call CREACAMPO("", "DETAPROFO", "FechPrepa", 8, 0)
    Call CREACAMPO("", "DETAPROFO", "FApro_Nope", 8, 0)
    Call CREACAMPO("", "DETAPROFO", "NroRemi", 10, 48)
    Call CREACAMPO("", "DETAPROFO", "TPedItem", 10, 32)
    Call CREACAMPO("", "DETAPROFO", "NroPed", 4, 0, 1)
    Call CREACAMPO("", "DETAPROFO", "TipoPed", 10, 16)
    Call CREACAMPO("", "DETAPROFO", "FechEmis", 8, 0)
    Call CREACAMPO("", "DETAPROFO", "NroClie", 4, 0, 1)
    Call CREACAMPO("", "DETAPROFO", "NroOcom", 10, 32)
    Call CREACAMPO("", "DETAPROFO", "Status", 3, 0)
    Call CREACAMPO("", "DETAPROFO", "ListPre", 3, 0)
    Call CREACAMPO("", "DETAPROFO", "CondPag", 3, 0)
    Call CREACAMPO("", "DETAPROFO", "Vended", 3, 0)
    Call CREACAMPO("", "DETAPROFO", "PorComVend", 6, 0)
    Call CREACAMPO("", "DETAPROFO", "NuOrItem", 3, 2)
    Call CREACAMPO("", "DETAPROFO", "CodiInt", 4, 0, 1)
    Call CREACAMPO("", "DETAPROFO", "CodiExt", 10, 48)
    Call CREACAMPO("", "DETAPROFO", "Descrip", 10, 128)
    Call CREACAMPO("", "DETAPROFO", "SubCod", 10, 32)
    Call CREACAMPO("", "DETAPROFO", "VenRep", 10, 16)
    Call CREACAMPO("", "DETAPROFO", "MoneEmis", 3, 2)
    Call CREACAMPO("", "DETAPROFO", "TipoCam", 7, 0)
    Call CREACAMPO("", "DETAPROFO", "DescLista", 7, 0)
    Call CREACAMPO("", "DETAPROFO", "NumeroSerie", 10, 64)
    Call CREACAMPO("", "DETAPROFO", "IdenLote", 10, 32)
    Call CREACAMPO("", "DETAPROFO", "LISTA_DE_ITEM", 3, 0)
    Call CREACAMPO("", "DETAPROFO", "ObservaItem", 10, 1024)
    Call CREACAMPO("", "DETAPROFO", "VALOBADA_1", 10, 512)
    
    
End Sub
Sub CREATABLA_CONTACSCOM()
    Call ABRECLIEN
    Call CREACAMPO("", "CONTACSCOM", "CODIEMPR", 3, 0)
    Call CREACAMPO("", "CONTACSCOM", "CODICLI", 10, 12)
    Call CREACAMPO("", "CONTACSCOM", "ORDI_CTAC", 3, 0)
    Call CREACAMPO("", "CONTACSCOM", "NOME_CTAC", 10, 64)
    Call CREACAMPO("", "CONTACSCOM", "CARG_CTAC", 10, 64)
    Call CREACAMPO("", "CONTACSCOM", "MAIL_CTAC", 10, 128)
    Call CREACAMPO("", "CONTACSCOM", "TELE_CTAC", 10, 64)
    Call CREACAMPO("", "CONTACSCOM", "CELU_CTAC", 10, 64)
    Call CREACAMPO("", "CONTACSCOM", "MARBORRA", 3, 0)

End Sub

Function ValidaExisteCampo%(NOMBRECAMPO$, NombreTabla$)
    SQLX$ = "SELECT * From INFORMATION_SCHEMA.Columns WITH(NOLOCK)"
    SQLX$ = SQLX$ + " WHERE COLUMN_NAME = '" & NOMBRECAMPO$ & "'"
    SQLX$ = SQLX$ + " AND TABLE_NAME = '" & NombreTabla$ & "'"
    Set rs = READSET(SQLX$)
    With rs
      If .EOF Then
        'NO EXISTE
        ValidaExisteCampo% = 0
      Else
        ValidaExisteCampo% = 1
      End If
    End With
    rs.Close
    Set rs = Nothing
End Function

Sub CREATABLA_COTIDAICHI()

    Call CREACAMPO("", "COTIDAICHI", "Descripcion", 10, 64)
    If ValidaExisteCampo%("COTIDAICHI_ID", "COTIDAICHI") < 1 Then
     SQLX$ = "ALTER TABLE COTIDAICHI ADD COTIDAICHI_ID INT IDENTITY"
     FLEXCONN.Execute (SQLX$)
    End If
       
   Call CREACAMPO("", "COTIDAICHI", "NRO_COTIZA", 4, 0, 1)
   Call CREACAMPO("", "COTIDAICHI", "VRS_COTIZA", 4, 0, 1)
   Call CREACAMPO("", "COTIDAICHI", "MODO", 3, 0, 1)

   Call CREACAMPO("", "COTIDAICHI", "FECHEMISI", 8, 0)
   Call CREACAMPO("", "COTIDAICHI", "FECHCOMPRA", 8, 0)
   Call CREACAMPO("", "COTIDAICHI", "FeSolEnt", 8, 0)
   Call CREACAMPO("", "COTIDAICHI", "Cod_Inte", 4, 0, 1)
   Call CREACAMPO("", "COTIDAICHI", "Codigo", 10, 24)
   Call CREACAMPO("", "COTIDAICHI", "MODELO", 10, 128) 'CODIGO DEL CLIENTE
   Call CREACAMPO("", "COTIDAICHI", "ARMADO", 10, 6) 'ARMADO
   Call CREACAMPO("", "COTIDAICHI", "NUMARMADO", 3, 0) 'NUMERO ARMADO
   
   Call CREACAMPO("", "COTIDAICHI", "ESTADO", 10, 12)
   Call CREACAMPO("", "COTIDAICHI", "Ctac_Cli", 10, 48)
   Call CREACAMPO("", "COTIDAICHI", "NUME_CLI", 4, 0, 1)
   Call CREACAMPO("", "COTIDAICHI", "Codi_Cli", 10, 12, 1)
   Call CREACAMPO("", "COTIDAICHI", "CANTIDAD", 7, 0)
   Call CREACAMPO("", "COTIDAICHI", "LARGO", 7, 0)
   Call CREACAMPO("", "COTIDAICHI", "ANCHO", 7, 0)
   Call CREACAMPO("", "COTIDAICHI", "SUPERFX1", 7, 0)
   Call CREACAMPO("", "COTIDAICHI", "SUPERFXLOTE", 7, 0)
   'CARAS
   Call CREACAMPO("", "COTIDAICHI", "CARAS", 3, 0)
   Call CREACAMPO("", "COTIDAICHI", "Timate", 3, 0, 1)
   Call CREACAMPO("", "COTIDAICHI", "ESPMAT", 3, 0)
   Call CREACAMPO("", "COTIDAICHI", "ESPCOBRE", 3, 0)
   Call CREACAMPO("", "COTIDAICHI", "PTH", 3, 0)
   Call CREACAMPO("", "COTIDAICHI", "PISTAS", 3, 0)
        
   Call CREACAMPO("", "COTIDAICHI", "TEXCARAS", 10, 24)
   Call CREACAMPO("", "COTIDAICHI", "TEXTIMATE", 10, 24)
   Call CREACAMPO("", "COTIDAICHI", "TEXESPMAT", 10, 24)
   Call CREACAMPO("", "COTIDAICHI", "TEXESPECOBRE", 10, 24)
   Call CREACAMPO("", "COTIDAICHI", "TEXPTH", 10, 24)
   Call CREACAMPO("", "COTIDAICHI", "TEXPISTAS", 10, 24)
   '
   'PALELIZADO Y CIRCUITO
   Call CREACAMPO("", "COTIDAICHI", "CIRCUITO", 10, 6)
   Call CREACAMPO("", "COTIDAICHI", "PANELIZADO", 10, 6)

   'ESTADO ACATUAL
   Call CREACAMPO("", "COTIDAICHI", "FtainfTec", 10, 6) 'falta informacion tecnica
   Call CREACAMPO("", "COTIDAICHI", "PrvApMtra", 10, 6) 'previa aprobacion muestra
   Call CREACAMPO("", "COTIDAICHI", "PrvAppgo", 10, 6) 'previa aprobacion pago
   Call CREACAMPO("", "COTIDAICHI", "NoStdar", 10, 6) 'no standar
   
   'PROGRAMAPCB
   Call CREACAMPO("", "COTIDAICHI", "PrgPCB", 3, 0) 'programa pcb
   Call CREACAMPO("", "COTIDAICHI", "TexPrgPCB", 10, 24) 'texto programa pcb

   'MASCARA ANTISOLDANTE
   Call CREACAMPO("", "COTIDAICHI", "LS", 10, 6) 'L.S.
   Call CREACAMPO("", "COTIDAICHI", "LC", 10, 6) 'L.C..
   Call CREACAMPO("", "COTIDAICHI", "MASCSOLDANTE", 3, 0) 'MASCARA ANTISOLDANTE
   Call CREACAMPO("", "COTIDAICHI", "TexMASCSOLDANTE", 10, 24) 'texto MASCARA ANTISOLDANTE
   'LEYENDA COMPONENTES
   Call CREACAMPO("", "COTIDAICHI", "LEYECOM", 3, 0) 'LEYENDA COMPONENTE
   Call CREACAMPO("", "COTIDAICHI", "TexLEYECOM", 10, 24) 'texto LEYENDA COMPONENTE
   Call CREACAMPO("", "COTIDAICHI", "LCOMP_LS", 10, 6) 'LEYEDA COMPONETES L.S.
   Call CREACAMPO("", "COTIDAICHI", "LCOMP_LC", 10, 6) 'LEYEDA COMPONETESL.C.
  
   'EXCEDENTE
   Call CREACAMPO("", "COTIDAICHI", "PRDEXACTA", 10, 6) 'PRODUCCION EXACTA
   Call CREACAMPO("", "COTIDAICHI", "ACEXCTE", 10, 6) 'ACEPTA EXCEDENTE
   Call CREACAMPO("", "COTIDAICHI", "TOLERANCIA", 3, 0) 'TOLERANCIA
   
   'PLAZOENTREGA Y ADICIONAL
   Call CREACAMPO("", "COTIDAICHI", "PLZ_ENT", 3, 0) 'DIAS DE PLAZO DE ENTREGA
   Call CREACAMPO("", "COTIDAICHI", "PLZ_ADIC_ARM", 3, 0) 'DIAS DE PLAZO DE ARMADO

   'ESPECIALES
   Call CREACAMPO("", "COTIDAICHI", "ESP_GRAFITO", 10, 6) 'GRAFITO
   Call CREACAMPO("", "COTIDAICHI", "ESP_GOMA", 10, 6) 'GOMA
   Call CREACAMPO("", "COTIDAICHI", "ESP_NI_AU", 10, 6) 'NIQUEL ORO
   Call CREACAMPO("", "COTIDAICHI", "ESP_OTRO", 10, 6) 'OTRO
   Call CREACAMPO("", "COTIDAICHI", "ESP_SUPMM2", 7, 0) 'SUPERFICIE EN MM2

   'CORTE DE PLACA
   Call CREACAMPO("", "COTIDAICHI", "ROUTER", 10, 6)
   Call CREACAMPO("", "COTIDAICHI", "SCORING", 10, 6)
   'ENTREGA
   Call CREACAMPO("", "COTIDAICHI", "PROTOTIPO", 10, 6)
   '
   Call CREACAMPO("", "COTIDAICHI", "TEST_ELECT", 10, 6) 'TEST ELECTRICO
   Call CREACAMPO("", "COTIDAICHI", "PNLIZADO", 10, 6) 'PANELIZADO
   Call CREACAMPO("", "COTIDAICHI", "MASCFONDO", 10, 6) 'MASCARA FONDO
   Call CREACAMPO("", "COTIDAICHI", "PCSO_INI_MASC", 10, 6) 'PROCESO INICIO MASCARA
   
   Call CREACAMPO("", "COTIDAICHI", "REFPANELIZADO", 10, 24) 'MASCARA FONDO
   Call CREACAMPO("", "COTIDAICHI", "REFMASDEFONDO", 10, 24) 'PROCESO INICIO MASCARA

   'MONTAJE
   Call CREACAMPO("", "COTIDAICHI", "MTJ_SMDBOT1", 3, 0) 'COMPONENTES SMD
   Call CREACAMPO("", "COTIDAICHI", "MTJ_SMDTOP1", 3, 0) 'COMPONENTES SMD
   Call CREACAMPO("", "COTIDAICHI", "MTJ_INSBOT1", 3, 0) 'COMPONENTES DE INSERCION
   Call CREACAMPO("", "COTIDAICHI", "MTJ_INSTOP1", 3, 0) 'COMPONENTES DE INSERCION
   
   Call CREACAMPO("", "COTIDAICHI", "MTJ_SMDBOT2", 3, 0) 'COMPONENTES SMD
   Call CREACAMPO("", "COTIDAICHI", "MTJ_SMDTOP2", 3, 0) 'COMPONENTES SMD
   Call CREACAMPO("", "COTIDAICHI", "MTJ_INSBOT2", 3, 0) 'COMPONENTES DE INSERCION
   Call CREACAMPO("", "COTIDAICHI", "MTJ_INSTOP2", 3, 0) 'COMPONENTES DE INSERCION
   
   'RESUMEN
   Call CREACAMPO("", "COTIDAICHI", "GASPREPA", 7, 0) 'GASTOS PREPARATORIOS
   Call CREACAMPO("", "COTIDAICHI", "CONPAG", 3, 0) 'NUMERO DE CONDICION DE PAGO
   Call CREACAMPO("", "COTIDAICHI", "PRC_FJO_DM2", 7, 0) 'PRECIO FIJO DM2
   Call CREACAMPO("", "COTIDAICHI", "DISENO", 3, 0) 'DISEÑO
   Call CREACAMPO("", "COTIDAICHI", "PRC_DM2", 7, 0) 'PRECIO DM2
   Call CREACAMPO("", "COTIDAICHI", "Preunit", 7, 0) 'PRECIO UNITARIO
   Call CREACAMPO("", "COTIDAICHI", "PrcTtPlca", 7, 0) 'PRECIO TOTAL PLACAS
   Call CREACAMPO("", "COTIDAICHI", "PrcTtOpcion", 7, 0) 'PRECIO TOTAL OPERACION
   Call CREACAMPO("", "COTIDAICHI", "PreunitCto", 7, 0) 'PRECIO UNITARIO CIRCUITO
   Call CREACAMPO("", "COTIDAICHI", "Preunit_Arm", 7, 0) 'PRECIO UNITARIO DE ARAMADO
   Call CREACAMPO("", "COTIDAICHI", "Preunit_Cpte", 7, 0) 'PRECIO UNITARIO DE COMPONENTE
   Call CREACAMPO("", "COTIDAICHI", "PrcTtPlcaSA", 7, 0) 'PRECIO TOTAL PLACAS S/ARMADO
   Call CREACAMPO("", "COTIDAICHI", "TtOpcion", 7, 0) 'PRECIO TOTAL OPCION
   
    Call CREACAMPO("", "COTIDAICHI", "Prcdm2Costo", 7, 0)
    Call CREACAMPO("", "COTIDAICHI", "PrcUnitCosto", 7, 0)
    Call CREACAMPO("", "COTIDAICHI", "totPlcaCosto", 7, 0)
    Call CREACAMPO("", "COTIDAICHI", "PrcTtOpcionCosto", 7, 0)
   
   Call CREACAMPO("", "COTIDAICHI", "CantStencil", 7, 0) 'CANTIDAD TOTAL STENCIL
   Call CREACAMPO("", "COTIDAICHI", "StTotalusd", 7, 0) 'TOTAL DOLARES STENCIL
  
   Call CREACAMPO("", "COTIDAICHI", "ANTICIPO", 7, 0) 'ANTICIPO
   Call CREACAMPO("", "COTIDAICHI", "SALDO", 7, 0) 'SALDO

   
   Call CREACAMPO("", "COTIDAICHI", "ObservaAdmin", 10, 2048) 'OBSERVACIONES DE ADMINISTARACION
   Call CREACAMPO("", "COTIDAICHI", "NroOcom", 10, 32) 'NUMERO DE ORDEN DE COMPRA DE CLIENTE
   Call CREACAMPO("", "COTIDAICHI", "ObservaPlanta", 10, 2048) 'OBSERVACIONES DE PLANTA
   Call CREACAMPO("", "COTIDAICHI", "ObservCliente", 10, 2048) 'OBSERVACIONES AL CLIENTE
   Call CREACAMPO("", "COTIDAICHI", "ObservPenal", 10, 2048)
   Call CREACAMPO("", "COTIDAICHI", "NUSER", 3, 0)
   Call CREACAMPO("", "COTIDAICHI", "CANTICAPAS", 3, 0)
   Call CREACAMPO("", "COTIDAICHI", "NroPed", 4, 0, 1)
   
   Call CREACAMPO("", "DOCURELCOT", "NRO_COTIZA", 4, 0, 1)
   Call CREACAMPO("", "DOCURELCOT", "Descripcion", 10, 24)
   Call CREACAMPO("", "DOCURELCOT", "Ruta", 10, 512)
   Call CREACAMPO("", "DOCURELCOT", "Marborra", 3, 0)
   '
   
    Call CREACAMPO("", "TABLAUM2", "INTERPOLAR", 10, 8)
    If ValidaExisteCampo%("TABLAUM2_ID", "TABLAUM2") < 1 Then
     SQLX$ = "ALTER TABLE TABLAUM2 ADD TABLAUM2_ID INT IDENTITY"
     FLEXCONN.Execute (SQLX$)
    End If

   Call CREACAMPO("", "TABLAUM2", "Lote_M2", 7, 0, 1)
   Call CREACAMPO("", "TABLAUM2", "FR4D16SEMI35", 7, 0)
   Call CREACAMPO("", "TABLAUM2", "FR2D16SEMI35", 7, 0)
   Call CREACAMPO("", "TABLAUM2", "FR4D16Doble", 7, 0)
   
   Call CREACAMPO("", "TABLAUM2", "PoliesterSimple", 7, 0)
   Call CREACAMPO("", "TABLAUM2", "Polymida", 7, 0)
   Call CREACAMPO("", "TABLAUM2", "FR4S24SEMI35", 7, 0)
   
   Call CREACAMPO("", "TABLAUM2", "FR4S32SEMI35", 7, 0)
   Call CREACAMPO("", "TABLAUM2", "FR4S08SEMI35", 7, 0)
   Call CREACAMPO("", "TABLAUM2", "FR4S10SEMI35", 7, 0)
   Call CREACAMPO("", "TABLAUM2", "FR4S12SEMI35", 7, 0)
   Call CREACAMPO("", "TABLAUM2", "FactorCosto", 7, 0)
   Call CREACAMPO("", "TABLAUM2", "Marborra", 3, 0)
   
    Call CREACAMPO("", "TABLAUM2COS", "INTERPOLAR", 10, 8)
    If ValidaExisteCampo%("TABLAUM2COS_ID", "TABLAUM2COS") < 1 Then
     SQLX$ = "ALTER TABLE TABLAUM2COS ADD TABLAUM2COS_ID INT IDENTITY"
     FLEXCONN.Execute (SQLX$)
    End If

   Call CREACAMPO("", "TABLAUM2COS", "Lote_M2", 7, 0, 1)
   Call CREACAMPO("", "TABLAUM2COS", "FR4SimpleCu", 7, 0)
   Call CREACAMPO("", "TABLAUM2COS", "FR4SimpleCel", 7, 0)
   Call CREACAMPO("", "TABLAUM2COS", "FR2SimpleCu", 7, 0)
   
   Call CREACAMPO("", "TABLAUM2COS", "FR4DobleCu", 7, 0)
   Call CREACAMPO("", "TABLAUM2COS", "FR4DobleCel", 7, 0)
   Call CREACAMPO("", "TABLAUM2COS", "PoliesterSimple", 7, 0)
   Call CREACAMPO("", "TABLAUM2COS", "FR4S24SEMI35", 7, 0)
   
   Call CREACAMPO("", "TABLAUM2COS", "FR4S32SEMI35", 7, 0)
   Call CREACAMPO("", "TABLAUM2COS", "Polymida", 7, 0)
   Call CREACAMPO("", "TABLAUM2COS", "FR4S08SEMI35", 7, 0)
   Call CREACAMPO("", "TABLAUM2COS", "FR4S10SEMI35", 7, 0)
   
   Call CREACAMPO("", "TABLAUM2COS", "FR4S12SEMI35", 7, 0)
   Call CREACAMPO("", "TABLAUM2COS", "FactorCosto", 7, 0)
   Call CREACAMPO("", "TABLAUM2COS", "Marborra", 3, 0)
   

    Call CREACAMPO("", "UM2PTH", "INTERPOLAR", 10, 8)
    If ValidaExisteCampo%("UM2PTH_ID", "UM2PTH") < 1 Then
     SQLX$ = "ALTER TABLE UM2PTH ADD UM2PTH_ID INT IDENTITY"
     FLEXCONN.Execute (SQLX$)
    End If

   Call CREACAMPO("", "UM2PTH", "Lote_M2", 7, 0, 1)
   Call CREACAMPO("", "UM2PTH", "FR4D16PTH18", 7, 0)
   Call CREACAMPO("", "UM2PTH", "FR4D24PTH35", 7, 0)
   Call CREACAMPO("", "UM2PTH", "FR4D32PTH35", 7, 0)
   
   Call CREACAMPO("", "UM2PTH", "FR4D16PTH70", 7, 0)
   Call CREACAMPO("", "UM2PTH", "FR4D08PTH35", 7, 0)
   Call CREACAMPO("", "UM2PTH", "FR4D10PTH35", 7, 0)
   
   Call CREACAMPO("", "UM2PTH", "FR4D12PTH35", 7, 0)
   Call CREACAMPO("", "UM2PTH", "FR4D04PTH35", 7, 0)
   Call CREACAMPO("", "UM2PTH", "FactorCosto", 7, 0)
   Call CREACAMPO("", "UM2PTH", "Marborra", 3, 0)


   
   
     Call CREACAMPO("", "UM2PTHCOS", "INTERPOLAR", 10, 8)
    If ValidaExisteCampo%("UM2PTHCOS_ID", "UM2PTHCOS") < 1 Then
     SQLX$ = "ALTER TABLE UM2PTHCOS ADD UM2PTHCOS_ID INT IDENTITY"
     FLEXCONN.Execute (SQLX$)
    End If

   Call CREACAMPO("", "UM2PTHCOS", "Lote_M2", 7, 0, 1)
   Call CREACAMPO("", "UM2PTHCOS", "FR4D16PTH18", 7, 0)
   Call CREACAMPO("", "UM2PTHCOS", "FR4D24PTH35", 7, 0)
   Call CREACAMPO("", "UM2PTHCOS", "FR4D32PTH35", 7, 0)
   
   Call CREACAMPO("", "UM2PTHCOS", "FR4D16PTH70", 7, 0)
   Call CREACAMPO("", "UM2PTHCOS", "FR4D08PTH35", 7, 0)
   Call CREACAMPO("", "UM2PTHCOS", "FR4D10PTH35", 7, 0)
   
   Call CREACAMPO("", "UM2PTHCOS", "FR4D12PTH35", 7, 0)
   Call CREACAMPO("", "UM2PTHCOS", "FR4D04PTH35", 7, 0)
   Call CREACAMPO("", "UM2PTHCOS", "FactorCosto", 7, 0)
   Call CREACAMPO("", "UM2PTHCOS", "Marborra", 3, 0)
   '
   'TABLA SECTORES
   Call CREACAMPO("", "SECTORES", "SECTOR", 10, 256)
   If ValidaExisteCampo%("SECTORES_ID", "SECTORES") < 1 Then
     SQLX$ = "ALTER TABLE SECTORES ADD SECTORES_ID INT IDENTITY"
     FLEXCONN.Execute (SQLX$)
   End If
   Call CREACAMPO("", "SECTORES", "ID", 7, 0)
   Call CREACAMPO("", "SECTORES", "SECUENPROD", 3, 0) 'SECUENCIA PRODUCTIVA
   Call CREACAMPO("", "SECTORES", "Marborra", 3, 0)
   
   'TABLA SecuenMaster
   Call CREACAMPO("", "SecuenMaster", "Referencia", 10, 128)
   If ValidaExisteCampo%("SecuenMaster_ID", "SecuenMaster") < 1 Then
     SQLX$ = "ALTER TABLE SecuenMaster ADD SecuenMaster_ID INT IDENTITY"
     FLEXCONN.Execute (SQLX$)
   End If
   Call CREACAMPO("", "SecuenMaster", "CodSecuen", 3, 0, 1)
   Call CREACAMPO("", "SecuenMaster", "StringSecuen", 10, 1024)
   Call CREACAMPO("", "SecuenMaster", "Filtros", 10, 1024) 'SECUENCIA PRODUCTIVA
   Call CREACAMPO("", "SecuenMaster", "CIRela", 4, 0)
   Call CREACAMPO("", "SecuenMaster", "CodExtRela", 10, 24) ' codigo relacionado
   Call CREACAMPO("", "SecuenMaster", "CodCliente", 10, 64) ' codigo relacionado
   Call CREACAMPO("", "SecuenMaster", "Observaciones", 10, 512)
   Call CREACAMPO("", "SecuenMaster", "OrdGuardado", 3, 0, 1) ' ORDEN CON EL CUAL SE GUARDO
   
   Call CREACAMPO("", "SecuenMaster", "Marborra", 3, 0)
   
End Sub
Sub CREATABLAPROVINCIAS()
    Call ABRECONEXION
    Call CREACAMPO("", "PROVINCIA", "ID", 4, 0, 1)
    Call CREACAMPO("", "PROVINCIA", "NOMBRE", 10, 128)
    Call CREACAMPO("", "PROVINCIA", "codigo31662", 10, 64)
    Call CREACAMPO("", "PROVINCIA", "CODFLEX", 10, 24)
    Call CREACAMPO("", "PROVINCIA", "Marborra", 3, 0)
    
    Call CREACAMPO("", "LOCALIDAD", "ID", 4, 0, 1)
    Call CREACAMPO("", "LOCALIDAD", "provincia_id", 4, 0, 1)
    Call CREACAMPO("", "LOCALIDAD", "NOMBRE", 10, 64)
    Call CREACAMPO("", "LOCALIDAD", "codigopostal", 10, 64)
    Call CREACAMPO("", "LOCALIDAD", "Marborra", 3, 0)
    
End Sub
Sub CREATABLAS_CASHFLOW()
    Call ABRECONEXION
    'CREA CABEZERA FLUJO CAJA
    Call CREACAMPO("", "ENCAFLOW", "REFERENCIA1", 10, 256)
    If ValidaExisteCampo%("ENCAFLOW_ID", "ENCAFLOW") < 1 Then
     SQLX$ = "ALTER TABLE ENCAFLOW ADD ENCAFLOW_ID INT IDENTITY"
     FLEXCONN.Execute (SQLX$)
     'VA A TENER RELACION CON ID_CABEZERA  DE TABLA BACKUPFLOW
    End If
    Call CREACAMPO("", "ENCAFLOW", "REFERENCIA2", 10, 256)
    Call CREACAMPO("", "ENCAFLOW", "FECHAYHO", 8, 0)
    Call CREACAMPO("", "ENCAFLOW", "FEDES", 8, 0)
    Call CREACAMPO("", "ENCAFLOW", "FEHAS", 8, 0)
    Call CREACAMPO("", "ENCAFLOW", "INTERVALO", 10, 12)
    '
    'CREA TABLA BACKUPFLOW DETALLE
    Call CREACAMPO("", "BACKUPFLOW", "DOCUORIGEN", 10, 128)
    If ValidaExisteCampo%("BACKUPFLOW_ID", "BACKUPFLOW") < 1 Then
     SQLX$ = "ALTER TABLE BACKUPFLOW ADD BACKUPFLOW_ID INT IDENTITY"
     FLEXCONN.Execute (SQLX$)
    End If
    Call CREACAMPO("", "BACKUPFLOW", "FECHEMISI", 8, 0)
    Call CREACAMPO("", "BACKUPFLOW", "FECHVEN", 8, 0)
    Call CREACAMPO("", "BACKUPFLOW", "ACREDITACION", 8, 0)
    Call CREACAMPO("", "BACKUPFLOW", "NUPROVE", 4, 2, 1)
    Call CREACAMPO("", "BACKUPFLOW", "NUCLIEN", 4, 2, 1)
    Call CREACAMPO("", "BACKUPFLOW", "IMPORTE", 7, 0)
    Call CREACAMPO("", "BACKUPFLOW", "TIPODOCU", 10, 128)
    Call CREACAMPO("", "BACKUPFLOW", "ID_CABEZERA", 6, 0, 1)
    Call CREACAMPO("", "BACKUPFLOW", "Mone_Emis", 3, 2)
    Call CREACAMPO("", "BACKUPFLOW", "COTIZA_DOLAR", 7, 0)
    Call CREACAMPO("", "BACKUPFLOW", "DescriMov", 10, 128)
    Call CREACAMPO("", "BACKUPFLOW", "MOVCASHFLOW", 10, 64)
    Call CREACAMPO("", "BACKUPFLOW", "ORDENAMIENTO", 4, 0)
    Call CREACAMPO("", "BACKUPFLOW", "CodInCue", 3, 0, 1)
    Call CREACAMPO("", "BACKUPFLOW", "CodExCue", 10, 16, 1)
    Call CREACAMPO("", "BACKUPFLOW", "CUCONTA_BANCO", 10, 16, 1)
End Sub
Sub CREATABLARECETACTP()

    Call CREACAMPO("", "RECETACTP", "RASOCLI", 10, 128)
    If ValidaExisteCampo%("RECETACTP_ID", "RECETACTP") < 1 Then
        SQLX$ = "ALTER TABLE RECETACTP ADD RECETACTP_ID INT IDENTITY"
        FLEXCONN.Execute (SQLX$)
    End If
    
    Call CREACAMPO("", "RECETACTP", "NROCOTIZACION", 4, 0, 1)
    Call CREACAMPO("", "RECETACTP", "CODINT", 4, 0, 1)
    Call CREACAMPO("", "RECETACTP", "STATUS", 3, 0, 1)
    Call CREACAMPO("", "RECETACTP", "NUMRECETACTP", 4, 0, 1)
    Call CREACAMPO("", "RECETACTP", "NUCLIEN", 4, 0, 1)
    Call CREACAMPO("", "RECETACTP", "ESPECIFICLI", 10, 512)
    Call CREACAMPO("", "RECETACTP", "CANTIDAD", 6, 4)
    Call CREACAMPO("", "RECETACTP", "JUSTA", 10, 12)
    Call CREACAMPO("", "RECETACTP", "NUVEND", 3, 2)
    Call CREACAMPO("", "RECETACTP", "CAPKGS", 6, 0)
    Call CREACAMPO("", "RECETACTP", "CANTIUSOS", 10, 24)
    Call CREACAMPO("", "RECETACTP", "IMPRESION", 10, 24)
    Call CREACAMPO("", "RECETACTP", "CANTICARAS", 3, 0)
    Call CREACAMPO("", "RECETACTP", "CPOLARGOCM", 6, 0)
    Call CREACAMPO("", "RECETACTP", "CPOANCHOCM", 6, 0)
    Call CREACAMPO("", "RECETACTP", "CPOALTOCM", 6, 0)
    Call CREACAMPO("", "RECETACTP", "ESLOJOCM", 6, 0)
    
    Call CREACAMPO("", "RECETACTP", "ESLMEDI1CM", 6, 0)
    Call CREACAMPO("", "RECETACTP", "ESLMED2CM", 6, 0)
    Call CREACAMPO("", "RECETACTP", "VALVVCDIAMETRO", 6, 0)
    Call CREACAMPO("", "RECETACTP", "VALVVCLARGO", 6, 0)
    Call CREACAMPO("", "RECETACTP", "VALVVCSLARGOSINO", 10, 12)
    Call CREACAMPO("", "RECETACTP", "VALVVCSLARGO", 10, 12)
    Call CREACAMPO("", "RECETACTP", "POLLARGO", 10, 24)
    Call CREACAMPO("", "RECETACTP", "GRMJECUERPO", 10, 48)
    Call CREACAMPO("", "RECETACTP", "GRMJEFONDO", 10, 48)
    

    Call CREACAMPO("", "RECETACTP", "GRMJETAPA", 10, 48)
    Call CREACAMPO("", "RECETACTP", "LAMCPOLAMIN", 10, 48)
    Call CREACAMPO("", "RECETACTP", "LAMFDOLAM", 10, 48)
    Call CREACAMPO("", "RECETACTP", "PDOCTIPO", 10, 48)
    Call CREACAMPO("", "RECETACTP", "PDOCUBIC", 10, 48)
    Call CREACAMPO("", "RECETACTP", "COMPARTIMENTADO", 10, 48)
    Call CREACAMPO("", "RECETACTP", "LINERSINO", 10, 48)

    Call CREACAMPO("", "RECETACTP", "LINERTIPO", 10, 48)
    Call CREACAMPO("", "RECETACTP", "DETAETIQ", 10, 48)
    Call CREACAMPO("", "RECETACTP", "DETAETIQ2", 10, 48)
    Call CREACAMPO("", "RECETACTP", "EMBUFARDO", 10, 48)
    Call CREACAMPO("", "RECETACTP", "TIPOPALLET", 10, 48)
    Call CREACAMPO("", "RECETACTP", "UNXPALLET", 3, 0)
    Call CREACAMPO("", "RECETACTP", "LINER", 6, 0)
    Call CREACAMPO("", "RECETACTP", "CODBIGBAG", 10, 48)
    Call CREACAMPO("", "RECETACTP", "PRODUCTO", 10, 48)
    Call CREACAMPO("", "RECETACTP", "ESPIMPRECTE", 10, 48)
    Call CREACAMPO("", "RECETACTP", "ESPIMIPCTEP", 10, 48)

    Call CREACAMPO("", "RECETACTP", "COLORES", 10, 512)
    Call CREACAMPO("", "RECETACTP", "CANTCOLORES", 3, 0)
    
    Call CREACAMPO("", "RECETACTP", "CPOTIPOCOST", 10, 48)
    Call CREACAMPO("", "RECETACTP", "CPOCOSTPBILO", 10, 48)
    Call CREACAMPO("", "RECETACTP", "CPOCOSTUEXT", 10, 48)

    Call CREACAMPO("", "RECETACTP", "ESLTIPOCOST", 10, 48)
    Call CREACAMPO("", "RECETACTP", "ESLANCHOCM", 10, 48)
    Call CREACAMPO("", "RECETACTP", "ESLPERIMET", 10, 48)
    Call CREACAMPO("", "RECETACTP", "VALVVDDIAM", 6, 0)
    Call CREACAMPO("", "RECETACTP", "VALVVDLARGO", 6, 0)
    Call CREACAMPO("", "RECETACTP", "VALVVDSLARGO", 10, 48)
    Call CREACAMPO("", "RECETACTP", "POLLERALAMIN", 10, 48)
    Call CREACAMPO("", "RECETACTP", "LAMTAPALAM", 10, 48)
    Call CREACAMPO("", "RECETACTP", "LAMTELVALVLAM", 10, 48)
    Call CREACAMPO("", "RECETACTP", "CINTLINER", 10, 48)
    Call CREACAMPO("", "RECETACTP", "PDOCMEDIDA", 10, 48)
    
    Call CREACAMPO("", "RECETACTP", "PDOCCANTIDAD", 6, 0)
    Call CREACAMPO("", "RECETACTP", "COMPARTLAMAMP", 10, 48)
    Call CREACAMPO("", "RECETACTP", "LINERALTO", 10, 48)
    Call CREACAMPO("", "RECETACTP", "LINERMICRON", 6, 0)
    Call CREACAMPO("", "RECETACTP", "OBSGENERAL", 10, 1024)
    Call CREACAMPO("", "RECETACTP", "FechEmisi", 8, 0)
    Call CREACAMPO("", "RECETACTP", "CODIMASTER", 10, 24, 1)

    Call CREACAMPO("", "RECETACTP", "VPDvcsegcm", 10, 48)
    Call CREACAMPO("", "RECETACTP", "VPDCargaCm", 10, 48)
    Call CREACAMPO("", "RECETACTP", "PDFUELLELINERCU", 10, 48)
    Call CREACAMPO("", "RECETACTP", "VPDDISTANCIABOCAFONDOCM", 10, 48)
    Call CREACAMPO("", "RECETACTP", "VPDLARGOMAMAPAROCM", 10, 48)
    Call CREACAMPO("", "RECETACTP", "VPDVDESCSEGCM", 10, 48)
    Call CREACAMPO("", "RECETACTP", "VPDVDESCARGACM", 10, 48)
    Call CREACAMPO("", "RECETACTP", "ESLINGAMEDIDA2CM", 10, 48)
    Call CREACAMPO("", "RECETACTP", "VPDVDLARGO", 10, 48)
    Call CREACAMPO("", "RECETACTP", "VPDANCHOUTILMAMPARO", 10, 48)
    Call CREACAMPO("", "RECETACTP", "VPDANCHOUTILMAMPARO", 10, 48)
    Call CREACAMPO("", "RECETACTP", "VPDDIVISIONLATERALES", 10, 48)
    Call CREACAMPO("", "RECETACTP", "VPDPOLLERA", 10, 48)
    Call CREACAMPO("", "RECETACTP", "VPDTELASVALVLAMIN", 10, 48)
    Call CREACAMPO("", "RECETACTP", "VPDPOLLERAGM2", 10, 48)
    Call CREACAMPO("", "RECETACTP", "VPDVDGM2", 10, 48)
    Call CREACAMPO("", "RECETACTP", "VPDVCGM2", 10, 48)
    Call CREACAMPO("", "RECETACTP", "CPOCOSTUEXT", 10, 48)
    Call CREACAMPO("", "RECETACTP", "TELAVENTILADA", 10, 48)
    Call CREACAMPO("", "RECETACTP", "VPDTELAMAMPARO", 10, 48)
    Call CREACAMPO("", "RECETACTP", "LINERANCHO", 10, 48)
    Call CREACAMPO("", "RECETACTP", "NEWVERSIONDE", 10, 24, 1) 'CUANDO SE COPIA EN BASE A
    Call CREACAMPO("", "RECETACTP", "MICRONAJELINER", 10, 48)
    Call CREACAMPO("", "RECETACTP", "APROBADA", 3, 0, 1) 'CUANDO SE COPIA EN BASE A
    Call CREACAMPO("", "RECETACTP", "COTIZAR", 3, 0, 1) '
    Call CREACAMPO("", "RECETACTP", "PREPARO", 10, 32)
    Call CREACAMPO("", "RECETACTP", "APROBO", 10, 32)
    
    Call CREACAMPO("", "RECETACTP", "FEPREPARO", 8, 0)
    Call CREACAMPO("", "RECETACTP", "FEAPROBO", 8, 0)
    Call CREACAMPO("", "RECETACTP", "FUELLELINERCU", 6, 0)
    Call CREACAMPO("", "RECETACTP", "CODESPTECTE", 10, 64) 'CODIGO DE ESPECIFICACION DEL CLIENTE
    Call CREACAMPO("", "RECETACTP", "NoMostEnList", 3, 0)
    
    Call CREACAMPO("", "RECETACTP", "HILACHA", 10, 64) 'DATOS AGREGADOS EN OTRO FRM PARA LA FICHA TECNICA
    Call CREACAMPO("", "RECETACTP", "DETEMETAL", 10, 64) 'DATOS AGREGADOS EN OTRO FRM PARA LA FICHA TECNICA
    

    'TABLA DE CONFIGURACIN DE BIGBAG PARA BLOQUEAR O NO VALORES EN LA FICHA TECNICA
    Call CREACAMPO("", "CONFIGBB", "CODBIGBAG", 10, 48, 1)
    If ValidaExisteCampo%("CONFIGBB_ID", "CONFIGBB") < 1 Then
        SQLX$ = "ALTER TABLE CONFIGBB ADD CONFIGBB_ID INT IDENTITY"
        FLEXCONN.Execute (SQLX$)
    End If
    Call CREACAMPO("", "CONFIGBB", "DESCRIPCIONBB", 10, 128)
    Call CREACAMPO("", "CONFIGBB", "VC", 10, 16)
    Call CREACAMPO("", "CONFIGBB", "VCS", 10, 16)
    Call CREACAMPO("", "CONFIGBB", "VD", 10, 16)
    Call CREACAMPO("", "CONFIGBB", "VDS", 10, 16)
    Call CREACAMPO("", "CONFIGBB", "POLLERA", 10, 16)
    Call CREACAMPO("", "CONFIGBB", "COMPARTIMENTADO", 10, 16)
    Call CREACAMPO("", "CONFIGBB", "FONDOPANEL", 10, 16)
    Call CREACAMPO("", "CONFIGBB", "CUERPOLAT", 10, 16) 'CUEPO LATERAL
    Call CREACAMPO("", "CONFIGBB", "VALVCARGA", 10, 16) 'VALVULA DE CARGA
    Call CREACAMPO("", "CONFIGBB", "VALVDESCAR", 10, 16) 'VÁLVULA DESCARGA
    Call CREACAMPO("", "CONFIGBB", "CINTIPOLLERA", 10, 16) 'CINTILLA POLLERA
    Call CREACAMPO("", "CONFIGBB", "TAPAPOLLERA", 10, 16) 'TAPA / POLLERA
    Call CREACAMPO("", "CONFIGBB", "CINTDESC", 10, 16) 'CINTILLA DESCARGA
    Call CREACAMPO("", "CONFIGBB", "CINTDESCSEG", 10, 16) 'CINTILLA DESC. SEG
    Call CREACAMPO("", "CONFIGBB", "CINTCARGPOLL", 10, 16) 'CINTILLA CARGA/ POLL
    Call CREACAMPO("", "CONFIGBB", "CINTCARGSEG", 10, 16) 'CINTILLA CARG SEG
    Call CREACAMPO("", "CONFIGBB", "CINTLINER", 10, 16) 'CINTILLA LINER
    Call CREACAMPO("", "CONFIGBB", "MAMPARO", 10, 16)
    Call CREACAMPO("", "CONFIGBB", "HILO1", 10, 16)
    Call CREACAMPO("", "CONFIGBB", "HILO2", 10, 16)
    Call CREACAMPO("", "CONFIGBB", "HILO3", 10, 16)
    
    Call CREACAMPO("", "CONFIGBB", "LATERAL2", 10, 16) 'LATERAL 2
    Call CREACAMPO("", "CONFIGBB", "ESLINGA", 10, 16)
    Call CREACAMPO("", "CONFIGBB", "ESLPERIM", 10, 16) 'ESLINGA PERIMETRAL
    Call CREACAMPO("", "CONFIGBB", "ETIQUETA", 10, 16)
    Call CREACAMPO("", "CONFIGBB", "SOBRE", 10, 16)
    Call CREACAMPO("", "CONFIGBB", "TINTA", 10, 16)
    Call CREACAMPO("", "CONFIGBB", "DILUYENTE", 10, 16)
    Call CREACAMPO("", "CONFIGBB", "LINER", 10, 16)
    Call CREACAMPO("", "CONFIGBB", "ENVOLTORIO", 10, 16)
    Call CREACAMPO("", "CONFIGBB", "UPANEL", 10, 16)
    Call CREACAMPO("", "CONFIGBB", "MARBORRA", 3, 0, 1)
    
    
    
    'TABLA DE COTIZACION
    Call CREACAMPO("", "COTIZACTP", "RASOCLI", 10, 128)
    If ValidaExisteCampo%("COTIZACTP_ID", "COTIZACTP") < 1 Then
        SQLX$ = "ALTER TABLE COTIZACTP ADD COTIZACTP_ID INT IDENTITY"
        FLEXCONN.Execute (SQLX$)
    End If
    
    
    Call CREACAMPO("", "COTIZACTP", "NROCOTIZACION", 4, 0, 1)
    Call CREACAMPO("", "COTIZACTP", "STATUS", 3, 0, 1)
    Call CREACAMPO("", "COTIZACTP", "NUMRECETACTP", 4, 0, 1)
    Call CREACAMPO("", "COTIZACTP", "NUCLIEN", 4, 0, 1)
    Call CREACAMPO("", "COTIZACTP", "FECHEMISI", 8, 0)
    Call CREACAMPO("", "COTIZACTP", "CODIMASTER", 10, 24, 1)
    Call CREACAMPO("", "COTIZACTP", "NoMostEnList", 3, 0)
    Call CREACAMPO("", "COTIZACTP", "NoMostEnListBolas", 3, 0)

    'TABLA COTI1
    Call CREACAMPO("", "COTI1", "Codigo", 10, 24, 1)
    If ValidaExisteCampo%("COTI1_ID", "COTI1") < 1 Then
        SQLX$ = "ALTER TABLE COTI1 ADD COTI1_ID INT IDENTITY"
        FLEXCONN.Execute (SQLX$)
    End If
    
    Call CREACAMPO("", "COTI1", "CODINTELEM", 4, 0, 1)
    Call CREACAMPO("", "COTI1", "CODINTCONJ", 4, 0, 1)
    Call CREACAMPO("", "COTI1", "COMPONENTE", 10, 48)
    Call CREACAMPO("", "COTI1", "STATUS", 3, 0, 1)
    Call CREACAMPO("", "COTI1", "F_TECNICA", 10, 48)
    Call CREACAMPO("", "COTI1", "DESCRIPCION", 10, 64)
    Call CREACAMPO("", "COTI1", "FECHEMISI", 8, 0)
    
    Call CREACAMPO("", "COTI1", "COMPONENTE", 10, 48)
    Call CREACAMPO("", "COTI1", "MODUXBB", 3, 0, 1)
    Call CREACAMPO("", "COTI1", "CORTE", 7, 0)
    Call CREACAMPO("", "COTI1", "PRECIOXM", 7, 0)
    Call CREACAMPO("", "COTI1", "NAMEMSF", 10, 24)
    Call CREACAMPO("", "COTI1", "RESPACORTE", 7, 0)
    
    Call CREACAMPO("", "COTI1", "ORDEN", 3, 0)
    
    Call CREACAMPO("", "COTI1", "NROCOTIZACION", 4, 0, 1)
    Call CREACAMPO("", "COTI1", "MARBORRA", 3, 0)
    Call CREACAMPO("", "COTI1", "CodComponente", 10, 10)
    

    'TABLA COTI2
    Call CREACAMPO("", "COTI2", "SOBRE_PDOCMED", 10, 32)
    If ValidaExisteCampo%("COTI2_ID", "COTI2") < 1 Then
        SQLX$ = "ALTER TABLE COTI2 ADD COTI2_ID INT IDENTITY"
        FLEXCONN.Execute (SQLX$)
    End If
    Call CREACAMPO("", "COTI2", "NROCOTIZACION", 4, 0, 1)
    Call CREACAMPO("", "COTI2", "ETIQ_MODULO", 3, 0)
    Call CREACAMPO("", "COTI2", "MARBORRA", 3, 0)
    Call CREACAMPO("", "COTI2", "ETIQ_COSTO", 7, 0)
    Call CREACAMPO("", "COTI2", "SOBRE_DOCTIPO", 10, 32)
    Call CREACAMPO("", "COTI2", "SOBRE_MODULO", 3, 0)
    Call CREACAMPO("", "COTI2", "SOBRE_COSTO", 7, 0)
    Call CREACAMPO("", "COTI2", "LINER_ANCHO", 7, 0)
    Call CREACAMPO("", "COTI2", "IMPRE_MODULO", 3, 0)
    Call CREACAMPO("", "COTI2", "IMPRE_COSTO", 7, 0)
    Call CREACAMPO("", "COTI2", "LINER_ALTO", 7, 0)
    Call CREACAMPO("", "COTI2", "LINER_COSTO", 7, 0)
    Call CREACAMPO("", "COTI2", "LINER_MODULO", 3, 0)
    Call CREACAMPO("", "COTI2", "LINER_KG", 7, 0)
    Call CREACAMPO("", "COTI2", "ENVO_EMBF", 3, 0)
    Call CREACAMPO("", "COTI2", "LIBRE1TXT", 10, 128)
    Call CREACAMPO("", "COTI2", "ENVO_MOD", 3, 0)
    Call CREACAMPO("", "COTI2", "ENVO_COSTO", 7, 0)
    Call CREACAMPO("", "COTI2", "LIBRE1TXT", 10, 128)
    Call CREACAMPO("", "COTI2", "LIBRE2TXT", 10, 128)
    Call CREACAMPO("", "COTI2", "LIBRE3TXT", 10, 128)
    Call CREACAMPO("", "COTI2", "LIBRE1IMPO", 7, 0)
    Call CREACAMPO("", "COTI2", "LIBRE2IMPO", 7, 0)
    Call CREACAMPO("", "COTI2", "LIBRE3IMPO", 7, 0)
    Call CREACAMPO("", "COTI2", "MICRONALJELINER", 7, 0)
    
    Call CREACAMPO("", "COTI2", "SOBRE_DOCMED", 10, 32)
    Call CREACAMPO("", "COTI2", "MODOPAGO", 10, 128)
    Call CREACAMPO("", "COTI2", "FORMAENTREGA", 10, 1024)
    Call CREACAMPO("", "COTI2", "FECHENTRE", 8, 0)
    Call CREACAMPO("", "COTI2", "FECHEMISI", 8, 0)
    Call CREACAMPO("", "COTI2", "VALIMPRESION", 7, 0)
    
    
    
    
    Call CREACAMPO("", "COTI2", "SCRAPP", 7, 0)
    Call CREACAMPO("", "COTI2", "LOGYFLETE", 7, 0)
    Call CREACAMPO("", "COTI2", "IMPUESTO", 7, 0)
    Call CREACAMPO("", "COTI2", "COMISION", 7, 0)
    Call CREACAMPO("", "COTI2", "OTROS", 7, 0)
    Call CREACAMPO("", "COTI2", "GASTFINAN", 7, 0)
    Call CREACAMPO("", "COTI2", "OTRO1", 7, 0)
    Call CREACAMPO("", "COTI2", "OTRO2", 7, 0)
    Call CREACAMPO("", "COTI2", "ESTRUCTURA", 7, 0)
    
    Call CREACAMPO("", "COTI2", "COSTTOT", 7, 0)
    Call CREACAMPO("", "COTI2", "PJEUTILIDAD", 7, 0)
    Call CREACAMPO("", "COTI2", "PREVENTA", 7, 0)
    Call CREACAMPO("", "COTI2", "PREVENTAREAL", 7, 0)
    
    Call CREACAMPO("", "COTI2", "COSTMOBRA", 7, 0)
    Call CREACAMPO("", "COTI2", "COSTOSFIJOS", 7, 0)
    Call CREACAMPO("", "COTI2", "STATUS", 3, 0, 1)
    Call CREACAMPO("", "COTI2", "OBSERVACIONES", 12, 0)
    Call CREACAMPO("", "COTI2", "CANTICONFIRMADA", 6, 4)
    Call CREACAMPO("", "COTI2", "MANTOFERTA", 10, 64)
    Call CREACAMPO("", "COTI2", "PLAZOENTRE", 10, 64)
    Call CREACAMPO("", "COTI2", "CODESPECIFINT", 10, 64) 'CODIGO DE ESPECIFICACION INTERNA
    Call CREACAMPO("", "COTI2", "PORCENTCOMI", 7, 0)
    Call CREACAMPO("", "COTI2", "IMPOLINER", 7, 0) 'VALOR INGRESADO MANULAMENTE DEL LINER
    Call CREACAMPO("", "COTI2", "PREVENTAFINAL", 7, 0)
    Call CREACAMPO("", "COTI2", "ComPorReal", 7, 0) 'COMISION REAL INGRESADA MANUALMENTE SOLO REFERENCIA SE REPITE EN TABLA COTIPVTA
    
    




    '
    'TABLA CALCUCORTE

    Call CREACAMPO("", "CALCUCORTE", "CODBIGBAG", 10, 48, 1)
    If ValidaExisteCampo%("CALCUCORTE_ID", "CALCUCORTE") < 1 Then
        SQLX$ = "ALTER TABLE CALCUCORTE ADD CALCUCORTE_ID INT IDENTITY"
        FLEXCONN.Execute (SQLX$)
    End If
    Call CREACAMPO("", "CALCUCORTE", "DESCRIPCIONBB", 10, 48)
    Call CREACAMPO("", "CALCUCORTE", "LATERAL2", 10, 48)
    Call CREACAMPO("", "CALCUCORTE", "MARBORRA", 10, 48)
    Call CREACAMPO("", "CALCUCORTE", "CUERPOLAT", 10, 48)
    Call CREACAMPO("", "CALCUCORTE", "FONDOPANEL", 10, 48)
    Call CREACAMPO("", "CALCUCORTE", "UPANEL", 10, 48)
    Call CREACAMPO("", "CALCUCORTE", "POLLERA", 10, 48)
    Call CREACAMPO("", "CALCUCORTE", "TAPAPOLLERA", 10, 48)
    '
    'TABLA DOCURELFT
    Call CREACAMPO("", "DOCURELFT", "NUMRECETACTP", 4, 0, 1)
    Call CREACAMPO("", "DOCURELFT", "CODINT", 4, 0, 1)
    Call CREACAMPO("", "DOCURELFT", "Descripcion", 10, 64)
    Call CREACAMPO("", "DOCURELFT", "Ruta", 10, 512)
    Call CREACAMPO("", "DOCURELFT", "Marborra", 3, 0)

    'TABLA DOCURELFT
    Call CREACAMPO("", "DOCURELPVTA", "NROPVTA", 4, 0, 1)
    Call CREACAMPO("", "DOCURELPVTA", "CODINT", 4, 0, 1)
    Call CREACAMPO("", "DOCURELPVTA", "Descripcion", 10, 24)
    Call CREACAMPO("", "DOCURELPVTA", "Ruta", 10, 512)
    Call CREACAMPO("", "DOCURELPVTA", "Marborra", 3, 0)

    Call CREARVISTAFITECYCOTIZA
    Call CREAR_TABLA_ANOTAMASTER
    Call CREAR_TABLA_CODIGORANGO
    
End Sub
Sub CREAR_TABLA_ANOTAMASTER()
    Call ABRECONEXION
    Call CREACAMPO("", "ANOTAMASTER", "Codint", 4, 0, 2)
    Call CREACAMPO("", "ANOTAMASTER", "HISTORIAL", 12, 0)
    Call CREACAMPO("", "ANOTAMASTER", "Codigo", 10, 24, 1)
    Call CREACAMPO("", "ANOTAMASTER", "Descripcion", 10, 64)
End Sub
Sub CREAR_TABLA_CODIGORANGO()
    Call ABRECONEXION
    Call CREACAMPO("", "CODIGORANGO", "DESCRIRANGO", 12, 0)
    If ValidaExisteCampo%("CODIGORANGO_ID", "CODIGORANGO") < 1 Then
        SQLX$ = "ALTER TABLE CODIGORANGO ADD CODIGORANGO_ID INT IDENTITY"
        FLEXCONN.Execute (SQLX$)
    End If
    Call CREACAMPO("", "CODIGORANGO", "CIRANGO", 4, 0, 1)
    Call CREACAMPO("", "CODIGORANGO", "CODRANGO", 10, 32)
    Call CREACAMPO("", "CODIGORANGO", "Descripcion", 10, 64)
    Call CREACAMPO("", "CODIGORANGO", "CODMATPRIMA", 10, 24)
    Call CREACAMPO("", "CODIGORANGO", "CIMATPRIMA", 4, 0, 1)
    Call CREACAMPO("", "CODIGORANGO", "DESCRIMATPRIMA", 10, 64)
    Call CREACAMPO("", "CODIGORANGO", "Marborra", 3, 0)

End Sub

Sub CREARVISTAFITECYCOTIZA()
        If EXISTEVISTA%("FITECYCOTIZA") < 1 Then
           SQLX$ = "CREATE VIEW FITECYCOTIZA AS"
        Else
           SQLX$ = "ALTER VIEW FITECYCOTIZA AS"
        End If
       SQLX$ = SQLX$ + " SELECT R.NUMRECETACTP AS FT_NROFT,R.FechEmisi AS FT_FECHEMISI"
       SQLX$ = SQLX$ + " ,R.APROBADA AS FT_APROBADA , R.COTIZAR AS FT_COTIZAR, C.NoMostEnList AS COT_NoMostEnList"
       SQLX$ = SQLX$ + " ,C.NROCOTIZACION AS COT_NROCOT,C.FECHEMISI COT_FECHEMISI, R.NUCLIEN AS FT_CTACLTE"
       SQLX$ = SQLX$ + " ,R.RASOCLI AS FT_RASOCLI,R.STATUS AS FT_STAT,C.NUMRECETACTP AS COT_NROFT,R.CODIMASTER AS FT_CODIGO"
       SQLX$ = SQLX$ + " ,R.RECETACTP_ID AS FT_ID, C.NUMRECETACTP AS COT_ID,C.STATUS AS COT_STATUS, R.NUVEND AS FT_NUVEND"
       SQLX$ = SQLX$ + "  FROM COTIZACTP C "
       SQLX$ = SQLX$ + " RIGHT JOIN RECETACTP R ON  R.NUMRECETACTP =  C.NUMRECETACTP"
       FLEXCONN.Execute (SQLX$)

End Sub
Sub CREARTABLASOLICOTI()

    Call CREACAMPO("", "SOLICOTI", "CODBIGBAG", 10, 48, 1)
    If ValidaExisteCampo%("SOLICOTI_ID", "SOLICOTI") < 1 Then
        SQLX$ = "ALTER TABLE SOLICOTI ADD SOLICOTI_ID INT IDENTITY"
        FLEXCONN.Execute (SQLX$)
    End If
    Call CREACAMPO("", "SOLICOTI", "NUVEND", 3, 2)
    Call CREACAMPO("", "SOLICOTI", "CANTIDAD", 6, 4)
    Call CREACAMPO("", "SOLICOTI", "MARBORRA", 3, 0)
    Call CREACAMPO("", "SOLICOTI", "CODBIGBAG", 10, 48)
    Call CREACAMPO("", "SOLICOTI", "CODIMASTER", 10, 24, 1)
    Call CREACAMPO("", "SOLICOTI", "COTIZAR", 3, 0, 1)  '
    Call CREACAMPO("", "SOLICOTI", "NUMRECETACTP", 4, 0, 1)
    Call CREACAMPO("", "SOLICOTI", "CODINT", 4, 0, 1)
    Call CREACAMPO("", "SOLICOTI", "NUCLIEN", 4, 0, 1)
    Call CREACAMPO("", "SOLICOTI", "FECHEMISI", 8, 0)
    Call CREACAMPO("", "SOLICOTI", "NROCOTIZACION", 4, 0, 1)
    Call CREACAMPO("", "SOLICOTI", "DESCRIMASTER", 10, 64)
    Call CREACAMPO("", "SOLICOTI", "RASO_CLI", 10, 64)
    
    Call CREACAMPO("", "SOLICOTIPVTA", "MARBORRA", 3, 0)
    If ValidaExisteCampo%("SOLICOTIPVTA_ID", "SOLICOTIPVTA") < 1 Then
        SQLX$ = "ALTER TABLE SOLICOTIPVTA ADD SOLICOTIPVTA_ID INT IDENTITY"
        FLEXCONN.Execute (SQLX$)
    End If
    Call CREACAMPO("", "SOLICOTIPVTA", "NUVEND", 3, 2)
    Call CREACAMPO("", "SOLICOTIPVTA", "CANTIDAD", 6, 4)
    Call CREACAMPO("", "SOLICOTIPVTA", "CODIMASTER", 10, 24, 1)
    Call CREACAMPO("", "SOLICOTIPVTA", "REFECLIENTE", 10, 64) 'ES UNA REFERENCIA O IDENTIFICADOR O CODIGO DEL PRODUCTO DEL CLIENTE

    Call CREACAMPO("", "SOLICOTIPVTA", "COTIZAR", 3, 0, 1)  '
    Call CREACAMPO("", "SOLICOTIPVTA", "NROPVTA", 4, 0, 1)
    Call CREACAMPO("", "SOLICOTIPVTA", "CODINT", 4, 0, 1)
    Call CREACAMPO("", "SOLICOTIPVTA", "NUCLIEN", 4, 0, 1)
    Call CREACAMPO("", "SOLICOTIPVTA", "FECHEMISI", 8, 0)
    Call CREACAMPO("", "SOLICOTIPVTA", "NROCOTIZACION", 4, 0, 1)
    Call CREACAMPO("", "SOLICOTIPVTA", "DESCRIMASTER", 10, 64)
    Call CREACAMPO("", "SOLICOTIPVTA", "RASO_CLI", 10, 64)
    
    'COTIZACION DE PLANILLA DE VENTAS
    Call CREACAMPO("", "COTIPVTA", "MARBORRA", 3, 0)
    If ValidaExisteCampo%("COTIPVTA_ID", "COTIPVTA") < 1 Then
        SQLX$ = "ALTER TABLE COTIPVTA ADD COTIPVTA_ID INT IDENTITY"
        FLEXCONN.Execute (SQLX$)
    End If
    Call CREACAMPO("", "COTIPVTA", "NUVEND", 3, 2)
    Call CREACAMPO("", "COTIPVTA", "CANTIDAD", 6, 4)
    Call CREACAMPO("", "COTIPVTA", "NROPVTA", 4, 0, 1)
    Call CREACAMPO("", "COTIPVTA", "REFECLIENTE", 10, 64) 'ES UNA REFERENCIA O IDENTIFICADOR O CODIGO DEL PRODUCTO DEL CLIENTE
    Call CREACAMPO("", "COTIPVTA", "NUCLIEN", 4, 0, 1)
    Call CREACAMPO("", "COTIPVTA", "FECHEMISI", 8, 0)
    Call CREACAMPO("", "COTIPVTA", "NROCOTIZACION", 4, 0, 1)
    Call CREACAMPO("", "COTIPVTA", "RASO_CLI", 10, 64)
    
    Call CREACAMPO("", "COTIPVTA", "OBSERVACIONES", 10, 1024)
    Call CREACAMPO("", "COTIPVTA", "GRAMAJE", 7, 0)
    Call CREACAMPO("", "COTIPVTA", "PRECIO_PP", 7, 0)
    Call CREACAMPO("", "COTIPVTA", "PRECIOPAPEL", 7, 0)
    Call CREACAMPO("", "COTIPVTA", "ANCHOPAPEL", 7, 0)
    Call CREACAMPO("", "COTIPVTA", "COSTOIMPRE", 7, 0)
    Call CREACAMPO("", "COTIPVTA", "COSTOCONFE", 7, 0)
    Call CREACAMPO("", "COTIPVTA", "RODILLOS", 7, 0)
    Call CREACAMPO("", "COTIPVTA", "PRECIOVENTA", 7, 0)
    Call CREACAMPO("", "COTIPVTA", "PREVTAREAL", 7, 0)
    Call CREACAMPO("", "COTIPVTA", "PESO", 7, 0)
    Call CREACAMPO("", "COTIPVTA", "GRAMAJEPAPEL", 7, 0)
    Call CREACAMPO("", "COTIPVTA", "PESOPAPEL", 7, 0)
    Call CREACAMPO("", "COTIPVTA", "LARGOCORTE", 7, 0)
    Call CREACAMPO("", "COTIPVTA", "LARGOCORTE2", 7, 0)
    Call CREACAMPO("", "COTIPVTA", "COSTOPAPEL", 7, 0)
    Call CREACAMPO("", "COTIPVTA", "LARGOCAL", 7, 0)
    Call CREACAMPO("", "COTIPVTA", "STATUS", 3, 0)
    Call CREACAMPO("", "COTIPVTA", "ComPorReal", 7, 0) 'COMISION REAL INGRESADA MANUALMENTE SOLO REFERENCIA SE REPITE EN TABLA COTI2
    
    
    Call CREAR_TABLA_FTECNICACAMB
    
End Sub

Sub CREATABLA_UBILOTE_ENVAP()

    Call CREACAMPO("", "UBILOTE_ENVAP", "LOTE", 10, 24, 1)
    If ValidaExisteCampo%("UBILOTE_ENVAP_ID", "UBILOTE_ENVAP") < 1 Then
        SQLX$ = "ALTER TABLE UBILOTE_ENVAP ADD UBILOTE_ENVAP_ID INT IDENTITY"
        FLEXCONN.Execute (SQLX$)
    End If
    Call CREACAMPO("", "UBILOTE_ENVAP", "CANTIDAD", 6, 4)
    Call CREACAMPO("", "UBILOTE_ENVAP", "MARBORRA", 3, 0)
    Call CREACAMPO("", "UBILOTE_ENVAP", "RACK", 3, 0)
    Call CREACAMPO("", "UBILOTE_ENVAP", "DescriRack", 10, 64)
    Call CREACAMPO("", "UBILOTE_ENVAP", "FILA", 3, 0)
    Call CREACAMPO("", "UBILOTE_ENVAP", "COLUMNA", 3, 0)  '
    Call CREACAMPO("", "UBILOTE_ENVAP", "CODINT", 4, 0, 1)
    Call CREACAMPO("", "UBILOTE_ENVAP", "NUCLIEN", 4, 0, 1)
    Call CREACAMPO("", "UBILOTE_ENVAP", "FECHEMISI", 8, 0)
    Call CREACAMPO("", "UBILOTE_ENVAP", "FeSolEnt", 8, 0)
    Call CREACAMPO("", "UBILOTE_ENVAP", "IDSUBOR", 10, 12, 1)
    Call CREACAMPO("", "UBILOTE_ENVAP", "RASO_CLI", 10, 64)
    Call CREACAMPO("", "UBILOTE_ENVAP", "DescOper", 10, 48)
    Call CREACAMPO("", "UBILOTE_ENVAP", "NumeOper", 3, 3)
    Call CREACAMPO("", "UBILOTE_ENVAP", "NROPED", 4, 0, 1)
    Call CREACAMPO("", "UBILOTE_ENVAP", "ID_MOVISTO", 4, 0, 1)
    
End Sub
Sub CREATABLA_P_VTA_BOLSAS()

    'TABLA PLANILLA DE VENTAS BOLSAS CONTEPLAST
    Call CREACAMPO("", "P_VTA_BOLSAS", "RASOCLI", 10, 128)
    If ValidaExisteCampo%("P_VTA_BOLSAS_ID", "P_VTA_BOLSAS") < 1 Then
        SQLX$ = "ALTER TABLE P_VTA_BOLSAS ADD P_VTA_BOLSAS_ID INT IDENTITY"
        FLEXCONN.Execute (SQLX$)
    End If
    '
    Call CREACAMPO("", "P_VTA_BOLSAS", "NROPVTA", 4, 0, 1)
    Call CREACAMPO("", "P_VTA_BOLSAS", "CODINT", 4, 0, 1)
    Call CREACAMPO("", "P_VTA_BOLSAS", "CODIMASTER", 10, 32, 1)
    Call CREACAMPO("", "P_VTA_BOLSAS", "REFECLIENTE", 10, 64)

    Call CREACAMPO("", "P_VTA_BOLSAS", "STATUS", 3, 0, 1)
    Call CREACAMPO("", "P_VTA_BOLSAS", "NUCLIEN", 4, 0, 1)
    Call CREACAMPO("", "P_VTA_BOLSAS", "CANTIDAD", 6, 4)
    Call CREACAMPO("", "P_VTA_BOLSAS", "NUVEND", 3, 2)
    Call CREACAMPO("", "P_VTA_BOLSAS", "FECHEMISI", 8, 0)
    
    Call CREACAMPO("", "P_VTA_BOLSAS", "CAPKGS", 6, 0)
    Call CREACAMPO("", "P_VTA_BOLSAS", "IMPRESION", 10, 24)
    Call CREACAMPO("", "P_VTA_BOLSAS", "MAT_BOLSA", 10, 128)
    Call CREACAMPO("", "P_VTA_BOLSAS", "CANTICARAS", 3, 0)
    Call CREACAMPO("", "P_VTA_BOLSAS", "ALTOUTILCM", 6, 0)
    Call CREACAMPO("", "P_VTA_BOLSAS", "ANCHOTOTCM", 6, 0)
    Call CREACAMPO("", "P_VTA_BOLSAS", "COLORES", 10, 512)
    Call CREACAMPO("", "P_VTA_BOLSAS", "CANTCOLORES", 3, 0)
    Call CREACAMPO("", "P_VTA_BOLSAS", "FUELLE", 10, 24)
    Call CREACAMPO("", "P_VTA_BOLSAS", "SISTCARGA", 10, 24)
    Call CREACAMPO("", "P_VTA_BOLSAS", "GRAMAJE", 10, 48)
    Call CREACAMPO("", "P_VTA_BOLSAS", "LINER", 10, 24)
    Call CREACAMPO("", "P_VTA_BOLSAS", "PRODUCTO", 10, 64)
    Call CREACAMPO("", "P_VTA_BOLSAS", "PLENO", 10, 24)
    Call CREACAMPO("", "P_VTA_BOLSAS", "MEDIFUELLE", 10, 24)
    Call CREACAMPO("", "P_VTA_BOLSAS", "MICROPERFO", 10, 24)
    Call CREACAMPO("", "P_VTA_BOLSAS", "OBSERVACIONES", 10, 1024)
    Call CREACAMPO("", "P_VTA_BOLSAS", "REFECLIENTE", 10, 64)
    
    Call CREACAMPO("", "P_VTA_BOLSAS", "UNIPALLET", 3, 0)
    Call CREACAMPO("", "P_VTA_BOLSAS", "SISDESCA", 10, 48)
    Call CREACAMPO("", "P_VTA_BOLSAS", "SisCarga", 10, 48)
    Call CREACAMPO("", "P_VTA_BOLSAS", "TIPOHILO", 10, 48)
    Call CREACAMPO("", "P_VTA_BOLSAS", "UNIDPAQU", 6, 4)
    
    Call CREACAMPO("", "P_VTA_BOLSAS", "DescarCOLOR", 10, 512) 'COLORES DE SISTEMA DE DESCARGA
    Call CREACAMPO("", "P_VTA_BOLSAS", "DescarCantCol", 3, 0)  'COLORES DE SISTEMA DE DESCARGA
    Call CREACAMPO("", "P_VTA_BOLSAS", "APROBO", 10, 64)
    Call CREACAMPO("", "P_VTA_BOLSAS", "FEAPROBO", 8, 0)
    Call CREACAMPO("", "P_VTA_BOLSAS", "CantHilos", 3, 0)
    Call CREACAMPO("", "P_VTA_BOLSAS", "ObservAdicionales", 10, 1024) ' Se Cargan en la Orden de Fabricacion
    
    Call CREACAMPO("", "P_VTA_BOLSAS", "LARGOCORTE", 7, 0)
    Call CREACAMPO("", "P_VTA_BOLSAS", "EspecifCliente", 10, 128) ' especificacion del cliente
    Call CREACAMPO("", "P_VTA_BOLSAS", "DESCRICODCONFECCION", 10, 128)
    Call CREACAMPO("", "P_VTA_BOLSAS", "CODCONFECCION", 10, 32)
    
    Call CREACAMPO("", "P_VTA_BOLSAS", "EspImpEnva", 10, 128)
    Call CREACAMPO("", "P_VTA_BOLSAS", "HiloMaterial", 10, 64)
    Call CREACAMPO("", "P_VTA_BOLSAS", "TituEnvap", 10, 128)
    Call CREACAMPO("", "P_VTA_BOLSAS", "AnchoCadenta", 10, 16)
    Call CREACAMPO("", "P_VTA_BOLSAS", "Material", 10, 32)
    Call CREACAMPO("", "P_VTA_BOLSAS", "Colocado", 10, 16)
    Call CREACAMPO("", "P_VTA_BOLSAS", "AltoPolietileno", 10, 32)
    Call CREACAMPO("", "P_VTA_BOLSAS", "AnchoPolietileno", 10, 32)
    Call CREACAMPO("", "P_VTA_BOLSAS", "FuellePolietileno", 10, 32)
    Call CREACAMPO("", "P_VTA_BOLSAS", "Pabilo", 10, 32)
    Call CREACAMPO("", "P_VTA_BOLSAS", "CrepeTColor", 10, 32)
    Call CREACAMPO("", "P_VTA_BOLSAS", "TipoColocado", 10, 32)
    
    Call CREACAMPO("", "P_VTA_BOLSAS", "AbreFacil", 10, 32)
    Call CREACAMPO("", "P_VTA_BOLSAS", "Densidad", 10, 32)
    Call CREACAMPO("", "P_VTA_BOLSAS", "TipoCPSP", 10, 32)
    Call CREACAMPO("", "P_VTA_BOLSAS", "MicronajeLiner", 10, 32)
    Call CREACAMPO("", "P_VTA_BOLSAS", "PaqEnfundados", 10, 32)
    Call CREACAMPO("", "P_VTA_BOLSAS", "CertCalidad", 10, 32)
    Call CREACAMPO("", "P_VTA_BOLSAS", "FuelleLaterales", 10, 32)
    
    
    

    
    
    
End Sub
Sub CREAR_TABLA_CREACOD()
    
    Call CREACAMPO("", "CREACOD", "CAMPO", 10, 48)
    If ValidaExisteCampo%("CREACOD_ID", "CREACOD") < 1 Then
        SQLX$ = "ALTER TABLE CREACOD ADD CREACOD_ID INT IDENTITY"
        FLEXCONN.Execute (SQLX$)
    End If
    Call CREACAMPO("", "CREACOD", "CODIGO", 3, 0, 1)
    Call CREACAMPO("", "CREACOD", "FRMNAME", 10, 48)
    Call CREACAMPO("", "CREACOD", "Referencia", 10, 48)
    Call CREACAMPO("", "CREACOD", "T_DATO", 10, 48)
    Call CREACAMPO("", "CREACOD", "LONGITUD", 6, 0)
    Call CREACAMPO("", "CREACOD", "ROTULO", 10, 48)
    Call CREACAMPO("", "CREACOD", "FORMATO", 10, 48)
    Call CREACAMPO("", "CREACOD", "ID_REPLICA", 10, 48)
    Call CREACAMPO("", "CREACOD", "CAMPOECO", 10, 48)
    Call CREACAMPO("", "CREACOD", "SQLXCAMPOECO", 10, 1024)
    Call CREACAMPO("", "CREACOD", "ENCABEZADOSZELECHO", 10, 256)
    
    Call CREACAMPO("", "CREACOD", "FRAME", 10, 48)
    Call CREACAMPO("", "CREACOD", "Graba", 10, 48)
    Call CREACAMPO("", "CREACOD", "MUESTRA_RS", 10, 48)
    Call CREACAMPO("", "CREACOD", "ID01", 10, 48)
    Call CREACAMPO("", "CREACOD", "TABLA", 10, 24)
    Call CREACAMPO("", "CREACOD", "FECHEMISI", 8, 0)
    
    Call CREACAMPO("", "CREACOD", "MARBORRA", 3, 0)
End Sub

Sub CREAR_TABLA_FTECNICACAMB()
     'FICHA TECNICA CAMBACERES
    Call CREACAMPO("", "FTECNICACAMB", "Codigo", 10, 48)
    If ValidaExisteCampo%("FTECNICACAMB_ID", "FTECNICACAMB") < 1 Then
        SQLX$ = "ALTER TABLE FTECNICACAMB ADD FTECNICACAMB_ID INT IDENTITY"
        FLEXCONN.Execute (SQLX$)
    End If
    Call CREACAMPO("", "FTECNICACAMB", "CODINT", 4, 0, 1)
    Call CREACAMPO("", "FTECNICACAMB", "DESCRIP", 10, 64)
    Call CREACAMPO("", "FTECNICACAMB", "ANCHO_CM", 7, 0) 'ANCHO (CM)
    Call CREACAMPO("", "FTECNICACAMB", "GRAMOS_M2", 7, 0) 'GRAMOS M2
    Call CREACAMPO("", "FTECNICACAMB", "ARO_CM", 7, 0) 'ARO (CM)
    Call CREACAMPO("", "FTECNICACAMB", "LAMINADA", 10, 24) ' LAMINADA
    Call CREACAMPO("", "FTECNICACAMB", "DISTDEDOYARO_MM", 7, 0) 'DIST. ENTRE DEDO Y ARO (MM)
    
    Call CREACAMPO("", "FTECNICACAMB", "TramaCta_Cm", 7, 0) 'TramaCta_Cm
    Call CREACAMPO("", "FTECNICACAMB", "TramaAnc_Cm", 7, 0) 'ANCHO DE CINTA (CM)
    Call CREACAMPO("", "FTECNICACAMB", "TramaCob_CmMM", 7, 0) 'COBERTURA X CM (MM)
    Call CREACAMPO("", "FTECNICACAMB", "TramaDenier_GRMT", 7, 0) 'DENIER GR/9000 MTS
    
    Call CREACAMPO("", "FTECNICACAMB", "UrdCta_Cm", 7, 0) 'CINTAS        X CM
    Call CREACAMPO("", "FTECNICACAMB", "UrdAnc_Cm", 7, 0) 'ANCHO DE CINTA (CM)
    Call CREACAMPO("", "FTECNICACAMB", "UrdCob_CmMM", 7, 0) 'COBERTURA X CM (MM)
    Call CREACAMPO("", "FTECNICACAMB", "UrdDenier_GRMT", 7, 0) 'DENIER GR/9000 MTS
    
    Call CREACAMPO("", "FTECNICACAMB", "RefCta_Cm", 7, 0) 'CINTAS        X CM
    Call CREACAMPO("", "FTECNICACAMB", "RefAnc_Cm", 7, 0) 'ANCHO DE CINTA (CM)
    Call CREACAMPO("", "FTECNICACAMB", "RefCob_CmMM", 7, 0) 'COBERTURA X CM (MM)
    Call CREACAMPO("", "FTECNICACAMB", "RefDenier_GRMT", 7, 0) ' DENIER GR/9000 MTS
    
    Call CREACAMPO("", "FTECNICACAMB", "PicsX_MIN", 7, 0) 'PICS X MIN (PPM)
    Call CREACAMPO("", "FTECNICACAMB", "CantCintXMod", 7, 0) 'CANT DE CINTAS X MODULO
    Call CREACAMPO("", "FTECNICACAMB", "CantCintAdicRxM", 7, 0) 'CANT. DE CINTAS ADICIONALES REAL X
    Call CREACAMPO("", "FTECNICACAMB", "LubiriTime", 3, 0) 'tiempo de lubricacion
    Call CREACAMPO("", "FTECNICACAMB", "LubriUnMed", 10, 24) 'medida de tiempo
    Call CREACAMPO("", "FTECNICACAMB", "BreakHoraVal", 7, 0) '
    Call CREACAMPO("", "FTECNICACAMB", "BreakMinVal", 7, 0) '
    Call CREACAMPO("", "FTECNICACAMB", "BreakSegVal", 7, 0) '
    
    
    Call CREACAMPO("", "FTECNICACAMB", "BreakHoraUnMed", 10, 24) '
    Call CREACAMPO("", "FTECNICACAMB", "BreakMinUnMed", 10, 24) '
    Call CREACAMPO("", "FTECNICACAMB", "BreakSegUnMed", 10, 24) '
    Call CREACAMPO("", "FTECNICACAMB", "ResTransAnKG", 7, 0) 'RESISTENCIA  ANCHO (CM)
    Call CREACAMPO("", "FTECNICACAMB", "ResPjeElong", 7, 0) 'RESISTENCIA  % ELONGACION
    
    Call CREACAMPO("", "FTECNICACAMB", "TolAnchoCm", 7, 0) 'TOLERANCIA ANCHO (CM)
    Call CREACAMPO("", "FTECNICACAMB", "TolGramoM2", 7, 0) 'TOLERANCIA GRAMOS M2
    Call CREACAMPO("", "FTECNICACAMB", "BOBINASMT", 3, 0)  'BOBINA (MTS)
    
    Call CREACAMPO("", "FTECNICACAMB", "CintColCant", 3, 0) 'INTAS DE COLOR  cantidad
    Call CREACAMPO("", "FTECNICACAMB", "CintColCol", 10, 48) 'CINTAS DE COLOR  color
    
    Call CREACAMPO("", "FTECNICACAMB", "PosicionCm1", 7, 0) 'PosicionCm
    Call CREACAMPO("", "FTECNICACAMB", "PosicionCm2", 7, 0) 'PosicionCm
    Call CREACAMPO("", "FTECNICACAMB", "PosicionCm3", 7, 0) 'PosicionCm
    Call CREACAMPO("", "FTECNICACAMB", "PosicionCm4", 7, 0) 'PosicionCm
    Call CREACAMPO("", "FTECNICACAMB", "PosicionCm5", 7, 0) 'PosicionCm
    Call CREACAMPO("", "FTECNICACAMB", "PosicionCm6", 7, 0) 'PosicionCm
    Call CREACAMPO("", "FTECNICACAMB", "PosicionCm7", 7, 0) 'PosicionCm
    Call CREACAMPO("", "FTECNICACAMB", "PosicionCm8", 7, 0) 'PosicionCm
    Call CREACAMPO("", "FTECNICACAMB", "PosicionCm9", 7, 0) 'PosicionCm
    Call CREACAMPO("", "FTECNICACAMB", "PosicionCm10", 7, 0) 'PosicionCm

    Call CREACAMPO("", "FTECNICACAMB", "PuntoFreno", 3, 0) 'PosicionCm
    Call CREACAMPO("", "FTECNICACAMB", "Telarespos", 3, 0) 'PosicionCm
    Call CREACAMPO("", "FTECNICACAMB", "TensionarroHilo", 3, 0) 'PosicionCm
    Call CREACAMPO("", "FTECNICACAMB", "PesConVacKg", 7, 0) 'PosicionCm
    Call CREACAMPO("", "FTECNICACAMB", "PesoRolloKg", 7, 0) 'PosicionCm
    
    Call CREACAMPO("", "FTECNICACAMB", "TramaCob_Eficien", 7, 0) 'eficiencia
    Call CREACAMPO("", "FTECNICACAMB", "TramaCob_ValMxCta", 7, 0) 'PosicionCm
    Call CREACAMPO("", "FTECNICACAMB", "MARBORRA", 3, 0)
    
    Call CREACAMPO("", "FTECNICACAMB", "TramaCob_TIPOTELA", 10, 24) 'PosicionCm
    Call CREACAMPO("", "FTECNICACAMB", "TramaCob_GRAMAJE", 7, 0) 'PosicionCm
    '
    
    'TABLA FTMODULO
    Call CREACAMPO("", "FTMODULO", "Codigo", 10, 48)
    If ValidaExisteCampo%("FTMODULO_ID", "FTMODULO") < 1 Then
        SQLX$ = "ALTER TABLE FTMODULO ADD FTMODULO_ID INT IDENTITY"
        FLEXCONN.Execute (SQLX$)
    End If
    Call CREACAMPO("", "FTMODULO", "CODINT", 4, 0, 1)
    Call CREACAMPO("", "FTMODULO", "DESCRIP", 10, 64)
    Call CREACAMPO("", "FTMODULO", "MODULO", 3, 0, 1) '
    Call CREACAMPO("", "FTMODULO", "CANTIDAD", 3, 0) '
    Call CREACAMPO("", "FTMODULO", "REFUERZOS", 3, 0) '
    Call CREACAMPO("", "FTMODULO", "MARBORRA", 3, 0, 1) '
    
End Sub

Sub CREATABLA_FOTODETACIERRE()

            Call CREACAMPO("", "FOTODETACIERRE2", "CodXElem", 10, 24)
            If ValidaExisteCampo%("FOTODETACIERRE2_ID", "FOTODETACIERRE2") < 1 Then
              SQLX$ = "ALTER TABLE FOTODETACIERRE2 ADD FOTODETACIERRE2_ID INT IDENTITY"
              FLEXCONN.Execute (SQLX$)
            End If

            Call CREACAMPO("", "FOTODETACIERRE2", "NumeOper", 3, 2)
            Call CREACAMPO("", "FOTODETACIERRE2", "DescOper", 10, 48)
            Call CREACAMPO("", "FOTODETACIERRE2", "FACTORCORTE", 3, 0)
            Call CREACAMPO("", "FOTODETACIERRE2", "CORTE", 7, 8)
            Call CREACAMPO("", "FOTODETACIERRE2", "CANTCORTES", 7, 8)
            Call CREACAMPO("", "FOTODETACIERRE2", "IdenLote", 10, 32)
            Call CREACAMPO("", "FOTODETACIERRE2", "Peso", 7, 0)
            Call CREACAMPO("", "FOTODETACIERRE2", "CODRANGO", 10, 24)
            Call CREACAMPO("", "FOTODETACIERRE2", "SALDO", 7, 8)
            Call CREACAMPO("", "FOTODETACIERRE2", "IDMOVISTO", 7, 0, 1)

            Call CREACAMPO("", "FOTODETACIERRE3", "CODIGO", 10, 24)
            If ValidaExisteCampo%("FOTODETACIERRE3_ID", "FOTODETACIERRE3") < 1 Then
              SQLX$ = "ALTER TABLE FOTODETACIERRE3 ADD FOTODETACIERRE3_ID INT IDENTITY"
              FLEXCONN.Execute (SQLX$)
            End If
            Call CREACAMPO("", "FOTODETACIERRE3", "DESCRIPCION", 10, 64)
            Call CREACAMPO("", "FOTODETACIERRE3", "COMPONENTE", 10, 48)
            Call CREACAMPO("", "FOTODETACIERRE3", "CODINTELEM", 3, 0)
            Call CREACAMPO("", "FOTODETACIERRE3", "MODUXBB", 3, 0, 1)
            Call CREACAMPO("", "FOTODETACIERRE3", "CORTE", 7, 8)
            Call CREACAMPO("", "FOTODETACIERRE3", "NECESIDAD", 7, 8)
            Call CREACAMPO("", "FOTODETACIERRE3", "NECES_MTS", 7, 8)
            Call CREACAMPO("", "FOTODETACIERRE3", "DISPONIBILIDAD", 7, 8)
            Call CREACAMPO("", "FOTODETACIERRE3", "FALTA", 7, 8)
            Call CREACAMPO("", "FOTODETACIERRE3", "IDMOVISTO", 7, 0, 1)
            Call CREACAMPO("", "FOTODETACIERRE3", "CODRANGO", 10, 24)
            

End Sub

Sub CreaTablaPropCompra()

    Call CREACAMPO("", "CABEZAPropCompra", "Referencia", 10, 256)
    If ValidaExisteCampo%("CABEZAPropCompra_ID", "CABEZAPropCompra") < 1 Then
        SQLX$ = "ALTER TABLE CABEZAPropCompra ADD CABEZAPropCompra_ID INT IDENTITY"
        FLEXCONN.Execute (SQLX$)
    End If
    '
    Call CREACAMPO("", "PropCompra", "Codigo", 10, 24, 1)
    If ValidaExisteCampo%("PropCompra_ID", "PropCompra") < 1 Then
        SQLX$ = "ALTER TABLE PropCompra ADD PropCompra_ID INT IDENTITY"
        FLEXCONN.Execute (SQLX$)
    End If
   '
   Call CREACAMPO("", "PropCompra", "Descripcion", 10, 64)
   Call CREACAMPO("", "PropCompra", "Status", 3, 0)
   Call CREACAMPO("", "PropCompra", "Codint", 4, 0, 2)
   Call CREACAMPO("", "PropCompra", "UNIMED", 10, 16)
   Call CREACAMPO("", "PropCompra", "Necesidad", 7, 0)
   
   Call CREACAMPO("", "PropCompra", "MONEDA", 10, 12)
   Call CREACAMPO("", "PropCompra", "PreUnit", 6, 4)
   Call CREACAMPO("", "PropCompra", "ImpoPUnit", 6, 4)
   Call CREACAMPO("", "PropCompra", "CantSegLIndiv", 7, 0)
   Call CREACAMPO("", "PropCompra", "ImpoSegLIndiv", 6, 4)
   Call CREACAMPO("", "PropCompra", "StockDisp", 7, 0)
   Call CREACAMPO("", "PropCompra", "CantCalc_Stock", 7, 0)
   Call CREACAMPO("", "PropCompra", "ImpoSegStock", 6, 4)
   
   Call CREACAMPO("", "PropCompra", "Nume_Prov", 4, 0, 1)
   Call CREACAMPO("", "PropCompra", "Raso_Prov", 10, 128)

   Call CREACAMPO("", "PropCompra", "CantAComprar", 7, 0)
   Call CREACAMPO("", "PropCompra", "ImpoSegComprar", 6, 4)
   Call CREACAMPO("", "PropCompra", "MARBORRA", 3, 1)
   Call CREACAMPO("", "PropCompra", "Nume_Cli", 4, 0, 1)
   Call CREACAMPO("", "PropCompra", "ID_CABEZA", 4, 0, 1)
   Call CREACAMPO("", "PropCompra", "Raso_Cli", 10, 128)

 End Sub

Function TipoDatoCampo(TABLA$, CAMPO$)
    
    SQLX$ = "Select  DATA_TYPE from information_schema.columns WITH(NOLOCK) "
    SQLX$ = SQLX$ + " WHERE TABLE_NAME= '" & TABLA$ & "'  AND COLUMN_NAME= '" & CAMPO$ & "'"
    Set rs = READSET(SQLX$)
    
    TipoDatoCampo = SAFETEXT$(rs!DATA_TYPE)
    rs.Close
    Set rs = Nothing
End Function

Sub CREARVISTACotizaFtecBolsas()
        If EXISTEVISTA%("CotizaFtecBolsas") < 1 Then
           SQLX$ = "CREATE VIEW CotizaFtecBolsas AS"
        Else
           SQLX$ = "ALTER VIEW CotizaFtecBolsas AS"
        End If
        SQLX$ = SQLX$ + " select COT.NROCOTIZACION AS COT_NROCOTIZACION  ,COT.NUCLIEN AS COT_NUCLIEN"
        SQLX$ = SQLX$ + " ,COT.RASOCLI AS COT_RASOCLI, COT.NUMRECETACTP AS COT_NUMRECETACTP"
        SQLX$ = SQLX$ + " ,P_VTA.NUVEND AS P_VTA_NUVEND"
        SQLX$ = SQLX$ + " ,COT.FECHEMISI AS COT_FECHEMISI , COT.CODIMASTER AS COT_CODIMASTER"
        SQLX$ = SQLX$ + " ,COT.NoMostEnListBolas AS COT_NoMostEnListBolas"
        SQLX$ = SQLX$ + " ,COT.COTIZACTP_ID AS COT_COTIZACTP_ID"
        SQLX$ = SQLX$ + " FROM COTIZACTP COT"
        SQLX$ = SQLX$ + " INNER JOIN P_VTA_BOLSAS P_VTA "
        SQLX$ = SQLX$ + " ON COT.NUMRECETACTP = P_VTA.NroPvta"
        FLEXCONN.Execute (SQLX$)
        
End Sub

Sub CREATABLA_ORDENADOR()

            Call CREACAMPO("", "ORDENADOR", "CodXElem", 10, 24)
            If ValidaExisteCampo%("ORDENADOR_ID", "ORDENADOR") < 1 Then
              SQLX$ = "ALTER TABLE ORDENADOR ADD ORDENADOR_ID INT IDENTITY"
              FLEXCONN.Execute (SQLX$)
            End If

            Call CREACAMPO("", "ORDENADOR", "CAMPONUM1", 7, 0)
            Call CREACAMPO("", "ORDENADOR", "CAMPONUM2", 7, 0)
            Call CREACAMPO("", "ORDENADOR", "CAMPONUM3", 7, 0)
            Call CREACAMPO("", "ORDENADOR", "CAMPONUM4", 7, 0)
            Call CREACAMPO("", "ORDENADOR", "CAMPONUM5", 7, 0)
            Call CREACAMPO("", "ORDENADOR", "CAMPONUM6", 7, 0)
            Call CREACAMPO("", "ORDENADOR", "CAMPONUM7", 7, 0)
            Call CREACAMPO("", "ORDENADOR", "CAMPONUM8", 7, 0)
            Call CREACAMPO("", "ORDENADOR", "CAMPOTEXT1", 10, 64)
            Call CREACAMPO("", "ORDENADOR", "CAMPOTEXT2", 10, 64)
            Call CREACAMPO("", "ORDENADOR", "CAMPOTEXT3", 10, 64)
            Call CREACAMPO("", "ORDENADOR", "CAMPOTEXT4", 10, 64)
            Call CREACAMPO("", "ORDENADOR", "CAMPOTEXT5", 10, 64)
            Call CREACAMPO("", "ORDENADOR", "CAMPOTEXT6", 10, 64)
            Call CREACAMPO("", "ORDENADOR", "CAMPOTEXT7", 10, 64)
            Call CREACAMPO("", "ORDENADOR", "CAMPOTEXT8", 10, 64)
            Call CREACAMPO("", "ORDENADOR", "CAMPOTEXT9", 10, 64)
            Call CREACAMPO("", "ORDENADOR", "CAMPOTEXT10", 10, 64)
            Call CREACAMPO("", "ORDENADOR", "CAMPOTEXT11", 10, 64)
            Call CREACAMPO("", "ORDENADOR", "CAMPOTEXT12", 10, 64)
            Call CREACAMPO("", "ORDENADOR", "CAMPOTEXT13", 10, 64)
            Call CREACAMPO("", "ORDENADOR", "CAMPOTEXT14", 10, 64)
            Call CREACAMPO("", "ORDENADOR", "CAMPOTEXT15", 10, 64)
            Call CREACAMPO("", "ORDENADOR", "CAMPOTEXT16", 10, 64)

            Call CREACAMPO("", "ORDENADOR", "CAMPOFECHA1", 8, 0)
            Call CREACAMPO("", "ORDENADOR", "CAMPOFECHA2", 8, 0)
            Call CREACAMPO("", "ORDENADOR", "CAMPOFECHA3", 8, 0)
            Call CREACAMPO("", "ORDENADOR", "CAMPOFECHA4", 8, 0)

            

End Sub

Sub creaTablaNorfaEnvap()
     Call CREACAMPO("OFABRIC", "NorfaEnvap", "IdSubOr", 10, 12, 1)
     Call CREACAMPO("OFABRIC", "NorfaEnvap", "of_envap", 10, 12)
    
End Sub

Sub GENERA_TABLA_PROPUCOMP()
       Call CREACAMPO("", "ENCA_PROPUCOMP", "REFERENCIA", 10, 64)
       If ValidaExisteCampo%("ENCA_PROPUCOMP_ID", "ENCA_PROPUCOMP") < 1 Then
          SQLX$ = "ALTER TABLE ENCA_PROPUCOMP ADD ENCA_PROPUCOMP_ID INT IDENTITY"
          FLEXCONN.Execute (SQLX$)
        End If

       Call CREACAMPO("", "ENCA_PROPUCOMP", "Codint", 3, 0)
       Call CREACAMPO("", "ENCA_PROPUCOMP", "NumPropCompra", 7, 0)
       Call CREACAMPO("", "ENCA_PROPUCOMP", "FeInsert", 8, 0)
       Call CREACAMPO("", "ENCA_PROPUCOMP", "FeUpdate", 8, 0)
       Call CREACAMPO("", "ENCA_PROPUCOMP", "MARBORRA", 3, 0)
       Call CREACAMPO("", "ENCA_PROPUCOMP", "Cantidad", 7, 0)
       
 
        
        
       Call ABRECONEXION
       Call CREACAMPO("", "DET_PROPUCOMP", "MARBORRA", 3, 0, 1)
       If ValidaExisteCampo%("DET_PROPUCOMP_ID", "DET_PROPUCOMP") < 1 Then
          SQLX$ = "ALTER TABLE DET_PROPUCOMP ADD DET_PROPUCOMP_ID INT IDENTITY"
          FLEXCONN.Execute (SQLX$)
        End If

       Call CREACAMPO("", "DET_PROPUCOMP", "Codint", 3, 0)
       Call CREACAMPO("", "DET_PROPUCOMP", "Necesidad", 7, 0)
       Call CREACAMPO("", "DET_PROPUCOMP", "Indivisible", 7, 0)
       Call CREACAMPO("", "DET_PROPUCOMP", "CompSugerida", 7, 0)
       Call CREACAMPO("", "DET_PROPUCOMP", "Stock", 7, 0)
       Call CREACAMPO("", "DET_PROPUCOMP", "CompAuto", 7, 0)
       Call CREACAMPO("", "DET_PROPUCOMP", "Num_Prov", 6, 0)
       Call CREACAMPO("", "DET_PROPUCOMP", "Nume_Ocom", 4, 4, 1)
       Call CREACAMPO("", "DET_PROPUCOMP", "STATUS", 3, 0)
       Call CREACAMPO("", "DET_PROPUCOMP", "NumPropCompra", 7, 0)
       Call CREACAMPO("", "DET_PROPUCOMP", "OCPendEnt", 7, 0) 'orden de compras pendiente de entrega
       Call CREACAMPO("", "DET_PROPUCOMP", "Stomin", 6, 0)
       Call CREACAMPO("", "DET_PROPUCOMP", "Stomax", 6, 0)
       
       
       
End Sub
