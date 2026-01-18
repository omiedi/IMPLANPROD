Attribute VB_Name = "SQL_MGR"
Public wrkPredeterminado As Workspace, BASDAT, CREABA, CREATA, CREACA

Public MODICONSU%
Public IMPOSEJEC%
'
Public DESQUERY$
Public TIPQUERY%
Public DATBASE$
Public TABSELE$, TABSECU$, CAMPOMA1$, CAMPOMA2$, TIPOJOIN$, UNIONALL$
Public COLUMINI%, COLUMAXI%, COLUPROM%, COLUTOTA%, SALACUMU%
'
Public CANFI%
Public SELCAMPO$(128)
Public FORCAMPO$(128)
Public FORCOLUM$(256)
Public TOTCAMPO$(128)
Public GRUCAMPO$(128)
Public OCUCAMPO$(128)
Public ORDCAMPO$(128)
Public TITCAMPO$(128)
Public TABLAECO$(128)
Public FORMAECO$(128)
Public TABVALO$(128)
Public CAMVALO$(128)
Public CONVALO$(128)
'
Public CACONDI%
Public CAMPOCOND$(128), CAMPOSEL$(128)
Public OPERACOND$(128)
Public VAREFCOND$(128)
Public VALORCOND$(128)

Public TXT_CONSULTA$
Public TXT_CAREGIST$
'
Dim YACRECON%                         ' BANDERA PARA INDICAR QUE YA SE HA CREADO LA TABLA DE CONTROL
'^ SIMBOLO POTENCIA
Sub APLICAR_LOTES_A_RESERVAS(CI1%, CODI$, CANTIX, LOTE$, NORFA$, NSERIE$)
           '
           'VALIDO SI YA EXISTEN RESERVAS PASADAS DE TRAZAXSERIE A RESERVA
           'SI DETECTA QUE SI BORRA LOS REGISTROS
           'SI DETECTA QUE SI NO ACTUALIZA LA CANTIDAD ORIGINAL DE CANTRES
           
           ACTUALIZAR_CANTIDAD_RESERVADA_ORIGINAL% = 0
           CONDI$ = "IDSUBOR = '" & NORFA & "' AND COD_EXTE = '" & CODI$ & "' AND NUMEROSERIE = '" & NSERIE$ & "'"
           If CantRegistros&("RESERVA", CONDI$, "NOMESS") > 0 Then
              Call BORRAREGISTROS("RESERVA", CONDI$, REGAF&, "NOMESS")
              ACTUALIZAR_CANTIDAD_RESERVADA_ORIGINAL% = 1
           End If
           '
           If ACTUALIZAR_CANTIDAD_RESERVADA_ORIGINAL% > 0 Then
                'SE DESCUENTA DE LA CANTIDAD RESERVADA ORIGINAL DE CANTRES LA FILA QUE SE VA AGREGAR POR LA CANTIDAD AGREGADA
                'ES DECIR SI CANTRES = 5 Y SE VA A GENERAR UNA FILA NUEVA POR 1 CON SU LOTE THEN CANTRES = 4
                'SIEMPRE DEBERIA HABER UNA SOLA FILA PARA UN CODIGO, UNA OF Y QUE NO TENGA N.SERIE
                Dim rs1 As ADODB.Recordset
                Set rs1 = New ADODB.Recordset
                SQLX$ = "SELECT * FROM RESERVA WHERE IDSUBOR = '" & NORFA & "' AND COD_EXTE = '" & CODI$ & "' AND (NUMEROSERIE = '' OR NUMEROSERIE IS NULL)"
                rs1.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
                With rs1
                  Do While Not .EOF
                    !CANTRES = XVALO(!CANTRES) - XVALO(CANTIX)
                    USOIX = XVALO(!UsoUnit)
                    If XVALO(!CANTRES) <= 0 Then .Delete
                    .Update
                    .MoveNext
                  Loop
                End With
                rs1.Close: Set rs1 = Nothing
           End If
           '
           'AGREGA LA NUEVA FILA CON LOTE.
           Dim RS2 As ADODB.Recordset
           Set RS2 = New ADODB.Recordset
           SQLX$ = "SELECT * FROM RESERVA WHERE IDSUBOR = ''"
           RS2.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
           With RS2
                .AddNew
                !CANTRES = XVALO(CANTIX)
                !NUMEROSERIE = NSERIE$
                On Error Resume Next
                !CODIEMPR = CodiEmp%
                On Error GoTo 0
                !cod_inte = CI1%
                !Cod_Exte = CODI$
                !Descrip = Left$(DESCRIT0$(CI1%), 64)
                !idsubor = Left$(NORFA$, 12)
                !FechRes = Now
                !CANTCONS = 0
                !Saldo_Entre = CANTIX
                !idenlote = Left$(LOTE$, 16)
                !DEPOSITO = 1
                !UsoUnit = USOIX
                !NuOrItem = 0
                !NUME_SOLINT = 0 ' INICIALIZACION DE CAMPO
               .Update
           End With
           RS2.Close: Set RS2 = Nothing
'           Call ACSALOTE(CI1%, LOTE$, "NOMESS")

 
End Sub


Sub CARGA_CONSULTA(CODQUERY$)
    '
    ENCONTROSQL% = 0
    SQLX$ = "SELECT * FROM SQL_QUERIES WITH(NOLOCK)"
    SQLX$ = SQLX$ + " WHERE COD_QUERY = '" & CODQUERY$ & "'"
    Set RS = READSET(SQLX$)
    With RS
      If .EOF Then
        ENCONTROSQL% = -1
      Else
        ENCONTROSQL% = 1
        GoTo 50
      End If
    End With
    If ENCONTROSQL% < 0 Then
      'SI NO LO ENCONTRO DESTRUYE EL RECORDSET
      RS.Close
      Set RS = Nothing
    End If
    
    Call ABRETACON
    strSQL = "SELECT * FROM SQL_QUERIES"
    strSQL = strSQL + " WHERE COD_QUERY = '" & CODQUERY$ & "'"
    On Error Resume Next
    'Set MIRS = FLEXCONN.Execute(STRSQL)
    Set MIRS = BaDaConsul.OpenRecordset(strSQL, dbOpenSnapshot)
    MIRS.MoveFirst
    If Err Then
       On Error GoTo 0
       Call MENSERR(24, "No se ha Encontrado la Consulta Predefinida '" + CODQUERY$ + "'.")
       If APLINVO$ <> "" Then
          Exit Sub    ' Call FINAL("")
       End If
      If DERACCE%("CFGREPS", "Configuración de Reportes") > 1 Then
          FGENREP07.Show 1
       End If
       Exit Sub
    End If
    '
50  CANFI% = 0: CACONDI% = 0
    If ENCONTROSQL% > 0 Then
        If Not (RS.EOF And RS.BOF) Then
           DESQUERY$ = RS!DES_QUERY
           TIPQUERY% = XVALO(RS!TIP_QUERY)
           TTXX$ = RS!PAR_QUERY
           Call DECOD_CONSULTA(TTXX$)
        End If
        RS.Close
        Set RS = Nothing
        Exit Sub
    End If
    
    If Not (MIRS.EOF And MIRS.BOF) Then
       DESQUERY$ = MIRS!DES_QUERY
       TIPQUERY% = XVALO(MIRS!TIP_QUERY)
       TTXX$ = MIRS!PAR_QUERY
       Call DECOD_CONSULTA(TTXX$)
    End If
    '
End Sub

Sub DECOD_CONSULTA(TTYY$)
    '
    TTXX$ = TTYY$
    CANFI% = 0: CACONDI% = 0
    TABSECU$ = "": CAMPOMA1$ = "": CAMPOMA2$ = "": TIPOJOIN$ = ""
    COLUMINI% = 0: COLUMAXI% = 0: COLUPROM% = 0: COLUTOTA% = 0: SALACUMU% = 0
    '
    While Len(TTXX$) > 0
       PPXX% = InStr(TTXX$, Chr$(124))
       If PPXX% < 1 Then PPXX% = 1 + Len(TTXX$)
       COMAI$ = Left$(TTXX$, PPXX% - 1)
       TTXX$ = Mid$(TTXX$, PPXX% + 1)
       '
       If Left$(COMAI$, 10) = "DATABASE =" Then
           DATBASE$ = UCase$(TRIM$(Mid$(COMAI$, 11)))
           GoTo 19
         ElseIf Left$(COMAI$, 10) = "SELTABLE =" Then
           TABSELE$ = UCase$(TRIM$(Mid$(COMAI$, 11)))
         ElseIf Left$(COMAI$, 10) = "SELTASEC =" Then
           TABSECU$ = UCase$(TRIM$(Mid$(COMAI$, 11)))
         ElseIf Left$(COMAI$, 10) = "TIPOJOIN =" Then
           TIPOJOIN$ = UCase$(TRIM$(Mid$(COMAI$, 11)))
         ElseIf Left$(COMAI$, 10) = "UNIONALL =" Then
           UNIONALL$ = UCase$(TRIM$(Mid$(COMAI$, 11)))
         ElseIf Left$(COMAI$, 10) = "MATCHFI1 =" Then
           CAMPOMA1$ = UCase$(TRIM$(Mid$(COMAI$, 11, 64)))
         ElseIf Left$(COMAI$, 10) = "MATCHFI2 =" Then
           CAMPOMA2$ = UCase$(TRIM$(Mid$(COMAI$, 11, 64)))
         ElseIf Left$(COMAI$, 10) = "SELFIELD =" Then
           CANFI% = CANFI% + 1
           FISELX$ = TRIM$(Mid$(COMAI$, 11))
           SELCAMPO$(CANFI%) = UCase$(TRIM$(Mid$(FISELX$, 1, 64)))
           FORCAMPO$(CANFI%) = UCase$(TRIM$(Mid$(FISELX$, 65, 8)))
           OCUCAMPO$(CANFI%) = TRIM$(Mid$(FISELX$, 73, 1))
           GRUCAMPO$(CANFI%) = TRIM$(Mid$(FISELX$, 75, 1))
           TOTCAMPO$(CANFI%) = TRIM$(Mid$(FISELX$, 76, 1))
           ORDCAMPO$(CANFI%) = UCase$(TRIM$(Mid$(FISELX$, 77, 4)))
           TITCAMPO$(CANFI%) = TRIM$(Mid$(FISELX$, 81, 64))
           TABLAECO$(CANFI%) = TRIM$(Mid$(FISELX$, 145, 16))
           FORMAECO$(CANFI%) = TRIM$(Mid$(FISELX$, 161, 16))
           '
           TABVALO$(CANFI%) = ""
           CAMVALO$(CANFI%) = ""
'           CONVALO$(CANFI%) = ""
           TAUXI$ = Mid$(FISELX$, 169, 640)
           CONVALO$(CANFI%) = TRIM$(Mid$(TAUXI$, 129, 512))
           If UCase$(TRIM$(Left$(SELCAMPO$(CANFI%), 8))) = "VALOBADA" Then
              TABVALO$(CANFI%) = TRIM$(Mid$(TAUXI$, 1, 64))
              CAMVALO$(CANFI%) = TRIM$(Mid$(TAUXI$, 65, 64))
'              CONVALO$(CANFI%) = TRIM$(Mid$(TAUXI$, 129, 512))
           End If
           '
         ElseIf Left$(COMAI$, 10) = "CAMFECHA =" Then
           NUMFI% = 1
           If CANFI% < NUMFI% Then CANFI% = NUMFI%
           FISELX$ = TRIM$(Mid$(COMAI$, 11))
           SELCAMPO$(NUMFI%) = UCase$(TRIM$(Mid$(FISELX$, 1, 64)))
           FORCAMPO$(NUMFI%) = UCase$(TRIM$(Mid$(FISELX$, 65, 10)))
           GRUCAMPO$(NUMFI%) = TRIM$(Mid$(FISELX$, 75, 1))
           TOTCAMPO$(NUMFI%) = TRIM$(Mid$(FISELX$, 76, 1))
           ORDCAMPO$(NUMFI%) = UCase$(TRIM$(Mid$(FISELX$, 77, 4)))
           TITCAMPO$(NUMFI%) = TRIM$(Mid$(FISELX$, 81, 64))
           TABLAECO$(NUMFI%) = TRIM$(Mid$(FISELX$, 145, 16))
           FORMAECO$(NUMFI%) = TRIM$(Mid$(FISELX$, 161, 16))
         ElseIf Left$(COMAI$, 10) = "CAMCLASE =" Then
           NUMFI% = 2
           If CANFI% < NUMFI% Then CANFI% = NUMFI%
           FISELX$ = TRIM$(Mid$(COMAI$, 11))
           SELCAMPO$(NUMFI%) = UCase$(TRIM$(Mid$(FISELX$, 1, 64)))
           FORCAMPO$(NUMFI%) = UCase$(TRIM$(Mid$(FISELX$, 65, 10)))
           GRUCAMPO$(NUMFI%) = TRIM$(Mid$(FISELX$, 75, 1))
           TOTCAMPO$(NUMFI%) = TRIM$(Mid$(FISELX$, 76, 1))
           ORDCAMPO$(NUMFI%) = UCase$(TRIM$(Mid$(FISELX$, 77, 4)))
           TITCAMPO$(NUMFI%) = TRIM$(Mid$(FISELX$, 81, 64))
           TABLAECO$(NUMFI%) = TRIM$(Mid$(FISELX$, 145, 16))
           FORMAECO$(NUMFI%) = TRIM$(Mid$(FISELX$, 161, 16))
         ElseIf Left$(COMAI$, 10) = "CAMVALOR =" Then
           NUMFI% = 3
           If CANFI% < NUMFI% Then CANFI% = NUMFI%
           FISELX$ = Mid$(COMAI$, 12)
           SELCAMPO$(NUMFI%) = UCase$(TRIM$(Mid$(FISELX$, 1, 64)))
           FORCAMPO$(NUMFI%) = UCase$(TRIM$(Mid$(FISELX$, 65, 10)))
           GRUCAMPO$(NUMFI%) = TRIM$(Mid$(FISELX$, 75, 1))
           TOTCAMPO$(NUMFI%) = TRIM$(Mid$(FISELX$, 76, 1))
           ORDCAMPO$(NUMFI%) = UCase$(TRIM$(Mid$(FISELX$, 77, 4)))
           TITCAMPO$(NUMFI%) = TRIM$(Mid$(FISELX$, 81, 64))
           TABLAECO$(NUMFI%) = TRIM$(Mid$(FISELX$, 145, 16))
           FORMAECO$(NUMFI%) = TRIM$(Mid$(FISELX$, 161, 16))
         ElseIf Left$(COMAI$, 10) = "CAMVALO2 =" Then
           NUMFI% = 4
           If CANFI% < NUMFI% Then CANFI% = NUMFI%
           FISELX$ = Mid$(COMAI$, 12)
           SELCAMPO$(NUMFI%) = UCase$(TRIM$(Mid$(FISELX$, 1, 64)))
           FORCAMPO$(NUMFI%) = UCase$(TRIM$(Mid$(FISELX$, 65, 10)))
           GRUCAMPO$(NUMFI%) = TRIM$(Mid$(FISELX$, 75, 1))
           TOTCAMPO$(NUMFI%) = TRIM$(Mid$(FISELX$, 76, 1))
           ORDCAMPO$(NUMFI%) = UCase$(TRIM$(Mid$(FISELX$, 77, 4)))
           TITCAMPO$(NUMFI%) = TRIM$(Mid$(FISELX$, 81, 64))
           TABLAECO$(NUMFI%) = TRIM$(Mid$(FISELX$, 145, 16))
           FORMAECO$(NUMFI%) = TRIM$(Mid$(FISELX$, 161, 16))
         ElseIf Left$(COMAI$, 10) = "CONDITIO =" Then
           CACONDI% = CACONDI% + 1
           CDSELX$ = TRIM$(Mid$(COMAI$, 11))
           CAMPOCOND$(CACONDI%) = TRIM$(Left$(CDSELX$, 24))
           OPERACOND$(CACONDI%) = TRIM$(Mid$(CDSELX$, 26, 8))
           VAREFCOND$(CACONDI%) = TRIM$(Mid$(CDSELX$, 35))
         ElseIf Left$(COMAI$, 10) = "COL_MIN = " Then
           If TRIM$(Mid$(COMAI$, 11, 2)) = "SI" Then
              COLUMINI% = 1
           End If
         ElseIf Left$(COMAI$, 10) = "COL_MAX = " Then
           If TRIM$(Mid$(COMAI$, 11, 2)) = "SI" Then
              COLUMAXI% = 1
           End If
         ElseIf Left$(COMAI$, 10) = "COL_PRO = " Then
           If TRIM$(Mid$(COMAI$, 11, 2)) = "SI" Then
              COLUPROM% = 1
           End If
         ElseIf Left$(COMAI$, 10) = "COL_TOT = " Then
           If TRIM$(Mid$(COMAI$, 11, 2)) = "SI" Then
              COLUTOTA% = 1
           End If
         ElseIf Left$(COMAI$, 10) = "FOR_VAL = " Then
           If TRIM$(Mid$(COMAI$, 11, 8)) = "SALFINAL" Then
              SALACUMU% = 1
           End If
         Else
       End If
19  Wend
    '
End Sub
'
Sub AUDITACT(TABLA$, CAMPO$, CONDI$, VALORANT, NUEVALOR)
    ' Crea Tabla y Campos para Auditoria de Tablas
'    Static ATACT%
'    If ATACT% = 0 Then
'        Call CREACAMPO("", "AUDITACT", "AuTabla", 10, 64)
'        Call CREACAMPO("", "AUDITACT", "AuCampo", 10, 64)
'        Call CREACAMPO("", "AUDITACT", "AuCondi", 12, 0)
'        Call CREACAMPO("", "AUDITACT", "AuValAnt", 11, 0)
'        Call CREACAMPO("", "AUDITACT", "AuNueVa", 11, 0)
'        Call CREACAMPO("", "AUDITACT", "AuFeHor", 8, 0)
'        Call CREACAMPO("", "AUDITACT", "AuUser", 3, 0, 1)
'        Call CREACAMPO("", "AUDITACT", "AuUsuario", 10, 64)
'        Call CREACAMPO("", "AUDITACT", "AuAplica", 10, 32)
'        Call CREACAMPO("", "AUDITACT", "AuTerminal", 10, 64)
'        ATACT% = 1
'    End If
    If CStr(VALORANT) <> CStr(NUEVALOR) Then
        If UCase(Left$(TABLA$, 3)) <> "SQL" Then
            User% = USNBR% Mod 100
            Dim AUDITACT As ADODB.Recordset
            Set AUDITACT = New ADODB.Recordset
            SQLX$ = "SELECT * FROM AUDITACT WHERE AuUser = 0 "
25          On Error Resume Next
            AUDITACT.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
            If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
                On Error GoTo 0
                Call CapturaErrorOpen
                GoTo 25
            End If
            On Error GoTo 0
            With AUDITACT
                .AddNew
                On Error Resume Next
                !Autabla = TABLA$                  ' nombre tabla
                !AuCampo = CAMPO$                  ' nombre campo
                !AuCondi = CONDI$                  ' condicion
                !AuValAnt = VALORANT               ' valor anterior
                !AuNueva = NUEVALOR                ' valor nuevo
                !AuFeHor = Now                     ' fecha actual
                !AuUser = User%                    ' numero del usuario
                !AuUsuario = Left$(USERNAME$, 64)  ' nombre del usuario activo
                !AuAplica = Left$(App.EXEName, 32) ' nombre de la aplicacion
                !AuTerminal = Left$(NOMTER$, 64)   ' nombre del terminal
                !CODIEMPR = CodiEmp%               ' codigo de empresa
                .Update
                If Err Then
                   On Error GoTo 0
                   TTXX$ = "Error al Grabar Auditoria " & TABLA$ & "   " & CAMPO$ & "  " & CONDI$ & "   " & VALORANT & "   " & NUEVALOR
                   MsgBox TTXX$
                End If
                On Error GoTo 0
            End With
            AUDITACT.Close
        End If
    End If
End Sub

Sub ARMA_CONSULTA(Optional VUELTANUMERO%)
    '
    KLIK% = 0
    If VUELTANUMERO% > 0 And CONSIDERAR_MOVIVIE% > 0 Then TABSELE$ = "MOVIVIE"

    SQLX$ = "SELECT ": SQLZ$ = "": SQLT$ = ""
    VUELTA& = 0: VUELTB& = 0: VUELTC& = 0: AGRUPA% = 0
    For i& = 1 To CANFI%
       CAMPO$ = SELCAMPO$(i&)
       If Left$(UCase$(CAMPO$), 8) = "VALOBADA" Then
          Call CREACAMPO("", TABSELE$, CAMPO$, 10, 256)
          CAMPO$ = TABSELE$ + "." + CAMPO$
       End If
       ALIASCAMPO$ = CAMPO$
       If GRUCAMPO$(i&) = "1" Then
          CAMPO1$ = CAMPO$
          If InStr(CAMPO1$, ".") > 0 Then
              PPXYZ% = InStr(CAMPO1$, ".")
              CAMPO1$ = Mid$(CAMPO1$, PPXYZ% + 1)
          End If
          ALIASCAMPO$ = "Sx" + CAMPO1$
          CAMPO$ = "SUM(" + CAMPO$ + ") AS " + ALIASCAMPO$
          AGRUPA% = 1
       End If
       If TRIM$(CAMPO$) <> "" Then
          If VUELTA& = 0 Then
              SQLX$ = SQLX$ + CAMPO$
             Else
              SQLX$ = SQLX$ + ", " + CAMPO$
          End If
          VUELTA& = VUELTA& + 1
       End If
       '
       ORDEX$ = TRIM$(UCase$(ORDCAMPO$(i&)))
       If ORDEX$ = "ASC" Or ORDEX$ = "DESC" Then
          If VUELTB& > o Then SQLZ$ = SQLZ$ + ", "
          SQLZ$ = SQLZ$ + ALIASCAMPO$ + " " + ORDEX$
          VUELTB& = VUELTB& + 1
       End If
       '
       AGRUX$ = GRUCAMPO$(i&)
       If AGRUX$ <> "1" Then
          If VUELTC& > o Then SQLT$ = SQLT$ + ", "
          SQLT$ = SQLT$ + CAMPO$ + " "
          VUELTC& = VUELTC& + 1
       End If
    Next i&
    If VUELTA& = 0 Then SQLX$ = "SELECT *"
    '
    TAPRIN$ = TABSELE$
    SQLY$ = " FROM " + TAPRIN$
    '
    TASEC$ = TABSECU$
    If TASEC$ <> "" Then
       CAMPRIN$ = CAMPOMA1$
       If CAMPRIN$ <> "" Then
          CAMSEC$ = CAMPOMA2$
          If CAMSEC$ <> "" Then
             If TIPOJOIN$ = "" Then TIPOJOIN$ = "INNER"
             SQLY$ = SQLY$ + " " & TIPOJOIN$ & " JOIN " + TASEC$
             SQLY$ = SQLY$ + " WITH(NOLOCK) ON " + TAPRIN$ + "." + CAMPRIN$
             SQLY$ = SQLY$ + " = " + TASEC$ + "." + CAMSEC$
          End If
       End If
    End If
    '
    If TASEC$ = "" And CAMPOMA2$ = "" Then SQLY$ = SQLY$ + " WITH(NOLOCK) "
    '
    If CACONDI% > 0 Then
       jj& = 0
       For i& = 1 To CACONDI%
'          If JJ& < 1 Then
'                SQLY$ = SQLY$ + " WHERE "
'             Else
'                SQLY$ = SQLY$ + " AND "
'          End If
          'INTERVENCION OR
          If jj& < 1 Then
              SQLY$ = SQLY$ + " WHERE "
          ElseIf UCase$(Left$(TRIM$(CAMPOCOND$(i&)), 3)) = "OR " Then
              SQLY$ = SQLY$ + " "
          Else
               SQLY$ = SQLY$ + " AND "
          End If

          SQLY$ = SQLY$ + CAMPOCOND$(i&) + " "
          SQLY$ = SQLY$ + OPERACOND$(i&) + " "
          SQLY$ = SQLY$ + VAREFCOND$(i&)
          jj& = jj& + 1
       Next i&
    End If
    '
    TXT_CONSULTA$ = SQLX$ + SQLY$
       If AGRUPA% = 1 Then TXT_CONSULTA = TXT_CONSULTA$ + " GROUP BY " + SQLT$
       If SQLZ$ <> "" Then TXT_CONSULTA = TXT_CONSULTA$ + " ORDER BY " + SQLZ$
       
    If UNIONALL$ <> "" Then
        POSI% = 1
        PrimeraTabla$ = TRIM$(TAPRIN$)
        SegundaTabla$ = TRIM$(UNIONALL$)
        ConsultaSql$ = TXT_CONSULTA$
85      POSI% = InStr(POSI%, ConsultaSql$, PrimeraTabla$)
        If POSI% > 0 Then
            ConsultaSql$ = Mid$(ConsultaSql$, 1, POSI% - 1) & SegundaTabla$ & Mid$(ConsultaSql$, POSI% + Len(PrimeraTabla$))
            POSI% = POSI% + Len(SegundaTabla$) + 1: GoTo 85
        End If
        TXT_CONSULTA$ = SQLX$ + SQLY$ & " UNION ALL " & ConsultaSql$
        
        PrimeraConsulta$ = "SELECT COUNT(*) AS CANREGIS" + SQLY$
        SegundaConsulta$ = PrimeraConsulta$
        POSI% = InStr(1, SQLY$, PrimeraTabla$)
        SegundaConsulta$ = "SELECT COUNT(*) AS CANREGIS" + Mid$(SQLY$, 1, POSI% - 1) & " " & SegundaTabla$ & " " & Mid$(SQLY$, POSI% + Len(PrimeraTabla))
        TXT_CAREGIST$ = "SELECT (" & PrimeraConsulta & ") + (" & SegundaConsulta$ & ")"
    Else
        TXT_CAREGIST$ = "SELECT COUNT(*) AS CANREGIS" + SQLY$
    End If
    '
End Sub
'
Sub CARGADEBASE(ADAT$, TBXX$, OKX%)
    '
    TBXX$ = "": OKX% = 0
    ADAT1$ = TRIM$(ADAT$)
    If ADAT1$ = "" Then Exit Sub
    '
    SQLX$ = "SELECT * FROM SQL_QUERIES WITH(NOLOCK)"
    SQLX$ = SQLX$ + " WHERE COD_QUERY = '" & ADAT1$ & "'"
    Set RS = READSET(SQLX$)
    If Not RS.EOF Then
       CANFI% = 0: CACONDI% = 0
       TTXX$ = RS!PAR_QUERY
       GoTo 40
    End If
    
    Call ABRETACON
    strSQL = "SELECT * FROM SQL_QUERIES"
    strSQL = strSQL + " WHERE COD_QUERY = '" & ADAT1$ & "'"
    On Error Resume Next
    Set MIRS = BaDaConsul.OpenRecordset(strSQL, dbOpenSnapshot)
    MIRS.MoveFirst
    If Err Then
       On Error GoTo 0
       Exit Sub           ' CONSULTA NO DEFINIDA
    End If
    On Error GoTo 0
    If (MIRS.EOF And MIRS.BOF) Then Exit Sub   ' CONSULTA NO DEFINIDA
    '
    CANFI% = 0: CACONDI% = 0
    TTXX$ = MIRS!PAR_QUERY
    
40  Call DECOD_CONSULTA(TTXX$)
    Call ARMA_CONSULTA
    '
    Screen.MousePointer = 11
    Call ABRECONEXION
    On Error Resume Next
    Set RS = FLEXCONN.Execute(FILTSQL$(TXT_CAREGIST$))
    If Err Then Exit Sub        'Imposible Ejecutar la Consulta
    On Error GoTo 0
    FIN& = RS!CANREGIS
    '
        MsgBox "VA A EJECUTAR LA CONSULTA"

    Set RS = New ADODB.Recordset
    On Error Resume Next
    Set RS = FLEXCONN.Execute(FILTSQL$(TXT_CONSULTA))
    If Err Then Exit Sub
    '
    ' GENERA EL TEXTO
    AX1% = 0: AX2% = 0
    For i& = 1 To RS.Fields.Count
      If FORCAMPO$(i&) = "" Then
         FORCAMPO$(i&) = "A16": ACAMPO% = 16
         TICAMPO% = RS.Fields(i& - 1).Type
         If TICAMPO% = 2 Then FORCAMPO$(i&) = "F7.0": ACAMPO% = 7                       ' INTEGER
         If TICAMPO% = 3 Then FORCAMPO$(i&) = "F7.0": ACAMPO% = 7                       ' INTEGER
         If TICAMPO% = 5 Then FORCAMPO$(i&) = "F12.2": ACAMPO% = 12                      ' NUMERICO
         If TICAMPO% = 202 Then FORCAMPO$(i&) = "A48": ACAMPO% = 48                        ' TEXTO
         If TICAMPO% = 135 Then FORCAMPO$(i&) = "D8": ACAMPO% = 8     ' FECHA
      End If
      ACAMPO% = XVALO(Mid$(FORCAMPO$(i&), 2))
      '
      If i& = 1 Then
           AX1% = ACAMPO%
         Else
           AX2% = AX2% + ACAMPO%
      End If
    Next i&
    '
    AXX% = AX1% + 2 + AX2%
    TBXX$ = Space$(FIN& * AXX%)
    '
    RWW& = 0
    With RS
       Do While Not .EOF
          For i& = 1 To .Fields.Count
            On Error Resume Next
            valor = .Fields(i& - 1).Value
            If i& = 1 Then
                 TXTX$ = FORMATNUM$(valor, FORCAMPO$(i&)) + "  "
               Else
                 TXTX$ = TXTX$ + FORMATNUM$(valor, FORCAMPO$(i&))
            End If
          Next i&
          Mid$(TBXX$, RWW& * AXX% + 1, AXX%) = TXTX$
          RWW& = RWW& + 1
       .MoveNext
       Loop
    End With
    '
    TBXX$ = Left$(TBXX$, RWW& * AXX%)
    OKX% = 1
    '
End Sub

Function EXISCONSUSQL%(CODMO$)
        EXISCONSUSQL% = 0
        SQLX$ = "SELECT * FROM SQL_QUERIES WITH(NOLOCK) WHERE COD_QUERY = '" & CODMO$ & "'"
        Set RS = READSET(SQLX$)
        With RS
          If Not .EOF Then EXISCONSUSQL% = 1
        End With
        RS.Close
        Set RS = Nothing
 End Function


Sub GrabarMensaje(SUCURSAL, NUOC&, MENSAXE$, USUARIOX%)

    SQLX$ = "Select * from MENSA_SUCURSAL"
    SQLX$ = SQLX$ + " WHERE O_COMPRA_SUC = " & NUOC&
    SQLX$ = SQLX$ + " AND SUCURSAL = " & SUCURSAL
    '
    Dim RS As ADODB.Recordset
    Set RS = New ADODB.Recordset
    '
    RS.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
    '
    With RS
       If .EOF Then
         .AddNew
       End If
       !SUCURSAL = SUCURSAL
       !O_COMPRA_SUC = NUOC&
       !USUARIO = USUARIOX% 'EL DE LA SUCURSAL QUE ENVIO
       !Status = 0
       !Mensaje = MENSAXE$
      .Update
    End With
    RS.Close
    Set RS = Nothing
    '
End Sub

Function IDCONSULTABLAIMPRE()
'   SQLX$ = "DROP TABLE TABLA_IMPRE": FLEXCONN.Execute (SQLX$)
   
   Call GENERATABLAIMPRE
   MAXTIMPRE = 1 + XVALO(VALOBADA("TABLA_IMPRE", "MAX(TABLA_IMPRE_ID)", "TABLA_IMPRE_ID > 0", "NOMESS"))
   SQLX$ = "INSERT INTO TABLA_IMPRE (IDCONSUL,CAMPO1) "
   SQLX$ = SQLX$ + " Values ("
   SQLX$ = SQLX$ + "" & MAXTIMPRE
   SQLX$ = SQLX$ + ",'BORRAR' "
   SQLX$ = SQLX$ + ")"
   Dim RS As New ADODB.Recordset
   FLEXCONN.Execute (SQLX$)
   If Len(SAFETEXT(MAXTIMPRE)) > 8 Then MAXTIMPRE = XVALO(Right(SAFETEXT(MAXTIMPRE), 8))
   IDCONSULTABLAIMPRE = MAXTIMPRE
   
End Function

Function ListaValoresInsert$(Origen$, ListaCampos$, CONDI$)
'asignar a textolote la lista de campos
'OJO DEVUELVE EL STRING CON LOS PARENTESIS
   Call TEXTOLOTES$(ListaCampos$, ",")
   CANTTEXLOT = UBound(CADENATEX$)

   acumulaStringValores$ = ""
   SQLX$ = "  SELECT " & ListaCampos$
   SQLX$ = SQLX$ + " FROM  " & Origen$ & " WITH(NOLOCK)"
   SQLX$ = SQLX$ + " WHERE " & CONDI$
   Set RS = READSET(SQLX$)
   With RS
       Do While Not .EOF
         If TRIM$(acumulaStringValores$) <> "" Then 'es por que ya ingreso una vez en el for, separa el string de valores x coma
              acumulaStringValores$ = acumulaStringValores$ & "," 'devolveria algo asi:(val1,val2),(val1,val2),(val1,val2)
         End If
         For i& = 1 To CANTTEXLOT
              A = CADENATEX$(i&)
            If i& = 1 Then
               acumulaStringValores$ = acumulaStringValores$ & "("
            End If
            CampoName$ = RS.Fields(i& - 1).Name
            TICAMPO% = RS.Fields(i& - 1).Type
             If TICAMPO% = 2 Or TICAMPO% = 3 Or TICAMPO% = 5 Then
                  VALORCAMPO = TRIM$(FORMATNUM$(RS.Fields(i& - 1).Value, "F16.5"))
                  If Right(VALORCAMPO, 6) = ".00000" Then VALORCAMPO = SAFETEXT$(XVALO(VALORCAMPO))
                  If XVALO(VALORCAMPO) < 0.00005 Then
                     VALORCAMPO = 0
                  End If
               ElseIf TICAMPO% = 135 Then
                  FE% = FOR_DDMMYYYY%(SAFETEXT$(RS.Fields(i& - 1).Value))
                  VALORCAMPO = Chr(39) & FETEXCOR1$(FE%) & Chr(39)               ' FECHA
               Else
                  VALORCAMPO = Chr(39) & SAFETEXT$(RS.Fields(i& - 1).Value) & Chr(39)
             End If
             If i& = 1 Then
                acumulaStringValores$ = acumulaStringValores$ & VALORCAMPO
             Else
                acumulaStringValores$ = acumulaStringValores$ & "," & VALORCAMPO
             End If
         Next i&
         acumulaStringValores$ = acumulaStringValores$ & ")"
         i& = 1 ' por si pega mas de una vuelta y debe generar un script de insert masivo
         .MoveNext
       Loop
   End With
   RS.Close
   Set RS = Nothing
   ListaValoresInsert$ = acumulaStringValores$

End Function

Function VALIDARCAMPOFECHA%(CAMPOSAGRABAR$)
       'VALIDA SI ALGUNO DE LOS CAMPOS A GRABAR EN LA TABLA DATASQL ES CAMPOFE1 O CAMPOFE2 QUE SON DE FECHAS
       'SI ENCUENTRA ESTOS CAMPOS VALIDA QUE LA FECHA SEA VALIDA, SI ES VALIDA DE TODOS MODOS LA FORMATEA A DD/MM/AA
       'O SEA QUE SE PUEDE INGRESAR CON OTRO FORMATO PERO LO CAMBIA AL FORMATO DD/MM/AA
       VALIDARCAMPOFECHA% = 0
       Dim strArray() As String
       On Error Resume Next
       strArray = Split(CAMPOSAGRABAR$, ",")
       For i& = 0 To UBound(strArray)
           If VALIDARCAMPOFECHA% < 0 Then Exit For
           NOMB_CAMPO$ = UCase(TRIM$(strArray(i&)))
           If NOMB_CAMPO$ = "CAMPOFE1" Or NOMB_CAMPO$ = "CAMPOFE2" Then
              For J& = 1 To VENTABNEW.MSF.Rows - 1
                 fech$ = VENTABNEW.MSF.TextMatrix(J&, i& + 1)
                 If IsDate(fech$) Then
                    FE% = FECHANUM(fech$) 'FUNCIONA CON DD/MM/AA
                    If FE% < 1 Then
                       FE% = CVINT(fech$) ' POR SI VIENE CON OTRO FORMATO DD/MM/AAAA
                    End If
                    If FE% < 1 Then GoTo 50
                    VENTABNEW.MSF.TextMatrix(J&, i& + 1) = FECHATEX$(FE%)
                 Else
50                  Call COMUNI("Atención En Fila: " & J& & " La Fecha Ingresada No es Válida")
                    VALIDARCAMPOFECHA% = -1
                    Exit For
                 End If
              Next J&
           End If
       Next i&
       


End Function

Function VALOBADA(NOTAB$, NOCAM$, CONDI$, Optional MODOMESS$)
   'PASANDOLE COMO ARGUMENTO EL NOMBRE DEL CAMPO , EL DE LA TABLA Y LA CONDICION _
   DEVUELVE EN VALOBADA EL VALOR DE LA ENCONTRADO SEGUN LOS ARGUMENTOS ENVIADOS

   Call ABRECONEXION                ' PARA EVITAR QUE SE CAIGA POR ESTAR CERRADA - 9/7/2010
   '
   CONDI3$ = CONDI$
   PPXX% = InStr(CONDI$, "['")
   If PPXX% > 0 Then
      PPYY% = InStr(PPXX%, CONDI$, "']")
      If PPYY% > PPXX% Then
        Code$ = TRIM$(Mid$(CONDI$, PPXX% + 2, PPYY% - PPXX% - 2))
        CIII& = CODINT%(Code$)
        CONDI3$ = Left$(CONDI$, PPXX% - 1) & CIII& & Mid$(CONDI$, PPYY% + 2)
      End If
   End If
   '
   ' SI NO SE DETALLA EL CAMPO O LA CONDICION SALE SIN MENSAJE. EPB 6/10/2011
   'If NOCAM$ = "" Or CONDI3$ = "" Then Exit Function   ' NUEVO
   '
   SQLX$ = "SELECT " + NOCAM$ + " AS VALORSE "
   
   If TRIM$(NOTAB$) = "" Then GoTo 35   ' NUEVO
   
   SQLX$ = SQLX$ + " FROM " + NOTAB$
   If InStr(UCase$(NOTAB$), "INNER JOIN") < 1 Then
      If InStr(UCase$(NOTAB$), "WITH(NOLOCK)") < 1 Then
         SQLX$ = SQLX$ + " WITH(NOLOCK)" ' AGREGADO EL WITH(NOLOCK) - 9-7-2010
      End If
   End If
   '

35 If TRIM$(CONDI3$) = "" Then GoTo 45  ' NUEVO

   SQLX$ = SQLX$ + " WHERE " + CONDI3$

45 Dim MIRS As ADODB.Recordset
   Set MIRS = New ADODB.Recordset
   '
   VALOBADA = ""
   MIRS.CursorLocation = adUseClient
   MIRS.CursorType = adOpenForwardOnly
   MIRS.LockType = adLockReadOnly
   On Error Resume Next
   'Set MIRS = FLEXCONN.Execute(FILTSQL$(SQLX$, MODOMESS$))
   MIRS.Open FILTSQL$(SQLX$), FLEXCONN, , , adCmdText
   If Err Then
     On Error GoTo 0
     If MODOMESS$ <> "NOMESS" Then
        TTXX$ = "Imposible Ejecutar Consulta\   \'" + SQLX$ + "'"
        If COMALTER%(TTXX$ + "\\Continuar\Interrumpir") = 2 Then
           Call FINAL("")
        End If
      End If
     GoTo 90
   End If
   On Error GoTo 0
   With MIRS
     If Not (.BOF And .EOF) Then
       VTP% = VarType(!VALORSE)
       If VTP% = 8 Or VTP% < 2 Then
           VALOBADA = SAFETEXT(!VALORSE)
         Else
           VALOBADA = !VALORSE
       End If
     End If
   End With
   '
90 On Error Resume Next
   MIRS.Close
   Set MIRS = Nothing
   On Error GoTo 0
   '
End Function

Function CantRegistros&(NOTAB$, Optional CONDI$, Optional MODOMESS$, Optional TIMUT, Optional CAMPODISTINCT$)
   'Cuenta la cantidad de registros de una tabla , la condicion es opcional
    '
    CantRegistros& = 0
    '
'    If XVALO(TIMUT) > 0 Then
'      TIMUTAN = FLEXCONN.CommandTimeout
'      FLEXCONN.CommandTimeout = TIMUT
'    End If
    '
    SQLX$ = "SELECT COUNT(*) AS CANTIDAD"
    If CAMPODISTINCT$ <> "" Then SQLX$ = "SELECT COUNT(DISTINCT " & CAMPODISTINCT$ & ") AS CANTIDAD"
    SQLX$ = SQLX$ + " FROM " & NOTAB$
    If InStr(UCase$(NOTAB), "JOIN") < 1 Then
       SQLX$ = SQLX$ + " WITH(NOLOCK) "
    End If
    If CONDI$ <> "" Then
      SQLX$ = SQLX$ + " WHERE  " & CONDI$
    End If
    '
    On Error Resume Next
    Set MIRS = FLEXCONN.Execute(FILTSQL$(SQLX$, MODOMESS$))
    If Err Then
      ERRNU = Err.Number
      ERRDE$ = Err.Description
      On Error GoTo 0
      TTZZ$ = "Error " & ERRNU & " - " & ERRDE$
      Call FRATEXTO(TTZZ$, 96)
      TTXX$ = ""
      For KKI% = 1 To CARETEX%
        TTXX$ = TTXX$ + "\" + RETEX$(KKI%)
      Next KKI%
      Call COMUNI(TTXX$)
      Exit Function
    End If
    On Error GoTo 0
    '
    With MIRS
     If Not (.EOF And .EOF) Then
       CantRegistros& = XVALO(!Cantidad)
     End If
    End With
    '
End Function

Function EXISTELATABLA%(NOTAB$)
    ' devuelve '0 - cero' si no existe ... '1 - uno' si existe NOTAB$
    ' creada por EPB el 7/11/2010
    '
    EXISS% = 0
    If TRIM$(NOTAB$) <> "" Then
       SQLX$ = "SELECT COUNT(*) AS CANTIDAD FROM " & NOTAB$
       On Error Resume Next
       Set MIRS = FLEXCONN.Execute(FILTSQL$(SQLX$, "NOMESS"))
       If Err Then
          MIRS.Close
          On Error GoTo 0
          GoTo 90
       End If
       On Error GoTo 0
       MIRS.Close
       EXISS% = 1
    End If
    '
90  EXISTELATABLA% = EXISS%
    '
End Function

Sub ACTUREGISTRO_NUE(TABLA$, CAMPO$, CONDI$, NUEVALOR, REGISTROSAFECTADOS&, Optional MODO$, Optional TIMUT)
   '
'   If XVALO(TIMUT) > 0 Then
'      TIMUTAN = FLEXCONN.CommandTimeout
'      FLEXCONN.CommandTimeout = TIMUT
'   End If
   '
   NUMREG& = CantRegistros&(TABLA$, CONDI$, MODO$)
   If NUMREG& > 0 Then
      Openforms = DoEvents
      ' Valor Anterior para Tabla de Auditoria
      If NUMREG& = 1 Then
        VALORANT = VALOBADA(TABLA$, CAMPO$, CONDI$, "NOMESS")
        Else
        VALORANT = "Varios" 'Si es mas de un registro, escribe "Varios"
      End If
      '
      TPC = TIPODECAMPO(TABLA$, CAMPO$)
      If TPC = 99 Or TPC = 34 Or TPC = 173 Then
        Dim MIRS As ADODB.Recordset
        Set MIRS = New ADODB.Recordset
        SQLX$ = "SELECT " & CAMPO$ & " AS NUCAM FROM " & TABLA & " WHERE " & CONDI$
25      On Error Resume Next
        MIRS.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
        If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
           On Error GoTo 0
           Call CapturaErrorOpen
           GoTo 25
        End If
        On Error GoTo 0
        RETRATA% = 0
        '
28      With MIRS
          If Not (.BOF And .EOF) Then
            On Error Resume Next
            !NUCAM = NUEVALOR
            .Update
            If Err Then
               ERRNU = Err.Number
               ERRDE$ = Err.Description
               On Error GoTo 0
               '
               If InStr(UCase$(MODO$), "RETRY") > 0 Then
                  RETRATA% = RETRATA% + 1
                  If RETRATA% < 8 Then GoTo 28
               End If
               '
               LONGI& = Len(CStr(NUEVALOR))
               LOMAX& = LONCAMPO(TABLA$, CAMPO$)
               TTXX$ = "ERROR No Recuperable en Transacción\de Actualización (ACTUREGISTRO)"
               TTXX$ = TTXX$ + "\al Procesar Campo: '" + CAMPO$ + "' (Long: " & LONGI& & " - Max: " & LOMAX& & ") en Tabla: '" + TABLA$ + "'."
               TTXX$ = TTXX$ + "\   \El Error Informado por el Servidor SQL es:"
                  TTZZ$ = "Error " & ERRNU & " - " & ERRDE$
                  Call FRATEXTO(TTZZ$, 96)
                  For KKI% = 1 To CARETEX%
                     TTXX$ = TTXX$ + "\" + RETEX$(KKI%)
                  Next KKI%
               TTXX$ = TTXX$ + "\   \Probable Saturación de la Capacidad\de la Base de Datos."
               TTXX$ = TTXX$ + "\   \Verifique Incremento Automatico del Tamaño (10 %)"
               TTXX$ = TTXX$ + "\y Crecimiento Ilimitado (Unrestricted Growth),"
               TTXX$ = TTXX$ + "\tanto en el Archivo de la B.D. Principal"
               TTXX$ = TTXX$ + "\como Tambien en el Historico de Transacciones (Transaction Log) !!!"
               TTXX$ = TTXX$ + "\   \Consulte Mesa de Ayuda FLEXOFT."
               If (USNBR% Mod 100) <> 1 Then
                    Call COMUNI(TTXX$)
                    Call FINAL("")
                  Else
                    OPXX% = COMALTER%(TTXX$ + "\\Re-Intentar Transaccion\Continuar (Resultados Impredecibles !!!)\Cancelar Proceso")
                    If OPXX% <= 1 Then GoTo 28
                    If OPXX% = 2 Then GoTo 35
                    Call FINAL("")
               End If
            End If
35          On Error GoTo 0
          End If
        End With
        '
        MIRS.Close
        Set MIRS = Nothing
        GoTo 99
      End If
      '
      VALOCO$ = SAFETEXT$(NUEVALOR)
      VALOCO$ = REPLACAR$(VALOCO$, "'", " ")
      '
      ' ESTO QUE SIGUE CORREGIDO POR EPB EL 16/5/2009
      If TPC = 231 Then  ' ES UN CAMPO TEXTO
           LONGI = LONCAMPO(TABLA$, CAMPO$)
           VALOCO$ = Left$(VALOCO$, LONGI)
        ElseIf TPC = 62 Or TPC = 56 Then
           VALOCO$ = TRIM$(FORMATNUM$(NUEVALOR, "F20.8"))
           If VALOCO$ = "" Then VALOCO$ = "0"
           If Right$(VALOCO$, 9) = ".00000000" Then VALOCO$ = Left$(VALOCO$, Len(VALOCO$) - 9)
      End If
      '
      If TPC = 58 Then
         If IsNumeric(NUEVALOR) = True Then
              FEXI% = NUEVALOR
           Else
              FFXX$ = NUEVALOR
              FEXI% = FECHANUM(FFXX$)
         End If
         If FEXI% > 0 Then
            VALOCO$ = FETEXCOR1$(FEXI%)
         End If
      End If
      '
'COMENTO PARA QUE NO AGREGUE LA COMILLA SIMPLE
'      If TPC = 58 Or TPC = 231 Or TPC = 99 Then ' FECHA - TEXTO - MEMO
'         VALOCO$ = "'" + VALOCO$ + "'"
'      End If
      '
'      SQLX$ = "UPDATE " + TABLA$
'      SQLX$ = SQLX$ + " SET " + CAMPO$ + " = " & VALOCO$
'      SQLX$ = SQLX$ & " WHERE " + CONDI$
      '
      
        Dim MIRS1 As ADODB.Recordset
        Set MIRS1 = New ADODB.Recordset
        SQLX$ = "SELECT " & CAMPO$ & " AS NUCAM FROM " & TABLA & " WHERE " & CONDI$
30      On Error Resume Next
        MIRS1.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
        If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
           On Error GoTo 0
           Call CapturaErrorOpen
           GoTo 30
        End If
        On Error GoTo 0
        '
        With MIRS1
            On Error Resume Next
            Do While Not .EOF
               'ESTA GRABANDO BIEN PERO LOS NUMEROS CON . LO TOMA COMO UN VALOR ENTERO
               !NUCAM = VALOCO$
               .Update
               If Err Then GoSub ERROR_GRABACION
               .MoveNext
            Loop
            On Error GoTo 0
         End With
         MIRS1.Close
         Set MIRS1 = Nothing
      
      If InStr(UCase$(MODO$), "NOAUDIT") < 1 Then
         Call AUDITACT(TABLA$, CAMPO$, CONDI$, VALORANT, NUEVALOR)
      End If
   End If
99 REGISTROSAFECTADOS& = NUMREG&
      
Exit Sub
      
      
      
ERROR_GRABACION:
      
      If Err Then
         ERRNU = Err.Number
         ERRDE$ = Err.Description
         On Error GoTo 0
         If InStr(UCase$(MODO$), "NOMESS") < 1 Then
           TTYY$ = " Imposible Ejecutar Actualizacion:\  "
           TTXX$ = "'" + SQLX$ + "'"
           Call FRATEXTO(TTXX$, 96)
           For KKI% = 1 To CARETEX%
              TTYY$ = TTYY$ + "\" + RETEX$(KKI%)
           Next KKI%
           TTYY$ = TTYY$ + "\   "
           TTYY$ = TTYY$ + "\El Error Informado por el Servidor SQL es:"
           TTXX$ = "Error " & ERRNU & " - " & ERRDE$
           Call FRATEXTO(TTXX$, 96)
           For KKI% = 1 To CARETEX%
              TTYY$ = TTYY$ + "\" + RETEX$(KKI%)
           Next KKI%
           TTYY$ = TTYY$ + "\   \Comunique Este Mensaje a Mesa de Ayuda FLEXOFT"
           Call MENSERR(24, TTYY$)
         End If
      End If
   '
'   If XVALO(TIMUT) > 0 Then
'      If XVALO(TIMUTAN) > 0 Then
'         FLEXCONN.CommandTimeout = TIMUTAN
'      End If
'   End If
   '
End Sub

Sub ACTUREGISTRO(TABLA$, CAMPO$, CONDI$, NUEVALOR, REGISTROSAFECTADOS&, Optional MODO$, Optional TIMUT)
   '
'   If XVALO(TIMUT) > 0 Then
'      TIMUTAN = FLEXCONN.CommandTimeout
'      FLEXCONN.CommandTimeout = TIMUT
'   End If
   '
   Call CREATABLA_BCONTROL 'CREA LOS CAMPOS DE BASE AUDITACT
   '
   NUMREG& = CantRegistros&(TABLA$, CONDI$, MODO$)
   If NUMREG& > 0 Then
      Openforms = DoEvents
      ' Valor Anterior para Tabla de Auditoria
      If NUMREG& = 1 Then
        VALORANT = VALOBADA(TABLA$, CAMPO$, CONDI$, "NOMESS")
        Else
        VALORANT = "Varios" 'Si es mas de un registro, escribe "Varios"
      End If
      '
      TPC = TIPODECAMPO(TABLA$, CAMPO$)
      If TPC = 99 Or TPC = 34 Or TPC = 173 Then
        Dim MIRS As ADODB.Recordset
        Set MIRS = New ADODB.Recordset
        SQLX$ = "SELECT " & CAMPO$ & " AS NUCAM FROM " & TABLA & " WHERE " & CONDI$
25      On Error Resume Next
        MIRS.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
        If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
           On Error GoTo 0
           Call CapturaErrorOpen
           GoTo 25
        End If
        On Error GoTo 0
        RETRATA% = 0
        '
28      With MIRS
          If Not (.BOF And .EOF) Then
            On Error Resume Next
            !NUCAM = NUEVALOR
            .Update
            If Err Then
               ERRNU = Err.Number
               ERRDE$ = Err.Description
               On Error GoTo 0
               '
               If InStr(UCase$(MODO$), "RETRY") > 0 Then
                  RETRATA% = RETRATA% + 1
                  If RETRATA% < 8 Then GoTo 28
               End If
               '
               LONGI& = Len(CStr(NUEVALOR))
               LOMAX& = LONCAMPO(TABLA$, CAMPO$)
               TTXX$ = "ERROR No Recuperable en Transacción\de Actualización (ACTUREGISTRO)"
               TTXX$ = TTXX$ + "\al Procesar Campo: '" + CAMPO$ + "' (Long: " & LONGI& & " - Max: " & LOMAX& & ") en Tabla: '" + TABLA$ + "'."
               TTXX$ = TTXX$ + "\   \El Error Informado por el Servidor SQL es:"
                  TTZZ$ = "Error " & ERRNU & " - " & ERRDE$
                  Call FRATEXTO(TTZZ$, 96)
                  For KKI% = 1 To CARETEX%
                     TTXX$ = TTXX$ + "\" + RETEX$(KKI%)
                  Next KKI%
               TTXX$ = TTXX$ + "\   \Probable Saturación de la Capacidad\de la Base de Datos."
               TTXX$ = TTXX$ + "\   \Verifique Incremento Automatico del Tamaño (10 %)"
               TTXX$ = TTXX$ + "\y Crecimiento Ilimitado (Unrestricted Growth),"
               TTXX$ = TTXX$ + "\tanto en el Archivo de la B.D. Principal"
               TTXX$ = TTXX$ + "\como Tambien en el Historico de Transacciones (Transaction Log) !!!"
               TTXX$ = TTXX$ + "\   \Consulte Mesa de Ayuda FLEXOFT."
               If (USNBR% Mod 100) <> 1 Then
                    Call COMUNI(TTXX$)
                    Call FINAL("")
                  Else
                    OPXX% = COMALTER%(TTXX$ + "\\Re-Intentar Transaccion\Continuar (Resultados Impredecibles !!!)\Cancelar Proceso")
                    If OPXX% <= 1 Then GoTo 28
                    If OPXX% = 2 Then GoTo 35
                    Call FINAL("")
               End If
            End If
35          On Error GoTo 0
          End If
        End With
        '
        MIRS.Close
        Set MIRS = Nothing
        GoTo 99
      End If
      '
      VALOCO$ = SAFETEXT$(NUEVALOR)
      VALOCO$ = REPLACAR$(VALOCO$, "'", " ")
      '
      ' ESTO QUE SIGUE CORREGIDO POR EPB EL 16/5/2009
      If TPC = 231 Then  ' ES UN CAMPO TEXTO
           LONGI = LONCAMPO(TABLA$, CAMPO$)
           If LONGI > 0 Then VALOCO$ = Left$(VALOCO$, LONGI)
        ElseIf TPC = 62 Or TPC = 56 Then
           VALOCO$ = TRIM$(FORMATNUM$(NUEVALOR, "F20.8"))
           If VALOCO$ = "" Then VALOCO$ = "0"
           If Right$(VALOCO$, 9) = ".00000000" Then VALOCO$ = Left$(VALOCO$, Len(VALOCO$) - 9)
      End If
      '
      If TPC = 58 Then
         If IsNumeric(NUEVALOR) = True Then
              FEXI% = NUEVALOR
           Else
              FFXX$ = NUEVALOR
              FEXI% = FECHANUM(FFXX$)
         End If
         If FEXI% > 0 Then
            VALOCO$ = FETEXCOR1$(FEXI%)
         End If
      End If
      '
      If TPC = 58 Or TPC = 231 Or TPC = 99 Then ' FECHA - TEXTO - MEMO
         VALOCO$ = "'" + VALOCO$ + "'"
      End If
      '
      SQLX$ = "UPDATE " + TABLA$
      SQLX$ = SQLX$ + " SET " + CAMPO$ + " = " & VALOCO$
      SQLX$ = SQLX$ & " WHERE " + CONDI$
      '
      ERRNU = 0: ERRDESC$ = ""
      On Error Resume Next
      FLEXCONN.Execute (FILTSQL$(SQLX$))
      If Err Then
         ERRNU = Err.Number
         ERRDE$ = Err.Description
         On Error GoTo 0
         If InStr(UCase$(MODO$), "NOMESS") < 1 Then
           TTYY$ = " Imposible Ejecutar Actualizacion:\  "
           TTXX$ = "'" + SQLX$ + "'"
           Call FRATEXTO(TTXX$, 96)
           For KKI% = 1 To CARETEX%
              TTYY$ = TTYY$ + "\" + RETEX$(KKI%)
           Next KKI%
           TTYY$ = TTYY$ + "\   "
           TTYY$ = TTYY$ + "\El Error Informado por el Servidor SQL es:"
           TTXX$ = "Error " & ERRNU & " - " & ERRDE$
           Call FRATEXTO(TTXX$, 96)
           For KKI% = 1 To CARETEX%
              TTYY$ = TTYY$ + "\" + RETEX$(KKI%)
           Next KKI%
           TTYY$ = TTYY$ + "\   \Comunique Este Mensaje a Mesa de Ayuda FLEXOFT"
           Call MENSERR(24, TTYY$)
         End If
      End If
      If InStr(UCase$(MODO$), "NOAUDIT") < 1 Then
         Call AUDITACT(TABLA$, CAMPO$, CONDI$, VALORANT, NUEVALOR) ' COMENTADO PROVISORIAMENTE 11/09/15
      End If
   End If
99 REGISTROSAFECTADOS& = NUMREG&
   '
'   If XVALO(TIMUT) > 0 Then
'      If XVALO(TIMUTAN) > 0 Then
'         FLEXCONN.CommandTimeout = TIMUTAN
'      End If
'   End If
   '
End Sub
Sub ACTUREGISTRO2(TABLA$, CAMPO$, CONDI$, NUEVALOR, REGISTROSAFECTADOS&, Optional MODO$, Optional TIMUT)
   '
'   If XVALO(TIMUT) > 0 Then
'      TIMUTAN = FLEXCONN.CommandTimeout
'      FLEXCONN.CommandTimeout = TIMUT
'   End If
   '
   Dim CAMP_CONSULTA1(), VAL_CONSULTA1(), CAMP_CONSULTA2(), VAL_CONSULTA2()
   Dim Campox As Field
   NUMREG& = CantRegistros&(TABLA$, CONDI$, MODO$)
   If NUMREG& > 0 Then
      Openforms = DoEvents
      ' Valor Anterior para Tabla de Auditoria
      If NUMREG& = 1 Then
        VALORANT = VALOBADA(TABLA$, CAMPO$, CONDI$, "NOMESS")
        Else
        VALORANT = "Varios" 'Si es mas de un registro, escribe "Varios"
      End If
      '
      If InStr(CAMPO$, ",") > 0 Then
         CAMP$ = CAMPO$
         Do Until InStr(CAMP$, ",") = 0
            POS1% = InStr(CAMP$, ",")
            If POS1% > 0 Then
              CAMPO1$ = Mid$(CAMP$, 1, POS1% - 1)
              campo2$ = Mid$(CAMP$, 1, POS1% - 1)
              POS2% = InStr(NUEVALOR, ",")
              VALOR2 = Mid$(NUEVALOR, 1, POS2% - 1)
              VALORACT1 = Mid$(NUEVALOR, 1, POS2% - 1)
            Else
              
            End If
15          TPC = TIPODECAMPO(TABLA$, CAMPO1$)
            If TPC = 99 Or TPC = 34 Or TPC = 173 Then
               L& = L& + 1
               ReDim Preserve CAMP_CONSULTA1(L&), VAL_CONSULTA1(L&)
               CAMP_CONSULTA1(L&) = CAMPO1$
               VAL_CONSULTA1(L&) = VALORACT1
               CAMPO1$ = " " & CAMPO1$ & "  AS NUCAM" & L& & " , "
               CAMP$ = CAMPO1$ = Mid$(CAMP$, POS1% + 1)
               ACUMUCAMP1$ = ACUMUCAMP1$ + CAMPO1$
            Else
               U& = U& + 1
               ReDim Preserve CAMP_CONSULTA2(U&), VAL_CONSULTA2(U&)
               CAMP_CONSULTA2(U&) = campo2$
               VAL_CONSULTA2(U&) = VALORACT1
               campo2$ = " " & campo2$ & "  AS NUCAM" & U& & " , "
               CAMP$ = Mid$(CAMP$, POS1% + 1)
               ACUMUCAMP2$ = ACUMUCAMP2$ + campo2$
              'GoTo 99
           End If
         Loop
         'CAMPO1$ = CAMPO1$ & CAMP$
'         If CAMP$ <> "" Then
'           GoTo 15 ' LEE EL ULTIMO
      End If
      '
      CAMPO$ = CAMPO1$
      'If TPC = 99 Or TPC = 34 Or TPC = 173 Then
      If UBound(CAMP_CONSULTA1) > 0 Then
        Dim MIRS As ADODB.Recordset
        Set MIRS = New ADODB.Recordset
        SQLX$ = "SELECT " & CAMPO$ & "  FROM " & TABLA & " WHERE " & CONDI$
25      On Error Resume Next
        MIRS.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
        If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
           On Error GoTo 0
           Call CapturaErrorOpen
           GoTo 25
        End If
        On Error GoTo 0
        RETRATA% = 0
        '
28      With MIRS
          If Not (.BOF And .EOF) Then
            On Error Resume Next
            For i& = 1 To UBound(CAMP_CONSULTA1)
               Campox = CAMP_CONSULTA1(i&)
               !Campox = VAL_CONSULTA1
            Next i&
'               !NUCAM = NUEVALOR
            .Update
            
            If Err Then
               ERRNU = Err.Number
               ERRDE$ = Err.Description
               On Error GoTo 0
               '
               If InStr(UCase$(MODO$), "RETRY") > 0 Then
                  RETRATA% = RETRATA% + 1
                  If RETRATA% < 8 Then GoTo 28
               End If
               '
               LONGI& = Len(CStr(NUEVALOR))
               LOMAX& = LONCAMPO(TABLA$, CAMPO$)
               TTXX$ = "ERROR No Recuperable en Transacción\de Actualización (ACTUREGISTRO)"
               TTXX$ = TTXX$ + "\al Procesar Campo: '" + CAMPO$ + "' (Long: " & LONGI& & " - Max: " & LOMAX& & ") en Tabla: '" + TABLA$ + "'."
               TTXX$ = TTXX$ + "\   \El Error Informado por el Servidor SQL es:"
                  TTZZ$ = "Error " & ERRNU & " - " & ERRDE$
                  Call FRATEXTO(TTZZ$, 96)
                  For KKI% = 1 To CARETEX%
                     TTXX$ = TTXX$ + "\" + RETEX$(KKI%)
                  Next KKI%
               TTXX$ = TTXX$ + "\   \Probable Saturación de la Capacidad\de la Base de Datos."
               TTXX$ = TTXX$ + "\   \Verifique Incremento Automatico del Tamaño (10 %)"
               TTXX$ = TTXX$ + "\y Crecimiento Ilimitado (Unrestricted Growth),"
               TTXX$ = TTXX$ + "\tanto en el Archivo de la B.D. Principal"
               TTXX$ = TTXX$ + "\como Tambien en el Historico de Transacciones (Transaction Log) !!!"
               TTXX$ = TTXX$ + "\   \Consulte Mesa de Ayuda FLEXOFT."
               If (USNBR% Mod 100) <> 1 Then
                    Call COMUNI(TTXX$)
                    Call FINAL("")
                  Else
                    OPXX% = COMALTER%(TTXX$ + "\\Re-Intentar Transaccion\Continuar (Resultados Impredecibles !!!)\Cancelar Proceso")
                    If OPXX% <= 1 Then GoTo 28
                    If OPXX% = 2 Then GoTo 35
                    Call FINAL("")
               End If
            End If
35          On Error GoTo 0
          End If
        End With
        '
        MIRS.Close
        Set MIRS = Nothing
        GoTo 99
      End If
      '
      VALOCO$ = SAFETEXT$(NUEVALOR)
      VALOCO$ = REPLACAR$(VALOCO$, "'", " ")
      '
      ' ESTO QUE SIGUE CORREGIDO POR EPB EL 16/5/2009
      'If TPC = 231 Then  ' ES UN CAMPO TEXTO
      If UBound(CAMP_CONSULTA2) > 0 Then
           LONGI = LONCAMPO(TABLA$, CAMPO$)
           VALOCO$ = Left$(VALOCO$, LONGI)
        ElseIf TPC = 62 Or TPC = 56 Then
           VALOCO$ = TRIM$(FORMATNUM$(NUEVALOR, "F20.8"))
           If VALOCO$ = "" Then VALOCO$ = "0"
           If Right$(VALOCO$, 9) = ".00000000" Then VALOCO$ = Left$(VALOCO$, Len(VALOCO$) - 9)
      End If
      '
      If TPC = 58 Then
         If IsNumeric(NUEVALOR) = True Then
              FEXI% = NUEVALOR
           Else
              FFXX$ = NUEVALOR
              FEXI% = FECHANUM(FFXX$)
         End If
         If FEXI% > 0 Then
            VALOCO$ = FETEXCOR1$(FEXI%)
         End If
      End If
      '
      If TPC = 58 Or TPC = 231 Or TPC = 99 Then ' FECHA - TEXTO - MEMO
         VALOCO$ = "'" + VALOCO$ + "'"
      End If
      '
      SQLX$ = "UPDATE " + TABLA$
      SQLX$ = SQLX$ + " SET " + CAMPO$ + " = " & VALOCO$
      SQLX$ = SQLX$ & " WHERE " + CONDI$
      '
      ERRNU = 0: ERRDESC$ = ""
      On Error Resume Next
      FLEXCONN.Execute (FILTSQL$(SQLX$))
      If Err Then
         ERRNU = Err.Number
         ERRDE$ = Err.Description
         On Error GoTo 0
         If InStr(UCase$(MODO$), "NOMESS") < 1 Then
           TTYY$ = " Imposible Ejecutar Actualizacion:\  "
           TTXX$ = "'" + SQLX$ + "'"
           Call FRATEXTO(TTXX$, 96)
           For KKI% = 1 To CARETEX%
              TTYY$ = TTYY$ + "\" + RETEX$(KKI%)
           Next KKI%
           TTYY$ = TTYY$ + "\   "
           TTYY$ = TTYY$ + "\El Error Informado por el Servidor SQL es:"
           TTXX$ = "Error " & ERRNU & " - " & ERRDE$
           Call FRATEXTO(TTXX$, 96)
           For KKI% = 1 To CARETEX%
              TTYY$ = TTYY$ + "\" + RETEX$(KKI%)
           Next KKI%
           TTYY$ = TTYY$ + "\   \Comunique Este Mensaje a Mesa de Ayuda FLEXOFT"
           Call MENSERR(24, TTYY$)
         End If
      End If
      If InStr(UCase$(MODO$), "NOAUDIT") < 1 Then
         Call AUDITACT(TABLA$, CAMPO$, CONDI$, VALORANT, NUEVALOR)
      End If
   End If
99 REGISTROSAFECTADOS& = NUMREG&
   '
'   If XVALO(TIMUT) > 0 Then
'      If XVALO(TIMUTAN) > 0 Then
'         FLEXCONN.CommandTimeout = TIMUTAN
'      End If
'   End If
   '
End Sub

Sub PoneEnCeroNull(TABLA$, CAMPO$, REGISTROSAFECTADOS&, Optional MODO$)
   '
   'PONE EN 0 O VACIO LOS CAMPOS NULL
   '
   SCMA = Screen.MousePointer
   Screen.MousePointer = 11
   '
   CONDI$ = CAMPO$ + " IS NULL "
   NUMREG& = CantRegistros&(TABLA$, CONDI$)
   If NUMREG& > 0 Then
      TPC = TIPODECAMPO(TABLA$, CAMPO$)
      If TPC = 58 Or TPC = 231 Or TPC = 99 Then ' FECHA - TEXTO - MEMO
         VALOCO = "''"
      Else
         VALOCO = 0
      End If
      SQLX$ = "UPDATE " + TABLA$
      SQLX$ = SQLX$ + " SET " + CAMPO$ + " = " & VALOCO
      SQLX$ = SQLX$ & " WHERE " + CONDI$
      
      ERRNU = 0: ERRDESC$ = ""
      On Error Resume Next
      FLEXCONN.Execute (FILTSQL$(SQLX$))
      If Err Then
         ERRNU = Err.Number
         ERRDE$ = UCase$(Err.Description)
         On Error GoTo 0
         '
         If InStr(ERRDE$, "TIEMPO DE ESPERA") > 0 Or InStr(ERRDE$, "TIMEOUT") > 0 Then
            ' PONE EN CERO POR BARRIDA
            Call COMUNI("FLEXOFT Debe Ahora Ajustar a Cero\el Valor del Campo '" + CAMPO$ + "'\en la Tabla '" + TABLA$ + "'\   \Esto Puede Demorar Varios Minutos")
            SQLX$ = "SELECT " + CAMPO$ + " AS VALORSE FROM " + TABLA$ & " WHERE " + CONDI$
            Dim MOVISE As ADODB.Recordset
            Set MOVISE = New ADODB.Recordset
            '
25          On Error Resume Next
            MOVISE.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
            If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
               On Error GoTo 0
               Call CapturaErrorOpen
               GoTo 25
            End If
            On Error GoTo 0
            '
            jj& = 0
            With MOVISE
               Do While Not .EOF
                  jj& = jj& + 1: Call TRACE(20, jj&, NUMREG&)
                  !VALORSE = VALOCO
                  .Update
               .MoveNext
               Loop
            End With
            MOVISE.Close
            Set MOVISE = Nothing
            GoTo 99
         End If
         '
         '
         If InStr(UCase$(MODO$), "NOMESS") < 1 Then
            TTYY$ = " Imposible Ejecutar Actualizacion:\  "
            TTXX$ = "'" + SQLX$ + "'"
            Call FRATEXTO(TTXX$, 96)
            For KKI% = 1 To CARETEX%
               TTYY$ = TTYY$ + "\" + RETEX$(KKI%)
            Next KKI%
            TTYY$ = TTYY$ + "\   "
            TTYY$ = TTYY$ + "\El Error Informado por el Servidor SQL es:"
            TTXX$ = "Error " & ERRNU & " - " & ERRDE$
            Call FRATEXTO(TTXX$, 96)
            For KKI% = 1 To CARETEX%
               TTYY$ = TTYY$ + "\" + RETEX$(KKI%)
            Next KKI%
            TTYY$ = TTYY$ + "\   \Comunique Este Mensaje a Mesa de Ayuda FLEXOFT"
            Call MENSERR(24, TTYY$)
         End If
      End If
   End If
   '
99 REGISTROSAFECTADOS& = NUMREG&
   Screen.MousePointer = SCMA
   '
End Sub
'
Sub CreaRegistro(NOTAB$, NOCAM$, NUEVALOR, Optional MODO$)

      TPC = TIPODECAMPO(NOTAB$, NOCAM$)

      VALOCO$ = SAFETEXT$(NUEVALOR)
      VALOCO$ = REPLACAR$(VALOCO$, "'", " ")
      '
      ' ESTO QUE SIGUE CORREGIDO POR EPB EL 16/5/2009
      If TPC = 231 Then  ' ES UN CAMPO TEXTO
           LONGI = LONCAMPO(NOTAB$, NOCAM$)
           VALOCO$ = Left$(VALOCO$, LONGI)
        ElseIf TPC = 62 Then
           VALOCO$ = TRIM$(FORMATNUM$(NUEVALOR, "F20.6"))
           If VALOCO$ = "" Then VALOCO$ = "0"
      End If
      '
      If TPC = 58 Then
         FFXX$ = NUEVALOR
         FEXI% = FECHANUM(FFXX$)
         If FEXI% > 0 Then
            VALOCO$ = FETEXCOR1$(FEXI%)
         End If
      End If
      '
      If TPC = 58 Or TPC = 231 Or TPC = 99 Then ' FECHA - TEXTO - MEMO
         VALOCO$ = "'" + VALOCO$ + "'"
      End If
      '
      SQLX$ = "INSERT INTO " & NOTAB$
      SQLX$ = SQLX$ + " (" + NOCAM$
      SQLX$ = SQLX$ + ")"
      SQLX$ = SQLX$ + " Values (" & VALOCO$
      SQLX$ = SQLX$ + ")"
      '
      On Error Resume Next
       FLEXCONN.Execute (FILTSQL$(SQLX$))
      If Err Then
         On Error GoTo 0
         If InStr(UCase$(MODO$), "NOMESS") < 1 Then
           Call MENSERR(24, " Imposible Ejecutar Actualizacion:\  \'" + SQLX$ + "'")
         End If
      End If
      CONDI$ = NOCAM$ & " = " & VALOCO$ & " AND CodiEmpr IS NULL"
      Call ACTUREGISTRO(NOTAB$, "CODIEMPR", CONDI$, CodiEmp%, RAFE&, "NOMESS" + "/" + MODO$)
      '
End Sub

Function TIPODECAMPO(NTABLA$, NCAMPO$)
    '
    Call ABRECONEXION
    TIPODECAMPO = 0
    If EXISTELCAMPO%(NBASE$, NTABLA$, NCAMPO$, TCAMPO, LOCAMPO) > 0 Then
       TIPODECAMPO = TCAMPO
    End If
    If TIPODECAMPO > 0 Then Exit Function
    '
    ' vieja rutina
    Dim MIRS As ADODB.Recordset
    Set MIRS = New ADODB.Recordset
    strSQL = "Select systypes.name, systypes.xtype from sysobjects inner join syscolumns on "
    strSQL = strSQL & "sysobjects.id = syscolumns.id inner join systypes "
    strSQL = strSQL & "on syscolumns.xtype = systypes.xtype "
    strSQL = strSQL & "where sysobjects.name = '" & NTABLA$ & "' "
    strSQL = strSQL & "and syscolumns.name ='" & NCAMPO$ & "'"
    Set MIRS = FLEXCONN.Execute(FILTSQL$(strSQL))
    If Not (MIRS.EOF And MIRS.BOF) Then
        TIPODECAMPO = MIRS!xtype
    End If

End Function
Function TIPODECAMPO_DESTINO(NTABLA$, NCAMPO$)
    '
    TIPODECAMPO_DESTINO = 0
    If EXISTELCAMPO_DESTINO%(NBASE$, NTABLA$, NCAMPO$, TCAMPO, LOCAMPO) > 0 Then
       TIPODECAMPO_DESTINO = TCAMPO
    End If
    If TIPODECAMPO_DESTINO > 0 Then Exit Function
    '
    ' vieja rutina
    Dim MIRS As ADODB.Recordset
    Set MIRS = New ADODB.Recordset
    strSQL = "Select systypes.name, systypes.xtype from sysobjects inner join syscolumns on "
    strSQL = strSQL & "sysobjects.id = syscolumns.id inner join systypes "
    strSQL = strSQL & "on syscolumns.xtype = systypes.xtype "
    strSQL = strSQL & "where sysobjects.name = '" & NTABLA$ & "' "
    strSQL = strSQL & "and syscolumns.name ='" & NCAMPO$ & "'"
    Set MIRS = FLEXCONN.Execute(FILTSQL$(strSQL))
    If Not (MIRS.EOF And MIRS.BOF) Then
        TIPODECAMPO_DESTINO = MIRS!xtype
    End If

End Function

'
Function LONCAMPO(TABLA$, CAMPO$)
   '
   LONCAMPO = 0
   If EXISTELCAMPO%(NBASE$, TABLA$, CAMPO$, TCAMPO, LOCAMPO) > 0 Then
      LONCAMPO = LOCAMPO
   End If
   If LONCAMPO > 0 Then Exit Function
   '
   ' vieja rutina
   'STRSQL = "SELECT SYSCOLUMNS.NAME AS NAME, SYSCOLUMNS.XTYPE AS XTYPE1, SYSCOLUMNS.LENGTH AS LENG FROM syscolumns "
   Dim MIRS As ADODB.Recordset
   Set MIRS = New ADODB.Recordset
   strSQL = "SELECT SYSCOLUMNS.LENGTH AS LENG FROM syscolumns "
   strSQL = strSQL + "INNER JOIN sysobjects "
   strSQL = strSQL & "ON sysobjects.id = syscolumns.id "
   strSQL = strSQL & "where sysobjects.name = '" & TABLA$ & "' "
   strSQL = strSQL & "AND syscolumns.name = '" & UCase$(CAMPO$) & "'"
   Set MIRS = FLEXCONN.Execute(FILTSQL$(strSQL))
   If Not (MIRS.EOF And MIRS.BOF) Then
      LONCAMPO = Int(MIRS!LENG / 2)
   End If
   '
End Function
'
Function INDICEDECAMPO%(TABLA$, CAMPO$)
   '
   Dim MIRS As ADODB.Recordset
   Set MIRS = New ADODB.Recordset
   strSQL = "Select * from sysindexes WHERE id = OBJECT_ID('" & TABLA$ & "') AND NAME = 'I" + CAMPO$ + "' "
   CONIN% = 0
   '
   On Error Resume Next
   Set MIRS = FLEXCONN.Execute(strSQL)
       If Not (MIRS.EOF And MIRS.BOF) Then CONIN% = 1                         ' el indice ya existe
   MIRS.Close
   Set MIRS = Nothing
   On Error GoTo 0
   '
   INDICEDECAMPO% = CONIN%
End Function
'
Sub GRAVALCON(USNUM%, PROCE$, CLAVE$, VALCON$)
    '
    If TRIM$(UCase$(VALCONTROL$(USNUM%, PROCE$, CLAVE$))) = TRIM$(UCase$(VALCON$)) Then
       Exit Sub
    End If
    '
    USNU% = USNUM% Mod 100
    PROBUS$ = Left$(TRIM$(UCase$(PROCE$)), 128)
    CODBUS$ = Left$(TRIM$(UCase$(CLAVE$)), 128)
    '
    Dim PARATEMP As ADODB.Recordset
    Set PARATEMP = New ADODB.Recordset
    
    SQLX$ = "SELECT * FROM USER_PARMS "
    SQLX$ = SQLX$ + " WHERE USUARIO = " & USNU%
    SQLX$ = SQLX$ + " AND PROCESO = '" & PROBUS$ & "'"
    SQLX$ = SQLX$ + " AND CLAVE = '" & CODBUS$ & "'"
    PARATEMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
    '
    With PARATEMP
       If .EOF Then
         .AddNew   ' YA CONVERTIDO
       End If
       !USUARIO = USNU%
       !Proceso = PROBUS$
       !CLAVE = CODBUS$
       !ValoCon = VALCON$
       On Error Resume Next
       !CODIEMPR = CodiEmp%
       On Error GoTo 0
       .Update
    End With
    '
90  PARATEMP.Close
    
'    If TRIM$(UCase$(VALCONTROL$(USNUM%, PROCE$, CLAVE$))) = TRIM$(UCase$(VALCON$)) Then
'       Exit Sub
'    End If
'    '
'    USNU% = USNUM% Mod 100
'    PROBUS$ = Left$(TRIM$(UCase$(PROCE$)), 128)
'    CODBUS$ = Left$(TRIM$(UCase$(CLAVE$)), 128)
'    '
'    Call ABRETACON
'    SQLX$ = "SELECT * FROM USER_PARMS "
'    SQLX$ = SQLX$ + " WHERE USUARIO = " & USNU%
'    SQLX$ = SQLX$ + " AND PROCESO = '" & PROBUS$ & "'"
'    SQLX$ = SQLX$ + " AND CLAVE = '" & CODBUS$ & "'"
'    Set PARATEMP = BaDaConsul.OpenRecordset(SQLX$, dbOpenDynaset)
'    '
'    With PARATEMP
'       If Not (.BOF And .EOF) Then
'           .Edit
'         Else
'           .AddNew   ' YA CONVERTIDO
'           On Error Resume Next
'           !CodiEmpr = CodiEmp%
'           On Error GoTo 0
'       End If
'       !USUARIO = USNU%
'       !Proceso = PROBUS$
'       !CLAVE = CODBUS$
'       !ValoCon = VALCON$
'       On Error Resume Next
'       !CodiEmpr = CodiEmp%
'       On Error GoTo 0
'       .Update
'    End With
'    '
'90  PARATEMP.Close
    '
End Sub
'
Function VALCONTROL$(USNUM%, PROCE$, CLAVE$)
    '
    VCXXX$ = ""
    USNU% = USNUM% Mod 100
    PROBUS$ = Left$(TRIM$(UCase$(PROCE$)), 128)
    CODBUS$ = Left$(TRIM$(UCase$(CLAVE$)), 128)
    '
    Dim PARATEMP As ADODB.Recordset
    Set PARATEMP = New ADODB.Recordset
    SQLX$ = "SELECT * FROM USER_PARMS "
    SQLX$ = SQLX$ + " WHERE USUARIO = " & USNU%
    SQLX$ = SQLX$ + " AND PROCESO = '" & PROBUS$ & "'"
    SQLX$ = SQLX$ + " AND CLAVE = '" & CODBUS$ & "'"
    PARATEMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
    '
    With PARATEMP
       If Not (.BOF And .EOF) Then
          VCXXX$ = SAFETEXT$(!ValoCon)
       End If
    End With
    PARATEMP.Close: Set PARATEMP = Nothing
    '
    VALCONTROL$ = VCXXX$

'    VCXXX$ = ""
'    USNU% = USNUM% Mod 100
'    PROBUS$ = Left$(TRIM$(UCase$(PROCE$)), 128)
'    CODBUS$ = Left$(TRIM$(UCase$(CLAVE$)), 128)
'    '
'    Call ABRETACON
'    SQLX$ = "SELECT * FROM USER_PARMS "
'    SQLX$ = SQLX$ + " WHERE USUARIO = " & USNU%
'    SQLX$ = SQLX$ + " AND PROCESO = '" & PROBUS$ & "'"
'    SQLX$ = SQLX$ + " AND CLAVE = '" & CODBUS$ & "'"
'    Set PARATEMP = BaDaConsul.OpenRecordset(SQLX$, dbOpenSnapshot)
'    '
'    With PARATEMP
'       If Not (.BOF And .EOF) Then
'          VCXXX$ = !ValoCon
'       End If
'    End With
'    PARATEMP.Close
'    '
'    VALCONTROL$ = VCXXX$
    '
End Function

Sub ENUMERA(TABLA$, NOMBREDECAMPO$, CONDICION$)
    '
    Call CREACAMPO("", TABLA$, NOMBREDECAMPO$, 4, 0)
    SQLX$ = "SELECT " + NOMBREDECAMPO$ + " AS NUMERORDEN FROM " + TABLA$ + " WHERE " & CONDICION$
    '
    Dim ENUME As ADODB.Recordset
    Set ENUME = New ADODB.Recordset
25  On Error Resume Next
    ENUME.Open FILTSQL$(SQLX$), FLEXCONN, adOpenStatic, adLockPessimistic, adCmdText

    If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
       On Error GoTo 0
       Call CapturaErrorOpen
       GoTo 25
    End If
    On Error GoTo 0
    '
    jj& = 0
    With ENUME
       If Not (.BOF And .EOF) Then
          Do While Not .EOF
            jj& = jj& + 1
            !NUMERORDEN = jj&
            .Update
          .MoveNext
          Loop
       End If
    End With
    ENUME.Close
    Set ENUME = Nothing
    '
End Sub

Sub RENUMERA(TABLA$, NOMBREDECAMPO$, CONDICION$)
    '
    Dim valor(), CANREG&
    CANREG& = CantRegistros&(TABLA$, CONDICION$)
    ReDim valor(CANREG&)
    '
    Call CREACAMPO("", TABLA$, NOMBREDECAMPO$, 4, 0)
    SQLX$ = "SELECT " + NOMBREDECAMPO$ + " AS NUMERORDEN FROM " + TABLA$ + " WHERE " & CONDICION$
    '
    Dim ENUME As ADODB.Recordset
    Set ENUME = New ADODB.Recordset
25  On Error Resume Next
    ENUME.Open FILTSQL$(SQLX$), FLEXCONN, adOpenStatic, adLockPessimistic, adCmdText
    If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
      On Error GoTo 0
      Call CapturaErrorOpen
      GoTo 25
    End If
    On Error GoTo 0
    '
    jj& = 0
    VALOMAX = 0
    With ENUME
       If Not (.BOF And .EOF) Then
          Do While Not .EOF
            jj& = jj& + 1
            NUMACT = XVALO(!NUMERORDEN)
            '
            If NUMACT < 1 Then
               NUMACT = VALOMAX + 1
               !NUMERORDEN = NUMACT
               .Update
               If NUMACT > VALOMAX Then VALOMAX = NUMACT
            End If
            '
            For KK& = 1 To jj&
               If valor(KK&) = NUMACT Then
                  NUMACT = VALOMAX + 1
                  !NUMERORDEN = NUMACT
                  .Update
                  If NUMACT > VALOMAX Then VALOMAX = NUMACT
               End If
            Next KK&
            valor(jj&) = NUMACT
            If NUMACT > VALOMAX Then VALOMAX = NUMACT
          .MoveNext
          Loop
       End If
    End With
    ENUME.Close
    Set ENUME = Nothing
    '
End Sub

Sub CARGALISTA(NOMBREDELISTA, TABLA$, LISTADECAMPOS$, CONDICION$, Optional MODOSEL$)
    '
        '
    Dim NOCAMPO$(), FORMX$()
    ReDim NOCAMPO$(1), FORMX$(1)
    '
    MODOSE$ = TRIM$(UCase$(SAFETEXT$(MODOSEL$)))
    '
    If InStr(MODOSE$, "NOCLEAR") > 0 Then GoTo 10
    On Error Resume Next
    NOMBREDELISTA.Clear
    If Err Then
       On Error GoTo 0
       Call MENSERR(24, "Imposible Limpiar Lista '" & NOMBREDELISTA & "'.")
       Call FINAL("")
    End If
    On Error GoTo 0
    
    '
10  TIPODELISTA = "GRID"
    On Error Resume Next
    CANCOL& = NOMBREDELISTA.Cols
    If Err Then
       On Error GoTo 0
       TIPODELISTA = "LIST"
    End If
    On Error GoTo 0
    '
    If TIPODELISTA = "GRID" Then
       NOMBREDELISTA.Rows = 2
       ALTUREN = NOMBREDELISTA.RowHeight(0) + NOMBREDELISTA.RowHeight(1)
       While ALTUREN < NOMBREDELISTA.Height
          NOMBREDELISTA.Rows = NOMBREDELISTA.Rows + 1
          ALTUREN = ALTUREN + NOMBREDELISTA.RowHeight(NOMBREDELISTA.Rows - 1)
       Wend
    End If
    '
    LSTX$ = TRIM$(LISTADECAMPOS$)
    CANCA% = 0
    SQLX$ = ""
    While LSTX$ <> ""
       PPXX% = InStr(LSTX$, ";")
       PPXX2% = InStr(LSTX$, ",")
       If PPXX2% > 0 And PPXX2% < PPXX% Then PPXX% = PPXX2%
       If PPXX% < 1 Then PPXX% = 1 + Len(LSTX$)
       '
       CANCA% = CANCA% + 1
       If CANCA% >= UBound(NOCAMPO$) Then
          ReDim Preserve NOCAMPO$(CANCA%), FORMX$(CANCA%)
       End If
       NOCX$ = Left$(LSTX$, PPXX% - 1)
       LSTX$ = Mid$(LSTX$, PPXX% + 1)
       '
       PPXX% = InStr(NOCX$, "/")
       If PPXX% < 1 Then PPXX% = 1 + Len(NOCX$)
       NOCAMPO$(CANCA%) = Left$(NOCX$, PPXX% - 1)
       FORMX$(CANCA%) = Mid$(NOCX$, PPXX% + 1)
       If FORMX$(CANCA%) = "" Then FORMX$(CANCA%) = "A16"
       '
       If TIPODELISTA = "GRID" Then
          NOMBREDELISTA.Cols = CANCA%
       End If
       '
       If SQLX$ <> "" Then SQLX$ = SQLX$ + ", "
       SQLX$ = SQLX$ + NOCAMPO$(CANCA%) + " AS CAMPO" + TRIM$(Str$(CANCA%))
    Wend
    '
    TABLA1$ = TABLA$: TABLA2$ = "": IJO$ = ""
    PPXX% = InStr(TABLA$, "+")
    If PPXX% > 0 Then
       TABLA1$ = TRIM$(Left$(TABLA$, PPXX% - 1))
       TABLA$ = TRIM$(Mid$(TABLA$, PPXX% + 1))
       PPXX% = InStr(TABLA$, "/")
       If PPXX% > 0 Then
         TABLA2$ = TRIM$(Left$(TABLA$, PPXX% - 1))
         IJO$ = TRIM$(Mid$(TABLA$, PPXX% + 1))
         TABLA$ = TABLA1$ + " INNER JOIN " + TABLA2$ + " ON " + IJO$
       End If
    End If
    
    If InStr(MODOSE$, "DISTINCT") > 0 Then SQLX$ = "DISTINCT " + SQLX$
    SQLX$ = "SELECT " + SQLX$ + " FROM " + TABLA$
    If TRIM$(CONDICION$) <> "" Then
       SQLX$ = SQLX$ + " WHERE " & CONDICION$
    End If
    On Error Resume Next
    Set TABTMP = FLEXCONN.Execute(FILTSQL$(SQLX$))
    If Err Then
       On Error GoTo 0
       MsgBox "Imposible Ejecutar Consulta SQL: " + vbCrLf + vbCrLf & SQLX$, vbExclamation
       Call FINAL("")
    End If
    '
    With TABTMP
       'CARGA EL LISTADO CON LAS COTIZACIONES YA INGRESADAS
       TTXX$ = ""
       J% = 0: Fila& = 0
       Do While Not .EOF
           TTXX$ = "": COLUI% = 0
           '
           If TIPODELISTA = "GRID" Then
              Fila& = Fila& + 1
37            If Fila& > NOMBREDELISTA.Rows - 1 Then
                 NOMBREDELISTA.Rows = NOMBREDELISTA.Rows + 1
                 GoTo 37
              End If
           End If
           '
           For U& = 1 To .Fields.Count
              valor = .Fields(U& - 1).Value
              FORMATX$ = FORMX$(U&)
              TECO$ = ""
              PPXX% = InStr(FORMATX$, "+")
              If PPXX% > 0 Then
                 TECO$ = TRIM$(Mid$(FORMATX$, PPXX% + 1))
                 FORMATX$ = Left$(FORMATX$, PPXX% - 1)
              End If
              '
              ENCAB$ = ""
              PPXX% = InStr(FORMATX$, "/")
              If PPXX% > 0 Then
                 ENCAB$ = Mid$(FORMATX$, PPXX% + 1)
                 FORMATX$ = Left$(FORMATX$, PPXX% - 1)
              End If
              TTYY$ = FORMATNUM$(valor, FORMATX$)
              If TRIM$(TTYY$) = "" Then
                  TTYY$ = Space$(XVALO(Mid$(FORMATX$, 2)))
              End If
              If TIPODELISTA = "LIST" Then
                   If TTXX$ <> "" Then TTXX$ = TTXX$ + " "
                   TTXX$ = TTXX$ + TTYY$
                 ElseIf TIPODELISTA = "GRID" Then
                   COLUI% = COLUI% + 1
17                 If COLUI% > NOMBREDELISTA.Cols Then
                      NOMBREDELISTA.Cols = NOMBREDELISTA.Cols + 1
                      GoTo 17
                   End If
                   If Fila& = 1 Then
                      FORMALTER.TEPRUEBA.Font = NOMBREDELISTA.Font
                      FORMALTER.TEPRUEBA.FontSize = NOMBREDELISTA.FontSize
                      FORMALTER.TEPRUEBA.Caption = "l" + TMASCARA$(FORMATX$)
                      NOMBREDELISTA.ColWidth(COLUI% - 1) = FORMALTER.TEPRUEBA.Width
                      NOMBREDELISTA.ColAlignment(COLUI% - 1) = 6
                      If Left$(FORMATX$, 1) = "A" Or Left$(FORMATX$, 1) = "C" Or Left$(FORMATX$, 1) = "D" Then
                         NOMBREDELISTA.ColAlignment(COLUI% - 1) = 1
                      End If
                      NOMBREDELISTA.FontBold = True
                      NOMBREDELISTA.TextMatrix(0, COLUI% - 1) = ENCAB$
                      NOMBREDELISTA.FontBold = False
                   End If
                   If Left$(FORMATX$, 1) = "A" Then TTYY$ = valor       ' PARA QUE EN LA GRILLA LO PONGA COMPLETO
                   NOMBREDELISTA.TextMatrix(Fila&, COLUI% - 1) = TTYY$
              End If
              '
              If TECO$ <> "" Then
                 PPXX% = InStr(TECO$, "/")
                 If PPXX% < 1 Then PPXX% = 1 + Len(TECO$)
                 FORMATX$ = Mid$(TECO$, PPXX% + 1)
                 TECO$ = TRIM$(UCase$(Left$(TECO$, PPXX% - 1)))
                 '
                 ENCAB$ = ""
                 PPXX% = InStr(FORMATX$, "/")
                 If PPXX% > 0 Then
                    ENCAB$ = Mid$(FORMATX$, PPXX% + 1)
                    FORMATX$ = Left$(FORMATX$, PPXX% - 1)
                 End If
                 If FORMATX$ = "" Then FORMATX$ = "A16"
                 '
                 VALOECO = Null
                 If Left$(TECO$, 5) = "DEUDO" Or Left$(TECO$, 5) = "CLIEN" Then
                   VALOECO = rasocli$(XVALO(valor))
                 ElseIf Left$(TECO$, 5) = "PROVE" Or Left$(TECO$, 5) = "ACRED" Then
                   VALOECO = rasocli$((-1) * XVALO(valor))
                 ElseIf Left$(TECO$, 5) = "MASTE" Then
                   TXVALO$ = valor
                     If Len(TXVALO$) > 4 Then
                          CIXI% = CODINT%(TXVALO$)
                        Else
                          CIXI% = XVALO(valor)
                     End If
                   VALOECO = DESCRIT0$(CIXI%)
                 ElseIf EXISTE%(LUDAT$ + TECO$ + ".RFS") Then
                   LFF% = LENFIELD%(TECO$, 1)
                   ABUS$ = MKS$(XVALO(valor))
                   On Error Resume Next
                   If LFF% = 2 Then ABUS$ = MKI$(Int(XVALO(valor)))
                   If LFF% = 1 Then ABUS$ = Chr$(Int(XVALO(valor)))
                   If LFF% > 4 Then ABUS$ = TRIM$(valor) 'ESTO LO AGREGUE POR QUE SI ABUS$ ES TEXTO NO LO CONSIDERA
                   On Error GoTo 0
                   VALOECO = ECOARCH$(TECO$, ABUS$)
                 End If
                 '
                 TTYY$ = FORMATNUM$(VALOECO, FORMATX$)
                 If TRIM$(TTYY$) = "" Then
                     TTYY$ = Space$(XVALO(Mid$(FORMATX$, 2)))
                 End If
                 If TIPODELISTA = "LIST" Then
                      If TTXX$ <> "" Then TTXX$ = TTXX$ + " "
                      TTXX$ = TTXX$ + TTYY$
                    ElseIf TIPODELISTA = "GRID" Then
                      COLUI% = COLUI% + 1
27                    If COLUI% > NOMBREDELISTA.Cols Then
                         NOMBREDELISTA.Cols = NOMBREDELISTA.Cols + 1
                         GoTo 27
                      End If
                      If Fila& = 1 Then
                         FORMALTER.TEPRUEBA.Font = NOMBREDELISTA.Font
                         FORMALTER.TEPRUEBA.FontSize = NOMBREDELISTA.FontSize
                         FORMALTER.TEPRUEBA.Caption = "l" + TMASCARA$(FORMATX$)
                         NOMBREDELISTA.ColWidth(COLUI% - 1) = FORMALTER.TEPRUEBA.Width
                         NOMBREDELISTA.ColAlignment(COLUI% - 1) = 6
                         If Left$(FORMATX$, 1) = "A" Or Left$(FORMATX$, 1) = "C" Or Left$(FORMATX$, 1) = "D" Then
                            NOMBREDELISTA.ColAlignment(COLUI% - 1) = 1
                         End If
                         NOMBREDELISTA.FontBold = True
                         NOMBREDELISTA.TextMatrix(0, COLUI% - 1) = ENCAB$
                         NOMBREDELISTA.FontBold = False
                      End If
                      If Left$(FORMATX$, 1) = "A" Then TTYY$ = VALOECO       ' PARA QUE EN LA GRILLA LO PONGA COMPLETO
                      NOMBREDELISTA.TextMatrix(Fila&, COLUI% - 1) = TTYY$
                 End If
                 '
              End If
           Next U&
           '
           If TIPODELISTA = "LIST" Then
               NOMBREDELISTA.AddItem TTXX$
           End If
           '
       .MoveNext
       Loop
    End With
    '
    If TIPODELISTA = GRID Then
    End If
    '
    TABTMP.Close

    'CODIGO ANTERIOR
'''''''    Dim NOCAMPO$(), FORMX$()
'''''''    ReDim NOCAMPO$(1), FORMX$(1)
'''''''    '
'''''''    MODOSE$ = TRIM$(UCase$(SAFETEXT$(MODOSEL$)))
'''''''    '
'''''''    If InStr(MODOSE$, "NOCLEAR") > 0 Then GoTo 10
'''''''    On Error Resume Next
'''''''    NOMBREDELISTA.Clear
'''''''    If Err Then
'''''''       On Error GoTo 0
'''''''       Call MENSERR(24, "Imposible Limpiar Lista '" & NOMBREDELISTA & "'.")
'''''''       Call FINAL("")
'''''''    End If
'''''''    On Error GoTo 0
'''''''
'''''''    '
'''''''10  TIPODELISTA = "GRID"
'''''''    On Error Resume Next
'''''''    CANCOL& = NOMBREDELISTA.Cols
'''''''    If Err Then
'''''''       On Error GoTo 0
'''''''       TIPODELISTA = "LIST"
'''''''    End If
'''''''    On Error GoTo 0
'''''''    '
'''''''    If TIPODELISTA = "GRID" Then
'''''''       NOMBREDELISTA.Rows = 2
'''''''       ALTUREN = NOMBREDELISTA.RowHeight(0) + NOMBREDELISTA.RowHeight(1)
'''''''       While ALTUREN < NOMBREDELISTA.Height
'''''''          NOMBREDELISTA.Rows = NOMBREDELISTA.Rows + 1
'''''''          ALTUREN = ALTUREN + NOMBREDELISTA.RowHeight(NOMBREDELISTA.Rows - 1)
'''''''       Wend
'''''''    End If
'''''''    '
'''''''    LSTX$ = TRIM$(LISTADECAMPOS$)
'''''''    CANCA% = 0
'''''''    sqlx$ = ""
'''''''    While LSTX$ <> ""
'''''''       PPXX% = InStr(LSTX$, ";")
'''''''       PPXX2% = InStr(LSTX$, ",")
'''''''       If PPXX2% > 0 And PPXX2% < PPXX% Then PPXX% = PPXX2%
'''''''       If PPXX% < 1 Then PPXX% = 1 + Len(LSTX$)
'''''''       '
'''''''       CANCA% = CANCA% + 1
'''''''       If CANCA% >= UBound(NOCAMPO$) Then
'''''''          ReDim Preserve NOCAMPO$(CANCA%), FORMX$(CANCA%)
'''''''       End If
'''''''       NOCX$ = Left$(LSTX$, PPXX% - 1)
'''''''       LSTX$ = Mid$(LSTX$, PPXX% + 1)
'''''''       '
'''''''       PPXX% = InStr(NOCX$, "/")
'''''''       If PPXX% < 1 Then PPXX% = 1 + Len(NOCX$)
'''''''       NOCAMPO$(CANCA%) = Left$(NOCX$, PPXX% - 1)
'''''''       FORMX$(CANCA%) = Mid$(NOCX$, PPXX% + 1)
'''''''       If FORMX$(CANCA%) = "" Then FORMX$(CANCA%) = "A16"
'''''''       '
'''''''       If TIPODELISTA = "GRID" Then
'''''''          NOMBREDELISTA.Cols = CANCA%
'''''''       End If
'''''''       '
'''''''       If sqlx$ <> "" Then sqlx$ = sqlx$ + ", "
'''''''       sqlx$ = sqlx$ + NOCAMPO$(CANCA%) + " AS CAMPO" + TRIM$(Str$(CANCA%))
'''''''    Wend
'''''''    '
'''''''    TABLA1$ = TABLA$: TABLA2$ = "": IJO$ = ""
'''''''    PPXX% = InStr(TABLA$, "+")
'''''''    If PPXX% > 0 Then
'''''''       TABLA1$ = TRIM$(Left$(TABLA$, PPXX% - 1))
'''''''       TABLA$ = TRIM$(Mid$(TABLA$, PPXX% + 1))
'''''''       PPXX% = InStr(TABLA$, "/")
'''''''       If PPXX% > 0 Then
'''''''         TABLA2$ = TRIM$(Left$(TABLA$, PPXX% - 1))
'''''''         IJO$ = TRIM$(Mid$(TABLA$, PPXX% + 1))
'''''''         TABLA$ = TABLA1$ + " INNER JOIN " + TABLA2$ + " ON " + IJO$
'''''''       End If
'''''''    End If
'''''''
'''''''    If InStr(MODOSE$, "DISTINCT") > 0 Then sqlx$ = "DISTINCT " + sqlx$
'''''''    sqlx$ = "SELECT " + sqlx$ + " FROM " + TABLA$
'''''''    If TRIM$(CONDICION$) <> "" Then
'''''''       sqlx$ = sqlx$ + " WHERE " & CONDICION$
'''''''    End If
'''''''    On Error Resume Next
'''''''    Set TABTMP = FLEXCONN.Execute(FILTSQL$(sqlx$))
'''''''    If Err Then
'''''''       On Error GoTo 0
'''''''       MsgBox "Imposible Ejecutar Consulta SQL: " + vbCrLf + vbCrLf & sqlx$, vbExclamation
'''''''       Call FINAL("")
'''''''    End If
'''''''    '
'''''''    With TABTMP
'''''''       'CARGA EL LISTADO CON LAS COTIZACIONES YA INGRESADAS
'''''''       TTXX$ = ""
'''''''       j% = 0: FILA& = 0
'''''''       Do While Not .EOF
'''''''           TTXX$ = "": COLUI% = 0
'''''''           '
'''''''           If TIPODELISTA = "GRID" Then
'''''''              FILA& = FILA& + 1
'''''''37            If FILA& > NOMBREDELISTA.Rows - 1 Then
'''''''                 NOMBREDELISTA.Rows = NOMBREDELISTA.Rows + 1
'''''''                 GoTo 37
'''''''              End If
'''''''           End If
'''''''           '
'''''''           For U& = 1 To .Fields.Count
'''''''              VALOR = .Fields(U& - 1).Value
'''''''              If FORMX$(CANCA%) = "" Then FORMX$(CANCA%) = "A16" = FORMX$(U&)
'''''''              TECO$ = ""
'''''''              FORMATX$ = FORMX$(CANCA%)
'''''''              PPXX% = InStr(FORMATX$, "+")
'''''''              If PPXX% > 0 Then
'''''''                 TECO$ = TRIM$(Mid$(FORMATX$, PPXX% + 1))
'''''''                 FORMATX$ = Left$(FORMATX$, PPXX% - 1)
'''''''              End If
'''''''              '
'''''''              ENCAB$ = ""
'''''''              PPXX% = InStr(FORMATX$, "/")
'''''''              If PPXX% > 0 Then
'''''''                 ENCAB$ = Mid$(FORMATX$, PPXX% + 1)
'''''''                 FORMATX$ = Left$(FORMATX$, PPXX% - 1)
'''''''              End If
'''''''              TTYY$ = FORMATNUM$(VALOR, FORMATX$)
'''''''              If TRIM$(TTYY$) = "" Then
'''''''                  TTYY$ = Space$(XVALO(Mid$(FORMATX$, 2)))
'''''''              End If
'''''''              If TIPODELISTA = "LIST" Then
'''''''                   If TTXX$ <> "" Then TTXX$ = TTXX$ + " "
'''''''                   TTXX$ = TTXX$ + TTYY$
'''''''                 ElseIf TIPODELISTA = "GRID" Then
'''''''                   COLUI% = COLUI% + 1
'''''''17                 If COLUI% > NOMBREDELISTA.Cols Then
'''''''                      NOMBREDELISTA.Cols = NOMBREDELISTA.Cols + 1
'''''''                      GoTo 17
'''''''                   End If
'''''''                   If FILA& = 1 Then
'''''''                      FORMALTER.TEPRUEBA.Font = NOMBREDELISTA.Font
'''''''                      FORMALTER.TEPRUEBA.FontSize = NOMBREDELISTA.FontSize
'''''''                      FORMALTER.TEPRUEBA.Caption = "l" + TMASCARA$(FORMATX$)
'''''''                      NOMBREDELISTA.ColWidth(COLUI% - 1) = FORMALTER.TEPRUEBA.Width
'''''''                      NOMBREDELISTA.ColAlignment(COLUI% - 1) = 6
'''''''                      If Left$(FORMATX$, 1) = "A" Or Left$(FORMATX$, 1) = "C" Or Left$(FORMATX$, 1) = "D" Then
'''''''                         NOMBREDELISTA.ColAlignment(COLUI% - 1) = 1
'''''''                      End If
'''''''                      NOMBREDELISTA.FontBold = True
'''''''                      NOMBREDELISTA.TextMatrix(0, COLUI% - 1) = ENCAB$
'''''''                      NOMBREDELISTA.FontBold = False
'''''''                   End If
'''''''                   If Left$(FORMATX$, 1) = "A" Then TTYY$ = VALOR       ' PARA QUE EN LA GRILLA LO PONGA COMPLETO
'''''''                   NOMBREDELISTA.TextMatrix(FILA&, COLUI% - 1) = TTYY$
'''''''              End If
'''''''              '
'''''''              If TECO$ <> "" Then
'''''''                 PPXX% = InStr(TECO$, "/")
'''''''                 If PPXX% < 1 Then PPXX% = 1 + Len(TECO$)
'''''''                 FORMATX$ = Mid$(TECO$, PPXX% + 1)
'''''''                 TECO$ = TRIM$(UCase$(Left$(TECO$, PPXX% - 1)))
'''''''                 '
'''''''                 ENCAB$ = ""
'''''''                 PPXX% = InStr(FORMATX$, "/")
'''''''                 If PPXX% > 0 Then
'''''''                    ENCAB$ = Mid$(FORMATX$, PPXX% + 1)
'''''''                    FORMATX$ = Left$(FORMATX$, PPXX% - 1)
'''''''                 End If
'''''''                 If FORMATX$ = "" Then FORMATX$ = "A16"
'''''''                 '
'''''''                 VALOECO = Null
'''''''                 If Left$(TECO$, 5) = "DEUDO" Or Left$(TECO$, 5) = "CLIEN" Then
'''''''                   VALOECO = RASOCLI$(XVALO(VALOR))
'''''''                 ElseIf Left$(TECO$, 5) = "PROVE" Or Left$(TECO$, 5) = "ACRED" Then
'''''''                   VALOECO = RASOCLI$((-1) * XVALO(VALOR))
'''''''                 ElseIf Left$(TECO$, 5) = "MASTE" Then
'''''''                   TXVALO$ = VALOR
'''''''                     If Len(TXVALO$) > 4 Then
'''''''                          CIXI% = CODINT%(TXVALO$)
'''''''                        Else
'''''''                          CIXI% = XVALO(VALOR)
'''''''                     End If
'''''''                   VALOECO = DESCRIT0$(CIXI%)
'''''''                 ElseIf EXISTE%(LUDAT$ + TECO$ + ".RFS") Then
'''''''                   LFF% = LENFIELD%(TECO$, 1)
'''''''                   ABUS$ = MKS$(XVALO(VALOR))
'''''''                   On Error Resume Next
'''''''                   If LFF% = 2 Then ABUS$ = MKI$(Int(XVALO(VALOR)))
'''''''                   If LFF% = 1 Then ABUS$ = Chr$(Int(XVALO(VALOR)))
'''''''                   If LFF% > 4 Then ABUS$ = TRIM$(VALOR) 'ESTO LO AGREGUE POR QUE SI ABUS$ ES TEXTO NO LO CONSIDERA
'''''''                   On Error GoTo 0
'''''''                   VALOECO = ECOARCH$(TECO$, ABUS$)
'''''''                 End If
'''''''                 '
'''''''                 TTYY$ = FORMATNUM$(VALOECO, FORMATX$)
'''''''                 If TRIM$(TTYY$) = "" Then
'''''''                     TTYY$ = Space$(XVALO(Mid$(FORMATX$, 2)))
'''''''                 End If
'''''''                 If TIPODELISTA = "LIST" Then
'''''''                      If TTXX$ <> "" Then TTXX$ = TTXX$ + " "
'''''''                      TTXX$ = TTXX$ + TTYY$
'''''''                    ElseIf TIPODELISTA = "GRID" Then
'''''''                      COLUI% = COLUI% + 1
'''''''27                    If COLUI% > NOMBREDELISTA.Cols Then
'''''''                         NOMBREDELISTA.Cols = NOMBREDELISTA.Cols + 1
'''''''                         GoTo 27
'''''''                      End If
'''''''                      If FILA& = 1 Then
'''''''                         FORMALTER.TEPRUEBA.Font = NOMBREDELISTA.Font
'''''''                         FORMALTER.TEPRUEBA.FontSize = NOMBREDELISTA.FontSize
'''''''                         FORMALTER.TEPRUEBA.Caption = "l" + TMASCARA$(FORMATX$)
'''''''                         NOMBREDELISTA.ColWidth(COLUI% - 1) = FORMALTER.TEPRUEBA.Width
'''''''                         NOMBREDELISTA.ColAlignment(COLUI% - 1) = 6
'''''''                         If Left$(FORMATX$, 1) = "A" Or Left$(FORMATX$, 1) = "C" Or Left$(FORMATX$, 1) = "D" Then
'''''''                            NOMBREDELISTA.ColAlignment(COLUI% - 1) = 1
'''''''                         End If
'''''''                         NOMBREDELISTA.FontBold = True
'''''''                         NOMBREDELISTA.TextMatrix(0, COLUI% - 1) = ENCAB$
'''''''                         NOMBREDELISTA.FontBold = False
'''''''                      End If
'''''''                      If Left$(FORMATX$, 1) = "A" Then TTYY$ = VALOECO       ' PARA QUE EN LA GRILLA LO PONGA COMPLETO
'''''''                      NOMBREDELISTA.TextMatrix(FILA&, COLUI% - 1) = TTYY$
'''''''                 End If
'''''''                 '
'''''''              End If
'''''''           Next U&
'''''''           '
'''''''           If TIPODELISTA = "LIST" Then
'''''''               NOMBREDELISTA.AddItem TTXX$
'''''''           End If
'''''''           '
'''''''       .MoveNext
'''''''       Loop
'''''''    End With
'''''''    '
'''''''    If TIPODELISTA = GRID Then
'''''''    End If
'''''''    '
'''''''    TABTMP.Close
End Sub

Sub CARGALISEL(NOMBREDELISTA$, TABLA$, LISTADECAMPOS$, CONDICION$, Optional MODOSEL$)
    '
    Dim NOCAMPO$(), FORMX$()
    ReDim NOCAMPO$(1), FORMX$(1)
    '
    SCMA = Screen.MousePointer
    Screen.MousePointer = 11
    Call ABRECONEXION
    '
    MODOSE$ = UCase$(SAFETEXT$(MODOSEL$))
    LSTX$ = TRIM$(LISTADECAMPOS$)
    CANCA% = 0
    SQLX$ = ""
    ATOTA% = 1
    While LSTX$ <> ""
       PPXX% = InStr(LSTX$, ";")
       PPXX2% = InStr(LSTX$, ",")
       If PPXX2% > 0 And PPXX2% < PPXX% Then PPXX% = PPXX2%
       If PPXX% < 1 Then PPXX% = 1 + Len(LSTX$)
       '
       CANCA% = CANCA% + 1
       If CANCA% >= UBound(NOCAMPO$) Then
          ReDim Preserve NOCAMPO$(CANCA%), FORMX$(CANCA%)
       End If
       NOCX$ = Left$(LSTX$, PPXX% - 1)
       LSTX$ = Mid$(LSTX$, PPXX% + 1)
       '
       PPXX% = InStr(NOCX$, "/")
       If PPXX% < 1 Then PPXX% = 1 + Len(NOCX$)
       NOCAMPO$(CANCA%) = Left$(NOCX$, PPXX% - 1)
       FORMX$(CANCA%) = Mid$(NOCX$, PPXX% + 1)
       If FORMX$(CANCA%) = "" Then FORMX$(CANCA%) = "A16"
       '
       'ATOTA% = ATOTA% + 1 + XVALO(Mid$(FORMX$(CANCA%), 2))
       ATOTA% = ATOTA% + XVALO(Mid$(FORMX$(CANCA%), 2)) + 2
       '
       TECO$ = ""
       PPXX% = InStr(FORMX$(CANCA%), "+")
       If PPXX% > 0 Then
          TECO$ = UCase$(TRIM$(Mid$(FORMX$(CANCA%), PPXX% + 1)))
          PPXX% = InStr(TECO$, "/")
          If PPXX% > 0 Then
             FORMATX$ = TRIM$(Mid$(TECO$, PPXX% + 1))
             ATOTA% = ATOTA% + Len(TMASCARA$(FORMATX$))
          End If
       End If
       
       If CANCA% = 1 Then ACOL1% = XVALO(Mid$(FORMX$(CANCA%), 2))
       '
       If SQLX$ <> "" Then SQLX$ = SQLX$ + ", "
       SQLX$ = SQLX$ + NOCAMPO$(CANCA%) + " AS CAMPO" + TRIM$(Str$(CANCA%))
    Wend
    '
    ACOL2% = ATOTA% - 2 - ACOL1%
    '
    NOLISTA$ = NOMBREDELISTA$
    ESLOCAL% = 0
    If Left$(NOLISTA$, 1) = "!" Then
       NOLISTA$ = Mid$(NOLISTA$, 2)
       ESLOCAL% = 1
    End If
    '
    
    TABLA1$ = TABLA$: TABLA2$ = "": IJO$ = ""
    PPXX% = InStr(TABLA$, "+")
    If PPXX% > 0 Then
       TABLA1$ = TRIM$(Left$(TABLA$, PPXX% - 1))
       TABLA$ = TRIM$(Mid$(TABLA$, PPXX% + 1))
       PPXX% = InStr(TABLA$, "/")
       If PPXX% > 0 Then
         TABLA2$ = TRIM$(Left$(TABLA$, PPXX% - 1))
         IJO$ = TRIM$(Mid$(TABLA$, PPXX% + 1))
         TABLA$ = TABLA1$ + " INNER JOIN " + TABLA2$ + " ON " + IJO$
       End If
    End If
    '
    
    If InStr(MODOSE$, "DISTINCT") > 0 Then SQLX$ = "DISTINCT " + SQLX$
    SQLX$ = "SELECT " + SQLX$ + " FROM " + TABLA$
'    'VALIDACION QUE SE PASE EN LA CONSULTA WITH NOLOCK
'    If InStr(UCase$(SQLX$), "WITH(NOLOCK)") < 1 Then SQLX$ = SQLX$ + " WITH(NOLOCK) "
    '
    If TRIM$(CONDICION$) <> "" Then
       SQLX$ = SQLX$ + " WHERE " & CONDICION$
    End If
    '
    On Error Resume Next
    Set TABTMP = FLEXCONN.Execute(FILTSQL$(SQLX$))
   
'    Set TABTMP = READSET(SQLX$)
    If Err Then
       MsgBox Err.Description
       On Error GoTo 0
       MsgBox "Imposible Ejecutar Consulta SQL: " + vbCrLf + vbCrLf & SQLX$, vbExclamation
       Call FINAL("")
    End If
    '
    PPXX% = InStr(UCase$(CONDICION$), "ORDER BY"): If PPXX% < 1 Then PPXX% = 1 + Len(CONDICION$)
    CONDICI1$ = TRIM$(Left$(CONDICION$, PPXX% - 1))
    PPXX% = InStr(UCase$(CONDICI1$), "GROUP BY"): If PPXX% < 1 Then PPXX% = 1 + Len(CONDICI1$)
    CONDICAN$ = TRIM$(Left$(CONDICI1$, PPXX% - 1))
    '
    CANREGI& = CantRegistros(TABLA$, CONDICAN$)
    TELIS$ = Space$(CANREGI& * ATOTA%)
    POSIN& = 1
    '
    If InStr(MODOSE$, "NOCLEAR") > 0 Then
       If Left$(NOMBREDELISTA$, 1) = "!" Then
           TELISA$ = LOADFILE$(LUCOM$ + Mid$(NOMBREDELISTA$, 2) + ".SEL")
           GoTo 15
       End If
       Call ABRECONTROL
       QTR& = CantRegistros("SQL_LISELEC", "NomLiSele = '" & NOMBREDELISTA$ & "'")
       If QTR& > 0 Then
          TELISA$ = VALOBADA("SQL_LISELEC", "LiSelec", "NomLiSele = '" & NOMBREDELISTA$ & "'")
15        TELIS$ = TELISA$ + TELIS$
          POSIN& = 1 + Len(TELISA$)
       End If
    End If
    '
    
    With TABTMP
       'CARGA EL LISTADO CON LAS COTIZACIONES YA INGRESADAS
       TTXX$ = ""
       J% = 0: KKU& = 0
       Do While Not .EOF
           TTXX$ = ""
           KKU& = KKU& + 1
           Call TRACE(20, KKU&, CANREGI&)
           For U& = 1 To .Fields.Count
              valor = .Fields(U& - 1).Value
              FORMATX$ = FORMX$(U&)
              TECO$ = ""
              PPXX% = InStr(FORMATX$, "+")
              If PPXX% > 0 Then
                 TECO$ = UCase$(TRIM$(Mid$(FORMATX$, PPXX% + 1)))
                 FORMATX$ = TRIM$(Left$(FORMATX$, PPXX% - 1))
              End If
              TTYY$ = FORMATNUM$(valor, FORMATX$)
              If TRIM$(TTYY$) = "" Then
                  TTYY$ = Space$(XVALO(Mid$(FORMATX$, 2)))
              End If
              If TTXX$ <> "" And TTYY$ <> "" Then TTXX$ = TTXX$ + " "
              TTXX$ = TTXX$ + TTYY$
              If U& = 1 Then TTXX$ = TTXX$ + " "   ' PARA QUE QUEDEN 2 ESPACIOS
              '
              If TECO$ <> "" Then
                 PPXX% = InStr(TECO$, "/")
                 If PPXX% < 1 Then PPXX% = 1 + Len(TECO$)
                 FORMATX$ = Mid$(TECO$, PPXX% + 1)
                 TECO$ = TRIM$(UCase$(Left$(TECO$, PPXX% - 1)))
                 If FORMATX$ = "" Then FORMATX$ = "A16"
                 '
                 VALOECO = Null
                 If Left$(TECO$, 5) = "DEUDO" Or Left$(TECO$, 5) = "CLIEN" Then
                   VALOECO = rasocli$(XVALO(valor))
                 ElseIf Left$(TECO$, 5) = "PROVE" Or Left$(TECO$, 5) = "ACRED" Then
                   VALOECO = rasocli$((-1) * XVALO(valor))
                 ElseIf Left$(TECO$, 5) = "MASTE" Then
                   TXVALO$ = valor
                     If Len(TXVALO$) > 4 Then
                          CIXI% = CODINT%(TXVALO$)
                        Else
                          CIXI% = XVALO(valor)
                     End If
                   VALOECO = DESCRIT0$(CIXI%)
                 ElseIf EXISTE%(LUDAT$ + TECO$ + ".RFS") Then
                   ABUS$ = valor
                   If TIPFIELD%(TECO$, 1) <> 6 Then
                        LFF% = LENFIELD%(TECO$, 1)
                        ABUS$ = MKS$(XVALO(valor))
                        On Error Resume Next
                        If LFF% = 2 Then ABUS$ = MKI$(Int(XVALO(valor)))
                        If LFF% = 1 Then ABUS$ = Chr$(Int(XVALO(valor)))
                        On Error GoTo 0
                   End If
                   VALOECO = ECOARCH$(TECO$, ABUS$)
                 End If
                 '
                 TTYY$ = FORMATNUM$(VALOECO, FORMATX$)
                 If TRIM$(TTYY$) = "" Then
                     TTYY$ = Space$(XVALO(Mid$(FORMATX$, 2)))
                 End If
                 If TTXX$ <> "" Then TTXX$ = TTXX$ + " "
                 TTXX$ = TTXX$ + TTYY$
                 '
              End If
           Next U&
           '
           Mid$(TELIS$, POSIN&, ATOTA%) = TTXX$
           POSIN& = POSIN& + ATOTA%
           '
       .MoveNext
       Loop
    End With
    
    TABTMP.Close
    '
    TELIS$ = Left$(TELIS$, POSIN& - 1)        ' PARA RECORTAR LOS BLANCOS AL FINAL
    '
    Call ABRECONTROL
    QTR& = CantRegistros("SQL_LISELEC", "NomLiSele = '" & NOMBREDELISTA$ & "'")
    
    '
    GRABAR_ARCHIVO% = 0
    Dim RS As ADODB.Recordset
    Set RS = New ADODB.Recordset
    CONDI$ = "Select * from SQL_LISELEC where NomLiSele = '" & NOMBREDELISTA$ & "'"
25  On Error Resume Next
    RS.Open FILTSQL$(CONDI$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
    If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
       On Error GoTo 0
       Call CapturaErrorOpen
       GoTo 25
    End If
    On Error GoTo 0
    With RS
         If QTR& < 1 Then
           .AddNew
           !NomLiSele = NOMBREDELISTA$
         End If
         On Error Resume Next
         !CODIEMPR = CodiEmp%
         On Error GoTo 0
         !ACol_1 = ACOL1%
         !ACol_2 = ACOL2%
         !Tabla_Origen = TABLA$
         !Lista_Campos = LISTADECAMPOS$
         !Condi_Selec = CONDICION$
         GRABAR_ARCHIVO% = 0
         If Left$(NOMBREDELISTA$, 1) = "!" Then
              GRABAR_ARCHIVO% = 1
              Call SAVEFILE(LUCOM$ + Mid$(NOMBREDELISTA$, 2) + ".SEL", TELIS$)
            Else
              !Fech_Actu = Now 'AHORA$
              !LiSelec = TELIS$
          End If
         .Update
    End With
    RS.Close
    Set RS = Nothing
    If GRABAR_ARCHIVO% = 1 Then
       Call SAVEFILE(LUCOM$ + Mid$(NOMBREDELISTA$, 2) + ".SEL", TELIS$)
    End If
'''''    CODIGO ORIGINAL
'''''    If QTR& < 1 Then
'''''        Call CreaRegistro("SQL_LISELEC", "NomLiSele", NOMBREDELISTA$)
'''''    End If
'''''    Call ACTUREGISTRO("SQL_LISELEC", "ACol_1", "NomLiSele = '" & NOMBREDELISTA$ & "'", ACOL1%, RAFE&)
'''''    Call ACTUREGISTRO("SQL_LISELEC", "ACol_2", "NomLiSele = '" & NOMBREDELISTA$ & "'", ACOL2%, RAFE&)
'''''    Call ACTUREGISTRO("SQL_LISELEC", "Tabla_Origen", "NomLiSele = '" & NOMBREDELISTA$ & "'", TABLA$, RAFE&)
'''''    Call ACTUREGISTRO("SQL_LISELEC", "Lista_Campos", "NomLiSele = '" & NOMBREDELISTA$ & "'", LISTADECAMPOS$, RAFE&)
'''''    Call ACTUREGISTRO("SQL_LISELEC", "Condi_Selec", "NomLiSele = '" & NOMBREDELISTA$ & "'", CONDICION$, RAFE&)
'''''    If Left$(NOMBREDELISTA$, 1) = "!" Then
'''''          Call SAVEFILE(LUCOM$ + Mid$(NOMBREDELISTA$, 2) + ".SEL", TELIS$)
'''''        Else
'''''          Call ACTUREGISTRO("SQL_LISELEC", "Fech_Actu", "NomLiSele = '" & NOMBREDELISTA$ & "'", AHORA$, RAFE&)
'''''          Call ACTUREGISTRO("SQL_LISELEC", "LiSelec", "NomLiSele = '" & NOMBREDELISTA$ & "'", TELIS$, RAFE&)
'''''    End If
    '
    Call FRESHECHO(NOMBREDELISTA$, "NORECAR")
    '
    If EXISTE%(LUDAT$ + NOLISTA$ + ".RFS") > 0 Then
       TVX1% = TIPFIELD%(NOLISTA$, 1): TVX2% = TIPFIELD%(NOLISTA$, 2)
       LFX1% = LENFIELD%(NOLISTA$, 1): LFX2% = LENFIELD%(NOLISTA$, 2)
       PRX1% = POSFIELD%(NOLISTA$, 1): PRX2% = POSFIELD%(NOLISTA$, 2)
       ATOSEL% = RECLEN%(NOLISTA$)
       REBLA$ = REGBLAN$(NOLISTA$)
       AREBLA& = Len(REBLA$)
       '
       TEDAT$ = String$(CANREGI& * AREBLA&, 0)
       POSIN& = 1
       jj& = 0
       FIN& = Len(TELIS$)
       '
       For U& = 1 To FIN& Step ATOTA
         Call TRACE(20, U&, FIN&)
         TTXX$ = Mid$(TELIS$, U&, ATOTA%)
         
         VXX1$ = Left$(TTXX$, ACOL1%)
         VXX2$ = Mid$(TTXX$, ACOL1% + 2)
         '
         TTYY$ = REBLA$
         Call REPLA(TTYY$, CBIN$(VXX1$, TVX1%, LFX1%), 1, LFX1%)
         Call REPLA(TTYY$, AJUSTI$(VXX2$, LFX2%), 1 + LFX1%, LFX2%)
         jj& = jj& + 1
         ' Call GRAREG(NOMBREDELISTA$, TTYY$, JJ&)   ' CORRECTO ASI LA
         On Error Resume Next
         
         Mid$(TEDAT$, POSIN&, AREBLA&) = TTYY$
         POSIN& = POSIN& + AREBLA&
         On Error GoTo 0
       Next U&                                     ' GRABA EN WKF
       '
       Call CIERRARCH(NOMBREDELISTA$)
       If Left$(NOMBREDELISTA$, 1) <> "!" Then
            Call SAVEFILE(LUDAT$ + NOMBREDELISTA$ + ".DAT", TEDAT$)
          Else
            Call SAVEFILE(LUCOM$ + Mid$(NOMBREDELISTA$, 2) + ".WKF", TEDAT$)
            
       End If
    End If
    '
    
    Screen.MousePointer = SCMA
    '
End Sub

Sub BORRALISEL(NOMBREDELISTA$)
    Call BORRAREGISTROS("SQL_LISELEC", "NomLiSele = '" & NOMBREDELISTA$ & "'", RAFE&, "NOMESS")
End Sub
'
Sub BORRAREGISTROS(TABLA$, CONDI$, REGISTROSAFECTADOS&, Optional MODO$)
   '
   NUMREG& = CantRegistros&(TABLA$, CONDI$)
   If NUMREG& > 0 Then
      '
      SQLX$ = "DELETE FROM " + TABLA$
      SQLX$ = SQLX$ & " WHERE " + CONDI$
      '
      On Error Resume Next
       FLEXCONN.Execute (FILTSQL$(SQLX$))
      If Err Then
         On Error GoTo 0
         If MODO$ <> "NOMESS" Then
           Call MENSERR(24, " Imposible Ejecutar Eliminacion:\  \'" + SQLX$ + "'")
         End If
      End If
   End If
   REGISTROSAFECTADOS& = NUMREG&
   '
End Sub

Sub ACONSUL_PRE(FORNAME$)
    '
    
    Call COPYSTRU("INDIASIE", "INCONSQL")
    Call ABRETACON
    jj& = 0
    '
    strSQL$ = "SELECT COD_QUERY FROM SQL_MENUS WITH(NOLOCK) WHERE Form_Name = '" & FORNAME$ & "' ORDER BY Num_Orden Asc "
    Set MIRS = READSET(strSQL$)
    'On Error Resume Next
    'Set MIRS = BaDaConsul.OpenRecordset(strSQL, dbOpenSnapshot)
   ' MIRS.MoveFirst
    If MIRS.EOF Then
       Call MENSERR(24, "No Hay Consultas Preconfiguradas\para este Formulario.")
       MIRS.Close: Set MIRS = Nothing
       GoTo 80
    End If
    Do While Not MIRS.EOF
       COCU$ = MIRS!COD_QUERY
       'DECU$ = ECOARCH$("INCONSQL", COCU$)
       CONDI$ = "COD_QUERY = '" & COCU$ & "'"
       DECU$ = VALOBADA("SQL_QUERIES", "DES_QUERY", CONDI$, "NOMESS")
       TTXX$ = AJUSTI$(COCU$, 16) + DECU$
       jj& = jj& + 1
       Call GRAREG("!INDIASIE", TTXX$, jj&)
    MIRS.MoveNext
    Loop
    MIRS.Close: Set MIRS = Nothing
    GoTo 90 ' SI LLEGO ACA SALTEA LA CONSULTA MDB

80  strSQL$ = "SELECT COD_QUERY FROM SQL_MENUS  WHERE Form_Name = '" & FORNAME$ & "' ORDER BY Num_Orden Asc "
    Call ABRETACON
    On Error Resume Next
    Set MIRS = BaDaConsul.OpenRecordset(strSQL, dbOpenSnapshot)
    MIRS.MoveFirst
    If Err Then
       On Error GoTo 0
       Call MENSERR(24, "No Hay Consultas Preconfiguradas\para este Formulario.")
       MIRS.Close: Set MIRS = Nothing
       GoTo 99
    End If
    Do While Not MIRS.EOF
       On Error GoTo 0
       COCU$ = MIRS!COD_QUERY
       'DECU$ = ECOARCH$("INCONSQL", COCU$)
       CONDI$ = "COD_QUERY = '" & COCU$ & "'"
       DECU$ = VALOTACO("SQL_QUERIES", "DES_QUERY", CONDI$)
       TTXX$ = AJUSTI$(COCU$, 16) + DECU$
       jj& = jj& + 1
       Call GRAREG("!INDIASIE", TTXX$, jj&)
    MIRS.MoveNext
    Loop
    MIRS.Close: Set MIRS = Nothing


90  Call SETULTREG("!INDIASIE", jj&)
    Call FRESHECHO("!INDIASIE")
    '
10  RSEL$ = REGSEL$("!INDIASIE/Consultas y Reportes Preconfigurados")
    If Len(RSEL$) > 4 Then
       MODUCONE$ = "?" + TRIM$(Left$(RSEL$, 16))
       DECONEC$ = TRIM$(Mid$(RSEL$, 17, 48))
       Call CONECRUN(MODUCONE$, DECONEC$)
       GoTo 10
    End If
    '
99 End Sub

Sub CONECTA_CONSUL_PRE(FORNAME$)
    '
    If (USNBR% Mod 100) <> 1 Then
       Call MENSERR(24, "Re-Configuración de Accesos\a Consultas Pre-Configuradas\Requiere Privilegios de Supervisor.")
       Exit Sub
    End If
    '
    Call COPYSTRU("INDIASIE", "INCONSQL")
 '**************************************
    strSQL$ = "SELECT COD_QUERY, DES_QUERY FROM SQL_QUERIES WITH(NOLOCK) ORDER BY COD_QUERY ASC"
    Set MIRS = READSET(strSQL$)
    If MIRS.EOF Then
       GoTo 40
    End If
    '
    jj& = 0
    On Error GoTo 0
    If Not (MIRS.EOF And MIRS.BOF) Then
       Do While Not MIRS.EOF
         TTXX$ = AJUSTI$(SAFETEXT$(MIRS!COD_QUERY), 16) + SAFETEXT$(MIRS!DES_QUERY)
         jj& = jj& + 1
         Call GRAREG("INCONSQL", TTXX$, jj&)
       MIRS.MoveNext
       Loop
       Call SETULTREG("INCONSQL", jj&)
    End If
    MIRS.Close: Set MIRS = Nothing
    
    jj& = 0
    strSQL$ = "SELECT COD_QUERY FROM SQL_MENUS WITH(NOLOCK) WHERE Form_Name = '" & FORNAME$ & "' ORDER BY Num_Orden Asc "
    Set MIRS = READSET(strSQL$)
    If MIRS.EOF Then
       GoTo 25
    End If
    Do While Not MIRS.EOF
       COCU$ = SAFETEXT$(MIRS!COD_QUERY)
       DECU$ = ECOARCH$("INCONSQL", COCU$)
       TTXX$ = AJUSTI$(COCU$, 16) + DECU$
       jj& = jj& + 1
       Call GRAREG("!INDIASIE", TTXX$, jj&)
    MIRS.MoveNext
    Loop
25  MIRS.Close: Set MIRS = Nothing
    Call SETULTREG("!INDIASIE", jj&)
    If jj& > 0 Then GoTo 60 ' SI TOMO INFORMACION DEL SQL NO LEE EL MDB
    
'***************************************
    
    
    
40  Call ABRETACON
    strSQL = "SELECT COD_QUERY, DES_QUERY FROM SQL_QUERIES ORDER BY COD_QUERY ASC"
    On Error Resume Next
    Set MIRS = BaDaConsul.OpenRecordset(strSQL$, dbOpenSnapshot)
    MIRS.MoveFirst
    If Err Then
       On Error GoTo 0
       Call MENSERR(24, "No se Encontraron Consultas Preconfiguradas.")
       Exit Sub
    End If
    '
    jj& = 0
    On Error GoTo 0
    If Not (MIRS.EOF And MIRS.BOF) Then
       Do While Not MIRS.EOF
         TTXX$ = AJUSTI$(SAFETEXT$(MIRS!COD_QUERY), 16) + SAFETEXT$(MIRS!DES_QUERY)
         jj& = jj& + 1
         Call GRAREG("INCONSQL", TTXX$, jj&)
       MIRS.MoveNext
       Loop
       Call SETULTREG("INCONSQL", jj&)
    End If
    '
    jj& = 0
    strSQL$ = "SELECT COD_QUERY FROM SQL_MENUS WHERE Form_Name = '" & FORNAME$ & "' ORDER BY Num_Orden Asc "
    On Error Resume Next
    Set MIRS = BaDaConsul.OpenRecordset(strSQL$, dbOpenSnapshot)
    MIRS.MoveFirst
    If Err Then
       On Error GoTo 0
       GoTo 45
    End If
    On Error GoTo 0
    Do While Not MIRS.EOF
       COCU$ = SAFETEXT$(MIRS!COD_QUERY)
       DECU$ = ECOARCH$("INCONSQL", COCU$)
       TTXX$ = AJUSTI$(COCU$, 16) + DECU$
       jj& = jj& + 1
       Call GRAREG("!INDIASIE", TTXX$, jj&)
    MIRS.MoveNext
    Loop
45  MIRS.Close
    Call SETULTREG("!INDIASIE", jj&)
    '
60  VTB% = VENTABU%("ACONPRE")
    If VTB% >= 0 Then
    '**************************************
       FLEXCONN.Execute ("DELETE FROM SQL_MENUS WHERE Form_Name = '" & FORNAME$ & "'")
       Set MIRS = New ADODB.Recordset
       strSQL$ = "SELECT * FROM SQL_MENUS WHERE Form_Name = '" & FORNAME$ & "' ORDER BY Num_Orden Asc "
       'Set MIRS = BaDaConsul.OpenRecordset(strSQL$, dbOpenDynaset)
       MIRS.Open FILTSQL$(strSQL$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
       For J& = 1 To ULTREG&("!INDIASIE")
          XX$ = REGLEIDO$("!INDIASIE", J&)
          MIRS.AddNew   ' YA CONVERTIDO
              On Error Resume Next
          MIRS!CODIEMPR = CodiEmp%
              On Error GoTo 0
          MIRS!Form_Name = FORNAME$
          MIRS!COD_QUERY = TRIM$(Left$(XX$, 16))
          MIRS!Num_Orden = J&
          MIRS.Update
       Next J&
       MIRS.Close: Set MIRS = Nothing
    
    
    '**************************************
       Call ABRETACON
       BaDaConsul.Execute ("DELETE FROM SQL_MENUS WHERE Form_Name = '" & FORNAME$ & "'")
       strSQL$ = "SELECT * FROM SQL_MENUS WHERE Form_Name = '" & FORNAME$ & "' ORDER BY Num_Orden Asc "
       Set MIRS = BaDaConsul.OpenRecordset(strSQL$, dbOpenDynaset)
       For J& = 1 To ULTREG&("!INDIASIE")
          XX$ = REGLEIDO$("!INDIASIE", J&)
          MIRS.AddNew   ' YA CONVERTIDO
              On Error Resume Next
          MIRS!CODIEMPR = CodiEmp%
              On Error GoTo 0
          MIRS!Form_Name = FORNAME$
          MIRS!COD_QUERY = TRIM$(Left$(XX$, 16))
          MIRS!Num_Orden = J&
          MIRS.Update
       Next J&
       MIRS.Close
    End If
    Call CIERRARCH("*.*")
    '
End Sub


Sub CapturaErrorOpen()
      ' Viene de captura de error en los casos que no se realiza el OPEN correctamente
      ' Si cancela realiza un call final
      ' Si reintenta vuelve y continua
      '
      OPX1% = COMALTER%("Servidor Transitoriamente Ocupado.\No se Ha Completado la Transacción\Espere unos Segundos y Pulse Reintentar\\Reintentar\Cancelar")
      If OPX1% = 2 Then
         OPX2% = COMALTER%("¿ Realmente Desea Cancelar\la Transacción Iniciada ?\   \Pueden Producirse Resultados Impredecibles !!!\Se Recomienda Esperar Algunos Segundos y Reintentar.\\Reintentar Transacción\Cancelar")
         If OPX2% = 2 Then
            Call FINAL("")
         End If
      End If
End Sub

'
Sub REGAUDITRA(Tipo$, Descripcion$, NUMERO$)
    '
    Call ABRECONEXION
    Call CREATABLA_BCONTROL
    Call ABRECONEXION
    '
    Dim Audi As ADODB.Recordset
    Set Audi = New ADODB.Recordset
    CONDI$ = "Select * from AUDITRA where TipoTra = 'x'"
25  On Error Resume Next
    Audi.Open FILTSQL$(CONDI$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
    If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
       On Error GoTo 0
       Call CapturaErrorOpen
       GoTo 25
    End If
    On Error GoTo 0
    With Audi
         .AddNew   ' YA CONVERTIDO
         On Error Resume Next
         !CODIEMPR = CodiEmp%
         On Error GoTo 0
         !FeReTra = Now
         !TIPOTRA = Left(SAFETEXT(Tipo$), 8)
         !DescriTra = Left(SAFETEXT(Descripcion$), 64)
         !NuCompro = Left(SAFETEXT(NUMERO$), 24)
         !UserTra = Left(SAFETEXT(USERNAME$), 16)
         !IdTermi = Left(SAFETEXT(NOMTER$), 24)
         .Update
    End With
    Audi.Close
    Set Audi = Nothing
    
End Sub
'
Function FILTSQL$(STSQLX, Optional MODOMESS$)
    '
    FILTSQL$ = STSQLX
    STSQLX1$ = STSQLX
    NTABLAX$ = ""
    '
    If InStr(UCase$(FILTSQL$), "INNER JOIN ") > 0 Then GoTo 80 ' DEBE RESOLVERSE EN ORIGEN
    If InStr(UCase$(FILTSQL$), "LEFT JOIN ") > 0 Then GoTo 80 ' DEBE RESOLVERSE EN ORIGEN
    If InStr(UCase$(FILTSQL$), "RIGHT JOIN ") > 0 Then GoTo 80 ' DEBE RESOLVERSE EN ORIGEN
    If Left$(UCase$(TRIM$(FILTSQL$)), 11) = "SELECT NAME" Then GoTo 80
    If Left$(UCase$(TRIM$(FILTSQL$)), 22) = "SELECT SYSCOLUMNS.NAME" Then GoTo 80
    If InStr(UCase$(TRIM$(FILTSQL$)), "FROM SYSINDEXES") > 0 Then GoTo 80
    If InStr(UCase$(TRIM$(FILTSQL$)), "CODIEMPR IS NULL") > 0 Then GoTo 80
    '
    PPXX% = InStr(UCase$(STSQLX1$), "FROM")   ' CASO SELECT
       If PPXX% < 1 Then
          If InStr(UCase$(STSQLX1$), "DELETE") > 0 Then
             PPXX% = InStr(UCase$(STSQLX1$), "DELETE") + 2
          End If
       End If
       If PPXX% < 1 Then
          If InStr(UCase$(STSQLX1$), "UPDATE") > 0 Then
             PPXX% = InStr(UCase$(STSQLX1$), "UPDATE") + 2
          End If
       End If
       '
    If PPXX% > 0 Then
      PPYY% = InStr(PPXX% + 6, UCase$(STSQLX1$), " ")
        If PPYY% < PPXX% + 5 Then PPYY% = 1 + Len(STSQLX1$)
      NTABLA$ = UCase$(TRIM$(Mid$(STSQLX1$, PPXX% + 5, PPYY% - PPXX% - 5)))
      '
      ' unificadas
      If NTABLA$ = "SYSOBJECTS" Then GoTo 80
      If NTABLA$ = "MASTER" Then GoTo 80
      If NTABLA$ = "FORMULAS" Then GoTo 80
      If NTABLA$ = "FORMUPAT" Then GoTo 80
      If NTABLA$ = "SECOPER" Then GoTo 80
      If NTABLA$ = "PLANCUEN" Then GoTo 80
      If NTABLA$ = "PROVED12" Then GoTo 80
      If NTABLA$ = "DEUDOR12" Then GoTo 80
      If NTABLA$ = "BSEGCOM" Then GoTo 80
      If NTABLA$ = "LISTAPRE" Then GoTo 80
      If NTABLA$ = "PREVENTA" Then GoTo 80
      If NTABLA$ = "DESCOMBI" Then GoTo 80
      If NTABLA$ = "PRESUVEN" Then GoTo 80
      If NTABLA$ = "MOVISTO" Then GoTo 80
      If NTABLA$ = "OCOEMI" Then GoTo 80
      If NTABLA$ = "ENCOCABI" Then GoTo 80
      If NTABLA$ = "ENCASOLINT" Then GoTo 80
      If NTABLA$ = "DETASOLINT" Then GoTo 80
      If NTABLA$ = "COTIRECI" Then GoTo 80
      If NTABLA$ = "BOLRECEP" Then GoTo 80
      If NTABLA$ = "MOVIREPA" Then GoTo 80
      If NTABLA$ = "MACONSIG" Then GoTo 80
      If NTABLA$ = "TABEJCON" Then GoTo 80
      If NTABLA$ = "UNEGOCIOS" Then GoTo 80
      'If NTABLA$ = "ORDEFABR" Then GoTo 80
      If NTABLA$ = "OPERFAB" Then GoTo 80
      'If NTABLA$ = "RESERVA" Then GoTo 80
      If NTABLA$ = "INFOFAB" Then GoTo 80
      If NTABLA$ = "LOTINGRE" Then GoTo 80
      If NTABLA$ = "TANUMSE" Then GoTo 80
      If NTABLA$ = "CONTROLFIN" Then GoTo 80
      If NTABLA$ = "BALAREQMO" Then GoTo 80
      If NTABLA$ = "RECONCILIA" Then GoTo 80
      If NTABLA$ = "UBIPRE" Then GoTo 80
      If NTABLA$ = "NEBRUNE" Then GoTo 80
      If NTABLA$ = "COTELEF" Then GoTo 80
      If NTABLA$ = "CORTRAB" Then GoTo 80
      If NTABLA$ = "DETOTRA" Then GoTo 80
      If NTABLA$ = "TARINTER" Then GoTo 80
      If NTABLA$ = "CORRPROG" Then GoTo 80
      If NTABLA$ = "ENCAREMI" Then GoTo 80
      If NTABLA$ = "PROVECOD" Then GoTo 80
      If NTABLA$ = "DETACARTEM" Then GoTo 80
      If NTABLA$ = "ENSAR072" Then GoTo 80
      If NTABLA$ = "SEGELECT" Then GoTo 80
      If NTABLA$ = "ENSAR070" Then GoTo 80
      If NTABLA$ = "MEDIFIN" Then GoTo 80
      If NTABLA$ = "ENSAR074" Then GoTo 80
      If NTABLA$ = "PLADIASTO" Then GoTo 80
      If NTABLA$ = "INOCONF" Then GoTo 80
      If NTABLA$ = "PADRON" Then GoTo 80
      If NTABLA$ = "COMPXEMPLE" Then GoTo 80
      If NTABLA$ = "TAPROYEC" Then GoTo 80
      If NTABLA$ = "INFOPROY" Then GoTo 80
      If NTABLA$ = "SUBCONJUNTOS" Then GoTo 80
      If NTABLA$ = "AUDITRA" Then GoTo 80
      If NTABLA$ = "TRANSPORTES" Then GoTo 80
      If NTABLA$ = "SUCENTRE" Then GoTo 80
      If NTABLA$ = "OPORTUNI" Then GoTo 80
      If NTABLA$ = "EQUIPOS" Then GoTo 80
      If NTABLA$ = "GALEIMAG" Then GoTo 80
      If NTABLA$ = "REPUMAQ" Then GoTo 80
      If NTABLA$ = "CONTADORES" Then GoTo 80
      If NTABLA$ = "LISEMPAQ" Then GoTo 80
      If NTABLA$ = "ENCEMPAQ" Then GoTo 80
      If NTABLA$ = "HERRAMEN" Then GoTo 80
      If NTABLA$ = "HERRAOPER" Then GoTo 80
      If NTABLA$ = "AUDITACT" Then GoTo 80
      If NTABLA$ = "CONTACSCOM" Then GoTo 80
      If NTABLA$ = "REPARA" Then GoTo 80
      If NTABLA$ = "REPUREPA" Then GoTo 80
      If NTABLA$ = "BADAMODE" Then GoTo 80
      If NTABLA$ = "LIQUIDACOS" Then GoTo 80
      If NTABLA$ = "LIQUIRECEP" Then GoTo 80
      If NTABLA$ = "CARPEEXP" Then GoTo 80
      If NTABLA$ = "PLAMANPREV" Then GoTo 80
      If NTABLA$ = "DISPONIB" Then GoTo 80
      If NTABLA$ = "TIEMOBRA" Then GoTo 80
      '
      ' separadas
      If NTABLA$ = "SQL_LISELEC" Then GoTo 60
      If NTABLA$ = "CAJABANC" Then GoTo 60
      If NTABLA$ = "CABANVIE" Then GoTo 60
      If NTABLA$ = "SADEUPEN" Then GoTo 60
      If NTABLA$ = "SACREPEN" Then GoTo 60
      If NTABLA$ = "APLICOBRA" Then GoTo 60
      If NTABLA$ = "APLIPAGO" Then GoTo 60
      If NTABLA$ = "ENCABEASIEN" Then GoTo 60
      If NTABLA$ = "DETAASIEN" Then GoTo 60
      If NTABLA$ = "PEDENCA" Then GoTo 60
      If NTABLA$ = "PEDDETA" Then GoTo 60
      If NTABLA$ = "OCOMENCA" Then GoTo 60
      If NTABLA$ = "OCOMDETA" Then GoTo 60
      If NTABLA$ = "CARPEIMP" Then GoTo 60
      If NTABLA$ = "AUTOENT" Then GoTo 60
      If NTABLA$ = "DETAIVA" Then GoTo 60
      If NTABLA$ = "PDOC_TBL" Then GoTo 60
      If NTABLA$ = "PLAGASPEN" Then GoTo 60
      If NTABLA$ = "FLEXCRL" Then GoTo 60
      If NTABLA$ = "ORDEFABR" Then GoTo 60
      If NTABLA$ = "ORFAETI" Then GoTo 60
      If NTABLA$ = "RESERVA" Then GoTo 60
      If NTABLA$ = "DETFACOMSQL" Then GoTo 60
      '
      GoTo 80                               ' las que no se detallan van unificadas
'      If NTABLA$ = "" Then GoTo 80
'      If InStr(UCase$(MODOMESS$), "NOMESS") < 1 Then
'         MsgBox "Que Hago con Tabla " + NTABLA$
'      End If
'      '
60    PPXX% = InStr(UCase$(STSQLX1$), "WHERE")
      If PPXX% > 0 Then
          IZQUI$ = Left$(STSQLX1$, PPXX% + 5)
          DERE$ = Mid$(STSQLX1$, PPXX% + 5)
        ElseIf InStr(UCase$(STSQLX1$), "ORDER") > 0 Then
          PPXX1% = InStr(UCase$(STSQLX1$), "ORDER")
          IZQUI$ = Left$(STSQLX1$, PPXX1% - 1) + " WHERE "
          DERE$ = Mid$(STSQLX1$, PPXX1%)
        Else
          IZQUI$ = STSQLX1$ + " WHERE"
          DERE$ = ""
      End If
      FILTSQL$ = IZQUI$ + " CodiEmpr = " & CodiEmp%
      If DERE$ <> "" Then
         If Left$(TRIM$(UCase$(DERE$)), 5) <> "ORDER" Then
            FILTSQL$ = FILTSQL$ & " AND"
         End If
         FILTSQL$ = FILTSQL$ + " " + DERE$
      End If
      '
    End If
    '
80  End Function
'
Function READSET(SQXX$, Optional TIMUT, Optional SINNOLOCK%) As ADODB.Recordset
    '
    Call ABRECONEXION
    
    If XVALO(TIMUT) > 0 Then
       FLEXCONN.CommandTimeout = TIMUT
       FIN& = FLEXCONN.CommandTimeout
    End If
    '
10  Set READSET = New ADODB.Recordset

    If InStr(UCase$(SQXX$), "WITH(NOLOCK)") < 1 And SINNOLOCK% = 0 Then
       TTX1$ = "Error de Invocación en Función 'READSET' - Falta WITH(NOLOCK) En Consulta: " + vbCrLf
       TTX1$ = TTX1$ + SQXX$
       MsgBox TTX1$
       Call FINAL("")
    End If
    '
    SQYY$ = FILTSQL$(SQXX$)
    On Error Resume Next
    'Set READSET = FLEXCONN.Execute(SQYY$)
    READSET.CursorLocation = adUseClient
    READSET.CursorType = adOpenForwardOnly
    READSET.LockType = adLockReadOnly
    READSET.Open SQYY$, FLEXCONN, , , adCmdText
   
    If Err Then
        EDES$ = UCase$(Err.Description)
        On Error GoTo 0
        If InStr(EDES$, "TIMEOUT") > 0 Or InStr(EDES$, "ESPERA") > 0 Then
             OPX% = COMALTER%("Tiempo de Espera (" & FLEXCONN.CommandTimeout & ") Agotado.\Desea Re-Intentar Prolongando Tiempo de Espera ?\   \En Cualquier Caso Comunique Este Mensaje a FLEXOFT\\Si, Reintentar\No, Cancelar")
             If OPX% = 1 Then
                  FLEXCONN.CommandTimeout = 2 * FLEXCONN.CommandTimeout
                  FIN& = FLEXCONN.CommandTimeout
                  GoTo 10
                Else
                  Call FINAL("")
             End If
           Else
             TTXX$ = "Error no Recuperable al Ejecutar Consulta SQL: " + SQXX$ + vbCrLf + vbCrLf
             TTXX$ = TTXX$ + "El Mensaje de Error Informado Dice: " + EDES$ + vbCrLf + vbCrLf
             TTXX$ = TTXX$ + "Al Pulsar 'Aceptar' se Cerrará la Aplicación."
             MsgBox TTXX$
             Call FINAL("")
        End If
    End If
    '
    On Error GoTo 0
    '
End Function


Function EXISTELREP%(FORNAME$)
    '
    EXREP% = 0
    SQLX$ = "SELECT * FROM SQL_QUERIES WITH(NOLOCK) WHERE COD_QUERY = '" & FORNAME$ & "'"
    Set RS = READSET(SQLX$)
    With RS
      If Not .EOF Then
        EXREP% = 1
        GoTo 99
      End If
    End With
    RS.Close
    Set RS = Nothing
    
    Call ABRETACON
    '
    strSQL = "SELECT * FROM SQL_QUERIES WHERE COD_QUERY = '" & FORNAME$ & "'"
    On Error Resume Next
    Set MIRS = BaDaConsul.OpenRecordset(strSQL, dbOpenSnapshot)
    MIRS.MoveFirst
    If Err Then
       On Error GoTo 0
       GoTo 25
    End If
    On Error GoTo 0
    Do While Not MIRS.EOF
        EXREP% = 1
        GoTo 99
    MIRS.MoveNext
    Loop
25  MIRS.Close
    '
99  EXISTELREP% = EXREP%
    '
End Function
Function EXISCONSUMDB%(CODMO$)
        Call ABRETACON
        EXISCONSUMDB% = 0
        SQLX$ = "SELECT * FROM SQL_QUERIES  WHERE COD_QUERY = '" & CODMO$ & "'"
        Set RS = BaDaConsul.OpenRecordset(SQLX$, dbOpenSnapshot)
        With RS
          If Not (.BOF And .EOF) Then EXISCONSUMDB% = 1
        End With
        RS.Close
        Set RS = Nothing
End Function

Function CONTROL_SQL$(PROCE$, CLAVE$)
    '
    If YACRECON% < 1 Then
       Call CREACAMPO("", "FLEXCRL", "PROBUS", 10, 32, 1)
       Call CREACAMPO("", "FLEXCRL", "CLABUS", 10, 32, 1)
       Call CREACAMPO("", "FLEXCRL", "VALCONTROL", 12, 0, 0)
       YACRECON% = 1
    End If
    '
    CONDI$ = "PROBUS = '" & PROCE$ & "' AND CLABUS = '" & CLAVE$ & "'"
    CONTROL_SQL$ = VALOBADA("FLEXCRL", "VALCONTROL", CONDI$, "NOMESS")
    '
    If CONTROL_SQL$ = "" Then
        ' SI NO LO ENCUENTRA BUSCA EN LA TABLA DAT
        PROBUS$ = AJUSTI$(PROCE$, 8)
        If TRIM$(PROCE$) = "" Then PROBUS$ = ""
        CODBUS$ = AJUSTI$(CLAVE$, 8)
        '
        CONTROL_SQL$ = ""
        
        RFF$ = RECFILT$("FLEXCRL", 2, CODBUS$)
        For U& = 1 To Len(RFF$) Step 4
           REBU& = CVS(Mid$(RFF$, U&, 4))
           XI$ = REGLEIDO$("FLEXCRL", REBU&)
           If PROBUS$ = "" Or PROBUS$ = Left$(XI$, 8) Then
              CONTROL_SQL$ = TRIM$(Mid$(XI$, 17, 80))
              Exit For
           End If
        Next U&
    End If
    '
End Function
'
Sub GRACONTROL_SQL(PROCE$, CLAVE$, VALCON$)
    '
    If YACRECON% < 1 Then
       Call CREACAMPO("", "FLEXCRL", "PROBUS", 10, 32, 1)
       Call CREACAMPO("", "FLEXCRL", "CLABUS", 10, 32, 1)
       Call CREACAMPO("", "FLEXCRL", "VALCONTROL", 12, 0, 0)
       YACRECON% = 1
    End If
    '
    CONDI$ = "PROBUS = '" & PROCE$ & "' AND CLABUS = '" & CLAVE$ & "'"
    If CantRegistros("FLEXCRL", CONDI$, "NOMESS") < 1 Then
10     RNDX$ = RANDSTRING$(8)
       If CantRegistros("FLEXCRL", "PROBUS = '" & RNDX$ & "'", "NOMESS") > 0 Then GoTo 10
       Call CreaRegistro("FLEXCRL", "PROBUS", RNDX$, "NOMESS")
       Call ACTUREGISTRO("FLEXCRL", "CLABUS", "PROBUS = '" & RNDX$ & "'", CLAVE$, RAFE&, "NOMESS")
       Call ACTUREGISTRO("FLEXCRL", "PROBUS", "PROBUS = '" & RNDX$ & "'", PROCE$, RAFE&, "NOMESS")
    End If
    Call ACTUREGISTRO("FLEXCRL", "VALCONTROL", CONDI$, VALCON$, RAFE&, "NOMESS")
    '
End Sub
'
'Sub CARGALISTA_LIST_GRID(NOMBREDELISTA, TABLA$, LISTADECAMPOS$, CONDICION$, Optional MODOSEL$, Optional ENCABEZADO$, Optional COLUMNA_AGRUPA%)
'    '
'    Dim NOCAMPO$(), FORMX$()
'    ReDim NOCAMPO$(1), FORMX$(1)
'    '
'    MODOSE$ = TRIM$(UCase$(SAFETEXT$(MODOSEL$)))
'    '
'    '
'    TIPODELISTA = "GRID"
'    On Error Resume Next
'    CANCOL& = NOMBREDELISTA.Cols
'    If Err Then
'       On Error GoTo 0
'       TIPODELISTA = "LIST"
'    End If
'    On Error GoTo 0
'    '
'    If TIPODELISTA = "LIST" Then
'        On Error Resume Next
'        NOMBREDELISTA.Clear
'        If Err Then
'           On Error GoTo 0
'           Call MENSERR(24, "Imposible Limpiar Lista '" & NOMBREDELISTA & "'.")
'           Call FINAL("")
'        End If
'        On Error GoTo 0
'    ElseIf TIPODELISTA = "GRID" Then
'        NOMBREDELISTA.Rows = 1
'    End If
'
'    If TIPODELISTA = "GRID" Then
'       NOMBREDELISTA.Redraw = False
'       NOMBREDELISTA.Rows = 2
'       ALTUREN = NOMBREDELISTA.RowHeight(0) + NOMBREDELISTA.RowHeight(1)
'       While ALTUREN < NOMBREDELISTA.Height
'          NOMBREDELISTA.Rows = NOMBREDELISTA.Rows + 1
'          ALTUREN = ALTUREN + NOMBREDELISTA.RowHeight(NOMBREDELISTA.Rows - 1)
'       Wend
'       If TRIM$(ENCABEZADO$) <> "" Then
'          NOMBREDELISTA.FormatString = ENCABEZADO$
'       End If
'
'    End If
'    '
'    LSTX$ = TRIM$(LISTADECAMPOS$)
'    CANCA% = 0
'    SQLX$ = ""
'    While LSTX$ <> ""
'       PPXX% = InStr(LSTX$, ";")
'       PPXX2% = InStr(LSTX$, ",")
'       If PPXX2% > 0 And PPXX2% < PPXX% Then PPXX% = PPXX2%
'       If PPXX% < 1 Then PPXX% = 1 + Len(LSTX$)
'       '
'       CANCA% = CANCA% + 1
'       If CANCA% >= UBound(NOCAMPO$) Then
'          ReDim Preserve NOCAMPO$(CANCA%), FORMX$(CANCA%)
'       End If
'       NOCX$ = Left$(LSTX$, PPXX% - 1)
'       LSTX$ = Mid$(LSTX$, PPXX% + 1)
'       '
'       PPXX% = InStr(NOCX$, "/")
'       If PPXX% < 1 Then PPXX% = 1 + Len(NOCX$)
'       NOCAMPO$(CANCA%) = Left$(NOCX$, PPXX% - 1)
'       FORMX$(CANCA%) = Mid$(NOCX$, PPXX% + 1)
'       If FORMX$(CANCA%) = "" Then FORMX$(CANCA%) = "A16"
'       '
'       If TIPODELISTA = "GRID" Then
'          NOMBREDELISTA.Cols = CANCA%
'       End If
'       '
'       If SQLX$ <> "" Then SQLX$ = SQLX$ + ", "
'       SQLX$ = SQLX$ + NOCAMPO$(CANCA%) + " AS CAMPO" + TRIM$(Str$(CANCA%))
'    Wend
'    '
'    TABLA1$ = TABLA$: TABLA2$ = "": IJO$ = ""
'    PPXX% = InStr(TABLA$, "+")
'    If PPXX% > 0 Then
'       TABLA1$ = TRIM$(Left$(TABLA$, PPXX% - 1))
'       TABLA$ = TRIM$(Mid$(TABLA$, PPXX% + 1))
'       PPXX% = InStr(TABLA$, "/")
'       If PPXX% > 0 Then
'         TABLA2$ = TRIM$(Left$(TABLA$, PPXX% - 1))
'         IJO$ = TRIM$(Mid$(TABLA$, PPXX% + 1))
'         TABLA$ = TABLA1$ + " INNER JOIN " + TABLA2$ + " ON " + IJO$
'       End If
'    End If
'
'    If InStr(MODOSE$, "DISTINCT") > 0 Then SQLX$ = "DISTINCT " + SQLX$
'    SQLX$ = "SELECT " + SQLX$ + " FROM " + TABLA$
'    If TRIM$(CONDICION$) <> "" Then
'       SQLX$ = SQLX$ + " WHERE " & CONDICION$
'    End If
'    On Error Resume Next
'    Set TABTMP = FLEXCONN.Execute(FILTSQL$(SQLX$))
'    If Err Then
'       On Error GoTo 0
'       MsgBox "Imposible Ejecutar Consulta SQL: " + vbCrLf + vbCrLf & SQLX$, vbExclamation
'       Call FINAL("")
'    End If
'    '
'    With TABTMP
'       TTXX$ = ""
'       j% = 0: FILA& = 0
'       Do While Not .EOF
'           TTXX$ = "": COLUI% = 0
'           '
'           If TIPODELISTA = "GRID" Then
'              FILA& = FILA& + 1
'37            If FILA& > NOMBREDELISTA.Rows - 1 Then
'                 NOMBREDELISTA.Rows = NOMBREDELISTA.Rows + 1
'                 GoTo 37
'              End If
'           End If
'           '
'           For U& = 1 To .Fields.Count
'              VALOR = .Fields(U& - 1).Value
'              FORMATX$ = FORMX$(U&) ' FORMATO
'              TECO$ = ""
'              PPXX% = InStr(FORMATX$, "+")
'              If PPXX% > 0 Then
'                 TECO$ = TRIM$(Mid$(FORMATX$, PPXX% + 1))
'                 FORMATX$ = Left$(FORMATX$, PPXX% - 1)
'              End If
'              '
'              ENCAB$ = ""
'              PPXX% = InStr(FORMATX$, "/")
'              If PPXX% > 0 Then
'                 ENCAB$ = Mid$(FORMATX$, PPXX% + 1)
'                 FORMATX$ = Left$(FORMATX$, PPXX% - 1)
'              End If
'              TTYY$ = FORMATNUM$(VALOR, FORMATX$)
'              If TRIM$(TTYY$) = "" Then
'                  TTYY$ = Space$(XVALO(Mid$(FORMATX$, 2)))
'              End If
'              If TIPODELISTA = "LIST" Then
'                   If TTXX$ <> "" Then TTXX$ = TTXX$ + " "
'                   TTXX$ = TTXX$ + TTYY$
'              ElseIf TIPODELISTA = "GRID" Then
'                   COLUI% = COLUI% + 1
'17                 If COLUI% > NOMBREDELISTA.Cols Then
'                      NOMBREDELISTA.Cols = NOMBREDELISTA.Cols + 1
'                      GoTo 17
'                   End If
'                   If FILA& = 1 Then
'                      FORMALTER.TEPRUEBA.Font = NOMBREDELISTA.Font
'                      FORMALTER.TEPRUEBA.FontSize = NOMBREDELISTA.FontSize
'                      FORMALTER.TEPRUEBA.Caption = "l" + TMASCARA$(FORMATX$)
'                      NOMBREDELISTA.ColWidth(COLUI% - 1) = FORMALTER.TEPRUEBA.Width
'                      NOMBREDELISTA.COLALIGNEMENT(COLUI% - 1) = 7
'                      If Left$(FORMATX$, 1) = "A" Or Left$(FORMATX$, 1) = "C" Or Left$(FORMATX$, 1) = "D" Then
'                         NOMBREDELISTA.COLALIGNEMENT(COLUI% - 1) = 1
'                      End If
'                      NOMBREDELISTA.FontBold = True
'                      NOMBREDELISTA.TextMatrix(0, COLUI% - 1) = ENCAB$
'                      NOMBREDELISTA.FontBold = False
'                   End If
'                   If Left$(FORMATX$, 1) = "A" Then NOMBREDELISTA.COLALIGNEMENT(COLUI% - 1) = 1: TTYY$ = VALOR   ' PARA QUE EN LA GRILLA LO PONGA COMPLETO
'
''                   TTYY$ = "HOLA OMAR"
'                   NOMBREDELISTA.TextMatrix(FILA&, COLUI% - 1) = TTYY$
'              End If
'              '
'              'PARA QUE AGRUPE POR LA COLUMNA PASADA COMO PARAMETRO
'              If COLUMNA_AGRUPA% > 0 Then
'                NOMBREDELISTA.MergeCells = flexMergeRestrictColumns
'                NOMBREDELISTA.MergeCol(COLUMNA_AGRUPA%) = True
'              End If
'              '
'              If TECO$ <> "" Then
'                 PPXX% = InStr(TECO$, "/")
'                 If PPXX% < 1 Then PPXX% = 1 + Len(TECO$)
'                 FORMATX$ = Mid$(TECO$, PPXX% + 1)
'                 TECO$ = TRIM$(UCase$(Left$(TECO$, PPXX% - 1)))
'                 '
'                 ENCAB$ = ""
'                 PPXX% = InStr(FORMATX$, "/")
'                 If PPXX% > 0 Then
'                    ENCAB$ = Mid$(FORMATX$, PPXX% + 1)
'                    FORMATX$ = Left$(FORMATX$, PPXX% - 1)
'                 End If
'                 If FORMATX$ = "" Then FORMATX$ = "A16"
'                 '
'                 VALOECO = Null
'                 If Left$(TECO$, 5) = "DEUDO" Or Left$(TECO$, 5) = "CLIEN" Then
'                   VALOECO = RASOCLI$(XVALO(VALOR))
'                 ElseIf Left$(TECO$, 5) = "PROVE" Or Left$(TECO$, 5) = "ACRED" Then
'                   VALOECO = RASOCLI$((-1) * XVALO(VALOR))
'                 ElseIf Left$(TECO$, 5) = "MASTE" Then
'                   TXVALO$ = VALOR
'                     If Len(TXVALO$) > 4 Then
'                          CIXI% = CODINT%(TXVALO$)
'                        Else
'                          CIXI% = XVALO(VALOR)
'                     End If
'                   VALOECO = DESCRIT0$(CIXI%)
'                 ElseIf EXISTE%(LUDAT$ + TECO$ + ".RFS") Then
'                   LFF% = LENFIELD%(TECO$, 1)
'                   ABUS$ = MKS$(XVALO(VALOR))
'                   On Error Resume Next
'                   If LFF% = 2 Then ABUS$ = MKI$(Int(XVALO(VALOR)))
'                   If LFF% = 1 Then ABUS$ = Chr$(Int(XVALO(VALOR)))
'                   On Error GoTo 0
'                   VALOECO = ECOARCH$(TECO$, ABUS$)
'                 End If
'                 '
'                 TTYY$ = FORMATNUM$(VALOECO, FORMATX$)
'                 If TRIM$(TTYY$) = "" Then
'                     TTYY$ = Space$(XVALO(Mid$(FORMATX$, 2)))
'                 End If
'                 If TIPODELISTA = "LIST" Then
'                      If TTXX$ <> "" Then TTXX$ = TTXX$ + " "
'                      TTXX$ = TTXX$ + TTYY$
'                    ElseIf TIPODELISTA = "GRID" Then
'                      COLUI% = COLUI% + 1
'27                    If COLUI% > NOMBREDELISTA.Cols Then
'                         NOMBREDELISTA.Cols = NOMBREDELISTA.Cols + 1
'                         GoTo 27
'                      End If
'                      If FILA& = 1 Then
'                         FORMALTER.TEPRUEBA.Font = NOMBREDELISTA.Font
'                         FORMALTER.TEPRUEBA.FontSize = NOMBREDELISTA.FontSize
'                         FORMALTER.TEPRUEBA.Caption = "l" + TMASCARA$(FORMATX$)
'                         NOMBREDELISTA.ColWidth(COLUI% - 1) = FORMALTER.TEPRUEBA.Width
'                         NOMBREDELISTA.COLALIGNEMENT(COLUI% - 1) = 7
'                         If Left$(FORMATX$, 1) = "A" Or Left$(FORMATX$, 1) = "C" Or Left$(FORMATX$, 1) = "D" Then
'                            NOMBREDELISTA.COLALIGNEMENT(COLUI% - 1) = 1
'                         End If
'                         NOMBREDELISTA.FontBold = True
'                         NOMBREDELISTA.TextMatrix(0, COLUI% - 1 + 1) = ENCAB$
'                         NOMBREDELISTA.FontBold = False
'                      End If
'                      If Left$(FORMATX$, 1) = "A" Then TTYY$ = VALOECO       ' PARA QUE EN LA GRILLA LO PONGA COMPLETO
'                      NOMBREDELISTA.TextMatrix(FILA&, COLUI% - 1) = TTYY$
'                 End If
'                 '
'              End If
'           Next U&
'           '
'           If TIPODELISTA = "LIST" Then
'               NOMBREDELISTA.AddItem TTXX$
'           End If
'           '
'       .MoveNext
'       Loop
'    End With
'    '
'    If TIPODELISTA = GRID Then
'       NOMBREDELISTA.Redraw = True
'    End If
'    '
'    TABTMP.Close
'End Sub

Function MAXNUMCOD&(TABLA$, CAMPO$)
   CONDI$ = CAMPO$ & " > 0"
   MAXNUMCOD& = XVALO(VALOBADA(TABLA$, "MAX(" & CAMPO$ & ")", CONDI$))
   
End Function


Sub ACTUALIZA_DAT_INDIPEDI()
    SQLX$ = "SELECT NROPED,REFERENCIA FROM PEDENCA WITH(NOLOCK) WHERE NROPED > 0 ORDER BY NROPED DESC"
'    Call CARGALISEL("INDIPEDI", "PEDENCA", "NROPED/F6.0 ; REFERENCIA/A60", "NROPED > 0 ORDER BY NROPED DESC")
    Set PEDENCTEMP = READSET(SQLX$)
    '
    jj& = 0
    TTXX$ = ""
    Screen.MousePointer = 11
    With PEDENCTEMP
        Do While Not .EOF
             NRODES& = XVALO(!NROPED)
             REFE$ = Left$(SAFETEXT$(!REFERENCIA), 60)
             Call REPLA(TTXX$, MKS$(NRODES&), 1, 4)
             Call REPLA(TTXX$, REFE$, 5, 60)
             jj& = jj& + 1
             Call GRAREG("INDIPEDI", TTXX$, jj&)
        .MoveNext
        Loop
    End With
    
    PEDENCTEMP.Close
    Set PEDENCTEMP = Nothing
    '
    Call SETULTREG("INDIPEDI", jj&)
    Call CIERRARCH("INDIPEDI")
    Call FRESHECHO("INDIPEDI")
    Screen.MousePointer = 1
End Sub

Sub GrabarArchivo(PathArchivoCompleto$, TABLA$, Optional MODOMESS$)
    On Error GoTo ERROR_GUARDAR
    Screen.MousePointer = 11
    EPAC$ = TRIM$(UCase$(EMPREAC$))
    ES_RIVA_CAMBA_QUIL% = 0
    If InStr(EPAC$, "CAMBACERES") > 0 Or InStr(EPAC$, "CONTEPLAST") > 0 Or InStr(EPAC$, "QUILMES") > 0 Then ES_RIVA_CAMBA_QUIL% = 1

    
    'GRABA STOCK MINIMO Y MAXIMO
    '***************************
    If TRIM$(UCase$(TABLA$)) = "STOCKSUGXSUC" Then
         NARCHI1% = FILEOPEN%(PathArchivoCompleto$, 1, 256) 'ABRE EL ARCHIVO DE LECTURA
         Do While Not EOF(NARCHI1%) ' RECORRE
            Line Input #NARCHI1%, XX$ ' LEE CADA REGISTRO
            CONT% = CONT% + 1
            Call TEXTOLOTES$(XX$, ";")
            CI1% = XVALO(CADENATEX$(2))
            If CI1% > 0 Then
              STMINI = XVALO(CADENATEX$(7))
              STMAXI = XVALO(CADENATEX$(8))
              CONDI$ = "CODINT = " & CI1%
              Call ACTUREGISTRO("MASTER", "STOMIN", CONDI$, STMINI, REGAF&, "NOMESS")
              Call ACTUREGISTRO("MASTER", "STOMAX", CONDI$, STMAXI, REGAF&, "NOMESS")
            End If
         Loop
        Close #NARCHI1%
    'GRABAR EXCEPCION
    '****************
    ElseIf TRIM$(UCase$(TABLA$)) = "EXCEPCION" Then
         NARCHI1% = FILEOPEN%(PathArchivoCompleto$, 1, 256) 'ABRE EL ARCHIVO DE LECTURA
         Do While Not EOF(NARCHI1%) ' RECORRE
            Line Input #NARCHI1%, XX$ ' LEE CADA REGISTRO
            CONT% = CONT% + 1
            Call TEXTOLOTES$(XX$, ";")
            CI1% = XVALO(CADENATEX$(2))
            If CI1% > 0 Then
              NCLIE = XVALO(CADENATEX$(1))
              FE% = CVINT(CADENATEX$(3))
              USUARIOX% = XVALO(CADENATEX$(4))
              NUOC& = XVALO(CADENATEX$(5))
              CANTIEXCEPCION = XVALO(CADENATEX$(6))
              '
              Call GRABAREXCEPCION(CI1%, NCLIE, FE%, USUARIOX%, NUOC&, CANTIEXCEPCION)
            End If
         Loop
         Close #NARCHI1%
    'GRABA OC_SUCURSAL
    '*****************
    ElseIf TRIM$(UCase$(TABLA$)) = "OC_SUCURSAL" Then
         VUELTA% = 0: INDX% = 1
         NARCHI1% = FILEOPEN%(PathArchivoCompleto$, 1, 256) 'ABRE EL ARCHIVO DE LECTURA
         Do While Not EOF(NARCHI1%) ' RECORRE

         
            Line Input #NARCHI1%, XX$ ' LEE CADA REGISTRO
            CONT% = CONT% + 1
            Call TEXTOLOTES$(XX$, ";")
            If TRIM$(CADENATEX$(1)) = "||MENSAJE||" Then GoTo 60
            CLTESUCUR = XVALO(CADENATEX$(1))
            CI1% = XVALO(CADENATEX$(2))
            'ACA AGREGAR EL CODIGO, VALIDAR SI CADENTA TEXT ES = 13
            If ES_RIVA_CAMBA_QUIL% > 0 Then 'SI ES CONTE,CAMBA,QUILMES
              If UBound(CADENATEX$) >= 13 Then 'SI TIENE DATOS EN POSICION 13
                codigoExterno$ = TRIM$(CADENATEX$(13)) 'ESTO ES PARA DETECTAR EL INTERNO DEL MISMO CODIGO EXTERNO,
                CI1% = CODINT%(codigoExterno$)         'P.Q LOS CODIGOS EXTERNOS DEBEN SER IGUALES, NO ASI EL INTERNO
              End If
            End If

            If CI1% > 0 Then
              NROPEDIDOSUCURSAL = XVALO(CADENATEX$(3))
              CLIENTEPEDIDO = XVALO(CADENATEX$(4))
              USUARIOX% = XVALO(CADENATEX$(5))
              NUOC& = XVALO(CADENATEX$(6))
              CANTICOMPRA = XVALO(CADENATEX$(7))
              FESOLENTRE% = FECHANUM(TRIM$(CADENATEX$(8)))
              VERSIONX% = XVALO(CADENATEX$(9))
              FEMISI% = FECHANUM(TRIM$(CADENATEX$(10)))
              OCOMPRACLIENTE$ = ""
              If UBound(CADENATEX$) >= 11 Then
                 OCOMPRACLIENTE$ = TRIM$(CADENATEX$(11))
                 RAZOCLIENTE_SUC$ = TRIM$(CADENATEX$(12))
              End If
              
              If VUELTA% = 0 Then
                CONDI$ = "O_COMPRA_SUC = " & NUOC&
                CONDI$ = CONDI$ + " AND SUCURSAL = " & CLTESUCUR
                Call ACTUREGISTRO("OC_SUCURSAL", "MARBORRA", CONDI$, 1, REGAF&, "NOMESS")
                VUELTA% = 1
              End If

              Call GrabarOcompraSucursal(CLTESUCUR, CI1%, NROPEDIDOSUCURSAL, CLIENTEPEDIDO, USUARIOX%, NUOC&, FEMISI%, CANTICOMPRA, FESOLENTRE%, VERSIONX%, OCOMPRACLIENTE$, RAZOCLIENTE_SUC$, codigoExterno$)
            End If
            If TRIM$(CADENATEX$(1)) = "||MENSAJE||" Then
60             'INDX% = INDX% + 1
               For i% = 2 To UBound(CADENATEX$)
                 ACUMUMENXAJE$ = ACUMUMENXAJE$ + CADENATEX$(i%) + vbCrLf
               Next i%
               Call GrabarMensaje(CLTESUCUR, NUOC&, ACUMUMENXAJE$, USUARIOX%)
                  
            End If

         Loop
         CONDI3$ = CONDI$ + " AND MARBORRA = 1"
         Call BORRAREGISTROS("OC_SUCURSAL", CONDI3$, REGAF&, "NOMESS")

         Close #NARCHI1%
    'GRABA AVISO DE GENERACION DE ORDEN DE VENTA EN AVELLANEDA PARA QUE SUCURSAL REGISTRE
    '*****************
    ElseIf TRIM$(UCase$(TABLA$)) = "OCOMENCA" Then
         VUELTA% = 0
         NARCHI1% = FILEOPEN%(PathArchivoCompleto$, 1, 256) 'ABRE EL ARCHIVO DE LECTURA
         Do While Not EOF(NARCHI1%) ' RECORRE
            Line Input #NARCHI1%, XX$ ' LEE CADA REGISTRO
            CONT% = CONT% + 1
            Call TEXTOLOTES$(XX$, ";")
            NCLIE = XVALO(CADENATEX$(1))
            NPED& = XVALO(CADENATEX$(2))
            OCOMPRA& = XVALO(CADENATEX$(3))
            FECHEMODIF% = FECHANUM(TRIM$(CADENATEX$(4)))
            CI1% = XVALO(CADENATEX$(5))
            Codigo$ = TRIM$(CADENATEX$(6))
            DESCRIX$ = TRIM$(CADENATEX$(7))
            CANTPEDIDA = XVALO(CADENATEX$(8))
            CANTENTREGADA = XVALO(CADENATEX$(9))
            ACCION$ = TRIM$(CADENATEX$(10))
            NUMEUSUARIO% = XVALO(CADENATEX$(11))
            If OCOMPRA& > 0 Then
              Call GrabarPedidoEnOcompraSucursal(OCOMPRA&, NPED&)
              If VUELTA% = 0 Then
                CONDI$ = "NRO_OCOM = " & OCOMPRA&
                CONDI$ = CONDI$ + " AND nroped = " & NPED&
                Call ACTUREGISTRO("DetaOVenta", "MARBORRA", CONDI$, 1, REGAF&, "NOMESS")
                VUELTA% = 1
              End If
              Call ACTUALIZARTABLA_DetaOventa(NCLIE, NPED&, OCOMPRA&, FECHEMODIF%, CI1%, Codigo$, DESCRIX$, CANTPEDIDA, CANTENTREGADA, ACCION$, NUMEUSUARIO%)
            End If
         Loop
         Close #NARCHI1%
        '
         CONDI2$ = CONDI$ + " AND MARBORRA = 1"
         Call BORRAREGISTROS("DetaOVenta", CONDI2$, REGAF&, "NOMESS")
    End If


ERROR_GUARDAR:
    If Err <> 0 Then
       Screen.MousePointer = 1
      Call COMUNI("Error al Guardar, No se Realizó Completamente\Vuelva a Intentarlo Nuevamente\Si esta Situación Persiste - Consulte Con su Soporte de Sistemas" & vbCrLf & Err.Number & " " & Err.Description)
      Exit Sub
    Else
      Call Kill(PathArchivoCompleto$)
      'If MODOMESS$ <> "NOMESS" Then Call COMUNI("Se Han Descargado Los Archivos Correspondientes")
    End If

      Screen.MousePointer = 1
End Sub

Sub GrabarArchivoCTP(PathArchivoCompleto$, TABLA$, Optional MODOMESS$)
    On Error GoTo ERROR_GUARDAR
    Screen.MousePointer = 11
    '****************
    
    'GRABAR REMITO DE CONTEPLAST
    
         VUELTA% = 0: INDX% = 1
         NARCHI1% = FILEOPEN%(PathArchivoCompleto$, 1, 1024) 'ABRE EL ARCHIVO DE LECTURA
         Do While Not EOF(NARCHI1%) ' RECORRE
            Line Input #NARCHI1%, XX$ ' LEE CADA REGISTRO
            Call TEXTOLOTES$(XX$, ";")
            CANTTEXLOT = UBound(CADENATEX$)
            CI1% = XVALO(CADENATEX$(4))
            If CI1% > 0 Then
                CONT% = CONT% + 1
                
                NROREMITO$ = TRIM$(CADENATEX$(1))
                MONEDAX% = XVALO(CADENATEX$(2))
                TIPOCAMBIO# = XVALO(CADENATEX$(3))
                
                CODEXTEX$ = TRIM$(CADENATEX$(5))
                ORDEN% = XVALO(CADENATEX$(6))
                DESCRILAX$ = TRIM$(CADENATEX$(7))
                CANTSALIX = XVALO(TRIM$(CADENATEX$(8)))
                CANTINGREX = XVALO(CADENATEX$(9))
                NUMEORPEX = XVALO(TRIM$(CADENATEX$(10)))
                NROPED& = XVALO(CADENATEX$(11))
                VALOUNIX# = XVALO(CADENATEX$(12))
                NUMECLIX = XVALO(CADENATEX$(13))
                
                FECHMOX = FECHANUM(CADENATEX$(14))
                LOTEX$ = TRIM$(CADENATEX$(15))
                CONDVTA% = XVALO(CADENATEX$(16))
                VENDEX = XVALO(CADENATEX$(17))
                DOMIENT1$ = TRIM$(CADENATEX$(18))
                RSOCIAL$ = TRIM$(CADENATEX$(19))
                NAMEVENDX$ = ""
                If CANTTEXLOT >= 20 Then
                  NAMEVENDX$ = TRIM$(CADENATEX$(20))
                End If
                ORDENCOMPRA$ = ""
                If CANTTEXLOT >= 21 Then
                  ORDENCOMPRA$ = TRIM$(CADENATEX$(21))
                End If
                If CANTTEXLOT >= 22 Then
                  NROVENDENVAP% = XVALO(CADENATEX$(22))
                End If
                ObservaItem$ = ""
                If CANTTEXLOT >= 23 Then
                  ObservaItem$ = TRIM$(CADENATEX$(23))
                End If
                DESCRIDOS$ = ""
                If CANTTEXLOT >= 24 Then
                  DESCRIDOS$ = TRIM$(CADENATEX$(24))
                End If
                
                If VUELTA% = 0 Then
                    CONDI$ = "DOPRILAR = '" & NROREMITO$ & "'"
                    Call ACTUREGISTRO("MOVISTOCTP", "MARBORRA", CONDI$, 1, REGAF&, "NOMESS")
                    VUELTA% = 1
                 End If
    
                 Call GrabaRemitoSucursalCTP(NROREMITO$, MONEDAX%, TIPOCAMBIO#, CI1%, CODEXTEX$, ORDEN%, DESCRILAX$, CANTINGREX, CANTSALIX _
                 , NUMEORPEX, NROPED&, VALOUNIX#, NUMECLIX, FECHMOX, LOTEX$, CONDPAGX, VENDEX, DOMIENT1$, RSOCIAL$, NAMEVENDX$, ORDENCOMPRA$, NROVENDENVAP%, ObservaItem$, DESCRIDOS$)
            End If
         Loop
         CONDI3$ = CONDI$ + " AND MARBORRA = 1"
         Call BORRAREGISTROS("MOVISTOCTP", CONDI3$, REGAF&, "NOMESS")
         Close #NARCHI1%


ERROR_GUARDAR:
    If Err <> 0 Then
       Screen.MousePointer = 1
       MsgBox Err.Description
      Call COMUNI("Error al Guardar, No se Realizó Completamente\Vuelva a Intentarlo Nuevamente\Si esta Situación Persiste - Consulte Con su Soporte de Sistemas" & vbCrLf & Err.Number & " " & Err.Description)
      Exit Sub
    Else
      Call Kill(PathArchivoCompleto$)
    End If

      Screen.MousePointer = 1
End Sub



 Sub ACTUALIZARTABLA_DetaOventa(NCLIE, NPED&, OCOMPRA&, FECHEMODIF%, CI1%, Codigo$, DESCRIX$, CANTPEDIDA, CANTENTREGADA, ACCION$, NUMEUSUARIO%)
   
   'siempre genera un archivo un regisro nuevo antes de la llamada hace un marborra = 1 y luego los borra.
   '*****************************************************************************************************
   SQLX$ = "Select * from DetaOVenta"
   SQLX$ = SQLX$ + " WHERE NRO_OCOM = " & NUOC&
   SQLX$ = SQLX$ + "  AND nroped = " & NPED&
   Dim RS As ADODB.Recordset
   Set RS = New ADODB.Recordset
   
   RS.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
   '
   With RS
    .AddNew
    !NROPED = NPED&
    !nro_ocom = OCOMPRA&
    !CODINT = CI1%
    !CODIEXT = Codigo$
    !Descrip = DESCRIX$
    !CANPED = CANTPEDIDA
    !CANTENT = CANTENTREGADA
    !FECHAMODIF = CVDAT(FECHEMODIF%)
    !ACCION = ACCION$
    !NCtePedido = NCLIE
    !MARBORRA = 0
    .Update
   End With

 End Sub

Sub GRABAREXCEPCION(CI1%, NCLIE, FE%, USUARIOX%, NUOC&, CANTIEXCEPCION)
   
   SQLX$ = "Select * from Excepcion"
   SQLX$ = SQLX$ + " WHERE NUME_OCOM = " & NUOC&
   SQLX$ = SQLX$ + " AND CODINT = " & CI1%
   SQLX$ = SQLX$ + " AND NUCLIEN = " & NCLIE
    
   Dim RS As ADODB.Recordset
   Set RS = New ADODB.Recordset

   RS.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
   '
   With RS
       If .EOF Then
         .AddNew
       End If
       On Error Resume Next
       !CODIEMPR = CodiEmp%
       On Error GoTo 0
       
       !nume_ocom = NUOC&
       !NUCLIEN = NCLIE
       !CODINT = CI1%
       !Codigo = CODEXT$(CI1%)
       !FECHA = CVDAT(FE%)
       !USUARIO = USUARIOX%
       !CANTIEXCEPCION = CANTIEXCEPCION
      .Update
   End With
End Sub

Sub GrabarPedidoEnOcompraSucursal(OCOMPRA&, NPED&)
   
   SQLX$ = "Select Nume_Pedi from OCOMENCA "
   SQLX$ = SQLX$ + " WHERE NUME_OCOM = " & OCOMPRA&
   Dim RS As ADODB.Recordset
   Set RS = New ADODB.Recordset

   RS.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
   '
   With RS
     If Not .EOF Then
      !Nume_Pedi = NPED&
      .Update
     End If
   End With
   RS.Close
   Set RS = Nothing

End Sub

Sub GrabarOcompraSucursal(SUCURSAL, CI1%, NROPEDIDOSUCURSAL, CLIENTEPEDIDO, USUARIOXSUCURSAL%, NUOC&, FEMISI%, CANTICOMPRA, FESOLENTRE%, VERSIONX%, OCOMPRACTESUCURSAL$, RAZOCLIENTE_SUC$, Optional codigoExterno$)
   SQLX$ = "Select * from OC_SUCURSAL"
   SQLX$ = SQLX$ + " WHERE O_COMPRA_SUC = " & NUOC&
   SQLX$ = SQLX$ + " AND SUCURSAL = " & SUCURSAL
    
   Dim RS As ADODB.Recordset
   Set RS = New ADODB.Recordset

   RS.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
   '
   With RS
'       If .EOF Then
         .AddNew
'       End If
       On Error Resume Next
       !CODIEMPR = CodiEmp%
       On Error GoTo 0

       !SUCURSAL = SUCURSAL
       !CODINT = CI1%
       !NROPED_SUC = NROPEDIDOSUCURSAL
       !NCLIEN_PEDIDO = CLIENTEPEDIDO
       !USUARIO = USUARIOXSUCURSAL%
       !O_COMPRA_SUC = NUOC&
       !Cantidad = CANTICOMPRA
       !NPED_GENERADO_EN_AVELLANEDA = 0
       !VERSION_COMPRA_SUC = VERSIONX%
       !FECHA_SOLICI_COMPRA = FETEXCOR1$(FESOLENTRE%)
       !FECHEMISI = CVDAT(FEMISI%)
       !MARBORRA = 0
       !Status = 0
       !OComCliente = OCOMPRACTESUCURSAL$
       !RAZ_CTE_SUC = Left$(RAZOCLIENTE_SUC$, 32) 'LA RAZON SOCIAL DEL CLIENTE DE LA SUCURSAL
       !Cod_Exte = Left$(codigoExterno$, 24)
      .Update
   End With
   RS.Close
   Set RS = Nothing

End Sub
Sub GrabaRemitoSucursalCTP(NROREMITO$, MONEDAX%, TIPOCAMBIO#, CI1%, CODEXTEX$, ORDEN%, DESCRILAX$, CANTINGREX, CANTSALIX _
             , NUMEORPEX, NROPED&, VALOUNIX#, NUMECLIX, FECHMOX, LOTEX$, CONDPAGX, VENDEX, DOMIENT1$, RSOCIAL$, NAMEVENDEDOR$, ORDENCOMPRA$, NROVENDENVAP%, ObservaItemx$, DESCRIDOSX$)
'   Remito realizado en conteplast se graba en la tabla movistoctp para luego pasar a waldbot
   
   SQLX$ = "Select * from MOVISTOCTP"
   SQLX$ = SQLX$ + " WHERE DOPRILAR  = '" & NROREMITO$ & "'"
    
   Dim RS As ADODB.Recordset
   Set RS = New ADODB.Recordset

   RS.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
   '
   With RS
        .AddNew
        On Error Resume Next
        !CODIEMPR = CodiEmp%
        On Error GoTo 0

        !FeReMovi = Now
        !FechMov = FECHATEX$(FECHMOX)
        !DESCRI_LAR = DESCRILAX$
        !idenlote = LOTEX$
        !CODIMOVI = "RT"
        !DOPRILAR = NROREMITO$
        !NuOrItem = ORDEN%
        !NUPEDCLI = NROPED&
        !NUMEORDPED = NUMEORPEX
        !VALOUNIT = ROUND(VALOUNIX#, 3)
        !ValTotIt = VALOUNIX# * CANTSALIX
        !MoneVal = MONEDAX%
        !TIPOCAM = TIPOCAMBIO#
        !NUME_CLIE = NUMECLIX
        !raso_cli = rasocli$(NUMECLIX)
        !CANTSALID = CANTSALIX
        !CANTINGRE = CANTINGREX
        !CondPag = CONDPAGX
        !VENDED = VENDEX
        !cod_inte = CI1%
        !Cod_Exte = CODEXTEX$
        !Descrip = Left$(DESCRILAX$, 64)
        !raso_cli = RSOCIAL$
        !DOMIENT = DOMIENT1$
        !FeRemAWb = Now 'fecha remitido a waldbot
        !PasoaWb = 0 'marca de pasado a Waldbot
'        MOVISTOCTP = 0
        !MARBORRA = 0
        !NAMEVEND = NAMEVENDEDOR$
        !ORDENCOMPRA = Left(ORDENCOMPRA$, 32)
        !NroVend_Envap = NROVENDENVAP% 'viene de pedenca
        !ObservaItem = ObservaItemx$
        !DESCRIDOS = DESCRIDOSX$
        .Update
   End With
   RS.Close
   Set RS = Nothing

End Sub

Function GENERA_CONSULTA_SQL_CUENTAS_RETIB$()

    VUELTA% = 0
    CONDICIONX$ = ""
    If EXISTE%(LUDAT$ & "PROVINB") < 1 Then Exit Function
    
    Screen.MousePointer = 11
    
    For i& = 1 To ULTREG&("PROVINB")
        X$ = REGLEIDO$("PROVINB", i&)
        CuentaRetencionIIBB% = CVI(Mid$(X$, 29, 2))
        CodigoJurisdiccion% = CVI(Mid$(X$, 37, 2))
        If CuentaRetencionIIBB% > 0 And CodigoJurisdiccion% > 0 Then
          If VUELTA% = 0 Then
            VUELTA% = 1
            CONDICIONX$ = " (CUECONTAI = " & CuentaRetencionIIBB%
          Else
            CONDICIONX$ = CONDICIONX$ & " OR CUECONTAI = " & CuentaRetencionIIBB%
          End If
        ElseIf CuentaRetencionIIBB% > 0 And CodigoJurisdiccion% = 0 Then
            GoTo ErrorJurisdiccion
        End If
    Next i&
    
    If TRIM$(CONDICIONX$) <> "" Then CONDICIONX$ = CONDICIONX$ & ")" 'CIERRO LA CONSULTA
    GENERA_CONSULTA_SQL_CUENTAS_RETIB$ = CONDICIONX$
    
    Call CIERRARCH("PROVINB")
    
    Screen.MousePointer = 1
    
    Exit Function
    
ErrorJurisdiccion:
    Screen.MousePointer = 1
    GENERA_CONSULTA_SQL_CUENTAS_RETIB$ = ""
    Exit Function

End Function
Function EXISTESQR_ENSQL%(NOMBSQR$)
    EXISTESQR_ENSQL% = 0
    CONDI$ = " COD_QUERY = '" & NOMBSQR$ & "'"
    If CantRegistros&("SQL_QUERIES", CONDI$, "NOMESS") > 0 Then EXISTESQR_ENSQL% = 1
End Function
Sub CARGAPEDIDOS(Optional STAT%, Optional ModoOrdenamiento%)
   
   Static CARGADO%
   If CARGADO% > 0 And IsFormLoaded(FRMZELECHO) Then GoTo 50
     
   CARGADO% = 1
   CONDI$ = "NROPED > 0 "
   If InStr(1, EMPREAC$, "TANIT") Then CONDI$ = "NROPED > 199000 "
   If InStr(EMPREAC$, "MEDITEA") > 0 Then CONDI$ = " NROPED > 50000 "
   'FIN& = CantRegistros("PEDENCA", CONDI$, "NOMESS")
   SQLX$ = " SELECT NROPED, REFERENCIA FROM PEDENCA WITH(NOLOCK)"
   SQLX$ = SQLX$ + " WHERE " & CONDI$
   If STAT% > 0 Then CONDI$ = " AND Status = " & STAT%
   If ModoOrdenamiento% = 0 Then
      SQLX$ = SQLX$ + " ORDER BY Nroped "
   Else
      SQLX$ = SQLX$ + " ORDER BY Nroped desc"
   End If
   Campos$ = "NRO-PEDIDO/F14.0;REFERENCIA/A60"
   Call CARGA_ENCABEZADO(FRMZELECHO.MSF1, Campos$, FRMZELECHO, , , ANTOT)
   Call CARGA_ENCABEZADO(FRMZELECHO.MSF2, Campos$, FRMZELECHO, , , ANTOT)
   FRMZELECHO.Width = 9000
   FRMZELECHO.Caption = "Indice General de Pedidos (Pendientes-Cumplidos-Anulados) (" & FIN& & " Registros)"
   FRMZELECHO.MSF2.Rows = FIN& + 2
   Set RS = READSET(SQLX$)
   
   FIN& = RS.RecordCount
   FRMZELECHO.MSF2.Rows = FIN& + 2
   FRMZELECHO.MSF2.Redraw = False
   With RS
      Do While Not .EOF
        i& = i& + 1
        FRMZELECHO.MSF2.TextMatrix(i&, 1) = XVALO(!NROPED)
        FRMZELECHO.MSF2.TextMatrix(i&, 2) = SAFETEXT$(!REFERENCIA)
        .MoveNext
      Loop
    End With
    RS.Close
    Set RS = Nothing
    FRMZELECHO.MSF2.Redraw = True
    '
50  FRMZELECHO.Show 1
    
    NROPED = XVALO(RESP_ZELE_VECT(1))
    '
End Sub

Function ObtenerIdTablaAuxiliar(NombreTabla As String) As Long
    
    ObtenerIdTablaAuxiliar = XVALO(VALOBADA("ENCAPRFUN", "ID_LISTA", "LISTA = '" & NombreTabla & "'", "NOMESS"))
    
End Function

Private Function ListaSelDataSql(nombre As String) As Boolean

    ' OBTIENE ID DE LA TABLA, DESCRIPCION Y CONSULTA PARA DATASQL
    Dim RstENCAPRFUN As ADODB.Recordset
    ConsultaSql$ = "SELECT ID_LISTA, LISTA, DESCRIPCION, CONSULTASQL FROM ENCAPRFUN WITH(NOLOCK) WHERE LISTA='" & nombre & "'"
    Set RstENCAPRFUN = READSET(ConsultaSql$)
    If RstENCAPRFUN.EOF And RstENCAPRFUN.BOF Then
        ListaSeleccionDataSql = False
        Exit Function
    End If
    ID_Lista& = RstENCAPRFUN!ID_Lista
    Descripcion$ = RstENCAPRFUN!Descripcion
    ConsultaSqlListaSeleccion$ = RstENCAPRFUN!ConsultaSql
    RstENCAPRFUN.Close
    Set RstENCAPRFUN = Nothing
    
    ' OBTIENE LISTA DE CAMPOS PARA ARMAR GRILLA
    Dim RstENCAPRF As ADODB.Recordset
    ConsultaSql$ = "SELECT ENCABEZADO FROM ENCAPRF WITH(NOLOCK) WHERE ID_LISTA=" & CStr(ID_Lista&)
    Set RstENCAPRF = READSET(ConsultaSql$)
    If RstENCAPRF.EOF And RstENCAPRF.BOF Then
        ListaSeleccionDataSql = False
        Exit Function
    End If
    CantidadDeCampos% = 0
    With RstENCAPRF
        Do While Not .EOF
            Campos$ = Campos$ & !Encabezado & ";"
            CantidadDeCampos% = CantidadDeCampos% + 1
            .MoveNext
        Loop
    End With
    Campos$ = Mid$(Campos$, 1, Len(Campos$) - 1)
    RstENCAPRF.Close
    Set RstENCAPRF = Nothing

    ' ARMA ZELECHO CON LOS DATOS DE LA TABLA
    Call CARGA_ENCABEZADO(FRMZELECHO.MSF2, Campos$, FRMZELECHO)
    Call CARGA_ENCABEZADO(FRMZELECHO.MSF1, Campos$, FRMZELECHO)
    FRMZELECHO.Caption = Descripcion$
    
    ' CARGA GRILLA CON LOS DATOS DE LA TABLA
    Dim Rst As ADODB.Recordset
    ConsultaSql$ = ConsultaSqlListaSeleccion$ & " WHERE ID_LISTA=" & CStr(ID_Lista&)
    Set Rst = READSET(ConsultaSql$)
    With Rst
        Do While Not .EOF
            FRMZELECHO.MSF2.Rows = FRMZELECHO.MSF2.Rows + 1
            For i% = 0 To CantidadDeCampos% - 1
                FRMZELECHO.MSF2.TextMatrix(FRMZELECHO.MSF2.Rows - 1, i% + 1) = .Fields(i%)
            Next i%
            .MoveNext
        Loop
    End With
    On Error Resume Next
    Rst.Close
    Set Rst = Nothing
    On Error GoTo 0
    
    ' DETERMINA ANCHO A PARTIR DEL ANCHO DE LA GRILLA
    FRMZELECHO.Width = FRMZELECHO.MSF2.Width + 200
    ' DESPLIEGA LISTA
    FRMZELECHO.Show 1
    
    ListaSeleccionDataSql = True

End Function

Sub ABM_TABLA(IDLIST&, Optional NOMBRELISTA$, Optional VALIDARDUPFILA1$, Optional ORDENPORCAMPOS$)
    If IDLIST& < 1 Then
       IDLIST& = XVALO(VALOBADA("ENCAPRFUN", "ID_LISTA", "LISTA = '" & NOMBRELISTA$ & "'"))
       If IDLIST& < 1 Then Exit Sub
    End If
    Unload VENTABNEW
    'CARGO LOS DATOS UNICOS DE ENCABEZADOS
     
    Dim obj As New RECORXET
    Set rs1 = obj.RS_TAB_COMP_ENCAPRFUN(IDLIST&)
    MAXIMOCAMPOS = XVALO(rs1!canticampos)
    VENTANA.Inicializar = 0
    VENTANA.AgregaRegistro = XVALO(rs1!ADDREGISTROS)
    VENTANA.UtilizaFormula = XVALO(rs1!USAFORMULA)
    VENTANA.HabilitaBorrar = XVALO(rs1!BORRAR)
    VENTANA.HabilitaInsertar = XVALO(rs1!Insertar)
    VENTANA.NoMostrarBotonTilde = XVALO(rs1!NOTILDE)
    VENTANA.TITULO = SAFETEXT$(rs1!TITURECORD)
    CONSU$ = SAFETEXT$(rs1!ConsultaSql)
    LISTAX$ = SAFETEXT$(rs1!Lista)
    Descrip$ = SAFETEXT$(rs1!Descripcion)
    
    SQLX$ = "SELECT * FROM ENCAPRF WITH(NOLOCK) WHERE ID_LISTA = " & IDLIST&
    SQLX$ = SQLX$ + " ORDER BY ENCAPRF_ID ASC"
    Set RS = READSET(SQLX$)
    With RS
    J% = 0
      Do While Not .EOF
        J% = J% + 1
        If J% = 1 Then
          ID1$ = TRIM$(SAFETEXT$(!Encabezado))
          Call TEXTOLOTES(ID1$, "/")
          On Error Resume Next
            FORMATID = TRIM$(UCase$(CADENATEX$(2)))
            LONGFORM = Mid$(FORMATID, 2)
          On Error GoTo 0
        End If
        ENCAX1$ = TRIM$(SAFETEXT$(!Encabezado))
        
        FORMUX1 = TRIM$(SAFETEXT$(!Formula))
        CAMPEDI1$ = TRIM$(SAFETEXT$(!CAMPEDITAB))
        COLECO1$ = TRIM$(SAFETEXT$(!COLCAMECO))
        TXTTOTX1$ = TRIM$(SAFETEXT$(!TXTTOTAL))
        TOTAX1$ = TRIM$(SAFETEXT$(!Total))
        DEJLIB1$ = TRIM$(SAFETEXT$(!DejarLibre))
        If ENCAX1$ <> "" Then ENCAX$ = ENCAX$ + ENCAX1$ + ";" 'ENCABEZADO
        If FORMUX1 <> "" Then VENTANA.Formula(J%) = FORMUX1
        If TRIM(UCase(CAMPEDI1$)) = "S" Then CAMPEDI$ = CAMPEDI$ & J% & ";" 'CAMPOS EDITABLES
        If COLECO1$ <> "" Then VENTANA.ColEcoCampo(J%) = COLECO1$
        If TXTTOTX1$ <> "" Then VENTANA.TxtUbicadoGrillaTotal(J%) = TXTTOTX1$
        If TOTAX1$ <> "" Then VENTANA.TOTALES(J%) = TOTAX1$
        If DEJLIB1$ <> "" Then VENTANA.DejarLibre(J%) = DEJLIB1$
        .MoveNext
      Loop
    End With
    RS.Close
    Set RS = Nothing
    If ENCAX$ <> "" Then
      If Right(ENCAX$, 1) = ";" Then ENCAX$ = Left$(ENCAX$, Len(ENCAX$) - 1)
      VENTANA.Campox = ENCAX$
    End If
    '
    If CAMPEDI$ <> "" Then
      If Right(CAMPEDI$, 1) = ";" Then CAMPEDI$ = Left$(CAMPEDI$, Len(CAMPEDI$) - 1)
      VENTANA.CampEditables = CAMPEDI$
    End If
    '
    VENTANA.Inicializar = 1
    '
    rs1.Close
    Set rs1 = Nothing
    
    SQLX$ = CONSU$ + " WHERE ID_LISTA = " & IDLIST&
    ORDENA$ = "": If ORDENPORCAMPOS$ <> "" Then ORDENA$ = " ORDER BY " & ORDENPORCAMPOS$
    Set rs1 = READSET(SQLX$ + ORDENA$, , 1)
    Call Carga_MSF_Recordset(rs1, VENTANA.Campox, VENTABNEW.MSF)
    rs1.Close
    Set rs1 = Nothing
    VENTABNEW.MSF.COL = 1
    If ORDENPORCAMPOS$ = "" Then VENTABNEW.MSF.Sort = 3
'    If VENTABNEW.MSF.Rows < 2 Then VENTABNEW.MSF.Rows = 2
    VENTABNEW.Caption = Descrip$
50  VENTABNEW.Show 1
      If VENTABNEW.APROBADO% > 0 Then
        If VENTANA.ValidarDuplicadoFila1 <> "N" Then
          DUPLIC = VALIDAR_DUPLICADO(VENTABNEW.MSF, 1, 1)
          If DUPLIC > 0 Then
            Call COMUNI("Código Duplicado En Fila: " & DUPLIC)
            GoTo 50
          End If
      End If
      
        Call GRABASQLDAT(SQLX$, RET_CAMPOS$, RET_TABLA$, RET_CONDI$)
        If VALIDARCAMPOFECHA%(RET_CAMPOS$) < 0 Then GoTo 50
        'MARCO CON MARBORRA
        Call ACTUREGISTRO(RET_TABLA$, "MARBORRA", RET_CONDI$, 1, REGAF&)
        Call TEXTOLOTES$(RET_CAMPOS$, ",")
        On Error Resume Next
        CANTICAMP = UBound(CADENATEX$)
        ORD% = 0
        On Error GoTo 0
        If CANTICAMP > 0 Then
          For i& = 1 To VENTABNEW.MSF.Rows - 1
             If TRIM$(VENTABNEW.MSF.TextMatrix(i&, 1)) = "" And Left(CAMPEDI$, 1) < 2 Then
                GoTo 60
             ElseIf Left(CAMPEDI$, 1) > 1 Then 'QUE EL MISMO NO ESTA EDITABLE
                'SE AGREGO ESTE ELSE PARA QUE SI NO ESTA EDITABLE LA PRIMER COLUMNA
                'ENUMERE AUTOMATICAMENTE (RECOMENDAMOS QUE SI SE USA A MODO NO EDITABLE NO SE PUEDAN BORRAR REGISTROS
                'NI AGREGAR POR QUE PODRIA CAMBIAR LOS CODIGOS AUTOGENERADOS )'22/03/18'
                ORD% = ORD% + 1
                VENTABNEW.MSF.TextMatrix(i&, 1) = ORD%
             End If
             
             VALOX$ = IDLIST& & ","
             VALOX$ = VALOX$ & Chr(39) & LISTAX$ & Chr(39) & ","
             For k& = 1 To VENTABNEW.MSF.Cols - 1
                DATOX$ = Chr(39) & TRIM$(VENTABNEW.MSF.TextMatrix(i&, k&)) & Chr(39)
                If k& = 1 Then
                  LETFORM = Left(FORMATID, 1)
                  If LETFORM = "F" Then
                    CEROX$ = String$(LONGFORM, "0")
                    DATOX$ = VENTABNEW.MSF.TextMatrix(i&, k&)
                    DATOX$ = Right$(CEROX$ & TRIM$(DATOX$), XVALO(LONGFORM))
                    DATOX$ = Chr(39) & DATOX$ & Chr(39)
                   End If
                Else
'                  DATOX$ = Chr(39) & VENTABNEW.MSF.TextMatrix(i&, k&) & Chr(39)
                End If
                VALOX$ = VALOX$ + DATOX$ + ","
             Next k&
             If Right$(VALOX$, 1) = "," Then VALOX$ = Left$(VALOX$, Len(VALOX$) - 1)
             SQLX$ = "INSERT INTO " + RET_TABLA$
               PRICAMP$ = "ID_LISTA, LISTA, "
             SQLX$ = SQLX$ + " (" + PRICAMP$ + RET_CAMPOS$ + ") "
             SQLX$ = SQLX$ + " VALUES (" & TRIM$(VALOX$) & ")"
             FLEXCONN.Execute SQLX$
60        Next i&
        End If
        Call BORRAREGISTROS(RET_TABLA$, RET_CONDI$ & " AND MARBORRA = 1", REGAF&, "NOMESS")
    End If
    Set VENTANA = Nothing
    Unload VENTABNEW
    Set VENTANA = Nothing
End Sub

Sub GRABASQLDAT(SQLX$, RET_CAMPO$, RET_TABLA$, RET_CONDI$)
   'RET_TABLA$ (DATASQL)
   'RET_CAMPO$ (CAMPO1, CAMPO2)
   'RET_CONDI$ (ID_LISTA = 12)
   CONSU$ = UCase$(SQLX$)
   POS1% = InStr(CONSU$, " ")
   POS2% = InStr(CONSU$, " FROM ")
   POS3% = POS2% + 6 ' DONDE COMIENZA FROM + LONGITUD DE FROM Y 2 ESPACIOS
   POS4% = InStr(CONSU$, "WITH(NOLOCK)")
   POS5% = InStr(CONSU$, "WHERE ")
   POS6% = InStr(CONSU$, "ORDER ")
   POS5% = POS5% + 6
   RET_CAMPO$ = Mid$(CONSU$, POS1%, POS2% - POS1%)
   RET_TABLA$ = Mid$(CONSU$, POS3%, POS4% - POS3%)
   If POS6% > 0 Then
     RET_CONDI$ = Mid$(CONSU$, POS5%, POS6%)
   Else
     RET_CONDI$ = Mid$(CONSU$, POS5%)
   End If
End Sub
Sub CARGA_TABLA_IMPRE(MSF As Object, Campos$, NAMECONS$, DESCRICONS$, TABLA1$, Optional RET_NAMECONS$, Optional BorrTabImpre%, Optional SOLO_SELECCIONADOS%, Optional APARTIR_COLUM%)
    '
    'MODIFICACION REALIZADA
    '1.SE GENERA UN ID UNICO PARA LA NUEVA CONSULTA CON LA FUNCION IDCONSULTABLAIMPRE
    'ESTA CREA UN REGISTRO CON EL NUEVO ID QUE AL FINAL SE ELIMINA, PERO DURANTE TODO EL PROCESO SE UTILIZA
    'SE AGREGA ESTE ID COMO CONDICION A MOSTRAR EN LA TABLA_IMPRE
    'SE VALIDA SI EXISTE EL SQF EN ESE CASO SE GENERA CON OTRO NOMBRE AGREGANDO 01 ...02....HASTA NO ENCONTRAR NINGUNA C/MISMO NOMBRE
    'SE REALIZA UN BLOQUEO DESDE EL COMIENZO HASTA QUE ARMA LA CONSULTA (SQF)
    'BORRA LOS REGISTROS DE LA TABLA_IMPRE MENORES AL DIA DE HOY
    'PARA PRESENTAR EL GENERADOR DE REPORTES SE DEBERA LLAMAR DE LA SIGUIENTE FORMA:
'    Call FINAL("?" & RET_NAMECONS$) QUE ES EL PARAMETRO DE RETORNO POR SI SE MODIFICO EL NOMBRE POR QUE YA EXISTIA LA CONSULTA
'    Call EliminarSQF(RET_NAMECONS$) 'LUEGO DE PRESENTAR LA CONSULTA SE DEBERA BORRAR INMEDIATAMENTE.
    
    
    'APARTIR_COLUM% SI NO TIENE VALOR ES A PARTIR DE  COLUMNA 1, SI SE LE PASA APARTIR_COLUM% = -1 ES A PARTIR DE 0
    If APARTIR_COLUM% = 0 Then APARTIR_COLUM% = 1
    If APARTIR_COLUM% < 0 Then APARTIR_COLUM% = 0
    If BorrTabImpre% > 0 Then
      CONDI$ = "TABLA_IMPRE_ID > 0 "
      Call BORRAREGISTROS("TABLA_IMPRE", CONDI$, REGAF&, "NOMESS")
    End If
    
    RET_NAMECONS$ = NAMECONS$ ' VALOR DE RETORNO POR DEFECTO
    Call BLOQARCH("LICOREP")
    IDCONSUL = IDCONSULTABLAIMPRE ' funcion que crea un registro y devuelve el nuevo id
'    A = TRIM$(Left$(List2.List(i&), 24)) + " " + TRIM$(Mid$(List2.List(i&), 26, 8)) + " " + TRIM$(Mid$(List2.List(i&), 35))
    TX$ = ""
    Call REPLA(TX$, "IDCONSUL", 1, 24)
    Call REPLA(TX$, "=", 26, 8)
    Call REPLA(TX$, TRIM$(SAFETEXT$(IDCONSUL)), 36, 14)
    CONDITIONE$ = TX$
    'GENERACION DE SQF
    'SI LA CONSULTA EXISTE GENERA UNA NUEVA
    i& = 0
    If BorrTabImpre% > 0 Then GoTo 8

5   If EXISCONSUSQL%(RET_NAMECONS$) > 0 Then
       RET_NAMECONS$ = Mid$(RET_NAMECONS$, 1, Len(RET_NAMECONS$) - 2)
       i& = i& + 1
       RET_NAMECONS$ = RET_NAMECONS$ & FORMATOCON0$(i&, 2)
       GoTo 5
    Else
      GoTo 6
    End If
    '
6   i& = 0
7   If EXISCONSUMDB%(RET_NAMECONS$) > 0 Then
       RET_NAMECONS$ = Mid$(RET_NAMECONS$, 1, Len(RET_NAMECONS$) - 2)
       i& = i& + 1
       RET_NAMECONS$ = RET_NAMECONS$ & FORMATOCON0$(i&, 2)
       GoTo 7
    End If
    
8  CONSUB$ = ""
    CONSUB$ = "CODQUERY = " + TRIM$(RET_NAMECONS$) + Chr$(124)
    CONSUB$ = CONSUB$ + "DESQUERY = " + TRIM$(RET_NAMECONS$) + Chr$(124)
    CONSUB$ = CONSUB$ + "TIPQUERY = " + "   0" + Chr$(124)
    CONSUB$ = CONSUB$ + "SELTABLE = " + TABLA1$ + Chr$(124)
    '
    'CARGO LOS CAMPOS Y EL FORMATO
    Call TEXTOLOTES(Campos$, ";")
    On Error Resume Next
    For i& = 1 To UBound(CADENATEX$)
       Dim strArray() As String
       On Error Resume Next
       strArray = Split(CADENATEX$(i&), "/")
       TIT_CAMPO$ = strArray(0)
       FORMATX$ = strArray(1)
       OCULTA$ = "0"
       If XVALO(Mid$(FORMATX$, 2)) < 1 Then OCULTA$ = "1"
       CAMPOSE$ = Space$(168)
       Mid$(CAMPOSE$, 1, 64) = "CAMPO" & i&
       Mid$(CAMPOSE$, 65, 8) = AJUSTI$(FORMATX$, 8)
       Mid$(CAMPOSE$, 73, 1) = OCULTA$
       Mid$(CAMPOSE$, 81, 64) = AJUSTI$(TIT_CAMPO$, 64)
       CAMPO$ = CAMPOSE$
       CONSUB$ = CONSUB$ + "SELFIELD = " + CAMPO$ + Chr$(124)
    Next i&
    '
    'PARA QUE SE MUESTRE TAL CUAL SE CARGA SE ORDENA POR EL ID
    Mid$(CAMPOSE$, 1, 64) = "TABLA_IMPRE_ID"
    Mid$(CAMPOSE$, 65, 8) = AJUSTI$("F12", 8)
    Mid$(CAMPOSE$, 73, 1) = "1"
    MODOSOR$ = "ASC"
    Mid$(CAMPOSE$, 77, 4) = MODOSOR$
    Mid$(CAMPOSE$, 81, 64) = AJUSTI$("", 64)
    CAMPO$ = CAMPOSE$
    CONSUB$ = CONSUB$ + "SELFIELD = " + CAMPO$ + Chr$(124)
    CONSUB$ = CONSUB$ + "CONDITIO = " + CONDITIONE$ + Chr$(124) ' condicion con identificador
    '
    On Error GoTo 0
    '
       
10     CODMO$ = RET_NAMECONS$
       DESMO$ = DESCRICONS$
       TEXMO$ = CONSUB$
       '
       If EXISTE_TABLA%("SQL_QUERIES") > 0 Then
          Call GRABA_CONSULTA_SQL(CODMO$, DESMO$, 0, TEXMO$, OK%)
       End If
       
       Call ABRETACON
       strSQL = "SELECT * FROM SQL_QUERIES WHERE COD_QUERY = '" & CODMO$ & "'"
       Set CONSUSQL = BaDaConsul.OpenRecordset(strSQL, dbOpenDynaset)
       With CONSUSQL
          If (.BOF And .EOF) Then
            .AddNew   ' YA CONVERTIDO
            On Error Resume Next
            !CODIEMPR = CodiEmp%
            On Error GoTo 0
            GoTo 15
          End If
          'If COMALTER%("Ya Existe una Consulta Registrada\Bajo Ese Código.\\Cancela Guardar\Sobre-Escribir") <> 2 Then GoTo 20
          .Edit
15        !COD_QUERY = CODMO$
          !DES_QUERY = DESMO$
          !TIP_QUERY = 0        ' REPORTE TABULADO
          !PAR_QUERY = TEXMO$
          .Update
       End With
       '
20     On Error Resume Next
       CONSUSQL.Close
       On Error GoTo 0
       Set CONSUSQL = Nothing
       BaDaConsul.Close
       '
       Call LIBARCH("LICOREP")
30     If EXISTE%("F:\FLEXOFT\NOVEDAD\SQL_QUERY\.") > 0 Then
           ASALI$ = "F:\FLEXOFT\NOVEDAD\SQL_QUERY\" + CODMO$ + ".SQF"
           Call SAVEFILE(ASALI$, TEXMO$)
       End If
       '
       'HASTA AQUI LA GENERACION DEL ARCHIVO SQF.
       '
       'BORRA LA TABLA
       VUELTA% = 0
       A = FETEXCOR1(HOY(HO$))
       
40     Call BORRAREGISTROS(TABLA1$, "FECHA < '" & FETEXCOR1(HOY(HO$)) & "'", REGAF&, "NOMESS")
       If CantRegistros(TABLA1$, "FECHA < '" & FETEXCOR1(HOY(HO$)) & "'", "NOMESS") > 0 Then
          VUELTA% = VUELTA% + 1
          If VUELTA% > 5 Then
             Call COMUNI("Imposible Borrar Tabla: " & TABLA1$ & " Vuelva a Intentar ")
          End If
          GoTo 40
       End If
       '
       Call T_Espera(1 / 2)
       FE% = HOY(HO$)
       'CARGO LOS DATOS EN LA TABLA
       
       SQLX$ = "SELECT * FROM " & TABLA1$
       Dim RS As ADODB.Recordset
       Set RS = New ADODB.Recordset
65     On Error Resume Next
       RS.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
       '
       If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
           On Error GoTo 0
           Call CapturaErrorOpen
           GoTo 65
       End If
       On Error GoTo 0
       With RS
           For i& = 1 To MSF.Rows - 1 'RECORRE LA FILA
           
              If SOLO_SELECCIONADOS% > 0 Then
                   If MSF.TextMatrix(i&, 0) <> "*" Then
                      GoTo 70 'NEXT
                   End If
              End If
              
             .AddNew
             !IDCONSUL = IDCONSUL
             !FECHA = CVDAT(FE%)
             If APARTIR_COLUM% = 1 Then
                For J& = APARTIR_COLUM% To MSF.Cols - 1 'RECORRE AHORA CADA CELDA
                   Select Case J&
                     Case 1: !CAMPO1 = TRIM$(MSF.TextMatrix(i&, J&)) 'GRABA SEGUN LA CELDA QUE LEE
                     Case 2: !campo2 = TRIM$(MSF.TextMatrix(i&, J&))
                     Case 3: !CAMPO3 = TRIM$(MSF.TextMatrix(i&, J&))
                     Case 4: !CAMPO4 = TRIM$(MSF.TextMatrix(i&, J&))
                     Case 5: !CAMPO5 = TRIM$(MSF.TextMatrix(i&, J&))
                     Case 6: !CAMPO6 = TRIM$(MSF.TextMatrix(i&, J&))
                     Case 7: !CAMPO7 = TRIM$(MSF.TextMatrix(i&, J&))
                     Case 8: !CAMPO8 = TRIM$(MSF.TextMatrix(i&, J&))
                     Case 9: !CAMPO9 = TRIM$(MSF.TextMatrix(i&, J&))
                     Case 10: !CAMPO10 = TRIM$(MSF.TextMatrix(i&, J&))
                     Case 11: !CAMPO11 = TRIM$(MSF.TextMatrix(i&, J&))
                     Case 12: !CAMPO12 = TRIM$(MSF.TextMatrix(i&, J&))
                     Case 13: !CAMPO13 = TRIM$(MSF.TextMatrix(i&, J&))
                     Case 14: !CAMPO14 = TRIM$(MSF.TextMatrix(i&, J&))
                     Case 15: !CAMPO15 = TRIM$(MSF.TextMatrix(i&, J&))
                     Case 16: !CAMPO16 = TRIM$(MSF.TextMatrix(i&, J&))
                     Case 17: !CAMPO17 = TRIM$(MSF.TextMatrix(i&, J&))
                     Case 18: !CAMPO18 = TRIM$(MSF.TextMatrix(i&, J&))
                     Case 19: !CAMPO19 = TRIM$(MSF.TextMatrix(i&, J&))
                     Case 20: !CAMPO20 = TRIM$(MSF.TextMatrix(i&, J&))
                     Case 21: !CAMPO21 = TRIM$(MSF.TextMatrix(i&, J&))
                     Case 22: !CAMPO22 = TRIM$(MSF.TextMatrix(i&, J&))
                     Case 23: !CAMPO23 = TRIM$(MSF.TextMatrix(i&, J&))
                     Case 24: !CAMPO24 = TRIM$(MSF.TextMatrix(i&, J&))
                     Case 25: !CAMPO25 = TRIM$(MSF.TextMatrix(i&, J&))
                     Case 26: !CAMPO26 = TRIM$(MSF.TextMatrix(i&, J&))
                     Case 27: !CAMPO27 = TRIM$(MSF.TextMatrix(i&, J&))
                     Case 28: !CAMPO28 = TRIM$(MSF.TextMatrix(i&, J&))
                     Case 29: !CAMPO29 = TRIM$(MSF.TextMatrix(i&, J&))
                     Case 30: !CAMPO30 = TRIM$(MSF.TextMatrix(i&, J&))
                     Case 31: !CAMPO31 = TRIM$(MSF.TextMatrix(i&, J&))

                   End Select
                Next J&
             ElseIf APARTIR_COLUM% = 0 Then
                    For J& = APARTIR_COLUM% To MSF.Cols - 1 'RECORRE AHORA CADA CELDA
                       Select Case J&
                         Case 0: !CAMPO1 = TRIM$(MSF.TextMatrix(i&, J&))
                         Case 1: !campo2 = TRIM$(MSF.TextMatrix(i&, J&)) 'GRABA SEGUN LA CELDA QUE LEE
                         Case 2: !CAMPO3 = TRIM$(MSF.TextMatrix(i&, J&))
                         Case 3: !CAMPO4 = TRIM$(MSF.TextMatrix(i&, J&))
                         Case 4: !CAMPO5 = TRIM$(MSF.TextMatrix(i&, J&))
                         Case 5: !CAMPO6 = TRIM$(MSF.TextMatrix(i&, J&))
                         Case 6: !CAMPO7 = TRIM$(MSF.TextMatrix(i&, J&))
                         Case 7: !CAMPO8 = TRIM$(MSF.TextMatrix(i&, J&))
                         Case 8: !CAMPO9 = TRIM$(MSF.TextMatrix(i&, J&))
                         Case 9: !CAMPO10 = TRIM$(MSF.TextMatrix(i&, J&))
                         Case 10: !CAMPO11 = TRIM$(MSF.TextMatrix(i&, J&))
                         Case 11: !CAMPO12 = TRIM$(MSF.TextMatrix(i&, J&))
                         Case 12: !CAMPO13 = TRIM$(MSF.TextMatrix(i&, J&))
                         Case 13: !CAMPO14 = TRIM$(MSF.TextMatrix(i&, J&))
                         Case 14: !CAMPO15 = TRIM$(MSF.TextMatrix(i&, J&))
                         Case 15: !CAMPO16 = TRIM$(MSF.TextMatrix(i&, J&))
                         Case 16: !CAMPO17 = TRIM$(MSF.TextMatrix(i&, J&))
                         Case 17: !CAMPO18 = TRIM$(MSF.TextMatrix(i&, J&))
                         Case 18: !CAMPO19 = TRIM$(MSF.TextMatrix(i&, J&))
                         Case 19: !CAMPO20 = TRIM$(MSF.TextMatrix(i&, J&))
                         Case 20: !CAMPO21 = TRIM$(MSF.TextMatrix(i&, J&))
                         Case 21: !CAMPO22 = TRIM$(MSF.TextMatrix(i&, J&))
                         Case 22: !CAMPO23 = TRIM$(MSF.TextMatrix(i&, J&))
                         Case 23: !CAMPO24 = TRIM$(MSF.TextMatrix(i&, J&))
                         Case 24: !CAMPO25 = TRIM$(MSF.TextMatrix(i&, J&))
                         Case 25: !CAMPO26 = TRIM$(MSF.TextMatrix(i&, J&))
                         Case 26: !CAMPO27 = TRIM$(MSF.TextMatrix(i&, J&))
                         Case 27: !CAMPO28 = TRIM$(MSF.TextMatrix(i&, J&))
                         Case 28: !CAMPO29 = TRIM$(MSF.TextMatrix(i&, J&))
                         Case 29: !CAMPO30 = TRIM$(MSF.TextMatrix(i&, J&))
                         Case 30: !CAMPO31 = TRIM$(MSF.TextMatrix(i&, J&))
                         
                       End Select
                Next J&
             End If
             .Update
70         Next i&
       End With
       
       RS.Close
       Set RS = Nothing
       
       Call BORRAREGISTROS(TABLA1$, "IDCONSUL = " & IDCONSUL & " AND CAMPO1 = 'BORRAR'", REGAF&, "NOMESS")

End Sub
Function NAMEUSERSQL$(NUSER%)
        NAMEUSERSQL$ = VALOBADA("USUARIOSSQL", "NOMBRE", "NUSER = " & NUSER%, "NOMESS")
End Function

Function ORFASEL_ENVAPLAST$(Optional INCLUIR_CERRADAS%, Optional AgregaAnul%)
    If IsFormLoaded(FRMZELECHO) Then
        ESTAENMEMORIA = 1
    Else
        Call HookForm(FRMZELECHO.MSF2) 'estaba tirando un error cuando se cierra el formulario con doble clic
        Call HookForm(FRMZELECHO.MSF1) 'probablemente por que hace un hide, SI NO ESTA EN MEMORIA SE APLICA LA FUNCION.
    End If

    Call CARGA_ENCABEZADO(FRMZELECHO.MSF2, "O.Fabricación/a16;Código/A16;Descripción/A24;Cant.Proy/F10.1;Cant.Aprob./F10.1;Nro.Ped./f7;Cliente/f7;R.Social/a24;Ord.Item/F8.0", FRMZELECHO)
    Call CARGA_ENCABEZADO(FRMZELECHO.MSF1, "O.Fabricación/a16;Código/A16;Descripción/A24;Cant.Proy/F10.1;Cant.Aprob./F10.1;Nro.Ped./f7;Cliente/f7;R.Social/a24;Ord.Item/F8.0", FRMZELECHO)
    
    FRMZELECHO.Caption = "Ordenes de Fabricación Pendientes de Cierre"
    If INCLUIR_CERRADAS% > 0 Then FRMZELECHO.Caption = "Ordenes de Fabricación Pendientes y Cerradas"
    Dim obj As New RECORXET
    Set rs1 = obj.RS_ORFPEN_ENVA(INCLUIR_CERRADAS%, AgregaAnul%)
    Call Carga_MSF_Recordset(rs1, "O.Fabricación/a16;Código/A16;Descripción/A24;Cant.Proy/F8.1;Cant.Aprob./F8.1;Nro.Ped./f7;Cliente/f7;R.Social/a24;Ord.Item/F8.0", FRMZELECHO.MSF2, 1)
    Set obj = Nothing
    rs1.Close
    Set rs1 = Nothing
    
    FRMZELECHO.Width = 13500
    FRMZELECHO.Show 1
    
    ORFASEL_ENVAPLAST$ = TRIM$(RESP_ZELE_VECT(1))
End Function


Function OrdFabSegunPedido$(NROPED&, CI1%, Optional NORIT%)
    CONDI$ = "NroPed  =" & NROPED& & " AND Cod_Inte =" & CI1%
    REGAF& = CantRegistros&("ORDEFABR", CONDI$, "NOMESS")
    If REGAF& > 1 And NORIT% > 0 Then
       CONDI$ = CONDI$ + " AND NuOrProd = " & NORIT%
       OrdFabSegunPedido$ = VALOBADA("ORDEFABR", "IdSubOr", CONDI$)
    Else
      OrdFabSegunPedido$ = VALOBADA("ORDEFABR", "IdSubOr", CONDI$)
    End If
End Function
Function PrecioUnitarioSegunPedido#(NROPED&, CI1%, Optional NordenPedido%)
   CONDI$ = "NROPED = " & NROPED&
   CONDI$ = CONDI$ + " AND CODIINT = " & CI1%
   If NordenPedido% > 0 Then CONDI$ = CONDI$ + " AND NUORITEM = " & NordenPedido%
   
   
   PrecioUnitarioSegunPedido# = XVALO(VALOBADA("PEDDETA", "PREUNID", CONDI$, "NOMESS"))
End Function
Sub EliminarSQF(NameSQF$)
    ARCHIX$ = NameSQF$
    CONDI$ = "COD_QUERY = '" & ARCHIX$ & "'"
    Call BORRAREGISTROS("SQL_QUERIES", CONDI$, REGAF&, "NOMESS")
    
On Error GoTo ERRORSQLCTRLMDB
    Call ABRETACON
    strSQL = "SELECT * FROM SQL_QUERIES where COD_QUERY = '" & ARCHIX$ & "'"
    Set MIRS = BaDaConsul.OpenRecordset(strSQL, dbOpenDynaset)
    If Not (MIRS.EOF And MIRS.BOF) Then
       MIRS.Delete
    End If
ERRORSQLCTRLMDB:
   Call COMUNI("Error De Apertura de Base de Datos SQCTRL\Informe esta Situación a su Soporte de Sistemas")
End Sub
Sub CambiaGriddeDiseñadorFRM(archivo$, NombreConsOriginal$, NombreConsNuevo$)
  'ESTA RUTINA SE USA CUANDO SE LLAMA A MOSTRAR UN SQR (DISEÑADOR FRM) QUE GENERAR UNA CONSULTA EN LA TABLA_IMPRE
'ej.:NombreConsOriginal$ = SITUXOF
    'NombreConsNuevo$ = SITUX03
    'ARCHIVO ES EL ARCHIVO QUE DEFIENE EL DISEÑADOR DE FRM EJ. OFFT-ENV.SQR
    
    '******FUNCIONA CORRECTAMENTE EL TEMA ES QUE EL ARCHIVO SI TIENE GUARADA UNA IMAGEN SE LA BORRA POR QUE NO
    'LA VUELVE A COPIAR CORRECTAMENTE.******

    RUTAX$ = LUDAT$
    rutayarchivo$ = RUTAX$ + archivo$
    NARCHI1% = FILEOPEN%(rutayarchivo$, 1, 1512) ' ABRE EL ARCHIVO DE ESCRITURA
    'recorro el archivo cuando llego a la fila del grid cambio a cual consulta deberia llamar
    Do While Not EOF(NARCHI1%) ' RECORRE
       Line Input #NARCHI1%, XX$ ' LEE CADA REGISTRO
       If Left(XX$, 12) = "TIPO = GRID|" Then
         If InStr(XX$, NombreConsOriginal$) > 0 Then
            X$ = REPLACAR(XX$, NombreConsOriginal$, NombreConsNuevo$)
            XX$ = X$
         End If
       End If
       TX$ = TX$ + XX$ + vbCrLf
   Loop
'   Print #NARCHI2%, TX$ ' ESCRIBE EN EL ARCHIVO EL NUEVO TEXTO
   Close #NARCHI1%
   NARCHI2% = FILEOPEN%(rutayarchivo$, 2, 512) 'ABRE EL ARCHIVO DE LECTURA
   Print #NARCHI2%, TX$
   Close #NARCHI2%
End Sub
Sub EnvaplastAgregaItemsAlista1(n%)
     'SI LA EMPRESA ES ENVAPLAST
     'LE ASIGNA A TODOS LOS CLIENTES LA LISTA 1
     'CARGA TODOS LOS ITEMS EN LA LISTA DE PRECIOS EN LA LISTA 1 CON PRECIO UNITARIO 0.0001
     EPAC$ = TRIM$(UCase$(EMPREAC$))
     If InStr(EPAC$, "ENVAPLAST") < 1 Then Exit Sub
     If n% < 1 Then n% = 1
     
     If CantRegistros("LISTAPRE", "COD_LISTA = 1", "NOMESS") < 1 Then
        Call COMUNI("Debe Crear la Lista de Precios Nro 1")
        Exit Sub
     End If
     '
     If CantRegistros("LISTAPRE", "DECIMALES >= 3", "NOMESS") < 1 Then
        Call COMUNI("A Lista 1 Debe Asignar 3 o 4 Decimales")
        Exit Sub
     End If
     '
     '
     ARCHIPRE$ = TRIM$(Str$(n%))
     While Len(ARCHIPRE$) < 3: ARCHIPRE$ = "0" + ARCHIPRE$: Wend
     ARCHIPRE$ = "LIPRE" + ARCHIPRE$
     Call COPYSTRU("LIPREVEN", ARCHIPRE$)
     LF2% = LENFIELD(ARCHIPRE$, 2)
     If INCLUALIAS% = 0 Then
       INCLUALIAS% = (-1)
       If Control$("", "IDSECUN") = "ALIAS" Then INCLUALIAS% = 1
     End If
     '
     CADECI% = XVALO(VALOBADA("LISTAPRE", "DECIMALES", "cod_lista=" & n%))
     If CADECI% > 4 Then CADECI% = 4
   '
     Screen.MousePointer = 11


     SQLX$ = "UPDATE DEUDOR12 SET Lpre_Cli = 1 WHERE Lpre_Cli < 1 OR Lpre_Cli IS NULL" 'SI TIENE ALGUNA LISTA ASIGNADA NO APLICA.
     FLEXCONN.Execute (SQLX$)
     '
     SQLX$ = "Select m.codint,codigo,DESCRIPCION from master m WITH(NOLOCK)"
     SQLX$ = SQLX$ + " where STATUS >= 0 AND  not exists (select 1 from preventa where preventa.Cod_Pieza  = m.codint) ORDER BY m.codint"
     Set RS = READSET(SQLX$)
     Dim CII1%(), CODX$(), DESX$()
     With RS
        Do While Not .EOF
           i& = i& + 1
           ReDim Preserve CII1%(i&), CODX$(i&), DESX$(i&)
           CII1%(i&) = XVALO(!CODINT)
           CODX$(i&) = SAFETEXT$(!Codigo)
           DESX$(i&) = SAFETEXT$(!Descripcion)
           .MoveNext
        Loop
        RS.Close
        Set RS = Nothing
        '
        If i& < 1 Then Screen.MousePointer = 1: Exit Sub  'SI NO HAY ITEMS SIN GRABAR EN LA LISTA SALE.
        '
        On Error GoTo salir
        
        MAXORDEN = XVALO(VALOBADA("PREVENTA", "MAX(Num_Orden)", "COD_PIEZA > 0 ", "NOMESS"))
        For i& = 1 To UBound(CII1%)
           SQLX$ = " INSERT INTO PREVENTA "
           SQLX$ = SQLX$ + "(Codiempr,Cod_Lista,Cod_Pieza,Num_Orden,Desc_Pieza,Denom_Pieza,Carac_Control,Num_Original "
           SQLX$ = SQLX$ + ",Precio_Unitario, Cant_Limite,Porc_Limite)"
           SQLX$ = SQLX$ + " VALUES("
           SQLX$ = SQLX$ + "" & CodiEmp%
           SQLX$ = SQLX$ + ",1"
           SQLX$ = SQLX$ + ", " & CII1%(i&)
           SQLX$ = SQLX$ + ", " & MAXORDEN + i&
           SQLX$ = SQLX$ + ",'" & DESX$(i&) & "'"
           SQLX$ = SQLX$ + ",'" & DESX$(i&) & "'"
           SQLX$ = SQLX$ + ",''"
           SQLX$ = SQLX$ + ",0"
           SQLX$ = SQLX$ + ",0.0010"
           SQLX$ = SQLX$ + ",0"
           SQLX$ = SQLX$ + ",0"
           SQLX$ = SQLX$ + ")"
           FLEXCONN.Execute (SQLX$)
           
           TTXX$ = REGBLAN$(ARCHIPRE$)
           CII% = CII1%(i&)
           Call REPLA(TTXX$, MKI$(CII%), 1, 2)
           Call REPLA(TTXX$, DESX$(i&), 3, LF2%)
           MRX$ = ""
           Call REPLA(TTXX$, MRX$, 3 + LF2%, 1)
           Call REPLA(TTXX$, "", 4 + LF2%, 16)
               PRUNI# = XVALO(0.001)
            Call REPLA(TTXX$, MKD$(PRUNI#), 20 + LF2%, 8)
            Call REPLA(TTXX$, MKS$(0), 28 + LF2%, 2)
            Call REPLA(TTXX$, MKD$(0), 30 + LF2%, 1)
            '
            PRELI# = PRUNI#
            CDC$ = TRIM$(Str$(CADECI%))
            PREX$ = " " + FORMATNUM$(PRELI#, "F12." + CDC$) + "  "
            Call REPLA(TTXX$, PREX$, 1 + LF2% - Len(PREX$), 2 + Len(PREX$))
            Call REPLA(TTXX$, MRX$, 2 + LF2%, 1)
            '
            If INCLUALIAS% > 0 Then
              ALX$ = TRIM$(NUPLANO$(CII%))
              If ALX$ <> "" Then
                Call REPLA(TTXX$, " (" + ALX$ + ")", LF2% - 16 - Len(ALX$), 3 + Len(ALX$))
              End If
            End If
            jj& = jj& + 1
            'Call GRAREG(ARCHIPRE$, TTXX$, JJ&)

           
            Call REGAPP(ARCHIPRE$, TTXX$)
        Next i&
     End With
     Screen.MousePointer = 1
     UREG& = ULTREG&(ARCHIPRE$)
     Call SETULTREG(ARCHIPRE$, UREG&)
     Call CIERRARCH(ARCHIPRE$)
     Call FRESHECHO("!ARCHIPRE$")


salir:
End Sub
Function MEDIDA_ENVAPLAST$(CIXI%, RET_CORTE#)
      RET_CORTE# = 0
      MEDI$ = ""
      SQLX$ = "SELECT TOP 1 CORTE,Descripcion"
      SQLX$ = SQLX$ + " FROM ESPECIFIC WITH(NOLOCK) "
      SQLX$ = SQLX$ + " WHERE COD_INTE = " & CIXI% & " ORDER BY ID"
      J& = 0
      Set RS = READSET(SQLX$)
      With RS
        Do While Not .EOF
            J& = J& + 1
            If J& = 1 Then
              textox$ = SAFETEXT$(!Descripcion)
              X = TEXTOLOTES$(textox$, "X")
              On Error Resume Next
              For k% = 1 To UBound(CADENATEX$)
                If k% < UBound(CADENATEX$) Then
                  If k% > 1 Then
                    MEDI$ = MEDI$ & " x " & CADENATEX$(k%)
                  Else
                    MEDI$ = CADENATEX$(k%)
                  End If
                End If
              Next k%
              RET_CORTE# = TRIM$(FORMATNUM$(XVALO(!corte), "F10.2"))
              On Error GoTo 0
           End If
           .MoveNext
        Loop
        
      End With
      MEDIDA_ENVAPLAST$ = MEDI$
      RS.Close
      Set RS = Nothing
End Function

Function VersionBD%()
    'SQLX$ = "SELECT @@VERSION AS 'SQL Server Version';"
    SQLX$ = "SELECT @@VERSION"
    Set RS = FLEXCONN.Execute(SQLX$)
    VERSIOX$ = TRIM$(RS.Fields(0))
    'If XVALO(Mid$(VERSIOX$, 22)) <> 2000 Then ' VALIDO QUE SEA > a la version 2000 el varchar(max) da error
    'Microsoft SQL Server 2017 (RTM) - 14.0.1000.169 (X64)   Aug 22 2017 17:04:49   Copyright (C) 2017 Microsoft Corporation  Express Edition (64-bit) on Windows Server 2012 Datacenter 6.2 <X64> (Build 9200: )
    VersionBD% = XVALO(Mid$(VERSIOX$, 22, 4))
    
End Function
Sub CARGARCOMBOUSERTEL(CMB As comboBox)
   'CARGA UN COMBO SEGUN USUARIOS DE LISTA USUARIOS SQL
   CMB.Clear
   SQLX$ = "SELECT CAMPO5,CAMPO3 FROM DATASQL WITH(NOLOCK) WHERE LISTA = 'USUARIOSQL' ORDER BY CAMPO5"
   Set RS = READSET(SQLX$)
   With RS
    Do While Not .EOF
       Call REPLA(TX$, UCase$(SAFETEXT$(!CAMPO5)), 1, 6)
       Call REPLA(TX$, UCase$(SAFETEXT$(!CAMPO3)), 8, 24)
       CMB.AddItem TX$
      .MoveNext
    Loop
   End With
   
  
End Sub
Function VALOADMIN(NOTAB$, NOCAM$, CONDI$, Optional MODOMESS$)

    VALOADMIN = VALOBADA(NOTAB$, NOCAM$, CONDI$, MODOMESS$)

End Function
Function MAXNUMCAMPO&(TABLA$, CAMPO$, Optional CONDI$)
   'SIN NOMESS PARA ESTAR CONCIENTE SI NO GENERA EL SIGUIENTE NUMERO CORRECTAMENTE
   If TRIM$(CONDI$) = "" Then
      CONDI$ = CAMPO$ & " > 0"
   End If
   Screen.MousePointer = 11
   MAXNUMCAMPO& = XVALO(VALOBADA(TABLA$, "MAX(" & CAMPO$ & ")", CONDI$))
   Screen.MousePointer = 1
   
End Function
Function SELECGRUPOACTU()
   
   On Error Resume Next
   RESP_ZELE_VECT(1) = ""
   On Error GoTo 0
   Campos$ = "ID GRUPO/F10;REFERENCIA/A48"
   Call CARGA_ENCABEZADO(FRMZELECHO.MSF2, Campos$, FRMZELECHO)
   Call CARGA_ENCABEZADO(FRMZELECHO.MSF1, Campos$, FRMZELECHO)
   FRMZELECHO.Caption = ""
 
   Dim obj As New RECORXET
   Set RS = obj.RS_GRUPOACTU()
   Call Carga_MSF_Recordset(RS, Campos$, FRMZELECHO.MSF2, 1)
   Set obj = Nothing
   RS.Close
   Set RS = Nothing

   FRMZELECHO.Width = 8500
   FRMZELECHO.Show 1
   If TRIM$(RESP_ZELE_VECT(1)) <> "" Then
     SELECGRUPOACTU = TRIM$(RESP_ZELE_VECT(1))
   End If
End Function
Sub ASOCIAGACTU()
    If DERACCE%("MODCOSTO/NOMESS", "Modificar Costos y Precios") < 2 Then Exit Sub
    'ASOCIA ITEMS A UN GRUPO DE ACTUALIZACION
20  Campos$ = "CODINT/F8;Código/A20;Descripción/A40;Cód.Grupo/F10;Referencia/A24"
    Campos$ = "CODINT/F8;Código/A20;Descripción/A40;Cód.Grupo/F10;Referencia/A24;M./F0;Costo Rep/F0"
    Call CARGA_ENCABEZADO(SELFAMILI.GRID, Campos$, SELFAMILI, 1, 0, ANTOT, , 1)
    Call CARGA_ENCABEZADO(SELFAMILI.MSF_2, Campos$, SELFAMILI, 1, 0, ANTOT, , 1)
    'SELFAMILI.Command3.Caption = "Aceptar"
    SELFAMILI.Command3.Visible = False
    SELFAMILI.mnuActuCostos.Visible = True
    SELFAMILI.Command3.ToolTipText = "Asociar Selección a Grupo de Actualización"
    '
    Screen.MousePointer = 11
    Dim obj As New RECORXET
    Set rs1 = obj.RS_MasterComp(1, 0, 0)
    Call Carga_MSF_Recordset(rs1, Campos$, SELFAMILI.GRID, 1)
    Set obj = Nothing
    rs1.Close
    Set rs1 = Nothing
    Screen.MousePointer = 1
    'LE AGREGO LA REFERENCIA DEL GRUPO DIRECTAEMNTE A LA GRILLA.
    For i& = 1 To SELFAMILI.GRID.Rows - 1
        GRACT% = XVALO(SELFAMILI.GRID.TextMatrix(i&, 4))
        SELFAMILI.GRID.TextMatrix(i&, 5) = ""
        If GRACT% > 0 Then
           SELFAMILI.GRID.TextMatrix(i&, 5) = REFEGRUPOACTU$(GRACT%)
        End If
    Next i&
    Screen.MousePointer = 1
    SELFAMILI.Show 1

    If SELFAMILI.HaySeleccion% > 0 Then
       SELFAMILI.HaySeleccion% = 0
       GRACT% = SELECGRUPOACTU
       Screen.MousePointer = 11
       REFEGRUP$ = TRIM$(RESP_ZELE_VECT(2))
       If GRACT% < 1 Then
          OPX% = COMALTER%("Atención No se Ha Seleccionado Ningún Grupo de Actualización\Si Selecciona Grabar y Blanquear los Item Seleccionados Borrará el Dato Actual\\Cancelar\Grabar y Blanquear")
       Else
          OPX% = COMALTER%("Confirma Asociar Los Items Seleccionados al Grupo de Actualización: " & REFEGRUP$ & "\\Cancelar\Grabar")
       End If
       '
       Dim MSF As Object
       Set MSF = SELFAMILI.GRID
       If TRIM$(SELFAMILI.TXTBUSCAR) <> "" Then
          Set MSF = SELFAMILI.MSF_2
       End If
       ACTU& = 0
       If OPX% = 2 Then
          For i& = 1 To MSF.Rows - 1
            MARCA$ = TRIM$(MSF.TextMatrix(i&, 0))
            If MARCA$ = "*" Then
               ACTU& = ACTU& + 1
               CI1% = XVALO(MSF.TextMatrix(i&, 1))
               SQLX$ = "UPDATE MASTER SET GR_ACTU = " & GRACT%
               SQLX$ = SQLX$ + " WHERE CODINT =  " & CI1%
               FLEXCONN.Execute (SQLX$)
            End If
          Next i&
          Screen.MousePointer = 1
          Call COMUNI("Se Han Actualizado Los Códigos Seleccionados (" & ACTU& & ")")
          Exit Sub

       Else
          Screen.MousePointer = 1
          SELFAMILI.HaySeleccion% = 0
          Exit Sub
       End If
    End If
    Screen.MousePointer = 1
End Sub
Function NOMBRE_TEC_EXT$(NUMTEC&)
  If NUMTEC& < 1 Then NOMBRE_TEC_EXT$ = ""
  NOMBRE_TEC_EXT$ = VALOBADA("DATASQL", "CAMPO2", "LISTA = 'TECNIEXT' AND CAMPO1 = '" & FORMATOCON0$(NUMTEC&, 9) & "'", "NOMESS")
End Function
Function NCLI_ASOC_TECNICO(NUREP&)
  NUMTEC& = TECNICO_REPA&(NUREP&)
  'DEVUELVE EL CLIENTE ASOCIADO AL TECNICO
  NCLI_ASOC_TECNICO = XVALO(VALOBADA("DATASQL", "CAMPO3", "LISTA = 'TECNIEXT' AND CAMPO1 = '" & FORMATOCON0$(NUMTEC&, 9) & "'", "NOMESS"))
  
End Function

Function TECNICO_REPA&(NUREP&)
   CONDI$ = "NUMREPA = " & NUREP&
   TECNICO_REPA& = XVALO(VALOBADA("REPARA", "TEC_EXT", CONDI$, "NOMESS"))
End Function
Function SUMREPU(NROREP&)
    'SUMA LOS COSTOS TOMANDO EN CUENTA LAS SALIDAS DEL   MOVISTO
    'CUANDO SE REALIZA EL CIERRE DE LA  ORDEN EN EL CAMPO IdCalidad SE LE PONE EL NRO DE REPARACION.
    'Y ESTE LO USA A MODO DE FILTRO.
    SQLX$ = "SELECT MO.COD_INTE,MO.CANTSALID, M.COSCAL,M.COSUNI "
    SQLX$ = SQLX$ + " FROM MOVISTO MO INNER JOIN MASTER M WITH(NOLOCK) ON MO.COD_INTE = M.CODINT "
    SQLX$ = SQLX$ + " WHERE MO.IdCalidad = " & NROREP&
    Set rs1 = READSET(SQLX$)
    '
    ACUMUCOSCAL = 0
    With rs1
     Do While Not .EOF
       CI1% = XVALO(!cod_inte)
       CANTI = XVALO(!CANTSALID)
       COCAL = XVALO(!COSCAL): If COCAL <= 0 Then COCAL = XVALO(!COSUNI)
       ACUMUCOSCAL = ACUMUCOSCAL + (CANTI * COCAL)
       .MoveNext
     Loop
    End With
    rs1.Close
    Set rs1 = Nothing
    SUMREPU = ACUMUCOSCAL
End Function

Function SUMMOBRA(NREP&)
    SUMMOBRA = XVALO(VALOBADA("REPARA", "MOBRA_EN_GTIA", "NUMREPA = " & NREP&, "NOMESS"))
End Function

Function SUMMOBRA_TODAS(NREP&)
    MANO_OBRA_TRADICIONAL = XVALO(VALOBADA("REPUREPA", "TOT_MANO_OBRA", "NUMREPA = " & NUREP&, "NOMESS"))
    MANO_OBRA_EN_GARANTIA = XVALO(VALOBADA("REPUREPA", "MOBRA_EN_GTIA", "NUMREPA = " & NUREP&, "NOMESS"))
    SUMMOBRA_TODAS = MANO_OBRA_TRADICIONAL + MANO_OBRA_EN_GARANTIA
End Function

'
Sub GeneraPedidoDeRecepcionFacturada(NP, RETORNO_PEDIDO, facturaprov$, PorCjeGananciaAHP#, OCOMPRA&)
       
       If PorCjeGananciaAHP# = 0 Then
          PorCjeGananciaAHP# = Control("CFGMINEP", "PJGANAHP")
       End If
       'SE GENERA UN PEDIDO DESDE MINEPARTS PARA CLIENTE AHP
       'SE TOMAN LOS DATOS DE UNA RECEPCION QUE YA HA SIDO FACTUARDA
       'ESTA SE REALIZA ON DEMAND PARA QUE EL USUARIO LA REALICE CUANDO YA SE GENERARON LOS GASTOS PARA DICHA FACTURA
'Call mostrarDatosx("PorCjeGananciaAHP#: " & PorCjeGananciaAHP#)
'Call mostrarDatosx("COEF_GANANCIA#: " & ROUND(1 + (PorCjeGananciaAHP# / 100), 2))


       NCAHP = XVALO(Control("CFGMINEP", "CLTEMINE")) 'CLIENTE AHP CONFIGURADO
       CONDI$ = "facturaprov = '" & facturaprov$ & "' And Nume_cli = " & NP & " AND nume_ocom = " & OCOMPRA&
'Call mostrarDatosx("CONDI$: " & CONDI$)

       MONEMIX% = XVALO(VALOBADA("OCOMENCA", "Mone_emis", "NUME_OCOM = " & OCOMPRA&, "NOMESS"))
'Call mostrarDatosx("MONEMIX%: " & MONEMIX%)
       OKX% = 0
       RETORNO_PEDIDO = 0
       RSOC$ = TRIM$(rasocli$(-1)) 'PARA REFRESCAR EL RECORDSET
       FEMI% = HOY(HO$)
       Call ABREPEDCLI
       '
       On Error GoTo ERROR_GUARDAR
       FLEXCONN.BeginTrans
       
       Screen.MousePointer = 11
       FEMI% = HOY(HO$)
       FEX$ = FECHATEX$(FEMI%)
       Dim RS As ADODB.Recordset
       Set RS = New ADODB.Recordset
       SQLX$ = "SELECT * FROM  PEDENCA  "
       SQLX$ = SQLX$ + " WHERE NroClie = -1"
25     On Error Resume Next
       RS.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
       If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
              On Error GoTo 0
              Call CapturaErrorOpen
              GoTo 25
           End If
           On Error GoTo 0
       With RS
         .AddNew
         !CODIEMPR = CodiEmp%
         NROPEDI1& = NUPEDIDO& ' FUNCION

         !NROPED = NROPEDI1&
         !FeSolEnt = CVDAT(FEMI%)
         RSOC$ = TRIM$(rasocli$(NCAHP))
         REFE$ = ""
         If OCOMPRA& > 0 Then REFE$ = facturaprov$
         If REFE$ <> "" Then
            REFE$ = REFE$ + " - " + FECHATEX$(FEMI%)
         Else
            REFE$ = FECHATEX$(FEMI%)
         End If
         FEX = FEMI%
         LOMA% = 48 - Len(REFE$)
         REFE$ = Left$(rasocli$(-1 * NP), LOMA%) + " - " & REFE$
         
         !REFERENCIA = Left$(REFE$, 128)
         !FECHEMIS = CVDAT(FEMI%)
         !NroClie = NCAHP
         !RasoClie = RSOC$
         !NROOCOM = Left$(TRIM$(OCOMPRA&), 16)
         !DOMIENT = ""
         !LOCAENT = ""
         !TranspEnt = ""
         !NTranspo = 0
         !NSucent = 0
         !Observa = ""
         !Status = 0
         !ListPre = 0
         !CondPag = 0
         !VENDED = 0
         !TDescApli = ""
         .Update
       End With
       RS.Close
       Set RS = Nothing
       
       ANOTA$ = ""
       TTXYZ$ = Space$(76)
       HOII% = HOY(HOS$)
       Call AGREGA_ANOTAPED_NEW(NROPEDI1&, "Generación de Orden de Venta:")
       IKITEM% = 0
       '
       
       SQLX$ = "SELECT * FROM  PEDDETA  "
       SQLX$ = SQLX$ + " WHERE NROPED = " & NROPEDI1&
       
       Set RsPeddeta = New ADODB.Recordset
30     On Error Resume Next
       RsPeddeta.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
       If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
              On Error GoTo 0
              Call CapturaErrorOpen
              GoTo 30
       End If
       On Error GoTo 0
       
       IKITEM% = 0
       ConsultaSql$ = "SELECT * FROM detfacomsql with(nolock)"
       ConsultaSql$ = ConsultaSql$ + " where " & CONDI$
       ConsultaSql$ = ConsultaSql$ + " order by NordenBrItem"
'Call mostrarDatosx("ConsultaSql$: " & ConsultaSql$)
       IKITEM% = 0
       Set rs1 = READSET(ConsultaSql$)
       With rs1
          Do While Not .EOF
                CIXI% = XVALO(!cod_inte)
                CANTI = XVALO(!CantFacturada)
                NBREC& = XVALO(!NUMEBORE)
                LOTEX1$ = "BR-" & FORMATOCON0(NBREC&, 6) & "-" & FORMATOCON0(XVALO(!NordenBrItem), 2)
                COSFOB# = XVALO(!PreunitS_Factura)
'Call mostrarDatosx("COSFOB#: " & ROUND(COSFOB#, 2))

                PORCEN_GASNAC# = PORCENT_COSTO_NACIONALIZACION(NBREC&, MONEMIX%)
'Call mostrarDatosx("PORCEN_GASNAC#: " & ROUND(PORCEN_GASNAC#, 2))
                preunx = XVALO(COSFOB# * PORCEN_GASNAC#)
                COEF_GANANCIA# = 1 + (PorCjeGananciaAHP# / 100)
                preunx = preunx * COEF_GANANCIA#
'Call mostrarDatosx(CODEXT$(CIXI%) & "  preunx: " & ROUND(preunx, 2))
                IKITEM% = IKITEM% + 1
'If IKITEM% > 1 Then RsPeddeta.Close: Exit Sub
                With RsPeddeta
                     .AddNew
                     !CODIEMPR = CodiEmp%
                     IKITEM% = IKITEM% + 1
                     !NROPED = NROPEDI1&
                     !FECHEMIS = CVDAT(FEMI%)
                     !NroClie = NCAHP
                     !RasoClie = Left$(TRIM$(rasocli$(NCAHP)), 64)
                     !NROOCOM = Left$(TRIM$(OCOMPRA&), 16)
                     !FeSolEnt = CVDAT(FEMI%)
                     !Status = 0
                     !ListPre = 0
                     !CondPag = 0
                     !VENDED = 0
                     !PorComVend = 0
                     !NuOrItem = IKITEM%
                     !CODIINT = CIXI%
                     !CODIEXT = CODEXT$(CIXI%)
                     !SubCod = ""
                     !Descrip = DESCRIT0$(CIXI%)
                     !TLargo = 0
                     !TAncho = 0
                     !TEspesor = 0
                     !CBase = ""
                     !TTALLE = ""
                     !TCOLOR = ""
                     !VenRep = ""   ' Venta o Reparacion
                     !MoneEmis = MONEMIX%
                     'PREUNI# = PREUNI
                     !TIPOCAM = TICAMONE(MONEMIX%)
                     
                     !PreList = preunx
                     !DescLista = 0
                     !DescCombi = ""
                     !DescRubro = 0
                     !DescTotal = 0
                     !CanAsig = 0
                     !CANPED = CANTI
                     !PreUnid = preunx
                     !ImpoTot = CANTI * preunx
                     !CANTENT = 0
                     !CanSaldo = 0
                     !FeUltEnt = CVDAT(0)
                     !NroRemi = ""
                     !TPedItem = ""
                     !idenlote = LOTEX1$
                     !Num_Prov = NP
                    .Update

                     '
                     TTXYZ$ = Space$(76)
                     Call REPLA(TTXYZ$, "- " + CODEXT$(CIXI%), 11, 18)
                     Call REPLA(TTXYZ$, DESCRIT0$(CIXI%), 29, 38)
                     Call REPLA(TTXYZ$, CSTRING$(MKS$(CANTI), 3, 10, 1, 2), 57, 10)
                     ANOTA$ = ANOTA$ + TTXYZ$ + Chr$(10) + Chr$(13)
                     
                     If LOTEX1$ <> "" Then
                      'GUARDO EN EL ANTOADOR EL NUMERO DE LOTE
                      TELOTE$ = "Lote Reservado: " + LOTEX1$
                      TELOTEXX$ = Space$(76)
                      Call REPLA(TELOTEXX$, TELOTE$, 29, 38)
                      ANOTA$ = ANOTA$ + TELOTEXX$ + Chr$(10) + Chr$(13)
                     End If
                     
                   End With
                   rs1.MoveNext
            Loop
          End With
          '
          On Error Resume Next
          RsPeddeta.Close
          Set RsPeddeta = Nothing
'          RS1.Close
'          Set RS1 = Nothing
          On Error GoTo 0
          'GUARDO EN EL ANTOADOR EL NUMERO DE PROV.
          NUMEPROVTEX$ = "N°Prov: " + TRIM$(Str$(NPROVX)) & "-" & facturaprov$ '(INT15-372)
          
          TELOTEXX$ = Space$(76)
          Call REPLA(TELOTEXX$, NUMEPROVTEX$, 29, 38)
          ANOTA$ = ANOTA$ + TELOTEXX$ + Chr$(10) + Chr$(13)

          Call AGREGA_ANOTAPED_NEW(NROPEDI1&, ANOTA$)
          '
          SQLX$ = "UPDATE DETFACOMSQL SET NROPED = " & NROPEDI1&
          SQLX$ = SQLX$ + " WHERE FACTURAPROV = '" & facturaprov$ & "'"
          SQLX$ = SQLX$ + " AND nume_ocom = " & OCOMPRA&
          FLEXCONN.Execute (SQLX$)
          
98        Screen.MousePointer = 1
          RETORNO_PEDIDO = NROPEDI1&
          
ERROR_GUARDAR:
    If Err.Number <> 0 Then
       FLEXCONN.RollbackTrans
       Call COMUNI("ERROR AL GUARDAR NO SE REALIZO LA TRANSACCION" & vbCrLf & Err.Number & " " & Err.Description)
       On Error GoTo 0
    Else
      FLEXCONN.CommitTrans
      OKX% = 1
      Call COMUNI("TRANSACCIÓN REALIZADA")
    End If


End Sub


Sub GENERAPEDIDO(NUREP&, REMI$, RETORNO_PEDIDO, OKX%, Optional PIVA%)
       NTECNICO% = TECNICO_REPA&(NUREP&)
       NC_ASOC_TECNICO = NCLI_ASOC_TECNICO(NUREP&)
       
       OKX% = 0
       RETORNO_PEDIDO = 0
       Call ABREPEDCLI
       '
       Screen.MousePointer = 11
       APREC$ = "LIPRE" + "001"
       PRLST# = PRELISTA#(APREC$, CIXI%)
       FEMI% = HOY(HO$)
       FEX$ = FECHATEX$(FEMI%)
       Dim RS As ADODB.Recordset
       Set RS = New ADODB.Recordset
       SQLX$ = "SELECT * FROM  PEDENCA  "
       SQLX$ = SQLX$ + " WHERE NroClie = -1"
25     On Error Resume Next
       RS.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
       If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
              On Error GoTo 0
              Call CapturaErrorOpen
              GoTo 25
           End If
           On Error GoTo 0
       With RS
         .AddNew
         !CODIEMPR = CodiEmp%
         NROPEDI1& = NUPEDIDO& ' FUNCION
         !NROPED = NROPEDI1&
         !FeSolEnt = CVDAT(FEMI%)
         RSOC$ = rasocli$(NC_ASOC_TECNICO)
         TXTREFE$ = TRIM$(Left$(RSOC$, 24)) & " - " & FECHATEX$(FEMI%) & "(Rep.N°: " & NUREP& & ")"
         !REFERENCIA = Left$(TXTREFE$, 128)
         !FECHEMIS = CVDAT(FEMI%)
         !NroClie = NC_ASOC_TECNICO
         !RasoClie = RSOC$
         !NROOCOM = ""
         !DOMIENT = DOMICLI$(NC_ASOC_TECNICO)
         !LOCAENT = LOCACLI$(NC_ASOC_TECNICO)
         !TranspEnt = ""
         !NTranspo = 0
         !NSucent = 0
         !Observa = ""
         !Status = 0
         !ListPre = 0
         !CondPag = 0
         !VENDED = 0
         !TDescApli = ""
         .Update
       End With
       RS.Close
       Set RS = Nothing
       
       ANOTA$ = ""
       TTXYZ$ = Space$(76)
       HOII% = HOY(HOS$)
       Call REPLA(TTXYZ$, HOS$, 1, 8)
       Call REPLA(TTXYZ$, "Generación de Orden de Venta:", 11, 56)
       Call REPLA(TTXYZ$, USERTER$, 69, 24)
       ANOTA$ = ANOTA$ + TTXYZ$ + Chr$(10) + Chr$(13)
       IKITEM% = 0
       '
       SQLX$ = "SELECT * FROM  PEDDETA  "
       SQLX$ = SQLX$ + " WHERE NroClie = -1"
       Set RsPeddeta = New ADODB.Recordset
30     On Error Resume Next
       RsPeddeta.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
       If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
              On Error GoTo 0
              Call CapturaErrorOpen
              GoTo 30
       End If
       On Error GoTo 0
       
       ConsultaSql$ = "SELECT * FROM REPUREPA WITH(NOLOCK) WHERE NUMREPA=" & CStr(NUREP&)
       Set RsRepurepa = READSET(ConsultaSql$)
       With RsRepurepa
          Do While Not .EOF
                CIXI% = XVALO(!cod_inte)
                CANTI = XVALO(!Cant_Sol)
                PRECARGA = XVALO(!prun_neto)
                PREUNI = PRECARGA
                CAC# = 0
                PORDESCUIN = XVALO(!PORDESC)
                If CIXI% > 0 And CANTI > 0 Then
                   If PORDESCUIN > 0 Then
                     PREUNI = PREUNI * (100 - PORDESCUIN) / 100
                   End If
                With RsPeddeta
                     .AddNew
                     !CODIEMPR = CodiEmp%
                     IKITEM% = IKITEM% + 1
                     !NROPED = NROPEDI1&
                     !FECHEMIS = CVDAT(FEMI%)
                     !NroClie = NC_ASOC_TECNICO
                     !RasoClie = rasocli$(NC_ASOC_TECNICO)
                     !NROOCOM = ""
                     !FeSolEnt = CVDAT(FEMI%)
                     !Status = 0
                     !ListPre = 0
                     !CondPag = 0
                     !VENDED = 0
                     !PorComVend = 0
                     !NuOrItem = IKITEM%
                     !CODIINT = CIXI%
                     !CODIEXT = CODEXT$(CIXI%)
                     !SubCod = ""
                     !Descrip = DESCRIT0$(CIXI%)
                     !TLargo = 0
                     !TAncho = 0
                     !TEspesor = 0
                     !CBase = ""
                     !TTALLE = ""
                     !TCOLOR = ""
                     !VenRep = ""   ' Venta o Reparacion
                     !MoneEmis = 1
                     'PREUNI# = PREUNI
                     !TIPOCAM = 1
                     !PreList = PRECARGA
                     !DescLista = 0
                     !DescCombi = ""
                     !DescRubro = 0
                     !DescTotal = 0
                     !CanAsig = 0
                     !CANPED = CANTI
                     !PreUnid = PREUNI
                     !ImpoTot = PREUNI * CANTI
                     !CANTENT = 0
                     !CanSaldo = 0
                     !FeUltEnt = CVDAT(0)
                     !NroRemi = ""
                     !TPedItem = ""
                     '
                     TTXYZ$ = Space$(76)
                     Call REPLA(TTXYZ$, "- " + CODEXT$(CIXI%), 11, 18)
                     Call REPLA(TTXYZ$, DESCRIT0$(CIXI%), 29, 38)
                     Call REPLA(TTXYZ$, CSTRING$(MKS$(1), 3, 10, 1, 2), 57, 10)
                     ANOTA$ = ANOTA$ + TTXYZ$ + Chr$(10) + Chr$(13)
                    .Update
                   End With
                  End If
                  RsRepurepa.MoveNext
            Loop
          End With
          '
          On Error Resume Next
          RsPeddeta.Close
          Set RsPeddeta = Nothing
          RsRepurepa.Close
          Set RsRepurepa = Nothing
          On Error GoTo 0
          
          Call REPLA(TTXX$, "*** Técnico Externo ***", 11, 56)
          ANOTA$ = ANOTA$ + TTXX$ + Chr$(10) + Chr$(13)
          ANOTA$ = ANOTA$ + Space$(10) + String$(56, "-") + Chr$(10) + Chr$(13)
          SQLX$ = "SELECT ANOTAPED FROM PEDENCA WITH(NOLOCK)"
          SQLX$ = SQLX$ + " WHERE NroPed = " & NROPEDI1& & ""
          Set PEDDDENCA = READSET(SQLX$)
          ANOTB$ = SAFETEXT$(PEDDDENCA!AnotaPed)
          ANOTA$ = ANOTB$ + ANOTA$
          ANOTA$ = REPLACAR$(ANOTA$, "'", "")
          
          SQLX$ = "UPDATE PEDENCA SET AnotaPed=" & "'" & ANOTA$ & "' WHERE NroPed=" & NROPEDI1&
          FLEXCONN.Execute (SQLX$)
          '
          OKX% = 1
          '
          RETORNO_PEDIDO = NROPEDI1&
          

98        Screen.MousePointer = 1
   '
End Sub

Sub GENERAPEDIDO_VIEJO(NUREP&, REMI$, RETORNO_PEDIDO, OKX%, Optional PIVA%)
       NTECNICO% = TECNICO_REPA&(NUREP&)
       NC_ASOC_TECNICO = NCLI_ASOC_TECNICO(NUREP&)
       
       OKX% = 0
       RETORNO_PEDIDO = 0
       Call ABREPEDCLI
       '
       

       Screen.MousePointer = 11
       APREC$ = "LIPRE" + "001"
       PRLST# = PRELISTA#(APREC$, CIXI%)
       FEMI% = HOY(HO$)
       FEX$ = FECHATEX$(FEMI%)
       Dim RS As ADODB.Recordset
       Set RS = New ADODB.Recordset
       SQLX$ = "SELECT * FROM  PEDENCA  "
       SQLX$ = SQLX$ + " WHERE NroClie = -1"
25     On Error Resume Next
       RS.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
       If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
              On Error GoTo 0
              Call CapturaErrorOpen
              GoTo 25
           End If
           On Error GoTo 0
       With RS
         .AddNew
         !CODIEMPR = CodiEmp%
         NROPEDI1& = NUPEDIDO& ' FUNCION
         !NROPED = NROPEDI1&
         !FeSolEnt = CVDAT(FEMI%)
         RSOC$ = rasocli$(NC_ASOC_TECNICO)
         TXTREFE$ = TRIM$(Left$(RSOC$, 24)) & " - " & FECHATEX$(FEMI%) & "(Rep.N°: " & NUREP& & ")"
         !REFERENCIA = Left$(TXTREFE$, 128)
         !FECHEMIS = CVDAT(FEMI%)
         !NroClie = NC_ASOC_TECNICO
         !RasoClie = RSOC$
         !NROOCOM = ""
         !DOMIENT = DOMICLI$(NC_ASOC_TECNICO)
         !LOCAENT = LOCACLI$(NC_ASOC_TECNICO)
         !TranspEnt = ""
         !NTranspo = 0
         !NSucent = 0
         !Observa = ""
         !Status = 0
         !ListPre = 0
         !CondPag = 0
         !VENDED = 0
         !TDescApli = ""
         .Update
       End With
       RS.Close
       Set RS = Nothing
       
       ANOTA$ = ""
       TTXYZ$ = Space$(76)
       HOII% = HOY(HOS$)
       Call REPLA(TTXYZ$, HOS$, 1, 8)
       Call REPLA(TTXYZ$, "Generación de Orden de Venta:", 11, 56)
       Call REPLA(TTXYZ$, USERTER$, 69, 24)
       ANOTA$ = ANOTA$ + TTXYZ$ + Chr$(10) + Chr$(13)
       IKITEM% = 0
       '
       SQLX$ = "SELECT * FROM  PEDDETA  "
       SQLX$ = SQLX$ + " WHERE NroClie = -1"
       Set rs1 = New ADODB.Recordset
30     On Error Resume Next
       rs1.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
       If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
              On Error GoTo 0
              Call CapturaErrorOpen
              GoTo 30
       End If
       On Error GoTo 0
       
       CMO$ = "RP"
       CONDI$ = "CodiMovi = '" & CMO$ & "'"
       CONDI$ = CONDI$ + " AND DoPriLAR = '" & REMI$ & "' AND CANTSALID > 0 AND CODIMOVI= 'RP'"
       SQLX$ = "SELECT COD_INTE,NuorItem,DESCRI_LAR, CantSalid, cantingre, NUPEDCLI, ValoUnitIva, Nume_Clie, FECHMOV,IDENLOTE FROM MOVISTO WITH(NOLOCK)"
       SQLX$ = SQLX$ + " WHERE " & CONDI$ & "  ORDER BY NuOrItem"
       Set RS = READSET(SQLX$)
       With RS
          Do While Not .EOF
                NUORIT% = XVALO(!NuOrItem)
                CIXI% = XVALO(!cod_inte)
                CANTI = XVALO(!CANTSALID)
                LOTE$ = SAFETEXT$(!idenlote)
                PRECARGA = XVALO(!VALOUNITIVA)
                PREUNI = PRECARGA
                CAC# = 0
                PORDESCUIN = XVALO(VALOBADA("REPUREPA", "PORDESC", "NUMREPA = " & NUREP&, "NOMESS")) ' DESCUENTO QUE SE AGREGA EN PRESUPUESTO
                If CIXI% > 0 And CANTI > 0 Then
                   If PORDESCUIN > 0 Then
                     PREUNI = PREUNI * (100 - PORDESCUIN) / 100
'                     TOT# = CDbl(CAN * PREUNI)
                   End If
'                 COCAL = COSCALC#(CIXI%)
'                 ACUMUCOSCAL = ACUMUCOSCAL + (CANTI * COCAL)
                 With rs1
                     .AddNew
                     !CODIEMPR = CodiEmp%
                     IKITEM% = IKITEM% + 1
                     !NROPED = NROPEDI1&
                     !FECHEMIS = CVDAT(FEMI%)
                     !NroClie = NC_ASOC_TECNICO
                     !RasoClie = rasocli$(NC_ASOC_TECNICO)
                     !NROOCOM = ""
                     !FeSolEnt = CVDAT(CVINT(FEX1$)) ' LO GENERA GUANDO GRABA EL PEDENCA
                     !Status = 0              ' Status  de Aprobación
                     !ListPre = 0
                     !CondPag = 0
                     !VENDED = 0
                     !PorComVend = 0
                     !NuOrItem = IKITEM%
                     !CODIINT = CIXI%
                     !CODIEXT = CODEXT$(CIXI%)
                     !SubCod = ""
                     !Descrip = DESCRIT0$(CIXI%)
                     !TLargo = 0
                     !TAncho = 0
                     !TEspesor = 0
                     !CBase = ""
                     !TTALLE = ""
                     !TCOLOR = ""
                     !VenRep = ""   ' Venta o Reparacion
                     !MoneEmis = 1
                     'PREUNI# = PREUNI
                     !TIPOCAM = 1
                     !PreList = PRECARGA
                     !DescLista = 0
                     !DescCombi = ""
                     !DescRubro = 0
                     !DescTotal = 0
                     !CanAsig = 0
                     !CANPED = CANTI
                     !PreUnid = PREUNI
                     !ImpoTot = PREUNI * CANTI
                     !CANTENT = 0
                     !CanSaldo = 0
                     !FeUltEnt = CVDAT(0)
                     !NroRemi = ""
                     !TPedItem = ""
                     '
                     TTXYZ$ = Space$(76)
                     Call REPLA(TTXYZ$, "- " + CODEXT$(CIXI%), 11, 18)
                     Call REPLA(TTXYZ$, DESCRIT0$(CIXI%), 29, 38)
                     Call REPLA(TTXYZ$, CSTRING$(MKS$(1), 3, 10, 1, 2), 57, 10)
                     ANOTA$ = ANOTA$ + TTXYZ$ + Chr$(10) + Chr$(13)
                    .Update
                   End With
                  End If
                  RS.MoveNext
            Loop
          End With
          '
          On Error Resume Next
          rs1.Close
          Set rs1 = Nothing
          RS.Close
          Set RS = Nothing
          On Error GoTo 0
          
          Call REPLA(TTXX$, "*** Técnico Externo ***", 11, 56)
          ANOTA$ = ANOTA$ + TTXX$ + Chr$(10) + Chr$(13)
          ANOTA$ = ANOTA$ + Space$(10) + String$(56, "-") + Chr$(10) + Chr$(13)
          SQLX$ = "SELECT ANOTAPED FROM PEDENCA WITH(NOLOCK)"
          SQLX$ = SQLX$ + " WHERE NroPed = " & NROPEDI1& & ""
          Set PEDDDENCA = READSET(SQLX$)
          ANOTB$ = SAFETEXT$(PEDDDENCA!AnotaPed)
          ANOTA$ = ANOTB$ + ANOTA$
          ANOTA$ = REPLACAR$(ANOTA$, "'", "")
          
          SQLX$ = "UPDATE PEDENCA SET AnotaPed=" & "'" & ANOTA$ & "' WHERE NroPed=" & NROPEDI1&
          FLEXCONN.Execute (SQLX$)
          '
          OKX% = 1
          '
          RETORNO_PEDIDO = NROPEDI1&
          

98        Screen.MousePointer = 1
   '
End Sub


Function NUPEDIDO&()
   '
   NPP& = 1 + XVALO(VALOBADA("PEDENCA", "MAX(NROPED)", "NROPED >= 0", "NOMESS"))
   '
   NUPEDIDO& = NPP&
   '
End Function


Sub ACTUALIZA_GARANT_UNUREP(NUREP&, NUSERI$)
       'ACTUALIZA EL ESTADO,EL UTLIIMO NRO DE REPARACION Y LA FECHA DE VENCIMIENTO DE GARANTIA DE REPARACION
       SQLX$ = "SELECT FevenGaRep,UnumRep FROM TANUMSE "
       SQLX$ = SQLX$ + " WHERE  NUMEROSERIE = '" & NUSERI$ & "'"

       Set rs1 = New ADODB.Recordset
30     On Error Resume Next
       rs1.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
       If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
              On Error GoTo 0
              Call CapturaErrorOpen
              GoTo 30
       End If
       
       On Error GoTo 0
       If Not rs1.EOF Then
         rs1!FevenGaRep = CVDAT(DIASPOST(HOY(HO$), 180))
         rs1!UnumRep = NUREP&
         rs1.Update
       End If
       rs1.Close
       Set rs1 = Nothing

End Sub
Function MUESTRATECNICO()
    Campos$ = "Cód.Téc./f9;Nombre/A32;CTACLIENTE/A0;R.SOCCTE/A0;CTAPROV/A0;RSOCPROV/A0"
    Call CARGA_ENCABEZADO(FRMZELECHO.MSF2, Campos$, FRMZELECHO)
    Call CARGA_ENCABEZADO(FRMZELECHO.MSF1, Campos$, FRMZELECHO)
    FRMZELECHO.Caption = ""

    Dim obj As New RECORXET
    Set rs1 = obj.RS_LISTA_DE_TÉCNICOS()
    Call Carga_MSF_Recordset(rs1, Campos$, FRMZELECHO.MSF2)
    Set obj = Nothing
    rs1.Close
    Set rs1 = Nothing
    FRMZELECHO.Width = 8500
    FRMZELECHO.Show 1
    MUESTRATECNICO = XVALO(RESP_ZELE_VECT(1))
End Function

Public Function RS_LISTAREPA(CONDI$) As ADODB.Recordset
  'SOLO TRAE LOS QUE TIENEN TECNICO
  Dim RS_FUN As ADODB.Recordset
  Set RS_FUN = New ADODB.Recordset
  SQLX$ = "SELECT R.NUMREPA,R.REFEREN,R.TEC_EXT,D.CAMPO2 FROM REPARA R  WITH(NOLOCK)"
  SQLX$ = SQLX$ + " INNER JOIN DATASQL D ON CONVERT(int, D.campo1)= R.TEC_EXT"
  SQLX$ = SQLX$ + " WHERE D.LISTA = 'TECNIEXT' AND " & CONDI$
  Set RS_FUN = READSET(SQLX$)
  Set RS_LISTAREPA = RS_FUN
  Set RS_FUN = Nothing
End Function
Sub PRESENTA_TABLAS_AUXILIARES()

  Call GENERAR_ESTRUCTURAS_SQL
  Call CARGA_ENCABEZADO(FRMZELECHO.MSF2, "ID LISTA/F12;NOMBRE LISTA/A32;DESCRIPCION/A48", FRMZELECHO)
  Call CARGA_ENCABEZADO(FRMZELECHO.MSF1, "ID LISTA/F12;NOMBRE LISTA/A32;DESCRIPCION/A48", FRMZELECHO)
  FRMZELECHO.Caption = "LISTA DE ESTRUCTURAS PRF Y RFS"

  Dim obj As New RECORXET
  Set rs1 = obj.RS_ENCAPRFLISTAS()
  Call Carga_MSF_Recordset(rs1, "ID LISTA/F12;NOMBRE LISTA/A32;DESCRIPCION/A48", FRMZELECHO.MSF2)
  Set obj = Nothing
  rs1.Close
  Set rs1 = Nothing
  EPAC$ = UCase$(TRIM$(EMPREAC$))
  '

  FRMZELECHO.Width = 8500
  FRMZELECHO.Show 1
  IDLIST& = XVALO(RESP_ZELE_VECT(1))
  If InStr(EPAC$, "ENVAPLAST") > 0 Then
    LIXTA$ = TRIM$(RESP_ZELE_VECT(2))
    If LIXTA$ = "MATERIALES ENVAPLAST" Then
       Call COMUNI("Atención !!! Editar Tabla de Familias Desde Maestro de Artículos")
       Exit Sub
    End If
  End If

  Call ABM_TABLA(IDLIST&)
End Sub

Function RESULTADO_CONTROL_EYM$(TABLA$, CAMPO$, NORFA$, NSERIE$)
  'DEVUELVE EL RESULTADO DEL CONTROL FINAL O REGISTRO DE DATOS
  RESULTADO_CONTROL_EYM$ = VALOBADA(TABLA$, CAMPO$, "IDSUBOR = '" & NORFA$ & "' AND NUMEROSERIE = '" & NSERIE$ & "' ORDER BY REVISION DESC", "NOMESS")
End Function
Function NOMBRE_CONTROL_EYM$(CI1%)
  'DEVUELVE EL NOMBRE DEL FROMULARIO QUE ESTA VINCULADO EL CODIGO
  NOMBRE_CONTROL_EYM$ = VALOBADA("CONTROEYM", "C_FRMCONTROL", "COD_INTE = " & CI1%, "NOMESS")
End Function
Function REFE_TABLA_CONTROL_EYM$(CI1%)
  'DEVUELVE EL NOMBRE DE LA TABLA INVOLUCRADA SEGUN EL CONTROL QUE ESTA VINCULADO EL CODIGO
  REFE_TABLA_CONTROL_EYM$ = VALOBADA("CONTROEYM", "REF_CONTROL", "COD_INTE = " & CI1%)
  'Agregue esta fila POR QUE USA LA MISMA TABLA.
  If TRIM$(UCase$(REFE_TABLA_CONTROL_EYM$)) = "REGCFINC15EYM" Then REFE_TABLA_CONTROL_EYM$ = "CONTFINEYM"
End Function


Function EXCEPCION_CONTROL_EYM$(CI1%, REFE_TABLA_CONTROL_EYM$)
  'DEVUELVE EL NOMBRE DE LA LISTA DE EXCEPCIONES SEGUN LA REFERENCIA Y EL CODIGO
  CONDI$ = "REF_CONTROL = '" & REFE_TABLA_CONTROL_EYM$ & "' AND COD_INTE = " & CI1%
  EXCEPCION_CONTROL_EYM$ = VALOBADA("CONTROEYM", "EXCEPCIONES", "COD_INTE = " & CI1%, "NOMESS")
End Function


Sub GENERAR_RESERVAS_DE_TRAZABLE_EYM(NORFA$, NSERIE$)
           '
           SQLX$ = "SELECT COD_INTE,COD_EXTE,TRAZABLE,NUMEROSERIE,IDENLOTE, SUM(CANTI)AS CANTI "
           SQLX$ = SQLX$ + " FROM TRAZAXSERIE WITH(NOLOCK)  WHERE (TRAZABLE = 1 OR TRAZABLE < 0)"
           SQLX$ = SQLX$ + " AND IDSUBOR = '" & NORFA$ & "' AND NUMEROSERIE = '" & NSERIE$ & "'"
'           SQLX$ = SQLX$ + " AND (FUE_A_RESER < 1 OR FUE_A_RESER  IS NULL)" 'NO PARECE UTIL POR AHORA
           SQLX$ = SQLX$ + " GROUP BY COD_INTE,COD_EXTE,TRAZABLE,NUMEROSERIE,IDENLOTE"
           Set RS = READSET(SQLX$)
           With RS
             Do While Not .EOF
               CI1% = XVALO(!cod_inte)
               CODI$ = SAFETEXT$(!Cod_Exte)
               CANTIX = XVALO(!CANTI)
               LOTE$ = SAFETEXT$(!idenlote)
               Call APLICAR_LOTES_A_RESERVAS(CI1%, CODI$, CANTIX, LOTE$, NORFA$, NSERIE$)
               .MoveNext
             Loop
           End With
           RS.Close: Set RS = Nothing
End Sub

Function ControlItemsSinCargo%()
     ControlItemsSinCargo% = 0
     'If CONTROL("ADMIPED", "ITEMSCGO") <> "SI" Then ControlItemsSinCargo% = 0
     If Control("ADMIPED", "ITEMSCGO") <> "SI" Then GoTo 99
     
     ' VERIFICA SI EL USUARIO TIENE DERECHOS PARA ENTREGA SIN CARGO CUALQUIER ARTICULO
     NumeroUsuario% = USNBR%
     IdTablaUsuariosConDerechos& = ObtenerIdTablaAuxiliar("USUASINCAR")
     CondicionSql$ = "codiempr=" & CStr(CODIEMPR%) & " and id_lista=" & CStr(IdTablaUsuariosConDerechos&) & " and campo1='" & CStr(NumeroUsuario%) & "'"
     PERMISO$ = UCase(TRIM$(SAFETEXT$(VALOBADA("DATASQL", "campo3", CondicionSql$))))
     If PERMISO$ = "S" Or PERMISO$ = "SI" Then
        GoTo 99
     End If
     
     '
     '0= TODO BIEN
     '1= TIENE ITEMS SIN CARGO PERO ESTAN APROBADOS
     '-1 TIENE ITEMS SIN CARGO QUE ESTAN EN LA LISTA Y NECESITAN APROBACION Y SU VALOR ES 0
     '   O BEIN SI SUPERAN EL PORCENTAJE DE TOLERANCIA (SOLO SI TOLERANCIA ES > 0)
     '-2 TIENE ITEMS  QUE NO ESTAN EN LA LISTA Y LA DIFERENCIA ES MAYOR A LA TOLERADA CONTRA LA LISTA
     '-3 TIENE ITEMS CON VALOR 0 QUE NO ESTAN EN LA LISTA (SE CONSIDERA POR SI NO HAY PORCENTAJE DE TOLERANCIA
     PorcentajeTolerancia = ROUND(XVALO(Control("ADMIPED", "TOLELPRE")), 2)
     
     For U& = 1 To ULTREG&("!CARDETPE")
       X$ = REGLEIDO$("!CARDETPE", U&)
       CAN = CVS(Mid$(X$, 47, 4))
       If CAN > 0.05 Then
         '
         CIXI% = CVI(Left$(X$, 2))
         If CIXI% > 0 Then
           NLISST% = CVS(Mid$(X$, 85, 4))
           PLISTA# = PRELISTA#(SAFETEXT$(NLISST%), CIXI%)
           '
           'LISTA_PANTALLA% = XVALO(Text21(0).TEXT)
           MON% = MONEMI%
           TCMB = XVALO(TICAMBIS)
           
           PRLIST# = CVS(Mid$(X$, 39, 4))
           PRUN# = CVS(Mid$(X$, 73, 4))     ' CVS(Mid$(X$, 51, 4))
           IMPTOT# = CVD(Mid$(X$, 55, 8))
           
           EXISTLISTSC% = ExisteListaSinCargo%(CIXI%)
           
           If EXISTLISTSC% = 0 Then 'CASO QUE FIGUREN EN LA LISTA Y NO NECESITEN AUTORIZACION
                ControlItemsSinCargo% = 1
           End If
           
           If EXISTLISTSC% = 1 Then 'CASO QUE FIGUREN EN LA LISTA PERO NECESITEN AUTORIZACION
              If IMPTOT# <= 0.01 Then
                ControlItemsSinCargo% = -1
                Exit For
              End If
           End If
           ' VALIDA TOLERANCIA EN ARTICULOS QUE NO EXISTEN EN LA LISTA
           ' Y EN AQUELLOS QUE ESTAN MARCADO COMO PEND DE APROBACION CUYO PRECIO ES MAYOR A 0.01
           If EXISTLISTSC% = -1 Or EXISTLISTSC% = 1 Then 'CASO QUE NO EXISTA EN LA LISTA Y TENGAN UNA DIFERENCA MAYOR AL PORCENTAJE DE TOLERANCIA
              If PorcentajeTolerancia > 0 Then
                DIF = ToleranciaEnPorcentaje(PLISTA#, PRUN#)
                If DIF > PorcentajeTolerancia Then
                    ControlItemsSinCargo% = -1
                    Exit For
                End If
              Else
                'CASO QUE SI NO TIENE PORCENTAJE DE TOLERANCIA Y NO ESTAN EN LA LISTA DE APROBADOS DIRECTOS
                 If EXISTLISTSC% = -1 And IMPTOT# <= 0.01 Then
                   ControlItemsSinCargo% = -3
                   Exit For
                End If
              End If
           End If
           
         End If
       End If
     Next U&
     
99  Exit Function
     
End Function
Function ExisteListaSinCargo%(CIXI%)
  '0=EXISTE EN LA LISTA Y NO PRECISA APROBACION
  '1 = EXISTE EN LA LISTA Y ESTA PARA PENDIENTE DE APROBACION
  '-1 = NO EXISTE EN LA LISTA
  CODI$ = CODEXT$(CIXI%)
  CONDI$ = "LISTA  = 'ART_SIN_CARGO' AND CAMPO1 = '" & CODI$ & "'"
  SQLX$ = "SELECT CAMPO3 FROM DATASQL WITH(NOLOCK) WHERE " & CONDI$
  Set RS = READSET(SQLX$)
  If Not (RS.BOF And RS.EOF) Then ' SI EXISTE EN LA LISTA DEVUELVE EL VALOR DEL CAMPO3 DONDE SE PONE 1 O VACIO
     ExisteListaSinCargo% = XVALO(RS!CAMPO3)
  Else
     ExisteListaSinCargo% = -1 'SI LA CONSULTA NO ARROJA NADA DEVUELVE -1 (NO EXISTE EL CODIGO EN LA LISTA
  End If
  RS.Close
  Set RS = Nothing
End Function
Function SELECCION_FACTURAPROFORMA%(NC, Optional CAMPOS_FORMAT$, Optional CAMPSEL$, Optional MODO%)
    'PRESENTA LISTA DE PROFORMAS PENDIENTES DEL CLIENTE ACTIVO
    SELECCION_FACTURAPROFORMA% = 0
    If TRIM$(CAMPOS_FORMAT$) = "" Then
       'SI NO ESTA PARAMETRIZADO CAMPOS_FORMAT$ MUESTRA ESTOS CAMPOS POR DEFECTO
       Campos$ = "Fecha/D8;N°Prof./F8;Referencia/A84;Imp.Neto/F12.2;Imp.Total/F12.2;Pedido/F8"
    Else
       Campos$ = CAMPOS_FORMAT$ 'CAMPOS QUE VIENEN COMO PARAMETRO
    End If
    
    Call CARGA_ENCABEZADO(FRMZELECHO.MSF2, Campos$, FRMZELECHO)
    Call CARGA_ENCABEZADO(FRMZELECHO.MSF1, Campos$, FRMZELECHO)
'    FRMZELECHO.Caption = "Proformas Pendientes de Cliente: " & NC
    FRMZELECHO.Caption = "Facturas Proforma"
    
    If TRIM$(CAMPSEL$) = "" Then
       'SI NO ESTA PARAMETRIZADO CAMPSEL$ MUESTRA ESTOS CAMPOS POR DEFECTO
       CAMPOS_SELECT$ = "Fechemisi,NumeComp,REFERENCIA,TotNeto,ImpoTot,NROPED"
    Else
       CAMPOS_SELECT$ = CAMPSEL$ 'CAMPOS QUE VIENEN COMO PARAMETRO
    End If
    
    Dim obj As New RECORXET
    Set rs1 = obj.RS_FACT_PROFORMA_PEN(NC, CAMPOS_SELECT$, MODO%)
    Call Carga_MSF_Recordset(rs1, Campos$, FRMZELECHO.MSF2, 1)
    Set obj = Nothing
    rs1.Close
    Set rs1 = Nothing
    FRMZELECHO.Width = 14800
    FRMZELECHO.Show 1
    SELECCION_FACTURAPROFORMA% = XVALO(RESP_ZELE_VECT(2))
End Function

Function FECHAACRED%(Tipo%, FECHE%, FEVEN%)
    'CONFIGURACION DE TABLA DE DEMORA DE DIAS DE ACREDITACION USADO PARA CASHFLOW Y  LAS TARJETAS DE CREDITO
    FECHAACRED% = FEVEN% 'POR DEFECTO LE AGREGA LA FECHA DE VENCIMIENTO
    CONDI$ = "LISTA = 'CONFIMECO' AND CAST(CAMPO1 AS Integer) = " & Tipo%
    CANDIAS% = XVALO(VALOBADA("DATASQL", "CAMPO3", CONDI$, "NOMESS"))
    If CANDIAS% > 0 Then
       CANDIAS% = CANT_DIASPOSTHABILES(FECHE%, CANDIAS%)
       FECHAACRED% = DIASPOST(FECHE%, CANDIAS%)
       A = FECHATEX$(FECHAACRED%)
    End If
    '
End Function


  Sub MOSTRARTABLASCONFIGURADAS(FORMUX$)
      Call GENERAR_ESTRUCTURAS_SQL
      Call CARGA_ENCABEZADO(FRMZELECHO.MSF2, "ID LISTA/F12;NOMBRE LISTA/A32;DESCRIPCION/A48", FRMZELECHO)
      Call CARGA_ENCABEZADO(FRMZELECHO.MSF1, "ID LISTA/F12;NOMBRE LISTA/A32;DESCRIPCION/A48", FRMZELECHO)
      FRMZELECHO.Caption = "LISTA DE ESTRUCTURAS PRF Y RFS"
      J& = 0
      
      SQLX$ = "SELECT CAMPO1,CAMPO2 FROM  DATASQL  WITH(NOLOCK) WHERE LISTA ='TABDASQL' AND CAMPO3 = '" & FORMUX$ & "'"
      Set RS = READSET(SQLX$)
      With RS
        Do While Not .EOF
          J& = J& + 1
          FRMZELECHO.MSF2.Rows = J& + 1
          FRMZELECHO.MSF2.TextMatrix(J&, 1) = (SAFETEXT$(!CAMPO1))
          FRMZELECHO.MSF2.TextMatrix(J&, 2) = (SAFETEXT$(!campo2))
          DESXRI$ = VALOBADA("ENCAPRFUN", "DESCRIPCION", "LISTA = '" & TRIM$(FRMZELECHO.MSF2.TextMatrix(J&, 2) & "'"))
          FRMZELECHO.MSF2.TextMatrix(J&, 3) = DESXRI$
          .MoveNext
        Loop
      
      End With
      FRMZELECHO.Width = 12500
      FRMZELECHO.Show 1
      IDLIST& = XVALO(RESP_ZELE_VECT(1))
    
      Call ABM_TABLA(IDLIST&)
    
End Sub

Function AsientoTieneCuentaMadreProveedor&(NUASIEN$, CTACBLE%)
               'devuelve EL NRO DE PROVEEDOR si el asiento tiene alugna de las cuentas de contrapartida que es cuentamadre de  algun proveedor.
               AsientoTieneCuentaMadreProveedor& = 0
               SQLX$ = "SELECT D.CUECOINT  FROM DetaAsien D  WITH(NOLOCK) "
               SQLX$ = SQLX$ + " INNER JOIN EncaBeAsien E ON D.NUIDASIEN = E.NUIDASIEN "
               SQLX$ = SQLX$ + " WHERE (E.TIPOASIEN < 1 OR E.TIPOASIEN IS NULL)"
               SQLX$ = SQLX$ + " AND D.NUIDASIEN = '" & NUASIEN$ & "'"
               SQLX$ = SQLX$ + " AND  D.CUECOINT <> " & CTACBLE%
               SQLX$ = SQLX$ + " AND E.STATASIEN >= 0"
               Set RS = READSET(SQLX$)
               With RS
                  Do While Not .EOF
                    CCBLE% = XVALO(!CUECOINT)
                    ESCCBLEMADREPROV% = XVALO(VALOBADA("PROVED12", "NUME_CLI", "CuMa_Cli = " & CCBLE%, "NOMESS"))
                    If ESCCBLEMADREPROV% > 0 Then
                       AsientoTieneCuentaMadreProveedor& = ESCCBLEMADREPROV%
                       Exit Do
                    End If
                    .MoveNext
                  Loop
               End With
               RS.Close
               Set RS = Nothing

End Function
Sub ANULARSALIDANUMEROSERIESEGUNREMITO(REMX$, NC, FEMOV)
            'BLANQUEA EL NUMERO DE REMITO,FECHA, CLIENTE EN TANUMSE DE LOS EQUIPOS QUE SALIERON CON EL REMITO
            'BORRA EN EL PEDDETA EL NRO DE SERIE Y REGISTRA EN EL ANOTADOR

            SQLX$ = "SELECT NumRemito, NumFACTURA, NumeCli, FechEntrega, FeVenGar,DepoNumse,NUMEROSERIE FROM TANUMSE "
            SQLX$ = SQLX$ + " WHERE numremito = '" & REMX$ & "'"
            SQLX$ = SQLX$ + " AND numecli = " & NC
            SQLX$ = SQLX$ + " AND FECHENTREGA = '" & FEMOV & "'"
            
            Dim RSTNSE1 As ADODB.Recordset
            Set RSTNSE1 = New ADODB.Recordset
35          On Error Resume Next
            RSTNSE1.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
            If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
                On Error GoTo 0
                Call CapturaErrorOpen
                GoTo 35
            End If
            On Error GoTo 0
            Dim NSERIEX$(): J% = 0
            With RSTNSE1
              Do While Not .EOF
               !NumRemito = ""
               !NUMFACTURA = ""
               !NUMECLI = 0
               !FECHENTREGA = CVDAT(0)
               !FEVENGAR = CVDAT(0)
               !DEPONUMSE = DEPX%
               NUSERIE$ = SAFETEXT$(!NUMEROSERIE)
               
               J% = J% + 1
               ReDim Preserve NSERIEX$(J%)
               NSERIEX$(J%) = NUSERIE$
               
               .Update
               .MoveNext
              Loop
            End With
            RSTNSE1.Close
            Set RSTNSE1 = Nothing
            '
            '
'            RECORRER EL VECTOR Y REALIZAR LA ANOTACION Y TAMBIEN BORRAR EN EL PEDETA
            On Error Resume Next
            cantseries% = UBound(NSERIEX$)
            On Error GoTo 0
            If cantseries% > 0 Then
              For i% = 1 To cantseries%
                    NUSERRR$ = NSERIEX$(i%)
                    CONDI$ = "NUMEROSERIE = '" & NUSERRR$ & "'"
                    Call ACTUREGISTRO("PEDDETA", "NUMEROSERIE", CONDI$, "", REG&, "NOMESS")
                    'GRABACION EN EL ANOTADOR (HISTORIAL DE EQUIPOS)
                    CUERPO$ = "Reingreso (Anulación de Remito: " & REMX$ & ")"
                    Call ANOTANUSE(NUSERRR$, CUERPO$)
              Next i%
            End If

End Sub


Function LISTASELECGENERAL(NOMBRELISTA$, Optional ANCHO1$, Optional ANCHO2$, Optional MOSTRARSOLOREFE%, Optional CONDILISTAGENERAL$)
   'USAR PARA LISTAR TABLAS DE DATASQL CON DOS CAMPOS
   On Error Resume Next
   RESP_ZELE_VECT(1) = ""
   On Error GoTo 0
   'SI SE QUIERE MODIFICAR EL FORMATO Y EL ANCHO DE ALGUN COLUMNA SE DEBERA PASAR POR EJ.: A24
   'USO CONCRETO SI EL CODIGO NO ES N UMERICO SE PUEDE PASAR A PESOS, SI EL ANCHO DE LA DESCRICPCION ES CORTO
   If MOSTRARSOLOREFE% > 0 Then
      Campos$ = "CODIGO/A0;REFERENCIA/A32"
   Else
      Campos$ = "CODIGO/A8;REFERENCIA/A32"
   End If
   If TRIM$(ANCHO1$) <> "" Then Campos$ = REPLACAR(Campos$, "CODIGO/F8", "CODIGO/" & ANCHO1$)
   If TRIM$(ANCHO2$) <> "" Then Campos$ = REPLACAR(Campos$, "REFERENCIA/A32", "REFERENCIA/" & ANCHO2$)
   
   Call CARGA_ENCABEZADO(FRMZELECHO.MSF2, Campos$, FRMZELECHO)
   Call CARGA_ENCABEZADO(FRMZELECHO.MSF1, Campos$, FRMZELECHO)
   FRMZELECHO.Caption = NOMBRELISTA$
   NOMBRELISTA$ = TRIM$(UCase$(NOMBRELISTA$))
   Dim obj As New RECORXET
   'LA CONDILISTAGENERAL$ VIENE SOLO DE PARAMETRO DE LA LLAMADA NO SE COMPLETA AQUI
   Set RS = obj.RS_LISTASELECGENRAL(NOMBRELISTA$, CONDILISTAGENERAL$)
   Call Carga_MSF_Recordset(RS, Campos$, FRMZELECHO.MSF2, 1)
   Set obj = Nothing
   RS.Close
   Set RS = Nothing

   FRMZELECHO.Width = 8500
   If XVALO(Mid(ANCHO1$, 2)) + XVALO(Mid(ANCHO2$, 2)) > 64 Then
      FRMZELECHO.Width = 11500
   End If
   'PRESENTA EN EL CAPTION  EL NOMBE DE LA LISTA Y LA DESCRIPCION
   CONDI$ = "LISTA = '" & NOMBRELISTA$ & "'"
   KAPTION$ = VALOBADA("ENCAPRFUN", "LISTA  +  ' - ' +  DESCRIPCION", CONDI$, "NOMESS")
   FRMZELECHO.Caption = KAPTION$
   
   FRMZELECHO.Show 1
   If TRIM$(RESP_ZELE_VECT(1)) <> "" Then
     LISTASELECGENERAL = TRIM$(RESP_ZELE_VECT(1))
   Else
     LISTASELECGENERAL = ""
   End If

End Function
Function NumNuevaCotiza()
    NumNuevaCotiza = 1 + XVALO(VALOBADA("COTIZACTP", "TOP 1 NROCOTIZACION", "NROCOTIZACION > 0 ORDER BY NROCOTIZACION  DESC", "NOMESS"))
End Function
Function NumNuevaCotizaPVTA()
    NumNuevaCotizaPVTA = 1 + XVALO(VALOBADA("COTIPVTA", "TOP 1 NROCOTIZACION", "NROCOTIZACION > 0 ORDER BY NROCOTIZACION  DESC", "NOMESS"))
End Function

Function NumNuevaReceta()
    NumNuevaReceta = 1 + XVALO(VALOBADA("RECETACTP", "TOP 1 NUMRECETACTP", "NUMRECETACTP > 0 ORDER BY NUMRECETACTP DESC", "NOMESS"))
End Function

Function FICHATECNCIA_SEGUN_CODIGO&(CI1%)
    CONDI$ = "CODINT = " & CI1%
    FICHATECNCIA_SEGUN_CODIGO& = XVALO(VALOBADA("RECETACTP", "NUMRECETACTP", CONDI$, "NOMESS"))
End Function
Function NUEVONUMEROPRES&()
    NUEVONUMEROPRES& = 1 + XVALO(VALOBADA("PRESUVEN", "TOP 1 NROPRES", "NROPRES > 0 ORDER BY NROPRES  DESC", "NOMESS"))
End Function

Function EstadoAprobacionReparaEym%(NUREP&)
   CONDI$ = "NUMREPA = " & NUREP&
   EstadoAprobacionReparaEym% = XVALO(VALOBADA("REPARA", "STAT_APROBACION", CONDI$, "NOMESS"))
End Function

Function NumNuevaPVta&()
    'NUEVO NUMERO DE PLANILLA DE VENTAS
    NumNuevaPVta& = 1 + XVALO(VALOBADA("P_VTA_BOLSAS", "TOP 1 NROPVTA", "NROPVTA > 0 ORDER BY NROPVTA DESC", "NOMESS"))
End Function

Function CamposdeUnaTabla$(TABLA$, Optional MostrarIdentity%, Optional CEmp$)
    'GENERA UN STRING CON TODOS LOS CAMPOS DE UNA TABLA SEPARADOS POR ,
    CamposdeUnaTabla$ = "" 'si CEmp$ tiene valor reemplaza el campo Codiempr por dicho valor
    CEmp$ = TRIM$(CEmp$)
    cnt = 0
    Set RS = FLEXCONN.Execute("sp_columns [" & TABLA$ & "]")
    Do Until RS.EOF
        namex$ = TRIM$(RS!column_name)
        If MostrarIdentity% = 0 Then
          If UCase(namex$) = TABLA$ & "_ID" Then
           GoTo 30
          End If
        End If
        If CEmp$ <> "" And TRIM$(UCase(namex$)) = "CODIEMPR" Then namex$ = CEmp$
        cnt = cnt + 1
        If cnt = 1 Then
           CamposdeUnaTabla$ = CamposdeUnaTabla$ & namex$
        Else
           CamposdeUnaTabla$ = CamposdeUnaTabla$ & "," & namex$
        End If
30      RS.MoveNext
     Loop
     Set RS = Nothing
End Function

Sub DuplicarRegistro(TABLA$, CONDI$, Optional Ret_NuevoIDGenerado&)
    'SI ObtenerID& > 0 OBTIENE EL ID DEL REGISTRO CREADO PERO LA TALBA DE POSEER UN CAMPO IDENTITY
    'DUPLICA UN REGISTRO SEGUN UNA CONDICION USAR PREFERENTEMENTE PARA TABLAS CON CAMPO UNIVOCO
    '  OJO SI LA CONDICION NO ES PRECISA PUEDE DUPICAR MAS DE UN REGISTRO ***
    If TRIM$(CONDI$) = "" Then Exit Sub
    If TRIM$(TABLA$) = "" Then Exit Sub
    If EXISTE_TABLA(TABLA$) < 1 Then Exit Sub
    Campos$ = CamposdeUnaTabla$(TABLA$)
'    On Error GoTo ERROR_GUARDAR
'
'    FLEXCONN.BeginTrans
    SQLX$ = "INSERT INTO " & TABLA$
    SQLX$ = SQLX$ + "(" & Campos$ & ")"
    SQLX$ = SQLX$ + " SELECT " & Campos$
    SQLX$ = SQLX$ + " FROM " & TABLA$
    SQLX$ = SQLX$ + " Where " & CONDI$
    FLEXCONN.Execute SQLX$
    If Ret_NuevoIDGenerado& > 0 Then
        SQLX$ = "SELECT @@IDENTITY AS NUEVOID FROM " & TABLA$ & " WITH(NOLOCK)"
        Set RS = READSET(SQLX$)
        Ret_NuevoIDGenerado& = RS!NUEVOID
    End If
    
'ERROR_GUARDAR:
'    If Err.Number <> 0 Then
'       FLEXCONN.RollbackTrans
'       Call COMUNI("ERROR AL GUARDAR NO SE REALIZO LA TRANSACCION" & vbCrLf & Err.Number & " " & Err.Description)
'       On Error GoTo 0
'       Exit Sub
'    Else
'
'      FLEXCONN.CommitTrans
'    End If
    
End Sub
Function MuestraAsignacionCamposdeUnaTabla$(TABLA$, Optional MostrarIdentity%, Optional CEmp$)
    'GENERA UN STRING CON TODOS LOS CAMPOS DE UNA TABLA SEPARADOS POR ,
'    CamposdeUnaTabla$ = "" 'si CEmp$ tiene valor reemplaza el campo Codiempr por dicho valor
    Asig$ = ""
    Campostodos$ = CamposdeUnaTabla$(TABLA$)
    Dim strArray() As String
    On Error Resume Next
    strArray = Split(Campostodos$, ",")
    For i& = 0 To UBound(strArray)
       Asig$ = Asig$ + "!" & strArray(i&) & " = " & "RS0!" & strArray(i&) & vbCrLf
    Next i&
    MuestraAsignacionCamposdeUnaTabla$ = Asig$
End Function
Sub DuplicarRegistroDeUnaBaseAOtra(BaseOrigen$, TablaOrigen$, BaseDestino$, TablaDestino$, CONDI$, Optional NOCerrarConexion%, Optional CEmp%, Optional NOREALIZARDELETE%)
    'GENERA UNA TRANSACCION DE UNA TABLA DE UNA BASE A OTRA
    'LA IDEA ES CUANDO EN UNA TABLA ORIGEN SE GENERA UNA TRANSACCION LA DUPLIQUE EN OTRA TABLA IGUAL DE OTRA BASE
    'PRIMERO BORRA EL REGISTRO DE LA TABLA DESTINO Y LUEGO HACE UN INSERT EN LA MISMA BAJO MISMA CONDICION
    
    Origen$ = BaseOrigen$ & ".dbo." & TablaOrigen$
    DESTINO$ = BaseDestino$ & ".dbo." & TablaDestino$
    '
'    If CantRegistros(DESTINO$, Condi$) < 1 Then Exit Sub
    
'    On Error GoTo ERROR_GUARDAR
''    FLEXCONN.Close
''    Call ABRECONEXION
'    FLEXCONN.BeginTrans
    'PRIMERO BORRA EL REGISTRO DE LA TABLA DESTINO Y LUEGO HACE UN INSERT EN LA MISMA BAJO MISMA CONDICION
    If NOREALIZARDELETE% < 1 Then 'USAR EN EL CASO QUE SE NECESITE DUPLICAR DE UNA MISMA TABLA
        SQLX$ = "DELETE FROM " & DESTINO$ & " WHERE " & CONDI$
        FLEXCONN.Execute (SQLX$)
    End If
    '
    'INSERTA EL/LOS REGISTRO EN LA TABLA DESTINO DE LA TABLA ORIGEN SEGUN CONDICION
    SQLX$ = "INSERT INTO " & DESTINO$ & " (" & CamposdeUnaTabla$(TablaDestino$) & ")"
    SQLX$ = SQLX$ + " SELECT " & CamposdeUnaTabla$(TablaOrigen$, 0, SAFETEXT$(CEmp%))
    SQLX$ = SQLX$ + " From " & Origen$ & " Where " & CONDI$
    FLEXCONN.Execute (SQLX$)
    
'    CANT111 = CantRegistros("" & Origen$ & "", Condi$) '
'    MsgBox "ahp " & CANT111
'    cant222 = CantRegistros("" & DESTINO$ & "", Condi$) '
'    MsgBox "MINEPARTS " & cant222
    
'ERROR_GUARDAR:
'    If Err.Number <> 0 Then
'       XX = Err.Description
'       FLEXCONN.RollbackTrans
'       Call COMUNI("ERROR AL REALIZAR TRANSACCIÓN DE SINCRONIZACION" & vbCrLf & Err.Number & " " & Err.Description)
'       On Error GoTo 0
'       Exit Sub
'    Else
'      FLEXCONN.CommitTrans
'      'Call COMUNI("Proceso Realizado Exitosamente")
'    End If
    'puede o no refrescar la conexion, por que si ingresa a esta rutina a travez de un loop de un recordset
    'cuando vuelve al recordset el hecho de haber cerrado la conexion expira el recordset
    If NOCerrarConexion% < 1 Then
      FLEXCONN.Close
      Call ABRECONEXION
    End If
End Sub

Function ScriptDeInsert$(BaseOrigen$, TablaOrigen$, BaseDestino$, TablaDestino$, CONDI$, Optional NOCerrarConexion%, Optional CEmp%, Optional NOREALIZARDELETE%)
    'GENERA la instruccion del insert para ejecutar O ENVIAR AL DROP BOX Y QUE SE EJECUTE EN OTRO SERVIDOR

    ScriptDeInsert$ = ""
    Origen$ = BaseOrigen$ & ".dbo." & TablaOrigen$
    DESTINO$ = BaseDestino$ & ".dbo." & TablaDestino$
    '
    ListaCampos$ = CamposdeUnaTabla$(TablaDestino$)
    ListaValores$ = ListaValoresInsert$(TablaDestino$, ListaCampos$, CONDI$)
    
    'INSERTA EL/LOS REGISTRO EN LA TABLA DESTINO DE LA TABLA ORIGEN SEGUN CONDICION
    SQLX$ = "INSERT INTO " & DESTINO$ & " (" & ListaCampos$ & ")"
    SQLX$ = SQLX$ + " VALUES  " & ListaValores$
    
    ScriptDeInsert$ = SQLX$
    
End Function

Function EmpresaNoRealizaCambios%()
   EmpresaNoRealizaCambios% = 0
   EPAC$ = TRIM$(UCase$(EMPREAC$))
   If InStr(EPAC$, "MINEPARTS") > 0 Then
      Call COMUNI("No se Admiten Realizar Cambios Desde Empresa Mineparts\Ingresar a Empresa AHP Para Realizar los Mismos")
      EmpresaNoRealizaCambios% = 1
   End If
End Function
Sub DUPLICARCOMPRAMINEPARTS(NrodeOrdenDeCompraGenerado&, Nroprov%)
    'DUPLICA LA COMPRA TAL CUAL SE GRABO EN MINEPARTS
    'Genero la duplicacion en Mineparts tal cual la compra realizada en AHP
    
    '*******   ATENCION SOLO SE GENERARAN LAS ORDENES DE COMPRA DE AHP NO DE MINEPART  *******
    '*******   SI DEBEN GENERAR SOLO PARA MINPARTS QUE LA GENEREN DESDE AHP Y EN TODO  *******
    '*******   CASO QUE LA DEN POR CUMPLIDA                                            *******
    
    CONDI$ = "NUME_OCOM = " & NrodeOrdenDeCompraGenerado&
    Call DuplicarRegistroDeUnaBaseAOtra("AHP", "OCOMENCA", "MINEPARTS", "OCOMENCA", CONDI$, 0, 1)
    Call DuplicarRegistroDeUnaBaseAOtra("AHP", "DETAIVA", "MINEPARTS", "DETAIVA", CONDI$, 0, 1)
    Call DuplicarRegistroDeUnaBaseAOtra("AHP", "OCOMDETA", "MINEPARTS", "OCOMDETA", CONDI$, 0, 1) 'ver si funciona con mas de un registro
    '
    'ACTUALIZO OCOMDETA DEMINEPARTS
    FE% = HOY(HO$): FEC$ = FECHATEX$(FE%)
    RASO$ = rasocli$((-1) * (Nroprov% + 1)) ' para que refresque
    RASO$ = rasocli$((-1) * Nroprov%)
    StringRef$ = FEC$ & " - " & "RASO$"
    CONDI$ = "NUME_OCOM = " & NrodeOrdenDeCompraGenerado& ' el nro de orden de compra creado en ahp
    SQLX$ = "UPDATE Mineparts.dbo.OCOMDETA SET Npro_Ocom = " & Nroprov%
    SQLX$ = SQLX$ + " , RaSo_Prov = '" & RASO$ & "'"
    SQLX$ = SQLX$ + " WHERE " & CONDI$
    FLEXCONN.Execute (SQLX$)

    'Actualizo orden de compra generada orignalmente en Ahp al prov. del exterior pasandola a Mineparts local
    'leo el nro de proveedor Mineparts
    NprovMineparts = XVALO(Control("CFGMINEP", "PRFEMINE"))
    IdGeneradoEnMineparts = XVALO(VALOBADA("Mineparts.Dbo.OCOMENCA", "OCOMENCA_ID", CONDI$, "Nomess"))
    '
    
    FLEXCONN.Close
    Call ABRECONEXION
    '
    'ACTUALIZO LA ORDEN DE COMPRA DE AHP
    FE% = HOY(HO$): FEC$ = FECHATEX$(FE%)
    StringRef$ = FEC$ & " - " & "Mineparts"
    CONDI$ = "NUME_OCOM = " & NrodeOrdenDeCompraGenerado& ' el nro de orden de compra creado en ahp
    SQLX$ = "UPDATE AHP.dbo.OCOMENCA SET Refe_Ocom = '" & StringRef$ & "'"
    SQLX$ = SQLX$ + " , Npro_Ocom = " & NprovMineparts
    SQLX$ = SQLX$ + " , RaSo_Prov = 'Mineparts'"
    SQLX$ = SQLX$ + " , Via_Entrega = '1'"
    SQLX$ = SQLX$ + " , IdCompExt = " & IdGeneradoEnMineparts
    SQLX$ = SQLX$ + " WHERE " & CONDI$
    FLEXCONN.Execute (SQLX$)
    
     'ACTUALIZO OCOMDETA DE AHP
'    FE% = HOY(HO$): fec$ = FECHATEX$(FE%)
'    RASO$ = rasocli$((-1) * Nroprov%)
'    StringRef$ = fec$ & " - " & "RASO$"
    CONDI$ = "NUME_OCOM = " & NrodeOrdenDeCompraGenerado& ' el nro de orden de compra creado en ahp
    SQLX$ = "UPDATE AHP.dbo.OCOMDETA SET Npro_Ocom = " & NprovMineparts
    SQLX$ = SQLX$ + " , RaSo_Prov = 'Mineparts'"
    SQLX$ = SQLX$ + " WHERE " & CONDI$
    FLEXCONN.Execute (SQLX$)

    'Fin actualizacion Ocomenca
End Sub
    
    
Function StringAggFlex$(TABLA$, CONDI$, CAMPO$, SEPARADOR$)
'STRING_AGG
    StringAggFlex$ = ""
    TCAMPO = TIPODECAMPO(TABLA$, CAMPO$)
    If TCAMPO = 10 Or TCAMPO = 12 Or TCAMPO = 231 Then    ' SI ES CAMPO DE TEXTO LE AGRGEGA COMILLAS SIMPLES
      COMILLASSIMPLES$ = Chr$(39)
    End If
    Dim RS As New ADODB.Recordset
    SQLX$ = "SELECT " & CAMPO$
    SQLX$ = SQLX$ + " FROM " & TABLA$ & "  WITH(NOLOCK)"
    SQLX$ = SQLX$ + " WHERE " & CONDI$
    Set RS = READSET(SQLX$)
    With RS
        ' Recorrer el recordset y concatenar los valores
        Do Until RS.EOF
            StringAggFlex$ = StringAggFlex$ & COMILLASSIMPLES$ & TRIM$(.Fields(0)) & COMILLASSIMPLES$ & SEPARADOR$
            RS.MoveNext
        Loop

    End With

    
    ' Eliminar la última coma y espacio agregados
    If Len(StringAggFlex$) > 2 Then
        StringAggFlex$ = Left(StringAggFlex$, Len(StringAggFlex$) - 1)
    End If
    
    ' Cerrar el recordset y la conexión a la base de datos
    RS.Close
    Set RS = Nothing
    

End Function

    

Function TipoObjectSql$(nombre$)
    SQLX$ = "SELECT   name AS ObjectName, type_desc As ObjectType"
    SQLX$ = SQLX$ + " From  sys.objects WITH(NOLOCK)"
    SQLX$ = SQLX$ + " Where name = '" & nombre$ & "'"
    Set RS = READSET(SQLX$)
    With RS
       TIPOOBJETO$ = UCase$(TRIM$(SAFETEXT$(!ObjectType)))
       If TIPOOBJETO$ = "USER_TABLE" Then
           TipoObjectSql$ = "TABLA"
       ElseIf TIPOOBJETO$ = "SQL_STORED_PROCEDURE" Then
           TipoObjectSql$ = "STORED"
       ElseIf TIPOOBJETO$ = "VIEW" Then
           TipoObjectSql$ = "VISTA"
       Else
           TipoObjectSql$ = "DESCONOCIDO"
       End If
    End With
    RS.Close
    Set RS = Nothing
End Function

