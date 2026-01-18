Attribute VB_Name = "SQL_MGR"
Public wrkPredeterminado As Workspace, BASDAT, CREABA, CREATA, CREACA

Public MODICONSU%
Public IMPOSEJEC%
'
Public DESQUERY$
Public TIPQUERY%
Public DATBASE$
Public TABSELE$, TABSECU$, CAMPOMA1$, CAMPOMA2$
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

Sub CARGA_CONSULTA(CODQUERY$)
    '
    Call ABRETACON
    STRSQL = "SELECT * FROM SQL_QUERIES"
    STRSQL = STRSQL + " WHERE COD_QUERY = '" & CODQUERY$ & "'"
    On Error Resume Next
    'Set MIRS = FLEXCONN.Execute(STRSQL)
    Set MIRS = BaDaConsul.OpenRecordset(STRSQL, dbOpenSnapshot)
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
    CANFI% = 0: CACONDI% = 0
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
    TABSECU$ = "": CAMPOMA1$ = "": CAMPOMA2$ = ""
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
                !CodiEmpr = CodiEmp%               ' codigo de empresa
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

Sub ARMA_CONSULTA(Optional vueltanumero%)
    '
    KLIK% = 0
    SQLX$ = "SELECT ": SQLZ$ = "": SQLT$ = ""
    VUELTA& = 0: VUELTB& = 0: VUELTC& = 0: AGRUPA% = 0
    For I& = 1 To CANFI%
       CAMPO$ = SELCAMPO$(I&)
       If Left$(UCase$(CAMPO$), 8) = "VALOBADA" Then
          Call CREACAMPO("", TABSELE$, CAMPO$, 10, 256)
          CAMPO$ = TABSELE$ + "." + CAMPO$
       End If
       ALIASCAMPO$ = CAMPO$
       If GRUCAMPO$(I&) = "1" Then
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
       ORDEX$ = TRIM$(UCase$(ORDCAMPO$(I&)))
       If ORDEX$ = "ASC" Or ORDEX$ = "DESC" Then
          If VUELTB& > O Then SQLZ$ = SQLZ$ + ", "
          SQLZ$ = SQLZ$ + ALIASCAMPO$ + " " + ORDEX$
          VUELTB& = VUELTB& + 1
       End If
       '
       AGRUX$ = GRUCAMPO$(I&)
       If AGRUX$ <> "1" Then
          If VUELTC& > O Then SQLT$ = SQLT$ + ", "
          SQLT$ = SQLT$ + CAMPO$ + " "
          VUELTC& = VUELTC& + 1
       End If
    Next I&
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
             SQLY$ = SQLY$ + " INNER JOIN " + TASEC$
             SQLY$ = SQLY$ + " WITH(NOLOCK) ON " + TAPRIN$ + "." + CAMPRIN$
             SQLY$ = SQLY$ + " = " + TASEC$ + "." + CAMSEC$
          End If
       End If
    End If
    '
    If TASEC$ = "" And CAMPOMA2$ = "" Then SQLY$ = SQLY$ + " WITH(NOLOCK) "
    '
    If CACONDI% > 0 Then
       JJ& = 0
       For I& = 1 To CACONDI%
          If JJ& < 1 Then
                SQLY$ = SQLY$ + " WHERE "
             Else
                SQLY$ = SQLY$ + " AND "
          End If
          SQLY$ = SQLY$ + CAMPOCOND$(I&) + " "
          SQLY$ = SQLY$ + OPERACOND$(I&) + " "
          SQLY$ = SQLY$ + VAREFCOND$(I&)
          JJ& = JJ& + 1
       Next I&
    End If
    '
    TXT_CONSULTA$ = SQLX$ + SQLY$
       If AGRUPA% = 1 Then TXT_CONSULTA = TXT_CONSULTA$ + " GROUP BY " + SQLT$
       If SQLZ$ <> "" Then TXT_CONSULTA = TXT_CONSULTA$ + " ORDER BY " + SQLZ$
    TXT_CAREGIST$ = "SELECT COUNT(*) AS CANREGIS" + SQLY$
    '
' MsgBox TXT_CONSULTA
End Sub
'
Sub CARGADEBASE(ADAT$, TBXX$, OKX%)
    '
    TBXX$ = "": OKX% = 0
    ADAT1$ = TRIM$(ADAT$)
    If ADAT1$ = "" Then Exit Sub
    '
    Call ABRETACON
    STRSQL = "SELECT * FROM SQL_QUERIES"
    STRSQL = STRSQL + " WHERE COD_QUERY = '" & ADAT1$ & "'"
    On Error Resume Next
    Set MIRS = BaDaConsul.OpenRecordset(STRSQL, dbOpenSnapshot)
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
    Call DECOD_CONSULTA(TTXX$)
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
    Set RS = New ADODB.Recordset
    On Error Resume Next
    Set RS = FLEXCONN.Execute(FILTSQL$(TXT_CONSULTA))
    If Err Then Exit Sub
    '
    ' GENERA EL TEXTO
    AX1% = 0: AX2% = 0
    For I& = 1 To RS.Fields.Count
      If FORCAMPO$(I&) = "" Then
         FORCAMPO$(I&) = "A16": ACAMPO% = 16
         TICAMPO% = RS.Fields(I& - 1).Type
         If TICAMPO% = 2 Then FORCAMPO$(I&) = "F7.0": ACAMPO% = 7                       ' INTEGER
         If TICAMPO% = 3 Then FORCAMPO$(I&) = "F7.0": ACAMPO% = 7                       ' INTEGER
         If TICAMPO% = 5 Then FORCAMPO$(I&) = "F12.2": ACAMPO% = 12                      ' NUMERICO
         If TICAMPO% = 202 Then FORCAMPO$(I&) = "A48": ACAMPO% = 48                        ' TEXTO
         If TICAMPO% = 135 Then FORCAMPO$(I&) = "D8": ACAMPO% = 8     ' FECHA
      End If
      ACAMPO% = XVALO(Mid$(FORCAMPO$(I&), 2))
      '
      If I& = 1 Then
           AX1% = ACAMPO%
         Else
           AX2% = AX2% + ACAMPO%
      End If
    Next I&
    '
    AXX% = AX1% + 2 + AX2%
    TBXX$ = Space$(FIN& * AXX%)
    '
    RWW& = 0
    With RS
       Do While Not .EOF
          For I& = 1 To .Fields.Count
            On Error Resume Next
            VALOR = .Fields(I& - 1).Value
            If I& = 1 Then
                 TXTX$ = FORMATNUM$(VALOR, FORCAMPO$(I&)) + "  "
               Else
                 TXTX$ = TXTX$ + FORMATNUM$(VALOR, FORCAMPO$(I&))
            End If
          Next I&
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

Function VALOBADA(NOTAB$, NOCAM$, CONDI$, Optional MODOMESS$)
   'PASANDOLE COMO ARGUMENTO EL NOMBRE DEL CAMPO , EL DE LA TABLA Y LA CONDICION _
   DEVUELVE EN VALOBADA EL VALOR DE LA ENCONTRADO SEGUN LOS ARGUMENTOS ENVIADOS

   Call ABRECONEXION                ' PARA EVITAR QUE SE CAIGA POR ESTAR CERRADA - 9/7/2010
   '
   condi3$ = CONDI$
   PPXX% = InStr(CONDI$, "['")
   If PPXX% > 0 Then
      PPYY% = InStr(PPXX%, CONDI$, "']")
      If PPYY% > PPXX% Then
        CODE$ = TRIM$(Mid$(CONDI$, PPXX% + 2, PPYY% - PPXX% - 2))
        CIII& = CODINT(CODE$)
        condi3$ = Left$(CONDI$, PPXX% - 1) & CIII& & Mid$(CONDI$, PPYY% + 2)
      End If
   End If
   '
   ' SI NO SE DETALLA EL CAMPO O LA CONDICION SALE SIN MENSAJE. EPB 6/10/2011
   If NOCAM$ = "" Or condi3$ = "" Then Exit Function
   '
   SQLX$ = "SELECT " + NOCAM$ + " AS VALORSE "
   SQLX$ = SQLX$ + " FROM " + NOTAB$
   If InStr(UCase$(NOTAB$), "INNER JOIN") < 1 Then
      If InStr(UCase$(NOTAB$), "WITH(NOLOCK)") < 1 Then
         SQLX$ = SQLX$ + " WITH(NOLOCK)" ' AGREGADO EL WITH(NOLOCK) - 9-7-2010
      End If
   End If
   '
   SQLX$ = SQLX$ + " WHERE " + condi3$
   Dim MIRS As ADODB.Recordset
   Set MIRS = New ADODB.Recordset
   '
   VALOBADA = ""
   On Error Resume Next
   Set MIRS = FLEXCONN.Execute(FILTSQL$(SQLX$, MODOMESS$))
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

Function CantRegistros&(NOTAB$, Optional CONDI$, Optional MODOMESS$, Optional TIMUT)
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
Sub ACTUREGISTRO2(TABLA$, CAMPO$, CONDI$, NUEVALOR, REGISTROSAFECTADOS&, Optional MODO$, Optional TIMUT)
   '
'   If XVALO(TIMUT) > 0 Then
'      TIMUTAN = FLEXCONN.CommandTimeout
'      FLEXCONN.CommandTimeout = TIMUT
'   End If
   '
   Dim CAMP_CONSULTA1(), VAL_CONSULTA1(), CAMP_CONSULTA2(), VAL_CONSULTA2()
   Dim CAMPOX As Field
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
              CAMPO2$ = Mid$(CAMP$, 1, POS1% - 1)
              pos2% = InStr(NUEVALOR, ",")
              VALOR2 = Mid$(NUEVALOR, 1, pos2% - 1)
              VALORACT1 = Mid$(NUEVALOR, 1, pos2% - 1)
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
               CAMP_CONSULTA2(U&) = CAMPO2$
               VAL_CONSULTA2(U&) = VALORACT1
               CAMPO2$ = " " & CAMPO2$ & "  AS NUCAM" & U& & " , "
               CAMP$ = Mid$(CAMP$, POS1% + 1)
               ACUMUCAMP2$ = ACUMUCAMP2$ + CAMPO2$
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
            For I& = 1 To UBound(CAMP_CONSULTA1)
               CAMPOX = CAMP_CONSULTA1(I&)
               !CAMPOX = VAL_CONSULTA1
            Next I&
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
            JJ& = 0
            With MOVISE
               Do While Not .EOF
                  JJ& = JJ& + 1: Call TRACE(20, JJ&, NUMREG&)
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
    STRSQL = "Select systypes.name, systypes.xtype from sysobjects inner join syscolumns on "
    STRSQL = STRSQL & "sysobjects.id = syscolumns.id inner join systypes "
    STRSQL = STRSQL & "on syscolumns.xtype = systypes.xtype "
    STRSQL = STRSQL & "where sysobjects.name = '" & NTABLA$ & "' "
    STRSQL = STRSQL & "and syscolumns.name ='" & NCAMPO$ & "'"
    Set MIRS = FLEXCONN.Execute(FILTSQL$(STRSQL))
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
    STRSQL = "Select systypes.name, systypes.xtype from sysobjects inner join syscolumns on "
    STRSQL = STRSQL & "sysobjects.id = syscolumns.id inner join systypes "
    STRSQL = STRSQL & "on syscolumns.xtype = systypes.xtype "
    STRSQL = STRSQL & "where sysobjects.name = '" & NTABLA$ & "' "
    STRSQL = STRSQL & "and syscolumns.name ='" & NCAMPO$ & "'"
    Set MIRS = FLEXCONN.Execute(FILTSQL$(STRSQL))
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
   STRSQL = "SELECT SYSCOLUMNS.LENGTH AS LENG FROM syscolumns "
   STRSQL = STRSQL + "INNER JOIN sysobjects "
   STRSQL = STRSQL & "ON sysobjects.id = syscolumns.id "
   STRSQL = STRSQL & "where sysobjects.name = '" & TABLA$ & "' "
   STRSQL = STRSQL & "AND syscolumns.name = '" & UCase$(CAMPO$) & "'"
   Set MIRS = FLEXCONN.Execute(FILTSQL$(STRSQL))
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
   STRSQL = "Select * from sysindexes WHERE id = OBJECT_ID('" & TABLA$ & "') AND NAME = 'I" + CAMPO$ + "' "
   CONIN% = 0
   '
   On Error Resume Next
   Set MIRS = FLEXCONN.Execute(STRSQL)
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
    Call ABRETACON
    SQLX$ = "SELECT * FROM USER_PARMS "
    SQLX$ = SQLX$ + " WHERE USUARIO = " & USNU%
    SQLX$ = SQLX$ + " AND PROCESO = '" & PROBUS$ & "'"
    SQLX$ = SQLX$ + " AND CLAVE = '" & CODBUS$ & "'"
    Set PARATEMP = BaDaConsul.OpenRecordset(SQLX$, dbOpenDynaset)
    '
    With PARATEMP
       If Not (.BOF And .EOF) Then
           .Edit
         Else
           .AddNew   ' YA CONVERTIDO
           On Error Resume Next
           !CodiEmpr = CodiEmp%
           On Error GoTo 0
       End If
       !USUARIO = USNU%
       !PROCESO = PROBUS$
       !CLAVE = CODBUS$
       !Valocon = VALCON$
       On Error Resume Next
       !CodiEmpr = CodiEmp%
       On Error GoTo 0
       .Update
    End With
    '
90  PARATEMP.Close
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
    Call ABRETACON
    SQLX$ = "SELECT * FROM USER_PARMS "
    SQLX$ = SQLX$ + " WHERE USUARIO = " & USNU%
    SQLX$ = SQLX$ + " AND PROCESO = '" & PROBUS$ & "'"
    SQLX$ = SQLX$ + " AND CLAVE = '" & CODBUS$ & "'"
    Set PARATEMP = BaDaConsul.OpenRecordset(SQLX$, dbOpenSnapshot)
    '
    With PARATEMP
       If Not (.BOF And .EOF) Then
          VCXXX$ = !Valocon
       End If
    End With
    PARATEMP.Close
    '
    VALCONTROL$ = VCXXX$
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
    JJ& = 0
    With ENUME
       If Not (.BOF And .EOF) Then
          Do While Not .EOF
            JJ& = JJ& + 1
            !NUMERORDEN = JJ&
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
    Dim VALOR(), CANREG&
    CANREG& = CantRegistros&(TABLA$, CONDICION$)
    ReDim VALOR(CANREG&)
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
    JJ& = 0
    VALOMAX = 0
    With ENUME
       If Not (.BOF And .EOF) Then
          Do While Not .EOF
            JJ& = JJ& + 1
            NUMACT = XVALO(!NUMERORDEN)
            '
            If NUMACT < 1 Then
               NUMACT = VALOMAX + 1
               !NUMERORDEN = NUMACT
               .Update
               If NUMACT > VALOMAX Then VALOMAX = NUMACT
            End If
            '
            For KK& = 1 To JJ&
               If VALOR(KK&) = NUMACT Then
                  NUMACT = VALOMAX + 1
                  !NUMERORDEN = NUMACT
                  .Update
                  If NUMACT > VALOMAX Then VALOMAX = NUMACT
               End If
            Next KK&
            VALOR(JJ&) = NUMACT
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
       j% = 0: FILA& = 0
       Do While Not .EOF
           TTXX$ = "": COLUI% = 0
           '
           If TIPODELISTA = "GRID" Then
              FILA& = FILA& + 1
37            If FILA& > NOMBREDELISTA.Rows - 1 Then
                 NOMBREDELISTA.Rows = NOMBREDELISTA.Rows + 1
                 GoTo 37
              End If
           End If
           '
           For U& = 1 To .Fields.Count
              VALOR = .Fields(U& - 1).Value
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
              TTYY$ = FORMATNUM$(VALOR, FORMATX$)
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
                   If FILA& = 1 Then
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
                   If Left$(FORMATX$, 1) = "A" Then TTYY$ = VALOR       ' PARA QUE EN LA GRILLA LO PONGA COMPLETO
                   NOMBREDELISTA.TextMatrix(FILA&, COLUI% - 1) = TTYY$
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
                   VALOECO = RASOCLI$(XVALO(VALOR))
                 ElseIf Left$(TECO$, 5) = "PROVE" Or Left$(TECO$, 5) = "ACRED" Then
                   VALOECO = RASOCLI$((-1) * XVALO(VALOR))
                 ElseIf Left$(TECO$, 5) = "MASTE" Then
                   TXVALO$ = VALOR
                     If Len(TXVALO$) > 4 Then
                          CIXI% = CODINT%(TXVALO$)
                        Else
                          CIXI% = XVALO(VALOR)
                     End If
                   VALOECO = DESCRIT0$(CIXI%)
                 ElseIf EXISTE%(LUDAT$ + TECO$ + ".RFS") Then
                   LFF% = LENFIELD%(TECO$, 1)
                   ABUS$ = MKS$(XVALO(VALOR))
                   On Error Resume Next
                   If LFF% = 2 Then ABUS$ = MKI$(Int(XVALO(VALOR)))
                   If LFF% = 1 Then ABUS$ = Chr$(Int(XVALO(VALOR)))
                   If LFF% > 4 Then ABUS$ = TRIM$(VALOR) 'ESTO LO AGREGUE POR QUE SI ABUS$ ES TEXTO NO LO CONSIDERA
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
                      If FILA& = 1 Then
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
                      NOMBREDELISTA.TextMatrix(FILA&, COLUI% - 1) = TTYY$
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
End Sub

Sub CARGALISEL(NOMBREDELISTA$, TABLA$, LISTADECAMPOS$, CONDICION$, Optional MODOSEL$)
    '
    'ejemplo de eco: Num_Prov/F9+PROVED12/A48
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
       j% = 0: KKU& = 0
       Do While Not .EOF
           TTXX$ = ""
           KKU& = KKU& + 1
           Call TRACE(20, KKU&, CANREGI&)
           For U& = 1 To .Fields.Count
              VALOR = .Fields(U& - 1).Value
              FORMATX$ = FORMX$(U&)
              TECO$ = ""
              PPXX% = InStr(FORMATX$, "+")
              If PPXX% > 0 Then
                 TECO$ = UCase$(TRIM$(Mid$(FORMATX$, PPXX% + 1)))
                 FORMATX$ = TRIM$(Left$(FORMATX$, PPXX% - 1))
              End If
              TTYY$ = FORMATNUM$(VALOR, FORMATX$)
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
                   VALOECO = RASOCLI$(XVALO(VALOR))
                 ElseIf Left$(TECO$, 5) = "PROVE" Or Left$(TECO$, 5) = "ACRED" Then
                   VALOECO = RASOCLI$((-1) * XVALO(VALOR))
                 ElseIf Left$(TECO$, 5) = "MASTE" Then
                   TXVALO$ = VALOR
                     If Len(TXVALO$) > 4 Then
                          CIXI% = CODINT%(TXVALO$)
                        Else
                          CIXI% = XVALO(VALOR)
                     End If
                   VALOECO = DESCRIT0$(CIXI%)
                 ElseIf EXISTE%(LUDAT$ + TECO$ + ".RFS") Then
                   ABUS$ = VALOR
                   If TIPFIELD%(TECO$, 1) <> 6 Then
                        LFF% = LENFIELD%(TECO$, 1)
                        ABUS$ = MKS$(XVALO(VALOR))
                        On Error Resume Next
                        If LFF% = 2 Then ABUS$ = MKI$(Int(XVALO(VALOR)))
                        If LFF% = 1 Then ABUS$ = Chr$(Int(XVALO(VALOR)))
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
    If QTR& < 1 Then
        Call CreaRegistro("SQL_LISELEC", "NomLiSele", NOMBREDELISTA$)
    End If
    Call ACTUREGISTRO("SQL_LISELEC", "ACol_1", "NomLiSele = '" & NOMBREDELISTA$ & "'", ACOL1%, RAFE&)
    Call ACTUREGISTRO("SQL_LISELEC", "ACol_2", "NomLiSele = '" & NOMBREDELISTA$ & "'", ACOL2%, RAFE&)
    Call ACTUREGISTRO("SQL_LISELEC", "Tabla_Origen", "NomLiSele = '" & NOMBREDELISTA$ & "'", TABLA$, RAFE&)
    Call ACTUREGISTRO("SQL_LISELEC", "Lista_Campos", "NomLiSele = '" & NOMBREDELISTA$ & "'", LISTADECAMPOS$, RAFE&)
    Call ACTUREGISTRO("SQL_LISELEC", "Condi_Selec", "NomLiSele = '" & NOMBREDELISTA$ & "'", CONDICION$, RAFE&)
    If Left$(NOMBREDELISTA$, 1) = "!" Then
          Call SAVEFILE(LUCOM$ + Mid$(NOMBREDELISTA$, 2) + ".SEL", TELIS$)
        Else
          Call ACTUREGISTRO("SQL_LISELEC", "Fech_Actu", "NomLiSele = '" & NOMBREDELISTA$ & "'", AHORA$, RAFE&)
          Call ACTUREGISTRO("SQL_LISELEC", "LiSelec", "NomLiSele = '" & NOMBREDELISTA$ & "'", TELIS$, RAFE&)
    End If
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
       JJ& = 0
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
         JJ& = JJ& + 1
         ' Call GRAREG(NOMBREDELISTA$, TTYY$, JJ&)   ' CORRECTO ASI LA
         Mid$(TEDAT$, POSIN&, AREBLA&) = TTYY$
         POSIN& = POSIN& + AREBLA&
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
    JJ& = 0
    '
    STRSQL = "SELECT COD_QUERY FROM SQL_MENUS WHERE Form_Name = '" & FORNAME$ & "' ORDER BY Num_Orden Asc "
    On Error Resume Next
    Set MIRS = BaDaConsul.OpenRecordset(STRSQL, dbOpenSnapshot)
    MIRS.MoveFirst
    If Err Then
       On Error GoTo 0
       Call MENSERR(24, "No Hay Consultas Preconfiguradas\para este Formulario.")
       GoTo 99
    End If
    On Error GoTo 0
    Do While Not MIRS.EOF
       COCU$ = MIRS!COD_QUERY
       'DECU$ = ECOARCH$("INCONSQL", COCU$)
       CONDI$ = "COD_QUERY = '" & COCU$ & "'"
       DECU$ = VALOTACO("SQL_QUERIES", "DES_QUERY", CONDI$)
       TTXX$ = AJUSTI$(COCU$, 16) + DECU$
       JJ& = JJ& + 1
       Call GRAREG("!INDIASIE", TTXX$, JJ&)
    MIRS.MoveNext
    Loop
    MIRS.Close
    Call SETULTREG("!INDIASIE", JJ&)
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
    Call ABRETACON
    STRSQL = "SELECT COD_QUERY, DES_QUERY FROM SQL_QUERIES ORDER BY COD_QUERY ASC"
    On Error Resume Next
    Set MIRS = BaDaConsul.OpenRecordset(STRSQL, dbOpenSnapshot)
    MIRS.MoveFirst
    If Err Then
       On Error GoTo 0
       Call MENSERR(24, "No se Encontraron Consultas Preconfiguradas.")
       Exit Sub
    End If
    '
    JJ& = 0
    On Error GoTo 0
    If Not (MIRS.EOF And MIRS.BOF) Then
       Do While Not MIRS.EOF
         TTXX$ = AJUSTI$(MIRS!COD_QUERY, 16) + MIRS!DES_QUERY
         JJ& = JJ& + 1
         Call GRAREG("INCONSQL", TTXX$, JJ&)
       MIRS.MoveNext
       Loop
       Call SETULTREG("INCONSQL", JJ&)
    End If
    '
    JJ& = 0
    STRSQL = "SELECT COD_QUERY FROM SQL_MENUS WHERE Form_Name = '" & FORNAME$ & "' ORDER BY Num_Orden Asc "
    On Error Resume Next
    Set MIRS = BaDaConsul.OpenRecordset(STRSQL, dbOpenSnapshot)
    MIRS.MoveFirst
    If Err Then
       On Error GoTo 0
       GoTo 25
    End If
    On Error GoTo 0
    Do While Not MIRS.EOF
       COCU$ = MIRS!COD_QUERY
       DECU$ = ECOARCH$("INCONSQL", COCU$)
       TTXX$ = AJUSTI$(COCU$, 16) + DECU$
       JJ& = JJ& + 1
       Call GRAREG("!INDIASIE", TTXX$, JJ&)
    MIRS.MoveNext
    Loop
25    MIRS.Close
    Call SETULTREG("!INDIASIE", JJ&)
    '
    VTB% = VENTABU%("ACONPRE")
    If VTB% >= 0 Then
       BaDaConsul.Execute ("DELETE FROM SQL_MENUS WHERE Form_Name = '" & FORNAME$ & "'")
       STRSQL = "SELECT * FROM SQL_MENUS WHERE Form_Name = '" & FORNAME$ & "' ORDER BY Num_Orden Asc "
       Set MIRS = BaDaConsul.OpenRecordset(STRSQL, dbOpenDynaset)
       For j& = 1 To ULTREG&("!INDIASIE")
          XX$ = REGLEIDO$("!INDIASIE", j&)
          MIRS.AddNew   ' YA CONVERTIDO
              On Error Resume Next
          MIRS!CodiEmpr = CodiEmp%
              On Error GoTo 0
          MIRS!Form_Name = FORNAME$
          MIRS!COD_QUERY = TRIM$(Left$(XX$, 16))
          MIRS!Num_Orden = j&
          MIRS.Update
       Next j&
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
         !CodiEmpr = CodiEmp%
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
Function READSET(SQXX$, Optional TIMUT) As ADODB.Recordset
    '
    Call ABRECONEXION
    
    If XVALO(TIMUT) > 0 Then
       FLEXCONN.CommandTimeout = TIMUT
       FIN& = FLEXCONN.CommandTimeout
    End If
    '
10  Set READSET = New ADODB.Recordset

    If InStr(UCase$(SQXX$), "WITH(NOLOCK)") < 1 Then
       TTX1$ = "Error de Invocación en Función 'READSET' - Falta WITH(NOLOCK) En Consulta: " + vbCrLf
       TTX1$ = TTX1$ + SQXX$
       MsgBox TTX1$
       Call FINAL("")
    End If
    '
    SQYY$ = FILTSQL$(SQXX$)
    On Error Resume Next
    Set READSET = FLEXCONN.Execute(SQYY$)
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
    On Error GoTo 0
    '
End Function


Function EXISTELREP%(FORNAME$)
    '
    EXREP% = 0
    Call ABRETACON
    '
    STRSQL = "SELECT * FROM SQL_QUERIES WHERE COD_QUERY = '" & FORNAME$ & "'"
    On Error Resume Next
    Set MIRS = BaDaConsul.OpenRecordset(STRSQL, dbOpenSnapshot)
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
Sub CARGALISTA_LIST_GRID(NOMBREDELISTA, TABLA$, LISTADECAMPOS$, CONDICION$, Optional MODOSEL$, Optional ENCABEZADO$, Optional COLUMNA_AGRUPA%)
    '
    Dim NOCAMPO$(), FORMX$()
    ReDim NOCAMPO$(1), FORMX$(1)
    '
    MODOSE$ = TRIM$(UCase$(SAFETEXT$(MODOSEL$)))
    '
    '
    TIPODELISTA = "GRID"
    On Error Resume Next
    CANCOL& = NOMBREDELISTA.Cols
    If Err Then
       On Error GoTo 0
       TIPODELISTA = "LIST"
    End If
    On Error GoTo 0
    '
    If TIPODELISTA = "LIST" Then
        On Error Resume Next
        NOMBREDELISTA.Clear
        If Err Then
           On Error GoTo 0
           Call MENSERR(24, "Imposible Limpiar Lista '" & NOMBREDELISTA & "'.")
           Call FINAL("")
        End If
        On Error GoTo 0
    ElseIf TIPODELISTA = "GRID" Then
        NOMBREDELISTA.Rows = 1
    End If
    
    If TIPODELISTA = "GRID" Then
       NOMBREDELISTA.Redraw = False
       NOMBREDELISTA.Rows = 2
       ALTUREN = NOMBREDELISTA.RowHeight(0) + NOMBREDELISTA.RowHeight(1)
       While ALTUREN < NOMBREDELISTA.Height
          NOMBREDELISTA.Rows = NOMBREDELISTA.Rows + 1
          ALTUREN = ALTUREN + NOMBREDELISTA.RowHeight(NOMBREDELISTA.Rows - 1)
       Wend
       If TRIM$(ENCABEZADO$) <> "" Then
          NOMBREDELISTA.FormatString = ENCABEZADO$
       End If
       
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
       TTXX$ = ""
       j% = 0: FILA& = 0
       Do While Not .EOF
           TTXX$ = "": COLUI% = 0
           '
           If TIPODELISTA = "GRID" Then
              FILA& = FILA& + 1
37            If FILA& > NOMBREDELISTA.Rows - 1 Then
                 NOMBREDELISTA.Rows = NOMBREDELISTA.Rows + 1
                 GoTo 37
              End If
           End If
           '
           For U& = 1 To .Fields.Count
              VALOR = .Fields(U& - 1).Value
              FORMATX$ = FORMX$(U&) ' FORMATO
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
              TTYY$ = FORMATNUM$(VALOR, FORMATX$)
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
                   If FILA& = 1 Then
                      FORMALTER.TEPRUEBA.Font = NOMBREDELISTA.Font
                      FORMALTER.TEPRUEBA.FontSize = NOMBREDELISTA.FontSize
                      FORMALTER.TEPRUEBA.Caption = "l" + TMASCARA$(FORMATX$)
                      NOMBREDELISTA.ColWidth(COLUI% - 1) = FORMALTER.TEPRUEBA.Width
                      NOMBREDELISTA.COLALIGNEMENT(COLUI% - 1) = 7
                      If Left$(FORMATX$, 1) = "A" Or Left$(FORMATX$, 1) = "C" Or Left$(FORMATX$, 1) = "D" Then
                         NOMBREDELISTA.COLALIGNEMENT(COLUI% - 1) = 1
                      End If
                      NOMBREDELISTA.FontBold = True
                      NOMBREDELISTA.TextMatrix(0, COLUI% - 1) = ENCAB$
                      NOMBREDELISTA.FontBold = False
                   End If
                   If Left$(FORMATX$, 1) = "A" Then NOMBREDELISTA.COLALIGNEMENT(COLUI% - 1) = 1: TTYY$ = VALOR   ' PARA QUE EN LA GRILLA LO PONGA COMPLETO
                   
'                   TTYY$ = "HOLA OMAR"
                   NOMBREDELISTA.TextMatrix(FILA&, COLUI% - 1) = TTYY$
              End If
              '
              'PARA QUE AGRUPE POR LA COLUMNA PASADA COMO PARAMETRO
              If COLUMNA_AGRUPA% > 0 Then
                NOMBREDELISTA.MergeCells = flexMergeRestrictColumns
                NOMBREDELISTA.MergeCol(COLUMNA_AGRUPA%) = True
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
                   VALOECO = RASOCLI$(XVALO(VALOR))
                 ElseIf Left$(TECO$, 5) = "PROVE" Or Left$(TECO$, 5) = "ACRED" Then
                   VALOECO = RASOCLI$((-1) * XVALO(VALOR))
                 ElseIf Left$(TECO$, 5) = "MASTE" Then
                   TXVALO$ = VALOR
                     If Len(TXVALO$) > 4 Then
                          CIXI% = CODINT%(TXVALO$)
                        Else
                          CIXI% = XVALO(VALOR)
                     End If
                   VALOECO = DESCRIT0$(CIXI%)
                 ElseIf EXISTE%(LUDAT$ + TECO$ + ".RFS") Then
                   LFF% = LENFIELD%(TECO$, 1)
                   ABUS$ = MKS$(XVALO(VALOR))
                   On Error Resume Next
                   If LFF% = 2 Then ABUS$ = MKI$(Int(XVALO(VALOR)))
                   If LFF% = 1 Then ABUS$ = Chr$(Int(XVALO(VALOR)))
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
                      If FILA& = 1 Then
                         FORMALTER.TEPRUEBA.Font = NOMBREDELISTA.Font
                         FORMALTER.TEPRUEBA.FontSize = NOMBREDELISTA.FontSize
                         FORMALTER.TEPRUEBA.Caption = "l" + TMASCARA$(FORMATX$)
                         NOMBREDELISTA.ColWidth(COLUI% - 1) = FORMALTER.TEPRUEBA.Width
                         NOMBREDELISTA.COLALIGNEMENT(COLUI% - 1) = 7
                         If Left$(FORMATX$, 1) = "A" Or Left$(FORMATX$, 1) = "C" Or Left$(FORMATX$, 1) = "D" Then
                            NOMBREDELISTA.COLALIGNEMENT(COLUI% - 1) = 1
                         End If
                         NOMBREDELISTA.FontBold = True
                         NOMBREDELISTA.TextMatrix(0, COLUI% - 1 + 1) = ENCAB$
                         NOMBREDELISTA.FontBold = False
                      End If
                      If Left$(FORMATX$, 1) = "A" Then TTYY$ = VALOECO       ' PARA QUE EN LA GRILLA LO PONGA COMPLETO
                      NOMBREDELISTA.TextMatrix(FILA&, COLUI% - 1) = TTYY$
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
       NOMBREDELISTA.Redraw = True
    End If
    '
    TABTMP.Close
End Sub

Function MAXNUMCOD&(TABLA$, CAMPO$)
   CONDI$ = CAMPO$ & " > 0"
   MAXNUMCOD& = XVALO(VALOBADA(TABLA$, "MAX(" & CAMPO$ & ")", CONDI$))
   
End Function

Sub GENERATABLAIMPRE()
   '
   Call ABRECONEXION
   Call CREACAMPO("", "TABLA_IMPRE", "CAMPO1", 10, 48)
   Call CREACAMPO("", "TABLA_IMPRE", "CAMPO2", 10, 48)
   Call CREACAMPO("", "TABLA_IMPRE", "CAMPO3", 10, 48)
   Call CREACAMPO("", "TABLA_IMPRE", "CAMPO4", 10, 48)
   Call CREACAMPO("", "TABLA_IMPRE", "CAMPO5", 10, 48)
   Call CREACAMPO("", "TABLA_IMPRE", "CAMPO6", 10, 48)
   Call CREACAMPO("", "TABLA_IMPRE", "CAMPO7", 10, 48)
   Call CREACAMPO("", "TABLA_IMPRE", "CAMPO8", 10, 48)
   Call CREACAMPO("", "TABLA_IMPRE", "CAMPO9", 10, 48)
   Call CREACAMPO("", "TABLA_IMPRE", "CAMPO10", 10, 48)
   Call CREACAMPO("", "TABLA_IMPRE", "CAMPO11", 10, 48)
   Call CREACAMPO("", "TABLA_IMPRE", "CAMPO12", 10, 48)
   Call CREACAMPO("", "TABLA_IMPRE", "CAMPO13", 10, 48)
   Call CREACAMPO("", "TABLA_IMPRE", "CAMPO14", 10, 48)
   Call CREACAMPO("", "TABLA_IMPRE", "CAMPO15", 10, 48)
   Call CREACAMPO("", "TABLA_IMPRE", "CAMPO16", 10, 48)
   '
End Sub

