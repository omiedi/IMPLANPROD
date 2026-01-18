Attribute VB_Name = "MODULO_SP"
    Sub CREAR_STORED_PROCEDURE()
    Static YaCreoStoredProcedure%
    If YaCreoStoredProcedure% > 0 Then Exit Sub
    YaCreoStoredProcedure% = 1
    '
'por ahora lo comento por que no funciona bien, solo deberia mostrar saldos pendientes a fecha hasta y no en un rango especifico
'    'AMBEST SALDO DE FACTURAS ANTICIPOS SIN APLICAR (ENTRE FECHAS)
'    If EXISTE_SP%("ZPXANTICIPOSPENDIENTES") > 0 Then
'      SQLX$ = "alter Procedure ZPXANTICIPOSPENDIENTES "
'    Else
'      SQLX$ = "CREATE Procedure ZPXANTICIPOSPENDIENTES "
'    End If
'    SQLX$ = SQLX$ + "("
'
'    SQLX$ = SQLX$ + "@FECHA1 SMALLDATETIME, @FECHA2 SMALLDATETIME"
'
'    SQLX$ = SQLX$ + ")"
'    SQLX$ = SQLX$ + " AS"
'    SQLX$ = SQLX$ + " Select SUM(IMPODEBE) - SUM(IMPOHABE) AS RESULTADO"
'    SQLX$ = SQLX$ + " from factuanti with(nolock) WHERE  (tipomovim = 'FVEN' OR tipomovim = 'AJUS')"
'    SQLX$ = SQLX$ + " AND FECHEMISI >= @FECHA1 AND FECHEMISI <= @FECHA2"
'
'    FLEXCONN.Execute (SQLX$)
'-------------------------------------------------------------------------------------------------------------------
    
    'AMBEST SALDO DE FACTURAS ANTICIPOS SIN APLICAR (A UNA FECHA)
    If EXISTE_SP%("ZPXANTICIPOSPEND_ACUMULADOS") > 0 Then
      SQLX$ = "alter Procedure ZPXANTICIPOSPEND_ACUMULADOS "
    Else
      SQLX$ = "CREATE Procedure ZPXANTICIPOSPEND_ACUMULADOS "
    End If
    SQLX$ = SQLX$ + "("
    
    SQLX$ = SQLX$ + " @FECHA2 SMALLDATETIME"
    
    SQLX$ = SQLX$ + ")"
    SQLX$ = SQLX$ + " AS"
    SQLX$ = SQLX$ + " Select SUM(IMPODEBE) - SUM(IMPOHABE) AS RESULTADO"
    SQLX$ = SQLX$ + " from factuanti with(nolock) WHERE  (tipomovim = 'FVEN' OR tipomovim = 'AJUS')"
    SQLX$ = SQLX$ + " AND FECHEMISI <= @FECHA2"

    FLEXCONN.Execute (SQLX$)
'-------------------------------------------------------------------------------------------------------------------
  'AMBEST SALDO DE FACTURAS ANTICIPOS SIN APLICAR (A UNA FECHA) ESTA REPETIDO PERO NO BORRAR.
    If EXISTE_SP%("ANTICIPOSPEND_ACUM_PREVIOS") > 0 Then
      SQLX$ = "alter Procedure ANTICIPOSPEND_ACUM_PREVIOS "
    Else
      SQLX$ = "CREATE Procedure ANTICIPOSPEND_ACUM_PREVIOS "
    End If
    SQLX$ = SQLX$ + "("
    
    SQLX$ = SQLX$ + " @FECHA2 SMALLDATETIME"
    
    SQLX$ = SQLX$ + ")"
    SQLX$ = SQLX$ + " AS"
    SQLX$ = SQLX$ + " Select SUM(IMPODEBE) - SUM(IMPOHABE) AS RESULTADO"
    SQLX$ = SQLX$ + " from factuanti with(nolock) WHERE  (tipomovim = 'FVEN' OR tipomovim = 'AJUS')"
    SQLX$ = SQLX$ + " AND FECHEMISI <= @FECHA2"

    FLEXCONN.Execute (SQLX$)
 '-------------------------------------------------------------------------------------------------
 'COSTO CALCULADO DEL MASTER PARAMETRO CODIGO INTERNO
If EXISTE_SP%("COSTOCALCULADO") > 0 Then
  SQLX$ = "Alter "
Else
  SQLX$ = "Create "
  EJECUTA_DIRECTO% = 1
End If
SQLX$ = SQLX$ + " PROCEDURE COSTOCALCULADO "
SQLX$ = SQLX$ + "( "
SQLX$ = SQLX$ + "@Codint INT "
SQLX$ = SQLX$ + ") "
SQLX$ = SQLX$ + "AS "
SQLX$ = SQLX$ + " "
SQLX$ = SQLX$ + "SELECT CosCal "
SQLX$ = SQLX$ + "FROM Master "
SQLX$ = SQLX$ + "WHERE Codint =  @Codint "
SQLX$ = SQLX$ + " "
SQLX$ = SQLX$ + "EXEC COSTOCALCULADO 4 "
If EJECUTA_DIRECTO% > 0 Then
  FLEXCONN.Execute (SQLX$)
Else
  If CAMBIOSP("COSTOCALCULADO", SQLX$) Then
   FLEXCONN.Execute (SQLX$)
  End If
End If

End Sub
Function CAMBIOSP(NAME_SP$, TEXTONUEVO$) As Boolean
'-------------------------------------------------------------------------------------------------------------------
'VALIDA SI HUBO CAMBIOS EN EL STORED PROCEDURE

    TXNEW$ = TEXTONUEVO$
    SQLX$ = "exec sp_helptext " & NAME_SP$
    Dim RS1 As ADODB.Recordset
    Set RS1 = New ADODB.Recordset
    On Error Resume Next
    RS1.Open SQLX$, FLEXCONN, adOpenKeyset, adLockReadOnly, adCmdText
    If Abs(Err.Number) > 0 Then On Error GoTo 0: GoTo 90
    On Error GoTo 0
    TEXTOEXISTENTE$ = ""
    With RS1
      Do While Not .EOF
        TEXTOEXISTENTE$ = TEXTOEXISTENTE$ & RS1!TEXT
        .MoveNext
      Loop
    End With
    RS1.Close
    Set RS1 = Nothing
    'EN EL SQLSERVER QUEDA COMO CREATE POR ESO LO CAMBIO PARA LA COMPARACION
    If Left$(TRIM$(TEXTONUEVO$), 6) = "ALTER " Then
       TXNEW$ = REPLACAR(TXNEW$, "ALTER ", "CREATE ")
    End If
    'EN EL SQLSERVER LE AGREGA [dbo].[AB] (LO QUITO PARA LA COMPARACION)
    TEXTOEXISTENTE$ = REPLACAR$(TEXTOEXISTENTE$, "[dbo].[" & NAME_SP$ & "]", NAME_SP$)
    
    If TEXTOEXISTENTE$ = TXNEW$ Then
       CAMBIOSP = False
    Else
90     CAMBIOSP = True
    End If
    
End Function

Function EXISTE_SP%(SP$)
    EXISTE_SP% = 0
    Call ABRECONEXION
    
    SQLX$ = "IF EXISTS (SELECT * FROM sysobjects WHERE id = OBJECT_ID(N'[dbo].[" & SP$
    SQLX$ = SQLX$ + "]'))"
    SQLX$ = SQLX$ + " BEGIN"
    SQLX$ = SQLX$ + " select 'El objeto existe' "
    SQLX$ = SQLX$ + " End"
    SQLX$ = SQLX$ + " IF OBJECT_ID('dbo." & SP$ & "'"
    SQLX$ = SQLX$ + ") IS  NULL"
    SQLX$ = SQLX$ + " BEGIN"
    SQLX$ = SQLX$ + " select 'El objeto no existe'"
    SQLX$ = SQLX$ + " End"
    Dim RS As New ADODB.Recordset
    
    Set RS = FLEXCONN.Execute(SQLX$)
    RESULTADO$ = RS.Fields(0)
    If RESULTADO$ = "El objeto existe" Then EXISTE_SP% = 1
    RS.Close
    Set RS = Nothing
    

End Function

