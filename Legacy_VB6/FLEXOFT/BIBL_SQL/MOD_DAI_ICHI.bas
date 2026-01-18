Attribute VB_Name = "MOD_DAI_ICHI"
Function SELCARAS() As Integer
   
   On Error Resume Next
   RESP_ZELE_VECT(1) = ""
   On Error GoTo 0
   CAMPOS$ = "CODIGO/F8;REFERENCIA/A32"
   Call CARGA_ENCABEZADO(FRMZELECHO.msf2, CAMPOS$, FRMZELECHO)
   Call CARGA_ENCABEZADO(FRMZELECHO.MSF1, CAMPOS$, FRMZELECHO)
   FRMZELECHO.Caption = ""
 
   Dim obj As New RECORXET
   Set RS = obj.RS_CARAS()
   Call Carga_MSF_Recordset(RS, CAMPOS$, FRMZELECHO.msf2, 1)
   Set obj = Nothing
   RS.Close
   Set RS = Nothing

   FRMZELECHO.Width = 8500
   FRMZELECHO.Show 1
   If TRIM$(RESP_ZELE_VECT(1)) <> "" Then
     SELCARAS = TRIM$(RESP_ZELE_VECT(1))
   End If
End Function
Function SELSINO() As Integer
   
   On Error Resume Next
   RESP_ZELE_VECT(1) = ""
   On Error GoTo 0
   CAMPOS$ = "CODIGO/F8;REFERENCIA/A32"
   Call CARGA_ENCABEZADO(FRMZELECHO.msf2, CAMPOS$, FRMZELECHO)
   Call CARGA_ENCABEZADO(FRMZELECHO.MSF1, CAMPOS$, FRMZELECHO)
   FRMZELECHO.Caption = ""
 
   Dim obj As New RECORXET
   Set RS = obj.RS_SINO()
   Call Carga_MSF_Recordset(RS, CAMPOS$, FRMZELECHO.msf2, 1)
   Set obj = Nothing
   RS.Close
   Set RS = Nothing

   FRMZELECHO.Width = 8500
   FRMZELECHO.Show 1
   If TRIM$(RESP_ZELE_VECT(1)) <> "" Then
     SELSINO = TRIM$(RESP_ZELE_VECT(1))
   End If
End Function

Function REFECARAS$(COD%)
   REFECARAS$ = VALOBADA("DATASQL", "CAMPO2", "LISTA  = 'CARAS' AND CAST(CAMPO1 AS INTEGER)= " & COD%, "NOMESS")
'CAST(SUBSTRING(" & CAMPO1$ & ",1,16) AS Integer )
End Function
Function REFESINO$(COD%)
   REFESINO$ = VALOBADA("DATASQL", "CAMPO2", "LISTA  = 'TABLASI_NO' AND CAST(CAMPO1 AS INTEGER)= " & COD%, "NOMESS")
End Function


Function SELESPEMATE() As Integer
   
   On Error Resume Next
   RESP_ZELE_VECT(1) = ""
   On Error GoTo 0
   CAMPOS$ = "CODIGO/F8;REFERENCIA/A32"
   Call CARGA_ENCABEZADO(FRMZELECHO.msf2, CAMPOS$, FRMZELECHO)
   Call CARGA_ENCABEZADO(FRMZELECHO.MSF1, CAMPOS$, FRMZELECHO)
   FRMZELECHO.Caption = ""
 
   Dim obj As New RECORXET
   Set RS = obj.RS_ESPEMATE()
   Call Carga_MSF_Recordset(RS, CAMPOS$, FRMZELECHO.msf2, 1)
   Set obj = Nothing
   RS.Close
   Set RS = Nothing

   FRMZELECHO.Width = 8500
   FRMZELECHO.Show 1
   If TRIM$(RESP_ZELE_VECT(1)) <> "" Then
     SELESPEMATE = TRIM$(RESP_ZELE_VECT(1))
   End If
End Function
Function REFEESPEMATE$(COD%)
   REFEESPEMATE$ = VALOBADA("DATASQL", "CAMPO2", "LISTA  = 'ESPEMATE' AND CAST(CAMPO1 AS INTEGER)= " & COD%, "NOMESS")
End Function


Function SELEPTH() As Integer
   
   On Error Resume Next
   RESP_ZELE_VECT(1) = ""
   On Error GoTo 0
   CAMPOS$ = "CODIGO/F8;REFERENCIA/A32"
   Call CARGA_ENCABEZADO(FRMZELECHO.msf2, CAMPOS$, FRMZELECHO)
   Call CARGA_ENCABEZADO(FRMZELECHO.MSF1, CAMPOS$, FRMZELECHO)
   FRMZELECHO.Caption = ""
 
   Dim obj As New RECORXET
   Set RS = obj.RS_ESPEPTH()
   Call Carga_MSF_Recordset(RS, CAMPOS$, FRMZELECHO.msf2, 1)
   Set obj = Nothing
   RS.Close
   Set RS = Nothing

   FRMZELECHO.Width = 8500
   FRMZELECHO.Show 1
   If TRIM$(RESP_ZELE_VECT(1)) <> "" Then
     SELEPTH = TRIM$(RESP_ZELE_VECT(1))
   End If
End Function
Function REFEPTH$(COD%)
   REFEPTH$ = VALOBADA("DATASQL", "CAMPO2", "LISTA  = 'PTH' AND CAST(CAMPO1 AS INTEGER)= " & COD%, "NOMESS")
End Function


Function SELEPISTA() As Integer
   
   On Error Resume Next
   RESP_ZELE_VECT(1) = ""
   On Error GoTo 0
   CAMPOS$ = "CODIGO/F8;REFERENCIA/A32"
   Call CARGA_ENCABEZADO(FRMZELECHO.msf2, CAMPOS$, FRMZELECHO)
   Call CARGA_ENCABEZADO(FRMZELECHO.MSF1, CAMPOS$, FRMZELECHO)
   FRMZELECHO.Caption = ""
 
   Dim obj As New RECORXET
   Set RS = obj.RS_PISTAS()
   Call Carga_MSF_Recordset(RS, CAMPOS$, FRMZELECHO.msf2, 1)
   Set obj = Nothing
   RS.Close
   Set RS = Nothing

   FRMZELECHO.Width = 8500
   FRMZELECHO.Show 1
   If TRIM$(RESP_ZELE_VECT(1)) <> "" Then
     SELEPISTA = TRIM$(RESP_ZELE_VECT(1))
   End If
End Function
Function REFEPISTA$(COD%)
   REFEPISTA$ = VALOBADA("DATASQL", "CAMPO2", "LISTA  = 'PISTAS' AND CAST(CAMPO1 AS INTEGER)= " & COD%, "NOMESS")
End Function





Function SELEPECOBRE() As Integer
   
   On Error Resume Next
   RESP_ZELE_VECT(1) = ""
   On Error GoTo 0
   CAMPOS$ = "CODIGO/F8;REFERENCIA/A32"
   Call CARGA_ENCABEZADO(FRMZELECHO.msf2, CAMPOS$, FRMZELECHO)
   Call CARGA_ENCABEZADO(FRMZELECHO.MSF1, CAMPOS$, FRMZELECHO)
   FRMZELECHO.Caption = ""
 
   Dim obj As New RECORXET
   Set RS = obj.RS_ESPECOBRE()
   Call Carga_MSF_Recordset(RS, CAMPOS$, FRMZELECHO.msf2, 1)
   Set obj = Nothing
   RS.Close
   Set RS = Nothing

   FRMZELECHO.Width = 8500
   FRMZELECHO.Show 1
   If TRIM$(RESP_ZELE_VECT(1)) <> "" Then
     SELEPECOBRE = TRIM$(RESP_ZELE_VECT(1))
   End If
End Function
Function REFEESPECOBRE(COD%)
   REFEESPECOBRE = VALOBADA("DATASQL", "CAMPO2", "LISTA  = 'ESPECOBRE' AND CAST(CAMPO1 AS INTEGER)= " & COD%, "NOMESS")
End Function



Function SELEPCB() As Integer
   
   On Error Resume Next
   RESP_ZELE_VECT(1) = ""
   On Error GoTo 0
   CAMPOS$ = "CODIGO/F8;REFERENCIA/A32"
   Call CARGA_ENCABEZADO(FRMZELECHO.msf2, CAMPOS$, FRMZELECHO)
   Call CARGA_ENCABEZADO(FRMZELECHO.MSF1, CAMPOS$, FRMZELECHO)
   FRMZELECHO.Caption = ""
 
   Dim obj As New RECORXET
   Set RS = obj.RS_PCB()
   Call Carga_MSF_Recordset(RS, CAMPOS$, FRMZELECHO.msf2, 1)
   Set obj = Nothing
   RS.Close
   Set RS = Nothing

   FRMZELECHO.Width = 8500
   FRMZELECHO.Show 1
   If TRIM$(RESP_ZELE_VECT(1)) <> "" Then
     SELEPCB = TRIM$(RESP_ZELE_VECT(1))
   End If
End Function
Function REFEPCB$(COD%)
   REFEPCB$ = VALOBADA("DATASQL", "CAMPO2", "LISTA  = 'prograpcb' AND CAST(CAMPO1 AS INTEGER)= " & COD%, "NOMESS")
End Function


Function SELEMASCANTSOLDANTE() As Integer
   
   On Error Resume Next
   RESP_ZELE_VECT(1) = ""
   On Error GoTo 0
   CAMPOS$ = "CODIGO/F8;REFERENCIA/A32"
   Call CARGA_ENCABEZADO(FRMZELECHO.msf2, CAMPOS$, FRMZELECHO)
   Call CARGA_ENCABEZADO(FRMZELECHO.MSF1, CAMPOS$, FRMZELECHO)
   FRMZELECHO.Caption = ""
 
   Dim obj As New RECORXET
   Set RS = obj.RS_MASCANTSOLDANTE()
   Call Carga_MSF_Recordset(RS, CAMPOS$, FRMZELECHO.msf2, 1)
   Set obj = Nothing
   RS.Close
   Set RS = Nothing

   FRMZELECHO.Width = 8500
   FRMZELECHO.Show 1
   If TRIM$(RESP_ZELE_VECT(1)) <> "" Then
     SELEMASCANTSOLDANTE = TRIM$(RESP_ZELE_VECT(1))
   End If
End Function
Function REFEMASCANTSOLDANTE$(COD%)
   REFEMASCANTSOLDANTE$ = VALOBADA("DATASQL", "CAMPO2", "LISTA  = 'MASCANTSOLDANTE' AND CAST(CAMPO1 AS INTEGER)= " & COD%, "NOMESS")
End Function

Function SELELEYECOMP() As Integer
   
   On Error Resume Next
   RESP_ZELE_VECT(1) = ""
   On Error GoTo 0
   CAMPOS$ = "CODIGO/F8;REFERENCIA/A32"
   Call CARGA_ENCABEZADO(FRMZELECHO.msf2, CAMPOS$, FRMZELECHO)
   Call CARGA_ENCABEZADO(FRMZELECHO.MSF1, CAMPOS$, FRMZELECHO)
   FRMZELECHO.Caption = ""
 
   Dim obj As New RECORXET
   Set RS = obj.RS_LEYECOMP()
   Call Carga_MSF_Recordset(RS, CAMPOS$, FRMZELECHO.msf2, 1)
   Set obj = Nothing
   RS.Close
   Set RS = Nothing

   FRMZELECHO.Width = 8500
   FRMZELECHO.Show 1
   If TRIM$(RESP_ZELE_VECT(1)) <> "" Then
     SELELEYECOMP = TRIM$(RESP_ZELE_VECT(1))
   End If
End Function
Function REFELEYECOMP$(COD%)
   REFELEYECOMP$ = VALOBADA("DATASQL", "CAMPO2", "LISTA  = 'LEYECOMP' AND CAST(CAMPO1 AS INTEGER)= " & COD%, "NOMESS")
End Function

Function NumNuevaCotizaDaichi()
    NumNuevaCotizaDaichi = 1 + XVALO(VALOBADA("COTIDAICHI", "TOP 1 NRO_COTIZA", "NRO_COTIZA > 0 ORDER BY NRO_COTIZA DESC", "NOMESS"))
End Function
Function LIMITEINFERIOR(TABLA$, valor, RET_ID)
   'DEVUELVE EL VALOR INFERIOR DE LA TABLA LOGARITMICA
   CONDI$ = "LOTE_M2 <= " & valor & " ORDER BY LOTE_M2 DESC"
   
   LIMITEINFERIOR = XVALO(VALOBADA(TABLA$, "LOTE_M2", CONDI$, "NOMESS"))
   RET_ID = XVALO(VALOBADA(TABLA$, TABLA$ & "_ID", CONDI$, "NOMESS"))
End Function
Function LIMITESUPERIOR(TABLA$, valor, RET_ID)
   'DEVUELVE EL VALOR INFERIOR DE LA TABLA LOGARITMICA
   CONDI$ = "LOTE_M2 >= " & valor & " ORDER BY LOTE_M2 "
   LIMITESUPERIOR = XVALO(VALOBADA(TABLA$, "LOTE_M2", CONDI$, "NOMESS"))
   RET_ID = XVALO(VALOBADA(TABLA$, TABLA$ & "_ID", CONDI$, "NOMESS"))
End Function
Function VALORSUPERIOR(TABLA$, CAMPO$, ID_SUP)
   'DEVUELVE EL VALOR SUPERIOR BUSCA POR EL ID DE LA TABLA LOGARITMICA
   CONDI$ = TABLA$ & "_ID = " & ID_SUP
   VALORSUPERIOR = XVALO(VALOBADA(TABLA$, CAMPO$, CONDI$, "NOMESS"))
End Function

Function VALORINFERIOR(TABLA$, CAMPO$, ID_INF)
   'DEVUELVE EL VALOR INFERIOR POR EL ID DE LA TABLA LOGARITMICA
   CONDI$ = TABLA$ & "_ID = " & ID_INF
   VALORINFERIOR = XVALO(VALOBADA(TABLA$, CAMPO$, CONDI$, "NOMESS"))
End Function

Function FACTORCOSTOSUPERIOR(TABLA$, ID_SUP)
   'DEVUELVE EL VALOR SUPERIOR SEGUN EL ID DE LA TABLA LOGARITMICA
   CONDI$ = TABLA$ & "_ID = " & ID_SUP
   FACTORCOSTOSUPERIOR = XVALO(VALOBADA(TABLA$, "FactorCosto", CONDI$, "NOMESS"))
End Function

Function FACTORCOSTOINFERIOR(TABLA$, ID_INF)
   'DEVUELVE EL VALOR INFERIOR SEGUN EL ID DE LA TABLA LOGARITMICA
   CONDI$ = TABLA$ & "_ID = " & ID_INF
   FACTORCOSTOINFERIOR = XVALO(VALOBADA(TABLA$, "FactorCosto", CONDI$, "NOMESS"))
End Function

Function RESULTADOTABALGRIT(TABLA$, CAMPO$, valor)
    LIMIINF = LIMITEINFERIOR(TABLA$, valor, RET_ID_INF)
    LIMSUP = LIMITESUPERIOR(TABLA$, valor, RET_ID_SUP)
    If (LIMIINF + LIMSUP) <= 0 Then RESULTADOTABALGRIT = 0: Exit Function
    '
    If LIMIINF <= 0 Then LIMIINF = LIMSUP: RET_ID_INF = RET_ID_SUP
    If LIMSUP <= 0 Then LIMPSUP = LIMIINF: RET_ID_SUP = RET_ID_INF
    '
    VALINF = VALORINFERIOR(TABLA$, CAMPO$, RET_ID_INF)
    VALSUP = VALORSUPERIOR(TABLA$, CAMPO$, RET_ID_SUP)
    '
    VALINF = ROUND(VALINF, 3)
    VALSUP = ROUND(VALSUP, 3)
'    VALINF = VALINF / FACTORCOSTOINFERIOR(TABLA$, RET_ID_INF)
'    VALSUP = VALSUP * FACTORCOSTOSUPERIOR(TABLA$, RET_ID_SUP)
    On Error Resume Next
    If VALINF <> VALSUP Then
       RESULTADOTABALGRIT = XVALO(VALINF) + (XVALO(valor) - XVALO(LIMIINF)) / (XVALO(LIMSUP) - XVALO(LIMIINF)) * (XVALO(VALSUP) - XVALO(VALINF))
    Else
       'en caso que valor inferior y valorsuperior sean iguales no interpola devuelve el correspondiente a la fila unica
       RESULTADOTABALGRIT = VALINF
    End If
    On Error GoTo 0 '
End Function

   Function CALCULACOSTO(LoteM2, CARAS$, TIPOMATERIAL$, TXTESPESOR, TXTESPCOBRE, TXTPTH$, RET_COSTONORMAL)
        
        CALCULACOSTO = 0
         RET_COSTONORMAL = 0
         If LoteM2 > 0 Then
           If TRIM$(CARAS$) <> "" Then
              If TRIM$(TIPOMATERIAL$) <> "" Then
                 If XVALO(TXTESPESOR) > 0 Then
                    If XVALO(TXTESPCOBRE) > 0 Then
                      If TRIM$(TXTPTH$) <> "" Then
20                        CONDI$ = "LISTA = 'CONFIGLOGA' AND CAMPO2 = '" & TRIM$(CARAS$) & "'"
                          CONDI$ = CONDI$ + " AND CAMPO3 = '" & TRIM$(TIPOMATERIAL$) & "'"
                          CONDI$ = CONDI$ + " AND CAMPO4 = '" & TRIM$(TXTESPESOR) & "'"
                          CONDI$ = CONDI$ + " AND CAMPO5 = '" & TRIM$(TXTESPCOBRE) & "'"
                          CONDI$ = CONDI$ + " AND CAMPO6 = '" & TRIM$(TXTPTH$) & "'"
                          SQLX$ = "SELECT * FROM DATASQL WITH(NOLOCK) WHERE " & CONDI$
                          Set RS = READSET(SQLX$)
                          With RS
                            Do While Not .EOF
                              TABLA$ = TRIM$(SAFETEXT$(!CAMPO8))
                              CAMPO$ = TRIM$(SAFETEXT$(!CAMPO1))
                              If TABLA$ <> "" And CAMPO$ <> "" Then
                                'PUEDE SER UM2PTH, UM2PTHCOS O TABLAUM2 , TABLAUM2COS
                                If Right$(TABLA$, 3) = "COS" Then
                                   CALCULACOSTO = RESULTADOTABALGRIT(TABLA$, CAMPO$, TRIM$(FORMATNUM$(LoteM2, "f16.3")))
                                 Else
                                   RET_COSTONORMAL = RESULTADOTABALGRIT(TABLA$, CAMPO$, TRIM$(FORMATNUM$(LoteM2, "f16.3")))
                                 End If
                              End If
                              .MoveNext
                            Loop
                          End With
                      End If
                    End If
                 End If
              End If
           End If
         End If
   End Function

   
   Function VALIDARINTEGRIDAD_TABLA_CALCULO_LOGARITMICA%()
      'SE VALIDA SI EXISTEN DATOS QUE PUEDEN ARROJAR MAS DE UN REGISTRO COMO RESPUESTA ANTE UNA CONSULTA
      ' ES DECIR SI LA COMBINACION DE LOS DATOS PARA LA FUNCION: CALCULACOSTO PUEDE SER AMBIGUA
      ' SI ASI FUESE DEVULEVE LA CANTIDAD DE REPETICION DE RESULTADO PARA LA MISMA COMBINACION
      
      VALIDARINTEGRIDAD_TABLA_CALCULO_LOGARITMICA% = 0
      SQLX$ = "SELECT COUNT(*), CAMPO1,CAMPO2,CAMPO3,CAMPO4,CAMPO5,CAMPO6, CAMPO8  "
      SQLX$ = SQLX$ + " FROM DATASQL WITH(NOLOCK) WHERE LISTA = 'CONFIGLOGA' "
      SQLX$ = SQLX$ + " GROUP BY CAMPO1,CAMPO2,CAMPO3,CAMPO4,CAMPO5,CAMPO6, CAMPO8  "
      SQLX$ = SQLX$ + " HAVING COUNT(*) > 1"
      Set RS = READSET(SQLX$)
      With RS
        Do While Not .EOF
         VALIDARINTEGRIDAD_TABLA_CALCULO_LOGARITMICA% = VALIDARINTEGRIDAD_TABLA_CALCULO_LOGARITMICA% + 1
         .MoveNext
        Loop
      End With
     
   End Function
   
Function CALCULAPORCENTAJE_ANTICIPO(TEXTOCONDIPAGO$, IMPORTE)
   CALCULAPORCENTAJE_ANTICIPO = 0
   CONP$ = TRIM$(TEXTOCONDIPAGO$)
   CONP$ = Left$(CONP$, 2)
   If IsNumeric(CONP$) Then
      PORCENTAJE = Int(XVALO(CONP$))
      IMPORTE = XVALO(IMPORTE)
      CALCULAPORCENTAJE_ANTICIPO = FORMATNUM$(ResultadoValorSegúnPorcentaje(PORCENTAJE, IMPORTE), "F12.0")
   End If

End Function

Function SecuenciasOptionCheck%(indexObjeto%, ret_Descrip$, ret_Grupo$, ret_TipoObjet$)
   SecuenciasOptionCheck = -1
   'SI NO ENCUENTRA COINCIDENCIA DEVUELVE -1
   'RETORNA 3 VALORES DE LA TABLA DATASQL QUE GUARDA LA RELACION DE CADA OBJETO CON EL TIPO DE DATO
   SQLX$ = "SELECT CAMPO2,CAMPO3,CAMPO4 FROM DATASQL WITH(NOLOCK) WHERE LISTA = 'FILTROSEC'  AND CAST(CAMPO1 AS INTEGER)  = " & indexObjeto%
   Set RS = READSET(SQLX$)
   With RS
      ret_Descrip$ = SAFETEXT$(!CAMPO2)      'EJEMPLO = PTH
      ret_Grupo$ = SAFETEXT$(!CAMPO3)        'EJEMPLO = CIRCUITO
      ret_TipoObjet$ = SAFETEXT$(!CAMPO4)    'EJEMPLO = OptionButton
      SecuenciasOptionCheck = 1
   End With
    
End Function
Sub AltaSecuencia(TEXTOX$)
   'DA DE ALTA UNA NUEVA SECUENCIA
   On Error GoTo ERROR_GUARDAR
   FLEXCONN.BeginTrans
   CONDI$ = "ID > 0"
   PROXIMASEC& = 1 + XVALO(VALOBADA("SECTORES", "Max(SECUENPROD)", CONDI$, "NOMESS"))
   ID& = 1 + XVALO(VALOBADA("SECTORES", "Max(ID)", CONDI$, "NOMESS"))
   MBORRA% = 0
        
   SQLX$ = "INSERT INTO SECTORES "
   SQLX$ = SQLX$ + " (ID,SECUENPROD,SECTOR,CODIEMPR,MARBORRA)"
   SQLX$ = SQLX$ + " VALUES(" & ID&
   SQLX$ = SQLX$ + " ," & PROXIMASEC&
   SQLX$ = SQLX$ + " ,'" & TEXTOX$ & "'"
   SQLX$ = SQLX$ + " ," & CodiEmp%
   SQLX$ = SQLX$ + " ," & MBORRA%
   SQLX$ = SQLX$ + ")"
   FLEXCONN.Execute (SQLX$)
ERROR_GUARDAR:
    If Err.Number <> 0 Then
       FLEXCONN.RollbackTrans
       Call COMUNI("ERROR AL GUARDAR NO SE REALIZO LA TRANSACCION" & vbCrLf & Err.Number & " " & Err.Description)
       On Error GoTo 0
    Else
      FLEXCONN.CommitTrans
      Call COMUNI("Transacción Realizada")
    End If

End Sub
Function CANT_VERSIONES_EXISTENTES&(NROCOTI&)
   CANT_VERSIONES_EXISTENTES& = CantRegistros("COTIDAICHI", "NRO_COTIZA = " & NROCOTI&, "NOMESS")
End Function
Function MODOCOTIZA%(NCOT&)
   MODOCOTIZA% = XVALO(VALOBADA("COTIDAICHI", "MODO", "NRO_COTIZA  = " & NCOT& & " AND MODO > 0", "NOMESS"))
End Function

Sub GRABAMODO(NCOT&)
   If CANT_VERSIONES_EXISTENTES&(NCOT&) > 0 Then
      MODOX% = MODOCOTIZA%(NCOT&)
      Call ACTUREGISTRO("COTIDAICHI", "MODO", "NRO_COTIZA = " & NCOT&, MODOX%, REGAF&, "NOMESS")
   End If

End Sub

Function COTIZACIONDIFERENTESCODIGOS%(NROCOTI&, Optional RET_MODELO$)
       COTIZACIONDIFERENTESCODIGOS% = 0
       'ACA VALIDO SI ES DE DIFERENTES CODIGOS O MISMO CODIGO DIFERENTES CANTIDADES
       VUELTA% = 0: COTIZACIONDIFERENTESCODIGOS% = 0
       SQLX$ = "select  cod_inte,nro_cotiza from cotidaichi WITH(NOLOCK) WHERE NRO_COTIZA = " & NROCOTI& & " group by nro_cotiza, cod_inte "
       Set RSX = READSET(SQLX$)
       
       Do While Not RSX.EOF
            VUELTA% = VUELTA% + 1
            COTIZACIONDIFERENTESCODIGOS% = COTIZACIONDIFERENTESCODIGOS% + 1
            'SI ALGUNO NO TIENE CODIGO DAICHI LE ASIGNA -1 Y SALE DEL WHILE y RETORNA EL MODELO
            If XVALO(RSX!cod_inte) < 1 Then
               COTIZACIONDIFERENTESCODIGOS% = -1
               'VALOBADA POR QUE SI NO ROMPE EL GROUP BY DADO QUE EL MODELO SE ESCRIBE MANUALMENTE.
               CONDI$ = "NRO_COTIZA = " & NROCOTI& & " AND COD_INTE < 1 "
               RET_MODELO$ = VALOBADA("cotidaichi", "MODELO", CONDI$, "NOMESS")
               Exit Do
            End If
         RSX.MoveNext

       Loop
       RSX.Close: Set RSX = Nothing



End Function

