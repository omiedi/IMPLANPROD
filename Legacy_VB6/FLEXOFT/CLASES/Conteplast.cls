VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Conteplast"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Private Type PropCplast
    FILASECTORGRID As String
    
    nombre As String
    Domicilio As String
    localidad As String
    mailcliente As String
    TeleCli As String
    CeluCli As String
    NotaCli As String
    CtacCli As String
    CuitCli As String
End Type
'Declaramos el vector
Private Atrib() As PropCplast
Function CantAcumuladaSubProceso&(IdRegPadrex&)
   'DEVUELVE LA CANTIDAD ACUMULADA EN UN SUBPROCESO AGRUPANDO POR EL ID DEL MOVIMIENTO ORIGINAL PADRE
   'EJEMPLO DE UN CORTE DE LATERALES SE VAN INFORMANDO EL PROCESO DE IMRESION Y PUEDE QUE
   'SE HAGAN PARCIALES SOBRE UN CORTE Y LOTE, CADA VEZ QUE SE GUARDA UN SUB-PROCESO SE ANOTA EL ID DEL PADRE
   CantAcumuladaSubProceso& = XVALO(VALOBADA("INFOFAB", "SUM(CANTREAL * FACTORCORTE)", "IDREGPADRE = " & IdRegPadrex&, "NOMESS"))
End Function

Function CANTIDADINFORMADA(NORFA$, NOMBRESECTOR$, Optional NOINFORMARLINEA%, Optional MODOX%)
    'LAS UBICACIONES EN RACK:  IDCOORDUBIC=0
    'LAS UBICACIONES EN SECORES LAS < 100 SON LAS DE IMPRESION O ACCESORIOS
    'LAS UBICACIONES EN LINEAS LAS > 100  Y < 200 SON LAS DE LINEA
    'VEMOS COMO SIGUE
    'MODOX% = 0 NO FILTRA
    'MODOX% = 1 CON IDREGPADRE > 0
    'MODOX% = 2 CON OPERFUCIONADA > 0
    'MODOX% = 3 CON IDREGPADRE > 0 Y OPERFUCIONADA <> ''
    'MODOX% = 4 CON IDREGPADRE < 1"
    CONDI$ = "IDSUBOR = '" & NORFA$ & "'"
    If TRIM$(NOMBRESECTOR$) <> "" Then CONDI$ = CONDI$ + " AND DESCOPER = '" & NOMBRESECTOR$ & "'"
    If NOINFORMARLINEA% > 0 Then 'NO INFORMA SI INGRESO A LINEA
       CONDI$ = CONDI$ + " AND (IDCOORDUBIC < 100 OR IDCOORDUBIC IS NULL)"
    End If
    If MODOX% = 1 Then
       CONDI$ = CONDI$ + " AND (IDREGPADRE > 0 AND IdRegPadre  IS NOT NULL)"
    ElseIf MODOX% = 2 Then
       CONDI$ = CONDI$ + " AND (OPERFUCIONADA <> '' AND OPERFUCIONADA  IS NOT NULL)"
    ElseIf MODOX% = 3 Then
       CONDI$ = CONDI$ + " AND (OPERFUCIONADA <> '' AND OPERFUCIONADA  IS NOT NULL)"
       CONDI$ = CONDI$ + " AND (IDREGPADRE > 0 AND IdRegPadre  IS NOT NULL)"
    ElseIf MODOX% = 4 Then
       CONDI$ = CONDI$ + " AND (IDREGPADRE < 1 OR IdRegPadre  IS  NULL)"
    End If
    
    CANTIDADINFORMADA = XVALO(VALOBADA("INFOFAB", "SUM(CANTREAL)", CONDI$, "NOMESS"))
    
End Function

Function CONSULTA_OR$(TABLA$, CAMPO_OR$, CONDI$, CAMPO_DEVOLUCION_OR$)
    'CAMPO_OR$= AL CAMPO QUE VA A USAR EN LA CONSULTA DE LA FU OR EJ: (CODMATPRIMA = 10 OR CODMATPRIMA = 12)
    'CONDI$ =  "CODRANGO = 'ABC-123"
    'CAMPO_DEVOLUCION_OR$ = VA A SER EL CAMPO QUE VA A UTILIZAR PARA LA CONDI EN LA LLAMADA _
    'EJ: CAMPO_DEVOLUCION_OR$= COD_EXTE (X Q' EN LA LLAMADA VOY A REALIZAR UN CONSULTA EN EL MOVISTO _
    'BAJO LA CONDICION QUE ME DEVULEVA ESTA FUNCION (COD_EXTE = BCR234 OR COD_EXTE = TRE765)
    'ESTA SERIA LA DEVOLUCION: (COD_EXTE = BCR234 OR COD_EXTE = TRE765)
    J& = 0
    CONSULTA_OR$ = ""
    SQLX$ = "SELECT " & CAMPO_OR$ & " AS CAMPOOR FROM " & TABLA$ & " WITH(NOLOCK)"
    SQLX$ = SQLX$ + " WHERE " & CONDI$
    Set RS = READSET(SQLX$)
    With RS
    Do While Not .EOF
      With RS
        J& = J& + 1
        If CONSULTA_OR$ = "" Then
          CONSULTA_OR$ = "( " & CAMPO_DEVOLUCION_OR$ & " = '" & SAFETEXT$(!CAMPOOR) & "'"
        Else
          CONSULTA_OR$ = CONSULTA_OR$ & " Or " & CAMPO_DEVOLUCION_OR$ & " = '" & SAFETEXT$(!CAMPOOR) & "'"
        End If
        .MoveNext
      End With
    Loop
    End With
    If CONSULTA_OR$ <> "" Then CONSULTA_OR$ = CONSULTA_OR$ & ")"
    RS.Close
    Set RS = Nothing


End Function

Sub ConsumoInfofab(FECHA%, CantCortes, cod_inte, idsubor, Cod_Exte, Descrip, NUMEOPER, descoper, idenlote, codielem, CodXElem, FACTORCORTE, corte, Observa$, REFERENCIA$, idMov&)
       'idMov& es el que se genero cuando se dio de alta en el movisto el movimiento OF (INGRESO DE P.TERMINADO)

        SQLX$ = "INSERT INTO INFOFAB"
        SQLX$ = SQLX$ + "([IdSubOr],[CodiEmpr],[Cod_Inte]"
        SQLX$ = SQLX$ + " ,[Cod_Exte],[Descrip],[NumeOper],[DescOper]"
        SQLX$ = SQLX$ + " ,[CantReal],[FeHorLib]"
        SQLX$ = SQLX$ + " ,[Observa],[Referen],[COORDUBIC],[IDCOORDUBIC]"
        SQLX$ = SQLX$ + " ,[IdenLote],[CodiElem],[CodXElem],[FACTORCORTE] ,[FeHorInf],[Consumo],[IDMOVISTO],[CantSalid]"
        SQLX$ = SQLX$ + " ,[CodiMovi],[CORTE],[PESO])"
        
        SQLX$ = SQLX$ + "  Values "
        
        SQLX$ = SQLX$ + "  ('" & idsubor & "'"
        SQLX$ = SQLX$ + " ,0"
        SQLX$ = SQLX$ + " ," & cod_inte
        SQLX$ = SQLX$ + "  ,'" & Cod_Exte & "'"
        SQLX$ = SQLX$ + "  ,'" & Descrip & "'"
        SQLX$ = SQLX$ + "  , " & NUMEOPER
        SQLX$ = SQLX$ + "  , '" & descoper & "'"
        SQLX$ = SQLX$ + "  , 0"
        SQLX$ = SQLX$ + "  ,'" & FETEXCOR1$(FECHA%) & "'"
        SQLX$ = SQLX$ + "  ,'" & Observa & "'"
        SQLX$ = SQLX$ + "  , '" & REFERENCIA$ & "'"
        SQLX$ = SQLX$ + "  ,'LINEA'"
        SQLX$ = SQLX$ + "  ,100"
        SQLX$ = SQLX$ + "  ,'" & idenlote & "'"
        SQLX$ = SQLX$ + "  , " & codielem
        SQLX$ = SQLX$ + "  ,'" & CodXElem & "'"
        SQLX$ = SQLX$ + "  ," & FACTORCORTE
        SQLX$ = SQLX$ + "  ,'" & FETEXCOR1$(FECHA%) & "'"
        SQLX$ = SQLX$ + "  , 0"
        SQLX$ = SQLX$ + "  , " & idMov&  'idMov& es el que se genero cuando se dio de alta en el movisto el movimiento OF (INGRESO DE P.TERMINADO)
        SQLX$ = SQLX$ + "  ,  '" & TRIM$(FORMATNUM$(CantCortes, "F14.2")) & "'"
        SQLX$ = SQLX$ + "  ,'CF'"
        SQLX$ = SQLX$ + "  , '" & TRIM$(FORMATNUM$(corte, "F14.2")) & "'"
        SQLX$ = SQLX$ + "  , 0"
        SQLX$ = SQLX$ + " )"
'Call mostrarDatosx("1 " & SQLX$)

        FLEXCONN.Execute (SQLX$)
End Sub

Sub ConsumoSegunLosCodigosTransferidosCTP(NORFA$, DEPOCORTE%)



   Dim CTP As New Conteplast
   CICONJ% = CTP.CODIFAB2%(NORFA$)
   'RECORRO LOS CONSUMOS DEL MOVISTO SEGUN LA OF
   SQLX$ = "SELECT  COD_INTE,CANTSALID,MOVISTO_ID FROM MOVISTO WITH(NOLOCK)"
   SQLX$ = SQLX$ + " WHERE CODIMOVI = 'CF' "
   SQLX$ = SQLX$ + " AND DOPRICOR ='" & NORFA$ & "'"
   SQLX$ = SQLX$ + " AND CANTSALID > 0 AND (IDENLOTE = '' OR IDENLOTE IS NULL)"
   SQLX$ = SQLX$ + " ORDER BY MOVISTO_ID DESC"
   Set RS = READSET(SQLX$)
'Call mostrarDatosx(SQLX$)
   With RS
     Do While Not .EOF
       CI1% = XVALO(!cod_inte)
       CANTIRANGO = XVALO(!CANTSALID)
       ID& = XVALO(!MOVISTO_ID)
       'SI ES CODIGO RANGO CONSULTO EL MATERIAL TRANSFERIDO PARA LA ORDEN AL DEPOSITO DE CORTE
       'Y REEMPLAZO LOS CODIGOS RANGO POR LOS TRANSFERIDOS QUE CORRESPONDAN AL RANGO ACTIVO
       If ESCODIGORANGO&(CI1%) > 0 Then
'MsgBox "ES CODIGO RANGO"

         CODIRANGO$ = CODEXT$(CI1%)
         CONDI$ = "CODRANGO = '" & CODIRANGO$ & "'"
         CONDIOR$ = CTP.CONSULTA_OR("CODIGORANGO", "CODMATPRIMA", CONDI$, "COD_EXTE")
         SQLX$ = "SELECT SUM(CANTINGRE-CANTSALID) AS RESULTADO ,COD_INTE,COD_EXTE,DESCRIP,IDENLOTE"
         SQLX$ = SQLX$ + " FROM MOVISTO WITH(NOLOCK)"
         SQLX$ = SQLX$ + " where (DOSECLAR =   '" & TRIM(NORFA$) & "'  OR DOPRICOR = '" & TRIM(NORFA$) & "')"
         SQLX$ = SQLX$ + "  AND DEPOMOVI = " & DEPOCORTE%
         If CONDIOR$ <> "" Then SQLX$ = SQLX$ + " AND " & CONDIOR$
         SQLX$ = SQLX$ + " GROUP BY COD_INTE,COD_EXTE,DESCRIP,IDENLOTE"
         SQLX$ = SQLX$ + " Having Sum(CANTINGRE - CANTSALID) > 0.05"
'Call mostrarDatosx(SQLX$)
         Set rs1 = READSET(SQLX$)
         With rs1
           Do While Not .EOF
              CIMTP = XVALO(!cod_inte)
              CANTITRANS = XVALO(!Resultado)
'MsgBox "CANTIDAD TRANSFERIDA " & CANTITRANS & " LOTE: " & SAFETEXT$(!IDENLOTE)
              SQLX$ = "UPDATE MOVISTO SET COD_INTE = " & XVALO(!cod_inte)
              SQLX$ = SQLX$ + "  ,COD_EXTE = '" & SAFETEXT$(!Cod_Exte) & "'"
              SQLX$ = SQLX$ + "  ,DESCRIP = '" & SAFETEXT$(!Descrip) & "'"
              SQLX$ = SQLX$ + "  ,DESCRI_LAR = '" & SAFETEXT$(!Descrip) & "'"
              SQLX$ = SQLX$ + "  ,IDENLOTE = '" & SAFETEXT$(!idenlote) & "'"
              If CANTITRANS < CANTIRANGO Then
                'si la cantidad transferida es menor a la necesidad actualiza cantsalid
                SQLX$ = SQLX$ + "  ,CANTSALID = '" & FORMATNUM$(CANTITRANS, "F16.5") & "'"
                '
                'Antes de grabar genera duplica el registro por la diferencia(manteniendo el codigo rango)
                NEWID& = 1 'PARA QUE DEVUELVA ID
                '
                'SQLX1$
                Call DuplicarRegistro("MOVISTO", "MOVISTO_ID = " & ID&, NEWID&)
                Diferencia = CANTIRANGO - CANTITRANS: If Diferencia < 1 Then Diferencia = CANTIRANGO
'MsgBox "Diferencia " & Diferencia
                SQLX1$ = "UPDATE MOVISTO SET CANTSALID = '" & FORMATNUM$(Diferencia, "F16.5") & "'"
                SQLX1$ = SQLX1$ + "  WHERE  MOVISTO_ID = " & NEWID&
'Call mostrarDatosx(SQLX1$)
                FLEXCONN.Execute SQLX1$
'MsgBox "Diferencia Aplicado a id " & NEWID&

                'ahora actualiza el registro con el codigo transfereido real incluyendo la cantidad   SQLX$
                SQLX$ = SQLX$ + "  WHERE  MOVISTO_ID = " & ID&
'Call mostrarDatosx(SQLX$)
                FLEXCONN.Execute SQLX$
                ID& = NEWID& 'SI ENTRA POR ESTE IF ACTUALIZA EL ID P Q EL NUEVO REGISTRO ES DUPLICADO DEL ORIGINAL POR LA DIFERENCIA
                             'DADO QUE SI EN LA SIGUIENTE VUELTA NO ENTRA POR EL ELSE  APLICA DIRECTO POR EL ID
              Else
                'SI VIENE ACA ES QUE ALCANZA CON LA CANTIAD LEIDA DEL CODIGO TRANSFERIDO
                SQLX$ = SQLX$ + "  WHERE  MOVISTO_ID = " & ID&
'MsgBox "Aplicacion directa a id " & ID&
'Call mostrarDatosx(SQLX$)
                FLEXCONN.Execute SQLX$
                Exit Do
              End If
              CANTIRANGO = CANTIRANGO - CANTITRANS
              If CANTIRANGO <= 0 Then Exit Do
             .MoveNext
           Loop
         End With
       End If
   
       .MoveNext
     Loop
   End With
   Set CTP = Nothing
End Sub

Function FechaEntregaSegunNroPedyOrdenPedido%(NROPED&, ORDENPED%)
   CONDI$ = "NROPED = " & NROPED& & " AND  NuOrItem = " & ORDENPED%
   FechaEntregaSegunNroPedyOrdenPedido% = CVINT(VALOBADA("PEDDETA", "FeSolEnt", CONDI$, "NOMESS"))

End Function

Function CordenadasUbicaciónSegunMovistoid$(IDMOVISTO&)
    'CONSULTA CON CONCATENACION DE CAMPOS SEPARADOS POR PUNTO Y COMA
    SQLX$ = "select rtrim(CONVERT(char, rack ) )+ ';'"
    SQLX$ = SQLX$ + " + rtrim(CONVERT(char, fila) )+ ';'"
    SQLX$ = SQLX$ + " +  rtrim(CONVERT(char, columna)) AS RESULTADO"
    SQLX$ = SQLX$ + " from ubilote_envap WITH(NOLOCK) "
    SQLX$ = SQLX$ + " Where id_movisto = " & IDMOVISTO&
    Set RS = READSET(SQLX$)
    CordenadasUbicaciónSegunMovistoid$ = SAFETEXT$(RS!Resultado)
End Function
Function ImprimeOrdenesSegunCoti1(CI1%, ROTULOSECTOR$, Optional NROCOT&) As Boolean
'INFORMO FRANCO QUE ESTA Y OTRAS SE VALIDAN CONTRA LA COTIZACION 1 SI TIENEN CODIGO GENERA
'valido en la tabla de coti1 si la misma posee informado para rotulo parametro un codigo es true
  longitud% = TRIM$(Len(ROTULOSECTOR$))
  
   CONDI$ = "CODINTCONJ = " & CI1%
   CONDI$ = CONDI$ + " AND left(COMPONENTE," & longitud% & ") = '" & ROTULOSECTOR$ & "'"
   If NROCOT& > 0 Then CONDI$ = CONDI$ + " AND NROCOTIZACION = " & NROCOT&
   ImprimeOrdenesSegunCoti1 = False
   If XVALO(VALOBADA("COTI1", "CODINTELEM", CONDI$, "NOMESS")) > 0 Then
      ImprimeOrdenesSegunCoti1 = True
   End If
End Function

Sub mostrarDatosx(textox$)
   OBSDETALLE.Text28 = textox$
   OBSDETALLE.Show 1
End Sub

Function MTSCONSUMIDOS(NORFA$, Optional LOTE$, Optional DEPOSITO%)

   SQLX$ = "select SUM(CANTSALID) AS CONSU, COD_EXTE from movisto WITH(NOLOCK) where codimovi = 'Cf'"
   SQLX$ = SQLX$ + " AND DOPRILAR = '" & NORFA$ & "' "
   If TRIM$(LOTE$) <> "" Then SQLX$ = SQLX$ + " AND IDENLOTE = '" & LOTE$ & "'"
   If DEPOSITO% > 0 Then SQLX$ = SQLX$ + " AND DEPOMOVI = " & DEPOSITO%
   SQLX$ = SQLX$ + " GROUP BY COD_EXTE"
   If TRIM$(LOTE$) <> "" Then SQLX$ = SQLX$ + " , IDENLOTE"
   
   Set RS = READSET(SQLX$)
   With RS
    If Not .EOF Then
       MTSCONSUMIDOS = XVALO(!CONSU)
    Else
       MTSCONSUMIDOS = 0
    End If
   End With
   RS.Close
   Set RS = Nothing

End Function

Function NroFTecnicaXCodigo(CI1%)
   CONDI$ = "CODINT = " & CI1%
   NroFTecnicaXCodigo = XVALO(VALOBADA("RECETACTP", "NUMRECETACTP", CONDI$, "NOMESS"))
End Function
Function NroFTecnicaXCodigoBolsa(CI1%)
   CONDI$ = "CODINT = " & CI1%
   NroFTecnicaXCodigoBolsa = XVALO(VALOBADA("P_VTA_BOLSAS", "NroPvta", CONDI$, "NOMESS"))
End Function
Function AltoUtilFTecnicaBolsa(FT&)
   CONDI$ = "NROPVTA = " & FT&
   AltoUtilFTecnicaBolsa = XVALO(VALOBADA("P_VTA_BOLSAS", "ALTOUTILCM", CONDI$, "NOMESS"))
End Function
Function MatBolsaFTecnicaBolsa(FT&)
   CONDI$ = "NROPVTA = " & FT&
   MatBolsaFTecnicaBolsa = VALOBADA("P_VTA_BOLSAS", "MAT_BOLSA", CONDI$, "NOMESS")
End Function

Function CantHiloBolsaFTecnicaBolsa(FT&)
   CONDI$ = "NROPVTA = " & FT&
   CantHiloBolsaFTecnicaBolsa = VALOBADA("P_VTA_BOLSAS", "CantHilos", CONDI$, "NOMESS")
End Function
Function ObservAdicionalesFTecnicaBolsa(FT&)
   CONDI$ = "NROPVTA = " & FT&
   ObservAdicionalesFTecnicaBolsa = VALOBADA("P_VTA_BOLSAS", "ObservAdicionales", CONDI$, "NOMESS")
End Function


'Public Property Get RUTULOSECTORGRID(i As Variant) As String
'   On Error Resume Next
'   RUTULOSECTORGRID = Atrib(i).RUTULOSECTORGRID
'   On Error GoTo 0
'End Property
''
'Public Property Let RUTULOSECTORGRID(VALOR As String)
'   On Error Resume Next
'   Atrib(i).FILASECTORGRID = VALOR
'   On Error GoTo 0
'End Property
Public Function ROTULOSECTORGRID(CODICONJ, codielem) As String
   COMPONENTEX = ""
   'DEVUELVE EL TEXTO DEL ROTULO DE LA GRILA QUE REFIERE AL SECTOR EN LA COTIZACION (EJEMPLO POLLERA,ESLINGA)
   SQLX$ = "SELECT COMPONENTE FROM coti1 WITH(NOLOCK)"
   SQLX$ = SQLX + " Where codintconj = " & CODICONJ
   SQLX$ = SQLX + " and codintelem =  " & codielem
   Set RS = READSET(SQLX$)
   With RS
     If Not .EOF Then
       COMPONENTEX = XVALO(!COMPONENTE)
     End If
   End With
   
   RS.Close
   Set RS = Nothing
   'AHORA LE QUITO LO QUE ESTA LUEGO DEL PARENTESIS EJ "TAPA (200-200-50)" ME QUEDE TAPA
   ROTULOSECTORGRID = COMPONENTEX
   If TRIM$(ROTULOSECTORGRID) <> "" Then
       POS1% = InStr(ROTULOSECTORGRID, "(")
       If POS1% > 0 Then
          ROTULOSECTORGRID = TRIM$(Mid$(ROTULOSECTORGRID, 1, POS1%))
       End If
   End If
End Function
Public Function ImprimeOTImpresion(CI1%, ROTULOSECTOR$) As Boolean
'    ORDEN DE IMPRESIÓN: “Impresión”: Si
    IMPRIME$ = VALOBADA("RECETACTP", "IMPRESION", "CODINT = " & CI1%, "NOMESS")
    If IMPRIME$ = "SI" Then
       ImprimeOTImpresion = True
    Else
       ImprimeOTImpresion = False
    End If
End Function
Public Function ImprimeOTCuerpos(CI1%, ROTULOSECTOR$) As Boolean
'INFORMO LAUTARO QUE SE DEBE CUMPLIR LAS TRES SITUACIONES
'    ORDEN DE CORTE DE CUERPOS: “Código de confección”: debe figurar “CFT” ; “Alto (Cm)”: tiene cualquier valor ingresado y “Cuerpo (g/m2)”: tiene cualquier valor ingresado
    CODBB$ = VALOBADA("RECETACTP", "CODBIGBAG", "CODINT = " & CI1%, "NOMESS")
    ALTO = XVALO(VALOBADA("RECETACTP", "CPOALTOCM", "CODINT = " & CI1%, "NOMESS"))
    CPOGMJE$ = TRIM$(VALOBADA("RECETACTP", "GRMJECUERPO", "CODINT = " & CI1%, "NOMESS"))
    If UCase$(Left(TRIM$(CODBB$), 3)) = "CFT" And ALTO > 0 And CPOGMJE$ <> "" Then
       ImprimeOTCuerpos = True
    Else
       ImprimeOTCuerpos = False
    End If
End Function
Function CODIGOBIGBAG$(CI1%)
    CODIGOBIGBAG$ = VALOBADA("RECETACTP", "CODBIGBAG", "CODINT = " & CI1%, "NOMESS")
End Function
Function TIPOPALLET$(CI1%)
    TIPOPALLET$ = VALOBADA("RECETACTP", "TIPOPALLET", "CODINT = " & CI1%, "NOMESS")
End Function
Function EMBUFARDO$(CI1%)
    EMBUFARDO = VALOBADA("RECETACTP", "EMBUFARDO", "CODINT = " & CI1%, "NOMESS")
End Function
Function UNXPALLET(CI1%)
    UNXPALLET = VALOBADA("RECETACTP", "UNXPALLET", "CODINT = " & CI1%, "NOMESS")
End Function
Function OBSGENERAL$(CI1%)
    OBSGENERAL$ = VALOBADA("RECETACTP", "OBSGENERAL", "CODINT = " & CI1%, "NOMESS")
End Function

Function JUSTA$(CI1%)
    JUSTA$ = VALOBADA("RECETACTP", "JUSTA", "CODINT = " & CI1%, "NOMESS")
End Function
Function DESCRIBIGBAG$(CODBB$)
    DESCRIBIGBAG$ = VALOBADA("CONFIGBB", "DESCRIPCIONBB", "CODBIGBAG = '" & CODBB$ & "'", "NOMESS")
    
End Function

Public Function ImprimeOTLaterales(CI1%, ROTULOSECTOR$) As Boolean
'INFORMO LAUTARO QUE SE DEBE CUMPLIR LAS TRES SITUACIONES
'ORDEN DE CORTE DE LATERALES: “Código de confección”: debe figurar “CFP/CFPM” y “Alto cm”: tiene algún valor ingresado
    CODBB$ = VALOBADA("RECETACTP", "CODBIGBAG", "CODINT = " & CI1%, "NOMESS")
    ALTO = XVALO(VALOBADA("RECETACTP", "CPOALTOCM", "CODINT = " & CI1%, "NOMESS"))
    CPOGMJE$ = TRIM$(VALOBADA("RECETACTP", "GRMJECUERPO", "CODINT = " & CI1%, "NOMESS"))
    If (UCase$(Left(TRIM$(CODBB$), 3)) = "CFP" Or UCase$(Left(TRIM$(CODBB$), 4)) = "CFPM") And ALTO > 0 Then
       ImprimeOTLaterales = True
    Else
       ImprimeOTLaterales = False
    End If
End Function

Public Function ImprimeOTFondos(CI1%, ROTULOSECTOR$) As Boolean
'INFORMO LAUTARO QUE SE DEBE CUMPLIR LAS TRES SITUACIONES
'ORDEN DE CORTE DE FONDOS: “Código de confección”: Cualquier “CFT/CFPM”; “largo (Cm)”: tiene cualquier valor ingresado y “Fondo (g/m2)”: tiene cualquier valor ingresado
    CODBB$ = VALOBADA("RECETACTP", "CODBIGBAG", "CODINT = " & CI1%, "NOMESS")
    Largo = XVALO(VALOBADA("RECETACTP", "CPOLARGOCM", "CODINT = " & CI1%, "NOMESS"))
    GMJEFONDO$ = TRIM$(VALOBADA("RECETACTP", "GRMJEFONDO", "CODINT = " & CI1%, "NOMESS"))
    If (UCase$(Left(TRIM$(CODBB$), 3)) = "CFT" Or UCase$(Left(TRIM$(CODBB$), 4)) = "CFPM") And Largo > 0 And GMJEFONDO$ <> "" Then
       ImprimeOTFondos = True
    Else
       ImprimeOTFondos = False
    End If
End Function


Public Function ImprimeOTUPanel(CI1%, ROTULOSECTOR$) As Boolean
'INFORMO LAUTARO QUE SE DEBE CUMPLIR LAS TRES SITUACIONES
'ORDEN DE CORTE DE U-PANEL: “Código de confección”: Cualquier “CFP”;  “largo (Cm)”: tiene cualquier valor ingresado y “Fondo (g/m2)”: tiene cualquier valor ingresado
    CODBB$ = VALOBADA("RECETACTP", "CODBIGBAG", "CODINT = " & CI1%, "NOMESS")
    Largo = XVALO(VALOBADA("RECETACTP", "CPOLARGOCM", "CODINT = " & CI1%, "NOMESS"))
    GMJEFONDO$ = TRIM$(VALOBADA("RECETACTP", "GRMJEFONDO", "CODINT = " & CI1%, "NOMESS"))
    If (UCase$(Left(TRIM$(CODBB$), 3)) = "CFP" Or UCase$(Left(TRIM$(CODBB$), 4)) = "CFPM") And Largo > 0 And GMJEFONDO$ <> "" Then
       ImprimeOTUPanel = True
    Else
       ImprimeOTUPanel = False
    End If
End Function


Public Function ImprimeOTPollera(CI1%, ROTULOSECTOR$) As Boolean
'INFORMO LAUTARO QUE SE DEBE CUMPLIR LAS TRES SITUACIONES
'ORDEN DE CORTE DE POLLERA: “Código de confección”: CFT 11/41, CFP 11/41, CFPM 11/41; “Pollera (largo en cm)”: tiene cualquier valor ingresado y “Pollera (g/m2)”: tiene cualquier valor ingresado
    CODBB$ = TRIM$(UCase$(VALOBADA("RECETACTP", "CODBIGBAG", "CODINT = " & CI1%, "NOMESS")))
    'POLLERALARGO PUEDE CONTENER NO O UN VALOR POR ESO LE HAGO UN XVALO
    POLLARGO = XVALO(VALOBADA("RECETACTP", "POLLARGO", "CODINT = " & CI1%, "NOMESS"))
    POLLERAGM2$ = TRIM$(VALOBADA("RECETACTP", "VPDPOLLERAGM2", "CODINT = " & CI1%, "NOMESS"))
    ImprimeOTPollera = False
    Select Case CODBB$
       Case "CFT11", "CFT41", "CFP11", "CFP41", "CFPM11", "CFPM41"
          If POLLARGO > 0 And POLLERAGM2 <> "" Then
             ImprimeOTPollera = True
          End If
    End Select
    
End Function

Public Function ImprimeOTTapas(CI1%, ROTULOSECTOR$) As Boolean
'INFORMO LAUTARO QUE SE DEBE CUMPLIR LAS TRES SITUACIONES
'ORDEN DE CORTE DE TAPAS: “Código de confección”: CFT 12/13/42/43, CFP 12/13/42/43, CFPM 12/13/42/43; “VC (Diámetro en cm)”: tiene cualquier valor ingresado y “Tapa (g/m2)”: tiene cualquier valor ingresado
    CODBB$ = TRIM$(UCase$(VALOBADA("RECETACTP", "CODBIGBAG", "CODINT = " & CI1%, "NOMESS")))
    'POLLERALARGO PUEDE CONTENER NO O UN VALOR POR ESO LE HAGO UN XVALO
    VCDIAMETROCMS = XVALO(VALOBADA("RECETACTP", "VALVVCDIAMETRO", "CODINT = " & CI1%, "NOMESS"))
    GRAMAJETAPA$ = TRIM$(VALOBADA("RECETACTP", "GRMJETAPA", "CODINT = " & CI1%, "NOMESS"))
    ImprimeOTTapas = False
    Select Case CODBB$
       Case "CFT12", "CFT13", "CFT42", "CFT43", "CFP12", "CFP13", "CFP42", "CFP43", "CFPM12", "CFPM13", "CFPM42", "CFPM43"
          If VCDIAMETROCMS > 0 And GRAMAJETAPA$ <> "" And UCase$(GRAMAJETAPA$) <> "NO" Then
             ImprimeOTTapas = True
          End If
    End Select
    
End Function
Public Function ImprimeOTMamparos(CI1%, ROTULOSECTOR$) As Boolean
'INFORMO LAUTARO QUE SE DEBE CUMPLIR LAS TRES SITUACIONES
'ORDEN DE CORTE DE MAMPAROS: Si “Código de confección”: Cualquier “CFPM”; “Compartimentado”: Si ; “Largo del mamparo (Cm)”: tiene cualquier valor ingresado
    CODBB$ = TRIM$(UCase$(VALOBADA("RECETACTP", "CODBIGBAG", "CODINT = " & CI1%, "NOMESS")))
    'POLLERALARGO PUEDE CONTENER NO O UN VALOR POR ESO LE HAGO UN XVALO
    KOMPARTIMENTADO = TRIM$(UCase$((VALOBADA("RECETACTP", "COMPARTIMENTADO", "CODINT = " & CI1%, "NOMESS"))))
    LARGOMAMPARO = XVALO(VALOBADA("RECETACTP", "VPDLARGOMAMAPAROCM", "CODINT = " & CI1%, "NOMESS"))
    ImprimeOTMamparos = False
    If Left$(CODBB$, 4) = "CFPM" And KOMPARTIMENTADO = "SI" And LARGOMAMPARO > 0 Then
      ImprimeOTMamparos = True
    End If
End Function
Public Function ImprimeOTEslingas(CI1%, ROTULOSECTOR$) As Boolean
'INFORMO LAUTARO QUE SE DEBE CUMPLIR LAS TRES SITUACIONES
'ORDEN DE CORTE DE ESLINGAS: “Eslinga (ojo cm)”, “Eslinga (medida 1 cm)” y “Eslinga (medida 2 cm)”: tienen cualquier valor ingresado
    ESLOJOCM = XVALO(VALOBADA("RECETACTP", "ESLOJOCM", "CODINT = " & CI1%, "NOMESS"))
    'POLLERALARGO PUEDE CONTENER NO O UN VALOR POR ESO LE HAGO UN XVALO
    ESLMEDIDA1CM = XVALO((VALOBADA("RECETACTP", "ESLMEDI1CM", "CODINT = " & CI1%, "NOMESS")))
    ESLMEDIDA2CM = XVALO(VALOBADA("RECETACTP", "ESLMED2CM", "CODINT = " & CI1%, "NOMESS"))
    ImprimeOTEslingas = False
    If ESLOJOCM > 0 And ESLMEDIDA1CM > 0 And ESLMEDIDA2CM > 0 Then
        ImprimeOTEslingas = True
    End If
End Function
'DEBERIA CONSIDERARSE PRIMERO SI DEBE IMPRIMIRSE SEGUN ESTAS FUNCIONES Y LUEGO VER CON LA FUNCION ROTULOSECTORGRID
'PARA PASAR LOS PARAMETROS DE IMPRESION Y MANDAR A IMPRMIR EL ARCHIVO SEGUN LO CONFIGURADO EN ORDEFABR



Sub AgregarOperacionesSegúnCódigo(CI1%, Optional NROCOTIZACION&)
    'RECORRO LOS SECTORES PRODUCTIVOS Y DE ACUERDO A ESTOS VEO SI SEGUN CODIGO LE CORRRESONDE OPERACION
    'SI LE CORRESONDE GENERO LAS OPERACIONES PARA EL CODIGO.
    Dim CTPL As New Conteplast
    'SON LOS SECTORES QUE ESTAN CONFIGURADOS EN SETUP DE ORDENES DE FABRICACION DONDE ASOCIA FRM.
    
    'ESTE ES PARA EL CASO DE LAS BOLSAS CON CODIGO CTP-
    Codigo$ = CODEXT$(CI1%)
    If Left$(Codigo$, 4) = "CTP-" Then 'EL SISTEMA GENERA AUTOMATICAMENTE PARA LAS BOLSAS CON EL PREFIJO CTP-
       Call AGREGAROPERACION(CI1%, "FABRICACIÓN DE BOLSA")
       Exit Sub
    End If
    '
    FIN& = ULTREG&("SECPROD")
    For REG& = 1 To FIN&
      SPX$ = REGLEIDO$("SECPROD", REG&)
      If CVS(SPX$) >= 1 Then
        TTXYZ$ = FILTERGETRECORD$("SECPROD", 1, SPX$)
        CODSECT = CVS(Left(TTXYZ$, 4))
        REFE$ = TRIM$(Mid$(TTXYZ$, 5, 78))
        FRMIM$ = UCase$(TRIM$(Mid$(TTXYZ$, 83, 8)))

        Select Case REFE$
           Case "IMPRESION"
              If CTPL.ImprimeOTImpresion(CI1%, REFE$) Then
                 Call AGREGAROPERACION(CI1%, REFE$)
              End If
           Case "CUERPO"
              If CTPL.ImprimeOrdenesSegunCoti1(CI1%, REFE$, NROCOTIZACION&) Then
                 Call AGREGAROPERACION(CI1%, REFE$)
              End If
           Case "LATERALES"
              If CTPL.ImprimeOrdenesSegunCoti1(CI1%, REFE$, NROCOTIZACION&) Then
                 Call AGREGAROPERACION(CI1%, REFE$)
              End If
           Case "FONDO"
              If CTPL.ImprimeOrdenesSegunCoti1(CI1%, REFE$, NROCOTIZACION&) Then
                 Call AGREGAROPERACION(CI1%, REFE$)
              End If
           Case "UPANEL"
              If CTPL.ImprimeOrdenesSegunCoti1(CI1%, REFE$, NROCOTIZACION&) Then
                 Call AGREGAROPERACION(CI1%, REFE$)
              End If
           Case "POLLERA"
              If CTPL.ImprimeOrdenesSegunCoti1(CI1%, REFE$, NROCOTIZACION&) Then
                 Call AGREGAROPERACION(CI1%, REFE$)
              End If
           Case "TAPA"
              If CTPL.ImprimeOrdenesSegunCoti1(CI1%, REFE$, NROCOTIZACION&) Then
                 Call AGREGAROPERACION(CI1%, REFE$)
              End If
           Case "MAMPARO [CM]"
              If CTPL.ImprimeOrdenesSegunCoti1(CI1%, REFE$, NROCOTIZACION&) Then
                 Call AGREGAROPERACION(CI1%, REFE$)
              End If
           Case "ESLINGAS [CM]"
              If CTPL.ImprimeOrdenesSegunCoti1(CI1%, REFE$, NROCOTIZACION&) Then
                 Call AGREGAROPERACION(CI1%, REFE$)
              End If
           Case "VÁLVULAS DE CARGA [CM]"
              If CTPL.ImprimeOrdenesSegunCoti1(CI1%, REFE$, NROCOTIZACION&) Then
                 Call AGREGAROPERACION(CI1%, REFE$)
              End If
           Case "VÁLVULAS DE DESCARGA [CM]"
              If CTPL.ImprimeOrdenesSegunCoti1(CI1%, REFE$, NROCOTIZACION&) Then
                 Call AGREGAROPERACION(CI1%, REFE$)
              End If
           Case "CINTILLAS C [CM]"
              If CTPL.ImprimeOrdenesSegunCoti1(CI1%, REFE$, NROCOTIZACION&) Then
                 Call AGREGAROPERACION(CI1%, REFE$)
              End If
           Case "CINTILLA CS [CM]"
              If CTPL.ImprimeOrdenesSegunCoti1(CI1%, REFE$, NROCOTIZACION&) Then
                 Call AGREGAROPERACION(CI1%, REFE$)
              End If
           Case "CINTILLA POLLERA [CM]"
              If CTPL.ImprimeOrdenesSegunCoti1(CI1%, REFE$, NROCOTIZACION&) Then
                 Call AGREGAROPERACION(CI1%, REFE$)
              End If
           Case "CINTILLA DS [CM]"
              If CTPL.ImprimeOrdenesSegunCoti1(CI1%, REFE$, NROCOTIZACION&) Then
                 Call AGREGAROPERACION(CI1%, REFE$)
              End If
           Case "CINTILLAS D [CM]"
              If ImprimeOrdenesSegunCoti1(CI1%, REFE$, NROCOTIZACION&) Then
                 Call AGREGAROPERACION(CI1%, REFE$)
              End If
           Case "CINTILLA LINER [CM]"
              If ImprimeOrdenesSegunCoti1(CI1%, REFE$, NROCOTIZACION&) Then
                 Call AGREGAROPERACION(CI1%, REFE$)
              End If
           Case "ESLINGA PERIMETRAL [CM]"
              If ImprimeOrdenesSegunCoti1(CI1%, REFE$, NROCOTIZACION&) Then
                 Call AGREGAROPERACION(CI1%, REFE$)
              End If
           Case "FONDO U"
              If ImprimeOrdenesSegunCoti1(CI1%, REFE$, NROCOTIZACION&) Then
                 Call AGREGAROPERACION(CI1%, REFE$)
              End If
           Case "COMPARTIMENTO"
              If ImprimeOrdenesSegunCoti1(CI1%, REFE$, NROCOTIZACION&) Then
                 Call AGREGAROPERACION(CI1%, REFE$)
              End If
           Case "ANTIFUGA"
              If ImprimeOrdenesSegunCoti1(CI1%, REFE$, NROCOTIZACION&) Then
                 Call AGREGAROPERACION(CI1%, REFE$)
              End If
              Case "TUBO [CM]"
              If ImprimeOrdenesSegunCoti1(CI1%, REFE$, NROCOTIZACION&) Then
                 Call AGREGAROPERACION(CI1%, REFE$)
              End If
              Case "ETIQUETA"
              If ImprimeOrdenesSegunCoti1(CI1%, REFE$, NROCOTIZACION&) Then
                 Call AGREGAROPERACION(CI1%, REFE$)
              End If
           Case "PORTAETIQUETA"
              If ImprimeOrdenesSegunCoti1(CI1%, REFE$, NROCOTIZACION&) Then
                 Call AGREGAROPERACION(CI1%, REFE$)
              End If
           Case "SOMBRERO"
              If ImprimeOrdenesSegunCoti1(CI1%, REFE$, NROCOTIZACION&) Then
                 Call AGREGAROPERACION(CI1%, REFE$)
              End If
           Case "CORDON D"
              If ImprimeOrdenesSegunCoti1(CI1%, REFE$, NROCOTIZACION&) Then
                 Call AGREGAROPERACION(CI1%, REFE$)
              End If

           Case "CORDON DS"
              If ImprimeOrdenesSegunCoti1(CI1%, REFE$, NROCOTIZACION&) Then
                 Call AGREGAROPERACION(CI1%, REFE$)
              End If

        End Select
        
      End If
    Next REG&
    '
    'GENERA OPERACION SI LLEVA LINER
    NROFT& = NroFTecnicaXCodigo(CI1%)
    LINER_SI_NO$ = UCase$(TRIM$(VALOBADA("RECETACTP", "LINERSINO", "NUMRECETACTP = " & NROFT&, "NOMESS")))
    If LINER_SI_NO$ = "SI" Then
       'SOLO AGREGA OPERACION SE IMPRIME ON DEMAND
       Call AGREGAROPERACION(CI1%, "LINER")
    End If
    '

    
End Sub
Sub AGREGAROPERACION(CI1%, NombreOperacion$)
   'AGREGA SOLO SI NO EXISTE
   CONDI$ = "DESCOPER = '" & NombreOperacion$ & "' And CODICONJ = " & CI1%
   

   If CantRegistros("Secoper", CONDI$, "NOMESS") > 0 Then Exit Sub
   CANTOP% = CantRegistros("Secoper", "CODICONJ = " & CI1%, "NOMESS")
   OPERACIONMAYOR% = XVALO(VALOBADA("SECOPER", "MAX(NUMEOPER)", "CODICONJ = " & CI1%, "NOMESS"))
   FE% = HOY(HO$)

   SQLX$ = "SELECT * FROM Secoper "
   Dim RS As ADODB.Recordset
   Set RS = New ADODB.Recordset
25 On Error Resume Next
   RS.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
    If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
      On Error GoTo 0
      Call CapturaErrorOpen
      GoTo 25
    End If
    On Error GoTo 0
    With RS
      .AddNew
      !CODICONJ = CI1%
      !CODXCONJ = CODEXT$(CI1%)
      ULI& = CANTOP%
      NOPY% = 10 * (1 + Int(OPERACIONMAYOR%) / 10)
      !NUMEOPER = NOPY%
      '!AltOper = Asc(Mid$(RESECO$(LL% - 1), 88, 1))
      !descoper = TRIM$(NombreOperacion$)
      !StatOper = 1
      !Fech_Revis = CVDAT(FE%)
      .Update
    End With


End Sub

Function FRMSECTOR$(NOMBRESECTOR$)
    NOMBRESECTOR$ = UCase$(TRIM$(NOMBRESECTOR$))
    POS1% = InStr(NOMBRESECTOR$, "(")
    If POS1% > 0 Then NOMBRESECTOR$ = TRIM$(Mid$(NOMBRESECTOR$, 1, POS1% - 1))

    FIN& = ULTREG&("SECPROD")
    For REG& = 1 To FIN&
      SPX$ = REGLEIDO$("SECPROD", REG&)
      If CVS(SPX$) >= 1 Then
        TTXYZ$ = FILTERGETRECORD$("SECPROD", 1, SPX$)
        CODSECT = CVS(Left(TTXYZ$, 4))
        REFE$ = UCase$(TRIM$(Mid$(TTXYZ$, 5, 78)))
        If REFE$ = NOMBRESECTOR$ Then
          FRMSECTOR$ = UCase$(TRIM$(Mid$(TTXYZ$, 83, 8)))
          Exit For
        End If
      End If
    Next REG&
End Function

Function NumeOperXNombre%(CI1%, NombreOperacion$)
'  NombreOperacion$ = REPLACAR$(NombreOperacion$, "[CM]", "")
  NombreOperacion$ = UCase$(TRIM$(NombreOperacion$))
  CONDI$ = "CODICONJ = " & CI1%

  CONDI$ = CONDI$ + " AND DESCOPER = '" & NombreOperacion$ & "'"
  NumeOperXNombre% = XVALO(VALOBADA("SECOPER", "NUMEOPER", CONDI$, "NOMESS"))
  
End Function

Function ValidarCodigosOF%(CICONJ%, CIELEM%)
   REG& = 0
   CONDI$ = " CODICONJ =" & CICONJ%
   CONDI$ = CONDI$ + " AND (CODIELEM = " & CIELEM%
   CIRANGO& = CODIGORANGO&(CIELEM%)
   If CIRANGO& > 0 Then
      CONDI$ = CONDI$ + " OR CODIELEM = " & CIRANGO&
   End If
   CONDI$ = CONDI$ + " )"
   ValidarCodigosOF% = CantRegistros&("FORMULAS", CONDI$, "NOMESS")

End Function

Function LARGOCTP(CI1%)
    LARGOCTP = XVALO(VALOBADA("RECETACTP", "CPOLARGOCM", "CODINT = " & CI1%, "NOMESS"))
End Function
Function ALTOCTP(CI1%)
    ALTOCTP = XVALO(VALOBADA("RECETACTP", "CPOALTOCM", "CODINT = " & CI1%, "NOMESS"))
End Function
Function ANCHOCTP(CI1%)
    ANCHOCTP = XVALO(VALOBADA("RECETACTP", "CPOANCHOCM", "CODINT = " & CI1%, "NOMESS"))
End Function
Function CodigoOriginalCoti1SegunComponente(NROCOTIZACION&, NOMBRESECTOR$)
   'PASA QUE PUEDE VENIR EL SECTOR COMO UPANEL O UPANEL [CM] POR QUE FUE UN CAMBIO QUE SE PIDIO EN EL MEDIO
   'Y QUEDE ATRAPADO EN MEDIO DE DOS CIRCUNSTANCIAS
   NOMBRESECTOR$ = TRIM$(NOMBRESECTOR$)
   HastaParentesis% = InStr(NOMBRESECTOR$, "(")
   HastaCorchete% = InStr(NOMBRESECTOR$, "[")
   If HastaParentesis% > 0 Then
     NOMBRESECTOR$ = TRIM$(Mid$(NOMBRESECTOR$, 1, HastaParentesis%))
     LongNombreSector% = Len(NOMBRESECTOR$)
   ElseIf HastaCorchete% > 0 Then
     NOMBRESECTOR$ = TRIM$(Mid$(NOMBRESECTOR$, 1, HastaCorchete% - 1))
     LongNombreSector% = Len(NOMBRESECTOR$)
   Else
     LongNombreSector% = Len(NOMBRESECTOR$)
   End If
   CONDI$ = "NROCOTIZACION= " & NROCOTIZACION& & " AND  left(COMPONENTE," & LongNombreSector% & ") = '" & NOMBRESECTOR$ & "'"
   CodigoOriginalCoti1SegunComponente = XVALO(VALOBADA("COTI1", "CODINTELEM", CONDI$, "NOMESS"))
   
End Function
Function CodintPlanVenta(NroPvta&)
   CodintPlanVenta = XVALO(VALOBADA("P_VTA_BOLSAS", "Codint", "CODINT > 0 and NROPVTA = " & NroPvta&, "NOMESS"))
End Function
Sub CargaOperacionesInformadas(NORFA$, CIFAB%, MSF As msFlexGrid, Optional MAYOROIGUAL&, Optional DESCRIPCIONOPERACION$, Optional MENOROIGUAL&, Optional UBICACIONESRACK$)
    '< 40 SECTOR DE CORTE
    '>= 100 LINEA <= LINEA
    '>=200 EMBALAJE
    '>= 300 DEPURACION POR CIERRE DE OF.
    
    MSF.Rows = 1
    Dim CTP As New Conteplast
    CANTI = CantiProyectadaOF2(NORFA$)
    'DEVUELVE LAS OPERACIONES INFORAMDAS SEGUN OF.
    SQLX$ = " SELECT * FROM INFOFAB WITH(NOLOCK)"
    SQLX$ = SQLX$ + " WHERE IDSUBOR = '" & NORFA$ & "'"
    SQLX$ = SQLX$ + " AND CANTREAL > 0 "
    If XVALO(MENOROIGUAL&) > 0 Then SQLX$ = SQLX$ + " AND (IDCOORDUBIC <= " & MENOROIGUAL& & " or IDCOORDUBIC is null)"
    If XVALO(MAYOROIGUAL&) > 0 Then SQLX$ = SQLX$ + " AND IDCOORDUBIC >= " & MAYOROIGUAL&
    
    If TRIM$(DESCRIPCIONOPERACION$) <> "" Then SQLX$ = SQLX$ + " AND DESCOPER = '" & DESCRIPCIONOPERACION$ & "'"
    SQLX$ = SQLX$ + " order by  NUMEOPER,FeHorInf,Cod_inte,IDENLOTE"

    Set RS = READSET(SQLX$)
    With RS
      Do While Not .EOF
        REG& = REG& + 1
        MSF.Rows = REG& + 1
'        VER DE GUARDAR EL CODIGO ELEMENTO EN EL INFOFAB PARA LUEGO LEERLO
        DESOP$ = SAFETEXT$(!descoper)
        CIELEM% = XVALO(!codielem)
        CONDI$ = "Codintconj = " & CIFAB%
        CONDI$ = CONDI$ + " and codintelem = " & CIELEM%
        If DESOP$ <> "" Then
           CONDI$ = CONDI$ + " and LEFT(COMPONENTE"
           CONDI$ = CONDI$ + " ," & Len(DESOP$)
           CONDI$ = CONDI$ + ") = '" & DESOP$ & "'"
        End If
       CORTEX = XVALO(VALOBADA("Coti1", "CORTE", CONDI$, "NOMESS")): If CORTEX < 1 Then CORTEX = 1
        MSF.TextMatrix(REG&, 1) = SAFETEXT$(!Cod_Exte)
        MSF.TextMatrix(REG&, 2) = SAFETEXT$(!Descrip)
        MSF.TextMatrix(REG&, 3) = SAFETEXT$(!idenlote)
        MSF.TextMatrix(REG&, 4) = FORMATNUM$(XVALO(!CANTREAL) * 1, "F10.1")
        MSF.TextMatrix(REG&, 5) = CI1%
        MSF.TextMatrix(REG&, 6) = XVALO(!NUMEOPER)
        MSF.TextMatrix(REG&, 7) = SAFETEXT$(!descoper)
        MSF.TextMatrix(REG&, 8) = SAFETEXT$(!TiemEjec)
        MSF.TextMatrix(REG&, 9) = SAFETEXT$(!CodXElem)
        MSF.TextMatrix(REG&, 10) = SAFETEXT$(!COORDUBIC)
        MSF.TextMatrix(REG&, 11) = XVALO(!NuOrInf) ' NUMERO DE ORDEN DE INFORME
        MSF.TextMatrix(REG&, 12) = XVALO(!IDCOORDUBIC) ' NUMERO DE CODIGO DE UBICACION
        MSF.TextMatrix(REG&, 13) = XVALO(!NUMOPERARIO) ' NUMERO DE OPERARIO
        MSF.TextMatrix(REG&, 14) = SAFETEXT(!NameOperario) ' NOMBRE DE OPERARIO
        MSF.TextMatrix(REG&, 15) = TRIM$(FORMATNUM$(XVALO(!Peso), "F12.2")) ' PESO INFORMADO
        MSF.TextMatrix(REG&, 16) = XVALO(!INFOFAB_ID)
80      .MoveNext
      Loop
    End With
    RS.Close
    Set RS = Nothing
    
End Sub

Sub NuevaCargaSaldoOperacionesInformadasALinea(NORFA$, CIFAB%, MSF As msFlexGrid, Optional MAYOROIGUAL&, Optional DESCRIPCIONOPERACION$, Optional MENOROIGUAL&, Optional UBICACIONESRACK$)
    '< 40 SECTOR DE CORTE
    '>= 100 LINEA <= LINEA
    '>=200 EMBALAJE
    '>= 300 DEPURACION POR CIERRE DE OF.
    
'TRAE TODOS LOS CORTES QUE ESTAN INFORMADOS QUE TODAVIA NO PASARON A LINEA, NO SE AGRUPA NI SE CALCULA SALDOS
'POR QUE ESTOS SE VAN A CONSUMIR UNA VEZ QUE ESTEN EN LINEA

    MSF.Rows = 1
    Dim CTP As New Conteplast
    CANTI = CantiProyectadaOF2(NORFA$)
    SQLX$ = " SELECT * FROM INFOFAB WITH(NOLOCK)"
    SQLX$ = SQLX$ + " WHERE IDSUBOR = '" & NORFA$ & "'"
    SQLX$ = " SELECT SUM(CASE WHEN cantreal > 0 THEN PESO  ELSE -PESO END) AS PesoKg,COUNT(*) AS CANTREG"
    SQLX$ = SQLX$ + " ,SUM(CASE WHEN cantreal > 0 THEN cantreal  ELSE -cantsalid END) AS saldoCantiCortes,"
    SQLX$ = SQLX$ + "  COUNT(*) AS CANTREG ,descrip,cod_inte,idenlote, NumeOper,CODIELEM"
    SQLX$ = SQLX$ + "     ,descoper,codxelem ,Cod_Exte,COORDUBIC,IDCOORDUBIC"
    SQLX$ = SQLX$ + "     ,NUMOPERARIO,NUMOPERARIO,NameOperario"
    SQLX$ = SQLX$ + " From  InfoFab WITH(NOLOCK)"
    SQLX$ = SQLX$ + " WHERE    IDSUBOR = '" & NORFA$ & "' AND Cod_Inte = " & CIFAB%
    If XVALO(MENOROIGUAL&) > 0 Then SQLX$ = SQLX$ + " AND (IDCOORDUBIC <= " & MENOROIGUAL& & " or IDCOORDUBIC is null)"
    If XVALO(MAYOROIGUAL&) > 0 Then SQLX$ = SQLX$ + " AND IDCOORDUBIC >= " & MAYOROIGUAL&
    
    If TRIM$(DESCRIPCIONOPERACION$) <> "" Then SQLX$ = SQLX$ + " AND DESCOPER = '" & DESCRIPCIONOPERACION$ & "'"
    
    SQLX$ = SQLX$ + " group by Cod_Exte,idenlote,NumeOper,COORDUBIC,IDCOORDUBIC,idenlote,descrip,cod_inte,idenlote, descoper"
    SQLX$ = SQLX$ + " , codxelem,NUMOPERARIO,NUMOPERARIO,NameOperario,CODIELEM"
    SQLX$ = SQLX$ + " order by COORDUBIC, NUMEOPER,Cod_inte,IDENLOTE"
    Set RS = READSET(SQLX$)
    With RS
      Do While Not .EOF
        REG& = REG& + 1
        MSF.Rows = REG& + 1
'        VER DE GUARDAR EL CODIGO ELEMENTO EN EL INFOFAB PARA LUEGO LEERLO
        DESOP$ = SAFETEXT$(!descoper)
        CIELEM% = XVALO(!codielem)
        CONDI$ = "Codintconj = " & CIFAB%
        CONDI$ = CONDI$ + " and codintelem = " & CIELEM%
        If DESOP$ <> "" Then
           CONDI$ = CONDI$ + " and LEFT(COMPONENTE"
           CONDI$ = CONDI$ + " ," & Len(DESOP$)
           CONDI$ = CONDI$ + ") = '" & DESOP$ & "'"
        End If
       CORTEX = XVALO(VALOBADA("Coti1", "CORTE", CONDI$, "NOMESS")): If CORTEX < 1 Then CORTEX = 1
        MSF.TextMatrix(REG&, 1) = SAFETEXT$(!Cod_Exte)
        MSF.TextMatrix(REG&, 2) = SAFETEXT$(!Descrip)
        MSF.TextMatrix(REG&, 3) = SAFETEXT$(!idenlote)
        MSF.TextMatrix(REG&, 4) = FORMATNUM$(XVALO(!saldoCantiCortes) * 1, "F10.1")
        MSF.TextMatrix(REG&, 5) = CI1%
        MSF.TextMatrix(REG&, 6) = XVALO(!NUMEOPER)
        MSF.TextMatrix(REG&, 7) = SAFETEXT$(!descoper)
        MSF.TextMatrix(REG&, 8) = ""
        MSF.TextMatrix(REG&, 9) = SAFETEXT$(!CodXElem)
        MSF.TextMatrix(REG&, 10) = SAFETEXT$(!COORDUBIC)
        MSF.TextMatrix(REG&, 11) = XVALO(!CANTREG) ' NUMERO DE ORDEN DE INFORME
        MSF.TextMatrix(REG&, 12) = XVALO(!IDCOORDUBIC) ' NUMERO DE CODIGO DE UBICACION
        MSF.TextMatrix(REG&, 13) = XVALO(!NUMOPERARIO) ' NUMERO DE OPERARIO
        MSF.TextMatrix(REG&, 14) = "" ' NOMBRE DE OPERARIO
        MSF.TextMatrix(REG&, 15) = TRIM$(FORMATNUM$(XVALO(!PESOKG), "F12.2")) ' PESO INFORMADO
        MSF.TextMatrix(REG&, 16) = 0
80      .MoveNext
      Loop
    End With
    RS.Close
    Set RS = Nothing
    
End Sub

Sub NuevaCargaOperacionesInformadasSegunUbicaciones(NORFA$, CIFAB%, MSF As msFlexGrid, Optional MAYOROIGUAL&, Optional DESCRIPCIONOPERACION$, Optional MENOROIGUAL&, Optional UBICACIONESRACK$, Optional EXCEPCIONES$)
    'SI SE INGRESAN EXECPECIONES DEBEN TENER EL MISMO ORDEN QUE LOS RACK (UBICACIONESRACK$)
    'EJEMPLO UBICACIONESRACK$= "2;3;4;5" SI LAS EXCEPCIONES SON SOLO PARA EL NUMERO 3
    'EXCEPCIONES$ = 0;1,2;0;0 ES DECIR EL RACK 2 NO TIENE, EL 3 TIENE EL PISO 1 Y 2 EXCEPTUADOS
    'NOTESE QUE LOS RACK SE SEPARAN POR ';' Y DENTRO DE ESTE  SEPARADOS POR ',' CADA PISO
    'ENTONCES VA A MIRAR LAS UBICACIONES UBICACIONESRACK$= "2;3;4;5"
    'CON EXCEPCION DE (SEGUN EJEMPLO: "0;1,2;0;0") (EL RACK 2 LO MUESTRA COMPLETO (ESTA EN 0),EL 3 EXCEPTUA EL PISO 1 Y 2 ,
    'Y MUESTRA COMPLETOS EL 4 Y EL 5 X Q' ESTAN EN 0
    '
    'ESTA ES LA NUEVA DONDE SE COMIENZAN A REALIZAR MOVIMIENTOS DE SALIDA
    'ESTO ES A PARTIR DE LA OF 2341
    
    MSF.Rows = 1
    Dim CTP As New Conteplast
    CANTI = CantiProyectadaOF2(NORFA$)
    '***
    '/*/*/ EN LA PRIMERA GRILLA SE MUESTRA SIN AGRUPAR TODO LOS CORTES INFORMADOS QUE TODAVIA NO PASARON A LINEA.
    '***
    
    SQLX$ = "SELECT  CantReal "
    SQLX$ = SQLX$ + " ,descrip,cod_inte,idenlote, NumeOper,descoper,codxelem"
    SQLX$ = SQLX$ + " ,Cod_Exte,COORDUBIC,IDCOORDUBIC,CODIELEM,INFOFAB_ID,NuOrInf"
    SQLX$ = SQLX$ + " ,CantReal,NUMOPERARIO,TiemEjec,NUMOPERARIO,NameOperario, Peso"
    SQLX$ = SQLX$ + " FROM INFOFAB WITH(NOLOCK)"
    SQLX$ = SQLX$ + " WHERE    IDSUBOR = '" & NORFA$ & "' AND Cod_Inte = " & CIFAB%
    SQLX$ = SQLX$ + " AND CANTREAL > 0 "
    If XVALO(MENOROIGUAL&) > 0 Then SQLX$ = SQLX$ + " AND (IDCOORDUBIC <= " & MENOROIGUAL& & " or IDCOORDUBIC is null  OR COORDUBIC = 'SEMITERMINADOS'"
    If XVALO(MAYOROIGUAL&) > 0 Then SQLX$ = SQLX$ + " AND IDCOORDUBIC >= " & MAYOROIGUAL&
    If TRIM$(DESCRIPCIONOPERACION$) <> "" Then SQLX$ = SQLX$ + " AND DESCOPER = '" & DESCRIPCIONOPERACION$ & "'"
    
    If TRIM$(UBICACIONESRACK$) <> "" Then '
       SQLX$ = SQLX$ + " OR COORDUBIC like  '%;%'"
       Dim strArray() As String
       On Error Resume Next
       strArray = Split(UBICACIONESRACK$, ";")
       CANTVECTOR = UBound(strArray)
       '
       Dim strArray2() As String
       strArray2 = Split(EXCEPCIONES$, ";")
       CANTVECTOR2 = UBound(strArray2)
       
       Dim strArray3() As String
       
       SQLXPISO$ = ""
       For i% = 0 To CANTVECTOR
         NUMERACK$ = TRIM$(strArray(i%))
         LENRACK% = Len(NUMERACK$)

         If i% = 0 Then
           SQLX$ = SQLX$ + " and (left(COORDUBIC," & LENRACK% & ") = '" & NUMERACK$ & "'"
         Else
            SQLX$ = SQLX$ + " or left(COORDUBIC," & LENRACK% & ") = '" & NUMERACK$ & "'"
         End If
       Next i%
       SQLX$ = SQLX$ + "))"
       '
       'EXCEPCIONES DE PISOS
       For I2% = 0 To CANTVECTOR2
            PISOEXCEPCION$ = TRIM$(strArray2(I2%))
            NUMERACK$ = (strArray(I2%))
            If XVALO(PISOEXCEPCION$) > 0 Then
                strArray3 = Split(PISOEXCEPCION$, ",")
                CANTVECTOR3 = UBound(strArray3)
                For i3% = 0 To CANTVECTOR3
                   PISOEXCEPCIONINDIV$ = (strArray3(i3%))
                   RackyPisoExcepcion$ = NUMERACK$ & ";" & PISOEXCEPCIONINDIV$
                   LENEXEPCION% = Len(RackyPisoExcepcion$)
                   SQLXPISO$ = SQLXPISO$ + " and left(COORDUBIC," & LENEXEPCION% & ") <> '" & RackyPisoExcepcion$ & "'"
                Next i3%
             End If
       Next I2%
      SQLX$ = SQLX$ + SQLXPISO$
    End If
         
    On Error GoTo 0
       
'    SQLX$ = SQLX$ + " group by Cod_Exte,idenlote,NumeOper,COORDUBIC,IDCOORDUBIC,idenlote,descrip,cod_inte,idenlote"
'    SQLX$ = SQLX$ + " ,CODIELEM  ,descoper , codxelem,NuOrInf"
    SQLX$ = SQLX$ + " order by INFOFAB_ID,COORDUBIC, NUMEOPER,Cod_inte,IDENLOTE"



    Set RS = READSET(SQLX$)
    With RS
      Do While Not .EOF
        REG& = REG& + 1
        MSF.Rows = REG& + 1
'        VER DE GUARDAR EL CODIGO ELEMENTO EN EL INFOFAB PARA LUEGO LEERLO
        DESOP$ = SAFETEXT$(!descoper)
        CIELEM% = XVALO(!codielem)
        CONDI$ = "Codintconj = " & CIFAB%
        CONDI$ = CONDI$ + " and codintelem = " & CIELEM%
        If DESOP$ <> "" Then
           CONDI$ = CONDI$ + " and LEFT(COMPONENTE"
           CONDI$ = CONDI$ + " ," & Len(DESOP$)
           CONDI$ = CONDI$ + ") = '" & DESOP$ & "'"
        End If
        CORTEX = XVALO(VALOBADA("Coti1", "CORTE", CONDI$, "NOMESS")): If CORTEX < 1 Then CORTEX = 1
'        IDINFOFAB = XVALO(!INFOFAB_ID)
        MSF.TextMatrix(REG&, 1) = SAFETEXT$(!Cod_Exte)
        MSF.TextMatrix(REG&, 2) = SAFETEXT$(!Descrip)
        MSF.TextMatrix(REG&, 3) = SAFETEXT$(!idenlote)
        MSF.TextMatrix(REG&, 4) = FORMATNUM$(XVALO(!CANTREAL) * 1, "F10.1")
        MSF.TextMatrix(REG&, 5) = CI1%
        MSF.TextMatrix(REG&, 6) = XVALO(!NUMEOPER)
        MSF.TextMatrix(REG&, 7) = SAFETEXT$(!descoper)
        MSF.TextMatrix(REG&, 8) = SAFETEXT$(!TiemEjec)
        MSF.TextMatrix(REG&, 9) = SAFETEXT$(!CodXElem)
        MSF.TextMatrix(REG&, 10) = SAFETEXT$(!COORDUBIC)
        MSF.TextMatrix(REG&, 11) = XVALO(!NuOrInf) ' NUMERO DE ORDEN DE INFORME
        MSF.TextMatrix(REG&, 12) = XVALO(!IDCOORDUBIC) ' NUMERO DE CODIGO DE UBICACION
        MSF.TextMatrix(REG&, 13) = XVALO(!NUMOPERARIO) ' NUMERO DE OPERARIO
        MSF.TextMatrix(REG&, 14) = SAFETEXT(!NameOperario) ' NOMBRE DE OPERARIO
        MSF.TextMatrix(REG&, 15) = TRIM$(FORMATNUM$(XVALO(!Peso), "F12.2")) ' PESO INFORMADO
        MSF.TextMatrix(REG&, 16) = XVALO(!INFOFAB_ID) '
80      .MoveNext
      Loop
    End With
    RS.Close
    Set RS = Nothing
    
End Sub

Sub CargaOperacionesInformadasSegunUbicaciones(NORFA$, CIFAB%, MSF As msFlexGrid, Optional MAYOROIGUAL&, Optional DESCRIPCIONOPERACION$, Optional MENOROIGUAL&, Optional UBICACIONESRACK$, Optional EXCEPCIONES$)
    'SI SE INGRESAN EXECPECIONES DEBEN TENER EL MISMO ORDEN QUE LOS RACK (UBICACIONESRACK$)
    'EJEMPLO UBICACIONESRACK$= "2;3;4;5" SI LAS EXCEPCIONES SON SOLO PARA EL NUMERO 3
    'EXCEPCIONES$ = 0;1,2;0;0 ES DECIR EL RACK 2 NO TIENE, EL 3 TIENE EL PISO 1 Y 2 EXCEPTUADOS
    'NOTESE QUE LOS RACK SE SEPARAN POR ';' Y DENTRO DE ESTE  SEPARADOS POR ',' CADA PISO
    'ENTONCES VA A MIRAR LAS UBICACIONES UBICACIONESRACK$= "2;3;4;5"
    'CON EXCEPCION DE (SEGUN EJEMPLO: "0;1,2;0;0") (EL RACK 2 LO MUESTRA COMPLETO (ESTA EN 0),EL 3 EXCEPTUA EL PISO 1 Y 2 ,
    'Y MUESTRA COMPLETOS EL 4 Y EL 5 X Q' ESTAN EN 0
    MSF.Rows = 1
    Dim CTP As New Conteplast
    CANTI = CantiProyectadaOF2(NORFA$)
    'DEVUELVE LAS OPERACIONES INFORAMDAS SEGUN OF. FILTRADAS POR TAMBIEN UBICACIONES EN RACK
    SQLX$ = " SELECT * FROM INFOFAB I WITH(NOLOCK)"
    SQLX$ = SQLX$ + " WHERE NOT EXISTS"
    SQLX$ = SQLX$ + " (SELECT * FROM INFOFAB AS I2"
    SQLX$ = SQLX$ + " WHERE I.INFOFAB_ID = I2.IDREGPADRE)"
    SQLX$ = SQLX$ + " AND IDSUBOR = '" & NORFA$ & "'"
    If XVALO(MENOROIGUAL&) > 0 Then SQLX$ = SQLX$ + " AND (IDCOORDUBIC <= " & MENOROIGUAL& & " or IDCOORDUBIC is null  OR COORDUBIC = 'SEMITERMINADOS')"
    If XVALO(MAYOROIGUAL&) > 0 Then SQLX$ = SQLX$ + " AND IDCOORDUBIC >= " & MAYOROIGUAL&
    If TRIM$(DESCRIPCIONOPERACION$) <> "" Then SQLX$ = SQLX$ + " AND DESCOPER = '" & DESCRIPCIONOPERACION$ & "'"
    
    If TRIM$(UBICACIONESRACK$) <> "" Then '
       SQLX$ = SQLX$ + " and COORDUBIC like  '%;%'"
       Dim strArray() As String
       On Error Resume Next
       strArray = Split(UBICACIONESRACK$, ";")
       CANTVECTOR = UBound(strArray)
       '
       Dim strArray2() As String
       strArray2 = Split(EXCEPCIONES$, ";")
       CANTVECTOR2 = UBound(strArray2)
       
       Dim strArray3() As String
       
       SQLXPISO$ = ""
       For i% = 0 To CANTVECTOR
         NUMERACK$ = TRIM$(strArray(i%))
         LENRACK% = Len(NUMERACK$)

         If i% = 0 Then
           SQLX$ = SQLX$ + " and (left(COORDUBIC," & LENRACK% & ") = '" & NUMERACK$ & "'"
         Else
            SQLX$ = SQLX$ + " or left(COORDUBIC," & LENRACK% & ") = '" & NUMERACK$ & "'"
         End If
       Next i%
       SQLX$ = SQLX$ + ")"
       '
       'EXCEPCIONES DE PISOS
       For I2% = 0 To CANTVECTOR2
            PISOEXCEPCION$ = TRIM$(strArray2(I2%))
            NUMERACK$ = (strArray(I2%))
            If XVALO(PISOEXCEPCION$) > 0 Then
                strArray3 = Split(PISOEXCEPCION$, ",")
                CANTVECTOR3 = UBound(strArray3)
                For i3% = 0 To CANTVECTOR3
                   PISOEXCEPCIONINDIV$ = (strArray3(i3%))
                   RackyPisoExcepcion$ = NUMERACK$ & ";" & PISOEXCEPCIONINDIV$
                   LENEXEPCION% = Len(RackyPisoExcepcion$)
                   SQLXPISO$ = SQLXPISO$ + " and left(COORDUBIC," & LENEXEPCION% & ") <> '" & RackyPisoExcepcion$ & "'"
                Next i3%
             End If
       Next I2%
      SQLX$ = SQLX$ + SQLXPISO$
    End If
         
    On Error GoTo 0
       
    
    SQLX$ = SQLX$ + " order by COORDUBIC, NUMEOPER,FeHorInf,Cod_inte,IDENLOTE"

    Set RS = READSET(SQLX$)
    With RS
      Do While Not .EOF
        REG& = REG& + 1
        MSF.Rows = REG& + 1
'        VER DE GUARDAR EL CODIGO ELEMENTO EN EL INFOFAB PARA LUEGO LEERLO
        DESOP$ = SAFETEXT$(!descoper)
        CIELEM% = XVALO(!codielem)
        CONDI$ = "Codintconj = " & CIFAB%
        CONDI$ = CONDI$ + " and codintelem = " & CIELEM%
        If DESOP$ <> "" Then
           CONDI$ = CONDI$ + " and LEFT(COMPONENTE"
           CONDI$ = CONDI$ + " ," & Len(DESOP$)
           CONDI$ = CONDI$ + ") = '" & DESOP$ & "'"
        End If
       CORTEX = XVALO(VALOBADA("Coti1", "CORTE", CONDI$, "NOMESS")): If CORTEX < 1 Then CORTEX = 1
        MSF.TextMatrix(REG&, 1) = SAFETEXT$(!Cod_Exte)
        MSF.TextMatrix(REG&, 2) = SAFETEXT$(!Descrip)
        MSF.TextMatrix(REG&, 3) = SAFETEXT$(!idenlote)
        MSF.TextMatrix(REG&, 4) = FORMATNUM$(XVALO(!CANTREAL) * 1, "F10.1")
        MSF.TextMatrix(REG&, 5) = CI1%
        MSF.TextMatrix(REG&, 6) = XVALO(!NUMEOPER)
        MSF.TextMatrix(REG&, 7) = SAFETEXT$(!descoper)
        MSF.TextMatrix(REG&, 8) = SAFETEXT$(!TiemEjec)
        MSF.TextMatrix(REG&, 9) = SAFETEXT$(!CodXElem)
        MSF.TextMatrix(REG&, 10) = SAFETEXT$(!COORDUBIC)
        MSF.TextMatrix(REG&, 11) = XVALO(!NuOrInf) ' NUMERO DE ORDEN DE INFORME
        MSF.TextMatrix(REG&, 12) = XVALO(!IDCOORDUBIC) ' NUMERO DE CODIGO DE UBICACION
        MSF.TextMatrix(REG&, 13) = XVALO(!NUMOPERARIO) ' NUMERO DE OPERARIO
        MSF.TextMatrix(REG&, 14) = SAFETEXT(!NameOperario) ' NOMBRE DE OPERARIO
        MSF.TextMatrix(REG&, 15) = TRIM$(FORMATNUM$(XVALO(!Peso), "F12.2")) ' PESO INFORMADO
        MSF.TextMatrix(REG&, 16) = XVALO(!INFOFAB_ID) '
80      .MoveNext
      Loop
    End With
    RS.Close
    Set RS = Nothing
    
End Sub


Function DescriOperSegunCodigo(CI1%, NOPER%)
     'OJO NO CONSIDERA CONDIGO
     CONDI1$ = "CodIConj = " & CI1% & " and  NumeOper = " & NOPER%
     DescriOperSegunCodigo = VALOBADA("SECOPER", "DescOper", CONDI1$)
End Function
Function IDSEGUNUBICACION&(TEXTOUBICACION$)
    CONDI$ = "LISTA = 'SECTPLANTA' AND CAMPO2 = '" & TEXTOUBICACION$ & "'"
    IDSEGUNUBICACION& = XVALO(VALOBADA("DATASQL", "CAMPO1", CONDI$, "NOMESS"))
End Function

Function UBICACIONSEGUNID$(ID&)
    CONDI$ = "LISTA = 'SECTPLANTA' AND CONVERT(INT,CAMPO1) = " & ID&
    UBICACIONSEGUNID$ = VALOBADA("DATASQL", "CAMPO2", CONDI$, "NOMESS")
End Function

Function ValidarMatDeConsumoTransferidoALinea%(CIFAB%, NROCOTIZA&, CantiaCerrar, NORFA$)
        SQLX$ = "SELECT  C.CODIGO, C.DESCRIPCION, C.COMPONENTE,C.CODINTELEM, C.MODUXBB,C.CORTE"
        SQLX$ = SQLX$ + " , SUM(C.MODUXBB * C.CORTE * 1) AS NECESIDAD_EN_M"
        SQLX$ = SQLX$ + " ,SUM(C.MODUXBB * 1) AS NECESIDAD_DE_CORTES "
        SQLX$ = SQLX$ + " FROM COTI1 C  WITH(NOLOCK)"
        SQLX$ = SQLX$ + " Where CODINTCONJ = " & CIFAB%
        SQLX$ = SQLX$ + " And CODINTELEM > 0 And NROCOTIZACION =  " & NROCOTIZA&
        SQLX$ = SQLX$ + " And LEFT(COMPONENTE,8) <> 'CINTILLA' And LEFT(COMPONENTE,4) <> 'HILO'" 'EXCEPCIONES
        SQLX$ = SQLX$ + " GROUP BY CODIGO, C.DESCRIPCION, C.COMPONENTE,C.CODINTELEM, C.MODUXBB,C.CORTE"
        Set RS = READSET(SQLX$)
        With RS
            Do While Not .EOF
               DESCIPCIONOPERACION$ = SAFETEXT(!COMPONENTE)
               rr = SaldoEnLineaCortesInfofab(NORFA$, DESCIPCIONOPERACION$)
               If SaldoEnLineaCortesInfofab(NORFA$, DESCIPCIONOPERACION$) < CantiaCerrar Then
                  ValidarMatDeConsumoTransferidoALinea% = (-1)
                  XXX$ = "O.F.: " & NORFA$
                  XXX$ = XXX$ & " - Cierre x " & CantiaCerrar
                  OPX% = COMALTER%(XXX$ & "\ Faltan Transferir a Linea Material: " & DESCIPCIONOPERACION$ & "\\Cancelar\Continuar")
                  If OPX% < 2 Then
                    Exit Do
                  End If
               End If
               .MoveNext
            Loop
        End With
End Function

Function Zelecho_OperacionesSegunCodiConjunto(CI1%)
   Dim CTP As New Conteplast
   OperacionesDelCodigoLeido$ = ""
   Call CARGA_ENCABEZADO(FRMZELECHO.msf2, "Operación/A30;Código./A18;Descripción/A32;Formulario/A12;Necesidad;U.Med./A6; CI/f6", FRMZELECHO)
   Call CARGA_ENCABEZADO(FRMZELECHO.MSF1, "Operación/A30;Código./A18;Descripción/A32;Formulario/A12;Necesidad;U.Med./A6; CI/f6", FRMZELECHO)
   
   REG& = 0
   SQLX$ = "SELECT * FROM FORMULAS WITH(NOLOCK) WHERE CODICONJ =" & CI1%
   SQLX$ = SQLX$ + " ORDER BY COMPONENTE"
   Set RS = READSET(SQLX$)
   With RS
     Do While Not .EOF
         REG& = REG& + 1
         CIELEM% = XVALO(!codielem)
         FRMZELECHO.msf2.Rows = REG& + 1
         NOMBRESECTOR$ = SAFETEXT$(!COMPONENTE)
         FRMZELECHO.msf2.TextMatrix(REG&, 1) = NOMBRESECTOR$
         FRMZELECHO.msf2.TextMatrix(REG&, 2) = SAFETEXT$(!CodXElem)
         FRMZELECHO.msf2.TextMatrix(REG&, 3) = SAFETEXT$(DESCRIT0$(CIELEM%))
         FRMZELECHO.msf2.TextMatrix(REG&, 4) = CTP.FRMSECTOR(NOMBRESECTOR$)
         FRMZELECHO.msf2.TextMatrix(REG&, 5) = XVALO(!usobruto) * XVALO(TXTCANPROY)
         FRMZELECHO.msf2.TextMatrix(REG&, 6) = SAFETEXT$(!UMED_FORM)
         FRMZELECHO.msf2.TextMatrix(REG&, 7) = CIELEM%
         'SI LA UNIDAD DE CONSUMO ES
'         If TRIM$(UCase$(SAFETEXT$(!UMED_FORM))) = "CM" Then FRMZELECHO.msf2.TextMatrix(REG&, 6) = XVALO(!USOBRUTO) / 100
         
       .MoveNext
     Loop
   End With
   RS.Close
   Set RS = Nothing
   FRMZELECHO.Show 1
   OperacionesDelCodigoLeido$ = RESP_ZELE_VECT(1)

End Function
Function PesoLeidoBalanza#(CI1%, IDTARA%, Optional OK%, Optional NoValidarCodigo%, Optional TARAEDITABLE#)
    PesoLeidoBalanza# = 0
    OK% = 0
    If NoValidarCodigo% > 0 Then 'OPCION PARA CONTEMPLAR SI SOLO PESO SIN MOSTRAR CODIGO
      BALANZA.TXTDESCRIPCION = ""
    Else
      BALANZA.TXTDESCRIPCION = DESCRIT0$(CI1%)
    End If
    
    On Error Resume Next
    BALANZA.Show 1
    If Err.Number > 0 Then
        MsgBox "Atención Falta Formulario Balanza"
    End If
    On Error GoTo 0

    If BALANZA.Respuesta = "APROBO" Then
        OK% = 1
        PesoLeidoBalanza# = XVALO(BALANZA.TXTPESO)
        If IDTARA% > 0 And TARAEDITABLE# <= 0 Then
            PesoLeidoBalanza# = PesoLeidoBalanza# - TaraSegunIDTabla#(IDTARA%)
        Else
            PesoLeidoBalanza# = PesoLeidoBalanza# - XVALO(TARAEDITABLE#)
        End If
        Unload BALANZA
    End If
    
End Function
Function OperacionesDelCodigoLeidoGeneral$(CIELEM%, NORFA$, Optional ret_CantRegistros&)
   CI1% = CODIFAB2%(NORFA$)
   Dim CTP As New Conteplast
   OperacionesDelCodigoLeidoGeneral$ = ""
   Call CARGA_ENCABEZADO(FRMZELECHO.msf2, "Operación/A30;Código./A18;Descripción/A32;Formulario/A12;Necesidad;U.Med./A6; CI/f6;Nop/F4", FRMZELECHO)
   Call CARGA_ENCABEZADO(FRMZELECHO.MSF1, "Operación/A30;Código./A18;Descripción/A32;Formulario/A12;Necesidad;U.Med./A6; CI/f6;Nop/F4", FRMZELECHO)
   REG& = 0
   SQLX$ = "SELECT * FROM FORMULAS WITH(NOLOCK) WHERE CODICONJ =" & CI1%
   SQLX$ = SQLX$ + " AND (CODIELEM = " & CIELEM%
   CIRANGO& = CODIGORANGO&(CIELEM%)
   If CIRANGO& > 0 Then
      SQLX$ = SQLX$ + " OR CODIELEM = " & CIRANGO&
   End If
   SQLX$ = SQLX$ + " )"
   SQLX$ = SQLX$ + " ORDER BY NUORIT"
   Set RS = READSET(SQLX$)
   With RS
     Do While Not .EOF
         REG& = REG& + 1
         CIELEM% = XVALO(!codielem)
         FRMZELECHO.msf2.Rows = REG& + 1
         NOMBRESECTOR$ = SAFETEXT$(!COMPONENTE)
         NOP% = CTP.NumeOperXNombre(CI1%, NOMBRESECTOR$)
         FRMZELECHO.msf2.TextMatrix(REG&, 1) = NOMBRESECTOR$
         FRMZELECHO.msf2.TextMatrix(REG&, 2) = SAFETEXT$(!CodXElem)
         FRMZELECHO.msf2.TextMatrix(REG&, 3) = SAFETEXT$(DESCRIT0$(CIELEM%))
         FRMZELECHO.msf2.TextMatrix(REG&, 4) = CTP.FRMSECTOR(NOMBRESECTOR$)
         FRMZELECHO.msf2.TextMatrix(REG&, 5) = XVALO(!usobruto) * XVALO(TXTCANPROY)
         FRMZELECHO.msf2.TextMatrix(REG&, 6) = SAFETEXT$(!UMED_FORM)
         FRMZELECHO.msf2.TextMatrix(REG&, 7) = CIELEM%
         FRMZELECHO.msf2.TextMatrix(REG&, 8) = NOP%
         
         'SI LA UNIDAD DE CONSUMO ES
'         If TRIM$(UCase$(SAFETEXT$(!UMED_FORM))) = "CM" Then FRMZELECHO.msf2.TextMatrix(REG&, 6) = XVALO(!USOBRUTO) / 100
         
       .MoveNext
     Loop
   End With
   RS.Close
   Set RS = Nothing
   If REG& > 1 Then 'si son mas de un registro muestra la lista
     FRMZELECHO.Show 1
     OperacionesDelCodigoLeidoGeneral$ = RESP_ZELE_VECT(1)
'     TXTNUMEOPER = RESP_ZELE_VECT(8)
'     Me.TXTNECESIDAD = RESP_ZELE_VECT(5)
     Exit Function
   End If
   'si es un solo registro llega aqui y asigna directamente
   OperacionesDelCodigoLeidoGeneral$ = NOMBRESECTOR$
'   Me.TXTNECESIDAD = FRMZELECHO.MSF2.TextMatrix(REG&, 5)
   ret_CantRegistros& = REG&
   'EL CONSUMO TOTAL DEBERIA SER EL CORTE POR MODXBB

End Function


Function InformesDeScrap&(Des%, Has%, PERIO$)
    Campos$ = "Fecha/D8;" '1
    Campos$ = Campos$ + "C.Sect./F6;" '2
    Campos$ = Campos$ + "Descripción/A32;"      '3
    Campos$ = Campos$ + "C.Categ./F7;"      '4
    Campos$ = Campos$ + "Descripción/A32;"   '5
    Campos$ = Campos$ + "Cód.Art./A10;"       '6
    Campos$ = Campos$ + "Descripción/A32;"        '7
    Campos$ = Campos$ + "N°Dep./F6;"     '8
    Campos$ = Campos$ + "CI/A0;" '9
    Campos$ = Campos$ + "ID/A0;" '10
    Campos$ = Campos$ + "KG./F8.2" '11

    Call CARGA_ENCABEZADO(FRMZELECHO.msf2, Campos$, FRMZELECHO)
    Call CARGA_ENCABEZADO(FRMZELECHO.MSF1, Campos$, FRMZELECHO)
    FRMZELECHO.Caption = "Informes de Scrap "
    If PERIO$ <> "" Then
       FRMZELECHO.Caption = "Informes de Scrap (" & PERIO$ & ")"
    End If
    
    Dim obj As New RECORXET
    Set rs1 = obj.RS_InformesScrap(Des%, Has%)
    Call Carga_MSF_Recordset(rs1, Campos$, FRMZELECHO.msf2, 1)
    Set obj = Nothing
    rs1.Close
    Set rs1 = Nothing
    FRMZELECHO.Width = 12500
    FRMZELECHO.Show 1
    InformesDeScrap& = XVALO(RESP_ZELE_VECT(10))
End Function
Function InformesDeScrapReporte(Des%, Has%, PERIO$)
    Campos$ = "Fecha/D8;" '1
    Campos$ = Campos$ + "C.Sector/F10;" '2
    Campos$ = Campos$ + "Descripción/A32;"      '3
    Campos$ = Campos$ + "C.Cat./F5;"      '4
    Campos$ = Campos$ + "Descripción/A32;"   '5
    Campos$ = Campos$ + "Cód./A5;"       '6
    Campos$ = Campos$ + "Descripción/A32;"        '7
    Campos$ = Campos$ + "N°Dep./F6;"     '8
    Campos$ = Campos$ + "CI/A0;" '9
    Campos$ = Campos$ + "ID/A0;" '10
    Campos$ = Campos$ + "KG./F9.2" '11

    Call CARGA_ENCABEZADO(GRILLAGRAL.GRID, Campos$, GRILLAGRAL, 1)
    Call CARGA_ENCABEZADO(GRILLAGRAL.MSF_2, Campos$, GRILLAGRAL, 1)
    TX$ = "Informes de Scrap"
    If PERIO$ <> "" Then
        TX$ = "Informes de Scrap (" & PERIO$ & ")"
    End If
    GRILLAGRAL.Caption = TX$

    Dim obj As New RECORXET
    Set rs1 = obj.RS_InformesScrap(Des%, Has%)
    Call Carga_MSF_Recordset(rs1, Campos$, GRILLAGRAL.GRID)
    Set obj = Nothing
    rs1.Close
    Set rs1 = Nothing
    GRILLAGRAL.Width = 12500
    
    '*******************************************
    GRILLAGRAL.mnuMenu1.Caption = "Imprimir"
    GRILLAGRAL.mnuMenu1.Visible = True
    GRILLAGRAL.GRID.ZOrder 0
    GRILLAGRAL.TXTBUSCAR = ""
    
    GRILLAGRAL.SinEfectoClickEnGrilla% = 1

    
    GRILLAGRAL.Show 1
    Dim MSF As msFlexGrid
    If GRILLAGRAL.MENU1 > 0 Then
        Set MSF = GRILLAGRAL.GRID
        If GRILLAGRAL.MSF_2.Rows > 1 Then Set MSF = GRILLAGRAL.MSF_2
        Call CARGA_TABLA_IMPRE(MSF, Campos$, "INFOSCRP", TX$, "TABLA_IMPRE", RET_NAMECONS$, 1)
        Call FINAL("?" & RET_NAMECONS$)
        GRILLAGRAL.MENU1 = 0
        Screen.MousePointer = 1
    End If

End Function
Sub AnularInformeScrap(idinfScrap&)
   If idinfScrap& < 1 Then Exit Sub
   
   'borra el informe de scrap segun el id y tambien borra el que se genero en el movisto

   '
   'BORRO DE LA TABLA DE DETALLE EL MOVIMIENTO SEGUN EL ID
   SQLX$ = "DELETE FROM DETASCRAP WHERE DETASCRAP_ID = " & idinfScrap&
   FLEXCONN.Execute (SQLX$)
   
   'BORRO DE LA TABLA MOVISTO SEGUN EL ID DE LA TABLA DETASCRAP CON EL FORMATO SC.nnn por que cuando se guarda
   'le graba en Doseccor y Doseclar con SC. seguido del numero de id de DETASCRAP
   SQLX$ = "DELETE FROM MOVISTO WHERE CODIMOVI = 'SC' AND DoSecCor = '" & "SC." & idinfScrap& & "'"
   FLEXCONN.Execute (SQLX$)
   

End Sub
Sub CARGARARCHIVOSENCOMBO(FTEC&, COMBO As comboBox)
      COMBO.Clear
      SQLX$ = "SELECT * FROM DOCURELFT WITH(NOLOCK) WHERE NUMRECETACTP  = " & FTEC&
      Dim DOCTMP As ADODB.Recordset
      Set DOCTMP = FLEXCONN.Execute(FILTSQL$(SQLX$))
      With DOCTMP
        Do While Not .EOF
           DESCRI$ = SAFETEXT$(!Descripcion)
           RUTAX$ = SAFETEXT$(!Ruta$)
           Call REPLA(TX$, DESCRI$, 1, 24)
           Call REPLA(TX$, RUTAX$, 32, 512)
           COMBO.AddItem TX$
           .MoveNext
        Loop
      End With
End Sub
 Function SELEC_FTECNICAYCOTI(Optional CONDIPARAM$) As Long
    Erase RESP_ZELE_VECT

    Campos$ = "N°F.Tec./F9;Fecha/D8;N°Cotiz./F9;Fecha/D8;Cta./F8;Razon Social/A32;Stat/F5;Código/A20;FT_id/a0;COT_ID/A0;Nro.Vend./F10"
    Call CARGA_ENCABEZADO(FRMZELECHO.msf2, Campos$, FRMZELECHO)
    Call CARGA_ENCABEZADO(FRMZELECHO.MSF1, Campos$, FRMZELECHO)
    FRMZELECHO.Caption = ""
    
    Dim obj As New RECORXET
    CAMPOSTABLA$ = "FT_NROFT,FT_FECHEMISI, COT_NROCOT,COT_FECHEMISI, FT_CTACLTE,FT_RASOCLI,FT_STAT,FT_CODIGO,FT_ID,COT_ID,FT_NUVEND"
    TABLA$ = "FITECYCOTIZA" 'ES UNA VISTA
    CONDI$ = "FT_NROFT > 0 "
    If TRIM$(CONDIPARAM$) <> "" Then CONDI$ = CONDI$ + CONDIPARAM$
    
    CONDI$ = CONDI$ + " ORDER BY COT_NROCOT DESC"
    Set rs1 = obj.RS_GENERAL(TABLA$, CAMPOSTABLA$, CONDI$)
    Call Carga_MSF_Recordset(rs1, Campos$, FRMZELECHO.msf2, 1)
    On Error Resume Next
    Set obj = Nothing
    rs1.Close
    Set rs1 = Nothing
    On Error GoTo 0
        
    FRMZELECHO.Width = 12000
50  FRMZELECHO.Show 1
    If XVALO(RESP_ZELE_VECT(1)) > 0 Then '
       SELEC_FTECNICAYCOTI = XVALO(RESP_ZELE_VECT(1))
    End If
End Function
Sub TransfIngresada(MSF As msFlexGrid, DepoIngeso%, Optional FILTRO$)

  MSF.Rows = 1: REG& = 0
  DEPOCORTE% = XVALO(Control("", "DEPODEST"))
  SQLX$ = "SELECT  COD_INTE,COD_EXTE,DESCRIP, IDENLOTE,DOPRILAR"
  SQLX$ = SQLX$ + " FROM MOVISTO WITH(NOLOCK)"
  SQLX$ = SQLX$ + " where DEPOMOVI = " & DepoIngeso%
  SQLX$ = SQLX$ + " and cantingre > 0"
  If TRIM$(FILTRO$) <> "" Then SQLX$ = SQLX$ + " and " & FILTRO$
  
  SQLX$ = SQLX$ + " GROUP BY COD_INTE,COD_EXTE,DESCRIP, IDENLOTE,DOPRILAR"
  SQLX$ = SQLX$ + " ORDER BY IDENLOTE "
  Set RS = READSET(SQLX$)
  With RS
    Do While Not .EOF
       LOTE$ = SAFETEXT$(!idenlote)
       CONDI$ = "DEPOMOVI = " & DepoIngeso% & " AND IDENLOTE = '" & LOTE$ & "'"
'      PesoPorMetro# = PESMETRO(CI1%)
       CANTMETROSBOBINAING = XVALO(VALOBADA("MOVISTO", "Sum(CANTINGRE)", CONDI$, "NOMESS"))
       CANTMETROSBOBINASAL = XVALO(VALOBADA("MOVISTO", "Sum(CANTSALID)", CONDI$, "NOMESS"))
       CANTMETROSBOBINA = XVALO(CANTMETROSBOBINAING) - (XVALO(CANTMETROSBOBINASAL) + 0.5)
       If CANTMETROSBOBINA > 0 Then
            CONDI$ = "DEPOMOVI = " & DepoIngeso% & " AND IDENLOTE = '" & LOTE$ & "' AND CANTINGRE > 0"
            CANTKILOSBOBINAMAS = XVALO(VALOBADA("MOVISTO", "SUM(PESOKG)", CONDI$, "NOMESS"))
            CONDI$ = "DEPOMOVI = " & DepoIngeso% & " AND IDENLOTE = '" & LOTE$ & "' AND CANTSALID > 0"
            CANTKILOSBOBINAMENOS = XVALO(VALOBADA("MOVISTO", "SUM(PESOKG)", CONDI$, "NOMESS"))
            PESOENKG# = CANTKILOSBOBINAMAS - CANTKILOSBOBINAMENOS
            REG& = REG& + 1
            MSF.Rows = REG& + 1
            CI1% = XVALO(!cod_inte)
            MSF.TextMatrix(REG&, 1) = SAFETEXT$(!Cod_Exte)
            MSF.TextMatrix(REG&, 2) = SAFETEXT$(!Descrip)
            PesoPorMetro# = CoeficientePesoPorMetroSegunLote#(LOTE$)
            PESOSALDOBOBINA = CANTMETROSBOBINA * PesoPorMetro#
            MSF.TextMatrix(REG&, 3) = LOTE$
            MSF.TextMatrix(REG&, 4) = FORMATNUM$(XVALO(PESOENKG#), "F10.2")
            MSF.TextMatrix(REG&, 5) = FORMATNUM$(CANTMETROSBOBINA, "F10.1")
            MSF.TextMatrix(REG&, 6) = XVALO(!cod_inte)
            MSF.TextMatrix(REG&, 7) = SAFETEXT$(!DOPRILAR)
       End If
      
       .MoveNext
    Loop
 End With
 RS.Close
End Sub
Sub LinerIngresadaEnDeposito(MSF As msFlexGrid, DepoIngeso%, familia$)

  MSF.Rows = 1: REG& = 0
  SQLX$ = "SELECT  SUM(M.cantingre - M.Cantsalid),M.COD_INTE,M.COD_EXTE,M.DESCRIP, M.IDENLOTE,MA.FAMILIA,M.Cantsalid"
  SQLX$ = SQLX$ + " FROM MOVISTO M WITH(NOLOCK)"
  SQLX$ = SQLX$ + " INNER JOIN MASTER MA ON MA.codint =  M.COD_INTE"
  SQLX$ = SQLX$ + " where M.DEPOMOVI = " & DepoIngeso%
  SQLX$ = SQLX$ + " and   MA.FAMILIA = '" & familia$ & "'"
  SQLX$ = SQLX$ + " GROUP BY M.COD_INTE,M.COD_EXTE,M.DESCRIP, M.IDENLOTE,MA.FAMILIA,M.Cantsalid"
  SQLX$ = SQLX$ + " HAVING SUM(M.cantingre - M.Cantsalid) > 0"
  SQLX$ = SQLX$ + " ORDER BY M.IDENLOTE "
  Set RS = READSET(SQLX$)
  With RS
    Do While Not .EOF
       LOTE$ = SAFETEXT$(!idenlote)
       CONDI$ = "DEPOMOVI = " & DepoIngeso% & " AND IDENLOTE = '" & LOTE$ & "'"
'      PesoPorMetro# = PESMETRO(CI1%)
       CANTI = XVALO(VALOBADA("MOVISTO", "Sum(CANTINGRE - CANTSALID)", CONDI$, "NOMESS"))
       If CANTI > 0 Then

            REG& = REG& + 1
            MSF.Rows = REG& + 1
            CI1% = XVALO(!cod_inte)
            MSF.TextMatrix(REG&, 1) = SAFETEXT$(!Cod_Exte)
            MSF.TextMatrix(REG&, 2) = SAFETEXT$(!Descrip)
            MSF.TextMatrix(REG&, 3) = LOTE$
            MSF.TextMatrix(REG&, 4) = FORMATNUM$(CANTI, "F10.1")
            MSF.TextMatrix(REG&, 5) = XVALO(!cod_inte)
            MSF.TextMatrix(REG&, 6) = ""
       End If
      
       .MoveNext
    Loop
 End With
 RS.Close
End Sub
Sub CargaMatTranfXOrden(MSF1 As msFlexGrid, NORFA$, DEPOCORTE%, Optional FILTROOPERACION%, Optional AGREGARCODIGORANGO%, Optional ConsiderarConsumo%, Optional CamposOrden$)
  'CARGA LAS TRANSFERENCIAS REALIZADAS POR OF AL DEPOSITO DE CORTE
  
'**** AGREGAR EL CAMPO KG EN LA CONSULTA Y VER SI TRAE LOS KG Y Q NO LOS CALCULE COMO AHORA.
  EPAC$ = UCase$(TRIM$(EMPREAC$))

  MSF1.Rows = 1: REG& = 0
  SQLX$ = "SELECT SUM(CANTINGRE-CANTSALID) AS RESULTADO ,COD_INTE,COD_EXTE,DESCRIP, IDENLOTE,ID_OPERACION,PESOKG FROM MOVISTO WITH(NOLOCK)"
  SQLX$ = SQLX$ + " where "
  If ConsiderarConsumo% > 0 Then 'SI CONSIDERA CONSUMO VA A MOSTRAR EL RESULTADO DE LO TRANSFERIDO MENOS EL CONSUMO DE LA ORDEN FAB
     SQLX$ = SQLX$ + " (DOSECLAR = '" & TRIM(NORFA$) & "' or  DOPRICOR = '" & TRIM(NORFA$) & "')"
  Else
     SQLX$ = SQLX$ + " DOSECLAR = '" & TRIM(NORFA$) & "'"
  End If
   
  SQLX$ = SQLX$ + " AND DEPOMOVI = " & DEPOCORTE%
  If FILTROOPERACION% > 0 Then SQLX$ = SQLX$ + " AND ID_OPERACION = " & FILTROOPERACION%
  
  SQLX$ = SQLX$ + " GROUP BY COD_INTE,COD_EXTE,DESCRIP, IDENLOTE,ID_OPERACION,PESOKG"
  If TRIM$(CamposOrden$) <> "" Then
     SQLX$ = SQLX$ + " ORDER BY " & CamposOrden$
  Else
     SQLX$ = SQLX$ + " ORDER BY IDENLOTE "
  End If
'  Call CARGA_ENCABEZADO(MSF2, "CODIGO/A20;Descripción/A40;LOTE/A16;Saldo.Transf.Mts./f18.1;CI/f6;Op.N°/F6;Kg.Transf.(Orig.)/F16.2", Me)
'  Call CARGA_ENCABEZADO(MSF3, "CI1/A0;Còdigo/A16;Descripciòn/A32;Lote/a12;Consumo/F10.0;Kgs./F8.0", Me)
  DEPOCORTE% = XVALO(Control("", "DEPODEST"))
  

  
  Set RS = READSET(SQLX$)
  With RS
    Do While Not .EOF
      If XVALO(!Resultado) > 0 Then
      
       REG& = REG& + 1
       MSF1.Rows = REG& + 1
       CI1% = XVALO(!cod_inte)
       MSF1.TextMatrix(REG&, 1) = SAFETEXT$(!Cod_Exte)
       MSF1.TextMatrix(REG&, 2) = SAFETEXT$(!Descrip)
       LOTE$ = SAFETEXT$(!idenlote)

       CANTMETROSBOBINA = XVALO(!Resultado)
       PesoPorMetro# = CoeficientePesoPorMetroSegunLote#(LOTE$)
       PESOSALDOBOBINA = CANTMETROSBOBINA * PesoPorMetro#
       If InStr(EPAC$, "QUILMES") > 0 Then
         MSF1.TextMatrix(REG&, 4) = STODEPFECHENMETROS#(CI1%, DEPOCORTE%, HOY(HO$), LOTE$)
         PESOSALDOBOBINA = STODEPFECHENKG#(CI1%, DEPOCORTE%, HOY(HO$), LOTE$)
       End If
'       If Unimed(CI1%) = "U." Then
'         CANTMETROSBOBINA = SaldoLotexMetros(CI1%, LOTE$, DEPOCORTE%)
'         PESOSALDOBOBINA = SaldoLotexKilos(CI1%, LOTE$, DEPOCORTE%)
'       Else
'         PesoPorMetro# = CoeficientePesoPorMetroSegunLote#(LOTE$)
'         PESOSALDOBOBINA = CANTMETROSBOBINA * PesoPorMetro#
'       End If

        
       MSF1.TextMatrix(REG&, 3) = LOTE$
       MSF1.TextMatrix(REG&, 4) = FORMATNUM$(CANTMETROSBOBINA, "F10.1")
       MSF1.TextMatrix(REG&, 5) = XVALO(!cod_inte)
       MSF1.TextMatrix(REG&, 6) = XVALO(!ID_OPERACION)
       MSF1.TextMatrix(REG&, 7) = FORMATNUM$(XVALO(!PESOKG), "F10.2")
       If AGREGARCODIGORANGO% > 0 Then
         CIRANGO% = CODIGORANGO&(CI1%)
         If CIRANGO% < 1 Then
           CODIGOX$ = "========"
           DESCRIXO$ = "======================"
         Else
           CODIGOX$ = CODEXT$(CIRANGO%)
           DESCRIXO$ = DESCRIT0$(CIRANGO%)
         End If
         MSF1.TextMatrix(REG&, 8) = CODIGOX$
         MSF1.TextMatrix(REG&, 9) = DESCRIXO$
       End If
      End If
       .MoveNext
    Loop
 End With
 RS.Close
 Set RS = Nothing

End Sub

Sub CargaMatTranfXOrdenConsumoFinalBolsasQuilmes(MSF1 As msFlexGrid, NORFA$, DEPOCORTE%, Optional FILTROOPERACION%, Optional AGREGARCODIGORANGO%, Optional ConsiderarConsumo%, Optional CamposOrden$)
  'CARGA LAS TRANSFERENCIAS REALIZADAS POR OF AL DEPOSITO DE CORTE
  
'**** AGREGAR EL CAMPO KG EN LA CONSULTA Y VER SI TRAE LOS KG Y Q NO LOS CALCULE COMO AHORA.
  EPAC$ = UCase$(TRIM$(EMPREAC$))

  MSF1.Rows = 1: REG& = 0
  SQLX$ = "SELECT SUM(CANTINGRE-CANTSALID) AS RESULTADO ,COD_INTE,COD_EXTE,DESCRIP, IDENLOTE,ID_OPERACION,PESOKG FROM MOVISTO WITH(NOLOCK)"
  SQLX$ = SQLX$ + " where "
  If ConsiderarConsumo% > 0 Then 'SI CONSIDERA CONSUMO VA A MOSTRAR EL RESULTADO DE LO TRANSFERIDO MENOS EL CONSUMO DE LA ORDEN FAB
     SQLX$ = SQLX$ + " (DOSECLAR = '" & TRIM(NORFA$) & "' or  DOPRICOR = '" & TRIM(NORFA$) & "')"
  Else
     SQLX$ = SQLX$ + " DOSECLAR = '" & TRIM(NORFA$) & "'"
  End If
   
  SQLX$ = SQLX$ + " AND DEPOMOVI = " & DEPOCORTE%
  If FILTROOPERACION% > 0 Then SQLX$ = SQLX$ + " AND ID_OPERACION = " & FILTROOPERACION%
  
  SQLX$ = SQLX$ + " GROUP BY COD_INTE,COD_EXTE,DESCRIP, IDENLOTE,ID_OPERACION,PESOKG"
  If TRIM$(CamposOrden$) <> "" Then
     SQLX$ = SQLX$ + " ORDER BY " & CamposOrden$
  Else
     SQLX$ = SQLX$ + " ORDER BY IDENLOTE "
  End If
'  Call CARGA_ENCABEZADO(MSF2, "CODIGO/A20;Descripción/A40;LOTE/A16;Saldo.Transf.Mts./f18.1;CI/f6;Op.N°/F6;Kg.Transf.(Orig.)/F16.2", Me)
'  Call CARGA_ENCABEZADO(MSF3, "CI1/A0;Còdigo/A16;Descripciòn/A32;Lote/a12;Consumo/F10.0;Kgs./F8.0", Me)
'  DEPOCORTE% = XVALO(Control("", "DEPODEST"))
'  Call mostrarDatosx(SQLX$)


  
  Set RS = READSET(SQLX$)
  With RS
    Do While Not .EOF
      'ESTA MODIFICACION ES POR QUE INGRESA EN EL MISMO DEPOSITO  Y SEGUN LA CONDICION NO MUESTRA EL REGISTRO 3) IF-000120-1 OF-000120-1 QUE ES EL INGRESO POR OF
     'ENTONCES ESTABA MOSTRANDO 0 COMO RESULTADO.
     ' cantingre   cantsalid   COD_INTE    COD_EXTE    DESCRIP depomovi    doseclar    dopricor    IDENLOTE    ID_OPERACION    PESOKG
            '''''0   1   804 ITP-004634-03   IMPRESO - LOPEZ MAURICIO GERMAN 05-04634-02-B   3   OF-000115-1 TR.102  OF-000120000    0   445
            '''''1   0   804 ITP-004634-03   IMPRESO - LOPEZ MAURICIO GERMAN 05-04634-02-B   3   OF-000115-1 TR.102  OF-000120000    0   445
            '''''1   0   804 ITP-004634-03   IMPRESO - LOPEZ MAURICIO GERMAN 05-04634-02-B   3   IF-000120-1 OF-000120-1 OF-000120000    0   445
            '''''0   1   804 ITP-004634-03   IMPRESO - LOPEZ MAURICIO GERMAN 05-04634-02-B   3   OF-000115-1 TR.101  OF-000120000    0   445
            '''''1   0   804 ITP-004634-03   IMPRESO - LOPEZ MAURICIO GERMAN 05-04634-02-B   3   OF-000115-1 TR.101  OF-000120000    0   445
      'DE ESTA MANERA VEO SI EL SALDO EN ESE DEPOSITO ES > 0
      LOTE$ = SAFETEXT$(!idenlote)
      CONDI$ = "DEPOMOVI = " & DEPOCORTE%
      CONDI$ = CONDI$ + " and idenlote = '" & LOTE$ & "'"
      CANTI = XVALO(VALOBADA("MOVISTO", "SUM(CANTINGRE-CANTSALID)", CONDI$, "NOMESS"))
      'HASTA AQUI
      If CANTI > 0 Then
       REG& = REG& + 1
       MSF1.Rows = REG& + 1
       CI1% = XVALO(!cod_inte)
       MSF1.TextMatrix(REG&, 1) = SAFETEXT$(!Cod_Exte)
       MSF1.TextMatrix(REG&, 2) = SAFETEXT$(!Descrip)
       LOTE$ = SAFETEXT$(!idenlote)

       CANTMETROSBOBINA = CANTI
       PesoPorMetro# = CoeficientePesoPorMetroSegunLote#(LOTE$)
       PESOSALDOBOBINA = CANTMETROSBOBINA * PesoPorMetro#
       If InStr(EPAC$, "QUILMES") > 0 Then
         MSF1.TextMatrix(REG&, 4) = STODEPFECHENMETROS#(CI1%, DEPOCORTE%, HOY(HO$), LOTE$)
         PESOSALDOBOBINA = STODEPFECHENKG#(CI1%, DEPOCORTE%, HOY(HO$), LOTE$)
       End If
'       If Unimed(CI1%) = "U." Then
'         CANTMETROSBOBINA = SaldoLotexMetros(CI1%, LOTE$, DEPOCORTE%)
'         PESOSALDOBOBINA = SaldoLotexKilos(CI1%, LOTE$, DEPOCORTE%)
'       Else
'         PesoPorMetro# = CoeficientePesoPorMetroSegunLote#(LOTE$)
'         PESOSALDOBOBINA = CANTMETROSBOBINA * PesoPorMetro#
'       End If

        
       MSF1.TextMatrix(REG&, 3) = LOTE$
       MSF1.TextMatrix(REG&, 4) = FORMATNUM$(CANTMETROSBOBINA, "F10.1")
       MSF1.TextMatrix(REG&, 5) = XVALO(!cod_inte)
       MSF1.TextMatrix(REG&, 6) = XVALO(!ID_OPERACION)
       MSF1.TextMatrix(REG&, 7) = FORMATNUM$(XVALO(!PESOKG), "F10.2")
       If AGREGARCODIGORANGO% > 0 Then
         CIRANGO% = CODIGORANGO&(CI1%)
         If CIRANGO% < 1 Then
           CODIGOX$ = "========"
           DESCRIXO$ = "======================"
         Else
           CODIGOX$ = CODEXT$(CIRANGO%)
           DESCRIXO$ = DESCRIT0$(CIRANGO%)
         End If
         MSF1.TextMatrix(REG&, 8) = CODIGOX$
         MSF1.TextMatrix(REG&, 9) = DESCRIXO$
       End If
      End If
       .MoveNext
    Loop
 End With
 RS.Close
 Set RS = Nothing

End Sub


Sub NuevaCargaMatTranfXOrden(MSF1 As msFlexGrid, NORFA$, DEPOCORTE%)
  'CARGA LAS TRANSFERENCIAS REALIZADAS POR OF AL DEPOSITO DE CORTE

  MSF1.Rows = 1: REG& = 0
  SQLX$ = " SELECT SUM(CANTINGRE-CANTSALID) AS RESULTADO ,COD_INTE,COD_EXTE,DESCRIP,ID_OPERACION,"
  SQLX$ = SQLX$ + "   IDENLOTE FROM MOVISTO WITH(NOLOCK) where  (DOSECLAR = '" & TRIM(NORFA$) & "' or"
  SQLX$ = SQLX$ + "   DOPRICOR = '" & TRIM(NORFA$) & "') AND DEPOMOVI = " & DEPOCORTE%
  SQLX$ = SQLX$ + "   GROUP BY COD_INTE,COD_EXTE,DESCRIP, IDENLOTE,ID_OPERACION"
  SQLX$ = SQLX$ + "   HAVING SUM(CANTINGRE-CANTSALID) > 0.5"
  SQLX$ = SQLX$ + "   ORDER BY IDENLOTE"

  DEPOCORTE% = XVALO(Control("", "DEPODEST"))
  Set RS = READSET(SQLX$)
  With RS
    Do While Not .EOF
       REG& = REG& + 1
       MSF1.Rows = REG& + 1
       CI1% = XVALO(!cod_inte)
       MSF1.TextMatrix(REG&, 1) = SAFETEXT$(!Cod_Exte)
       MSF1.TextMatrix(REG&, 2) = SAFETEXT$(!Descrip)
       LOTE$ = SAFETEXT$(!idenlote)
       CANTMETROSBOBINA = XVALO(!Resultado)
       PesoPorMetro# = CoeficientePesoPorMetroSegunLote#(LOTE$)
       PESOSALDOBOBINA = CANTMETROSBOBINA * PesoPorMetro#
       
       
       MSF1.TextMatrix(REG&, 3) = LOTE$
       MSF1.TextMatrix(REG&, 4) = FORMATNUM$(CANTMETROSBOBINA, "F10.1")
       MSF1.TextMatrix(REG&, 5) = XVALO(!cod_inte)
       MSF1.TextMatrix(REG&, 6) = XVALO(!ID_OPERACION)
       MSF1.TextMatrix(REG&, 7) = FORMATNUM$(XVALO(PESOSALDOBOBINA), "F10.2")
       .MoveNext
    Loop
 End With
 RS.Close
 Set RS = Nothing

End Sub


Sub CargaMatTranfXOrdenNoConsideraConsumos(MSF1 As msFlexGrid, NORFA$, DEPOCORTE%, Optional FILTROOPERACION%, Optional AGREGARCODIGORANGO%, Optional ConsiderarConsumo%, Optional CamposOrden$)
  'CARGA LAS TRANSFERENCIAS REALIZADAS POR OF AL DEPOSITO DE CORTE NO CONSIDERA CONSUMOS

  MSF1.Rows = 1: REG& = 0
  SQLX$ = "SELECT SUM(CANTINGRE-CANTSALID) AS RESULTADO ,COD_INTE,COD_EXTE,DESCRIP, IDENLOTE,ID_OPERACION FROM MOVISTO WITH(NOLOCK)"
  SQLX$ = SQLX$ + " where "
  SQLX$ = SQLX$ + " DOSECLAR = '" & TRIM(NORFA$) & "'"
   
  SQLX$ = SQLX$ + " AND DEPOMOVI = " & DEPOCORTE%
  SQLX$ = SQLX$ + " AND codimovi = 'TR'"
  If FILTROOPERACION% > 0 Then SQLX$ = SQLX$ + " AND ID_OPERACION = " & FILTROOPERACION%
  
  SQLX$ = SQLX$ + " GROUP BY COD_INTE,COD_EXTE,DESCRIP, IDENLOTE,ID_OPERACION"
  If TRIM$(CamposOrden$) <> "" Then
     SQLX$ = SQLX$ + " ORDER BY " & CamposOrden$
  Else
     SQLX$ = SQLX$ + " ORDER BY IDENLOTE "
  End If
'  Call CARGA_ENCABEZADO(MSF2, "CODIGO/A20;Descripción/A40;LOTE/A16;Saldo.Transf.Mts./f18.1;CI/f6;Op.N°/F6;Kg.Transf.(Orig.)/F16.2", Me)
'  Call CARGA_ENCABEZADO(MSF3, "CI1/A0;Còdigo/A16;Descripciòn/A32;Lote/a12;Consumo/F10.0;Kgs./F8.0", Me)

  Set RS = READSET(SQLX$)
  With RS
    Do While Not .EOF
      If XVALO(!Resultado) > 0 Then
       REG& = REG& + 1
       MSF1.Rows = REG& + 1
       CI1% = XVALO(!cod_inte)
       MSF1.TextMatrix(REG&, 1) = SAFETEXT$(!Cod_Exte)
       MSF1.TextMatrix(REG&, 2) = SAFETEXT$(!Descrip)
       LOTE$ = SAFETEXT$(!idenlote)
        CANTMETROSBOBINA = XVALO(!Resultado)
'        PesoPorMetro# = PESMETRO(CI1%)
        PesoPorMetro# = CoeficientePesoPorMetroSegunLote#(LOTE$)
        PESOSALDOBOBINA = CANTMETROSBOBINA * PesoPorMetro#
       MSF1.TextMatrix(REG&, 3) = LOTE$
       MSF1.TextMatrix(REG&, 4) = FORMATNUM$(CANTMETROSBOBINA, "F10.1")
       MSF1.TextMatrix(REG&, 5) = XVALO(!cod_inte)
       MSF1.TextMatrix(REG&, 6) = XVALO(!ID_OPERACION)
       MSF1.TextMatrix(REG&, 7) = FORMATNUM$(XVALO(PESOSALDOBOBINA), "F10.1")
       If AGREGARCODIGORANGO% > 0 Then
         CIRANGO% = CODIGORANGO&(CI1%)
         If CIRANGO% < 1 Then
           CODIGOX$ = "========"
           DESCRIXO$ = "======================"
         Else
           CODIGOX$ = CODEXT$(CIRANGO%)
           DESCRIXO$ = DESCRIT0$(CIRANGO%)
         End If
         MSF1.TextMatrix(REG&, 8) = CODIGOX$
         MSF1.TextMatrix(REG&, 9) = DESCRIXO$
       End If
      End If
       .MoveNext
    Loop
 End With
 RS.Close
 Set RS = Nothing

End Sub

Sub MostrarConsumoSegunOf(MSF As msFlexGrid, NORFA$, FRM As Form)
   MSF.Rows = 1: REG& = 0
'  Call CARGA_ENCABEZADO(MSF2, "CODIGO/A20;Descripción/A40;LOTE/A16;Saldo.Transf.Mts./f18.1;CI/f6;Op.N°/F6;Kg.Transf.(Orig.)/F16.2", Me)
'  Call CARGA_ENCABEZADO(MSF3, "CI1/A0;Còdigo/A16;Descripciòn/A32;Lote/a12;Consumo/F10.0;Kgs./F8.0", Me)

   SQLX$ = "SELECT SUM(CANTSALID) AS RESULTADO ,COD_INTE,COD_EXTE,DESCRIP, IDENLOTE,ID_OPERACION "
   SQLX$ = SQLX$ + "  FROM MOVISTO WITH(NOLOCK) "
   SQLX$ = SQLX$ + "  where  DOPRILAR = '" & NORFA$ & "' AND CODIMOVI = 'CF'"
   SQLX$ = SQLX$ + "  GROUP BY COD_INTE,COD_EXTE,DESCRIP, IDENLOTE,ID_OPERACION ORDER BY IDENLOTE"
   
   Set RS = READSET(SQLX$)
   With RS
    Do While Not .EOF
      If XVALO(!Resultado) <> 0 Then
               REG& = REG& + 1
               MSF.Rows = REG& + 1
               CI1% = XVALO(!cod_inte)
               LOTE$ = SAFETEXT$(!idenlote)
               MSF.TextMatrix(REG&, 1) = CI1%
               MSF.TextMatrix(REG&, 2) = SAFETEXT$(!Cod_Exte)
               MSF.TextMatrix(REG&, 3) = SAFETEXT$(!Descrip)
               MSF.TextMatrix(REG&, 4) = LOTE$
               CANTMETROSBOBINA = XVALO(!Resultado)
               MSF.TextMatrix(REG&, 5) = FORMATNUM$(CANTMETROSBOBINA, "F10.1")
               sumametros = sumametros + CANTMETROSBOBINA
               PesoPorMetro# = CoeficientePesoPorMetroSegunLote#(LOTE$)
               MSF.TextMatrix(REG&, 6) = FORMATNUM$(CANTMETROSBOBINA * PesoPorMetro#, "F10.1")
               sumakilos = sumakilos + XVALO(MSF.TextMatrix(REG&, 6))

      End If
       .MoveNext
    Loop
  End With
  RS.Close
  Set RS = Nothing
  MSF.Rows = REG& + 3
  MSF.TextMatrix(REG& + 2, 5) = FORMATNUM$(sumametros, "F12.1")
  MSF.TextMatrix(REG& + 2, 6) = FORMATNUM$(sumakilos, "F12.1")
  Call COLOREAR_GRILLA_SOLO_UNA_AFILA(MSF, &HE2D3C0, REG& + 2)
  
  FRM.TXTMETROSCONSUMIDOS.TEXT = TRIM$(FORMATNUM$(sumametros, "F12.1"))
  FRM.TXTKILOSCONSUMIDOS.TEXT = TRIM$(FORMATNUM$(sumakilos, "F12.1"))
  
'  Call ControlarsiExistenCodigosQueNoEstenEnGrillaSuperior(ContConsumo.MSF3, ContConsumo.MSF2, ContConsumo)
   
End Sub

Sub ControlarsiExistenCodigosQueNoEstenEnGrillaSuperior(MSFDESDE As msFlexGrid, MSFHACIA As msFlexGrid, FRM As Form)
    'REVISA QUE LOS CODIGOS CONSUMIDOS ESTEN EN LA GRILLA DE TRANSFERIDOS
     'SI NO ES ASI LOS AGREGA EN LA GRILLA DE ARRIBA EJEMPLO DE LAS CINTILLAS
    For i& = 1 To MSFDESDE.Rows - 1
       Existex% = 0
       CodigoA$ = MSFDESDE.TextMatrix(i&, 2)
       For J& = 1 To MSFHACIA.Rows - 1
          CodigoB$ = MSFHACIA.TextMatrix(J&, 1)
          If CodigoA$ = CodigoB$ Then
            Existex% = 1
            Exit For
          End If
       Next J&
       If Existex% = 0 Then 'LO AGREGA MAS REGISTROS CON LOS DATOS DE CONSMUMO QUE NO ESTABAN COMO TRANSFEREIDOS
          MSFHACIA.Rows = MSFHACIA.Rows + 1
          REG& = MSFHACIA.Rows - 1
          MSFHACIA.TextMatrix(REG&, 1) = MSFDESDE.TextMatrix(i&, 2)
          MSFHACIA.TextMatrix(REG&, 2) = MSFDESDE.TextMatrix(i&, 3)
          MSFHACIA.TextMatrix(REG&, 4) = MSFDESDE.TextMatrix(i&, 5)
          MSFHACIA.TextMatrix(REG&, 5) = MSFDESDE.TextMatrix(i&, 1)
          MSFHACIA.TextMatrix(REG&, 7) = MSFDESDE.TextMatrix(i&, 6)
          
       End If
    Next i&
'      Call CARGA_ENCABEZADO(MSF2, "CODIGO/A20;Descripción/A40;LOTE/A16;Saldo.Transf.Mts./f18.1;CI/f6;Op.N°/F6;Kg.Transf.(Orig.)/F16.2", Me)
'  Call CARGA_ENCABEZADO(MSF3, "CI1/A0;Còdigo/A16;Descripciòn/A32;Lote/a12;Consumo/F10.0;Kgs./F8.0", Me)
    'RECALCULA EL TOTAL CON LOS NUEVOS REGISTROS AGREGADOS CON LOS DATOS DE ARRIBA
    sumametros = SUMACOLUMNA#(MSFHACIA, 4, 1)
    sumakilos = SUMACOLUMNA#(MSFHACIA, 7, 1)
    
    MSFHACIA.Rows = REG& + 3
    MSFHACIA.TextMatrix(REG& + 2, 4) = FORMATNUM$(sumametros, "F12.1")
    MSFHACIA.TextMatrix(REG& + 2, 7) = FORMATNUM$(sumakilos, "F12.1")
    FRM.TXTMETROSTRANSFERIDOS = FORMATNUM$(sumametros, "F12.1")
    FRM.TXTKILTRANSFERIDOS = FORMATNUM$(sumakilos, "F12.1")
    
    Call COLOREAR_GRILLA_SOLO_UNA_AFILA(MSFHACIA, &HE2D3C0, REG& + 2)
End Sub
Sub CargaMatTranfTodasLasOrdenes(MSF1 As msFlexGrid, DEPOCORTE%, AbiertasoCerradas%, Optional AGREGARCODIGORANGO%, Optional ConsiderarConsumo%)
  'CARGA LAS TRANSFERENCIAS REALIZADAS POR OF AL DEPOSITO DE CORTE
  MSF1.Rows = 1: REG& = 0

  SQLX$ = "SELECT SUM(M.CANTINGRE-M.CANTSALID) AS RESULTADO ,M.COD_INTE,M.COD_EXTE,M.DESCRIP " ',M.DOPRICOR,DOSECLAR,o.FechCierr
  SQLX$ = SQLX$ + "  FROM MOVISTO M WITH(NOLOCK)"
  SQLX$ = SQLX$ + "  INNER JOIN ORDEFABR O ON (O.IdSubOr  = M.DoSecLar     OR O.IDSUBOR = M.DOPRICOR)"
  SQLX$ = SQLX$ + "  INNER JOIN MASTER MAS ON M.COD_INTE = MAS.CODINT"
  SQLX$ = SQLX$ + " where (M.DOSECLAR LIKE '_F-______-1'  OR M.DOSECLAR LIKE 'AJ.%')" ' QUE SEA OF, PF O AJUSTE"
'  SQLX$ = SQLX$ + " (left(M.DOSECLAR,3) = 'OF-'  or  left(M.DOPRICOR,3)   = 'OF-' )  "
  If AbiertasoCerradas% = 3 Or AbiertasoCerradas% = 1 Then
      SQLX$ = SQLX$ + " AND O.StatuOrf = " & AbiertasoCerradas%
  End If

  SQLX$ = SQLX$ + " AND M.DEPOMOVI = " & DEPOCORTE%
  SQLX$ = SQLX$ + " AND MAS.TIMATE <> 49" 'QUE NO SEA CODIGO RANGO
  
  SQLX$ = SQLX$ + " GROUP BY M.COD_INTE,M.COD_EXTE,M.DESCRIP" ',IDSUBOR,M.DOPRICOR,DOSECLAR,o.FechCierr
  SQLX$ = SQLX$ + " ORDER BY M.COD_EXTE "
  sumametros = 0: sumakilos = 0
  
  Set RS = READSET(SQLX$)
  With RS
    Do While Not .EOF
      If XVALO(!Resultado) <> 0 Then
'       If Left$(!DOSECLAR, 3) = "OF-" Then
               REG& = REG& + 1
               MSF1.Rows = REG& + 1
               CI1% = XVALO(!cod_inte)
               MSF1.TextMatrix(REG&, 1) = CI1%
               MSF1.TextMatrix(REG&, 2) = SAFETEXT$(!Cod_Exte)
               MSF1.TextMatrix(REG&, 3) = SAFETEXT$(!Descrip)
               CANTMETROSBOBINA = XVALO(!Resultado)
'               MSF1.TextMatrix(REG&, 4) = SAFETEXT$(!DOSECLAR)
               MSF1.TextMatrix(REG&, 4) = FORMATNUM$(CANTMETROSBOBINA, "F10.1")
               sumametros = sumametros + CANTMETROSBOBINA
               'PESMETRO trae el peso del maestro y lo multiplica por los metros
               MSF1.TextMatrix(REG&, 5) = FORMATNUM$(CANTMETROSBOBINA * PESMETRO(CI1%), "F10.1")
               sumakilos = sumakilos + XVALO(MSF1.TextMatrix(REG&, 5))
'               MSF1.TextMatrix(REG&, 7) = FECHATEX$(CVINT(!FechCierr))
'       End If
      End If
       .MoveNext
    Loop
 End With
 RS.Close
 Set RS = Nothing
 MSF1.Rows = REG& + 3
 MSF1.TextMatrix(REG& + 2, 4) = FORMATNUM$(sumametros, "F12.1")
 MSF1.TextMatrix(REG& + 2, 5) = FORMATNUM$(sumakilos, "F12.1")
 Call COLOREAR_GRILLA_SOLO_UNA_AFILA(MSF1, &HE2D3C0, REG& + 2)

 

End Sub

Sub CargaMatTranfXOrdenParaCierre(MSF1 As msFlexGrid, NORFA$, DEPOCORTE%, Optional FILTROOPERACION%, Optional AGREGARCODIGORANGO%)
   'SE CONSIDERA TAMBIEN LO TRANSFERIDO AL DEPOSITO DE ESLINGA
   DEPOESLINGA% = XVALO(Control("CONTEPLA", "DEPOESLI"))
  'ES CargaMatTranfXOrden MODIFICADO PARA PRESENTAR LA PANTALLA DE EVALUACION PARA CIERRE DE FABRICACION
  SQLX$ = "SELECT SUM(CANTINGRE-CANTSALID) AS RESULTADO ,COD_INTE,COD_EXTE,DESCRIP,ID_OPERACION"
  SQLX$ = SQLX$ + "  FROM MOVISTO WITH(NOLOCK) where (DOSECLAR = '" & TRIM(NORFA$) & "' or doprilar ='" & TRIM(NORFA$) & "')"
  SQLX$ = SQLX$ + "  AND (DEPOMOVI = " & DEPOCORTE%
  SQLX$ = SQLX$ + "  Or DEPOMOVI = " & DEPOESLINGA%
  SQLX$ = SQLX$ + " )"
  If FILTROOPERACION% > 0 Then SQLX$ = SQLX$ + " AND ID_OPERACION = " & FILTROOPERACION%
    
  SQLX$ = SQLX$ + " GROUP BY COD_INTE,COD_EXTE,DESCRIP,ID_OPERACION"
  SQLX$ = SQLX$ + " having  SUM(CANTINGRE-CANTSALID) >0 "
  SQLX$ = SQLX$ + " ORDER BY ID_OPERACION,COD_INTE "
  MSF1.Rows = 1: REG& = 0
'  SQLX$ = "SELECT SUM(CANTINGRE-CANTSALID) AS RESULTADO ,COD_INTE,COD_EXTE,DESCRIP, IDENLOTE,ID_OPERACION FROM MOVISTO WITH(NOLOCK)"
'  SQLX$ = SQLX$ + " where DOSECLAR = '" & TRIM(NORFA$) & "'  AND DEPOMOVI = " & DEPOCORTE%
'  SQLX$ = SQLX$ + " GROUP BY COD_INTE,COD_EXTE,DESCRIP, IDENLOTE,ID_OPERACION"
'  SQLX$ = SQLX$ + " ORDER BY ID_OPERACION,COD_INTE,IDENLOTE "
  CIFAB% = CODIFAB2%(NORFA$)
  Set RS = READSET(SQLX$)
  With RS
    Do While Not .EOF
      If XVALO(!Resultado) > 0 Then
       REG& = REG& + 1
       MSF1.Rows = REG& + 1
       CI1% = XVALO(!cod_inte)
       MSF1.TextMatrix(REG&, 1) = SAFETEXT$(!Cod_Exte)
       MSF1.TextMatrix(REG&, 2) = SAFETEXT$(!Descrip)
       LOTE$ = ""
       CANTMETROSBOBINA = XVALO(!Resultado)
       NUMOPER% = XVALO(!ID_OPERACION)
       DESCOPE$ = DescriOperSegunCodigo(CIFAB%, NUMOPER%)
'      PesoPorMetro# = PESMETRO(CI1%)
       PesoPorMetro# = CoeficientePesoPorMetroSegunLote#(LOTE$)
       PESOSALDOBOBINA = CANTMETROSBOBINA * PesoPorMetro#
       
       MSF1.TextMatrix(REG&, 3) = LOTE$
       MSF1.TextMatrix(REG&, 4) = FORMATNUM$(CANTMETROSBOBINA, "F10.2")
       MSF1.TextMatrix(REG&, 5) = XVALO(!cod_inte)
       MSF1.TextMatrix(REG&, 6) = XVALO(!ID_OPERACION)
       MSF1.TextMatrix(REG&, 7) = DESCOPE$
       MSF1.TextMatrix(REG&, 8) = FORMATNUM$(XVALO(PESOSALDOBOBINA), "F10.2")
       
       If AGREGARCODIGORANGO% > 0 Then
         CIRANGO% = CODIGORANGO&(CI1%)
         If CIRANGO% < 1 Then
           CODIGOX$ = "========"
           DESCRIXO$ = "======================"
         Else
           CODIGOX$ = CODEXT$(CIRANGO%)
           DESCRIXO$ = DESCRIT0$(CIRANGO%)
         End If
         MSF1.TextMatrix(REG&, 9) = TRIM$(CODIGOX$)
         MSF1.TextMatrix(REG&, 10) = TRIM$(DESCRIXO$)
         
       End If
      End If
       .MoveNext
    Loop
 End With
 RS.Close
 Set RS = Nothing
End Sub

Sub GrabarOperacion(NORFA$, CODIEMPR%, CIFAB%, NOP%, CANTI, LOTE$, CIELEM%, CANTICORTEXBOLSON, NUMEOPERARIO%, CORTEENMETROS, PESOKG, RET_NEWORDEN%, Optional Ret_NuevoIDGenerado&, Optional IDCOORDUBIX&, Optional COORDUBIX$, Optional IDREGPADRE&, Optional DESCRIOPERACION$, Optional IdMovistoGenerado&)
        FE% = HOY(HO$)
        fech$ = FETEXCOR1$(FE%)
        HORAX$ = Time
        HORAX$ = Left$(HORANUM(HORAX$), 4)
        HORACENTECIIMAL = CHoraSentecimal(HORAX$)
        nombre$ = TRIM$(ECOARCH$("PADRON", MKI$(NUMEOPERARIO%)))
        REFEOF$ = ReferenSegunOf2$(NORFA$)
        NEWORDEN% = 1 + UltimoNumOrdenSegunOFyOperacion2&(NORFA$, NOP%)
        RET_NEWORDEN% = NEWORDEN% 'VALOR DE RETORNO
        CODIEXFAB$ = CODEXT$(CIFAB%)
        DESCRIFAB$ = DESCRIT0$(CIFAB%)
        If TRIM$(DESCRIOPERACION$) = "" Then
           DESCRIOPERACION$ = DescriOperSegunCodigo(CIFAB%, NOP%)
        End If
        CANTPROG = CantiProyectadaOF2(NORFA$)
        CodExtElem$ = CODEXT$(CIELEM%)
        
        Call ABREORFAB
        SQLX$ = " SELECT * FROM InfoFab "
        Dim INFOFFFAB As ADODB.Recordset
        Set INFOFFFAB = New ADODB.Recordset
        INFOFFFAB.Open FILTSQL$(SQLX$), FLEXCONN, adOpenDynamic, adLockPessimistic, adCmdText
        With INFOFFFAB
            .AddNew
            On Error Resume Next
            !CODIEMPR = CodiEmp%
            On Error GoTo 0
            !idsubor = NORFA$
            !DescriOF = REFEOF$
            !NuOrInf = NEWORDEN%
            !cod_inte = CIFAB%
            !Cod_Exte = CODIEXFAB$
            !Descrip = DESCRIFAB$
            !NUMEOPER = NOP%
            !AltOper = 0
            !descoper = DESCRIOPERACION$
            !ReTraba = 0
            !CanOrPro = CANTPROG
            !CANTREAL = CANTI
            !FeHorLib = Now
            !Hora_Des = ROUND(HORACENTECIIMAL, 0)
            !Hora_Has = ROUND(HORACENTECIIMAL, 0)
            !TiemEjec = HORACENTECIIMAL
            !MotiRecha = 0
            !Status = 0
            !idenlote = LOTE$
            !codielem = CIELEM%
            !CodXElem = CodExtElem$
            !FACTORCORTE = CANTICORTEXBOLSON
            !FEHORINF = Now
            !COORDUBIC = COORDUBIX$
            !IDCOORDUBIC = IDCOORDUBIX&
            !NUMOPERARIO = NUMEOPERARIO%
            !NameOperario = nombre$
            !corte = CORTEENMETROS
            !Peso = PESOKG
            !IDREGPADRE = XVALO(IDREGPADRE&)
            !COORDUBIC = "SEMITERMINADOS"  'SECTOR$
            !IDCOORDUBIC = 3 'IDSECTOR
            !IDMOVISTO = XVALO(IdMovistoGenerado&)
            
            .Update
        End With
        Set INFOFFFAB = Nothing
        '
        'GRABAR EN OPERFAF
        SQLX$ = "SELECT @@IDENTITY AS NUEVOID FROM INFOFAB WITH(NOLOCK)"
        Set RS = READSET(SQLX$)
        Ret_NuevoIDGenerado& = RS!NUEVOID

        'CANTIDAD ACUMULADA DE LA OPERACION
        CONDI$ = "IdSubOr = '" & NORFA$ & "' AND NumeOper = " & NOP% & " AND RETRABA < 1"
        CAREAL# = XVALO(VALOBADA("INFOFAB", "SUM(CantReal)", CONDI$, "NOMESS"))
        
        'TIEMO INFORMADO ACUMULADA DE LA OPERACION
        CONDI$ = "IdSubOr = '" & NORFA$ & "' AND NumeOper = " & NOP%
        TIEINF = XVALO(VALOBADA("INFOFAB", "SUM(TiemEjec)", CONDI$, "NOMESS"))
        
        SQLZ$ = " SELECT * FROM OPERFAB "
        SQLZ$ = SQLZ$ + " WHERE IdSubOr = '" & NORFA$ & "' AND NumeOper = " & NOP%
        Dim OPERBU As ADODB.Recordset
        Set OPERBU = New ADODB.Recordset
        OPERBU.Open FILTSQL$(SQLZ$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText

        With OPERBU
         On Error Resume Next
         .MoveFirst
         If Err < 1 Then
            On Error GoTo 0
           !CANTREAL = CAREAL#
           !CANTPEND = XVALO(!CANTPROG) - CAREAL#
             If XVALO(!CANTPEND) < 0 Or STATUOP% = 3 Or STATUOP% = 8 Then !CANTPEND = 0
             CANTPE# = XVALO(!CANTPEND)
             MODO% = CAINFANT%
           !TiemEjec = TIEINF
           !TiemPend = TIEMAQUI(Ci%, NOPE%, CANTPE#, MODO%)  ' 1 es para que no considere preparacion
              If !TiemPend < 0 Or STATUOP% = 3 Or STATUOP% = 0 Then !TiemPend = 0
           !StatOper = STATUOP%
           .Update
         End If
         On Error GoTo 0
   End With

        
 End Sub
 
Sub PRINTCINTILLA(CIOF%, NORFA$, MEDIDA, CANTI, FE%)
    '
    RUTAEXCEL$ = Control("COTICONT", "EXFABCIN")
    If EXISTE(RUTAEXCEL$) < 1 Then
       Call COMUNI("Falta Configurar Ruta de Excel Impresión de Cintillas (Configurable en Ficha Técnica)")
       Exit Sub
    End If

    With VENTABNEW.MSF
       .Rows = 6: .Cols = 2
       .TextMatrix(1, 1) = NORFA$
       .TextMatrix(2, 1) = FECHATEX$(FE%)
       .TextMatrix(3, 1) = TRIM$(MEDIDA)
       .TextMatrix(4, 1) = CODEXT$(CIOF%)
       .TextMatrix(5, 1) = CANTI
    End With
    
    PDES$ = Control$(IDENTER$, "PRINTSER") 'PUERTO DE IMPRESION CONFIGURADO PARA IMPRESION DE COTIZACION
    If PDES$ = "" Then PDES$ = Control$("*", "PRINTSER")

    'OJO SI SE LLAMA A PANTALLA A VENTABNEW CUANDO SE ACEPTA BORRA LAS FILAS QUE ESTAN VACIAS
'    VENTABNEW.Show 1
    Call Param_Print_Excel(RUTAEXCEL$, "HOJA1", VENTABNEW.MSF, PDES$, "1", 1, 11, "", 1, 0)                         'imprime en excel la 1,2 Y 3 columna de mi msf en la columna 1 a partir de la fila 11 en el excel (este graba E imprime)
    Unload VENTABNEW
End Sub
Function TaraSegunIDTabla#(ID%)
    CONDI$ = "LISTA = 'TABLATARAS' AND CONVERT(INT,CAMPO1) = " & ID%
    TaraSegunIDTabla# = XVALO(VALOBADA("DATASQL", "CAMPO3", CONDI$, "NOMESS"))
End Function

Sub CargaOperacionesInformadasMasSubProcesos(NORFA$, CIFAB%, MSF As msFlexGrid, Optional MAYOROIGUAL&, Optional DESCRIPCIONOPERACION$, Optional MENOROIGUAL&)
    MSF.Rows = 1
    Dim CTP As New Conteplast
    CANTI = CantiProyectadaOF2(NORFA$)
    
    'DEVUELVE LAS OPERACIONES INFORAMDAS SEGUN OF. CON EXCEPCION DE LAS QUE TIENEN SUB-PROCESO
    SQLX$ = " SELECT * FROM INFOFAB I WITH(NOLOCK)"
    
    
    SQLX$ = SQLX$ + " WHERE NOT EXISTS"
    SQLX$ = SQLX$ + " (SELECT * FROM INFOFAB AS I2"
    SQLX$ = SQLX$ + " WHERE I.INFOFAB_ID = I2.IDREGPADRE)"
    SQLX$ = SQLX$ + " and IDSUBOR = '" & NORFA$ & "'"
    
    
    If XVALO(MENOROIGUAL&) > 0 Then SQLX$ = SQLX$ + " AND (IDCOORDUBIC <= " & MENOROIGUAL& & " or IDCOORDUBIC is null)"
    If XVALO(MAYOROIGUAL&) > 0 Then SQLX$ = SQLX$ + " AND IDCOORDUBIC >= " & MAYOROIGUAL&
    
    If TRIM$(DESCRIPCIONOPERACION$) <> "" Then SQLX$ = SQLX$ + " AND DESCOPER = '" & DESCRIPCIONOPERACION$ & "'"
    SQLX$ = SQLX$ + " ORDER BY NUMEOPER,FeHorInf,Cod_inte,IDENLOTE"
    Set RS = READSET(SQLX$)
    With RS
      Do While Not .EOF
        REG& = REG& + 1
        MSF.Rows = REG& + 1
'        VER DE GUARDAR EL CODIGO ELEMENTO EN EL INFOFAB PARA LUEGO LEERLO
        DESOP$ = SAFETEXT$(!descoper)
        CIELEM% = XVALO(!codielem)
        CONDI$ = "Codintconj = " & CIFAB%
        CONDI$ = CONDI$ + " and codintelem = " & CIELEM%
        If DESOP$ <> "" Then
           CONDI$ = CONDI$ + " and LEFT(COMPONENTE"
           CONDI$ = CONDI$ + " ," & Len(DESOP$)
           CONDI$ = CONDI$ + ") = '" & DESOP$ & "'"
        End If
       CORTEX = XVALO(VALOBADA("Coti1", "CORTE", CONDI$, "NOMESS")): If CORTEX < 1 Then CORTEX = 1
        MSF.TextMatrix(REG&, 1) = SAFETEXT$(!Cod_Exte)
        MSF.TextMatrix(REG&, 2) = SAFETEXT$(!Descrip)
        MSF.TextMatrix(REG&, 3) = SAFETEXT$(!idenlote)
        MSF.TextMatrix(REG&, 4) = FORMATNUM$(XVALO(!CANTREAL) * 1, "F10.1")
        MSF.TextMatrix(REG&, 5) = CIELEM%
        MSF.TextMatrix(REG&, 6) = XVALO(!NUMEOPER)
        MSF.TextMatrix(REG&, 7) = SAFETEXT$(!descoper)
        MSF.TextMatrix(REG&, 8) = SAFETEXT$(!TiemEjec)
        MSF.TextMatrix(REG&, 9) = SAFETEXT$(!CodXElem)
        MSF.TextMatrix(REG&, 10) = SAFETEXT$(!COORDUBIC)
        MSF.TextMatrix(REG&, 11) = XVALO(!NuOrInf) ' NUMERO DE ORDEN DE INFORME
        MSF.TextMatrix(REG&, 12) = XVALO(!IDCOORDUBIC) ' NUMERO DE CODIGO DE UBICACION
        MSF.TextMatrix(REG&, 13) = XVALO(!NUMOPERARIO) ' NUMERO DE OPERARIO
        MSF.TextMatrix(REG&, 14) = SAFETEXT(!NameOperario) ' NOMBRE DE OPERARIO
        MSF.TextMatrix(REG&, 15) = TRIM$(FORMATNUM$(XVALO(!Peso), "F12.2")) ' PESO INFORMADO
        MSF.TextMatrix(REG&, 16) = XVALO(!INFOFAB_ID) '
        MSF.TextMatrix(REG&, 17) = XVALO(!IDREGPADRE) '
        MSF.TextMatrix(REG&, 18) = XVALO(!OperFucionada) '
        If XVALO(!IDREGPADRE) > 0 Or XVALO(!OperFucionada) > 0 Then Call COLOREAR_GRILLA_SOLO_UNA_AFILA(MSF, &HE2D3C0, REG&)
        MSF.TextMatrix(REG&, 19) = XVALO(!IDMOVISTO)
80      .MoveNext
      Loop
    End With
    RS.Close
    Set RS = Nothing
    
    'RECORRO LA GRILLA Y ME QUEDO SOLO CON LAS OPERACIONES QUE NO POSEEN SUB-OPERACIONES Y LAS
    'SUB-OPERACIONES INFORMADAS
    i& = 0
    While i& < MSF.Rows - 1
        i& = i& + 1
        CANTI = XVALO(MSF.TextMatrix(i&, 4))
        IDX& = XVALO(MSF.TextMatrix(i&, 16))
        IDPADRE1& = XVALO(MSF.TextMatrix(i&, 17))
        OPERFU& = XVALO(MSF.TextMatrix(i&, 18))
        If OPERFU& > 0 Then Call cmdEliminaItem(MSF, i&): i& = i& - 1: GoTo 85
        If IDPADRE1& < 1 Then 'SI TIENE INFORMADO QUE ES HIJO DE UN PADRE LO DEJO
            For J& = 1 To MSF.Rows - 1 'SI NO ES ASI BUSCO SI ESE PADRE TIENE UN SUBPROCESO
                IDPADRE& = XVALO(MSF.TextMatrix(J&, 17)) 'Y SI LO TIENE ELIMINO EL REGISTRO PADRE
                If IDX& = IDPADRE& Or OPERFU& > 0 Then
                   If CANTI < 1 Then 'AGREGUE ESTE IF POR QUE COMO AHORA DESCUENTA LA CANTIDAD DEL SUB-PROCESO
                      Call cmdEliminaItem(MSF, i&) 'EN EL CASO QUE NO HAYA SIDO POR EL TOTAL QUEDA UN SALDO(25_08_20)
                      i& = i& - 1
                      Exit For
                   End If
                End If
                
            Next J&
        End If
85  Wend

End Sub

Function ListaSelecRemitosTallerConteplast$()

                TituloLista$ = "Remitos de Entrega a Taller"
                Campos$ = "O.Fab./A12;Remito Nro/A14;Fecha/D8;Cta/F7;Proveedor/A24;Còdigo/A20;Descripciòn/A32"
                Call CARGA_ENCABEZADO(FRMZELECHO.msf2, Campos$, FRMZELECHO)
                Call CARGA_ENCABEZADO(FRMZELECHO.MSF1, Campos$, FRMZELECHO)
                FRMZELECHO.Caption = TituloLista$
                SQLX$ = "SELECT DOPRILAR, FROM MOVTALLER WHERE "
    
                SQLX$ = "SELECT M.DOPRILAR,M.DocuOrig,M.FECHMOV,M.NUME_PROV,P.Raso_Cli,O.Cod_Exte,O.Descrip  FROM MOVTALLER M"
                SQLX$ = SQLX$ + " LEFT JOIN Proved12 P ON P.Nume_Cli = M.Nume_Prov"
                SQLX$ = SQLX$ + " LEFT JOIN ORDEFABR O ON O.IdSubOr = M.DOPRILAR"
                SQLX$ = SQLX$ + " WHERE CODIMOVI = 'ST'"
                SQLX$ = SQLX$ + " GROUP BY  M.DOPRILAR,M.DocuOrig,M.FECHMOV,M.NUME_PROV,P.Raso_Cli,O.Cod_Exte,O.Descrip"
                SQLX$ = SQLX$ + " ORDER BY DOPRILAR DESC"

                Set Rst = FLEXCONN.Execute(SQLX$)
                Call Carga_MSF_Recordset(Rst, Campos$, FRMZELECHO.msf2, 1)
                On Error Resume Next
                Rst.Close
                Set Rst = Nothing
                On Error GoTo 0
                    
                FRMZELECHO.Width = 16500
                FRMZELECHO.Show 1
                
                Remito$ = TRIM(RESP_ZELE_VECT(2))
                
                ListaSelecRemitosTallerConteplast$ = Remito$
End Function

Function ListaSelecRecepcionTallerConteplast$()

                TituloLista$ = "Recepciones de Taller"
                Campos$ = "O.Fab./A12;Recepciòn Nro/A14;Referencia/A32;Fecha/D8;Cta/F7;Proveedor/A24"
                Call CARGA_ENCABEZADO(FRMZELECHO.msf2, Campos$, FRMZELECHO)
                Call CARGA_ENCABEZADO(FRMZELECHO.MSF1, Campos$, FRMZELECHO)
                FRMZELECHO.Caption = TituloLista$
    
                SQLX$ = "SELECT M.DOPRILAR,M.DocuOrig,M.NUME_PROV,M.FECHMOV,P.Raso_Cli,O.Cod_Exte,O.Descrip  FROM MOVTALLER M"
                SQLX$ = SQLX$ + " LEFT JOIN Proved12 P ON P.Nume_Cli = M.Nume_Prov"
                SQLX$ = SQLX$ + " LEFT JOIN ORDEFABR O ON O.IdSubOr = M.DOPRILAR"
                SQLX$ = SQLX$ + " WHERE CODIMOVI = 'DT'"
                SQLX$ = SQLX$ + " GROUP BY  M.DocuOrig,M.DOPRILAR,M.NUME_PROV,M.FECHMOV,P.Raso_Cli,O.Cod_Exte,O.Descrip"
                SQLX$ = SQLX$ + " ORDER BY M.FECHMOV DESC"

                Set Rst = FLEXCONN.Execute(SQLX$)
                Call Carga_MSF_Recordset(Rst, Campos$, FRMZELECHO.msf2, 1)
                On Error Resume Next
                Rst.Close
                Set Rst = Nothing
                On Error GoTo 0
                    
                FRMZELECHO.Width = 16500
                FRMZELECHO.Show 1
                
                DOCUMENTO$ = TRIM(RESP_ZELE_VECT(2))
                
                ListaSelecRecepcionTallerConteplast$ = DOCUMENTO$
End Function


Function ListaIdInfofabSegunRemitoMovTaller$(NROREMITO$)
        'DEVUELVE UNA CADENA DE NUMEROS DE ID SEPARADOS POR COMA
        VUELTA% = 0
        SQLX$ = "SELECT idcalidad FROM MOVTALLER M WITH(NOLOCK)"
        SQLX$ = SQLX$ + " WHERE CODIMOVI = 'ST' and M.DOCUORIG = '" & NROREMITO$ & "'"

        Set Rst = READSET(SQLX$)
        With Rst
          Do While Not .EOF
            IDREG& = XVALO(!IdCalidad)
            If VUELTA% = 0 Then
               ListaIdInfofabSegunRemitoMovTaller$ = ListaIdInfofabSegunRemitoMovTaller$ & IDREG&
               VUELTA% = 1
            Else
               ListaIdInfofabSegunRemitoMovTaller$ = ListaIdInfofabSegunRemitoMovTaller$ & "," & IDREG&
            End If
            .MoveNext
          Loop
        End With

        Rst.Close
        Set Rst = Nothing
        On Error GoTo 0
        ListaIdInfofabSegunRemitoMovTaller$ = TRIM$(ListaIdInfofabSegunRemitoMovTaller$)
                    
End Function

Sub DuplicaRegistrosDeSaldoXDiferencia(CantidadDelRegistroSeleccionado&, IDDelRegistroSeleccionado&, CantidadIngresadaXUsuario, KilosRegistrosSeleccionado, NumeroOperacionRegistrosSeleccionado%, KilosIngresadoXUsuario#, Optional ESFUCION%)
        'CRITERIO DE RUTINA
        'SE UTILIZA EN EL CASO QUE SE INFORMA DE UNA OPERACION UN SUB-PROCESO DONDE LA CANTIDAD DE ESTE
        'ES MENOR A LA CANTIDAD DEL REGISTRO SELECCIONADO
        'EJ.: HAY UNA OPERACION POLLERA CON CANTIDAD 100 U. DE CORTE Y PESO 200 KG
        'SE SELECCIONA Y SE INFORMA DE ESTE 60 U. POR UN PESO QUE MARRQUE LA BALANZA (ESTE ULTIMO NO SE CONSIDERA PARA EL PRORATEO)
        'ENTONCES SE ACTUALIZA EL REGISTRO SELECCIONADO CON LA CANTIDAD DE CORTES INGRESADA Y EL PESO PROPORCIONAL SEGUN COEFICIENTE DEL PESO X UNIDAD
        'Y SE GENERA UN NUEVO REGISTRO X LA DIFERENCIA Y EL NUEVO PESO MISMO CRITERIO
        
        
        CANTIDIF = XVALO(CantidadDelRegistroSeleccionado&) - XVALO(CantidadIngresadaXUsuario)
        KILOSDIF = XVALO(KilosRegistrosSeleccionado) - XVALO(KilosIngresadoXUsuario#)
        'SE GENERA UN NUEVO REGISTRO CON
        CONDI$ = "INFOFAB_ID =" & IDDelRegistroSeleccionado&
        If ESFUCION% > 0 Then
           Call ACTUREGISTRO("INFOFAB", "CANTREAL", CONDI$, CantidadIngresadaXUsuario, REGAF&, "NOMESS")
        End If
        NROOF$ = VALOBADA("INFOFAB", "IDSUBOR", CONDI$, "NOMESS")
'        '**** EL CAMPO ORIGINAL SE MANTIENE COMO ESTA
'                   ESTE NO DEBERIA MOSTRARSE POR QUE ES PADRE Y DEBERIAN QUEDAR A LA VISTA EL DUPLICADO Y EL
'                    QUE SE GENERA NUEVO QUE LO HACE POR FUERA DE ESTA RUTINA

'        SQLX$ = "UPDATE INFOFAB SET CANTREAL = " & Diferencia&
'        SQLX$ = SQLX$ + "  ,PESO = '" & FORMATNUM$(KILOCALCU#, "F16.5") & "'"
'        SQLX$ = SQLX$ + "  WHERE " & CONDI$
'        FLEXCONN.Execute SQLX$

        NEWID& = 1 'EN AL RUTINA DuplicarRegistro Si NEWID& >  0 DEVUELVE EL NUEVO ID GENERADO
        Call DuplicarRegistro("INFOFAB", CONDI$, NEWID&)
        'BUSCO EL VALOR SIGUIENTE DE INFORME SEGUN OF Y OPER
        CONDI$ = "IdSubOr = '" & NROOF$ & "' AND NumeOper = " & NumeroOperacionRegistrosSeleccionado%
        NUEVOMONUMORDENOP% = 1 + XVALO(VALOBADA("INFOFAB", "MAX(nuorinf)", CONDI$, "NOMESS"))
        
        'ACTUALIZO EL CAMPO CREADO
        CONDI$ = "INFOFAB_ID =" & NEWID&
        SQLX$ = "UPDATE INFOFAB SET CANTREAL = " & CANTIDIF 'XVALO(CantidadIngresadaXUsuario)
        SQLX$ = SQLX$ + "  ,nuorinf = " & NUEVOMONUMORDENOP%
        SQLX$ = SQLX$ + "  ,PESO = '" & FORMATNUM$(KILOSDIF, "F16.5") & "'" 'FORMATNUM$(KilosIngresadoXUsuario#, "F16.5") & "'"
        SQLX$ = SQLX$ + "  ,IDREGPADRE = 0 " ' & IDDelRegistroSeleccionado&
        SQLX$ = SQLX$ + "  ,CantOrigCte = 0 " 'LA CANTIDAD ORIGINAL VA A 0 SE MANTIENE EN EL REGISTRO ORIGINAL
        
        
        SQLX$ = SQLX$ + "  WHERE  " & CONDI$
        FLEXCONN.Execute SQLX$
   
End Sub
Function FuncionLINER_SI_NO$(CI1%)
    NROFT& = NroFTecnicaXCodigo(CI1%)
    FuncionLINER_SI_NO$ = UCase$(TRIM$(VALOBADA("RECETACTP", "LINERSINO", "NUMRECETACTP = " & NROFT&, "NOMESS")))
End Function
Function CodEspTecCliente$(CI1%)
    'CODIGO ESPECIFICACION TECNICA CLIENTE QUE SE GUARDA EN AL FICHA TECNICA
    NROFT& = NroFTecnicaXCodigo(CI1%)
    CodEspTecCliente$ = UCase$(TRIM$(VALOBADA("RECETACTP", "CODESPTECTE", "NUMRECETACTP = " & NROFT&, "NOMESS")))
End Function
Sub PRINTETIQUETACTP(CI1%, NORFA$, CODIGOEDITABLE$, CANTI, FECHA, CANTPAQ, Optional IDMOVISTO&, Optional CANTIMPRE%, Optional TextoQuilmes$)
    If CI1% < 1 Then
       Call COMUNI("No Hay un Código Activo")
       Exit Sub
    End If
    '
    NORFA$ = TRIM$(NORFA$)
    If NORFA$ = "" Then
       Call COMUNI("Falta Orden de Fabriación")
       Exit Sub
    End If
    '
    PDES$ = Control$(IDENTER$, "PRINTSER") 'PUERTO DE IMPRESION CONFIGURADO PARA IMPRESION DE COTIZACION
    If PDES$ = "" Then PDES$ = Control$("*", "PRINTSER")

    '
    'RUTAEXCEL$ = "X:\CONTEPLAST\FLEXOFT\EXCEL\EtiquetaCTP.xlsx" 'CONTROL("COTICONT", "EXCEETIQ")
    RUTAEXCEL$ = Control("COTICONT", "EXCEETIQ")
    
    If EXISTE(RUTAEXCEL$) < 1 Then
       Call COMUNI("Falta Configurar Ruta de Plantilla Excel de Formulario de Etiqueta")
       Exit Sub
    End If
    FORIPRE$ = TRIM$(Control("ETIQPRTE", "ETPROTER"))
    If FORIPRE$ = "" Then
       Call COMUNI("Falta Configurar Etiqueta de Cierre de Orden de Fabricación\Configurar Formulario desde la Ficha Técnica > Congiguración General")
       Exit Sub
    End If
    '
'1 razon social
'2 orden de compra
'3 codigo p.t.
'4 medidas
'5 descripcion de p.t.
'6 orden de fabricacion
'7 fecha del cierre
'8 Cantidad de paquetes
'9 Cantidad de Bolsas

    Dim CTP As New Conteplast
    NP& = NroPedidoSegunOf2&(NORFA$)
    NC = cliente_SEGUN_PEDIDO(XVALO(NP&))
    RASO$ = rasocli$(NC)
    If TextoQuilmes$ <> "" Then 'a pedido de centurion se cambia la orden de compra por un texto de referencia
       OC$ = TextoQuilmes$      'si viene vacio imprime la o.compra
    Else
       OC$ = OrdenCompraPed$(NP&): If OC$ = "" Then OC$ = "-"
    End If
    CODI$ = CODEXT$(CI1%)
    DESKRI$ = DESCRIT0$(CI1%)
    NF& = FICHATECNCIA_SEGUN_CODIGO&(CI1%)
    ANCHOX = CTP.ANCHOCTP(CI1%)
    LARGOX = CTP.LARGOCTP(CI1%)
    ALTOX = CTP.ALTOCTP(CI1%)
    CODCLIENTE$ = CTP.CodEspTecCliente$(CI1%)
    MEDIDAS$ = SAFETEXT$(ANCHOX) & " X " & LARGOX & " X " & ALTOX & " CM."
    '
'    IdCalidad
    CANTIPAQUETE$ = TRIM$(InputBox$("Informar", "Ingrese Cantidad de Paquetes", XVALO(CANTPAQ)))
    If XVALO(CANTIPAQUETE$) < 1 Then CANTIPAQUETE$ = "1"
    
    With VENTABNEW.MSF
       .Rows = 11: .Cols = 2
       .TextMatrix(1, 1) = RASO$
       .TextMatrix(2, 1) = OC$
       .TextMatrix(3, 1) = CODI$
       .TextMatrix(4, 1) = CODCLIENTE$
       .TextMatrix(5, 1) = DESKRI$
       .TextMatrix(6, 1) = NORFA$
       .TextMatrix(7, 1) = FECHA
       .TextMatrix(8, 1) = CANTIPAQUETE$
       .TextMatrix(9, 1) = CANTI
    End With
      
     IMAG$ = "QR_ETIQPTERM.JPG"
     IDENTIFICADOR$ = "QR_ETIQPTERM"

     On Error Resume Next
     Kill "C:\FLEXOFT\" & IMAG$
     TEXTOQR$ = CODI$ & "|" & NORFA$ & "|" & CANTI & "|" & NC & "|" & CI1% & "|" & NP& & "|" & IDMOVISTO&

     On Error GoTo 0
     'RUTINA QUE LLAMA A APLICACION DE .NET PARA GENERAR LA IMAGEN DEL CODIGO QR
     QRX$ = TRIM$(IMAGENCODBARRA$("C:\FLEXOFT\", IMAG$, TEXTOQR$, "QR_CODE", 0, 0))
     Call T_Espera(3)
'     Image1.Picture = LoadPicture("C:\FLEXOFT\" & IMAG$)
     Call GRABAR_QR("C:\FLEXOFT\" & IMAG$, IDENTIFICADOR$)

    'OJO SI SE LLAMA A PANTALLA A VENTABNEW CUANDO SE ACEPTA BORRA LAS FILAS QUE ESTAN VACIAS
'    VENTABNEW.Show 1
    If CANTIMPRE% < 1 Then CANTIMPRE% = 1
    For i% = 1 To CANTIMPRE%
      Call Param_Print_Excel(RUTAEXCEL$, "ETIQUETAS", VENTABNEW.MSF, PDES$, "1", 1, 7, "", 1, 0, "C:\FLEXOFT\" & IMAG$, 550, 29, 130, 130)                     'imprime en excel la 1,2 Y 3 columna de mi msf en la columna 1 a partir de la fila 11 en el excel (este graba E imprime)
    Next i%
    
    Unload VENTABNEW
    
    'ACTUALIZO LA CANTIDAD DE PAQUETES
    Call ACTUREGISTRO("MOVISTO", "IDCALIDAD", "MOVISTO_ID = " & IDMOVISTO&, CANTIPAQUETE$, REGAF&, "NOMESS")
    
End Sub
Function MetInformados(NORFA$, LOTE$, Optional NOP%)
    
'    SQLX$ = "select sum(corte*cantreal) AS RESULTADO from INFOFAB WITH(NOLOCK) "
'    SQLX$ = SQLX$ + " where idsubor = '" & NORFA$ & "'"
'    SQLX$ = SQLX$ + " and idenlote = '" & LOTE$ & "'"
'    If NOP% > 0 Then SQLX$ = SQLX$ + " and numeoper = " & NOP%
'    SQLX$ = SQLX$ + " and (IdRegPadre is null or IdRegPadre < 1)"
'    SQLX$ = SQLX$ + " and (OperFucionada is null or OperFucionada < 1)"
'
'    Set rs = READSET(SQLX$)
    If NOP% > 0 Then
        SQLX$ = "select sum(corte*CantOrigCte) AS RESULTADO from INFOFAB WITH(NOLOCK) "
        SQLX$ = SQLX$ + " where idsubor = '" & NORFA$ & "'"
        SQLX$ = SQLX$ + " and idenlote = '" & LOTE$ & "'"
        SQLX$ = SQLX$ + " and numeoper = " & NOP%
    Else
        SQLX$ = "select sum(corte*cantreal) AS RESULTADO from INFOFAB WITH(NOLOCK) "
        SQLX$ = SQLX$ + " where idsubor = '" & NORFA$ & "'"
        SQLX$ = SQLX$ + " and idenlote = '" & LOTE$ & "'"
        SQLX$ = SQLX$ + " and (IdRegPadre is null or IdRegPadre < 1)"
        SQLX$ = SQLX$ + " and (OperFucionada is null or OperFucionada < 1)"
    End If
    
    Set RS = READSET(SQLX$)

    MetInformados = XVALO(RS!Resultado)
    
End Function
Function CotiBolsaAprobada%(NROCOTI&)
   CotiBolsaAprobada% = XVALO(VALOBADA("COTIPVTA", "STATUS", "NROCOTIZACION=" & NROCOTI&, "NOMESS"))
End Function

Sub DevoluciondeBobina(MSF1 As msFlexGrid, REG&, NORFA$, DEPOCONSUM%, DEPOALMACEN%, Optional SaldoEnOrigen)
'    Se debe devolver con el stock que tenga en el deposito 2 (como en la otra)
'    con eso tenemos bien los metros,
'    y se debe informar los kg, que tiene la bobina.

    CodigoProducto$ = TRIM$(MSF1.TextMatrix(REG&, 1))
    NOP% = XVALO(MSF1.TextMatrix(REG&, 6))
    DEPOCORTE% = DEPOCONSUM%
    If DEPOCORTE% < 1 Then
       Call COMUNI("Falta Definir Depósito de Corte")
       Exit Sub
    End If
    'LECTURA PESO DESDE BALANZA
    CI1% = XVALO(MSF1.TextMatrix(REG&, 5))
    METROSTRANSFERIDOS = XVALO(MSF1.TextMatrix(REG&, 4))
    LOTE$ = MSF1.TextMatrix(REG&, 3)
    CONDI$ = "COD_INTE = " & CI1%
    CONDI$ = CONDI$ + " AND IDENLOTE = '" & LOTE$ & "'"
    CONDI$ = CONDI$ + " AND DOSECLAR = '" & TRIM(NORFA$) & " '"
    CONDI$ = CONDI$ + " AND DEPOMOVI = " & DEPOCORTE%
    CANTREG& = CantRegistros&("MOVISTO", CONDI$, "NOMESS")
    
'    If CANTREG& Mod 2 = 0 Then
'       Call COMUNI("Atención!! La Bobina Seleccionada Ya ha tenido Ingreso y Salida Para La Orden de Fabricación Activa")
'       Screen.MousePointer = 1
'       Exit Sub
'    End If
    
'   PesoPorMetro# = PESMETRO(CI1%)
    PesoPorMetro# = CoeficientePesoPorMetroSegunLote#(LOTE$)
    If PesoPorMetro# < 0.0005 Then
        Call COMUNI("Falta Configurar Desde la Ficha del Artículo\El Peso por Metro Lineal\del Producto " & CodigoProducto$ & "")
        GoTo 99
    End If
    
'    BALANZA.txtCodigo = CodigoProducto$
'    BALANZA.txtDescripcion = DESCRIT0$(CI1%)
'    BALANZA.Show 1
'    If BALANZA.Respuesta = "APROBO" Then
     PESOLEIDO# = PesoLeidoBalanza(CI1%, 1, OK%)
'        PesoLeido# = XVALO(BALANZA.TXTPESO)
     If OK% > 0 Then
        ' CALCULAR METROS LINEALES
        SaldoBobinaEnMetrosLineales# = PESOLEIDO# / PesoPorMetro#
        SALDO = TRIM$(FORMATNUM$(SaldoBobinaEnMetrosLineales#, "F10.1"))
        
    Else
        GoTo 99
    End If
    '
    
    'QUITE LA VALIDACION SI HABIA 2 MOVIMIENTOS QUE ARRIBA QUEDO COMENTADA
    'AGREGUE ESTE PARRAFO VALIDANDO QUE NO SE PUEDA RETORNAR MAS QUE EL SALDO + UN 5 %
    If PESOLEIDO# < 0.05 Then GoTo 99
    PESOLEID$ = TRIM$(FORMATNUM$(PESOLEIDO#, "F12.2"))
    If (PESOLEIDO#) > (SaldoEnOrigen * 1.05) Then 'si el peso de la balanza es > (+ 5 %) a el saldo de origen
      Call COMUNI("Imposible Retornar Bobina\El Peso a Devolver Es Superior al Saldo de Origen (" & PESOLEID$ & ")")
      GoTo 99
    End If
    
    'caso en que se quiere devolver mas del saldo del lote en el deposito 2
    SaldoDelLote = SaldoLote(CI1%, LOTE$, 2)
    If SaldoBobinaEnMetrosLineales# > SaldoDelLote Then
       SaldoBobinaEnMetrosLineales# = SaldoDelLote
       If SaldoBobinaEnMetrosLineales# > 0.05 Then
          OPX% = COMALTER%("Atenciòn !!!\El saldo del Lote seleccionado es de " & FORMATNUM$(SaldoDelLote, "F9.0") & "\Solo Podrà Regresar a Depòsito: " & FORMATNUM$(SaldoDelLote, "F9.0") & "\\Cancelar\Generar Devoluciòn")
         If OPX% < 2 Then
           GoTo 99
          Else
           GoTo 80 'transfiere el maximo del saldo del lote
         End If
       Else 'si  SaldoBobinaEnMetrosLineales# no tiene saldo
         Call COMUNI("Imposible Transferir Devoluciòn El Còdigo y Lote Activo No Posee Saldo")
         GoTo 99
       End If
    End If
    'TRANSFERENCIA
80  FECHA% = HOY(HO$)
    CMOV$ = "TR"


    NCOMP& = ProNroComprobante(CMOV$)
    DOPRI$ = "TR." + TRIM$(Str$(NCOMP&))
    DOSEC$ = NORFA$
      
    DESCRIDEPO$ = ECOARCH$("TADEPOSI", Chr$(DEPOCONSUM%))
    REFE$ = "Tr.A " & DESCRIDEPO$ & "-" & TXTOF
    NCOMP& = ProNroComprobante(CMOV$)
    Call MOVISTOK(FECHA%, CI1%, LOTE$, CMOV$, DOPRI$, DOPRI$, DOSEC$, DOSEC$, REFE$, 0, 0, 1, NC%, DEPOCONSUM%, (-1) * SaldoBobinaEnMetrosLineales#, , , , , , , , , , NOP%)
    Call T_Espera(1 / 10) 'PARA QUE HAYA DIFERENCIA ENTRE UNA GRABACION Y LA OTRA
    DESCRIDEPO$ = TRIM$(ECOARCH$("TADEPOSI", Chr$(DEPOALMACEN%)))
    REFE$ = "Tr.A " & DESCRIDEPO$ & "-" & TRIM$(TXTOF)
    Call MOVISTOK(FECHA%, CI1%, LOTE$, CMOV$, DOPRI$, DOPRI$, DOSEC$, DOSEC$, REFE$, 0, 0, 1, NC%, DEPOALMACEN%, SaldoBobinaEnMetrosLineales#, , , , , , , , , , NOP%)
    

   
99  Exit Sub
    
End Sub
Sub NuevaDevoluciondeBobina(CIELEM%, NORFA$, LOTE$, DEPOSAL%, DEPOING%, NOP%)
'    Se debe devolver con el stock que tenga en el deposito 2 (como en la otra)
'    con eso tenemos bien los metros,
'    y se debe informar los kg, que tiene la bobina.
    Flag_EsFinBobina% = 0
    CONDI$ = "  IDENLOTE = '" & LOTE$ & "'"
    CONDI$ = CONDI$ + " AND  DEPOMOVI = " & DEPOSAL%
    CONDI$ = CONDI$ + " AND  COD_INTE  = " & CIELEM%

    SQLX$ = " SELECT  SUM(CASE WHEN CANTINGRE > 0 THEN PESOKG  ELSE -PESOKG END) AS PESOKG"
    SQLX$ = SQLX$ + " ,SUM(CANTINGRE - CANTSALID) As SaldoMetros"
    SQLX$ = SQLX$ + " From Movisto WITH(NOLOCK)"
    SQLX$ = SQLX$ + " Where " & CONDI$
    Set RS = READSET(SQLX$)
    If Not RS.EOF Then
       METROS# = XVALO(RS!SaldoMetros)
       'KILOGRAMOS# = XVALO(rs!PESOKG)
    End If
    RS.Close
    Set RS = Nothing
    
    If METROS# <= 0.05 Then
       Call COMUNI("El Registro Seleccionado no Posee Saldo en Depósito de Corte")
       GoTo 99
    End If
    KILOGRAMOS# = PesoLeidoBalanza(CI1%, 1, OK%)
    If OK% < 1 Then GoTo 99
    
    If KILOGRAMOS# <= 0.05 Then
        OPX% = COMALTER%("Atenciòn !!! Muy Importante: \Los Kilogramos Informados son : " & KILOGRAMOS# & "\De Continuar con la Devolución se Registrará un Ajuste Dejando sin Saldo La Bobina\\Cancelar\Generar Devoluciòn Y Ajuste de Fin de Bobina")
        If OPX% < 2 Then GoTo 99
        Flag_EsFinBobina% = 1
    End If

    OPX% = COMALTER%("Atenciòn !!! Situaciòn de Devoluciòn: \Total Metros a Reingresar: " & FORMATNUM$(METROS#, "F9.1") & "\Total Kg a Reingresar: " & FORMATNUM$(KILOGRAMOS#, "F9.1") & "\\Cancelar\Generar Devoluciòn al Depósito: " & DEPOING%)
    If OPX% < 2 Then
       GoTo 99
    End If
   ''INTERVENCION DE ALERTA CUANDO SUPERA EL 5%
    PesoTeoricoSaldoBd# = PesoTeoricoSaldoActual#(LOTE$)
    If CompararValoresEnPorcentaje(PesoTeoricoSaldoBd, KILOGRAMOS#) > 5 Then
        OPX% = COMALTER%("Atención !!\El Peso Informado (" & TRIM$(FORMATNUM$(KILOGRAMOS#, "F8.2")) & " Kg.)\Vs. el Peso del Saldo Registrado (" & TRIM$(FORMATNUM$(PesoTeoricoSaldoBd#, "F8.2")) & " Kg.)\Supera el 5 % de Tolerancia \\Cancelar\Retornar de Todos Modos")
        If OPX% < 2 Then GoTo 99
        If OPX% = 2 Then
           If FIRMAELECTRONICAOK% <= 0 Then GoTo 99 'si no lo aprueba la firma electronica sale, si no sigue
        End If
    End If
    ''FIN INTERVENCION DE ALERTA CUANDO SUPERA EL 5%
      

    'TRANSFERENCIA
80  FECHA% = HOY(HO$)
    CMOV$ = "TR"

    DEPOCONSUM% = XVALO(DEPOSAL%)
    DEPOALMACEN% = XVALO(DEPOING%)
    SALDOENMETROS# = METROS#
    SALDOENKGRAMOS# = KILOGRAMOS#
    NCOMP& = ProNroComprobante(CMOV$)
    DOPRI$ = "TR." + TRIM$(Str$(NCOMP&))
    DOSEC$ = NORFA$
      
    DESCRIDEPO$ = ECOARCH$("TADEPOSI", Chr$(DEPOCONSUM%))
    REFE$ = "Tr.A " & DESCRIDEPO$ & "-" & TXTOF
    NCOMP& = ProNroComprobante(CMOV$)
    On Error GoTo ERROR_GUARDAR
    FLEXCONN.BeginTrans

    Call MOVISTOK(FECHA%, CIELEM%, LOTE$, CMOV$, DOPRI$, DOPRI$, DOSEC$, DOSEC$, REFE$, 0, 0, 1, NC%, DEPOCONSUM%, (-1) * SALDOENMETROS#, , , , , , , , , , NOP%, SALDOENMETROS#, SALDOENKGRAMOS#)
    Call T_Espera(1 / 10) 'PARA QUE HAYA DIFERENCIA ENTRE UNA GRABACION Y LA OTRA
    DESCRIDEPO$ = TRIM$(ECOARCH$("TADEPOSI", Chr$(DEPOALMACEN%)))
    REFE$ = "Tr.A " & DESCRIDEPO$ & "-" & TRIM$(TXTOF)
    '
    Call MOVISTOK(FECHA%, CIELEM%, LOTE$, CMOV$, DOPRI$, DOPRI$, DOSEC$, DOSEC$, REFE$, 0, 0, 1, NC, DEPOALMACEN%, SALDOENMETROS#, , , , , , , , , , NOP%, SALDOENMETROS#, SALDOENKGRAMOS#)
    'primero veo si tengo que grabar los kg de en el movimiento br del lote
    
    Call ActualizarPesoEnBrdeMovisto(LOTE$)
    pesoSegunLote# = ROUND(PesoOriginalSegúnLote#(LOTE$), 2)
    
    'validar si es unidad de medida m.
    If Unimed(CIELEM%) <> "M." Then Exit Sub
    'obtengo el peso teorico segun lo que esta en la tabla, que es lo que fui descontando en cada consumo del total
    
    PesoTeoricoSaldoBd# = PesoTeoricoSaldoActual#(LOTE$)
    CONDI$ = "IDENLOTE = '" & LOTE$ & "'"
    MetrosSegunLoteIniciales# = ROUND(XVALO(VALOBADA("BOLRECEP", "CantNom", CONDI$, "NOMESS")), 2)
    PesoXMetro# = pesoSegunLote# / MetrosSegunLoteIniciales#
    
    'comparo los kg que informaron en el retorno en este mismo retorno

    'ajuste en kg = pesoInfoEnvap - PesoTeoricoSaldoBd#  en este orden si es negativo va a ser cantsalid de mts else cantingre mts
    If CompararValoresEnPorcentaje(PesoTeoricoSaldoBd, KILOGRAMOS#) > 1 Then
       'si es diferente obtengo la diferencia
        ajusteEnKG# = KILOGRAMOS# - PesoTeoricoSaldoBd#
        ajusteEnMts# = ajusteEnKG# / PesoXMetro# 'EN MOVISTOK EL PESOKG Y LOS METROS LO GUARDO CON ABS()
        If Flag_EsFinBobina% > 0 Then 'si es fin de bobina pesokg = 0
          If ajusteEnMts# < 0 Then    'siempre deberia ser negativo, pero lo pongo del lado que corresponda
           ajusteEnMts# = -1 * METROS# 'para que ajuste por la misma cantidad del saldo
          Else
           ajusteEnMts# = METROS#
          End If
        End If
        DOCUNUM& = ProNroComprobante("DR")
        DOPRILA$ = "DR." & Str$(DOCUNUM&)
        DOPRICO$ = DOPRILA$
        DOSELA$ = NORFA$
        DOSECO$ = DOSELA$
        REFE$ = "Dif en Retorno " & NORFA$
        FF% = HOY(HO$)
        DEP% = DEPOALMACEN%
        '
        Call MOVISTOK(FF%, CIELEM%, LOTE, "DR", DOPRILA$, DOPRICO$, DOSELA$, DOSECO$, REFE$, 1, 0, 1, 0, DEP%, ajusteEnMts#, 0, 0, "", "", "", 0, , , , NOP%, Abs(ajusteEnMts#), Abs(ajusteEnKG#))
        '
    End If
ERROR_GUARDAR:
    If Err <> 0 Then
       FLEXCONN.RollbackTrans
       On Error GoTo 0
       Call COMUNI("ERROR AL GUARDAR NO SE REALIZO LA TRANSACCION")
       Exit Sub
    Else
      FLEXCONN.CommitTrans
      Call COMUNI("Grabación Realizado Exitosamente")
    End If
    
99  Exit Sub



End Sub

Function FIRMAELECTRONICAOK%()
     FIRMAELEC.Text1 = ""
     FIRMAELEC.Text2 = ""
     FIRMAELEC.Label1 = ""
     FIRMAELECTRONICAOK% = -1
     FIRMAELEC.Show 1

     If FIRMAELEC.ESTADO > 0 Then
        FIRMAELECTRONICAOK% = 1
        Exit Function
     Else
        If NIVCONFI%(XVALO(FIRMAELEC.Label1)) < 3 Then
         Call COMUNI("Usuario Con Nivel de Clasificación Inválido Para Esta Operación")
         FIRMAELECTRONICAOK% = -1
         Exit Function
        End If
     End If

End Function

Function CompararValoresEnPorcentaje(A As Double, B As Double) As Double
    Dim Diferencia As Double
    Dim promedio As Double
    Dim porcentajeDiferencia As Double

    If A = B Then
        CompararValores = 0
        Exit Function
    End If
    '60 Y 58
    Diferencia = Abs(A - B) '2
    promedio = (A + B) / 2  '59
    porcentajeDiferencia = (Diferencia / promedio) * 100

    CompararValoresEnPorcentaje = ROUND(porcentajeDiferencia, 2)
End Function
Function SectorSinParentesis$(NOMBRESECTOR$)
   'PASA QUE PUEDE VENIR EL SECTOR COMO UPANEL O UPANEL [CM] POR QUE FUE UN CAMBIO QUE SE PIDIO EN EL MEDIO
   'Y QUEDE ATRAPADO EN MEDIO DE DOS CIRCUNSTANCIAS
   NOMBRESECTOR$ = TRIM$(NOMBRESECTOR$)
   HastaParentesis% = InStr(NOMBRESECTOR$, "(")
   HastaCorchete% = InStr(NOMBRESECTOR$, "[")
   If HastaParentesis% > 0 Then
     NOMBRESECTOR$ = TRIM$(Mid$(NOMBRESECTOR$, 1, HastaParentesis% - 1))
     LongNombreSector% = Len(NOMBRESECTOR$)
   ElseIf HastaCorchete% > 0 Then
     NOMBRESECTOR$ = TRIM$(Mid$(NOMBRESECTOR$, 1, HastaCorchete% - 1))
     LongNombreSector% = Len(NOMBRESECTOR$)
   Else
     LongNombreSector% = Len(NOMBRESECTOR$)
   End If
   SectorSinParentesis$ = Left(NOMBRESECTOR$, LongNombreSector%)

End Function
Function SectorSinParentesisConCorchetes$(NOMBRESECTOR$)
   'LE QUITA LOS PARENTESIS PERO NO LOS CORCHETES
   NOMBRESECTOR$ = TRIM$(NOMBRESECTOR$)
   HastaParentesis% = InStr(NOMBRESECTOR$, "(")
   If HastaParentesis% > 0 Then
     NOMBRESECTOR$ = TRIM$(Mid$(NOMBRESECTOR$, 1, HastaParentesis% - 1))
     LongNombreSector% = Len(NOMBRESECTOR$)
   Else
     LongNombreSector% = Len(NOMBRESECTOR$)
   End If
   
   SectorSinParentesisConCorchetes$ = Left(NOMBRESECTOR$, LongNombreSector%)

End Function

Function APLICARCONSUMO%(NORFA$, DescripOper$, CANTI_ACERRAR, Optional ValorIdcoordubic%)
'****  tener en cuenta la operacion como filtro *****
'ESTA FUNCION DEVUELVE 1 SI PUDO APLICAR EL CONSUMO A LA OPERACION PARAMETRIZADA
'RECORRE LA TABLA INFOFAB REVISANDO LOS INFORMES DE CORTES QUE ESTAN EN LA LINEA
'Y REGISTRA EN EL CAMPO CONSUMO LA CANTIDAD, EN LA MEDIDA QUE CONSUMO TENGA UN SALDO (cantreal - CONSUMO)
'VALIDO PARA BOLSONES
       If ValorIdcoordubic < 1 Then ValorIdcoordubic = 100 'si es vacio queda como estaba en 100(linea), si no
        APLICARCONSUMO% = 0                                'deberia venir con valor 400 de entrega taller
        CIFAB% = CODIFAB2%(NORFA$)
        CantDeCierre = CANTI_ACERRAR
        'ACA TENGO LOS SALDOS DE CADA INFORME DE FABRICACION
        SQLX$ = "SELECT   I.DescOper"
        SQLX$ = SQLX$ + "  ,I.IdenLote,(I.cantreal - I.CONSUMO) AS SALDO,CONSUMO, IDREGPADRE"
        SQLX$ = SQLX$ + " FROM infofab I "
        
        SQLX$ = SQLX$ + " WHERE NOT EXISTS (SELECT * FROM INFOFAB AS I2"
        SQLX$ = SQLX$ + " WHERE I.INFOFAB_ID = I2.IDREGPADRE)"
        
        SQLX$ = SQLX$ + " AND idcoordubic = " & ValorIdcoordubic%
        SQLX$ = SQLX$ + " and IdSubOr = '" & NORFA$ & "'"
        SQLX$ = SQLX$ + " AND left(I.DescOper," & Len(DescripOper$)
        SQLX$ = SQLX$ + " )= '" & DescripOper$ & "'"
        SQLX$ = SQLX$ + " order by saldo "
        

        Dim RS As ADODB.Recordset
        Set RS = New ADODB.Recordset
  
        RS.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
        With RS
          Do While Not .EOF
            SaldoActual = XVALO(!SALDO)
'MsgBox "SALDO ACTUAL DE " & DescripOper$ & ": " & SaldoActual
            If SaldoActual >= CantDeCierre Then
               Aconsumir = CantDeCierre
            Else
               Aconsumir = SaldoActual
            End If
            CantDeCierre = CantDeCierre - Aconsumir
'MsgBox "A Consumir (Consumo Previo: " & XVALO(!CONSUMO) & ") " & DescripOper$ & ": " & Aconsumir
            !CONSUMO = XVALO(!CONSUMO) + Aconsumir
            .Update
            ConsumoAplicado = ConsumoAplicado + Aconsumir
            If ConsumoAplicado >= CANTI_ACERRAR Then
               APLICARCONSUMO% = 1
               Exit Do 'si ya cumplio con la asignacion del consumo sale
            End If
            .MoveNext
          Loop
        End With
        RS.Close
        Set RS = Nothing
        
'MsgBox "LLEGO AL FINAL"
'ACTUALIZA EL CONSUMO EN LA TAPA SOLO ESTO ES PARA CUANDO HAY FUCION DE TAPA CON TAPA COMPLETA
        If TRIM$(UCase$(DescripOper$)) = "TAPA" Then
           SQLX$ = " SELECT IDREGPADRE,CONSUMO FROM INFOFAB WITH(NOLOCK)"
           SQLX$ = SQLX$ + " WHERE DESCOPER = 'TAPA COMPLETA'"
           SQLX$ = SQLX$ + " AND IDSUBOR = '" & NORFA$ & "'"
           SQLX$ = SQLX$ + " AND IDREGPADRE > 0 "
           Set rs1 = READSET(SQLX$)
           With rs1
              Do While Not .EOF
                 CONSUMODELREGISTRO = XVALO(!CONSUMO)
                 IDREGPA& = XVALO(!IDREGPADRE)
                 SQLX$ = "UPDATE INFOFAB SET CONSUMO = " & CONSUMODELREGISTRO
                 SQLX$ = SQLX$ + " WHERE INFOFAB_ID = " & IDREGPA&
                 FLEXCONN.Execute (SQLX$)
                 .MoveNext
              Loop
           End With
       End If

End Function

Function NuevaAPLICARCONSUMO%(NORFA$, DescripOper$, CANTI_ACERRAR, NUMEOPER%, idDelMovisto, Optional ValorIdcoordubic%)
'****  tener en cuenta la operacion como filtro *****
'ESTA FUNCION DEVUELVE 1 SI PUDO APLICAR EL CONSUMO A LA OPERACION PARAMETRIZADA
'RECORRE LA TABLA INFOFAB REVISANDO LOS INFORMES DE CORTES QUE ESTAN EN LA LINEA
'Y REGISTRA EN EL CAMPO CONSUMO LA CANTIDAD, EN LA MEDIDA QUE CONSUMO TENGA UN SALDO (cantreal - cantsalid)
'VALIDO PARA BOLSONES

       'idMov& es el que se genero cuando se dio de alta en el movisto el movimiento OF (INGRESO DE P.TERMINADO)
        
       If ValorIdcoordubic < 1 Then ValorIdcoordubic = 100 'si es vacio queda como estaba en 100(linea), si no
        NuevaAPLICARCONSUMO% = 0                                'deberia venir con valor 400 de entrega taller
        CIFAB% = CODIFAB2%(NORFA$)
        CantDeCierre = CANTI_ACERRAR
        'ACA TENGO LOS SALDOS DE CADA INFORME DE FABRICACION
        
10      SQLX$ = "SELECT  SUM(CASE WHEN cantreal > 0 THEN PESO  ELSE -PESO END) AS kg,"
        SQLX$ = SQLX$ + " SUM(CASE WHEN cantreal > 0 THEN cantreal  ELSE -cantsalid END) AS Saldo,"
        SQLX$ = SQLX$ + " COUNT(*) AS CANTREG"
        SQLX$ = SQLX$ + " ,cod_inte,idsubor,Cod_Exte,descrip,numeoper,descoper,idenlote,codielem,CodXElem,FACTORCORTE,corte"
        SQLX$ = SQLX$ + " From InfoFab WITH(NOLOCK)"
        SQLX$ = SQLX$ + " WHERE    IDSUBOR = '" & NORFA$ & "' AND IDCOORDUBIC = " & ValorIdcoordubic%
        SQLX$ = SQLX$ + "  and LEFT(DescOper,8) <> 'CINTILLA' and LEFT(DescOper,4) <> 'HILO'"
        SQLX$ = SQLX$ + " AND left(DescOper," & Len(DescripOper$)
        SQLX$ = SQLX$ + " )= '" & DescripOper$ & "'"
        SQLX$ = SQLX$ + "  group by cod_inte,idsubor,Cod_Exte,descrip,numeoper,descoper,idenlote,codielem"
        SQLX$ = SQLX$ + "  ,corte ,CodXElem,FACTORCORTE"
        SQLX$ = SQLX$ + "  HAVING SUM(CASE WHEN cantreal > 0 THEN cantreal  ELSE -cantsalid END)> 0"
        SQLX$ = SQLX$ + "  order by  NUMEOPER,Cod_inte,IDENLOTE"
        
        Set RS = READSET(SQLX$)
  
'        RS.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
        With RS
          Do While Not .EOF
            SaldoActual = XVALO(!SALDO)
'MsgBox "SALDO ACTUAL DE " & DescripOper$ & ": " & SaldoActual
            If SaldoActual >= CantDeCierre Then
               Aconsumir = CantDeCierre
            Else
               Aconsumir = SaldoActual
            End If
            CantDeCierre = CantDeCierre - Aconsumir
            CI1% = XVALO(!cod_inte)
            NORFA$ = SAFETEXT$(!idsubor)
            Codigo$ = SAFETEXT$(!Cod_Exte)
            DESCRIX$ = SAFETEXT$(!Descrip)
            NOPER% = XVALO(!NUMEOPER)
            DesxOper$ = SAFETEXT$(!descoper)
            LOTE$ = SAFETEXT$(!idenlote)
            CIELEM% = XVALO(!codielem)
            CodExtElem = SAFETEXT$(!CodXElem)
            FACTORCORTE = FORMATNUM$(XVALO(!FACTORCORTE), "F16.4")
            XORTE = FORMATNUM$(XVALO(!corte), "F16.4")
            Observa$ = "Cierre X " & CANTI_ACERRAR / FACTORCORTE & " el Día: " & FECHATEX$(HOY(HO$))
            REFERENCIA$ = "Cierre X " & CANTI_ACERRAR & " el Día: " & FECHATEX$(HOY(HO$))
            idMov& = idDelMovisto
            Aconsumir = FORMATNUM$(XVALO(Aconsumir), "F16.4")

            Call ConsumoInfofab(HOY(HO$), Aconsumir, CI1%, NORFA$, Codigo$, DESCRIX$, NOPER%, DesxOper$, LOTE$, CIELEM%, CodExtElem, FACTORCORTE, XORTE, Observa$, REFERENCIA$, idMov&)
'            !CONSUMO = XVALO(!CONSUMO) + Aconsumir
'            .Update
            ConsumoAplicado = ConsumoAplicado + Aconsumir
            If ConsumoAplicado >= CANTI_ACERRAR Then
               NuevaAPLICARCONSUMO% = 1
               Exit Do 'si ya cumplio con la asignacion del consumo sale
            Else
               GoTo 10
            End If
            
            .MoveNext
          Loop
        End With
        RS.Close
        Set RS = Nothing
        
        If ConsumoAplicado < CANTI_ACERRAR Then 'SI NO PUDO APLICAR TODO
           CantiSinAplicar = CANTI_ACERRAR - XVALO(ConsumoAplicado)
           idMov& = idDelMovisto
           CONDI$ = "NUMEOPER = " & NUMEOPER% & " AND IDSUBOR ='" & NORFA$ & "' AND COD_INTE = " & CIFAB%
           XXORTE = FORMATNUM$(XVALO(VALOBADA("INFOFAB", "CORTE", CONDI$, "NOMESS")), "F16.4")
           If XVALO(XXORTE) <= 0 Then
              XXORTE = "0"
           End If
           FACTORXORTE = XVALO(VALOBADA("INFOFAB", "FACTORCORTE", CONDI$, "NOMESS"))
           CantiSinAplicar = FORMATNUM$(XVALO(CantiSinAplicar), "F16.4")

           Call ConsumoInfofab(HOY(HO$), CantiSinAplicar, CIFAB%, NORFA$, "", "", NUMEOPER%, DescripOper$, "", 0, "", FACTORXORTE, XXORTE, "SinSaldo", "SinSaldo", idMov&)
        End If
'MsgBox "LLEGO AL FINAL"
'ACTUALIZA EL CONSUMO EN LA TAPA SOLO ESTO ES PARA CUANDO HAY FUCION DE TAPA CON TAPA COMPLETA
        If TRIM$(UCase$(DescripOper$)) = "TAPA" Then
           SQLX$ = " SELECT IDREGPADRE,CONSUMO FROM INFOFAB WITH(NOLOCK)"
           SQLX$ = SQLX$ + " WHERE DESCOPER = 'TAPA COMPLETA'"
           SQLX$ = SQLX$ + " AND IDSUBOR = '" & NORFA$ & "'"
           SQLX$ = SQLX$ + " AND IDREGPADRE > 0 "
           Set rs1 = READSET(SQLX$)
           With rs1
              Do While Not .EOF
                 CONSUMODELREGISTRO = XVALO(!CONSUMO)
                 IDREGPA& = XVALO(!IDREGPADRE)
                 SQLX$ = "UPDATE INFOFAB SET CONSUMO = " & CONSUMODELREGISTRO
                 SQLX$ = SQLX$ + " WHERE INFOFAB_ID = " & IDREGPA&
                 FLEXCONN.Execute (SQLX$)
                 .MoveNext
              Loop
           End With
       End If

End Function


Sub DescontarConsumoPorAnulacion(NORFA$, DescripOper$, CANTI_ACERRAR)

        
        CIFAB% = CODIFAB2%(NORFA$)
        CantDeCierre = CANTI_ACERRAR
        'ACA TENGO LOS SALDOS DE CADA INFORME DE FABRICACION
        SQLX$ = "SELECT   I.DescOper"
        SQLX$ = SQLX$ + "  ,I.IdenLote,(I.cantreal - I.CONSUMO) AS SALDO,CONSUMO"
        SQLX$ = SQLX$ + " FROM infofab I "
        
        SQLX$ = SQLX$ + " WHERE NOT EXISTS (SELECT * FROM INFOFAB AS I2"
        SQLX$ = SQLX$ + " WHERE I.INFOFAB_ID = I2.IDREGPADRE)"
        
        SQLX$ = SQLX$ + " AND (idcoordubic = 100 or idcoordubic = 300)" 'por si esta depurado
        SQLX$ = SQLX$ + " and IdSubOr = '" & NORFA$ & "'"
        SQLX$ = SQLX$ + " AND left(I.DescOper," & Len(DescripOper$)
        SQLX$ = SQLX$ + " )= '" & DescripOper$ & "'"
        SQLX$ = SQLX$ + " order by idcoordubic asc ,Consumo desc"
        Dim RS As ADODB.Recordset
        Set RS = New ADODB.Recordset
        RS.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
        With RS
          Do While Not .EOF
            CONSUMODELREGISTRO = XVALO(!CONSUMO)
            If CONSUMODELREGISTRO >= CantDeCierre Then 'si lo que tiene descontar alcanza con el valor del consumo
               Adescontar = CantDeCierre
               !CONSUMO = XVALO(!CONSUMO) - Adescontar
               .Update
               Exit Do
            Else 'si no alcanza pone el consumo en 0
               Adescontar = CONSUMODELREGISTRO
               !CONSUMO = 0
               .Update
            End If
            'va actualizando el saldo descontado de la cantidad de cierre.
            CantDeCierre = CantDeCierre - Adescontar
        
            .MoveNext
          Loop
        End With
        RS.Close
        Set RS = Nothing



End Sub

Sub RecorrerOperacionesSegunOf(NORFA$, CANTI_ACERRAR, NROCOTI&, Optional IDCOORDUBIC%)
'EN EL CASO QUE SEA DE TALLER 'IDCOORDUBIC' VIENE CON VALOR 400
'RECORRE LA ESTRCUTURA SEGUN COTIZACIN EN TABLA COTI1
'Y LLAMA A FUNCION APLICARCONSUMO% QUE APLICA EN INFOFAB LA CANTIDAD CONSUMIDA
        If NROCOTI& < 1 Then Exit Sub
        Dim CTP As New Conteplast
        If CANTI_ACERRAR < 1 Then CANTI_ACERRAR = 1
        SQLX$ = "SELECT C.CODIGO, C.DESCRIPCION, C.COMPONENTE,C.CODINTELEM, C.MODUXBB,C.CORTE"
        SQLX$ = SQLX$ + ", SUM(C.MODUXBB * C.CORTE * "
        SQLX$ = SQLX$ & CANTI_ACERRAR
        SQLX$ = SQLX$ + ") AS NECESIDAD_EN_M" 'ESTO EN PRINCIPIO NO LO UTILIZO LO DEJO POR LAS DUDAS
        
        SQLX$ = SQLX$ + ", SUM(C.MODUXBB * "
        SQLX$ = SQLX$ & CANTI_ACERRAR
        SQLX$ = SQLX$ + ") AS NECESIDAD_DE_CORTES"
        
        SQLX$ = SQLX$ + " FROM COTI1 C  WITH(NOLOCK)"
        SQLX$ = SQLX$ + "  WHERE CODINTELEM > 0 "
        SQLX$ = SQLX$ + " And NROCOTIZACION = " & NROCOTI&
        SQLX$ = SQLX$ + " GROUP BY CODIGO, C.DESCRIPCION, C.COMPONENTE,C.CODINTELEM, C.MODUXBB,C.CORTE"

        Set RS = READSET(SQLX$)
        With RS
          Do While Not .EOF
            Descrioper$ = SAFETEXT$(!COMPONENTE)
            Descrioper$ = TRIM$(CTP.SectorSinParentesis$(Descrioper$))

            CantCortesCalculadosNecesarios = XVALO(!NECESIDAD_DE_CORTES)
'            APLICA EL CONSUMO EN INFOFAB
            APLICONSUMO% = APLICARCONSUMO%(NORFA$, Descrioper$, CantCortesCalculadosNecesarios, IDCOORDUBIC%)
            .MoveNext
          Loop
        End With
        RS.Close: Set RS = Nothing

End Sub
Function DescripcionyCoordenadasUbicacion$(CORDENADAUBICACION$)
'recibe ej 3;4;8
'devuelve Producto Terminado 3;4;8
     If XVALO(CORDENADAUBICACION$) > 0 Then
        COORDENADANUM$ = CORDENADAUBICACION$
        CONDI$ = "CONVERT(int,campo1) = " & XVALO(CORDENADAUBICACION)
        CONDI$ = CONDI$ + " AND LISTA = 'CONFUBIX'"
        DESCORDENADAUBICACION$ = VALOBADA("DATASQL", "CAMPO2", CONDI$, "NOMESS")
      End If
      DescripcionyCoordenadasUbicacion$ = DESCORDENADAUBICACION$ & " " & COORDENADANUM
End Function
Sub RecuperarConsumo_x_AnulaciónDeCierre(NROOF$, CANTI_ACERRAR, NROCOTI&)
        If NROCOTI& < 1 Then Exit Sub
        Dim CTP As New Conteplast
        CANTI_ACERRAR = TRIM$(FORMATNUM$(CANTI_ACERRAR, "f12.1"))
        SQLX$ = "SELECT C.CODIGO, C.DESCRIPCION, C.COMPONENTE,C.CODINTELEM, C.MODUXBB,C.CORTE"
        SQLX$ = SQLX$ + ", SUM(C.MODUXBB * C.CORTE * "
        SQLX$ = SQLX$ & CANTI_ACERRAR
        SQLX$ = SQLX$ + ") AS NECESIDAD_EN_M" 'ESTO EN PRINCIPIO NO LO UTILIZO LO DEJO POR LAS DUDAS
        
        SQLX$ = SQLX$ + ", SUM(C.MODUXBB * "
        SQLX$ = SQLX$ & CANTI_ACERRAR
        SQLX$ = SQLX$ + ") AS NECESIDAD_DE_CORTES"
        
        SQLX$ = SQLX$ + " FROM COTI1 C  WITH(NOLOCK)"
        SQLX$ = SQLX$ + "  WHERE CODINTELEM > 0 "
        SQLX$ = SQLX$ + " And NROCOTIZACION = " & NROCOTI&
        SQLX$ = SQLX$ + " GROUP BY CODIGO, C.DESCRIPCION, C.COMPONENTE,C.CODINTELEM, C.MODUXBB,C.CORTE"

        Set RS = READSET(SQLX$)
        With RS
          Do While Not .EOF
            Descrioper$ = SAFETEXT$(!COMPONENTE)
            Descrioper$ = TRIM$(CTP.SectorSinParentesis$(Descrioper$))
            CantCortesCalculadosNecesarios = XVALO(!NECESIDAD_DE_CORTES)
'            APLICA EL CONSUMO EN INFOFAB
            Call DescontarConsumoPorAnulacion(NROOF$, Descrioper$, CantCortesCalculadosNecesarios)
            .MoveNext
          Loop
        End With
        RS.Close: Set RS = Nothing

        Set CTP = Nothing




End Sub

Function SaldoEnLineaCortesInfofab(NORFA$, Descrioper$)
'        DEVUELVE EL SALDO DE UNA SECTOR AGRUPADO (EJEMPLO SI ES LATERALES Y LATERALES-SEMITERMINADOS)
'        CONSIDERA LA CANTIDAD INFORMADA MENOS LOS CONSUMOS INFORMADOS EN LOS CIERRES DE OF.
        'FUNCTION ISNULL ES PARA QUE SI ES NULL LO CONSIDERE COMO 0
        SectSinParentesis$ = SectorSinParentesis$(Descrioper$)
        CantCaracteres% = Len(SectSinParentesis$)
        'ESTA CONSULTA PRIMERO REVISA SI TIENE UNA OPERACION FUCIONADA
        SQLX$ = "SELECT OperFucionada,left(descoper,"
        SQLX$ = SQLX$ & CantCaracteres% & ") "
        SQLX$ = SQLX$ + " FROM infofab I  WITH(NOLOCK)"
        SQLX$ = SQLX$ + " WHERE "
        SQLX$ = SQLX$ + " idcoordubic = 100 and IdSubOr = '" & NORFA$ & "'"
        SQLX$ = SQLX$ + " and left(descoper," & CantCaracteres%
        SQLX$ = SQLX$ + " )= '" & SectSinParentesis$ & "'"
        Set rs1 = READSET(SQLX$)
    
        If Not rs1.EOF Then
             NOPERFUCION& = XVALO(rs1!OperFucionada) 'SI TIENE OPERACION FUCIONADA
        End If                                       'BUSCA EL SALDO DE LA OP.FUCIONADA EJ. TAPA COMPLETA
        rs1.Close                                    'Y TOMA EL SALDO DE ESTA PARA VALIDAR. Y NO SIGUE
        Set rs1 = Nothing
        If NOPERFUCION& > 0 Then
           SaldoOperacionFucionada = XVALO(VALOBADA("INFOFAB", "SUM(ISNULL(cantreal,0) - ISNULL(CONSUMO,0))", "INFOFAB_ID = " & NOPERFUCION&, "NOMESS"))
           SaldoEnLineaCortesInfofab = SaldoOperacionFucionada
          GoTo 99
        End If
      '============================================================================================
       SQLX$ = "SELECT   FACTORCORTE,I.CORTE,left(descoper,"
        SQLX$ = SQLX$ & CantCaracteres% & ") ,SUM(ISNULL(I.cantreal,0) - ISNULL(I.CONSUMO,0)) AS SALDO"
        SQLX$ = SQLX$ + " FROM infofab I  WITH(NOLOCK)"
        SQLX$ = SQLX$ + " WHERE NOT EXISTS (SELECT * FROM INFOFAB AS I2 WHERE I.INFOFAB_ID = I2.IDREGPADRE)"
        SQLX$ = SQLX$ + " AND idcoordubic = 100 and IdSubOr = '" & NORFA$ & "'"
        SQLX$ = SQLX$ + " and left(descoper," & CantCaracteres%
        SQLX$ = SQLX$ + " )= '" & SectSinParentesis$ & "'"
        SQLX$ = SQLX$ + " group by FACTORCORTE,I.CORTE,left(descoper,"
        SQLX$ = SQLX$ & CantCaracteres%
        SQLX$ = SQLX$ + " )"
        Set RS = READSET(SQLX$)
       With RS
          If Not .EOF Then
             SaldoEnLineaCortesInfofab = XVALO(RS!SALDO)
          Else
             SaldoEnLineaCortesInfofab = 0
          End If
        End With
        RS.Close
        Set RS = Nothing
99
End Function
Function ValidarNecesidadCompBolsas(NORFA$, DEPOCORTE%, CantiaCerrar)
    'DEVUELVE -1 SI FALTA DISPONIBILIDAD DE MATERIAL TRANSFERIDO PARA LA ORDEN REVISANDO POR CODIGO
    CIFAB% = CODIFAB2%(NORFA$)
    
    SQLX$ = "SELECT * FROM FORMULAS WITH(NOLOCK) WHERE CodIConj=" & CIFAB%
    Set RstFormula = READSET(SQLX$)
    J& = 0
    With RstFormula
        If Not (.EOF And .BOF) Then
            Do While Not .EOF
                CodigoIntComponentes% = XVALO(!codielem)
                CodigoExtComponentes$ = SAFETEXT$(!CodXElem)
                CantidadUnitariaComponente = XVALO(!usobruto)
                NECESIDAD = CantiaCerrar * CantidadUnitariaComponente

                SALDO = DisponibleXOrdenyCódigoCTP(NORFA$, CodigoIntComponentes%, DEPOCORTE%)
                .MoveNext
            Loop
        End If
    End With
    RstFormula.Close
    Set RstFormula = Nothing
End Function
Function DisponibleXOrdenyCódigoCTP(NORFA$, CodigoIntComponentes%, DEPOCORTE%)
    Dim CTP As New Conteplast
    'TENER EN CUENTA QUE SI EL CODIGO QUE SE PASA ES UN CODIGO RANGO
    'VA A DEVOLVER EL ACUMULADOD DE TODOS LOS QUE DENTRO DE LA CONSULTA CORRESPONDAN AL CODIGO RANGO
    
    CONDI$ = "COD_INTE = " & CodigoIntComponentes%
    If ESCODIGORANGO&(CodigoIntComponentes%) > 0 Then
       CONDI$ = CTP.CONSULTA_OR("CODIGORANGO", "CODMATPRIMA", CONDI$, "COD_INTE")
    End If
    
    SQLX$ = " SELECT SUM(CANTINGRE-CANTSALID) AS RESULTADO ,COD_INTE,Cod_Exte "
    SQLX$ = SQLX$ + " FROM MOVISTO WITH(NOLOCK)"
    SQLX$ = SQLX$ + " where (DOSECLAR = '" & NORFA$ & "' or doprilar = '" & NORFA$ & "') "
    SQLX$ = SQLX$ + " AND DEPOMOVI = " & DEPOCORTE%
    SQLX$ = SQLX$ + " AND " & CONDI$
    SQLX$ = SQLX$ + " GROUP BY COD_INTE"
    Set RS = READSET(SQLX$)
    With RS
      Do While Not .EOF
        DisponibleXOrdenyCódigoCTP = DisponibleXOrdenyCódigoCTP + XVALO(Resultado)
        .MoveNext
      Loop
    End With
     
    Set CTP = Nothing
End Function

Function CantiTransferida_X_OF_Y_CodigosAsociadosalRango#(NORFA$, CONDIOR$, Optional CODIRANGO$)
'CantiTransferida_X_OF_Y_CodigosAsociadosalRango# devuelve la canti transferida para el param NORFA$ Y LOS CODIGOS REALES DEL RANGO

   CONDIOR$ = TRIM$(CONDIOR$)
   Rango$ = TRIM$(Rango$)
   CantiTransferida_X_OF_Y_CodigosAsociadosalRango# = 0
   If CONDIOR$ = "" And Rango$ = "" Then Exit Function 'SI NO SE LE PASA LA CONDIOR NI EL CODIGO RANGO SE VA
   If CONDIOR$ = "" Then 'SI NO TIENE CONDIOR LA ARMA CON EL CODIGO RANGO
      If ESCODIGORANGO&(CODINT%(CODIRANGO$)) > 0 Then 'VERIFICA QUE SEA CODIGO RANGO
        CONDI$ = "CODRANGO = '" & CODIRANGO$ & "'"
        CONDIOR$ = CONSULTA_OR("CODIGORANGO", "CODMATPRIMA", CONDI$, "COD_EXTE")
      End If
   End If
   
   SQLX$ = "SELECT  sum(CANTSALID-CANTINGRE) AS CANTI  FROM MOVISTO WITH(NOLOCK)"
   SQLX$ = SQLX$ + " WHERE CODIMOVI = 'TR'  and depomovi = 1"
   SQLX$ = SQLX$ + " AND  DOSECLAR = '" & NORFA$ & "'"
   If CONDIOR$ <> "" Then SQLX$ = SQLX$ + " AND " & CONDIOR$
'   Call mostrarDatosx("2 " & SQLX$)

   Set RS = READSET(SQLX$)
   With RS
     If .EOF <> True Then
        CantiTransferida_X_OF_Y_CodigosAsociadosalRango# = XVALO(!CANTI)
     End If
   End With
   RS.Close
   Set RS = Nothing
   
   
End Function
Sub ACTUBASEWALDBOTT(Optional NoBorrarArchivoOrigen%)
   Call GRACONTROL("BASEWALD", "ESTAGRAB", "NO") 'LO INICIALIZO EN NO POR LAS DUDAS
'ATENCION (ACTUBASEWALDBOTT_ANT_DESCOMPRIME) TIENE EL CODIGO PARA TOMAR UN ARCHIVO ZIP O RAR  Y DESCOMPRIMIRLO EN OTRA CARPETA
   If Control("BASEWALD", "ESTAGRAB") = "SI" Then Exit Sub
   On Error GoTo ERROR_GUARDAR
   RUTADROP$ = Control("COTICONT", "ARCDROPB")
   NAMEARCHIVO$ = "SGIW14.MDB" ' BUSCOARCHIVO$("SGIW14.MDB")
   RUTADESTINO$ = Control("CONWALDB", "RUTAWALD")
'MsgBox "RUTADROOP: " & RUTADROP$  '\\PC10\uSERS\Omar\Dropbox\IMPOR_IMPLANSOFT
'MsgBox "NAMEARCHIVO$: " & NAMEARCHIVO$  'ESTO NO LO ENCONTRO VACIO
'MsgBox "RUTADESTINO$: " & RUTADESTINO$  'X:\FLEXOFT\WALDBOT\SGIW11.MDB
   'comparacion de fechas de archivo
   FectaArchivoDrop# = FILEDATE#(RUTADROP$ & "\" & NAMEARCHIVO$)
   FechaArchivoDestino# = FILEDATE#(RUTADESTINO$) 'ESTA RUTA ESTA CON EL ARCHIVO INCLUIDO
'MsgBox "FectaArchivoDrop#: " & FectaArchivoDrop#  'ESTE VIENE EN 0
'MsgBox "FechaArchivoDestino#: " & FechaArchivoDestino# 'ESTE TRAE INFO
   
   If Abs(FectaArchivoDrop# - FechaArchivoDestino#) <= 0.00002 Then 'si son misma fecha y hora se va
      Call GRACONTROL("BASEWALD", "ESTAGRAB", "NO")
      Exit Sub
   End If
   '
   POSRIGHT% = InStrRev(RUTADESTINO$, "\")
   RUTADESTINO$ = Mid$(RUTADESTINO$, 1, POSRIGHT% - 1) ' le quito de la cadena el archivo por un cambio que hubo y no tener que revisar todo
'MsgBox "RUTADESTINO$: " & RUTADESTINO$
   If EXISTE(RUTADROP$ & "\" & NAMEARCHIVO$) > 0 Then
'MsgBox "EXISTE RUTA" & RUTADROP$ & "\" & NAMEARCHIVO$
        Call GRACONTROL("BASEWALD", "ESTAGRAB", "SI") 'SI ENTRO BLOQUEALA
        If EXISTE(RUTADESTINO$ & "\" & Mid$(NAMEARCHIVO$, 1, Len(NAMEARCHIVO$) - 4)) > 0 Then
           Kill RUTADESTINO$ & "\" & Mid$(NAMEARCHIVO$, 1, Len(NAMEARCHIVO$) - 4) 'ELIMINA EL ARCHIVO NO ZIPEADO
        End If
'MsgBox "VA A COPIAR EN LA CARPETA WALDBOT"
        FileCopy RUTADROP$ & "\" & "SGIW14.mdb", RUTADESTINO$ & "\" & "SGIW14.MDB"
        
        Screen.MousePointer = 11
        Call T_Espera(20)
        Call GRACONTROL("BASEWALD", "ESTAGRAB", "NO")
        Screen.MousePointer = 1
        
   End If
   
ERROR_GUARDAR:
        If Err <> 0 Then
           Call GRACONTROL("BASEWALD", "ESTAGRAB", "NO")
        End If

End Sub
Function BUSCOARCHIVO$(PARTEDELNOMBRE$)
On Error GoTo errsub
'PARTDELNOMBRE NO SE ESTA USANDO ESTA HARCODEADO EL NOMBRE DEL ARCHIVO
RUTADROP$ = TRIM$(Control("COTICONT", "ARCDROPB"))
If RUTADROP$ = "" Then
  Call COMUNI("FALTA CONFIGURAR RUTA DE ARCHIVO DE INTERCAMBIO CON BASE WALDBOTT")
  Exit Function
End If
'Set Fso = New FileSystemObject ' ESTO ESTABA COMENTADO PERO SI NO ESTA ACTIVO NO LEE EL DIRECTORIO
    ' Obtiene el directorio
    Set El_Directorio = fso.GetFolder(RUTADROP$ + "")
    ' Lista los ficheros de esta carpeta
    For Each El_Archivo In El_Directorio.Files
       'Añade la ruta
       MsgBox "El_Archivo.Name " & El_Archivo.Name
       NOMBREARCH$ = El_Archivo.Name
       'If Left$(UCase$(NOMBREARCH$), 14) = "SGIW11.MDB.ZIP" Then ' ANTES PARA BUSCAR ZIPEADO
       If TRIM$(Left$(UCase$(NOMBREARCH$), 10)) = "SGIW14.MDB" Then
          BUSCOARCHIVO$ = NOMBREARCH$
           MsgBox "BUSCOARCHIVO$ " & BUSCOARCHIVO$
          Exit For
       End If

    Next El_Archivo
    Set CTP = Nothing
Exit Function

'Error
errsub:
'Error de permiso denegado
If Err.Number = 70 Then
   Resume Next

ElseIf Err.Number = 91 Then
   Screen.MousePointer = vbDefault
   Exit Function
Else
   'MsgBox Err.Description, vbCritical
   Exit Function
End If
End Function

Function CostoTelasSegunCotizacion1#(NROCOTIZA)
'DEVUELVE EL COSTO DE TELAS QUE ES LA SUMA DE LA SOLAPA DE COTIZACION 1
'ES LA CANTIDAD DE CORTES POR EL USO DE CADA CORTE (QUE ESTA EN CM POR /100) POR EL PRECIO X METRO
CONDI$ = "NROCOTIZACION = " & NROCOTIZA

CostoTelasSegunCotizacion1# = XVALO(VALOBADA("COTI1", "SUM((MODUXBB *(CORTE/100))  * PRECIOXM )", CONDI$, "NOMESS"))
'select SUM((MODUXBB *(CORTE/100))  * PRECIOXM ) AS CORTE from COTI1 where NROCOTIZACION = 1910


End Function
Function CostoLinerSegunCotizacion2#(NROCOTIZA)
CONDI$ = "NROCOTIZACION = " & NROCOTIZA

CostoLinerSegunCotizacion2# = XVALO(VALOBADA("COTI2", "IMPOLINER", CONDI$, "NOMESS"))
End Function

Function SUMADEOTROSSegunCotizacion2#(NROCOTIZA)
    CONDI$ = "NROCOTIZACION = " & NROCOTIZA
    SUMADEOTROSSegunCotizacion2# = XVALO(VALOBADA("COTI2", "SUM(ETIQ_COSTO + SOBRE_COSTO + LINER_COSTO+LIBRE1IMPO+LIBRE2IMPO+LIBRE3IMPO)", CONDI$, "NOMESS"))

End Function
Function PorcentajeOtrosSegunCotizacion2#(NROCOTIZA)
    CONDI$ = "NROCOTIZACION = " & NROCOTIZA
    PorcentajeOtrosSegunCotizacion2# = XVALO(VALOBADA("COTI2", "SUM(ETIQ_COSTO + SOBRE_COSTO + LINER_COSTO+LIBRE1IMPO+LIBRE2IMPO+LIBRE3IMPO)", CONDI$, "NOMESS"))

End Function
Function ScrapSegunCotizacion2#(NROCOTIZA)
    CONDI$ = "NROCOTIZACION = " & NROCOTIZA
    ScrapSegunCotizacion2# = XVALO(VALOBADA("COTI2", "SCRAPP", CONDI$, "NOMESS"))

End Function
Function ComisionSegunCotizacion2#(NROCOTIZA)
    CONDI$ = "NROCOTIZACION = " & NROCOTIZA
    ComisionSegunCotizacion2# = XVALO(VALOBADA("COTI2", "COMISION", CONDI$, "NOMESS"))
End Function
Function ManoObraSegunCotizacion2#(NROCOTIZA)
    CONDI$ = "NROCOTIZACION = " & NROCOTIZA
    ManoObraSegunCotizacion2# = XVALO(VALOBADA("COTI2", "COSTMOBRA", CONDI$, "NOMESS"))
End Function
Function CostosFijosSegunCotizacion2#(NROCOTIZA)
    CONDI$ = "NROCOTIZACION = " & NROCOTIZA
    CostosFijosSegunCotizacion2# = XVALO(VALOBADA("COTI2", "COSTOSFIJOS", CONDI$, "NOMESS"))
End Function
Function CostoTotalSegunCotizacion2#(NROCOTIZA)
    CONDI$ = "NROCOTIZACION = " & NROCOTIZA
    CostoTotalSegunCotizacion2# = XVALO(VALOBADA("COTI2", "COSTTOT", CONDI$, "NOMESS"))
End Function
Function PorcentajeUtilidadSegunCotizacion2#(NROCOTIZA)
    CONDI$ = "NROCOTIZACION = " & NROCOTIZA
    PorcentajeUtilidadSegunCotizacion2# = XVALO(VALOBADA("COTI2", "PJEUTILIDAD", CONDI$, "NOMESS"))
End Function


Sub RecuperarLoteCintilla(Optional NORFA$, Optional DEPO%)


'   Dim CTP As New Conteplast
   CICONJ% = CODIFAB2%(NORFA$)
   '
   ' Esto Aplica a las Cintillas como producto terminado, no a los cierres de bolsones
   EsCintilla% = 0
   If TRIM$(UCase$(CODFAMIL$(CICONJ%))) = "CC" Then EsCintilla% = 1

   'RECORRO LOS CONSUMOS DEL MOVISTO SEGUN LA OF LOS QUE NO POSEEN LOTES
   SQLX$ = "SELECT  COD_INTE,CANTSALID,MOVISTO_ID FROM MOVISTO WITH(NOLOCK)"
   SQLX$ = SQLX$ + " WHERE CODIMOVI = 'CF' "
   SQLX$ = SQLX$ + " AND DOPRICOR ='" & NORFA$ & "'"
   SQLX$ = SQLX$ + " AND CANTSALID > 0 AND (IDENLOTE = '' OR IDENLOTE IS NULL)"
   SQLX$ = SQLX$ + " ORDER BY MOVISTO_ID DESC"
   Set RS = READSET(SQLX$)
'Call CTP.mostrarDatosx(SQLX$)
   With RS
     Do While Not .EOF
       CI1% = XVALO(!cod_inte)
       CANTIRANGO = XVALO(!CANTSALID)
       ID& = XVALO(!MOVISTO_ID)
       'SI ES CODIGO RANGO CONSULTO EL MATERIAL TRANSFERIDO PARA LA ORDEN AL DEPOSITO DE CORTE
       'Y REEMPLAZO LOS CODIGOS RANGO POR LOS TRANSFERIDOS QUE CORRESPONDAN AL RANGO ACTIVO
       If ESCODIGORANGO&(CI1%) > 0 Then
'MsgBox "ES CODIGO RANGO"

         CODIRANGO$ = CODEXT$(CI1%)
         CONDI$ = "CODRANGO = '" & CODIRANGO$ & "'"
         CONDIOR$ = CTP.CONSULTA_OR("CODIGORANGO", "CODMATPRIMA", CONDI$, "COD_EXTE")
         SQLX$ = "SELECT SUM(CANTINGRE-CANTSALID) AS RESULTADO ,COD_INTE,COD_EXTE,DESCRIP,IDENLOTE"
         SQLX$ = SQLX$ + " FROM MOVISTO WITH(NOLOCK)"
         SQLX$ = SQLX$ + " where DEPOMOVI = " & DEPO%
         'SI ES CINTILLA NO FILTRA POR ORDEN DE FABRICACION
         If EsCintilla% < 1 Then SQLX$ = SQLX$ + "  AND (DOSECLAR =   '" & TRIM(NORFA$) & "'  OR DOPRICOR = '" & TRIM(NORFA$) & "')"

         If CONDIOR$ <> "" Then SQLX$ = SQLX$ + " AND " & CONDIOR$
         SQLX$ = SQLX$ + " GROUP BY COD_INTE,COD_EXTE,DESCRIP,IDENLOTE"
         SQLX$ = SQLX$ + " Having Sum(CANTINGRE - CANTSALID) > 0"
'Call CTP.mostrarDatosx(SQLX$)
         Set rs1 = READSET(SQLX$)
         With rs1
           Do While Not .EOF
              CIMTP = XVALO(!cod_inte)
              CANTITRANS = XVALO(!Resultado)
'MsgBox "CANTIDAD TRANSFERIDA " & CANTITRANS & " LOTE: " & SAFETEXT$(!IDENLOTE)
              SQLX$ = "UPDATE MOVISTO SET COD_INTE = " & XVALO(!cod_inte)
              SQLX$ = SQLX$ + "  ,COD_EXTE = '" & SAFETEXT$(!Cod_Exte) & "'"
              SQLX$ = SQLX$ + "  ,DESCRIP = '" & SAFETEXT$(!Descrip) & "'"
              SQLX$ = SQLX$ + "  ,DESCRI_LAR = '" & SAFETEXT$(!Descrip) & "'"
              SQLX$ = SQLX$ + "  ,IDENLOTE = '" & SAFETEXT$(!idenlote) & "'"
              SQLX$ = SQLX$ + "  ,DEPOMOVI = " & DEPO%
              If CANTITRANS < CANTIRANGO Then
                'si la cantidad transferida es menor a la necesidad actualiza cantsalid
                SQLX$ = SQLX$ + "  ,CANTSALID = '" & FORMATNUM$(CANTITRANS, "F16.5") & "'"
                '
                'Antes de grabar genera duplica el registro por la diferencia(manteniendo el codigo rango)
                NEWID& = 1 'PARA QUE DEVUELVA ID
                '
                'SQLX1$
                Call DuplicarRegistro("MOVISTO", "MOVISTO_ID = " & ID&, NEWID&)
                Diferencia = CANTIRANGO - CANTITRANS: If Diferencia < 1 Then Diferencia = CANTIRANGO
'MsgBox "Diferencia " & Diferencia
                SQLX1$ = "UPDATE MOVISTO SET CANTSALID = '" & FORMATNUM$(Diferencia, "F16.5") & "'"
                SQLX1$ = SQLX1$ + "  WHERE  MOVISTO_ID = " & NEWID&
'Call CTP.mostrarDatosx(SQLX1$)
                FLEXCONN.Execute SQLX1$
'MsgBox "Diferencia Aplicado a id " & NEWID&

                'ahora actualiza el registro con el codigo transfereido real incluyendo la cantidad   SQLX$
                SQLX$ = SQLX$ + "  WHERE  MOVISTO_ID = " & ID&
'Call CTP.mostrarDatosx(SQLX$)
                FLEXCONN.Execute SQLX$
                ID& = NEWID& 'SI ENTRA POR ESTE IF ACTUALIZA EL ID P Q EL NUEVO REGISTRO ES DUPLICADO DEL ORIGINAL POR LA DIFERENCIA
                             'DADO QUE SI EN LA SIGUIENTE VUELTA NO ENTRA POR EL ELSE  APLICA DIRECTO POR EL ID
              Else
                'SI VIENE ACA ES QUE ALCANZA CON LA CANTIAD LEIDA DEL CODIGO TRANSFERIDO
                SQLX$ = SQLX$ + "  WHERE  MOVISTO_ID = " & ID&
'MsgBox "Aplicacion directa a id " & ID&
'Call CTP.mostrarDatosx(SQLX$)
                FLEXCONN.Execute SQLX$
                Exit Do
              End If
              CANTIRANGO = CANTIRANGO - CANTITRANS
              If CANTIRANGO <= 0 Then Exit Do
             .MoveNext
           Loop
         End With
       End If

       .MoveNext
     Loop
   End With
   Set CTP = Nothing

End Sub


Function CODIFAB2%(NORFX$, Optional TABLA$)
    CICON% = 0
    '
    If TRIM$(TABLA$) = "" Then TABLA$ = "OrdeFabr"
    If TRIM$(NORFX$) <> "" Then
      Call ABREORFAB
      Dim ORDEFA As ADODB.Recordset
      SQLX$ = "SELECT Cod_Inte FROM " & TABLA$ & " WHERE IdSubOr = '" & NORFX$ & "'"
      Set ORDEFA = FLEXCONN.Execute(FILTSQL$(SQLX$))
      '
      If Not (ORDEFA.BOF And ORDEFA.EOF) Then
         CICON% = ORDEFA!cod_inte
      End If
      ORDEFA.Close
      Set ORDEFA = Nothing
    End If
    '
    CODIFAB2% = CICON%
End Function
Function CantiProyectadaOF2(NUMORFA$)
   '
   CONDI$ = "IdSubOr = '" & NUMORFA$ & "'"
   
   CantiProyectadaOF2 = XVALO(VALOBADA("ordefabr", "CANPROY", CONDI$, "NOMESS"))
'   Call ABREORFAB
'   '
'   SQLX$ = " SELECT * FROM ordefabr with(nolock)"
'   SQLX$ = SQLX$ + " WHERE IdSubOr = '" & NUMORFA$ & "' "
'   '
'   Set INFOFABTMP = READSET(SQLX$)
'
'   CantiProyectadaOF = 0
'   With INFOFABTMP
'    CantiProyectadaOF = XVALO(!CANPROY)
'   End With
   '
End Function
Function ReferenSegunOf2$(NORFA$)
    ReferenSegunOf2$ = VALOBADA("ORDEFABR", "Referen", "IdSubOr ='" & NORFA$ & "'")
End Function
Function UltimoNumOrdenSegunOFyOperacion2&(NORFA$, NOPE%)
  CONDI$ = "IdSubOr = '" & NORFA$ & "' AND NumeOper = " & NOPE%
  UltimoNumOrdenSegunOFyOperacion2& = CantRegistros&("INFOFAB", CONDI$, "NOMESS")
End Function

Function NroPedidoSegunOf2&(NORFA$)
    NroPedidoSegunOf2& = XVALO(VALOBADA("ORDEFABR", "NroPed", "IdSubOr ='" & NORFA$ & "'"))
End Function
Function NroOrdenFabricacionSegunPedido$(NROPED&)
    NroOrdenFabricacionSegunPedido$ = VALOBADA("ORDEFABR", "IDSUBOR", "NROPED =" & NROPED&, "NOMESS")
End Function

Function PesoPromedioDeCierresFabricSegunOF(NORFA$)
    CONDI$ = "IdenLote =  '" & NORFA$ & "' AND CANTINGRE > 0 AND CODIMOVI = 'OF'"
    PesoPromedioDeCierresFabricSegunOF = XVALO(VALOBADA("MOVISTO", "SUM(PESOKG) / SUM(CantIngre)", CONDI$, "NOMESS"))
End Function
Function SaldoTallerSegunOF&(NORFA$)
   SQLX$ = "select sum(cantsalid - cantingre)FROM MOVTALLER WITH(NOLOCK) "
   SQLX$ = SQLX$ + " where doprilar = '" & NORFA$ & "'"
   SQLX$ = SQLX$ + " and (Codimovi = 'DT' OR CODIMOVI = 'ST')" ' DEVOLUCION TALLER Y SALIDA TALLER
   Set RS = READSET(SQLX$)
   SaldoTallerSegunOF& = XVALO(RS.Fields(0).Value)
   
End Function
Function UltimoMovistoIdGeneradoSegunOf(NORFA$)
    'ESTO SE UTILIZA PARA QEU LUEGO DE GENERAR UNA ORDEN DE FABRICACION
   'SIRVE PARA REALIZARLO INMEDIATAMENTE LUEGO DE LA GENERACION DE LA OF
   CONDI$ = "CODIMOVI = 'OF' AND CANTINGRE > 0"
   CONDI$ = CONDI$ + " AND DOPRILAR = '" & NORFA$ & "'"
   CONDI$ = CONDI$ + " ORDER BY MOVISTO_ID  DESC"
   ID& = XVALO(VALOBADA("MOVISTO", "TOP 1 MOVISTO_ID", CONDI$, "NOMESS"))
   UltimoMovistoIdGeneradoSegunOf = ID& 'retorono de id
End Function
