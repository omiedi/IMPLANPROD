Attribute VB_Name = "PedVsFabric"
Sub APLICAR_DIFER_A_FABRICAR(MSF As MSFlexGrid)
   For i& = 2 To MSF.Rows - 1 Step 4 ' RECORRO LAS FILAS STEP 4 POR QUE CADA CUATRO ENCUENTRA EL CODIGO
      CI1% = CODINT%(MSF.TextMatrix(i&, 0))
      If CI1% > 0 Then 'LEE EL CODIGO
            For J% = 1 To MSF.Cols - 1 'RECORRE LAS COLUMNAS
               NECESIDAD = XVALO(MSF.TextMatrix(i&, J%))
               If NECESIDAD > 0 Then ' SI TIENE VALOR (PEDIDO)
                  STOCKS1$ = TRIM$(MSF.TextMatrix(i& + 1, J%))
                  If STOCKS1$ <> "" Then
                        Erase CADENATEX$
                        Call TEXTOLOTES$(STOCKS1$, "/")
                        STO = XVALO(CADENATEX$(1))
                        AFABRICAR = NECESIDAD - STO
                        
                        If AFABRICAR > 0 Then
                           MSF.TextMatrix(i& + 2, J%) = FORMATNUM$(AFABRICAR, "F9")
                           Call COLOREAR_GRILLA_SOLACELDA(MSF, vbRed, i& + 2, J%)
                        End If
                   Else
                        'SI NO ENTRO EN EL IF A FABRICAR TODA LA NECESIDAD.
                        MSF.TextMatrix(i& + 2, J%) = FORMATNUM$(NECESIDAD, "F9")
                        Call COLOREAR_GRILLA_SOLACELDA(MSF, vbRed, i& + 2, J%)

                   End If
               End If
            Next J%
      End If
   Next i&

End Sub

Sub APLICAR_STOCKS(MSF As MSFlexGrid)
   'APLICA EL STOCK PARA CADA CODIGO
   'EJEMPLO TIENE PARA UNA FECHA PEDIDOS POR 4 Y PARA OTRA PEDIDOS POR 12, SIENDO EL STOCK DE 10
   'EN LA PRIMERA LE APLICA 4/10 (COMPLETA LOS 4 DE UN STOCK DE 10)
   'EN LA SEGUNDA LE APLICA 6/6  (COMPLETA SOLO 6 RESTANTES POR QUE ES EL STOCK QUE LE QUEDA COMO SALDO
    
   For i& = 2 To MSF.Rows - 1 Step 4 ' RECORRO LAS FILAS STEP 4 POR QUE CADA CUATRO ENCUENTRA EL CODIGO
      CI1% = CODINT%(MSF.TextMatrix(i&, 0))
      If CI1% > 0 Then 'LEE EL CODIGO
         ST1 = STOCKACT(CI1%) 'LEE EL STOCK
         If ST1 > 0 Then
            For J% = 1 To MSF.Cols - 1 'RECORRE LAS COLUMNAS
               NECESIDAD = XVALO(MSF.TextMatrix(i&, J%))
               If NECESIDAD > 0 Then ' SI TIENE VALOR (PEDIDO)
                  Call COLOREAR_GRILLA_SOLACELDA(MSF, &HFFFFC0, i& + 1, J%)
                  If ST1 < NECESIDAD Then
                     NECESIDAD = ST1
                     Call COLOREAR_GRILLA_SOLACELDA(MSF, vbYellow, i& + 1, J%)
                  End If
                  tEXsTOCK = NECESIDAD & "/" & ST1
                  MSF.TextMatrix(i& + 1, J%) = FORMATNUM$(tEXsTOCK, "F9")
                  ST1 = ST1 - NECESIDAD
                  If ST1 <= 0 Then Exit For 'SI SE COMIO EL STOCK PASA AL SIGUIENTE
               End If
            Next J%
          End If
         End If
   Next i&
End Sub
Function ExcedentedeFabrica(CI1%)
  'DEVUELVE QUE HAY EN O.FABRICACION GENERADAS RESTANDOLE LA CANTIDAD PEDIDA DEL MISMO CODIGO
  '
  FabrPendiente = FabricacionPendiente(CI1%)
  CantPedPend = CANTIDAD_PEDIDA_PENDIENTE(CI1%)
  '
  ExcedentedeFabrica = FabrPendiente - CantPedPend
  If ExcedentedeFabrica < 0 Then ExcedentedeFabrica = 0
End Function
Function DiaNoLaborable%(DIA%)
     DIAFERIADO% = 0
     '
     'VALIDA LOS DIAS NO LABORABLES FIJOS (DOMINGO,LUNES,MARTES,....)
     XX$ = CONTROL("CONFABPE", "NOLABORA") 'ESTE CONTROL GUARDA SEGUN LAS OPCIONES ELEGIDAS EJ. DIAS NO LABORABLES SAB Y DOM "0;6"
     If InStr(XX$, SAFETEXT$(DIASEM%(DIA%))) > 0 Then
       DIAFERIADO% = 1
       GoTo 99
     End If
     '
     For U& = 1 To ULTREG&("CALFERI")
         If CVI(Left$(REGLEIDO$("CALFERI", U&), 2)) = DIA% Then
             DIAFERIADO% = 1
             GoTo 99
         End If
     Next U&
     '
     AAA = DIASEM%(DIA%)
99   DiaNoLaborable% = DIAFERIADO%
End Function

Function FechaFinalTotalPendFabricacion%(MOLDE%)
   'Devuelve la fecha final de las ordenes de fabricacion de todos los articulos que compartan el mismo molde
    CANTIUNIDADES = CantiUnidPendienteCierreXMolde(MOLDE%)
    If CantiConfigMoldes(MOLDE%) > 0 Then
      DIASNETOS = 1 'cantidad de dias neto por defecto
      CANMOLDES% = CantiConfigMoldes(MOLDE%)
      If CANTIUNIDADES > CANMOLDES% Then
         DIASNETOSCONDECIMAL = CANTIUNIDADES / CANMOLDES%
         DIASNETOS = Int(DIASNETOSCONDECIMAL)
         If DIASNETOSCONDECIMAL > DIASNETOS Then
           DIASNETOS = DIASNETOS + 1
         End If
       End If
    Else
       Call COMUNI("Falta Configurar Cantidad de Moldes (Modelo: " & ECOARCH$("MOLDES", MKI$(MOLDE%)) & ")")
    End If
    '
    DiasReales% = CantDiasDemoraReal%(FechaInicio%, CantiDias%)
    
End Function
Function CantDiasDemoraReal%(FechaInicio%, CantiDias%)
   FE1% = DIASPOST(FechaInicio%, CantiDias%)
   AA = FECHATEX$(FE1% + 2)
   DIAX% = FechaInicio%
   For DIAX% = DIAX% To FE1% '
     CantDemoraReal% = CantDemoraReal% + 1
     If DiaNoLaborable(DIAX%) > 0 Then
       CantDemoraReal% = CantDemoraReal% + 1
     End If
   Next DIAX%
   CantDiasDemoraReal% = DIASPOST(FechaInicio%, CantDemoraReal%)
End Function
Function CantiUnidPendienteCierreXMolde(MOLDE%)
   'Recorrer las ordenes de fabricacion con status < 3 corresponde al mismo molde
   'tomar la diferencia de canproy - canapro y acumularla
   ACUMU_SALDOPEN = 0
   SQLX$ = "SELECT M.MOLDE, M.CODINT, O.IDSUBOR,O.Cod_Inte,O.STATUS,O.CANPROY , O.CANAPRO,"
   SQLX$ = SQLX$ + " FROM MASTER M INNER JOIN ORDEFABR O"
   SQLX$ = SQLX$ + " ON M.CODINT = O.Cod_Inte"
   SQLX$ = SQLX$ + " WHERE O.STATUS < 3 "
   SQLX$ = SQLX$ + " AND O.CANPROY > O.CANAPRO"
   SQLX$ = SQLX$ + " AND M.MOLDE = & Molde% "
   SQLX$ = SQLX$ + " ORDER BY O.IDSUBOR"
   Set RS = READSET(SQLX$)
   With RS
     Do While Not .EOF
       CAPROY = XVALO(!CANPROY)
       CANPRO = XVALO(!CANAPRO)
       SALDOPEN = CAPROY - CANPRO
       ACUMU_SALDOPEN = ACUMU_SALDOPEN + SALDOPEN
       .MoveNext
     Loop
   End With
   RS.Close
   Set RS = Nothing
   CantiUnidPendienteCierreXMolde = ACUMU_SALDOPEN
   '
End Function
Function DemoraenDiasFabricaciondeItem%(CI1%, CANTIUNIDADES)
   'Devuelve la Cantidad de dias que demora un item en fabricarse segun la cantidad
      DIASNETOS% = 1 'cantidad de dias neto por defecto
      Moldex% = MOLDE%(CI1%)
      CANMOLDES% = CantiConfigMoldes(Moldex%)
      If CANMOLDES% < 1 Then
        Call COMUNI("Falta Configurar Cantidad de Moldes (Modelo: " & ECOARCH$("MOLDES", MKI$(Moldex%)))
        Exit Function
      End If

      UNME$ = Unimed$(CI1%)
      '
      CAN = CANTIUNIDADES 'guardo el valor para que no vuelva aguas arriba camibado
      If UNME$ = "M2" Then
         CANTIUNIDADES = CANTIUNIDADES * 6
      End If
      '
      If CANTIUNIDADES > CANMOLDES% Then
         DIASNETOSCONDECIMAL = CANTIUNIDADES / CANMOLDES%
         DIASNETOS% = Int(DIASNETOSCONDECIMAL)
         If DIASNETOSCONDECIMAL > DIASNETOS% Then
           DIASNETOS% = DIASNETOS% + 1
           
         End If
       End If
       '
       CANTIUNIDADES = CAN ' ACTUALIZO AL VALOR ORIGINAL, PARA QUE NO VUELVA AGUAS ARRIBA CAMBIADO
    '

    DiasReales1% = CantDiasDemoraReal%(HOY(HO$), DIASNETOS%)
    DIFERENCIAREAL = DiasReales1% - HOY(HO$)
    DemoraenDiasFabricaciondeItem% = DIFERENCIAREAL
    
End Function

Function FechaExcedenteFabricacion(CI1%, CANTI)
   'CON ExcedenteFabricacion SE LA CANTIDAD QUE TENGO COMO EXCEDENTE
   'FechaExcedenteFabricacion =PARA CUANDO TENGO ESA CANTIDAD.
   'FECHA DE DISPONIBILIDAD = LA FECHA CUANDO SE TERMINARIA DE FABRICAR Y SE PRODUCIRIAN EL EXCEDENTE TOTAL DE DICHA CANTIDAD
   'DEBERIAMOS SABER LA FECHA, SEGUN LA CANTIDAD.
End Function
Sub MOSTRAR_KARDEX(CI1%)
   If CI1% > 0 Then
    CODI$ = CODEXT$(CI1%)
    If CODI$ <> "" Then
      FIKARDEX09.Text1 = CODI$
      FIKARDEX09.Show 1
    Else
      Call COMUNI("Código Inexistente o No Válido")
    End If
   End If
End Sub

Function UltimoN°Pedido(CI1%)
     'DEVUELVE EL ULTIMO NUMERO DE PEDIDO DE UN ITEM (QUE CONTENGA CANTIDAD PENDIENTE)
     UltimoN°Pedido = 0
     SQLX$ = " SELECT NROPED,FechEmis FROM PEDDETA WITH(NOLOCK) "
     SQLX$ = SQLX$ + "WHERE CodiInt = " & CI1% & " "
     SQLX$ = SQLX$ + "AND Status < 8 "   ' para que no este anulada ni cumplida
     SQLX$ = SQLX$ + "AND CantEnt < (CanPed - 0.005) "
     SQLX$ = SQLX$ + "ORDER BY NroPed DESC "
     '
     'ATENCION NO REALIZO UN SUM DE LAS CANTIDADES POR SI EN PEDIDOS HAY ENTREGAS MAYORES A LA CANTIDAD SOLICITADA
     Set PEDTEMP = READSET(SQLX$)
     With PEDTEMP
        If Not .EOF Then
          UltimoN°Pedido = XVALO(!NROPED)
        End If
     End With
End Function
Function CantPendEntregaConsideraMolde(Moldex%, NroPedidoLimite&)
     'DEVUELVE LA CANTIDAD PENDIENTE DE ENTREGA DE TODOS LOS PEDIDOS DE TODOS LOS ITEMS QUE COMPARTAN MOLDE HASTA EL NROPEDIDO_
     'PASADO COMO PARAMETRO.
     CantPendEntregaConsideraMolde = 0
     
     SQLX$ = " SELECT  P.CantEnt AS CANTEN, P.CanPed AS CANPED FROM PEDDETA P WITH(NOLOCK) INNER JOIN MASTER M ON P.CODIINT = M.CODINT"
     SQLX$ = SQLX$ + " WHERE M.MOLDE = " & Moldex%
     SQLX$ = SQLX$ + " AND P.Status < 8 "   ' para que no este anulada ni cumplida
     SQLX$ = SQLX$ + " AND P.CantEnt < (P.CanPed - 0.005) "
     SQLX$ = SQLX$ + " AND P.NROPED <= " & NroPedidoLimite&
     SQLX$ = SQLX$ + " ORDER BY P.FeSolEnt,P.NROPED DESC "
     '
     'ATENCION NO REALIZO UN SUM DE LAS CANTIDADES POR SI EN PEDIDOS HAY ENTREGAS MAYORES A LA CANTIDAD SOLICITADA
     Set RS = READSET(SQLX$)
     With RS
       Do While Not .EOF
        If Not .EOF Then
          CantPendEntregaConsideraMolde = CantPendEntregaConsideraMolde + (XVALO(!CANPED) - XVALO(!CANTEN))
        End If
        .MoveNext
       Loop
     End With
     RS.Close
     Set RS = Nothing
     
End Function

Function DemoraenDiasExcedenteFabricaciondeItem%(CI1%)
  'devuelve los dias que tarda en fabricar para completar todos los pedidos que compartan el molde
  UltNroped& = UltimoN°Pedido(CI1%)
  Moldex% = MOLDE%(CI1%)
  CantPendEntr = CantPendEntregaConsideraMolde(Moldex%, UltNroped&)
  DemoraenDiasExcedenteFabricaciondeItem% = DemoraenDiasFabricaciondeItem%(CI1%, CantPendEntr)
End Function
Function CANTIXM2(CI1%, NLIS%)
       'si no lo encuentra devuelve -1
       RFF$ = RECFILT$("LISTASM2", 1, MKI$(CI1%))
       FIN& = Len(RFF$)
       For UK& = 1 To FIN& Step 4
          U& = CVS(Mid$(RFF$, UK&, 4))
          X$ = REGLEIDO$("LISTASM2", U&)
          NL% = Asc(Mid$(X$, 3, 1))
          If NL% = NLIS% Then
            CANTIXM2 = CVD(Mid$(X$, 4, 8))
            Exit For
          End If
       Next UK&
       If CANTIXM2 <= 0 Then CANTIXM2 = -1
End Function


Sub CARGA_ENCA_DIAG(MSF As MSFlexGrid)
   MSF.Rows = 3
   MSF.Clear
   HOYX% = HOY(HO$)
   MSF.Cols = 120
   MSF.ColWidth(0) = 1750
   For i& = 1 To MSF.Cols - 1
        J% = J% + 1
        MSF.Cols = J% + 1
        MMXX1$ = METEXTO$(HOYX%)
        MSF.TextMatrix(0, J%) = MMXX1$
        MSF.ColWidth(J%) = 750:   MSF.ColAlignment(J%) = flexAlignCenterTop
        MSF.TextMatrix(1, J%) = Left$(DSEM$(HOYX%), 3) & "-" & Left$(FECHATEX$(HOYX%), 2) & Space(10) & SAFETEXT$(HOYX%) 'GUARDO DE MANERA OCULTA LA FECHANUM
        If DiaNoLaborable%(HOYX%) > 0 Then
          Call COLOREAR_TEXTO_SOLO_UNA_AFILA(MSF, vbRed, 1, J%)
        End If
        HOYX% = DIASPOST(HOYX%, 1)
   Next i&
   '
   MSF.MergeRow(0) = True
   MSF.MergeCells = flexMergeFree
   MSF.FixedRows = 2
   Call MostrarDiagrama(DIAGPREM.MSF1, DIAGPREM.msf2)
End Sub
Sub MostrarDiagrama(MSF As MSFlexGrid, GRILLAPEDIDOS As MSFlexGrid)
   '1.cargar encabezado
   Screen.MousePointer = 11
   HOYX% = HOY(HO$)
   REG& = 1
   VUELTA% = 0
   MSF.Rows = 2
   MSF.Redraw = False
   CONDI$ = "CANPED > CANTENT AND P.STATUS < 8"
   SQLX$ = "SELECT P.CODIINT, P.CANPED AS PEDIDO ,P.CANTENT AS ENTREGADO, M.Umedida ,P.FESOLENT,P.NROPED AS NP,P.NROCLIE AS NCL"
   SQLX$ = SQLX$ + " FROM PEDDETA P WITH(NOLOCK) INNER JOIN MASTER M WITH(NOLOCK) ON P.CODIINT = M.CODINT"
   SQLX$ = SQLX$ + " Where " & CONDI$
   'SQLX$ = SQLX$ + " GROUP BY P.CODIINT,M.Umedida, FESOLENT,P.NROPED,P.NROCLIE"
   SQLX$ = SQLX$ + " ORDER BY CODIINT,FESOLENT, NROPED,  NROCLIE"
   Set RS = READSET(SQLX$)
   With RS
       Do While Not .EOF
         CI1% = XVALO(!CODIINT)
         If (CI1% <> CI_ANT% And CI_ANT% > 0) Or VUELTA% = 0 Then

           'MSF.TextMatrix(REG&, DiferenciaDias%) = MSF.TextMatrix(REG&, DiferenciaDias%) & "    " & PEDIDOS(REG&, DiferenciaDias%)
           'IMPRIMO EL CODIGO EN LA COLUMNA 0 EN DOS FILAS
           REG& = REG& + 1
           MSF.Rows = REG& + 1
           MSF.TextMatrix(REG&, 0) = CODEXT$(CI1%)
           MSF.RowHeight(REG&) = 300
           
           REG& = REG& + 1
           MSF.Rows = REG& + 1
           MSF.TextMatrix(REG&, 0) = "STOCK"
           MSF.RowHeight(REG&) = 300
           
           REG& = REG& + 1
           MSF.Rows = REG& + 1
           MSF.TextMatrix(REG&, 0) = "A FABRICAR"
           MSF.RowHeight(REG&) = 300
           '
           REG& = REG& + 1
           MSF.Rows = REG& + 1
           MSF.RowHeight(REG&) = 200

           VUELTA% = 1
'           If REG& > 2 Then
'             ReDim Preserve PEDIDOS(REG&, MSF.Cols)
'           End If
         End If
         '
         REG& = REG& - 3
         'CALCULA LA DIFERENCIA EN DIAS PARA VER DONDE UBICO EL PEDIDO EN LA GRILLA.
         FESOLICITUDENTREGA% = CVINT(!FeSolEnt)
         AAA = FECHATEX$(FESOLICITUDENTREGA%)
         If FESOLICITUDENTREGA% < HOYX% Then
            FESOLICITUDENTREGA% = HOYX%
         End If
         PendEntrega = XVALO(!PEDIDO) - XVALO(!ENTREGADO): If PendEntrega < 0 Then PendEntrega = 0
         FECHAaUBICAR% = FESOLICITUDENTREGA%
20       If DiaNoLaborable%(FESOLICITUDENTREGA%) > 0 Then
               FESOLICITUDENTREGA% = DIASPOST(FESOLICITUDENTREGA%, 1): GoTo 20
         End If
         AAA = FECHATEX$(FECHAaUBICAR%)
         DiferenciaDias% = 1 + DiferenciaDiasEntreFechas(HOYX%, FESOLICITUDENTREGA%)
         If DiferenciaDias% < 1 Then DiferenciaDias% = 1
         MSF.TextMatrix(REG&, DiferenciaDias%) = FORMATNUM$(XVALO(MSF.TextMatrix(REG&, DiferenciaDias%)) + PendEntrega, "F9.0")
         Call COLOREAR_GRILLA_SOLACELDA(MSF, &HC0E0C0, REG&, DiferenciaDias%)
         Call CARGARGRILLA_PEDIDOS(GRILLAPEDIDOS, REG&, DiferenciaDias%, XVALO(!NP))
         'PEDIDOS(REG&, DiferenciaDias%) = PEDIDOS(REG&, DiferenciaDias%) & XVALO(!NP) & "|" & XVALO(!NCL) & "|" & PendEntrega & ";"
         CI_ANT% = CI1%
         REG& = REG& + 3
         .MoveNext
       Loop
   End With
   RS.Close
   Set RS = Nothing
   Call APLICAR_STOCKS(MSF)
   Call APLICAR_DIFER_A_FABRICAR(MSF)
   MSF.Redraw = True
   Screen.MousePointer = 1
End Sub
Sub PresentarPedidos(MSF As MSFlexGrid, MSF1 As MSFlexGrid)
   Call MostrarDiagrama(MSF, MSF1)
End Sub

Sub CARGARGRILLA_PEDIDOS(MSF As MSFlexGrid, REG&, COL%, NROPED&)
    TEXTOPOSICION$ = TRIM$(SAFETEXT$(REG&) & ";" & SAFETEXT$(COL%))
    'LO BUSCA SI LO ENCUENTRA LE CONCATENA EL NRO DE PEDIDO Y TERMINA LA RUTINA
    For i& = 1 To MSF.Rows - 1
       If TRIM$(MSF.TextMatrix(i&, 1)) = TEXTOPOSICION$ Then
          MSF.TextMatrix(i&, 2) = MSF.TextMatrix(i&, 2) & NROPED& & ";"
          Exit Sub
       End If
    Next i&
    '
    'SI NO LO ENCUENTRA
    MSF.Rows = MSF.Rows + 1
    MSF.TextMatrix(MSF.Rows - 1, 1) = TEXTOPOSICION$
    MSF.TextMatrix(MSF.Rows - 1, 2) = NROPED& & ";"
End Sub

Function CantiConfigMoldes(Moldex%)
        X$ = FILTERGETRECORD$("MOLCANTI", 1, MKI$(Moldex%))
        CantiConfigMoldes = CVI(Mid(X$, 3, 2))
        
End Function
Function ClienteSegunNroPedido(NUMPE&)
   CONDI$ = "NROPED = " & NUMPE&
   ClienteSegunNroPedido = XVALO(VALOBADA("PEDENCA", "NroClie", CONDI$, "NOMESS"))
End Function

Function DESCARTEXOF(NORFA$, INSOLAPASUP%, INDSOLAPAINF%, Optional FE%, Optional MAQ$)
    'SI NO SE QUIERE FILTRAR POR INDSOLAPAINF% PASAR COMO PARAMETRO -1
    SQLX$ = " select SUM(CANT_DESC) as descarte from DESCARCONS WITH(NOLOCK)"
    SQLX$ = SQLX$ + " where IDSUBOR = '" & NORFA$ & "'"
    If INSOLAPASUP% >= 0 Then SQLX$ = SQLX$ + " AND INDISOLSUP = " & INSOLAPASUP%
    If INDSOLAPAINF% >= 0 Then SQLX$ = SQLX$ + " AND  INDISOLINF = " & INDSOLAPAINF%
    If FE% > 0 Then SQLX$ = SQLX$ + " AND  FECHA = '" & FETEXCOR1$(FE%) & "'"
    If TRIM$(MAQ$) <> "" Then SQLX$ = SQLX$ + " AND  NUMMAQUI = '" & MAQ$ & "'"
    Set RS = READSET(SQLX$)
    With RS
      DESCARTEXOF = XVALO(!descarte)
    End With
    
    

End Function
