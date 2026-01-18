Attribute VB_Name = "MAST_FUN_SQL"
Public PUNCODIG As String
Public BUFERCOD As String * 2048
Public NUSU%, COSU%(64), SUCU$(64), ADEPO$(64), DISTO%(64), CATEO#(64), CAFIN#(64)
Dim CIANT%, XXX$
Public PHABIT%()
Public UPDATMASTER%, UPDAPLANOS%
'
Dim MasterTemp As ADODB.Recordset
Sub ACTUALZACION_MULTIPLE_LISTAS_SELECCIOON(archivo$)
  'SOLO DEBERIA ACTUALIZAR CUANDO ESTA GRABANDO EN EL PRIMERO EN LA LISTA.
  'Y SOLO EN EL CASO QUE LOS DOS TENGAN EL MISMO CONNSQL
  
End Sub

Sub AltaM2AMaster(CODIGOX$)
    'aca da de alta de la tabla MASTER2  A LA TABLA MASTER
    If EmpresaNoRealizaCambios% > 0 Then
       Exit Sub
    End If
    
    CODIGOX$ = TRIM$(CODIGOX$)
    CI1X = CODINT_NEW
    SQLX$ = "INSERT INTO MASTER(Ferevis,Observa,Especificacion,DesCompAut,DescriComp,DepoEntra,DepoCon,Codigo,Descripcion,Descridos,"
    SQLX$ = SQLX$ + " Status , CODINT , Alias, Umedida, TIMATE, CRITREP, Monecos, FECOSTO, PorDescCos, COSUNI, PORNACIO, PorComVen, PROVHABI, STOMIN,"
    SQLX$ = SQLX$ + " Stomax , Lorepos, LINDIVIS, PESUNI, PESMETRO,Color,Talle,Descompo,Grupiva,Grucoven,CLinea,Cfamilia,NPlano,Revision,Trazable,"
    SQLX$ = SQLX$ + " Revi_Formu , FechApFor, UsuApFor, Cod_Interna, Bien_Cap, Codigo_Ncm, Codigo_Sec, Pro_Umed, Iva_Id, ArchiPdf,"
    SQLX$ = SQLX$ + " Familia,Lindiven,CanDeci,ConsuProm,CosCal,FeCosCal,Stat_Movi,Clas_ATC,Vers_BC,FeVig_BC,"
    SQLX$ = SQLX$ + " FeRev_BC,CantApro,LeTiCom,PiezasXCaja,M3Caja,KiloNetoCaja,KiloBrutoCaja,Autono_Proy,UsuPrePes,"
    SQLX$ = SQLX$ + " UsuPreDol,MANO_OBRA_REAL,Tamano_Etiqueta,CueContaI,Molde)"
    '
    'VER QUE LA INFO SE MIGRA CON CODINT= 0
    SQLX$ = SQLX$ + " select Ferevis,Observa,Especificacion,DesCompAut,DescriComp,DepoEntra,DepoCon,Codigo,Descripcion,Descridos,"
    SQLX$ = SQLX$ + " Status , " & CI1X & " , Alias, Umedida, TIMATE, CRITREP, Monecos, FECOSTO, PorDescCos, COSUNI, PORNACIO, PorComVen, PROVHABI, STOMIN,"
    SQLX$ = SQLX$ + " Stomax , Lorepos, LINDIVIS, PESUNI, PESMETRO,Color,Talle,Descompo,Grupiva,Grucoven,CLinea,Cfamilia,NPlano,Revision,Trazable,"
    SQLX$ = SQLX$ + " Revi_Formu , FechApFor, UsuApFor, Cod_Interna, Bien_Cap, Codigo_Ncm, Codigo_Sec, Pro_Umed, Iva_Id, ArchiPdf,"
    SQLX$ = SQLX$ + " Familia,Lindiven,CanDeci,ConsuProm,CosCal,FeCosCal,Stat_Movi,Clas_ATC,Vers_BC,FeVig_BC,"
    SQLX$ = SQLX$ + " FeRev_BC , CantApro, LeTiCom, PiezasXCaja, M3Caja, KiloNetoCaja, KiloBrutoCaja, Autono_Proy, USUPREPES, "
    SQLX$ = SQLX$ + " UsuPreDol,MANO_OBRA_REAL,Tamano_Etiqueta,CueContaI,Molde FROM MASTER2 WHERE CODIGO = '" & CODIGOX$ & "'"
    On Error Resume Next
    FLEXCONN.Execute (SQLX$)
     If Err Then
       On Error GoTo 0
       Call COMUNI("No Se Realizó La Transacción para El Código : " & CODIGOX$)
       Exit Sub
    End If
   
    SQLX$ = "DELETE FROM MASTER2 WHERE CODIGO = '" & CODIGOX$ & "'"
    FLEXCONN.Execute (SQLX$)
    If Err Then
       On Error GoTo 0
       Call COMUNI("No Se Realizó El Borrado para El Código : " & CODIGOX$)
       Exit Sub
    End If
    '
    EPAC$ = TRIM$(UCase$(EMPREAC$))
    If InStr(EPAC$, "MINEPARTS") < 1 And InStr(EPAC$, "AHP") > 0 Then
      CONDI$ = "CODINT = " & CI1X
      Call DuplicarRegistroDeUnaBaseAOtra("AHP", "MASTER", "MINEPARTS", "MASTER", CONDI$)
    End If
    
End Sub

Function ESEQUIPO_EYM%(CI1%)
    ESEQUIPO_EYM% = 0
    Static LATABLAEXISTE%
    '
    If EXISTE_TABLA%("COMBINA_IVA") > 0 Then
          LATABLAEXISTE% = 1
    Else
          LATABLAEXISTE% = -1
          Exit Function
    End If
    CONDI$ = "CodintEq = " & CI1%
    If CantRegistros("COMBINA_IVA", CONDI$, "NOMESS") > 0 Then
       ESEQUIPO_EYM% = 1
    End If
End Function

Function ESACCESORIO_EYM%(CI1%)
    ESACCESORIO_EYM% = 0
    Static LATABLAEXISTE%
    '
    If EXISTE_TABLA%("COMBINA_IVA") > 0 Then
          LATABLAEXISTE% = 1
    Else
          LATABLAEXISTE% = -1
          Exit Function
    End If
    CONDI$ = "CodintACC = " & CI1%
    If CantRegistros("COMBINA_IVA", CONDI$, "NOMESS") > 0 Then
       ESACCESORIO_EYM% = 1
    End If
End Function
Function EXISTE_ESPECIFIC_ENVAPLAST%(CI1%)
   If EXISTE_TABLA%("ESPECIFIC") > 0 Then
      If CantRegistros("ESPECIFIC", "COD_INTE = " & CI1%, "NOMESS") > 0 Then
         EXISTE_ESPECIFIC_ENVAPLAST% = 1
      End If
   End If
End Function


Function GENERAR_NUEVA_VERSION$(CI1%)

   COD$ = UCase$(CODEXT$(CI1%))
   POS1% = InStrRev(COD$, "-") 'la ultima posicion
   COD$ = Left$(COD$, POS1%)
   CONDI$ = "CODIGO LIKE '" & COD$ & "%' ORDER BY CODIGO DESC"
   ULTI_VERSION$ = VALOBADA("MASTER", "CODIGO", CONDI$, "NOMESS")
   VERSION_COMPLETA$ = Mid$(ULTI_VERSION$, POS1% + 1, Len(Mid$(ULTI_VERSION$, POS1% + 1)))
   PRIMERA_LETRA$ = Left$(VERSION_COMPLETA$, 1): ASK1% = Asc(PRIMERA_LETRA$)
   If PRIMERA_LETRA$ = "Z" Then
       Call COMUNI("Imposible Asignar Nueva Versión - Genere Nuevo Código "): GENERAR_NUEVA_VERSION$ = "": Exit Function
   End If
   
   CANTI_LETRAS% = Len(VERSION_COMPLETA$)
   
   If CANTI_LETRAS% = 1 Then ' SI TIENE  1 CARACTER
      If ASK1% < 90 Then 'SI ES MENOR A 90 AVANZA 1
         ASK1% = ASK1% + 1
      End If
   End If
   '
   PRIMERA_LETRA$ = Chr$(ASK1%)
   '
   GENERAR_NUEVA_VERSION$ = COD$ & PRIMERA_LETRA$
End Function

Sub ZELECHOAHP(STAT%, Optional ModoOrdenamiento%)
   Static CARGADO%
   If CARGADO% > 0 And IsFormLoaded(FRMZELECHO) Then GoTo 50
   CARGADO% = 1
   CONDI$ = ""
   If STAT% > 0 Then CONDI$ = "Status = " & STAT%
   FIN& = CantRegistros("MASTER", CONDI$, "NOMESS")
   FIN& = FIN& + CantRegistros("MASTER2", CONDI$, "NOMESS")
   
   SQLX$ = " SELECT CODINT, CODIGO, DESCRIPCION FROM MASTER WITH(NOLOCK)"
   If STAT% > 0 Then SQLX$ = SQLX$ + " WHERE " & CONDI$
   SQLX$ = SQLX$ + " UNION ALL  "
   SQLX$ = SQLX$ + " SELECT CODINT, CODIGO, DESCRIPCION FROM MASTER2 WITH(NOLOCK)"
   If STAT% > 0 Then SQLX$ = SQLX$ + " WHERE STATUS = " & STAT%
   If ModoOrdenamiento% = 0 Then
        SQLX$ = SQLX$ + " ORDER BY DESCRIPCION"
   ElseIf ModoOrdenamiento% = 1 Then
        SQLX$ = SQLX$ + " ORDER BY CODIGO"
   End If
   
   Campos$ = "ci/F0;Código/A20;Descripción/A48"
   Call CARGA_ENCABEZADO(FRMZELECHO.MSF1, Campos$, FRMZELECHO, , , ANTOT)
   Call CARGA_ENCABEZADO(FRMZELECHO.msf2, Campos$, FRMZELECHO, , , ANTOT)
   FRMZELECHO.Width = 9000
   FRMZELECHO.Caption = "Tabla de Artículos (" & FIN& & " Registros)"
   FRMZELECHO.msf2.Rows = FIN& + 2

   Set rs = READSET(SQLX$)
   FRMZELECHO.msf2.Redraw = False
   With rs
      Do While Not .EOF
        i& = i& + 1
        FRMZELECHO.msf2.TextMatrix(i&, 1) = XVALO(!CODINT)
        FRMZELECHO.msf2.TextMatrix(i&, 2) = SAFETEXT$(!Codigo)
        FRMZELECHO.msf2.TextMatrix(i&, 3) = SAFETEXT$(!Descripcion)
        .MoveNext
      Loop
    End With
    rs.Close
    Set rs = Nothing
    FRMZELECHO.msf2.Redraw = True
    '
50  FRMZELECHO.Show 1
    
    CI1 = XVALO(RESP_ZELE_VECT(1))
    CODI$ = RESP_ZELE_VECT(2)
    '
   '
   
   

End Sub


Sub Agregar_Registro_Master_A_Master2(CIMAST, VALIDAFORMULA%)
     'Ferevis,Observa,Especificacion,DesCompAut,DescriComp,DepoEntra,DepoCon,Codigo,Descripcion,Descridos
     'Status,Codint,Alias,Umedida,Timate,Critrep,Monecos,Fecosto,PorDescCos,Cosuni,Pornacio,PorComVen,Provhabi,Stomin
     'Stomax,Lorepos,Lindivis,Pesuni,PesMetro
     'Color,Talle,Descompo,Grupiva,Grucoven
     'CLinea,Cfamilia,NPlano,Revision,Trazable
     'Revi_Formu,FechApFor,UsuApFor,Cod_Interna,Bien_Cap
     'Codigo_Ncm,Codigo_Sec,Pro_Umed,Iva_Id,ArchiPdf
     'Familia,Lindiven,CanDeci,ConsuProm,CosCal
     'FeCosCal,Stat_Movi,Clas_ATC,Vers_BC,FeVig_BC
     'FeRev_BC,CantApro,LeTiCom,PiezasXCaja,M3Caja,KiloNetoCaja,KiloBrutoCaja,Autono_Proy,UsuPrePes
     'UsuPreDol,MANO_OBRA_REAL,Tamano_Etiqueta,CueContaI,Molde
     
     '********************************************************************************
     'NO CONSIDERO ACA EL TEMA DE MINEPARTS POR QUE LAS ALTAS SE DEBEN DAR SOLO DE AHP
     '********************************************************************************
     
    CI1X = XVALO(CIMAST)
    If Abs(CI1X) <= 0 Then Exit Sub
    '
    'LOS ITEMS QUE EXISTEN EN FORMULAS O POSEEN ESTRUCURA NO SE ELIMINAN
    If VALIDAFORMULA% > 0 Then
       'VALIDA SI CONTIENE FORMULA
       CONDI$ = "CODICONJ = " & CIMAST
       If CantRegistros("FORMULAS", CONDI$, "NOMESS") > 0 Then Exit Sub
       'VALIDA SI ES PARTE DE UNA FORMULA
       CONDI$ = "CODIELEM = " & CIMAST
       If CantRegistros("FORMULAS", CONDI$, "NOMESS") > 0 Then Exit Sub
    End If
    '
    'ESTOS 2 PARRAFOS ES UNA SOLA CONSULTA PARA COPIAR UN REGISTRO DE LA TABLA MASTER A LA TABLA MASTER2
    SQLX$ = "INSERT INTO MASTER2(Ferevis,Observa,Especificacion,DesCompAut,DescriComp,DepoEntra,DepoCon,Codigo,Descripcion,Descridos,"
    SQLX$ = SQLX$ + " Status , CODINT, Alias, Umedida, TIMATE, CRITREP, Monecos, FECOSTO, PorDescCos, COSUNI, PORNACIO, PorComVen, PROVHABI, STOMIN,"
    SQLX$ = SQLX$ + " Stomax , Lorepos, LINDIVIS, PESUNI, PESMETRO,Color,Talle,Descompo,Grupiva,Grucoven,CLinea,Cfamilia,NPlano,Revision,Trazable,"
    SQLX$ = SQLX$ + " Revi_Formu , FechApFor, UsuApFor, Cod_Interna, Bien_Cap, Codigo_Ncm, Codigo_Sec, Pro_Umed, Iva_Id, ArchiPdf,"
    SQLX$ = SQLX$ + " Familia,Lindiven,CanDeci,ConsuProm,CosCal,FeCosCal,Stat_Movi,Clas_ATC,Vers_BC,FeVig_BC,"
    SQLX$ = SQLX$ + " FeRev_BC,CantApro,LeTiCom,PiezasXCaja,M3Caja,KiloNetoCaja,KiloBrutoCaja,Autono_Proy,UsuPrePes,"
    SQLX$ = SQLX$ + " UsuPreDol,MANO_OBRA_REAL,Tamano_Etiqueta,CueContaI,Molde)"
    '
    'VER QUE LA INFO SE MIGRA CON CODINT= 0 YA QUE CUANDO VUELVA A LA TBL MASTER VA A TOMAR EL CODIGO INTERNO CORRESPONDIENTE.
    SQLX$ = SQLX$ + " select Ferevis,Observa,Especificacion,DesCompAut,DescriComp,DepoEntra,DepoCon,Codigo,Descripcion,Descridos,"
    SQLX$ = SQLX$ + " Status , 0, Alias, Umedida, TIMATE, CRITREP, Monecos, FECOSTO, PorDescCos, COSUNI, PORNACIO, PorComVen, PROVHABI, STOMIN,"
    SQLX$ = SQLX$ + " Stomax , Lorepos, LINDIVIS, PESUNI, PESMETRO,Color,Talle,Descompo,Grupiva,Grucoven,CLinea,Cfamilia,NPlano,Revision,Trazable,"
    SQLX$ = SQLX$ + " Revi_Formu , FechApFor, UsuApFor, Cod_Interna, Bien_Cap, Codigo_Ncm, Codigo_Sec, Pro_Umed, Iva_Id, ArchiPdf,"
    SQLX$ = SQLX$ + " Familia,Lindiven,CanDeci,ConsuProm,CosCal,FeCosCal,Stat_Movi,Clas_ATC,Vers_BC,FeVig_BC,"
    SQLX$ = SQLX$ + " FeRev_BC , CantApro, LeTiCom, PiezasXCaja, M3Caja, KiloNetoCaja, KiloBrutoCaja, Autono_Proy, USUPREPES, "
    SQLX$ = SQLX$ + " UsuPreDol,MANO_OBRA_REAL,Tamano_Etiqueta,CueContaI,Molde FROM MASTER WHERE CODINT = " & CI1X
    On Error Resume Next
    FLEXCONN.Execute (SQLX$)
     If Err Then
       On Error GoTo 0
       Call COMUNI("No Se Realizó La Transacción para El Código Interno: " & CI1X)
       Exit Sub
    End If
    '
    'SI PUDO DAR DE ALTA EN LA TABLA MASTER2 BORRA DE LA TABLA MASTER.
    SQLX$ = "DELETE FROM MASTER WHERE CODINT = " & CI1X
    FLEXCONN.Execute (SQLX$)
    If Err Then
       On Error GoTo 0
       Call COMUNI("No Se Realizó El Borrado para El Código Interno: " & CI1X)
       Exit Sub
    End If
    'SI NO HAY ERROR REALIZA LA TRANSACCION
End Sub


'
Function CODINT_NEW()
   'ESTA CONSULTA DEVUELVE EL PRIMER HUECO
   'SE ORDENA LA TABLA MASTER SEGUN CODINT Y SE GENERA OTRO CAMPO QUE ENUMERA EN FORMA ASCENDENTE TAMBIEN
   'ARRANCANDO A PARTIR DE DONDE EL CODINT ES DIFERENTE AL ORDEN ENTONCES EL PRIMERO SIEMPRE VA A SER EL PRIMER HUECO.
   'SI NO DEVUELVE NADA ES POR QUE NO HAY HUECOS, EN ESE CASO LE ASIGNA EL VALOR DE NUMAS + 1
'   sqlx$ = "SELECT TOP 1 ORDEN, CODINT From (SELECT ROW_NUMBER() OVER (ORDER BY CODINT) AS Orden, CODINT  FROM MASTER WITH(NOLOCK)) T1"
'   sqlx$ = sqlx$ + " GROUP BY CODINT, ORDEN Having CODINT <> ORDEN"
'   Set RS = READSET(sqlx$)
'   With RS
'     If Not .EOF Then
'        CODINT_NEW = XVALO(!ORDEN)
'     Else
'        CODINT_NEW = NUMAS% + 1
'     End If
'   End With
'   RS.Close
'   Set RS = Nothing
'Function CODINT_NEW()
   'ESTA CONSULTA DEVUELVE EL PRIMER HUECO
   'SE ORDENA LA TABLA MASTER SEGUN CODINT Y SE GENERA OTRO CAMPO QUE ENUMERA EN FORMA ASCENDENTE TAMBIEN
   'ARRANCANDO A PARTIR DE DONDE EL CODINT ES DIFERENTE AL ORDEN ENTONCES EL PRIMERO SIEMPRE VA A SER EL PRIMER HUECO.
   'SI NO DEVUELVE NADA ES POR QUE NO HAY HUECOS, EN ESE CASO LE ASIGNA EL VALOR DE NUMAS + 1
   If VersionBD% <= 2000 Then 'SE AGREGA ESTO POR QUE SI ES VERSION 2000 NO LO ACEPTA
      CODINT_NEW = NUMAS% + 1 'CASO EYM
      Exit Function
   End If
   
   SQLX$ = "SELECT TOP 1 ORDEN, CODINT From (SELECT ROW_NUMBER() OVER (ORDER BY CODINT) AS Orden, CODINT  FROM MASTER WITH(NOLOCK)) T1"
   SQLX$ = SQLX$ + " GROUP BY CODINT, ORDEN Having CODINT <> ORDEN"
   Set rs = READSET(SQLX$)
   With rs
    
     If Not .EOF Then
        
        CODINT_NEW = XVALO(!ORDEN)
     Else
        CODINT_NEW = NUMAS% + 1
     End If
   End With
   
   rs.Close
   Set rs = Nothing
   CONDI$ = "CODINT = " & CODINT_NEW
   
   If CantRegistros("MASTER", CONDI$, "NOMESS") > 0 Then
      CODINT_NEW = NUMAS% + 1
   End If
End Function


Function CostoRecpMovisto(CI1%, RECEP&, Optional MONEX%)
'TENER EN CUENTA QUE EN EL CASO QUE EN UNA MISMA ORDEN DE RECEPCION HAYA MAS DE UNA ORDEN DE COMPRA
'PUEDE DEVOLVER VALORES QUE NO CORRESPONDAN CUANDO SE REPITA EL CODIGO  CON DIFRENTEES PRECIOS,
'ESTO ES RARO QUE SUCEDA EN AHP POR QUE LA MAYORIA SON DOLARES Y NO DEBERIA CAMBIAR EL MISMO,
'AVISAR QUE HAGAN UNA RECEPCION POR ORDEN DE COMPRA

   CONDI$ = "COD_INTE = " & CI1%
   CONDI$ = CONDI$ + " AND SUBSTRING(DOPRICOR,4,7) = " & RECEP&
   CONDI$ = CONDI$ + " AND CODIMOVI = 'BR'"
   CONDI$ = CONDI$ + " AND ORDER BY IDENLOTE DESC"
   MON% = XVALO(VALOBADA("MOVISTO", "MONEVAL", CONDI$, "NOMESS"))
   TIPCAM# = XVALO(VALOBADA("MOVISTO", "TIPOCAM", CONDI$, "NOMESS"))
   COSMOVISTO# = XVALO(VALOBADA("MOVISTO", "VALOUNIT", CONDI$, "NOMESS"))
   COEFCON = 1
   If MONEX% <> MON% And MONEX% > 0 Then
      COEFCON = TICAMONE(MON%) / TICAMONE(MONEX%)
   End If
   COSMOVISTO# = COSMOVISTO# * COEFCON
   
   'SI ES 100 DOLARES Y PASO MONEDA $ = COSMOVISTO# * 9  = 900
   
End Function

Sub ActualizarCostoMovisto(CI1%, NBORE&, NUEVOVALOR#, MONEDAX%)
   'ACTUALIZA EL PRECIO DEL MOVISTO VERIFICA SI ES LA MISMA MONEDA
   CONDI$ = "COD_INTE = " & CI1%
   CONDI$ = CONDI$ + " AND SUBSTRING(DOPRICOR,4,7) = " & RECEP&
   CONDI$ = CONDI$ + " AND CODIMOVI = 'BR'"
   Call ACTUREGISTRO("MOVISTO", "VALOUNIT", CONDI$, FORMATNUM$(NUEVOVALOR#, "F16.5"), REGAF&, "NOMESS")
   Call ACTUREGISTRO("MOVISTO", "MONEVAL", CONDI$, MONEDAX%, REGAF&, "NOMESS")
   If MONEDAX% = 1 Then Call ACTUREGISTRO("MOVISTO", "TIPOCAM", CONDI$, 1, REGAF&, "NOMESS")
End Sub
Sub AGRECOD(CODINUE$, DESCRINUE$, TIMA%, UNID$, OKX%, Optional TRAZ%, Optional NPLANO$)
   '
   Screen.MousePointer = 11
   '
   'validaciones
   If Len(CODINUE$) < 3 Or Len(CODINUE$) > 15 Then
        Call MENSERR(24, "El Codigo de Articulo no debe ser Inferior a tres Cararcteres. No se ha podido dar del Alta el Articulo.")
        Screen.MousePointer = 1
        GoTo 99
   End If
   '
   If Len(DESCRINUE$) <= 0 Then
        Call MENSERR(24, "El Articulo no tiene Descripcion. No se ha Podido dar de Alta el Articulo.")
        Screen.MousePointer = 1
        GoTo 99
   End If
   '
   CODII% = CODINT_NEW
'   CODII% = NUMAS% + 1
'20 On Error Resume Next
'   Call APERBASDAT("MASTER")
'   SQLX$ = "SELECT MAX(CODINT)AS CODIMAXIM FROM MASTER "
'   Dim MAXIMMMO As ADODB.Recordset
'   Set MAXIMMMO = FLEXCONN.Execute(FILTSQL$(SQLX$))
'   With MAXIMMMO
'      If !CODIMAXIM < 1 Then
'         CODII% = 1
'       Else
'         CODII% = XVALO(!CODIMAXIM) + 1
'      End If
'   End With
'   MAXIMMMO.Close
'   Set MAXIMMMO = Nothing
   '
   SQLX$ = "SELECT * FROM MASTER "
   SQLX$ = SQLX$ + " WHERE CODIGO = '" & TRIM$(CODINUE$) & "' "
   Dim COCODTMP As ADODB.Recordset
   Set COCODTMP = New ADODB.Recordset
25 On Error Resume Next
   COCODTMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenDynamic, adLockPessimistic, adCmdText
   If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
      On Error GoTo 0
      Call CapturaErrorOpen
      GoTo 25
   End If
   On Error GoTo 0
   '
   With COCODTMP
     On Error Resume Next
     .MoveFirst
     If Err Then 'el codigo no es pre-existente LO aGREGO
       On Error GoTo 0
       SQLX$ = "SELECT * FROM MASTER "
       Set MASTTTER = New ADODB.Recordset
26     On Error Resume Next
       MASTTTER.Open FILTSQL$(SQLX$), FLEXCONN, adOpenDynamic, adLockPessimistic, adCmdText
       If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
           On Error GoTo 0
           Call CapturaErrorOpen
           GoTo 26
       End If
       On Error GoTo 0
       '
       With MASTTTER
         .AddNew   ' YA CONVERTIDO
         On Error Resume Next
         !CODIEMPR = CodiEmp%
         On Error GoTo 0
           !Codigo = CODINUE$
           !Descripcion = DESCRINUE$
           !Status = 1
           !CODINT = CODII%
           !UMEDIDA = UNID$
           !TIMATE = TIMA%
              If TRAZ% <> 1 Then TRAZ% = 0
           !TRAZABLE = TRAZ%
           .Update
       End With
       '
       UPDATMASTER% = 1
       Call GENECOMAS
       OKX% = 1
       Exit Sub
       '
     End If
     On Error GoTo 0
     Call MENSERR(24, "Código de Articulo no válido o Pre-Existente. No se ha Podido dar de Alta el Articulo.")
     GoTo 99
   End With
   '
99 Screen.MousePointer = 1
End Sub
Function IDCalculoSegunFechaOF(FEXC$)
   FECHG% = CVINT(FEXC$)
   CONDI$ = " FECALCULO < '" & FETEXCOR3$(FECHG%, FEXC$) & "' ORDER BY FECALCULO DESC"
   IDCalculoSegunFechaOF = XVALO(VALOBADA("COBERFA", "TOP 1 ID", CONDI$, "NOMESS"))
End Function
Function IDCalculoSegunFechaOC(NUMOC&)
   CONDI$ = " NUME_OCOM = " & NUMOC& '
'   FEYHO$ = "16/11/2016 05:48:00 P.M."
   FEYHO$ = VALOBADA("OCOMENCA", "FGEN_OCOM ", CONDI$, "NOMESS")
'   MsgBox "FECHA Y HORA DE TABLA OCOMENCA  " & vbCrLf & FEYHO$
   
   FECHG% = CVINT(FEYHO$)
   CONDI$ = " FECALCULO < '" & FETEXCOR3$(FECHG%, FEYHO$) & "' ORDER BY FECALCULO DESC"
'   MsgBox "CONDICION DE CONSULTA (IDCalculoSegunFechaOC) " & vbCrLf & CONDI$
   IDCalculoSegunFechaOC = XVALO(VALOBADA("COBERFA", "TOP 1 ID", CONDI$, "NOMESS"))
End Function


'
Sub GENECOMAS()
  '
      If UPDATMASTER% > 0 Then
        '
        Screen.MousePointer = 11
        '
        CONDI$ = " CODIGO <> '' AND STATUS > 0 ORDER BY CODIGO ASC"
        If Control$("", "IDSECUN") = "ALIAS" Then
            Call CARGALISEL("ECOMAST", "MASTER", "CODIGO/A24;DESCRIPCION/A36;NPLANO/A24", CONDI$, "NOMESS")
        Else
            Call CARGALISEL("ECOMAST", "MASTER", "CODIGO/A24;DESCRIPCION/A50", CONDI$, "NOMESS")
        End If
        '
      End If
      '
      If UPDATMASTER% > 0 Or UPDAPLANOS% > 0 Then
          CONDI$ = " CODIGO <> '' AND STATUS > 0 AND NPLANO <> ''"
          Call CARGALISEL("IPLANOS", "MASTER", "NPLANO/A24;DESCRIPCION/A54;CODINT/F5.0", CONDI$, "NOMESS")
      End If
      '
      
      UPDATMASTER% = 0
      UPDAPLANOS% = 0
      Screen.MousePointer = 1
      '

    'Set MASTERECO = Articulos.OpenRecordset("SELECT CODIGO,DESCRIPCION,CODINT,NPLANO FROM Master WHERE STATUS > 0 ORDER BY CODIGO ASC;", dbOpenDynaset, dbReadOnly)
    'sql
''    Dim MASTERECO As ADODB.Recordset
''    Set MASTERECO = New ADODB.Recordset
''    MASTERECO.CursorType = adOpenStatic
''    MASTERECO.LockType = adLockReadOnly
''
''    STRSQL = "SELECT CODIGO,DESCRIPCION,CODINT,NPLANO FROM Master WITH(NOLOCK) WHERE STATUS > 0 ORDER BY CODIGO ASC"
''25  On Error Resume Next
''    MASTERECO.Open FILTSQL$(STRSQL), FLEXCONN, , , adCmdText
''    If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
''      On Error GoTo 0
''      Call CapturaErrorOpen
''      GoTo 25
''    End If
''    On Error GoTo 0
''    '
''    CAREC& = CantRegistros("MASTER", "STATUS > 0")
''    If CAREC& < 1 Then GoTo 15
''    '
''    FIN& = CAREC&
''    REBLA$ = REGBLAN$("ECOMAST")
''    JJ& = 0: JJ1& = 0
''    '
''10  Call TRACE(20, JJ&, FIN&)
''    EL1$ = REBLA$
''    JJ1& = JJ1& + 1
''    CODEXX$ = MASTERECO!CODIGO
''    '
''    If TRIM$(CODEXX$) <> "" Then
''      LOFI1% = LENFIELD%("ECOMAST", 1)
''      POFI2% = POSFIELD%("ECOMAST", 2)
''      POFI3% = POSFIELD%("ECOMAST", 3)
''      Call REPLA(EL1$, CODEXX$, 1, LOFI1%)
''        DEPLANO$ = TRIM$(MASTERECO!Descripcion)
''        If IsNull(MASTERECO!NPLANO) = False Then
''          NXPLANO$ = TRIM$(MASTERECO!NPLANO)
''          If NXPLANO$ <> "" Then
''            'If NXPLANO$ <> TRIM$(MASTERECO!CODIGO) Then
''            If InStr(TRIM$(MASTERECO!CODIGO), NXPLANO$) < 1 Then
''              While Len(DEPLANO$ + NXPLANO$) > 43
''                DEPLANO$ = Left$(DEPLANO$, Len(DEPLANO$) - 1)
''              Wend
''              DEPLANO$ = DEPLANO$ + " (" + NXPLANO$ + ")"
''            End If
''          End If
''        End If
''      Call REPLA(EL1$, DEPLANO$, POFI2%, 46)
''      If MASTERECO!CODINT <= 32767 Then
''          Call REPLA(EL1$, MKI$(MASTERECO!CODINT), POFI3%, 2)
''      End If
''      JJ& = JJ& + 1
''      Call GRAREG("ECOMAST", EL1$, JJ&)
''    End If
''    '
''12  If JJ1& < CAREC& Then
''      MASTERECO.MoveNext
''      GoTo 10
''    End If
''    '
''    MASTERECO.Close
''    Call SETULTREG("ECOMAST", JJ&)
''    Call CIERRARCH("ECOMAST")
''    Call FRESHECHO("ECOMAST")
''    Call BAJALDISCO
''  End If
  '
''15 If UPDATMASTER% > 0 Or UPDAPLANOS% > 0 Then
''    Screen.MousePointer = 11
''
''    'Set MASTERECO = Articulos.OpenRecordset("SELECT NPlano,CODIGO,DESCRIPCION,CODINT FROM Master WHERE STATUS > 0 AND NPLANO <> '' ORDER BY NPLANO ASC;", dbOpenDynaset, dbReadOnly)
''    SQLX$ = " select NPlano,CODIGO,DESCRIPCION,CODINT"
''    SQLX$ = SQLX$ + " from Master"
''    SQLX$ = SQLX$ + " WHERE STATUS > 0 "
''    SQLX$ = SQLX$ + " AND NPLANO <> ''"
''    ' SQLX$ = SQLX$ + " AND Npro_Ocom = " & NC% & " "
''
''    'Dim MASTERECO As ADODB.Recordset
''    Set MASTERECO = FLEXCONN.Execute(FILTSQL$(SQLX$))
''
''    On Error Resume Next
''    MASTERECO.MoveLast
''    If Err Then GoTo 99
''    '
''    CAREC& = MASTERECO.RecordCount
''    If CAREC& < 1 Then GoTo 99
''    '
''    FIN& = CAREC&
''    REBLA$ = REGBLAN$("IPLANOS")
''    JJ& = 0
''    MASTERECO.MoveFirst
''    '
''20  Call TRACE(20, JJ&, FIN&)
''    EL1$ = REBLA$
''    Call REPLA(EL1$, MASTERECO!NPLANO, 1, 24)
''    Call REPLA(EL1$, MASTERECO!Descripcion, 25, 54)
''    Call REPLA(EL1$, MKI$(MASTERECO!CODINT), 79, 2)
''    JJ& = JJ& + 1
''    Call GRAREG("IPLANOS", EL1$, JJ&)
''    '
''    If JJ& < CAREC& Then
''      MASTERECO.MoveNext
''      GoTo 20
''    End If
''    '
''    MASTERECO.Close
''    Set MASTERECO = Nothing
''    Call SETULTREG("IPLANOS", JJ&)
''    Call CIERRARCH("IPLANOS")
''    Call FRESHECHO("IPLANOS")
''    Call BAJALDISCO
''   End If
''   '
''99 UPDATMASTER% = 0
''   UPDAPLANOS% = 0
''   Screen.MousePointer = 1
End Sub
Sub GENEECOMABAJ()
    CONDI$ = " CODIGO <> '' AND STATUS < 0 ORDER BY CODIGO ASC"
    Call CARGALISEL("ECOMABAJ", "MASTER", "CODIGO/A24;DESCRIPCION/A50", CONDI$, "NOMESS")
End Sub

'
Function NUMAS%()
   '
   If INCLUALIAS% = 0 Then
     INCLUALIAS% = (-1)
     If Control$("", "IDSECUN") = "ALIAS" Then INCLUALIAS% = 1
   End If
   '
   Call APERBASDAT("MASTER")
   On Error Resume Next
   
   Dim MIRS As ADODB.Recordset
   Set MIRS = New ADODB.Recordset
   strSQL = ("SELECT max(codint) as maximo FROM MASTER")
   Set MIRS = FLEXCONN.Execute(FILTSQL$(strSQL))
   If Not IsNull(MIRS!Maximo) Then
      NUMAS% = MIRS!Maximo
   Else
      NUMAS% = 0
   End If
   MIRS.Close
   Set MIRS = Nothing
   
End Function

Function PRICOD%()
   '
   PRICOD% = 0
   Call APERBASDAT("MASTER")
   'Set MasterTp = Articulos.OpenRecordset("SELECT Codigo, Status, Codint FROM Master WHERE Status >=0 order by Codigo ASC", dbOpenSnapshot)
    strSQL = "SELECT MIN(Codigo)as Primero FROM Master WHERE Status >=0  "
    Dim MasterTp As ADODB.Recordset
    Set MasterTp = FLEXCONN.Execute(FILTSQL$(strSQL))
    If Not IsNull(MasterTp!Primero) Then 'si no esta vacio
       PRICOD% = CODINT%(MasterTp!Primero)
    End If
    On Error GoTo 0
    MasterTp.Close
    Set MasterTp = Nothing
End Function

Function ULTCOD%()
   '
   ULTCOD% = 0
   Call APERBASDAT("MASTER")
   'Set MasterTp = Articulos.OpenRecordset("SELECT Codigo, Status, Codint FROM Master WHERE Status >=0 order by Codigo ASC", dbOpenSnapshot)
   strSQL = "SELECT MAX(Codigo) as Ultimo FROM Master WHERE Status >=0 "
   Dim MasterTp As ADODB.Recordset
   Set MasterTp = FLEXCONN.Execute(FILTSQL$(strSQL))
   If Not IsNull(MasterTp!Ultimo) Then 'si no es null (devuelve un sola fila)
      ULTCOD% = CODINT%(MasterTp!Ultimo)
    End If
   On Error GoTo 0
   MasterTp.Close
   Set MasterTp = Nothing
   '
End Function

Function STATUIT%(Ci%)
   Static CIANT%, STATUIA%, FIDAMASANT#
   '
   ' = 1 --> Existe
   ' = 0 --> No existe el item
   '
   If STATUIA% < 1 Then
     Call APERBASDAT("MASTER")
   End If
   '
   If Ci% = CIANT% Then
     If STATUIA% > 0 Then
       STATUIT% = STATUIA%
       On Error Resume Next
       CIAA% = MasterTemp!CODINT
       On Error GoTo 0
       If Ci% = CIAA% Then Exit Function
     End If
   End If
   '*********NOTA : SI NO ENTRA AQUI DA  UN ERROR EN CODEXT **********
   STATUIT% = 0
   If Ci% > 0 Then
     strSQL = "SELECT * FROM Master WITH(NOLOCK) WHERE CODINT=" & Ci%
     Set MasterTemp = FLEXCONN.Execute(FILTSQL$(strSQL))
     If Not (MasterTemp.EOF And MasterTemp.BOF) Then 'si no esta vacio
         STATUIT% = 1
     End If
     On Error GoTo 0
   End If
   '
   CIANT% = Ci%
   STATUIA% = STATUIT%
   '
End Function
 
Function STATITEM%(Ci%)
'\\\OT-06-0353-MH-29/08/06///
   ' DEVUELVE  1 SI ESTÁ ACTIVO
   '          -1 SI ESTA DE BAJA
   On Error Resume Next
   If STATUIT%(Ci%) <> 0 Then STATITEM% = MasterTemp!Status
'\\\OT-06-0353-MH-FIN///
End Function
'
Function CODINT%(CODEXTE$)
   '
   Static FIDAMASANT#, LONCODI%
   '
   YaControloMaster2% = 0
10 CODINT% = 0
   If LONCODI% < 1 Then
      LONCODI% = XVALO(Control$("MASTER", "LOMAXCOD"))
      If LONCODI% < 16 Or LONCODI% > 24 Then LONCODI% = 16
   End If
   '
   If INCLUALIAS% = 0 Then
     INCLUALIAS% = (-1)
     If Control$("", "IDSECUN") = "ALIAS" Then INCLUALIAS% = 1
   End If
   '
   If CODEXTE$ <> "" Then
     '
     Call APERBASDAT("MASTER")
     On Error Resume Next
     CDXX$ = TRIM$(CODEXTE$): CDYY$ = AJUSTI$(CODEXTE$, LONCODI% - 1) + Chr$(96)
     SQLX$ = "Select * from Master where CODIGO ='" + CDXX$ + "' OR CODIGO = '" & CDYY$ & "' "
     '
     If INCLUALIAS% = 1 Then SQLX$ = SQLX$ + "OR NPlano = '" + CDXX$ + "' "
     Set MasterTemp = FLEXCONN.Execute(FILTSQL$(SQLX$))
     If Not (MasterTemp.EOF And MasterTemp.BOF) Then
        CODINT% = MasterTemp!CODINT * MasterTemp!Status
     End If
     '
     If CODINT% = 0 Then
       If EXISMASTER2%(CODEXTE$) > 0 And YaControloMaster2% < 1 Then
         Call AltaM2AMaster(CODEXTE$)
         YaControloMaster2% = 1 ' por las dudas para que no entre en un bucle
         GoTo 10
       End If
     End If

   End If
   '
End Function

Function CODINPLA%(NUPLAX$)
   '
   If NUPLAX$ <> "" Then
     Call APERBASDAT("MASTER")
     strSQL = "SELECT codint FROM MASTER WHERE NPLANO ='" & TRIM$(NPLAX$) + "'"
     Dim MIRS As ADODB.Recordset
     Set MIRS = FLEXCONN.Execute(FILTSQL$(strSQL))
     On Error Resume Next
     If Not (MIRS.EOF And MIRS.BOF) Then
      CODINPLA% = MIRS!CODINT
     End If
     MIRS.Close
     Set MIRS = Nothing
   
     'On Error Resume Next
     'Master.MoveFirst
     'Master.FindFirst "NPLANO ='" + TRIM$(NPLAX$) + "'"
     'If Master.NoMatch = False Then
     '  CODINPLA% = Master!CODINT
     'End If
     On Error GoTo 0
   End If
   '
End Function

Function CODEXT$(Ci%)
   '
   If Ci% = 0 Then
        CODEXT$ = ""
     Else
        If STATUIT%(Ci%) <> 0 Then
            On Error Resume Next
            CODEXT$ = TRIM$(MasterTemp!Codigo)
            If Err Then GoTo 20
          Else
20          On Error GoTo 0
            Call APERBASDAT("MASTER")
            strSQL = "SELECT Codigo FROM MASTER WHERE Codint = " & Ci%
            Dim MIRS As ADODB.Recordset
            Set MIRS = FLEXCONN.Execute(FILTSQL$(strSQL))
            If Not (MIRS.EOF And MIRS.BOF) Then
              CODEXT$ = MIRS!Codigo
            End If
            MIRS.Close
            Set MIRS = Nothing
        End If
   End If
   '
End Function
 
Function DESCRIT$(Ci%)
   On Error Resume Next
   DESCRIT$ = ""
   If Ci% < 1 Then
       If InStr(UCase$(App.EXEName), "WINTAB") < 1 Then
          DESCRIT$ = "Material No Codificado"
       End If
     Else
       DESCRIT$ = DESCRIT0$(Ci%)
       If TRIM$(DESCRIT$) <> "" Then DESCRIT$ = DESCRIT$ + " (" + Unimed$(Ci%) + ")"
   End If
End Function

Function DESCRIT0$(Ci%)
   On Error Resume Next
   DESCRIT0$ = ""
   If Ci% < 1 Then
       If InStr(UCase$(App.EXEName), "WINTAB") < 1 Then
         DESCRIT0$ = "Material No Codificado"
       End If
     Else
       If STATUIT%(Ci%) <> 0 Then DESCRIT0$ = TRIM$(MasterTemp!Descripcion)
   End If
End Function

Function DESCRIT2$(Ci%)
   On Error Resume Next
   DESCRIT2$ = ""
   If STATUIT%(Ci%) <> 0 Then DESCRIT2$ = TRIM$(MasterTemp!DESCRIDOS)
End Function
Function USUPREPES$(Ci%)
   On Error Resume Next
   USUPREPES$ = ""
   If STATUIT%(Ci%) <> 0 Then USUPREPES$ = TRIM$(MasterTemp!USUPREPES)
End Function
Function USUPREDOL$(Ci%)
   On Error Resume Next
   USUPREDOL$ = ""
   If STATUIT%(Ci%) <> 0 Then USUPREDOL$ = TRIM$(MasterTemp!USUPREDOL)
End Function

Function Unimed$(Ci%)
   '
   On Error Resume Next
   If STATUIT%(Ci%) <> 0 Then Unimed$ = TRIM$(MasterTemp!UMEDIDA)
   If TRIM$(Unimed$) = "" Then Unimed$ = "U."
   '
End Function

Function DEUNIMED$(Ci%)
   UME$ = Unimed$(Ci%)
   DEUNIMED$ = UME$
   DEUNI$ = TRIM$(ECOARCH$("UMEDIDA", UME$))
   If DEUNI$ <> "" Then
      DEUNIMED$ = DEUNI$
   End If
End Function
'
Function CODFAMIL$(Ci%)
   '
   On Error Resume Next
   If STATUIT%(Ci%) <> 0 Then CODFAMIL$ = TRIM$(MasterTemp!familia)
   '
End Function



Function TIMATE%(Ci%)
   On Error Resume Next
   If STATUIT%(Ci%) <> 0 Then TIMATE% = CDbl(MasterTemp!TIMATE)
End Function

Function CRITREP%(Ci%)
   On Error Resume Next
   If STATUIT%(Ci%) <> 0 Then CRITREP% = CDbl(MasterTemp!CRITREP)
End Function

Function GRUCOVEI%(Ci%)
   On Error Resume Next
   If STATUIT%(Ci%) > 0 Then GRUCOVEI% = CDbl(MasterTemp!Grucoven)
End Function

Function GRUPIVA%(Ci%)
   Static GRUMAXI%
   If GRUMAXI% = 0 Then
     GRUMAXI% = ULTREG&("GRUPIVA")
   End If
   On Error Resume Next
   
   'CODIGO AGREGADO PARA CONSIDRERAR IVAS COMBINADOS
   If VALIDA_TASAIVA$ = "SI" Then
    If Ci% > 0 Then
      If XVALO(TASAIVAX_CI(Ci%)) > 0 Then
         URE& = ULTREG&("GRUPIVA")
         'If URE& > 4 Then URE& = 4
         For U& = 1 To URE&
           X$ = REGLEIDO$("GRUPIVA", U&)
           GIV = CVS(Mid$(X$, 13, 4))
           If GIV = TASAIVAX_CI(Ci%) Then
              GRUPIVA% = U& - 1
              Exit For
           End If
         Next U&
         Call CIERRARCH("GRUPIVA")
         Exit Function
     End If
    End If
   End If
   'FIN CODIGO AGREGADO PARA CONSIDRERAR IVAS COMBINADOS

   
   If STATUIT%(Ci%) > 0 Then GRUPIVA% = CDbl(MasterTemp!GRUPIVA)
   If GRUPIVA% > GRUMAXI% Then GRUPIVA% = 0
End Function

Function LOTESTAN(Ci%)
   On Error Resume Next
   If STATUIT%(Ci%) > 0 Then LOTESTAN = CDbl(MasterTemp!LOREPOS)
   If LOTESTAN < 0 Then LOTESTAN = 100 * Abs(LOTESTAN)
End Function
  
Function COSUNI(Ci%)
   '
   On Error Resume Next
   If STATUIT%(Ci%) > 0 Then COSUNI = CDbl(MasterTemp!COSUNI)
   '
End Function
'
Function MONECOSTO%(Ci%)
   '
   MONECOSTO% = 1
   On Error Resume Next
   If STATUIT%(Ci%) <> 0 Then MONECOSTO% = CDbl(MasterTemp!Monecos)
   '
End Function

Function FECOSTO%(Ci%)
   On Error Resume Next
   If STATUIT%(Ci%) <> 0 Then FECOST$ = Format(MasterTemp!FECOSTO, "dd/mm/yy")
   FECOSTO% = FECHANUM(FECOST$)
End Function
'
Function COSPESOS(Ci%)
    COSPESOS = COSUNI(Ci%) * TICAMONE(MONECOSTO%(Ci%)) * (1 + PORNACIO(Ci%) / 100)

End Function
'
Function PORNACIO(Ci%)
   PORNACIO = 0
   On Error Resume Next
   If STATUIT%(Ci%) <> 0 Then PORNACIO = CDbl(MasterTemp!PORNACIO)
   On Error GoTo 0
End Function
'
Function COEFNACIO(Ci%)
   COEFNACIO = 1 + PORNACIO(Ci%) / 100
End Function
'
Function PORCOMVE(Ci%)
   PORCOMVE = 0
   On Error Resume Next
   If STATUIT%(Ci%) <> 0 Then PORCOMVE = CDbl(MasterTemp!PorComVen)
   On Error GoTo 0
End Function
'
Function MANO_OBRA_REAL(Ci%)
   MANO_OBRA_REAL = 0
   On Error Resume Next
   If STATUIT%(Ci%) <> 0 Then MANO_OBRA_REAL = CDbl(MasterTemp!MANO_OBRA_REAL)
   On Error GoTo 0
End Function
'

Function PROVHABI%(Ci%)
   PROVHABI% = 0        ' FABRICACION PROPIA
   On Error Resume Next
   If STATUIT%(Ci%) <> 0 Then PROVHABI% = CDbl(MasterTemp!PROVHABI)
   On Error GoTo 0
   '
   If PROVHABI% > 0 Then
     If CantRegistros("PROVED12", "Nume_Cli = " & PROVHABI%) < 1 Then
       PROVHABI% = 0
     End If
   End If
End Function
 
Function COLOR$(Ci%)
   On Error Resume Next
   If STATUIT%(Ci%) <> 0 Then COLOR$ = TRIM$(MasterTemp!COLOR)
   On Error GoTo 0
End Function
'
Function TALLE$(Ci%)
   On Error Resume Next
   If STATUIT%(Ci%) <> 0 Then TALLE$ = TRIM$(MasterTemp!TALLE)
End Function
'
Function PESUNI(Ci%)
   On Error Resume Next
   If STATUIT%(Ci%) <> 0 Then PESUNI = CDbl(MasterTemp!PESUNI)
End Function
'
Function PESMETRO(Ci%)
   On Error Resume Next
   If STATUIT%(Ci%) <> 0 Then PESMETRO = CDbl(MasterTemp!PESMETRO)
End Function
'
Function LTITPE$(CP)
   'convierte peso unitario en titulo de hilado
   LCLR$ = ""
   If CP > 0 Then
        CANDE% = 5: If CP < 10 Then CANDE% = 7
        LCLR$ = CSTRING$(MKS$(CP), 3, 10, CANDE%, 2)
      ElseIf CP < 0 Then
        LCLR$ = TRIM$(CSTRING$(MKS$(Abs(CP)), 3, 8, 2, 2))
        LCLR$ = REPLACAR$(LCLR$, ".", "/")
   End If
   LTITPE$ = LCLR$
   '
End Function

Function OBSERVAIT$(Ci%)
   On Error Resume Next
   If STATUIT%(Ci%) <> 0 Then OBSERVAIT$ = MasterTemp!Observa
End Function
'
Function ESPECIFIT$(Ci%)
   On Error Resume Next
   If STATUIT%(Ci%) <> 0 Then ESPECIFIT$ = MasterTemp!ESPECIFICACION
End Function
'
Function NUPLANO$(Ci%)
   '
   NUPLANO$ = ""
   On Error Resume Next
   If STATUIT%(Ci%) <> 0 Then NUPLANO$ = TRIM$(MasterTemp!NPLANO)
   '
End Function
Function CODINTERNA$(Ci%)
   '
   CODINTERNA$ = ""
   On Error Resume Next
   If STATUIT%(Ci%) <> 0 Then CODINTERNA$ = TRIM$(MasterTemp!Cod_Interna)
   '
End Function

'\\\OT-05-0702-WAR-01/11/05///
Function ATRIB1$(Ci%) 'ATRIBUTO UNO
   '
   ATRIB1$ = ""
   On Error Resume Next
   If STATUIT%(Ci%) <> 0 Then ATRIB1$ = TRIM$(MasterTemp!TALLE)
   '
End Function
Function ATRIB2$(Ci%) 'ATRIBUTO DOS
   '
   ATRIB2$ = ""
   On Error Resume Next
   If STATUIT%(Ci%) <> 0 Then ATRIB2$ = TRIM$(MasterTemp!COLOR)
   '
End Function
'
Function ARMEMOTEC$(Ci%)
   '
   ARMEMOTEC$ = VALOBADA("MASTER", "ARCHIPDF", "CODINT=" & Ci%)
   '
End Function
'
Function CONSUMOPROMEDIO#(Ci%)
   '
   CONSUMOPROMEDIO# = XVALO(VALOBADA("MASTER", "consuprom", "CODINT=" & Ci%))
   '
End Function
Function COSCALC#(Ci%)
   '
   COSCALC# = XVALO(VALOBADA("MASTER", "CosCal", "CODINT=" & Ci%))
   '
End Function
Function FECOSCAL%(Ci%)
   '
   FECOSCAL% = FECHANUM("01/01/80")
   On Error Resume Next
   FECOSCAL% = CVINT%(VALOBADA("MASTER", "FeCosCal", "CODINT=" & Ci%))
   On Error GoTo 0
   '
End Function
'
Function UNICOM$(Ci%, Optional NPROVE%)
   '
   UNICO$ = ""
   '
   Call ABRECOMPRAS
   If Ci% > 0 Then  ' que no lo haga con material libre
     SQLX$ = "SELECT * FROM OCOMDETA "
     SQLX$ = SQLX$ + "WHERE Cod_Inte = " & Ci% & " "
     SQLX$ = SQLX$ + "AND Stat_Ocom < 9 "
       If NPROVE% > 0 Then
         SQLX$ = SQLX$ + "AND Npro_Ocom = " & NPROVE% & " "
       End If
     SQLX$ = SQLX$ + "ORDER BY Femi_Ocom ASC, Nume_Ocom ASC"
     'Set OcomdetaTemp = Compras.OpenRecordset(SQLX$, dbOpenSnapshot)
     Dim OcomdetaTemp As ADODB.Recordset
     Set OcomdetaTemp = New ADODB.Recordset
25   On Error Resume Next
     OcomdetaTemp.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockReadOnly, adCmdText
     If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
      On Error GoTo 0
      Call CapturaErrorOpen
      GoTo 25
     End If
     On Error GoTo 0
     '
     On Error Resume Next
     With OcomdetaTemp
       .MoveLast
       If Err < 1 Then
         On Error Resume Next
         UNICO$ = SAFETEXT$(!Uni_Com)
         On Error GoTo 0
       End If
     End With
     '
     If ECOARCH$("UNIMED", UNICO$) = "" Then UNICO$ = ""
     If UNICO$ = "" Then UNICO$ = Unimed$(Ci%)
   End If
   '
   UNICOM$ = UNICO$
   '
End Function
'
Function ULTICOM$(NPRO%, CIIX%)
   '
   ' trae el registro de OCOMDETA correspondiente a la
   ' última compra no anulada del artículo y proveedor
   ' con precio registrado mayor que cero
   '
   Dim OcomdetaTemp  As ADODB.Recordset
   VUELTA% = 0
   '
20 UCO$ = REGBLAN$("OCOMDETA")
   Call REPLA(UCO$, MKI$(NPRO%), 7, 2)
   Call REPLA(UCO$, MKI$(CIIX%), 9, 2)
   '
   Call ABRECOMPRAS
   SQLX$ = "SELECT * FROM OCOMDETA "
   SQLX$ = SQLX$ + "WHERE Cod_Inte = " & CIIX% & " "
   SQLX$ = SQLX$ + "AND Npro_Ocom = " & NPRO% & " "
   If VUELTA% = 0 Then SQLX$ = SQLX$ + "AND Pre_Unit >= 0.00005 "
   SQLX$ = SQLX$ + "AND Stat_Ocom < 9 "
   SQLX$ = SQLX$ + "ORDER BY Femi_Ocom DESC, Nume_Ocom DESC"
   Set OcomdetaTemp = FLEXCONN.Execute(FILTSQL$(SQLX$))
   
   With OcomdetaTemp
     If Not (.BOF And .EOF) Then
         Call REPLA(UCO$, MKS$(!nume_ocom), 1, 4)
         Call REPLA(UCO$, MKI$(FECHANUM(Format(!Femi_Ocom, "dd/mm/yy"))), 5, 2)
         Call REPLA(UCO$, SAFETEXT$(!Uni_Mas), 11, 4)
         Call REPLA(UCO$, SAFETEXT$(!Uni_Com), 59, 4)
         Call REPLA(UCO$, MKS$(!Coef_Unids), 63, 4)
         Call REPLA(UCO$, Chr$(!Mone_Emis), 122, 1)
         Call REPLA(UCO$, MKD$(!Pre_Unit), 67, 8)
         Call REPLA(UCO$, MKS$(!Desc_Item), 75, 4)
       Else
         If VUELTA% < 1 Then
           VUELTA% = VUELTA% + 1
           GoTo 20
         End If
     End If
   End With
   '
   ULTICOM$ = UCO$
   '
End Function

Function ULTIPRE#(NPRO%, CIIX%)
  '
  'DEVUELVE EL ULTIMO PRECIO DE COMPRA DEL ARTICULO CIIX%
  'AL PROVEEDOR NPRO%
  '
  UC$ = ULTICOM$(NPRO%, CIIX%)
  MONULCO% = Asc(Mid$(UC$, 122, 1)): If MONULCO% < 1 Then MONULCO% = 1
  ULTIPRE# = CVD(Mid$(UC$, 67, 8)) * TICAMONE(MONULCO%)
  '
End Function

Function REVICOD$(Ci%)
   REVICOD$ = ""
   On Error Resume Next
   If STATUIT%(Ci%) <> 0 Then REVICOD$ = TRIM$(MasterTemp!Revision)
End Function
'
Function FEREVI%(Ci%)
   On Error Resume Next
   If STATUIT%(Ci%) <> 0 Then FEREVIST$ = Format(MasterTemp!Ferevis, "dd/mm/yy")
   FEREVI% = FECHANUM(FEREVIST$)
End Function
'
Function FEREFOR%(Ci%)   ' FECHA REVISION FORMULA
  '
  FRF% = 0
  On Error Resume Next
  If STATUIT%(Ci%) <> 0 Then
      FRF% = CVINT(MasterTemp!Revi_Formu)
      FRF2% = CVINT(MasterTemp!FechApFor)
      If FRF2% > 0 Then FRF% = FRF2%
  End If
  FEREFOR% = FRF%
  '
End Function
'
Function USREFOR$(Ci%)    ' USUARIO RESPONSABLE REVISION
  '
  URF$ = ""
  If Ci% > 0 Then
    XX$ = FILTERGETRECORD$("REVIFORM", 1, MKI$(Ci%))
    URF$ = TRIM$(Mid$(XX$, 5, 28))
  End If
  USREFOR$ = URF$
  '
End Function
'
' ***************************************************************************
'
Function CCOLOR%(TCX$)
    '
    TC$ = TRIM$(UCase$(TCX$))
    CC% = 0
    '
    If TC$ <> "" Then
       For U& = 1 To ULTREG&("TACOLOR")
           XX$ = TRIM$(Left$(REGLEIDO$("TACOLOR", U&), 31))
           If XX$ = TC$ Then
              CC% = U&
              GoTo 99
           End If
       Next U&
       '
       CC% = 1 + ULTREG&("TACOLOR")
       Call REGAPP("TACOLOR", TC$)
       Call CIERRARCH("TACOLOR")
       Call FRESHECHO("TACOLOR")
    End If
    '
99  CCOLOR% = CC%
    '
End Function
'

Function lColor$(CC%)
   '
   LCLR$ = ""
   If CC% > 0 Then
      If CC% <= ULTREG&("TACOLOR") Then
         U& = CC%
         LCLR$ = TRIM$(Left$(REGLEIDO$("TACOLOR", U&), 31))
      End If
   End If
   '
99 lColor$ = LCLR$
   '
End Function
'
' ***************************************************************************
'

Function CTALLE%(TCX$)
    '
    TC$ = TRIM$(UCase$(TCX$))
    CC% = 0
    '
    If TC$ <> "" Then
       For U& = 1 To ULTREG&("TATALLES")
           XX$ = TRIM$(Left$(REGLEIDO$("TATALLES", U&), 15))
           If XX$ = TC$ Then
              CC% = U&
              GoTo 99
           End If
       Next U&
       '
       If ULTREG&("TATALLES") >= 255 Then
            Call MENSERR(24, "Saturación de Tabla de Talles.\Imposible Agregar Talle '" + TC$ + "' a Tabla.")
            CC% = 0
          Else
            CC% = 1 + ULTREG&("TATALLES")
            Call REGAPP("TATALLES", TC$)
            Call CIERRARCH("TATALLES")
            Call FRESHECHO("TATALLES")
       End If
    End If
    '
99  CTALLE% = CC%
    '
End Function
'
Function LTALLE$(CC%)
   '
   LCLR$ = ""
   If CC% > 0 Then
      If CC% <= ULTREG&("TATALLES") Then
         U& = CC%
         LCLR$ = TRIM$(Left$(REGLEIDO$("TATALLES", U&), 15))
      End If
   End If
   '
99 LTALLE$ = LCLR$
   '
End Function
'
' ***************************************************************************
'
Function TIPOHIL$(Ci%)
'
TIPOHIL$ = TALLE$(Ci%)
End Function
'
' ***************************************************************************
'

Function TITULO$(Ci%)
'
If Ci% = CIANT% Then GoTo 10
'
TITULO$ = ""
If Ci% > 0 Then
   If Ci% <= NUMAS% Then
       REG& = Ci% + 1: CIANT% = Ci%
       XXX$ = REGLEIDO$("MASTER", REG&)
10     CP = CVS(Mid$(XXX$, 99, 4))
       If CP < 0 Then
          LCLR$ = TRIM$(CSTRING$(MKS$(Abs(CP)), 3, 8, 2, 2))
          TITULO$ = REPLACAR$(LCLR$, ".", "/")
       End If
   End If
End If
End Function
'
' ***************************************************************************
'

' ***************************************************************************
'

Function COSCALCU(Ci%, MODOCO%)
  '
  COU = 0
  RFG$ = RECFILT$("RECOSCAL", 1, MKI$(Ci%))
  If Len(RFG$) >= 4 Then
    REGB& = CVS(RFG$)
    If REGB& > 0 Then
      Q$ = REGLEIDO$("RECOSCAL", REGB&)
      COU = CVD(Mid$(Q$, 89, 8))
    End If
  End If
  '
  MODOCO% = 1
  If COU < 0.0005 Then
    COU = COSUNI(Ci%) * TICAMONE(MONECOSTO%(Ci%)) * (1 + PORNACIO(Ci%) / 100)
    MODOCO% = 2
  End If
  '
  COSCALCU = COU
  '
End Function

Function COSTOSUMI(Ci%, MODOCO%)
  '
  COU = 0
  RFG$ = RECFILT$("RECOSCAL", 1, MKI$(Ci%))
  If Len(RFG$) >= 4 Then
    REGB& = CVS(RFG$)
    If REGB& > 0 Then
      Q$ = REGLEIDO$("RECOSCAL", REGB&)
      COU = CVD(Mid$(Q$, 33, 8))        ' COSTO TOTAL DE SUMINISTROS
    End If
  End If
  '
  MODOCO% = 1
  If COU < 0.0005 Then
    COU = COSUNI(Ci%) * TICAMONE(MONECOSTO%(Ci%)) * (1 + PORNACIO(Ci%) / 100)
    MODOCO% = 2
  End If
  '
  COSTOSUMI = COU
  '
End Function
' ***************************************************************************
'
Function PRECALCU(Ci%)
  '
  COU = 0
  RFG$ = RECFILT$("LICOSCAL", 1, MKI$(Ci%))
  If Len(RFG$) >= 4 Then
    REGB& = CVS(RFG$)
    If REGB& > 0 Then
      Q$ = REGLEIDO$("LICOSCAL", REGB&)
      COU = CVD(Mid$(Q$, 97, 8))
    End If
  End If
  '
  PRECALCU = COU
  '
End Function
' ***************************************************************************
'
Function COSMATCAL(Ci%, MONECO%)
   '
   ' determina el costo calculado de suministros (M.Prima,
   ' Piezas y Trabajos de Terceros en la misma forma que
   ' el programa ANACOS01. Si MONECO%=1 en pesos
   '                       Si MONECO%=2 en dolares
   '
   Static DP%(32767), FESTRU#
   
   If Ci% < 1 Then GoTo 90
   CICON% = Ci%
   '
   NVI% = NUINV%
   NMM% = NUMAS%
   '
   Call EXPLOMUL(Ci%)
   SUMPPROP = 0: SUMPREC = 0
   SUPIPROP = 0: SUPIREC = 0
   SUSEPROP = 0: SUSEREC = 0
   '
   For KKI% = 1 To CONTAMUL%
     For kkj% = KKI% + 1 To CONTAMUL%
       If CIK%(kkj%) = CIK%(KKI%) Then
         CANT(KKI%) = CANT(KKI%) + CANT(kkj%)
         CANT(kkj%) = 0
       End If
     Next kkj%
   Next KKI%
   '
   ' costo de materias primas y piezas de terceros
   '
   SUPRI = 0: SUPIE = 0: SUSER = 0
   '
   FIN& = CONTAMUL
   For KKI% = 1 To CONTAMUL
     If Abs(CANT(KKI%)) < 0.0000000000005 Then GoTo 19
     If KKI% >= CONTAMUL% Or NIVE%(KKI% + 1) <= NIVE%(KKI%) Then  ' NO TIENE ESTRUCTURA
15      MONEI% = MONECOSTO%(CIK%(KKI%))
        DEMONE$ = DEMONECO$(MONEI%)
        COSUDOL = COSUNI(CIK%(KKI%))
        If MONEI% > 1 Then
            COSUN = COSUDOL * TICAMONE(MONEI%) * (1 + PORNACIO(CIK%(KKI%)) / 100)
          Else
            FEBAS% = FECOSTO%(CIK%(KKI%))
            FEACT% = HOY(HOS$)
            COSUN = COSUDOL * IAJUINF#(FEBAS%, FEACT%)
        End If
        '
        ' CONVERSION DE UNIDADES
        UNXX$ = Unimed$(CIK%(KKI%))
        USON = CANT(KKI%)
        Call CONVERUNID(USON, UNXX$)
        '
        Z$ = Space$(256)
        Call REPLA(Z$, CODEXT$(CIK%(KKI%)), 1, 15)
        DESCRI$ = DESCRIT0$(CIK%(KKI%))
        Call REPLA(Z$, DESCRI$, 17, 23)
        Call REPLA(Z$, DESCRI$, 129, 40)
        ' Call REPLA(Z$, CSTRING$(MKS$(CANT(KKI%)), 3, 10, 3, 2), 40, 10)
        Call REPLA(Z$, CSTRING$(MKS$(USON), 3, 10, 3, 2), 40, 10)
        ' Call REPLA(Z$, UNIMED$(CIK%(KKI%)), 52, 4)
        Call REPLA(Z$, UNXX$, 52, 4)
        Call REPLA(Z$, CSTRING$(MKS$(COSUDOL), 3, 10, DECICALCOS%, 2), 169, 10)
        Call REPLA(Z$, DEMONE$, 179, 10)
        Call REPLA(Z$, CSTRING$(MKS$(COSUN), 3, 10, DECICALCOS%, 2), 189, 10)
        UNIME$ = UCase$(Unimed$(CIK%(KKI%)))
        If Left$(UNIME$, 1) = "U" Or Left$(UNIME$, 2) = "PZ" Then
             If KKI% = 1 Then
                  SUPIPROP = SUPIPROP + CANT(KKI%) * COSUN
                Else
                  SUPIREC = SUPIREC + CANT(KKI%) * COSUN
             End If
             Call REPLA(Z$, CSTRING$(MKS$(CANT(KKI%) * COSUN), 3, 10, DECICALCOS%, 2), 65, 10)
             SUPIE = SUPIE + CANT(KKI%) * COSUN
          Else
             If KKI% = 1 Then
                  SUMPPROP = SUMPPROP + COSUN * CANT(KKI%)
                Else
                  SUMPREC = SUMPREC + COSUN * CANT(KKI%)
             End If
             Call REPLA(Z$, CSTRING$(MKS$(CANT(KKI%) * COSUN), 3, 10, DECICALCOS%, 2), 55, 10)
             SUPRI = SUPRI + CANT(KKI%) * COSUN
        End If
        Call REPLA(Z$, CSTRING$(MKS$(CANT(KKI%) * COSUN), 3, 10, DECICALCOS%, 2), 85, 10)
        '
      Else
        '
        If COSUNI(CIK%(KKI%)) > 0 Then
          PHB% = PROVHABI%(CIK%(KKI%))
          If PHB% > 0 Then
             If PROVINT%(PHB%) = 0 Then
               MONEI% = MONECOSTO%(CIK%(KKI%))
               DEMONE$ = DEMONECO$(MONEI%)
               COSUDOL = COSUNI(CIK%(KKI%))
               If MONEI% > 1 Then
                   COSUN = COSUDOL * TICAMONE(MONEI%) * (1 + PORNACIO(CIK%(KKI%)) / 100)
                 Else
                   FEBAS% = FECOSTO%(CIK%(KKI%))
                   FEACT% = HOY(HOS$)
                   COSUN = COSUDOL * IAJUINF#(FEBAS%, FEACT%)
               End If
               '
               Z$ = Space$(256)
               Call REPLA(Z$, CODEXT$(CIK%(KKI%)), 1, 15)
               DESCRI$ = DESCRIT0$(CIK%(KKI%))
               Call REPLA(Z$, DESCRI$, 17, 23)
               Call REPLA(Z$, DESCRI$, 129, 40)
               Call REPLA(Z$, CSTRING$(MKS$(CANT(KKI%)), 3, 10, 3, 2), 40, 10)
               Call REPLA(Z$, Unimed$(CIK%(KKI%)), 52, 4)
               Call REPLA(Z$, CSTRING$(MKS$(COSUDOL), 3, 10, DECICALCOS%, 2), 169, 10)
               Call REPLA(Z$, DEMONE$, 179, 10)
               Call REPLA(Z$, CSTRING$(MKS$(COSUN), 3, 10, DECICALCOS%, 2), 189, 10)
               If KKI% = 1 Then
                    SUSEPROP = SUSEPROP + COSUN * CANT(KKI%)
                  Else
                    SUSEREC = SUSEREC + CANT(KKI%) * COSUN
               End If
               Call REPLA(Z$, CSTRING$(MKS$(CANT(KKI%) * COSUN), 3, 10, DECICALCOS%, 2), 75, 10)
               SUSER = SUSER + CANT(KKI%) * COSUN
               Call REPLA(Z$, CSTRING$(MKS$(CANT(KKI%) * COSUN), 3, 10, DECICALCOS%, 2), 85, 10)
             End If
           End If
        End If
        '
     End If
19 Next KKI%

   A1 = TRIM$(CSTRING$(MKS$(SUMPPROP), 3, 12, DECICALCOS%, 2))
   A2 = TRIM$(CSTRING$(MKS$(SUMPREC), 3, 12, DECICALCOS%, 2))
   A3 = TRIM$(CSTRING$(MKS$(SUPIPROP), 3, 12, DECICALCOS%, 2))
   A4 = TRIM$(CSTRING$(MKS$(SUPIREC), 3, 12, DECICALCOS%, 2))
   A5 = TRIM$(CSTRING$(MKS$(SUSEPROP), 3, 12, DECICALCOS%, 2))
   A6 = TRIM$(CSTRING$(MKS$(SUSEREC), 3, 12, DECICALCOS%, 2))
   '
   SUMA1 = 0
   SUMA1 = SUMA1 + Val(A1)
   SUMA1 = SUMA1 + Val(A2)
   SUMA1 = SUMA1 + Val(A3)
   SUMA1 = SUMA1 + Val(A4)
   SUMA1 = SUMA1 + Val(A5)
   SUMA1 = SUMA1 + Val(A6)
   SUCOS = SUMA1

   '
'   Static DP%(32767)
'   Static FESTRU#, NVI%, NMM%
'   '
'   SUCOS = 0
'   CICON% = CI%
'   If CICON% < 1 Or CICON% > NUMAS% Then GoTo 90
'   '
'   ' IDENTIFICANDO LOS ITEMS SIN EXPLOSION
'   '
''   FESTRUC# = FileDateTime(LUDAT$ + "ESTRUNUE.DAT")
''   If FESTRUC# <> FESTRU# Then
''     '
''     NVI% = NUINV%
''     NMM% = NUMAS%
''     '
''     FESTRU# = FESTRUC#
''     FIN& = ULTREG&("ESTRUNUE")
''     For I& = 1 To FIN&
''       Call TRACE(20, I&, FIN&)
''       R0$ = REGLEIDO$("ESTRUNUE", I&)
''       CI1% = CVI(Left$(R0$, 2))
''       If CI1% > 0 Then
''         If CI1% <= NMM% Then
''           If CVS(Mid$(R0$, 5, 4)) > 0.00005 Then
''             DP%(CI1%) = 1
''           End If
''         End If
''       End If
''     Next I&
''   End If
'   '
'   Call EXPLOMUL(CICON%)
'   '
'   FIN& = CONTAMUL
'   For KKI% = 1 To CONTAMUL
'     COSUDOL = COSUNI(CIK%(KKI%))
'     If DP%(CIK%(KKI%)) < 1 Then   ' NO TIENE ESTRUCTURA
'        '
'        MONEI% = MONECOSTO%(CIK%(KKI%))
'        COSUN = COSUDOL * TICAMONE(MONEI%) * (1 + PORNACIO(CIK%(KKI%)) / 100)
'        SUCOS = SUCOS + CANT(KKI%) * COSUN
'        '
'      Else
'        '
'        If COSUDOL > 0 Then
'          PHB% = PROVHABI%(CIK%(KKI%))
'          If PHB% > 0 Then
'            If PROVINT%(PHB%) = 0 Then
'              MONEI% = MONECOSTO%(CIK%(KKI%))
'              COSUN = COSUDOL * TICAMONE(MONEI%) * (1 + PORNACIO(CIK%(KKI%)) / 100)
'              SUCOS = SUCOS + CANT(KKI%) * COSUN
'            End If
'          End If
'        End If
'        '
'     End If
'   Next KKI%
   '
   If MONECO% = 2 Then
      TCM = TICAMONE(2)
      If TCM < 1 Then TCM = 1
      SUCOS = SUCOS / TCM
   End If
   '
90 COSMATCAL = SUCOS
   '
End Function

'
' ***************************************************************************
'

' ***************************************************************************
'
Function PROVINT%(NPRX%)
   '
   Static PVIT%(32767), CPVI%
   '
   If CPVI% < 1 Then
     CPVI% = 1
     For UI& = 0 To 32767
       PVIT%(UI&) = 0
     Next UI&
     '
     For U& = 1 To ULTREG&("TAPROINT")
       X$ = REGLEIDO$("TAPROINT", U&)
       NPPX% = CVI(Left$(X$, 2))
       If NPPX% > 0 Then
         PVIT%(NPPX%) = 1
       End If
     Next U&
     Call CIERRARCH("TAPROINT")
   End If
   '
   PROVINT% = 0
   If NPRX% > 0 Then
     PROVINT% = PVIT%(NPRX%)
   End If
   '
End Function
'
' ***************************************************************************
'
Function PREBASE(Ci%)
'
PREBASE = COSUNI(Ci%)
End Function
'
' ***************************************************************************
'

' ***************************************************************************
'
Function PORCEIVA(CILOCAL%)
   '
   Static CTX%, URE&, PRIVA(15), REMA As String * 128
   If CTX% < 1 Then
     CTX% = 1
     URE& = ULTREG&("GRUPIVA")
     'If URE& > 4 Then URE& = 4
     For U& = 1 To URE&
       X$ = REGLEIDO$("GRUPIVA", U&)
       PRIVA(U& - 1) = CVS(Mid$(X$, 13, 4))
     Next U&
     Call CIERRARCH("GRUPIVA")
   End If
   '
   GRIVA% = GRUPIVA%(CILOCAL%)
   If GRIVA% < 0 Or GRIVA% > URE& - 1 Then GRIVA% = 0
   '
   'CODIGO AGREGADO PARA CONSIDRERAR IVAS COMBINADOS
   If VALIDA_TASAIVA$ = "SI" Then
    If CILOCAL% > 0 Then
      If XVALO(TASAIVAX_CI(CILOCAL%)) > 0 Then
       PORCEIVA = TASAIVAX_CI(CILOCAL%) 'VECTOR DE COMBINACION DE IVAS DE EQUIPOS CON ACCESORIOS
       Exit Function
      End If
    End If
   End If
   'FIN CODIGO AGREGADO PARA CONSIDRERAR IVAS COMBINADOS
   
   PORCEIVA = PRIVA(GRIVA%)
   '
End Function
'
' ***************************************************************************
'
Function STOMIN(Ci%)
   On Error Resume Next
   If STATUIT%(Ci%) <> 0 Then STOMIN = CDbl(MasterTemp!STOMIN)
   If STOMIN < 0 Then STOMIN = 100 * Abs(STOMIN)
End Function
'
' ***************************************************************************
'
' ***************************************************************************
'\\\OT-05-0720-WAR-25/11/05///
Function LINDIVIS(Ci%)
   On Error Resume Next
   If STATUIT%(Ci%) <> 0 Then LINDIVIS = CDbl(MasterTemp!LINDIVIS)
End Function

Function LINDIVEN(Ci%)
   On Error Resume Next
   If STATUIT%(Ci%) <> 0 Then LINDIVEN = CDbl(MasterTemp!LINDIVEN)
End Function

'\\\OT-05-0720-WAR-FIN///
'
' ***************************************************************************
'
Function LOREPOMA(Ci%)
   On Error Resume Next
   If STATUIT%(Ci%) <> 0 Then LOREPOMA = CDbl(MasterTemp!STOMAX)
End Function
'
' ***************************************************************************
'
Sub MUESTRADATCOD(CODDX$)
   '
   CICDX% = CODINT%(CODDX$)
   If CICDX% > 0 Then
     If CICDX% <= NUMAS% Then
       If EXISTE%(LUGRA$ + CODDX$ + ".JPG") = 0 Then
            REMASX& = CICDX% + 1
            RMCDX$ = REGBLAN$("MASTER") + String$(52, 0)
            Call REPLA(RMCDX$, MKI$(CICDX%), 1, 2)
            Call REPLA(RMCDX$, DESCRIT0$(CICDX%), 5, 64)
            Call REPLA(RMCDX$, CODEXT$(CICDX%), 77, 16)
            Call REPLA(RMCDX$, Chr$(TIMATE%(CICDX%)), 93, 1)
            Call REPLA(RMCDX$, Chr$(CRITREP%(CICDX%)), 94, 1)
            Call REPLA(RMCDX$, Unimed$(CICDX%), 109, 2)
            Call REPLA(RMCDX$, MKS$(COSUNI(CICDX%)), 111, 4)
            Call REPLA(RMCDX$, MKI$(FECOSTO%(CICDX%)), 115, 2)
            Call REPLA(RMCDX$, MKI$(PROVHABI%(CICDX%)), 123, 2)
            Call REPLA(RMCDX$, MKS$(STOCKACT(CICDX%)), 129, 4)
            Call REPLA(RMCDX$, MKS$(STOMIN(CICDX%)), 133, 4)
            Call REPLA(RMCDX$, MKS$(LOTESTAN(CICDX%)), 137, 4)
            Call REPLA(RMCDX$, MKS$(LOREPOMA(CICDX%)), 141, 4)
            Call REPLA(RMCDX$, TALLE$(CICDX%), 145, 16)
            Call REPLA(RMCDX$, COLOR$(CICDX%), 161, 16)
            Call REPLA(RMCDX$, MKS$(PESUNI(CICDX%)), 177, 4)
            '
            Call REPLA(RMCDX$, NUPLANO$(CICDX%), 181, 20) ' CODIGO ALIAS
            Call REPLA(RMCDX$, CODINTERNA$(CICDX%), 201, 48) ' CODIFICACION INTERNA
            '
            If DERACCE%("VERCOSTO", "Visualizar Costos y Precios") < 1 Then
              Call REPLA(RMCDX$, MKS$(99999999999#), 111, 4)
            End If
            OBXXX$ = OBJPLA$("MUESTRADATCOD", RMCDX$)
       Else
            '
            '\\\OT-1069-FL-24/11/08/// Carga los datos en el formulario FORDACOD07 en caso que el
            'articulo tenga asignado una imagen, sino carga el obejeto "MUESTRADATCOD"
            FORDACOD07.Label2.Caption = CODDX$
            FORDACOD07.Show 1
            '\\\OT-1069-FL-FIN///
       End If
     End If
   End If
End Sub

Sub MUESTRADATPRO(NPROV$)
   NPY& = XVALO(NPROV$)
   If NPY& > 0 Then
     If NPY& <= 32767 Then
       NPX% = NPY&
       If ECOARCH$("PROVED1", MKI$(NPX%)) <> "" Then
         RMCDX$ = String$(1024, 0)
         NC1% = (-1) * NPX%
         Call REPLA(RMCDX$, MKI$(NPX%), 1, 2)
         Call REPLA(RMCDX$, rasocli$(NC1%), 3, 62)
         Call REPLA(RMCDX$, TEL2CLI$(NC1%), 65, 32)
         Call REPLA(RMCDX$, DOMICLI$(NC1%), 97, 64)
         Call REPLA(RMCDX$, CPOSCLI$(NC1%), 161, 16)
         Call REPLA(RMCDX$, LOCACLI$(NC1%), 177, 48)
         Call REPLA(RMCDX$, TEL1CLI$(NC1%), 257, 24)
         Call REPLA(RMCDX$, PAISCLI$(NC1%), 225, 32)
         Call REPLA(RMCDX$, PROVCLI$(NC1%), 529, 1)
         Call REPLA(RMCDX$, FAXCLI$(NC1%), 281, 24)
         Call REPLA(RMCDX$, EMAILCLI$(NC1%), 305, 48)
         Call REPLA(RMCDX$, Chr$(PIVACLI%(NC1%)), 530, 1)
         Call REPLA(RMCDX$, CuitCli$(NC1%), 513, 16)
         Call REPLA(RMCDX$, Chr$(TICUEPRO%(NC1%)), 531, 1)
         Call REPLA(RMCDX$, Chr$(CPAGCLI%(NC1%)), 532, 1)
         OBXXX$ = OBJPLA$("MUESTRADATPRO", RMCDX$)
       End If
     End If
   End If
End Sub

Sub REPUNCODIG()
   '
Exit Sub
   Screen.MousePointer = 11
   NX% = FILEOPEN%(LUDAT$ + "PUNCODIG.DAT", 0, 2048)
   BFC$ = "": JI% = 0
   For U& = 1 To NUMAS%
      X$ = REGLEIDO$("MASTER", U& + 1)
      BFC$ = BFC$ + Mid$(X$, 77, 16)
      If Len(BFC$) >= 2048 Then
         BUFERCOD$ = BFC$
         JI% = JI% + 1
         Put #NX%, JI%, BUFERCOD$
         BFC$ = ""
      End If
   Next U&
   '
   If Len(BFC$) > 0 Then
      While Len(BFC$) < 2048
         BFC$ = BFC$ + Chr$(0)
      Wend
      BUFERCOD$ = BFC$
      JI% = JI% + 1
      Put #NX%, JI%, BUFERCOD$
   End If
   '
   For JI1% = JI% + 1 To LOF(NX%) / 2048
      BUFERCOD$ = String$(2048, 0)
      Put #NX%, JI1%, BUFERCOD$
   Next JI1%
   '
   Close #NX%: NX% = 0
   PUNCODIG$ = ""
   Screen.MousePointer = 1
   '
End Sub
'
Sub CARDEPOS()
   '
   If NUSU% < 1 Then
      SMP = Screen.MousePointer
      Screen.MousePointer = 11
      FIN& = ULTREG&("SUCURSAL")
      For U& = 1 To FIN&
         If U& <= 62 Then
            X$ = REGLEIDO$("SUCURSAL", U&)
            NUSU% = U&
            If Len(X$) >= 64 Then
               COSU%(U&) = Asc(Mid$(X$, 64, 1))
               SUCU$(U&) = Left$(X$, 20)
               ADEPO$(U&) = Mid$(X$, 21, 8)
               If TRIM$(ADEPO$(U&)) = "" Then ADEPO$(U&) = Left$(X$, 8)
               CATEO#(U&) = 0
               CAFIN#(U&) = 0
               DISTO%(U&) = 0
               If UCase$(Mid$(X$, 62, 2) = "NO") Then
               DISTO%(U&) = (-1)
               End If
            End If
         End If
         '
         If U& = 1 Then
            ' ACTUALIZA LOS REGISTROS DE MOVIMIENTOS EN DEPOSITO 0 (CERO)
            CONDI$ = "COD_INTE > 0 AND DEPOMOVI = 0"
            QTR& = CantRegistros("MOVISTO", CONDI$)
            If QTR& > 0 Then
               Call ACTUREGISTRO("MOVISTO", "DEPOMOVI", CONDI$, COSU%(1), RAFE&, "NOMESS")
            End If
         End If
      Next U&
      '
      If NUSU% < 1 Then
         NUSU% = 1
         COSU%(1) = 1
         SUCU$(1) = "Deposito Central"
         ADEPO$(1) = "Central"
         CATEO#(1) = 0
         CAFIN#(1) = 0
      End If
      Screen.MousePointer = SMP
   End If
   '
End Sub
'
Function CANTI_ENPROCESO_SECADO(CI1%, DIASECADO, Optional FECHA_VARIABLE%)
   '
   'LA FECHA VARIABLE ES POR SI SE PASA UNA FECHA POSTERIOR ENTONCES VA A CONSIDERAR SI ESTA SECO PARA DICHA FECHA
   'SI NO TOMA LA DEL DIA.
   If FECHA_VARIABLE% <= 0 Then
      FECHA_VARIABLE% = HOY(HO$)
   End If
   'TIEMPO DE SECADO  CONSIDERADO PARA PREMOTEC
   If DIASECADO < 1 Then
     DIASECADO = XVALO(Control("CONFABPE", "DIASSECA"))
     If DIASECADO < 1 Then DIASECADO = 15
   End If
   
   '
   CONDIX$ = "COD_INTE = " & CI1%
   CONDIX$ = CONDIX$ + "  AND CODIMOVI = 'OF'"
   CONDIX$ = CONDIX$ + " AND datediff(D, FECHMOV,'" & FETEXCOR1$(FECHA_VARIABLE%) & "') < " & DIASECADO
   
   CANTI_ENPROCESO_SECADO = XVALO(VALOBADA("MOVISTO", "sum(CANTINGRE)", CONDIX$, ""))
End Function


Function STOCKACT(Ci%)
   '
   Static MIRS As ADODB.Recordset
   
   Call CARDEPOS
   Call APERBASDAT("STOCKS")
   '
'''   STAC = 0
'''   STRSQL = "Select sum(CantIngre) AS SUINGRE, SUM(CantSalid) as SUSALI from movisto WITH(NOLOCK) where cod_inte = " & CI%
'''   '
'''   Set MIRS = New ADODB.Recordset
'''   Set MIRS = FLEXCONN.Execute(FILTSQL$(STRSQL))
'''   If Not (MIRS.EOF And MIRS.BOF) Then   ' IsNull(mirs!SALDO) Then
'''      STAC = MIRS!SUINGRE - MIRS!SUSALI
'''   End If
'''   MIRS.Close
   'Call CARDEPOS
    Static USA_SALOTEDP%
    If USA_SALOTEDP% = 0 Then
       USA_SALOTEDP% = -1
       'SI ESTA EL CONTROL Y EXISTE LA TABLA DE SALDOS DE LOTES
       If Control("CONSLOTE", "SALOTEDE") = "SI" And EXISTE_TABLA%("SALOTEDEP") > 0 Then
         USA_SALOTEDP% = 1
       End If
    End If

    Condi1$ = ""
    For UI% = 1 To NUSU%
     If DISTO%(UI%) < 0 Then
        Condi1$ = Condi1$ + " AND DepoMovi <> " & COSU%(UI%)
     End If
    Next UI%

    If USA_SALOTEDP% <= 0 Then
        '
        CONDI$ = "Cod_inte = " & Ci%
        If Condi1$ <> "" Then CONDI$ = CONDI$ & " AND (" & Mid$(Condi1$, 5) & ")"
        'CONDI$ = "cod_inte = " & CI%
        SUINGRE = XVALO(VALOBADA("MOVISTO", "sum(CantIngre)", CONDI$, "NOMESS"))
        SUSALI = XVALO(VALOBADA("MOVISTO", "sum(CantSalid)", CONDI$, "NOMESS"))
        STACT = SUINGRE - SUSALI
        'STOCKACT = XVALO(STACT) - STONODIS(CI%)  'No considerar depositos no disponibles
        STOCKACT = XVALO(STACT)
        Exit Function
    Else
        strSQL$ = "Select SUM(SALDO) AS ACUMUSALDO  FROM SALOTEDEP WITH(NOLOCK)"
        strSQL$ = strSQL$ + " WHERE Cod_Inte = " & Ci%
        strSQL$ = strSQL$ + " AND SALDO > 0 "
        If Condi1$ <> "" Then CONDI$ = CONDI$ & " AND (" & Mid$(Condi1$, 5) & ")"
          
        Set MIRS = FLEXCONN.Execute(FILTSQL$(strSQL))
        If Not (MIRS.EOF And MIRS.BOF) Then
           SL = ROUND(XVALO(MIRS!ACUMUSALDO), CADECI%)
        End If
        STOCKACT = SL
        Exit Function
    End If
    
   '
End Function
'
Function STONODIS(Ci%)
   '
   Static CTX%, CONDI$
   If CTX% < 1 Then
      Call CARDEPOS
      CONDI$ = ""
      For UI% = 1 To NUSU%
        If DISTO%(UI%) < 0 Then
           CONDI$ = CONDI$ + " OR DepoMovi = " & COSU%(UI%)
        End If
      Next UI%
   End If
   '
   strSQL$ = "Cod_inte = " & Ci% & " AND (" & Mid$(CONDI$, 4) & ")"
   SUINGRE = XVALO(VALOBADA("MOVISTO", "SUM(CantIngre)", strSQL$, "NOMESS"))
   SUSALI = XVALO(VALOBADA("MOVISTO", "SUM(CantSalid)", strSQL$, "NOMESS"))
   '
   STONODIS = SUINGRE - SUSALI ' STAC
   '
End Function
'
Function STODEPOS(Ci%, DEPO%)
   '
   STAC = 0
   If DEPO% < 1 Then
      STAC = STOCKACT(Ci%)
      GoTo 99
   End If
   '
   Call CARDEPOS
   Call APERBASDAT("STOCKS")
   '
   strSQL$ = "Cod_inte = " & Ci%
   If DEPO% > 1 Then
        strSQL$ = strSQL$ + " And DepoMovi= " & DEPO%
      Else
        strSQL$ = strSQL$ + " And DepoMovi <= " & DEPO%
   End If
   '
   SUINGRE = XVALO(VALOBADA("MOVISTO", "SUM(CantIngre)", strSQL$, "NOMESS"))
   SUSALI = XVALO(VALOBADA("MOVISTO", "SUM(CantSalid)", strSQL$, "NOMESS"))
   STAC = SUINGRE - SUSALI ' STAC
   '
99 STODEPOS = STAC
   '
End Function
'
Function STODEPOS_NUEVO(Ci%, DEPO%)
   '
   STAC = 0
   If DEPO% < 1 Then
      STAC = STOCKACT(Ci%)
      GoTo 99
   End If
   '
   Call CARDEPOS
   Call APERBASDAT("STOCKS")
   '
   strSQL$ = "Cod_inte = " & Ci%
   If DEPO% > 1 Then
        strSQL$ = strSQL$ + " And DepoMovi= " & DEPO%
      Else
        strSQL$ = strSQL$ + " And DepoMovi <= " & DEPO%
   End If
   '
   SQLX$ = "SELECT SUM(CANTINGRE) as SUINGRE, SUM(CANTSALID) AS SUSALI FROM MOVISTO WITH(NOLOCK) WHERE " & strSQL$
   'SUINGRE = XVALO(VALOBADA("MOVISTO", "SUM(CantIngre)", STRSQL$, "NOMESS"))
   'SUSALI = XVALO(VALOBADA("MOVISTO", "SUM(CantSalid)", STRSQL$, "NOMESS"))
   Dim Rst As ADODB.Recordset
   Set Rst = READSET(SQLX$)
   If Not (Rst.BOF And Rst.EOF) Then
        STAC = XVALO(Rst!SUINGRE) - XVALO(Rst!SUSALI) ' STAC
   End If
   Rst.Close
   Set Rst = Nothing
   '
99 STODEPOS_NUEVO = STAC
   '
End Function

Function STODEPFECH(Ci%, DEPO%, fech%, Optional LOTE$, Optional DEPO_NOT_IN$)
   'DEPO_NOT_IN$ DEBE SER UNA CADENA EJEMPLO: "1,4,7" DE ESTA MANERA NO CONSIDERARA ESTOS 3 DEPOSITOS
   ' DEVUELVE LA EXISTENCIA EN EL DEPÓSITO DEPO% A LA FECHA FECH% - 04/04/07 (OT-07-0141)
   ' SI DEPO_NOT_IN$ = 0 DEVUELVE EXISTENCIA TOTAL DEL ITEM A ESA FECHA
      If CI1% = 39 Then
        XX = 1
      End If

   If Ci% > 0 Then
      SCMA = Screen.MousePointer
      Screen.MousePointer = 11
      '
      CARRAS# = 0
      HASDAT = FETEXCOR1$(fech%)
      Call APERBASDAT("STOCKS")
      SQLX$ = " SELECT sum(CantIngre) AS SUINGRE, SUM(Cantsalid) as SUSALI FROM MOVISTO "
      SQLX$ = SQLX$ + "WHERE Cod_Inte = " & Ci% & " "
      SQLX$ = SQLX$ + "AND FechMov <= '" & HASDAT & "'"
      If DEPO% > 0 Then
         SQLX$ = SQLX$ + " AND DepoMovi = " & DEPO%
      End If
      '
      If TRIM$(DEPO_NOT_IN$) <> "" Then  'DEPOSITOS QUE NO DEBE CONSIDERAR
         SQLX$ = SQLX$ + " AND DepoMovi NOT IN (" & DEPO_NOT_IN$ & ")"
      End If
      '
      If TRIM$(LOTE$) <> "" Then
         SQLX$ = SQLX$ + " AND IDENLOTE = '" & LOTE$ & "'"
      End If
      Dim MoviTemp As ADODB.Recordset
      Set MoviTemp = FLEXCONN.Execute(FILTSQL$(SQLX$))
      With MoviTemp
          If Not (MoviTemp.EOF And MoviTemp.BOF) Then   ' IsNull(mirs!SALDO) Then
             CARRAS# = XVALO(!SUINGRE) - XVALO(!SUSALI) ' SE AGREGO EL XVALO PARA QUE NO GENERE ERROR EN LOS CASOS NULL
          End If
      End With
      Set MoviTemp = Nothing
      '
   End If
   '
   STODEPFECH = CARRAS#
   Screen.MousePointer = SCMA
   '
End Function

Function STODEPFECHENKG#(CI1%, DEPO%, fech%, Optional LOTE$, Optional DEPO_NOT_IN$)
   'DEPO_NOT_IN$ DEBE SER UNA CADENA EJEMPLO: "1,4,7" DE ESTA MANERA NO CONSIDERARA ESTOS 3 DEPOSITOS
      'STOCK A FECHA EN KG
      '

      KGINGRE# = 0
      HASDAT = FETEXCOR1$(fech%)
      STODEPFECHENKG# = 0
      If CI1% < 1 Then Exit Function
      
      'PRIMERO LOS KG INGRESADOS
      SQLX$ = " SELECT sum(PESOKG) AS SUINGRE FROM MOVISTO WITH(NOLOCK) "
      SQLX$ = SQLX$ + " WHERE Cod_Inte = " & CI1% & " "
      SQLX$ = SQLX$ + " AND FechMov <= '" & HASDAT & "'"
      SQLX$ = SQLX$ + " AND CANTINGRE > 0"
      If DEPO% > 0 Then
         SQLX$ = SQLX$ + " AND DepoMovi = " & DEPO%
      End If
      '
      If TRIM$(DEPO_NOT_IN$) <> "" Then  'DEPOSITOS QUE NO DEBE CONSIDERAR
         SQLX$ = SQLX$ + " AND DepoMovi NOT IN (" & DEPO_NOT_IN$ & ")"
      End If
      '
      If TRIM$(LOTE$) <> "" Then
         SQLX$ = SQLX$ + " AND IDENLOTE = '" & LOTE$ & "'"
      End If
      Set MoviTemp = READSET(SQLX$)
      With MoviTemp
          If Not (MoviTemp.EOF And MoviTemp.BOF) Then   ' IsNull(mirs!SALDO) Then
             KGINGRE# = XVALO(!SUINGRE)  ' SE AGREGO EL XVALO PARA QUE NO GENERE ERROR EN LOS CASOS NULL
          End If
      End With
      Set MoviTemp = Nothing

      'AHORA LOS KG SALIDOS EN MISMO RANGO
      SQLX$ = " SELECT sum(PESOKG) AS SUSALID FROM MOVISTO WITH(NOLOCK) "
      SQLX$ = SQLX$ + " WHERE Cod_Inte = " & CI1% & " "
      SQLX$ = SQLX$ + " AND FechMov <= '" & HASDAT & "'"
      SQLX$ = SQLX$ + " AND CANTSALID > 0"
      If DEPO% > 0 Then
         SQLX$ = SQLX$ + " AND DepoMovi = " & DEPO%
      End If
      '
      If TRIM$(DEPO_NOT_IN$) <> "" Then  'DEPOSITOS QUE NO DEBE CONSIDERAR
         SQLX$ = SQLX$ + " AND DepoMovi NOT IN (" & DEPO_NOT_IN$ & ")"
      End If
      
      If TRIM$(LOTE$) <> "" Then
         SQLX$ = SQLX$ + " AND IDENLOTE = '" & LOTE$ & "'"
      End If
      Set MoviTemp = READSET(SQLX$)
      With MoviTemp
          If Not (MoviTemp.EOF And MoviTemp.BOF) Then   ' IsNull(mirs!SALDO) Then
             KGSALIDA# = XVALO(!SUSALID)  ' SE AGREGO EL XVALO PARA QUE NO GENERE ERROR EN LOS CASOS NULL
          End If
      End With
      Set MoviTemp = Nothing
      STODEPFECHENKG# = KGINGRE# - KGSALIDA#
End Function

Function STODEPFECHENMETROS#(CI1%, DEPO%, fech%, Optional LOTE$, Optional DEPO_NOT_IN$)
   'DEPO_NOT_IN$ DEBE SER UNA CADENA EJEMPLO: "1,4,7" DE ESTA MANERA NO CONSIDERARA ESTOS 3 DEPOSITOS

      'STOCK A FECHA EN METROS
      '
      If CI1% = 614 Then
      A = 1
      End If
      METROSINGRE# = 0
      HASDAT = FETEXCOR1$(fech%)
      STODEPFECHENMETROS# = 0
      If CI1% < 1 Then Exit Function
      
      'PRIMERO LOS KG INGRESADOS
      SQLX$ = " SELECT sum(METROS) AS SUINGRE FROM MOVISTO WITH(NOLOCK) "
      SQLX$ = SQLX$ + " WHERE Cod_Inte = " & CI1% & " "
      SQLX$ = SQLX$ + " AND FechMov <= '" & HASDAT & "'"
      SQLX$ = SQLX$ + " AND CANTINGRE > 0"
      If DEPO% > 0 Then
         SQLX$ = SQLX$ + " AND DepoMovi = " & DEPO%
      End If
      '
      If TRIM$(DEPO_NOT_IN$) <> "" Then  'DEPOSITOS QUE NO DEBE CONSIDERAR
         SQLX$ = SQLX$ + " AND DepoMovi NOT IN (" & DEPO_NOT_IN$ & ")"
      End If

      If TRIM$(LOTE$) <> "" Then
         SQLX$ = SQLX$ + " AND IDENLOTE = '" & LOTE$ & "'"
      End If
      Set MoviTemp = READSET(SQLX$)
      With MoviTemp
          If Not (MoviTemp.EOF And MoviTemp.BOF) Then   ' IsNull(mirs!SALDO) Then
             METROSINGRE# = XVALO(!SUINGRE)  ' SE AGREGO EL XVALO PARA QUE NO GENERE ERROR EN LOS CASOS NULL
          End If
      End With
      Set MoviTemp = Nothing

      'AHORA LOS KG SALIDOS EN MISMO RANGO
      SQLX$ = " SELECT sum(METROS) AS SUSALID FROM MOVISTO WITH(NOLOCK) "
      SQLX$ = SQLX$ + " WHERE Cod_Inte = " & CI1% & " "
      SQLX$ = SQLX$ + " AND FechMov <= '" & HASDAT & "'"
      SQLX$ = SQLX$ + " AND CANTSALID > 0"
      If DEPO% > 0 Then
         SQLX$ = SQLX$ + " AND DepoMovi = " & DEPO%
      End If
      '
      If TRIM$(DEPO_NOT_IN$) <> "" Then  'DEPOSITOS QUE NO DEBE CONSIDERAR
         SQLX$ = SQLX$ + " AND DepoMovi NOT IN (" & DEPO_NOT_IN$ & ")"
      End If

      If TRIM$(LOTE$) <> "" Then
         SQLX$ = SQLX$ + " AND IDENLOTE = '" & LOTE$ & "'"
      End If
      Set MoviTemp = READSET(SQLX$)
      With MoviTemp
          If Not (MoviTemp.EOF And MoviTemp.BOF) Then   ' IsNull(mirs!SALDO) Then
             METROSALIDA# = XVALO(!SUSALID)  ' SE AGREGO EL XVALO PARA QUE NO GENERE ERROR EN LOS CASOS NULL
          End If
      End With
      Set MoviTemp = Nothing
      STODEPFECHENMETROS# = METROSINGRE# - METROSALIDA#
End Function


'
Function STOCKASI(Ci%)
    '
10  ASII = 0
    UREOD& = ULTREG&("MAVENDID")
    RGXX$ = RECFILT$("MAVENDID", 20, MKI$(Ci%))
    '
    For U& = 1 To Len(RGXX$) Step 4
      REG& = CVS(Mid$(RGXX$, U&, 4))
      If REG& > 0 Then
        If REG& <= UREOD& Then
          X$ = REGLEIDO$("MAVENDID", REG&)
          If Asc(Mid$(X$, 63, 1)) <= 1 Then  ' STATUS 1
            CASIG = CVS(Mid$(X$, 177, 4)) - CVS(Mid$(X$, 225, 4))
            If CASIG < 0 Then CASIG = 0
            ASII = ASII + CASIG
          End If
        End If
      End If
    Next U&
    '
    STOCKASI = ASII
    '
End Function
'
Function STOCKRES(Ci%, Optional LOTE$)
    '
10  ASII = 0
    '
35  Call ABRESERVA
    SQLX$ = " SELECT SUM(CantRes) AS SURESE, SUM(CantCons) AS SUCONS FROM RESERVA WITH(NOLOCK) "
    SQLX$ = SQLX$ + "WHERE Cod_Inte = " & Ci% & " "
    SQLX$ = SQLX$ + "AND CantCons < (CantRes - 0.005) "
    If TRIM$(LOTE$) <> "" Then
        SQLX$ = SQLX$ + " AND IDENLOTE='" & TRIM$(LOTE$) & "'"
    End If
    'SQLX$ = SQLX$ + "ORDER BY IdSubOr ASC "
    '
'    Dim ReserTemp As ADODB.Recordset
'    Set ReserTemp = New ADODB.Recordset
'    ReserTemp.CursorType = adOpenDynamic
'    ReserTemp.LockType = adLockReadOnly
'25  On Error Resume Next
'    ReserTemp.Open FILTSQL$(SQLX$), FLEXCONN, , , adcmtext
'    If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
'      On Error GoTo 0
'      Call CapturaErrorOpen
'      GoTo 25
'    End If
'    On Error GoTo 0
    Set ReserTemp = READSET(SQLX$)
    '
    With ReserTemp
          If Not (ReserTemp.EOF And ReserTemp.BOF) Then
            SARE = XVALO(!SURESE) - XVALO(!SUCONS)
          End If
          If SARE < 0 Then SARE = 0
          If TRIM$(UCase$(Unimed$(Ci%))) = "GK" Then
             SARE = SARE / 1000
          End If

    End With
    '
    STOCKRES = SARE
    '
End Function

Function STOCKRES_POR_OFABRICACION(NORFA$, CI1%)
    '
    Call ABRESERVA
    SQLX$ = " SELECT SUM(CantRes) AS SURESE, SUM(CantCons) AS SUCONS FROM RESERVA WITH(NOLOCK) "
    SQLX$ = SQLX$ + " WHERE IDSUBOR = '" & NORFA$ & "' "
    SQLX$ = SQLX$ + " AND Cod_Inte = " & CI1%
    SQLX$ = SQLX$ + "AND CantCons < (CantRes - 0.005) "
    Set ReserTemp = READSET(SQLX$)
    '
    With ReserTemp
          If Not (ReserTemp.EOF And ReserTemp.BOF) Then
            SARE = XVALO(!SURESE) - XVALO(!SUCONS)
          End If
          If SARE < 0 Then SARE = 0
          If TRIM$(UCase$(Unimed$(Ci%))) = "GK" Then
             SARE = SARE / 1000
          End If

    End With
    '
    STOCKRES_POR_OFABRICACION = SARE
    '
End Function

Function STOCK_RESERV_SIN_ARMADOR(CIXI%)

        'DEVUELVE EL STOCK RESERVADO CONSIDERANDO SOLO LAS ORDENES DE FABRICACION QUE 'NO' ESTEN ASIGNADAS A ARMADORES
        '
        Reservado = 0
        SQLX$ = "SELECT SUM(RESERVA.CANTRES)AS CANTRES  , SUM(RESERVA.CANTCONS) AS CANTCONS "
        SQLX$ = SQLX$ + " FROM RESERVA WITH(NOLOCK) INNER JOIN ORDEFABR WITH(NOLOCK)"
        SQLX$ = SQLX$ + " ON RESERVA.IDSUBOR = ORDEFABR.IDSUBOR"
        SQLX$ = SQLX$ + " Where (OrdeFabr.ARMADOR < 1 OR OrdeFabr.ARMADOR IS NULL)"
        SQLX$ = SQLX$ + " AND RESERVA.COD_INTE = " & CIXI%
        SQLX$ = SQLX$ + " AND ORDEFABR.STATUORF < 3  "
        SQLX$ = SQLX$ + " AND RESERVA.CantCons < (RESERVA.CantRes - 0.005)"
        SQLX$ = SQLX$ + " AND RESERVA.CodiEmpr = " & CodiEmp%
        SQLX$ = SQLX$ + " AND ORDEFABR.CodiEmpr = " & CodiEmp%
        SQLX$ = SQLX$ + " GROUP BY RESERVA.COD_INTE"
        '
        Dim RESTMP As ADODB.Recordset
        Set RESTMP = FLEXCONN.Execute(SQLX$) ' SIN FILTSQL POR QUE INNER JOIN NO CORRERSPONDE
        With RESTMP
            If Not (RESTMP.EOF And RESTMP.BOF) Then
              Reservado = XVALO(!CANTRES) - XVALO(!CANTCONS)
            End If
            If Reservado < 0 Then Reservado = 0
            If TRIM$(UCase$(Unimed$(CIXI%))) = "GK" Then
               Reservado = Reservado / 1000
            End If
        End With
        RESTMP.Close
        Set RESTMP = Nothing
        '
        STOCK_RESERV_SIN_ARMADOR = Reservado
        '
End Function

'
Function STORESOF(Ci%, NORFAI&)
    '
    NBRS$ = TRIM$(Str$(NORFAI&))
    While Len(NBRS$) < 6
      NBRS$ = "0" + NBRS$
    Wend
    NBRS$ = "OF-" + NBRS$ + "-%"
    '
10  ASII = 0
    '
35  Call APERBASDAT("RESERVA")
    SQLX$ = " SELECT SUM(CantRes) AS SURESE, SUM(CantCons) AS SUCONS  FROM RESERVA WITH(NOLOCK) "
    SQLX$ = SQLX$ + "WHERE Cod_Inte = " & Ci% & " "
    SQLX$ = SQLX$ + "AND IdSubOr LIKE '" & NBRS$ & "' "
    SQLX$ = SQLX$ + "AND CantCons < (CantRes - 0.005) "
    '
    'Dim ReserTemp As ADODB.Recordset
    'Set ReserTemp = FLEXCONN.Execute(FILTSQL$(SQLX$))
    
    Set ReserTemp = READSET(SQLX$)

    With ReserTemp
          If Not (ReserTemp.EOF And ReserTemp.BOF) Then
            ASII = XVALO(!SURESE) - XVALO(!SUCONS)
          End If
          If ASII < 0 Then ASII = 0
    End With
    If TRIM$(UCase$(Unimed$(Ci%))) = "GK" Then
      ASII = ASII / 1000
    End If
    Set ReserTemp = Nothing
    '
    STORESOF = ASII
    '
End Function
'
Function STOCKDIS(Ci%)
    '
    AA1 = STOCKACT(Ci%)
    aa2 = STOCKASI(Ci%)
    aa3 = STOCKRES(Ci%)
    ' AA4 = STONODIS(CI%)   ' SUPRIMIDO PORQUE YA SE CONSIDERA EN STOCKACT
    
    STOCKDIS = AA1 - aa2 - aa3  ' STOCKACT(CI%) - STOCKASI(CI%) - STOCKRES(CI%) - STONODIS(CI%)
    '
End Function
'

Function STOCKDISNEVEN(Ci%)
    '
    STOCKDISNEVEN = STODEPOS(Ci%, DEPOEXPE%) - STOCKASI(Ci%) - STOCKRES(Ci%) - PEDNOASI(Ci%)
    '
End Function
'

Function STOCKDISNE(Ci%)
    '
    STOCKDISNE = STOCKDIS(Ci%) - PEDNOASI(Ci%)
    '
End Function
'
Function STORFAB(Ci%)
   '
   STORF = 0
   SQLX$ = "SELECT SUM(CANPROY) AS SUCANP, SUM(CANAPRO)AS SUCANAPRO FROM ORDEFABR "
   SQLX$ = SQLX$ + " WHERE COD_INTE = " & Ci% & ""
   SQLX$ = SQLX$ + " AND CANAPRO < CANPROY  "
   SQLX$ = SQLX$ + "AND StatuOrf < 3 "   ' para que no este anulada/cumplida/cerrada
   Dim OFFFABRIC As ADODB.Recordset
   Set OFFFABRIC = FLEXCONN.Execute(FILTSQL$(SQLX$))

   With OFFFABRIC
    If Not (OFFFABRIC.EOF And OFFFABRIC.BOF) Then
      STORF = XVALO(!SUCANP) - XVALO(!SUCANAPRO)
    End If
    If STORF < 0 Then STORF = 0
   End With
   OFFFABRIC.Close
   STORFAB = STORF
   '
End Function

Function STOCOPEN(Ci%)
     '
'     CAPEN = 0
   STOCOPEN = 0
   SQLX$ = "SELECT SUM(CANT_OCOM) AS SUOCOM, SUM(CANT_REC) AS SURECI FROM OCOMDETA"
   SQLX$ = SQLX$ + " WHERE COD_INTE = " & Ci% & ""
   SQLX$ = SQLX$ + " AND CANT_REC < CANT_OCOM "
   SQLX$ = SQLX$ + " AND Stat_Ocom < 9 "
   SQLX$ = SQLX$ + " AND N_OrServ < 1 "
   Dim OCOMMMDETA As ADODB.Recordset
   Set OCOMMMDETA = FLEXCONN.Execute(FILTSQL$(SQLX$))

   With OCOMMMDETA
    If Not (OCOMMMDETA.EOF And OCOMMMDETA.BOF) Then
      STOCOPEN = XVALO(!SUOCOM) - XVALO(!SURECI)
    End If
    If STOCOPEN < 0 Then STOCOPEN = 0
   End With
   OCOMMMDETA.Close
   Set OCOMMMDETA = Nothing

'18   RFF$ = RECFILT$("OCOMDETA", 5, MKI$(CI%))
'     For UK& = 1 To Len(RFF$) Step 4
'        U& = CVS(Mid$(RFF$, UK&, 4))
'        X$ = REGLEIDO$("OCOMDETA", U&)
'        CIIX% = CVI(Mid$(X$, 9, 2))
'        STAT% = Asc(Mid$(X$, 124, 1))
'        If STAT% < 9 Then      ' no esta anulada
'          NUORSE& = CVS(Mid$(X$, 117, 4))
'          If NUORSE& < 1 Then
'            COEFCAN = CVS(Mid$(X$, 63, 4))
'            If COEFCAN < 0.005 Then COEFCAN = 1
'            CACOM = CVS(Mid$(X$, 55, 4)) * COEFCAN
'            CAREC = CVS(Mid$(X$, 99, 4)) * COEFCAN
'            CAPEN = CACOM - CAREC
'          End If
'        End If
'     Next UK&
'     '
'     If CAPEN < 0 Then CAPEN = 0
'     STOCOPEN = CAPEN

End Function

Function STOLOTES(Ci%)
    '
    STLOT = 0
    '
    If Ci% > 0 Then
      Call APERBASDAT("STOCKS")
      SQLX$ = " SELECT * FROM LOTINGRE "
      SQLX$ = SQLX$ + "WHERE Cod_Inte = " & Ci% & " "
      'Set LotIngTemp = Stocks.OpenRecordset(SQLX$, dbOpenDynaset)
      Dim LOTINGTEMP As ADODB.Recordset
      Set LOTINGTEMP = New ADODB.Recordset
26    On Error Resume Next
      LOTINGTEMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockReadOnly, adCmdText
      If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
           On Error GoTo 0
           Call CapturaErrorOpen
           GoTo 26
      End If
      On Error GoTo 0
      '
      With LOTINGTEMP
25      On Error Resume Next
        .MoveFirst
        If Err < 1 Then
          On Error GoTo 0
          Do While Not .EOF
            STLOT = STLOT + XVALO(!CanSaldo)
          .MoveNext
          Loop
        End If
        On Error GoTo 0
      End With
      LOTINGTEMP.Close
      Set LOTINGTEMP = Nothing
    End If
    '
    STOLOTES = STLOT
    '
End Function
'
Function UBISTO$(Ci%)
    '
20  UBX$ = ""
    RGXX$ = RECFILT$("UBISTOCK", 1, MKI$(Ci%))
    For U& = 1 To Len(RGXX$) Step 4
        REG& = CVS(Mid$(RGXX$, U&, 4))
        X$ = REGLEIDO$("UBISTOCK", REG&)
        If CVI(Left$(X$, 2)) <> Ci% Then
           X$ = REGLEIDO$("UBISTOCK", 1)
           Call GRAREG("UBISTOCK", X$, 1)
           Call CIERRARCH("UBISTOCK")
           GoTo 20
        End If
        '
        If TRIM$(Mid$(X$, 3, 24)) <> "" Then
           If UBX$ <> "" Then UBX$ = UBX$ + ", "
           UBX$ = UBX$ + TRIM$(Mid$(X$, 3, 24))
        End If
    Next U&
    '
    UBISTO$ = UBX$
    '
End Function

Function UBISTO1$(Ci%)
    '
20  SQLX$ = "SELECT DISTINCT UBICACION FROM UBICALOT WITH(NOLOCK) WHERE CODIGO=" & Ci% & " AND CANTIDAD > 0.05"
    Set Rst = READSET(SQLX$)
    '
    UB$ = ""
    With Rst
        Do While Not .EOF
            UB$ = UB$ + SAFETEXT$(!UBICACION) & "; "
            .MoveNext
        Loop
    End With
    '
    On Error Resume Next
    Rst.Close
    Set Rst = Nothing
    On Error GoTo 0
    '
    UBISTO1$ = UB$
    '
End Function

'
Function GRABAUBICACION(CIXI%, CAN#, UBICA$, LOTE$)
    '
    UBICATMP$ = UBICA$
    NN% = InStr(1, UBICATMP$, Chr$(124))
    If NN% > 0 Then
        SE$ = TRIM$(Mid$(UBICATMP$, 1, NN% - 1))
        UBICATMP$ = Mid$(UBICATMP$, NN% + 1)
    End If
    NN% = InStr(1, UBICATMP$, Chr$(124))
    If NN% > 0 Then
        es$ = TRIM$(Mid$(UBICATMP$, 1, NN% - 1))
        UBICATMP$ = Mid$(UBICATMP$, NN% + 1)
    End If
    NN% = InStr(1, UBICATMP$, Chr$(124))
    If NN% > 0 Then
        FA$ = TRIM$(Mid$(UBICATMP$, 1, NN% - 1))
        UBICATMP$ = Mid$(UBICATMP$, NN% + 1)
    End If
    NN% = InStr(1, UBICATMP$, Chr$(124))
    If NN% < 1 Then
        co$ = TRIM$(Mid$(UBICATMP$, 1))
    End If
    '
    If SE$ = "" Or es$ = "" Or FA$ = "" Or co$ = "" Then Exit Function
    '
100 ABUS$ = RANDSTRING$(-10)
    CONDI$ = "LOTE = '" & ABUS$ & "'"
    If CantRegistros("UBICALOT", CONDI$) > 0 Then GoTo 100
    '
    Call CreaRegistro("UBICALOT", "LOTE", ABUS$, "NOMESS")
    Call ACTUREGISTRO("UBICALOT", "SECTOR", CONDI$, SE$, REGAF&, "NOMESS")
    Call ACTUREGISTRO("UBICALOT", "ESTANTERIA", CONDI$, es$, REGAF&, "NOMESS")
    Call ACTUREGISTRO("UBICALOT", "FILA", CONDI$, FA$, REGAF&, "NOMESS")
    Call ACTUREGISTRO("UBICALOT", "COLUMNA", CONDI$, co$, REGAF&, "NOMESS")
    Call ACTUREGISTRO("UBICALOT", "CODIGO", CONDI$, CIXI%, REGAF&, "NOMESS")
    Call ACTUREGISTRO("UBICALOT", "CANTIDAD", CONDI$, CAN#, REGAF&, "NOMESS")
        UBICACION$ = SE$ & es$ & FA$ & co$
    Call ACTUREGISTRO("UBICALOT", "UBICACION", CONDI$, UBICACION$, REGAF&, "NOMESS")
    Call ACTUREGISTRO("UBICALOT", "LOTE", CONDI$, LOTE$, REGAF&, "NOMESS")
    '
End Function

Function UBISTONUE$(CIXI%, CAN#)
    '
    TX$ = ""
    If VOLUMENCOD#(CIXI%) > 0.00005 Then
        ' UBICACION LIBRE DENTRO DE LAS ASIGNADAS AL CODIGO
        TX$ = UBICACIONLIBRE$(CIXI%, CAN#)
        
        ' UBICACION OCUPADA DE LAS ASIGNADAS AL CODIGO, PERO CON ESPACIO SUFICIENTE
        If TX$ = "" Then TX$ = UBICACIONOCUPADA$(CIXI%, CAN#)
        
        ' UBICACION PERTENECENTE A LA FAMILIA CON ESPACIO SUFICIENTE. SE CONSIDERAN TAMBIEN A LOS
        ' PULMONES QUE LE FUERON ASIGNADO UNA FAMILIA
        If TX$ = "" Then TX$ = UBICACIONFAMILIA$(CIXI%, CAN#)
        
        ' PULMON CON ESPACIO DISPONIBLE PERTENECIENTE A LA ESTANTERIA DONDE ESTA
        ' ASIGNADO EL PRODUCTO
        If TX$ = "" Then TX$ = PULMONDEESTANTERIA$(CIXI%, CAN#)
        
        ' PULMON CON ESPACIO DISPONIBLE PERTENECIENTE AL SECTOR DONDE ESTA
        ' ASIGNADO EL PRODUCTO
        If TX$ = "" Then TX$ = PULMONDELSECTOR$(CIXI%, CAN#)
        
        ' PULMON CON ESPACIO DISPONIBLE PERTENECIENTE A UN SECTOR DONDE NO ESTA
        ' ASIGNADO EL PRODUCTO
        If TX$ = "" Then TX$ = PULMONOTROSECTOR$(CIXI%, CAN#)
        
        'If TX$ = "" Then TX$ = UBICACIONMANUL$(CIXI%, CAN#)
    End If
    UBISTONUE$ = TX$

End Function

Function UBICACIONLIBRE$(CIXI%, CAN#)
    '
    UBICACIONLIBRE$ = ""
    '
    ' VERIFICA SI EL CODIGO ESTA ASIGNADO A ALGUNA UBICACION
        SQLX$ = "SELECT * FROM UBICACOD WITH(NOLOCK) WHERE CODIGO=" & CIXI% & " ORDER BY SECTOR,ESTANTERIA,FILA,COLUMNA ASC"
        Set Rst = READSET(SQLX$)
        With Rst
            Do While Not .EOF
                ' BUSCA SI LA UBICACION (SECTOR+ESTANTE+FILA+COLUMNA) ESTA VACIO.
                ' ESTO SIGNIFICA QUE NO HAYAN REGISTROS EN UBICALOT
                SE$ = SAFETEXT$(!SECTOR)
                es$ = SAFETEXT$(!ESTANTERIA)
                fi$ = SAFETEXT$(!Fila)
                co$ = SAFETEXT$(!COLUMNA)
                CONDI$ = "SECTOR='" & SE$ & "'" & " AND ESTANTERIA='" & es$ & "'" & " AND FILA='" & fi$ & "'" & " AND COLUMNA='" & co$ & "'"
                REG0& = CantRegistros("UBICALOT", CONDI$)
                If REG0& < 1 Then
                    ' CONTROLA QUE EL VOLUMEN DE LAS CAJAS CABE EN LA UBICACION
                    LOTINDI# = LINDIVEN(CIXI%): If LOTINDI# < 0.00005 Then LOTINDI# = 0
                    M3CAJAS# = VOLUMENCOD#(CIXI%) * (CAN# / LOTINDI#)
                    M3UBICA# = XVALO(VALOBADA("UBICAVOL", "VOLUMEN", CONDI$)) ' VOLUMEN DE LA UBICACION
                    If M3UBICA# > M3CAJAS# Then
                        UBICACIONLIBRE$ = SE$ & Chr$(124) & es$ & Chr$(124) & fi$ & Chr$(124) & co$ ' ASIGNA UBICACION
                        Exit Do
                    End If
                End If
                .MoveNext
            Loop
        End With
        On Error Resume Next
        Rst.Close
        Set Rst = Nothing
        On Error GoTo 0
    '
End Function

Function UBICACIONOCUPADA$(CIXI%, CAN#)
    '
    UBICACIONOCUPADA$ = ""
    '
    ' VERIFICA SI EL CODIGO ESTA ASIGNADO A ALGUNA UBICACION
    SQLX$ = "SELECT * FROM UBICACOD WITH(NOLOCK) WHERE CODIGO=" & CIXI% & " ORDER BY SECTOR, ESTANTERIA, FILA, COLUMNA ASC"
    Set Rst = READSET(SQLX$)
    With Rst
        Do While Not .EOF
            M3UBICA# = 0
            ' CALCULA EL VOLUMEN TOTAL OCUPADO DE LA UBICACION
            SE$ = SAFETEXT$(!SECTOR)
            es$ = SAFETEXT$(!ESTANTERIA)
            fi$ = SAFETEXT$(!Fila)
            co$ = SAFETEXT$(!COLUMNA)
            CONDI$ = "SECTOR='" & SE$ & "'" & " AND ESTANTERIA='" & es$ & "'" & " AND FILA='" & fi$ & "'" & " AND COLUMNA='" & co$ & "'"
            VOLUMENOCUPADO# = XVALO(VALOBADA("UBICAVOL", "VOLUMEN", CONDI$))
            SQLX$ = "SELECT * FROM UBICALOT WITH(NOLOCK) WHERE " & CONDI$
            Set RSTLOTE = READSET(SQLX$)
            With RSTLOTE
                Do While Not .EOF
                    CIXX% = XVALO(!Codigo)
                    CANTI# = XVALO(!Cantidad)
                    LOTINDI# = LINDIVEN(CIXX%): If LOTINDI# < 0.00005 Then LOTINDI# = 0
                    VOLUMENLOTE# = VOLUMENCOD#(CIXX%) * (CANTI# / LOTINDI#)
                    '
                    M3UBICA# = M3UBICA# + VOLUMENLOTE#
                    .MoveNext
                Loop
            End With
            ' CALCULA VOLUMEN DEL PRODUCTO QUE SE QUIERE UBICAR
            LOTINDI# = LINDIVEN(CIXI%): If LOTINDI# < 0.00005 Then LOTINDI# = 0
            M3CAJAS# = VOLUMENCOD#(CIXI%) * (CAN# / LOTINDI#)
            ' CONTROLA QUE EL VOLUMEN DE LAS CAJAS CABE EN LA UBICACION
            If M3UBICA# + M3CAJAS# <= VOLUMENOCUPADO# Then
                UBICACIONOCUPADA$ = SE$ & Chr$(124) & es$ & Chr$(124) & fi$ & Chr$(124) & co$
                Exit Do
            End If
            .MoveNext
        Loop
    End With
    On Error Resume Next
    Rst.Close
    Set Rst = Nothing
    On Error GoTo 0
        '
End Function

Function UBICACIONFAMILIA$(CIXI%, CAN#)
    '
    UBICACIONFAMILIA$ = ""
    '
    ' OBTENGO LAS UBICACIONES CORRESPONDIENTE A LA FAMILIA DEL CODIGO
    FAMI$ = SAFETEXT$(VALOBADA("MASTER", "FAMILIA", "CODINT=" & CIXI%))
    SQLX$ = "SELECT * FROM UBICAVOL WITH(NOLOCK) WHERE FAMILIA='" & FAMI$ & "' ORDER BY SECTOR,ESTANTERIA,FILA,COLUMNA ASC"
    Set Rst = READSET(SQLX$)
    With Rst
        Do While Not .EOF
            M3UBICA# = 0
            ' CALCULA EL VOLUMEN TOTAL OCUPADO DE LA UBICACION
            SE$ = TRIM$(SAFETEXT$(!SECTOR))
            es$ = TRIM$(SAFETEXT$(!ESTANTERIA))
            fi$ = TRIM$(SAFETEXT$(!Fila))
            co$ = TRIM$(SAFETEXT$(!COLUMNA))
            CONDI$ = "SECTOR='" & SE$ & "'" & " AND ESTANTERIA='" & es$ & "'" & " AND FILA='" & fi$ & "'" & " AND COLUMNA='" & co$ & "'"
            VOLUMENOCUPADO# = XVALO(VALOBADA("UBICAVOL", "VOLUMEN", CONDI$))
            SQLX$ = "SELECT * FROM UBICALOT WITH(NOLOCK) WHERE " & CONDI$
            Set RSTLOTE = READSET(SQLX$)
            With RSTLOTE
                Do While Not .EOF
                    CIXX% = XVALO(!Codigo)
                    LOTINDI# = LINDIVEN(CIXX%): If LOTINDI# < 0.00005 Then LOTINDI# = 0
                    CANTI# = XVALO(!Cantidad)
                    VOLUMENLOTE# = VOLUMENCOD#(CIXX%) * (CANTI# / LOTINDI#)
                    '
                    M3UBICA# = M3UBICA# + VOLUMENLOTE#
                    .MoveNext
                Loop
            End With
            ' CALCULA VOLUMEN DEL PRODUCTO QUE SE QUIERE UBICAR
            LOTINDI# = LINDIVEN(CIXI%): If LOTINDI# < 0.00005 Then LOTINDI# = 0
            M3CAJAS# = VOLUMENCOD#(CIXI%) * (CAN# / LOTINDI#)
            ' CONTROLA QUE EL VOLUMEN DE LAS CAJAS CABE EN LA UBICACION
            If M3UBICA# + M3CAJAS# <= VOLUMENOCUPADO# Then
                UBICACIONFAMILIA$ = SE$ & Chr$(124) & es$ & Chr$(124) & fi$ & Chr$(124) & co$
                Exit Do
            End If
20          .MoveNext
        Loop
        On Error Resume Next
        RSTLOTE.Close
        Set RSTLOTE = Nothing
        On Error GoTo 0
    End With
    On Error Resume Next
    Rst.Close
    Set Rst = Nothing
    On Error GoTo 0

End Function

Function PULMONDEESTANTERIA$(CIXI%, CAN#)
    '
    PULMONDEESTANTERIA$ = ""
    '
    ' OBTENGO LAS UBICACIONES CORRESPONDIENTE AL CODIGO
    SQLX$ = "SELECT DISTINCT SECTOR, ESTANTERIA FROM UBICACOD WITH(NOLOCK) WHERE CODIGO=" & CIXI%
    Set RSTSEC = READSET(SQLX$)
    Do While Not RSTSEC.EOF
        ' OBTIENE LOS PULMONES CORRESPONDIENTES AL SECTOR Y ESTANTERIA QUE LE FUE ASIGNADA
        SE$ = SAFETEXT$(RSTSEC!SECTOR)
        es$ = SAFETEXT$(RSTSEC!ESTANTERIA)
        SQLX$ = "SELECT * FROM UBICAVOL WITH(NOLOCK) WHERE SECTOR='" & SE$ & "' AND ESTANTERIA='" & es$ & "' AND (FAMILIA='' OR FAMILIA IS NULL) ORDER BY SECTOR,ESTANTERIA,FILA,COLUMNA ASC"
        Set Rst = READSET(SQLX$)
        With Rst
            Do While Not .EOF
                M3UBICA# = 0
                ' BUSCA LA UBICACION ENTRE LOS PULMONES QUE CORRESPONDEN AL/LOS SECTOR/ES QUE PERTENECE EL PRODUCTO
                fi$ = TRIM$(SAFETEXT$(!Fila))
                co$ = TRIM$(SAFETEXT$(!COLUMNA))
                ' CALCULA EL VOLUMEN OCUPADO DE LAS UBICACIONES (PULMONES)
                CONDI$ = "SECTOR='" & SE$ & "'" & " AND ESTANTERIA='" & es$ & "'" & " AND FILA='" & fi$ & "'" & " AND COLUMNA='" & co$ & "'"
                VOLUMENOCUPADO# = XVALO(VALOBADA("UBICAVOL", "VOLUMEN", CONDI$))
                SQLX$ = "SELECT * FROM UBICALOT WITH(NOLOCK) WHERE " & CONDI$
                Set RSTLOTE = READSET(SQLX$)
                With RSTLOTE
                    Do While Not .EOF
                        CIXX% = XVALO(!Codigo)
                        CANTI# = XVALO(!Cantidad)
                        LOTINDI# = LINDIVEN(CIXX%): If LOTINDI# < 0.00005 Then LOTINDI# = 0
                        VOLUMENLOTE# = VOLUMENCOD#(CIXX%) * (CANTI# / LOTINDI#)
                        '
                        M3UBICA# = M3UBICA# + VOLUMENLOTE#
                        .MoveNext
                    Loop
                End With
                ' CALCULA VOLUMEN DEL PRODUCTO QUE SE QUIERE UBICAR
                LOTINDI# = LINDIVEN(CIXI%): If LOTINDI# < 0.00005 Then LOTINDI# = 0
                M3CAJAS# = VOLUMENCOD#(CIXI%) * (CAN# / LOTINDI#)
                ' CONTROLA QUE EL VOLUMEN DE LAS CAJAS CABE EN LA UBICACION
                If M3UBICA# + M3CAJAS# <= VOLUMENOCUPADO# Then
                    PULMONDEESTANTERIA$ = SE$ & Chr$(124) & es$ & Chr$(124) & fi$ & Chr$(124) & co$
                    Exit Do
                End If
20              .MoveNext
            Loop
        End With
        On Error Resume Next
        Rst.Close
        Set Rst = Nothing
        On Error GoTo 0
        RSTSEC.MoveNext
    Loop
    On Error Resume Next
    RSTSEC.Close
    Set Rst = Nothing
    On Error GoTo 0
    '
End Function

Function PULMONDELSECTOR$(CIXI%, CAN#)
    '
    PULMONDELSECTOR$ = ""
    '
    ' OBTENGO LAS UBICACIONES CORRESPONDIENTE A LA FAMILIA DEL CODIGO
    SQLX$ = "SELECT DISTINCT SECTOR FROM UBICACOD WITH(NOLOCK) WHERE CODIGO=" & CIXI%
    Set RSTSEC = READSET(SQLX$)
    Do While Not RSTSEC.EOF
        SECTOR$ = SAFETEXT$(RSTSEC!SECTOR)
        SQLX$ = "SELECT * FROM UBICAVOL WITH(NOLOCK) WHERE SECTOR='" & SECTOR$ & "' AND (FAMILIA='' OR FAMILIA IS NULL) ORDER BY SECTOR,ESTANTERIA,FILA,COLUMNA ASC"
        Set Rst = READSET(SQLX$)
        With Rst
            Do While Not .EOF
                M3UBICA# = 0
                ' BUSCA LA UBICACION ENTRE LOS PULMONES QUE CORRESPONDEN AL/LOS SECTOR/ES QUE PERTENECE EL PRODUCTO
                SE$ = TRIM$(SAFETEXT$(!SECTOR))
                es$ = TRIM$(SAFETEXT$(!ESTANTERIA))
                fi$ = TRIM$(SAFETEXT$(!Fila))
                co$ = TRIM$(SAFETEXT$(!COLUMNA))
                ' CALCULA EL VOLUMEN OCUPADO DE LAS UBICACIONES (PULMONES)
                CONDI$ = "SECTOR='" & SE$ & "'" & " AND ESTANTERIA='" & es$ & "'" & " AND FILA='" & fi$ & "'" & " AND COLUMNA='" & co$ & "'"
                VOLUMENOCUPADO# = XVALO(VALOBADA("UBICAVOL", "VOLUMEN", CONDI$))
                SQLX$ = "SELECT * FROM UBICALOT WITH(NOLOCK) WHERE " & CONDI$
                Set RSTLOTE = READSET(SQLX$)
                With RSTLOTE
                    Do While Not .EOF
                        CIXX% = XVALO(!Codigo)
                        CANTI# = XVALO(!Cantidad)
                        LOTINDI# = LINDIVEN(CIXX%): If LOTINDI# < 0.00005 Then LOTINDI# = 0
                        VOLUMENLOTE# = VOLUMENCOD#(CIXX%) * (CANTI# / LOTINDI#)
                        '
                        M3UBICA# = M3UBICA# + VOLUMENLOTE#
                        .MoveNext
                    Loop
                End With
                ' CALCULA VOLUMEN DEL PRODUCTO QUE SE QUIERE UBICAR
                LOTINDI# = LINDIVEN(CIXI%): If LOTINDI# < 0.00005 Then LOTINDI# = 0
                M3CAJAS# = VOLUMENCOD#(CIXI%) * (CAN# / LOTINDI#)
                ' CONTROLA QUE EL VOLUMEN DE LAS CAJAS CABE EN LA UBICACION
                If M3UBICA# + M3CAJAS# <= VOLUMENOCUPADO# Then
                    PULMONDELSECTOR$ = SE$ & Chr$(124) & es$ & Chr$(124) & fi$ & Chr$(124) & co$
                    Exit Do
                End If
20              .MoveNext
            Loop
        End With
        On Error Resume Next
        Rst.Close
        Set Rst = Nothing
        On Error GoTo 0
        RSTSEC.MoveNext
    Loop
    On Error Resume Next
    RSTSEC.Close
    Set Rst = Nothing
    On Error GoTo 0
    '
End Function

Function PULMONOTROSECTOR$(CIXI%, CAN#)
    '
    PULMONOTROSECTOR$ = ""
    '
    ' OBTENGO LOS SECTORES DONDE ESTA EL PRODUCTO PARA OBVIARLOS
    SQLX$ = "SELECT DISTINCT SECTOR FROM UBICACOD WITH(NOLOCK) WHERE CODIGO=" & CIXI%
    Set RSTSEC = READSET(SQLX$)
    i% = 0: CONDI$ = ""
    Do While Not RSTSEC.EOF
        '
        SECTOR$ = TRIM$(SAFETEXT$(RSTSEC!SECTOR))
        i% = i% + 1
        If i% = 1 Then
            CONDI$ = CONDI$ & " SECTOR<>'" & SECTOR$ & "'"
        Else
            CONDI$ = CONDI$ & " AND SECTOR<>'" & SECTOR$ & "'"
        End If
        '
        RSTSEC.MoveNext
    Loop
    '
    SQLX$ = "SELECT * FROM UBICAVOL WITH(NOLOCK) WHERE " & CONDI$ & " AND (FAMILIA='' OR FAMILIA IS NULL) ORDER BY SECTOR,ESTANTERIA,FILA,COLUMNA ASC"
    If CONDI$ = "" Then SQLX$ = "SELECT * FROM UBICAVOL WITH(NOLOCK) WHERE (FAMILIA='' OR FAMILIA IS NULL) ORDER BY SECTOR,ESTANTERIA,FILA,COLUMNA ASC"
    Set Rst = READSET(SQLX$)
    With Rst
        Do While Not .EOF
            M3UBICA# = 0
            ' BUSCA LA UBICACION ENTRE LOS PULMONES QUE CORRESPONDEN AL/LOS SECTOR/ES QUE PERTENECE EL PRODUCTO
            SE$ = TRIM$(SAFETEXT$(!SECTOR))
            es$ = TRIM$(SAFETEXT$(!ESTANTERIA))
            fi$ = TRIM$(SAFETEXT$(!Fila))
            co$ = TRIM$(SAFETEXT$(!COLUMNA))
            ' CALCULA EL VOLUMEN OCUPADO DE LAS UBICACIONES (PULMONES)
            CONDI$ = "SECTOR='" & SE$ & "'" & " AND ESTANTERIA='" & es$ & "'" & " AND FILA='" & fi$ & "'" & " AND COLUMNA='" & co$ & "' AND (FAMILIA='' OR FAMILIA IS NULL)"
            VOLUMENOCUPADO# = XVALO(VALOBADA("UBICAVOL", "VOLUMEN", CONDI$))
            SQLX$ = "SELECT * FROM UBICALOT WITH(NOLOCK) WHERE " & CONDI$
            Set RSTLOTE = READSET(SQLX$)
            With RSTLOTE
                Do While Not .EOF
                    CIXX% = XVALO(!Codigo)
                    CANTI# = XVALO(!Cantidad)
                    LOTINDI# = LINDIVEN(CIXX%): If LOTINDI# < 0.00005 Then LOTINDI# = 0
                    VOLUMENLOTE# = VOLUMENCOD#(CIXX%) * (CANTI# / LOTINDI#)
                    '
                    M3UBICA# = M3UBICA# + VOLUMENLOTE#
                    .MoveNext
                Loop
            End With
            ' CALCULA VOLUMEN DEL PRODUCTO QUE SE QUIERE UBICAR
            LOTINDI# = LINDIVEN(CIXX%): If LOTINDI# < 0.00005 Then LOTINDI# = 0
            M3CAJAS# = VOLUMENCOD#(CIXI%) * (CAN# / LOTINDI#)
            ' CONTROLA QUE EL VOLUMEN DE LAS CAJAS CABE EN LA UBICACION
            If M3UBICA# + M3CAJAS# <= VOLUMENOCUPADO# Then
                PULMONOTROSECTOR$ = SE$ & Chr$(124) & es$ & Chr$(124) & fi$ & Chr$(124) & co$
                Exit Do
            End If
20          .MoveNext
        Loop
        On Error Resume Next
        RSTLOTE.Close
        Set RSTLOTE = Nothing
        On Error GoTo 0
    End With
    On Error Resume Next
    Rst.Close
    Set Rst = Nothing
    On Error GoTo 0
    '
End Function

Function CAMBIAUBICACION$()
        '
'        Call SELECHO("MASTER/Maestro de Articulos")
'        COD$ = TRIM$(VALACT1$("MASTER"))
'        If COD$ <> "" Then
'            CIXI% = CODINT%(COD$)
'            LOTINDI# = LINDIVEN(CIXI%)
'            If LOTINDI# < 0.00005 Then
'                Call COMUNI("Imposible Continuar. Falta Lote Indivisible de Ventas")
'                Exit Function
'            End If
'            '
'            Call ARMALILOT(CIXI%, , , , 1)  ' PRESENTA SOLO LOTES CON SALDO
'            Call SELECHO("!LILOTSE")
'            '
'            LOTE$ = TRIM$(VALACT1$("!LILOTSE"))
'            LOTERESTO$ = TRIM$(VALACT2$("!LILOTSE"))
'            If LOTE$ = "" Then Exit Function
'            '
'            FAMI$ = SAFETEXT$(VALOBADA("MASTER", "FAMILIA", "CODINT=" & CIXI%))
'            SECTOR$ = TRIM$(SAFETEXT$(VALOBADA("UBICAVOL", "SECTOR", "FAMILIA='" & FAMI$ & "'")))
'            '
'            If SECTOR$ <> "" Then
'                Call CARGALISEL("UBIMANU", "UBICAVOL", "SECTOR/A4;ESTANTERIA/A4;FILA/A4;COLUMNA/A4;FAMILIA/A4", "SECTOR='" & SECTOR$ & "'")
'            Else
'                Call CARGALISEL("UBIMANU", "UBICAVOL", "SECTOR/A4;ESTANTERIA/A4;FILA/A4;COLUMNA/A4;FAMILIA/A4", "SECTOR<>''")
'            End If
'10          Call SELECHO("UBIMANU/Ubicaciones del Sector '" & SECTOR$ & "'")
'
'            SE$ = TRIM$(VALACT1$("UBIMANU"))
'            If SE$ = "" Then Exit Function
'
'            TXTX$ = VALACT2$("UBIMANU")
'            ES$ = TRIM$(Mid$(TXTX$, 1, 4))
'            FI$ = TRIM$(Mid$(TXTX$, 6, 4))
'            CO$ = TRIM$(Mid$(TXTX$, 11, 4))
'            '
'            CAN# = XVALO(VALOBADA("UBICALOT", "CANTIDAD", "LOTE='" & LOTE$ & "'"))
'            UBAC$ = TRIM$(SAFETEXT$(VALOBADA("UBICALOT", "UBICACION", "LOTE='" & LOTE$ & "'")))
'            '
'            CONDI$ = "SECTOR='" & SE$ & "'" & " AND ESTANTERIA='" & ES$ & "'" & " AND FILA='" & FI$ & "'" & " AND COLUMNA='" & CO$ & "'"
'            VOLUMENOCUPADO# = XVALO(VALOBADA("UBICAVOL", "VOLUMEN", CONDI$))
'            SQLX$ = "SELECT * FROM UBICALOT WITH(NOLOCK) WHERE " & CONDI$
'            Set RSTLOTE = READSET(SQLX$)
'            With RSTLOTE
'                Do While Not .EOF
'                    CIXX% = XVALO(!Codigo)
'                    CANTI# = XVALO(!CANTIDAD)
'                    LOTINDI# = LINDIVEN(CIXX%): If LOTINDI# < 0.00005 Then LOTINDI# = 1
'                    VOLUMENLOTE# = VOLUMENCOD#(CIXX%) * (CANTI# / LOTINDI#)
'                    '
'                    M3UBICA# = M3UBICA# + VOLUMENLOTE#
'                    .MoveNext
'                Loop
'            End With
'            ' CALCULA VOLUMEN DEL PRODUCTO QUE SE QUIERE UBICAR
'            LOTINDI# = LINDIVEN(CIXI%)
'            M3CAJAS# = VOLUMENCOD#(CIXI%) * (CAN# / LOTINDI#)
'            ' CONTROLA QUE EL VOLUMEN DE LAS CAJAS CABE EN LA UBICACION
'            If M3UBICA# + M3CAJAS# <= VOLUMENOCUPADO# Then
'                REUBICA$ = SE$ & Chr$(124) & ES$ & Chr$(124) & FI$ & Chr$(124) & CO$
'            End If
'            If REUBICA$ = "" Then
'                OPX% = COMALTER%("La Ubicación Elegida No Posee Espacio Suficiente.\\Elegir Otra Ubicacion\Cancelar")
'                If OPX% < 1 Or OPX% > 2 Then Exit Function
'                If OPX% = 1 Then GoTo 10
'            End If
'            OPX% = COMALTER%("¿Esta Seguro Cambiar la Ubicacion del Lote '" & LOTE$ & "'\De '" & UBAC$ & "' Por '" & REUBICA$ & "'?\\No, Conservar\Si, Cambiar")
'            If OPX% = 2 Then
'                Call BORRAREGISTROS("UBICALOT", "LOTE='" & LOTE$ & "'", REGAF&)
'                Call GRABAUBICACION(CIXI%, CAN#, REUBICA$, LOTE$)
'                Call COMUNI("Cambio de Ubicación Realizada")
'            End If
'            '
'        End If
        '
End Function

Function SEPRODUC(Ci%)
    '
20  SPROD = 0
    Call APERBASDAT("ORFAB")
    '
    SQLX$ = "SELECT NumeOrFa, SecciPro FROM ORDEFABR "
    SQLX$ = SQLX$ + "WHERE Cod_Inte = " & Ci% & " "
    SQLX$ = SQLX$ + "ORDER BY NumeOrFa DESC;"
    'Set ORFABTEMP = Ofabric.OpenRecordset(SQLX$, dbOpenSnapshot)
    Dim ORFABTEMP As ADODB.Recordset
    Set ORFABTEMP = New ADODB.Recordset
25  On Error Resume Next
    ORFABTEMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenDynamic, adLockPessimistic, adCmdText
    If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
      On Error GoTo 0
      Call CapturaErrorOpen
      GoTo 25
    End If
    On Error GoTo 0
    '
    On Error Resume Next
    With ORFABTEMP
      .MoveFirst
      If Err < 1 Then
        SPROD = XVALO(!SecciPro)
      End If
      On Error GoTo 0
    End With
    Set ORFABTEMP = Nothing
    '
    SEPRODUC = SPROD
    '
End Function
'
Function PEDNOASI(Ci%)
    '
10  ASII = 0
    RGXX$ = RECFILT$("ODESEMI", 6, MKI$(Ci%))
    '
    For i& = 1 To Len(RGXX$) Step 4
        REG& = CVS(Mid$(RGXX$, i&, 4))
        X$ = REGLEIDO$("ODESEMI", REG&)
        CIP% = CVI(Mid$(X$, 11, 2))
        If CIP% <> Ci% Then        ' fuerza reindexacion
            MPANTE = Screen.MousePointer
            Screen.MousePointer = 11
            X$ = REGLEIDO$("ODESEMI", 1)
            Call GRAREG("ODESEMI", X$, 1)
            Call CIERRARCH("ODESEMI")
            Screen.MousePointer = MPANTE
            GoTo 10
        End If
        '
        ASII = ASII + CVS(Mid$(X$, 17, 4))
        '
    Next i&
    '
    PEDNOASI = ASII
    '
End Function
'

Function PRELISTA#(APREC$, CIIXXI%)
   '
   Call APERBASDAT("PEDCLI")
   '
   If UCase$(Left$(APREC$, 5)) = "LIPRE" Then
       NLIS% = XVALO(Mid$(APREC$, 6, 3))
     Else
       NLIS% = XVALO(APREC$)
   End If

   PRELISTA# = 0: VUELTA% = 0
        If CIIXXI% > 0 Then
          If CIIXXI% <= NUMAS% Then
            SQLX$ = "SELECT * FROM Preventa WITH(NOLOCK)"
            SQLX$ = SQLX$ + "WHERE Cod_Lista =" & NLIS% & " "
            SQLX$ = SQLX$ + "and Cod_pieza =" & CIIXXI% & ""
'            Dim PRODUC As ADODB.Recordset
'            Set PRODUC = FLEXCONN.Execute(FILTSQL$(SQLX$))
            Set PRODUC = READSET(SQLX$)
            IncluirIva% = XVALO(VALOBADA("LISTAPRE", "INCLUIVA", "Cod_Lista =" & NLIS% & " "))

            With PRODUC
               On Error Resume Next
               .MoveFirst
               FIN& = .RecordCount
               If Err < 1 Then
                  If IncluirIva% = 1 Then ' SI INCLUYE IVA CALCULA EL PRECIO SIN IVA
                    PRELISTA# = (XVALO(!Precio_Unitario)) / (1 + PORCEIVA(CIIXXI%) / 100)
                  Else
                    PRELISTA# = XVALO(!Precio_Unitario)
                  End If
               End If
               On Error GoTo 0
            End With
            Set PRODUC = Nothing
          End If
        End If
        
   If CONVERPRECIO% = 1 Then
     'NLIS% = XVALO(Mid$(APREC$, 6, 3))
     PRELISTA# = PRELISTA# * TICAMONE(MONELISTA%(NLIS%)) / TICAMONE(MONEMI%)
   End If
   '
End Function

Function PRELISTACT#(CIX%, TIMONE%, Optional SOLOLISVIG%)
    'SOLOLISVIG% PARAMETRO OPCIONAL 0:LISTA VIGENTE SI ES 0 TRAE COSTO REPOSICION
    '                               1:SOLO DEVUELVE LISTA VIGENTE
    'TIMONE% , PARAMETRO QUE DEVUELVE CON EL TIPO DE MONEDAD DE LA LISTA
    'SE CONSIDERA QUE SOLO HAY UNA LISTA VIGENTE.
    'SI SE QUIERE QUE TRAIGA LA LISTA VIGENTE QUE CONTENGA ESE ARTICULO VOY A TENER
    'QUE AGREGAR EN PUNTO 2 LA CONDICION TAMBIEN DEL   'AND COD_PIEZA = " & CI%'
    '
    Ci% = CIX%
    '
'1.-TOMAR LA FECHA DEL DIA
    FECHAX = FETEXCOR1$(HOY(HO$))
    '----------------------------------------------------------------
    '
'2.-BUSCO LA LISTA VIGENTE
    'BUSCA QUE LA FECHA DE VIGENCIA SEA <= A LA FECHA DE HOY Y LO ORDENA DE MANERA DESCENDENTE
    'PARA QUE PRESENTE LA MAS CERCANA A HOY
    CONDI$ = "COD_LISTA > 0 AND VIGENCIA <= '" & FECHAX & "'"
    CONDI$ = CONDI$ + " ORDER BY VIGENCIA DESC"
    LISTA_ACTIVA% = XVALO(VALOBADA("LISTAPRE", "COD_LISTA", CONDI$, "NOMESS"))
    '------------------------------------------------------------------
    '
'3.-BUSCO EL CODIGO EN LA LISTA
    CONDI2$ = "COD_LISTA = " & LISTA_ACTIVA% & " AND COD_PIEZA = " & Ci%
    VALORLISTA# = XVALO(VALOBADA("PREVENTA", "PRECIO_UNITARIO", CONDI2$))
    '-------------------------------------------------------------------
    '
'4.- SI ENCUENTRA EL CODIGO EN LA LISTA Y EL PRECIO ES > 0 SALE
     If VALORLISTA# > 0 Or SOLOLISVIG% > 0 Then 'SI TIENE PRECIO EN LISTA O SI ES 0 PERO SOLO DEVUELVE LISTA VIGENTE
          PRELISTACT# = VALORLISTA#
          TIMONE% = MONELISTA%(LISTA_ACTIVA%)
          If TIMONE% < 1 Then TIMONE% = 1
          Exit Function
     End If
    '----------------------------------------------------------------
    '
'6.-EN CASO QUE SEA 0 DEVUELVE EL COSTO DE REPOSICION,
    PRELISTACT# = COSUNI(Ci%)
    TIMONE% = MONECOSTO%(Ci%)
    If TIMONE% < 1 Then TIMONE% = 1
    '
End Function
Function MONELISTA%(ILISTA%)
    Call ABREPEDCLI
    ML% = XVALO(VALOBADA("LISTAPRE", "Moneda", "Cod_Lista = " & ILISTA%))
    If ML% < 1 Then ML% = 1
    MONELISTA% = ML%
End Function

Function DECILISTA%(ILISTA%)
    Call ABREPEDCLI
    ML% = XVALO(VALOBADA("LISTAPRE", "Decimales", "Cod_Lista = " & ILISTA%))
    If ML% < 2 Then ML% = 2
    If ML% > 5 Then ML% = 5
    DECILISTA% = ML%
End Function


Function DENOLISTA$(ILISTA%)
    Call ABREPEDCLI
    ILIS% = ILISTA%: If ILIS% > 255 Then ILIS% = 0
    DNL$ = "Precios Concertados (Pesos)"
    If ILIS% > 0 Then
       DNL$ = TRIM$(SAFETEXT$(VALOBADA("LISTAPRE", "Denom_Lista", "Cod_Lista = " & ILISTA%)))
       MNL% = MONELISTA%(ILIS%)
       If MNL% = 2 Then
           DNL$ = DNL$ + " (Dólares)"
         Else
           DNL$ = DNL$ + " (Pesos)"
       End If
    End If
    DENOLISTA$ = DNL$
End Function
'
Function DEPOPRE%(Ci%)
'
DPP% = 0
TMT% = TIMATE%(Ci%)
For U& = 1 To ULTREG&("TABTIMA")
   X$ = REGLEIDO$("TABTIMA", U&)
   If Asc(Left$(X$, 1)) = TMT% Then
      DPP% = Asc(Mid$(X$, 17, 1))
      GoTo 10
   End If
Next U&
'
10 If DPP% > 0 Then
     If ECOARCH$("ECODEP", Chr$(DPP%)) = "" Then DPP% = 0
     If DPP% = 99 Then DPP% = 0
End If
'
DEPOPRE% = DPP%
'
End Function
'
Function DEPOEXPE%()
    '
    Static CTXX%, DPXXP%
    If CTXX% = 0 Then
      CTXX% = 1
      DPXXP% = 0
      DPXX% = XVALO(Control$("", "DEPOEXPE"))
      If DPXX% > 0 Then
        If DPXX% < 99 Then
          If ECOARCH$("TADEPOSI", Chr$(DPXX%)) <> "" Then
            DPXXP% = DPXX%
          End If
        End If
      End If
    End If
    '
    DEPOEXPE% = DPXXP%
    '
End Function
'
Function FEPLANO$(CIXI%)
   '
   NPLX$ = ""
   RNC$ = RECFILT$("INDIPLAN", 1, MKI$(CIXI%))
   If Len(RNC$) >= 4 Then
     REPLANO& = CVS(RNC$)
     If REPLANO& > 0 Then
       If REPLANO& <= ULTREG&("INDIPLAN") Then
         RPLA$ = REGLEIDO$("INDIPLAN", REPLANO&)
         FEX = CVI(Mid$(RPLA$, 23, 2))
         NPLX$ = FECHATEX$(FEX)
       End If
     End If
   End If
   '
   FEPLANO$ = NPLX$
   '
End Function

Function DEPENART%(COD1%) 'DEPOSITO POR OMISION PARA ENTRADA POR O/F
   '
   Call APERBASDAT("MASTER")
   If COD1% > 0 Then
       'Master.FindFirst ("CODINT = " & COD1%)
       'If Master.NoMatch = False Then
       SQLX$ = "SELECT DepoEntra FROM MASTER "
       SQLX$ = SQLX$ + " WHERE CODINT = " & COD1% & ""
       Dim MASTTTER As New ADODB.Recordset
       Set MASTTTER = FLEXCONN.Execute(FILTSQL$(SQLX$))
       With MASTTTER
         On Error Resume Next
         .MoveFirst
         If Err < 1 Then
          DEPENART% = XVALO(!DEPOENTRA)
         End If
         On Error GoTo 0
       End With
       Set MASTTTER = Nothing
   End If
   '
End Function
'
Function DEPCONART%(COD2%) 'DEPOSITO POR OMIDION PARA CONSUMO DE O/F
   '
   Call APERBASDAT("MASTER")
   If COD2% > 0 Then
       'Master.FindFirst ("CODINT = " & COD2%)
       'If Master.NoMatch = False Then
       SQLX$ = "SELECT DepoCon FROM MASTER "
       SQLX$ = SQLX$ + " WHERE CODINT = " & COD2% & ""
       Dim MASTTTER As New ADODB.Recordset
       Set MASTTTER = FLEXCONN.Execute(FILTSQL$(SQLX$))
       With MASTTTER
         On Error Resume Next
         .MoveFirst
         If Err < 1 Then
          DEPCONART% = XVALO(!DEPOCON)
         End If
         On Error GoTo 0
       End With
       Set MASTTTER = Nothing
   End If
  '
End Function

Function PZPLANO$(CIXI%)
   '
   NPLX$ = ""
   RNC$ = RECFILT$("INDIPLAN", 1, MKI$(CIXI%))
   If Len(RNC$) >= 4 Then
     REPLANO& = CVS(RNC$)
     If REPLANO& > 0 Then
       If REPLANO& <= ULTREG&("INDIPLAN") Then
         RPLA$ = REGLEIDO$("INDIPLAN", REPLANO&)
         NPLX$ = TRIM$(Mid$(RPLA$, 53, 12))
       End If
     End If
   End If
   '
   PZPLANO$ = NPLX$
   '
End Function

Sub PRIDOGRAF(CIXI%, IGRA%)
   '
   Static URE&, RUTARGRA$
   '
   If IGRA% < 0 Or IGRA% > 3 Then
      Exit Sub
   End If
   Screen.MousePointer = 11
   '
   If URE& = 0 Then
     URE& = ULTREG&("INDIPLAN")
     If URE& < 1 Then
       URE& = (-1)
     End If
     '
     RUTARGRA$ = TRIM$(Control$("", "RUTARGRA"))
     If RUTARGRA$ = "" Then RUTARGRA$ = TRIM$(LUDAT$)
     If Right$(RUTARGRA$, 1) <> "\" Then
       RUTARGRA$ = RUTARGRA$ + "\"
     End If
   End If
   '
   PRIFORGEN$ = Mid$(TRIM$(Control$("", "FORGRAFI")), 2)
   CAIMINI% = 0
   '
   If PRIFORGEN$ <> "" Then
     If CARDOGRAF%(CIXI%, IGRA%, AGRAFIC$) = 1 Then
       HIMAGEN = FORMALTER.Image3.Picture.Width
       VIMAGEN = FORMALTER.Image3.Picture.Height
       If HIMAGEN > VIMAGEN Then
            PRIFORM$ = "H" + PRIFORGEN$
          Else
            PRIFORM$ = "V" + PRIFORGEN$
       End If
       '
       If EXISTE%(LUDAT$ + PRIFORM$ + ".FRM") Then
         If CIXI% > 0 Then
           If CIXI% <= NUMAS% Then
             '
             CAPARDOC% = 0
             CACOLTAB1% = 0
             CAITAB1% = 0
             CACOLTAB2% = 0
             CAITAB2% = 0
             '
             CODPARAM$(1) = "DES-MAT": DOCPARAM$(1) = TRIM$(CODEXT$(CIXI%)) + " - " + DESCRIT0$(CIXI%)
             CAPARDOC% = 1
             '
             COPLANO$ = AGRAFIC$
             If COPLANO$ <> "" Then
               If EXISTE%(COPLANO$) > 0 Then
                 CAIMINI% = 1
                 CODIMAGEN$(1) = "GRAFIPIE"
                 ARCHIMAGEN$(1) = COPLANO$
               End If
             End If
           End If
         End If
       End If
       Call PRINTDOC("@" + PRIFORM$)
     End If
   End If
   Screen.MousePointer = 1
   '
End Sub
'
Function CARDOGRAF%(CIXI%, IGRA%, AGRAFIC$)
   '
   Static URE&, RUTARGRA$
   '
   If URE& = 0 Then
     URE& = ULTREG&("INDIPLAN")
     If URE& < 1 Then
       URE& = (-1)
     End If
     '
     RUTARGRA$ = TRIM$(Control$("", "RUTARGRA"))
     If RUTARGRA$ = "" Then RUTARGRA$ = TRIM$(LUDAT$)
     If Right$(RUTARGRA$, 1) <> "\" Then
       RUTARGRA$ = RUTARGRA$ + "\"
     End If
   End If
   '
   LPX% = 0: AGRAFIC$ = ""
   If IGRA% < 0 Or IGRA% > 3 Then
      Exit Function
   End If
   '
   ' BUSCAR ARCHIVO DEL PLANO DE CIIX%
   '
   APLANO$ = "": Index = 0
   If CIXI% > 0 Then
     If CIXI% <= NUMAS% Then
       Screen.MousePointer = 11
       RNC$ = RECFILT$("INDIPLAN", 1, MKI$(CIXI%))
       If Len(RNC$) >= 4 Then
         REPLANO& = CVS(RNC$)
         If REPLANO& > 0 Then
           If REPLANO& <= URE& Then
             RPLA$ = REGLEIDO$("INDIPLAN", REPLANO&)
             APLANO$ = TRIM$(Mid$(RPLA$, 65 + 48 * IGRA%, 48))
             If APLANO$ <> "" Then
               FORMALTER.Image3.Picture = LoadPicture("")
               COPLANO$ = RUTARGRA$ + APLANO$
               If COPLANO$ <> "" Then
                 If EXISTE%(COPLANO$) > 0 Then
                   FORMALTER.Image3.Picture = LoadPicture(COPLANO$)
                   LPX% = 1
                   AGRAFIC$ = COPLANO$
                 End If
               End If
             End If
           End If
         End If
       End If
     End If
   End If
   '
   CARDOGRAF% = LPX%
   '
End Function

Function DEGRUVEN$(cl%)
   DX$ = "Varios"
   If cl% > 0 Then DX$ = ECOARCH$("GRUCOVEN", MKI$(cl%))
   If DX$ = "" Then Beep: cl% = 0: DX$ = "Varios"
   DEGRUVEN$ = DX$
End Function
'
'
Function NUHOPROF$(CIXI%, NOPE%, Optional ALTEROP%)
   '
   NHPF$ = ""
   '
   ALTOP% = 0
   On Error Resume Next
   ALTOP% = ALTEROP%
   On Error GoTo 0
   '
   SQLX$ = "SELECT * FROM SecOper "
   SQLX$ = SQLX$ + "WHERE CodIConj=" & CIXI% & " "
   SQLX$ = SQLX$ + "AND NumeOper = " & NOPE% & " "
   If ALTOP% > 0 Then
        SQLX$ = SQLX$ + "AND AltOper = " & ALTOP% & " "
      Else
        'SQLX$ = SQLX$ + "AND (AltOper = " & ALTOP% & " OR ISNULL(AltOper) = true) "
        SQLX$ = SQLX$ + "AND (AltOper = " & ALTOP% & " OR AltOper IS NULL) "
       
   End If
   SQLX$ = SQLX$ + "AND StatOper = 1"
   'Set SecopTemp = Articulos.OpenRecordset(SQLX$, dbOpenSnapshot, dbReadOnly)
   Dim SECOPTEMP As ADODB.Recordset
   Set SECOPTEMP = FLEXCONN.Execute(FILTSQL$(SQLX$))
   CAIT% = 0
   On Error Resume Next
   SECOPTEMP.MoveLast
   If Err Then GoTo 90
   CAREC% = SECOPTEMP.RecordCount
   If CAREC% < 1 Then GoTo 90
   '
   SECOPTEMP.MoveFirst
   With SECOPTEMP
     NHPF$ = SAFETEXT$(!HoProFab)
   End With
   Set SECOPTEMP = Nothing
   '
90 NUHOPROF$ = NHPF$
   '
End Function
'
Function MAQUISTAN$(CIXI%, NOPE%, Optional ALTEROP%)
   '
   NPLX$ = ""
   '
   ALTOP% = 0
   On Error Resume Next
   ALTOP% = ALTEROP%
   On Error GoTo 0
   '
   SQLX$ = "SELECT * FROM SecOper "
   SQLX$ = SQLX$ + "WHERE CodIConj=" & CIXI% & " "
   SQLX$ = SQLX$ + "AND NumeOper = " & NOPE% & " "
   If ALTOP% > 0 Then
        SQLX$ = SQLX$ + "AND AltOper = " & ALTOP% & " "
      Else
        'SQLX$ = SQLX$ + "AND (AltOper = " & ALTOP% & " OR ISNULL(AltOper) = true) "
        SQLX$ = SQLX$ + "AND AltOper = " & ALTOP% & " OR AltOper IS NULL "
   End If
   SQLX$ = SQLX$ + "AND StatOper = 1"
   'Set SecopTemp = Articulos.OpenRecordset(SQLX$, dbOpenSnapshot, dbReadOnly)
   Dim SECOPTEMP As ADODB.Recordset
   Set SECOPTEMP = FLEXCONN.Execute(FILTSQL$(SQLX$))
   '
   CAIT% = 0
   On Error Resume Next
   SECOPTEMP.MoveLast
   If Err Then GoTo 90
   CAREC% = SECOPTEMP.RecordCount
   If CAREC% < 1 Then GoTo 90
   '
   SECOPTEMP.MoveFirst
   With SECOPTEMP
     NPLX$ = SAFETEXT$(!CodMaq0)
   End With
   Set SECOPTEMP = Nothing
   '
90 MAQUISTAN$ = NPLX$
   '
End Function
'
Function TIESTAN(CIXI%, NOPE%, CANTID#, MODO%, Optional ALTEROP%)
   '
   ' MODO% = 0 ==> incluye tiempo de preparación
   ' MODO% > 0 ==> no considera tiempo de preparación
   '
   Call ABRESECOP
   TSTA = 0
   '
   ALTOP% = 0
   On Error Resume Next
   ALTOP% = ALTEROP%
   On Error GoTo 0
   '
   SQLX$ = "SELECT * FROM SecOper "
   SQLX$ = SQLX$ + "WHERE CodIConj= " & CIXI% & " "
   SQLX$ = SQLX$ + "AND NumeOper = " & NOPE% & " "
   If ALTOP% > 0 Then
        SQLX$ = SQLX$ + "AND AltOper = " & ALTOP% & " "
      Else
        'SQLX$ = SQLX$ + "AND (AltOper = " & ALTOP% & " OR ISNULL(AltOper) = true) "
        SQLX$ = SQLX$ + "AND (AltOper = " & ALTOP% & " OR  AltOper IS NULL )"
   End If
   SQLX$ = SQLX$ + "AND StatOper = 1"
   'Set SecopTemp = Articulos.OpenRecordset(SQLX$, dbOpenSnapshot, dbReadOnly)
   Dim SECOPTEMP As ADODB.Recordset
   Set SECOPTEMP = FLEXCONN.Execute(FILTSQL$(SQLX$))

   
   CAIT% = 0
   'On Error Resume Next
   'SECOPTEMP.MoveLast
   'If Err Then GoTo 90
   'CAREC% = SECOPTEMP.RecordCount
   'If CAREC% < 1 Then GoTo 90
   '
   'SECOPTEMP.MoveFirst
   With SECOPTEMP
     If Not (.EOF And .BOF) Then
        CALOTES = 1
        CAOPS = XVALO(!CantiOps)
          If CAOPS < 0.01 Then CAOPS = 1
        CICLOHE = XVALO(!PieCiclo)
          If CICLOHE < 1 Then CICLOHE = 1
        MINPREP = 60 * XVALO(!HorsPrep) * CALOTES
          If MODO% > 0 Then MINPREP = 0
        MINFIJO = 60 * !HorsFijo * CALOTES
        MINVARI = CANTID# * !MinuStan
        MINHERR = CANTID# * 60 * !HorsMHer / CICLOHE
        MINDEMO = (MINFIJO + MINVARI + MINHERR) * !PorDemor / 100
        MINLIMP = 60 * !HorsLimp * CALOTES
        MINLOTE = MINPREP + MINFIJO + MINVARI + MINDEMO + MINLIMP
        '
        TSTA = CAOPS * MINLOTE
     End If
     '
   End With
   Set SECOPTEMP = Nothing
   

   

   '
90 TIESTAN = TSTA
   '
End Function
'
Function TIEMAQUI(CIXI%, NOPE%, CANTID#, MODO%, Optional ALTEROP%)
   '
   ' MODO% = 0 ==> incluye tiempo de preparación
   ' MODO% > 0 ==> no considera tiempo de preparación
   '
   Call ABRESECOP
   TSTA = 0
   
   ALTOP% = 0
   On Error Resume Next
   ALTOP% = ALTEROP%
   On Error GoTo 0
   '
   SQLX$ = "SELECT * FROM SecOper "
   SQLX$ = SQLX$ + "WHERE CodIConj=" & CIXI% & " "
   SQLX$ = SQLX$ + " AND NumeOper = " & NOPE% & " "
   If ALTOP% > 0 Then
        SQLX$ = SQLX$ + "AND AltOper = " & ALTOP% & " "
      Else
        'SQLX$ = SQLX$ + "AND (AltOper = " & ALTOP% & " OR ISNULL(AltOper) = true) "
        SQLX$ = SQLX$ + "AND (AltOper = " & ALTOP% & " OR AltOper is NULL)  "
   End If
   SQLX$ = SQLX$ + "AND StatOper = 1"
   'Set SecopTemp = Articulos.OpenRecordset(SQLX$, dbOpenSnapshot, dbReadOnly)
   Dim SECOPTEMP As ADODB.Recordset
   Set SECOPTEMP = FLEXCONN.Execute(FILTSQL$(SQLX$))
   '
   With SECOPTEMP
     CALOTES = 1
     CICLOHE = XVALO(!PieCiclo)
       If CICLOHE < 1 Then CICLOHE = 1
     MINPREP = 60 * XVALO(!HorsPrep) * CALOTES
       If MODO% > 0 Then MINPREP = 0
     MINFIJO = 60 * XVALO(!HorsFijo) * CALOTES
     MINVARI = CANTID# * XVALO(!MinuStan)
     MINHERR = CANTID# * 60 * XVALO(!HorsMHer) / CICLOHE
     MINDEMO = (MINFIJO + MINVARI + MINHERR) * XVALO(!PorDemor) / 100
     MINLIMP = 60 * XVALO(!HorsLimp) * CALOTES
     MINLOTE = MINPREP + MINFIJO + MINVARI + MINDEMO + MINLIMP
     '
     TSTA = MINLOTE
     '
   End With
   Set SECOPTEMP = Nothing
   '
90 TIEMAQUI = TSTA
   '
End Function
'\\\OT-07-0028-RG-25/01/07///
Function CADSTAN(CIXI%, NOPE%, Optional ALTEROP%)
   CADSTAN = 0
   TSTANBA = TIEMAQUI(CIXI%, NOPE%, 1, 1, ALTEROP%)
   If TSTANBA > 0 Then CADSTAN = 60 / TSTANBA
   If CADSTAN >= 1000 Then CADSTAN = Int(CADSTAN + 0.5)
End Function
'\\\OT-07-0028-RG-FIN///
Function CANOPS%(CIXI%)
   '
   CNP% = 0
   SQLX$ = "SELECT count(*) as CANOPE from SecOper"
   SQLX$ = SQLX$ + " WHERE CodIConj=" & CIXI% & ""
   SQLX$ = SQLX$ + " AND StatOper = 1"
   Dim SECOPTEMP As ADODB.Recordset
   Set SECOPTEMP = FLEXCONN.Execute(FILTSQL$(SQLX$))
   '
   On Error Resume Next
   CNP% = SECOPTEMP!CANOPE    'RecordCount
   On Error GoTo 0
   Set SECOPTEMP = Nothing
   '
90 CANOPS% = CNP%
   '
End Function
'
Function DESCRIOP$(CIXI%, NOPE%, Optional ALTEROP%)
   '
   Call APERBASDAT("MASTER")
   '
   ALTOP% = 0
   On Error Resume Next
   ALTOP% = ALTEROP%
   On Error GoTo 0
   '
   NPLX$ = ""
   '
   CONDI$ = "CodIConj=" & CIXI% & " "
   CONDI$ = CONDI$ + " AND NumeOper = " & NOPE% & " "
   If ALTOP% > 0 Then
        CONDI$ = CONDI$ + " AND AltOper = " & ALTOP% & " "
      Else
        CONDI$ = CONDI$ + " AND (AltOper = " & ALTOP% & " OR AltOper IS NULL ) "
   End If
   CONDI$ = CONDI$ + " AND StatOper = 1"
   '
   Descr$ = VALOBADA("SecOper", "DescOper", CONDI$, "NOMESS")
   CAIT% = 0

   If TRIM$(Descr$) <> "" Then
      NPLX$ = Descr$
   End If
   '
   DESCRIOP$ = NPLX$
   '
End Function
'
Function ULTIMOP%(CIXI%)
   '
   Call APERBASDAT("MASTER")
   ULTX1% = 0
   '
   SQLX$ = "SELECT * FROM SecOper "
   SQLX$ = SQLX$ + "WHERE CodIConj=" & CIXI% & " "
   SQLX$ = SQLX$ + "AND StatOper = 1 "
   SQLX$ = SQLX$ + "ORDER BY NumeOper DESC"
   'Set SecopTemp = Articulos.OpenRecordset(SQLX$, dbOpenSnapshot, dbReadOnly)
   Dim SECOPTEMP As ADODB.Recordset
   Set SECOPTEMP = FLEXCONN.Execute(FILTSQL$(SQLX$))

   On Error Resume Next
   SECOPTEMP.MoveFirst
   If Err < 1 Then
      ULTX1% = XVALO(SECOPTEMP!NUMEOPER)
   End If
   On Error GoTo 0
   Set SECOPTEMP = Nothing
   '
90 ULTIMOP% = ULTX1%
   '
End Function
'
Function CANOPREP(CIXI%, NOPE%)  ' CANTIDAD DE OPERARIOS DE PREPARACION PARA LA OPERACION
   '
   NPLY = 1
   'Set SecopTemp = Articulos.OpenRecordset("SELECT Nuoper_Prep FROM SecOper WHERE CodIConj=" & CIXI% & "AND NumeOper = " & NOPE% & " AND StatOper = 1", dbOpenSnapshot, dbReadOnly)
   SQLX$ = "SELECT Nuoper_Prep FROM SecOper "
   SQLX$ = SQLX$ + "WHERE CodIConj=" & CIXI% & ""
   SQLX$ = SQLX$ + "AND NumeOper = " & NOPE% & ""
   SQLX$ = SQLX$ + "AND StatOper = 1"
   Dim SECOPTEMP As ADODB.Recordset
   Set SECOPTEMP = FLEXCONN.Execute(FILTSQL$(SQLX$))

   CAIT% = 0
   On Error Resume Next
   SECOPTEMP.MoveLast
   If Err Then GoTo 90
   CAREC% = SECOPTEMP.RecordCount
   '
   If CAREC% > 0 Then
     SECOPTEMP.MoveFirst
     If SECOPTEMP!NUOPER_PREP > 0 Then
       NPLY = SECOPTEMP!NUOPER_PREP
     End If
   End If
   Set SECOPTEMP = Nothing
   '
90 CANOPREP = NPLY
   '
End Function
'
Function CATOPREP%(CIXI%, NOPE%)  ' CATEGORIA DE OPERARIOS DE PREPARACION PARA LA OPERACION
   '
   NPLY = 0
   'Set SecopTemp = Articulos.OpenRecordset("SELECT Catoper_Prep FROM SecOper WHERE CodIConj=" & CIXI% & "AND NumeOper = " & NOPE% & " AND StatOper = 1", dbOpenSnapshot, dbReadOnly)
   SQLX$ = "SELECT Catoper_Prep FROM SecOper "
   SQLX$ = SQLX$ + " WHERE CodIConj=" & CIXI% & ""
   SQLX$ = SQLX$ + " AND NumeOper = " & NOPE% & ""
   SQLX$ = SQLX$ + " AND StatOper = 1"
   Dim SECOPTEMP As ADODB.Recordset
   Set SECOPTEMP = FLEXCONN.Execute(FILTSQL$(SQLX$))

   CAIT% = 0
'   On Error Resume Next
'   SECOPTEMP.MoveLast
'   If Err Then GoTo 90
'   CAREC% = SECOPTEMP.RecordCount
'   '
'   If CAREC% > 0 Then
'     SECOPTEMP.MoveFirst
'     If SECOPTEMP!CATOPER_PREP > 0 Then
    If Not SECOPTEMP.EOF Then
       NPLY = XVALO(SECOPTEMP!CATOPER_PREP)
     End If
'   End If
   Set SECOPTEMP = Nothing
   '
90 CATOPREP% = NPLY
   '
End Function
'

Function OBSEROP$(CIXI%, NOPE%)
   '
   NPLX$ = ""
   SQLX$ = "SELECT ObseOper FROM SecOper "
   SQLX$ = SQLX$ + "WHERE CodIConj=" & CIXI% & ""
   SQLX$ = SQLX$ + "AND NumeOper = " & NOPE% & " "
   SQLX$ = SQLX$ + "AND StatOper = 1"
   Dim SECOPTEMP As ADODB.Recordset
   Set SECOPTEMP = FLEXCONN.Execute(FILTSQL$(SQLX$))

   On Error Resume Next
   NPLX$ = SECOPTEMP!OBSEOPER
   On Error GoTo 0
   Set SECOPTEMP = Nothing
   '
90 OBSEROP$ = NPLX$
   '
End Function

Function IAJUINF#(FEBAS%, FEACT%)
   '
   Static FIDATA#, AJUIN$
   Static CAPER%, DESPE%(512), HASPE%(512), IAJU#(512)
   '
   If AJUIN$ = "" Or FIDATIM#("FLEXCRL") <> FIDATA# Then
     FIDATA# = FIDATIM#("FLEXCRL")
     AJUIN$ = Control$("ANACOSTO", "AJUINFLA")
     If AJUIN$ <> "SI" Then AJUIN$ = "NO"
   End If
   '
   If AJUIN$ <> "SI" Then
     IAJC# = 1
     GoTo 99
   End If
   '
   'CARGA EL AJUSTE EN EL VECTOR PARA TODOS LOS MESES DE LA TABLA
   If CAPER% < 1 Then
     CAPER% = ULTREG&("INDIAJU")
     For U& = 1 To CAPER%
       XX$ = REGLEIDO$("INDIAJU", U&)
       DESPE%(U&) = CVI(Mid$(XX$, 17, 2))
       HASPE%(U&) = CVI(Mid$(XX$, 19, 2))
       IAJU#(U&) = CVD(Mid$(XX$, 21, 8))
     Next U&
     If CAPER% < 1 Then
       CAPER% = 1
       DESPE%(1) = (-32767)
       HASPE%(1) = 32767
       IAJU#(1) = 1
     End If
   End If
   '
   'LEE EL VALOR DEL MES DESDE
10 For KKI% = 1 To CAPER%
     If HASPE%(KKI%) >= FEBAS% Then
       IAJ1# = IAJU#(KKI%)
       GoTo 20
     End If
   Next KKI%
   IAJ1# = IAJU#(CAPER%)
   '
   'LEE EL VALOR DEL MES HASTA
20 For KKI% = 1 To CAPER%
     If HASPE%(KKI%) >= FEACT% Then
       IAJ2# = IAJU#(KKI%)
       '================================================================================================
       'SE AGREGA ESTAS 2 LINEAS PARA QUE CUANDO CORRE EL CALCULO SEGURAMENTE EL MES ACTUAL NO VA A TENER EL INDICE DE INFLACION.
       'ENTONCES COMPARA SI EL INDICE DEL MES ACTUAL ES < AL ANTERIOR SE QUEDA CON EL DEL MES ANTERIOR,
       'OCURRIA QUE EL MES ACTUAL TENIA 1 Y POR EJEMPLO EL ANTERIOR 1.8 PERO SE QUEDABA CON 1 Y NO ACTUALIZABA EL COSTO.
       IAJ3# = IAJU#(KKI% - 1)
       If IAJ2# < IAJ3# And IAJ2# = 1 Then IAJ2# = IAJ3# 'COMPARA QUE SEA = 1 POR QUE PUEDE SER MENOR EN CASO DE DEFLACION (EN ESE CASO VA A QUEDARSE CON EL ULTIMO)
       'FIN AGREGADO ====================================================================================
       GoTo 30
     End If
   Next KKI%
   IAJ2# = IAJU#(CAPER%)
   '
30 IAJC# = 1
   If IAJ1# >= 1 Then
     IAJC# = IAJ2# / IAJ1#
   End If
   '
99 IAJUINF# = IAJC#
   '
End Function

Function RVPLANO$(CIXI%)
   '
   NPLX$ = ""
   RNC$ = RECFILT$("INDIPLAN", 1, MKI$(CIXI%))
   If Len(RNC$) >= 4 Then
     REPLANO& = CVS(RNC$)
     If REPLANO& > 0 Then
       If REPLANO& <= ULTREG&("INDIPLAN") Then
         RPLA$ = REGLEIDO$("INDIPLAN", REPLANO&)
         NPLX$ = TRIM$(Mid$(RPLA$, 19, 4))
       End If
     End If
   End If
   '
   RVPLANO$ = NPLX$
   '
End Function

Function NUHOREI$(CIXI%)
   '
   NPLX$ = ""
   RNC$ = RECFILT$("DATADINS", 1, MKI$(CIXI%))
   URN& = ULTREG&("DATADINS")
   '
   If Len(RNC$) >= 4 Then
     RENOTA& = CVS(Left$(RNC$, 4))
     If RENOTA& > 0 Then
       If RENOTA& <= URN& Then
         TTXX$ = REGLEIDO$("DATADINS", RENOTA&)
         NPLX$ = Mid$(TTXX$, 5, 16)
       End If
     End If
   End If
   '
   NUHOREI$ = NPLX$
   '
End Function

Function ESPECIFI$(CIXI%)
   '
   NPLX$ = ""
   RNC$ = RECFILT$("ESPECIFI", 1, MKI$(CIXI%))
   URN& = ULTREG&("ESPECIFI")
   '
   For U& = 1 To Len(RNC$) Step 4
     RENOTA& = CVS(Mid$(RNC$, U&, 4))
     If RENOTA& > 0 Then
       If RENOTA& <= URN& Then
         TTYY$ = RTrim$(Mid$(REGLEIDO$("ESPECIFI", RENOTA&), 3))
         If NPLX$ <> "" Then
           NPLX$ = NPLX$ + Chr$(13) + Chr$(10)
         End If
         NPLX$ = NPLX$ + RTrim$(TTYY$)
       End If
     End If
   Next U&
   '
   ESPECIFI$ = NPLX$
   '
End Function

Function NUHOJES$(CIXI%)
   '
   NPLX$ = ""
   RNC$ = RECFILT$("INDIPLAN", 1, MKI$(CIXI%))
   If Len(RNC$) >= 4 Then
     REPLANO& = CVS(RNC$)
     If REPLANO& > 0 Then
       If REPLANO& <= ULTREG&("INDIPLAN") Then
         RPLA$ = REGLEIDO$("INDIPLAN", REPLANO&)
         NPLX$ = TRIM$(Mid$(RPLA$, 33, 16))
       End If
     End If
   End If
   '
   NUHOJES$ = NPLX$
   '
End Function

Function RVHOJES$(CIXI%)
   '
   NPLX$ = ""
   RNC$ = RECFILT$("INDIPLAN", 1, MKI$(CIXI%))
   If Len(RNC$) >= 4 Then
     REPLANO& = CVS(RNC$)
     If REPLANO& > 0 Then
       If REPLANO& <= ULTREG&("INDIPLAN") Then
         RPLA$ = REGLEIDO$("INDIPLAN", REPLANO&)
         NPLX$ = TRIM$(Mid$(RPLA$, 49, 4))
       End If
     End If
   End If
   '
   RVHOJES$ = NPLX$
   '
End Function

Function LEERPLANO%(CIXXI%, ARCHIPLANO$)
   '
   Static URE&, RUTARGRA$
   '
   If URE& = 0 Then
     URE& = ULTREG&("INDIPLAN")
     If URE& < 1 Then
       URE& = (-1)
     End If
     '
     RUTARGRA$ = TRIM$(Control$("", "RUTARGRA"))
     If RUTARGRA$ = "" Then RUTARGRA$ = TRIM$(LUDAT$)
     If Right$(RUTARGRA$, 1) <> "\" Then
       RUTARGRA$ = RUTARGRA$ + "\"
     End If
   End If
   '
   LPX% = 0: ARCHIPLANO$ = ""
   ' BUSCAR ARCHIVO DEL PLANO DE CIIX%
   APLANO$ = "": Index = 0
   If CIXXI% > 0 Then
     If CIXXI% <= NUMAS% Then
       Screen.MousePointer = 11
       RNC$ = RECFILT$("INDIPLAN", 1, MKI$(CIXXI%))
       If Len(RNC$) >= 4 Then
         REPLANO& = CVS(RNC$)
         If REPLANO& > 0 Then
           If REPLANO& <= URE& Then
             RPLA$ = REGLEIDO$("INDIPLAN", REPLANO&)
             APLANO$ = TRIM$(Mid$(RPLA$, 65, 48))
             If APLANO$ <> "" Then
               FORMALTER.Image3.Picture = LoadPicture("")
               COPLANO$ = RUTARGRA$ + APLANO$
               If COPLANO$ <> "" Then
                 If EXISTE%(COPLANO$) > 0 Then
                   FORMALTER.Image3.Picture = LoadPicture(COPLANO$)
                   LPX% = 1
                   ARCHIPLANO$ = COPLANO$
                 End If
               End If
             End If
           End If
         End If
       End If
     End If
   End If
   '
   LEERPLANO% = LPX%
   '
End Function

Function NUPLIEN$(CIXI%)
   '
   NPLX$ = ""
   RNC$ = RECFILT$("DATADINS", 1, MKI$(CIXI%))
   URN& = ULTREG&("DATADINS")
   '
   If Len(RNC$) >= 4 Then
     RENOTA& = CVS(Left$(RNC$, 4))
     If RENOTA& > 0 Then
       If RENOTA& <= URN& Then
         TTXX$ = REGLEIDO$("DATADINS", RENOTA&)
         NPLX$ = Mid$(TTXX$, 25, 16)
       End If
     End If
   End If
   '
   NUPLIEN$ = NPLX$
   '
End Function

Function RVPLIEN$(CIXI%)
   '
   NPLX$ = ""
   RNC$ = RECFILT$("DATADINS", 1, MKI$(CIXI%))
   URN& = ULTREG&("DATADINS")
   '
   If Len(RNC$) >= 4 Then
     RENOTA& = CVS(Left$(RNC$, 4))
     If RENOTA& > 0 Then
       If RENOTA& <= URN& Then
         TTXX$ = REGLEIDO$("DATADINS", RENOTA&)
         NPLX$ = Mid$(TTXX$, 41, 4)
       End If
     End If
   End If
   '
   RVPLIEN$ = NPLX$
   '
End Function

Function RVHOREI$(CIXI%)
   '
   NPLX$ = ""
   RNC$ = RECFILT$("DATADINS", 1, MKI$(CIXI%))
   URN& = ULTREG&("DATADINS")
   '
   If Len(RNC$) >= 4 Then
     RENOTA& = CVS(Left$(RNC$, 4))
     If RENOTA& > 0 Then
       If RENOTA& <= URN& Then
         TTXX$ = REGLEIDO$("DATADINS", RENOTA&)
         NPLX$ = Mid$(TTXX$, 21, 4)
       End If
     End If
   End If
   '
   RVHOREI$ = NPLX$
   '
End Function
'\\\OT-06-0105-WAR-23/03/06///
Function DESCOMPO%(Ci%)
   '
   DESCOMPO% = 0
   On Error Resume Next
   If STATUIT%(Ci%) <> 0 Then DESCOMPO% = MasterTemp!DesCompAut
   '
 '\\\OT-06-0105-WAR-FIN//
End Function
'
Function CODORIG$(CIXXX%, NCXXX)
   '
   ASELE$ = "REPLACOC"
   If NCXXX < 0 Then ASELE$ = "REPLACOP"
   '
   CDXXX$ = CODEXT$(CIXXX%)
   UREP& = ULTREG&(ASELE$)
   RFF$ = RECFILT$(ASELE$, 3, MKI$(CIXXX%))
   '
   For KKI% = 1 To Len(RFF$) Step 4
     REREP& = CVS(Mid$(RFF$, KKI%, 4))
     If REREP& > 0 Then
       If REREP& <= UREP& Then
         XXYY$ = REGLEIDO$(ASELE$, REREP&)
         If CVI(Mid$(XXYY$, 1, 2)) = Abs(NCXXX) Then
           CDYYY$ = TRIM$(Mid$(XXYY$, 5, 28))
           If CDYYY$ <> "" Then
             CDXXX$ = CDYYY$
             GoTo 99
           End If
         End If
       End If
     End If
   Next KKI%
   '
99 CODORIG$ = CDXXX$
   '
End Function

Function LOTESEL$(CIXI%)
   '
   Dim TLOTE$(256), CALOTE#(256)
   LSX$ = ""
   '
   If CIXI% > 0 Then
     CANLOT% = 0
     RFF$ = RECFILT$("MOVISTO", 2, MKI$(CIXI%))
     For U& = 1 To Len(RFF$) Step 4
       REMO& = CVS(Mid$(RFF$, U&, 4))
       X$ = REGLEIDO$("MOVISTO", REMO&)
       ENTRA# = CVD(Mid$(X$, 88, 8))
       If ENTRA# >= 0.05 Then
         IDENLO$ = TRIM$(Mid$(X$, 5, 16))
         If IDENLO$ = "" Then IDENLO$ = TRIM$(Mid$(X$, 23, 10))
         If IDENLO$ = "" Then IDENLO$ = TRIM$(Mid$(X$, 33, 10))
         IDENLO$ = AJUSTI$(IDENLO$, 16)
         '
         For KKI% = 1 To CANLOT%
           If Left$(TLOTE$(KKI%), 16) = IDENLO$ Then
             CALOTE#(KKI%) = CALOTE#(KKI%) + ENTRA#
             TTXX$ = TLOTE$(KKI%)
             Call REPLA(TTXX$, CSTRING$(MKD$(CALOTE#(KKI%)), 4, 12, 1, 2), 27, 12)
             TLOTE$(KKI%) = TTXX$
             GoTo 19
           End If
         Next KKI%
         FEX = CVI(Left$(X$, 2))
         CANLOT% = CANLOT% + 1
         TTXX$ = AJUSTI$(IDENLO$, 40)
         Call REPLA(TTXX$, FECHATEX$(FEX), 19, 8)
         Call REPLA(TTXX$, CSTRING$(MKD$(ENTRA#), 4, 12, 1, 2), 27, 12)
         TLOTE$(CANLOT%) = TTXX$
         CALOTE#(KKI%) = ENTRA#
         '
       End If
19   Next U&
     '
     For U& = 1 To Len(RFF$) Step 4
       REMO& = CVS(Mid$(RFF$, U&, 4))
       X$ = REGLEIDO$("MOVISTO", REMO&)
       SALID# = CVD(Mid$(X$, 96, 8))
       If SALID# >= 0.05 Then
         IDENLO$ = AJUSTI$(Mid$(X$, 5, 16), 16)
         IDENLOX$ = TRIM$(IDENLO$)
         '
35       For KKI% = 1 To CANLOT%
           If Left$(TLOTE$(KKI%), 16) = IDENLO$ Or IDENLOX$ = "" Then
             If SALID# <= CALOTE#(KKI%) Then
                 CALOTE#(KKI%) = CALOTE#(KKI%) - SALID#
                 SALID# = 0
               Else
                 SALID# = SALID# - CALOTE#(KKI%)
                 CALOTE#(KKI%) = 0
             End If
           End If
         Next KKI%
         If SALID# >= 0.05 Then
           IDENLOX$ = ""
           GoTo 35
         End If
         '
       End If
     Next U&
     '
     REBLA$ = REGBLAN$("STOLOTE"): JJ& = 0
     For KKI% = 1 To CANLOT%
       If CALOTE#(KKI%) > 0.05 Then
         ZZ$ = REBLA$
         Call REPLA(ZZ$, TLOTE$(KKI%), 1, 38)
         Call REPLA(ZZ$, MKD$(CALOTE#(KKI%)), 41, 8)
         JJ& = JJ& + 1
         Call GRAREG("!STOLOTE", ZZ$, JJ&)
       End If
     Next KKI%
     Call SETULTREG("!STOLOTE", JJ&)
     Call CIERRARCH("!STOLOTE")
     '
     Call FRESHECHO("!STOLOTE")
     Call SELECHO("!STOLOTE/Id.Lote                       F.Alta                Q.Orig                      Saldo")
     LSX$ = VALACT1$("!STOLOTE")
   End If
   '
   LOTESEL$ = LSX$
   '
End Function

Function GRUCOFAB%(CIJK%)
  '
  ' GRUCOFAB%=1 --> Compra
  ' GRUCOFAB%=2 --> Fabricacion
  '
  Static GRF(32767) As Integer
  '
  If GRF%(CIJK%) < 1 Then
    GRF%(CIJK%) = 1
    Call AREXPLO(CIJK%, 1, 1)
    '
    If CAIT% > 0 Then
      GRF%(CIJK%) = 2
      PBL% = PROVHABI%(CIJK%)
      If PBL% > 0 Then
        If PROVINT%(PBL%) = 0 Then
          GRF%(CIJK%) = 1
        End If
      End If
    End If
  End If
  '
  GRUCOFAB% = GRF%(CIJK%)
  '
End Function
' ***************************************************************************
Function GRUCOFAB_V2%(CIJK%)
  '
  ' GRUCOFAB%=1 --> Compra
  ' GRUCOFAB%=2 --> Fabricacion
  '
  Static GRF(32767) As Integer
  '
  If GRF%(CIJK%) < 1 Then
    GRF%(CIJK%) = 1
    Call AREXPLO(CIJK%, 1, 1)
    '
    If CAIT% > 0 Then
      GRF%(CIJK%) = 2
      PBL% = PROVHABI%(CIJK%)
      If PBL% > 0 Then
        ' SI EL ARTICULO TIENE PROVEEDOR HABITUAL Y A SU VEZ SECUENCIA DE OPERACIONES.
        ' SE CONSIDERA QUE SE FABRICA PERO EVENTUALMENTE SE COMPRA
        If CantRegistros("SECOPER", "CodIConj=" & CStr(CIJK%)) = 0 Then
          GRF%(CIJK%) = 1
        End If
      End If
    End If
  End If
  '
  GRUCOFAB_V2% = GRF%(CIJK%)
  '
End Function

'
Function COTIPIF$(Ci%)
  RECO$ = FILTERGETRECORD$("COEQTIP", 1, MKI$(Ci%))
  COTIPIF$ = TRIM$(Mid$(RECO$, 9, 24))
End Function

Function NUINV%()
'
X$ = REGLEIDO$("INVERT", 1&)
NUINV% = CVI(Mid$(X$, 17, 2))
End Function
' ***************************************************************************

Sub DESHASCOD(COD1$, COD2$, CIKX1%, CIKX2%)
   '
   LONCODI% = XVALO(Control$("MASTER", "LOMAXCOD"))
   If LONCODI% < 16 Or LONCODI% > 24 Then LONCODI% = 16

10 COD1$ = "": COD2$ = "": CIKX1% = 0: CIKX2% = 0
   CI1% = PRICOD%
   CI2% = ULTCOD%
   VDEF$ = MKI$(CI1%) + MKI$(CI2%)
   X$ = OBJPLA$("DESHASCOD", MKI$(CI1%) + MKI$(CI2%))
   If Len(X$) < 4 Then
     Exit Sub
   End If
   '
   Screen.MousePointer = 11
   '
   CIKX1% = CVI(Left$(X$, 2))
   COD1$ = CODEXT$(CIKX1%): If TRIM$(COD1$) = "" Then Exit Sub
   CIKX2% = CVI(Mid$(X$, 3, 2))
   COD2$ = CODEXT$(CIKX2%): If TRIM$(COD2$) = "" Then Exit Sub
   If CIKX1% < 1 Or CIKX2% < 1 Then GoTo 10
   If COD2$ < COD1$ Then GoTo 10
   '
End Sub
'
Function SECPERTE(EQP$)
   '
   EQP1$ = TRIM$(EQP$)
   If EQP1$ <> "" Then
      Y$ = FILTERGETRECORD$("PADMAQ", 1, EQP1$)
      Sec = CVS(Mid$(Y$, 37, 4))
      If Sec > 0 Then
         SECPERTE = Sec
      End If
   End If
   '
End Function

Function FECREALOT%(IDLOT$)
   '
   FECREALOT% = 0
   '
   Call APERBASDAT("STOCKS")
   SQLX$ = " SELECT FECHALTA FROM LOTINGRE WITH(NOLOCK) "
   SQLX$ = SQLX$ + "WHERE IdenLote = '" & IDLOT$ & "' "
   Set rs = READSET(SQLX$)
   '
   With rs
   If Not .EOF Then
      FECREALOT% = CVINT(!FechAlta)
   End If
   End With
   rs.Close
   Set rs = Nothing
   '
End Function

Function SALNOCONS(IDLOT$, CIXI%)
   '
   SALNOCONS = 0
   '
   Call APERBASDAT("STOCKS")
   SQLX$ = " SELECT * FROM LOTINGRE "
   SQLX$ = SQLX$ + "WHERE IdenLote = '" & IDLOT$ & "' "
   SQLX$ = SQLX$ + "AND Cod_Inte = " & CIXI% & " "
   '
   Dim LOTINGTEMP As ADODB.Recordset
   Set LOTINGTEMP = FLEXCONN.Execute(FILTSQL$(SQLX$))
   '
   With LOTINGTEMP
25   On Error Resume Next
     .MoveFirst
     If Err < 1 Then
        SALNOCONS = XVALO(!CanAlta) - XVALO(!CanConsu)
     End If
     On Error GoTo 0
   End With
   Set LOTINGTEMP = Nothing
   If SALNOCONS < 0 Then SALNOCONS = 0
   '
End Function


Sub ALTACODIGO(VDEF$, OKX%)
   '
   If EmpresaNoRealizaCambios% > 0 Then
      Exit Sub
   End If
   
   Call ABREMASTER 'abre conexion
   '
   LONCODI% = XVALO(Control$("MASTER", "LOMAXCOD"))
   If LONCODI% < 15 Or LONCODI% > 24 Then LONCODI% = 15
   '
10 OBX$ = OBJPLA$("ALTACODIGOSQL", VDEF$)
   If OBX$ = "" Then GoTo 99
   VDEF$ = OBX$
   CODINUE$ = TRIM$(Mid$(OBX$, 77, 24))
   '
   If Len(CODINUE$) < 3 Then
      Call MENSERR(24, "Ingreso no Válido.\Longitud Código Menor que 3 Caracteres.")
      GoTo 10
   End If
   '
   If Len(CODINUE$) > LONCODI% Then
      Call MENSERR(24, "Ingreso no Válido.\Longitud Código Mayor que " & LONCODI% & " Caracteres.")
      GoTo 10
   End If
   '
   If CODINT%(CODINUE$) <> 0 Then
      '\\\OT-06-0263-WAR-05/07/06///
      CIX1% = CODINT%(CODINUE$)
      If CIX1% < 0 Then 'El Articulo esta de Baja
        RTA1% = COMALTER("Código de Baja. Puede Ahora Restaurar\ el Código de Baja.\\Restaurar\Continuar")
        If RTA1% = 1 Then 'Restaurar
          Call ABREMASTER 'abre conexion
          
          'With Master
          '  .FindFirst ("CODINT = " & Abs(CIX1%))
          '  If .NoMatch = False Then
          '    .Edit
          '    !Status = 1
          '    If Mid$(Master!CODIGO, 16, 1) = Chr$(96) Then
          '      Master!CODIGO = CODINUE$
          '    End If
          '    .Update
          'SQL
          strSQL = ("Update Master set Status=1 where Codint=") & Abs(CIX1%)
          FLEXCONN.Execute strSQL
          
              UPDATMASTER% = 1
              Call GENECOMAS
              OKX% = 1 'OK
              GoTo 99
            End If
          'End With
        
        End If
      'End If
      '\\\OT-06-0263-WAR-FIN///
      Call MENSERR(24, "Ingreso no Válido.\Código Pre-Existente.")
      GoTo 10
   End If
   GRCONTA% = CVI(Mid$(OBX$, 3, 2))
   If ECOARCH$("GRUCOVEN", MKI(GRCONTA%)) = "" Then
     GRCONTA% = 0
   End If
   '
   If TRIM$(Mid$(OBX$, 5, 64)) = "" Then
      Call MENSERR(24, "Ingreso no Válido.\Falta Descripcion")
      GoTo 10
   End If
   DENUE$ = TRIM$(Mid$(OBX$, 5, 64))
   '
   TIMA% = Asc(Mid$(OBX$, 101, 1))
   If ECOARCH$("TATIMAT", Chr$(TIMA%)) = "" Then
      Call MENSERR(24, "Tipo de Material no Válido")
      GoTo 10
   End If
   '
   UNID$ = Mid$(OBX$, 109, 2)
   If ECOARCH$("UNIMED", UNID$) = "" Then
      Call MENSERR(24, "Unidad de Medida no Válida")
      GoTo 10
   End If
   
  If ECOARCH$("GRUCOVEN", MKI(GRCONTA%)) = "" Then
    GRCONTA% = 0
   End If

   '
   Screen.MousePointer = 11
   'CODII% = NUMAS% + 1
   CODII% = CODINT_NEW
   '
20 On Error Resume Next
   Call ABREMASTER 'abreconexion
   '
ESAHP% = 0
EPAC$ = TRIM$(UCase$(EMPREAC$))
If InStr(EPAC$, "AHP") > 0 Then
   ESAHP% = 1
End If

   Dim RMaster As ADODB.Recordset
   Set RMaster = New ADODB.Recordset
   RMaster.CursorType = adOpenDynamic
   RMaster.LockType = adLockPessimistic
   strSQL = ("Select * from Master where codint=" & CODII%)
25 On Error Resume Next
   RMaster.Open FILTSQL$(strSQL), FLEXCONN, , , adCmdText
   If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
      On Error GoTo 0
      Call CapturaErrorOpen
      GoTo 25
   End If
   On Error GoTo 0
   
   If Not (RMaster.EOF And RMaster.BOF) Then 'si no esta vacio por lo tanto existe
      Call MENSERR(24, "Ingreso no Válido.\Código Pre-Existente.")
      GoTo 10
   End If

   With RMaster
      .AddNew   ' YA CONVERTIDO
      On Error Resume Next
      !CODIEMPR = CodiEmp%
      On Error GoTo 0
      !Codigo = CODINUE$
      !Descripcion = DENUE$
      !Status = 1
      !CODINT = CODII%
      !UMEDIDA = UNID$
      !TIMATE = TIMA%
      !Grucoven = GRCONTA%
      If ESAHP% > 0 Then
        !TRAZABLE = 1
      End If
      On Error Resume Next
      .Update
      OKX% = 1
      If Err Then
         Call MENSERR(24, "Alta no Procesada Correctamente")
         OKX% = 0
      End If
      On Error GoTo 0
   End With
   '
   RMaster.Close
   Set RMaster = Nothing
   UPDATMASTER% = 1
   '
   If InStr(EPAC$, "MINEPARTS") < 1 And InStr(EPAC$, "AHP") > 0 Then
       CONDI$ = "CODINT = " & CODII%
       Call DuplicarRegistroDeUnaBaseAOtra("AHP", "MASTER", "MINEPARTS", "MASTER", CONDI$)
   End If
   
   If InStr(EPAC$, "CAMBACERES") > 0 And InStr(EPAC$, "QUILMES") < 1 Then
       CONDI$ = "CODINT = " & CODII%
       ' LE AGREGO DATOS POR DEFECTO
       SQLX$ = "UPDATE MASTER SET CONSUMO_X_CORTE = 1,TRAZABLE= 1,TIMATE=41,UMEDIDA= 'U.',LINDIVIS =  1 , STOMAX = 1 "
       If UCase$(Left$(TRIM$(CODINUE$), 4)) = "ITP-" Then SQLX$ = SQLX$ + " , FAMILIA = 'TI'"
       SQLX$ = SQLX$ + " WHERE " & CONDI$
       FLEXCONN.Execute (SQLX$)
'       'GENERA EL ARCHIVO EN EL DROPBOX.
'       ScriptParaInsert$ = ScriptDeInsert$("CAMBACERES", "MASTER", "QUILMES", "MASTER", CONDI$, , CodiEmp%)
'       Call GeneraArchivoStringAltaMaster(CODII%, CONDI$, ScriptParaInsert$)
   End If
   ' generar un registro para los saldos de stock
   'no se usa mas porque se remplaza por funcion stockact
'   Call ABRESTOCKS
'   Dim RSalStock As ADODB.Recordset
'   Set RSalStock = New ADODB.Recordset
'   RSalStock.CursorType = adOpenDynamic
'   RSalStock.LockType = adLockPessimistic
'
'   RSalStock.Open ("Select * from SalStock where CodiInte=" & CODII%), FLEXCONN, , , adCmdText
'
'   If (RSalStock.EOF And RSalStock.BOF) Then 'si esta vacio por lo tanto no existe
'     With RSalStock
'       .AddNew
'       !CodiInte = CODII%
'       !VeSalSto = String$(256, 0)
'       .Update
'       .Close
'     End With
'   End If
'
     '
99 Screen.MousePointer = 1
Set RMaster = Nothing
End Sub

Sub crearCodigoTelaImpresa(CI1Bolsa%, RET_NUEVOCOMPO%)
    If CI1Bolsa% < 1 Then Exit Sub
    SQLX$ = " SELECT CODIGO FROM MASTER WITH(NOLOCK) WHERE CODINT = " & CI1Bolsa%
    Set rs = READSET(SQLX$)
    If Not rs.EOF Then
      Codigo$ = rs!Codigo
    End If
    OPX% = COMALTER%("Atenciòn !!! El Sistema va a Crear un Nuevo Còdigo de Bobina Base Impresa para el Còdigo: " & Codigo$ & "\\Cancelar\Crear Nuevo Còdigo")
    If OPX% <= 1 Then Exit Sub
    
       VDEF$ = String$(128, 0)
       '
     
       CODIGOTELAIMPRESA$ = "I" + Right(Codigo$, Len(Codigo$) - 1)
       'PASA LOS DATOS DEL
       LONCODI% = XVALO(Control$("MASTER", "LOMAXCOD"))
       EXISTECOD% = CODINT%(CODIGOTELAIMPRESA$)
       If EXISTECOD% > 0 Then
          Call COMUNI("Imposible dar de Alta - Código Existente \(Còdigo: " & CODIGOTELAIMPRESA$ & ") Genere el Alta Manualmente con otro Còdigo")
          Exit Sub
       End If
       
       CI_NEW% = CODINT_NEW
       Call REPLA(VDEF$, MKI$(CI_NEW%), 1, 2)

        Call REPLA(VDEF$, "Tela Impresa de Bolsa Còd.: " & Codigo$, 5, 64)
        Call REPLA(VDEF$, CODIGOTELAIMPRESA$, 77, LONCODI%)
        Call REPLA(VDEF$, "U.", 109, 2)
        Call REPLA(VDEF$, Chr$(1), 101, 1)
        
       '
       OKX% = 0
       
       Call ALTACODIGO(VDEF$, OKX%)
       If OKX% < 1 Then Exit Sub
       If LONCODI% < 15 Or LONCODI% > 24 Then LONCODI% = 15
       If OKX% = 1 Then
          CODINUE$ = Mid$(VDEF$, 77, LONCODI%)
          RET_NUEVOCOMPO% = CI_NEW%
          ' *****  EL CODIGO SIGUIENTE SE PASO A LA RUTINA ALTACODIGO  *****
'          SQLX$ = "UPDATE MASTER SET CONSUMO_X_CORTE = 1,TRAZABLE= 1,TIMATE=41,UMEDIDA= 'U.',LINDIVIS =  1 , STOMAX = 1 "
'          SQLX$ = SQLX$ + " WHERE CODINT = " & CI_NEW%
'          FLEXCONN.Execute (SQLX$)
       End If
       '
       
       Screen.MousePointer = 11
       NOCERRARFRM% = 1 ' NO PERMITE CERRAR EL FORMULARIO
       If OKX% = 1 And ESPORCOPIA% > 0 Then
        Dim RMaster As ADODB.Recordset
        Set RMaster = New ADODB.Recordset
        strSQL$ = "Select * from Master WITH(NOLOCK) where codint=" & CII%
        Set RMaster = READSET(strSQL$)
        If Not (RMaster.EOF And RMaster.BOF) Then
           CACAM% = RMaster.Fields.Count
           For KU% = 1 To CACAM%
             Dato$ = SAFETEXT$(RMaster.Fields(KU% - 1).Value)
             NOCAMPO$ = TRIM$(UCase$(RMaster.Fields(KU% - 1).Name))
             If NOCAMPO$ <> "CODIGO" Then
                If NOCAMPO$ <> "DESCRIPCION" Then
                   If NOCAMPO$ <> "CODINT" Then
                      If NOCAMPO$ <> "ALIAS" Then
                         If NOCAMPO$ <> "UMEDIDA" Then
                           CONDI$ = "CODINT=" & CI_NEW%
                           Call ACTUREGISTRO("MASTER", NOCAMPO$, CONDI$, Dato$, REGAF&, "NOMESS")
                         End If
                      End If
                   End If
                End If
              End If
           Next KU%
        End If
       End If
       'ACTUALIZO EL MASTER.
       CONDI$ = " CODIGO <> '' AND STATUS > 0 ORDER BY CODIGO ASC"
       If Control$("", "IDSECUN") = "ALIAS" Then
           Call CARGALISEL("ECOMAST", "MASTER", "CODIGO/A24;DESCRIPCION/A36;NPLANO/A24", CONDI$, "NOMESS")
       Else
           Call CARGALISEL("ECOMAST", "MASTER", "CODIGO/A24;DESCRIPCION/A50", CONDI$, "NOMESS")
       End If
        '
       Screen.MousePointer = 1
End Sub


Sub GeneraArchivoStringAltaMaster(CI1%, CONDI$, StringInsert$)
        'ESTO ES ESTANDO PARADO EN  CAMBACERES
        
        'EJEMPLO GENERA UN ARCHIVO Y LO GUARDA EN LA RUTA CONFIGURADA
        'EJEMPLO:SEPARA LA CONDICION DEL INSERT CON: |~|
        'CODINT = 1143
        '|~|INSERT INTO QUILMES.dbo.MASTER (Ferevis,CodiEmpr,Observa,Especificacion,DesCompAut,DescriComp,.........
        ESCAMBACERES% = 0
        EPAC$ = TRIM$(UCase$(EMPREAC$))
        If InStr(EPAC$, "CAMBACERES") > 0 Then
           ESCAMBACERES% = 1
        End If
        '
        longCondi = Len(CONDI$)
        If longCondi < 1 Then Exit Sub
        If longCondi > 1024 Then
           Call COMUNI("Imposible Procesar - Condicion Supera los 1024 caracteres")
           Exit Sub
        End If
        longStringInsert = Len(StringInsert$)
        If longStringInsert < 1 Then Exit Sub
        
        Ruta$ = ""
        '
        CARPETASINCRO$ = TRIM$(Control("MASTER", "DROPMAST"))
        Ruta$ = CARPETASINCRO$
        If Ruta$ = "" Then
           Call COMUNI("Imposible Continuar.\Falta Definir Carpeta de Intercambio de Alta Item Cambaceres")
           Exit Sub
        End If
        
        A$ = Date
        A$ = REPLACAR(A$, "/", "_")
        B$ = Time
        B$ = REPLACAR(B$, ".", "")
        B$ = REPLACAR(B$, ":", "")
        NombreArchivo$ = "AltaCodigo_" & TRIM$(CStr(CI1%))
        NARCHI2% = FILEOPEN%(Ruta$ + "\" + NombreArchivo$ & "_" & A$ & " " & B$ & ".TXT", 2, 1024)    ' ABRE EL ARCHIVO DE ESCRITURA
        
        TX$ = TRIM$(SAFETEXT$(CONDI$)) & "|~|"
        Print #NARCHI2%, TX$
        Print #NARCHI2%, StringInsert$

        Close #NARCHI2%
        '''
End Sub




Sub EjecutarALtaMasterQuilmes(PathArchivoCompleto$, OK%)
'_________________________________________________________________________________________________________
   'SE PUEDE PASAR CUALQUIER TABLA A CA SE DETECTA LUEGO DEL INSERT INTO  A QUE TABLA PERTENECE
'______________ESTO ES PARADO EN QUILMES___________________________________________________________________________________________
        OK% = 0
        Screen.MousePointer = 11
        '****************
        'EJEMPLO DE ARCHIVO A LLER
        'CODINT = 1143
        '|~|INSERT INTO QUILMES.dbo.MASTER (Ferevis,CodiEmpr,Observa,Especificacion,DesCompAut,DescriComp,.......

        
         VUELTA% = 0: INDX% = 1
         LECTURA$ = ""
         NARCHI1% = FILEOPEN%(PathArchivoCompleto$, 1, 1024) 'ABRE EL ARCHIVO DE LECTURA
         Do While Not EOF(NARCHI1%) ' RECORRE
            Line Input #NARCHI1%, XX$ ' LEE CADA REGISTRO
            LECTURA$ = LECTURA$ & XX$
         Loop
         Close #NARCHI1%
         
         strArray = Split(LECTURA$, "|~|")
         CONDI$ = TRIM$(strArray(0))
         If CONDI$ <> "" Then
            SQLX$ = (strArray(1))
            Pos_Tabla% = InStr(TRIM$(UCase$(SQLX$)), "INSERT INTO")
            If Pos_Tabla% > 0 Then
                sqlxMayusculas$ = TRIM$(UCase$(SQLX$))
                restoSqlx$ = LTrim$(Mid$(sqlxMayusculas$, Pos_Tabla% + 11))
                PosEspacioDespuesDeNombreTabla% = InStr(restoSqlx$, " ")
                If PosEspacioDespuesDeNombreTabla% > 0 Then
                    On Error GoTo ERROR_GUARDAR
                    TABLA$ = TRIM$(Mid$(SQLX$, Pos_Tabla% + 11, PosEspacioDespuesDeNombreTabla%))
                    strSQL$ = "DELETE  " & TABLA$ & " WHERE " & CONDI$
                    FLEXCONN.BeginTrans
                    FLEXCONN.Execute (strSQL$)
                    FLEXCONN.Execute (SQLX$)
                End If
            End If
         End If
'         FALTA VER EL TEMA DE LA ESTRUCTURA
ERROR_GUARDAR:
    If Err <> 0 Then
       Screen.MousePointer = 1
       FLEXCONN.RollbackTrans
       On Error GoTo 0
       Call COMUNI("ERROR AL GUARDAR NO SE REALIZO LA TRANSACCION")
       Exit Sub
    Else
      FLEXCONN.CommitTrans
      Call Kill(PathArchivoCompleto$)
      UPDATMASTER% = 1: Call GENECOMAS
  '
      OK% = 1
    End If
         
      Screen.MousePointer = 1
         
         

End Sub


Sub ALTARELACODIGO(VDEF$, OKX%)
   '
   If EmpresaNoRealizaCambios% > 0 Then
      GoTo 99
   End If
   
   Call ABREMASTER 'abre conexion
   '
   OBX$ = VDEF$
   OKX% = 0
   LONCODI% = XVALO(Control$("MASTER", "LOMAXCOD"))
   If LONCODI% < 15 Or LONCODI% > 24 Then LONCODI% = 15
   '
   CODINUE$ = TRIM$(Mid$(OBX$, 77, 24))
   '
   If Len(CODINUE$) < 3 Then
      Call MENSERR(24, "Ingreso no Válido.\Longitud Código Menor que 3 Caracteres.")
      GoTo 99
   End If
   '
   If Len(CODINUE$) > LONCODI% Then
      Call MENSERR(24, "Ingreso no Válido.\Longitud Código Mayor que " & LONCODI% & " Caracteres.")
      GoTo 99
   End If
   '
   If CODINT%(CODINUE$) <> 0 Then
      CIX1% = CODINT%(CODINUE$)
      If CIX1% < 0 Then 'El Articulo esta de Baja
        RTA1% = COMALTER("Código de Baja. Puede Ahora Restaurar\ el Código de Baja.\\Restaurar\Continuar")
        If RTA1% = 1 Then 'Restaurar
          Call ABREMASTER 'abre conexion
          
          strSQL = ("Update Master set Status=1 where Codint=") & Abs(CIX1%)
          FLEXCONN.Execute strSQL
          UPDATMASTER% = 1: Call GENECOMAS
              OKX% = 1 'OK
              GoTo 99
            End If
        End If
      Call MENSERR(24, "Ingreso no Válido.\Código Pre-Existente.")
      GoTo 99
   End If
   GRCONTA% = 0
   CODIORIGINAL% = CVI(Mid$(OBX$, 67, 2))
   If CODIORIGINAL% < 1 Then
    If TRIM$(Mid$(OBX$, 3, 64)) = "" Then
       Call MENSERR(24, "Ingreso no Válido.\Falta Descripcion")
       GoTo 99
    End If
   End If
   DENUE$ = TRIM$(Mid$(OBX$, 3, 64))
   '
   TIMA% = 1
   '
   UNID$ = "U."
   '
   Screen.MousePointer = 11
   'SI TIENE  CODIGO DE RELACION RELACIONA EN LA TABLA CODICLIENTE Y SALE
   If CODIORIGINAL% > 0 Then
      '
      'VALIDA SI YA ESTA RELACIONADO
      CONDI$ = "Codigo_Cliente = '" & CODINUE$ & "' AND CODINT = " & CODIORIGINAL%
      EXISTECODNUE% = CantRegistros&("CODICLIENTE", CONDI$, "NOMESS")
      If EXISTECODNUE% > 0 Then
         Call COMUNI("Código yá Relacionado Con Item Seleccionado")
         GoTo 99
      End If
      Dim rs As New ADODB.Recordset
      SQLX$ = " SELECT * FROM CODICLIENTE "
      SQLX$ = SQLX$ + " WHERE CODINT = " & CODIORIGINAL%
26    On Error Resume Next
      rs.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
      If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
         On Error GoTo 0
         Call CapturaErrorOpen
         GoTo 26
      End If
      On Error GoTo 0
      With rs
         .AddNew
         On Error Resume Next
         !CODIEMPR = CodiEmp%
         On Error GoTo 0
         !Codigo_Cliente = UCase$(CODINUE$)
         !Status = 1
         !CODINT = CODIORIGINAL%
         On Error Resume Next
         .Update
         OKX% = 2
         If Err Then
           Call MENSERR(24, "Relación de Código no Procesada Correctamente")
           OKX% = 0
         End If
         On Error GoTo 0
      End With
      rs.Close
      Set rs = Nothing
      GoTo 99
   End If
   'FIN RELACION CODIGO CLIENTE (SI ENTRO ACA SE VA )
   CODII% = CODINT_NEW
   '
20 On Error Resume Next
   Call ABREMASTER 'abreconexion
   '
   Dim RMaster As ADODB.Recordset
   Set RMaster = New ADODB.Recordset
   RMaster.CursorType = adOpenKeyset
   RMaster.LockType = adLockPessimistic
   strSQL = ("Select * from Master where codint=" & CODII%)
25 On Error Resume Next
   RMaster.Open FILTSQL$(strSQL), FLEXCONN, , , adCmdText
   If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
      On Error GoTo 0
      Call CapturaErrorOpen
      GoTo 25
   End If
   On Error GoTo 0
   
   If Not (RMaster.EOF And RMaster.BOF) Then 'si no esta vacio por lo tanto existe
      Call MENSERR(24, "Ingreso no Válido.\Código Pre-Existente.")
      GoTo 99
   End If

   With RMaster
      .AddNew   ' YA CONVERTIDO
      On Error Resume Next
      !CODIEMPR = CodiEmp%
      On Error GoTo 0
      !Codigo = CODINUE$
      !Descripcion = DENUE$
      !Status = 1
      !CODINT = CODII%
      !UMEDIDA = UNID$
      !TIMATE = TIMA%
      !Grucoven = GRCONTA%
      !TRAZABLE = 1
      On Error Resume Next
      .Update
      OKX% = 1
      If Err Then
         Call MENSERR(24, "Alta no Procesada Correctamente")
         OKX% = 0
      End If
      On Error GoTo 0
   End With
   '
   RMaster.Close
   Set RMaster = Nothing
   '
   EPAC$ = TRIM$(UCase$(EMPREAC$))

   If OKX% = 1 Then
      If InStr(EPAC$, "MINEPARTS") < 1 And InStr(EPAC$, "AHP") > 0 Then
       CONDI$ = "CODINT = " & CODII%
       Call DuplicarRegistroDeUnaBaseAOtra("AHP", "MASTER", "MINEPARTS", "MASTER", CONDI$)
      End If
   End If
   
99 Screen.MousePointer = 1
On Error Resume Next
Set RMaster = Nothing
On Error GoTo 0
End Sub


Function ESTRAZABLE%(Ci%, Optional REAL%)
   '
   
   On Error Resume Next
   If STATUIT%(Ci%) <> 0 Then ESTRAZABLE% = CDbl(MasterTemp!TRAZABLE)
   'los que son -1 tambien llevan n.serie pero con esta marca no genera n.serie
   If ESTRAZABLE% = -1 And REAL% > 0 Then
      ESTRAZABLE% = -1
      Exit Function
   End If
   
   If ESTRAZABLE% = -1 Then ESTRAZABLE% = 2
   '
End Function
Function MOLDE%(CI1%)
   '
   On Error Resume Next
   If STATUIT%(CI1%) <> 0 Then MOLDE% = CDbl(MasterTemp!MOLDE)
   '
End Function

Function MARCALOT$(CIXI%, IDLOT$)
    '
    LTP1$ = ""
    Call APERBASDAT("STOCKS")
    '
    SQLX$ = " SELECT Marca FROM LOTINGRE WITH(NOLOCK)"
    SQLX$ = SQLX$ + " WHERE Cod_Inte = " & CIXI%
    SQLX$ = SQLX$ + " AND IdenLote = '" & IDLOT$ & "' "
    '
    Dim LOTINGTEMP As ADODB.Recordset
    Set LOTINGTEMP = FLEXCONN.Execute(FILTSQL$(SQLX$))
    With LOTINGTEMP
       If Not (.EOF And .BOF) Then
          LTP1$ = SAFETEXT$(!MARCA)
       End If
    End With
    '
    Set LOTINGTEMP = Nothing
    MARCALOT$ = LTP1$
    '
End Function
         
Function LOTORIG$(CIXI%, IDLOT$)
    '
    LTP1$ = ""
    Call APERBASDAT("STOCKS")
    '
    SQLX$ = " SELECT LoteCol_Prov FROM LOTINGRE WITH(NOLOCK)"
    SQLX$ = SQLX$ + " WHERE Cod_Inte = " & CIXI%
    SQLX$ = SQLX$ + " AND IdenLote = '" & IDLOT$ & "' "
    '
    Dim LOTINGTEMP As ADODB.Recordset
    Set LOTINGTEMP = FLEXCONN.Execute(FILTSQL$(SQLX$))
    With LOTINGTEMP
       If Not (.EOF And .BOF) Then
          LTP1$ = SAFETEXT$(!LOTECOL_PROV)
       End If
    End With
    '
    Set LOTINGTEMP = Nothing
    LOTORIG$ = LTP1$
    '
End Function
'
Function PROXILOTE$(LOTEORI$)
    '
    NCR$ = TRIM$(LOTEORI$)
    PPXX% = InStr(NCR$, "/")
    If PPXX% > 0 Then
       NCR$ = Left$(NCR$, PPXX% - 1) ' VERIFICA LA EXISTENCIA DE LA BARRA IDENTIFICADORA DE SUBLOTE
    End If
    '
    Call APERBASDAT("STOCKS")
    For KKII% = 65 To 90
      NCR1$ = NCR$ + "/" + Chr$(KKII%) ' ARMA EL NUEVO SUBLOTE
      QTR& = CantRegistros("LOTINGRE", "IdenLote = '" & NCR1$ & "'") ' VERIFICA LA EXISTENCIA DEL SUBLOTE
      If QTR& < 1 Then Exit For
    Next KKII%
    ' OJO QUE SI SALE POR AQUÍ QUEDA EN /Z
    PROXILOTE$ = NCR1$
    '
End Function
'
Function UBILOTE$(CIXI%, IDLOT$, Optional MODO%)
    '
    LTP1$ = ""
    Call APERBASDAT("STOCKS")
    '
    SQLX$ = " SELECT DepoUbic, Ubicfis FROM LOTINGRE WITH(NOLOCK)"
    SQLX$ = SQLX$ + " WHERE Cod_Inte = " & CIXI%
    SQLX$ = SQLX$ + " AND IdenLote = '" & IDLOT$ & "' "
    '
    Dim LOTINGTEMP As ADODB.Recordset
    Set LOTINGTEMP = FLEXCONN.Execute(FILTSQL$(SQLX$))
    DPX% = 0: UBX$ = ""
    With LOTINGTEMP
       If Not (.EOF And .BOF) Then
          DPX% = XVALO(!DEPOUBIC)
          UBX$ = SAFETEXT$(!UbicFis)
       End If
    End With
    '
    LOTINGTEMP.Close
    Set LOTINGTEMP = Nothing
    '
    DPXX$ = ""
    If MODO% < 1 Then
      If DPX% > 0 Then
        DPXX$ = ECOARCH$("TADEPOSI", Chr$(DPX%))
        If DPXX$ <> "" Then DPXX$ = DPXX$ + " - "
      End If
    End If
    '
    UBILOTE$ = DPXX$ + UBX$
    '
End Function
'
Function DEPOLOTE%(CIXI%, IDLOT$)
    '
    Call APERBASDAT("STOCKS")
    CONDI$ = "COD_INTE = " & CIXI% & " AND IDENLOTE ='" & IDLOT$ & "'"
    DEPOLOTE% = XVALO(VALOBADA("LOTINGRE", "DEPOUBIC", CONDI$, "NOMESS"))
    '
End Function
'
Function AUTONOMIASM$(Ci%)
   'Leer dato en campo Autono_Proy
   PSM$ = ""
   CONDI$ = "CODINT = " & Ci%
   PSM$ = TRIM$(VALOBADA("MASTER", "Autono_Proy", CONDI$, "NOMESS"))
   '
   If PSM$ <> "" Then
      AUTONOMIASM$ = PSM$
      GoTo 99
   End If
   '
   TIMAT% = TIMATE%(Ci%)
   If TIMAT% > 0 Then ' FUNCION QUE DEVUELVE EL LEAD TIME SEGUN EL TIPO DE MATERIAL
    XX$ = FILTERGETRECORD$("TABTIMA", 1, MKI$(TIMAT%))
    PSM$ = Mid$(XX$, 82, 6)
   End If
99 AUTONOMIASM$ = PSM$

End Function
Function LEAD_TIME%(Ci%)

   'Leer dato en campo LeTiCom y comparar para seguir
   LTME% = 0
   CONDI$ = "CODINT = " & Ci%
   LCOM$ = SAFETEXT(VALOBADA("MASTER", "LeTiCom", CONDI$, "NOMESS"))
   
   If LCOM$ <> "" Then
      LTME% = XVALO(LCOM$)
      GoTo 99
   End If
   
   TIMAT% = TIMATE%(Ci%)
   If TIMAT% > 0 Then ' FUNCION QUE DEVUELVE EL LEAD TIME SEGUN EL TIPO DE MATERIAL
    XX$ = FILTERGETRECORD$("TABTIMA", 1, MKI$(TIMAT%))
    LTME% = CVI(Mid$(XX$, 18, 2))
   End If
99 LEAD_TIME% = LTME%

End Function

Function VALCODPROV%(NPRO%, CIX%)  ' FUNCION QUE VALIDA QUE EL CODIGO DEL ARTICULO
    '                              SE ENCUENTRE ASOCIADO A UN PROVEEDOR
    Static CODPROVE%
    '
    If CODPROVE% = 0 Then
       CODPROVE% = -1
       If Control$("OCOMPRA", "VALCOPRO") = "SI" Then CODPROVE% = 1
    End If
    VALCODPROV% = 1
    If CODPROVE% < 1 Or CIX% < 1 Then Exit Function
    '
    CONDI$ = "Num_Prov = " & NPRO% & " and Cod_Inte = " & CIX%
    If CantRegistros("PROVECOD", CONDI$) < 1 Then
      Call MENSERR(24, "Código Artículo '" & CODEXT$(CIX%) & "' no Relacionado al\Proveedor Seleccionado")
      VALCODPROV% = 0
    End If
    '
End Function

Sub CONVERUNID(USON, UNXX$)
    '
10  If Abs(USON) > 0 Then
      If Abs(USON) < 1 Then
        For KKIL& = 1 To ULTREG&("UNIMED")
          RLX$ = REGLEIDO$("UNIMED", KKIL&)
          ULEI$ = UCase$(TRIM$(Mid$(RLX$, 1, 4)))
          If ULEI$ = UCase$(UNXX$) Then
            UREP$ = UCase$(TRIM$(Mid$(RLX$, 33, 4)))
            If UREP$ <> ULEI$ Then
              FACO = CVS(Mid$(RLX$, 37, 4))
              If FACO > 1 Then
                UNXX$ = TRIM$(Mid$(RLX$, 33, 4))
                USON = USON * FACO
                GoTo 10
              End If
            End If
          End If
        Next KKIL&
      End If
    End If
End Sub

Function FACONVERT(UNISTO$, UNIFOR$)
    '
    FACOV = 1
    UNSTO$ = TRIM$(UCase$(UNISTO$))
    UNFOR$ = TRIM$(UCase$(UNIFOR$))
    '
10  If UCase$(UNSTO$) <> UCase$(UNFOR$) Then
      For KKIL& = 1 To ULTREG&("UNIMED")
         RLX$ = REGLEIDO$("UNIMED", KKIL&)
         ULEI$ = UCase$(TRIM$(Mid$(RLX$, 33, 4)))
         If ULEI$ = UNFOR$ Then
           UREP$ = UCase$(TRIM$(Mid$(RLX$, 1, 4)))
           If UREP$ <> ULEI$ Then
              FACO = CVS(Mid$(RLX$, 37, 4))
              If FACO > 1 Then
                UNFOR$ = TRIM$(UCase$(Mid$(RLX$, 1, 4)))
                FACOV = FACOV * FACO
                GoTo 10
              End If
           End If
         End If
      Next KKIL&
    End If
    '
    FACONVERT = FACOV
    '
End Function

Function REDON_INDIVIS(Ci%, CANT#, Optional MODO$)
   '
   'SI MODO$ ES IGUAL A VENTA ES CALCULA EL LOTE INDIVISIBLE DE VENTAS
   If UCase$(MODO$) = "VENTA" Then LINDI = LINDIVEN(Ci%): GoTo 10
   LINDI = LINDIVIS(Ci%)
   '
10 If LINDI < 0.00001 Then
      REDON_INDIVIS = CANT#
      Exit Function
   End If
   CANTNUE& = Int(CANT# / LINDI)
   Resto# = CANT# - (CANTNUE& * XVALO(LINDI))
   
   If Resto# > 0 Then
      CANTNUE& = CANTNUE& + 1
   End If
   
   REDON_INDIVIS = LINDI * CANTNUE&
   
End Function

Function PRESENVASE$(Ci%)
    
    PRVASE$ = ""
    If Ci% > 0 Then
      CAN = ROUND(LINDIVIS(Ci%), 0)
      If CAN >= 1 Then
        EMBA$ = TRIM$(ATRIB1$(Ci%))
        If EMBA$ = "" Then EMBA$ = "Bultos"
        TIPUN$ = Unimed$(Ci%)
        '
        PRVASE$ = EMBA$ & " de " & CAN & " " & TIPUN$
        '
      End If
    End If
    '
99  PRESENVASE$ = PRVASE$

End Function

Function CANBULTOS(Ci%, CANT#, Optional MODO$)
   '
   CANB = 0
   '
   ' SI MODO ES VENTA TRAE EL LOTE INDIVISIBLE DE VENTAS
   If Ci% > 0 Then
     If CANT# > 0 Then
       If UCase$(MODO$) = "VENTA" Then
          CANPORBTO = ROUND(LINDIVEN(Ci%), 0)
       Else
          CANPORBTO = ROUND(LINDIVIS(Ci%), 0)
       End If
       '
       If CANPORBTO >= 1 Then
          CANB = Int(CANT# / CANPORBTO)
          Resto# = CANT# - (CANB * CANPORBTO)
          If Resto# > 0 Then
            CANB = CANB + 1
          End If
       End If
     End If
   End If
   '
   CANBULTOS = CANB
   '
End Function
Function CANBULTOSXKG(Ci%, CANT#)
   '
   CANB = 0
   '
   PESOX = PESUNI(Ci%)
   ' SI MODO ES VENTA TRAE EL LOTE INDIVISIBLE DE VENTAS
   If Ci% > 0 Then
     If CANT# > 0 Then
       CANPORBTO = ROUND(PESOX, 0)
       '
       If CANPORBTO >= 1 Then
          CANB = Int(CANT# / CANPORBTO)
          Resto# = CANT# - (CANB * CANPORBTO)
          If Resto# > 0 Then
            CANB = CANB + 1
          End If
       End If
     End If
   End If
   '
   CANBULTOSXKG = CANB
   '
End Function


Function BIENCAP%(Ci%)
   '
    BICA% = 0
    If Ci% < 1 Then GoTo 99
    '
    CONDI$ = "CODINT = " & Ci%
    BICA% = XVALO(VALOBADA("MASTER", "BIEN_CAP", CONDI$, "NOMESS"))
   
99  BIENCAP% = BICA%
   
End Function

Function ConsuSegunCorte%(Ci%)
   '
   ConsuSegunCorte% = 0
   On Error Resume Next
   If STATUIT%(Ci%) <> 0 Then ConsuSegunCorte% = XVALO(MasterTemp!Consumo_X_Corte)
   '
End Function

Function STATMOVI%(Ci%)
   '
   STATMOVI% = 0
   On Error Resume Next
   If STATUIT%(Ci%) <> 0 Then STATMOVI% = XVALO(MasterTemp!Stat_Movi)
   '
End Function

Function UNIMEAFIP%(Ci%)

   'MODIFICADO 09 POR QUE NO DEVOLIVA POR DEFECTO LA 'UNIDAD' 7
   UNIMEAFIP% = 7 ' POR DEFECTO DEVUELVE UNIDAD
   
   If Ci% < 1 Then Exit Function
   '
   UNIM$ = VALOBADA("MASTER", "UMEDIDA", "CODINT =" & Ci%, "NOMESS")
   X$ = FILTERGETRECORD$("UNIMED", 1, TRIM$(UNIM$))
   UMEDAFIP% = CVI(Mid$(X$, 41, 2))
   '
   If UMEDAFIP% > 0 Then
     UNIMEAFIP% = UMEDAFIP%
   End If
   '
End Function
Sub CALCUVALRES(CODIMAQ$, Optional FECHASEL%)
    ' GRABA EL VALOR RESIDUAL EN LA TABLA
    
    'CALCULO LA DIFERENCIA DE AÑOS CALCULAR EL VALOR RESIDUAL
    CODEQUIPO$ = TRIM$(CODIMAQ$)
    '
    CALCUGENERAL% = 0
    'VALIDO SI SE PASO FECHA HASTA SI NO TOMA LA DE HOY
    If FECHASEL% > 0 Then ' FECHA DE RECALCULO TOTAL
      FE% = FECHASEL%
      CALCUGENERAL% = 1
    Else
      FE% = HOY(HO$) ' SI NO SE LE PASA FECHA TOMA LA DEL DIA
    End If
    '
    'FECHA  PARA TOMAR AÑO
    fech = FECHATEXLAR$(FE%): ANIOACTUAL = Mid$(fech, 7)
    'FECHA ACTIVACION PARA TOMAR EL AÑO
    Call ABREMAQUI
    CONDI$ = "CODIGO_EQU =  '" & CODEQUIPO$ & "'"
    FECHACTIV% = CVINT(VALOBADA("EQUIPOS", "FECHAACTIV_EQU", CONDI$, "NOMESS"))
    FECACTIV = FECHATEXLAR$(FECHACTIV%): ANIOACTIV = XVALO(Mid$(FECACTIV, 7))
    IMPORTEVTA = XVALO(VALOBADA("EQUIPOS", "ImpVta_Equ", CONDI$, "NOMESS"))

    '
    'CALCULO EL PORCENTAJE DE DEPECIACION EN BASE AL PLAZO DE AMORTIZACION CONTABLE
    PLAZAMORT% = 0
    PLAZAMORT% = XVALO(VALOBADA("EQUIPOS", "PLAZOAMORT_EQU", CONDI$, "NOMESS"))
    If PLAZAMORT% > 0 Then
      DEPRECANUAL = 100 / PLAZAMORT%
    End If
    '
    'CALCULO EL PORCENTAJE DE DEPECIACION EN BASE AL PLAZO DE AMORTIZACION FISCAL
    PLAZAMORTFIS% = 0
    PLAZAMORTFIS% = XVALO(VALOBADA("EQUIPOS", "PLAZOAMORTFISC_EQU", CONDI$, "NOMESS"))
    If PLAZAMORTFIS% > 0 Then
      DEPRECANUALFIS# = 100 / PLAZAMORTFIS%
    End If
    VALORORIG = XVALO(VALOBADA("EQUIPOS", "VALORIGPESOS_EQU", CONDI$, "NOMESS"))
    ANIOSTRANSC = 1 + (XVALO(ANIOACTUAL) - XVALO(ANIOACTIV))
    '
    If ANIOSTRANSC <= 0 Then
       VALORCALCULADO = VALORORIG
       ANIOSTRANSC = 0
       GoTo 50
    End If
    '
    Resultado = XVALO((VALORORIG * ANIOSTRANSC) * (DEPRECANUAL / 100))
    RESULTADOFISC = (VALORORIG * ANIOSTRANSC * ROUND(DEPRECANUALFIS#, 2))
    '
    'SI EL RESUTADO ES MAYOR AL VALOR ORIGINAL O ESTA VENDIDO
    If Resultado > VALORORIG Or IMPORTEVTA > 0 Then
         NUEVOVALOR = 0
    Else
         NUEVOVALOR = VALORORIG - XVALO(Resultado)
    End If
    '
    If RESULTADOFISC > VALORORIG Or IMPORTEVTA > 0 Then
         NUEVOVALORFISC = 0
    Else
         NUEVOVALORFISC = VALORORIG - XVALO(RESULTADOFISC)
    End If

    '
    'VALIDA QUE LA FECHA ACTUAL O LA FECHA DE CALCULO NO SEA < A LA FECHA DE ACTIVACION
    'ESTO ES PARA QUE SI EL BIEN TIENE FECHA DE ACTIVACION POSTERIOR A LA FECHA DE
    ' CALCULO O LA DE HOY EL VALOR RESIDUAL VA A SER 0 POR QUE A ESA FECHA EL BIEN NO EXISTE
    If FE% < FECHACTIV% Then NUEVOVALOR = 0
    '
    'VALIIDO SI ES POR CALCULO GENERAL ACTUALIZO EL MISMO Y LA FECHA DE CALCULO
50  If CALCUGENERAL% > 0 Then
       Call ACTUREGISTRO("EQUIPOS", "ValResCalc_Equ", CONDI$, NUEVOVALOR, REGAF&, "NOMESS")
       Call ACTUREGISTRO("EQUIPOS", "ValResCalcFisc_Equ", CONDI$, NUEVOVALORFISC, REGAF&, "NOMESS")
       FECHACALCU = FETEXCOR1$(FE%) 'FECHA DE CALCULO
       Call ACTUREGISTRO("EQUIPOS", "FechCalc_Equ", CONDI$, FECHACALCU, REGAF&, "NOMESS")
    
    Else ' SI ES ACTUALIZACION INDEPENDIENTE ACTUALIZO EL VALOR RESIDUAL COMUN
       Call ACTUREGISTRO("EQUIPOS", "ValorResid_Equ", CONDI$, NUEVOVALOR, REGAF&, "NOMESS")
       Call ACTUREGISTRO("EQUIPOS", "ANIOSTRANSC_EQU", CONDI$, ANIOSTRANSC, REGAF&, "NOMESS")
       Call ACTUREGISTRO("EQUIPOS", "ValorResidFisc_Equ", CONDI$, NUEVOVALORFISC, REGAF&, "NOMESS")
    End If
   
   'SE GRABA SIEMPRE
    Call ACTUREGISTRO("EQUIPOS", "DEPRECANUAL_EQU", CONDI$, FORMATNUM$(DEPRECANUAL, "F12.2"), REGAF&, "NOMESS")
    Call ACTUREGISTRO("EQUIPOS", "DEPRECANUALFISC_EQU", CONDI$, FORMATNUM$(DEPRECANUALFIS, "F12.2"), REGAF&, "NOMESS")
    '
End Sub

 Function DERIMPO(Ci%)
   On Error Resume Next
   TIMA% = TIMATE%(Ci%)
   XX$ = FILTERGETRECORD$("TABTIMA", 1, MKI$(TIMA%))
   DERIM = CVS(Mid$(XX$, 88, 4))
   DERIMPO = DERIM
End Function

Function DEREXPO(Ci%)
   On Error Resume Next
   TIMA% = TIMATE%(Ci%)
   XX$ = FILTERGETRECORD$("TABTIMA", 1, MKI$(TIMA%))
   DEREX = CVS(Mid$(XX$, 92, 4))
   DEREXPO = DERIM
End Function

Function CODINTBARRA%(UPC$)
   '
   CODI% = 0
   If TRIM(UPC$) = "" Then GoTo 99
   '
   CODI% = XVALO(VALOBADA("MASTER", "CODINT", "COLOR = '" & UPC$ & "'", "NOMESS"))
   
99 CODINTBARRA% = CODI%
End Function

Function VOLUMENCOD#(Ci%)
    '
    M3Caja# = XVALO(VALOBADA("MASTER", "M3CAJA", "CODINT=" & Ci%, "NOMESS"))
    VOLUMENCOD# = M3Caja#
    '
End Function

Sub ADD_NOVEDAD(CI1%, NPROV%, PREC_ACT#, PREC_NUE#, OBSERVAC$, MONED%, MODO$)
    'NOVEDADES DE CAMBIOS DE PRECIO
    
    ' 'LISTA' SI ES ACTUALIZACION DE LISTA
    '"COSTO" SI ES ACTUALIZACION DE COSTO POR  PROVEEDOR
    User% = USNBR% Mod 100
    FE% = HOY(HO$)
    '
    'TOMA EL VALOR EN DOLARES
    If XVALO(MONED%) < 1 Then MONED% = 1
    COTIZAX = TICAMONE(MONED%)
    
    '
    'ESTO ES PARA CUANDO SE INICIALIZA UN ITEM QUE NO REGISTRA IMPORTE ACTUAL
    If PREC_ACT# < 0.005 Then PREC_ACT# = PREC_NUE#
    
    If CI1% > 0 Then
       SQLX$ = "SELECT * FROM NOVEDADES "
       SQLX$ = SQLX$ + " WHERE CODINT = -1"
       '
       Set Rst = New ADODB.Recordset
25     On Error Resume Next
       Rst.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
       If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
         On Error GoTo 0
         Call CapturaErrorOpen
         GoTo 25
       End If
       On Error GoTo 0
       
       With Rst
          .AddNew
          !CODINT = CI1%
          !Nume_Usuar = User%
          !Fech_Mod = Now
          !Prec_Ant = PREC_ACT# * XVALO(COTIZAX) ' EL PRECIO  SE GUARDA SIEMPRE EN PESOS
          !Prec_New = PREC_NUE# * XVALO(COTIZAX) ' EL PRECIO  SE GUARDA SIEMPRE EN PESOS
          !NumProv = NPROV%
          !OBSERVACION = Left$(OBSERVAC$, 512)
          !Moneda = MONED%
          !Lista_Repos = MODO$
          !COTIZA_DOLAR = XVALO(COTIZAX)
          .Update
       End With
       '
       Rst.Close
       Set Rst = Nothing
       '
    End If
    '
End Sub
Function COSTO_SEGUN_PROV(NPROV&, CI1%)
   'COSTO SEGUN PROVEEDOR GUARDADO EN LA TABLA RELACODPROV
   CONDI$ = "NUM_PROV = " & NPROV& & " AND COD_INTE = " & CI1%
   COSTO_SEGUN_PROV = XVALO(VALOBADA("RELACODPROV", "PRECIO", CONDI$, "NOMESS"))
End Function
Function FECHA_COSTO_POR_PROV%(NPROV&, CI1%)
 'FECHA DE COSTO SEGUN PROVEEDOR GUARDADO EN LA TABLA RELACODPROV
   CONDI$ = "NUM_PROV = " & NPROV& & " AND COD_INTE = " & CI1%
   FECHA_COSTO_POR_PROV% = CVINT(VALOBADA("RELACODPROV", "FECHA", CONDI$, "NOMESS"))
End Function

Function CODICLIENTE&(Codigo$)
  'FUNCION CREADA PARA AHP QUE DEVUELVE EL CODIGO INTERNO PASANDOLE EL CODIGO EXTERNO DEL CLIENTE
  LONGIMAX% = XVALO(Control$("MASTER", "LOMAXCOD"))
  If LONGIMAX% < 15 Then LONGIMAX% = 15
  If Len(Codigo$) > LONGIMAX% Then
    Call COMUNI("LONGITUD EXCESIVA DE CODIGO \ Máximo Acatual: " & LONGIMAX% & " Caracteres")
    Exit Function
  End If
  '
  CODI$ = TRIM$(Codigo$)
  If CODI$ = "" Then Exit Function
  '
  CONDI$ = "Codigo_Cliente = '" & Codigo$ & "' AND STATUS >= 0"
  CODICLIENTE& = XVALO(VALOBADA("CODICLIENTE", "CODINT", CONDI$, "NOMESS"))
  '
End Function

Function SALDO_LOTE_PEDIDO(Ci%, Optional LOTE$, Optional NPED&)
   '
   Call ABRECONEXION 'stored procedure
   'SI SE LE PASA EL NUMERO DE PEDIDO, ESTE, NO SE TIENE EN CUENTA
    SQLX$ = " EXECUTE SALDOLOTEPEDIDO2  " & Ci%
    SQLX$ = SQLX$ + ", '" & LOTE$ & "'"
    SQLX$ = SQLX$ + ", " & NPED&
    SQLX$ = SQLX$ + ", " & CodiEmp%
    
    Dim rs As New ADODB.Recordset
    Set rs = FLEXCONN.Execute(SQLX$)
    SALDO_LOTE_PEDIDO = XVALO(rs!Reservado)
    rs.Close: Set rs = Nothing
   Exit Function
   
   CONDI$ = "CANPED > CantEnt AND STATUS < 8 "
   CONDI$ = CONDI$ + " AND CodiInt =  " & Ci%
   If TRIM$(LOTE$) <> "" Then CONDI$ = CONDI$ + " AND IDENLOTE = '" & LOTE$ & "'"
   '
   ''SI SE LE PASA EL NUMERO DE PEDIDO ESTE NO SE TIENE EN CUENTA
   If NPED& > 0 Then CONDI$ = CONDI$ + "  And NROPED <> " & NPED&
   '
   SQLX$ = "SELECT SUM(CANPED) AS RESERVADO, SUM(CANTENT) AS CONSUMO FROM PEDDETA WITH(NOLOCK)"
   SQLX$ = SQLX$ + " WHERE " & CONDI$
   Set rs = READSET(SQLX$)
   SALDO_LOTE_PEDIDO = XVALO(rs!Reservado) - XVALO(rs!CONSUMO)
   
   
'   RESERVADO = XVALO(VALOBADA("PEDDETA", "SUM(CANPED)", CONDI$, "NOMESS"))
'   CONSUMO = XVALO(VALOBADA("PEDDETA", "SUM(CANTENT)", CONDI$, "NOMESS"))
'   SALDO_LOTE_PEDIDO = RESERVADO - CONSUMO

End Function
Function SALDO_SIN_LOTE_PEDIDO(Ci%, RetornoStringNroPedidos$, Optional NPED&)
   'TRAE DE ESTE CODIGO SI HAY ALGUN PENDIENTE SIN LOTE ASIGNADO, PARA PRESENTAR EN UN LABEL
   'POR QUE SI NO TIENE LOTE ASIGNADO NO LO MUESTRA COMO RESERVADO EN LA PANTALLA DEL COTIZADOR CON STOCK
   
   
   Call ABRECONEXION 'stored procedure
   
   StringNroPedidos = ""
   SALDO_SIN_LOTE_PEDIDO = 0
   CONDI$ = "CANPED > CantEnt AND STATUS < 8 "
   CONDI$ = CONDI$ + " AND CodiInt =  " & Ci%
   CONDI$ = CONDI$ + " AND (IDENLOTE = '' OR IDENLOTE  IS NULL)"
   '
   ''SI SE LE PASA EL NUMERO DE PEDIDO ESTE NO SE TIENE EN CUENTA
   If NPED& > 0 Then CONDI$ = CONDI$ + "  And NROPED <> " & NPED&
   '
   cnt% = 0
   SQLX$ = "SELECT SUM(CANPED) AS RESERVADO, SUM(CANTENT) AS CONSUMO,NROPED"
   SQLX$ = SQLX$ + " FROM PEDDETA WITH(NOLOCK)"
   SQLX$ = SQLX$ + " WHERE " & CONDI$
   SQLX$ = SQLX$ + " GROUP BY NROPED"
   Set rs = READSET(SQLX$)
   With rs
     Do While Not .EOF
      SALDO_SIN_LOTE_PEDIDO = SALDO_SIN_LOTE_PEDIDO + (XVALO(rs!Reservado) - XVALO(rs!CONSUMO))
      If cnt% < 1 Then
        RetornoStringNroPedidos$ = TRIM$(CStr(XVALO(!NROPED)))
      Else
        RetornoStringNroPedidos$ = RetornoStringNroPedidos$ & "," & TRIM$(CStr(XVALO(!NROPED)))
      End If
      cnt% = cnt% + 1
      .MoveNext
    Loop
  End With
  
'   RESERVADO = XVALO(VALOBADA("PEDDETA", "SUM(CANPED)", CONDI$, "NOMESS"))
'   CONSUMO = XVALO(VALOBADA("PEDDETA", "SUM(CANTENT)", CONDI$, "NOMESS"))
'   SALDO_LOTE_PEDIDO = RESERVADO - CONSUMO

End Function
Function RESERVAS_PEDIDAS_X_VIA_PROV(Ci%, VIA%, PROV%)
   '
   Call ABRECONEXION
   'SI SE LE PASA EL NUMERO DE PEDIDO ESTE NO SE TIENE EN CUENTA
   ACUMU_ST = 0
   CONDI$ = "CANPED > CantEnt AND STATUS < 8 "
   CONDI$ = CONDI$ + " AND CodiInt =  " & Ci%
   CONDI$ = CONDI$ + " AND IDENLOTE <> ''"
   '
   SQLX$ = "SELECT SUM(CANPED) AS PEDIDA, SUM(CANTENT) AS ENTREGADA"
   SQLX$ = SQLX$ + " FROM PEDDETA WITH(NOLOCK) WHERE " & CONDI$
   Set rs = READSET(SQLX$)
   With rs
      Do While Not .EOF
        LOTE$ = SAFETEXT$(!idenlote)
        PROVEEDOR_LOTE% = PROVELOTE%(LOTE$)
        VIA_LOTE% = VIALOTE%(LOTE$)
        If PROVEEDOR_LOTE% = PROV% And VIA_LOTE% = VIA% Then
          St = PEDIDA - ENTREGADA
          ACUMU_ST = ACUMU_ST + St
        End If
        .MoveNext
      Loop
   End With

End Function

Sub CARGALOTE_PED(NUPED&)
    '
    'CARGA LA INFORMACION PARA ASIGNAR LOTES A LOS ITEMS DE UN PEDIDO.
    LOTESPED.List2.Clear
    For U& = 1 To ULTREG&("!CARDETPE")
      X$ = REGLEIDO$("!CARDETPE", U&)
      CAN = CVS(Mid$(X$, 47, 4))
      If CAN > 0.05 Then
        '
        CIXI% = CVI(Left$(X$, 2))
        If CIXI% > 0 Then
            TX$ = ""
            NUORIT% = CVI(Mid$(X$, 132, 2))
            LOTEX$ = Mid$(X$, 3, 16)
            DESCUTEX$ = TRIM$(Mid$(X$, 75, 14))
            PRELI = FORMATNUM$(CVS(Mid$(X$, 39, 4)), "F12.2")
            COEFCON = COEFDESCU(DESCUTEX$)
            NUPRESU& = CVS(Mid$(X$, 125, 4))
            FE% = CVI(Mid$(X$, 63, 2))
            '
            Call REPLA(TX$, CODEXT$(CIXI%), 1, 16) 'OJO-LONGITUD DE CODIGO EXTERNO
            Call REPLA(TX$, DESCRIT$(CIXI%), 17, 32)
            Call REPLA(TX$, FORMATNUM(CAN, "F11.1"), 50, 11)
            Call REPLA(TX$, LOTEX$, 62, 16)
            Call REPLA(TX$, FORMATNUM$(PRELI, "F12.2"), 80, 12)
            Call REPLA(TX$, SAFETEXT$(NUPRESU&), 94, 8)
            Call REPLA(TX$, FECHATEX$(FE%), 104, 8)
            Call REPLA(TX$, DESCUTEX$, 114, 14)
            LOTESPED.List2.AddItem TX$
         End If
        End If
    Next U&
    '
   'Presentar formulario para asignar lotes
   If LOTESPED.List2.ListCount > 0 Then
     LOTESPED.Label16 = NUPED&
30   LOTESPED.Show 1
     If LOTESPED.List2.ListCount < 1 Or LOTESPED.Label7 = "FALTAN_LOTES" Then
          LOTESPED.LIMPIAR
          OK% = (-1)
          Exit Sub
     Else
         'PRIMERO VALIDO QUE TODOS TENGAN LOTE
         For i& = 1 To LOTESPED.List2.ListCount
           LOTE$ = TRIM$(Mid$(LOTESPED.List2.List(i& - 1), 62, 16))
           Ci% = CODINT%(Mid$(LOTESPED.List2.List(i& - 1), 1, 16))
           If LOTE$ = "" Then
              Call COMUNI("Imposible Continuar!!\Falta Lote Para el Código: " & CODEXT(Ci%))
              GoTo 30 ' PRESENTA NUEVAMENTE EL LOTESPED
           End If
         Next i&

         J& = 0
         'SI EL ARCHIVO TENIA CODIGOS NO TRAZABLES LOS CARGA PRIMERO
         For YX& = 1 To ULTREG&("!CARDETPE")
          YY$ = REGLEIDO$("!CARDETPE", YX&)
          Ci% = CVI(Mid$(YY$, 1, 2))
          CAN = CVD(Mid$(YY$, 19, 8))
          If Ci% > 0 Then
            If Ci% <= NUMAS% Then
              If ESTRAZABLE%(Ci%) < 1 Then
                 J& = J& + 1
                 Call REPLA(YY$, MKI$(Int(J&)), 132, 2)
                 Call GRAREG("!CARDETPE", YY$, J&)
              End If
            End If
          End If
        Next YX&
        'FE% = HOY(HO$)
        '
        'RECORRE LOS LOTES TRAZABLES Y LOS GRABA A CONTINUACION EN CARDETPE.
        For i& = 1 To LOTESPED.List2.ListCount
             'CODIGO DE LIST DE LOTESPED
              Ci% = CODINT%(TRIM$(Left$(LOTESPED.List2.List(i& - 1), 16)))
              
              CANTI = XVALO(Mid$(LOTESPED.List2.List(i& - 1), 50, 11))
              LOTE$ = TRIM$(Mid$(LOTESPED.List2.List(i& - 1), 62, 16))
              PRELI = XVALO(Mid$(LOTESPED.List2.List(i& - 1), 80, 12))
              'USOII = XVALO(Mid$(LOTESPED.List2.List(I& - 1), 85, 12))
              NUPRESU& = XVALO(Mid$(LOTESPED.List2.List(i& - 1), 94, 8))
              FE% = FECHANUM(Mid$(LOTESPED.List2.List(i& - 1), 104, 8))
              DESCUTEX$ = Mid$(LOTESPED.List2.List(i& - 1), 114, 14)
              '
              J& = J& + 1
              ZZ$ = ""
              Call REPLA(ZZ$, MKI$(Ci%), 1, 2)
              Call REPLA(ZZ$, LOTE$, 3, 16)
              Call REPLA(ZZ$, MKS$(PRELI), 39, 4)
              Call REPLA(ZZ$, MKS$(CANTI), 47, 4)   ' CANTIDAD
              Call REPLA(ZZ$, MKS$(PRELI), 51, 4)
              Call REPLA(ZZ$, MKS$(PRELI), 65, 4)
              Call REPLA(ZZ$, MKI$(FE%), 63, 2)
              Call REPLA(ZZ$, DESCUTEX$, 75, 14)
              Call REPLA(ZZ$, MKS$(NUPRESU&), 125, 4)
              Call REPLA(ZZ$, MKI$(Int(J&)), 132, 2)
              '
              Call GRAREG("!CARDETPE", ZZ$, J&)
        Next i&
        '
        Call SETULTREG("!CARDETPE", J&)
     End If
   End If
   '
   LOTESPED.List2.Clear

End Sub
Function CANTICONSU(CI1%, Optional Des%, Optional Has%)

     Dim CONFECHA%
     
     CONFECHA% = 0
     If Des% > 0 And Has% > 0 Then
        CONFECHA% = 1
        DES1 = FETEXCOR1$(Des%)
        HAS1 = FETEXCOR1$(Has%)
     End If
     '
     CANTICONSU = 0
     SQLX$ = " SELECT SUM(CANTINGRE) AS INGRE1, SUM(CANTSALID) AS SALI1   FROM MOVISTO WITH(NOLOCK) "
     SQLX$ = SQLX$ + " WHERE COD_INTE = " & CI1%
     SQLX$ = SQLX$ + " AND (CODIMOVI = 'RT' OR CODIMOVI = 'CF' OR  CODIMOVI = 'TR' OR  CODIMOVI = 'VM')"
     If CONFECHA% > 0 Then
        SQLX$ = SQLX$ + " AND FECHMOV >= '" & DES1 & "' "
        SQLX$ = SQLX$ + " AND FECHMOV <= '" & HAS1 & "' "
     End If
     SQLX$ = SQLX$ + " GROUP BY COD_INTE"

     Set rs = READSET(SQLX$)
     With rs
       Do While Not .EOF
           CANTI = XVALO(!SALI1) - XVALO(!INGRE1)
           If CANTI < 0.05 Then CANTI = 0
           .MoveNext
       Loop
     End With
     '
     rs.Close
     Set rs = Nothing
     '
     CANTICONSU = CANTI
     '
End Function

Function CANTIRECIBIDA(CI1%, Optional Des%, Optional Has%)

     Dim CONFECHA%
     
     CONFECHA% = 0
     If Des% > 0 And Has% > 0 Then
        CONFECHA% = 1
        DES1 = FETEXCOR1$(Des%)
        HAS1 = FETEXCOR1$(Has%)
     End If
     '
     CANTI = 0
     ACUMULACANTI = 0
     SQLX$ = " SELECT SUM(CanToApro) AS APRO, SUM(CanToRech) AS RECHA ,SUM(CanToFalta) AS FALTA  FROM BOLRECEP WITH(NOLOCK) "
     SQLX$ = SQLX$ + " WHERE COD_INTE = " & CI1%
     If CONFECHA% > 0 Then
        SQLX$ = SQLX$ + " AND FechRecep >= '" & DES1 & "' "
        SQLX$ = SQLX$ + " AND FechRecep <= '" & HAS1 & "' "
     End If
     SQLX$ = SQLX$ + " GROUP BY COD_INTE"

     Set rs = READSET(SQLX$)
     With rs
       Do While Not .EOF
           CANTI = XVALO(!APRO) - (XVALO(!RECHA) + XVALO(!FALTA))
           If CANTI < 0.05 Then CANTI = 0
         .MoveNext
       Loop
     End With
     rs.Close
     Set rs = Nothing
     AA = CODEXT$(CI1%)
     '
     CANTIRECIBIDA = CANTI
     '
End Function


Function CANTICOMPRADA(CI1%, TOTALCOMPRA#, Optional Des%, Optional Has%)

     Dim CONFECHA%
    
     CONFECHA% = 0
     If Des% > 0 And Has% > 0 Then
        CONFECHA% = 1
        DES1 = FETEXCOR1$(Des%)
        HAS1 = FETEXCOR1$(Has%)
     End If
     
'     Funcion que me devuelva la cantidad comprada
'    fecha >= desde and fecha <= hasta.
'    statcom < 8
     ACUMULACANTI = 0
     SQLX$ = "Select  COD_INTE, UNIDS_ocom,  PrUn_Neto_Pes from OCOMDETA WITH(NOLOCK)"
     SQLX$ = SQLX$ + " Where Stat_Ocom < 8 "
     SQLX$ = SQLX$ + " AND  Stat_Ocom >= 0 " 'EXCLUYE LAS O/C's NO APROB.
     SQLX$ = SQLX$ + " AND COD_INTE = " & CI1%
     If CONFECHA% > 0 Then
        SQLX$ = SQLX$ + " AND FEMI_OCOM >= '" & DES1 & "'"
        SQLX$ = SQLX$ + " AND FEMI_OCOM <= '" & HAS1 & "'"
     End If
     SQLX$ = SQLX$ + " ORDER BY NUME_OCOM"
     Set rs = READSET(SQLX$)
     With rs
       Do While Not .EOF
         CANTI = XVALO(!Unids_Ocom)
         PREUNIX# = XVALO(!PrUn_Neto_Pes)
         Subtotal# = CANTI * PREUNIX#
         AcumTotal# = AcumTotal# + Subtotal#
         ACUMULACANTI = ACUMULACANTI + CANTI
         .MoveNext
       Loop
     End With
     '
     rs.Close
     Set rs = Nothing

     CANTICOMPRADA = 0
     TOTALCOMPRA# = 0
     If ACUMULACANTI > 0.005 Then
        TOTALCOMPRA# = AcumTotal#
        CANTICOMPRADA = ACUMULACANTI
     End If
     '
End Function

Function CompraPendientesAcumuladasPos(CI1%, Optional NUOC&, Optional SIN_PEDIDO%, Optional FECALCULOX, Optional INCLU_ANUL%)
     'SI FECALCULO% > 0 CONSIDERA QUE LAS ORDENES DE COMPRA SEAN POSTERIOR AL CALCULO DEL MRP
     FECALCULOX = TRIM$(FECALCULOX)
     FECHYHO$ = FECALCULOX
15   Call ABRECOMPRAS
     CompraPendientesAcumuladasPos = 0
     
     SQLX$ = " SELECT  D.Coef_Unids, D.cant_ocom, D.Cant_Rec,D.NUME_OCOM FROM OCOMDETA D WITH(NOLOCK)  "
     SQLX$ = SQLX$ + " INNER JOIN OCOMENCA O ON O.NUME_OCOM = D.NUME_OCOM "
     SQLX$ = SQLX$ + " WHERE D.Cod_Inte = " & CI1% & " "
     SQLX$ = SQLX$ + " AND D.Stat_Ocom < 8 "   ' para que no este anulada ni cumplida
     If INCLU_ANUL% < 1 Then SQLX$ = SQLX$ + " AND  D.Stat_Ocom <> -1 " 'Para que no  hayan O/C's sin aprobar
     SQLX$ = SQLX$ + " AND D.Cant_Rec < (D.Cant_Ocom - 0.005) "
     SQLX$ = SQLX$ + " AND D.N_ORSERV < 1 "
     If NUOC& > 0 Then SQLX$ = SQLX$ + "AND D.NUME_OCOM <> " & NUOC&
     asa = FETEXCOR3$(CVINT(FECALCULOX), FECHYHO$)
     If FECALCULOX <> "" Then SQLX$ = SQLX$ + " AND  O.FGen_Ocom > '" & FETEXCOR3$(CVINT(FECALCULOX), FECHYHO$) & "'"
     SQLX$ = SQLX$ + "ORDER BY Nume_Ocom ASC "
     '
     Set OcoTemp = READSET(SQLX$)
     With OcoTemp
         Do While Not .EOF
             COEFCAN = XVALO(!Coef_Unids)
             If COEFCAN < 0.005 Then COEFCAN = 1
             CACOM = XVALO(!cant_ocom) * COEFCAN
             CAREC = XVALO(!Cant_Rec) * COEFCAN
             CAPEN = CACOM - CAREC
             If SIN_PEDIDO% > 0 Then
               CONDI$ = "NUME_CLI > 0 AND  NUME_OCOM = " & XVALO(!nume_ocom)
               If CantRegistros&("OCOMENCA", CONDI$, "NOMESS") > 0 Then GoTo 50
             End If
             CompraPendientesAcumuladasPos = CompraPendientesAcumuladasPos + CAPEN
50           .MoveNext
         Loop
     End With
     OcoTemp.Close
     Set OcoTemp = Nothing
End Function
Function CompraPendientesAcumuladas(CI1%, Optional NUOC&, Optional SIN_PEDIDO%, Optional FECALCULOX)
     
15   Call ABRECOMPRAS
     CompraPendientesAcumuladas = 0
     
     SQLX$ = " SELECT  Coef_Unids, cant_ocom, Cant_Rec,NUME_OCOM FROM OCOMDETA WITH(NOLOCK) "
     SQLX$ = SQLX$ + "WHERE Cod_Inte = " & CI1% & " "
     SQLX$ = SQLX$ + "AND Stat_Ocom < 8 "   ' para que no este anulada ni cumplida
     SQLX$ = SQLX$ + "AND  Stat_Ocom <> -1 " 'Para que no  hayan O/C's sin aprobar
     SQLX$ = SQLX$ + "AND Cant_Rec < (Cant_Ocom - 0.005) "
     SQLX$ = SQLX$ + "AND N_ORSERV < 1 "
     If NUOC& > 0 Then SQLX$ = SQLX$ + "AND NUME_OCOM <> " & NUOC&
     SQLX$ = SQLX$ + "ORDER BY Nume_Ocom ASC "
     '
     Set OcoTemp = READSET(SQLX$)
     With OcoTemp
         Do While Not .EOF
             COEFCAN = XVALO(!Coef_Unids)
             If COEFCAN < 0.005 Then COEFCAN = 1
             CACOM = XVALO(!cant_ocom) * COEFCAN
             CAREC = XVALO(!Cant_Rec) * COEFCAN
             CAPEN = CACOM - CAREC
             If SIN_PEDIDO% > 0 Then
               CONDI$ = "NUME_CLI > 0 AND  NUME_OCOM = " & XVALO(!nume_ocom)
               If CantRegistros&("OCOMENCA", CONDI$, "NOMESS") > 0 Then GoTo 50
             End If
             CompraPendientesAcumuladas = CompraPendientesAcumuladas + CAPEN
50           .MoveNext
         Loop
     End With
     
End Function

Function FabricacionPendiente(CI1%)
     Call ABREORFAB
     TOFABRI = 0
     'TOTAL DE CANTIDAD PENDIENTE EN ORDENES DE FABRICACION PENDIENTES DE CIERRE.
     SQLX$ = " SELECT CanProy,  CanApro,  NUMEORFA, FechEnSol, Referen, FechProyEntre, FechRecal FROM ORDEFABR WITH(NOLOCK) "
     SQLX$ = SQLX$ + "WHERE Cod_Inte = " & CI1% & " "
     SQLX$ = SQLX$ + "AND StatuOrf < 3 "   ' para que no este anulada/cumplida/cerrada
     SQLX$ = SQLX$ + "AND CanApro < (CanProy - 0.005) "
     SQLX$ = SQLX$ + "ORDER BY NumeOrFa ASC "
     '
     Set OfaTemp = READSET(SQLX$)
     With OfaTemp
         Do While Not .EOF
           CACOM = XVALO(!CANPROY)
           CAREC = XVALO(!CanApro)
           CAPEN = CACOM - CAREC
           If CAPEN > 0 Then
             TOFABRI = TOFABRI + CAPEN
           End If
         .MoveNext
         Loop
     End With
     '
     FabricacionPendiente = TOFABRI
End Function

Function VIAENTRE_COMPRA(NUMERO_ORD_COMPRA&)
   'DEVUELVE LA VIA CARGADA QUE SE CARGO EN LA ORDEN DE COMPRA.
   VIAENTRE_COMPRA = XVALO(VALOBADA("OCOMENCA", "VIA_ENTREGA", "NUME_OCOM = " & NUMERO_ORD_COMPRA&, "NOMESS"))
End Function
Function CANFAPEN(Ci%) 'DEVUELVE LA CANTIDAD TOTAL EN PROCESO DE FABRICACIÓN DE UN ARTÍCULO
'\\\OT-07-0386-FL-17/10/07///
    CANFA# = 0
    If Ci% > 0 Then
        Call APERBASDAT("ORFAB")
        SQLX$ = "SELECT * FROM ORDEFABR WITH(NOLOCK)"
        SQLX$ = SQLX$ + "WHERE CANAPRO < CANPROY "
        SQLX$ = SQLX$ + "AND COD_INTE = " & Ci% & " "
        SQLX$ = SQLX$ + "AND StatuOrf < 3 "
        Set OrdefabrTemp = READSET(SQLX$)
        '
        With OrdefabrTemp
             Do While Not .EOF
               CANFA# = CANFA# + (XVALO(!CANPROY) - XVALO(!CanApro))
             .MoveNext
             Loop
        End With
    End If
    CANFAPEN = CANFA#
    '\\\OT-07-0386-FIN///
End Function
Function CANOCPEN(CI1%) ' DEVUELVE LAS CANTIDADES COMPRADAS PENDIENTES DE RECEPCIÓN
    ' '\\\OT-07-0386-FL-16/10/07///
    CANPEN# = 0
    If CI1% > 0 Then
        Call APERBASDAT("COMPRAS")
        SQLX$ = "SELECT *  FROM OCOMDETA WITH(NOLOCK)"
        SQLX$ = SQLX$ + "WHERE CANT_OCOM > CANT_REC "
        SQLX$ = SQLX$ + "AND COD_INTE = " & CI1% & " "
        SQLX$ = SQLX$ + "AND STAT_OCOM < 3 "
        Set OcomdetaTemp = READSET(SQLX$)
        '
        With OcomdetaTemp
             Do While Not .EOF
               CANPEN# = CANPEN# + (XVALO(!cant_ocom) - XVALO(!Cant_Rec))
             .MoveNext
             Loop
        End With
    End If
    CANOCPEN = CANPEN#
   ' '\\\OT-07-0386-FIN///
End Function
Function CANRECEP(CI1%, Optional FECALCU)
    Call ABRESTOCKS
    'SI VIENE EL PARAMETRO DE LA FECHA DE CALCULO DEL MRP AGREGA LA CONDICION QUE LA FECHA SEA POSTERIOR
    FECALCU = TRIM$(FECALCU)
    FECHYHO$ = FECALCU
    FECALCU1 = FETEXCOR3$(CVINT(FECALCU), FECHYHO$)
    SQLX$ = " SELECT FEYHORecep, CANTNOM,CanToApro,FechRecep  FROM BOLRECEP WITH(NOLOCK)"
    SQLX$ = SQLX$ + " WHERE Status = 0 "
    SQLX$ = SQLX$ + " AND Cod_Inte = " & CI1% & " "
    '
    Set COBRECEPTMP = READSET(SQLX$)
    '
    TORECEPE = 0
    With COBRECEPTMP
         JJ& = 0
         k& = 0
         Do While Not .EOF
            If FECALCU <> "" Then
              If IsDate(SAFETEXT$(!FEYHORecep)) Then
                 FEYHOR = TRIM$(SAFETEXT$(!FEYHORecep))
                 FECALCBD = FETEXCOR3$(CVINT(FEYHOR), SAFETEXT$(FEYHOR))
                 If FECALCBD <= FECALCU1 Then GoTo 50
              Else
                 If SAFETEXT$(!FECHRECEP) <= FECALCU Then GoTo 50
              End If
            End If
            CARECEP = XVALO(!CantNom) - XVALO(!CANTOAPRO)
            If CARECEP > 0 Then
               TORECEPE = TORECEPE + CARECEP
            End If
50       .MoveNext
         Loop
    End With
    CANRECEP = TORECEPE
    'NOTA: AGREGO DENTRO DEL DO WHILE EL FILTRO DE LA FECHA DE CALCULO PARA QUE SI EL CAMPO FEYHORecep TIENE VALOR FECHA
    'COMPARE CON ESTE SI NO COMPARE CON EL FECHRECEP, EL MOTIVO ES POR QUE EL CAMPO FEYHORECEP ES NUEVO .
End Function

Function CANTIDAD_PEDIDA_PENDIENTE(CI1%)

     Call ABREPEDCLI
     TOPEDID = 0
     SQLX$ = " SELECT NROPED, NroClie, CanPed,  CantEnt, FechEmis, AcuDespa, AcuPrepa, FeSolEnt FROM PEDDETA WITH(NOLOCK) "
     SQLX$ = SQLX$ + "WHERE CodiInt = " & CI1% & " "
     SQLX$ = SQLX$ + "AND VenRep = 'V' "
     SQLX$ = SQLX$ + "AND Status < 8 "   ' para que no este anulada ni cumplida
     SQLX$ = SQLX$ + "AND CantEnt < (CanPed - 0.005) "
'     If CONSIDERAR_DESPACHO% > 0 Then
'        SQLX$ = SQLX$ + "AND STADESPA > 0" 'QUE NO ESTE INFORMADO COMO DESPACHADO
'     End If
     SQLX$ = SQLX$ + "ORDER BY NroPed ASC "
     '
     'ATENCION NO REALIZO UN SUM DE LAS CANTIDADES POR SI EN PEDIDOS HAY ENTREGAS MAYORES A LA CANTIDAD SOLICITADA
     Set PEDTEMP = READSET(SQLX$)
     With PEDTEMP
         Do While Not .EOF
           CAPEN = XVALO(!CANPED) - XVALO(!CANTENT)
           If CAPEN < 0 Then CAPEN = 0
           '
           If CAPEN > 0 Then
             TOPEDID = TOPEDID + CAPEN
           End If
           .MoveNext
         Loop
     End With
     CANTIDAD_PEDIDA_PENDIENTE = TOPEDID

End Function

Function CODISECC#(Ci%)
    '
    CODISECC# = 0
    If ExisteArchivo%(LUDAT$, "CODSEPRO.RFS") = 1 Then
        RFF$ = FILTERGETRECORD$("CODSEPRO", 1, MKI$(Ci%))
        CODISECC# = CVS(Mid$(RFF$, 3, 4))
    End If
    '
End Function
Function PRECIODEVENTA_AHP(CI1%, CTAPROV, VIAENTR%)
     'procedimiento almacenado
     PRECIODEVENTA_AHP = 0
     SQLX$ = " EXECUTE PRECIOVENTAAHP " & XVALO(CTAPROV)
     SQLX$ = SQLX$ + ", " & CodiEmp%
     SQLX$ = SQLX$ + ", " & CI1%
     SQLX$ = SQLX$ + ", " & VIAENTR%
     Set rs = FLEXCONN.Execute(SQLX$)
     PREXVENTAX = XVALO(rs.Fields(0))
     If PREXVENTAX > 0 Then
      PRECIODEVENTA_AHP = PREXVENTAX
     End If

'     SQLX$ = "SELECT PRECIOVENTA FROM COSTOXPROVE WITH(NOLOCK)"
'     SQLX$ = SQLX$ + " WHERE CODINT = " & CI1%
'     SQLX$ = SQLX$ + " AND Nume_Cli =" & CTAPROV
'     SQLX$ = SQLX$ + " AND Via_Entrega = " & VIAENTR%
'     Set RS = READSET(SQLX$)
'     If Not RS.EOF Then PRECIODEVENTA_AHP = XVALO(RS!PRECIOVENTA)
'     RS.Close
'     Set RS = Nothing

End Function
Sub CARUBICSQL(CIII%, NPRO)
    '
    Dim REGUBI&(512)
    If CIII% <= 0 Then
        Call MENSERR(24, "Código de Artículo no Válido")
        Exit Sub
    End If
    '
'    If NPRO < 1 Then
''      Call Label13_DblClick
'      Exit Sub
'    End If
    '
'    If NPRO <= 0 Then
'        Call MENSERR(24, "Proveedor no Válido")
'        Exit Sub
'    End If
    '
    Screen.MousePointer = 11
    Call BLANARCH("!CARUBIC")
    Call SETULTREG("!CARUBIC", 0)
    '
    K1& = 0
    SQLX$ = "SELECT * FROM UBICASQL  WITH(NOLOCK) "
    SQLX$ = SQLX$ + " WHERE COD_INTE = " & CIII%
    If NPRO > 0 Then SQLX$ = SQLX$ + " AND Nume_Pro = " & NPRO
    SQLX$ = SQLX$ + " ORDER BY CODIUBICA"
    Set rs = READSET(SQLX$)
    With rs
      Do While Not .EOF
         CODUB$ = TRIM$(SAFETEXT$(!CODIUBICA))
         Observa$ = SAFETEXT$(!OBSERVACION)
         If CODUB$ <> "" Then
            Call REPLA(TX$, MKI$(CIII%), 1, 2)
            Call REPLA(TX$, CODUB$, 3, 24)
            Call REPLA(TX$, MKS$(0), 27, 4)
            Call REPLA(TX$, Observa$, 31, 30)
            Call REPLA(TX$, MKS$(NPRO), 61, 4)
            Call REGAPP("!CARUBIC", TX$)
            K1& = K1& + 1: REGUBI&(K1&) = REG&
         End If
         .MoveNext
      Loop
    End With
    rs.Close
    Set rs = Nothing
    
    Screen.MousePointer = 1
    '
    VTB% = VENTABU%("CARUBIC/BORDEIZQ=20/BORDESUP=12/NC/TITUPAN=" + TRIM$(CODEXT$(CIII%)) + " - " + DESCRIT$(CIII%))
    '
    If VTB% >= 0 Then
        CONDI$ = "COD_INTE = " & CIII%
        If NPRO > 0 Then CONDI$ = CONDI$ + " AND Nume_Pro = " & NPRO
        '
        Call ACTUREGISTRO("UBICASQL", "MARBORRA", CONDI$, 1, REGAF&, "NOMESS")
        
        Screen.MousePointer = 11
        k& = ULTREG&("!CARUBIC")
        K2& = 0
        For i& = 1 To k&
            TX$ = REGLEIDO$("!CARUBIC", i&)
            If TRIM$(Mid$(TX$, 3, 24)) <> "" Then
                '
                CODUB$ = TRIM$(Mid$(TX$, 3, 24))
                Observa$ = TRIM$(Mid$(TX$, 31, 30)) 'OJO ESTA POSICION ME PARECE QUE ESTRA MAL DEBERIA SER 31,30
                Call INSERTUBICASQL(CIII%, CODUB$, Observa$, NPRO)
            End If
            '
        Next i&
        '
        CONDI$ = CONDI$ + " AND MARBORRA = 1"
        Call BORRAREGISTROS("UBICASQL", CONDI$, REGAF&, "NOMESS")
        '
       
    '
    End If
    Call CIERRARCH("UBISTOCK")
    Screen.MousePointer = 1
    '
End Sub
Sub INSERTUBICASQL(CI1%, CODIUBICA$, OBSERVACION$, NumeProv)
        SQLX$ = "INSERT INTO UBICASQL"
        SQLX$ = SQLX$ + " ([Cod_Inte], [CodiEmpr], [CODIUBICA], [CANTIDAD], [OBSERVACION], [Nume_Pro], [MARBORRA])"
        SQLX$ = SQLX$ + " VALUES(" & CI1%
        SQLX$ = SQLX$ + ", " & CodiEmp%
        SQLX$ = SQLX$ + ", '" & CODIUBICA$ & "'"
        SQLX$ = SQLX$ + ", 0"
        SQLX$ = SQLX$ + ", '" & OBSERVACION$ & "'"
        SQLX$ = SQLX$ + ", " & NumeProv
        SQLX$ = SQLX$ + ", 0"
        SQLX$ = SQLX$ + ")"
        FLEXCONN.Execute (SQLX$)
   
End Sub

Function COSTO_SEGUN_COMPRA(CI1%, NUOC&, Optional NUORDEN%)
   COSTO_SEGUN_COMPRA = 0
   SQLX$ = "SELECT PrUn_Neto FROM OCOMDETA WITH(NOLOCK)"
   SQLX$ = SQLX$ + " WHERE STAT_OCOM < 8"
   SQLX$ = SQLX$ + " AND COD_INTE = " & CI1%
   If NUORDEN% > 0 Then SQLX$ = SQLX$ + " AND NuOrd_Item = " & NUORDEN%
   SQLX$ = SQLX$ + " AND NUME_OCOM = " & NUOC&
   Set rs = READSET(SQLX$)
   With rs
      COSTO_SEGUN_COMPRA = XVALO(!prun_neto)
   End With
   rs.Close
   Set rs = Nothing
End Function
Sub RUTINA_ACTU_COSTOPROVE(NPROV, CodiEmp%, CI1%, VIAENTREX%, prexiventa#)
   SQLX$ = " " & NPROV
   SQLX$ = SQLX$ + " ," & CodiEmp%
   SQLX$ = SQLX$ + " ," & CI1%
   SQLX$ = SQLX$ + " ," & VIAENTREX%
   SQLX$ = SQLX$ + " ," & FORMATNUM$(prexiventa#, "F16.4")
   'EL ULTIMO PARAMETRO ES EL DE LA LISTA DE PREWCIOS SI NO SE PASA POR DEFECTO EN EL SP LO GRABA CON 1
   FLEXCONN.Execute "ACTU_COSTOPROVE_2 " & SQLX$
End Sub
Function COEFI_COMPRA_SIEMPRE_DOLAR(NUOC&)
   'PARA QUE SIEMPRE DEVUELVA PARA CALCULAR LOS DOLARES
   'SI ES MONEDA > 1 (DOLAR) DEVUELVE 1 PARA QUE MANTENGA EL VALOR DE LA MONEDA EXTRANJERA
   'SI ES MONEDA 1 (PESOS) DEVUELVE EL VALOR DE COTIZA DOLAR PARA QUE LO TRANSFORME LOS PESOS A DOLARES AL MOMENTO DE COMPRA
   COEFI_COMPRA_SIEMPRE_DOLAR = 1 ' POR DEFECTO COEFICIENTE = 1
   SQLX$ = "SELECT Mone_Emis,COTIZA_DOLAR FROM OCOMENCA WITH(NOLOCK)"
   SQLX$ = SQLX$ + " WHERE Nume_Ocom = " & NUOC&
   Set rs = READSET(SQLX$)
   MONEX% = XVALO(rs!Mone_Emis)
   If MONEX% > 1 Then 'SI MONEDA DE COMPRA ES > 1 TOMA LA COTIZACION DEL DOLAR
      COEFI_COMPRA_SIEMPRE_DOLAR = 1
   Else
      COEFI_COMPRA_SIEMPRE_DOLAR = XVALO(rs!COTIZA_DOLAR)
   End If
   rs.Close
   Set rs = Nothing
End Function

Function ProveedItemSegunPedido&(NUPED&, CI1%, NUORD%)
   ProveedItemSegunPedido& = 0
   SQLX$ = "SELECT Num_Prov FROM PEDDETA WITH(NOLOCK)"
   SQLX$ = SQLX$ + " WHERE NroPed = " & NUPED&
   SQLX$ = SQLX$ + " AND CODIINT = " & CI1%
   SQLX$ = SQLX$ + " AND NuorItem = " & NUORD%
   Set rs = READSET(SQLX$)
   With rs
     If Not .EOF Then
       ProveedItemSegunPedido& = XVALO(!Num_Prov)
     End If
   End With
   rs.Close
   Set rs = Nothing
End Function

Function RECE_POST_FECHA(CI1%, FECHYHORA)
    RECE_POST_FECHA = 0
    FECHYH$ = TRIM$(FECHYHORA)
    FECHYHORAX = FETEXCOR3$(CVINT(FECHYHORA), FECHYH$)
'     FECHYHORAX = FETEXCOR1$(CVINT(FECHYHORA))'le quite la hora por que en ambest daba error devolvia '05/dic/2016 610:51:'
'    If Len(FECHYHORA) > 11 Then B = Mid$(FECHYHORA, 11)
'    FECHYHORA = A & B
    If IsDate(FECHYHORA) = False Then Exit Function
    'DEVUELVE LAS RECEPCIONES APROBADAS POSTERIORES A UNA FECHA DE UN ARTICULO
    SQLX$ = "select SUM(CANTINGRE) AS INGRE, SUM(CANTSALID) AS SALI FROM MOVISTO WITH(NOLOCK) WHERE  COD_INTE = " & CI1%
    SQLX$ = SQLX$ + " AND COD_EXTE <> '' AND CODIMOVI = 'BR' AND FEREMOVI > '" & FECHYHORAX & "'"
    Set rs = READSET(SQLX$)
    If Not rs.EOF Then RECE_POST_FECHA = XVALO(rs!INGRE) - XVALO(rs!sali)
    
    rs.Close
    Set rs = Nothing
End Function
Function CIERRE_POST_FECHA_MRP(CI1%, FECHGENOF%)
    'DEVUELVE LO INGRESADO COMO ORDEN DE FABRAICACION AL MOVISTO POSTERIOR A UNA FECHA (QUE EN ESTE CASO ES EL DE GENERACION DE LA OF)
    CIERRE_POST_FECHA_MRP = 0
    FECHAX = FETEXCOR1$(FECHGENOF%)
    If IsDate(FECHAX) = False Then Exit Function
    'DEVUELVE LAS RECEPCIONES APROBADAS POSTERIORES A UNA FECHA DE UN ARTICULO
    SQLX$ = "select SUM(CANTINGRE) AS INGRE  FROM MOVISTO WITH(NOLOCK) WHERE  COD_INTE = " & CI1%
    SQLX$ = SQLX$ + " AND CODIMOVI = 'OF' AND FEREMOVI > '" & FECHAX & "'"
    Set rs = READSET(SQLX$)
    If Not rs.EOF Then CIERRE_POST_FECHA_MRP = XVALO(rs!INGRE)
    
    rs.Close
    Set rs = Nothing
End Function

Function FECHAULTIMO_CALCULO()
  'DEVUELVE FECHA DE ULTIMO CALCULO SEGUN TABLA OCOEMIHIS
  SQLX$ = "SELECT FECALCULO FROM COBERFA WITH(NOLOCK) ORDER BY FECALCULO DESC"
  Set rs = READSET(SQLX$)
  FECHAULTIMO_CALCULO = ""
  If Not rs.EOF Then FECHAULTIMO_CALCULO = SAFETEXT$(rs!FECALCULO)
  rs.Close
  Set rs = Nothing

End Function
 Function NECECOMPRA(CI1, Optional Semana%)
   SQLX2$ = "": NECECOMPRA = 0
   'SEMANA% = 1 DEVUELVE NECESIDAD DE LA SEMANA ACTIVA HACIA ATRAS
   'SEMANA% = 2 DEVUELVE NECESIDAD DE LA SEGUNDA SEMANA POSTERIOR A LA ACTIVA
   'SEMANA% = 0 DEVUELVE NECESIDAD TOTAL SIN CONSIDERAR SEMANAS.
   '
   'CALCULA QUE DIA DE LA SEMANA ES DIASEM DEVUELVE DE 0 A 6 - 0 ES DOMINGO
   DS = DIASEM%(HOY(HO$))
   CADI% = 6 - DS
   DSPOS% = DIASPOST(HOY(HO$), CADI%) 'BUSCA QUE DIA TERMINA LA SEMANA QUE VA A SER EL SABADO
   If Semana% = 1 Then
      FESEMACTIVA = FETEXCOR1$(DSPOS%)
      SQLX2$ = " AND fechrecep  <= '" & FESEMACTIVA & "'"
   ElseIf Semana% = 2 Then  'SI ES LA SEGUNDA SEMANA CONSULTA POR EL RANGO DESDE EL PRINCIPIO DE LA SEGUNDA SEMANA A LA ACTIVA - HASTA EL FIN DE ESTA.
      FESEMADDE = FETEXCOR1$(DSPOS% + 1)
      FESEMAHST = FETEXCOR1$(DSPOS% + 7)
      SQLX2$ = " AND fechrecep  >= '" & FESEMADDE & "'"
      SQLX2$ = SQLX2$ + " AND fechrecep  <= '" & FESEMAHST & "'"
   End If
   '
   SQLX$ = "SELECT CANTIDAD AS CANTI, PERIODO FROM OCOEMI WITH(NOLOCK) WHERE COD_INTE = " & CI1
   SQLX$ = SQLX$ + SQLX2$ + " ORDER BY fechrecep "
   Set rs = READSET(SQLX$)
   Do While Not rs.EOF
     NECECOMPRA = NECECOMPRA + XVALO(rs!CANTI) ' ACUMULA
     PERIOX$ = SAFETEXT$(rs!PERIODO)
     rs.MoveNext
   Loop
   rs.Close
   Set rs = Nothing

End Function
 Function NECECOMPRA_MRP(CI1, Optional Semana%)
 
   SQLX2$ = "": NECECOMPRA_MRP = 0
   'SEMANA% = 1 DEVUELVE NECESIDAD DE LA SEMANA ACTIVA HACIA ATRAS
   'SEMANA% = 2 DEVUELVE NECESIDAD DE LA SEGUNDA SEMANA POSTERIOR A LA ACTIVA
   'SEMANA% = 0 DEVUELVE NECESIDAD TOTAL SIN CONSIDERAR SEMANAS.
   '
   'CALCULA QUE DIA DE LA SEMANA ES DIASEM DEVUELVE DE 0 A 6 - 0 ES DOMINGO
   DS = DIASEM%(HOY(HO$))
   CADI% = 6 - DS
   DSPOS% = DIASPOST(HOY(HO$), CADI%) 'BUSCA QUE DIA TERMINA LA SEMANA QUE VA A SER EL SABADO
   If Semana% = 1 Then
      FESEMACTIVA = FETEXCOR1$(DSPOS%)
      SQLX2$ = " AND FecNecesi  <= '" & FESEMACTIVA & "'"
   ElseIf Semana% = 2 Then  'SI ES LA SEGUNDA SEMANA CONSULTA POR EL RANGO DESDE EL PRINCIPIO DE LA SEGUNDA SEMANA A LA ACTIVA - HASTA EL FIN DE ESTA.
      FESEMADDE = FETEXCOR1$(DSPOS% + 1)
      FESEMAHST = FETEXCOR1$(DSPOS% + 7)
      SQLX2$ = " AND FecNecesi  >= '" & FESEMADDE & "'"
      SQLX2$ = SQLX2$ + " AND FecNecesi  <= '" & FESEMAHST & "'"
   End If
   '
   SQLX$ = "SELECT TOP 1 ID FROM COBERFA WITH(NOLOCK) ORDER BY ID DESC"
   Set RS0 = READSET(SQLX$): IDMAX& = XVALO(RS0!ID): If IDMAX& < 1 Then Exit Function
   SQLX$ = " select CANTIDISP from coberfa WITH(NOLOCK) "
   SQLX$ = SQLX$ + " Where COD_INTE = " & CI1
   SQLX$ = SQLX$ + "  AND TIPOMOVI = 14 " ' COMPRA
   SQLX$ = SQLX$ + "  AND ID = " & IDMAX&
   SQLX$ = SQLX$ + SQLX2$
   Set rs = READSET(SQLX$)
   Do While Not rs.EOF
     NECECOMPRA_MRP = NECECOMPRA_MRP + XVALO(rs!CANTIDISP) ' ACUMULA
     rs.MoveNext
   Loop
   rs.Close
   Set rs = Nothing


End Function



 Function NECEFABRICAR(IDSUB$, Optional Semana%)
 
   SQLX2$ = "": NECEFABRICAR = 0
   'SEMANA% = 1 DEVUELVE NECESIDAD DE LA SEMANA ACTIVA HACIA ATRAS
   'SEMANA% = 2 DEVUELVE NECESIDAD DE LA SEGUNDA SEMANA POSTERIOR A LA ACTIVA
   'SEMANA% = 0 DEVUELVE NECESIDAD TOTAL SIN CONSIDERAR SEMANAS.
   '
   'CALCULA QUE DIA DE LA SEMANA ES DIASEM DEVUELVE DE 0 A 6 - 0 ES DOMINGO
   
   DS = DIASEM%(HOY(HO$))
   CADI% = 6 - DS
   DSPOS% = DIASPOST(HOY(HO$), CADI%) 'BUSCA QUE DIA TERMINA LA SEMANA QUE VA A SER EL SABADO
   If Semana% = 1 Then
      FESEMACTIVA = FETEXCOR1$(DSPOS%)
      SQLX2$ = " AND FECHNECE  <= '" & FESEMACTIVA & "'"
   ElseIf Semana% = 2 Then  'SI ES LA SEGUNDA SEMANA CONSULTA POR EL RANGO DESDE EL PRINCIPIO DE LA SEGUNDA SEMANA A LA ACTIVA - HASTA EL FIN DE ESTA.
      FESEMADDE = FETEXCOR1$(DSPOS% + 1)
      FESEMAHST = FETEXCOR1$(DSPOS% + 7)
      SQLX2$ = " AND FECHNECE  >= '" & FESEMADDE & "'"
      SQLX2$ = SQLX2$ + " AND FECHNECE  <= '" & FESEMAHST & "'"
   End If
   '
   SQLX$ = "SELECT CANT_PROG AS CANTI, PERIODO FROM PLAFABR  WITH(NOLOCK) WHERE COD_INTE = " & CI1
   SQLX$ = SQLX$ + SQLX2$ + " ORDER BY FECHNECE "
   Set rs = READSET(SQLX$)
   Do While Not rs.EOF
     NECEFABRICAR = NECEFABRICAR + XVALO(rs!CANTI) ' ACUMULA
     PERIOX$ = SAFETEXT$(rs!PERIO_PROGRAM)
     rs.MoveNext
   Loop
   rs.Close
   Set rs = Nothing


End Function
 Function NECEFABRICAR_MRP(CI1, IDSUBOX$, Optional FECHGEN$, Optional Semana%)
   'SEMANA% = 1 DEVUELVE NECESIDAD DE LA FECHA DE FABRICACION  HACIA ATRAS
   'SEMANA% = 2 DEVUELVE NECESIDAD DE LA FECHA POSTERIOR A LA ACTIVA
   'SEMANA% = 0 DEVUELVE NECESIDAD TOTAL SIN CONSIDERAR SEMANAS.
   '
   'CALCULA QUE DIA DE LA SEMANA ES DIASEM DEVUELVE DE 0 A 6 - 0 ES DOMINGO
   FECHGEN$ = TRIM$(FECHGEN$)
   If FECHGEN$ = "" Then FECHGEN$ = VALOBADA("ORDEFABR", "FECHGEN", "IDSUBOR = '" & IDSUBOX$ & "'", "NOMESS")
   FECHG% = CVINT(FECHGEN$)

   CONDI$ = " FECALCULO < '" & FETEXCOR3$(FECHG%, FECHGEN$) & "' ORDER BY FECALCULO DESC"
   ID1 = XVALO(VALOBADA("COBERFA", "ID", CONDI$, "NOMESS"))
   DS = DIASEM%(HOY(HO$))
   CADI% = 6 - DS
   DSPOS% = DIASPOST(FECHG%, CADI%) 'BUSCA QUE DIA TERMINA LA SEMANA QUE VA A SER EL SABADO
   If Semana% = 1 Then
      FESEMACTIVA = FETEXCOR1$(DSPOS%)
      SQLX2$ = " AND FecNecesi  <= '" & FESEMACTIVA & "'"
   ElseIf Semana% = 2 Then  'SI ES LA SEGUNDA SEMANA CONSULTA POR EL RANGO DESDE EL PRINCIPIO DE LA SEGUNDA SEMANA A LA ACTIVA - HASTA EL FIN DE ESTA.
      FESEMADDE = FETEXCOR1$(DSPOS% + 1)
      FESEMAHST = FETEXCOR1$(DSPOS% + 7)
      SQLX2$ = " AND FecNecesi  >= '" & FESEMADDE & "'"
      SQLX2$ = SQLX2$ + " AND FecNecesi  <= '" & FESEMAHST & "'"
   End If
   '
   SQLX$ = " select SUM(CANTIDISP) AS DISPO from coberfa WITH(NOLOCK) "
   SQLX$ = SQLX$ + " Where COD_INTE = " & CI1
   SQLX$ = SQLX$ + "  AND TIPOMOVI = 15 " ' of
   SQLX$ = SQLX$ + "  AND ID = " & ID1
   SQLX$ = SQLX$ + SQLX2$
   Set rs = READSET(SQLX$)
   Do While Not rs.EOF
     NECEFABRICAR_MRP = NECEFABRICAR_MRP + XVALO(rs!Dispo) ' ACUMULA
     rs.MoveNext
   Loop
   rs.Close
   Set rs = Nothing
End Function
Function EXISMASTER2%(CODIGOX$)
   EXISMASTER2% = 0
   If CantRegistros("MASTER2", "CODIGO = '" & CODIGOX$ & "'", "NOMESS") > 0 Then EXISMASTER2% = 1
End Function

Function COSCALCUHISTO(Ci%, MODOCO%, PERIODO$, NIVELCOSTO%)

    Dim dbCostoCalculado As Double
    Dim sCampo As String
    
    CondicionSql$ = "PERIODO='" & PERIODO$ & "' AND COD_INTE=" & CStr(Ci%) & " ORDER BY FECOSCAL DESC"
    sCampo = "TotaSuminis"
    If NIVELCOSTO% = 1 Then
        sCampo = "TotCostoDirecto"
    ElseIf NIVELCOSTO% = 2 Then
        sCampo = "CostoStandard"
    End If
    dbCostoCalculado = XVALO(VALOBADA("HISTOCOSTOS", sCampo, CondicionSql$))
    
    If dbCostoCalculado < 0.00005 Then
        dbCostoCalculado = COSCALCU(Ci%, MODOCO%)
    End If
    
    COSCALCUHISTO = dbCostoCalculado

End Function
Function REVISO$(CI1%)
    CONDI$ = "COD_INTE = " & CI1%
    REVISO = TRIM$(VALOBADA("ESPECIFIC", "REVISO", CONDI$, "NOMESS"))
End Function
Function APROBO$(CI1%)
    CONDI$ = "COD_INTE = " & CI1%
    APROBO$ = TRIM$(VALOBADA("ESPECIFIC", "APROBO", CONDI$, "NOMESS"))
End Function
Function PREPARO$(CI1%)
    CONDI$ = "COD_INTE = " & CI1%
    PREPARO$ = TRIM$(VALOBADA("ESPECIFIC", "preparo", CONDI$, "NOMESS"))
End Function
Function EXISTE_COD_ESPECIFIC(CI1%)
   CONDI$ = "COD_INTE = " & CI1%
   EXISTE_COD_ESPECIFIC = CantRegistros&("ESPECIFIC", CONDI$, "NOMESS")
End Function
Function PROXIMO_ID(TABLA$, CAMPO$, CONDI$)
   SQLX$ = "SELECT TOP 1 " & CAMPO$
   SQLX$ = SQLX$ + " AS PRO_NUMERO"
   SQLX$ = SQLX$ + " FROM " & TABLA$
   SQLX$ = SQLX$ + " WITH(NOLOCK)"
   SQLX$ = SQLX$ + " ORDER BY " & CAMPO$
   SQLX$ = SQLX$ + " DESC"
   Set rs = READSET(SQLX$)
   PROXIMO_ID = 1 + XVALO(rs!PRO_NUMERO)
End Function


Function REFEGRUPOACTU$(GRACT%)
      DESCRITABL$ = "GRUPOACTU"
      GRT& = GRACT%
      REFEGRUPOACTU$ = VALOBADA("DATASQL", "CAMPO2", "LISTA = '" & DESCRITABL$ & "' AND CAMPO1 = '" & FORMATOCON0$(GRT&, 12) & "'", "NOMESS")
End Function

Function ArticuloGrupoContable(CodigoInterno&) As Integer

    GRUPO% = XVALO(VALOBADA("ARTIGRUCONTA", "GRUCOVEN", "CODINT=" & CStr(CodigoInterno&)))
    ArticuloGrupoContable = GRUPO%
    
End Function
 
Function CodigoArba$(CI1%)
    CONDI$ = "CODINT = " & CI1%
    CodigoArba$ = TRIM$(VALOBADA("MASTER", "Cod_Arba", CONDI$, "NOMESS"))
End Function

Function ListaSeleccionArticulos(Optional RetornaSegundoCampoEco As String, Optional Refresh As Boolean = False, Optional CONDI$) As String

    TituloLista$ = "Maestro de articulos"
    If TRIM$(FRMZELECHO.Caption) = TituloLista$ And Refresh = False Then GoTo 50
    
    Call CARGA_ENCABEZADO(FRMZELECHO.msf2, "Codigo/A24;Descripcion/A64", FRMZELECHO)
    Call CARGA_ENCABEZADO(FRMZELECHO.MSF1, "Codigo/A24;Descripcion/A64", FRMZELECHO)
    FRMZELECHO.Caption = TituloLista$
    
    Dim obj As New RECORXET
    Set rs1 = obj.RS_MAESTRODEARTICULOS(CONDI$)
    Call Carga_MSF_Recordset(rs1, "Codigo/A24;Descripcion/A64", FRMZELECHO.msf2, 1)
    On Error Resume Next
    Set obj = Nothing
    rs1.Close
    Set rs1 = Nothing
    On Error GoTo 0
        
    FRMZELECHO.Width = 10000
50  FRMZELECHO.Show 1

    CodigoArticulo$ = TRIM(RESP_ZELE_VECT(1))
    ListaSeleccionArticulos = CodigoArticulo$
    RetornaSegundoCampoEco = TRIM(RESP_ZELE_VECT(2))
    
End Function

Function CANTIDAD_PEDIDA(CI1%, Des%, Has%)
     'DEVUELVE LA CANTIDAD PEDIDA DE UN ITEM EN UN RANGO DE TIEMPO SIN IMPORTAR SI TIENE ENTREGAS
     Call ABREPEDCLI
     TOPEDID = 0
     SQLX$ = " SELECT SUM(CanPed) AS CANTIPED FROM PEDDETA WITH(NOLOCK) "
     SQLX$ = SQLX$ + " WHERE CodiInt = " & CI1% & " "
     SQLX$ = SQLX$ + " AND VenRep = 'V' "
     SQLX$ = SQLX$ + " AND Status < 8 "   ' para que no este anulada ni cumplida
     SQLX$ = SQLX$ + " AND FechEmis >= '" & FETEXCOR1$(Des%) & "'"
     SQLX$ = SQLX$ + " AND FechEmis <= '" & FETEXCOR1$(Has%) & "'"
     '
     'ATENCION NO REALIZO UN SUM DE LAS CANTIDADES POR SI EN PEDIDOS HAY ENTREGAS MAYORES A LA CANTIDAD SOLICITADA
     Set PEDTEMP = READSET(SQLX$)
     With PEDTEMP
        TOPEDID = XVALO(!CANTIPED)
     End With
     PEDTEMP.Close
     Set PEDTEMP = Nothing
     
     CANTIDAD_PEDIDA = TOPEDID

End Function
Function CANTI_INFO_DESPA(CI1%)
     CANTI_INFO_DESPA = 0
     SQLX$ = " SELECT SUM(CanPed) AS CANTIPED , SUM(CANTENT) AS CANTIEN FROM PEDDETA WITH(NOLOCK) "
     SQLX$ = SQLX$ + " WHERE CodiInt = " & CI1% & " "
     SQLX$ = SQLX$ + " AND VenRep = 'V' "
     SQLX$ = SQLX$ + " AND Status < 8 "   ' para que no este anulada ni cumplida
     SQLX$ = SQLX$ + " AND StaDespa > 0 "
     SQLX$ = SQLX$ + " AND CanPed > CANTENT "
     
'     SQLX$ = SQLX$ + " AND FechEmis >= '" & FETEXCOR1$(DES%) & "'"
'     SQLX$ = SQLX$ + " AND FechEmis <= '" & FETEXCOR1$(HAS%) & "'"
     '
     'ATENCION NO REALIZO UN SUM DE LAS CANTIDADES POR SI EN PEDIDOS HAY ENTREGAS MAYORES A LA CANTIDAD SOLICITADA
     Set PEDTEMP = READSET(SQLX$)
     With PEDTEMP
        YADESPA = XVALO(!CANTIPED) - XVALO(!CANTIEN)
     End With
     PEDTEMP.Close
     Set PEDTEMP = Nothing
     If YADESPA < 0 Then YADESPA = 0
     
     CANTI_INFO_DESPA = YADESPA

End Function
Function DefineModelo(CI1%)

End Function
Function ControlAsientoImpo(CI1%, CANTI, LOTE$, SOLOTRAZABLES_n%)
    ControlAsientoImpo = 0
    If Abs(XVALO(CANTI)) < 1 Then Exit Function
    If ESTRAZABLE%(CI1%) <> SOLOTRAZABLES_n% Then Exit Function 'SOLO INGESA SI SON TRAZABLES = AL PARAMETRO

    CODI$ = CODEXT$(CI1%)
    Dim obj As New Art_Importados
    Call obj.VeriImportado(CODI$)
    If obj.EsImportado < 1 Then GoTo 99
    If LOTE$ = "" Then 'SI NO TIENE LOTE LE ASIGNA EL COSTO DE REPOSICION
       ControlAsientoImpo = XVALO(COSUNI(CI1%))
       GoTo 99
    End If

    CONDI$ = "IDENLOTE = '" & LOTE$ & "' AND COD_INTE = " & CI1%
    IMPOUNT = XVALO(VALOBADA("DETFACOMSQL", "PreunitC_Bonif", CONDI$, "NOMESS"))
    ControlAsientoImpo = CANTI * IMPOUNT
    'SI NO ENCUENTRA LE ASIGNA EL COSTO DE REPOSICION
    If ControlAsientoImpo < 0.005 Then
      ControlAsientoImpo = CANTI * XVALO(COSUNI(CI1%))
    End If
99  Set obj = Nothing
End Function

Function ANCHOMASTER(Ci%)
   'ENVAPLAST
   On Error Resume Next
   If STATUIT%(Ci%) <> 0 Then ANCHOMASTER = XVALO(MasterTemp!ANCHO)
   '
End Function
Function ANCHOESPECIFIC(CI1%)
    CONDI$ = "COD_INTE = " & CI1% & " ORDER BY ID"
    'EJEMPLO DE INFORMACION QUE DEVUELVE: 47X76X0X0,8
    'HAGO UN XVALO PARA QUEDARME CON EL PRIMER VALOR HASTA LA X EN ESTE CASO EL 47
    ANCHOESPECIFIC = XVALO(VALOBADA("ESPECIFIC", "TOP 1 DESCRIPCION", CONDI$, "NOMESS"))
End Function
Function CORTEESPECIFIC(CI1%)
    CONDI$ = "COD_INTE = " & CI1% & " ORDER BY ID"
    'EJEMPLO DE INFORMACION QUE DEVUELVE: 47X76X0X0,8
    'HAGO UN XVALO PARA QUEDARME CON EL PRIMER VALOR HASTA LA X EN ESTE CASO EL 47
    CORTEESPECIFIC = UCase(VALOBADA("ESPECIFIC", "TOP 1 CORTE", CONDI$, "NOMESS"))
End Function

Function DescripXNumCodigo$(CI1%, ARGU_BUSQUE$)
   'BUSCA EN LA TABLA ESPECIFIC SEGUN EL CODIGO INTERNO DEL PRODUCTO TERMINADO SI TIENE ALGUN COMPONENTE
   'QUE COMIENCE CON LOS PARAMETROS PASADOS COMO SEGUNDO ARGUMENTO.
   CantCaract% = Len(ARGU_BUSQUE$)
   CONDI$ = "COD_INTE = " & CI1% & " AND LEFT(CODIDATO," & CantCaract% & ") = '" & ARGU_BUSQUE$ & "'"
   CODIEXT_ESPECIFIC$ = VALOBADA("ESPECIFIC", "CODIDATO", CONDI$)
   CiComp% = CODINT%(CODIEXT_ESPECIFIC$)
   If CiComp% > 0 Then
      DescripXNumCodigo$ = DESCRIT0$(CiComp%)
   Else
      DescripXNumCodigo$ = ""
   End If
End Function
Function ANCHOXNumCodigo$(CI1%)
   'BUSCA EN LA TABLA ESPECIFIC SEGUN EL CODIGO INTERNO DEL PRODUCTO TERMINADO SI TIENE ALGUN COMPONENTE
   'UBICADO CON EL ITEM 'PLIEGO EXTERIOR' DEVUELVE EL ANCHO DE ESTE DEL MASTER
   CONDI$ = "COD_INTE = " & CI1% & " AND ITEM = 'PLIEGO EXTERIOR'"
   CODIEXT_ESPECIFIC$ = VALOBADA("ESPECIFIC", "CODIDATO", CONDI$)
   CiComp% = CODINT%(CODIEXT_ESPECIFIC$)
   If CiComp% > 0 Then
       strArray = Split(AnchoColorGramaje$(CiComp%), ";")
       ANCHOXNumCodigo$ = strArray(0)
   Else
       ANCHOXNumCodigo$ = ""
   End If
End Function
Function AnchoColorGramaje$(CI1%)
   AnchoColorGramaje$ = " ; ; ;"
   SQLX$ = "SELECT ANCHO,GRAMAJE, COLOR, MICRONES FROM MASTER WITH(NOLOCK)"
   SQLX$ = SQLX$ + " WHERE CODINT = " & CI1%
   Set rs = READSET(SQLX$)
   If Not rs.EOF Then
        AnchoColorGramaje$ = XVALO(rs!ANCHO) & ";"
        AnchoColorGramaje$ = AnchoColorGramaje$ & XVALO(rs!GRAMAJE) & ";"
        AnchoColorGramaje$ = AnchoColorGramaje$ & XVALO(rs!COLOR) & ";"
        AnchoColorGramaje$ = AnchoColorGramaje$ & XVALO(rs!MICRONES)
   End If
   rs.Close
   Set rs = Nothing
End Function


Sub ACTUALIZAVECTORDEIVA()
    For i& = 1 To 32767
       TASAIVAX_CI(i&) = 0
    Next i&
    
    Static LATABLAEXISTE%, URE&
    If VALIDA_TASAIVA$ = "" Then
       VALIDA_TASAIVA$ = Control("FACTURA", "VALITIVA")
       If VALIDA_TASAIVA$ <> "SI" Then VALIDA_TASAIVA$ = "NO"
    End If
    '
    If EXISTE_TABLA%("COMBINA_IVA") > 0 Then
          LATABLAEXISTE% = 1
    Else
          LATABLAEXISTE% = -1
          Exit Sub
    End If

    If VALIDA_TASAIVA$ <> "SI" Or LATABLAEXISTE% = 0 Then
       Exit Sub
    End If
    '
    If URE& <= 0 Then
        URE& = ULTREG&("GRUPIVA")
        For U& = 1 To URE&
           X$ = REGLEIDO$("GRUPIVA", U&)
           TAIVA(U& - 1) = CVS(Mid$(X$, 13, 4))
        Next U&
    End If

    'VERIFICAR PRIMERO SI DENTRO DEL VECTOR HAY ALGUN ITEM EQUIPO Y PONER EN OTRO VECTOR ESTOS.
    'SI ES ASI VERIFICAR QUE EL RESTO MACHEE CON ALGUNO DE ESTOS
    For i& = 1 To UBound(CI_VECTOR_TIVA)
       CI1% = XVALO(CI_VECTOR_TIVA(i&))
       If CI1% > 0 Then
          CONDI$ = "CodintEq = " & CI1%
          If CantRegistros("COMBINA_IVA", CONDI$, "NOMESS") > 0 Then
             J& = J& + 1
             If XVALO(TAIVA(1)) = XVALO(PORCEIVA(CI1%)) Then  'SI ES IGUAL A LA TAZA MINIMA LO CONSIDERA
                If J& = 1 Then
                   KONDI$ = KONDI$ + " (CodintEq = " & CI1%   'Y ARMA UNA CONDICION PARA LOS ACCESORIOS
                Else
                   KONDI$ = KONDI$ + " OR CodintEq = " & CI1% 'LIGADOS A ESTOS
                End If
             End If
          End If
       End If
    Next i&
    '
    'ahora busco bajo la condicion que exista como accesorio y vinculado a alguno de los equipos
    'ejemplo si en la misma fila esta el equipo 100 y el accesorio acc2 si coincide le aplico la tasa minima.
    If TRIM$(KONDI$) <> "" Then KONDI$ = KONDI$ + ")"
    
    If J& > 0 Then
        For i& = 1 To UBound(CI_VECTOR_TIVA)
           CI1% = XVALO(CI_VECTOR_TIVA(i&))
           If CI1% > 0 Then
              CONDI$ = "CodintACC = " & CI1%
              If TRIM$(KONDI$) <> "" Then
                 CONDI$ = CONDI$ + " AND " + KONDI$
              End If
              If CantRegistros("COMBINA_IVA", CONDI$, "NOMESS") > 0 Then
                 TASAIVAX_CI(CI1%) = XVALO(TAIVA(1)) 'SI LOS ACCESORIOS MACHEAN CON LOS EQUIPOS LE ASIGNA LA TAZA MINIMA
              End If
           End If
        Next i&
    End If
    
    
End Sub

Sub TABLA_COMBINA_EQUIPOYACCESORIOS()
        Call CREATABLA_ARTICULOS
        VENTANA.Campox = "Cód.Equipo/A24;Descripción/A40;Cód.Accesorio/A24;Descripción/A40"
        VENTANA.Inicializar = 0
        VENTANA.AgregaRegistro = 1
        VENTANA.UtilizaFormula = 1
        VENTANA.HabilitaBorrar = 1
        VENTANA.HabilitaInsertar = 1
        VENTANA.NoMostrarBotonTilde = 0
        VENTANA.TITULO = "COMBINACIÓN DE EQUIPOS CON ACCESORIOS (TASA DE IVA)"
        VENTANA.CampEditables = "1;3"
        VENTANA.ColEcoCampo(1) = "CODIGO;DESCRIPCION;MASTER"
        VENTANA.ColEcoCampo(3) = "CODIGO;DESCRIPCION;MASTER"
        VENTANA.Inicializar = 1
        Campos$ = "CodigoEq,DescripcionEq,CodigoAcc,DescripcionAcc"
        Dim obj As New RECORXET
        Set rs1 = obj.RS_COMBINAEQUIACC(Campos$, "ORDER BY CodigoEq")
        Call Carga_MSF_Recordset(rs1, "Cód.Equipo/A24;Descripción/A32;Cód.Accesorio;Descripción/A32", VENTABNEW.MSF)
        Set obj = Nothing
        rs1.Close
        Set rs1 = Nothing
        '*****************************************************'

30      VENTABNEW.Show 1
        Dim MSF As Object
        Set MSF = VENTABNEW.MSF
        '
        If VENTABNEW.APROBADO% < 1 Then Exit Sub

        If VENTABNEW.APROBADO% > 0 Then
             If FilaVacia%(VENTABNEW.MSF, 1) < 1 Then
               VALIEC1& = VALIDAR_ECO(VENTABNEW.MSF, 1, 0)
               If VALIEC1& > 0 Then
                 Call COMUNI("Faltan Datos En Fila: " & VALIEC1&)
                 GoTo 30
               End If
               VALIEC2& = VALIDAR_ECO(VENTABNEW.MSF, 3, 0)

               If VALIEC2& > 0 Then
                 Call COMUNI("Faltan Datos En Fila: " & VALIEC2&)
                 GoTo 30
               End If
             End If
        End If
        '
        'TRANSACCIÓN GRABACION
        On Error GoTo ERROR_GUARDAR
        FLEXCONN.BeginTrans
        Call BORRAREGISTROS("COMBINA_IVA", "CodintEq > 0", REGAF&, "NOMESS")
        For REG& = 1 To MSF.Rows - 1
          CODIEQ$ = TRIM$(MSF.TextMatrix(REG&, 1))
          CODIACC$ = TRIM$(MSF.TextMatrix(REG&, 3))
          DESCREQ$ = TRIM$(MSF.TextMatrix(REG&, 2))
          DESCRACC$ = TRIM$(MSF.TextMatrix(REG&, 4))
          CodintEq = CODINT(CODIEQ$)
          CODINTACC = CODINT(CODIACC$)

          If CodintEq > 0 And CODINTACC <> "" Then
             SQLX$ = "INSERT INTO COMBINA_IVA "
             SQLX$ = SQLX$ + " (CodintEq,CodigoEq,DescripcionEq,CodintAcc,CodigoAcc,DescripcionAcc) "
             SQLX$ = SQLX$ + " VALUES(" & CodintEq
             SQLX$ = SQLX$ + " ,'" & CODIEQ$ & "'"
             SQLX$ = SQLX$ + " ,'" & DESCREQ$ & "'"
             SQLX$ = SQLX$ + " ," & CODINTACC
             SQLX$ = SQLX$ + " ,'" & CODIACC$ & "'"
             SQLX$ = SQLX$ + " ,'" & DESCRACC$ & "'"
             SQLX$ = SQLX$ + ")"
             FLEXCONN.Execute (SQLX$)
          End If
        Next REG&
        Set VENTANA = Nothing


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
End Sub


Sub GuardarCodigoCliente(CI1%, NC, CODICLIENTEX$)
   If CI1% < 1 Then Exit Sub
   If NC < 1 Then Exit Sub
   If Len(TRIM$(CODICLIENTEX$)) > 28 Then
      Call COMUNI("Atención!!! Código de Cliente Supera el Máximo Permitido de 28 Caracteres")
      Exit Sub
   End If
   ASELE$ = "REPLACOC"
   If NC < 0 Then ASELE$ = "REPLACOP"
   '
   Encontro% = 0
   CDXXX$ = CODEXT$(CI1%)
   UREP& = ULTREG&(ASELE$)
   RFF$ = RECFILT$(ASELE$, 3, MKI$(CI1%))
   '
   For KKI% = 1 To Len(RFF$) Step 4
     REREP& = CVS(Mid$(RFF$, KKI%, 4))
     If REREP& > 0 Then
       If REREP& <= UREP& Then
         XXYY$ = REGLEIDO$(ASELE$, REREP&)
         If CVI(Mid$(XXYY$, 1, 2)) = Abs(NC) Then
           CDYYY$ = TRIM$(Mid$(XXYY$, 5, 28))
           If CDYYY$ <> "" Then
              Call REPLA(XXYY$, TRIM$(CODICLIENTEX$), 5, 28)
              Call GRAREG(ASELE$, XXYY$, REREP&)
              Encontro% = 1
              GoTo 99
           End If
         End If
       End If
     End If
   Next KKI%
   
   If Encontro% < 1 Then
    Call REPLA(XXX$, MKI$(NC), 1, 2)
    Call REPLA(XXX$, MKI$(CI1%), 3, 2)
    Call REPLA(XXX$, CODICLIENTEX$, 5, 28)
    Call REPLA(XXX$, "", 33, 96)
    Call REGAPP(ASELE$, XXX$)
   End If
   '
99
End Sub

Function NUEVOORDENCODIGOMASTERDELCLIENTE%(NC1&)
   NUEVOORDENCODIGOMASTERDELCLIENTE% = 1
   NUCLIENTE$ = FORMATOCON0$(NC1&, 6)
   SQLX$ = "SELECT MAX(CAST(SUBSTRING(CODIGO,8,2) AS INT)) AS MAXIMOORDEN FROM MASTER WITH(NOLOCK)"
   SQLX$ = SQLX$ + " WHERE CAST(SUBSTRING(CODIGO,1,6) AS VARCHAR) =  '" & NUCLIENTE$ & "'"
   Set rs = READSET(SQLX$)
   With rs
      If Not .EOF Then
         NUEVOORDENCODIGOMASTERDELCLIENTE% = 1 + XVALO(!MAXIMOORDEN)
      End If
   End With
   rs.Close
   Set rs = Nothing
End Function
Function NUEVOORDENCODIGOMASTERDELCLIENTEBOLSA%(NC1&)
   NUEVOORDENCODIGOMASTERDELCLIENTEBOLSA% = 1
   NUCLIENTE$ = FORMATOCON0$(NC1&, 6)
   TEXTCOMPLETO$ = "CTP-" & NUCLIENTE$
   'CTP-001234-00
   '1234567890123
   SQLX$ = "SELECT MAX(CAST(SUBSTRING(CODIGO,12,2) AS INT)) AS MAXIMOORDEN FROM MASTER WITH(NOLOCK)"
   SQLX$ = SQLX$ + " WHERE CAST(SUBSTRING(CODIGO,1,10) AS VARCHAR) =  '" & TEXTCOMPLETO$ & "'"
   Set rs = READSET(SQLX$)
   With rs
      If Not .EOF Then
         NUEVOORDENCODIGOMASTERDELCLIENTEBOLSA% = 1 + XVALO(!MAXIMOORDEN)
      End If
   End With
   rs.Close
   Set rs = Nothing
End Function
Sub ANOTAHISMASTER(CI1%, TXTX$, Optional FF%, Optional MODO$)
    '
    If CI1% > 0 Then
       '
       HOII% = HOY(HOS$)
       fech$ = HOS$
       If FF% > 0 Then
          fech$ = FECHATEX(FF%) ' SI TRAE FECHA EL PARAMETRO TOMO ESTA
       End If
       '
       Call FRATEXTO(TXTX$, 56) ' FRACCIONO EL TEXTO
       '
       For KKI% = 1 To CARETEX% ' RECORRO DESDE 1 A LA CANTIDAD DE FRACCIONES
           TTYY$ = Space$(96)
           If KKI% = 1 Then
              Call REPLA(TTYY$, fech$, 1, 8) ' SI ES LA PRIMERA VEZ GRABA LA FECHA Y USUARIO
              Call REPLA(TTYY$, USERTER$, 69, 24)
           End If
           Call REPLA(TTYY$, RETEX$(KKI%), 11, 56) ' SIEMPRE QUE PASA GUARDA EL CENTRO DEL TEXTO
           ANOTANTE$ = ANOTANTE$ + TTYY$ + vbCrLf
       Next KKI%
       '
       ANOTANTE$ = ANOTANTE$ & String$(88, "=") & vbCrLf

       
       SQLX$ = "SELECT HISTORIAL,CODINT,Codigo,Descripcion FROM ANOTAMASTER"
       SQLX$ = SQLX$ + " WHERE CODINT = " & CI1%
       Dim rs As New ADODB.Recordset
       Set rs = New ADODB.Recordset
25     On Error Resume Next
       rs.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
       If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
           On Error GoTo 0
           Call CapturaErrorOpen
           GoTo 25
       End If
       On Error GoTo 0
       '
       With rs
         If (.BOF And .EOF) Then
            .AddNew
            !CODINT = CI1%
            !Codigo = CODEXT$(CI1%)
            !Descripcion = DESCRIT0$(CI1%)
         End If
            ANOTRIOR$ = SAFETEXT$(!Historial)
            !Historial = ANOTRIOR$ + ANOTANTE$
            .Update
       End With
       rs.Close
       Set rs = Nothing
       '
    End If
    '
End Sub

Function NUEVAVERSIONCODIGOMASTERDELCLIENTE%(NC, CI1%)
   NUEVAVERSIONCODIGOMASTERDELCLIENTE% = 0
   '123456-89-11
   'SI NO TIENE DATOS VA A DEVOLVER 0 PARA QUE SE ASIGNE VERSION 0
   'SI YA TIENE UNA VERSION 0 VA A DEVOLVER 1
   NC1& = NC
   NUCLIENTE$ = FORMATOCON0$(NC1&, 6)
   SQLX$ = "SELECT MAX(CAST(SUBSTRING(CODIGO,11,2) AS INT) AS MAXIMOVERSION )FROM MASTER WITH(NOLOCK)"
   SQLX$ = SQLX$ + "FROM MASTER WHERE CAST(SUBSTRING(CODIGO,4,6) AS VARCHAR) =  " ' & NUCLIENTE$ & "'"
   Set rs = READSET(SQLX$)
   With rs
      If Not .EOF Then
         NUEVAVERSIONCODIGOMASTERDELCLIENTE% = 1 + XVALO(!MAXIMOVERSION)
      End If
   End With
   rs.Close
   Set rs = Nothing
End Function

Function GENERAR_NUEVA_VERSION_NUMERICA%(PARCIALCODIGO$)

   COD$ = PARCIALCODIGO$
'   SELECT * FROM MASTER WHERE LEFT(CODIGO,10)  LIKE '00-00001-%'
'   ORDER BY CAST(SUBSTRING(CODIGO, 11,2) AS INTEGER)
   CONDI$ = "LEFT(CODIGO,10)  LIKE '" & COD$ & "%'"
   CANATIVERSIONES% = CantRegistros&("MASTER", CONDI$, "NOMESS")
   If CANATIVERSIONES% < 1 Then
     GENERAR_NUEVA_VERSION_NUMERICA% = 0
   Else
     GENERAR_NUEVA_VERSION_NUMERICA% = CANATIVERSIONES%
   End If
   '
End Function
Function GENERAR_NUEVA_INDICE_CTA_99000%(NC)
   'NO HAY QUE SUMARLE 1 POR QUE LA PRIMERA VEZ INICIALIZA EN 0 EN VEZ DE 1
   NC1& = NC
   NC_FORMATEADO$ = FORMATOCON0$(NC1&, 6) & "-"
   CONDI$ = "LEFT(CODIGO,7)  LIKE '" & NC_FORMATEADO$ & "%'"
   CantiReg& = CantRegistros&("MASTER", CONDI$, "NOMESS")
   If CantiReg& < 1 Then
     GENERAR_NUEVA_INDICE_CTA_99000% = 0
   Else
     If CantiReg& < 10000 Then
        GENERAR_NUEVA_INDICE_CTA_99000% = CantiReg&
     Else
        Screen.MousePointer = 11
        CANTIDADABORRAR = 9900
        CODIGOX$ = "099000-00-00"
        For i& = 0 To CANTIDADABORRAR
           'SI LLEGA A TENER MAS DE 9999 REGISTROS DEL CLIENTE 99000 BORRA LOS 9900 PRIMEROS
           'VERLO COMO FUNCIONA REALMENTE ESTO A FUTURO
           'CUANDO LUEGO DE BORRAR ESTO SE VAYAN DANDO DE ALTA LOS CODIGOS A PARTIR DEL 099000-01-00
           'VA A LLEGAR AL 099000-99-00 y estos van a ser existentes habria que ver de eliminarlos en ese momento.
           CONDI$ = "CODIGO= '" & CODIGOX$ & "'"
           Call BORRAREGISTROS("MASTER", CONDI$, REGAF&, "NOMESS")
           indice$ = Mid$(CODIGOX$, 8, 2) & Right(CODIGOX$, 2)
           INDINUM& = XVALO(indice$) + 1
           INDICEFORMAT$ = FORMATOCON0$(INDINUM&, 4)
           DELMEDIO$ = Mid$(INDICEFORMAT$, 1, 2)
           ULTIMOS2$ = Right(INDICEFORMAT$, 2)
           CODIGOX$ = NC_FORMATEADO$ & DELMEDIO$ & "-" & ULTIMOS2$
        Next i&
        Screen.MousePointer = 1
     End If
   End If
   '
   
   '
End Function
Function CodigoIntSegunLote(LOTE$)
   CodigoIntSegunLote = XVALO(VALOBADA("Bolrecep", "Cod_inte", "Idenlote = '" & LOTE$ & "'", "Nomess"))
End Function
Function CodigoIntSegunLoteDelMovisto(LOTE$)
   'caso de fabricacion
   'BUSCA SI ES UN PRODUCTO FABRICADO  PRIMERO, SI NO LO BUSCA POR RECEPCION
   LOTE$ = TRIM$(LOTE$)
   If Left$(LOTE$, 3) = "OF-" Then
      CodigoIntSegunLoteDelMovisto = XVALO(VALOBADA("Movisto", "Cod_inte", "Idenlote = '" & LOTE$ & "' and Codimovi = 'OF' AND CANTINGRE > 0", "NOMESS"))
   End If
   
   If Left$(LOTE$, 3) = "BR-" Then
      CodigoIntSegunLoteDelMovisto = XVALO(VALOBADA("Movisto", "Cod_inte", "Idenlote = '" & LOTE$ & "' and Codimovi = 'BR' AND CANTINGRE > 0", "NOMESS"))
   End If
End Function
Function CODIGORANGO&(CIMP)
    'DEVUELVE EL CODIGO RANGO DE UNA MATERIA PRIMA
    CODIGORANGO& = XVALO(VALOBADA("CODIGORANGO", "CIRANGO", "CIMATPRIMA = " & CIMP, "NOMESS"))
End Function
Function ESCODIGORANGO&(CI1%)
  'devuelve >= 1 si es codigo rango
  ESCODIGORANGO& = CantRegistros&("CODIGORANGO", "CIRANGO = " & CI1%, "NOMESS")
End Function
Function CODIGORANGOEXTE$(CIMP)
    'DEVUELVE EL CODIGO RANGO DE UNA MATERIA PRIMA
    CODIGORANGOEXTE$ = VALOBADA("CODIGORANGO", "CODRANGO", "CIMATPRIMA = " & CIMP, "NOMESS")
End Function
Function CODIGORANGOEXTE_SEGUNCODIGORANGOINTERNO$(CODINTRANGO)
    'DEVUELVE EL CODIGO RANGO EXTERNO SEGUN EL CODIGO INTERNO DE LA TABLA CODIGORANGO
    CODIGORANGOEXTE_SEGUNCODIGORANGOINTERNO$ = VALOBADA("CODIGORANGO", "CODRANGO", "CIRANGO = " & CODINTRANGO, "NOMESS")
End Function
Function CODIGORANGOINTE_SEGUNCODIGORANGOEXTERNO$(CODEXRANGO$)
    'DEVUELVE EL CODIGO RANGO INTERNORNO SEGUN EL CODIGO EXTERNO DE LA TABLA CODIGORANGO
    CODIGORANGOINTE_SEGUNCODIGORANGOEXTERNO$ = VALOBADA("CODIGORANGO", "CIRANGO", "CODRANGO = '" & CODEXRANGO$ & "'", "NOMESS")
End Function
Function CodigodeCotizacion&(COT&, COMPONENTE$)
   'devuelve el codigo rango pasandole el nro de cotizacion y el identificador de la operacion
   COMPONENTE$ = TRIM$(COMPONENTE$)
   LONGITCOMPONENTE% = Len(COMPONENTE$)
   CONDI$ = "NROCOTIZACION = " & COT& & " AND LEFT(COMPONENTE," & LONGITCOMPONENTE% & " ) = '" & COMPONENTE$ & "'"
   CICOMPO& = XVALO(VALOBADA("COTI1", "CODINTELEM", CONDI$, "NOMESS"))
   CodigodeCotizacion& = CODIGORANGO&(CICOMPO&)
End Function

Function ListaSeleccion_ListasDePrecios() As Long

    Campos$ = "Id/I5;Denominación/A32"
    Call CARGA_ENCABEZADO(FRMZELECHO.MSF1, Campos$, FRMZELECHO, , , ANTOT)
    Call CARGA_ENCABEZADO(FRMZELECHO.msf2, Campos$, FRMZELECHO, , , ANTOT)
    FRMZELECHO.Caption = "Lista de Precios"
    
    ConsultaSql$ = "select cod_lista, denom_lista from listapre with(nolock) order by cod_lista"
    Dim rst_ListasPrecio As ADODB.Recordset
    Set rst_ListasPrecio = READSET(ConsultaSql$)
    
    FRMZELECHO.msf2.Rows = rst_ListasPrecio.RecordCount + 1
    fi& = 0
    With rst_ListasPrecio
        Do While Not .EOF
            fi& = fi& + 1
            FRMZELECHO.msf2.TextMatrix(fi&, 1) = CStr(!Cod_Lista)
            FRMZELECHO.msf2.TextMatrix(fi&, 2) = !DENOM_LISTA
        
            .MoveNext
        Loop
    End With
    
    On Error Resume Next
    rst_ListasPrecio.Close
    Set rst_ListasPrecio = Nothing
    On Error GoTo 0
    
    FRMZELECHO.Width = 5000
    FRMZELECHO.Show 1

    CodigoLista& = XVALO(RESP_ZELE_VECT(1))

    ListaSeleccion_ListasDePrecios = CodigoLista&
    
End Function

Sub BAJAARTICULO(CII%)
   If CII% <= 0 Then Exit Sub
   LONCODI% = XVALO(Control$("MASTER", "LOMAXCOD"))
   If LONCODI% < 15 Or LONCODI% > 24 Then LONCODI% = 15
   
   COD$ = AJUSTI$(CODEXT$(CII%), LONCODI%)
   CONDI$ = "CodiElem = " & CII%
   QTR1& = CantRegistros&("Formulas", CONDI$)
   QTR2& = CantRegistros&("FormuPAT", CONDI$)
   If QTR1& > 0 Or QTR2& > 0 Then
      'Call MENSERR(24, "Imposible Baja de este Código\por Ser Componente de Fórmula Activa.\   \Consulte Implosion y Desconecte\de los Productos Padre Correspondientes.")
      'SE INCORPORA EN EL MENSAJE LA LISTA DE CODIGOS PADRES (ACTIVOS O DE BAJA)
      SQLX$ = "SELECT * FROM FORMULAS WITH(NOLOCK) WHERE CODIELEM=" & CII%
      Set rs = READSET(SQLX$)
      II& = 0
      Do While Not rs.EOF
            Codigo$ = CODEXT$(Abs(XVALO(rs!CODICONJ)))
            LISTACODIGOS$ = LISTACODIGOS$ & Codigo$ & ";"
            II& = II& + 1
            If II& = 4 Then
                LISTACODIGOS$ = LISTACODIGOS$ & "\"
                II& = 0
            End If
        rs.MoveNext
      Loop
      'SE CAMBIA MENSAJE PORQUE SE HA DEJADO LA POSIBILIDAD DE ELIMINAR EL ITEM
      Call MENSERR(24, "El Codigo que Intenta dar de Baja\Es Componente de la Fórmula de Producto (Activo o de Baja)\De los Productos Padre Correspondientes:\ \" & LISTACODIGOS$)
      'GoTo 99
   End If
   '
   If COMALTER%(" Baja Lógica del Codigo '" + TRIM$(COD$) + "' \" + TRIM$(DESCRIT$(CII%)) + "\\Cancelar\Confirmar") <> 2 Then GoTo 99
   '
   Screen.MousePointer = 11
   CII% = CODINT%(COD$)
   If CII% > 0 Then
     strSQL = "Update Master set Status=-1 where Codint=" & CII%
     FLEXCONN.Execute strSQL, cdadreg
'     Call GENEECOMABAJ 'ACTUALIZA LISTA DE SELECCION
   End If
   '
99 Call GENECOMAS
   Screen.MousePointer = 1

End Sub
Function Es_componente%(CODICONJ%, codielem%)
     'DEVUELVE > 0 SI   CodiElem% ES COMPONENTE DEL CodiConj%
     CONDI$ = "CodIConj = " & CODICONJ%
     CONDI$ = CONDI$ + " AND CodiElem = " & codielem%
     Es_componente% = CantRegistros&("FORMULAS", CONDI$, "NOMESS")
End Function
Function ComponenteSegunCodiConjunYFamilia%(CODICONJ%, CONDIFAMILIA$)
     'DEVUELVE el codigo interno del componente si solo hay uno que coincida con la familia
     'si coinciden mas de uno devuelve 0
     'usado para cambaceres para devolver la bobina base
     'EJEMPLO CONDIFAMILIA$ = "FAMILIA = 'TP' OR FAMILIA = 'TT'"
     CantiReg& = 0
     ComponenteSegunCodiConjunYFamilia% = 0
     SQLX$ = "select COUNT(*) as Resul from formulas f WITH(NOLOCK)"
     SQLX$ = SQLX$ + " inner join master m on m.Codint = f.CodiElem "
     SQLX$ = SQLX$ + " where codiconj = " & CODICONJ%
     SQLX$ = SQLX$ + " and (" & CONDIFAMILIA$ & ")"
     Set rs = READSET(SQLX$)
     If Not rs.EOF Then
       CantiReg& = XVALO(rs!RESUL)
     End If

     If CantiReg& <> 1 Then ComponenteSegunCodiConjunYFamilia% = 0 'no lo encuentra o hay mas de un registro de coincidencia
     
     SQLX$ = "select  F.CodiElem from formulas f WITH(NOLOCK)"
     SQLX$ = SQLX$ + " inner join master m on m.Codint = f.CodiElem "
     SQLX$ = SQLX$ + " where codiconj = " & CODICONJ%
     SQLX$ = SQLX$ + " and (" & CONDIFAMILIA$ & ")"
     Set rs = READSET(SQLX$)
     If Not rs.EOF Then
       ComponenteSegunCodiConjunYFamilia% = XVALO(rs!codielem)
     End If
     rs.Close
     Set rs = Nothing
     

     
End Function

Function DESCRIPCIONCODIGORANGO(CODRANGO$)
    DESCRIPCIONCODIGORANGO = VALOBADA("CODIGORANGO", "Descripcion", "CODRANGO = '" & CODRANGO$ & "'", "NOMESS")
End Function
Function PorcentajeMinorista(CI1%)
   'DEVUELVE EL PORCENTAJE QUE ESTA EN LA TABLA COSMINORISTA, EN CASO QUE DEVUELVA 0 , BUSCA EL CONFIGURADO A MODO GENERAL
   CONDI$ = "COD_INTE = " & CI1%
   PorcentajeMinorista = XVALO(VALOBADA("COSMINORISTA", "PORCENTAJE", CONDI$, "NOMESS"))
   If PorcentajeMinorista <= 0.05 Then
      PorcentajeMinorista = ROUND(XVALO(Control$("ANACOSTO", "PORMINOR")), 2)
   End If
   

End Function

Sub GrabarPorcentajeMinorista(CI1%, PORCENTA)
   If CI1% < 1 Then Exit Sub
   'SI NO EXISTE EL REGISTRO LO CREA CON EL CODIGO Y LA FECHA DE INICIO, Y LUEGO AGREGA LA MODIFICACION
   'SI YA EXISTE SOLO AGREGA LA MODIFICACION
   CONDI$ = "Cod_Inte=" & CI1%
   FE% = HOY(HO$)
   fech = FETEXCOR1$(FE%)

   EXISTEREG& = CantRegistros&("COSMINORISTA", CONDI$, "NOMESS")
   If EXISTEREG& < 1 Then

        SQLX$ = "INSERT INTO COSMINORISTA"
        SQLX$ = SQLX$ + " (COD_INTE,FECHINICIO)"
        SQLX$ = SQLX$ + " VALUES(" & CI1%
        SQLX$ = SQLX$ + " ,'" & fech & "'"
        SQLX$ = SQLX$ + " )"
        FLEXCONN.Execute (SQLX$)
   End If
   NumeroUsuario% = USNBR%

   SQLX$ = "UPDATE COSMINORISTA SET "
   SQLX$ = SQLX$ + "NumeroUsuario = " & NumeroUsuario%
   SQLX$ = SQLX$ + ",Porcentaje= " & XVALO(PORCENTA)
   SQLX$ = SQLX$ + ",UltUpdate =   '" & fech & "'"
   FLEXCONN.Execute (SQLX$)

End Sub
Function Uso_de_componente#(CODICONJ%, codielem%)
     'DEVUELVE EL USO UNITARIO SEGUN EL CODICONJUNTO Y EL ELEMENTO
     ' **** OJO ACUMULA SI DENTRO DE LA FORMULA DE CODICONJ ESTA REPETIDO ****
     CONDI$ = "CodIConj = " & CODICONJ%
     CONDI$ = CONDI$ + " AND CodiElem = " & codielem%
     Uso_de_componente# = XVALO(VALOBADA("FORMULAS", "SUM(USOBRUTO)", CONDI$, "NOMESS"))
End Function

Function AgregarRegistroDetFacomSql(CI1%, LOTE$)
'AHP
'CARGA UN REGISTRO CON LA CANTIDAD APROBADA PARA QUE NO QUEDE PENDIENTE DE FACTURACION EN LA CARGA POR DETALLE DE PROVEEDOR
'CON CODIGO, COD_INTE, IDENLOTE, CANTFACTURADA = CANTIDAD RECIBIDA DEL BOLRECEP, EL NUMEBORE, FECHEMISI, FECHCONTA
',CANTISEGUNREMITO,CANTIAPROBADA,CANTPENDFACT=0,MARBORRA =0,NUME_CLI,NUME_OCOM
        '
        'VULEVO A VERIFICAR SI NO SE INGRESO A DETFACOMSQL
        SQLX$ = "SELECT Cod_Inte, idenlote,FECHRECEP FROM BOLRECEP B WITH(NOLOCK)"
        SQLX$ = SQLX$ + " WHERE NOT EXISTS (SELECT 1 FROM DETFACOMSQL D"
        SQLX$ = SQLX$ + " Where B.Cod_Inte = D.Cod_Inte And B.idenlote = D.idenlote) AND B.IDENLOTE = '" & LOTE$ & "' AND B.COD_INTE =" & CI1%
        SQLX$ = SQLX$ + " ORDER BY FECHRECEP DESC"
        Set rs = READSET(SQLX$)
        rs.Close
        Set rs = Nothing
        
        SQLX$ = "SELECT * FROM BOLRECEP WITH(NOLOCK)"
        SQLX$ = SQLX$ + " WHERE IDENLOTE = '" & LOTE$ & "'"
        Set rs = READSET(SQLX$)
        With rs
           If Not .EOF Then
            CANTFACTURADAX = XVALO(!CANTOAPRO)
            NUMEBOREX = XVALO(!NUMEBORE)
            FECHEMISIX = CVDAT(CVINT(!FECHRECEP))
            NUME_CLIX = XVALO(!NumProv)
            NUME_OCOMX = XVALO(!NUMOC)
           End If
        End With
        rs.Close
        Set rs = Nothing
        
        If CANTFACTURADAX > 0 Then  'SI ENTRO EN LA CONSULTA DE ARRIBA CANTFACTURADAX = XVALO(!CANTOAPRO),,, CANTFACTURADAX DEBE SER > 0 , SI NO NO INSERTA EN DETFACOMSQL
            SQLX$ = "INSERT INTO DETFACOMSQL"
            SQLX$ = SQLX$ + " (COD_INTE,IDENLOTE,CANTFACTURADA,NUMEBORE,FECHEMISI,FECHCONTA,CANTISEGUNREMITO,CANTIAPROBADA,CANTPENDFACT,MARBORRA,NUME_CLI,NUME_OCOM)"
            
            SQLX$ = SQLX$ + " VALUES(" & CI1%
            SQLX$ = SQLX$ + " ,'" & LOTE$ & "'"
            SQLX$ = SQLX$ + " ," & CANTFACTURADAX
            SQLX$ = SQLX$ + ", " & NUMEBOREX
            SQLX$ = SQLX$ + " ,'" & FECHEMISIX & "'"
            SQLX$ = SQLX$ + " ,'" & FECHEMISIX & "'"
            SQLX$ = SQLX$ + ", " & CANTFACTURADAX
            SQLX$ = SQLX$ + ", " & CANTFACTURADAX
            SQLX$ = SQLX$ + ", 0"
            SQLX$ = SQLX$ + ", 0"
            SQLX$ = SQLX$ + ", " & NUME_CLIX
            SQLX$ = SQLX$ + ", " & NUME_OCOMX
            SQLX$ = SQLX$ + ")"
            FLEXCONN.Execute (SQLX$)
       End If

End Function
