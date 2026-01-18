Attribute VB_Name = "EscribeEstrucSql"

Sub GENERAR_ESTRUCTURAS_SQL()
   Call ABRECONEXION
   Call CREATABLA_ENCAPRF
   EPAC$ = TRIM$(UCase$(EMPREAC$))
   Call GENERAR_ESTRUCTURAS1
   Call GENERAR_ESTRUCTURAS2
   
   If InStr(1, EPAC$, "CONTEPLAST") > 0 Or InStr(1, EPAC$, "CAMBACERES") > 0 Or InStr(1, EPAC$, "PERGAMINO") > 0 Then
     Call GENERAR_ESTRUCTURAS_CTP
   End If
   '
End Sub
Sub GENERAR_ESTRUCTURAS_CTP()
       '|INICIO TIPOTELACMB
        '
        CONDI$ = " LISTA =   'TIPOTELACMB'"
        EXILISTA% = CantRegistros&("ENCAPRFUN", CONDI$, "NOMESS")
        If EXILISTA% > 0 Then
          FE% = 16350
          A = FECHATEX$(FE%)
          
          FEBD% = CVINT(VALOBADA("ENCAPRFUN", "FECHEMISI", CONDI$, "NOMESS"))
          GENERAREST% = 0
          If FEBD% < FE% Then
            IDLISTAX& = XVALO(VALOBADA("ENCAPRFUN", "ID_LISTA", CONDI$, "NOMESS"))
            Call BORRAREGISTROS("ENCAPRFUN", "ID_LISTA = " & IDLISTAX&, REGAF&, "NOMESS")
            Call BORRAREGISTROS("ENCAPRF", "ID_LISTA = " & IDLISTAX&, REGAF&, "NOMESS")
            GENERAREST% = 1
          End If
        Else
          GENERAREST% = 1
          CONDI$ = " ID_LISTA  > 0 "
          IDLISTAX& = 1 + XVALO(VALOBADA("ENCAPRFUN", "MAX(ID_LISTA)", CONDI$, "NOMESS"))
        End If
        If GENERAREST% > 0 Then
          SQLX$ = "INSERT INTO  ENCAPRFUN (ID_LISTA ,LISTA, Descripcion, ADDREGISTROS, USAFORMULA, BORRAR, INSERTAR, NOTILDE "
          SQLX$ = SQLX$ + ", CANTICAMPOS, CodiEmpr, USARECORD, FECHEMISI, TITURECORD, ConsultaSql)"
          SQLX$ = SQLX$ + "VALUES ( " & IDLISTAX&
          SQLX$ = SQLX$ + ",'TIPOTELACMB'"
          SQLX$ = SQLX$ + ",'TIPO DE TELA CAMBACERES Y REFILES(TIPOTELACMB)'"
          SQLX$ = SQLX$ + ",1"
          SQLX$ = SQLX$ + ",1"
          SQLX$ = SQLX$ + " ,0"
          SQLX$ = SQLX$ + " ,0"
          SQLX$ = SQLX$ + " ,0"
          SQLX$ = SQLX$ + " ,3"
          SQLX$ = SQLX$ + " ,0"
          SQLX$ = SQLX$ + " ,0"
          SQLX$ = SQLX$ + " ,'30/03/19'"
          SQLX$ = SQLX$ + " ,''"
          SQLX$ = SQLX$ + " ,'SELECT CAMPO1, CAMPO2,CAMPO3,CAMPO4,CAMPO5 FROM DATASQL WITH(NOLOCK)'"
          SQLX$ = SQLX$ + ")"
          FLEXCONN.Execute (SQLX$)
            SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
            SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
            SQLX$ = SQLX$ + " ,0"
            SQLX$ = SQLX$ + " ,0"
            SQLX$ = SQLX$ + " , 'CODIGO/F4'"
            SQLX$ = SQLX$ + " , '     '"
            SQLX$ = SQLX$ + " , 'S'"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " )"
            FLEXCONN.Execute (SQLX$)
        
        '
            SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
            SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
            SQLX$ = SQLX$ + " ,0"
            SQLX$ = SQLX$ + " ,0"
            SQLX$ = SQLX$ + " , 'Tipo de Tela/A24'"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " , 'S'"
            SQLX$ = SQLX$ + " , '     '"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " )"
            FLEXCONN.Execute (SQLX$)
        
            SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
            SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
            SQLX$ = SQLX$ + " ,0"
            SQLX$ = SQLX$ + " ,0"
            SQLX$ = SQLX$ + " , 'Refiles/A7'"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " , 'S'"
            SQLX$ = SQLX$ + " , '     '"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " )"
            FLEXCONN.Execute (SQLX$)
        
            SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
            SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
            SQLX$ = SQLX$ + " ,0"
            SQLX$ = SQLX$ + " ,0"
            SQLX$ = SQLX$ + " , 'Valor gr-m/F12.1'"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " , 'S'"
            SQLX$ = SQLX$ + " , '     '"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " )"
            FLEXCONN.Execute (SQLX$)
            
            SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
            SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
            SQLX$ = SQLX$ + " ,0"
            SQLX$ = SQLX$ + " ,0"
            SQLX$ = SQLX$ + " , 'Factor/F7.0'"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " , 'S'"
            SQLX$ = SQLX$ + " , '     '"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " )"
            FLEXCONN.Execute (SQLX$)
        '
        End If
        '|FIN TIPOTELACMB
'|INICIO DIAMECOR
'
CONDI$ = " LISTA =   'DIAMECOR'"
EXILISTA% = CantRegistros&("ENCAPRFUN", CONDI$, "NOMESS")
If EXILISTA% > 0 Then
  FE% = 16961
  FEBD% = CVINT(VALOBADA("ENCAPRFUN", "FECHEMISI", CONDI$, "NOMESS"))
  GENERAREST% = 0
  If FEBD% < FE% Then
    IDLISTAX& = XVALO(VALOBADA("ENCAPRFUN", "ID_LISTA", CONDI$, "NOMESS"))
    Call BORRAREGISTROS("ENCAPRFUN", "ID_LISTA = " & IDLISTAX&, REGAF&, "NOMESS")
    Call BORRAREGISTROS("ENCAPRF", "ID_LISTA = " & IDLISTAX&, REGAF&, "NOMESS")
    GENERAREST% = 1
  End If
Else
  GENERAREST% = 1
  CONDI$ = " ID_LISTA  > 0 "
  IDLISTAX& = 1 + XVALO(VALOBADA("ENCAPRFUN", "MAX(ID_LISTA)", CONDI$, "NOMESS"))
End If
If GENERAREST% > 0 Then
  SQLX$ = "INSERT INTO  ENCAPRFUN (ID_LISTA ,LISTA, Descripcion, ADDREGISTROS, USAFORMULA, BORRAR, INSERTAR, NOTILDE "
  SQLX$ = SQLX$ + ", CANTICAMPOS, CodiEmpr, USARECORD, FECHEMISI, TITURECORD, ConsultaSql)"
  SQLX$ = SQLX$ + "VALUES ( " & IDLISTAX&
  SQLX$ = SQLX$ + ",'DIAMECOR'"
  SQLX$ = SQLX$ + ",'DIAMETROS PARA O.F. CORTES'"
  SQLX$ = SQLX$ + ",1"
  SQLX$ = SQLX$ + ",1"
  SQLX$ = SQLX$ + " ,1"
  SQLX$ = SQLX$ + " ,1"
  SQLX$ = SQLX$ + " ,0"
  SQLX$ = SQLX$ + " ,4"
  SQLX$ = SQLX$ + " ,0"
  SQLX$ = SQLX$ + " ,0"
  SQLX$ = SQLX$ + " ,'01/10/20'"
  SQLX$ = SQLX$ + " ,''"
  SQLX$ = SQLX$ + " ,'SELECT CAMPO1,CAMPO2 FROM DATASQL WITH(NOLOCK)'"
  SQLX$ = SQLX$ + ")"
  FLEXCONN.Execute (SQLX$)
    SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
    SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
    SQLX$ = SQLX$ + " ,0"
    SQLX$ = SQLX$ + " ,0"
    SQLX$ = SQLX$ + " , 'CODIGO/F8'"
    SQLX$ = SQLX$ + " , ' '"
    SQLX$ = SQLX$ + " , 'S'"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " )"
    FLEXCONN.Execute (SQLX$)


'
    SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
    SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
    SQLX$ = SQLX$ + " ,0"
    SQLX$ = SQLX$ + " ,0"
    SQLX$ = SQLX$ + " , 'DIAMETRO (CM.)/A14'"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , 'S'"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " )"
    FLEXCONN.Execute (SQLX$)

'
End If
'|FIN DIAMECOR




'|INICIO UNIDPAQU
'
CONDI$ = " LISTA =   'UNIDPAQU'"
EXILISTA% = CantRegistros&("ENCAPRFUN", CONDI$, "NOMESS")
If EXILISTA% > 0 Then
  FE% = 16899
  FEBD% = CVINT(VALOBADA("ENCAPRFUN", "FECHEMISI", CONDI$, "NOMESS"))
  GENERAREST% = 0
  If FEBD% < FE% Then
    IDLISTAX& = XVALO(VALOBADA("ENCAPRFUN", "ID_LISTA", CONDI$, "NOMESS"))
    Call BORRAREGISTROS("ENCAPRFUN", "ID_LISTA = " & IDLISTAX&, REGAF&, "NOMESS")
    Call BORRAREGISTROS("ENCAPRF", "ID_LISTA = " & IDLISTAX&, REGAF&, "NOMESS")
    GENERAREST% = 1
  End If
Else
  GENERAREST% = 1
  CONDI$ = " ID_LISTA  > 0 "
  IDLISTAX& = 1 + XVALO(VALOBADA("ENCAPRFUN", "MAX(ID_LISTA)", CONDI$, "NOMESS"))
End If
If GENERAREST% > 0 Then
  SQLX$ = "INSERT INTO  ENCAPRFUN (ID_LISTA ,LISTA, Descripcion, ADDREGISTROS, USAFORMULA, BORRAR, INSERTAR, NOTILDE "
  SQLX$ = SQLX$ + ", CANTICAMPOS, CodiEmpr, USARECORD, FECHEMISI, TITURECORD, ConsultaSql)"
  SQLX$ = SQLX$ + "VALUES ( " & IDLISTAX&
  SQLX$ = SQLX$ + ",'UNIDPAQU'"
  SQLX$ = SQLX$ + ",'UNID. X PAQUETES FICHA TÉCNICA BOLSAS'"
  SQLX$ = SQLX$ + ",1"
  SQLX$ = SQLX$ + ",1"
  SQLX$ = SQLX$ + " ,0"
  SQLX$ = SQLX$ + " ,0"
  SQLX$ = SQLX$ + " ,0"
  SQLX$ = SQLX$ + " ,3"
  SQLX$ = SQLX$ + " ,0"
  SQLX$ = SQLX$ + " ,0"
  SQLX$ = SQLX$ + " ,'03/08/20'"
  SQLX$ = SQLX$ + " ,''"
  SQLX$ = SQLX$ + " ,'SELECT CAMPO1,CAMPO2 FROM DATASQL WITH(NOLOCK)'"
  SQLX$ = SQLX$ + ")"
  FLEXCONN.Execute (SQLX$)
    SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
    SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
    SQLX$ = SQLX$ + " ,0"
    SQLX$ = SQLX$ + " ,0"
    SQLX$ = SQLX$ + " , 'Código/f8'"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " )"
    FLEXCONN.Execute (SQLX$)

'
    SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
    SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
    SQLX$ = SQLX$ + " ,0"
    SQLX$ = SQLX$ + " ,0"
    SQLX$ = SQLX$ + " , 'Referencia/a48'"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , 'S'"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " )"
    FLEXCONN.Execute (SQLX$)

'
End If
'|FIN UNIDPAQU





'|INICIO TIPOHILO
'
CONDI$ = " LISTA =   'TIPOHILO'"
EXILISTA% = CantRegistros&("ENCAPRFUN", CONDI$, "NOMESS")
If EXILISTA% > 0 Then
  FE% = 16899
  FEBD% = CVINT(VALOBADA("ENCAPRFUN", "FECHEMISI", CONDI$, "NOMESS"))
  GENERAREST% = 0
  If FEBD% < FE% Then
    IDLISTAX& = XVALO(VALOBADA("ENCAPRFUN", "ID_LISTA", CONDI$, "NOMESS"))
    Call BORRAREGISTROS("ENCAPRFUN", "ID_LISTA = " & IDLISTAX&, REGAF&, "NOMESS")
    Call BORRAREGISTROS("ENCAPRF", "ID_LISTA = " & IDLISTAX&, REGAF&, "NOMESS")
    GENERAREST% = 1
  End If
Else
  GENERAREST% = 1
  CONDI$ = " ID_LISTA  > 0 "
  IDLISTAX& = 1 + XVALO(VALOBADA("ENCAPRFUN", "MAX(ID_LISTA)", CONDI$, "NOMESS"))
End If
If GENERAREST% > 0 Then
  SQLX$ = "INSERT INTO  ENCAPRFUN (ID_LISTA ,LISTA, Descripcion, ADDREGISTROS, USAFORMULA, BORRAR, INSERTAR, NOTILDE "
  SQLX$ = SQLX$ + ", CANTICAMPOS, CodiEmpr, USARECORD, FECHEMISI, TITURECORD, ConsultaSql)"
  SQLX$ = SQLX$ + "VALUES ( " & IDLISTAX&
  SQLX$ = SQLX$ + ",'TIPOHILO'"
  SQLX$ = SQLX$ + ",'TIPO HILO FICHA TÉCNICA BOLSAS '"
  SQLX$ = SQLX$ + ",1"
  SQLX$ = SQLX$ + ",1"
  SQLX$ = SQLX$ + " ,0"
  SQLX$ = SQLX$ + " ,0"
  SQLX$ = SQLX$ + " ,0"
  SQLX$ = SQLX$ + " ,3"
  SQLX$ = SQLX$ + " ,0"
  SQLX$ = SQLX$ + " ,0"
  SQLX$ = SQLX$ + " ,'03/08/20'"
  SQLX$ = SQLX$ + " ,''"
  SQLX$ = SQLX$ + " ,'SELECT CAMPO1,CAMPO2 FROM DATASQL WITH(NOLOCK)'"
  SQLX$ = SQLX$ + ")"
  FLEXCONN.Execute (SQLX$)
    SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
    SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
    SQLX$ = SQLX$ + " ,0"
    SQLX$ = SQLX$ + " ,0"
    SQLX$ = SQLX$ + " , 'Código/f8'"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " )"
    FLEXCONN.Execute (SQLX$)

'
    SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
    SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
    SQLX$ = SQLX$ + " ,0"
    SQLX$ = SQLX$ + " ,0"
    SQLX$ = SQLX$ + " , 'Referencia/a48'"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , 'S'"
    SQLX$ = SQLX$ + " , ' '"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " )"
    FLEXCONN.Execute (SQLX$)

'
End If
'|FIN TIPOHILO

'|INICIO HILOMATERIAL
'
CONDI$ = " LISTA =   'HILOMATERIAL'"
EXILISTA% = CantRegistros&("ENCAPRFUN", CONDI$, "NOMESS")
If EXILISTA% > 0 Then
  FE% = 16899
  FEBD% = CVINT(VALOBADA("ENCAPRFUN", "FECHEMISI", CONDI$, "NOMESS"))
  GENERAREST% = 0
  If FEBD% < FE% Then
    IDLISTAX& = XVALO(VALOBADA("ENCAPRFUN", "ID_LISTA", CONDI$, "NOMESS"))
    Call BORRAREGISTROS("ENCAPRFUN", "ID_LISTA = " & IDLISTAX&, REGAF&, "NOMESS")
    Call BORRAREGISTROS("ENCAPRF", "ID_LISTA = " & IDLISTAX&, REGAF&, "NOMESS")
    GENERAREST% = 1
  End If
Else
  GENERAREST% = 1
  CONDI$ = " ID_LISTA  > 0 "
  IDLISTAX& = 1 + XVALO(VALOBADA("ENCAPRFUN", "MAX(ID_LISTA)", CONDI$, "NOMESS"))
End If
If GENERAREST% > 0 Then
  SQLX$ = "INSERT INTO  ENCAPRFUN (ID_LISTA ,LISTA, Descripcion, ADDREGISTROS, USAFORMULA, BORRAR, INSERTAR, NOTILDE "
  SQLX$ = SQLX$ + ", CANTICAMPOS, CodiEmpr, USARECORD, FECHEMISI, TITURECORD, ConsultaSql)"
  SQLX$ = SQLX$ + "VALUES ( " & IDLISTAX&
  SQLX$ = SQLX$ + ",'HILOMATERIAL'"
  SQLX$ = SQLX$ + ",'HILOMATERIAL FICHA TÉCNICA BOLSAS '"
  SQLX$ = SQLX$ + ",1"
  SQLX$ = SQLX$ + ",1"
  SQLX$ = SQLX$ + " ,0"
  SQLX$ = SQLX$ + " ,0"
  SQLX$ = SQLX$ + " ,0"
  SQLX$ = SQLX$ + " ,3"
  SQLX$ = SQLX$ + " ,0"
  SQLX$ = SQLX$ + " ,0"
  SQLX$ = SQLX$ + " ,'03/08/20'"
  SQLX$ = SQLX$ + " ,''"
  SQLX$ = SQLX$ + " ,'SELECT CAMPO1,CAMPO2 FROM DATASQL WITH(NOLOCK)'"
  SQLX$ = SQLX$ + ")"
  FLEXCONN.Execute (SQLX$)
    SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
    SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
    SQLX$ = SQLX$ + " ,0"
    SQLX$ = SQLX$ + " ,0"
    SQLX$ = SQLX$ + " , 'Código/f8'"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " )"
    FLEXCONN.Execute (SQLX$)

'
    SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
    SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
    SQLX$ = SQLX$ + " ,0"
    SQLX$ = SQLX$ + " ,0"
    SQLX$ = SQLX$ + " , 'Referencia/a48'"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , 'S'"
    SQLX$ = SQLX$ + " , ' '"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " )"
    FLEXCONN.Execute (SQLX$)

'
End If
'|FIN HILOMATERIAL



'|INICIO SISDESCA
'
CONDI$ = " LISTA =   'SISDESCA'"
EXILISTA% = CantRegistros&("ENCAPRFUN", CONDI$, "NOMESS")
If EXILISTA% > 0 Then
  FE% = 16899
  FEBD% = CVINT(VALOBADA("ENCAPRFUN", "FECHEMISI", CONDI$, "NOMESS"))
  GENERAREST% = 0
  If FEBD% < FE% Then
    IDLISTAX& = XVALO(VALOBADA("ENCAPRFUN", "ID_LISTA", CONDI$, "NOMESS"))
    Call BORRAREGISTROS("ENCAPRFUN", "ID_LISTA = " & IDLISTAX&, REGAF&, "NOMESS")
    Call BORRAREGISTROS("ENCAPRF", "ID_LISTA = " & IDLISTAX&, REGAF&, "NOMESS")
    GENERAREST% = 1
  End If
Else
  GENERAREST% = 1
  CONDI$ = " ID_LISTA  > 0 "
  IDLISTAX& = 1 + XVALO(VALOBADA("ENCAPRFUN", "MAX(ID_LISTA)", CONDI$, "NOMESS"))
End If
If GENERAREST% > 0 Then
  SQLX$ = "INSERT INTO  ENCAPRFUN (ID_LISTA ,LISTA, Descripcion, ADDREGISTROS, USAFORMULA, BORRAR, INSERTAR, NOTILDE "
  SQLX$ = SQLX$ + ", CANTICAMPOS, CodiEmpr, USARECORD, FECHEMISI, TITURECORD, ConsultaSql)"
  SQLX$ = SQLX$ + "VALUES ( " & IDLISTAX&
  SQLX$ = SQLX$ + ",'SISDESCA'"
  SQLX$ = SQLX$ + ",'SISTEMA DE DESCARGA FICHA TECNICA BOLSAS'"
  SQLX$ = SQLX$ + ",1"
  SQLX$ = SQLX$ + ",1"
  SQLX$ = SQLX$ + " ,0"
  SQLX$ = SQLX$ + " ,0"
  SQLX$ = SQLX$ + " ,0"
  SQLX$ = SQLX$ + " ,3"
  SQLX$ = SQLX$ + " ,0"
  SQLX$ = SQLX$ + " ,0"
  SQLX$ = SQLX$ + " ,'03/08/20'"
  SQLX$ = SQLX$ + " ,''"
  SQLX$ = SQLX$ + " ,'SELECT CAMPO1,CAMPO2 FROM DATASQL WITH(NOLOCK)'"
  SQLX$ = SQLX$ + ")"
  FLEXCONN.Execute (SQLX$)
    SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
    SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
    SQLX$ = SQLX$ + " ,0"
    SQLX$ = SQLX$ + " ,0"
    SQLX$ = SQLX$ + " , 'Código/f8'"
    SQLX$ = SQLX$ + " , ' '"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " )"
    FLEXCONN.Execute (SQLX$)

'
    SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
    SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
    SQLX$ = SQLX$ + " ,0"
    SQLX$ = SQLX$ + " ,0"
    SQLX$ = SQLX$ + " , 'Referencia/a48'"
    SQLX$ = SQLX$ + " , ' '"
    SQLX$ = SQLX$ + " , 's'"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " )"
    FLEXCONN.Execute (SQLX$)

'
End If
'|FIN SISDESCA


'|INICIO SisCarga
'
CONDI$ = " LISTA =   'SisCarga'"
EXILISTA% = CantRegistros&("ENCAPRFUN", CONDI$, "NOMESS")
If EXILISTA% > 0 Then
  FE% = 16899
  FEBD% = CVINT(VALOBADA("ENCAPRFUN", "FECHEMISI", CONDI$, "NOMESS"))
  GENERAREST% = 0
  If FEBD% < FE% Then
    IDLISTAX& = XVALO(VALOBADA("ENCAPRFUN", "ID_LISTA", CONDI$, "NOMESS"))
    Call BORRAREGISTROS("ENCAPRFUN", "ID_LISTA = " & IDLISTAX&, REGAF&, "NOMESS")
    Call BORRAREGISTROS("ENCAPRF", "ID_LISTA = " & IDLISTAX&, REGAF&, "NOMESS")
    GENERAREST% = 1
  End If
Else
  GENERAREST% = 1
  CONDI$ = " ID_LISTA  > 0 "
  IDLISTAX& = 1 + XVALO(VALOBADA("ENCAPRFUN", "MAX(ID_LISTA)", CONDI$, "NOMESS"))
End If
If GENERAREST% > 0 Then
  SQLX$ = "INSERT INTO  ENCAPRFUN (ID_LISTA ,LISTA, Descripcion, ADDREGISTROS, USAFORMULA, BORRAR, INSERTAR, NOTILDE "
  SQLX$ = SQLX$ + ", CANTICAMPOS, CodiEmpr, USARECORD, FECHEMISI, TITURECORD, ConsultaSql)"
  SQLX$ = SQLX$ + "VALUES ( " & IDLISTAX&
  SQLX$ = SQLX$ + ",'SisCarga'"
  SQLX$ = SQLX$ + ",'SISTEMA DE CARGA  FICHA TÉCNICA BOLSAS'"
  SQLX$ = SQLX$ + ",1"
  SQLX$ = SQLX$ + ",1"
  SQLX$ = SQLX$ + " ,0"
  SQLX$ = SQLX$ + " ,0"
  SQLX$ = SQLX$ + " ,0"
  SQLX$ = SQLX$ + " ,3"
  SQLX$ = SQLX$ + " ,0"
  SQLX$ = SQLX$ + " ,0"
  SQLX$ = SQLX$ + " ,'03/08/20'"
  SQLX$ = SQLX$ + " ,''"
  SQLX$ = SQLX$ + " ,'SELECT CAMPO1,CAMPO2 FROM DATASQL WITH(NOLOCK)'"
  SQLX$ = SQLX$ + ")"
  FLEXCONN.Execute (SQLX$)
    SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
    SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
    SQLX$ = SQLX$ + " ,0"
    SQLX$ = SQLX$ + " ,0"
    SQLX$ = SQLX$ + " , 'Código/f8'"
    SQLX$ = SQLX$ + " , ' '"
    SQLX$ = SQLX$ + " , ' '"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " )"
    FLEXCONN.Execute (SQLX$)

'
    SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
    SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
    SQLX$ = SQLX$ + " ,0"
    SQLX$ = SQLX$ + " ,0"
    SQLX$ = SQLX$ + " , 'Referencia/a32'"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , 'S'"
    SQLX$ = SQLX$ + " , ' '"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " )"
    FLEXCONN.Execute (SQLX$)

'
End If
'|FIN SisCarga

        '|INICIO CONFUBPT
        '
        CONDI$ = " LISTA =   'CONFUBPT'"
        EXILISTA% = CantRegistros&("ENCAPRFUN", CONDI$, "NOMESS")
        If EXILISTA% > 0 Then
          FE% = 16759
          FEBD% = CVINT(VALOBADA("ENCAPRFUN", "FECHEMISI", CONDI$, "NOMESS"))
          GENERAREST% = 0
          If FEBD% < FE% Then
            IDLISTAX& = XVALO(VALOBADA("ENCAPRFUN", "ID_LISTA", CONDI$, "NOMESS"))
            Call BORRAREGISTROS("ENCAPRFUN", "ID_LISTA = " & IDLISTAX&, REGAF&, "NOMESS")
            Call BORRAREGISTROS("ENCAPRF", "ID_LISTA = " & IDLISTAX&, REGAF&, "NOMESS")
            GENERAREST% = 1
          End If
        Else
          GENERAREST% = 1
          CONDI$ = " ID_LISTA  > 0 "
          IDLISTAX& = 1 + XVALO(VALOBADA("ENCAPRFUN", "MAX(ID_LISTA)", CONDI$, "NOMESS"))
        End If
        If GENERAREST% > 0 Then
          SQLX$ = "INSERT INTO  ENCAPRFUN (ID_LISTA ,LISTA, Descripcion, ADDREGISTROS, USAFORMULA, BORRAR, INSERTAR, NOTILDE "
          SQLX$ = SQLX$ + ", CANTICAMPOS, CodiEmpr, USARECORD, FECHEMISI, TITURECORD, ConsultaSql)"
          SQLX$ = SQLX$ + "VALUES ( " & IDLISTAX&
          SQLX$ = SQLX$ + ",'CONFUBPT'"
          SQLX$ = SQLX$ + ",'Config. Ubic.Productos Terminados'"
          SQLX$ = SQLX$ + ",1"
          SQLX$ = SQLX$ + ",1"
          SQLX$ = SQLX$ + " ,0"
          SQLX$ = SQLX$ + " ,1"
          SQLX$ = SQLX$ + " ,0"
          SQLX$ = SQLX$ + " ,4"
          SQLX$ = SQLX$ + " ,0"
          SQLX$ = SQLX$ + " ,0"
          SQLX$ = SQLX$ + " ,'23/03/20'"
          SQLX$ = SQLX$ + " ,''"
          SQLX$ = SQLX$ + " ,'SELECT CAMPO1, CAMPO2,CAMPO3,CAMPO4 FROM DATASQL WITH(NOLOCK) '"
          SQLX$ = SQLX$ + ")"
          FLEXCONN.Execute (SQLX$)
          
          SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
          SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
          SQLX$ = SQLX$ + " ,0"
          SQLX$ = SQLX$ + " ,0"
          SQLX$ = SQLX$ + " , 'N° DE RACK/F8'"
          SQLX$ = SQLX$ + " , ' '"
          SQLX$ = SQLX$ + " , 'S'"
          SQLX$ = SQLX$ + " , ' '"
          SQLX$ = SQLX$ + " , ' '"
          SQLX$ = SQLX$ + " , ' '"
          SQLX$ = SQLX$ + " , ' '"
          SQLX$ = SQLX$ + " )"
          FLEXCONN.Execute (SQLX$)
            
            SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
            SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
            SQLX$ = SQLX$ + " ,0"
            SQLX$ = SQLX$ + " ,0"
            SQLX$ = SQLX$ + " , 'DESCRIPCION/A32'"
            SQLX$ = SQLX$ + " , ' '"
            SQLX$ = SQLX$ + " , 'S'"
            SQLX$ = SQLX$ + " , ' '"
            SQLX$ = SQLX$ + " , ' '"
            SQLX$ = SQLX$ + " , ' '"
            SQLX$ = SQLX$ + " , ' '"
            SQLX$ = SQLX$ + " )"
            FLEXCONN.Execute (SQLX$)
            '
            SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
            SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
            SQLX$ = SQLX$ + " ,0"
            SQLX$ = SQLX$ + " ,0"
            SQLX$ = SQLX$ + " , 'CANT.FILAS./F9'"
            SQLX$ = SQLX$ + " , ' '"
            SQLX$ = SQLX$ + " , 'S'"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " )"
            FLEXCONN.Execute (SQLX$)
        
            '
            SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
            SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
            SQLX$ = SQLX$ + " ,0"
            SQLX$ = SQLX$ + " ,0"
            SQLX$ = SQLX$ + " , 'CANT.COL./F9'"
            SQLX$ = SQLX$ + " , ' '"
            SQLX$ = SQLX$ + " , 'S'"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " )"
            FLEXCONN.Execute (SQLX$)
        
        '
        End If
        '|FIN CONFUBPT


'|INICIO OPERFUSI
'
CONDI$ = " LISTA =   'OPERFUSI'"
EXILISTA% = CantRegistros&("ENCAPRFUN", CONDI$, "NOMESS")
If EXILISTA% > 0 Then
  FE% = 16841
  FEBD% = CVINT(VALOBADA("ENCAPRFUN", "FECHEMISI", CONDI$, "NOMESS"))
  GENERAREST% = 0
  If FEBD% < FE% Then
    IDLISTAX& = XVALO(VALOBADA("ENCAPRFUN", "ID_LISTA", CONDI$, "NOMESS"))
    Call BORRAREGISTROS("ENCAPRFUN", "ID_LISTA = " & IDLISTAX&, REGAF&, "NOMESS")
    Call BORRAREGISTROS("ENCAPRF", "ID_LISTA = " & IDLISTAX&, REGAF&, "NOMESS")
    GENERAREST% = 1
  End If
Else
  GENERAREST% = 1
  CONDI$ = " ID_LISTA  > 0 "
  IDLISTAX& = 1 + XVALO(VALOBADA("ENCAPRFUN", "MAX(ID_LISTA)", CONDI$, "NOMESS"))
End If
If GENERAREST% > 0 Then
  SQLX$ = "INSERT INTO  ENCAPRFUN (ID_LISTA ,LISTA, Descripcion, ADDREGISTROS, USAFORMULA, BORRAR, INSERTAR, NOTILDE "
  SQLX$ = SQLX$ + ", CANTICAMPOS, CodiEmpr, USARECORD, FECHEMISI, TITURECORD, ConsultaSql)"
  SQLX$ = SQLX$ + "VALUES ( " & IDLISTAX&
  SQLX$ = SQLX$ + ",'OPERFUSI'"
  SQLX$ = SQLX$ + ",'OPERACIONES FUSIONADAS'"
  SQLX$ = SQLX$ + ",1"
  SQLX$ = SQLX$ + ",1"
  SQLX$ = SQLX$ + " ,0"
  SQLX$ = SQLX$ + " ,0"
  SQLX$ = SQLX$ + " ,0"
  SQLX$ = SQLX$ + " ,4"
  SQLX$ = SQLX$ + " ,0"
  SQLX$ = SQLX$ + " ,0"
  SQLX$ = SQLX$ + " ,'09/06/20'"
  SQLX$ = SQLX$ + " ,''"
  SQLX$ = SQLX$ + " ,'SELECT CAMPO1,CAMPO2,CAMPO3 FROM DATASQL WITH(NOLOCK) '"
  SQLX$ = SQLX$ + ")"
  FLEXCONN.Execute (SQLX$)
    SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
    SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
    SQLX$ = SQLX$ + " ,0"
    SQLX$ = SQLX$ + " ,0"
    SQLX$ = SQLX$ + " , 'CODIGO/F8'"
    SQLX$ = SQLX$ + " , ' '"
    SQLX$ = SQLX$ + " , 'S'"
    SQLX$ = SQLX$ + " , ' '"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " )"
    FLEXCONN.Execute (SQLX$)

'
    SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
    SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
    SQLX$ = SQLX$ + " ,0"
    SQLX$ = SQLX$ + " ,0"
    SQLX$ = SQLX$ + " , 'DESCRIPCION/A48'"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , 'S'"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " )"
    FLEXCONN.Execute (SQLX$)

'
    SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
    SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
    SQLX$ = SQLX$ + " ,0"
    SQLX$ = SQLX$ + " ,0"
    SQLX$ = SQLX$ + " , 'SV/A0'"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " )"
    FLEXCONN.Execute (SQLX$)

'
End If
'|FIN OPERFUSI


'|INICIO CATESCRP
'
CONDI$ = " LISTA =   'CATESCRP'"
EXILISTA% = CantRegistros&("ENCAPRFUN", CONDI$, "NOMESS")
If EXILISTA% > 0 Then
  FE% = 16829
  FEBD% = CVINT(VALOBADA("ENCAPRFUN", "FECHEMISI", CONDI$, "NOMESS"))
  GENERAREST% = 0
  If FEBD% < FE% Then
    IDLISTAX& = XVALO(VALOBADA("ENCAPRFUN", "ID_LISTA", CONDI$, "NOMESS"))
    Call BORRAREGISTROS("ENCAPRFUN", "ID_LISTA = " & IDLISTAX&, REGAF&, "NOMESS")
    Call BORRAREGISTROS("ENCAPRF", "ID_LISTA = " & IDLISTAX&, REGAF&, "NOMESS")
    GENERAREST% = 1
  End If
Else
  GENERAREST% = 1
  CONDI$ = " ID_LISTA  > 0 "
  IDLISTAX& = 1 + XVALO(VALOBADA("ENCAPRFUN", "MAX(ID_LISTA)", CONDI$, "NOMESS"))
End If
If GENERAREST% > 0 Then
  SQLX$ = "INSERT INTO  ENCAPRFUN (ID_LISTA ,LISTA, Descripcion, ADDREGISTROS, USAFORMULA, BORRAR, INSERTAR, NOTILDE "
  SQLX$ = SQLX$ + ", CANTICAMPOS, CodiEmpr, USARECORD, FECHEMISI, TITURECORD, ConsultaSql)"
  SQLX$ = SQLX$ + "VALUES ( " & IDLISTAX&
  SQLX$ = SQLX$ + ",'CATESCRP'"
  SQLX$ = SQLX$ + ",'CATEGORIAS DE SCRAP'"
  SQLX$ = SQLX$ + ",1"
  SQLX$ = SQLX$ + ",1"
  SQLX$ = SQLX$ + " ,0"
  SQLX$ = SQLX$ + " ,0"
  SQLX$ = SQLX$ + " ,0"
  SQLX$ = SQLX$ + " ,4"
  SQLX$ = SQLX$ + " ,0"
  SQLX$ = SQLX$ + " ,0"
  SQLX$ = SQLX$ + " ,'29/05/20'"
  SQLX$ = SQLX$ + " ,''"
  SQLX$ = SQLX$ + " ,'SELECT CAMPO1,CAMPO2,CAMPO3 FROM DATASQL WITH(NOLOCK)'"
  SQLX$ = SQLX$ + ")"
  FLEXCONN.Execute (SQLX$)
    SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
    SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
    SQLX$ = SQLX$ + " ,0"
    SQLX$ = SQLX$ + " ,0"
    SQLX$ = SQLX$ + " , 'Código/f8'"
    SQLX$ = SQLX$ + " , ' '"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " )"
    FLEXCONN.Execute (SQLX$)

'
    SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
    SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
    SQLX$ = SQLX$ + " ,0"
    SQLX$ = SQLX$ + " ,0"
    SQLX$ = SQLX$ + " , 'CATEGORIA/A64'"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , 'S'"
    SQLX$ = SQLX$ + " , ' '"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " )"
    FLEXCONN.Execute (SQLX$)

'
    SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
    SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
    SQLX$ = SQLX$ + " ,0"
    SQLX$ = SQLX$ + " ,0"
    SQLX$ = SQLX$ + " , 'SV/A0'"
    SQLX$ = SQLX$ + " , ' '"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " )"
    FLEXCONN.Execute (SQLX$)

'
End If
'|FIN CATESCRP
'|INICIO SECTSCRP
'
CONDI$ = " LISTA =   'SECTSCRP'"
EXILISTA% = CantRegistros&("ENCAPRFUN", CONDI$, "NOMESS")
If EXILISTA% > 0 Then
  FE% = 16829
  FEBD% = CVINT(VALOBADA("ENCAPRFUN", "FECHEMISI", CONDI$, "NOMESS"))
  GENERAREST% = 0
  If FEBD% < FE% Then
    IDLISTAX& = XVALO(VALOBADA("ENCAPRFUN", "ID_LISTA", CONDI$, "NOMESS"))
    Call BORRAREGISTROS("ENCAPRFUN", "ID_LISTA = " & IDLISTAX&, REGAF&, "NOMESS")
    Call BORRAREGISTROS("ENCAPRF", "ID_LISTA = " & IDLISTAX&, REGAF&, "NOMESS")
    GENERAREST% = 1
  End If
Else
  GENERAREST% = 1
  CONDI$ = " ID_LISTA  > 0 "
  IDLISTAX& = 1 + XVALO(VALOBADA("ENCAPRFUN", "MAX(ID_LISTA)", CONDI$, "NOMESS"))
End If
If GENERAREST% > 0 Then

  SQLX$ = "INSERT INTO  ENCAPRFUN (ID_LISTA ,LISTA, Descripcion, ADDREGISTROS, USAFORMULA, BORRAR, INSERTAR, NOTILDE "
  SQLX$ = SQLX$ + ", CANTICAMPOS, CodiEmpr, USARECORD, FECHEMISI, TITURECORD, ConsultaSql)"
  SQLX$ = SQLX$ + "VALUES ( " & IDLISTAX&
  SQLX$ = SQLX$ + ",'SECTSCRP'"
  SQLX$ = SQLX$ + ",'SECTORES DE ORIGEN DE SCRAP'"
  SQLX$ = SQLX$ + ",1"
  SQLX$ = SQLX$ + ",1"
  SQLX$ = SQLX$ + " ,0"
  SQLX$ = SQLX$ + " ,0"
  SQLX$ = SQLX$ + " ,0"
  SQLX$ = SQLX$ + " ,4"
  SQLX$ = SQLX$ + " ,0"
  SQLX$ = SQLX$ + " ,0"
  SQLX$ = SQLX$ + " ,'29/05/20'"
  SQLX$ = SQLX$ + " ,''"
  SQLX$ = SQLX$ + " ,'SELECT CAMPO1,CAMPO2,CAMPO3 FROM DATASQL WITH(NOLOCK)'"
  SQLX$ = SQLX$ + ")"
  FLEXCONN.Execute (SQLX$)
    SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
    SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
    SQLX$ = SQLX$ + " ,0"
    SQLX$ = SQLX$ + " ,0"
    SQLX$ = SQLX$ + " , 'Código/F8'"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ' '"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " )"
    FLEXCONN.Execute (SQLX$)

'
    SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
    SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
    SQLX$ = SQLX$ + " ,0"
    SQLX$ = SQLX$ + " ,0"
    SQLX$ = SQLX$ + " , 'Sector/A32'"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , 'S'"
    SQLX$ = SQLX$ + " , ' '"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " )"
    FLEXCONN.Execute (SQLX$)

'
    SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
    SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
    SQLX$ = SQLX$ + " ,0"
    SQLX$ = SQLX$ + " ,0"
    SQLX$ = SQLX$ + " , 'SV/A0'"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " )"
    FLEXCONN.Execute (SQLX$)

'
End If
'|FIN SECTSCRP


'|INICIO SECTPLANTA
'
        CONDI$ = " LISTA =   'SECTPLANTA'"
        EXILISTA% = CantRegistros&("ENCAPRFUN", CONDI$, "NOMESS")
        If EXILISTA% > 0 Then
          FE% = 16788
          FEBD% = CVINT(VALOBADA("ENCAPRFUN", "FECHEMISI", CONDI$, "NOMESS"))
          GENERAREST% = 0
          If FEBD% < FE% Then
            IDLISTAX& = XVALO(VALOBADA("ENCAPRFUN", "ID_LISTA", CONDI$, "NOMESS"))
            Call BORRAREGISTROS("ENCAPRFUN", "ID_LISTA = " & IDLISTAX&, REGAF&, "NOMESS")
            Call BORRAREGISTROS("ENCAPRF", "ID_LISTA = " & IDLISTAX&, REGAF&, "NOMESS")
            GENERAREST% = 1
          End If
        Else
          GENERAREST% = 1
          CONDI$ = " ID_LISTA  > 0 "
          IDLISTAX& = 1 + XVALO(VALOBADA("ENCAPRFUN", "MAX(ID_LISTA)", CONDI$, "NOMESS"))
        End If
        If GENERAREST% > 0 Then
          SQLX$ = "INSERT INTO  ENCAPRFUN (ID_LISTA ,LISTA, Descripcion, ADDREGISTROS, USAFORMULA, BORRAR, INSERTAR, NOTILDE "
          SQLX$ = SQLX$ + ", CANTICAMPOS, CodiEmpr, USARECORD, FECHEMISI, TITURECORD, ConsultaSql)"
          SQLX$ = SQLX$ + "VALUES ( " & IDLISTAX&
          SQLX$ = SQLX$ + ",'SECTPLANTA'"
          SQLX$ = SQLX$ + ",'SECTORES EN PLANTA'"
          SQLX$ = SQLX$ + ",1"
          SQLX$ = SQLX$ + ",1"
          SQLX$ = SQLX$ + " ,1"
          SQLX$ = SQLX$ + " ,1"
          SQLX$ = SQLX$ + " ,0"
          SQLX$ = SQLX$ + " ,5"
          SQLX$ = SQLX$ + " ,0"
          SQLX$ = SQLX$ + " ,0"
          SQLX$ = SQLX$ + " ,'20/04/20'"
          SQLX$ = SQLX$ + " ,''"
          SQLX$ = SQLX$ + " ,'SELECT CAMPO1,CAMPO2,CAMPO3,CAMPO4 FROM DATASQL WITH(NOLOCK)'"
          SQLX$ = SQLX$ + ")"
          FLEXCONN.Execute (SQLX$)
            SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
            SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
            SQLX$ = SQLX$ + " ,0"
            SQLX$ = SQLX$ + " ,0"
            SQLX$ = SQLX$ + " , 'CODIGO/F8'"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " , 'S'"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " )"
            FLEXCONN.Execute (SQLX$)
        
        '
            SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
            SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
            SQLX$ = SQLX$ + " ,0"
            SQLX$ = SQLX$ + " ,0"
            SQLX$ = SQLX$ + " , 'SECTOR/A24'"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " , 'S'"
            SQLX$ = SQLX$ + " , ' '"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " )"
            FLEXCONN.Execute (SQLX$)
        
        '
            SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
            SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
            SQLX$ = SQLX$ + " ,0"
            SQLX$ = SQLX$ + " ,0"
            SQLX$ = SQLX$ + " , 'NADA/A0'"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " , ' '"
            SQLX$ = SQLX$ + " , ' '"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " )"
            FLEXCONN.Execute (SQLX$)
        
        '
            SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
            SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
            SQLX$ = SQLX$ + " ,0"
            SQLX$ = SQLX$ + " ,0"
            SQLX$ = SQLX$ + " , 'NADA/A0'"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " )"
            FLEXCONN.Execute (SQLX$)
        
        '
        End If
        '|FIN SECTPLANTA
        
        '|INICIO ALTOUTIL
        '
        CONDI$ = " LISTA =   'ALTOUTIL'"
        EXILISTA% = CantRegistros&("ENCAPRFUN", CONDI$, "NOMESS")
        If EXILISTA% > 0 Then
          FE% = 16778
          FEBD% = CVINT(VALOBADA("ENCAPRFUN", "FECHEMISI", CONDI$, "NOMESS"))
          GENERAREST% = 0
          If FEBD% < FE% Then
            IDLISTAX& = XVALO(VALOBADA("ENCAPRFUN", "ID_LISTA", CONDI$, "NOMESS"))
            Call BORRAREGISTROS("ENCAPRFUN", "ID_LISTA = " & IDLISTAX&, REGAF&, "NOMESS")
            Call BORRAREGISTROS("ENCAPRF", "ID_LISTA = " & IDLISTAX&, REGAF&, "NOMESS")
            GENERAREST% = 1
          End If
        Else
          GENERAREST% = 1
          CONDI$ = " ID_LISTA  > 0 "
          IDLISTAX& = 1 + XVALO(VALOBADA("ENCAPRFUN", "MAX(ID_LISTA)", CONDI$, "NOMESS"))
        End If
        If GENERAREST% > 0 Then
          SQLX$ = "INSERT INTO  ENCAPRFUN (ID_LISTA ,LISTA, Descripcion, ADDREGISTROS, USAFORMULA, BORRAR, INSERTAR, NOTILDE "
          SQLX$ = SQLX$ + ", CANTICAMPOS, CodiEmpr, USARECORD, FECHEMISI, TITURECORD, ConsultaSql)"
          SQLX$ = SQLX$ + "VALUES ( " & IDLISTAX&
          SQLX$ = SQLX$ + ",'ALTOUTIL'"
          SQLX$ = SQLX$ + ",'ALTO UTIL EN CM.'"
          SQLX$ = SQLX$ + ",1"
          SQLX$ = SQLX$ + ",1"
          SQLX$ = SQLX$ + " ,1"
          SQLX$ = SQLX$ + " ,1"
          SQLX$ = SQLX$ + " ,0"
          SQLX$ = SQLX$ + " ,4"
          SQLX$ = SQLX$ + " ,0"
          SQLX$ = SQLX$ + " ,0"
          SQLX$ = SQLX$ + " ,'10/04/20'"
          SQLX$ = SQLX$ + " ,''"
          SQLX$ = SQLX$ + " ,'SELECT CAMPO1, CAMPO2,CAMPO3 FROM DATASQL WITH(NOLOCK)'"
          SQLX$ = SQLX$ + ")"
          FLEXCONN.Execute (SQLX$)
            SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
            SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
            SQLX$ = SQLX$ + " ,0"
            SQLX$ = SQLX$ + " ,0"
            SQLX$ = SQLX$ + " , 'CODIGO/F8'"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " , 'S'"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " )"
            FLEXCONN.Execute (SQLX$)
        
        '
            SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
            SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
            SQLX$ = SQLX$ + " ,0"
            SQLX$ = SQLX$ + " ,0"
            SQLX$ = SQLX$ + " , 'ALTO UTIL CM/F12.2'"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " , 'S'"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " )"
            FLEXCONN.Execute (SQLX$)
        
        '
            SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
            SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
            SQLX$ = SQLX$ + " ,0"
            SQLX$ = SQLX$ + " ,0"
            SQLX$ = SQLX$ + " , 'NADA/A0'"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " )"
            FLEXCONN.Execute (SQLX$)
        
        '
        End If
        '|FIN ALTOUTIL

        '|INICIO RODICOLO
        '
        CONDI$ = " LISTA =   'RODICOLO'"
        EXILISTA% = CantRegistros&("ENCAPRFUN", CONDI$, "NOMESS")
        If EXILISTA% > 0 Then
          FE% = 16779
          FEBD% = CVINT(VALOBADA("ENCAPRFUN", "FECHEMISI", CONDI$, "NOMESS"))
          GENERAREST% = 0
          If FEBD% < FE% Then
            IDLISTAX& = XVALO(VALOBADA("ENCAPRFUN", "ID_LISTA", CONDI$, "NOMESS"))
            Call BORRAREGISTROS("ENCAPRFUN", "ID_LISTA = " & IDLISTAX&, REGAF&, "NOMESS")
            Call BORRAREGISTROS("ENCAPRF", "ID_LISTA = " & IDLISTAX&, REGAF&, "NOMESS")
            GENERAREST% = 1
          End If
        Else
          GENERAREST% = 1
          CONDI$ = " ID_LISTA  > 0 "
          IDLISTAX& = 1 + XVALO(VALOBADA("ENCAPRFUN", "MAX(ID_LISTA)", CONDI$, "NOMESS"))
        End If
        If GENERAREST% > 0 Then
          SQLX$ = "INSERT INTO  ENCAPRFUN (ID_LISTA ,LISTA, Descripcion, ADDREGISTROS, USAFORMULA, BORRAR, INSERTAR, NOTILDE "
          SQLX$ = SQLX$ + ", CANTICAMPOS, CodiEmpr, USARECORD, FECHEMISI, TITURECORD, ConsultaSql)"
          SQLX$ = SQLX$ + "VALUES ( " & IDLISTAX&
          SQLX$ = SQLX$ + ",'RODICOLO'"
          SQLX$ = SQLX$ + ",'TABLA RODILLOS COLORES'"
          SQLX$ = SQLX$ + ",1"
          SQLX$ = SQLX$ + ",1"
          SQLX$ = SQLX$ + " ,1"
          SQLX$ = SQLX$ + " ,1"
          SQLX$ = SQLX$ + " ,0"
          SQLX$ = SQLX$ + " ,6"
          SQLX$ = SQLX$ + " ,0"
          SQLX$ = SQLX$ + " ,0"
          SQLX$ = SQLX$ + " ,'11/04/20'"
          SQLX$ = SQLX$ + " ,''"
          SQLX$ = SQLX$ + " ,'Select campo1, campo2,campo3,campo4,campo5 from datasql with(nolock)'"
          SQLX$ = SQLX$ + ")"
          FLEXCONN.Execute (SQLX$)
            SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
            SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
            SQLX$ = SQLX$ + " ,0"
            SQLX$ = SQLX$ + " ,0"
            SQLX$ = SQLX$ + " , 'Codigo/F8'"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " , 's'"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " )"
            FLEXCONN.Execute (SQLX$)
        
        '
            SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
            SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
            SQLX$ = SQLX$ + " ,0"
            SQLX$ = SQLX$ + " ,0"
            SQLX$ = SQLX$ + " , 'Rod-LPAP/F12.2'"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " , 's'"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " )"
            FLEXCONN.Execute (SQLX$)
        
        '
            SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
            SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
            SQLX$ = SQLX$ + " ,0"
            SQLX$ = SQLX$ + " ,0"
            SQLX$ = SQLX$ + " , 'EXTRA/F12.2'"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " , 's'"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " )"
            FLEXCONN.Execute (SQLX$)
        
        '
            SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
            SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
            SQLX$ = SQLX$ + " ,0"
            SQLX$ = SQLX$ + " ,0"
            SQLX$ = SQLX$ + " , 'CantColor/F9'"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " , 'S'"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " )"
            FLEXCONN.Execute (SQLX$)
        
        '
            SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
            SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
            SQLX$ = SQLX$ + " ,0"
            SQLX$ = SQLX$ + " ,0"
            SQLX$ = SQLX$ + " , 'Nada2/A0'"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " )"
            FLEXCONN.Execute (SQLX$)
        
        '
        End If
        '|FIN RODICOLO


'|INICIO ANCHOTOT
        '
        CONDI$ = " LISTA =   'ANCHOTOT'"
        EXILISTA% = CantRegistros&("ENCAPRFUN", CONDI$, "NOMESS")
        If EXILISTA% > 0 Then
          FE% = 16776
          FEBD% = CVINT(VALOBADA("ENCAPRFUN", "FECHEMISI", CONDI$, "NOMESS"))
          GENERAREST% = 0
          If FEBD% < FE% Then
            IDLISTAX& = XVALO(VALOBADA("ENCAPRFUN", "ID_LISTA", CONDI$, "NOMESS"))
            Call BORRAREGISTROS("ENCAPRFUN", "ID_LISTA = " & IDLISTAX&, REGAF&, "NOMESS")
            Call BORRAREGISTROS("ENCAPRF", "ID_LISTA = " & IDLISTAX&, REGAF&, "NOMESS")
            GENERAREST% = 1
          End If
        Else
          GENERAREST% = 1
          CONDI$ = " ID_LISTA  > 0 "
          IDLISTAX& = 1 + XVALO(VALOBADA("ENCAPRFUN", "MAX(ID_LISTA)", CONDI$, "NOMESS"))
        End If
        If GENERAREST% > 0 Then
          SQLX$ = "INSERT INTO  ENCAPRFUN (ID_LISTA ,LISTA, Descripcion, ADDREGISTROS, USAFORMULA, BORRAR, INSERTAR, NOTILDE "
          SQLX$ = SQLX$ + ", CANTICAMPOS, CodiEmpr, USARECORD, FECHEMISI, TITURECORD, ConsultaSql)"
          SQLX$ = SQLX$ + "VALUES ( " & IDLISTAX&
          SQLX$ = SQLX$ + ",'ANCHOTOT'"
          SQLX$ = SQLX$ + ",'ANCHO TOTAL CM'"
          SQLX$ = SQLX$ + ",1"
          SQLX$ = SQLX$ + ",1"
          SQLX$ = SQLX$ + " ,1"
          SQLX$ = SQLX$ + " ,1"
          SQLX$ = SQLX$ + " ,0"
          SQLX$ = SQLX$ + " ,3"
          SQLX$ = SQLX$ + " ,0"
          SQLX$ = SQLX$ + " ,0"
          SQLX$ = SQLX$ + " ,'08/04/20'"
          SQLX$ = SQLX$ + " ,''"
          SQLX$ = SQLX$ + " ,'SELECT CAMPO1, CAMPO2,CAMPO3,CAMPO4 FROM DATASQL WITH(NOLOCK)'"
          SQLX$ = SQLX$ + ")"
          FLEXCONN.Execute (SQLX$)
            SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
            SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
            SQLX$ = SQLX$ + " ,0"
            SQLX$ = SQLX$ + " ,0"
            SQLX$ = SQLX$ + " , 'Código/f8'"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " , 's'"
            SQLX$ = SQLX$ + " , ' '"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " )"
            FLEXCONN.Execute (SQLX$)
            '
            SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
            SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
            SQLX$ = SQLX$ + " ,0"
            SQLX$ = SQLX$ + " ,0"
            SQLX$ = SQLX$ + " , 'Medida Cm./F8.2'"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " , 's'"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " , ' '"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " )"
            FLEXCONN.Execute (SQLX$)
            
            SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
            SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
            SQLX$ = SQLX$ + " ,0"
            SQLX$ = SQLX$ + " ,0"
            SQLX$ = SQLX$ + " , 'APAP/F8.2'"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " , 'S'"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " )"
            FLEXCONN.Execute (SQLX$)
            
            SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
            SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
            SQLX$ = SQLX$ + " ,0"
            SQLX$ = SQLX$ + " ,0"
            SQLX$ = SQLX$ + " , 'NV/A0'"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " )"
            FLEXCONN.Execute (SQLX$)
            '
        End If
        '|FIN ANCHOTOT



        '|INICIO MATEBOLS
        '
        CONDI$ = " LISTA =   'MATEBOLS'"
        EXILISTA% = CantRegistros&("ENCAPRFUN", CONDI$, "NOMESS")
        If EXILISTA% > 0 Then
          FE% = 16777
          FEBD% = CVINT(VALOBADA("ENCAPRFUN", "FECHEMISI", CONDI$, "NOMESS"))
          GENERAREST% = 0
          If FEBD% < FE% Then
            IDLISTAX& = XVALO(VALOBADA("ENCAPRFUN", "ID_LISTA", CONDI$, "NOMESS"))
            Call BORRAREGISTROS("ENCAPRFUN", "ID_LISTA = " & IDLISTAX&, REGAF&, "NOMESS")
            Call BORRAREGISTROS("ENCAPRF", "ID_LISTA = " & IDLISTAX&, REGAF&, "NOMESS")
            GENERAREST% = 1
          End If
        Else
          GENERAREST% = 1
          CONDI$ = " ID_LISTA  > 0 "
          IDLISTAX& = 1 + XVALO(VALOBADA("ENCAPRFUN", "MAX(ID_LISTA)", CONDI$, "NOMESS"))
        End If
        If GENERAREST% > 0 Then
          SQLX$ = "INSERT INTO  ENCAPRFUN (ID_LISTA ,LISTA, Descripcion, ADDREGISTROS, USAFORMULA, BORRAR, INSERTAR, NOTILDE "
          SQLX$ = SQLX$ + ", CANTICAMPOS, CodiEmpr, USARECORD, FECHEMISI, TITURECORD, ConsultaSql)"
          SQLX$ = SQLX$ + "VALUES ( " & IDLISTAX&
          SQLX$ = SQLX$ + ",'MATEBOLS'"
          SQLX$ = SQLX$ + ",'MATERIAL (BOLSA)'"
          SQLX$ = SQLX$ + ",1"
          SQLX$ = SQLX$ + ",1"
          SQLX$ = SQLX$ + " ,1"
          SQLX$ = SQLX$ + " ,1"
          SQLX$ = SQLX$ + " ,0"
          SQLX$ = SQLX$ + " ,3"
          SQLX$ = SQLX$ + " ,0"
          SQLX$ = SQLX$ + " ,0"
          SQLX$ = SQLX$ + " ,'09/04/20'"
          SQLX$ = SQLX$ + " ,''"
          SQLX$ = SQLX$ + " ,'SELECT CAMPO1, CAMPO2,CAMPO3,CAMPO4,CAMPO5,CAMPO6 FROM DATASQL WITH(NOLOCK)'"
          SQLX$ = SQLX$ + ")"
          FLEXCONN.Execute (SQLX$)
            SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
            SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
            SQLX$ = SQLX$ + " ,0"
            SQLX$ = SQLX$ + " ,0"
            SQLX$ = SQLX$ + " , 'Código/F8'"
            SQLX$ = SQLX$ + " , ' '"
            SQLX$ = SQLX$ + " , 's'"
            SQLX$ = SQLX$ + " , ' '"
            SQLX$ = SQLX$ + " , ' '"
            SQLX$ = SQLX$ + " , ' '"
            SQLX$ = SQLX$ + " , ' '"
            SQLX$ = SQLX$ + " )"
            FLEXCONN.Execute (SQLX$)
        
        '
            SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
            SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
            SQLX$ = SQLX$ + " ,0"
            SQLX$ = SQLX$ + " ,0"
            SQLX$ = SQLX$ + " , 'Descripción/A64'"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " , 's'"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " )"
            FLEXCONN.Execute (SQLX$)
            
            SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
            SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
            SQLX$ = SQLX$ + " ,0"
            SQLX$ = SQLX$ + " ,0"
            SQLX$ = SQLX$ + " , 'Gramaje gr-m2/F12.2'"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " , 'S'"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " )"
            FLEXCONN.Execute (SQLX$)
        
            SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
            SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
            SQLX$ = SQLX$ + " ,0"
            SQLX$ = SQLX$ + " ,0"
            SQLX$ = SQLX$ + " , 'Papel/F10.2'"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " , 'S'"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " )"
            FLEXCONN.Execute (SQLX$)
            
            SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
            SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
            SQLX$ = SQLX$ + " ,0"
            SQLX$ = SQLX$ + " ,0"
            SQLX$ = SQLX$ + " , 'Largo Cm./F10.2'"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " , 'S'"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " )"
            FLEXCONN.Execute (SQLX$)
            
            SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
            SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
            SQLX$ = SQLX$ + " ,0"
            SQLX$ = SQLX$ + " ,0"
            SQLX$ = SQLX$ + " , 'NV/A0'"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " )"
            FLEXCONN.Execute (SQLX$)
        
        '
        End If
        '|FIN MATEBOLS
        
        '|INICIO DATOSLIB
        '
        CONDI$ = " LISTA =   'DATOSLIB'"
        EXILISTA% = CantRegistros&("ENCAPRFUN", CONDI$, "NOMESS")
        If EXILISTA% > 0 Then
          FE% = 17183
          FEBD% = CVINT(VALOBADA("ENCAPRFUN", "FECHEMISI", CONDI$, "NOMESS"))
          GENERAREST% = 0
          If FEBD% < FE% Then
            IDLISTAX& = XVALO(VALOBADA("ENCAPRFUN", "ID_LISTA", CONDI$, "NOMESS"))
            Call BORRAREGISTROS("ENCAPRFUN", "ID_LISTA = " & IDLISTAX&, REGAF&, "NOMESS")
            Call BORRAREGISTROS("ENCAPRF", "ID_LISTA = " & IDLISTAX&, REGAF&, "NOMESS")
            GENERAREST% = 1
          End If
        Else
          GENERAREST% = 1
          CONDI$ = " ID_LISTA  > 0 "
          IDLISTAX& = 1 + XVALO(VALOBADA("ENCAPRFUN", "MAX(ID_LISTA)", CONDI$, "NOMESS"))
        End If
        If GENERAREST% > 0 Then
          SQLX$ = "INSERT INTO  ENCAPRFUN (ID_LISTA ,LISTA, Descripcion, ADDREGISTROS, USAFORMULA, BORRAR, INSERTAR, NOTILDE "
          SQLX$ = SQLX$ + ", CANTICAMPOS, CodiEmpr, USARECORD, FECHEMISI, TITURECORD, ConsultaSql)"
          SQLX$ = SQLX$ + "VALUES ( " & IDLISTAX&
          SQLX$ = SQLX$ + ",'DATOSLIB'"
          SQLX$ = SQLX$ + ",'DATOS ANEXOS DE FICHA TECNICA'"
          SQLX$ = SQLX$ + ",1"
          SQLX$ = SQLX$ + ",1"
          SQLX$ = SQLX$ + " ,1"
          SQLX$ = SQLX$ + " ,1"
          SQLX$ = SQLX$ + " ,0"
          SQLX$ = SQLX$ + " ,3"
          SQLX$ = SQLX$ + " ,0"
          SQLX$ = SQLX$ + " ,0"
          SQLX$ = SQLX$ + " ,'31/03/21'"
          SQLX$ = SQLX$ + " ,''"
          SQLX$ = SQLX$ + " ,'SELECT CAMPO1, CAMPO2 FROM DATASQL WITH(NOLOCK)'"
          SQLX$ = SQLX$ + ")"
          FLEXCONN.Execute (SQLX$)
            SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
            SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
            SQLX$ = SQLX$ + " ,0"
            SQLX$ = SQLX$ + " ,0"
            SQLX$ = SQLX$ + " , 'Rótulo/A48'"
            SQLX$ = SQLX$ + " , ' '"
            SQLX$ = SQLX$ + " , 'S'"
            SQLX$ = SQLX$ + " , ' '"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " )"
            FLEXCONN.Execute (SQLX$)
        
        '
            SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
            SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
            SQLX$ = SQLX$ + " ,0"
            SQLX$ = SQLX$ + " ,0"
            SQLX$ = SQLX$ + " , 'Texto/A48'"
            SQLX$ = SQLX$ + " , ' '"
            SQLX$ = SQLX$ + " , 'S'"
            SQLX$ = SQLX$ + " , ' '"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " )"
            FLEXCONN.Execute (SQLX$)
        
        '
        End If
        '|FIN DATOSLIB
        '|INICIO CODBOL
        '
        FE% = HOY(HO$)
        CONDI$ = " LISTA =   'CODBOL'"
        EXILISTA% = CantRegistros&("ENCAPRFUN", CONDI$, "NOMESS")
        If EXILISTA% > 0 Then
          FE% = 17964
          FEBD% = CVINT(VALOBADA("ENCAPRFUN", "FECHEMISI", CONDI$, "NOMESS"))
          GENERAREST% = 0
          If FEBD% < FE% Then
            IDLISTAX& = XVALO(VALOBADA("ENCAPRFUN", "ID_LISTA", CONDI$, "NOMESS"))
            Call BORRAREGISTROS("ENCAPRFUN", "ID_LISTA = " & IDLISTAX&, REGAF&, "NOMESS")
            Call BORRAREGISTROS("ENCAPRF", "ID_LISTA = " & IDLISTAX&, REGAF&, "NOMESS")
            GENERAREST% = 1
          End If
        Else
          GENERAREST% = 1
          CONDI$ = " ID_LISTA  > 0 "
          IDLISTAX& = 1 + XVALO(VALOBADA("ENCAPRFUN", "MAX(ID_LISTA)", CONDI$, "NOMESS"))
        End If
        If GENERAREST% > 0 Then
          SQLX$ = "INSERT INTO  ENCAPRFUN (ID_LISTA ,LISTA, Descripcion, ADDREGISTROS, USAFORMULA, BORRAR, INSERTAR, NOTILDE "
          SQLX$ = SQLX$ + ", CANTICAMPOS, CodiEmpr, USARECORD, FECHEMISI, TITURECORD, ConsultaSql)"
          SQLX$ = SQLX$ + "VALUES ( " & IDLISTAX&
          SQLX$ = SQLX$ + ",'CODBOL'"
          SQLX$ = SQLX$ + ",'CODIGOS CONFECCION BOLSAS'"
          SQLX$ = SQLX$ + ",1"
          SQLX$ = SQLX$ + ",1"
          SQLX$ = SQLX$ + " ,1"
          SQLX$ = SQLX$ + " ,0"
          SQLX$ = SQLX$ + " ,0"
          SQLX$ = SQLX$ + " ,7"
          SQLX$ = SQLX$ + " ,0"
          SQLX$ = SQLX$ + " ,0"
          SQLX$ = SQLX$ + " ,'12/02/23'"
          SQLX$ = SQLX$ + " ,''"
          SQLX$ = SQLX$ + " ,'SELECT CAMPO1,CAMPO2 FROM DATASQL WITH(NOLOCK)'"
          SQLX$ = SQLX$ + ")"
          FLEXCONN.Execute (SQLX$)
            SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
            SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
            SQLX$ = SQLX$ + " ,0"
            SQLX$ = SQLX$ + " ,0"
            SQLX$ = SQLX$ + " , 'CODIGO/A12'"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " , 'S'"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " )"
            FLEXCONN.Execute (SQLX$)
        
        '
            SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
            SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
            SQLX$ = SQLX$ + " ,0"
            SQLX$ = SQLX$ + " ,0"
            SQLX$ = SQLX$ + " , 'DESCRIPCIÓN/A128'"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " , 'S'"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " )"
            FLEXCONN.Execute (SQLX$)
  

        '
        End If
        '|FIN CODBOL
        
        
        '|INICIO CTP_CREPE
        '
        FE% = HOY(HO$)
        CONDI$ = " LISTA =   'CTP_CREPE'"
        EXILISTA% = CantRegistros&("ENCAPRFUN", CONDI$, "NOMESS")
        If EXILISTA% > 0 Then
          FE% = 17991
          FEBD% = CVINT(VALOBADA("ENCAPRFUN", "FECHEMISI", CONDI$, "NOMESS"))
          GENERAREST% = 0
          If FEBD% < FE% Then
            IDLISTAX& = XVALO(VALOBADA("ENCAPRFUN", "ID_LISTA", CONDI$, "NOMESS"))
            Call BORRAREGISTROS("ENCAPRFUN", "ID_LISTA = " & IDLISTAX&, REGAF&, "NOMESS")
            Call BORRAREGISTROS("ENCAPRF", "ID_LISTA = " & IDLISTAX&, REGAF&, "NOMESS")
            GENERAREST% = 1
          End If
        Else
          GENERAREST% = 1
          CONDI$ = " ID_LISTA  > 0 "
          IDLISTAX& = 1 + XVALO(VALOBADA("ENCAPRFUN", "MAX(ID_LISTA)", CONDI$, "NOMESS"))
        End If
        If GENERAREST% > 0 Then
          SQLX$ = "INSERT INTO  ENCAPRFUN (ID_LISTA ,LISTA, Descripcion, ADDREGISTROS, USAFORMULA, BORRAR, INSERTAR, NOTILDE "
          SQLX$ = SQLX$ + ", CANTICAMPOS, CodiEmpr, USARECORD, FECHEMISI, TITURECORD, ConsultaSql)"
          SQLX$ = SQLX$ + "VALUES ( " & IDLISTAX&
          SQLX$ = SQLX$ + ",'CTP_CREPE'"
          SQLX$ = SQLX$ + ",'CODIGOS PAPELES CREPE'"
          SQLX$ = SQLX$ + ",1"
          SQLX$ = SQLX$ + ",1"
          SQLX$ = SQLX$ + " ,1"
          SQLX$ = SQLX$ + " ,0"
          SQLX$ = SQLX$ + " ,0"
          SQLX$ = SQLX$ + " ,7"
          SQLX$ = SQLX$ + " ,0"
          SQLX$ = SQLX$ + " ,0"
          SQLX$ = SQLX$ + " ,'07/03/23'"
          SQLX$ = SQLX$ + " ,''"
          SQLX$ = SQLX$ + " ,'SELECT CAMPO1,CAMPO2 FROM DATASQL WITH(NOLOCK)'"
          SQLX$ = SQLX$ + ")"
          FLEXCONN.Execute (SQLX$)
            SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
            SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
            SQLX$ = SQLX$ + " ,0"
            SQLX$ = SQLX$ + " ,0"
            SQLX$ = SQLX$ + " , 'CODIGO/A12'"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " , 'S'"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " )"
            FLEXCONN.Execute (SQLX$)
        
        '
            SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
            SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
            SQLX$ = SQLX$ + " ,0"
            SQLX$ = SQLX$ + " ,0"
            SQLX$ = SQLX$ + " , 'DESCRIPCIÓN/A128'"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " , 'S'"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " )"
            FLEXCONN.Execute (SQLX$)
  

        '
        End If
        '|FIN CTP_CREPE
        
    '|INICIO CTP_ABREFACIL
        '
        FE% = HOY(HO$)
        CONDI$ = " LISTA =   'CTP_ABREFACIL'"
        EXILISTA% = CantRegistros&("ENCAPRFUN", CONDI$, "NOMESS")
        If EXILISTA% > 0 Then
          FE% = 17991
          FEBD% = CVINT(VALOBADA("ENCAPRFUN", "FECHEMISI", CONDI$, "NOMESS"))
          GENERAREST% = 0
          If FEBD% < FE% Then
            IDLISTAX& = XVALO(VALOBADA("ENCAPRFUN", "ID_LISTA", CONDI$, "NOMESS"))
            Call BORRAREGISTROS("ENCAPRFUN", "ID_LISTA = " & IDLISTAX&, REGAF&, "NOMESS")
            Call BORRAREGISTROS("ENCAPRF", "ID_LISTA = " & IDLISTAX&, REGAF&, "NOMESS")
            GENERAREST% = 1
          End If
        Else
          GENERAREST% = 1
          CONDI$ = " ID_LISTA  > 0 "
          IDLISTAX& = 1 + XVALO(VALOBADA("ENCAPRFUN", "MAX(ID_LISTA)", CONDI$, "NOMESS"))
        End If
        If GENERAREST% > 0 Then
          SQLX$ = "INSERT INTO  ENCAPRFUN (ID_LISTA ,LISTA, Descripcion, ADDREGISTROS, USAFORMULA, BORRAR, INSERTAR, NOTILDE "
          SQLX$ = SQLX$ + ", CANTICAMPOS, CodiEmpr, USARECORD, FECHEMISI, TITURECORD, ConsultaSql)"
          SQLX$ = SQLX$ + "VALUES ( " & IDLISTAX&
          SQLX$ = SQLX$ + ",'CTP_ABREFACIL'"
          SQLX$ = SQLX$ + ",'CODIGOS ABREFACIL'"
          SQLX$ = SQLX$ + ",1"
          SQLX$ = SQLX$ + ",1"
          SQLX$ = SQLX$ + " ,1"
          SQLX$ = SQLX$ + " ,0"
          SQLX$ = SQLX$ + " ,0"
          SQLX$ = SQLX$ + " ,7"
          SQLX$ = SQLX$ + " ,0"
          SQLX$ = SQLX$ + " ,0"
          SQLX$ = SQLX$ + " ,'07/03/23'"
          SQLX$ = SQLX$ + " ,''"
          SQLX$ = SQLX$ + " ,'SELECT CAMPO1,CAMPO2 FROM DATASQL WITH(NOLOCK)'"
          SQLX$ = SQLX$ + ")"
          FLEXCONN.Execute (SQLX$)
            SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
            SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
            SQLX$ = SQLX$ + " ,0"
            SQLX$ = SQLX$ + " ,0"
            SQLX$ = SQLX$ + " , 'CODIGO/A12'"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " , 'S'"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " )"
            FLEXCONN.Execute (SQLX$)
        
        '
            SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
            SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
            SQLX$ = SQLX$ + " ,0"
            SQLX$ = SQLX$ + " ,0"
            SQLX$ = SQLX$ + " , 'DESCRIPCIÓN/A128'"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " , 'S'"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " )"
            FLEXCONN.Execute (SQLX$)
  

        '
        End If
        '|FIN CTP_ABREFACIL
        
        Call GeneraTablaAuxiliarCodigoDescripcion("CTP_TIPO_COLOCADO", 17991, "TIPO / COLOCADO")
        Call GeneraTablaAuxiliarCodigoDescripcion("CTP_DENSIDAD", 17991, "DENSIDAD")
        Call GeneraTablaAuxiliarCodigoDescripcion("CTP_TipoC/S_Pestaña", 17991, "Con o Sin Pestaña")
        
        End Sub
Sub GENERAR_ESTRUCTURAS1()
    EPAC$ = TRIM$(UCase$(EMPREAC$))
    
    '|INICIO CTP.PDCOTIPO
            '
            CONDI$ = " LISTA =   'CTP.PDCOTIPO'"
            EXILISTA% = CantRegistros&("ENCAPRFUN", CONDI$, "NOMESS")
            If EXILISTA% > 0 Then
              FE% = 16588
              FEBD% = CVINT(VALOBADA("ENCAPRFUN", "FECHEMISI", CONDI$, "NOMESS"))
              GENERAREST% = 0
              If FEBD% < FE% Then
                IDLISTAX& = XVALO(VALOBADA("ENCAPRFUN", "ID_LISTA", CONDI$, "NOMESS"))
                Call BORRAREGISTROS("ENCAPRFUN", "ID_LISTA = " & IDLISTAX&, REGAF&, "NOMESS")
                Call BORRAREGISTROS("ENCAPRF", "ID_LISTA = " & IDLISTAX&, REGAF&, "NOMESS")
                GENERAREST% = 1
              End If
            Else
              GENERAREST% = 1
              CONDI$ = " ID_LISTA  > 0 "
              IDLISTAX& = 1 + XVALO(VALOBADA("ENCAPRFUN", "MAX(ID_LISTA)", CONDI$, "NOMESS"))
            End If
            If GENERAREST% > 0 Then
              SQLX$ = "INSERT INTO  ENCAPRFUN (ID_LISTA ,LISTA, Descripcion, ADDREGISTROS, USAFORMULA, BORRAR, INSERTAR, NOTILDE "
              SQLX$ = SQLX$ + ", CANTICAMPOS, CodiEmpr, USARECORD, FECHEMISI, TITURECORD, ConsultaSql)"
              SQLX$ = SQLX$ + "VALUES ( " & IDLISTAX&
              SQLX$ = SQLX$ + ",'CTP.PDCOTIPO'"
              SQLX$ = SQLX$ + ",'PORTA DOCUMENTO TIPO'"
              SQLX$ = SQLX$ + ",1"
              SQLX$ = SQLX$ + ",1"
              SQLX$ = SQLX$ + " ,1"
              SQLX$ = SQLX$ + " ,1"
              SQLX$ = SQLX$ + " ,0"
              SQLX$ = SQLX$ + " ,3"
              SQLX$ = SQLX$ + " ,0"
              SQLX$ = SQLX$ + " ,0"
              SQLX$ = SQLX$ + " ,'12/11/19'"
              SQLX$ = SQLX$ + " ,''"
              SQLX$ = SQLX$ + " ,'SELECT CAMPO1,CAMPO2 FROM DATASQL WITH(NOLOCK)'"
            '
              SQLX$ = SQLX$ + ")"
              FLEXCONN.Execute (SQLX$)
                SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
                SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
                SQLX$ = SQLX$ + " ,0"
                SQLX$ = SQLX$ + " ,0"
                SQLX$ = SQLX$ + " , 'CODIGO/F8'"
                SQLX$ = SQLX$ + " , ''"
                SQLX$ = SQLX$ + " , ' '"
                SQLX$ = SQLX$ + " , ''"
                SQLX$ = SQLX$ + " , ''"
                SQLX$ = SQLX$ + " , ''"
                SQLX$ = SQLX$ + " , ''"
                SQLX$ = SQLX$ + " )"
                FLEXCONN.Execute (SQLX$)
            
            '
                SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
                SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
                SQLX$ = SQLX$ + " ,0"
                SQLX$ = SQLX$ + " ,0"
                SQLX$ = SQLX$ + " , 'REFERENCIA/A32'"
                SQLX$ = SQLX$ + " , ''"
                SQLX$ = SQLX$ + " , 'S'"
                SQLX$ = SQLX$ + " , ''"
                SQLX$ = SQLX$ + " , ''"
                SQLX$ = SQLX$ + " , ''"
                SQLX$ = SQLX$ + " , ''"
                SQLX$ = SQLX$ + " )"
                FLEXCONN.Execute (SQLX$)
            
            '
            End If
'|FIN CTP.PDCOTIPO

    '|INICIO CTP.PCTCOT
            '
            CONDI$ = " LISTA =   'CTP.PCTCOT'"
            EXILISTA% = CantRegistros&("ENCAPRFUN", CONDI$, "NOMESS")
            If EXILISTA% > 0 Then
              FE% = 16588
              FEBD% = CVINT(VALOBADA("ENCAPRFUN", "FECHEMISI", CONDI$, "NOMESS"))
              GENERAREST% = 0
              If FEBD% < FE% Then
                IDLISTAX& = XVALO(VALOBADA("ENCAPRFUN", "ID_LISTA", CONDI$, "NOMESS"))
                Call BORRAREGISTROS("ENCAPRFUN", "ID_LISTA = " & IDLISTAX&, REGAF&, "NOMESS")
                Call BORRAREGISTROS("ENCAPRF", "ID_LISTA = " & IDLISTAX&, REGAF&, "NOMESS")
                GENERAREST% = 1
              End If
            Else
              GENERAREST% = 1
              CONDI$ = " ID_LISTA  > 0 "
              IDLISTAX& = 1 + XVALO(VALOBADA("ENCAPRFUN", "MAX(ID_LISTA)", CONDI$, "NOMESS"))
            End If
            If GENERAREST% > 0 Then
              SQLX$ = "INSERT INTO  ENCAPRFUN (ID_LISTA ,LISTA, Descripcion, ADDREGISTROS, USAFORMULA, BORRAR, INSERTAR, NOTILDE "
              SQLX$ = SQLX$ + ", CANTICAMPOS, CodiEmpr, USARECORD, FECHEMISI, TITURECORD, ConsultaSql)"
              SQLX$ = SQLX$ + "VALUES ( " & IDLISTAX&
              SQLX$ = SQLX$ + ",'CTP.PCTCOT'"
              SQLX$ = SQLX$ + ",'PORCENTAJES DE COTIZACION'"
              SQLX$ = SQLX$ + ",1"
              SQLX$ = SQLX$ + ",1"
              SQLX$ = SQLX$ + " ,1"
              SQLX$ = SQLX$ + " ,1"
              SQLX$ = SQLX$ + " ,0"
              SQLX$ = SQLX$ + " ,4"
              SQLX$ = SQLX$ + " ,0"
              SQLX$ = SQLX$ + " ,0"
              SQLX$ = SQLX$ + " ,'12/11/19'"
              SQLX$ = SQLX$ + " ,''"
              SQLX$ = SQLX$ + " ,'SELECT CAMPO1,CAMPO2,CAMPO3 FROM DATASQL WITH(NOLOCK)'"
              SQLX$ = SQLX$ + ")"
              FLEXCONN.Execute (SQLX$)
                SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
                SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
                SQLX$ = SQLX$ + " ,0"
                SQLX$ = SQLX$ + " ,0"
                SQLX$ = SQLX$ + " , 'CODIGO/F8'"
                SQLX$ = SQLX$ + " , ''"
                SQLX$ = SQLX$ + " , ' '"
                SQLX$ = SQLX$ + " , ''"
                SQLX$ = SQLX$ + " , ''"
                SQLX$ = SQLX$ + " , ''"
                SQLX$ = SQLX$ + " , ''"
                SQLX$ = SQLX$ + " )"
                FLEXCONN.Execute (SQLX$)
            
            '
                SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
                SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
                SQLX$ = SQLX$ + " ,0"
                SQLX$ = SQLX$ + " ,0"
                SQLX$ = SQLX$ + " , 'REFERENCIA/A32'"
                SQLX$ = SQLX$ + " , ''"
                SQLX$ = SQLX$ + " , 'S'"
                SQLX$ = SQLX$ + " , ' '"
                SQLX$ = SQLX$ + " , ''"
                SQLX$ = SQLX$ + " , ''"
                SQLX$ = SQLX$ + " , ''"
                SQLX$ = SQLX$ + " )"
                FLEXCONN.Execute (SQLX$)
            
            '
                SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
                SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
                SQLX$ = SQLX$ + " ,0"
                SQLX$ = SQLX$ + " ,0"
                SQLX$ = SQLX$ + " , '%/F5.2'"
                SQLX$ = SQLX$ + " , ''"
                SQLX$ = SQLX$ + " , 'S'"
                SQLX$ = SQLX$ + " , ''"
                SQLX$ = SQLX$ + " , ''"
                SQLX$ = SQLX$ + " , ''"
                SQLX$ = SQLX$ + " , ''"
                SQLX$ = SQLX$ + " )"
                FLEXCONN.Execute (SQLX$)
            
            '
            End If
 '|FIN CTP.PCTCOT

 '|INICIO Ctp.MP
            '
            CONDI$ = " LISTA =   'Ctp.MP'"
            EXILISTA% = CantRegistros&("ENCAPRFUN", CONDI$, "NOMESS")
            If EXILISTA% > 0 Then
              FE% = 16546
              FEBD% = CVINT(VALOBADA("ENCAPRFUN", "FECHEMISI", CONDI$, "NOMESS"))
              GENERAREST% = 0
              If FEBD% < FE% Then
                IDLISTAX& = XVALO(VALOBADA("ENCAPRFUN", "ID_LISTA", CONDI$, "NOMESS"))
                Call BORRAREGISTROS("ENCAPRFUN", "ID_LISTA = " & IDLISTAX&, REGAF&, "NOMESS")
                Call BORRAREGISTROS("ENCAPRF", "ID_LISTA = " & IDLISTAX&, REGAF&, "NOMESS")
                GENERAREST% = 1
              End If
            Else
              GENERAREST% = 1
              CONDI$ = " ID_LISTA  > 0 "
              IDLISTAX& = 1 + XVALO(VALOBADA("ENCAPRFUN", "MAX(ID_LISTA)", CONDI$, "NOMESS"))
            End If
            If GENERAREST% > 0 Then
              SQLX$ = "INSERT INTO  ENCAPRFUN (ID_LISTA ,LISTA, Descripcion, ADDREGISTROS, USAFORMULA, BORRAR, INSERTAR, NOTILDE "
              SQLX$ = SQLX$ + ", CANTICAMPOS, CodiEmpr, USARECORD, FECHEMISI, TITURECORD, ConsultaSql)"
              SQLX$ = SQLX$ + "VALUES ( " & IDLISTAX&
              SQLX$ = SQLX$ + ",'Ctp.MP'"
              SQLX$ = SQLX$ + ",'TABLA CONFIG TELA MAMPARO MP '"
              SQLX$ = SQLX$ + ",1"
              SQLX$ = SQLX$ + ",1"
              SQLX$ = SQLX$ + " ,1"
              SQLX$ = SQLX$ + " ,0"
              SQLX$ = SQLX$ + " ,0"
              SQLX$ = SQLX$ + " ,4"
              SQLX$ = SQLX$ + " ,0"
              SQLX$ = SQLX$ + " ,0"
              SQLX$ = SQLX$ + " ,'02/10/19'"
              SQLX$ = SQLX$ + " ,''"
              SQLX$ = SQLX$ + " ,'SELECT CAMPO1,CAMPO2,CAMPO3,CAMPO4 FROM DATASQL WITH(NOLOCK)'"
              SQLX$ = SQLX$ + ")"
              FLEXCONN.Execute (SQLX$)
                SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
                SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
                SQLX$ = SQLX$ + " ,0"
                SQLX$ = SQLX$ + " ,0"
                SQLX$ = SQLX$ + " , 'CODIGO/F8'"
                SQLX$ = SQLX$ + " , ''"
                SQLX$ = SQLX$ + " , ' '"
                SQLX$ = SQLX$ + " , ''"
                SQLX$ = SQLX$ + " , ''"
                SQLX$ = SQLX$ + " , ''"
                SQLX$ = SQLX$ + " , ''"
                SQLX$ = SQLX$ + " )"
                FLEXCONN.Execute (SQLX$)
            
            '
                SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
                SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
                SQLX$ = SQLX$ + " ,0"
                SQLX$ = SQLX$ + " ,0"
                SQLX$ = SQLX$ + " , 'COLUMNA1/F8'"
                SQLX$ = SQLX$ + " , ''"
                SQLX$ = SQLX$ + " , 'S'"
                SQLX$ = SQLX$ + " , ''"
                SQLX$ = SQLX$ + " , ''"
                SQLX$ = SQLX$ + " , ''"
                SQLX$ = SQLX$ + " , ''"
                SQLX$ = SQLX$ + " )"
                FLEXCONN.Execute (SQLX$)
            
                SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
                SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
                SQLX$ = SQLX$ + " ,0"
                SQLX$ = SQLX$ + " ,0"
                SQLX$ = SQLX$ + " , 'COLUMNA2/F8'"
                SQLX$ = SQLX$ + " , ''"
                SQLX$ = SQLX$ + " , 'S'"
                SQLX$ = SQLX$ + " , ''"
                SQLX$ = SQLX$ + " , ''"
                SQLX$ = SQLX$ + " , ''"
                SQLX$ = SQLX$ + " , ''"
                SQLX$ = SQLX$ + " )"
                FLEXCONN.Execute (SQLX$)
                
                SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
                SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
                SQLX$ = SQLX$ + " ,0"
                SQLX$ = SQLX$ + " ,0"
                SQLX$ = SQLX$ + " , 'NV/A0'"
                SQLX$ = SQLX$ + " , ''"
                SQLX$ = SQLX$ + " , ''"
                SQLX$ = SQLX$ + " , ''"
                SQLX$ = SQLX$ + " , ''"
                SQLX$ = SQLX$ + " , ''"
                SQLX$ = SQLX$ + " , ''"
                SQLX$ = SQLX$ + " )"
                FLEXCONN.Execute (SQLX$)
           
            '
            End If
            '|FIN Ctp.MP
            'SE REEMPLAZO POR LA TABLA CALCUCORTE
    If InStr(1, EPAC$, "CONTEPLAST") > 0 Then
'            '|INICIO MEDICORTE
'            '
'            CONDI$ = " LISTA =   'MEDICORTE'"
'            EXILISTA% = CantRegistros&("ENCAPRFUN", CONDI$, "NOMESS")
'            If EXILISTA% > 0 Then
'              FE% = 16599
'              FEBD% = CVINT(VALOBADA("ENCAPRFUN", "FECHEMISI", CONDI$, "NOMESS"))
'              GENERAREST% = 0
'              If FEBD% < FE% Then
'                IDLISTAX& = XVALO(VALOBADA("ENCAPRFUN", "ID_LISTA", CONDI$, "NOMESS"))
'                Call BORRAREGISTROS("ENCAPRFUN", "ID_LISTA = " & IDLISTAX&, REGAF&, "NOMESS")
'                Call BORRAREGISTROS("ENCAPRF", "ID_LISTA = " & IDLISTAX&, REGAF&, "NOMESS")
'                GENERAREST% = 1
'              End If
'            Else
'              GENERAREST% = 1
'              CONDI$ = " ID_LISTA  > 0 "
'              IDLISTAX& = 1 + XVALO(VALOBADA("ENCAPRFUN", "MAX(ID_LISTA)", CONDI$, "NOMESS"))
'            End If
'            If GENERAREST% > 0 Then
'              SQLX$ = "INSERT INTO  ENCAPRFUN (ID_LISTA ,LISTA, Descripcion, ADDREGISTROS, USAFORMULA, BORRAR, INSERTAR, NOTILDE "
'              SQLX$ = SQLX$ + ", CANTICAMPOS, CodiEmpr, USARECORD, FECHEMISI, TITURECORD, ConsultaSql)"
'              SQLX$ = SQLX$ + "VALUES ( " & IDLISTAX&
'              SQLX$ = SQLX$ + ",'MEDICORTE'"
'              SQLX$ = SQLX$ + ",'FORMULAS CORTE (Separados x (;) EJ:ALTO;*;2;+;LARGO;*;2;+;7)'"
'
'              SQLX$ = SQLX$ + ",1"
'              SQLX$ = SQLX$ + ",1"
'              SQLX$ = SQLX$ + " ,1"
'              SQLX$ = SQLX$ + " ,1"
'              SQLX$ = SQLX$ + " ,0"
'              SQLX$ = SQLX$ + " ,9"
'              SQLX$ = SQLX$ + " ,0"
'              SQLX$ = SQLX$ + " ,1"
'              SQLX$ = SQLX$ + " ,'23/11/19'"
'              SQLX$ = SQLX$ + " ,'CORTES'"
'              SQLX$ = SQLX$ + " ,'SELECT CAMPO1,CAMPO2,CAMPO3,CAMPO4,CAMPO5, CAMPO6,CAMPO7,CAMPO8 FROM DATASQL WITH(NOLOCK)'"
'              SQLX$ = SQLX$ + ")"
'              FLEXCONN.Execute (SQLX$)
'                SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
'                SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
'                SQLX$ = SQLX$ + " ,0"
'                SQLX$ = SQLX$ + " ,0"
'                SQLX$ = SQLX$ + " , 'COD.BIGBAG/A12'"
'                SQLX$ = SQLX$ + " , ''"
'                SQLX$ = SQLX$ + " , 'S'"
'                SQLX$ = SQLX$ + " , " & "'CAMPO1;CAMPO2;DATASQL;LISTA ='" & "'CODBB'" & "''"
'                SQLX$ = SQLX$ + " , ''"
'                SQLX$ = SQLX$ + " , ''"
'                SQLX$ = SQLX$ + " , ''"
'                SQLX$ = SQLX$ + " )"
'                FLEXCONN.Execute (SQLX$)
'
'            '
'                SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
'                SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
'                SQLX$ = SQLX$ + " ,0"
'                SQLX$ = SQLX$ + " ,0"
'                SQLX$ = SQLX$ + " , 'DESCRIPCIÓN/A64'"
'                SQLX$ = SQLX$ + " , ''"
'                SQLX$ = SQLX$ + " , ''"
'                SQLX$ = SQLX$ + " , ''"
'                SQLX$ = SQLX$ + " , ''"
'                SQLX$ = SQLX$ + " , ''"
'                SQLX$ = SQLX$ + " , ''"
'                SQLX$ = SQLX$ + " )"
'                FLEXCONN.Execute (SQLX$)
'
'            '
'                SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
'                SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
'                SQLX$ = SQLX$ + " ,0"
'                SQLX$ = SQLX$ + " ,0"
'                SQLX$ = SQLX$ + " , 'LATERALES/A24'"
'                SQLX$ = SQLX$ + " , ' '"
'                SQLX$ = SQLX$ + " , 'S'"
'                SQLX$ = SQLX$ + " , ''"
'                SQLX$ = SQLX$ + " , ''"
'                SQLX$ = SQLX$ + " , ''"
'                SQLX$ = SQLX$ + " , ''"
'                SQLX$ = SQLX$ + " )"
'                FLEXCONN.Execute (SQLX$)
'
'            '
'                SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
'                SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
'                SQLX$ = SQLX$ + " ,0"
'                SQLX$ = SQLX$ + " ,0"
'                SQLX$ = SQLX$ + " , 'CUERPO/A24'"
'                SQLX$ = SQLX$ + " , ''"
'                SQLX$ = SQLX$ + " , 'S'"
'                SQLX$ = SQLX$ + " , ''"
'                SQLX$ = SQLX$ + " , ''"
'                SQLX$ = SQLX$ + " , ''"
'                SQLX$ = SQLX$ + " , ''"
'                SQLX$ = SQLX$ + " )"
'                FLEXCONN.Execute (SQLX$)
'
'            '
'                SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
'                SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
'                SQLX$ = SQLX$ + " ,0"
'                SQLX$ = SQLX$ + " ,0"
'                SQLX$ = SQLX$ + " , 'FONDO/A24'"
'                SQLX$ = SQLX$ + " , ''"
'                SQLX$ = SQLX$ + " , 'S'"
'                SQLX$ = SQLX$ + " , ''"
'                SQLX$ = SQLX$ + " , ''"
'                SQLX$ = SQLX$ + " , ''"
'                SQLX$ = SQLX$ + " , ''"
'                SQLX$ = SQLX$ + " )"
'                FLEXCONN.Execute (SQLX$)
'
'            '
'                SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
'                SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
'                SQLX$ = SQLX$ + " ,0"
'                SQLX$ = SQLX$ + " ,0"
'                SQLX$ = SQLX$ + " , 'U PANEL/A24'"
'                SQLX$ = SQLX$ + " , ''"
'                SQLX$ = SQLX$ + " , 'S'"
'                SQLX$ = SQLX$ + " , ''"
'                SQLX$ = SQLX$ + " , ''"
'                SQLX$ = SQLX$ + " , ''"
'                SQLX$ = SQLX$ + " , ''"
'                SQLX$ = SQLX$ + " )"
'                FLEXCONN.Execute (SQLX$)
'
'            '
'                SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
'                SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
'                SQLX$ = SQLX$ + " ,0"
'                SQLX$ = SQLX$ + " ,0"
'                SQLX$ = SQLX$ + " , 'TAPA/A24'"
'                SQLX$ = SQLX$ + " , ''"
'                SQLX$ = SQLX$ + " , 'S'"
'                SQLX$ = SQLX$ + " , ''"
'                SQLX$ = SQLX$ + " , ''"
'                SQLX$ = SQLX$ + " , ''"
'                SQLX$ = SQLX$ + " , ''"
'                SQLX$ = SQLX$ + " )"
'                FLEXCONN.Execute (SQLX$)
'
'            '
'                SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
'                SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
'                SQLX$ = SQLX$ + " ,0"
'                SQLX$ = SQLX$ + " ,0"
'                SQLX$ = SQLX$ + " , 'POLLERA/A24'"
'                SQLX$ = SQLX$ + " , ''"
'                SQLX$ = SQLX$ + " , 'S'"
'                SQLX$ = SQLX$ + " , ''"
'                SQLX$ = SQLX$ + " , ''"
'                SQLX$ = SQLX$ + " , ''"
'                SQLX$ = SQLX$ + " , ''"
'                SQLX$ = SQLX$ + " )"
'                FLEXCONN.Execute (SQLX$)
'
'            '
'            End If
'            '|FIN MEDICORTE
    
            '|INICIO CODBB
            '
            CONDI$ = " LISTA =   'CODBB'"
            EXILISTA% = CantRegistros&("ENCAPRFUN", CONDI$, "NOMESS")
            If EXILISTA% > 0 Then
              FE% = 16599
              FEBD% = CVINT(VALOBADA("ENCAPRFUN", "FECHEMISI", CONDI$, "NOMESS"))
              GENERAREST% = 0
              If FEBD% < FE% Then
                IDLISTAX& = XVALO(VALOBADA("ENCAPRFUN", "ID_LISTA", CONDI$, "NOMESS"))
                Call BORRAREGISTROS("ENCAPRFUN", "ID_LISTA = " & IDLISTAX&, REGAF&, "NOMESS")
                Call BORRAREGISTROS("ENCAPRF", "ID_LISTA = " & IDLISTAX&, REGAF&, "NOMESS")
                GENERAREST% = 1
              End If
            Else
              GENERAREST% = 1
              CONDI$ = " ID_LISTA  > 0 "
              IDLISTAX& = 1 + XVALO(VALOBADA("ENCAPRFUN", "MAX(ID_LISTA)", CONDI$, "NOMESS"))
            End If
            If GENERAREST% > 0 Then
              SQLX$ = "INSERT INTO  ENCAPRFUN (ID_LISTA ,LISTA, Descripcion, ADDREGISTROS, USAFORMULA, BORRAR, INSERTAR, NOTILDE "
              SQLX$ = SQLX$ + ", CANTICAMPOS, CodiEmpr, USARECORD, FECHEMISI, TITURECORD, ConsultaSql)"
              SQLX$ = SQLX$ + "VALUES ( " & IDLISTAX&
              SQLX$ = SQLX$ + ",'CODBB'"
              SQLX$ = SQLX$ + ",'CODIGOS BIG BAG'"
              SQLX$ = SQLX$ + ",1"
              SQLX$ = SQLX$ + ",1"
              SQLX$ = SQLX$ + " ,1"
              SQLX$ = SQLX$ + " ,0"
              SQLX$ = SQLX$ + " ,0"
              SQLX$ = SQLX$ + " ,7"
              SQLX$ = SQLX$ + " ,0"
              SQLX$ = SQLX$ + " ,0"
              SQLX$ = SQLX$ + " ,'23/11/19'"
              SQLX$ = SQLX$ + " ,''"
              SQLX$ = SQLX$ + " ,'SELECT CAMPO1,CAMPO2,CAMPO3,CAMPO4,CAMPO5,CAMPO6 FROM DATASQL WITH(NOLOCK)'"
              SQLX$ = SQLX$ + ")"
              FLEXCONN.Execute (SQLX$)
                SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
                SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
                SQLX$ = SQLX$ + " ,0"
                SQLX$ = SQLX$ + " ,0"
                SQLX$ = SQLX$ + " , 'CODIGO/A12'"
                SQLX$ = SQLX$ + " , ''"
                SQLX$ = SQLX$ + " , 'S'"
                SQLX$ = SQLX$ + " , ''"
                SQLX$ = SQLX$ + " , ''"
                SQLX$ = SQLX$ + " , ''"
                SQLX$ = SQLX$ + " , ''"
                SQLX$ = SQLX$ + " )"
                FLEXCONN.Execute (SQLX$)
            
            '
                SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
                SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
                SQLX$ = SQLX$ + " ,0"
                SQLX$ = SQLX$ + " ,0"
                SQLX$ = SQLX$ + " , 'DESCRIPCIÓN/A128'"
                SQLX$ = SQLX$ + " , ''"
                SQLX$ = SQLX$ + " , 'S'"
                SQLX$ = SQLX$ + " , ''"
                SQLX$ = SQLX$ + " , ''"
                SQLX$ = SQLX$ + " , ''"
                SQLX$ = SQLX$ + " , ''"
                SQLX$ = SQLX$ + " )"
                FLEXCONN.Execute (SQLX$)
            
            '
                SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
                SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
                SQLX$ = SQLX$ + " ,0"
                SQLX$ = SQLX$ + " ,0"
                SQLX$ = SQLX$ + " , 'OTRODATO1/A0'"
                SQLX$ = SQLX$ + " , ''"
                SQLX$ = SQLX$ + " , ''"
                SQLX$ = SQLX$ + " , ''"
                SQLX$ = SQLX$ + " , ''"
                SQLX$ = SQLX$ + " , ''"
                SQLX$ = SQLX$ + " , ''"
                SQLX$ = SQLX$ + " )"
                FLEXCONN.Execute (SQLX$)
            
            '
                SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
                SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
                SQLX$ = SQLX$ + " ,0"
                SQLX$ = SQLX$ + " ,0"
                SQLX$ = SQLX$ + " , 'OTRODATO/A0'"
                SQLX$ = SQLX$ + " , ''"
                SQLX$ = SQLX$ + " , ''"
                SQLX$ = SQLX$ + " , ''"
                SQLX$ = SQLX$ + " , ''"
                SQLX$ = SQLX$ + " , ''"
                SQLX$ = SQLX$ + " , ''"
                SQLX$ = SQLX$ + " )"
                FLEXCONN.Execute (SQLX$)
            
            '
                SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
                SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
                SQLX$ = SQLX$ + " ,0"
                SQLX$ = SQLX$ + " ,0"
                SQLX$ = SQLX$ + " , 'OTRODATO3/A0'"
                SQLX$ = SQLX$ + " , ''"
                SQLX$ = SQLX$ + " , ''"
                SQLX$ = SQLX$ + " , ''"
                SQLX$ = SQLX$ + " , ''"
                SQLX$ = SQLX$ + " , ''"
                SQLX$ = SQLX$ + " , ''"
                SQLX$ = SQLX$ + " )"
                FLEXCONN.Execute (SQLX$)
            
            '
                SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
                SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
                SQLX$ = SQLX$ + " ,0"
                SQLX$ = SQLX$ + " ,0"
                SQLX$ = SQLX$ + " , 'NV/A0'"
                SQLX$ = SQLX$ + " , ''"
                SQLX$ = SQLX$ + " , ''"
                SQLX$ = SQLX$ + " , ''"
                SQLX$ = SQLX$ + " , ''"
                SQLX$ = SQLX$ + " , ''"
                SQLX$ = SQLX$ + " , ''"
                SQLX$ = SQLX$ + " )"
                FLEXCONN.Execute (SQLX$)
            
            '
            End If
            '|FIN CODBB

            '|INICIO Ctp.Caras
            '
            CONDI$ = " LISTA =   'Ctp.Caras'"
            EXILISTA% = CantRegistros&("ENCAPRFUN", CONDI$, "NOMESS")
            If EXILISTA% > 0 Then
              FE% = 16546
              FEBD% = CVINT(VALOBADA("ENCAPRFUN", "FECHEMISI", CONDI$, "NOMESS"))
              GENERAREST% = 0
              If FEBD% < FE% Then
                IDLISTAX& = XVALO(VALOBADA("ENCAPRFUN", "ID_LISTA", CONDI$, "NOMESS"))
                Call BORRAREGISTROS("ENCAPRFUN", "ID_LISTA = " & IDLISTAX&, REGAF&, "NOMESS")
                Call BORRAREGISTROS("ENCAPRF", "ID_LISTA = " & IDLISTAX&, REGAF&, "NOMESS")
                GENERAREST% = 1
              
              End If
            Else
              GENERAREST% = 1
              CONDI$ = " ID_LISTA  > 0 "
              IDLISTAX& = 1 + XVALO(VALOBADA("ENCAPRFUN", "MAX(ID_LISTA)", CONDI$, "NOMESS"))
            End If
            If GENERAREST% > 0 Then
              SQLX$ = "INSERT INTO  ENCAPRFUN (ID_LISTA ,LISTA, Descripcion, ADDREGISTROS, USAFORMULA, BORRAR, INSERTAR, NOTILDE "
              SQLX$ = SQLX$ + ", CANTICAMPOS, CodiEmpr, USARECORD, FECHEMISI, TITURECORD, ConsultaSql)"
              SQLX$ = SQLX$ + "VALUES ( " & IDLISTAX&
              SQLX$ = SQLX$ + ",'Ctp.Caras'"
              SQLX$ = SQLX$ + ",'CANTIDAD DE CARAS '"
              SQLX$ = SQLX$ + ",1"
              SQLX$ = SQLX$ + ",1"
              SQLX$ = SQLX$ + " ,1"
              SQLX$ = SQLX$ + " ,0"
              SQLX$ = SQLX$ + " ,0"
              SQLX$ = SQLX$ + " ,3"
              SQLX$ = SQLX$ + " ,0"
              SQLX$ = SQLX$ + " ,0"
              SQLX$ = SQLX$ + " ,'02/10/19'"
              SQLX$ = SQLX$ + " ,''"
              SQLX$ = SQLX$ + " ,'SELECT CAMPO1,CAMPO2,CAMPO3 FROM DATASQL WITH(NOLOCK)'"
              SQLX$ = SQLX$ + ")"
              FLEXCONN.Execute (SQLX$)
                SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
                SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
                SQLX$ = SQLX$ + " ,0"
                SQLX$ = SQLX$ + " ,0"
                SQLX$ = SQLX$ + " , 'CODIGO/F8'"
                SQLX$ = SQLX$ + " , ''"
                SQLX$ = SQLX$ + " , ' '"
                SQLX$ = SQLX$ + " , ''"
                SQLX$ = SQLX$ + " , ''"
                SQLX$ = SQLX$ + " , ''"
                SQLX$ = SQLX$ + " , ''"
                SQLX$ = SQLX$ + " )"
                FLEXCONN.Execute (SQLX$)
            
            '
                SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
                SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
                SQLX$ = SQLX$ + " ,0"
                SQLX$ = SQLX$ + " ,0"
                SQLX$ = SQLX$ + " , 'CANTIDAD/F8'"
                SQLX$ = SQLX$ + " , ''"
                SQLX$ = SQLX$ + " , 'S'"
                SQLX$ = SQLX$ + " , ''"
                SQLX$ = SQLX$ + " , ''"
                SQLX$ = SQLX$ + " , ''"
                SQLX$ = SQLX$ + " , ''"
                SQLX$ = SQLX$ + " )"
                FLEXCONN.Execute (SQLX$)
            
                SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
                SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
                SQLX$ = SQLX$ + " ,0"
                SQLX$ = SQLX$ + " ,0"
                SQLX$ = SQLX$ + " , 'NV/A0'"
                SQLX$ = SQLX$ + " , ''"
                SQLX$ = SQLX$ + " , ''"
                SQLX$ = SQLX$ + " , ''"
                SQLX$ = SQLX$ + " , ''"
                SQLX$ = SQLX$ + " , ''"
                SQLX$ = SQLX$ + " , ''"
                SQLX$ = SQLX$ + " )"
                FLEXCONN.Execute (SQLX$)
           
            '
            End If
            '|FIN Ctp.Caras
            
            

'|INICIO CTP.LINER
            '
            CONDI$ = " LISTA =   'CTP.LINER'"
            EXILISTA% = CantRegistros&("ENCAPRFUN", CONDI$, "NOMESS")
            If EXILISTA% > 0 Then
              FE% = 16547
              FEBD% = CVINT(VALOBADA("ENCAPRFUN", "FECHEMISI", CONDI$, "NOMESS"))
              GENERAREST% = 0
              If FEBD% < FE% Then
                IDLISTAX& = XVALO(VALOBADA("ENCAPRFUN", "ID_LISTA", CONDI$, "NOMESS"))
                Call BORRAREGISTROS("ENCAPRFUN", "ID_LISTA = " & IDLISTAX&, REGAF&, "NOMESS")
                Call BORRAREGISTROS("ENCAPRF", "ID_LISTA = " & IDLISTAX&, REGAF&, "NOMESS")
                GENERAREST% = 1
              End If
            Else
              GENERAREST% = 1
              CONDI$ = " ID_LISTA  > 0 "
              IDLISTAX& = 1 + XVALO(VALOBADA("ENCAPRFUN", "MAX(ID_LISTA)", CONDI$, "NOMESS"))
            End If
            If GENERAREST% > 0 Then
              SQLX$ = "INSERT INTO  ENCAPRFUN (ID_LISTA ,LISTA, Descripcion, ADDREGISTROS, USAFORMULA, BORRAR, INSERTAR, NOTILDE "
              SQLX$ = SQLX$ + ", CANTICAMPOS, CodiEmpr, USARECORD, FECHEMISI, TITURECORD, ConsultaSql)"
              SQLX$ = SQLX$ + "VALUES ( " & IDLISTAX&
              SQLX$ = SQLX$ + ",'CTP.LINER'"
              SQLX$ = SQLX$ + ",'LINER '"
              SQLX$ = SQLX$ + ",1"
              SQLX$ = SQLX$ + ",1"
              SQLX$ = SQLX$ + " ,0"
              SQLX$ = SQLX$ + " ,0"
              SQLX$ = SQLX$ + " ,0"
              SQLX$ = SQLX$ + " ,3"
              SQLX$ = SQLX$ + " ,0"
              SQLX$ = SQLX$ + " ,0"
              SQLX$ = SQLX$ + " ,'03/10/19'"
              SQLX$ = SQLX$ + " ,''"
              SQLX$ = SQLX$ + " ,'SELECT CAMPO1, CAMPO2,CAMPO3 FROM DATASQL WITH(NOLOCK)'"
              SQLX$ = SQLX$ + ")"
              FLEXCONN.Execute (SQLX$)
                SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
                SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
                SQLX$ = SQLX$ + " ,0"
                SQLX$ = SQLX$ + " ,0"
                SQLX$ = SQLX$ + " , 'ID/F8'"
                SQLX$ = SQLX$ + " , ''"
                SQLX$ = SQLX$ + " , ' '"
                SQLX$ = SQLX$ + " , ''"
                SQLX$ = SQLX$ + " , ''"
                SQLX$ = SQLX$ + " , ''"
                SQLX$ = SQLX$ + " , ''"
                SQLX$ = SQLX$ + " )"
                FLEXCONN.Execute (SQLX$)
            
            '
                SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
                SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
                SQLX$ = SQLX$ + " ,0"
                SQLX$ = SQLX$ + " ,0"
                SQLX$ = SQLX$ + " , 'LINER/F8'"
                SQLX$ = SQLX$ + " , ''"
                SQLX$ = SQLX$ + " , 'S'"
                SQLX$ = SQLX$ + " , ''"
                SQLX$ = SQLX$ + " , ''"
                SQLX$ = SQLX$ + " , ''"
                SQLX$ = SQLX$ + " , ''"
                SQLX$ = SQLX$ + " )"
                FLEXCONN.Execute (SQLX$)
            
                SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
                SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
                SQLX$ = SQLX$ + " ,0"
                SQLX$ = SQLX$ + " ,0"
                SQLX$ = SQLX$ + " , 'NV/A0'"
                SQLX$ = SQLX$ + " , ''"
                SQLX$ = SQLX$ + " , ''"
                SQLX$ = SQLX$ + " , ''"
                SQLX$ = SQLX$ + " , ''"
                SQLX$ = SQLX$ + " , ''"
                SQLX$ = SQLX$ + " , ''"
                SQLX$ = SQLX$ + " )"
                FLEXCONN.Execute (SQLX$)
            
            '
            End If
            '|FIN CTP.LINER
            
            '|INICIO CTP.COSTURA
            '
            CONDI$ = " LISTA =   'CTP.COSTURA'"
            EXILISTA% = CantRegistros&("ENCAPRFUN", CONDI$, "NOMESS")
            If EXILISTA% > 0 Then
              FE% = 16547
              FEBD% = CVINT(VALOBADA("ENCAPRFUN", "FECHEMISI", CONDI$, "NOMESS"))
              GENERAREST% = 0
              If FEBD% < FE% Then
                IDLISTAX& = XVALO(VALOBADA("ENCAPRFUN", "ID_LISTA", CONDI$, "NOMESS"))
                Call BORRAREGISTROS("ENCAPRFUN", "ID_LISTA = " & IDLISTAX&, REGAF&, "NOMESS")
                Call BORRAREGISTROS("ENCAPRF", "ID_LISTA = " & IDLISTAX&, REGAF&, "NOMESS")
                GENERAREST% = 1
              End If
            Else
              GENERAREST% = 1
              CONDI$ = " ID_LISTA  > 0 "
              IDLISTAX& = 1 + XVALO(VALOBADA("ENCAPRFUN", "MAX(ID_LISTA)", CONDI$, "NOMESS"))
            End If
            If GENERAREST% > 0 Then
              SQLX$ = "INSERT INTO  ENCAPRFUN (ID_LISTA ,LISTA, Descripcion, ADDREGISTROS, USAFORMULA, BORRAR, INSERTAR, NOTILDE "
              SQLX$ = SQLX$ + ", CANTICAMPOS, CodiEmpr, USARECORD, FECHEMISI, TITURECORD, ConsultaSql)"
              SQLX$ = SQLX$ + "VALUES ( " & IDLISTAX&
              SQLX$ = SQLX$ + ",'CTP.COSTURA'"
              SQLX$ = SQLX$ + ",'COSUTRA'"
              SQLX$ = SQLX$ + ",1"
              SQLX$ = SQLX$ + ",1"
              SQLX$ = SQLX$ + " ,0"
              SQLX$ = SQLX$ + " ,0"
              SQLX$ = SQLX$ + " ,0"
              SQLX$ = SQLX$ + " ,3"
              SQLX$ = SQLX$ + " ,0"
              SQLX$ = SQLX$ + " ,0"
              SQLX$ = SQLX$ + " ,'03/10/19'"
              SQLX$ = SQLX$ + " ,''"
              SQLX$ = SQLX$ + " ,'SELECT CAMPO1, CAMPO2,CAMPO3 FROM DATASQL WITH(NOLOCK)'"
              SQLX$ = SQLX$ + ")"
              FLEXCONN.Execute (SQLX$)
                SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
                SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
                SQLX$ = SQLX$ + " ,0"
                SQLX$ = SQLX$ + " ,0"
                SQLX$ = SQLX$ + " , 'ID/F8'"
                SQLX$ = SQLX$ + " , ''"
                SQLX$ = SQLX$ + " , ' '"
                SQLX$ = SQLX$ + " , ''"
                SQLX$ = SQLX$ + " , ''"
                SQLX$ = SQLX$ + " , ''"
                SQLX$ = SQLX$ + " , ''"
                SQLX$ = SQLX$ + " )"
                FLEXCONN.Execute (SQLX$)
            
            '
                SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
                SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
                SQLX$ = SQLX$ + " ,0"
                SQLX$ = SQLX$ + " ,0"
                SQLX$ = SQLX$ + " , 'COSTURA/A32'"
                SQLX$ = SQLX$ + " , ''"
                SQLX$ = SQLX$ + " , 'S'"
                SQLX$ = SQLX$ + " , ' '"
                SQLX$ = SQLX$ + " , ''"
                SQLX$ = SQLX$ + " , ''"
                SQLX$ = SQLX$ + " , ''"
                SQLX$ = SQLX$ + " )"
                FLEXCONN.Execute (SQLX$)
            
                SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
                SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
                SQLX$ = SQLX$ + " ,0"
                SQLX$ = SQLX$ + " ,0"
                SQLX$ = SQLX$ + " , 'NV/A0'"
                SQLX$ = SQLX$ + " , ''"
                SQLX$ = SQLX$ + " , ''"
                SQLX$ = SQLX$ + " , ''"
                SQLX$ = SQLX$ + " , ''"
                SQLX$ = SQLX$ + " , ''"
                SQLX$ = SQLX$ + " , ''"
                SQLX$ = SQLX$ + " )"
                FLEXCONN.Execute (SQLX$)
            
            '
            End If
            '|FIN CTP.COSTURA

            '|INICIO CTP.ES
            '
            CONDI$ = " LISTA =   'CTP.ES'"
            EXILISTA% = CantRegistros&("ENCAPRFUN", CONDI$, "NOMESS")
            If EXILISTA% > 0 Then
              FE% = 16547
              FEBD% = CVINT(VALOBADA("ENCAPRFUN", "FECHEMISI", CONDI$, "NOMESS"))
              GENERAREST% = 0
              If FEBD% < FE% Then
                IDLISTAX& = XVALO(VALOBADA("ENCAPRFUN", "ID_LISTA", CONDI$, "NOMESS"))
                Call BORRAREGISTROS("ENCAPRFUN", "ID_LISTA = " & IDLISTAX&, REGAF&, "NOMESS")
                Call BORRAREGISTROS("ENCAPRF", "ID_LISTA = " & IDLISTAX&, REGAF&, "NOMESS")
                GENERAREST% = 1
              End If
            Else
              GENERAREST% = 1
              CONDI$ = " ID_LISTA  > 0 "
              IDLISTAX& = 1 + XVALO(VALOBADA("ENCAPRFUN", "MAX(ID_LISTA)", CONDI$, "NOMESS"))
            End If
            If GENERAREST% > 0 Then
              SQLX$ = "INSERT INTO  ENCAPRFUN (ID_LISTA ,LISTA, Descripcion, ADDREGISTROS, USAFORMULA, BORRAR, INSERTAR, NOTILDE "
              SQLX$ = SQLX$ + ", CANTICAMPOS, CodiEmpr, USARECORD, FECHEMISI, TITURECORD, ConsultaSql)"
              SQLX$ = SQLX$ + "VALUES ( " & IDLISTAX&
              SQLX$ = SQLX$ + ",'CTP.ES'"
              SQLX$ = SQLX$ + ",'ES'"
              SQLX$ = SQLX$ + ",1"
              SQLX$ = SQLX$ + ",1"
              SQLX$ = SQLX$ + " ,0"
              SQLX$ = SQLX$ + " ,0"
              SQLX$ = SQLX$ + " ,0"
              SQLX$ = SQLX$ + " ,4"
              SQLX$ = SQLX$ + " ,0"
              SQLX$ = SQLX$ + " ,0"
              SQLX$ = SQLX$ + " ,'03/10/19'"
              SQLX$ = SQLX$ + " ,''"
              SQLX$ = SQLX$ + " ,'SELECT CAMPO1,CAMPO2,CAMPO3 FROM DATASQL WITH(NOLOCK)'"
              SQLX$ = SQLX$ + ")"
              FLEXCONN.Execute (SQLX$)
                SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
                SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
                SQLX$ = SQLX$ + " ,0"
                SQLX$ = SQLX$ + " ,0"
                SQLX$ = SQLX$ + " , 'ID/F8'"
                SQLX$ = SQLX$ + " , ''"
                SQLX$ = SQLX$ + " , ''"
                SQLX$ = SQLX$ + " , ''"
                SQLX$ = SQLX$ + " , ''"
                SQLX$ = SQLX$ + " , ''"
                SQLX$ = SQLX$ + " , ''"
                SQLX$ = SQLX$ + " )"
                FLEXCONN.Execute (SQLX$)
            
            '
                SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
                SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
                SQLX$ = SQLX$ + " ,0"
                SQLX$ = SQLX$ + " ,0"
                SQLX$ = SQLX$ + " , 'ES/A32'"
                SQLX$ = SQLX$ + " , ''"
                SQLX$ = SQLX$ + " , 'S'"
                SQLX$ = SQLX$ + " , ''"
                SQLX$ = SQLX$ + " , ''"
                SQLX$ = SQLX$ + " , ''"
                SQLX$ = SQLX$ + " , ''"
                SQLX$ = SQLX$ + " )"
                FLEXCONN.Execute (SQLX$)
            
            '
                SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
                SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
                SQLX$ = SQLX$ + " ,0"
                SQLX$ = SQLX$ + " ,0"
                SQLX$ = SQLX$ + " , 'NV/A0'"
                SQLX$ = SQLX$ + " , ''"
                SQLX$ = SQLX$ + " , ''"
                SQLX$ = SQLX$ + " , ''"
                SQLX$ = SQLX$ + " , ''"
                SQLX$ = SQLX$ + " , ''"
                SQLX$ = SQLX$ + " , ''"
                SQLX$ = SQLX$ + " )"
                FLEXCONN.Execute (SQLX$)
            
            '
            End If
            '|FIN CTP.ES


            '|INICIO CTP.ESCOST
            '
            CONDI$ = " LISTA =   'CTP.ESCOST'"
            EXILISTA% = CantRegistros&("ENCAPRFUN", CONDI$, "NOMESS")
            If EXILISTA% > 0 Then
              FE% = 16547
              FEBD% = CVINT(VALOBADA("ENCAPRFUN", "FECHEMISI", CONDI$, "NOMESS"))
              GENERAREST% = 0
              If FEBD% < FE% Then
                IDLISTAX& = XVALO(VALOBADA("ENCAPRFUN", "ID_LISTA", CONDI$, "NOMESS"))
                Call BORRAREGISTROS("ENCAPRFUN", "ID_LISTA = " & IDLISTAX&, REGAF&, "NOMESS")
                Call BORRAREGISTROS("ENCAPRF", "ID_LISTA = " & IDLISTAX&, REGAF&, "NOMESS")
                GENERAREST% = 1
              End If
            Else
              GENERAREST% = 1
              CONDI$ = " ID_LISTA  > 0 "
              IDLISTAX& = 1 + XVALO(VALOBADA("ENCAPRFUN", "MAX(ID_LISTA)", CONDI$, "NOMESS"))
            End If
            If GENERAREST% > 0 Then
              SQLX$ = "INSERT INTO  ENCAPRFUN (ID_LISTA ,LISTA, Descripcion, ADDREGISTROS, USAFORMULA, BORRAR, INSERTAR, NOTILDE "
              SQLX$ = SQLX$ + ", CANTICAMPOS, CodiEmpr, USARECORD, FECHEMISI, TITURECORD, ConsultaSql)"
              SQLX$ = SQLX$ + "VALUES ( " & IDLISTAX&
              SQLX$ = SQLX$ + ",'CTP.ESCOST'"
              SQLX$ = SQLX$ + ",'ESCOST'"
              SQLX$ = SQLX$ + ",1"
              SQLX$ = SQLX$ + ",1"
              SQLX$ = SQLX$ + " ,0"
              SQLX$ = SQLX$ + " ,0"
              SQLX$ = SQLX$ + " ,0"
              SQLX$ = SQLX$ + " ,4"
              SQLX$ = SQLX$ + " ,0"
              SQLX$ = SQLX$ + " ,0"
              SQLX$ = SQLX$ + " ,'03/10/19'"
              SQLX$ = SQLX$ + " ,''"
              SQLX$ = SQLX$ + " ,'SELECT CAMPO1,CAMPO2,CAMPO3 FROM DATASQL WITH(NOLOCK)'"
              SQLX$ = SQLX$ + ")"
              FLEXCONN.Execute (SQLX$)
                SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
                SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
                SQLX$ = SQLX$ + " ,0"
                SQLX$ = SQLX$ + " ,0"
                SQLX$ = SQLX$ + " , 'ID/F8'"
                SQLX$ = SQLX$ + " , ''"
                SQLX$ = SQLX$ + " , ''"
                SQLX$ = SQLX$ + " , ''"
                SQLX$ = SQLX$ + " , ''"
                SQLX$ = SQLX$ + " , ''"
                SQLX$ = SQLX$ + " , ''"
                SQLX$ = SQLX$ + " )"
                FLEXCONN.Execute (SQLX$)
            
            '
                SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
                SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
                SQLX$ = SQLX$ + " ,0"
                SQLX$ = SQLX$ + " ,0"
                SQLX$ = SQLX$ + " , 'ESCOST/A32'"
                SQLX$ = SQLX$ + " , ''"
                SQLX$ = SQLX$ + " , 'S'"
                SQLX$ = SQLX$ + " , ''"
                SQLX$ = SQLX$ + " , ''"
                SQLX$ = SQLX$ + " , ''"
                SQLX$ = SQLX$ + " , ''"
                SQLX$ = SQLX$ + " )"
                FLEXCONN.Execute (SQLX$)
            
            '
                SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
                SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
                SQLX$ = SQLX$ + " ,0"
                SQLX$ = SQLX$ + " ,0"
                SQLX$ = SQLX$ + " , 'NV/A0'"
                SQLX$ = SQLX$ + " , ''"
                SQLX$ = SQLX$ + " , ''"
                SQLX$ = SQLX$ + " , ''"
                SQLX$ = SQLX$ + " , ''"
                SQLX$ = SQLX$ + " , ''"
                SQLX$ = SQLX$ + " , ''"
                SQLX$ = SQLX$ + " )"
                FLEXCONN.Execute (SQLX$)
            
            '
            End If
            '|FIN CTP.ESCOST

            '|INICIO CTP.VALV
            '
            CONDI$ = " LISTA =   'CTP.VALV'"
            EXILISTA% = CantRegistros&("ENCAPRFUN", CONDI$, "NOMESS")
            If EXILISTA% > 0 Then
              FE% = 16547
              FEBD% = CVINT(VALOBADA("ENCAPRFUN", "FECHEMISI", CONDI$, "NOMESS"))
              GENERAREST% = 0
              If FEBD% < FE% Then
                IDLISTAX& = XVALO(VALOBADA("ENCAPRFUN", "ID_LISTA", CONDI$, "NOMESS"))
                Call BORRAREGISTROS("ENCAPRFUN", "ID_LISTA = " & IDLISTAX&, REGAF&, "NOMESS")
                Call BORRAREGISTROS("ENCAPRF", "ID_LISTA = " & IDLISTAX&, REGAF&, "NOMESS")
                GENERAREST% = 1
              End If
            Else
              GENERAREST% = 1
              CONDI$ = " ID_LISTA  > 0 "
              IDLISTAX& = 1 + XVALO(VALOBADA("ENCAPRFUN", "MAX(ID_LISTA)", CONDI$, "NOMESS"))
            End If
            If GENERAREST% > 0 Then
              SQLX$ = "INSERT INTO  ENCAPRFUN (ID_LISTA ,LISTA, Descripcion, ADDREGISTROS, USAFORMULA, BORRAR, INSERTAR, NOTILDE "
              SQLX$ = SQLX$ + ", CANTICAMPOS, CodiEmpr, USARECORD, FECHEMISI, TITURECORD, ConsultaSql)"
              SQLX$ = SQLX$ + "VALUES ( " & IDLISTAX&
              SQLX$ = SQLX$ + ",'CTP.VALV'"
              SQLX$ = SQLX$ + ",'VALVULA'"
              SQLX$ = SQLX$ + ",1"
              SQLX$ = SQLX$ + ",1"
              SQLX$ = SQLX$ + " ,0"
              SQLX$ = SQLX$ + " ,0"
              SQLX$ = SQLX$ + " ,0"
              SQLX$ = SQLX$ + " ,4"
              SQLX$ = SQLX$ + " ,0"
              SQLX$ = SQLX$ + " ,0"
              SQLX$ = SQLX$ + " ,'03/10/19'"
              SQLX$ = SQLX$ + " ,''"
              SQLX$ = SQLX$ + " ,'SELECT CAMPO1,CAMPO2,CAMPO3,CAMPO4 FROM DATASQL WITH(NOLOCK)'"
              SQLX$ = SQLX$ + ")"
              FLEXCONN.Execute (SQLX$)
                SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
                SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
                SQLX$ = SQLX$ + " ,0"
                SQLX$ = SQLX$ + " ,0"
                SQLX$ = SQLX$ + " , 'ID/F8'"
                SQLX$ = SQLX$ + " , ''"
                SQLX$ = SQLX$ + " , ''"
                SQLX$ = SQLX$ + " , ''"
                SQLX$ = SQLX$ + " , ''"
                SQLX$ = SQLX$ + " , ''"
                SQLX$ = SQLX$ + " , ''"
                SQLX$ = SQLX$ + " )"
                FLEXCONN.Execute (SQLX$)
            
            '
                SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
                SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
                SQLX$ = SQLX$ + " ,0"
                SQLX$ = SQLX$ + " ,0"
                SQLX$ = SQLX$ + " , 'VALV./A32'"
                SQLX$ = SQLX$ + " , ''"
                SQLX$ = SQLX$ + " , 'S'"
                SQLX$ = SQLX$ + " , ''"
                SQLX$ = SQLX$ + " , ''"
                SQLX$ = SQLX$ + " , ''"
                SQLX$ = SQLX$ + " , ''"
                SQLX$ = SQLX$ + " )"
                FLEXCONN.Execute (SQLX$)
            
                SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
                SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
                SQLX$ = SQLX$ + " ,0"
                SQLX$ = SQLX$ + " ,0"
                SQLX$ = SQLX$ + " , 'V Desc Seg (cm)./F14'"
                SQLX$ = SQLX$ + " , ''"
                SQLX$ = SQLX$ + " , 'S'"
                SQLX$ = SQLX$ + " , ''"
                SQLX$ = SQLX$ + " , ''"
                SQLX$ = SQLX$ + " , ''"
                SQLX$ = SQLX$ + " , ''"
                SQLX$ = SQLX$ + " )"
                FLEXCONN.Execute (SQLX$)
            
            '
                SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
                SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
                SQLX$ = SQLX$ + " ,0"
                SQLX$ = SQLX$ + " ,0"
                SQLX$ = SQLX$ + " , 'NV/A0'"
                SQLX$ = SQLX$ + " , ''"
                SQLX$ = SQLX$ + " , ''"
                SQLX$ = SQLX$ + " , ''"
                SQLX$ = SQLX$ + " , ''"
                SQLX$ = SQLX$ + " , ''"
                SQLX$ = SQLX$ + " , ''"
                SQLX$ = SQLX$ + " )"
                FLEXCONN.Execute (SQLX$)
            
            '
            End If
            '|FIN CTP.VALV


            '|INICIO CTP.GRAMAJE
            '
            CONDI$ = " LISTA =   'CTP.GRAMAJE'"
            EXILISTA% = CantRegistros&("ENCAPRFUN", CONDI$, "NOMESS")
            If EXILISTA% > 0 Then
              FE% = 16547
              FEBD% = CVINT(VALOBADA("ENCAPRFUN", "FECHEMISI", CONDI$, "NOMESS"))
              GENERAREST% = 0
              If FEBD% < FE% Then
                IDLISTAX& = XVALO(VALOBADA("ENCAPRFUN", "ID_LISTA", CONDI$, "NOMESS"))
                Call BORRAREGISTROS("ENCAPRFUN", "ID_LISTA = " & IDLISTAX&, REGAF&, "NOMESS")
                Call BORRAREGISTROS("ENCAPRF", "ID_LISTA = " & IDLISTAX&, REGAF&, "NOMESS")
                GENERAREST% = 1
              End If
            Else
              GENERAREST% = 1
              CONDI$ = " ID_LISTA  > 0 "
              IDLISTAX& = 1 + XVALO(VALOBADA("ENCAPRFUN", "MAX(ID_LISTA)", CONDI$, "NOMESS"))
            End If
            If GENERAREST% > 0 Then
              SQLX$ = "INSERT INTO  ENCAPRFUN (ID_LISTA ,LISTA, Descripcion, ADDREGISTROS, USAFORMULA, BORRAR, INSERTAR, NOTILDE "
              SQLX$ = SQLX$ + ", CANTICAMPOS, CodiEmpr, USARECORD, FECHEMISI, TITURECORD, ConsultaSql)"
              SQLX$ = SQLX$ + "VALUES ( " & IDLISTAX&
              SQLX$ = SQLX$ + ",'CTP.GRAMAJE'"
              SQLX$ = SQLX$ + ",'GRAMAJE'"
              SQLX$ = SQLX$ + ",1"
              SQLX$ = SQLX$ + ",1"
              SQLX$ = SQLX$ + " ,0"
              SQLX$ = SQLX$ + " ,0"
              SQLX$ = SQLX$ + " ,0"
              SQLX$ = SQLX$ + " ,4"
              SQLX$ = SQLX$ + " ,0"
              SQLX$ = SQLX$ + " ,0"
              SQLX$ = SQLX$ + " ,'03/10/19'"
              SQLX$ = SQLX$ + " ,''"
              SQLX$ = SQLX$ + " ,'SELECT CAMPO1,CAMPO2,CAMPO3 FROM DATASQL WITH(NOLOCK)'"
              SQLX$ = SQLX$ + ")"
              FLEXCONN.Execute (SQLX$)
                SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
                SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
                SQLX$ = SQLX$ + " ,0"
                SQLX$ = SQLX$ + " ,0"
                SQLX$ = SQLX$ + " , 'ID/F8'"
                SQLX$ = SQLX$ + " , ''"
                SQLX$ = SQLX$ + " , ''"
                SQLX$ = SQLX$ + " , ''"
                SQLX$ = SQLX$ + " , ''"
                SQLX$ = SQLX$ + " , ''"
                SQLX$ = SQLX$ + " , ''"
                SQLX$ = SQLX$ + " )"
                FLEXCONN.Execute (SQLX$)
            
            '
                SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
                SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
                SQLX$ = SQLX$ + " ,0"
                SQLX$ = SQLX$ + " ,0"
                SQLX$ = SQLX$ + " , 'GRAMAJE/A32'"
                SQLX$ = SQLX$ + " , ''"
                SQLX$ = SQLX$ + " , 'S'"
                SQLX$ = SQLX$ + " , ''"
                SQLX$ = SQLX$ + " , ''"
                SQLX$ = SQLX$ + " , ''"
                SQLX$ = SQLX$ + " , ''"
                SQLX$ = SQLX$ + " )"
                FLEXCONN.Execute (SQLX$)
            
            '
                SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
                SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
                SQLX$ = SQLX$ + " ,0"
                SQLX$ = SQLX$ + " ,0"
                SQLX$ = SQLX$ + " , 'NV/A0 '"
                SQLX$ = SQLX$ + " , ''"
                SQLX$ = SQLX$ + " , ''"
                SQLX$ = SQLX$ + " , ''"
                SQLX$ = SQLX$ + " , ''"
                SQLX$ = SQLX$ + " , ''"
                SQLX$ = SQLX$ + " , ''"
                SQLX$ = SQLX$ + " )"
                FLEXCONN.Execute (SQLX$)
            
            '
            End If
            '|FIN CTP.GRAMAJE
                '|INICIO CTP.TAPA_GR
                '
                CONDI$ = " LISTA =   'CTP.TAPA_GR'"
                EXILISTA% = CantRegistros&("ENCAPRFUN", CONDI$, "NOMESS")
                If EXILISTA% > 0 Then
                  FE% = 16547
                  FEBD% = CVINT(VALOBADA("ENCAPRFUN", "FECHEMISI", CONDI$, "NOMESS"))
                  GENERAREST% = 0
                  If FEBD% < FE% Then
                    IDLISTAX& = XVALO(VALOBADA("ENCAPRFUN", "ID_LISTA", CONDI$, "NOMESS"))
                    Call BORRAREGISTROS("ENCAPRFUN", "ID_LISTA = " & IDLISTAX&, REGAF&, "NOMESS")
                    Call BORRAREGISTROS("ENCAPRF", "ID_LISTA = " & IDLISTAX&, REGAF&, "NOMESS")
                    GENERAREST% = 1
                  End If
                Else
                  GENERAREST% = 1
                  CONDI$ = " ID_LISTA  > 0 "
                  IDLISTAX& = 1 + XVALO(VALOBADA("ENCAPRFUN", "MAX(ID_LISTA)", CONDI$, "NOMESS"))
                End If
                If GENERAREST% > 0 Then
                  SQLX$ = "INSERT INTO  ENCAPRFUN (ID_LISTA ,LISTA, Descripcion, ADDREGISTROS, USAFORMULA, BORRAR, INSERTAR, NOTILDE "
                  SQLX$ = SQLX$ + ", CANTICAMPOS, CodiEmpr, USARECORD, FECHEMISI, TITURECORD, ConsultaSql)"
                  SQLX$ = SQLX$ + "VALUES ( " & IDLISTAX&
                  SQLX$ = SQLX$ + ",'CTP.TAPA_GR'"
                  SQLX$ = SQLX$ + ",'TAPA GR'"
                  SQLX$ = SQLX$ + ",1"
                  SQLX$ = SQLX$ + ",1"
                  SQLX$ = SQLX$ + " ,0"
                  SQLX$ = SQLX$ + " ,0"
                  SQLX$ = SQLX$ + " ,0"
                  SQLX$ = SQLX$ + " ,4"
                  SQLX$ = SQLX$ + " ,0"
                  SQLX$ = SQLX$ + " ,0"
                  SQLX$ = SQLX$ + " ,'03/10/19'"
                  SQLX$ = SQLX$ + " ,''"
                  SQLX$ = SQLX$ + " ,'SELECT CAMPO1,CAMPO2,CAMPO3 FROM DATASQL WITH(NOLOCK)'"
                  SQLX$ = SQLX$ + ")"
                  FLEXCONN.Execute (SQLX$)
                    SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
                    SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
                    SQLX$ = SQLX$ + " ,0"
                    SQLX$ = SQLX$ + " ,0"
                    SQLX$ = SQLX$ + " , 'ID/F8'"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " , ' '"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " )"
                    FLEXCONN.Execute (SQLX$)
                
                '
                    SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
                    SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
                    SQLX$ = SQLX$ + " ,0"
                    SQLX$ = SQLX$ + " ,0"
                    SQLX$ = SQLX$ + " , 'TAPA GR/A32'"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " , 'S'"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " )"
                    FLEXCONN.Execute (SQLX$)
                
                '
                    SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
                    SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
                    SQLX$ = SQLX$ + " ,0"
                    SQLX$ = SQLX$ + " ,0"
                    SQLX$ = SQLX$ + " , 'NV/A0'"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " )"
                    FLEXCONN.Execute (SQLX$)
                
                '
                End If
                '|FIN CTP.TAPA_GR

                '|INICIO CTP.POLLERA
                '
                CONDI$ = " LISTA =   'CTP.POLLERA'"
                EXILISTA% = CantRegistros&("ENCAPRFUN", CONDI$, "NOMESS")
                If EXILISTA% > 0 Then
                  FE% = 16547
                  FEBD% = CVINT(VALOBADA("ENCAPRFUN", "FECHEMISI", CONDI$, "NOMESS"))
                  GENERAREST% = 0
                  If FEBD% < FE% Then
                    IDLISTAX& = XVALO(VALOBADA("ENCAPRFUN", "ID_LISTA", CONDI$, "NOMESS"))
                    Call BORRAREGISTROS("ENCAPRFUN", "ID_LISTA = " & IDLISTAX&, REGAF&, "NOMESS")
                    Call BORRAREGISTROS("ENCAPRF", "ID_LISTA = " & IDLISTAX&, REGAF&, "NOMESS")
                    GENERAREST% = 1
                  End If
                Else
                  GENERAREST% = 1
                  CONDI$ = " ID_LISTA  > 0 "
                  IDLISTAX& = 1 + XVALO(VALOBADA("ENCAPRFUN", "MAX(ID_LISTA)", CONDI$, "NOMESS"))
                End If
                If GENERAREST% > 0 Then
                  SQLX$ = "INSERT INTO  ENCAPRFUN (ID_LISTA ,LISTA, Descripcion, ADDREGISTROS, USAFORMULA, BORRAR, INSERTAR, NOTILDE "
                  SQLX$ = SQLX$ + ", CANTICAMPOS, CodiEmpr, USARECORD, FECHEMISI, TITURECORD, ConsultaSql)"
                  SQLX$ = SQLX$ + "VALUES ( " & IDLISTAX&
                  SQLX$ = SQLX$ + ",'CTP.POLLERA'"
                  SQLX$ = SQLX$ + ",'POLLERA'"
                  SQLX$ = SQLX$ + ",1"
                  SQLX$ = SQLX$ + ",1"
                  SQLX$ = SQLX$ + " ,0"
                  SQLX$ = SQLX$ + " ,0"
                  SQLX$ = SQLX$ + " ,0"
                  SQLX$ = SQLX$ + " ,4"
                  SQLX$ = SQLX$ + " ,0"
                  SQLX$ = SQLX$ + " ,0"
                  SQLX$ = SQLX$ + " ,'03/10/19'"
                  SQLX$ = SQLX$ + " ,''"
                  SQLX$ = SQLX$ + " ,'SELECT CAMPO1,CAMPO2,CAMPO3 FROM DATASQL WITH(NOLOCK)'"
                  SQLX$ = SQLX$ + ")"
                  FLEXCONN.Execute (SQLX$)
                    SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
                    SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
                    SQLX$ = SQLX$ + " ,0"
                    SQLX$ = SQLX$ + " ,0"
                    SQLX$ = SQLX$ + " , 'ID/F8'"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " )"
                    FLEXCONN.Execute (SQLX$)
                
                '
                    SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
                    SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
                    SQLX$ = SQLX$ + " ,0"
                    SQLX$ = SQLX$ + " ,0"
                    SQLX$ = SQLX$ + " , 'POLLERA/A32'"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " , 'S'"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " )"
                    FLEXCONN.Execute (SQLX$)
                
                '
                    SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
                    SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
                    SQLX$ = SQLX$ + " ,0"
                    SQLX$ = SQLX$ + " ,0"
                    SQLX$ = SQLX$ + " , 'NV/A0'"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " )"
                    FLEXCONN.Execute (SQLX$)
                
                '
                End If
                '|FIN CTP.POLLERA
    
    
                 '|INICIO CTP.PD
                '
                CONDI$ = " LISTA =   'CTP.PD'"
                EXILISTA% = CantRegistros&("ENCAPRFUN", CONDI$, "NOMESS")
                If EXILISTA% > 0 Then
                  FE% = 16547
                  FEBD% = CVINT(VALOBADA("ENCAPRFUN", "FECHEMISI", CONDI$, "NOMESS"))
                  GENERAREST% = 0
                  If FEBD% < FE% Then
                    IDLISTAX& = XVALO(VALOBADA("ENCAPRFUN", "ID_LISTA", CONDI$, "NOMESS"))
                    Call BORRAREGISTROS("ENCAPRFUN", "ID_LISTA = " & IDLISTAX&, REGAF&, "NOMESS")
                    Call BORRAREGISTROS("ENCAPRF", "ID_LISTA = " & IDLISTAX&, REGAF&, "NOMESS")
                    GENERAREST% = 1
                  End If
                Else
                  GENERAREST% = 1
                  CONDI$ = " ID_LISTA  > 0 "
                  IDLISTAX& = 1 + XVALO(VALOBADA("ENCAPRFUN", "MAX(ID_LISTA)", CONDI$, "NOMESS"))
                End If
                If GENERAREST% > 0 Then
                  SQLX$ = "INSERT INTO  ENCAPRFUN (ID_LISTA ,LISTA, Descripcion, ADDREGISTROS, USAFORMULA, BORRAR, INSERTAR, NOTILDE "
                  SQLX$ = SQLX$ + ", CANTICAMPOS, CodiEmpr, USARECORD, FECHEMISI, TITURECORD, ConsultaSql)"
                  SQLX$ = SQLX$ + "VALUES ( " & IDLISTAX&
                  SQLX$ = SQLX$ + ",'CTP.PD'"
                  SQLX$ = SQLX$ + ",'PD'"
                  SQLX$ = SQLX$ + ",1"
                  SQLX$ = SQLX$ + ",1"
                  SQLX$ = SQLX$ + " ,0"
                  SQLX$ = SQLX$ + " ,0"
                  SQLX$ = SQLX$ + " ,0"
                  SQLX$ = SQLX$ + " ,4"
                  SQLX$ = SQLX$ + " ,0"
                  SQLX$ = SQLX$ + " ,0"
                  SQLX$ = SQLX$ + " ,'03/10/19'"
                  SQLX$ = SQLX$ + " ,''"
                  SQLX$ = SQLX$ + " ,'SELECT CAMPO1,CAMPO2,CAMPO3 FROM DATASQL WITH(NOLOCK)'"
                  SQLX$ = SQLX$ + ")"
                  FLEXCONN.Execute (SQLX$)
                    SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
                    SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
                    SQLX$ = SQLX$ + " ,0"
                    SQLX$ = SQLX$ + " ,0"
                    SQLX$ = SQLX$ + " , 'ID/F8'"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " )"
                    FLEXCONN.Execute (SQLX$)
                
                '
                    SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
                    SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
                    SQLX$ = SQLX$ + " ,0"
                    SQLX$ = SQLX$ + " ,0"
                    SQLX$ = SQLX$ + " , 'PD/A32'"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " , 'S'"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " )"
                    FLEXCONN.Execute (SQLX$)
                
                '
                    SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
                    SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
                    SQLX$ = SQLX$ + " ,0"
                    SQLX$ = SQLX$ + " ,0"
                    SQLX$ = SQLX$ + " , 'NV/A0'"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " )"
                    FLEXCONN.Execute (SQLX$)
                
                '
                End If
                '|FIN CTP.PD
                
                
                 '|INICIO CTP.PDT
                '
                CONDI$ = " LISTA =   'CTP.PDT'"
                EXILISTA% = CantRegistros&("ENCAPRFUN", CONDI$, "NOMESS")
                If EXILISTA% > 0 Then
                  FE% = 16547
                  FEBD% = CVINT(VALOBADA("ENCAPRFUN", "FECHEMISI", CONDI$, "NOMESS"))
                  GENERAREST% = 0
                  If FEBD% < FE% Then
                    IDLISTAX& = XVALO(VALOBADA("ENCAPRFUN", "ID_LISTA", CONDI$, "NOMESS"))
                    Call BORRAREGISTROS("ENCAPRFUN", "ID_LISTA = " & IDLISTAX&, REGAF&, "NOMESS")
                    Call BORRAREGISTROS("ENCAPRF", "ID_LISTA = " & IDLISTAX&, REGAF&, "NOMESS")
                    GENERAREST% = 1
                  End If
                Else
                  GENERAREST% = 1
                  CONDI$ = " ID_LISTA  > 0 "
                  IDLISTAX& = 1 + XVALO(VALOBADA("ENCAPRFUN", "MAX(ID_LISTA)", CONDI$, "NOMESS"))
                End If
                If GENERAREST% > 0 Then
                  SQLX$ = "INSERT INTO  ENCAPRFUN (ID_LISTA ,LISTA, Descripcion, ADDREGISTROS, USAFORMULA, BORRAR, INSERTAR, NOTILDE "
                  SQLX$ = SQLX$ + ", CANTICAMPOS, CodiEmpr, USARECORD, FECHEMISI, TITURECORD, ConsultaSql)"
                  SQLX$ = SQLX$ + "VALUES ( " & IDLISTAX&
                  SQLX$ = SQLX$ + ",'CTP.PDT'"
                  SQLX$ = SQLX$ + ",'PDT'"
                  SQLX$ = SQLX$ + ",1"
                  SQLX$ = SQLX$ + ",1"
                  SQLX$ = SQLX$ + " ,0"
                  SQLX$ = SQLX$ + " ,0"
                  SQLX$ = SQLX$ + " ,0"
                  SQLX$ = SQLX$ + " ,4"
                  SQLX$ = SQLX$ + " ,0"
                  SQLX$ = SQLX$ + " ,0"
                  SQLX$ = SQLX$ + " ,'03/10/19'"
                  SQLX$ = SQLX$ + " ,''"
                  SQLX$ = SQLX$ + " ,'SELECT CAMPO1,CAMPO2,CAMPO3 FROM DATASQL WITH(NOLOCK)'"
                  SQLX$ = SQLX$ + ")"
                  FLEXCONN.Execute (SQLX$)
                    SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
                    SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
                    SQLX$ = SQLX$ + " ,0"
                    SQLX$ = SQLX$ + " ,0"
                    SQLX$ = SQLX$ + " , 'ID/F8'"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " )"
                    FLEXCONN.Execute (SQLX$)
                
                '
                    SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
                    SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
                    SQLX$ = SQLX$ + " ,0"
                    SQLX$ = SQLX$ + " ,0"
                    SQLX$ = SQLX$ + " , 'PDT/A48'"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " , 'S'"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " )"
                    FLEXCONN.Execute (SQLX$)
                
                '
                    SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
                    SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
                    SQLX$ = SQLX$ + " ,0"
                    SQLX$ = SQLX$ + " ,0"
                    SQLX$ = SQLX$ + " , 'NV/A0'"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " )"
                    FLEXCONN.Execute (SQLX$)
                
                '
                End If
                '|FIN CTP.PDT
               
                 '|INICIO CTP.PDTA
                '
                CONDI$ = " LISTA =   'CTP.PDTA'"
                EXILISTA% = CantRegistros&("ENCAPRFUN", CONDI$, "NOMESS")
                If EXILISTA% > 0 Then
                  FE% = 16547
                  FEBD% = CVINT(VALOBADA("ENCAPRFUN", "FECHEMISI", CONDI$, "NOMESS"))
                  GENERAREST% = 0
                  If FEBD% < FE% Then
                    IDLISTAX& = XVALO(VALOBADA("ENCAPRFUN", "ID_LISTA", CONDI$, "NOMESS"))
                    Call BORRAREGISTROS("ENCAPRFUN", "ID_LISTA = " & IDLISTAX&, REGAF&, "NOMESS")
                    Call BORRAREGISTROS("ENCAPRF", "ID_LISTA = " & IDLISTAX&, REGAF&, "NOMESS")
                    GENERAREST% = 1
                  End If
                Else
                  GENERAREST% = 1
                  CONDI$ = " ID_LISTA  > 0 "
                  IDLISTAX& = 1 + XVALO(VALOBADA("ENCAPRFUN", "MAX(ID_LISTA)", CONDI$, "NOMESS"))
                End If
                If GENERAREST% > 0 Then
                  SQLX$ = "INSERT INTO  ENCAPRFUN (ID_LISTA ,LISTA, Descripcion, ADDREGISTROS, USAFORMULA, BORRAR, INSERTAR, NOTILDE "
                  SQLX$ = SQLX$ + ", CANTICAMPOS, CodiEmpr, USARECORD, FECHEMISI, TITURECORD, ConsultaSql)"
                  SQLX$ = SQLX$ + "VALUES ( " & IDLISTAX&
                  SQLX$ = SQLX$ + ",'CTP.PDTA'"
                  SQLX$ = SQLX$ + ",'PDTA'"
                  SQLX$ = SQLX$ + ",1"
                  SQLX$ = SQLX$ + ",1"
                  SQLX$ = SQLX$ + " ,0"
                  SQLX$ = SQLX$ + " ,0"
                  SQLX$ = SQLX$ + " ,0"
                  SQLX$ = SQLX$ + " ,4"
                  SQLX$ = SQLX$ + " ,0"
                  SQLX$ = SQLX$ + " ,0"
                  SQLX$ = SQLX$ + " ,'03/10/19'"
                  SQLX$ = SQLX$ + " ,''"
                  SQLX$ = SQLX$ + " ,'SELECT CAMPO1,CAMPO2,CAMPO3 FROM DATASQL WITH(NOLOCK)'"
                  SQLX$ = SQLX$ + ")"
                  FLEXCONN.Execute (SQLX$)
                    SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
                    SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
                    SQLX$ = SQLX$ + " ,0"
                    SQLX$ = SQLX$ + " ,0"
                    SQLX$ = SQLX$ + " , 'ID/F8'"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " )"
                    FLEXCONN.Execute (SQLX$)
                
                '
                    SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
                    SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
                    SQLX$ = SQLX$ + " ,0"
                    SQLX$ = SQLX$ + " ,0"
                    SQLX$ = SQLX$ + " , 'PDTA/A32'"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " , 'S'"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " )"
                    FLEXCONN.Execute (SQLX$)
                
                '
                    SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
                    SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
                    SQLX$ = SQLX$ + " ,0"
                    SQLX$ = SQLX$ + " ,0"
                    SQLX$ = SQLX$ + " , 'NV/A0'"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " )"
                    FLEXCONN.Execute (SQLX$)
                
                '
                End If
                '|FIN CTP.PDTA
  
  
                '|INICIO CTP.COMP
                '
                CONDI$ = " LISTA =   'CTP.COMP'"
                EXILISTA% = CantRegistros&("ENCAPRFUN", CONDI$, "NOMESS")
                If EXILISTA% > 0 Then
                  FE% = 16547
                  FEBD% = CVINT(VALOBADA("ENCAPRFUN", "FECHEMISI", CONDI$, "NOMESS"))
                  GENERAREST% = 0
                  If FEBD% < FE% Then
                    IDLISTAX& = XVALO(VALOBADA("ENCAPRFUN", "ID_LISTA", CONDI$, "NOMESS"))
                    Call BORRAREGISTROS("ENCAPRFUN", "ID_LISTA = " & IDLISTAX&, REGAF&, "NOMESS")
                    Call BORRAREGISTROS("ENCAPRF", "ID_LISTA = " & IDLISTAX&, REGAF&, "NOMESS")
                    GENERAREST% = 1
                  End If
                Else
                  GENERAREST% = 1
                  CONDI$ = " ID_LISTA  > 0 "
                  IDLISTAX& = 1 + XVALO(VALOBADA("ENCAPRFUN", "MAX(ID_LISTA)", CONDI$, "NOMESS"))
                End If
                If GENERAREST% > 0 Then
                  SQLX$ = "INSERT INTO  ENCAPRFUN (ID_LISTA ,LISTA, Descripcion, ADDREGISTROS, USAFORMULA, BORRAR, INSERTAR, NOTILDE "
                  SQLX$ = SQLX$ + ", CANTICAMPOS, CodiEmpr, USARECORD, FECHEMISI, TITURECORD, ConsultaSql)"
                  SQLX$ = SQLX$ + "VALUES ( " & IDLISTAX&
                  SQLX$ = SQLX$ + ",'CTP.COMP'"
                  SQLX$ = SQLX$ + ",'COMP'"
                  SQLX$ = SQLX$ + ",1"
                  SQLX$ = SQLX$ + ",1"
                  SQLX$ = SQLX$ + " ,0"
                  SQLX$ = SQLX$ + " ,0"
                  SQLX$ = SQLX$ + " ,0"
                  SQLX$ = SQLX$ + " ,4"
                  SQLX$ = SQLX$ + " ,0"
                  SQLX$ = SQLX$ + " ,0"
                  SQLX$ = SQLX$ + " ,'03/10/19'"
                  SQLX$ = SQLX$ + " ,''"
                  SQLX$ = SQLX$ + " ,'SELECT CAMPO1,CAMPO2,CAMPO3 FROM DATASQL WITH(NOLOCK)'"
                  SQLX$ = SQLX$ + ")"
                  FLEXCONN.Execute (SQLX$)
                    SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
                    SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
                    SQLX$ = SQLX$ + " ,0"
                    SQLX$ = SQLX$ + " ,0"
                    SQLX$ = SQLX$ + " , 'ID/F8'"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " )"
                    FLEXCONN.Execute (SQLX$)
                
                '
                    SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
                    SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
                    SQLX$ = SQLX$ + " ,0"
                    SQLX$ = SQLX$ + " ,0"
                    SQLX$ = SQLX$ + " , 'COMP/F10'"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " , 'S'"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " )"
                    FLEXCONN.Execute (SQLX$)
                
                '
                    SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
                    SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
                    SQLX$ = SQLX$ + " ,0"
                    SQLX$ = SQLX$ + " ,0"
                    SQLX$ = SQLX$ + " , 'NV/A0'"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " )"
                    FLEXCONN.Execute (SQLX$)
                
                '
                End If
                '|FIN CTP.COMP
                
                  '|INICIO CTP.LINT
                '
                CONDI$ = " LISTA =   'CTP.LINT'"
                EXILISTA% = CantRegistros&("ENCAPRFUN", CONDI$, "NOMESS")
                If EXILISTA% > 0 Then
                  FE% = 16547
                  FEBD% = CVINT(VALOBADA("ENCAPRFUN", "FECHEMISI", CONDI$, "NOMESS"))
                  GENERAREST% = 0
                  If FEBD% < FE% Then
                    IDLISTAX& = XVALO(VALOBADA("ENCAPRFUN", "ID_LISTA", CONDI$, "NOMESS"))
                    Call BORRAREGISTROS("ENCAPRFUN", "ID_LISTA = " & IDLISTAX&, REGAF&, "NOMESS")
                    Call BORRAREGISTROS("ENCAPRF", "ID_LISTA = " & IDLISTAX&, REGAF&, "NOMESS")
                    GENERAREST% = 1
                  End If
                Else
                  GENERAREST% = 1
                  CONDI$ = " ID_LISTA  > 0 "
                  IDLISTAX& = 1 + XVALO(VALOBADA("ENCAPRFUN", "MAX(ID_LISTA)", CONDI$, "NOMESS"))
                End If
                If GENERAREST% > 0 Then
                  SQLX$ = "INSERT INTO  ENCAPRFUN (ID_LISTA ,LISTA, Descripcion, ADDREGISTROS, USAFORMULA, BORRAR, INSERTAR, NOTILDE "
                  SQLX$ = SQLX$ + ", CANTICAMPOS, CodiEmpr, USARECORD, FECHEMISI, TITURECORD, ConsultaSql)"
                  SQLX$ = SQLX$ + "VALUES ( " & IDLISTAX&
                  SQLX$ = SQLX$ + ",'CTP.LINT'"
                  SQLX$ = SQLX$ + ",'LINT'"
                  SQLX$ = SQLX$ + ",1"
                  SQLX$ = SQLX$ + ",1"
                  SQLX$ = SQLX$ + " ,0"
                  SQLX$ = SQLX$ + " ,0"
                  SQLX$ = SQLX$ + " ,0"
                  SQLX$ = SQLX$ + " ,4"
                  SQLX$ = SQLX$ + " ,0"
                  SQLX$ = SQLX$ + " ,0"
                  SQLX$ = SQLX$ + " ,'03/10/19'"
                  SQLX$ = SQLX$ + " ,''"
                  SQLX$ = SQLX$ + " ,'SELECT CAMPO1,CAMPO2,CAMPO3 FROM DATASQL WITH(NOLOCK)'"
                  SQLX$ = SQLX$ + ")"
                  FLEXCONN.Execute (SQLX$)
                    SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
                    SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
                    SQLX$ = SQLX$ + " ,0"
                    SQLX$ = SQLX$ + " ,0"
                    SQLX$ = SQLX$ + " , 'ID/F8'"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " )"
                    FLEXCONN.Execute (SQLX$)
                
                '
                    SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
                    SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
                    SQLX$ = SQLX$ + " ,0"
                    SQLX$ = SQLX$ + " ,0"
                    SQLX$ = SQLX$ + " , 'LINT/A32'"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " , 'S'"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " )"
                    FLEXCONN.Execute (SQLX$)
                
                '
                    SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
                    SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
                    SQLX$ = SQLX$ + " ,0"
                    SQLX$ = SQLX$ + " ,0"
                    SQLX$ = SQLX$ + " , 'NV/A0'"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " )"
                    FLEXCONN.Execute (SQLX$)
                
                '
                End If
                '|FIN CTP.LINT
              
                '|INICIO CTP.EMB
                '
                CONDI$ = " LISTA =   'CTP.EMB'"
                EXILISTA% = CantRegistros&("ENCAPRFUN", CONDI$, "NOMESS")
                If EXILISTA% > 0 Then
                  FE% = 16547
                  FEBD% = CVINT(VALOBADA("ENCAPRFUN", "FECHEMISI", CONDI$, "NOMESS"))
                  GENERAREST% = 0
                  If FEBD% < FE% Then
                    IDLISTAX& = XVALO(VALOBADA("ENCAPRFUN", "ID_LISTA", CONDI$, "NOMESS"))
                    Call BORRAREGISTROS("ENCAPRFUN", "ID_LISTA = " & IDLISTAX&, REGAF&, "NOMESS")
                    Call BORRAREGISTROS("ENCAPRF", "ID_LISTA = " & IDLISTAX&, REGAF&, "NOMESS")
                    GENERAREST% = 1
                  End If
                Else
                  GENERAREST% = 1
                  CONDI$ = " ID_LISTA  > 0 "
                  IDLISTAX& = 1 + XVALO(VALOBADA("ENCAPRFUN", "MAX(ID_LISTA)", CONDI$, "NOMESS"))
                End If
                If GENERAREST% > 0 Then
                  SQLX$ = "INSERT INTO  ENCAPRFUN (ID_LISTA ,LISTA, Descripcion, ADDREGISTROS, USAFORMULA, BORRAR, INSERTAR, NOTILDE "
                  SQLX$ = SQLX$ + ", CANTICAMPOS, CodiEmpr, USARECORD, FECHEMISI, TITURECORD, ConsultaSql)"
                  SQLX$ = SQLX$ + "VALUES ( " & IDLISTAX&
                  SQLX$ = SQLX$ + ",'CTP.EMB'"
                  SQLX$ = SQLX$ + ",'EMB'"
                  SQLX$ = SQLX$ + ",1"
                  SQLX$ = SQLX$ + ",1"
                  SQLX$ = SQLX$ + " ,0"
                  SQLX$ = SQLX$ + " ,0"
                  SQLX$ = SQLX$ + " ,0"
                  SQLX$ = SQLX$ + " ,4"
                  SQLX$ = SQLX$ + " ,0"
                  SQLX$ = SQLX$ + " ,0"
                  SQLX$ = SQLX$ + " ,'03/10/19'"
                  SQLX$ = SQLX$ + " ,''"
                  SQLX$ = SQLX$ + " ,'SELECT CAMPO1,CAMPO2,CAMPO3 FROM DATASQL WITH(NOLOCK)'"
                  SQLX$ = SQLX$ + ")"
                  FLEXCONN.Execute (SQLX$)
                    SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
                    SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
                    SQLX$ = SQLX$ + " ,0"
                    SQLX$ = SQLX$ + " ,0"
                    SQLX$ = SQLX$ + " , 'ID/F8'"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " )"
                    FLEXCONN.Execute (SQLX$)
                
                '
                    SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
                    SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
                    SQLX$ = SQLX$ + " ,0"
                    SQLX$ = SQLX$ + " ,0"
                    SQLX$ = SQLX$ + " , 'EMB/F10'"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " , 'S'"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " )"
                    FLEXCONN.Execute (SQLX$)
                
                '
                    SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
                    SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
                    SQLX$ = SQLX$ + " ,0"
                    SQLX$ = SQLX$ + " ,0"
                    SQLX$ = SQLX$ + " , 'NV/A0'"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " )"
                    FLEXCONN.Execute (SQLX$)
                
                '
                End If
                '|FIN CTP.EMB
                
                
                '|INICIO CTP.ETIQ
                '
                CONDI$ = " LISTA =   'CTP.ETIQ'"
                EXILISTA% = CantRegistros&("ENCAPRFUN", CONDI$, "NOMESS")
                If EXILISTA% > 0 Then
                  FE% = 16547
                  FEBD% = CVINT(VALOBADA("ENCAPRFUN", "FECHEMISI", CONDI$, "NOMESS"))
                  GENERAREST% = 0
                  If FEBD% < FE% Then
                    IDLISTAX& = XVALO(VALOBADA("ENCAPRFUN", "ID_LISTA", CONDI$, "NOMESS"))
                    Call BORRAREGISTROS("ENCAPRFUN", "ID_LISTA = " & IDLISTAX&, REGAF&, "NOMESS")
                    Call BORRAREGISTROS("ENCAPRF", "ID_LISTA = " & IDLISTAX&, REGAF&, "NOMESS")
                    GENERAREST% = 1
                  End If
                Else
                  GENERAREST% = 1
                  CONDI$ = " ID_LISTA  > 0 "
                  IDLISTAX& = 1 + XVALO(VALOBADA("ENCAPRFUN", "MAX(ID_LISTA)", CONDI$, "NOMESS"))
                End If
                If GENERAREST% > 0 Then
                  SQLX$ = "INSERT INTO  ENCAPRFUN (ID_LISTA ,LISTA, Descripcion, ADDREGISTROS, USAFORMULA, BORRAR, INSERTAR, NOTILDE "
                  SQLX$ = SQLX$ + ", CANTICAMPOS, CodiEmpr, USARECORD, FECHEMISI, TITURECORD, ConsultaSql)"
                  SQLX$ = SQLX$ + "VALUES ( " & IDLISTAX&
                  SQLX$ = SQLX$ + ",'CTP.ETIQ'"
                  SQLX$ = SQLX$ + ",'ETIQ'"
                  SQLX$ = SQLX$ + ",1"
                  SQLX$ = SQLX$ + ",1"
                  SQLX$ = SQLX$ + " ,0"
                  SQLX$ = SQLX$ + " ,0"
                  SQLX$ = SQLX$ + " ,0"
                  SQLX$ = SQLX$ + " ,4"
                  SQLX$ = SQLX$ + " ,0"
                  SQLX$ = SQLX$ + " ,0"
                  SQLX$ = SQLX$ + " ,'03/10/19'"
                  SQLX$ = SQLX$ + " ,''"
                  SQLX$ = SQLX$ + " ,'SELECT CAMPO1,CAMPO2,CAMPO3 FROM DATASQL WITH(NOLOCK)'"
                  SQLX$ = SQLX$ + ")"
                  FLEXCONN.Execute (SQLX$)
                    SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
                    SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
                    SQLX$ = SQLX$ + " ,0"
                    SQLX$ = SQLX$ + " ,0"
                    SQLX$ = SQLX$ + " , 'ID/F8'"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " )"
                    FLEXCONN.Execute (SQLX$)
                
                '
                    SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
                    SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
                    SQLX$ = SQLX$ + " ,0"
                    SQLX$ = SQLX$ + " ,0"
                    SQLX$ = SQLX$ + " , 'ETIQ/A64'"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " , 'S'"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " )"
                    FLEXCONN.Execute (SQLX$)
                
                '
                    SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
                    SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
                    SQLX$ = SQLX$ + " ,0"
                    SQLX$ = SQLX$ + " ,0"
                    SQLX$ = SQLX$ + " , 'NV/A0'"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " )"
                    FLEXCONN.Execute (SQLX$)
                
                '
                End If
                '|FIN CTP.ETIQ
                
                  
                  '|INICIO CTP.PAB
                '
                CONDI$ = " LISTA =   'CTP.PAB'"
                EXILISTA% = CantRegistros&("ENCAPRFUN", CONDI$, "NOMESS")
                If EXILISTA% > 0 Then
                  FE% = 16547
                  FEBD% = CVINT(VALOBADA("ENCAPRFUN", "FECHEMISI", CONDI$, "NOMESS"))
                  GENERAREST% = 0
                  If FEBD% < FE% Then
                    IDLISTAX& = XVALO(VALOBADA("ENCAPRFUN", "ID_LISTA", CONDI$, "NOMESS"))
                    Call BORRAREGISTROS("ENCAPRFUN", "ID_LISTA = " & IDLISTAX&, REGAF&, "NOMESS")
                    Call BORRAREGISTROS("ENCAPRF", "ID_LISTA = " & IDLISTAX&, REGAF&, "NOMESS")
                    GENERAREST% = 1
                  End If
                Else
                  GENERAREST% = 1
                  CONDI$ = " ID_LISTA  > 0 "
                  IDLISTAX& = 1 + XVALO(VALOBADA("ENCAPRFUN", "MAX(ID_LISTA)", CONDI$, "NOMESS"))
                End If
                If GENERAREST% > 0 Then
                  SQLX$ = "INSERT INTO  ENCAPRFUN (ID_LISTA ,LISTA, Descripcion, ADDREGISTROS, USAFORMULA, BORRAR, INSERTAR, NOTILDE "
                  SQLX$ = SQLX$ + ", CANTICAMPOS, CodiEmpr, USARECORD, FECHEMISI, TITURECORD, ConsultaSql)"
                  SQLX$ = SQLX$ + "VALUES ( " & IDLISTAX&
                  SQLX$ = SQLX$ + ",'CTP.PAB'"
                  SQLX$ = SQLX$ + ",'PAB'"
                  SQLX$ = SQLX$ + ",1"
                  SQLX$ = SQLX$ + ",1"
                  SQLX$ = SQLX$ + " ,0"
                  SQLX$ = SQLX$ + " ,0"
                  SQLX$ = SQLX$ + " ,0"
                  SQLX$ = SQLX$ + " ,4"
                  SQLX$ = SQLX$ + " ,0"
                  SQLX$ = SQLX$ + " ,0"
                  SQLX$ = SQLX$ + " ,'03/10/19'"
                  SQLX$ = SQLX$ + " ,''"
                  SQLX$ = SQLX$ + " ,'SELECT CAMPO1,CAMPO2,CAMPO3 FROM DATASQL WITH(NOLOCK)'"
                  SQLX$ = SQLX$ + ")"
                  FLEXCONN.Execute (SQLX$)
                    SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
                    SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
                    SQLX$ = SQLX$ + " ,0"
                    SQLX$ = SQLX$ + " ,0"
                    SQLX$ = SQLX$ + " , 'ID/F8'"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " )"
                    FLEXCONN.Execute (SQLX$)
                
                '
                    SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
                    SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
                    SQLX$ = SQLX$ + " ,0"
                    SQLX$ = SQLX$ + " ,0"
                    SQLX$ = SQLX$ + " , 'PAB/A32'"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " , 'S'"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " )"
                    FLEXCONN.Execute (SQLX$)
                
                '
                    SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
                    SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
                    SQLX$ = SQLX$ + " ,0"
                    SQLX$ = SQLX$ + " ,0"
                    SQLX$ = SQLX$ + " , 'NV/A0'"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " )"
                    FLEXCONN.Execute (SQLX$)
                
                '
                End If
                '|FIN CTP.PAB

                '|INICIO CTP.FAR
                '
                CONDI$ = " LISTA =   'CTP.FAR'"
                EXILISTA% = CantRegistros&("ENCAPRFUN", CONDI$, "NOMESS")
                If EXILISTA% > 0 Then
                  FE% = 16547
                  FEBD% = CVINT(VALOBADA("ENCAPRFUN", "FECHEMISI", CONDI$, "NOMESS"))
                  GENERAREST% = 0
                  If FEBD% < FE% Then
                    IDLISTAX& = XVALO(VALOBADA("ENCAPRFUN", "ID_LISTA", CONDI$, "NOMESS"))
                    Call BORRAREGISTROS("ENCAPRFUN", "ID_LISTA = " & IDLISTAX&, REGAF&, "NOMESS")
                    Call BORRAREGISTROS("ENCAPRF", "ID_LISTA = " & IDLISTAX&, REGAF&, "NOMESS")
                    GENERAREST% = 1
                  End If
                Else
                  GENERAREST% = 1
                  CONDI$ = " ID_LISTA  > 0 "
                  IDLISTAX& = 1 + XVALO(VALOBADA("ENCAPRFUN", "MAX(ID_LISTA)", CONDI$, "NOMESS"))
                End If
                If GENERAREST% > 0 Then
                  SQLX$ = "INSERT INTO  ENCAPRFUN (ID_LISTA ,LISTA, Descripcion, ADDREGISTROS, USAFORMULA, BORRAR, INSERTAR, NOTILDE "
                  SQLX$ = SQLX$ + ", CANTICAMPOS, CodiEmpr, USARECORD, FECHEMISI, TITURECORD, ConsultaSql)"
                  SQLX$ = SQLX$ + "VALUES ( " & IDLISTAX&
                  SQLX$ = SQLX$ + ",'CTP.FAR'"
                  SQLX$ = SQLX$ + ",'FAR'"
                  SQLX$ = SQLX$ + ",1"
                  SQLX$ = SQLX$ + ",1"
                  SQLX$ = SQLX$ + " ,0"
                  SQLX$ = SQLX$ + " ,0"
                  SQLX$ = SQLX$ + " ,0"
                  SQLX$ = SQLX$ + " ,4"
                  SQLX$ = SQLX$ + " ,0"
                  SQLX$ = SQLX$ + " ,0"
                  SQLX$ = SQLX$ + " ,'03/10/19'"
                  SQLX$ = SQLX$ + " ,''"
                  SQLX$ = SQLX$ + " ,'SELECT CAMPO1,CAMPO2,CAMPO3 FROM DATASQL WITH(NOLOCK)'"
                  SQLX$ = SQLX$ + ")"
                  FLEXCONN.Execute (SQLX$)
                    SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
                    SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
                    SQLX$ = SQLX$ + " ,0"
                    SQLX$ = SQLX$ + " ,0"
                    SQLX$ = SQLX$ + " , 'ID/F8'"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " )"
                    FLEXCONN.Execute (SQLX$)
                
                '
                    SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
                    SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
                    SQLX$ = SQLX$ + " ,0"
                    SQLX$ = SQLX$ + " ,0"
                    SQLX$ = SQLX$ + " , 'FAR/A32'"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " , 'S'"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " )"
                    FLEXCONN.Execute (SQLX$)
                
                '
                    SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
                    SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
                    SQLX$ = SQLX$ + " ,0"
                    SQLX$ = SQLX$ + " ,0"
                    SQLX$ = SQLX$ + " , 'NV/A0'"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " )"
                    FLEXCONN.Execute (SQLX$)
                
                '
                End If
                '|FIN CTP.FAR
    
    
                '|INICIO CTP.FORMMAMPA
                '
                CONDI$ = " LISTA =   'CTP.FORMMAMPA'"
                EXILISTA% = CantRegistros&("ENCAPRFUN", CONDI$, "NOMESS")
                If EXILISTA% > 0 Then
                  FE% = 16547
                  FEBD% = CVINT(VALOBADA("ENCAPRFUN", "FECHEMISI", CONDI$, "NOMESS"))
                  GENERAREST% = 0
                  If FEBD% < FE% Then
                    IDLISTAX& = XVALO(VALOBADA("ENCAPRFUN", "ID_LISTA", CONDI$, "NOMESS"))
                    Call BORRAREGISTROS("ENCAPRFUN", "ID_LISTA = " & IDLISTAX&, REGAF&, "NOMESS")
                    Call BORRAREGISTROS("ENCAPRF", "ID_LISTA = " & IDLISTAX&, REGAF&, "NOMESS")
                    GENERAREST% = 1
                  End If
                Else
                  GENERAREST% = 1
                  CONDI$ = " ID_LISTA  > 0 "
                  IDLISTAX& = 1 + XVALO(VALOBADA("ENCAPRFUN", "MAX(ID_LISTA)", CONDI$, "NOMESS"))
                End If
                If GENERAREST% > 0 Then
                  SQLX$ = "INSERT INTO  ENCAPRFUN (ID_LISTA ,LISTA, Descripcion, ADDREGISTROS, USAFORMULA, BORRAR, INSERTAR, NOTILDE "
                  SQLX$ = SQLX$ + ", CANTICAMPOS, CodiEmpr, USARECORD, FECHEMISI, TITURECORD, ConsultaSql)"
                  SQLX$ = SQLX$ + "VALUES ( " & IDLISTAX&
                  SQLX$ = SQLX$ + ",'CTP.FORMMAMPA'"
                  SQLX$ = SQLX$ + ",'FORMULA PARA MAMPARO'"
                  SQLX$ = SQLX$ + ",1"
                  SQLX$ = SQLX$ + ",1"
                  SQLX$ = SQLX$ + " ,0"
                  SQLX$ = SQLX$ + " ,0"
                  SQLX$ = SQLX$ + " ,0"
                  SQLX$ = SQLX$ + " ,4"
                  SQLX$ = SQLX$ + " ,0"
                  SQLX$ = SQLX$ + " ,0"
                  SQLX$ = SQLX$ + " ,'03/10/19'"
                  SQLX$ = SQLX$ + " ,''"
                  SQLX$ = SQLX$ + " ,'SELECT CAMPO1,CAMPO2,CAMPO3 FROM DATASQL WITH(NOLOCK)'"
                  SQLX$ = SQLX$ + ")"
                  FLEXCONN.Execute (SQLX$)
                    SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
                    SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
                    SQLX$ = SQLX$ + " ,0"
                    SQLX$ = SQLX$ + " ,0"
                    SQLX$ = SQLX$ + " , 'ID/F8'"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " )"
                    FLEXCONN.Execute (SQLX$)
                
                '
                    SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
                    SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
                    SQLX$ = SQLX$ + " ,0"
                    SQLX$ = SQLX$ + " ,0"
                    SQLX$ = SQLX$ + " , 'F.MAMPARO/A48'"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " , 'S'"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " )"
                    FLEXCONN.Execute (SQLX$)
                
                '
                    SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
                    SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
                    SQLX$ = SQLX$ + " ,0"
                    SQLX$ = SQLX$ + " ,0"
                    SQLX$ = SQLX$ + " , 'NV/A0'"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " )"
                    FLEXCONN.Execute (SQLX$)
                
                '
                End If
                '|FIN CTP.FORMMAMPA
    
                '|INICIO CTP.PALLET
                '
                CONDI$ = " LISTA =   'CTP.PALLET'"
                EXILISTA% = CantRegistros&("ENCAPRFUN", CONDI$, "NOMESS")
                If EXILISTA% > 0 Then
                  FE% = 16547
                  FEBD% = CVINT(VALOBADA("ENCAPRFUN", "FECHEMISI", CONDI$, "NOMESS"))
                  GENERAREST% = 0
                  If FEBD% < FE% Then
                    IDLISTAX& = XVALO(VALOBADA("ENCAPRFUN", "ID_LISTA", CONDI$, "NOMESS"))
                    Call BORRAREGISTROS("ENCAPRFUN", "ID_LISTA = " & IDLISTAX&, REGAF&, "NOMESS")
                    Call BORRAREGISTROS("ENCAPRF", "ID_LISTA = " & IDLISTAX&, REGAF&, "NOMESS")
                    GENERAREST% = 1
                  End If
                Else
                  GENERAREST% = 1
                  CONDI$ = " ID_LISTA  > 0 "
                  IDLISTAX& = 1 + XVALO(VALOBADA("ENCAPRFUN", "MAX(ID_LISTA)", CONDI$, "NOMESS"))
                End If
                If GENERAREST% > 0 Then
                  SQLX$ = "INSERT INTO  ENCAPRFUN (ID_LISTA ,LISTA, Descripcion, ADDREGISTROS, USAFORMULA, BORRAR, INSERTAR, NOTILDE "
                  SQLX$ = SQLX$ + ", CANTICAMPOS, CodiEmpr, USARECORD, FECHEMISI, TITURECORD, ConsultaSql)"
                  SQLX$ = SQLX$ + "VALUES ( " & IDLISTAX&
                  SQLX$ = SQLX$ + ",'CTP.PALLET'"
                  SQLX$ = SQLX$ + ",'PALLET'"
                  SQLX$ = SQLX$ + ",1"
                  SQLX$ = SQLX$ + ",1"
                  SQLX$ = SQLX$ + " ,0"
                  SQLX$ = SQLX$ + " ,0"
                  SQLX$ = SQLX$ + " ,0"
                  SQLX$ = SQLX$ + " ,4"
                  SQLX$ = SQLX$ + " ,0"
                  SQLX$ = SQLX$ + " ,0"
                  SQLX$ = SQLX$ + " ,'03/10/19'"
                  SQLX$ = SQLX$ + " ,''"
                  SQLX$ = SQLX$ + " ,'SELECT CAMPO1,CAMPO2,CAMPO3 FROM DATASQL WITH(NOLOCK)'"
                  SQLX$ = SQLX$ + ")"
                  FLEXCONN.Execute (SQLX$)
                    SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
                    SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
                    SQLX$ = SQLX$ + " ,0"
                    SQLX$ = SQLX$ + " ,0"
                    SQLX$ = SQLX$ + " , 'ID/F8'"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " )"
                    FLEXCONN.Execute (SQLX$)
                
                '
                    SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
                    SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
                    SQLX$ = SQLX$ + " ,0"
                    SQLX$ = SQLX$ + " ,0"
                    SQLX$ = SQLX$ + " , 'PALLET/A48'"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " , 'S'"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " )"
                    FLEXCONN.Execute (SQLX$)
                
                '
                    SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
                    SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
                    SQLX$ = SQLX$ + " ,0"
                    SQLX$ = SQLX$ + " ,0"
                    SQLX$ = SQLX$ + " , 'NV/A0'"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " )"
                    FLEXCONN.Execute (SQLX$)
                
                '
                End If
                '|FIN CTP.PALLET
    
    
                '|INICIO CTP.CANTIUSOS
                '
                CONDI$ = " LISTA =   'CTP.CANTIUSOS'"
                EXILISTA% = CantRegistros&("ENCAPRFUN", CONDI$, "NOMESS")
                If EXILISTA% > 0 Then
                  FE% = 16548
                  FEBD% = CVINT(VALOBADA("ENCAPRFUN", "FECHEMISI", CONDI$, "NOMESS"))
                  GENERAREST% = 0
                  If FEBD% < FE% Then
                    IDLISTAX& = XVALO(VALOBADA("ENCAPRFUN", "ID_LISTA", CONDI$, "NOMESS"))
                    Call BORRAREGISTROS("ENCAPRFUN", "ID_LISTA = " & IDLISTAX&, REGAF&, "NOMESS")
                    Call BORRAREGISTROS("ENCAPRF", "ID_LISTA = " & IDLISTAX&, REGAF&, "NOMESS")
                    GENERAREST% = 1
                  End If
                Else
                  GENERAREST% = 1
                  CONDI$ = " ID_LISTA  > 0 "
                  IDLISTAX& = 1 + XVALO(VALOBADA("ENCAPRFUN", "MAX(ID_LISTA)", CONDI$, "NOMESS"))
                End If
                If GENERAREST% > 0 Then
                  SQLX$ = "INSERT INTO  ENCAPRFUN (ID_LISTA ,LISTA, Descripcion, ADDREGISTROS, USAFORMULA, BORRAR, INSERTAR, NOTILDE "
                  SQLX$ = SQLX$ + ", CANTICAMPOS, CodiEmpr, USARECORD, FECHEMISI, TITURECORD, ConsultaSql)"
                  SQLX$ = SQLX$ + "VALUES ( " & IDLISTAX&
                  SQLX$ = SQLX$ + ",'CTP.CANTIUSOS'"
                  SQLX$ = SQLX$ + ",'CANTIDAD DE USOS'"
                  SQLX$ = SQLX$ + ",1"
                  SQLX$ = SQLX$ + ",1"
                  SQLX$ = SQLX$ + " ,0"
                  SQLX$ = SQLX$ + " ,0"
                  SQLX$ = SQLX$ + " ,0"
                  SQLX$ = SQLX$ + " ,4"
                  SQLX$ = SQLX$ + " ,0"
                  SQLX$ = SQLX$ + " ,0"
                  SQLX$ = SQLX$ + " ,'04/10/19'"
                  SQLX$ = SQLX$ + " ,''"
                  SQLX$ = SQLX$ + " ,'SELECT CAMPO1,CAMPO2,CAMPO3 FROM DATASQL WITH(NOLOCK)'"
                  SQLX$ = SQLX$ + ")"
                  FLEXCONN.Execute (SQLX$)
                    SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
                    SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
                    SQLX$ = SQLX$ + " ,0"
                    SQLX$ = SQLX$ + " ,0"
                    SQLX$ = SQLX$ + " , 'ID/F8'"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " )"
                    FLEXCONN.Execute (SQLX$)
                
                '
                    SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
                    SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
                    SQLX$ = SQLX$ + " ,0"
                    SQLX$ = SQLX$ + " ,0"
                    SQLX$ = SQLX$ + " , 'CANTIDAD DE USOS/A48'"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " , 'S'"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " )"
                    FLEXCONN.Execute (SQLX$)
                
                '
                    SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
                    SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
                    SQLX$ = SQLX$ + " ,0"
                    SQLX$ = SQLX$ + " ,0"
                    SQLX$ = SQLX$ + " , 'NV/A0'"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " )"
                    FLEXCONN.Execute (SQLX$)
                
                '
                End If
                '|FIN CTP.CANTIUSOS
    
    
    
    
                 '|INICIO CTP.EMPRESA
                '
                CONDI$ = " LISTA =   'CTP.EMPRESA'"
                EXILISTA% = CantRegistros&("ENCAPRFUN", CONDI$, "NOMESS")
                If EXILISTA% > 0 Then
                  FE% = 16547
                  FEBD% = CVINT(VALOBADA("ENCAPRFUN", "FECHEMISI", CONDI$, "NOMESS"))
                  GENERAREST% = 0
                  If FEBD% < FE% Then
                    IDLISTAX& = XVALO(VALOBADA("ENCAPRFUN", "ID_LISTA", CONDI$, "NOMESS"))
                    Call BORRAREGISTROS("ENCAPRFUN", "ID_LISTA = " & IDLISTAX&, REGAF&, "NOMESS")
                    Call BORRAREGISTROS("ENCAPRF", "ID_LISTA = " & IDLISTAX&, REGAF&, "NOMESS")
                    GENERAREST% = 1
                  End If
                Else
                  GENERAREST% = 1
                  CONDI$ = " ID_LISTA  > 0 "
                  IDLISTAX& = 1 + XVALO(VALOBADA("ENCAPRFUN", "MAX(ID_LISTA)", CONDI$, "NOMESS"))
                End If
                If GENERAREST% > 0 Then
                  SQLX$ = "INSERT INTO  ENCAPRFUN (ID_LISTA ,LISTA, Descripcion, ADDREGISTROS, USAFORMULA, BORRAR, INSERTAR, NOTILDE "
                  SQLX$ = SQLX$ + ", CANTICAMPOS, CodiEmpr, USARECORD, FECHEMISI, TITURECORD, ConsultaSql)"
                  SQLX$ = SQLX$ + "VALUES ( " & IDLISTAX&
                  SQLX$ = SQLX$ + ",'CTP.EMPRESA'"
                  SQLX$ = SQLX$ + ",'EMPRESA'"
                  SQLX$ = SQLX$ + ",1"
                  SQLX$ = SQLX$ + ",1"
                  SQLX$ = SQLX$ + " ,0"
                  SQLX$ = SQLX$ + " ,0"
                  SQLX$ = SQLX$ + " ,0"
                  SQLX$ = SQLX$ + " ,4"
                  SQLX$ = SQLX$ + " ,0"
                  SQLX$ = SQLX$ + " ,0"
                  SQLX$ = SQLX$ + " ,'03/10/19'"
                  SQLX$ = SQLX$ + " ,''"
                  SQLX$ = SQLX$ + " ,'SELECT CAMPO1,CAMPO2,CAMPO3 FROM DATASQL WITH(NOLOCK)'"
                  SQLX$ = SQLX$ + ")"
                  FLEXCONN.Execute (SQLX$)
                    SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
                    SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
                    SQLX$ = SQLX$ + " ,0"
                    SQLX$ = SQLX$ + " ,0"
                    SQLX$ = SQLX$ + " , 'ID/F8'"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " )"
                    FLEXCONN.Execute (SQLX$)
                
                '
                    SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
                    SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
                    SQLX$ = SQLX$ + " ,0"
                    SQLX$ = SQLX$ + " ,0"
                    SQLX$ = SQLX$ + " , 'EMPRESA/A48'"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " , 'S'"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " )"
                    FLEXCONN.Execute (SQLX$)
                
                '
                    SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
                    SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
                    SQLX$ = SQLX$ + " ,0"
                    SQLX$ = SQLX$ + " ,0"
                    SQLX$ = SQLX$ + " , 'NV/A0'"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " )"
                    FLEXCONN.Execute (SQLX$)
                
                '
                End If
                '|FIN CTP.EMPRESA
   
                 '|INICIO CTP.ESL
                '
                CONDI$ = " LISTA =   'CTP.ESL'"
                EXILISTA% = CantRegistros&("ENCAPRFUN", CONDI$, "NOMESS")
                If EXILISTA% > 0 Then
                  FE% = 16547
                  FEBD% = CVINT(VALOBADA("ENCAPRFUN", "FECHEMISI", CONDI$, "NOMESS"))
                  GENERAREST% = 0
                  If FEBD% < FE% Then
                    IDLISTAX& = XVALO(VALOBADA("ENCAPRFUN", "ID_LISTA", CONDI$, "NOMESS"))
                    Call BORRAREGISTROS("ENCAPRFUN", "ID_LISTA = " & IDLISTAX&, REGAF&, "NOMESS")
                    Call BORRAREGISTROS("ENCAPRF", "ID_LISTA = " & IDLISTAX&, REGAF&, "NOMESS")
                    GENERAREST% = 1
                  End If
                Else
                  GENERAREST% = 1
                  CONDI$ = " ID_LISTA  > 0 "
                  IDLISTAX& = 1 + XVALO(VALOBADA("ENCAPRFUN", "MAX(ID_LISTA)", CONDI$, "NOMESS"))
                End If
                If GENERAREST% > 0 Then
                  SQLX$ = "INSERT INTO  ENCAPRFUN (ID_LISTA ,LISTA, Descripcion, ADDREGISTROS, USAFORMULA, BORRAR, INSERTAR, NOTILDE "
                  SQLX$ = SQLX$ + ", CANTICAMPOS, CodiEmpr, USARECORD, FECHEMISI, TITURECORD, ConsultaSql)"
                  SQLX$ = SQLX$ + "VALUES ( " & IDLISTAX&
                  SQLX$ = SQLX$ + ",'CTP.ESL'"
                  SQLX$ = SQLX$ + ",'ESL'"
                  SQLX$ = SQLX$ + ",1"
                  SQLX$ = SQLX$ + ",1"
                  SQLX$ = SQLX$ + " ,0"
                  SQLX$ = SQLX$ + " ,0"
                  SQLX$ = SQLX$ + " ,0"
                  SQLX$ = SQLX$ + " ,4"
                  SQLX$ = SQLX$ + " ,0"
                  SQLX$ = SQLX$ + " ,0"
                  SQLX$ = SQLX$ + " ,'03/10/19'"
                  SQLX$ = SQLX$ + " ,''"
                  SQLX$ = SQLX$ + " ,'SELECT CAMPO1,CAMPO2,CAMPO3 FROM DATASQL WITH(NOLOCK)'"
                  SQLX$ = SQLX$ + ")"
                  FLEXCONN.Execute (SQLX$)
                    SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
                    SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
                    SQLX$ = SQLX$ + " ,0"
                    SQLX$ = SQLX$ + " ,0"
                    SQLX$ = SQLX$ + " , 'ID/F8'"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " )"
                    FLEXCONN.Execute (SQLX$)
                
                '
                    SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
                    SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
                    SQLX$ = SQLX$ + " ,0"
                    SQLX$ = SQLX$ + " ,0"
                    SQLX$ = SQLX$ + " , 'ESL/F10'"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " , 'S'"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " )"
                    FLEXCONN.Execute (SQLX$)
                
                '
                    SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
                    SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
                    SQLX$ = SQLX$ + " ,0"
                    SQLX$ = SQLX$ + " ,0"
                    SQLX$ = SQLX$ + " , 'NV/A0'"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " )"
                    FLEXCONN.Execute (SQLX$)
                
                '
                End If
                '|FIN CTP.ESL
   
    
                '|INICIO CTP.LI
                '
                CONDI$ = " LISTA =   'CTP.LI'"
                EXILISTA% = CantRegistros&("ENCAPRFUN", CONDI$, "NOMESS")
                If EXILISTA% > 0 Then
                  FE% = 16547
                  FEBD% = CVINT(VALOBADA("ENCAPRFUN", "FECHEMISI", CONDI$, "NOMESS"))
                  GENERAREST% = 0
                  If FEBD% < FE% Then
                    IDLISTAX& = XVALO(VALOBADA("ENCAPRFUN", "ID_LISTA", CONDI$, "NOMESS"))
                    Call BORRAREGISTROS("ENCAPRFUN", "ID_LISTA = " & IDLISTAX&, REGAF&, "NOMESS")
                    Call BORRAREGISTROS("ENCAPRF", "ID_LISTA = " & IDLISTAX&, REGAF&, "NOMESS")
                    GENERAREST% = 1
                  End If
                Else
                  GENERAREST% = 1
                  CONDI$ = " ID_LISTA  > 0 "
                  IDLISTAX& = 1 + XVALO(VALOBADA("ENCAPRFUN", "MAX(ID_LISTA)", CONDI$, "NOMESS"))
                End If
                If GENERAREST% > 0 Then
                  SQLX$ = "INSERT INTO  ENCAPRFUN (ID_LISTA ,LISTA, Descripcion, ADDREGISTROS, USAFORMULA, BORRAR, INSERTAR, NOTILDE "
                  SQLX$ = SQLX$ + ", CANTICAMPOS, CodiEmpr, USARECORD, FECHEMISI, TITURECORD, ConsultaSql)"
                  SQLX$ = SQLX$ + "VALUES ( " & IDLISTAX&
                  SQLX$ = SQLX$ + ",'CTP.LI'"
                  SQLX$ = SQLX$ + ",'LI'"
                  SQLX$ = SQLX$ + ",1"
                  SQLX$ = SQLX$ + ",1"
                  SQLX$ = SQLX$ + " ,0"
                  SQLX$ = SQLX$ + " ,0"
                  SQLX$ = SQLX$ + " ,0"
                  SQLX$ = SQLX$ + " ,4"
                  SQLX$ = SQLX$ + " ,0"
                  SQLX$ = SQLX$ + " ,0"
                  SQLX$ = SQLX$ + " ,'03/10/19'"
                  SQLX$ = SQLX$ + " ,''"
                  SQLX$ = SQLX$ + " ,'SELECT CAMPO1,CAMPO2,CAMPO3 FROM DATASQL WITH(NOLOCK)'"
                  SQLX$ = SQLX$ + ")"
                  FLEXCONN.Execute (SQLX$)
                    SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
                    SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
                    SQLX$ = SQLX$ + " ,0"
                    SQLX$ = SQLX$ + " ,0"
                    SQLX$ = SQLX$ + " , 'ID/F8'"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " )"
                    FLEXCONN.Execute (SQLX$)
                
                '
                    SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
                    SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
                    SQLX$ = SQLX$ + " ,0"
                    SQLX$ = SQLX$ + " ,0"
                    SQLX$ = SQLX$ + " , 'LI/A32'"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " , 'S'"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " )"
                    FLEXCONN.Execute (SQLX$)
                
                '
                    SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
                    SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
                    SQLX$ = SQLX$ + " ,0"
                    SQLX$ = SQLX$ + " ,0"
                    SQLX$ = SQLX$ + " , 'NV/A0'"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " )"
                    FLEXCONN.Execute (SQLX$)
                
                '
                End If
                '|FIN CTP.LI
    
    End If
    
    
    
    
    
    
    
    
    
    '|INICIO ARMAINFAGRU
    '
    CONDI$ = " LISTA =   'ARMAINFAGRU'"
    EXILISTA% = CantRegistros&("ENCAPRFUN", CONDI$, "NOMESS")
    If EXILISTA% > 0 Then
      FE% = 16465
      FEBD% = CVINT(VALOBADA("ENCAPRFUN", "FECHEMISI", CONDI$, "NOMESS"))
      GENERAREST% = 0
      If FEBD% < FE% Then
        IDLISTAX& = XVALO(VALOBADA("ENCAPRFUN", "ID_LISTA", CONDI$, "NOMESS"))
        Call BORRAREGISTROS("ENCAPRFUN", "ID_LISTA = " & IDLISTAX&, REGAF&, "NOMESS")
        Call BORRAREGISTROS("ENCAPRF", "ID_LISTA = " & IDLISTAX&, REGAF&, "NOMESS")
        GENERAREST% = 1
      End If
    Else
      GENERAREST% = 1
      CONDI$ = " ID_LISTA  > 0 "
      IDLISTAX& = 1 + XVALO(VALOBADA("ENCAPRFUN", "MAX(ID_LISTA)", CONDI$, "NOMESS"))
    End If
    If GENERAREST% > 0 Then
      SQLX$ = "INSERT INTO  ENCAPRFUN (ID_LISTA ,LISTA, Descripcion, ADDREGISTROS, USAFORMULA, BORRAR, INSERTAR, NOTILDE "
      SQLX$ = SQLX$ + ", CANTICAMPOS, CodiEmpr, USARECORD, FECHEMISI, TITURECORD, ConsultaSql)"
      SQLX$ = SQLX$ + "VALUES ( " & IDLISTAX&
      SQLX$ = SQLX$ + ",'ARMAINFAGRU'"
      SQLX$ = SQLX$ + ",'ARMADO DE INFORMES AGRUPADOS'"
      SQLX$ = SQLX$ + ",1"
      SQLX$ = SQLX$ + ",1"
      SQLX$ = SQLX$ + " ,1"
      SQLX$ = SQLX$ + " ,1"
      SQLX$ = SQLX$ + " ,0"
      SQLX$ = SQLX$ + " ,6"
      SQLX$ = SQLX$ + " ,0"
      SQLX$ = SQLX$ + " ,0"
      SQLX$ = SQLX$ + " ,'17/07/2019'"
      SQLX$ = SQLX$ + " ,''"
      SQLX$ = SQLX$ + " ,'SELECT CAMPO1, CAMPO2, CAMPO3, CAMPO4, CAMPO5, CAMPO6, CAMPO7 FROM DATASQL WITH(NOLOCK)'"
      SQLX$ = SQLX$ + ")"
      FLEXCONN.Execute (SQLX$)
      
        SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
        SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
        SQLX$ = SQLX$ + " ,0"
        SQLX$ = SQLX$ + " ,0"
        SQLX$ = SQLX$ + " , 'Ord/F3.0'"
        SQLX$ = SQLX$ + " , ''"
        SQLX$ = SQLX$ + " , 'N'"
        SQLX$ = SQLX$ + " , ''"
        SQLX$ = SQLX$ + " , ''"
        SQLX$ = SQLX$ + " , ''"
        SQLX$ = SQLX$ + " , ''"
        SQLX$ = SQLX$ + " )"
        FLEXCONN.Execute (SQLX$)
    
    '
        SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
        SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
        SQLX$ = SQLX$ + " ,0"
        SQLX$ = SQLX$ + " ,0"
        SQLX$ = SQLX$ + " , 'Inf/F3.0'"
        SQLX$ = SQLX$ + " , ''"
        SQLX$ = SQLX$ + " , 'S'"
        SQLX$ = SQLX$ + " , " & "'CAMPO1;CAMPO2;DATASQL;LISTA ='" & "'INFOCONAGRU'" & "''"
        SQLX$ = SQLX$ + " , ''"
        SQLX$ = SQLX$ + " , ''"
        SQLX$ = SQLX$ + " , ''"
        SQLX$ = SQLX$ + " )"
        FLEXCONN.Execute (SQLX$)
    
    '
        SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
        SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
        SQLX$ = SQLX$ + " ,0"
        SQLX$ = SQLX$ + " ,0"
        SQLX$ = SQLX$ + " , 'Nombre/A24'"
        SQLX$ = SQLX$ + " , ''"
        SQLX$ = SQLX$ + " , 'N'"
        SQLX$ = SQLX$ + " , ''"
        SQLX$ = SQLX$ + " , ''"
        SQLX$ = SQLX$ + " , ''"
        SQLX$ = SQLX$ + " , ''"
        SQLX$ = SQLX$ + " )"
        FLEXCONN.Execute (SQLX$)
    
    '
        SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
        SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
        SQLX$ = SQLX$ + " ,0"
        SQLX$ = SQLX$ + " ,0"
        SQLX$ = SQLX$ + " , 'Nivel/F3.0'"
        SQLX$ = SQLX$ + " , ''"
        SQLX$ = SQLX$ + " , 'S'"
        SQLX$ = SQLX$ + " , ''"
        SQLX$ = SQLX$ + " , ''"
        SQLX$ = SQLX$ + " , ''"
        SQLX$ = SQLX$ + " , ''"
        SQLX$ = SQLX$ + " )"
        FLEXCONN.Execute (SQLX$)
    
    '
        SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
        SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
        SQLX$ = SQLX$ + " ,0"
        SQLX$ = SQLX$ + " ,0"
        SQLX$ = SQLX$ + " , 'Descripción/A64'"
        SQLX$ = SQLX$ + " , ''"
        SQLX$ = SQLX$ + " , 'S'"
        SQLX$ = SQLX$ + " , ''"
        SQLX$ = SQLX$ + " , ''"
        SQLX$ = SQLX$ + " , ''"
        SQLX$ = SQLX$ + " , ''"
        SQLX$ = SQLX$ + " )"
        FLEXCONN.Execute (SQLX$)
    
    '
        SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
        SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
        SQLX$ = SQLX$ + " ,0"
        SQLX$ = SQLX$ + " ,0"
        SQLX$ = SQLX$ + " , 'Cuenta/A16'"
        SQLX$ = SQLX$ + " , ' '"
        SQLX$ = SQLX$ + " , 'S'"
        SQLX$ = SQLX$ + " , 'codexcue;denocue;plancuen'"
        SQLX$ = SQLX$ + " , ' '"
        SQLX$ = SQLX$ + " , ''"
        SQLX$ = SQLX$ + " , ''"
        SQLX$ = SQLX$ + " )"
        FLEXCONN.Execute (SQLX$)
    
    '
        SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
        SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
        SQLX$ = SQLX$ + " ,0"
        SQLX$ = SQLX$ + " ,0"
        SQLX$ = SQLX$ + " , 'Denominacion/A48'"
        SQLX$ = SQLX$ + " , ' '"
        SQLX$ = SQLX$ + " , 'N'"
        SQLX$ = SQLX$ + " , ' '"
        SQLX$ = SQLX$ + " , ''"
        SQLX$ = SQLX$ + " , ''"
        SQLX$ = SQLX$ + " , ''"
        SQLX$ = SQLX$ + " )"
        FLEXCONN.Execute (SQLX$)
    
    '
    End If
    '|FIN ARMAINFAGRU
    
    '|INICIO INFOCONAGRU
    '
    CONDI$ = " LISTA =   'INFOCONAGRU'"
    EXILISTA% = CantRegistros&("ENCAPRFUN", CONDI$, "NOMESS")
    If EXILISTA% > 0 Then
      FE% = 16465
      FEBD% = CVINT(VALOBADA("ENCAPRFUN", "FECHEMISI", CONDI$, "NOMESS"))
      GENERAREST% = 0
      If FEBD% < FE% Then
        IDLISTAX& = XVALO(VALOBADA("ENCAPRFUN", "ID_LISTA", CONDI$, "NOMESS"))
        Call BORRAREGISTROS("ENCAPRFUN", "ID_LISTA = " & IDLISTAX&, REGAF&, "NOMESS")
        Call BORRAREGISTROS("ENCAPRF", "ID_LISTA = " & IDLISTAX&, REGAF&, "NOMESS")
        GENERAREST% = 1
      End If
    Else
      GENERAREST% = 1
      CONDI$ = " ID_LISTA  > 0 "
      IDLISTAX& = 1 + XVALO(VALOBADA("ENCAPRFUN", "MAX(ID_LISTA)", CONDI$, "NOMESS"))
    End If
    If GENERAREST% > 0 Then
      SQLX$ = "INSERT INTO  ENCAPRFUN (ID_LISTA ,LISTA, Descripcion, ADDREGISTROS, USAFORMULA, BORRAR, INSERTAR, NOTILDE "
      SQLX$ = SQLX$ + ", CANTICAMPOS, CodiEmpr, USARECORD, FECHEMISI, TITURECORD, ConsultaSql)"
      SQLX$ = SQLX$ + "VALUES ( " & IDLISTAX&
      SQLX$ = SQLX$ + ",'INFOCONAGRU'"
      SQLX$ = SQLX$ + ",'INFORMES CONTABLES AGRUPADOS'"
      SQLX$ = SQLX$ + ",1"
      SQLX$ = SQLX$ + ",0"
      SQLX$ = SQLX$ + " ,1"
      SQLX$ = SQLX$ + " ,0"
      SQLX$ = SQLX$ + " ,0"
      SQLX$ = SQLX$ + " ,3"
      SQLX$ = SQLX$ + " ,0"
      SQLX$ = SQLX$ + " ,0"
      SQLX$ = SQLX$ + " ,'17/07/2019'"
      SQLX$ = SQLX$ + " ,''"
      SQLX$ = SQLX$ + " ,'select campo1, campo2 from datasql with(nolock)'"
      SQLX$ = SQLX$ + ")"
      FLEXCONN.Execute (SQLX$)
        SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
        SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
        SQLX$ = SQLX$ + " ,0"
        SQLX$ = SQLX$ + " ,0"
        SQLX$ = SQLX$ + " , 'Codigo/F3.0'"
        SQLX$ = SQLX$ + " , ' '"
        SQLX$ = SQLX$ + " , 'N'"
        SQLX$ = SQLX$ + " , ' '"
        SQLX$ = SQLX$ + " , ''"
        SQLX$ = SQLX$ + " , ''"
        SQLX$ = SQLX$ + " , ''"
        SQLX$ = SQLX$ + " )"
        FLEXCONN.Execute (SQLX$)
    
    '
        SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
        SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
        SQLX$ = SQLX$ + " ,0"
        SQLX$ = SQLX$ + " ,0"
        SQLX$ = SQLX$ + " , 'Nombre/a64'"
        SQLX$ = SQLX$ + " , ''"
        SQLX$ = SQLX$ + " , 's'"
        SQLX$ = SQLX$ + " , ''"
        SQLX$ = SQLX$ + " , ''"
        SQLX$ = SQLX$ + " , ''"
        SQLX$ = SQLX$ + " , ''"
        SQLX$ = SQLX$ + " )"
        FLEXCONN.Execute (SQLX$)
    
    '
    End If
    '|FIN INFOCONAGRU
    
    '
    '|INICIO DOCEXCONFIN
    '
    CONDI$ = " LISTA =   'DOCEXCONFIN'"
    EXILISTA% = CantRegistros&("ENCAPRFUN", CONDI$, "NOMESS")
    If EXILISTA% > 0 Then
      FE% = 16444
      FEBD% = CVINT(VALOBADA("ENCAPRFUN", "FECHEMISI", CONDI$, "NOMESS"))
      GENERAREST% = 0
      If FEBD% < FE% Then
        IDLISTAX& = XVALO(VALOBADA("ENCAPRFUN", "ID_LISTA", CONDI$, "NOMESS"))
        Call BORRAREGISTROS("ENCAPRFUN", "ID_LISTA = " & IDLISTAX&, REGAF&, "NOMESS")
        Call BORRAREGISTROS("ENCAPRF", "ID_LISTA = " & IDLISTAX&, REGAF&, "NOMESS")
        GENERAREST% = 1
      End If
    Else
      GENERAREST% = 1
      CONDI$ = " ID_LISTA  > 0 "
      IDLISTAX& = 1 + XVALO(VALOBADA("ENCAPRFUN", "MAX(ID_LISTA)", CONDI$, "NOMESS"))
    End If
    If GENERAREST% > 0 Then
      SQLX$ = "INSERT INTO  ENCAPRFUN (ID_LISTA ,LISTA, Descripcion, ADDREGISTROS, USAFORMULA, BORRAR, INSERTAR, NOTILDE "
      SQLX$ = SQLX$ + ", CANTICAMPOS, CodiEmpr, USARECORD, FECHEMISI, TITURECORD, ConsultaSql)"
      SQLX$ = SQLX$ + "VALUES ( " & IDLISTAX&
      SQLX$ = SQLX$ + ",'DOCEXCONFIN'"
      SQLX$ = SQLX$ + ",'DOCUMENTOS EXTERNOS PARA CONTROL FINAL'"
      SQLX$ = SQLX$ + ",1"
      SQLX$ = SQLX$ + ",1"
      SQLX$ = SQLX$ + " ,1"
      SQLX$ = SQLX$ + " ,1"
      SQLX$ = SQLX$ + " ,0"
      SQLX$ = SQLX$ + " ,6"
      SQLX$ = SQLX$ + " ,0"
      SQLX$ = SQLX$ + " ,0"
      SQLX$ = SQLX$ + " ,'28/06/2019'"
      SQLX$ = SQLX$ + " ,''"
      SQLX$ = SQLX$ + " ,'SELECT CAMPO1, CAMPO2, CAMPO3, CAMPO4, CAMPO5 FROM DATASQL WITH(NOLOCK)'"
      SQLX$ = SQLX$ + ")"
      FLEXCONN.Execute (SQLX$)
        SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
        SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
        SQLX$ = SQLX$ + " ,0"
        SQLX$ = SQLX$ + " ,0"
        SQLX$ = SQLX$ + " , 'Id/F0'"
        SQLX$ = SQLX$ + " , ''"
        SQLX$ = SQLX$ + " , 'N'"
        SQLX$ = SQLX$ + " , ''"
        SQLX$ = SQLX$ + " , ''"
        SQLX$ = SQLX$ + " , ''"
        SQLX$ = SQLX$ + " , ''"
        SQLX$ = SQLX$ + " )"
        FLEXCONN.Execute (SQLX$)
    
    '
        SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
        SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
        SQLX$ = SQLX$ + " ,0"
        SQLX$ = SQLX$ + " ,0"
        SQLX$ = SQLX$ + " , 'Codigo/A16'"
        SQLX$ = SQLX$ + " , ''"
        SQLX$ = SQLX$ + " , 'S'"
        SQLX$ = SQLX$ + " , 'codigo;descripcion;master'"
        SQLX$ = SQLX$ + " , ''"
        SQLX$ = SQLX$ + " , ''"
        SQLX$ = SQLX$ + " , ''"
        SQLX$ = SQLX$ + " )"
        FLEXCONN.Execute (SQLX$)
    
    '
        SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
        SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
        SQLX$ = SQLX$ + " ,0"
        SQLX$ = SQLX$ + " ,0"
        SQLX$ = SQLX$ + " , 'Descripción/A48'"
        SQLX$ = SQLX$ + " , ''"
        SQLX$ = SQLX$ + " , 'N'"
        SQLX$ = SQLX$ + " , ''"
        SQLX$ = SQLX$ + " , ''"
        SQLX$ = SQLX$ + " , ''"
        SQLX$ = SQLX$ + " , ''"
        SQLX$ = SQLX$ + " )"
        FLEXCONN.Execute (SQLX$)
    
    '
        SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
        SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
        SQLX$ = SQLX$ + " ,0"
        SQLX$ = SQLX$ + " ,0"
        SQLX$ = SQLX$ + " , 'Nombre/A48'"
        SQLX$ = SQLX$ + " , ''"
        SQLX$ = SQLX$ + " , 'S'"
        SQLX$ = SQLX$ + " , ''"
        SQLX$ = SQLX$ + " , ''"
        SQLX$ = SQLX$ + " , ''"
        SQLX$ = SQLX$ + " , ''"
        SQLX$ = SQLX$ + " )"
        FLEXCONN.Execute (SQLX$)
    
    '
        SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
        SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
        SQLX$ = SQLX$ + " ,0"
        SQLX$ = SQLX$ + " ,0"
        SQLX$ = SQLX$ + " , 'Ruta/A64'"
        SQLX$ = SQLX$ + " , ''"
        SQLX$ = SQLX$ + " , 'S'"
        SQLX$ = SQLX$ + " , ''"
        SQLX$ = SQLX$ + " , ''"
        SQLX$ = SQLX$ + " , ''"
        SQLX$ = SQLX$ + " , ''"
        SQLX$ = SQLX$ + " )"
        FLEXCONN.Execute (SQLX$)
    
    '
    End If
    '|FIN DOCEXCONFIN
    
    
    '
    '|INICIO CANCONART
    '
    CONDI$ = " LISTA =   'CANCONART'"
    EXILISTA% = CantRegistros&("ENCAPRFUN", CONDI$, "NOMESS")
    If EXILISTA% > 0 Then
      FE% = 16442
      FEBD% = CVINT(VALOBADA("ENCAPRFUN", "FECHEMISI", CONDI$, "NOMESS"))
      GENERAREST% = 0
      If FEBD% < FE% Then
        IDLISTAX& = XVALO(VALOBADA("ENCAPRFUN", "ID_LISTA", CONDI$, "NOMESS"))
        Call BORRAREGISTROS("ENCAPRFUN", "ID_LISTA = " & IDLISTAX&, REGAF&, "NOMESS")
        Call BORRAREGISTROS("ENCAPRF", "ID_LISTA = " & IDLISTAX&, REGAF&, "NOMESS")
        GENERAREST% = 1
      End If
    Else
      GENERAREST% = 1
      CONDI$ = " ID_LISTA  > 0 "
      IDLISTAX& = 1 + XVALO(VALOBADA("ENCAPRFUN", "MAX(ID_LISTA)", CONDI$, "NOMESS"))
    End If
    If GENERAREST% > 0 Then
      SQLX$ = "INSERT INTO  ENCAPRFUN (ID_LISTA ,LISTA, Descripcion, ADDREGISTROS, USAFORMULA, BORRAR, INSERTAR, NOTILDE "
      SQLX$ = SQLX$ + ", CANTICAMPOS, CodiEmpr, USARECORD, FECHEMISI, TITURECORD, ConsultaSql)"
      SQLX$ = SQLX$ + "VALUES ( " & IDLISTAX&
      SQLX$ = SQLX$ + ",'CANCONART'"
      SQLX$ = SQLX$ + ",'CANTIDAD DE CONTROLES POR ARTICULO'"
      SQLX$ = SQLX$ + ",1"
      SQLX$ = SQLX$ + ",1"
      SQLX$ = SQLX$ + " ,1"
      SQLX$ = SQLX$ + " ,1"
      SQLX$ = SQLX$ + " ,0"
      SQLX$ = SQLX$ + " ,4"
      SQLX$ = SQLX$ + " ,0"
      SQLX$ = SQLX$ + " ,0"
      SQLX$ = SQLX$ + " ,'26/06/2019'"
      SQLX$ = SQLX$ + " ,''"
      SQLX$ = SQLX$ + " ,'SELECT CAMPO1, CAMPO2, CAMPO3 FROM DATASQL WITH(NOLOCK)'"
      SQLX$ = SQLX$ + ")"
      FLEXCONN.Execute (SQLX$)
        SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
        SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
        SQLX$ = SQLX$ + " ,0"
        SQLX$ = SQLX$ + " ,0"
        SQLX$ = SQLX$ + " , 'Código/A16'"
        SQLX$ = SQLX$ + " , ''"
        SQLX$ = SQLX$ + " , 's'"
        SQLX$ = SQLX$ + " , 'codigo;descripcion;master'"
        SQLX$ = SQLX$ + " , ''"
        SQLX$ = SQLX$ + " , ''"
        SQLX$ = SQLX$ + " , ''"
        SQLX$ = SQLX$ + " )"
        FLEXCONN.Execute (SQLX$)
    
    '
        SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
        SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
        SQLX$ = SQLX$ + " ,0"
        SQLX$ = SQLX$ + " ,0"
        SQLX$ = SQLX$ + " , 'Descripción/A64'"
        SQLX$ = SQLX$ + " , ''"
        SQLX$ = SQLX$ + " , 'n'"
        SQLX$ = SQLX$ + " , ' '"
        SQLX$ = SQLX$ + " , ''"
        SQLX$ = SQLX$ + " , ''"
        SQLX$ = SQLX$ + " , ''"
        SQLX$ = SQLX$ + " )"
        FLEXCONN.Execute (SQLX$)
    
    '
        SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
        SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
        SQLX$ = SQLX$ + " ,0"
        SQLX$ = SQLX$ + " ,0"
        SQLX$ = SQLX$ + " , 'Cant. Ok/F8.0'"
        SQLX$ = SQLX$ + " , ''"
        SQLX$ = SQLX$ + " , 's'"
        SQLX$ = SQLX$ + " , ''"
        SQLX$ = SQLX$ + " , ''"
        SQLX$ = SQLX$ + " , ''"
        SQLX$ = SQLX$ + " , ''"
        SQLX$ = SQLX$ + " )"
        FLEXCONN.Execute (SQLX$)
    
    '
    End If
    '|FIN CANCONART
    
    '|INICIO LISCOSREPA
    '
    CONDI$ = " LISTA =   'LISCOSREPA'"
    EXILISTA% = CantRegistros&("ENCAPRFUN", CONDI$, "NOMESS")
    If EXILISTA% > 0 Then
      FE% = 16441
      FEBD% = CVINT(VALOBADA("ENCAPRFUN", "FECHEMISI", CONDI$, "NOMESS"))
      GENERAREST% = 0
      If FEBD% < FE% Then
        IDLISTAX& = XVALO(VALOBADA("ENCAPRFUN", "ID_LISTA", CONDI$, "NOMESS"))
        Call BORRAREGISTROS("ENCAPRFUN", "ID_LISTA = " & IDLISTAX&, REGAF&, "NOMESS")
        Call BORRAREGISTROS("ENCAPRF", "ID_LISTA = " & IDLISTAX&, REGAF&, "NOMESS")
        GENERAREST% = 1
      End If
    Else
      GENERAREST% = 1
      CONDI$ = " ID_LISTA  > 0 "
      IDLISTAX& = 1 + XVALO(VALOBADA("ENCAPRFUN", "MAX(ID_LISTA)", CONDI$, "NOMESS"))
    End If
    If GENERAREST% > 0 Then
      SQLX$ = "INSERT INTO  ENCAPRFUN (ID_LISTA ,LISTA, Descripcion, ADDREGISTROS, USAFORMULA, BORRAR, INSERTAR, NOTILDE "
      SQLX$ = SQLX$ + ", CANTICAMPOS, CodiEmpr, USARECORD, FECHEMISI, TITURECORD, ConsultaSql)"
      SQLX$ = SQLX$ + "VALUES ( " & IDLISTAX&
      SQLX$ = SQLX$ + ",'LISCOSREPA'"
      SQLX$ = SQLX$ + ",'LISTA DE COSTOS DE REPARACION'"
      SQLX$ = SQLX$ + ",1"
      SQLX$ = SQLX$ + ",1"
      SQLX$ = SQLX$ + " ,1"
      SQLX$ = SQLX$ + " ,0"
      SQLX$ = SQLX$ + " ,0"
      SQLX$ = SQLX$ + " ,4"
      SQLX$ = SQLX$ + " ,0"
      SQLX$ = SQLX$ + " ,0"
      SQLX$ = SQLX$ + " ,'25/06/2019'"
      SQLX$ = SQLX$ + " ,''"
      SQLX$ = SQLX$ + " ,'SELECT CAMPO1, CAMPO2, CAMPO3 FROM DATASQL WITH(NOLOCK)'"
      SQLX$ = SQLX$ + ")"
      FLEXCONN.Execute (SQLX$)
        SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
        SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
        SQLX$ = SQLX$ + " ,0"
        SQLX$ = SQLX$ + " ,0"
        SQLX$ = SQLX$ + " , 'Código/a16'"
        SQLX$ = SQLX$ + " , ''"
        SQLX$ = SQLX$ + " , 's'"
        SQLX$ = SQLX$ + " , 'codigo;descripcion;master'"
        SQLX$ = SQLX$ + " , ''"
        SQLX$ = SQLX$ + " , ''"
        SQLX$ = SQLX$ + " , ''"
        SQLX$ = SQLX$ + " )"
        FLEXCONN.Execute (SQLX$)
    
    '
        SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
        SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
        SQLX$ = SQLX$ + " ,0"
        SQLX$ = SQLX$ + " ,0"
        SQLX$ = SQLX$ + " , 'Descripción/a32'"
        SQLX$ = SQLX$ + " , ''"
        SQLX$ = SQLX$ + " , 'n'"
        SQLX$ = SQLX$ + " , ''"
        SQLX$ = SQLX$ + " , ''"
        SQLX$ = SQLX$ + " , ''"
        SQLX$ = SQLX$ + " , ''"
        SQLX$ = SQLX$ + " )"
        FLEXCONN.Execute (SQLX$)
    
    '
        SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
        SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
        SQLX$ = SQLX$ + " ,0"
        SQLX$ = SQLX$ + " ,0"
        SQLX$ = SQLX$ + " , 'Costo/F10.2'"
        SQLX$ = SQLX$ + " , ''"
        SQLX$ = SQLX$ + " , 's'"
        SQLX$ = SQLX$ + " , ' '"
        SQLX$ = SQLX$ + " , ''"
        SQLX$ = SQLX$ + " , ''"
        SQLX$ = SQLX$ + " , ''"
        SQLX$ = SQLX$ + " )"
        FLEXCONN.Execute (SQLX$)
    
    '
    End If
    '|FIN LISCOSREPA
    
    
'     *** PUNTOS DE VENTA DE EXPORTA SIMPLE
'    |INICIO PTOVTAEXSIM
    
    CONDI$ = " LISTA =   'PTOVTAEXSIM'"
    EXILISTA% = CantRegistros&("ENCAPRFUN", CONDI$, "NOMESS")
    If EXILISTA% > 0 Then
      FE% = 16399
      FEBD% = CVINT(VALOBADA("ENCAPRFUN", "FECHEMISI", CONDI$, "NOMESS"))
      GENERAREST% = 0
      If FEBD% < FE% Then
        IDLISTAX& = XVALO(VALOBADA("ENCAPRFUN", "ID_LISTA", CONDI$, "NOMESS"))
        Call BORRAREGISTROS("ENCAPRFUN", "ID_LISTA = " & IDLISTAX&, REGAF&, "NOMESS")
        Call BORRAREGISTROS("ENCAPRF", "ID_LISTA = " & IDLISTAX&, REGAF&, "NOMESS")
        GENERAREST% = 1
      End If
    Else
      GENERAREST% = 1
      CONDI$ = " ID_LISTA  > 0 "
      IDLISTAX& = 1 + XVALO(VALOBADA("ENCAPRFUN", "MAX(ID_LISTA)", CONDI$, "NOMESS"))
    End If
    If GENERAREST% > 0 Then
      SQLX$ = "INSERT INTO  ENCAPRFUN (ID_LISTA ,LISTA, Descripcion, ADDREGISTROS, USAFORMULA, BORRAR, INSERTAR, NOTILDE "
      SQLX$ = SQLX$ + ", CANTICAMPOS, CodiEmpr, USARECORD, FECHEMISI, TITURECORD, ConsultaSql)"
      SQLX$ = SQLX$ + "VALUES ( " & IDLISTAX&
      SQLX$ = SQLX$ + ",'PTOVTAEXSIM'"
      SQLX$ = SQLX$ + ",'PUNTOS DE VENTA EXPORTA SIMPLE'"
      SQLX$ = SQLX$ + ",1"
      SQLX$ = SQLX$ + ",1"
      SQLX$ = SQLX$ + " ,0"
      SQLX$ = SQLX$ + " ,0"
      SQLX$ = SQLX$ + " ,0"
      SQLX$ = SQLX$ + " ,5"
      SQLX$ = SQLX$ + " ,0"
      SQLX$ = SQLX$ + " ,0"
      SQLX$ = SQLX$ + " ,'15/05/2019'"
      SQLX$ = SQLX$ + " ,''"
      SQLX$ = SQLX$ + " ,'SELECT CAMPO1, CAMPO2, CAMPO3,CAMPO4 FROM DATASQL WITH(NOLOCK)'"

      SQLX$ = SQLX$ + " )"
      FLEXCONN.Execute (SQLX$)
        SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
        SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
        SQLX$ = SQLX$ + " ,0"
        SQLX$ = SQLX$ + " ,0"
        SQLX$ = SQLX$ + " , 'P.Venta/A5'"
        SQLX$ = SQLX$ + " , ''"
        SQLX$ = SQLX$ + " , 's'"
        SQLX$ = SQLX$ + " , ''"
        SQLX$ = SQLX$ + " , ''"
        SQLX$ = SQLX$ + " , ''"
        SQLX$ = SQLX$ + " , ''"
        SQLX$ = SQLX$ + " )"
        FLEXCONN.Execute (SQLX$)

    '
        SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
        SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
        SQLX$ = SQLX$ + " ,0"
        SQLX$ = SQLX$ + " ,0"
        SQLX$ = SQLX$ + " , 'Cuenta/F8'"
        SQLX$ = SQLX$ + " , ''"
        SQLX$ = SQLX$ + " , 's'"
        SQLX$ = SQLX$ + " , 'nume_cli;raso_cli;proved12'"
        SQLX$ = SQLX$ + " , ''"
        SQLX$ = SQLX$ + " , ''"
        SQLX$ = SQLX$ + " , ''"
        SQLX$ = SQLX$ + " )"
        FLEXCONN.Execute (SQLX$)

    '
        SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
        SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
        SQLX$ = SQLX$ + " ,0"
        SQLX$ = SQLX$ + " ,0"
        SQLX$ = SQLX$ + " , 'Proveedor/A64'"
        SQLX$ = SQLX$ + " , '     '"
        SQLX$ = SQLX$ + " , 'n'"
        SQLX$ = SQLX$ + " , '     '"
        SQLX$ = SQLX$ + " , ''"
        SQLX$ = SQLX$ + " , ''"
        SQLX$ = SQLX$ + " , ''"
        SQLX$ = SQLX$ + " )"
        FLEXCONN.Execute (SQLX$)

    '
        SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
        SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
        SQLX$ = SQLX$ + " ,0"
        SQLX$ = SQLX$ + " ,0"
        SQLX$ = SQLX$ + " , 'Id/a0'"
        SQLX$ = SQLX$ + " , ''"
        SQLX$ = SQLX$ + " , 's'"
        SQLX$ = SQLX$ + " , '     '"
        SQLX$ = SQLX$ + " , '     '"
        SQLX$ = SQLX$ + " , ''"
        SQLX$ = SQLX$ + " , ''"
        SQLX$ = SQLX$ + " )"
        FLEXCONN.Execute (SQLX$)

    '
    End If
'    |FIN PTOVTAEXSIM
    If InStr(1, EPAC$, "DAI ICHI") > 0 Then
    
            '|INICIO FILTROSEC
            '
            CONDI$ = " LISTA =   'FILTROSEC'"
            EXILISTA% = CantRegistros&("ENCAPRFUN", CONDI$, "NOMESS")
            If EXILISTA% > 0 Then
              FE% = 16387
              FEBD% = CVINT(VALOBADA("ENCAPRFUN", "FECHEMISI", CONDI$, "NOMESS"))
              GENERAREST% = 0
              If FEBD% < FE% Then
                IDLISTAX& = XVALO(VALOBADA("ENCAPRFUN", "ID_LISTA", CONDI$, "NOMESS"))
                Call BORRAREGISTROS("ENCAPRFUN", "ID_LISTA = " & IDLISTAX&, REGAF&, "NOMESS")
                Call BORRAREGISTROS("ENCAPRF", "ID_LISTA = " & IDLISTAX&, REGAF&, "NOMESS")
                GENERAREST% = 1
              End If
            Else
              GENERAREST% = 1
              CONDI$ = " ID_LISTA  > 0 "
              IDLISTAX& = 1 + XVALO(VALOBADA("ENCAPRFUN", "MAX(ID_LISTA)", CONDI$, "NOMESS"))
            End If
            If GENERAREST% > 0 Then
              SQLX$ = "INSERT INTO  ENCAPRFUN (ID_LISTA ,LISTA, Descripcion, ADDREGISTROS, USAFORMULA, BORRAR, INSERTAR, NOTILDE "
              SQLX$ = SQLX$ + ", CANTICAMPOS, CodiEmpr, USARECORD, FECHEMISI, TITURECORD, ConsultaSql)"
              SQLX$ = SQLX$ + "VALUES ( " & IDLISTAX&
              SQLX$ = SQLX$ + ",'FILTROSEC'"
              SQLX$ = SQLX$ + ",'Filtros Armado de  Secuencias'"
              SQLX$ = SQLX$ + ",1"
              SQLX$ = SQLX$ + ",1"
              SQLX$ = SQLX$ + " ,0"
              SQLX$ = SQLX$ + " ,0"
              SQLX$ = SQLX$ + " ,0"
              SQLX$ = SQLX$ + " ,3"
              SQLX$ = SQLX$ + " ,0"
              SQLX$ = SQLX$ + " ,1"
              SQLX$ = SQLX$ + " ,'03/05/19'"
              SQLX$ = SQLX$ + " ,'FILTRO SECUENCIAS'"
              SQLX$ = SQLX$ + " ,'SELECT CAMPO1, CAMPO2,CAMPO3,CAMPO4 FROM DATASQL WITH(NOLOCK)'"
              SQLX$ = SQLX$ + ")"
              FLEXCONN.Execute (SQLX$)
                SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
                SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
                SQLX$ = SQLX$ + " ,0"
                SQLX$ = SQLX$ + " ,0"
                SQLX$ = SQLX$ + " , 'Código/F8'"
                SQLX$ = SQLX$ + " , ''"
                SQLX$ = SQLX$ + " , '     '"
                SQLX$ = SQLX$ + " , ''"
                SQLX$ = SQLX$ + " , ''"
                SQLX$ = SQLX$ + " , ''"
                SQLX$ = SQLX$ + " , ''"
                SQLX$ = SQLX$ + " )"
                FLEXCONN.Execute (SQLX$)
            
            '
                SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
                SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
                SQLX$ = SQLX$ + " ,0"
                SQLX$ = SQLX$ + " ,0"
                SQLX$ = SQLX$ + " , 'Descripción/A48     '"
                SQLX$ = SQLX$ + " , ''"
                SQLX$ = SQLX$ + " , 's'"
                SQLX$ = SQLX$ + " , ''"
                SQLX$ = SQLX$ + " , ''"
                SQLX$ = SQLX$ + " , ''"
                SQLX$ = SQLX$ + " , ''"
                SQLX$ = SQLX$ + " )"
                FLEXCONN.Execute (SQLX$)
            
                SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
                SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
                SQLX$ = SQLX$ + " ,0"
                SQLX$ = SQLX$ + " ,0"
                SQLX$ = SQLX$ + " , 'GRUPO/A48     '"
                SQLX$ = SQLX$ + " , ''"
                SQLX$ = SQLX$ + " , 's'"
                SQLX$ = SQLX$ + " , ''"
                SQLX$ = SQLX$ + " , ''"
                SQLX$ = SQLX$ + " , ''"
                SQLX$ = SQLX$ + " , ''"
                SQLX$ = SQLX$ + " )"
                FLEXCONN.Execute (SQLX$)
            
                SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
                SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
                SQLX$ = SQLX$ + " ,0"
                SQLX$ = SQLX$ + " ,0"
                SQLX$ = SQLX$ + " , 'OBJETO/A12     '"
                SQLX$ = SQLX$ + " , ''"
                SQLX$ = SQLX$ + " , 's'"
                SQLX$ = SQLX$ + " , ''"
                SQLX$ = SQLX$ + " , ''"
                SQLX$ = SQLX$ + " , ''"
                SQLX$ = SQLX$ + " , ''"
                SQLX$ = SQLX$ + " )"
                FLEXCONN.Execute (SQLX$)
           
            '
            End If
            '|FIN FILTROSEC
    
    
    
            'COMIENZO
        'INICIO CONFIGLOGA
        '
        CONDI$ = " LISTA =   'CONFIGLOGA'"
        EXILISTA% = CantRegistros&("ENCAPRFUN", CONDI$, "NOMESS")
        If EXILISTA% > 0 Then
          FE% = 16347
          FEBD% = CVINT(VALOBADA("ENCAPRFUN", "FECHEMISI", CONDI$, "NOMESS"))
          GENERAREST% = 0
          If FEBD% < FE% Then
            IDLISTAX& = XVALO(VALOBADA("ENCAPRFUN", "ID_LISTA", CONDI$, "NOMESS"))
            Call BORRAREGISTROS("ENCAPRFUN", "ID_LISTA = " & IDLISTAX&, REGAF&, "NOMESS")
            Call BORRAREGISTROS("ENCAPRF", "ID_LISTA = " & IDLISTAX&, REGAF&, "NOMESS")
            GENERAREST% = 1
          End If
        Else
          GENERAREST% = 1
          CONDI$ = " ID_LISTA  > 0 "
          IDLISTAX& = 1 + XVALO(VALOBADA("ENCAPRFUN", "MAX(ID_LISTA)", CONDI$, "NOMESS"))
        End If
        If GENERAREST% > 0 Then
          SQLX$ = "INSERT INTO  ENCAPRFUN (ID_LISTA ,LISTA, Descripcion, ADDREGISTROS, USAFORMULA, BORRAR, INSERTAR, NOTILDE "
          SQLX$ = SQLX$ + ", CANTICAMPOS, CodiEmpr, USARECORD, FECHEMISI, TITURECORD, ConsultaSql)"
          SQLX$ = SQLX$ + "VALUES ( " & IDLISTAX&
          SQLX$ = SQLX$ + ",'CONFIGLOGA'"
          SQLX$ = SQLX$ + ",'CONFIGURACION DE TABLA LOGARITMICA'"
          SQLX$ = SQLX$ + ",1"
          SQLX$ = SQLX$ + ",1"
          SQLX$ = SQLX$ + " ,1"
          SQLX$ = SQLX$ + " ,1"
          SQLX$ = SQLX$ + " ,0"
          SQLX$ = SQLX$ + " ,9"
          SQLX$ = SQLX$ + " ,0"
          SQLX$ = SQLX$ + " ,1"
          SQLX$ = SQLX$ + " ,'27/03/19'"
          SQLX$ = SQLX$ + " ,''"
          SQLX$ = SQLX$ + " ,'SELECT CAMPO1,CAMPO2,CAMPO3, CAMPO4,CAMPO5,CAMPO6,CAMPO7,CAMPO8 FROM DATASQL WITH(NOLOCK)'"
          SQLX$ = SQLX$ + ")"
          FLEXCONN.Execute (SQLX$)
            SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
            SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
            SQLX$ = SQLX$ + " ,0"
            SQLX$ = SQLX$ + " ,0"
            SQLX$ = SQLX$ + " , 'CAMPO-CODIGO/A18'"
            SQLX$ = SQLX$ + " , '     '"
            SQLX$ = SQLX$ + " , 'S'"
            SQLX$ = SQLX$ + " , '     '"
            SQLX$ = SQLX$ + " , '     '"
            SQLX$ = SQLX$ + " , '     '"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " )"
            FLEXCONN.Execute (SQLX$)
        
        '
            SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
            SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
            SQLX$ = SQLX$ + " ,0"
            SQLX$ = SQLX$ + " ,0"
            SQLX$ = SQLX$ + " , 'CARA/A18'"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " , 'S'"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " )"
            FLEXCONN.Execute (SQLX$)
        
        '
            SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
            SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
            SQLX$ = SQLX$ + " ,0"
            SQLX$ = SQLX$ + " ,0"
            SQLX$ = SQLX$ + " , 'TIPO MATERIAL/A18'"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " , 'S'"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " )"
            FLEXCONN.Execute (SQLX$)
        
        '
            SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
            SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
            SQLX$ = SQLX$ + " ,0"
            SQLX$ = SQLX$ + " ,0"
            SQLX$ = SQLX$ + " , 'ESP. MATERIAL/F12.2'"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " , 'S'"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " )"
            FLEXCONN.Execute (SQLX$)
        
        '
            SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
            SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
            SQLX$ = SQLX$ + " ,0"
            SQLX$ = SQLX$ + " ,0"
            SQLX$ = SQLX$ + " , 'ESP. COBRE/F12'"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " , 'S'"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " )"
            FLEXCONN.Execute (SQLX$)
        
        '
            SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
            SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
            SQLX$ = SQLX$ + " ,0"
            SQLX$ = SQLX$ + " ,0"
            SQLX$ = SQLX$ + " , 'PTH/A6'"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " , 'S'"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " )"
            FLEXCONN.Execute (SQLX$)
        
        '
            SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
            SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
            SQLX$ = SQLX$ + " ,0"
            SQLX$ = SQLX$ + " ,0"
            SQLX$ = SQLX$ + " , 'PISTAS/A12     '"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " , 'S'"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " )"
            FLEXCONN.Execute (SQLX$)
        
        
            SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
            SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
            SQLX$ = SQLX$ + " ,0"
            SQLX$ = SQLX$ + " ,0"
            SQLX$ = SQLX$ + " , 'TABLA/A18     '"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " , 'S'"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " )"
            FLEXCONN.Execute (SQLX$)
        '
        End If
        '|FIN CONFIGLOGA
            

            
    '|INICIO CARAS
        '
        CONDI$ = " LISTA =   'CARAS'"
        EXILISTA% = CantRegistros&("ENCAPRFUN", CONDI$, "NOMESS")
        If EXILISTA% > 0 Then
          FE% = 16308
          FEBD% = CVINT(VALOBADA("ENCAPRFUN", "FECHEMISI", CONDI$, "NOMESS"))
          GENERAREST% = 0
          If FEBD% < FE% Then
            IDLISTAX& = XVALO(VALOBADA("ENCAPRFUN", "ID_LISTA", CONDI$, "NOMESS"))
            Call BORRAREGISTROS("ENCAPRFUN", "ID_LISTA = " & IDLISTAX&, REGAF&, "NOMESS")
            Call BORRAREGISTROS("ENCAPRF", "ID_LISTA = " & IDLISTAX&, REGAF&, "NOMESS")
            GENERAREST% = 1
          End If
        Else
          GENERAREST% = 1
          CONDI$ = " ID_LISTA  > 0 "
          IDLISTAX& = 1 + XVALO(VALOBADA("ENCAPRFUN", "MAX(ID_LISTA)", CONDI$, "NOMESS"))
        End If
        If GENERAREST% > 0 Then
          SQLX$ = "INSERT INTO  ENCAPRFUN (ID_LISTA ,LISTA, Descripcion, ADDREGISTROS, USAFORMULA, BORRAR, INSERTAR, NOTILDE "
          SQLX$ = SQLX$ + ", CANTICAMPOS, CodiEmpr, USARECORD, FECHEMISI, TITURECORD, ConsultaSql)"
          SQLX$ = SQLX$ + "VALUES ( " & IDLISTAX&
          SQLX$ = SQLX$ + ",'CARAS'"
          SQLX$ = SQLX$ + ",'CONFIGURACION DE CARAS '"
          SQLX$ = SQLX$ + ",1"
          SQLX$ = SQLX$ + ",1"
          SQLX$ = SQLX$ + " ,0"
          SQLX$ = SQLX$ + " ,0"
          SQLX$ = SQLX$ + " ,0"
          SQLX$ = SQLX$ + " ,3"
          SQLX$ = SQLX$ + " ,0"
          SQLX$ = SQLX$ + " ,1"
          SQLX$ = SQLX$ + " ,'20/02/19'"
          SQLX$ = SQLX$ + " ,'CARAS '"
          SQLX$ = SQLX$ + " ,'SELECT CAMPO1,CAMPO2 FROM DATASQL WITH(NOLOCK)'"
          SQLX$ = SQLX$ + ")"
          FLEXCONN.Execute (SQLX$)
            SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
            SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
            SQLX$ = SQLX$ + " ,0"
            SQLX$ = SQLX$ + " ,0"
            SQLX$ = SQLX$ + " , 'CODIGO/F8'"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " , 'S'"
            SQLX$ = SQLX$ + " , '     '"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " )"
            FLEXCONN.Execute (SQLX$)
        
        '
            SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
            SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
            SQLX$ = SQLX$ + " ,0"
            SQLX$ = SQLX$ + " ,0"
            SQLX$ = SQLX$ + " , 'REFERENCIA/A32'"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " , 'S'"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " )"
            FLEXCONN.Execute (SQLX$)
        
        '
        End If
        '|FIN CARAS
        
        '|INICIO ESPEMATE
        '
        CONDI$ = " LISTA =   'ESPEMATE'"
        EXILISTA% = CantRegistros&("ENCAPRFUN", CONDI$, "NOMESS")
        If EXILISTA% > 0 Then
          FE% = 16308
          FEBD% = CVINT(VALOBADA("ENCAPRFUN", "FECHEMISI", CONDI$, "NOMESS"))
          GENERAREST% = 0
          If FEBD% < FE% Then
            IDLISTAX& = XVALO(VALOBADA("ENCAPRFUN", "ID_LISTA", CONDI$, "NOMESS"))
            Call BORRAREGISTROS("ENCAPRFUN", "ID_LISTA = " & IDLISTAX&, REGAF&, "NOMESS")
            Call BORRAREGISTROS("ENCAPRF", "ID_LISTA = " & IDLISTAX&, REGAF&, "NOMESS")
            GENERAREST% = 1
          End If
        Else
          GENERAREST% = 1
          CONDI$ = " ID_LISTA  > 0 "
          IDLISTAX& = 1 + XVALO(VALOBADA("ENCAPRFUN", "MAX(ID_LISTA)", CONDI$, "NOMESS"))
        End If
        If GENERAREST% > 0 Then
          SQLX$ = "INSERT INTO  ENCAPRFUN (ID_LISTA ,LISTA, Descripcion, ADDREGISTROS, USAFORMULA, BORRAR, INSERTAR, NOTILDE "
          SQLX$ = SQLX$ + ", CANTICAMPOS, CodiEmpr, USARECORD, FECHEMISI, TITURECORD, ConsultaSql)"
          SQLX$ = SQLX$ + "VALUES ( " & IDLISTAX&
          SQLX$ = SQLX$ + ",'ESPEMATE'"
          SQLX$ = SQLX$ + ",'ESPESOR DE MATERIALES'"
          SQLX$ = SQLX$ + ",1"
          SQLX$ = SQLX$ + ",1"
          SQLX$ = SQLX$ + " ,0"
          SQLX$ = SQLX$ + " ,0"
          SQLX$ = SQLX$ + " ,0"
          SQLX$ = SQLX$ + " ,3"
          SQLX$ = SQLX$ + " ,0"
          SQLX$ = SQLX$ + " ,1"
          SQLX$ = SQLX$ + " ,'20/02/19'"
          SQLX$ = SQLX$ + " ,''"
          SQLX$ = SQLX$ + " ,'SELECT CAMPO1, CAMPO2 FROM DATASQL WITH(NOLOCK)'"
          SQLX$ = SQLX$ + ")"
          FLEXCONN.Execute (SQLX$)
            SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
            SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
            SQLX$ = SQLX$ + " ,0"
            SQLX$ = SQLX$ + " ,0"
            SQLX$ = SQLX$ + " , 'CODIGO/F8'"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " , 'S'"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " )"
            FLEXCONN.Execute (SQLX$)
        
        '
            SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
            SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
            SQLX$ = SQLX$ + " ,0"
            SQLX$ = SQLX$ + " ,0"
            SQLX$ = SQLX$ + " , 'REFERENCIA/A30'"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " , 'S'"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " )"
            FLEXCONN.Execute (SQLX$)
        
        '
        End If
        '|FIN ESPEMATE
        '|INICIO ESPECOBRE
        '
        CONDI$ = " LISTA =   'ESPECOBRE'"
        EXILISTA% = CantRegistros&("ENCAPRFUN", CONDI$, "NOMESS")
        If EXILISTA% > 0 Then
          FE% = 16308
          FEBD% = CVINT(VALOBADA("ENCAPRFUN", "FECHEMISI", CONDI$, "NOMESS"))
          GENERAREST% = 0
          If FEBD% < FE% Then
            IDLISTAX& = XVALO(VALOBADA("ENCAPRFUN", "ID_LISTA", CONDI$, "NOMESS"))
            Call BORRAREGISTROS("ENCAPRFUN", "ID_LISTA = " & IDLISTAX&, REGAF&, "NOMESS")
            Call BORRAREGISTROS("ENCAPRF", "ID_LISTA = " & IDLISTAX&, REGAF&, "NOMESS")
            GENERAREST% = 1
          End If
        Else
          GENERAREST% = 1
          CONDI$ = " ID_LISTA  > 0 "
          IDLISTAX& = 1 + XVALO(VALOBADA("ENCAPRFUN", "MAX(ID_LISTA)", CONDI$, "NOMESS"))
        End If
        If GENERAREST% > 0 Then
          SQLX$ = "INSERT INTO  ENCAPRFUN (ID_LISTA ,LISTA, Descripcion, ADDREGISTROS, USAFORMULA, BORRAR, INSERTAR, NOTILDE "
          SQLX$ = SQLX$ + ", CANTICAMPOS, CodiEmpr, USARECORD, FECHEMISI, TITURECORD, ConsultaSql)"
          SQLX$ = SQLX$ + "VALUES ( " & IDLISTAX&
          SQLX$ = SQLX$ + ",'ESPECOBRE'"
          SQLX$ = SQLX$ + ",'ESPESOR DE COBRE'"
          SQLX$ = SQLX$ + ",1"
          SQLX$ = SQLX$ + ",1"
          SQLX$ = SQLX$ + " ,0"
          SQLX$ = SQLX$ + " ,0"
          SQLX$ = SQLX$ + " ,0"
          SQLX$ = SQLX$ + " ,3"
          SQLX$ = SQLX$ + " ,0"
          SQLX$ = SQLX$ + " ,1"
          SQLX$ = SQLX$ + " ,'20/02/19'"
          SQLX$ = SQLX$ + " ,'ESPECOBRE'"
          SQLX$ = SQLX$ + " ,'SELECT CAMPO1, CAMPO2 FROM DATASQL WITH(NOLOCK)'"
          SQLX$ = SQLX$ + ")"
          FLEXCONN.Execute (SQLX$)
            SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
            SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
            SQLX$ = SQLX$ + " ,0"
            SQLX$ = SQLX$ + " ,0"
            SQLX$ = SQLX$ + " , 'CODIGO/F8'"
            SQLX$ = SQLX$ + " , '     '"
            SQLX$ = SQLX$ + " , 'S'"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " )"
            FLEXCONN.Execute (SQLX$)
        
        '
            SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
            SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
            SQLX$ = SQLX$ + " ,0"
            SQLX$ = SQLX$ + " ,0"
            SQLX$ = SQLX$ + " , 'ESPESOR/A12'"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " , 'S'"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " )"
            FLEXCONN.Execute (SQLX$)
        
        '
        End If
        '|FIN ESPECOBRE
        '|INICIO PTH
        '
        CONDI$ = " LISTA =   'PTH'"
        EXILISTA% = CantRegistros&("ENCAPRFUN", CONDI$, "NOMESS")
        If EXILISTA% > 0 Then
          FE% = 16308
          FEBD% = CVINT(VALOBADA("ENCAPRFUN", "FECHEMISI", CONDI$, "NOMESS"))
          GENERAREST% = 0
          If FEBD% < FE% Then
            IDLISTAX& = XVALO(VALOBADA("ENCAPRFUN", "ID_LISTA", CONDI$, "NOMESS"))
            Call BORRAREGISTROS("ENCAPRFUN", "ID_LISTA = " & IDLISTAX&, REGAF&, "NOMESS")
            Call BORRAREGISTROS("ENCAPRF", "ID_LISTA = " & IDLISTAX&, REGAF&, "NOMESS")
            GENERAREST% = 1
          End If
        Else
          GENERAREST% = 1
          CONDI$ = " ID_LISTA  > 0 "
          IDLISTAX& = 1 + XVALO(VALOBADA("ENCAPRFUN", "MAX(ID_LISTA)", CONDI$, "NOMESS"))
        End If
        If GENERAREST% > 0 Then
          SQLX$ = "INSERT INTO  ENCAPRFUN (ID_LISTA ,LISTA, Descripcion, ADDREGISTROS, USAFORMULA, BORRAR, INSERTAR, NOTILDE "
          SQLX$ = SQLX$ + ", CANTICAMPOS, CodiEmpr, USARECORD, FECHEMISI, TITURECORD, ConsultaSql)"
          SQLX$ = SQLX$ + "VALUES ( " & IDLISTAX&
          SQLX$ = SQLX$ + ",'PTH'"
          SQLX$ = SQLX$ + ",'PTH'"
          SQLX$ = SQLX$ + ",1"
          SQLX$ = SQLX$ + ",1"
          SQLX$ = SQLX$ + " ,0"
          SQLX$ = SQLX$ + " ,0"
          SQLX$ = SQLX$ + " ,0"
          SQLX$ = SQLX$ + " ,3"
          SQLX$ = SQLX$ + " ,0"
          SQLX$ = SQLX$ + " ,1"
          SQLX$ = SQLX$ + " ,'20/02/19'"
          SQLX$ = SQLX$ + " ,'PTH'"
          SQLX$ = SQLX$ + " ,'SELECT CAMPO1,CAMPO2 FROM DATASQL WITH(NOLOCK)'"
        '
          SQLX$ = SQLX$ + ")"
          FLEXCONN.Execute (SQLX$)
            SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
            SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
            SQLX$ = SQLX$ + " ,0"
            SQLX$ = SQLX$ + " ,0"
            SQLX$ = SQLX$ + " , 'CODIGO/F8'"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " , 'S'"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " )"
            FLEXCONN.Execute (SQLX$)
        
        '
            SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
            SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
            SQLX$ = SQLX$ + " ,0"
            SQLX$ = SQLX$ + " ,0"
            SQLX$ = SQLX$ + " , 'REFERENCIA/A12'"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " , 'S'"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " )"
            FLEXCONN.Execute (SQLX$)
        
        '
        End If
        '|FIN PTH
        
        '|INICIO PISTAS
        '
        CONDI$ = " LISTA =   'PISTAS'"
        EXILISTA% = CantRegistros&("ENCAPRFUN", CONDI$, "NOMESS")
        If EXILISTA% > 0 Then
          FE% = 16308
          FEBD% = CVINT(VALOBADA("ENCAPRFUN", "FECHEMISI", CONDI$, "NOMESS"))
          GENERAREST% = 0
          If FEBD% < FE% Then
            IDLISTAX& = XVALO(VALOBADA("ENCAPRFUN", "ID_LISTA", CONDI$, "NOMESS"))
            Call BORRAREGISTROS("ENCAPRFUN", "ID_LISTA = " & IDLISTAX&, REGAF&, "NOMESS")
            Call BORRAREGISTROS("ENCAPRF", "ID_LISTA = " & IDLISTAX&, REGAF&, "NOMESS")
            GENERAREST% = 1
          End If
        Else
          GENERAREST% = 1
          CONDI$ = " ID_LISTA  > 0 "
          IDLISTAX& = 1 + XVALO(VALOBADA("ENCAPRFUN", "MAX(ID_LISTA)", CONDI$, "NOMESS"))
        End If
        If GENERAREST% > 0 Then
          SQLX$ = "INSERT INTO  ENCAPRFUN (ID_LISTA ,LISTA, Descripcion, ADDREGISTROS, USAFORMULA, BORRAR, INSERTAR, NOTILDE "
          SQLX$ = SQLX$ + ", CANTICAMPOS, CodiEmpr, USARECORD, FECHEMISI, TITURECORD, ConsultaSql)"
          SQLX$ = SQLX$ + "VALUES ( " & IDLISTAX&
          SQLX$ = SQLX$ + ",'PISTAS'"
          SQLX$ = SQLX$ + ",'TABLA DE PISTAS'"
          SQLX$ = SQLX$ + ",1"
          SQLX$ = SQLX$ + ",1"
          SQLX$ = SQLX$ + " ,0"
          SQLX$ = SQLX$ + " ,0"
          SQLX$ = SQLX$ + " ,0"
          SQLX$ = SQLX$ + " ,3"
          SQLX$ = SQLX$ + " ,0"
          SQLX$ = SQLX$ + " ,1"
          SQLX$ = SQLX$ + " ,'20/02/19'"
          SQLX$ = SQLX$ + " ,'PISTAS'"
          SQLX$ = SQLX$ + " ,'SELECT CAMPO1,CAMPO2 FROM DATASQL WITH(NOLOCK)'"
          SQLX$ = SQLX$ + ")"
          FLEXCONN.Execute (SQLX$)
            SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
            SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
            SQLX$ = SQLX$ + " ,0"
            SQLX$ = SQLX$ + " ,0"
            SQLX$ = SQLX$ + " , 'CODIGO/F8'"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " , 'S'"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " )"
            FLEXCONN.Execute (SQLX$)
        
        '
            SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
            SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
            SQLX$ = SQLX$ + " ,0"
            SQLX$ = SQLX$ + " ,0"
            SQLX$ = SQLX$ + " , 'REFERENCIA/A32'"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " , 'S'"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " , '     '"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " )"
            FLEXCONN.Execute (SQLX$)
        
        '
        End If
        '|FIN PISTAS
    End If
    '|INICIO Ctp.CantHilos
    '
    CONDI$ = " LISTA =   'Ctp.CantHilos'"
    EXILISTA% = CantRegistros&("ENCAPRFUN", CONDI$, "NOMESS")
    If EXILISTA% > 0 Then
      FE% = 16546
      FEBD% = CVINT(VALOBADA("ENCAPRFUN", "FECHEMISI", CONDI$, "NOMESS"))
      GENERAREST% = 0
      If FEBD% < FE% Then
        IDLISTAX& = XVALO(VALOBADA("ENCAPRFUN", "ID_LISTA", CONDI$, "NOMESS"))
        Call BORRAREGISTROS("ENCAPRFUN", "ID_LISTA = " & IDLISTAX&, REGAF&, "NOMESS")
        Call BORRAREGISTROS("ENCAPRF", "ID_LISTA = " & IDLISTAX&, REGAF&, "NOMESS")
        GENERAREST% = 1
      
      End If
    Else
      GENERAREST% = 1
      CONDI$ = " ID_LISTA  > 0 "
      IDLISTAX& = 1 + XVALO(VALOBADA("ENCAPRFUN", "MAX(ID_LISTA)", CONDI$, "NOMESS"))
    End If
    If GENERAREST% > 0 Then
      SQLX$ = "INSERT INTO  ENCAPRFUN (ID_LISTA ,LISTA, Descripcion, ADDREGISTROS, USAFORMULA, BORRAR, INSERTAR, NOTILDE "
      SQLX$ = SQLX$ + ", CANTICAMPOS, CodiEmpr, USARECORD, FECHEMISI, TITURECORD, ConsultaSql)"
      SQLX$ = SQLX$ + "VALUES ( " & IDLISTAX&
      SQLX$ = SQLX$ + ",'Ctp.CantHilos'"
      SQLX$ = SQLX$ + ",'CANTIDAD DE HILOS BOLSAS '"
      SQLX$ = SQLX$ + ",1"
      SQLX$ = SQLX$ + ",1"
      SQLX$ = SQLX$ + " ,1"
      SQLX$ = SQLX$ + " ,0"
      SQLX$ = SQLX$ + " ,0"
      SQLX$ = SQLX$ + " ,3"
      SQLX$ = SQLX$ + " ,0"
      SQLX$ = SQLX$ + " ,0"
      SQLX$ = SQLX$ + " ,'02/10/19'"
      SQLX$ = SQLX$ + " ,''"
      SQLX$ = SQLX$ + " ,'SELECT CAMPO1,CAMPO2,CAMPO3 FROM DATASQL WITH(NOLOCK)'"
      SQLX$ = SQLX$ + ")"
      FLEXCONN.Execute (SQLX$)
        SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
        SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
        SQLX$ = SQLX$ + " ,0"
        SQLX$ = SQLX$ + " ,0"
        SQLX$ = SQLX$ + " , 'CODIGO/F8'"
        SQLX$ = SQLX$ + " , ''"
        SQLX$ = SQLX$ + " , ' '"
        SQLX$ = SQLX$ + " , ''"
        SQLX$ = SQLX$ + " , ''"
        SQLX$ = SQLX$ + " , ''"
        SQLX$ = SQLX$ + " , ''"
        SQLX$ = SQLX$ + " )"
        FLEXCONN.Execute (SQLX$)
    
    '
        SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
        SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
        SQLX$ = SQLX$ + " ,0"
        SQLX$ = SQLX$ + " ,0"
        SQLX$ = SQLX$ + " , 'CANTIDAD/F8'"
        SQLX$ = SQLX$ + " , ''"
        SQLX$ = SQLX$ + " , 'S'"
        SQLX$ = SQLX$ + " , ''"
        SQLX$ = SQLX$ + " , ''"
        SQLX$ = SQLX$ + " , ''"
        SQLX$ = SQLX$ + " , ''"
        SQLX$ = SQLX$ + " )"
        FLEXCONN.Execute (SQLX$)
    
        SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
        SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
        SQLX$ = SQLX$ + " ,0"
        SQLX$ = SQLX$ + " ,0"
        SQLX$ = SQLX$ + " , 'NV/A0'"
        SQLX$ = SQLX$ + " , ''"
        SQLX$ = SQLX$ + " , ''"
        SQLX$ = SQLX$ + " , ''"
        SQLX$ = SQLX$ + " , ''"
        SQLX$ = SQLX$ + " , ''"
        SQLX$ = SQLX$ + " , ''"
        SQLX$ = SQLX$ + " )"
        FLEXCONN.Execute (SQLX$)
   
    '
    End If
    '|FIN Ctp.CantHilos


'|INICIO  CTEPROCOMPRA
'
CONDI$ = " LISTA =   'CTEPROCOMPRA'"
EXILISTA% = CantRegistros&("ENCAPRFUN", CONDI$, "NOMESS")
If EXILISTA% > 0 Then
  FE% = 18137
  FEBD% = CVINT(VALOBADA("ENCAPRFUN", "FECHEMISI", CONDI$, "NOMESS"))
  GENERAREST% = 0
  If FEBD% < FE% Then
    IDLISTAX& = XVALO(VALOBADA("ENCAPRFUN", "ID_LISTA", CONDI$, "NOMESS"))
    Call BORRAREGISTROS("ENCAPRFUN", "ID_LISTA = " & IDLISTAX&, REGAF&, "NOMESS")
    Call BORRAREGISTROS("ENCAPRF", "ID_LISTA = " & IDLISTAX&, REGAF&, "NOMESS")
    GENERAREST% = 1
  End If
Else
  GENERAREST% = 1
  CONDI$ = " ID_LISTA  > 0 "
  IDLISTAX& = 1 + XVALO(VALOBADA("ENCAPRFUN", "MAX(ID_LISTA)", CONDI$, "NOMESS"))
End If
If GENERAREST% > 0 Then
  SQLX$ = "INSERT INTO  ENCAPRFUN (ID_LISTA ,LISTA, Descripcion, ADDREGISTROS, USAFORMULA, BORRAR, INSERTAR, NOTILDE "
  SQLX$ = SQLX$ + ", CANTICAMPOS, CodiEmpr, USARECORD, FECHEMISI, TITURECORD, ConsultaSql)"
  SQLX$ = SQLX$ + "VALUES ( " & IDLISTAX&
  SQLX$ = SQLX$ + ",'CTEPROCOMPRA'"
  SQLX$ = SQLX$ + ",'CTEPROCOMPRA'"
  SQLX$ = SQLX$ + ",1"
  SQLX$ = SQLX$ + ",1"
  SQLX$ = SQLX$ + " ,1"
  SQLX$ = SQLX$ + " ,1"
  SQLX$ = SQLX$ + " ,0"
  SQLX$ = SQLX$ + " ,5"
  SQLX$ = SQLX$ + " ,0"
  SQLX$ = SQLX$ + " ,0"
  SQLX$ = SQLX$ + " ,'11/7/2023'"
  SQLX$ = SQLX$ + " ,''"
  SQLX$ = SQLX$ + " ,'SELECT CAMPO1, CAMPO2, CAMPO3, CAMPO4, CAMPO5 , CAMPO6,CAMPO7 FROM DATASQL WITH(NOLOCK)'"
  SQLX$ = SQLX$ + ")"
  FLEXCONN.Execute (SQLX$)
    SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
    SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
    SQLX$ = SQLX$ + " ,0"
    SQLX$ = SQLX$ + " ,0"
    SQLX$ = SQLX$ + " , 'Ord/F10'"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , 'N'"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " )"
    FLEXCONN.Execute (SQLX$)

'
    SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
    SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
    SQLX$ = SQLX$ + " ,0"
    SQLX$ = SQLX$ + " ,0"
    SQLX$ = SQLX$ + " , 'Cta/I10'"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , 'S'"
    SQLX$ = SQLX$ + " , 'nume_cli;raso_cli;deudor12;stat_cli>0;order by raso_cli'"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " )"
    FLEXCONN.Execute (SQLX$)

'
    SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
    SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
    SQLX$ = SQLX$ + " ,0"
    SQLX$ = SQLX$ + " ,0"
    SQLX$ = SQLX$ + " , 'Cliente/A32'"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , 'N'"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ' '"
    SQLX$ = SQLX$ + " )"
    FLEXCONN.Execute (SQLX$)

'
    SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
    SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
    SQLX$ = SQLX$ + " ,0"
    SQLX$ = SQLX$ + " ,0"
    SQLX$ = SQLX$ + " , 'Cta/I10'"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , 'S'"
    SQLX$ = SQLX$ + " , 'nume_cli;raso_cli;proved12;stat_cli>0;order by raso_cli'"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " )"
    FLEXCONN.Execute (SQLX$)

'
    SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
    SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
    SQLX$ = SQLX$ + " ,0"
    SQLX$ = SQLX$ + " ,0"
    SQLX$ = SQLX$ + " , 'Proveedor/A32'"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , 'N'"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " )"
    FLEXCONN.Execute (SQLX$)
    
    SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
    SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
    SQLX$ = SQLX$ + " ,0"
    SQLX$ = SQLX$ + " ,0"
    SQLX$ = SQLX$ + " , 'Referencia/A48'"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , 'S'"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " )"
    FLEXCONN.Execute (SQLX$)

    SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
    SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
    SQLX$ = SQLX$ + " ,0"
    SQLX$ = SQLX$ + " ,0"
    SQLX$ = SQLX$ + " , 'Ruta(256)/A256'"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , 'S'"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " )"
    FLEXCONN.Execute (SQLX$)


'
End If
'|FIN CTEPROCOMPRA
'|INICIO CTP.HILACHA
'|INICIO CTP.HILACHA
                '
                CONDI$ = " LISTA =   'CTP.HILACHA'"
                EXILISTA% = CantRegistros&("ENCAPRFUN", CONDI$, "NOMESS")
                If EXILISTA% > 0 Then
                  FE% = 16547
                  FEBD% = CVINT(VALOBADA("ENCAPRFUN", "FECHEMISI", CONDI$, "NOMESS"))
                  GENERAREST% = 0
                  If FEBD% < FE% Then
                    IDLISTAX& = XVALO(VALOBADA("ENCAPRFUN", "ID_LISTA", CONDI$, "NOMESS"))
                    Call BORRAREGISTROS("ENCAPRFUN", "ID_LISTA = " & IDLISTAX&, REGAF&, "NOMESS")
                    Call BORRAREGISTROS("ENCAPRF", "ID_LISTA = " & IDLISTAX&, REGAF&, "NOMESS")
                    GENERAREST% = 1
                  End If
                Else
                  GENERAREST% = 1
                  CONDI$ = " ID_LISTA  > 0 "
                  IDLISTAX& = 1 + XVALO(VALOBADA("ENCAPRFUN", "MAX(ID_LISTA)", CONDI$, "NOMESS"))
                End If
                If GENERAREST% > 0 Then
                  SQLX$ = "INSERT INTO  ENCAPRFUN (ID_LISTA ,LISTA, Descripcion, ADDREGISTROS, USAFORMULA, BORRAR, INSERTAR, NOTILDE "
                  SQLX$ = SQLX$ + ", CANTICAMPOS, CodiEmpr, USARECORD, FECHEMISI, TITURECORD, ConsultaSql)"
                  SQLX$ = SQLX$ + "VALUES ( " & IDLISTAX&
                  SQLX$ = SQLX$ + ",'CTP.HILACHA'"
                  SQLX$ = SQLX$ + ",'TAPA GR'"
                  SQLX$ = SQLX$ + ",1"
                  SQLX$ = SQLX$ + ",1"
                  SQLX$ = SQLX$ + " ,0"
                  SQLX$ = SQLX$ + " ,0"
                  SQLX$ = SQLX$ + " ,0"
                  SQLX$ = SQLX$ + " ,4"
                  SQLX$ = SQLX$ + " ,0"
                  SQLX$ = SQLX$ + " ,0"
                  SQLX$ = SQLX$ + " ,'03/10/19'"
                  SQLX$ = SQLX$ + " ,''"
                  SQLX$ = SQLX$ + " ,'SELECT CAMPO1,CAMPO2,CAMPO3 FROM DATASQL WITH(NOLOCK)'"
                  SQLX$ = SQLX$ + ")"
                  FLEXCONN.Execute (SQLX$)
                    SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
                    SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
                    SQLX$ = SQLX$ + " ,0"
                    SQLX$ = SQLX$ + " ,0"
                    SQLX$ = SQLX$ + " , 'ID/F8'"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " , ' '"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " )"
                    FLEXCONN.Execute (SQLX$)
                
                '
                    SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
                    SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
                    SQLX$ = SQLX$ + " ,0"
                    SQLX$ = SQLX$ + " ,0"
                    SQLX$ = SQLX$ + " , 'TAPA GR/A32'"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " , 'S'"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " )"
                    FLEXCONN.Execute (SQLX$)
                
                '
                    SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
                    SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
                    SQLX$ = SQLX$ + " ,0"
                    SQLX$ = SQLX$ + " ,0"
                    SQLX$ = SQLX$ + " , 'NV/A0'"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " )"
                    FLEXCONN.Execute (SQLX$)
                
                '
                End If
                '|FIN CTP.HILACHA


'|INICIO CTP.DETEMETAL
                '
                CONDI$ = " LISTA =   'CTP.DETEMETAL'"
                EXILISTA% = CantRegistros&("ENCAPRFUN", CONDI$, "NOMESS")
                If EXILISTA% > 0 Then
                  FE% = 16547
                  FEBD% = CVINT(VALOBADA("ENCAPRFUN", "FECHEMISI", CONDI$, "NOMESS"))
                  GENERAREST% = 0
                  If FEBD% < FE% Then
                    IDLISTAX& = XVALO(VALOBADA("ENCAPRFUN", "ID_LISTA", CONDI$, "NOMESS"))
                    Call BORRAREGISTROS("ENCAPRFUN", "ID_LISTA = " & IDLISTAX&, REGAF&, "NOMESS")
                    Call BORRAREGISTROS("ENCAPRF", "ID_LISTA = " & IDLISTAX&, REGAF&, "NOMESS")
                    GENERAREST% = 1
                  End If
                Else
                  GENERAREST% = 1
                  CONDI$ = " ID_LISTA  > 0 "
                  IDLISTAX& = 1 + XVALO(VALOBADA("ENCAPRFUN", "MAX(ID_LISTA)", CONDI$, "NOMESS"))
                End If
                If GENERAREST% > 0 Then
                  SQLX$ = "INSERT INTO  ENCAPRFUN (ID_LISTA ,LISTA, Descripcion, ADDREGISTROS, USAFORMULA, BORRAR, INSERTAR, NOTILDE "
                  SQLX$ = SQLX$ + ", CANTICAMPOS, CodiEmpr, USARECORD, FECHEMISI, TITURECORD, ConsultaSql)"
                  SQLX$ = SQLX$ + "VALUES ( " & IDLISTAX&
                  SQLX$ = SQLX$ + ",'CTP.DETEMETAL'"
                  SQLX$ = SQLX$ + ",'TAPA GR'"
                  SQLX$ = SQLX$ + ",1"
                  SQLX$ = SQLX$ + ",1"
                  SQLX$ = SQLX$ + " ,0"
                  SQLX$ = SQLX$ + " ,0"
                  SQLX$ = SQLX$ + " ,0"
                  SQLX$ = SQLX$ + " ,4"
                  SQLX$ = SQLX$ + " ,0"
                  SQLX$ = SQLX$ + " ,0"
                  SQLX$ = SQLX$ + " ,'03/10/19'"
                  SQLX$ = SQLX$ + " ,''"
                  SQLX$ = SQLX$ + " ,'SELECT CAMPO1,CAMPO2,CAMPO3 FROM DATASQL WITH(NOLOCK)'"
                  SQLX$ = SQLX$ + ")"
                  FLEXCONN.Execute (SQLX$)
                    SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
                    SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
                    SQLX$ = SQLX$ + " ,0"
                    SQLX$ = SQLX$ + " ,0"
                    SQLX$ = SQLX$ + " , 'ID/F8'"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " , ' '"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " )"
                    FLEXCONN.Execute (SQLX$)
                
                '
                    SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
                    SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
                    SQLX$ = SQLX$ + " ,0"
                    SQLX$ = SQLX$ + " ,0"
                    SQLX$ = SQLX$ + " , 'TAPA GR/A32'"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " , 'S'"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " )"
                    FLEXCONN.Execute (SQLX$)
                
                '
                    SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
                    SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
                    SQLX$ = SQLX$ + " ,0"
                    SQLX$ = SQLX$ + " ,0"
                    SQLX$ = SQLX$ + " , 'NV/A0'"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " )"
                    FLEXCONN.Execute (SQLX$)
                
                '
                End If
                '|FIN CTP.DETEMETAL



End Sub
Sub GENERAR_ESTRUCTURAS2()

        '|INICIO COLORES
        '
        CONDI$ = " LISTA =   'COLORES'"
        EXILISTA% = CantRegistros&("ENCAPRFUN", CONDI$, "NOMESS")
        If EXILISTA% > 0 Then
          FE% = 16612
          FEBD% = CVINT(VALOBADA("ENCAPRFUN", "FECHEMISI", CONDI$, "NOMESS"))
          GENERAREST% = 0
          If FEBD% < FE% Then
            IDLISTAX& = XVALO(VALOBADA("ENCAPRFUN", "ID_LISTA", CONDI$, "NOMESS"))
            Call BORRAREGISTROS("ENCAPRFUN", "ID_LISTA = " & IDLISTAX&, REGAF&, "NOMESS")
            Call BORRAREGISTROS("ENCAPRF", "ID_LISTA = " & IDLISTAX&, REGAF&, "NOMESS")
            GENERAREST% = 1
          End If
        Else
          GENERAREST% = 1
          CONDI$ = " ID_LISTA  > 0 "
          IDLISTAX& = 1 + XVALO(VALOBADA("ENCAPRFUN", "MAX(ID_LISTA)", CONDI$, "NOMESS"))
        End If
        If GENERAREST% > 0 Then
          SQLX$ = "INSERT INTO  ENCAPRFUN (ID_LISTA ,LISTA, Descripcion, ADDREGISTROS, USAFORMULA, BORRAR, INSERTAR, NOTILDE "
          SQLX$ = SQLX$ + ", CANTICAMPOS, CodiEmpr, USARECORD, FECHEMISI, TITURECORD, ConsultaSql)"
          SQLX$ = SQLX$ + "VALUES ( " & IDLISTAX&
          SQLX$ = SQLX$ + ",'COLORES'"
          SQLX$ = SQLX$ + ",'TABLA DE COLORES'"
          SQLX$ = SQLX$ + ",1"
          SQLX$ = SQLX$ + ",1"
          SQLX$ = SQLX$ + " ,1"
          SQLX$ = SQLX$ + " ,1"
          SQLX$ = SQLX$ + " ,0"
          SQLX$ = SQLX$ + " ,3"
          SQLX$ = SQLX$ + " ,0"
          SQLX$ = SQLX$ + " ,0"
          SQLX$ = SQLX$ + " ,'04/12/19'"
          SQLX$ = SQLX$ + " ,''"
          SQLX$ = SQLX$ + " ,'SELECT CAMPO1,CAMPO2 FROM DATASQL WITH(NOLOCK)'"
          SQLX$ = SQLX$ + ")"
          FLEXCONN.Execute (SQLX$)
            SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
            SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
            SQLX$ = SQLX$ + " ,0"
            SQLX$ = SQLX$ + " ,0"
            SQLX$ = SQLX$ + " , 'CODIGO/F8'"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " , 'N'"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " )"
            FLEXCONN.Execute (SQLX$)
        
            '
            SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
            SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
            SQLX$ = SQLX$ + " ,0"
            SQLX$ = SQLX$ + " ,0"
            SQLX$ = SQLX$ + " , 'COLOR/A24'"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " , 'S'"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " )"
            FLEXCONN.Execute (SQLX$)
        
        '
        End If
        '|FIN COLORES
        
        
        '|INICIO CONFUBIX
        '
        CONDI$ = " LISTA =   'CONFUBIX'"
        EXILISTA% = CantRegistros&("ENCAPRFUN", CONDI$, "NOMESS")
        If EXILISTA% > 0 Then
          FE% = 16759
          FEBD% = CVINT(VALOBADA("ENCAPRFUN", "FECHEMISI", CONDI$, "NOMESS"))
          GENERAREST% = 0
          If FEBD% < FE% Then
            IDLISTAX& = XVALO(VALOBADA("ENCAPRFUN", "ID_LISTA", CONDI$, "NOMESS"))
            Call BORRAREGISTROS("ENCAPRFUN", "ID_LISTA = " & IDLISTAX&, REGAF&, "NOMESS")
            Call BORRAREGISTROS("ENCAPRF", "ID_LISTA = " & IDLISTAX&, REGAF&, "NOMESS")
            GENERAREST% = 1
          End If
        Else
          GENERAREST% = 1
          CONDI$ = " ID_LISTA  > 0 "
          IDLISTAX& = 1 + XVALO(VALOBADA("ENCAPRFUN", "MAX(ID_LISTA)", CONDI$, "NOMESS"))
        End If
        If GENERAREST% > 0 Then
          SQLX$ = "INSERT INTO  ENCAPRFUN (ID_LISTA ,LISTA, Descripcion, ADDREGISTROS, USAFORMULA, BORRAR, INSERTAR, NOTILDE "
          SQLX$ = SQLX$ + ", CANTICAMPOS, CodiEmpr, USARECORD, FECHEMISI, TITURECORD, ConsultaSql)"
          SQLX$ = SQLX$ + "VALUES ( " & IDLISTAX&
          SQLX$ = SQLX$ + ",'CONFUBIX'"
          SQLX$ = SQLX$ + ",'CONFIGURACION DE UBICACIONES'"
          SQLX$ = SQLX$ + ",1"
          SQLX$ = SQLX$ + ",1"
          SQLX$ = SQLX$ + " ,0"
          SQLX$ = SQLX$ + " ,1"
          SQLX$ = SQLX$ + " ,0"
          SQLX$ = SQLX$ + " ,4"
          SQLX$ = SQLX$ + " ,0"
          SQLX$ = SQLX$ + " ,0"
          SQLX$ = SQLX$ + " ,'23/03/20'"
          SQLX$ = SQLX$ + " ,''"
          SQLX$ = SQLX$ + " ,'SELECT CAMPO1, CAMPO2,CAMPO3,CAMPO4 FROM DATASQL WITH(NOLOCK) '"
          SQLX$ = SQLX$ + ")"
          FLEXCONN.Execute (SQLX$)
          
          SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
          SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
          SQLX$ = SQLX$ + " ,0"
          SQLX$ = SQLX$ + " ,0"
          SQLX$ = SQLX$ + " , 'N° DE RACK/F8'"
          SQLX$ = SQLX$ + " , ' '"
          SQLX$ = SQLX$ + " , 'S'"
          SQLX$ = SQLX$ + " , ' '"
          SQLX$ = SQLX$ + " , ' '"
          SQLX$ = SQLX$ + " , ' '"
          SQLX$ = SQLX$ + " , ' '"
          SQLX$ = SQLX$ + " )"
          FLEXCONN.Execute (SQLX$)
            
            SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
            SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
            SQLX$ = SQLX$ + " ,0"
            SQLX$ = SQLX$ + " ,0"
            SQLX$ = SQLX$ + " , 'DESCRIPCION/A32'"
            SQLX$ = SQLX$ + " , ' '"
            SQLX$ = SQLX$ + " , 'S'"
            SQLX$ = SQLX$ + " , ' '"
            SQLX$ = SQLX$ + " , ' '"
            SQLX$ = SQLX$ + " , ' '"
            SQLX$ = SQLX$ + " , ' '"
            SQLX$ = SQLX$ + " )"
            FLEXCONN.Execute (SQLX$)
            '
            SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
            SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
            SQLX$ = SQLX$ + " ,0"
            SQLX$ = SQLX$ + " ,0"
            SQLX$ = SQLX$ + " , 'CANT.FILAS./F9'"
            SQLX$ = SQLX$ + " , ' '"
            SQLX$ = SQLX$ + " , 'S'"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " )"
            FLEXCONN.Execute (SQLX$)
        
            '
            SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
            SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
            SQLX$ = SQLX$ + " ,0"
            SQLX$ = SQLX$ + " ,0"
            SQLX$ = SQLX$ + " , 'CANT.COL./F9'"
            SQLX$ = SQLX$ + " , ' '"
            SQLX$ = SQLX$ + " , 'S'"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " )"
            FLEXCONN.Execute (SQLX$)
        
        '
        End If
        '|FIN CONFUBIX

        '|INICIO PROGRAPCB
        '
        EPAC$ = TRIM$(UCase$(EMPREAC$))
        CONDI$ = " LISTA =   'PROGRAPCB'"
        EXILISTA% = CantRegistros&("ENCAPRFUN", CONDI$, "NOMESS")
        If EXILISTA% > 0 Then
          FE% = 16308
          FEBD% = CVINT(VALOBADA("ENCAPRFUN", "FECHEMISI", CONDI$, "NOMESS"))
          GENERAREST% = 0
          If FEBD% < FE% Then
            IDLISTAX& = XVALO(VALOBADA("ENCAPRFUN", "ID_LISTA", CONDI$, "NOMESS"))
            Call BORRAREGISTROS("ENCAPRFUN", "ID_LISTA = " & IDLISTAX&, REGAF&, "NOMESS")
            Call BORRAREGISTROS("ENCAPRF", "ID_LISTA = " & IDLISTAX&, REGAF&, "NOMESS")
            GENERAREST% = 1
          End If
        Else
          GENERAREST% = 1
          CONDI$ = " ID_LISTA  > 0 "
          IDLISTAX& = 1 + XVALO(VALOBADA("ENCAPRFUN", "MAX(ID_LISTA)", CONDI$, "NOMESS"))
        End If
        If GENERAREST% > 0 Then
          SQLX$ = "INSERT INTO  ENCAPRFUN (ID_LISTA ,LISTA, Descripcion, ADDREGISTROS, USAFORMULA, BORRAR, INSERTAR, NOTILDE "
          SQLX$ = SQLX$ + ", CANTICAMPOS, CodiEmpr, USARECORD, FECHEMISI, TITURECORD, ConsultaSql)"
          SQLX$ = SQLX$ + "VALUES ( " & IDLISTAX&
          SQLX$ = SQLX$ + ",'PROGRAPCB'"
          SQLX$ = SQLX$ + ",'TABLA PROGRAMA PCB'"
          SQLX$ = SQLX$ + ",1"
          SQLX$ = SQLX$ + ",1"
          SQLX$ = SQLX$ + " ,0"
          SQLX$ = SQLX$ + " ,0"
          SQLX$ = SQLX$ + " ,0"
          SQLX$ = SQLX$ + " ,3"
          SQLX$ = SQLX$ + " ,0"
          SQLX$ = SQLX$ + " ,1"
          SQLX$ = SQLX$ + " ,'20/02/19'"
          SQLX$ = SQLX$ + " ,'PCB'"
          SQLX$ = SQLX$ + " ,'SELECT CAMPO1,CAMPO2 FROM DATASQL WITH(NOLOCK)'"
          SQLX$ = SQLX$ + ")"
          FLEXCONN.Execute (SQLX$)
            SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
            SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
            SQLX$ = SQLX$ + " ,0"
            SQLX$ = SQLX$ + " ,0"
            SQLX$ = SQLX$ + " , 'CODIGO/F8'"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " , 'S'"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " )"
            FLEXCONN.Execute (SQLX$)
        
        '
            SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
            SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
            SQLX$ = SQLX$ + " ,0"
            SQLX$ = SQLX$ + " ,0"
            SQLX$ = SQLX$ + " , 'REFERENCIA/A48'"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " , 'S'"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " )"
            FLEXCONN.Execute (SQLX$)
        
        '
        End If
                '|FIN PROGRAPCB
        '|INICIO MASCANTSOLDANTE
        '
        CONDI$ = " LISTA =   'MASCANTSOLDANTE'"
        EXILISTA% = CantRegistros&("ENCAPRFUN", CONDI$, "NOMESS")
        If EXILISTA% > 0 Then
          FE% = 16308
          FEBD% = CVINT(VALOBADA("ENCAPRFUN", "FECHEMISI", CONDI$, "NOMESS"))
          GENERAREST% = 0
          If FEBD% < FE% Then
            IDLISTAX& = XVALO(VALOBADA("ENCAPRFUN", "ID_LISTA", CONDI$, "NOMESS"))
            Call BORRAREGISTROS("ENCAPRFUN", "ID_LISTA = " & IDLISTAX&, REGAF&, "NOMESS")
            Call BORRAREGISTROS("ENCAPRF", "ID_LISTA = " & IDLISTAX&, REGAF&, "NOMESS")
            GENERAREST% = 1
          End If
        Else
          GENERAREST% = 1
          CONDI$ = " ID_LISTA  > 0 "
          IDLISTAX& = 1 + XVALO(VALOBADA("ENCAPRFUN", "MAX(ID_LISTA)", CONDI$, "NOMESS"))
        End If
        If GENERAREST% > 0 Then
          SQLX$ = "INSERT INTO  ENCAPRFUN (ID_LISTA ,LISTA, Descripcion, ADDREGISTROS, USAFORMULA, BORRAR, INSERTAR, NOTILDE "
          SQLX$ = SQLX$ + ", CANTICAMPOS, CodiEmpr, USARECORD, FECHEMISI, TITURECORD, ConsultaSql)"
          SQLX$ = SQLX$ + "VALUES ( " & IDLISTAX&
          SQLX$ = SQLX$ + ",'MASCANTSOLDANTE'"
          SQLX$ = SQLX$ + ",'MASCARA ANTISOLDANTE'"
          SQLX$ = SQLX$ + ",1"
          SQLX$ = SQLX$ + ",1"
          SQLX$ = SQLX$ + " ,0"
          SQLX$ = SQLX$ + " ,0"
          SQLX$ = SQLX$ + " ,0"
          SQLX$ = SQLX$ + " ,3"
          SQLX$ = SQLX$ + " ,0"
          SQLX$ = SQLX$ + " ,1"
          SQLX$ = SQLX$ + " ,'20/02/19'"
          SQLX$ = SQLX$ + " ,'MASCANTSOLDANTE'"
          SQLX$ = SQLX$ + " ,'SELECT CAMPO1,CAMPO2 FROM DATASQL WITH(NOLOCK)'"
          SQLX$ = SQLX$ + ")"
          FLEXCONN.Execute (SQLX$)
            SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
            SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
            SQLX$ = SQLX$ + " ,0"
            SQLX$ = SQLX$ + " ,0"
            SQLX$ = SQLX$ + " , 'CODIGO/F8'"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " , 'S'"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " )"
            FLEXCONN.Execute (SQLX$)
        
        '
            SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
            SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
            SQLX$ = SQLX$ + " ,0"
            SQLX$ = SQLX$ + " ,0"
            SQLX$ = SQLX$ + " , 'REFERENCIA/A48'"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " , 'S'"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " )"
            FLEXCONN.Execute (SQLX$)
        
        '
        End If
        '|FIN MASCANTSOLDANTE
        
        
        
        '|INICIO LEYECOMP
        '
        CONDI$ = " LISTA =   'LEYECOMP'"
        EXILISTA% = CantRegistros&("ENCAPRFUN", CONDI$, "NOMESS")
        If EXILISTA% > 0 Then
          FE% = 16308
          FEBD% = CVINT(VALOBADA("ENCAPRFUN", "FECHEMISI", CONDI$, "NOMESS"))
          GENERAREST% = 0
          If FEBD% < FE% Then
            IDLISTAX& = XVALO(VALOBADA("ENCAPRFUN", "ID_LISTA", CONDI$, "NOMESS"))
            Call BORRAREGISTROS("ENCAPRFUN", "ID_LISTA = " & IDLISTAX&, REGAF&, "NOMESS")
            Call BORRAREGISTROS("ENCAPRF", "ID_LISTA = " & IDLISTAX&, REGAF&, "NOMESS")
            GENERAREST% = 1
          End If
        Else
          GENERAREST% = 1
          CONDI$ = " ID_LISTA  > 0 "
          IDLISTAX& = 1 + XVALO(VALOBADA("ENCAPRFUN", "MAX(ID_LISTA)", CONDI$, "NOMESS"))
        End If
        If GENERAREST% > 0 Then
          SQLX$ = "INSERT INTO  ENCAPRFUN (ID_LISTA ,LISTA, Descripcion, ADDREGISTROS, USAFORMULA, BORRAR, INSERTAR, NOTILDE "
          SQLX$ = SQLX$ + ", CANTICAMPOS, CodiEmpr, USARECORD, FECHEMISI, TITURECORD, ConsultaSql)"
          SQLX$ = SQLX$ + "VALUES ( " & IDLISTAX&
          SQLX$ = SQLX$ + ",'LEYECOMP'"
          SQLX$ = SQLX$ + ",'COLOR LEYENDA DE COMPONENTES'"
          SQLX$ = SQLX$ + ",1"
          SQLX$ = SQLX$ + ",1"
          SQLX$ = SQLX$ + " ,0"
          SQLX$ = SQLX$ + " ,0"
          SQLX$ = SQLX$ + " ,0"
          SQLX$ = SQLX$ + " ,3"
          SQLX$ = SQLX$ + " ,0"
          SQLX$ = SQLX$ + " ,1"
          SQLX$ = SQLX$ + " ,'20/02/19'"
          SQLX$ = SQLX$ + " ,''"
          SQLX$ = SQLX$ + " ,'SELECT CAMPO1,CAMPO2 FROM DATASQL WITH(NOLOCK)'"
          SQLX$ = SQLX$ + ")"
          FLEXCONN.Execute (SQLX$)
            SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
            SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
            SQLX$ = SQLX$ + " ,0"
            SQLX$ = SQLX$ + " ,0"
            SQLX$ = SQLX$ + " , 'CODIGO/F8'"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " , 'S'"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " )"
            FLEXCONN.Execute (SQLX$)
        
        '
            SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
            SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
            SQLX$ = SQLX$ + " ,0"
            SQLX$ = SQLX$ + " ,0"
            SQLX$ = SQLX$ + " , 'REFERFENCIA/A48'"
            SQLX$ = SQLX$ + " , '     '"
            SQLX$ = SQLX$ + " , 'S'"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " )"
            FLEXCONN.Execute (SQLX$)
        
        '
        End If
        '|FIN LEYECOMP
        
        
 

    '*********************** - --- ---- FIN DAIICHI ------ ------ *************************************
    
    If InStr(1, EPAC$, "TER") > 0 Then
        '|INICIO ARMATER
        '
        CONDI$ = " LISTA =   'ARMATER'"
        EXILISTA% = CantRegistros&("ENCAPRFUN", CONDI$, "NOMESS")
        If EXILISTA% > 0 Then
          FE% = 16113
          FEBD% = CVINT(VALOBADA("ENCAPRFUN", "FECHEMISI", CONDI$, "NOMESS"))
          GENERAREST% = 0
          If FEBD% < FE% Then
            IDLISTAX& = XVALO(VALOBADA("ENCAPRFUN", "ID_LISTA", CONDI$, "NOMESS"))
            Call BORRAREGISTROS("ENCAPRFUN", "ID_LISTA = " & IDLISTAX&, REGAF&, "NOMESS")
            Call BORRAREGISTROS("ENCAPRF", "ID_LISTA = " & IDLISTAX&, REGAF&, "NOMESS")
            GENERAREST% = 1
          End If
        Else
          GENERAREST% = 1
          CONDI$ = " ID_LISTA  > 0 "
          IDLISTAX& = 1 + XVALO(VALOBADA("ENCAPRFUN", "MAX(ID_LISTA)", CONDI$, "NOMESS"))
        End If
        If GENERAREST% > 0 Then
          SQLX$ = "INSERT INTO  ENCAPRFUN (ID_LISTA ,LISTA, Descripcion, ADDREGISTROS, USAFORMULA, BORRAR, INSERTAR, NOTILDE "
          SQLX$ = SQLX$ + ", CANTICAMPOS, CodiEmpr, USARECORD, FECHEMISI, TITURECORD, ConsultaSql)"
          SQLX$ = SQLX$ + "VALUES ( " & IDLISTAX&
          SQLX$ = SQLX$ + ",'ARMATER'"
          SQLX$ = SQLX$ + ",'TABLA DE ARMADORES'"
          SQLX$ = SQLX$ + ",1"
          SQLX$ = SQLX$ + ",0"
          SQLX$ = SQLX$ + " ,1"
          SQLX$ = SQLX$ + " ,1"
          SQLX$ = SQLX$ + " ,0"
          SQLX$ = SQLX$ + " ,2"
          SQLX$ = SQLX$ + " ,0"
          SQLX$ = SQLX$ + " ,0"
          SQLX$ = SQLX$ + " ,'17/09/2018'"
          SQLX$ = SQLX$ + " ,''"
          SQLX$ = SQLX$ + " ,'SELECT CAMPO1, CAMPO2 FROM DATASQL WITH(NOLOCK)'"
          SQLX$ = SQLX$ + ")"
          FLEXCONN.Execute (SQLX$)
          
            SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
            SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
            SQLX$ = SQLX$ + " ,0"
            SQLX$ = SQLX$ + " ,0"
            SQLX$ = SQLX$ + " , 'C/A1'"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " , 's'"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " )"
            FLEXCONN.Execute (SQLX$)
          
            SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
            SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
            SQLX$ = SQLX$ + " ,0"
            SQLX$ = SQLX$ + " ,0"
            SQLX$ = SQLX$ + " , 'Armador/A48'"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " , 's'"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " )"
            FLEXCONN.Execute (SQLX$)
        '
        End If
        '|FIN ARMATER
    End If
    
    If InStr(1, EPAC$, "ARCON") > 0 Then
        '|INICIO FOCALFA
        '
        CONDI$ = " LISTA =   'FOCALFA'"
        EXILISTA% = CantRegistros&("ENCAPRFUN", CONDI$, "NOMESS")
        If EXILISTA% > 0 Then
          FE% = 16113
          FEBD% = CVINT(VALOBADA("ENCAPRFUN", "FECHEMISI", CONDI$, "NOMESS"))
          GENERAREST% = 0
          If FEBD% < FE% Then
            IDLISTAX& = XVALO(VALOBADA("ENCAPRFUN", "ID_LISTA", CONDI$, "NOMESS"))
            Call BORRAREGISTROS("ENCAPRFUN", "ID_LISTA = " & IDLISTAX&, REGAF&, "NOMESS")
            Call BORRAREGISTROS("ENCAPRF", "ID_LISTA = " & IDLISTAX&, REGAF&, "NOMESS")
            GENERAREST% = 1
          End If
        Else
          GENERAREST% = 1
          CONDI$ = " ID_LISTA  > 0 "
          IDLISTAX& = 1 + XVALO(VALOBADA("ENCAPRFUN", "MAX(ID_LISTA)", CONDI$, "NOMESS"))
        End If
        If GENERAREST% > 0 Then
          SQLX$ = "INSERT INTO  ENCAPRFUN (ID_LISTA ,LISTA, Descripcion, ADDREGISTROS, USAFORMULA, BORRAR, INSERTAR, NOTILDE "
          SQLX$ = SQLX$ + ", CANTICAMPOS, CodiEmpr, USARECORD, FECHEMISI, TITURECORD, ConsultaSql)"
          SQLX$ = SQLX$ + "VALUES ( " & IDLISTAX&
          SQLX$ = SQLX$ + ",'FOCALFA'"
          SQLX$ = SQLX$ + ",'FORMULARIOS DE CALIDAD POR FAMILIA'"
          SQLX$ = SQLX$ + ",1"
          SQLX$ = SQLX$ + ",0"
          SQLX$ = SQLX$ + " ,1"
          SQLX$ = SQLX$ + " ,1"
          SQLX$ = SQLX$ + " ,0"
          SQLX$ = SQLX$ + " ,4"
          SQLX$ = SQLX$ + " ,0"
          SQLX$ = SQLX$ + " ,0"
          SQLX$ = SQLX$ + " ,'17/09/2018'"
          SQLX$ = SQLX$ + " ,''"
          SQLX$ = SQLX$ + " ,'SELECT CAMPO1, CAMPO2, CAMPO3, CAMPO4 FROM DATASQL WITH(NOLOCK)'"
          SQLX$ = SQLX$ + ")"
          FLEXCONN.Execute (SQLX$)
          
            SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
            SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
            SQLX$ = SQLX$ + " ,0"
            SQLX$ = SQLX$ + " ,0"
            SQLX$ = SQLX$ + " , 'C/A1'"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " , 's'"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " )"
            FLEXCONN.Execute (SQLX$)
          
            SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
            SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
            SQLX$ = SQLX$ + " ,0"
            SQLX$ = SQLX$ + " ,0"
            SQLX$ = SQLX$ + " , 'Familia/A5'"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " , 's'"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " )"
            FLEXCONN.Execute (SQLX$)
        
        '
            SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
            SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
            SQLX$ = SQLX$ + " ,0"
            SQLX$ = SQLX$ + " ,0"
            SQLX$ = SQLX$ + " , 'Ruta/A128'"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " , 's'"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " )"
            FLEXCONN.Execute (SQLX$)
        
        '
            SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
            SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
            SQLX$ = SQLX$ + " ,0"
            SQLX$ = SQLX$ + " ,0"
            SQLX$ = SQLX$ + " , 'Hoja/A24'"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " , 's'"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " )"
            FLEXCONN.Execute (SQLX$)
        
        '
        End If
        '|FIN FOCALFA
    End If
    
    If InStr(EPAC$, "MEDITEA") > 0 Or InStr(EPAC$, "EYM") > 0 Then
        '|INICIO ADUANA
        '
        CONDI$ = " LISTA =   'ADUANA'"
        EXILISTA% = CantRegistros&("ENCAPRFUN", CONDI$, "NOMESS")
        If EXILISTA% > 0 Then
          FE% = 16348
          FEBD% = CVINT(VALOBADA("ENCAPRFUN", "FECHEMISI", CONDI$, "NOMESS"))
          GENERAREST% = 0
          If FEBD% < FE% Then
            IDLISTAX& = XVALO(VALOBADA("ENCAPRFUN", "ID_LISTA", CONDI$, "NOMESS"))
            Call BORRAREGISTROS("ENCAPRFUN", "ID_LISTA = " & IDLISTAX&, REGAF&, "NOMESS")
            Call BORRAREGISTROS("ENCAPRF", "ID_LISTA = " & IDLISTAX&, REGAF&, "NOMESS")
            GENERAREST% = 1
          End If
        Else
          GENERAREST% = 1
          CONDI$ = " ID_LISTA  > 0 "
          IDLISTAX& = 1 + XVALO(VALOBADA("ENCAPRFUN", "MAX(ID_LISTA)", CONDI$, "NOMESS"))
        End If
        If GENERAREST% > 0 Then
          SQLX$ = "INSERT INTO  ENCAPRFUN (ID_LISTA ,LISTA, Descripcion, ADDREGISTROS, USAFORMULA, BORRAR, INSERTAR, NOTILDE "
          SQLX$ = SQLX$ + ", CANTICAMPOS, CodiEmpr, USARECORD, FECHEMISI, TITURECORD, ConsultaSql)"
          SQLX$ = SQLX$ + "VALUES ( " & IDLISTAX&
          SQLX$ = SQLX$ + ",'ADUANA'"
          SQLX$ = SQLX$ + ",'TABLA DE ADUANA'"
          SQLX$ = SQLX$ + ",1"
          SQLX$ = SQLX$ + ",1"
          SQLX$ = SQLX$ + " ,0"
          SQLX$ = SQLX$ + " ,0"
          SQLX$ = SQLX$ + " ,0"
          SQLX$ = SQLX$ + " ,3"
          SQLX$ = SQLX$ + " ,0"
          SQLX$ = SQLX$ + " ,1"
          SQLX$ = SQLX$ + " ,'28/03/19'"
          SQLX$ = SQLX$ + " ,''"
          SQLX$ = SQLX$ + " ,'SELECT CAMPO1,CAMPO2 FROM DATASQL WITH(NOLOCK)'"
          SQLX$ = SQLX$ + ")"
          FLEXCONN.Execute (SQLX$)
            SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
            SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
            SQLX$ = SQLX$ + " ,0"
            SQLX$ = SQLX$ + " ,0"
            SQLX$ = SQLX$ + " , 'Código/F8'"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " , '     '"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " )"
            FLEXCONN.Execute (SQLX$)
        
        '
            SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
            SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
            SQLX$ = SQLX$ + " ,0"
            SQLX$ = SQLX$ + " ,0"
            SQLX$ = SQLX$ + " , 'Referencia/A24'"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " , 'S'"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " )"
            FLEXCONN.Execute (SQLX$)
        
        '
        End If
        '|FIN ADUANA
    
    
    
    
    
    
    
                '|INICIO USUASINCAR
                '
                CONDI$ = " LISTA =   'USUASINCAR'"
                EXILISTA% = CantRegistros&("ENCAPRFUN", CONDI$, "NOMESS")
                If EXILISTA% > 0 Then
                  FE% = 16113
                  FEBD% = CVINT(VALOBADA("ENCAPRFUN", "FECHEMISI", CONDI$, "NOMESS"))
                  GENERAREST% = 0
                  If FEBD% < FE% Then
                    IDLISTAX& = XVALO(VALOBADA("ENCAPRFUN", "ID_LISTA", CONDI$, "NOMESS"))
                    Call BORRAREGISTROS("ENCAPRFUN", "ID_LISTA = " & IDLISTAX&, REGAF&, "NOMESS")
                    Call BORRAREGISTROS("ENCAPRF", "ID_LISTA = " & IDLISTAX&, REGAF&, "NOMESS")
                    GENERAREST% = 1
                  End If
                Else
                  GENERAREST% = 1
                  CONDI$ = " ID_LISTA  > 0 "
                  IDLISTAX& = 1 + XVALO(VALOBADA("ENCAPRFUN", "MAX(ID_LISTA)", CONDI$, "NOMESS"))
                End If
                If GENERAREST% > 0 Then
                  SQLX$ = "INSERT INTO  ENCAPRFUN (ID_LISTA ,LISTA, Descripcion, ADDREGISTROS, USAFORMULA, BORRAR, INSERTAR, NOTILDE "
                  SQLX$ = SQLX$ + ", CANTICAMPOS, CodiEmpr, USARECORD, FECHEMISI, TITURECORD, ConsultaSql)"
                  SQLX$ = SQLX$ + "VALUES ( " & IDLISTAX&
                  SQLX$ = SQLX$ + ",'USUASINCAR'"
                  SQLX$ = SQLX$ + ",'USUARIOS CON DERECHO A ENTREGAR MATERIAL SIN CARGO'"
                  SQLX$ = SQLX$ + ",1"
                  SQLX$ = SQLX$ + ",0"
                  SQLX$ = SQLX$ + " ,1"
                  SQLX$ = SQLX$ + " ,1"
                  SQLX$ = SQLX$ + " ,0"
                  SQLX$ = SQLX$ + " ,2"
                  SQLX$ = SQLX$ + " ,0"
                  SQLX$ = SQLX$ + " ,0"
                  SQLX$ = SQLX$ + " ,'17/09/2018'"
                  SQLX$ = SQLX$ + " ,''"
                  SQLX$ = SQLX$ + " ,'SELECT CAMPO1, CAMPO2, CAMPO3 FROM DATASQL WITH(NOLOCK)'"
                  SQLX$ = SQLX$ + ")"
                  FLEXCONN.Execute (SQLX$)
                  
                    SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
                    SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
                    SQLX$ = SQLX$ + " ,0"
                    SQLX$ = SQLX$ + " ,0"
                    SQLX$ = SQLX$ + " , 'C/A1'"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " , 'n'"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " )"
                    FLEXCONN.Execute (SQLX$)
                  
                    SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
                    SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
                    SQLX$ = SQLX$ + " ,0"
                    SQLX$ = SQLX$ + " ,0"
                    SQLX$ = SQLX$ + " , 'Usuario/A48'"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " , 'n'"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " )"
                    FLEXCONN.Execute (SQLX$)
                    
                    SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
                    SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
                    SQLX$ = SQLX$ + " ,0"
                    SQLX$ = SQLX$ + " ,0"
                    SQLX$ = SQLX$ + " , 'Permiso(S-N)/A14'"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " , 's'"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " )"
                    FLEXCONN.Execute (SQLX$)
                    
                '
                End If
    
          '|INICIO TECNICOS-MDT
    
            CONDI$ = " LISTA =   'TECNICOS-MDT'"
            EXILISTA% = CantRegistros&("ENCAPRFUN", CONDI$, "NOMESS")
            If EXILISTA% > 0 Then
              FE% = 15925
              FEBD% = CVINT(VALOBADA("ENCAPRFUN", "FECHEMISI", CONDI$, "NOMESS"))
              GENERAREST% = 0
              If FEBD% < FE% Then
                IDLISTAX& = XVALO(VALOBADA("ENCAPRFUN", "ID_LISTA", CONDI$, "NOMESS"))
                Call BORRAREGISTROS("ENCAPRFUN", "ID_LISTA = " & IDLISTAX&, REGAF&, "NOMESS")
                Call BORRAREGISTROS("ENCAPRF", "ID_LISTA = " & IDLISTAX&, REGAF&, "NOMESS")
                GENERAREST% = 1
              End If
            Else
              GENERAREST% = 1
              CONDI$ = " ID_LISTA  > 0 "
              IDLISTAX& = 1 + XVALO(VALOBADA("ENCAPRFUN", "MAX(ID_LISTA)", CONDI$, "NOMESS"))
            End If
            If GENERAREST% > 0 Then
              SQLX$ = "INSERT INTO  ENCAPRFUN (ID_LISTA ,LISTA, Descripcion, ADDREGISTROS, USAFORMULA, BORRAR, INSERTAR, NOTILDE "
              SQLX$ = SQLX$ + ", CANTICAMPOS, CodiEmpr, USARECORD, FECHEMISI, TITURECORD, ConsultaSql)"
              SQLX$ = SQLX$ + "VALUES ( " & IDLISTAX&
              SQLX$ = SQLX$ + ",'TECNICOS-MDT'"
              SQLX$ = SQLX$ + ",'TABLA DE TECNICOS'"
              SQLX$ = SQLX$ + ",1"
              SQLX$ = SQLX$ + ",1"
              SQLX$ = SQLX$ + " ,1"
              SQLX$ = SQLX$ + " ,1"
              SQLX$ = SQLX$ + " ,0"
              SQLX$ = SQLX$ + " ,3"
              SQLX$ = SQLX$ + " ,0"
              SQLX$ = SQLX$ + " ,1"
              SQLX$ = SQLX$ + " ,'21/03/18'"
              SQLX$ = SQLX$ + " ,'TECNICOSMDT'"
              SQLX$ = SQLX$ + " ,'SELECT CAMPO1, CAMPO2 FROM DATASQL WITH(NOLOCK) '"
              SQLX$ = SQLX$ + ")"
              FLEXCONN.Execute (SQLX$)
                SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
                SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
                SQLX$ = SQLX$ + " ,0"
                SQLX$ = SQLX$ + " ,0"
                SQLX$ = SQLX$ + " , 'CODIGO/F6'"
                SQLX$ = SQLX$ + " , '     '"
                SQLX$ = SQLX$ + " , 'S'"
                SQLX$ = SQLX$ + " , ''"
                SQLX$ = SQLX$ + " , ''"
                SQLX$ = SQLX$ + " , ''"
                SQLX$ = SQLX$ + " , ''"
                SQLX$ = SQLX$ + " )"
                FLEXCONN.Execute (SQLX$)
            
            '
                SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
                SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
                SQLX$ = SQLX$ + " ,0"
                SQLX$ = SQLX$ + " ,0"
                SQLX$ = SQLX$ + " , 'NOMBRE/A32'"
                SQLX$ = SQLX$ + " , ''"
                SQLX$ = SQLX$ + " , 'S'"
                SQLX$ = SQLX$ + " , ''"
                SQLX$ = SQLX$ + " , ''"
                SQLX$ = SQLX$ + " , ''"
                SQLX$ = SQLX$ + " , ''"
                SQLX$ = SQLX$ + " )"
                FLEXCONN.Execute (SQLX$)
            
            '
            End If
            ''|FIN TECNICOS-MDT


            '|INICIO MEDIOS COBRANZA-MDT
                '
                CONDI$ = " LISTA =   'MEDIOS COBRANZA-MDT'"
                EXILISTA% = CantRegistros&("ENCAPRFUN", CONDI$, "NOMESS")
                If EXILISTA% > 0 Then
                  FE% = 15926
                  FEBD% = CVINT(VALOBADA("ENCAPRFUN", "FECHEMISI", CONDI$, "NOMESS"))
                  GENERAREST% = 0
                  If FEBD% < FE% Then
                    IDLISTAX& = XVALO(VALOBADA("ENCAPRFUN", "ID_LISTA", CONDI$, "NOMESS"))
                    Call BORRAREGISTROS("ENCAPRFUN", "ID_LISTA = " & IDLISTAX&, REGAF&, "NOMESS")
                    Call BORRAREGISTROS("ENCAPRF", "ID_LISTA = " & IDLISTAX&, REGAF&, "NOMESS")
                    GENERAREST% = 1
                  End If
                Else
                  GENERAREST% = 1
                  CONDI$ = " ID_LISTA  > 0 "
                  IDLISTAX& = 1 + XVALO(VALOBADA("ENCAPRFUN", "MAX(ID_LISTA)", CONDI$, "NOMESS"))
                End If
                If GENERAREST% > 0 Then
                  SQLX$ = "INSERT INTO  ENCAPRFUN (ID_LISTA ,LISTA, Descripcion, ADDREGISTROS, USAFORMULA, BORRAR, INSERTAR, NOTILDE "
                  SQLX$ = SQLX$ + ", CANTICAMPOS, CodiEmpr, USARECORD, FECHEMISI, TITURECORD, ConsultaSql)"
                  SQLX$ = SQLX$ + "VALUES ( " & IDLISTAX&
                  SQLX$ = SQLX$ + ",'MEDIOS COBRANZA-MDT'"
                  SQLX$ = SQLX$ + ",'TABLA DE MEDIOS DE COBRANZA'"
                  SQLX$ = SQLX$ + ",1"
                  SQLX$ = SQLX$ + ",1"
                  SQLX$ = SQLX$ + " ,0"
                  SQLX$ = SQLX$ + " ,0"
                  SQLX$ = SQLX$ + " ,0"
                  SQLX$ = SQLX$ + " ,4"
                  SQLX$ = SQLX$ + " ,0"
                  SQLX$ = SQLX$ + " ,1"
                  SQLX$ = SQLX$ + " ,'22/03/18'"
                  SQLX$ = SQLX$ + " ,'TABMEDCOB'"
                  SQLX$ = SQLX$ + " ,'SELECT CAMPO1,CAMPO2,CAMPO3 FROM DATASQL WITH(NOLOCK)'"
                  SQLX$ = SQLX$ + ")"
                  FLEXCONN.Execute (SQLX$)
                    SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
                    SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
                    SQLX$ = SQLX$ + " ,0"
                    SQLX$ = SQLX$ + " ,0"
                    SQLX$ = SQLX$ + " , 'CÓDIGO/F7'"
                    SQLX$ = SQLX$ + " , '     '"
                    SQLX$ = SQLX$ + " , '     '"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " )"
                    FLEXCONN.Execute (SQLX$)
                
                '
                    SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
                    SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
                    SQLX$ = SQLX$ + " ,0"
                    SQLX$ = SQLX$ + " ,0"
                    SQLX$ = SQLX$ + " , 'REFERENCIA/A48'"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " , 'S'"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " )"
                    FLEXCONN.Execute (SQLX$)
                
                '
                    SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
                    SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
                    SQLX$ = SQLX$ + " ,0"
                    SQLX$ = SQLX$ + " ,0"
                    SQLX$ = SQLX$ + " , 'PORCENTAJE/F12.2'"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " , 'S'"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " )"
                    FLEXCONN.Execute (SQLX$)
                
                '
                End If
                '|FIN MEDIOS COBRANZA-MDT
                
                '|INICIO TECNIEXT
                '
                CONDI$ = " LISTA =   'TECNIEXT'"
                EXILISTA% = CantRegistros&("ENCAPRFUN", CONDI$, "NOMESS")
                If EXILISTA% > 0 Then
                  FE% = 16004
                  FEBD% = CVINT(VALOBADA("ENCAPRFUN", "FECHEMISI", CONDI$, "NOMESS"))
                  GENERAREST% = 0
                  If FEBD% < FE% Then
                    IDLISTAX& = XVALO(VALOBADA("ENCAPRFUN", "ID_LISTA", CONDI$, "NOMESS"))
                    Call BORRAREGISTROS("ENCAPRFUN", "ID_LISTA = " & IDLISTAX&, REGAF&, "NOMESS")
                    Call BORRAREGISTROS("ENCAPRF", "ID_LISTA = " & IDLISTAX&, REGAF&, "NOMESS")
                    GENERAREST% = 1
                  End If
                Else
                  GENERAREST% = 1
                  CONDI$ = " ID_LISTA  > 0 "
                  IDLISTAX& = 1 + XVALO(VALOBADA("ENCAPRFUN", "MAX(ID_LISTA)", CONDI$, "NOMESS"))
                End If
                If GENERAREST% > 0 Then
                  SQLX$ = "INSERT INTO  ENCAPRFUN (ID_LISTA ,LISTA, Descripcion, ADDREGISTROS, USAFORMULA, BORRAR, INSERTAR, NOTILDE "
                  SQLX$ = SQLX$ + ", CANTICAMPOS, CodiEmpr, USARECORD, FECHEMISI, TITURECORD, ConsultaSql)"
                  SQLX$ = SQLX$ + "VALUES ( " & IDLISTAX&
                  SQLX$ = SQLX$ + ",'TECNIEXT'"
                  SQLX$ = SQLX$ + ",'Asociación Técnicos Externos'"
                  SQLX$ = SQLX$ + ",1"
                  SQLX$ = SQLX$ + ",1"
                  SQLX$ = SQLX$ + " ,0"
                  SQLX$ = SQLX$ + " ,0"
                  SQLX$ = SQLX$ + " ,0"
                  SQLX$ = SQLX$ + " ,7"
                  SQLX$ = SQLX$ + " ,0"
                  SQLX$ = SQLX$ + " ,0"
                  SQLX$ = SQLX$ + " ,'04/06/18'"
                  SQLX$ = SQLX$ + " ,''"
                  SQLX$ = SQLX$ + " ,'SELECT CAMPO1,CAMPO2,CAMPO3, CAMPO4,CAMPO5,CAMPO6 FROM DATASQL WITH(NOLOCK)'"
                  SQLX$ = SQLX$ + ")"
                  FLEXCONN.Execute (SQLX$)
                    SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
                    SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
                    SQLX$ = SQLX$ + " ,0"
                    SQLX$ = SQLX$ + " ,0"
                    SQLX$ = SQLX$ + " , 'Cod.Tec./f9'"
                    SQLX$ = SQLX$ + " , '     '"
                    SQLX$ = SQLX$ + " , '     '"
                    SQLX$ = SQLX$ + " , '     '"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " )"
                    FLEXCONN.Execute (SQLX$)
                
                '
                    SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
                    SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
                    SQLX$ = SQLX$ + " ,0"
                    SQLX$ = SQLX$ + " ,0"
                    SQLX$ = SQLX$ + " , 'Nombre/A32'"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " , 'S'"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " )"
                    FLEXCONN.Execute (SQLX$)
                
                '
                    SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
                    SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
                    SQLX$ = SQLX$ + " ,0"
                    SQLX$ = SQLX$ + " ,0"
                    SQLX$ = SQLX$ + " , 'Cta.Clte/F9'"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " , 'S'"
                    SQLX$ = SQLX$ + " , 'NUME_CLI;RASO_CLI;DEUDOR12'"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " )"
                    FLEXCONN.Execute (SQLX$)
                
                '
                    SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
                    SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
                    SQLX$ = SQLX$ + " ,0"
                    SQLX$ = SQLX$ + " ,0"
                    SQLX$ = SQLX$ + " , 'R.Social/A32'"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " , '     '"
                    SQLX$ = SQLX$ + " , '     '"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " )"
                    FLEXCONN.Execute (SQLX$)
                
                '
                    SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
                    SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
                    SQLX$ = SQLX$ + " ,0"
                    SQLX$ = SQLX$ + " ,0"
                    SQLX$ = SQLX$ + " , 'Cta.Prov./F9'"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " , 'S'"
                    SQLX$ = SQLX$ + " , 'NUME_CLI;RASO_CLI;PROVED12'"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " )"
                    FLEXCONN.Execute (SQLX$)
                
                '
                    SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
                    SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
                    SQLX$ = SQLX$ + " ,0"
                    SQLX$ = SQLX$ + " ,0"
                    SQLX$ = SQLX$ + " , 'R.Social/A32'"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " , ''"
                    SQLX$ = SQLX$ + " )"
                    FLEXCONN.Execute (SQLX$)
                
                '
                End If
'|FIN TECNIEXT
            '|INICIO Equipos_Importados
            '
            CONDI$ = " LISTA =   'Equipos_Importados'"
            EXILISTA% = CantRegistros&("ENCAPRFUN", CONDI$, "NOMESS")
            If EXILISTA% > 0 Then
              FE% = 16146
              FEBD% = CVINT(VALOBADA("ENCAPRFUN", "FECHEMISI", CONDI$, "NOMESS"))
              GENERAREST% = 0
              If FEBD% < FE% Then
                IDLISTAX& = XVALO(VALOBADA("ENCAPRFUN", "ID_LISTA", CONDI$, "NOMESS"))
                Call BORRAREGISTROS("ENCAPRFUN", "ID_LISTA = " & IDLISTAX&, REGAF&, "NOMESS")
                Call BORRAREGISTROS("ENCAPRF", "ID_LISTA = " & IDLISTAX&, REGAF&, "NOMESS")
                GENERAREST% = 1
              End If
            Else
              GENERAREST% = 1
              CONDI$ = " ID_LISTA  > 0 "
              IDLISTAX& = 1 + XVALO(VALOBADA("ENCAPRFUN", "MAX(ID_LISTA)", CONDI$, "NOMESS"))
            End If
            If GENERAREST% > 0 Then
              SQLX$ = "INSERT INTO  ENCAPRFUN (ID_LISTA ,LISTA, Descripcion, ADDREGISTROS, USAFORMULA, BORRAR, INSERTAR, NOTILDE "
              SQLX$ = SQLX$ + ", CANTICAMPOS, CodiEmpr, USARECORD, FECHEMISI, TITURECORD, ConsultaSql)"
              SQLX$ = SQLX$ + "VALUES ( " & IDLISTAX&
              SQLX$ = SQLX$ + ",'Equipos_Importados'"
              SQLX$ = SQLX$ + ",'Listado de Equipos Importados'"
              SQLX$ = SQLX$ + ",1"
              SQLX$ = SQLX$ + ",1"
              SQLX$ = SQLX$ + " ,1"
              SQLX$ = SQLX$ + " ,1"
              SQLX$ = SQLX$ + " ,0"
              SQLX$ = SQLX$ + " ,3"
              SQLX$ = SQLX$ + " ,0"
              SQLX$ = SQLX$ + " ,1"
              SQLX$ = SQLX$ + " ,'18/10/18'"
              SQLX$ = SQLX$ + " ,''"
              SQLX$ = SQLX$ + " ,'SELECT CAMPO1,CAMPO2 FROM DATASQL WITH(NOLOCK) '"
              SQLX$ = SQLX$ + ")"
              FLEXCONN.Execute (SQLX$)
                SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
                SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
                SQLX$ = SQLX$ + " ,0"
                SQLX$ = SQLX$ + " ,0"
                SQLX$ = SQLX$ + " , 'CODIGO/A24'"
                SQLX$ = SQLX$ + " , ''"
                SQLX$ = SQLX$ + " , 'S'"
                SQLX$ = SQLX$ + " , 'CODIGO;DESCRIPCION;MASTER'"
                SQLX$ = SQLX$ + " , ''"
                SQLX$ = SQLX$ + " , ''"
                SQLX$ = SQLX$ + " , ''"
                SQLX$ = SQLX$ + " )"
                FLEXCONN.Execute (SQLX$)
            
            '
                SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
                SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
                SQLX$ = SQLX$ + " ,0"
                SQLX$ = SQLX$ + " ,0"
                SQLX$ = SQLX$ + " , 'DESCRIPCION/A48'"
                SQLX$ = SQLX$ + " , ''"
                SQLX$ = SQLX$ + " , ''"
                SQLX$ = SQLX$ + " , ''"
                SQLX$ = SQLX$ + " , ''"
                SQLX$ = SQLX$ + " , ''"
                SQLX$ = SQLX$ + " , ''"
                SQLX$ = SQLX$ + " )"
                FLEXCONN.Execute (SQLX$)
            
            '
            End If
            '|FIN Equipos_Importados
                            
                            
             '|INICIO ART_SIN_CARGO
            '
            CONDI$ = " LISTA =   'ART_SIN_CARGO'"
            EXILISTA% = CantRegistros&("ENCAPRFUN", CONDI$, "NOMESS")
            If EXILISTA% > 0 Then
              FE% = 16150
              FEBD% = CVINT(VALOBADA("ENCAPRFUN", "FECHEMISI", CONDI$, "NOMESS"))
              GENERAREST% = 0
              If FEBD% < FE% Then
                IDLISTAX& = XVALO(VALOBADA("ENCAPRFUN", "ID_LISTA", CONDI$, "NOMESS"))
                Call BORRAREGISTROS("ENCAPRFUN", "ID_LISTA = " & IDLISTAX&, REGAF&, "NOMESS")
                Call BORRAREGISTROS("ENCAPRF", "ID_LISTA = " & IDLISTAX&, REGAF&, "NOMESS")
                GENERAREST% = 1
              End If
            Else
              GENERAREST% = 1
              CONDI$ = " ID_LISTA  > 0 "
              IDLISTAX& = 1 + XVALO(VALOBADA("ENCAPRFUN", "MAX(ID_LISTA)", CONDI$, "NOMESS"))
            End If
            If GENERAREST% > 0 Then
              SQLX$ = "INSERT INTO  ENCAPRFUN (ID_LISTA ,LISTA, Descripcion, ADDREGISTROS, USAFORMULA, BORRAR, INSERTAR, NOTILDE "
              SQLX$ = SQLX$ + ", CANTICAMPOS, CodiEmpr, USARECORD, FECHEMISI, TITURECORD, ConsultaSql)"
              SQLX$ = SQLX$ + "VALUES ( " & IDLISTAX&
              SQLX$ = SQLX$ + ",'ART_SIN_CARGO'"
              SQLX$ = SQLX$ + ",'LISTADO DE ARTICULOS SIN CARGO'"
              SQLX$ = SQLX$ + ",1"
              SQLX$ = SQLX$ + ",1"
              SQLX$ = SQLX$ + " ,1"
              SQLX$ = SQLX$ + " ,1"
              SQLX$ = SQLX$ + " ,0"
              SQLX$ = SQLX$ + " ,4"
              SQLX$ = SQLX$ + " ,0"
              SQLX$ = SQLX$ + " ,1"
              SQLX$ = SQLX$ + " ,'22/10/18'"
              SQLX$ = SQLX$ + " ,''"
              SQLX$ = SQLX$ + " ,'SELECT CAMPO1,CAMPO2,CAMPO3 FROM DATASQL WITH(NOLOCK)'"
              SQLX$ = SQLX$ + ")"
              FLEXCONN.Execute (SQLX$)
                SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
                SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
                SQLX$ = SQLX$ + " ,0"
                SQLX$ = SQLX$ + " ,0"
                SQLX$ = SQLX$ + " , 'CÓDIGO/A16'"
                SQLX$ = SQLX$ + " , '     '"
                SQLX$ = SQLX$ + " , 'S'"
                SQLX$ = SQLX$ + " , 'CODIGO;DESCRIPCION;MASTER'"
                SQLX$ = SQLX$ + " , ''"
                SQLX$ = SQLX$ + " , ''"
                SQLX$ = SQLX$ + " , ''"
                SQLX$ = SQLX$ + " )"
                FLEXCONN.Execute (SQLX$)
            
            '
                SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
                SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
                SQLX$ = SQLX$ + " ,0"
                SQLX$ = SQLX$ + " ,0"
                SQLX$ = SQLX$ + " , 'DESCRIPCIÓN/A32'"
                SQLX$ = SQLX$ + " , ''"
                SQLX$ = SQLX$ + " , ''"
                SQLX$ = SQLX$ + " , '     '"
                SQLX$ = SQLX$ + " , ''"
                SQLX$ = SQLX$ + " , ''"
                SQLX$ = SQLX$ + " , ''"
                SQLX$ = SQLX$ + " )"
                FLEXCONN.Execute (SQLX$)
            
            '
                SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
                SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
                SQLX$ = SQLX$ + " ,0"
                SQLX$ = SQLX$ + " ,0"
                SQLX$ = SQLX$ + " , 'CON AUTORIZACIÓN/F20'"
                SQLX$ = SQLX$ + " , '     '"
                SQLX$ = SQLX$ + " , 'S'"
                SQLX$ = SQLX$ + " , '     '"
                SQLX$ = SQLX$ + " , ''"
                SQLX$ = SQLX$ + " , ''"
                SQLX$ = SQLX$ + " , ''"
                SQLX$ = SQLX$ + " )"
                FLEXCONN.Execute (SQLX$)
            
            '
            End If
            '|FIN ART_SIN_CARGO
                                       
                            
    End If
    
'|INICIO GRUPOACTU
'
CONDI$ = " LISTA =   'GRUPOACTU'"
EXILISTA% = CantRegistros&("ENCAPRFUN", CONDI$, "NOMESS")
If EXILISTA% > 0 Then
  FE% = 15984
  FEBD% = CVINT(VALOBADA("ENCAPRFUN", "FECHEMISI", CONDI$, "NOMESS"))
  GENERAREST% = 0
  If FEBD% < FE% Then
    IDLISTAX& = XVALO(VALOBADA("ENCAPRFUN", "ID_LISTA", CONDI$, "NOMESS"))
    Call BORRAREGISTROS("ENCAPRFUN", "ID_LISTA = " & IDLISTAX&, REGAF&, "NOMESS")
    Call BORRAREGISTROS("ENCAPRF", "ID_LISTA = " & IDLISTAX&, REGAF&, "NOMESS")
    GENERAREST% = 1
  End If
Else
  GENERAREST% = 1
  CONDI$ = " ID_LISTA  > 0 "
  IDLISTAX& = 1 + XVALO(VALOBADA("ENCAPRFUN", "MAX(ID_LISTA)", CONDI$, "NOMESS"))
End If
If GENERAREST% > 0 Then
  SQLX$ = "INSERT INTO  ENCAPRFUN (ID_LISTA ,LISTA, Descripcion, ADDREGISTROS, USAFORMULA, BORRAR, INSERTAR, NOTILDE "
  SQLX$ = SQLX$ + ", CANTICAMPOS, CodiEmpr, USARECORD, FECHEMISI, TITURECORD, ConsultaSql)"
  SQLX$ = SQLX$ + "VALUES ( " & IDLISTAX&
  SQLX$ = SQLX$ + ",'GRUPOACTU'"
  SQLX$ = SQLX$ + ",'GRUPOS DE ACTUALIZACION'"
  SQLX$ = SQLX$ + ",1"
  SQLX$ = SQLX$ + ",0"
  SQLX$ = SQLX$ + " ,0"
  SQLX$ = SQLX$ + " ,0"
  SQLX$ = SQLX$ + " ,0"
  SQLX$ = SQLX$ + " ,3"
  SQLX$ = SQLX$ + " ,0"
  SQLX$ = SQLX$ + " ,1"
  SQLX$ = SQLX$ + " ,'16/05/18'"
  SQLX$ = SQLX$ + " ,''"
  SQLX$ = SQLX$ + " ,'SELECT CAMPO1, CAMPO2 FROM DATASQL WITH(NOLOCK)'"
  SQLX$ = SQLX$ + ")"
  FLEXCONN.Execute (SQLX$)
    SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
    SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
    SQLX$ = SQLX$ + " ,0"
    SQLX$ = SQLX$ + " ,0"
    SQLX$ = SQLX$ + " , 'ID_GRUPO/F12'"
    SQLX$ = SQLX$ + " , '     '"
    SQLX$ = SQLX$ + " , 'S'"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " )"
    FLEXCONN.Execute (SQLX$)

'
    SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
    SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
    SQLX$ = SQLX$ + " ,0"
    SQLX$ = SQLX$ + " ,0"
    SQLX$ = SQLX$ + " , 'REFERENCIA/A32'"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , 'S'"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " )"
    FLEXCONN.Execute (SQLX$)

'
End If
'|FIN GRUPOACTU

'===============================================================================================
'===============================================================================================
'===============================================================================================
'|INICIO TIPODATO
'
CONDI$ = " LISTA =   'TIPODATO'"
EXILISTA% = CantRegistros&("ENCAPRFUN", CONDI$, "NOMESS")
If EXILISTA% > 0 Then
  FE% = 16078
  FEBD% = CVINT(VALOBADA("ENCAPRFUN", "FECHEMISI", CONDI$, "NOMESS"))
  GENERAREST% = 0
  If FEBD% < FE% Then
    IDLISTAX& = XVALO(VALOBADA("ENCAPRFUN", "ID_LISTA", CONDI$, "NOMESS"))
    Call BORRAREGISTROS("ENCAPRFUN", "ID_LISTA = " & IDLISTAX&, REGAF&, "NOMESS")
    Call BORRAREGISTROS("ENCAPRF", "ID_LISTA = " & IDLISTAX&, REGAF&, "NOMESS")
    GENERAREST% = 1
  End If
Else
  GENERAREST% = 1
  CONDI$ = " ID_LISTA  > 0 "
  IDLISTAX& = 1 + XVALO(VALOBADA("ENCAPRFUN", "MAX(ID_LISTA)", CONDI$, "NOMESS"))
End If
If GENERAREST% > 0 Then
  SQLX$ = "INSERT INTO  ENCAPRFUN (ID_LISTA ,LISTA, Descripcion, ADDREGISTROS, USAFORMULA, BORRAR, INSERTAR, NOTILDE "
  SQLX$ = SQLX$ + ", CANTICAMPOS, CodiEmpr, USARECORD, FECHEMISI, TITURECORD, ConsultaSql)"
  SQLX$ = SQLX$ + "VALUES ( " & IDLISTAX&
  SQLX$ = SQLX$ + ",'TIPODATO'"
  SQLX$ = SQLX$ + ",'TIPO DE DATOS CONFIGURACION N.SERIE'"
  SQLX$ = SQLX$ + ",1"
  SQLX$ = SQLX$ + ",1"
  SQLX$ = SQLX$ + " ,0"
  SQLX$ = SQLX$ + " ,0"
  SQLX$ = SQLX$ + " ,0"
  SQLX$ = SQLX$ + " ,4"
  SQLX$ = SQLX$ + " ,0"
  SQLX$ = SQLX$ + " ,1"
  SQLX$ = SQLX$ + " ,'14/08/18'"
  SQLX$ = SQLX$ + " ,'TIPO DE DATO N.SERIE'"
  SQLX$ = SQLX$ + " ,'SELECT CAMPO1, CAMPO2,CAMPO3 FROM DATASQL WITH(NOLOCK)  '"
  SQLX$ = SQLX$ + ")"
  FLEXCONN.Execute (SQLX$)
    SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
    SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
    SQLX$ = SQLX$ + " ,0"
    SQLX$ = SQLX$ + " ,0"
    SQLX$ = SQLX$ + " , 'CODIGO/F6'"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , '     '"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " )"
    FLEXCONN.Execute (SQLX$)

'
    SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
    SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
    SQLX$ = SQLX$ + " ,0"
    SQLX$ = SQLX$ + " ,0"
    SQLX$ = SQLX$ + " , 'T.DATO/A24'"
    SQLX$ = SQLX$ + " , '     '"
    SQLX$ = SQLX$ + " , 'S'"
    SQLX$ = SQLX$ + " , '     '"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " )"
    FLEXCONN.Execute (SQLX$)

'
    SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
    SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
    SQLX$ = SQLX$ + " ,0"
    SQLX$ = SQLX$ + " ,0"
    SQLX$ = SQLX$ + " , 'ABREVIATURA/A8'"
    SQLX$ = SQLX$ + " , '     '"
    SQLX$ = SQLX$ + " , 'S'"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " )"
    FLEXCONN.Execute (SQLX$)

'
End If
'|FIN TIPODATO
 If InStr(EPAC$, "EYM") > 0 Then
             '|INICIO EquiposNoGeneranNSerie
            '
            CONDI$ = " LISTA =   'EquiposNoGeneranNSerie'"
            EXILISTA% = CantRegistros&("ENCAPRFUN", CONDI$, "NOMESS")
            If EXILISTA% > 0 Then
              FE% = 16146
              FEBD% = CVINT(VALOBADA("ENCAPRFUN", "FECHEMISI", CONDI$, "NOMESS"))
              GENERAREST% = 0
              If FEBD% < FE% Then
                IDLISTAX& = XVALO(VALOBADA("ENCAPRFUN", "ID_LISTA", CONDI$, "NOMESS"))
                Call BORRAREGISTROS("ENCAPRFUN", "ID_LISTA = " & IDLISTAX&, REGAF&, "NOMESS")
                Call BORRAREGISTROS("ENCAPRF", "ID_LISTA = " & IDLISTAX&, REGAF&, "NOMESS")
                GENERAREST% = 1
              End If
            Else
              GENERAREST% = 1
              CONDI$ = " ID_LISTA  > 0 "
              IDLISTAX& = 1 + XVALO(VALOBADA("ENCAPRFUN", "MAX(ID_LISTA)", CONDI$, "NOMESS"))
            End If
            If GENERAREST% > 0 Then
              SQLX$ = "INSERT INTO  ENCAPRFUN (ID_LISTA ,LISTA, Descripcion, ADDREGISTROS, USAFORMULA, BORRAR, INSERTAR, NOTILDE "
              SQLX$ = SQLX$ + ", CANTICAMPOS, CodiEmpr, USARECORD, FECHEMISI, TITURECORD, ConsultaSql)"
              SQLX$ = SQLX$ + "VALUES ( " & IDLISTAX&
              SQLX$ = SQLX$ + ",'EquiposNoGeneranNSerie'"
              SQLX$ = SQLX$ + ",'Listado de Equipos No Generan N.Serie'"
              SQLX$ = SQLX$ + ",1"
              SQLX$ = SQLX$ + ",1"
              SQLX$ = SQLX$ + " ,1"
              SQLX$ = SQLX$ + " ,1"
              SQLX$ = SQLX$ + " ,0"
              SQLX$ = SQLX$ + " ,3"
              SQLX$ = SQLX$ + " ,0"
              SQLX$ = SQLX$ + " ,1"
              SQLX$ = SQLX$ + " ,'18/10/18'"
              SQLX$ = SQLX$ + " ,''"
              SQLX$ = SQLX$ + " ,'SELECT CAMPO1,CAMPO2 FROM DATASQL WITH(NOLOCK) '"
              SQLX$ = SQLX$ + ")"
              FLEXCONN.Execute (SQLX$)
                SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
                SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
                SQLX$ = SQLX$ + " ,0"
                SQLX$ = SQLX$ + " ,0"
                SQLX$ = SQLX$ + " , 'CODIGO/A24'"
                SQLX$ = SQLX$ + " , ''"
                SQLX$ = SQLX$ + " , 'S'"
                SQLX$ = SQLX$ + " , 'CODIGO;DESCRIPCION;MASTER'"
                SQLX$ = SQLX$ + " , ''"
                SQLX$ = SQLX$ + " , ''"
                SQLX$ = SQLX$ + " , ''"
                SQLX$ = SQLX$ + " )"
                FLEXCONN.Execute (SQLX$)
            
            '
                SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
                SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
                SQLX$ = SQLX$ + " ,0"
                SQLX$ = SQLX$ + " ,0"
                SQLX$ = SQLX$ + " , 'DESCRIPCION/A48'"
                SQLX$ = SQLX$ + " , ''"
                SQLX$ = SQLX$ + " , ''"
                SQLX$ = SQLX$ + " , ''"
                SQLX$ = SQLX$ + " , ''"
                SQLX$ = SQLX$ + " , ''"
                SQLX$ = SQLX$ + " , ''"
                SQLX$ = SQLX$ + " )"
                FLEXCONN.Execute (SQLX$)
            
            '
            End If
            '|FIN EquiposNoGeneranNSerie
            
            '|INICIO DISPOMEDI
            '
            CONDI$ = " LISTA =   'DISPOMEDI'"
            EXILISTA% = CantRegistros&("ENCAPRFUN", CONDI$, "NOMESS")
            If EXILISTA% > 0 Then
              FE% = 16334
              FEBD% = CVINT(VALOBADA("ENCAPRFUN", "FECHEMISI", CONDI$, "NOMESS"))
              GENERAREST% = 0
              If FEBD% < FE% Then
                IDLISTAX& = XVALO(VALOBADA("ENCAPRFUN", "ID_LISTA", CONDI$, "NOMESS"))
                Call BORRAREGISTROS("ENCAPRFUN", "ID_LISTA = " & IDLISTAX&, REGAF&, "NOMESS")
                Call BORRAREGISTROS("ENCAPRF", "ID_LISTA = " & IDLISTAX&, REGAF&, "NOMESS")
                GENERAREST% = 1
              End If
            Else
              GENERAREST% = 1
              CONDI$ = " ID_LISTA  > 0 "
              IDLISTAX& = 1 + XVALO(VALOBADA("ENCAPRFUN", "MAX(ID_LISTA)", CONDI$, "NOMESS"))
            End If
            If GENERAREST% > 0 Then
              SQLX$ = "INSERT INTO  ENCAPRFUN (ID_LISTA ,LISTA, Descripcion, ADDREGISTROS, USAFORMULA, BORRAR, INSERTAR, NOTILDE "
              SQLX$ = SQLX$ + ", CANTICAMPOS, CodiEmpr, USARECORD, FECHEMISI, TITURECORD, ConsultaSql)"
              SQLX$ = SQLX$ + "VALUES ( " & IDLISTAX&
              SQLX$ = SQLX$ + ",'DISPOMEDI'"
              SQLX$ = SQLX$ + ",'DISPOSITIVO DE MEDICION - REGISTRO DE DATOS'"
              SQLX$ = SQLX$ + ",1"
              SQLX$ = SQLX$ + ",1"
              SQLX$ = SQLX$ + " ,0"
              SQLX$ = SQLX$ + " ,0"
              SQLX$ = SQLX$ + " ,0"
              SQLX$ = SQLX$ + " ,3"
              SQLX$ = SQLX$ + " ,1"
              SQLX$ = SQLX$ + " ,1"
              SQLX$ = SQLX$ + " ,'14/03/19'"
              SQLX$ = SQLX$ + " ,'DISPODATO'"
              SQLX$ = SQLX$ + " ,'SELECT CAMPO1, CAMPO2 FROM DATASQL WITH(NOLOCK)'"
              SQLX$ = SQLX$ + ")"
              FLEXCONN.Execute (SQLX$)
                SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
                SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
                SQLX$ = SQLX$ + " ,1"
                SQLX$ = SQLX$ + " ,0"
                SQLX$ = SQLX$ + " , 'CODIGO/F8'"
                SQLX$ = SQLX$ + " , ''"
                SQLX$ = SQLX$ + " , ''"
                SQLX$ = SQLX$ + " , ''"
                SQLX$ = SQLX$ + " , ''"
                SQLX$ = SQLX$ + " , ''"
                SQLX$ = SQLX$ + " , ''"
                SQLX$ = SQLX$ + " )"
                FLEXCONN.Execute (SQLX$)
            
            '
                SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
                SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
                SQLX$ = SQLX$ + " ,1"
                SQLX$ = SQLX$ + " ,0"
                SQLX$ = SQLX$ + " , 'DISP.MEDICIÓN/A24'"
                SQLX$ = SQLX$ + " , ''"
                SQLX$ = SQLX$ + " , 'S'"
                SQLX$ = SQLX$ + " , ''"
                SQLX$ = SQLX$ + " , ''"
                SQLX$ = SQLX$ + " , ''"
                SQLX$ = SQLX$ + " , ''"
                SQLX$ = SQLX$ + " )"
                FLEXCONN.Execute (SQLX$)
            
            '
            End If
            '|FIN DISPOMEDI
        '|INICIO TIPOREP
        '
        CONDI$ = " LISTA =   'TIPOREP'"
        EXILISTA% = CantRegistros&("ENCAPRFUN", CONDI$, "NOMESS")
        If EXILISTA% > 0 Then
          FE% = 16714
          FEBD% = CVINT(VALOBADA("ENCAPRFUN", "FECHEMISI", CONDI$, "NOMESS"))
          GENERAREST% = 0
          If FEBD% < FE% Then
            IDLISTAX& = XVALO(VALOBADA("ENCAPRFUN", "ID_LISTA", CONDI$, "NOMESS"))
            Call BORRAREGISTROS("ENCAPRFUN", "ID_LISTA = " & IDLISTAX&, REGAF&, "NOMESS")
            Call BORRAREGISTROS("ENCAPRF", "ID_LISTA = " & IDLISTAX&, REGAF&, "NOMESS")
            GENERAREST% = 1
          End If
        Else
          GENERAREST% = 1
          CONDI$ = " ID_LISTA  > 0 "
          IDLISTAX& = 1 + XVALO(VALOBADA("ENCAPRFUN", "MAX(ID_LISTA)", CONDI$, "NOMESS"))
        End If
        If GENERAREST% > 0 Then
          SQLX$ = "INSERT INTO  ENCAPRFUN (ID_LISTA ,LISTA, Descripcion, ADDREGISTROS, USAFORMULA, BORRAR, INSERTAR, NOTILDE "
          SQLX$ = SQLX$ + ", CANTICAMPOS, CodiEmpr, USARECORD, FECHEMISI, TITURECORD, ConsultaSql)"
          SQLX$ = SQLX$ + "VALUES ( " & IDLISTAX&
          SQLX$ = SQLX$ + ",'TIPOREP'"
          SQLX$ = SQLX$ + ",'TABLA DE TIPOREPA'"
          SQLX$ = SQLX$ + ",1"
          SQLX$ = SQLX$ + ",1"
          SQLX$ = SQLX$ + " ,1"
          SQLX$ = SQLX$ + " ,1"
          SQLX$ = SQLX$ + " ,0"
          SQLX$ = SQLX$ + " ,3"
          SQLX$ = SQLX$ + " ,0"
          SQLX$ = SQLX$ + " ,0"
          SQLX$ = SQLX$ + " ,'10/02/20'"
          SQLX$ = SQLX$ + " ,''"
          SQLX$ = SQLX$ + " ,'SELECT CAMPO1,CAMPO2 FROM DATASQL WITH(NOLOCK)'"
          SQLX$ = SQLX$ + ")"
          FLEXCONN.Execute (SQLX$)
            SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
            SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
            SQLX$ = SQLX$ + " ,0"
            SQLX$ = SQLX$ + " ,0"
            SQLX$ = SQLX$ + " , 'CODIGO/F8'"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " , 'N'"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " )"
            FLEXCONN.Execute (SQLX$)
        
        '
            SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
            SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
            SQLX$ = SQLX$ + " ,0"
            SQLX$ = SQLX$ + " ,0"
            SQLX$ = SQLX$ + " , 'T.REPARACIÓN/A24'"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " , 'S'"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " )"
            FLEXCONN.Execute (SQLX$)
        
        '
        End If
        '|FIN TIPOREP
                            
    End If

'|INICIO CODMODELO
'
        CONDI$ = " LISTA =   'CODMODELO'"
        EXILISTA% = CantRegistros&("ENCAPRFUN", CONDI$, "NOMESS")
        If EXILISTA% > 0 Then
          FE% = 16113
          FEBD% = CVINT(VALOBADA("ENCAPRFUN", "FECHEMISI", CONDI$, "NOMESS"))
          GENERAREST% = 0
          If FEBD% < FE% Then
            IDLISTAX& = XVALO(VALOBADA("ENCAPRFUN", "ID_LISTA", CONDI$, "NOMESS"))
            Call BORRAREGISTROS("ENCAPRFUN", "ID_LISTA = " & IDLISTAX&, REGAF&, "NOMESS")
            Call BORRAREGISTROS("ENCAPRF", "ID_LISTA = " & IDLISTAX&, REGAF&, "NOMESS")
            GENERAREST% = 1
          End If
        Else
          GENERAREST% = 1
          CONDI$ = " ID_LISTA  > 0 "
          IDLISTAX& = 1 + XVALO(VALOBADA("ENCAPRFUN", "MAX(ID_LISTA)", CONDI$, "NOMESS"))
        End If
        If GENERAREST% > 0 Then
          SQLX$ = "INSERT INTO  ENCAPRFUN (ID_LISTA ,LISTA, Descripcion, ADDREGISTROS, USAFORMULA, BORRAR, INSERTAR, NOTILDE "
          SQLX$ = SQLX$ + ", CANTICAMPOS, CodiEmpr, USARECORD, FECHEMISI, TITURECORD, ConsultaSql)"
          SQLX$ = SQLX$ + "VALUES ( " & IDLISTAX&
          SQLX$ = SQLX$ + ",'CODMODELO'"
          SQLX$ = SQLX$ + ",'CÓDIGOS QUE DEFINEN MODELO Y QUE IMPRIMEN ET. GRANDE N.SERIE'"
          SQLX$ = SQLX$ + ",1"
          SQLX$ = SQLX$ + ",1"
          SQLX$ = SQLX$ + " ,1"
          SQLX$ = SQLX$ + " ,1"
          SQLX$ = SQLX$ + " ,0"
          SQLX$ = SQLX$ + " ,5"
          SQLX$ = SQLX$ + " ,0"
          SQLX$ = SQLX$ + " ,1"
          SQLX$ = SQLX$ + " ,'17/09/18'"
          SQLX$ = SQLX$ + " ,''"
          SQLX$ = SQLX$ + " ,'SELECT CAMPO1,CAMPO2,CAMPO3,CAMPO4 FROM DATASQL WITH(NOLOCK)'"
        '
          SQLX$ = SQLX$ + ")"
          FLEXCONN.Execute (SQLX$)
            SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
            SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
            SQLX$ = SQLX$ + " ,0"
            SQLX$ = SQLX$ + " ,0"
            SQLX$ = SQLX$ + " , 'Código/A16'"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " , 'S'"
            SQLX$ = SQLX$ + " , 'CODIGO;DESCRIPCION;MASTER'"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " )"
            FLEXCONN.Execute (SQLX$)
        
        '
            SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
            SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
            SQLX$ = SQLX$ + " ,0"
            SQLX$ = SQLX$ + " ,0"
            SQLX$ = SQLX$ + " , 'Descripción/A28'"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " , '     '"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " )"
            FLEXCONN.Execute (SQLX$)
        
        '
            SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
            SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
            SQLX$ = SQLX$ + " ,0"
            SQLX$ = SQLX$ + " ,0"
            SQLX$ = SQLX$ + " , 'Define Modelo(s)/A16'"
            SQLX$ = SQLX$ + " , '     '"
            SQLX$ = SQLX$ + " , 'S'"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " )"
            FLEXCONN.Execute (SQLX$)
        
        '
            SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
            SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
            SQLX$ = SQLX$ + " ,0"
            SQLX$ = SQLX$ + " ,0"
            SQLX$ = SQLX$ + " , 'Imp.Etiq.Gde.(s)/A14'"
            SQLX$ = SQLX$ + " , '     '"
            SQLX$ = SQLX$ + " , 'S'"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " )"
            FLEXCONN.Execute (SQLX$)
'
   End If
   '|FIN CODMODELO
 
 
        '|INICIO CONTROLEYM
        '
        '|INICIO CONTROLEYM
        '
        CONDI$ = " LISTA =   'CONTROLEYM'"
        EXILISTA% = CantRegistros&("ENCAPRFUN", CONDI$, "NOMESS")
        If EXILISTA% > 0 Then
          FE% = 16108
          FEBD% = CVINT(VALOBADA("ENCAPRFUN", "FECHEMISI", CONDI$, "NOMESS"))
          GENERAREST% = 0
          If FEBD% < FE% Then
            IDLISTAX& = XVALO(VALOBADA("ENCAPRFUN", "ID_LISTA", CONDI$, "NOMESS"))
            Call BORRAREGISTROS("ENCAPRFUN", "ID_LISTA = " & IDLISTAX&, REGAF&, "NOMESS")
            Call BORRAREGISTROS("ENCAPRF", "ID_LISTA = " & IDLISTAX&, REGAF&, "NOMESS")
            GENERAREST% = 1
          End If
        Else
          GENERAREST% = 1
          CONDI$ = " ID_LISTA  > 0 "
          IDLISTAX& = 1 + XVALO(VALOBADA("ENCAPRFUN", "MAX(ID_LISTA)", CONDI$, "NOMESS"))
        End If
        If GENERAREST% > 0 Then
          SQLX$ = "INSERT INTO  ENCAPRFUN (ID_LISTA ,LISTA, Descripcion, ADDREGISTROS, USAFORMULA, BORRAR, INSERTAR, NOTILDE "
          SQLX$ = SQLX$ + ", CANTICAMPOS, CodiEmpr, USARECORD, FECHEMISI, TITURECORD, ConsultaSql)"
          SQLX$ = SQLX$ + "VALUES ( " & IDLISTAX&
          SQLX$ = SQLX$ + ",'CONTROLEYM'"
          SQLX$ = SQLX$ + ",'TABLA DE CONTROLES Y EXCEPCIONES '"
          SQLX$ = SQLX$ + ",1"
          SQLX$ = SQLX$ + ",1"
          SQLX$ = SQLX$ + " ,1"
          SQLX$ = SQLX$ + " ,1"
          SQLX$ = SQLX$ + " ,0"
          SQLX$ = SQLX$ + " ,6"
          SQLX$ = SQLX$ + " ,0"
          SQLX$ = SQLX$ + " ,1"
          SQLX$ = SQLX$ + " ,'12/09/18'"
          SQLX$ = SQLX$ + " ,''"
          SQLX$ = SQLX$ + " ,'SELECT CAMPO1,CAMPO2,CAMPO3,CAMPO4,CAMPO5 FROM DATASQL WITH(NOLOCK) '"
          SQLX$ = SQLX$ + ")"
          FLEXCONN.Execute (SQLX$)
            SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
            SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
            SQLX$ = SQLX$ + " ,0"
            SQLX$ = SQLX$ + " ,0"
            SQLX$ = SQLX$ + " , 'Código/A16'"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " , 's'"
            SQLX$ = SQLX$ + " , 'CODIGO;DESCRIPCION;MASTER'"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " )"
            FLEXCONN.Execute (SQLX$)
        
        '
            SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
            SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
            SQLX$ = SQLX$ + " ,0"
            SQLX$ = SQLX$ + " ,0"
            SQLX$ = SQLX$ + " , 'Descripción/A32'"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " )"
            FLEXCONN.Execute (SQLX$)
        
        '
            SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
            SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
            SQLX$ = SQLX$ + " ,0"
            SQLX$ = SQLX$ + " ,0"
            SQLX$ = SQLX$ + " , 'Ref.Control/A16'"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " , 's'"
            SQLX$ = SQLX$ + " , 'CAMPO1;CAMPO2;DATASQL;LISTA = CONFICONTROL'"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " )"
            FLEXCONN.Execute (SQLX$)
        
        '
            SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
            SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
            SQLX$ = SQLX$ + " ,0"
            SQLX$ = SQLX$ + " ,0"
            SQLX$ = SQLX$ + " , 'Referencia/A24'"
            SQLX$ = SQLX$ + " , '     '"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " , '     '"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " )"
            FLEXCONN.Execute (SQLX$)
        
        '
            SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
            SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
            SQLX$ = SQLX$ + " ,0"
            SQLX$ = SQLX$ + " ,0"
            SQLX$ = SQLX$ + " , 'Excepciones ()/A48'"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " , 'S'"
            SQLX$ = SQLX$ + " , '     '"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " )"
            FLEXCONN.Execute (SQLX$)
        
        '
        End If
'|FIN CONTROLEYM

          '|INICIO PERSOCON
        
            CONDI$ = " LISTA =   'PERSOCON'"
            EXILISTA% = CantRegistros&("ENCAPRFUN", CONDI$, "NOMESS")
            If EXILISTA% > 0 Then
              FE% = 15925
              FEBD% = CVINT(VALOBADA("ENCAPRFUN", "FECHEMISI", CONDI$, "NOMESS"))
              GENERAREST% = 0
              If FEBD% < FE% Then
                IDLISTAX& = XVALO(VALOBADA("ENCAPRFUN", "ID_LISTA", CONDI$, "NOMESS"))
                Call BORRAREGISTROS("ENCAPRFUN", "ID_LISTA = " & IDLISTAX&, REGAF&, "NOMESS")
                Call BORRAREGISTROS("ENCAPRF", "ID_LISTA = " & IDLISTAX&, REGAF&, "NOMESS")
                GENERAREST% = 1
              End If
            Else
              GENERAREST% = 1
              CONDI$ = " ID_LISTA  > 0 "
              IDLISTAX& = 1 + XVALO(VALOBADA("ENCAPRFUN", "MAX(ID_LISTA)", CONDI$, "NOMESS"))
            End If
            If GENERAREST% > 0 Then
              SQLX$ = "INSERT INTO  ENCAPRFUN (ID_LISTA ,LISTA, Descripcion, ADDREGISTROS, USAFORMULA, BORRAR, INSERTAR, NOTILDE "
              SQLX$ = SQLX$ + ", CANTICAMPOS, CodiEmpr, USARECORD, FECHEMISI, TITURECORD, ConsultaSql)"
              SQLX$ = SQLX$ + "VALUES ( " & IDLISTAX&
              SQLX$ = SQLX$ + ",'PERSOCON'"
              SQLX$ = SQLX$ + ",'TABLA PERSONAL CONTROL'"
              SQLX$ = SQLX$ + ",1"
              SQLX$ = SQLX$ + ",1"
              SQLX$ = SQLX$ + " ,1"
              SQLX$ = SQLX$ + " ,1"
              SQLX$ = SQLX$ + " ,0"
              SQLX$ = SQLX$ + " ,3"
              SQLX$ = SQLX$ + " ,0"
              SQLX$ = SQLX$ + " ,1"
              SQLX$ = SQLX$ + " ,'21/03/18'"
              SQLX$ = SQLX$ + " ,'PERSOCON'"
              SQLX$ = SQLX$ + " ,'SELECT CAMPO1, CAMPO2 FROM DATASQL WITH(NOLOCK) '"
              SQLX$ = SQLX$ + ")"
              FLEXCONN.Execute (SQLX$)
                SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
                SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
                SQLX$ = SQLX$ + " ,0"
                SQLX$ = SQLX$ + " ,0"
                SQLX$ = SQLX$ + " , 'CODIGO/F6'"
                SQLX$ = SQLX$ + " , '     '"
                SQLX$ = SQLX$ + " , 'S'"
                SQLX$ = SQLX$ + " , ''"
                SQLX$ = SQLX$ + " , ''"
                SQLX$ = SQLX$ + " , ''"
                SQLX$ = SQLX$ + " , ''"
                SQLX$ = SQLX$ + " )"
                FLEXCONN.Execute (SQLX$)
            
            '
                SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
                SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
                SQLX$ = SQLX$ + " ,0"
                SQLX$ = SQLX$ + " ,0"
                SQLX$ = SQLX$ + " , 'NOMBRE/A32'"
                SQLX$ = SQLX$ + " , ''"
                SQLX$ = SQLX$ + " , 'S'"
                SQLX$ = SQLX$ + " , ''"
                SQLX$ = SQLX$ + " , ''"
                SQLX$ = SQLX$ + " , ''"
                SQLX$ = SQLX$ + " , ''"
                SQLX$ = SQLX$ + " )"
                FLEXCONN.Execute (SQLX$)
            
            '
            End If
            ''|FIN PERSOCON
            
            ''|INICIO PERSOAPRO
        
            CONDI$ = " LISTA =   'PERSOAPRO'"
            EXILISTA% = CantRegistros&("ENCAPRFUN", CONDI$, "NOMESS")
            If EXILISTA% > 0 Then
              FE% = 15925
              FEBD% = CVINT(VALOBADA("ENCAPRFUN", "FECHEMISI", CONDI$, "NOMESS"))
              GENERAREST% = 0
              If FEBD% < FE% Then
                IDLISTAX& = XVALO(VALOBADA("ENCAPRFUN", "ID_LISTA", CONDI$, "NOMESS"))
                Call BORRAREGISTROS("ENCAPRFUN", "ID_LISTA = " & IDLISTAX&, REGAF&, "NOMESS")
                Call BORRAREGISTROS("ENCAPRF", "ID_LISTA = " & IDLISTAX&, REGAF&, "NOMESS")
                GENERAREST% = 1
              End If
            Else
              GENERAREST% = 1
              CONDI$ = " ID_LISTA  > 0 "
              IDLISTAX& = 1 + XVALO(VALOBADA("ENCAPRFUN", "MAX(ID_LISTA)", CONDI$, "NOMESS"))
            End If
            If GENERAREST% > 0 Then
              SQLX$ = "INSERT INTO  ENCAPRFUN (ID_LISTA ,LISTA, Descripcion, ADDREGISTROS, USAFORMULA, BORRAR, INSERTAR, NOTILDE "
              SQLX$ = SQLX$ + ", CANTICAMPOS, CodiEmpr, USARECORD, FECHEMISI, TITURECORD, ConsultaSql)"
              SQLX$ = SQLX$ + "VALUES ( " & IDLISTAX&
              SQLX$ = SQLX$ + ",'PERSOAPRO'"
              SQLX$ = SQLX$ + ",'TABLA PERSONAL APROBACIÓN'"
              SQLX$ = SQLX$ + ",1"
              SQLX$ = SQLX$ + ",1"
              SQLX$ = SQLX$ + " ,1"
              SQLX$ = SQLX$ + " ,1"
              SQLX$ = SQLX$ + " ,0"
              SQLX$ = SQLX$ + " ,3"
              SQLX$ = SQLX$ + " ,0"
              SQLX$ = SQLX$ + " ,1"
              SQLX$ = SQLX$ + " ,'21/03/18'"
              SQLX$ = SQLX$ + " ,'PERSOAPRO'"
              SQLX$ = SQLX$ + " ,'SELECT CAMPO1, CAMPO2 FROM DATASQL WITH(NOLOCK) '"
              SQLX$ = SQLX$ + ")"
              FLEXCONN.Execute (SQLX$)
                SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
                SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
                SQLX$ = SQLX$ + " ,0"
                SQLX$ = SQLX$ + " ,0"
                SQLX$ = SQLX$ + " , 'CODIGO/F6'"
                SQLX$ = SQLX$ + " , '     '"
                SQLX$ = SQLX$ + " , 'S'"
                SQLX$ = SQLX$ + " , ''"
                SQLX$ = SQLX$ + " , ''"
                SQLX$ = SQLX$ + " , ''"
                SQLX$ = SQLX$ + " , ''"
                SQLX$ = SQLX$ + " )"
                FLEXCONN.Execute (SQLX$)
            
            '
                SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
                SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
                SQLX$ = SQLX$ + " ,0"
                SQLX$ = SQLX$ + " ,0"
                SQLX$ = SQLX$ + " , 'NOMBRE/A32'"
                SQLX$ = SQLX$ + " , ''"
                SQLX$ = SQLX$ + " , 'S'"
                SQLX$ = SQLX$ + " , ''"
                SQLX$ = SQLX$ + " , ''"
                SQLX$ = SQLX$ + " , ''"
                SQLX$ = SQLX$ + " , ''"
                SQLX$ = SQLX$ + " )"
                FLEXCONN.Execute (SQLX$)
            
            '
            End If
            ''|FIN PERSOAPRO
            '
        '|INICIO CONFICONTROL
        '
        CONDI$ = " LISTA =   'CONFICONTROL'"
        EXILISTA% = CantRegistros&("ENCAPRFUN", CONDI$, "NOMESS")
        If EXILISTA% > 0 Then
          FE% = 16092
          FEBD% = CVINT(VALOBADA("ENCAPRFUN", "FECHEMISI", CONDI$, "NOMESS"))
          GENERAREST% = 0
          If FEBD% < FE% Then
            IDLISTAX& = XVALO(VALOBADA("ENCAPRFUN", "ID_LISTA", CONDI$, "NOMESS"))
            Call BORRAREGISTROS("ENCAPRFUN", "ID_LISTA = " & IDLISTAX&, REGAF&, "NOMESS")
            Call BORRAREGISTROS("ENCAPRF", "ID_LISTA = " & IDLISTAX&, REGAF&, "NOMESS")
            GENERAREST% = 1
          End If
        Else
          GENERAREST% = 1
          CONDI$ = " ID_LISTA  > 0 "
          IDLISTAX& = 1 + XVALO(VALOBADA("ENCAPRFUN", "MAX(ID_LISTA)", CONDI$, "NOMESS"))
        End If
        If GENERAREST% > 0 Then
          SQLX$ = "INSERT INTO  ENCAPRFUN (ID_LISTA ,LISTA, Descripcion, ADDREGISTROS, USAFORMULA, BORRAR, INSERTAR, NOTILDE "
          SQLX$ = SQLX$ + ", CANTICAMPOS, CodiEmpr, USARECORD, FECHEMISI, TITURECORD, ConsultaSql)"
          SQLX$ = SQLX$ + "VALUES ( " & IDLISTAX&
          SQLX$ = SQLX$ + ",'CONFICONTROL'"
          SQLX$ = SQLX$ + ",'CONFIGURACION DE CONTROLES DE EQUIPOS'"
          SQLX$ = SQLX$ + ",1"
          SQLX$ = SQLX$ + ",1"
          SQLX$ = SQLX$ + " ,0"
          SQLX$ = SQLX$ + " ,0"
          SQLX$ = SQLX$ + " ,0"
          SQLX$ = SQLX$ + " ,3"
          SQLX$ = SQLX$ + " ,0"
          SQLX$ = SQLX$ + " ,1"
          SQLX$ = SQLX$ + " ,'28/08/18'"
          SQLX$ = SQLX$ + " ,'CONFICTRL'"
          SQLX$ = SQLX$ + " ,'SELECT CAMPO1, CAMPO2 FROM DATASQL WITH(NOLOCK) '  "
          SQLX$ = SQLX$ + ")"
          FLEXCONN.Execute (SQLX$)
        
            SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
            SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
            SQLX$ = SQLX$ + " ,0"
            SQLX$ = SQLX$ + " ,0"
            SQLX$ = SQLX$ + " , 'Texto Presentacion/A24'"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " , 'S'"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " )"
            FLEXCONN.Execute (SQLX$)
        
        '
            SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
            SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
            SQLX$ = SQLX$ + " ,0"
            SQLX$ = SQLX$ + " ,0"
            SQLX$ = SQLX$ + " , 'Frm Original/A32'"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " , 'S'"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " )"
            FLEXCONN.Execute (SQLX$)
        
        '
        End If
        '|FIN CONFICONTROL
            
            
     
'|INICIO MÁQ. DE USUARIO
'
'|INICIO LISTABCONTA
'
CONDI$ = " LISTA =   'LISTABCONTA'"
EXILISTA% = CantRegistros&("ENCAPRFUN", CONDI$, "NOMESS")
If EXILISTA% > 0 Then
  FE% = 16450
  FEBD% = CVINT(VALOBADA("ENCAPRFUN", "FECHEMISI", CONDI$, "NOMESS"))
  GENERAREST% = 0
  If FEBD% < FE% Then
    IDLISTAX& = XVALO(VALOBADA("ENCAPRFUN", "ID_LISTA", CONDI$, "NOMESS"))
    Call BORRAREGISTROS("ENCAPRFUN", "ID_LISTA = " & IDLISTAX&, REGAF&, "NOMESS")
    Call BORRAREGISTROS("ENCAPRF", "ID_LISTA = " & IDLISTAX&, REGAF&, "NOMESS")
    GENERAREST% = 1
  End If
Else
  GENERAREST% = 1
  CONDI$ = " ID_LISTA  > 0 "
  IDLISTAX& = 1 + XVALO(VALOBADA("ENCAPRFUN", "MAX(ID_LISTA)", CONDI$, "NOMESS"))
End If
If GENERAREST% > 0 Then
  SQLX$ = "INSERT INTO  ENCAPRFUN (ID_LISTA ,LISTA, Descripcion, ADDREGISTROS, USAFORMULA, BORRAR, INSERTAR, NOTILDE "
  SQLX$ = SQLX$ + ", CANTICAMPOS, CodiEmpr, USARECORD, FECHEMISI, TITURECORD, ConsultaSql)"
  SQLX$ = SQLX$ + "VALUES ( " & IDLISTAX&
  SQLX$ = SQLX$ + ",'LISTABCONTA'"
  SQLX$ = SQLX$ + ",'LISTA DE CUENTAS CONTABLES'"
  SQLX$ = SQLX$ + ",1"
  SQLX$ = SQLX$ + ",1"
  SQLX$ = SQLX$ + " ,1"
  SQLX$ = SQLX$ + " ,1"
  SQLX$ = SQLX$ + " ,0"
  SQLX$ = SQLX$ + " ,9"
  SQLX$ = SQLX$ + " ,0"
  SQLX$ = SQLX$ + " ,0"
  SQLX$ = SQLX$ + " ,'10/07/19'"
  SQLX$ = SQLX$ + " ,''"
  SQLX$ = SQLX$ + " ,'SELECT CAMPO1, CAMPO2,CAMPO3,CAMPO4,CAMPO5,CAMPO6,CAMPO7,CAMPO8  FROM DATASQL WITH(NOLOCK) '"
  SQLX$ = SQLX$ + ")"
  FLEXCONN.Execute (SQLX$)
    SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
    SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
    SQLX$ = SQLX$ + " ,0"
    SQLX$ = SQLX$ + " ,0"
    SQLX$ = SQLX$ + " , 'NV/A0'"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ' '"
    SQLX$ = SQLX$ + " , ' '"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " )"
    FLEXCONN.Execute (SQLX$)

'
    SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
    SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
    SQLX$ = SQLX$ + " ,0"
    SQLX$ = SQLX$ + " ,0"
    SQLX$ = SQLX$ + " , 'CODIGO/A20'"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , 'S'"
'    SQLX$ = SQLX$ + " , 'CAMPO1;CAMPO2;DATASQL;LISTA ='REFEINGA''"
    SQLX$ = SQLX$ + " , " & "'CAMPO1;CAMPO2;DATASQL;LISTA ='" & "'REFEINGA'" & "''"
    
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " )"
    FLEXCONN.Execute (SQLX$)

'
    SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
    SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
    SQLX$ = SQLX$ + " ,0"
    SQLX$ = SQLX$ + " ,0"
    SQLX$ = SQLX$ + " , 'DENOMINACIÓN/A32'"
    SQLX$ = SQLX$ + " , ' '"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " )"
    FLEXCONN.Execute (SQLX$)

'
    SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
    SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
    SQLX$ = SQLX$ + " ,0"
    SQLX$ = SQLX$ + " ,0"
    SQLX$ = SQLX$ + " , 'CODIGO/A20'"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , 'S'"
    SQLX$ = SQLX$ + " , 'codexcue;RTRIM(denocue);PLANCUEN'"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " )"
    FLEXCONN.Execute (SQLX$)

'
    SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
    SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
    SQLX$ = SQLX$ + " ,0"
    SQLX$ = SQLX$ + " ,0"
    SQLX$ = SQLX$ + " , 'DENOMINACIÓN/A32'"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " )"
    FLEXCONN.Execute (SQLX$)

'
    SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
    SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
    SQLX$ = SQLX$ + " ,0"
    SQLX$ = SQLX$ + " ,0"
    SQLX$ = SQLX$ + " , 'CODIGO/A24'"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , 'S'"
    SQLX$ = SQLX$ + " , 'codexcue;RTRIM(denocue);PLANCUEN'"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " )"
    FLEXCONN.Execute (SQLX$)

'
    SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
    SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
    SQLX$ = SQLX$ + " ,0"
    SQLX$ = SQLX$ + " ,0"
    SQLX$ = SQLX$ + " , 'DENOMINACIÓN/A32'"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ' '"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " )"
    FLEXCONN.Execute (SQLX$)

'
    SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
    SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
    SQLX$ = SQLX$ + " ,0"
    SQLX$ = SQLX$ + " ,0"
    SQLX$ = SQLX$ + " , 'NV/A0'"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " )"
    FLEXCONN.Execute (SQLX$)

'
End If
'|FIN LISTABCONTA



'|INICIO REFEINGA
'
CONDI$ = " LISTA =   'REFEINGA'"
EXILISTA% = CantRegistros&("ENCAPRFUN", CONDI$, "NOMESS")
If EXILISTA% > 0 Then
  FE% = 16522
  FEBD% = CVINT(VALOBADA("ENCAPRFUN", "FECHEMISI", CONDI$, "NOMESS"))
  GENERAREST% = 0
  If FEBD% < FE% Then
    IDLISTAX& = XVALO(VALOBADA("ENCAPRFUN", "ID_LISTA", CONDI$, "NOMESS"))
    Call BORRAREGISTROS("ENCAPRFUN", "ID_LISTA = " & IDLISTAX&, REGAF&, "NOMESS")
    Call BORRAREGISTROS("ENCAPRF", "ID_LISTA = " & IDLISTAX&, REGAF&, "NOMESS")
    GENERAREST% = 1
  End If
Else
  GENERAREST% = 1
  CONDI$ = " ID_LISTA  > 0 "
  IDLISTAX& = 1 + XVALO(VALOBADA("ENCAPRFUN", "MAX(ID_LISTA)", CONDI$, "NOMESS"))
End If
If GENERAREST% > 0 Then
  SQLX$ = "INSERT INTO  ENCAPRFUN (ID_LISTA ,LISTA, Descripcion, ADDREGISTROS, USAFORMULA, BORRAR, INSERTAR, NOTILDE "
  SQLX$ = SQLX$ + ", CANTICAMPOS, CodiEmpr, USARECORD, FECHEMISI, TITURECORD, ConsultaSql)"
  SQLX$ = SQLX$ + "VALUES ( " & IDLISTAX&
  SQLX$ = SQLX$ + ",'REFEINGA'"
  SQLX$ = SQLX$ + ",'REFERENCIAS DE INGRESOS Y GASTOS CASHFLOW Y BANCOS DE IMPACTO'"
  SQLX$ = SQLX$ + ",1"
  SQLX$ = SQLX$ + ",1"
  SQLX$ = SQLX$ + " ,0"
  SQLX$ = SQLX$ + " ,0"
  SQLX$ = SQLX$ + " ,0"
  SQLX$ = SQLX$ + " ,4"
  SQLX$ = SQLX$ + " ,0"
  SQLX$ = SQLX$ + " ,0"
  SQLX$ = SQLX$ + " ,'10/09/19'"
  SQLX$ = SQLX$ + " ,''"
  SQLX$ = SQLX$ + " ,'SELECT CAMPO1,CAMPO2,CAMPO3,CAMPO4,CAMPO5 FROM DATASQL WITH(NOLOCK)'"
  SQLX$ = SQLX$ + ")"
  FLEXCONN.Execute (SQLX$)
    SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
    SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
    SQLX$ = SQLX$ + " ,0"
    SQLX$ = SQLX$ + " ,0"
    SQLX$ = SQLX$ + " , 'CODIGO P.I.G/F8'"
    SQLX$ = SQLX$ + " , ' '"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " )"
    FLEXCONN.Execute (SQLX$)

'
    SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
    SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
    SQLX$ = SQLX$ + " ,0"
    SQLX$ = SQLX$ + " ,0"
    SQLX$ = SQLX$ + " , 'REF. PLAN I.G./A48'"
    SQLX$ = SQLX$ + " , ' '"
    SQLX$ = SQLX$ + " , 'S'"
    SQLX$ = SQLX$ + " , ' '"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " )"
    FLEXCONN.Execute (SQLX$)

    SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
    SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
    SQLX$ = SQLX$ + " ,0"
    SQLX$ = SQLX$ + " ,0"
    SQLX$ = SQLX$ + " , 'CTA.CBLE./A16'"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , 'S'"
    SQLX$ = SQLX$ + " , " & "'CAMPO1;CAMPO2;DATASQL;LISTA ='" & "'CUENTASCONBLES'" & "''"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ' '"
    SQLX$ = SQLX$ + " )"
    FLEXCONN.Execute (SQLX$)

'
    SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
    SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
    SQLX$ = SQLX$ + " ,0"
    SQLX$ = SQLX$ + " ,0"
    SQLX$ = SQLX$ + " , 'REFERENCIA/A40'"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ' '"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " )"
    FLEXCONN.Execute (SQLX$)

'
    SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
    SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
    SQLX$ = SQLX$ + " ,0"
    SQLX$ = SQLX$ + " ,0"
    SQLX$ = SQLX$ + " , 'NV/A0'"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " )"
    FLEXCONN.Execute (SQLX$)

'
End If
'|FIN REFEINGA

'|INICIO INGRSALI
'
CONDI$ = " LISTA =   'INGRSALI'"
EXILISTA% = CantRegistros&("ENCAPRFUN", CONDI$, "NOMESS")
If EXILISTA% > 0 Then
  FE% = 16935
  FEBD% = CVINT(VALOBADA("ENCAPRFUN", "FECHEMISI", CONDI$, "NOMESS"))
  GENERAREST% = 0
  If FEBD% < FE% Then
    IDLISTAX& = XVALO(VALOBADA("ENCAPRFUN", "ID_LISTA", CONDI$, "NOMESS"))
    Call BORRAREGISTROS("ENCAPRFUN", "ID_LISTA = " & IDLISTAX&, REGAF&, "NOMESS")
    Call BORRAREGISTROS("ENCAPRF", "ID_LISTA = " & IDLISTAX&, REGAF&, "NOMESS")
    GENERAREST% = 1
  End If
Else
  GENERAREST% = 1
  CONDI$ = " ID_LISTA  > 0 "
  IDLISTAX& = 1 + XVALO(VALOBADA("ENCAPRFUN", "MAX(ID_LISTA)", CONDI$, "NOMESS"))
End If
If GENERAREST% > 0 Then
  SQLX$ = "INSERT INTO  ENCAPRFUN (ID_LISTA ,LISTA, Descripcion, ADDREGISTROS, USAFORMULA, BORRAR, INSERTAR, NOTILDE "
  SQLX$ = SQLX$ + ", CANTICAMPOS, CodiEmpr, USARECORD, FECHEMISI, TITURECORD, ConsultaSql)"
  SQLX$ = SQLX$ + "VALUES ( " & IDLISTAX&
  SQLX$ = SQLX$ + ",'INGRSALI'"
  SQLX$ = SQLX$ + ",'INGRESOS Y SALIDAS CASHFLOW'"
  SQLX$ = SQLX$ + ",1"
  SQLX$ = SQLX$ + ",1"
  SQLX$ = SQLX$ + " ,1"
  SQLX$ = SQLX$ + " ,1"
  SQLX$ = SQLX$ + " ,0"
  SQLX$ = SQLX$ + " ,9"
  SQLX$ = SQLX$ + " ,0"
  SQLX$ = SQLX$ + " ,0"
  SQLX$ = SQLX$ + " ,'07/09/20'"
  SQLX$ = SQLX$ + " ,''"
  SQLX$ = SQLX$ + " ,'SELECT CAMPOFE1,CAMPO2,CAMPO3,CAMPO4,CAMPO5,CAMPO6,CAMPO7 FROM DATASQL WITH(NOLOCK)'"
'
  SQLX$ = SQLX$ + ")"
  FLEXCONN.Execute (SQLX$)
    SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
    SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
    SQLX$ = SQLX$ + " ,0"
    SQLX$ = SQLX$ + " ,0"
    SQLX$ = SQLX$ + " , 'Fecha/D8'"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , 'S'"
    SQLX$ = SQLX$ + " , ' '"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " )"
    FLEXCONN.Execute (SQLX$)

'
    SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
    SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
    SQLX$ = SQLX$ + " ,0"
    SQLX$ = SQLX$ + " ,0"
    SQLX$ = SQLX$ + " , 'COD.ING.SAL./F12'"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , 'S'"
    SQLX$ = SQLX$ + " , " & "'CAMPO1;CAMPO2;DATASQL;LISTA ='" & "'REFEINGA'" & "''"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " )"
    FLEXCONN.Execute (SQLX$)

'
    SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
    SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
    SQLX$ = SQLX$ + " ,0"
    SQLX$ = SQLX$ + " ,0"
    SQLX$ = SQLX$ + " , 'REFERENCIA/A32'"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ' '"
    SQLX$ = SQLX$ + " , ' '"
    SQLX$ = SQLX$ + " )"
    FLEXCONN.Execute (SQLX$)

'
    SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
    SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
    SQLX$ = SQLX$ + " ,0"
    SQLX$ = SQLX$ + " ,0"
    SQLX$ = SQLX$ + " , 'INGRESO/F14.2'"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , 'S'"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " )"
    FLEXCONN.Execute (SQLX$)

'
    SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
    SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
    SQLX$ = SQLX$ + " ,0"
    SQLX$ = SQLX$ + " ,0"
    SQLX$ = SQLX$ + " , 'SALIDA/F14.2'"
    SQLX$ = SQLX$ + " , ' '"
    SQLX$ = SQLX$ + " , 'S'"
    SQLX$ = SQLX$ + " , ' '"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " )"
    FLEXCONN.Execute (SQLX$)

'
    SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
    SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
    SQLX$ = SQLX$ + " ,0"
    SQLX$ = SQLX$ + " ,0"
    SQLX$ = SQLX$ + " , 'OBSERVACION/A48'"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , 'S'"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " )"
    FLEXCONN.Execute (SQLX$)

'
    SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
    SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
    SQLX$ = SQLX$ + " ,0"
    SQLX$ = SQLX$ + " ,0"
    SQLX$ = SQLX$ + " , 'NV/A0'"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " )"
    FLEXCONN.Execute (SQLX$)

'
End If
'|FIN INGRSALI


'|INICIO VENCIMIVA
'
CONDI$ = " LISTA =   'VENCIMIVA'"
EXILISTA% = CantRegistros&("ENCAPRFUN", CONDI$, "NOMESS")
If EXILISTA% > 0 Then
  FE% = 16490
  FEBD% = CVINT(VALOBADA("ENCAPRFUN", "FECHEMISI", CONDI$, "NOMESS"))
  GENERAREST% = 0
  If FEBD% < FE% Then
    IDLISTAX& = XVALO(VALOBADA("ENCAPRFUN", "ID_LISTA", CONDI$, "NOMESS"))
    Call BORRAREGISTROS("ENCAPRFUN", "ID_LISTA = " & IDLISTAX&, REGAF&, "NOMESS")
    Call BORRAREGISTROS("ENCAPRF", "ID_LISTA = " & IDLISTAX&, REGAF&, "NOMESS")
    GENERAREST% = 1
  End If
Else
  GENERAREST% = 1
  CONDI$ = " ID_LISTA  > 0 "
  IDLISTAX& = 1 + XVALO(VALOBADA("ENCAPRFUN", "MAX(ID_LISTA)", CONDI$, "NOMESS"))
End If
If GENERAREST% > 0 Then
  SQLX$ = "INSERT INTO  ENCAPRFUN (ID_LISTA ,LISTA, Descripcion, ADDREGISTROS, USAFORMULA, BORRAR, INSERTAR, NOTILDE "
  SQLX$ = SQLX$ + ", CANTICAMPOS, CodiEmpr, USARECORD, FECHEMISI, TITURECORD, ConsultaSql)"
  SQLX$ = SQLX$ + "VALUES ( " & IDLISTAX&
  SQLX$ = SQLX$ + ",'VENCIMIVA'"
  SQLX$ = SQLX$ + ",'VENCIMIENTOS DE PAGOS DE IMPUESTOS'"
  SQLX$ = SQLX$ + ",1"
  SQLX$ = SQLX$ + ",1"
  SQLX$ = SQLX$ + " ,1"
  SQLX$ = SQLX$ + " ,1"
  SQLX$ = SQLX$ + " ,0"
  SQLX$ = SQLX$ + " ,7"
  SQLX$ = SQLX$ + " ,0"
  SQLX$ = SQLX$ + " ,0"
  SQLX$ = SQLX$ + " ,'17/09/19'"
  SQLX$ = SQLX$ + " ,''"
  SQLX$ = SQLX$ + " ,'SELECT CAMPO1,CAMPO2,CAMPO3,CAMPO4,CAMPO5,CAMPO6 FROM DATASQL WITH(NOLOCK)'"
  SQLX$ = SQLX$ + ")"
  FLEXCONN.Execute (SQLX$)
    SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
    SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
    SQLX$ = SQLX$ + " ,0"
    SQLX$ = SQLX$ + " ,0"
    SQLX$ = SQLX$ + " , 'Fecha/a6'"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , 'S'"
    SQLX$ = SQLX$ + " , ' '"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " )"
    FLEXCONN.Execute (SQLX$)

'
    SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
    SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
    SQLX$ = SQLX$ + " ,0"
    SQLX$ = SQLX$ + " ,0"
    SQLX$ = SQLX$ + " , 'COD.ING.SAL./F12'"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , 'S'"
    SQLX$ = SQLX$ + " , " & "'CAMPO1;CAMPO2;DATASQL;LISTA ='" & "'REFEINGA'" & "''"

    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " )"
    FLEXCONN.Execute (SQLX$)

'
    SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
    SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
    SQLX$ = SQLX$ + " ,0"
    SQLX$ = SQLX$ + " ,0"
    SQLX$ = SQLX$ + " , 'REFERENCIA/A32'"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " )"
    FLEXCONN.Execute (SQLX$)

    SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
    SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
    SQLX$ = SQLX$ + " ,0"
    SQLX$ = SQLX$ + " ,0"
    SQLX$ = SQLX$ + " , 'CTA.CBLE./A18'"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , 'S'"
    SQLX$ = SQLX$ + " , " & "'CAMPO1;CAMPO2;DATASQL;LISTA ='" & "'CUENTASCONBLES'" & "''"

    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " )"
    FLEXCONN.Execute (SQLX$)

'
    SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
    SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
    SQLX$ = SQLX$ + " ,0"
    SQLX$ = SQLX$ + " ,0"
    SQLX$ = SQLX$ + " , 'REFERENCIA/A32'"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ' '"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " )"
    FLEXCONN.Execute (SQLX$)


'
    SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
    SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
    SQLX$ = SQLX$ + " ,0"
    SQLX$ = SQLX$ + " ,0"
    SQLX$ = SQLX$ + " , 'NV/A0'"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " )"
    FLEXCONN.Execute (SQLX$)

'
End If
'|FIN VENCIMIVA

'|INICIO CUENBANTA
'
CONDI$ = " LISTA =   'CUENBANTA'"
EXILISTA% = CantRegistros&("ENCAPRFUN", CONDI$, "NOMESS")
If EXILISTA% > 0 Then
  FE% = 16451
  FEBD% = CVINT(VALOBADA("ENCAPRFUN", "FECHEMISI", CONDI$, "NOMESS"))
  GENERAREST% = 0
  If FEBD% < FE% Then
    IDLISTAX& = XVALO(VALOBADA("ENCAPRFUN", "ID_LISTA", CONDI$, "NOMESS"))
    Call BORRAREGISTROS("ENCAPRFUN", "ID_LISTA = " & IDLISTAX&, REGAF&, "NOMESS")
    Call BORRAREGISTROS("ENCAPRF", "ID_LISTA = " & IDLISTAX&, REGAF&, "NOMESS")
    GENERAREST% = 1
  End If
Else
  GENERAREST% = 1
  CONDI$ = " ID_LISTA  > 0 "
  IDLISTAX& = 1 + XVALO(VALOBADA("ENCAPRFUN", "MAX(ID_LISTA)", CONDI$, "NOMESS"))
End If
If GENERAREST% > 0 Then
  SQLX$ = "INSERT INTO  ENCAPRFUN (ID_LISTA ,LISTA, Descripcion, ADDREGISTROS, USAFORMULA, BORRAR, INSERTAR, NOTILDE "
  SQLX$ = SQLX$ + ", CANTICAMPOS, CodiEmpr, USARECORD, FECHEMISI, TITURECORD, ConsultaSql)"
  SQLX$ = SQLX$ + "VALUES ( " & IDLISTAX&
  SQLX$ = SQLX$ + ",'CUENBANTA'"
  SQLX$ = SQLX$ + ",'CUENTAS BANCARIAS  SQL'"
  SQLX$ = SQLX$ + ",1"
  SQLX$ = SQLX$ + ",1"
  SQLX$ = SQLX$ + " ,0"
  SQLX$ = SQLX$ + " ,0"
  SQLX$ = SQLX$ + " ,0"
  SQLX$ = SQLX$ + " ,6"
  SQLX$ = SQLX$ + " ,0"
  SQLX$ = SQLX$ + " ,0"
  SQLX$ = SQLX$ + " ,'03/07/19'"
  SQLX$ = SQLX$ + " ,''"
  SQLX$ = SQLX$ + " ,'SELECT CAMPO1,CAMPO2,CAMPO3,CAMPO4,CAMPO5 FROM DATASQL WITH(NOLOCK)'"
  SQLX$ = SQLX$ + ")"
  FLEXCONN.Execute (SQLX$)
    SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
    SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
    SQLX$ = SQLX$ + " ,0"
    SQLX$ = SQLX$ + " ,0"
    SQLX$ = SQLX$ + " , 'CTA.CBLE/A16'"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , 'S'"
    SQLX$ = SQLX$ + " , 'CODEXCUE;DENOCUE;PLANCUEN'"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " )"
    FLEXCONN.Execute (SQLX$)

'
    SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
    SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
    SQLX$ = SQLX$ + " ,0"
    SQLX$ = SQLX$ + " ,0"
    SQLX$ = SQLX$ + " , 'DESCRIPCION/A32'"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ' '"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " )"
    FLEXCONN.Execute (SQLX$)

'
    SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
    SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
    SQLX$ = SQLX$ + " ,0"
    SQLX$ = SQLX$ + " ,0"
    SQLX$ = SQLX$ + " , 'D.H.ACRED./F8'"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , 'S'"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " )"
    FLEXCONN.Execute (SQLX$)

'
    SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
    SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
    SQLX$ = SQLX$ + " ,0"
    SQLX$ = SQLX$ + " ,0"
    SQLX$ = SQLX$ + " , '% COMIS./F10.2'"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , 'S'"
    SQLX$ = SQLX$ + " , ' '"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " )"
    FLEXCONN.Execute (SQLX$)

'
    SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
    SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
    SQLX$ = SQLX$ + " ,0"
    SQLX$ = SQLX$ + " ,0"
    SQLX$ = SQLX$ + " , 'N.V/A0'"
    SQLX$ = SQLX$ + " , ' '"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " )"
    FLEXCONN.Execute (SQLX$)

'
End If
'|FIN CUENBANTA

'|INICIO SALCUBAN
'
CONDI$ = " LISTA =   'SALCUBAN'"
EXILISTA% = CantRegistros&("ENCAPRFUN", CONDI$, "NOMESS")
If EXILISTA% > 0 Then
  FE% = 16451
  FEBD% = CVINT(VALOBADA("ENCAPRFUN", "FECHEMISI", CONDI$, "NOMESS"))
  GENERAREST% = 0
  If FEBD% < FE% Then
    IDLISTAX& = XVALO(VALOBADA("ENCAPRFUN", "ID_LISTA", CONDI$, "NOMESS"))
    Call BORRAREGISTROS("ENCAPRFUN", "ID_LISTA = " & IDLISTAX&, REGAF&, "NOMESS")
    Call BORRAREGISTROS("ENCAPRF", "ID_LISTA = " & IDLISTAX&, REGAF&, "NOMESS")
    GENERAREST% = 1
  End If
Else
  GENERAREST% = 1
  CONDI$ = " ID_LISTA  > 0 "
  IDLISTAX& = 1 + XVALO(VALOBADA("ENCAPRFUN", "MAX(ID_LISTA)", CONDI$, "NOMESS"))
End If
If GENERAREST% > 0 Then
  SQLX$ = "INSERT INTO  ENCAPRFUN (ID_LISTA ,LISTA, Descripcion, ADDREGISTROS, USAFORMULA, BORRAR, INSERTAR, NOTILDE "
  SQLX$ = SQLX$ + ", CANTICAMPOS, CodiEmpr, USARECORD, FECHEMISI, TITURECORD, ConsultaSql)"
  SQLX$ = SQLX$ + "VALUES ( " & IDLISTAX&
  SQLX$ = SQLX$ + ",'SALCUBAN'"
  SQLX$ = SQLX$ + ",'SALDO DE CUENTAS BANCARIAS'"
  SQLX$ = SQLX$ + ",1"
  SQLX$ = SQLX$ + ",1"
  SQLX$ = SQLX$ + " ,0"
  SQLX$ = SQLX$ + " ,0"
  SQLX$ = SQLX$ + " ,0"
  SQLX$ = SQLX$ + " ,4"
  SQLX$ = SQLX$ + " ,0"
  SQLX$ = SQLX$ + " ,0"
  SQLX$ = SQLX$ + " ,'03/07/19'"
  SQLX$ = SQLX$ + " ,''"
  SQLX$ = SQLX$ + " ,'SELECT CAMPO1,CAMPO2,CAMPO3 FROM DATASQL WITH(NOLOCK)'"
'
  SQLX$ = SQLX$ + ")"
  FLEXCONN.Execute (SQLX$)
    SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
    SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
    SQLX$ = SQLX$ + " ,0"
    SQLX$ = SQLX$ + " ,0"
    SQLX$ = SQLX$ + " , 'CTA.CBLE.BANCO/A12'"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , 'S'"
    SQLX$ = SQLX$ + " , " & "'CAMPO1;CAMPO2;DATASQL;LISTA ='" & "'CUENBANTA'" & "''"


    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " )"
    FLEXCONN.Execute (SQLX$)

'
    SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
    SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
    SQLX$ = SQLX$ + " ,0"
    SQLX$ = SQLX$ + " ,0"
    SQLX$ = SQLX$ + " , 'BANCO/A30'"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ' '"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " )"
    FLEXCONN.Execute (SQLX$)

'
    SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
    SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
    SQLX$ = SQLX$ + " ,0"
    SQLX$ = SQLX$ + " ,0"
    SQLX$ = SQLX$ + " , 'SALDO/f14.2'"
    SQLX$ = SQLX$ + " , ' '"
    SQLX$ = SQLX$ + " , 'S'"
    SQLX$ = SQLX$ + " , ' '"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " )"
    FLEXCONN.Execute (SQLX$)

'
End If
'|FIN SALCUBAN


'|INICIO CONFIMECO
'
CONDI$ = " LISTA =   'CONFIMECO'"
EXILISTA% = CantRegistros&("ENCAPRFUN", CONDI$, "NOMESS")
If EXILISTA% > 0 Then
  FE% = 16530
  FEBD% = CVINT(VALOBADA("ENCAPRFUN", "FECHEMISI", CONDI$, "NOMESS"))
  GENERAREST% = 0
  If FEBD% < FE% Then
    IDLISTAX& = XVALO(VALOBADA("ENCAPRFUN", "ID_LISTA", CONDI$, "NOMESS"))
    Call BORRAREGISTROS("ENCAPRFUN", "ID_LISTA = " & IDLISTAX&, REGAF&, "NOMESS")
    Call BORRAREGISTROS("ENCAPRF", "ID_LISTA = " & IDLISTAX&, REGAF&, "NOMESS")
    GENERAREST% = 1
  End If
Else
  GENERAREST% = 1
  CONDI$ = " ID_LISTA  > 0 "
  IDLISTAX& = 1 + XVALO(VALOBADA("ENCAPRFUN", "MAX(ID_LISTA)", CONDI$, "NOMESS"))
End If
If GENERAREST% > 0 Then
  SQLX$ = "INSERT INTO  ENCAPRFUN (ID_LISTA ,LISTA, Descripcion, ADDREGISTROS, USAFORMULA, BORRAR, INSERTAR, NOTILDE "
  SQLX$ = SQLX$ + ", CANTICAMPOS, CodiEmpr, USARECORD, FECHEMISI, TITURECORD, ConsultaSql)"
  SQLX$ = SQLX$ + "VALUES ( " & IDLISTAX&
  SQLX$ = SQLX$ + ",'CONFIMECO'"
  SQLX$ = SQLX$ + ",'CONFIGURACION DE ACREDITACION DE MEDIOS DE COBRANZA'"
  SQLX$ = SQLX$ + ",1"
  SQLX$ = SQLX$ + ",1"
  SQLX$ = SQLX$ + " ,0"
  SQLX$ = SQLX$ + " ,0"
  SQLX$ = SQLX$ + " ,0"
  SQLX$ = SQLX$ + " ,6"
  SQLX$ = SQLX$ + " ,0"
  SQLX$ = SQLX$ + " ,0"
  SQLX$ = SQLX$ + " ,'18/09/19'"
  SQLX$ = SQLX$ + " ,''"
  SQLX$ = SQLX$ + " ,'SELECT CAMPO1,CAMPO2,CAMPO3,CAMPO4,CAMPO5,CAMPO6,CAMPO7 FROM DATASQL WITH(NOLOCK)'"
  SQLX$ = SQLX$ + ")"
  FLEXCONN.Execute (SQLX$)
    SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
    SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
    SQLX$ = SQLX$ + " ,0"
    SQLX$ = SQLX$ + " ,0"
    SQLX$ = SQLX$ + " , 'COD.MEC./A16'"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , 'S'"
    SQLX$ = SQLX$ + " , " & "'CAMPO3;CAMPO2;DATASQL;LISTA ='" & "'CTASCBMC'" & "''"

    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " )"
    FLEXCONN.Execute (SQLX$)

'
    SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
    SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
    SQLX$ = SQLX$ + " ,0"
    SQLX$ = SQLX$ + " ,0"
    SQLX$ = SQLX$ + " , 'REFERENCIA/A64'"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ' '"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " )"
    FLEXCONN.Execute (SQLX$)

'
    SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
    SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
    SQLX$ = SQLX$ + " ,0"
    SQLX$ = SQLX$ + " ,0"
    SQLX$ = SQLX$ + " , 'D.H.ACRED./F9'"
    SQLX$ = SQLX$ + " , ' '"
    SQLX$ = SQLX$ + " , 'S'"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " )"
    FLEXCONN.Execute (SQLX$)

'
    SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
    SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
    SQLX$ = SQLX$ + " ,0"
    SQLX$ = SQLX$ + " ,0"
    SQLX$ = SQLX$ + " , '% COM./F6.2'"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , 'S'"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " )"
    FLEXCONN.Execute (SQLX$)

    SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
    SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
    SQLX$ = SQLX$ + " ,0"
    SQLX$ = SQLX$ + " ,0"
    SQLX$ = SQLX$ + " , 'CTA.CBLE./A16'"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , 'S'"
    SQLX$ = SQLX$ + " , " & "'CAMPO1;CAMPO2;DATASQL;LISTA ='" & "'CUENTASCONBLES'" & "''"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ' '"
    SQLX$ = SQLX$ + " )"
    FLEXCONN.Execute (SQLX$)

'
    SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
    SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
    SQLX$ = SQLX$ + " ,0"
    SQLX$ = SQLX$ + " ,0"
    SQLX$ = SQLX$ + " , 'REFERENCIA/A40'"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ' '"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " )"
    FLEXCONN.Execute (SQLX$)
'
    SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
    SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
    SQLX$ = SQLX$ + " ,0"
    SQLX$ = SQLX$ + " ,0"
    SQLX$ = SQLX$ + " , 'N.V./A0'"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " )"
    FLEXCONN.Execute (SQLX$)

'
End If
'|FIN CONFIMECO


'|INICIO CTASCBMC
'
CONDI$ = " LISTA =   'CTASCBMC'"
EXILISTA% = CantRegistros&("ENCAPRFUN", CONDI$, "NOMESS")
If EXILISTA% > 0 Then
  FE% = 16456
  FEBD% = CVINT(VALOBADA("ENCAPRFUN", "FECHEMISI", CONDI$, "NOMESS"))
  GENERAREST% = 0
  If FEBD% < FE% Then
    IDLISTAX& = XVALO(VALOBADA("ENCAPRFUN", "ID_LISTA", CONDI$, "NOMESS"))
    Call BORRAREGISTROS("ENCAPRFUN", "ID_LISTA = " & IDLISTAX&, REGAF&, "NOMESS")
    Call BORRAREGISTROS("ENCAPRF", "ID_LISTA = " & IDLISTAX&, REGAF&, "NOMESS")
    GENERAREST% = 1
  End If
Else
  GENERAREST% = 1
  CONDI$ = " ID_LISTA  > 0 "
  IDLISTAX& = 1 + XVALO(VALOBADA("ENCAPRFUN", "MAX(ID_LISTA)", CONDI$, "NOMESS"))
End If
If GENERAREST% > 0 Then
  SQLX$ = "INSERT INTO  ENCAPRFUN (ID_LISTA ,LISTA, Descripcion, ADDREGISTROS, USAFORMULA, BORRAR, INSERTAR, NOTILDE "
  SQLX$ = SQLX$ + ", CANTICAMPOS, CodiEmpr, USARECORD, FECHEMISI, TITURECORD, ConsultaSql)"
  SQLX$ = SQLX$ + "VALUES ( " & IDLISTAX&
  SQLX$ = SQLX$ + ",'CTASCBMC'"
  SQLX$ = SQLX$ + ",'CUENTAS CONTABLES DE MEDIOS DE COBRANZA'"
  SQLX$ = SQLX$ + ",0"
  SQLX$ = SQLX$ + ",1"
  SQLX$ = SQLX$ + " ,0"
  SQLX$ = SQLX$ + " ,0"
  SQLX$ = SQLX$ + " ,0"
  SQLX$ = SQLX$ + " ,6"
  SQLX$ = SQLX$ + " ,0"
  SQLX$ = SQLX$ + " ,0"
  SQLX$ = SQLX$ + " ,'08/07/19'"
  SQLX$ = SQLX$ + " ,''"
  SQLX$ = SQLX$ + " ,'SELECT CAMPO1,CAMPO2,CAMPO3,CAMPO4,CAMPO5 FROM DATASQL WITH(NOLOCK)'"

'
  SQLX$ = SQLX$ + ")"
  FLEXCONN.Execute (SQLX$)
    SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
    SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
    SQLX$ = SQLX$ + " ,0"
    SQLX$ = SQLX$ + " ,0"
    SQLX$ = SQLX$ + " , 'COD.COB/F8'"
    SQLX$ = SQLX$ + " , ' '"
    SQLX$ = SQLX$ + " , ' '"
    SQLX$ = SQLX$ + " , ' '"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " )"
    FLEXCONN.Execute (SQLX$)

'
    SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
    SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
    SQLX$ = SQLX$ + " ,0"
    SQLX$ = SQLX$ + " ,0"
    SQLX$ = SQLX$ + " , 'REFERENCIA/A32'"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ' '"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " )"
    FLEXCONN.Execute (SQLX$)

'
    SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
    SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
    SQLX$ = SQLX$ + " ,0"
    SQLX$ = SQLX$ + " ,0"
    SQLX$ = SQLX$ + " , 'CTACBLE/A14'"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ' '"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " )"
    FLEXCONN.Execute (SQLX$)

'
    SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
    SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
    SQLX$ = SQLX$ + " ,0"
    SQLX$ = SQLX$ + " ,0"
    SQLX$ = SQLX$ + " , 'CODMOV/A8'"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " )"
    FLEXCONN.Execute (SQLX$)

'
    SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
    SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
    SQLX$ = SQLX$ + " ,0"
    SQLX$ = SQLX$ + " ,0"
    SQLX$ = SQLX$ + " , ' N.V./A0'"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " )"
    FLEXCONN.Execute (SQLX$)

'
End If
'|FIN CTASCBMC

'|INICIO TABDASQL
'
CONDI$ = " LISTA =   'TABDASQL'"
EXILISTA% = CantRegistros&("ENCAPRFUN", CONDI$, "NOMESS")
If EXILISTA% > 0 Then
  FE% = 16470
  FEBD% = CVINT(VALOBADA("ENCAPRFUN", "FECHEMISI", CONDI$, "NOMESS"))
  GENERAREST% = 0
  If FEBD% < FE% Then
    IDLISTAX& = XVALO(VALOBADA("ENCAPRFUN", "ID_LISTA", CONDI$, "NOMESS"))
    Call BORRAREGISTROS("ENCAPRFUN", "ID_LISTA = " & IDLISTAX&, REGAF&, "NOMESS")
    Call BORRAREGISTROS("ENCAPRF", "ID_LISTA = " & IDLISTAX&, REGAF&, "NOMESS")
    GENERAREST% = 1
  End If
Else
  GENERAREST% = 1
  CONDI$ = " ID_LISTA  > 0 "
  IDLISTAX& = 1 + XVALO(VALOBADA("ENCAPRFUN", "MAX(ID_LISTA)", CONDI$, "NOMESS"))
End If
If GENERAREST% > 0 Then
  SQLX$ = "INSERT INTO  ENCAPRFUN (ID_LISTA ,LISTA, Descripcion, ADDREGISTROS, USAFORMULA, BORRAR, INSERTAR, NOTILDE "
  SQLX$ = SQLX$ + ", CANTICAMPOS, CodiEmpr, USARECORD, FECHEMISI, TITURECORD, ConsultaSql)"
  SQLX$ = SQLX$ + "VALUES ( " & IDLISTAX&
  SQLX$ = SQLX$ + ",'TABDASQL'"
  SQLX$ = SQLX$ + ",'LISTA DE TABLAS DATASQL'"
  SQLX$ = SQLX$ + ",1"
  SQLX$ = SQLX$ + ",1"
  SQLX$ = SQLX$ + " ,0"
  SQLX$ = SQLX$ + " ,1"
  SQLX$ = SQLX$ + " ,0"
  SQLX$ = SQLX$ + " ,5"
  SQLX$ = SQLX$ + " ,0"
  SQLX$ = SQLX$ + " ,0"
  SQLX$ = SQLX$ + " ,'22/07/19'"
  SQLX$ = SQLX$ + " ,''"
  SQLX$ = SQLX$ + " ,'SELECT CAMPO1,CAMPO2, CAMPO3, CAMPO4 FROM DATASQL WITH(NOLOCK)'"

'
  SQLX$ = SQLX$ + ")"
  FLEXCONN.Execute (SQLX$)
    SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
    SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
    SQLX$ = SQLX$ + " ,0"
    SQLX$ = SQLX$ + " ,0"
    SQLX$ = SQLX$ + " , 'ID-TABLA/F8'"
    SQLX$ = SQLX$ + " , ' '"
    SQLX$ = SQLX$ + " , 'S'"
    SQLX$ = SQLX$ + " , 'ID_LISTA;LISTA;ENCAPRFUN'"
    SQLX$ = SQLX$ + " , ' '"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " )"
    FLEXCONN.Execute (SQLX$)

'
    SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
    SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
    SQLX$ = SQLX$ + " ,0"
    SQLX$ = SQLX$ + " ,0"
    SQLX$ = SQLX$ + " , 'REFERENCIA/A48'"
    SQLX$ = SQLX$ + " , ' '"
    SQLX$ = SQLX$ + " , ' '"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " )"
    FLEXCONN.Execute (SQLX$)

'
    SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
    SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
    SQLX$ = SQLX$ + " ,0"
    SQLX$ = SQLX$ + " ,0"
    SQLX$ = SQLX$ + " , 'FORMULARIO/A32'"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , 'S'"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " )"
    FLEXCONN.Execute (SQLX$)

'
    SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
    SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
    SQLX$ = SQLX$ + " ,0"
    SQLX$ = SQLX$ + " ,0"
    SQLX$ = SQLX$ + " , 'NV/A0'"
    SQLX$ = SQLX$ + " , ' '"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " )"
    FLEXCONN.Execute (SQLX$)

'
End If
'|FIN TABDASQL

'|INICIO CUENTASCONBLES
'
CONDI$ = " LISTA =   'CUENTASCONBLES'"
EXILISTA% = CantRegistros&("ENCAPRFUN", CONDI$, "NOMESS")
If EXILISTA% > 0 Then
  FE% = 16522
  FEBD% = CVINT(VALOBADA("ENCAPRFUN", "FECHEMISI", CONDI$, "NOMESS"))
  GENERAREST% = 0
  If FEBD% < FE% Then
    IDLISTAX& = XVALO(VALOBADA("ENCAPRFUN", "ID_LISTA", CONDI$, "NOMESS"))
    Call BORRAREGISTROS("ENCAPRFUN", "ID_LISTA = " & IDLISTAX&, REGAF&, "NOMESS")
    Call BORRAREGISTROS("ENCAPRF", "ID_LISTA = " & IDLISTAX&, REGAF&, "NOMESS")
    GENERAREST% = 1
  End If
Else
  GENERAREST% = 1
  CONDI$ = " ID_LISTA  > 0 "
  IDLISTAX& = 1 + XVALO(VALOBADA("ENCAPRFUN", "MAX(ID_LISTA)", CONDI$, "NOMESS"))
End If
If GENERAREST% > 0 Then
  SQLX$ = "INSERT INTO  ENCAPRFUN (ID_LISTA ,LISTA, Descripcion, ADDREGISTROS, USAFORMULA, BORRAR, INSERTAR, NOTILDE "
  SQLX$ = SQLX$ + ", CANTICAMPOS, CodiEmpr, USARECORD, FECHEMISI, TITURECORD, ConsultaSql)"
  SQLX$ = SQLX$ + "VALUES ( " & IDLISTAX&
  SQLX$ = SQLX$ + ",'CUENTASCONBLES'"
  SQLX$ = SQLX$ + ",'CUENTAS CONTABLES'"
  SQLX$ = SQLX$ + ",1"
  SQLX$ = SQLX$ + ",1"
  SQLX$ = SQLX$ + " ,0"
  SQLX$ = SQLX$ + " ,0"
  SQLX$ = SQLX$ + " ,0"
  SQLX$ = SQLX$ + " ,4"
  SQLX$ = SQLX$ + " ,0"
  SQLX$ = SQLX$ + " ,0"
  SQLX$ = SQLX$ + " ,'10/09/19'"
  SQLX$ = SQLX$ + " ,''"
  SQLX$ = SQLX$ + " ,'SELECT CAMPO1,CAMPO2,CAMPO3 FROM DATASQL WITH(NOLOCK)'"
  SQLX$ = SQLX$ + ")"
  FLEXCONN.Execute (SQLX$)
    SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
    SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
    SQLX$ = SQLX$ + " ,0"
    SQLX$ = SQLX$ + " ,0"
    SQLX$ = SQLX$ + " , 'CODIGO/A24'"
    SQLX$ = SQLX$ + " , ' '"
    SQLX$ = SQLX$ + " , 'S'"
    SQLX$ = SQLX$ + " , 'codexcue;RTRIM(denocue);PLANCUEN '"
    SQLX$ = SQLX$ + " , ' '"
    SQLX$ = SQLX$ + " , ' '"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " )"
    FLEXCONN.Execute (SQLX$)

'
    SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
    SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
    SQLX$ = SQLX$ + " ,0"
    SQLX$ = SQLX$ + " ,0"
    SQLX$ = SQLX$ + " , 'REFERENCIA/A48'"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ' '"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " )"
    FLEXCONN.Execute (SQLX$)

'
    SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
    SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
    SQLX$ = SQLX$ + " ,0"
    SQLX$ = SQLX$ + " ,0"
    SQLX$ = SQLX$ + " , 'ULTIMO/A0'"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " )"
    FLEXCONN.Execute (SQLX$)

'
End If
'|FIN CUENTASCONBLES

'|INICIO RUBFIBAN
'ASOCIACION DE BANCOS A RUBROS FIJOS DEL CASHFLOW
CONDI$ = " LISTA =   'RUBFIBAN'"
EXILISTA% = CantRegistros&("ENCAPRFUN", CONDI$, "NOMESS")
If EXILISTA% > 0 Then
  FE% = 16526
  FEBD% = CVINT(VALOBADA("ENCAPRFUN", "FECHEMISI", CONDI$, "NOMESS"))
  GENERAREST% = 0
  If FEBD% < FE% Then
    IDLISTAX& = XVALO(VALOBADA("ENCAPRFUN", "ID_LISTA", CONDI$, "NOMESS"))
    Call BORRAREGISTROS("ENCAPRFUN", "ID_LISTA = " & IDLISTAX&, REGAF&, "NOMESS")
    Call BORRAREGISTROS("ENCAPRF", "ID_LISTA = " & IDLISTAX&, REGAF&, "NOMESS")
    GENERAREST% = 1
  End If
Else
  GENERAREST% = 1
  CONDI$ = " ID_LISTA  > 0 "
  IDLISTAX& = 1 + XVALO(VALOBADA("ENCAPRFUN", "MAX(ID_LISTA)", CONDI$, "NOMESS"))
End If
If GENERAREST% > 0 Then
  SQLX$ = "INSERT INTO  ENCAPRFUN (ID_LISTA ,LISTA, Descripcion, ADDREGISTROS, USAFORMULA, BORRAR, INSERTAR, NOTILDE "
  SQLX$ = SQLX$ + ", CANTICAMPOS, CodiEmpr, USARECORD, FECHEMISI, TITURECORD, ConsultaSql)"
  SQLX$ = SQLX$ + "VALUES ( " & IDLISTAX&
  SQLX$ = SQLX$ + ",'RUBFIBAN'"
  SQLX$ = SQLX$ + ",'ASOCIACIÓN DE BANCOS  A RUBROS FIJOS '"
  SQLX$ = SQLX$ + ",1"
  SQLX$ = SQLX$ + ",1"
  SQLX$ = SQLX$ + " ,1"
  SQLX$ = SQLX$ + " ,1"
  SQLX$ = SQLX$ + " ,0"
  SQLX$ = SQLX$ + " ,5"
  SQLX$ = SQLX$ + " ,0"
  SQLX$ = SQLX$ + " ,0"
  SQLX$ = SQLX$ + " ,'13/09/19'"
  SQLX$ = SQLX$ + " ,''"
  SQLX$ = SQLX$ + " ,'SELECT CAMPO1,CAMPO2,CAMPO3, CAMPO4 FROM DATASQL WITH(NOLOCK)'"
  SQLX$ = SQLX$ + ")"
  FLEXCONN.Execute (SQLX$)
    SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
    SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
    SQLX$ = SQLX$ + " ,0"
    SQLX$ = SQLX$ + " ,0"
    SQLX$ = SQLX$ + " , 'RUBRO/A18'"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , 'S'"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " )"
    FLEXCONN.Execute (SQLX$)

'
    SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
    SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
    SQLX$ = SQLX$ + " ,0"
    SQLX$ = SQLX$ + " ,0"
    SQLX$ = SQLX$ + " , 'CTA.CBLE./A18'"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , 'S'"
    SQLX$ = SQLX$ + " , " & "'CAMPO1;CAMPO2;DATASQL;LISTA ='" & "'CUENTASCONBLES'" & "''"

    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " )"
    FLEXCONN.Execute (SQLX$)

'
    SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
    SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
    SQLX$ = SQLX$ + " ,0"
    SQLX$ = SQLX$ + " ,0"
    SQLX$ = SQLX$ + " , 'REFERENCIA/A32'"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ' '"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " )"
    FLEXCONN.Execute (SQLX$)

'
    SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
    SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
    SQLX$ = SQLX$ + " ,0"
    SQLX$ = SQLX$ + " ,0"
    SQLX$ = SQLX$ + " , 'NV/A0'"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ' '"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " )"
    FLEXCONN.Execute (SQLX$)

'
End If
'|FIN RUBFIBAN

        '|INICIO TABLASI_NO
        '
        CONDI$ = " LISTA =   'TABLASI_NO'"
        EXILISTA% = CantRegistros&("ENCAPRFUN", CONDI$, "NOMESS")
        If EXILISTA% > 0 Then
          FE% = 16326
          FEBD% = CVINT(VALOBADA("ENCAPRFUN", "FECHEMISI", CONDI$, "NOMESS"))
          GENERAREST% = 0
          If FEBD% < FE% Then
            IDLISTAX& = XVALO(VALOBADA("ENCAPRFUN", "ID_LISTA", CONDI$, "NOMESS"))
            Call BORRAREGISTROS("ENCAPRFUN", "ID_LISTA = " & IDLISTAX&, REGAF&, "NOMESS")
            Call BORRAREGISTROS("ENCAPRF", "ID_LISTA = " & IDLISTAX&, REGAF&, "NOMESS")
            GENERAREST% = 1
          End If
        Else
          GENERAREST% = 1
          CONDI$ = " ID_LISTA  > 0 "
          IDLISTAX& = 1 + XVALO(VALOBADA("ENCAPRFUN", "MAX(ID_LISTA)", CONDI$, "NOMESS"))
        End If
        If GENERAREST% > 0 Then
          SQLX$ = "INSERT INTO  ENCAPRFUN (ID_LISTA ,LISTA, Descripcion, ADDREGISTROS, USAFORMULA, BORRAR, INSERTAR, NOTILDE "
          SQLX$ = SQLX$ + ", CANTICAMPOS, CodiEmpr, USARECORD, FECHEMISI, TITURECORD, ConsultaSql)"
          SQLX$ = SQLX$ + "VALUES ( " & IDLISTAX&
          SQLX$ = SQLX$ + ",'TABLASI_NO'"
          SQLX$ = SQLX$ + ",'TABLA SI - NO     '"
          SQLX$ = SQLX$ + ",1"
          SQLX$ = SQLX$ + ",1"
          SQLX$ = SQLX$ + " ,0"
          SQLX$ = SQLX$ + " ,0"
          SQLX$ = SQLX$ + " ,0"
          SQLX$ = SQLX$ + " ,3"
          SQLX$ = SQLX$ + " ,0"
          SQLX$ = SQLX$ + " ,0"
          SQLX$ = SQLX$ + " ,'06/03/19'"
          SQLX$ = SQLX$ + " ,''"
          SQLX$ = SQLX$ + " ,'SELECT CAMPO1, CAMPO2 FROM DATASQL WITH(NOLOCK)'"
          SQLX$ = SQLX$ + ")"
          FLEXCONN.Execute (SQLX$)
            SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
            SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
            SQLX$ = SQLX$ + " ,0"
            SQLX$ = SQLX$ + " ,0"
            SQLX$ = SQLX$ + " , 'CODIGO/F4'"
            SQLX$ = SQLX$ + " , '     '"
            SQLX$ = SQLX$ + " , 'S'"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " )"
            FLEXCONN.Execute (SQLX$)
        
        '
            SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
            SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
            SQLX$ = SQLX$ + " ,0"
            SQLX$ = SQLX$ + " ,0"
            SQLX$ = SQLX$ + " , 'REFERENCIA/A24'"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " , 'S'"
            SQLX$ = SQLX$ + " , '     '"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " )"
            FLEXCONN.Execute (SQLX$)
        
        '
        End If
        '|FIN TABLASI_NO

    '|INICIO MATCONLA
    '
    CONDI$ = " LISTA =   'MATCONLA'"
    EXILISTA% = CantRegistros&("ENCAPRFUN", CONDI$, "NOMESS")
    If EXILISTA% > 0 Then
      FE% = 17162
      FEBD% = CVINT(VALOBADA("ENCAPRFUN", "FECHEMISI", CONDI$, "NOMESS"))
      GENERAREST% = 0
      If FEBD% < FE% Then
        IDLISTAX& = XVALO(VALOBADA("ENCAPRFUN", "ID_LISTA", CONDI$, "NOMESS"))
        Call BORRAREGISTROS("ENCAPRFUN", "ID_LISTA = " & IDLISTAX&, REGAF&, "NOMESS")
        Call BORRAREGISTROS("ENCAPRF", "ID_LISTA = " & IDLISTAX&, REGAF&, "NOMESS")
        GENERAREST% = 1
      End If
    Else
      GENERAREST% = 1
      CONDI$ = " ID_LISTA  > 0 "
      IDLISTAX& = 1 + XVALO(VALOBADA("ENCAPRFUN", "MAX(ID_LISTA)", CONDI$, "NOMESS"))
    End If
    If GENERAREST% > 0 Then
      SQLX$ = "INSERT INTO  ENCAPRFUN (ID_LISTA ,LISTA, Descripcion, ADDREGISTROS, USAFORMULA, BORRAR, INSERTAR, NOTILDE "
      SQLX$ = SQLX$ + ", CANTICAMPOS, CodiEmpr, USARECORD, FECHEMISI, TITURECORD, ConsultaSql)"
      SQLX$ = SQLX$ + "VALUES ( " & IDLISTAX&
      SQLX$ = SQLX$ + ",'MATCONLA'"
      SQLX$ = SQLX$ + ",'MATERIAL DE CONSUMO LAMINADO'"
      SQLX$ = SQLX$ + ",1"
      SQLX$ = SQLX$ + ",1"
      SQLX$ = SQLX$ + " ,1"
      SQLX$ = SQLX$ + " ,1"
      SQLX$ = SQLX$ + " ,0"
      SQLX$ = SQLX$ + " ,6"
      SQLX$ = SQLX$ + " ,0"
      SQLX$ = SQLX$ + " ,0"
      SQLX$ = SQLX$ + " ,'10/03/21'"
      SQLX$ = SQLX$ + " ,''"
      SQLX$ = SQLX$ + " ,'SELECT CAMPO1, CAMPO2,CAMPO3,CAMPO4,CAMPO5 FROM DATASQL WITH(NOLOCK)   '"
      SQLX$ = SQLX$ + ")"
      FLEXCONN.Execute (SQLX$)
        SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
        SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
        SQLX$ = SQLX$ + " ,0"
        SQLX$ = SQLX$ + " ,0"
        SQLX$ = SQLX$ + " , 'CODIGO/A20'"
        SQLX$ = SQLX$ + " , ''"
        SQLX$ = SQLX$ + " , 'S'"
        SQLX$ = SQLX$ + " , 'CODIGO;DESCRIPCION;MASTER;STATUS >0'"
        SQLX$ = SQLX$ + " , ''"
        SQLX$ = SQLX$ + " , ''"
        SQLX$ = SQLX$ + " , ''"
        SQLX$ = SQLX$ + " )"
        FLEXCONN.Execute (SQLX$)
    
    '
        SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
        SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
        SQLX$ = SQLX$ + " ,0"
        SQLX$ = SQLX$ + " ,0"
        SQLX$ = SQLX$ + " , 'DESCRIPCION/A40'"
        SQLX$ = SQLX$ + " , ''"
        SQLX$ = SQLX$ + " , ''"
        SQLX$ = SQLX$ + " , ''"
        SQLX$ = SQLX$ + " , ''"
        SQLX$ = SQLX$ + " , ''"
        SQLX$ = SQLX$ + " , ''"
        SQLX$ = SQLX$ + " )"
        FLEXCONN.Execute (SQLX$)
    
    '
        SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
        SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
        SQLX$ = SQLX$ + " ,0"
        SQLX$ = SQLX$ + " ,0"
        SQLX$ = SQLX$ + " , 'METROS/F10'"
        SQLX$ = SQLX$ + " , ''"
        SQLX$ = SQLX$ + " , 'S'"
        SQLX$ = SQLX$ + " , ''"
        SQLX$ = SQLX$ + " , ''"
        SQLX$ = SQLX$ + " , ''"
        SQLX$ = SQLX$ + " , ''"
        SQLX$ = SQLX$ + " )"
        FLEXCONN.Execute (SQLX$)
    
    '
        SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
        SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
        SQLX$ = SQLX$ + " ,0"
        SQLX$ = SQLX$ + " ,0"
        SQLX$ = SQLX$ + " , 'CONS. X BOB./F12.2'"
        SQLX$ = SQLX$ + " , ''"
        SQLX$ = SQLX$ + " , 'S'"
        SQLX$ = SQLX$ + " , ''"
        SQLX$ = SQLX$ + " , ''"
        SQLX$ = SQLX$ + " , ''"
        SQLX$ = SQLX$ + " , ''"
        SQLX$ = SQLX$ + " )"
        FLEXCONN.Execute (SQLX$)
    
    '
        SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
        SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
        SQLX$ = SQLX$ + " ,0"
        SQLX$ = SQLX$ + " ,0"
        SQLX$ = SQLX$ + " , 'CI/A0'"
        SQLX$ = SQLX$ + " , ''"
        SQLX$ = SQLX$ + " , ' '"
        SQLX$ = SQLX$ + " , ''"
        SQLX$ = SQLX$ + " , ''"
        SQLX$ = SQLX$ + " , ''"
        SQLX$ = SQLX$ + " , ''"
        SQLX$ = SQLX$ + " )"
        FLEXCONN.Execute (SQLX$)
    
    '
    End If
    '|FIN MATCONLA
    
'|INICIO MATETRAN
'
CONDI$ = " LISTA =   'MATETRAN'"
EXILISTA% = CantRegistros&("ENCAPRFUN", CONDI$, "NOMESS")
If EXILISTA% > 0 Then
  FE% = 17177
  FEBD% = CVINT(VALOBADA("ENCAPRFUN", "FECHEMISI", CONDI$, "NOMESS"))
  GENERAREST% = 0
  If FEBD% < FE% Then
    IDLISTAX& = XVALO(VALOBADA("ENCAPRFUN", "ID_LISTA", CONDI$, "NOMESS"))
    Call BORRAREGISTROS("ENCAPRFUN", "ID_LISTA = " & IDLISTAX&, REGAF&, "NOMESS")
    Call BORRAREGISTROS("ENCAPRF", "ID_LISTA = " & IDLISTAX&, REGAF&, "NOMESS")
    GENERAREST% = 1
  End If
Else
  GENERAREST% = 1
  CONDI$ = " ID_LISTA  > 0 "
  IDLISTAX& = 1 + XVALO(VALOBADA("ENCAPRFUN", "MAX(ID_LISTA)", CONDI$, "NOMESS"))
End If
If GENERAREST% > 0 Then
  SQLX$ = "INSERT INTO  ENCAPRFUN (ID_LISTA ,LISTA, Descripcion, ADDREGISTROS, USAFORMULA, BORRAR, INSERTAR, NOTILDE "
  SQLX$ = SQLX$ + ", CANTICAMPOS, CodiEmpr, USARECORD, FECHEMISI, TITURECORD, ConsultaSql)"
  SQLX$ = SQLX$ + "VALUES ( " & IDLISTAX&
  SQLX$ = SQLX$ + ",'MATETRAN'"
  SQLX$ = SQLX$ + ",'DEFINICION DE MAT. TRANSF. TRANSFORMADO'"
  SQLX$ = SQLX$ + ",1"
  SQLX$ = SQLX$ + ",1"
  SQLX$ = SQLX$ + " ,1"
  SQLX$ = SQLX$ + " ,1"
  SQLX$ = SQLX$ + " ,0"
  SQLX$ = SQLX$ + " ,5"
  SQLX$ = SQLX$ + " ,0"
  SQLX$ = SQLX$ + " ,0"
  SQLX$ = SQLX$ + " ,'25/03/21'"
  SQLX$ = SQLX$ + " ,''"
  SQLX$ = SQLX$ + " ,'SELECT CAMPO1, CAMPO2, CAMPO3, CAMPO4 FROM DATASQL WITH(NOLOCK) '"
  SQLX$ = SQLX$ + ")"
  FLEXCONN.Execute (SQLX$)
    SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
    SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
    SQLX$ = SQLX$ + " ,0"
    SQLX$ = SQLX$ + " ,0"
    SQLX$ = SQLX$ + " , 'COD.FAMILIA/A12'"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , 'S'"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " )"
    FLEXCONN.Execute (SQLX$)

'
    SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
    SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
    SQLX$ = SQLX$ + " ,0"
    SQLX$ = SQLX$ + " ,0"
    SQLX$ = SQLX$ + " , 'CODIGO/A20'"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , 'S'"
    SQLX$ = SQLX$ + " , 'CODIGO;DESCRIPCION;MASTER;STATUS>=0'"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " )"
    FLEXCONN.Execute (SQLX$)

'
    SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
    SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
    SQLX$ = SQLX$ + " ,0"
    SQLX$ = SQLX$ + " ,0"
    SQLX$ = SQLX$ + " , 'DESCRIPCION/A36'"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " )"
    FLEXCONN.Execute (SQLX$)

'
    SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
    SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
    SQLX$ = SQLX$ + " ,0"
    SQLX$ = SQLX$ + " ,0"
    SQLX$ = SQLX$ + " , 'DEP.NUM./F10'"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , 'S'"
    SQLX$ = SQLX$ + " , ' '"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " )"
    FLEXCONN.Execute (SQLX$)

'
End If
'|FIN MATETRAN

'|INICIO  CONFIGURACION C.F. C15
'
CONDI$ = " LISTA =   ' CONFIGURACION C.F. C15'"
EXILISTA% = CantRegistros&("ENCAPRFUN", CONDI$, "NOMESS")
If EXILISTA% > 0 Then
  FE% = 18247
  FEBD% = CVINT(VALOBADA("ENCAPRFUN", "FECHEMISI", CONDI$, "NOMESS"))
  GENERAREST% = 0
  If FEBD% < FE% Then
    IDLISTAX& = XVALO(VALOBADA("ENCAPRFUN", "ID_LISTA", CONDI$, "NOMESS"))
    Call BORRAREGISTROS("ENCAPRFUN", "ID_LISTA = " & IDLISTAX&, REGAF&, "NOMESS")
    Call BORRAREGISTROS("ENCAPRF", "ID_LISTA = " & IDLISTAX&, REGAF&, "NOMESS")
    GENERAREST% = 1
  End If
Else
  GENERAREST% = 1
  CONDI$ = " ID_LISTA  > 0 "
  IDLISTAX& = 1 + XVALO(VALOBADA("ENCAPRFUN", "MAX(ID_LISTA)", CONDI$, "NOMESS"))
End If
If GENERAREST% > 0 Then
  SQLX$ = "INSERT INTO  ENCAPRFUN (ID_LISTA ,LISTA, Descripcion, ADDREGISTROS, USAFORMULA, BORRAR, INSERTAR, NOTILDE "
  SQLX$ = SQLX$ + ", CANTICAMPOS, CodiEmpr, USARECORD, FECHEMISI, TITURECORD, ConsultaSql)"
  SQLX$ = SQLX$ + "VALUES ( " & IDLISTAX&
  SQLX$ = SQLX$ + ",' CONFIGURACION C.F. C15'"
  SQLX$ = SQLX$ + ",'  CONFIGURACION CONTROL FINAL  C15'"
  SQLX$ = SQLX$ + ",1"
  SQLX$ = SQLX$ + ",1"
  SQLX$ = SQLX$ + " ,1"
  SQLX$ = SQLX$ + " ,1"
  SQLX$ = SQLX$ + " ,0"
  SQLX$ = SQLX$ + " ,21"
  SQLX$ = SQLX$ + " ,1"
  SQLX$ = SQLX$ + " ,1"
  SQLX$ = SQLX$ + " ,'06/11/2023'"
  SQLX$ = SQLX$ + " ,''"
  SQLX$ = SQLX$ + " ,'SELECT CAMPO1,CAMPO2,CAMPO3,CAMPO4,CAMPO5,CAMPO6,CAMPO7,CAMPO8,CAMPO9, CAMPO10,CAMPO11,CAMPO12,CAMPO13,CAMPO14,CAMPO15,CAMPO16,CAMPO17,CAMPO18,CAMPO19,CAMPO20 FROM DATASQL WITH(NOLOCK)'"
  SQLX$ = SQLX$ + ")"
  FLEXCONN.Execute (SQLX$)
    SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
    SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
    SQLX$ = SQLX$ + " ,1"
    SQLX$ = SQLX$ + " ,0"
    SQLX$ = SQLX$ + " , 'DB0001/A8'"
    SQLX$ = SQLX$ + " , ' '"
    SQLX$ = SQLX$ + " , 'S'"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " )"
    FLEXCONN.Execute (SQLX$)

'
    SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
    SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
    SQLX$ = SQLX$ + " ,1"
    SQLX$ = SQLX$ + " ,0"
    SQLX$ = SQLX$ + " , 'DB0002/A8'"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , 'S'"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " )"
    FLEXCONN.Execute (SQLX$)

'
    SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
    SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
    SQLX$ = SQLX$ + " ,1"
    SQLX$ = SQLX$ + " ,0"
    SQLX$ = SQLX$ + " , 'DB0003/A8'"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , 'S'"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " )"
    FLEXCONN.Execute (SQLX$)

'
    SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
    SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
    SQLX$ = SQLX$ + " ,1"
    SQLX$ = SQLX$ + " ,0"
    SQLX$ = SQLX$ + " , 'DB0004/A8'"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , 'S'"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " )"
    FLEXCONN.Execute (SQLX$)

'
    SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
    SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
    SQLX$ = SQLX$ + " ,1"
    SQLX$ = SQLX$ + " ,0"
    SQLX$ = SQLX$ + " , 'DB0005/A8'"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , 'S'"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " )"
    FLEXCONN.Execute (SQLX$)

'
    SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
    SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
    SQLX$ = SQLX$ + " ,1"
    SQLX$ = SQLX$ + " ,0"
    SQLX$ = SQLX$ + " , 'DB0006/A8'"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , 'S'"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " )"
    FLEXCONN.Execute (SQLX$)

'
    SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
    SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
    SQLX$ = SQLX$ + " ,1"
    SQLX$ = SQLX$ + " ,0"
    SQLX$ = SQLX$ + " , 'DB0007/A8'"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , 'S'"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " )"
    FLEXCONN.Execute (SQLX$)

'
    SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
    SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
    SQLX$ = SQLX$ + " ,1"
    SQLX$ = SQLX$ + " ,0"
    SQLX$ = SQLX$ + " , 'DB0008/A8'"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , 'S'"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " )"
    FLEXCONN.Execute (SQLX$)

'
    SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
    SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
    SQLX$ = SQLX$ + " ,1"
    SQLX$ = SQLX$ + " ,0"
    SQLX$ = SQLX$ + " , 'DB0009/A8'"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , 'S'"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " )"
    FLEXCONN.Execute (SQLX$)

'
    SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
    SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
    SQLX$ = SQLX$ + " ,1"
    SQLX$ = SQLX$ + " ,0"
    SQLX$ = SQLX$ + " , 'DB0010/A8'"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , 'S'"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " )"
    FLEXCONN.Execute (SQLX$)

'
    SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
    SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
    SQLX$ = SQLX$ + " ,1"
    SQLX$ = SQLX$ + " ,0"
    SQLX$ = SQLX$ + " , 'DB0011/A8'"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , 'S'"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " )"
    FLEXCONN.Execute (SQLX$)

'
    SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
    SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
    SQLX$ = SQLX$ + " ,1"
    SQLX$ = SQLX$ + " ,0"
    SQLX$ = SQLX$ + " , 'DB0012/A8'"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , 'S'"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " )"
    FLEXCONN.Execute (SQLX$)

'
    SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
    SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
    SQLX$ = SQLX$ + " ,1"
    SQLX$ = SQLX$ + " ,0"
    SQLX$ = SQLX$ + " , 'DB0013/A8'"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , 'S'"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " )"
    FLEXCONN.Execute (SQLX$)

'
    SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
    SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
    SQLX$ = SQLX$ + " ,1"
    SQLX$ = SQLX$ + " ,0"
    SQLX$ = SQLX$ + " , 'DB0014/A8'"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , 'S'"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " )"
    FLEXCONN.Execute (SQLX$)

'
    SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
    SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
    SQLX$ = SQLX$ + " ,1"
    SQLX$ = SQLX$ + " ,0"
    SQLX$ = SQLX$ + " , 'DB0015/A8'"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , 'S'"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " )"
    FLEXCONN.Execute (SQLX$)

'
    SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
    SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
    SQLX$ = SQLX$ + " ,1"
    SQLX$ = SQLX$ + " ,0"
    SQLX$ = SQLX$ + " , 'DB0016/A8'"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , 'S'"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " )"
    FLEXCONN.Execute (SQLX$)

'
    SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
    SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
    SQLX$ = SQLX$ + " ,1"
    SQLX$ = SQLX$ + " ,0"
    SQLX$ = SQLX$ + " , 'DB0017/A8'"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , 'S'"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " )"
    FLEXCONN.Execute (SQLX$)

'
    SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
    SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
    SQLX$ = SQLX$ + " ,1"
    SQLX$ = SQLX$ + " ,0"
    SQLX$ = SQLX$ + " , 'DB0018/A8'"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , 'S'"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " )"
    FLEXCONN.Execute (SQLX$)

'
    SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
    SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
    SQLX$ = SQLX$ + " ,1"
    SQLX$ = SQLX$ + " ,0"
    SQLX$ = SQLX$ + " , 'DB0019/A8'"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , 'S'"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " )"
    FLEXCONN.Execute (SQLX$)

'
    SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
    SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
    SQLX$ = SQLX$ + " ,1"
    SQLX$ = SQLX$ + " ,0"
    SQLX$ = SQLX$ + " , 'NroRegistro/F10'"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , 'S'"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " , ''"
    SQLX$ = SQLX$ + " )"
    FLEXCONN.Execute (SQLX$)

'
End If
'|FIN  CONFIGURACION C.F. C15
'|INICIO StockMinimoNeuquen
    '
    CONDI$ = " LISTA =   'StMinNeu'"
    EXILISTA% = CantRegistros&("ENCAPRFUN", CONDI$, "NOMESS")
    If EXILISTA% > 0 Then
      FE% = 16442
      FEBD% = CVINT(VALOBADA("ENCAPRFUN", "FECHEMISI", CONDI$, "NOMESS"))
      GENERAREST% = 0
      If FEBD% < FE% Then
        IDLISTAX& = XVALO(VALOBADA("ENCAPRFUN", "ID_LISTA", CONDI$, "NOMESS"))
        Call BORRAREGISTROS("ENCAPRFUN", "ID_LISTA = " & IDLISTAX&, REGAF&, "NOMESS")
        Call BORRAREGISTROS("ENCAPRF", "ID_LISTA = " & IDLISTAX&, REGAF&, "NOMESS")
        GENERAREST% = 1
      End If
    Else
      GENERAREST% = 1
      CONDI$ = " ID_LISTA  > 0 "
      IDLISTAX& = 1 + XVALO(VALOBADA("ENCAPRFUN", "MAX(ID_LISTA)", CONDI$, "NOMESS"))
    End If
    If GENERAREST% > 0 Then
      SQLX$ = "INSERT INTO  ENCAPRFUN (ID_LISTA ,LISTA, Descripcion, ADDREGISTROS, USAFORMULA, BORRAR, INSERTAR, NOTILDE "
      SQLX$ = SQLX$ + ", CANTICAMPOS, CodiEmpr, USARECORD, FECHEMISI, TITURECORD, ConsultaSql)"
      SQLX$ = SQLX$ + "VALUES ( " & IDLISTAX&
      SQLX$ = SQLX$ + ",'StMinNeu'"
      SQLX$ = SQLX$ + ",'Stock Minimo'"
      SQLX$ = SQLX$ + ",1"
      SQLX$ = SQLX$ + ",1"
      SQLX$ = SQLX$ + " ,1"
      SQLX$ = SQLX$ + " ,1"
      SQLX$ = SQLX$ + " ,0"
      SQLX$ = SQLX$ + " ,4"
      SQLX$ = SQLX$ + " ,0"
      SQLX$ = SQLX$ + " ,0"
      SQLX$ = SQLX$ + " ,'05/01/2025'"
      SQLX$ = SQLX$ + " ,''"
      SQLX$ = SQLX$ + " ,'SELECT CAMPO1, CAMPO2, CAMPO3 FROM DATASQL WITH(NOLOCK)'"
      SQLX$ = SQLX$ + ")"
      FLEXCONN.Execute (SQLX$)
        SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
        SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
        SQLX$ = SQLX$ + " ,0"
        SQLX$ = SQLX$ + " ,0"
        SQLX$ = SQLX$ + " , 'Código/A16'"
        SQLX$ = SQLX$ + " , ''"
        SQLX$ = SQLX$ + " , 's'"
        SQLX$ = SQLX$ + " , 'codigo;descripcion;master'"
        SQLX$ = SQLX$ + " , ''"
        SQLX$ = SQLX$ + " , ''"
        SQLX$ = SQLX$ + " , ''"
        SQLX$ = SQLX$ + " )"
        FLEXCONN.Execute (SQLX$)
    
    '
        SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
        SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
        SQLX$ = SQLX$ + " ,0"
        SQLX$ = SQLX$ + " ,0"
        SQLX$ = SQLX$ + " , 'Descripción/A64'"
        SQLX$ = SQLX$ + " , ''"
        SQLX$ = SQLX$ + " , 'n'"
        SQLX$ = SQLX$ + " , ' '"
        SQLX$ = SQLX$ + " , ''"
        SQLX$ = SQLX$ + " , ''"
        SQLX$ = SQLX$ + " , ''"
        SQLX$ = SQLX$ + " )"
        FLEXCONN.Execute (SQLX$)
    
    '
        SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
        SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
        SQLX$ = SQLX$ + " ,0"
        SQLX$ = SQLX$ + " ,0"
        SQLX$ = SQLX$ + " , 'St.Min/F10.0'"
        SQLX$ = SQLX$ + " , ''"
        SQLX$ = SQLX$ + " , 's'"
        SQLX$ = SQLX$ + " , ''"
        SQLX$ = SQLX$ + " , ''"
        SQLX$ = SQLX$ + " , ''"
        SQLX$ = SQLX$ + " , ''"
        SQLX$ = SQLX$ + " )"
        FLEXCONN.Execute (SQLX$)
    
    '
    End If
    '|FIN StockMinimoNeuquen

End Sub
