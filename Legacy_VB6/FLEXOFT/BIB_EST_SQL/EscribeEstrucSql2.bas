Attribute VB_Name = "EscribeEstrucSql"

Sub GENERAR_ESTRUCTURAS_SQL()




'|INICIO USUARIOSQL
'
''CONDI$ = " LISTA =   'USUARIOSQL   '"
''EXILISTA% = CantRegistros&("ENCAPRFUN", CONDI$, "NOMESS")
''If EXILISTA% > 0 Then
''  FE% = 15662
''  FEBD% = CVINT(VALOBADA("ENCAPRFUN", "FECHEMISI", CONDI$, "NOMESS"))
''  GENERAREST% = 0
''  If FEBD% < FE% Then
''    IDLISTAX& = XVALO(VALOBADA("ENCAPRFUN", "ID_LISTA", CONDI$, "NOMESS"))
''    Call BORRAREGISTROS("ENCAPRFUN", "ID_LISTA = " & IDLISTAX&, REGAF&, "NOMESS")
''    Call BORRAREGISTROS("ENCAPRF", "ID_LISTA = " & IDLISTAX&, REGAF&, "NOMESS")
''    GENERAREST% = 1
''  End If
''Else
''  GENERAREST% = 1
''  CONDI$ = " ID_LISTA  > 0 "
''  IDLISTAX& = 1 + XVALO(VALOBADA("ENCAPRFUN", "MAX(ID_LISTA)", CONDI$, "NOMESS"))
''End If
''If GENERAREST% > 0 Then
''  SQLX$ = "INSERT INTO  ENCAPRFUN (ID_LISTA ,LISTA, Descripcion, ADDREGISTROS, USAFORMULA, BORRAR, INSERTAR, NOTILDE "
''  SQLX$ = SQLX$ + ", CANTICAMPOS, CodiEmpr, USARECORD, FECHEMISI, TITURECORD, ConsultaSql)"
''  SQLX$ = SQLX$ + "VALUES ( " & IDLISTAX&
''  SQLX$ = SQLX$ + ",'USUARIOSQL   '"
''  SQLX$ = SQLX$ + ",'TABLA DE USUARIOS'"
''  SQLX$ = SQLX$ + ",0"
''  SQLX$ = SQLX$ + ",0"
''  SQLX$ = SQLX$ + " ,0"
''  SQLX$ = SQLX$ + " ,0"
''  SQLX$ = SQLX$ + " ,0"
''  SQLX$ = SQLX$ + " ,5"
''  SQLX$ = SQLX$ + " ,0"
''  SQLX$ = SQLX$ + " ,0"
''  SQLX$ = SQLX$ + " ,'14/08/17'"
''  SQLX$ = SQLX$ + " ,''"
''  SQLX$ = SQLX$ + " ,'SELECT CAMPO1, CAMPO2, CAMPO3, CAMPO4 FROM DATASQL WITH(NOLOCK)'"
''  SQLX$ = SQLX$ + ")"
''  FLEXCONN.Execute (SQLX$)
''    SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
''    SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
''    SQLX$ = SQLX$ + " ,0"
''    SQLX$ = SQLX$ + " ,0"
''    SQLX$ = SQLX$ + " , 'NUSER/F6'"
''    SQLX$ = SQLX$ + " , '     '"
''    SQLX$ = SQLX$ + " , '     '"
''    SQLX$ = SQLX$ + " , ''"
''    SQLX$ = SQLX$ + " , ''"
''    SQLX$ = SQLX$ + " , ''"
''    SQLX$ = SQLX$ + " , ''"
''    SQLX$ = SQLX$ + " )"
''    FLEXCONN.Execute (SQLX$)
''
'''
''    SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
''    SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
''    SQLX$ = SQLX$ + " ,0"
''    SQLX$ = SQLX$ + " ,0"
''    SQLX$ = SQLX$ + " , 'NOMBRE/A32'"
''    SQLX$ = SQLX$ + " , ''"
''    SQLX$ = SQLX$ + " , 'S'"
''    SQLX$ = SQLX$ + " , ''"
''    SQLX$ = SQLX$ + " , ''"
''    SQLX$ = SQLX$ + " , ''"
''    SQLX$ = SQLX$ + " , ''"
''    SQLX$ = SQLX$ + " )"
''    FLEXCONN.Execute (SQLX$)
''
'''
''    SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
''    SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
''    SQLX$ = SQLX$ + " ,0"
''    SQLX$ = SQLX$ + " ,0"
''    SQLX$ = SQLX$ + " , 'IDENTIFICADOR/A32'"
''    SQLX$ = SQLX$ + " , ''"
''    SQLX$ = SQLX$ + " , 'S'"
''    SQLX$ = SQLX$ + " , ''"
''    SQLX$ = SQLX$ + " , ''"
''    SQLX$ = SQLX$ + " , ''"
''    SQLX$ = SQLX$ + " , ''"
''    SQLX$ = SQLX$ + " )"
''    FLEXCONN.Execute (SQLX$)
''
'''
''    SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
''    SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
''    SQLX$ = SQLX$ + " ,0"
''    SQLX$ = SQLX$ + " ,0"
''    SQLX$ = SQLX$ + " , 'CLAVE/24'"
''    SQLX$ = SQLX$ + " , ''"
''    SQLX$ = SQLX$ + " , 'S'"
''    SQLX$ = SQLX$ + " , ''"
''    SQLX$ = SQLX$ + " , ''"
''    SQLX$ = SQLX$ + " , ''"
''    SQLX$ = SQLX$ + " , ''"
''    SQLX$ = SQLX$ + " )"
''    FLEXCONN.Execute (SQLX$)
''
'''
''End If
'''|FIN USUARIOSQL

'|INICIO MOTIOTRA
'
CONDI$ = " LISTA =   'MOTIOTRA'"
EXILISTA% = CantRegistros&("ENCAPRFUN", CONDI$, "NOMESS")
If EXILISTA% > 0 Then
  FE% = 15714
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
  SQLX$ = SQLX$ + ",'MOTIOTRA'"
  SQLX$ = SQLX$ + ",'MOTIVOS DE ORDEN DE TRABAJO'"
  SQLX$ = SQLX$ + ",1"
  SQLX$ = SQLX$ + ",0"
  SQLX$ = SQLX$ + " ,1"
  SQLX$ = SQLX$ + " ,1"
  SQLX$ = SQLX$ + " ,0"
  SQLX$ = SQLX$ + " ,3"
  SQLX$ = SQLX$ + " ,0"
  SQLX$ = SQLX$ + " ,1"
  SQLX$ = SQLX$ + " ,'02/10/17'"
  SQLX$ = SQLX$ + " ,'MOTITELE'"
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
'|FIN MOTIOTRA


'|INICIO TipoCons
'|INICIO TipoCons
'
CONDI$ = " LISTA =   'TipoCons'"
EXILISTA% = CantRegistros&("ENCAPRFUN", CONDI$, "NOMESS")
If EXILISTA% > 0 Then
  FE% = 15714
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
  SQLX$ = SQLX$ + ",'TipoCons'"
  SQLX$ = SQLX$ + ",'Tipo Consuta '"
  SQLX$ = SQLX$ + ",1"
  SQLX$ = SQLX$ + ",0"
  SQLX$ = SQLX$ + " ,1"
  SQLX$ = SQLX$ + " ,1"
  SQLX$ = SQLX$ + " ,0"
  SQLX$ = SQLX$ + " ,3"
  SQLX$ = SQLX$ + " ,0"
  SQLX$ = SQLX$ + " ,1"
  SQLX$ = SQLX$ + " ,'02/10/17'"
  SQLX$ = SQLX$ + " ,''"
  SQLX$ = SQLX$ + " ,'SELECT CAMPO1, CAMPO2 FROM DATASQL WITH(NOLOCK)'"
  SQLX$ = SQLX$ + ")"
  FLEXCONN.Execute (SQLX$)
    SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
    SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
    SQLX$ = SQLX$ + " ,0"
    SQLX$ = SQLX$ + " ,0"
    SQLX$ = SQLX$ + " , 'Código/F8'"
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
    SQLX$ = SQLX$ + " , 'T.Consulta/A32'"
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
'|FIN TipoCons


'|INICIO USUARIOSQL   
'
CONDI$ = " LISTA =   'USUARIOSQL   '
EXILISTA% = CantRegistros&("ENCAPRFUN", CONDI$ ,"NOMESS")
If EXILISTA% > 0 Then
  FE% = 15715
  FEBD% = CVINT(VALOBADA("ENCAPRFUN","FECHEMISI", CONDI$, "NOMESS"))
  GENERAREST% = 0
  If FEBD% < FE% Then
    IDLISTAX& = XVALO(VALOBADA("ENCAPRFUN","ID_LISTA", CONDI$,"NOMESS"))
    Call BORRAREGISTROS("ENCAPRFUN","ID_LISTA = " & IDLISTAX&,REGAF&,"NOMESS")
    Call BORRAREGISTROS("ENCAPRF","ID_LISTA = " & IDLISTAX&,REGAF&,"NOMESS")
    GENERAREST% = 1
  END IF  
ELSE
  GENERAREST% = 1
  CONDI$ =  " ID_LISTA  > 0 "
  IDLISTAX& = 1 + XVALO(VALOBADA("ENCAPRFUN","MAX(ID_LISTA)",CONDI$,"NOMESS"))
END IF  
If GENERAREST% > 0 Then  
  SQLX$ = "INSERT INTO  ENCAPRFUN (ID_LISTA ,LISTA, Descripcion, ADDREGISTROS, USAFORMULA, BORRAR, INSERTAR, NOTILDE "
  SQLX$ = SQLX$ + ", CANTICAMPOS, CodiEmpr, USARECORD, FECHEMISI, TITURECORD, ConsultaSql)"
  SQLX$ = SQLX$ + "VALUES ( " & IDLISTAX&
  SQLX$ = SQLX$ + ",'USUARIOSQL   '
  SQLX$ = SQLX$ + ",'TABLA DE USUARIOS'
  SQLX$ = SQLX$ + ",1
  SQLX$ = SQLX$ + ",0
  SQLX$ = SQLX$ + " ,1
  SQLX$ = SQLX$ + " ,1
  SQLX$ = SQLX$ + " ,0
  SQLX$ = SQLX$ + " ,6
  SQLX$ = SQLX$ + " ,0
  SQLX$ = SQLX$ + " ,0
  SQLX$ = SQLX$ + " ,'03/10/17'
  SQLX$ = SQLX$ + " ,''
  SQLX$ = SQLX$ + " ,'SELECT CAMPO1, CAMPO2, CAMPO3, CAMPO4,CAMPO5 FROM DATASQL WITH(NOLOCK)'
  SQLX$ = SQLX$ + ")
  FLEXCONN.Execute (SQLX$)
    SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
    SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
    SQLX$ = SQLX$ + " ,0
    SQLX$ = SQLX$ + " ,0
    SQLX$ = SQLX$ + " , 'NUSER/F6'
    SQLX$ = SQLX$ + " , '     '
    SQLX$ = SQLX$ + " , '     '
    SQLX$ = SQLX$ + " , ''
    SQLX$ = SQLX$ + " , ''
    SQLX$ = SQLX$ + " , '     '
    SQLX$ = SQLX$ + " , ''
    SQLX$ = SQLX$ + " )
    FLEXCONN.Execute (SQLX$)

'
    SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
    SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
    SQLX$ = SQLX$ + " ,0
    SQLX$ = SQLX$ + " ,0
    SQLX$ = SQLX$ + " , 'NOMBRE/A32'
    SQLX$ = SQLX$ + " , ''
    SQLX$ = SQLX$ + " , 'S'
    SQLX$ = SQLX$ + " , ''
    SQLX$ = SQLX$ + " , ''
    SQLX$ = SQLX$ + " , '     '
    SQLX$ = SQLX$ + " , '     '
    SQLX$ = SQLX$ + " )
    FLEXCONN.Execute (SQLX$)

'
    SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
    SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
    SQLX$ = SQLX$ + " ,0
    SQLX$ = SQLX$ + " ,0
    SQLX$ = SQLX$ + " , 'IDENTIFICADOR/A32'
    SQLX$ = SQLX$ + " , ''
    SQLX$ = SQLX$ + " , 'S'
    SQLX$ = SQLX$ + " , ''
    SQLX$ = SQLX$ + " , ''
    SQLX$ = SQLX$ + " , ''
    SQLX$ = SQLX$ + " , ''
    SQLX$ = SQLX$ + " )
    FLEXCONN.Execute (SQLX$)

'
    SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
    SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
    SQLX$ = SQLX$ + " ,0
    SQLX$ = SQLX$ + " ,0
    SQLX$ = SQLX$ + " , 'CLAVE/24'
    SQLX$ = SQLX$ + " , ''
    SQLX$ = SQLX$ + " , 'S'
    SQLX$ = SQLX$ + " , ''
    SQLX$ = SQLX$ + " , ''
    SQLX$ = SQLX$ + " , ''
    SQLX$ = SQLX$ + " , ''
    SQLX$ = SQLX$ + " )
    FLEXCONN.Execute (SQLX$)

'
    SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
    SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
    SQLX$ = SQLX$ + " ,0
    SQLX$ = SQLX$ + " ,0
    SQLX$ = SQLX$ + " , 'Abreviatura/A2'
    SQLX$ = SQLX$ + " , ''
    SQLX$ = SQLX$ + " , 'S'
    SQLX$ = SQLX$ + " , ''
    SQLX$ = SQLX$ + " , ''
    SQLX$ = SQLX$ + " , ''
    SQLX$ = SQLX$ + " , ''
    SQLX$ = SQLX$ + " )
    FLEXCONN.Execute (SQLX$)

'
End If 
'|FIN USUARIOSQL   

End Sub
