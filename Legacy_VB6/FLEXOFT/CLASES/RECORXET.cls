VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "RECORXET"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"

Public Function RS_CODIGENE() As ADODB.Recordset
  Dim RS_FUN As ADODB.Recordset
  Set RS_FUN = New ADODB.Recordset
  SQLX$ = "SELECT id, TitVentana FROM GENECODI WITH(NOLOCK)"
  Set RS_FUN = Nothing
  Set RS_FUN = READSET(SQLX$)
  Set RS_CODIGENE = RS_FUN
  Set RS_FUN = Nothing
End Function
Public Function RS_FactPend_de_GenerarPedidoAHP(Pendientes%) As ADODB.Recordset
  Dim RS_FUN As ADODB.Recordset
  Set RS_FUN = New ADODB.Recordset
  
  SQLX$ = "SELECT   D.nume_cli,P.Raso_Cli,D.facturaprov,D.fechemisi,D.numebore,D.nume_ocom,D.Nroped "
  SQLX$ = SQLX$ + " FROM DETFACOMSQL D WITH(NOLOCK) INNER JOIN PROVED12 P on d.nume_cli = p.nume_cli "
  If Pendientes% > 0 Then SQLX$ = SQLX$ + "  WHERE (D.NROPED < 1 OR D.NROPED IS NULL)"
  SQLX$ = SQLX$ + " group by D.nume_cli,P.Raso_Cli,D.facturaprov,D.fechemisi,D.numebore,D.nume_ocom,D.Nroped"
  SQLX$ = SQLX$ + " ORDER BY D.fechemisi DESC"
  Set RS_FUN = Nothing
  Set RS_FUN = READSET(SQLX$)
  Set RS_FactPend_de_GenerarPedidoAHP = RS_FUN
  Set RS_FUN = Nothing
End Function

Public Function RS_PEDIPEN() As ADODB.Recordset
  Dim RS_FUN As ADODB.Recordset
  Set RS_FUN = New ADODB.Recordset
  SQLX$ = "SELECT CodiExt, Descrip, CanPed, PreUnid FROM PEDDETA WITH(NOLOCK)  WHERE CanPed >= CANTENT AND STATUS < 8  "
  Set RS_FUN = READSET(SQLX$)
  Set RS_PEDIPEN = RS_FUN
  Set RS_FUN = Nothing
End Function
'*****************************************************'

Public Function RS_MasterMateriasPrimas() As ADODB.Recordset
  Dim RS_FUN As ADODB.Recordset
  Set RS_FUN = New ADODB.Recordset
  SQLX$ = "SELECT Codigo, Descripcion, Consumo_X_Corte, Talle FROM MASTER WITH(NOLOCK) WHERE Timate = 1 AND Status >= 0"
  Set RS_FUN = READSET(SQLX$)
  Set RS_MasterMateriasPrimas = RS_FUN
  Set RS_FUN = Nothing
End Function
Public Function RS_MATERIASPRIMASDECODIGORANGO(CIRANG) As ADODB.Recordset
  Dim RS_FUN As ADODB.Recordset
  Set RS_FUN = New ADODB.Recordset
  SQLX$ = "SELECT CODRANGO,Descripcion,CODMATPRIMA,DESCRIMATPRIMA,CIRANGO,CIMATPRIMA,CODIGORANGO_ID"
  SQLX$ = SQLX$ + " FROM CODIGORANGO WITH(NOLOCK)  WHERE CIRANGO = " & CIRANG & " ORDER BY CODMATPRIMA ASC"
  Set RS_FUN = READSET(SQLX$)
  Set RS_MATERIASPRIMASDECODIGORANGO = RS_FUN
  Set RS_FUN = Nothing

End Function

Public Function RS_Pedidos_Sin_OT() As ADODB.Recordset
  'NO CONSIDERA ANULADOS NI CUMPLIDOS
  'NOTA: QUITE ESTA CONDICION DE LA SELECT 1 FROM ORDEFABR.... POR QUE SI NO NO FUNCIONABA 04/05/20
  Dim RS_FUN As ADODB.Recordset
  Set RS_FUN = New ADODB.Recordset
  SQLX$ = "Select  NROCLIE,RaSoClie, FECHEMIS, NROPED,  CodiExt, Descrip, CANPED, FESOLeNT,NroOcom,nuoritem from PEDDETA WITH(NOLOCK) WHERE CanPed > CantEnt AND STATUS < 8 "
  SQLX$ = SQLX$ + " AND not exists (select 1 from ORDEFABR where PEDDETA.NROPED  = ORDEFABR.NROPED and peddeta.CodiInt = ordefabr.Cod_Inte   ) ORDER BY PEDDETA.NROPED"
  Set RS_FUN = READSET(SQLX$)
  Set RS_Pedidos_Sin_OT = RS_FUN
  Set RS_FUN = Nothing
  
End Function

Public Function RS_PEDIPEN2(Optional todos%, Optional NC) As ADODB.Recordset
  Dim RS_FUN As ADODB.Recordset
  Set RS_FUN = New ADODB.Recordset
  WHERE_AND$ = " WHERE "
  SQLX$ = "SELECT NROCLIE,RaSoClie, FECHEMIS, NROPED,  CodiExt, Descrip, CANPED, FESOLeNT FROM PEDDETA WITH(NOLOCK)"
  If todos% < 1 Then SQLX$ = SQLX$ + WHERE_AND$ & " CanPed >= CANTENT AND STATUS < 8  and  STATUS >= 0 ": WHERE_AND$ = " And "
  If NC > 0 Then SQLX$ = SQLX$ + WHERE_AND$ & " NROCLIE = " & NC
  SQLX$ = SQLX$ + " ORDER BY NROPED DESC"
  Set RS_FUN = READSET(SQLX$)
  Set RS_PEDIPEN2 = RS_FUN
  Set RS_FUN = Nothing
End Function
Public Function RS_VENTACARGA(NPED&) As ADODB.Recordset
  Dim RS_FUN As ADODB.Recordset
  Set RS_FUN = New ADODB.Recordset
  SQLX$ = "SELECT CodiExt, Descrip, CanPed FROM PEDDETA WITH(NOLOCK) WHERE NroPed = " & NPED&
  Set RS_FUN = READSET(SQLX$)
  Set RS_VENTACARGA = RS_FUN
  Set RS_FUN = Nothing
End Function

Public Function RS_ORFPEN_ENVA(Optional INCLUIR_CERRADAS%, Optional AgregaAnul%) As ADODB.Recordset
  'SELECHO DE ENVAPLAST EN ORDEN DE FABRICACION   *** si AgregaAnul% > 0 le concatena a la colum descripcion (Anul)
  Dim RS_FUN As ADODB.Recordset
  Set RS_FUN = New ADODB.Recordset
  If AgregaAnul% > 0 Then
     SQLX$ = "SELECT ORDEFABR.IdSubOr, ORDEFABR.Cod_Exte, CASE WHEN STATUORF = 9 THEN CONCAT('(Anul) ',ORDEFABR.Descrip) else ORDEFABR.Descrip End As Descrip, "
  Else
     SQLX$ = "SELECT ORDEFABR.IdSubOr, ORDEFABR.Cod_Exte,ORDEFABR.Descrip, "
  End If
  SQLX$ = SQLX$ + "  ORDEFABR.CanProy, ORDEFABR.CanApro, ORDEFABR.NroPed, PEDENCA.NroClie, PEDENCA.RasoClie,ORDEFABR.NUORPROD FROM ORDEFABR WITH(NOLOCK) INNER JOIN PEDENCA ON ORDEFABR.NroPed = PEDENCA.NroPed "
  If INCLUIR_CERRADAS% < 1 Then SQLX$ = SQLX$ + " WHERE ORDEFABR.StatuOrf < '3' AND ORDEFABR.CanProy > ORDEFABR.CanApro "
  SQLX$ = SQLX$ + " ORDER BY IDSUBOR DESC"
  Set RS_FUN = READSET(SQLX$)
  Set RS_ORFPEN_ENVA = RS_FUN
  Set RS_FUN = Nothing
End Function
Public Function RS_PediOrfa() As ADODB.Recordset
  Dim RS_FUN As ADODB.Recordset
  Set RS_FUN = New ADODB.Recordset
  SQLX$ = "SELECT PEDDETA.NroPed, PEDDETA.NroClie, PEDDETA.RaSoClie, PEDDETA.CodiExt, PEDDETA.DESCRIP, PEDDETA.CanPed, PEDDETA.FeSolEnt, PEDDETA.CantEnt, ORDEFABR.IdSubOr, ORDEFABR.CanProy, ORDEFABR.CanApro FROM PEDDETA WITH(NOLOCK) INNER JOIN ORDEFABR ON PEDDETA.NroPed = ORDEFABR.NroPed WHERE ORDEFABR.Cod_Inte = PEDDETA.CODIINT AND PEDDETA.CanPed > PEDDETA.CanTENT ORDER BY PEDDETA.NROPED"
  Set RS_FUN = READSET(SQLX$)
  Set RS_PediOrfa = RS_FUN
  Set RS_FUN = Nothing
End Function
Public Function RS_VALIDAR() As ADODB.Recordset
  'VALIDA QUE LA SUMA DE LOS CAMPOS IDEBECOMP E IHABECOMP SEAN IGUALES
  'ESTO SURGIO POR QUE EN MEDITEA HUBO UN PROBLEMA CON UNA FACTURA QUE GRABO TODO BIEN PERO
  'EN EL CAMPO IDEBECOMP GRABO EL IMPORTE DE UNA FACTURA ANTERIOR.
  
  Dim RS_FUN As ADODB.Recordset
  Set RS_FUN = New ADODB.Recordset
  SQLX$ = "SELECT CODICOM_LAR , SUM(IDebeComp) AS SxIDebeComp, SUM(IHabeComp) AS SxIHabeComp FROM CAJABANC WITH(NOLOCK) WHERE CODICOM_LAR LIKE 'FC%'  AND NUCLIEN > 0  GROUP BY CODICOM_LAR  HAVING  abs(SUM(IHabeComp) - SUM(IDebeComp)) > 0.05 order by CODICOM_LAR "
  Set RS_FUN = READSET(SQLX$)
  Set RS_VALIDAR = RS_FUN
  Set RS_FUN = Nothing
End Function
'*****************************************************'
Public Function RS_DEUDOR12() As ADODB.Recordset
  Dim RS_FUN As ADODB.Recordset
  Set RS_FUN = New ADODB.Recordset
  SQLX$ = "SELECT Nume_Cli, Raso_Cli FROM DEUDOR12  WITH(NOLOCK)  WHERE Stat_Cli >= 0 order by nume_cli"
  Set RS_FUN = READSET(SQLX$)
  Set RS_DEUDOR12 = RS_FUN
  Set RS_FUN = Nothing
End Function
'*****************************************************'
'ESTE TRAE TODOS LOS PEDIDOS pendientes y no NO funciona bien
'*****************************************************'
Public Function RS_PedidoPendxCte(NC) As ADODB.Recordset
  Dim RS_FUN As ADODB.Recordset
  Set RS_FUN = New ADODB.Recordset

  SQLX$ = "SELECT NroPed, Referencia FROM PEDENCA WITH(NOLOCK) WHERE Status < 8 AND Status >= 0 "
  If NC > 0 Then SQLX$ = SQLX$ + " AND NROCLIE = " & NC
  SQLX$ = SQLX$ + " ORDER BY NroPed DESC"
  Set RS_FUN = READSET(SQLX$)
  Set RS_PedidoPendxCte = RS_FUN
  Set RS_FUN = Nothing
End Function
'*****************************************************'
'*****************************************************'
Public Function RS_PedidoPendxCte2(NC) As ADODB.Recordset
  'ESTE FUNCIONA BIEN PARA MOSTRAR ENCABEZADOS DE PEDIDOS
  Dim RS_FUN As ADODB.Recordset
  Set RS_FUN = New ADODB.Recordset

  SQLX$ = "SELECT DISTINCT P.NroPed, P.Referencia FROM PEDENCA P WITH(NOLOCK)  "
  SQLX$ = SQLX$ + " INNER JOIN PEDDETA PP ON P.NROPED = PP.NROPED"
  SQLX$ = SQLX$ + " WHERE P.Status < 8 AND P.Status >= 0"
  SQLX$ = SQLX$ + " AND PP.Status < 8 "    ' NO ANULADOS NI CUMPLIDOS
  SQLX$ = SQLX$ + " AND PP.STATUS >= 0 "   ' APROBADOS
  SQLX$ = SQLX$ + " AND (((PP.CanPed - PP.CantEnt) >= 0.05) OR (PP.TipoPed = 'ABIERTO' AND PP.FinVigen >= getdate()))"

  If NC > 0 Then SQLX$ = SQLX$ + "  AND PP.NROCLIE = " & NC
  SQLX$ = SQLX$ + " ORDER BY P.NroPed DESC ,P.Referencia "
  Set RS_FUN = READSET(SQLX$)
  Set RS_PedidoPendxCte2 = RS_FUN
  Set RS_FUN = Nothing
End Function

'*****************************************************'
Public Function RS_PedidosVsOF(NROP&) As ADODB.Recordset
  Dim RS_FUN As ADODB.Recordset
  Set RS_FUN = New ADODB.Recordset
  SQLX$ = "SELECT PEDDETA.CodiExt, PEDDETA.Descrip, PEDDETA.CanPed, PEDDETA.CantEnt, ORDEFABR.IdSubOr, ORDEFABR.CanProy, ORDEFABR.CanApro, PEDDETA.NuOrItem FROM PEDDETA WITH(NOLOCK)  INNER JOIN ORDEFABR ON PEDDETA.NroPed = ORDEFABR.NroPed WHERE ORDEFABR.Cod_Inte = PEDDETA.CODIINT "
  If NROP& > 0 Then SQLX$ = SQLX$ + " AND PEDDETA.NROPED = " & NROP&
  SQLX$ = SQLX$ + " ORDER BY PEDDETA.NuOrItem ASC"
  
  Set RS_FUN = READSET(SQLX$)
  Set RS_PedidosVsOF = RS_FUN
  Set RS_FUN = Nothing
End Function
'*****************************************************'

Public Function RS_DETPED() As ADODB.Recordset
  Dim RS_FUN As ADODB.Recordset
  Set RS_FUN = New ADODB.Recordset
  SQLX$ = "SELECT CodiExt, Descrip, CanPed, PreUnid FROM PEDDETA WITH(NOLOCK)  WHERE NroPed < '10'"
  Set RS_FUN = READSET(SQLX$)
  Set RS_DETPED = RS_FUN
  Set RS_FUN = Nothing
End Function

Public Function RS_ENCAPRFLISTAS() As ADODB.Recordset
  Dim RS_FUN As ADODB.Recordset
  Set RS_FUN = New ADODB.Recordset
  SQLX$ = "SELECT DISTINCT ID_LISTA, LISTA,DESCRIPCION FROM ENCAPRFUN WITH(NOLOCK) WHERE  (MARBORRA IS NULL OR MARBORRA = 0)"
  Set RS_FUN = READSET(SQLX$)
  Set RS_ENCAPRFLISTAS = RS_FUN
  Set RS_FUN = Nothing
End Function
Public Function RS_ConsInfofab(NORFA$) As ADODB.Recordset
  Dim RS_FUN As ADODB.Recordset
  Set RS_FUN = New ADODB.Recordset
  SQLX$ = "select CodiElem, CODXELEM, cantreal,consumo,  idenlote,corte, FACTORCORTE from infofab WITH(NOLOCK)"
  SQLX$ = SQLX$ + " where idsubor = '" & NORFA$ & "' AND CANTREAL > CONSUMO"
  Set RS_FUN = READSET(SQLX$)
  Set RS_ConsInfofab = RS_FUN
  Set RS_FUN = Nothing
End Function
'Public Function RS_MUESTRAENCAPRF(Optional NROIDLIST&) As ADODB.Recordset
'  'MUESTRA LA LISTA DE CONFIGURACION SE LLAMA EN LA CARGA DE CONFIGURACION DE ESTRUCTURA
'  Dim RS_FUN As ADODB.Recordset
'  Set RS_FUN = New ADODB.Recordset
'  SQLX$ = "SELECT  TIPO_ID,  CAMPO, PROP_VENTANA, ENCABEZADO, ORDENPRES, FORMATO FROM ENCAPRF WITH(NOLOCK)  WHERE Status >= 0"
'  If NROIDLIST& > 0 Then SQLX$ = SQLX$ + " AND ID_LISTA = " & NROIDLIST&
'  Set RS_FUN = READSET(SQLX$)
'  Set RS_MUESTRAENCAPRF = RS_FUN
'  Set RS_FUN = Nothing
'End Function
Public Function RS_TAB_COMP_ENCAPRFUN(Optional NROIDLIST&) As ADODB.Recordset
  Dim RS_FUN As ADODB.Recordset
  Set RS_FUN = New ADODB.Recordset
  SQLX$ = "SELECT ID_LISTA, CodiEmpr, ADDREGISTROS, BORRAR, CANTICAMPOS, CONSULTASQL, ENCAPRFUN_ID,DESCRIPCION, FECHEMISI, INSERTAR, LISTA, MARBORRA, NOTILDE, USAFORMULA, USARECORD,TITURECORD FROM ENCAPRFUN WITH(NOLOCK)"
  If NROIDLIST& > 0 Then SQLX$ = SQLX$ + " WHERE ID_LISTA = " & NROIDLIST&
  Set RS_FUN = READSET(SQLX$)
  Set RS_TAB_COMP_ENCAPRFUN = RS_FUN
  Set RS_FUN = Nothing
End Function

''''''''''''''''
Public Function RS_DATASQL() As ADODB.Recordset
  Dim RS_FUN As ADODB.Recordset
  Set RS_FUN = New ADODB.Recordset
  SQLX$ = "SELECT CAMPO1, CAMPO2, CAMPO3, CAMPO4,CAMPO5, CAMPO62, CAMPO7, CAMPO8, CAMPO9, CAMPO10, CAMPO11, CAMPO12 FROM PEDDETA WITH(NOLOCK)  WHERE NroPed < '10'"
  Set RS_FUN = READSET(SQLX$)
  Set RS_DATASQL = RS_FUN
  Set RS_FUN = Nothing
End Function
Public Function RS_CODIMAQ(Optional textsector$) As ADODB.Recordset
  Dim RS_FUN As ADODB.Recordset
  Set RS_FUN = New ADODB.Recordset
  SQLX$ = "SELECT CAMPO1, CAMPO2 FROM DATASQL WITH(NOLOCK) WHERE LISTA  = 'CODIMAQ'"
  If TRIM$(textsector$) <> "" Then
    SEXTOR% = XVALO(textsector$)
    SQLX$ = SQLX$ + " AND CAMPO3 = " & SEXTOR%
  End If
  Set RS_FUN = READSET(SQLX$)
  Set RS_CODIMAQ = RS_FUN
  Set RS_FUN = Nothing
End Function
Public Function RS_ARCHIVOSADJ(CI1%) As ADODB.Recordset
  Dim RS_FUN As ADODB.Recordset
  Set RS_FUN = New ADODB.Recordset
  SQLX$ = "SELECT Descripcion, Ruta FROM DOCUREL WITH(NOLOCK)  WHERE Cod_Inte = " & CI1%
  Set RS_FUN = READSET(SQLX$)
  Set RS_ARCHIVOSADJ = RS_FUN
  Set RS_FUN = Nothing
End Function

Public Function RS_HORASXUSUARIO(USUARIO%, Desde%, Hasta%) As ADODB.Recordset
  Dim RS_FUN As ADODB.Recordset
  Set RS_FUN = New ADODB.Recordset
  SQLX$ = "SELECT FECHA, IDSUBOR, H_DESDE, H_HASTA FROM PDIARPROD1 WITH(NOLOCK)"
  SQLX$ = SQLX$ + "  WHERE PDIARPROD1_ID > 0"
  If USUARIO% > 0 Then
    SQLX$ = SQLX$ + "  AND (NUMAYUD1 =  " & USUARIO%
    SQLX$ = SQLX$ + "  OR NUMAYUD2 =  " & USUARIO%
    SQLX$ = SQLX$ + "  OR NUMAYUD3 =  " & USUARIO%
    SQLX$ = SQLX$ + "  OR NUMAYUD4 =  " & USUARIO% & ")"
  End If
  If Desde% > 0 Then SQLX$ = SQLX$ + "  AND FECHA >= '" & FETEXCOR1$(Desde%) & "'"
  If Hasta% > 0 Then SQLX$ = SQLX$ + "  AND FECHA <= '" & FETEXCOR1$(Hasta%) & "'"
  SQLX$ = SQLX$ + " ORDER BY FECHA"
  Set RS_FUN = READSET(SQLX$)
  Set RS_HORASXUSUARIO = RS_FUN
  Set RS_FUN = Nothing
End Function

Public Function RS_CARGADODESC(NORFA$, SOLAPAINF%, SOLAPASUP%) As ADODB.Recordset
  Dim RS_FUN As ADODB.Recordset
  Set RS_FUN = New ADODB.Recordset

  SQLX$ = "SELECT FECHA,NUMMAQUI,NAME_MAQUI,CODIDESC, DESCRIDESC, CANT_DESC,OBSERVACIONES FROM DESCARCONS WITH(NOLOCK)"
  SQLX$ = SQLX$ + " WHERE IDSUBOR = '" & NORFA$ & "'"
  SQLX$ = SQLX$ + " AND INDISOLINF= " & SOLAPAINF%
  SQLX$ = SQLX$ + " AND INDISOLSUP = " & SOLAPASUP%
  Set RS_FUN = READSET(SQLX$)
  Set RS_CARGADODESC = RS_FUN
  Set RS_FUN = Nothing
End Function

    'ESTO VA EN LA CLASE RECORXET
Public Function RS_SELECODIMAQUI(Optional textsector$) As ADODB.Recordset
      Dim RS_FUN As ADODB.Recordset
      Set RS_FUN = New ADODB.Recordset
      'recibo el parametro sector como texto, a pesar que es un nro. por que el 0 tambien es _
      un sector de este modo puedo validara si esta vacio o tiene un valor
      SQLX$ = "SELECT CAMPO1, CAMPO2 FROM DATASQL WITH(NOLOCK)  where lista = 'CODIMAQ'"
      If TRIM$(textsector$) <> "" Then
        SEXTOR% = XVALO(textsector$)
        SQLX$ = SQLX$ + " AND CAMPO3 = " & SEXTOR%
      End If
      Set RS_FUN = READSET(SQLX$)
      Set RS_SELECODIMAQUI = RS_FUN
      Set RS_FUN = Nothing
End Function
Public Function RS_PDPRODIMPRE(CODMAQ, FECHX)
    Dim RS_FUN As ADODB.Recordset
    Set RS_FUN = New ADODB.Recordset
    SQLX$ = "   SELECT P.IDSUBOR,P.NroPed,P.NroClie,P.NAMEAYUD1,P.NAMEAYUD2,P.H_DESDE,P.H_HASTA,P.H_TRABAJADAS,P.CANTOK FROM "
    SQLX$ = SQLX$ + " (SELECT IDSUBOR,NroPed,NroClie,NAMEAYUD1,NAMEAYUD2,H_DESDE,H_HASTA,H_TRABAJADAS,CANTOK, FECHA,NUMMAQUI  FROM PDIARPROD1 "
    SQLX$ = SQLX$ + " union All "
    SQLX$ = SQLX$ + " SELECT IDSUBOR,NroPed,NroClie,NAMEOP1,NAMEOP2,H_DESDE,H_HASTA,H_TRABAJADAS,CANTIDAD,FECHA,NUMMAQUI  FROM PDIARPROD2)  AS P "
    SQLX$ = SQLX$ + "  WHERE P.NUMMAQUI  = '" & CODMAQ & "' AND FECHA = '" & FECHX & "'"
    SQLX$ = SQLX$ + " ORDER BY P.IDSUBOR,P.FECHA "
    Set RS_FUN = READSET(SQLX$, , 1)
    Set RS_PDPRODIMPRE = RS_FUN
    Set RS_FUN = Nothing
End Function

    Public Function RS_USUSQL() As ADODB.Recordset
      Dim RS_FUN As ADODB.Recordset
      Set RS_FUN = New ADODB.Recordset
      SQLX$ = "SELECT NUSER, NOMBRE FROM USUARIOSSQL WITH(NOLOCK)  ORDER BY NOMBRE ASC"
      Set RS_FUN = READSET(SQLX$)
      Set RS_USUSQL = RS_FUN
      Set RS_FUN = Nothing
    End Function
    '*****************************************************'
    
    Public Function RS_USUSQLConClave(Ident$, KLAVE$) As ADODB.Recordset
      Dim RS_FUN As ADODB.Recordset
      Set RS_FUN = New ADODB.Recordset
      SQLX$ = "SELECT NUSER,NOMBRE,IDENTIFICADOR,CLAVE FROM USUARIOSSQL WITH(NOLOCK) "
      SQLX$ = SQLX$ + " WHERE IDENTIFICADOR = '" & Ident$ & "' AND CLAVE = '" & KLAVE$ & "' ORDER BY NOMBRE ASC"
      Set RS_FUN = READSET(SQLX$)
      Set RS_USUSQLConClave = RS_FUN
      Set RS_FUN = Nothing
    End Function


Public Function RS_VALOASIG() As ADODB.Recordset
  EPAC$ = TRIM$(UCase$(EMPREAC$))
  If InStr(EPAC$, "ENVAPLAST") Then Call DEPURAVALORASIGNADO_ENVAPLAST

  Dim RS_FUN As ADODB.Recordset
  Set RS_FUN = New ADODB.Recordset
  'HAGO UN REPLACE POR SI EL NUMERO ESTA GUARDADO CON COMA EN VEZ DE PUNTO.
  SQLX$ = "SELECT Talle FROM MASTER WITH(NOLOCK) where (talle <> '' AND talle is not null)  ORDER BY (case  when isnumeric(TALLE) = 1 then cast (REPLACE(TALLE,',','.') as float) end)"
  Set RS_FUN = READSET(SQLX$)
  
  Set RS_VALOASIG = RS_FUN
  Set RS_FUN = Nothing
End Function
'*****************************************************'
Sub DEPURAVALORASIGNADO_ENVAPLAST()
  'BORRA DE LA TABLA MASTER CAMPO TALLE LOS DATOS NO NUMERICOS
  'SOLO USAR EN ENVAPLAST
  Dim Rst As ADODB.Recordset
  Set Rst = New ADODB.Recordset

  SQLX$ = "SELECT Talle FROM MASTER  where talle <> ''  "
  Rst.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
  With Rst
    Do While Not .EOF
      If XVALO(!TALLE) <= 0 Then
        !TALLE = ""
        .Update
      End If
      .MoveNext
    Loop
  End With
  Rst.Close
  Set Rst = Nothing
End Sub
Public Function RS_DetaPedido(NPX&) As ADODB.Recordset
  Dim RS_FUN As ADODB.Recordset
  Set RS_FUN = New ADODB.Recordset
  SQLX$ = "SELECT codIint,CodiExt, Descrip, CanPed, PreUnid,FESOLENT,nroclie,NUORITEM FROM PEDDETA WITH(NOLOCK)  WHERE NROPED = " & NPX&
  Set RS_FUN = READSET(SQLX$)
  Set RS_DetaPedido = RS_FUN
  Set RS_FUN = Nothing
End Function

Public Function RS_ANCHOYGRAMAJE(CI1%) As ADODB.Recordset
  Dim RS_FUN As ADODB.Recordset
  Set RS_FUN = New ADODB.Recordset
  SQLX$ = "SELECT ANCHO, GRAMAJE, MICRONES FROM MASTER WITH(NOLOCK)  WHERE Codint = " & CI1%
  Set RS_FUN = READSET(SQLX$)
  Set RS_ANCHOYGRAMAJE = RS_FUN
  Set RS_FUN = Nothing
End Function
'*****************************************************'
Public Function RS_CONSUMOX3(NORFA$)
   Dim RS_FUN As ADODB.Recordset
   Set RS_FUN = New ADODB.Recordset
   SQLX$ = "SELECT C.CODINT,C.IDSUBOR,C.CODIGO,C.DESCRIPCION,SUM(C.CONSUMO)AS CONSUMOX,C.MARBORRA "
   SQLX$ = SQLX$ + " FROM (SELECT CODINT,IDSUBOR,CODIGO,DESCRIPCION,CONSUMO,MARBORRA"
   SQLX$ = SQLX$ + " From CONSUMO1"
   SQLX$ = SQLX$ + " Union All"
   SQLX$ = SQLX$ + " SELECT CODINT,IDSUBOR,CODIGO,DESCRIPCION,CONSUMO,MARBORRA FROM CONSUMO2"
   SQLX$ = SQLX$ + " Union All"
   SQLX$ = SQLX$ + " SELECT CODINT,IDSUBOR,CODIGO,DESCRIPCION,CONSUMO,MARBORRA FROM CONSUMO3)"
   SQLX$ = SQLX$ + " AS C WHERE C.IDSUBOR ='" & NORFA$ & "'"
   'SQLX$ = SQLX$ + " AND CODINT = " & CI1%
   SQLX$ = SQLX$ + " AND (C.MARBORRA < 1 OR C.MARBORRA IS NULL)"
   SQLX$ = SQLX$ + " GROUP BY C.CODINT,C.IDSUBOR,C.CODIGO,C.DESCRIPCION,MARBORRA"
   'SQLX$ = SQLX$ + " ORDER BY C.INDISOLSUP"
   Set RS_FUN = READSET(SQLX$, , 1)
   Set RS_CONSUMOX3 = RS_FUN
   Set RS_FUN = Nothing
End Function
Public Function RS_Situación_de_Pendientes_de_Entrega(Optional Desde%, Optional Hasta%) As ADODB.Recordset
  Dim RS_FUN As ADODB.Recordset
  Set RS_FUN = New ADODB.Recordset
  CONDI$ = " P.NroPed > 0 "
  CONDI$ = CONDI$ + " AND P.Status < 8 "
  CONDI$ = CONDI$ + " AND P.StaDespa < 2 "
  CONDI$ = CONDI$ + " AND P.CanPed > P.AcuDespa"
  CONDI$ = CONDI$ + " AND P.VenRep = 'V' "             ' Excluye Reparaciones
  CONDI$ = CONDI$ + " AND P.CanPed > P.CantEnt "
  If Desde% > 0 And Hasta% > 0 Then
    DESDEX$ = FETEXCOR1$(Desde%)
    HASTAX$ = FETEXCOR1$(Hasta%)
    CONDI$ = CONDI$ + "  AND P.Fechemis >= '" & DESDEX$ & "' AND P.Fechemis <= '" & HASTAX$ & "'"
  End If
  SQLX$ = "SELECT P.NroClie, PE.RasoClie, P.NroPed,P.CODIINT, P.CodiExt, P.Descrip, P.CanPed, P.CantEnt ,p.fechemis "
'  If Desde% > 0 And Hasta% > 0 Then SQLX$ = SQLX$ + " ,p.fechemis"
  SQLX$ = SQLX$ + " FROM PEDENCA PE WITH(NOLOCK) INNER JOIN PEDDETA P ON PE.NroPed = P.NroPed WHERE " & CONDI$
  Set RS_FUN = READSET(SQLX$)
  Set RS_Situación_de_Pendientes_de_Entrega = RS_FUN
  Set RS_FUN = Nothing
End Function
Public Function RS_PEDIDOS_TODOS_X_FECHA(Des%, Has%) As ADODB.Recordset
  'ESTE TIENE FECHA DE ENTREGA Y FECHA DE EMISION
  Dim RS_FUN As ADODB.Recordset
  Set RS_FUN = New ADODB.Recordset
  CONDI$ = " P.NroPed > 0 "
  CONDI$ = CONDI$ + " AND P.Status < 8 "
  CONDI$ = CONDI$ + " AND P.VenRep = 'V' "             ' Excluye Reparaciones
  CONDI$ = CONDI$ + " AND P.FeSolEnt >= '" & FETEXCOR1$(Des%) & "'"             ' Excluye Reparaciones
  CONDI$ = CONDI$ + " AND P.FeSolEnt <= '" & FETEXCOR1$(Has%) & " '"
  '
  SQLX$ = "SELECT P.NroClie, PE.RasoClie, P.NroPed,P.CODIINT, P.CodiExt, P.Descrip, P.CanPed, P.CantEnt,P.FeSolEnt, P.FechEmis "
  SQLX$ = SQLX$ + " FROM PEDENCA PE WITH(NOLOCK) INNER JOIN PEDDETA P ON PE.NroPed = P.NroPed WHERE " & CONDI$
  'POR QUE ES PARA JYM Y TIENE QUE TRAER DE CODIEMPR 0 Y 1
  RS_FUN.CursorLocation = adUseClient
  RS_FUN.CursorType = adOpenForwardOnly
  RS_FUN.LockType = adLockReadOnly
  RS_FUN.Open SQLX$, FLEXCONN, , , adCmdText

  'Set RS_FUN = READSET(SQLX$)
  Set RS_PEDIDOS_TODOS_X_FECHA = RS_FUN
  Set RS_FUN = Nothing
End Function

Public Function RS_DOTADIARIA(FE%, SOLASUX%, Optional NORFA$) As ADODB.Recordset
  'SI NO SE QUIERE VALIDAR POR SOLAPA PASAR PARAMETRO COMO -1
  Dim RS_FUN As ADODB.Recordset
  Set RS_FUN = New ADODB.Recordset
  SQLX$ = "SELECT FECHA, NUSER, NOMBRE, H_DESDE, H_HASTA,NUMMAQUI,NAME_MAQUI,DOTACION_ID FROM DOTACION WITH(NOLOCK) "
  SQLX$ = SQLX$ + " WHERE FECHA = '" & FETEXCOR1$(FE%) & "'"
  If SOLASUX% >= 0 Then SQLX$ = SQLX$ + " AND  INDISOLSUP = " & SOLASUX%
  If TRIM$(NORFA$) <> "" Then SQLX$ = SQLX$ + " AND  IDSUBOR = '" & NORFA$ & "'"
  Set RS_FUN = READSET(SQLX$)
  Set RS_DOTADIARIA = RS_FUN
  Set RS_FUN = Nothing
End Function

Public Function RS_HORASSININFORMAR() As ADODB.Recordset
    Dim RS_FUN As ADODB.Recordset
    Set RS_FUN = New ADODB.Recordset
    SQLX$ = "SELECT FECHA, SOLAPASUP, NUSER, NOMBRE, H_DESDE, H_HASTA FROM DOTACION WITH(NOLOCK)  WHERE MARBORRA < 1 AND (H_DESDE = '' OR H_HASTA = '' OR H_DESDE IS NULL OR H_HASTA IS NULL) ORDER BY FECHA DESC"
    Set RS_FUN = READSET(SQLX$)
    Set RS_HORASSININFORMAR = RS_FUN
    Set RS_FUN = Nothing
End Function
Public Function RS_HORASDOT(Des%, Has%) As ADODB.Recordset
    Dim RS_FUN As ADODB.Recordset
    Set RS_FUN = New ADODB.Recordset
    SQLX$ = "SELECT FECHA,IDSUBOR, SOLAPASUP, NUSER, NOMBRE, H_DESDE, H_HASTA, H_TRABAJADAS FROM DOTACION WITH(NOLOCK) WHERE FECHA >= '" & FETEXCOR1$(Des%) & "'  AND FECHA <= '" & FETEXCOR1$(Has%) & "' ORDER BY FECHA,SOLAPASUP, NUSER"
    Set RS_FUN = READSET(SQLX$)
    Set RS_HORASDOT = RS_FUN
    Set RS_FUN = Nothing
End Function
        

Public Function RS_MAESTRODECLIENTES() As ADODB.Recordset
    
    Dim RS_FUN As ADODB.Recordset
    SQLX$ = "SELECT NUME_CLI, RASO_CLI FROM DEUDOR12 WITH(NOLOCK) WHERE STAT_CLI > 0 ORDER BY RASO_CLI"
    Set RS_FUN = READSET(SQLX$)
    Set RS_MAESTRODECLIENTES = RS_FUN
    On Error Resume Next
    Set RS_FUN = Nothing
    On Error GoTo 0

End Function
Public Function RS_MAESTRODECLIENTESxVENDEDOR(NroVend As Integer) As ADODB.Recordset
    'FILTRA POR EL VENDEDOR DE DEUDOR12 SI ES 0 TRAE TODOS
    Dim RS_FUN As ADODB.Recordset
    SQLX$ = "SELECT NUME_CLI, RASO_CLI FROM DEUDOR12 WITH(NOLOCK) WHERE STAT_CLI > 0 "
    If NroVend > 0 Then SQLX$ = SQLX$ + " AND VEND_CLI = " & NroVend
    SQLX$ = SQLX$ + " ORDER BY RASO_CLI"
    Set RS_FUN = READSET(SQLX$)
    Set RS_MAESTRODECLIENTESxVENDEDOR = RS_FUN
    On Error Resume Next
    Set RS_FUN = Nothing
    On Error GoTo 0

End Function

Public Function RS_MAESTRODEPROVEEDORES() As ADODB.Recordset
    
    Dim RS_FUN As ADODB.Recordset
    SQLX$ = "SELECT NUME_CLI, RASO_CLI FROM PROVED12 WITH(NOLOCK) WHERE STAT_CLI > 0 ORDER BY RASO_CLI"
    Set RS_FUN = READSET(SQLX$)
    Set RS_MAESTRODEPROVEEDORES = RS_FUN
    On Error Resume Next
    Set RS_FUN = Nothing
    On Error GoTo 0

End Function

Public Function RS_MAESTRODEARTICULOS(Optional CONDI$) As ADODB.Recordset
    
    Dim RS_FUN As ADODB.Recordset
    SQLX$ = "SELECT CODIGO, DESCRIPCION FROM MASTER WITH(NOLOCK) WHERE STATUS > 0 "
    If TRIM$(CONDI$) <> "" Then SQLX$ = SQLX$ & " AND " & CONDI$
    SQLX$ = SQLX$ + " ORDER BY CODIGO"
    Set RS_FUN = READSET(SQLX$)
    Set RS_MAESTRODEARTICULOS = RS_FUN
    On Error Resume Next
    Set RS_FUN = Nothing
    On Error GoTo 0

End Function

Public Function RS_ESTADISTICA_CONSUMO_AGRUPADA_X_MES_ENCOLUMNADO(Des%, Has%, DEP%, SQLCODMOV$, TMAT%, C1%, C2%, Optional SOLOSALIDA%) As ADODB.Recordset
    Dim RS_FUN As ADODB.Recordset
    CODI1$ = CODEXT$(C1%)
    CODI2$ = CODEXT$(C2%)

    SQLX$ = "SELECT  YEAR(MOV.FechMov) AS ANO ,MONTH(MOV.FechMov) as mes1 ,"
    SQLX$ = SQLX$ + "  Case Month(MOV.FechMov)"
    SQLX$ = SQLX$ + "   WHEN '1' THEN 'Enero'"
    SQLX$ = SQLX$ + "   WHEN '2' THEN 'Febrero'"
    SQLX$ = SQLX$ + "   WHEN '3' THEN 'Marzo'"
    SQLX$ = SQLX$ + "   WHEN '4' THEN 'Abril'"
    SQLX$ = SQLX$ + "   WHEN '5' THEN 'Mayo'"
    SQLX$ = SQLX$ + "   WHEN '6' THEN 'Junio'"
    SQLX$ = SQLX$ + "   WHEN '7' THEN 'Julio'"
    SQLX$ = SQLX$ + "   WHEN '8' THEN 'Agosto'"
    SQLX$ = SQLX$ + "   WHEN '9' THEN 'Septiembre'"
    SQLX$ = SQLX$ + "   WHEN '10' THEN 'Octubre'"
    SQLX$ = SQLX$ + "   WHEN '11' THEN 'Noviembre'"
    SQLX$ = SQLX$ + "   WHEN '12' THEN 'Diciembre' "
    SQLX$ = SQLX$ + "   End"
    SQLX$ = SQLX$ + " MES "
    
    SQLX$ = SQLX$ + " ,MOV.Cod_Exte "
    
    SQLX$ = SQLX$ + " ,sum(MOV.CANTSALID)as sale"
    SQLX$ = SQLX$ + " ,sum(MOV.CANTINGRE)as INGRE"
    
    SQLX$ = SQLX$ + " , MOV.Cod_Inte,MOV.DESCRIP, MAS.coscal, MAS.STOMIN, MAS.leticom "
    SQLX$ = SQLX$ + " FROM MOVISTO MOV   INNER JOIN MASTER MAS ON MOV.COD_INTE = MAS.CODINT "
    SQLX$ = SQLX$ + " WHERE MOV.FechMov >= '" & FETEXCOR1$(Des%) & "' AND MOV.FechMov <= '" & FETEXCOR1$(Has%) & "'  and MOV.Cod_Inte >0"
    
    SQLX$ = SQLX$ + " AND MOV.COD_EXTE >= '" & CODI1$ & "' "
    SQLX$ = SQLX$ + " AND MOV.COD_EXTE <= '" & CODI2$ & "' "
    If TMAT% > 0 Then
         SQLX$ = SQLX$ + " AND MAS.Timate = " & TMAT%
    End If
    If DEP% > 0 Then
         SQLX$ = SQLX$ + " AND MOV.DEPOMOVI = " & DEP%
    End If
    If SOLOSALIDA% > 0 Then
        SQLX$ = SQLX$ + " AND MOV.CantSalid > 0"
    End If
    SQLX$ = SQLX$ + SQLCODMOV$
    
    SQLX$ = SQLX$ + " group by  YEAR(MOV.fechmov), MONTH(MOV.fechmov), MOV.COD_EXTE,"
    SQLX$ = SQLX$ + " MOV.COD_INTE,MOV.DESCRIP, MAS.coscal, MAS.STOMIN, MAS.leticom"
    
    SQLX$ = SQLX$ + " order by  MOV.COD_EXTE, YEAR(MOV.fechmov), MONTH(MOV.fechmov)"
    SQLX$ = SQLX$ + " "
    
    Set RS_FUN = READSET(SQLX$, , 1)
    Set RS_ESTADISTICA_CONSUMO_AGRUPADA_X_MES_ENCOLUMNADO = RS_FUN
    Set RS_FUN = Nothing
    
End Function

Public Function RS_ESTADISTICA_CONSUMO_AGRUPADA_X_MES(Des%, Has%, DEP%, SQLCODMOV$, TMAT%, C1%, C2%, Optional SOLOSALIDA%) As ADODB.Recordset
    Dim RS_FUN As ADODB.Recordset
    CODI1$ = CODEXT$(C1%)
    CODI2$ = CODEXT$(C2%)

    SQLX$ = "SELECT  YEAR(MOV.FechMov) AS ANO ,MONTH(MOV.FechMov) as mes ,"
    SQLX$ = SQLX$ + "  Case Month(MOV.FechMov)"
    SQLX$ = SQLX$ + "   WHEN '1' THEN 'Enero'"
    SQLX$ = SQLX$ + "   WHEN '2' THEN 'Febrero'"
    SQLX$ = SQLX$ + "   WHEN '3' THEN 'Marzo'"
    SQLX$ = SQLX$ + "   WHEN '4' THEN 'Abril'"
    SQLX$ = SQLX$ + "   WHEN '5' THEN 'Mayo'"
    SQLX$ = SQLX$ + "   WHEN '6' THEN 'Junio'"
    SQLX$ = SQLX$ + "   WHEN '7' THEN 'Julio'"
    SQLX$ = SQLX$ + "   WHEN '8' THEN 'Agosto'"
    SQLX$ = SQLX$ + "   WHEN '9' THEN 'Septiembre'"
    SQLX$ = SQLX$ + "   WHEN '10' THEN 'Octubre'"
    SQLX$ = SQLX$ + "   WHEN '11' THEN 'Noviembre'"
    SQLX$ = SQLX$ + "   WHEN '12' THEN 'Diciembre' "
    SQLX$ = SQLX$ + "   End"
    SQLX$ = SQLX$ + " MES "
    
    SQLX$ = SQLX$ + " ,MOV.Cod_Exte , MOV.DEPOMOVI, "
    SQLX$ = SQLX$ + " MOV.CodiMovi , "
    SQLX$ = SQLX$ + "MONEDA ="
    SQLX$ = SQLX$ + " case"
    SQLX$ = SQLX$ + " WHEN MOV.moneval < '2' THEN '$'"
    SQLX$ = SQLX$ + " WHEN MOV.MONEVAL = '2' THEN 'U$S'"
    SQLX$ = SQLX$ + " End "
    
    SQLX$ = SQLX$ + " ,sum(MOV.CANTSALID)as sale"
    SQLX$ = SQLX$ + " , sum(MOV.CANTINGRE) as ingre"
    If SOLOSALIDA% > 0 Then
      SQLX$ = SQLX$ + " ,sum(MOV.CantSalid * MOV.ValoUnit) as importe"
    Else
      SQLX$ = SQLX$ + " ,sum(MOV.CantSalid * MOV.ValoUnit) - SUM(MOV.cantingre * MOV.ValoUnit)  as importe"
    End If
    SQLX$ = SQLX$ + " , MOV.Cod_Inte,MOV.moneval"
    SQLX$ = SQLX$ + " FROM MOVISTO MOV   INNER JOIN MASTER MAS ON MOV.COD_INTE = MAS.CODINT "
    SQLX$ = SQLX$ + " WHERE MOV.FechMov >= '" & FETEXCOR1$(Des%) & "' AND MOV.FechMov <= '" & FETEXCOR1$(Has%) & "'  and MOV.Cod_Inte >0"
    
    SQLX$ = SQLX$ + " AND MOV.COD_EXTE >= '" & CODI1$ & "' "
    SQLX$ = SQLX$ + " AND MOV.COD_EXTE <= '" & CODI2$ & "' "
    If TMAT% > 0 Then
         SQLX$ = SQLX$ + " AND MAS.Timate = " & TMAT%
    End If
    If DEP% > 0 Then
         SQLX$ = SQLX$ + " AND MOV.DEPOMOVI = " & DEP%
    End If
    If SOLOSALIDA% > 0 Then
        SQLX$ = SQLX$ + " AND MOV.CantSalid > 0"
    End If
    SQLX$ = SQLX$ + SQLCODMOV$
    
    SQLX$ = SQLX$ + " group by  YEAR(MOV.fechmov), MONTH(MOV.fechmov), MOV.COD_EXTE, MOV.DEPOMOVI, "
    SQLX$ = SQLX$ + " MOV.CodiMovi, "
    SQLX$ = SQLX$ + " MOV.moneval,MOV.COD_INTE"
    
    SQLX$ = SQLX$ + " order by  YEAR(MOV.fechmov), MONTH(MOV.fechmov)"
    SQLX$ = SQLX$ + ", MOV.COD_INTE "
    
    Set RS_FUN = READSET(SQLX$, , 1)
    Set RS_ESTADISTICA_CONSUMO_AGRUPADA_X_MES = RS_FUN
    Set RS_FUN = Nothing
    
End Function
Public Function RS_CONDIPAGO() As ADODB.Recordset
  'CARGA LA CONDICIONES DE PAGO DE LA TABLA SQL ESTO ESTA HECHO PARA ENVAPLAST QUE LEE DE WALDBOTT Y LO GUARDA EN EL FLEXOFT
  Dim RS_FUN As ADODB.Recordset
  Set RS_FUN = New ADODB.Recordset
  SQLX$ = "SELECT CODICONPAG, REFERENCIA FROM LICONPAG WITH(NOLOCK) ORDER BY CODICONPAG ASC"
  Set RS_FUN = READSET(SQLX$)
  Set RS_CONDIPAGO = RS_FUN
  Set RS_FUN = Nothing
End Function

Public Function RS_MasterComp(SOLOACTIVOS%, Optional GRACT%, Optional CONCOSTO%) As ADODB.Recordset
  Dim RS_FUN As ADODB.Recordset
  Set RS_FUN = New ADODB.Recordset
  SQLX$ = "SELECT CODINT,Codigo, Descripcion,GR_ACTU,COLOR,MONECOS"
  If CONCOSTO% Then SQLX$ = SQLX$ + " ,COSUNI "
  SQLX$ = SQLX$ + " FROM MASTER WITH(NOLOCK) "
  SQLX$ = SQLX$ + " WHERE CODINT > 0"
  If SOLOACTIVOS% > 0 Then SQLX$ = SQLX$ + " AND STATUS >=0 "
  If GRACT% > 0 Then SQLX$ = SQLX$ + " AND GR_ACTU = " & GRACT%
  Set RS_FUN = READSET(SQLX$)
  Set RS_MasterComp = RS_FUN
  
  Set RS_FUN = Nothing
End Function
Public Function RS_GRUPOACTU() As ADODB.Recordset
  Dim RS_FUN As ADODB.Recordset
  Set RS_FUN = New ADODB.Recordset
  SQLX$ = "SELECT CAMPO1, CAMPO2 FROM DATASQL WITH(NOLOCK) WHERE LISTA  = 'GRUPOACTU'"
  Set RS_FUN = READSET(SQLX$)
  Set RS_GRUPOACTU = RS_FUN
  Set RS_FUN = Nothing
End Function

Public Function RS_BOLRECEPEN() As ADODB.Recordset
  Dim RS_FUN As ADODB.Recordset
  Set RS_FUN = New ADODB.Recordset
  SQLX$ = "SELECT NUMEBORE, FECHRECEP, NUMREM_PROV, RASO_PROV FROM BOLRECEP WITH(NOLOCK) WHERE STATUS <= 1 GROUP BY NUMEBORE, FECHRECEP, NUMREM_PROV, RASO_PROV ORDER BY NumeBoRe DESC"
  Set RS_FUN = READSET(SQLX$)
  Set RS_BOLRECEPEN = RS_FUN
  Set RS_FUN = Nothing
End Function
Public Function RS_ZELUSSLCT() As ADODB.Recordset
  Dim RS_FUN As ADODB.Recordset
  Set RS_FUN = New ADODB.Recordset
  SQLX$ = "SELECT CAMPO1, CAMPO3, CAMPO5 FROM DATASQL WITH(NOLOCK)  WHERE LISTA = 'USUARIOSQL'"
  Set RS_FUN = READSET(SQLX$)
  Set RS_ZELUSSLCT = RS_FUN
  Set RS_FUN = Nothing
End Function

Public Function RS_LISTA_DE_OT() As ADODB.Recordset
  Dim RS_FUN As ADODB.Recordset
  Set RS_FUN = New ADODB.Recordset
  SQLX$ = "SELECT Nume_OT, Referencia FROM CORTRAB WITH(NOLOCK) ORDER BY Nume_OT ASC"
  Set RS_FUN = READSET(SQLX$)
  Set RS_LISTA_DE_OT = RS_FUN
  Set RS_FUN = Nothing
End Function
'*****************************************************'
Public Function RS_TAREACT() As ADODB.Recordset
  Dim RS_FUN As ADODB.Recordset
  Set RS_FUN = New ADODB.Recordset
  SQLX$ = "SELECT Nume_Tar, Des_Corta FROM TARINTER WITH(NOLOCK) WHERE Nume_Tar > '0' ORDER BY Nume_Tar ASC"
  Set RS_FUN = READSET(SQLX$)
  Set RS_TAREACT = RS_FUN
  Set RS_FUN = Nothing
End Function
Public Function RS_Solicitud_COT_Fallidas() As ADODB.Recordset
  Dim RS_FUN As ADODB.Recordset
  Set RS_FUN = New ADODB.Recordset
  SQLX$ = "SELECT ENCACOT.ENCACOT_ID, ENCACOT.FechMov FROM ENCACOT WITH(NOLOCK)  INNER JOIN DETACOT ON ENCACOT.ENCACOT_ID = DETACOT.ID_ENCA  WHERE (ENCACOT.COT = ''  OR ENCACOT.COT IS NULL) ORDER BY ENCACOT_ID "
  Set RS_FUN = READSET(SQLX$)
  Set RS_Solicitud_COT_Fallidas = RS_FUN
  Set RS_FUN = Nothing
End Function
Public Function RS_Solicitud_COT_OK(Optional LETRA$) As ADODB.Recordset
  Dim RS_FUN As ADODB.Recordset
  Set RS_FUN = New ADODB.Recordset
  'TRAE TODOS LOS QUE HAN OBTENIDO COT
  SQLX$ = "SELECT D.ID_ENCA ,D.FECHMOV,D.Nume_Clie ,DEU.Raso_Cli ,D.DOPRILAR,E.COT FROM DETACOT D WITH(NOLOCK) INNER JOIN ENCACOT E ON D.ID_ENCA = E.ENCACOT_ID"
  SQLX$ = SQLX$ + " INNER JOIN Deudor12 DEU ON  DEU.Nume_Cli = D.Nume_Clie  WHERE (E.COT <> '' AND E.COT  IS NoT NULL)"
  If LETRA$ = "D" Then
    SQLX$ = SQLX$ + " AND CANTINGRE > 0 "
  Else
    SQLX$ = SQLX$ + " AND CANTSALID > 0 "
  End If
  SQLX$ = SQLX$ + "GROUP BY D.ID_ENCA ,D.FECHMOV,D.Nume_Clie ,DEU.Raso_Cli ,D.DOPRILAR,E.COT"
  SQLX$ = SQLX$ + " ORDER BY D.ID_ENCA"
  Set RS_FUN = READSET(SQLX$)
  Set RS_Solicitud_COT_OK = RS_FUN
  Set RS_FUN = Nothing
End Function

Public Function RS_COTSLECCIONADO(NROSEL&) As ADODB.Recordset
  Dim RS_FUN As ADODB.Recordset
  Set RS_FUN = New ADODB.Recordset
  SQLX$ = "SELECT Cod_Inte, Cod_Exte, Descrip, CantSalid, ValoUnit, PROVINCIA, PATENTE, DoPriLar, NUMPVTA, Nume_Clie, LOCALITRANSP, LOCALIDAD_ENTR, FeFormCot, FechMov, DOPRINUM, COPOSTAL, CUITTRANSP, DOM_ENTRE, DOMITRANSP,DNITRANS,FEREMITO FROM DETACOT WITH(NOLOCK) WHERE ID_ENCA = " & NROSEL&
  Set RS_FUN = READSET(SQLX$)
  Set RS_COTSLECCIONADO = RS_FUN
  Set RS_FUN = Nothing
End Function
Public Function RS_USUARIOYMAQUINA() As ADODB.Recordset
  Dim RS_FUN As ADODB.Recordset
  Set RS_FUN = New ADODB.Recordset
  SQLX$ = "SELECT NUSER, NOMBRE, IDMAQUINA FROM USUARIOSSQL WITH(NOLOCK)"
  Set RS_FUN = READSET(SQLX$)
  Set RS_USUARIOYMAQUINA = RS_FUN
  Set RS_FUN = Nothing
End Function


Public Function RS_PEDIDOS1() As ADODB.Recordset
  Dim RS_FUN As ADODB.Recordset
  Set RS_FUN = New ADODB.Recordset
  SQLX$ = "SELECT CodiExt, Descrip, CanPed, PreList FROM PEDDETA WITH(NOLOCK)"
  Set RS_FUN = READSET(SQLX$)
  Set RS_PEDIDOS1 = RS_FUN
  Set RS_FUN = Nothing
End Function
Public Function RS_PEDIDOS_TODOS() As ADODB.Recordset
  Dim RS_FUN As ADODB.Recordset
  Set RS_FUN = New ADODB.Recordset
  SQLX$ = "SELECT Nroped,Referencia FROM PEDENCA WITH(NOLOCK) order by nroped desc"
  Set RS_FUN = READSET(SQLX$)
  Set RS_PEDIDOS_TODOS = RS_FUN
  Set RS_FUN = Nothing
End Function
Public Function RS_DEPURACOMP(FECHOYTOLERA) As ADODB.Recordset
  Dim RS_FUN As ADODB.Recordset
  Set RS_FUN = New ADODB.Recordset

    SQLX$ = "Select NUME_OCOM,Npro_Ocom,RaSo_Prov,Cod_Exte,Descritem,FEMI_OCOM,Fentre_Sol,Cant_Ocom,Cant_Rec,SUM(Cant_Ocom-Cant_Rec) AS SALDO,NuOrd_Item  from OCOMDETA WITH(NOLOCK)"
    SQLX$ = SQLX$ + " Where Stat_Ocom < 8 "
    SQLX$ = SQLX$ + " AND Stat_Ocom <> -1 "
    SQLX$ = SQLX$ + " AND Cant_Rec < Cant_Ocom - 0.005 "
    SQLX$ = SQLX$ + " AND Fentre_Sol < '" & FECHOYTOLERA & "'"
    SQLX$ = SQLX$ + " group by NUME_OCOM,Npro_Ocom,RaSo_Prov,Cod_Exte,Descritem,FEMI_OCOM,Fentre_Sol,Cant_Ocom,Cant_Rec,NuOrd_Item"
    SQLX$ = SQLX$ + " ORDER BY Nume_Ocom, Cod_Exte, Fentre_Sol "
    Set RS_FUN = READSET(SQLX$)
    Set RS_DEPURACOMP = RS_FUN
    Set RS_FUN = Nothing
End Function
Public Function RS_ALUMNOS(Optional SOLO_DEBAJA%) As ADODB.Recordset
  Dim RS_FUN As ADODB.Recordset
  Set RS_FUN = New ADODB.Recordset
  SQLX$ = "SELECT Nume_Cli, Raso_Cli, DOCU_CLI, Tele_Cli, Mail_Cli,CELU_CLI FROM DEUDOR12 WITH(NOLOCK) "
  If SOLO_DEBAJA% < 1 Then
     SQLX$ = SQLX$ + "  Where Stat_Cli >= 0"
  Else
     SQLX$ = SQLX$ + "  Where Stat_Cli < 0 "
  End If
  
  SQLX$ = SQLX$ + "  ORDER BY Nume_Cli"
  Set RS_FUN = READSET(SQLX$)
  Set RS_ALUMNOS = RS_FUN
  Set RS_FUN = Nothing
End Function
Public Function RS_PROFESORES(Optional SOLO_DEBAJA%) As ADODB.Recordset
  Dim RS_FUN As ADODB.Recordset
  Set RS_FUN = New ADODB.Recordset
  SQLX$ = "SELECT Nume_Cli, Raso_Cli, DOCU_CLI, Tele_Cli, Mail_Cli,CELU_CLI FROM PROFESORES WITH(NOLOCK)  "
  If SOLO_DEBAJA% < 1 Then
     SQLX$ = SQLX$ + "  Where Stat_Cli >= 0"
  Else
     SQLX$ = SQLX$ + "  Where Stat_Cli < 0 "
  End If
  SQLX$ = SQLX$ + " ORDER BY Nume_Cli"
  
  Set RS_FUN = READSET(SQLX$)
  Set RS_PROFESORES = RS_FUN
  Set RS_FUN = Nothing
End Function

Public Function RS_TABLA_PROVINCIA() As ADODB.Recordset
  Dim RS_FUN As ADODB.Recordset
  Set RS_FUN = New ADODB.Recordset
  SQLX$ = "SELECT CODINUM, CODITEX,DESCRI_PROV  FROM TAPROVINCIA WITH(NOLOCK) ORDER BY CODINUM ASC"
  Set RS_FUN = READSET(SQLX$)
  Set RS_TABLA_PROVINCIA = RS_FUN
  Set RS_FUN = Nothing
End Function
    
    Public Function RS_CURSOS() As ADODB.Recordset
      Dim RS_FUN As ADODB.Recordset
      Set RS_FUN = New ADODB.Recordset
      SQLX$ = "SELECT CURSOS_ID,COD_AREA, AREA,DESCRIPCION,ValCurso,ValMatricula,ValDerExamen"
      SQLX$ = SQLX$ + " FROM CURSOS WITH(NOLOCK) ORDER BY COD_AREA,DESCRIPCION ASC"
      Set RS_FUN = READSET(SQLX$)
      Set RS_CURSOS = RS_FUN
      Set RS_FUN = Nothing
    End Function
Public Function RS_CURSOS_ACT(Optional NUMCURSO&, Optional SOLOBAJA%, Optional MAS28DIAS%) As ADODB.Recordset
  Dim RS_FUN As ADODB.Recordset
  Set RS_FUN = New ADODB.Recordset 'EL NUMCURSO ES COMO LA FAMILIA
  SQLX$ = "SELECT * FROM MASTER AS M WITH(NOLOCK) LEFT JOIN PROFESORES AS P ON P.NUME_CLI = M.ID_PROFESOR Where M.CODINT > 0"
  If NUMCURSO& > 0 Then SQLX$ = SQLX$ + "  AND M.ID_CURSO = " & NUMCURSO&
  If SOLOBAJA% > 0 Then
    SQLX$ = SQLX$ + "  AND M.STATUS < 0"
  Else
    SQLX$ = SQLX$ + " AND (M.STATUS >= 0 OR M.STATUS IS NULL)"
  End If
  If MAS28DIAS% < 1 Then
    FF% = HOY(HO$)
    FF% = DIANTE(FF%, 28)
    fech = FETEXCOR1$(FF%)
    SQLX$ = SQLX$ + " AND FeIniCurs >  '" & fech & "'"
  End If
  SQLX$ = SQLX$ + " ORDER BY FeIniCurs,H_DESDE_FN DESC"
  Set RS_FUN = READSET(SQLX$)
  Set RS_CURSOS_ACT = RS_FUN
  Set RS_FUN = Nothing
End Function
Public Function RS_CURSOS_ACT_2(Optional NUMCURSO&, Optional SOLOBAJA%, Optional MAS28DIAS%) As ADODB.Recordset
  'PARA MOSTRAR EN SELECHO
  Dim RS_FUN As ADODB.Recordset
  Set RS_FUN = New ADODB.Recordset 'EL NUMCURSO ES COMO LA FAMILIA

  SQLX$ = "SELECT  M.DESCRICURSO,M.FeIniCurs,M.FeFinCurs,M.H_DESDE,M.H_HASTA,M.CODINT FROM MASTER AS M WITH(NOLOCK) LEFT JOIN PROFESORES AS P ON P.NUME_CLI = M.ID_PROFESOR Where M.CODINT > 0"
  If SOLOBAJA% > 0 Then
    SQLX$ = SQLX$ + "  AND M.STATUS < 0"
  Else
    SQLX$ = SQLX$ + " AND (M.STATUS >= 0 OR M.STATUS IS NULL)"
  End If
  If NUMCURSO& > 0 Then
    SQLX$ = SQLX$ + "  AND COD_AREA = " & NUMCURSO&
  End If
  If MAS28DIAS% < 1 Then
    FF% = HOY(HO$)
    FF% = DIANTE(FF%, 28)
    fech = FETEXCOR1$(FF%)
    SQLX$ = SQLX$ + " AND FeIniCurs >  '" & fech & "'"
  End If
  SQLX$ = SQLX$ + " ORDER BY FeIniCurs DESC"
  Set RS_FUN = READSET(SQLX$)
  Set RS_CURSOS_ACT_2 = RS_FUN
  Set RS_FUN = Nothing
End Function

    Public Function RS_CURSOS_DEL_ALUMNO(NC) As ADODB.Recordset
      Dim RS_FUN As ADODB.Recordset
      Set RS_FUN = New ADODB.Recordset
      SQLX$ = "SELECT CURSOS_ID,FECHA,DESCRIPCION "
      SQLX$ = SQLX$ + " FROM ALUMNO_CURSO  WITH(NOLOCK) "
      SQLX$ = SQLX$ + " WHERE NUME_CLI = " & NC
      SQLX$ = SQLX$ + " ORDER BY FECHA DESC"
      Set RS_FUN = READSET(SQLX$)
      Set RS_CURSOS_DEL_ALUMNO = RS_FUN
      Set RS_FUN = Nothing
    End Function

Public Function ALUMNO_x_DIATEX(DIAZ$) As ADODB.Recordset
    Dim RS_FUN As ADODB.Recordset
    Set RS_FUN = New ADODB.Recordset
    SQLX$ = " select a.nume_cli, m.codint,m.descricurso,m.feinicurs,m.fefincurs,m.cantmeses,m.h_desde,m.h_hasta,m.desc_dia1,m.desc_dia2,"
    SQLX$ = SQLX$ + " id_profesor,id_aula from alumno_curso a inner join master m on a.cursos_id = m.codint"
    SQLX$ = SQLX$ + " Where datediff(D, m.FeFinCurs, GETDATE()) <= 1 AND (M.STATUS >= 0 OR M.STATUS IS NULL) "
    SQLX$ = SQLX$ + " and desc_dia1 = '" & DIAZ$ & "' or desc_dia2 = '" & DIAZ$ & "'"
    SQLX$ = SQLX$ + " order by a.nume_cli,h_desde_fn"
    Set RS_FUN = READSET(SQLX$)
    Set ALUMNO_x_DIATEX = RS_FUN
    Set RS_FUN = Nothing
End Function

Public Function DIASSEMANATEX() As ADODB.Recordset
    Dim RS_FUN As ADODB.Recordset
    Set RS_FUN = New ADODB.Recordset
    SQLX$ = " select CAMPO1,CAMPO2"
    SQLX$ = SQLX$ + " from DATASQL  WITH(NOLOCK)"
    SQLX$ = SQLX$ + " Where LISTA  = 'DIAS_SEMANA' "
    SQLX$ = SQLX$ + " order by CAMPO1"
    Set RS_FUN = READSET(SQLX$)
    Set DIASSEMANATEX = RS_FUN
    Set RS_FUN = Nothing
End Function


Public Function RS_ArtClasificado(familia) As ADODB.Recordset
    Dim RS_FUN As ADODB.Recordset
    Set RS_FUN = New ADODB.Recordset

    SQLX$ = "Select CODINT,CODIGO,DESCRIPCION  from MASTER WITH(NOLOCK)"
    SQLX$ = SQLX$ + " Where Familia = '" & familia & "'"
    Set RS_FUN = READSET(SQLX$)
    Set RS_ArtClasificado = RS_FUN
    Set RS_FUN = Nothing
End Function
Public Function RS_MOTIVOS() As ADODB.Recordset
  Dim RS_FUN As ADODB.Recordset
  Set RS_FUN = New ADODB.Recordset
  SQLX$ = "SELECT CAMPO1, CAMPO2 FROM DATASQL WITH(NOLOCK) WHERE LISTA  = 'MOTIVOS' ORDER BY CAMPO1"
  Set RS_FUN = READSET(SQLX$)
  Set RS_MOTIVOS = RS_FUN
  Set RS_FUN = Nothing
End Function

Public Function RS_TECNICOS_MDT() As ADODB.Recordset
  Dim RS_FUN As ADODB.Recordset
  Set RS_FUN = New ADODB.Recordset
  SQLX$ = "SELECT CAMPO1, CAMPO2 FROM DATASQL WITH(NOLOCK) WHERE LISTA  = 'TECNICOS-MDT' ORDER BY CAMPO2"
  Set RS_FUN = READSET(SQLX$)
  Set RS_TECNICOS_MDT = RS_FUN
  Set RS_FUN = Nothing
End Function
Public Function RS_MEDIOSCOBRANZA_MDT() As ADODB.Recordset
  Dim RS_FUN As ADODB.Recordset
  Set RS_FUN = New ADODB.Recordset
  SQLX$ = "SELECT CAMPO1, CAMPO2,CAMPO3 FROM DATASQL WITH(NOLOCK) WHERE LISTA  = 'MEDIOS COBRANZA-MDT' ORDER BY CAMPO1"
  Set RS_FUN = READSET(SQLX$)
  Set RS_MEDIOSCOBRANZA_MDT = RS_FUN
  Set RS_FUN = Nothing
End Function
Public Function RS_NEGATIVOS_MDT(PANOLFACTU%) As ADODB.Recordset
  Dim RS_FUN As ADODB.Recordset
  Set RS_FUN = New ADODB.Recordset
  SQLX$ = "SELECT COD_INTE,COD_EXTE,DESCRIP,SUM(CANTINGRE-CANTSALID) AS STOCKK FROM MOVISTO MO WITH(NOLOCK)"
  SQLX$ = SQLX$ + " INNER JOIN MASTER MA ON MA.CODINT = MO.COD_INTE "
  SQLX$ = SQLX$ + " Where DEPOMOVI = " & PANOLFACTU% & " AND (MA.TRAZABLE < 1 OR MA.TRAZABLE IS NULL)"
  SQLX$ = SQLX$ + " GROUP BY COD_INTE,COD_EXTE,DESCRIP HAVING SUM(CANTINGRE-CANTSALID) < 0 ORDER BY COD_INTE "
  Set RS_FUN = READSET(SQLX$)
  Set RS_NEGATIVOS_MDT = RS_FUN
  Set RS_FUN = Nothing
End Function
Public Function RS_RESPONSABLES_CME() As ADODB.Recordset 'CENTRO MEDICO ESCUELA
  Dim RS_FUN As ADODB.Recordset
  Set RS_FUN = New ADODB.Recordset
  SQLX$ = "SELECT CAMPO1, CAMPO2 FROM DATASQL WITH(NOLOCK) WHERE LISTA  = 'RESPONSABLES' ORDER BY CAMPO1"
  Set RS_FUN = READSET(SQLX$)
  Set RS_RESPONSABLES_CME = RS_FUN
  Set RS_FUN = Nothing
End Function
Public Function RS_RESERVAS_ENV(Optional PORARTICULO%) As ADODB.Recordset 'RESERVAS POR OF
  Dim RS_FUN As ADODB.Recordset
  Set RS_FUN = New ADODB.Recordset

  SQLX$ = " SELECT COD_EXTE, DESCRIP, IDSUBOR, FechRes, CantRes, CANTCONS, sum(CantRes - CANTCONS) as Saldo, IdenLote, Deposito, UsoUnit FROM RESERVA WITH(NOLOCK) "
  SQLX$ = SQLX$ + " WHERE COD_INTE > 0 AND CantRes-CantCons > 0 group by COD_EXTE, DESCRIP, IDSUBOR, FechRes, CantRes, CANTCONS,  IdenLote, Deposito, UsoUnit"
  If PORARTICULO% > 0 Then
     SQLX$ = SQLX$ + " ORDER BY COD_EXTE "
  Else
     SQLX$ = SQLX$ + " ORDER BY IdSubOr "
  End If
  
  Set RS_FUN = READSET(SQLX$)
  Set RS_RESERVAS_ENV = RS_FUN
  Set RS_FUN = Nothing
End Function
Public Function RS_RESERVAS_C_CTE_ENV(Optional PORARTICULO%) As ADODB.Recordset 'RESERVAS POR OF CON N DE CLIENTE
  Dim RS_FUN As ADODB.Recordset
  Set RS_FUN = New ADODB.Recordset

  SQLX$ = " SELECT R.COD_EXTE, R.DESCRIP, R.IDSUBOR, R.FechRes, R.CantRes, R.CANTCONS, sum(R.CantRes - R.CANTCONS) as Saldo,"
  SQLX$ = SQLX$ + " R.IdenLote, R.Deposito, R.UsoUnit,D.RASO_CLI FROM RESERVA R WITH(NOLOCK) "
  SQLX$ = SQLX$ + " INNER JOIN ORDEFABR O ON O.IDSUBOR = R.IDSUBOR INNER JOIN DEUDOR12 D ON D.NUME_CLI = O.NUME_CLI"
  SQLX$ = SQLX$ + " WHERE R.COD_INTE > 0 AND R.CantRes-R.CantCons > 0 "
  
  SQLX$ = SQLX$ + " group by R.COD_EXTE, R.DESCRIP, R.IDSUBOR, R.FechRes, R.CantRes, R.CANTCONS,  R.IdenLote, R.Deposito, R.UsoUnit,D.RASO_CLI"
  If PORARTICULO% > 0 Then
     SQLX$ = SQLX$ + " ORDER BY R.COD_EXTE "
  Else
     SQLX$ = SQLX$ + " ORDER BY R.IdSubOr "
  End If
  
  Set RS_FUN = READSET(SQLX$)
  Set RS_RESERVAS_C_CTE_ENV = RS_FUN
  Set RS_FUN = Nothing
End Function

Public Function RS_RESERVAS_ENVGR() As ADODB.Recordset 'GENERARL DE RESERVAS X OF
  Dim RS_FUN As ADODB.Recordset
  Set RS_FUN = New ADODB.Recordset

  SQLX$ = " SELECT IDSUBOR,COD_EXTE, DESCRIP, sum(CantRes - CANTCONS) as Saldo,  UsoUnit FROM RESERVA WITH(NOLOCK) "
  SQLX$ = SQLX$ + " WHERE COD_INTE > 0 AND CantRes-CantCons > 0 group by IDSUBOR,COD_EXTE, DESCRIP, UsoUnit"
  SQLX$ = SQLX$ + " ORDER BY IdSubOr,COD_EXTE "
  Set RS_FUN = READSET(SQLX$)
  Set RS_RESERVAS_ENVGR = RS_FUN
  Set RS_FUN = Nothing
End Function

Public Function RS_PREVTA(LP%, GRACT%) As ADODB.Recordset
  Dim RS_FUN As ADODB.Recordset
  Set RS_FUN = New ADODB.Recordset
   
  SQLX$ = "SELECT p.Cod_Pieza,m.codigo, m.Descripcion,M.GR_ACTU,COLOR,P.Precio_Unitario"
  SQLX$ = SQLX$ + " FROM PREVENTA P WITH(NOLOCK) INNER JOIN MASTER M ON M.CODINT = P.COD_PIEZA"
  SQLX$ = SQLX$ + " WHERE P.Cod_Lista = " & LP%
  SQLX$ = SQLX$ + " AND M.GR_ACTU = " & GRACT%
  Set RS_FUN = READSET(SQLX$)
  Set RS_PREVTA = RS_FUN
  Set RS_FUN = Nothing
End Function
Public Function RS_LISTASPRECIOS()
   Dim RS_FUN As ADODB.Recordset
   Set RS_FUN = New ADODB.Recordset
   SQLX$ = "SELECT Cod_Lista,DENOM_LISTA FROM LISTAPRE WITH(NOLOCK) ORDER BY Cod_Lista ASC"
   Set RS_FUN = READSET(SQLX$)
   Set RS_LISTASPRECIOS = RS_FUN
   Set RS_FUN = Nothing

End Function
Public Function RS_LISTA_DE_TÉCNICOS() As ADODB.Recordset
  Dim RS_FUN As ADODB.Recordset
  Set RS_FUN = New ADODB.Recordset
  SQLX$ = "SELECT CAMPO1, CAMPO2,CAMPO3,CAMPO4,CAMPO5,CAMPO6 FROM DATASQL WITH(NOLOCK)  WHERE LISTA = 'TECNIEXT' "
  Set RS_FUN = READSET(SQLX$)
  Set RS_LISTA_DE_TÉCNICOS = RS_FUN
  Set RS_FUN = Nothing
End Function
Public Function RS_PERSOCONTROL_EYM() As ADODB.Recordset
  Dim RS_FUN As ADODB.Recordset
  Set RS_FUN = New ADODB.Recordset
  SQLX$ = "SELECT CAMPO1, CAMPO2 FROM DATASQL WITH(NOLOCK) WHERE LISTA  = 'PERSOCON' ORDER BY CAMPO2"
  Set RS_FUN = READSET(SQLX$)
  Set RS_PERSOCONTROL_EYM = RS_FUN
  Set RS_FUN = Nothing
End Function
Public Function RS_PERSOAPRO_EYM() As ADODB.Recordset
  Dim RS_FUN As ADODB.Recordset
  Set RS_FUN = New ADODB.Recordset
  SQLX$ = "SELECT CAMPO1, CAMPO2 FROM DATASQL WITH(NOLOCK) WHERE LISTA  = 'PERSOAPRO' ORDER BY CAMPO2"
  Set RS_FUN = READSET(SQLX$)
  Set RS_PERSOAPRO_EYM = RS_FUN
  Set RS_FUN = Nothing
End Function
Public Function RS_CONFIGURACION_DE_FORMATO_DE_NSERIE() As ADODB.Recordset
  Dim RS_FUN As ADODB.Recordset
  Set RS_FUN = New ADODB.Recordset
  SQLX$ = "SELECT TIP_DATO,REFER_DATO, CANTICARACT, APARTIR, EXTREMO, COMPLETACON,MASCARA,MEDIDAS FROM FORMANSERIE WITH(NOLOCK)  WHERE CORRESPONDE_A = 'NUMEROS DE SERIE'"
  Set RS_FUN = READSET(SQLX$)
  Set RS_CONFIGURACION_DE_FORMATO_DE_NSERIE = RS_FUN
  Set RS_FUN = Nothing
End Function
Public Function RS_CONTROLES_EYM() As ADODB.Recordset
  Dim RS_FUN As ADODB.Recordset
  Set RS_FUN = New ADODB.Recordset
  SQLX$ = "SELECT CAMPO1, CAMPO2 FROM DATASQL WITH(NOLOCK) WHERE LISTA  = 'CONFICONTROL' ORDER BY CAMPO1"
  Set RS_FUN = READSET(SQLX$)
  Set RS_CONTROLES_EYM = RS_FUN
  Set RS_FUN = Nothing
End Function
Public Function RS_CONSUMOS(CONDI$) As ADODB.Recordset
  Dim RS_FUN As ADODB.Recordset
  Set RS_FUN = New ADODB.Recordset
  SQLX$ = "SELECT Cod_Exte, Descrip,CantSalid,IdenLote,OrdConsum FROM MOVISTO MO WITH(NOLOCK)"
  SQLX$ = SQLX$ + " INNER JOIN MASTER MA ON MA.CODINT = MO.COD_INTE Where " & CONDI$
  Set RS_FUN = READSET(SQLX$)
  Set RS_CONSUMOS = RS_FUN
  Set RS_FUN = Nothing
End Function
Public Function RS_RESERVAS(CONDI$) As ADODB.Recordset
  Dim RS_FUN As ADODB.Recordset
  Set RS_FUN = New ADODB.Recordset
  SQLX$ = "SELECT Cod_Exte, Descrip,RE.CANTRES,RE.CANTCONS,IdenLote,OrdConsum FROM RESERVA RE WITH(NOLOCK)"
  SQLX$ = SQLX$ + " INNER JOIN MASTER MA ON MA.CODINT = RE.COD_INTE Where " & CONDI$
  Set RS_FUN = READSET(SQLX$)
  Set RS_RESERVAS = RS_FUN
  Set RS_FUN = Nothing
End Function
Public Function RS_NSERIE(CONDI$) As ADODB.Recordset
  Dim RS_FUN As ADODB.Recordset
  Set RS_FUN = New ADODB.Recordset
  SQLX$ = "SELECT NUMEROSERIE,FECHALTA FROM TANUMSE WITH(NOLOCK) WHERE " & CONDI$
  Set RS_FUN = READSET(SQLX$)
  Set RS_NSERIE = RS_FUN
  Set RS_FUN = Nothing
End Function

Public Function RS_TODARECE() As ADODB.Recordset
  'LEVANTA TODAS LOS DETALLES DE UNA RECEPCION
  Dim RS_FUN As ADODB.Recordset
  Set RS_FUN = New ADODB.Recordset
  SQLX$ = "SELECT  NUMEBORE,NumProv, RaSo_Prov, FechRecep FROM BOLRECEP WITH(NOLOCK)"
  SQLX$ = SQLX$ + " GROUP BY  NUMEBORE,NumProv, RaSo_Prov, FechRecep ORDER BY NUMEBORE DESC"
  Set RS_FUN = READSET(SQLX$)
  Set RS_TODARECE = RS_FUN
  Set RS_FUN = Nothing
End Function
Public Function RS_TODAOF() As ADODB.Recordset
  'LEVANTA TODAS LOS DETALLES DE UNA RECEPCION
  Dim RS_FUN As ADODB.Recordset
  Set RS_FUN = New ADODB.Recordset
  SQLX$ = "SELECT  NUMEORFA,IDSUBOR, COD_INTE,COD_EXTE,DESCRIP ,CANPROY,FechGen FROM ORDEFABR WITH(NOLOCK)"
  SQLX$ = SQLX$ + " ORDER BY NUMEORFA DESC"
  Set RS_FUN = READSET(SQLX$)
  Set RS_TODAOF = RS_FUN
  Set RS_FUN = Nothing
End Function

Public Function RS_DETARECE(NBORE, OP%, Optional IncluirCantBultos%) As ADODB.Recordset
  'Call CREATABLA_EQUIPOS
  'LEVANTA TODAS LOS DETALLES DE UNA RECEPCION

  Dim RS_FUN As ADODB.Recordset
  Set RS_FUN = New ADODB.Recordset
  SQLX$ = "SELECT B.IdenLote, B.Cod_Exte, B.DescripItem, B.CantNom,B.Cod_Inte,M.Revision,D.DESV_TRANS,B.LoteCol_Prov"
  If IncluirCantBultos% > 0 Then SQLX$ = SQLX$ + ", B.CantidadBultos "
  SQLX$ = SQLX$ + " FROM BOLRECEP B WITH(NOLOCK)"
  SQLX$ = SQLX$ + " INNER JOIN MASTER M ON B.Cod_Inte = M.CODINT "
  SQLX$ = SQLX$ + " LEFT JOIN DES_TRANS D ON B.Cod_Inte   =  D.Cod_Inte  AND D.IDENLOTE = B.IDENLOTE  "
  SQLX$ = SQLX$ + " WHERE B.NUMEBORE = " & NBORE & " AND B.CODIEMPR = " & CodiEmp%
'  SQLX$ = SQLX$ + " AND D.IDENLOTE = O.IdSubOr"
  If OP% = 1 Then SQLX$ = SQLX$ + " AND M.TRAZABLE = 1"
  'If OP% = 2 Then SQLX$ = SQLX$ + " AND M.NOIMPRIMIBLE = 1"
  SQLX$ = SQLX$ + " ORDER BY B.COD_EXTE"
  Set RS_FUN = READSET(SQLX$)
  Set RS_DETARECE = RS_FUN
  Set RS_FUN = Nothing
End Function


Public Function RS_DETAORFAB(OF) As ADODB.Recordset
  'LEVANTA TODAS LOS DETALLES DE UNA FABRICACION
  Dim RS_FUN As ADODB.Recordset
  Set RS_FUN = New ADODB.Recordset
 
  'CAMPOS$ = "N°OF/A0;O.Fab/A16;CODIGO/A16;DESCRIPCION/A24;CANTIDAD/F10.1;CI/A0;REVISION/A12;D.TRANS./F8"
  SQLX$ = "SELECT O.NUMEORFA,O.IDSUBOR, O.COD_EXTE, O.DESCRIP, O.CanPROY,O.COD_INTE,M.Revision,D.DESV_TRANS FROM ORDEFABR O WITH(NOLOCK)"
  SQLX$ = SQLX$ + " INNER JOIN MASTER M ON O.Cod_Inte = M.CODINT "
  SQLX$ = SQLX$ + " LEFT JOIN DES_TRANS D ON O.Cod_Inte   =  D.Cod_Inte AND D.IDENLOTE = O.IdSubOr"
  SQLX$ = SQLX$ + " Where O.CODIEMPR = " & CodiEmp% & " And O.NUMEORFA = " & OF
  'SQLX$ = SQLX$ + " AND D.IDENLOTE = O.IdSubOr"
  Set RS_FUN = READSET(SQLX$)
  Set RS_DETAORFAB = RS_FUN
  Set RS_FUN = Nothing
End Function
Public Function RS_DETAORFAB_CAMBACERES(NORFA$) As ADODB.Recordset
  'LEVANTA TODAS LOS DETALLES DE UNA FABRICACION
  Dim RS_FUN As ADODB.Recordset
  Set RS_FUN = New ADODB.Recordset

  SQLX$ = "SELECT IDENLOTE,COD_EXTE,DESCRIP,CANTINGRE,METROS,PESOKG,COD_INTE FROM MOVISTO WITH(NOLOCK)"
  SQLX$ = SQLX$ + " Where CODIEMPR = " & CodiEmp% & " And DOPRILAR = '" & NORFA$ & "'"
  SQLX$ = SQLX$ + " AND CANTINGRE > 0 AND  CODIMOVI = 'OF'"
  SQLX$ = SQLX$ + " ORDER BY IDENLOTE"
  

  Set RS_FUN = READSET(SQLX$)
  Set RS_DETAORFAB_CAMBACERES = RS_FUN
 Set RS_FUN = Nothing
End Function
Public Function RS_SERIEEYM() As ADODB.Recordset
  Dim RS_FUN As ADODB.Recordset
  Set RS_FUN = New ADODB.Recordset
  SQLX$ = "SELECT T.COD_INTE, T.COD_EXTE,T.DESCRIPCION,T.IDSUBOR,T.NumeroSerie, T.FechAlta,T.TRAZAXSERIE_ID FROM TRAZAXSERIE T WITH(NOLOCK) "
  SQLX$ = SQLX$ + " INNER JOIN ORDEFABR O ON O.IDSUBOR = T.IDSUBOR WHERE (T.TRAZABLE = 2 OR T.TRAZABLE = -1)  "
  SQLX$ = SQLX$ + " AND O.StatuOrf < 3 ORDER BY O.IDSUBOR DESC"
  Set RS_FUN = READSET(SQLX$)
  Set RS_SERIEEYM = RS_FUN
  Set RS_FUN = Nothing
End Function
Public Function RS_OF_ABIERTAS(Optional Todas%, Optional INCLUYE_APROBADA%, Optional EXCEPCIONCODIGOFAMILIA$, Optional Mesx%) As ADODB.Recordset
  'si TODAS% > 0 no filtra por status < 3
  'si incluye aprobada trae tambien el campo canapro
  'SI TODAS% = 3  SOLO CERRADAS
  'SI TODAS% < 0  tra etodo
  
  Dim RS_FUN As ADODB.Recordset
  Set RS_FUN = New ADODB.Recordset
  SQLX$ = "SELECT IDSUBOR,Referen,CANPROY "
  If INCLUYE_APROBADA% > 0 Then SQLX$ = SQLX$ + " ,O.CANAPRO,O.ID_ORFA "
  SQLX$ = SQLX$ + " FROM ORDEFABR O WITH(NOLOCK) INNER JOIN MASTER M ON M.CODINT = O.COD_INTE"
  SQLX$ = SQLX$ + " WHERE O.COD_INTE > 0 "
  If Todas% < 1 And Todas% >= 0 Then SQLX$ = SQLX$ + " AND O.StatuOrf < 3 "
  If Todas% > 1 Then SQLX$ = SQLX$ + " AND O.StatuOrf = 3 "
  
  If TRIM$(EXCEPCIONCODIGOFAMILIA$) <> "" Then SQLX$ = SQLX$ + " AND (M.FAMILIA <> '" & EXCEPCIONCODIGOFAMILIA$ & "' OR M.FAMILIA  IS NULL)"
  If Mesx% > 0 And Mesx% <= 12 Then SQLX$ = SQLX$ + " AND MONTH(FechGen) =  " & Mesx%
  
  SQLX$ = SQLX$ + " ORDER BY O.IDSUBOR DESC"
  Set RS_FUN = READSET(SQLX$)
  Set RS_OF_ABIERTAS = RS_FUN
  Set RS_FUN = Nothing
End Function
Public Function RS_OF_ABIERTAS_2(Optional Todas%, Optional INCLUYE_APROBADA%) As ADODB.Recordset
  'si incluye aprobada trae tambien el campo canapro
  
  'ESTA INCLUYE CANTIDAD APROBADA
  Dim RS_FUN As ADODB.Recordset
  Set RS_FUN = New ADODB.Recordset
  SQLX$ = "SELECT IDSUBOR,Referen,CANPROY "
  If INCLUYE_APROBADA% > 0 Then SQLX$ = SQLX$ + " ,CANAPRO,ID_ORFA "
  SQLX$ = SQLX$ + " FROM ORDEFABR WITH(NOLOCK) "
  If Todas% < 1 Then SQLX$ = SQLX$ + " WHERE StatuOrf < 3 "
  SQLX$ = SQLX$ + " ORDER BY IDSUBOR DESC"
  Set RS_FUN = READSET(SQLX$)
  Set RS_OF_ABIERTAS_2 = RS_FUN
  Set RS_FUN = Nothing
  
End Function
Public Function RS_EXCEPCIONES() As ADODB.Recordset
  Dim RS_FUN As ADODB.Recordset
  Set RS_FUN = New ADODB.Recordset
  SQLX$ = "SELECT COD_INTE,COD_EXTE,DESCRIPCION,C_FRMCONTROL,REF_CONTROL,EXCEPCIONES FROM CONTROEYM WITH(NOLOCK) "
  SQLX$ = SQLX$ + " ORDER BY COD_EXTE "
  Set RS_FUN = READSET(SQLX$)
  Set RS_EXCEPCIONES = RS_FUN
  Set RS_FUN = Nothing
End Function
Public Function RS_COD_LLEVAN_CONTROL(TABLA$, Optional CondicionLike$) As ADODB.Recordset
  'TRAE SOLO LOS ITEMS QUE LLEVAN CONTROL FINAL
  'ES UN INNERJOIN DE UNA TABLA DE CONFIGURACION DE DATASQL
  Dim RS_FUN As ADODB.Recordset
  Set RS_FUN = New ADODB.Recordset
  SQLX$ = "SELECT T.NUMEROSERIE,T.COD_EXTE,T.DESCRIPCION,T.IDSUBOR,T.FECHALTA  FROM TRAZAXSERIE T  WITH(NOLOCK) INNER JOIN CONTROEYM C "
  SQLX$ = SQLX$ + " ON C.COD_EXTE= T.COD_EXTE  "
  SQLX$ = SQLX$ + " WHERE T.TRAZABLE = -1 AND C.REF_CONTROL = '" & TABLA$ & "'"
  If TRIM$(CondicionLike$) <> "" Then SQLX$ = SQLX$ + " AND " & CondicionLike$
  SQLX$ = SQLX$ + " ORDER BY FECHALTA DESC"
  Set RS_FUN = READSET(SQLX$)
  Set RS_COD_LLEVAN_CONTROL = RS_FUN
  Set RS_FUN = Nothing
End Function
Public Function RS_COD_LLEVAN_CONTROL_DB_C15(TABLA$) As ADODB.Recordset
  'TRAE SOLO LOS ITEMS QUE LLEVAN CONTROL FINAL
  'ES UN INNERJOIN DE UNA TABLA DE CONFIGURACION DE DATASQL
  Dim RS_FUN As ADODB.Recordset
  Set RS_FUN = New ADODB.Recordset
  SQLX$ = "SELECT T.NUMEROSERIE,T.COD_EXTE,T.DESCRIPCION,T.IDSUBOR,T.FECHALTA  FROM TRAZAXSERIE T  WITH(NOLOCK) INNER JOIN CONTROEYM C "
  SQLX$ = SQLX$ + " ON C.COD_EXTE= T.COD_EXTE  "
  SQLX$ = SQLX$ + " WHERE  C.REF_CONTROL = '" & TABLA$ & "'"
'  SQLX$ = SQLX$ + " AND C.COD_EXTE LIKE 'DB%' AND T.TRAZABLE = 2 " 'ASI ESTABA ANTES AHORA QUE AGREGARON OTRO NIVEL LA CAMBIE A LA DE ABAJO PARA QUE APAREZXCA PARA SELECCIONAR
  SQLX$ = SQLX$ + " AND C.COD_EXTE LIKE 'DB%' AND (T.TRAZABLE = 2 OR T.TRAZABLE = -1)"
  SQLX$ = SQLX$ + " ORDER BY FECHALTA DESC"
  Set RS_FUN = READSET(SQLX$)
  Set RS_COD_LLEVAN_CONTROL_DB_C15 = RS_FUN
  Set RS_FUN = Nothing
End Function

Public Function RS_T_MATERIAL_SQL() As ADODB.Recordset
  Dim RS_FUN As ADODB.Recordset
  Set RS_FUN = New ADODB.Recordset
    SQLX$ = "select CAMPO1,CAMPO2 FROM DATASQL WITH(NOLOCK)"
    SQLX$ = SQLX$ + " WHERE LISTA = 'TABTIMA' ORDER BY CAMPO1"
  Set RS_FUN = READSET(SQLX$)
  Set RS_T_MATERIAL_SQL = RS_FUN
  Set RS_FUN = Nothing
End Function
Public Function RS_CODMODELO(CI1%) As ADODB.Recordset
  Dim RS_FUN As ADODB.Recordset
  Set RS_FUN = New ADODB.Recordset
  SQLX$ = "select CAMPO1,CAMPO2,CAMPO3,CAMPO4 FROM DATASQL WITH(NOLOCK)"
  SQLX$ = SQLX$ + " WHERE LISTA = 'CODMODELO' AND CAMPO1 = '" & CODEXT$(CI1%) & "' ORDER BY CAMPO1"
  Set RS_FUN = READSET(SQLX$)
  Set RS_CODMODELO = RS_FUN
  Set RS_FUN = Nothing
End Function
Public Function RS_BOLRECEDETALLE(NBORE) As ADODB.Recordset
  'devuelve el detalle de una recepcion pasandole el valor numerico de la boleta de recepcion
  Dim RS_FUN As ADODB.Recordset
  Set RS_FUN = New ADODB.Recordset
  SQLX$ = "SELECT B.IdenLote, B.Cod_Exte, B.DescripItem, B.CantNom, B.Cod_Inte FROM BOLRECEP B WITH(NOLOCK)"
  SQLX$ = SQLX$ + " WHERE B.NUMEBORE = " & NBORE
  SQLX$ = SQLX$ + " ORDER BY NumeBoRe DESC"
  Set RS_FUN = READSET(SQLX$)
  Set RS_BOLRECEDETALLE = RS_FUN
  Set RS_FUN = Nothing
End Function
Public Function RS_BoleReceYDespaimpo(NBORE) As ADODB.Recordset
  'devuelve el detalle de una recepcion MAS LOS DATOS DE DESPACHO SE CREO PARA MEDITEA
  Dim RS_FUN As ADODB.Recordset
  Set RS_FUN = New ADODB.Recordset
  SQLX$ = "SELECT B.IdenLote, B.Cod_Exte, B.DescripItem, B.CantNom, B.Cod_Inte,D.FECHA,D.DESPACHO,D.ADUANA"
  SQLX$ = SQLX$ + " FROM BOLRECEP B WITH(NOLOCK) LEFT JOIN DESPAIMPO D"
  SQLX$ = SQLX$ + " ON B.NUMEBORE = D.NUMEBORE AND B.IDENLOTE = D.IDENLOTE"
  SQLX$ = SQLX$ + " WHERE B.NUMEBORE = " & NBORE
  SQLX$ = SQLX$ + " ORDER BY B.NumeBoRe DESC"
  Set RS_FUN = READSET(SQLX$)
  Set RS_BoleReceYDespaimpo = RS_FUN
  Set RS_FUN = Nothing
End Function
Public Function RS_OF_ABYCERR(Optional Todas%) As ADODB.Recordset
  Dim RS_FUN As ADODB.Recordset
  Set RS_FUN = New ADODB.Recordset
  SQLX$ = "SELECT IDSUBOR,Referen,CANPROY,CANAPRO,StatuOrf,FechGen FROM ORDEFABR WITH(NOLOCK) "
  If Todas% < 1 Then SQLX$ = SQLX$ + " WHERE StatuOrf < 3 "
  SQLX$ = SQLX$ + " ORDER BY IDSUBOR DESC"
  Set RS_FUN = READSET(SQLX$)
  Set RS_OF_ABYCERR = RS_FUN
  Set RS_FUN = Nothing
End Function
Public Function RS_OF_PENDIENTES_NSERIE_CON_BR(Optional Todas%) As ADODB.Recordset 'REALIZADA PARA MEDITEA
  Dim RS_FUN As ADODB.Recordset
  Set RS_FUN = New ADODB.Recordset
  SQLX$ = "Select DISTINCT O.IDSUBOR,O.REFEREN,O.CANPROY,B.NUMEBORE,B.RASO_PROV FROM ORDEFABR O  WITH(NOLOCK)"
  SQLX$ = SQLX$ + " INNER JOIN MASTER M ON O.COD_INTE = M.CODINT " 'MASTER PARA QUE TRAIGA LOS TRAZABLES 2
  SQLX$ = SQLX$ + " LEFT JOIN BOLRECEP B ON O.ORIGEN = B.NUMEBORE" 'BOLRECEP P QUE DEVULEVA EL NRO DE RECEPCION Y EL PROV
  SQLX$ = SQLX$ + " where not exists (select 1 from TANUMSE where O.IDSUBOR  = TANUMSE.IDSUBOR)" 'QUE NO EXISTAN EN EL TANUMSE (N.SERIE NO CREADO)
  SQLX$ = SQLX$ + " AND origen <> 'PN-0' "
  SQLX$ = SQLX$ + " AND M.TRAZABLE = 2"
  SQLX$ = SQLX$ + " ORDER BY IDSUBOR DESC"
  Set RS_FUN = READSET(SQLX$)
  Set RS_OF_PENDIENTES_NSERIE_CON_BR = RS_FUN
  Set RS_FUN = Nothing
End Function
Public Function RS_NSERIE_CON_BR() As ADODB.Recordset 'REALIZADA PARA MEDITEA
  Dim RS_FUN As ADODB.Recordset
  Set RS_FUN = New ADODB.Recordset
  SQLX$ = "Select DISTINCT T.NUMEROSERIE,T.COD_INTE,T.COD_EXTE,T.DESCRIP,T.NUMEBORE,B.RASO_PROV FROM TANUMSE  T  WITH(NOLOCK)"
  SQLX$ = SQLX$ + " LEFT JOIN BOLRECEP B ON  B.NUMEBORE = T.NUMEBORE" 'BOLRECEP P QUE DEVULEVA EL NRO DE RECEPCION Y EL PROV
  SQLX$ = SQLX$ + " ORDER BY T.COD_EXTE, T.NUMEROSERIE ASC "
  Set RS_FUN = READSET(SQLX$)
  Set RS_NSERIE_CON_BR = RS_FUN
  Set RS_FUN = Nothing
End Function
Public Function RS_FACT_PROFORMA_PEN(NC, Campos$, Optional MODO%) As ADODB.Recordset 'REALIZADA PARA MEDITEA
  'MODO = 0 SOLO PENDIENTES Y NO ANULADAS
  'MODO = 1 INCLUYE ANULADAS
  'MODO = 2 MUESTRA TODAS 'NO AGREGA FILTRO
  'MODO = 3 SOLO FACTURADAS Y NO ANULADAS
  'MODO = 4 SOLO SOLO ANULADAS
  
  Dim RS_FUN As ADODB.Recordset
  Set RS_FUN = New ADODB.Recordset
  SQLX$ = "Select " & Campos$
  SQLX$ = SQLX$ + " FROM PROFORMTL WITH(NOLOCK) "
  SQLX$ = SQLX$ + " WHERE NUCLIEN > 0"
  If NC > 0 Then SQLX$ = SQLX$ + " AND NUCLIEN = " & NC 'FILTRO POR CLIENTE
  If MODO% = 0 Then SQLX$ = SQLX$ + " AND (STATUS = 0 OR STATUS IS NULL ) AND STATUS < 9 AND (FacturaReal = '' OR FacturaReal IS NULL)"
  If MODO% = 1 Then SQLX$ = SQLX$ + " AND (STATUS = 0 OR STATUS IS NULL )" 'INCLUYE ANULADAS
  If MODO% = 3 Then SQLX$ = SQLX$ + " AND STATUS >= 1 AND STATUS < 9" 'SOLO FACTURADAS Y NO ANULADAS
  If MODO% = 4 Then SQLX$ = SQLX$ + " AND STATUS = 9" 'SOLO SOLO ANULADAS
  SQLX$ = SQLX$ + " ORDER BY FeHoReal ASC"
  Set RS_FUN = READSET(SQLX$)
  Set RS_FACT_PROFORMA_PEN = RS_FUN
  Set RS_FUN = Nothing
  
End Function
Public Function RS_COMBINAEQUIACC(Campos$, CONDI$) As ADODB.Recordset
          Dim RS_FUN As ADODB.Recordset
          Set RS_FUN = New ADODB.Recordset
          SQLX$ = "SELECT " & Campos$ & " FROM COMBINA_IVA WITH(NOLOCK)  "
          If TRIM$(CONDI$) <> "" Then SQLX$ = SQLX$ + CONDI$
          Set RS_FUN = READSET(SQLX$)
          Set RS_COMBINAEQUIACC = RS_FUN
          Set RS_FUN = Nothing
End Function
Public Function RS_SERIESDISPONIBLES(CI1%)
         'DEVUELVE LOS N.SERIES DISPONIBLES DE DEPOSITO >= 0 SEGUN CODIGO PARAMETRO
         'Y QUE NO ESTE NI ETREGADO NI ASIGNADO A ALGUN PEDIDO
         
         SQLX$ = "SELECT T.NUMEROSERIE, M.CODIGO, M.DESCRIPCION,T.DEPONUMSE FROM TANUMSE T"
         SQLX$ = SQLX$ + " INNER JOIN MASTER M WITH(NOLOCK) ON (T.COD_INTE = M.CODINT) "
         SQLX$ = SQLX$ + " WHERE T.Cod_Inte = " & CI1% & ""
         SQLX$ = SQLX$ + " AND not exists (select 1 from PEDDETA where PEDDETA.NUMEROSERIE  = T.NUMEROSERIE)"
         SQLX$ = SQLX$ + " AND (T.NumRemito = '' OR T.NumRemito IS NULL)"
         SQLX$ = SQLX$ + " AND T.DepoNumse >= 0" ' PARA QUE TRAIGA SOLO LOS TERMINADOS.
         Set RS_FUN = READSET(SQLX$)
         Set RS_SERIESDISPONIBLES = RS_FUN
         Set RS_FUN = Nothing

'         Set NUSERTMP = READSET(SQLX$)
'
'         With NUSERTMP
'           Do While Not .EOF
'              NUSERR$ = SAFETEXT$(!NUMEROSERIE)
''              NusePed$ = VALOBADA("PEDDETA", "NUMEROSERIE", "NUMEROSERIE = '" & NUSERR$ & "'")
''              If NusePed$ = "" Then
'                 Call REPLA(TTXX$, NUSERR$, 1, 16)
'                 COD$ = SAFETEXT$(!Codigo): DESC$ = SAFETEXT$(!Descripcion)
'                 Call REPLA(TTXX$, COD$ + "  " + DESC$, 17, 47)
''                 jj& = jj& + 1
''                 Call GRAREG("!INUSEDIS", TTXX$, jj&)
'              End If
'           .MoveNext
'           Loop
'         End With
'         NUSERTMP.Close
'         Set NUSERTMP = Nothing

End Function

        
Public Function RS_DocumVersionados(NroOriginal&, REFERENCIA$) As ADODB.Recordset
          Dim RS_FUN As ADODB.Recordset
          Set RS_FUN = New ADODB.Recordset
          CONDI$ = " NroOrig = " & NroOriginal&
          CONDI$ = CONDI$ + " AND REFERENCIA = '" & REFERENCIA$ & "'"
          SQLX$ = "SELECT NroOrig,FECHEMIS,NumDocPdoc,NVersion FROM IMAG_VERSION WITH(NOLOCK)  "
          If TRIM$(CONDI$) <> "" Then SQLX$ = SQLX$ + "WHERE " & CONDI$
          Set RS_FUN = READSET(SQLX$)
          Set RS_DocumVersionados = RS_FUN
          Set RS_FUN = Nothing
End Function


Public Function RS_CARAS() As ADODB.Recordset
  Dim RS_FUN As ADODB.Recordset
  Set RS_FUN = New ADODB.Recordset
  SQLX$ = "SELECT CAMPO1, CAMPO2 FROM DATASQL WITH(NOLOCK) WHERE LISTA  = 'CARAS'"
  Set RS_FUN = READSET(SQLX$)
  Set RS_CARAS = RS_FUN
  Set RS_FUN = Nothing
End Function
Public Function RS_SINO() As ADODB.Recordset
  Dim RS_FUN As ADODB.Recordset
  Set RS_FUN = New ADODB.Recordset
  SQLX$ = "SELECT CAMPO1, CAMPO2 FROM DATASQL WITH(NOLOCK) WHERE LISTA  = 'TABLASI_NO'"
  Set RS_FUN = READSET(SQLX$)
  Set RS_SINO = RS_FUN
  Set RS_FUN = Nothing
End Function
Public Function RS_ESPEMATE() As ADODB.Recordset
  Dim RS_FUN As ADODB.Recordset
  Set RS_FUN = New ADODB.Recordset
  SQLX$ = "SELECT CAMPO1, CAMPO2 FROM DATASQL WITH(NOLOCK) WHERE LISTA  = 'ESPEMATE'"
  Set RS_FUN = READSET(SQLX$)
  Set RS_ESPEMATE = RS_FUN
  Set RS_FUN = Nothing
End Function

Public Function RS_ESPECOBRE() As ADODB.Recordset
  Dim RS_FUN As ADODB.Recordset
  Set RS_FUN = New ADODB.Recordset
  SQLX$ = "SELECT CAMPO1, CAMPO2 FROM DATASQL WITH(NOLOCK) WHERE LISTA  = 'ESPECOBRE'"
  Set RS_FUN = READSET(SQLX$)
  Set RS_ESPECOBRE = RS_FUN
  Set RS_FUN = Nothing
End Function
Public Function RS_ESPEPTH() As ADODB.Recordset
  Dim RS_FUN As ADODB.Recordset
  Set RS_FUN = New ADODB.Recordset
  SQLX$ = "SELECT CAMPO1, CAMPO2 FROM DATASQL WITH(NOLOCK) WHERE LISTA  = 'PTH'"
  Set RS_FUN = READSET(SQLX$)
  Set RS_ESPEPTH = RS_FUN
  Set RS_FUN = Nothing
End Function
Public Function RS_PISTAS() As ADODB.Recordset
  Dim RS_FUN As ADODB.Recordset
  Set RS_FUN = New ADODB.Recordset
  SQLX$ = "SELECT CAMPO1, CAMPO2 FROM DATASQL WITH(NOLOCK) WHERE LISTA  = 'PISTAS'"
  Set RS_FUN = READSET(SQLX$)
  Set RS_PISTAS = RS_FUN
  Set RS_FUN = Nothing
End Function
Public Function RS_PCB() As ADODB.Recordset
  Dim RS_FUN As ADODB.Recordset
  Set RS_FUN = New ADODB.Recordset
  SQLX$ = "SELECT CAMPO1, CAMPO2 FROM DATASQL WITH(NOLOCK) WHERE LISTA  = 'prograpcb'"
  Set RS_FUN = READSET(SQLX$)
  Set RS_PCB = RS_FUN
  Set RS_FUN = Nothing
End Function
Public Function RS_MASCANTSOLDANTE() As ADODB.Recordset
  Dim RS_FUN As ADODB.Recordset
  Set RS_FUN = New ADODB.Recordset
  SQLX$ = "SELECT CAMPO1, CAMPO2 FROM DATASQL WITH(NOLOCK) WHERE LISTA  = 'MASCANTSOLDANTE'"
  Set RS_FUN = READSET(SQLX$)
  Set RS_MASCANTSOLDANTE = RS_FUN
  Set RS_FUN = Nothing
End Function
Public Function RS_LEYECOMP() As ADODB.Recordset
  Dim RS_FUN As ADODB.Recordset
  Set RS_FUN = New ADODB.Recordset
  SQLX$ = "SELECT CAMPO1, CAMPO2 FROM DATASQL WITH(NOLOCK) WHERE LISTA  = 'LEYECOMP'"
  Set RS_FUN = READSET(SQLX$)
  Set RS_LEYECOMP = RS_FUN
  Set RS_FUN = Nothing
End Function

 Function RS_ListaCotización(Optional NROCOT&) As ADODB.Recordset
  Dim RS_FUN As ADODB.Recordset
  Set RS_FUN = New ADODB.Recordset
  Campos$ = "N°Cotiz./F9;Versión/F9;Fecha/D8;Cant./F8.1;Pre-Dm2/F12.3;Precio/F12.2;Gasto/F12.2;PTH/A8;Faz/A10"
  Campos$ = Campos$ + ";Terminaciones/A12;Sup/F9.3;Est./A4;Código/A20;Modelo/A20;Cta.Clie/F10;Cliente/A24;ID/F9"
   

  SQLX$ = "SELECT C.NRO_COTIZA,C.VRS_COTIZA,C.FECHEMISI,C.CANTIDAD,C.PRC_DM2,C.Preunit"
  SQLX$ = SQLX$ + " ,C.GASPREPA,C.TEXPTH,C.TEXCARAS,(C.LC + ' _ ' + C.LS)AS TERMINACIONES "
  SQLX$ = SQLX$ + ",C.SUPERFXLOTE,C.ESTADO,UPPER(C.Codigo),UPPER(C.MODELO),(UPPER(C.Codi_Cli) + ' ' + CONVERT(CHAR,C.NUME_CLI)) AS CLIENTE,UPPER(D.RASO_CLI),UPPER(C.COTIDAICHI_ID) FROM COTIDAICHI C WITH(NOLOCK) "
  SQLX$ = SQLX$ + " LEFT JOIN DEUDOR12 D ON D.NUME_CLI = C.NUME_CLI"
  If NROCOT& > 0 Then SQLX$ = SQLX$ + " WHERE NRO_COTIZA = " & NROCOT&
  SQLX$ = SQLX$ + " ORDER BY NRO_COTIZA DESC"
  Set RS_FUN = READSET(SQLX$)
  Set RS_ListaCotización = RS_FUN
  Set RS_FUN = Nothing


End Function
Public Function RS_DISPOMEDI_EYM() As ADODB.Recordset
  Dim RS_FUN As ADODB.Recordset
  Set RS_FUN = New ADODB.Recordset
  SQLX$ = "SELECT CAMPO1, CAMPO2 FROM DATASQL WITH(NOLOCK) WHERE LISTA  = 'DISPOMEDI' ORDER BY CAMPO2"
  Set RS_FUN = READSET(SQLX$)
  Set RS_DISPOMEDI_EYM = RS_FUN
  Set RS_FUN = Nothing
End Function

Public Function RS_GENERAL(TABLA$, Campos$, CONDI$) As ADODB.Recordset
          Dim RS_FUN As ADODB.Recordset
          Set RS_FUN = New ADODB.Recordset
          VariableWithnolock$ = " WITH(NOLOCK)"
          If InStr(UCase$(TABLA$), " JOIN ") > 0 Then VariableWithnolock$ = "" 'SI TIENE INNER JOIN SE DEBE AGREGAR EN ORIGEN
          SQLX$ = "SELECT " & Campos$ & " FROM  " & TABLA$ & VariableWithnolock$ 'EJEMPLO mnuCodiRang_Click EN FRM COTICONT DE COTIZACO.VB
          If TRIM$(CONDI$) <> "" Then SQLX$ = SQLX$ + " WHERE " & CONDI$
          Set RS_FUN = READSET(SQLX$)
          Set RS_GENERAL = RS_FUN
          Set RS_FUN = Nothing
End Function
Public Function RS_CONSULTA_COMPLETA(SQLX$) As ADODB.Recordset
          Dim RS_FUN As ADODB.Recordset
          Set RS_FUN = New ADODB.Recordset
          Set RS_FUN = READSET(SQLX$)
          Set RS_CONSULTA_COMPLETA = RS_FUN
          Set RS_FUN = Nothing
End Function

Public Function RS_Tabla_Provincias_Config() As ADODB.Recordset
  Dim RS_FUN As ADODB.Recordset
  Set RS_FUN = New ADODB.Recordset
  SQLX$ = "SELECT CODIGO31662, NOMBRE,ID,CODFLEX FROM  PROVINCIA WITH(NOLOCK) ORDER BY ID"
  Set RS_FUN = READSET(SQLX$)
  Set RS_Tabla_Provincias_Config = RS_FUN
  Set RS_FUN = Nothing
End Function
'*****************************************************'

Public Function RS_ListaGrupoSecuencias() As ADODB.Recordset
    
    Dim RS_FUN As ADODB.Recordset
    SQLX$ = "SELECT CODSECUEN,REFERENCIA,CODCLIENTE,CODEXTRELA,CIRELA  FROM SecuenMaster WITH(NOLOCK) ORDER BY CODSECUEN"
    Set RS_FUN = READSET(SQLX$)
    Set RS_ListaGrupoSecuencias = RS_FUN
    On Error Resume Next
    Set RS_FUN = Nothing
    On Error GoTo 0

End Function
Public Function RS_CashFlowEmitidos() As ADODB.Recordset
    
    Dim RS_FUN As ADODB.Recordset
    SQLX$ = "SELECT ENCAFLOW_ID,REFERENCIA1,FECHAYHO,FEDES,FEHAS,INTERVALO FROM ENCAFLOW WITH(NOLOCK) ORDER BY FECHAYHO desc"
    Set RS_FUN = READSET(SQLX$)
    Set RS_CashFlowEmitidos = RS_FUN
    On Error Resume Next
    Set RS_FUN = Nothing
    On Error GoTo 0

End Function
Public Function RS_ProformasXPedido(NPED&) As ADODB.Recordset
    Dim RS_FUN As ADODB.Recordset
    SQLX$ = "SELECT NumeComp,Mone_Emis,TotNeto, FacturaReal,Referencia  FROM PROFORMTL WITH(NOLOCK) WHERE NROPED =  " & NPED&
    SQLX$ = SQLX$ + " AND STATUS < 9 ORDER BY PROFORMTL_ID"
    Set RS_FUN = READSET(SQLX$)
    Set RS_ProformasXPedido = RS_FUN
    On Error Resume Next
    Set RS_FUN = Nothing
    On Error GoTo 0

End Function

Public Function RS_LISTASELECGENRAL(NOMBRELISTA$, Optional CONDI$) As ADODB.Recordset
  'USAR PARA LISTAR TABLAS DE DATASQL CON DOS CAMPOS
  CONDI$ = TRIM$(CONDI$)
  If TRIM$(NOMBRELISTA$) = "" Then Exit Function
  Dim RS_FUN As ADODB.Recordset
  Set RS_FUN = New ADODB.Recordset
  SQLX$ = "SELECT CAMPO1, CAMPO2 FROM DATASQL WITH(NOLOCK) WHERE LISTA  = '" & NOMBRELISTA$ & "'"
  If CONDI$ <> "" Then SQLX$ = SQLX$ + " AND " & CONDI$
  Set RS_FUN = READSET(SQLX$)
  Set RS_LISTASELECGENRAL = RS_FUN
  Set RS_FUN = Nothing
End Function

Public Function RS_CONFIGBIGBAG() As ADODB.Recordset

  Dim RS_FUN As ADODB.Recordset
  Set RS_FUN = New ADODB.Recordset
  SQLX$ = "SELECT CODBIGBAG,DESCRIPCIONBB,CUERPOLAT,LATERAL2 ,FONDOPANEL,TAPAPOLLERA,POLLERA, UPANEL,VALVCARGA ,VALVDESCAR, "
  SQLX$ = SQLX$ + " CINTIPOLLERA,CINTDESC,CINTDESCSEG, CINTCARGPOLL,  CINTCARGSEG, CINTLINER, MAMPARO,HILO1, "
  SQLX$ = SQLX$ + " ESLINGA, ESLPERIM, ETIQUETA, SOBRE,LINER,ENVOLTORIO,CONFIGBB_ID"
  SQLX$ = SQLX$ + " FROM CONFIGBB WITH(NOLOCK)  WHERE CONFIGBB_ID > '0' ORDER BY CODBIGBAG ASC"
  Set RS_FUN = READSET(SQLX$)
  Set RS_CONFIGBIGBAG = RS_FUN
  Set RS_FUN = Nothing


End Function
Public Function RS_CODRANGO() As ADODB.Recordset

  Dim RS_FUN As ADODB.Recordset
  Set RS_FUN = New ADODB.Recordset
  SQLX$ = "SELECT CODRANGO,Descripcion,CODMATPRIMA,DESCRIMATPRIMA,CIRANGO,CIMATPRIMA,CODIGORANGO_ID"
  SQLX$ = SQLX$ + " FROM CODIGORANGO WITH(NOLOCK)  WHERE CODIGORANGO_ID > '0' ORDER BY CODRANGO ASC"
  Set RS_FUN = READSET(SQLX$)
  Set RS_CODRANGO = RS_FUN
  Set RS_FUN = Nothing

End Function

Public Function RS_REMITOSOFICIALES(MODO%) As ADODB.Recordset  'MEDITEA
  '0 TODOS
  '1 SIN ANULAR
  '2 ANULADOS
  Dim RS_FUN As ADODB.Recordset
  Set RS_FUN = New ADODB.Recordset
  SQLX$ = "SELECT DISTINCT FECHMOV,REMIOFI, REFERENCIA"
  SQLX$ = SQLX$ + " FROM MOVIOFI WITH(NOLOCK)  "
  SQLX$ = SQLX$ + " WHERE MOVIOFI_ID > 0 "
  If MODO% = 1 Then SQLX$ = SQLX$ + " AND STATUS < 9 "
  If MODO% = 2 Then SQLX$ = SQLX$ + " AND STATUS = 9 "
  SQLX$ = SQLX$ + " ORDER BY FECHMOV, REMIOFI ASC"
  Set RS_FUN = READSET(SQLX$)
  Set RS_REMITOSOFICIALES = RS_FUN
  Set RS_FUN = Nothing

End Function

Public Function RS_CONFIGSCRAP() As ADODB.Recordset  'CONTEPLAST

  Dim RS_FUN As ADODB.Recordset
  Set RS_FUN = New ADODB.Recordset
  SQLX$ = "SELECT CodSector,DescSector,CodCateg,DescCateg,CodArt,DescArt,DepIngre,CiArt,CONFSCRAP_ID"
  SQLX$ = SQLX$ + " FROM CONFSCRAP WITH(NOLOCK)  "
  SQLX$ = SQLX$ + " WHERE CONFSCRAP_ID > 0 "
  SQLX$ = SQLX$ + " ORDER BY CodSector, CodCateg ASC"
  Set RS_FUN = READSET(SQLX$)
  Set RS_CONFIGSCRAP = RS_FUN
  Set RS_FUN = Nothing

End Function

Public Function RS_InformesScrap(Optional Des%, Optional Has%) As ADODB.Recordset  'CONTEPLAST

  Dim RS_FUN As ADODB.Recordset
  Set RS_FUN = New ADODB.Recordset
  SQLX$ = "SELECT Fecha,CodSector,DescSector,CodCateg,DescCateg,CodArt,DescArt,DepIngre,CiArt,DETASCRAP_ID,KILOS"
  SQLX$ = SQLX$ + " FROM DETASCRAP WITH(NOLOCK)  "
  SQLX$ = SQLX$ + " WHERE DETASCRAP_ID > 0 "
  If Des% > 0 Then SQLX$ = SQLX$ + " AND  FECHA >=  '" & FETEXCOR1$(Des%) & "'"
  If Has% > 0 Then SQLX$ = SQLX$ + " AND  FECHA <=  '" & FETEXCOR1$(Has%) & "'"
  SQLX$ = SQLX$ + " ORDER BY CodSector, CodCateg ASC"
  Set RS_FUN = READSET(SQLX$)
  Set RS_InformesScrap = RS_FUN
  Set RS_FUN = Nothing

End Function
Public Function RS_TIPOTELAYREFILE(Optional CONDI$) As ADODB.Recordset
    
    Dim RS_FUN As ADODB.Recordset
    SQLX$ = "SELECT CAMPO1, CAMPO2,CAMPO3,CAMPO4,CAMPO5 FROM DATASQL WITH(NOLOCK)  "
    SQLX$ = SQLX$ + " WHERE LISTA  = 'TIPOTELACMB'"
    SQLX$ = SQLX$ + " ORDER BY CAMPO1"
    Set RS_FUN = READSET(SQLX$)
    Set RS_TIPOTELAYREFILE = RS_FUN
    On Error Resume Next
    Set RS_FUN = Nothing
    On Error GoTo 0

End Function
Public Function RS_OF_PendTaller() As ADODB.Recordset
    'LOS REGISTROS QUE TIENEN EL NRO DE MOVISTO_ID SON LOS QUE ESTAN PENDIENTES DE RECEPCION
    Dim RS_FUN As ADODB.Recordset
    SQLX$ = "SELECT FechMov, Nume_Prov,Raso_Cli, CANTINGRE,MOVTALLER_ID FROM MOVTALLER WITH(NOLOCK)"
    SQLX$ = SQLX$ + " LEFT join Proved12 ON NUME_PROV = Nume_Cli"
    SQLX$ = SQLX$ + " WHERE CODIMOVI = 'OF' "
    SQLX$ = SQLX$ + " AND (IDMOVISTO < 1 OR IDMOVISTO IS NULL) ORDER BY FECHMOV "
    
    Set RS_FUN = READSET(SQLX$)
    Set RS_OF_PendTaller = RS_FUN
    On Error Resume Next
    Set RS_FUN = Nothing
    On Error GoTo 0

End Function
Public Function RS_PedidoxCte(NC) As ADODB.Recordset
  Dim RS_FUN As ADODB.Recordset
  Set RS_FUN = New ADODB.Recordset

  SQLX$ = "SELECT NroPed, Referencia FROM PEDENCA WITH(NOLOCK) WHERE  Status >= 0 "
  If NC > 0 Then SQLX$ = SQLX$ + " AND NROCLIE = " & NC
  SQLX$ = SQLX$ + " ORDER BY NroPed DESC"
  Set RS_FUN = READSET(SQLX$)
  Set RS_PedidoxCte = RS_FUN
  Set RS_FUN = Nothing
End Function
Public Function RS_DetaPedidoAhp(NP&) As ADODB.Recordset
  Dim RS_FUN As ADODB.Recordset
  Set RS_FUN = New ADODB.Recordset
  SQLX$ = "SELECT PEDDETA.CodiExt, PEDDETA.Descrip, PEDDETA.CanPed, PEDDETA.CantEnt, PEDDETA.NuOrItem FROM PEDDETA WITH(NOLOCK)  "
  If NP& > 0 Then SQLX$ = SQLX$ + " WHERE PEDDETA.NROPED = " & NP&
  SQLX$ = SQLX$ + " ORDER BY PEDDETA.NuOrItem ASC"
  
  Set RS_FUN = READSET(SQLX$)
  Set RS_DetaPedidoAhp = RS_FUN
  Set RS_FUN = Nothing
End Function

Public Function RS_BOLRECE() As ADODB.Recordset
  Dim RS_FUN As ADODB.Recordset
  Set RS_FUN = New ADODB.Recordset
  SQLX$ = "SELECT NUMEBORE, FECHRECEP, NUMREM_PROV, RASO_PROV FROM BOLRECEP WITH(NOLOCK) GROUP BY NUMEBORE, FECHRECEP, NUMREM_PROV, RASO_PROV ORDER BY NumeBoRe DESC"
  Set RS_FUN = READSET(SQLX$)
  Set RS_BOLRECE = RS_FUN
  Set RS_FUN = Nothing
End Function
Public Function RS_FormulaYCorte(CI1%) As ADODB.Recordset
  Dim RS_FUN As ADODB.Recordset 'ESTO ES PARA PERGAMINO, SIEMPRE VA A MANTNER UNA SOLA ESTRUCTURA POR CODICONJ A DIFERENCIA DE CONTE QUE PUEDE HABER MAS DE UNA COTIZACION POR CODIGO
'  Set RS_FUN = New ADODB.Recordset
'    SQLX$ = "SELECT F.CODIELEM,F.CODXELEM,M.DESCRIPCION,NUORIT,Formulas_id,corte,MODUXBB FROM Formulas F WITH(NOLOCK) INNER JOIN MASTER M  ON M.CODINT = F.CODIELEM"
'    SQLX$ = SQLX$ + " WHERE CODICONJ = " & CI1%
'    SQLX$ = SQLX$ + " ORDER BY NuorIt ASC"
  
    SQLX$ = "SELECT C.CODINTELEM,C.CODIGO,M.DESCRIPCION,C.ORDEN,COTI1_ID,corte,MODUXBB,NROCOTIZACION,CodComponente,Componente"
    SQLX$ = SQLX$ + " FROM COTI1 C WITH(NOLOCK) INNER JOIN MASTER M  ON M.CODINT = C.CODINTELEM"
    SQLX$ = SQLX$ + " WHERE CODINTCONJ = " & CI1%
    SQLX$ = SQLX$ + " ORDER BY ORDEN ASC"

  Set RS_FUN = READSET(SQLX$)
  Set RS_FormulaYCorte = RS_FUN
  Set RS_FUN = Nothing
End Function
