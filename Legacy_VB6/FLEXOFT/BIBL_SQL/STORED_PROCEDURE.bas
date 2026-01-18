Attribute VB_Name = "STORED_PROCEDURE"

'
'************************************AYUDA************************************************
'Mensaje: Los tipos de datos ntext y varchar(max) son incompatibles con el operador add.
'en el sp: ANOTANUMSE_SP tuve que castear para poder concatenar, aparentemente protesta por qeu el tipo de dato de historial es ntext y el parametro
'es varchar(max) y no eran compatibles ejemplo:
'    SQLX$ = SQLX$ + "SET Historial = CAST(Historial AS VARCHAR(MAX)) + @Historial"
'si no ver el sp completo: ANOTANUMSE_SP
'===================================================================================================================000

Sub CREAR_STORED_PROCEDURE()
    Static YaCreoStoredProcedure%
    If YaCreoStoredProcedure% > 0 Then Exit Sub
    YaCreoStoredProcedure% = 1
'======================================================================================================
'CREA SP MUESTRA EL NUMERO MAXIMO DE RECIBO
 
    If EXISTE_SP%("PROXIMO_N_RECIBO") > 0 Then
      SQLX$ = "alter Procedure PROXIMO_N_RECIBO "
    Else
      SQLX$ = "CREATE Procedure PROXIMO_N_RECIBO "
      EJECUTA_DIRECTO% = 1
    End If
   
    CONDI$ = " TipoMovim = 'RCOB'"
    'CONDI$ = CONDI$ + " AND LEFT(CodiCom_Cor,2)= 'RB'"
    CONDI$ = CONDI$ + " AND OpciMovim= 'CCORR'"
    CONDI$ = CONDI$ + " AND FechEmisi >=  @FECHALU  "
    CONDI$ = CONDI$ + " AND NumeTalo =  @NumeTalo"
    CONDI$ = CONDI$ + " AND  Status < 9"
    '
    SQLX$ = SQLX$ + "("
    
    SQLX$ = SQLX$ + "@NumeTalo INTEGER, "
    SQLX$ = SQLX$ + "@FECHALU smalldatetime "
    
    SQLX$ = SQLX$ + ")"
    SQLX$ = SQLX$ + " AS"
     
    SQLX$ = SQLX$ + " SELECT MAX(NumeComp) as maximo FROM CajaBanc WITH(NOLOCK) "
    SQLX$ = SQLX$ + " WHERE " & CONDI$
    If EJECUTA_DIRECTO% > 0 Then
      FLEXCONN.Execute (SQLX$)
    Else
      If CAMBIOSP("PROXIMO_N_RECIBO", SQLX$) Then
        FLEXCONN.Execute (SQLX$)
      End If
    End If
'====================================================================================================================
'======================================================================================================
'CREA SP MUESTRA EL NUMERO MAXIMO DE REMITO
 
    If EXISTE_SP%("MAXIMO_N_REMITO") > 0 Then
      SQLX$ = "alter Procedure MAXIMO_N_REMITO "
    Else
      SQLX$ = "CREATE Procedure MAXIMO_N_REMITO "
      EJECUTA_DIRECTO% = 1
    End If
   
    CONDI$ = "  FechMov >= @FechMov "
    CONDI$ = CONDI$ + " AND CODIMOVI = @CODIMOVI "
    CONDI$ = CONDI$ + " AND NUMPVTA= @NUMPVTA   "
    CONDI$ = CONDI$ + " AND Codiempr= @Codiempr"

    '
    SQLX$ = SQLX$ + "("
    
    SQLX$ = SQLX$ + "@CODIMOVI nvarchar(2), "
    SQLX$ = SQLX$ + "@FechMov smalldatetime , "
    SQLX$ = SQLX$ + "@NUMPVTA  INT,"
    SQLX$ = SQLX$ + "@Codiempr SMALLINT"
    
    SQLX$ = SQLX$ + ")"
    SQLX$ = SQLX$ + " AS"
     
    SQLX$ = SQLX$ + " SELECT MAX(DOPRINUM) as maximo FROM MOVISTO WITH(NOLOCK) "
    SQLX$ = SQLX$ + " WHERE " & CONDI$
    If EJECUTA_DIRECTO% > 0 Then
      FLEXCONN.Execute (SQLX$)
    Else
      If CAMBIOSP("MAXIMO_N_REMITO", SQLX$) Then
        FLEXCONN.Execute (SQLX$)
      End If
    End If
    '======================================================================================================
    'PEDIDOS PENDIENTES - FACTURACION
    If EXISTE_SP%("PEDIDOS_PENDIENTES") > 0 Then
      SQLX$ = "alter Procedure PEDIDOS_PENDIENTES "
    Else
      SQLX$ = "CREATE Procedure PEDIDOS_PENDIENTES "
      EJECUTA_DIRECTO% = 1
        End If

    SQLX$ = SQLX$ + "  @NroClie INT, @CODIEMPR SMALLINT, @PEDIDO INT"
    SQLX$ = SQLX$ + " AS"
    SQLX$ = SQLX$ + " DECLARE @SQLX NVARCHAR(600)"
    
    SQLX$ = SQLX$ + " SET @SQLX = ' SELECT DISTINCT PEDENCA.NroPed, PEDENCA.NroClie, PEDENCA.FechEmis, PEDENCA.NroOcom FROM PEDENCA INNER JOIN  PEDDETA '"
    SQLX$ = SQLX$ + " SET @SQLX = @SQLX + '  ON (PEDENCA.NroPed = PEDDETA.NroPed)'"
    SQLX$ = SQLX$ + " SET @SQLX = @SQLX + '   WHERE PEDENCA.NroClie = '"
    SQLX$ = SQLX$ + " SET @SQLX = @SQLX  +  Cast(@NroClie As Char(8))"
    
    SQLX$ = SQLX$ + " SET @SQLX = @SQLX + '  AND PEDDETA.Status < 8  AND PEDDETA.CantEnt < (PEDDETA.CanPed - 0.005)  AND PEDENCA.Status < 9 '"
    SQLX$ = SQLX$ + " SET @SQLX = @SQLX + '  AND PEDENCA.CODIEMPR = ' +   Cast(@CODIEMPR As Char(6))"
    SQLX$ = SQLX$ + " SET @SQLX = @SQLX + '   AND  PEDDETA.CODIEMPR = ' + Cast(@CODIEMPR As Char(6))"
    SQLX$ = SQLX$ + "  IF @PEDIDO > 0 BEGIN"
    SQLX$ = SQLX$ + " SET @SQLX = @SQLX + ' AND PEDDETA.NroPed = ' + @PEDIDO"
    SQLX$ = SQLX$ + " End"
    SQLX$ = SQLX$ + " SET @SQLX = @SQLX + ' ORDER BY PEDENCA.NroPed ASC'"
    SQLX$ = SQLX$ + "EXEC(@SQLX)"
    
    If EJECUTA_DIRECTO% > 0 Then
      FLEXCONN.Execute (SQLX$)
    Else
      If CAMBIOSP("PEDIDOS_PENDIENTES", SQLX$) Then
        FLEXCONN.Execute (SQLX$)
      End If
    End If
'============================================================================================================
'DEVUELVE EL PRECIO DE VENTA EDITADO DE AHP
If EXISTE_SP%("PRECIOVENTAAHP") > 0 Then
  SQLX$ = "Alter "
Else
  SQLX$ = "Create "
  EJECUTA_DIRECTO% = 1
End If
    CONDI$ = "  Nume_Cli >= @Nume_Cli "
    CONDI$ = CONDI$ + " AND CodiEmpr = @CodiEmpr "
    CONDI$ = CONDI$ + " AND CODINT= @CODINT   "
    CONDI$ = CONDI$ + " AND Via_Entrega= @Via_Entrega"

    SQLX$ = SQLX$ + " PROCEDURE PRECIOVENTAAHP "
    SQLX$ = SQLX$ + "( "
    SQLX$ = SQLX$ + "@Nume_Cli int, "
    SQLX$ = SQLX$ + "@CodiEmpr smallint, "
    SQLX$ = SQLX$ + "@CODINT int, "
    SQLX$ = SQLX$ + "@Via_Entrega smallint "
    SQLX$ = SQLX$ + ") "
    SQLX$ = SQLX$ + "AS "
    SQLX$ = SQLX$ + " "
    SQLX$ = SQLX$ + "SELECT PRECIOVENTA "
    SQLX$ = SQLX$ + "FROM COSTOXPROVE "
    SQLX$ = SQLX$ + " WHERE " & CONDI$
    If EJECUTA_DIRECTO% > 0 Then
      FLEXCONN.Execute (SQLX$)
    Else
      If CAMBIOSP("PRECIOVENTAAHP", SQLX$) Then
       FLEXCONN.Execute (SQLX$)
      End If
    End If
'============================================================
'ALTA DE CLIENTE DEVUELVE DA EL ALTA Y DEVUELVE EL NUEVO NUMERO
'============================================================

If EXISTE_SP%("ALTADEUDOR12") > 0 Then
  SQLX$ = "Alter "
Else
  SQLX$ = "Create "
  EJECUTA_DIRECTO% = 1
End If
SQLX$ = SQLX$ + " PROCEDURE  "
SQLX$ = SQLX$ + " ALTADEUDOR12 ( "
SQLX$ = SQLX$ + "@Lentre0_Cli nvarchar  (96)=NULL, "
SQLX$ = SQLX$ + "@CodiEmpr smallint=NULL, "
SQLX$ = SQLX$ + "@Lentre1_Cli nvarchar  (96)=NULL, "
SQLX$ = SQLX$ + "@NTrans_Cli smallint=NULL, "
SQLX$ = SQLX$ + "@Lentre2_Cli nvarchar  (96)=NULL, "
SQLX$ = SQLX$ + "@Hora_Ent nvarchar  (96)=NULL, "
SQLX$ = SQLX$ + "@Zona_Cli nvarchar  (64)=NULL, "
SQLX$ = SQLX$ + "@Zona_Rep nvarchar  (16)=NULL, "
SQLX$ = SQLX$ + "@Gru_Div int =NULL, "
SQLX$ = SQLX$ + "@Nume_Cli int = NULL, "
SQLX$ = SQLX$ + "@Codi_Cli nvarchar  (24)=NULL, "
SQLX$ = SQLX$ + "@Raso_Cli nvarchar  (256)=NULL, "
SQLX$ = SQLX$ + "@Fant_Cli nvarchar  (64)=NULL, "
SQLX$ = SQLX$ + "@Stat_Cli smallint =NULL, "
SQLX$ = SQLX$ + "@Domi_Cli nvarchar  (128) =NULL, "
SQLX$ = SQLX$ + "@Cpos_Cli nvarchar  (32) =NULL, "
SQLX$ = SQLX$ + "@Loca_Cli nvarchar  (96) =NULL, "
SQLX$ = SQLX$ + "@Pais_Cli nvarchar  (96) =NULL, "
SQLX$ = SQLX$ + "@Prov_Cli nvarchar  (4) =NULL, "
SQLX$ = SQLX$ + "@Tele_Cli nvarchar  (64)  =NULL, "
SQLX$ = SQLX$ + "@Faxi_Cli nvarchar  (64) =NULL, "
SQLX$ = SQLX$ + "@Mail_Cli nvarchar  (128) =NULL, "
SQLX$ = SQLX$ + "@PagWeb nvarchar  (256) =NULL, "
SQLX$ = SQLX$ + "@Piva_Cli smallint =NULL, "
SQLX$ = SQLX$ + "@Cuit_Cli nvarchar  (48) =NULL, "
SQLX$ = SQLX$ + "@Nibr_Cli nvarchar  (48) =NULL, "
SQLX$ = SQLX$ + "@Prvd_Cli nvarchar  (48) =NULL, "
SQLX$ = SQLX$ + "@Ctac_Cli nvarchar  (96) =NULL, "
SQLX$ = SQLX$ + "@Hora_Cli nvarchar  (64) =NULL, "
SQLX$ = SQLX$ + "@Cpag_Cli smallint =NULL, "
SQLX$ = SQLX$ + "@Lpre_Cli smallint =NULL, "
SQLX$ = SQLX$ + "@Vend_Cli smallint =NULL, "
SQLX$ = SQLX$ + "@Tcpr_Cli smallint =NULL, "
SQLX$ = SQLX$ + "@CaIb_Cli nvarchar  (4) =NULL, "
SQLX$ = SQLX$ + "@Dsta_Cli nvarchar  (64) =NULL, "
SQLX$ = SQLX$ + "@Lcre_Cli float  (8) =NULL, "
SQLX$ = SQLX$ + "@Sald_Cli float  (8) =NULL, "
SQLX$ = SQLX$ + "@CuMa_Cli smallint =NULL, "
SQLX$ = SQLX$ + "@GrCo_Cli smallint =NULL, "
SQLX$ = SQLX$ + "@Nota_Cli ntext =NULL, "
SQLX$ = SQLX$ + "@Obser_Entre ntext =NULL, "
SQLX$ = SQLX$ + "@Recar_Embal float  (8) =NULL, "
SQLX$ = SQLX$ + "@Mar_Ya_Fac smallint =NULL, "
SQLX$ = SQLX$ + "@DePosIva nvarchar  (64) =NULL, "
SQLX$ = SQLX$ + "@LCred_CCte float  (8) =NULL, "
SQLX$ = SQLX$ + "@EnvMaiDeuda smallint  =NULL, "
SQLX$ = SQLX$ + "@Ctact_Cob nvarchar  (96) =NULL, "
SQLX$ = SQLX$ + "@Tele_Cob nvarchar  (96) =NULL, "
SQLX$ = SQLX$ + "@Mail_Cob nvarchar  (256) =NULL, "
SQLX$ = SQLX$ + "@DiasRec_Cob nvarchar  (128) =NULL, "
SQLX$ = SQLX$ + "@DiasPago_Cob nvarchar  (128) =NULL, "
SQLX$ = SQLX$ + "@Segui_Cob ntext =NULL, "
SQLX$ = SQLX$ + "@MailFactura_Cob nvarchar  (128) =NULL, "
SQLX$ = SQLX$ + "@CELU_CLI nvarchar  (256) =NULL, "
SQLX$ = SQLX$ + "@Lentre3_Cli nvarchar  (20) =NULL, "
SQLX$ = SQLX$ + "@JurisdiccioIIBB_CLI nvarchar  (4) =NULL, "
SQLX$ = SQLX$ + "@RegimenFacturaCredito smallint =NULL, "
SQLX$ = SQLX$ + "@Code         int =NULL output "
SQLX$ = SQLX$ + ") "
SQLX$ = SQLX$ + "AS "
SQLX$ = SQLX$ + " "
SQLX$ = SQLX$ + "set @Code = (SELECT MAX(NUME_CLI) FROM DEUDOR12) "
SQLX$ = SQLX$ + "SET @Code = 1 + @Code "
SQLX$ = SQLX$ + "INSERT INTO Deudor12 ( "
SQLX$ = SQLX$ + "CaIb_Cli, "
SQLX$ = SQLX$ + "CELU_CLI, "
SQLX$ = SQLX$ + "Codi_Cli, "
SQLX$ = SQLX$ + "CodiEmpr, "
SQLX$ = SQLX$ + "Cpag_Cli, "
SQLX$ = SQLX$ + "Cpos_Cli, "
SQLX$ = SQLX$ + "Ctac_Cli, "
SQLX$ = SQLX$ + "Ctact_Cob, "
SQLX$ = SQLX$ + "Cuit_Cli, "
SQLX$ = SQLX$ + "CuMa_Cli, "
SQLX$ = SQLX$ + "DePosIva, "
SQLX$ = SQLX$ + "DiasPago_Cob, "
SQLX$ = SQLX$ + "DiasRec_Cob, "
SQLX$ = SQLX$ + "Domi_Cli, "
SQLX$ = SQLX$ + "Dsta_Cli, "
SQLX$ = SQLX$ + "EnvMaiDeuda, "
SQLX$ = SQLX$ + "Fant_Cli, "
SQLX$ = SQLX$ + "Faxi_Cli, "
SQLX$ = SQLX$ + "GrCo_Cli, "
SQLX$ = SQLX$ + "Gru_Div, "
SQLX$ = SQLX$ + "Hora_Cli, "
SQLX$ = SQLX$ + "Hora_Ent, "
SQLX$ = SQLX$ + "JurisdiccioIIBB_CLI, "
SQLX$ = SQLX$ + "Lcre_Cli, "
SQLX$ = SQLX$ + "LCred_CCte, "
SQLX$ = SQLX$ + "Lentre0_Cli, "
SQLX$ = SQLX$ + "Lentre1_Cli, "
SQLX$ = SQLX$ + "Lentre2_Cli, "
SQLX$ = SQLX$ + "Lentre3_Cli, "
SQLX$ = SQLX$ + "Loca_Cli, "
SQLX$ = SQLX$ + "Lpre_Cli, "
SQLX$ = SQLX$ + "Mail_Cli, "
SQLX$ = SQLX$ + "Mail_Cob, "
SQLX$ = SQLX$ + "MailFactura_Cob, "
SQLX$ = SQLX$ + "Mar_Ya_Fac, "
SQLX$ = SQLX$ + "Nibr_Cli, "
SQLX$ = SQLX$ + "Nota_Cli, "
SQLX$ = SQLX$ + "NTrans_Cli, "
SQLX$ = SQLX$ + "Nume_Cli, "
SQLX$ = SQLX$ + "Obser_Entre, "
SQLX$ = SQLX$ + "PagWeb, "
SQLX$ = SQLX$ + "Pais_Cli, "
SQLX$ = SQLX$ + "Piva_Cli, "
SQLX$ = SQLX$ + "Prov_Cli, "
SQLX$ = SQLX$ + "Prvd_Cli, "
SQLX$ = SQLX$ + "Raso_Cli, "
SQLX$ = SQLX$ + "Recar_Embal, "
SQLX$ = SQLX$ + "RegimenFacturaCredito, "
SQLX$ = SQLX$ + "Sald_Cli, "
SQLX$ = SQLX$ + "Segui_Cob, "
SQLX$ = SQLX$ + "Stat_Cli, "
SQLX$ = SQLX$ + "Tcpr_Cli, "
SQLX$ = SQLX$ + "Tele_Cli, "
SQLX$ = SQLX$ + "Tele_Cob, "
SQLX$ = SQLX$ + "Vend_Cli, "
SQLX$ = SQLX$ + "Zona_Cli, "
SQLX$ = SQLX$ + "Zona_Rep "
SQLX$ = SQLX$ + ")VALUES( "
SQLX$ = SQLX$ + "@CaIb_Cli, "
SQLX$ = SQLX$ + "@CELU_CLI, "
SQLX$ = SQLX$ + "@Codi_Cli, "
SQLX$ = SQLX$ + "@CodiEmpr, "
SQLX$ = SQLX$ + "@Cpag_Cli, "
SQLX$ = SQLX$ + "@Cpos_Cli, "
SQLX$ = SQLX$ + "@Ctac_Cli, "
SQLX$ = SQLX$ + "@Ctact_Cob, "
SQLX$ = SQLX$ + "@Cuit_Cli, "
SQLX$ = SQLX$ + "@CuMa_Cli, "
SQLX$ = SQLX$ + "@DePosIva, "
SQLX$ = SQLX$ + "@DiasPago_Cob, "
SQLX$ = SQLX$ + "@DiasRec_Cob, "
SQLX$ = SQLX$ + "@Domi_Cli, "
SQLX$ = SQLX$ + "@Dsta_Cli, "
SQLX$ = SQLX$ + "@EnvMaiDeuda, "
SQLX$ = SQLX$ + "@Fant_Cli, "
SQLX$ = SQLX$ + "@Faxi_Cli, "
SQLX$ = SQLX$ + "@GrCo_Cli, "
SQLX$ = SQLX$ + "@Gru_Div, "
SQLX$ = SQLX$ + "@Hora_Cli, "
SQLX$ = SQLX$ + "@Hora_Ent, "
SQLX$ = SQLX$ + "@JurisdiccioIIBB_CLI, "
SQLX$ = SQLX$ + "@Lcre_Cli, "
SQLX$ = SQLX$ + "@LCred_CCte, "
SQLX$ = SQLX$ + "@Lentre0_Cli, "
SQLX$ = SQLX$ + "@Lentre1_Cli, "
SQLX$ = SQLX$ + "@Lentre2_Cli, "
SQLX$ = SQLX$ + "@Lentre3_Cli, "
SQLX$ = SQLX$ + "@Loca_Cli, "
SQLX$ = SQLX$ + "@Lpre_Cli, "
SQLX$ = SQLX$ + "@Mail_Cli, "
SQLX$ = SQLX$ + "@Mail_Cob, "
SQLX$ = SQLX$ + "@MailFactura_Cob, "
SQLX$ = SQLX$ + "@Mar_Ya_Fac, "
SQLX$ = SQLX$ + "@Nibr_Cli, "
SQLX$ = SQLX$ + "@Nota_Cli, "
SQLX$ = SQLX$ + "@NTrans_Cli, "
SQLX$ = SQLX$ + "@Code, "
SQLX$ = SQLX$ + "@Obser_Entre, "
SQLX$ = SQLX$ + "@PagWeb, "
SQLX$ = SQLX$ + "@Pais_Cli, "
SQLX$ = SQLX$ + "@Piva_Cli, "
SQLX$ = SQLX$ + "@Prov_Cli, "
SQLX$ = SQLX$ + "@Prvd_Cli, "
SQLX$ = SQLX$ + "@Raso_Cli, "
SQLX$ = SQLX$ + "@Recar_Embal, "
SQLX$ = SQLX$ + "@RegimenFacturaCredito, "
SQLX$ = SQLX$ + "@Sald_Cli, "
SQLX$ = SQLX$ + "@Segui_Cob, "
SQLX$ = SQLX$ + "@Stat_Cli, "
SQLX$ = SQLX$ + "@Tcpr_Cli, "
SQLX$ = SQLX$ + "@Tele_Cli, "
SQLX$ = SQLX$ + "@Tele_Cob, "
SQLX$ = SQLX$ + "@Vend_Cli, "
SQLX$ = SQLX$ + "@Zona_Cli, "
SQLX$ = SQLX$ + "@Zona_Rep) "
SQLX$ = SQLX$ + "SELECT @Code "
If EJECUTA_DIRECTO% > 0 Then
  
  FLEXCONN.Execute (SQLX$)
Else
  If CAMBIOSP("", SQLX$) Then
   FLEXCONN.Execute (SQLX$)
  End If
End If
'============================================================
'FIN  ALTA DE CLIENTE DEVUELVE DA EL ALTA Y DEVUELVE EL NUEVO NUMERO
'============================================================
'============================================================================================================
'TRIGGER PARA GUARDAR EL PROXIMO NUMERO DE REMITO EN OTRA TABLA (MAXREMITO)
'SOLO ACTUALIZA SI ES EL NUMERO QUE MACHEA EN LA TABLA ES >
'ESTO SE CONSIDERA PARA EL CASO QUE SE ELIMINE UN REMITO Y SE VUELVA A GRABAR CON EL MISMO NUMERO MODIFICANDO MANUALMENTE
'SI NO FUERA ASI PODRIA GUARDAR UN NUMERO MENOR Y MOSTRAR UN NUMERO MENOR AL MAXIMO AL SER CONSULTADO EL MAXREMITO
'SI NO LO ENCUENTRA (SEGUN P.VENTA, FECHALIMITE, CODIMOVI,CODIEMPR) REALIZA UN INSERT

''If EXISTE_SP%("DIS_MAXDOPRINUM_insertar") > 0 Then
''  SQLX$ = "Alter "
''Else
''  SQLX$ = "Create "
''  EJECUTA_DIRECTO% = 1
''End If
''SQLX$ = SQLX$ + " TRIGGER DIS_MAXDOPRINUM_insertar "
''SQLX$ = SQLX$ + "on MOVISTO "
''SQLX$ = SQLX$ + "for insert "
''SQLX$ = SQLX$ + "as "
''SQLX$ = SQLX$ + "BEGIN "
''SQLX$ = SQLX$ + "declare @FLIMITE smalldatetime "
''SQLX$ = SQLX$ + "DECLARE @Codiempr SMALLINT "
''SQLX$ = SQLX$ + "DECLARE @NUMPVTA  INT "
''SQLX$ = SQLX$ + "DECLARE @CODIMOVI nvarchar(2) "
''SQLX$ = SQLX$ + "DECLARE @DOPRINUM INT "
''SQLX$ = SQLX$ + "DECLARE @MAXDOPRINUM1 INT "
''SQLX$ = SQLX$ + "SET @NUMPVTA  =  (SELECT NUMPVTA FROM INSERTED ) "
''SQLX$ = SQLX$ + "SET @CODIMOVI  = (SELECT CODIMOVI FROM INSERTED ) "
''SQLX$ = SQLX$ + "SET @Codiempr  = (SELECT Codiempr FROM INSERTED ) "
''SQLX$ = SQLX$ + "SET @DOPRINUM  = (SELECT DOPRINUM FROM INSERTED ) "
''SQLX$ = SQLX$ + "SET @FLIMITE  = (SELECT FechMov FROM INSERTED ) "
''SQLX$ = SQLX$ + " "
''SQLX$ = SQLX$ + "SET @MAXDOPRINUM1 = (SELECT MAX(DOPRINUM) FROM MAXREMITO WHERE NUMPVTA = @NUMPVTA "
''SQLX$ = SQLX$ + "AND  CODIMOVI = @CODIMOVI AND Codiempr = @Codiempr AND FechMov = @FLIMITE) "
''SQLX$ = SQLX$ + " "
''SQLX$ = SQLX$ + "IF cast(@MAXDOPRINUM1 AS int) > 0  "
''SQLX$ = SQLX$ + "BEGIN "
''SQLX$ = SQLX$ + "IF cast(@MAXDOPRINUM1 AS int) < @DOPRINUM "
''SQLX$ = SQLX$ + "BEGIN "
''SQLX$ = SQLX$ + "UPDATE MAXREMITO "
''SQLX$ = SQLX$ + "SET DOPRINUM = @DOPRINUM "
''SQLX$ = SQLX$ + "WHERE NUMPVTA = @NUMPVTA "
''SQLX$ = SQLX$ + "AND  CODIMOVI = @CODIMOVI "
''SQLX$ = SQLX$ + "AND Codiempr = @Codiempr "
''SQLX$ = SQLX$ + "AND FechMov = @FLIMITE "
''SQLX$ = SQLX$ + "END "
''SQLX$ = SQLX$ + "END "
''SQLX$ = SQLX$ + "ELSE "
''SQLX$ = SQLX$ + "BEGIN IF cast(@DOPRINUM AS int) > 0 "
''SQLX$ = SQLX$ + "INSERT INTO MAXREMITO "
''SQLX$ = SQLX$ + "( NUMPVTA,CODIMOVI, CODIEMPR,  DOPRINUM,FechMov) "
''SQLX$ = SQLX$ + "VALUES ( @NUMPVTA, @CODIMOVI, @Codiempr , @DOPRINUM,@FLIMITE ) "
''SQLX$ = SQLX$ + "END "
''SQLX$ = SQLX$ + " "
''SQLX$ = SQLX$ + "END "
''If EJECUTA_DIRECTO% > 0 Then
''  FLEXCONN.Execute (SQLX$)
''Else
''  If CAMBIOSP("DIS_MAXDOPRINUM_insertar", SQLX$) Then
''   FLEXCONN.Execute (SQLX$)
''  End If
''End If
'============================================================================================================
'CREADO PARA AHP PARA ACTUALIZAR EL COSTO POR PROVEEDOR DE LA TABLA COSTOXPROVE
If EXISTE_SP%("ACTU_COSTOPROVE_2") > 0 Then
  SQLX$ = "Alter "
Else
  SQLX$ = "Create "
  EJECUTA_DIRECTO% = 1
End If
SQLX$ = SQLX$ + " PROCEDURE ACTU_COSTOPROVE_2 "
SQLX$ = SQLX$ + "( "
SQLX$ = SQLX$ + "@Nume_Cli int = 0, "
SQLX$ = SQLX$ + "@CodiEmpr smallint = 0, "
SQLX$ = SQLX$ + "@CODINT int = 0, "
SQLX$ = SQLX$ + "@Via_Entrega smallint = 0, "
SQLX$ = SQLX$ + "@PRECIOVENTA Decimal (16,4) = 0, "
SQLX$ = SQLX$ + "@LISTA smallint = 1 "
SQLX$ = SQLX$ + ") "
SQLX$ = SQLX$ + "AS "
SQLX$ = SQLX$ + "DECLARE @REGCOUNT INT "
SQLX$ = SQLX$ + "SET @REGCOUNT = (SELECT COUNT(*) FROM COSTOXPROVE "
SQLX$ = SQLX$ + "WHERE CodiEmpr =  @CodiEmpr "
SQLX$ = SQLX$ + "AND CODINT =  @CODINT "
SQLX$ = SQLX$ + "AND LISTA =  @LISTA "
SQLX$ = SQLX$ + "AND Nume_Cli =  @Nume_Cli "
SQLX$ = SQLX$ + "AND Via_Entrega =  @Via_Entrega) "
SQLX$ = SQLX$ + " "
SQLX$ = SQLX$ + "IF  @REGCOUNT > 0 "
SQLX$ = SQLX$ + "BEGIN "
SQLX$ = SQLX$ + "UPDATE COSTOXPROVE "
SQLX$ = SQLX$ + "SET PRECIOVENTA = @PRECIOVENTA "
SQLX$ = SQLX$ + " "
SQLX$ = SQLX$ + "WHERE CodiEmpr =  @CodiEmpr "
SQLX$ = SQLX$ + "AND CODINT =  @CODINT "
SQLX$ = SQLX$ + "AND LISTA =  @LISTA "
SQLX$ = SQLX$ + "AND Nume_Cli =  @Nume_Cli "
SQLX$ = SQLX$ + "AND Via_Entrega =  @Via_Entrega "
SQLX$ = SQLX$ + "END "
SQLX$ = SQLX$ + "ELSE "
SQLX$ = SQLX$ + "BEGIN "
SQLX$ = SQLX$ + "INSERT INTO COSTOXPROVE "
SQLX$ = SQLX$ + "(CodiEmpr,CODINT ,LISTA, Nume_Cli, Via_Entrega, PRECIOVENTA) "
SQLX$ = SQLX$ + "VALUES "
SQLX$ = SQLX$ + "(@CodiEmpr, @CODINT,@LISTA,@Nume_Cli,@Via_Entrega,@PRECIOVENTA) "
SQLX$ = SQLX$ + "END "
SQLX$ = SQLX$ + " "
If EJECUTA_DIRECTO% > 0 Then
  FLEXCONN.Execute (SQLX$)
Else
  If CAMBIOSP("ACTU_COSTOPROVE_2", SQLX$) Then
   FLEXCONN.Execute (SQLX$)
  End If
End If
'============================================================================================================

    'PEDIDOS PENDIENTES - FACTURACION
    If EXISTE_SP%("PEDIDOS_PENDIENTES_ENTREGA") > 0 Then
      SQLX$ = "alter Procedure PEDIDOS_PENDIENTES_ENTREGA "
    Else
      SQLX$ = "CREATE Procedure PEDIDOS_PENDIENTES_ENTREGA "
      EJECUTA_DIRECTO% = 1
        End If

    SQLX$ = SQLX$ + "  @NroClie INT, @CODIEMPR SMALLINT, @PEDIDO INT"
    SQLX$ = SQLX$ + " AS"
    SQLX$ = SQLX$ + " DECLARE @SQLX NVARCHAR(600)"
    
    SQLX$ = SQLX$ + " SET @SQLX = ' SELECT DISTINCT PEDENCA.NroPed, PEDENCA.NroClie, PEDENCA.FechEmis, PEDENCA.NroOcom,PEDENCA.Domient,PEDENCA.LOCAENT FROM PEDENCA INNER JOIN  PEDDETA '"
    SQLX$ = SQLX$ + " SET @SQLX = @SQLX + '  ON (PEDENCA.NroPed = PEDDETA.NroPed)'"
    SQLX$ = SQLX$ + " SET @SQLX = @SQLX + '   WHERE PEDENCA.NroClie = '"
    SQLX$ = SQLX$ + " SET @SQLX = @SQLX  +  Cast(@NroClie As Char(8))"
    
    SQLX$ = SQLX$ + " SET @SQLX = @SQLX + '  AND PEDDETA.Status < 8  AND PEDDETA.CantEnt < (PEDDETA.CanPed - 0.005)  AND PEDENCA.Status < 9 '"
    SQLX$ = SQLX$ + " SET @SQLX = @SQLX + '  AND PEDENCA.CODIEMPR = ' +   Cast(@CODIEMPR As Char(6))"
    SQLX$ = SQLX$ + " SET @SQLX = @SQLX + '   AND  PEDDETA.CODIEMPR = ' + Cast(@CODIEMPR As Char(6))"
    SQLX$ = SQLX$ + "  IF @PEDIDO > 0 BEGIN"
    SQLX$ = SQLX$ + " SET @SQLX = @SQLX + ' AND PEDDETA.NroPed = ' + @PEDIDO"
    SQLX$ = SQLX$ + " End"
    SQLX$ = SQLX$ + " SET @SQLX = @SQLX + ' ORDER BY PEDENCA.NroPed ASC'"
    SQLX$ = SQLX$ + "EXEC(@SQLX)"
    
    If EJECUTA_DIRECTO% > 0 Then
      FLEXCONN.Execute (SQLX$)
    Else
      If CAMBIOSP("PEDIDOS_PENDIENTES_ENTREGA", SQLX$) Then
        FLEXCONN.Execute (SQLX$)
      End If
    End If
'============================================================================================================

    'CUENTA CORRIENTE DE CLIENTE
    If EXISTE_SP%("CTACTE_CLIENTE") > 0 Then
      SQLX$ = "Alter "
    Else
      SQLX$ = "Create "
      EJECUTA_DIRECTO% = 1
    End If
    SQLX$ = SQLX$ + " PROCEDURE CTACTE_CLIENTE "
    SQLX$ = SQLX$ + "(@NuClien int, @CodiEmpr smallint"
    SQLX$ = SQLX$ + ") "
    SQLX$ = SQLX$ + "AS "
    SQLX$ = SQLX$ + "SELECT CODICOM_LAR, IDEBECOMP,IHABECOMP, Mone_Emis, FechEmisi, NumeComp, Tipo_Cam, ObserGen, ModoOpciMovim "
    SQLX$ = SQLX$ + " FROM CAJABANC "
    SQLX$ = SQLX$ + " WHERE NuClien =  @NuClien "
    SQLX$ = SQLX$ + " AND (TipoMovim =  'FVEN' OR TipoMovim =  'RCOB' OR TipoMovim =  'AJUD') "
    SQLX$ = SQLX$ + " And CodiEmpr = @CodiEmpr "
    SQLX$ = SQLX$ + " ORDER BY FECHEMISI,FEHOREAL ASC "
    If EJECUTA_DIRECTO% > 0 Then
      FLEXCONN.Execute (SQLX$)
    Else
      If CAMBIOSP("CTACTE_CLIENTE", SQLX$) Then
       FLEXCONN.Execute (SQLX$)
      End If
    End If
    'FIN CUENTA CORRIENTE DE CLIENTE
'============================================================================================================
'SALDO DE CLIENTE USADO EN CUENTA CORRIENTE
    If EXISTE_SP%("SALDOCLIENTE") > 0 Then
      SQLX$ = "Alter "
    Else
      SQLX$ = "Create "
      EJECUTA_DIRECTO% = 1
    End If
    SQLX$ = SQLX$ + " PROCEDURE SALDOCLIENTE "
    SQLX$ = SQLX$ + "( "
    SQLX$ = SQLX$ + "@MODOMONEI SMALLINT , @NuClien int, @CodiEmpr smallint "
    SQLX$ = SQLX$ + ") "
    SQLX$ = SQLX$ + "AS "
    SQLX$ = SQLX$ + "DECLARE @SQL NVARCHAR(512) "
    SQLX$ = SQLX$ + "SET @SQL = "
    SQLX$ = SQLX$ + "'SELECT CodiCom_Cor, "
    SQLX$ = SQLX$ + "CodiCom_Lar, "
    SQLX$ = SQLX$ + "SaldDebe, "
    SQLX$ = SQLX$ + "SaldAcre_Dola, "
    SQLX$ = SQLX$ + "SaldDebe_Dola, "
    SQLX$ = SQLX$ + "SaldAcre, "
    SQLX$ = SQLX$ + "FechEmisi, "
    SQLX$ = SQLX$ + "FeVencim, "
    SQLX$ = SQLX$ + "IbruComp "
    SQLX$ = SQLX$ + "FROM SADEUPEN "
    SQLX$ = SQLX$ + "WHERE NuClien = ' +   Cast(@NuClien As Char(8)) + "
    SQLX$ = SQLX$ + "'AND (ABS(SaldDebe)  >= 0.005 "
    SQLX$ = SQLX$ + "or ABS(SaldAcre)  >= 0.005' "
    SQLX$ = SQLX$ + "IF @MODOMONEI = 2  BEGIN SET @SQL=  @SQL + ' OR ABS(SaldAcre_Dola) >= 0.005 OR ABS(SaldDebe_Dola) >= 0.005 ' END "
    SQLX$ = SQLX$ + "SET @SQL=  @SQL + ') ORDER BY FeVencim ASC' "
    'SQLX$ = SQLX$ + "print @SQL "
    SQLX$ = SQLX$ + "EXEC(@SQL)"
    If EJECUTA_DIRECTO% > 0 Then
      FLEXCONN.Execute (SQLX$)
    Else
      If CAMBIOSP("SALDOCLIENTE", SQLX$) Then
       FLEXCONN.Execute (SQLX$)
      End If
    'FIN SALDO DE CLIENTE USADO EN CUENTA CORRIENTE
    End If
    '============================================================================================================
    'GRABA HISTORIAL EN TABLA TANUMSE PASANDOLE EL TEXTO A AGREGAR YA FORMATEADO
    SQLX$ = "SELECT @@VERSION AS 'SQL Server Version';"
    Set rs = FLEXCONN.Execute(SQLX$)
    VERSIOX$ = rs.Fields(0)
    If XVALO(Mid$(VERSIOX$, 22)) <> 2000 Then ' VALIDO QUE SEA > a la version 2000 el varchar(max) da error
    ' MICROSOFT SQL SERVER 2008
    'Microsoft SQL Server 2005 'Microsoft SQL Server  2000
    If EXISTE_SP%("ANOTANUMSE_SP") > 0 Then
          SQLX$ = "Alter "
        Else
          SQLX$ = "Create "
          EJECUTA_DIRECTO% = 1
        End If
        SQLX$ = SQLX$ + " PROCEDURE ANOTANUMSE_SP "
        SQLX$ = SQLX$ + "( "
        SQLX$ = SQLX$ + "@NumeroSerie nvarchar  (64), "
        SQLX$ = SQLX$ + "@Historial VARCHAR(MAX) "
        SQLX$ = SQLX$ + ") "
        SQLX$ = SQLX$ + "AS "
        SQLX$ = SQLX$ + "UPDATE TANUMSE "
        SQLX$ = SQLX$ + "SET Historial = ISNULL(CAST(Historial AS VARCHAR(MAX)),'') + @Historial"
        SQLX$ = SQLX$ + " WHERE NumeroSerie =  @NumeroSerie "
        If EJECUTA_DIRECTO% > 0 Then
          FLEXCONN.Execute (SQLX$)
        Else
          If CAMBIOSP("ANOTANUMSE_SP", SQLX$) Then
           FLEXCONN.Execute (SQLX$)
          End If
        End If
    End If
    '============================================================================================================
    If EXISTE_SP%("CARCONSIPEN_SP") > 0 Then
      SQLX$ = "Alter "
    Else
      SQLX$ = "Create "
      EJECUTA_DIRECTO% = 1
    End If
    SQLX$ = SQLX$ + " PROCEDURE CARCONSIPEN_SP "
    SQLX$ = SQLX$ + "( "
    SQLX$ = SQLX$ + "@Nume_Clie INT, "
    SQLX$ = SQLX$ + "@FechMov1 smalldatetime, "
    SQLX$ = SQLX$ + "@FechMov2 smalldatetime "
    SQLX$ = SQLX$ + ") "
    SQLX$ = SQLX$ + "AS "
    SQLX$ = SQLX$ + "SELECT  COD_INTE, CANTSALID, CANTINGRE, NuPedCli, FECHMOV, DoPriCor, MoneVal, VALOUNIT, NUORITEM  , OCOMPRA_RTOABIERTO "
    SQLX$ = SQLX$ + "FROM MOVISTO WITH(NOLOCK) "
    SQLX$ = SQLX$ + "WHERE Nume_Clie =  @Nume_Clie "
    SQLX$ = SQLX$ + "AND FechMov >=  @FechMov1 AND  FechMov <=  @FechMov2 "
    SQLX$ = SQLX$ + "AND (CodiMovi = 'RT'  OR CodiMovi =  'DV'  OR CodiMovi =  '*R') "
    SQLX$ = SQLX$ + "AND DoPriLar <> ''  AND ((DoSecLar is null) OR (DoSecLar = '') OR (DoSecLar LIKE 'P/F%')) "
    If EJECUTA_DIRECTO% > 0 Then
      FLEXCONN.Execute (SQLX$)
    Else
      If CAMBIOSP("CARCONSIPEN_SP", SQLX$) Then
       FLEXCONN.Execute (SQLX$)
      End If
    End If
    '============================================================================================================
    '============================================================================================================
    If EXISTE_SP%("CARCONSIPEN_REPA_SP") > 0 Then
      SQLX$ = "Alter "
    Else
      SQLX$ = "Create "
      EJECUTA_DIRECTO% = 1
    End If
    SQLX$ = SQLX$ + " PROCEDURE CARCONSIPEN_REPA_SP "
    SQLX$ = SQLX$ + "( "
    SQLX$ = SQLX$ + "@Nume_Clie INT, "
    SQLX$ = SQLX$ + "@FechMov1 smalldatetime, "
    SQLX$ = SQLX$ + "@FechMov2 smalldatetime "
    SQLX$ = SQLX$ + ") "
    SQLX$ = SQLX$ + "AS "
    SQLX$ = SQLX$ + "SELECT  COD_INTE, CANTSALID, CANTINGRE, NuPedCli, FECHMOV, DoPriCor, MoneVal, VALOUNIT, NUORITEM  , OCOMPRA_RTOABIERTO "
    SQLX$ = SQLX$ + "FROM MOVISTO WITH(NOLOCK) "
    SQLX$ = SQLX$ + "WHERE Nume_Clie =  @Nume_Clie "
    SQLX$ = SQLX$ + "AND FechMov >=  @FechMov1 AND  FechMov <=  @FechMov2 "
    SQLX$ = SQLX$ + "AND (CodiMovi = 'RT'  OR CodiMovi =  'DV'  OR CodiMovi =  '*R') "
    SQLX$ = SQLX$ + "AND DoPriLar <> ''  AND ((DoSecLar is null) OR (DoSecLar = '') OR (DoSecLar LIKE 'P/F%')) "
    If EJECUTA_DIRECTO% > 0 Then
      FLEXCONN.Execute (SQLX$)
    Else
      If CAMBIOSP("CARCONSIPEN_REPA_SP", SQLX$) Then
       FLEXCONN.Execute (SQLX$)
      End If
    End If
    '============================================================================================================
    'FUNCION INSERT DE TABLA UBICASQL (AHP)
    '============================================================================================================
    If EXISTE_SP%("SPX_ADDUBICACION") > 0 Then
      SQLX$ = "Alter "
    Else
      SQLX$ = "Create "
      EJECUTA_DIRECTO% = 1
    End If
    
    SQLX$ = SQLX$ + " PROCEDURE SPX_ADDUBICACION "
    SQLX$ = SQLX$ + "( "
    SQLX$ = SQLX$ + "@Cod_Inte int, "
    SQLX$ = SQLX$ + "@CodiEmpr smallint, "
    SQLX$ = SQLX$ + "@CODIUBICA nvarchar  (64), "
    SQLX$ = SQLX$ + "@CANTIDAD float  (8), "
    SQLX$ = SQLX$ + "@OBSERVACION nvarchar  (128), "
    SQLX$ = SQLX$ + "@Nume_Pro int, "
    SQLX$ = SQLX$ + "@MARBORRA smallint "
    SQLX$ = SQLX$ + ") "
    SQLX$ = SQLX$ + "AS "
    SQLX$ = SQLX$ + " "
    SQLX$ = SQLX$ + "INSERT INTO UBICASQL ( "
    SQLX$ = SQLX$ + "CANTIDAD, "
    SQLX$ = SQLX$ + "Cod_Inte, "
    SQLX$ = SQLX$ + "CodiEmpr, "
    SQLX$ = SQLX$ + "CODIUBICA, "
    SQLX$ = SQLX$ + "MARBORRA, "
    SQLX$ = SQLX$ + "Nume_Pro, "
    SQLX$ = SQLX$ + "OBSERVACION "
    SQLX$ = SQLX$ + ")VALUES( "
    SQLX$ = SQLX$ + "@CANTIDAD, "
    SQLX$ = SQLX$ + "@Cod_Inte, "
    SQLX$ = SQLX$ + "@CodiEmpr, "
    SQLX$ = SQLX$ + "@CODIUBICA, "
    SQLX$ = SQLX$ + "@MARBORRA, "
    SQLX$ = SQLX$ + "@Nume_Pro, "
    SQLX$ = SQLX$ + "@OBSERVACION "
    SQLX$ = SQLX$ + ") "
    If EJECUTA_DIRECTO% > 0 Then
      FLEXCONN.Execute (SQLX$)
    Else
      If CAMBIOSP("SPX_ADDUBICACION", SQLX$) Then
       FLEXCONN.Execute (SQLX$)
      End If
    End If
    '============================================================================================================
    'FUNCION INSERT DE TABLA UBICASQL (AHP)
    '============================================================================================================
    '============================================================================================================
    'SP SELECT DE TABLA BOLRECEP DEVUELVE UN RECORDSET DE LOS PROVEEDORES SEGUN CI% (AHP) O SEA TODOS LOS QUE EN ALGUN MOMENTO RECEPCIONARON EL CODIGO
    '============================================================================================================

    If EXISTE_SP%("ProveedorSegunCodigo") > 0 Then
      SQLX$ = "Alter "
    Else
      SQLX$ = "Create "
      EJECUTA_DIRECTO% = 1
    End If
    SQLX$ = SQLX$ + " PROCEDURE ProveedorSegunCodigo "
    SQLX$ = SQLX$ + "( "
    SQLX$ = SQLX$ + "@Cod_Inte   int  "
    SQLX$ = SQLX$ + ") "
    SQLX$ = SQLX$ + "AS "
    SQLX$ = SQLX$ + " "
    SQLX$ = SQLX$ + "SELECT DISTINCT NumProv "
    SQLX$ = SQLX$ + "FROM BOLRECEP "
    SQLX$ = SQLX$ + "WHERE Cod_Inte =  @Cod_Inte "
    SQLX$ = SQLX$ + " and (CANTOAPRO > 0 "
    SQLX$ = SQLX$ + " OR CanToRech > 0 "
    SQLX$ = SQLX$ + " OR  CanToFalta > 0) "
    
    If EJECUTA_DIRECTO% > 0 Then
      FLEXCONN.Execute (SQLX$)
    Else
      If CAMBIOSP("ProveedorSegunCodigo", SQLX$) Then
       FLEXCONN.Execute (SQLX$)
      End If
    End If

    '============================================================================================================
    'FIN SP SELECT DE TABLA BOLRECEP DEVUELVE UN RECORDSET DE LOS PROVEEDORES SEGUN CI% (AHP)
    '============================================================================================================
    '============================================================================================================
    'DEVUELVE EL NUMERO DE PROVEEDOR PASANDOLE EL NUMERO DE RECEPCION.
    '============================================================================================================
    If EXISTE_SP%("ProveedorSegunlote") > 0 Then
      SQLX$ = "Alter "
    Else
      SQLX$ = "Create "
      EJECUTA_DIRECTO% = 1
    End If
    SQLX$ = SQLX$ + " PROCEDURE ProveedorSegunlote "
    SQLX$ = SQLX$ + "( "
    SQLX$ = SQLX$ + "@NumeBoRe float  (8) "
    SQLX$ = SQLX$ + ") "
    SQLX$ = SQLX$ + "AS "
    SQLX$ = SQLX$ + " "
    SQLX$ = SQLX$ + "SELECT NumProv "
    SQLX$ = SQLX$ + "FROM BOLRECEP "
    SQLX$ = SQLX$ + "WHERE NumeBoRe =  @NumeBoRe "
    SQLX$ = SQLX$ + " "
    If EJECUTA_DIRECTO% > 0 Then
      FLEXCONN.Execute (SQLX$)
    Else
      If CAMBIOSP("ProveedorSegunlote", SQLX$) Then
       FLEXCONN.Execute (SQLX$)
      End If
    End If
    '============================================================================================================
    'FIN DEVUELVE EL NUMERO DE PROVEEDOR PASANDOLE EL NUMERO DE RECEPCION.
    '============================================================================================================
    
    '============================================================================================================
    ' DEVUELVE SEGUN EL LOTE LA CANTIDAD PENDIENTE DE ENTREGA DEL PEDDETA OPCIONAL NRO DE PEDIDO
    ' SI SE LE PASA N.PEDIDO ES PARA QUE NO CONSIDERE EL MISMO.
    '============================================================================================================
    If EXISTE_SP%("SALDOLOTEPEDIDO2") > 0 Then
      SQLX$ = "Alter "
    Else
      SQLX$ = "Create "
      EJECUTA_DIRECTO% = 1
    End If
    SQLX$ = SQLX$ + " PROCEDURE SALDOLOTEPEDIDO2 "
    SQLX$ = SQLX$ + "( "
    SQLX$ = SQLX$ + "@CodiInt INTEGER, "
    SQLX$ = SQLX$ + "@IdenLote NVARCHAR(32)= NULL, " 'PARAMETRO OPCIONAL
    SQLX$ = SQLX$ + "@NROPED INTEGER = NULL, "        'PARAMETRO OPCIONAL
    SQLX$ = SQLX$ + "@CodiEmpr smallint "

    SQLX$ = SQLX$ + ") "
    SQLX$ = SQLX$ + "AS DECLARE @SQL NVARCHAR(512) "
    SQLX$ = SQLX$ + " "
    SQLX$ = SQLX$ + " SET @SQL = 'SELECT (SUM(CANPED)- SUM(CantEnt)) AS RESERVADO FROM PEDDETA' "
    SQLX$ = SQLX$ + " SET @SQL = @SQL + ' WHERE CANPED > CANTENT ' "
    SQLX$ = SQLX$ + " SET @SQL = @SQL + ' AND STATUS < 8 ' "
    SQLX$ = SQLX$ + " SET @SQL = @SQL + ' AND CODIINT  = ' +   Cast( @CodiInt As Char(8)) "
    SQLX$ = SQLX$ + " IF @IdenLote <> ''  BEGIN SET @SQL = @SQL +  ' AND IDENLOTE = ' + CHAR(39) +   @IdenLote  + CHAR(39) END"
    SQLX$ = SQLX$ + " IF @NROPED > 0  BEGIN SET @SQL = @SQL + ' AND NROPED <> ' +  Cast( @NROPED As Char(8)) END "
    SQLX$ = SQLX$ + " SET @SQL = @SQL + ' and  CodiEmpr = ' +   Cast( @CodiEmpr As Char(4)) "

    SQLX$ = SQLX$ + "EXEC(@SQL) "
    SQLX$ = SQLX$ + " "
    SQLX$ = SQLX$ + " "
    SQLX$ = SQLX$ + " "
    If EJECUTA_DIRECTO% > 0 Then
      FLEXCONN.Execute (SQLX$)
    Else
      If CAMBIOSP("SALDOLOTEPEDIDO2", SQLX$) Then
       FLEXCONN.Execute (SQLX$)
      End If
    End If
   
    '============================================================================================================
    'FIN DEVUELVE SEGUN EL LOTE LA CANTIDAD PENDIENTE DE ENTREGA DEL PEDDETA OPCIONAL NRO DE PEDIDO
    ' SI SE LE PASA N.PEDIDO ES PARA QUE NO CONSIDERE EL MISMO.
    '============================================================================================================
    
    '============================================================================================================
    ' DA DE ALTA UN NUEVA SOLICITUD DE COT
    '
    '============================================================================================================
   
    
    If EXISTE_SP%("NUEVA_SOLIC_COT") > 0 Then
      SQLX$ = "Alter "
    Else
      SQLX$ = "Create "
      EJECUTA_DIRECTO% = 1
    End If
    SQLX$ = SQLX$ + " PROCEDURE NUEVA_SOLIC_COT "
    SQLX$ = SQLX$ + "( "
    SQLX$ = SQLX$ + "@CodiEmpr smallint, "
    SQLX$ = SQLX$ + "@NumUsuar smallint, "
    SQLX$ = SQLX$ + "@FechMov smalldatetime, "
    SQLX$ = SQLX$ + "@FeFtoCot nvarchar  (24), "
    SQLX$ = SQLX$ + "@COT nvarchar  (256)=NULL, "
    SQLX$ = SQLX$ + "@TEXTCOT nvarchar  (24)=NULL, "
    SQLX$ = SQLX$ + "@FeSalida smalldatetime, "
    SQLX$ = SQLX$ + "@Code int= NULL output "
    SQLX$ = SQLX$ + ") "
    SQLX$ = SQLX$ + "AS "
    SQLX$ = SQLX$ + " "
    SQLX$ = SQLX$ + "INSERT INTO ENCACOT ( "
    SQLX$ = SQLX$ + "CodiEmpr, "
    SQLX$ = SQLX$ + "COT, "
    SQLX$ = SQLX$ + "FechMov, "
    SQLX$ = SQLX$ + "FeFtoCot, "
    SQLX$ = SQLX$ + "NumUsuar, "
    SQLX$ = SQLX$ + "TEXTCOT, "
    SQLX$ = SQLX$ + "FeSalida "
    SQLX$ = SQLX$ + ")VALUES( "
    SQLX$ = SQLX$ + "@CodiEmpr, "
    SQLX$ = SQLX$ + "@COT, "
    SQLX$ = SQLX$ + "@FechMov, "
    SQLX$ = SQLX$ + "@FeFtoCot, "
    SQLX$ = SQLX$ + "@NumUsuar, "
    SQLX$ = SQLX$ + "@TEXTCOT, "
    SQLX$ = SQLX$ + "@FeSalida "
    
    SQLX$ = SQLX$ + ") "
    SQLX$ = SQLX$ + " SET @Code=@@Identity "
    SQLX$ = SQLX$ + " RETURN @Code"
    If EJECUTA_DIRECTO% > 0 Then
      FLEXCONN.Execute (SQLX$)
    Else
      If CAMBIOSP("NUEVA_SOLIC_COT", SQLX$) Then
       FLEXCONN.Execute (SQLX$)
      End If
    End If

    '============================================================================================================
    ' FIN DA DE ALTA UN NUEVA SOLICITUD DE COT
    '
    '============================================================================================================
    '============================================================================================================
    ' DA DE ALTA UN NUEVA SOLICITUD DE COT
    '
    '============================================================================================================
   
    
    If EXISTE_SP%("NUEVA_SOLIC_COT_2") > 0 Then
      SQLX$ = "Alter "
    Else
      SQLX$ = "Create "
      EJECUTA_DIRECTO% = 1
    End If
    SQLX$ = SQLX$ + " PROCEDURE NUEVA_SOLIC_COT_2 "
    SQLX$ = SQLX$ + "( "
    SQLX$ = SQLX$ + "@CodiEmpr smallint, "
    SQLX$ = SQLX$ + "@NumUsuar smallint, "
    SQLX$ = SQLX$ + "@FechMov smalldatetime, "
    SQLX$ = SQLX$ + "@FeFtoCot nvarchar  (24), "
    SQLX$ = SQLX$ + "@COT nvarchar  (256)=NULL, "
    SQLX$ = SQLX$ + "@TEXTCOT nvarchar  (24)=NULL, "
    SQLX$ = SQLX$ + "@FeSalida smalldatetime, "
    SQLX$ = SQLX$ + "@Code int= NULL output "
    SQLX$ = SQLX$ + ") "
    SQLX$ = SQLX$ + "AS "
    SQLX$ = SQLX$ + " "
    SQLX$ = SQLX$ + "INSERT INTO ENCACOT ( "
    SQLX$ = SQLX$ + "CodiEmpr, "
    SQLX$ = SQLX$ + "COT, "
    SQLX$ = SQLX$ + "FechMov, "
    SQLX$ = SQLX$ + "FeFtoCot, "
    SQLX$ = SQLX$ + "NumUsuar, "
    SQLX$ = SQLX$ + "TEXTCOT, "
    SQLX$ = SQLX$ + "FeSalida "
    SQLX$ = SQLX$ + ")VALUES( "
    SQLX$ = SQLX$ + "@CodiEmpr, "
    SQLX$ = SQLX$ + "@COT, "
    SQLX$ = SQLX$ + "@FechMov, "
    SQLX$ = SQLX$ + "@FeFtoCot, "
    SQLX$ = SQLX$ + "@NumUsuar, "
    SQLX$ = SQLX$ + "@TEXTCOT, "
    SQLX$ = SQLX$ + "@FeSalida "
    
    SQLX$ = SQLX$ + ") "
    SQLX$ = SQLX$ + " SET @Code=@@Identity "
    SQLX$ = SQLX$ + " RETURN @Code"
    If EJECUTA_DIRECTO% > 0 Then
      FLEXCONN.Execute (SQLX$)
    Else
      If CAMBIOSP("NUEVA_SOLIC_COT_2", SQLX$) Then
       FLEXCONN.Execute (SQLX$)
      End If
    End If

    '============================================================================================================
    ' FIN DA DE ALTA UN NUEVA SOLICITUD DE COT
    '
    '============================================================================================================


    '============================================================================================================
    ' DA DE ALTA EN LA TABLA DE DETALLE DE  COT
    '
    '============================================================================================================


    
    
    If EXISTE_SP%("AGREGA_DETACOT") > 0 Then
      SQLX$ = "Alter "
    Else
      SQLX$ = "Create "
      EJECUTA_DIRECTO% = 1
    End If
    SQLX$ = SQLX$ + " PROCEDURE AGREGA_DETACOT "
    SQLX$ = SQLX$ + "( "
    SQLX$ = SQLX$ + "@CodiEmpr smallint, "
    SQLX$ = SQLX$ + "@FechMov smalldatetime, "
    SQLX$ = SQLX$ + "@FeFormCot nvarchar  (24), "
    SQLX$ = SQLX$ + "@Cod_Inte smallint, "
    SQLX$ = SQLX$ + "@Cod_Exte nvarchar  (48), "
    SQLX$ = SQLX$ + "@Descrip nvarchar  (128), "
    SQLX$ = SQLX$ + "@ValoUnit float  (8), "
    SQLX$ = SQLX$ + "@Nume_Clie int, "
    SQLX$ = SQLX$ + "@CantIngre float  (8), "
    SQLX$ = SQLX$ + "@CantSalid float  (8), "
    SQLX$ = SQLX$ + "@NUMPVTA smallint, "
    SQLX$ = SQLX$ + "@DOPRINUM int, "
    SQLX$ = SQLX$ + "@DoPriLar nvarchar  (24), "
    SQLX$ = SQLX$ + "@ID_ENCA int, "
    
    SQLX$ = SQLX$ + "@DOM_ENTRE nvarchar  (128), "
    SQLX$ = SQLX$ + "@LOCALIDAD_ENTR nvarchar  (64), "
    SQLX$ = SQLX$ + "@COPOSTAL nvarchar  (24), "
    SQLX$ = SQLX$ + "@CUITTRANSP nvarchar  (24), "
    SQLX$ = SQLX$ + "@PROVINCIA nvarchar  (24), "
    SQLX$ = SQLX$ + "@DOMITRANSP  nvarchar  (64), "
    SQLX$ = SQLX$ + "@LOCALITRANSP nvarchar  (64), "
    SQLX$ = SQLX$ + "@PATENTE nvarchar  (24), "
    SQLX$ = SQLX$ + "@DNITRANS nvarchar  (24), "
    SQLX$ = SQLX$ + "@FEREMITO smalldatetime, "
    SQLX$ = SQLX$ + "@CODUNICO nvarchar  (24), "
    SQLX$ = SQLX$ + "@Code   int = NULL output " ' ++++++++**********************
    SQLX$ = SQLX$ + ") "
    SQLX$ = SQLX$ + "AS "
    SQLX$ = SQLX$ + " "
    SQLX$ = SQLX$ + "INSERT INTO DETACOT ( "
    SQLX$ = SQLX$ + "CantIngre, "
    SQLX$ = SQLX$ + "CantSalid, "
    SQLX$ = SQLX$ + "Cod_Exte, "
    SQLX$ = SQLX$ + "Cod_Inte, "
    SQLX$ = SQLX$ + "CodiEmpr, "
    SQLX$ = SQLX$ + "Descrip, "
    SQLX$ = SQLX$ + "DoPriLar, "
    SQLX$ = SQLX$ + "DOPRINUM, "
    SQLX$ = SQLX$ + "FechMov, "
    SQLX$ = SQLX$ + "FeFormCot, "
    SQLX$ = SQLX$ + "ID_ENCA, "
    SQLX$ = SQLX$ + "Nume_Clie, "
    SQLX$ = SQLX$ + "NUMPVTA, "
    SQLX$ = SQLX$ + "ValoUnit, "
    
    SQLX$ = SQLX$ + "DOM_ENTRE, "
    SQLX$ = SQLX$ + "LOCALIDAD_ENTR, "
    SQLX$ = SQLX$ + "COPOSTAL, "
    SQLX$ = SQLX$ + "CUITTRANSP, "
    SQLX$ = SQLX$ + "PROVINCIA, "
    SQLX$ = SQLX$ + "DOMITRANSP, "
    SQLX$ = SQLX$ + "LOCALITRANSP, "
    SQLX$ = SQLX$ + "PATENTE, "
    SQLX$ = SQLX$ + "DNITRANS, "
    SQLX$ = SQLX$ + "FEREMITO, "
    SQLX$ = SQLX$ + "CODUNICO "
   
    SQLX$ = SQLX$ + ")VALUES( "
    SQLX$ = SQLX$ + "@CantIngre, "
    SQLX$ = SQLX$ + "@CantSalid, "
    SQLX$ = SQLX$ + "@Cod_Exte, "
    SQLX$ = SQLX$ + "@Cod_Inte, "
    SQLX$ = SQLX$ + "@CodiEmpr, "
    SQLX$ = SQLX$ + "@Descrip, "
    SQLX$ = SQLX$ + "@DoPriLar, "
    SQLX$ = SQLX$ + "@DOPRINUM, "
    SQLX$ = SQLX$ + "@FechMov, "
    SQLX$ = SQLX$ + "@FeFormCot, "
    SQLX$ = SQLX$ + "@ID_ENCA, "
    SQLX$ = SQLX$ + "@Nume_Clie, "
    SQLX$ = SQLX$ + "@NUMPVTA, "
    SQLX$ = SQLX$ + "@ValoUnit, "
    SQLX$ = SQLX$ + "@DOM_ENTRE, "
    SQLX$ = SQLX$ + "@LOCALIDAD_ENTR, "
    SQLX$ = SQLX$ + "@COPOSTAL, "
    SQLX$ = SQLX$ + "@CUITTRANSP, "
    SQLX$ = SQLX$ + "@PROVINCIA, "
    SQLX$ = SQLX$ + "@DOMITRANSP, "
    SQLX$ = SQLX$ + "@LOCALITRANSP, "
    SQLX$ = SQLX$ + "@PATENTE, "
    SQLX$ = SQLX$ + "@DNITRANS, "
    SQLX$ = SQLX$ + "@FEREMITO,"
    SQLX$ = SQLX$ + "@CODUNICO"
    
    SQLX$ = SQLX$ + ") "
    SQLX$ = SQLX$ + " SET @Code=@@Identity "
    SQLX$ = SQLX$ + " RETURN @Code"
    SQLX$ = SQLX$ + " "
    If EJECUTA_DIRECTO% > 0 Then
      FLEXCONN.Execute (SQLX$)
    Else
      If CAMBIOSP("AGREGA_DETACOT", SQLX$) Then
       FLEXCONN.Execute (SQLX$)
      End If
    End If

    '============================================================================================================
    'PROCEDIMIENTO QUE GUARDA LOS DATOS DE IMPORTACION DE LOS ARTICULOS IMPORTADOS EN RECEPCION (MEDITEA)
    'SI EXISTE EN LA TALBA LA BR ACTUALIZA, SI NO EXISTE REALIZA UN INSERT INTO
    '============================================================================================================
    If EXISTE_SP%("GRABADATOSINGDESPACHO_V1") > 0 Then
      SQLX$ = "Alter "
    Else
      SQLX$ = "Create "
      EJECUTA_DIRECTO% = 1
    End If
    SQLX$ = SQLX$ + " PROCEDURE  GRABADATOSINGDESPACHO_V1"
    SQLX$ = SQLX$ + " ( "
    SQLX$ = SQLX$ + "@CodiEmpr SMALLINT = 0,"
    SQLX$ = SQLX$ + "@FECHA  SMALLDATETIME, "
    SQLX$ = SQLX$ + "@DESPACHO NVARCHAR(128), "
    SQLX$ = SQLX$ + "@ADUANA NVARCHAR(128), "
    SQLX$ = SQLX$ + "@LOTE NVARCHAR(24), "
    SQLX$ = SQLX$ + "@NUMEBORE INTEGER) "
    SQLX$ = SQLX$ + "AS "
    SQLX$ = SQLX$ + "BEGIN "
    SQLX$ = SQLX$ + "DECLARE @CANTI INT "
    SQLX$ = SQLX$ + "SET @CANTI =  (SELECT COUNT(*) FROM DESPAIMPO WHERE IDENLOTE = @LOTE ) "
    SQLX$ = SQLX$ + "IF @CANTI > 0 "
    SQLX$ = SQLX$ + "BEGIN "
    SQLX$ = SQLX$ + "UPDATE DESPAIMPO SET "
    SQLX$ = SQLX$ + "CODIEMPR = @CodiEmpr, "
    SQLX$ = SQLX$ + "FECHA = @FECHA, "
    SQLX$ = SQLX$ + "DESPACHO = @DESPACHO, "
    SQLX$ = SQLX$ + "ADUANA= @ADUANA, "
    SQLX$ = SQLX$ + "NUMEBORE= @NUMEBORE "
    SQLX$ = SQLX$ + "WHERE IDENLOTE = @LOTE "
    SQLX$ = SQLX$ + "END "
    SQLX$ = SQLX$ + "ELSE "
    SQLX$ = SQLX$ + "INSERT INTO DESPAIMPO( "
    SQLX$ = SQLX$ + "CODIEMPR, "
    SQLX$ = SQLX$ + "FECHA , "
    SQLX$ = SQLX$ + "DESPACHO, "
    SQLX$ = SQLX$ + "ADUANA, "
    SQLX$ = SQLX$ + "IDENLOTE, "
    SQLX$ = SQLX$ + "NUMEBORE) "
    SQLX$ = SQLX$ + "VALUES( "
    SQLX$ = SQLX$ + "@CodiEmpr, "
    SQLX$ = SQLX$ + "@FECHA, "
    SQLX$ = SQLX$ + "@DESPACHO, "
    SQLX$ = SQLX$ + "@ADUANA, "
    SQLX$ = SQLX$ + "@LOTE, "
    SQLX$ = SQLX$ + "@NUMEBORE) "
    SQLX$ = SQLX$ + "END "
    If EJECUTA_DIRECTO% > 0 Then
      FLEXCONN.Execute (SQLX$)
    Else
      If CAMBIOSP("GRABADATOSINGDESPACHO_V1", SQLX$) Then
       FLEXCONN.Execute (SQLX$)
      End If
    End If
    '============================================================================================================
    'FIN
    'PROCEDIMIENTO QUE GUARDA LOS DATOS DE IMPORTACION DE LOS ARTICULOS IMPORTADOS EN RECEPCION (MEDITEA)
    'SI EXISTE EN LA TALBA LA BR ACTUALIZA, SI NO EXISTE REALIZA UN INSERT INTO
    '============================================================================================================

    '============================================================================================================
    'PROCEDIMIENTO UPDATE EN TABLA TANUMSE PARA REEMPLAZAR UN NUMERO DE SERIE POR OTRO DEVUELVE EL ID DONDE
    'SE REALIZO LA ACTUALIZACION
    '============================================================================================================


    If EXISTE_SP%("UPDATETANUMSE") > 0 Then
      SQLX$ = "Alter "
    Else
      SQLX$ = "Create "
      EJECUTA_DIRECTO% = 1
    End If
    SQLX$ = SQLX$ + " PROCEDURE  "
    SQLX$ = SQLX$ + " UPDATETANUMSE( "
    SQLX$ = SQLX$ + "@CodiEmpr smallint, "
    SQLX$ = SQLX$ + "@Numeroserie nvarchar  (24) , "
    SQLX$ = SQLX$ + "@NuevoNumeroserie nvarchar  (24), "
    SQLX$ = SQLX$ + "@Mesanio nvarchar(48), "
    SQLX$ = SQLX$ + "@NUMEORD smallint, "
    SQLX$ = SQLX$ + "@COD_INTE INTEGER ,"
    SQLX$ = SQLX$ + "@COD_EXTE nvarchar (48), "
    SQLX$ = SQLX$ + "@DESCRIP nvarchar (128),"
    SQLX$ = SQLX$ + "@IDSUBOR nvarchar(24), "
    SQLX$ = SQLX$ + "@FECHALTA smalldatetime, "
    SQLX$ = SQLX$ + "@HISTORIAL ntext , "
    
    SQLX$ = SQLX$ + "@Code int= NULL output "
    SQLX$ = SQLX$ + ") "
    SQLX$ = SQLX$ + " "
    SQLX$ = SQLX$ + "AS "
    SQLX$ = SQLX$ + "BEGIN "
    SQLX$ = SQLX$ + " "
    SQLX$ = SQLX$ + "SET @Code = (SELECT tanumse_id FROM TANUMSE WHERE Numeroserie = @Numeroserie )"
    
    SQLX$ = SQLX$ + "UPDATE TANUMSE SET CodiEmpr= @CodiEmpr, "
    SQLX$ = SQLX$ + "Numeroserie = @NuevoNumeroserie, "
    SQLX$ = SQLX$ + "Mesanio = @Mesanio, "
    SQLX$ = SQLX$ + "NUMEORD= @NUMEORD, "
    SQLX$ = SQLX$ + "COD_INTE=@COD_INTE, "
    SQLX$ = SQLX$ + "COD_EXTE=@COD_EXTE, "
    SQLX$ = SQLX$ + "DESCRIP= @DESCRIP, "
    SQLX$ = SQLX$ + "IDSUBOR= @IDSUBOR, "
    SQLX$ = SQLX$ + "FECHALTA= @FECHALTA, "
    SQLX$ = SQLX$ + "HISTORIAL= @HISTORIAL "
    
    SQLX$ = SQLX$ + "WHERE Numeroserie = @Numeroserie "
    SQLX$ = SQLX$ + "RETURN @Code "
    SQLX$ = SQLX$ + "END "
    On Error Resume Next
    If EJECUTA_DIRECTO% > 0 Then
      FLEXCONN.Execute (SQLX$)
    Else
      If CAMBIOSP("", SQLX$) Then
       FLEXCONN.Execute (SQLX$)
      End If
    End If
    On Error GoTo 0



    'SQLX$ = SQLX$ + "SET @SQL = @SQL + 'Nume_Cli = ' "
'SQLX$ = SQLX$ + "SET @SQL = @SQL  +  Cast( @Nume_Cli As Char(8)) "

'SQLX$ = SQLX$ + "BEGIN SET @SQL=  @SQL + ' CaIb_Cli' END "
'SQLX$ = SQLX$ + "ELSE "
'SQLX$ = SQLX$ + "SET @SQL = @SQL + CHAR(39) + @CaIb_Cli + CHAR(39) "

'EJEMPLOS
'    If EXISTE_SP%("ANUFACTU1") > 0 Then
'      SQLX$ = "alter Procedure ANUFACTU1 "
'   Else
'      SQLX$ = "CREATE Procedure ANUFACTU1 "
'   End If
'
'    CONDI$ = " (TipoMovim = 'FVEN' OR TipoMovim = 'AJUD') "
'    CONDI$ = CONDI$ + " AND CodiCom_Lar = @DOCULA"
'    CONDI$ = CONDI$ + " AND NuCLIEN = @NCLI"
'    CONDI$ = CONDI$ + " AND FechEmisi >=  @FECHALU  "
'
'    SQLX$ = SQLX$ + "("
'
'    SQLX$ = SQLX$ + "@DOCULA NVARCHAR(48),"
'    SQLX$ = SQLX$ + "@NCLI INTEGER,"
'    SQLX$ = SQLX$ + "@FECHALU smalldatetime"
'
'    SQLX$ = SQLX$ + ")"
'    SQLX$ = SQLX$ + " AS"
'
'    SQLX$ = SQLX$ + " SELECT Cae,CLI_REAL "
'    SQLX$ = SQLX$ + " From CajaBanc"
'    SQLX$ = SQLX$ + " WHERE " & CONDI$
'    FLEXCONN.Execute (SQLX$)
''--------------------------------------------------
''ELIMINA FACTURA
'    If EXISTE_SP%("ANUFACTURA2") > 0 Then
'      SQLX$ = "alter Procedure ANUFACTURA2 "
'    Else
'      SQLX$ = "CREATE Procedure ANUFACTURA2 "
'    End If
'
'    CONDI$ = " CodiCom_Lar= @DOCULA"
'    CONDI$ = CONDI$ & " And NuClien= @NCLI And FechEmisi >= @FECHALU"
'
'    SQLX$ = SQLX$ + "("
'
'    SQLX$ = SQLX$ + "@DOCULA NVARCHAR(48),"
'    SQLX$ = SQLX$ + "@NCLI INTEGER,"
'    SQLX$ = SQLX$ + "@FECHALU smalldatetime"
'
'    SQLX$ = SQLX$ + ")"
'    SQLX$ = SQLX$ + " AS"
'    SQLX$ = SQLX$ + " DELETE CAJABANC WHERE " & CONDI$
'    FLEXCONN.Execute (SQLX$)
'    '-------------------------------------------------------------
'   ' MODIFICA FACTURA
'    If EXISTE_SP%("ANUFACTURA3") > 0 Then
'      SQLX$ = "alter Procedure ANUFACTURA3 "
'    Else
'      SQLX$ = "CREATE Procedure ANUFACTURA3 "
'    End If
'
'    CONDI$ = " CodiCom_Lar= @DOCULA"
'    CONDI$ = CONDI$ & " And NuClien= @NCLI And FechEmisi >= @FECHALU"
'
'    SQLX$ = SQLX$ + "("
'
'    SQLX$ = SQLX$ + "@DOCULA NVARCHAR(48),"
'    SQLX$ = SQLX$ + "@NCLI INTEGER,"
'    SQLX$ = SQLX$ + "@FECHALU smalldatetime"
'
'    SQLX$ = SQLX$ + ")"
'    SQLX$ = SQLX$ + " AS"
'    SQLX$ = SQLX$ + " UPDATE CAJABANC set  ValAbsComp = 0,IDEBECOMP = 0,IHABECOMP = 0 "
'    SQLX$ = SQLX$ + " Where " & CONDI$
'    FLEXCONN.Execute (SQLX$)
''---------------------------------------------------
'   ' MODIFICA FACTURA
'    If EXISTE_SP%("BORRASADEUPEN") > 0 Then
'      SQLX$ = "alter Procedure BORRASADEUPEN "
'    Else
'      SQLX$ = "CREATE Procedure BORRASADEUPEN "
'    End If
'
'    CONDI$ = " CodiCom_Lar= @DOCULA"
'    CONDI$ = CONDI$ & " And NuClien= @NCLI "
'
'    SQLX$ = SQLX$ + "("
'
'    SQLX$ = SQLX$ + "@DOCULA NVARCHAR(48),"
'    SQLX$ = SQLX$ + "@NCLI INTEGER"
'
'    SQLX$ = SQLX$ + ")"
'    SQLX$ = SQLX$ + " AS"
'    SQLX$ = SQLX$ + " DELETE SADEUPEN "
'    SQLX$ = SQLX$ + " Where " & CONDI$
'    FLEXCONN.Execute (SQLX$)
''------------------------------------------------------------
''ACTUALIZA EL MOVISTO EN LA ANULACION DE LA FACTURA
'    If EXISTE_SP%("ACTUMOVISTO") > 0 Then
'      SQLX$ = "alter Procedure ACTUMOVISTO "
'    Else
'      SQLX$ = "CREATE Procedure ACTUMOVISTO "
'    End If
'
'    CONDI$ = " (DoSecLar= @DOCULA  OR DoPriLar =@DOCULA)"
'    CONDI$ = CONDI$ & " And Nume_Clie= @NCLI "
'
'    SQLX$ = SQLX$ + "("
'
'    SQLX$ = SQLX$ + "@DOCULA NVARCHAR(48),"
'    SQLX$ = SQLX$ + "@NCLI INTEGER"
'
'    SQLX$ = SQLX$ + ")"
'    SQLX$ = SQLX$ + " AS"
'    SQLX$ = SQLX$ + " UPDATE MOVISTO set  DOSECCOR = ' ',DoPriLar = ' ' ,DOSECLAR =' ', CANTINGRE= 0,CANTSALID = 0 "
'    SQLX$ = SQLX$ + " Where " & CONDI$
'    FLEXCONN.Execute (SQLX$)
'
'    '------------------------------------------------------------
''ACTUALIZA EL MOVISTO EN LA ANULACION DE LA FACTURA EN MOVIREPA
'    If EXISTE_SP%("ACTUMOVIREPA") > 0 Then
'      SQLX$ = "alter Procedure ACTUMOVIREPA "
'    Else
'      SQLX$ = "CREATE Procedure ACTUMOVIREPA "
'    End If
'
'    CONDI$ = " (DoSecLar= @DOCULA  OR DoPriLar =@DOCULA)"
'    CONDI$ = CONDI$ & " And Nume_Clie= @NCLI "
'
'    SQLX$ = SQLX$ + "("
'
'    SQLX$ = SQLX$ + "@DOCULA NVARCHAR(48),"
'    SQLX$ = SQLX$ + "@NCLI INTEGER"
'
'    SQLX$ = SQLX$ + ")"
'    SQLX$ = SQLX$ + " AS"
'    SQLX$ = SQLX$ + " UPDATE MOVIREPA set  DOSECCOR = ' ',DoPriLar = ' ' ,DOSECLAR =' ', CANTINGRE= 0, CANTSALID = 0 "
'    SQLX$ = SQLX$ + " Where " & CONDI$
'    FLEXCONN.Execute (SQLX$)
''---------------------------------------------------------------------------------------------------------------
''borra aplicobra
'
'    If EXISTE_SP%("BORRAAPLICOBRA") > 0 Then
'      SQLX$ = "alter Procedure BORRAAPLICOBRA "
'    Else
'      SQLX$ = "CREATE Procedure BORRAAPLICOBRA "
'    End If
'
'    CONDI$ = "(CoDocOr_Lar =@DOCULA OR CoDocAp_Lar = @DOCULA) AND NuCLIEN = @NCLI"
'
'    SQLX$ = SQLX$ + "("
'
'    SQLX$ = SQLX$ + "@DOCULA NVARCHAR(48),"
'    SQLX$ = SQLX$ + "@NCLI INTEGER"
'
'    SQLX$ = SQLX$ + ")"
'    SQLX$ = SQLX$ + " AS"
'    SQLX$ = SQLX$ + " DELETE APLICOBRA "
'    SQLX$ = SQLX$ + " Where " & CONDI$
'    FLEXCONN.Execute (SQLX$)
'    '
''--------------------------------------------------------------------------------------------------
'
''ACTUALIZA EL MOVISTO EN LA ANULACION DE LA FACTURA EN MOVIREPA
'    If EXISTE_SP%("LIMPIATANUMSE") > 0 Then
'      SQLX$ = "alter Procedure LIMPIATANUMSE "
'    Else
'      SQLX$ = "CREATE Procedure LIMPIATANUMSE "
'    End If
'
'    CONDI$ = "NUMFACTURA=@DOCULA"
'
'    SQLX$ = SQLX$ + "("
'
'    SQLX$ = SQLX$ + "@DOCULA NVARCHAR(48)"
'
'    SQLX$ = SQLX$ + ")"
'    SQLX$ = SQLX$ + " AS"
'    SQLX$ = SQLX$ + " UPDATE TANUMSE set  NUMFACTURA = ''"
'    SQLX$ = SQLX$ + " Where " & CONDI$
'    FLEXCONN.Execute (SQLX$)
''--------------------------------------------------------------------------------------------------
'
''ACTUALIZA LA TABLA MASTER TODOS LOS CAMPOS
''-----------------------------------------------------------------------------------------------
''ACTUALIZACION GENERAL DE TABLA DE CLIENTES DEUDOR12
' EJECUTA_DIRECTO% = 0
'
'If EXISTE_SP%("ACTUDEUDOR12") > 0 Then
'  SQLX$ = "Alter "
'Else
'  SQLX$ = "Create "
'  EJECUTA_DIRECTO% = 1
'End If
''
'SQLX$ = SQLX$ + " PROCEDURE ACTUDEUDOR12 "
'SQLX$ = SQLX$ + "(@Lentre0_Cli nvarchar  (96) = NULL "
'SQLX$ = SQLX$ + ",@Lentre1_Cli nvarchar  (96) = NULL "
'SQLX$ = SQLX$ + ",@NTrans_Cli smallint = NULL "
'SQLX$ = SQLX$ + ",@Lentre2_Cli nvarchar  (96) = NULL "
'SQLX$ = SQLX$ + ",@Hora_Ent nvarchar  (96) = NULL "
'SQLX$ = SQLX$ + ",@Zona_Cli nvarchar  (64) = NULL "
'SQLX$ = SQLX$ + ",@Zona_Rep nvarchar  (16) = NULL "
'SQLX$ = SQLX$ + ",@Gru_Div int = NULL "
'SQLX$ = SQLX$ + ",@Nume_Cli int = NULL "
'SQLX$ = SQLX$ + ",@Codi_Cli nvarchar  (24) = NULL "
'SQLX$ = SQLX$ + ",@Raso_Cli nvarchar  (128) = NULL "
'SQLX$ = SQLX$ + ",@Fant_Cli nvarchar  (64) = NULL "
'SQLX$ = SQLX$ + ",@Stat_Cli smallint = NULL "
'SQLX$ = SQLX$ + ",@Domi_Cli nvarchar  (128) = NULL "
'SQLX$ = SQLX$ + ",@Cpos_Cli nvarchar  (32) = NULL "
'SQLX$ = SQLX$ + ",@Loca_Cli nvarchar  (96) = NULL "
'SQLX$ = SQLX$ + ",@Pais_Cli nvarchar  (96) = NULL "
'SQLX$ = SQLX$ + ",@Prov_Cli nvarchar  (4) = NULL "
'SQLX$ = SQLX$ + ",@Tele_Cli nvarchar  (64) = NULL "
'SQLX$ = SQLX$ + ",@Faxi_Cli nvarchar  (64) = NULL "
'SQLX$ = SQLX$ + ",@Mail_Cli nvarchar  (128) = NULL "
'SQLX$ = SQLX$ + ",@PagWeb nvarchar  (256) = NULL "
'SQLX$ = SQLX$ + ",@Piva_Cli smallint = NULL "
'SQLX$ = SQLX$ + ",@Cuit_Cli nvarchar  (48) = NULL "
'SQLX$ = SQLX$ + ",@Nibr_Cli nvarchar  (48) = NULL "
'SQLX$ = SQLX$ + ",@Prvd_Cli nvarchar  (48) = NULL "
'SQLX$ = SQLX$ + ",@Ctac_Cli nvarchar  (96) = NULL "
'SQLX$ = SQLX$ + ",@Hora_Cli nvarchar  (64) = NULL "
'SQLX$ = SQLX$ + ",@Cpag_Cli smallint = NULL "
'SQLX$ = SQLX$ + ",@Lpre_Cli smallint = NULL "
'SQLX$ = SQLX$ + ",@Vend_Cli smallint = NULL "
'SQLX$ = SQLX$ + ",@Tcpr_Cli smallint = NULL "
'SQLX$ = SQLX$ + ",@CaIb_Cli nvarchar  (4) = NULL "
'SQLX$ = SQLX$ + ",@Dsta_Cli nvarchar  (64) = NULL "
'SQLX$ = SQLX$ + ",@Lcre_Cli float  (8) = NULL "
'SQLX$ = SQLX$ + ",@Sald_Cli float  (8) = NULL "
'SQLX$ = SQLX$ + ",@CuMa_Cli smallint = NULL "
'SQLX$ = SQLX$ + ",@GrCo_Cli smallint = NULL "
'SQLX$ = SQLX$ + ",@Nota_Cli NVARCHAR (MAX) = NULL "
'SQLX$ = SQLX$ + ",@Obser_Entre NVARCHAR (MAX) = NULL "
'SQLX$ = SQLX$ + ",@Recar_Embal float  (8) = NULL "
'SQLX$ = SQLX$ + ",@Mar_Ya_Fac smallint = NULL "
'SQLX$ = SQLX$ + ",@CodiEmpr smallint = NULL "
'SQLX$ = SQLX$ + ",@DePosIva nvarchar  (64) = NULL "
'SQLX$ = SQLX$ + ",@LCred_CCte float  (8) = NULL "
'SQLX$ = SQLX$ + ",@EnvMaiDeuda smallint = NULL "
'SQLX$ = SQLX$ + ",@Ctact_Cob nvarchar  (96) = NULL "
'SQLX$ = SQLX$ + ",@Tele_Cob nvarchar  (96) = NULL "
'SQLX$ = SQLX$ + ",@Mail_Cob nvarchar  (128) = NULL "
'SQLX$ = SQLX$ + ",@DiasRec_Cob nvarchar  (128) = NULL "
'SQLX$ = SQLX$ + ",@DiasPago_Cob nvarchar  (128) = NULL "
'SQLX$ = SQLX$ + ",@Segui_Cob NVARCHAR (MAX) = NULL "
'SQLX$ = SQLX$ + ",@MailFactura_Cob nvarchar  (128) = NULL "
'SQLX$ = SQLX$ + ",@CELU_ClI nvarchar  (64) = NULL "
'SQLX$ = SQLX$ + ") "
'SQLX$ = SQLX$ + "AS "
'SQLX$ = SQLX$ + "DECLARE @SQL NVARCHAR(MAX) "
'SQLX$ = SQLX$ + "SET @SQL = ' UPDATE Deudor12 SET ' "
'SQLX$ = SQLX$ + "SET @SQL = @SQL + 'CaIb_Cli= ' "
'SQLX$ = SQLX$ + "IF (@CaIb_Cli = '' OR @CaIb_Cli IS NULL) "
'SQLX$ = SQLX$ + "BEGIN SET @SQL=  @SQL + ' CaIb_Cli' END "
'SQLX$ = SQLX$ + "ELSE "
'SQLX$ = SQLX$ + "SET @SQL = @SQL + CHAR(39) + @CaIb_Cli + CHAR(39) "
'SQLX$ = SQLX$ + "SET @SQL = @SQL + ',CELU_ClI= ' "
'SQLX$ = SQLX$ + "IF (@CELU_ClI = '' OR @CELU_ClI IS NULL) "
'SQLX$ = SQLX$ + "BEGIN SET @SQL=  @SQL + ' CELU_ClI' END "
'SQLX$ = SQLX$ + "ELSE "
'SQLX$ = SQLX$ + "SET @SQL = @SQL + CHAR(39) + @CELU_ClI + CHAR(39) "
'SQLX$ = SQLX$ + "SET @SQL = @SQL + ',Codi_Cli= ' "
'SQLX$ = SQLX$ + "IF (@Codi_Cli = '' OR @Codi_Cli IS NULL) "
'SQLX$ = SQLX$ + "BEGIN SET @SQL=  @SQL + ' Codi_Cli' END "
'SQLX$ = SQLX$ + "ELSE "
'SQLX$ = SQLX$ + "SET @SQL = @SQL + CHAR(39) + @Codi_Cli + CHAR(39) "
'SQLX$ = SQLX$ + "SET @SQL = @SQL + ',CodiEmpr= ' "
'SQLX$ = SQLX$ + "IF (@CodiEmpr = '' OR @CodiEmpr IS NULL) "
'SQLX$ = SQLX$ + "BEGIN SET @SQL=  @SQL + ' CodiEmpr' END "
'SQLX$ = SQLX$ + "ELSE "
'SQLX$ = SQLX$ + "SET @SQL = @SQL + CHAR(39) +  CAST(@CodiEmpr As Char(16)) + CHAR(39) "
'SQLX$ = SQLX$ + "SET @SQL = @SQL + ',Cpag_Cli= ' "
'SQLX$ = SQLX$ + "IF (@Cpag_Cli = '' OR @Cpag_Cli IS NULL) "
'SQLX$ = SQLX$ + "BEGIN SET @SQL=  @SQL + ' Cpag_Cli' END "
'SQLX$ = SQLX$ + "ELSE "
'SQLX$ = SQLX$ + "SET @SQL = @SQL + CHAR(39) +  CAST(@Cpag_Cli As Char(16)) + CHAR(39) "
'SQLX$ = SQLX$ + "SET @SQL = @SQL + ',Cpos_Cli= ' "
'SQLX$ = SQLX$ + "IF (@Cpos_Cli = '' OR @Cpos_Cli IS NULL) "
'SQLX$ = SQLX$ + "BEGIN SET @SQL=  @SQL + ' Cpos_Cli' END "
'SQLX$ = SQLX$ + "ELSE "
'SQLX$ = SQLX$ + "SET @SQL = @SQL + CHAR(39) + @Cpos_Cli + CHAR(39) "
'SQLX$ = SQLX$ + "SET @SQL = @SQL + ',Ctac_Cli= ' "
'SQLX$ = SQLX$ + "IF (@Ctac_Cli = '' OR @Ctac_Cli IS NULL) "
'SQLX$ = SQLX$ + "BEGIN SET @SQL=  @SQL + ' Ctac_Cli' END "
'SQLX$ = SQLX$ + "ELSE "
'SQLX$ = SQLX$ + "SET @SQL = @SQL + CHAR(39) + @Ctac_Cli + CHAR(39) "
'SQLX$ = SQLX$ + "SET @SQL = @SQL + ',Ctact_Cob= ' "
'SQLX$ = SQLX$ + "IF (@Ctact_Cob = '' OR @Ctact_Cob IS NULL) "
'SQLX$ = SQLX$ + "BEGIN SET @SQL=  @SQL + ' Ctact_Cob' END "
'SQLX$ = SQLX$ + "ELSE "
'SQLX$ = SQLX$ + "SET @SQL = @SQL + CHAR(39) + @Ctact_Cob + CHAR(39) "
'SQLX$ = SQLX$ + "SET @SQL = @SQL + ',Cuit_Cli= ' "
'SQLX$ = SQLX$ + "IF (@Cuit_Cli = '' OR @Cuit_Cli IS NULL) "
'SQLX$ = SQLX$ + "BEGIN SET @SQL=  @SQL + ' Cuit_Cli' END "
'SQLX$ = SQLX$ + "ELSE "
'SQLX$ = SQLX$ + "SET @SQL = @SQL + CHAR(39) + @Cuit_Cli + CHAR(39) "
'SQLX$ = SQLX$ + "SET @SQL = @SQL + ',CuMa_Cli= ' "
'SQLX$ = SQLX$ + "IF (@CuMa_Cli = '' OR @CuMa_Cli IS NULL) "
'SQLX$ = SQLX$ + "BEGIN SET @SQL=  @SQL + ' CuMa_Cli' END "
'SQLX$ = SQLX$ + "ELSE "
'SQLX$ = SQLX$ + "SET @SQL = @SQL + CHAR(39) +  CAST(@CuMa_Cli As Char(16)) + CHAR(39) "
'SQLX$ = SQLX$ + "SET @SQL = @SQL + ',DePosIva= ' "
'SQLX$ = SQLX$ + "IF (@DePosIva = '' OR @DePosIva IS NULL) "
'SQLX$ = SQLX$ + "BEGIN SET @SQL=  @SQL + ' DePosIva' END "
'SQLX$ = SQLX$ + "ELSE "
'SQLX$ = SQLX$ + "SET @SQL = @SQL + CHAR(39) + @DePosIva + CHAR(39) "
'SQLX$ = SQLX$ + "SET @SQL = @SQL + ',DiasPago_Cob= ' "
'SQLX$ = SQLX$ + "IF (@DiasPago_Cob = '' OR @DiasPago_Cob IS NULL) "
'SQLX$ = SQLX$ + "BEGIN SET @SQL=  @SQL + ' DiasPago_Cob' END "
'SQLX$ = SQLX$ + "ELSE "
'SQLX$ = SQLX$ + "SET @SQL = @SQL + CHAR(39) + @DiasPago_Cob + CHAR(39) "
'SQLX$ = SQLX$ + "SET @SQL = @SQL + ',DiasRec_Cob= ' "
'SQLX$ = SQLX$ + "IF (@DiasRec_Cob = '' OR @DiasRec_Cob IS NULL) "
'SQLX$ = SQLX$ + "BEGIN SET @SQL=  @SQL + ' DiasRec_Cob' END "
'SQLX$ = SQLX$ + "ELSE "
'SQLX$ = SQLX$ + "SET @SQL = @SQL + CHAR(39) + @DiasRec_Cob + CHAR(39) "
'SQLX$ = SQLX$ + "SET @SQL = @SQL + ',Domi_Cli= ' "
'SQLX$ = SQLX$ + "IF (@Domi_Cli = '' OR @Domi_Cli IS NULL) "
'SQLX$ = SQLX$ + "BEGIN SET @SQL=  @SQL + ' Domi_Cli' END "
'SQLX$ = SQLX$ + "ELSE "
'SQLX$ = SQLX$ + "SET @SQL = @SQL + CHAR(39) + @Domi_Cli + CHAR(39) "
'SQLX$ = SQLX$ + "SET @SQL = @SQL + ',Dsta_Cli= ' "
'SQLX$ = SQLX$ + "IF (@Dsta_Cli = '' OR @Dsta_Cli IS NULL) "
'SQLX$ = SQLX$ + "BEGIN SET @SQL=  @SQL + ' Dsta_Cli' END "
'SQLX$ = SQLX$ + "ELSE "
'SQLX$ = SQLX$ + "SET @SQL = @SQL + CHAR(39) + @Dsta_Cli + CHAR(39) "
'SQLX$ = SQLX$ + "SET @SQL = @SQL + ',EnvMaiDeuda= ' "
'SQLX$ = SQLX$ + "IF (@EnvMaiDeuda = '' OR @EnvMaiDeuda IS NULL) "
'SQLX$ = SQLX$ + "BEGIN SET @SQL=  @SQL + ' EnvMaiDeuda' END "
'SQLX$ = SQLX$ + "ELSE "
'SQLX$ = SQLX$ + "SET @SQL = @SQL + CHAR(39) +  CAST(@EnvMaiDeuda As Char(16)) + CHAR(39) "
'SQLX$ = SQLX$ + "SET @SQL = @SQL + ',Fant_Cli= ' "
'SQLX$ = SQLX$ + "IF (@Fant_Cli = '' OR @Fant_Cli IS NULL) "
'SQLX$ = SQLX$ + "BEGIN SET @SQL=  @SQL + ' Fant_Cli' END "
'SQLX$ = SQLX$ + "ELSE "
'SQLX$ = SQLX$ + "SET @SQL = @SQL + CHAR(39) + @Fant_Cli + CHAR(39) "
'SQLX$ = SQLX$ + "SET @SQL = @SQL + ',Faxi_Cli= ' "
'SQLX$ = SQLX$ + "IF (@Faxi_Cli = '' OR @Faxi_Cli IS NULL) "
'SQLX$ = SQLX$ + "BEGIN SET @SQL=  @SQL + ' Faxi_Cli' END "
'SQLX$ = SQLX$ + "ELSE "
'SQLX$ = SQLX$ + "SET @SQL = @SQL + CHAR(39) + @Faxi_Cli + CHAR(39) "
'SQLX$ = SQLX$ + "SET @SQL = @SQL + ',GrCo_Cli= ' "
'SQLX$ = SQLX$ + "IF (@GrCo_Cli = '' OR @GrCo_Cli IS NULL) "
'SQLX$ = SQLX$ + "BEGIN SET @SQL=  @SQL + ' GrCo_Cli' END "
'SQLX$ = SQLX$ + "ELSE "
'SQLX$ = SQLX$ + "SET @SQL = @SQL + CHAR(39) +  CAST(@GrCo_Cli As Char(16)) + CHAR(39) "
'SQLX$ = SQLX$ + "SET @SQL = @SQL + ',Gru_Div= ' "
'SQLX$ = SQLX$ + "IF (@Gru_Div = '' OR @Gru_Div IS NULL) "
'SQLX$ = SQLX$ + "BEGIN SET @SQL=  @SQL + ' Gru_Div' END "
'SQLX$ = SQLX$ + "ELSE "
'SQLX$ = SQLX$ + "SET @SQL = @SQL + CHAR(39) +  CAST(@Gru_Div As Char(16)) + CHAR(39) "
'SQLX$ = SQLX$ + "SET @SQL = @SQL + ',Hora_Cli= ' "
'SQLX$ = SQLX$ + "IF (@Hora_Cli = '' OR @Hora_Cli IS NULL) "
'SQLX$ = SQLX$ + "BEGIN SET @SQL=  @SQL + ' Hora_Cli' END "
'SQLX$ = SQLX$ + "ELSE "
'SQLX$ = SQLX$ + "SET @SQL = @SQL + CHAR(39) + @Hora_Cli + CHAR(39) "
'SQLX$ = SQLX$ + "SET @SQL = @SQL + ',Hora_Ent= ' "
'SQLX$ = SQLX$ + "IF (@Hora_Ent = '' OR @Hora_Ent IS NULL) "
'SQLX$ = SQLX$ + "BEGIN SET @SQL=  @SQL + ' Hora_Ent' END "
'SQLX$ = SQLX$ + "ELSE "
'SQLX$ = SQLX$ + "SET @SQL = @SQL + CHAR(39) + @Hora_Ent + CHAR(39) "
'SQLX$ = SQLX$ + "SET @SQL = @SQL + ',Lcre_Cli= ' "
'SQLX$ = SQLX$ + "IF (@Lcre_Cli = '' OR @Lcre_Cli IS NULL) "
'SQLX$ = SQLX$ + "BEGIN SET @SQL=  @SQL + ' Lcre_Cli' END "
'SQLX$ = SQLX$ + "ELSE "
'SQLX$ = SQLX$ + "SET @SQL = @SQL + CHAR(39) +  CAST(@Lcre_Cli As Char(16)) + CHAR(39) "
'SQLX$ = SQLX$ + "SET @SQL = @SQL + ',LCred_CCte= ' "
'SQLX$ = SQLX$ + "IF (@LCred_CCte = '' OR @LCred_CCte IS NULL) "
'SQLX$ = SQLX$ + "BEGIN SET @SQL=  @SQL + ' LCred_CCte' END "
'SQLX$ = SQLX$ + "ELSE "
'SQLX$ = SQLX$ + "SET @SQL = @SQL + CHAR(39) +  CAST(@LCred_CCte As Char(16)) + CHAR(39) "
'SQLX$ = SQLX$ + "SET @SQL = @SQL + ',Lentre0_Cli= ' "
'SQLX$ = SQLX$ + "IF (@Lentre0_Cli = '' OR @Lentre0_Cli IS NULL) "
'SQLX$ = SQLX$ + "BEGIN SET @SQL=  @SQL + ' Lentre0_Cli' END "
'SQLX$ = SQLX$ + "ELSE "
'SQLX$ = SQLX$ + "SET @SQL = @SQL + CHAR(39) + @Lentre0_Cli + CHAR(39) "
'SQLX$ = SQLX$ + "SET @SQL = @SQL + ',Lentre1_Cli= ' "
'SQLX$ = SQLX$ + "IF (@Lentre1_Cli = '' OR @Lentre1_Cli IS NULL) "
'SQLX$ = SQLX$ + "BEGIN SET @SQL=  @SQL + ' Lentre1_Cli' END "
'SQLX$ = SQLX$ + "ELSE "
'SQLX$ = SQLX$ + "SET @SQL = @SQL + CHAR(39) + @Lentre1_Cli + CHAR(39) "
'SQLX$ = SQLX$ + "SET @SQL = @SQL + ',Lentre2_Cli= ' "
'SQLX$ = SQLX$ + "IF (@Lentre2_Cli = '' OR @Lentre2_Cli IS NULL) "
'SQLX$ = SQLX$ + "BEGIN SET @SQL=  @SQL + ' Lentre2_Cli' END "
'SQLX$ = SQLX$ + "ELSE "
'SQLX$ = SQLX$ + "SET @SQL = @SQL + CHAR(39) + @Lentre2_Cli + CHAR(39) "
'SQLX$ = SQLX$ + "SET @SQL = @SQL + ',Loca_Cli= ' "
'SQLX$ = SQLX$ + "IF (@Loca_Cli = '' OR @Loca_Cli IS NULL) "
'SQLX$ = SQLX$ + "BEGIN SET @SQL=  @SQL + ' Loca_Cli' END "
'SQLX$ = SQLX$ + "ELSE "
'SQLX$ = SQLX$ + "SET @SQL = @SQL + CHAR(39) + @Loca_Cli + CHAR(39) "
'SQLX$ = SQLX$ + "SET @SQL = @SQL + ',Lpre_Cli= ' "
'SQLX$ = SQLX$ + "IF (@Lpre_Cli = '' OR @Lpre_Cli IS NULL) "
'SQLX$ = SQLX$ + "BEGIN SET @SQL=  @SQL + ' Lpre_Cli' END "
'SQLX$ = SQLX$ + "ELSE "
'SQLX$ = SQLX$ + "SET @SQL = @SQL + CHAR(39) +  CAST(@Lpre_Cli As Char(16)) + CHAR(39) "
'SQLX$ = SQLX$ + "SET @SQL = @SQL + ',Mail_Cli= ' "
'SQLX$ = SQLX$ + "IF (@Mail_Cli = '' OR @Mail_Cli IS NULL) "
'SQLX$ = SQLX$ + "BEGIN SET @SQL=  @SQL + ' Mail_Cli' END "
'SQLX$ = SQLX$ + "ELSE "
'SQLX$ = SQLX$ + "SET @SQL = @SQL + CHAR(39) + @Mail_Cli + CHAR(39) "
'SQLX$ = SQLX$ + "SET @SQL = @SQL + ',Mail_Cob= ' "
'SQLX$ = SQLX$ + "IF (@Mail_Cob = '' OR @Mail_Cob IS NULL) "
'SQLX$ = SQLX$ + "BEGIN SET @SQL=  @SQL + ' Mail_Cob' END "
'SQLX$ = SQLX$ + "ELSE "
'SQLX$ = SQLX$ + "SET @SQL = @SQL + CHAR(39) + @Mail_Cob + CHAR(39) "
'SQLX$ = SQLX$ + "SET @SQL = @SQL + ',MailFactura_Cob= ' "
'SQLX$ = SQLX$ + "IF (@MailFactura_Cob = '' OR @MailFactura_Cob IS NULL) "
'SQLX$ = SQLX$ + "BEGIN SET @SQL=  @SQL + ' MailFactura_Cob' END "
'SQLX$ = SQLX$ + "ELSE "
'SQLX$ = SQLX$ + "SET @SQL = @SQL + CHAR(39) + @MailFactura_Cob + CHAR(39) "
'SQLX$ = SQLX$ + "SET @SQL = @SQL + ',Mar_Ya_Fac= ' "
'SQLX$ = SQLX$ + "IF (@Mar_Ya_Fac = '' OR @Mar_Ya_Fac IS NULL) "
'SQLX$ = SQLX$ + "BEGIN SET @SQL=  @SQL + ' Mar_Ya_Fac' END "
'SQLX$ = SQLX$ + "ELSE "
'SQLX$ = SQLX$ + "SET @SQL = @SQL + CHAR(39) +  CAST(@Mar_Ya_Fac As Char(16)) + CHAR(39) "
'SQLX$ = SQLX$ + "SET @SQL = @SQL + ',Nibr_Cli= ' "
'SQLX$ = SQLX$ + "IF (@Nibr_Cli = '' OR @Nibr_Cli IS NULL) "
'SQLX$ = SQLX$ + "BEGIN SET @SQL=  @SQL + ' Nibr_Cli' END "
'SQLX$ = SQLX$ + "ELSE "
'SQLX$ = SQLX$ + "SET @SQL = @SQL + CHAR(39) + @Nibr_Cli + CHAR(39) "
'SQLX$ = SQLX$ + "SET @SQL = @SQL + ',Nota_Cli= ' "
'SQLX$ = SQLX$ + "IF (@Nota_Cli LIKE '' OR @Nota_Cli IS NULL) "
'SQLX$ = SQLX$ + "BEGIN SET @SQL=  @SQL + ' Nota_Cli' END "
'SQLX$ = SQLX$ + "ELSE "
'SQLX$ = SQLX$ + "SET @SQL = @SQL + CHAR(39) + @Nota_Cli + CHAR(39) "
'SQLX$ = SQLX$ + "SET @SQL = @SQL + ',NTrans_Cli= ' "
'SQLX$ = SQLX$ + "IF (@NTrans_Cli = '' OR @NTrans_Cli IS NULL) "
'SQLX$ = SQLX$ + "BEGIN SET @SQL=  @SQL + ' NTrans_Cli' END "
'SQLX$ = SQLX$ + "ELSE "
'SQLX$ = SQLX$ + "SET @SQL = @SQL + CHAR(39) +  CAST(@NTrans_Cli As Char(16)) + CHAR(39) "
'SQLX$ = SQLX$ + "SET @SQL = @SQL + ',Nume_Cli= ' "
'SQLX$ = SQLX$ + "IF (@Nume_Cli = '' OR @Nume_Cli IS NULL) "
'SQLX$ = SQLX$ + "BEGIN SET @SQL=  @SQL + ' Nume_Cli' END "
'SQLX$ = SQLX$ + "ELSE "
'SQLX$ = SQLX$ + "SET @SQL = @SQL + CHAR(39) +  CAST(@Nume_Cli As Char(16)) + CHAR(39) "
'SQLX$ = SQLX$ + "SET @SQL = @SQL + ',Obser_Entre= ' "
'SQLX$ = SQLX$ + "IF (@Obser_Entre LIKE '' OR @Obser_Entre IS NULL) "
'SQLX$ = SQLX$ + "BEGIN SET @SQL=  @SQL + ' Obser_Entre' END "
'SQLX$ = SQLX$ + "ELSE "
'SQLX$ = SQLX$ + "SET @SQL = @SQL + CHAR(39) + @Obser_Entre + CHAR(39) "
'SQLX$ = SQLX$ + "SET @SQL = @SQL + ',PagWeb= ' "
'SQLX$ = SQLX$ + "IF (@PagWeb = '' OR @PagWeb IS NULL) "
'SQLX$ = SQLX$ + "BEGIN SET @SQL=  @SQL + ' PagWeb' END "
'SQLX$ = SQLX$ + "ELSE "
'SQLX$ = SQLX$ + "SET @SQL = @SQL + CHAR(39) + @PagWeb + CHAR(39) "
'SQLX$ = SQLX$ + "SET @SQL = @SQL + ',Pais_Cli= ' "
'SQLX$ = SQLX$ + "IF (@Pais_Cli = '' OR @Pais_Cli IS NULL) "
'SQLX$ = SQLX$ + "BEGIN SET @SQL=  @SQL + ' Pais_Cli' END "
'SQLX$ = SQLX$ + "ELSE "
'SQLX$ = SQLX$ + "SET @SQL = @SQL + CHAR(39) + @Pais_Cli + CHAR(39) "
'SQLX$ = SQLX$ + "SET @SQL = @SQL + ',Piva_Cli= ' "
'SQLX$ = SQLX$ + "IF (@Piva_Cli = '' OR @Piva_Cli IS NULL) "
'SQLX$ = SQLX$ + "BEGIN SET @SQL=  @SQL + ' Piva_Cli' END "
'SQLX$ = SQLX$ + "ELSE "
'SQLX$ = SQLX$ + "SET @SQL = @SQL + CHAR(39) +  CAST(@Piva_Cli As Char(16)) + CHAR(39) "
'SQLX$ = SQLX$ + "SET @SQL = @SQL + ',Prov_Cli= ' "
'SQLX$ = SQLX$ + "IF (@Prov_Cli = '' OR @Prov_Cli IS NULL) "
'SQLX$ = SQLX$ + "BEGIN SET @SQL=  @SQL + ' Prov_Cli' END "
'SQLX$ = SQLX$ + "ELSE "
'SQLX$ = SQLX$ + "SET @SQL = @SQL + CHAR(39) + @Prov_Cli + CHAR(39) "
'SQLX$ = SQLX$ + "SET @SQL = @SQL + ',Prvd_Cli= ' "
'SQLX$ = SQLX$ + "IF (@Prvd_Cli = '' OR @Prvd_Cli IS NULL) "
'SQLX$ = SQLX$ + "BEGIN SET @SQL=  @SQL + ' Prvd_Cli' END "
'SQLX$ = SQLX$ + "ELSE "
'SQLX$ = SQLX$ + "SET @SQL = @SQL + CHAR(39) + @Prvd_Cli + CHAR(39) "
'SQLX$ = SQLX$ + "SET @SQL = @SQL + ',Raso_Cli= ' "
'SQLX$ = SQLX$ + "IF (@Raso_Cli = '' OR @Raso_Cli IS NULL) "
'SQLX$ = SQLX$ + "BEGIN SET @SQL=  @SQL + ' Raso_Cli' END "
'SQLX$ = SQLX$ + "ELSE "
'SQLX$ = SQLX$ + "SET @SQL = @SQL + CHAR(39) + @Raso_Cli + CHAR(39) "
'SQLX$ = SQLX$ + "SET @SQL = @SQL + ',Recar_Embal= ' "
'SQLX$ = SQLX$ + "IF (@Recar_Embal = '' OR @Recar_Embal IS NULL) "
'SQLX$ = SQLX$ + "BEGIN SET @SQL=  @SQL + ' Recar_Embal' END "
'SQLX$ = SQLX$ + "ELSE "
'SQLX$ = SQLX$ + "SET @SQL = @SQL + CHAR(39) +  CAST(@Recar_Embal As Char(16)) + CHAR(39) "
'SQLX$ = SQLX$ + "SET @SQL = @SQL + ',Sald_Cli= ' "
'SQLX$ = SQLX$ + "IF (@Sald_Cli = '' OR @Sald_Cli IS NULL) "
'SQLX$ = SQLX$ + "BEGIN SET @SQL=  @SQL + ' Sald_Cli' END "
'SQLX$ = SQLX$ + "ELSE "
'SQLX$ = SQLX$ + "SET @SQL = @SQL + CHAR(39) +  CAST(@Sald_Cli As Char(16)) + CHAR(39) "
'SQLX$ = SQLX$ + "SET @SQL = @SQL + ',Segui_Cob= ' "
'SQLX$ = SQLX$ + "IF (@Segui_Cob LIKE '' OR @Segui_Cob IS NULL) "
'SQLX$ = SQLX$ + "BEGIN SET @SQL=  @SQL + ' Segui_Cob' END "
'SQLX$ = SQLX$ + "ELSE "
'SQLX$ = SQLX$ + "SET @SQL = @SQL + CHAR(39) + @Segui_Cob + CHAR(39) "
'SQLX$ = SQLX$ + "SET @SQL = @SQL + ',Stat_Cli= ' "
'SQLX$ = SQLX$ + "IF (@Stat_Cli = '' OR @Stat_Cli IS NULL) "
'SQLX$ = SQLX$ + "BEGIN SET @SQL=  @SQL + ' Stat_Cli' END "
'SQLX$ = SQLX$ + "ELSE "
'SQLX$ = SQLX$ + "SET @SQL = @SQL + CHAR(39) +  CAST(@Stat_Cli As Char(16)) + CHAR(39) "
'SQLX$ = SQLX$ + "SET @SQL = @SQL + ',Tcpr_Cli= ' "
'SQLX$ = SQLX$ + "IF (@Tcpr_Cli = '' OR @Tcpr_Cli IS NULL) "
'SQLX$ = SQLX$ + "BEGIN SET @SQL=  @SQL + ' Tcpr_Cli' END "
'SQLX$ = SQLX$ + "ELSE "
'SQLX$ = SQLX$ + "SET @SQL = @SQL + CHAR(39) +  CAST(@Tcpr_Cli As Char(16)) + CHAR(39) "
'SQLX$ = SQLX$ + "SET @SQL = @SQL + ',Tele_Cli= ' "
'SQLX$ = SQLX$ + "IF (@Tele_Cli = '' OR @Tele_Cli IS NULL) "
'SQLX$ = SQLX$ + "BEGIN SET @SQL=  @SQL + ' Tele_Cli' END "
'SQLX$ = SQLX$ + "ELSE "
'SQLX$ = SQLX$ + "SET @SQL = @SQL + CHAR(39) + @Tele_Cli + CHAR(39) "
'SQLX$ = SQLX$ + "SET @SQL = @SQL + ',Tele_Cob= ' "
'SQLX$ = SQLX$ + "IF (@Tele_Cob = '' OR @Tele_Cob IS NULL) "
'SQLX$ = SQLX$ + "BEGIN SET @SQL=  @SQL + ' Tele_Cob' END "
'SQLX$ = SQLX$ + "ELSE "
'SQLX$ = SQLX$ + "SET @SQL = @SQL + CHAR(39) + @Tele_Cob + CHAR(39) "
'SQLX$ = SQLX$ + "SET @SQL = @SQL + ',Vend_Cli= ' "
'SQLX$ = SQLX$ + "IF (@Vend_Cli = '' OR @Vend_Cli IS NULL) "
'SQLX$ = SQLX$ + "BEGIN SET @SQL=  @SQL + ' Vend_Cli' END "
'SQLX$ = SQLX$ + "ELSE "
'SQLX$ = SQLX$ + "SET @SQL = @SQL + CHAR(39) +  CAST(@Vend_Cli As Char(16)) + CHAR(39) "
'SQLX$ = SQLX$ + "SET @SQL = @SQL + ',Zona_Cli= ' "
'SQLX$ = SQLX$ + "IF (@Zona_Cli = '' OR @Zona_Cli IS NULL) "
'SQLX$ = SQLX$ + "BEGIN SET @SQL=  @SQL + ' Zona_Cli' END "
'SQLX$ = SQLX$ + "ELSE "
'SQLX$ = SQLX$ + "SET @SQL = @SQL + CHAR(39) + @Zona_Cli + CHAR(39) "
'SQLX$ = SQLX$ + "SET @SQL = @SQL + ',Zona_Rep= ' "
'SQLX$ = SQLX$ + "IF (@Zona_Rep = '' OR @Zona_Rep IS NULL) "
'SQLX$ = SQLX$ + "BEGIN SET @SQL=  @SQL + ' Zona_Rep' END "
'SQLX$ = SQLX$ + "ELSE "
'SQLX$ = SQLX$ + "SET @SQL = @SQL + CHAR(39) + @Zona_Rep + CHAR(39) "
'SQLX$ = SQLX$ + " "
'SQLX$ = SQLX$ + "SET @SQL = @SQL + ' WHERE ' "
'SQLX$ = SQLX$ + "SET @SQL = @SQL + 'Nume_Cli = ' "
'SQLX$ = SQLX$ + "SET @SQL = @SQL  +  Cast( @Nume_Cli As Char(8)) "
'SQLX$ = SQLX$ + "EXEC(@SQL); "
'If EJECUTA_DIRECTO% > 0 Then
'  FLEXCONN.Execute (SQLX$)
'Else
'  If CAMBIOSP("ACTUDEUDOR12", SQLX$) Then
'   FLEXCONN.Execute (SQLX$)
'  End If
'End If
''
'
''----------------------------------------------------------------------------------------------------------------
''REEMPLAZA EL CODIGO DE CLIENTE EN CAJABANC PRIMERO SE LE PASA CLIENTE PARA LA CONDICION LUEGO EL DE REEMPLAZO
'If EXISTE_SP%("UPDATENUCLIENBANC") > 0 Then
'  SQLX$ = "Alter "
'Else
'  SQLX$ = "Create "
'  EJECUTA_DIRECTO% = 1
'End If
'SQLX$ = SQLX$ + " PROCEDURE UPDATENUCLIENBANC "
'SQLX$ = SQLX$ + "(@NuClien int = NULL "
'SQLX$ = SQLX$ + " ,@NuClienRemp int = NULL "
'SQLX$ = SQLX$ + " "
'SQLX$ = SQLX$ + ") "
'SQLX$ = SQLX$ + "AS "
'SQLX$ = SQLX$ + "DECLARE @SQL NVARCHAR(MAX) "
'SQLX$ = SQLX$ + "SET @SQL = ' UPDATE CAJABANC SET ' "
'SQLX$ = SQLX$ + "SET @SQL = @SQL + 'NuClien= ' "
'SQLX$ = SQLX$ + "IF (@NuClien = '' OR @NuClien IS NULL) "
'SQLX$ = SQLX$ + "BEGIN SET @SQL=  @SQL + ' NuClien' END "
'SQLX$ = SQLX$ + "ELSE "
'SQLX$ = SQLX$ + "SET @SQL = @SQL + CHAR(39) +  CAST(@NuClienRemp As Char(16)) + CHAR(39) "
'SQLX$ = SQLX$ + " "
'SQLX$ = SQLX$ + "SET @SQL = @SQL + ' WHERE ' "
'SQLX$ = SQLX$ + "SET @SQL = @SQL + 'NuClien = ' "
'SQLX$ = SQLX$ + "SET @SQL = @SQL  +  Cast( @NuClien As Char(8)) "
'SQLX$ = SQLX$ + "EXEC(@SQL); "
'If EJECUTA_DIRECTO% > 0 Then
'  FLEXCONN.Execute (SQLX$)
'Else
'  If CAMBIOSP("UPDATENUCLIENBANC", SQLX$) Then
'   FLEXCONN.Execute (SQLX$)
'  End If
'End If
'
''----------------------------------------------------------------------------------------------------------------
'
'    If EXISTE_SP%("UPDATENUCLIENSADEU") > 0 Then
'      SQLX$ = "Alter "
'    Else
'      SQLX$ = "Create "
'      EJECUTA_DIRECTO% = 1
'    End If
'    SQLX$ = SQLX$ + " PROCEDURE UPDATENUCLIENSADEU "
'    SQLX$ = SQLX$ + "(@NuClien int = NULL "
'    SQLX$ = SQLX$ + " ,@NuClienRemp int = NULL "
'    SQLX$ = SQLX$ + " "
'    SQLX$ = SQLX$ + ") "
'    SQLX$ = SQLX$ + "AS "
'    SQLX$ = SQLX$ + "DECLARE @SQL NVARCHAR(MAX) "
'    SQLX$ = SQLX$ + "SET @SQL = ' UPDATE SADEUPEN SET ' "
'    SQLX$ = SQLX$ + "SET @SQL = @SQL + 'NuClien= ' "
'    SQLX$ = SQLX$ + "IF (@NuClien = '' OR @NuClien IS NULL) "
'    SQLX$ = SQLX$ + "BEGIN SET @SQL=  @SQL + ' NuClien' END "
'    SQLX$ = SQLX$ + "ELSE "
'    SQLX$ = SQLX$ + "SET @SQL = @SQL + CHAR(39) +  CAST(@NuClienRemp As Char(16)) + CHAR(39) "
'    SQLX$ = SQLX$ + " "
'    SQLX$ = SQLX$ + "SET @SQL = @SQL + ' WHERE ' "
'    SQLX$ = SQLX$ + "SET @SQL = @SQL + 'NuClien = ' "
'    SQLX$ = SQLX$ + "SET @SQL = @SQL  +  Cast( @NuClien As Char(8)) "
'    SQLX$ = SQLX$ + "EXEC(@SQL); "
'    If EJECUTA_DIRECTO% > 0 Then
'      FLEXCONN.Execute (SQLX$)
'    Else
'      If CAMBIOSP("UPDATENUCLIENSADEU", SQLX$) Then
'       FLEXCONN.Execute (SQLX$)
'      End If
'    End If
'
'
''------------------------------------------------------------------------------------------------------------------------------
'
'    'SP QUE ACTUAILZA AL IGUAL QUE EL CARGALISEL LA TABLA SQL_LISELEC
'    If EXISTE_SP%("ACTUDEUDOR1") > 0 Then
'      SQLX$ = "Alter "
'    Else
'      SQLX$ = "Create "
'      EJECUTA_DIRECTO% = 1
'    End If
'    SQLX$ = SQLX$ + " PROCEDURE ACTUDEUDOR1 (@CodiEmpr INT = 0)"
'    SQLX$ = SQLX$ + "AS "
'    SQLX$ = SQLX$ + " "
'    SQLX$ = SQLX$ + "DECLARE    @Codigo INT,  @raso_cli varchar(48), @txtx varchar(max), @txnum varchar(20),  @txrasocli varchar(124) "
'    SQLX$ = SQLX$ + " "
'    SQLX$ = SQLX$ + "SET @Codigo=0 "
'    SQLX$ = SQLX$ + "set @txtx = '' "
'    SQLX$ = SQLX$ + "DECLARE ElCursor CURSOR  FOR "
'    SQLX$ = SQLX$ + "SELECT nume_cli, raso_cli "
'    SQLX$ = SQLX$ + "FROM deudor12 "
'    SQLX$ = SQLX$ + "where Stat_Cli >= 0 and Nume_Cli>0 ORDER BY Raso_Cli "
'    SQLX$ = SQLX$ + "OPEN ElCursor FETCH NEXT FROM ElCursor INTO @Codigo, @raso_cli "
'    SQLX$ = SQLX$ + "WHILE (@@FETCH_STATUS = 0) BEGIN "
'    SQLX$ = SQLX$ + "set @txnum = space(6) "
'    SQLX$ = SQLX$ + "set @txnum =  @txnum +  rtrim(ltrim(Cast( @Codigo As Char(6)))) "
'    SQLX$ = SQLX$ + "set @txnum = right(@txnum,6) "
'    SQLX$ = SQLX$ + "SET @txnum =  @txnum + space(2) "
'    SQLX$ = SQLX$ + " "
'    SQLX$ = SQLX$ + "set @txrasocli =   ltrim(@raso_cli) "
'    SQLX$ = SQLX$ + "set @txrasocli = @txrasocli + space(48) "
'    SQLX$ = SQLX$ + "set @txrasocli = left(@txrasocli,48) "
'    SQLX$ = SQLX$ + " "
'    SQLX$ = SQLX$ + "set  @txtx =  @txtx + @txnum "
'    SQLX$ = SQLX$ + "set  @txtx = @txtx + @txrasocli + space(3) "
'    SQLX$ = SQLX$ + " "
'    SQLX$ = SQLX$ + "set @txnum='' "
'    SQLX$ = SQLX$ + "set @txrasocli ='' "
'    SQLX$ = SQLX$ + "FETCH NEXT FROM ElCursor INTO @Codigo, @raso_cli "
'    SQLX$ = SQLX$ + " "
'    SQLX$ = SQLX$ + "END "
'    SQLX$ = SQLX$ + "CLOSE ElCursor "
'    SQLX$ = SQLX$ + "DEALLOCATE ElCursor "
'    SQLX$ = SQLX$ + "update sql_liselec set LiSelec = @txtx, fech_actu = getdate() "
'    SQLX$ = SQLX$ + "where NomLiSele = 'deudor1' and  CodiEmpr = @CodiEmpr "
'    If EJECUTA_DIRECTO% > 0 Then
'      FLEXCONN.Execute (SQLX$)
'    Else
'      If CAMBIOSP("ACTUDEUDOR1", SQLX$) Then
'       FLEXCONN.Execute (SQLX$)
'      End If
'    End If
''------------------------------------------------------------------------------------------------------------------------------------------
'    If EXISTE_SP%("ACTUINDICLIE") > 0 Then
'      SQLX$ = "Alter "
'    Else
'      SQLX$ = "Create "
'      EJECUTA_DIRECTO% = 1
'    End If
'
'    SQLX$ = SQLX$ + " PROCEDURE ACTUINDICLIE (@CodiEmpr INT = 0)"
'    SQLX$ = SQLX$ + "AS "
'    SQLX$ = SQLX$ + " "
'    SQLX$ = SQLX$ + "DECLARE    @Codigo VARCHAR(16),  @raso_cli varchar(50), @txtx varchar(max), @txnum varchar(20),  @txrasocli varchar(124) "
'    SQLX$ = SQLX$ + " "
'    SQLX$ = SQLX$ + "SET @Codigo='' "
'    SQLX$ = SQLX$ + "set @txtx = '' "
'    SQLX$ = SQLX$ + "DECLARE ElCursor CURSOR STATIC LOCAL FORWARD_ONLY FOR "
'    SQLX$ = SQLX$ + "SELECT Codi_Cli, raso_cli "
'    SQLX$ = SQLX$ + "FROM deudor12 "
'    SQLX$ = SQLX$ + "where Stat_Cli >= 0 AND Codi_Cli <> '' ORDER BY Raso_Cli "
'    SQLX$ = SQLX$ + "OPEN ElCursor FETCH NEXT FROM ElCursor INTO @Codigo, @raso_cli "
'    SQLX$ = SQLX$ + "WHILE (@@FETCH_STATUS = 0) BEGIN "
'    SQLX$ = SQLX$ + "set @txnum =  rtrim(ltrim(@Codigo)) "
'    SQLX$ = SQLX$ + "set @txnum = @txnum + space(12) "
'    SQLX$ = SQLX$ + "set @txnum = LEFT(@txnum,12) "
'    SQLX$ = SQLX$ + "set @txnum =  @txnum + space(2) "
'    SQLX$ = SQLX$ + " "
'    SQLX$ = SQLX$ + "set @txrasocli =   ltrim(@raso_cli) "
'    SQLX$ = SQLX$ + "set @txrasocli = @txrasocli + space(50) "
'    SQLX$ = SQLX$ + "set @txrasocli = LEFT(@txrasocli,50) "
'    SQLX$ = SQLX$ + " "
'    SQLX$ = SQLX$ + "set  @txtx =  @txtx + @txnum "
'    SQLX$ = SQLX$ + "set  @txtx = @txtx + @txrasocli + space(3) "
'    SQLX$ = SQLX$ + " "
'    SQLX$ = SQLX$ + "set @txnum='' "
'    SQLX$ = SQLX$ + "set @txrasocli ='' "
'    SQLX$ = SQLX$ + "FETCH NEXT FROM ElCursor INTO @Codigo, @raso_cli "
'    SQLX$ = SQLX$ + " "
'    SQLX$ = SQLX$ + "END "
'    SQLX$ = SQLX$ + "CLOSE ElCursor "
'    SQLX$ = SQLX$ + "DEALLOCATE ElCursor "
'    SQLX$ = SQLX$ + "update sql_liselec set LiSelec = @txtx, fech_actu = getdate() "
'    SQLX$ = SQLX$ + "where NomLiSele = 'INDICLIE' and  CodiEmpr = @CodiEmpr "
'
'    If EJECUTA_DIRECTO% > 0 Then
'      FLEXCONN.Execute (SQLX$)
'    Else
'      If CAMBIOSP("ACTUINDICLIE", SQLX$) Then
'       FLEXCONN.Execute (SQLX$)
'      End If
'    End If
'
'End Sub
'
'Sub VALOBADACAJABANC(CodiCom_Lar, NuClien, NumeComp, TipoMovim)
'
'    If EXISTE_SP%("VALOCABAN") > 0 Then
'      SQLX$ = "alter Procedure VALOCABAN "
'    Else
'      SQLX$ = "CREATE Procedure VALOCABAN "
'    End If
'
'    CONDI$ = "NUMFACTURA=@DOCULA"
'
'    SQLX$ = SQLX$ + "("
'
'    SQLX$ = SQLX$ + "@DOCULA NVARCHAR(48)= CODICOM_LAR,"
'    SQLX$ = SQLX$ + "@NuClien int=NuClien inT"
'    SQLX$ = SQLX$ + "@DOCULA NVARCHAR(48)= CODICOM_LAR,"
'    SQLX$ = SQLX$ + "@DOCULA NVARCHAR(48)= CODICOM_LAR,"
'    SQLX$ = SQLX$ + "@DOCULA NVARCHAR(48)= COD ICOM_LAR,"
'
'    SQLX$ = SQLX$ + ")"
'    SQLX$ = SQLX$ + " AS"
'
'    SQLX$ = SQLX$ + " UPDATE CAJABANC set  NUMFACTURA = ''"
'    SQLX$ = SQLX$ + " Where " & CONDI$
'    FLEXCONN.Execute (SQLX$)

    End Sub



Function EXISTE_TRIGGER%(NAMETRIGGER$)
    EXISTE_TRIGGER% = 0
    
    'VALIDO LA VERSION POR QUE SI ES 2000 NO FUNCIONA DA ERROR,  ESTO SUCEDIO EN JYM
    SQLX$ = "select cast(@@version as varchar)"
    Set rs = FLEXCONN.Execute(SQLX$)
    Version1$ = rs.Fields(0)
    If InStr(Version1$, "2000") > 0 Then
      Exit Function
    End If
    '
    NAMETRIGGER$ = TRIM$(NAMETRIGGER$)
    EXISTE_TRIGGER% = 0
    On Error Resume Next
    SQLX$ = "SELECT * FROM SYS.SYSOBJECTS WITH(NOLOCK) WHERE TYPE = 'TR' AND Name = '" & NAMETRIGGER$ & "'"
    Set rs = READSET(SQLX$)
    With rs
      If Not .EOF Then
        EXISTE_TRIGGER% = 1
      End If
    End With
    On Error GoTo 0
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
    Dim rs As New ADODB.Recordset
    
    Set rs = FLEXCONN.Execute(SQLX$)
    Resultado$ = rs.Fields(0)
    If Resultado$ = "El objeto existe" Then EXISTE_SP% = 1
    rs.Close
    Set rs = Nothing
    

End Function

Function CAMBIOSP(NAME_SP$, TEXTONUEVO$) As Boolean
'-------------------------------------------------------------------------------------------------------------------
'VALIDA SI HUBO CAMBIOS EN EL STORED PROCEDURE

    TXNEW$ = UCase$(TEXTONUEVO$)
    SQLX$ = "exec sp_helptext " & NAME_SP$
    Dim rs1 As ADODB.Recordset
    Set rs1 = New ADODB.Recordset
    On Error Resume Next
    rs1.Open SQLX$, FLEXCONN, adOpenKeyset, adLockReadOnly, adCmdText
    If Abs(Err.Number) > 0 Then On Error GoTo 0: GoTo 90
    On Error GoTo 0
    TEXTOEXISTENTE$ = ""
    With rs1
      Do While Not .EOF
        TEXTOEXISTENTE$ = TEXTOEXISTENTE$ & rs1!TEXT
        .MoveNext
      Loop
    End With
    rs1.Close
    Set rs1 = Nothing
    'EN EL SQLSERVER QUEDA COMO CREATE POR ESO LO CAMBIO PARA LA COMPARACION
    
    If UCase$(Left$(TRIM$(TXNEW$), 6)) = "ALTER " Then
       TXNEW$ = REPLACAR(TXNEW$, "ALTER ", "CREATE ")
    End If
    'EN EL SQLSERVER LE AGREGA [dbo].[AB] (LO QUITO PARA LA COMPARACION)
    TEXTOEXISTENTE$ = REPLACAR$(TEXTOEXISTENTE$, "[dbo].[" & NAME_SP$ & "]", NAME_SP$)
    TEXTOEXISTENTE$ = UCase(TEXTOEXISTENTE$)
    If TEXTOEXISTENTE$ = TXNEW$ Then
       CAMBIOSP = False
    Else
90     CAMBIOSP = True
    End If
    
End Function
'
'EJEMPLOS DE PROGRAMACION DE STORED PROCEDURE
'EJEMPLO DE UN PROCEDIMIENTO QUE SEGUN LAS CONDICIONES RETORNA DIFRENTES VALORES.

'Creamos un procedimiento que reciba un documento y nos retorne distintos valores según: sea nulo (1), no sea válido (2),
'no esté en la tabla "socios" (3), sea un socio deudor (4) o sea un socio sin deuda (0):

'ALTER procedure [dbo].[pa_documentovalido]
'  @documento char(8)=null
'  as
'   if @documento is null return 1
'   Else
'    if len(@documento)<8 return 2
'    Else
'     if not exists (select *from socios where documento=@documento) return 3
'     Else
'     begin
'      if exists (select *from inscriptos
'                where documentosocio=@documento and
'                matricula='n') return 4
'      else return 0
'     end;
'
'CODIGO QUE RETORNA EL VALOR RETURN DEL SP DE ARRIBA
'DECLARE @return_status int;
'EXEC @return_status = pa_documentovalido '30000000';
'SELECT 'Return Status' = @return_status;
'GO
'ahp
' 'DECLARE @saldo_actual DECIMAL(10, 2)DECLARE @cuenta_corriente
' 'TABLE (ID INT IDENTITY(1, 1),CANTINGRE DECIMAL(10, 2),CANTSALID DECIMAL(10, 2),
' 'Diferencia DECIMAL(10, 2),SALDO_ACTUAL DECIMAL(10, 2),IDENLOTE VARCHAR(20))
' 'INSERT INTO @cuenta_corriente (CANTINGRE, CANTSALID,IDENLOTE)
' 'SELECT CANTINGRE, CANTSALID,IDENLOTE FROM MOVISTO WITH(NOLOCK)
' 'Where COD_INTE = 24587 AND IDENLOTE ='' ORDER BY CANTSALID,MOVISTO_ID SET @saldo_actual = 0
' 'UPDATE @cuenta_corriente SET Diferencia = @saldo_actual, SALDO_ACTUAL = @saldo_actual + CANTINGRE - CANTSALID
' ', @saldo_actual = @saldo_actual + CANTINGRE - CANTSALID SELECT CANTINGRE, CANTSALID, ID,Diferencia,IDENLOTE
' 'FROM @cuenta_corriente  ORDER BY ID
'    'MUESTRA UN RECORDSET, MOSTRANDO PRIMERO LOS INGRESOS Y LUEGO LOS CONSUMOS
'    'PARA VER SI TIENE NEGATIVOS PARA EN OTRO PROCESO REPARAR DICHOS NEGATIVOS APLICANDO LOTES CON SALDO POSITIVO
'    If EXISTE_SP%("CalculaSaldoLotes") > 0 Then
'      SQLX$ = "alter Procedure CalculaSaldoLotes "
'    Else
'      SQLX$ = "CREATE Procedure CalculaSaldoLotes "
'      EJECUTA_DIRECTO% = 1
'    End If
'
'    CONDI$ = "  COD_INTE = @COD_INTE "
'    CONDI$ = CONDI$ + " AND IDENLOTE = @IDENLOTE "
'    CONDI$ = CONDI$ + " AND Codiempr= @Codiempr"
'
'    '
'    SQLX$ = SQLX$ + "("
'    SQLX$ = SQLX$ + " @cod_inte INT,"
'    SQLX$ = SQLX$ + " @idenlote VARCHAR(20)"
'    SQLX$ = SQLX$ + " AS"
'    SQLX$ = SQLX$ + " DECLARE @saldo_actual DECIMAL(10, 2)DECLARE @cuenta_corriente"
'    SQLX$ = SQLX$ + " TABLE (ID INT IDENTITY(1, 1),CANTINGRE DECIMAL(10, 2),CANTSALID DECIMAL(10, 2),"
'    SQLX$ = SQLX$ + " Diferencia DECIMAL(10, 2),SALDO_ACTUAL DECIMAL(10, 2),IDENLOTE VARCHAR(20))"
'
'    SQLX$ = SQLX$ + " INSERT INTO @cuenta_corriente (CANTINGRE, CANTSALID,IDENLOTE)"
'    SQLX$ = SQLX$ + " SELECT CANTINGRE, CANTSALID,IDENLOTE FROM MOVISTO WITH(NOLOCK)"
'    SQLX$ = SQLX$ + " WHERE " & CONDI$
'    SQLX$ = SQLX$ + " UPDATE @cuenta_corriente SET Diferencia = @saldo_actual, SALDO_ACTUAL = @saldo_actual + CANTINGRE - CANTSALID"
'    SQLX$ = SQLX$ + " , @saldo_actual = @saldo_actual + CANTINGRE - CANTSALID"
'    SQLX$ = SQLX$ + " SELECT CANTINGRE, CANTSALID, ID,Diferencia,IDENLOTE"
'    SQLX$ = SQLX$ + " FROM @cuenta_corriente  ORDER BY ID"
'    SQLX$ = SQLX$ + ")"
'
'
'    SQLX$ = SQLX$ + " SELECT MAX(DOPRINUM) as maximo FROM MOVISTO WITH(NOLOCK) "
'    SQLX$ = SQLX$ + " WHERE " & CONDI$
'    If EJECUTA_DIRECTO% > 0 Then
'      FLEXCONN.Execute (SQLX$)
'    Else
'      If CAMBIOSP("MAXIMO_N_REMITO", SQLX$) Then
'        FLEXCONN.Execute (SQLX$)
'      End If
'    End If


