Attribute VB_Name = "CLIE_FUN_SQL"
' ACCDEUDO.QBS                         ' accesos varios a archivo de deudores
'
Public ARCHE$(16)
'Public NC As Integer
'Public MODO As Integer

'Public NROPRO&
'Public NCLILLAMA%
'
Public VERICUIT%
'
Public TRANACT%    ' NUMERO DE TRANSPORTE ACTIVO - PARA PODER CAMBIARLO EN EL PEDIDO
Dim ClienTemp As ADODB.Recordset
Dim YACRECON%
Private Declare Sub Sleep Lib "KERNEL32" (ByVal dwMilliseconds As Long)



Function ImporteFacturaAplicada#(NCL, FACTU$)
  'EN PRINCIPIO LO ESTOY UTILIZANDO PARA SABER EL IMPORTE QUE SE APICO DE UNA NOTA DE CREDITO PARA RESTAR DEL SALDO
  'QUE PUEDE QUE NO SEA EL MISMO QUE LA N.C.  POR QUE SE APLICA HASTA EL MAXIMO PENDIENTE.
  CONDI$ = "IMPOHABE > 0 AND  DOCU_APLI = '" & FACTU$ & "' And NUCLIEN = " & NCL
  CONDI$ = CONDI$ + " AND TIPOMOVIM = 'AJUS' OR TIPOMOVIM = 'FVEN'"
  ImporteFacturaAplicada# = XVALO(VALOBADA("FACTUANTI", "IMPOHABE", CONDI$, "NOMESS"))
  
End Function

Sub MODICUENTA(NumeClie, Credito$, FACTURA$)
   CONDI$ = "NuClien = " & NumeClie & " AND CODICOM_LAR =  '" & FACTURA$ & "' AND TIPOMOVIM = 'FVEN'"
   CUENTAFVEN% = XVALO(VALOBADA("CAJABANC", "CUECONTAI", CONDI$, "NOMESS"))
   
   CONDI$ = "NuClien = " & NumeClie & " AND CODICOM_LAR =  '" & Credito$ & "' AND TIPOMOVIM = 'RCOB'"
   Call ACTUREGISTRO("CAJABANC", "CUECONTAI", CONDI$, CUENTAFVEN%, RegAf&, "NOMESS")
   '
'   CONDI$ = "NuClien = " & NumeClie & " AND CODICOM_LAR =  '" & CREDITO$ & "' AND TIPOMOVIM = 'CHRECI'"
'   Call ACTUREGISTRO("CAJABANC", "CUECONTAI", CONDI$, CUENTAFVEN%, REGAF&, "NOMESS")
'   '
'   CONDI$ = "NuClien = " & NumeClie & " AND CODICOM_LAR =  '" & FACTURA$ & "' AND TIPOMOVIM = 'VENTA'"
'   CUENTAVENTA% = XVALO(VALOBADA("CAJABANC", "CUECONTAI", CONDI$, "NOMESS"))

End Sub

Sub MONE_TIPOCAM_GASTOFLETE(CONDI$, MONE%, TIPOCAM#, GASTOFLETE#)
       'SOLO CASOS QUE DEVUELVA UN SOLO REGISTRO
       'EJEMPLO DE CONDICION PARA AHP  CONDIFLET$ = "CODICOM_LAR = '" & NUFACTU$ & "' AND TIPOMOVIM = 'COMPRA' and NUPROVE = " & NPROV& & "   AND LEFT(DESCRIMOV,13) = 'Flete Export.'"

       SQLX$ = "SELECT IDEBECOMP,MONE_EMIS,TIPO_CAM FROM CAJABANC WITH(NOLOCK) WHERE " & CONDI$
       Set rs = READSET(SQLX$)
       With rs
         If Not rs.EOF Then
            MONE% = XVALO(!Mone_Emis)
            TIPOCAM# = XVALO(!Tipo_Cam)
            GASTOFLETE# = XVALO(!Idebecomp)
         End If
       End With
       

End Sub

Function ProveedorRealExterior%(NPROV&)
   'DEVUELVE 0 SI EL PROVEEDOR ES EL MINEPARTS (SEGUN CONFIGURACION)
   'DEVUELVE 1 SI EL PROVEEDOR ES MINEPARTS (SEGUN CONFIGURACION)
   Static NprovMineparts
   If NprovMineparts < 1 Then
      NprovMineparts = XVALO(Control("CFGMINEP", "PRFEMINE"))
   End If
   '
   If NPROV& <> NprovMineparts Then ProveedorRealExterior% = 0
   If NPROV& = NprovMineparts Then ProveedorRealExterior% = 1
   
End Function

Function codProvinciaEntrega$(NC)
   CONDI$ = "NUME_CLI = " & NC
   codProvinciaEntrega$ = VALOBADA("DEUDOR12", "ProvEntre_Cli", CONDI$, "NOMESS")
End Function

Function TIENE_FACTURAS&(NC, DIAX%, FEULT_FACT%)


   'DEVUELVE 1 O MAS  SI COMPRO DENTRO DE LOS DIA PASADO COMO PARAMETRO AL DIA DE HOY
   'DEVUELVE 0 SI NO COMPRO DENTRO DE LOS DIA PASADO COMO PARAMETRO AL DIA DE HOY
    Des% = DIANTE(HOY(HO$), DIAX%)
    Desde = FETEXCOR1$(Des%)
    Hasta = FETEXCOR1$(HOY(HO$))
    '
    CONDI$ = "TIPOMOVIM = 'FVEN' AND FECHEMISI >= '" & Desde & "'"
    CONDI$ = CONDI$ + " AND FECHEMISI <= '" & Hasta & "'"
    CONDI$ = CONDI$ + " AND LEFT(CODICOM_LAR ,3) = 'FC-'"
    CONDI$ = CONDI$ + " AND NUCLIEN = " & NC
    '
    TIENE_FACTURAS& = CantRegistros&("CAJABANC", CONDI$, "NOMESS")
    '
    CONDI$ = "TIPOMOVIM = 'FVEN' "
    CONDI$ = CONDI$ + " AND NUCLIEN = " & NC
    CONDI$ = CONDI$ + " ORDER BY FECHEMISI DESC"
    
    FEULT_FACT% = CVINT(VALOBADA("CAJABANC", "FECHEMISI", CONDI$))
End Function


'
Function NUMECLI&(COPRO$)
   '
   NPRX& = 0
   CPRX$ = UCase$(TRIM$(COPRO$))
   If CPRX$ <> "" Then
     Call APERBASDAT("CLIEN")
     'sql
     Set ClienTemp = FLEXCONN.Execute(FILTSQL$("SELECT * FROM Deudor12 WHERE Codi_Cli = '" & CPRX$ & "'"))
     If Not (ClienTemp.EOF And ClienTemp.BOF) Then 'si la encuentra
       NPRX& = Abs(ClienTemp!nume_cli)
     End If
   End If
   '
90 NUMECLI& = NPRX&
   '
   
End Function
'
Function NUMEPRO%(COPRO$)
   '
   NPRX% = 0
   CPRX$ = UCase$(TRIM$(COPRO$))
   If CPRX$ <> "" Then
     Call APERBASDAT("PROVE")
     NPRX% = XVALO(VALOBADA("Proved12", "Nume_Cli", "Codi_Cli = '" & CPRX$ & "'", "NOMESS"))
   End If
   '
90 NUMEPRO% = NPRX%
   '
End Function

Function STATCLI%(NCLIE)
   '
   Static NCLIEA&, STATUIA%
   '
   ' = 1 --> Existe y está activo
   ' =-1 --> Existe y esta de baja
   ' = 0 --> No existe el item
   '
   NCLX& = Int(NCLIE)
   If NCLX& = NCLIEA& Then
     If STATUIA% > 0 Then
       STATCLI% = STATUIA%
       'Exit Function 'COMENTADO POR QUE CUANDO SE VUELVE A CONSULTAR SI ES MISMO CLIENTE DEVUELVE VACIO
     End If
   End If
   '
   STATCLI% = 0
   If NCLX& <> 0 Then
     Set ClienTemp = New ADODB.Recordset
     ClienTemp.CursorType = adOpenKeyset
     ClienTemp.LockType = adLockOptimistic
     '
     If NCLX& > 0 Then
          Call APERBASDAT("CLIEN")
          Set ClienTemp = FLEXCONN.Execute(FILTSQL$("SELECT * FROM Deudor12 WITH(NOLOCK) WHERE Nume_Cli = " & Abs(NCLX&)))
       Else
          Call APERBASDAT("PROVE")
          Set ClienTemp = FLEXCONN.Execute(FILTSQL$("SELECT * FROM Proved12 WITH(NOLOCK) WHERE Nume_Cli = " & Abs(NCLX&)))
     End If
     '
     If Not (ClienTemp.EOF And ClienTemp.BOF) Then 'si la encuentra
        STATCLI = 1
        If ClienTemp!Stat_Cli < 1 Then
           STATCLI = (-1)
        End If
     End If
     '
   End If
   '
90 NCLIEA& = NCLX&
   STATUIA% = STATCLI%
   
End Function
'
Function REGICLI&(NC)
   ' DEVUELVE CERO SI NO EXISTE Y 1 SI EXISTE AUNQUE ESTÉ DE BAJA
   RGCL& = Abs(STATCLI%(NC))
   REGICLI& = RGCL&
End Function
'
Function CODICLI$(NC)
   On Error Resume Next
   If STATCLI%(NC) <> 0 Then CODICLI$ = TRIM$(ClienTemp!Codi_Cli)
End Function
'
Function CODPROSPE$(DIELE$)
    '
    CPRX$ = ""
    If DIELE$ = "" Then GoTo 90
    '
    ' primero busca en la direccion principal de mail
    CONDI$ = "Mail_Cli = '" & DIELE$ & "'"
    CPRX$ = TRIM$(VALOBADA("BSEGCOM", "CODI_CLI", CONDI$, "NOMESS"))
    '
    If CPRX$ = "" Then
       CONDI$ = "MAIL_CTAC = '" & DIELE$ & "'"
       CPRX$ = TRIM$(VALOBADA("CONTACSCOM", "CODICLI", CONDI$, "NOMESS"))
    End If
    '
    If CPRX$ = "" Then
       CONDI$ = "SECTOR1 = '" & DIELE$ & "'"
       CPRX$ = TRIM$(VALOBADA("BSEGCOM", "CODI_CLI", CONDI$, "NOMESS"))
    End If
    '
    If CPRX$ = "" Then
       CONDI$ = "SECTOR2 = '" & DIELE$ & "'"
       CPRX$ = TRIM$(VALOBADA("BSEGCOM", "CODI_CLI", CONDI$, "NOMESS"))
    End If
    '
    If CPRX$ = "" Then
       CONDI$ = "SECTOR2 = '" & DIELE$ & "'"
       CPRX$ = TRIM$(VALOBADA("BSEGCOM", "CODI_CLI", CONDI$, "NOMESS"))
    End If
    '
90  CODPROSPE$ = CPRX$
    '
End Function

Function rasocli$(NC)
'    If NC > 0 Then
'      On Error Resume Next
'      Dim vec As New clientes
'      If Err > 0 Then
'          On Error GoTo 0
'          GoTo 10
'      Else
'        rasocli$ = TRIM$(VectCliente.nombre(NC))
'        Exit Function
'      End If
'    End If
'      nom$ = vec.nombre(50034)
'        dom1$ = vec.domicilio(50034)
'        Loc1$ = vec.localidad(50034)

10 On Error Resume Next
   If STATCLI%(NC) <> 0 Then rasocli$ = TRIM$(ClienTemp!raso_cli)
End Function
'
Function DOMICLI$(NC)
   On Error Resume Next
   If STATCLI%(NC) <> 0 Then DOMICLI$ = TRIM$(ClienTemp!domi_cli)
End Function
'
Function CPOSCLI$(NC)
   On Error Resume Next
   If STATCLI%(NC) <> 0 Then CPOSCLI$ = TRIM$(ClienTemp!Cpos_Cli)
End Function
'
Function LOCACLI$(NC)
   On Error Resume Next
   If STATCLI%(NC) <> 0 Then LOCACLI$ = TRIM$(ClienTemp!loca_Cli)
End Function
'
Function CPLOCLI$(NC)
   CPX$ = TRIM$(CPOSCLI$(NC))
   If CPX$ <> "" Then CPX$ = "(" + CPX$ + ") "
   CPLOCLI$ = TRIM$(CPX$ + LOCACLI$(NC))
End Function
'
Function PAISCLI$(NC)
   On Error Resume Next
   If STATCLI%(NC) <> 0 Then PAISCLI$ = TRIM$(ClienTemp!Pais_Cli)
End Function
'
Function TEL1CLI$(NC)
   On Error Resume Next
   If STATCLI%(NC) <> 0 Then TEL1CLI$ = TRIM$(ClienTemp!Tele_Cli)
End Function
'
Function TEL2CLI$(NC)
   On Error Resume Next
   If STATCLI%(NC) <> 0 Then TEL2CLI$ = TRIM$(ClienTemp!Fant_Cli)
End Function
'
Function ABRECLI$(NC)
   ABRECLI$ = TEL2CLI$(NC)
End Function
'
Function FANTCLI$(NC)
   FANTCLI$ = TEL2CLI$(NC)
   If FANTCLI$ = "" Then
     FANTCLI$ = UCase$(rasocli$(NC))
     PPXX% = InStr(FANTCLI$, " ")
     If PPXX% > 1 Then
       FANTCLI$ = Left$(FANTCLI$, PPXX% - 1)
     End If
   End If
End Function
'
Function FAXCLI$(NC)
   On Error Resume Next
   If STATCLI%(NC) <> 0 Then FAXCLI$ = TRIM$(ClienTemp!Faxi_Cli)
End Function
'
Function EMAILCLI$(NC)
   On Error Resume Next
   If STATCLI%(NC) <> 0 Then EMAILCLI$ = TRIM$(ClienTemp!Mail_Cli)
   '
   If TRIM$(EMAILCLI$) = "" Then ' SI NO TIENE MAIL BUSCA EN EL PROSPECT
     PROSPECT$ = CODICLI(NC)
     EMAILCLI$ = VALOBADA("BSEGCOM", "Mail_Cli", "Codi_Cli='" & PROSPECT$ & "'", "NOMESS")
   End If
End Function
'
Function CONTACTCLI$(NC)
   On Error Resume Next
   If STATCLI%(NC) <> 0 Then CONTACTCLI$ = TRIM$(ClienTemp!CTAC_CLI)
End Function
Function LUENTRECOMPRA$(NC)
   'LUGAR DE ENTREGA COMPRAS
   On Error Resume Next
   If STATCLI%(NC) <> 0 Then LUENTRECOMPRA$ = ClienTemp!LUENTCOMPRAS
End Function
Function OBSERVACOMPRA$(NC)
   'OBSERVACION DE ORDEN DE COMPRA
   On Error Resume Next
   If STATCLI%(NC) <> 0 Then OBSERVACOMPRA$ = ClienTemp!OBSCOMPRAS
End Function
'
Function HORARCLI$(NC)
   On Error Resume Next
   If STATCLI%(NC) <> 0 Then HORARCLI$ = TRIM$(ClienTemp!Hora_Cli)
End Function
'
Function CuitCli$(NC)
   On Error Resume Next
   If STATCLI%(NC) <> 0 Then CuitCli$ = TRIM$(ClienTemp!Cuit_Cli)
End Function
'
Function RegimenFacturaCredito%(NC)
    On Error Resume Next
    If STATCLI%(NC) <> 0 Then RegimenFacturaCredito% = ClienTemp!RegimenFacturaCredito
End Function
Function NIBRCLI$(NC)
   On Error Resume Next
   If STATCLI%(NC) <> 0 Then NIBRCLI$ = TRIM$(ClienTemp!Nibr_Cli)
End Function
'
Function JURISIIBBCLI$(NC)
   On Error Resume Next
   If STATCLI%(NC) <> 0 Then JURISIIBBCLI$ = TRIM$(ClienTemp!JurisdiccioIIBB_CLI)
End Function

Function NPROCLI$(NC)
   On Error Resume Next
   If STATCLI%(NC) <> 0 Then NPROCLI$ = TRIM$(ClienTemp!Prvd_Cli)
End Function
'
Function PROVCLI$(NC)
   On Error Resume Next
   If STATCLI%(NC) <> 0 Then PROVCLI$ = TRIM$(ClienTemp!Prov_Cli)
End Function
'
Function CPAGCLI%(NC)
   On Error Resume Next
   If STATCLI%(NC) <> 0 Then CPAGCLI% = ClienTemp!Cpag_Cli
   If CPAGCLI < 1 Then
     If NC > 0 Then
         CPAGCLI% = XVALO(Control$("AMACLI", "CONPAG")) ' PARA QUE TOME POR DEFAULT
       Else
         CPAGCLI% = XVALO(Control$("AMAPRO", "CONPAG")) ' PARA QUE TOME POR DEFAULT
     End If
   End If
End Function
'
Function CONDICOM$(NC)
   On Error Resume Next
   If STATCLI%(NC) <> 0 Then CONDICOM$ = ClienTemp!CCom_Cli
End Function
'
Function LIPRCLI%(NC)
   On Error Resume Next
   If STATCLI%(NC) <> 0 Then LIPRCLI% = ClienTemp!Lpre_Cli
   If LIPRCLI% < 1 Then
     LIPRCLI% = XVALO(Control$("AMACLI", "LISTAS")) ' PARA QUE TOME POR DEFAULT
   End If
End Function
'
Function TICUEPRO%(NC)
   On Error Resume Next
   If STATCLI%(NC) <> 0 Then TICUEPRO% = ClienTemp!Tcpr_Cli
End Function
'
Function VENDCLI%(NC)
   On Error Resume Next
   If STATCLI%(NC) <> 0 Then VENDCLI% = TRIM$(ClienTemp!Vend_Cli)
   If VENDCLI% < 0 Or VENDCLI% > 255 Then
      Call MENSERR(24, "Número de Vendedor no Válido\en Cliente Número " + TRIM$(Str$(NC)) + ".")
      VENDCLI% = 0
   End If
End Function
'
Function PIVACLI%(NC)
   On Error Resume Next
   If STATCLI%(NC) <> 0 Then PIVACLI% = TRIM$(ClienTemp!Piva_Cli)
End Function
'
Function CATIBRU$(NC)
   On Error Resume Next
   If STATCLI%(NC) <> 0 Then CATIBRU$ = TRIM$(ClienTemp!Caib_Cli)
End Function
'
Function LIMICRE(NC)
   LIMICRE = 0
   On Error Resume Next
   If STATCLI%(NC) <> 0 Then LIMICRE = CDbl(ClienTemp!Lcre_Cli)
End Function
'
Function DESCUCLI$(NC)
   On Error Resume Next
   If STATCLI%(NC) <> 0 Then DESCUCLI$ = TRIM$(ClienTemp!Dsta_Cli)
End Function
'
Function SALDCLI#(NC)
    Call ABRECABAN
'    ESTO QUE SIGUE SE MODIFICO EL 1/4/09 PARA EFICIENTIZAR - EPB
'    SQLX$ = "SELECT SUM(IDEBECOMP) AS SUDEBE, SUM(IHABECOMP) AS SUHABE FROM CAJABANC "
'    SQLX$ = SQLX$ + " WHERE NuClien = " & NC & " "
'    SQLX$ = SQLX$ + " AND (TipoMovim = 'FVEN' OR TipoMovim = 'RCOB' OR TipoMovim = 'AJUD') "
'    Dim RSX As ADODB.Recordset
'    Set RSX = FLEXCONN.Execute(SQLX$)
'
'    SALDCLI# = 0
'    On Error Resume Next
'    SALDCLI# = XVALO(RSX!Sudebe) - XVALO(RSX!Suhabe)
'    On Error GoTo 0
'    '
    CONDI$ = "NuClien = " & NC & " AND (TipoMovim = 'FVEN' OR TipoMovim = 'RCOB' OR TipoMovim = 'AJUD') "
    SUDE# = XVALO(VALOBADA("CAJABANC", "SUM(IDEBECOMP)", CONDI$))
    SUHA# = XVALO(VALOBADA("CAJABANC", "SUM(IHABECOMP)", CONDI$))
    SALDCLI# = SUDE# - SUHA#
    
End Function

Function SALDPRO#(NC)
    Call ABRECABAN
    
    CONDI$ = "NuProve = " & NC & " AND (TipoMovim = 'FCOM' OR TipoMovim = 'OPAG' OR TipoMovim = 'AJUP') "
    SUDE# = XVALO(VALOBADA("CAJABANC", "SUM(IDEBECOMP)", CONDI$))
    SUHA# = XVALO(VALOBADA("CAJABANC", "SUM(IHABECOMP)", CONDI$))
    SALDPRO# = SUDE# - SUHA#
    
End Function
'
Function PORDESCLI(NC)
    '
    PDXC = 0
    DESCX$ = REPLACAR$(DESCUCLI$(NC), "%", " ")
    '
    If InStr(DESCX$, "+") < 1 Then
        PDXC = XVALO(DESCX$)
      Else
        XDC% = InStr(DESCX$, "+")
        DESC1 = XVALO(Left$(DESCX$, XDC% - 1))
        DESC2 = XVALO(Mid$(DESCX$, XDC% + 1))
        PDXC = 100 * (1 - (1 - DESC1 / 100) * (1 - DESC2 / 100))
    End If
    '
99  PORDESCLI = PDXC
    '
End Function
'
Function PORCEDESCU(TDESCUEN$)
    '
    coefi = 1
    DESCX$ = TRIM$(REPLACAR$(TDESCUEN$, "%", ""))
    '
    Do While Len(DESCX$) > 0
       PPXX% = InStr(DESCX$, "+")
       If PPXX% < 1 Then PPXX% = 1 + Len(DESCX$)
       AKX1$ = Left$(DESCX$, PPXX% - 1)
       DESCX$ = TRIM$(Mid$(DESCX$, PPXX% + 1))
       '
       PDC = XVALO(AKX1$)
       coefi = coefi * (1 - PDC / 100)
    Loop
    '
99  PORCEDESCU = 100 * (1 - coefi)
    '
End Function
'
Function LUENTRECLI$(NC, OPCI%)
   '
   LUENTRECLI$ = ""
   On Error Resume Next
   If STATCLI%(NC) <> 0 Then
     If OPCI% = 0 Then
         LUENTRECLI$ = TRIM$((ClienTemp!Lentre0_Cli))
         If LUENTRECLI$ = "" Then
           LTXX$ = FILTERGETRECORD$("LENTRECL", 1, MKI$(Int(NC)))
           LUENTRECLI$ = TRIM$(Mid$(LTXX$, 3, 32))
         End If
         'CONFIGURACION PARA VERIFICAR SI PRESENTA EL DOMICILIO FISCAL COMO DOMICILIO DE ENTREGA EN EL CASO QUE ESTE VACIO
         If TRIM$(Control("MASTER", "ENTREPFI")) <> "NO" Then
            If LUENTRECLI$ = "" Then LUENTRECLI$ = DOMICLI$(NC)
         End If
       ElseIf OPCI% = 1 Then
         LUENTRECLI$ = (ClienTemp!Lentre1_Cli)
         If LUENTRECLI$ = "" Then
           LTXX$ = FILTERGETRECORD$("LENTRECL", 1, MKI$(Int(NC)))
           LUENTRECLI$ = TRIM$(Mid$(LTXX$, 33, 32))
         End If
         If TRIM$(Control("MASTER", "ENTREPFI")) <> "NO" Then
           If LUENTRECLI$ = "" Then LUENTRECLI$ = CPOSCLI$(NC) + " - " + LOCACLI$(NC)
         End If
       ElseIf OPCI% = 2 Then
         NTRX% = NUTRANCLI%(NC)
         If NTRX% < 1 Then
             LUENTRECLI$ = (ClienTemp!Lentre2_Cli)
             If LUENTRECLI$ = "" Then
               LTXX$ = FILTERGETRECORD$("LENTRECL", 1, MKI$(Int(NC)))
               LUENTRECLI$ = TRIM$(Mid$(LTXX$, 67, 32))
             End If
         Else
               LUENTRECLI$ = ECOARCH$("TRANSPOR", MKI$(NTRX%))
         End If
       ElseIf OPCI% = 3 Then
         LUENTRECLI$ = HORAENT$(NC)
       ElseIf OPCI% = 4 Then
         LUENTRECLI$ = (ClienTemp!Lentre3_Cli)
     End If
   End If
   '
End Function
'
Function DOMIENT$(NC)
   DOMIENT$ = LUENTRECLI$(NC, 0)
End Function
'
Function LOCAENT$(NC)
   LOCAENT$ = LUENTRECLI$(NC, 1)
End Function
'
Function CODPOSTALENT$(NC)
   CODPOSTALENT$ = LUENTRECLI$(NC, 4)
End Function
'
Function TRANENT$(NC)
   TRANENT$ = LUENTRECLI$(NC, 2)
End Function
'
Function NUTRANCLI%(NC, Optional SUC%)
   If SUC% = 0 Then
        If TRANACT% > 0 Then
            NUTRANCLI% = TRANACT%
          Else
            On Error Resume Next
            If STATCLI%(NC) <> 0 Then NUTRANCLI% = TRIM$(ClienTemp!NTrans_Cli)
        End If
    Else
        CONDI$ = "NUME_CLI=" & NC
        NUTRANCLI% = XVALO(VALOBADA("SUCENTRE", "CODI_TRA", CONDI$))
'        RFF$ = RECFILT$("SUCENTRE", 5, MKI$(XVALO(NC)))
'        For I1& = 1 To Len(RFF$) Step 4
'          I& = CVS(Mid$(RFF$, I1&))
'          X$ = REGLEIDO$("SUCENTRE", I&)
'          SUCU1% = CVI(Mid$(X$, 1, 2))
'          If SUCU1% = SUC% Then
'            NUTRANCLI% = CVI(Mid$(X$, 195, 2))
'            Exit For
'          End If
'        Next I1&
    End If
End Function
'
Function DENOTRANS$(NTRANI%)
     XX$ = VALOBADA("TRANSPORTES", "DENO_TRA", "CODI_TRA = " & NTRANI%, "NOMESS")
     DENOTRANS$ = TRIM$(XX$)
     
     If DENOTRANS$ = "" Then
        XX$ = FILTERGETRECORD$("TRANSPOR", 1, MKI$(NTRANI%))
        DENOTRANS$ = TRIM$(Mid$(XX$, 3, 32))
     End If
End Function

Function RASOTRANS$(NTRANI%)
    XX$ = VALOBADA("TRANSPORTES", "RASO_TRA", "CODI_TRA = " & NTRANI%)
    RASOTRANS$ = TRIM$(XX$)
    
    If RASOTRANS$ = "" Then
       XX$ = FILTERGETRECORD$("TRANSPOR", 1, MKI$(NTRANI%))
       RASOTRANS$ = TRIM$(Mid$(XX$, 35, 48))
    End If
End Function

Function DOMITRANS$(NTRANI%)
    XX$ = VALOBADA("TRANSPORTES", "DOMI_TRA", "CODI_TRA = " & NTRANI%, "NOMESS")
    DOMITRANS$ = TRIM$(XX)
    
    If DOMITRANS$ = "" Then
       XX$ = FILTERGETRECORD$("TRANSPOR", 1, MKI$(NTRANI%))
       DOMITRANS$ = TRIM$(Mid$(XX$, 83, 48))
    End If
End Function

Function LOCATRANS$(NTRANI%)
    XX$ = VALOBADA("TRANSPORTES", "LOCA_TRA", "CODI_TRA = " & NTRANI%, "NOMESS")
    LOCATRANS$ = TRIM$(XX$)
    
    If LOCATRANS$ = "" Then
       XX$ = FILTERGETRECORD$("TRANSPOR", 1, MKI$(NTRANI%))
       LOCATRANS$ = TRIM$(Mid$(XX$, 131, 48))
    End If
End Function

Function TELETRANS$(NTRANI%)
    XX$ = VALOBADA("TRANSPORTES", "TELE_TRA", "CODI_TRA = " & NTRANI%, "NOMESS")
    TELETRANS$ = TRIM$(XX$)
    
    If TELETRANS$ = "" Then
       XX$ = FILTERGETRECORD$("TRANSPOR", 1, MKI$(NTRANI%))
       TELETRANS$ = TRIM$(Mid$(XX$, 195, 48))
    End If
End Function

Function CTACTRANS$(NTRANI%)
    XX$ = VALOBADA("TRANSPORTES", "CTAC_TRA", "CODI_TRA = " & NTRANI%, "NOMESS")
    CTACTRANS$ = TRIM$(XX$)
    
    If CTACTRANS$ = "" Then
       XX$ = FILTERGETRECORD$("TRANSPOR", 1, MKI$(NTRANI%))
       CTACTRANS$ = TRIM$(Mid$(XX$, 339, 48))
    End If
End Function

Function CUITTRANS$(NTRANI%)
    XX$ = VALOBADA("TRANSPORTES", "CUIT_TRA", "CODI_TRA = " & NTRANI%, "NOMESS")
    CUITTRANS$ = TRIM$(XX$)
    
    If CUITTRANS$ = "" Then
       XX$ = FILTERGETRECORD$("TRANSPOR", 1, MKI$(NTRANI%))
       CUITTRANS$ = TRIM$(Mid$(XX$, 179, 15))
    End If
End Function
Function PROVTRANS$(NTRANI%) 'PROVINICIA DEL TRANSPORTE
    XX$ = VALOBADA("TRANSPORTES", "Prov_Tra", "CODI_TRA = " & NTRANI%, "NOMESS")
    PROVTRANS$ = TRIM$(XX$)
    
   End Function
Function PIVATRANS%(NTRANI%)
    '
    'SE VALIDA QUE EXISTA EL REGISTRO, SI EXISTE LO LEE Y SALE POR QUE LA POSCICION DE IVA PUEDE SER 0
    If CantRegistros("TRANSPORTES", "CODI_TRA = " & NTRANI%, "NOMESS") > 0 Then
      XX% = XVALO(VALOBADA("TRANSPORTES", "PIVA_TRA", "CODI_TRA = " & NTRANI%, "NOMESS"))
      PIVATRANS% = XX%
    Else
       TX$ = FILTERGETRECORD$("TRANSPOR", 1, MKI$(NTRANI%))
       PIVATRANS% = Asc(Mid$(TX$, 195, 1))
    End If
    '
End Function

Function FAXITRANS$(NTRANI%)
    XX$ = VALOBADA("TRANSPORTES", "FAXI_TRA", "CODI_TRA = " & NTRANI%, "NOMESS")
    FAXITRANS$ = TRIM$(XX$)
    
    If FAXITRANS$ = "" Then
       XX$ = FILTERGETRECORD$("TRANSPOR", 1, MKI$(NTRANI%))
       FAXITRANS$ = TRIM$(Mid$(XX$, 243, 48))
    End If
End Function

Function MAILTRANS$(NTRANI%)
    XX$ = VALOBADA("TRANSPORTES", "MAIL_TRA", "CODI_TRA = " & NTRANI%, "NOMESS")
    MAILTRANS$ = TRIM$(XX$)
    
    If MAILTRANS$ = "" Then
       XX$ = FILTERGETRECORD$("TRANSPOR", 1, MKI$(NTRANI%))
       MAILTRANS$ = TRIM$(Mid$(XX$, 291, 48))
    End If
End Function

Function NTRANCLI$(NC)           ' NOMBRE DEL TRANSPORTISTA
    NTRANCLI$ = ""
    If NC > 0 Then
      NTR% = NUTRANCLI%(NC)
      If NTR% > 0 Then
        NTRANCLI$ = DENOTRANS$(NTR%)
      End If
    End If
End Function

Function RTRANCLI$(NC)           ' R.SOCIAL DEL TRANSPORTISTA
    RTRANCLI$ = ""
    If NC > 0 Then
      NTR% = NUTRANCLI%(NC)
      If NTR% > 0 Then
        RTRANCLI$ = RASOTRANS$(NTR%)
      End If
    End If
End Function

Function DTRANCLI$(NC)           ' DOMICILIO DEL TRANSPORTISTA
    DTRANCLI$ = ""
    If NC > 0 Then
      NTR% = NUTRANCLI%(NC)
      If NTR% > 0 Then
        DTRANCLI$ = DOMITRANS$(NTR%)
      End If
    End If
End Function

Function LTRANCLI$(NC)           ' LOCALIDAD DEL TRANSPORTISTA
    LTRANCLI$ = ""
    If NC > 0 Then
      NTR% = NUTRANCLI%(NC)
      If NTR% > 0 Then
        LTRANCLI$ = LOCATRANS$(NTR%)
      End If
    End If
End Function

'
Function HORAENT$(NC)
   On Error Resume Next
   If STATCLI%(NC) <> 0 Then HORAENT$ = TRIM$(ClienTemp!Hora_Ent)
End Function
Function DifireCuenta%(NumeClie, Credito$, FACTURA$)
   DifireCuenta% = 0
   CONDI$ = "NuClien = " & NumeClie & " AND CODICOM_LAR =  '" & Credito$ & "' AND TIPOMOVIM = 'RCOB'"
   CUENCREDITO% = XVALO(VALOBADA("CAJABANC", "CUECONTAI", CONDI$, "NOMESS"))

   CONDI$ = "NuClien = " & NumeClie & " AND CODICOM_LAR =  '" & FACTURA$ & "' AND TIPOMOVIM = 'FVEN'"
   CUENFACTURA% = XVALO(VALOBADA("CAJABANC", "CUECONTAI", CONDI$, "NOMESS"))
   '
   If CUENCREDITO% <> CUENFACTURA% Then DifireCuenta% = 1
End Function

'
Function ZONACLI$(NC)
   On Error Resume Next
   If STATCLI%(NC) <> 0 Then ZONACLI$ = TRIM$(ClienTemp!Zona_Cli)
End Function
'
Function OBSENTRE$(NC)
   On Error Resume Next
   If STATCLI%(NC) <> 0 Then OBSENTRE$ = ClienTemp!Obser_Entre
End Function
'
Function POREMBAL(NC)
   On Error Resume Next
   If STATCLI%(NC) <> 0 Then POREMBAL = TRIM$(ClienTemp!Recar_Embal)
End Function
'
Function CHEPEACRE#(NC)
   '
   Call APERBASDAT("CABAN")
   '
   SUCHEPE# = 0
   FEXI% = DIANTE(DIANTE(HOY(HOS$)))
   AYER = FETEXCOR1$(FEXI%)
15 Call APERBASDAT("CABAN")
   SQLX$ = "SELECT FECHENTRE,ValAbsComp FROM CAJABANC WITH(NOLOCK)"
   SQLX$ = SQLX$ + "where TipoMovim = 'CHRECI' "
   SQLX$ = SQLX$ + "AND NuClien = " & NC & " "
   SQLX$ = SQLX$ + "AND Status > 0 "
   SQLX$ = SQLX$ + "AND FechAcred >= '" & AYER & "' "
   SQLX$ = SQLX$ + " AND OpciMovim = 'CH' "
   '
   Dim CABATEMP As ADODB.Recordset
   Set CABATEMP = New ADODB.Recordset
   Set CABATEMP = FLEXCONN.Execute(FILTSQL$(SQLX$))
   '
   With CABATEMP
     On Error Resume Next
     .MoveFirst
     If Err < 1 Then
       On Error GoTo 0
       Do While Not .EOF
'         If TRIM$(!OpciMovim) <> "CH" Then GoTo 19 ' no es un cheque
         If IsNull(!FECHENTRE) = False Then
           If CVINT(!FECHENTRE) > 33 Then
             GoTo 19    ' salio de cartera
           End If
         End If
         '
         SUCHEPE# = SUCHEPE# + XVALO(!ValAbsComp)
         '
19     .MoveNext
       Loop
     End If
   End With
   '
   CHEPEACRE# = SUCHEPE#
   '
   CABATEMP.Close
   Set CABATEMP = Nothing
End Function
'
Function CUECOCLI%(NC)
   CUECOCLI% = 0
   On Error Resume Next
   If STATCLI%(NC) <> 0 Then CUECOCLI% = ClienTemp!CuMa_Cli
   '
   If CUECOCLI% < 1 Then
     CUMA$ = Control$("AMACLI", "ECUECON")
     If NC < 0 Then
       CUMA$ = Control$("AMAPRO", "ECUECON")
     End If
     CUECOCLI% = CUEINT%(CUMA$)
   End If
End Function
'
Function CUEMADRE$(NC)
   '
   CUEMADRE$ = ""
   On Error Resume Next
   If STATCLI%(NC) <> 0 Then CUEMADRE$ = CODCUE$(ClienTemp!CuMa_Cli)
   '
   If TRIM$(CUEMADRE$) = "" Then
     CUEMADRE$ = Control$("AMACLI", "ECUECON")
     If NC < 0 Then
       CUEMADRE$ = Control$("AMAPRO", "ECUECON")
     End If
   End If
   '
End Function
'
Function GRUCOCLI%(NC)
   GRUCOCLI% = 0
   'On Error Resume Next
   'If STATCLI%(NC%) <> 0 Then GRUCOCLI% = ClienTemp!GrCo_Cli
   If NC <> 0 Then
     Dim CLITE As ADODB.Recordset
     Set CLITE = New ADODB.Recordset
     '
     If NC > 0 Then
          Call APERBASDAT("CLIEN")
          Set CLITE = FLEXCONN.Execute(FILTSQL$("SELECT GrCo_Cli FROM Deudor12 WITH(NOLOCK) WHERE Nume_Cli = " & Abs(NC)))
       Else
          Call APERBASDAT("PROVE")
          Set CLITE = FLEXCONN.Execute(FILTSQL$("SELECT GrCo_Cli FROM Proved12 WITH(NOLOCK) WHERE Nume_Cli = " & Abs(NC)))
     End If
     '
     If Not (CLITE.EOF And CLITE.BOF) Then 'si la encuentra
        GRUCOCLI% = XVALO(CLITE!GrCo_Cli)
     End If
   End If
   '
End Function
'
Function GRUDIVCLI(NC)
   GRUDIVCLI = NC
   On Error Resume Next
   If STATCLI%(NC) <> 0 Then
     GRDVC = ClienTemp!Gru_Div
     If GRDVC > 0 Then
       GRUDIVCLI = GRDVC
     End If
   End If
End Function
'
Function ORIGRUDI%(NC)
   Static CTXYX1%, ORGD%(32767)
   If CTXYX1% < 1 Then
      For KKU& = 1 To 32767
         ORGD%(KKU&) = 0
      Next KKU&
      Call APERBASDAT("CLIEN")
      'jandy sql
      'Set BCLIEN = Badaclie.OpenRecordset("SELECT * FROM Deudor12 WHERE Gru_Div > 0", dbOpenSnapshot)
      'On Error Resume Next
      strSQL = "SELECT * FROM Deudor12 WHERE Gru_Div > 0"
      Dim BCLIEN As ADODB.Recordset
      Set BCLIEN = FLEXCONN.Execute(FILTSQL$(strSQL))
      
      
      With BCLIEN
        On Error Resume Next
        .MoveFirst
        If Err < 1 Then
           On Error GoTo 0
           Do While Not .EOF
             GRD% = XVALO(!Gru_Div)
             ORGD%(GRD%) = XVALO(!nume_cli)
           .MoveNext
           Loop
        End If
        On Error GoTo 0
      End With
      BCLIEN.Close
      Set BCLIEN = Nothing
      
      CTXYX1% = 1
   End If
   '
   ORIGRUDI% = ORGD%(NC)
   '
End Function
'
Function ORDECHEQ$(NC)
   ORDECHEQ$ = rasocli$(NC)
   If NC < 0 Then
     On Error Resume Next
     If STATCLI%(NC) <> 0 Then
       If IsNull(ClienTemp!Ord_Cheq) = False Then
         ORDECHEQ$ = ClienTemp!Ord_Cheq
       End If
     End If
   End If
End Function
'
Function ANOTACLI$(NC)
   ANOTACLI$ = ""
   On Error Resume Next
   If STATCLI%(NC) <> 0 Then ANOTACLI$ = ClienTemp!nota_Cli
End Function
'
Sub NotaCli(NC, MODO%)
   '
   ' MODO% = 1    abre automaticamente si hay nota para el cliente
   ' MODO% = 2    se pidió apertura por botón de acceso
   '
   If rasocli$(NC) = "" Then Exit Sub
   '
   If NC > 0 Then
     If REGICLI(NC) > 0 Then
       Call APERBASDAT("CLIEN")
       'Set ClienTemp = Badaclie.OpenRecordset("SELECT * FROM Deudor12 WHERE Nume_Cli = " & NCLIE, dbOpenSnapshot)
       '
       Set ClienTemp = FLEXCONN.Execute(FILTSQL$("SELECT * FROM Deudor12 WHERE Nume_Cli = " & NC))
       On Error Resume Next '
       ClienTemp.MoveFirst
       If Err < 1 Then
         TXTANOTADOR$ = SAFETEXT$(ClienTemp!nota_Cli)
       End If
     End If
   End If
   If TRIM$(TXTANOTADOR$) <> "" Then
        Call FRATEXTO(TXTANOTADOR$, 78)
        For KKU1% = 1 To CARETEX%
            X$ = REGBLAN$("NOTACLI")
            Call REPLA(X$, RETEX$(KKU1%), 3, 78)
            JJ& = JJ& + 1
            Call GRAREG("!NOTACLI", X$, JJ&)
        Next KKU1%
        VTB% = VENTABU%("NOTACLI")
   End If
   
'   RNC$ = RECFILT$("NOTACLI", 1, MKI$(Int(NC)))
'   If Len(RNC$) > 0 Or MODO% = 2 Then
'      Screen.MousePointer = 11
'      JJ& = 0
'      For U& = 1 To Len(RNC$) Step 4
'         RENOTA& = CVS(Mid$(RNC$, U&, 4))
'         If RENOTA& > 0 Then
'            If RENOTA& <= ULTREG&("NOTACLI") Then
'               X$ = REGLEIDO$("NOTACLI", RENOTA&)
'               JJ& = JJ& + 1
'               Call GRAREG("!NOTACLI", X$, JJ&)
'            End If
'         End If
'      Next U&
'      Call SETULTREG("!NOTACLI", JJ&)
'      Call CIERRARCH("!NOTACLI")
'      Screen.MousePointer = 1
'      VTB% = VENTABU%("NOTACLI")
'      If VTB% >= 0 Then
'         Screen.MousePointer = 11
'         URN& = ULTREG&("!NOTACLI")
'         JJ& = ULTREG&("NOTACLI")
'         While Len(RNC$) < 4 * URN&
'            JJ& = JJ& + 1
'            RNC$ = RNC$ + MKS$(JJ&)
'         Wend
'         KK& = (-3)
'         For JJ& = 1 To ULTREG&("!NOTACLI")
'            X$ = REGLEIDO$("!NOTACLI", JJ&)
'            Call REPLA(X$, MKI$(Int(NC)), 1, 2)
'            KK& = KK& + 4
'            RENOTA& = CVS(Mid$(RNC$, KK&, 4))
'            Call GRAREG("NOTACLI", X$, RENOTA&)
'         Next JJ&
'         While KK& < Len(RNC$) - 3
'            KK& = KK& + 4
'            RENOTA& = CVS(Mid$(RNC$, KK&, 4))
'            Call GRAREG("NOTACLI", MKI$(Int(NC)) + Space$(78), RENOTA&)
'         Wend
'         Call CIERRARCH("NOTACLI")
'         Screen.MousePointer = 1
'      End If
'   End If
End Sub
'
Function DEPROVI$(PR$)
   DEPROVI$ = ECOARCH$("PROVINC", PR$)
End Function
'
Function DELIPRE$(cl%)
   DX$ = "PRECIOS LIBRES"
   If cl% > 0 Then If cl% < 256 Then DX$ = ECOARCH$("LISTAS", Chr$(cl%))
   If DX$ = "" Then Beep: cl% = 0: DX$ = "PRECIOS LIBRES"
   DELIPRE$ = DX$
End Function
'
Function DECONPA$(cl%)
   DX$ = "CONTADO"
   If cl% > 0 Then If cl% < 256 Then DX$ = ECOARCH$("CONPAG", Chr$(cl%))
   If DX$ = "" Then Beep: cl% = 0: DX$ = "CONTADO"
   DECONPA$ = DX$
End Function
'
Function DEVENDE$(cl)
   DX$ = "GERENCIA"
   If cl > 0 Then DX$ = ECOARCH$("VENDEDOR", Chr$(VENDCLI%(cl)))
   If DX$ = "" Then Beep: cl = 0: DX$ = "GERENCIA"
   DEVENDE$ = DX$
End Function
'
Function DEGRUCLI$(cl%)
   DX$ = "NO CLASIFICADO"
   If cl% > 0 Then DX$ = ECOARCH$("GRUCOCLI", MKI$(cl%))
   If DX$ = "" Then Beep: cl% = 0: DX$ = "NO CLASIFICADO"
   DEGRUCLI$ = DX$
End Function
'
Function POSIVAC$(NC)
    '
    CAPOS% = 7
    Dim PIVA$(7)
    '
    PIVA$(0) = "CONS.FIN."
    PIVA$(1) = "RESP.INSC"
    PIVA$(2) = "R.NO INSC"
    PIVA$(3) = "NO RESP."
    PIVA$(4) = "EXENTO"
    PIVA$(5) = "COM.EXT. "
    PIVA$(6) = "MONOTR."
    PIVA$(7) = "EXEN-PRO"
    '
    APV% = PIVACLI%(NC)
    PS$ = ""
    If APV% >= 0 Then
        If APV% <= CAPOS% Then
            PS$ = PIVA$(APV%)
        End If
    End If
    '
    POSIVAC$ = PS$
End Function
'
Function POSIVAL$(NC)
    '
    CAPOS% = 7
    Dim PIVA$(7)
    '
    PIVA$(0) = "A CONSUMIDOR FINAL"
    PIVA$(1) = "IVA RESPONSABLE INSCRIPTO"
    PIVA$(2) = "IVA RESPONSABLE NO INSCRIPTO"
    PIVA$(3) = "IVA NO RESPONSABLE"
    PIVA$(4) = "IVA EXENTO"
    PIVA$(5) = "OPs.COMERCIO EXTERIOR"
    PIVA$(6) = "CONTRIBUYENTE MONOTRIBUTO"
    PIVA$(7) = "EXENTO - ZONA PROMOVIDA"
    '
    If InStr(EMPREAC$, "SABO") > 0 Then
      PIVA$(5) = "OPERACIONES EXPORTACION"
    End If
    '
    APV% = PIVACLI%(NC)
    PS$ = ""
    If APV% >= 0 Then
        If APV% <= CAPOS% Then
            PS$ = PIVA$(APV%)
        End If
    End If
    '
POSIVAL$ = PS$
End Function
'
Function VALICUIT%(Cuit$)
    '
    CU$ = TRIM$(Cuit$)
    VC% = 0
    If Control$("FORMCUIT", "SUPRIVAL") = "SI" Then GoTo 900
    '
    If Len(CU$) < 11 Then GoTo 999                                    ' no valida
    '
    VX% = Asc(Mid$(CU$, 3))
    If VX% >= 48 Then
        If VX% <= 57 Then
            CU$ = Left$(CU$, 2) + "-" + Mid$(CU$, 3)
        End If
    End If
    '
    If Len(CU$) < 12 Then GoTo 999
    VX% = Asc(Mid$(CU$, 12))
    If VX% >= 48 Then
        If VX% <= 57 Then
            CU$ = Left$(CU$, 11) + "-" + Mid$(CU$, 12)
        End If
    End If
    '
    If Len(CU$) <> 13 Then GoTo 999                                   ' no valida
    '
    For J% = 1 To 13
        If J% <> 3 Then
            If J% <> 12 Then
                VX% = Asc(Mid$(CU$, J%, 1))
                If VX% < 48 Or VX% > 57 Then GoTo 999                 ' no valida
            End If
        End If
    Next J%
    '
    Mid$(CU$, 3, 1) = "-"
    Mid$(CU$, 12, 1) = "-"
    '
    SUMA% = 0
    SUMA% = SUMA% + 5 * XVALO(Mid$(CU$, 1, 1))
    SUMA% = SUMA% + 4 * XVALO(Mid$(CU$, 2, 1))
    SUMA% = SUMA% + 3 * XVALO(Mid$(CU$, 4, 1))
    SUMA% = SUMA% + 2 * XVALO(Mid$(CU$, 5, 1))
    SUMA% = SUMA% + 7 * XVALO(Mid$(CU$, 6, 1))
    SUMA% = SUMA% + 6 * XVALO(Mid$(CU$, 7, 1))
    SUMA% = SUMA% + 5 * XVALO(Mid$(CU$, 8, 1))
    SUMA% = SUMA% + 4 * XVALO(Mid$(CU$, 9, 1))
    SUMA% = SUMA% + 3 * XVALO(Mid$(CU$, 10, 1))
    SUMA% = SUMA% + 2 * XVALO(Mid$(CU$, 11, 1))
    '
    SUMB% = SUMA% Mod 11
    DIVER% = 0
    If SUMB% Mod 11 <> 0 Then DIVER% = 11 - SUMB%
    If DIVER% <> XVALO(Mid$(CU$, 13, 1)) Then GoTo 999                  ' no valida
    '
900 VC% = 1
    Cuit$ = CU$

999 VALICUIT% = VC%
    '
End Function

Function FormatearCuitValido$(CUITX$)
    'RECIBE UN CUIT VALIDO POR QUE ASI NOS ASEGURAMOS QUE TIENE TODOS LOS CARACTERERES
    'SI TIENE LOS SEPARADORES LO DEJA IGUAL
    'SI LE FALTA ALGUNO DE LOS 2 SE LOS AGREGA
    'EJEMPLO: RECIBE 20140119548 DEVUELVE 20-14011954-8
    'RECIBE 20-140119548 DEVUELVE 20-14011954-8
    'RECIBE 2014011954-8 DEVUELVE 20-14011954-8
    
    If InStr(Mid$(CUITX$, 3, 1), "-") < 1 Then
       CUIT1$ = CUITX$
       CUIT1$ = Left$(CUIT1$, 2) & "-"
       CUITX$ = CUIT1$ & Mid$(CUITX$, 3)
    Else
       CUIT1$ = Left$(CUITX$, 3)
    End If
    If InStr(Mid$(CUITX$, 12, 1), "-") < 1 Then
       Cuit2$ = CUITX$
       Cuit2$ = Mid$(Cuit2$, 4, 8) & "-" & Right(CUITX$, 1)
    Else
       Cuit2$ = Right$(CUITX$, 10)
    End If
   FormatearCuitValido$ = CUIT1$ & Cuit2$
End Function
               
Sub CARDATPRO(NC%)
    '
    Screen.MousePointer = 11
    '
    MODIPRO.Text2.TEXT = TEL2CLI$(NC%)
    MODIPRO.Text3.TEXT = rasocli$(NC%)
    MODIPRO.Text4.TEXT = DOMICLI$(NC%)
    MODIPRO.Text5.TEXT = CPOSCLI$(NC%)
    MODIPRO.Text6.TEXT = LOCACLI$(NC%)
    MODIPRO.Text10.TEXT = TEL1CLI$(NC%)
    MODIPRO.Text15.TEXT = CuitCli$(NC%)
    MODIPRO.Text18.TEXT = NIBRCLI$(NC%)
    MODIPRO.Text9.TEXT = PAISCLI$(NC%)
    MODIPRO.Text11.TEXT = FAXCLI$(NC%)
    MODIPRO.Text12.TEXT = EMAILCLI$(NC%)
    '
    MODIPRO.COTEXT(0).TEXT = TRIM$(PROVCLI$(NC%))
    MODIPRO.COTEXT(1).TEXT = TRIM$(Str$(PIVACLI%(NC%)))
    MODIPRO.COTEXT(2).TEXT = TRIM$(CATIBRU$(NC%))
    MODIPRO.COTEXT(3).TEXT = TRIM$(Str$(LIPRCLI%(NC%)))
    MODIPRO.COTEXT(4).TEXT = TRIM$(Str$(CPAGCLI%(NC%)))
    MODIPRO.COTEXT(5).TEXT = TRIM$(CUEMADRE$(NC%))
    '
    For II = 0 To 5
       ECO$ = ""
       If TRIM$(MODIPRO.COTEXT(II).TEXT) <> "" Then
          TV11% = TIPFIELD(ARCHE$(II), 1)
          LF11% = LENFIELD(ARCHE$(II), 1)
          ARGU$ = CBIN$(TRIM$(MODIPRO.COTEXT(II).TEXT), TV11%, LF11%)
          ECO$ = TRIM$(ECOARCH$(ARCHE$(II), ARGU$))
       End If
       '
       If ECO$ <> "" Then
             MODIPRO.DETEXT(II).TEXT = ECO$
          Else
             MODIPRO.DETEXT(II).TEXT = "No Definido"
       End If
    Next II
    '
    Screen.MousePointer = 1
    '
End Sub
'

Sub CARTAVALI()
    '
    If ULTREG&("CARETIB") < 4 Then
      IQ& = 0
      IQ& = IQ& + 1: Call GRAREG("CARETIB", "0I.Brutos No Responsable", IQ&)
      IQ& = IQ& + 1: Call GRAREG("CARETIB", "1Contribuyente Puro", IQ&)
      IQ& = IQ& + 1: Call GRAREG("CARETIB", "2Convenio Multilateral", IQ&)
      IQ& = IQ& + 1: Call GRAREG("CARETIB", "3Conv.Multil. (Comerciales)", IQ&)
      Call SETULTREG("CARETIB", IQ&)
      Call CIERRARCH("CARETIB")
    End If
    '
    If ULTREG&("TICUEPRO") < 2 Then
      IQ& = 0
      IQ& = IQ& + 1: Call GRAREG("TICUEPRO", Chr$(0) + "Cuenta Corriente", IQ&)
      IQ& = IQ& + 1: Call GRAREG("TICUEPRO", Chr$(1) + "Cuenta de Gastos", IQ&)
      Call SETULTREG("TICUEPRO", IQ&)
      Call CIERRARCH("TICUEPRO")
    End If
    '
    If ULTREG&("CAPOSIV") < 8 Then
      IQ& = 0
      IQ& = IQ& + 1: Call GRAREG("CAPOSIV", Chr$(0) + AJUSTI$("IVA Consumidor Final            ", 31) + AJUSTI$("Cons.Fin", 8) + String$(24, 0), IQ&)
      IQ& = IQ& + 1: Call GRAREG("CAPOSIV", Chr$(1) + AJUSTI$("IVA Responsable Inscripto   ", 31) + AJUSTI$("Resp.Ins", 8) + String$(24, 0), IQ&)
      IQ& = IQ& + 1: Call GRAREG("CAPOSIV", Chr$(2) + AJUSTI$("IVA Responsable no Inscripto", 31) + AJUSTI$("R.No Ins", 8) + String$(24, 0), IQ&)
      IQ& = IQ& + 1: Call GRAREG("CAPOSIV", Chr$(3) + AJUSTI$("IVA No Responsable          ", 31) + AJUSTI$("No Resp.", 8) + String$(24, 0), IQ&)
      IQ& = IQ& + 1: Call GRAREG("CAPOSIV", Chr$(4) + AJUSTI$("IVA Responsable Exento      ", 31) + AJUSTI$("Exento  ", 8) + String$(24, 0), IQ&)
      IQ& = IQ& + 1: Call GRAREG("CAPOSIV", Chr$(5) + AJUSTI$("IVA Operac.Comercio Exterior", 31) + AJUSTI$("Com.Ext.", 8) + String$(24, 0), IQ&)
          If InStr(EMPREAC$, "SABO") > 0 Then Call GRAREG("CAPOSIV", Chr$(5) + AJUSTI$("IVA Operaciones Exportacion", 31) + AJUSTI$("Export.", 8) + String$(24, 0), IQ&)
      IQ& = IQ& + 1: Call GRAREG("CAPOSIV", Chr$(6) + AJUSTI$("IVA Contrib.Monotributo     ", 31) + AJUSTI$("Monotrib", 8) + String$(24, 0), IQ&)
      IQ& = IQ& + 1: Call GRAREG("CAPOSIV", Chr$(7) + AJUSTI$("IVA Exento Zona Promovida   ", 31) + AJUSTI$("Monotrib", 8) + String$(24, 0), IQ&)
      Call SETULTREG("CAPOSIV", IQ&)
      Call CIERRARCH("CAPOSIV")
    End If
    '
    For i& = ULTREG&("LISTAS") To 1 Step -1
       XX$ = REGLEIDO$("LISTAS", i&)
       If Asc(Left$(XX$, 1)) > 0 Then GoTo 59
       Call GRAREG("LISTAS", String$(128, 0), i&)
    Next i&
59  Call CIERRARCH("LISTAS")
    '
     For i& = ULTREG&("CONPAG") To 1 Step -1
       XX$ = REGLEIDO$("CONPAG", i&)
       If Asc(Left$(XX$, 1)) > 0 Then GoTo 69
       Call GRAREG("CONPAG", String$(128, 0), i&)
    Next i&
69  Call CIERRARCH("CONPAG")
    '
    For i& = ULTREG&("VENDEDOR") To 1 Step -1
       XX$ = REGLEIDO$("VENDEDOR", i&)
       If Asc(Left$(XX$, 1)) > 0 Then GoTo 79
       Call GRAREG("VENDEDOR", String$(128, 0), i&)
    Next i&
79  Call CIERRARCH("VENDEDOR")
    '
     For i& = ULTREG&("PROVINC") To 1 Step -1
       XX$ = REGLEIDO$("PROVINC", i&)
       If Asc(Left$(XX$, 1)) > 32 Then GoTo 89
       Call GRAREG("PROVINC", String$(128, 0), i&)
    Next i&
89  Call CIERRARCH("PROVINC")
    '
    ARCHE$(0) = "PROVINC"
    ARCHE$(1) = "CAPOSIV"
    ARCHE$(2) = "CARETIB"
    ARCHE$(3) = "LISTAS"
    ARCHE$(4) = "CONPAG"
    ARCHE$(5) = "VENDEDOR"
    ARCHE$(6) = "ECUECON"
    ARCHE$(7) = "GRUCOCLI"
    '
End Sub

Sub SORTCLI(ARDEU%)
    '
    Dim DEU1 As String * 32
    Dim DEU2 As String * 128
    Dim REGDEU$(32767)
    Dim OCUCLI%(32767)
    '
    PCLI$ = String$(32767, 0)
    '
    LU$ = LUDAT$
    ARX1$ = LU$ + "DEUDOR1.DAT"
    ASELE$ = "DEUDOR1"
    If ARDEU% = 2 Then
        ARX1$ = LU$ + "PROVED1.DAT"
        ASELE$ = "PROVED1"
    End If
    On Error GoTo 100
    n1% = FreeFile
    Open ARX1$ For Random Access Read Write Lock Read Write As #n1% Len = 32
    '
    ARX2$ = LU$ + "DEUDOR2.DAT"
    If ARDEU% = 2 Then
        ARX2$ = LU$ + "PROVED2.DAT"
    End If
    On Error GoTo 200
    N2% = FreeFile
    Open ARX2$ For Random Access Read Write Lock Read Write As #N2% Len = 128
    '
    Screen.MousePointer = 11
    AGRECLI% = 0
    On Error GoTo 0
    Lock #n1%
    Lock #N2%
    '
    FORSELEC.LISORT.Clear
    For U& = 1 To LOF(n1%) / 32
        Get #n1%, U&, DEU1$
        Get #N2%, U&, DEU2$
        If CVI(Left$(DEU1$, 2)) > 0 Then
            '
            NCXI% = CVI(Left$(DEU1$, 2))
            If NCXI% <= 32767 Then
              If OCUCLI%(NCXI%) > 0 Then
                GoTo 19       ' descarta repetidos
              End If
              OCUCLI%(NCXI%) = 1
            End If
            '
            CLI$ = Mid$(DEU1$, 3, 30)
            Cuit$ = TRIM$(Mid$(DEU2$, 75, 15))
            If VERICUIT% <> (-1) Then
              If Cuit$ <> "" Then
                If VALICUIT%(Cuit$) <> 1 Then
                  Screen.MousePointer = 1
                  TECLI$ = "Cliente": If ARDEU% = 2 Then TECLI$ = "Proveedor"
                  If COMALTER%(TECLI$ + ": " + TRIM$(CLI$) + "\C.U.I.T.: " + Cuit$ + "\  \Número de C.U.I.T. no Válido.\\Continuar Revisión\Cancelar") = 2 Then
                     VERICUIT% = (-1)
                  End If
                  Screen.MousePointer = 11
                End If
              End If
            End If
            REX$ = Space$(36)
            Call REPLA(REX$, CLI$, 1, 30)
            Call REPLA(REX$, CSTRING$(MKS$(U&), 3, 6, 0, 2), 31, 6)
            FORSELEC.LISORT.AddItem REX$
            REGDEU$(U&) = DEU1$ + DEU2$
            '
          ElseIf CVI(Left$(DEU1$, 2)) < 0 Then
            '
            XX$ = DEU1$
            Call COPYSTRU("DEUDOR1", "DEUDOBAJ")
            Call REGAPP("DEUDOBAJ", XX$)
            Call CIERRARCH("DEUDOBAJ")
        End If
19  Next U&
    '
    URE& = 0
    For U& = 1 To FORSELEC.LISORT.ListCount
        FORSELEC.LISORT.ListIndex = U& - 1
        INDI& = XVALO(Mid$(FORSELEC.LISORT.TEXT, 31, 6))
        DEU1$ = Left$(REGDEU$(INDI&), 32)
        DEU2$ = Mid$(REGDEU$(INDI&), 33, 128)
        Put #n1%, U&, DEU1$
        Put #N2%, U&, DEU2$
        URE& = U&
    Next U&
    For U& = URE& + 1 To LOF(n1%) / 32
        DEU1$ = String$(32, 0)
        DEU2$ = String$(128, 0)
        Put #n1%, U&, DEU1$
        Put #N2%, U&, DEU2$
    Next U&
    '
    Unlock #N2%
    Unlock #n1%
    Close #N2%: N2% = 0
    Close #n1%: n1% = 0
    DEU1$ = REGLEIDO$(ASELE$, 1)
    Call GRAREG(ASELE$, DEU1$, 1)
    Call CIERRARCH(ASELE$)
    Screen.MousePointer = 1
    Exit Sub
    '
100 Resume 101
101 On Error GoTo 0
    Close #n1%: n1% = 0
    Call COMUNI("Imposible Ordenar - Archivo en Uso.")
    Exit Sub
    '
200 Resume 201
201 On Error GoTo 0
    Close #n1%: n1% = 0
    Close #N2%: N2% = 0
    Call COMUNI("Imposible Ordenar - Archivo en Uso.")
    Exit Sub
    '
End Sub
'

Function CUEIBRUPRO$(A$)
   '
   Static CTX%, COPROVI$, CUEPROVI$(64)
   If CTX% = 0 Then
     For i& = 1 To ULTREG&("PROVINB")
       X$ = REGLEIDO$("PROVINB", i&)
       If ECOARCH$("CUECON", Mid$(X$, 27, 2)) = "" Then
          Call REPLA(X$, MKI$(0), 27, 2)
       End If
       If TRIM$(Left$(X$, 1)) <> "" Then
          If Len(COPROVI$) < 64 Then
            COPROVI$ = COPROVI$ + Left$(X$, 1)
            CUEPROVI$(Len(COPROVI$)) = Mid$(ECOARCH$("CUECON", Mid$(X$, 27, 2)), 31, 12)
          End If
       End If
     Next i&
     Call CIERRARCH("PROVINB")
     CTX% = 1
   End If
   '
   A1$ = AJUSTI$(A$, 1)
   QQQ$ = Space$(12)
   PPXX% = InStr(COPROVI$, A1$)
   If PPXX% > 0 Then
     QQQ$ = CUEPROVI$(PPXX%)
   End If
   '
   CUEIBRUPRO$ = AJUSTI$(QQQ$, 12)
   '
End Function

Function CUEIPROVIN%(A$)
   '
   Static CTX%, COPROVI$, CUEPROVII%(64)
   If CTX% = 0 Then
     For i& = 1 To ULTREG&("PROVINB")
       X$ = REGLEIDO$("PROVINB", i&)
       If ECOARCH$("CUECON", Mid$(X$, 27, 2)) = "" Then
          Call REPLA(X$, MKI$(0), 27, 2)
       End If
       If TRIM$(Left$(X$, 1)) <> "" Then
          If Len(COPROVI$) < 64 Then
            COPROVI$ = COPROVI$ + Left$(X$, 1)
            CUEPROVII%(Len(COPROVI$)) = CVI(Mid$(X$, 27, 2))
          End If
       End If
     Next i&
     Call CIERRARCH("PROVINB")
     CTX% = 1
   End If
   '
   A1$ = AJUSTI$(A$, 1)
   QQI% = 0
   PPXX% = InStr(COPROVI$, A1$)
   If PPXX% > 0 Then
     QQI% = CUEPROVII%(PPXX%)
   End If
   '
   CUEIPROVIN% = QQI%
   '
End Function
Function CodigoRegimenRetIBXProvincia(LETRA_PROVIN$, Optional RET_CTACBLE_IB)
    XX$ = FILTERGETRECORD$("PROVINB", 1, LETRA_PROVIN$)
    RET_CTACBLE_IB = CVI(Mid$(XX$, 27, 2))
    CodigoRegimenRetIBXProvincia = CVS(Mid$(XX$, 37, 4))
End Function
Function CUEMECOB%(TIPOVAL%)
   '
   Static CMC0%
   '
   If CMC0% < 1 Then
     CMC0% = CUECOINT%(Control$("", "CUECHECA"))
     If CMC0% < 1 Then
       Call MENSERR(24, "Error de Configuración.\  \Falta Definir Cuenta Contable\Asociada a la Cartera de Valores.")
       On Error Resume Next
       OPCHEQ07.Show 1
       On Error GoTo 0
       Call CIERRARCH("*.*")
       Call FINAL("")
     End If
   End If
   '
   CMC% = CMC0%
   RFF$ = RECFILT$("MECOBRA", 1, MKI$(TIPOVAL%))
   If Len(RFF$) >= 4 Then
     REME& = CVS(Mid$(RFF$, 1, 4))
     If REME& > 0 Then
       CUCU$ = Mid$(REGLEIDO$("MECOBRA", REME&), 45, 12)
       CMC% = CUECOINT%(CUCU$)
       If CMC% < 1 Then CMC% = CMC0%
     End If
   End If
   '
   CUEMECOB% = CMC%
   '
End Function

Function CUEMEPAG%(TIPOVAL%)
   '
   Static CMC0%
   '
   If CMC0% < 1 Then
     CMC0% = CUECOINT%(Control$("", "CUECHECA"))
     If CMC0% < 1 Then
       Call MENSERR(24, "Error de Configuración.\  \Falta Definir Cuenta Contable\Asociada a la Cartera de Valores.")
       Call CIERRARCH("*.*")
       Call FINAL("")
     End If
   End If
   '
   CMC% = CMC0%
   RFF$ = RECFILT$("MEPAGO", 1, MKI$(TIPOVAL%))
   If Len(RFF$) >= 4 Then
     REME& = CVS(Mid$(RFF$, 1, 4))
     If REME& > 0 Then
       XX$ = REGLEIDO$("MEPAGO", REME&)
       CMC% = CVI(Mid$(XX$, 61, 2))
       If CMC% < 1 Then
         CUCU$ = Mid$(XX$, 45, 12)
         CMC% = CUECOINT%(CUCU$)
       End If
       If CMC% < 1 Then CMC% = CMC0%
     End If
   End If
   '
   CUEMEPAG% = CMC%
   '
End Function

Function CUECOINT%(CUEX$)
   '
   CUECOINT% = CUEINT%(CUEX$)
   '
End Function

Function CUECARCHE$()
   '
   CUECARCHE$ = Control$("", "CUECHECA")
   CUEICHEQ% = CUECOINT%(CUECARCHE$)
   If TRIM$(CUECARCHE$) = "" Or CUEICHEQ% < 1 Then
     CUECARCHE$ = ""
     Exit Function
   End If
   '
End Function

Function CUEICARCH%()
   '
   CUECHECA$ = Control$("", "CUECHECA")
   CUEICARCH% = CUECOINT%(CUECHECA$)
   If TRIM$(CUECHECA$) = "" Or CUEICARCH% < 1 Then
     CUEICARCH% = 0
   End If
   '
End Function

Function MONEMICO%(NPROX%)
   '
   MNC% = 1
   RFF$ = RECFILT$("CLAPROMO", 1, MKI$(Abs(NPROX%)))
   If Len(RFF$) >= 4 Then
     RECLA& = CVS(Left$(RFF$, 4))
     XX$ = REGLEIDO$("CLAPROMO", RECLA&)
     MNC% = Asc(Mid$(XX$, 31, 1))
   End If
   '
   MONEMICO% = MNC%
   '
End Function

Function MONEVACO%(NPROX%)
   '
   MNC% = 1
   RFF$ = RECFILT$("CLAPROMO", 1, MKI$(Abs(NPROX%)))
   If Len(RFF$) >= 4 Then
     RECLA& = CVS(Left$(RFF$, 4))
     XX$ = REGLEIDO$("CLAPROMO", RECLA&)
     MNC% = Asc(Mid$(XX$, 32, 1))
   End If
   '
   MONEVACO% = MNC%
   '
End Function

Sub SELRACLIEN(NCL1%, NCL2%, VENDSEL%)
   '
   NCMIN% = 32767: NCMAX% = 0
   IXXX$ = LOADFILE$(LUDAT$ + "DEUDOR1.X01")
   '
   For KKL& = 1 To Len(IXXX$) Step 2
     NCMX% = CVI(Mid$(IXXX$, KKL&, 2))
     If NCMX% > 0 Then
       If NCMX% < NCMIN% Then NCMIN% = NCMX%
       If NCMX% > NCMAX% Then NCMAX% = NCMX%
     End If
   Next KKL&
   '
10 VDEF$ = MKI$(NCMIN%) + MKI$(NCMAX%) + Chr$(0)
   OBX$ = OBJPLA$("DESHASCLIEVEN", VDEF$)
   If OBX$ = "" Then
     NCL1% = 0: NCL2% = 0
     Exit Sub
   End If
   '
   NCL1% = CVI(Left$(OBX$, 2))
   NCL2% = CVI(Mid$(OBX$, 3, 2))
   If NCL1% < 1 Or NCL2% < 1 Or NCL2% < NCL1% Then
     Call MENSERR(24, "Rango no Válido")
     GoTo 10
   End If
   VENDSEL% = Asc(Mid$(OBX$, 5, 1))
   '
End Sub

Sub SELRACLIENUE(NCL1, NCL2, VENDSEL%)
   '
   NCMIN = 999999: NCMAXI = 0
   IXXX$ = LOADFILE$(LUDAT$ + "DEUDOR12.X01")
   '
   For KKL& = 1 To Len(IXXX$) Step 4
     NCMX = CVS(Mid$(IXXX$, KKL&, 4))
     If NCMX > 0 Then
       If NCMX < NCMIN Then NCMIN = NCMX
       If NCMX > NCMAXI Then NCMAXI = NCMX
     End If
   Next KKL&
   '
10 VDEF$ = MKS$(NCMIN) + MKS$(NCMAXI) + Chr$(0)
   OBX$ = OBJPLA$("DESHASCLIEVENUE", VDEF$)
   If OBX$ = "" Then
     NCL1 = 0: NCL2 = 0
     Exit Sub
   End If
   '
   NCL1 = CVS(Left$(OBX$, 4))
   NCL2 = CVS(Mid$(OBX$, 5, 4))
   If NCL1 < 1 Or NCL2 < 1 Or NCL2 < NCL1 Then
     Call MENSERR(24, "Rango no Válido")
     GoTo 10
   End If
   VENDSEL% = Asc(Mid$(OBX$, 9, 1))
   '
End Sub
'
Function TELESUC$(NCLIEN, ISUC%)
 '\\\OT-07-0010-RG-12/01/07///
  TELESUC$ = SAFETEXT$(VALOBADA("SUCENTRE", "Tele_Suc", "NUME_CLI=" & NCLIEN & " and Nume_Suc=" & Str$(ISUC%)))
'  RFF$ = RECFILT$("SUCENTRE", 5, MKI$(NCLIEN))
'  For I1& = 1 To Len(RFF$) Step 4
'    I& = CVS(Mid$(RFF$, I1&))
'    X$ = REGLEIDO$("SUCENTRE", I&)
'    SUCU1% = CVI(Mid$(X$, 1, 2))
'    If SUCU1% = ISUC% Then
'      TELESUC$ = TRIM$(Mid$(X$, 163, 32))
'      Exit For
'    End If
'  Next I1&
  '\\\OT-07-0010-RG-FIN///
End Function
Function VALOPEDPEN1#(NCLIE)
   VALOPEDPEN1# = 0
   '
   Call APERBASDAT("PEDCLI")
   SQLX$ = " SELECT CANPED, CANTENT, PREUNID FROM PEDDETA "
   SQLX$ = SQLX$ + " INNER JOIN PEDENCA WITH(NOLOCK) ON PEDDETA.NroPed = PEDENCA.NroPed"
   SQLX$ = SQLX$ + " WHERE PEDENCA.STATUS < 9"
   SQLX$ = SQLX$ + " AND PEDENCA.NroClie = " & NCLIE & ""
   SQLX$ = SQLX$ + " AND PEDDETA.STATUS < 8 "
   SQLX$ = SQLX$ + " AND PEDDETA.CANPED > PEDDETA.CANTENT "
   SQLX$ = SQLX$ + " AND PEDENCA.CODIEMPR = " & CodiEmp% & " "
   SQLX$ = SQLX$ + " AND PEDdeta.CODIEMPR = " & CodiEmp% & " "
   '
   Dim PEDCLITEMP As ADODB.Recordset
   Set PEDCLITEMP = New ADODB.Recordset
   Set PEDCLITEMP = FLEXCONN.Execute(FILTSQL$(SQLX$))
   '
   With PEDCLITEMP
    Do While Not .EOF
        TotalItem# = (XVALO(!CANPED) - XVALO(!CANTENT)) * XVALO(!PreUnid)
        VALOPEDPEN1# = VALOPEDPEN1# + TotalItem#
        .MoveNext
    Loop
   End With
   PEDCLITEMP.Close
   Set PEDCLITEMP = Nothing
   '
'   VALOPEDPEN1# = 0
'   '
'   Call APERBASDAT("PEDCLI")
'   SQLX$ = " SELECT SUM((CANPED - CANTENT)* PREUNID) AS VAPEPEN FROM PEDDETA "
'   SQLX$ = SQLX$ + " INNER JOIN PEDENCA ON PEDDETA.NroPed = PEDENCA.NroPed"
'   SQLX$ = SQLX$ + " WHERE PEDENCA.STATUS < 9"
'   SQLX$ = SQLX$ + " AND PEDENCA.NroClie = " & NCLIE & ""
'   SQLX$ = SQLX$ + " AND PEDDETA.STATUS < 8 "
'   SQLX$ = SQLX$ + " AND PEDDETA.CANPED > PEDDETA.CANTENT "
'   SQLX$ = SQLX$ + " AND PEDENCA.CODIEMPR = " & CodiEmp% & " "
'   SQLX$ = SQLX$ + " AND PEDdeta.CODIEMPR = " & CodiEmp% & " "
'   '
'   Dim PEDCLITEMP As ADODB.Recordset
'   Set PEDCLITEMP = New ADODB.Recordset
'   Set PEDCLITEMP = FLEXCONN.Execute(FILTSQL$(SQLX$))
'   '
'   With PEDCLITEMP
'    If Not (.EOF And .BOF) Then
'      VALOPEDPEN1 = XVALO(!VAPEPEN)
'    End If
'   End With
'   PEDCLITEMP.Close
'   Set PEDCLITEMP = Nothing
'   '
End Function

Function ALIRETBRU#(NPROVE%, MINOIMP#, MEPAGO%)
    '
    Static CTCALCU%, ALIBRU(8), MINIMPO#(8), MEPAGX%(8)
    '
    ' Static PROVANT%, ALIRETBRUANT#, MINOIMPANT#, MEPAGOANT%
    ' Linea anterior comentada el 11/5/2009 - DOSIVAC: Si se generaban
    ' dos O/Pagos seguidas al MISMO PROVEEDOR, en el segundo traia la
    ' retencion de IIBB en cero. Al comentar la linea fuerza a repe-
    ' tir la búsqueda.
    '
    If NPROVE% > 0 And NPROVE% = PROVANT% Then GoTo 99
    '
    If Abs(NPROVE%) = 0 Then
        ALIRETBRUANT# = 0: MINOIMPANT# = 0: MEPAGOANT% = 0
        GoTo 99
    End If
    '
    If PIVACLI%(NPROVE% * (-1)) = 5 Then  ' Si el proveedor es importacion no se realiza calculo
        RETENC# = 0: MINOMPONIBLE# = 0: MEDIOPAGO% = 0
        GoTo 98
    End If
    '
    For U& = 1 To ULTREG&("TACATIBP")
       TCX$ = REGLEIDO$("TACATIBP", U&)
       TCY% = Asc(Mid$(TCX$, 1, 1))
       If TCY% > 0 Then
          If TCY% <= 8 Then
             ALIBRU(TCY%) = CVS(Mid$(TCX$, 33, 4))
             If ALIBRU(TCY%) > 0 Then CTCALCU% = 1
             MINIMPO#(TCY%) = CVS(Mid$(TCX$, 37, 4))
             MEPAGX%(TCY%) = CVI(Mid$(TCX$, 41, 2))
          End If
       End If
    Next U&
    '
    REBLA$ = REGBLAN$("DARIBPRO")
    RERET$ = FILTERGETRECORD$("DARIBPRO", 1, MKI$(Abs(NPROVE%)))
    CATIB1% = 1: PORCONVE = 100   ' VALORES POR DEFAULT
    If RERET$ <> REBLA$ Then
        CATIB1% = Asc(Mid$(RERET$, 3, 1))
        PORCONVE = CVS(Mid$(RERET$, 5, 4))
    End If
    
    RETENC# = ALIBRU(CATIB1%)
    MINOMPONIBLE# = MINIMPO#(CATIB1%)
    MEDIOPAGO% = MEPAGX%(CATIB1%)
    ' falta sacar el medio de pago
    '
    If EXISTE%(LUDAT$ + "PADIBRUR.TXT") > 0 Then
       CUITPROV$ = CuitCli$((-1) * Abs(NPROVE%))
       XX$ = REPADIBRU$(CUITPROV$, 2)
       If TRIM$(XX$) <> "" Then
          ret$ = Mid$(XX$, 48, 4)
          RETENC# = XVALO(ret$)
       End If
    End If
    '
98  PROVANT% = NPROVE%
    ALIRETBRUANT# = RETENC#
    MINOIMPANT# = MINOMPONIBLE#
    MEPAGOANT% = MEDIOPAGO%
    '
99  ALIRETBRU# = ALIRETBRUANT#
    MINOIMP# = MINOIMPANT#
    MEPAGO% = MEPAGOANT%
    '
End Function

Function ALIRETBRU_CABA#(NPROVE%, MINOIMP#, MEPAGO%)
    '
    Static CTCALCU%, ALIBRU(8), MINIMPO#(8), MEPAGX%(8)
    '
    If NPROVE% > 0 And NPROVE% = PROVANT% Then GoTo 99
    '
    If Abs(NPROVE%) = 0 Then
        ALIRETBRUANT# = 0: MINOIMPANT# = 0: MEPAGOANT% = 0
        GoTo 99
    End If
    '
    If PIVACLI%(NPROVE% * (-1)) = 5 Or PIVACLI%(NPROVE% * (-1)) = 0 Then  ' Si el proveedor es importacion no se realiza calculo
        RETENC# = 0: MINOMPONIBLE# = 0: MEDIOPAGO% = 0
        GoTo 98
    End If
    '
    For U& = 1 To ULTREG&("TACAIBCA")
       TCX$ = REGLEIDO$("TACAIBCA", U&)
       TCY% = Asc(Mid$(TCX$, 1, 1))
       If TCY% > 0 Then
          If TCY% <= 8 Then
             ALIBRU(TCY%) = CVS(Mid$(TCX$, 33, 4))
             If ALIBRU(TCY%) > 0 Then CTCALCU% = 1
             MINIMPO#(TCY%) = CVS(Mid$(TCX$, 37, 4))
             MEPAGX%(TCY%) = CVI(Mid$(TCX$, 41, 2))
          End If
       End If
    Next U&
    '
    
    ' *** CATEGORIA DE INGRESOS BRUTOS
    CATIB1% = XVALO(VALOBADA("PROVED12", "CATEGORIAIBCABA_CLI", "NUME_CLI=" & CStr(NPROVE%)))
    
    RETENC# = ALIBRU(CATIB1%)
    MINOMPONIBLE# = MINIMPO#(CATIB1%)
    MEDIOPAGO% = MEPAGX%(CATIB1%)
    ' falta sacar el medio de pago
    '
    If EXISTE%(LUDAT$ + "PAIBCABA.TXT") > 0 Or EXISTE%(LUDAT$ + "IIBBCABA.TXT") > 0 Then
       CUITPROV$ = CuitCli$((-1) * Abs(NPROVE%))
       XX$ = REPADIBRU_CABA$(CUITPROV$, 2)
       If TRIM$(XX$) <> "" Then
          ret$ = Mid$(XX$, 40, 4)
          RETENC# = XVALO(ret$)
       End If
    End If
    '
98  PROVANT% = NPROVE%
    ALIRETBRUANT# = RETENC#
    MINOIMPANT# = MINOMPONIBLE#
    MEPAGOANT% = MEDIOPAGO%
    '
99  ALIRETBRU_CABA# = ALIRETBRUANT#
    MINOIMP# = MINOIMPANT#
    MEPAGO% = MEPAGOANT%
    '
End Function

Function ALIRETBRU_TUCUMAN#(NPROVE%, MINOIMP#, MEPAGO%)
    '
    Static ALIRETBRUANT#, MINOIMPANT#, MEPAGOANT%, PROVANT%
    Dim ALIBRU#(8), MINIMPO#(8), MEPAGX%(8)
    '
    ' *** SI ES MISMO PROVEEDOR QUE ANTERIOR DEVUELVO LO ULTIMOS DATOS
    If NPROVE% > 0 And NPROVE% = PROVANT% Then GoTo 99
    '
    ' *** NO RECIBE PROVEEDOR
    If Abs(NPROVE%) = 0 Then
        ALIRETBRUANT# = 0: MINOIMPANT# = 0: MEPAGOANT% = 0
        GoTo 99
    End If
    '
    ' *** NO APLICA SI EL PROVEEDOR ES DEL EXTERIOR O CONSUMIDOR FINAL
    If PIVACLI%(NPROVE% * (-1)) = 5 Or PIVACLI%(NPROVE% * (-1)) = 0 Then
        ALIRETBRUANT# = 0: MINOIMPANT# = 0: MEPAGOANT% = 0
        GoTo 99
    End If
    '
    ' *** RECUPERA DATOS DE LAS CATEGORIAS
    For U& = 1 To ULTREG&("TACAIBTU")
        TCX$ = REGLEIDO$("TACAIBTU", U&)
        TCY% = Asc(Mid$(TCX$, 1, 1))
        If TCY% <= 8 Then
           ALIBRU(TCY%) = CVS(Mid$(TCX$, 33, 4))
           If ALIBRU(TCY%) > 0 Then CTCALCU% = 1
           MINIMPO#(TCY%) = CVS(Mid$(TCX$, 37, 4))
           MEPAGX%(TCY%) = CVI(Mid$(TCX$, 41, 2))
        End If
    Next U&
    '
    
    ' *** INSCRIPCION EN INGRESOS BRUTOS
    CATIBGRAL% = XVALO(VALOBADA("PROVED12", "Caib_Cli", "NUME_CLI=" & CStr(NPROVE%)))
    
    ' *** CATEGORIA INGRESOS BRUTOS
    CATIBTUCU% = XVALO(VALOBADA("PROVED12", "CATEGORIAIBTUCU_CLI", "NUME_CLI=" & CStr(NPROVE%)))
    ' *** DATOS DE LA CATEGORIA
    RETENC# = ALIBRU(CATIBTUCU%)
    MINOMPONIBLE# = MINIMPO#(CATIBTUCU%)
    MEDIOPAGO% = MEPAGX%(CATIBTUCU%)
    '
    ' *** BUSQUEDA EN PADRON
    Encontro% = 0
    If EXISTE%(LUDAT$ + "PAIBTUCU.TXT") > 0 Or EXISTE%(LUDAT$ + "IIBBTUCU.TXT") > 0 Then
       CUITPROV$ = CuitCli$((-1) * Abs(NPROVE%))
       XX$ = REPADIBRU_TUCUMAN$(CUITPROV$)
       If TRIM$(XX$) <> "" Then
          ALICUOTATXT$ = TRIM$(Mid$(XX$, 21))
          ALICUOTA# = XVALO(ALICUOTATXT$)
          ' SI ESTA EN CONVENIO MULTILATERAL, SE RETIENE EL 50% DE ALICUOTA QUE FIGURA EN EL PADRON
          ' SALVO QUE ESTE DENTRO DEL REGIMEN ESPECIAL DE CONVENIO MULTILATERAL (CATIBGRAL% = 3)
          ConvenioMultilateral$ = TRIM$(Mid$(XX$, 19, 2))
          If ConvenioMultilateral$ = "CM" And CATIBGRAL% <> 3 Then ALICUOTA# = (0.5) * ALICUOTA#

          ' SI LA ALICUOTA DEVUELTA ES 0, ENTONCES SE DEBE APLICAR 0.175 %
          If ALICUOTA# < 0.005 Then ALICUOTA = 0.175
          
          Encontro% = 1
       End If
    End If
    '
98  PROVANT% = NPROVE%
    ALIRETBRUANT# = ALICUOTA#
    ' *** SI EL CONTRIBUYENTE NO ESTA REGISTRADO EN EL PADRON,
    ' *** TOMA LA ALICUOTA CONFIGURADA EN LA CATEGORIA ASIGNADA AL PROVEEDOR
    If Encontro% = 0 Then ALIRETBRUANT# = ALIBRU(CATIBTUCU%)
    MINOIMPANT# = MINOMPONIBLE#
    MEPAGOANT% = MEDIOPAGO%
    '
99  ALIRETBRU_TUCUMAN# = ALIRETBRUANT#
    MINOIMP# = MINOIMPANT#
    MEPAGO% = MEPAGOANT%
    '
End Function

Function ALIRETBRU_MISIONES#(NPROVE%, MINOIMP#, MEPAGO%)
    '
    Static ALIBRU(8), MINIMPO#(8), MEPAGX%(8)
    '
    If Abs(NPROVE%) = 0 Then
        ALIRETBRUANT# = 0: MINOIMPANT# = 0: MEPAGOANT% = 0
        GoTo 99
    End If
    '
    If PIVACLI%(NPROVE% * (-1)) = 5 Or PIVACLI%(NPROVE% * (-1)) = 0 Then  ' Si el proveedor es importacion no se realiza calculo
        RETENC# = 0: MINOMPONIBLE# = 0: MEDIOPAGO% = 0
        GoTo 98
    End If
    '
    For U& = 1 To ULTREG&("TACAIBMI")
       TCX$ = REGLEIDO$("TACAIBMI", U&)
       TCY% = Asc(Mid$(TCX$, 1, 1))
       If TCY% > 0 Then
          If TCY% <= 8 Then
             ALIBRU(TCY%) = CVS(Mid$(TCX$, 33, 4))
             MINIMPO#(TCY%) = CVS(Mid$(TCX$, 37, 4))
             MEPAGX%(TCY%) = CVI(Mid$(TCX$, 41, 2))
          End If
       End If
    Next U&
    
    ' *** CATEGORIA DE INGRESOS BRUTOS
    CATIB1% = XVALO(VALOBADA("PROVED12", "CATEGORIAIBMISIONES_CLI", "NUME_CLI=" & CStr(NPROVE%)))
    
    RETENC# = ALIBRU(CATIB1%)
    MINOMPONIBLE# = MINIMPO#(CATIB1%)
    MEDIOPAGO% = MEPAGX%(CATIB1%)
    '
98  PROVANT% = NPROVE%
    ALIRETBRUANT# = RETENC#
    MINOIMPANT# = MINOMPONIBLE#
    MEPAGOANT% = MEDIOPAGO%
    '
99  ALIRETBRU_MISIONES# = ALIRETBRUANT#
    MINOIMP# = MINOIMPANT#
    MEPAGO% = MEPAGOANT%
    '
End Function


Function REPADIBRU_CABA$(CTTT$, Optional MODO%)    ' CTTT$ = CUIT DEL CLIENTE / PROVEEDOR
    '
    Static LREC%
    Static ARCHIIBB$
    Dim XX1 As String * 61
    
    FechaPAIBCABA$ = "01/01/80"
    If EXISTE%(LUDAT$ + "PAIBCABA.TXT") > 0 Then
        FechaPAIBCABA$ = Format(FileSystem.FileDateTime(LUDAT$ + "PAIBCABA.TXT"), "dd/mm/yy")
    End If
    FechaIIBBCABA$ = "01/01/80"
    If EXISTE%(LUDAT$ + "IIBBCABA.TXT") > 0 Then
        FechaIIBBCABA$ = Format(FileSystem.FileDateTime(LUDAT$ + "IIBBCABA.TXT"), "dd/mm/yy")
    Else
        Screen.MousePointer = 1
        Call COMUNI("Imposible Actualizar Padrón.\Falta Archivo IIBBCABA.TXT")
        GoTo 99
    End If
    
    If FECHANUM(FechaIIBBCABA$) > FECHANUM(FechaPAIBCABA$) Then
'        If EXISTE%(LUDAT$ + "PAIBCABA.TXT") > 0 Then
'            On Error Resume Next
'            Call Kill(LUDAT$ & "PAIBCABA.TXT")
'            If Err > 0 Then
'                Call MENSERR(24, "Error al Actualizar Padrón de Alicuotas de IIBB CABA")
'                On Error GoTo 0
'                RPDB$ = ""
'                Screen.MousePointer = 1
'                GoTo 99
'            End If
'            On Error GoTo 0
'        End If
'
'        ARCHIIBB$ = "IIBBCABA.TXT"
'        NARCHI1% = FILEOPEN%(LUDAT$ + "IIBBCABA.TXT", 1, 122)
'        NARCHI2% = FILEOPEN%(LUDAT$ + "PAIBCABA.TXT", 2, 61)
'        While Not EOF(NARCHI1%)
'            Line Input #NARCHI1%, REGISTROPADRON$
'            NREG& = NREG& + 1
'            'Print #NARCHI2%, Mid$(RegistroPadron$, 1, 61)
'            CUITT$ = Mid$(REGISTROPADRON$, 28, 11)
'            ORPAGO07.List3.AddItem CUITT$ & Mid$(REGISTROPADRON$, 1, 61)
'        Wend
'        Close #NARCHI1%
'        'Close #NARCHI2%
'
'        For II& = 1 To ORPAGO07.List3.ListCount - 1
'            Print #NARCHI2%, Mid$(ORPAGO07.List3.LIST(II&), 12, 61)
'        Next II&
'        Close #NARCHI2%
    End If
    '
    CUITPROV$ = CTTT$
    CUITPROVEEDOR$ = REPLACAR$(CUITPROV$, "-", "")
    '
    XX1 = Space(61)
    RPDB$ = ""
    NARCHI% = FILEOPEN%(LUDAT$ + "PAIBCABA.TXT", 0, 61)
    NREG& = LOF(NARCHI%) / 61
    '
    LI& = 0: LS& = NREG& + 1
11  E& = Int((LS& - LI&) / 2)
    If E& < 1 Then GoTo 99

    L& = LI& + E&
    If L& < 1 Or L& > NREG& Then GoTo 99
    Get #NARCHI%, L&, XX1
    CUITPADRON$ = Mid$(XX1, 17, 11)
       If CUITPADRON$ < CUITPROVEEDOR$ Then LI& = L&: GoTo 11
       If CUITPADRON$ > CUITPROVEEDOR$ Then LS& = L&: GoTo 11
    RPDB$ = XX1
    '
99  Close #NARCHI%
    REPADIBRU_CABA$ = RPDB$
    '
End Function

Function REPADIBRU_TUCUMAN$(CTTT$)    ' CTTT$ = CUIT DEL CLIENTE / PROVEEDOR
    '
    Static LREC%
    Static ARCHIIBB$
    Dim XX1 As String * 32
    '
    CUITPROV$ = CTTT$
    CUITPROVEEDOR$ = REPLACAR$(CUITPROV$, "-", "")
    '
    XX1 = Space(32)
    RPDB$ = ""
    NARCHI% = FILEOPEN%(LUDAT$ + "PAIBTUCU.TXT", 0, 32)
    NREG& = LOF(NARCHI%) / 32
    '
    LI& = 0: LS& = NREG& + 1
11  E& = Int((LS& - LI&) / 2)
    If E& < 1 Then GoTo 99

    L& = LI& + E&
    If L& < 1 Or L& > NREG& Then GoTo 99
    Get #NARCHI%, L&, XX1
    CUITPADRON$ = Mid$(XX1, 3, 11)
       If CUITPADRON$ < CUITPROVEEDOR$ Then LI& = L&: GoTo 11
       If CUITPADRON$ > CUITPROVEEDOR$ Then LS& = L&: GoTo 11
    RPDB$ = XX1
    '
99  Close #NARCHI%
    If TRIM$(Mid$(RPDB$, 23, 4)) = "----" Then RPDB$ = ""
    REPADIBRU_TUCUMAN$ = RPDB$
    '
End Function


'
Function ALIPERBRU_CABA#(NCLIE, MINIPER#, VARDO%)
   '
   Static CTTXX%, TASAA, TASAB, TASAEXP, MINIPERCE
   Dim TASAX#
   
   ' *** NO PERCIBE A CLIENTES DEL EXTERIOR NI A CONSUMIDORES FINALES
   If PIVACLI%(NCLIE) = 5 Or PIVACLI%(NCLIE) = 0 Then   ' Si el cliente es exportacion no se realiza calculo
       TASAX# = 0: MINIPERCE = 0
       GoTo 99
   End If
   '
   ' *** VALORES POR DEFAULT
   If CTTXX% < 1 Then
       TASAA = XVALO(Control$("PERIBCA", "TASA-A"))
       TASAB = XVALO(Control$("PERIBCA", "TASA-B"))
       TASAEXP = XVALO(Control$("PERIBCA", "TASA-EXP"))
       MINIPERCE = XVALO(Control$("PERIBCA", "MINIPER"))            ' MINIMO NO IMPONIBLE
       CTTXX% = 1
    End If
    '
    TASAX = TASAA
    If VARDO% = 2 Then TASAX = TASAB
    If VARDO% = 3 Then TASAX = TASAEXP
    '
    If ULTREG&("DAPIBCA") < 1 Then
        REPERCI$ = FILTERGETRECORD$("DAPIBCA", 1, MKI$(Int(NCLIE)))
        CATIB1% = Asc(Mid$(REPERCI$, 3, 1))
        PORCONVE = CVS(Mid$(REPERCI$, 5, 4))
    Else
        REPERCI$ = FILTERGETRECORD$("DAPIBCA", 1, MKS$(Int(NCLIE)))
        CATIB1% = Asc(Mid$(REPERCI$, 5, 1))
        PORCONVE = CVS(Mid$(REPERCI$, 6, 4))
    End If
    '
    If ULTREG&("DAPIBCA") > 0 Then NNY% = 10
    If NNY% = 0 Then NNY% = 1
    If UCase$(Mid$(REPERCI$, NNY%, 18)) = "EXENTO CERTIFICADO" Then
        TASAX = 0
        GoTo 99
    End If
    
    ' *** OBTIENE TASA DE PERCEPCION
    If EXISTE%(LUDAT$ + "PAIBCABA.TXT") > 0 Or EXISTE%(LUDAT$ + "IIBBCABA.TXT") > 0 Then
       CuitClien$ = CuitCli$(NCLIE)
       XX$ = REPADIBRU_CABA$(CuitClien$, 1)
       If TRIM$(XX$) <> "" Then
          PER$ = Mid$(XX$, 35, 4)
          PER$ = REPLACAR$(PER$, ",", ".")
          TASAX# = XVALO(PER$)
       End If
    End If
    '
99  ALIPERBRU_CABA# = TASAX#
    MINIPER# = MINIPERCE
    '
End Function

Function ALIPERBRU#(NCLIE, MINIPER#, VARDO%)
    '
    ' Static CLIEANT%, ALIPERBRUANT#, MINIPERANT#
    ' Linea anterior comentada el 8/5/2009 - DOSIVAC: Si se generaban
    ' dos pedidos seguidos al MISMO CLIENTE, en el segundo traia la
    ' percepcion de IIBB en cero. Al comentar la linea fuerza a repe-
    ' tir la búsqueda.
    '
    Static CTTXX%, TASAA, TASAB, TASAEXP, MINIPERCE
    
    If NCLIE > 0 And NCLIE = CLIEANT% Then GoTo 99
    '
    If Abs(NCLIE) = 0 Then
        ALIPERBRUANT# = 0: MINIPERANT# = 0
        GoTo 99
    End If
    '
    If PIVACLI%(NCLIE) = 5 Then   ' Si el cliente es exportacion no se realiza calculo
        TASAX = 0: MINIPERCE = 0
        GoTo 98
    End If
    '
    If CTTXX% < 1 Then
       TASAA = XVALO(Control$("PERIBRU", "TASA-A"))
       TASAB = XVALO(Control$("PERIBRU", "TASA-B"))
       TASAEXP = XVALO(Control$("PERIBRU", "TASA-EXP"))
       MINIPERCE = XVALO(Control$("PERIBRU", "MINIPER"))
       CTTXX% = 1
    End If
    '
    TASAX = TASAA
    If VARDO% = 2 Then TASAX = TASAB
    If VARDO% = 3 Then TASAX = TASAEXP
    '
    If ULTREG&("DAPIBCLN") < 1 Then
        REPERCI$ = FILTERGETRECORD$("DAPIBCLI", 1, MKI$(Int(NCLIE)))
        CATIB1% = Asc(Mid$(REPERCI$, 3, 1))
        PORCONVE = CVS(Mid$(REPERCI$, 5, 4))
    Else
        REPERCI$ = FILTERGETRECORD$("DAPIBCLN", 1, MKS$(Int(NCLIE)))
        CATIB1% = Asc(Mid$(REPERCI$, 5, 1))
        PORCONVE = CVS(Mid$(REPERCI$, 6, 4))
    End If
    '
    If PORCONVE > 100 Then PORCONVE = 100
    '
    NNY% = 9
    If ULTREG&("DAPIBCLN") > 0 Then NNY% = 10
    If UCase$(Mid$(REPERCI$, NNY%, 18)) = "EXENTO CERTIFICADO" Then
        TASAX = 0
        GoTo 98
    End If
    '
    If CATIB1% > 0 Then         ' NO ESTA EXENTO
      If PORCONVE >= 10.1 Then  ' PARTICIPACION CONVENIO MULTILATERAL
           TASAX = TASAX * PORCONVE / 100
         Else
           TASAX = 0
      End If
    End If
    If CATIB1% < 1 Then TASAX = 0
    '
    If EXISTE%(LUDAT$ + "PADIBRUP.TXT") > 0 Then
       CuitClien$ = CuitCli$(NCLIE)
       XX$ = REPADIBRU$(CuitClien$, 1)
       If TRIM$(XX$) <> "" Then
          PER$ = Mid$(XX$, 48, 4)
          TASAX = XVALO(PER$)
       End If
    End If
    '
98  'CLIEANT = NCLIE
    ALIPERBRUANT# = TASAX
    MINIPERANT# = MINIPERCE
    '
99  ALIPERBRU# = ALIPERBRUANT#
    MINIPER# = MINIPERANT#
    '
End Function
'
Function ALIPERBRU_MISIONES#(NCLIE, MINIPER#, VARDO%)
    '
    Static LEIDO%, TASAA#, TASAB#, MINIPERCE#, CODPROV$
    '
    If LEIDO% < 1 Then
       TASAA# = XVALO(Control$("PERIBMI", "TASA-A"))
       TASAB# = XVALO(Control$("PERIBMI", "TASA-B"))
       MINIPERCE# = XVALO(Control$("PERIBMI", "MINIPER"))
       CODPROV$ = TRIM$(Control$("PERIBMI", "CODPROV"))
       LEIDO% = 1
    End If
    '
    ' ESTABLECE TASA/ALICUOTA A PARTIR DE LA LETRA DEL COMPROBANTE
    TASAX# = TASAA#
    If VARDO% = 2 Then TASAX# = TASAB#
    If VARDO% = 3 Then TASAX# = 0
    '
    REPERCI$ = FILTERGETRECORD$("DAPIBMI", 1, MKS$(Int(NCLIE)))
    CATIB1% = Asc(Mid$(REPERCI$, 5, 1))
    
    ' EXENTO CONFIGURADO POR EL USUARIO DESDE LA FICHA DEL CLIENTE
    If CATIB1% = 0 Then
        TASAX# = 0
        GoTo 99
    End If
    
    ' EXCLUYE A LOS CLIENTES DEL EXTERIOR
    If PIVACLI%(NCLIE) = 5 Then
        TASAX# = 0: MINIPERCE# = 0
        GoTo 99
    End If
    '
99  ALIPERBRU_MISIONES# = TASAX#
    MINIPER# = MINIPERCE#
    '
End Function
Function ALIPERBRU_TUCUMAN#(NCLIE, MINIPER#, VARDO%)
    '
    Static CTTXX%, MINIPERCE#
    Dim TASAX#
    Dim RegistradoEnPadron As Boolean
   
    ' *** VALORES POR DEFAULT
    If CTTXX% < 1 Then
       MINIPERCE# = XVALO(Control$("PERIBTU", "MINIPER"))            ' MINIMO NO IMPONIBLE
       CTTXX% = 1
    End If
        
    ' *** NO PERCIBE A CLIENTES DEL EXTERIOR NI A CONSUMIDORES FINALES
    If PIVACLI%(NCLIE) = 5 Or PIVACLI%(NCLIE) = 0 Then   ' Si el cliente es exportacion no se realiza calculo
       TASAX# = 0: MINIPERCE = 0
       GoTo 99
    End If
    
    ' *** SOLO DEVUELVE PERCEPCION DE IIBB SI ESTA ANULANDO EXPLICITAMENTE UNA FACTURA
    If FORFASQL.TIFADOC% = 6 And FORFASQL.AnulaFacturaComprobante$ = "" Then GoTo 99
    ' *** SI ESTA ANULANDO FACTURA VERIFICA SI EN LA MISMA SE PERCIBIO IIBB POR TUCUMAN
    If FORFASQL.TIFADOC% = 6 And FORFASQL.AnulaFacturaComprobante$ <> "" Then
        CondicionSql$ = "nuclien=" & CStr(NCLIE) & " and codicom_lar='" & FORFASQL.AnulaFacturaComprobante$ & "' and tipomovim='ibru' and opcimovim='ib-tu' and varimovim='perc'"
        ImportePercibido# = XVALO(VALOBADA("cajabanc", "valabscomp", CondicionSql$))
        If ImportePercibido# < 0.005 Then GoTo 99
        
        ' *** VERIFICA SI ESTA QUERIENDO EMITIR UNA NOTA DE CREDITO CON FECHA MAYOR A LA FECHA DE PRESENTACION
        DiaVencimiento% = XVALO(Control$("PERIBTU", "DIAVENCI"))
        Dim FechaFactura, FechaLimite As Date
        FechaFactura = VALOBADA("cajabanc", "fechconta", CondicionSql$)
        FechaLimite = DateAdd("m", 1, FechaFactura)
        FechaLimite = DateSerial(Year(FechaLimite), Month(FechaLimite), DiaVencimiento%)
        
        If Now > FechaLimite Then GoTo 99
    End If
    
    '
    CATIBGRAL% = XVALO(VALOBADA("PROVED12", "Caib_Cli", "NUME_CLI=" & CStr(NCLIE)))
    REPERCI$ = FILTERGETRECORD$("DAPIBTU", 1, MKS$(NCLIE))
    SujetoPercepcion% = Asc(Mid$(REPERCI$, 5, 1))
    
    ' *** OBTIENE TASA DE PERCEPCION
    RegistradoEnPadron = False
    TASAX# = 0
    If EXISTE%(LUDAT$ + "PAIBTUCU.TXT") > 0 Or EXISTE%(LUDAT$ + "IIBBTUCU.TXT") > 0 Then
       CUITCLIENTE$ = CuitCli$(NCLIE)
       XX$ = REPADIBRU_TUCUMAN$(CUITCLIENTE$)
       If TRIM$(XX$) <> "" Then
          PER$ = TRIM$(Mid$(XX$, 21))
          PER$ = REPLACAR$(PER$, ",", ".")
          TASAX# = XVALO(PER$)
          
          ' SI ESTA EN CONVENIO MULTILATERAL, SE RETIENE EL 50% DE ALICUOTA QUE FIGURA EN EL PADRON
          ' SALVO QUE ESTE DENTRO DEL REGIMEN ESPECIAL DE CONVENIO MULTILATERAL (CATIBGRAL% = 3)
          Convenio$ = TRIM$(Mid$(XX$, 19, 2))
          If Convenio$ = "CM" And CATIBGRAL% <> 3 Then TASAX# = (0.5) * TASAX#
          
          RegistradoEnPadron = True
       End If
    End If
    '
99  ALIPERBRU_TUCUMAN# = TASAX#
    If RegistradoEnPadron = False And SujetoPercepcion% = 1 Then ALIPERBRU_TUCUMAN# = 5 ' SI EL CONTRIBUYENTE NO ESTA EN EL PADRON PERO DEBERIA PERCIBIRLE, APLICA UNA TASA DEL 5%
    MINIPER# = MINIPERCE#
    '
End Function
Function REPADIBRU$(CTTT$, Optional MODO%)    ' CTTT$ = CUIT DEL CLIENTE / PROVEEDOR
    '
    Static LREC%
    Static ARCHIIBB$
    Dim XX1 As String * 57
    
    If LREC% < 1 Then
       ARCHIIBB$ = "PADIBRUP.TXT"
       If MODO% = 2 Then ARCHIIBB$ = "PADIBRUR.TXT"
       NARCHI% = FILEOPEN%(LUDAT$ + ARCHIIBB$, 1, 128)
       Line Input #NARCHI%, XX$
       LREC% = 2 + Len(XX$)
       Close #NARCHI%
    End If
    '
    If LREC% <> 57 Then
       Call MENSERR(24, "Longitud de Registro Incorrecta (" + TRIM$(Str$(LREC%)) + ") en PADIBRU.TXT. Consulte.")
       Call FINAL("")
    End If
    '
    CUITPROV$ = CTTT$
    CUITPROVEEDOR$ = REPLACAR$(CUITPROV$, "-", "")
    '
    XX1 = Space$(57)
    RPDB$ = ""
    NARCHI% = FILEOPEN%(LUDAT$ + ARCHIIBB$, 0, LREC%)
    NREG& = LOF(NARCHI%) / 57
    '
    LI& = 0: LS& = NREG& + 1
11     E& = Int((LS& - LI&) / 2)
    If E& < 1 Then GoTo 30
'11  E& = Int((LS& - LI&) / 2) 'AGREGADO POR QUE PARA EL CUIT 30664029452 DEL CLIENTE 305 E& QUEDABA EN 0 Y NO DEVOLVIA EL STRING DEL  PADIBRU
'    If E& < 1 Then
'      If RPDB$ = "" Then 'PARA QUE EN ESTE CASO PASE A LA BUSQUEDA SECUENCIAL
'        GoTo 30
'      Else
'        GoTo 99
'      End If
'    End If

    L& = LI& + E&
    If L& < 1 Or L& > NREG& Then GoTo 30
    Get #NARCHI%, L&, XX1$
    CUITPADRON$ = Mid$(XX1$, 30, 11)
       If CUITPADRON$ < CUITPROVEEDOR$ Then LI& = L&: GoTo 11
       If CUITPADRON$ > CUITPROVEEDOR$ Then LS& = L&: GoTo 11
    RPDB$ = XX1$: GoTo 99          ' LO ENCONTRO EN BUSQUEDA BINARIA
    '
    ' NO LO ENCONTRO - HACE BUSQUEDA SECUENCIAL
30  Close #NARCHI%
    ' SE CAMBIA PORQUE SE ASUME QUE LOS CUITS ESTAN ORDENADOS,
    ' CON LO CUAL LA BUSQUEDA BINARIA SIEMPRE FUNCIONA. O SEA, QUE
    ' SI NO LO ENCUENTRA ES PORQUE NO EXISTE EN PADIBRU
    GoTo 99
    '
    NARCHI% = FILEOPEN%(LUDAT$ + ARCHIIBB$, 1, 128)
    Do While Not EOF(NARCHI%)
       Line Input #NARCHI%, XX$
       CUITPADRON$ = Mid$(XX$, 30, 11)
       If CUITPADRON$ = CUITPROVEEDOR$ Then
          RPDB$ = XX$: GoTo 99     ' LO ENCONTRO SECUENCIAL
       End If
    Loop
    '
99  Close #NARCHI%
    REPADIBRU$ = RPDB$
    '
End Function
'
Function HAYFACVEN%(NCLIE)
   '\\\OT-08-0179-OD-21/04/08///
   HOII% = HOY(HOS$)
   DIAHOY$ = FETEXCOR1$(HOII%)
   HFV% = 0
   '
   Call APERBASDAT("CABAN")
   If SALDCLI#(NCLIE) > 0 Then ' ME PARECE QUE NO ESTA FUNCIONANDO SALDCLI
     SQLX$ = "SELECT SUM (SALDDEBE - SALDACRE) AS DEUDATO "
     SQLX$ = SQLX$ + " FROM SADEUPEN WITH(NOLOCK)"
     SQLX$ = SQLX$ + " WHERE NUCLIEN =" & NCLIE
     SQLX$ = SQLX$ + " AND SALDDEBE > SALDACRE "
     SQLX$ = SQLX$ + " AND FeVencim < '" & DIAHOY & "' "
     Dim FACCCVEN As ADODB.Recordset
     Set FACCCVEN = FLEXCONN.Execute(FILTSQL$(SQLX$))
     '
     With FACCCVEN
       If Not (.EOF And .BOF) Then
         If XVALO(!DEUDATO) > 0 Then HFV% = 1
       End If
     End With
     FACCCVEN.Close
     Set FACCCVEN = Nothing
     '
   End If
   HAYFACVEN% = HFV%
   '\\\OT-08-0179-OD-FIN///
End Function

Sub ALTAPROSCLI(CODICL$)
   '
   Call APERBASDAT("CLIEN")
   '
   ' BUSCA EL PROXIMO NUMERO DE CLIENTE LIBRE - RUTINA MODIFICADA EL 15/04/2010
   Screen.MousePointer = 11
   PRONUCLI& = XVALO(Control$("", "PRONUCLI"))
   If PRONUCLI& < 0 Then
        NROPRO& = XVALO(VALOBADA("deudor12", "max(nume_cli)", "nume_cli>0"))
        GoTo 3
   Else
        If PRONUCLI& > 999999 Then PRONUCLI& = 0
        If InStr(1, EMPREAC$, "MEDITEA") > 0 Then
           If CantRegistros("DEUDOR12", "Nume_Cli = 40001") > 0 Then
                GoTo 2
           Else
                PRONUCLI& = 40000
           End If
        End If
   End If
2  NROPRO& = PRONUCLI&
   If NROPRO& < 0 Then NROPRO& = 0
3  NROPRO& = NROPRO& + 1

   NCLIE = NROPRO&
   
   ' *** PRESENTA VENTANA DE ALTA CON DATOS BASICOS SEGUN PROSPECTO
   ConsultaSql$ = "SELECT * FROM BSEGCOM WITH(NOLOCK) WHERE CODI_CLI='" & CODICL$ & "'"
   Dim Rst_BSEGCOM As ADODB.Recordset
   Set Rst_BSEGCOM = READSET(ConsultaSql$)
   With Rst_BSEGCOM
        If Not (.EOF And .BOF) Then
             RASO$ = SAFETEXT$(!raso_cli)
             DOMI$ = SAFETEXT$(!domi_cli)
             Mail_Cli$ = SAFETEXT$(!Mail_Cli)
             Cpos_Cli$ = SAFETEXT$(!Cpos_Cli)
             loca_Cli$ = SAFETEXT$(!loca_Cli$)
             Prov_Cli$ = SAFETEXT$(!Prov_Cli)
             Pais_Cli$ = SAFETEXT$(!Pais_Cli)
             Vend_Cli$ = SAFETEXT$(!Vend_Cli)
             Tele_Cli$ = SAFETEXT$(!Tele_Cli)
             FAX_CLI$ = SAFETEXT$(!FAX_CLI)
             CTAC_CLI$ = SAFETEXT$(!CTAC_CLI)
             nota_Cli$ = SAFETEXT$(!nota_Cli)
        Else
            Call MENSERR(24, "Problema al Acceder a los Datos del Prospecto")
            Rst_BSEGCOM.Close
            Set Rst_BSEGCOM = Nothing
            Exit Sub
        End If
   End With
   Rst_BSEGCOM.Close
   Set Rst_BSEGCOM = Nothing

   ' VERIFICA SI EN EL MAESTRO DE CLIENTES SE CONFIGURO PARA QUE LA RAZON SOCIAL ESTE SIEMPRE EN MAYUSCULA
   If Control$("MASTER", "RASOMAYU") = "SI" Then
        RASO$ = UCase$(RASO$)
   End If
   '
   If POXIVA% < 1 Then POXIVA% = 1
   VDEF$ = String$(128, 0)
   Call REPLA(VDEF$, CODICL$, 1, 12)
   Call REPLA(VDEF$, MKS$(NCLIE), 13, 4)
   Call REPLA(VDEF$, RASO$, 17, 62)
   Call REPLA(VDEF$, Chr$(1), 79, 1)
   Call REPLA(VDEF$, Space$(15), 80, 15)
   Call REPLA(VDEF$, DOMI$, 95, 49)
   '
10 Screen.MousePointer = 1
   OBX$ = OBJPLA$("ALTACLIENACC2", VDEF$)
   If OBX$ = "" Then GoTo 99
   VDEF$ = OBX$
   
   ' *** VALIDA DATOS BASICOS
   '
   NROPRO& = CVS(Mid$(OBX$, 13, 4))
   If NROPRO& <> NCLIE Then
     NC& = XVALO(VALOBADA("DEUDOR12", "Nume_Cli", "Nume_Cli=" & NROPRO&))
     If NC& > 0 Then GoTo 3
   End If
   NCLIE = NROPRO&
   '
   CODICL1$ = TRIM$(Mid$(OBX$, 1, 12))
   If CODICL1$ <> CODICL$ Then
      Call MENSERR(24, "Ingreso no Válido.\No Permitido Cambiar Codigo de Cuenta")
      Call REPLA(OBX$, CODICL$, 1, 12)
      GoTo 10
   End If
   '
   DCLINUE$ = TRIM$(Mid$(OBX$, 17, 62))
   If Len(DCLINUE$) < 1 Then
      Call MENSERR(24, "Ingreso no Válido.\Falta Razón Social del Cliente")
      GoTo 10
   End If
   ' VERIFICA SI EN EL MAESTRO DE CLIENTES SE CONFIGURO PARA QUE LA RAZON SOCIAL ESTE SIEMPRE EN MAYUSCULA
   If Control$("MASTER", "RASOMAYU") = "SI" Then
        DCLINUE$ = UCase$(DCLINUE$)
   End If
   '
   PIVCLI% = Asc(Mid$(OBX$, 79, 1))
   If PIVCLI% < 0 Or PIVCLI% > 7 Then
      Call MENSERR(24, "Ingreso no Válido.\Posición IVA no Figura en Tablas.")
      GoTo 10
   End If
   '
   
   CUTT$ = TRIM$(Mid$(OBX$, 80, 15))
   
   If Len(CUTT$) < 9 Then
        Opcion% = COMALTER("Se Ha Detectado el Ingreso de un DNI. ¿Desea Obtener el Cuit Asociado?\\Si, Obtener\No, Continuar")
        If Opcion% = 1 Then
            CUTT$ = ObtieneCuitDesdeDni(CUTT$)
            Call REPLA(OBX$, CUTT$, 80, 15)
            VDEF$ = OBX$
            GoTo 10
        End If
        If Opcion% < 1 Then GoTo 10
'        GoTo 20 'continuar
   Else
        CUTTSINGUION$ = REPLACAR$(CUTT$, "-", "")
        If Len(TRIM$(CUTTSINGUION$)) > 12 Then
            Call MENSERR(24, "El Límite Máximo Para Número de Documento es de 12 Dígitos")
            GoTo 10
        End If

   End If
   
   If VALICUIT%(CUTT$) <> 1 Then
     If PIVCLI% <> 0 Then
       If PIVCLI% <> 5 Then
         Call MENSERR(24, "Ingreso no Válido.\Número de C.U.I.T. Incorrecto.")
         GoTo 10
       End If
     End If
   End If
   '
   
   DOMI$ = TRIM$(Mid$(OBX$, 95, 49))
   If DOMI$ = "" Then
         Call MENSERR(24, "Falta Ingresar el Domicilio")
         GoTo 10
   End If
   '
   ' *** VALIDA POSICION DE IVA CONTRA AFIP
   Dim oAltaCliente As New DatosEmpresaAfip
   condicionIvaAfip% = oAltaCliente.ObtenerIdCondicionIvaAfip(CUTT$)
   If condicionIvaAfip% >= 0 And condicionIvaAfip% <> PIVCLI% Then
        denominacionPosIvaCliente$ = ECOARCH$("CAPOSIV", Chr$(PIVCLI%))
        denominacionPosIvaAfip$ = ECOARCH$("CAPOSIV", Chr$(condicionIvaAfip%))
        opcionElegida% = COMALTER%("Se Ha Encontrado que la Posición de Iva Ingresada para el Cliente\" & TRIM$(DCLINUE$) & " \No Coincide con la Registrada en Afip\ \Ingresado como: " & denominacionPosIvaCliente$ & "\Según Afip: " & denominacionPosIvaAfip$ & "\ \¿Desea Considerar la que Posee Afip?\\Si, Considerar Afip\No, Utilizar la Ingresada")
        If opcionElegida% = 1 Then
            Call REPLACE$(VDEF$, Chr$(condicionIvaAfip%), 79, 1)
            Call REPLACE$(VDEF$, oAltaCliente.Cuit, 80, 15)
            GoTo 10
        End If
   End If
   '
   ' *** VALIDA EL USO DEL CUIT EN OTRO CLIENTE
   '
   'SE CAMBIO Y SE VALIDA QUE NO ESTE REPETIDO Y SOLO PERMITE REPETIR CON DERECHOS
   If VALICUIT%(CUTT$) = 1 Then CUTT$ = FormatearCuitValido$(CUTT$)
   CONDI$ = "Cuit_Cli = '" & CUTT$ & "' AND nume_cli <> " & NCLIE
   razonSocial$ = TRIM$(VALOBADA("DEUDOR12", "RASO_CLI", CONDI$, "NOMESS"))
   If razonSocial$ <> "" Then
        Screen.MousePointer = 1
        OPX% = COMALTER%("Ciente: '" & razonSocial$ & "' Posee el Mismo Número de C.U.I.T.\\Cancelar\Grabar Igual")
        If OPX% < 2 Then Exit Sub
        '
        If DERACCE%("REPICLIE", "Repetir C.U.I.T. en Diferentes Clientes") < 2 Then
          Exit Sub
        End If
   End If

   ' *** AGREGAR EL NUEVO CLIENTE A LA TABLA
   '
   ' VERIFICA SI EL NUMERO DE CLIENTE ESTA OCUPADO. EN CASO DE QUE SI, SUMA UNO Y VUELVE A PRESENTAR
   ' LA VENTANA INICIAL. CASO CONTRARIO CONTINUA CON EL ALTA DEL CLIENTE
   Call BLOQARCH("DEUDOR2")
   Dim Rdeudor12 As ADODB.Recordset
   Set Rdeudor12 = New ADODB.Recordset
   Rdeudor12.CursorType = adOpenKeyset
   Rdeudor12.LockType = adLockPessimistic
25 On Error Resume Next
   Rdeudor12.Open FILTSQL$("Select * from Deudor12 where nume_cli=" & NCLIE), FLEXCONN, , , adCmdText
   If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
      On Error GoTo 0
      Call CapturaErrorOpen
      GoTo 25
   End If
   On Error GoTo 0
   If Not (Rdeudor12.EOF And Rdeudor12.BOF) Then
      NCLIE = NCLIE + 1
      Call LIBARCH("DEUDOR2")
      Call COMUNI("FLEXOFT ha Re-Numerado el Cliente por '" + TRIM$(FORMATNUM$(NCLIE, "I5")) + "'.\  \Verifique y Vuelva a Grabar.")
      VDEF$ = OBX$: Call REPLA(VDEF$, MKI$(NCLIE), 13, 2)
      GoTo 10
   End If
   CONDI$ = "Codi_Cli='" & CODICL$ & "'"
   With Rdeudor12
      .AddNew   ' YA CONVERTIDO
      On Error Resume Next
      !CODIEMPR = CodiEmp%
      On Error GoTo 0
      !Codi_Cli = CODICL$
      !nume_cli = TRIM$(Str$(NCLIE))
      !raso_cli = DCLINUE$
      !Mail_Cli = Mail_Cli$
      !Cuit_Cli = CUTT$
      !domi_cli = DOMI$
      !Cpos_Cli = Cpos_Cli$
      !loca_Cli = loca_Cli$
      !Prov_Cli = Prov_Cli$
      !Pais_Cli = Pais_Cli$
      !Vend_Cli = XVALO(Vend_Cli$)
      !Tele_Cli = Tele_Cli$
      !Faxi_Cli = Faxi_Cli$
      !CTAC_CLI = CTAC_CLI$
      !nota_Cli = nota_Cli$
     
      !Piva_Cli = PIVCLI% ' para que cambie la posicion de iva caundo viene del prospect
      !Stat_Cli = 1
      On Error Resume Next
      .Update
      If Err Then
         Call MENSERR(24, "Alta no Procesada Correctamente")
         Exit Sub
      End If
      
      ' ASIGNA EL NUMERO DE CLIENTE A LOS PRESUPUESTOS DEL PROSPECTO
      SQLX$ = "UPDATE PRESUVEN SET NroClie=" & TRIM$(FORMATNUM$(NCLIE, "I5")) & " WHERE Codi_Cli='" & CODICL$ & "'"
      FLEXCONN.Execute SQLX$
      On Error GoTo 0
   End With
   '
   Rdeudor12.Close
   Set Rdeudor12 = Nothing
   Call LIBARCH("DEUDOR2")
   '
   ' GENERA LISTA DE CLIENTES
   Call CARGALISEL("DEUDOR1", "DEUDOR12", "Nume_Cli/I6;Raso_Cli/A30", "NUME_CLI > 0 ORDER BY RASO_CLI ASC")
   
   ' GENERA LISTA PARA ECO
   If ULTREG&("DEUDOR12") > 0 Then
        XX$ = REGBLAN$("DEUDOR12")
        Call REPLA(XX$, MKS$(NCLIE), 1, 4)
        Call REPLA(XX$, DCLINUE$, 5, 28)
        Call REGAPP("DEUDOR12", XX$)
        Call FILESORT("DEUDOR12", "", 1, 1)
   Else
        Dim Rst As ADODB.Recordset
        SQLX$ = "select nume_cli, raso_cli from deudor12 with(nolock) where stat_cli>0"
        Set Rst = READSET(SQLX$)
        '
        With Rst
          JJ& = 0
          Do While Not .EOF
              NCLIEN = XVALO(!nume_cli)
              RASOC$ = SAFETEXT$(!raso_cli)
              '
              XX$ = REGBLAN$("DEUDOR12")
              Call REPLA(XX$, MKS$(NCLIEN), 1, 4)
              Call REPLA(XX$, RASOC$, 5, 28)
              JJ& = JJ& + 1
              Call GRAREG("DEUDOR12", XX$, JJ&)
              .MoveNext
          Loop
        End With
        On Error Resume Next
        Rst.Close
        Set Rst = Nothing
        On Error GoTo 0
        Call SETULTREG("DEUDOR12", JJ&)
        Call FRESHECHO("DEUDOR12")
        '
   End If
   '
   '
99 Screen.MousePointer = 1
End Sub


Function ObtieneCuitDesdeDni(Dni As String) As String

    Dim DniDigitos(8) As Integer
    Dim OperaDigitos(10) As Integer
    Dim X, Y, Z As Byte
    
    SEXO% = COMALTER%("Seleccione el Sexo\\Femenino\Masculino")
    If SEXO% = 1 Then
        X = 2: Y = 7
    ElseIf SEXO% = 2 Then
        X = 2: Y = 0
    Else
        Exit Function
    End If
    
    If IsNumeric(Dni) Then
        DniNumero# = XVALO(Dni)
        If Len(Dni) < 8 Then
            While Len(Dni) < 8
                Dni = "0" & Dni
            Wend
        End If
        For i% = 8 To 1 Step -1
            DniDigitos(i%) = DniNumero# Mod 10
            DniNumero# = Int(DniNumero# / 10)
        Next i%
        
        OperaDigitos(1) = X * 5
        OperaDigitos(2) = Y * 4
        OperaDigitos(3) = DniDigitos(1) * 3
        OperaDigitos(4) = DniDigitos(2) * 2
        OperaDigitos(5) = DniDigitos(3) * 7
        OperaDigitos(6) = DniDigitos(4) * 6
        OperaDigitos(7) = DniDigitos(5) * 5
        OperaDigitos(8) = DniDigitos(6) * 4
        OperaDigitos(9) = DniDigitos(7) * 3
        OperaDigitos(10) = DniDigitos(8) * 2
        
        For i% = 1 To 10
            SumaOperaDigitos% = SumaOperaDigitos% + OperaDigitos(i%)
        Next i%
        
        Resto% = SumaOperaDigitos% Mod 11
        
        If Resto% = 0 Then
            DIGITOVERIFICADOR% = 0
            GoTo 90
        End If
        
        If Resto% = 1 Then
            If SEXO% = 1 Then
                X = 2: Y = 3
                DIGITOVERIFICADOR% = 4
            End If
            If SEXO% = 2 Then
                X = 2: Y = 3
                DIGITOVERIFICADOR% = 9
            End If
            GoTo 90
        End If
                        
        DIGITOVERIFICADOR% = 11 - Resto%
        
90      Cuit$ = CStr(X) & CStr(Y)
        For i% = 1 To 8
            Cuit$ = Cuit$ & CStr(DniDigitos(i%))
        Next i%
        Cuit$ = Cuit$ & CStr(DIGITOVERIFICADOR%)
        
        ObtieneCuitDesdeDni = Cuit$
    End If
    
End Function
Sub MUESTRAVENCIDAS(NC%, DIAS%)
    '
Exit Sub
    HOII% = HOY(HO$)
    FLIMITE% = DIANTE(HOII%)
    Do While DIENTRE(FLIMITE%, HOII%) < DIAS%
    FLIMITE% = DIANTE(FLIMITE%)
    Loop
    
    FIMISQL$ = FETEXCOR1$(FLIMITE%)
    
    CONDI$ = "Nuclien = " + Str(NC%) + " AND FEVENCIM <= '" + FIMISQL$ + "' AND SaldDebe >= 0.005"
    XX = XVALO(VALOBADA("SADEUPEN", "COUNT (NUCLIEN)", CONDI$))
    '
    If XX > 0 Then
        SQLS$ = "SELECT * FROM SADEUPEN WHERE " + CONDI$ + ""
        Set DEUDATMP = FLEXCONN.Execute(FILTSQL$(SQLS$))
        With DEUDATMP
           Do While Not .EOF
             SALDEB# = XVALO(!SaldDebe)
             SALACR# = XVALO(!SaldAcre)
             DEUDA$ = FORMATNUM$(SALDEB# - SALACR#, "F12.2")
             IMPORIG# = XVALO(!IBruComp)
             FEVE = CVINT(!FEVENCIM)
             NUMFACT$ = SAFETEXT(!codicom_lar)
             FECHOR = CVINT(!FECHEMISI)
             Z$ = Space$(128)
             Call REPLA(Z$, NUMFACT$, 1, 18)
             Call REPLA(Z$, FECHATEX$(FECHOR), 21, 8)
             Call REPLA(Z$, FORMATNUM(IMPORIG#, "F11.2"), 30, 11)
             Call REPLA(Z$, FECHATEX$(FEVE), 42, 8)
             Call REPLA(Z$, DEUDA$, 50, 12)
             FACVENCI07.List1.AddItem (Z$)
               '
            .MoveNext
            Loop
        End With
        FACVENCI07.Show 1
        
    End If
    '
End Sub

Function DESCURECA(CP%)
   '
   DESREC& = 0
   '
   If CP% > 0 Then ' DEVUELVE EL DESCUENTO O RECARGO SEGUN LA CONDICION DE PAGO
    If CP% <= ULTREG&("CONPAG") Then
     REG& = CP%
     X$ = REGLEIDO$("CONPAG", REG&)
     DESREC = CVS(Mid$(X$, 44, 4))
    End If
   End If
   '
   DESCURECA = DESREC
   '
End Function

Sub GUARDAVALCOT(NUMECOTI$, CLAVEVAL$, valor$, Optional COLUVALOR)
    'ESTA RUTINA ACTUALIZA O CARGA UN NUEVO REGISTRO EN LA TABLA DETCOTIZA (DETALLE DE COTIZACIONES)
    'SEGUN EL NUMERO DE COTIZACION Y LA CLAVE
    '
    CONDI$ = "NroCoti = '" & NUMECOTI$ & "' AND ClaValor = '" & CLAVEVAL$ & "'"
    '
    rr& = CantRegistros&("DETCOTIZA", CONDI$)
    If rr& = 0 Then
        Call ABREPEDCLI
        SQLX$ = "SELECT * FROM DETCOTIZA"
        Dim DECOTIZATMP As ADODB.Recordset
        Set DECOTIZATMP = New ADODB.Recordset
25      On Error Resume Next
        DECOTIZATMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenDynamic, adLockPessimistic, adCmdText
        If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
           On Error GoTo 0
           Call CapturaErrorOpen
           GoTo 25
        End If
        On Error GoTo 0
        '
        With DECOTIZATMP
            .AddNew   ' YA CONVERTIDO
            On Error Resume Next
            !CODIEMPR = CodiEmp%
            On Error GoTo 0
                !NROCOTI = NUMECOTI$
                !ClaValor = CLAVEVAL$
                If XVALO(COLUVALOR) < 2 Then
                    !VALOR1 = TRIM$(valor$)
                ElseIf XVALO(COLUVALOR) = 2 Then
                    !VALOR2 = TRIM$(valor$)
                ElseIf XVALO(COLUVALOR) = 3 Then
                    !Valor3 = TRIM$(valor$)
                End If
            .Update
         End With
    Else
        If XVALO(COLUVALOR) < 2 Then
            Call ACTUREGISTRO("DETCOTIZA", "Valor1", CONDI$, TRIM$(valor$), RAF&, "NOMESS")
        ElseIf XVALO(COLUVALOR) = 2 Then
            Call ACTUREGISTRO("DETCOTIZA", "Valor2", CONDI$, TRIM$(valor$), RAF&, "NOMESS")
        ElseIf XVALO(COLUVALOR) = 3 Then
            Call ACTUREGISTRO("DETCOTIZA", "Valor3", CONDI$, TRIM$(valor$), RAF&, "NOMESS")
        End If
    End If
    '
End Sub

Function VALEIDCOT(NUMECOTI$, CLAVEVAL$, Optional COLUVALOR)
    'ESTA FUNCION DEVUELVE EL REGISTRO DE LA TABLA DETCOTIZA (DETALLE DE COTIZACIONES)
    'SEGUN EL NUMERO DE COTIZACION Y LA CLAVE
    '
    CONDI$ = "NroCoti = '" & NUMECOTI$ & "' AND ClaValor = '" & CLAVEVAL$ & "'"
    '
    If XVALO(COLUVALOR) < 2 Then
        VALEIDCOT = TRIM$(VALOBADA("DETCOTIZA", "Valor1", CONDI$, "NOMESS"))
    ElseIf XVALO(COLUVALOR) = 2 Then
        CONDI$ = CONDI$ & "AND VALOR2 IS NOT NULL"
        VALEIDCOT = TRIM$(VALOBADA("DETCOTIZA", "Valor2", CONDI$, "NOMESS"))
    ElseIf XVALO(COLUVALOR) = 3 Then
        CONDI$ = CONDI$ & "AND VALOR3 IS NOT NULL"
        VALEIDCOT = TRIM$(VALOBADA("DETCOTIZA", "Valor3", CONDI$, "NOMESS"))
    End If
    '
End Function

Function DESCOMBINA$(NCLIEN, CIX%, Optional MODODESC%)
   '
   'MODODESC% = DEVUELVE EL MODO DE DESCUENTO SI ES MAYOR A 0 = PRODUCTO BONIFICADO
   DCX$ = ""
   MODO% = 0
   MODOBON% = 0
   If NCLIEN > 0 Then
     If CIX% > 0 Then
       GRX% = GRUCOCLI%(NCLIEN)
       LNX% = GRUCOVEI%(CIX%)
       
       'CLIENTE Y CODIGO
       CONDI$ = "Nume_Cli = " & NCLIEN & " AND GrCo_Cli = 0 AND Cod_Inte=" & CIX% & " AND GruCoVen = 0"
       DCTEMP$ = VALOBADA("DESCOMBI", "PORDESCU", CONDI$, "NOMESS")
       MODOBON% = XVALO(VALOBADA("DESCOMBI", "MODO_DESC", CONDI$, "NOMESS"))
       If DCTEMP$ <> "" Then
          DCX$ = DCTEMP$
          MODO% = MODOBON%
          GoTo 99
       End If
       '
       'CLIENTE Y LINEA
       CONDI$ = "Nume_Cli = " & NCLIEN & " AND GrCo_Cli = 0 AND Cod_Inte = 0 AND GruCoVen = " & LNX%
       DCTEMP$ = VALOBADA("DESCOMBI", "PORDESCU", CONDI$, "NOMESS")
       MODOBON% = XVALO(VALOBADA("DESCOMBI", "MODO_DESC", CONDI$, "NOMESS"))
       If DCTEMP$ <> "" Then
          DCX$ = DCTEMP$
          MODO% = MODOBON%
          GoTo 99
       End If
       '
       'CLIENTE SOLO
       CONDI$ = "Nume_Cli = " & NCLIEN & " AND GrCo_Cli = 0 AND Cod_Inte = 0 AND GruCoVen = 0"
       DCTEMP$ = VALOBADA("DESCOMBI", "PORDESCU", CONDI$, "NOMESS")
       MODOBON% = XVALO(VALOBADA("DESCOMBI", "MODO_DESC", CONDI$, "NOMESS"))
       If DCTEMP$ <> "" Then
          MODO% = MODOBON%
          DCX$ = DCTEMP$
          GoTo 99
       End If
       '
       'GRUPO Y CODIGO
       CONDI$ = "Nume_Cli = 0 AND GrCo_Cli = " & GRX% & " AND Cod_Inte=" & CIX% & " AND GruCoVen = 0"
       DCTEMP$ = VALOBADA("DESCOMBI", "PORDESCU", CONDI$, "NOMESS")
       MODOBON% = XVALO(VALOBADA("DESCOMBI", "MODO_DESC", CONDI$, "NOMESS"))
       If DCTEMP$ <> "" Then
          MODO% = MODOBON%
          DCX$ = DCTEMP$
          GoTo 99
       End If
       '
       'GRUPO Y LINEA
       DCX$ = ""
       CONDI$ = "Nume_Cli = 0 AND GrCo_Cli = " & GRX% & " AND Cod_Inte = 0 AND GruCoVen = " & LNX%
       DCTEMP$ = VALOBADA("DESCOMBI", "PORDESCU", CONDI$, "NOMESS")
       MODOBON% = XVALO(VALOBADA("DESCOMBI", "MODO_DESC", CONDI$, "NOMESS"))
       If DCTEMP$ <> "" Then
          MODO% = MODOBON%
          DCX$ = DCTEMP$
          GoTo 99
       End If
       '
       'GRUPO Y TODOS LOS ARTICULOS
       DCX$ = ""
       CONDI$ = "Nume_Cli = 0 AND GrCo_Cli = " & GRX% & " AND Cod_Inte = 0 AND GruCoVen = 0"
       DCTEMP$ = VALOBADA("DESCOMBI", "PORDESCU", CONDI$, "NOMESS")
       MODOBON% = XVALO(VALOBADA("DESCOMBI", "MODO_DESC", CONDI$, "NOMESS"))
       If DCTEMP$ <> "" Then
          MODO% = MODOBON%
          DCX$ = DCTEMP$
          GoTo 99
       End If
       '
       'TODOS LOS CLIENTES Y ARTICULOS
       DCX$ = ""
       CONDI$ = "Nume_Cli = 0 AND GrCo_Cli = 0 AND Cod_Inte = " & CIX% & " AND GruCoVen = 0"
       DCTEMP$ = VALOBADA("DESCOMBI", "PORDESCU", CONDI$, "NOMESS")
       MODOBON% = XVALO(VALOBADA("DESCOMBI", "MODO_DESC", CONDI$, "NOMESS"))
       If DCTEMP$ <> "" Then
          MODO% = MODOBON%
          DCX$ = DCTEMP$
          GoTo 99
       End If
       '
       'TODOS LOS CLIENTES Y LINEA
       DCX$ = ""
       CONDI$ = "Nume_Cli = 0 AND GrCo_Cli = 0 AND Cod_Inte = 0 AND GruCoVen = " & LNX%
       DCTEMP$ = VALOBADA("DESCOMBI", "PORDESCU", CONDI$, "NOMESS")
       MODOBON% = XVALO(VALOBADA("DESCOMBI", "MODO_DESC", CONDI$, "NOMESS"))
       If DCTEMP$ <> "" Then
          MODO% = MODOBON%
          DCX$ = DCTEMP$
          GoTo 99
       End If
       'TODOS LOS CLIENTES Y TODOS
       DCX$ = ""
       CONDI$ = "Nume_Cli = 0 AND GrCo_Cli = 0 AND Cod_Inte = 0 AND GruCoVen = 0"
       DCTEMP$ = VALOBADA("DESCOMBI", "PORDESCU", CONDI$, "NOMESS")
       MODOBON% = XVALO(VALOBADA("DESCOMBI", "MODO_DESC", CONDI$, "NOMESS"))
       If DCTEMP$ <> "" Then
          DCX$ = DCTEMP$
          MODO% = MODOBON%
          GoTo 99
       End If
       '
     End If
   End If
   '
99    DESCOMBINA$ = DCX$
      MODODESC% = MODO%
End Function

Function LCred_CCte(NC)

   LCred_CCte = 0
   On Error Resume Next
   If STATCLI%(NC) <> 0 Then LCred_CCte = CDbl(ClienTemp!LCred_CCte)

End Function


Function EXEDELIMCC%(NCLIEN, VALORFA, VALOEXEDIDO)
    '
    ' VERIFICA SI EL LIMITE DE CREDITO EN CUENTA CORRIENTE ESTA SUPERADO POR
    ' LA SALDO + EL VALOR PASADO COMO PARAMETRO
    '  0 > NO
    '  1 > SI
    EXEDELIMCC% = 0
    VALOEXEDIDO = 0
    If Control("ADMIPED", "VRLICCTE") <> "SI" Then
       Exit Function
    End If
    '
    If (XVALO(VALORFA) + SALDCLI#(NCLIEN)) > LCred_CCte(NCLIEN) + 0.005 Then '
        VALOEXEDIDO = (XVALO(VALORFA) + SALDCLI#(NCLIEN)) - LCred_CCte(NCLIEN)
        EXEDELIMCC% = 1
    End If
    
End Function

Function INDIC#(NC)
   '
   ' indice de calidad de proveedores.
   On Error Resume Next
   If STATCLI%(NC) <> 0 Then INDIC# = TRIM$(ClienTemp!CALINTEG)
   '
End Function

Function CLASE$(indice#)
    ' devuelve la clase segun el indice
    ' si el indice es 0 o no existe el ejecutable devuelve vacio
    CLASE$ = ""
    
    If EXISTE%("CALPRO09.EXE") > 0 And indice# > 0 Then
        If indice# >= XVALO(Control("CALIPROV", "LIMCAT-A ")) Then CLASE$ = "A"
        If indice# >= XVALO(Control("CALIPROV", "LIMCAT-B ")) And indice# < XVALO(Control("CALIPROV", "LIMCAT-A ")) Then CLASE$ = "B"
        If indice# < XVALO(Control("CALIPROV", "LIMCAT-B ")) Then CLASE$ = "C"
    End If
    
End Function

Function ENVMAIDEUDA%(NC)
   ' Devuelve 0 = "No envia", y 1 = "Envia"
   ENVMAIDEUDA% = 0
   On Error Resume Next
   If STATCLI%(NC) <> 0 Then ENVMAIDEUDA% = XVALO(ClienTemp!ENVMAIDEUDA)
   '
End Function

Sub MUESTRADOC(DOCUNU&, Optional MODOX$, Optional Tipo$, Optional NUBUSQUEDA$, Optional NUCOMPLE$, Optional VARCONTROL$, Optional MENSAJEDOC$, Optional NoPuedeImprimir%)
    '
'    DOCUNU& = DOCUNUME&
    'VALIDACION SI DOCUNU& < 1 BUSCA EN NUEVA FUNCION DE BUSQUEDA DE NUMERO DE DOCUMENTO
    If DOCUNU& < 1 Then
      VARCONTROL1$ = VARCONTROL$ 'MODIFICADO POR QUE EN DOCID ESTE PARAMETRO SE UTILIZA PARA EL CASO DE FECHAS.
      If TRIM$(UCase$(VARCONTROL$)) = "NOMESS" Then VARCONTROL1$ = ""
      DOCUNU& = DOCID&(Tipo$, NUBUSQUEDA$, NUCOMPLE$, VARCONTROL1$)
    End If
    '
    If DOCUNU& < 1 Then
      If VARCONTROL$ <> "NOMESS" Then
         Call COMUNI("Documento no Encontrado")
      End If
      Exit Sub
    End If
    '
    Call ABRECONEXION
    CONDI$ = ("NUMEDOC =" & DOCUNU&)
    REFX$ = TRIM$(SAFETEXT$(VALOBADA("PDOC_TBL", "REFERENCIA", CONDI$, "NOMESS")))
    If REFX$ = "" Then
       XXX$ = FILTERGETRECORD$("PDOCLST", 1, MKS$(DOCUNU&))
       REFX$ = TRIM$(Mid$(XXX$, 5, 60))
    End If
    EDITFORM.Label98 = REFX$
    '
5   For KKI% = 0 To 5
      EDITFORM.Picture1(KKI%).Cls
      EDITFORM.Picture1(KKI%).DrawWidth = 3
      EDITFORM.Picture1(KKI%).DrawStyle = 0
      EDITFORM.Picture1(KKI%).FillStyle = 1
    Next KKI%
    '
    'AQUI VALIDO SI EXISTE EN SQL
    TPSP$ = VALOBADA("PDOC_TBL", "IMAGEN", CONDI$, "NOMESS")
    '
    'SI EXISTE NO BUSCA EN ARCHIVO
    If TPSP$ <> "" Then GoTo 96
    
    Dim TD0 As String * 512
    ASPL$ = "PDOCSPL.DAT"
    NX0% = FILEOPEN%(LUDAT$ + ASPL$, 0, 512)
    Get #NX0%, 1, TD0$
    If CVS(Left$(TD0$, 4)) > DOCUNU& Then
55    Close #NX0%
      ASPL$ = "PDOCSPL.DT0"
      NX0% = FILEOPEN%(LUDAT$ + ASPL$, 0, 512)
    End If
    '
    URE& = Abs(LOF(NX0%) / 512)
    LI& = 0: LS& = URE&
6   E& = Int((LS& - LI&) / 2): If E& = 0 Then GoTo 7
    L& = LI& + E&: If L& < 1 Or L& > URE& + 1 Then GoTo 7
    Get #NX0%, L&, TD0$
    DCC& = CVS(Left$(TD0$, 4))
    If DCC& < DOCUNU& Then LI& = L&: GoTo 6
    If DCC& > DOCUNU& Then LS& = L&: GoTo 6
    GoTo 8
    '
7   If ASPL$ = "PDOCSPL.DAT" Then
       GoTo 55
    End If
    Screen.MousePointer = 1
    Call MENSERR(24, "Documento no Registrado")
    GoTo 99
    '
8   TPSP$ = ""
    Screen.MousePointer = 11
    While L& > 1
      Get #NX0%, L&, TD0$
      DCC& = CVS(Left$(TD0$, 4))
      If DCC& < DOCUNU& Then GoTo 9
      L& = L& - 1
    Wend
    '
9   CAPAGINAS% = 0
    For KKL& = L& To URE&
      Get #NX0%, KKL&, TD0$
      TBUF$ = TD0$
      If CVS(Left$(TBUF$, 4)) > DOCUNU& Then GoTo 95
      If CVS(Left$(TBUF$, 4)) = DOCUNU& Then
        TPSP$ = TPSP$ + Mid$(TBUF$, 5)
      End If
    Next KKL&
    '
95  Close #NX0%
96  If MODOX$ = "RECUP" Then
20      PPXXY& = InStr(TPSP$, "00000030")
        If PPXXY& > 0 Then
           Mid$(TPSP$, PPXXY&, 8) = "00000029"
           GoTo 20
        End If
21      PPXXY& = InStr(TPSP$, "00000031")
        If PPXXY& > 0 Then
           Mid$(TPSP$, PPXXY&, 8) = "00000029"
           GoTo 21
        End If
22      PPXXY& = InStr(TPSP$, "00000032")
        If PPXXY& > 0 Then
           Mid$(TPSP$, PPXXY&, 8) = "00000029"
           GoTo 22
        End If
    End If
    '
    SPOOLSTRING$ = TPSP$
    For KKI% = 0 To 5
      EDITFORM.Picture1(KKI%).Cls
      EDITFORM.Picture1(KKI%).Top = 0
      EDITFORM.Picture1(KKI%).Refresh
      EDITFORM.Picture1(KKI%).Height = 7170
      EDITFORM.Picture1(KKI%).Width = 11700
      EDITFORM.Picture1(KKI%).DrawWidth = 3
      EDITFORM.Picture1(KKI%).DrawStyle = 0
      EDITFORM.Picture1(KKI%).FillStyle = 1
    Next KKI%
    ALTUPAGINA = 7170
    ANCHOPAGI = 0
    TOPEPAGINA = 0
    PAGINACTIVA% = 0
    CAPAGINAS% = 0
    VUELTAREP% = 0
    '
10  PPXX% = InStr(TPSP$, Chr$(255))
    If Len(TPSP$) > 0 Then
      If PPXX% < 1 Then
         ' ESTO ES UN PARCHE PARA ARREGLAR EL PROBLEMA DEL 19/7/2011 DE QUE EN EL
         ' SERVIDOR GRABABA CHR$(63) EN VEX DE 255 CUANDO LA CONFIGURACION AVAN-
         ' ZADA DE IDIOMA ERA RUSO.
         If VUELTAREP% = 0 Then
            For KKL& = 1 To Len(TPSP$)
               If Asc(Mid$(TPSP$, KKL&, 1)) = 63 Then
                  'Mid$(TPSP$, KKL&, 1) = Chr$(255)
                  TPSP$ = REPLACAR$(TPSP$, Chr$(63), Chr$(255))
               End If
            Next KKL&
            VUELTAREP% = 1
            'SPOOLSTRING$ = TPSP$
            GoTo 10
         End If
       End If
    End If
    ' PPXX% = InStr(TPSP$, Chr$(63))   ' OJO !!!!
    If PPXX% > 0 Then
      COPRIAT$ = Left$(TPSP$, PPXX% - 1)
      TPSP$ = Mid$(TPSP$, PPXX% + 1)
      Call TRACE(20, FIN& - Len(TPSP$), FIN&)
      '
      If Len(COPRIAT$) > 0 Then
        COPRI% = Asc(Left$(COPRIAT$, 1))
        ATRIP$ = Mid$(COPRIAT$, 2)
        Call SHOWCOMMAND(COPRI%, ATRIP$)
      End If
      GoTo 10
      '
    End If
    '
    Screen.MousePointer = 1
    PAGINACTIVA% = 0
    For KKI% = 0 To 5
       EDITFORM.Picture1(KKI%).Visible = False
    Next KKI%
    EDITFORM.Picture1(PAGINACTIVA%).Visible = True
    EDITFORM.Picture1(PAGINACTIVA%).Refresh
    EDITFORM.VScroll1.Min = 0
    EDITFORM.VScroll1.Value = 0
    EDITFORM.VScroll1.Max = 1
    '
    'INTERVENCION MUESTRA MENSAJE POR NO IMPRIMIR
    If TRIM$(MENSAJEDOC$) <> "" Then
      EDITFORM.Label3.Visible = True
      EDITFORM.Frame2.Left = 120
      EDITFORM.Label3 = MENSAJEDOC$
    Else
      EDITFORM.Label3.Visible = False
      EDITFORM.Frame2.Left = 3120
      EDITFORM.Label3 = ""
    End If
    '
    If NoPuedeImprimir% > 0 Then
       EDITFORM.Command4.Enabled = False
       EDITFORM.Command12.Enabled = False
    Else
       EDITFORM.Command4.Enabled = True
       EDITFORM.Command12.Enabled = True
    End If
    'FIN DE INTERVENCION MUESTRA MENSAJE POR NO IMPRIMIR
    '
'    If ALTUPAGINA > EDITFORM.Frame1.Height Then
'       EDITFORM.VScroll1.Max = (ALTUPAGINA - EDITFORM.Frame1.Height) / 100
'    End If

    ' AJUSTA ALTURA DEL FORMULARIO EDITFORM Y LA IMAGEN DEL COMPROBANTE
    EDITFORM.Height = 8670  ' ALTO POR DEFAULT
    AlturaDisponible = Screen.Height
    AlturaDisponible = AlturaDisponible * 0.9 ' UTILIZO EL 90% DE DISPONIBILIDAD
    ' VERIFICA SI EL ALTO SEGUN LA RESOLUCION ES MAYOR A LA ALTURA POR DEFAULT DEL FORMULARIO
    If AlturaDisponible > EDITFORM.Height Then
       ' AJUSTA EL ALTO DEL FORMULARIO DEPENDIENDO DE LA SUMA DE LAS PARTES DEL FORMULARIO EDITFORM: ENCABEZADO + IMAGEN (PICTURE) + SCROLL HORIZONTAL
       If (EDITFORM.Frame1.Top + ALTUPAGINA + EDITFORM.HScroll1.Height + 400) < AlturaDisponible Then
            EDITFORM.Height = EDITFORM.Frame1.Top + ALTUPAGINA + EDITFORM.HScroll1.Height + 400
            EDITFORM.Frame1.Height = ALTUPAGINA
            EDITFORM.Picture1(0).Height = ALTUPAGINA
            EDITFORM.Picture1(1).Height = ALTUPAGINA
            EDITFORM.Picture1(2).Height = ALTUPAGINA
            EDITFORM.Picture1(3).Height = ALTUPAGINA
            EDITFORM.Picture1(4).Height = ALTUPAGINA
            EDITFORM.Picture1(5).Height = ALTUPAGINA
            EDITFORM.VScroll1.Height = ALTUPAGINA
            EDITFORM.HScroll1.Top = EDITFORM.VScroll1.Top + ALTUPAGINA
       Else ' SI LA SUMA ANTERIOR NO CABE EN LA ALTURA DISPONIBLE, AJUSTA TODO A ESTA ULTIMA. A ESTA ALTURA, LA DISPONIBLE ES MAYOR A LA POR DEFECTO DE EDITFORM
            EDITFORM.Height = AlturaDisponible + 400
            EDITFORM.Frame1.Height = AlturaDisponible - (EDITFORM.Frame1.Top + EDITFORM.HScroll1.Height + 100)
            EDITFORM.Picture1(0).Height = ALTUPAGINA
            EDITFORM.Picture1(1).Height = ALTUPAGINA
            EDITFORM.Picture1(2).Height = ALTUPAGINA
            EDITFORM.Picture1(3).Height = ALTUPAGINA
            EDITFORM.Picture1(4).Height = ALTUPAGINA
            EDITFORM.Picture1(5).Height = ALTUPAGINA
            EDITFORM.VScroll1.Height = EDITFORM.Frame1.Height
            EDITFORM.HScroll1.Top = EDITFORM.VScroll1.Top + EDITFORM.Frame1.Height
       End If
       ' AJUSTE EL LARGO DEL SCROLL VERTICAL PARA PODER VISUALIZAR LO QUE NO SE LLEGA A VER DE LA IMAGEN (PICTURE)
       EDITFORM.VScroll1.Max = (ALTUPAGINA - EDITFORM.Frame1.Height) / 100
    End If
    
    ' AJUSTA ANCHO DEL FORMULARIO EDITFORM Y LA IMAGEN DEL COMPROBANTE
    EDITFORM.Width = 12000
    AnchoDisponible = Screen.Width '/ Screen.TwipsPerPixelX
    AnchoDisponible = AnchoDisponible * 0.8 ' UTILIZO EL 80% DE DISPONIBLE
    If AnchoDisponible > EDITFORM.Width Then
        ' AJUSTA EL ANCHO DEPENDIENDO SI EL ANCHO DE LA IMAGEN CABE EN EL ANCHO DISPONIBLE.
        ' LA IMAGEN TIENE CASI EL MISMO ANCHO QUE EL FORMULARIO, POR ESO ES QUE SE TOMA EL MISMO PARA COMPARA CON EL DISPONIBLE
        If ANCHOPAGI < AnchoDisponible Then
            If ANCHOPAGI > EDITFORM.Width Then ' AJUSTA SOLO SI
                EDITFORM.Width = ANCHOPAGI + (EDITFORM.VScroll1.Width + 200)
                EDITFORM.Frame1.Width = ANCHOPAGI
                EDITFORM.Picture1(0).Width = ANCHOPAGI
                EDITFORM.Picture1(1).Width = ANCHOPAGI
                EDITFORM.Picture1(2).Width = ANCHOPAGI
                EDITFORM.Picture1(3).Width = ANCHOPAGI
                EDITFORM.Picture1(4).Width = ANCHOPAGI
                EDITFORM.Picture1(5).Width = ANCHOPAGI
                EDITFORM.VScroll1.Left = ANCHOPAGI
                EDITFORM.HScroll1.Width = ANCHOPAGI
            End If
        Else ' EL ANCHO DE LA IMAGEN ES MAYOR QUE EL DISPONIBLE
            EDITFORM.Width = AnchoDisponible
            EDITFORM.Frame1.Width = AnchoDisponible - (EDITFORM.VScroll1.Width + 200)
            EDITFORM.Picture1(0).Width = EDITFORM.Frame1.Width
            EDITFORM.Picture1(1).Width = EDITFORM.Frame1.Width
            EDITFORM.Picture1(2).Width = EDITFORM.Frame1.Width
            EDITFORM.Picture1(3).Width = EDITFORM.Frame1.Width
            EDITFORM.Picture1(4).Width = EDITFORM.Frame1.Width
            EDITFORM.Picture1(5).Width = EDITFORM.Frame1.Width
            EDITFORM.VScroll1.Left = EDITFORM.Frame1.Width
            EDITFORM.HScroll1.Width = EDITFORM.Frame1.Width
        End If
        EDITFORM.HScroll1.Max = (-1) * (ANCHOPAGI - EDITFORM.Frame1.Width) / 100
    End If
    
    EDITFORM.Command3.Enabled = False
'    EDITFORM.Command1.Enabled = False
    '
    On Error Resume Next
    If UCase$(MODOX$) = "PRINT" Then
        Call EDITFORM.Command4_Click
      Else
        EDITFORM.Show 1
    End If
    On Error GoTo 0

99  End Sub


Sub SAVESPOOL(DESCRIDOC$, OKX%, Optional RET_DOCUNU&)
   '
   Static CACOMM%, PRICOMM$(128), COMMINDEX$
   OKX% = 0
   '
   If CACOMM% < 1 Then
      PRICOMM$(1) = "ORIENTATION"
      PRICOMM$(2) = "NEWPAGE"
      PRICOMM$(3) = "ENDDOC"
      PRICOMM$(4) = "HEIGHT"
      PRICOMM$(5) = "WIDTH"
      PRICOMM$(6) = "SCALEMODE"
      PRICOMM$(7) = "FORECOLOR"
      PRICOMM$(8) = "DRAWMODE"
      PRICOMM$(9) = "DRAWWIDTH"
      PRICOMM$(10) = "DRAWSTYLE"
      PRICOMM$(11) = "FILLSTYLE"
      PRICOMM$(12) = "LINE"
      PRICOMM$(13) = "FONTNAME"
      PRICOMM$(14) = "FONTITALIC"
      PRICOMM$(15) = "FONTSIZE"
      PRICOMM$(16) = "FONTBOLD"
      PRICOMM$(17) = "CURRENTX"
      PRICOMM$(18) = "CURRENTY"
      PRICOMM$(19) = "PRINT"
      PRICOMM$(20) = "FONTUNDERLINE"
      PRICOMM$(21) = "PAINTPICTURE"
      CACOMM% = 21
      COMMINDEX$ = ""
      For KKII% = 1 To CACOMM%
        COMMINDEX$ = COMMINDEX$ + AJUSTI$(PRICOMM$(KKII%), 16)
      Next KKII%
   End If
   '
   'If FORMALTER.PSPOOL.ListCount < 1 Then
'MsgBox "Entré en SAVESPOOL con CAREVEC = " & CAREVEC&
   If CAREVEC& < 1 Then
      Exit Sub
   End If
   '
'10 FIN& = FORMALTER.PSPOOL.ListCount
10 FIN& = CAREVEC&
   TPSP$ = String$(4096, 0): PINSE& = 1
   For KKIL& = 1 To FIN&
      Call TRACE(20, KKIL&, FIN&)
      'TTTT$ = TRIM$(FORMALTER.PSPOOL.List(KKIL& - 1))
      TTTT$ = TRIM$(VECPRICO$(KKIL&))
      PCOMM$ = TRIM$(Left$(TTTT$, 16))
      PATRI$ = Mid$(TTTT$, 17)
      KKKK% = InStr(COMMINDEX$, PCOMM$)
      If (KKKK% Mod 16) = 1 Then
        KKJJ% = (KKKK% + 15) / 16
        TPSQ$ = Chr$(KKJJ%) + PATRI$ + Chr$(255)
        If PINSE& + Len(TPSQ$) > Len(TPSP$) Then
          TPSP$ = TPSP$ + String$(4096, 0)
        End If
        Mid$(TPSP$, PINSE&, Len(TPSQ$)) = TPSQ$
        PINSE& = PINSE& + Len(TPSQ$)
        '+ Chr$(KKJJ%) + PATRI$ + Chr$(255)
        GoTo 19
      End If
      For KKJJ% = 1 To CACOMM%
        If PRICOMM$(KKJJ%) = PCOMM$ Then
           TPSQ$ = Chr$(KKJJ%) + PATRI$ + Chr$(255)
           If PINSE& + Len(TPSQ$) > Len(TPSP$) Then
             TPSP$ = TPSP$ + String$(4096, 0)
           End If
           Mid$(TPSP$, PINSE&, Len(TPSQ$)) = TPSQ$
           PINSE& = PINSE& + Len(TPSQ$)
           'TPSP$ = TPSP$ + Chr$(KKJJ%) + PATRI$ + Chr$(255)
           'If CANCELPRINT% < 1 Then Call PRINTCOMMAND(KKJJ%, PATRI$)
           GoTo 19
        End If
      Next KKJJ%
19 Next KKIL&
   TPSP$ = Left$(TPSP$, PINSE& - 1)
   '
20 Call BLOQARCH("PDOCLST")
   If Err Then
21    On Error GoTo 0
      Call COMUNI("Error al Bloquear Tablas'PDOCLST'.\  \Espere un Instante y Pulse 'Aceptar'\para Reanudar Ejecución.")
      GoTo 20
   End If
   On Error GoTo 0
   '
   UDOCU& = 0
   UREGI& = ULTREG&("PDOCLST")
   If UREGI& > 0 Then
      UDOCU& = CVS(Left$(REGLEIDO$("PDOCLST", UREGI&), 4))
   End If
   PDLS$ = String$(256, 0)
   DOCUNU& = 1 + UDOCU&
   HOII% = HOY(HOS$)
   HORX% = Int(HORANUM(Time$))
   Call REPLA(PDLS$, MKS$(DOCUNU&), 1, 4)
   Call REPLA(PDLS$, DESCRIDOC$, 5, 60)
   Call REPLA(PDLS$, MKI$(HOII%), 65, 2)
   Call REPLA(PDLS$, MKI$(HORX%), 67, 2)
   Call REPLA(PDLS$, USERNAME$, 69, 30)
   Call REPLA(PDLS$, NOMTER$, 99, 30)
     On Error Resume Next
   Call REPLA(PDLS$, Printer.DeviceName, 129, 128)
     On Error GoTo 0
   Call REGAPP("PDOCLST", PDLS$)
   Call CIERRARCH("PDOCLST")
   '
   Call ABRECONEXION
   Call CREACAMPO("", "PDOC_TBL", "NUMEDOC", 4, 0, 1)
   Call CREACAMPO("", "PDOC_TBL", "Referencia", 10, 128)
   Call CREACAMPO("", "PDOC_TBL", "FechayHora", 8, 0)
   Call CREACAMPO("", "PDOC_TBL", "Usuario", 10, 64)
   Call CREACAMPO("", "PDOC_TBL", "Terminal", 10, 64)
   Call CREACAMPO("", "PDOC_TBL", "Impresora", 10, 64)
   Call CREACAMPO("", "PDOC_TBL", "Imagen", 11, 0)
   Call CREACAMPO("", "PDOC_TBL", "TIPODOCU", 10, 32, 1)
   Call CREACAMPO("", "PDOC_TBL", "NUMECOMPLE", 10, 32, 1)
   Call CREACAMPO("", "PDOC_TBL", "NUMEBUSQUE", 10, 16, 1)
   '
   KKJJ% = 0
   'CONDI$ = "NUMEDOC =" & DOCUNU&
   SQLX$ = "SELECT * FROM  PDOC_TBL"
   SQLX$ = SQLX$ + " WHERE NUMEDOC =" & DOCUNU&
   
   Dim RSPDSL As ADODB.Recordset
   Set RSPDSL = New ADODB.Recordset
26 On Error Resume Next
   RSPDSL.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
   If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
        On Error GoTo 0
        Call CapturaErrorOpen
        GoTo 26
   End If
   On Error GoTo 0
   '
   With RSPDSL
       .AddNew
       On Error Resume Next
       !CODIEMPR = CodiEmp%
       On Error GoTo 0
       !NUMEDOC = DOCUNU&
       !REFERENCIA = Left$(DESCRIDOC$, 128)
       !FECHAYHORA = Now
       !USUARIO = Left$(USERNAME$, 64)
       !TERMINAL = Left$(NOMTER$, 64)
       !IMPRESORA = Left$(Printer.DeviceName, 64)
       !IMAGEN = TPSP$
       !TIPODOCU = Left$(TRIM$(TIPODOC$), 32)
       !NUMECOMPLE = NUMEDOC$
           NUBUSQ$ = NUBUSQUE$(NUMEDOC$, TIPODOC$)
       !NUMEBUSQUE = Left$(NUBUSQ$, 16)
       .Update
   End With
   RSPDSL.Close
   Set RSPDSL = Nothing
   RET_DOCUNU& = DOCUNU& 'VALOR DE RETORNO

   
'   Call CreaRegistro("PDOC_TBL", "NUMEDOC", DOCUNU&, "NOMESS")
'   Call ACTUREGISTRO("PDOC_TBL", "REFERENCIA", CONDI$, DESCRIDOC$, REGAF&, "NOMESS")
'   Call ACTUREGISTRO("PDOC_TBL", "FECHAYHORA", CONDI$, Now, REGAF&, "NOMESS")
'   Call ACTUREGISTRO("PDOC_TBL", "USUARIO", CONDI$, USERNAME$, REGAF&, "NOMESS")
'   Call ACTUREGISTRO("PDOC_TBL", "TERMINAL", CONDI$, NOMTER$, REGAF&, "NOMESS")
'   Call ACTUREGISTRO("PDOC_TBL", "IMPRESORA", CONDI$, Printer.DeviceName, REGAF&, "NOMESS")
'   Call ACTUREGISTRO("PDOC_TBL", "IMAGEN", CONDI$, tpsp$, REGAF&, "NOMESS")
'   '
'   Call ACTUREGISTRO("PDOC_TBL", "TIPODOCU", CONDI$, TRIM$(TIPODOC$), REGAF&, "NOMESS")
'   Call ACTUREGISTRO("PDOC_TBL", "NUMECOMPLE", CONDI$, NUMEDOC$, REGAF&, "NOMESS")
'   NUBUSQ$ = NUBUSQUE$(NUMEDOC$, TIPODOC$)
'   Call ACTUREGISTRO("PDOC_TBL", "NUMEBUSQUE", CONDI$, NUBUSQ$, REGAF&, "NOMESS")
   '
   Call CIERRARCH("PIMGLST")
   Call CIERRARCH("PIMGSPL")
   '
30 On Error Resume Next
   Printer.PrintQuality = (-1)
   MATRIXPUN% = 0
   If InStr(Printer.DeviceName, "810") > 0 Then MATRIXPUN% = 1
   If InStr(Printer.DeviceName, "286") > 0 Then MATRIXPUN% = 1
   If InStr(UCase$(Printer.DeviceName), "APEX") > 0 Then MATRIXPUN% = 1
   If InStr(UCase$(Printer.DeviceName), "FACTURA") > 0 Then MATRIXPUN% = 1
   If InStr(UCase$(Printer.DeviceName), "REMITO") > 0 Then MATRIXPUN% = 1
   'If InStr(Printer.DriverName, "EPSON9") > 0 Then MATRIXPUN% = 1
   On Error GoTo 0
   If MATRIXPUN% = 1 Then
      FORMALTER.List1.Clear
   End If
   '
   If CANCELPRINT% < 1 Then
     If HOJASUEL% = 1 Then
       Call COMUNI("Coloque Formulario Correspondiente a\" + TIPODOC$ + " " + NUMEDOC$ + "\y Pulse Botón 'Aceptar Mensaje'.")
     End If
   End If
   '
   'ASIGNO LA RUTA A UNA VARIABLE POR QUE CUANDO PRESENTA LA VENTANA PARA EL PDF SE PIERDE LA RUTA TOMA LA DE C:\FLEXOFT
   RUXA$ = CurDir  'MAS ABAJO LA ASIGNO VER COMO
   '
   If SENDBYMAIL$ <> "" Then
   
     ' *** LEE IMPRESORA PREDETERMINADA PARA LUEGO REGRESAR A ELLA
     IMPREDET$ = Printer.DeviceName
     Printer.TrackDefault = False
     
     ' *** NOMBRE DEL ARCHIVO PDF
     ARANEX$ = DEDOCOR$
     If ARANEX$ = "" Then ARANEX$ = "PDF"
     On Error Resume Next
     Kill LUCOM$ + "" + ARANEX$ + ".PDF"
     On Error GoTo 0
     
     ' *** DETECTA IMPRESORA VIRTUAL INSTALADA
     Dim noStart As Boolean
     ModoBullzip% = 0
     ModoPdfCreator% = 0
     For Each PRX1 In Printers
           ImpresoraNombre = PRX1.DeviceName
           
           If InStr(UCase$(ImpresoraNombre), "BIOPDF") > 0 Or InStr(UCase$(PRX1.DeviceName), "BULLZIP") > 0 Then
                If InStr(UCase$(ImpresoraNombre), "PDFCREATOR") > 0 Then
                End If
                
                If InStr(UCase$(ImpresoraNombre), "BULLZIP") > 0 Then ImpresoraNombre = "BULLZIP"
                If InStr(UCase$(ImpresoraNombre), "BIOPDF") > 0 Then ImpresoraNombre = "BIOPDF"
                SETTINGS_PROGID = ImpresoraNombre & ".PDFSettings"
                UTIL_PROGID = ImpresoraNombre & ".PDFUtil"
                Dim prtidx As Integer
                Dim sPrinterName As String
                Dim settings As Object
                Dim util As Object
                
                Set util = CreateObject(UTIL_PROGID)
                sPrinterName = util.defaultprintername
                
                Rem -- Configure the PDF print job
                Set settings = CreateObject(SETTINGS_PROGID)
                settings.printerName = sPrinterName
                settings.SetValue "Output", LUCOM$ + ARANEX$ & ".PDF"
                settings.SetValue "ConfirmOverwrite", "no"
                settings.SetValue "ShowSaveAS", "never"
                settings.SetValue "ShowSettings", "never"
                settings.SetValue "ShowPDF", "no"
                settings.SetValue "RememberLastFileName", "no"
                settings.SetValue "RememberLastFolderName", "no"
                settings.WriteSettings True
                Rem -- Find the index of the printer
                prtidx = PrinterIndex(sPrinterName)
                    
                Rem -- Set the current printer
                Set Printer = Printers(prtidx)
                ModoBullzip% = 1
                Exit For
             End If
     Next
    
     If ModoBullzip% = 1 Then GoTo 35
     
     ModoPdfCreator% = 0
     For Each PRX1 In Printers
         ImpresoraNombre = PRX1.DeviceName
         
         If InStr(UCase$(ImpresoraNombre), "PDFCREATOR") > 0 Then
              Set PDFCreator1 = CreateObject("PDFCREATOR.clsPDFCreator")
              Dim opt As Object
              
               With PDFCreator1
                .cVisible = True
                If .cStart("/NoProcessingAtStartup") = False Then
                  If .cStart("/NoProcessingAtStartup", True) = False Then
                   Exit Sub
                  End If
                  .cVisible = True
                End If
                Set opt = .cOptions
                .cClearCache
                noStart = False
               End With
              
               With opt
                opt.AutosaveDirectory = LUCOM$
                opt.AutosaveFilename = ARANEX$ & ".PDF"
                opt.UseAutosave = 1
                opt.UseAutosaveDirectory = 1
                opt.AutosaveFormat = 0 ' PDF
               End With
               Set PDFCreator1.cOptions = opt
               Set Printer = Printers(PrinterIndex("PDFCreator"))
                
               ModoPdfCreator% = 1
               Exit For
         End If
     Next
     
   End If
   '
35 On Error Resume Next

   Printer.EndDoc   ' INICIALIZA LA IMPRESORA
   Printer.ScaleMode = 1

   If Err Then
     Resume
     On Error GoTo 0
     If COMALTER%("Error de Comunicacion con la Impresora\\Resolver y Reintentar\Cancelar Impresion") <> 2 Then
       GoTo 35
     End If
     GoTo 40
   End If
   '
   FUERADEAREA% = 0
   FUENODISPO% = 0
   For KKIIL& = 1 To CAREVEC&
      Call TRACE(20, KKIIL&, CAREVEC&)
      TTTT$ = TRIM$(VECPRICO$(KKIIL&))
      PCOMM$ = TRIM$(Left$(TTTT$, 16))
      PATRI$ = Mid$(TTTT$, 17)
      For KKJJ% = 1 To CACOMM%
        If PRICOMM$(KKJJ%) = PCOMM$ Then
           If CANCELPRINT% < 1 Then Call PRINTCOMMAND(KKJJ%, PATRI$)
               TTXX$ = TTTT$ + Chr$(10) + Chr$(13)
               TTXX$ = TTXX$ + "SCALEMODE: " & Printer.ScaleMode & Chr$(10) & Chr$(13)
               TTXX$ = TTXX$ + "SCALEWIDTH: " & Printer.ScaleWidth & Chr$(10) & Chr$(13)
               TTXX$ = TTXX$ + "SCALELEFT: " & Printer.ScaleLeft & Chr$(10) & Chr$(13)
               TTXX$ = TTXX$ + "SCALEHEIGHT: " & Printer.ScaleHeight & Chr$(10) & Chr$(13)
               TTXX$ = TTXX$ + "SCALETOP: " & Printer.ScaleTop & Chr$(10) & Chr$(13)
               TTXX$ = TTXX$ + "PAPERSIZE: " & Printer.PaperSize
           GoTo 39
        End If
      Next KKJJ%
39 Next KKIIL&
   '
   If FUERADEAREA% > 0 Then
      Call MENSERR(24, "MENSAJE CRÍTICO: Se Han Encontrado Leyendas\Fuera de Area de Impresión Horizontal o Vertical.\   \FLEXOFT ha Suprimido estas Leyendas.\  \Se Recomienda Revisar y Corregir el Tamaño de Papel\Configurado en Esta Estación de Trabajo\para la Impresora " + REPLACAR$(Printer.DeviceName, "\", "/") + ".")
   End If
   FUERADEAREA% = 0
   '
   If FUENODISPO% > 0 Then
      Call MENSERR(24, "MENSAJE CRÍTICO: El Formulario Electrónico Actual\Utiliza Fuentes que no Están Disponibles\en la Impresora " + REPLACAR$(Printer.DeviceName, "\", "/") + ".\   \Se Recomienda Revisar el Formulario Electrónico\o Utilizar Otra Impresora de la Red.")
   End If
   FUENODISPO% = 0
   '
40 ReDim VECPRICO$(0)
   CAREVEC& = 0
   Call CLOSEFORMPRIN("")
   '
   If SENDBYMAIL$ <> "" Then
     Call SETPRINPORT("RESTORE")
      ' *** VERIFICA SI EL ARCHIVO YA FUE CREADO
     If ModoBullzip% = 1 Then
        Dim runonce As String
41      runonce = settings.GetSettingsFilePath(True)
        VUELTA% = 0
        While Dir(runonce, vbNormal) <> ""
            Call T_Espera(0.1)
            VUELTA% = VUELTA% + 1
            If VUELTA% > 15 Then
                TX1$ = "Generando Archivo PDF "
                OPX% = COMALTER%(TX1$ & " \\Continuar\Cancelar")
                If OPX% = 2 Then
                    Call MENSERR(24, "Imposible Enviar por Mail.\  \No se ha Generado Correctamente\el Archivo Adjunto en Formato PDF.")
                    Exit Sub
                Else
                    GoTo 41
                End If
            End If
        Wend
        ' SE AGREGA ESTA DEMORA PORQUE EXISTE UN INTERVALO DE TIEMPO ENTRE QUE EL BULLZIP ELIMINA EL ARCHIVO .INI
        ' CON LA CONFIGURACION QUE GENERA EL PDF Y LA GRABACION DEL PDF EN LA RUTA ASIGNADA
        TINI = Timer: While Timer < TINI + 1: Wend
        Call T_Espera(2)
        EDITFORM.Label1.Caption = ARANEX$
        GoTo 60
        '
     End If
     '
      ' *** VERIFICA SI EL ARCHIVO YA FUE CREADO
     If ModoPdfCreator% = 1 Then
        PDFCreator1.cPrinterStop = False
        Sleep 1000
       
        If noStart = False Then
            PDFCreator1.cClose
            While PDFCreator1.cProgramIsRunning
                DoEvents
                Sleep 1000
            Wend
        End If
        Set PDFCreator1 = Nothing
        Set opt = Nothing
        
46      While Dir(LUCOM$ & ARANEX$ & ".PDF", vbNormal) = ""
            Call T_Espera(0.1)
            VUELTA% = VUELTA% + 1
            If VUELTA% > 15 Then
                TX1$ = "Generando Archivo PDF "
                OPX% = COMALTER%(TX1$ & " \\Continuar\Cancelar")
                If OPX% = 2 Then
                    Call MENSERR(24, "Imposible Enviar por Mail.\  \No se ha Generado Correctamente\el Archivo Adjunto en Formato PDF.")
                    Exit Sub
                Else
                    VUELTA% = 0
                    GoTo 46
                End If
            End If
        Wend
        GoTo 60
     End If
     
49   TINI = Timer
50   If EXISTE%(LUCOM$ + "" + TRIM$(EDITFORM.Label1) + ".PDF") < 1 Then
       If Timer < TINI + 30 Then
         GoTo 50
       End If
       '
       TX1$ = "Generando Archivo PDF "
       OPX% = COMALTER%(TX1$ & " \\Continuar\Cancelar")
       If OPX% = 2 Then
           Call MENSERR(24, "Imposible Enviar por Mail.\  \No se ha Generado Correctamente\el Archivo Adjunto en Formato PDF.")
           Exit Sub
       End If
       If EXISTE%(LUCOM$ + "" + TRIM$(EDITFORM.Label1) + ".PDF") < 1 Then GoTo 49
     End If
     'ANTES DEL ENVIAMAIL LE ASIGNO LA RUTA POR  QUE QUEDO EN C:\FLEXOFT EN VEZ DE X:\GABAL\FLEXOFT\WINDOWS
     'X:\GABAL\FLEXOFT\WINDOWS (EJEMPLO DE  RUTA PARA VER COMO SE ASIGNA PARTICIONANDO EL TEXTO)
60   ChDrive (Left$(RUXA$, 1)) 'ASIGNO LA UNIDAD DE RED PRIMER POSICION DE LA RUTA (X)
     ChDir Mid$(RUXA$, 4)      'LE ASIGNO LA RUTA DESPUES DE LA BARRA  (FLEXOFT\WINDOWS)
     'FIN DE ASIGNACION DE RUTA PERDIDA
     Call T_Espera(2)
     Call ENVIAPORMAIL
     '
   End If
   OKX% = 1
   '
End Sub

Function EXISTEIMPRESORAPDF%()
    EXISTEIMPRESORAPDF% = -1
    'VALIDO SI EXISTE LA IMPRESORA PDF
    For Each PRX1 In Printers
       ImpresoraNombre = PRX1.DeviceName
       If InStr(UCase$(ImpresoraNombre), "PDFWRITER") > 0 Or InStr(UCase$(ImpresoraNombre), "BIOPDF") > 0 Or InStr(UCase$(PRX1.DeviceName), "BULLZIP") > 0 Then
         EXISTEIMPRESORAPDF% = 1
         Exit For
       End If
    Next

End Function
Function PrinterIndex(ByVal printerName As String) As Integer
    Dim i As Integer
    
    For i = 0 To Printers.Count - 1
        If LCase(Printers(i).DeviceName) Like LCase(printerName) Then
            PrinterIndex = i
            Exit Function
        End If
    Next
    PrinterIndex = -1
End Function

Function NUBUSQUE$(NUCOMP$, TIPOFA$)
               
            'DEVUELVE EL NUMERO DE BUSQUEDA SEGUN EL NUMERO COMPLETO
 
            NUBUSX$ = NUCOMP$
            TIPOX$ = UCase$(TIPOFA$)
            '
            'CASOS: Remito Nro.0001-00040261  **  Remito Nro.RT-X-0001-00000008
            POSREMI% = 6: If Left$(NUCOMP$, 5) = "RT-X-" Then POSREMI% = 11
 
            If Left$(TIPOX$, Len("REMITO")) = "REMITO" Then NUBUSX$ = TRIM$(Mid$(NUCOMP$, POSREMI%))
            If Left$(TIPOX$, Len("FACTURA")) = "FACTURA" Then NUBUSX$ = TRIM$(Mid$(NUCOMP$, 6))
            '
            If TIPOX$ = "DEPOSITO BANCARIO" Then NUBUSX$ = TRIM$(Mid$(NUCOMP$, 4))
            If TIPOX$ = "NOTA DE CREDITO" Then NUBUSX$ = TRIM$(Mid$(NUCOMP$, 6))
            If TIPOX$ = "NOTA DE DEBITO" Then NUBUSX$ = TRIM$(Mid$(NUCOMP$, 6))
            If TIPOX$ = "ORDEN DE FABRICACIÓN" Then NUBUSX$ = TRIM$(Mid$(NUCOMP$, 4))
            If TIPOX$ = "REMITO DE DEVOLUCIÓN" Then NUBUSX$ = TRIM$(Mid$(NUCOMP$, 4))
            If TIPOX$ = "REMITO DE TRASLADO" Then NUBUSX$ = TRIM$(Mid$(NUCOMP$, 4))
            If TIPOX$ = "VALE DE CONSUMO DE MATERIALES" Then NUBUSX$ = TRIM$(Mid$(NUCOMP$, 4))
            If TIPOX$ = "DESPACHO DE MATERIALES" Then NUBUSX$ = TRIM$(Mid$(NUCOMP$, 4))
            If TIPOX$ = "REM-REPA" Then NUBUSX$ = TRIM$(Mid$(NUCOMP$, 4))
            '
            NUBUSQUE$ = TRIM$(FORMATNUM$(Val(NUBUSX$), "F16.0"))
            '
 End Function

Function OCOMEXISTE%(NCLIEN, NORCOM$, Optional NROPED&)
   '
      'VALIDACION SI LA ORDEN DE COMPRA EXISTE PARA ESTE CLIENTE
      OCOMEXISTE% = 0
      If NCLIEN < 1 Or TRIM$(NORCOM$) = "" Then Exit Function
      '
      CONDI$ = "NROCLIE = " & NCLIEN & " And NROOCOM = '" & NORCOM$ & "'"
      CONDI$ = CONDI$ + "AND  Status <> 9 " 'QUE LA ORDEN DE COMPRA NO ESTE ANULADA
      '
      'SI SE PASA NUMERO DE PEDIDO SE AGREGA CONDI QUE SEA DIFRENTE A ESTE
      If XVALO(NROPED&) > 0 Then
         CONDI$ = CONDI$ + " And NROPED <> " & NROPED&
      End If
      '
      EXISTEOCOMPRA% = CantRegistros("PEDENCA", CONDI$)
      '
      If EXISTEOCOMPRA% > 0 Then
        OCOMEXISTE% = 1
      End If
      '
End Function

Function OCOMEXISTENEUQUEN%(NCLIEN, NORCOM$, Optional ORCOMPRAORIGINAL&)
      'EN ARCON EN SUCURSALES REALIZAN ORDEN DE COMPRA HACIA AVELLANEDA Y SE PASAN LOS DATOS DEL CLIENTE QUE HIZO EL PEDIDO Y LA ORDEN DE COMPRA
      'VALIDACION SI LA ORDEN DE COMPRA EXISTE PARA ESTE CLIENTE
      OCOMEXISTENEUQUEN% = 0
      If NCLIEN < 1 Or TRIM$(NORCOM$) = "" Then Exit Function
      '
      CONDI$ = "Nume_Cli = " & NCLIEN & " And OComCliente = '" & NORCOM$ & "'"
      CONDI$ = CONDI$ + " AND  Stat_Ocom <> 9 " 'QUE LA ORDEN DE COMPRA NO ESTE ANULADA
      If ORCOMPRAORIGINAL& > 0 Then CONDI$ = CONDI$ + " AND Nume_Ocom <> " & ORCOMPRAORIGINAL&
      '
      EXISTEOCOMPRA% = CantRegistros("OCOMENCA", CONDI$)
      '
      If EXISTEOCOMPRA% > 0 Then
        OCOMEXISTENEUQUEN% = 1
      End If
      '
End Function

Function Control$(PROCE$, CLAVE$)
    '
    Call ABRECONEXION
    If YACRECON% < 1 Then
       Call CREACAMPO("", "FLEXCRL", "PROBUS", 10, 32, 1)
       Call CREACAMPO("", "FLEXCRL", "CLABUS", 10, 32, 1)
       Call CREACAMPO("", "FLEXCRL", "VALCONTROL", 12, 0, 0)
       YACRECON% = 1
    End If
    '
    CONDI$ = "PROBUS = '" & PROCE$ & "' AND CLABUS = '" & CLAVE$ & "'"
    Control$ = VALOBADA("FLEXCRL", "VALCONTROL", CONDI$, "NOMESS")
    '
    If Control$ = "" Then
        If CantRegistros("FLEXCRL", CONDI$, "NOMESS") > 0 Then
           Exit Function             ' PARA POSIBILITAR QUE DEVUELVA BLANCOS
        End If
        '
        ' SI NO LO ENCUENTRA BUSCA EN LA TABLA DAT
        PROBUS$ = AJUSTI$(PROCE$, 8)
        If TRIM$(PROCE$) = "" Then PROBUS$ = ""
        CODBUS$ = AJUSTI$(CLAVE$, 8)
        '
        Control$ = ""
        
        RFF$ = RECFILT$("FLEXCRL", 2, CODBUS$)
        For U& = 1 To Len(RFF$) Step 4
           REBU& = CVS(Mid$(RFF$, U&, 4))
           XI$ = REGLEIDO$("FLEXCRL", REBU&)
           If PROBUS$ = "" Or PROBUS$ = Left$(XI$, 8) Then
              Control$ = TRIM$(Mid$(XI$, 17, 80))
              Exit For
           End If
        Next U&
    End If
    '
End Function
'
Sub GRACONTROL(PROCE$, CLAVE$, VALCON$)
    '
    Call ABRECONEXION
    If YACRECON% < 1 Then
       Call CREACAMPO("", "FLEXCRL", "PROBUS", 10, 32, 1)
       Call CREACAMPO("", "FLEXCRL", "CLABUS", 10, 32, 1)
       Call CREACAMPO("", "FLEXCRL", "VALCONTROL", 12, 0, 0)
       YACRECON% = 1
    End If
    '
    CONDI$ = "PROBUS = '" & PROCE$ & "' AND CLABUS = '" & CLAVE$ & "'"
    If CantRegistros("FLEXCRL", CONDI$, "NOMESS") < 1 Then
10     RNDX$ = RANDSTRING$(-8)
       If CantRegistros("FLEXCRL", "PROBUS = '" & RNDX$ & "'", "NOMESS") > 0 Then GoTo 10
       Call CreaRegistro("FLEXCRL", "PROBUS", RNDX$, "NOMESS")
       Call ACTUREGISTRO("FLEXCRL", "CLABUS", "PROBUS = '" & RNDX$ & "'", CLAVE$, RAFE&, "NOMESS")
       Call ACTUREGISTRO("FLEXCRL", "PROBUS", "PROBUS = '" & RNDX$ & "'", PROCE$, RAFE&, "NOMESS")
    End If
    Call ACTUREGISTRO("FLEXCRL", "VALCONTROL", CONDI$, VALCON$, RAFE&, "NOMESS")
    '
End Sub
'
Function DOCID&(Tipo$, NUBUSQUEDA$, Optional NUCOMPLE$, Optional VARCONTROL$)
    '
    TIPO1$ = UCase$(TRIM$(Tipo$))
    
    NUCOMPLE1$ = UCase$(TRIM$(NUCOMPLE$))
    NUBUSQUEDA1$ = UCase$(TRIM$(NUBUSQUEDA$))
    VARCONTROL1$ = TRIM$(VARCONTROL$)
    '
    '
    Call ABRECONEXION
    Call CREACAMPO("", "PDOC_TBL", "TIPODOCU", 10, 32, 1)
    Call CREACAMPO("", "PDOC_TBL", "NUMECOMPLE", 10, 32, 1)
    Call CREACAMPO("", "PDOC_TBL", "NUMEBUSQUE", 10, 16, 1)
    '
    CONDI$ = "TIPODOCU = '" & TIPO1$ & "'"
    CONDI$ = CONDI$ + " AND (RTRIM(LTRIM(NUMEBUSQUE)) = '" & NUBUSQUEDA1$ & "'"
    'NUMERO COMPLETO
    If TRIM$(NUCOMPLE1$) <> "" Then
       CONDI$ = CONDI$ + " AND RTRIM(LTRIM(NUMECOMPLE)) = '" & NUCOMPLE1$ & "'"
    End If
    CONDI$ = CONDI$ + ") "
    'VARIABLE DE CONTROL
    If VARCONTROL1$ <> "" Then
       'SI EL PARAMETRO VARCONTROL TIENE UNA FECHA ENTONCES BUSCA EN EL CAMPO FECHAYHORA DE PDOC_TBL
       FF1% = FECHANUM(VARCONTROL1$)
       If FF1% > 0 Then
         FECHA1$ = FETEXCOR1$(FF1%): FECHA2$ = FETEXCOR1$(FF1%, "HASTA")
         CONDI$ = CONDI$ + " AND FECHAYHORA >='" & FECHA1$ & "' AND FECHAYHORA <='" & FECHA2$ & "'"
       Else
         CONDI$ = CONDI$ + " AND REFERENCIA LIKE '%" & VARCONTROL$ & "%'"
       End If
    End If
    '
    
    'AGREGADO PARA VISUALIZAR DUPLICADOS
    If CantRegistros&("PDOC_TBL", CONDI$, "NOMESS") > 1 And TIPO1$ = "FACTURA" And VARCONTROL1$ = "" Then
       Call CARGALISEL("LISFACTX", "PDOC_TBL", "REFERENCIA/A64", CONDI$, "NOMESS")
       Call SELECHO("LISFACTX/SELECCION DE FACTURA")
       REFE$ = TRIM$(VALACT1$("LISFACTX"))
       If REFE$ <> "" Then
          CONDI$ = CONDI$ + " AND REFERENCIA LIKE '%" & REFE & "%'"
       End If
    End If
    '
    If CantRegistros&("PDOC_TBL", CONDI$, "NOMESS") > 1 And TIPO1$ = "RECIBO" Then
       Call CARGALISEL("LISRECI", "PDOC_TBL", "REFERENCIA/A64", CONDI$, "NOMESS")
       Call SELECHO("LISRECI/SELECCION DE Recibo")
       REFE$ = TRIM$(VALACT1$("LISRECI"))
       If REFE$ <> "" Then
          CONDI$ = CONDI$ + " AND REFERENCIA LIKE '%" & REFE & "%'"
       End If
    End If

    'FIN DE AGREGADO PARA VISUALIZAR DUPLICADOS
    '
    'SE AGREGA ESTA CONDICION PARA QUE BUSQUE DE ADELANTE HACIA ATRAS
    CONDI$ = CONDI$ + " ORDER BY NUMEDOC DESC"
    
    DOCID& = XVALO(VALOBADA("PDOC_TBL", "NUMEDOC", CONDI$, "NOMESS"))
    '
    If DOCID& > 0 Then Exit Function
    '
    'ESTO ES PARA QUE NO BUSQUE CUANDO SE TRATA DE FACTURA ELECTRONICA YA ESTOS
    'COMPROBANTES NO ESTAN REGISTRADOS EN EL PDOCSPL
    If InStr(TIPO1$, "FACTURA") > 0 Or InStr(TIPO1$, "NOTA DE DEBITO") > 0 Or InStr(TIPO1$, "NOTA DE CREDITO") > 0 Then
       If InStr(1, EMPREAC$, "FERVI") > 0 Then
            FE% = FECHANUM(VARCONTROL1$): FF% = FECHANUM("13/01/13")
            If FE% > FF% Then
                'VARCONTROL1$ = ""
                Exit Function
            End If
       Else
        If Control$("FACTURA", "SALIDA") = "FELECT" Then Exit Function
       End If
    End If
    '
    For UI& = ULTREG&("PDOCLST") To 1 Step -1
      XX$ = Left$(REGLEIDO$("PDOCLST", UI&), 64)
      If UI& = 268843 Then
      A = B
      End If
      If InStr(UCase$(XX$), TIPO1$) > 4 Then
        '
        'VALIDACION QUE SI ES PRESUPUESTO NO TRAIGA UN PRESUPUESTO DE REPARACION
        If TIPO1$ = "PRESUPUESTO" Then
           If InStr(UCase$(XX$), "PRESUPUESTO DE ") > 4 Then GoTo 90 'NEXT
        End If
        '
        PPXX% = InStr(XX$, "NRO.")
        If PPXX% > 0 Then
          XX$ = Left$(XX$, PPXX% - 1) + " " + Mid$(XX$, PPXX% + 4)
        End If
         '
        If TIPO1$ = "NOTA DE PEDIDO" Then 'ORDEN DE COMPRA
          While Len(NUBUSQUEDA1$) < 6: NUBUSQUEDA1$ = "0" + NUBUSQUEDA1$: Wend
          NUDOCU1$ = " " + NUBUSQUEDA1$ + " "
          NUDOCU2$ = " " + NUBUSQUEDA1$ + "-"

          'CASO DE ORDEN DE COMPRA BUSCA POR UN OR CON DOS VARIABLES
          If InStr(XX$, NUDOCU1$) > 4 Or InStr(XX$, NUDOCU2$) > 4 Then
            DOCID& = CVS(Left$(XX$, 4))
            Exit Function
          End If
        '
        'CASO FACTURA O NOTA DE ...
        ElseIf InStr(UCase$(TIPO1$), "FACTURA") > 0 Or InStr(UCase$(TIPO1$), "NOTA DE") > 0 Then
          FE% = FECHANUM(VARCONTROL1$)
          'FEMIX% = CVI(Mid$(XX$, 65, 2))
          If InStr(XX$, NUCOMPLE1$) > 4 Then
'            If InStr(XX$, TIPO1$) > 4 Then  SE MODIFICA PARA QUE COMPARE CON  MAYUSCULAS
            If InStr(UCase$(XX$), TIPO1$) > 4 Then
                'SE OMITE LA COMPARACION POR RAZON SOCIAL, ASUMIENDO QUE EXISTE
                'LA POSIBILIDAD DE QUE ESTA CAMBIE
                'If FEMIX% = FE% Then
                  DOCID& = CVS(Left$(XX$, 4))
                  Exit Function
                'End If
            End If
          End If
        '
        'RESTO
        Else
            ' se agregan los espacios para contemplar el caso que al buscar el numero 712
            ' no traiga el 5712 y el espacio al final para que si busco el 1 no me traiga el 1700
            TXT1$ = XX$
            If UCase(Tipo$) = "BOLETA DE RECEPCIÓN" Then
                NN% = InStr(5, XX$, "-")
                TXT1$ = Mid$(XX$, 5, NN% - 5)
            End If
            If InStr(TXT1$, " " & NUCOMPLE1$ & " ") > 4 Then
              DOCID& = CVS(Left$(XX$, 4))
              Exit Function
            End If
        End If
        
      End If
90  Next UI&
    '
End Function


'Function DOCID_RESPALDO&(Tipo$, NUBUSQUEDA$, Optional NUCOMPLE$, Optional VARCONTROL$)
'    '
'    TIPO1$ = UCase$(TRIM$(Tipo$))
'    NUCOMPLE1$ = UCase$(TRIM$(NUCOMPLE$))
'    NUBUSQUEDA1$ = UCase$(TRIM$(NUBUSQUEDA$))
'    VARCONTROL1$ = TRIM$(VARCONTROL$)
'    '
'    Call ABRECONEXION
'    Call CREACAMPO("", "PDOC_TBL", "TIPODOCU", 10, 32, 1)
'    Call CREACAMPO("", "PDOC_TBL", "NUMECOMPLE", 10, 32, 1)
'    Call CREACAMPO("", "PDOC_TBL", "NUMEBUSQUE", 10, 16, 1)
'    '
'    CONDI$ = "TIPODOCU = '" & TIPO1$ & "'"
'    CONDI$ = CONDI$ + " AND (NUMEBUSQUE = '" & NUBUSQUEDA1$ & "'"
'    'NUMERO COMPLETO
'    If TRIM$(NUCOMPLE1$) <> "" Then
'       CONDI$ = CONDI$ + " AND NUMECOMPLE = '" & NUCOMPLE1$ & "'"
'    End If
'    CONDI$ = CONDI$ + ") "
'    'VARIABLE DE CONTROL
'    If VARCONTROL1$ <> "" Then
'       'SI EL PARAMETRO VARCONTROL TIENE UNA FECHA ENTONCES BUSCA EN EL CAMPO FECHAYHORA DE PDOC_TBL
'       FF1% = FECHANUM(VARCONTROL1$)
'       If FF1% > 0 Then
'         FECHA1$ = FETEXCOR1$(FF1%): FECHA2$ = FETEXCOR1$(FF1%, "HASTA")
'         CONDI$ = CONDI$ + " AND FECHAYHORA >='" & FECHA1$ & "' AND FECHAYHORA <='" & FECHA2$ & "'"
'       Else
'         CONDI$ = CONDI$ + " AND REFERENCIA LIKE '%" & VARCONTROL1$ & "%'"
'       End If
'    End If
'    '
'
'    'AGREGADO PARA VISUALIZAR DUPLICADOS
'    If CantRegistros&("PDOC_TBL", CONDI$, "NOMESS") > 1 And TIPO1$ = "FACTURA" And VARCONTROL1$ = "" Then
'       Call CARGALISEL("LISFACTX", "PDOC_TBL", "REFERENCIA/A64", CONDI$, "NOMESS")
'       Call SELECHO("LISFACTX/SELECCION DE FACTURA")
'       REFE$ = TRIM$(VALACT1$("LISFACTX"))
'       If RESP$ <> "" Then
'          CONDI$ = CONDI$ + " AND REFERENCIA LIKE '%" & REFE & "%'"
'       End If
'    End If
'    'FIN DE AGREGADO PARA VISUALIZAR DUPLICADOS
'    '
'    'SE AGREGA ESTA CONDICION PARA QUE BUSQUE DE ADELANTE HACIA ATRAS
'    CONDI$ = CONDI$ + " ORDER BY NUMEDOC DESC"
'
'    DOCID& = XVALO(VALOBADA("PDOC_TBL", "NUMEDOC", CONDI$, "NOMESS"))
'    '
'    If DOCID& > 0 Then Exit Function
'    '
'    'ESTO ES PARA QUE NO BUSQUE CUANDO SE TRATA DE FACTURA ELECTRONICA YA ESTOS
'    'COMPROBANTES NO ESTAN REGISTRADOS EN EL PDOCSPL
'    If InStr(TIPO1$, "FACTURA") > 0 Or InStr(TIPO1$, "NOTA DE") > 0 Then
'        If CONTROL$("FACTURA", "SALIDA") = "FELECT" Then Exit Function
'    End If
'    '
'
'50  For UI& = ULTREG&("PDOCLST") To 1 Step -1
'      XX$ = Left$(REGLEIDO$("PDOCLST", UI&), 64)
'      If InStr(UCase$(XX$), TIPO1$) > 4 Then
'        '
'        'VALIDACION QUE SI ES PRESUPUESTO NO TRAIGA UN PRESUPUESTO DE REPARACION
'        If TIPO1$ = "PRESUPUESTO" Then
'           If InStr(UCase$(XX$), "PRESUPUESTO DE ") > 4 Then GoTo 90 'NEXT
'        End If
'        '
'        PPXX% = InStr(XX$, "NRO.")
'        If PPXX% > 0 Then
'          XX$ = Left$(XX$, PPXX% - 1) + " " + Mid$(XX$, PPXX% + 4)
'        End If
'         '
'        If TIPO1$ = "NOTA DE PEDIDO" Then 'ORDEN DE COMPRA
'          While Len(NUBUSQUEDA1$) < 6: NUBUSQUEDA1$ = "0" + NUBUSQUEDA1$: Wend
'          NUDOCU1$ = " " + NUBUSQUEDA1$ + " "
'          NUDOCU2$ = " " + NUBUSQUEDA1$ + "-"
'
'          'CASO DE ORDEN DE COMPRA BUSCA POR UN OR CON DOS VARIABLES
'          If InStr(XX$, NUDOCU1$) > 4 Or InStr(XX$, NUDOCU2$) > 4 Then
'            DOCID& = CVS(Left$(XX$, 4))
'            Exit Function
'          End If
'        '
'        'CASO FACTURA O NOTA DE ...
'        ElseIf InStr(UCase$(TIPO1$), "FACTURA") > 0 Or InStr(UCase$(TIPO1$), "NOTA DE") > 0 Then
'          FE% = FECHANUM(VARCONTROL1$)
'          FEMIX% = CVI(Mid$(XX$, 65, 2))
'          If InStr(XX$, NUCOMPLE1$) > 4 Then
''            If InStr(XX$, TIPO1$) > 4 Then  SE MODIFICA PARA QUE COMPARE CON  MAYUSCULAS
'            If InStr(UCase$(XX$), TIPO1$) > 4 Then
'                'SE OMITE LA COMPARACION POR RAZON SOCIAL, ASUMIENDO QUE EXISTE
'                'LA POSIBILIDAD DE QUE ESTA CAMBIE
'                If FEMIX% = FE% Then
'                  DOCID& = CVS(Left$(XX$, 4))
'                  Exit Function
'                End If
'            End If
'          End If
'        '
'        'RESTO
'        Else
'            ' se agregan los espacios para contemplar el caso que al buscar el numero 712
'            ' no traiga el 5712 y el espacio al final para que si busco el 1 no me traiga el 1700
'            TXT1$ = XX$
'            If UCase(Tipo$) = "BOLETA DE RECEPCIÓN" Then
'                NN% = InStr(5, XX$, "-")
'                TXT1$ = Mid$(XX$, 5, NN% - 5)
'            End If
'            If InStr(TXT1$, " " & NUCOMPLE1$ & " ") > 4 Then
'              DOCID& = CVS(Left$(XX$, 4))
'              Exit Function
'            End If
'        End If
'
'      End If
'90  Next UI&
    '
'End Function
Function FACTUFERVI&(Tipo$, NUCOMPLE$, FE%)
   
     TIPO1$ = UCase$(TRIM$(Tipo$))
     NUCOMPLE1$ = UCase$(TRIM$(NUCOMPLE$))
     If InStr(UCase$(TIPO1$), "FACTURA") > 0 Or InStr(UCase$(TIDOCUM$), "NOTA DE") > 0 Then
      NUCOMPLE1$ = TRIM$(NUCOMPLE1$)
      While Len(NUCOMPLE1$) < 8: NUCOMPLE1$ = "0" + NUCOMPLE1$: Wend
     End If

50   For UI& = ULTREG&("PDOCLST") To 1 Step -1
      XX$ = Left$(REGLEIDO$("PDOCLST", UI&), 64)
      If InStr(UCase$(XX$), TIPO1$) > 4 Then
        '
        'CASO FACTURA O NOTA DE ...
          FEMIX% = CVI(Mid$(XX$, 65, 2))
          If InStr(XX$, NUCOMPLE1$) > 4 Then
            If InStr(UCase$(XX$), TIPO1$) > 4 Then
                'SE OMITE LA COMPARACION POR RAZON SOCIAL, ASUMIENDO QUE EXISTE
                'LA POSIBILIDAD DE QUE ESTA CAMBIE
                If FEMIX% = FE% Then
                  FACTUFERVI& = CVS(Left$(XX$, 4))
                  Exit Function
                End If
            End If
        End If
      End If
90  Next UI&

End Function
Sub CONTACTO_MAIL(CODICX$, NCL)

    COCLI$ = TRIM$(CODICX$)
    FRMLISTMAIL.LICHESEL.Clear
    '
    Call CREATABLA_CONTACSCOM
    
    If CantRegistros&("CONTACSCOM", "CODICLI = '" & COCLI$ & "'", "NOMESS") > 0 Then
       'CARGO PRIMERO EL DEL MAESTRO
       MAILX$ = EMAILCLI$(XVALO(NCL))
       If InStr(MAILX$, "@") > 0 Then
           Call REPLA(TX$, MAILX$, 44, 64)
           FRMLISTMAIL.LICHESEL.AddItem TX$
       End If
       '
       'RECORRO LA TABLA CONTACSCOM Y CARGO LOS MAIL EN EL LIST
       SQLX$ = "SELECT * FROM CONTACSCOM WITH(NOLOCK)"
       SQLX$ = SQLX$ + " WHERE CODICLI = '" & COCLI$ & "'  ORDER BY MAIL_CTAC"
       Set MAITMP = New ADODB.Recordset
       Set MAITMP = FLEXCONN.Execute(FILTSQL$(SQLX$))
       With MAITMP
         Do While Not .EOF
           MAILX$ = SAFETEXT$(!MAIL_CTAC)
           If InStr(MAILX$, "@") > 0 Then
              nomb$ = SAFETEXT$(!NOME_CTAC)
              CARG$ = SAFETEXT$(!CARG_CTAC)
              Call REPLA(TX$, nomb$, 1, 20)
              Call REPLA(TX$, CARG$, 22, 20)
              Call REPLA(TX$, MAILX$, 44, 64)
              FRMLISTMAIL.LICHESEL.AddItem TX$
           End If
           .MoveNext
         Loop
       End With
       '
       'PRESENTO EL FRM DE CARGA DE MAIL
       If FRMLISTMAIL.LICHESEL.ListCount > 0 Then
         FRMLISTMAIL.Show 1
       End If
    End If
End Sub

Sub MUESTRASTSCREEN(NCLI, NUDOCUM$)
    '
    NCLI2 = NCLI
    NUFACTU$ = NUDOCUM$
    '
    'SE AGREGA POR QUE AHORA GUARDA EL CLIENTE ORIGINAL PARA CASOS DE GRUPO DIVISIONAL
    CONDI$ = "CODICOM_LAR = '" & NUFACTU$ & "'AND  NUCLIEN = " & NCLI
    NCLI1 = XVALO(VALOBADA("CAJABANC", "CLI_REAL", CONDI$, "NOMESS"))
    If NCLI1 > 0 Then NCLI2 = NCLI1
    '
    TIPOCOM$ = Mid$(NUFACTU$, 1, 2)
    '
    CONDI$ = "DOSECLAR='" & NUFACTU$ & "' AND NUME_CLIE= " & NCLI2 & " AND COD_INTE > 0"
    CANTREG& = XVALO(CantRegistros("MOVISTO", CONDI$, "NOMESS"))
    '
    TABLA1$ = "MOVISTO"
    CANTREGREPA& = 0
    If CANTREG& < 1 Then
        'VALIDO SI ES UNA REPARACION
        Condi1$ = "DOSECLAR='" & NUFACTU$ & "' AND NUME_CLIE= " & NCLI2 & ""
        CANTREGREPA& = XVALO(CantRegistros("MOVIREPA", Condi1$, "NOMESS"))
        If CANTREGREPA& > 0 Then TABLA1$ = "MOVIREPA"
    End If
    '
    'SI EL CAMPO DE Descri_Lar ES NULL O VACIO SE LE AGREGA EL DE LA DESCRIPCION
    '****************************************************************************
    SQLX$ = "UPDATE  " & TABLA1$ & " SET Descri_Lar = DESCRIP "
    SQLX$ = SQLX$ + " WHERE (Descri_Lar IS NULL OR Descri_Lar LIKE '') "
    SQLX$ = SQLX$ + " AND DOSECLAR = '" & NUFACTU$ & "' AND NUME_CLIE=" & NCLI2
    FLEXCONN.Execute (FILTSQL$(SQLX$))
    '
    '****************************************************************************
    LETRA$ = Mid$(NUFACTU$, 18, 1)
    FORIPRE$ = TRIM$(Control$("", "FORFAC-" + LETRA$))
    If CANTREG& = 0 And CANTREGREPA& = 0 Then TIPOCOM$ = "(T.Libre)"
    TIFADOC% = 2
        
    If Mid$(NUFACTU$, 1, 2) = "NC" Then TIFADOC% = 6
    If Mid$(NUFACTU$, 1, 2) = "ND" Then TIFADOC% = 3
    '
    If TIPOCOM$ = "(T.Libre)" Then
      If LETRA$ <> "E" Then   ' EXCLUYE FACTURAS DE EXPORTACION
        FORIPRE$ = TRIM$(Control$("", "FACTEXLI"))
        If UCase$(LETRA$) = "B" Then
          FORIPRE$ = TRIM$(Control$("", "FACTELIB"))
        End If
      End If
    End If
    '
    ' FACTURA EN MODO FACTURA DE CREDITO
    ModoCredito& = 0
    CondicionSql$ = "CODICOM_LAR = '" & NUFACTU$ & "' AND  NUCLIEN = " & NCLI & " AND MODOOPCIMOVIM='CR'"
    ModoCredito& = CantRegistros("CAJABANC", CondicionSql$)
    If ModoCredito& > 0 Then
        FORIPRE$ = TRIM$(Control$("", "FACCRE-" + LETRA$))
    End If
    
    'FORMULARIOS DE REPARACION
    If CANTREGREPA& > 0 And TIFADOC% = 2 Then
       FORIPRE$ = "FCR" + Mid$(FORIPRE$, 4)
    End If
    '
    '\\\OT-06-0038-WAR-01/02/06///
    FORIALT$ = ""
    If TIFADOC% = 3 Or TIFADOC% = 6 Then 'SI NC O ND
      If TIFADOC% = 3 Then ' SI NOTA DE DEBITO
        If UCase$(LETRA$) = "A" Then
          FORIALT$ = Control$("", "FORNDE-A") 'ND-A
          ElseIf UCase$(LETRA$) = "B" Then
          FORIALT$ = Control$("", "FORNDE-B") 'ND-B
          
          ElseIf UCase$(LETRA$) = "E" Then
          FORIALT$ = Control$("", "FORNDE-E") 'ND-E
        End If
      End If
      '
      If TIFADOC% = 6 Then ' SI NOTA DE CREDITO
        If UCase$(LETRA$) = "A" Then
          FORIALT$ = Control$("", "FORNCR-A") 'NC-A
          ElseIf UCase$(LETRA$) = "B" Then
          FORIALT$ = Control$("", "FORNCR-B") 'NC-B
          ElseIf UCase$(LETRA$) = "E" Then
          FORIALT$ = Control$("", "FORNCR-E") 'NC-E
        End If
      End If
    End If
    If TRIM$(FORIALT$) <> "" Then
      FORIPRE$ = FORIALT$
    End If
    '
    ' NOTA DE DEBITO EN MODO FACTURA DE CREDITO
    If ModoCredito& > 0 And TIFADOC% = 3 Then
        FORIPRE$ = TRIM$(Control$("", "NDBCRE-" + LETRA$))
    End If
    
    ' NOTA DE CREDITO EN MODO FACTURA DE CREDITO
    If ModoCredito& > 0 And TIFADOC% = 6 Then
        FORIPRE$ = TRIM$(Control$("", "NCRCRE-" + LETRA$))
    End If
    '
    ARCSQR$ = FORIPRE$
    If EXISTE%(LUDAT$ + ARCSQR$ + ".SQR") = 0 Then
       Call COMUNI("Falta Archivo " & ARCSQR$ + ".SQR")
    End If
    '
    'SE LE AGREGA POR DEFECTO LA LEYENDA ORIGINAL.
    FILTX$ = TRIM$(NUFACTU$) & ";" & TRIM$(Str$(NCLI2) & ";" & "ORIGINAL")
    
    'SE LE AGREGA CODIGO QR AFIP
    CONDI$ = "CODICOM_LAR = '" & NUFACTU$ & "' AND  NUCLIEN = " & NCLI
    TextoQRAfip$ = TRIM$(SAFETEXT$(VALOBADA("cajabanc", "CodigoQRAfip", CONDI$)))
    If TextoQRAfip$ <> "" Then
        TEXTOQR$ = TextoQRAfip$
        IMAG$ = "IMAGFACTU_QR.JPG"
        IDENTIFICADOR$ = "QR_FACTURA"
        Anchox% = 20
        ALTOX% = 20
        '
        On Error Resume Next
        Kill LUCOM$ & IMAG$
        On Error GoTo 0
        'RUTINA QUE LLAMA A APLICACION DE .NET PARA GENERAR LA IMAGEN DEL CODIGO QR
'        QRX$ = TRIM$(IMAGENCODBARRA$(LUCOM$, IMAG$, TEXTOQR$, "QR_CODE", ANCHOX%, ALTOX%))
        Call T_Espera(3)
        'STDSCREEN.Image1.Picture = LoadPicture(LUCOM$ & IMAG$)
'        Call GRABAR_QR(LUCOM$ & IMAG$, IDENTIFICADOR$)
        
'        Call RecortaImagenQRFactura
        
        FILTX$ = FILTX$ & ";QR_FACTURA"
    Else
        IMAG$ = "IMAGFACTU_QR.JPG"
        On Error Resume Next
        Kill LUCOM$ & IMAG$
        On Error GoTo 0
    End If
    
    CONDI$ = "DOSECLAR='" & NUFACTU$ & "' AND NUME_CLIE=" & NCLI2
    'SE INCORPORA PARA CONTEMPLAR A LAS REPARACIONES Y OBTENER DE MOVIREPA LA MONEDA
    TABMOV$ = "MOVISTO"
    If CANTREGREPA& > 0 Then TABMOV$ = "MOVIREPA"
    MONEMI% = XVALO(VALOBADA(TABMOV$, "MONEVAL", CONDI$))
    Call STDFORM(ARCSQR$, FILTX$)
    '
End Sub

Sub ACTUCOSTO_PROVEEDOR()
10  Call SELECHO("PROVED1")
    NCX1% = XVALO(VALACT1$("PROVED1"))
    If NCX1% > 0 Then
       '
       Call APERBASDAT("MASTER")
       '
       '
       'Lectura desde SQL
       SQLX$ = "SELECT * FROM MASTER with(NOLOCK) "
       SQLX$ = SQLX$ + " WHERE PROVHABI = " & NCX1% & " AND Status > 0"
       SQLX$ = SQLX$ + " ORDER BY CODIGO ASC"
       '
       Dim MasterBaja As New ADODB.Recordset
       Set MasterBaja = FLEXCONN.Execute(FILTSQL$(SQLX$))
       '
       With MasterBaja
          If Err Then
            On Error GoTo 0
            Call MENSERR(24, "No hay Registros que Coincidan con el Proveedor Elegido .")
          End If
          '
          i& = 0
          Do While Not .EOF
             COD$ = SAFETEXT$(!Codigo)
             MON% = XVALO(!Monecos)
             PODES1 = ROUND#(XVALO(!PorDescCos), 1): If PODES1 > 99 Then PODES1 = 0
             COS1# = XVALO(!COSUNI / (1 - PODES1 / 100))
             COS2# = XVALO(!COSUNI)
             fech = SAFETEXT$(!FECOSTO)
             '
             Y$ = REGBLAN$("ACTUCOS")
             CODI% = CODINT%(COD$)
             Call REPLA(Y$, MKI$(CODI%), 1, 2)
             Call REPLA(Y$, MKI$(MON%), 3, 2)
             Call REPLA(Y$, MKD$(COS1#), 5, 8)
             Call REPLA(Y$, MKS$(PODES1), 13, 4)
             Call REPLA(Y$, MKD$(COS2#), 17, 8)
             FCOS% = CVINT%(fech)
             Call REPLA(Y$, MKI$(FCOS%), 25, 2)
             i& = i& + 1
             Call GRAREG("!ACTUCOS", Y$, i&)
            .MoveNext
            '
          Loop
          Call SETULTREG("!ACTUCOS", i&)
       End With
       '
       
40     JUN% = VENTABU%("ACTUCOS/TITUPAN=Costo de Reposición - " + rasocli$((-1) * NCX1%))
       '
       If JUN% < 0 Then GoTo 10
       
'       If JUN% >= 0 Then
       
       
       EPAC$ = TRIM$(UCase$(EMPREAC$))
       If InStr(EPAC$, "J Y M") > 0 Then

        OBSERVAC.Caption = "Observación General de Cambio de Precio x Proveedor (Mandataria)"
        OBSERVAC.Text28 = ""
        OBSERVAC.Show 1                        'LLAMA A VENTANA DE OBSERVACIONES
        If OBSERVAC.Label1.Caption <> "OK" Or TRIM$(OBSERVAC.Text28) = "" Then
          Call COMUNI("Debe Ingresar un Texto General para el Cambio de Precio\No se van a Registrar el Mismo")
          GoTo 40                             'SI NO SE INGRESA OBSERVACION VUELVE A LA VENTANA AL INPUT BOX
        End If
        OBSERVACION$ = OBSERVAC.Text28
       End If
       If TRIM$(OBSERVACION$) = "" Then OBSERVACION$ = "Cambio de Precio Opción: x Proveedor"
       
       HOII% = HOY(HOS$)
       '
       Dim COSTOANTERIOR#(32767)
       For i& = 1 To 32767
         COSTOANTERIOR#(i&) = 0
       Next i&
       
        Call ABREMASTER
        For J& = 1 To ULTREG&("!ACTUCOS")
            Z$ = REGLEIDO$("!ACTUCOS", J&)
            COD3% = CVI(Mid$(Z$, 1, 2))
            PODES1 = CVS(Mid$(Z$, 13, 4)): If PODES1 < 0 Or PODES1 > 99 Then PODES1 = 0
            COU = (CVD(Mid$(Z$, 5, 8))) * (1 - PODES1 / 100)
            MONEI% = CVI(Mid$(Z$, 3, 2))
            FCOS% = CVI(Mid$(Z$, 25, 2))
            PHAB% = PROVHABI%(COD3%)
            '
            If PHAB% > 0 And NCX1% <> PHAB% Then
               TTXX$ = "Para el Artículo " + TRIM$(CODEXT$(COD3%)) + "\el Proveedor Habitual Registrado es\"
               TTXX$ = TTXX$ + TRIM$(rasocli$((-1) * PHAB%)) + "\  \"
               TTXX$ = TTXX$ + "Desea Cambiar el nuevo Proveedor Habitual a\"
               TTXX$ = TTXX$ + rasocli$((-1) * NCX1%)
               If COMALTER%(TTXX$ + "\\No, Conservar Anterior\Si, Cambiar") = 1 Then GoTo 10
            End If
            '
            ' Grabacion en Tabla SQL
            SQLX$ = " SELECT * FROM Master"
            SQLX$ = SQLX$ + " WHERE CODINT = " & COD3% & ""
            Dim MODMASTER As ADODB.Recordset
            Set MODMASTER = New ADODB.Recordset
            MODMASTER.Open FILTSQL$(SQLX$), FLEXCONN, adOpenDynamic, adLockPessimistic, adCmdText
            On Error Resume Next
                With MODMASTER
                    .MoveFirst
                    If Err < 1 Then
                        On Error GoTo 0
                        !PROVHABI = NCX1%
                        !FECOSTO = CVDAT(FCOS%)
                        'EN EL CASO QUE CAMBIE EL COSTO  PONE LA FECHA DE HOY
                        COSTOANTERIOR#(COD3%) = XVALO(!COSUNI)
                        If ROUND(XVALO(!COSUNI), 4) <> ROUND(COU, 4) Then
                           !FECOSTO = Now
                        End If
                        !COSUNI = COU
                        !PorDescCos = PODES1
                        !Monecos = MONEI%
                       .Update
                    End If
                 End With
                
        Next J&
        '
        'RECORRO NUEVAMENTE PARA ACTUALIZAR LAS NOVEDADES
        For J& = 1 To ULTREG&("!ACTUCOS")
            Z$ = REGLEIDO$("!ACTUCOS", J&)
            COD3% = CVI(Mid$(Z$, 1, 2))
            PODES1 = CVS(Mid$(Z$, 13, 4)): If PODES1 < 0 Or PODES1 > 99 Then PODES1 = 0
            COU = (CVD(Mid$(Z$, 5, 8))) * (1 - PODES1 / 100)
            MONEI% = CVI(Mid$(Z$, 3, 2))
            PHAB% = NCX1%
            
            If ROUND(COSTOANTERIOR#(COD3%), 4) <> ROUND(COU, 4) Then
               PREC_NVO# = XVALO(COU)
               Call ADD_NOVEDAD(COD3%, NCX1%, COSTOANTERIOR#(COD3%), PREC_NVO#, OBSERVACION$, MONEI%, "COSTO")
            End If
        Next J&
    End If
End Sub
Sub VERISALDO(NC%, TXTSALDO As TextBox)
   CONDI$ = "NUCLIEN = " & NC% & " AND (TIPOMOVIM = 'FHON' OR TIPOMOVIM = 'AJUD') AND STATUS < 9"
''''   IDEBEX# = XVALO(VALOBADA("CAJABANC", "SUM(IDEBECOMP)", CONDI$, "NOMESS"))
''''   CONDI$ = "NUCLIEN = " & NC% & " AND (TIPOMOVIM = 'RHON' ND TIPOMOVIM = 'AJUSALD') AND STATUS < 9"
''''   IHABEX# = XVALO(VALOBADA("CAJABANC", "SUM(IHABECOMP)", CONDI$, "NOMESS"))
''''   SALDOX# = IDEBEX# - IHABEX#
''''   If SALDOX# > 0.005 Then
''''      Call COMUNI("Saldo Cliente: " & FORMATNUM$(SALDOX#, "F12.2"))
''''   End If
''''   TXTSALDO = FORMATNUM$(SALDOX#, "F12.2")
   'PRESENTA MENSAJE Y LO IMPRIME EN EL TEXTBOX PASADO COMO PARAMETRO
   
   TXTSALDO.ForeColor = vbBlack
   CONDI$ = "NUCLIEN = " & NC% & " AND (TIPOMOVIM = 'FHON' OR TIPOMOVIM = 'AJUD' OR TIPOMOVIM = 'RHON') AND STATUS < 9"
   IDEBEX# = XVALO(VALOBADA("CAJABANC", "SUM(IDEBECOMP)", CONDI$, "NOMESS"))
'   CONDI$ = "NUCLIEN = " & NC% & " AND (TIPOMOVIM = 'RHON' ND TIPOMOVIM = 'AJUSALD') AND STATUS < 9"
   IHABEX# = XVALO(VALOBADA("CAJABANC", "SUM(IHABECOMP)", CONDI$, "NOMESS"))
   SALDOX# = IDEBEX# - IHABEX#
   If SALDOX# > 0.005 Then
      Call COMUNI("Saldo Cliente: " & FORMATNUM$(SALDOX#, "F12.2"))
      TXTSALDO.ForeColor = vbRed
   End If
   '
   TXTSALDO = FORMATNUM$(SALDOX#, "F12.2")
   

End Sub

Function Vali_Derechos_Lpre(NC, CI1%, NLISTA%, PRECIO_VENTA#)
    '
    'DEVUELVE 1 TIENE DERECHOS PARA HACER CAMBIOS DE LO MODIFICADO
    'DEVUELVE 0 SI SE MODIFICO SOLO LA LISTA DE PRECIOS Y NO TENIA DERECHO PARA ESTO
    'DEVUELVE -1   SI MODIFICO EL PRECIO SOLAMENTE Y NO TENIA DERECHOS
    'DEVUELVE-2 SI MODIFICO EL PRECIO Y LA LISTA SIN DERECHOS
    'DEVUELVE -3 SI EL CODIGO NO PERTENCE A LA LISTA
    'SI TIENE DERECHOS DE LECTURA Y ESCRITURA PARA LA LISTA ACTIVA APRUEBA DIRECTAMENTE.
    
    '**************************************************************************************************************************************
    '**********ATENCION SI EL CODIGO NO EXISTE EN LA LISTA NO SE VA A SOLUCIONAR CON MODIFICAR Y CAMBIAR LA LISTA Y SOBRE-ESCRIBIRLO ******
    '**********SE DEBERA MODIFICAR PEDIDO - ELIMINAR EL ITEM DEL DETALLE Y CON LA LISTA ACTIVA CORRESPONDIENTE INGRESARLO COMO UNO NUEVO **
    '**************************************************************************************************************************************
    Valid% = 1
    If NLISTA% < 0 Then NLISTA% = 0
    '
    'SI ES EL ADMINISTRADOR APRUEBA DIRECTAMENTE.
    USER1% = USNBR% Mod 100
    If USER1% = 1 Then GoTo 99
    '
    'VALIDACION SI EL CODIGO NO PERTENECE A LA LISTA DEVUELVE -3 Y SE VA
    If CantRegistros("PREVENTA", "Cod_Pieza = " & CI1%) < 1 Then
       Valid% = -3
       GoTo 99
    End If
    '
    NUMLISTA_CLIENTE% = LIPRCLI%(NC)
    CONDI$ = "Cod_Lista = " & NLISTA%
    CONDI$ = CONDI$ + " AND USUARIO = " & USER1%
    DEREX% = XVALO(VALOBADA("DERELISTAPRE", "DERECHO", CONDI$, "NOMESS"))
    If DEREX% >= 2 Then ' SI TIENE DERECHOS = 2 APRUEBA TODAS LAS OPCIONES
       GoTo 99
    ElseIf DEREX% <= 1 Then 'SI TIENE DERECHOS <= 1 Y ES LISTA 0 DEVUELVE -2 DIRECTAMENTE POR Q' PARA ESTA LISTA NECESITA LOS 2 DERECHOS
       If NLISTA% = 0 Then  'POR QUE ES IMPOSIBLE VALIDAR SI CAMBIA EL PRECIO.
          Valid% = -2
          GoTo 99
       End If
    End If
    '-------------------------------------------------------------------------------------
    'SI ES LA LISTA DEL CLIENTE APRUEBA LISTA O TIENE DERECHO > 0 MINIMAMENTE PUEDE CAMBIAR LA LISTA.
    'EN ESTE CASO EL DERECHO SOLO PODRIA SER 0 O 1
    If NUMLISTA_CLIENTE% <> NLISTA% And DEREX% <= 0 Then
      Valid% = Valid% - 1
    End If  '
    '--------------------------------------------------------------------------------------
    'SI CAMBIO EL PRECIO DEVUELVE -1 , o -2 SI TAMBIEN SE CAMBIO LA LISTA DE PRECIOS SIN DERECHOS
    If Abs(ROUND(PRECIO_VENTA#, 2) - ROUND(PRELISTA#(TRIM$(Str$(NLISTA%)), CI1%), 2)) > 0.05 Then
      Valid% = Valid% - 2
    End If
    '
    '---------------------------------------------------------------------------------------
99   Vali_Derechos_Lpre = Valid%
End Function
Function PROVELOTE%(LOTE$)
    'CASO INICIAL PARA LOS QUE TIENEN EL LOTE INICIAL QUE LE ASIGNE UN PROVEEDOR SEGUN
    'SI TERMINA CON Q ES EL PROVEEDOR AHP - USA
EPAC$ = TRIM$(UCase$(EMPREAC$))
If InStr(EPAC$, "AHP") > 0 Then
    If UCase$(Left(LOTE$, 2)) = "LI" Then
      CONDI$ = "idenlote = '" & LOTE$ & "'"
      CODEX$ = CODEXT$(XVALO(VALOBADA("LOTINGRE", "COD_INTE", CONDI$)))
      If Right(CODEX$, 1) = "Q" Then
         PROVELOTE% = 10014   ' FAMA
      Else
         PROVELOTE% = 1        'AHP - USA
      End If
      Exit Function ' SE VA
    End If
End If
    
    'FORMA ORIGINAL
    'DEVUELVE EL PROVEEDOR QUE INGRESO EL LOTE PASADO COMO PARAMETRO
    
    NROLOTE& = XVALO(Mid$(LOTE$, 4, 6))
    SQLX$ = " EXECUTE ProveedorSegunlote  " & NROLOTE&
    Dim rs As New ADODB.Recordset
    Set rs = FLEXCONN.Execute(SQLX$)
    PROVELOTE% = XVALO(rs!NumProv)
    rs.Close: Set rs = Nothing
    '
    'MODIFICACION PARA MINEPARTS 'SI EL PROVEEDOR ES MINEPARTS LE DEVUELVO EL REAL DEL EXTERIOR.
    NPROV& = PROVELOTE%
    If ProveedorRealExterior%(NPROV&) > 0 Then
        InfoProveMineparts% = CtaProv_ViaEnt_PjeNacionalizacion%(LOTE$, Ret_ProvExterior, Ret_ViaEntExt, Ret_PjeNacion, Ret_CoefNacion)
        If InfoProveMineparts% > 0 Then
           PROVELOTE% = Ret_ProvExterior
        End If
    End If
    
End Function

Function VIALOTE%(LOTE$)
    'DEVUELVE LA VIA DE ENTREGA QUE INGRESO EL LOTE PASADO COMO PARAMETRO
    
    NROLOTE& = XVALO(Mid$(LOTE$, 4, 6))
    CONDI$ = "IDENLOTE = '" & LOTE$ & "'"
    VIALOTE% = XVALO(VALOBADA("BOLRECEP", "Via_Entrega", CONDI$, "NOMESS"))
    If VIALOTE% < 1 Then
       NOC& = XVALO(VALOBADA("BOLRECEP", "NumOC", CONDI$, "NOMESS"))
       VIALOTE% = VIAENTRE_COMPRA(NOC&)
    End If

'SE MODIFICO PARA QUE VALIDE SEGUN EL IDENLOTE Y NO POR EL NUMEBORE EL IDENLOTE HILA MAS FINO

'    'DEVUELVE LA VIA DE ENTREGA QUE INGRESO EL LOTE PASADO COMO PARAMETRO
'    NROLOTE& = XVALO(Mid$(LOTE$, 4, 6))
'    CONDI$ = "NumeBoRe = " & NROLOTE&
'    VIALOTE% = XVALO(VALOBADA("BOLRECEP", "Via_Entrega", CONDI$, "NOMESS"))
'    If VIALOTE% < 1 Then
'       NOC& = XVALO(VALOBADA("BOLRECEP", "NumOC", CONDI$, "NOMESS"))
'       VIALOTE% = VIAENTRE_COMPRA(NOC&)
'    End If
End Function


'Function PORCENT_COSTO_NACIONALIZACION_RESPALDO_30_05_19(NUMEBORE&, Optional MONEDX%, Optional MONECOMP%)
'    PORCENT_COSTO_NACIONALIZACION = 0: GASTOS_APLICADOS_DOLAR# = 0: GASTOS_APLICADOS_PESOS# = 0: RET_GASTOFLETE# = 0
'    If NUMEBORE& < 1 Then Exit Function
'    '
'    If MONEDX% < 1 Then MONEDX% = 2 ' POR DEFECTO DOLAR
'    '
'    'la funcionalidad es obtener el porcentaje que hay de diferencia entre el valor de compra y de los gasotos
'    'asociados a la factura que se cargo en base a la recepcion que machea con esta orden de compra.
'    '----
'    '
'    'toma el numero de factura, el proveedor y la moneda en que se gurardo la factura que machea con la recepcion br
'    CONDI$ = " NUMEBORE =" & NUMEBORE&
'    NUFACTU$ = TRIM$(VALOBADA("DETFACOMSQL ", "FACTURAPROV", CONDI$, "NOMESS"))
'    If NUFACTU$ = "" Then Exit Function
'    NUMEOCOM& = XVALO(VALOBADA("DETFACOMSQL ", "NUME_OCOM", CONDI$, "NOMESS"))
'    NPROV& = XVALO(VALOBADA("DETFACOMSQL", "Nume_cli", CONDI$, "NOMESS"))
'
'    'OBTENGO EL VALOR DE COMPRA y el tipo de moneda  segun la factura que tomo en el parrafo anterior
'    CONDI$ = " FACTURAPROV = '" & NUFACTU$ & "' AND Nume_cli = " & NPROV&
''    NUFACTU$ = TRIM$(VALOBADA("FACTURAYGASTOS", "FACTURAPROV", CONDI$, "NOMESS"))
'    MONEFACTU% = XVALO(VALOBADA("FACTURAYGASTOS", "MoneVal", CONDI$, "NOMESS"))
'    VALCOMPRA# = XVALO(VALOBADA("FACTURAYGASTOS", "VAL_COMPRA", CONDI$, "NOMESS"))
'    NPROV& = XVALO(VALOBADA("FACTURAYGASTOS", "Nume_cli", CONDI$, "NOMESS"))
'    '
'    'desde aqui combio
'    'siempre lleva a dolares
'    If MONEFACTU% <> MONEDX% Then
'       COEFCON = TICAMONE(MONEFACTU%) / TICAMONE(MONEDX%)
'       VALCOMPRA# = COEFCON * VALCOMPRA#
'    End If
'    '
'    'si obtuvo
'    If NUFACTU$ <> "" Then
'       '
'       CONDI$ = "FACTURAPROV_ASOC = '" & NUFACTU$ & "' AND MoneVal = " & 1 & " and Nume_cli_ASOC = " & NPROV& & " and  TIPO_FACTURA = 'G'"
'       GASTOS_APLICADOS_PESOS# = XVALO(VALOBADA("FACTURAYGASTOS", "Sum(IMPO_GASTO)", CONDI$, "NOMESS"))
'
'       '''''''''' GASTO DE FLETE
'       'EL FLETE SIEMPRE ESTA EN PESOS EN EL CAJABANC.
'       CONDIFLET$ = "CODICOM_LAR = '" & NUFACTU$ & "' AND TIPOMOVIM = 'COMPRA' and NUPROVE = " & NPROV& & "   AND LEFT(DESCRIMOV,13) = 'Flete Export.'"
'       Call MONE_TIPOCAM_GASTOFLETE(CONDIFLET$, RET_MONEBANC%, RET_TIPOCAM#, RET_GASTOFLETE#)
'       '
'       If RET_GASTOFLETE# > 0 Then
'          GASTOS_APLICADOS_PESOS# = GASTOS_APLICADOS_PESOS# + RET_GASTOFLETE#
'       End If
'       '''''''''''FIN GASTO DE FLETE
'
'       If 1 <> MONEDX% Then
'           COEFCON = TICAMONE(1) / TICAMONE(MONEDX%)
'           GASTOS_APLICADOS_PESOS# = COEFCON * GASTOS_APLICADOS_PESOS# ' AQUI LOS CONVIERTO
'       End If
'       '
'       CONDI$ = "FACTURAPROV_ASOC = '" & NUFACTU$ & "' AND MoneVal = " & 2 & " and Nume_cli_ASOC = " & NPROV& & " and  TIPO_FACTURA = 'G'"
'       GASTOS_APLICADOS_DOLAR# = XVALO(VALOBADA("FACTURAYGASTOS", "Sum(IMPO_GASTO)", CONDI$, "NOMESS"))
'       If 2 <> MONEDX% Then
'          COEFCON = TICAMONE(2) / TICAMONE(MONEDX%)
'          GASTOS_APLICADOS_DOLAR# = COEFCON * GASTOS_APLICADOS_DOLAR#
'       End If
'
'       TOTAL_COMPRA_MAS_GASTO# = VALCOMPRA# + GASTOS_APLICADOS_PESOS# + GASTOS_APLICADOS_DOLAR#
'       PORCENT_COSTO_NACIONALIZACION = PORCENT_ENTRE_2_VALORES(TOTAL_COMPRA_MAS_GASTO#, VALCOMPRA#)
'       '
'
'       PORCENT_COSTO_NACIONALIZACION = 1 + (PORCENT_COSTO_NACIONALIZACION) / 100
'    End If
'
'
'    Exit Function
'
'
'    'hasta aqui
'
'    If MONECOMP% <> MONEDX% Then
'       COEFCON = TICAMONE(MONEFACTU%) / TICAMONE(MONEDX%)
'       VALCOMPRA# = COEFCON * VALCOMPRA#
'    End If
'
'
'    'OBTENGO EL GASTO RELACIONADO CON ESTA FACTURA ACUMULADO FILTRANDO POR MONEDA, POR SI SE CARGARON GASTOS CON DIFERENTES MONEDAS.
'    If NUFACTU$ <> "" Then
'       '
'       CONDI$ = "FACTURAPROV_ASOC = '" & NUFACTU$ & "' AND MoneVal = " & 1 & " and Nume_cli_ASOC = " & NPROV& & " and  TIPO_FACTURA = 'G'"
'       GASTOS_APLICADOS_PESOS# = XVALO(VALOBADA("FACTURAYGASTOS", "Sum(IMPO_GASTO)", CONDI$, "NOMESS"))
'
''       '''''''''' GASTO DE FLETE
''       'EL FLETE SIEMPRE ESTA EN PESOS EN EL CAJABANC.
''       CONDIFLET$ = "CODICOM_LAR = '" & NUFACTU$ & "' AND TIPOMOVIM = 'COMPRA' and NUPROVE = " & NPROV& & "   AND LEFT(DESCRIMOV,13) = 'Flete Export.'"
''       Call MONE_TIPOCAM_GASTOFLETE(CONDIFLET$, RET_MONEBANC%, RET_TIPOCAM#, RET_GASTOFLETE#)
''       '
''       If RET_GASTOFLETE# > 0 Then
''          GASTOS_APLICADOS_PESOS# = GASTOS_APLICADOS_PESOS# + RET_GASTOFLETE#
''       End If
''       '''''''''''FIN GASTO DE FLETE
'
'       If 1 <> MONEDX% Then
'           COEFCON = TICAMONE(1) / TICAMONE(MONEDX%)
'           GASTOS_APLICADOS_PESOS# = COEFCON * GASTOS_APLICADOS_PESOS# ' AQUI LOS CONVIERTO
'       End If
'       '
'       CONDI$ = "FACTURAPROV_ASOC = '" & NUFACTU$ & "' AND MoneVal = " & 2 & " and Nume_cli_ASOC = " & NPROV& & " and  TIPO_FACTURA = 'G'"
'       GASTOS_APLICADOS_DOLAR# = XVALO(VALOBADA("FACTURAYGASTOS", "Sum(IMPO_GASTO)", CONDI$, "NOMESS"))
'       If 2 <> MONEDX% Then
'          COEFCON = TICAMONE(2) / TICAMONE(MONEDX%)
'          GASTOS_APLICADOS_DOLAR# = COEFCON * GASTOS_APLICADOS_DOLAR#
'       End If
'
'       TOTAL_COMPRA_MAS_GASTO# = VALCOMPRA# + GASTOS_APLICADOS_PESOS# + GASTOS_APLICADOS_DOLAR#
'       PORCENT_COSTO_NACIONALIZACION = PORCENT_ENTRE_2_VALORES(TOTAL_COMPRA_MAS_GASTO#, VALCOMPRA#)
'       '
'       PORCENT_COSTO_NACIONALIZACION = 1 + (PORCENT_COSTO_NACIONALIZACION) / 100
'    End If
'    '
'End Function '
Function PORCENT_COSTO_NACIONALIZACION(NUMEBORE&, Optional MONEDX%, Optional MONECOMP%)
    PORCENT_COSTO_NACIONALIZACION = 0: GASTOS_APLICADOS_DOLAR# = 0: GASTOS_APLICADOS_PESOS# = 0: RET_GASTOFLETE# = 0
    If NUMEBORE& < 1 Then Exit Function
    '
    If MONEDX% < 1 Then MONEDX% = 2 ' POR DEFECTO DOLAR
    '
'Call mostrarDatosx("MONEDX% DENTRO DE PORCENT_COSTO_NACIONALIZACION: " & MONEDX%)
    
    'la funcionalidad es obtener el porcentaje que hay de diferencia entre el valor de compra y de los gasotos
    'asociados a la factura que se cargo en base a la recepcion que machea con esta orden de compra.
    '----
    '
    'toma el numero de factura, el proveedor y la moneda en que se gurardo la factura que machea con la recepcion br
    CONDI$ = " NUMEBORE =" & NUMEBORE&
'Call mostrarDatosx("CONDI$ DE NUFACTU$,NUMEOCOM&,NPROV& DE DETFACOMSQL: " & CONDI$)
    NUFACTU$ = TRIM$(VALOBADA("DETFACOMSQL ", "FACTURAPROV", CONDI$, "NOMESS"))
    If NUFACTU$ = "" Then Exit Function
    NUMEOCOM& = XVALO(VALOBADA("DETFACOMSQL ", "NUME_OCOM", CONDI$, "NOMESS"))
    NPROV& = XVALO(VALOBADA("DETFACOMSQL", "Nume_cli", CONDI$, "NOMESS"))
    
    'AQUI CAMBIO MINEPART
    If ProveedorRealExterior%(NPROV&) > 0 Then 'SI EL PROVEEDOR ES MINEPART
      CONDI$ = " NUMEBORE =" & NUMEBORE&       'LE APLICA EL QUE SE GUARDA CUANDO SE RECIBE EN EL BOLREDET
'Call mostrarDatosx("CONDI$ de ProveedorRealExterior%: " & CONDI$)
      
      LOTE$ = VALOBADA("DETFACOMSQL", "IDENLOTE", CONDI$, "NOMESS") 'PARA CUMPLIR CON LA FUNCION DE ABAJO
      InfoProveMineparts% = CtaProv_ViaEnt_PjeNacionalizacion%(LOTE$, Ret_ProvExterior, Ret_ViaEntExt, Ret_PjeNacion, Ret_CoefNacion)
      PORCENT_COSTO_NACIONALIZACION = Ret_CoefNacion ' ES EL COEFICIENTE QUE RETORNA LA FUNCION
      
      Exit Function
    End If
    
    '
    'SE AGREGO ESTE PARRAFO DONDE SE TOMA EL VALOR DE COTIZACION AL MOMENTO DE GUARDAR LA  FACTURA DEL PROVEEDOR
    COTIZACION_SEGUN_MONEDA_AL_MOMENTO_FACTURADO = TipoCambioFacturaProveedor(NPROV&, NUFACTU$, RET_TIPOMONEDA%)
'Call mostrarDatosx("COTIZACION_SEGUN_MONEDA_AL_MOMENTO_FACTURADO: " & ROUND(COTIZACION_SEGUN_MONEDA_AL_MOMENTO_FACTURADO, 2))
    Dim COTIZFACTU(6)
    COTIZFACTU(RET_TIPOMONEDA) = COTIZACION_SEGUN_MONEDA_AL_MOMENTO_FACTURADO
XX = COTIZFACTU(RET_TIPOMONEDA)
'Call mostrarDatosx("COTIZFACTU(RET_TIPOMONEDA): " & ROUND(XX, 2))
    'SI NO LLEGA A TRAER EL VALOR DE LA COTIZACION TOMA EL TICAMONE DEL DIA, (DEBERIA TRAER INFO SIEMPRE)
    If COTIZFACTU(RET_TIPOMONEDA) <= 0 Then COTIZFACTU(RET_TIPOMONEDA) = TICAMONE(RET_TIPOMONEDA)
XX = COTIZFACTU(RET_TIPOMONEDA)
'Call mostrarDatosx("COTIZFACTU(RET_TIPOMONEDA ASIGNADO): " & ROUND(XX, 2))
    
    '
    'OBTENGO EL VALOR DE COMPRA y el tipo de moneda  segun la factura que tomo en el parrafo anterior
    CONDI$ = " FACTURAPROV = '" & NUFACTU$ & "' AND Nume_cli = " & NPROV&
'Call mostrarDatosx("CONDI$ DE MONEFACTU%,VALCOMPRA# Y NPROV&: " & CONDI$)
'    NUFACTU$ = TRIM$(VALOBADA("FACTURAYGASTOS", "FACTURAPROV", CONDI$, "NOMESS"))
    MONEFACTU% = XVALO(VALOBADA("t", "MoneVal", CONDI$, "NOMESS"))
'    VALCOMPRA# = XVALO(VALOBADA("FACTURAYGASTOS", "VAL_COMPRA", CONDI$, "NOMESS"))
'Call mostrarDatosx("VAL_COMPRA: " & VALCOMPRA#)

    '12/11/21 SE CAMBIO PARA QUE TRAIGA EL VALOR DE LA FACTURA COMO BASE EN VEZ DEL DE LA ORDEN DECOMPRA
    VALCOMPRA# = XVALO(VALOBADA("FACTURAYGASTOS", "IMPORTE_FACTURA", CONDI$, "NOMESS")) 'ANTES VAL_COMPRA
    NPROV& = XVALO(VALOBADA("FACTURAYGASTOS", "Nume_cli", CONDI$, "NOMESS"))
'Call mostrarDatosx("IMPORTE_FACTURA: " & VALCOMPRA#)
    '
    'desde aqui combio
    'siempre lleva a dolares
    If MONEFACTU% <> MONEDX% Then
       'COEFCON = TICAMONE(MONEFACTU%) / TICAMONE(MONEDX%)
       COEFCON = COTIZFACTU(RET_TIPOMONEDA) / COTIZFACTU(RET_TIPOMONEDA)
       VALCOMPRA# = COEFCON * VALCOMPRA#
    End If
    '
'Call mostrarDatosx("VALCOMPRA#: " & ROUND(VALCOMPRA#, 2))
    'si obtuvo
    If NUFACTU$ <> "" Then
'Call mostrarDatosx("NUFACTU$: " & NUFACTU$)
       '
       CONDI$ = "FACTURAPROV_ASOC = '" & NUFACTU$ & "' AND MoneVal = " & 1 & " and Nume_cli_ASOC = " & NPROV& & " and  TIPO_FACTURA = 'G'"
       GASTOS_APLICADOS_PESOS# = XVALO(VALOBADA("FACTURAYGASTOS", "Sum(IMPO_GASTO)", CONDI$, "NOMESS"))
'Call mostrarDatosx("CONDI$ DE GASTOS_APLICADOS_PESOS#: " & CONDI$)
'Call mostrarDatosx("GASTOS_APLICADOS_PESOS#: " & ROUND(GASTOS_APLICADOS_PESOS#, 2))
       
       '''''''''' GASTO DE FLETE
       'EL FLETE SIEMPRE ESTA EN PESOS EN EL CAJABANC.
       CUECON% = CtaCbleRelaProveedor%(NPROV&)
'       CONDIFLET$ = "CODICOM_LAR = '" & NUFACTU$ & "' AND TIPOMOVIM = 'COMPRA' and NUPROVE = " & NPROV& & "   AND (LEFT(DESCRIMOV,13) = 'Flete Export.')"
       CONDIFLET$ = "CODICOM_LAR = '" & NUFACTU$ & "' AND TIPOMOVIM = 'COMPRA' and NUPROVE = " & NPROV& & "   AND CUECONTAI = " & CUECON%
'Call mostrarDatosx("CONDIFLET$: " & CONDIFLET$)
       Call MONE_TIPOCAM_GASTOFLETE(CONDIFLET$, RET_MONEBANC%, RET_TIPOCAM#, RET_GASTOFLETE#)
       '
'Call mostrarDatosx("RET_GASTOFLETE#: " & ROUND(RET_GASTOFLETE#, 2))

       If RET_GASTOFLETE# > 0 Then
          GASTOS_APLICADOS_PESOS# = GASTOS_APLICADOS_PESOS# + RET_GASTOFLETE#
       End If
       '''''''''''FIN GASTO DE FLETE
       
       If 1 <> MONEDX% Then
           'COEFCON = TICAMONE(1) / TICAMONE(MONEDX%)
           COEFCON = TICAMONE(1) / COTIZFACTU(RET_TIPOMONEDA)
           GASTOS_APLICADOS_PESOS# = COEFCON * GASTOS_APLICADOS_PESOS# ' AQUI LOS CONVIERTO
       End If
       '
       CONDI$ = "FACTURAPROV_ASOC = '" & NUFACTU$ & "' AND MoneVal = " & 2 & " and Nume_cli_ASOC = " & NPROV& & " and  TIPO_FACTURA = 'G'"
       GASTOS_APLICADOS_DOLAR# = XVALO(VALOBADA("FACTURAYGASTOS", "Sum(IMPO_GASTO)", CONDI$, "NOMESS"))
       If 2 <> MONEDX% Then
          'COEFCON = TICAMONE(2) / TICAMONE(MONEDX%)
          COEFCON = COTIZFACTU(RET_TIPOMONEDA) / TICAMONE(MONEDX%)
          GASTOS_APLICADOS_DOLAR# = COEFCON * GASTOS_APLICADOS_DOLAR#
       End If
'Call mostrarDatosx("GASTOS_APLICADOS_DOLAR#: " & ROUND(GASTOS_APLICADOS_DOLAR#, 2))

       TOTAL_COMPRA_MAS_GASTO# = VALCOMPRA# + GASTOS_APLICADOS_PESOS# + GASTOS_APLICADOS_DOLAR#
'Call mostrarDatosx("TOTAL_COMPRA_MAS_GASTO#: " & ROUND(TOTAL_COMPRA_MAS_GASTO#, 2))

       PORCENT_COSTO_NACIONALIZACION = PORCENT_ENTRE_2_VALORES(TOTAL_COMPRA_MAS_GASTO#, VALCOMPRA#)
       '
       PORCENT_COSTO_NACIONALIZACION = 1 + (PORCENT_COSTO_NACIONALIZACION) / 100
'Call mostrarDatosx("PORCENT_COSTO_NACIONALIZACION: " & ROUND(PORCENT_COSTO_NACIONALIZACION, 2))
    End If
    '
    Exit Function
    '
    'hasta aqui
    
    If MONECOMP% <> MONEDX% Then
       COEFCON = TICAMONE(MONEFACTU%) / TICAMONE(MONEDX%)
       VALCOMPRA# = COEFCON * VALCOMPRA#
    End If
   
    
    'OBTENGO EL GASTO RELACIONADO CON ESTA FACTURA ACUMULADO FILTRANDO POR MONEDA, POR SI SE CARGARON GASTOS CON DIFERENTES MONEDAS.
    If NUFACTU$ <> "" Then
       '
       CONDI$ = "FACTURAPROV_ASOC = '" & NUFACTU$ & "' AND MoneVal = " & 1 & " and Nume_cli_ASOC = " & NPROV& & " and  TIPO_FACTURA = 'G'"
       GASTOS_APLICADOS_PESOS# = XVALO(VALOBADA("FACTURAYGASTOS", "Sum(IMPO_GASTO)", CONDI$, "NOMESS"))
       
'       '''''''''' GASTO DE FLETE
'       'EL FLETE SIEMPRE ESTA EN PESOS EN EL CAJABANC.
'       CONDIFLET$ = "CODICOM_LAR = '" & NUFACTU$ & "' AND TIPOMOVIM = 'COMPRA' and NUPROVE = " & NPROV& & "   AND LEFT(DESCRIMOV,13) = 'Flete Export.'"
'       Call MONE_TIPOCAM_GASTOFLETE(CONDIFLET$, RET_MONEBANC%, RET_TIPOCAM#, RET_GASTOFLETE#)
'       '
'       If RET_GASTOFLETE# > 0 Then
'          GASTOS_APLICADOS_PESOS# = GASTOS_APLICADOS_PESOS# + RET_GASTOFLETE#
'       End If
'       '''''''''''FIN GASTO DE FLETE

       If 1 <> MONEDX% Then
           COEFCON = TICAMONE(1) / TICAMONE(MONEDX%)
           GASTOS_APLICADOS_PESOS# = COEFCON * GASTOS_APLICADOS_PESOS# ' AQUI LOS CONVIERTO
       End If
       '
       CONDI$ = "FACTURAPROV_ASOC = '" & NUFACTU$ & "' AND MoneVal = " & 2 & " and Nume_cli_ASOC = " & NPROV& & " and  TIPO_FACTURA = 'G'"
       GASTOS_APLICADOS_DOLAR# = XVALO(VALOBADA("FACTURAYGASTOS", "Sum(IMPO_GASTO)", CONDI$, "NOMESS"))
       If 2 <> MONEDX% Then
          COEFCON = TICAMONE(2) / TICAMONE(MONEDX%)
          GASTOS_APLICADOS_DOLAR# = COEFCON * GASTOS_APLICADOS_DOLAR#
       End If

       TOTAL_COMPRA_MAS_GASTO# = VALCOMPRA# + GASTOS_APLICADOS_PESOS# + GASTOS_APLICADOS_DOLAR#
       PORCENT_COSTO_NACIONALIZACION = PORCENT_ENTRE_2_VALORES(TOTAL_COMPRA_MAS_GASTO#, VALCOMPRA#)
       '
       PORCENT_COSTO_NACIONALIZACION = 1 + (PORCENT_COSTO_NACIONALIZACION) / 100
    End If
    '
End Function '

Function PORCENT_COSTO_NACIONALIZACION_modificado(NUMEBORE&, Optional MONEDX%, Optional MONECOMP%)
    PORCENT_COSTO_NACIONALIZACION_modificado = 0: GASTOS_APLICADOS_DOLAR# = 0
    If NUMEBORE& < 1 Then Exit Function
    '
    If MONEDX% < 1 Then MONEDX% = 2 ' POR DEFECTO DOLAR
    '
    
    CONDI$ = " NUMEBORE =" & NUMEBORE&
    NUFACTU$ = TRIM$(VALOBADA("DETFACOMSQL ", "FACTURAPROV", CONDI$, "NOMESS"))
    NUMEOCOM& = XVALO(VALOBADA("DETFACOMSQL ", "NUME_OCOM", CONDI$, "NOMESS"))
    NPROV& = XVALO(VALOBADA("DETFACOMSQL", "Nume_cli", CONDI$, "NOMESS"))
    
    'OBTENGO EL VALOR DE COMPRA
    CONDI$ = " FACTURAPROV = '" & NUFACTU$ & "' AND Nume_cli = " & NPROV&
'    NUFACTU$ = TRIM$(VALOBADA("FACTURAYGASTOS", "FACTURAPROV", CONDI$, "NOMESS"))
    MONEFACTU% = XVALO(VALOBADA("FACTURAYGASTOS", "MoneVal", CONDI$, "NOMESS"))
    VALCOMPRA# = XVALO(VALOBADA("FACTURAYGASTOS", "VAL_COMPRA", CONDI$, "NOMESS"))
    NPROV& = XVALO(VALOBADA("FACTURAYGASTOS", "Nume_cli", CONDI$, "NOMESS"))
    '
    If MONECOMP% <> MONEDX% Then
       COEFCON = TICAMONE(MONEFACTU%) / TICAMONE(MONEDX%)
       VALCOMPRA# = COEFCON * VALCOMPRA#
    End If

    
    'OBTENGO EL GASTO RELACIONADO CON ESTA FACTURA ACUMULADO FILTRANDO POR MONEDA, POR SI SE CARGARON GASTOS CON DIFERENTES MONEDAS.
    If NUFACTU$ <> "" Then
       '
       CONDI$ = "FACTURAPROV_ASOC = '" & NUFACTU$ & "' AND MoneVal = " & 1 & " and Nume_cli_ASOC = " & NPROV& & " and  TIPO_FACTURA = 'G'"
       GASTOS_APLICADOS_PESOS# = XVALO(VALOBADA("FACTURAYGASTOS", "Sum(IMPO_GASTO)", CONDI$, "NOMESS"))
       If 1 <> MONEDX% Then
          COEFCON = TICAMONE(1) / TICAMONE(MONEDX%)
          GASTOS_APLICADOS_PESOS# = COEFCON * GASTOS_APLICADOS_PESOS# ' AQUI LOS CONVIERTO
       End If
       '
       'EL FLETE SIEMPRE ESTA EN PESOS EN EL CAJABANC.
       CONDIFLET$ = "CODICOM_LAR = '" & NUFACTU$ & "' AND TIPOMOVIM = 'COMPRA' and NUPROVE = " & NPROV& & "   AND LEFT(DESCRIMOV,13) = 'Flete Export.'"
       Call MONE_TIPOCAM_GASTOFLETE(CONDIFLET$, RET_MONEBANC%, RET_TIPOCAM#, RET_GASTOFLETE#)
       '
       If RET_GASTOFLETE# > 0 Then
         If MONECOMP% <> 1 Then
           If RET_MONEBANC% = 1 Then RET_TIPOCAM# = TICAMONE(MONEDX%)
           COEFCON = TICAMONE(1) / RET_TIPOCAM#
           RET_GASTOFLETE# = COEFCON * RET_GASTOFLETE#
         End If
       End If
       CONDI$ = "FACTURAPROV_ASOC = '" & NUFACTU$ & "' AND MoneVal = " & 2 & " and Nume_cli_ASOC = " & NPROV& & " and  TIPO_FACTURA = 'G'"
       GASTOS_APLICADOS_DOLAR# = XVALO(VALOBADA("FACTURAYGASTOS", "Sum(IMPO_GASTO)", CONDI$, "NOMESS"))
       If 2 <> MONEDX% Then
          COEFCON = TICAMONE(2) / TICAMONE(MONEDX%)
          GASTOS_APLICADOS_DOLAR# = COEFCON * GASTOS_APLICADOS_DOLAR#
       End If
       TOTAL_COMPRA_MAS_GASTO# = VALCOMPRA# + GASTOS_APLICADOS_PESOS# + GASTOS_APLICADOS_DOLAR# + RET_GASTOFLETE#
       PORCENT_COSTO_NACIONALIZACION_modificado = PORCENT_ENTRE_2_VALORES(TOTAL_COMPRA_MAS_GASTO#, VALCOMPRA#)
       '
       PORCENT_COSTO_NACIONALIZACION_modificado = 1 + (PORCENT_COSTO_NACIONALIZACION_modificado) / 100
    End If
    '
    
End Function

Function MONELOTE%(LOTE$)
    MONELOTE% = 1
    CONDI$ = "IDENLOTE = '" & LOTE$ & "'"
    NUOC& = XVALO(VALOBADA("BOLRECEP", "NumOC", CONDI$, "NOMESS"))
    CI1% = XVALO(VALOBADA("BOLRECEP", "COD_INTE", CONDI$, "NOMESS"))
    If NUOC& > 0 Then
       CONDI$ = "STAT_OCOM < 9 AND COD_INTE= " & CI1% & " AND NUME_OCOM = " & NUOC&
       MONELOTE% = XVALO(VALOBADA("OCOMDETA", "Mone_Emis", CONDI$, "NOMESS"))
    End If
    If MONELOTE% < 1 Then
       CONDI$ = "STAT_OCOM = 9 AND COD_INTE= " & CI1% & " AND NUME_OCOM = " & NUOC&
       MONELOTE% = XVALO(VALOBADA("OCOMDETA", "Mone_Emis", CONDI$, "NOMESS"))
    End If


End Function
Function PRECIOUNITCOMPRA#(LOTE$)
    'revisar por que esta mal falta el codigo y que status < 9
    PRECIOUNITCOMPRA# = 0
    CONDI$ = "IDENLOTE = '" & LOTE$ & "'"
    NUOC& = XVALO(VALOBADA("BOLRECEP", "NumOC", CONDI$, "NOMESS"))
    If NUOC& > 0 Then
       CONDI$ = "NUME_OCOM = " & NUOC&
       PRECIOUNITCOMPRA# = XVALO(VALOBADA("OCOMDETA", "Pre_Unit", CONDI$, "NOMESS"))
    End If

End Function
Function NumeroSolicitud&(NPROV, NPRES&)
   'DEVUELVE EL NUMERO DE SOLICITUD SEGUN EN NRO DE PRESUPUESTO Y EL DEL PROVEEDOR.
   NumeroSolicitud& = XVALO(VALOBADA("PRESUCOMPRA", "NumSolCompra", "NroPres = " & NPRES& & " and  Num_Prov = " & NPROV, "NOMESS"))
End Function

Function ProxNumSolCompra&()
   ProxNumSolCompra& = 1 + XVALO(VALOBADA("PRESUCOMPRA", "MAX(NumSolCompra)", "NumSolCompra > 0", "NOMESS"))
End Function
Function NUEVONUMPRESCOMPRA&()
    'DEVUELVE EL VALOR SIGUIENTE AL MAXIMO DEL CAMPO NROPRES  DE LA TABLA PRESUCOMPRA
    NUEVONUMPRESCOMPRA& = 1 + XVALO(VALOBADA("PRESUCOMPRA", "MAX(NroPres)", "NroPres > 0", "NOMESS"))
End Function

Function OrdenCompraSegunSolicitud&(NSOLIC&)
    CONDI$ = "Medi_Ocom = " & NSOLIC&
    CONDI$ = CONDI$ + " AND STAT_OCOM < 8"
    OrdenCompraSegunSolicitud& = XVALO(VALOBADA("OCOMDETA", "nume_ocom", CONDI$, "NOMESS"))
End Function

Function TIENE_FACTURAS_ANTICIPO%(NC)
   '0 NO TIENE FACTURA ANTICIPO
   '1 SI TIENE FACTURA ANTICIPO

   TIENE_FACTURAS_ANTICIPO% = 0
   CONDI$ = "NUCLIEN = " & NC
   If CantRegistros&("FACTUANTI", CONDI$, "NOMESS") Then TIENE_FACTURAS_ANTICIPO% = 1
End Function
Function EsFacturaAnticipo%(NCL, FACTU$)
   '0 NO ES FACTURA ANTICIPO
   '1 ES FACTURA ANTICIPO
   
   EsFacturaAnticipo% = 0
   If Left$(FACTU$, 3) = "NC-" Then
     CONDI$ = "IMPOHABE > 0 AND  DOCU_APLI = '" & FACTU$ & "' And NUCLIEN = " & NCL
     If CantRegistros&("FACTUANTI", CONDI$, "NOMESS") > 0 Then EsFacturaAnticipo% = 1
   Else
     CONDI$ = "IMPODEBE > 0 AND (DOCU_APLI = '' OR DOCU_APLI IS NULL) AND DOCU_ORIG = '" & FACTU$ & "' And NUCLIEN = " & NCL
     If CantRegistros&("FACTUANTI", CONDI$, "NOMESS") > 0 Then EsFacturaAnticipo% = 1
   End If
End Function
Function EsAplicacionDeAnticipo%(NCL, FACTU$)
   '0 NO ES APLICACION DE  FACTURA ANTICIPO
   '1 SI ES APLICACION DE FACTURA ANTICIPO

   EsAplicacionDeAnticipo% = 0
   CONDI$ = "IMPOHABE > 0 And DOCU_APLI = '" & FACTU$ & "' And NUCLIEN = " & NCL
   If CantRegistros&("FACTUANTI", CONDI$, "NOMESS") > 0 Then EsAplicacionDeAnticipo% = 1
   '
End Function

Function ImporteRealFactura(NCL, FACTU$)
   'DEVUELVE EL IMPORTE REAL DE UNA FACTURA A LA CUAL SE LE APLICO UN ANTICIPO (AMBEST) Y QUE POR ESTE MOTIVO LA MISMA SE GRABA X EJEMPLO CON 0.01
   CONDI$ = "DOSECLAR = '" & FACTU$ & "' And NUME_CLIE = " & NCL
   CONDI$ = CONDI$ + " AND (DESCRI_LAR not LIKE 'Ant.FC-%' and DESCRI_LAR not LIKE 'Dto Anticipo FC-%')"
   'CONDI$ = CONDI$ + "AND DESCRI_LAR like 'Desc.:%' "
   ImpRealFact = XVALO(VALOBADA("MOVISTO", "SUM(ValTotItIva)", CONDI$, "NOMESS"))
   'ver si no se puede traer de cajabanc: no no se puede por que en el movisto esta separado los descuentos
   'y el anticipo.
   If ImpRealFact <= 0.005 Then
      ImpRealFact = XVALO(VALOBADA("MOVIPRU", "SUM(ValTotItIva)", CONDI$, "NOMESS"))
   End If
   
   If UCase$(Right(FACTU$, 1)) = "A" Then
      ImpRealFact = ImpRealFact * 1.21
   End If
   ImporteRealFactura = ImpRealFact
   
End Function
Function SaldoPendFacturaAnticipo#(NCL, FACTU$)
   'DEVUELVE EL SALDO PENDIENTE DE APLICACION DE LA FACTURA ANTICIPO
   'SE  AGREGO EL SUM A IMPODEBE POR QUE CUANDO HAY APLICACIONES ANTERIORES PUEDE SER QUE DEVUELVA 0 SEGUN EL ORDEN
   CONDI$ = "DOCU_ORIG = '" & FACTU$ & "' AND TIPOMOVIM = 'FVEN' And NUCLIEN = " & NCL
   ImpOriginal# = XVALO(VALOBADA("FACTUANTI", "SUM(IMPODEBE)", CONDI$, "NOMESS"))
   Aplicaciones# = XVALO(VALOBADA("FACTUANTI", "Sum(IMPOHABE)", CONDI$, "NOMESS"))
   '
   SALDOX# = ImpOriginal# - Aplicaciones#
   If SALDOX# < 0.005 Then SALDOX# = 0
   '
   SaldoPendFacturaAnticipo# = SALDOX#
   '
End Function

'Function MostrarFactAnticConSaldo$(NC)
'    'Call CARGA_ENCABEZADO(FRMZELECHO.MSF1, "Cta.Prov/F9;R.Social/A40;O.Compra/F9;Sol.Compra/F12", FRMZELECHO, , , ANTOT)
'    Campos$ = "Fact.Anticipo/A20;Imp.Fact.Ant./F12.2;Imp.Aplicado/F12.2;Saldo-Tot/F12.2"
'    Call CARGA_ENCABEZADO(FRMZELECHO.MSF1, Campos$, FRMZELECHO, , , ANTOT)
'    Call CARGA_ENCABEZADO(FRMZELECHO.msf2, Campos$, FRMZELECHO, , , ANTOT)
'    If ANTOT < 5500 Then ANTOT = 5500
'    FRMZELECHO.Width = ANTOT + 1500
'    FRMZELECHO.Caption = "Facturas Anticipo de Cliente N°: " & NC
'    '
'    FRMZELECHO.MSF1.Rows = 1
'
'    SQLX$ = " Select NUCLIEN,SUM(IMPODEBE) AS IDEBE, SUM(IMPOHABE)AS IHABE , DOCU_ORIG  from factuanti with(nolock)"
'    SQLX$ = SQLX$ + " where nuclien =" & NC
'    SQLX$ = SQLX$ + " AND TIPOMOVIM = 'FVEN'"
'    SQLX$ = SQLX$ + " GROUP BY NUCLIEN,DOCU_ORIG"
'    Set RS1 = READSET(SQLX$)
'    With RS1
'        Do While Not .EOF
'             NUFACTU$ = TRIM$(SAFETEXT$(!DOCU_ORIG))
'             '
'             j& = j& + 1
'             FRMZELECHO.msf2.Rows = j& + 1
'             FRMZELECHO.msf2.TextMatrix(j&, 1) = NUFACTU$
'             FRMZELECHO.msf2.TextMatrix(j&, 2) = FORMATNUM$((XVALO(!IDEBE)), "G12.2")
'             FRMZELECHO.msf2.TextMatrix(j&, 3) = FORMATNUM$((XVALO(!IHABE)), "G12.2")
'             SALDOFACTURA# = XVALO(!IDEBE) - XVALO(!IHABE)
'             SALDO# = SALDO# + XVALO(!IDEBE) - XVALO(!IHABE)
'             FRMZELECHO.msf2.TextMatrix(j&, 4) = FORMATNUM$(SALDOFACTURA#, "G12.2")
'             If j& Mod 2 Then Call COLOREAR_GRILLA_SOLO_UNA_AFILA(FRMZELECHO.msf2, COLOR1, FRMZELECHO.msf2.Rows - 1)
'             DOCUORIG_ANT$ = DOCUORIG$
'            .MoveNext
'        Loop
'    End With
'    RS1.Close
'    Set RS1 = Nothing
'    FRMZELECHO.Show 1
'    MostrarFactAnticConSaldo$ = RESP_ZELE_VECT(1)
'End Function
Sub GRAANTICIPO(NumeClie, Credito$, NFACT$, IMPOAPLI#)
        '
        'VALIDO SI ES UNA FACTURA ANTICIPO
        If UCase$(Left$(TRIM$(Credito$), 3)) <> "NC-" Then Exit Sub
        
        If EsFacturaAnticipo%(NumeClie, NFACT$) < 1 Then Exit Sub
        '
        TXTX$ = "Factura: " & FACTURA$
        TXTX$ = TXTX$ + " Corresponde a una Factura Anticipo\Desea Informar el Importe Descontando del Saldo a Entregar"
        TXTX$ = TXTX$ + "\\Imputar Sin Descontar de Saldo a Entregar\Imputar Descontando del Saldo a Entregar"
        OPX% = COMALTER%(TXTX$)
        If OPX% < 2 Then Exit Sub 'SI NO ELIGE QUE DESCUENTA DEL SALDO SE VA
        '
        'INFORMA EN LA TABLA FACTUANTI, DESCONTANDO LA NOTA DE CREDITO COMO SALDO DE ENTREGA
           SALPENDIENTE# = SaldoPendFacturaAnticipo#(NumeClie, NFACT$)
           If IMPOAPLI# > SALPENDIENTE# Then IMPOAPLI# = SALPENDIENTE#
           NCLI1 = NumeClie
           IMPOAPLICAR# = IMPOAPLI#
           TiMovi1$ = "AJUS"
           BASEIMPODE1# = 0
           BASEIMPOHA1# = 0
           COTI_DOLAR1# = TICAMONE(MONEMI%)
           VaMovi1$ = Right$(NFACT$, 1)

           Set RS10 = New ADODB.Recordset
           SQLX$ = "SELECT * FROM FACTUANTI WHERE DOCU_ORIG = '' "
333        On Error Resume Next
           RS10.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
           If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
              On Error GoTo 0
              Call CapturaErrorOpen
              GoTo 333
           End If
           On Error GoTo 0
           '
           With RS10
                  .AddNew
                  On Error Resume Next
                  !CODIEMPR = CodiEmp%
                  On Error GoTo 0
                  !DOCU_ORIG = NFACT$
                  !VARIMOVIM = VaMovi1$
                  !DOCU_APLI = Credito$
                  !IMPODEBE = 0
                  !IMPOHABE = IMPOAPLICAR#
                  !NUCLIEN = Int(GRUDIVCLI(NCLI1))
                  !Cli_Real = NCLI1
                  !RefeCom = rasocli$(NCLI1)
                  !NumeCom = XVALO(Mid$(NFCY$, 9, 8))
                  !FECHEMISI = FECHATEX$(HOY(HO$))
                  !Status = 0
                  !TipoMovim = TiMovi1$
                  !COTIZA_DOLAR = COTI_DOLAR#
                  !BASEIMPODEBE = 0
                  !BASEIMPOHABE = BASEIMPOHA1#
                  !USUARIO = USERNAME$
                  !OBSERVACIONES = "APLICACION NOTA DE CREDITO A FC. ANTICIPO"
                  !Mone_Emis = MONEMI%
                  .Update
           End With
           Screen.MousePointer = 1
End Sub
Function CantiYaFacturadaSegunBr(NBORE&, CI1%)
   'DEVUELVE EL TOTAL FACTURADO DE CADA ARTICULO, SI UN ARTICULO ESTA REPETIDO EN LA  MISMA RECEPCICON
   'LOS ACUMULA A DIFERENCIA DE LA FUNCION 'CantiYaFacturadaSegunBrID' QUE DIFERENCIA POR EL IDENLOTE DONDE CADA ITEM
   'TIENE SU IDENTIFICADOR UNIVOCO
   CantiYaFacturadaSegunBr = 0
   SQLX$ = "SELECT SUM(CantFacturada) AS CANTIDADFACTURADA FROM DETFACOMSQL WITH(NOLOCK)"
   SQLX$ = SQLX$ + " WHERE NumeBoRe = " & NBORE&
   SQLX$ = SQLX$ + " AND COD_INTE = " & CI1%
   Set rs = READSET(SQLX$)
   With rs
      CantiYaFacturadaSegunBr = XVALO(!CANTIDADFACTURADA)
   End With
   '
End Function
Function CantiYaFacturadaSegunBrID(IDLOT$, CI1%)
   CantiYaFacturadaSegunBrID = 0
   'DEVUELVE LA CANTIDAD FACTURADA SEGUN UN LOTE OSEA CONSIDERA EL IDENLOTE = BR-000101-1 PARA DIFERENCIAR SI
   'UN MISMO CODIGO ESTA CARGADO MAS DE UNA VEZ EN LA ORDEN DE COMPRA PROBABLEMENTE CON DIFERENTE PRECIO
   'ESTO SE DA EN AHP.
   SQLX$ = "SELECT SUM(CantFacturada) AS CANTIDADFACTURADA FROM DETFACOMSQL WITH(NOLOCK)"
   SQLX$ = SQLX$ + " WHERE IDENLOTE = '" & IDLOT$ & "'"
   SQLX$ = SQLX$ + " AND COD_INTE = " & CI1%
   Set rs = READSET(SQLX$)
   With rs
      CantiYaFacturadaSegunBrID = XVALO(!CANTIDADFACTURADA)
   End With
   
End Function

Function ValorCompraFacturaExpo#(NPROV, NUFACTURA$, Optional MON%, Optional SUMARAJUSTE%)
   'DEVUELVE EL VALOR DE COMPRA GUARDADO EN LA TABLA FACUTRA Y GASTOS
   'EL MISMO SE GUARDA EN LA MISMA MONEDA QUE LA FACTURA CARGADA.
   'SI SE LE PASA PARAMETRO MON%, SI ES DIFERENTE A LA MONEDA EN QEU ESTA GUARDADA DEVUELVE EL VALOR PERO CALCUALADO A CAMBIO ACTUAL.
    VALOR_COMPRA# = 0
    If MON% < 1 Then MONRESU% = 1
    CONDI$ = "Nume_Cli = " & NPROV & " AND TIPO_FACTURA = 'M'"
    CONDI$ = CONDI$ + " AND FACTURAPROV = '" & NUFACTURA$ & "'"
    SQLX1$ = SQLX1$ + " SELECT VAL_COMPRA,MONEVAL ,TIPOCAM, AJUSTE_VAL_COMPRA FROM FACTURAYGASTOS WITH(NOLOCK) WHERE " & CONDI$
   ' SQLX1$ = SQLX1$ + " AND VAL_COMPRA > 0 AND VAL_COMPRA <> IMPORTE_FACTURA" ' NO MUESTRO  LAS QUE ES EL MISMO VALOR P Q NO TIENE SENTIDO
    Set RS1 = READSET(SQLX1$)
    With RS1
         VALOR_COMPRA# = XVALO(!VAL_COMPRA)
         VALCOM_AJUST = 0
         If SUMARAJUSTE% > 0 Then
           VALCOM_AJUST = XVALO(!AJUSTE_VAL_COMPRA)
           VALOR_COMPRA# = VALOR_COMPRA# + VALCOM_AJUST
         End If
         If MON% > 0 Then
            MONFAC% = XVALO(!MoneVal)
            If MON% <> MONFAC% Then
               TCAMBIO# = XVALO(!TIPOCAM)
               COEFCON = TCAMBIO# / TICAMONE(MON%)
               VALOR_COMPRA# = VALOR_COMPRA# * COEFCON
            End If
         End If
    End With
    RS1.Close
    Set RS1 = Nothing
    ValorCompraFacturaExpo# = VALOR_COMPRA#
End Function

Function GASFLETE#(NPROV, NUFACTURA$, Optional MON%)
   'DEVUELVE EL GASTO DE FLETE DEL CAJABANC
   GASFLETE# = 0
   CUEI% = CtaCbleRelaProveedor%(NPROV)
   If CUEI% > o Then
       CONDI$ = "NUPROVE  = " & NPROV & " AND CUECONTAI = " & CUEI%
       CONDI$ = CONDI$ + " AND CODICOM_LAR = '" & NUFACTURA$ & "'"
       SQLX1$ = SQLX1$ + " SELECT IDEBECOMP,MONE_EMIS, TIPO_CAM FROM CAJABANC WITH(NOLOCK) WHERE " & CONDI$
       Set RS1 = READSET(SQLX1$)
       With RS1
         IMPO# = XVALO(!Idebecomp)
         If MON% <> 1 Then
            MONFAC% = XVALO(!Mone_Emis)
            If MON% = MONFAC% Then
               TCAMBIO# = XVALO(!Tipo_Cam)
               'COEFCON = TCAMBIO# / TICAMONE(MON%)
               IMPO# = IMPO# / TCAMBIO#
            End If
         End If

       End With
       RS1.Close
       Set RS1 = Nothing
       GASFLETE# = IMPO#
   End If

End Function

Function CtaCbleRelaProveedor%(NPROV)
   NPROVTEXT$ = SAFETEXT$(NPROV)
   CONDI$ = " PROBUS = 'PROVEEDOR CUENTA CONTABLE' AND CLABUS = '" & NPROV & "'"
   CtaCbleRelaProveedor% = XVALO(VALOBADA("FLEXCRL", "VALCONTROL", CONDI$, "NOMESS"))
   
End Function
Function CtaCbleRelaProveedorXDefectoEnDetalle%(NPROV)
   NPROVTEXT$ = SAFETEXT$(NPROV)
   CONDI$ = " PROBUS = 'CTA X DEFECTO CARGA DETALLE' AND CLABUS = '" & NPROV & "'"
   CtaCbleRelaProveedorXDefectoEnDetalle% = XVALO(VALOBADA("FLEXCRL", "VALCONTROL", CONDI$, "NOMESS"))
End Function

Function GastosAplicados(NUFACTU$, NPROV, MONEX%)
       'SE CONVIERTE A $ POR EL TIPO DE CAMBIO DE LA FACTURA MADRE
       CONDI$ = "FACTURAPROV= '" & NUFACTU$ & "'and  TIPO_FACTURA = 'M' AND Nume_Cli = " & NPROV

       
       
       VALOCAMBIO# = XVALO(VALOBADA("FACTURAYGASTOS", "TipoCam", CONDI$, "NOMESS"))
       'OBTIEN EL IMPORTE EN PESOS SI MONEX ES <> A PESOS TRANSFORMA DOLAR
       CONDI$ = "FACTURAPROV_ASOC = '" & NUFACTU$ & "' AND MoneVal = " & 1 & " and  TIPO_FACTURA = 'G' AND Nume_Cli_Asoc = " & NPROV
       GASTOS_APLICADOS_PESOS# = XVALO(VALOBADA("FACTURAYGASTOS", "Sum(IMPO_GASTO)", CONDI$, "NOMESS"))
       If MONEX% <> 1 Then
          GASTOS_APLICADOS_CONVERTIDO_A_DOLARES# = GASTOS_APLICADOS_PESOS# / VALOCAMBIO#
          GASTOS_APLICADOS_PESOS# = GASTOS_APLICADOS_CONVERTIDO_A_DOLARES#
       End If
       '
       'SI SON U$S Y LA MONEX ES PESOS LO TRANSFORMA A PESOS
       CONDI$ = "FACTURAPROV_ASOC = '" & NUFACTU$ & "' AND MoneVal = " & 2 & " and  TIPO_FACTURA = 'G' AND Nume_Cli_Asoc = " & NPROV
       GASTOS_APLICADOS_DOLAR# = XVALO(VALOBADA("FACTURAYGASTOS", "Sum(IMPO_GASTO)", CONDI$, "NOMESS"))
       If MONEX% = 1 Then
          GASTOS_APLICADOS_CONVERTIDO_A_PESOS# = GASTOS_APLICADOS_DOLAR# * VALOCAMBIO#
          GASTOS_APLICADOS_DOLAR# = GASTOS_APLICADOS_CONVERTIDO_A_PESOS#
       End If
       
       '
       GastosAplicados = GASTOS_APLICADOS_PESOS# + GASTOS_APLICADOS_DOLAR#
       
End Function
Function GastosAplicados2(NUFACTU$, NPROV, MONEX%)
'ESTA FUNCION DEVUELVE LOS GASTOS  APLICADOS A DIFERENTES FACTURAS QUE SE REALIZO CON ESTA FACTURA PASADA COMO PARAMETRO
'LA OTRA FUNCION DEVUELVE EL GASTO QUE SE LE APLICO A LA FACTURA PASADA COMO PARAMETRO
       'SE CONVIERTE A $ POR EL TIPO DE CAMBIO DE LA FACTURA MADRE
       CONDI$ = "FACTURAPROV= '" & NUFACTU$ & "'and  TIPO_FACTURA = 'M' AND Nume_Cli = " & NPROV
       VALOCAMBIO# = XVALO(VALOBADA("FACTURAYGASTOS", "TipoCam", CONDI$, "NOMESS"))
       'OBTIEN EL IMPORTE EN PESOS SI MONEX ES <> A PESOS TRANSFORMA DOLAR
       CONDI$ = "FACTURAPROV = '" & NUFACTU$ & "' AND MoneVal = " & 1 & " and  TIPO_FACTURA = 'G'"
       GASTOS_APLICADOS_PESOS# = XVALO(VALOBADA("FACTURAYGASTOS", "Sum(IMPO_GASTO)", CONDI$, "NOMESS"))
       If MONEX% <> 1 Then
          GASTOS_APLICADOS_CONVERTIDO_A_DOLARES# = GASTOS_APLICADOS_PESOS# / VALOCAMBIO#
          GASTOS_APLICADOS_PESOS# = GASTOS_APLICADOS_CONVERTIDO_A_DOLARES#
       End If
       '
       'SI SON U$S Y LA MONEX ES PESOS LO TRANSFORMA A PESOS
       CONDI$ = "FACTURAPROV_ASOC = '" & NUFACTU$ & "' AND MoneVal = " & 2 & " and  TIPO_FACTURA = 'G' "
       GASTOS_APLICADOS_DOLAR# = XVALO(VALOBADA("FACTURAYGASTOS", "Sum(IMPO_GASTO)", CONDI$, "NOMESS"))
       If MONEX% = 1 Then
          GASTOS_APLICADOS_CONVERTIDO_A_PESOS# = GASTOS_APLICADOS_DOLAR# * VALOCAMBIO#
          GASTOS_APLICADOS_DOLAR# = GASTOS_APLICADOS_CONVERTIDO_A_PESOS#
       End If
       '
       GastosAplicados2 = GASTOS_APLICADOS_PESOS# + GASTOS_APLICADOS_DOLAR#
       
End Function

Function PROVEEDOR_SEGUN_PEDIDO(NPED&, NUMEORDPED%, CI1%)
    'DEVUELVE EL NUMERO DE PROVEEDOR DEL PEDIDO EL MISMO VIENE DE LA COTIZACION DE AHP
    CONDI$ = "NROPED " & NPED& & "   AND CODIINT = " & CI1%
    Condi1$ = " AND NUORITEM = " & NUMEORDPED%
    'SI NO DEVUELVE POR LA CONDICION, BUSCA FILTRANDO SOLO POR CODIGO Y NROPEDIDO.
    PROVEEDOR_SEGUN_PEDIDO = XVALO(VALOBADA("PEDDETA", "NUM_PROV", CONDI$ + Condi1$, "NOMESS"))
    If PROVEEDOR_SEGUN_PEDIDO < 1 Then
       PROVEEDOR_SEGUN_PEDIDO = XVALO(VALOBADA("PEDDETA", "NUM_PROV", CONDI$, "NOMESS"))
    End If
    '
End Function
Function ProxNumSolCompraPedido&()
   ProxNumSolCompraPedido& = 1 + XVALO(VALOBADA("PEDICOMPRA", "MAX(NumSolCompra)", "NumSolCompra > 0", "NOMESS"))
End Function
Function cliente_SEGUN_PEDIDO(NPED&)
    'DEVUELVE EL NUMERO DE PROVEEDOR DEL PEDIDO EL MISMO VIENE DE LA COTIZACION DE AHP
    CONDI$ = "NROPED =" & NPED&
    cliente_SEGUN_PEDIDO = XVALO(VALOBADA("pedenca", "NroClie", CONDI$, "NOMESS"))
    '
End Function
Function TipoCambioCompra(LOTE$)
    TipoCambioCompra = 0
    CONDI$ = "IDENLOTE = '" & LOTE$ & "'"
    NUOC& = XVALO(VALOBADA("BOLRECEP", "NumOC", CONDI$, "NOMESS"))
    CI1% = XVALO(VALOBADA("BOLRECEP", "COD_INTE", CONDI$, "NOMESS"))
    If NUOC& > 0 Then
       CONDI$ = "NUME_OCOM = " & NUOC&
       TipoCambioCompra = XVALO(VALOBADA("OCOMENCA", "COTIZA_DOLAR", CONDI$, "NOMESS"))
    End If
    '
End Function

Function PRECIOVENTA_X_FORMULA(COSTO1#, VIA%, CI%, CALCULAR_VALOR_SOBRE_FOB%, Optional NPROV, Optional EDITADO%)
   'EN COSTO1# YA VIENE CON EL COSTO NACIONALIZADO, COMO BANDERA TIENE EL PARAMETRO 'CALCULAR_VALOR_SOBRE_FOB%' SI ESTA ES > 0 NO LE CALCULA EL COSTOSOBREFOB DE
   'LA FORMULA STANDARD.
'   NPROV = XVALO(NPROV)
'
'*****************************************************  IMPORTANTE  *****************************
'=======================================================================================================================================
'EN LA TABLA PORCCOEF DEBERA TENER GRABADO LOS COEFICIENTES PARA CADA VIA REPETIDO N VECES TANTOS COEFICIENTES TENGA.
'PARA QUE EL CASO QUE UNA VIA Y COEFICIENTE NO TENGA PROVEEDOR ASIGNADO, SI SE DESEA QUE POR DEFECTO CALCULE LOS PORCENTAJES DEBERA _
'TENER EN DICHA TABLA LA INFORMACION PARA CADA VIA SIN PROVEEDOR ESTAS VAN A SER LAS QUE EL SISTEMA VA A TOMRA POR DEFECTO.
'=================================================================================================

'RESUMEN
'-------
'SI ES LOCAL ENTRE EN EL PRIMER CASE Y SI ENCUENTRA EL PORCENTAJE (QUE ES UNO SOLO) POR EL PROVEEDOR LO LEE, SI EXISTE EN LA TABLA CON PROVEEDOR ASIGNADO.
'BUSCA SI HAY PORCENTAJE A NIVEL GENERAL PARA LA VIA LOCAL, SI EXISTE EN LA TABLA SIN PROVEEDOR ASIGNADO
'
'SI EDITADO% = 1 THEN BUSCA SI TIENE PRECIO DE LISTA EDITADO , SI LO TIENE MUESTRA ESTE
' If EDITADO% > 0 Then
'    SQLX$ = " EXECUTE PRECIOVENTAAHP " & NPROV
'    SQLX$ = SQLX$ + ", " & CodiEmp%
'    SQLX$ = SQLX$ + ", " & CI%
'    SQLX$ = SQLX$ + ", " & VIA%
'    Set RS = FLEXCONN.Execute(SQLX$)
'    PREXVENTAX = XVALO(RS.Fields(0))
'    If PREXVENTAX > 0 Then
'      PRECIOVENTA_X_FORMULA = PREXVENTAX
'      Exit Function
'    End If
' End If
    If EDITADO% >= 0 Then
      PREXVENTAX = PRECIODEVENTA_AHP(CI%, NPROV, VIA%)
      If PREXVENTAX > 0.005 Then
        PRECIOVENTA_X_FORMULA = PREXVENTAX
        Exit Function
      End If
    End If
    
   If EXISTE%(LUDAT$ + "PORCCOEF.RFS") < 1 Then Exit Function
   PRECIOVENTA_X_FORMULA = 0
   Select Case VIA%
    Case 1 'CASO VIA LOCAL -
        'CASO VIA LOCAL - SI EN LA TABLA ESTA GUARDADO CON EL PROVEEDOR SELECCIONADO VA  A ENTRAR EN ESTE IF Y VA A TOMAR EL VALOR DENTRO DE ESTA
        If NPROV > 0 Then
            RFF$ = RECFILT$("!PORCCOEF", 6, MKI$(NPROV))
            FIN& = Len(RFF$)
            For UK& = 1 To FIN& Step 4
               U& = CVS(Mid$(RFF$, UK&, 4))
               X$ = REGLEIDO$("!PORCCOEF", U&)
               VIA1% = CVI(Mid$(X$, 1, 2))
               If VIA1% = VIA% Then
                  PORC = CVD(Mid$(X$, 5, 8))
                  PORC = 1 + (PORC / 100)
                  Exit For
               End If
            Next UK&
        End If
'ACA VA A ENTRAR EN LOS CASOS QUE VENGA CON NUMERO DE PROVEEDOR Y NO HAYA ENCONTRADO EL PORCENTAJE EN EL IF ANTERIOR, VA A TOMAR EL INDICE GENERAL DE LA VIA LOCAL
        If PORC < 0.005 Then
            RFF$ = RECFILT$("PORCCOEF", 1, MKI$(VIA%))
            For KKI% = 1 To Len(RFF$) Step 4
             RELI& = CVS(Mid$(RFF$, KKI%, 4))
             If RELI& > 0 Then
               ZZ$ = REGLEIDO$("PORCCOEF", RELI&)
               If CVI(Mid$(ZZ$, 13, 2)) = 0 Then 'VALIDACION DE PROVEEDOR
                 PORC = CVD(Mid$(ZZ$, 5, 8))
                 PORC = 1 + (PORC / 100)
                 Exit For
               End If
             End If
            Next KKI%
        End If
        '
        PRECIOVENTA_X_FORMULA = XVALO(COSTO1#) * PORC  '"SEGUN FORMULA"
        Exit Function
        
        
    Case Else 'RESTO DE LAS VIAS DE PROVEEEDORES DE EXPORTACION
    'ACA VA A BUSCAR LOS QUE MACHEEN CON EL PROVEEDOR
        RFF$ = RECFILT$("PORCCOEF", 1, MKI$(VIA%))
        For KKI% = 1 To Len(RFF$) Step 4
         RELI& = CVS(Mid$(RFF$, KKI%, 4))
         If RELI& > 0 Then
           ZZ$ = REGLEIDO$("PORCCOEF", RELI&)
           If CVI(Mid$(ZZ$, 13, 2)) = NPROV Or NPROV = 0 Then 'VALIDACION DE PROVEEDOR
                If CVI(Mid$(ZZ$, 3, 2)) = 2 Then
                    PORCSOBREFOB = CVD(Mid$(ZZ$, 5, 8))
                    If Abs(PORCSOBREFOB) < 0.005 Then
                       Call COMUNI("Imposible Calcular Precio de Venta Según Fórmula - Falta Porcentaje Sobre Valor Fob")
                       Exit Function
                    End If
                ElseIf CVI(Mid$(ZZ$, 3, 2)) = 3 Then
                    COSTOKG = CVD(Mid$(ZZ$, 5, 8))
                    If Abs(COSTOKG) < 0.005 Then
                       Call COMUNI("Imposible Calcular Precio de Venta Según Fórmula - Falta Costo x Kg.")
                       Exit Function
                    End If
                ElseIf CVI(Mid$(ZZ$, 3, 2)) = 4 Then
                    INVERSO = CVD(Mid$(ZZ$, 5, 8))
                    If Abs(INVERSO) < 0.005 Then
                       Call COMUNI("Imposible Calcular Precio de Venta Según Fórmula - Falta Porcentaje Inverso")
                       Exit Function
                    End If
                End If
           End If
         End If
        Next KKI%
   End Select
   '
   'SE AGREGA ESTE IF PARA EL CASO QUE NO ENTRE EN EL IF DE ARRIBA  -  TOME LOS VALORES GENERALES
   'VUELVE A RECORRER Y EN EL CASO QUE NO TENGA VALOR ALGUNO DE LOS COEFICIENTES SE LOS AGREGUE
   If XVALO(PORCSOBREFOB) = 0 Or XVALO(COSTOKG) = 0 Or XVALO(INVERSO) = 0 Then
        '
        RFF$ = RECFILT$("PORCCOEF", 1, MKI$(VIA%))
        For KKI% = 1 To Len(RFF$) Step 4
         RELI& = CVS(Mid$(RFF$, KKI%, 4))
         If RELI& > 0 Then
           ZZ$ = REGLEIDO$("PORCCOEF", RELI&)
           If CVI(Mid$(ZZ$, 13, 2)) = 0 Then  'VALIDACION DE PROVEEDOR
                If CVI(Mid$(ZZ$, 3, 2)) = 2 And XVALO(PORCSOBREFOB) = 0 Then ' SI NO TIENE VALOR LE AGREGA EL GENERAL
                    PORCSOBREFOB = CVD(Mid$(ZZ$, 5, 8))
                    If Abs(PORCSOBREFOB) < 0.005 Then
                       Call COMUNI("Imposible Calcular Precio de Venta Según Fórmula - Falta Porcentaje Sobre Valor Fob")
                       Exit Function
                    End If
                    
                ElseIf CVI(Mid$(ZZ$, 3, 2)) = 3 And XVALO(COSTOKG) = 0 Then ' SI NO TIENE VALOR LE AGREGA EL GENERAL
                    COSTOKG = CVD(Mid$(ZZ$, 5, 8))
                    If Abs(COSTOKG) < 0.005 Then
                       Call COMUNI("Imposible Calcular Precio de Venta Según Fórmula - Falta Costo x Kg.")
                       Exit Function
                    End If
                ElseIf CVI(Mid$(ZZ$, 3, 2)) = 4 And XVALO(INVERSO) = 0 Then ' SI NO TIENE VALOR LE AGREGA EL GENERAL
                    INVERSO = CVD(Mid$(ZZ$, 5, 8))
                    If Abs(INVERSO) < 0.005 Then
                       Call COMUNI("Imposible Calcular Precio de Venta Según Fórmula - Falta Porcentaje Inverso")
                       Exit Function
                    End If
                End If
           End If
         End If
        Next KKI%
   End If
   
   'PVTA EN USD = (FOB * 1,3150 + KG de peso * 11,0625) / 0,70
   
   PORCSOBREFOB = 1 + (PORCSOBREFOB / 100)
   'SI TIENE GASTOS DE NACIONALIZACION CARGADOS NO DEBE CALCUALAR EL PORCENTAJE SOBRE FOB
   If CALCULAR_VALOR_SOBRE_FOB% > 0 Then
      VALOR1 = XVALO(COSTO1#) * PORCSOBREFOB ' COSTO  * PORCENTAJE SOBRE FOB
   Else
      VALOR1 = XVALO(COSTO1#)
   End If
   
   CONDI$ = "CODINT = " & CI%
   PESOUNITARIO = XVALO(VALOBADA("MASTER", "PESUNI", CONDI$, "NOMESS"))
   VALOR2 = COSTOKG * PESOUNITARIO
   PRECIOVENTA_X_FORMULA = (VALOR1 + VALOR2)
   If INVERSO <> 0 Then
      PRECIOVENTA_X_FORMULA = (VALOR1 + VALOR2) / INVERSO
   End If
   
End Function

Function COTIZACIONCABANC(DOCUMENX$, NC)
  'ESTA FUNCION DEVUELVE SEGUN ELDOCUMENTO EL VALOR DE COTIZACION DE LA FACTURA EN EL CASO QUE EL TIPO DE MONEDA SEA >  0
  'SI LA MONEDA ES 1 DEVUELVE 1
  'DE ESTE MODO PUEDO MULTIPLICAR POR EJEMPLO EL IMPORTE QUE ME DEVUELVE EL MOVISTO POR EL VALOR DE LA FUNCION
  'QUE VA A SER UNO SI ES MONEDA UNO O SI NO LO DEBERA DIVIDIR POR EL VALOR DE COTIZADOLAR
  COTIZACIONCABANC = 1
  DOCU$ = TRIM$(DOCUMENX$)
  If DOCU$ = "" Then Exit Function
  SQLX$ = "SELECT Mone_Emis, COTIZA_DOLAR FROM CAJABANC WITH(NOLOCK)"
  SQLX = SQLX$ + " WHERE CODICOM_LAR = '" & DOCU$ & "' AND NuClien = " & NC
  SQLX$ = SQLX$ + " AND MONE_EMIS < 2"
  Set rs = READSET(SQLX$)
  If Not rs.EOF Then COTIZACIONCABANC = XVALO(rs!COTIZA_DOLAR)
  rs.Close
  Set rs = Nothing
End Function
Function ULTIMACOMPRA_SEGUN_PROV_VIA&(CI1%, NPROV, VIAENTR%)
    ULTIMACOMPRA_SEGUN_PROV_VIA& = 0
    SQLX$ = " SELECT TOP 1 OC.Nume_Ocom  FROM OCOMENCA OC WITH(NOLOCK) INNER JOIN OCOMDETA OD"
    SQLX$ = SQLX$ + " ON OC.NUME_OCOM = OD.NUME_OCOM"
    SQLX$ = SQLX$ + " WHERE OC.Npro_Ocom = " & NPROV
    SQLX$ = SQLX$ + " AND OC.Via_Entrega = " & VIAENTR%
    SQLX$ = SQLX$ + " AND  OD.Cod_Inte = " & CI1%
    SQLX$ = SQLX$ + " AND OD.Cant_Rec > 0"
    SQLX$ = SQLX$ + " ORDER BY OC.Nume_Ocom  DESC"
    Set rs = READSET(SQLX$)
    With rs
      If Not .EOF Then ULTIMACOMPRA_SEGUN_PROV_VIA& = XVALO(!nume_ocom)
    End With
    rs.Close
    Set rs = Nothing
End Function
Function CostoUltimaCompra#(CI1%)
    'DEVUELVE EL VALOR DE LA ULTIMA ULTIMA ORDEN DE COMPRA REALIZADA Y QUE YA SE HAYA RECIBIDO
    'TENGO CAMPOS DE MAS PERO CREO QUE VOY A USAR PARA AHP
    'SE RESPETO EL MISMO FORMATO QUE TIENE LA GRILLA DE AMASAHP SOLO SE ORDENO POR FECHA DE EMISION
    EPAC$ = UCase$(EMPREAC$)

    NprovMineparts = XVALO(Control("CFGMINEP", "PRFEMINE"))
    SQLX$ = " SELECT top 1  o.Cod_Inte,o.Npro_Ocom, o.RaSo_Prov, oc.Via_Entrega,Pre_Unit,o.Femi_Ocom ,o.Mone_Emis,COTIZA_DOLAR  "
    SQLX$ = SQLX$ + " FROM OCOMDETA o WITH(NOLOCK) INNER JOIN OCOMENCA OC ON O.nume_ocom = OC.nume_ocom "
    SQLX$ = SQLX$ + " WHERE O.Cod_Inte = " & CI1% & "  and O.stat_ocom < 9 and O.stat_ocom >= 0 and  Cant_Rec > 0  "
    SQLX$ = SQLX$ + " And O.stat_ocom < 9 And O.stat_ocom >= 0 And Cant_Rec > 0"
    If InStr(EPAC$, "AHP") > 0 Then SQLX$ = SQLX$ + " And o.Npro_Ocom <> " & NprovMineparts
    SQLX$ = SQLX$ + " group by O.COD_INTE, O.NPRO_OCOM,O.RaSo_Prov,OC.VIA_ENTREGA,Pre_Unit,o.Femi_Ocom ,o.Mone_Emis,COTIZA_DOLAR  "
    SQLX$ = SQLX$ + " ORDER BY o.Femi_Ocom DESC"
    Set rs = READSET(SQLX$)
    With rs
        CostoUltimaCompra# = XVALO(rs!Pre_Unit)
        If XVALO(rs!Mone_Emis) < 2 Then
           COTI# = XVALO(rs!COTIZA_DOLAR) 'SI ES MONEDA PESOS CALCULO EL VALOR EN DOLARES
           If COTI# > 0 Then
                CostoUltimaCompra# = CostoUltimaCompra# / COTI#
           Else
                CostoUltimaCompra# = 0
           End If
        End If
    End With

End Function
Function PROV_SegCtaCble(CUENCONTA%)

    PROV_SegCtaCble = 0
    SQLX$ = "SELECT * FROM Proved12 WITH(NOLOCK) "
    SQLX$ = SQLX$ + " WHERE CuMa_Cli = " & CUENCONTA%
    Set rs = READSET(SQLX$)
    With rs
     If Not (.EOF And .BOF) Then
       PROV_SegCtaCble = XVALO(!nume_cli)
     End If
    End With
    
End Function

Function CtaRetIIBBSegunProvyProvincia%(DenoABreviadaMecobra$)
        'BUSCA POR LA DENOMINACION ABREVIADA LA CUENTA CONTABLE EN MECOBRA
        'BUSCA EL PROVEEDOR DONDE LA CUENTA MADRE  DEL PROVEEDOR
        'OBTENIENDO LA PROVINCIA A PARTIR DEL NPROVEEDOR Y BUSCANDO EN LA TABLA PROVINB.DAT LA CUENTA CONTABLE ASOCIADA DE IIBB
        ' MAS EL CODIGO DE REGIMEN
        CtaRetIIBBSegunProvyProvincia% = 0
        XX$ = FILTERGETRECORD$("MECOBRA", 2, DenoABreviadaMecobra$) ' LEO EL REGISTRO SEGUN EL TEXTO QUE TIENE EN APLICOBRA QUE ES EL QUE  PONE EN EL LABEL 17
        CTACBLEMECOBRAX$ = TRIM$(Mid$(XX$, 45, 12)) 'LEE LA CUENTA CONTABLE DE MECOBRA
        CUECO% = CUECOINT%(CTACBLEMECOBRAX$)
        NC = XVALO(VALOBADA("PROVED12", "Nume_Cli", "CuMa_Cli = " & CUECO%, "NOMESS")) 'BUSCA EL PROVEEDOR SEGUN LA CUENTA CONTABLE
        If NC > 0 Then
          PROVINCIAX$ = PROVCLI$(-1 * NC)
          X = CodigoRegimenRetIBXProvincia(PROVINCIAX$, RET_REIB%) 'EL PARAMETRO RET_REIB% ES DE RETORNO
          CtaRetIIBBSegunProvyProvincia% = RET_REIB%
        End If
End Function

'Function EsCtaProveedor(CUECO%)
' 'devuelve la cuanta de retencion de ingresos brutos segun la cuenta contable  de la tabla provinb.dat
'        'BUSCA EL PROVEEDOR DONDE LA CUENTA MADRE  DEL PROVEEDOR
'        'OBTENIENDO LA PROVINCIA A PARTIR DEL NPROVEEDOR Y BUSCANDO EN LA TABLA PROVINB.DAT LA CUENTA CONTABLE ASOCIADA DE IIBB
'        ' MAS EL CODIGO DE REGIMEN
'        CtaRetIIBBSegunProvyProvincia% = 0
'
'        NC = XVALO(VALOBADA("PROVED12", "Nume_Cli", "CuMa_Cli = " & CUECO%, "NOMESS")) 'BUSCA EL PROVEEDOR SEGUN LA CUENTA CONTABLE
'        If NC > 0 Then
'          PROVINCIAX$ = PROVCLI$(-1 * NC)
'          x = CodigoRegimenRetIBXProvincia(PROVINCIAX$, RET_REIB%) 'EL PARAMETRO RET_REIB% ES DE RETORNO
'          CtaRetIIBBSegunProvyProvincia% = RET_REIB%
'        End If
'End Function

Function PROV_SEGUN_EMECOBRA%(DenoABreviadaMecobra$)
        'BUSCA LA CUENTA CONTABLE SEGUN LA DENOMINACION ABREVIADA DEL MECOBRA
        'CON ESTO BUSCA LA CUENTA CONTABLE QUE ESTAN EN MECOBRA
        'Y MIRA SI TIENE UN PROVEEDOR CON ESA CUENTA CONTABLE SI ES ASI TRAE EL PROVEEDOR
        NC = 0
        XX$ = FILTERGETRECORD$("MECOBRA", 2, TRIM$(DenoABreviadaMecobra$)) ' LEO EL REGISTRO SEGUN EL TEXTO QUE TIENE EN APLICOBRA QUE ES EL QUE  PONE EN EL LABEL 17
        CTACBLEMECOBRAX$ = TRIM$(Mid$(XX$, 45, 12)) 'LEE LA CUENTA CONTABLE DE MECOBRA
        CUECO% = CUECOINT%(CTACBLEMECOBRAX$)
        NC = XVALO(VALOBADA("PROVED12", "Nume_Cli", "CuMa_Cli = " & CUECO%, "NOMESS")) 'BUSCA EL PROVEEDOR SEGUN LA CUENTA CONTABLE
        If NC > 0 Then
          PROV_SEGUN_EMECOBRA% = NC
        End If

End Function
Sub IMPRIME_OCOMPRA(NPX&)
    DOCUNU& = 0
    TIDOCUM$ = "Nota de Pedido"
    EPAC$ = UCase$(EMPREAC$)
    If InStr(EPAC$, "AHP") > 0 Then TIDOCUM$ = "ORDEN DE COMPRA"

    NUDOCU$ = TRIM$(Str$(NPX&))
    Call MUESTRADOC(DOCUNU&, "", TIDOCUM$, NUDOCU$)

End Sub
Function PrimerClienteActivo&(Optional CLIPRO$)
   CLIPRO$ = TRIM$(UCase$(CLIPRO))
   Tablax$ = "DEUDOR12"
   If CLIPRO$ = "PROVEEDOR" Then Tablax$ = "PROVED12"
      
   SQLX$ = "SELECT NUME_CLI FROM " & Tablax$ & "  with(nolock)"
   SQLX$ = SQLX$ + "WHERE Stat_Cli > 0 "
   SQLX$ = SQLX$ + "ORDER BY Nume_Cli "
   Set ClieTemp = READSET(SQLX$)
   '
   PrimerClienteActivo& = 0
   With ClieTemp
     On Error Resume Next
     .MoveFirst
     If Err < 1 Then PrimerClienteActivo& = XVALO(!nume_cli)
     On Error GoTo 0
   End With
   
   ClieTemp.Close
   Set ClienTemp = Nothing

End Function
Function UltimoClienteActivo&(Optional CLIPRO$)
   CLIPRO$ = TRIM$(UCase$(CLIPRO))
   Tablax$ = "DEUDOR12"
   If CLIPRO$ = "PROVEEDOR" Then Tablax$ = "PROVED12"

   SQLX$ = "SELECT NUME_CLI FROM " & Tablax$ & "  with(nolock)"
   SQLX$ = SQLX$ + "WHERE Stat_Cli > 0 "
   SQLX$ = SQLX$ + "ORDER BY NUME_CLI desc "
   Set ClieTemp = READSET(SQLX$)
   '
   With ClieTemp
     On Error Resume Next
     .MoveFirst
     If Err < 1 Then UltimoClienteActivo& = XVALO(!nume_cli)
     On Error GoTo 0
   End With
   ClieTemp.Close
   Set ClienTemp = Nothing

End Function
Function OrdenCompraPed$(NPED&)
 CONDI$ = ("NROPED = " & NPED&)
 OrdenCompraPed$ = VALOBADA("PEDENCA", "NROOCOM", CONDI$, "NOMESS")
End Function

Function ListaSeleccionClientes(Optional RetornaSegundoCampoEco As String, Optional Refresh As Boolean = False) As Long

    TituloLista$ = "Maestro de clientes"
    If TRIM$(FRMZELECHO.Caption) = TituloLista$ And Refresh = False Then GoTo 50
    
    Call CARGA_ENCABEZADO(FRMZELECHO.msf2, "Cuenta/I10;Razón Social/A64", FRMZELECHO)
    Call CARGA_ENCABEZADO(FRMZELECHO.MSF1, "Cuenta/I10;Razón Social/A64", FRMZELECHO)
    FRMZELECHO.Caption = TituloLista$
    
    Dim obj As New RECORXET
    Set RS1 = obj.RS_MAESTRODECLIENTES()
    Call Carga_MSF_Recordset(RS1, "Cuenta/I10;Razon Social/A64", FRMZELECHO.msf2, 1)
    On Error Resume Next
    Set obj = Nothing
    RS1.Close
    Set RS1 = Nothing
    On Error GoTo 0
        
    FRMZELECHO.Width = 10000
50  FRMZELECHO.Show 1

    NumeroCliente$ = TRIM(RESP_ZELE_VECT(1))
    ListaSeleccionClientes = XVALO(NumeroCliente$)
    RetornaSegundoCampoEco = TRIM(RESP_ZELE_VECT(2))
    
End Function
Function ListaSeleccionClientesFiltradoXVendedor(Optional RetornaSegundoCampoEco As String, Optional Refresh As Boolean = False, Optional NroVend As Integer) As Long

    TituloLista$ = "Maestro de clientes (Vendedor: " & NroVend & ")"
    If TRIM$(FRMZELECHO.Caption) = TituloLista$ And Refresh = False Then GoTo 50

    Call CARGA_ENCABEZADO(FRMZELECHO.msf2, "Cuenta/I10;Razón Social/A64", FRMZELECHO)
    Call CARGA_ENCABEZADO(FRMZELECHO.MSF1, "Cuenta/I10;Razón Social/A64", FRMZELECHO)
    FRMZELECHO.Caption = TituloLista$

    Dim obj As New RECORXET
    Set RS1 = obj.RS_MAESTRODECLIENTESxVENDEDOR(NroVend)
    Call Carga_MSF_Recordset(RS1, "Cuenta/I10;Razon Social/A64", FRMZELECHO.msf2, 1)
    On Error Resume Next
    Set obj = Nothing
    RS1.Close
    Set RS1 = Nothing
    On Error GoTo 0

    FRMZELECHO.Width = 10000
50  FRMZELECHO.Show 1

    NumeroCliente$ = TRIM(RESP_ZELE_VECT(1))
    ListaSeleccionClientesFiltradoXVendedor = XVALO(NumeroCliente$)
    RetornaSegundoCampoEco = TRIM(RESP_ZELE_VECT(2))
    
End Function

Function NCLLICONTEPLAST(VENDESEGUNUSUARIO%)
   
   'NCLIE = ListaSeleccionClientes("", True)
   '1 SI TIENE DERECHOS O ES ADIMINISTRADOR VE TODOS (PRESENTAR CON NOMESS)
   '2 SI ES VENDEDOR VE LOS FILTRADOS POR VENDEDOR
   '3 SI NO ES VENDEDOR MENSAJE QUE NO PUEDE VISAULIZAR CLIENTES QUE SOLICITE DERECHO. VISUALIZAR DERECHO
   
   USNU% = USNBR% Mod 100
   If USNU% = 1 Or DERACCE%("VECLICTP/NOMESS", "VISUALIZAR TODOS LOS CLIENTES") >= 2 Then 'CON NOMESS NO PRESENTA MENSAJE
      NCLIE = ListaSeleccionClientes("", True)
      GoTo 99
   End If
   
   If VENDESEGUNUSUARIO% > 0 Then  ' ES VENDEDOR SOLO VE SUS CLIENTES
      NCLIE = ListaSeleccionClientesFiltradoXVendedor("", True, VENDESEGUNUSUARIO%)
      GoTo 99
   Else
      If DERACCE%("VECLICTP", "VER TODOS LOS CLIENTES ") < 2 Then 'ACA PRESENTA EL MENSAJE
        NCLIE = 0
        GoTo 99
      End If
   End If
99 NCLLICONTEPLAST = NCLIE
End Function
Function ListaSeleccionDocumVersionados&(NroOriginal&, REFERENCIA$)

    TituloLista$ = "Imagenes Digitales Versionadas"
    If TRIM$(FRMZELECHO.Caption) = TituloLista$ Then GoTo 50
    Campos$ = "Nro Original/f18;Fecha/D8;Identificador/a16;Versión/A8"
    Call CARGA_ENCABEZADO(FRMZELECHO.msf2, Campos$, FRMZELECHO)
    Call CARGA_ENCABEZADO(FRMZELECHO.MSF1, Campos$, FRMZELECHO)
    FRMZELECHO.Caption = TituloLista$
    
    Dim obj As New RECORXET
    Set RS1 = obj.RS_DocumVersionados(NroOriginal&, REFERENCIA$)
    Call Carga_MSF_Recordset(RS1, Campos$, FRMZELECHO.msf2, 1)
    On Error Resume Next
    Set obj = Nothing
    RS1.Close
    Set RS1 = Nothing
    On Error GoTo 0
        
    FRMZELECHO.Width = 10000
50  FRMZELECHO.Show 1

    NroDocum& = XVALO(RESP_ZELE_VECT(3))
    ListaSeleccionDocumVersionados& = NroDocum&
    
End Function

Sub GRABAR_VERSION_IMAGEN(NRO_ORIG&, NUMEDOCUM&, REFERENCIA$)
        If NUMEDOCUM& < 1 Then Exit Sub
        If NRO_ORIG& < 1 Then Exit Sub
        
        'LA REFERENCIA DEBE SER EL TIPODOC$ (LA MISMA QUE APLICA EN PDOC_TBL)
        'NRO_ORIG& ES EL NRO DEL PEDIDO POR EJEMPLO
        'NumDocPdoc ES EL NRO DEL CAMPO NUMEDOC CON EL CUAL SE GUARDO EN EL PDOC_TBL (ES UN VALOR DE RETORNO DEL PRINTDOC)
        'EL NUMERO DE VERSION ARRANCA DE 0 POR ESO LO MANEJO CON CANTEGISTRO
        CONDI$ = "NroOrig = " & NRO_ORIG& & " AND Referencia = '" & REFERENCIA$ & "'"
        CREG& = CantRegistros("IMAG_VERSION", CONDI$, "NOMESS")
        NPROX_VERSION& = CREG&
        
        SQLX$ = "INSERT INTO IMAG_VERSION "
        SQLX$ = SQLX$ + " (NroOrig,Referencia,NumDocPdoc,NVERSION,FechEmis,CODIEMPR)"
        SQLX$ = SQLX$ + " VALUES(" & NRO_ORIG&
        SQLX$ = SQLX$ + " ,'" & REFERENCIA$ & "'"
        SQLX$ = SQLX$ + " ," & NUMEDOCUM&
        SQLX$ = SQLX$ + " ," & NPROX_VERSION&
        SQLX$ = SQLX$ + " ,'" & FETEXCOR1$(HOY(HO$)) & "'"
        SQLX$ = SQLX$ + " ," & CodiEmp%
        
        SQLX$ = SQLX$ + ")"
        FLEXCONN.Execute (SQLX$)
End Sub

Function ListaSeleccionProveedores(Optional RetornaSegundoCampoEco As String, Optional Refresh As Boolean = False) As Long

    TituloLista$ = "Maestro de proveedores"
    If TRIM$(FRMZELECHO.Caption) = TituloLista$ And Refresh = False Then GoTo 50
    
    Call CARGA_ENCABEZADO(FRMZELECHO.msf2, "Cuenta/I10;Razón Social/A64", FRMZELECHO)
    Call CARGA_ENCABEZADO(FRMZELECHO.MSF1, "Cuenta/I10;Razón Social/A64", FRMZELECHO)
    FRMZELECHO.Caption = TituloLista$
    
    Dim obj As New RECORXET
    Set RS1 = obj.RS_MAESTRODEPROVEEDORES()
    Call Carga_MSF_Recordset(RS1, "Cuenta/I10;Razon Social/A64", FRMZELECHO.msf2, 1)
    On Error Resume Next
    Set obj = Nothing
    RS1.Close
    Set RS1 = Nothing
    On Error GoTo 0
        
    FRMZELECHO.Width = 10000
50  FRMZELECHO.Show 1
    
    NumeroProveedor$ = TRIM(RESP_ZELE_VECT(1))
    ListaSeleccionProveedores = XVALO(NumeroProveedor$)
    RetornaSegundoCampoEco = TRIM(RESP_ZELE_VECT(2))
    
End Function

Function CodigoCtaContablePorGrupoVenta(grupoVenta%) As Integer

    CuentaContable% = XVALO(VALOBADA("GRUVECUECON", "CODINCUE", "GRUCOVEN=" & CStr(grupoVenta%)))
    
    CodigoCtaContablePorGrupoVenta = CuentaContable%
End Function

Sub ESTAXART(Optional CI1%, Optional LISTAOCOSTO$)
   If CI1% < 1 Then
     Call SELECHO("MASTER")
     COD$ = TRIM$(VALACT1$("MASTER"))
     CI1% = CODINT%(COD$)
   End If
   
   If CI1% < 1 Then Exit Sub
   '
   LISTAOCOSTO$ = TRIM$(LISTAOCOSTO$)
   If LISTAOCOSTO$ = "" Then LISTAOCOSTO$ = "LISTA"
20 Call DESHASFECHA(DESDEX%, HASTAX%, PERIO$)
   If PERIO$ = "" Then Exit Sub

   DESFE = FETEXCOR1$(DESDEX%)
   HASFE = FETEXCOR1$(HASTAX%, "HASTA")
   '
   'VALIDO SI HAY NOVEDADES CON MAS DE UNA MONEDA
   CONDI$ = " Fech_Mod  >=  '" & DESFE & " '"
   CONDI$ = CONDI$ + " and Fech_Mod  <=  '" & HASFE & " '"
   CONDI$ = CONDI$ + "  AND Lista_Repos = '" & LISTAOCOSTO$ & "'"
   CONDI$ = CONDI$ + "  AND CODINT = " & CI1%
   CONDI2$ = "  AND MONEDA <= 1"
   CONDI3$ = "  AND MONEDA > 1"
   If CantRegistros("NOVEDADES", CONDI$ + CONDI2$, "NOMESS") > 0 Then PESO% = 1
   If CantRegistros("NOVEDADES", CONDI$ + CONDI3$, "NOMESS") > 0 Then DOL% = 1
   'SI HAY MAS DE UNA MONEDA PREGUNTO CUAL QUIERE VER
   If PESO% + DOL% > 1 Then
60    OPX% = COMALTER%("Se Han Detectado Cambios de Precios en 2 Tipos de Monedas\\Ver Cambios en Pesos\Ver Cambios en Dolares")
      If OPX% < 1 Or OPX% > 2 Then GoTo 20
      TIPOMONE% = OPX%
   ' SI HAY SOLO UNA MONEDA MUESTRA DIRECTAMENTE ESTA.
   Else
      TIPOMONE% = 1
      If DOL% > 0 Then TIPOMONE% = 2
   End If
   SQLX$ = "SELECT  * FROM NOVEDADES WITH(NOLOCK) "
   SQLX$ = SQLX$ + " WHERE   " & CONDI$
   If TIPOMONE% = 1 Then SQLX$ = SQLX$ + " AND MONEDA <= " & TIPOMONE%
   If TIPOMONE% = 2 Then SQLX$ = SQLX$ + " AND MONEDA >= " & TIPOMONE%
   SQLX$ = SQLX$ + " ORDER BY FECH_MOD"
   '
   INFO_GRILLA07.MSF.Redraw = False
   Screen.MousePointer = 11
   'Call CARGA_ENCABEZADO_X_ART
     With INFO_GRILLA07
        .MSF.row = 1
        .MSF.Cols = 10
        NUCOLU% = 1: .TEPRUEBA = "90": .MSF.ColWidth(NUCOLU% - 1) = .TEPRUEBA.Width: .MSF.ColAlignment(NUCOLU% - 1) = 6: .MSF.TextMatrix(0, NUCOLU% - 1) = ""
        '
        NUCOLU% = 2: .TEPRUEBA = "1234567890": .MSF.ColWidth(NUCOLU% - 1) = .TEPRUEBA.Width: .MSF.ColAlignment(NUCOLU% - 1) = 1: .MSF.TextMatrix(0, NUCOLU% - 1) = "Fecha"
        NUCOLU% = 3: .TEPRUEBA = "12345678902": .MSF.ColWidth(NUCOLU% - 1) = .TEPRUEBA.Width: .MSF.ColAlignment(NUCOLU% - 1) = 6: .MSF.TextMatrix(0, NUCOLU% - 1) = "$.Ant"
        NUCOLU% = 4: .TEPRUEBA = "12345678902": .MSF.ColWidth(NUCOLU% - 1) = .TEPRUEBA.Width: .MSF.ColAlignment(NUCOLU% - 1) = 6: .MSF.TextMatrix(0, NUCOLU% - 1) = "$.Nuevo"
        NUCOLU% = 5: .TEPRUEBA = "12345678": .MSF.ColWidth(NUCOLU% - 1) = .TEPRUEBA.Width: .MSF.ColAlignment(NUCOLU% - 1) = 6: .MSF.TextMatrix(0, NUCOLU% - 1) = "%"
        '
        NUCOLU% = 6: .TEPRUEBA = "123456": .MSF.ColWidth(NUCOLU% - 1) = .TEPRUEBA.Width: .MSF.ColAlignment(NUCOLU% - 1) = 6: .MSF.TextMatrix(0, NUCOLU% - 1) = "Mon"
        NUCOLU% = 7: .TEPRUEBA = "1234567890123456": .MSF.ColWidth(NUCOLU% - 1) = .TEPRUEBA.Width: .MSF.ColAlignment(NUCOLU% - 1) = 1: .MSF.TextMatrix(0, NUCOLU% - 1) = "Usuario"
        NUCOLU% = 8: .TEPRUEBA = "1234567890123456789+123456789+123456789+123456789+123456789+": .MSF.ColWidth(NUCOLU% - 1) = .TEPRUEBA.Width: .MSF.ColAlignment(NUCOLU% - 1) = 1: .MSF.TextMatrix(0, NUCOLU% - 1) = "Observación"
        '
        NUCOLU% = 9: .TEPRUEBA = "12345678902": .MSF.ColWidth(NUCOLU% - 1) = .TEPRUEBA.Width: .MSF.ColAlignment(NUCOLU% - 1) = 6: .MSF.TextMatrix(0, NUCOLU% - 1) = "M.Original"
        NUCOLU% = 10: .TEPRUEBA = "12345678902": .MSF.ColWidth(NUCOLU% - 1) = .TEPRUEBA.Width: .MSF.ColAlignment(NUCOLU% - 1) = 6: .MSF.TextMatrix(0, NUCOLU% - 1) = "Cotiz."
       
     End With
        '
    

   REG& = 0
   Set rs = READSET(SQLX$)
   With rs
    Do While Not .EOF
      REG& = REG& + 1
      '
      INFO_GRILLA07.MSF.Rows = REG& + 1
      Call TRACE(24, REG&, FIN&)
      '
      INFO_GRILLA07.MSF.TextMatrix(REG&, 1) = FECHATEX$(CVINT(!Fech_Mod))
      P_ANT# = XVALO(!Prec_Ant)
      P_NEW# = XVALO(!Prec_New)
      INFO_GRILLA07.MSF.TextMatrix(REG&, 2) = FORMATNUM$(P_ANT#, "F12.2")
      INFO_GRILLA07.MSF.TextMatrix(REG&, 3) = FORMATNUM$(P_NEW#, "F12.2")
       PJE = ((P_NEW# / P_ANT#) - 1) * 100
      INFO_GRILLA07.MSF.TextMatrix(REG&, 4) = FORMATNUM$(PJE, "F6.2")
      INFO_GRILLA07.MSF.TextMatrix(REG&, 5) = FORMATNUM$(TIPOMONE%, "F2.0")
       
      INFO_GRILLA07.MSF.TextMatrix(REG&, 6) = TRIM$(ECOARCH$("FLEXUSR", MKI$(XVALO(!Nume_Usuar))))
      INFO_GRILLA07.MSF.TextMatrix(REG&, 7) = SAFETEXT$(!OBSERVACION)

      .MoveNext
    Loop
   End With
   
   rs.Close
   Set rs = Nothing
   Screen.MousePointer = 1
   If REG& < 1 Then
      Call COMUNI("No Se Detectaron Cambios en el Rango Seleccionado")
      GoTo 20
   End If
   
   INFO_GRILLA07.MSF.Redraw = True
   
   If TIPOMONE% < 1 Then TIPOMONE% = 1
   DEMONE$ = DEMONECO$(TIPOMONE%)
   INFO_GRILLA07.Caption = "Cambios de Precios en " & DEMONE$ & "  - Art.: " & CODEXT$(CI1%) & " - Desde: " & FECHATEX$(DESDEX%) & " - Hasta: " & FECHATEX$(HASTAX%)
   Call COLOREAR_GRILLASQL(INFO_GRILLA07.MSF, &HEFF1F1)
   '
   INFO_GRILLA07.Command7.Visible = True
   INFO_GRILLA07.CI1X = CI1%
   '
   INFO_GRILLA07.Show 1
   '
   If PESO% + DOL% > 1 Then GoTo 60
   '
   GoTo 20
   
End Sub

Sub ESTAXPRV()
   
10  Call SELECHO("PROVED1")
    NCX1% = XVALO(VALACT1$("PROVED1"))
    If NCX1% < 1 Then
       OPX% = COMALTER%("No se Seleccionó Ningún Proveedor\\Cancelar\Ver Novedades de Todos Los Artículos")
       If OPX% < 2 Then Exit Sub
    End If
   '
20 Call DESHASFECHA(DESDEX%, HASTAX%, PERIO$)
   If PERIO$ = "" Then GoTo 10

   DESFE = FETEXCOR1$(DESDEX%)
   HASFE = FETEXCOR1$(HASTAX%, "HASTA")
   '
   
   'VALIDO SI HAY NOVEDADES CON MAS DE UNA MONEDA
   CONDI$ = " Fech_Mod  >=  '" & DESFE & " '"
   CONDI$ = CONDI$ + " and Fech_Mod  <=  '" & HASFE & " '"
   CONDI$ = CONDI$ + "  AND Lista_Repos = 'COSTO'"
   If NCX1% > o Then CONDI$ = CONDI$ + "  AND NumProv  = " & NCX1%
'   CONDI2$ = "  AND MONEDA <= 1"
'   condi3$ = "  AND MONEDA > 1"
   PESO% = 0: DOL% = 0
'   If CantRegistros("NOVEDADES", CONDI$ + CONDI2$, "NOMESS") > 0 Then PESO% = 1
'   If CantRegistros("NOVEDADES", CONDI$ + condi3$, "NOMESS") > 0 Then DOL% = 1
'   'SI HAY MAS DE UNA MONEDA PREGUNTO CUAL QUIERE VER
'   If PESO% + DOL% > 1 Then
'60    OPX% = COMALTER%("Se Han Detectado Cambios de Precios en 2 Tipos de Monedas\\Ver Cambios en Pesos\Ver Cambios en Dolares")
'      If OPX% < 1 Or OPX% > 2 Then GoTo 20
'      TIPOMONE% = OPX%
'   ' SI HAY SOLO UNA MONEDA MUESTRA DIRECTAMENTE ESTA.
'   Else
'      TIPOMONE% = 1
'      If DOL% > 0 Then TIPOMONE% = 2
'   End If
'   '
'   REG& = 0
     With INFO_GRILLAJM
        .MSF.Rows = 1
        .MSF.Cols = 9
        NUCOLU% = 1: .TEPRUEBA = "90": .MSF.ColWidth(NUCOLU% - 1) = .TEPRUEBA.Width: .MSF.ColAlignment(NUCOLU% - 1) = 6: .MSF.TextMatrix(0, NUCOLU% - 1) = ""
        NUCOLU% = 2: .TEPRUEBA = "12345678901234567890": .MSF.ColWidth(NUCOLU% - 1) = .TEPRUEBA.Width: .MSF.ColAlignment(NUCOLU% - 1) = 1: .MSF.TextMatrix(0, NUCOLU% - 1) = "Còdigo"
        NUCOLU% = 3: .TEPRUEBA = "123456789012345678901234567890123": .MSF.ColWidth(NUCOLU% - 1) = .TEPRUEBA.Width: .MSF.ColAlignment(NUCOLU% - 1) = 1: .MSF.TextMatrix(0, NUCOLU% - 1) = "Descripciòn"
        NUCOLU% = 4: .TEPRUEBA = "1234567890": .MSF.ColWidth(NUCOLU% - 1) = .TEPRUEBA.Width: .MSF.ColAlignment(NUCOLU% - 1) = 1: .MSF.TextMatrix(0, NUCOLU% - 1) = "Fecha 1"
        NUCOLU% = 5: .TEPRUEBA = "12345678902": .MSF.ColWidth(NUCOLU% - 1) = .TEPRUEBA.Width: .MSF.ColAlignment(NUCOLU% - 1) = 6: .MSF.TextMatrix(0, NUCOLU% - 1) = "$.Ant"
        NUCOLU% = 6: .TEPRUEBA = "1234567890": .MSF.ColWidth(NUCOLU% - 1) = .TEPRUEBA.Width: .MSF.ColAlignment(NUCOLU% - 1) = 1: .MSF.TextMatrix(0, NUCOLU% - 1) = "Fecha 2"
        NUCOLU% = 7: .TEPRUEBA = "12345678902": .MSF.ColWidth(NUCOLU% - 1) = .TEPRUEBA.Width: .MSF.ColAlignment(NUCOLU% - 1) = 6: .MSF.TextMatrix(0, NUCOLU% - 1) = "$.Nuevo"
        NUCOLU% = 8: .TEPRUEBA = "12345678": .MSF.ColWidth(NUCOLU% - 1) = .TEPRUEBA.Width: .MSF.ColAlignment(NUCOLU% - 1) = 6: .MSF.TextMatrix(0, NUCOLU% - 1) = "%"
        NUCOLU% = 9: .TEPRUEBA = "12345": .MSF.ColWidth(NUCOLU% - 1) = .TEPRUEBA.Width: .MSF.ColAlignment(NUCOLU% - 1) = 6: .MSF.TextMatrix(0, NUCOLU% - 1) = "Mon"
        
        '
    End With
   
   '
   SQLX$ = "SELECT CODINT  FROM NOVEDADES WITH(NOLOCK) "
   SQLX$ = SQLX$ + " WHERE   " & CONDI$
'   If TIPOMONE% = 1 Then SQLX$ = SQLX$ + " AND MONEDA <= " & TIPOMONE%
'   If TIPOMONE% = 2 Then SQLX$ = SQLX$ + " AND MONEDA >= " & TIPOMONE%
   SQLX$ = SQLX$ + " GROUP BY CODINT "
   '
   INFO_GRILLAJM.MSF.Rows = 1
   REG& = 0
   INFO_GRILLAJM.MSF.Redraw = False
   Screen.MousePointer = 11
   Set rs = READSET(SQLX$)
   With rs
    Do While Not .EOF
      
      '
      '
      'LEO LA INFORMACION EN PARTICULAR POR CODIGO TENIENDO EN CUENTA ORDENAMIENTOS
      CI1% = XVALO(!CODINT)
      
      'TOMO LA FECHA PRIMERA DENTRO DEL RANGO
      Condi1$ = "CODINT = " & CI1%
      Condi1$ = Condi1$ + " and Fech_Mod  >=  '" & DESFE & " '"
      Condi1$ = Condi1$ + " and Fech_Mod  <=  '" & HASFE & " '"
      Condi1$ = Condi1$ + "  AND NumProv  = " & NCX1%
      Condi1$ = Condi1$ + " ORDER BY Fech_Mod ASC"
      FE1% = CVINT(VALOBADA("NOVEDADES", "FECH_MOD", Condi1$, "NOMESS"))
      If FE1% < 1 Then
        GoTo 60
      End If
      '
      'LEO EL  PRIMER PRECIO
      CONDI2$ = "CODINT = " & CI1%
      CONDI2$ = CONDI2$ + " and Fech_Mod  >=  '" & FETEXCOR1$(FE1%) & " '"
      CONDI2$ = CONDI2$ + " and Fech_Mod  <=  '" & FETEXCOR1$(FE1%, "HASTA") & " '"
      CONDI2$ = CONDI2$ + "  AND NumProv  = " & NCX1%
      CONDI2$ = CONDI2$ + " ORDER BY Fech_Mod ASC"
      PRECIO1# = XVALO(VALOBADA("NOVEDADES", "Prec_Ant", CONDI2$, "NOMESS"))
      '
      'LEO LA ULTIAMA FECHA DENTRO DEL RANGO
      CONDI3$ = "CODINT = " & CI1%
      CONDI3$ = CONDI3$ + " and Fech_Mod  >=  '" & DESFE & " '"
      CONDI3$ = CONDI3$ + " and Fech_Mod  <=  '" & HASFE & " '"
      CONDI3$ = CONDI3$ + "  AND NumProv  = " & NCX1%
      CONDI3$ = CONDI3$ + " ORDER BY Fech_Mod DESC"
      FE2% = CVINT(VALOBADA("NOVEDADES", "FECH_MOD", CONDI3$, "NOMESS"))
      '
      'LEO EL ULTIMO  PRECIO
      CONDI4$ = "CODINT = " & CI1%
      CONDI4$ = CONDI4$ + " and Fech_Mod  >=  '" & FETEXCOR1$(FE2%) & " '"
      CONDI4$ = CONDI4$ + " and Fech_Mod  <=  '" & FETEXCOR1$(FE2%, "HASTA") & " '"
      CONDI4$ = CONDI4$ + "  AND NumProv  = " & NCX1%
      CONDI4$ = CONDI4$ + " ORDER BY Fech_Mod DESC"

      PRECIO2# = XVALO(VALOBADA("NOVEDADES", "Prec_New", CONDI4$, "NOMESS"))
      '
      'CALCULO EL PORCENTAJE DE MODIFICACION ENTRE EL PRIMER VALOR Y EL ULTIMO SEGUN RANGO
      If PRECIO1# > 0.005 Then
         PJE = ((PRECIO2# / PRECIO1#) - 1) * 100
      ElseIf PRECIO2# < 0.005 Then
         PJE = -100
      'CASO QUE SEA INGRESO DE UN ITEM QUE SE CARGA EL PRECIO POR PRIMERA VEZ  X Q' CUANDO SE GRABE EN LA TABLA NOVEDADES NO VA A TRAER NINGUN COSTO ANTERIOR
      ElseIf PRECIO1# < 0.005 Then
         PJE = 0
      End If
      '
      REG& = REG& + 1
      INFO_GRILLAJM.MSF.Rows = REG& + 1

      INFO_GRILLAJM.MSF.TextMatrix(REG&, 1) = CODEXT$(CI1%)
      INFO_GRILLAJM.MSF.TextMatrix(REG&, 2) = DESCRIT0$(CI1%)
      INFO_GRILLAJM.MSF.TextMatrix(REG&, 3) = FECHATEX$(FE1%)
      INFO_GRILLAJM.MSF.TextMatrix(REG&, 4) = FORMATNUM$(PRECIO1#, "F12.2")
      INFO_GRILLAJM.MSF.TextMatrix(REG&, 5) = FECHATEX$(FE2%)
      INFO_GRILLAJM.MSF.TextMatrix(REG&, 6) = FORMATNUM$(PRECIO2#, "F12.2")
      INFO_GRILLAJM.MSF.TextMatrix(REG&, 7) = FORMATNUM$(PJE, "F9.2")
'      INFO_GRILLAJM.MSF.TextMatrix(REG&, 8) = FORMATNUM$(TIPOMONE%, "F2.0")
      '
60    .MoveNext
    Loop
   End With
   
   rs.Close
   Set rs = Nothing
   Screen.MousePointer = 1

   INFO_GRILLAJM.MSF.Redraw = True
   
   If REG& < 1 Then
      Call COMUNI("No Se Detectaron Cambios en el Rango Seleccionado")
      GoTo 20
   End If

   If TIPOMONE% < 1 Then TIPOMONE% = 1
   DEMONE$ = DEMONECO$(TIPOMONE%)

   INFO_GRILLAJM.Caption = "Cambios de Precios en " & DEMONE$ & " - Prov.: " & rasocli(NCX1% * -1) & " - Desde: " & FECHATEX$(DESDEX%) & " - Hasta: " & FECHATEX$(HASTAX%)
   Call COLOREAR_GRILLASQL(INFO_GRILLAJM.MSF, &HC0E0FF)
   INFO_GRILLAJM.NPROV% = NCX1%
   INFO_GRILLAJM.Command7.Visible = True
   INFO_GRILLAJM.Show 1
   
   
   GoTo 10
End Sub

Function CtaCliSegunCuit(CUITX$)
   CUITX$ = REPLACAR(CUITX$, "-", "")
   CONDI$ = "Cuit_Cli = '" & CUITX$ & "' OR Cuit_Cli = '" & CUITX$ & "' OR Cuit_Cli = '" & FormatearCuitValido$(CUITX$) & "'"
   NC = XVALO(VALOBADA("DEUDOR12", "NUME_CLI", CONDI$, "NOMESS"))
   CtaCliSegunCuit = NC
End Function

Function CodPostal$(NC)
   'DOMICILIO DE ENTREGA DE DEUDOR12
   On Error Resume Next
   If STATCLI%(NC) <> 0 Then CodPostal$ = TRIM$(ClienTemp!Lentre3_Cli)
End Function
Function ALUMONOSEL(Optional SOLODEBAJA%)

    Call CARGA_ENCABEZADO(FRMZELECHO.msf2, "Legajo/f7;Nombres/A24;Documento/A13;Teléfono/A24;Mail/A32;Celular/A20", FRMZELECHO)
    Call CARGA_ENCABEZADO(FRMZELECHO.MSF1, "Legajo/f7;Nombres/A24;Documento/A13;Teléfono/A24;Mail/A32;Celular/A20", FRMZELECHO)
    FRMZELECHO.Caption = "TABLA DE ALULMNOS"
    If SOLODEBAJA% > 0 Then FRMZELECHO.Caption = "ALUMNOS DADOS DE BAJA"

    Dim obj As New RECORXET
    Set RS1 = obj.RS_ALUMNOS(SOLODEBAJA%)
    Call Carga_MSF_Recordset(RS1, "Legajo/f6;Nombres/A24;Documento/A10;Teléfono/A12;Mail/A24;Celular/A12", FRMZELECHO.msf2)
    Set obj = Nothing
    RS1.Close
    Set RS1 = Nothing
    FRMZELECHO.Width = 14800
    FRMZELECHO.Show 1
    ALUMONOSEL = XVALO(RESP_ZELE_VECT(1))
    
End Function
Function PROFESORSEL(Optional SOLODEBAJA%)

    Call CARGA_ENCABEZADO(FRMZELECHO.msf2, "Legajo/f7;Nombres/A24;Documento/A13;Teléfono/A24;Mail/A32;Celular/A20", FRMZELECHO)
    Call CARGA_ENCABEZADO(FRMZELECHO.MSF1, "Legajo/f7;Nombres/A24;Documento/A13;Teléfono/A24;Mail/A32;Celular/A20", FRMZELECHO)
    FRMZELECHO.Caption = "TABLA DE PROFESORES ACTIVOS"
    If SOLODEBAJA% > 0 Then FRMZELECHO.Caption = "PROFESORES DADOS DE BAJA"
    Dim obj As New RECORXET
    Set RS1 = obj.RS_PROFESORES(SOLODEBAJA%)
    Call Carga_MSF_Recordset(RS1, "Legajo/f6;Nombres/A24;Documento/A10;Teléfono/A18;Mail/A28;Celular/A12", FRMZELECHO.msf2)
    Set obj = Nothing
    RS1.Close
    Set RS1 = Nothing
    FRMZELECHO.Width = 14800
    FRMZELECHO.Show 1
    PROFESORSEL = XVALO(RESP_ZELE_VECT(1))
    
End Function

Sub ALTAPROVINCIASSQL(CODINUM&, codProvinciaFlx$, nombreP$)
        SQLX$ = "INSERT INTO TAPROVINCIA "
        SQLX$ = SQLX$ + " (CODINUM,CODITEX,DESCRI_PROV)"
        SQLX$ = SQLX$ + " VALUES(" & CODINUM&
        SQLX$ = SQLX$ + " ,'" & codProvinciaFlx$ & "'"
        SQLX$ = SQLX$ + " ,'" & nombreP$ & "'"
        SQLX$ = SQLX$ + ")"
        FLEXCONN.Execute (SQLX$)
End Sub

Function PROVELOTENUM(NROLOTE&)
    SQLX$ = " EXECUTE ProveedorSegunlote  " & NROLOTE&
    Dim rs As New ADODB.Recordset
    Set rs = FLEXCONN.Execute(SQLX$)
    PROVELOTENUM = XVALO(rs!NumProv)
    rs.Close: Set rs = Nothing
End Function
Function DeudaProforma(NC)
   'DEVUELVE EL TOTAL DEL PENDIENTE DE FACTURAS PROFORMAS NO FACTURADAS
   'SI NC > 0 FILTRA POR CLIENTE
   CONDI$ = "(STATUS = 0 OR STATUS IS NULL ) AND STATUS < 9 AND (FacturaReal = '' OR FacturaReal IS NULL) "
   If NC > 0 Then CONDI$ = CONDI$ + " AND NUCLIEN = " & NC
   DeudaProforma = XVALO(VALOBADA("PROFORMTL", "SUM(ImpoTot)", CONDI$, "NOMESS"))
End Function

Function LETRAPROFORMA$(NPROF&)
   CONDI$ = "NUMECOMP =  " & NPROF&
   LETRAPROFORMA$ = VALOBADA("PROFORMTL", "VARIMOVIM", CONDI$, "NOMESS")
End Function
Function EsProformaXDetalle%(NPROF&)
   CONDI$ = "NUMECOMP =  " & NPROF&
   EsProformaXDetalle% = CantRegistros&("DETAPROFO", CONDI$, "NOMESS")
End Function
Function PedidoTeineProforma&(NP&)
   CONDI$ = "NROPED =  " & NP&
   PedidoTeineProforma& = XVALO(VALOBADA("PROFORMTL", "NUMECOMP", CONDI$, "NOMESS"))
End Function

Function ListaSeleccionProvincia() As Long
    ListaSeleccionProvincia = 0
    TituloLista$ = "Maestro de Provincias"
    Campos$ = "Id Prov/f10;Provincia/A64;Cód.Alfa/A9"
    Call CARGA_ENCABEZADO(FRMZELECHO.msf2, Campos$, FRMZELECHO)
    Call CARGA_ENCABEZADO(FRMZELECHO.MSF1, Campos$, FRMZELECHO)
    FRMZELECHO.Caption = TituloLista$
    
    Dim obj As New RECORXET
    Set RS1 = obj.RS_GENERAL("PROVINCIA", "ID,NOMBRE,CODIGO31662", "ID > 0 ORDER BY ID ")
    Call Carga_MSF_Recordset(RS1, Campos$, FRMZELECHO.msf2, 1)
    On Error Resume Next
    Set obj = Nothing
    RS1.Close
    Set RS1 = Nothing
    On Error GoTo 0
        
    FRMZELECHO.Width = 12000
50  FRMZELECHO.Show 1
    If XVALO(RESP_ZELE_VECT(1)) > 0 Then
       ListaSeleccionProvincia = XVALO(RESP_ZELE_VECT(1))
    End If
    
End Function
Function ListaSeleccionLocalidades(Optional ProvNum%) As Long
    ListaSeleccionLocalidades = 0
    TituloLista$ = "Maestro de Localidades"
    Campos$ = "Id Loc/f10;Localidad/A64;Cód.Postal/f9"
    Call CARGA_ENCABEZADO(FRMZELECHO.msf2, Campos$, FRMZELECHO)
    Call CARGA_ENCABEZADO(FRMZELECHO.MSF1, Campos$, FRMZELECHO)
    FRMZELECHO.Caption = TituloLista$
    
    Dim obj As New RECORXET
    If ProvNum% > 0 Then
       CONDI$ = "PROVINCIA_ID = " & ProvNum%
    Else
       CONDI$ = "ID > 0"
    End If
    CONDI$ = CONDI$ + " ORDER BY ID"
    Set RS1 = obj.RS_GENERAL("LOCALIDAD", "ID,NOMBRE,CODIGOPOSTAL", CONDI$)
    Call Carga_MSF_Recordset(RS1, Campos$, FRMZELECHO.msf2, 1)
    On Error Resume Next
    Set obj = Nothing
    RS1.Close
    Set RS1 = Nothing
    On Error GoTo 0
        
    FRMZELECHO.Width = 12000
50  FRMZELECHO.Show 1
    If XVALO(RESP_ZELE_VECT(1)) > 0 Then
      ListaSeleccionLocalidades = XVALO(RESP_ZELE_VECT(1))
    End If
    
End Function

Function TipoCambioFacturaProveedor(NPROVE, DOCUFACTURA$, Optional RETORNO_MONEMIS%)
    RETORNO_MONEMIS% = 0
    TipoCambioFacturaProveedor = 0
    'DEVUELVE EL TIPO DE CAMBIO SEGUN LA FACTURA Y EL NUMERO DEL PROVEEDOR
    'CREADA PARA AHP PARA TOMAR CON QUE COTIZACION SE GUARDO LA FACTURA Y USAR ESTE DATO EN EL CALCULO DE PORCENTAJE DE NACIONALIZACION
    SQLX$ = "Select mone_emis, tipo_cam  from cajabanc WITH(NOLOCK) "
    SQLX$ = SQLX$ + " WHERE tipomovim  = 'fcom' "
    SQLX$ = SQLX$ + " AND nuprove = " & NPROVE
    SQLX$ = SQLX$ + " AND CODICOM_LAR = '" & DOCUFACTURA$ & "'"
    Set rs = READSET(SQLX$)
    RETORNO_MONEMIS% = XVALO(rs!Mone_Emis)
    TipoCambioFacturaProveedor = XVALO(rs!Tipo_Cam)
    rs.Close
    Set rs = Nothing
End Function
Function CUTACBLESEGUNLEYENDA$(LEYENDA$)
   'SE REALIZA LA CONSULTA DE ESTA MANERA POR EL TEMA QUE EN LA TABLA DATASQL, ESTA GUARDANDO
   '3 TABLAS ASOCIADAS Y NO ESTAN EN LA MISMA COLUMNA LA CUENTA CONTABLE ASOCIADA.
   'CASHFLOW
   CUTACBLESEGUNLEYENDA$ = ""
   SQLX$ = "select upper(CAMPO1) AS CAMPO1,upper(CAMPO2)as CAMPO2, upper(CAMPO3) AS CAMPO3, upper(CAMPO4) AS CAMPO4, upper(CAMPO5) AS CAMPO5, lista from datasql WITH(NOLOCK) where lista = 'RUBFIBAN' or lista = 'REFEINGA' OR LISTA = 'VENCIMIVA'"
   Set rs = READSET(SQLX$)
   With rs
      Do While Not .EOF
        Select Case SAFETEXT$(!Lista)
          Case "REFEINGA"
            CTABLE$ = SAFETEXT$(!CAMPO3)
            LEYLEIDO$ = SAFETEXT$(!campo2)
          Case "RUBFIBAN"
            CTABLE$ = SAFETEXT$(!campo2)
            LEYLEIDO$ = SAFETEXT$(!CAMPO1)
          Case "VENCIMIVA"
            CTABLE$ = SAFETEXT$(!CAMPO5)
            LEYLEIDO$ = SAFETEXT$(!CAMPO4)
        End Select
        If LEYLEIDO$ = UCase(LEYENDA$) Then
           CUTACBLESEGUNLEYENDA$ = TRIM$(CTABLE$)
           Exit Do
        End If
        .MoveNext
      Loop
   
   End With
   
   
End Function

Function CtaProv_ViaEnt_PjeNacionalizacion%(LOTE$, RetProvExt, RetViaEnt, RetPjeNacionalizacion, RetCoeficienteNacionalizacion)
   '******   AHP    ******
   '******   DEVUELVE 1 SI ENCONTRO DATOS    ******
   'ESTA FUNCION DEVUELVE SEGUN EL LOTE LOS VALORES QUE SE GUARDA CUANDO SE CARGA LA FACTURA DEL PROVEEDOR
   'MINEPART PARADO EN AHP,EL RetPjeNacionalizacion LO GUARDA COMO TAL Y POR CALCULO TAMBIEN LO DEVUELVE COMO COEFICIENTE
   CtaProv_ViaEnt_PjeNacionalizacion% = 0 ' POR DEFECTO
   '
   SQLX$ = "SELECT ProvExterior,PjeNacionaliz,ViaExterior"
   SQLX$ = SQLX$ + " FROM BOLRECEP WITH(NOLOCK)"
   SQLX$ = SQLX$ + " WHERE IDENLOTE = '" & TRIM$(LOTE$) & "'"
   Set rs = READSET(SQLX$)
   With rs
     If Not .EOF Then
       RetProvExt = XVALO(!ProvExterior)
       RetViaEnt = XVALO(!ViaExterior)
       RetPjeNacionalizacion = XVALO(!PjeNacionaliz)
       RetCoeficienteNacionalizacion = 1 + (RetPjeNacionalizacion / 100)
       CtaProv_ViaEnt_PjeNacionalizacion% = 1
     End If
   End With
   rs.Close
   Set rs = Nothing
   '
   
End Function
Function CantidadPendienteSegunPedidoCodigo(NPED&, CI1%)
    'DEVUELVE LA CANTIDAD PENDIENTE DE UN CODIGO EN UN PEDIDO
    CONDI$ = "NROPED =" & NPED& & " AND CODIINT = " & CI1%
    CONDI$ = CONDI$ + " AND CODIINT = " & CI1%
    CONDI$ = CONDI$ + " AND STATUS < 8 AND STATUS >= 0 "
    CantidadPendienteSegunPedidoCodigo = XVALO(VALOBADA("PEDDETA", "SUM(CANPED -CantEnt)", CONDI$, "NOMESS"))
    '
End Function
Function CantidadPedidaItemEnPedido&(NPED&, CI1%)
    'existencia de UN CODIGO EN UN PEDIDO - devuelve la cantidad pedida
    CONDI$ = "NROPED =" & NPED& & " AND CODIINT = " & CI1%
    CONDI$ = CONDI$ + " AND CODIINT = " & CI1%
    'CONDI$ = CONDI$ + " AND STATUS < 8 AND STATUS >= 0 "
    CantidadPedidaItemEnPedido& = XVALO(VALOBADA("PEDDETA", "SUM(CANPED)", CONDI$, "NOMESS"))
    '
End Function

Function ExisteNumeroProveedorEnConfiguraciónCompraInterna&(NC)
    'DEVUELVE EL ID (DATASQL_ID) SI LO ENCUENTRA EN DATASQL, SI NO DEVULEVE 0
    If NC < 1 Then ExisteNumeroProveedorEnConfiguraciónCompraInterna& = 0
    'EL CAMPO4 ES EL NRO DE PROVEEDOR
    CONDI$ = "LISTA = 'CTEPROCOMPRA' AND CAMPO6 <> 'LOCAL' AND CAMPO4 = '" & TRIM$(Str(Abs(NC))) & "'"
    ExisteNumeroProveedorEnConfiguraciónCompraInterna& = XVALO(VALOBADA("DATASQL", "DATASQL_ID", CONDI$, "NOMESS"))
End Function

Function NroClienteSegunIDompraInterna&(ID&)
    'DEVUELVE EL NRO DE CLIENTE ASOCIADO AL PROVEEDOR SEGUN EL ID (DATASQL_ID).
    If ID& < 1 Then NroClienteSegunIDompraInterna& = 0
    
    CONDI$ = "LISTA = 'CTEPROCOMPRA' AND DATASQL_ID =" & ID&
    NroClienteSegunIDompraInterna& = XVALO(VALOBADA("DATASQL", "CAMPO2", CONDI$, "NOMESS"))
End Function

Function NroClientePropioCompraInterna&()
    'DEVUELVE EL NRO DE CLIENTE PROPIO EL QUE EN LA REFERENCIA ESTA ROTULADO COMO LOCAL (DATASQL_ID).
    NroClientePropioCompraInterna& = 0
    
    CONDI$ = "LISTA = 'CTEPROCOMPRA' AND UPPER(CAMPO6) ='LOCAL'"
    NroClientePropioCompraInterna& = XVALO(VALOBADA("DATASQL", "CAMPO2", CONDI$, "NOMESS"))
End Function

Function RutaCarpetaSegunNCCompraInterna$(NC)
    'DEVUELVE la RUTA SEGUN EL CLIENTE
    RutaCarpetaSegunNCCompraInterna$ = 0
    NCL1$ = TRIM$(Str(NC))
    
    CONDI$ = "LISTA = 'CTEPROCOMPRA' AND CAMPO2 ='" & NCL1$ & "'"
    RutaCarpetaSegunNCCompraInterna$ = VALOBADA("DATASQL", "CAMPO7", CONDI$, "NOMESS")
End Function
Function RutaCarpetaSegunNProvCompraInterna$(NPROV)
    'DEVUELVE la RUTA SEGUN EL PROVEEDOR
    RutaCarpetaSegunNProvCompraInterna$ = ""
    NPROV1$ = TRIM$(Str(NPROV))
    
    CONDI$ = "LISTA = 'CTEPROCOMPRA' AND CAMPO4 ='" & NPROV1$ & "'"
    RutaCarpetaSegunNProvCompraInterna$ = VALOBADA("DATASQL", "CAMPO7", CONDI$, "NOMESS")
End Function
Function NClteSegunNProvCompraInterna&(NPROV)
    'DEVUELVE la RUTA SEGUN EL PROVEEDOR
    NClteSegunNProvCompraInterna& = 0
    NPROV1$ = TRIM$(Str(NPROV))
    'EL CAMPO4 ES EL NRO DE cliente
    CONDI$ = "LISTA = 'CTEPROCOMPRA' AND CAMPO4 ='" & NPROV1$ & "'"
    NClteSegunNProvCompraInterna& = XVALO(VALOBADA("DATASQL", "CAMPO2", CONDI$, "NOMESS"))
End Function
Function RutaCarpetaLocalCompraInterna$()
    'DEVUELVE la ruta del que esta marcado como local (en cada sucursal debera estar marcada la que corresponda local)
    RutaCarpetaLocalCompraInterna$ = ""
    
    CONDI$ = "LISTA = 'CTEPROCOMPRA' AND UPPER(CAMPO6) ='LOCAL'"
    RutaCarpetaLocalCompraInterna$ = VALOBADA("DATASQL", "CAMPO7", CONDI$, "NOMESS")
End Function
Function StringDeNroClientesComprasInternas$()
    CONDI$ = "LISTA = 'CTEPROCOMPRA'"
    StringDeNroClientesComprasInternas$ = StringAggFlex$("DATASQL", CONDI$, "CAMPO2", ",")
End Function
Function StringDeNroProveedoresComprasInternas$()
    CONDI$ = "LISTA = 'CTEPROCOMPRA'"
    StringDeNroProveedoresComprasInternas$ = StringAggFlex$("DATASQL", CONDI$, "CAMPO4", ",")
End Function
Function RutaCarpetaRemitosNeuquen$()
    'DEVUELVE la ruta del que en el campo 6 esta guardada como 'REMITOSNEUQUEN'
    'comparte la misma tanto en neuquen como en avellaneda.
    RutaCarpetaRemitosNeuquen$ = ""
    
    CONDI$ = "LISTA = 'CTEPROCOMPRA' AND UPPER(CAMPO6) ='REMITOSNEUQUEN'"
    RutaCarpetaRemitosNeuquen$ = TRIM$(VALOBADA("DATASQL", "CAMPO7", CONDI$, "NOMESS"))
End Function

Function NroClienteSegunPedido&(NROPED&)
    'DEVUELVE EL NRO DE CLIENTE SEGUN EL NRO DE PEDIDO
    NroClienteSegunPedido& = 0
    CONDI$ = "NROPED = " & NROPED&
    NroClienteSegunPedido& = XVALO(VALOBADA("PEDENCA", "NROCLIE", CONDI$, "NOMESS"))
End Function

