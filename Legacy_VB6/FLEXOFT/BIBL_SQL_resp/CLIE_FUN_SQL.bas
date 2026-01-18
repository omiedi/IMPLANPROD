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
Function TIENE_FACTURAS&(NC, DIAX%, FEULT_FACT%)
   'DEVUELVE 1 O MAS  SI COMPRO DENTRO DE LOS DIA PASADO COMO PARAMETRO AL DIA DE HOY
   'DEVUELVE 0 SI NO COMPRO DENTRO DE LOS DIA PASADO COMO PARAMETRO AL DIA DE HOY
    DES% = DIANTE(HOY(HO$), DIAX%)
    DESDE = FETEXCOR1$(DES%)
    HASTA = FETEXCOR1$(HOY(HO$))
    '
    CONDI$ = "TIPOMOVIM = 'FVEN' AND FECHEMISI >= '" & DESDE & "'"
    CONDI$ = CONDI$ + " AND FECHEMISI <= '" & HASTA & "'"
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
       NPRX& = Abs(ClienTemp!NUME_CLI)
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
       Exit Function
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

Function RASOCLI$(NC)
   On Error Resume Next
   If STATCLI%(NC) <> 0 Then RASOCLI$ = TRIM$(ClienTemp!Raso_Cli)
End Function
'
Function DOMICLI$(NC)
   On Error Resume Next
   If STATCLI%(NC) <> 0 Then DOMICLI$ = TRIM$(ClienTemp!Domi_Cli)
End Function
'
Function CPOSCLI$(NC)
   On Error Resume Next
   If STATCLI%(NC) <> 0 Then CPOSCLI$ = TRIM$(ClienTemp!Cpos_Cli)
End Function
'
Function LOCACLI$(NC)
   On Error Resume Next
   If STATCLI%(NC) <> 0 Then LOCACLI$ = TRIM$(ClienTemp!Loca_Cli)
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
   If STATCLI%(NC) <> 0 Then TEL1CLI$ = TRIM$(ClienTemp!TELE_Cli)
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
     FANTCLI$ = UCase$(RASOCLI$(NC))
     PPXX% = InStr(FANTCLI$, " ")
     If PPXX% > 1 Then
       FANTCLI$ = Left$(FANTCLI$, PPXX% - 1)
     End If
   End If
End Function
'
Function FAXCLI$(NC)
   On Error Resume Next
   If STATCLI%(NC) <> 0 Then FAXCLI$ = TRIM$(ClienTemp!FAXI_Cli)
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
   If STATCLI%(NC) <> 0 Then CONTACTCLI$ = TRIM$(ClienTemp!Ctac_Cli)
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
Function CUITCLI$(NC)
   On Error Resume Next
   If STATCLI%(NC) <> 0 Then CUITCLI$ = TRIM$(ClienTemp!Cuit_Cli)
End Function
'
Function NIBRCLI$(NC)
   On Error Resume Next
   If STATCLI%(NC) <> 0 Then NIBRCLI$ = TRIM$(ClienTemp!Nibr_Cli)
End Function
'
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
    COEFI = 1
    DESCX$ = TRIM$(REPLACAR$(TDESCUEN$, "%", ""))
    '
    Do While Len(DESCX$) > 0
       PPXX% = InStr(DESCX$, "+")
       If PPXX% < 1 Then PPXX% = 1 + Len(DESCX$)
       AKX1$ = Left$(DESCX$, PPXX% - 1)
       DESCX$ = TRIM$(Mid$(DESCX$, PPXX% + 1))
       '
       PDC = XVALO(AKX1$)
       COEFI = COEFI * (1 - PDC / 100)
    Loop
    '
99  PORCEDESCU = 100 * (1 - COEFI)
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
      STRSQL = "SELECT * FROM Deudor12 WHERE Gru_Div > 0"
      Dim BCLIEN As ADODB.Recordset
      Set BCLIEN = FLEXCONN.Execute(FILTSQL$(STRSQL))
      
      
      With BCLIEN
        On Error Resume Next
        .MoveFirst
        If Err < 1 Then
           On Error GoTo 0
           Do While Not .EOF
             GRD% = XVALO(!Gru_Div)
             ORGD%(GRD%) = XVALO(!NUME_CLI)
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
   ORDECHEQ$ = RASOCLI$(NC)
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
   If STATCLI%(NC) <> 0 Then ANOTACLI$ = ClienTemp!Nota_Cli
End Function
'
Sub NOTACLI(NC, MODO%)
   '
   ' MODO% = 1    abre automaticamente si hay nota para el cliente
   ' MODO% = 2    se pidió apertura por botón de acceso
   '
   If RASOCLI$(NC) = "" Then Exit Sub
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
         TXTANOTADOR$ = SAFETEXT$(ClienTemp!Nota_Cli)
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
Function DELIPRE$(CL%)
   DX$ = "PRECIOS LIBRES"
   If CL% > 0 Then If CL% < 256 Then DX$ = ECOARCH$("LISTAS", Chr$(CL%))
   If DX$ = "" Then Beep: CL% = 0: DX$ = "PRECIOS LIBRES"
   DELIPRE$ = DX$
End Function
'
Function DECONPA$(CL%)
   DX$ = "CONTADO"
   If CL% > 0 Then If CL% < 256 Then DX$ = ECOARCH$("CONPAG", Chr$(CL%))
   If DX$ = "" Then Beep: CL% = 0: DX$ = "CONTADO"
   DECONPA$ = DX$
End Function
'
Function DEVENDE$(CL)
   DX$ = "GERENCIA"
   If CL > 0 Then DX$ = ECOARCH$("VENDEDOR", Chr$(VENDCLI%(CL)))
   If DX$ = "" Then Beep: CL = 0: DX$ = "GERENCIA"
   DEVENDE$ = DX$
End Function
'
Function DEGRUCLI$(CL%)
   DX$ = "NO CLASIFICADO"
   If CL% > 0 Then DX$ = ECOARCH$("GRUCOCLI", MKI$(CL%))
   If DX$ = "" Then Beep: CL% = 0: DX$ = "NO CLASIFICADO"
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
Function VALICUIT%(CUIT$)
    '
    CU$ = TRIM$(CUIT$)
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
    For j% = 1 To 13
        If j% <> 3 Then
            If j% <> 12 Then
                VX% = Asc(Mid$(CU$, j%, 1))
                If VX% < 48 Or VX% > 57 Then GoTo 999                 ' no valida
            End If
        End If
    Next j%
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
    CUIT$ = CU$

999 VALICUIT% = VC%
    '
End Function

Sub CARDATPRO(NC%)
    '
    Screen.MousePointer = 11
    '
    MODIPRO.Text2.TEXT = TEL2CLI$(NC%)
    MODIPRO.Text3.TEXT = RASOCLI$(NC%)
    MODIPRO.Text4.TEXT = DOMICLI$(NC%)
    MODIPRO.Text5.TEXT = CPOSCLI$(NC%)
    MODIPRO.Text6.TEXT = LOCACLI$(NC%)
    MODIPRO.Text10.TEXT = TEL1CLI$(NC%)
    MODIPRO.Text15.TEXT = CUITCLI$(NC%)
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
    For I& = ULTREG&("LISTAS") To 1 Step -1
       XX$ = REGLEIDO$("LISTAS", I&)
       If Asc(Left$(XX$, 1)) > 0 Then GoTo 59
       Call GRAREG("LISTAS", String$(128, 0), I&)
    Next I&
59  Call CIERRARCH("LISTAS")
    '
     For I& = ULTREG&("CONPAG") To 1 Step -1
       XX$ = REGLEIDO$("CONPAG", I&)
       If Asc(Left$(XX$, 1)) > 0 Then GoTo 69
       Call GRAREG("CONPAG", String$(128, 0), I&)
    Next I&
69  Call CIERRARCH("CONPAG")
    '
    For I& = ULTREG&("VENDEDOR") To 1 Step -1
       XX$ = REGLEIDO$("VENDEDOR", I&)
       If Asc(Left$(XX$, 1)) > 0 Then GoTo 79
       Call GRAREG("VENDEDOR", String$(128, 0), I&)
    Next I&
79  Call CIERRARCH("VENDEDOR")
    '
     For I& = ULTREG&("PROVINC") To 1 Step -1
       XX$ = REGLEIDO$("PROVINC", I&)
       If Asc(Left$(XX$, 1)) > 32 Then GoTo 89
       Call GRAREG("PROVINC", String$(128, 0), I&)
    Next I&
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
    N1% = FreeFile
    Open ARX1$ For Random Access Read Write Lock Read Write As #N1% Len = 32
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
    Lock #N1%
    Lock #N2%
    '
    FORSELEC.LISORT.Clear
    For U& = 1 To LOF(N1%) / 32
        Get #N1%, U&, DEU1$
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
            CUIT$ = TRIM$(Mid$(DEU2$, 75, 15))
            If VERICUIT% <> (-1) Then
              If CUIT$ <> "" Then
                If VALICUIT%(CUIT$) <> 1 Then
                  Screen.MousePointer = 1
                  TECLI$ = "Cliente": If ARDEU% = 2 Then TECLI$ = "Proveedor"
                  If COMALTER%(TECLI$ + ": " + TRIM$(CLI$) + "\C.U.I.T.: " + CUIT$ + "\  \Número de C.U.I.T. no Válido.\\Continuar Revisión\Cancelar") = 2 Then
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
        Put #N1%, U&, DEU1$
        Put #N2%, U&, DEU2$
        URE& = U&
    Next U&
    For U& = URE& + 1 To LOF(N1%) / 32
        DEU1$ = String$(32, 0)
        DEU2$ = String$(128, 0)
        Put #N1%, U&, DEU1$
        Put #N2%, U&, DEU2$
    Next U&
    '
    Unlock #N2%
    Unlock #N1%
    Close #N2%: N2% = 0
    Close #N1%: N1% = 0
    DEU1$ = REGLEIDO$(ASELE$, 1)
    Call GRAREG(ASELE$, DEU1$, 1)
    Call CIERRARCH(ASELE$)
    Screen.MousePointer = 1
    Exit Sub
    '
100 Resume 101
101 On Error GoTo 0
    Close #N1%: N1% = 0
    Call COMUNI("Imposible Ordenar - Archivo en Uso.")
    Exit Sub
    '
200 Resume 201
201 On Error GoTo 0
    Close #N1%: N1% = 0
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
     For I& = 1 To ULTREG&("PROVINB")
       X$ = REGLEIDO$("PROVINB", I&)
       If ECOARCH$("CUECON", Mid$(X$, 27, 2)) = "" Then
          Call REPLA(X$, MKI$(0), 27, 2)
       End If
       If TRIM$(Left$(X$, 1)) <> "" Then
          If Len(COPROVI$) < 64 Then
            COPROVI$ = COPROVI$ + Left$(X$, 1)
            CUEPROVI$(Len(COPROVI$)) = Mid$(ECOARCH$("CUECON", Mid$(X$, 27, 2)), 31, 12)
          End If
       End If
     Next I&
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
     For I& = 1 To ULTREG&("PROVINB")
       X$ = REGLEIDO$("PROVINB", I&)
       If ECOARCH$("CUECON", Mid$(X$, 27, 2)) = "" Then
          Call REPLA(X$, MKI$(0), 27, 2)
       End If
       If TRIM$(Left$(X$, 1)) <> "" Then
          If Len(COPROVI$) < 64 Then
            COPROVI$ = COPROVI$ + Left$(X$, 1)
            CUEPROVII%(Len(COPROVI$)) = CVI(Mid$(X$, 27, 2))
          End If
       End If
     Next I&
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
        TOTALITEM# = (XVALO(!CANPED) - XVALO(!CantEnt)) * XVALO(!PreUnid)
        VALOPEDPEN1# = VALOPEDPEN1# + TOTALITEM#
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
       CUITPROV$ = CUITCLI$((-1) * Abs(NPROVE%))
       XX$ = REPADIBRU$(CUITPROV$, 2)
       If TRIM$(XX$) <> "" Then
          RET$ = Mid$(XX$, 48, 4)
          RETENC# = XVALO(RET$)
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
       cuitclien$ = CUITCLI$(NCLIE)
       XX$ = REPADIBRU$(cuitclien$, 1)
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
      
'   PRONUCLI& = XVALO(Control$("", "PRONUCLI"))
'   '
'2  NROPRO& = PRONUCLI&: If NROPRO& < 0 Then NROPRO& = 0
'3  NROPRO& = NROPRO& + 1
'   If NROPRO& > 999999 Then
'      If PRONUCLI& > 0 Then
'           PRONUCLI& = 0: GoTo 2  ' VUELVE A EMPEZAR
'         Else
'           Call MENSERR(24, "No hay Numeros de Cliente Libres\para Alta Automática. Intente Alta Manual.")
'           Call FINAL("")
'      End If
'   End If
   If CantRegistros("DEUDOR12", "Nume_Cli = " & NROPRO&) > 0 Then
      GoTo 3
   End If
   '
   NC& = XVALO(VALOBADA("DEUDOR12", "Nume_Cli", "Nume_Cli=" & NROPRO&))
   If NC& > 0 Then GoTo 3
   NCLIE = NROPRO&
   '
   ' PRESENTA VENTANA DE ALTA CON DATOS BASICOS SEGUN PROSPECTO
   '
   'VERIFICO SI TIENEN GUARDADO POSICION DE IVA
   CONDI$ = " CODI_CLI = '" & CODICL$ & "'"
   'POXIVA% = XVALO(VALOBADA("BSEGCOM", "PIVA_CLI", CONDI$))
   CUITX$ = VALOBADA("BSEGCOM", "CUIT_CLI", CONDI$, "NOMESS")
   RASO$ = VALOBADA("BSEGCOM", "RASO_CLI", CONDI$, "NOMESS")
   If POXIVA% < 1 Then POXIVA% = 1
   VDEF$ = String$(128, 0)
   Call REPLA(VDEF$, CODICL$, 1, 12)
   Call REPLA(VDEF$, MKS$(NCLIE), 13, 4)
   Call REPLA(VDEF$, RASO$, 17, 62)
   Call REPLA(VDEF$, Chr$(1), 79, 1)
   Call REPLA(VDEF$, Space$(15), 80, 15)
   '
10 Screen.MousePointer = 1
   OBX$ = OBJPLA$("ALTACLIENACCN", VDEF$)
   If OBX$ = "" Then GoTo 99
   VDEF$ = OBX$
   
'   VDEF$ = String$(128, 0)
'   Call REPLA(VDEF$, CODICL$, 1, 12)
'   Call REPLA(VDEF$, MKI$(NCLIE%), 13, 2)
'   RASO$ = SAFETEXT$(VALOBADA("BSEGCOM", "Raso_Cli", "Codi_Cli='" & CODICL$ & "'"))
'   Call REPLA(VDEF$, RASO$, 15, 62)
'   Call REPLA(VDEF$, Chr$(POXIVA%), 77, 1)
'   Call REPLA(VDEF$, CUITX$, 78, 15)
'   '
'10 Screen.MousePointer = 1
'   OBX$ = OBJPLA$("ALTACLIENACC", VDEF$)
'   If OBX$ = "" Then GoTo 99
'   '
   ' VALIDA DATOS BASICOS
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
   '
   PIVCLI% = Asc(Mid$(OBX$, 79, 1))
   If PIVCLI% < 0 Or PIVCLI% > 6 Then
      Call MENSERR(24, "Ingreso no Válido.\Posición IVA no Figura en Tablas.")
      GoTo 10
   End If
   '
   CUTT$ = TRIM$(Mid$(OBX$, 80, 15))
   If VALICUIT%(CUTT$) <> 1 Then
     If PIVCLI% <> 0 Then
       If PIVCLI% <> 5 Then
         Call MENSERR(24, "Ingreso no Válido.\Número de C.U.I.T. Incorrecto.")
         GoTo 10
       End If
     End If
   End If
   '
   ' VALIDA EL USO DEL CUIT EN OTRO CLIENTE
   '
   If CUTT$ <> "" Then
     NCLIX = XVALO(VALOBADA("DEUDOR12", "Nume_Cli", "Cuit_Cli = '" & CUTT$ & "'"))
     If NCLIX > 0 Then
        If COMALTER%("Existe Registrada una Cuenta con el Mismo C.U.I.T.\  \'" + TRIM$(RASOCLI$(NCLIX)) + "' (" + TRIM$(Str$(NCLIX)) + ")\\Cancelar\Registrar Igual") <> 2 Then
           GoTo 10
        Else 'VALIDAR SI TIENE DERECHOS PARA DAR DE ALTA CLIENTE CON CUIT EXISTENTE
           If DERACCE%("ALCUITEX", "Alta de Cliente con C.U.I.T. Existente") < 2 Then
              GoTo 10
           End If
        End If
     End If
   End If
   '
   ' AGREGAR EL NUEVO CLIENTE A LA TABLA
   '
   Call BLOQARCH("DEUDOR2")
   Dim Rdeudor12 As ADODB.Recordset
   Set Rdeudor12 = New ADODB.Recordset
   Rdeudor12.CursorType = adOpenDynamic
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
      Call COMUNI("FLEXOFT ha Re-Numerado el Cliente por '" + TRIM$(FORMATNUM$(NCLIE, "I5")) + "'.\  \Verifique y Vuelva a Grabar.")
      Call LIBARCH("DEUDOR2")
      VDEF$ = OBX$: Call REPLA(VDEF$, MKI$(NCLIE), 13, 2)
      GoTo 10
   End If
   CONDI$ = "Codi_Cli='" & CODICL$ & "'"
   With Rdeudor12
      .AddNew   ' YA CONVERTIDO
      On Error Resume Next
      !CodiEmpr = CodiEmp%
      On Error GoTo 0
      !Codi_Cli = CODICL$
      !NUME_CLI = TRIM$(Str$(NCLIE))
        RASO$ = SAFETEXT$(VALOBADA("BSEGCOM", "Raso_Cli", CONDI$))
      !Raso_Cli = RASO$
      !Mail_Cli = SAFETEXT$(VALOBADA("BSEGCOM", "Mail_Cli", CONDI$))
      !Cuit_Cli = CUTT$
      !Domi_Cli = SAFETEXT$(VALOBADA("BSEGCOM", "Domi_Cli", CONDI$))
      !Cpos_Cli = SAFETEXT$(VALOBADA("BSEGCOM", "Cpos_Cli", CONDI$))
      !Loca_Cli = SAFETEXT$(VALOBADA("BSEGCOM", "Loca_Cli", CONDI$))
      !Prov_Cli = SAFETEXT$(VALOBADA("BSEGCOM", "Prov_Cli", CONDI$))
      !Pais_Cli = SAFETEXT$(VALOBADA("BSEGCOM", "Pais_Cli", CONDI$))
      !Loca_Cli = SAFETEXT$(VALOBADA("BSEGCOM", "Loca_Cli", CONDI$))
      !Vend_Cli = XVALO(VALOBADA("BSEGCOM", "Vend_Cli", CONDI$))
      !Loca_Cli = SAFETEXT$(VALOBADA("BSEGCOM", "Loca_Cli", CONDI$))
      !TELE_Cli = SAFETEXT$(VALOBADA("BSEGCOM", "Tele_Cli", CONDI$))
      !FAXI_Cli = SAFETEXT$(VALOBADA("BSEGCOM", "Fax_Cli", CONDI$))
      !Ctac_Cli = SAFETEXT$(VALOBADA("BSEGCOM", "Ctac_Cli", CONDI$))
      !Nota_Cli = SAFETEXT$(VALOBADA("BSEGCOM", "Nota_Cli", CONDI$))
      '!Nibr_Cli = SAFETEXT$(VALOBADA("BSEGCOM", "NIBR_CLI", CONDI$)) ' INGRESO BRUTOS
      
      '!Piva_Cli = 1
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
'   UPDATCLI% = 1
   '
   Call CARGALISEL("DEUDOR1", "DEUDOR12", "Nume_Cli/F6.0;Raso_Cli/A30", "NUME_CLI > 0 ORDER BY RASO_CLI ASC")
   '
'   TTXX$ = MKI$(NCLIE%) + RASO$
'   Call FILTERPUTRECORD("DEUDOR1", 1, MKI$(NCLIE%), TTXX$)
'   'Call FILESORT("DEUDOR1", "", 1, 1, "ASC") se comento esta linea para mejorar la velocidad y evitar que se pisen al dar de alta clientes desde el prospect en Meditea
'   Call FRESHECHO("DEUDOR1")
   '
99 Screen.MousePointer = 1
End Sub


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
             SALACR# = XVALO(!SALDAcre)
             DEUDA$ = FORMATNUM$(SALDEB# - SALACR#, "F12.2")
             IMPORIG# = XVALO(!IBruComp)
             FEVE = CVINT(!FEVENCIM)
             NUMFACT$ = SAFETEXT(!CODICOM_LAR)
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

Sub GUARDAVALCOT(NUMECOTI$, CLAVEVAL$, VALOR$, Optional COLUVALOR)
    'ESTA RUTINA ACTUALIZA O CARGA UN NUEVO REGISTRO EN LA TABLA DETCOTIZA (DETALLE DE COTIZACIONES)
    'SEGUN EL NUMERO DE COTIZACION Y LA CLAVE
    '
    CONDI$ = "NroCoti = '" & NUMECOTI$ & "' AND ClaValor = '" & CLAVEVAL$ & "'"
    '
    RR& = CantRegistros&("DETCOTIZA", CONDI$)
    If RR& = 0 Then
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
            !CodiEmpr = CodiEmp%
            On Error GoTo 0
                !NroCoti = NUMECOTI$
                !ClaValor = CLAVEVAL$
                If XVALO(COLUVALOR) < 2 Then
                    !VALOR1 = TRIM$(VALOR$)
                ElseIf XVALO(COLUVALOR) = 2 Then
                    !VALOR2 = TRIM$(VALOR$)
                ElseIf XVALO(COLUVALOR) = 3 Then
                    !Valor3 = TRIM$(VALOR$)
                End If
            .Update
         End With
    Else
        If XVALO(COLUVALOR) < 2 Then
            Call ACTUREGISTRO("DETCOTIZA", "Valor1", CONDI$, TRIM$(VALOR$), RAF&, "NOMESS")
        ElseIf XVALO(COLUVALOR) = 2 Then
            Call ACTUREGISTRO("DETCOTIZA", "Valor2", CONDI$, TRIM$(VALOR$), RAF&, "NOMESS")
        ElseIf XVALO(COLUVALOR) = 3 Then
            Call ACTUREGISTRO("DETCOTIZA", "Valor3", CONDI$, TRIM$(VALOR$), RAF&, "NOMESS")
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

Function CLASE$(INDICE#)
    ' devuelve la clase segun el indice
    ' si el indice es 0 o no existe el ejecutable devuelve vacio
    CLASE$ = ""
    
    If EXISTE%("CALPRO09.EXE") > 0 And INDICE# > 0 Then
        If INDICE# >= XVALO(Control("CALIPROV", "LIMCAT-A ")) Then CLASE$ = "A"
        If INDICE# >= XVALO(Control("CALIPROV", "LIMCAT-B ")) And INDICE# < XVALO(Control("CALIPROV", "LIMCAT-A ")) Then CLASE$ = "B"
        If INDICE# < XVALO(Control("CALIPROV", "LIMCAT-B ")) Then CLASE$ = "C"
    End If
    
End Function

Function ENVMAIDEUDA%(NC)
   ' Devuelve 0 = "No envia", y 1 = "Envia"
   ENVMAIDEUDA% = 0
   On Error Resume Next
   If STATCLI%(NC) <> 0 Then ENVMAIDEUDA% = XVALO(ClienTemp!ENVMAIDEUDA)
   '
End Function

Sub MUESTRADOC(DOCUNU&, Optional MODOX$, Optional Tipo$, Optional NUBUSQUEDA$, Optional NUCOMPLE$, Optional VARCONTROL$)
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
                  Mid$(TPSP$, KKL&, 1) = Chr$(255)
               End If
            Next KKL&
            VUELTAREP% = 1
            SPOOLSTRING$ = TPSP$
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
    If ALTUPAGINA > EDITFORM.Frame1.Height Then
       EDITFORM.VScroll1.Max = (ALTUPAGINA - EDITFORM.Frame1.Height) / 100
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


Sub SAVESPOOL(DESCRIDOC$, OKX%)
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
       !CodiEmpr = CodiEmp%
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
     IMPREDET$ = Printer.DeviceName
     Printer.TrackDefault = False
     For Each PRX1 In Printers
       If InStr(UCase$(PRX1.DeviceName), "PDFWRITER") > 0 Or InStr(UCase$(PRX1.DeviceName), "PDF4U") > 0 Or InStr(UCase$(PRX1.DeviceName), "BULLZIP") > 0 Then
         Printer.EndDoc
         On Error Resume Next
         Set Printer = PRX1
         If Err Then
           On Error GoTo 0
           GoTo 32
         End If
         On Error GoTo 0
         ARANEX$ = DEDOCOR$
         If ARANEX$ = "" Then ARANEX$ = "PDF"
         On Error Resume Next
         Kill LUCOM$ + "" + ARANEX$ + ".PDF"
         On Error GoTo 0
         
         EDITFORM.Label1 = ARANEX$
         EDITFORM.Timer1.Enabled = True
         GoTo 35
       End If
     Next
     '

32   Call MENSERR(24, "Imposible Direccionar a Impresora Virtual PDF")
   End If
   '
35 On Error Resume Next

'MsgBox "Comienza la Impresión"
   Printer.EndDoc   ' INICIALIZA LA IMPRESORA
   Printer.ScaleMode = 1
        'MsgBox "SCALEMODE: " & Printer.ScaleMode
        'MsgBox "SCALEWIDTH: " & Printer.ScaleWidth
        'MsgBox "SCALELEFT: " & Printer.ScaleLeft
        'MsgBox "SCALEHEIGHT: " & Printer.ScaleHeight
        'MsgBox "SCALETOP: " & Printer.ScaleTop
        'MsgBox "PAPERSIZE: " & Printer.PaperSize
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
   'For KKII% = 1 To FORMALTER.PSPOOL.ListCount
   For KKIIL& = 1 To CAREVEC&
      'TTTT$ = TRIM$(FORMALTER.PSPOOL.List(KKII% - 1))
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
               'MsgBox TTXX$
           GoTo 39
        End If
      Next KKJJ%
39 Next KKIIL&
   '
'MsgBox "Termina la Impresión"
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
'40 FORMALTER.PSPOOL.Clear
40 ReDim VECPRICO$(0)
   CAREVEC& = 0
   Call CLOSEFORMPRIN("")
   '
   If SENDBYMAIL$ <> "" Then
      'VALIDACION PARA EL CASO SI ES BULLZIP LA IMPRESORA PRESENTA MENSAJE
      For Each PRX10 In Printers
        If InStr(UCase$(PRX10.DeviceName), "BULLZIP") > 0 Then
           DoEvents
           TX1$ = "Generando Archivo PDF "
45         OPX% = COMALTER%(TX1$ & " \\Continuar\Cancelar")
           If OPX% = 2 Then Exit Sub  'SI CANCELA SALE
           If EXISTE%(LUCOM$ + "" + TRIM$(EDITFORM.Label1) + ".PDF") < 1 Then GoTo 45
         End If
      Next
      '
     TINI = Timer
50   If EXISTE%(LUCOM$ + "" + TRIM$(EDITFORM.Label1) + ".PDF") < 1 Then
       If Timer < TINI + 10 Then
         GoTo 50
       End If
       '
       Call MENSERR(24, "Imposible Enviar por Mail.\  \No se ha Generado Correctamente\el Archivo Adjunto en Formato PDF.")
       Exit Sub
     End If
     'ANTES DEL ENVIAMAIL LE ASIGNO LA RUTA POR  QUE QUEDO EN C:\FLEXOFT EN VEZ DE X:\GABAL\FLEXOFT\WINDOWS
     'X:\GABAL\FLEXOFT\WINDOWS (EJEMPLO DE  RUTA PARA VER COMO SE ASIGNA PARTICIONANDO EL TEXTO)
     ChDrive (Left$(RUXA$, 1)) 'ASIGNO LA UNIDAD DE RED PRIMER POSICION DE LA RUTA (X)
     ChDir Mid$(RUXA$, 4)      'LE ASIGNO LA RUTA DESPUES DE LA BARRA  (FLEXOFT\WINDOWS)
     'FIN DE ASIGNACION DE RUTA PERDIDA
     Call ENVIAPORMAIL
     
   End If
   OKX% = 1
   '
End Sub

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
         CONDI$ = CONDI$ + " AND REFERENCIA LIKE '%" & VARCONTROL1$ & "%'"
       End If
    End If
    '
    
    'AGREGADO PARA VISUALIZAR DUPLICADOS
    If CantRegistros&("PDOC_TBL", CONDI$, "NOMESS") > 1 And TIPO1$ = "FACTURA" And VARCONTROL1$ = "" Then
       Call CARGALISEL("LISFACTX", "PDOC_TBL", "REFERENCIA/A64", CONDI$, "NOMESS")
       Call SELECHO("LISFACTX/SELECCION DE FACTURA")
       REFE$ = TRIM$(VALACT1$("LISFACTX"))
       If Resp$ <> "" Then
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
    Call ABRECLIEN
    Call CREACAMPO("", "CONTACSCOM", "CODIEMPR", 3, 0)
    Call CREACAMPO("", "CONTACSCOM", "CODICLI", 10, 12)
    Call CREACAMPO("", "CONTACSCOM", "ORDI_CTAC", 3, 0)
    Call CREACAMPO("", "CONTACSCOM", "NOME_CTAC", 10, 64)
    Call CREACAMPO("", "CONTACSCOM", "CARG_CTAC", 10, 64)
    Call CREACAMPO("", "CONTACSCOM", "MAIL_CTAC", 10, 128)
    Call CREACAMPO("", "CONTACSCOM", "TELE_CTAC", 10, 64)
    Call CREACAMPO("", "CONTACSCOM", "CELU_CTAC", 10, 64)
    Call CREACAMPO("", "CONTACSCOM", "MARBORRA", 3, 0)
    '
    If CantRegistros&("CONTACSCOM", "CODICLI = '" & COCLI$ & "'", "NOMESS") > 0 Then
       'CARGO PRIMERO EL DEL MAESTRO
       MAILX$ = EMAILCLI$(XVALO(NCL))
       If InStr(MAILX$, "@") > 0 Then
           Call REPLA(TX$, MAILX$, 44, 32)
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
              NOMB$ = SAFETEXT$(!NOME_CTAC)
              CARG$ = SAFETEXT$(!CARG_CTAC)
              Call REPLA(TX$, NOMB$, 1, 20)
              Call REPLA(TX$, CARG$, 22, 20)
              Call REPLA(TX$, MAILX$, 44, 32)
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
        CONDI1$ = "DOSECLAR='" & NUFACTU$ & "' AND NUME_CLIE= " & NCLI2 & ""
        CANTREGREPA& = XVALO(CantRegistros("MOVIREPA", CONDI1$, "NOMESS"))
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
    ARCSQR$ = FORIPRE$
    If EXISTE%(LUDAT$ + ARCSQR$ + ".SQR") = 0 Then
       Call COMUNI("Falta Archivo " & ARCSQR$ + ".SQR")
    End If
    '
    'SE LE AGREGA POR DEFECTO LA LEYENDA ORIGINAL.
    FILTX$ = TRIM$(NUFACTU$) & ";" & TRIM$(Str$(NCLI2) & ";" & "ORIGINAL")
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
          I& = 0
          Do While Not .EOF
             COD$ = SAFETEXT$(!CODIGO)
             MON% = XVALO(!Monecos)
             PODES1 = ROUND#(XVALO(!PorDescCos), 1): If PODES1 > 99 Then PODES1 = 0
             COS1# = XVALO(!COSUNI / (1 - PODES1 / 100))
             COS2# = XVALO(!COSUNI)
             FECH = SAFETEXT$(!FECOSTO)
             '
             Y$ = REGBLAN$("ACTUCOS")
             CODI% = CODINT%(COD$)
             Call REPLA(Y$, MKI$(CODI%), 1, 2)
             Call REPLA(Y$, MKI$(MON%), 3, 2)
             Call REPLA(Y$, MKD$(COS1#), 5, 8)
             Call REPLA(Y$, MKS$(PODES1), 13, 4)
             Call REPLA(Y$, MKD$(COS2#), 17, 8)
             FCOS% = CVINT%(FECH)
             Call REPLA(Y$, MKI$(FCOS%), 25, 2)
             I& = I& + 1
             Call GRAREG("!ACTUCOS", Y$, I&)
            .MoveNext
            '
          Loop
          Call SETULTREG("!ACTUCOS", I&)
       End With
       '
       
40     JUN% = VENTABU%("ACTUCOS/TITUPAN=Costo de Reposición - " + RASOCLI$((-1) * NCX1%))
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
        Observacion$ = OBSERVAC.Text28
       End If
       If TRIM$(Observacion$) = "" Then Observacion$ = "Cambio de Precio Opción: x Proveedor"
       
       HOII% = HOY(HOS$)
       '
       Dim COSTOANTERIOR#(32767)
       For I& = 1 To 32767
         COSTOANTERIOR#(I&) = 0
       Next I&
       
        Call ABREMASTER
        For j& = 1 To ULTREG&("!ACTUCOS")
            Z$ = REGLEIDO$("!ACTUCOS", j&)
            COD3% = CVI(Mid$(Z$, 1, 2))
            PODES1 = CVS(Mid$(Z$, 13, 4)): If PODES1 < 0 Or PODES1 > 99 Then PODES1 = 0
            COU = (CVD(Mid$(Z$, 5, 8))) * (1 - PODES1 / 100)
            MONEI% = CVI(Mid$(Z$, 3, 2))
            FCOS% = CVI(Mid$(Z$, 25, 2))
            PHAB% = PROVHABI%(COD3%)
            '
            If PHAB% > 0 And NCX1% <> PHAB% Then
               TTXX$ = "Para el Artículo " + TRIM$(CODEXT$(COD3%)) + "\el Proveedor Habitual Registrado es\"
               TTXX$ = TTXX$ + TRIM$(RASOCLI$((-1) * PHAB%)) + "\  \"
               TTXX$ = TTXX$ + "Desea Cambiar el nuevo Proveedor Habitual a\"
               TTXX$ = TTXX$ + RASOCLI$((-1) * NCX1%)
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
                
        Next j&
        '
        'RECORRO NUEVAMENTE PARA ACTUALIZAR LAS NOVEDADES
        For j& = 1 To ULTREG&("!ACTUCOS")
            Z$ = REGLEIDO$("!ACTUCOS", j&)
            COD3% = CVI(Mid$(Z$, 1, 2))
            PODES1 = CVS(Mid$(Z$, 13, 4)): If PODES1 < 0 Or PODES1 > 99 Then PODES1 = 0
            COU = (CVD(Mid$(Z$, 5, 8))) * (1 - PODES1 / 100)
            MONEI% = CVI(Mid$(Z$, 3, 2))
            PHAB% = NCX1%
            
            If ROUND(COSTOANTERIOR#(COD3%), 4) <> ROUND(COU, 4) Then
               PREC_NVO# = XVALO(COU)
               Call ADD_NOVEDAD(COD3%, NCX1%, COSTOANTERIOR#(COD3%), PREC_NVO#, Observacion$, MONEI%, "COSTO")
            End If
        Next j&
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
    VALID% = 1
    If NLISTA% < 0 Then NLISTA% = 0
    '
    'SI ES EL ADMINISTRADOR APRUEBA DIRECTAMENTE.
    USER1% = USNBR% Mod 100
    If USER1% = 1 Then GoTo 99
    '
    'VALIDACION SI EL CODIGO NO PERTENECE A LA LISTA DEVUELVE -3 Y SE VA
    If CantRegistros("PREVENTA", "Cod_Pieza = " & CI1%) < 1 Then
       VALID% = -3
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
          VALID% = -2
          GoTo 99
       End If
    End If
    '-------------------------------------------------------------------------------------
    'SI ES LA LISTA DEL CLIENTE APRUEBA LISTA O TIENE DERECHO > 0 MINIMAMENTE PUEDE CAMBIAR LA LISTA.
    'EN ESTE CASO EL DERECHO SOLO PODRIA SER 0 O 1
    If NUMLISTA_CLIENTE% <> NLISTA% And DEREX% <= 0 Then
      VALID% = VALID% - 1
    End If  '
    '--------------------------------------------------------------------------------------
    'SI CAMBIO EL PRECIO DEVUELVE -1 , o -2 SI TAMBIEN SE CAMBIO LA LISTA DE PRECIOS SIN DERECHOS
    If Abs(ROUND(PRECIO_VENTA#, 2) - ROUND(PRELISTA#(TRIM$(Str$(NLISTA%)), CI1%), 2)) > 0.05 Then
      VALID% = VALID% - 2
    End If
    '
    '---------------------------------------------------------------------------------------
99   Vali_Derechos_Lpre = VALID%
End Function
Function PROVELOTE%(LOTE$)
    'DEVUELVE EL PROVEEDOR QUE INGRESO EL LOTE PASADO COMO PARAMETRO
    NROLOTE& = XVALO(Mid$(LOTE$, 4, 6))
    CONDI$ = "NumeBoRe = " & NROLOTE&
    PROVELOTE% = XVALO(VALOBADA("BOLRECEP", "NumProv", CONDI$))
End Function
Function VIALOTE%(LOTE$)
    'DEVUELVE LA VIA DE ENTREGA QUE INGRESO EL LOTE PASADO COMO PARAMETRO
    NROLOTE& = XVALO(Mid$(LOTE$, 4, 6))
    CONDI$ = "NumeBoRe = " & NROLOTE&
    VIALOTE% = XVALO(VALOBADA("BOLRECEP", "Via_Entrega", CONDI$, "NOMESS"))
    If VIALOTE% < 1 Then
       NOC& = XVALO(VALOBADA("BOLRECEP", "NumOC", CONDI$, "NOMESS"))
       VIALOTE% = VIAENTRE_COMPRA(NOC&)
    End If
End Function


Function PORCENT_COSTO_NACIONALIZACION(NUMEBORE&, Optional MONEDX%)
    PORCENT_COSTO_NACIONALIZACION = 0
    If NUMEBORE& < 1 Then Exit Function
    '
    If MONEDX% < 1 Then MONEDX% = 2 ' POR DEFECTO DOLAR
    'OBTENGO EL VALOR DE COMPRA Y EL NUMERO DE FACTURA DEL LOTE
    CONDI$ = " NUMEBORE = " & NUMEBORE&
    NUFACTU$ = TRIM$(VALOBADA("FACTURAYGASTOS", "FACTURAPROV", CONDI$, "NOMESS"))
    MONEFACTU% = XVALO(VALOBADA("FACTURAYGASTOS", "MoneVal", CONDI$, "NOMESS"))
    VALCOMPRA# = XVALO(VALOBADA("FACTURAYGASTOS", "VAL_COMPRA", CONDI$, "NOMESS"))
    '
    If MONECOMP% <> MONEDX% Then
       COEFCON = TICAMONE(MONEFACTU%) / TICAMONE(MONEDX%)
       VALCOMPRA# = COEFCON * VALCOMPRA#
    End If

    
    'OBTENGO EL GASTO RELACIONADO CON ESTA FACTURA ACUMULADO FILTRANDO POR MONEDA, POR SI SE CARGARON GASTOS CON DIFERENTES MONEDAS.
    If NUFACTU$ <> "" Then
       '
       CONDI$ = "FACTURAPROV_ASOC = '" & NUFACTU$ & "' AND MoneVal = " & 1
       GASTOS_APLICADOS_PESOS# = XVALO(VALOBADA("FACTURAYGASTOS", "Sum(IMPO_GASTO)", CONDI$, "NOMESS"))
       If 1 <> MONEDX% Then
          COEFCON = TICAMONE(1) / TICAMONE(MONEDX%)
          GASTOS_APLICADOS_PESOS# = COEFCON * GASTOS_APLICADOS_PESOS# ' AQUI LOS CONVIERTO
       End If
       '
       CONDI$ = "FACTURAPROV_ASOC = '" & NUFACTU$ & "' AND MoneVal = " & 2
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
    
End Function

Function MONELOTE%(LOTE$)
    MONELOTE% = 1
    CONDI$ = "IDENLOTE = '" & LOTE$ & "'"
    NUOC& = XVALO(VALOBADA("BOLRECEP", "NumOC", CONDI$, "NOMESS"))
    If NUOC& > 0 Then
       CONDI$ = "NUME_OCOM = " & NUOC&
       MONELOTE% = XVALO(VALOBADA("OCOMDETA", "Mone_Emis", CONDI$, "NOMESS"))
    End If

End Function
Function NumeroSolicitud&(NPROV, NPRES&)
   'DEVUELVE EL NUMERO DE SOLICITUD SEGUN EN NRO DE PRESUPUESTO Y EL DEL PROVEEDOR.
   NumeroSolicitud& = XVALO(VALOBADA("PRESUCOMPRA", "NumSolCompra", "NroPres = " & NPRES& & " and  Num_Prov = " & NPROV, "NOMESS"))
End Function

Function ProxNumSolCompra&()
   ProxNumSolCompra& = 1 + XVALO(VALOBADA("PRESUCOMPRA", "MAX(NumSolCompra)", "NumSolCompra > 0", "NOMESS"))
End Function
Function OrdenCompraSegunSolicitud&(NSOLIC&)
    CONDI$ = "Medi_Ocom = " & NSOLIC&
    OrdenCompraSegunSolicitud& = XVALO(VALOBADA("OCOMDETA", "nume_ocom", CONDI$, "NOMESS"))
End Function

