Attribute VB_Name = "MOVIST07_SQL"
Public CAITPE%, CIPE%(2048), CADESPA(2048), REPE&(2048)
Public NOARRAS%
Public CIMOVI%(32767)
'
Public CAMOVILO%
Public CIMOVIM%(32767), CAMOVIM(32767), CAITMO%
Public TIMOVILO%(512), CANINGLO#(512), CANSALLO#(512), CANRESLO#(512), FEMOVILO%(512), DEMOVILO$(512)
'
Public APROBOLE%      ' BANDERA PARA INDICAR QUE ES PROCESO DE APROBACION DE BR
'

Sub RECUSALSTOK()
   'se dejo solo para que compile
End Sub
Sub ACSALOC(ID_OCOM$)
   '
   ' ACTUALIZA MOVISTO CON EL VALOR DE DocuOrig TOMADO DE LOTINGRE
   Screen.MousePointer = 11
   CONDI$ = "CodiMovi = 'BR' "
   CONDI$ = CONDI$ + " AND DocuOrig  IS NULL "
   QTR& = CantRegistros("MOVISTO", CONDI$)
   If QTR& > 0 Then
      JJ& = 0
      Dim MOVISEL As ADODB.Recordset
      Set MOVISEL = New ADODB.Recordset
      Set MOVISEL = FLEXCONN.Execute(FILTSQL$("SELECT * FROM MOVISTO WHERE " + CONDI$))
      With MOVISEL
         Do While Not .EOF
            JJ& = JJ& + 1
            Call TRACE(20, JJ&, QTR&)
            CIXI% = XVALO(!Cod_Inte)
            IDLO$ = SAFETEXT$(!IdenLote)
            DOCUCO$ = ""
            If CIXI% > 0 Then
              If IDLO$ <> "" Then
                DOCUCO$ = VALOBADA("LOTINGRE", "DOCUORIG", "COD_INTE = " & CIXI% & " AND IdenLote = '" & IDLO$ & "'")
              End If
            End If
            CONDI1$ = CONDI$ + " AND COD_INTE = " & CIXI% & " AND IdenLote = '" & IDLO$ & "'"
            Call ACTUREGISTRO("MOVISTO", "DocuOrig", CONDI1$, DOCUCO$, RAFE&, "NOMESS")
         .MoveNext
         Loop
      End With
      MOVISEL.Close
      Set MOVISEL = Nothing
   End If
   '
   ' ACTUALIZA SALDOS DE OCOMDETA
   Screen.MousePointer = 11
   NUMPED% = XVALO(VALOBADA("OCOMENCA", "Nume_Pedi", "ID_COMPROB='" & ID_OCOM$ & "'"))
   NUMOCO% = XVALO(VALOBADA("OCOMENCA", "Nume_Ocom", "ID_COMPROB='" & ID_OCOM$ & "'"))
   sqlx$ = ""
   If NUMPED% > 0 Then
        sqlx$ = "SELECT * FROM OCOMDETA WHERE Nume_Pedi = " & NUMPED%
     ElseIf NUMOCO% > 0 Then
        sqlx$ = "SELECT *FROM OCOMDETA WHERE Nume_Ocom = " & NUMOCO%
   End If
   If sqlx$ = "" Then Exit Sub
   '
   Dim RSOC As ADODB.Recordset
   Set RSOC = New ADODB.Recordset
25 On Error Resume Next
   RSOC.Open FILTSQL$(sqlx$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
   '
   If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
      On Error GoTo 0
      Call CapturaErrorOpen
      GoTo 25
   End If
   On Error GoTo 0
   '
   With RSOC
      Do While Not .EOF
        CIXI% = XVALO(!Cod_Inte)
        CONDI$ = "CodiMovi = 'BR' "
        CONDI$ = CONDI$ + " AND DocuOrig = '" & ID_OCOM & "' "
        CONDI$ = CONDI$ + " AND COD_INTE = " & CIXI%
        CANRE = XVALO(VALOBADA("MOVISTO", "SUM(CantIngre)", CONDI$))
        CANRECI# = XVALO(VALOBADA("LOTINGRE", "SUM(CANALTA)", "DOCUORIG = '" & ID_OCOM$ & "' AND COD_INTE = " & CIXI%))
        If CANRECI# > CANRE Then
          CANRE = CANRECI#
        End If
        '
        CANOC = XVALO(!cant_ocom)
        STAOC = XVALO(!Stat_Ocom)
        !Cant_Rec = CANRE
        If STAOC < 3 Then
          If CANRE >= CANOC Then
             !Stat_Ocom = 3
          End If
        End If
        .Update
      .MoveNext
      Loop
   End With
   RSOC.Close
   Set RSOC = Nothing
   Screen.MousePointer = 1
   '
End Sub
Sub RECULOTE(CIXI%)
    '
    Exit Sub
    
    'ARMO EL LOTE CON FORMATO
    '************************
10  ABUS1$ = RANDSTRING$(-6)
    LOTE$ = "LA-" + ABUS1$ + "-1"
    '
    'VERIFICO QUE NO SEA EXISTENTE
    '*****************************
    CONDI$ = "IDENLOTE = '" & LOTE$ & "'"
    If CantRegistros("MOVISTO", CONDI$) > 0 Then GoTo 10
    If CantRegistros("LOTINGRE", CONDI$) > 0 Then GoTo 10
    '
    'GENERO EL LOTE
    '**************
    HOII% = HOY(HOS$)
    Call CREALOTIN(HOII%, 0, 0, CIXI%, LOTE$, "", "", 1)
    '
    'ACTUALIZO EN LA TABLA MOVISTO
    '*****************************
    CONDI$ = "(IDENLOTE = '' OR IDENLOTE IS NULL) AND Cod_Inte = " & CIXI%
    Call ACTUREGISTRO("MOVISTO", "IDENLOTE", CONDI$, LOTE$, REGAF&, "NOMESS")
    '
    Call ACSALOTE(CIXI%, LOTE$, "NOMESS")
    '
End Sub

'
Function SALDOCOM(NUOC$, CIXI%) 'OBTIENE EL SALDO PEDNIENTE POR LA ORDEN DE COMPRA Y EL CODIGO DEL ITEM
    '
    Static NUOCOA$
    SALDOCOM = 0
    '
    If NUOC$ <> NUOCOA$ Then
       Call ACSALOC(NUOC$)
       NUOCOA$ = NUOC$
    End If
    '
    If Mid$(NUOC$, 1, 2) = "OC" Then
        NUMOCPR% = XVALO(Mid$(NUOC$, 4))
        CONDI$ = "NUME_OCOM = " & NUMOCPR% & " And COD_INTE = " & CIXI% & " AND STAT_OCOM < 8 "
      ElseIf Mid$(NUOC$, 1, 2) = "NP" Then
        NUMOCPR% = XVALO(Mid$(NUOC$, 4))
        CONDI$ = "NUME_PEDI = " & NUMOCPR% & " And COD_INTE = " & CIXI% & " AND STAT_OCOM < 8 "
    End If
    CANTOCPR = XVALO(VALOBADA("OCOMDETA", "SUM(CANT_OCOM)", CONDI$, "NOMESS"))
    CANTRECI = XVALO(VALOBADA("OCOMDETA", "SUM(CANT_REC)", CONDI$, "NOMESS"))
    SALD = (CANTOCPR - CANTRECI)
    If SALD < 0 Then SALD = 0
    '
    SALDOCOM = SALD
    '
End Function
'
Function PUEDEMOVI%(DEPOMO%, Optional MODO$)
   ' La funcion revisa si en el deposito elegido es posible el movimiento de salida
   ' El movimiento es lícito si
   '   - no hay control de negativos
   '   - es un movimiento de entrada (CAMOVIM() > 0)
   '   - en todos los casos ABS(CAMOVIM) <= SALDEPOS(CI%,DEPOMO%)
   ' En los casos anteriores sobre PUEDEMOVI% devuelve '1'
   ' Caso contrario presenta lista de los codigos con stock insuficiente
   ' y devuelve '-1' sobre PUEDEMOVI%
   '
   ' Si en MODO$ le pasamos 'NODESC' obvia la descomposicion
   ' en caso de items de descomposicion automatica
   Static CONEGATI%
   If CONEGATI% = 0 Then
     CONEGATI% = (-1)
     If TRIM$(Control$("STOCKS", "CONEGATI")) = "SI" Then
       CONEGATI% = 1
     End If
   End If
   '
   PUEDEMOVI% = 1
   If CONEGATI% > 0 Then
     ' BLANQUEAR LISTA DE CODIGOS CON STOCK INSUFICIENTE
     '\\\OT-05-0730-WAR-09/11/05///
     GoSub COPYCONTROLA
     Call BLANARCH("!COSTNEG")
     '\\\OT-05-0730-WAR-FIN///
     For KKI% = 1 To CAITMO%
       For KKJ% = KKI% + 1 To CAITMO%
         If CIMOVIM%(KKI%) = CIMOVIM%(KKJ%) Then
           CAMOVIM(KKI%) = CAMOVIM(KKI%) + CAMOVIM(KKJ%)
           CAMOVIM(KKJ%) = 0
         End If
       Next KKJ%
     Next KKI%
     '
     CAITMO1% = CAITMO%
     For KKI% = 1 To CAITMO1%
       If CAMOVIM(KKI%) < (-0.005) Then
         If Left$(UCase$(MODO$), 6) <> "NODESC" Then
           If DESCOMPO%(CIMOVIM%(KKI%)) > 0 Then
              '
              If STOCKACT(CIMOVIM%(KKI%)) < 0 Then
                 ' PUEDEMOVI% = (-1)    . se suprime porque al ser descomp. autom puede moverlo igual
                 Call COMUNI("Stock Negativo (Independiente de esta Transaccion)\ para Codigo " & CODEXT(CIMOVIM%(KKI%)) & " con Descomposicion Automatica")
              End If
              '
              Call AREXPLO(CIMOVIM%(KKI%), 1, 1)
              For KKJ% = 1 To CAIT%
                CAITMO% = CAITMO% + 1
                CIMOVIM%(CAITMO%) = CIJ%(KKJ%)
                CAMOVIM(CAITMO%) = USOI(KKJ%) * CAMOVIM(KKI%)
              Next KKJ%
              CAMOVIM(KKI%) = 0
           End If
         End If
       End If
     Next KKI%
     '
     For KKI% = 1 To CAITMO%
       For KKJ% = KKI% + 1 To CAITMO%
         If CIMOVIM%(KKI%) = CIMOVIM%(KKJ%) Then
           CAMOVIM(KKI%) = CAMOVIM(KKI%) + CAMOVIM(KKJ%)
           CAMOVIM(KKJ%) = 0
         End If
       Next KKJ%
     Next KKI%
     '
     j& = 0: LICONEG$ = ""
     Y$ = REGBLAN("COSTNEG")
     For KKI% = 1 To CAITMO%
       If CAMOVIM(KKI%) < (-0.005) Then
         DEPMOI% = DEPOMO%: If DEPOMO% = 99 Then DEPMOI% = DEPOPRE%(CIMOVIM%(KKI%))
         'SI SACO LA MISMA CANTIDAD QUE TENGO NO ME DEJA SACARLO
         VUELTA% = 1
         '
25       If UCase$(MODO$) = "CONSIDERA_RESERVAS" Then
            STDP = STODEPOS_NUEVO(CIMOVIM%(KKI%), DEPMOI%) - STOCK_RESERV_SIN_ARMADOR(CIMOVIM%(KKI%))
         Else
            STDP = (STODEPOS_NUEVO(CIMOVIM%(KKI%), DEPMOI%))
         End If
         '
         DIF1 = (STDP + CAMOVIM(KKI%))
         If DIF1 < (-0.005) Then
           If VUELTA% = 1 Then
              VUELTA% = 2
              Call VESARRAS(CIMOVIM%(KKI%))
              If TIBASE$ = "MDB" Then GoTo 25
           End If
           ' AGREGAR A LISTA DE CODIGOS NEGATIVOS
           '\\\OT-05-0730-WAR-09/11/05///
           CANTI2 = Abs(CAMOVIM(KKI%))
           DIF2 = Abs(STDP - CANTI2)
           Call REPLA(Y$, MKI$(CIMOVIM%(KKI%)), 1, 2)
              If LICONEG$ <> "" Then LICONEG$ = LICONEG$ + "; "
              LICONEG$ = LICONEG$ + CODEXT$(CIMOVIM%(KKI%))
           Call REPLA(Y$, MKS$(CANTI2), 5, 4)
           Call REPLA(Y$, MKS$(STDP), 21, 4)
           Call REPLA(Y$, MKS$(DIF2), 25, 4)
           j& = j& + 1
           Call GRAREG("!COSTNEG", Y$, j&)
           '\\\OT-05-0730-WAR-FIN///
           PUEDEMOVI% = (-1)
         End If
       End If
     Next KKI%
     '
     '\\\OT-05-0730-WAR-09/11/05///
     Call SETULTREG("!COSTNEG", j&)
     Call CIERRARCH("!COSTNEG")
     If j& > 0 Then 'SI HAY ALGUN NEGATIVO
     ' PRESENTAR MENSAJE ADVERTENCIA CON LIST DE LOS CODIGOS CON STOCK INSUFICIENTE
       Call COMUNI("Esta Transacción producirá Negativos\en el Déposito Elegido para\los Siguientes Items:\   \" + LICONEG$)
       DE2$ = ECOARCH("TADEPOSI", Chr$(DEPOMO%))
       NUDE$ = FORMATNUM$(DEPOMO%, "F2.0")
       VTB% = VENTABU%("COSTONEG/NC/TITUPAN=Items con Stock Insuficiente en Déposito: " + NUDE$ + " - " + DE2$) 'TRIM$(RASOCLI$(NC1%)))
     End If
     '\\\OT-05-0730-WAR-FIN///
   End If
   '
   Exit Function
   '*********************************************************************************
COPYCONTROLA:
   If EXISTE%(LUDAT$ + "COSTNEG.RFS") < 1 Then
      Call COPYSTRU("MATEROF", "COSTNEG")
      '
      AA1$ = LOADFILE$("COSTONEG.PRF")
85    PPXX% = InStr(AA1$, "MATEROF")
      If PPXX% > 0 Then
         Mid$(AA1$, PPXX%, 7) = "COSTNEG"
         Call SAVEFILE("COSTONEG.PRF", AA1$)
         GoTo 85
      End If
   End If
Return
'
End Function
'
Sub MOVISTOK(FECHA%, CI%, LOTE$, CMOV$, DOPRILA$, DOPRICO$, DOSELA$, DOSECO$, REFE$, NUORIT%, VUNI#, MONEII%, NCLIE, DEPO%, CANTI, Optional NUPECLIEN&, Optional NORITPE%, Optional DELARGA$, Optional DOCOMPRA$, Optional IDMOVIM$, Optional PRECONIVA#, Optional TIPOTRAN%, Optional ORDCOM_RTOABIER$)
   '
   Static DECONSUM$    ' CTX%, REBLAN$, RELIBR$, UREMOVI&
   '
   CADEC% = XVALO(CNTDC$(CI%)) ' CANTIDAD DE DECIMALES CONFIGURADOS
   '
   HOII% = HOY(HOS$)
   Screen.MousePointer = 11
   If CANTI < 0 Then
     If DECONSUM$ = "" Then
       DECONSUM$ = "MANUA"
       If Control$("", "DECONSUM") = "AUTOM" Then
         DECONSUM$ = "AUTOM"
       End If
     End If
   End If
   '
   SMP = Screen.MousePointer
   USERN% = USNBR% Mod 100
   If USERN% < 1 Then USERN% = 1
   Call CARDEPOS
   '
   'Call VERISAL(CI%)               ' verificar saldos cruzados
   '
   DEPOXI% = DEPO%
   If DEPO% = 99 Then DEPOXI% = DEPOPRE%(CI%)
   If DECONSUM$ = "AUTOM" Then
     If CANTI < 0 Then
       DEPII% = DEPO%
       DEPOXI% = DEPOSIBLE%(CI%, CANTI, DEPII%)   'Obtengo el  deposito de movimiento.
     End If
   End If
   '
   'CANTIX = (Int(10000 * CANTI + 0.5)) / 10000
   CANTIX = ROUND#(CANTI, CADEC%)
   '
   If NOCLOSE% <> 1 Then Call APERBASDAT("STOCKS")
   '
   ' If FECHA% > HOII% + 30 Or FECHA% < HOII% - 100 Then FECHA% = HOII%
   
   'sql
   '\\\OT-08-0839-OD-12/08/08///
   NUMPVTA1% = 0
   DOPRINUM1& = 0
   If Left$(CMOV$, 1) <> "*" Then
     TABLAMR$ = "Movisto" 'SI TIENE CMOV$ = "*" GRABA EN MOVIREPA
     If UCase$(CMOV$) = "RT" Then ' AGREGADO PARA LEER DATOS NUMERICOS DE REMITO Y PUNTO DE VENAT.
        If Left$(UCase$(DOPRILA$), 2) = "RT" Then
          NUMPVTA1% = XVALO(Mid$(DOPRILA$, 6, 4))
          DOPRINUM1& = XVALO(Mid$(DOPRILA$, 11, 8))
        End If
     End If
     '
    Else
      TABLAMR$ = "MOVIREPA"
   End If
   
   '\\\OT-08-0839-OD-FIN///

'   Dim Movisto As ADODB.Recordset
'   Set Movisto = New ADODB.Recordset
'   Movisto.CursorType = adOpenDynamic
'   Movisto.LockType = adLockPessimistic
'   strsql = "Select *  from " & TABLAMR$
'   strsql = strsql + " Where Cod_Inte = " & CI%
'   strsql = strsql + " and DepoMovi =" & DEPOXI%
'   Movisto.Open strsql, FLEXCONN, , , adCmdText
   DOPRILA1$ = TRIM$(DOPRILA$)
   Dim Movisto As ADODB.Recordset
   Set Movisto = New ADODB.Recordset
   STRSQL = "Select *  from " & TABLAMR$
   STRSQL = STRSQL + " WHERE DoPriLar = '" & DOPRILA1$ & "'"
25 On Error Resume Next
   'Movisto.Open strsql, FLEXCONN, , , adCmdText
   Movisto.Open FILTSQL$(STRSQL), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
   '
   If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
      On Error GoTo 0
      Call CapturaErrorOpen
      GoTo 25
   End If
   On Error GoTo 0
   '
   With Movisto
     .AddNew   ' YA CONVERTIDO
     On Error Resume Next
     !CodiEmpr = CodiEmp%
     On Error GoTo 0
     !NUMPVTA = NUMPVTA1%
     !DOPRINUM = DOPRINUM1&
     !Fechmov = CVDAT(FECHA%)
     !Cod_Inte = CI%
     !Cod_Exte = Left$(CODEXT$(CI%), 24)
        DESCIRIX1$ = Left(DESCRIT0$(CI%), 64)
     !Descrip = DESCIRIX1$
     !IdenLote = Left$(LOTE$, 16)
     !CodiMovi = Left$(CMOV$, 2)
     !DoPriCor = Left$(DOPRICO$, 12)
       DOPRILAX$ = TRIM$(DOPRILA$): If DOPRILAX$ = "" Then DOPRILAX$ = DOPRICO$
     !DoPriLar = Left(DOPRILAX$, 24)
     !DoSecCor = Left$(DOSECO$, 12)
       DOSECLAX$ = DOSELA$: If DOSECLAX$ = "" Then DOSECLAX$ = DOSECO$
     !DoSecLar = Left(DOSECLAX$, 24)
     !nuoritem = NUORIT%
     !DOCUORIG = Left$(DOCOMPRA$, 24)
     !IdenMovi = Left$(IDMOVIM$, 32)
     !NUPEDCLI = 0: !NumeOrdPed = 0
     !Descri_Lar = DELARGA$ ' DESCRIPCION DE TEXTO LIBRE
        If TRIM$(DELARGA$) = "" Then !Descri_Lar = DESCIRIX1$
       If Left$(DOSECO$, 3) = "PC." Then
           !NUPEDCLI = XVALO(Mid$(DOSECO$, 4))
         Else
           On Error Resume Next
           !NUPEDCLI = NUPECLIEN
           !NumeOrdPed = NORITPE%
           On Error GoTo 0
       End If
     '\\\OT-08-0751-OD-12/09/08///
     !ReferCor = Left(REFE$, 64)
     '\\\OT-08-0751-OD-FIN///
        If Left$(REFE$, 7) = "RECHAZO" Or Left$(REFE$, 8) = "FALTANTE" Then
          locon$ = AJUSTI$(LOTE$, 24) + FORMATNUM$(Abs(CANTIX), "F8." & CNTDC$ & "")
          !LotConsum = locon$
        End If
     !ValoUnit = VUNI#
     !VALTOTIT = ROUND(VUNI# * (ROUND(Abs(CANTIX), CADEC%)), 2)
     !MoneVal = MONEII%
       On Error Resume Next      ' POR SI NO ESTA DEFINIDO EL CAMPO
     !TipoCam = TICAMONE(MONEII%)
       On Error GoTo 0
     !DepoMovi = DEPOXI%
     If NCLIE >= 0 Then
         !Nume_Clie = NCLIE
         !Nume_Prov = 0
       Else
         !Nume_Clie = 0
         !Nume_Prov = Abs(NCLIE)
     End If
     If CANTIX >= 0 Then
         !CantIngre = ROUND#(CANTIX, CADEC%) ' CADEC% = CANTIDAD DE DECIMALES CONFIGURADOS
         !Cantsalid = 0
       Else
         !CantIngre = 0
         !Cantsalid = ROUND#(Abs(CANTIX), CADEC%)
         
     End If
     !CantSaldo = 0
     !FeReMovi = Now
     !NumUsuar = USERN%
     !VALOUNITIVA = PRECONIVA#
     !VALTOTITIVA = ROUND(PRECONIVA# * (ROUND(Abs(CANTIX), CADEC%)), 2)
     !VERIFICADOPOR = 0
     !FHor_Verif = FETEXCOR1$(CVINT(NUNCA$))
     On Error Resume Next
     !TIPOTRAN = TIPOTRAN% 'TIPO DE TRANSFERENCIA AGREGADO POR ALEJPA
     On Error GoTo 0
     !OCOMPRA_RTOABIERTO = Left$(ORDCOM_RTOABIER$, 64)
     
     .Update
   End With
   '
  
   If NOCLOSE% <> 1 Then
     Movisto.Close
     Set Movisto = Nothing
   End If
   '
   If APROBOLE% < 1 Then
      If ESTRAZABLE%(CI%) = 1 Then
         Call ACSALOTE(CI%, LOTE$, "NOMESS")
      End If
   End If
   '
   Screen.MousePointer = SMP
   '
End Sub
'
Function DEPOSIBLE%(CI%, CAMOVI, DEPO%)
   '
   Static DPNG%
   '
   If DPNG% = 0 Then
     DPNG% = XVALO(Control$("", "DEPONEGA"))
     If DPNG% < 1 Then DPNG% = (-1)
   End If
   '
   DEPOX% = DEPO%
   Call CARDEPOS
   CASALIDA = Abs(CAMOVI)
   If DEPOX% > 0 Then
      '
      RGG& = CI% + 1
      STRI$ = REGLEIDO$("SALDEPOS", RGG&)
      SUCX% = IDEPOS%(DEPOX%)
      If CVS(Mid$(STRI$, 4 * SUCX% - 3, 4)) >= CASALIDA - 0.05 Then
         DEPOSI% = DEPOX%
         GoTo 99
      End If
      '
      DEPOX% = 0
      GoTo 20
      '
    Else
      '
20    RGG& = CI% + 1
      STRI$ = REGLEIDO$("SALDEPOS", RGG&)
      For JII% = 1 To NUSU%
        If CVS(Mid$(STRI$, 4 * JII% - 3, 4)) >= CASALIDA - 0.05 Then
           DEPOSI% = COSU%(JII%)
           GoTo 99
        End If
      Next JII%
      '
      Call MENSERR(24, "Codigo " + TRIM$(CODEXT$(CI%)) + "\Stock Insuficiente en Depósitos Examinados.\  \Se Producirá un Negativo de Stock.")
      DEPOSI% = DPNG%
      If DEPOSI% < 1 Then DEPOSI% = COSU%(1)
      GoTo 99
      '
   End If
   '
99 DEPOSIBLE% = DEPOSI%
   '
End Function


Sub VAREMAT()
   '
   NRO = 1 + ULTREG&("MOVISTO")
   Call CIERRARCH("MOVISTO")
   Call BLOQARCH("MOVISTO")
   Call BLOQARCH("TRANSFE")
   '
   For U& = 1 To ULTREG&("VAREMAT")
      X$ = REGLEIDO$("VAREMAT", U&)
      CI1% = CVI(Mid$(X$, 5, 2))
      DP1% = Asc(Mid$(X$, 3, 1))
      CI2% = CVI(Mid$(X$, 5, 2))
      DP2% = Asc(Mid$(X$, 4, 1))
      FECHA% = CVI(Mid$(X$, 1, 2))
      If FECHA% < 1 Then FECHA% = HOY(HO$)
      LOTE$ = Mid$(X$, 15, 16)
      CAN = CVS(Mid$(X$, 11, 4))
      REFE0$ = ""
      '
      If CI1% > 0 Then
       If CI1% <= NUMAS% Then
        If DP1% > 0 And DP1% < 99 Then
         If ECOARCH$("ECODEP", Chr$(DP1%)) <> "" Then
          If CI2% > 0 Then
           If CI2% <= NUMAS% Then
            If DP2% > 0 And DP2% < 99 Then
             If ECOARCH$("ECODEP", Chr$(DP2%)) <> "" Then
              If CAN >= 0.1 Then
               '
               NC% = 0
               VUNI# = 0
               MONE$ = ""
               CMOV$ = "TR"
               DOPRI$ = "TR." + TRIM$(Str$(NRO))
               DOSEC$ = "TR." + Mid$(HO$, 8, 1) + Mid$(HO$, 4, 2) + Left$(HO$, 2)
               REFE$ = "Vale Mat.a Dep." + TRIM$(Str$(DP2%))
               CANTI = (-1) * CAN
               'Call MovistoK(FECHA%, CI1%, LOTE$, CMOV$, DOPRI$, DOSEC$, REFE$, VUNI#, MONE$, NC%, DP1%, CANTI)
               '
               COD$ = LTrim$(RTrim$(CODEXT$(CI1%)))
               REFE$ = "Vale Mat.de Dep." + TRIM$(Str$(DP1%))
               CANTI = CAN
               'Call MovistoK(FECHA%, CI2%, LOTE$, CMOV$, DOPRI$, DOSEC$, REFE$, VUNI#, MONE$, NC%, DP2%, CANTI)
               '
               XXX$ = REGBLAN$("TRANSFE")
               Call REPLA(XXX$, MKI$(FECHA%), 1, 2)
               Call REPLA(XXX$, Chr$(DP1%), 3, 1)
               Call REPLA(XXX$, Chr$(DP2%), 4, 1)
               Call REPLA(XXX$, MKI$(CI1%), 5, 2)
               Call REPLA(XXX$, MKI$(CI2%), 7, 2)
               Call REPLA(XXX$, LOTE$, 9, 16)
               Call REPLA(XXX$, LOTE$, 25, 16)
               Call REPLA(XXX$, MKS$(CAN), 41, 4)
               Call REPLA(XXX$, MKS$(NRO), 45, 4)
               Call REGAPP("TRANSFE", XXX$)
               '
               Call REPLA(X$, String$(32, 0), 1, 32)
               Call GRAREG("VAREMAT", X$, U&)
               '
              End If
             End If
            End If
           End If
          End If
         End If
        End If
       End If
      End If
      '
29 Next U&
   '
   Call CIERRARCH("MOVISTO")
   Call CIERRARCH("TRANSFE")
   '
   j& = 0
   For U& = 1 To ULTREG&("VAREMAT")
     X$ = REGLEIDO$("VAREMAT", U&)
     If X$ <> String$(32, 0) Then
       j& = j& + 1
       Call GRAREG("VAREMAT", X$, j&)
     End If
   Next U&
   Call SETULTREG("VAREMAT", j&)
   Call CIERRARCH("VAREMAT")
   '
   OKEY% = 1
   If j& > 0 Then
     OKEY% = (-1)
   End If
   '
End Sub
'
Function ProNroComprobante(TM$) 'Tipo de movimiento (CodiMovi) en movisto
    Call ABRESTOCKS
    '
'    SQLX$ = "SELECT max(cast(substring(DoPriCor, 4,6)as int))as Maximo FROM MOVISTO"
'    SQLX$ = SQLX$ + " WHERE CodiMovi = '" & TM$ & "'"
    sqlx$ = "SELECT cast(substring(DoPriCor, 4,6)as int) as Maximo FROM MOVISTO WITH(NOLOCK)"
    sqlx$ = sqlx$ + " WHERE CodiMovi = '" & TM$ & "' order by maximo desc"
    UNUM& = 0
    Dim MIRS As ADODB.Recordset
    On Error Resume Next
    Set MIRS = FLEXCONN.Execute(FILTSQL$(sqlx$))
    If Not IsNull(MIRS!maximo) Then
        UNUM& = MIRS!maximo
    End If
    MIRS.Close
    On Error GoTo 0
    ProNroComprobante = 1 + UNUM&
    Set MIRS = Nothing
End Function

'
Sub NUENTRECONSI(NPROVE%, Optional NROCOMX$)
        '
        '\\\OT-5-861-FG-09/12/05///
        Dim CANPREASI()
        '\\\OT-5-861-FG-FIN///
        '
        LLAMACOMPRA% = 0
        If NPROVE% > 0 Then LLAMACOMPRA% = 1
        '
        UTILIZAREMITOVENTAS% = 0
        'MODFIICACION PARA QUE CONSIDERE SI UTILIZA EL FORMULARIO DE REMITOS PARA CONSIGNACIONNES
        If Control("OCOMPRA", "UTIRTVTA") = "SI" Then
             UTILIZAREMITOVENTAS% = 1
             FORIPRE$ = TRIM$(Control$("", "FOREMIFA"))
             If FORIPRE$ = "" Then
               Call COMUNI("Error de Configuración:\  \Falta Definir Formulario\de Impresión de Remito.")
               GoTo 9999
             Else
               LU$ = LUDAT$
               If EXISTE%(LU$ + FORIPRE$ + ".FRM") <> 1 Then
                 RMCC$ = REPLACAR$(LU$ + FORIPRE$ + ".FRM", "\", "/")
                 Call COMUNI("Error de Configuración:\  \No Existe Formulario de Impresión\de Remito de Facturación:\  \'" + RMCC$ + "'.")
                 GoTo 9999
               End If
             End If
        Else
             FORIPRE$ = TRIM$(Control$("", "FOREMICO"))
             If FORIPRE$ <> "" Then
                 LU$ = LUDAT$
                 If EXISTE%(LU$ + FORIPRE$ + ".FRM") <> 1 Then
                     RMCC$ = REPLACAR$(LU$ + FORIPRE$ + ".FRM", "\", "/")
                     Call COMUNI("Error de Configuración:\  \No Existe Formulario de Impresión\de Remito en Consignación:\  \'" + RMCC$ + "'.")
                     GoTo 9999
                 End If
             End If
        End If
        '
        ' PONE EN CERO LOS CAMPOS DE CANTIDADES QUE CONTENGAN UN NULO
        Call PoneEnCeroNull("MACONSIG", "cant_asig", REGAF&)
        Call PoneEnCeroNull("MACONSIG", "cant_entreg", REGAF&)
        Call PoneEnCeroNull("MACONSIG", "cant_rendid", REGAF&)
        Call PoneEnCeroNull("MACONSIG", "cant_devu", REGAF&)
        '
        DEPOCON% = XVALO(Control$("", "DEPOCONS"))
        DepOrig% = DEPOCON%
        DEPOSI$ = ECOARCH$("TADEPOSI", Chr$(DepOrig%))
        '
        'Call BLIMESS("Revisando Archivo de Movimientos de Stock")
1100    Screen.MousePointer = 11
        CMO$ = TRIM$(Control$("", "CODRECON"))
        If CMO$ = "" Then CMO$ = "RC"
        '
1002    If UTILIZAREMITOVENTAS% > 0 Then
            NUPUVEN$ = TRIM$(Control$("", "PREFIREM"))
            If XVALO(NUPUVEN$) = 0 Then NUPUVEN$ = "0001"
            NROPRO& = 0
            Call NUREMIFA(NROPRO&, NUPUVEN$)
            REMI$ = CMO$ + "." + TRIM$(Str$(NROPRO&))
        Else
            UNRO& = ProNroComprobante(CMO$)
            REMI$ = CMO$ + "-" + TRIM$(Str$(UNRO&))
        End If
        '
        Call FRESHALL
        '
1105    Call CIERRARCH("*.*")
        '\\\OT-5-861-FG-9/12/05///
        ReDim CANPREASI(NUMAS%)
        '\\\OT-5-861-FG-FIN///
        '
        Screen.MousePointer = 1
        If NPROVE% > 0 Then
             EBOL$ = MKI$(NPROVE%)
          Else
             EBOL$ = OBJPLA$("SELPROVE", "")
             If EBOL$ = "" Then
                GoTo 9999
             End If
        End If
        '
        NC% = CVI(EBOL$)
        If NC% < 1 Or ECOARCH$("PROVED1", MKI$(NC%)) = "" Then
            Call MENSERR(24, "Proveedor Inexistente")
            NPROVE% = 0
            GoTo 1105
        End If
        '
        NC1% = (-1) * NC%
        '
        Y$ = PEDY$: If Y$ = "" Then Y$ = Space$(128)
        Call REPLA(Y$, MKS$(NROBOL&), 1, 4)
        Call REPLA(Y$, EBOL$, 7, 2)
        Call REPLA(Y$, MKI$(HOY(HO$)), 5, 2)
        Call REPLA(Y$, RASOCLI$(NC1%), 9, 32)
        Call REPLA(Y$, DOMICLI$(NC1%), 41, 32)
        CLOCA$ = TRIM$(CPOSCLI$(NC1%)): If CLOCA$ <> "" Then CLOCA$ = CLOCA$ + " "
        Call REPLA(Y$, CLOCA$ + LOCACLI$(NC1%), 73, 32)
1110    Call REPLA(Y$, REMI$, 105, 16)
        Call REPLA(Y$, Chr$(DepOrig%), 121, 1)
        '
1111    PEDY$ = OBJPLA$("REMICON-ENCB", Y$)
        If PEDY$ = "" Then GoTo 9998
        FF% = CVI(Mid$(PEDY$, 5, 2))
        REMI$ = TRIM$(Mid$(PEDY$, 105, 16))
        If TRIM$(REMI$) = "" Then
            Y$ = PEDY$
            Call MENSERR(24, "Falta Numero de Remito")
            GoTo 1111
        End If
        '
        If UTILIZAREMITOVENTAS% > 0 Then
            REMIMANUAL$ = TRIM$(InputBox("Numero de Remito", "Ingreso Manual", Str$(NROPRO&)))
            If REMIMANUAL$ <> "" Then
                REMI$ = CMO$ & "." & REMIMANUAL$
            End If
        End If
        '
        STRSQL = "Select Cod_Inte,DoPriCor from Movisto"
        STRSQL = STRSQL + " where DoPriCor = '" & REMI$ & "'"
        Dim MIRS As ADODB.Recordset
        Set MIRS = FLEXCONN.Execute(FILTSQL$(STRSQL))
        'With UNRECONTMP
        With MIRS
          '.FindFirst "DoPriCor = '" & REMI$ & "'"
          'If .NoMatch = False Then
          If Not (.EOF And .BOF) Then
            Call MENSERR(24, "Número de Remito ya Registrado.")
            GoTo 1111
          End If
        End With
        MIRS.Close
        Set MIRS = Nothing
        '
        DepOrig% = Asc(Mid$(PEDY$, 121, 1))
        '
        If DepOrig% < 1 Or ECOARCH$("TADEPOSI", Chr$(DepOrig%)) = "" Then
          Call MENSERR(24, "Depósito de Salida no Válido")
          DepOrig% = XVALO(Control$("", "DEPOCONS"))
          GoTo 1110
        End If
        '
        DEPOSI$ = ECOARCH$("TADEPOSI", Chr$(DepOrig%))
        '
1112    Call CIERRARCH("*.*")
        Call BLANARCH("!BOLREDET")
        '
        'OPCION DE AGRUPA (GABAL)
        If Control("OCOMPRA", "AGRUCONS") = "SI" Then
              Call LISTMATPROV(NC%, DepOrig%, NROCOMX$)
              GoTo 1113 ' VA A LA VENTANA
        End If
        '
        Screen.MousePointer = 11
        FFOO% = FECHANUM("01/01/13")
        HAYOC% = 0
        Call APERBASDAT("STOCKS")
        sqlx$ = "SELECT * FROM MACONSIG WITH(NOLOCK)"
        sqlx$ = sqlx$ + " WHERE Cant_Asig > Cant_Entreg "
        sqlx$ = sqlx$ + " AND Cant_Asig >= 0.05 "
        sqlx$ = sqlx$ + " AND Npro_Ocom = " & NC% & " "
        If NROCOMX$ <> "" Then
           sqlx$ = sqlx$ + "AND Nro_Ocom Like '" & NROCOMX$ & "%'"
        End If
        '
        Dim MACONSITMP As ADODB.Recordset
        Set MACONSITMP = FLEXCONN.Execute(FILTSQL$(sqlx$))
        With MACONSITMP
          Do While .EOF = False
           CAPENE# = XVALO(!cant_asig) - XVALO(!cant_entreg)
           NUOC& = XVALO(Mid$(SAFETEXT$(!Nro_Ocom), 4, 6))
           NORCOMPRA$ = SAFETEXT$(!Nro_Ocom)
           If (STATOCOM%(NUOC&) < 8 And STATOCOM%(NUOC&) >= 0) Then
               '
               CI% = XVALO(!Cod_Inte)
               'STODISPO# = CDbl(STOCKACT(CI%))      CODIGO ANTERIOR !!!
               '\\\OT-5-861-FG-9/12/05///
               STODISPO# = CDbl(STODEPOS(CI%, DepOrig%) - CANPREASI(CI%))
               '\\\OT-5-861-FG-FIN///
               If STODISPO# < 0 Then STODISPO# = 0
               CANENTRE# = CAPENE# ' CANTIDAD PENDIENTE SEGUN FORMULA
               QUEDAPEN# = 0       ' QUEDA PENDIENTE
               If STODISPO# < CAPENE# Then
                  CAPENE# = STODISPO#
                  QUEDAPEN# = CANENTRE# - STODISPO#   ' QUEDA PENDIENTE ENTREGA FUTURA
               End If
               '
               SOBR# = 0
               If InStr(1, EMPREAC$, "FERVI") > 0 Then
                    CONDI$ = "Cod_Inte = " & CI% & " AND NPRO_OCOM=" & NC%
                    CONDI$ = CONDI$ & " AND (CANT_ENTREG-CANT_RENDID > 0 OR CANT_ASIG > CANT_ENTREG OR CANT_DEVU <> 0)"
                    'CONDI1$ = "Cod_Inte = " & CI% & " AND (NRO_OCOM='OC-000000-00') "
                    'CONDI1$ = CONDI1$ + " AND NPRO_OCOM=" & NC% & " AND FEMI_OCOM>'" & FETEXCOR1$(FFOO%) & "'"
                    '
                    'CAS# = XVALO(VALOBADA("MACONSIG", "SUM(CANT_ASIG)", CONDI$))
                    CEN# = XVALO(VALOBADA("MACONSIG", "SUM(CANT_ENTREG)", CONDI$))
                    'CEN1# = XVALO(VALOBADA("MACONSIG", "SUM(CANT_ENTREG)", CONDI1$))
                    CRE# = XVALO(VALOBADA("MACONSIG", "SUM(CANT_RENDID)", CONDI$))
                    CDE# = XVALO(VALOBADA("MACONSIG", "SUM(CANT_DEVU)", CONDI$))
                    STOTER# = CEN# - CRE# - CDE#
               End If
               '
               Y$ = REGBLAN$("!BOLREDET")
               ESTRA& = 2 - ESTRAZABLE%(CI%) 'SE AGREGA ESTO PARA GUARDAR DE ACUERDO A LA
               Call REPLA(Y$, MKS$(ESTRA&), 1, 4) ' TRAZABILIDAD
               Call REPLA(Y$, MKI$(CI%), 83, 2)
               Call REPLA(Y$, UNIMED$(CI%), 85, 4)
               Call REPLA(Y$, MKD$(CAPENE#), 89, 8)
               Call REPLA(Y$, NORCOMPRA$, 119, 16)
               Call REPLA(Y$, MKS$(NUOC&), 151, 4)
               Call REPLA(Y$, MKD$(QUEDAPEN#), 163, 8)  ' QUEDA PENDIENTE
               Call REPLA(Y$, MKD$(CANENTRE#), 171, 8)  ' CANTIDAD PENDIENTE SEGUN FORMULA
               If InStr(1, EMPREAC$, "FERVI") > 0 Then Call REPLA(Y$, MKD$(Abs(STOTER#)), 171, 8)
               Call REPLA(Y$, MKS$(I&), 189, 4)

               HAYOC% = 1
               Call REGAPP("!BOLREDET", Y$)
               '\\\OT-5-861-FG-9/12/05///
               ' ACUMULA LO PRE-ASIGNADO POR CASO DE ITEMS QUE APARECEN VARIAS VECES
               If CI% > 0 Then
                 If CI% <= NUMAS% Then
                   CANPREASI(CI%) = CANPREASI(CI%) + CAPENE#
                 End If
               End If
               '\\\OT-5-861-FG-FIN///
           End If
           .MoveNext
          Loop
        End With
        MACONSITMP.Close
        Set MACONSITMP = Nothing
        '
555     If HAYOC% = 0 Then
            Call COMUNI("No hay Consignaciones Pendientes\para este Proveedor.")
        End If
        '
1113    Archiv$ = "ECONSIG"
        If EXISTE%("ECONSIM.PRF") > 0 Then Archiv$ = "ECONSIM" ' PARA MEDITEA DONDE SE AGREGAR LA CANT. PENDIENTE Y EL TOTAL A ENTREGAR SEGUN FORMULA
        VTB% = VENTABU%(Archiv$ & "/NC/TITUPAN=Consignación de Materiales - " + TRIM$(RASOCLI$(NC1%)))
        If VTB% < 0 Or ULTREG&("!BOLREDET") < 1 Then
            X$ = "": Y$ = "": YY$ = ""
            EBOL$ = "": PEDY$ = "": DESCRI$ = ""
            GoTo 1105
        End If
        '
        Call ASIGNANLOTE(OK%)
        If OK% < 0 Then GoTo 1113
        '
        For YX& = 1 To ULTREG&("!BOLREDET")
          YY$ = REGLEIDO$("!BOLREDET", YX&)
          CI% = CVI(Mid$(YY$, 83, 2))
          CAN# = CVD(Mid$(YY$, 89, 8))
          If CI% > 0 Then
            If CI% <= NUMAS% Then
              If CAN# > 0.005 Then
                GoTo 11141 'GoTo 1114 \\\OT-05-0730-WAR-10/11/05///
              End If
            End If
          End If
        Next YX&
        Call MENSERR(24, "Consignación Cancelada !!!\Faltan Cantidades o Códigos.")
        GoTo 1113
        '\\\OT-05-0730-WAR-10/11/05/// COMENTE LO SIG.
'1114    If TRIM$(FORIPRE$) <> "" Then
'           DTP$ = OBJPLA$("TRANSPOR", DATRANSPO$)
'           If DTP$ = "" Then GoTo 1113         ' CANCELADO
'           DATRANSPO$ = DTP$
'        End If
        '\\\OT-05-0730-WAR-FIN/// HASTA ACA COMENTE
11141   CONTA% = ULTREG&("!BOLREDET")
        '
1115    If CONTA% > 0 Then
        '
        ' ESTE BLOQUE AGREGADO 9/11/05 POR CONTROL NEGATIVOS DE STOCK
        CAITMO% = 0
        For YX& = 1 To ULTREG&("!BOLREDET")
             YY$ = REGLEIDO$("!BOLREDET", YX&)
             CI% = CVI(Mid$(YY$, 83, 2))
             CAN# = CVD(Mid$(YY$, 89, 8))
             If CI% > 0 Then
               If CI% <= NUMAS% Then
                 If CAN# > 0.005 Then
                   CAITMO% = CAITMO% + 1
                   CIMOVIM%(CAITMO%) = CI%
                   CAMOVIM(CAITMO) = (-1) * CAN#
                 End If
               End If
             End If
          Next YX&
          If PUEDEMOVI%(DepOrig%) < 0 Then GoTo 1113 ' CANCELA SALIDA
          '
          '\\\OT-05-0730-WAR-10/11/05///
1114      If TRIM$(FORIPRE$) <> "" Then
             DTP$ = OBJPLA$("TRANSPOR", DATRANSPO$)
             If DTP$ = "" Then GoTo 1113         ' CANCELADO
             DATRANSPO$ = DTP$
          End If
          '\\\OT-05-0730-WAR-FIN///
          '
          Call BLOQARCH("MACONSIG")
          Screen.MousePointer = 11  'Call BLIMESS("Registrando Recepcion ...")
          '
          'OPCION DE AGRUPA (GABAL)
          If Control("OCOMPRA", "AGRUCONS") = "SI" Then
              Call REMICONSARMA(NC%, REMI$, DepOrig%, NROOCOM$)
              GoTo 1130 ' IMPRIME
          End If
          '
          AINDI% = 0
          NUITEM = 0
          For YX& = 1 To ULTREG&("!BOLREDET")
             YY$ = REGLEIDO$("!BOLREDET", YX&)
             CI% = CVI(Mid$(YY$, 83, 2))
             CAN# = CVD(Mid$(YY$, 89, 8))
             LOTE$ = TRIM$(Mid$(YY$, 69, 12))
             NUMECOMPRA$ = TRIM$(Mid$(YY$, 119, 16)) ' NUM.O.C. COMPLETO (TEXTO)
             If CI% > 0 Then
               If CI% <= NUMAS% Then
                 If CAN# > 0.005 Then
                    NUOC& = CVS(Mid$(YY$, 151, 4))
                    DOSEC$ = ""
                    Dim MaConsig As ADODB.Recordset
                    Set MaConsig = New ADODB.Recordset
                    sqlx$ = " select * from MACONSIG"
                    'SQLX$ = SQLX$ + " WHERE substring(Nro_Ocom, 4, 6) =  " & NUOC & " " '\\\OT-06-0206-WAR-29/05/06///
                    sqlx$ = sqlx$ + " WHERE Nro_Ocom =  '" & NUMECOMPRA$ & "' " 'MODIFICADO PARA QUE COMPARE NUMERO DE ORDEN DE COMPRA MAS EL ORDEN DEL ITEM
                    sqlx$ = sqlx$ + " AND Cod_Inte = " & CI% & ""
                    sqlx$ = sqlx$ + " AND Npro_Ocom = " & NC% & " "
                    sqlx$ = sqlx$ + " AND Cant_Asig > Cant_Entreg "
                    'SE AGREGA LA LINEA DE ARRIBA PARA NETEAR SOBRE LO QUE REALMENTE ESTA PENDIENTE
26                  On Error Resume Next
                    MaConsig.Open FILTSQL$(sqlx$), FLEXCONN, adOpenDynamic, adLockPessimistic, adCmdText
                    If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
                       On Error GoTo 0
                       Call CapturaErrorOpen
                       GoTo 26
                    End If
                    On Error GoTo 0
                    '
                    With MaConsig
                      If Not (.EOF And .BOF) Then 'si existe, actualiza
                           RESTO# = 0
                           DOSEC$ = SAFETEXT$(!Nro_Ocom)
                           CANTPEND# = XVALO(!cant_asig) - XVALO(!cant_entreg) 'AGREGADO
                           CANTOTAL# = CAN#
                           'SI LA CANTIDAD CONSIGANDA ES MENOR A LA ENTREGADA  SOLO ENTREGA EL PENDIENTE
                           If CAN# > CANTPEND# Then
                               RESTO# = CAN# - CANTPEND# 'AGREGADO CALCULA EL SALDO RESTANTE
                               CAN# = CANTPEND# '
                           End If 'AGREGADO
                           !cant_entreg = CAN# + XVALO(!cant_entreg)
                           '
                           If FF% > CVINT(!FECHA_ENTREG) Then
                             !FECHA_ENTREG = CVDAT(FF%)
                           End If
                           !NREMI_CONSIG = TRIM$(Mid$(REMI$, 4))
                           .Update
                           If RESTO# > 0 Then
                              CAN# = RESTO#
                              GoTo 70
                           End If
                           GoTo 1120
                        Else
70                         .AddNew   ' YA CONVERTIDO
                           On Error Resume Next
                           !CodiEmpr = CodiEmp%
                           On Error GoTo 0
                           !Femi_Ocom = CVDAT(FF%)
                           !Npro_Ocom = NC%
                           !Raso_Prov = RASOCLI(NC% * (-1))
                            NOCOM$ = TRIM$(Str$(NUOC&))
                            While Len(NOCOM$) < 6: NOCOM$ = "0" + NOCOM$: Wend
                            NOCOM$ = "OC-" + Left$(NOCOM$, 6) + "-00"
                           !Nro_Ocom = NOCOM$
                           !Cod_Inte = CI%
                           !Cod_Exte = CODEXT(CI%)
                           !NREMI_CONSIG = TRIM$(Mid$(REMI$, 4))
                           !Unidad = UNIMED$(CI%)
                           !cant_asig = CAN#
                           !cant_entreg = CAN#
                           !cant_rendid = 0
                           !cant_devu = 0
                           !FECH_ENTREG = CVDAT(FF%)
                           !LOTE_OCOM = LOTE$
                           NORD% = 0
                           If TRIM$(NROCOMX$) <> "" Then
                              NORD% = 1 + XVALO(VALOBADA("MACONSIG", "MAX(NUORDITEM)", "Nro_Ocom Like '" & NROCOMX$ & "%'", "NOMESS"))
                           End If
                           !NUORDITEM = NORD%
                           .Update
                      End If
                    End With
                 End If
               End If
             End If
             '
             ' If RESTO# > 0 Then CAN# = CANTOTAL#  desprolijo y peligroso !!!!!
             CAN# = CVD(Mid$(YY$, 89, 8))
1120         CANTI = (-1) * CAN#
             REFE$ = REMI$
             NUITEM = NUITEM + 1
             NUORIT% = NUITEM
             Call MOVISTOK(FF%, CI%, LOTE$, CMO$, REMI$, REMI$, DOSEC$, DOSEC$, REFE$, NUORIT%, 0, 0, NC1%, DepOrig%, CANTI)
             '
          Next YX&
          '
          
'''          For YX& = 1 To ULTREG&("!BOLREDET")
'''             YY$ = REGLEIDO$("!BOLREDET", YX&)
'''             CI% = CVI(Mid$(YY$, 83, 2))
'''             CAN# = CVD(Mid$(YY$, 89, 8))
'''             If CI% > 0 Then
'''               If CI% <= NUMAS% Then
'''                 If CAN# > 0.005 Then
'''                   CAITMO% = CAITMO% + 1
'''                   CIMOVIM%(CAITMO%) = CI%
'''                   CAMOVIM(CAITMO) = (-1) * CAN#
'''                 End If
'''               End If
'''             End If
'''          Next YX&
'''          If PUEDEMOVI%(DepOrig%) < 0 Then GoTo 1113 ' CANCELA SALIDA
'''          '
'''          '\\\OT-05-0730-WAR-10/11/05///
'''1114      If TRIM$(FORIPRE$) <> "" Then
'''           DTP$ = OBJPLA$("TRANSPOR", DATRANSPO$)
'''           If DTP$ = "" Then GoTo 1113         ' CANCELADO
'''           DATRANSPO$ = DTP$
'''          End If
'''          '\\\OT-05-0730-WAR-FIN///
'''          '
'''          Call BLOQARCH("MACONSIG")
'''          Screen.MousePointer = 11  'Call BLIMESS("Registrando Recepcion ...")
'''          '
'''          'OPCION DE AGRUPA (GABAL)
'''          If Control("OCOMPRA", "AGRUCONS") = "SI" Then
'''              Call REMICONSARMA(NC%, REMI$, DepOrig%)
'''              GoTo 1130 ' IMPRIME
'''          End If
'''          '
'''          AINDI% = 0
'''          NUITEM = 0
'''          For YX& = 1 To ULTREG&("!BOLREDET")
'''             YY$ = REGLEIDO$("!BOLREDET", YX&)
'''             CI% = CVI(Mid$(YY$, 83, 2))
'''             CAN# = CVD(Mid$(YY$, 89, 8))
'''             LOTE$ = TRIM$(Mid$(YY$, 69, 12))
'''             NUMECOMPRA$ = TRIM$(Mid$(YY$, 119, 16)) ' NUM.O.C. COMPLETO (TEXTO)
'''             If CI% > 0 Then
'''               If CI% <= NUMAS% Then
'''                 If CAN# > 0.005 Then
'''                    NUOC& = CVS(Mid$(YY$, 151, 4))
'''                    DOSEC$ = ""
'''
'''                    Dim MaConsig As ADODB.Recordset
'''                    Set MaConsig = New ADODB.Recordset
'''                    SQLX$ = " select * from MACONSIG"
'''                    'SQLX$ = SQLX$ + " WHERE substring(Nro_Ocom, 4, 6) =  " & NUOC & " " '\\\OT-06-0206-WAR-29/05/06///
'''                    SQLX$ = SQLX$ + " WHERE Nro_Ocom =  '" & NUMECOMPRA$ & "' " 'MODIFICADO PARA QUE COMPARE NUMERO DE ORDEN DE COMPRA MAS EL ORDEN DEL ITEM
'''                    SQLX$ = SQLX$ + " AND Cod_Inte = " & CI% & ""
'''                    SQLX$ = SQLX$ + " AND Npro_Ocom = " & NC% & " "
'''26                  On Error Resume Next
'''                    MaConsig.Open FILTSQL$(SQLX$), FLEXCONN, adOpenDynamic, adLockPessimistic, adCmdText
'''                    If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
'''                       On Error GoTo 0
'''                       Call CapturaErrorOpen
'''                       GoTo 26
'''                    End If
'''                    On Error GoTo 0
'''                    '
'''                    With MaConsig
'''                      If Not (.EOF And .BOF) Then 'si existe, actualiza
'''                           DOSEC$ = SAFETEXT$(!Nro_Ocom)
'''                           TOENTRE# = XVALO(!Cant_Entreg) + CAN#
'''                           !Cant_Entreg = TOENTRE#
'''
'''                           If FF% > CVINT(!Fecha_Entreg) Then
'''                             !Fecha_Entreg = CVDAT(FF%)
'''                           End If
'''                           !NRemi_Consig = TRIM$(Mid$(REMI$, 4))
'''                           .Update
'''                           GoTo 1120
'''                         Else
'''                           .AddNew   ' YA CONVERTIDO
'''                           On Error Resume Next
'''                           !CodiEmpr = CodiEmp%
'''                           On Error GoTo 0
'''                           !Femi_Ocom = CVDAT(FF%)
'''                           !Npro_Ocom = NC%
'''                           !Raso_Prov = RASOCLI(NC% * (-1))
'''                            NOCOM$ = TRIM$(Str$(NUOC&))
'''                            While Len(NOCOM$) < 6: NOCOM$ = "0" + NOCOM$: Wend
'''                            NOCOM$ = "OC-" + Left$(NOCOM$, 6) + "-00"
'''                           !Nro_Ocom = NOCOM$
'''                           !Cod_Inte = CI%
'''                           !COD_EXTE = CODEXT(CI%)
'''                           !NRemi_Consig = TRIM$(Mid$(REMI$, 4))
'''                           !Unidad = UNIMED$(CI%)
'''                           !Cant_Asig = CAN#
'''                           !Cant_Entreg = CAN#
'''                           !Fech_Entreg = CVDAT(FF%)
'''                           !LOTE_OCOM = LOTE$
'''                           .Update
'''                      End If
'''                    End With
'''                 End If
'''               End If
'''             End If
'''             '
'''1120         CANTI = (-1) * CAN#
'''             REFE$ = REMI$
'''             NUITEM = NUITEM + 1
'''             NUORIT% = NUITEM
'''             Call MOVISTOK(FF%, CI%, LOTE$, CMO$, REMI$, REMI$, DOSEC$, DOSEC$, REFE$, NUORIT%, 0, 0, NC1%, DepOrig%, CANTI)
'''             '
'''          Next YX&
          
          MaConsig.Close
          Set MaConsig = Nothing
          '
          Call CIERRARCH("MACONSIG")
          '
          Call FRESHECHO("IREPRIPE")
          Screen.MousePointer = 1
          '
1130    If UTILIZAREMITOVENTAS% > 0 Then
             CMO2$ = TRIM$(Control$("", "CODREFAC"))
             If CMO2$ = "" Then CMO2$ = "RT"
             While Len(NUPUVEN$) < 4: NUPUVEN$ = "0" + NUPUVEN$: Wend
             NROREMITO$ = TRIM$(Mid$(REMI$, 4)): While Len(NROREMITO$) < 8: NROREMITO$ = "0" + NROREMITO$: Wend
             NUREMLAR$ = CMO2$ & "-X-" & NUPUVEN$ & "-" & NROREMITO$
             NUREMCOR$ = CMO2$ & "." & Str$(XVALO(NUPUVEN$)) & "-" & Str$(XVALO(NROREMITO$))
             Call MOVISTOK(FF%, 0, "", CMO2$, NUREMLAR$, NUREMCOR$, REMI$, REMI$, "CONSIGNACION A PROVEEDOR: " & RASOCLI$((-1) * NC%), 1, 0, 0, ((-1) * NC%), 1, 1)
        End If

        If TRIM$(FORIPRE$) <> "" Then
            '
            FEX = FF%
            FECHDOC$ = FECHATEX$(FEX)
            NUMEDOC$ = TRIM$(REMI$)
            TIPODOC$ = "Remito de Traslado"
            '
            CODPARAM$(1) = "TRANSPOR": DOCPARAM$(1) = TRIM$(Mid$(DTP$, 1, 32))
            CODPARAM$(2) = "DOMI-TRA": DOCPARAM$(2) = TRIM$(Mid$(DTP$, 33, 32))
            CODPARAM$(3) = "CUIT-TRA": DOCPARAM$(3) = TRIM$(Mid$(DTP$, 65, 16))
            CODPARAM$(4) = "PATE-TRA": DOCPARAM$(4) = TRIM$(Mid$(DTP$, 81, 16))
            CODPARAM$(5) = "COND-TRA": DOCPARAM$(5) = TRIM$(Mid$(DTP$, 97, 32))
            CODPARAM$(6) = "DOCUCOND": DOCPARAM$(6) = TRIM$(Mid$(DTP$, 129, 24))
            CODPARAM$(7) = "LOCA-TRA": DOCPARAM$(7) = TRIM$(Mid$(DTP$, 159, 32))
            CODPARAM$(8) = "NUBULTOS": DOCPARAM$(8) = TRIM$(Mid$(DTP$, 153, 6))
            ' IMPRESION DEL CODIGO DE BARRAS PARA PICKING DEL REMITO EN CONSIGNACION
                NUCL$ = Right("000000" & TRIM$(Str(NC%)), 6)
                NUREMI$ = Right("00000000" & TRIM$(Mid$(REMI$, 4)), 8) 'MODIFIC. P'LONGITUD FIJA
                CADENA$ = NUCL$ + NUREMI$
                BARRA$ = CODIBARRA$(CADENA$, 2)
            CODPARAM$(9) = "REF-MAT3": DOCPARAM$(9) = BARRA$
            CAPARDOC% = 9
            '
            CODTABU1$(1) = "COD-MAT"
            CODTABU1$(2) = "DES-MAT"
            CODTABU1$(3) = "CANTIDAD"
            CODTABU1$(4) = "CAKILOS"
            CODTABU1$(5) = "UNIMED"
            CODTABU1$(6) = "N-OCOMPR"
            CODTABU1$(7) = "NRO-LOTE"
            CODTABU1$(8) = "UBI-MPRI"
            CACOLTAB1% = 8
            '
            DESTIDOC% = NC1%
            CAITAB1% = 0
            '
            'ORDENO EL ARCHIVO RESPENTANDO EL ORDEN PRIMERO TRAZABLES
            Call TRACENTRAL("BOLREDET", "")
            Call FILESORT("BOLREDET", "", 1, 1)
            Call TRALOCAL("BOLREDET", "")
            
            For YX& = 1 To ULTREG&("!BOLREDET")
               YY$ = REGLEIDO$("!BOLREDET", YX&)
               CI% = CVI(Mid$(YY$, 83, 2))
               CAN# = CVD(Mid$(YY$, 89, 8))
               If CI% > 0 Then
                  If CI% <= NUMAS% Then
                     If CAN# > 0.005 Then
                        '
                        If ESTRAZABLE%(CI%) > 0 Then ESTRAZA% = 1 ' PARA DEJAR UN ESPACIO CUANDO CAMBIA DE TRAZABLE A NO TRAZABLE
                        If ESTRAZA% > 0 And ESTRAZABLE%(CI%) < 1 Then CAITAB1% = CAITAB1% + 1: ESTRAZA% = 0
                        '
                        UNID$ = TRIM$(Mid$(YY$, 85, 4))
                        If UNID$ = "" Then UNID$ = UNIMED$(CI%)
                        CANS$ = CSTRING$(MKD$(CAN#), 3, 9, 1, 2)
                        KILO$ = CSTRING$(MKD$(CAN# * PESUNI(CI%)), 3, 12, 3, 2)
                        NULOTE$ = TRIM$(Mid$(YY$, 69, 12))
                        NUOCO$ = CSTRING$(Mid$(YY$, 151, 4), 3, 7, 0, 2)
                        '
                        CAITAB1% = CAITAB1% + 1
                        TETABU1$(1, CAITAB1%) = CODEXT$(CI%)
                        TETABU1$(2, CAITAB1%) = DESCRIT$(CI%)
                        TETABU1$(3, CAITAB1%) = CANS$
                        TETABU1$(4, CAITAB1%) = KILO$
                        TETABU1$(5, CAITAB1%) = UNID$
                        TETABU1$(6, CAITAB1%) = NUOCO$
                        TETABU1$(7, CAITAB1%) = NULOTE$
                        TETABU1$(8, CAITAB1%) = UBISTO$(CI%)
                        '
                        Call CARLINSEP  'Carga linea de separacion entre codigo y codigo

                     End If
                  End If
               End If
            Next YX&
            '
            Call PRINTDOC("@" & FORIPRE$)   ' Remito en Consignacion
            '
          End If
          '
1199    End If
        '
        If InStr(UCase$(EMPREAC$), "MEDITEA") > 0 And TRIM$(NROCOMX$) <> "" Then
           Call FORCOMP07.GENERALISTAPICK(NROCOMX$, "PRINT") 'GENERA LISTA DE PICKING
        End If
        
9998    Call CIERRARCH("*.*")
        '
        X$ = "": Y$ = "": YY$ = ""
        EBOL$ = "": PEDY$ = "": DESCRI$ = ""
        '
        If LLAMACOMPRA% = 0 Then
           GoTo 1100
        End If
        '
9999    Call CIERRARCH("*.*")
        Call LIMPIATETAB
        '
End Sub
'
Sub ASIGNANLOTE(OK%)
        '
        OK% = 0 ' 0 NO TIENE CODIGOS TRAZABLES O REALIZO CORRECTAMTE LA ASIGNACION DE LOTE
        '-1 CANCELO
        TRAZCOMPO.List2.Clear
        
        'RECORRO EL BOLREDET Y SI HAY ARTICULOS TRAZABLES VA CARGANDO EL
        'CODIGO PARA QUE SE LE ASIGNEN LOS LOTES EN LA VENTANA DEL TRAZCOMPO
        For YX& = 1 To ULTREG&("!BOLREDET")
          YY$ = REGLEIDO$("!BOLREDET", YX&)
          CI% = CVI(Mid$(YY$, 83, 2))
          CAN# = CVD(Mid$(YY$, 89, 8))
          NUOC& = CVS(Mid$(YY$, 151, 4))
          NORCOMPRA$ = TRIM$(Mid$(YY$, 119, 16))
          If CI% > 0 Then
            If CI% <= NUMAS% Then
              If CAN# > FRACMIN(CI%) Then
                If ESTRAZABLE%(CI%) > 0 Then
                   TX$ = ""
                   Call REPLA(TX$, CODEXT$(CI%), 1, 16)
                   Call REPLA(TX$, DESCRIT$(CI%), 17, 32)
                   Call REPLA(TX$, FORMATNUM(CAN#, "F10.1"), 50, 11)
                   Call REPLA(TX$, Str$(NUOC&), 80, 10) ' guardo en list oculto n.ocompra
                   Call REPLA(TX$, NORCOMPRA$, 90, 16) ' guardo en list oculto n.ocompra texto
                   TRAZCOMPO.List2.AddItem TX$
                End If
             End If
           End If
          End If
        Next YX&
   
        'Presentar formulario para asignar lotes

        If TRAZCOMPO.List2.ListCount > 0 Then
            TRAZCOMPO.Show 1
            If TRAZCOMPO.List2.ListCount < 1 Then ' CASO QUE NO HAYA CODIGOS TRAZABLES
                 TRAZCOMPO.LIMPIAR
                 OK% = (-1)
            Else
                If TRAZCOMPO.Label7 = "FALTAN_LOTES" Then
                    OK% = (-1): Exit Sub
                End If
                '
                j& = 0
                For YX& = 1 To ULTREG&("!BOLREDET")
                  YY$ = REGLEIDO$("!BOLREDET", YX&)
                  CI% = CVI(Mid$(YY$, 83, 2))
                  CAN# = CVD(Mid$(YY$, 89, 8))
                  If CI% > 0 Then
                    If CI% <= NUMAS% Then
                      If ESTRAZABLE%(CI%) < 1 Then
                         j& = j& + 1
                         Call GRAREG("!BOLREDET", YY$, j&)
                      End If
                    End If
                  End If
                Next YX&
                
                For I& = 1 To TRAZCOMPO.List2.ListCount
                     'CODIGO DE LIST DE TRAZCOMPO
                      CI% = CODINT%(TRIM$(Left$(TRAZCOMPO.List2.LIST(I& - 1), 16)))
                      'GRABO EN EL BOLRDET EL NUMERO DE LOTE
                      LOTE$ = TRIM$(Mid$(TRAZCOMPO.List2.LIST(I& - 1), 62, 16))
                      CANTI# = XVALO(Mid$(TRAZCOMPO.List2.LIST(I& - 1), 50, 11))
                      NUOC& = XVALO(TRIM$(Mid$(TRAZCOMPO.List2.LIST(I& - 1), 80, 10)))
                      NORCOMPRA$ = TRIM$(Mid$(TRAZCOMPO.List2.LIST(I& - 1), 90, 16))
                      '
                      YY$ = REGBLAN("!BOLREDET")
                      Call REPLA(YY$, LOTE$, 69, 12)
                      Call REPLA(YY$, MKI$(CI%), 83, 2)
                      Call REPLA(YY$, UNIMED$(CI%), 85, 4)
                      Call REPLA(YY$, MKD$(CANTI#), 89, 8)
                      Call REPLA(YY$, NORCOMPRA$, 119, 16)
                      Call REPLA(YY$, MKS$(NUOC&), 151, 4)
                      Call REPLA(YY$, MKS$(0), 189, 4)
                      j& = j& + 1
                      Call GRAREG("!BOLREDET", YY$, j&)
                Next I&
                '
                Call SETULTREG("!BOLREDET", j&)
           End If
        End If
        TRAZCOMPO.List2.Clear
        
        

End Sub
'
Sub DESCACONSI(NCX%, CIX%, CANX#, NUOCO&, Optional FERECEP%, Optional NUBOL&, Optional IDENLO$, Optional MODO$)
    '
Call DESCACONSI_NUE(NCX%, CIX%, CANX#, NUOCO&, FERECEP%, NUBOL&, IDENLO$, MODO$) ' a partir del 2 de noviembre de 2011
Exit Sub
    '
    CAREN# = 0
    CAPEN# = CANX#
    '
    VUELTA% = 1
    sqlx$ = "SELECT * From MACONSIG where Npro_Ocom = " & NCX% & " "
    sqlx$ = sqlx$ + " AND Cant_Rendid < Cant_Entreg - 0.001 "
    sqlx$ = sqlx$ + " AND Cod_Inte=" & CIX%
    '
20  Dim MaConsig As ADODB.Recordset
    Set MaConsig = New ADODB.Recordset
25  On Error Resume Next
    MaConsig.Open FILTSQL$(sqlx$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
    If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
       On Error GoTo 0
       Call CapturaErrorOpen
       GoTo 25
    End If
    On Error GoTo 0
    '
    ' AQUÍ DEVOLUCION DIRECTA DEL MATERIAL RECIBIDO
    With MaConsig
      Do While Not (.EOF)
         NUOCA& = XVALO(Mid$(!Nro_Ocom, 4, 6))
         If NUOCA& = NUOCO& Or (NUOCA& = 0 And VUELTA% = 2) Then ' coincide o/c o bien consignacion sin o/c
            CACON# = XVALO(!cant_entreg) - XVALO(!cant_rendid)
               If CACON# < 0 Then CACON# = 0
            CAREN# = CAPEN#
               If CAREN# > CACON# Then CAREN# = CACON#
            CAPEN# = CAPEN# - CAREN#
'           !Cant_Rendid = XVALO(!Cant_Rendid) + CAREN#
'           !Saldo_Pend = XVALO(!Cant_Entreg) - XVALO(!Cant_Rendid)
            !cant_devu = XVALO(!cant_devu) + CAREN#
            !Saldo_Pend = XVALO(!cant_entreg) - (XVALO(!cant_rendid) + XVALO(!cant_devu))
            .Update
         End If
      .MoveNext
      Loop
    End With
    MaConsig.Close
    '
    If CAPEN# > 0.005 Then
       If VUELTA% = 1 Then
          VUELTA% = 2
          GoTo 20
       End If
    End If
    '
    ' AQUI RENDICION DE COMPONENTES DEL MATERIAL RECIBIDO
    If CAPEN# >= 0.05 Then
       CODDX$ = CODEXT$(CIX%)
       CAPENDE = CAPEN#
       '
       FE% = HOY(HO$)
       FECH = FETEXCOR1$(FE%)

       Call LEEEXPLO(CODDX$, CAPENDE)
       For j% = 1 To CAIT%
          CIY% = CIJ%(j%)
          CANY# = CAPEN# * CDbl(USOI(j%))
          '
          VUELTA% = 1
          sqlx$ = "SELECT * From MACONSIG where Npro_Ocom = " & NCX% & " "
          sqlx$ = sqlx$ + " AND Cant_Rendid < Cant_Entreg - 0.001 "
          sqlx$ = sqlx$ + " AND Cod_Inte=" & CIY%
18        Set MaConsig = New ADODB.Recordset
26        On Error Resume Next
          MaConsig.Open FILTSQL$(sqlx$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
          If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
             On Error GoTo 0
             Call CapturaErrorOpen
             GoTo 26
          End If
          On Error GoTo 0
          '
          With MaConsig
             Do While Not (.EOF)
                NUOCA& = XVALO(Mid$(!Nro_Ocom, 4, 6))
                If NUOCA& = NUOCO& Or (NUOCA& = 0 And VUELTA% = 2) Or (NUOCA& <> NUOCO& And VUELTA% = 3) Then ' coincide o/c o bien consignacion sin o/c
                   CACON# = XVALO(!cant_entreg) - XVALO(!cant_rendid) - XVALO(!cant_devu)
                      If CACON# < 0 Then CACON# = 0
                   CAREN# = CANY#
                     If CAREN# > CACON# Then CAREN# = CACON#
                     !cant_rendid = XVALO(!cant_rendid) + CAREN#
                     !Saldo_Pend = XVALO(!cant_entreg) - XVALO(!cant_rendid)

                   .Update
                   CANY# = CANY# - CAREN#
                End If
             .MoveNext
             Loop
          End With
          MaConsig.Close
          '
          If CANY# > 0.005 Then
             If VUELTA% = 1 Then
                VUELTA% = 2
                GoTo 18
             ElseIf VUELTA% = 2 Then
                VUELTA% = 3
                GoTo 18
             ElseIf VUELTA% = 3 Then
100             ABUS$ = RANDSTRING$(-10)
                CONDI$ = "NRO_OCOM = '" & ABUS$ & "'"
                If CantRegistros("MACONSIG", CONDI$) > 0 Then GoTo 100
                Call CreaRegistro("MACONSIG", "NRO_OCOM", ABUS$, "NOMESS")
                Call ACTUREGISTRO("MACONSIG", "NPRO_OCOM", CONDI$, NCX%, REGAF&, "NOMESS")
                Call ACTUREGISTRO("MACONSIG", "RASO_PROV", CONDI$, RASOCLI(-1 * NCX%), REGAF&, "NOMESS")
                Call ACTUREGISTRO("MACONSIG", "SECC_EJEC", CONDI$, 0, REGAF&, "NOMESS")
                Call ACTUREGISTRO("MACONSIG", "Femi_Ocom", CONDI$, FECH, REGAF&, "NOMESS")
                Call ACTUREGISTRO("MACONSIG", "Deno_Secc", CONDI$, "", REGAF&, "NOMESS")
                Call ACTUREGISTRO("MACONSIG", "Nro_Orfa", CONDI$, "", REGAF&, "NOMESS")
                Call ACTUREGISTRO("MACONSIG", "Cod_Inte", CONDI$, CIY%, REGAF&, "NOMESS")
                Call ACTUREGISTRO("MACONSIG", "Cod_Exte", CONDI$, CODEXT$(CIY%), REGAF&, "NOMESS")
                Call ACTUREGISTRO("MACONSIG", "Descritem", CONDI$, DESCRIT0$(CIY%), REGAF&, "NOMESS")
                Call ACTUREGISTRO("MACONSIG", "Lote_Ocom", CONDI$, "", REGAF&, "NOMESS")
                Call ACTUREGISTRO("MACONSIG", "Unidad", CONDI$, UNIMED(CIY%), REGAF&, "NOMESS")
                Call ACTUREGISTRO("MACONSIG", "Cant_Asig", CONDI$, 0, REGAF&, "NOMESS")
                Call ACTUREGISTRO("MACONSIG", "Cant_Entreg", CONDI$, 0, REGAF&, "NOMESS")
                Call ACTUREGISTRO("MACONSIG", "Fech_Entreg", CONDI$, FECH, REGAF&, "NOMESS")
                Call ACTUREGISTRO("MACONSIG", "NRemi_Consig", CONDI$, "", REGAF&, "NOMESS")
                Call ACTUREGISTRO("MACONSIG", "Valo_Unit", CONDI$, 0, REGAF&, "NOMESS")
                Call ACTUREGISTRO("MACONSIG", "Moneda", CONDI$, MONEMI%, REGAF&, "NOMESS")
                Call ACTUREGISTRO("MACONSIG", "Cant_Rendid", CONDI$, CANY#, REGAF&, "NOMESS")
                Call ACTUREGISTRO("MACONSIG", "Saldo_Pend", CONDI$, 0, REGAF&, "NOMESS")
                Call ACTUREGISTRO("MACONSIG", "Observa", CONDI$, "", REGAF&, "NOMESS")
                Call ACTUREGISTRO("MACONSIG", "Valo_Tot", CONDI$, 0, REGAF&, "NOMESS")
                Call ACTUREGISTRO("MACONSIG", "NRO_OCOM", CONDI$, "OC-000000-00", REGAF&, "NOMESS")
             End If
          End If
       Next j%
    End If
    '
    Set MaConsig = Nothing
    '
End Sub
'
Sub DESCACONSI_NUE(NCX%, CIX%, CANX#, NUOCO&, Optional FERECEP%, Optional NUBOL&, Optional IDENLO$, Optional MODO$)
    '
    ' CREA LA TABLA RENCONSIG - RENDIMIENTO DE CONSIGNACIONES
    '
    Call CREACAMPO("STOCKS", "RENCONSIG", "Fe_Recep", 8, 0)
    Call CREACAMPO("STOCKS", "RENCONSIG", "NumeBoRe", 6, 4, 1)
    Call CREACAMPO("STOCKS", "RENCONSIG", "Nro_Ocom", 10, 12, 1)
    Call CREACAMPO("STOCKS", "RENCONSIG", "Nu_Prove", 3, 2, 1)
    Call CREACAMPO("STOCKS", "RENCONSIG", "Codi_Conj", 3, 0, 1)
    Call CREACAMPO("STOCKS", "RENCONSIG", "Cant_Conj", 7, 8)
    Call CREACAMPO("STOCKS", "RENCONSIG", "Codi_Elem", 3, 0, 1)
    Call CREACAMPO("STOCKS", "RENCONSIG", "Cant_Rendid", 7, 8)
    Call CREACAMPO("STOCKS", "RENCONSIG", "TipoMovim", 10, 4, 1)
    Call CREACAMPO("STOCKS", "RENCONSIG", "MarProce", 3, 0, 1)
    '
    ' traslada a la tabla los movimientos acumulados de MACONSIG
    If CantRegistros("RENCONSIG") < 1 Then
       ' ABRE LA TABLA DE RENDICION DE CONSIGNACIONES
       sqlx$ = "SELECT * From RENCONSIG where Nu_Prove = " & NCX% & " "
       Dim RenConsig As ADODB.Recordset
       Set RenConsig = New ADODB.Recordset
       RenConsig.Open FILTSQL$(sqlx$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
       '
       '
       sqlx$ = "SELECT Npro_Ocom, Cod_Inte, Nro_Ocom, "
       sqlx$ = sqlx$ + " SUM(Cant_Rendid) AS SxCant_Rendid, SUM(Cant_Devu) AS SxCant_Devu "
       sqlx$ = sqlx$ + " FROM MACONSIG WITH(NOLOCK) "
       sqlx$ = sqlx$ + " WHERE Cant_Rendid > 0 OR Cant_Devu > 0 "
       sqlx$ = sqlx$ + " GROUP BY Npro_Ocom , Cod_Inte , Nro_Ocom "
       sqlx$ = sqlx$ + " ORDER BY Npro_Ocom ASC, Cod_Inte ASC, Nro_Ocom ASC"
       Dim RENCO As ADODB.Recordset
       Set RENCO = READSET(sqlx$, 999)
       '
       Screen.MousePointer = 11
       With RENCO
          Do While Not .EOF
             HOII% = HOY(HO$)
             NORCOM$ = SAFETEXT$(!Nro_Ocom)
             NPROL& = XVALO(!Npro_Ocom)
             CIELE% = XVALO(!Cod_Inte)
             CAREN = XVALO(!SxCant_Rendid)
             CADEV = XVALO(!SxCant_Devu)
             '
             If Abs(CAREN) > 0.005 Then
                 RenConsig.AddNew
                 RenConsig!Fe_Recep = CVDAT(HOII%)
                 RenConsig!NUMEBORE = 0
                 RenConsig!Nro_Ocom = NORCOM$
                 RenConsig!Nu_Prove = NPROL&
                 RenConsig!Codi_Conj = 0
                 RenConsig!Cant_Conj = 0
                 RenConsig!Codi_Elem = CIELE%
                 RenConsig!cant_rendid = CAREN
                 RenConsig!TipoMovim = "QR"
                 RenConsig!MarProce = 0
                 RenConsig.Update
             End If
             '
             If Abs(CADEV) > 0.005 Then
                 RenConsig.AddNew
                 RenConsig!Fe_Recep = CVDAT(HOII%)
                 RenConsig!NUMEBORE = 0
                 RenConsig!Nro_Ocom = NORCOM$
                 RenConsig!Nu_Prove = NPROL&
                 RenConsig!Codi_Conj = 0
                 RenConsig!Cant_Conj = 0
                 RenConsig!Codi_Elem = CIELE%
                 RenConsig!cant_rendid = CADEV
                 RenConsig!TipoMovim = "QD"
                 RenConsig!MarProce = 0
                 RenConsig.Update
             End If
             '
          .MoveNext
          Loop
       End With
       Screen.MousePointer = 1
       '
       RenConsig.Close: Set RenConsig = Nothing
       RENCO.Close: Set RENCO = Nothing
    End If
    '
    ' VERIFICA SI ES UN MATERIA PRIMA Y NO UN MATERIAL CONSIGNADO.
    ' EN CASO DE SER MP SE VA DE LA RUTINA
    CODDX$ = CODEXT$(CIX%)
    CAPENDE# = CANX#
    Call LEEEXPLO(CODDX$, CAPENDE)
    If CAIT% < 1 And MODO$ <> "DV" Then Exit Sub
    '
    ' ABRE LA TABLA DE RENDICION DE CONSIGNACIONES
    sqlx$ = "SELECT * From RENCONSIG where Nu_Prove = " & NCX% & " "
    Dim RenConsigNew As ADODB.Recordset
    Set RenConsigNew = New ADODB.Recordset
    RenConsigNew.Open FILTSQL$(sqlx$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
    '
    FE% = FERECEP%: If FE% < 1 Then FE% = HOY(HO$)
    FECH = CVDAT(FE%)
    NOCOM$ = TRIM$(Str$(NUOCO&))
    While Len(NOCOM$) < 6: NOCOM$ = "0" + NOCOM$: Wend
    NORCOM$ = "OC-" + Left$(NOCOM$, 6) + "%"
    '
    ' AQUI HAY QUE LLER LA ESTRUCTURA PRIMERO EN EL MACONSIG Y SOLO SI NO HAY REGISTROS EN EL
    ' MACONSIG ENTONCES EN LA FORMULA
    With RenConsigNew
        If CAIT% > 0 Then
              ' AQUI RENDICION DE COMPONENTES
              For j% = 1 To CAIT%
                 CIY% = CIJ%(j%)
                 CANY# = CAPENDE# * CDbl(USOI(j%))
                 '
                 .AddNew
                 !Fe_Recep = FECH
                 !NUMEBORE = NUBOL&
                 !Nro_Ocom = NORCOM$
                 !Nu_Prove = NCX%
                 !Codi_Conj = CIX%
                 !Cant_Conj = CAPENDE#
                 !Codi_Elem = CIY%
                 !cant_rendid = CANY#
                 !TipoMovim = "CC"
                 !MarProce = 0
                 .Update
              Next j%
        Else
              ' AQUÍ DEVOLUCION DIRECTA DEL MATERIAL RECIBIDO
              .AddNew
              !Fe_Recep = FECH
              !NUMEBORE = NUBOL&
              !Nro_Ocom = NORCOM$
              !Nu_Prove = NCX%
              !Codi_Conj = 0
              !Cant_Conj = 0
              !Codi_Elem = CIX%
              !cant_rendid = CAPENDE#
              !TipoMovim = "DV"
              !MarProce = 0
              .Update
        End If
    End With
    '
    RenConsigNew.Close
    Set RenConsigNew = Nothing
    '
End Sub
'
Sub AJUMACON_VIEJA()
    '
    SCMA = Screen.MousePointer
    Screen.MousePointer = 11
    '
    CONDI$ = "MARPROCE = 0"
    FIN& = CantRegistros("RENCONSIG", CONDI$)
    sqlx$ = "SELECT * FROM RENCONSIG WITH(NOLOCK) WHERE " + CONDI$
    Dim SERECO As ADODB.Recordset
    Set SERECO = READSET(sqlx$)
    '
    Screen.MousePointer = 11
    U& = 0
    With SERECO
       Do While Not .EOF
          U& = U& + 1
          Call TRACE(20, U&, FIN&)
          NORCOM$ = Left$(SAFETEXT$(!Nro_Ocom), 9) + "%"
          CIX% = XVALO(!Codi_Elem)
          NCX% = XVALO(!Nu_Prove)
          '
          SQLY$ = "SELECT SUM(CANT_RENDID) AS SUREN FROM RENCONSIG WITH(NOLOCK) "
          SQLY$ = SQLY$ + " WHERE (NRO_OCOM LIKE '" & NORCOM$ & "'"
             If Left$(NORCOM$, 9) = "OC-000000" Then
                 SQLY$ = SQLY$ + " OR NRO_OCOM = '' OR NRO_OCOM IS NULL)"
               Else
                 SQLY$ = SQLY$ + ")"
             End If
          SQLY$ = SQLY$ + " AND Codi_Elem = " & CIX% & " "
          SQLY$ = SQLY$ + " AND Nu_Prove = " & NCX%
          'sqly$ = sqly$ + " AND (TIPOMOVIM = 'CC' OR TIPOMOVIM = 'QR')"
          SQLY$ = SQLY$ + " AND (TIPOMOVIM = 'DV' OR TIPOMOVIM = 'CC' OR TIPOMOVIM = 'QR')"
          Set QRENDI = READSET(SQLY$)
          SURENDI = XVALO(QRENDI!SUREN)
          CONDI$ = "(Nro_Ocom LIKE '" & NORCOM$ & "'"
             If Left$(NORCOM$, 9) = "OC-000000" Then
                 CONDI$ = CONDI$ + " OR NRO_OCOM = '' OR NRO_OCOM IS NULL)"
               Else
                 CONDI$ = CONDI$ + ")"
             End If
             CONDI$ = CONDI$ + " AND Cod_Inte = " & CIX% & " AND Npro_Ocom = " & NCX%
             '
          Call ACTUREGISTRO("MACONSIG", "Cant_Rendid", CONDI$, 0, RAFE&)
          Call ACTUREGISTRO("MACONSIG", "Cant_Devu", CONDI$, 0, RAFE&)
          Call ACTUREGISTRO("MACONSIG", "Saldo_Pend", CONDI$, 0, RAFE&)
          '
          sqlx$ = "SELECT * FROM MACONSIG WHERE " + CONDI$
          Dim CONSINGTEMP As ADODB.Recordset
          Set CONSINGTEMP = New ADODB.Recordset
25        On Error Resume Next
          CONSINGTEMP.Open sqlx$, FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
          If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
             On Error GoTo 0
             Call CapturaErrorOpen
             GoTo 25
          End If
          On Error GoTo 0
          '
          With CONSINGTEMP
             Do While Not .EOF
                SALPENDEVU = XVALO(!cant_entreg) - (XVALO(!cant_rendid) + XVALO(!cant_devu))
                If SALPENDEVU > 0 Then
                   DEVU = SALPENDEVU
                   If DEVU > SURENDI Then DEVU = SURENDI
                   !cant_rendid = XVALO(!cant_rendid) + DEVU
                   .Update
                   SURENDI = SURENDI - DEVU
                   If SURENDI < FRACMIN(CIX%) Then GoTo 39
                End If
             .MoveNext
             Loop
          End With
          '
          If SURENDI > 0 Then
             HOII% = HOY(HO$)
             FECH = CVDAT(HOII%)
100          ABUS$ = RANDSTRING$(-10)
             CONDI$ = "NRO_OCOM = '" & ABUS$ & "'"
             If CantRegistros("MACONSIG", CONDI$) > 0 Then GoTo 100
             Call CreaRegistro("MACONSIG", "NRO_OCOM", ABUS$, "NOMESS")
             Call ACTUREGISTRO("MACONSIG", "NPRO_OCOM", CONDI$, NCX%, REGAF&, "NOMESS")
             Call ACTUREGISTRO("MACONSIG", "RASO_PROV", CONDI$, RASOCLI(-1 * NCX%), REGAF&, "NOMESS")
             Call ACTUREGISTRO("MACONSIG", "SECC_EJEC", CONDI$, 0, REGAF&, "NOMESS")
             Call ACTUREGISTRO("MACONSIG", "Femi_Ocom", CONDI$, FECH, REGAF&, "NOMESS")
             Call ACTUREGISTRO("MACONSIG", "Deno_Secc", CONDI$, "", REGAF&, "NOMESS")
             Call ACTUREGISTRO("MACONSIG", "Nro_Orfa", CONDI$, "", REGAF&, "NOMESS")
             Call ACTUREGISTRO("MACONSIG", "Cod_Inte", CONDI$, CIX%, REGAF&, "NOMESS")
             Call ACTUREGISTRO("MACONSIG", "Cod_Exte", CONDI$, CODEXT$(CIX%), REGAF&, "NOMESS")
             Call ACTUREGISTRO("MACONSIG", "Descritem", CONDI$, DESCRIT0$(CIX%), REGAF&, "NOMESS")
             Call ACTUREGISTRO("MACONSIG", "Lote_Ocom", CONDI$, "", REGAF&, "NOMESS")
             Call ACTUREGISTRO("MACONSIG", "Unidad", CONDI$, UNIMED(CIX%), REGAF&, "NOMESS")
             Call ACTUREGISTRO("MACONSIG", "Cant_Asig", CONDI$, 0, REGAF&, "NOMESS")
             Call ACTUREGISTRO("MACONSIG", "Cant_Entreg", CONDI$, 0, REGAF&, "NOMESS")
             Call ACTUREGISTRO("MACONSIG", "Fech_Entreg", CONDI$, FECH, REGAF&, "NOMESS")
             Call ACTUREGISTRO("MACONSIG", "NRemi_Consig", CONDI$, "", REGAF&, "NOMESS")
             Call ACTUREGISTRO("MACONSIG", "Valo_Unit", CONDI$, 0, REGAF&, "NOMESS")
             Call ACTUREGISTRO("MACONSIG", "Moneda", CONDI$, MONEMI%, REGAF&, "NOMESS")
             Call ACTUREGISTRO("MACONSIG", "Cant_Rendid", CONDI$, SURENDI, REGAF&, "NOMESS")
             Call ACTUREGISTRO("MACONSIG", "Saldo_Pend", CONDI$, 0, REGAF&, "NOMESS")
             Call ACTUREGISTRO("MACONSIG", "Observa", CONDI$, "", REGAF&, "NOMESS")
             Call ACTUREGISTRO("MACONSIG", "Valo_Tot", CONDI$, 0, REGAF&, "NOMESS")
             Call ACTUREGISTRO("MACONSIG", "NRO_OCOM", CONDI$, Left$(NORCOM$, Len(NORCOM$) - 1), REGAF&, "NOMESS")
          End If
          '
39        CONSINGTEMP.Close
          Set CONSINGTEMP = Nothing
          '
          SQLY$ = "SELECT SUM(CANT_RENDID) AS SUREN FROM RENCONSIG WITH(NOLOCK) "
          SQLY$ = SQLY$ + " WHERE NRO_OCOM LIKE '" & NORCOM$ & "' "
          SQLY$ = SQLY$ + " AND Codi_Elem = " & CIX% & " "
          SQLY$ = SQLY$ + " AND Nu_Prove = " & NCX%
          'sqly$ = sqly$ + " AND (TIPOMOVIM = 'DV' OR TIPOMOVIM = 'QD' OR TIPOMOVIM = 'AJ')"
          SQLY$ = SQLY$ + " AND (TIPOMOVIM = 'QD' OR TIPOMOVIM = 'AJ')"
          Set QRENDI = READSET(SQLY$)
          SURENDI = XVALO(QRENDI!SUREN)
          'PARA CASOS DE AJUSTES QUE AUMENTAN EL STOCK EN EL ARMADOR SE CREA UN NUEVO REGISTRO
          'COMO SI FUERA UNA ENTREGA DE MATERIAL EXTRA
          If SURENDI < 0 Then GoTo 40
          '
          CONDI$ = "(Nro_Ocom LIKE '" & NORCOM$ & "'"
             If Left$(NORCOM$, 9) = "OC-000000" Then
                 CONDI$ = CONDI$ + " OR NRO_OCOM = '' OR NRO_OCOM IS NULL)"
               ElseIf Left$(NORCOM$, 3) = "AR-" Then
                 CONDI$ = "CANT_ENTREG > 0 "
               Else
                 CONDI$ = CONDI$ + ")"
             End If
             CONDI$ = CONDI$ + " AND Cod_Inte = " & CIX% & " AND Npro_Ocom = " & NCX%
             '
          sqlx$ = "SELECT * FROM MACONSIG WHERE " + CONDI$
          Set CONSINGTEMP = New ADODB.Recordset
35        On Error Resume Next
          CONSINGTEMP.Open sqlx$, FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
          If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
             On Error GoTo 0
             Call CapturaErrorOpen
             GoTo 35
          End If
          On Error GoTo 0
          '
          With CONSINGTEMP
             Do While Not .EOF
                SALPENDEVU = XVALO(!cant_entreg) - (XVALO(!cant_rendid) + XVALO(!cant_devu))
                If SALPENDEVU > 0.05 Then
                   DEVU = SALPENDEVU
                   If DEVU > SURENDI Then DEVU = SURENDI
                   'If XVALO(!CANT_DEVU) + DEVU < 0 Then DEVU = (-1) * XVALO(!CANT_DEVU)
                   !cant_devu = XVALO(!cant_devu) + DEVU
                   .Update
                   SURENDI = SURENDI - DEVU
                   If Abs(SURENDI) < FRACMIN(CIX%) Then GoTo 49
                End If
             .MoveNext
             Loop
          End With
          '
40        If Abs(SURENDI) > 0 Then
             HOII% = HOY(HO$)
             FECH = CVDAT(HOII%)
101          ABUS$ = RANDSTRING$(-10)
             CONDI$ = "NRO_OCOM = '" & ABUS$ & "'"
             If CantRegistros("MACONSIG", CONDI$) > 0 Then GoTo 101
             Call CreaRegistro("MACONSIG", "NRO_OCOM", ABUS$, "NOMESS")
             Call ACTUREGISTRO("MACONSIG", "NPRO_OCOM", CONDI$, NCX%, REGAF&, "NOMESS")
             Call ACTUREGISTRO("MACONSIG", "RASO_PROV", CONDI$, RASOCLI(-1 * NCX%), REGAF&, "NOMESS")
             Call ACTUREGISTRO("MACONSIG", "SECC_EJEC", CONDI$, 0, REGAF&, "NOMESS")
             Call ACTUREGISTRO("MACONSIG", "Femi_Ocom", CONDI$, FECH, REGAF&, "NOMESS")
             Call ACTUREGISTRO("MACONSIG", "Deno_Secc", CONDI$, "", REGAF&, "NOMESS")
             Call ACTUREGISTRO("MACONSIG", "Nro_Orfa", CONDI$, "", REGAF&, "NOMESS")
             Call ACTUREGISTRO("MACONSIG", "Cod_Inte", CONDI$, CIX%, REGAF&, "NOMESS")
             Call ACTUREGISTRO("MACONSIG", "Cod_Exte", CONDI$, CODEXT$(CIX%), REGAF&, "NOMESS")
             Call ACTUREGISTRO("MACONSIG", "Descritem", CONDI$, DESCRIT0$(CIX%), REGAF&, "NOMESS")
             Call ACTUREGISTRO("MACONSIG", "Lote_Ocom", CONDI$, "", REGAF&, "NOMESS")
             Call ACTUREGISTRO("MACONSIG", "Unidad", CONDI$, UNIMED(CIX%), REGAF&, "NOMESS")
             Call ACTUREGISTRO("MACONSIG", "Cant_Asig", CONDI$, 0, REGAF&, "NOMESS")
                CANENT = 0: If SURENDI < 0 Then CANENT = Abs(SURENDI)
             Call ACTUREGISTRO("MACONSIG", "Cant_Entreg", CONDI$, CANENT, REGAF&, "NOMESS")
             Call ACTUREGISTRO("MACONSIG", "Fech_Entreg", CONDI$, FECH, REGAF&, "NOMESS")
             Call ACTUREGISTRO("MACONSIG", "NRemi_Consig", CONDI$, "", REGAF&, "NOMESS")
             Call ACTUREGISTRO("MACONSIG", "Valo_Unit", CONDI$, 0, REGAF&, "NOMESS")
             Call ACTUREGISTRO("MACONSIG", "Moneda", CONDI$, MONEMI%, REGAF&, "NOMESS")
             Call ACTUREGISTRO("MACONSIG", "Cant_Rendid", CONDI$, 0, REGAF&, "NOMESS")
                CANDEV = SURENDI: If SURENDI < 0 Then CANDEV = 0
             Call ACTUREGISTRO("MACONSIG", "Cant_Devu", CONDI$, CANDEV, REGAF&, "NOMESS")
             Call ACTUREGISTRO("MACONSIG", "Saldo_Pend", CONDI$, 0, REGAF&, "NOMESS")
             Call ACTUREGISTRO("MACONSIG", "Observa", CONDI$, "", REGAF&, "NOMESS")
             Call ACTUREGISTRO("MACONSIG", "Valo_Tot", CONDI$, 0, REGAF&, "NOMESS")
                If SURENDI < 0 Then NORCOM$ = "OC-000000%"
             Call ACTUREGISTRO("MACONSIG", "NRO_OCOM", CONDI$, Left$(NORCOM$, Len(NORCOM$) - 1), REGAF&, "NOMESS")
          End If
          '
          On Error Resume Next
49        CONSINGTEMP.Close
          Set CONSINGTEMP = Nothing
          On Error GoTo 0
'          If SURENDI > 0 Then
'          QTR& = CantRegistros("MACONSIG", CONDI$)
'          If QTR& > 0 Then
'             Call ACTUREGISTRO("MACONSIG", "Cant_Devu", CONDI$, SURENDI / QTR&, RAFE&)
'          End If
'          '
       .MoveNext
       Loop
    End With
    '
    SERECO.Close
    Set SERECO = Nothing
    '
    If U& > 0 Then
        CONDI$ = "MARPROCE = 0"
        Call ACTUREGISTRO("RENCONSIG", "MARPROCE", CONDI$, 1, RAFE&)
    End If
    '
    Screen.MousePointer = SCMA
    '
End Sub

Sub AJUMACON()
    '
    SCMA = Screen.MousePointer
    Screen.MousePointer = 11
    '
    CONDI$ = "MARPROCE = 0"
    FIN& = CantRegistros("RENCONSIG", CONDI$)
    sqlx$ = "SELECT * FROM RENCONSIG WITH(NOLOCK) WHERE " + CONDI$
    Dim SERECO As ADODB.Recordset
    Set SERECO = READSET(sqlx$)
    '
    Screen.MousePointer = 11
    U& = 0
    With SERECO
       Do While Not .EOF
          U& = U& + 1
          Call TRACE(20, U&, FIN&)
          NORCOM$ = Left$(SAFETEXT$(!Nro_Ocom), 9) + "%"
          CIX% = XVALO(!Codi_Elem)
          NCX% = XVALO(!Nu_Prove)
          '
          SQLY$ = "SELECT SUM(CANT_RENDID) AS SUREN FROM RENCONSIG WITH(NOLOCK) "
          SQLY$ = SQLY$ + " WHERE (NRO_OCOM LIKE '" & NORCOM$ & "'"
             If Left$(NORCOM$, 9) = "OC-000000" Then
                 SQLY$ = SQLY$ + " OR NRO_OCOM = '' OR NRO_OCOM IS NULL)"
               Else
                 SQLY$ = SQLY$ + ")"
             End If
          SQLY$ = SQLY$ + " AND Codi_Elem = " & CIX% & " "
          SQLY$ = SQLY$ + " AND Nu_Prove = " & NCX%
          'sqly$ = sqly$ + " AND (TIPOMOVIM = 'CC' OR TIPOMOVIM = 'QR')"
          SQLY$ = SQLY$ + " AND (TIPOMOVIM = 'DV' OR TIPOMOVIM = 'CC' OR TIPOMOVIM = 'QR')"
          Set QRENDI = READSET(SQLY$)
          SURENDI = XVALO(QRENDI!SUREN)
          CONDI$ = "(Nro_Ocom LIKE '" & NORCOM$ & "'"
             If Left$(NORCOM$, 9) = "OC-000000" Then
                 CONDI$ = CONDI$ + " OR NRO_OCOM = '' OR NRO_OCOM IS NULL)"
               Else
                 CONDI$ = CONDI$ + ")"
             End If
             CONDI$ = CONDI$ + " AND Cod_Inte = " & CIX% & " AND Npro_Ocom = " & NCX%
             '
          sqlx$ = "SELECT * FROM MACONSIG WHERE " + CONDI$
          Dim RSCONS As ADODB.Recordset
          Set RSCONS = New ADODB.Recordset
25        On Error Resume Next
          RSCONS.Open sqlx$, FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
          If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
             On Error GoTo 0
             Call CapturaErrorOpen
             GoTo 25
          End If
          On Error GoTo 0
          '
          With RSCONS
             Do While Not .EOF
               !cant_rendid = 0
               !cant_devu = 0
               !Saldo_Pend = 0
               .Update
               .MoveNext
             Loop
          End With
          RSCONS.Close
          Set RSCONS = Nothing
          
'          Call ACTUREGISTRO("MACONSIG", "Cant_Rendid", CONDI$, 0, RAFE&)
'          Call ACTUREGISTRO("MACONSIG", "Cant_Devu", CONDI$, 0, RAFE&)
'          Call ACTUREGISTRO("MACONSIG", "Saldo_Pend", CONDI$, 0, RAFE&)
          '
          sqlx$ = "SELECT * FROM MACONSIG WHERE " + CONDI$
          Dim CONSINGTEMP As ADODB.Recordset
          Set CONSINGTEMP = New ADODB.Recordset
26        On Error Resume Next
          CONSINGTEMP.Open sqlx$, FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
          If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
             On Error GoTo 0
             Call CapturaErrorOpen
             GoTo 26
          End If
          On Error GoTo 0
          '
          With CONSINGTEMP
             Do While Not .EOF
                SALPENDEVU = XVALO(!cant_entreg) - (XVALO(!cant_rendid) + XVALO(!cant_devu))
                If SALPENDEVU > 0 Then
                   DEVU = SALPENDEVU
                   If DEVU > SURENDI Then DEVU = SURENDI
                   !cant_rendid = XVALO(!cant_rendid) + DEVU
                   .Update
                   SURENDI = SURENDI - DEVU
                   If SURENDI < FRACMIN(CIX%) Then GoTo 39
                End If
             .MoveNext
             Loop
          End With
          '
          On Error Resume Next
          CONSINGTEMP.Close
          Set CONSINGTEMP = Nothing
          On Error GoTo 0
          '
          If SURENDI > 0 Then
             HOII% = HOY(HO$)
             FECH = CVDAT(HOII%)
'             If CantRegistros("MACONSIG", CONDI$) > 0 Then GoTo 100

             sqlx$ = "SELECT * FROM MACONSIG WHERE NRO_OCOM = ''"
             Dim RSCON As ADODB.Recordset
             Set RSCON = New ADODB.Recordset
27              On Error Resume Next
             RSCON.Open sqlx$, FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
             If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
               On Error GoTo 0
               Call CapturaErrorOpen
               GoTo 27
             End If
             On Error GoTo 0
            '
             With RSCON
               .AddNew
               !Nro_Ocom = Left$(NORCOM$, Len(NORCOM$) - 1)
               !Npro_Ocom = NCX%
               !Raso_Prov = Left$(RASOCLI(-1 * NCX%), 64)
               !Secc_Ejec = 0
               !Femi_Ocom = FECH
               !Deno_Secc = ""
               !NRO_ORFA = ""
               !Cod_Inte = CIX%
               !Cod_Exte = CODEXT$(CIX%)
               !Descritem = DESCRIT0$(CIX%)
               !LOTE_OCOM = ""
               !Unidad = UNIMED(CIX%)
               !cant_asig = 0
               !cant_entreg = 0
               !Femi_Ocom = FECH
               !FECH_ENTREG = FECH
               !NREMI_CONSIG = ""
               !Valo_Unit = 0
               !MONEDA = MONEMI%
               !cant_rendid = SURENDI
               !Saldo_Pend = 0
               !Observa = ""
               !Valo_Tot = 0
               .Update
             End With
             RSCON.Close
             Set RSCON = Nothing

          End If
          '
          '
39        On Error Resume Next
          CONSINGTEMP.Close
          Set CONSINGTEMP = Nothing
          On Error GoTo 0

          SQLY$ = "SELECT SUM(CANT_RENDID) AS SUREN FROM RENCONSIG WITH(NOLOCK) "
          SQLY$ = SQLY$ + " WHERE NRO_OCOM LIKE '" & NORCOM$ & "' "
          SQLY$ = SQLY$ + " AND Codi_Elem = " & CIX% & " "
          SQLY$ = SQLY$ + " AND Nu_Prove = " & NCX%
          'sqly$ = sqly$ + " AND (TIPOMOVIM = 'DV' OR TIPOMOVIM = 'QD' OR TIPOMOVIM = 'AJ')"
          SQLY$ = SQLY$ + " AND (TIPOMOVIM = 'QD' OR TIPOMOVIM = 'AJ')"
          Set QRENDI = READSET(SQLY$)
          SURENDI = XVALO(QRENDI!SUREN)
          'PARA CASOS DE AJUSTES QUE AUMENTAN EL STOCK EN EL ARMADOR SE CREA UN NUEVO REGISTRO
          'COMO SI FUERA UNA ENTREGA DE MATERIAL EXTRA
          If SURENDI < 0 Then GoTo 40
          '
          CONDI$ = "(Nro_Ocom LIKE '" & NORCOM$ & "'"
             If Left$(NORCOM$, 9) = "OC-000000" Then
                 CONDI$ = CONDI$ + " OR NRO_OCOM = '' OR NRO_OCOM IS NULL)"
               ElseIf Left$(NORCOM$, 3) = "AR-" Then
                 CONDI$ = "CANT_ENTREG > 0 "
               Else
                 CONDI$ = CONDI$ + ")"
             End If
             CONDI$ = CONDI$ + " AND Cod_Inte = " & CIX% & " AND Npro_Ocom = " & NCX%
             '
          sqlx$ = "SELECT * FROM MACONSIG WHERE " + CONDI$
          Dim CONSINGTEMP1 As ADODB.Recordset
          Set CONSINGTEMP1 = New ADODB.Recordset
35        On Error Resume Next
          CONSINGTEMP1.Open sqlx$, FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
          If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
             On Error GoTo 0
             Call CapturaErrorOpen
             GoTo 35
          End If
          On Error GoTo 0
          '
          With CONSINGTEMP1
             Do While Not .EOF
                SALPENDEVU = XVALO(!cant_entreg) - (XVALO(!cant_rendid) + XVALO(!cant_devu))
                If SALPENDEVU > 0.05 Then
                   DEVU = SALPENDEVU
                   If DEVU > SURENDI Then DEVU = SURENDI
                   'If XVALO(!CANT_DEVU) + DEVU < 0 Then DEVU = (-1) * XVALO(!CANT_DEVU)
                   !cant_devu = XVALO(!cant_devu) + DEVU
                   .Update
                   SURENDI = SURENDI - DEVU
                   If Abs(SURENDI) < FRACMIN(CIX%) Then GoTo 49
                End If
             .MoveNext
             Loop
          End With
          '
          On Error Resume Next
          CONSINGTEMP1.Close
          Set CONSINGTEMP1 = Nothing
          On Error GoTo 0
          '
40        If Abs(SURENDI) > 0 Then
             sqlx$ = "SELECT * FROM MACONSIG WHERE NRO_OCOM = ''"
             Dim RSCON2 As ADODB.Recordset
             Set RSCON2 = New ADODB.Recordset
28              On Error Resume Next
             RSCON2.Open sqlx$, FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
             If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
               On Error GoTo 0
               Call CapturaErrorOpen
               GoTo 28
             End If
             On Error GoTo 0
            '
             With RSCON2

               HOII% = HOY(HO$)
               FECH = CVDAT(HOII%)
'101          ABUS$ = RANDSTRING$(-10)
'             CONDI$ = "NRO_OCOM = '" & ABUS$ & "'"
'             If CantRegistros("MACONSIG", CONDI$) > 0 Then GoTo 101
'             Call CreaRegistro("MACONSIG", "NRO_OCOM", ABUS$, "NOMESS")
               .AddNew
                  If SURENDI < 0 Then NORCOM$ = "OC-000000%"
               !Nro_Ocom = Left$(NORCOM$, Len(NORCOM$) - 1)
               !Npro_Ocom = NCX%
               !Raso_Prov = RASOCLI(-1 * NCX%)
               !Secc_Ejec = 0
               !Femi_Ocom = FECH
               !Deno_Secc = ""
               !NRO_ORFA = ""
               !Cod_Inte = CIX%
               !Cod_Exte = CODEXT$(CIX%)
               !Descritem = DESCRIT0$(CIX%)
               !LOTE_OCOM = ""
               !Unidad = UNIMED(CIX%)
               !cant_asig = 0
                  CANENT = 0: If SURENDI < 0 Then CANENT = Abs(SURENDI)
               !cant_entreg = CANENT
               !Femi_Ocom = FECH
               !FECH_ENTREG = FECH
               !NREMI_CONSIG = ""
               !MONEDA = MONEMI%
               !cant_rendid = 0
               !Saldo_Pend = 0
               !Valo_Tot = 0
                  CANDEV = SURENDI: If SURENDI < 0 Then CANDEV = 0
               !cant_devu = CANDEV
               !Observa = ""
               .Update
             End With
             RSCON2.Close
             Set RSCON2 = Nothing
          End If
          '
49        On Error Resume Next
          CONSINGTEMP1.Close
          Set CONSINGTEMP1 = Nothing
          On Error GoTo 0

          .MoveNext
       Loop
    End With
    '
    SERECO.Close
    Set SERECO = Nothing
    '
    If U& > 0 Then
        CONDI$ = "MARPROCE = 0"
        Call ACTUREGISTRO("RENCONSIG", "MARPROCE", CONDI$, 1, RAFE&)
    End If
    '
    Screen.MousePointer = SCMA
    '
End Sub

'\\\OT-05-0419-WAR-26/06/05///
Sub NUREMIFA(NROPRO&, Optional NPVX$) 'NUMERO DE PUNTO DE VENTA
    '
    Static CMO$, FLIM%, YARECU%
    Screen.MousePointer = 11
    '
    If CMO$ = "" Then
      CMO$ = TRIM$(Control$("", "CODREFAC"))
      If CMO$ = "" Then CMO$ = "RT"
    End If
    '
    If FLIM% < 1 Then
      FENUE$ = Control$("", "FENUREFA")
      If FENUE$ = "" Then FENUE$ = "18/04/03"
      FLIM% = FECHANUM(FENUE$)
      If FLIM% > HOY(HO$) Then FLIM% = HOY(HO$)
    End If
    FLINUME = CDbl(CVDAT(FLIM%))
    '
    FEXUL = HOY(HOS$)
    FLIDIAS = CDbl(CVDAT(DIASANTES(FEXUL, 15)))
    '
    FLIMI = FLIDIAS: MODORAPIDO% = 1
    If FLINUME > FLIMI Then
       FLIMI = FLINUME
       MODORAPIDO% = 0
    End If
    '
10  CAMOLE% = 0
    NROMA& = 0
    RETOLE& = 0
    EPAC$ = TRIM$(UCase$(EMPREAC$))
    '
    FLID = FETEXCOR1$(FLIM%)
    Call APERBASDAT("STOCKS")
    '
    ' ESTO ES PARA RECUPERAR EL ERROR DE GUIDI - INCOHERENCIA (5/03/2009)
    If YARECU% > 0 Then GoTo 20
    If InStr(EPAC$, "GUIDI") < 1 Then GoTo 20
    '
    sqlx$ = "UPDATE MOVISTO SET CANTINGRE = 0 WHERE CANTINGRE IS NULL "
    FLEXCONN.Execute sqlx$
    sqlx$ = "UPDATE MOVISTO SET CANTSALID = 0 WHERE CANTSALID IS NULL "
    FLEXCONN.Execute sqlx$
    QTR& = CantRegistros("MOVISTO", "CANTSALDO IS NULL")
    If QTR& > 0 Then
        sqlx$ = "UPDATE MOVISTO SET CantSaldo = 0 WHERE CantSaldo IS NULL "
        FLEXCONN.Execute sqlx$
    End If
    QTR& = CantRegistros("MOVISTO", "VALOUNIT IS NULL")
    If QTR& > 0 Then
        sqlx$ = "UPDATE MOVISTO SET VALOUNIT = 0 WHERE VALOUNIT IS NULL "
        FLEXCONN.Execute sqlx$
    End If
    QTR& = CantRegistros("MOVISTO", "COD_INTE IS NULL")
    If QTR& > 0 Then
        sqlx$ = "UPDATE MOVISTO SET COD_INTE = 0 WHERE COD_INTE IS NULL "
        FLEXCONN.Execute sqlx$
    End If
    QTR& = CantRegistros("MOVISTO", "NUMUSUAR IS NULL")
    If QTR& > 0 Then
        sqlx$ = "UPDATE MOVISTO SET NUMUSUAR = 0 WHERE NUMUSUAR IS NULL "
        FLEXCONN.Execute sqlx$
    End If
    '
    ' ESTO ES PARA RECUPERAR EL ERROR DE GUIDI - ERROR 94 (6/4/2009)
    sqlx$ = "UPDATE PEDDETA SET TIPOPED= ' ' WHERE TIPOPED IS NULL "
    FLEXCONN.Execute sqlx$
    sqlx$ = "UPDATE PEDDETA SET CantEnt = 0 WHERE CantEnt IS NULL "
    FLEXCONN.Execute sqlx$
    sqlx$ = "UPDATE PEDDETA SET CanPed = 0 WHERE CanPed IS NULL "
    FLEXCONN.Execute sqlx$
    sqlx$ = "UPDATE PEDDETA SET NroPed = 0 WHERE NroPed IS NULL "
    FLEXCONN.Execute sqlx$
    YARECU% = 1
    '
    'SE CAMBIA PARA QUE SE REALICE LA BUSQUEDA DEL VALOR MAXIMO EN UN CAMPO NUMERICO.
20  NUPU = XVALO(NPVX$)
    CONDI$ = "FechMov >= '" & FLID & "'"
    CONDI$ = CONDI$ + " AND CodiMovi = '" & CMO$ & "' "
    CONDI$ = CONDI$ + " AND NUMPVTA = " & NUPU
    If InStr(1, EMPREAC$, "GABAL") > 0 Then
        CONDI$ = CONDI$ & " AND Codiempr =" & CodiEmp%
    End If
    REMITONUM& = XVALO(VALOBADA("MOVISTO", "MAX(DOPRINUM)", CONDI$, ""))
    If REMITONUM& > 0 Then NROMA& = REMITONUM&: GoTo 84
    '
    'MODO TRADICIONAL
    PARTIZ$ = TRIM$(NPVX$): While Len(PARTIZ$) < 4: PARTIZ$ = "0" + PARTIZ$: Wend
    PARTIZ$ = "RT-X-" + Left$(PARTIZ$, 4)
    'STRSQL = "select max(cast(substring(DoPriLar,11,8)as int)) as maximo from movisto "
    STRSQL$ = "select max(DoPriLar) as maximo from movisto with(nolock) "
    STRSQL$ = STRSQL$ & "WHERE FechMov >= '" & FLID & "'"
    STRSQL$ = STRSQL$ & " AND CodiMovi = '" & CMO$ & "' "
    STRSQL$ = STRSQL$ & " AND DOPRILAR LIKE '" + PARTIZ$ + "%' "
    If InStr(1, EMPREAC$, "GABAL") > 0 Then
        STRSQL$ = STRSQL$ & " AND Codiempr =" & CodiEmp%
    End If
    '
    Dim MOVITMP As ADODB.Recordset
    Set MOVITMP = New ADODB.Recordset
    MOVITMP.CursorLocation = adUseServer
'25  On Error Resume Next
'    MOVITMP.Open FILTSQL$(STRSQL), FLEXCONN, adOpenStatic, adLockReadOnly, adCmdText
'    If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
'      On Error GoTo 0
'      Call CapturaErrorOpen
'      GoTo 25
'    End If
'    On Error GoTo 0
    'Set MOVITMP = FLEXCONN.Execute(FILTSQL$(STRSQL))
    Set MOVITMP = READSET(STRSQL$)
    '
    On Error Resume Next
    If Err Then
       On Error GoTo 0
       Call MENSERR(24, "Imposible Determinar el Próximo Número.\Debe Introducirlo Manualmente.")
       NROMA& = 0
       GoTo 84
    End If
    On Error GoTo 0
    With MOVITMP
      On Error Resume Next
      .MoveFirst
      If Err Then
        If MODORAPIDO% = 1 Then
           On Error GoTo 0
           MODORAPIDO% = 0
           FLIMI = FLINUME
           GoTo 10
        End If
        '
        Call MENSERR(24, "Imposible Determinar Número de Próximo Remito")
        On Error GoTo 0
        GoTo 85
      End If
      On Error GoTo 0
      '
      On Error Resume Next
      NROMA& = XVALO(Mid$(!maximo, 11))
      On Error GoTo 0
    End With
    MOVITMP.Close
    Set MOVITMP = Nothing
    '
84  If NROMA& < 0 Then NROMA& = 0
85  If NROPRO& < 1 Then NROPRO& = 1 + NROMA&
    Screen.MousePointer = 1
End Sub
'\\\OT-05-0419-WAR-FIN///

'
Sub DATADI(NCIY%, CIIY%, ORCO$, NORI$, NPLA$)
    '
    ORCO$ = "": NORI$ = "": NPLA$ = ""
    '
    For U& = 1 To ULTREG&("DATADI")
        X$ = REGLEIDO$("DATADI", U&)
        If CVI(Left$(X$, 2)) = CIIY% Then
            If CVI(Mid$(X$, 3, 2)) = NCIY% Then
                ORCO$ = TRIM$(Mid$(X$, 5, 12))
                NORI$ = TRIM$(Mid$(X$, 17, 16))
                NPLA$ = TRIM$(Mid$(X$, 33, 16))
                Exit Sub
            End If
        End If
    Next U&
    '
End Sub
'

Function IDEPOS%(DPXX%)
   '
   IDPX% = 1
   Call CARDEPOS
   For UI% = 1 To NUSU%
      If COSU%(UI%) = DPXX% Then
         IDPX% = UI%
         GoTo 99
      End If
   Next UI%
   '
99 IDEPOS% = IDPX%
   '
End Function

Sub REMIRECHA(NPROVE%, DEPX%, REMIRECH$)
   '
   FORIPRE$ = TRIM$(Control$("", "REMIRECH"))
   If FORIPRE$ = "" Then Exit Sub
   '
   If FORIPRE$ <> "" Then
      LU$ = LUDAT$
      If EXISTE%(LU$ + FORIPRE$ + ".FRM") <> 1 Then
         RMCC$ = REPLACAR$(LU$ + FORIPRE$ + ".FRM", "\", "/")
         Call COMUNI("Error de Configuración:\  \No Existe Formulario de Impresión\de Remito por Rechazo:\  \'" + RMCC$ + "'.")
         Exit Sub
      End If
   End If
   '
   Screen.MousePointer = 11
   DepOrig% = DEPX%
   DEPOSI$ = ECOARCH$("TADEPOSI", Chr$(DepOrig%))
   '
   NC% = Abs(NPROVE%)
   DESTIDOC% = (-1) * NC%
   '
   FEX = HOY(HOS$)
   FECHDOC$ = FECHATEX$(FEX)
   NUMEDOC$ = REMIRECH$
   TIPODOC$ = "Remito de Devolución"
   '
   DTP$ = OBJPLA$("TRANSPOR", DATRANSPO$)
   DATRANSPO$ = DTP$
   '
   CODPARAM$(1) = "TRANSPOR": DOCPARAM$(1) = TRIM$(Mid$(DTP$, 1, 32))
   CODPARAM$(2) = "DOMI-TRA": DOCPARAM$(2) = TRIM$(Mid$(DTP$, 33, 32))
   CODPARAM$(3) = "CUIT-TRA": DOCPARAM$(3) = TRIM$(Mid$(DTP$, 65, 16))
   CODPARAM$(4) = "PATE-TRA": DOCPARAM$(4) = TRIM$(Mid$(DTP$, 81, 16))
   CODPARAM$(5) = "COND-TRA": DOCPARAM$(5) = TRIM$(Mid$(DTP$, 97, 32))
   CODPARAM$(6) = "DOCUCOND": DOCPARAM$(6) = TRIM$(Mid$(DTP$, 129, 24))
   '
   CAPARDOC% = 6
   '
   CODTABU1$(1) = "COD-MAT"
   CODTABU1$(2) = "DES-MAT"
   CODTABU1$(3) = "CANTIDAD"
   CODTABU1$(4) = "CAKILOS"
   CACOLTAB1% = 4
   '
   AINDI% = 0
   CAITAB1% = 0
   Call LIMPIATETAB
   '
   'CONTROLES DE IMPRESION DE DESCRIPCION
   DESCPRI% = 0: DESCOMP% = 0: Especif% = 0
   If Control("BOLRECEP", "DESCPRIM") = "SI" Then DESCPRI% = 1
   If Control("BOLRECEP", "DESCOMPL") = "SI" Then DESCOMP% = 1
   If Control("BOLRECEP", "ESPECIFI") = "SI" Then Especif% = 1
   If DESCOMP% + Especif% < 1 Then DESCPRI% = 1 'POR SI NO SE SELECCIONO NADA
   '
   For YX& = 1 To ULTREG&("!BOLREDET")
      YY$ = REGLEIDO$("!BOLREDET", YX&)
      CI% = CVI(Mid$(YY$, 83, 2))
      CAN# = CVD(Mid$(YY$, 171, 8))
      LOTE$ = Mid$(YY$, 69, 12)
      '
      If CI% > 0 Then
         If CI% <= NUMAS% Then
            If CAN# > 0.005 Then
               UNID$ = UNIMED$(CI%)
               CANS$ = CSTRING$(MKD$(CAN#), 3, 9, 1, 2)
               KILO$ = CSTRING$(MKD$(CAN# * PESUNI(CI%)), 3, 12, 3, 2)
               '
               CAITAB1% = CAITAB1% + 1
               TETABU1$(1, CAITAB1%) = CODEXT$(CI%)
'               TETABU1$(2, CAITAB1%) = DESCRIT$(CI%)
               TETABU1$(3, CAITAB1%) = CANS$
               TETABU1$(4, CAITAB1%) = KILO$
               '
               DESPRIN$ = DESCRIT$(CI%)
               DESCOMPLEM$ = DESCRIT2$(CI%)
               ESPECIFICAC$ = ESPECIFIT$(CI%)
               '
               DIMPRE$ = ""
               If DESCPRI% > 0 Then
                If DESPRIN$ <> "" Then
                   DIMPRE$ = DESPRIN$
                End If
               End If
               '
               'DESCRIPCION COMPLEMENTARIA
               If DESCOMP% > 0 Then
                If DESCOMPLEM$ <> "" Then
                  If DIMPRE$ <> "" Then
                    DIMPRE$ = DIMPRE$ + vbCrLf
                  End If
                  DIMPRE$ = DIMPRE$ + DESCOMPLEM$
                End If
               End If
               '
               'ESPECIFICACION
               If Especif% > 0 Then
                If ESPECIFICAC$ <> "" Then
                  If DIMPRE$ <> "" Then
                    DIMPRE$ = DIMPRE$ + vbCrLf
                  End If
                  DIMPRE$ = DIMPRE$ + ESPECIFICAC$
                End If
               End If
               '
               'IMPRESION DE LA DESCRIPCION CON TODAS SUS VARIANTES CONFIGURADAS
               FORTELIB$ = TRIM$(UCase$(ATRIFORM$(FORIPRE$, "DES-MAT/FORMATO")))
               ATELIB% = XVALO(Mid$(FORTELIB$, 2))
               Call FRATEXTO(DIMPRE$, ATELIB%)
               For KKU1% = 1 To CARETEX%
                    TETABU1$(2, CAITAB1%) = RETEX$(KKU1%)
                    CAITAB1% = CAITAB1% + 1
               Next KKU1%
               CAITAB1% = CAITAB1% - 1
               '
               'LINEA DE SEPARACION
               Call CARLINSEP
            End If
         End If
      End If
   Next YX&
   '
   Call PRINTDOC("REMIRECH")   ' Remito de Devolucion
   '
   Call CIERRARCH("*.*")
   Screen.MousePointer = 1
   '
End Sub
'
Sub VALEMAT()
     '\\\OT-06-0244-WAR-27/06/06///
     RTA1% = ALTERNA%("Registrar Vale de Consumo \Anular Vale de Consumo")
     Select Case RTA1%
       Case 1
         GoTo 500
       Case 2
         Call ANUVALEMAT
         GoTo 9999
       Case Else
        GoTo 9999
     End Select
     '\\\OT-06-0244-WAR-FIN///
500  DepOrig% = XVALO(Control$("", "DEPOCEN"))
     DEPOSI$ = ECOARCH$("TADEPOSI", Chr$(DepOrig%))
     '
1100 Screen.MousePointer = 11
     CMO$ = TRIM$(Control$("", "CODCONSU"))
     If CMO$ = "" Then CMO$ = "VC"
     '
     Call APERBASDAT("STOCKS")
     'SQLX$ = "SELECT * FROM MOVISTO "
     'SQLX$ = SQLX$ + "WHERE CodiMovi = '" & CMO$ & "' "
     'SQLX$ = SQLX$ + "ORDER BY VAL(MID$(DoPriCor,4,7)) DESC"
     '
     'Set MOVISTEMP = Stocks.OpenRecordset(SQLX$, dbOpenSnapshot)
         
     UNRO& = 0
'     'SQLX$ = " SELECT max(substring(DoPriCor, 4,12)) as Maximo FROM MOVISTO" se comento esta linea por que le faltaba el cast abajo agregado(por ese motivo devolvia mal el valoar maximo)
'     SQLX$ = " SELECT max(cast(substring(DoPriCor, 4,12)as int)) as Maximo FROM MOVISTO"
'     SQLX$ = SQLX$ + " WHERE CodiMovi = '" & CMO$ & "' "
'
'     Dim MIRS As ADODB.Recordset
'     Set MIRS = FLEXCONN.Execute(FILTSQL$(SQLX$), Cantidad)
'     If Not IsNull(MIRS!maximo) Then
'          UNRO& = XVALO(MIRS!maximo)
'     End If
     
     'On Error Resume Next
     'With MOVISTEMP
     '  .MoveFirst
     '  If Err < 1 Then
     '    UNRO& = XVALO(Mid$(!DoPriCor, 4, 7))
     '  End If
     '  On Error GoTo 0
     'End With
     '
     NVM$ = TRIM$(ProNroComprobante(CMO$))
     UNRO& = XVALO(NVM$)
     '
1002 NROVAL& = UNRO& + 1
'     MIRS.Close
'     Set MIRS = Nothing
 
     Call FRESHALL
     '
1105 Call CIERRARCH("*.*")
     '
1111 HOII% = HOY(HOS$)
     VDEF$ = MKS$(NROVAL&) + MKI$(HOII%) + Chr$(DepOrig%) + Chr$(0) + Space$(32)
     PEDY$ = OBJPLA$("VALEMATER", VDEF$)
     If PEDY$ = "" Then GoTo 9999
     FF% = CVI(Mid$(PEDY$, 5, 2))
     DepOrig% = Asc(Mid$(PEDY$, 7, 1))
     DEPOSI$ = ECOARCH$("TADEPOSI", Chr$(DepOrig%))
     REFE$ = Mid$(PEDY$, 9, 32)
     CCOS1$ = Mid$(PEDY$, 41, 8)
     '\\\OT-06-0244-WAR-28/06/06///
     If VALIPROY%(CCOS1$, FF%) = (-1) Then
        Call MENSERR(24, "Centro de Costos No Vigente\Según Fecha Seleccionada.")
        GoTo 1111
     End If
     '\\\OT-06-0244-WAR-FIN///
     If TRIM$(REFE$) = "" Then
        Call MENSERR(24, "Referencia o Motivo es Mandatoria")
        GoTo 1111
     End If
     If DepOrig% < 1 Or DepOrig% >= 99 Or TRIM$(DEPOSI$) = "" Then
        Call MENSERR(24, "Depósito de Salida no Válido")
        GoTo 1111
     End If
     '
1112    Call SETULTREG("!BOLREDET", 0)
     Call CIERRARCH("*.*")
     '
     Screen.MousePointer = 11
     '
     fff$ = REPLACAR$(FECHATEX$(FF%), "/", "-")
     AVENTA$ = "BOLREDET"
     If EXISTE%("TRAZAB07.EXE") > 0 Or EXISTE%("TRAZAB08.EXE") > 0 Then AVENTA$ = "MOVITRAZ"
1113 VTB% = VENTABU%(AVENTA$ + "/NC/TITUPAN=Vale de Materiales - " + TRIM$(DEPOSI$) + " - " + fff$)
     If VTB% < 0 Or ULTREG&("!BOLREDET") < 1 Then
         GoTo 1105
     End If
     '
     CONTA% = ULTREG&("!BOLREDET")
     '
1115 If CONTA% > 0 Then
       Screen.MousePointer = 11  'Call BLIMESS("Registrando Recepcion ...")
       '
       '\\\OT-05-0730-WAR-10/11/05///
       CAITMO% = 0
       For YX& = 1 To ULTREG&("!BOLREDET")
           YY$ = REGLEIDO$("!BOLREDET", YX&)
           CI% = CVI(Mid$(YY$, 83, 2))
           CAN# = CVD(Mid$(YY$, 89, 8))
           If CI% > 0 Then
             If CI% <= NUMAS% Then
               If CAN# > 0.005 Then
                 CAITMO% = CAITMO% + 1
                 CIMOVIM%(CAITMO%) = CI%
                 CAMOVIM(CAITMO) = (-1) * CAN#
               End If
             End If
           End If
       Next YX&
       If PUEDEMOVI%(DepOrig%) < 0 Then GoTo 1113 ' CANCELA SALIDA
       '\\\OT-05-0730-WAR-FIN///
       '
       ' VALIDA LOS DATOS INGRESADOS
       HUBOSELOT% = 0
       For YX& = 1 To ULTREG&("!BOLREDET")
          YY$ = REGLEIDO$("!BOLREDET", YX&)
          CI% = CVI(Mid$(YY$, 83, 2))
          If ESTRAZABLE%(CI%) > O Then
              ' OBLIGA A INGRESAR EL LOTE
              ' COMPARA CANTIDAD DESPACHADA CON SALDO DEL LOTE
              CANTSA = CVD(Mid$(YY$, 89, 8))
              DEPX% = DepOrig%
              LOTEX$ = TRIM$(Mid$(YY$, 107, 12))
              If LOTEX$ <> "" Then
                 If SaldoLote(CI%, LOTEX$, DEPX%) >= CANTSA Then
                    GoTo 19
                 End If
              End If
              LOTEX$ = LOTESELEC$(CI%, DEPX%, CANTSA)
              HUBOSELOT% = 1
              If LOTEX$ = "" Then
                 GoTo 1113
              End If
19             Call REPLA(YY$, LOTEX$, 69, 12)
              Call REPLA(YY$, LOTEX$, 107, 12)
              Call GRAREG("!BOLREDET", YY$, YX&)
          End If
       Next YX&
       Call CIERRARCH("!BOLREDET")
       If HUBOSELOT% = 1 Then GoTo 1113
       '
       Call SETULTREG("!OBSERVOF", 0)
       Call CIERRARCH("!OBSERVOF")
       VTB% = VENTABU%("OBSERVOF/TITUPAN=Vale de Materiales - " + TRIM$(DEPOSI$) + " - " + fff$)
       '
       AINDI% = 0
       NUORIT% = (-1)
       For YX& = 1 To ULTREG&("!BOLREDET")
         YY$ = REGLEIDO$("!BOLREDET", YX&)
         CI% = CVI(Mid$(YY$, 83, 2))
         CAN# = CVD(Mid$(YY$, 89, 8))
         LOTE$ = Mid$(YY$, 69, 12)
           If TRIM$(LOTE$) = "" Then LOTE$ = Mid$(YY$, 107, 12)
         REMI$ = CMO$ + "." + TRIM$(Str$(NROVAL&))
         If CI% > 0 Then
           If CI% <= NUMAS% Then
             If CAN# > 0.005 Then
               DOSEC$ = ""
1120           CANTI = (-1) * CAN#
               NUORIT% = NUORIT% + 1
               '\\\OT-06-0244-WAR-20/06/06///
               'Si Tiene Centro de Costos => Valorizo
               VUNI# = 0
               If CCOS1$ <> "" Then
                 VUNI# = COSPESOS(CI%)
                 DOSEC$ = "CC-" + CCOS1$
               End If
               '\\\OT-06-0244-WAR-FIN///
               Call MOVISTOK(FF%, CI%, LOTE$, CMO$, REMI$, REMI$, DOSEC$, DOSEC$, REFE$, NUORIT%, VUNI#, 0, 0, DepOrig%, CANTI) '\\\OT-06-0244-WAR-20/06/06///
             End If
           End If
         End If
       Next YX&
       Call CIERRARCH("MOVISTO")
       '
       FORIPRE$ = TRIM$(Control$("", "FORCONSU"))
       If FORIPRE$ <> "" Then
         LU$ = LUDAT$
         If EXISTE%(LU$ + FORIPRE$ + ".FRM") <> 1 Then
           RMCC$ = REPLACAR$(LU$ + FORIPRE$ + ".FRM", "\", "/")
           Call COMUNI("Error de Configuración:\  \No Existe Formulario de Impresión\de Vale de Consumo de Materiales:\  \'" + RMCC$ + "'.")
           GoTo 9999
         End If
       End If
       '
       If TRIM$(FORIPRE$) <> "" Then
         '
         'VALIDACION DE MODO DE IMPRESION
         MODOIMPRE07.Check7.Caption = "Imprime Vale de Consumo"
         MODOIMPRE07.Label1.Caption = "" 'LIMPIA EL LABEL DONDE PONE OK, SI ACEPTA
         MODOIMPRE07.Check8.Value = 0
         Screen.MousePointer = 1
         MODOIMPRE07.Show 1
         CANCELPRINT% = 0
         If MODOIMPRE07.Label1.Caption = "OK" Then
          If MODOIMPRE07.Check8.Value > 0 Then CANCELPRINT% = 1
         End If
         Screen.MousePointer = 11
         '
         FEX = FF%
         FECHDOC$ = FECHATEX$(FEX)
         NUMEDOC$ = REMI$
         TIPODOC$ = "Vale de Consumo de Materiales"
         '
         CODPARAM$(1) = "REF-MAT1": DOCPARAM$(1) = REFE$
         CODPARAM$(2) = "DEPOSITO": DOCPARAM$(2) = DEPOSI$
         ' Muestra el Centro de Costo en el formulario de impresion
         CODPARAM$(3) = "REF-MAT2": DOCPARAM$(3) = TRIM$(CCOS1$) + " - " + TRIM$(ECOARCH("TACENCOS", TRIM$(CCOS1$)))
         CAPARDOC% = 3
         '
         CODTABU1$(1) = "COD-MPRI"
         CODTABU1$(2) = "DES-MPRI"
         CODTABU1$(3) = "CAN-MPRI"
         CODTABU1$(4) = "CAKILOS"
         CODTABU1$(5) = "UNI-MPRI"
         CODTABU1$(6) = "UBI-MPRI"
         CACOLTAB1% = 6
         '
         DESTIDOC% = 0
         CAITAB1% = 0
         '
         For YX& = 1 To ULTREG&("!BOLREDET")
           YY$ = REGLEIDO$("!BOLREDET", YX&)
           CI% = CVI(Mid$(YY$, 83, 2))
           CAN# = CVD(Mid$(YY$, 89, 8))
           If CI% > 0 Then
             If CI% <= NUMAS% Then
               If CAN# > 0.005 Then
                 UNID$ = TRIM$(Mid$(YY$, 85, 4))
                 If UNID$ = "" Then UNID$ = UNIMED$(CI%)
                 CANS$ = CSTRING$(MKD$(CAN#), 3, 9, 1, 2)
                 KILO$ = CSTRING$(MKD$(CAN# * PESUNI(CI%)), 3, 12, 3, 2)
                 NULOTE$ = TRIM$(Mid$(YY$, 69, 12))
                 NUOCO$ = CSTRING$(Mid$(YY$, 151, 4), 3, 7, 0, 2)
                 '
                 CAITAB1% = CAITAB1% + 1
                 TETABU1$(1, CAITAB1%) = CODEXT$(CI%)
                 TETABU1$(2, CAITAB1%) = DESCRIT$(CI%)
                 TETABU1$(3, CAITAB1%) = CANS$
                 TETABU1$(4, CAITAB1%) = KILO$
                 TETABU1$(5, CAITAB1%) = UNID$
                 TETABU1$(6, CAITAB1%) = UBISTO$(CI%) '
               End If
             End If
           End If
         Next YX&
         '
         Call PRINTDOC("FORCONSU")   ' Remito en Consignacion
       End If
       '
1199 End If
     Screen.MousePointer = 1
     GoTo 1100
     '
9999 Screen.MousePointer = 1
     Call CIERRARCH("*.*")
     '
End Sub

Sub DESCRAP()
     '
     EPAC$ = TRIM$(UCase$(EMPREAC$))
     DepOrig% = XVALO(Control$("", "DEPOCEN"))
     If InStr(EPAC$, "DOSIVA") > 0 Then
       DepOrig% = 91
     End If
     DEPOSI$ = ECOARCH$("TADEPOSI", Chr$(DepOrig%))
     '
1100 Screen.MousePointer = 11
     CMO$ = TRIM$(Control$("", "CODSCRAP"))
     If CMO$ = "" Then CMO$ = "SC"
     '
     Call APERBASDAT("STOCKS")
     sqlx$ = " SELECT max(substring(DoPriCor, 4,12)) as Maximo FROM MOVISTO"
     sqlx$ = sqlx$ + " WHERE CodiMovi = '" & CMO$ & "'"
      
     UNRO& = 0
     Dim MIRS As ADODB.Recordset
     Set MIRS = FLEXCONN.Execute(FILTSQL$(sqlx$))
     If Not IsNull(MIRS!maximo) Then  'Revisar_Ale (ver sc-dopricor of-000010-1)
        UNRO& = XVALO(MIRS!maximo)
     End If
1002 NROVAL& = UNRO& + 1
     MIRS.Close
     Set MIRS = Nothing
    
     Call FRESHALL

     'SQLX$ = "SELECT * FROM MOVISTO "
     'SQLX$ = SQLX$ + "WHERE CodiMovi = '" & CMO$ & "' "
     'SQLX$ = SQLX$ + "ORDER BY substring(DoPriCor,4,12) DESC"
     '
     'Set MOVISTEMP = Stocks.OpenRecordset(SQLX$, dbOpenSnapshot)
     'UNRO& = 0
    'Set MoviTemp = Stocks.OpenRecordset(sqlx$, dbOpenSnapshot)
    'On Error Resume Next
    'MoviTemp.MoveFirst
    'If Err < 1 Then
      'UNUM& = XVALO(Mid$(MoviTemp!DoPriCor, 4))
    'End If
    
     
     'On Error Resume Next
     'With MOVISTEMP
     '  .MoveFirst
     '  If Err < 1 Then
     '    UNRO& = XVALO(Mid$(!DoPriCor, 4, 7))
     '  End If
     '  On Error GoTo 0
     'End With
     '
'1002'NROVAL& = UNRO& + 1
     
     'Call FRESHALL

     '
1105 Call CIERRARCH("*.*")
     '
1111 HOII% = HOY(HOS$)
     VDEF$ = MKS$(NROVAL&) + MKI$(HOII%) + Chr$(DepOrig%) + Chr$(0) + Space$(32) + USERNAME$
     PEDY$ = OBJPLA$("DESCARSCRAP", VDEF$)
     If PEDY$ = "" Then GoTo 9999
     FF% = CVI(Mid$(PEDY$, 5, 2))
     DepOrig% = Asc(Mid$(PEDY$, 7, 1))
     DEPOSI$ = ECOARCH$("TADEPOSI", Chr$(DepOrig%))
     REFE$ = Mid$(PEDY$, 9, 32)
     If TRIM$(REFE$) = "" Then
       Call MENSERR(24, "Referencia o Motivo es Mandatoria")
       GoTo 1111
     End If
     If DepOrig% < 1 Or DepOrig% >= 99 Or TRIM$(DEPOSI$) = "" Then
       Call MENSERR(24, "Depósito de Salida no Válido")
       GoTo 1111
     End If
     '
1112 Call SETULTREG("!BOLREDET", 0)
     Call CIERRARCH("*.*")
     '
     Screen.MousePointer = 11
     '
     fff$ = REPLACAR$(FECHATEX$(FF%), "/", "-")
     '
     AVENTA$ = "BOLREDET"
     If EXISTE%("TRAZAB07.EXE") > 0 Or EXISTE%("TRAZAB08.EXE") > 0 Then AVENTA$ = "MOVITRAZ"
1113 VTB% = VENTABU%(AVENTA$ + "/NC/TITUPAN=Descarte de Scrap - " + TRIM$(DEPOSI$) + " - " + fff$)
     If VTB% < 0 Or ULTREG&("!BOLREDET") < 1 Then
       GoTo 1105
     End If
     '
     CONTA% = ULTREG&("!BOLREDET")
     '
1115 If CONTA% > 0 Then
       '
       Screen.MousePointer = 11  'Call BLIMESS("Registrando Recepcion ...")
       '
       '
       '\\\OT-05-0730-WAR-10/11/05///
       CAITMO% = 0
       For YX& = 1 To ULTREG&("!BOLREDET")
           YY$ = REGLEIDO$("!BOLREDET", YX&)
           CI% = CVI(Mid$(YY$, 83, 2))
           CAN# = CVD(Mid$(YY$, 89, 8))
           If CI% > 0 Then
             If CI% <= NUMAS% Then
               If CAN# > 0.005 Then
                 CAITMO% = CAITMO% + 1
                 CIMOVIM%(CAITMO%) = CI%
                 CAMOVIM(CAITMO) = (-1) * CAN#
               End If
             End If
           End If
       Next YX&
       If PUEDEMOVI%(DepOrig%) < 0 Then GoTo 1113 ' CANCELA SALIDA
       '\\\OT-05-0730-WAR-FIN///
       '
       ' VALIDA LOS DATOS INGRESADOS
       HUBOSELOT% = 0
       For YX& = 1 To ULTREG&("!BOLREDET")
          YY$ = REGLEIDO$("!BOLREDET", YX&)
          CI% = CVI(Mid$(YY$, 83, 2))
          If ESTRAZABLE%(CI%) > O Then
              ' OBLIGA A INGRESAR EL LOTE
              ' COMPARA CANTIDAD DESPACHADA CON SALDO DEL LOTE
              CANTSA = CVD(Mid$(YY$, 89, 8))
              DEPX% = DepOrig%
              LOTEX$ = TRIM$(Mid$(YY$, 107, 12))
              If LOTEX$ <> "" Then
                 If SaldoLote(CI%, LOTEX$, DEPX%) >= CANTSA Then
                    GoTo 19
                 End If
              End If
              LOTEX$ = LOTESELEC$(CI%, DEPX%, CANTSA)
              HUBOSELOT% = 1
              If LOTEX$ = "" Then
                 GoTo 1113
              End If
19            Call REPLA(YY$, LOTEX$, 69, 12)
              Call REPLA(YY$, LOTEX$, 107, 12)
              Call GRAREG("!BOLREDET", YY$, YX&)
          End If
       Next YX&
       Call CIERRARCH("!BOLREDET")
       If HUBOSELOT% = 1 Then GoTo 1113
       '
       Call SETULTREG("!OBSERVOF", 0)
       Call CIERRARCH("!OBSERVOF")
       VTB% = VENTABU%("OBSERVOF/TITUPAN=Descarte de Scrap - " + TRIM$(DEPOSI$) + " - " + fff$)
       '
       AINDI% = 0
       For YX& = 1 To ULTREG&("!BOLREDET")
         YY$ = REGLEIDO$("!BOLREDET", YX&)
         CI% = CVI(Mid$(YY$, 83, 2))
         CAN# = CVD(Mid$(YY$, 89, 8))
         LOTE$ = Mid$(YY$, 69, 12)
         REMI$ = CMO$ + "." + TRIM$(Str$(NROVAL&))
         If CI% > 0 Then
           If CI% <= NUMAS% Then
             If CAN# > 0.005 Then
               DOSEC$ = REMI$
1120           CANTI = (-1) * CAN#
               NUORIT% = NUORIT% + 1
               Call MOVISTOK(FF%, CI%, LOTE$, CMO$, REMI$, REMI$, DOSEC$, DOSEC$, REFE$, NUORIT%, 0, 0, 0, DepOrig%, CANTI)
             End If
           End If
         End If
         '
       Next YX&
       Call CIERRARCH("MOVISTO")
       '
       FORIPRE$ = TRIM$(Control$("", "FORSCRAP"))
       If FORIPRE$ <> "" Then
         LU$ = LUDAT$
         If EXISTE%(LU$ + FORIPRE$ + ".FRM") <> 1 Then
           RMCC$ = REPLACAR$(LU$ + FORIPRE$ + ".FRM", "\", "/")
           Call COMUNI("Error de Configuración:\  \No Existe Formulario de Impresión\de Vale de Descarte de Scrap:\  \'" + RMCC$ + "'.")
           GoTo 9999
         End If
       End If
       '
       If TRIM$(FORIPRE$) <> "" Then
         '
         'VALIDACION DE MODO DE IMPRESION
         MODOIMPRE07.Check7.Caption = "Imprime Vale de Scrap"
         MODOIMPRE07.Label1.Caption = "" 'LIMPIA EL LABEL DONDE PONE OK, SI ACEPTA
         MODOIMPRE07.Check8.Value = 0
         Screen.MousePointer = 1
         MODOIMPRE07.Show 1
         CANCELPRINT% = 0
         If MODOIMPRE07.Label1.Caption = "OK" Then
          If MODOIMPRE07.Check8.Value > 0 Then CANCELPRINT% = 1
         End If
         Screen.MousePointer = 11
         '
         FEX = FF%
         FECHDOC$ = FECHATEX$(FEX)
         NUMEDOC$ = REMI$
         TIPODOC$ = "Vale de Descarte de Scrap"
         '
         CODPARAM$(1) = "REF-MAT1": DOCPARAM$(1) = REFE$
         CODPARAM$(2) = "DEPOSITO": DOCPARAM$(2) = DEPOSI$
         CAPARDOC% = 2
         '
         CODTABU1$(1) = "COD-MPRI"
         CODTABU1$(2) = "DES-MPRI"
         CODTABU1$(3) = "CAN-MPRI"
         CODTABU1$(4) = "CAKILOS"
         CODTABU1$(5) = "UNI-MPRI"
         CODTABU1$(6) = "UBI-MPRI"
         CACOLTAB1% = 6
         '
         DESTIDOC% = 0
         CAITAB1% = 0
         '
         For YX& = 1 To ULTREG&("!BOLREDET")
           YY$ = REGLEIDO$("!BOLREDET", YX&)
           CI% = CVI(Mid$(YY$, 83, 2))
           CAN# = CVD(Mid$(YY$, 89, 8))
           If CI% > 0 Then
             If CI% <= NUMAS% Then
               If CAN# > 0.005 Then
                 UNID$ = TRIM$(Mid$(YY$, 85, 4))
                 If UNID$ = "" Then UNID$ = UNIMED$(CI%)
                 CANS$ = CSTRING$(MKD$(CAN#), 3, 9, 1, 2)
                 KILO$ = CSTRING$(MKD$(CAN# * PESUNI(CI%)), 3, 12, 3, 2)
                 NULOTE$ = TRIM$(Mid$(YY$, 69, 12))
                 NUOCO$ = CSTRING$(Mid$(YY$, 151, 4), 3, 7, 0, 2)
                 '
                 CAITAB1% = CAITAB1% + 1
                 TETABU1$(1, CAITAB1%) = CODEXT$(CI%)
                 TETABU1$(2, CAITAB1%) = DESCRIT$(CI%)
                 TETABU1$(3, CAITAB1%) = CANS$
                 TETABU1$(4, CAITAB1%) = KILO$
                 TETABU1$(5, CAITAB1%) = UNID$
                 TETABU1$(6, CAITAB1%) = UBISTO$(CI%) '
               End If
             End If
           End If
         Next YX&
         '
         Call PRINTDOC("FORSCRAP")   ' Remito en Consignacion
         '
       End If
       '
1199 End If
     Screen.MousePointer = 1
     GoTo 1100
     '
9999 Screen.MousePointer = 1
     Call CIERRARCH("*.*")
     '
End Sub

'
'\\\OT-06-0244-WAR-27/06/06///
Sub ANUVALEMAT()
  '
  Screen.MousePointer = 11
  '
  If DERACCE%("ANUVAMAT", "Anulacion de Vale de Materiales") < 2 Then
    GoTo 99
  End If
  '
  Call APERBASDAT("STOCKS")
  CMO$ = TRIM$(Control$("", "CODCONSU"))
  If CMO$ = "" Then CMO$ = "VC"
  '
  'sqlx$ = ""
  'sqlx$ = "SELECT  DoPRICor, ReferCor, FechMov FROM MOVISTO "
  'sqlx$ = sqlx$ + "WHERE CodiMovi = '" & CMO$ & "' "
  'sqlx$ = sqlx$ + " Order by VAL(MID$(DoPriCor,4,7)) ASC"
  'Set VAMTMP = Stocks.OpenRecordset(sqlx$, 4)
  'With VAMTMP
  '  On Error Resume Next
  '  .MoveFirst
  '  If Err Then
  '    Call MENSERR(24, "No Existen Vales de Materiales\para su Anulación")
  '    On Error GoTo 0
  '    GoTo 99
  '  End If
  '  On Error GoTo 0
  Dim VAMTMP As ADODB.Recordset
  Set VAMTMP = New ADODB.Recordset
  sqlx$ = ""
  sqlx$ = "SELECT  DoPRICor, ReferCor, FechMov FROM MOVISTO "
  sqlx$ = sqlx$ + "WHERE CodiMovi = '" & CMO$ & "' "
  sqlx$ = sqlx$ + " Order by substring(DoPriCor,4,12) ASC"
  
  Set VAMTMP = FLEXCONN.Execute(FILTSQL$(sqlx$))
  If (VAMTMP.EOF And VAMTMP.BOF) Then 'si esta vacio
      Call MENSERR(24, "No Existen Vales de Materiales\para su Anulación")
      On Error GoTo 0
      GoTo 99
  End If
  With VAMTMP
  
    JJ& = 0
    REBLA$ = REGBLAN$("INVALEMA")
    Do While .EOF = False
      YY$ = REBLA$
      NUVAMA& = XVALO(Mid$(!DoPriCor, 4))
      If NUMEA1& <> NUVAMA& Then
        Call REPLA(YY$, MKS$(NUVAMA&), 1, 4)
        REFER1$ = ""
        F1% = CVINT(!Fechmov)
        F2$ = FECHATEX$(F1%)
        REFER1$ = F2$ + " - " + SAFETEXT$(!ReferCor)
        Call REPLA(YY$, REFER1$, 5, 60)
        JJ& = JJ& + 1
        Call GRAREG("INVALEMA", YY$, JJ&)
        NUMEA1& = NUVAMA&
      End If
      .MoveNext
    Loop
    Call SETULTREG("INVALEMA", JJ&)
    Call FRESHECHO("INVALEMA")
  End With
  '
5 Call SELECHO("INVALEMA")
  NUVAMA1& = XVALO(VALACT1$("INVALEMA"))
  If NUVAMA1& > 0 Then
    OPX1% = COMALTER%("¿ Confirma Anular el Vale: " + CMO$ + "." + TRIM$(Str$(NUVAMA1&)) + " ?\\No, Cancelar\Si, Anular")
    If OPX1% = 2 Then
      DOCVAMA1$ = CMO$ + "." + TRIM$(Str$(NUVAMA1&))
      STRSQL = "Delete from movisto where DoPricor='" & DOCVAMA1$ & "' "
      FLEXCONN.Execute STRSQL, CANTIDAD
      If CANTIDAD = 0 Then
          Call MENSERR(24, "No Existen Vales de Materiales\para su Anulación")
          On Error GoTo 0
          GoTo 99
      Else
         Call COMUNI("Anulación Completa")
      End If
      
      'With Movisto
      '  DOCVAMA1$ = CMO$ + "." + TRIM$(Str$(NUVAMA1&))
      '  SQLX1$ = "DoPriCor = '" & DOCVAMA1$ & "' "
      '  .FindFirst SQLX1$
7     '  If .NoMatch = False Then
      '    .Delete
      '    .FindNext SQLX1$
      '    GoTo 7
      '  End If
      'End With
      
      'Call COMUNI("Anulación Completa")
    End If
  End If
  '
99 Screen.MousePointer = 1
End Sub
'\\\OT-06-0244-WAR-27/06/06///

Function STATOCOM%(NUOCOMX&)
   Static NOCOA&, STAOC%
   ' DEVUELVE EL STATUS DE LA ORDEN DE COMPRA
   If NUOCOMX& <> NOCOA& Then
     STAOC% = 9
     NOCOA& = NUOCOMX&
     Call APERBASDAT("COMPRAS")
     '
     STRSQL = "Select Nume_Ocom,Stat_Ocom from OcomEnca where Nume_Ocom=" & NUOCOMX&
     Dim OCOMENCA As ADODB.Recordset
     Set OCOMENCA = FLEXCONN.Execute(FILTSQL$(STRSQL))
     If Not (OCOMENCA.EOF And OCOMENCA.BOF) Then
       STAOC% = OCOMENCA!Stat_Ocom
     End If
     OCOMENCA.Close
     Set OCOMENCA = Nothing
   End If
   '
   STATOCOM% = STAOC%
End Function

Sub CREALOTIN(FALTA%, FEFAB%, FEVEN%, CIXI%, IDLOT$, MARCAX$, LOTORI$, CANOM#, _
              Optional LARGOM, Optional ANCHOMM, Optional DPX%, Optional UBX$, _
              Optional PROVED$, Optional DOCIN$, Optional ENVAS$, Optional STALIB$, _
              Optional CANSEC, Optional DESTIFA$)
   '
   IDLOTX$ = TRIM$(IDLOT$)
   '
   Call APERBASDAT("STOCKS")
   sqlx$ = " SELECT * FROM LOTINGRE "
   sqlx$ = sqlx$ + "WHERE Cod_Inte = " & CIXI% & " "
   sqlx$ = sqlx$ + "AND IdenLote = '" & IDLOTX$ & "' "
   '
   Dim LOTINGTEMP As ADODB.Recordset
   Set LOTINGTEMP = New ADODB.Recordset
25 On Error Resume Next
   LOTINGTEMP.Open FILTSQL$(sqlx$), FLEXCONN, adOpenDynamic, adLockPessimistic, adCmdText
   If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
      On Error GoTo 0
      Call CapturaErrorOpen
      GoTo 25
   End If
   On Error GoTo 0
   '
   With LOTINGTEMP
     If (.EOF And .BOF) Then
       .AddNew   ' YA CONVERTIDO
       On Error Resume Next
       !CodiEmpr = CodiEmp%
       On Error GoTo 0
       !CanConsu = 0
       !CanReser = 0
     End If
     !Cod_Inte = CIXI%
     !IdenLote = Left$(IDLOTX$, 32)
     !FechAlta = CVDAT(FALTA%)
     !FechFab = CVDAT(FEFAB%)
     !FechVenc = CVDAT(FEVEN%)
     !CanAlta = CANOM#
     If !CANSALDO < 0.05 Or !CANSALDO > CANOM# Then
        !CANSALDO = CANOM#
     End If
     !AnchoMili = ANCHOMM
     !MetroIni = LARGOM
     !SALMETRO = LARGOM
     !DepoUbic = DPX%
     !UBICFIS = Left$(UBX$, 32)
     !MARCA = Left$(MARCAX$, 32)
     !LOTECOL_PROV = Left$(LOTORI$, 16)
     !DocuIngre = Left$(DOCIN$, 24)
     !CANTIDAD = CANSEC
     !CodEnvase = Left$(ENVAS$, 64)
     !ORIGEN = Left$(PROVED$, 255)
     !Libera = Left$(STALIB$, 16)
     !Cod_Desti = Left$(DESTIFA$, 16)
     .Update
   End With
   LOTINGTEMP.Close
   Set LOTINGTEMP = Nothing
   '
   ' APLICA MARCA Y LOTE DE ORIGEN A TODOS LOS SUB-LOTES
   sqlx$ = "UPDATE LOTINGRE SET MARCA = '" & MARCAX$ & "', LoteCol_Prov = '" & LOTORI$ & "'"
   sqlx$ = sqlx$ + " WHERE Cod_Inte = " & CIXI% & " "
   sqlx$ = sqlx$ + "AND IdenLote LIKE '" & IDLOTX$ & "%' "
   FLEXCONN.Execute sqlx$
   '
   Call ACSALOTE(CIXI%, IDLOTX$, "NOMESS")
   '
End Sub
'

Sub ACSALOTE(CIXI%, IDLO$, Optional MODOX$)    ' ACTUALIZA SALDOS POR LOTE
    '
    IDLOTX$ = TRIM$(IDLO$)
    If IDLOTX$ = "" Then
       Exit Sub
    End If
    '
    Call APERBASDAT("STOCKS")
    'SQLZ$ = " SELECT SUM(CantIngre) AS SUING, SUM(CantSalid) AS SUCON FROM MOVISTO WITH(NOLOCK)"
    SQLZ$ = " SELECT CODIMOVI, CantIngre, CantSalid, FECHMOV FROM MOVISTO WITH(NOLOCK)"
    SQLZ$ = SQLZ$ + " WHERE Cod_Inte = " & CIXI% & " "
    SQLZ$ = SQLZ$ + " AND IdenLote = '" & IDLOTX$ & "'"
    SQLZ$ = SQLZ$ + " ORDER BY FECHMOV ASC, FEREMOVI ASC"
    '
    Dim MOVCONTEMP As ADODB.Recordset
    Set MOVCONTEMP = New ADODB.Recordset
    Set MOVCONTEMP = FLEXCONN.Execute(FILTSQL$(SQLZ$))
    
    SUINGRE# = 0: SUCONSU# = 0: FEPRIMO$ = ""
    With MOVCONTEMP
      If Not (.BOF And .EOF) Then
         Do While Not .EOF
            If FEPRIMO$ = "" Then FEPRIMO$ = FECHATEX$(CVINT(!Fechmov))
            CMOX$ = TRIM$(SAFETEXT$(!CodiMovi))
            CAING# = XVALO(!CantIngre)
            CASAL# = XVALO(!Cantsalid)
            If CAING# > 0 Then
               If CMOX$ = "IR" Or CMOX$ = "BR" Or CMOX$ = "OP" Or CMOX$ = "OF" Or CMOX$ = "FL" Or CMOX$ = "DV" Then GoTo 15
               If CMOX$ = "OD" Then If SUINGRE# < 0.00005 Then GoTo 15
            End If
            If CMOX$ = "AJ" Then
               FEMOVI$ = FECHATEX$(CVINT(!Fechmov))
               If FEMOVI$ = FEPRIMO$ Then
                  CAING# = CAING# - CASAL#
                  CASAL# = 0
                  GoTo 15 ' LOS AJUSTES DEL MISMO DIA DE ALTA SE SUMAN
               End If
            End If
            'If TRIM$(SAFETEXT$(!CodiMovi)) = "CC" Then
            ' RESTANTES MOVIMIENTOS SON TODOS POR SALIDA
            CASAL# = CASAL# - CAING#
            CAING# = 0
            ' End If
15          SUINGRE# = SUINGRE# + CAING#
            SUCONSU# = SUCONSU# + CASAL#
         .MoveNext
         Loop
      End If
    End With
    MOVCONTEMP.Close
    Set MOVCONTEMP = Nothing
    '
    Call APERBASDAT("RESERVA")
    SQLZ$ = " SELECT SUM(CantRes) AS SURES, SUM (CantCons) AS SUCON FROM RESERVA WITH(NOLOCK)"
    SQLZ$ = SQLZ$ + " WHERE Cod_Inte = " & CIXI% & " "
    SQLZ$ = SQLZ$ + " AND IdenLote = '" & IDLOTX$ & "'"
    SQLZ$ = SQLZ$ + " AND CantRes >= CantCons"
    SQLZ$ = SQLZ$ + " AND (Status = 0 OR STATUS IS NULL)"     ' para no considerar los marcados para anular
    '
    Dim RESERTMP As ADODB.Recordset
    Set RESERTMP = New ADODB.Recordset
    Set RESERTMP = FLEXCONN.Execute(FILTSQL$(SQLZ$))
    '
    SURESER# = 0
    With RESERTMP
      On Error Resume Next
      SURESER# = XVALO(!SURES) - XVALO(!SUCON)
      On Error GoTo 0
    End With
    RESERTMP.Close
    Set RESERTMP = Nothing
    '
    CONDI$ = "Cod_Inte = " & CIXI% & " "
    CONDI$ = CONDI$ + " AND IdenLote = '" & IDLOTX$ & "'"
    CONDI$ = CONDI$ + " AND Kilo_Asig <> 0"
    SUASIG# = XVALO(VALOBADA("BALAREQMO", "SUM(Kilo_Asig)", CONDI$))
    SUASIG# = SUASIG# - XVALO(VALOBADA("BALAREQMO", "SUM(Kilo_Consu)", CONDI$))
    If SUASIG# < 0 Then SUASIG# = 0
    SURESER# = SURESER# + SUASIG#
    '
20  sqlx$ = "SELECT * FROM LOTINGRE "
    sqlx$ = sqlx$ + " WHERE IdenLote = '" & IDLOTX$ & "'"
    sqlx$ = sqlx$ + " AND Cod_Inte = " & CIXI% & " "
    '
    Dim LOTINSEL As ADODB.Recordset
    Set LOTINSEL = New ADODB.Recordset
23  On Error Resume Next
    LOTINSEL.Open FILTSQL$(sqlx$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
    If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
      On Error GoTo 0
      Call CapturaErrorOpen
      GoTo 23
    End If
    On Error GoTo 0
    '
    UBIACT$ = ""
    With LOTINSEL
       On Error Resume Next
       .MoveFirst
       If Err Then
          On Error GoTo 0
          If UCase$(MODOX$) <> "NOMESS" Then
             Call MENSERR(24, "Lote '" + IDLOTX$ + "' no Encontrado.\Probable Falta Trazar Ingreso.")
          End If
          Call CONECRUN("TRAZAB07", "Registro de Trazabilidad de Componentes")
          Exit Sub
       End If
       On Error GoTo 0
       '
       CSALDO# = SUINGRE# - SUCONSU# - SURESER#
       CSALNE# = SUINGRE# - SUCONSU#
       KIBRUT# = 0
       If UCase$(TRIM$(UNIMED$(CIXI%))) = "KG" Then
          KIBRUT# = CSALNE# + XVALO(!Tara)
       End If
       '
       FRAMI = FRACMIN(CIXI%)
       If CSALDO# < (-1) * FRAMI Then
          If UCase$(MODOX$) <> "NOMESS" Then
             Call MENSERR(24, "Exceso de Consumo en Código '" + CODEXT$(CIXI%) + "' - Lote '" + IDLOTX$ + "'\Saldo Ajustado a Cero.")
          End If
          CSALDO# = 0
          GoTo 25
       End If
       '
       CANCON = XVALO(!CanConsu)
       CANRES = XVALO(!CanReser)
       If Abs(CSALDO#) < FRAMI Then CSALDO# = 0
       If Abs(CSALNE#) < FRAMI Then CSALNE# = 0
       '
       If Abs(SUINGRE# - XVALO(!CanAlta)) >= FRAMI / 10 Then
            GoTo RECULOTE
          ElseIf Abs(CSALDO# - XVALO(!CANSALDO)) >= FRAMI / 10 Then
            GoTo RECULOTE
          ElseIf Abs(CSALNE# - XVALO(!SalNeto)) >= FRAMI / 10 Then
            GoTo RECULOTE
          ElseIf Abs(SUCONSU# - CANCON) >= FRAMI / 10 Then
            GoTo RECULOTE
          ElseIf Abs(SURESER# - CANRES) >= FRAMI / 10 Then
            GoTo RECULOTE
          Else
            ancho# = XVALO(!AnchoMili)
            If ancho# > 0 Then
               LARGONU# = 1000 * CSALNE# / ancho# ' SE AGREGA PARA CUANDO SE REALIZO UN AJUSTE POR UNA CANTIDAD MAYOR A LA DEL LOTE
               If Abs(LARGONU# - XVALO(!SALMETRO)) >= 0.05 Then
                  GoTo RECULOTE
               End If
            End If
       End If
       GoTo 35
       '
RECULOTE:
       If UCase$(MODOX$) <> "NOMESS" Then
          Call MENSERR(24, "Saldos en Código '" + CODEXT$(CIXI%) + "' - Lote '" + IDLOTX$ + "'\Ajustado Según Consumos.")
       End If
       '
25     !CanAlta = SUINGRE#
       !CanConsu = SUCONSU#
       !CanReser = SURESER#
       !CANSALDO = CSALDO#
       !SalNeto = CSALNE#
       !KilBruto = KIBRUT#
       '
       ' GRABA EL LARGO AJUSTADO AL CONSUMO
       ancho# = XVALO(!AnchoMili)
       If ancho# > 0 Then
          LARGONU# = 1000 * CSALNE# / ancho# ' SE AGREGA PARA CUANDO SE REALIZO UN AJUSTE POR UNA CANTIDAD MAYOR A LA DEL LOTE
          If LARGONU# < 0 Then
             LARGONU# = 0
          End If
          !SALMETRO = LARGONU#
       End If
       .Update
    End With
    '
35  LOTINSEL.Close
    Set LOTINSEL = Nothing
    '
    If UBIACT$ <> "" Then
       If CantRegistros("LOTINGRE", "UbicFis = '" & UBIACT$ & "'") = 0 Then
          Call ACTUREGISTRO("UBIPRE", "Status", "CodiUbic = '" & UBIACT$ & "'", 0, REG&)
       End If
    End If
    '

End Sub
'
Sub CREALOTES(CIXI%)
   '
   Screen.MousePointer = 11
   Dim TLOTE$(), CALOTE#()
   '
   'GENERAR LOS MOVIMIENTOS DE ENTRADA DE LOTES
   Call APERBASDAT("STOCKS")
   sqlx$ = " SELECT * FROM MOVISTO "
   sqlx$ = sqlx$ + "WHERE Cod_Inte = " & CIXI% & " "
   sqlx$ = sqlx$ + "AND CantIngre > 0 "
   sqlx$ = sqlx$ + "AND CODIMOVI <> 'TR' "
   '
   Dim MOVCONTEMP As ADODB.Recordset
   Set MOVCONTEMP = New ADODB.Recordset
15 On Error Resume Next
   MOVCONTEMP.Open FILTSQL$(sqlx$), FLEXCONN, adOpenStatic, adLockPessimistic, adCmdText
   If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
      On Error GoTo 0
      Call CapturaErrorOpen
      GoTo 15
   End If
   On Error GoTo 0
   '
   CAMOVI& = 0
   With MOVCONTEMP
     On Error Resume Next
     .MoveFirst
     If Err < 1 Then
       Do While Not .EOF
         ENTRA# = XVALO(!CantIngre)
         If ENTRA# >= 0.05 Then
           IDENLO$ = TRIM$(!IdenLote): RECAR% = 0
           If IDENLO$ = "" Then IDENLO$ = TRIM$(!DoPriCor): RECAR% = 1
           If IDENLO$ = "" Then IDENLO$ = TRIM$(!DoSecCor): RECAR% = 1
           IDENLO$ = AJUSTI$(IDENLO$, 16)
           If RECAR% = 1 Then
             !IdenLote = TRIM$(IDENLO$)
             .Update
           End If
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
           FEX = CVINT(!Fechmov)
           CANLOT% = CANLOT% + 1
           ReDim Preserve TLOTE$(CANLOT%), CALOTE#(CANLOT%)
           TTXX$ = AJUSTI$(IDENLO$, 40)
           Call REPLA(TTXX$, FECHATEX$(FEX), 19, 8)
           Call REPLA(TTXX$, CSTRING$(MKD$(ENTRA#), 4, 12, 1, 2), 27, 12)
           TLOTE$(CANLOT%) = TTXX$
           CALOTE#(KKI%) = ENTRA#
         End If
19     .MoveNext
       Loop
     End If
   End With
   MOVCONTEMP.Close
   Set MOVCONTEMP = Nothing
   '
   ' VERIFICAR QUE COINCIDEN CON LOS DE TABLA LOTINGRE
   sqlx$ = " SELECT * FROM LOTINGRE "
   sqlx$ = sqlx$ + "WHERE Cod_Inte = " & CIXI% & " "
   'SQLX$ = SQLX$ + " ORDER BY FechAlta ASC "
   'Set LotIngTemp = Stocks.OpenRecordset(SQLX$, dbOpenDynaset)

   Dim LOTINGTEMP As ADODB.Recordset
   Set LOTINGTEMP = New ADODB.Recordset
22 On Error Resume Next
   LOTINGTEMP.Open FILTSQL$(sqlx$), FLEXCONN, adOpenDynamic, adLockPessimistic, adCmdText
   If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
      On Error GoTo 0
      Call CapturaErrorOpen
      GoTo 22
   End If
   On Error GoTo 0
   '
   With LOTINGTEMP
25   On Error Resume Next
     .MoveFirst
     If Err Then
       On Error GoTo 0
       GoTo 28
     End If
     On Error GoTo 0
     Do While Not .EOF
       IDLOTAB$ = AJUSTI$(!IdenLote, 16)
       For KKI% = 1 To CANLOT%
         If Left$(TLOTE$(KKI%), 16) = IDLOTAB$ Then
           FEALTA = CVDAT(FECHANUM(Mid$(TLOTE$(KKI%), 19, 8)))
           CANORIG = CALOTE#(KKI%)
           If FEALTA <> !FechAlta Or Abs(CANORIG - XVALO(!CanAlta)) > 0.05 Then
             .Delete
             GoTo 25
           End If
           Call ACSALOTE(CIXI%, TRIM$(IDLOTAB$), "NOMESS")
           TLOTE$(KKI%) = ""
         End If
       Next KKI%
     .MoveNext
     Loop
     '
28   For KKI% = 1 To CANLOT%
       If TLOTE$(KKI%) <> "" Then
         IDLOTAB$ = TRIM$(Left$(TLOTE$(KKI%), 16))
         FEALTA = CVDAT(FECHANUM(Mid$(TLOTE$(KKI%), 19, 8)))
         CANORIG = CALOTE#(KKI%)
         .AddNew   ' YA CONVERTIDO
         On Error Resume Next
         !CodiEmpr = CodiEmp%
         On Error GoTo 0
         !Cod_Inte = CIXI%
         !IdenLote = TRIM$(IDLOTAB$)
         !FechAlta = FEALTA
         !CanAlta = CANORIG
         !CANSALDO = CANORIG
         .Update
         Call ACSALOTE(CIXI%, TRIM$(IDLOTAB$), "NOMESS")
       End If
     Next KKI%
   End With
   LOTINGTEMP.Close
   Set LOTINGTEMP = Nothing
   '
   Screen.MousePointer = 1
   '
End Sub
'
Sub VERSALIDA(CIXI%)
   '
   Screen.MousePointer = 11
   Call APERBASDAT("STOCKS")
   '
   Dim IDLOTING$(), FELOTING() As Date, CANSALDOX#()
   '
   PORFIFO% = 0
   sqlx$ = " SELECT * FROM LOTINGRE "
   sqlx$ = sqlx$ + "WHERE Cod_Inte = " & CIXI% & " "
   sqlx$ = sqlx$ + "ORDER BY FechAlta ASC ;"
   'Set LotIngTemp = Stocks.OpenRecordset(SQLX$, dbOpenDynaset)
   Dim LOTINGTEMP As ADODB.Recordset
   Set LOTINGTEMP = New ADODB.Recordset
20 On Error Resume Next
   LOTINGTEMP.Open FILTSQL$(sqlx$), FLEXCONN, adOpenKeyset, adLockReadOnly, adCmdText
   If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
      On Error GoTo 0
      Call CapturaErrorOpen
      GoTo 25
   End If
   On Error GoTo 0
   '
   CANLOTES% = 0
   With LOTINGTEMP
     On Error Resume Next
     .MoveFirst
     If Err < 1 Then
       On Error GoTo 0
       Do While Not .EOF
         CANLOTES% = CANLOTES% + 1
         ReDim Preserve IDLOTING$(CANLOTES%), FELOTING(CANLOTES%), CANSALDOX#(CANLOTES%)
         IDLOTING$(CANLOTES%) = TRIM$(!IdenLote)
         FELOTING(CANLOTES%) = !FechAlta
         CANSALDOX#(CANLOTES%) = XVALO(!CanAlta)
       .MoveNext
       Loop
     End If
     On Error GoTo 0
   End With
   LOTINGTEMP.Close
   Set LOTINGTEMP = Nothing
   '
   sqlx$ = " SELECT * FROM MOVISTO "
   sqlx$ = sqlx$ + "WHERE Cod_Inte = " & CIXI% & " "
   sqlx$ = sqlx$ + "AND CantSalid > 0 "
   'SQLX$ = SQLX$ + "ORDER BY FechMov ASC ;"
   '
'AQUI VERIFICA TODOS LOS MOVIMIENTOS DE SALIDA Y AJUSTA QUE EN TODOS LOS
'CASOS LA SUMA DE LOS MOVIMIENTOS DE SALIDA POR LOTES SEA MENOR O IGUAL
'AL TOTAL DEL MOVIMIENTO DE SALIDA.
   'Set MoviTemp = Stocks.OpenRecordset(SQLX$, dbOpenDynaset)
   Dim MoviTemp As ADODB.Recordset
   Set MoviTemp = New ADODB.Recordset
21 On Error Resume Next
   MoviTemp.Open FILTSQL$(sqlx$), FLEXCONN, adOpenDynamic, adLockPessimistic, adCmdText
   If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
      On Error GoTo 0
      Call CapturaErrorOpen
      GoTo 21
   End If
   On Error GoTo 0
   '
   With MoviTemp
     On Error Resume Next
     .MoveFirst
     If Err < 1 Then    ' no hay registros de salida
       '
       Do While Not .EOF
         SALIDPE# = ROUND(!Cantsalid, 1)
         TSALOT$ = ""
         On Error Resume Next
           TSALOT$ = SAFETEXT$(!LotConsum)
         On Error GoTo 0
         If SALIDPE# < 0.05 Then
           TSALOT$ = ""
           '.Edit
           !LotConsum = ""
           .Update
           PORFIFO% = 1
           GoTo 19
         End If
15       CALOCONS# = 0
         For KKL& = 1 To Len(TSALOT$) Step 32
           LOTCO$ = TRIM$(Mid$(TSALOT$, KKL&, 24))
           For KKII% = 1 To CANLOTES%
             If LOTCO$ = IDLOTING$(KKII%) Then GoTo 16
           Next KKII%
           GoTo 18
           '
16         CACONS# = ROUND(XVALO(Mid$(TSALOT$, KKL& + 24, 8)), 1)
           If LOTCO$ = "" Or CACONS# < 0.05 Then
18           TSALOT$ = Left$(TSALOT$, KKL& - 1) + Mid$(TSALOT$, KKL& + 32)
             '.Edit
             !LotConsum = TSALOT$
             .Update
             PORFIFO% = 1
             GoTo 15
           End If
           If CACONS# > SALIDPE# - CALOCONS# Then
             CACONS# = ROUND(SALIDPE# - CALOCONS#, 1)
             If CACONS# > CANSALDOX#(KKII%) Then
               CACONS# = CANSALDOX#(KKII%)
             End If
             If CACONS# < 0.05 Then GoTo 18
             CACON1$ = FORMATNUM(CACONS#, "F8.1")
             Call REPLA(TSALOT$, CACON1$, KKL& + 24, 8)
             '.Edit
             !LotConsum = TSALOT$
             .Update
             PORFIFO% = 1
             GoTo 15
           End If
           CALOCONS# = ROUND(CALOCONS# + CACONS#, 1)
           CANSALDOX#(KKII%) = ROUND(CANSALDOX#(KKII%) - CACONS#, 1)
         Next KKL&
         '
19     .MoveNext
       Loop
     End If
   End With
   '
   '
   '
   With MoviTemp
     On Error Resume Next
     .MoveFirst
     If Err < 1 Then    ' no hay registros de salida
       '
       Do While Not .EOF
         SALIDPE# = ROUND(!Cantsalid, 1)
         TSALOT$ = ""
         On Error Resume Next
           TSALOT$ = SAFETEXT$(!LotConsum)
         On Error GoTo 0
         If SALIDPE# < 0.05 Then GoTo 29
         '
         FEMOVI = Fechmov
25       CALOCONS# = 0
         For KKL& = 1 To Len(TSALOT$) Step 32
           CACONS# = ROUND(XVALO(Mid$(TSALOT$, KKL& + 24, 8)), 1)
           CALOCONS# = CALOCONS# + CACONS#
         Next KKL&
         DIFEPEN# = ROUND(SALIDPE# - CALOCONS#)
         If DIFEPEN# >= 0.05 Then
           For KKII% = 1 To CANLOTES%
             If CANSALDOX#(KKII%) >= 0.05 Then
               If FELOTING(KKII%) >= FEMOVI Or Left$(IDLOTING$(KKII%), 2) = "AS" Then
                 CANLOT# = DIFEPEN#
                 If CANLOT# > CANSALDOX#(KKII%) Then CANLOT# = CANSALDOX#(KKII%)
                 TSALOT$ = TSALOT$ + AJUSTI$(IDLOTING$(KKII%), 24)
                 TSALOT$ = TSALOT$ + FORMATNUM$(CANLOT#, "F8.1")
                 '.Edit
                 !LotConsum = TSALOT$
                 .Update
                 PORFIFO% = 1
                 CANSALDOX#(KKII%) = CANSALDOX#(KKII%) - CANLOT#
                 DIFEPEN# = DIFEPEN# - CANLOT#
                 If DIFEPEN# < 0.05 Then GoTo 29
               End If
             End If
           Next KKII%
         End If
         '
29     .MoveNext
       Loop
     End If
   End With
   '
   With LOTINGTEMP
     On Error Resume Next
     .MoveFirst
     If Err < 1 Then
       On Error GoTo 0
       Do While Not .EOF
         IDLOTE$ = TRIM$(!IdenLote)
         For KKI% = 1 To CANLOTES%
           If TRIM$(IDLOTING$(KKI%)) = IDLOTE$ Then
             If Abs(!CANSALDO - CANSALDOX#(KKI%)) >= 0.05 Then
               '.Edit
               !CANSALDO = CANSALDOX#(KKI%)
               .Update
               PORFIFO% = 1
             End If
             GoTo 39
           End If
         Next KKI%
39     .MoveNext
       Loop
     End If
     On Error GoTo 0
   End With
   MoviTemp.Close
   Set MoviTemp = Nothing
   '
   Screen.MousePointer = 1
   If PORFIFO% > 0 Then
     Call COMUNI("Saldos por Lote Re-Calculados\por Criterio F.I.F.O.")
   End If
   '
End Sub

'
Sub ARMALILOT2(CIXI%)
    '
    'USO ASEMBLI
    Static VERIRECLE%
    '
    If VERIRECLE% < 1 Then
       VERIRECLE% = 1
       If RECLEN%("LILOTSE") < 60 Then
          Call MENSERR(24, "Longitud de Registro Incorrecta\para Controlador 'LILOTSE'.\   \Reinstale el Controlador .RFS")
          Call FINAL("")
       End If
    End If
    '
    Call APERBASDAT("STOCKS")
    sqlx$ = " SELECT * FROM LOTINGRE WITH(NOLOCK)"
    sqlx$ = sqlx$ + "WHERE Cod_Inte = " & CIXI% & " "
    sqlx$ = sqlx$ + "ORDER BY FechAlta ASC, IdenLote ASC "
    Set LOTINGTEMP = READSET(sqlx$)
    '
    REBLA$ = REGBLAN$("LILOTSE")
    JJ& = 0
    With LOTINGTEMP
25     On Error Resume Next
       .MoveFirst
       If Err < 1 Then
         On Error GoTo 0
         Do While Not .EOF
           TTXX$ = REBLA$
             FEX = CVINT(!FechAlta)
           Call REPLA(TTXX$, TRIM$(SAFETEXT$(!IdenLote)), 1, 16)
           Call REPLA(TTXX$, FECHATEX$(FEX), 17, 8)
           Call REPLA(TTXX$, FORMATNUM$(XVALO(!CanAlta), "F12,1"), 25, 12)
             SALFIS = !CanAlta - XVALO(!CanConsu): If SALFIS < 0 Then SALFIS = 0
             SALOGI = !CanAlta - XVALO(!CanConsu) - XVALO(!CanReser): If SALOGI < 0 Then SALOGI = 0
           Call REPLA(TTXX$, FORMATNUM$(SALFIS, "F12,1"), 37, 12)
           Call REPLA(TTXX$, FORMATNUM$(SALOGI, "F12,1"), 49, 12)
           JJ& = JJ& + 1
           Call GRAREG("!LILOTSE", TTXX$, JJ&)
         .MoveNext
         Loop
       End If
    End With
    Call SETULTREG("!LILOTSE", JJ&)
    Call FRESHECHO("!LILOTSE")
    '
End Sub

Sub ARMALILOT(CIXI%, Optional DEPO%, Optional FESTO%, Optional PORANCHO%, Optional SOLOCONSALDO%, Optional ESTADO$)
    '
    SCMA = Screen.MousePointer
    Call APERBASDAT("STOCKS")
    Screen.MousePointer = 11
       '
       ConsEst$ = ""
    If ESTADO$ <> "" Then
       ConsEst$ = " and Libera='" & ESTADO$ & "'"
    End If
    '
    'AVERIGUAR SI TIENEN LOTE (SI NO TIENE GENERA UN NUEVO LOTE PARA TODOS LOS QUE NO TIENEN)
    '****************************************************************************************
    CONDI$ = "(IDENLOTE = '' OR IDENLOTE IS NULL) AND Cod_Inte = " & CIXI%
    If CantRegistros&("MOVISTO", CONDI$, "NOMESS") > 0 Then
      Call RECULOTE(CIXI%)
    End If
    '
    CONDI$ = "Cod_Inte = " & CIXI% & ConsEst$
    If SOLOCONSALDO% = 1 Then
       CONDI$ = CONDI$ & " AND CANALTA > CANCONSU "
    End If
    sqlx$ = " SELECT * FROM LOTINGRE WITH(NOLOCK) "
    sqlx$ = sqlx$ + "WHERE " & CONDI$
    If PORANCHO% = 1 Then
         sqlx$ = sqlx$ + " ORDER BY AnchoMili ASC, FechAlta ASC"
      Else
         sqlx$ = sqlx$ + " ORDER BY FechAlta ASC "
    End If
    '
    FIN& = CantRegistros("LOTINGRE", CONDI$)
    '
    Dim LOTINGTEMP As ADODB.Recordset
    Set LOTINGTEMP = New ADODB.Recordset
    Set LOTINGTEMP = FLEXCONN.Execute(FILTSQL$(sqlx$))
    '
    REBLA$ = REGBLAN$("LILOTSE")
    JJ& = 0
    With LOTINGTEMP
25     On Error Resume Next
       .MoveFirst
       If Err < 1 Then
         On Error GoTo 0
         Do While Not .EOF
           TTXX$ = REBLA$
           FEX = CVINT(!FechAlta)
           IDLO$ = SAFETEXT$(!IdenLote)
           ANCHX = XVALO(!AnchoMili)
           FEVE% = CVINT(!FechVenc)
           SLOTE = SaldoLote(CIXI%, IDLO$, DEPO%, FESTO%)
           If SOLOCONSALDO% = 1 Then
              If SLOTE < 0.05 Then
                 GoTo 29
              End If
           End If
           '
           Call REPLA(TTXX$, TRIM$(IDLO$), 1, 16)
           Call REPLA(TTXX$, FECHATEX$(FEX), 17, 8)
           Call REPLA(TTXX$, FORMATNUM$(!CanAlta, "F10." & CNTDC$(CIXI%)), 25, 10)
           Call REPLA(TTXX$, FORMATNUM$(SLOTE, "F10." & CNTDC$(CIXI%)), 35, 10)
           Call REPLA(TTXX$, FECHATEX$(FEVE%), 47, 8)
              ANCHA$ = FORMATNUM$(ANCHX, "F8.0")
              If TRIM$(ANCHA$) <> "" Then ANCHA$ = ANCHA$ + " mm"
           Call REPLA(TTXX$, AJUSTD$(ANCHA$, 11), 55, 11)
               LARGX = 0: If ANCHX > 0 Then LARGX = 1000 * SLOTE / ANCHX
               LARGA$ = FORMATNUM$(LARGX, "F8.0")
               If TRIM$(LARGA$) <> "" Then LARGA$ = LARGA$ + " Mt"
           Call REPLA(TTXX$, AJUSTD$(LARGA$, 11), 66, 11)
           '
           JJ& = JJ& + 1
           Call GRAREG("!LILOTSE", TTXX$, JJ&)
           Call TRACE(20, JJ&, FIN&)
29       .MoveNext
         Loop
       End If
    End With
    LOTINGTEMP.Close
    Set LOTINGTEMP = Nothing
    Call SETULTREG("!LILOTSE", JJ&)
    Call FRESHECHO("!LILOTSE")
    Screen.MousePointer = SCMA
    '
End Sub
'
Sub ARMALISLOTES(CIXI%, Optional DEPO%, Optional FESTO%, Optional PORANCHO%, Optional SOLOCONSALDO%)
    '
    Call APERBASDAT("STOCKS")
    sqlx$ = " SELECT * FROM LOTINGRE "
    sqlx$ = sqlx$ + "WHERE Cod_Inte = " & CIXI%
    sqlx$ = sqlx$ + " ORDER BY FechAlta ASC "
    '
    Dim LOTINGTEMP As ADODB.Recordset
    Set LOTINGTEMP = New ADODB.Recordset
    Set LOTINGTEMP = FLEXCONN.Execute(FILTSQL$(sqlx$))

    REBLA$ = REGBLAN$("LILOTSE")
    JJ& = 0
    With LOTINGTEMP
25     If Not (.BOF And .EOF) Then
         Do While Not .EOF
           TTXX$ = REBLA$
           FEX = CVINT(!FechAlta)
           IDLO$ = SAFETEXT$(!IdenLote)
           FEVE% = CVINT(!FechVenc)
           SLOTE = SaldoLote(CIXI%, IDLO$, DEPO%, FESTO%)
           If SOLOCONSALDO% = 1 Then
              If SLOTE < 0.05 Then
                 GoTo 29
              End If
           End If
           '
           CANRES = XVALO(!CanReser)
           SALOGIC = SLOTE - CANRES
           If SALOGIC < 0 Then SALOGIC = 0
           Call REPLA(TTXX$, TRIM$(IDLO$), 1, 16)
           Call REPLA(TTXX$, FECHATEX$(FEX), 17, 8)
           Call REPLA(TTXX$, FORMATNUM$(!CanAlta, "F10.1"), 25, 10)
           Call REPLA(TTXX$, FORMATNUM$(SLOTE, "F10.1"), 35, 10)
           Call REPLA(TTXX$, FECHATEX$(FEVE%), 47, 8)
           If DEPO% < 1 Then
              Call REPLA(TTXX$, FORMATNUM$(CANRES, "F10.1"), 55, 10)
              Call REPLA(TTXX$, FORMATNUM$(SALOGIC, "F10.1"), 65, 10)
           End If
           '
           JJ& = JJ& + 1
           Call GRAREG("!LILOTSE", TTXX$, JJ&)
29       .MoveNext
         Loop
       End If
    End With
    LOTINGTEMP.Close
    Set LOTINGTEMP = Nothing
    Call SETULTREG("!LILOTSE", JJ&)
    Call FRESHECHO("!LILOTSE")
    '
End Sub
Sub ARMALISLOTES_SIN_SALDOLOTE(CIXI%, Optional DEPO%, Optional FESTO%, Optional PORANCHO%, Optional SOLOCONSALDO%)
    '
    Call APERBASDAT("STOCKS")
    sqlx$ = " SELECT * FROM LOTINGRE "
    sqlx$ = sqlx$ + "WHERE Cod_Inte = " & CIXI%
    sqlx$ = sqlx$ + " ORDER BY FechAlta ASC "
    '
    Dim LOTINGTEMP As ADODB.Recordset
    Set LOTINGTEMP = New ADODB.Recordset
    Set LOTINGTEMP = FLEXCONN.Execute(FILTSQL$(sqlx$))

    REBLA$ = REGBLAN$("LILOTSE")
    JJ& = 0
    With LOTINGTEMP
25     If Not (.BOF And .EOF) Then
         Do While Not .EOF
           TTXX$ = REBLA$
           FEX = CVINT(!FechAlta)
           IDLO$ = SAFETEXT$(!IdenLote)
           FEVE% = CVINT(!FechVenc)
'           SLOTE = SaldoLote(CIXI%, IDLO$, DEPO%, FESTO%)
           SLOTE = XVALO(!CanAlta) - XVALO(!CanConsu)
           If SOLOCONSALDO% = 1 Then
              If SLOTE < 0.05 Then
                 GoTo 29
              End If
           End If
           '
           CANRES = XVALO(!CanReser)
           SALOGIC = SLOTE - CANRES
           If SALOGIC < 0 Then SALOGIC = 0
           Call REPLA(TTXX$, TRIM$(IDLO$), 1, 16)
           Call REPLA(TTXX$, FECHATEX$(FEX), 17, 8)
           Call REPLA(TTXX$, FORMATNUM$(!CanAlta, "F10.1"), 25, 10)
           Call REPLA(TTXX$, FORMATNUM$(SLOTE, "F10.1"), 35, 10)
           Call REPLA(TTXX$, FECHATEX$(FEVE%), 47, 8)
           If DEPO% < 1 Then
              Call REPLA(TTXX$, FORMATNUM$(CANRES, "F10.1"), 55, 10)
              Call REPLA(TTXX$, FORMATNUM$(SALOGIC, "F10.1"), 65, 10)
           End If
           '
           JJ& = JJ& + 1
           Call GRAREG("!LILOTSE", TTXX$, JJ&)
29       .MoveNext
         Loop
       End If
    End With
    LOTINGTEMP.Close
    Set LOTINGTEMP = Nothing
    Call SETULTREG("!LILOTSE", JJ&)
    Call FRESHECHO("!LILOTSE")
    '
End Sub


Sub LISTMATPROV(NC%, DepOrig%, NROOCOM$)
        
        Screen.MousePointer = 11
        HAYOC% = 0
        Call APERBASDAT("STOCKS")
        Call SETULTREG("!BOLREDET", 0)
        '
        'CONSULTA DE AGRUPAMIENTO PARA QUE PRESENTE DE A UN CODIGO
        sqlx$ = "SELECT Cod_Inte, SUM(Cant_Asig) AS SxCant_Asig, SUM(Cant_Entreg) AS SxCant_Entreg , SUM(Cant_Devu) AS SxCant_Devu "
        sqlx$ = sqlx$ + " FROM MACONSIG WHERE Npro_Ocom = " & NC% & " "
        ' SQLX$ = SQLX$ + " AND (Cant_Asig > Cant_Entreg or Cant_Devu <> 0 )"
        sqlx$ = sqlx$ + " GROUP BY Cod_Inte ORDER BY Cod_Inte ASC"
        
        Dim MACONSITMP As ADODB.Recordset
        Set MACONSITMP = FLEXCONN.Execute(FILTSQL$(sqlx$))
        With MACONSITMP
          Do While .EOF = False
           'CANTIDAD PENDIENTE DE ENTREGA EN GENERAL
           CODIGO$ = CODEXT$(XVALO(!Cod_Inte))
           CAPENE# = XVALO(!SxCant_Asig) - (XVALO(!SxCant_Entreg) - XVALO(!SxCant_Devu))
           
           If CAPENE# < 0 Then CAPENE# = 0
               CI% = XVALO(!Cod_Inte)
               If CAPENE# > FRACMIN(CI%) Then
                    'AJUSTA CANTIDAD RESPETANDO EL LOTE INDIVISIBLE
                    If LINDIVIS(CI%) > FRACMIN(CI%) Then
                       'VER SI SE PRESENTA MENSAJE DE AVISO DE AJUSTE SEGUN LOTE INDIVISIBLE
                       CAPENE# = ROUND#(REDON_INDIVIS(CI%, CAPENE#), XVALO(CNTDC$))
                    End If
                    'AJUSTA CANTIDAD SEGUN EL STOCK DISPONIBLE EN EL DEPOSITO CONFIGURADO
                    STODISPO# = CDbl(STODEPOS(CI%, DepOrig%))
                    If STODISPO# < 0 Then STODISPO# = 0
                    If STODISPO# < CAPENE# Then CAPENE# = STODISPO#
                    '
                    Y$ = REGBLAN$("!BOLREDET")
                    Call REPLA(Y$, MKI$(CI%), 83, 2)
                    Call REPLA(Y$, UNIMED$(CI%), 85, 4)
                    Call REPLA(Y$, MKD$(CAPENE#), 89, 8)
                    Call REPLA(Y$, MKS$(0), 151, 4) ' SIN ORDEN DE COMPRA POR AGRUPAMIENTO
                    Call REPLA(Y$, MKS$(0), 189, 4)
                    Call REPLA(Y$, "", 119, 16) ' SIN ORDEN DE COMPRA POR AGRUPAMIENTO
                    HAYOC% = 1
                    Call REGAPP("!BOLREDET", Y$)
                End If
           .MoveNext
          Loop
        End With
        MACONSITMP.Close
        Set MACONSITMP = Nothing
        '
        
555     If HAYOC% = 0 Then
            Screen.MousePointer = 1
            Call COMUNI("No hay Consignaciones Pendientes\para este Proveedor.")
        End If
        '
End Sub
'
Sub REMICONSARMA(NPROVE%, REMI$, DepOrig%, NROOCOM$)
          '
Call REMICONSARMA_NUE(NPROVE%, REMI$, DepOrig%)
Call REORGAMATCON(NPROVE%)
Exit Sub
          'GENERA REMITO EN CONSIGANACION
          NC1% = (-1) * NPROVE%
          '
          CMO$ = TRIM$(Control$("", "CODRECON"))
          If CMO$ = "" Then CMO$ = "RC"

          FF% = HOY(HO$)
          For YX& = 1 To ULTREG&("!BOLREDET")
             YY$ = REGLEIDO$("!BOLREDET", YX&)
             CI% = CVI(Mid$(YY$, 83, 2))
             CAN# = CVD(Mid$(YY$, 89, 8))
             If CI% > 0 Then
               If CI% <= NUMAS% Then
                 If CAN# > 0.005 Then
                    NUOC& = CVS(Mid$(YY$, 151, 4))
                    DOSEC$ = ""
 
                    Dim MaConsig As ADODB.Recordset
                    Set MaConsig = New ADODB.Recordset
                    sqlx$ = " select * from MACONSIG"
                    sqlx$ = sqlx$ + " WHERE Cant_Asig > Cant_Entreg "
                    sqlx$ = sqlx$ + " AND Cant_Asig >= 0.05 "
                    sqlx$ = sqlx$ + " AND Cod_Inte = " & CI% & ""
                    sqlx$ = sqlx$ + " AND Npro_Ocom = " & NPROVE% & " "
25                  On Error Resume Next
                    MaConsig.Open FILTSQL$(sqlx$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
                    If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
                       On Error GoTo 0
                       Call CapturaErrorOpen
                       GoTo 25
                    End If
                    On Error GoTo 0
                    '
                    CANTIASIG_ENT = 0
                    NUITEM = 0
                    SALPENDIAREMITIR = CAN#
                    CANTIASIGPENDIENTE = 0
                    DIFERENCASO1 = 0
                    With MaConsig
                      'ASIGNA ENTREGA SEGUN CANTIDAD ASIGNADA
                      Do While Not .EOF
                         If SALPENDIAREMITIR < 0.05 Then
                            Exit Do
                         End If
                         
                           DOSEC$ = SAFETEXT$(!Nro_Ocom)
                           CANASIG = XVALO(!cant_asig)
                           CANENTRE = XVALO(!cant_entreg)
                           CANTIASIG_ENT = CANASIG - CANENTRE
                           
                           If SALPENDIAREMITIR < CANTIASIG_ENT Then CANTIASIG_ENT = SALPENDIAREMITIR
                               
                           'CASO 1 -- SI TIENE YA CANTIDAD ENTREGADA IGUALO CON LA CANTIDAD ASIGNADA
                           ' y GENERO UNO NUEVO
                           If CANENTRE > 0.05 Then
                              !cant_asig = CANENTRE
                              .Update
                              GoTo 60 ' NUEVO REGISTRO POR EL SALDO PENDIENTE A REMITIR
                           End If
                           '
                           'CASO 2 -- QUE EL SALDO A REMTIR SEA MENOR AL ASIGNADO
                           If SALPENDIAREMITIR < CANASIG Then
                              !cant_asig = CANTIASIG_ENT
                              CANTIASIGPENDIENTE = CANASIG - SALPENDIAREMITIR
                           End If
                           '
                           'ASIGNO CANTIDAD ENTREGADA
                           !cant_entreg = CANTIASIG_ENT
                           !NREMI_CONSIG = REMI$
                           If FF% > CVINT(!FECHA_ENTREG) Then
                             !FECHA_ENTREG = CVDAT(FF%)
                           End If
                           .Update
                           '
                           'GRABO EN MOVISTO
                           CANTI = (-1) * CANTIASIG_ENT
                           REFE$ = REMI$
                           NUITEM = NUITEM + 1
                           NUORIT% = NUITEM
                           Call MOVISTOK(FF%, CI%, LOTE$, CMO$, REMI$, REMI$, DOSEC$, DOSEC$, REFE$, NUORIT%, 0, 0, NC1%, DepOrig%, CANTI)
                           '
                           'ACTUALIZO EL SALDO A REMITIR
                           SALPENDIAREMITIR = SALPENDIAREMITIR - CANTIASIG_ENT
                           '
                           'SI CANTIDAD ASIGNADA ES MAYOR A LA ENTREGADA
                           If CANTIASIGPENDIENTE >= 0.05 Then GoTo 60 'GENERA UNO NUEVO
                           '
                           GoTo 80 ' PASO AL SIGUIENTE

60                         .AddNew
                           On Error Resume Next
                           !CodiEmpr = CodiEmp%
                           On Error GoTo 0
                           !Femi_Ocom = CVDAT(FF%)
                           !Npro_Ocom = NPROVE%
                           !Raso_Prov = RASOCLI(NC1%)
                           !Nro_Ocom = DOSEC$
                           !Cod_Inte = CI%
                           !Cod_Exte = CODEXT(CI%)
                           !NREMI_CONSIG = REMI$
                           !Unidad = UNIMED$(CI%)
                           !cant_asig = CANTIASIG_ENT
                           !cant_entreg = CANTIASIG_ENT
                           !cant_rendid = 0
                           !cant_devu = 0
                           !FECHA_ENTREG = CVDAT(FF%)
                           !Descritem = DESCRIT0$(CI%)

                           '
                           'CASO QUE EL SALDO A REMTIR SEA < A LA CANTIDAD CONSIGNADA
                           If CANTIASIGPENDIENTE > 0 Then
                               CANTIASIG_ENT = CANTIASIGPENDIENTE
                               !NREMI_CONSIG = ""
                               !cant_asig = CANTIASIG_ENT
                               !cant_entreg = 0
                               !FECHA_ENTREG = CVDAT(0)
                               .Update
                               CANTIASIGPENDIENTE = 0
                               GoTo 80
                           End If
                           '
                           'CASO QUE HUBO ENTREGADO AGREGA LA DIFERENCIA
                           If DIFERENCASO1 > 0 Then
                               !NREMI_CONSIG = ""
                               !cant_asig = DIFERENCASO1
                               !cant_entreg = 0
                               !FECHA_ENTREG = CVDAT(0)
                               .Update
                               DIFERENCASO1 = 0
                               GoTo 80
                           End If
                           .Update
                           '
                           'GRABO MOVIMIENTO EN MOVISTO
                           CANTI = (-1) * CANTIASIG_ENT
                           REFE$ = REMI$
                           NUITEM = NUITEM + 1
                           NUORIT% = NUITEM
                           Call MOVISTOK(FF%, CI%, LOTE$, CMO$, REMI$, REMI$, DOSEC$, DOSEC$, REFE$, NUORIT%, 0, 0, NC1%, DepOrig%, CANTI)
                           '
                           'ACTUALLIZO VARIABLES
                           SALPENDIAREMITIR = SALPENDIAREMITIR - CANTIASIG_ENT
                           CANTIASIGPENDIENTE = 0
                           DIFERENCASO1 = CANASIG - (CANENTRE + CAN#)
80                      .MoveNext
                      Loop
                      '
                      ' SI RECORRIO EL ARTICULO Y TODAVIA TIENE SALDO PENDIENTE A REMITIR
                      ' CREA UN  REGISTRO NUEVO POR EL SALDO
                      If SALPENDIAREMITIR > 0.05 Then
                         CANTIASIG_ENT = SALPENDIAREMITIR
                         DOSEC$ = ""
                         GoTo 60
                      End If
                      '
                      If DIFERENCASO1 > 0.05 Then
                        CANTIASIG_ENT = DIFERENCASO1
                        GoTo 60
                      End If
                      
                    End With
                 End If
               End If
             End If
             '
          Next YX&
          MaConsig.Close
          Set MaConsig = Nothing
          Call CIERRARCH("MACONSIG")
          '
          Call FRESHECHO("IREPRIPE")
          Screen.MousePointer = 1
          '
End Sub
Sub REMICONSARMA_NUE(NPROVE%, REMI$, DepOrig%)

          'GENERA REMITO EN CONSIGNACION
          NC1% = (-1) * NPROVE%
          CMO$ = TRIM$(Control$("", "CODRECON"))
          If CMO$ = "" Then CMO$ = "RC"
          FF% = HOY(HO$)
          '
          Call PoneEnCeroNull("MACONSIG", "CANT_DEVU", REGAF&)
          '
          'RECORRO LISTA DE MATERIALES A ENTREGAR
          For YX& = 1 To ULTREG&("!BOLREDET")
             YY$ = REGLEIDO$("!BOLREDET", YX&)
             CI% = CVI(Mid$(YY$, 83, 2))
             CAN# = CVD(Mid$(YY$, 89, 8))
             If CI% > 0 Then
               If CI% <= NUMAS% Then
                 If CAN# >= 0.05 Then
                    'OBTENGO LISTA DE CONSIGNACIONES QUE TIENEN SALDO PENDIENTE DE ENTREGA
                    Dim MaConsig As ADODB.Recordset
                    Set MaConsig = New ADODB.Recordset
                    sqlx$ = " select * from MACONSIG"
                    sqlx$ = sqlx$ + " WHERE Cant_Asig > (Cant_Entreg - Cant_Devu) "
                    sqlx$ = sqlx$ + " AND Cant_Asig >= 0.05 "
                    sqlx$ = sqlx$ + " AND Cod_Inte = " & CI% & ""
                    sqlx$ = sqlx$ + " AND Npro_Ocom = " & NPROVE% & " "
25                  On Error Resume Next
                    MaConsig.Open FILTSQL$(sqlx$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
                    If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
                       On Error GoTo 0
                       Call CapturaErrorOpen
                       GoTo 25
                    End If
                    On Error GoTo 0
                    '
                    CANTENTREGA = 0
                    NUITEM = 0
                    SALPENDIAREMITIR = CAN#
                    With MaConsig
                      'ASIGNA ENTREGA SEGUN CANTIDAD ASIGNADA
                      Do While Not .EOF
                         '
                         CANASIG = XVALO(!cant_asig)
                         CANENTRE = XVALO(!cant_entreg)
                         CANDEVU = XVALO(!cant_devu)
                         CANTENTREGA = CANASIG - (CANENTRE - CANDEVU)
                           
                         If CANTENTREGA < 0.05 Then GoTo 80 ' SI SE ENTREGO MAS DE LO ASIGNADO CONTINUA CON EL PROXIMO
                         If SALPENDIAREMITIR < CANTENTREGA Then CANTENTREGA = SALPENDIAREMITIR
                               
                         !cant_entreg = CANENTRE + CANTENTREGA
                         !NREMI_CONSIG = REMI$
                         .Update
                         '
                         'GRABA MOVIMIENTO DE STOCK DE SALIDA DEL MATERIAL
                         CANTI = (-1) * CANTENTREGA
                         REFE$ = REMI$
                         NUITEM = NUITEM + 1
                         NUORIT% = NUITEM
                         Call MOVISTOK(FF%, CI%, LOTE$, CMO$, REMI$, REMI$, DOSEC$, DOSEC$, REFE$, NUORIT%, 0, 0, NC1%, DepOrig%, CANTI)
                         '
                         'ACTUALIZO EL SALDO A REMITIR
                         SALPENDIAREMITIR = SALPENDIAREMITIR - CANTENTREGA
                         '
80                      .MoveNext
                      Loop
                      '
                      ' SI RECORRIO EL ARTICULO Y TODAVIA TIENE SALDO PENDIENTE A REMITIR
                      ' CREA UN  REGISTRO NUEVO POR EL SALDO
                      If SALPENDIAREMITIR > 0.05 Then
                           .AddNew
                           On Error Resume Next
                           !CodiEmpr = CodiEmp%
                           On Error GoTo 0
                           !Femi_Ocom = CVDAT(FF%)
                           !Npro_Ocom = NPROVE%
                           !Raso_Prov = RASOCLI(NC1%)
                           !Nro_Ocom = ""
                           !Cod_Inte = CI%
                           !Cod_Exte = CODEXT(CI%)
                           !NREMI_CONSIG = REMI$
                           !Unidad = UNIMED$(CI%)
                           !cant_asig = 0
                           !cant_entreg = SALPENDIAREMITIR
                           !cant_rendid = 0
                           !cant_devu = 0
                           !FECHA_ENTREG = CVDAT(FF%)
                           !Descritem = DESCRIT0$(CI%)
                           .Update
                           '
                           'GRABO MOVIMIENTO EN MOVISTO
                           CANTI = (-1) * SALPENDIAREMITIR
                           REFE$ = REMI$
                           NUITEM = NUITEM + 1
                           NUORIT% = NUITEM
                           Call MOVISTOK(FF%, CI%, LOTE$, CMO$, REMI$, REMI$, DOSEC$, DOSEC$, REFE$, NUORIT%, 0, 0, NC1%, DepOrig%, CANTI)
                           '
                      End If
                      'ACTUALLIZO VARIABLES
                      SALPENDIAREMITIR = 0
                      NUITEM = 0
                      '
                    End With
                 End If
               End If
             End If
             '
          Next YX&
          MaConsig.Close
          Set MaConsig = Nothing
          Screen.MousePointer = 1
          '
End Sub

Sub REORGAMATCON(NPROVE%)

    Call ABRECONEXION
    Call COPYSTRU("BOLREDET", "REOMACO")
        
    sqlx$ = "SELECT * FROM MACONSIG WITH(NOLOCK) "
    sqlx$ = sqlx$ + " WHERE Cant_Asig > (Cant_Entreg - Cant_Devu) "
    sqlx$ = sqlx$ + " AND Cant_Asig >= 0.05 "
    sqlx$ = sqlx$ + " AND Npro_Ocom = " & Str(NPROVE%) & " "
    Set MaConsig1 = READSET(sqlx$)
    
    I& = 0
    Call SETULTREG("!REOMACO", 0)
    With MaConsig1
        Do While Not .EOF
            CI% = XVALO(!Cod_Inte)
            NROOCOM$ = SAFETEXT$(!Nro_Ocom)
            CASIG = XVALO(!cant_asig)
            CENTRE = XVALO(!cant_entreg)
            CDEVU = XVALO(!cant_devu)
            SALDOACONSIG# = CASIG - (CENTRE - CDEVU)
            '
            If SALDOACONSIG# > 0 Then
                I& = I& + 1
                XX$ = REGBLAN$("REOMACO")
                Call REPLA(XX$, NROOCOM$, 69, 12)
                Call REPLA(XX$, MKI$(CI%), 83, 2)
                Call REPLA(XX$, MKD$(SALDOACONSIG#), 163, 8)
                Call REPLA(XX$, MKD$(0), 171, 8)
                Call GRAREG("!REOMACO", XX$, I&)
            End If
            .MoveNext
        Loop
    End With
    MaConsig1.Close
    Call SETULTREG("!REOMACO", I&)

    For JJ& = 1 To ULTREG&("!REOMACO")
        XX$ = REGLEIDO$("!REOMACO", JJ&)
        CI% = CVI(Mid$(XX$, 83, 2))
        COD$ = CODEXT$(CI%)
        SALDOACONSIG# = CVD(Mid$(XX$, 163, 8))
        TOTALDISTRI# = 0
        '
        Dim MATCONSINOC As ADODB.Recordset
        Set MATCONSINOC = New ADODB.Recordset
        sqlx1$ = "SELECT * FROM MACONSIG "
        sqlx1$ = sqlx1$ + " WHERE Cant_Entreg > (Cant_Rendid + Cant_Devu) "
        sqlx1$ = sqlx1$ + " AND Cant_Entreg >= 0.05 "
        sqlx1$ = sqlx1$ + " AND Cod_Inte=" & CI%
        sqlx1$ = sqlx1$ + " AND Npro_Ocom = " & NPROVE% & " "
        sqlx1$ = sqlx1$ + " AND (Nro_Ocom= '' OR Nro_Ocom like 'OC-000000%' OR Nro_Ocom is null)"
        '
25      On Error Resume Next
        MATCONSINOC.Open FILTSQL$(sqlx1$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
        If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
           On Error GoTo 0
           Call CapturaErrorOpen
           GoTo 25
        End If
        On Error GoTo 0
        '
        With MATCONSINOC
            Do While Not .EOF
                If SALDOACONSIG# > 0.05 Then
                    MATDISPO = XVALO(!cant_entreg) - (XVALO(!cant_rendid) + XVALO(!cant_devu))
                    If MATDISPO > 0 Then
                        MATDISTRI = SALDOACONSIG#
                        If MATDISTRI > MATDISPO Then MATDISTRI = MATDISPO
                        CDEVU = XVALO(!cant_devu)
                        !cant_devu = CDEVU + MATDISTRI
                        .Update
                        SALDOACONSIG# = SALDOACONSIG# - MATDISTRI
                        TOTALDISTRI# = TOTALDISTRI# + MATDISTRI
                    End If
                Else
                    Exit Do
                End If
                .MoveNext
            Loop
        End With
        MATCONSINOC.Close
        '
        Call REPLA(XX$, MKD$(TOTALDISTRI#), 171, 8)
        Call GRAREG("!REOMACO", XX$, JJ&)
    Next JJ&
    
    For JJ& = 1 To ULTREG&("!REOMACO")
        XX$ = REGLEIDO$("!REOMACO", JJ&)
        NROOCOM$ = Mid$(XX$, 69, 12)
        CI% = CVI(Mid$(XX$, 83, 2))
        TOTALDISTRI# = CVD(Mid$(XX$, 171, 8))
        '
        If TOTALDISTRI# > 0.05 Then
            CONDI$ = "COD_INTE=" & CI% & " AND NPRO_OCOM=" & NPROVE% & " AND NRO_OCOM='" & TRIM$(NROOCOM$) & "'"
            CENTRE = XVALO(VALOBADA("MACONSIG", "CANT_ENTREG", CONDI$))
            TOTALDISTRI# = TOTALDISTRI# + CENTRE
            Call ACTUREGISTRO("MACONSIG", "CANT_ENTREG", CONDI$, TOTALDISTRI#, REGAF&)
        End If
    Next JJ&
    
End Sub

'
Function SaldoFinal(CIXI%, NDEP%)
'select sum(cantingre)as AcuIngre, sum(cantsalid) as AcuEgre  from movisto
   'Where cod_inte = 140 And depomovi = 1
   'Insert Into SalStocks

Dim MIRS As ADODB.Recordset
Dim AcuIngre As Single
Dim AcuEgre As Single
STRSQL = "Select sum(cantingre)as AcuIngre, sum(cantsalid) as AcuEgre  from movisto"
STRSQL = STRSQL + " Where cod_inte = " & CIXI% & " And depomovi = " & NDEP%
Set MIRS = FLEXCONN.Execute(FILTSQL$(STRSQL))
If Not (MIRS.EOF And MIRS.BOF) Then
   AcuIngre = MIRS!AcuIngre
   AcuEgre = MIRS!AcuEgre
End If
SaldoFinal = AcuIngre - AcuEgre
End Function
'
Sub ACSALOT(CIXI%, IDLO$, SUINGRE#, SUCONSU#, SURESER#, CSALDO#, NPROVEI%, Optional MODOX$) ' ACTUALIZA SALDOS POR LOTE
    '
    CAMOVILO% = 0
    SUINGRE# = 0
    SUCONSU# = 0
    SURESER# = 0
    CSALDO# = 0
    NPROVEI% = 0
    '
    IDLOTX$ = TRIM$(IDLO$)
    If IDLOTX$ = "" Then
       Exit Sub
    End If
    '
    sqlx$ = "SELECT * FROM LOTINGRE "
    sqlx$ = sqlx$ + "WHERE IdenLote = '" & IDLOTX$ & "'"
    sqlx$ = sqlx$ + "AND Cod_Inte = " & CIXI%
    '
    Dim LOTINSEL As ADODB.Recordset
    Set LOTINSEL = New ADODB.Recordset
15  On Error Resume Next
    LOTINSEL.Open FILTSQL$(sqlx$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
    If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
      On Error GoTo 0
      Call CapturaErrorOpen
      GoTo 15
   End If
   On Error GoTo 0
   '
   With LOTINSEL
       On Error Resume Next
       .MoveFirst
       If Err Then
          On Error GoTo 0
          Call MENSERR(24, "Lote '" + IDLOTX$ + "' no Encontrado.\Probable Falta Trazar Ingreso.")
          Call CONECRUN("TRAZAB07", "Registro de Trazabilidad de Componentes")
          Exit Sub
       End If
       On Error GoTo 0
   End With
    '
20 With LOTINSEL
       CSALDO# = SUINGRE# - SUCONSU# - SURESER#
       If CSALDO# < (-0.05) Then
          Call MENSERR(24, "Exceso de Consumo en Código '" + CODEXT$(CIXI%) + "' - Lote '" + IDLOTX$ + "'\Saldo Ajustado a Cero.")
          CSALDO# = 0
          GoTo 25
       End If
       '
       If IsNull(!CanConsu) Then GoTo 25
       If IsNull(!CanReser) Then GoTo 25
       
       If Abs(SUINGRE# - XVALO(!CanAlta)) >= 0.05 Or Abs(CSALDO# - XVALO(!CANSALDO)) >= 0.05 Or Abs(SUCONSU# - XVALO(!CanConsu)) >= 0.05 Or Abs(SURESER# - XVALO(!CanReser)) >= 0.05 Then
          If UCase$(MODOX$) <> "NOMESS" Then
             Call MENSERR(24, "Saldos en Código '" + CODEXT$(CIXI%) + "' - Lote '" + IDLOTX$ + "'\Ajustado Según Ingresos, Reservas y Consumos.")
          End If
25        !CanAlta = SUINGRE#
          !CanConsu = SUCONSU#
          !CanReser = SURESER#
          !CANSALDO = CSALDO#
          .Update
       End If
   End With
   LOTINSEL.Close
   Set LOTINSEL = Nothing
    '
End Sub

Function SaldoLote(CIXI%, IDLO$, Optional NDEP%, Optional FEFIN%)
    '
    CADECI% = XVALO(CNTDC$)
    SL = 0
    FLEXCONN.Execute (FILTSQL$("SET DATEFORMAT dmy"))
    '
    If CIXI% > 0 Then             ' SOLO SI HAY UN CODIGO ACTIVO - 27/02/2010 - epb
       Dim MIRS As ADODB.Recordset
       Dim AcuMovi As Single
       STRSQL$ = "Select SUM(CANTINGRE) AS ACUIN, SUM(CANTSALID) AS ACUSA FROM MOVISTO WITH(NOLOCK)"
       STRSQL$ = STRSQL$ + " WHERE Cod_Inte = " & CIXI%
       If TRIM$(IDLO$) <> "" Then STRSQL$ = STRSQL$ + " AND IdenLote = '" & IDLO$ & "'"
       If NDEP% > 0 Then STRSQL$ = STRSQL$ + " AND DepoMovi = " & NDEP%
       If FEFIN% > 33 Then STRSQL$ = STRSQL$ + " AND FechMov <= '" & FETEXCOR1$(FEFIN%) & " '"
       Set MIRS = FLEXCONN.Execute(FILTSQL$(STRSQL))
       If Not (MIRS.EOF And MIRS.BOF) Then
          SL = ROUND(XVALO(MIRS!ACUIN) - XVALO(MIRS!ACUSA), CADECI%)
       End If
    End If
    SaldoLote = SL
    '
End Function

Function DEPOLOT%(CIXI%, IDLO$, Optional FEFIN%)
    '
    DL% = 0
    If CIXI% < 1 Then GoTo 90
    '
    Dim MIRS As ADODB.Recordset
    Dim AcuMovi As Single
    STRSQL$ = "Select SUM(CANTINGRE-CANTSALID) AS AcuMovi, DepoMovi FROM MOVISTO WITH(NOLOCK)"
    STRSQL$ = STRSQL$ + " WHERE Cod_Inte = " & CIXI%
    If TRIM$(IDLO$) <> "" Then STRSQL$ = STRSQL$ + " AND IdenLote = '" & IDLO$ & "'"
    If FEFIN% > 33 Then STRSQL$ = STRSQL$ + " AND FechMov <= '" & FETEXCOR1$(FEFIN%) & " '"
    STRSQL$ = STRSQL$ + " GROUP BY DepoMovi"
    Set MIRS = FLEXCONN.Execute(FILTSQL$(STRSQL))
    '
    DL% = 0
    If Not (MIRS.EOF And MIRS.BOF) Then
      With MIRS
        Do While Not .EOF
          If !AcuMovi >= 0.05 Then
            DL% = XVALO(!DepoMovi)
            GoTo 90
          End If
        .MoveNext
        Loop
      End With
    End If
    '
90  DEPOLOT% = DL%
    '
End Function
'
Sub CARDETMOVLOT(CIXI%, IDLOTX$)
    '
    CAMOVILO% = 0
    If CIXI% < 1 Or TRIM$(IDLOTX$) = "" Then Exit Sub
    '
    Call APERBASDAT("STOCKS")
    SQLZ$ = " SELECT * FROM MOVISTO WITH(NOLOCK) "
    SQLZ$ = SQLZ$ + "WHERE Cod_Inte = " & CIXI% & " "
    SQLZ$ = SQLZ$ + "AND IdenLote = '" & IDLOTX$ & "' "
    SQLZ$ = SQLZ$ + "ORDER BY FechMov ASC"
    '
    Dim MOVCONTEMP As ADODB.Recordset
    Set MOVCONTEMP = New ADODB.Recordset
25  On Error Resume Next
    MOVCONTEMP.Open FILTSQL$(SQLZ$), FLEXCONN, adOpenKeyset, adLockReadOnly, adCmdText
    If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
      On Error GoTo 0
      Call CapturaErrorOpen
      GoTo 25
    End If
    On Error GoTo 0
    '
    SUINGRE# = 0: SUCONSU# = 0
    With MOVCONTEMP
      On Error Resume Next
      .MoveFirst
      If Err < 1 Then
         On Error GoTo 0
         Do While Not .EOF
            CINGRE# = XVALO(!CantIngre)
            CCONSU# = XVALO(!Cantsalid)
29          SUINGRE# = SUINGRE# + CINGRE#
            SUCONSU# = SUCONSU# + CCONSU#
            '
            CAMOVILO% = CAMOVILO% + 1
            TIMOVILO%(CAMOVILO%) = 1: If Abs(CINGRE#) < 0.05 Then TIMOVILO%(CAMOVILO%) = 2
            CANINGLO#(CAMOVILO%) = CINGRE#
            CANSALLO#(CAMOVILO%) = CCONSU#
            CANRESLO#(CAMOVILO%) = 0
            FEMOVILO%(CAMOVILO%) = CVINT(!Fechmov)
            DEMOVILO$(CAMOVILO%) = SAFETEXT$(!DoPriLar)
               If Left$(!ReferCor, 7) = "RECHAZO" Then DEMOVILO$(CAMOVILO%) = "RECHAZO"
               If Left$(!ReferCor, 8) = "FALTANTE" Then DEMOVILO$(CAMOVILO%) = "FALTANTE"
         .MoveNext
         Loop
      End If
      On Error GoTo 0
    End With
    MOVCONTEMP.Close
    Set MOVCONTEMP = Nothing
    '
    Call APERBASDAT("RESERVA")
    SQLZ$ = " SELECT * FROM RESERVA WITH(NOLOCK) "
    SQLZ$ = SQLZ$ + "WHERE Cod_Inte = " & CIXI% & " "
    SQLZ$ = SQLZ$ + "AND IdenLote = '" & IDLOTX$ & "' "
    SQLZ$ = SQLZ$ + "AND Status = 0 "     ' para no considerar los marcados para anular
    SQLZ$ = SQLZ$ + "ORDER BY FechRes ASC"
    '
    Dim ReserTemp As ADODB.Recordset
    Set ReserTemp = New ADODB.Recordset
    Set ReserTemp = FLEXCONN.Execute(FILTSQL$(SQLZ$))
    '
    SURESER# = 0
    With ReserTemp
      If Not (.EOF And .BOF) Then
         Do While Not .EOF
           CANRE = 0: CANCO = 0
           On Error Resume Next
           CANRE = XVALO(!CANTRES)
           CANCO = XVALO(!CANTCONS)
           On Error GoTo 0
           CRESER# = CANRE - CANCO
           If CRESER# >= 0.05 Then
              CAMOVILO% = CAMOVILO% + 1
              TIMOVILO%(CAMOVILO%) = 3
              CANINGLO#(CAMOVILO%) = 0
              CANSALLO#(CAMOVILO%) = 0
              CANRESLO#(CAMOVILO%) = CRESER#
              FEMOVILO%(CAMOVILO%) = CVINT(!FechRes)
              DEMOVILO$(CAMOVILO%) = SAFETEXT$(!IDSUBOR)
              SURESER# = SURESER# + CRESER#
           End If
         .MoveNext
         Loop
      End If
      On Error GoTo 0
    End With
    ReserTemp.Close
    Set ReserTemp = Nothing
    '
End Sub
'
Sub BLANPEDIPE(NCLIE, FF%, NROREMIT$, CI%, CAN#, NUPEDIORI&, NUORITPE%)
    ' blanquear pedidos pendientes
    '
    If NUPEDIORI& < 1 Then
       Exit Sub    ' si no corresponde a ningún pedido no blanquea
    End If
    '
    Call APERBASDAT("PEDCLI")
    CANBLA# = CAN#
    VUELTA% = 1
    '
    ' PRIMERO BUSCA POR NRO-PED, NU-ORD-ITEM y CODIGO
10  sqlx$ = " SELECT * FROM PEDDETA "
    sqlx$ = sqlx$ + "WHERE NroPed = " & NUPEDIORI& & " "   ' COINCIDE NÚMERO DE PEDIDO
    sqlx$ = sqlx$ + "AND NroClie = " & NCLIE & " "        ' COINCIDE CLIENTE
    sqlx$ = sqlx$ + "AND CodiInt = " & CI% & " "           ' COINCIDE CÓDIGO
    sqlx$ = sqlx$ + "AND Status >= 0 "                     ' APROBADOS
    sqlx$ = sqlx$ + "AND Status < 8 "                      ' NO ANULADOS NI CUMPLIDOS
    sqlx$ = sqlx$ + "AND (CanPed > CantEnt OR TipoPed = 'ABIERTO') "  ' QUE TENGAN SALDO DE ENTREGA
    If VUELTA% <= 1 Then 'en el renglon de arriba se le agregaron los parentesis por que si no estaba evaluando mal la consulta
      sqlx$ = sqlx$ + " AND NuOrItem = " & NUORITPE% & " "  ' COINCIDE NRO DE ORDEN
    End If
    'SQLX$ = SQLX$ + "ORDER BY FeSolEnt ASC"
    'Set PEDSEL = Ventas.OpenRecordset(SQLX$, dbOpenDynaset)
    Dim PEDSEL As ADODB.Recordset
    Set PEDSEL = New ADODB.Recordset
25  On Error Resume Next
    PEDSEL.Open FILTSQL$(sqlx$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
    If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
      On Error GoTo 0
      Call CapturaErrorOpen
      GoTo 25
    End If
    On Error GoTo 0

    '
    With PEDSEL
      On Error Resume Next
      .MoveFirst
      If Err < 1 Then
         On Error GoTo 0
         Do While Not .EOF
         AA = XVALO(!NROPED)
            CSALDO# = XVALO(!CanPed) - XVALO(!CantEnt): If CSALDO# < 0 Then CSALDO# = 0
            If SAFETEXT$(!TipoPed) = "ABIERTO" Then CSALDO# = CANBLA#
            CANAPLI# = CANBLA#: If CANAPLI# > CSALDO# Then CANAPLI# = CSALDO#
              '.Edit
              !CantEnt = XVALO(!CantEnt) + CANAPLI#
              !FeUltEnt = CVDAT(FF%)
              !NroRemi = NROREMIT$
              If IsNull(!TipoPed) = True Or !TipoPed <> "ABIERTO" Then
                !CANSALDO = XVALO(!CanPed) - XVALO(!CantEnt)
              End If
            .Update
            CANBLA# = CANBLA# - CANAPLI#
            If CANBLA# < 0.05 Then
               PEDSEL.Close
               Set PEDSEL = Nothing
                              
               Exit Sub
            End If
         .MoveNext
         Loop
      End If
      On Error GoTo 0
    End With
    '
    PEDSEL.Close
    Set PEDSEL = Nothing
    If CANBLA# >= 0.05 Then
       If VUELTA% < 2 Then
          VUELTA% = 2
          GoTo 10
       End If
    End If
    '
    TTXX$ = "La Cantidad Remitida de " + TRIM$(FORMATNUM$(CAN#, "F10.1")) + " del Artículo"
    TTXX$ = TTXX$ + "\'" + CODEXT$(CI%) + "' con el Comprobante '" + NROREMIT$ + "' "
    TTXX$ = TTXX$ + "\no Pudo Descontarse Totalmente del Pedido '" + TRIM$(Str$(NUPEDIORI&)) + "'."
    TTXX$ = TTXX$ + "\   \Consulte al Soporte de Sistemas FLEXOFT."
    Call MENSERR(24, TTXX$)
    '
End Sub

Function FECHVENLOT%(NUBOLREC$)
    FECHAVEN% = 0
    Call APERBASDAT("STOCKS")
         sqlx$ = " SELECT * FROM BOLRECEP "
         sqlx$ = sqlx$ + "WHERE IdenLote = '" & NUBOLREC$ & "' "
         sqlx$ = sqlx$ + "ORDER BY FechRecep ASC, IdenLote ASC "
         'Set BolRecepTemp = Stocks.OpenRecordset(SQLX$, dbOpenSnapshot)
         Dim BolRecepTemp As ADODB.Recordset
         Set BolRecepTemp = FLEXCONN.Execute(FILTSQL$(sqlx$))

         With BolRecepTemp
           On Error Resume Next
           .MoveFirst
           If Err < 1 Then
              FECHAVEN% = CVINT(!PLAZO_USO)
           End If
           On Error GoTo 0
        End With
        Set BolRecepTemp = Nothing
      FECHVENLOT% = FECHAVEN%
End Function

Sub VERIDESPRE(NPEDI&, NORITPE%, HUBORECU%)
   '
   Static DEPEX%, DETER%, EPACXX$
   Static NPEDXA&, KKF%, CIKU%(8192), NPKU&(8192), CAYADES(8192)
   '
   If EPACXX$ = "" Then
     EPACXX$ = TRIM$(UCase$(EMPREAC$))
   End If
   '
   If InStr(EPACXX$, "DOSIVA") < 1 Then
     Exit Sub
   End If
   '
   ' DETERMINA DEPÓSITOS DE EXPEDICIÓN Y TERMINADOS
   If DEPEX% < 1 Or DETER% < 1 Then
     DEPEX% = Int(XVALO(Control$("", "DEPOEXPE")))
     DETER% = Int(XVALO(Control$("", "DEPOTERM")))
     If DEPEX% < 1 Or DETER% < 1 Or DEPEX% = DETER% Then
       Call MENSERR(24, "Proceso Imposible.\  \Falta Configurar Depósitos\de Expedición\y Productos Terminados.")
       Call FINAL("")
     End If
   End If
   '
    Dim MIRS As ADODB.Recordset
10  HUBORECU% = 0
    sqlx$ = " SELECT * FROM PEDDETA "
    sqlx$ = sqlx$ + "WHERE NroPed = " & NPEDI& & " "   ' COINCIDE NÚMERO DE PEDIDO
    If NORITPE% > 0 Then
      sqlx$ = sqlx$ + " AND NuOrItem = " & NORITPE% & " "  ' COINCIDE NRO DE ORDEN
    End If
    '
    Dim DESPAPE As ADODB.Recordset
    Set DESPAPE = New ADODB.Recordset
15  On Error Resume Next
    DESPAPE.Open FILTSQL$(sqlx$), FLEXCONN, adOpenDynamic, adLockPessimistic, adCmdText
    If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
      On Error GoTo 0
      Call CapturaErrorOpen
      GoTo 15
    End If
    On Error GoTo 0
    '
    With DESPAPE
      On Error Resume Next
      .MoveFirst
      If Err < 1 Then
         On Error GoTo 0
         Do While Not .EOF
           STATX% = XVALO(!StaDespa)
           REPARX% = 0: If !VenRep = "R" Then REPARX% = 1
           NPX& = XVALO(!NROPED)
           FEMIPED% = CVINT(!FechEmis)
           NC% = XVALO(!nroclie)
           CI% = XVALO(!CODIINT)
           AAA = CODEXT(CI%)
           CAPEDID = XVALO(!CanPed)           ' CANTIDAD PEDIDA
           NORPE% = XVALO(!nuoritem)
           If CI% > 0 And CI% <= NUMAS% Then
              DSLX$ = TRIM$(FORMATNUM$(NPX&, "F8.0"))
              While Len(DSLX$) < 6: DSLX$ = "0" + DSLX$: Wend
              DSLX$ = DSLX$ + "-" + TRIM$(Str$(NORPE%))
              '
              ACUPREA = XVALO(!AcuPrepa): ACUPRE = 0
              ACUYENA = XVALO(!AcuDespa): ACUYEN = 0
              '
              STRSQL$ = "SELECT SUM(CantIngre) as ACUIN, SUM(CantSalid) as ACSAL FROM MOVISTO WITH(NOLOCK)"
              STRSQL$ = STRSQL$ + " WHERE Cod_Inte = " & CI%
              STRSQL$ = STRSQL$ + " AND DepoMovi = " & DEPEX%
              STRSQL$ = STRSQL$ + " AND NuPedCli = " & NPEDI&
              STRSQL$ = STRSQL$ + " AND (NumeOrdPed = " & NORPE% & " OR DOSECLAR = '" & DSLX$ & "') "
               'STRSQL$ = STRSQL$ + NRO DE ORDEN DEL ITEM EN EL  PEDIDO
               
              Set MIRS = FLEXCONN.Execute(FILTSQL$(STRSQL$ + " AND CodiMovi = 'PP'"))
              'If Not IsNull(mirs!ACUP) Then ACUPRE = mirs!ACUP
              ACUPRE = XVALO(MIRS!ACUIN) - XVALO(MIRS!ACSAL)
              Set MIRS = FLEXCONN.Execute(FILTSQL$(STRSQL$ + " AND CodiMovi = 'DM'"))
              'If Not IsNull(mirs!ACUP) Then ACUYEN = (-1) * mirs!ACUP
              ACUYEN = XVALO(MIRS!ACSAL) - XVALO(MIRS!ACUIN)
             MIRS.Close
              '     '
              If FEMIPED% < FECHANUM("07/11/05") Then
                 ' IR A BUSCAR DESPACHOS REALIZADOS CON ANTERIORIDAD
                 ' DEVOLVIENDO SOBRE CAYENANT
                 NPEDXX& = NPX&
                 REDES& = 1  ' PARA FORZAR EL CÁLCULO
                 GoSub CARYENANT
                 ACUYEN = ACUYEN + CAYENANT
              End If
              '
              ' CORRECCION DEL 28/02/06 - SI DESPACHADO ES > PREPARADO
              ' ENTONCES PREPARADO = DESPACHADO
              If ACUYEN > ACUPRE Then
                ACUPRE = ACUYEN
              End If
              '
              ' GRABA LOS VALORES ACUMULADOS DE PREPARACION Y DESPACHO
              If STATX% <> 1 Then
                If Abs(ACUPRE - ACUPREA) > 0.05 Then HUBORECU% = 1
                If Abs(ACUYEN - ACUYENA) > 0.05 Then HUBORECU% = 1
              End If
              '
              !AcuPrepa = ACUPRE
              !AcuDespa = ACUYEN
              If STATX% Mod 10 < 2 Then
                STATNUE% = STATX% Mod 10
                If ACUPRE >= CAPEDID - 0.05 Then
                  If ACUYEN >= CAPEDID - 0.05 Then
                    STATNUE% = 3      ' LO PASA A CUMPLIDO SI ACUMULADO MAYOS A PEDIDO
                  End If
                End If
                !StaDespa = STATNUE%
              End If
              .Update
              '
           End If
         .MoveNext
         Loop
      End If
    End With
    Exit Sub
   '
   '
CARYENANT:
   CAYENANT = 0
   If NPEDXX& <> NPEDXA& Or REDES& < 1 Then
     KKF% = 0
   End If
   If REDES& < 1 Then GoTo 29
   '
   For KKV% = 1 To KKF%
     If CIKU%(KKV%) = CI% Then
       If NPKU&(KKV%) = NPEDXX& Then
         GoTo 25
       End If
     End If
   Next KKV%
   '
   KKF% = KKF% + 1
   CIKU%(KKF%) = CI%
   NPKU&(KKF%) = NPEDXX&
   CAYADES(KKF%) = 0
   ABUPED$ = "PL." + TRIM$(Str$(NPEDXX&))
   CADESPAR = 0
   STRSQL$ = "SELECT SUM(CantSalid) as ACUP1 FROM MOVISTO WITH(NOLOCK)"
   STRSQL$ = STRSQL$ + " WHERE Cod_Inte = " & CI%
   STRSQL$ = STRSQL$ + " AND DoPriCor = '" & ABUPED$ & "' "
   STRSQL$ = STRSQL$ + " AND FechMov < '07/NOV/2005' "
   Set MIRS = FLEXCONN.Execute(FILTSQL$(STRSQL$))
   If Not IsNull(MIRS!ACUP1) Then
      CADESPAR = MIRS!ACUP1
   End If
   MIRS.Close
   '           '     '
   If CADESPAR >= 0.005 Then
        CAYADES(KKF%) = CAYADES(KKF%) + CADESPAR
   End If
   KKV% = KKF%
   '
25 If CAYADES(KKV%) <= CAPEDID Then
       CAYENANT = CAYADES(KKV%)
       CAYADES(KKV%) = 0
     Else
       CAYADES(KKV%) = CAYADES(KKV%) - CAPEDID
       CAYENANT = CAPEDID
   End If
   '
29 Return
'
End Sub

Sub PRINSTIKAPRO(CI%, LOTE$, DEPO%)
    '
    Call PRINSTIKAPRO1(CI%, LOTE$, DEPO%)
'MsgBox "Sticker de Aprobación " & TRIM$(CODEXT$(CI%)) & " - " & LOTE$ & " - " & DEPO%
Exit Sub
    FORIPRE$ = Control$("", "STIKAPRO")
    If TRIM$(FORIPRE$) <> "" Then
      LU$ = LUDAT$
      If EXISTE%(LU$ + FORIPRE$ + ".FRM") > 0 Then
        '
        TIPODOC$ = "'Stiker de Aprobación"
        HOII% = HOY(HOS$)
        FECHDOC$ = HOS$
        NUMEDOC$ = LOTE$
        '
        DESTIDOC% = 0
        CACOLTAB1% = 0
        CAITAB1% = 0
        '
        Call PRINTDOC("STIKAPRO")   '
        '
      End If
    End If
    '
End Sub

Sub PRINSTIKAPRO1(CI%, LOTE$, Optional DEPO%, Optional CANTI)
    '
    FORIPRE$ = TRIM$(Control$("", "STIKAPRO"))
    If FORIPRE$ = "" Then Exit Sub
    '
    Call ACSALOTE(CI%, LOTE$, "NOMESS")
    '
    PFAC$ = Control$(IDENTER$, "PORTSTIK")
    If PFAC$ = "" Then PFAC$ = Control$("*", "PORTSTIK")
    If PFAC$ = "" Then Exit Sub ' POR SI NOHAY NINGUNA IMPRESORA CONFIGURADA
    Call SETPRINPORT(PFAC$)
    
    Call ABRESTOCKS
    '
    CONDI$ = "Cod_Inte = " & CI% & " AND IdenLote = '" & LOTE$ & "'"
    '
    If DEPO% > 0 Then
        CONDI$ = CONDI$ & " AND DepoUbic = " & DEPO%
    End If
    '
    CODIGO$ = CODEXT$(CI%)
    DESXRIP$ = DESCRIT0$(CI%)
    LOTEE$ = LOTE$
    DEPOSITO% = XVALO(VALOBADA("LOTINGRE", "DepoUbic", CONDI$))
    UBIFISICA$ = VALOBADA("LOTINGRE", "UbicFis", CONDI$)
    LARGO# = XVALO(VALOBADA("LOTINGRE", "MetroIni", CONDI$))
    ancho# = XVALO(VALOBADA("LOTINGRE", "AnchoMili", CONDI$))
    FF% = CVINT(VALOBADA("LOTINGRE", "FechAlta", CONDI$))
    FECHARECEP$ = FECHATEX$(FF%)
    CANTINIC# = XVALO(VALOBADA("LOTINGRE", "CanAlta", CONDI$))
    CANTICON# = XVALO(VALOBADA("LOTINGRE", "CanConsu", CONDI$))
    CANTISAL# = CANTINIC# - CANTICON#: If CANTISAL# < 0 Then CANTISAL# = 0
    If CANTINIC# > 0 Then
       If ancho# > 0 Then LARGO# = 1000 * CANTINIC# / ancho# ' SE AGREGA PARA CUANDO SE REALIZO UN AJUSTE POR UNA CANTIDAD MAYOR A LA DEL LOTE
       LARGO# = LARGO# * CANTISAL# / CANTINIC# ' para que aparezca el saldo
    End If
    ' GRABA EL LARGO AJUSTADO
    CSALNE# = XVALO(VALOBADA("LOTINGRE", "CanSaldo", CONDI$))
    If ancho# > 0.05 Then LARGO# = 1000 * CSALNE# / ancho#
    Call ACTUREGISTRO("LOTINGRE", "SalMetro", CONDI$, LARGO#, RAFE&, "NOMESS")
    '
    REMIPRO$ = VALOBADA("LOTINGRE", "DocuIngre", CONDI$)
    CODESTI$ = VALOBADA("LOTINGRE", "Cod_Desti", CONDI$)
       If Left$(CODESTI$, 2) = "OF" Or Left$(CODESTI$, 2) = "SF" Then
          CIXI% = XVALO(VALOBADA("ORDEFABR", "Cod_Inte", "IdSubOr = '" & CODESTI$ & "' "))
          If CIXI% > 0 Then
             CODESTI$ = DESCRIT0$(CIXI%)
          End If
       End If
       '
    LOTEPROV$ = VALOBADA("LOTINGRE", "LoteCol_Prov", CONDI$)
    Tara# = XVALO(VALOBADA("LOTINGRE", "Tara", CONDI$))
    POTENCIA# = XVALO(VALOBADA("LOTINGRE", "Potencia", CONDI$))
    ESTADO$ = VALOBADA("LOTINGRE", "Libera", CONDI$)
       If Left$(TRIM$(UCase$(ESTADO$)), 4) = "CUAR" Then ESTADO$ = "CUARENTENA"
       If Left$(TRIM$(UCase$(ESTADO$)), 4) = "APRO" Then ESTADO$ = "APROBADO"
       If Left$(TRIM$(UCase$(ESTADO$)), 4) = "RECH" Then ESTADO$ = "RECHAZADO"
    FF1% = CVINT(VALOBADA("LOTINGRE", "Fechco_Rep", CONDI$))
    FECHARETEST$ = FECHATEX$(FF1%)
    '
    If Control("BOLRECEP", "FORMATO") = "TXT" Then
        TXTX$ = TRIM$(FORIPRE$)
        If TRIM$(FORIPRE$) <> "" Then
            GoSub PRINZEBRA
          Else
            Call MENSERR(24, "Falta Definir Archivo Transferencia de Impresión")
        End If
        GoTo 90
    End If
    '
    TIPODOC$ = "Etiqueta Identificador de Lote"
    CODPARAM$(1) = "COD-MAT": DOCPARAM$(1) = CODIGO$
    CODPARAM$(2) = "NRO-LOTE": DOCPARAM$(2) = LOTEE$
    CODPARAM$(3) = "DEPOSITO": DOCPARAM$(3) = DEPOSITO%
    CODPARAM$(4) = "UBI-MPRI": DOCPARAM$(4) = UBIFISICA$
    CODPARAM$(5) = "MEDIDA-1": DOCPARAM$(5) = FORMATNUM$(LARGO#, "F15.4")
    CODPARAM$(6) = "MEDIDA-2": DOCPARAM$(6) = FORMATNUM$(ancho#, "F15.4")
    CODPARAM$(7) = "CANTIDAD"
    CODPARAM$(8) = "REF-MAT1"
    '
    If XVALO(CANTI) > 0 Then
       DOCPARAM$(7) = FORMATNUM$(CANTI, "F15." & CNTDC$)
       'DOCPARAM$(8) = AJUSTI$(CODIGO$, 16) + AJUSTI$(LOTEE$, 12) + FORMATNUM$(CANTI, "F8.0")
       DOCPARAM$(8) = TRIM(CODIGO$) + Chr$(124) + TRIM(LOTEE$) + Chr(124) + TRIM(FORMATNUM$(CANTI, "F8.0"))
    Else
       DOCPARAM$(7) = FORMATNUM$(CANTISAL#, "F15." & CNTDC$)
          TXXT$ = AJUSTI$(CODIGO$, 16) + AJUSTI$(LOTEE$, 16)
       DOCPARAM$(8) = TXXT$
    End If
    '
    CODPARAM$(9) = "DES-MAT": DOCPARAM$(9) = DESXRIP$
    CODPARAM$(10) = "MEDIDA-3": DOCPARAM$(10) = AJUSTI$(CODIGO$, 16) + AJUSTI$(LOTEE$, 12)
    CODPARAM$(11) = "FECH-REC": DOCPARAM$(11) = FECHARECEP$
    CODPARAM$(12) = "NRO-REMI": DOCPARAM$(12) = REMIPRO$
    CODPARAM$(13) = "NRO-AENT": DOCPARAM$(13) = CODESTI$
    CODPARAM$(14) = "PAR-MPRI": DOCPARAM$(14) = LOTEPROV$
    CODPARAM$(15) = "PES-CONJ": DOCPARAM$(15) = FORMATNUM$(Tara#, "F10.3")
    CODPARAM$(16) = "COL-CONJ": DOCPARAM$(16) = FORMATNUM$(POTENCIA#, "F10.3")
    CODPARAM$(17) = "REF-MAT2": DOCPARAM$(17) = ESTADO$
    CODPARAM$(18) = "FECH-LAN": DOCPARAM$(18) = FECHARETEST$
    CODPARAM$(19) = "CAN-CONJ": DOCPARAM$(19) = ""
    CODPARAM$(20) = "UNI-CONJ": DOCPARAM$(20) = ""
    '
    If UCase(UNIMED$(CI%)) = "KG" Or UCase(UNIMED$(CI%)) = "GR" Then
        CODPARAM$(19) = "CAN-CONJ": DOCPARAM$(19) = FORMATNUM$((CANTISAL# + Tara#), "F10.3")
        CODPARAM$(20) = "UNI-CONJ": DOCPARAM$(20) = UNIMED$(CI%)
        Else
        CODPARAM$(20) = "UNI-CONJ": DOCPARAM$(20) = "KG"
    End If
    '
    For pos% = Len(LOTEE$) To 1 Step -1
       LETRAX$ = Mid$(LOTEE$, pos%, 1)
       If LETRAX$ = "-" Then GoTo 85
       If Asc(LETRAX$) < 48 Or Asc(LETRAX$) > 57 Then
          GoTo 84
       End If
    Next pos%
    '
84  pos% = 1 + Len(LOTEE$)
    ' POS% = InStr(4, LOTEE$, "-")    ' LEE DESDE POS 4 PARA EVITAR EL PRIMER GUION
    ' If POS% < 4 Then POS% = 1 + Len(LOTEE$)
85  CANBU = XVALO(Mid$(LOTEE$, pos% + 1)) ' CANTIDAD DE BULTOS
    If CANBU < 1 Then CANBU = 1
    CANBUL$ = TRIM$(CStr(CANBU))
    LTE$ = Mid(LOTEE$, 1, pos% - 1)
    CONDI$ = "IDENLOTE LIKE '" & LTE$ & "%" & "'"
    CANITEM$ = CantRegistros&("LOTINGRE", CONDI$) ' CANTIDAD DE ITEMS
    CODPARAM$(21) = "REF-MAT6": DOCPARAM$(21) = CANBUL$ & " de " & CANITEM$
    '
    CODPARAM$(22) = "REF-MAT4": DOCPARAM$(22) = ""
    CODPARAM$(23) = "REF-MAT5": DOCPARAM$(23) = ""
    '
    DEST$ = SAFETEXT(VALOBADA("LOTINGRE", "COD_DESTI", CONDI$))      ' obtiene la S/F
    OP$ = SAFETEXT$(VALOBADA("OrdeFabr", "Destino", "IDSUBOR='" & DEST$ & "'"))
    If OP$ <> "" Then
       CADENA$ = "O.P. " & OP$
       DOCPARAM$(22) = CADENA$
    End If
    '
    PROVX$ = SAFETEXT$(VALOBADA("LOTINGRE", "ORIGEN", CONDI$))
    DOCPARAM$(23) = PROVX$
    If TRIM$(PROVX$) = "" Then
       CLIX% = XVALO(SAFETEXT$(VALOBADA("ORDEFABR", "NUME_CLI", "IDSUBOR='" & DEST$ & "'")))
       DOCPARAM$(23) = FANTCLI$(CLIX%)
    End If
    '
    CONDI$ = "COD_INTE = " & CI% & " AND IDENLOTE = '" & LOTE$ & "'"
    PRECINT$ = VALOBADA("LOTINGRE", "PRECINTO", CONDI$, "NOMESS")
    CODPARAM$(24) = "REF-MAT3": DOCPARAM$(24) = PRECINT$
    '
    CODPARAM$(25) = "DESTINO": DOCPARAM$(25) = DEST$
    '
    CODPARAM$(26) = "UNIMED": DOCPARAM$(26) = UNIMED$(CI%)
    '
    'Si es etiqueta de Cierre de Fabricacion entonces imprime esto:
    If Mid$(LOTE$, 1, 2) = "OP" Then
       LX$ = Mid$(LOTE$, 4, 9)
       CONX$ = "NUORPROD =  '" & LX$ & "'"
       SF$ = VALOBADA("ORDEFABR", "IDSUBOR", CONX$)
       OPCLI$ = VALOBADA("ORDEFABR", "DESTINO", CONX$)
       CLIX% = XVALO(VALOBADA("ORDEFABR", "NUME_CLI", CONX$))
       CLIFA$ = FANTCLI$(CLIX%)
       LOT$ = SAFETEXT(VALOBADA("ORDEFABR", "IDENLOTE", CONX$))
       DOCPARAM$(2) = LOT$          ' Lote
       DOCPARAM$(14) = ""           ' Lote / Analisis / Ir
       DOCPARAM$(22) = OPCLI$       ' OP Cliente
       DOCPARAM$(23) = CLIFA$       ' Cliente
       DOCPARAM$(25) = SF$          ' Numero de SF
    End If
    '
    CANTPARR& = 26
    CAPARDOC = CANTPARR&
    CAITAB1% = 0
    '
    Call PRINTDOC("@" + FORIPRE$)
    '
90  Call SETPRINPORT("RESTORE")
    Exit Sub
    '
PRINZEBRA:
    '
    TXTX$ = LOADFILE(LUDAT$ + FORIPRE$ + ".TXT")
    Call REPLATEXT(TXTX$, Chr$(13) + Chr$(10), Chr$(10))
    Call REPLATEXT(TXTX$, "@CODI-MAT", CODIGO$)
    Call REPLATEXT(TXTX$, "@NRO-LOTE", LOTEE$)
    Call REPLATEXT(TXTX$, "@DEPOSITO", Str$(DEPOSITO%))
    Call REPLATEXT(TXTX$, "@UBI-MPRI", UBIFISICA$)
    Call REPLATEXT(TXTX$, "@MEDIDA-1", FORMATNUM(LARGO#, "F10.1"))
    Call REPLATEXT(TXTX$, "@MEDIDA-2", FORMATNUM(ancho#, "F10.1"))
    Call REPLATEXT(TXTX$, "@CANTIDAD", FORMATNUM(CANTIDAD#, "F10.1"))
    Call REPLATEXT(TXTX$, "@DESC-MAT", DESXRIP$)
    Call REPLATEXT(TXTX$, "@REF-MAT1", AJUSTI$(CODIGO$, 16) + AJUSTI$(LOTEE$, 12))
    '
    TXTX1$ = TXTX$
    ARCHOUT$ = LUCOM$ + FORIPRE$ + ".PRN"
    Call SAVEFILE(ARCHOUT$, TXTX$)
    'PFAC$ = TRIM$(CONTROL$(IDENTER$, "PORTSTIK"))
    'Call DOSCOMMAND("PRINT " + ARCHOUT$)
    '
75  On Error Resume Next
    Printer.Print TXTX1$
    If Err Then
       On Error GoTo 0
       OPX% = COMALTER%("Error de la Impresora\\Re-Intentar\Abandonar")
       If OPX% = 1 Then GoTo 75
    End If
    On Error GoTo 0
    '
Return
'
End Sub

Function LOTESELEC$(CIXI%, DEPX%, CANTSA)
    SELECLOT07.Label4 = TRIM$(Str$(DEPX%))
    SELECLOT07.Label5 = FORMATNUM$(CANTSA, "F10.1")
    SELECLOT07.Label1 = CODEXT$(CIXI%)
    SELECLOT07.Label28 = ""
    '
    Call ARMALILOT(CIXI%, DEPX%)
    For U& = 1 To ULTREG("!LILOTSE")
       XX$ = REGLEIDO$("!LILOTSE", U&)
       IDLO$ = TRIM$(Left$(XX$, 16))
       If SaldoLote(CIXI%, IDLO$, DEPX%) >= CANTSA Then
          SELECLOT07.Label28 = IDLO$
          GoTo 90
       End If
    Next U&
    Call MENSERR(24, "No hay Lotes Disponibles\en Cantidad Suficiente\en el Depósito Elegido,")
    '
90  SCMA = Screen.MousePointer
    Screen.MousePointer = 1
    SELECLOT07.Show 1
    LOTESELEC$ = TRIM$(SELECLOT07.Label28)
    Screen.MousePointer = SCMA
End Function

Function ESTAPREPARADO%(NPEDX&)
    '\\\OT-08-0784-OD-06/09/08///
    
    EstPrep = 0
    '
    'GENERO LA CONDICION QUE ESTE MARCADO COMO VENTA ,QUE ESTE ACTIVO,QUE NO ESTE
    'DESPACHADO, SI ACUMULO ALGUNA PREPARACION , SI COINCIDE EL NUMERO DE PEDIDO
    
    CONDICION$ = " VenRep = 'V'"
    CONDICION$ = CONDICION$ + " AND Status < 8 AND StaDespa <= 3"
    CONDICION$ = CONDICION$ + " AND AcuPrepa > 0 AND NroPed = " & NPEDX& & ""
    '
    EstPrep = XVALO(VALOBADA("PEDDETA", "ACUPREPA", CONDICION$))
    If EstPrep > 0 Then EstPrep = 1
    '
    ESTAPREPARADO% = EstPrep
    '\\\OT-08-0784-OD-FIN///

End Function

Sub AGREGA_ANOTAPED(NPPXX&, TXTX$)
    '
    If NPPXX& > 0 Then
       Call ABREPEDCLI
       sqlx$ = "SELECT ANOTAPED FROM PEDENCA"
       sqlx$ = sqlx$ + " WHERE NroPed = " & NPPXX& & ""
       Dim PEDDDENCA As New ADODB.Recordset
       Set PEDDDENCA = New ADODB.Recordset
25     On Error Resume Next
       PEDDDENCA.Open FILTSQL$(sqlx$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
       If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
           On Error GoTo 0
           Call CapturaErrorOpen
           GoTo 25
       End If
       On Error GoTo 0
       '
       With PEDDDENCA
         If Not (.BOF And .EOF) Then
            ANOTANTE$ = !AnotaPed
            ANOTA$ = ANOTANTE$ + TXTX$ + Chr$(10) + Chr$(13)
            !AnotaPed = ANOTA$
            .Update
         End If
       End With
       PEDDDENCA.Close
       Set PEDDDENCA = Nothing
       '
    End If
    '
End Sub

Sub AGREGA_ANOTAPED_NEW(NPPXX&, TXTX$)
    '
    If NPPXX& > 0 Then
        HOII% = HOY(HOS$)
        FECH$ = HOS$
        If FF% > 0 Then
              FECH$ = FECHATEX(FF%) ' SI TRAE FECHA EL PARAMETRO TOMO ESTA
        End If
        '
        Call FRATEXTO(TXTX$, 56) ' FRACCIONO EL TEXTO
    
        For KKI% = 1 To CARETEX% ' RECORRO DESDE 1 A LA CANTIDAD DE FRACCIONES
            TTYY$ = Space$(96)
            If KKI% = 1 Then
               Call REPLA(TTYY$, FECH$, 1, 8) ' SI ES LA PRIMERA VEZ GRABA LA FECHA Y USUARIO
               Call REPLA(TTYY$, USERTER$, 69, 24)
            End If
            Call REPLA(TTYY$, RETEX$(KKI%), 11, 56) ' SIEMPRE QUE PASA GUARDA EL CENTRO DEL TEXTO
            ANOTANTE$ = ANOTANTE$ + TTYY$ + vbCrLf
        Next KKI%
        '
        ANOTANTE$ = ANOTANTE$ & String$(88, "=") & vbCrLf

       Call ABREPEDCLI
       sqlx$ = "SELECT ANOTAPED FROM PEDENCA"
       sqlx$ = sqlx$ + " WHERE NroPed = " & NPPXX& & ""
       Dim PEDDDENCA As New ADODB.Recordset
       Set PEDDDENCA = New ADODB.Recordset
25     On Error Resume Next
       PEDDDENCA.Open FILTSQL$(sqlx$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
       If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
           On Error GoTo 0
           Call CapturaErrorOpen
           GoTo 25
       End If
       On Error GoTo 0
       '
       With PEDDDENCA
         If Not (.BOF And .EOF) Then
            NOTAORIG$ = SAFETEXT$(!AnotaPed)
            ANOTA$ = NOTAORIG$ + ANOTANTE$
            !AnotaPed = ANOTA$
            .Update
         End If
       End With
       PEDDDENCA.Close
       Set PEDDDENCA = Nothing
       '
    End If
    '
    
End Sub
Sub AGREGA_ANOTAPROSPECT(NPPXX$, TXTX$)
    '
    If NPPXX$ <> "" Then
       Call ABRECLIEN
       sqlx$ = "SELECT HIS_CONTAC FROM BSEGCOM"
       sqlx$ = sqlx$ + " WHERE Codi_Cli = '" & NPPXX$ & "'"
       Dim BSEGCOMTMP As ADODB.Recordset
       Set BSEGCOMTMP = New ADODB.Recordset
25     On Error Resume Next
       BSEGCOMTMP.Open FILTSQL$(sqlx$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
       If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
           On Error GoTo 0
           Call CapturaErrorOpen
           GoTo 25
       End If
       On Error GoTo 0
       '
       With BSEGCOMTMP
         If Not (.BOF And .EOF) Then
            ANOTANTE$ = SAFETEXT(!His_Contac)
            ANOTA$ = ANOTANTE$ + TXTX$ + Chr$(10) + Chr$(13)
            !His_Contac = ANOTA$
            .Update
         End If
       End With
       BSEGCOMTMP.Close
       Set BSEGCOMTMP = Nothing
       '
    End If
    '
End Sub

Function OCOPENDIENTE%(NUOCO&, Optional NORIT%)
   '
   Resultado% = 0
   Call ABRECOMPRAS
   '
   sqlx$ = "SELECT * FROM OCOMDETA WITH(NOLOCK) "
   sqlx$ = sqlx$ + " WHERE Nume_Ocom =" & NUOCO& & ""
   '
   If NORIT% > 0 Then
     sqlx$ = sqlx$ + " AND NuOrd_Item = " & NORIT% & ""
   End If
   '
   Dim ORDCOMTMP As ADODB.Recordset
   Set ORDCOMTMP = FLEXCONN.Execute(FILTSQL$(sqlx$))
   With ORDCOMTMP
     Do While Not .EOF
       If XVALO(!Stat_Ocom) < 8 And XVALO(!Cant_Rec) < XVALO(!cant_ocom) - 0.001 Then
         Resultado% = 1
         Exit Do
       End If
     .MoveNext
     Loop
   End With
   ORDCOMTMP.Close
   Set ORDCOMTMP = Nothing
   '
   If Resultado% = 1 Then
     CAMSTATUS% = VALOBADA("OCOMENCA", "Stat_Ocom", "Nume_Ocom = " & NUOCO& & "")
     If CAMSTATUS% = 9 Then
        Resultado% = 0
     End If
   End If
   '
   OCOPENDIENTE% = Resultado%

End Function

Sub AGREGA_ANOTASIM(NPPXX&, TXTX$)
    '
    If NPPXX& > 0 Then
       Call ABRECOMPRAS
       sqlx$ = "SELECT AnotaSolint FROM ENCASOLINT"
       sqlx$ = sqlx$ + " WHERE Nume_Solint = " & NPPXX& & ""
       Dim ANOTASOLTMP As New ADODB.Recordset
       Set ANOTASOLTMP = New ADODB.Recordset
25     On Error Resume Next
       ANOTASOLTMP.Open FILTSQL$(sqlx$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
       If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
           On Error GoTo 0
           Call CapturaErrorOpen
           GoTo 25
       End If
       On Error GoTo 0
       '
       With ANOTASOLTMP
         If Not (.BOF And .EOF) Then
            ANOTANTE$ = SAFETEXT(!AnotaSolint)
            ANOTA$ = ANOTANTE$ + TXTX$ + Chr$(10) + Chr$(13)
            !AnotaSolint = ANOTA$
            .Update
         End If
       End With
       ANOTASOLTMP.Close
       Set ANOTASOLTMP = Nothing
       '
    End If
    '
End Sub

Function PERMITEUBI%(CI%, DESTI$, UBICAPROP$, Optional DEPO%)
    'Esta Función Valida si es posible Ubicar un Item  en
    'una determinada posición
    Static DEPOA%
    PERMITEUBI% = -1
    '
    If INTERPRE% < 1 Then
       If EXISTE%("DECAOT09.EXE") < 1 Then
          PERMITEUBI% = 1
          Exit Function
       End If
    End If
    '
    If UBICAPROP$ = "***" Then
       PERMITEUBI% = 1
       Exit Function
    End If
    '
    If DEPO% <> DEPOA% Then
       DEPOA% = DEPO%
       Call GENTAUBI(DEPOA%)
    End If
    '
    CODE$ = TRIM$(UCase$(CODEXT$(CI%)))
    DESTIX$ = TRIM$(UCase$(DESTI$))
    '
    If ULTREG("!TAUBICA") < 1 Then PERMITEUBI% = 1: Exit Function
    XX$ = FILTERGETRECORD$("!TAUBICA", 1, UBICAPROP$)
    CODSTA$ = UCase$(TRIM$(Mid$(XX$, 17, 16)))
    If InStr(CODSTA$, "LIBRE") > 0 Then PERMITEUBI% = 1: Exit Function
    If CODE$ <> "" And InStr(CODSTA$, CODE$) > 0 Then PERMITEUBI% = 1: Exit Function
    If DESTI$ <> "" Then
       CONDI$ = "UbicFis = '" & UBICAPROP$ & "' "
       CONDI$ = CONDI$ & " AND CANALTA > CANCONSU "
       CONDI$ = CONDI$ & " and Cod_Desti = '" & DESTI$ & "'"
       If DEPO% > 0 Then CONDI$ = CONDI$ + " AND DEPOUBIC = " & DEPO% & " "
       If CantRegistros("LOTINGRE", CONDI$) > 0 Then
          PERMITEUBI% = 1: Exit Function
       End If
    End If
    '
    If TRIM$(UBICAPROP$) = "" Then
         Call MENSERR(24, "Ubicación no Válida")
       Else
         Call MENSERR(24, "No Válido Colocar " & CODEXT(CI%) & " en Posición " & UBICAPROP$)
    End If
Exit Function
'
' *************************************************************************
'   ESTA ES LA VIEJA VERSION DISCONTINUADA AL 9/03/2010
    CONDI0$ = "UbicFis = '" & UBICAPROP$ & "' AND CANALTA > CANCONSU "
    If DEPO% > 0 Then CONDI0$ = CONDI0$ + " AND DEPOUBIC% = " & DEPO%
    If CantRegistros("LOTINGRE", CONDI0$) = 0 Then
       PERMITEUBI% = 1
       Exit Function
    End If
    '
    CONDI$ = "Cod_Inte =" & CI% & " and " & CONDI0$
    If CantRegistros("LOTINGRE", CONDI$) > 0 Then
       PERMITEUBI% = 1
       Exit Function
    End If
    '
    CONDI$ = CONDI0$ & "' and Cod_Desti='" & DESTI$ & "'"
    If CantRegistros("LOTINGRE", CONDI$) > 0 Then
       PERMITEUBI% = 1
       Exit Function
    End If
    '
    Call COMUNI("No Válido Colocar " & CODEXT(CI%) & " en Posición " & UBICAPROP$)
End Function

Sub APARTIMA()
    '
    Call ABREMASTER
    Call ABRESTOCKS
    '
    CONDI$ = " TIMATE IS NULL "
    Call ACTUREGISTRO("LOTINGRE", "TIMATE", CONDI$, 0, REGAF&)
    '
    sqlx$ = " SELECT  MASTER.TIMATE AS MATIMATE, LOTINGRE.TIMATE AS LOTIMATE "
    sqlx$ = sqlx$ + " FROM MASTER "
    sqlx$ = sqlx$ + " INNER JOIN LOTINGRE "
    sqlx$ = sqlx$ + " ON MASTER.CODINT = LOTINGRE.COD_INTE "
    sqlx$ = sqlx$ + " WHERE MASTER.TIMATE <> LOTINGRE.TIMATE "
    Dim RESETMP As ADODB.Recordset
    Set RESETMP = New ADODB.Recordset
25  On Error Resume Next
    RESETMP.Open FILTSQL$(sqlx$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
    If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
      On Error GoTo 0
      Call CapturaErrorOpen
      GoTo 25
    End If
    On Error GoTo 0
    With RESETMP
      Do While Not .EOF
        If XVALO(!LOTIMATE) <> XVALO(!MATIMATE) Then
         !LOTIMATE = XVALO(!MATIMATE)
         .Update
        End If
      .MoveNext
      Loop
    End With
    RESETMP.Close
    Set RESETMP = Nothing
    '
    CONDI$ = " PESUNI IS NULL "
    Call ACTUREGISTRO("MASTER", "PESUNI", CONDI$, 0, REGAF&)
    CONDI$ = " CANSALDO IS NULL "
    Call ACTUREGISTRO("LOTINGRE", "CANSALDO", CONDI$, 0, REGAF&)
    '
    sqlx$ = " SELECT  MASTER.PESUNI AS METCUA, LOTINGRE.CANSALDO AS SALDLOT , LOTINGRE.SalMeCua"
    sqlx$ = sqlx$ + " FROM MASTER "
    sqlx$ = sqlx$ + " INNER JOIN LOTINGRE "
    sqlx$ = sqlx$ + " ON MASTER.CODINT = LOTINGRE.COD_INTE "
    sqlx$ = sqlx$ + " WHERE MASTER.PESUNI > 0 AND LOTINGRE.CANSALDO > 0 "
    Dim SALMETTMP As ADODB.Recordset
    Set SALMETTMP = New ADODB.Recordset
26  On Error Resume Next
    SALMETTMP.Open FILTSQL$(sqlx$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
       If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
      On Error GoTo 0
      Call CapturaErrorOpen
      GoTo 26
    End If
    On Error GoTo 0
   
    With SALMETTMP
      Do While Not .EOF
         SUPERM2 = XVALO(!METCUA) * XVALO(!SALDLOT)
         !SalMeCua = SUPERM2
         .Update
      .MoveNext
      Loop
    End With
    SALMETTMP.Close
    Set SALMETTMP = Nothing

End Sub

Sub UBIRECOM(CIXI%, LOTE$, Optional DEP%)
    '
    Call GENTAUBI(DEP%)
    CODE$ = TRIM$(UCase$(CODEXT$(CIXI%)))
    DESTI$ = TRIM$(VALOBADA("LOTINGRE", "COD_DESTI", "IDENLOTE ='" & TRIM$(LOTE$) & "'"))
    '
    JJ& = 0
    For U& = 1 To ULTREG&("!TAUBICA")
       XX$ = REGLEIDO$("!TAUBICA", U&)
       CODSTA$ = UCase$(TRIM$(Mid$(XX$, 17, 16)))
       REFERE$ = TRIM$(Mid$(XX$, 33, 32))
       If InStr(CODSTA$, "LIBRE") > 0 Then GoTo 15
       If CODE$ <> "" And InStr(CODSTA$, CODE$) > 0 Then GoTo 15
       If DESTI$ <> "" And InStr(CODSTA$, DESTI$) > 0 Then GoTo 15
       GoTo 19
       '
15     If REFERE$ <> UREFE$ Then
          UREFE$ = REFERE$
          JJ& = JJ& + 1
          Call GRAREG("!TAUBICA", XX$, JJ&)
       End If
19  Next U&
    '
    Call SETULTREG("!TAUBICA", JJ&)
    Call FRESHECHO("!TAUBICA")
Exit Sub
    
    
    
'***************************************************************
    UBIPOSI$ = ""
    '
    CODE$ = CODEXT$(CIXI%)
    DESTI$ = TRIM$(VALOBADA("LOTINGRE", "COD_DESTI", "IDENLOTE ='" & TRIM$(LOTE$) & "'"))
    If DESTI$ = "" Then GoTo 200
    '
100 'Ubicaciones por Mismo Destino
    CONDI$ = "SELECT * FROM LOTINGRE WHERE COD_DESTI ='" & DESTI$ & "'"
    If ECOARCH$("TADEPOSI", Chr$(DEP%)) <> "" Then CONDI$ = CONDI$ & " AND DEPOUBIC = " & DEP%               ' SE AGREGA QUE LA BUSQUEDA SEA EN EL MISMO DEPOSITO"
    '
    Dim UBICA As ADODB.Recordset
    Set UBICA = FLEXCONN.Execute(FILTSQL$(CONDI$))
    Do While Not (UBICA.EOF)
       'GENERACION DE UBICACIONES
       UBIPO$ = AJUSTI$(UBICA!UBICFIS, 12)
       If TRIM$(UBIPO$) <> "" Then
          If InStr(UBIPOSI$, UBIPO$) < 1 Then
             UBIPOSI$ = UBIPOSI$ + UBIPO$
          End If
       End If
       UBICA.MoveNext
    Loop
    '
    UBICA.Close
    Set UBICA = Nothing
    '

200 'Ubicaciones por Mismo Material
    CONDI$ = "SELECT DISTINCT(UBICFIS)AS UBICACION  from lotingre where COD_INTE = " & CIXI% & " AND UBICFIS <> ''"
    If ECOARCH$("TADEPOSI", Chr$(DEP%)) <> "" Then CONDI$ = CONDI$ & " AND DEPOUBIC = " & DEP% ' SE AGREGA QUE LA BUSQUEDA SEA EN EL MISMO DEPOSITO
    '
    Set UBICA = FLEXCONN.Execute(FILTSQL$(CONDI$))
    Do While Not (UBICA.EOF)
       'GENERACION DE UBICACIONES
       UBIPO$ = AJUSTI$(UBICA!UBICACION, 12)
       If InStr(UBIPOSI$, UBIPO$) < 1 Then
          UBIPOSI$ = UBIPOSI$ + UBIPO$
       End If
    UBICA.MoveNext
    Loop
    '
    UBICA.Close
    Set UBICA = Nothing
    '
300 '1era Ubicacion Libre de cada Referencia
    REFANT$ = ""
    sqlx$ = "SELECT * FROM UBIPRE "
    If ECOARCH$("TADEPOSI", Chr$(DEP%)) <> "" Then
      sqlx$ = sqlx$ + " WHERE DEPOUBIC = " & DEP%
    End If
    sqlx$ = sqlx$ & " ORDER BY REFERENCIA"
    Dim UBIPRETMP As ADODB.Recordset
    Set UBIPRETMP = FLEXCONN.Execute(FILTSQL$(sqlx$))
    With UBIPRETMP
      Do While Not .EOF
        REFACT$ = TRIM$(!REFERENCIA)
        If REFACT$ <> REFANT$ Then
           UBIX$ = SAFETEXT$(!CODIUBIC)
           If CantRegistros("LOTINGRE", "UBICFIS = '" & UBIX$ & "'") < 1 Then
              REFANT$ = REFACT$
              UBIPO$ = AJUSTI$(UBIX$, 12)
              If InStr(UBIPOSI$, UBIPO$) < 1 Then
                 UBIPOSI$ = UBIPOSI$ + UBIPO$
              End If
           End If
        End If
        .MoveNext
      Loop
    End With
    '
    'Se Genera el Archivo TAUBICA
    JJ& = 0
    For U& = 1 To ULTREG&("!TAUBICA")
       XX$ = REGLEIDO$("!TAUBICA", U&)
       CODU$ = AJUSTI$(XX$, 12)
       If InStr(UBIPOSI$, CODU$) > 0 Then
         ABUX$ = UCase$(Mid$(XX$, 17, 24))
         If InStr(ABUX$, UCase$(CODE$)) > 0 Or (DESTI$ <> "" And InStr(ABUX, UCase$(DESTI$)) > 0) Or InStr(ABUX$, "LIBRE") > 0 Then
           JJ& = JJ& + 1
           Call GRAREG("!TAUBICA", XX$, JJ&)
         End If
       End If
    Next U&
    Call SETULTREG("!TAUBICA", JJ&)
    Call FRESHECHO("!TAUBICA")
End Sub

Sub ANOTANUSE(NSERIE$, TXTX$, Optional FF%, Optional MODO$)
    '
    If NSERIE$ <> "" Then
       '
'       ExNumSrie& = CantRegistros&("TANUMSE", "NUMEROSERIE = '" & NSERIE$ & "'") ' VALIDACION DE NUMERO DE SERIE
'       If ExNumSrie& < 1 Then
'          If MODO$ <> "NOMESS" Then
'             Call MENSERR(24, "Número de Serie: '" & NSERIE$ & "' Inexistente")
'          Else
'             Exit Sub
'          End If
'       End If
       '
       HOII% = HOY(HOS$)
       FECH$ = HOS$
       If FF% > 0 Then
          FECH$ = FECHATEX(FF%) ' SI TRAE FECHA EL PARAMETRO TOMO ESTA
       End If
       '
       Call FRATEXTO(TXTX$, 56) ' FRACCIONO EL TEXTO
       '
       For KKI% = 1 To CARETEX% ' RECORRO DESDE 1 A LA CANTIDAD DE FRACCIONES
           TTYY$ = Space$(96)
           If KKI% = 1 Then
              Call REPLA(TTYY$, FECH$, 1, 8) ' SI ES LA PRIMERA VEZ GRABA LA FECHA Y USUARIO
              Call REPLA(TTYY$, USERTER$, 69, 24)
           End If
           Call REPLA(TTYY$, RETEX$(KKI%), 11, 56) ' SIEMPRE QUE PASA GUARDA EL CENTRO DEL TEXTO
           ANOTANTE$ = ANOTANTE$ + TTYY$ + vbCrLf
       Next KKI%
       '
       ANOTANTE$ = ANOTANTE$ & String$(88, "=") & vbCrLf
       '
'       CONDI$ = "NUMEROSERIE = '" & NSERIE$ & "'"
'       '
'       ANOTRIOR$ = VALOBADA("TANUMSE", "HISTORIAL", CONDI$, "NOMESS") 'LA NOTA GUARDADA EN LA TABLA
'       '
'       ANOTA$ = ANOTRIOR$ & ANOTANTE$   ' CONCATENO LA ANTERIOR CON LA NUEVA
'       Call ACTUREGISTRO("TANUMSE", "HISTORIAL", CONDI$, ANOTA$, REGAF&, "NOMESS") 'GRABO
       '
       'SE REEMPLAZA PARA EVITAR EL ACTUREGISTRO, Y EL VALOBADA Y PASAR TODO EN UNA CONSULTA.
       sqlx$ = "SELECT HISTORIAL FROM TANUMSE"
       sqlx$ = sqlx$ + " WHERE NUMEROSERIE = '" & NSERIE$ & "'"
       Dim RS As New ADODB.Recordset
       Set RS = New ADODB.Recordset
25     On Error Resume Next
       RS.Open FILTSQL$(sqlx$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
       If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
           On Error GoTo 0
           Call CapturaErrorOpen
           GoTo 25
       End If
       On Error GoTo 0
       '
       With RS
         If Not (.BOF And .EOF) Then
            ANOTRIOR$ = SAFETEXT$(!Historial)
            !Historial = ANOTRIOR$ + ANOTANTE$
            .Update
         End If
       End With
       RS.Close
       Set RS = Nothing
       '
    End If
    '
End Sub
Function FORMATEA_ANOTADOR$(TXTX$, Optional FF%)
    '
    HOII% = HOY(HOS$)
    FECH$ = HOS$
    If FF% > 0 Then
       FECH$ = FECHATEX(FF%) ' SI TRAE FECHA EL PARAMETRO TOMO ESTA
    End If
    '
    Call FRATEXTO(TXTX$, 56) ' FRACCIONO EL TEXTO
    '
    For KKI% = 1 To CARETEX% ' RECORRO DESDE 1 A LA CANTIDAD DE FRACCIONES
        TTYY$ = Space$(96)
        If KKI% = 1 Then
           Call REPLA(TTYY$, FECH$, 1, 8) ' SI ES LA PRIMERA VEZ GRABA LA FECHA Y USUARIO
           Call REPLA(TTYY$, USERTER$, 69, 24)
        End If
        Call REPLA(TTYY$, RETEX$(KKI%), 11, 56) ' SIEMPRE QUE PASA GUARDA EL CENTRO DEL TEXTO
        ANOTANTE$ = ANOTANTE$ + TTYY$ + vbCrLf
    Next KKI%
    '
    ANOTANTE$ = ANOTANTE$ & String$(88, "=") & vbCrLf
    '
    FORMATEA_ANOTADOR$ = ANOTANTE$
    '
End Function

Sub GENTAUBI(Optional DEP%)
   '
   Dim CODUBI$(), OCUPUBI$(), DESTUBI$(), DENUBI$()
   Call APERBASDAT("STOCKS")
   SCMA = Screen.MousePointer
   Screen.MousePointer = 11
   CONDI$ = ""
   sqlx$ = "SELECT * FROM UBIPRE WITH(NOLOCK)"
   If ECOARCH$("TADEPOSI", Chr$(DEP%)) <> "" Then
      sqlx$ = sqlx$ + " WHERE DEPOUBIC = " & DEP%
      CONDI$ = "DEPOUBIC = " & DEP%
   End If
   sqlx$ = sqlx$ + " ORDER BY CODIUBIC ASC "
   Set LISUBIC = FLEXCONN.Execute(FILTSQL$(sqlx$))
   '
   Screen.MousePointer = 11
   FIN& = CantRegistros("UBIPRE", CONDI$)
   JJ& = 0
   U& = 0
   '
   If FIN& < 1 Then
     Call MENSERR(24, "No Existen Ubicaciones en el Depósito Donde se Encuentra el Lote.")
     GoTo 99
   End If
   ReDim CODUBI$(FIN&), OCUPUBI$(FIN&), DESTUBI$(FIN&), DENUBI$(FIN&)
   '
   With LISUBIC
      Do While Not .EOF
        JJ& = JJ& + 1
        Call TRACE(20, JJ&, FIN&)
        CODUBI$(JJ&) = UCase$(TRIM$(SAFETEXT$(!CODIUBIC)))
        DENUBI$(JJ&) = TRIM$(SAFETEXT$(!REFERENCIA))
      .MoveNext
      Loop
   End With
   '
   SQLY$ = "SELECT * FROM LOTINGRE WITH(NOLOCK) "  ' SELECCIONA LOS LOTES DISPONIBLES
   SQLY$ = SQLY$ + " INNER JOIN MASTER ON LOTINGRE.COD_INTE = MASTER.CODINT "
   SQLY$ = SQLY$ + "WHERE  (CanAlta - CanConsu) > 0 "
   SQLY$ = SQLY$ + " AND UBICFIS <> '' "
   If DEP% > 0 Then
      SQLY$ = SQLY$ + " AND DEPOUBIC = " & DEP%
   End If
   SQLY$ = SQLY$ + " ORDER BY UBICFIS"
   Dim LODISPO As ADODB.Recordset
   Set LODISPO = New ADODB.Recordset
   LODISPO.CursorLocation = adUseClient
   Set LODISPO = FLEXCONN.Execute(FILTSQL$(SQLY$))
   '
   KK& = 0: LOMAXI = 16
   With LODISPO
      Do While Not .EOF
        UBIFI$ = TRIM$(UCase$(SAFETEXT$(!UBICFIS)))
        For LL& = 1 To FIN&
           If CODUBI$(LL&) = UBIFI$ Then
              CODEPRO$ = OCUPUBI$(LL&)
              CODEDES$ = DESTUBI$(LL&)
              CIXI% = XVALO(!Cod_Inte)
              CODE$ = SAFETEXT$(!CODIGO)     '  CODEXT$(CIXI%)
              DEST$ = SAFETEXT$(!Cod_Desti)
              If InStr(CODEPRO$, CODE$) < 1 Then
                If CODEPRO$ <> "" Then
                     CODEPRO$ = "******"
                   Else
                     CODEPRO$ = " " + CODE$ + " "
                End If
              End If
              If CODEDES$ = "" Then
                If TRIM$(DEST$) <> "" Then
                  CODEDES$ = ">> " + TRIM$(DEST$) + " <<"
                End If
              End If
              OCUPUBI$(LL&) = CODEPRO$
              DESTUBI$(LL&) = CODEDES$
              If Len(CODEPRO$) > LOMAXI Then LOMAXI = Len(CODEPRO$)
              If Len(CODEDES$) > LOMAXI Then LOMAXI = Len(CODEDES$)
              GoTo 29
           End If
        Next LL&
29    .MoveNext
      Loop
   End With
   '
   For JJ& = 1 To FIN&
      If OCUPUBI$(JJ&) = "" Then
        If DESTUBI$(JJ&) = "" Then
          CODEPRO$ = " libre "
          GoTo 39
        End If
      End If
      CODEPRO$ = OCUPUBI$(JJ&)
      If CODEPRO$ = "******" Then
         CODEPRO$ = DESTUBI$(JJ&)
         If TRIM$(CODEPRO$) = "" Then
            CODEPRO$ = "<< VARIOS >>"
         End If
      End If
      '
39    If TRIM$(CODEPRO$) = "" Then CODEPRO$ = " libre "
      While Len(CODEPRO$) < LOMAXI: CODEPRO$ = "." + CODEPRO$ + ".": Wend
      CODEPRO$ = Left$(CODEPRO$, LOMAXI)
      If Left$(CODEPRO$, 1) = " " Then Mid$(CODEPRO$, 1, 1) = "."
      If Right$(CODEPRO$, 1) = " " Then CODEPRO$ = Left$(CODEPRO$, Len(CODEPRO$) - 1) + "."
      '
      TTXX$ = AJUSTI$(CODUBI$(JJ&), 16) + CODEPRO$ + " - " + DENUBI$(JJ&)
      Call GRAREG("!TAUBICA", TTXX$, JJ&)
      U& = JJ&
   Next JJ&
   '
99 Call SETULTREG("!TAUBICA", U&)
   Call FRESHECHO("!TAUBICA")
   Screen.MousePointer = 1
   '
End Sub

Function LOTSELEC$(CI%)
   '
   SELCODLOT07.Show 1
   '
   CI% = CODINT(SELCODLOT07.Text1)
   '
   If CI% > 0 Then
      If SELCODLOT07.Label28 <> "" Then
         LOTSELEC$ = SELCODLOT07.Label28
      End If
   End If
   '
End Function

Function FeVenGarant%(NUSERIE$)
     '
     FEGARANTIA% = 0
     '
     ' Devuelve la fecha que vence la garantia
     If TRIM$(NUSERIE$) = "" Then GoTo 99
     '
     FETANMUSE% = CVINT(VALOBADA("TANUMSE", "FevenGar", "NUMEROSERIE = '" & NUSERIE$ & "'", "NOMESS"))
     If FETANMUSE% > 0 Then
           FEGARANTIA% = FETANMUSE%
     Else
        FACTU$ = VALOBADA("TANUMSE", "NUMFACTURA", "NUMEROSERIE = '" & NUSERIE$ & "'", "NOMESS")
        CONDI$ = "CODICOM_LAR= '" & FACTU$ & "' AND TIPOMOVIM = 'FVEN'"
        FECAJABAN% = CVINT(VALOBADA("CAJABANC", "FECHEMISI", CONDI$, "NOMESS"))
        If FECAJABAN% > 0 Then
           FEGARANTIA% = DIASPOST(FECAJABAN%, 730)
        End If
     End If
99   FeVenGarant% = FEGARANTIA%

End Function

Function FeVenGarantRep%(NUSERIE$)
    '
    FEGARANTIRE% = 0
    '
    'Devuelve la fecha que vence la garantia por reparacion
    If TRIM$(NUSERIE$) = "" Then GoTo 99
    '
    CONDI$ = "NUMEROSERIE = '" & NUSERIE$ & "'ORDER BY FeVenGaRep DESC"
    FETANMUSE% = CVINT(VALOBADA("TANUMSE", "FeVenGaRep", CONDI$, "NOMESS"))
    If FETANMUSE% > 0 Then
       FEGARANTIRE = FETANMUSE%
    Else
       CONDI$ = "NUMEROSERIE = '" & NUSERIE$ & "'  ORDER BY FechEnt DESC"
       FECREPARA% = CVINT(VALOBADA("REPARA", "FechEnt", CONDI$, "NOMESS"))
       If FECREPARA% > 0 Then
         FEGARANTIRE% = DIASPOST(FECREPARA%, 180)
       End If
    End If
99  FeVenGarantRep% = FEGARANTIRE%
    '
End Function

Sub ASIGNALOTES(Archi$, RESUL$)
   'SELECCION DE LOTES A ASIGNAR AL PEDIDO
   RESUL$ = "APRO"
   On Error Resume Next
   ASIGLOT07.List1.Clear
   If Err Then
      On Error GoTo 0
      Call MENSERR(24, "Falta Formulario 'ASIGLOT07.FRM' en Proyecto '" + App.EXEName + "'.")
      Call FINAL("")
   End If
   On Error GoTo 0
   '
   ASIGLOT07.List2.Clear
   'AHP
   ESCARDEFAC% = 0: TODOSTIENLOTE% = 0
   For U& = 1 To ULTREG&(Archi$)
       XX$ = REGLEIDO$(Archi$, U&)
       If InStr(Archi$, "BOLREDET") > 0 Then
          CCII% = CVI(Mid$(XX$, 83, 2))
          SALPEND = CVD(Mid$(XX$, 89, 8))
        ElseIf InStr(Archi$, "CARDEFAC") > 0 Then
          CCII% = CVI(Mid$(XX$, 5, 2))
          SALPEND = CVS(Mid$(XX$, 1, 4))
          LOTEX1$ = TRIM$(Mid$(XX$, 267, 16))
          ESCARDEFAC% = 1 'AGREGADO LOTE AHP
       End If
       '
       TXTX$ = ""
       If CCII% > 0 And ESTRAZABLE%(CCII%) = 1 Then
          Call REPLA(TXTX$, CODEXT$(CCII%), 1, 16)
          Call REPLA(TXTX$, DESCRIT$(CCII%), 18, 33)
          Call REPLA(TXTX$, FORMATNUM$(SALPEND, "F10.1"), 54, 12)
          Call REPLA(TXTX$, TRIM$(FORMATNUM$(U&, "I5")), 80, 5)
          'AGREGADO LOTE AHP
          If ESCARDEFAC% > 0 Then 'AHP
             If LOTEX1$ <> "" Then
               TODOSTIENLOTE% = TODOSTIENLOTE% + 1
               Call REPLA(TXTX$, LOTEX1$ & ";", 85, 120)
               Call REPLA(TXTX$, FORMATNUM$(SALPEND, "F10.1"), 64, 10)
             End If
          End If
          'FIN AGREGADO LOTE AHP
          ASIGLOT07.List1.AddItem (TXTX$)
       End If
   Next U&
   '
   
   If ASIGLOT07.List1.ListCount > 0 Then
     'AGREGADO AHP SI TODOS TIENEN LOTE (VIENEN DE PEDIDO) SE ASIGNO RESPUESTA 'APRO' Y NO MUESTRO ASIGLOT07
     If TODOSTIENLOTE% = ASIGLOT07.List1.ListCount Or InStr(UCase$(EMPREAC$), "AHP") > 0 Then
      ASIGLOT07.VALIDA = "APRO"
     Else
      ASIGLOT07.Show 1
     End If
     RESUL$ = ASIGLOT07.VALIDA
   End If
   '
End Sub



Sub DETLOTREM(CI%, REG&)
    '
    If ESTRAZABLE%(CI%) = 1 Then
       TOTALOTE$ = ""
       For UY& = 0 To ASIGLOT07.List1.ListCount - 1
           YY$ = ASIGLOT07.List1.LIST(UY&)
           If XVALO(Mid$(YY$, 80, 5)) = REG& Then
               TOTALOTE$ = TRIM$(Mid$(YY$, 85))
           End If
       Next UY&
       '
5      NN% = InStr(TOTALOTE$, ";")
       If NN% > 0 Then
           LOTTE$ = Mid$(TOTALOTE$, 1, NN% - 1)
           '
           CONDI$ = "IDENLOTE = '" & LOTTE$ & "'"
           FEALT$ = (VALOBADA("LOTINGRE", "FECHALTA", CONDI$, "NOMESS"))
           FEEXX = CVINT(FEALT$)
           FEE$ = FECHATEX$(FEEXX)
           ANCHOO = XVALO(VALOBADA("LOTINGRE", "ANCHOMILI", CONDI$, "NOMESS"))
           CANTI = XVALO(VALOBADA("LOTINGRE", "CANALTA", CONDI$, "NOMESS"))
           TURNOP$ = VALOBADA("LOTINGRE", "TurnoProd", CONDI$, "NOMESS")
           '
           LOTTE$ = AJUSTI$(LOTTE$, 12) & " - " & TRIM$(FORMATNUM$(CANTI, "F10.1")) & " " & UNIMED$(CI%)
           If ANCHOO > 0 Then
             LOTTE$ = LOTTE$ & " - " & TRIM$(FORMATNUM$(ANCHOO, "F10.1")) & " mm"
           End If
           '
           LOTTE$ = LOTTE$ & " - " & FEE$
           If TRIM$(TURNOP$) <> "" Then
              LOTTE$ = LOTTE$ & " - " & TURNOP$
           End If
           '
           Call AGREGATEXTO(LOTTE$, 2)
           TOTALOTE$ = Mid$(TOTALOTE$, NN% + 1)
           GoTo 5
       End If
    End If
    '
End Sub

Function FRACMIN(CIIXI%)
    '
    CANMINI = ((1 / 10) ^ XVALO(CNTDC$(CIIXI%))) / 2
    FRACCC$ = TRIM$(FORMATNUM$(CANMINI, "F10." & XVALO(CNTDC$) + 1 & ""))
    FRACMIN = XVALO(FRACCC$)
    '
End Function

Sub ACTUBIPRE(UBIVIE$, UBINUE$, DEP%)   ' ACTULIZA EL ESTADO DE LAS UBICIONES SEGUN LOS MOVIMIENTO DE UBICACION
   '
   If UBIVIE$ <> "" Then
     CONDI$ = "UBICFIS = '" & UBIVIE$ & "' AND DEPOUBIC = " & DEP%
     REGS& = CantRegistros&("LOTINGRE", CONDI$) ' VERIFICA EN LOTINGRE SI EXSITEN ITEMS OCUPANDO DICHA UBICACION
     If REGS& <= 1 Then
       CONDI$ = "CodiUbic = '" & UBIVIE$ & "' AND DEPOUBIC = " & DEP%
       Call ACTUREGISTRO("UBIPRE", "STATUS", CONDI$, 0, REGS&, "NOMESS")  'SI NO EXISTEN ITEM EN LA POSICION, LA LIBERA
     End If
   End If
   '
   CONDI$ = "CodiUbic = '" & UBINUE$ & "' AND DEPOUBIC = " & DEP%
   Call ACTUREGISTRO("UBIPRE", "STATUS", CONDI$, 1, REGS&, "NOMESS") ' OCUPA LA NUEVA POSICION
   '
End Sub


Function VALIDEPOSI%(DEP%)
   '
   Vali% = 0
   If ECOARCH$("TADEPOSI", Chr(DEP%)) <> "" Then
      Vali% = 1
   End If
   '
   VALIDEPOSI% = Vali%
End Function

'*********************************************************
'!!!! AQUI ESTABA LA RUTINA VALIDASTOCK(Optional DepOrig%)
'!!!! LA MISMA SE MOVIO AL FORMULARIO FINFACTUR07
'*********************************************************

Function STOCK_PROV(NPROV&, CI1%, Optional VIA%)
    'SI TIENE PROBLEMAS DE VELOCIDAD VA A VER QUE REEMPLAZAR EL INNER JOIN POR RECORRER EL BOLRECEP
    ' DONDE PROVEEDOR = NPROV& Y COD_INTE = CI1% CADA CASO VER LOS MOVMIENTOS QUE TIENE EN MVISTO.
    sqlx$ = " SELECT  SUM(M.CANTSALID) AS SALID , SUM(M.CANTINGRE) AS INGRE "
    sqlx$ = sqlx$ + " FROM BOLRECEP B WITH(NOLOCK)  INNER JOIN MOVISTO  M WITH(NOLOCK) ON B.IDENLOTE = M.IDENLOTE "
    sqlx$ = sqlx$ + "  Where B.NUMPROV = " & NPROV&
    sqlx$ = sqlx$ + "  AND M.COD_INTE = " & CI1%
    If VIA% > 0 Then sqlx$ = sqlx$ + "  AND B.VIA_ENTREGA = " & VIA%
    sqlx$ = sqlx$ + "  GROUP BY M.COD_INTE, B.NUMPROV"
    Set RS = READSET(sqlx$)
    With RS
      ST = XVALO(!Ingre) - XVALO(!SALID)
    End With
    '
    STOCK_PROV = ST
    '
End Function
Function CANT_PEDIDA_PENDIENTE(CI1%)
     'CONSIDERO LA CANTIDAD PEDIDA PENDIENTE DE ENTREGA
     sqlx$ = " SELECT SUM(CanPed) AS CPED,  SUM(cantEnt) AS CENT FROM PEDDETA WITH(NOLOCK) "
     sqlx$ = sqlx$ + " WHERE CodiInt = " & CI1%
     sqlx$ = sqlx$ + " AND Status < 8 "   ' para que no este anulada ni cumplida
     sqlx$ = sqlx$ + " GROUP BY CodiInt "
     '
     Set RS = READSET(sqlx$)
     With RS
           CAPEN = XVALO(!CPED) - XVALO(!CENT)
           If CAPEN < 0 Then CAPEN = 0
     End With
     RS.Close
     Set RS = Nothing
     CANT_PEDIDA_PENDIENTE = CAPEN
End Function

Function CANT_RECEP_PENDIENTE_APROB(NPROV, CI1%)
     sqlx$ = " SELECT SUM(CANTNOM) AS CANOM, SUM(CanToApro) AS CAPRO FROM BOLRECEP WITH(NOLOCK) "
     sqlx$ = sqlx$ + " WHERE Status = 0 "
     sqlx$ = sqlx$ + " AND Cod_Inte = " & CI1%
     If NPROV > 0 Then sqlx$ = sqlx$ + " AND NumProv = " & NPROV
     sqlx$ = sqlx$ + " GROUP BY   COD_INTE "
     If NPROV > 0 Then sqlx$ = sqlx$ + ",  NumProv"
     '
     Set RS = READSET(sqlx$)
     With RS
          CARECEP = XVALO(!CANOM) - XVALO(!CAPRO)
     End With
     RS.Close
     Set RS = Nothing
     CANT_RECEP_PENDIENTE_APROB = CARECEP
End Function
Function COMPRA_PEND_DE_ENTR(NPROV, CI1%)
     '
     'SI SE LE PASA NPROV& = 0 NO TIENE EN CUENTA EL PROVEEDOR DEVUELVE TODO EL PENDIENTE DEL CODIGO
     sqlx$ = " SELECT cant_ocom, Cant_Rec, Coef_Unids FROM OCOMDETA WITH(NOLOCK) "
     sqlx$ = sqlx$ + " WHERE Cod_Inte = " & CI1% & " "
     If NPROV > 0 Then sqlx$ = sqlx$ + " AND Npro_Ocom = " & NPROV
     sqlx$ = sqlx$ + " AND Stat_Ocom < 8 "   ' para que no este anulada ni cumplida
     sqlx$ = sqlx$ + " AND  Stat_Ocom <> -1 " 'Para que no  hayan O/C's sin aprobar
     sqlx$ = sqlx$ + " AND Cant_Rec < (Cant_Ocom - 0.005) "
     '
     Set RS = READSET(sqlx$)
     With RS
         Do While Not .EOF
             COEFCAN = XVALO(!Coef_Unids)
             If COEFCAN < 0.005 Then COEFCAN = 1
             CACOM = XVALO(!cant_ocom) * COEFCAN
             CAREC = XVALO(!Cant_Rec) * COEFCAN
             CAPEN = CACOM - CAREC
             If CAPEN > 0 Then
               TOSUMIN = TOSUMIN + CAPEN
             End If
             .MoveNext
         Loop
     End With
     RS.Close
     Set RS = Nothing
     COMPRA_PEND_DE_ENTR = TOSUMIN
End Function
Sub CARGA_VECTORES_DE_MATERIALES_REMITIDOS()
'    Dim CODINTERNO%()
'    ReDim Preserve CODINTERNO%(NUMAS%)
    sqlx$ = "SELECT COD_INTE, SUM(CANTSALID) AS CANSAL FROM MOVISTO WITH(NOLOCK)"
    sqlx$ = sqlx$ + " WHERE CODIMOVI = 'RT'"
    sqlx$ = sqlx$ + " AND CANTSALID > 0 "
    sqlx$ = sqlx$ + " AND COD_INTE > 0 "
    sqlx$ = sqlx$ + " GROUP BY  COD_INTE"
    sqlx$ = sqlx$ + " ORDER BY COD_INTE"
    Set RS = READSET(sqlx$)
    With RS
     Do While Not .EOF
           CI1% = XVALO(!Cod_Inte)
           CANTISALE = XVALO(!CANSAL)
           Call AREXPLO(CI1%, 1, 1)
           For KKJ% = 1 To CAIT%
             CIMOVIM%(CI1%) = CIJ%(KKJ%)
             CAMOVIM(CI1%) = CAMOVIM(CI1%) + USOI(KKJ%) * CANTISALE
           Next KKJ%
           CAMOVIM(CI1%) = 0
           CI1% = 0
           .MoveNext
      Loop
    End With
    RS.Close
    Set RS = Nothing
    Erase CIMOVIM%, CAMOVIM
'    Call ACTUREGISTRO("MASTER", "STATUS", "CODINT = " & CI1%, 1, REGAF&, "NOMESS")
End Sub




