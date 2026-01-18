Attribute VB_Name = "MOVIST04"
Public CAITPE%, CIPE%(2048), CADESPA(2048), REPE&(2048)
Public NOARRAS%
Public CIMOVI%(32767)
'
Public CIMOVIM%(32767), CAMOVIM(32767), CAITMO%
'
Public CAMOVILO%
Public TIMOVILO%(512), CANINGLO#(512), CANSALLO#(512), CANRESLO#(512), FEMOVILO%(512), DEMOVILO$(512)
Public FORCAMPO$(128)

Function PUEDEMOVI%(DEPOMO%)
   ' La funcion revisa si en el deposito elegido es posible el movimiento de salida
   ' El movimiento es lícito si
   '   - no hay control de negativos
   '   - es un movimiento de entrada (CAMOVIM() > 0)
   '   - en todos los casos ABS(CAMOVIM) <= SALDEPOS(CI%,DEPOMO%)
   ' En los casos anteriores sobre PUEDEMOVI% devuelve '1'
   ' Caso contrario presenta lista de los codigos con stock insuficiente
   ' y devuelve '-1' sobre PUEDEMOVI%
   '
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
         If DESCOMPO%(CIMOVIM%(KKI%)) > 0 Then
           Call AREXPLO(CIMOVIM%(KKI%), 1, 1)
           For KKJ% = 1 To CAIT%
             CAITMO% = CAITMO% + 1
             CIMOVIM%(CAITMO%) = CIJ%(KKJ%)
             CAMOVIM(CAITMO%) = USOI(KKJ%) * CAMOVIM(KKI%)
           Next KKJ%
           CAMOVIM(KKI%) = 0
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
     j& = 0
     Y$ = REGBLAN("COSTNEG")
     For KKI% = 1 To CAITMO%
       If CAMOVIM(KKI%) < (-0.005) Then
         DEPMOI% = DEPOMO%: If DEPOMO% = 99 Then DEPMOI% = DEPOPRE%(CIMOVIM%(KKI%))
         'SI SACO LA MISMA CANTIDAD QUE TENGO NO ME DEJA SACARLO
         VUELTA% = 1
         Call VESARRDEP(CIMOVIM%(KKI%), DEPMOI%)   ' agregado por negativo de asembli
25       STDP = (STODEPOS(CIMOVIM%(KKI%), DEPMOI%))
         DIF1 = (STDP + CAMOVIM(KKI%))
         If DIF1 < (-0.005) Then
           If VUELTA% = 1 Then
              VUELTA% = 2
              Call VESARRAS(CIMOVIM%(KKI%))
              GoTo 25
           End If
           ' AGREGAR A LISTA DE CODIGOS NEGATIVOS
           '\\\OT-05-0730-WAR-09/11/05///
           CANTI2 = Abs(CAMOVIM(KKI%))
           DIF2 = Abs(STDP - CANTI2)
           Call REPLA(Y$, MKI$(CIMOVIM(KKI%)), 1, 2)
           Call REPLA(Y$, MKS$(CANTI2), 5, 4)
           Call REPLA(Y$, MKS$(STDP), 21, 4)
           Call REPLA(Y$, MKS$(DIF2), 25, 4)
           Call REPLA(Y$, MKI$(DEPMOI%), 29, 2)
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
     On Error Resume Next
     Kill LUCOM$ + "COSTONEG.WKF"
     On Error GoTo 0
     If j& > 0 Then 'SI HAY ALGUN NEGATIVO
     ' PRESENTAR MENSAJE ADVERTENCIA CON LIST DE LOS CODIGOS CON STOCK INSUFICIENTE
       Call COMUNI("Esta Transacción producirá Negativos\en el Déposito Elegido para\Algunos Items")
       DE2$ = ECOARCH("ECODEP", MKI$(DEPOMO%))
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
Sub MOVISTOK(FECHA%, CI%, LOTE$, CMOV$, DOPRILA$, DOPRICO$, DOSELA$, DOSECO$, REFE$, NUORIT%, VUNI#, MONEII%, NC%, DEPO%, CANTI, Optional NUPECLIEN&, Optional PRECONIVA#, Optional DELARGA$)
   '
   Static DECONSUM$    ' CTX%, REBLAN$, RELIBR$, UREMOVI&
   '
   Static COSREPOSI%

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
   '//OT-07-0464//
   If COSREPOSI% = 0 Then
       COSREPOSI% = 1    ' COSTO INFORMADO - COSPESOS
       If Control$("", "COSMOVI") = "COSTOCALCULADO" Then
          COSREPOSI% = 2 ' COSTO CALCULADO - COSCALCU
       End If
   End If
   '//OT-07-0464//
   SMP = Screen.MousePointer
   USERN% = USNBR% Mod 100
   If USERN% < 1 Then USERN% = 1
   Call CARDEPOS
   '
   'Call VERISAL(CI%)               ' verificar saldos cruzados
   '
   DEPOXI% = DEPO%
   If DECONSUM$ = "AUTOM" Then
     If CANTI < 0 Then
       DEPII% = DEPO%
       DEPOXI% = DEPOSIBLE%(CI%, CANTI, DEPII%)   'Obtengo el  deposito de movimiento.
     End If
   End If
   '
   CANTIX = (Int(10000 * CANTI + 0.5)) / 10000
   '
   If NOCLOSE% <> 1 Then Call APERBASDAT("STOCKS")
   '
   ' If FECHA% > HOII% + 30 Or FECHA% < HOII% - 100 Then FECHA% = HOII%
   ' COSTO UNITARIO (1) / CALCULADO (2)
   COSTOREPOS# = COSPESOS(CI%)
   If COSREPOSI% = 2 Then COSTOREPOS# = COSCALCU(CI%, 1)
   '
   With Movisto
     .AddNew
     !FECHMOV = CVDAT(FECHA%)
     !COD_INTE = CI%
     !Cod_Exte = Left$(CODEXT$(CI%), 24)
     !Descrip = Left$(DESCRIT0$(CI%), 64)
     !IdenLote = Left$(TRIM$(LOTE$), 16)
     !CodiMovi = Left$(CMOV$, 2)
     !DoPriCor = Left$(DOPRICO$, 12)
       DOPRILAX$ = DOPRILA$: If DOPRILAX$ = "" Then DOPRILAX$ = DOPRICO$
     !DOPRILAR = Left$(DOPRILAX$, 24)
     !DoSecCor = Left$(DOSECO$, 12)
       DOSECLAX$ = DOSELA$: If DOSECLAX$ = "" Then DOSECLAX$ = DOSECO$
     !DoSecLar = Left$(DOSECLAX$, 24)
     !NuorItem = NUORIT%
     !NuPedCli = 0
     !DESCRI_LAR = DELARGA$ ' DESCRIPCION DE TEXTO LIBRE
       If Left$(DOSECO$, 3) = "PC." Then
           !NuPedCli = XVALO(Mid$(DOSECO$, 4))
         Else
           On Error Resume Next
           !NuPedCli = NUPECLIEN
       End If
     '\\\OT-08-751-OD-12/09/08///
     !ReferCor = Left(REFE$, 64)
     '\\\OT-08-751-OD-FIN///
        If Left$(REFE$, 7) = "RECHAZO" Or Left$(REFE$, 8) = "FALTANTE" Then
          LOCON$ = AJUSTI$(LOTE$, 24) + FORMATNUM$(Abs(CANTIX), "F8.1")
          !LotConsum = LOCON$
        End If
     !CostoRep = COSTOREPOS#
     !VALOUNIT = VUNI#
     !MoneVal = MONEII%
       On Error Resume Next      ' POR SI NO ESTA DEFINIDO EL CAMPO
     !TipoCam = TICAMONE(MONEII%)
       On Error GoTo 0
     !DEPOMOVI = DEPOXI%
     If NC% >= 0 Then
         !Nume_Clie = NC%
         !Nume_Prov = 0
       Else
         !Nume_Clie = 0
         !Nume_Prov = Abs(NC%)
     End If
     If CANTIX >= 0 Then
         !CANTINGRE = CANTIX
         !CANTSALID = 0
       Else
         !CANTINGRE = 0
         !CANTSALID = Abs(CANTIX)
     End If
     !CantSaldo = 0
     !FeReMovi = Now
     !NumUsuar = USERN%
     !VALOUNITIVA = PRECONIVA#
     .Update
   End With
   '
   If NOCLOSE% <> 1 Then
     Movisto.Close
   End If
   '
   If NOARRAS% <> 1 Then
     Call VESARRDEP(CI%, DEPOXI%)
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
End Function
   
Sub VESARRDEP(CIXI%, NDEP%)
   '
   If CIXI% > 0 Then
     '
     SCMA = Screen.MousePointer
     Screen.MousePointer = 11
     '
     CARRAS# = 0
     '
     If NOCLOSE% <> 1 Then Call APERBASDAT("STOCKS")
     SQLX$ = " SELECT * FROM MOVISTO "
     SQLX$ = SQLX$ + "WHERE Cod_Inte = " & CIXI% & " "
     SQLX$ = SQLX$ + "AND DepoMovi = " & NDEP% & " "
     '
     Set MoviTemp = Stocks.OpenRecordset(SQLX$, dbOpenSnapshot)
     CAMOVI& = 0
     With MoviTemp
       On Error Resume Next
       .MoveFirst
       If Err < 1 Then
         Do While Not .EOF
           CAMOVI& = CAMOVI& + 1
           CARRAS# = CARRAS# + !CANTINGRE - !CANTSALID
         .MoveNext
         Loop
       End If
     End With
     '
     COLUX% = 1
     For KI% = 1 To NUSU%
       If NDEP% = COSU%(KI%) Then
         COLUX% = KI%
         Exit For
       End If
     Next KI%
     '
     With SalStock
       .FindFirst "CodiInte = " & CIXI%
       If .NoMatch = False Then
           .Edit
           VESAL$ = !VeSalSto
         Else
           .AddNew
           !CodiInte = CIXI%
           VESAL$ = String$(256, 0)
       End If
       Mid$(VESAL$, 4 * COLUX% - 3, 4) = MKS$(CARRAS#)
       !VeSalSto = VESAL$
       .Update
     End With
     If NOCLOSE% <> 1 Then SalStock.Close
     '
   End If
   '
99 End Sub
'
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
               '
               COD$ = LTrim$(RTrim$(CODEXT$(CI1%)))
               REFE$ = "Vale Mat.de Dep." + TRIM$(Str$(DP1%))
               CANTI = CAN
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
Sub NUENTRECONSI(NPROVE%)
        '
        '\\\OT-5-861-FG-09/12/05///
        Dim CANPREASI()
        '\\\OT-5-861-FG-FIN///
        '
        LLAMACOMPRA% = 0
        If NPROVE% > 0 Then LLAMACOMPRA% = 1
        '
        FORIPRE$ = TRIM$(Control$("", "FOREMICO"))
        If FORIPRE$ <> "" Then
            LU$ = LUDAT$
            If EXISTE%(LU$ + FORIPRE$ + ".FRM") <> 1 Then
                RMCC$ = REPLACAR$(LU$ + FORIPRE$ + ".FRM", "\", "/")
                Call COMUNI("Error de Configuración:\  \No Existe Formulario de Impresión\de Remito en Consignación:\  \'" + RMCC$ + "'.")
                GoTo 9999
            End If
        End If
        '
        DepoCon% = XVALO(Control$("", "DEPOCONS"))
        DEPORIG% = DepoCon%
        DEPOSI$ = ECOARCH$("TADEPOSI", Chr$(DEPORIG%))
        '
        'Call BLIMESS("Revisando Archivo de Movimientos de Stock")
1100    Screen.MousePointer = 11
        CMO$ = TRIM$(Control$("", "CODRECON"))
        If CMO$ = "" Then CMO$ = "RC"
        '
        Call APERBASDAT("STOCKS")
        SQLX$ = "SELECT * FROM MOVISTO "
        SQLX$ = SQLX$ + " WHERE CODIMOVI='RC' "
        Set UNRECONTMP = Stocks.OpenRecordset(SQLX$, 4)
        '
        With UNRECONTMP
            On Error Resume Next
            .MoveFirst
            If Err Then
                  UNRO& = 0
            Else
                Max = XVALO(Mid$(!DoPriCor, 4))
                Do While .EOF = False
                  If XVALO(Mid$(!DoPriCor, 4)) > Max Then
                    Max = XVALO(Mid$(!DoPriCor, 4))
                  End If
                  .MoveNext
                Loop
                UNRO& = Max
            End If
        End With
        '
1002    REMI$ = CMO$ + "-" + TRIM$(Str$(UNRO& + 1))
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
        Call REPLA(Y$, Chr$(DEPORIG%), 121, 1)
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
        With UNRECONTMP
          .FindFirst "DoPriCor = '" & REMI$ & "'"
          If .NoMatch = False Then
            Call MENSERR(24, "Número de Remito ya Registrado.")
            GoTo 1111
          End If
        End With
        '
        DEPORIG% = Asc(Mid$(PEDY$, 121, 1))
        If DepoCon% > 0 Then
          If DEPORIG% < 1 Or ECOARCH$("TADEPOSI", Chr$(DEPORIG%)) = "" Then
            Call MENSERR(24, "Depósito de Salida no Válido")
            DEPORIG% = XVALO(Control$("", "DEPOCONS"))
            GoTo 1110
          End If
        End If
        DEPOSI$ = ECOARCH$("TADEPOSI", Chr$(DEPORIG%))
        '
1112    Call CIERRARCH("*.*")
        Call BLANARCH("!BOLREDET")
        '
        Screen.MousePointer = 11
        HAYOC% = 0
        Call APERBASDAT("STOCKS")
        SQLX$ = "SELECT * FROM MACONSIG "
        SQLX$ = SQLX$ + " WHERE (Cant_Asig - Cant_Entreg ) > 0.005 "
        SQLX$ = SQLX$ + " AND Npro_Ocom = " & NC% & " "
        Set MACONSITMP = Stocks.OpenRecordset(SQLX$, 4)
        With MACONSITMP
         On Error Resume Next
         .MoveFirst
         If Err Then
            On Error GoTo 0
            GoTo 555
         Else
            On Error GoTo 0
          Do While .EOF = False
            CAPENE# = XVALO(!Cant_Asig) - XVALO(!Cant_Entreg)
            NUOC& = XVALO(Mid$(!Nro_Ocom, 4, 6))
            NORCOMPRA$ = SAFETEXT(!Nro_Ocom)
            If (STATOCOM%(NUOC&) < 8 And STATOCOM%(NUOC&) >= 0) Then
               CI% = !COD_INTE
               'STODISPO# = CDbl(STOCKACT(CI%))      CODIGO ANTERIOR !!!
               '\\\OT-5-861-FG-9/12/05///
               STODISPO# = CDbl(STOCKACT(CI%) - CANPREASI(CI%))
               '\\\OT-5-861-FG-FIN///
               If STODISPO# < 0 Then STODISPO# = 0
               If STODISPO# < CAPENE# Then CAPENE# = STODISPO#
               '
               Y$ = REGBLAN$("!BOLREDET")
               Call REPLA(Y$, MKI$(CI%), 83, 2)
               Call REPLA(Y$, UNIMED$(CI%), 85, 4)
               Call REPLA(Y$, MKD$(CAPENE#), 89, 8)
               Call REPLA(Y$, MKS$(NUOC&), 151, 4)
               Call REPLA(Y$, MKS$(I&), 189, 4)
               '
               Call REPLA(Y$, NORCOMPRA$, 119, 16)     'agrega el numero de la orden de compra al BOLREDET
               '
               HAYOC% = 1
               Call REGAPP("!BOLREDET", Y$)
               '\\\OT-5-861-FG-9/12/05///
               ' ACUMULA LO PRE-ASIGNADO POR CASO DE ITEMS QUE APARECEN VARIAS VECES
               If CI% > 0 Then
                 If CI% <= NUMAS% Then
                   CANPREASI(CI%) = CANPREASI(CI%) + CAPENE#
                 End If
               End If
            End If
            '\\\OT-5-861-FG-FIN///
            .MoveNext
          Loop
          End If
        End With
        '
555     If HAYOC% = 0 Then
            Call COMUNI("No hay Consignaciones Pendientes\para este Proveedor.")
        End If
        '
1113    VTB% = VENTABU%("ECONSIG/NC/TITUPAN=Consignación de Materiales - " + TRIM$(RASOCLI$(NC1%)))
        If VTB% < 0 Or ULTREG&("!BOLREDET") < 1 Then
            X$ = "": Y$ = "": YY$ = ""
            EBOL$ = "": PEDY$ = "": DESCRI$ = ""
            GoTo 1105
        End If
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
          If PUEDEMOVI%(DEPORIG%) < 0 Then GoTo 1113 ' CANCELA SALIDA
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
          AINDI% = 0
          NUITEM = 0
          For YX& = 1 To ULTREG&("!BOLREDET")
             YY$ = REGLEIDO$("!BOLREDET", YX&)
             CI% = CVI(Mid$(YY$, 83, 2))
             CAN# = CVD(Mid$(YY$, 89, 8))
             LOTE$ = Mid$(YY$, 69, 12)
             '
             NUMECOMPRA$ = TRIM$(Mid$(YY$, 119, 16))    ' numero de orden de compra
             '
             If CI% > 0 Then
               If CI% <= NUMAS% Then
                 If CAN# > 0.005 Then
                    NUOC& = CVS(Mid$(YY$, 151, 4))
                    DOSEC$ = ""
                    With MaConsig
                        'SQLX$ = " VAL(Mid$(Nro_Ocom, 4, 6)) =  " & NUOC & " " '\\\OT-06-0206-WAR-29/05/06///
                        SQLX$ = " Nro_Ocom = '" & NUMECOMPRA$ & "' "      ' compara con el numero de la orden de compra y numero de item
                        SQLX$ = SQLX$ + " AND Cod_Inte = " & CI% & ""
                        SQLX$ = SQLX$ + " AND Npro_Ocom = " & NC% & " "
                        .FindFirst SQLX$
                        If .NoMatch = False Then
                            .Edit
                            DOSEC$ = !Nro_Ocom
                            TOENTRE# = !Cant_Entreg + CAN#
                            !Cant_Entreg = TOENTRE#
                            If FF% > CVINT(!Fecha_Entreg) Then
                               !Fecha_Entreg = CVDAT(FF%)
                            End If
                            !NRemi_Consig = TRIM$(Mid$(REMI$, 4))
                            .Update
                            GoTo 1120
                        End If
                    End With
                  End If
               End If
             End If
             '
             With MaConsig
                .AddNew
                  !Femi_Ocom = CVDAT(FF%)
                  !Npro_Ocom = NC%
                  !RaSo_Prov = RASOCLI(NC% * (-1))
                   NOCOM$ = TRIM$(Str$(NUOC&))
                   While Len(NOCOM$) < 6: NOCOM$ = "0" + NOCOM$: Wend
                   NOCOM$ = "OC-" + Left$(NOCOM$, 6) + "-00"
                  !Nro_Ocom = NOCOM$
                  !COD_INTE = CI%
                  !Cod_Exte = CODEXT(CI%)
                  !NRemi_Consig = TRIM$(Mid$(REMI$, 4))
                  !Unidad = UNIMED$(CI%)
                  !Cant_Asig = CAN#
                  !Cant_Entreg = CAN#
                  !Cant_Rendid = 0
                  !Fecha_Entreg = CVDAT(FF%)
                .Update
             End With
             '
1120         CANTI = (-1) * CAN#
             REFE$ = REMI$
             NUITEM = NUITEM + 1
             NUORIT% = NUITEM
             'Call MovistoK(FF%, CI%, LOTE$, CMO$, REMI$, DOSEC$, RASOCLI$(NC1%), 0, "$", NC1%, DEPORIG%, CANTI)
             Call MOVISTOK(FF%, CI%, LOTE$, CMO$, REMI$, REMI$, DOSEC$, DOSEC$, REFE$, NUORIT%, 0, 0, NC1%, DEPORIG%, CANTI)
             '
          Next YX&
          Call CIERRARCH("MACONSIG")
          '
          Call FRESHECHO("IREPRIPE")
          Screen.MousePointer = 1
          '
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
            CAPARDOC% = 8
            '
            CODTABU1$(1) = "COD-MAT"
            CODTABU1$(2) = "DES-MAT"
            CODTABU1$(3) = "CANTIDAD"
            CODTABU1$(4) = "CAKILOS"
            CODTABU1$(5) = "UNIMED"
            CODTABU1$(6) = "N-OCOMPR"
            CACOLTAB1% = 6
            '
            DESTIDOC% = NC1%
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
                        TETABU1$(6, CAITAB1%) = NUOCO$
                        '
                     End If
                  End If
               End If
            Next YX&
            '
            Call PRINTDOC("FOREMICO")   ' Remito en Consignacion
            '
          End If
          '
1199    End If
        '
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
        '
         
End Sub
'Sub ENTRECONSI(NPROVE%)
        '/////***RUTINA ANTIGUA****////
'        LLAMACOMPRA% = 0
'        If NPROVE% > 0 Then LLAMACOMPRA% = 1
'        '
'        FORIPRE$ = TRIM$(CONTROL$("", "FOREMICO"))
'        If FORIPRE$ <> "" Then
'            LU$ = LUDAT$
'            If EXISTE%(LU$ + FORIPRE$ + ".FRM") <> 1 Then
'                RMCC$ = REPLACAR$(LU$ + FORIPRE$ + ".FRM", "\", "/")
'                Call COMUNI("Error de Configuración:\  \No Existe Formulario de Impresión\de Remito en Consignación:\  \'" + RMCC$ + "'.")
'                GoTo 9999
'            End If
'        End If
'        '
'        DEPOCON% = XVALO(CONTROL$("", "DEPOCONS"))
'        DEPORIG% = DEPOCON%
'        DEPOSI$ = ECOARCH$("TADEPOSI", Chr$(DEPORIG%))
'        '
'        'Call BLIMESS("Revisando Archivo de Movimientos de Stock")
'1100    Screen.MousePointer = 11
'        CMO$ = TRIM$(CONTROL$("", "CODRECON"))
'        If CMO$ = "" Then CMO$ = "RC"
'        FIN& = ULTREG&("MOVISTO")
'        For U& = FIN& To 1 Step -1
'            Call TRACE(20, FIN& - U& + 1, FIN&)
'            XX$ = REGLEIDO$("MOVISTO", U&)
'            If UCase$(Mid$(XX$, 21, 2)) = CMO$ Then
'                UNRO& = XVALO(Mid$(XX$, 26, 7))
'                GoTo 1002
'            End If
'        Next U&
'        UNRO& = 0
'        '
'1002    REMI$ = CMO$ + "-" + TRIM$(Str$(UNRO& + 1))
'        Call FRESHALL
'        '
'1105    Call CIERRARCH("*.*")
'        '
'        Screen.MousePointer = 1
'        If NPROVE% > 0 Then
'             EBOL$ = MKI$(NPROVE%)
'          Else
'             EBOL$ = OBJPLA$("SELPROVE", "")
'             If EBOL$ = "" Then
'                GoTo 9999
'             End If
'        End If
'        '
'        NC% = CVI(EBOL$)
'        If NC% < 1 Or ECOARCH$("PROVED1", MKI$(NC%)) = "" Then
'            Call MENSERR(24, "Proveedor Inexistente")
'            NPROVE% = 0
'            GoTo 1105
'        End If
'        '
'        NC1% = (-1) * NC%
'        '
'        Y$ = PEDY$: If Y$ = "" Then Y$ = Space$(128)
'        Call REPLA(Y$, MKS$(NROBOL&), 1, 4)
'        Call REPLA(Y$, EBOL$, 7, 2)
'        Call REPLA(Y$, MKI$(HOY(HO$)), 5, 2)
'        Call REPLA(Y$, RASOCLI$(NC1%), 9, 32)
'        Call REPLA(Y$, DOMICLI$(NC1%), 41, 32)
'        CLOCA$ = TRIM$(CPOSCLI$(NC1%)): If CLOCA$ <> "" Then CLOCA$ = CLOCA$ + " "
'        Call REPLA(Y$, CLOCA$ + LOCACLI$(NC1%), 73, 32)
'1110    Call REPLA(Y$, REMI$, 105, 16)
'        Call REPLA(Y$, Chr$(DEPORIG%), 121, 1)
'        '
'1111    PEDY$ = OBJPLA$("REMICON-ENCB", Y$)
'        If PEDY$ = "" Then GoTo 9998
'        FF% = CVI(Mid$(PEDY$, 5, 2))
'        REMI$ = TRIM$(Mid$(PEDY$, 105, 16))
'        If TRIM$(REMI$) = "" Then
'            Y$ = PEDY$
'            Call MENSERR(24, "Falta Numero de Remito")
'            GoTo 1111
'        End If
'        '
'        DEPORIG% = Asc(Mid$(PEDY$, 121, 1))
'        If DEPOCON% > 0 Then
'          If DEPORIG% < 1 Or ECOARCH$("TADEPOSI", Chr$(DEPORIG%)) = "" Then
'            Call MENSERR(24, "Depósito de Salida no Válido")
'            DEPORIG% = XVALO(CONTROL$("", "DEPOCONS"))
'            GoTo 1110
'          End If
'        End If
'        DEPOSI$ = ECOARCH$("TADEPOSI", Chr$(DEPORIG%))
'        '
'1112    Call CIERRARCH("*.*")
'        Call BLANARCH("!BOLREDET")
'        '
'        Screen.MousePointer = 11
'        HAYOC% = 0
'        For I& = 1 To ULTREG&("MACONSIG")
'            X$ = REGLEIDO$("MACONSIG", I&)
'            If CVI(Mid$(X$, 3, 2)) = NC% Then
'                CAPENE# = CVD(Mid$(X$, 55, 8)) - CVD(Mid$(X$, 63, 8))
'                If CAPENE# > 0 Then
'                    NUOC& = XVALO(Mid$(X$, 24, 6))
'                    CI% = CVI(Mid$(X$, 33, 2))
'                    STODISPO# = CDbl(STOCKACT(CI%))
'                    If STODISPO# < 0 Then STODISPO# = 0
'                    If STODISPO# < CAPENE# Then CAPENE# = STODISPO#
'                    '
'                    Y$ = REGBLAN$("!BOLREDET")
'                    Call REPLA(Y$, MKI$(CI%), 83, 2)
'                    Call REPLA(Y$, UNIMED$(CI%), 85, 4)
'                    Call REPLA(Y$, MKD$(CAPENE#), 89, 8)
'                    Call REPLA(Y$, MKS$(NUOC&), 151, 4)
'                    Call REPLA(Y$, MKS$(I&), 189, 4)
'                    HAYOC% = 1
'                    Call REGAPP("!BOLREDET", Y$)
'                End If
'            End If
'        Next I&
'        '
'        If HAYOC% = 0 Then
'            Call COMUNI("No hay Consignaciones Pendientes\para este Proveedor.")
'        End If
'        '
'1113    VTB% = VENTABU%("ECONSIG/NC/TITUPAN=Consignación de Materiales - " + TRIM$(RASOCLI$(NC1%)))
'        If VTB% < 0 Or ULTREG&("!BOLREDET") < 1 Then
'            X$ = "": Y$ = "": YY$ = ""
'            EBOL$ = "": PEDY$ = "": DESCRI$ = ""
'            GoTo 1105
'        End If
'        '
'        For YX& = 1 To ULTREG&("!BOLREDET")
'          YY$ = REGLEIDO$("!BOLREDET", YX&)
'          CI% = CVI(Mid$(YY$, 83, 2))
'          CAN# = CVD(Mid$(YY$, 89, 8))
'          If CI% > 0 Then
'            If CI% <= NUMAS% Then
'              If CAN# > 0.005 Then
'                GoTo 1114
'              End If
'            End If
'          End If
'        Next YX&
'        Call MENSERR(24, "Consignación Cancelada !!!\Faltan Cantidades o Códigos.")
'        GoTo 1113
'        '
'1114    If TRIM$(FORIPRE$) <> "" Then
'           DTP$ = OBJPLA$("TRANSPOR", DATRANSPO$)
'           If DTP$ = "" Then GoTo 1113         ' CANCELADO
'           DATRANSPO$ = DTP$
'        End If
'        '
'11141   CONTA% = ULTREG&("!BOLREDET")
'        '
'1115    If CONTA% > 0 Then
'          '
'          Call CIERRARCH("MOVISTO")
'          Call BLOQARCH("MOVISTO")
'          Call BLOQARCH("MACONSIG")
'          Screen.MousePointer = 11  'Call BLIMESS("Registrando Recepcion ...")
'          '
'          AINDI% = 0
'          For YX& = 1 To ULTREG&("!BOLREDET")
'             YY$ = REGLEIDO$("!BOLREDET", YX&)
'             CI% = CVI(Mid$(YY$, 83, 2))
'             CAN# = CVD(Mid$(YY$, 89, 8))
'             LOTE$ = Mid$(YY$, 69, 12)
'             If CI% > 0 Then
'               If CI% <= NUMAS% Then
'                 If CAN# > 0.005 Then
'                    NUOC& = CVS(Mid$(YY$, 151, 4))
'                    REGOC& = CVS(Mid$(YY$, 189, 4))
'                    DOSEC$ = ""
'                    If REGOC& > 0 Then
'                       If REGOC& <= ULTREG&("MACONSIG") Then
'                          ZZ$ = REGLEIDO$("MACONSIG", REGOC&)
'                          If CVI(Mid$(ZZ$, 3, 2)) = NC% Then
'                             If CVI(Mid$(ZZ$, 33, 2)) = CI% Then
'                                DOSEC$ = Mid$(ZZ$, 21, 12)
'                                TOENTRE# = CVD(Mid$(ZZ$, 63, 8)) + CAN#
'                                Call REPLA(ZZ$, MKD$(TOENTRE#), 63, 8)
'                                If FF% > CVI(Mid$(ZZ$, 71, 2)) Then
'                                    Call REPLA(ZZ$, MKI$(FF%), 71, 2)
'                                End If
'                                Call GRAREG("MACONSIG", ZZ$, REGOC&)
'                                GoTo 1120
'                             End If
'                          End If
'                       End If
'                    End If
'                    '
'                    ZZ$ = REGBLAN$("MACONSIG")
'                    Call REPLA(ZZ$, MKI$(FF%), 1, 2)
'                    Call REPLA(ZZ$, MKI$(NC%), 3, 2)
'                       NOCOM$ = TRIM$(Str$(NUOC&))
'                       While Len(NOCOM$) < 6: NOCOM$ = "0" + NOCOM$: Wend
'                       NOCOM$ = "OC-" + Left$(NOCOM$, 6) + "-00"
'                    Call REPLA(ZZ$, NOCOM$, 21, 12)
'                    Call REPLA(ZZ$, MKI$(CI%), 33, 2)
'                    Call REPLA(ZZ$, UNIMED$(CI%), 51, 4)
'                    Call REPLA(ZZ$, MKD$(CAN#), 55, 8)
'                    Call REPLA(ZZ$, MKD$(CAN#), 63, 8)
'                    Call REPLA(ZZ$, MKI$(FF%), 71, 2)
'                    Call REGAPP("MACONSIG", ZZ$)
'                    '
'1120                CANTI = (-1) * CAN#
'                    'Call MovistoK(FF%, CI%, LOTE$, CMO$, REMI$, DOSEC$, RASOCLI$(NC1%), 0, "$", NC1%, DEPORIG%, CANTI)
'                 End If
'               End If
'             End If
'             '
'          Next YX&
'          Call CIERRARCH("MOVISTO")
'          Call CIERRARCH("MACONSIG")
'          '
'          Call FRESHECHO("IREPRIPE")
'          Screen.MousePointer = 1
'          '
'          If TRIM$(FORIPRE$) <> "" Then
'            '
'            FEX = FF%
'            FECHDOC$ = FECHATEX$(FEX)
'            NUMEDOC$ = TRIM$(REMI$)
'            TIPODOC$ = "Remito de Traslado"
'            '
'            CODPARAM$(1) = "TRANSPOR": DOCPARAM$(1) = TRIM$(Mid$(DTP$, 1, 32))
'            CODPARAM$(2) = "DOMI-TRA": DOCPARAM$(2) = TRIM$(Mid$(DTP$, 33, 32))
'            CODPARAM$(3) = "CUIT-TRA": DOCPARAM$(3) = TRIM$(Mid$(DTP$, 65, 16))
'            CODPARAM$(4) = "PATE-TRA": DOCPARAM$(4) = TRIM$(Mid$(DTP$, 81, 16))
'            CODPARAM$(5) = "COND-TRA": DOCPARAM$(5) = TRIM$(Mid$(DTP$, 97, 32))
'            CODPARAM$(6) = "DOCUCOND": DOCPARAM$(6) = TRIM$(Mid$(DTP$, 129, 24))
'            CODPARAM$(7) = "LOCA-TRA": DOCPARAM$(7) = TRIM$(Mid$(DTP$, 159, 32))
'            CODPARAM$(8) = "NUBULTOS": DOCPARAM$(8) = TRIM$(Mid$(DTP$, 153, 6))
'            CAPARDOC% = 8
'            '
'            CODTABU1$(1) = "COD-MAT"
'            CODTABU1$(2) = "DES-MAT"
'            CODTABU1$(3) = "CANTIDAD"
'            CODTABU1$(4) = "CAKILOS"
'            CODTABU1$(5) = "UNIMED"
'            CODTABU1$(6) = "N-OCOMPR"
'            CACOLTAB1% = 6
'            '
'            DESTIDOC% = NC1%
'            CAITAB1% = 0
'            '
'            For YX& = 1 To ULTREG&("!BOLREDET")
'               YY$ = REGLEIDO$("!BOLREDET", YX&)
'               CI% = CVI(Mid$(YY$, 83, 2))
'               CAN# = CVD(Mid$(YY$, 89, 8))
'               If CI% > 0 Then
'                  If CI% <= NUMAS% Then
'                     If CAN# > 0.005 Then
'                        UNID$ = TRIM$(Mid$(YY$, 85, 4))
'                        If UNID$ = "" Then UNID$ = UNIMED$(CI%)
'                        CANS$ = CSTRING$(MKD$(CAN#), 3, 9, 1, 2)
'                        KILO$ = CSTRING$(MKD$(CAN# * PESUNI(CI%)), 3, 12, 3, 2)
'                        NULOTE$ = TRIM$(Mid$(YY$, 69, 12))
'                        NUOCO$ = CSTRING$(Mid$(YY$, 151, 4), 3, 7, 0, 2)
'                        '
'                        CAITAB1% = CAITAB1% + 1
'                        TETABU1$(1, CAITAB1%) = CODEXT$(CI%)
'                        TETABU1$(2, CAITAB1%) = DESCRIT$(CI%)
'                        TETABU1$(3, CAITAB1%) = CANS$
'                        TETABU1$(4, CAITAB1%) = KILO$
'                        TETABU1$(5, CAITAB1%) = UNID$
'                        TETABU1$(6, CAITAB1%) = NUOCO$
'                        '
'                     End If
'                  End If
'               End If
'            Next YX&
'            '
'            Call PRINTDOC("FOREMICO")   ' Remito en Consignacion
'            '
'          End If
'          '
'1199    End If
'        '
'9998    Call CIERRARCH("*.*")
'        '
'        X$ = "": Y$ = "": YY$ = ""
'        EBOL$ = "": PEDY$ = "": DESCRI$ = ""
'        '
'        If LLAMACOMPRA% = 0 Then
'           GoTo 1100
'        End If
'        '
'9999    Call CIERRARCH("*.*")
'        '
'End Sub
''

Sub DESCACONSI(NCX%, cix%, CANX#, NUOCO&, Optional MODODESCA$)
    '
    CAREN# = 0
    CAPEN# = CANX#
    MODES$ = UCase$(SAFETEXT$(MODODESCA$))
    '
    Call APERBASDAT("STOCKS")
    SQLX$ = "Npro_Ocom = " & NCX% & " "
    SQLX$ = SQLX$ + " AND Cant_Rendid < Cant_Entreg - 0.001 "
     
    With MaConsig
      VUELTA% = 1
8     .FindFirst SQLX$
10    If .NoMatch = False Then
        If !COD_INTE = cix% Then
          NUOCA& = XVALO(Mid$(!Nro_Ocom, 4, 6))
          If NUOCA& = NUOCO& Or (NUOCA& = 0 And VUELTA% = 2) Then ' coincide o/c o bien consignacion sin o/c
            CACON# = !Cant_Entreg - !Cant_Rendid
              If CACON# < 0 Then CACON# = 0
            CAREN# = CAPEN#
              If CAREN# > CACON# Then CAREN# = CACON#
            CAPEN# = CAPEN# - CAREN#
            .Edit
            If MODES$ = "DEVOLU" Then
                 !Cant_Entreg = !Cant_Entreg - CAREN#
               Else
                 !Cant_Rendid = !Cant_Rendid + CAREN#
            End If
            !Saldo_Pend = !Cant_Entreg - !Cant_Rendid
            .Update
          End If
        End If
        .FindNext SQLX$
        GoTo 10
      End If
      If CAPEN# > 0.005 Then
        If VUELTA% = 1 Then
          VUELTA% = 2
          GoTo 8
        End If
      End If
    End With
    '
    If CAPEN# >= 0.05 Then
      CODDX$ = CODEXT$(cix%)
      CAPENDE = CAPEN#
      '
      Call LEEEXPLO(CODDX$, CAPENDE)
      For j% = 1 To CAIT%
        CIY% = CIJ%(j%)
        CANY# = CAPEN# * CDbl(USOI(j%))
        '
        With MaConsig
          VUELTA% = 1
18        .FindFirst SQLX$
20        If .NoMatch = False Then
            If !COD_INTE = CIY% Then
              NUOCA& = XVALO(Mid$(!Nro_Ocom, 4, 6))
              If NUOCA& = NUOCO& Or (NUOCA& = 0 And VUELTA% = 2) Then ' coincide o/c o bien consignacion sin o/c
                CACON# = !Cant_Entreg - !Cant_Rendid
                  If CACON# < 0 Then CACON# = 0
                CAREN# = CANY#
                  If CAREN# > CACON# Then CAREN# = CACON#
                .Edit
                !Cant_Rendid = !Cant_Rendid + CAREN#
                !Saldo_Pend = !Cant_Entreg - !Cant_Rendid
                .Update
                CANY# = CANY# - CAREN#
              End If
            End If
            .FindNext SQLX$
            GoTo 20
          End If
          If CANY# > 0.005 Then
            If VUELTA% = 1 Then
              VUELTA% = 2
              GoTo 18
            End If
          End If
        End With
        '
      Next j%
    End If
    '
End Sub
'\\\OT-05-0419-WAR-26/06/05///
Sub NUREMIFA(NROPRO&, Optional NUPUVEN$) 'NUMERO DE PUNTO DE VENTA
    '
    Static CMO$, FLIM%
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
    'NUEVO
    'Flid# = CDbl(CVDAT(FLIM%))
    'FLI1 = CVDAT(FLIM%)
    Call APERBASDAT("STOCKS")
    SQLX$ = "SELECT * FROM MOVISTO "
    'SQLX$ = SQLX$ + " WHERE CDBL(FechMov) >= " & Flid# & " "
    SQLX$ = SQLX$ + " WHERE FechMov >= " & FLIMI & " "
    SQLX$ = SQLX$ + " AND CodiMovi = '" & CMO$ & "' "
    SQLX$ = SQLX$ + " AND DoPriLar LIKE '" & CMO$ + "*' "
    SQLX$ = SQLX$ + " AND DoPriLar <> '' "
    If TRIM(NUPUVEN$) <> "" Then
      SQLX$ = SQLX$ + " AND VAL(MID$(DoPriLar,6,4)) =  " & XVALO(NUPUVEN$) & " "
    End If
    SQLX$ = SQLX$ + "ORDER BY VAL(MID$(DoPriLar,11)) DESC"
    '
    Set MOVITMP = Stocks.OpenRecordset(SQLX$, dbOpenSnapshot)
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
      NROMA& = XVALO(Mid(!DOPRILAR, 11, 8))
      '
      If NROPRO& > 0 Then
        Do While Not .EOF
          If XVALO(Mid$(!DOPRILAR, 11, 8)) = NROPRO& Then
            NROPRO& = 0
            GoTo 84
          End If
          If XVALO(Mid$(!DOPRILAR, 11, 8)) < NROPRO& Then
            GoTo 84
          End If
        .MoveNext
        Loop
      End If
    End With
    '
84  If NROMA& < 0 Then NROMA& = 0
85  If NROPRO& < 1 Then NROPRO& = 1 + NROMA&
99  Screen.MousePointer = 1
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
   DEPORIG% = DEPX%
   DEPOSI$ = ECOARCH$("TADEPOSI", Chr$(DEPORIG%))
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
   CODPARAM$(7) = "NUBULTOS": DOCPARAM$(7) = TRIM$(Mid$(DTP$, 153, 6))
   CAPARDOC% = 7
   '
   CODTABU1$(1) = "COD-MAT"
   CODTABU1$(2) = "DES-MAT"
   CODTABU1$(3) = "CANTIDAD"
   CODTABU1$(4) = "CAKILOS"
   CACOLTAB1% = 4
   '
   AINDI% = 0
   CAITAB1% = 0
   For YX& = 1 To ULTREG&("!BOLREDET")
      YY$ = REGLEIDO$("!BOLREDET", YX&)
      CI% = CVI(Mid$(YY$, 83, 2))
      CAN# = CVD(Mid$(YY$, 171, 8))
      LOTE$ = Mid$(YY$, 69, 12)
      If CI% > 0 Then
         If CI% <= NUMAS% Then
            If CAN# > 0.005 Then
               UNID$ = UNIMED$(CI%)
               CANS$ = CSTRING$(MKD$(CAN#), 3, 9, 1, 2)
               KILO$ = CSTRING$(MKD$(CAN# * PESUNI(CI%)), 3, 12, 3, 2)
               '
               CAITAB1% = CAITAB1% + 1
               TETABU1$(1, CAITAB1%) = CODEXT$(CI%)
               TETABU1$(2, CAITAB1%) = DESCRIT$(CI%)
               TETABU1$(3, CAITAB1%) = CANS$
               TETABU1$(4, CAITAB1%) = KILO$
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
'
Sub GENKARDEX(CODL$)
   '
   Static DEPXX%, PRIDEP%
   If PRIDEP% < 1 Then
     Call CARDEPOS
     PRIDEP% = COSU%(1)
   End If
   '
   LU$ = LUDAT$
   If EXISTE%(LU$ + "COKARDEX.RFS") < 1 Then
     Call COPYSTRU("MOVISTO", "COKARDEX")
   End If
   REBLA$ = REGBLAN$("COKARDEX")
   '
5  CIXX% = CODINT%(CODL$)
   HOII% = HOY(HOS$)
   VDEF$ = MKI$(CIXX%) + Chr$(DEPXX%) + Chr$(0) + MKI$(1) + MKI$(HOII%) + "  "
   OBX$ = OBJPLA$("RANKARDEX", VDEF$)
   If OBX$ = "" Then Exit Sub
   '
   CIXX% = CVI(Left$(OBX$, 2))
   If CIXX% < 1 Or CIXX% > NUMAS% Then GoTo 5
   CODD$ = CODEXT$(CIXX%)
   '
   DEPXX% = Asc(Mid$(OBX$, 3, 1))
   If DEPXX% < 0 Then DEPXX% = 0
   '
   DESDE = CDbl(CVDAT(CVI(Mid$(OBX$, 5, 2))))
   HASTA = CDbl(CVDAT(CVI(Mid$(OBX$, 7, 2))))
   If DESDE > HASTA Then GoTo 5
   '
   INCLUVIE% = 0
   INCLU$ = TRIM$(UCase$(Mid$(OBX$, 9, 1)))
   If INCLU$ <> "" And INCLU$ <> "N" Then INCLUVIE% = 1
   '
6  CIDD% = CIXX%
   'UREMO& = ULTREG&("MOVISTO")
   JJ& = 0
   If CIDD% > 0 Then
     Call APERBASDAT("STOCKS")
     SQLX$ = " SELECT * FROM MOVISTO "
     SQLX$ = SQLX$ + "WHERE Cod_Inte = " & CIDD% & " "
     SQLX$ = SQLX$ + "AND Cdbl(FechMov) >= " & DESDE & " "
     SQLX$ = SQLX$ + "AND INT(Cdbl(FechMov)) <= " & HASTA & " "
     If DEPXX% > 0 Then
       If DEPXX% = PRIDEP% Then
           SQLX$ = SQLX$ + "AND (DepoMovi = " & DEPXX% & " OR DepoMovi = 0 OR ISNULL(DepoMovi) = True OR ISEMPTY(DepoMovi) = True) "
         Else
           SQLX$ = SQLX$ + "AND DepoMovi = " & DEPXX% & " "
       End If
     End If
     If INCLUVIE% = 1 Then
       SQLX$ = SQLX$ + "AND CodiMovi <> 'KU' "
     End If
     SQLX$ = SQLX$ + "ORDER BY FechMov ASC, FeReMovi ASC; "
     '
     Set MoviTemp = Stocks.OpenRecordset(SQLX$, dbOpenSnapshot)
     With MoviTemp
       On Error Resume Next
       .MoveFirst
       If Err < 1 Then
         Do While Not .EOF
           YY$ = REBLA$
           Call REPLA(YY$, MKI$(CVINT(!FECHMOV)), 1, 2)
           Call REPLA(YY$, MKI$(!COD_INTE), 3, 2)
           Call REPLA(YY$, !IdenLote, 5, 16)
           Call REPLA(YY$, !CodiMovi, 21, 2)
           Call REPLA(YY$, !DoPriCor, 23, 10)
           Call REPLA(YY$, !DoSecCor, 33, 10)
           Call REPLA(YY$, !ReferCor, 43, 30)
           Call REPLA(YY$, MKD$(!VALOUNIT), 73, 8)
             MONEX$ = "$": If !MoneVal = 2 Then MONEX$ = "U$"
           Call REPLA(YY$, MONEX$, 81, 2)
           Call REPLA(YY$, Chr$(!DEPOMOVI), 83, 1)
           Call REPLA(YY$, MKI$(!Nume_Clie), 84, 2)
           Call REPLA(YY$, MKI$(!Nume_Prov), 86, 2)
           Call REPLA(YY$, MKD$(!CANTINGRE), 88, 8)
           Call REPLA(YY$, MKD$(!CANTSALID), 96, 8)
           Call REPLA(YY$, MKD$(0), 104, 8)
           Call REPLA(YY$, CVINT(!FeReMovi), 122, 2)
           Call REPLA(YY$, Chr$(!NumUsuar), 126, 1)
           JJ& = JJ& + 1
           Call GRAREG("COKARDEX", YY$, JJ&)
         .MoveNext
         Loop
       End If
     End With
     Set MoviTemp = Nothing
     Movisto.Close
     '
     On Error GoTo 0
     If INCLUVIE% = 1 Then
        Call APERBASDAT("STOCKVIE")
        SQLX$ = " SELECT * FROM MOVISTO "
        SQLX$ = SQLX$ + "WHERE Cod_Inte = " & CIDD% & " "
        SQLX$ = SQLX$ + "AND Cdbl(FechMov) >= " & DESDE & " "
        SQLX$ = SQLX$ + "AND INT(Cdbl(FechMov)) <= " & HASTA & " "
        If DEPXX% > 0 Then
          SQLX$ = SQLX$ + "AND DepoMovi = " & DEPXX% & " "
        End If
        SQLX$ = SQLX$ + "AND CodiMovi <> 'KU' "
        SQLX$ = SQLX$ + "ORDER BY FechMov ASC, FeReMovi ASC; "
        '
        Set MoviTemp = StockVie.OpenRecordset(SQLX$, dbOpenSnapshot)
        With MoviTemp
          On Error Resume Next
          .MoveFirst
          If Err < 1 Then
            Do While Not .EOF
              YY$ = REBLA$
              Call REPLA(YY$, MKI$(CVINT(!FECHMOV)), 1, 2)
              Call REPLA(YY$, MKI$(!COD_INTE), 3, 2)
              Call REPLA(YY$, !IdenLote, 5, 16)
              Call REPLA(YY$, !CodiMovi, 21, 2)
              Call REPLA(YY$, !DoPriCor, 23, 10)
              Call REPLA(YY$, !DoSecCor, 33, 10)
              Call REPLA(YY$, !ReferCor, 43, 30)
              Call REPLA(YY$, MKD$(!VALOUNIT), 73, 8)
                MONEX$ = "$": If !MoneVal = 2 Then MONEX$ = "U$"
              Call REPLA(YY$, MONEX$, 81, 2)
              Call REPLA(YY$, Chr$(!DEPOMOVI), 83, 1)
              Call REPLA(YY$, MKI$(!Nume_Clie), 84, 2)
              Call REPLA(YY$, MKI$(!Nume_Prov), 86, 2)
              Call REPLA(YY$, MKD$(!CANTINGRE), 88, 8)
              Call REPLA(YY$, MKD$(!CANTSALID), 96, 8)
              Call REPLA(YY$, MKD$(0), 104, 8)
              Call REPLA(YY$, CVINT(!FeReMovi), 122, 2)
              Call REPLA(YY$, Chr$(!NumUsuar), 126, 1)
              JJ& = JJ& + 1
              Call GRAREG("COKARDEX", YY$, JJ&)
            .MoveNext
            Loop
          End If
        End With
        '
     End If
     Call SETULTREG("COKARDEX", JJ&)
     Call CIERRARCH("COKARDEX")
     Screen.MousePointer = 1
     Call CONECRUN("*COKARDEX/" + CODD$, "Ficha Kardex Indirecta") ' + "/WAIT")
     Screen.MousePointer = 1
    Else
     Call MENSERR(24, "Codigo: " + TRIM$(CODD$) + "\sin Movimientos Según Condición SQL.")
   End If
   '
   'If TRIM$(CODL$) = "" Then GoTo 5
   Screen.MousePointer = 1
   GoTo 5
   '
End Sub

Sub VALEMAT()
     '\\\OT-06-0244-WAR-27/06/06///
     RTA1% = ALTERNA%("Registrar Vale de Consumo \Anular Vale de Consumo\Ver y Re-Imprimir")
     Select Case RTA1%
       Case 1
         GoTo 500
       Case 2
         Call ANUVALEMAT
         GoTo 9999
       Case 3
         ' \\\OT-8-973-27/10/08
         Call REIMPRIMEVALE
         ' \\\OT-8-973-27/10/08-FIN
         GoTo 9999
       Case Else
        GoTo 9999
     End Select
     '\\\OT-06-0244-WAR-FIN///
500  DEPORIG% = XVALO(Control$("", "DEPOCEN"))
     DEPOSI$ = ECOARCH$("TADEPOSI", Chr$(DEPORIG%))
     '
1100 Screen.MousePointer = 11
     CMO$ = TRIM$(Control$("", "CODCONSU"))
     If CMO$ = "" Then CMO$ = "VC"
     '
     Call APERBASDAT("STOCKS")
     SQLX$ = "SELECT * FROM MOVISTO "
     SQLX$ = SQLX$ + "WHERE CodiMovi = '" & CMO$ & "' "
     SQLX$ = SQLX$ + "ORDER BY VAL(MID$(DoPriCor,4,7)) DESC"
     '
     Set MOVISTEMP = Stocks.OpenRecordset(SQLX$, dbOpenSnapshot)
     UNRO& = 0
     On Error Resume Next
     With MOVISTEMP
       .MoveFirst
       If Err < 1 Then
         UNRO& = XVALO(Mid$(!DoPriCor, 4, 7))
       End If
       On Error GoTo 0
     End With
     '
1002 NROVAL& = UNRO& + 1
     Call FRESHALL
     '
1105 Call CIERRARCH("*.*")
     '
1111 HOII% = HOY(HOS$)
     VDEF$ = MKS$(NROVAL&) + MKI$(HOII%) + Chr$(DEPORIG%) + Chr$(0) + Space$(32)
     PEDY$ = OBJPLA$("VALEMATER", VDEF$)
     If PEDY$ = "" Then GoTo 9999
     FF% = CVI(Mid$(PEDY$, 5, 2))
     DEPORIG% = Asc(Mid$(PEDY$, 7, 1))
     DEPOSI$ = ECOARCH$("TADEPOSI", Chr$(DEPORIG%))
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
     If DEPORIG% < 1 Or DEPORIG% >= 99 Or TRIM$(DEPOSI$) = "" Then
        Call MENSERR(24, "Depósito de Salida no Válido")
        GoTo 1111
     End If
     '
1112    Call SETULTREG("!BOLREDET", 0)
     Call CIERRARCH("*.*")
     '
     Screen.MousePointer = 11
     '
     FFF$ = REPLACAR$(FECHATEX$(FF%), "/", "-")
1113 VTB% = VENTABU%("BOLREDET/NC/TITUPAN=Vale de Materiales - " + TRIM$(DEPOSI$) + " - " + FFF$)
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
       If PUEDEMOVI%(DEPORIG%) < 0 Then GoTo 1113 ' CANCELA SALIDA
       '\\\OT-05-0730-WAR-FIN///
       '
       ' VALIDA LOS DATOS INGRESADOS
       If EXISTE%("TRAZAB04.EXE") > 0 Then
          HUBOSELOT% = 0
          PUNCODII$ = ""
          For YX& = 1 To ULTREG&("!BOLREDET")
             YY$ = REGLEIDO$("!BOLREDET", YX&)
             CI% = CVI(Mid$(YY$, 83, 2))
             For JJI% = 1 To Len(PUNCODII$) Step 2
                If CVI(Mid$(PUNCODII$, JJI%, 2)) = CI% Then
                  Call MENSERR(24, "Codigo '" + TRIM$(CODEXT$(CI%)) + "' Repetido.")
                  GoTo 1113
                End If
             Next JJI%
             PUNCODII$ = PUNCODII$ + MKI$(CI%)
             ' OBLIGA A INGRESAR EL LOTE
             ' COMPARA CANTIDAD DESPACHADA CON SALDO DEL LOTE
             CANTSA = CVD(Mid$(YY$, 89, 8))
             DEPX% = DEPORIG%
             LOTEX$ = TRIM$(Mid$(YY$, 107, 12))
             If LOTEX$ <> "" Then
                If SALOTE(CI%, LOTEX$, 1) >= CANTSA Then
                   GoTo 19
                End If
             End If
             LOTEX$ = LOTESELEC$(CI%, CANTSA)
             HUBOSELOT% = 1
             If LOTEX$ = "" Then
                GoTo 1113
             End If
19           Call REPLA(YY$, LOTEX$, 69, 12)
             Call REPLA(YY$, LOTEX$, 107, 12)
             Call GRAREG("!BOLREDET", YY$, YX&)
          Next YX&
          Call CIERRARCH("!BOLREDET")
          If HUBOSELOT% = 1 Then GoTo 1113
       End If
       '
       Call SETULTREG("!OBSERVOF", 0)
       Call CIERRARCH("!OBSERVOF")
       VTB% = VENTABU%("OBSERVOF/TITUPAN=Vale de Materiales - " + TRIM$(DEPOSI$) + " - " + FFF$)
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
               Call MOVISTOK(FF%, CI%, LOTE$, CMO$, REMI$, REMI$, DOSEC$, DOSEC$, REFE$, NUORIT%, VUNI#, 0, 0, DEPORIG%, CANTI) '\\\OT-06-0244-WAR-20/06/06///
               Call ACSALOTE(CI%, LOTE$, "NOMESS")
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
         FEX = FF%
         FECHDOC$ = FECHATEX$(FEX)
         NUMEDOC$ = REMI$
         TIPODOC$ = "Vale de Consumo de Materiales"
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
     RTA1% = ALTERNA%("Registrar Vale de Scrap\Ver y Re-Imprimir")
     Select Case RTA1%
       Case 1
         GoTo 1100
       Case 2
         ' \\\OT-8-973-27/10/08
         Call REIMPRIMESCRAP
         GoTo 9999
         ' \\\OT-8-973-27/10/08-FIN
       Case Else
        GoTo 9999
     End Select
     '
     EPAC$ = TRIM$(UCase$(EMPREAC$))
     DEPORIG% = XVALO(Control$("", "DEPOCEN"))
     If InStr(EPAC$, "DOSIVA") > 0 Then
       DEPORIG% = 91
     End If
     DEPOSI$ = ECOARCH$("TADEPOSI", Chr$(DEPORIG%))
     '
1100 Screen.MousePointer = 11
     CMO$ = TRIM$(Control$("", "CODSCRAP"))
     If CMO$ = "" Then CMO$ = "SC"
     '
     Call APERBASDAT("STOCKS")
     SQLX$ = "SELECT * FROM MOVISTO "
     SQLX$ = SQLX$ + "WHERE CodiMovi = '" & CMO$ & "' "
     SQLX$ = SQLX$ + "ORDER BY VAL(MID$(DoPriCor,4,7)) DESC"
     '
     Set MOVISTEMP = Stocks.OpenRecordset(SQLX$, dbOpenSnapshot)
     UNRO& = 0
     On Error Resume Next
     With MOVISTEMP
       .MoveFirst
       If Err < 1 Then
         UNRO& = XVALO(Mid$(!DoPriCor, 4, 7))
       End If
       On Error GoTo 0
     End With
     '
1002 NROVAL& = UNRO& + 1
     Call FRESHALL
     '
1105 Call CIERRARCH("*.*")
     '
1111 HOII% = HOY(HOS$)
     VDEF$ = MKS$(NROVAL&) + MKI$(HOII%) + Chr$(DEPORIG%) + Chr$(0) + Space$(32) + USERNAME$
     PEDY$ = OBJPLA$("DESCARSCRAP", VDEF$)
     If PEDY$ = "" Then GoTo 9999
     FF% = CVI(Mid$(PEDY$, 5, 2))
     DEPORIG% = Asc(Mid$(PEDY$, 7, 1))
     DEPOSI$ = ECOARCH$("TADEPOSI", Chr$(DEPORIG%))
     REFE$ = Mid$(PEDY$, 9, 32)
     If TRIM$(REFE$) = "" Then
       Call MENSERR(24, "Referencia o Motivo es Mandatoria")
       GoTo 1111
     End If
     If DEPORIG% < 1 Or DEPORIG% >= 99 Or TRIM$(DEPOSI$) = "" Then
       Call MENSERR(24, "Depósito de Salida no Válido")
       GoTo 1111
     End If
     '
1112 Call SETULTREG("!BOLREDET", 0)
     Call CIERRARCH("*.*")
     '
     Screen.MousePointer = 11
     '
     FFF$ = REPLACAR$(FECHATEX$(FF%), "/", "-")
'
1113 VTB% = VENTABU%("BOLREDET/NC/TITUPAN=Descarte de Scrap - " + TRIM$(DEPOSI$) + " - " + FFF$)
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
       If PUEDEMOVI%(DEPORIG%) < 0 Then GoTo 1113 ' CANCELA SALIDA
       '\\\OT-05-0730-WAR-FIN///
       '
       ' VALIDA LOS DATOS INGRESADOS
       If EXISTE%("TRAZAB04.EXE") > 0 Then
          HUBOSELOT% = 0: PUNCODII$ = ""
          For YX& = 1 To ULTREG&("!BOLREDET")
             YY$ = REGLEIDO$("!BOLREDET", YX&)
             CI% = CVI(Mid$(YY$, 83, 2))
             For JJI% = 1 To Len(PUNCODII$) Step 2
                If CVI(Mid$(PUNCODII$, JJI%, 2)) = CI% Then
                  Call MENSERR(24, "Codigo '" + TRIM$(CODEXT$(CI%)) + "' Repetido.")
                  GoTo 1113
                End If
             Next JJI%
             PUNCODII$ = PUNCODII$ + MKI$(CI%)
             ' OBLIGA A INGRESAR EL LOTE
             ' COMPARA CANTIDAD DESPACHADA CON SALDO DEL LOTE
             CANTSA = CVD(Mid$(YY$, 89, 8))
             DEPX% = DEPORIG%
             LOTEX$ = TRIM$(Mid$(YY$, 107, 12))
             If LOTEX$ <> "" Then
                If SALOTE(CI%, LOTEX$, 1) >= CANTSA Then
                   GoTo 19
                End If
             End If
             LOTEX$ = LOTESELEC$(CI%, CANTSA)
             HUBOSELOT% = 1
             If LOTEX$ = "" Then
                GoTo 1113
             End If
19           Call REPLA(YY$, LOTEX$, 69, 12)
             Call REPLA(YY$, LOTEX$, 107, 12)
             Call GRAREG("!BOLREDET", YY$, YX&)
          Next YX&
          Call CIERRARCH("!BOLREDET")
          If HUBOSELOT% = 1 Then GoTo 1113
       End If
       '
       Call SETULTREG("!OBSERVOF", 0)
       Call CIERRARCH("!OBSERVOF")
       VTB% = VENTABU%("OBSERVOF/TITUPAN=Descarte de Scrap - " + TRIM$(DEPOSI$) + " - " + FFF$)
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
               DOSEC$ = ""
1120           CANTI = (-1) * CAN#
               NUORIT% = NUORIT% + 1
               Call MOVISTOK(FF%, CI%, LOTE$, CMO$, REMI$, REMI$, DOSEC$, DOSEC$, REFE$, NUORIT%, 0, 0, 0, DEPORIG%, CANTI)
               Call ACSALOTE(CI%, LOTE$, "NOMESS")
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

Sub VESARRGLO()
   '
   SCMA = Screen.MousePointer
   Screen.MousePointer = 11
   '
   Call CARDEPOS
   '
   ' EL 19 DE NOVIEMBRE DE 2007 SE MODIFICÓ PARA LIMITAR LA CONSULTA SQL A 32
   ' ARTICULOS DIFERENTES. PASADO ESE NUMERO SE PONE EL REPITECON% EN 1 Y PRO-
   ' SESA EL BLOQUE SELECCIONADO. LUEGO REPITE PARA EL PROXIMO BLOQUE DE AR-
   ' TICULOS.
   '
10 Call APERBASDAT("STOCKS")
   ATX% = 0: CACODI& = 0
   REPITECON% = 0
   For KU& = 1 To NUMAS%
     If CIMOVI%(KU&) > 0 Then
       CIMOVI%(KU&) = 2
       CACODI& = CACODI& + 1
       If ATX% < 1 Then
           SQLX$ = " SELECT * FROM MOVISTO "
           SQLX$ = SQLX$ + "WHERE Cod_Inte = " & KU& & " "
           ATX% = 1
         Else
           SQLX$ = SQLX$ + "OR Cod_Inte = " & KU& & " "
       End If
     End If
     If CACODI& >= 32 Then
        REPITECON% = 1
        GoTo 20
     End If
   Next KU&
   '
20 If CACODI& < 1 Then GoTo 99
   '
   Set MoviTemp = Stocks.OpenRecordset(SQLX$, dbOpenSnapshot)
   KKX& = 0
   For KU& = 1 To NUMAS%
     If CIMOVI%(KU&) = 2 Then
       KKX& = KKX& + 1
       Call TRACE(20, KKX&, CACODI&)
       '
       For KI% = 0 To NUSU%
         CATEO#(KI%) = 0
       Next KI%
       '
       CAMOVI& = 0
       With MoviTemp
         On Error Resume Next
         .MoveFirst
         If Err < 1 Then
           Do While Not .EOF
             If !COD_INTE = KU& Then
               CAMOVI& = CAMOVI& + 1
               CANTI# = !CANTINGRE - !CANTSALID
               DEP% = !DEPOMOVI
               '
               COLUX% = 1
               For KI% = 1 To NUSU%
                 If DEP% = COSU%(KI%) Then
                   COLUX% = KI%
                   Exit For
                 End If
               Next KI%
               '
               CATEO#(COLUX%) = CATEO#(COLUX%) + CANTI#
             End If
           .MoveNext
           Loop
         End If
       End With
       '
       VESAL$ = String$(256, 0)
       For KI% = 1 To NUSU%
         Mid$(VESAL$, 4 * KI% - 3, 4) = MKS$(CATEO#(KI%))
       Next KI%
       '
       With SalStock
         .FindFirst "CodiInte = " & KU&
         If .NoMatch = False Then
             .Edit
           Else
             .AddNew
             !CodiInte = CIXI%
         End If
         !VeSalSto = VESAL$
         .Update
       End With
       '
       CIMOVI%(KU&) = 0
       '
     End If
   Next KU&
   '
   If REPITECON% > 0 Then GoTo 10
   '
99 End Sub

Sub VESARRGLO2()
   '
   SCMA = Screen.MousePointer
   Screen.MousePointer = 11
   '
   Call CARDEPOS
   '
   ' EL 19 DE NOVIEMBRE DE 2007 SE MODIFICÓ PARA LIMITAR LA CONSULTA SQL A 32
   ' ARTICULOS DIFERENTES. PASADO ESE NUMERO SE PONE EL REPITECON% EN 1 Y PRO-
   ' SESA EL BLOQUE SELECCIONADO. LUEGO REPITE PARA EL PROXIMO BLOQUE DE AR-
   ' TICULOS.
   '
10 Call APERBASDAT("STOCKS")
   ATX% = 0: CACODI& = 0
   REPITECON% = 0
   For KU& = 1 To NUMAS%
     If CIMOVI%(KU&) > 0 Then
       CIMOVI%(KU&) = 2
       CACODI& = CACODI& + 1
       If ATX% < 1 Then
           SQLX$ = " SELECT COD_INTE, DEPOMOVI, SUM(CANTINGRE) AS INGRE, SUM(CANTSALID) AS SALID FROM MOVISTO "
           SQLX$ = SQLX$ + "WHERE Cod_Inte = " & KU& & " "
           ATX% = 1
         Else
           SQLX$ = SQLX$ + "OR Cod_Inte = " & KU& & " "
       End If
     End If
'     If CACODI& >= 32 Then
'        REPITECON% = 1
'        GoTo 20
'     End If
   Next KU&
   '
20 If CACODI& < 1 Then GoTo 99
   SQLX$ = SQLX$ + " GROUP BY COD_INTE, DEPOMOVI "
   '
   Set MoviTemp = Stocks.OpenRecordset(SQLX$, dbOpenSnapshot)
   KKX& = 0
   For KU& = 1 To NUMAS%
     If CIMOVI%(KU&) = 2 Then
       KKX& = KKX& + 1
       Call TRACE(20, KKX&, CACODI&)
       '
       For KI% = 0 To NUSU%
         CATEO#(KI%) = 0
       Next KI%
       '
       CAMOVI& = 0
       With MoviTemp
         On Error Resume Next
         .MoveFirst
         If Err < 1 Then
           Do While Not .EOF
             If !COD_INTE = KU& Then
               CAMOVI& = CAMOVI& + 1
               CANTI# = !INGRE - !SALID
               DEP% = !DEPOMOVI
               '
               COLUX% = 1
               For KI% = 1 To NUSU%
                 If DEP% = COSU%(KI%) Then
                   COLUX% = KI%
                   Exit For
                 End If
               Next KI%
               '
               CATEO#(COLUX%) = CATEO#(COLUX%) + CANTI#
             End If
           .MoveNext
           Loop
         End If
       End With
       '
       VESAL$ = String$(256, 0)
       For KI% = 1 To NUSU%
         Mid$(VESAL$, 4 * KI% - 3, 4) = MKS$(CATEO#(KI%))
       Next KI%
       '
       With SalStock
         .FindFirst "CodiInte = " & KU&
         If .NoMatch = False Then
             .Edit
           Else
             .AddNew
             !CodiInte = CIXI%
         End If
         !VeSalSto = VESAL$
         .Update
       End With
       '
       CIMOVI%(KU&) = 0
       '
     End If
   Next KU&
   '
   If REPITECON% > 0 Then GoTo 10
   '
99 End Sub
Sub VESARRGLO1()
   '
   SCMA = Screen.MousePointer
   Screen.MousePointer = 11
   '
   Call CARDEPOS
   '
10 Call APERBASDAT("STOCKS")
   ATX% = 0: CACODI& = 0
   REPITECON% = 0
   For KU& = 1 To NUMAS%
     If CIMOVI%(KU&) > 0 Then
       CIMOVI%(KU&) = 2
       CACODI& = CACODI& + 1
     End If
   Next KU&
   '
20 If CACODI& < 1 Then GoTo 99
   '
   KKX& = 0
   For KU& = 1 To NUMAS%
     If CIMOVI%(KU&) = 2 Then
       KKX& = KKX& + 1
       Call TRACE(20, KKX&, CACODI&)
       '
       For KI% = 0 To NUSU%
         CATEO#(KI%) = 0
       Next KI%
       '
       CAMOVI& = 0
       SQLX$ = "SELECT DEPOMOVI, SUM(CANTINGRE) AS INGRE, SUM(CANTSALID) AS SALID FROM MOVISTO WHERE COD_INTE=" & KU& & " GROUP BY DEPOMOVI"
       Set MoviTemp = Stocks.OpenRecordset(SQLX$, dbOpenSnapshot)
       With MoviTemp
         On Error Resume Next
         .MoveFirst
         If Err < 1 Then
           Do While Not .EOF
             If !COD_INTE = KU& Then
               CAMOVI& = CAMOVI& + 1
               CANTI# = !INGRE - !SALID
               DEP% = !DEPOMOVI
               '
               COLUX% = 1
               For KI% = 1 To NUSU%
                 If DEP% = COSU%(KI%) Then
                   COLUX% = KI%
                   Exit For
                 End If
               Next KI%
               '
               CATEO#(COLUX%) = CATEO#(COLUX%) + CANTI#
             End If
             .MoveNext
           Loop
         End If
       End With
       '
       VESAL$ = String$(256, 0)
       For KI% = 1 To NUSU%
         Mid$(VESAL$, 4 * KI% - 3, 4) = MKS$(CATEO#(KI%))
       Next KI%
       '
       With SalStock
         .FindFirst "CodiInte = " & KU&
         If .NoMatch = False Then
             .Edit
           Else
             .AddNew
             !CodiInte = CIXI%
         End If
         !VeSalSto = VESAL$
         .Update
       End With
       '
       CIMOVI%(KU&) = 0
       '
     End If
   Next KU&
   '
   If REPITECON% > 0 Then GoTo 10
   '
99 End Sub

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
  SQLX$ = ""
  SQLX$ = "SELECT  DoPRICor, ReferCor, FechMov FROM MOVISTO "
  SQLX$ = SQLX$ + "WHERE CodiMovi = '" & CMO$ & "' "
  SQLX$ = SQLX$ + " Order by VAL(MID$(DoPriCor,4,7)) ASC"
  Set VAMTMP = Stocks.OpenRecordset(SQLX$, 4)
  With VAMTMP
    On Error Resume Next
    .MoveFirst
    If Err Then
      Call MENSERR(24, "No Existen Vales de Materiales\para su Anulación")
      On Error GoTo 0
      GoTo 99
    End If
    On Error GoTo 0
    JJ& = 0
    REBLA$ = REGBLAN$("INVALEMA")
    Do While .EOF = False
      YY$ = REBLA$
      NUVAMA& = XVALO(Mid$(!DoPriCor, 4))
      If NUMEA1& <> NUVAMA& Then
        Call REPLA(YY$, MKS$(NUVAMA&), 1, 4)
        REFER1$ = ""
        F1% = CVINT(!FECHMOV)
        F2$ = FECHATEX$(F1%)
        REFER1$ = F2$ + " - " + !ReferCor
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
      With Movisto
        DOCVAMA1$ = CMO$ + "." + TRIM$(Str$(NUVAMA1&))
        SQLX1$ = "DoPriCor = '" & DOCVAMA1$ & "' "
        .FindFirst SQLX1$
7       If .NoMatch = False Then
          .Delete
          .FindNext SQLX1$
          GoTo 7
        End If
      End With
      Call COMUNI("Anulación Completa")
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
     OcomEnca.FindFirst "Nume_Ocom = " & NUOCOMX&
     If OcomEnca.NoMatch = False Then
       STAOC% = OcomEnca!Stat_Ocom
     End If
   End If
   '
   STATOCOM% = STAOC%
End Function

Sub CREALOTIN(FALTA%, CIXI%, IDLOT$, CANOM#)
   '
   IDLOTX$ = TRIM$(IDLOT$)
   '
   Call APERBASDAT("STOCKS")
   SQLX$ = " SELECT * FROM LOTINGRE "
   SQLX$ = SQLX$ + "WHERE Cod_Inte = " & CIXI% & " "
   SQLX$ = SQLX$ + "AND IdenLote = '" & IDLOTX$ & "' "
   Set LotIngTemp = Stocks.OpenRecordset(SQLX$, dbOpenDynaset)
   '
   With LotIngTemp
     NUEVO% = 0
25   On Error Resume Next
     .MoveFirst
     If Err Then
         On Error GoTo 0
         .AddNew
         NUEVO% = 1
       Else
         On Error GoTo 0
         .Edit
     End If
     !COD_INTE = CIXI%
     !IdenLote = IDLOTX$
     If NUEVO% = 1 Then !FechAlta = CVDAT(FALTA%)   ' PARA QUE NO SOBRE-ESCRIBA LA ORIGINAL
     !CanAlta = CANOM#
     If !CanSaldo < 0.05 Or !CanSaldo > CANOM# Then
        !CanSaldo = CANOM#
     End If
     .Update
   End With
   '
   Call ACSALOTE(CIXI%, IDLOT$, "NOMESS")
   '
End Sub
'
Sub ACSALOTE(CIXI%, IDLO$, Optional MODOX$)    ' ACTUALIZA SALDOS POR LOTE
    '
    Call ACSALOT(CIXI%, IDLO$, CANTIN#, CACONRE#, CARESRE#, CSALDO#, NPROVX%, MODOX$)
    Exit Sub
' MODIFICADO PARA UNIFICACION DE CRITERIOS

'    IDLOTX$ = TRIM$(IDLO$)
'    If IDLOTX$ = "" Then
'       Exit Sub
'    End If
'    '
'    Call APERBASDAT("STOCKS")
'    SQLZ$ = " SELECT * FROM MOVISTO "
'    SQLZ$ = SQLZ$ + "WHERE Cod_Inte = " & CIXI% & " "
'    SQLZ$ = SQLZ$ + "ORDER BY FechMov ASC"
'    Set MOVCONTEMP = Stocks.OpenRecordset(SQLZ$, dbOpenSnapshot)
'    '
'    SUINGRE# = 0: SUCONSU# = 0
'    With MOVCONTEMP
'      On Error Resume Next
'      .MoveFirst
'      If Err < 1 Then
'         On Error GoTo 0
'         Do While Not .EOF
'           IDLOTI$ = ""
'           On Error Resume Next
'           IDLOTI$ = !IdenLote
'           On Error GoTo 0
'           If TRIM$(IDLOTI$) = IDLOTX$ Then
'              CENTRA# = !CantIngre
'              CSALI# = !Cantsalid
'              If !CodiMovi = "TR" Then
'                 If Left$(IDLOTX$, 2) <> "TR" And Left$(IDLOTX$, 2) <> "AS" Then
'                   CSALI# = CSALI# - CENTRA#
'                   CENTRA# = 0
'                 End If
'              End If
'              SUINGRE# = SUINGRE# + CENTRA#
'              SUCONSU# = SUCONSU# + CSALI#
'           End If
'         .MoveNext
'         Loop
'      End If
'    End With
'    '
'    Call APERBASDAT("RESERVA")
'    SQLZ$ = " SELECT * FROM RESERVA "
'    SQLZ$ = SQLZ$ + "WHERE Cod_Inte = " & CIXI% & " "
'    SQLZ$ = SQLZ$ + "AND Status = 0 "     ' para no considerar los marcados para anular
'    SQLZ$ = SQLZ$ + "ORDER BY FechRes ASC"
'    Set ReserTemp = Stocks.OpenRecordset(SQLZ$, dbOpenSnapshot)
'    '
'    SURESER# = 0
'    With ReserTemp
'      On Error Resume Next
'      .MoveFirst
'      If Err < 1 Then
'         On Error GoTo 0
'         Do While Not .EOF
'           CSALI# = !CantRes - !CantCons
'           If CSALI# >= 0.05 Then
'             LCONSU$ = ""
'             On Error Resume Next
'             LCONSU$ = !IdenLote
'             On Error GoTo 0
'             If LCONSU$ = IDLOTX$ Then
'                SURESER# = SURESER# + CSALI#
'             End If
'           End If
'         .MoveNext
'         Loop
'      End If
'    End With
'    '
'20  SQLX$ = "SELECT * FROM LOTINGRE "
'    SQLX$ = SQLX$ + "WHERE IdenLote = '" & IDLOTX$ & "'"
'    SQLX$ = SQLX$ + "AND Cod_Inte = " & CIXI%
'    Set LOTINSEL = Stocks.OpenRecordset(SQLX$, dbOpenDynaset) ' para poder ajustar cantidad saldo
'    With LOTINSEL
'       On Error Resume Next
'       .MoveFirst
'       If Err Then
'          On Error GoTo 0
'          Call MENSERR(24, "Lote '" + IDLOTX$ + "' no Encontrado.\Probable Falta Trazar Ingreso.")
'          Exit Sub
'       End If
'       On Error GoTo 0
'       '
'       CSALDO# = SUINGRE# - SUCONSU# - SURESER#
'       If CSALDO# < (-0.05) Then
'          Call MENSERR(24, "Exceso de Consumo en Código '" + CODEXT$(CIXI%) + "' - Lote '" + IDLOTX$ + "'\Saldo Ajustado a Cero.")
'          CSALDO# = 0
'          GoTo 25
'       End If
'0
'       If IsNull(!CanConsu) = True Or IsEmpty(!CanConsu) = True Then GoTo 25
'       If IsNull(!CanReser) = True Or IsEmpty(!CanReser) = True Then GoTo 25
'       If Abs(SUINGRE# - !CanAlta) >= 0.05 Or Abs(CSALDO# - !CanSaldo) >= 0.05 Or Abs(SUCONSU# - !CanConsu) >= 0.05 Or Abs(SURESER# - !CanReser) >= 0.05 Then
'          If UCase$(MODOX$) <> "NOMESS" Then
'             Call MENSERR(24, "Saldos en Código '" + CODEXT$(CIXI%) + "' - Lote '" + IDLOTX$ + "'\Ajustado Según Ingresos, Reservas y Consumos.")
'          End If
'25        .Edit
'          !CanAlta = SUINGRE#
'          !CanConsu = SUCONSU#
'          !CanReser = SURESER#
'          !CanSaldo = CSALDO#
'          .Update
'       End If
'    End With
'    '
End Sub
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
    SQLX$ = "SELECT * FROM LOTINGRE "
    SQLX$ = SQLX$ + "WHERE IdenLote = '" & IDLOTX$ & "'"
    SQLX$ = SQLX$ + "AND Cod_Inte = " & CIXI%
    Set LOTINSEL = Stocks.OpenRecordset(SQLX$, dbOpenDynaset) ' para poder ajustar cantidad saldo
    With LOTINSEL
       On Error Resume Next
       .MoveFirst
       If Err Then
          On Error GoTo 0
          Call MENSERR(24, "Lote '" + IDLOTX$ + "' no Encontrado.\Probable Falta Trazar Ingreso.")
          Exit Sub
       End If
       On Error GoTo 0
    End With
    '
    Call APERBASDAT("STOCKS")
    SQLZ$ = " SELECT * FROM MOVISTO "
    SQLZ$ = SQLZ$ + "WHERE Cod_Inte = " & CIXI% & " "
    SQLZ$ = SQLZ$ + "AND IdenLote = '" & IDLOTX$ & "' "
    SQLZ$ = SQLZ$ + "ORDER BY FechMov ASC"
    Set MOVCONTEMP = Stocks.OpenRecordset(SQLZ$, dbOpenSnapshot)
    '
    SUINGRE# = 0: SUCONSU# = 0
    With MOVCONTEMP
      On Error Resume Next
      .MoveFirst
      If Err < 1 Then
         On Error GoTo 0
         Do While Not .EOF
            CINGRE# = !CANTINGRE
            CCONSU# = !CANTSALID
            If !CodiMovi = "TR" Then
               If !DoPriCor = !IdenLote And !CANTINGRE > 0 Then GoTo 29
               CCONSU# = CCONSU# - CINGRE#
               CINGRE# = 0
            End If
29          SUINGRE# = SUINGRE# + CINGRE#
            SUCONSU# = SUCONSU# + CCONSU#
            '
            ' aqui prever varios ingresos
            If Abs(CINGRE#) >= 0.05 Then
               For KKI% = 1 To CAMOVILO%
                  If TIMOVILO%(KKI%) = 1 Then
                     CANINGLO#(CAMOVILO%) = CANINGLO#(CAMOVILO%) + CINGRE#
                     CANSALLO#(CAMOVILO%) = CANSALLO#(CAMOVILO%) + CCONSU#
                     GoTo 32
                  End If
               Next KKI%
            End If
            '
            CAMOVILO% = CAMOVILO% + 1
            TIMOVILO%(CAMOVILO%) = 1: If Abs(CINGRE#) < 0.05 Then TIMOVILO%(CAMOVILO%) = 2
            CANINGLO#(CAMOVILO%) = CINGRE#
            CANSALLO#(CAMOVILO%) = CCONSU#
            CANRESLO#(CAMOVILO%) = 0
            FEMOVILO%(CAMOVILO%) = CVINT(!FECHMOV)
            DEMOVILO$(CAMOVILO%) = !DOPRILAR
               '///OT-8-836-12/09/08
               If Left$(SAFETEXT(!ReferCor), 7) = "RECHAZO" Then DEMOVILO$(CAMOVILO%) = "RECHAZO"
               If Left$(SAFETEXT(!ReferCor), 8) = "FALTANTE" Then DEMOVILO$(CAMOVILO%) = "FALTANTE"
               '///OT-8-836-FIN
               ' aca faltan los numeros PNL/PLC
            '
32          If !CANTINGRE >= 0.05 Then
               If NPROVEI% < 1 Then
                  NPROVEI% = !Nume_Prov
               End If
            End If
            '
         .MoveNext
         Loop
      End If
    End With
    '
    Call APERBASDAT("RESERVA")
    SQLZ$ = " SELECT * FROM RESERVA "
    SQLZ$ = SQLZ$ + "WHERE Cod_Inte = " & CIXI% & " "
    SQLZ$ = SQLZ$ + "AND IdenLote = '" & IDLOTX$ & "' "
    SQLZ$ = SQLZ$ + "AND Status = 0 "     ' para no considerar los marcados para anular
    SQLZ$ = SQLZ$ + "ORDER BY FechRes ASC"
    Set ReserTemp = Stocks.OpenRecordset(SQLZ$, dbOpenSnapshot)
    '
    SURESER# = 0
    With ReserTemp
      On Error Resume Next
      .MoveFirst
      If Err < 1 Then
         On Error GoTo 0
         Do While Not .EOF
           CRESER# = !CantRes - !CantCons
           If CRESER# >= 0.05 Then
              CAMOVILO% = CAMOVILO% + 1
              TIMOVILO%(CAMOVILO%) = 3
              CANINGLO#(CAMOVILO%) = 0
              CANSALLO#(CAMOVILO%) = 0
              CANRESLO#(CAMOVILO%) = CRESER#
              FEMOVILO%(CAMOVILO%) = CVINT(!FechRes)
              DEMOVILO$(CAMOVILO%) = !IdSubOr
              SURESER# = SURESER# + CRESER#
           End If
         .MoveNext
         Loop
      End If
    End With
    '
20  With LOTINSEL
       CSALDO# = SUINGRE# - SUCONSU# - SURESER#
       If CSALDO# < (-0.05) Then
          Call MENSERR(24, "Exceso de Consumo en Código '" + CODEXT$(CIXI%) + "' - Lote '" + IDLOTX$ + "'\Saldo Ajustado a Cero.")
          CSALDO# = 0
          GoTo 25
       End If
       '
       If IsNull(!CanConsu) = True Or IsEmpty(!CanConsu) = True Then GoTo 25
       If IsNull(!CanReser) = True Or IsEmpty(!CanReser) = True Then GoTo 25
       If Abs(SUINGRE# - !CanAlta) >= 0.05 Or Abs(CSALDO# - !CanSaldo) >= 0.05 Or Abs(SUCONSU# - !CanConsu) >= 0.05 Or Abs(SURESER# - !CanReser) >= 0.05 Then
          If UCase$(MODOX$) <> "NOMESS" Then
             Call MENSERR(24, "Saldos en Código '" + CODEXT$(CIXI%) + "' - Lote '" + IDLOTX$ + "'\Ajustado Según Ingresos, Reservas y Consumos.")
          End If
25        .Edit
          !CanAlta = SUINGRE#
          !CanConsu = SUCONSU#
          !CanReser = SURESER#
          !CanSaldo = CSALDO#
          .Update
       End If
    End With
    '
End Sub
'
Function SALPORLOTES(CIXI%)
    '
    Call APERBASDAT("STOCKS")
    SQLX$ = " SELECT * FROM LOTINGRE WITH(NOLOCK)"
    SQLX$ = SQLX$ + "WHERE Cod_Inte = " & CIXI% & " "
    Set LotIngTemp = Stocks.OpenRecordset(SQLX$, dbOpenSnapshot)
    '
    SALPLO = 0
    With LotIngTemp
         Do While Not .EOF
             SALFIS = !CanAlta - !CanConsu: If SALFIS < 0 Then SALFIS = 0
             SALPLO = SALPLO + SALFIS
         .MoveNext
         Loop
    End With
    '
    SALPORLOTES = SALPLO
    '
End Function
'
Function SALOTE(CIXI%, IDLO$, Optional MODO%)
    ' DEVUELVE EL SALDO DEL LOTE
    ' SI MODO%=1 ENTONCES EL NETO (DESCONTADO EL RESERVADO)
    ' SI MODO%=2 ENTONCES EL BRUTO (DESCONTANDO SOLO EL CONSUMIDO)
    '
    Call ACSALOTE(CIXI%, IDLO$, "NOMESS")
    '
20  SALOTE = 0
    SQLX$ = "SELECT * FROM LOTINGRE "
    SQLX$ = SQLX$ + "WHERE IdenLote = '" & IDLO$ & "'"
    Set LOTINSEL = Stocks.OpenRecordset(SQLX$, dbOpenSnapshot)
    With LOTINSEL
       On Error Resume Next
       .MoveFirst
       If Err < 1 Then
          SALOTE = !CanAlta - !CanConsu
          If MODO% = 1 Then
             SALOTE = SALOTE - !CanReser
          End If
       End If
    End With
    If SALOTE < 0 Then SALOTE = 0
    '
End Function

Sub CREALOTES(CIXI%)
   '
   Screen.MousePointer = 11
   Dim TLOTE$(), CALOTE#()
   '
   'GENERAR LOS MOVIMIENTOS DE ENTRADA DE LOTES
   Call APERBASDAT("STOCKS")
   SQLX$ = " SELECT * FROM MOVISTO "
   SQLX$ = SQLX$ + "WHERE Cod_Inte = " & CIXI% & " "
   SQLX$ = SQLX$ + "AND CantIngre > 0 "
   SQLX$ = SQLX$ + "ORDER BY FechMov ASC "
   '
   Set MoviTemp = Stocks.OpenRecordset(SQLX$, dbOpenDynaset)
   CAMOVI& = 0
   With MoviTemp
     On Error Resume Next
     .MoveFirst
     If Err < 1 Then
       Do While Not .EOF
         ENTRA# = !CANTINGRE
         If ENTRA# >= 0.05 Then
           IDENLO$ = TRIM$(!IdenLote): RECAR% = 0
           If IDENLO$ = "" Then IDENLO$ = TRIM$(!DoPriCor): RECAR% = 1
           If IDENLO$ = "" Then IDENLO$ = TRIM$(!DoSecCor): RECAR% = 1
           IDENLO$ = AJUSTI$(IDENLO$, 16)
           If RECAR% = 1 Then
             .Edit
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
           FEX = CVINT(!FECHMOV)
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
   '
   ' VERIFICAR QUE COINCIDEN CON LOS DE TABLA LOTINGRE
   SQLX$ = " SELECT * FROM LOTINGRE "
   SQLX$ = SQLX$ + "WHERE Cod_Inte = " & CIXI% & " "
   SQLX$ = SQLX$ + "ORDER BY FechAlta ASC "
   Set LotIngTemp = Stocks.OpenRecordset(SQLX$, dbOpenDynaset)
   '
   With LotIngTemp
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
           If FEALTA <> !FechAlta Or Abs(CANORIG - !CanAlta) > 0.05 Then
             .Delete
             GoTo 25
           End If
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
         .AddNew
         !COD_INTE = CIXI%
         !IdenLote = TRIM$(IDLOTAB$)
         !FechAlta = FEALTA
         !CanAlta = CANORIG
         !CanSaldo = CANORIG
         .Update
       End If
     Next KKI%
   End With
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
   SQLX$ = " SELECT * FROM LOTINGRE "
   SQLX$ = SQLX$ + "WHERE Cod_Inte = " & CIXI% & " "
   SQLX$ = SQLX$ + "ORDER BY FechAlta ASC "
   Set LotIngTemp = Stocks.OpenRecordset(SQLX$, dbOpenDynaset)
   '
   CANLOTES% = 0
   With LotIngTemp
     On Error Resume Next
     .MoveFirst
     If Err < 1 Then
       On Error GoTo 0
       Do While Not .EOF
         CANLOTES% = CANLOTES% + 1
         ReDim Preserve IDLOTING$(CANLOTES%), FELOTING(CANLOTES%), CANSALDOX#(CANLOTES%)
         IDLOTING$(CANLOTES%) = TRIM$(!IdenLote)
         FELOTING(CANLOTES%) = !FechAlta
         CANSALDOX#(CANLOTES%) = !CanAlta
       .MoveNext
       Loop
     End If
   End With
   '
   SQLX$ = " SELECT * FROM MOVISTO "
   SQLX$ = SQLX$ + "WHERE Cod_Inte = " & CIXI% & " "
   SQLX$ = SQLX$ + "AND CantSalid > 0 "
   SQLX$ = SQLX$ + "ORDER BY FechMov ASC "
   '
'AQUI VERIFICA TODOS LOS MOVIMIENTOS DE SALIDA Y AJUSTA QUE EN TODOS LOS
'CASOS LA SUMA DE LOS MOVIMIENTOS DE SALIDA POR LOTES SEA MENOR O IGUAL
'AL TOTAL DEL MOVIMIENTO DE SALIDA.
   Set MoviTemp = Stocks.OpenRecordset(SQLX$, dbOpenDynaset)
   With MoviTemp
     On Error Resume Next
     .MoveFirst
     If Err < 1 Then    ' no hay registros de salida
       '
       Do While Not .EOF
         SALIDPE# = ROUND(!CANTSALID, 1)
         TSALOT$ = ""
         On Error Resume Next
           TSALOT$ = !LotConsum
         On Error GoTo 0
         If SALIDPE# < 0.05 Then
           TSALOT$ = ""
           .Edit
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
             .Edit
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
             .Edit
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
   With MoviTemp
     On Error Resume Next
     .MoveFirst
     If Err < 1 Then    ' no hay registros de salida
       '
       Do While Not .EOF
         SALIDPE# = ROUND(!CANTSALID, 1)
         TSALOT$ = ""
         On Error Resume Next
           TSALOT$ = !LotConsum
         On Error GoTo 0
         If SALIDPE# < 0.05 Then GoTo 29
         '
         FEMOVI = FECHMOV
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
                 .Edit
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
   With LotIngTemp
     On Error Resume Next
     .MoveFirst
     If Err < 1 Then
       On Error GoTo 0
       Do While Not .EOF
         IDLOTE$ = TRIM$(!IdenLote)
         For KKI% = 1 To CANLOTES%
           If TRIM$(IDLOTING$(KKI%)) = IDLOTE$ Then
             If Abs(!CanSaldo - CANSALDOX#(KKI%)) >= 0.05 Then
               .Edit
               !CanSaldo = CANSALDOX#(KKI%)
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
   '
   Screen.MousePointer = 1
   If PORFIFO% > 0 Then
     Call COMUNI("Saldos por Lote Re-Calculados\por Criterio F.I.F.O.")
   End If
   '
End Sub

Sub ARMALILOT(CIXI%)
    '
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
    SQLX$ = " SELECT * FROM LOTINGRE "
    SQLX$ = SQLX$ + "WHERE Cod_Inte = " & CIXI% & " "
    SQLX$ = SQLX$ + "ORDER BY FechAlta ASC, IdenLote ASC "
    Set LotIngTemp = Stocks.OpenRecordset(SQLX$, dbOpenDynaset)
    '
    REBLA$ = REGBLAN$("LILOTSE")
    JJ& = 0
    With LotIngTemp
25     On Error Resume Next
       .MoveFirst
       If Err < 1 Then
         On Error GoTo 0
         Do While Not .EOF
           TTXX$ = REBLA$
             FEX = CVINT(!FechAlta)
           Call REPLA(TTXX$, TRIM$(!IdenLote), 1, 16)
           Call REPLA(TTXX$, FECHATEX$(FEX), 17, 8)
           Call REPLA(TTXX$, FORMATNUM$(!CanAlta, "F12,1"), 25, 12)
             SALFIS = !CanAlta - !CanConsu: If SALFIS < 0 Then SALFIS = 0
             SALOGI = !CanAlta - !CanConsu - !CanReser: If SALOGI < 0 Then SALOGI = 0
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
'
Sub VERILOTCO()
    '
' ESTA RUTINA TOMA LOS LOTES DECLARADOR POR DETALLE EN CAMPO 'LOTCONSUM' Y ABRE
' EL MOVIMIENTO EN TANTOS MOVIMIENTOS PARCIALES COMO LOTES PARCIALES SE CONSUMIERON
' COMO BLANQUEA EL CAMPO 'LOTCONSUM' LA SEGUNDA VEZ QUE ENTRA YA NO EJECUTA.
    '
    Dim VALOU(256)
    '
    Call ABRESTOCKS
    Screen.MousePointer = 11
    SQLX$ = "SELECT * FROM MOVISTO "
    SQLX$ = SQLX$ + "WHERE (LOTCONSUM <> '') " ' OR ISEMPTY(LOTCONSUM) = FALSE OR ISNULL(LOTCONSUM) = FALSE) "
    Set MOSEL = Stocks.OpenRecordset(SQLX$, dbOpenDynaset)
    '
    On Error Resume Next
    MOSEL.MoveFirst
    If Err Then GoTo 90 ' NO HAY REGISTROS PARA PROCESAR
    '
    On Error GoTo 0
    With MOSEL
       Do While Not .EOF
31        LCONSU$ = ""
          On Error Resume Next
          LCONSU$ = !LotConsum
          On Error GoTo 0
          ILOCON$ = TRIM$(Left$(LCONSU$, 24))
          CANCON = XVALO(Mid$(LCONSU$, 25, 8))
          If ILOCON$ <> "" Then
            If TRIM$(!IdenLote) = "" Then
              .Edit
              !IdenLote = ILOCON$
              !CANTSALID = CANCON
              .Update
              LCONSU$ = Mid$(LCONSU$, 33)
33            If TRIM$(LCONSU$) <> "" Then
                 For KKU& = 1 To .Fields.Count
                    VALOU(KKU&) = .Fields(KKU& - 1).Value
                 Next KKU&
                 ILOCON$ = TRIM$(Left$(LCONSU$, 24))
                 CANCON = XVALO(Mid$(LCONSU$, 25, 8))
                 If ILOCON$ <> "" Then
                    Movisto.AddNew
                    For KKU& = 1 To Movisto.Fields.Count
                       Movisto.Fields(KKU& - 1).Value = VALOU(KKU&)
                    Next KKU&
                    Movisto!IdenLote = ILOCON$
                    Movisto!CANTSALID = CANCON
                    Movisto.Update
                    LCONSU$ = Mid$(LCONSU$, 33)
                    GoTo 33
                 End If
              End If
            End If
          End If
          '
          .Edit
          !LotConsum = ""
          .Update
          '
       .MoveNext
       Loop
    End With
    '
90  On Error GoTo 0
    Screen.MousePointer = 1
    '
End Sub

Function LOTORIG$(NUBOLREC$)
    LOTEPROV$ = ""
    Call APERBASDAT("STOCKS")
         SQLX$ = " SELECT * FROM BOLRECEP "
         SQLX$ = SQLX$ + "WHERE IdenLote = '" & NUBOLREC$ & "' "
         SQLX$ = SQLX$ + "ORDER BY FechRecep ASC, IdenLote ASC "
         Set BolRecepTemp = Stocks.OpenRecordset(SQLX$, dbOpenSnapshot)
         '
         With BolRecepTemp
           On Error Resume Next
           .MoveFirst
           If Err < 1 Then
             LOTEPROV$ = !LoteCol_Prov
           End If
           On Error GoTo 0
        End With
      LOTORIG$ = LOTEPROV$
End Function

Function FECHVENLOT%(NUBOLREC$)
    FECHAVEN% = 0
    Call APERBASDAT("STOCKS")
         SQLX$ = " SELECT * FROM BOLRECEP "
         SQLX$ = SQLX$ + "WHERE IdenLote = '" & NUBOLREC$ & "' "
         SQLX$ = SQLX$ + "ORDER BY FechRecep ASC, IdenLote ASC "
         Set BolRecepTemp = Stocks.OpenRecordset(SQLX$, dbOpenSnapshot)
         '
         With BolRecepTemp
           On Error Resume Next
           .MoveFirst
           If Err < 1 Then
              FECHAVEN% = CVINT(!PLAZO_USO)
           End If
           On Error GoTo 0
        End With
      FECHVENLOT% = FECHAVEN%
End Function

Sub BLANPEDIPE(NCLIE%, FF%, NROREMIT$, CI%, CAN#, NUPEDIORI&, NUORITPE%)
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
10  SQLX$ = " SELECT * FROM PEDDETA "
    SQLX$ = SQLX$ + "WHERE NroPed = " & NUPEDIORI& & " "   ' COINCIDE NÚMERO DE PEDIDO
    SQLX$ = SQLX$ + "AND NroClie = " & NCLIE% & " "        ' COINCIDE CLIENTE
    SQLX$ = SQLX$ + "AND CodiInt = " & CI% & " "           ' COINCIDE CÓDIGO
    SQLX$ = SQLX$ + "AND Status >= 0 "                     ' APROBADOS
    SQLX$ = SQLX$ + "AND Status < 8 "                      ' NO ANULADOS NI CUMPLIDOS
    SQLX$ = SQLX$ + "AND CanPed > CantEnt "                ' QUE TENGAN SALDO DE ENTREGA
    If VUELTA% <= 1 Then
      SQLX$ = SQLX$ + " AND NuOrItem = " & NUORITPE% & " "  ' COINCIDE NRO DE ORDEN
    End If
    SQLX$ = SQLX$ + "ORDER BY FeSolEnt ASC"
    Set PEDSEL = Ventas.OpenRecordset(SQLX$, dbOpenDynaset)
    '
    With PEDSEL
      On Error Resume Next
      .MoveFirst
      If Err < 1 Then
         On Error GoTo 0
         Do While Not .EOF
            CSALDO# = !CANPED - !CantEnt: If CSALDO# < 0 Then CSALDO# = 0
            If !TipoPed = "ABIERTO" Then CSALDO# = CANBLA#
            CANAPLI# = CANBLA#: If CANAPLI# > CSALDO# Then CANAPLI# = CSALDO#
            .Edit
              !CantEnt = !CantEnt + CANAPLI#
              !FeUltEnt = CVDAT(FF%)
              !NroRemi = NROREMIT$
              If IsNull(!TipoPed) = True Or !TipoPed <> "ABIERTO" Then
                !CanSaldo = !CANPED - !CantEnt
              End If
            .Update
            CANBLA# = CANBLA# - CANAPLI#
            If CANBLA# < 0.05 Then
               PEDSEL.Close
               Exit Sub
            End If
         .MoveNext
         Loop
      End If
    End With
    '
    PEDSEL.Close
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

Function LOTESELEC$(CIXI%, CANTSA)
    '
    SELECLOT04.Label5 = FORMATNUM$(CANTSA, "F10.1")
    SELECLOT04.Label1 = CODEXT$(CIXI%)
    SELECLOT04.Label28 = ""
    '
    Call ARMALILOT(CIXI%)
    For U& = 1 To ULTREG("!LILOTSE")
       XX$ = REGLEIDO$("!LILOTSE", U&)
       IDLO$ = TRIM$(Left$(XX$, 16))
       If SALOTE(CIXI%, IDLO$, 1) >= CANTSA Then
          SELECLOT04.Label28 = IDLO$
          GoTo 90
       End If
    Next U&
    Call MENSERR(24, "No hay Lotes Disponibles\en Cantidad Suficiente\en el Depósito Elegido,")
    '
90  SCMA = Screen.MousePointer
    Screen.MousePointer = 1
    SELECLOT04.Show 1
    LOTESELEC$ = TRIM$(SELECLOT04.Label28)
    Screen.MousePointer = SCMA
End Function

Sub REIMPRIMEVALE()
    ' \\\OT-8-973-27/10/08
    Screen.MousePointer = 11
    CMO$ = TRIM$(Control$("", "CODCONSU"))
    If CMO$ = "" Then CMO$ = "VC"
    '
    Call APERBASDAT("STOCKS")
    SQLX$ = "SELECT DISTINCT DoPriCor, FechMov, ReferCor FROM MOVISTO "
    SQLX$ = SQLX$ + "WHERE CodiMovi = '" & CMO$ & "' "
    SQLX$ = SQLX$ + "ORDER BY DoPriCor DESC"
    '
    Set MOVISTEMP = Stocks.OpenRecordset(SQLX$, dbOpenSnapshot)
    UNRO& = 0
    On Error Resume Next
    With MOVISTEMP
      .MoveFirst
      If Err > 0 Then
         Call MENSERR(24, "No Existen Registros que Cumplan\con la Condición SQL.")
         On Error GoTo 0
         GoTo 99
      End If
      On Error GoTo 0
      j& = 0
      Call FRESHECHO("!SELVALE")
      Do While Not .EOF
         REMI$ = !DoPriCor
         FF% = CVINT(!FECHMOV)
         F1$ = FECHATEX$(FF%)
         REFE$ = SAFETEXT(!ReferCor)
         REFE1$ = F1$ + " - " + REFE$
         Z$ = Space$(40) + String$(8, 0)
         Call REPLA(Z$, REMI$, 1, 10)
         Call REPLA(Z$, REFE1$, 11, 30)
         j& = j& + 1
         Call GRAREG("!SELVALE", Z$, j&)
      .MoveNext
      Loop
    End With
    '
    Call SETULTREG("!SELVALE", j&)
    Call CIERRARCH("!SELVALE")
    Screen.MousePointer = 1
    '
15  RERE$ = REGSEL$("!SELVALE")
    If Len(RERE$) <= 4 Then GoTo 99
    
    NUDOCU$ = TRIM$(Mid$(RERE$, 1, 10))
    TIDOCUM$ = "Vale de Consumo de Materiales"
    Screen.MousePointer = 11
    '
    For UI& = ULTREG&("PDOCLST") To 1 Step -1
        XX$ = Left$(REGLEIDO$("PDOCLST", UI&), 64)
         If InStr(XX$, TIDOCUM$) > 4 Then
            If InStr(XX$, NUDOCU$) > 4 Then
                PPXX% = InStr(XX$, "Nro.")
                If PPXX% > 0 Then
                 XX$ = Left$(XX$, PPXX% - 1) + " " + Mid$(XX$, PPXX% + 4)
                End If
                If InStr(XX$, NUDOCU$) > 4 Then
                 DOCUNU& = CVS(Left$(XX$, 4))
                 GoTo 50
                End If
            End If
         End If
    Next UI&
    Screen.MousePointer = 1
40  Call COMUNI("No se Encontró el Documento Impreso")
    GoTo 15
    '
50  Call MUESTRADOC(DOCUNU&)
    GoTo 15
     
99  Exit Sub
    ' \\\OT-8-973-27/10/08-FIN
End Sub

Sub REIMPRIMESCRAP()
    ' \\\OT-8-973-27/10/08
    Screen.MousePointer = 11
    CMO$ = TRIM$(Control$("", "CODSCRAP"))
    If CMO$ = "" Then CMO$ = "SC"
    '
    Call APERBASDAT("STOCKS")
    SQLX$ = "SELECT DISTINCT DoPriCor, FechMov, ReferCor FROM MOVISTO "
    SQLX$ = SQLX$ + "WHERE CodiMovi = '" & CMO$ & "' "
    SQLX$ = SQLX$ + "ORDER BY DoPriCor DESC"
    '
    Set MOVISTEMP = Stocks.OpenRecordset(SQLX$, dbOpenSnapshot)
    UNRO& = 0
    On Error Resume Next
    With MOVISTEMP
      .MoveFirst
      If Err > 0 Then
         Call MENSERR(24, "No Existen Registros que Cumplan\con la Condición SQL.")
         On Error GoTo 0
         GoTo 99
      End If
      On Error GoTo 0
      j& = 0
      Call FRESHECHO("!SELVALE")
      Do While Not .EOF
         REMI$ = !DoPriCor
         FF% = CVINT(!FECHMOV)
         F1$ = FECHATEX$(FF%)
         REFE$ = SAFETEXT(!ReferCor)
         REFE1$ = F1$ + " - " + REFE$
         Z$ = Space$(40) + String$(8, 0)
         Call REPLA(Z$, REMI$, 1, 10)
         Call REPLA(Z$, REFE1$, 11, 30)
         j& = j& + 1
         Call GRAREG("!SELVALE", Z$, j&)
      .MoveNext
      Loop
    End With
    '
    Call SETULTREG("!SELVALE", j&)
    Call CIERRARCH("!SELVALE")
    Screen.MousePointer = 1
    '
15  RERE$ = REGSEL$("!SELVALE")
    If Len(RERE$) <= 4 Then GoTo 99
    
    NUDOCU$ = TRIM$(Mid$(RERE$, 1, 10))
    TIDOCUM$ = "Vale de Descarte de Scrap"
    Screen.MousePointer = 11
    '
    For UI& = ULTREG&("PDOCLST") To 1 Step -1
        XX$ = Left$(REGLEIDO$("PDOCLST", UI&), 64)
         If InStr(XX$, TIDOCUM$) > 4 Then
            If InStr(XX$, NUDOCU$) > 4 Then
                PPXX% = InStr(XX$, "Nro.")
                If PPXX% > 0 Then
                 XX$ = Left$(XX$, PPXX% - 1) + " " + Mid$(XX$, PPXX% + 4)
                End If
                If InStr(XX$, NUDOCU$) > 4 Then
                 DOCUNU& = CVS(Left$(XX$, 4))
                 GoTo 50
                End If
            End If
         End If
    Next UI&
    Screen.MousePointer = 1
40  Call COMUNI("No se Encontró el Documento Impreso")
    GoTo 15
    '
50  Call MUESTRADOC(DOCUNU&)
    GoTo 15
    
99  Exit Sub
    ' \\\OT-8-973-27/10/08-FIN
End Sub


Sub RECUSALSTOK()
   Call ABRESTOCKS
   Call CARDEPOS
   Dim SalCodDepo As Recordset
   
   Dim VESAL$()
   '
   NMS% = NUMAS%
   ReDim VESAL$(NMS%)
   For IX% = 1 To NMS%
      VESAL$(IX%) = String$(256, 0)
   Next IX%
   '
   SQLX$ = "SELECT COD_INTE, DEPOMOVI, SUM(CANTINGRE) AS INGRESO, SUM(CANTSALID) AS SALIDA FROM MOVISTO GROUP BY COD_INTE, DEPOMOVI ORDER BY COD_INTE, DEPOMOVI"
   '
   Set SalCodDepo = Stocks.OpenRecordset(SQLX$, dbOpenDynaset)
     
     With SalCodDepo
       On Error Resume Next
       .MoveFirst
       If Err < 1 Then
         Do While Not .EOF
            CIXI% = XVALO(!COD_INTE)
            If CIXI% < 1 Or CIXI% > NMS% Then GoTo 20
            SALDEPO# = XVALO(!INGRESO) - XVALO(!SALIDA)
            DEPO% = XVALO(!DEPOMOVI)
            COLUX% = 1
            For KI% = 1 To NUSU%
             If DEPO% = COSU%(KI%) Then
               COLUX% = KI%
               Exit For
             End If
            Next KI%
            Mid$(VESAL$(CIXI%), 4 * COLUX% - 3, 4) = MKS$(SALDEPO#)
20          .MoveNext
         Loop
       End If
       On Error GoTo 0
     End With
     
     For IX% = 1 To NMS%
      SalStock.FindFirst "CodiInte = " & IX%
      If SalStock.NoMatch = False Then
         SalStock.Edit
       Else
         SalStock.AddNew
         SalStock!CodiInte = IX%
      End If
        SalStock!VeSalSto = VESAL$(IX%)
        SalStock.Update
     Next IX%
     
     SalCodDepo.Close
     SalStock.Close
    
     
End Sub
Sub REVISANUMERO(NPED&)
    '
    'REVISION DE NUMEROS DE ORDEN EN LOS ITEMS DADO EL PEDIDO
    'EN EL CASO QUE ALGUNO ESTE EN 0 O REPTIDO LO CAMBIA POR EL EL PROXIMO INMEDIATO AL MAXIMO
    '
     Call ABREPEDCLI
     NUITANTERIOR% = 0
     SQLX$ = "SELECT * FROM PEDDETA "
     SQLX$ = SQLX$ + " WHERE NroPed = " & NPED& & " "
     SQLX$ = SQLX$ + " AND Status < 8 "
     SQLX$ = SQLX$ + " ORDER BY NUORITEM ASC"
     Set PEDDETTEMP = Ventas.OpenRecordset(SQLX$, dbOpenDynaset)
     With PEDDETTEMP
      Do While Not .EOF
         NUIT% = XVALO(!NuorItem)
         If NUIT% = 0 Or NUIT% = NUITANTERIOR% Then ' SI ES 0  O ESTA REPETIDO
             .Edit
            !NuorItem = PRONULIBRE%(NPED&)
            .Update
         End If
         NUITANTERIOR% = NUIT%
        .MoveNext
      Loop
     End With
    
End Sub
Function PRONULIBRE%(NUPEDIDO&) ' DEVULELVE DADO EL PEDIDO EL SIGUIENTE NUMERO DE ORDEN LIBRE
     '
     PRONULIB% = 0
     '
     SQLX$ = "SELECT MAX(NUORITEM)AS MAXIMO FROM PEDDETA "
     SQLX$ = SQLX$ + "WHERE NroPed = " & NUPEDIDO& & " "
     SQLX$ = SQLX$ + "AND Status < 8 "
     Set PEDDETTEMP = Ventas.OpenRecordset(SQLX$, dbOpenSnapshot)
     With PEDDETTEMP
       NUMEORMAX% = XVALO(!MAXIMO)
     End With
     PRONULIB% = 1 + NUMEORMAX%
     '
     PRONULIBRE% = PRONULIB%
End Function

