Attribute VB_Name = "CUEC_FUN_SQL"
' los siguentes son generales en varias de las rutinas
Public CoComCor$
Public CoComLar$
Public NumeClie
Public NumeClieL
Public NumeProv%
Public FechEmi
Public NiveCla%
' los siguientes para CajaBanc
Public Referen$
Public FeContab As Date
Public FeAcredi As Date
Public CuentaCo%
Public CentroCos$
Public ItConcep$ 'nuevo p/ concepto
Public TiMovi$
Public OpMovi$
Public VaMovi$
Public Descrip$
Public NumeTal As Integer
Public NumeCom As Long
Public StatInit
Public ImpoDeb As Double
Public ValoDola As Double
Public BasImpo As Double
Public Observa$
Public NumCai$
Public FechCai As Date
Public InclIva As String
Public IdentiCae$
Public NumeroCae$
Public ObservaFactu$
Public CondPag%
Public INCOTERM$
Public NroPedi$
Public OrdCompra$
Public NCliReal
Public AtraCance&
Public AtraCheque&
'descuentos (pordescu y descuento aparte de gabal)
Public Descue1$
Public Descue2$
Public Impodesc1 As Double
Public Impodesc2 As Double
Public COTI_DOLAR As Double
Public CodigoImpuesto As Integer
Public CodigoRegimen As Integer
'
' los siguientes para grabar vencimientos e importes
Public IBruTot#
Public IIvaTot#
Public NuorVen%
Public FeVenc
Public ImVenc#
'
' los siguientes para grabar aplicacion de pagos
Public ImpApli#
Public RefApli$
Public DocOrLar$
Public DocOrCor$
Public TotOrig#
Public IvaOrig#
Public MonOri%
Public ImpApliOr#


Public COTIMONEM#
'
Public EsPorPedido& ' Define si va a la factura desde pedido ( 0 normal) (1 viene de pedido)
'


Sub GRACABAN()
   'jandy sql ver funcion noclose% para sql
   
   If NOCLOSE% = 0 Then Call APERBASDAT("CABAN")
   
   HOII% = HOY(HOS$)
   strsql = "Select * from CajaBanc where CodiCom_Lar='" & Left$(CoComLar$, 24) & "'"
   Dim CajaBanc As ADODB.Recordset
   Set CajaBanc = New ADODB.Recordset
   
25 On Error Resume Next
   CajaBanc.Open FILTSQL$(strsql), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
   If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
      On Error GoTo 0
      Call CapturaErrorOpen
      GoTo 25
   End If
   On Error GoTo 0
   With CajaBanc
     .AddNew   ' YA CONVERTIDO
     On Error Resume Next
     !CodiEmpr = CodiEmp%
     On Error GoTo 0
      !CodiCom_Cor = Left$(CoComCor$, 12)
      !CodiCom_Lar = Left$(CoComLar$, 24)
      !RefeCom = Left$(Referen$, 48)
      !FECHCONTA = SafeDate(FeContab)
      !FechAcred = SafeDate(FeAcredi)
      !CUECONTAI = CuentaCo%
      !CenCosto = Left$(CentroCos$, 8)
      !ItemConc = Left$(ItConcep$, 8) 'nuevo p/ concepto
        NumCli = NumeClieL
        If NumCli = 0 Then NumCli = NumeClie
      !nuclien = NumCli
      !NuProve = NumeProv%
      !TipoMovim = Left$(TiMovi$, 8)
      !OpciMovim = Left$(OpMovi$, 8)
      !VARIMOVIM = Left$(VaMovi$, 8)
      !DescriMov = Left$(Descrip$, 48)
      !FechEmisi = FechEmi
      !CodiEmpr = CodiEmp%
      !CodiSucu = CodiSuc%
      !NumeTalo = NumeTal
      !NumeComp = NumeCom
      !Status = StatInit
        On Error Resume Next
      !Mone_Emis = MONEMI%
      !Tipo_Cam = TICAMONE(MONEMI%)
        If COTIMONEM# >= 0.1 Then !Tipo_Cam = COTIMONEM#
        On Error GoTo 0
      !ValAbsComp = Abs(ImpoDeb#)
      If ImpoDeb# > 0 Then
          !Idebecomp = ImpoDeb#
          !IHABECOMP = 0
        Else
          !Idebecomp = 0
          !IHABECOMP = Abs(ImpoDeb#)
      End If
      !Impo_Dola = Abs(ValoDola#)
      !BASEIMPO = Abs(BasImpo#)
      !NROCAI = TRIM$(Left$(NumCai$, 24))
        On Error Resume Next
      !IncSubIva = InclIva$
        On Error GoTo 0
      !FEVTOCAI = SafeDate(FechCai)
      !OBSERGEN = Left$(Observa$, 64)
      !IdenAlta = Left$(IDACTU$, 8)          ' FECHA - HORA - USUARIO - TERMINAL
      !IdenBaja = String$(6, 0)
      !FeHoReal = Now()
      !NuAsiDiar = ""
      !Niv_Clasi = NiveCla%
      !CAE = TRIM$(Left$(NumeroCae$, 32))
      !IDCAE = TRIM$(Left$(IdentiCae$, 32))
      !IDCAENUM = XVALO(IdentiCae$)
      !CODCARPIMP = CODCARPE$          ' NO HACE FALTA EL SAFETEXT
      !Notas = ObservaFactu$
      !INCOTERMS = Left$(INCOTERM$, 12) ' (EXPORTACION)
      !CondPag = CondPag% ' CONDICION DE PAGO
      !PEDIDOS = NroPedi$ ' LISTA DE PEDIDOS QUE SE FACTURAN
      !OCompra = OrdCompra$ ' CONTENIDO DEL TEXT14 (LISTA DE ORDENES DE COMPRA SEGUN LOS PEDIDOS)
      !Cli_Real = NCliReal
      !SimMone = SIMBOLPES$(MONEMI%) ' GRABA EL SIMBOLO DE LA MONEDA
      !AtraCance = AtraCance&
      !AtraCheque = AtraCheque&
      '
      'DEscuentos general y descuento especial (Gabal) para presentar en F.Electronica.
      !Descu1 = Descue1$
      !Descu2 = Descue2$
      !Impodesc1 = Impodesc1#
      !Impodesc2 = Impodesc2#
      '
      'SE VALIDA PARA QUE SOLO GRABE EN FACTURAS
      If TiMovi$ = "FVEN" Or TiMovi$ = "VENTA" Or UCase$(OpMovi$) = "DEFI" Then
        MONX% = XVALO(Control("FACTURA", "MONTEXFA")) 'MONEDA CONSIDERADA PARA PRESENTAR EN TEXTO DE FACTURA
        If MONEMI% = 1 And MONX% >= 1 Then
          COTI_DOLAR# = TICAMONE(MONX%)
        End If
        !COTIZA_DOLAR = COTI_DOLAR#
      Else
        If COTI_DOLAR# <= 0 Then COTI_DOLAR# = 1
        !COTIZA_DOLAR = COTI_DOLAR#
      End If
      !CodigoImpuesto = CodigoImpuesto
      !CodigoRegimen = CodigoRegimen
      '
     .Update
   End With
   CajaBanc.Close
   Set CajaBanc = Nothing
   'If NOCLOSE% = 0 Then CajaBanc.Close
   
   '
End Sub
'
Sub GRACREPEN()
   '
   If NOCLOSE% = 0 Then Call APERBASDAT("CABAN")
   'jandy SQL
   strsql = "Select * from SacrePen where CodiCom_Lar='" & Left$(CoComLar$, 24) & "'"
   Dim Sacrepen As ADODB.Recordset
   Set Sacrepen = New ADODB.Recordset
25 On Error Resume Next
   Sacrepen.Open FILTSQL$(strsql), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
   If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
      On Error GoTo 0
      Call CapturaErrorOpen
      GoTo 25
   End If
   On Error GoTo 0

   With Sacrepen
     .AddNew   ' YA CONVERTIDO
     On Error Resume Next
     !CodiEmpr = CodiEmp%
     On Error GoTo 0
      !CodiEmpr = CodiEmp%
      !CodiSucu = CodiSuc%
      !nuclien = NumeClie
      !NuProve = NumeProv%
      !CodCliPro = Left$(CODICLI$(NumeClie - NumeProv%), 16)
      !RaSoCliPro = Left$(RASOCLI$(NumeClie - NumeProv%), 48)
      !CodiCom_Cor = Left$(CoComCor$, 12)
      !CodiCom_Lar = Left$(CoComLar$, 24)
      !FechEmisi = SafeDate(FechEmi)
      !IBruComp = IBruTot#
      !IIvaComp = IIvaTot#
      !OrVencim = NuorVen%
      !FEVENCIM = SafeDate(FeVenc)
      
      !Mone_Emis = MONEMI%
      If MONEMI% < 0 Then !Mone_Emis = 0
      On Error GoTo 0
    
      TICAMO# = TICAMONE(MONEMI%)
      If COTIMONEM# >= 0.1 Then TICAMO# = COTIMONEM#
      If TICAMO# < 0.05 Then TICAMO# = 1
      If MONEMI% > 1 Then
            !Ibru_Dola = ROUND(IBruTot# / TICAMO#, 2)
      Else
            !Ibru_Dola = 0
      End If
      
      If ImVenc# > 0 Then
          !SaldAcre = ImVenc#
          !SaldDebe = 0
          !SaldAcre_DOLA = ROUND#(ImVenc# / TICAMO#)
          !SaldDebe_DOLA = 0
        Else
          !SaldDebe = Abs(ImVenc#)
          !SaldAcre = 0
          !SaldDebe_DOLA = ROUND#(Abs(ImVenc# / TICAMO#))
          !SaldAcre_DOLA = 0
      End If
       If MONEMI% = (-1) Then
          !SaldAcre_DOLA = 0
          !SaldDebe_DOLA = 0
      End If
      !Status = 0
      !Niv_Clasi = NiveCla%
     .Update
   End With
   Sacrepen.Close
   Set Sacrepen = Nothing
   
   'If NOCLOSE% = 0 Then SacrePen.Close
   '
End Sub

Sub GRADEUPEN()
   '
   Call APERBASDAT("CABAN")
   
   strsql = "Select * from SaDeuPen where CodiCom_Lar='" & Left$(CoComLar$, 24) & "'"
   Dim SaDeuPen As ADODB.Recordset
   Set SaDeuPen = New ADODB.Recordset
25 On Error Resume Next
   SaDeuPen.Open FILTSQL$(strsql), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
   If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
      On Error GoTo 0
      Call CapturaErrorOpen
      GoTo 25
   End If
   On Error GoTo 0

   With SaDeuPen
     .AddNew   ' YA CONVERTIDO
     On Error Resume Next
     !CodiEmpr = CodiEmp%
     On Error GoTo 0
      !CodiEmpr = CodiEmp%
      !CodiSucu = CodiSuc%
      !nuclien = NumeClie
      !NuProve = NumeProv%
      !CodCliPro = Left$(CODICLI$(NumeClie - NumeProv%), 16)
      !RaSoCliPro = Left$(RASOCLI$(NumeClie - NumeProv%), 48)
      !CodiCom_Cor = Left$(CoComCor$, 12)
      !CodiCom_Lar = Left$(CoComLar$, 24)
      !FechEmisi = SafeDate(FechEmi)
        On Error Resume Next
      !Mone_Emis = MONEMI%
        On Error GoTo 0
      !IBruComp = IBruTot#
      !IIvaComp = IIvaTot#
        TICAMO# = TICAMONE(MONEMI%)
        If COTIMONEM# >= 0.1 Then TICAMO# = COTIMONEM#
        If TICAMO# < 0.05 Then TICAMO# = 1
      If MONEMI% > 1 Then
           !Ibru_Dola = ROUND(IBruTot# / TICAMO#, 2)
         Else
           !Ibru_Dola = 0
      End If
      !OrVencim = NuorVen%
      !FEVENCIM = SafeDate(FeVenc)
      If ImVenc > 0 Then
          !SaldDebe = ImVenc#
          !SaldAcre = 0
          !SaldDebe_DOLA = ROUND#(ImVenc# / TICAMO#)
          !SaldAcre_DOLA = 0
        Else
          !SaldAcre = Abs(ImVenc#)
          !SaldDebe = 0
          !SaldAcre_DOLA = Abs(ImVenc# / TICAMO#)
          !SaldDebe_DOLA = 0
      End If
      If Left$(CoComLar$, 2) = "RB" Or MONEMI% < 2 Then
         !SaldDebe_DOLA = 0
         !SaldAcre_DOLA = 0
      End If
      !Status = 0
     .Update
   End With
   SaDeuPen.Close
   
   Set SaDeuPen = Nothing
   
End Sub

Sub GRAPLIPAG()
   
   Dim Aplipago As ADODB.Recordset
   Set Aplipago = New ADODB.Recordset
   strsql = "Select * from ApliPago where NuProve=" & NumeProv%
25 On Error Resume Next
   Aplipago.Open FILTSQL$(strsql), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
   If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
      On Error GoTo 0
      Call CapturaErrorOpen
      GoTo 25
   End If
   On Error GoTo 0

   With Aplipago
     .AddNew   ' YA CONVERTIDO
     On Error Resume Next
     !CodiEmpr = CodiEmp%
     On Error GoTo 0
      !CodiEmpr = CodiEmp%
      !CodiSucu = CodiSuc%
      !NuProve = NumeProv%
      !CodiProv = Left$(CODICLI$((-1) * NumeProv%), 16)
      !RasoProv = Left$(RASOCLI$((-1) * NumeProv%), 48)
      !codocap_lar = Left$(CoComLar, 18)
      !CODOCAP_COR = Left$(CoComCor$, 12)
      !IMPOAPLI = ImpApli#
      !refeapli = Left$(RefApli$, 32)
      !codocor_lar = Left$(DocOrLar$, 18)
      !CODOCOR_COR = Left$(DocOrCor$, 12)
      !ItotOrig = TotOrig#
      !IIVAORIG = IvaOrig#
      !ImpoApli_Ori = ImpApliOr#
     .Update
   End With
   Aplipago.Close
   Set Aplipago = Nothing
End Sub
'
Sub GRAPLICOB()
   Call APERBASDAT("CABAN")
   
   Dim Aplicobra As ADODB.Recordset
   Set Aplicobra = New ADODB.Recordset
   strsql = "Select * from ApliCobra where NuClien=" & NumeClie
25 On Error Resume Next
   Aplicobra.Open FILTSQL$(strsql), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
   If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
      On Error GoTo 0
      Call CapturaErrorOpen
      GoTo 25
   End If
   On Error GoTo 0
   
   With Aplicobra
     .AddNew   ' YA CONVERTIDO
     On Error Resume Next
     !CodiEmpr = CodiEmp%
     On Error GoTo 0
      !CodiEmpr = CodiEmp%
      !CodiSucu = CodiSuc%
      !nuclien = NumeClie
      !CODICLIE = Left$(CODICLI$(NumeClie), 16)
      !RasoClie = Left$(RASOCLI$(NumeClie), 48)
      !codocap_lar = Left$(CoComLar, 18)
      !CODOCAP_COR = Left$(CoComCor$, 12)
      !IMPOAPLI = ImpApli#
      !refeapli = Left$(RefApli$, 32)
      !codocor_lar = Left$(DocOrLar$, 18)
      !CODOCOR_COR = Left$(DocOrCor$, 12)
      !ItotOrig = TotOrig#
      !IIVAORIG = IvaOrig#
     .Update
   End With
   Aplicobra.Close
   Set Aplicobra = Nothing
End Sub
'
Sub RAPLIPRO(NCX%)
    '
    'If InStr(UCase$(EPAC$), "FERVI") > 0 Then Exit Sub  ' provisorio FFFERVI
    If DERACCE%("RAPLIPRO", "Reconstrucción Automática Composición de Deuda") < 2 Then
       Call CIERRARCH("*.*")
       Call FINAL("")
    End If
    '
    Dim DOCULAR$(), DOCUCOR$(), IMPO#(), IORI#(), FFDOC%(), FVENDOC%()
    Dim STATDOC%(), CODIEM%(), CODISU%(), NICLADOC%(), SALDOLA#()
    RAMAX& = 0
    ATENTI% = 1
    '
    If TICUEPRO%((-1) * NCX%) = 1 Then
       Call APERBASDAT("CABAN")
       NUDOCU% = 0
       GoTo 24
    End If
    '
100 Screen.MousePointer = 11
    '
    SALDO# = 0: SARRAS# = 0
    NUDOCU% = 0
    RFF$ = RECFILT$("CUECOPRO", 1, MKI$(NCX%))
    For KU& = 1 To Len(RFF$) Step 4
      RELE& = CVS(Mid$(RFF$, KU&, 4))
      XX$ = REGLEIDO$("CUECOPRO", RELE&)
      NUDOCU% = NUDOCU% + 1
      If NUDOCU% > RAMAX& Then
        RAMAX& = RAMAX& + 128
        ReDim Preserve DOCULAR$(RAMAX&), DOCUCOR$(RAMAX&), IORI#(RAMAX&), IMPO#(RAMAX&), FFDOC%(RAMAX&), FVENDOC%(RAMAX&), CODIEM%(RAMAX&), CODISU%(RAMAX&), STATDOC%(RAMAX&), NICLADOC%(RAMAX&), SALDOLA#(RAMAX&)
      End If
      DOCULAR$(NUDOCU%) = TRIM$(Mid$(XX$, 33, 24))
      DOCUCOR$(NUDOCU%) = TRIM$(Mid$(XX$, 57, 12))
      IMPO#(NUDOCU%) = CVD(Mid$(XX$, 25, 8))
      IORI#(NUDOCU%) = Abs(IMPO#(NUDOCU%))
      FFDOC%(NUDOCU%) = CVI(Mid$(XX$, 3, 2))
      CODIEM%(NUDOCU%) = CVI(Mid$(XX$, 69, 2))
      CODISU%(NUDOCU%) = CVI(Mid$(XX$, 71, 2))
      STATDOC%(NUDOCU%) = CVI(Mid$(XX$, 73, 2))
      NICLADOC%(NUDOCU%) = 0
    Next KU&
    '
    Call APERBASDAT("CABAN")
    
    Dim CuecoTemp As ADODB.Recordset
    Set CuecoTemp = New ADODB.Recordset
    CuecoTemp.CursorLocation = adUseClient
    SQLX$ = "SELECT * FROM CAJABANC "
    SQLX$ = SQLX$ + "WHERE NuProve = " & NCX% & " "
    SQLX$ = SQLX$ + "AND (TipoMovim = 'FCOM' "
    SQLX$ = SQLX$ + "OR TipoMovim = 'AJUP' "
    SQLX$ = SQLX$ + "OR TipoMovim = 'OPAG') "
    SQLX$ = SQLX$ + "ORDER BY FechEmisi, FeHoReal ASC;"
    'Set CuecoTemp = CueCorri.OpenRecordset(SQLX$, dbOpenSnapshot)
25  On Error Resume Next
    CuecoTemp.Open FILTSQL$(SQLX$), FLEXCONN, adOpenStatic, adLockReadOnly, adCmdText
    If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
      On Error GoTo 0
      Call CapturaErrorOpen
      GoTo 25
    End If
    On Error GoTo 0
    
    '
    With CuecoTemp
      On Error Resume Next
      .MoveLast
      If Err < 1 Then
        .MoveFirst
        Do While Not .EOF
          ' ESTE BLOQUE AGREGADO PARA UNIFICAR LOS AJUSTES POR DIFERENCIA DE CAMBIO
          ' JUNTO CON LA O/PAGO CORRESPONDIENTE EN EL CASO DE GRAMY - VANINA
          COCOLA$ = !CodiCom_Lar
          COCOCO$ = !CodiCom_Cor
          If Left$(COCOLA$, 4) = "AS-X" Then
            If Left$(!OBSERGEN, 6) = "O.Pago" Then
              COCOLA$ = "OP-X" + Mid$(!CodiCom_Lar, 5)
              COCOCO$ = "OP-X" + Mid$(!CodiCom_Cor, 5)
            End If
          End If
          If Left$(COCOLA$, 2) = "OP" Then
              For KKII% = 1 To NUDOCU%
                If DOCULAR$(KKII%) = COCOLA$ Then
                  IMPO#(KKII%) = IMPO#(KKII%) + !Idebecomp - !IHABECOMP
                  MEMI% = !Mone_Emis
                  If MEMI% > 1 Then
                     TICAEMI = !Tipo_Cam
                     If TICAEMI > 0.1 Then
                       SALDOLA#(KKII%) = SALDOLA#(KKII%) + (!Idebecomp - IHABECOMP) / TICAEMI
                     End If
                  End If
                  GoTo 119
                End If
              Next KKII%
          End If
          ' FIN DE BLOQUE AGREGADO
        
        
          NUDOCU% = NUDOCU% + 1
          If NUDOCU% > RAMAX& Then
            RAMAX& = RAMAX& + 128
            ReDim Preserve DOCULAR$(RAMAX&), DOCUCOR$(RAMAX&), IORI#(RAMAX&), IMPO#(RAMAX&), FFDOC%(RAMAX&), FVENDOC%(RAMAX&), CODIEM%(RAMAX&), CODISU%(RAMAX&), STATDOC%(RAMAX&), NICLADOC%(RAMAX&), SALDOLA#(RAMAX&)
          End If
          DOCULAR$(NUDOCU%) = !CodiCom_Lar
          DOCUCOR$(NUDOCU%) = !CodiCom_Cor
          IMPO#(NUDOCU%) = !Idebecomp - !IHABECOMP
          IORI#(NUDOCU%) = !ValAbsComp
          FFDOC%(NUDOCU%) = CVINT(!FechEmisi)
          CODIEM%(NUDOCU%) = !CodiEmpr
          CODISU%(NUDOCU%) = !CodiSucu
          STATDOC%(NUDOCU%) = !Status
          NICLADOC%(NUDOCU%) = XVALO(!Niv_Clasi)
119     .MoveNext
        Loop
      End If
    End With
    CuecoTemp.Close
    Set CuecoTemp = Nothing
    '
    ' blanquea si pasa por cero
    FIN& = NUDOCU%
    SUMAPEN# = 0
    For KKI% = 1 To NUDOCU%
      SUMAPEN# = SUMAPEN# + IMPO#(KKI%)
      If Abs(SUMAPEN#) < 0.005 Then
        For KKJ% = 1 To KKI%
          IMPO#(KKJ%) = 0
        Next KKJ%
      End If
    Next KKI%
    '
    Dim ApliTemp As ADODB.Recordset
    Set ApliTemp = New ADODB.Recordset
    ApliTemp.CursorLocation = adUseClient
    SQLX$ = "SELECT * FROM APLIPAGO "
    SQLX$ = SQLX$ + "WHERE NuProve = " & NCX% & " "
    'Set ApliTemp = CueCorri.OpenRecordset(SQLX$, dbOpenSnapshot)
26  On Error Resume Next
    ApliTemp.Open FILTSQL$(SQLX$), FLEXCONN, adOpenStatic, adLockReadOnly, adCmdText
    If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
      On Error GoTo 0
      Call CapturaErrorOpen
      GoTo 26
    End If
    On Error GoTo 0
    '
'    With ApliTemp
'      On Error Resume Next
'      .MoveLast
'      If Err < 1 Then
'        .MoveFirst
'        Do While Not .EOF
'          IAPLI# = !IMPOAPLI
'          RECI$ = !CODOCAP_LAR
'          FACU$ = !CODOCOR_LAR
'          '
'          IRECIMA# = 0: IFACMAX# = 0
'          For KKI% = 1 To NUDOCU%
'             If DOCULAR$(KKI%) = RECI$ Then
'               If IMPO#(KKI%) > 0 Then
'                 IRECIMA# = IRECIMA# + IMPO#(KKI%)
'               End If
'             End If
'             If DOCULAR$(KKI%) = FACU$ Then
'               If IMPO#(KKI%) < 0 Then
'                 IFACMAX# = IFACMAX# - IMPO#(KKI%)
'               End If
'             End If
'          Next KKI%
'          '
'          If IAPLI# > IRECIMA# Then IAPLI# = IRECIMA#
'          If IAPLI# > IFACMAX# Then IAPLI# = IFACMAX#
'          '
'          IRECIMA# = IAPLI#
'          IFACMAX# = IAPLI#
'          '
'          If IAPLI# > 0 Then
'             For KKI% = 1 To NUDOCU%
'                If DOCULAR$(KKI%) = RECI$ Then
'                   If IMPO#(KKI%) > 0 Then
'                      IAP1# = IRECIMA#
'                      If IAP1# > Abs(IMPO#(KKI%)) Then IAP1# = Abs(IMPO#(KKI%))
'                      IMPO#(KKI%) = IMPO#(KKI%) - IAP1#
'                      IRECIMA# = IRECIMA# - IAP1#
'                   End If
'                End If
'                If DOCULAR$(KKI%) = FACU$ Then
'                   If IMPO#(KKI%) < 0 Then
'                      IAP2# = IFACMAX#
'                      If IAP2# > Abs(IMPO#(KKI%)) Then IAP2# = Abs(IMPO#(KKI%))
'                      IMPO#(KKI%) = IMPO#(KKI%) + IAP2#
'                      IFACMAX# = IFACMAX# - IAP2#
'                   End If
'                End If
'             Next KKI%
'          End If
'          '
'        .MoveNext
'        Loop
'      End If
'    End With
    With ApliTemp
      On Error Resume Next
      .MoveLast
      If Err < 1 Then
        .MoveFirst
        Do While Not .EOF
          IAPLI# = !IMPOAPLI
          IAPDOL# = !ImpoApli_Ori
          RECI$ = TRIM$(!codocap_lar)
          FACU$ = TRIM$(!codocor_lar)
          COEFIX = 1: If IAPLI# > 0 Then COEFIX = IAPLI# / IAPDOL#

          '
          DOCUAP% = 0: DOCUOR% = 0
          For KKI% = 1 To NUDOCU%
            If DOCULAR$(KKI%) = RECI$ And DOCUAP% < 1 Then DOCUAP% = KKI%
            If DOCULAR$(KKI%) = FACU$ And DOCUOR% < 1 Then DOCUOR% = KKI%
          Next KKI%
          '
          If DOCUAP% > 0 And DOCUOR% > 0 Then
            SIGAP% = 1: If IAPLI# < 0 Then SIGAP% = (-1)
            If SIGAP% * IAPLI# > (-1) * SIGAP% * IMPO#(DOCUOR%) Then IAPLI# = (-1) * SIGAP% * IMPO#(DOCUOR%)
            IMPO#(DOCUOR%) = IMPO#(DOCUOR%) + IAPLI#
            IMPO#(DOCUAP%) = IMPO#(DOCUAP%) - IAPLI#
          End If
          '
        .MoveNext
        Loop
      End If
    End With
    '
    ApliTemp.Close
    Set ApliTemp = Nothing
    '
    SUMAPEN# = 0
    FIN& = NUDOCU%
    For KKI% = 1 To NUDOCU%
      U& = KKI%: Call TRACE(20, U&, FIN&)
      SUMAPEN# = SUMAPEN# + IMPO#(KKI%)
      If Abs(SUMAPEN#) < 0.005 Then
        For KKJ% = 1 To KKI%
          IMPO#(KKJ%) = 0
        Next KKJ%
      End If
    Next KKI%
    '
    'HASTA ACA LLEGUE 11/11/14
    If NCX% <> 5000 Then    ' que no lo haga en DOSIVAC - VARIOS EN CUENTA CORRIENTE
       For KKI% = 1 To NUDOCU%
         For KKJ% = KKI% - 1 To 1 Step -1
           If Abs(IMPO#(KKI%) + IMPO#(KKJ%)) < 0.005 Then
             SALDOLA#(KKI%) = SALDOLA#(KKI%) - IAPDOL#
             If MODOMONEI% <> 2 Then SALDOLA#(KKI%) = 0
             IMPO#(KKI%) = 0
             IMPO#(KKJ%) = 0
           End If
         Next KKJ%
       Next KKI%
    End If
    '
    KKM% = 0
    SUMAPEN# = 0
    For KKI% = 1 To NUDOCU%
      If STATDOC%(KKI%) = 8 Then
        SUMAPEN# = SUMAPEN# + IMPO#(KKI%)
        SUMAPENDOL# = SUMAPENDOL# + SALDOLA#(KKI%)
      End If
    Next KKI%
    '
    If Abs(SUMAPEN#) < 0.005 Then
      For KKI% = 1 To NUDOCU%
        If STATDOC%(KKI%) = 8 Then
          SALDOLA#(KKI%) = 0
          IMPO#(KKI%) = 0
        End If
      Next KKI%
      GoTo 19
    End If
    '
    For KKI% = NUDOCU% To 1 Step -1
      If STATDOC%(KKI%) = 8 Then
        If Sgn(IMPO#(KKI%)) <> Sgn(SUMAPEN#) Then
            IMPO#(KKI%) = 0
          Else
            If Abs(SUMAPEN#) >= Abs(IMPO#(KKI%)) Then
                SUMAPEN# = SUMAPEN# - IMPO#(KKI%)
              Else
                IMPO#(KKI%) = SUMAPEN#
                SUMAPEN# = 0
            End If
        End If
      End If
    Next KKI%
    '
    'BUSCA LAS FECHAS DE VENCIMIENTO ANTERIORES
Dim Sacrepen As ADODB.Recordset
19  For U& = 1 To NUDOCU%
      If Abs(IMPO#(U&)) >= 0.005 Then
        ABUS$ = "Select * from Sacrepen where "
        ABUS$ = ABUS$ + "NuProve = " & NCX% & " "
        ABUS$ = ABUS$ + "AND CodiEmpr = " & CODIEM%(U&) & " "
        ABUS$ = ABUS$ + "AND CodiSucu = " & CODISU%(U&) & " "
        ABUS$ = ABUS$ + "AND NuClien = 0 "
        ABUS$ = ABUS$ + "AND NuProve = " & NCX% & " "
        ABUS$ = ABUS$ + "AND CodiCom_Lar = '" & DOCULAR$(U&) & "' "
        Set Sacrepen = FLEXCONN.Execute(FILTSQL$(ABUS$))
        '
        FVENDOC%(U&) = FFDOC%(U&)
        
        With Sacrepen
         If Not (.EOF And .BOF) Then
         ' .FindFirst ABUS$
         ' If .NoMatch = False Then
            FVENDOC%(U&) = CVINT(!FEVENCIM)
            If FVENDOC%(U&) < FFDOC%(U&) Then
              FVENDOC%(U&) = FFDOC%(U&)
            End If
          End If
        End With
        Sacrepen.Close
        Set Sacrepen = Nothing
      End If
    Next U&
    '
    If CONTROLGRABA% = 1 Then GoTo 24
    Screen.MousePointer = 1
    '
    ' SE COMENTA PARA QUE SIEMPRE ACTUALICE LA COMPOSICION DE DEUDA. AL PREGUNTAR SE DEJABA LA POSIBILIDAD
    ' DE QUE EL USUARIO OPTE POR NO ACTUALIZAR LA DEUDA Y DESPUES EL REPORTE DE CUENTAS POR PAGAR APARECIA
    ' DESACTUALIZADO.
    '
'    TTXX$ = "Se ha Completado la Reconstrucción\de la Composición de la Deuda de\   \'" + TRIM$(RASOCLI$((-1) * NCX%))
'    If COMALTER%(TTXX$ + "'\   \Desea Grabar la Nueva Composición de la Deuda ?\\Cancelar\Grabar") <> 2 Then
'       GoTo 999
'    End If
    '
24  Screen.MousePointer = 11
    YAPRIRE% = 0
    ABUS$ = "NuProve = " & NCX% & " "
    
    Dim MIRS As ADODB.Recordset
    strsql = "Select * from Sacrepen "
    strsql = strsql & "where NuProve = " & NCX%
    Set MIRS = FLEXCONN.Execute(FILTSQL$(strsql))
    With MIRS
      If Not (.EOF And .BOF) Then
         MIRS.Close
         strsql_D = "Delete from Sacrepen where NuProve = " & NCX%
         Call EJECUTACOMANDO(strsql_D)
      End If
    End With
    On Error Resume Next
    MIRS.Close
    Set MIRS = Nothing
    On Error GoTo 0
    HOII% = HOY(HO$)
    FIN& = NUDOCU%
    For U& = 1 To NUDOCU%
      Call TRACE(20, U&, FIN&)
      If Abs(IMPO#(U&)) >= 0.005 Then
        CoComCor$ = DOCUCOR$(U&)
        CoComLar$ = DOCULAR$(U&)
'        CodiEmp% = CODIEM%(U&)
'        CodiSuc% = CODISU%(U&)
        NumeClie = 0
        NumeProv% = NCX%
        FechEmi = CVDAT(FFDOC%(U&))
        IBruTot# = IORI#(U&)
        '
        If DOCULAR$(U&) <> DOCULARANT$ Then NuorVen% = 0
        NuorVen% = NuorVen% + 1
        '
        FeVenc = CVDAT(FVENDOC%(U&)) ' TOMADOS DE REGISTRO ANTERIOR
        ImVenc# = (-1) * IMPO#(U&)
        DOCULARANT$ = DOCULAR$(U&)
        NiveCla% = NICLADOC%(U&)
        '
        COTIMONEM# = 0
        ImVencDol = (-1) * SALDOLA#(U&)
        If MODOMONEI% = 2 Then
           MONEMI% = MONEMISORP%(DOCULARANT$, NCX%) Mod 6
           If Abs(SALDOLA#(U&)) >= 0.005 Then
             COTIMONEM# = ImVenc# / ((-1) * SALDOLA#(U&))
           End If
           If MONEMI% < 2 Then ImVencDol = 0
        End If
        '
        Call GRACREPEN
      End If
    Next U&
    NiveCla% = 0
    '
999 Call CIERRARCH("CUECORR")
    Call LIBARCH("CUECORR")
    Screen.MousePointer = 1
    ATENTI% = 0
    '
End Sub


Function INEFAC#(NROX$, FEMI, NPROX%) ' IMPORTE NETO DE LA FACTURA
    '
    Static NROXA$, FEMIA, NPROXA%, SUNEFA#
    '
    If NROX$ = NROXA$ Then
      If FEMI = FEMIA Then
        If NPROX% = NPROXA% Then
          GoTo 90
        End If
      End If
    End If
    '
    FFII% = FEMI
    FECHAD = FETEXCOR1$(FFII%)
    SUNEFA# = 0
    Call APERBASDAT("CABAN")
    Dim MIRS As ADODB.Recordset
    SQLX$ = "SELECT SUM(IDEBECOMP - IHABECOMP) AS SALDO FROM CAJABANC "
    SQLX$ = SQLX$ + "WHERE TipoMovim = 'COMPRA' "
    SQLX$ = SQLX$ + "AND FechConta = '" & FECHAD & "' "
    SQLX$ = SQLX$ + "AND CodiCom_Lar = '" & NROX$ & "' "
    SQLX$ = SQLX$ + "AND NuProve = " & Abs(NPROX%) & " "
    
    Set MIRS = FLEXCONN.Execute(FILTSQL$(SQLX$))
    If Not IsNull(MIRS!SALDO) Then
       SUNEFA# = MIRS!SALDO
    End If
    '
    NROXA$ = NROX$
    FEMIA = FEMI
    NPROXA% = NPROX%
    '
90  INEFAC# = SUNEFA#
    '
End Function

Function MONEMISOR%(NROX$, FEMI, NCLIEN%) ' MONEDA DE EMISION DE LA FACTURA
    '
    FFII% = FEMI
    FFECHEMI = FETEXCOR1$(FFII%)
    MONEMISOR% = 1
    '
    SQLX$ = "SELECT * FROM CAJABANC "
    SQLX$ = SQLX$ + "WHERE (TipoMovim = 'FVEN' OR TipoMovim = 'RCOB' OR TipoMovim = 'AJUD') "
    SQLX$ = SQLX$ + "AND CodiCom_Lar = '" & NROX$ & "' "
    SQLX$ = SQLX$ + "AND NuClien = " & Abs(NCLIEN%) & " "
    SQLX$ = SQLX$ + "AND FechConta = '" & FFECHEMI & "'"
    '
    Call APERBASDAT("CABAN")
    Dim CUECOPAR As ADODB.Recordset
    Set CUECOPAR = New ADODB.Recordset
25  On Error Resume Next
    CUECOPAR.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockReadOnly, adCmdText  '=  FLEXCONN.EXECUTE (FILTSQL$ (SQLX$))
    If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
      On Error GoTo 0
      Call CapturaErrorOpen
      GoTo 25
   End If
   On Error GoTo 0
   
    With CUECOPAR
       On Error Resume Next
       If Not (.EOF And .BOF) Then
         On Error GoTo 0
         MONEMISOR% = !Mone_Emis
       End If
       On Error GoTo 0
    End With
    CUECOPAR.Close
    Set CUECOPAR = Nothing
    '
    If MONEMISOR% < 1 Then MONEMISOR% = 1
    MONEMISOR% = MONEMISOR% Mod 6   ' ENJUAGUE GRAMY
    '
End Function
Function MONEMISORP%(NROX$, NPROVE%) ' MONEDA DE EMISION DE LA FACTURA
    '
    SQLX$ = "SELECT * FROM CAJABANC WITH(NOLOCK)"
    SQLX$ = SQLX$ + "WHERE (TipoMovim = 'FCOM' OR TipoMovim = 'OPAG' OR TipoMovim = 'AJUP') "
    SQLX$ = SQLX$ + "AND CodiCom_Lar = '" & NROX$ & "' "
    SQLX$ = SQLX$ + "AND NuProve = " & Abs(NPROVE%) & " "
    SQLX$ = SQLX$ + "ORDER BY FechConta DESC"
    Call APERBASDAT("CABAN")
'    Set CUECOPAR = CueCorri.OpenRecordset(SQLX$, dbOpenSnapshot)
    Set CUECOPAR = READSET(SQLX$)
    '
    MONEMISORP% = 1
    On Error Resume Next
    CUECOPAR.MoveFirst
    If Err < 1 Then
      On Error GoTo 0
      MONEMISORP% = XVALO(CUECOPAR!Mone_Emis)
    End If
    On Error GoTo 0
    '
    If MONEMISORP% < 1 Then MONEMISORP% = 1
    MONEMISORP% = MONEMISORP% Mod 6   ' ENJUAGUE GRAMY
    '
End Function
'
Function TICAMISOR(NROX$, FEMI, NCLIEN%) ' TIPO DE CAMBIO DE EMISION DE LA FACTURA
    '
    FFII% = FEMI
    FFECHEMI = FETEXCOR1$(FFII%)
    TICAMISOR = 1
    '
    SQLX$ = "SELECT * FROM CAJABANC "
    SQLX$ = SQLX$ + "WHERE (TipoMovim = 'FVEN' OR TipoMovim = 'RCOB' OR TipoMovim = 'AJUD') "
    SQLX$ = SQLX$ + " AND CodiCom_Lar = '" & NROX$ & "' "
    SQLX$ = SQLX$ + " AND NuClien = " & Abs(NCLIEN%) & " "
    SQLX$ = SQLX$ + " AND FechConta = '" & FFECHEMI & "'"
    '
    Call APERBASDAT("CABAN")
    Dim CUECOPAR As ADODB.Recordset
    Set CUECOPAR = New ADODB.Recordset
25  On Error Resume Next
    CUECOPAR.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockReadOnly, adCmdText  '=  FLEXCONN.EXECUTE (FILTSQL$ (SQLX$))
    If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
      On Error GoTo 0
      Call CapturaErrorOpen
      GoTo 25
   End If
   On Error GoTo 0
    '
    With CUECOPAR
       On Error Resume Next
       If Not (.EOF And .BOF) Then
         On Error GoTo 0
         TICAMISOR = !Tipo_Cam
       End If
       On Error GoTo 0
    End With
    '
    CUECOPAR.Close
    Set CUECOPAR = Nothing
    '
    If TICAMISOR < 0.01 Then TICAMISOR = 1
    '
End Function
'
Function TICAMISORP(NROX$, FEMI, NPROVE%) ' TIPO DE CAMBIO DE EMISION DE LA FACTURA
    '
    FFII% = FEMI
    FECHAD = FETEXCOR1$(FFII%)
    
    SQLX$ = "SELECT * FROM CAJABANC WITH(NOLOCK)"
    SQLX$ = SQLX$ + "WHERE (TipoMovim = 'FCOM' OR TipoMovim = 'OPAG' OR TipoMovim = 'AJUP') "
    SQLX$ = SQLX$ + "AND CodiCom_Lar = '" & NROX$ & "' "
    SQLX$ = SQLX$ + "AND NuProve = " & Abs(NPROVE%) & " "
    SQLX$ = SQLX$ + "AND FechConta = '" & FECHAD & "' "
    
    Set CUECOPAR = READSET(SQLX$)
    '
    TICAMISORP = 1
    FFII% = FEMI
    With CUECOPAR
        If .EOF <> True Then
          On Error Resume Next
          TICAMISORP = XVALO(!Tipo_Cam)
          On Error GoTo 0
        End If
    End With
    '
    If TICAMISORP < 0.01 Then TICAMISORP = 1
    '
End Function
Sub COMPENSACP()

  VDEF$ = Space$(61)
  HOII% = HOY(HO$)
  '
  Call REPLA(VDEF$, MKI$(HOII%), 1, 2)
  Call REPLA(VDEF$, MKI$(0), 3, 2)
  Call REPLA(VDEF$, MKI$(0), 5, 2)
  '
  CUECOMPE$ = Control$("SVD451S", "CUEVE-10")
  If CUECOMPE$ = "" Or CUECOINT%(CUECOMPE$) < 0 Then
    Call MENSERR(24, "Imposible Compensar Cuentas.\  \Falta Definir Cuenta Contable\Correspondiente a Ajustes de Saldos Deudores.")
    Exit Sub
  End If
  '
5 OBJ1$ = OBJPLA$("COMPENSACPN", VDEF$)
  If TRIM(OBJ1$) <> "" Then
    VDEF$ = OBJ1$
    'VALIDACIONES
    F1% = CVI(Mid$(OBJ1$, 1, 2))
    If F1% < 1 Or F1% > HOII% Then
      Call MENSERR(24, "Fecha no Válida")
      GoTo 5
    End If
    NCLI = CVS(Mid$(OBJ1$, 3, 4))
    If NCLI < 1 Or TRIM$(RASOCLI$(NCLI)) = "" Then
      Call MENSERR(24, "Cliente no Válido")
      GoTo 5
    End If
    NP1% = CVI(Mid$(OBJ1$, 7, 2))
    If NP1% < 1 Or TRIM$(RASOCLI$(NP1% * (-1))) = "" Then
      Call MENSERR(24, "Proveedor no Válido")
      GoTo 5
    End If
    IP1# = CVD(Mid$(OBJ1$, 9, 8))
    If IP1# < 0.005 Then
      Call MENSERR(24, "Importe no Válido")
      GoTo 5
    End If
    REF1$ = Mid$(OBJ1$, 17)
    If TRIM$(REF1$) = "" Then
      Call MENSERR(24, "Referencia no Válida")
      GoTo 5
    End If
    'PROCESO NUMERO DE COMPESACION
    FEX = F1%
    FETX$ = FECHATEX$(FEX)
    NRECIB$ = "CC-" + Mid$(FETX$, 7, 2) + Mid$(FETX$, 4, 2) + Mid$(FETX$, 1, 2) + "-"
    '
    'jandy sql
    
    'SQLX$ = "SELECT * FROM CAJABANC "
    Dim RECITEMP As ADODB.Recordset
    SQLX$ = "Select max(substring(CodiCom_Cor,11,2)) as Maximo from CajaBanc"
    SQLX$ = SQLX$ + " WHERE TipoMovim = 'AJUD' "
    'SQLX$ = SQLX$ + " AND CodiCom_Cor LIKE '" & NRECIB$ + "*' "
    SQLX$ = SQLX$ + " AND CodiCom_Cor LIKE '" & NRECIB$ + "%' "
    'SQLX$ = SQLX$ + " ORDER BY CodiCom_Cor DESC"
    Set RECITEMP = FLEXCONN.Execute(FILTSQL$(SQLX$))
    '
    ORMA% = (-1)
    'Set RECITEMP = CueCorri.OpenRecordset(SQLX$, dbOpenSnapshot)
    If Not IsNull(RECITEMP!maximo) Then
        ORMA% = XVALO(RECITEMP!maximo)
    End If
   ' With RECITEMP
   '   On Error Resume Next
   '   .MoveFirst
   '   If Err < 1 Then
   '     ORMA% = XVALO(Mid$(!CodiCom_Cor, 11, 2))
   '   End If
   ' End With
    
    ORMX$ = TRIM$(Str$(ORMA% + 1))
    While Len(ORMX$) < 2: ORMX$ = "0" + ORMX$: Wend
    NRECIB$ = NRECIB$ + Left$(ORMX$, 2)
    
    'GRACABAN PARA LA CTA. CTE. DEL PROVEEDOR

20  NC% = (-1) * NP1%
    Call APERBASDAT("CABAN")
    '
    CoComCor$ = NRECIB$
    CoComLar$ = NRECIB$
'    CodiEmp% = 0
'    CodiSuc% = SUC%
    NumeClie = 0
    NumeProv% = NP1%
    FechEmi = CVDAT(F1%)
    '
    Referen$ = REF1$  'RASOCLI$(NC%)
    FeContab = CVDAT(F1%)
    FeAcredi = CVDAT(F1%)
    NumeTal = 0
    NumeCom = XVALO(Mid$(NRECIB$, 4, 6))
    Observa$ = REF1$ 'RASOCLI$(NC%)
    '
    CuentaCo% = CUECOINT%(CUEMADRE$(NC%))
    TiMovi = "AJUP"            ' COMPENSACION DE CUENTAS
    OpMovi = "CC"              ' COMPENSACION DE CUENTAS
    VaMovi = "CCORR"           ' CUENTA CORRIENTE
    Descrip = "Compensacion - " + RASOCLI$(NCLI)
    ImpoDeb# = IP1#
    Call GRACABAN
    '
    ImpoDeb# = (-1) * IP1#
    CuentaCo% = CUECOINT%(CUECOMPE$)
    TiMovi = "AJUSALPR"        ' AJUSTE DE SALDO PROVEEDORES
    OpMovi = "CC"              ' AJUSTE DE SALDO
    VaMovi = "NO-IMP"          ' NO IMPUTADO
    Call GRACABAN
    '
    '*************GRABA DEUDA EN PROVEEDOR****************
    SIGNO% = 1
    IBruTot# = IP1#
    IIvaTot# = 0
    '
    NuorVen% = 1
    FeVenc = CVDAT(F1%)
    ImVenc# = IP1# * (-1)
    Call GRACREPEN
    '
    'Graba la Contra-Partida en CLIENTE.-
30  Call APERBASDAT("CABAN")
    '
    CoComCor$ = NRECIB$
    CoComLar$ = NRECIB$
'    CodiEmp% = 0
'    CodiSuc% = SUC%
    NumeClie = NCLI
    NumeProv% = 0
    FechEmi = CVDAT(F1%)
    '
    Referen$ = REF1$  'RASOCLI$(NC1%)
    FeContab = CVDAT(F1%)
    FeAcredi = CVDAT(F1%)
    NumeTal = 0
    NumeCom = XVALO(Mid$(NRECIB$, 4, 6))
    Observa$ = REF1$ 'RASOCLI$(NC1%)
    '
    CuentaCo% = CUECOINT%(CUEMADRE$(NCLI))
    TiMovi = "AJUD"            ' COMPENSACION DE CUENTAS
    OpMovi = "CC"              ' COMPENSACION DE CUENTAS
    VaMovi = "CCORR"           ' CUENTA CORRIENTE
    Descrip = "Compensacion - " + RASOCLI$(NC%)
    ImpoDeb# = IP1# * (-1)
    Call GRACABAN
    '
    ImpoDeb# = IP1#
    CuentaCo% = CUECOINT%(CUECOMPE$)
    TiMovi = "AJUSALD"         ' AJUSTE DE SALDO DEUDORES
    OpMovi = "CC"              ' AJUSTE DE SALDO
    VaMovi = "NO-IMP"          ' NO IMPUTADO
    Call GRACABAN
    '
    '*************GRABA DEUDA EN CLIENTE****************
    SIGNO% = 1
    IBruTot# = IP1#
    IIvaTot# = 0
    '
    NuorVen% = 1
    FeVenc = CVDAT(F1%)
    ImVenc# = IP1# * (-1)
    Call GRADEUPEN
    '
    FORIPRE$ = TRIM$(Control$("", "FORCOMPE"))
    If FORIPRE$ = "" Then
        Call COMUNI("Error de Configuración:\  \Falta Definir Formulario\de Impresión de Compensación de Cuentas.")
        Exit Sub
    End If
    '
    If FORIPRE$ <> "" Then
        LU$ = LUDAT$
        If EXISTE%(LU$ + FORIPRE$ + ".FRM") <> 1 Then
            RMCC$ = REPLACAR$(LU$ + FORIPRE$ + ".FRM", "\", "/")
            Call COMUNI("Error de Configuración:\  \No Existe Formulario de Impresión\de Compensación de Cuentas:\  \'" + RMCC$ + "'.")
            Exit Sub
        End If
    End If
    '
    FECHDOC$ = FECHATEX$(F1%)
    NUMEDOC$ = TRIM$(Str$(NURETGA&))
    NUMEDOC$ = NRECIB$
    TIPODOC$ = "Compensación de Cuentas"
    '
    CODPARAM$(1) = "REF-MAT1"
    CODPARAM$(2) = "IMP-TOTA"
    CAPARDOC% = 2
    '
    DOCPARAM$(1) = TRIM$(REF1$)
    DOCPARAM$(2) = TRIM$(CSTRING$(MKD$(IP1#), 3, 10, 2, 2))
    '
    CODTABU1$(1) = "CUEN-CLI"
    CODTABU1$(2) = "RASO-CLI"
    CACOLTAB1% = 2
    '
    CAITAB1% = 0
    CAITAB1% = CAITAB1% + 1
    TETABU1$(1, CAITAB1%) = TRIM$(Str$(NCLI)) 'CLIENTE
    TETABU1$(2, CAITAB1%) = TRIM$(RASOCLI$(NCLI)) 'CLIENTE
    CAITAB1% = CAITAB1% + 1
    TETABU1$(1, CAITAB1%) = TRIM$(Str$(NP1%)) 'PROVEEDOR
    TETABU1$(2, CAITAB1%) = TRIM$(RASOCLI$(-1 * (NP1%))) 'PROVEEDOR
    '
    Call PRINTDOC("FORCOMPE")
    '\\\OT-06-0144-WAR-FIN///
    '
    Call COMUNI("Compensación Procesada.")
    'CajaBanc.Close
    '
  End If
'\\\OT-06-0112-WAR-FIN///
End Sub

Function SALDEUPEN#(NUFACX$)
   '
   ' DEVUELVE EL SALDO DEUDOR PENDIENTE DEL COMPROBANTE LARGO O CORTO NUFACX$
   '
   Call ABRECONEXION
   SDP# = 0
   NNX$ = TRIM$(NUFACX$)
   If NNX$ <> "" Then
      Call ABRECABAN
      SQLX$ = "SELECT SUM(SaldDebe - SALDAcre) AS SalDeuP "
      SQLX$ = SQLX$ + " FROM SADEUPEN "
      SQLX$ = SQLX$ + "WHERE CodiCom_Lar = '" & NNX$ & "' "
      SQLX$ = SQLX$ + " OR CodiCom_Cor = '" & NNX$ & "'"
      Set SADEUPE = FLEXCONN.Execute(FILTSQL$(SQLX$))
      '
      With SADEUPE
         If Not IsNull(!SalDeuP) Then
            SDP# = (!SalDeuP)
          End If
      End With
      SADEUPE.Close
      Set SADEUPE = Nothing
   End If
   '
20 SALDEUPEN# = SDP#
   '
End Function

Sub RECUSALDOL(DOCULA$, NCX%)
    '
    SALDOLA# = 0
    RFF$ = RECFILT$("CARBIMON", 2, MKI$(NCX%))
    FIN& = Len(RFF$)
    For KKU& = 1 To FIN& Step 4
       Call TRACE(20, KKU&, FIN&)
       U& = CVS(Mid$(RFF$, KKU&, 4))
       XX$ = REGLEIDO$("CARBIMON", U&)
       NUDOCX$ = TRIM$(Mid$(XX$, 13, 18))
       If NUDOCX$ = DOCULA$ Then
          DEBEDOL# = ROUND#(CVD(Mid$(XX$, 97, 8)))
          HABEDOL# = ROUND#(CVD(Mid$(XX$, 105, 8)))
          SALDOLA# = SALDOLA# + DEBEDOL# - HABEDOL#
       End If
191 Next KKU&
    '
    Call APERBASDAT("CABAN")
    strsql = "Select * from SaDeuPen where NuClien=" & NCX%
    strsql = strsql & " And CODICOM_LAR = '" & DOCULA$ & "'"
    
    Dim SaDeuPenAcre As ADODB.Recordset
    Set SaDeuPenAcre = New ADODB.Recordset
25  On Error Resume Next
    SaDeuPenAcre.Open FILTSQL$(strsql), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
    If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
      On Error GoTo 0
      Call CapturaErrorOpen
      GoTo 25
   End If
   On Error GoTo 0
    
    With SaDeuPenAcre
      If Not (.EOF And .BOF) Then
        If SALDOLA# > 0 Then
             !SaldDebe_DOLA = SALDOLA#
             !SaldAcre_DOLA = 0
           Else
             !SaldDebe_DOLA = 0
             !SaldAcre_DOLA = Abs(SALDOLA#)
        End If
        .Update
        SALDOLA# = 0
      End If
    End With
    '
    SaDeuPenAcre.Close
    Set SaDeuPenAcre = Nothing
End Sub
'
Function VERICONCILIA%(COCOLA$)
   '
   ' Se invoca antes de anular un comprobante y devuelve:
   ' 0 (cero) si alguno de los registros que se van a eliminar tiene marca de conciliacion bancaria
   ' 1 (uno) si ninguno está conciliado --> se puede anular.
   '
   VCC% = 0
   SQLX$ = " SELECT *  FRoM CAJABANC WITH(NOLOCK)"
   SQLX$ = SQLX$ + " WHERE CODICOM_LAR = '" & COCOLA$ & "'"
   SQLX$ = SQLX$ + " AND MarConsi <> '' AND MarConsi IS NOT NULL "
  
   Dim DEPTMP As ADODB.Recordset
   Set DEPTMP = New ADODB.Recordset
   Set DEPTMP = FLEXCONN.Execute(FILTSQL$(SQLX$))
   '
   If DEPTMP.BOF And DEPTMP.EOF Then
      VCC% = 1
   End If
   '
   If VCC% < 1 Then
      Call MENSERR(24, "Anulación Imposible !!!\Movimiento Bancario Conciliado.")
   End If
   VERICONCILIA% = VCC%
End Function
'
Function SUMAVALFAC#(FACTU$)
    '\\\OT-08-0801-OD-04/09/08///
    TOTALFAC# = 0
    '
    'CREO LA CONSULTA EN EL MOVISTO DONDE EL CAMPO DOSECLAR ES IGUAL A FACTU$
    'FUNCIONALIDAD PASANDOLE COMO ARGUMENTO EL NUMERO DE FACTURA LARGO _
    DEVUELVE EL IMPORTE DE LA MISMA
    '
    If FACTU$ <> "" Then
        SQLX$ = "SELECT CANTSALID ,VALOUNIT ,TIPOCAM FROM MOVISTO"
        SQLX$ = SQLX$ + " WHERE CODIMOVI = 'RT' AND DOSECLAR = '" & FACTU$ & "'"
        Dim MOVITMP As ADODB.Recordset
        Set MOVITMP = FLEXCONN.Execute(FILTSQL$(SQLX$))
        '
        'RECORRO EL RECORDSET Y ACUMULO EN TOTALFAC
        With MOVITMP
            Do While Not .EOF
               TipKamb# = XVALO(!TipoCam)
               If TipKamb# < 0.005 Then TipKamb# = 1
               TOTALFAC# = TOTALFAC# + ROUND(((XVALO(!CANTSALID) * XVALO(!VALOUNIT)) * XVALO(TipKamb#)))
               .MoveNext
            Loop
        End With
        '
        MOVITMP.Close
        Set MOVITMP = Nothing
        '
    End If
    SUMAVALFAC# = TOTALFAC#
    '
    '\\\OT-08-0801-OD-FIN///

End Function

Function YAESTACHEQUE%(BANCO$, CHEQUE$)
    '
    'VALIDA QUE EL CHEQUE NO ESTE INGRESADO
    'DEL MISMO BANCO - NUMERO
    '0 NO INGRESADO - 1 INGRESADO
    '
    YAESTACHEQUE% = 0
    If TRIM$(CHEQUE$) = "" Then Exit Function
    '
    ABUS$ = AJUSTI$(BANCO$, 18)
    ABUS$ = ABUS$ + AJUSTI$(CHEQUE$, 14)
    '
    Call APERBASDAT("CABAN")
    SQLX$ = "TIPOMOVIM = 'CHRECI' "
    SQLX$ = SQLX$ + " AND DESCRIMOV LIKE '" & ABUS$ + "%'"
    '
    QTR& = CantRegistros("CAJABANC", SQLX$)
    If QTR& > 0 Then YAESTACHEQUE% = 1
    '
End Function
Function YAESTACHEQUE_OP%(BANCO$, CHEQUE$, ORDPAGO$)
    '
    'DEVUELVE 0 SI EL CHEQUE NO FIGURA EN B.DATOS,
    'DEVUELVE 1 SI EL CHEQUE FIGURA EN B.DATOS.
    'EN CASO QUE EL CHEQUE FIGURE EN B.D. DEVUELVE EN ORDPAGO$ EL NUMERO DE ORDEN DE PAGO
    
    YAESTACHEQUE_OP% = 0
    ORDPAGO$ = ""
    If TRIM$(CHEQUE$) = "" Then Exit Function
    '
    ABUS$ = AJUSTI$(BANCO$, 18)
    ABUS$ = ABUS$ + AJUSTI$(CHEQUE$, 14)
    '
    Call APERBASDAT("CABAN")
    SQLX$ = "TIPOMOVIM = 'CHENTR'"
    SQLX$ = SQLX$ + " AND DESCRIMOV LIKE '" & ABUS$ + "%'"
    '
    QTR& = CantRegistros("CAJABANC", SQLX$)
    If QTR& > 0 Then
      YAESTACHEQUE_OP% = 1
      ORDPAGO$ = VALOBADA("CAJABANC", "CODICOM_LAR", SQLX$, "NOMESS")
    End If
    '
End Function

Function NUOCOCLI$(NPED&, CIXI%)
    '
    NUOCOCLI$ = ""
    '
    Call APERBASDAT("PEDCLI")
    SQLX$ = " SELECT * FROM PEDDETA "
    SQLX$ = SQLX$ + " WHERE NroPed = " & NPED& & " "
    SQLX$ = SQLX$ + " AND Status < 8 "    ' NO ANULADOS NI CUMPLIDOS
    SQLX$ = SQLX$ + " AND STATUS >= 0 "   ' APROBADOS
    SQLX$ = SQLX$ + " AND CodiInt = " & CIXI% & " "
    '
    'Set PEDTEMP = Ventas.OpenRecordset(sqlx$, dbOpenSnapshot)
    '
    Dim PEDTEMP As ADODB.Recordset
    Set PEDTEMP = FLEXCONN.Execute(FILTSQL$(SQLX$))

    With PEDTEMP
      On Error Resume Next
      .MoveFirst
      If Err < 1 Then
        NUOCOCLI$ = !NROOCOM
      End If
    End With
    PEDTEMP.Close
    Set PEDTEMP = Nothing
    '
End Function
'
Function NOCOMCLI$(NPED&)
    '
    NOCOMCLI$ = ""
    '
    Call APERBASDAT("PEDCLI")
    SQLX$ = " SELECT * FROM PEDENCA "
    SQLX$ = SQLX$ + " WHERE NroPed = " & NPED& & " "
    'Set PEDTEMP = Ventas.OpenRecordset(SQLX$, dbOpenSnapshot)
    Dim PEDTEMP As ADODB.Recordset
    Set PEDTEMP = FLEXCONN.Execute(FILTSQL$(SQLX$))
    '
    With PEDTEMP
      On Error Resume Next
      .MoveFirst
      If Err < 1 Then
        On Error GoTo 0
        NOCOMCLI$ = !NROOCOM
      End If
      On Error GoTo 0
    End With
    PEDTEMP.Close
    Set PEDTEMP = Nothing
    '
End Function
'
Function DESCUPEDI$(NPED&)     ' DESCUENTO GENERAL DEL PEDIDO
    '
    DESCUPEDI$ = ""
    '
    Call APERBASDAT("PEDCLI")
    SQLX$ = " SELECT TDescApli, Dto_Ret_Cte FROM PEDENCA "
    SQLX$ = SQLX$ + " WHERE NroPed = " & NPED& & " "
    Dim PEDTEMP As ADODB.Recordset
    Set PEDTEMP = FLEXCONN.Execute(FILTSQL$(SQLX$))
    '
    With PEDTEMP
       If Not (.BOF And .EOF) Then
          DESCUPEDI$ = TRIM$(SAFETEXT$(!TDescApli))
          DESC_RETI_CLTE$ = TRIM$(SAFETEXT$(!Dto_Ret_Cte))
       End If
    End With
    PEDTEMP.Close
    Set PEDTEMP = Nothing
    '
    'Agregado  PARA QUE TRAIGA EL DESCUENTO DE RETIRA CLIENTE
    If DESC_RETI_CLTE$ <> "" Then
        SIGNO$ = ""
        If XVALO(DESC_RETI_CLTE$) > 0 And DESCUPEDI$ <> "" Then SIGNO$ = "+"
        DESCUPEDI$ = DESC_RETI_CLTE$ & SIGNO$ & DESCUPEDI$
    End If
    '
End Function
'
Function NUMOTIVO%(NPED&)
   '
   'DEVUELVE EL NUMERO DE MOTIVO ASIGNADO AL PEDIDO
   NUMTV% = 0
   If NPED& < 1 Then GoTo 99
   '
   CONDI$ = "NROPED =" & NPED&
   NUMTV% = XVALO(VALOBADA("PEDENCA", "OV_MOTIVO", CONDI$, "NOMESS"))
99 NUMOTIVO% = NUMTV%
   '
End Function
'
Function MONEPEDI%(NPED&)     ' MONEDA DE EMISION GENERAL DEL PEDIDO
    '
    MONEPEDI% = 0
    '
    Call APERBASDAT("PEDCLI")
    SQLX$ = " SELECT MoneEmis FROM PEDDETA "
    SQLX$ = SQLX$ + " WHERE NroPed = " & NPED& & " "
    Dim PEDTEMP As ADODB.Recordset
    Set PEDTEMP = FLEXCONN.Execute(FILTSQL$(SQLX$))
    '
    With PEDTEMP
       If Not (.BOF And .EOF) Then
          MONEPEDI% = XVALO(!MoneEmis)
       End If
    End With
    PEDTEMP.Close
    Set PEDTEMP = Nothing
    '
    If MONEPEDI% < 1 Then MONEPEDI% = 1
    '
End Function
'
Function PRECIOPED#(NPED&, CODIMAT%, MONEPE%)
    '
    PRECIOPED# = 0
    MONEPE% = 1
    '
    If NPED& > 0 Then
      Call APERBASDAT("PEDCLI")
      SQLX$ = " SELECT * FROM PEDDETA "
      SQLX$ = SQLX$ + " WHERE NroPed = " & NPED& & " "
      SQLX$ = SQLX$ + " AND CodiInt = " & CODIMAT% & " "
      SQLX$ = SQLX$ + " AND Status < 9 "    ' NO ANULADOS
      '
      ' JANDY SQL
      
      'Set PEDTEMP = Ventas.OpenRecordset(SQLX$, dbOpenSnapshot)
      '
      Dim PEDTEMP As ADODB.Recordset
      Set PEDTEMP = FLEXCONN.Execute(FILTSQL$(SQLX$))
      With PEDTEMP
        On Error Resume Next
        .MoveFirst
        If Err < 1 Then
          PRECIOPED# = !PreUnid
          MONEPE% = !MoneEmis
        End If
        On Error GoTo 0
      End With
    End If
    '
End Function
'
Function MONEMIPED%(NPED&)   ' MONEDA DE EMISION DEL PEDIDO
    '
    MONEMIPED% = 1    ' POR DEFAULT PESOS
    '
    If NPED& > 0 Then
      Call APERBASDAT("PEDCLI")
      SQLX$ = " SELECT * FROM PEDDETA "
      SQLX$ = SQLX$ + " WHERE NroPed = " & NPED& & " "
      SQLX$ = SQLX$ + " AND Status < 9 "    ' NO ANULADOS
      'Set PEDTEMP = Ventas.OpenRecordset(SQLX$, dbOpenSnapshot)
      Dim PEDTEMP As ADODB.Recordset
      Set PEDTEMP = FLEXCONN.Execute(FILTSQL$(SQLX$))
      
      '
      With PEDTEMP
        On Error Resume Next
        .MoveFirst
        If Err < 1 Then
          MONEMIPED% = !MoneEmis
        End If
        On Error GoTo 0
      End With
      Set PEDTEMP = Nothing
    End If
    '
End Function
'
Function SUBCODI$(NPED&, CIXI%)
    '
    SUBCODI$ = ""
    '
    Call APERBASDAT("PEDCLI")
    SQLX$ = " SELECT * FROM PEDDETA "
    SQLX$ = SQLX$ + " WHERE NroPed = " & NPED& & " "
    SQLX$ = SQLX$ + " AND Status < 8 "    ' NO ANULADOS NI CUMPLIDOS
    SQLX$ = SQLX$ + " AND STATUS >= 0 "   ' APROBADOS
    SQLX$ = SQLX$ + " AND CodiInt = " & CIXI% & " "
    '
    Dim PEDTEMP As ADODB.Recordset
    Set PEDTEMP = FLEXCONN.Execute(FILTSQL$(SQLX$))
    'Set PEDTEMP = Ventas.OpenRecordset(SQLX$, dbOpenSnapshot)
    '
    
    With PEDTEMP
      On Error Resume Next
      .MoveFirst
      If Err < 1 Then
        If IsNull(!SubCod) = False Then
          SUBCODI$ = !SubCod
        End If
      End If
    End With
    PEDTEMP.Close
    Set PEDTEMP = Nothing
    '
End Function
'
Function VALIDAOV%(NPED&)

   '0 = NO AUTORIZADO
   '1 = AUTORIZADO
   '
   Call ABREPEDCLI
   '
   'VALIDACION DE PEDIDOS CON MOTIVO
   Vali% = 1
   '
   'VALIDACION DE TIPO DE MOTIVO (SI NO TIENE MOTIVO MARCA 1 (APROBADO))
   NUMOT% = NUMOTIVO%(NPED&)
   If NUMOT% < 1 Then Vali% = 1: GoTo 99
   '
   MENSA$ = ""
   '
   CONDI$ = "NROPED = " & NPED&
   DESTINO$ = TRIM$(VALOBADA("PEDENCA", "OV_Destino", CONDI$, "NOMESS"))
   VEHICULO$ = TRIM$(VALOBADA("PEDENCA", "OV_Vehiculo", CONDI$, "NOMESS"))
   CONDIDESPA$ = TRIM$(VALOBADA("PEDENCA", "OV_CondDespa", CONDI$, "NOMESS"))
   FESALI = CVINT(VALOBADA("PEDENCA", "OV_Fesalida", CONDI$, "NOMESS"))
   FEREINGRE = CVINT(VALOBADA("PEDENCA", "OV_FeRegreso", CONDI$, "NOMESS"))
   If FESALI > FEREINGRE Then
      MENSA$ = "Fecha De Ingreso Mayor a Fecha de Salida"
      Vali% = 0: GoTo 95
   End If
   '
   Observa$ = VALOBADA("PEDENCA", "OV_Observaciones", CONDI$, "NOMESS")
   NUSERIE$ = VALOBADA("PEDENCA", "OV_NSerieDev", CONDI$, "NOMESS")
   '
   'VALIDO  SI SE INGRESO UNA FECHA DE REGRESO QUE NO SEA ANTERIOR A LA DE HOY
   FEX = HOY(HOS$)
   If FEREINGRE > 0 Then
    If FEREINGRE < FEX Then
       MENSA$ = " Fecha de Regreso Anterior a la de Hoy"
       Vali% = 0: GoTo 95
    End If
   End If
   '
   'VALIDACION QUE TENGA VENDEDOR
   If NUMOT% <> 5 Then ' 5 ES DONACIONES Y DEBERIA PERMITIR QUE NO TENGA VENDEDOR ASOCIADO
    CONDI$ = "NROPED = " & NPED&
    VENDE% = XVALO(VALOBADA("PEDENCA", "VENDED", CONDI$, "NOMESS"))
    If VENDE% < 1 Then
       MENSA$ = "Pedido Con Motivo Debe Ingresar un Vendedor "
       Vali% = 0: GoTo 95
    End If
   End If
   '
   
    'VALIDACIONES DE ACUERDO A TIPO DE MOTIVO
    
    Select Case NUMOT%
     Case 1  ' GIRAS / CONGRESOS
       '
       If VEHICULO$ = "" Then MENSA$ = "Falta Ingresar Vehículo": Vali% = 0: GoTo 95
        
       If DESTINO$ = "" Then MENSA$ = "Falta Ingresar Destino": Vali% = 0: GoTo 95
       If FESALI < 1 Then MENSA$ = "Falta Ingresar Fecha de Salida": Vali% = 0: GoTo 95
       If FEREINGRE < 1 Then MENSA$ = "Falta Ingresar Fecha de Regreso": Vali% = 0: GoTo 95
       If CONDIDESPA$ = "" Then MENSA$ = "Falta Ingresar Condición de Despacho": Vali% = 0: GoTo 95
       '
     Case 2 'DEMOSTRACION
       If VEHICULO$ = "" Then MENSA$ = "Falta Ingresar Vehículo": Vali% = 0: GoTo 95
       If FESALI < 1 Then MENSA$ = "Falta Ingresar Fecha de Salida": Vali% = 0: GoTo 95
       If FEREINGRE < 1 Then MENSA$ = "Falta Ingresar Fecha de Regreso": Vali% = 0: GoTo 95
       '
    Case 3  ' PRESTAMO
       If FESALI < 1 Then MENSA$ = "Falta Ingresar Fecha de Salida": Vali% = 0: GoTo 95
       If FEREINGRE < 1 Then MENSA$ = "Falta Ingresar Fecha de Regreso": Vali% = 0: GoTo 95
       If Observa$ = "" Then MENSA$ = "Falta Ingresar Observación/Motivo de Prestamo": Vali% = 0: GoTo 95
       '
    Case 4  ' CAMBIO
       If FESALI < 1 Then MENSA$ = "Falta Ingresar Fecha de Salida": Vali% = 0: GoTo 95
       If FEREINGRE < 1 Then MENSA$ = "Falta Ingresar Fecha de Regreso": Vali% = 0: GoTo 95
       If Observa$ = "" Then MENSA$ = "Falta Ingresar Observación/Motivo de Cambio": Vali% = 0: GoTo 95
       'VALIDAR EN EL CASO QUE EL MOTIVO SEA CAMBIO QUE NO SE HAYA CARGADO MAS DE 1 ITEM
       DESCMOTI$ = ECOARCH$("TAMOTPRE", Chr$(NUMOT%))
       
       'VALIDO QUE NO TENGA MAS DE 1 REGISTRO
       CONDI$ = "NROPED = " & NPED&
       CANTREPED& = CantRegistros&("PEDDETA", CONDI$)
       If CANTREPED& > 1 Then
           MENSA$ = (" Pedido con Motivo: '" & TRIM$(DESCMOTI$) & "' Con más de 1 Item")
          Vali% = 0
          GoTo 95
       Else
          ' SI TIENE UN SOLO REGISTRO VALIDO QUE CANTIDAD NO SEA > 1
          CAN = XVALO(VALOBADA("PEDDETA", "CANPED", CONDI$, "NOMESS"))
          If CAN > 1 Then
             MENSA$ = ("Pedido con Motivo: '" & TRIM$(DESCMOTI$) & "' Por más de 1 Equipo")
            Vali% = 0
            GoTo 95
          End If
       End If
       '
       'VALIDACION DE NUMERO DE SERIE
       If NUSERIE$ = "" Then MENSA$ = "Falta Ingresar Numero de Serie": Vali% = 0: GoTo 95
       CANTREG& = CantRegistros("TANUMSE", "NUMEROSERIE = '" & NUSERIE$ & "'")
       If CANTREG& < 1 Then
           MENSA$ = ("Número de  Serie No Válido"): Vali% = 0: GoTo 95
       End If
       CONDI1$ = "NUMEROSERIE = '" & NUSERIE$ & "'AND (NUMFACTURA = ''  OR NUMFACTURA IS NULL OR  NUMREMITO <> '' OR NUMREMITO IS NULL) "
       CANTREG& = CantRegistros("TANUMSE", CONDI1$)
       If CANTREG& < 1 Then
           MENSA$ = "Número de Serie Ingresado Para Cambio\Registrado como Facturado y/o Remitido": Vali% = 0: GoTo 95
       End If
       '
    Case 5 ' Donaciones
       If FESALI < 1 Then MENSA$ = "Falta Ingresar Fecha de Salida": GoTo 95
       If DESTINO$ = "" Then MENSA$ = "Falta Ingresar Destino": GoTo 95
       '
    End Select

    CONDI$ = "NROPED = " & NPED&
    AUTORI1$ = VALOBADA("PEDENCA", "OV_Autoriza1", CONDI$, "NOMESS")
    If AUTORI1$ = "" Then
       MENSA$ = "Falta Autorización de Primer Autorizante"
       Vali% = 0: GoTo 95
    End If
    '
    AUTORI2$ = VALOBADA("PEDENCA", "OV_Autoriza2", CONDI$, "NOMESS")
    If AUTORI2$ = "" Then
       MENSA$ = "Falta Autorización de  Supervisor"
       Vali% = 0: GoTo 95
    End If
    '
    If MENSA$ <> "" Then
95     Call COMUNI("Pedido Nro.: " & NPED& & " Motivo: " & NUMOT% & "\" & MENSA$)
    End If
    '
99  VALIDAOV% = Vali%
    '
End Function


Function REFECONCI_AUTOMA$(CODICONCI$)
    REFECONCI_AUTOMA$ = VALOBADA("CONCILIA", "REFECONCI", "CODICONCI = '" & CODICONCI$ & "'")
End Function
Function CUENTA_BCO_CONCILIA%(CODICONCI$)
    CUENTA_BCO_CONCILIA% = XVALO(VALOBADA("CONCILIA", "CUENTACONTBANCO", "CODICONCI = '" & CODICONCI$ & "'"))
End Function

