Attribute VB_Name = "CUEC_FUN"
' los siguentes son generales en varias de las rutinas
Public CoComCor$
Public CoComLar$
Public NumeClie%
Public CliReal%
Public SucuClie%
Public NumeProv%
Public FechEmi
'
' los siguientes para CajaBanc
Public CondiPago%
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
Public SuBalan As Double
Public AtraCheque&
Public AtraCance&
Public IdentiCae$ 'AGREGADO PARA FACTURA ELECTRONICA
Public NumeroCae$ 'AGREGADO PARA FACTURA ELECTRONICA
Public ObservaFactu$ ' AGREAGADO OBSERVACION DE FACTURA
Public ValorFob$ ' AGREGADO PARA EL VALOR FOB, CIF, ETC
Public DespaImpo$
Public CuitComprador$
'
' los siguientes para grabar vencimientos e importes
Public IBruTot#
Public IIvaTot#
Public NuorVen%
Public FeVenc
Public ImVenc#
Public ImVencDol#
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
'
Public Descue1$ 'DESCUENTO SE GUARDA EN CAJABANC

Public COTIMONEM#

Sub GRACABAN()
   '
   If NOCLOSE% = 0 Then Call APERBASDAT("CABAN")
   With CajaBanc
     .AddNew
      !CodiCom_Cor = Left$(CoComCor$, 12)
      !CODICOM_LAR = Left$(CoComLar$, 24)
      !Refecom = Left$(Referen$, 48)
      !FechConta = FeContab
      !FechAcred = FeAcredi
      !CueContaI = CuentaCo%
      !CENCOSTO = Left$(CentroCos$, 8)
      !ItemConc = Left$(ItConcep$, 8) 'nuevo p/ concepto
      !NuClien = NumeClie%
      !Cli_Real = CliReal%
      !Sucu_Clie = SucuClie%
      !NuProve = NumeProv%
      !TipoMovim = Left$(TiMovi$, 8)
      !OPCIMOVIM = Left$(OpMovi$, 8)
      !VARIMOVIM = Left$(VaMovi$, 8)
      !Descrimov = Left$(Descrip$, 48)
      !FechEmisi = FechEmi
      !CodiEmpr = CodiEmp%
      !CodiSucu = CodiSuc%
      !NUMETALO = NumeTal
      !NumeComp = NumeCom
      !Status = StatInit
        On Error Resume Next
      !Mone_Emis = MONEMI%
      !Tipo_Cam = TICAMONE(MONEMI%)
        If COTIMONEM# >= 0.1 Then !Tipo_Cam = COTIMONEM#
        On Error GoTo 0
        IDEBEX# = ROUND(ImpoDeb#)
      !ValAbsComp = Abs(IDEBEX#)
      If IDEBEX# > 0 Then
          !IDEBECOMP = IDEBEX#
          !IHABECOMP = 0
        Else
          !IDEBECOMP = 0
          !IHABECOMP = Abs(IDEBEX#)
      End If
      SuBalan# = SuBalan# + IDEBEX#
      !Impo_Dola = Abs(ValoDola#)
      !BASEIMPO = Abs(BasImpo#)
      !NROCAI = NumCai$
        On Error Resume Next
      !IncSubIva = InclIva$
        On Error GoTo 0
      !FEVTOCAI = FechCai
      !OBSERGEN = Left$(Observa$, 64)
      !IdenAlta = Left$(IDACTU$, 8)          ' FECHA - HORA - USUARIO - TERMINAL
      !IdenBaja = String$(6, 0)
      !FeHoReal = Now
      !NuAsiDiar = ""
      !AtraCance = AtraCance&
      !AtraCheque = AtraCheque&
      !CONDPAG = CondiPago%
      'AGREGADO PARA FACTURA ELECTRONICA
      !CAE = TRIM$(Left$(NumeroCae$, 32))
      !IDCAE = TRIM$(Left$(IdentiCae$, 32))
      !Notas = ObservaFactu$
      !Incoterms = ValorFob$
      'HASTA AQUI AGREGADO PARA FACTURA ELECTRONICA
      !DespaImpo = DespaImpo$
      !CUIT = CuitComprador$
      
      'DEscuentos general y descuento especial (Gabal) para presentar en F.Electronica.
      !Descu1 = Descue1$
      
     .Update
   End With
   If NOCLOSE% = 0 Then CajaBanc.Close
   '
End Sub
'
Sub GRACREPEN()
   '
   If NOCLOSE% = 0 Then Call APERBASDAT("CABAN")
   With SacrePen
     .AddNew
      !CodiEmpr = CodiEmp%
      !CodiSucu = CodiSuc%
      !NuClien = NumeClie%
      !NuProve = NumeProv%
      !CodCliPro = Left$(CODICLI$(NumeClie% - NumeProv%), 16)
      !RaSoCliPro = Left$(RASOCLI$(NumeClie% - NumeProv%), 48)
      !CodiCom_Cor = Left$(CoComCor$, 12)
      !CODICOM_LAR = Left$(CoComLar$, 24)
      !FechEmisi = FechEmi
        On Error Resume Next
      !Mone_Emis = MONEMI%
      If MONEMI% < 0 Then !Mone_Emis = 0
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
      !FEVENCIM = FeVenc
      If ImVenc > 0 Then
          !SaldAcre = ImVenc#
          !SaldDebe = 0
          !SaldAcre_Dola = ROUND#(ImVenc# / TICAMO#)
          !SaldDebe_Dola = 0
        Else
          !SaldDebe = Abs(ImVenc#)
          !SaldAcre = 0
          !SaldDebe_Dola = ROUND#(Abs(ImVenc# / TICAMO#))
          !SaldAcre_Dola = 0
      End If
      If MONEMI% = (-1) Then
          !SaldAcre_Dola = 0
          !SaldDebe_Dola = 0
      End If
      !Status = 0
     .Update
   End With
   If NOCLOSE% = 0 Then SacrePen.Close
   '
End Sub

Sub GRACREPEN1()
   '
   If NOCLOSE% = 0 Then Call APERBASDAT("CABAN")
   With SacrePen
     .AddNew
      !CodiEmpr = CodiEmp%
      !CodiSucu = CodiSuc%
      !NuClien = NumeClie%
      !NuProve = NumeProv%
      !CodCliPro = Left$(CODICLI$(NumeClie% - NumeProv%), 16)
      !RaSoCliPro = Left$(RASOCLI$(NumeClie% - NumeProv%), 48)
      !CodiCom_Cor = Left$(CoComCor$, 12)
      !CODICOM_LAR = Left$(CoComLar$, 24)
      !FechEmisi = FechEmi
        On Error Resume Next
      !Mone_Emis = MONEMI%
      If MONEMI% < 0 Then !Mone_Emis = 0
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
      !FEVENCIM = FeVenc
      If ImVenc > 0 Then
           !SaldAcre = ImVenc#
           !SaldDebe = 0
         Else
           !SaldDebe = Abs(ImVenc#)
           !SaldAcre = 0
      End If
      If ImVencDol# > 0 Then
           !SaldAcre_Dola = ImVencDol#
           !SaldDebe_Dola = 0
         Else
           !SaldDebe_Dola = ROUND#(Abs(ImVencDol#))
           !SaldAcre_Dola = 0
      End If
      !Status = 0
     .Update
   End With
   If NOCLOSE% = 0 Then SacrePen.Close
   '
End Sub

Sub GRADEUPEN()
   '
   If NOCLOSE% = 0 Then Call APERBASDAT("CABAN")
   
   With SaDeuPen
     .AddNew
      !CodiEmpr = CodiEmp%
      !CodiSucu = CodiSuc%
      !NuClien = NumeClie%
      !Sucu_Clie = SucuClie%
      !NuProve = NumeProv%
      !CodCliPro = Left$(CODICLI$(NumeClie% - NumeProv%), 16)
      !RaSoCliPro = Left$(RASOCLI$(NumeClie% - NumeProv%), 48)
      !CodiCom_Cor = Left$(CoComCor$, 12)
      !CODICOM_LAR = Left$(CoComLar$, 24)
      !FechEmisi = FechEmi
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
      !FEVENCIM = FeVenc
      If ImVenc > 0 Then
          !SaldDebe = ImVenc#
          !SaldAcre = 0
          !SaldDebe_Dola = ROUND#(ImVenc# / TICAMO#)
          !SaldAcre_Dola = 0
        Else
          !SaldAcre = Abs(ImVenc#)
          !SaldDebe = 0
          !SaldAcre_Dola = Abs(ImVenc# / TICAMO#)
          !SaldDebe_Dola = 0
      End If
      If Left$(CoComLar$, 2) = "RB" Or MONEMI% < 2 Then
         !SaldDebe_Dola = 0
         !SaldAcre_Dola = 0
      End If
      !Status = 0
      !CONDPAG = CondiPago%
     .Update
   End With
   If NOCLOSE% = 0 Then SaDeuPen.Close
   '
End Sub

Sub GRAPLIPAG()
   If NOCLOSE% = 0 Then Call APERBASDAT("CABAN")
   With ApliPago
     .AddNew
      !CodiEmpr = CodiEmp%
      !CodiSucu = CodiSuc%
      !NuProve = NumeProv%
      !CodiProv = Left$(CODICLI$((-1) * NumeProv%), 16)
      !RasoProv = Left$(RASOCLI$((-1) * NumeProv%), 48)
      !CoDocAp_Lar = Left$(CoComLar, 18)
      !CoDocAp_Cor = Left$(CoComCor$, 12)
      !ImpoApli = ImpApli#
      !RefeApli = Left$(RefApli$, 32)
      !CoDocOr_Lar = Left$(DocOrLar$, 18)
      !CoDocOr_Cor = Left$(DocOrCor$, 12)
      !ItotOrig = TotOrig#
      !IIVAORIG = IvaOrig#
      !Monemi_Ori = MonOri%
      !ImpoApli_Ori = ImpApliOr#
     .Update
   End With
   'ApliPago.Close
End Sub
'
Sub GRAPLICOB()
   If NOCLOSE% = 0 Then Call APERBASDAT("CABAN")
   With ApliCobra
     .AddNew
      !CodiEmpr = CodiEmp%
      !CodiSucu = CodiSuc%
      !NuClien = NumeClie%
      !CodiClie = Left$(CODICLI$(NumeClie%), 16)
      !RasoClie = Left$(RASOCLI$(NumeClie%), 48)
      !CoDocAp_Lar = Left$(CoComLar, 18)
      !CoDocAp_Cor = Left$(CoComCor$, 12)
      !ImpoApli = ImpApli#
      !RefeApli = Left$(RefApli$, 32)
      !CoDocOr_Lar = Left$(DocOrLar$, 18)
      !CoDocOr_Cor = Left$(DocOrCor$, 12)
      !ItotOrig = TotOrig#
      !IIVAORIG = IvaOrig#
     .Update
   End With
   'ApliCobra.Close
End Sub
'
Sub RAPLIPRO(NCX%)
    '
    Dim DOCULAR$(), DOCUCOR$(), IMPO#(), IORI#(), FFDOC%(), FVENDOC%(), STATDOC%(), CODIEM%(), CODISU%(), SALDOLA#()
    RAMAX& = 0
    ATENTI% = 1
    '
    MODOMONEI% = 0
    MODOMONE$ = Left$(TRIM$(UCase$(CONTROL$("GESACRED", "MODOMONE"))), 8)
    If MODOMONE$ = "MULTIMON" Then
       MODOMONEI% = 2
    End If
    '
    MONEMIA% = MONEMI%
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
        ReDim Preserve DOCULAR$(RAMAX&), DOCUCOR$(RAMAX&), IORI#(RAMAX&), IMPO#(RAMAX&), FFDOC%(RAMAX&), FVENDOC%(RAMAX&), CODIEM%(RAMAX&), CODISU%(RAMAX&), STATDOC%(RAMAX&), SALDOLA#(RAMAX&)
      End If
      DOCULAR$(NUDOCU%) = TRIM$(Mid$(XX$, 33, 24))
      DOCUCOR$(NUDOCU%) = TRIM$(Mid$(XX$, 57, 12))
      IMPO#(NUDOCU%) = CVD(Mid$(XX$, 25, 8))
      IORI#(NUDOCU%) = Abs(IMPO#(NUDOCU%))
      FFDOC%(NUDOCU%) = CVI(Mid$(XX$, 3, 2))
      CODIEM%(NUDOCU%) = CVI(Mid$(XX$, 69, 2))
      CODISU%(NUDOCU%) = CVI(Mid$(XX$, 71, 2))
      STATDOC%(NUDOCU%) = CVI(Mid$(XX$, 73, 2))
      SALDOLA#(NUDOCU%) = 0
    Next KU&
    '
    Call APERBASDAT("CABAN")
    SQLX$ = "SELECT * FROM CAJABANC "
    SQLX$ = SQLX$ + "WHERE NuProve = " & NCX% & " "
    SQLX$ = SQLX$ + "AND (TipoMovim = 'FCOM' "
    SQLX$ = SQLX$ + "OR TipoMovim = 'AJUP' "
    SQLX$ = SQLX$ + "OR TipoMovim = 'OPAG') "
    SQLX$ = SQLX$ + "ORDER BY FechEmisi, FeHoReal ASC;"
    Set CuecoTemp = CueCorri.OpenRecordset(SQLX$, dbOpenSnapshot)
    '
    With CuecoTemp
      On Error Resume Next
      .MoveLast
      If Err < 1 Then
        .MoveFirst
        Do While Not .EOF
          ' ESTE BLOQUE AGREGADO PARA UNIFICAR LOS AJUSTES POR DIFERENCIA DE CAMBIO
          ' JUNTO CON LA O/PAGO CORRESPONDIENTE EN EL CASO DE GRAMY - VANINA
          COCOLA$ = !CODICOM_LAR
          COCOCO$ = !CodiCom_Cor
          If Left$(COCOLA$, 4) = "AS-X" Then
            If Left$(!OBSERGEN, 6) = "O.Pago" Then
              COCOLA$ = "OP-X" + Mid$(!CODICOM_LAR, 5)
              COCOCO$ = "OP-X" + Mid$(!CodiCom_Cor, 5)
            End If
          End If
          If Left$(COCOLA$, 2) = "OP" Then
              For KKII% = 1 To NUDOCU%
                If DOCULAR$(KKII%) = COCOLA$ Then
                  IMPO#(KKII%) = IMPO#(KKII%) + !IDEBECOMP - !IHABECOMP
                  MEMI% = !Mone_Emis
                  If MEMI% > 1 Then
                     TICAEMI = !Tipo_Cam
                     If TICAEMI > 0.1 Then
                       SALDOLA#(KKII%) = SALDOLA#(KKII%) + (!IDEBECOMP - IHABECOMP) / TICAEMI
                     End If
                  End If
                  GoTo 119
                End If
              Next KKII%
          End If
          ' FIN DE BLOQUE AGREGADO
                    '
          NUDOCU% = NUDOCU% + 1
          If NUDOCU% > RAMAX& Then
            RAMAX& = RAMAX& + 128
            ReDim Preserve DOCULAR$(RAMAX&), DOCUCOR$(RAMAX&), IORI#(RAMAX&), IMPO#(RAMAX&), FFDOC%(RAMAX&), FVENDOC%(RAMAX&), CODIEM%(RAMAX&), CODISU%(RAMAX&), STATDOC%(RAMAX&), SALDOLA#(RAMAX&)
          End If
          DOCULAR$(NUDOCU%) = COCOLA$   ' !CodiCom_Lar
          DOCUCOR$(NUDOCU%) = COCOCO$   ' !CodiCom_Cor
          IMPO#(NUDOCU%) = !IDEBECOMP - !IHABECOMP
          IORI#(NUDOCU%) = !ValAbsComp
          FFDOC%(NUDOCU%) = CVINT(!FechEmisi)
          CODIEM%(NUDOCU%) = !CodiEmpr
          CODISU%(NUDOCU%) = !CodiSucu
          STATDOC%(NUDOCU%) = !Status
          '
          SALDOLA#(NUDOCU%) = 0
          MEMI% = !Mone_Emis
          If MEMI% > 1 Then
             TICAEMI = !Tipo_Cam
             If TICAEMI > 0.1 Then
               SALDOLA#(NUDOCU%) = IMPO#(NUDOCU%) / TICAEMI
             End If
          End If
          If MODOMONEI% <> 2 Then SALDOLA#(NUDOCU%) = 0
119     .MoveNext
        Loop
      End If
    End With
    '
    ' blanquea si pasa por cero
    SUMAPEN# = 0: SUMAPENDOL# = 0
    FIN& = NUDOCU%
    For KKI% = 1 To NUDOCU%
      U& = KKI%: Call TRACE(20, U&, FIN&)
      SUMAPEN# = SUMAPEN# + IMPO#(KKI%)
      SUMAPENDOL# = SUMAPENDOL# + SALDOLA#(KKI%)
      If Abs(SUMAPEN#) + Abs(SUMAPENDOL#) < 0.005 Then
        For KKJ% = 1 To KKI%
          IMPO#(KKJ%) = 0
          SALDOLA#(KKJ%) = 0
        Next KKJ%
      End If
    Next KKI%
    '
    SQLX$ = "SELECT * FROM APLIPAGO "
    SQLX$ = SQLX$ + "WHERE NuProve = " & NCX% & " "
    Set ApliTemp = CueCorri.OpenRecordset(SQLX$, dbOpenSnapshot)
    '
    With ApliTemp
      On Error Resume Next
      .MoveLast
      If Err < 1 Then
        .MoveFirst
        Do While Not .EOF
          IAPLI# = !ImpoApli
          IAPDOL# = !ImpoApli_Ori
          If MODOMONEI% <> 2 Then IAPDOL# = 0
          RECI$ = TRIM$(!CoDocAp_Lar)
          FACU$ = TRIM$(!CoDocOr_Lar)
          COEFIX = 1: If IAPLI# > 0 Then COEFIX = IAPLI# / IAPDOL#
          '
          IRECIMA# = 0: IFACMAX# = 0
          For KKI% = 1 To NUDOCU%
             If DOCULAR$(KKI%) = RECI$ Then
               If IMPO#(KKI%) > 0 Then
                 IRECIMA# = IRECIMA# + IMPO#(KKI%)
               End If
             End If
             If DOCULAR$(KKI%) = FACU$ Then
               If IMPO#(KKI%) < 0 Then
                 IFACMAX# = IFACMAX# - IMPO#(KKI%)
               End If
             End If
          Next KKI%
          '
          If IAPLI# > IRECIMA# Then IAPLI# = IRECIMA#
          If IAPLI# > IFACMAX# Then IAPLI# = IFACMAX#
          '
          IRECIMA# = IAPLI#
          IFACMAX# = IAPLI#
          '
          If IAPLI# > 0 Then
             For KKI% = 1 To NUDOCU%
                If DOCULAR$(KKI%) = RECI$ Then
                   If IMPO#(KKI%) > 0 Then
                      IAP1# = IRECIMA#
                      If IAP1# > Abs(IMPO#(KKI%)) Then IAP1# = Abs(IMPO#(KKI%))
                      IMPO#(KKI%) = IMPO#(KKI%) - IAP1#
                      SALDOLA#(KKI%) = SALDOLA#(KKI%) - IAPDOL#
                      If MODOMONEI% <> 2 Then SALDOLA#(KKI%) = 0
                      IRECIMA# = IRECIMA# - IAP1#
                   End If
                End If
                If DOCULAR$(KKI%) = FACU$ Then
                   If IMPO#(KKI%) < 0 Then
                      IAP2# = IFACMAX#
                      If IAP2# > Abs(IMPO#(KKI%)) Then IAP2# = Abs(IMPO#(KKI%))
                      IMPO#(KKI%) = IMPO#(KKI%) + IAP2#
                      SALDOLA#(KKI%) = SALDOLA#(KKI%) + IAPDOL#
                      If MODOMONEI% <> 2 Then SALDOLA#(KKI%) = 0
                      IFACMAX# = IFACMAX# - IAP2#
                   End If
                End If
             Next KKI%
          End If
          '
        .MoveNext
        Loop
      End If
    End With
    '
    SUMAPEN# = 0: SUMAPENDOL# = 0
    FIN& = NUDOCU%
    For KKI% = 1 To NUDOCU%
      U& = KKI%: Call TRACE(20, U&, FIN&)
      SUMAPEN# = SUMAPEN# + IMPO#(KKI%)
      SUMAPENDOL# = SUMAPENDOL# + SALDOLA#(KKI%)
      If MODOMONEI% <> 2 Then SUMAPENDOL# = 0
      If Abs(SUMAPEN#) + Abs(SUMAPENDOL#) < 0.005 Then
        For KKJ% = 1 To KKI%
          IMPO#(KKJ%) = 0
          SALDOLA#(KKJ%) = 0
        Next KKJ%
      End If
    Next KKI%
    '
    For KKI% = 1 To NUDOCU%
      For KKJ% = KKI% - 1 To 1 Step -1
        If Abs(IMPO#(KKI%) + IMPO#(KKJ%)) < 0.005 Then
          If Abs(SALDOLA#(KKI%) + SALDOLA#(KKJ%)) < 0.005 Then
            IMPO#(KKI%) = 0
            IMPO#(KKJ%) = 0
            SALDOLA#(KKI%) = 0
            SALDOLA#(KKJ%) = 0
          End If
        End If
      Next KKJ%
    Next KKI%
    
    KKM% = 0
    SUMAPEN# = 0: SUMAPENDOL# = 0
    For KKI% = 1 To NUDOCU%
      If STATDOC%(KKI%) = 8 Then
        SUMAPEN# = SUMAPEN# + IMPO#(KKI%)
        SUMAPENDOL# = SUMAPENDOL# + SALDOLA#(KKI%)
        If MODOMONEI% <> 2 Then SUMAPENDOL# = 0
      End If
    Next KKI%
    '
    If Abs(SUMAPEN#) + Abs(SUMAPENDOL#) < 0.005 Then
      For KKI% = 1 To NUDOCU%
        If STATDOC%(KKI%) = 8 Then
          IMPO#(KKI%) = 0
          SALDOLA#(KKI%) = 0
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
19  For U& = 1 To NUDOCU%
      If Abs(IMPO#(U&)) >= 0.005 Then
        ABUS$ = "NuProve = " & NCX% & " "
        ABUS$ = ABUS$ + "AND CodiEmpr = " & CODIEM%(U&) & " "
        ABUS$ = ABUS$ + "AND CodiSucu = " & CODISU%(U&) & " "
        ABUS$ = ABUS$ + "AND NuClien = 0 "
        ABUS$ = ABUS$ + "AND NuProve = " & NCX% & " "
        ABUS$ = ABUS$ + "AND CodiCom_Lar = '" & DOCULAR$(U&) & "' "
        '
        FVENDOC%(U&) = FFDOC%(U&)
        With SacrePen
          .FindFirst ABUS$
          If .NoMatch = False Then
            FVENDOC%(U&) = CVINT(!FEVENCIM)
            If FVENDOC%(U&) < FFDOC%(U&) Then
              FVENDOC%(U&) = FFDOC%(U&)
            End If
          End If
        End With
      End If
    Next U&
    '
    If CONTROLGRABA% = 1 Then GoTo 24
    Screen.MousePointer = 1
    TTXX$ = "Se ha Completado la Reconstrucción\de la Composición de la Deuda de\   \'" + TRIM$(RASOCLI$((-1) * NCX%))
    If COMALTER%(TTXX$ + "'\   \Desea Grabar la Nueva Composición de la Deuda ?\\Cancelar\Grabar") <> 2 Then
       GoTo 999
    End If
    '
24  Screen.MousePointer = 11
    YAPRIRE% = 0
    ABUS$ = "NuProve = " & NCX% & " "
    With SacrePen
25    If YAPRIRE% = 0 Then
          .FindFirst ABUS$
          YAPRIRE% = 1
        Else
          .FindNext ABUS$
      End If
      If .NoMatch = False Then
        .Delete
        GoTo 25
      End If
    End With
    '
    HOII% = HOY(HO$)
    FIN& = NUDOCU%
    For U& = 1 To NUDOCU%
      Call TRACE(20, U&, FIN&)
      If Abs(IMPO#(U&)) >= 0.005 Then    ' SOLO SE GUARDAN LOS QUE TIENEN SALDO EN PESOS - 16/09/07
        CoComCor$ = DOCUCOR$(U&)
        CoComLar$ = DOCULAR$(U&)
        CodiEmp% = CODIEM%(U&)
        CodiSuc% = CODISU%(U&)
        NumeClie% = 0
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
        COTIMONEM# = 0
        ImVencDol = (-1) * SALDOLA#(U&)
        If MODOMONEI% = 2 Then
           MONEMI% = MONEMISORP%(DOCULARANT$, NCX%) Mod 6
           If Abs(SALDOLA#(U&)) >= 0.005 Then
             COTIMONEM# = ImVenc# / ((-1) * SALDOLA#(U&))
           End If
           If MONEMI% < 2 Then ImVencDol = 0
        End If
If Left$(DOCULARANT$, 2) = "ND" Then
AAA = bbb
End If
        Call GRACREPEN1
        '
      End If
    Next U&
    '
999 Call CIERRARCH("CUECORR")
    Call LIBARCH("CUECORR")
    Screen.MousePointer = 1
    ATENTI% = 0
    MONEMI% = MONEMIA%
    '
End Sub

Sub RAPLIPROnue(NCX%)
    '
    Dim DOCULAR$(), DOCUCOR$(), IMPO#(), IORI#(), FFDOC%(), FVENDOC%(), STATDOC%(), CODIEM%(), CODISU%(), SALDOLA#()
    RAMAX& = 0
    ATENTI% = 1
    '
    MODOMONEI% = 0
    MODOMONE$ = Left$(TRIM$(UCase$(CONTROL$("GESACRED", "MODOMONE"))), 8)
    If MODOMONE$ = "MULTIMON" Then
       MODOMONEI% = 2
    End If
    '
    MONEMIA% = MONEMI%
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
        ReDim Preserve DOCULAR$(RAMAX&), DOCUCOR$(RAMAX&), IORI#(RAMAX&), IMPO#(RAMAX&), FFDOC%(RAMAX&), FVENDOC%(RAMAX&), CODIEM%(RAMAX&), CODISU%(RAMAX&), STATDOC%(RAMAX&), SALDOLA#(RAMAX&)
      End If
      DOCULAR$(NUDOCU%) = TRIM$(Mid$(XX$, 33, 24))
      DOCUCOR$(NUDOCU%) = TRIM$(Mid$(XX$, 57, 12))
      IMPO#(NUDOCU%) = CVD(Mid$(XX$, 25, 8))
      IORI#(NUDOCU%) = Abs(IMPO#(NUDOCU%))
      FFDOC%(NUDOCU%) = CVI(Mid$(XX$, 3, 2))
      CODIEM%(NUDOCU%) = CVI(Mid$(XX$, 69, 2))
      CODISU%(NUDOCU%) = CVI(Mid$(XX$, 71, 2))
      STATDOC%(NUDOCU%) = CVI(Mid$(XX$, 73, 2))
      SALDOLA#(NUDOCU%) = 0
    Next KU&
    '
    Call APERBASDAT("CABAN")
    SQLX$ = "SELECT * FROM CAJABANC "
    SQLX$ = SQLX$ + "WHERE NuProve = " & NCX% & " "
    SQLX$ = SQLX$ + "AND (TipoMovim = 'FCOM' "
    SQLX$ = SQLX$ + "OR TipoMovim = 'AJUP' "
    SQLX$ = SQLX$ + "OR TipoMovim = 'OPAG') "
    SQLX$ = SQLX$ + "ORDER BY FechEmisi, FeHoReal ASC;"
    Set CuecoTemp = CueCorri.OpenRecordset(SQLX$, dbOpenSnapshot)
    '
    With CuecoTemp
      On Error Resume Next
      .MoveLast
      If Err < 1 Then
        .MoveFirst
        Do While Not .EOF
          ' ESTE BLOQUE AGREGADO PARA UNIFICAR LOS AJUSTES POR DIFERENCIA DE CAMBIO
          ' JUNTO CON LA O/PAGO CORRESPONDIENTE EN EL CASO DE GRAMY - VANINA
          COCOLA$ = !CODICOM_LAR
          COCOCO$ = !CodiCom_Cor
          If Left$(COCOLA$, 4) = "AS-X" Then
            If Left$(!OBSERGEN, 6) = "O.Pago" Then
              COCOLA$ = "OP-X" + Mid$(!CODICOM_LAR, 5)
              COCOCO$ = "OP-X" + Mid$(!CodiCom_Cor, 5)
            End If
          End If
          If Left$(COCOLA$, 2) = "OP" Then
              For KKII% = 1 To NUDOCU%
                If DOCULAR$(KKII%) = COCOLA$ Then
                  IMPO#(KKII%) = IMPO#(KKII%) + !IDEBECOMP - !IHABECOMP
                  MEMI% = !Mone_Emis
                  If MEMI% > 1 Then
                     TICAEMI = !Tipo_Cam
                     If TICAEMI > 0.1 Then
                       SALDOLA#(KKII%) = SALDOLA#(KKII%) + (!IDEBECOMP - IHABECOMP) / TICAEMI
                     End If
                  End If
                  GoTo 119
                End If
              Next KKII%
          End If
          ' FIN DE BLOQUE AGREGADO
                    '
          NUDOCU% = NUDOCU% + 1
          If NUDOCU% > RAMAX& Then
            RAMAX& = RAMAX& + 128
            ReDim Preserve DOCULAR$(RAMAX&), DOCUCOR$(RAMAX&), IORI#(RAMAX&), IMPO#(RAMAX&), FFDOC%(RAMAX&), FVENDOC%(RAMAX&), CODIEM%(RAMAX&), CODISU%(RAMAX&), STATDOC%(RAMAX&), SALDOLA#(RAMAX&)
          End If
          DOCULAR$(NUDOCU%) = COCOLA$   ' !CodiCom_Lar
          DOCUCOR$(NUDOCU%) = COCOCO$   ' !CodiCom_Cor
          IMPO#(NUDOCU%) = !IDEBECOMP - !IHABECOMP
          IORI#(NUDOCU%) = !ValAbsComp
          FFDOC%(NUDOCU%) = CVINT(!FechEmisi)
          CODIEM%(NUDOCU%) = !CodiEmpr
          CODISU%(NUDOCU%) = !CodiSucu
          STATDOC%(NUDOCU%) = !Status
          '
          SALDOLA#(NUDOCU%) = 0
          MEMI% = !Mone_Emis
          If MEMI% > 1 Then
             TICAEMI = !Tipo_Cam
             If TICAEMI > 0.1 Then
               SALDOLA#(NUDOCU%) = IMPO#(NUDOCU%) / TICAEMI
             End If
          End If
          If MODOMONEI% <> 2 Then SALDOLA#(NUDOCU%) = 0
119     .MoveNext
        Loop
      End If
    End With
    '
    ' blanquea si pasa por cero
    SUMAPEN# = 0: SUMAPENDOL# = 0
    FIN& = NUDOCU%
    For KKI% = 1 To NUDOCU%
      U& = KKI%: Call TRACE(20, U&, FIN&)
      SUMAPEN# = SUMAPEN# + IMPO#(KKI%)
      SUMAPENDOL# = SUMAPENDOL# + SALDOLA#(KKI%)
      If Abs(SUMAPEN#) + Abs(SUMAPENDOL#) < 0.005 Then
        For KKJ% = 1 To KKI%
          IMPO#(KKJ%) = 0
          SALDOLA#(KKJ%) = 0
        Next KKJ%
      End If
    Next KKI%
    '
    SQLX$ = "SELECT * FROM APLIPAGO "
    SQLX$ = SQLX$ + "WHERE NuProve = " & NCX% & " "
    Set ApliTemp = CueCorri.OpenRecordset(SQLX$, dbOpenSnapshot)
    '
    With ApliTemp
      On Error Resume Next
      .MoveLast
      If Err < 1 Then
        On Error GoTo 0
        .MoveFirst
        Do While Not .EOF
          IAPLI# = XVALO(!ImpoApli)
          IAPDOL# = XVALO(!ImpoApli_Ori)
          If MODOMONEI% <> 2 Then IAPDOL# = 0
          RECI$ = TRIM$(!CoDocAp_Lar)
          FACU$ = TRIM$(!CoDocOr_Lar)
          COEFIX = 1: If IAPDOL# > 0 Then COEFIX = IAPLI# / IAPDOL#
          '
          IRECIMA# = 0: IFACMAX# = 0
          For KKI% = 1 To NUDOCU%
             If DOCULAR$(KKI%) = RECI$ Then
               If IMPO#(KKI%) > 0 Then
                 IRECIMA# = IRECIMA# + IMPO#(KKI%)
               End If
             End If
             If DOCULAR$(KKI%) = FACU$ Then
               If IMPO#(KKI%) < 0 Then
                 IFACMAX# = IFACMAX# - IMPO#(KKI%)
               End If
             End If
          Next KKI%
          '
          If IAPLI# > IRECIMA# Then IAPLI# = IRECIMA#
          If IAPLI# > IFACMAX# Then IAPLI# = IFACMAX#
          '
          IRECIMA# = IAPLI#
          IFACMAX# = IAPLI#
          '
          If IAPLI# > 0 Then
             For KKI% = 1 To NUDOCU%
                If DOCULAR$(KKI%) = RECI$ Then
                   If IMPO#(KKI%) > 0 Then
                      IAP1# = IRECIMA#
                      If IAP1# > Abs(IMPO#(KKI%)) Then IAP1# = Abs(IMPO#(KKI%))
                      IMPO#(KKI%) = IMPO#(KKI%) - IAP1#
                      ' ESTE BLOQUE SE AGREGÓ PARA LA CUENTA ELKO DE GRAMY
                      ' QUE FACTURA EN DOLARES Y PAGA EN PESOS
                      If IAPDOL# > SALDOLA#(KKI%) Then IAPDOL# = SALDOLA#(KKI%)
                      '
                      SALDOLA#(KKI%) = SALDOLA#(KKI%) - IAPDOL#
                      If MODOMONEI% <> 2 Then SALDOLA#(KKI%) = 0
                      IRECIMA# = IRECIMA# - IAP1#
                   End If
                End If
                If DOCULAR$(KKI%) = FACU$ Then
                   If IMPO#(KKI%) < 0 Then
                      IAP2# = IFACMAX#
                      If IAP2# > Abs(IMPO#(KKI%)) Then IAP2# = Abs(IMPO#(KKI%))
                      IMPO#(KKI%) = IMPO#(KKI%) + IAP2#
                      ' ESTE BLOQUE SE AGREGÓ PARA LA CUENTA ELKO DE GRAMY
                      ' QUE FACTURA EN DOLARES Y PAGA EN PESOS
                      If IAPDOL# > (-1) * SALDOLA#(KKI%) Then IAPDOL# = (-1) * SALDOLA#(KKI%)
                      '
                      SALDOLA#(KKI%) = SALDOLA#(KKI%) + IAPDOL#
                      If MODOMONEI% <> 2 Then SALDOLA#(KKI%) = 0
                      IFACMAX# = IFACMAX# - IAP2#
                   End If
                End If
             Next KKI%
          End If
          '
        .MoveNext
        Loop
      End If
    End With
    '
    SUMAPEN# = 0: SUMAPENDOL# = 0
    FIN& = NUDOCU%
    For KKI% = 1 To NUDOCU%
      U& = KKI%: Call TRACE(20, U&, FIN&)
      SUMAPEN# = SUMAPEN# + IMPO#(KKI%)
      SUMAPENDOL# = SUMAPENDOL# + SALDOLA#(KKI%)
      If MODOMONEI% <> 2 Then SUMAPENDOL# = 0
      If Abs(SUMAPEN#) + Abs(SUMAPENDOL#) < 0.005 Then
        For KKJ% = 1 To KKI%
          IMPO#(KKJ%) = 0
          SALDOLA#(KKJ%) = 0
        Next KKJ%
      End If
    Next KKI%
    '
    For KKI% = 1 To NUDOCU%
      For KKJ% = KKI% - 1 To 1 Step -1
        If Abs(IMPO#(KKI%) + IMPO#(KKJ%)) < 0.005 Then
          If Abs(SALDOLA#(KKI%) + SALDOLA#(KKJ%)) < 0.005 Then
            IMPO#(KKI%) = 0
            IMPO#(KKJ%) = 0
            SALDOLA#(KKI%) = 0
            SALDOLA#(KKJ%) = 0
          End If
        End If
      Next KKJ%
    Next KKI%
    
    KKM% = 0
    SUMAPEN# = 0: SUMAPENDOL# = 0
    For KKI% = 1 To NUDOCU%
      If STATDOC%(KKI%) = 8 Then
        SUMAPEN# = SUMAPEN# + IMPO#(KKI%)
        SUMAPENDOL# = SUMAPENDOL# + SALDOLA#(KKI%)
        If MODOMONEI% <> 2 Then SUMAPENDOL# = 0
      End If
    Next KKI%
    '
    If Abs(SUMAPEN#) + Abs(SUMAPENDOL#) < 0.005 Then
      For KKI% = 1 To NUDOCU%
        If STATDOC%(KKI%) = 8 Then
          IMPO#(KKI%) = 0
          SALDOLA#(KKI%) = 0
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
19  For U& = 1 To NUDOCU%
      If Abs(IMPO#(U&)) >= 0.005 Then
        ABUS$ = "NuProve = " & NCX% & " "
        ABUS$ = ABUS$ + "AND CodiEmpr = " & CODIEM%(U&) & " "
        ABUS$ = ABUS$ + "AND CodiSucu = " & CODISU%(U&) & " "
        ABUS$ = ABUS$ + "AND NuClien = 0 "
        ABUS$ = ABUS$ + "AND NuProve = " & NCX% & " "
        ABUS$ = ABUS$ + "AND CodiCom_Lar = '" & DOCULAR$(U&) & "' "
        '
        FVENDOC%(U&) = FFDOC%(U&)
        With SacrePen
          .FindFirst ABUS$
          If .NoMatch = False Then
            FVENDOC%(U&) = CVINT(!FEVENCIM)
            If FVENDOC%(U&) < FFDOC%(U&) Then
              FVENDOC%(U&) = FFDOC%(U&)
            End If
          End If
        End With
      End If
    Next U&
    '
    If CONTROLGRABA% = 1 Then GoTo 24
    Screen.MousePointer = 1
    TTXX$ = "Se ha Completado la Reconstrucción\de la Composición de la Deuda de\   \'" + TRIM$(RASOCLI$((-1) * NCX%))
    If COMALTER%(TTXX$ + "'\   \Desea Grabar la Nueva Composición de la Deuda ?\\Cancelar\Grabar") <> 2 Then
       GoTo 999
    End If
    '
24  Screen.MousePointer = 11
    YAPRIRE% = 0
    ABUS$ = "NuProve = " & NCX% & " "
    With SacrePen
25    If YAPRIRE% = 0 Then
          .FindFirst ABUS$
          YAPRIRE% = 1
        Else
          .FindNext ABUS$
      End If
      If .NoMatch = False Then
        .Delete
        GoTo 25
      End If
    End With
    '
    HOII% = HOY(HO$)
    FIN& = NUDOCU%
    For U& = 1 To NUDOCU%
      Call TRACE(20, U&, FIN&)
      If Abs(IMPO#(U&)) >= 0.005 Then    ' SOLO SE GUARDAN LOS QUE TIENEN SALDO EN PESOS - 16/09/07
        CoComCor$ = DOCUCOR$(U&)
        CoComLar$ = DOCULAR$(U&)
        CodiEmp% = CODIEM%(U&)
        CodiSuc% = CODISU%(U&)
        NumeClie% = 0
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
        COTIMONEM# = 0
        ImVencDol = (-1) * SALDOLA#(U&)
        If MODOMONEI% = 2 Then
           MONEMI% = MONEMISORP%(DOCULARANT$, NCX%) Mod 6
           If Abs(SALDOLA#(U&)) >= 0.005 Then
             COTIMONEM# = ImVenc# / ((-1) * SALDOLA#(U&))
           End If
           If MONEMI% < 2 Then ImVencDol = 0
        End If
If Left$(DOCULARANT$, 2) = "ND" Then
AAA = bbb
End If
        Call GRACREPEN1
        '
      End If
    Next U&
    '
999 Call CIERRARCH("CUECORR")
    Call LIBARCH("CUECORR")
    Screen.MousePointer = 1
    ATENTI% = 0
    MONEMI% = MONEMIA%
    '
End Sub

Function INEFACVIE#(NROX$, FEMI, NPROX%) ' IMPORTE NETO DE LA FACTURA
    '
    FFII% = FEMI
    FECHAD = Int(CDbl(CVDAT(FFII%)))
    SQLX$ = "SELECT * FROM CAJABANC "
    SQLX$ = SQLX$ + "WHERE INT(CDBL(FechConta)) = " & FECHAD & " "
    SQLX$ = SQLX$ + "AND CodiCom_Lar = '" & NROX$ & "' "
    SQLX$ = SQLX$ + "AND NuProve = " & Abs(NPROX%) & " "
    SQLX$ = SQLX$ + "AND TipoMovim = 'COMPRA' "
    '
    Call APERBASDAT("CABAN")
    Set CuecoTemp = CueCorri.OpenRecordset(SQLX$, dbOpenSnapshot)
    '
    SUNEFA# = 0
    With CuecoTemp
      On Error Resume Next
      .MoveLast
      If Err < 1 Then
        .MoveFirst
        Do While Not .EOF
          SUNEFA# = SUNEFA# + !IDEBECOMP - !IHABECOMP
        .MoveNext
        Loop
      End If
    End With
    '
    INEFACVIE# = SUNEFA#
    '
End Function

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
    FECHAD = Int(CDbl(CVDAT(FFII%)))
    SQLX$ = "SELECT * FROM CAJABANC "
    SQLX$ = SQLX$ + "WHERE TipoMovim = 'COMPRA' "
    SQLX$ = SQLX$ + "AND INT(CDBL(FechConta)) = " & FECHAD & " "
    SQLX$ = SQLX$ + "AND CodiCom_Lar = '" & NROX$ & "' "
    SQLX$ = SQLX$ + "AND NuProve = " & Abs(NPROX%) & " "
    Call APERBASDAT("CABAN")
    Set CUECOPAR = CueCorri.OpenRecordset(SQLX$, dbOpenDynaset)
    '
    SUNEFA# = 0
    With CUECOPAR
      On Error Resume Next
      .MoveFirst
10    If Err < 1 Then
        SUNEFA# = SUNEFA# + !IDEBECOMP - !IHABECOMP
        .MoveNext
        GoTo 10
      End If
    End With
    '
    NROXA$ = NROX$
    FEMIA = FEMI
    NPROXA% = NPROX%
    '
90  INEFAC# = SUNEFA#
    '
End Function

Function INEFACVIE1#(NROX$, FEMI, NPROX%) ' IMPORTE NETO DE LA FACTURA
    '
    Static CTX%, CUECOPAR
    If CTX% = 0 Then
      SQLX$ = "SELECT * FROM CAJABANC "
      SQLX$ = SQLX$ + "WHERE TipoMovim = 'COMPRA' "
      Call APERBASDAT("CABAN")
      Set CUECOPAR = CueCorri.OpenRecordset(SQLX$, dbOpenDynaset)
      CTX% = 1
    End If
    '
    FFII% = FEMI
    FECHAD = Int(CDbl(CVDAT(FFII%)))
    SQLX$ = "INT(CDBL(FechConta)) = " & FECHAD & " "
    SQLX$ = SQLX$ + "AND CodiCom_Lar = '" & NROX$ & "' "
    SQLX$ = SQLX$ + "AND NuProve = " & Abs(NPROX%) & " "
    '
    SUNEFA# = 0
    With CUECOPAR
      .FindFirst SQLX$
10    If .NoMatch = False Then
        SUNEFA# = SUNEFA# + !IDEBECOMP - !IHABECOMP
        .FindNext SQLX$
        GoTo 10
      End If
    End With
    '
    INEFACVIE1# = SUNEFA#
    '
End Function

Function MONEMISOR%(NROX$, FEMI, NCLIEN%) ' MONEDA DE EMISION DE LA FACTURA
    '
    SQLX$ = "SELECT * FROM CAJABANC "
    SQLX$ = SQLX$ + "WHERE (TipoMovim = 'FVEN' OR TipoMovim = 'RCOB' OR TipoMovim = 'AJUD') "
    SQLX$ = SQLX$ + "AND CodiCom_Lar = '" & NROX$ & "' "
    SQLX$ = SQLX$ + "AND NuClien = " & Abs(NCLIEN%) & " "
    Call APERBASDAT("CABAN")
    Set CUECOPAR = CueCorri.OpenRecordset(SQLX$, dbOpenSnapshot)
    '
    MONEMISOR% = 1
    FFII% = FEMI
    FECHAD = Int(CDbl(CVDAT(FFII%)))
    CUECOPAR.FindFirst "INT(CDBL(FechConta)) = " & FECHAD & " "
    If CUECOPAR.NoMatch = False Then
      On Error Resume Next
      MONEMISOR% = CUECOPAR!Mone_Emis
      On Error GoTo 0
    End If
    '
    If MONEMISOR% < 1 Then MONEMISOR% = 1
    MONEMISOR% = MONEMISOR% Mod 6   ' ENJUAGUE GRAMY
    '
End Function
'
Function MONEMISORP%(NROX$, NPROVE%) ' MONEDA DE EMISION DE LA FACTURA
    '
    SQLX$ = "SELECT * FROM CAJABANC "
    SQLX$ = SQLX$ + "WHERE (TipoMovim = 'FCOM' OR TipoMovim = 'OPAG' OR TipoMovim = 'AJUP') "
    SQLX$ = SQLX$ + "AND CodiCom_Lar = '" & NROX$ & "' "
    SQLX$ = SQLX$ + "AND NuProve = " & Abs(NPROVE%) & " "
    SQLX$ = SQLX$ + "ORDER BY FechConta DESC"
    Call APERBASDAT("CABAN")
    Set CUECOPAR = CueCorri.OpenRecordset(SQLX$, dbOpenSnapshot)
    '
    MONEMISORP% = 1
    On Error Resume Next
    CUECOPAR.MoveFirst
    If Err < 1 Then
      On Error GoTo 0
      MONEMISORP% = CUECOPAR!Mone_Emis
    End If
    On Error GoTo 0
    '
    If MONEMISORP% < 1 Then MONEMISORP% = 1
    MONEMISORP% = MONEMISORP% Mod 6   ' ENJUAGUE GRAMY
    '
End Function
'
Function TICAMISORP(NROX$, FEMI, NPROVE%) ' TIPO DE CAMBIO DE EMISION DE LA FACTURA
    '
    SQLX$ = "SELECT * FROM CAJABANC "
    SQLX$ = SQLX$ + "WHERE (TipoMovim = 'FCOM' OR TipoMovim = 'OPAG' OR TipoMovim = 'AJUP') "
    SQLX$ = SQLX$ + "AND CodiCom_Lar = '" & NROX$ & "' "
    SQLX$ = SQLX$ + "AND NuProve = " & Abs(NPROVE%) & " "
    Call APERBASDAT("CABAN")
    Set CUECOPAR = CueCorri.OpenRecordset(SQLX$, dbOpenSnapshot)
    '
    TICAMISORP = 1
    FFII% = FEMI
    FECHAD = Int(CDbl(CVDAT(FFII%)))
    CUECOPAR.FindFirst "INT(CDBL(FechConta)) = " & FECHAD & " "
    If CUECOPAR.NoMatch = False Then
      On Error Resume Next
      TICAMISORP = CUECOPAR.Tipo_Cam
      On Error GoTo 0
    End If
    '
    If TICAMISORP < 0.01 Then TICAMISORP = 1
    '
End Function
'
Function IMPORORIG(NROX$, NPROVE%) ' MONEDA DE EMISION DE LA FACTURA
    '
    SQLX$ = "SELECT * FROM CAJABANC "
    SQLX$ = SQLX$ + "WHERE (TipoMovim = 'FCOM' OR TipoMovim = 'OPAG' OR TipoMovim = 'AJUP') "
    SQLX$ = SQLX$ + "AND CodiCom_Lar = '" & NROX$ & "' "
    SQLX$ = SQLX$ + "AND NuProve = " & Abs(NPROVE%) & " "
    SQLX$ = SQLX$ + "ORDER BY FechConta DESC"
    Call APERBASDAT("CABAN")
    Set CUECOPAR = CueCorri.OpenRecordset(SQLX$, dbOpenSnapshot)
    '
    IMPORORIG = 0
    On Error Resume Next
    CUECOPAR.MoveFirst
    If Err < 1 Then
      On Error GoTo 0
      IMPORORIG = CUECOPAR!IHABECOMP - CUECOPAR!IDEBECOMP
    End If
    On Error GoTo 0
    '
End Function
'
Function TICAMISOR(NROX$, FEMI, NCLIEN%) ' TIPO DE CAMBIO DE EMISION DE LA FACTURA
    '
    SQLX$ = "SELECT * FROM CAJABANC "
    SQLX$ = SQLX$ + "WHERE (TipoMovim = 'FVEN' OR TipoMovim = 'RCOB' OR TipoMovim = 'AJUD') "
    SQLX$ = SQLX$ + "AND CodiCom_Lar = '" & NROX$ & "' "
    SQLX$ = SQLX$ + "AND NuClien = " & Abs(NCLIEN%) & " "
    Call APERBASDAT("CABAN")
    Set CUECOPAR = CueCorri.OpenRecordset(SQLX$, dbOpenSnapshot)
    '
    TICAMISOR = 1
    FFII% = FEMI
    FECHAD = Int(CDbl(CVDAT(FFII%)))
    CUECOPAR.FindFirst "INT(CDBL(FechConta)) = " & FECHAD & " "
    If CUECOPAR.NoMatch = False Then
      On Error Resume Next
      TICAMISOR = CUECOPAR.Tipo_Cam
      On Error GoTo 0
    End If
    '
    If TICAMISOR < 0.01 Then TICAMISOR = 1
    '
End Function
'
Sub COMPENSACP()
'\\\OT-06-0112-WAR-30/03/06///
  VDEF$ = Space$(61)
  HOII% = HOY(HO$)
  '
  Call REPLA(VDEF$, MKI$(HOII%), 1, 2)
  Call REPLA(VDEF$, MKI$(0), 3, 2)
  Call REPLA(VDEF$, MKI$(0), 5, 2)
  '
  CUECOMPE$ = CONTROL$("SVD451S", "CUEVE-10")
  If CUECOMPE$ = "" Or CUECOINT%(CUECOMPE$) < 0 Then
    Call MENSERR(24, "Imposible Compensar Cuentas.\  \Falta Definir Cuenta Contable\Correspondiente a Ajustes de Saldos Deudores.")
    Exit Sub
  End If
  '
5 OBJ1$ = OBJPLA$("COMPENSACP", VDEF$)
  If TRIM(OBJ1$) <> "" Then
    VDEF$ = OBJ1$
    'VALIDACIONES
    F1% = CVI(Mid$(OBJ1$, 1, 2))
    If F1% < 1 Or F1% > HOII% Then
      Call MENSERR(24, "Fecha no Válida")
      GoTo 5
    End If
    NC1% = CVI(Mid$(OBJ1$, 3, 2))
    If NC1% < 1 Or TRIM$(RASOCLI$(NC1%)) = "" Then
      Call MENSERR(24, "Cliente no Válido")
      GoTo 5
    End If
    NP1% = CVI(Mid$(OBJ1$, 5, 2))
    If NP1% < 1 Or TRIM$(RASOCLI$(NP1% * (-1))) = "" Then
      Call MENSERR(24, "Proveedor no Válido")
      GoTo 5
    End If
    IP1# = CVD(Mid$(OBJ1$, 7, 8))
    If IP1# < 0.005 Then
      Call MENSERR(24, "Importe no Válido")
      GoTo 5
    End If
    REF1$ = Mid$(OBJ1$, 15)
    If TRIM$(REF1$) = "" Then
      Call MENSERR(24, "Referencia no Válida")
      GoTo 5
    End If
    'PROCESO NUMERO DE COMPESACION
    FEX = F1%
    FETX$ = FECHATEX$(FEX)
    NRECIB$ = "CC-" + Mid$(FETX$, 7, 2) + Mid$(FETX$, 4, 2) + Mid$(FETX$, 1, 2) + "-"
    '
    SQLX$ = "SELECT * FROM CAJABANC "
    SQLX$ = SQLX$ + "WHERE TipoMovim = 'AJUD' "
    SQLX$ = SQLX$ + "AND CodiCom_Cor LIKE '" & NRECIB$ + "*' "
    SQLX$ = SQLX$ + "ORDER BY CodiCom_Cor DESC"
    '
    ORMA% = (-1)
    Set RECITEMP = CueCorri.OpenRecordset(SQLX$, dbOpenSnapshot)
    With RECITEMP
      On Error Resume Next
      .MoveFirst
      If Err < 1 Then
        ORMA% = XVALO(Mid$(!CodiCom_Cor, 11, 2))
      End If
    End With
    ORMX$ = TRIM$(Str$(ORMA% + 1))
    While Len(ORMX$) < 2: ORMX$ = "0" + ORMX$: Wend
    NRECIB$ = NRECIB$ + Left$(ORMX$, 2)
    'GRACABAN PARA LA CTA. CTE. DEL PROVEEDOR
20  NC% = (-1) * NP1%
    Call APERBASDAT("CABAN")
    '
    CoComCor$ = NRECIB$
    CoComLar$ = NRECIB$
    CodiEmp% = 0
    CodiSuc% = SUC%
    NumeClie% = 0
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
    Descrip = "Compensacion - " + RASOCLI$(NC1%)
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
    CodiEmp% = 0
    CodiSuc% = SUC%
    NumeClie% = NC1%
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
    CuentaCo% = CUECOINT%(CUEMADRE$(NC1%))
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
    VaMovi = "COMP-C"          ' NO IMPUTADO
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
    '\\\OT-06-0144-WAR-18/04/06///
    'IMPRESIÓN DE LA COMPENSACIÓN
    '
    FORIPRE$ = TRIM$(CONTROL$("", "FORCOMPE"))
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
    TETABU1$(1, CAITAB1%) = TRIM$(Str$(NC1%)) 'CLIENTE
    TETABU1$(2, CAITAB1%) = TRIM$(RASOCLI$(NC1%)) 'CLIENTE
    CAITAB1% = CAITAB1% + 1
    TETABU1$(1, CAITAB1%) = TRIM$(Str$(NP1%)) 'PROVEEDOR
    TETABU1$(2, CAITAB1%) = TRIM$(RASOCLI$(-1 * (NP1%))) 'PROVEEDOR
    '
    Call PRINTDOC("FORCOMPE")
    '\\\OT-06-0144-WAR-FIN///
    '
    Call COMUNI("Compensación Procesada.")
    CajaBanc.Close
    '
  End If
'\\\OT-06-0112-WAR-FIN///
End Sub

Function SALDEUPEN#(NUFACX$)
   '
   ' DEVUELVE EL SALDO DEUDOR PENDIENTE DEL COMPROBANTE LARGO O CORTO NUFACX$
   '
   SDP# = 0
   NNX$ = TRIM$(NUFACX$)
   If NNX$ <> "" Then
      Call ABRECABAN
      SQLY$ = "SELECT * FROM SADEUPEN "
      SQLY$ = SQLY$ + "WHERE CodiCom_Lar = '" & NNX$ & "' OR CodiCom_Cor = '" & NNX$ & "'"
      Set SADEUPE = CueCorri.OpenRecordset(SQLY$, dbOpenSnapshot)
      With SADEUPE
         On Error Resume Next
         .MoveFirst
         If Err Then
            On Error GoTo 0
            GoTo 20                  ' NO FIGURA PENDIENTE
         End If
         On Error GoTo 0
         Do While Not .EOF
            SDP# = SDP# + !SaldDebe - !SaldAcre
         .MoveNext
         Loop
      End With
   End If
   '
20 SALDEUPEN# = SDP#
   '
End Function

Function SALDACREPEN#(NUFACX$, NPRX%, MEMIX%)
   '
   ' DEVUELVE EL SALDO ACREEDOR PENDIENTE DEL COMPROBANTE LARGO O CORTO NUFACX$
   '
   SDP# = 0
   NNX$ = TRIM$(NUFACX$)
   If NNX$ <> "" Then
      Call ABRECABAN
      SQLY$ = "SELECT * FROM SACREPEN "
      SQLY$ = SQLY$ + "WHERE (CodiCom_Lar = '" & NNX$ & "' OR CodiCom_Cor = '" & NNX$ & "') "
      SQLY$ = SQLY$ + "AND NuProve = " & NPRX%
      Set SADEUPE = CueCorri.OpenRecordset(SQLY$, dbOpenSnapshot)
      With SADEUPE
         On Error Resume Next
         .MoveFirst
         If Err Then
            On Error GoTo 0
            GoTo 20                  ' NO FIGURA PENDIENTE
         End If
         On Error GoTo 0
         Do While Not .EOF
            If MEMIX% <= 1 Then
                SDP# = SDP# + !SaldAcre - !SaldDebe
              Else
                SDP# = SDP# + !SaldAcre_Dola - !SaldDebe_Dola
            End If
         .MoveNext
         Loop
      End With
   End If
   '
20 SALDACREPEN# = SDP#
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
    With SaDeuPen
      .FindFirst "NuClien = " & NCX% & " AND CODICOM_LAR = '" & DOCULA$ & "'"
201   If .NoMatch = False Then
        .Edit
        If SALDOLA# > 0 Then
             !SaldDebe_Dola = SALDOLA#
             !SaldAcre_Dola = 0
           Else
             !SaldDebe_Dola = 0
             !SaldAcre_Dola = Abs(SALDOLA#)
        End If
        .Update
        SALDOLA# = 0
        .FindNext "NuClien = " & NCX% & " AND CODICOM_LAR = '" & DOCULA$ & "'"
        GoTo 201
      End If
    End With
    '
End Sub
Function YAESTACHEQUE%(BANCO$, CHEQUE$)
    '
    YAESTACHK% = 0
    ABUS$ = AJUSTI$(BANCO$, 18)
    ABUS$ = ABUS$ + AJUSTI$(CHEQUE$, 14)
    '
    Call APERBASDAT("CABAN")
    SQLX$ = "SELECT * FROM CAJABANC"
    SQLX$ = SQLX$ + " WHERE TIPOMOVIM = 'CHRECI' "
    SQLX$ = SQLX$ + " AND DESCRIMOV LIKE '" & ABUS$ + "*'"
    Set OrpaTemp = CueCorri.OpenRecordset(SQLX$, dbOpenSnapshot)
    '
    With OrpaTemp
      On Error Resume Next
      .MoveFirst
      If Err < 1 Then
        On Error GoTo 0
        YAESTACHK% = 1
      End If
    End With
    '
    YAESTACHEQUE% = YAESTACHK%
    '
End Function

