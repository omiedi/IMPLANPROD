Attribute VB_Name = "CUECORR1"
Public MODORPA%
Public MODOSORVAL%
Public TIDOCOM$, LETRAFA$, NUDOCUCOM$
Public TOFAC#(10), FVII%(10), IMV#(10), IMV1#(10), GRU%(10), IPAR#(10)
'
Public NEGRA#(8), TASIVA(8), IMPIVA#(8)
Public PERAFIP#(4)
Public PERIBRU#(2), JURIBRU$(2)
'
Public TOTIMPU#, TONEGRA#, TOEXEN#, TOTIVA#, TORETIV#, TOTFAC#, ALICUIVA

'
Sub CUECORR(FE%, NC%, TI%, VA%, PRF%, NRO&, REFE$, MO$, TOFAC#(), FV%(), SD#(), CUE%(), IMC#(), OPCI%, FEMIREA%)

Static CAREG&, USER%, LU$, N1%, N11%, N21%, A0$, A1$, A2$, X02$, X04$, CTX%

' OPCI% = 1 - GRABA FACTURA
'         2 - VERIFICA NUMERO REPETIDO
'         3 - GRABA SIN VERIFICAR REPETICION Y CERRAR ARCHIVOS
'         4 - DETERMINA PROXIMO NUMERO DEBITOS PROPIOS

    USER% = USNBR% Mod 100
    If USER% < 1 Then USER% = 1
    LU$ = LUDAT$

    If (OPCI% = 3 And CTX% > 0) Then GoTo 31
    CTX% = 1

    CAREG& = ULTREG&("CUECORR")

31  If OPCI% = 4 Then
       Screen.MousePointer = 11
       ULNUM& = 0
       For I& = CAREG& To 1 Step -1
          A0$ = REGLEIDO$("CUECORR", I&)
          TNV$ = Mid$(A0$, 7, 2)
          If TNV$ = "DP" Then
             NRV& = XVALO(Mid$(A0$, 17, 8))
             If NRV& > ULNUM& Then ULNUM& = NRV&
             CDP% = CDP% + 1
             If CDP% >= 10 Then GoTo 32
           End If
       Next I&
       '
32     Screen.MousePointer = 1
       NRO& = ULNUM& + 1
       GoTo 4000
    End If
    
    If TI% = 42 Then
       TINUVA$ = DOCUACTIVO$
       If TRIM$(TINUVA$) = "" Then
          TINUVA$ = NUDOCUCOM$
       End If
       GoTo 35
    End If
    '
    PRFD# = CDbl(PRF%)
    NROD# = CDbl(NRO&)
    TIX% = TI% Mod 10
    If TIX% = 0 Then TINUVA$ = "AS"
    If TIX% = 1 Then TINUVA$ = "VC"
    If TIX% = 2 Then
       TINUVA$ = "FC"
       If NC% < 0 Then
          If TI% = 22 Then TINUVA$ = "TF"
          If TI% = 32 Then TINUVA$ = "RH"
       End If
    End If
    If TIX% = 3 Then TINUVA$ = "ND"
    If TIX% = 4 Then
       TINUVA$ = "RB"
       If NC% < 0 Then
          TINUVA$ = "OP"
          If TI% = 14 Then TINUVA$ = "FF"
       End If
    End If
    '
    If TIX% = 5 Then TINUVA$ = "DC"
    If TIX% = 6 Then TINUVA$ = "NC"
    If TIX% = 7 Then TINUVA$ = "DP"
    '
    VAS$ = Chr$(64 + VA%)
    TINUVA$ = TINUVA$ + "-" + VAS$ + "-" + NUMSTRI0$(PRFD#, 4, 0, 0) + "-" + NUMSTRI0$(NROD#, 8, 0, 0)
    '
35  NUDOCUCOM$ = TINUVA$
    If OPCI% = 3 Then GoTo 39
    GoSub 500                                   ' analiza repeticion de numero

    If (OPCI% = 1 And REPEI% > 0) Then
       Call MENSERR(24, "Operacion Abortada por Repeticion de Documento " + TIVANU$)
       OPCI% = (-1)
    End If
    '
    If REPEI% > 0 Then GoTo 200                  ' no graba si esta repetido
    If OPCI% = 2 Then GoTo 200                  ' verifica numero repetido

39  RECORD$ = MKI$(FE%)
    If NC% > 0 Then RECORD$ = RECORD$ + MKI$(NC%) + MKI$(0)
    If NC% <= 0 Then RECORD$ = RECORD$ + MKI$(0) + MKI$((-1) * NC%)


    RECORD$ = RECORD$ + TINUVA$

    REFE1$ = Left$(REFE$ + Space$(24), 24)
    REFE2$ = Mid$(REFE$, 25)
    RECORD$ = RECORD$ + REFE1$

    MO$ = Left$(MO$ + Space$(2), 2)
    RECORD$ = RECORD$ + MO$

    Total$ = ""
    For U% = 1 To 8
       Total$ = Total$ + MKD$(TOFAC#(U%))
    Next U%
    RECORD$ = RECORD$ + Total$

    TOTALD# = TOFAC#(0)
    If TOTALD# > 0 Then
       RECORD$ = RECORD$ + MKD$(TOTALD#) + String$(8, 0) + String$(8, 0)
       GoTo 49
    End If
    RECORD$ = RECORD$ + String$(8, 0) + MKD$(Abs(TOTALD#)) + String$(8, 0)

49  ' continue

    VENCI$ = ""
    For UI% = 1 To 6
       VENCI$ = VENCI$ + MKI$(FV%(UI%)) + MKD$(SD#(UI%))
    Next UI%
    RECORD$ = RECORD$ + VENCI$

    IMPU$ = ""
    For UI% = 1 To 8
       IMPU$ = IMPU$ + MKI$(CUE%(UI%)) + MKD$(IMC#(UI%))
    Next UI%
    RECORD$ = RECORD$ + IMPU$

99  RECORD$ = RECORD$ + MKI$(FEMIREA%) + String$(9, 0)

    HOO = HOY(HO$)
    While HOO > 32767
       HOO = HOO - 10000
    Wend
    HOI% = HOO
    RECORD$ = RECORD$ + MKI$(HOI%)

    HP = 100 * XVALO(Left$(Time$, 2)) + XVALO(Mid$(Time$, 4, 2))
    While HP > 32767
       HP = HP - 10000
    Wend
    HPI% = HP
    RECORD$ = RECORD$ + MKI$(HPI%)

    RECORD$ = RECORD$ + Chr$(USER%)
    CUIXX$ = ""
    For KKII% = 1 To Len(REFE2$)
      KKJJ% = Asc(Mid$(REFE2$, KKII%))
      If KKJJ% > 47 Then
        If KKJJ% < 58 Then
          CUIXX$ = CUIXX$ + Chr$(KKJJ%)
        End If
      End If
    Next KKII%
    CUIXX$ = Left$(CUIXX$ + "000000000000", 12)
    CUIYY$ = ""
    For KKII% = 1 To 12 Step 2
      CUIYY$ = CUIYY$ + Chr$(XVALO(Mid$(CUIXX$, KKII%, 2)))
    Next KKII%
    RECORD$ = RECORD$ + CUIYY$
    '
110 Call REGAPP("CUECORR", RECORD$)
    CAREG& = ULTREG&("CUECORR")


150 ' falta actualizar saldo de la cuenta

200 If OPCI% = 3 Then GoTo 4000

    If OPCI% = 2 Or REPEI% > 0 Then GoTo 4000
    RGX& = REGICLI&(NC%)
    If RGX& > 0 Then
       If RGX& <= ULTREG&("PROVED2") Then
          DX$ = REGLEIDO$("PROVED2", RGX&)
          SALDO# = CVD(Mid$(DX$, 109, 8)) - TOTALD#
          Call REPLA(DX$, MKD$(SALDO#), 109, 8)
          Call GRAREG("PROVED2", DX$, RGX&)
       End If
    End If
    '
    GoTo 4000
    '
500 If OPCI% = 2 Then
       Screen.MousePointer = 11
    End If
    REPEI% = 0
    '
    TINUVA1$ = TINUVA$: Mid$(TINUVA1$, 4, 1) = "X"
    NCX% = NC%
    PPX% = 2
    If NCX% <= 0 Then
       NCX% = Abs(NC%)
       PPX% = 4
    End If
    ARGU$ = MKI$(NCX%)
    RFC$ = RECFILT$("CUECORR", PPX%, MKI$(NCX%))
    '
    For UU& = 1 To Len(RFC$) Step 4
       REG& = CVS(Mid$(RFC$, UU&, 4))
       If (REG& < 1 Or REG& > ULTREG&("CUECORR")) Then
          Call MENSERR(24, "Error de Acceso a Registro " + Str$(REG&) + " de Archivo 'CUECORR.DAT'")
          Call FINAL("")
       End If
       '
       A0$ = REGLEIDO$("CUECORR", REG&)
       TINUVB$ = Mid$(A0$, 7, 18)
       Mid$(TINUVB$, 4, 1) = "X"
       If UCase$(Left$(TINUVA1$, 2)) <> "AS" Then
         If UCase$(TINUVB$) = UCase$(TINUVA1$) Then
           REPEI% = 1: NRO& = 0
           Screen.MousePointer = 1
           Call MENSERR(24, "Documento " + TINUVA$ + " Repetido")
           GoTo 599
         End If
       End If
509 Next UU&
    '
    Screen.MousePointer = 1
599 Return


4000 End Sub
'

Sub APLIPAG(FE%, NC%, TI%, VA%, PRF%, NRO&, IMAPLI#, REFE$, DOCORI$, ITOTORI#, IVAORI#)


    ARAPLI$ = "APLICOB": If NC% < 0 Then ARAPLI$ = "APLIPAG"
    PRFD# = CDbl(PRF%)
    NROD# = CDbl(NRO&)
    '
    TIX% = TI% Mod 10
    TINUVA$ = "FC"                                               ' por default
    If TIX% = 0 Then TINUVA$ = "AS"
    If TIX% = 1 Then TINUVA$ = "VC"
    If TIX% = 3 Then TINUVA$ = "ND"
      If TIX% = 4 Then
        TINUVA$ = "RB"
        If NC% < 0 Then
          TINUVA$ = "OP"
          If TI% = 14 Then TINUVA$ = "FF"
        End If
      End If
    If TIX% = 5 Then TINUVA$ = "DC"
    If TIX% = 6 Then TINUVA$ = "NC"
    If TIX% = 7 Then TINUVA$ = "DP"
    '
    If TIX% = 2 Then
      TINUVA$ = "FC"
      If NC% < 0 Then
        If TI% = 22 Then TINUVA$ = "TF"
        If TI% = 32 Then TINUVA$ = "RH"
      End If
    End If

    VAS$ = Chr$(64 + VA%): If VAS$ < "A" Or VAS$ > "Z" Then VAS$ = "X"
    TINUVA$ = TINUVA$ + "-" + VAS$ + "-" + NUMSTRI0$(PRFD#, 4, 0, 0) + "-" + NUMSTRI0$(NROD#, 8, 0, 0)

    RECORD$ = MKI$(FE%)
    RECORD$ = RECORD$ + MKI$(Abs(NC%))

    RECORD$ = RECORD$ + TINUVA$
    RECORD$ = RECORD$ + MKD$(IMAPLI#)

    REFE$ = Left$(REFE$ + Space$(24), 24)
    RECORD$ = RECORD$ + REFE$

    RECORD$ = RECORD$ + DOCORI$ + MKD$(ITOTORI#) + MKD$(IVAORI#)

    RECORD$ = RECORD$ + String$(8, 0)

    Call REGAPP(ARAPLI$, RECORD$)
    Call CIERRARCH(ARAPLI$)

4000 End Sub
'


Sub NUORPA(NROPRO&)
    '
    Screen.MousePointer = 11
    '
    FENUE$ = CONTROL$("", "FENUORPA")
    FLIM% = FECHANUM(FENUE$)
    If FLIM% > HOY(HO$) Then FLIM% = HOY(HO$)
    '
    CAMOLE% = 0
    NROMA& = 0
    For U& = ULTREG&("CUECORR") To 1 Step -1
        XX$ = REGLEIDO$("CUECORR", U&)
        CAMOLE% = CAMOLE% + 1
        If CAMOLE% > 512 Then GoTo 90
        '
        If Mid$(XX$, 7, 2) = "OP" Then
           FEMO% = CVI(Mid$(XX$, 1, 2))
           If FEMO% >= FLIM% Then
              NROLE& = XVALO(Mid$(XX$, 18, 7))
              If NROLE& = NROPRO& Then
                If Mid$(XX$, 12, 4) <> "anul" Then
                  NROPRO& = 0
                End If
              End If
              If NROLE& > NROMA& Then NROMA& = NROLE&
           End If
        End If
    Next U&
    '
90  If NROPRO& < 1 Then NROPRO& = 1 + NROMA&
    Screen.MousePointer = 1
    '
End Sub

Sub SIGUEORPA1()
    '
    Screen.MousePointer = 11
    If MODORPA% = 0 Then
        MODORPA% = 1       ' por emision
        If CONTROL$("ORPASTA", "MODORDE") = "VENCIM" Then
           MODORPA% = 2
        End If
    End If
    '
    NCXI = XVALO(FORPASTA.Text1.TEXT)
    If NCXI < 1 Or NCXI > 32767 Then
        Call MENSERR(24, "Numero de Cuenta no Válido")
        FORPASTA.Text1.SetFocus
        Exit Sub
    End If
    '
    NC% = NCXI
    If ECOARCH$("PROVED1", MKI$(NC%)) = "" Then
        Call MENSERR(24, "Cuenta no Figura\en Base de Datos de Proveedores")
        FORPASTA.Text1.SetFocus
        Exit Sub
    End If
    '
    FE$ = TRIM$(FORPASTA.Text9.TEXT)
    FF% = FECHANUM(FE$)
    HOI% = HOY(HO$)
    If FF% < 1 Or FF% > HOI% + 1 Then
        Call MENSERR(24, "Fecha no Válida")
        FORPASTA.Text9.SetFocus
        Exit Sub
    End If
    '
    If DIENTRE%(FF%, HOI%) > 55 Then
        Call MENSERR(24, "Fecha no Válida")
        FORPASTA.Text9.SetFocus
        Exit Sub
    End If
    '
    FECHALIM% = XVALO(CONTROL$("", "CIEMECOM"))
    If FF% <= FECHALIM% Then
        Call MENSERR(24, "Fecha no Válida.\  \Ese Período Mensual Está Cerrado.")
        FORPASTA.Text9.SetFocus
        Exit Sub
    End If
     '
    Call CIERRARCH("*.*")
    '
    Screen.MousePointer = 11
    NC% = XVALO(FORPASTA.Text1.TEXT)
    If NC% > 0 Then
        FORPASTA.Text2.TEXT = RASOCLI$((-1) * NC%)
        FORPASTA.Text3.TEXT = DOMICLI$((-1) * NC%)
        FORPASTA.Text4.TEXT = CPOSCLI$((-1) * NC%)
        FORPASTA.Text5.TEXT = LOCACLI$((-1) * NC%)
        'FORPASTA.Text7.TEXT = TEL1CLI$(NC%)
        'FORPASTA.Text10.TEXT = CSTRING$(MKD$(SALDCLI#(NC%)), 4, 13, 2, 2)
        FORPASTA.Refresh
    End If
    '
100 FORPASTA.List1.Clear
    FORPASTA.List1.Refresh
    FORPASTA.List2.Clear
    FORPASTA.List2.Refresh
    '
    z$ = Space$(128)
    Call REPLA(z$, "Anticipo no Aplicado", 1, 20)
    Call REPLA(z$, HO$, 23, 8)
    Call REPLA(z$, "       0", 119, 8)
    Call REPLA(z$, " 0", 127, 2)
    FORPASTA.List2.AddItem z$
    SUMADEU# = 0: SUMAFAC# = 0
    '
    If TICUEPRO%((-1) * NC%) = 1 Then        ' GASTOS
       Call BLOQARCH("!CARGAST")
       Call BLANARCH("!CARGAST")
       GoTo 200
    End If
    '
    MAXDEU& = ULTREG&("CUECORR")
    RFD$ = RECFILT$("CUECORR", 4, MKI$(NC%))
    '
    RESCUPRO.LINDICE.Clear
    For U& = 1 To Len(RFD$) Step 4
        REG& = CVS(Mid$(RFD$, U&, 4))
        x$ = REGLEIDO$("CUECORR", REG&)
        If CVI(Mid$(x$, 5, 2)) <> NC% Then
            x$ = REGLEIDO$("CUECORR", 1)
            Call GRAREG("CUECORR", x$, 1)
            Call CIERRARCH("CUECORR")
            GoTo 100
        End If
        '
        FEMI = CVI(Left$(x$, 2))
        IMPO# = CVD(Mid$(x$, 115, 8)) - CVD(Mid$(x$, 123, 8))
        SALDEU$ = Mid$(x$, 139, 60)
        For KKI% = 1 To 6
           FEVE = CVI(Mid$(SALDEU$, 10 * KKI% - 9, 2))
           SALDO# = CVD(Mid$(SALDEU$, 10 * KKI% - 7, 8))
           If Abs(SALDO#) >= 0.005 Then
              If MODORPA% = 1 Then
                  FF% = FEMI      ' orden emision
                Else
                  FF% = FEVE      ' orden vencimiento
              End If
              z$ = CSTRING$(MKI$(FF%), 1, 6, 0, 2) + CSTRING$(MKS$(REG&), 4, 8, 0, 2) + CSTRING$(MKI$(KKI%), 1, 4, 0, 2)
              RESCUPRO.LINDICE.AddItem z$
           End If
        Next KKI%
    Next U&
    '
    SUMADEU# = 0: SUMAFAC# = 0
    For U& = 1 To RESCUPRO.LINDICE.ListCount
        RESCUPRO.LINDICE.ListIndex = U& - 1
        REG& = XVALO(Mid$(RESCUPRO.LINDICE.TEXT, 7, 8))
        KKI% = XVALO(Mid$(RESCUPRO.LINDICE.TEXT, 15, 4))
        If REG& > 0 Then
          If REG& <= MAXDEU& Then
            If KKI% > 0 Then
              If KKI% <= 6 Then
                x$ = REGLEIDO$("CUECORR", REG&)
                TIX$ = Mid$(x$, 7, 2)
                SG% = (-1)
                TIPOS$ = "Factura": If TIX$ = "AS" Then TIPOS$ = "Ajuste de Saldo"
                                    If TIX$ = "TF" Then TIPOS$ = "Ticket-Factura"
                                    If TIX$ = "RH" Then TIPOS$ = "Rec.Honorarios"
                                    If TIX$ = "ND" Then TIPOS$ = "Nota de Debito"
                                    If TIX$ = "CP" Then TIPOS$ = "Crédito Propio"
                                    If TIX$ = "CO" Then TIPOS$ = "Comprobante"
                                    If TIX$ = "OP" Then TIPOS$ = "Orden de Pago": SG% = 1
                                    If TIX$ = "FF" Then TIPOS$ = "O/P Fondo Fijo": SG% = 1
                                    If TIX$ = "NC" Then TIPOS$ = "Nota de Credito": SG% = 1
                                    If TIX$ = "DP" Then TIPOS$ = "Debito Propio": SG% = 1
                                    If TIX$ = "CC" Then TIPOS$ = "Comp.Cuenta": SG% = 1
                                    '
                IMPO# = CVD(Mid$(x$, 115, 8)) - CVD(Mid$(x$, 123, 8))
                SALDO# = CVD(Mid$(x$, 139 + 10 * KKI% - 8, 8))
                '
                If TIX$ = "FC" Or TIX$ = "ND" Or TIX$ = "CP" Or TIX$ = "TF" Or TIX$ = "RH" Or TIX$ = "CO" Or (TIX$ = "AS" And IMPO# < 0) Or (TIX$ = "OP" And Mid$(x$, 12, 4) = "anul") Then
                  If SALDO# >= 0.005 Then
                    '
                    SUMAFAC# = SUMAFAC# + IMPO#
                    IMPO# = Abs(IMPO#)
                    SUMADEU# = SUMADEU# + SALDO#
                    '
                    NROX$ = Mid$(x$, 7, 18)
                    FEMI = CVI(Mid$(x$, 1, 2))
                    FEVE = CVI(Mid$(x$, 139 + 10 * KKI% - 10, 2))
                    '
                    IMPTEX$ = CSTRING$(MKD$(IMPO#), 3, 13, 2, 2)
                    SALTEX$ = CSTRING$(MKD$(SALDO#), 3, 13, 2, 2)
                    REGTEX$ = CSTRING$(MKS$(REG&), 3, 8, 0, 2)
                    IRETEX$ = CSTRING$(MKI$(KKI%), 1, 2, 0, 2)
                    '
                    z$ = Space$(128)
                    Call REPLA(z$, NROX$, 1, 18)
                    Call REPLA(z$, FECHATEX$(FEMI), 23, 8)
                    Call REPLA(z$, IMPTEX$, 32, 13)
                    Call REPLA(z$, FECHATEX$(FEVE), 47, 8)
                    Call REPLA(z$, SALTEX$, 56, 13)
                    Call REPLA(z$, REGTEX$, 119, 8)
                    Call REPLA(z$, IRETEX$, 127, 2)
                    '
                    FORPASTA.List2.AddItem z$
                    '
                  End If
                End If
              End If
            End If
          End If
        End If
    Next U&
    '
200 FORPASTA.List2.ListIndex = -1
    FORPASTA.List2.Refresh
    FORPASTA.Label8.Caption = CSTRING$(MKD$(SUMADEU#), 4, 14, 2, 2)
    FORPASTA.Label8.Refresh
    FORPASTA.Label10.Caption = ""
    FORPASTA.Label10.Refresh
    FORPASTA.List2.Visible = True
    FORPASTA.List2.SetFocus
    If FORPASTA.List2.ListCount > 0 Then
        VACON% = 1
        FORPASTA.List2.ListIndex = 0
        VACON% = 0
    End If
    '
    Screen.MousePointer = 11
    FORPASTA.List1.Clear
    '
    RESCUPRO.LINDICE.Clear
    If TRIM$(CONTROL$("ORPASTA", "UTICHECA")) = "SI" Then
       FIN& = ULTREG&("BADACHE")
       For I& = 1 To FIN&
         Call TRACE(20, I&, FIN&)
         x$ = REGLEIDO$("BADACHE", I&)
         If CVI(Mid$(x$, 133, 2)) < 1 Then
           FF% = CVI(Mid$(x$, 59, 2))
           ICHE# = CVD(Mid$(x$, 49, 8))
           If MODOSORVAL% = 0 Then
               z$ = CSTRING$(MKI$(FF%), 1, 6, 0, 2) + CSTRING$(MKS$(I&), 4, 8, 0, 2) + CSTRING$(MKI$(KKI%), 1, 4, 0, 2)
             Else
               z$ = CSTRING$(MKD$(ICHE#), 4, 6, 0, 2) + CSTRING$(MKS$(I&), 4, 8, 0, 2) + CSTRING$(MKI$(KKI%), 1, 4, 0, 2)
           End If
           RESCUPRO.LINDICE.AddItem z$
         End If
       Next I&
    End If
    '
    For U& = 1 To RESCUPRO.LINDICE.ListCount
      RESCUPRO.LINDICE.ListIndex = U& - 1
      REG& = XVALO(Mid$(RESCUPRO.LINDICE.TEXT, 7, 8))
      x$ = REGLEIDO$("BADACHE", REG&)
      If CVI(Mid$(x$, 133, 2)) < 1 Then
         ICHE# = CVD(Mid$(x$, 49, 8))
         TIVA$ = TRIM$(Left$(x$, 18))
         NUCHE$ = TRIM$(Mid$(x$, 19, 14))
         z$ = Space$(128)
         Call REPLA(z$, "Ch.3", 1, 4)
         If Left$(UCase$(TIVA$), 3) = "DOC" Then
            Call REPLA(z$, "Doc3", 1, 4)
         End If
         Call REPLA(z$, TIVA$, 7, 18)
         Call REPLA(z$, NUCHE$, 25, 14)
         FEMI = CVI(Mid$(x$, 57, 2))
         Call REPLA(z$, FECHATEX$(FEMI), 39, 8)
         FEVE = CVI(Mid$(x$, 59, 2))
         Call REPLA(z$, FECHATEX$(FEVE), 49, 8)
         Call REPLA(z$, CSTRING$(MKD$(ICHE#), 3, 12, 2, 2), 57, 12)
         Call REPLA(z$, CSTRING$(MKS$(REG&), 3, 8, 0, 2), 121, 8)
         FORPASTA.List1.AddItem z$
      End If
    Next U&
    '
    FORPASTA.List1.Refresh
    Screen.MousePointer = 1
    '
    If TICUEPRO%((-1) * NC%) = 1 Then        ' GASTOS
       Call SIGUEORPA5
    End If
    '
End Sub
'

Sub SIGUEORPA3()
    '
    If FORPASTA.List2.ListCount < 1 Then
        Exit Sub
    End If
    '
    Call BLANARCH("!CARGAPLI")
    For U& = 1 To FORPASTA.List2.ListCount
        FORPASTA.List2.ListIndex = U& - 1
        z$ = FORPASTA.List2.TEXT
        NROX$ = Mid$(z$, 1, 20)
        FEMII% = FECHANUM(Mid$(z$, 23, 8))
        IPED# = CDbl(XVALO(Mid$(z$, 32, 13)))
        FEVEI% = FECHANUM(Mid$(z$, 47, 8))
        SALD# = CDbl(XVALO(Mid$(z$, 56, 13)))
        ACRE# = CDbl(XVALO(Mid$(z$, 70, 14)))
        DOCU$ = Mid$(z$, 1, 18)
        REG2& = XVALO(Mid$(z$, 119, 8))
        IRE% = XVALO(Mid$(z$, 127, 2))
        '
        y$ = REGBLAN$("!CARGAPLI")
        Call REPLA(y$, NROX$, 1, 20)
        Call REPLA(y$, MKI$(FEMII%), 21, 2)
        Call REPLA(y$, MKD$(IPED#), 23, 8)
        Call REPLA(y$, MKI$(FEVEI%), 31, 2)
        Call REPLA(y$, MKD$(SALD#), 33, 8)
        Call REPLA(y$, MKD$(ACRE#), 41, 8)
        Call REPLA(y$, DOCU$, 49, 18)
        Call REPLA(y$, MKS$(REG2&), 67, 4)
        Call REPLA(y$, MKI$(IRE%), 71, 2)
        '
        Call REGAPP("!CARGAPLI", y$)
    Next U&
    Call CIERRARCH("!CARGAPLI")
    '
200 ATRI$ = "/NC"
    ATRI$ = ATRI$ + "/TITUPAN=Aplicación de Pagos"
    ATRI$ = ATRI$ + "/BORDESUP=W" + TRIM$(Str$(FORPASTA.Top + FORPASTA.Picture2.Top))
    ATRI$ = ATRI$ + "/BORDEIZQ=W" + TRIM$(Str$(FORPASTA.Left + FORPASTA.Picture2.Left))
    '
    VTB% = VENTABU%("CARGAPLI" + ATRI$)
    '
    VACON% = 1
    If VTB% >= 0 Then
        FORPASTA.List2.Clear
        For U& = 1 To ULTREG&("!CARGAPLI")
            y$ = REGLEIDO$("!CARGAPLI", U&)
            NROX$ = Mid$(y$, 1, 20)
            If TRIM$(NROX$) = "Anticipo no Aplicado" Then
               HOI% = HOY(HO$)
               ACRE# = CVD(Mid$(y$, 41, 8))
               Call REPLA(y$, MKI$(HOI%), 31, 2)
               Call REPLA(y$, MKD$(ACRE#), 23, 8)
               Call REPLA(y$, MKD$(ACRE#), 33, 8)
               Call GRAREG("!CARGAPLI", y$, U&)
            End If
            FEMI = CVI(Mid$(y$, 21, 2))
            IPED# = CVD(Mid$(y$, 23, 8))
            IMPTEX$ = CSTRING$(MKD$(IPED#), 4, 13, 2, 2)
            FEVE = CVI(Mid$(y$, 31, 2))
            SALD# = CVD(Mid$(y$, 33, 8))
            SALTEX$ = CSTRING$(MKD$(SALD#), 4, 13, 2, 2)
            ACRE# = CVD(Mid$(y$, 41, 8))
            '
            If ACRE# < 0 Then
                Call MENSERR(24, "No se Permite Aplicación Negativa")
                GoTo 200
            End If
            If U& > 1 Then
              If ACRE# > SALD# Then
                Call MENSERR(24, "El Importe Acreditado Debe Ser\Inferior al Saldo del Documento")
                 GoTo 200
              End If
            End If
            ACRTEX$ = CSTRING$(MKD$(ACRE#), 4, 14, 2, 2)
            DOCU$ = Mid$(y$, 49, 18)
            REG2& = CVS(Mid$(y$, 67, 4))
            IRE% = CVI(Mid$(y$, 71, 2))
            REGTEX$ = CSTRING$(MKS$(REG2&), 3, 8, 0, 2)
            IRETEX$ = CSTRING$(MKI$(IRE%), 1, 2, 0, 2)
            '
            z$ = Space$(128)
            Call REPLA(z$, NROX$, 1, 20)
            Call REPLA(z$, FECHATEX$(FEMI), 23, 8)
            Call REPLA(z$, IMPTEX$, 32, 13)
            Call REPLA(z$, FECHATEX$(FEVE), 47, 8)
            Call REPLA(z$, SALTEX$, 56, 13)
            Call REPLA(z$, ACRTEX$, 70, 14)
            Call REPLA(z$, REGTEX$, 119, 8)
            Call REPLA(z$, IRETEX$, 127, 2)
            FORPASTA.List2.AddItem z$
            '
        Next U&
    End If
    '
    SUMACRE# = 0
    For U& = 1 To FORPASTA.List2.ListCount
        FORPASTA.List2.ListIndex = U& - 1
        z$ = FORPASTA.List2.TEXT
        SUMACRE# = SUMACRE# + CDbl(XVALO(Mid$(z$, 70, 14)))
    Next U&
    '
    FORPASTA.Label10.Caption = CSTRING$(MKD$(SUMACRE#), 4, 14, 2, 2)
    FORPASTA.Label10.Refresh
    FORPASTA.Label18.Caption = TRIM$(CSTRING$(MKD$(SUMACRE#), 4, 14, 2, 2))
    FORPASTA.Label18.Refresh
    FORPASTA.Label20.Caption = CSTRING$(MKD$(XVALO(FORPASTA.Label18.Caption) - XVALO(FORPASTA.Label14.Caption) - XVALO(FORPASTA.Label21.Caption)), 4, 14, 2, 2)
    FORPASTA.Label20.Refresh
    '
    FORPASTA.List2.Refresh
    VACON% = 1
    '
End Sub
'

Sub SIGUEORPA4()
   '
   VACON% = 1
   FORPASTA.List4.Clear
10 For U& = 1 To FORPASTA.List1.ListCount
     z$ = FORPASTA.List1.List(U& - 1)
     If XVALO(Mid$(z$, 70, 14)) >= 0.005 Then
       FORPASTA.List4.AddItem z$
     End If
   Next U&
   FORPASTA.List1.Clear
   For U& = 1 To FORPASTA.List4.ListCount
     FORPASTA.List1.AddItem FORPASTA.List4.List(U& - 1)
   Next U&
   FORPASTA.List1.Refresh
   '
   SUMAVAL# = XVALO(FORPASTA.Label21.Caption)
   For U& = 1 To ULTREG&("!CARMEPA")
      x$ = REGLEIDO$("!CARMEPA", U&)
      ICHE# = CVD(Mid$(x$, 31, 8))
      COVA% = CVI(Left$(x$, 2))
      TIVA$ = ECOARCH$("MEPAGO", MKI$(COVA%))
      NUCHE$ = TRIM$(Mid$(x$, 19, 12))
      '
      z$ = Space$(128)
      Call REPLA(z$, "Prop", 1, 4)
      Call REPLA(z$, TIVA$, 7, 18)
      Call REPLA(z$, NUCHE$, 25, 14)
      FEMI = CVI(Mid$(x$, 39, 2))
      Call REPLA(z$, FECHATEX$(FEMI), 39, 8)
      FEVE = CVI(Mid$(x$, 41, 2))
      Call REPLA(z$, FECHATEX$(FEVE), 49, 8)
      Call REPLA(z$, CSTRING$(MKD$(ICHE#), 3, 12, 2, 2), 57, 12)
      Call REPLA(z$, CSTRING$(MKD$(ICHE#), 3, 14, 2, 2), 70, 14)
      Call REPLA(z$, CSTRING$(MKI$(COVA%), 1, 6, 0, 2), 115, 6)
      FORPASTA.List1.AddItem z$
      SUMAVAL# = SUMAVAL# + ICHE#
      FORPASTA.Label21.Caption = CSTRING$(MKD$(SUMAVAL#), 4, 14, 2, 2)
    '
   Next U&
   '
   FORPASTA.List1.Refresh
   FORPASTA.Label21.Refresh
   FORPASTA.Label20.Caption = CSTRING$(MKD$(XVALO(FORPASTA.Label18.Caption) - XVALO(FORPASTA.Label14.Caption) - SUMAVAL#), 4, 14, 2, 2)
   FORPASTA.Label20.Refresh
   '
   U& = 0
20 If U& < FORPASTA.List2.ListCount Then
      FORPASTA.List2.ListIndex = U&
      z$ = FORPASTA.List2.TEXT
      If XVALO(Mid$(z$, 70, 14)) < 0.005 Then
           FORPASTA.List2.RemoveItem U&
         Else
           U& = U& + 1
      End If
      GoTo 20
   End If
   '
   VACON% = 0
   '
   FORPASTA.Command5.Enabled = True
   '
End Sub
'

Sub SIGUEORPA5()
    '
200 ATRI$ = "/NC"
    ATRI$ = ATRI$ + "/TITUPAN=Distribución de Pagos"
    ATRI$ = ATRI$ + "/BORDESUP=W" + TRIM$(Str$(FORPASTA.Top + FORPASTA.Picture2.Top))
    ATRI$ = ATRI$ + "/BORDEIZQ=W" + TRIM$(Str$(FORPASTA.Left + FORPASTA.Picture2.Left))
    '
305 VTB% = VENTABU%("CARGAST" + ATRI$)
    If VTB% < 0 Then
       If COMALTER%("Realmente Desea Cancelar\la Distribución Ingresada ?\\No, Continuar\Si, Cancelar") <> 2 Then
           GoTo 305
       End If
       Call CIERRARCH("!CARGAST")
       Call BLANARCH("!CARGAST")
    End If
    '
    SUMACRE# = 0
    V& = 0: UF& = ULTREG&("!CARGAST")
    For U& = 1 To UF&
       x$ = REGLEIDO$("!CARGAST", U&)
       If ECOARCH$("ECUECON", Left$(x$, 12)) <> "" Then
       CUE% = CVI(Mid$(REGACT$("ECUECON"), 41, 2))
       Call REPLA(x$, MKI$(CUE%), 21, 2)
       SUMACRE# = SUMACRE# + CVD(Mid$(x$, 13, 8))
       V& = V& + 1
       Call GRAREG("!CARGAST", x$, V&)
       End If
    Next U&
    '
    Call SETULTREG("!CARGAST", V&)
    If V& < UF& Then
       Call COMUNI("Se han Detectado Importes Mal Aplicados.\  \Revise y Confirme la Distribución de Pagos.")
       GoTo 305
    End If
    '
    FORPASTA.Label18.Caption = CSTRING$(MKD$(SUMACRE#), 4, 14, 2, 2)
    FORPASTA.Label18.Refresh
    FORPASTA.Label20.Caption = CSTRING$(MKD$(XVALO(FORPASTA.Label18.Caption) - XVALO(FORPASTA.Label14.Caption) - XVALO(FORPASTA.Label21.Caption)), 4, 14, 2, 2)
    FORPASTA.Label20.Refresh
    '
    FORPASTA.List2.ListIndex = 0
    z$ = FORPASTA.List2.TEXT
    IACRE$ = CSTRING$(MKD$(SUMACRE#), 4, 14, 2, 2)
    Call REPLA(z$, IACRE$, 70, 14)
    VACON% = 1
    FORPASTA.List2.AddItem z$, ILISTA&
    FORPASTA.List2.RemoveItem ILISTA& + 1
    FORPASTA.List2.Refresh
    FORPASTA.List2.ListIndex = 0
    VACON% = 0
    '
End Sub
'

Sub PRIORPASTA()
    '
    HOI% = HOY(HO$)
    '
    FORIPRE$ = TRIM$(CONTROL$("", "FORPASTA"))
    If FORIPRE$ = "" Then
        Call COMUNI("Error de Configuración:\  \Falta Definir Formulario\de Impresión de Orden de Pago.")
        Exit Sub
    End If
    '
    If FORIPRE$ <> "" Then
        LU$ = LUDAT$
        If EXISTE%(LU$ + FORIPRE$ + ".FRM") <> 1 Then
            RMCC$ = REPLACAR$(LU$ + FORIPRE$ + ".FRM", "\", "/")
            Call COMUNI("Error de Configuración:\  \No Existe Formulario de Impresión\de Recibo de Cobranza:\  \'" + RMCC$ + "'.")
            Exit Sub
        End If
    End If
    '
    CODFOR$ = FORIPRE$
    '
    CONTA% = FORPASTA.List1.ListCount
    If CONTA% < 1 Then
        Call COMUNI("No se ha Ingresado Detalle de Valores")
        Exit Sub
    End If
    '
    If XVALO(FORPASTA.Label14.Caption) + XVALO(FORPASTA.Label21.Caption) < 0.005 Then
        Call COMUNI("No se ha Ingresado Detalle de Valores")
        Exit Sub
    End If
    '
    AVAHOJA$ = TRIM$(CONTROL$("ORPASTA", "AVAHOJA"))
    Screen.MousePointer = 11
    Call BLANARCH("!PRINRECI")
    '
    ILISTA& = FORPASTA.List1.ListIndex
    TOLISTA& = FORPASTA.List1.TopIndex
    '
    TCUIT$ = "": MODOSABO% = 0
    If InStr(UCase$(EMPREAC$), "SABO") > 0 Then MODOSABO% = 1
    '
    For U& = 1 To CONTA%
        FORPASTA.List1.ListIndex = U& - 1
        x$ = FORPASTA.List1.TEXT
        If TRIM$(x$) = "" Then GoTo 19
        TIPOS$ = Left$(x$, 4)
        BANCO$ = Mid$(x$, 7, 18)
        NUVA$ = Mid$(x$, 25, 14)
        IMVA# = XVALO(Mid$(x$, 70, 14))
        If IMVA# < 0.005 Then GoTo 19
        '
        FEMII% = FECHANUM(Mid$(x$, 39, 8)): If FEMII% < 1 Then FEMII% = HOI%
        FEVEI% = FECHANUM(Mid$(x$, 49, 8)): If FEVEI% < 1 Then FEMII% = HOI%
        PROTER$ = ""
        CLEA% = 48
        '
        BBAN$ = TRIM$(UCase$(BANCO$))
        If BBAN$ <> "" Then
          If InStr(TCUIT$, "CUIT " + BBAN$) < 1 Then
            CUBAN$ = CUITBAN$(BBAN$)
            If CUBAN$ <> "" Then
              If TCUIT$ <> "" Then TCUIT$ = TCUIT$ + ", "
              TCUIT$ = TCUIT$ + "CUIT " + BBAN$ + ": " + CUBAN$
            End If
          End If
        End If
        '
10      z$ = REGBLAN$("!PRINRECI")
        Call REPLA(z$, TIPOS$, 1, 4)
        Call REPLA(z$, NUVA$, 5, 12)
        Call REPLA(z$, BANCO$, 17, 16)
        Call REPLA(z$, MKD$(IMVA#), 33, 8)
        Call REPLA(z$, MKI$(0), 51, 2)
        FEX = FEVEI%
        If FEX > 0 Then
            Call REPLA(z$, FECHATEX$(FEX), 41, 8)
        End If
        Call REGAPP("!PRINRECI", z$)
        '
19  Next U&
    '
    If XVALO(FORPASTA.Label14) >= 0.005 Then
       z$ = REGBLAN$("!PRINRECI")
       Call REPLA(z$, "R.IB", 1, 4)
       Call REPLA(z$, TRIM$(FORPASTA.Label25), 5, 12)
       'Call REPLA(Z$, "Ret.IB.Bs.As.", 17, 16)
       Call REPLA(z$, "Ret.I.Brutos.", 17, 16)
       Call REPLA(z$, MKD$(XVALO(FORPASTA.Label14)), 33, 8)
       FEX = HOI%
       Call REPLA(z$, FECHATEX$(FEX), 41, 8)
       Call REGAPP("!PRINRECI", z$)
    End If
    '
    V0& = ULTREG&("!PRINRECI")
    '
    ILISTA& = FORPASTA.List2.ListIndex
    TOLISTA& = FORPASTA.List2.TopIndex
    '
    V& = 0
    If TICUEPRO%((-1) * NC%) = 1 Then
       For U& = 1 To ULTREG&("!CARGAST")
          x$ = REGLEIDO$("!CARGAST", U&)
          DOCU$ = Left$(x$, 12)
          If MODOSABO% = 1 Then
            If InStr(UCase$(DOCU$), "ANTICIPO NO") > 0 Then
              DOCU$ = "a/c Facturas"
            End If
          End If
          SALPE# = 0
          CANCE# = CVD(Mid$(x$, 13, 8))
          If Abs(CANCE#) >= 0.005 Then
             V& = V& + 1
             If V& <= V0& Then
                  z$ = REGLEIDO$("!PRINRECI", V&)
                Else
                  z$ = REGBLAN$("!PRINRECI")
             End If
             Call REPLA(z$, DOCU$, 53, 18)
             Call REPLA(z$, MKD$(CANCE#), 71, 8)
             Call REPLA(z$, MKD$(SALPE#), 79, 8)
             Call GRAREG("!PRINRECI", z$, V&)
          End If
       Next U&
       Call CIERRARCH("!CARGAST")
       GoTo 25
    End If
    '
    For U& = 1 To FORPASTA.List2.ListCount
        FORPASTA.List2.ListIndex = U& - 1
        x$ = FORPASTA.List2.TEXT
        DOCU$ = Mid$(x$, 1, 20)
        If MODOSABO% = 1 Then
          If InStr(UCase$(DOCU$), "ANTICIPO NO") > 0 Then
            DOCU$ = "a/c Facturas"
          End If
        End If
        SALPE# = CDbl(XVALO(Mid$(x$, 56, 13)))
        CANCE# = CDbl(XVALO(Mid$(x$, 70, 14)))
        If Abs(CANCE#) >= 0.005 Then
            V& = V& + 1
            If V& <= V0& Then
                  z$ = REGLEIDO$("!PRINRECI", V&)
               Else
                  z$ = REGBLAN$("!PRINRECI")
            End If
            FEFACA% = FECHANUM(Mid$(x$, 23, 8))
            Call REPLA(z$, MKI$(FEFACA%), 51, 2)
            Call REPLA(z$, DOCU$, 53, 18)
            Call REPLA(z$, MKD$(CANCE#), 71, 8)
            Call REPLA(z$, MKD$(SALPE#), 79, 8)
            Call GRAREG("!PRINRECI", z$, V&)
        End If
    Next U&
    '
25  V0& = ULTREG&("!PRINRECI")
    '
    TOTAVAL# = CDbl(XVALO(FORPASTA.Label14.Caption) + XVALO(FORPASTA.Label21.Caption))
    TOTACAN# = CDbl(XVALO(FORPASTA.Label18.Caption)) + CDbl(XVALO(FORPASTA.Label20.Caption))
33  ITEXTO$ = NUMTEX$(TOTAVAL#) + " ctvs"
    LOTEX% = XVALO(ATRIFORM$(CODFOR$, "LOIMPLET")): If LOTEX% < 20 Then LOTEX% = 20
    Call FRATEX(ITEXTO$, NRX1$, NRX2$, NRX3$, LOTEX%)
    IMPLET1$ = "": IMPLET2$ = "": IMPLET3$ = ""
    If TRIM$(NRX1$) <> "" Then
       IMPLET1$ = TRIM$(NRX1$) + " " + "***************************************************************************************************************"
    End If
    If TRIM$(NRX2$) <> "" Then
       IMPLET2$ = TRIM$(NRX2$) + " " + "***************************************************************************************************************"
    End If
    If TRIM$(NRX3$) <> "" Then
       IMPLET3$ = TRIM$(NRX3$) + " " + "***************************************************************************************************************"
    End If
    '
    Call CIERRARCH("!PRINRECI")
    '
30  If TRIM$(FORIPRE$) <> "" Then
        '
        CANCELPRINT% = 0
        If CONTROL$("ORPASTA", "SUPRINT") = "SI" Then CANCELPRINT% = 1
        '
        CODFOR$ = TRIM$(UCase$(FORIPRE$))
        MARIZW = XVALO(ATRIFORM$(CODFOR$, "MARIZQMM"))
        MARSUW = XVALO(ATRIFORM$(CODFOR$, "MARSUPMM"))
        '
        INTERLI = XVALO(ATRIFORM$(CODFOR$, "INTERLIN"))
        If INTERLI < 1 Then
            Call COMUNI("Error de Configuración de Interlinea\en Formulario '" + CODFOR$ + ".FRM'.")
            INTERLI = 1
        End If
        '
        DESMM = XVALO(ATRIFORM$(CODFOR$, "DETDESMM"))
        HASMM = XVALO(ATRIFORM$(CODFOR$, "DETHASMM"))
        '
        CACOPI% = XVALO(ATRIFORM$(CODFOR$, "CAN-COP"))
        If CACOPI% < 1 Or CACOPI% > 6 Then
            Call COMUNI("Error de Configuración de Cantidad de Copias\en Formulario '" + CODFOR$ + ".FRM'.")
            CACOPI% = 1
        End If
        '
        ORIENTA% = 1
        PORI$ = TRIM$(UCase$(ATRIFORM$(CODFOR$, "ORIENTA")))
        If PORI$ = "PAISAJE" Or PORI$ = "2" Then
            ORIENTA% = 2
        End If
        '
        PORIAN = Printer.Orientation  ' retiene orientacion
        '
        NRORPA$ = TRIM$(Mid$(FORPASTA.Text11.TEXT, 4))
        '
        Screen.MousePointer = 11
        If AVAHOJA$ <> "NEWPAGE" Then
            'Printer.Orientation = 1
            Call SETSPOOL("ORIENTATION", "1")
            If ORIENTA% = 2 Then
                'Printer.Orientation = 2
                Call SETSPOOL("ORIENTATION", "2")
            End If
        End If
        '
        NC1% = (-1) * NC%
        For COP% = 1 To CACOPI%
            '
            If AVAHOJA$ = "NEWPAGE" Then
               If PRINTERCONTROL% = 1 Then
                  'Printer.NewPage
                  Call SETSPOOL("NEWPAGE", "")
               End If
            End If
            PRINTERCONTROL% = 1
            '
            SUMAVA# = 0: SUMACAN# = 0
            Call CARFORWIN
            '
            Call PRINTFORWIN(TRIM$(Str$(NC%)), "CUEN-CLI", 0, 0)
            Call PRINTFORWIN(TRIM$(RASOCLI$(NC1%)), "RASO-CLI", 0, 0)
            Call PRINTFORWIN(DOMICLI$(NC1%), "DOMI-CLI", 0, 0)
            Call PRINTFORWIN(CPOSCLI$(NC1%), "CPOS-CLI", 0, 0)
            Call PRINTFORWIN(LOCACLI$(NC1%), "LOCA-CLI", 0, 0)
            Call PRINTFORWIN(CUITCLI$(NC1%), "CUIT-CLI", 0, 0)
            Call PRINTFORWIN(POSIVAL$(NC1%), "PIVA-CLI", 0, 0)
            FF% = FECHANUM(TRIM$(FORPASTA.Text9.TEXT))
            Call PRINTFORWIN(FETEXTO$(FF%), "FECH-EMI", 0, 0)
            Call PRINTFORWIN(NRORPA$, "NUME-COM", 0, 0)
            '
            Call PRINTFORWIN(TCUIT$, "TOTA-ITS", 0, 0)
            '
            YAC = 0
            For U& = 1 To ULTREG&("!OBSERVOF")
               TXOB$ = TRIM$(REGLEIDO$("!OBSERVOF", U&))
               Call PRINTFORWIN(TXOB$, "OBSERVA", 0, YAC)
               YAC = YAC + INTERLI
            Next U&
            Call CIERRARCH("!OBSERVOF")
            '
            YAC = DESMM
            For YX& = 1 To ULTREG&("!PRINRECI")
               '
               YAC = YAC + INTERLI
               '
               If YAC > HASMM - 2 * INTERLI Then
                   YAC = YAC + INTERLI
                   Call PRINTFORWIN("Transporte", "BANC-CHE", 0, YAC)
                   Call PRINTFORWIN(CSTRING$(MKD$(SUMAVA#), 4, 12, 2, 2), "IMPO-CHE", 0, YAC)
                   Call PRINTFORWIN("Transporte", "DOCU-APL", 0, YAC)
                   Call PRINTFORWIN(CSTRING$(MKD$(SUMACAN#), 4, 12, 2, 2), "IMPO-CAN", 0, YAC)
                   '
                   Call PRINTFORWIN(String$(20, "*"), "TOTA-VAL", 0, 0)
                   Call PRINTFORWIN(String$(20, "*"), "TOTA-CAN", 0, 0)
                   Call PRINTFORWIN(String$(96, "*"), "IMP-LET1", 0, 0)
                   Call PRINTFORWIN(String$(96, "*"), "IMP-LET2", 0, 0)
                   '
                   If AVAHOJA$ = "NEWPAGE" Then
                        'Printer.NewPage
                        Call SETSPOOL("NEWPAGE", "")
                      Else
                       'Printer.EndDoc
                       Call SETSPOOL("ENDDOC", "")
                   End If
                   '
                   Call CARFORWIN
                   Call PRINTFORWIN(TRIM$(Str$(NC1%)), "CUEN-CLI", 0, 0)
                   Call PRINTFORWIN(TRIM$(RASOCLI$(NC1%)), "RASO-CLI", 0, 0)
                   Call PRINTFORWIN(DOMICLI$(NC1%), "DOMI-CLI", 0, 0)
                   Call PRINTFORWIN(CPOSCLI$(NC1%), "CPOS-CLI", 0, 0)
                   Call PRINTFORWIN(LOCACLI$(NC1%), "LOCA-CLI", 0, 0)
                   Call PRINTFORWIN(CUITCLI$(NC1%), "CUIT-CLI", 0, 0)
                   Call PRINTFORWIN(POSIVAL$(NC1%), "PIVA-CLI", 0, 0)
                   FEX = FF%
                   Call PRINTFORWIN(FETEXTO$(FF%), "FECH-EMI", 0, 0)
                   Call PRINTFORWIN(NRORPA$, "NUME-COM", 0, 0)
                   '
                   Call PRINTFORWIN(TCUIT$, "TOTA-ITS", 0, 0)
                   '
                   YAC = DESMM + INTERLI
                   Call PRINTFORWIN("Transporte", "BANC-CHE", 0, YAC)
                   Call PRINTFORWIN(CSTRING$(MKD$(SUMAVA#), 4, 12, 2, 2), "IMPO-CHE", 0, YAC)
                   Call PRINTFORWIN("Transporte", "DOCU-APL", 0, YAC)
                   Call PRINTFORWIN(CSTRING$(MKD$(SUMACAN#), 4, 12, 2, 2), "IMPO-CAN", 0, YAC)
                   YAC = YAC + 2 * INTERLI
                   '
                End If
                '
                YY$ = REGLEIDO$("!PRINRECI", YX&)
                Call PRINTFORWIN(Left$(YY$, 4), "TIPO-VAL", 0, YAC)
                Call PRINTFORWIN(Mid$(YY$, 5, 12), "NUME-CHE", 0, YAC)
                Call PRINTFORWIN(Mid$(YY$, 17, 16), "BANC-CHE", 0, YAC)
                Call PRINTFORWIN(CSTRING$(Mid$(YY$, 33, 8), 4, 12, 2, 2), "IMPO-CHE", 0, YAC)
                SUMAVA# = SUMAVA# + CVD(Mid$(YY$, 33, 8))
                Call PRINTFORWIN(CSTRING$(Mid$(YY$, 41, 8), 6, 8, 0, 1), "FECH-CHE", 0, YAC)
                  FEXUU = CVI(Mid$(YY$, 51, 2))
                If FEXUU > 0 Then Call PRINTFORWIN(FECHATEX$(FEXUU), "FECH-LAN", 0, YAC)
                Call PRINTFORWIN(Mid$(YY$, 53, 18), "DOCU-APL", 0, YAC)
                Call PRINTFORWIN(CSTRING$(Mid$(YY$, 71, 8), 4, 12, 2, 2), "IMPO-CAN", 0, YAC)
                SUMACAN# = SUMACAN# + CVD(Mid$(YY$, 71, 8))
                Call PRINTFORWIN(CSTRING$(Mid$(YY$, 79, 8), 4, 12, 2, 2), "IMPO-PEN", 0, YAC)
                '
            Next YX&
            '
            Call PRINTFORWIN(CSTRING$(MKD$(TOTAVAL#), 4, 12, 2, 2), "TOTA-VAL", 0, 0)
            Call PRINTFORWIN(CSTRING$(MKD$(TOTACAN#), 4, 12, 2, 2), "TOTA-CAN", 0, 0)
            Call PRINTFORWIN(IMPLET1$, "IMP-LET1", 0, 0)
            Call PRINTFORWIN(IMPLET2$, "IMP-LET2", 0, 0)
                   '
            If AVAHOJA$ <> "NEWPAGE" Then
                 'Printer.EndDoc
                 Call SETSPOOL("ENDDOC", "")
            End If
            '
        Next COP%
        '
        Call CIERRARCH("*.*")
        DESCRIDOC$ = "Orden de Pago Nro. " + TRIM$(NRORPA$) + " - " + TRIM$(RASOCLI$(NC1%))
        FEX = FF%
        DESCRIDOC$ = DESCRIDOC$ + " - " + FECHATEX$(FEX)
        '
        Call SAVESPOOL(DESCRIDOC$, OKX%)
        Screen.MousePointer = 1
        If COMALTER%("Orden de Pago Numero\'" + NRORPA$ + "'\Emitida y Registrada\\Repetir Impresion\Otra Orden de Pago") = 1 Then GoTo 30
        '
        If AVAHOJA$ <> "NEWPAGE" Then
            Printer.Orientation = PORIAN
        End If
        '
    End If
    Call CIERRARCH("*.*")
    '
End Sub

Function CUITBAN$(BBAN$)
   '
   CUIBA$ = ""
   BBAX$ = TRIM$(BBAN$)
   If Len(BBAX$) < 4 Then GoTo 90
   '
   For U& = 1 To ULTREG&("MEPAGO")
     XUX$ = REGLEIDO$("MEPAGO", U&)
     If TRIM$(Mid$(XUX$, 3, 12)) = BBAX$ Then
       CUIBA$ = TRIM$(Mid$(XUX$, 31, 14))
       If Len(CUIBA$) < 4 Then CUIBA$ = ""
       GoTo 90
     End If
   Next U&
   '
   For U& = 1 To ULTREG&("MECOBRA")
     XUX$ = REGLEIDO$("MECOBRA", U&)
     If InStr(Mid$(XUX$, 3, 12), BBAX$) > 0 Then
       CUIBA$ = TRIM$(Mid$(XUX$, 31, 14))
       If Len(CUIBA$) < 11 Then CUIBA$ = ""
       GoTo 90
     End If
   Next U&
   '
90 CUITBAN$ = CUIBA$
   '
End Function
'
Sub GRAORPASTA(OKX%)
    '
    Call CIERRARCH("!OBSERVOF")
    Call BLANARCH("!OBSERVOF")
    '
    NRO& = XVALO(Mid$(FORPASTA.Text11.TEXT, 4))
    VALIDA% = VENTABU%("OBSERVOF/TITUPAN=Orden de Pagos Numero" + Str$(NRO&))
    If VALIDA% < 0 Then
       OKX% = (-1): Exit Sub
    End If
    '
    HP% = HOY(HOS$)
    OKX% = 0
    Call BLOQARCH("CUECORR")
    Call BLOQARCH("PROVED2")
    Call BLOQARCH("DETPAGAS")
    Call BLOQARCH("APLIPAG")
    Call BLOQARCH("BADACHE")
    Call BLOQARCH("CHEQUEMI")
    OKX% = 1                    ' bloqueó archivos
    '
    NC% = XVALO(FORPASTA.Text1.TEXT)
    NC1% = (-1) * NC%
    NRO& = XVALO(Mid$(FORPASTA.Text11.TEXT, 4))
    NROPRO& = NRO&
    Call NUORPA(NROPRO&)
    If NROPRO& <> NRO& Then
       Call CIERRARCH("CUECORR")
       Call CIERRARCH("PROVED2")
       Call CIERRARCH("DETPAGAS")
       Call CIERRARCH("APLIPAG")
       Call CIERRARCH("BADACHE")
       Call CIERRARCH("CHEQUEMI")
       Call MENSERR(24, "Al Grabar esta Orden de Pago\el Número Propuesto Estaba Ocupado.\  \Verifique el Nuevo Número Propuesto\y Vuelva a Grabar e Imprimir.")
       FORPASTA.Text11.TEXT = "OP." + TRIM$(Str$(NROPRO&))
       OKX% = 0
       Exit Sub
    End If
    '
    Tipo% = 4: VA% = 1: MO$ = "$"
    TOFAC#(0) = XVALO(FORPASTA.Label14.Caption) + XVALO(FORPASTA.Label21.Caption)
    TOFAC#(1) = 0
    TOFAC#(2) = TOFAC#(0)
    '
    FC% = FECHANUM(FORPASTA.Text9.TEXT)
    FVII%(1) = FC%
    IMV#(1) = 0
    FORPASTA.List2.ListIndex = 0
    x$ = FORPASTA.List2.TEXT
    DOCU$ = Mid$(x$, 1, 20)
    If DOCU$ = "Anticipo no Aplicado" Then
        CANCE# = CDbl(XVALO(Mid$(x$, 70, 14)))
        If Abs(CANCE#) >= 0.005 Then
           IMV#(1) = (-1) * CANCE#
        End If
    End If
    '
    Call CUECORR(FC%, (-1) * NC%, Tipo%, VA%, SUC%, NRO&, REFE$, MO$, TOFAC#(), FVII%(), IMV#(), GRU%(), IPAR#(), 1, FC%)
    OKX% = 2         ' grabó Orden de Pago
    '
    'If DESCAJ# > 0 Then
    '  TIPO% = 5
    '  TOFAC#(0) = DESCAJ#
    '  TOFAC#(1) = (Int(100 * DESCAJ# / 1.21 + 0.5)) / 100
    '  TOFAC#(2) = DESCAJ# - TOFAC#(1)
    '  FVII%(1) = 0: IMV#(1) = 0
    '
    '  Call CUECORR(FC%, (-1) * NC%, TIPO%, VA%, SUC%, NRO&, REFE$, MO$, TOFAC#(), FVII%(), IMV#(), GRU%(), IPAR#(), 1)
    '
    'End If
     '
800  ' *********************************** aqui blanquear facturas canceladas
     '
     If TICUEPRO%(NC1%) = 0 Then
       '
       XX% = 0
       For U& = 1 To FORPASTA.List2.ListCount
         FORPASTA.List2.ListIndex = U& - 1
         z$ = FORPASTA.List2.TEXT
         DOCU$ = Left$(z$, 20)
         If DOCU$ <> "Anticipo no Aplicado" Then
           CANCE# = CDbl(XVALO(Mid$(z$, 70, 14)))
           If CANCE# > 0 Then
              NREG& = XVALO(TRIM$(Mid$(z$, 119, 8)))
              If NREG& > 0 Then
                If NREG& <= ULTREG&("CUECORR") Then
                  ORD% = XVALO(Mid$(z$, 127, 2))
                  If ORD% > 0 Then
                    If ORD% <= 6 Then
                       '
                       SS$ = REGLEIDO$("CUECORR", NREG&)
                       SALDO$ = Mid$(SS$, 139, 60)
                       SALDEU# = CVD(Mid$(SALDO$, 10 * ORD% - 7, 8))
                       SALDEU# = SALDEU# - CANCE#
                       Call REPLA$(SALDO$, MKD$(SALDEU#), 10 * ORD% - 7, 8)
                       '
                       If Abs(SALDEU#) < 0.005 Then
                         SALDO$ = Left$(SALDO$, 10 * (ORD% - 1)) + Mid$(SALDO$, 10 * ORD% + 1) + String$(10, 0)
                       End If
                       '
                       Call REPLA(SS$, SALDO$, 139, 60)
                       Call GRAREG("CUECORR", SS$, NREG&)
                    End If
                  End If
                End If
              End If
           End If
         End If
       Next U&
       '
     End If
     'Call CIERRARCH("CUECORR")
     OKX% = 3         ' depuró facturas pendientes
     '
820  If TICUEPRO%(NC1%) = 1 Then
        RC& = REGICLI&(NC1%)
        DX$ = REGLEIDO$("PROVED2", RC&)
        DX$ = REPLACE$(DX$, MKD$(0), 109, 8)
        Call GRAREG("PROVED2", DX$, RC&)
        '
        For K1& = 1 To ULTREG&("!CARGAST")
           z$ = REGLEIDO$("!CARGAST", K1&)
           ZX$ = REGBLAN$("DETPAGAS")
           ZX$ = REPLACE$(ZX$, MKI$(NC%), 1, 2)
           ZX$ = REPLACE$(ZX$, MKS$(NRO&), 3, 4)
           ZX$ = REPLACE$(ZX$, MKI$(FC%), 7, 2)
           ZX$ = REPLACE$(ZX$, Mid$(z$, 13, 10), 9, 10)
           AA = CVD(Mid$(z$, 3))
           Call REGAPP("DETPAGAS", ZX$)
        Next K1&
        'Call CIERRARCH("DETPAGAS")
        GoTo 840
     End If


830  ' *************************************** aqui grabar archivo APLIPAG.DAT
     '
       For U& = 1 To FORPASTA.List2.ListCount
         FORPASTA.List2.ListIndex = U& - 1
         z$ = FORPASTA.List2.TEXT
         DOCU$ = Left$(z$, 20)
         If DOCU$ <> "Anticipo no Aplicado" Then
           CANCE# = CDbl(XVALO(Mid$(z$, 70, 14)))
           If CANCE# > 0 Then
              NREG& = XVALO(TRIM$(Mid$(z$, 119, 8)))
              If NREG& > 0 Then
                If NREG& <= ULTREG&("CUECORR") Then
                  SS$ = REGLEIDO$("CUECORR", NREG&)
                  DOCORI$ = Mid$(SS$, 7, 18)
                  SALDO$ = Mid$(SS$, 139, 60)
                  SS77$ = Mid$(SS$, 51)
                  ITOTORI# = Abs(CVD(Mid$(SS$, 115, 8)) - CVD(Mid$(SS$, 123, 8)))
                  IVAORI# = Abs(CVD(Mid$(SS77$, 25, 8)) + CVD(Mid$(SS77$, 33, 8)) + CVD(Mid$(SS77$, 41, 8)))
                  '
                  Call APLIPAG(FC%, NC1%, 4, 1, SUC%, NRO&, CANCE#, RefApli$, DOCORI$, ITOTORI#, IVAORI#)
                  '
                End If
              End If
           End If
         End If
       Next U&
       OKX% = 4       ' grabó aplicaciones
     '
840  'Call CIERRARCH("APLIPAG")
     'Call CIERRARCH("DETPAGAS")
     '
     ' ************************************** aqui grabar archivo CHEQUEMI.DAT
     '
     For U& = 1 To FORPASTA.List1.ListCount
        'FORPASTA.List1.ListIndex = U& - 1
        'Z$ = FORPASTA.List1.TEXT
        z$ = FORPASTA.List1.List(U& - 1)
        TIDOS$ = Left$(z$, 4)
        Tipo% = XVALO(Mid$(z$, 115, 6))
        BANCO$ = Mid$(z$, 7, 18)
        NUCHE$ = Mid$(z$, 25, 14)
        IMCHE# = XVALO(Mid$(z$, 70, 14))
        FECHE% = FECHANUM(Mid$(z$, 39, 8)): If FECHE% < 1 Then FECHE% = HO%
        FEVEN% = FECHANUM(Mid$(z$, 49, 8)): If FEVEN% < FECHE% Then FEVEN% = FECHE%
        '
        If IMCHE# > 0 Then
           '
           If TIDOS$ = "Ch.3" Or TIDOS$ = "Doc3" Then
              Tipo% = 999
              REGBA& = XVALO(TRIM$(Mid$(z$, 121, 8)))
              If REGBA& > 0 Then
                 If REGBA& <= ULTREG&("BADACHE") Then
                    HOII% = HOY(HOS$)
                    TTS$ = REGLEIDO$("BADACHE", REGBA&)
                    Call REPLA(TTS$, RASOCLI$(NC1%), 103, 30)
                    Call REPLA(TTS$, MKI$(FC%), 133, 2)
                    Call REPLA(TTS$, "Orden de Pago Nro." + Str$(NRO&), 135, 40)
                    Call REPLA(TTS$, HOS$ + " - " + Left$(Time$, 5) + " hs - Term: " + NOMTER$, 175, 40)
                    Call REPLA(TTS$, "Usuario: " + TRIM$(USERNAME$), 215, 40)
                    Call GRAREG("BADACHE", TTS$, REGBA&)
                 End If
              End If
           End If
           '
           z$ = REGBLAN$("CHEQUEMI")
           z$ = REPLACE$(z$, MKI$(NC%), 1, 2)
           z$ = REPLACE$(z$, NUCHE$, 3, 12)
           z$ = REPLACE$(z$, BANCO$, 15, 16)
           z$ = REPLACE$(z$, MKI$(FECHE%), 31, 2)
           z$ = REPLACE$(z$, MKD$(IMCHE#), 33, 8)
           z$ = REPLACE$(z$, MKS$(NRO&), 41, 4)
           z$ = REPLACE$(z$, MKI$(Tipo%), 45, 2)
           z$ = REPLACE$(z$, MKI$(FC%), 47, 2)
           z$ = REPLACE$(z$, MKI$(FEVEN%), 49, 2)
           Call REGAPP("CHEQUEMI", z$)
           '
       End If
       '
     Next U&
     '
     ' Ret I.Brutos Buenos Aires
     If XVALO(FORPASTA.Label14) >= 0.005 Then
        NURETBA& = 1
        If ULTREG&("SURETIB") > 0 Then
          Z1$ = REGLEIDO$("SURETIB", ULTREG&("SURETIB"))
          NURETBA& = 1 + CVS(Mid$(Z1$, 3, 4))
        End If
        FORPASTA.Label25 = TRIM$(Str$(NURETBA&))
        '
        Z1$ = REGBLAN$("SURETIB")
        Call REPLA(Z1$, MKI$(FC%), 1, 2)
        Call REPLA(Z1$, MKS$(NURETBA&), 3, 4)
        Call REPLA(Z1$, MKS$(NRO&), 7, 4)
        Call REPLA(Z1$, RASOCLI$(NC1%), 21, 30)
        Call REPLA(Z1$, CUITCLI$(NC1%), 51, 14)
        Call REPLA(Z1$, NIBRCLI$(NC1%), 65, 14)
        Call REPLA(Z1$, MKD$(XVALO(FORPASTA.Label23)), 79, 8)
        Call REPLA(Z1$, MKD$(XVALO(FORPASTA.Label14)), 87, 8)
        Call REGAPP("SURETIB", Z1$)
        '
        z$ = REGBLAN$("CHEQUEMI")
        z$ = REPLACE$(z$, MKI$(NC%), 1, 2)
        NUCHE$ = TRIM$(Str$(NURETBA&))
        z$ = REPLACE$(z$, NUCHE$, 3, 12)
        'Z$ = REPLACE$(Z$, "Ret.IB.Bs.As.", 15, 16)
        z$ = REPLACE$(z$, "Ret.I.Brutos.", 15, 16)
        z$ = REPLACE$(z$, MKI$(FC%), 31, 2)
        z$ = REPLACE$(z$, MKD$(XVALO(FORPASTA.Label14)), 33, 8)
        z$ = REPLACE$(z$, MKS$(NRO&), 41, 4)
        z$ = REPLACE$(z$, MKI$(XVALO(FORPASTA.Label24)), 45, 2)
        z$ = REPLACE$(z$, MKI$(FC%), 47, 2)
        z$ = REPLACE$(z$, MKI$(FC%), 49, 2)
        Call REGAPP("CHEQUEMI", z$)
     End If
     '
     Call CIERRARCH("SURETIB")
     'Call CIERRARCH("CHEQUEMI")
     OKX% = 5          ' grabo cheques de terceros
     OKX% = 6          ' grabo cheques emitidos
     '
890  Call CIERRARCH("PROVED2")
     Call CIERRARCH("DETPAGAS")
     Call CIERRARCH("APLIPAG")
     Call CIERRARCH("BADACHE")
     Call CIERRARCH("CHEQUEMI")
     Call CIERRARCH("CUECORR")
     Call CIERRARCH("*.*")
     Call BAJALDISCO
     OKX% = 7          ' LIBERO ARCHIVOS
     '
End Sub
'

Sub SIGUEFACO1()
    '
    Dim NRO$(10)
    Screen.MousePointer = 11
    '
    NCXI = XVALO(FORFACOM.Text1.TEXT)
    If NCXI < 1 Or NCXI > 32767 Then
        Call MENSERR(24, "Numero de Cuenta no Válido")
        Screen.MousePointer = 1
        FORFACOM.Text1.SetFocus
        Exit Sub
    End If
    '
    NC% = NCXI
    If ECOARCH$("ACREDOR", MKI$(NC%)) = "" Then
        Call MENSERR(24, "Cuenta no Figura\en Base de Datos de Acreedores")
        Screen.MousePointer = 1
        FORFACOM.Text1.SetFocus
        Exit Sub
    End If
    '
    FE$ = TRIM$(FORFACOM.Text9.TEXT)
    FF% = FECHANUM(FE$)
    HOI% = HOY(HO$)
    If FF% < 1 Or FF% > HOI% + 1 Then
        Screen.MousePointer = 1
        Call MENSERR(24, "Fecha no Válida")
        FORFACOM.Text9.SetFocus
        Exit Sub
    End If
    '
    If DIENTRE%(FF%, HOI%) > 55 Then
        Screen.MousePointer = 1
        Call MENSERR(24, "Fecha no Válida")
        FORFACOM.Text9.SetFocus
        Exit Sub
    End If
   '
    Call CIERRARCH("*.*")
    '
    Screen.MousePointer = 11
    NC% = XVALO(FORFACOM.Text1.TEXT)
    If NC% > 0 Then
      If Left$(RASOCLI$((-1) * NC%), 6) <> "VARIOS" Then
        FORFACOM.Text2.TEXT = RASOCLI$((-1) * NC%)
        FORFACOM.Text3.TEXT = DOMICLI$((-1) * NC%)
        FORFACOM.Text4.TEXT = CPOSCLI$((-1) * NC%)
        FORFACOM.Text5.TEXT = LOCACLI$((-1) * NC%)
        FORFACOM.Refresh
      End If
    End If
    '
    CUITT$ = CUITCLI$((-1) * NC%)
    If VALICUIT%(CUITT$) <> 1 Then
      If PIVACLI%((-1) * NC%) <> 5 Then
        Call MENSERR(24, "Imposible Procesar Documento.\Número de C.U.I.T. no Válido")
        FORFACOM.Text1.SetFocus
        Screen.MousePointer = 1
        Exit Sub
      End If
    End If
    '
    If TRIM$(FORFACOM.Text1.TEXT) = "" Then
       Screen.MousePointer = 1
       FORFACOM.Text11.SetFocus
       Exit Sub
    End If
    '
    LETRAFA$ = FORFACOM.Label20.Caption
    If TIDOCOM$ <> "CO" Then    ' evalua numero ingresado
      NROX$ = TRIM$(FORFACOM.Text11.TEXT)
      While Len(NROX$) > 0
         If Asc(Left$(NROX$, 1)) < 48 Or Asc(Left$(NROX$, 1)) > 57 Then
              NROX$ = Mid$(NROX$, 2)
            Else
              GoTo 101
         End If
      Wend
      '
101   If NROX$ = "" Then
         Screen.MousePointer = 1
         FORFACOM.Text11.SetFocus
         Exit Sub
      End If
      '
      SUC% = 0
      C1% = 1: For c% = 1 To 10:  NRO$(c%) = "": Next c%
      '
      For c% = 1 To Len(NROX$)
         C2% = Asc(Mid$(NROX$, c%, 1))
         If C2% = 32 Then GoTo 102
         If (C2% < 48 Or C2% > 57) Then C1% = C1% + 1: GoTo 102
         NRO$(C1%) = NRO$(C1%) + Chr$(C2%)
102   Next c%
      '
      If C1% = 1 Then
         SUC% = 0
         VD# = XVALO(NRO$(1))
         NRO$(2) = NRO$(1)
         NRO$(1) = "0"
      End If
      '
      If XVALO(NRO$(2)) < 1 Or Len(NRO$(1)) > 4 Or Len(NRO$(2)) > 8 Then
          Call MENSERR(24, "Numero de Documento no Permitido")
          Screen.MousePointer = 1
          FORFACOM.Text11.SetFocus
          Exit Sub
      End If
      '
      While Len(NRO$(1)) < 4: NRO$(1) = "0" + NRO$(1): Wend
      While Len(NRO$(2)) < 8: NRO$(2) = "0" + NRO$(2): Wend
      '
      LETRAFA$ = "C": If PIVACLI%((-1) * NC%) = 1 Then LETRAFA$ = "A"
      NROX$ = TIDOCOM$ + "-" + LETRAFA$ + "-" + NRO$(1) + "-" + NRO$(2)
      FORFACOM.Text11.TEXT = NROX$
      FORFACOM.Text11.Refresh
      FORFACOM.Label20.Caption = LETRAFA$
      FORFACOM.Label20.Refresh
      '
    End If
    '
    Screen.MousePointer = 11
    NROY1$ = NROX$
    If TIDOCOM$ <> "CO" Then    ' evalua numero ingresado
        Call REPLA(NROY1$, "X", 4, 1)  'para busqueda de repeticion
      Else
        NROY1$ = "CO-" + NROX$
    End If
    NROY$ = AJUSTI$(NROY1$, 18)
    '
    URECU& = ULTREG&("CUECORR")
    RFF$ = RECFILT$("CUECORR", 4, MKI$(NC%))
    For KKL& = 1 To Len(RFF$) Step 4
      RECUEC& = CVS(Mid$(RFF$, KKL&, 4))
      If RECUEC& > 0 Then
        If RECUEC& <= URECU& Then
          NROZ$ = Mid$(REGLEIDO$("CUECORR", RECUEC&), 7, 18)
          If Left$(NROZ$, 2) <> "CO" Then
            Call REPLA(NROZ$, "X", 4, 1)
          End If
          If TRIM$(NROY$) = TRIM$(NROZ$) Then
            Screen.MousePointer = 1
            Call MENSERR(24, "Documento ya Registrado")
            FORFACOM.Text1.SetFocus
            Exit Sub
          End If
        End If
      End If
    Next KKL&
    Screen.MousePointer = 1
    '
100 FORFACOM.List1.Clear
    FORFACOM.List1.Refresh
    FORFACOM.List2.Clear
    FORFACOM.List2.Refresh
    '
    MAXDEU& = ULTREG&("CUECORR")
    RFD$ = RECFILT$("CUECORR", 4, MKI$(NC%))
    PUCUEN$ = ""
    '
    For U& = 1 To Len(RFD$) Step 4
        REG& = CVS(Mid$(RFD$, U&, 4))
        x$ = REGLEIDO$("CUECORR", REG&)
        If CVI(Mid$(x$, 5, 2)) <> NC% Then
            x$ = REGLEIDO$("CUECORR", 1)
            Call GRAREG("CUECORR", x$, 1)
            Call CIERRARCH("CUECORR")
            GoTo 100
        End If
        '
        NROZ$ = Mid$(x$, 7, 18)
        If Left$(NROZ$, 3) <> "CO-" Then
           Call REPLA(NROZ$, "X", 4, 1)  'para busqueda de repeticion
        End If
        '
        If NROZ$ = NROY$ Then
           Call MENSERR(24, "Documento Pre-existente")
           FORFACOM.Text11.TEXT = ""
           FORFACOM.Text11.Refresh
           FORFACOM.Text11.SetFocus
           Screen.MousePointer = 1
           Exit Sub
        End If
        '
        Y6$ = Mid$(x$, 199, 80)
        For KKI% = 1 To 80 Step 10
          CUEB$ = Mid$(Y6$, KKI%, 2)
          CUEI% = CVI(CUEB$)
          If CUEI% > 0 Then
            For KKJ% = 1 To Len(PUCUEN$) Step 2
              If Mid$(PUCUEN$, KKJ%, 2) = CUEB$ Then GoTo 110
            Next KKJ%
            PUCUEN$ = PUCUEN$ + CUEB$
          End If
110     Next KKI%
    Next U&
    '
    PUCUEX$ = ""
    For KKI% = 1 To Len(PUCUEN$) Step 2
       CUEB$ = Mid$(PUCUEN$, KKI%, 2)
       If ECOARCH$("CUECON", CUEB$) <> "" Then
          PUCUEX$ = PUCUEX$ + CUEB$
          If Len(PUCUEX$) > 16 Then
             PUCUEX$ = Mid$(PUCUEX$, 3)
          End If
       End If
    Next KKI%
    '
    RESCUPRO.LINDICE.Clear
    FORFACOM.List1.Clear
    For KKI% = 1 To Len(PUCUEX$) Step 2
       CUEB$ = Mid$(PUCUEX$, KKI%, 2)
       z$ = Space$(128)
       Call REPLA(z$, CSTRING$(CUEB$, 1, 5, 0, 2), 1, 5)
       DECUE$ = ECOARCH$("CUECON", CUEB$)
       DECUE$ = TRIM$(TRIM$(Mid$(DECUE$, 31, 12)) + " " + TRIM$(Left$(DECUE$, 30)))
       Call REPLA(z$, DECUE$, 8, 30)
       RESCUPRO.LINDICE.AddItem z$
    Next KKI%
    '
    For U& = 1 To RESCUPRO.LINDICE.ListCount
       RESCUPRO.LINDICE.ListIndex = U& - 1
       z$ = RESCUPRO.LINDICE.TEXT
       FORFACOM.List1.AddItem z$
    Next U&
    '
    FORFACOM.List1.Refresh
    '
    If TIDOCOM$ = "NC" Then
      RESCUPRO.LINDICE.Clear
      FORFACOM.List2.Clear
      For U& = 1 To Len(RFD$) Step 4
        REG& = CVS(Mid$(RFD$, U&, 4))
        x$ = REGLEIDO$("CUECORR", REG&)
        FEMI = CVI(Left$(x$, 2))
        IMPO# = CVD(Mid$(x$, 115, 8)) - CVD(Mid$(x$, 123, 8))
        SALDEU$ = Mid$(x$, 139, 60)
        For KKI% = 1 To 6
           FEVE = CVI(Mid$(SALDEU$, 10 * KKI% - 9, 2))
           SALDO# = CVD(Mid$(SALDEU$, 10 * KKI% - 7, 8))
           If Abs(SALDO#) >= 0.005 Then
              If MODORPA% = 1 Then
                  FF% = FEMI      ' orden emision
                Else
                  FF% = FEVE      ' orden vencimiento
              End If
              z$ = CSTRING$(MKI$(FF%), 1, 6, 0, 2) + CSTRING$(MKS$(REG&), 4, 8, 0, 2) + CSTRING$(MKI$(KKI%), 1, 4, 0, 2)
              RESCUPRO.LINDICE.AddItem z$
           End If
        Next KKI%
      Next U&
      '
      SUMADEU# = 0: SUMAFAC# = 0
      For U& = 1 To RESCUPRO.LINDICE.ListCount
        RESCUPRO.LINDICE.ListIndex = U& - 1
        REG& = XVALO(Mid$(RESCUPRO.LINDICE.TEXT, 7, 8))
        KKI% = XVALO(Mid$(RESCUPRO.LINDICE.TEXT, 15, 4))
        If REG& > 0 Then
          If REG& <= MAXDEU& Then
            If KKI% > 0 Then
              If KKI% <= 6 Then
                x$ = REGLEIDO$("CUECORR", REG&)
                TIX$ = Mid$(x$, 7, 2)
                SG% = (-1)
                TIPOS$ = "Factura": If TIX$ = "AS" Then TIPOS$ = "Ajuste de Saldo"
                                    If TIX$ = "TF" Then TIPOS$ = "Ticket-Factura"
                                    If TIX$ = "RH" Then TIPOS$ = "Rec.Honorarios"
                                    If TIX$ = "ND" Then TIPOS$ = "Nota de Debito"
                                    If TIX$ = "CP" Then TIPOS$ = "Crédito Propio"
                                    If TIX$ = "OP" Then TIPOS$ = "Orden de Pago": SG% = 1
                                    If TIX$ = "FF" Then TIPOS$ = "O/P Fondo Fijo": SG% = 1
                                    If TIX$ = "NC" Then TIPOS$ = "Nota de Credito": SG% = 1
                                    If TIX$ = "DP" Then TIPOS$ = "Debito Propio": SG% = 1
                                    If TIX$ = "CC" Then TIPOS$ = "Comp.Cuenta": SG% = 1
                                    '
                IMPO# = CVD(Mid$(x$, 115, 8)) - CVD(Mid$(x$, 123, 8))
                SALDO# = CVD(Mid$(x$, 139 + 10 * KKI% - 8, 8))
                '
                If TIX$ = "FC" Or TIX$ = "ND" Or TIX$ = "CP" Or TIX$ = "TF" Or TIX$ = "RH" Or (TIX$ = "AS" And IMPO# < 0) Then
                  If SALDO# >= 0.005 Then
                    '
                    SUMAFAC# = SUMAFAC# + IMPO#
                    IMPO# = Abs(IMPO#)
                    SUMADEU# = SUMADEU# + SALDO#
                    '
                    NROX$ = Mid$(x$, 7, 18)
                    FEMI = CVI(Mid$(x$, 1, 2))
                    FEVE = CVI(Mid$(x$, 139 + 10 * KKI% - 10, 2))
                    '
                    IMPTEX$ = CSTRING$(MKD$(IMPO#), 3, 13, 2, 2)
                    SALTEX$ = CSTRING$(MKD$(SALDO#), 3, 13, 2, 2)
                    REGTEX$ = CSTRING$(MKS$(REG&), 3, 8, 0, 2)
                    IRETEX$ = CSTRING$(MKI$(KKI%), 1, 2, 0, 2)
                    '
                    z$ = Space$(128)
                    Call REPLA(z$, NROX$, 1, 18)
                    Call REPLA(z$, FECHATEX$(FEMI), 23, 8)
                    Call REPLA(z$, IMPTEX$, 32, 13)
                    Call REPLA(z$, FECHATEX$(FEVE), 47, 8)
                    Call REPLA(z$, SALTEX$, 56, 13)
                    Call REPLA(z$, REGTEX$, 119, 8)
                    Call REPLA(z$, IRETEX$, 127, 2)
                    '
                    FORFACOM.List2.AddItem z$
                    '
                  End If
                End If
              End If
            End If
          End If
        End If
      Next U&
    End If
    '
    FORFACOM.List2.ListIndex = -1
    FORFACOM.List2.Refresh
    FORFACOM.Label8.Caption = CSTRING$(MKD$(SUMADEU#), 4, 14, 2, 2)
    FORFACOM.Label8.Refresh
    FORFACOM.Label10.Caption = ""
    FORFACOM.Label10.Refresh
    FORFACOM.List2.Visible = True
    If FORFACOM.List2.ListCount > 0 Then
        VACON% = 1
        FORFACOM.List2.ListIndex = 0
        VACON% = 0
    End If
    '
    Screen.MousePointer = 1
    If FORFACOM.Check1 Then
        '\\\OT-05-0527-WAR-04/08/05///
        FORFACOM.List1.Clear 'Para que Borre las Imputaciones Contables por Default.
        '\\\OT-05-0527-WAR-FIN///
        Call FORFACOM.SIGUEBORE(XVALO(FORFACOM.Text1)) 'SIGUE CON BR'S SIN FACTUAR
      Else
        Call SIGUEFACO2
    End If
    '
End Sub
'

Sub SIGUEFACO2()
    '
    Call BLANARCH("!CARICOM")
    For U& = 1 To FORFACOM.List1.ListCount
        FORFACOM.List1.ListIndex = U& - 1
        z$ = FORFACOM.List1.TEXT
        CUE% = XVALO(Left$(z$, 5))
        IPA# = XVALO(Mid$(z$, 40, 12))
        '
        y$ = REGBLAN$("!CARICOM")
        Call REPLA(y$, MKI$(CUE%), 1, 2)
        Call REPLA(y$, MKD$(IPA#), 3, 8)
        Call REGAPP("!CARICOM", y$)
    Next U&
    Call CIERRARCH("!CARICOM")
    
    '
200 ATRI$ = "/NC"
    ATRI$ = ATRI$ + "/TITUPAN=Imputacion Contable"
    ATRI$ = ATRI$ + "/BORDESUP=W" + TRIM$(Str$(FORFACOM.Top + FORFACOM.Picture1.Top))
    ATRI$ = ATRI$ + "/BORDEIZQ=W" + TRIM$(Str$(FORFACOM.Left + FORFACOM.Picture1.Left))
    '
    VTB% = VENTABU%("CARIFACO" + ATRI$)
    '
    If VTB% >= 0 Then
       RESCUPRO.LINDICE.Clear
       FORFACOM.List1.Clear
       TOTIMPU# = 0
       For U& = 1 To ULTREG&("!CARICOM")
          y$ = REGLEIDO$("!CARICOM", U&)
          CUEB$ = Left$(y$, 2)
          If CVI(CUEB$) > 0 Then
             z$ = Space$(128)
             Call REPLA(z$, CSTRING$(CUEB$, 1, 5, 0, 2), 1, 5)
             DECUE$ = ECOARCH$("CUECON", CUEB$)
             DECUE$ = TRIM$(TRIM$(Mid$(DECUE$, 31, 12)) + " " + TRIM$(Left$(DECUE$, 30)))
             Call REPLA(z$, DECUE$, 8, 32)
             IMBB$ = Mid$(y$, 3, 8)
             If CVD(IMBB$) < 0 Then
                Call MENSERR(24, "No se Admiten Valores Negativos")
                GoTo 200
             End If
             Call REPLA(z$, CSTRING$(IMBB$, 3, 13, 2, 2), 38, 13)
             TOTIMPU# = TOTIMPU# + CVD(IMBB$)
             RESCUPRO.LINDICE.AddItem z$
          End If
       Next U&
       '
       For U& = 1 To RESCUPRO.LINDICE.ListCount
          RESCUPRO.LINDICE.ListIndex = U& - 1
          z$ = RESCUPRO.LINDICE.TEXT
          FORFACOM.List1.AddItem z$
       Next U&
       '
       For KKI% = 1 To 5
          FORFACOM.IMPOGRA(KKI%) = Space$(11)
          FORFACOM.IMPOIVA(KKI%) = Space$(8)
       Next KKI%
       '
       If LETRAFA$ = "A" Then
           FORFACOM.ITOGRA.Caption = CSTRING$(MKD$(TOTIMPU#), 4, 11, 2, 2)
           FORFACOM.ITOGRA.Refresh
         Else
           FORFACOM.ITOGRA.Caption = Space$(11)
           FORFACOM.IMPOGRA(4).TEXT = CSTRING$(MKD$(TOTIMPU#), 4, 11, 2, 2)
           FORFACOM.IMPOGRA(4).Refresh
       End If
    End If
    FORFACOM.List1.Refresh
    FORFACOM.IMPOGRA(1).SetFocus
    '\\\OT-05-0527-WAR-04/08/05///
    FORFACOM.Label28 = FORMATNUM$(TOTIMPU#, "F11.2")
    '\\\OT-05-0527-WAR-FIN///
    If TOTIMPU# < 0.01 Then
       FORFACOM.Text1.SetFocus
    End If
    '
End Sub

Sub SIGUEFACO3()
   '
      For UI% = 1 To 5
        FORFACOM.IMPOGRA(UI%).TEXT = Space$(11)
        FORFACOM.IMPOIVA(KKI%) = Space$(8)
      Next UI%
      If LETRAFA$ = "A" Then
           FORFACOM.ITOGRA.Caption = Space$(11)
           FORFACOM.ITOGRA.Caption = CSTRING$(MKD$(TOTIMPU#), 4, 11, 2, 2)
           FORFACOM.ITOGRA.Refresh
         Else
           FORFACOM.ITOGRA.Caption = Space$(11)
           FORFACOM.IMPOGRA(4).TEXT = CSTRING$(MKD$(TOTIMPU#), 4, 11, 2, 2)
           FORFACOM.IMPOGRA(4).Refresh
      End If
      TOEXEN# = XVALO(FORFACOM.IMPOGRA(4).TEXT)
      FORFACOM.IMPOGRA(1).SetFocus
   TONEGRA# = TOTIMPU# - TOEXEN#
   FORFACOM.ITOGRA.Caption = CSTRING$(MKD$(TONEGRA#), 4, 11, 2, 2)
   FORFACOM.ITOGRA.Refresh
   '
End Sub


Sub SIGUEFACO5()
    '
    If FORFACOM.List2.ListCount < 1 Then
        Exit Sub
    End If
    '
    Call BLANARCH("!CARGAPLI")
    For U& = 1 To FORFACOM.List2.ListCount
        FORFACOM.List2.ListIndex = U& - 1
        z$ = FORFACOM.List2.TEXT
        NROX$ = Mid$(z$, 1, 20)
        FEMII% = FECHANUM(Mid$(z$, 23, 8))
        IPED# = CDbl(XVALO(Mid$(z$, 32, 13)))
        FEVEI% = FECHANUM(Mid$(z$, 47, 8))
        SALD# = CDbl(XVALO(Mid$(z$, 56, 13)))
        ACRE# = CDbl(XVALO(Mid$(z$, 70, 14)))
        DOCU$ = Mid$(z$, 1, 18)
        REG2& = XVALO(Mid$(z$, 119, 8))
        IRE% = XVALO(Mid$(z$, 127, 2))
        '
        y$ = REGBLAN$("!CARGAPLI")
        Call REPLA(y$, NROX$, 1, 20)
        Call REPLA(y$, MKI$(FEMII%), 21, 2)
        Call REPLA(y$, MKD$(IPED#), 23, 8)
        Call REPLA(y$, MKI$(FEVEI%), 31, 2)
        Call REPLA(y$, MKD$(SALD#), 33, 8)
        Call REPLA(y$, MKD$(ACRE#), 41, 8)
        Call REPLA(y$, DOCU$, 49, 18)
        Call REPLA(y$, MKS$(REG2&), 67, 4)
        Call REPLA(y$, MKI$(IRE%), 71, 2)
        '
        Call REGAPP("!CARGAPLI", y$)
    Next U&
    Call CIERRARCH("!CARGAPLI")
    '
200 ATRI$ = "/NC"
    ATRI$ = ATRI$ + "/TITUPAN=Aplicación de Pagos"
    ATRI$ = ATRI$ + "/BORDESUP=W" + TRIM$(Str$(FORFACOM.Top + FORFACOM.Frame8.Top))
    ATRI$ = ATRI$ + "/BORDEIZQ=W" + TRIM$(Str$(FORFACOM.Left + FORFACOM.Picture2.Left))
    '
    VTB% = VENTABU%("CARGAPLI" + ATRI$)
    '
    VACON% = 1
    If VTB% >= 0 Then
        FORFACOM.List2.Clear
        For U& = 1 To ULTREG&("!CARGAPLI")
            y$ = REGLEIDO$("!CARGAPLI", U&)
            NROX$ = Mid$(y$, 1, 20)
            If TRIM$(NROX$) = "Anticipo no Aplicado" Then
               HOI% = HOY(HO$)
               ACRE# = CVD(Mid$(y$, 41, 8))
               Call REPLA(y$, MKI$(HOI%), 31, 2)
               Call REPLA(y$, MKD$(ACRE#), 23, 8)
               Call REPLA(y$, MKD$(ACRE#), 33, 8)
               Call GRAREG("!CARGAPLI", y$, U&)
            End If
            FEMI = CVI(Mid$(y$, 21, 2))
            IPED# = CVD(Mid$(y$, 23, 8))
            IMPTEX$ = CSTRING$(MKD$(IPED#), 4, 13, 2, 2)
            FEVE = CVI(Mid$(y$, 31, 2))
            SALD# = CVD(Mid$(y$, 33, 8))
            SALTEX$ = CSTRING$(MKD$(SALD#), 4, 13, 2, 2)
            ACRE# = CVD(Mid$(y$, 41, 8))
            '
            If ACRE# < 0 Then
                Call MENSERR(24, "No se Permite Aplicación Negativa")
                GoTo 200
            End If
            If U& > 1 Then
              If ACRE# > SALD# Then
                Call MENSERR(24, "El Importe Acreditado Debe Ser\Inferior al Saldo del Documento")
                 GoTo 200
              End If
            End If
            ACRTEX$ = CSTRING$(MKD$(ACRE#), 4, 14, 2, 2)
            DOCU$ = Mid$(y$, 49, 18)
            REG2& = CVS(Mid$(y$, 67, 4))
            IRE% = CVI(Mid$(y$, 71, 2))
            REGTEX$ = CSTRING$(MKS$(REG2&), 3, 8, 0, 2)
            IRETEX$ = CSTRING$(MKI$(IRE%), 1, 2, 0, 2)
            '
            z$ = Space$(128)
            Call REPLA(z$, NROX$, 1, 20)
            Call REPLA(z$, FECHATEX$(FEMI), 23, 8)
            Call REPLA(z$, IMPTEX$, 32, 13)
            Call REPLA(z$, FECHATEX$(FEVE), 47, 8)
            Call REPLA(z$, SALTEX$, 56, 13)
            Call REPLA(z$, ACRTEX$, 70, 14)
            Call REPLA(z$, REGTEX$, 119, 8)
            Call REPLA(z$, IRETEX$, 127, 2)
            FORFACOM.List2.AddItem z$
            '
        Next U&
    End If
    '
    SUMACRE# = 0
    For U& = 1 To FORFACOM.List2.ListCount
        FORFACOM.List2.ListIndex = U& - 1
        z$ = FORFACOM.List2.TEXT
        SUMACRE# = SUMACRE# + CDbl(XVALO(Mid$(z$, 70, 14)))
    Next U&
    '
    FORFACOM.Label10.Caption = CSTRING$(MKD$(SUMACRE#), 4, 14, 2, 2)
    FORFACOM.Label10.Refresh
    '
    FORFACOM.List2.Refresh
    VACON% = 0
    '
End Sub
'

Sub ANUFACO()

4000 ' anulacion de documentos
     ' ***********************
     '
     ' 1.- Ingresar Numero de Proveedor
     '     ****************************
     '
10   Screen.MousePointer = 1
     Call SELECHO("PROVED1")
     NC% = XVALO(VALACT1$("PROVED1"))
     If NC% < 1 Then Exit Sub
     '
     FECHALIM% = XVALO(CONTROL$("", "CIEMECOM"))
     '
100  Call BLANARCH("!SELCUECO")
     MAXDEU& = ULTREG&("CUECORR")
     RFD$ = RECFILT$("CUECORR", 4, MKI$(NC%))
     '
     For U& = 1 To Len(RFD$) Step 4
        REG& = CVS(Mid$(RFD$, U&, 4))
        x$ = REGLEIDO$("CUECORR", REG&)
        If CVI(Mid$(x$, 5, 2)) <> NC% Then
            x$ = REGLEIDO$("CUECORR", 1)
            Call GRAREG("CUECORR", x$, 1)
            Call CIERRARCH("CUECORR")
            GoTo 100
        End If
        '
        If CVI(Left$(x$, 2)) > FECHALIM% Then
           TIDOC$ = UCase$(Mid$(x$, 7, 2))
           If TIDOC$ = "FC" Or TIDOC$ = "ND" Or TIDOC$ = "CP" Or TIDOC$ = "NC" Or TIDOC$ = "TF" Or TIDOC$ = "RH" Or TIDOC$ = "CO" Then
              IMPO# = Abs(CVD(Mid$(x$, 115, 8))) + Abs(CVD(Mid$(x$, 123, 8)))
              If IMPO# >= 0.005 Then
                 IMPOS$ = NUMSTRI$(IMPO#, 14, 2, 1)
                 y$ = MKS$(REG&) + Mid$(x$, 7, 18) + IMPOS$
                 Call REGAPP("!SELCUECO", y$)
              End If
           End If
        End If
     Next U&
     Call CIERRARCH("!SELCUECO")
     If ULTREG&("!SELCUECO") < 1 Then
        Call MENSERR(24, "No hay Documentos en la Cuenta Elegida\que Puedan Ser Anulados.")
        GoTo 10
     End If
     Call FRESHECHO("!SELCUECO")
     '
     Call SELECHO("!SELCUECO")
     If VALACT1$("!SELCUECO") = "" Then GoTo 10
     REGCUEC& = XVALO(VALACT1$("!SELCUECO"))
     If REGCUEC& < 1 Then GoTo 10
     '
     x$ = REGLEIDO$("CUECORR", REGCUEC&)
     IMPO# = Abs(CVD(Mid$(x$, 115, 8))) + Abs(CVD(Mid$(x$, 123, 8)))
     IMPOS$ = Left$("$ " + TRIM$(NUMSTRI$(IMPO#, 14, 2, 1)) + Space$(18), 18)
     '
     SALDO# = 0
     SALDEU$ = Mid$(x$, 139, 60)
     For KKI% = 1 To 6
        FEVE = CVI(Mid$(SALDEU$, 10 * KKI% - 9, 2))
        If FEVE > 0 Then
           SALDO# = SALDO# + Abs(CVD(Mid$(SALDEU$, 10 * KKI% - 7, 8)))
        End If
     Next KKI%
     ATX% = 0
     If Abs(SALDO# - IMPO#) >= 0.005 Then
        TTXX$ = "Sobre este Comprobante ya hay\Créditos o Pagos Aplicados al Mismo."
        TTXX$ = TTXX$ + "\  \En Caso de Confirmarse la Anulación\Deberá Permitirse la Re-Construcción Automática\de la Composición de Deuda de la Cuenta Elegida."
        TTXX$ = TTXX$ + "\   \Desea Camcelar o Continuar el Proceso ?"
        If COMALTER%(TTXX$ + "\\Cancelar\Continuar") <> 2 Then
           GoTo 10
        End If
        ATX% = 1
     End If
     '
     NF$ = TRIM$(Mid$(x$, 7, 18))
     If COMALTER%("Documento: " + NF$ + "\  Importe: " + IMPOS$ + "\\Abandona\Confirma Anulacion") <> 2 Then GoTo 10
     '
     ' anulacion
     ' *********
     Screen.MousePointer = 11
     Call BLOQARCH("CUECORR")
     Call BLOQARCH("CUECOSS")
     Call BLOQARCH("CAJABANC")
     Call BLOQARCH("PROVED1")
     Call BLOQARCH("PROVED2")
     '
     FEMOA% = CVI(Left$(x$, 2))
     NCLPR$ = Mid$(x$, 3, 4)
     NPRX% = CVI(Mid$(x$, 5, 2))
     NDOCA$ = Mid$(x$, 7, 18)
     XD# = CVD(Mid$(x$, 115, 8))
     XH# = CVD(Mid$(x$, 123, 8))
     '
     Call REPLA$(x$, "anul", 12, 4)
     Call REPLA(x$, String$(228, 0), 51, 228)
     Call GRAREG("CUECORR", x$, REGCUEC&)
     Call CIERRARCH("CUECORR")
     '
     If REGCUEC& <= ULTREG&("CUECOSS") Then
        X1$ = REGLEIDO$("CUECOSS", REGCUEC&)
        Call REPLA$(X1$, "anul", 12, 4)
        Call REPLA(X1$, String$(228, 0), 51, 228)
        Call REPLA(X1$, String$(212, 0), 301, 212)
        Call GRAREG("CUECOSS", X1$, REGCUEC&)
     End If
     '
     FOFIJO% = 0
     XF$ = REGLEIDO$("CUECORR", REGCUEC& + 1)
     If CVI(Left$(XF$, 2)) = FEMOA% Then
       If Mid$(XF$, 3, 4) = NCLPR$ Then
         If Mid$(XF$, 7, 18) = "FF" + Mid$(NDOCA$, 3) Then
           If Abs(CVD(Mid$(XF$, 115, 8)) - CVD(Mid$(XF$, 123, 8)) + XD# - XH#) < 0.01 Then
             FOFIJO% = 1
           End If
         End If
       End If
     End If
     '
     If FOFIJO% = 1 Then
       Call REPLA$(XF$, "anul", 12, 4)
       Call REPLA(XF$, String$(228, 0), 51, 228)
       Call GRAREG("CUECORR", XF$, REGCUEC& + 1)
       Call CIERRARCH("CUECORR")
       '
       If REGCUEC& < ULTREG&("CUECOSS") Then
         X1$ = REGLEIDO$("CUECOSS", REGCUEC& + 1)
         Call REPLA$(X1$, "anul", 12, 4)
         Call REPLA(X1$, String$(228, 0), 51, 228)
         Call REPLA(X1$, String$(212, 0), 301, 212)
         Call GRAREG("CUECOSS", X1$, REGCUEC& + 1)
       End If
       Call CIERRARCH("CUECOSS")
     End If
     '
80   If Len(TRIM$(NF$)) < 18 Then
        NRECIB$ = TRIM$(NF$)
        While Len(NRECIB$) > 12: NRECIB$ = Mid$(NRECIB$, 2): Wend
      Else
        NRECIB$ = Left$(NF$, 5) + Mid$(NF$, 9, 1) + "-" + Mid$(NF$, 14, 5)
     End If
     NRECIB$ = TRIM$(NRECIB$)
     '
     For U& = 1 To ULTREG&("CAJABANC")
       XF$ = REGLEIDO$("CAJABANC", U&)
       If TRIM$(Left$(XF$, 12)) = NRECIB$ Then
         If CVI(Mid$(XF$, 61, 2)) = FEMOA% Then
           If Abs(CVI(Mid$(XF$, 67, 2))) = NPRX% Then
             Call REPLA(XF$, String$(32, 0), 145, 32)
             Call GRAREG("CAJABANC", XF$, U&)
           End If
         End If
       End If
     Next U&
     '
85   RFG$ = RECFILT$("PERCIBFA", 2, MKI$(NPRX%))
     For UKY& = 1 To Len(RFG$) Step 4
       RGY& = CVS(Mid$(RFG$, UKY&, 4))
       RGZ$ = REGLEIDO$("PERCIBFA", RGY&)
       If TRIM$(Mid$(RGZ$, 5, 18)) = TRIM$(NF$) Then
         If CVI(Left$(RGZ$, 2)) = FEMOA% Then
           Call REPLA(RGZ$, String$(48, 0), 81, 48)
           Call GRAREG("PERCIBFA", RGZ$, RGY&)
         End If
       End If
     Next UKY&
     Call CIERRARCH("PERCIBFA")
     '
90   Call CIERRARCH("CAJABANC")
     Call CIERRARCH("CUECORR")
     Call CIERRARCH("CUECOSS")
     '
     For U& = 1 To ULTREG&("PROVED1")
        D1$ = Left$(REGLEIDO$("PROVED1", U&), 2)
        If CVI(D1$) = NC% Then
           D2$ = REGLEIDO$("PROVED2", U&)
           SALDO# = CVD(Mid$(D2$, 109, 8)) + XD# - XH#
           Call REPLA(D2$, MKD$(SALDO#), 109, 8)
           Call GRAREG("PROVED2", D2$, U&)
           GoTo 99
        End If
     Next U&
     '
     '\\\OT-05-0527-WAR-09/08/05///
99   JJ& = 0
     For j& = 1 To ULTREG("DETFACOM")
       XX$ = REGLEIDO("DETFACOM", j&)
       NP1% = CVI(Mid$(XX$, 1, 2))
       NF1$ = TRIM$(Mid$(XX$, 3, 18))
       If NF$ <> NF1$ Then ' <> COMPROBANTE
         JJ& = JJ& + 1
         Call GRAREG("DETFACOM", XX$, JJ&)
       ElseIf NC% <> NP1% Then ' MISMO COMPROBANTE Y <> PROVEEDOR
         JJ& = JJ& + 1
         Call GRAREG("DETFACOM", XX$, JJ&)
       End If
     Next j&
     Call SETULTREG("DETFACOM", JJ&)
     '
     JJ& = 0
     For j& = 1 To ULTREG("FACBONCO")
       XX$ = REGLEIDO("FACBONCO", j&)
       NP1% = CVI(Mid$(XX$, 1, 2))
       NF1$ = TRIM$(Mid$(XX$, 9, 18))
       If NF$ <> NF1$ Then ' <> COMPROBANTE
         JJ& = JJ& + 1
         Call GRAREG("FACBONCO", XX$, JJ&)
       ElseIf NC% <> NP1% Then ' MISMO COMPROBANTE Y <> PROVEEDOR
         JJ& = JJ& + 1
         Call GRAREG("FACBONCO", XX$, JJ&)
       End If
     Next j&
     Call SETULTREG("FACBONCO", JJ&)
     '
     '\\\OT-05-0527-WAR-FIN///
     Call CIERRARCH("PROVED1")
     Call CIERRARCH("PROVED2")
     '
     If ATX% = 1 Then
        Call REAPLICREPRO(NC%)
     End If
     '
     GoTo 10
     '
End Sub
'

Sub ANUORPA()       ' anulacion de ordenes de pago
     '
     ' 1.- Ingresar Numero de Proveedor
     '     ****************************
     '
     FECHALIM% = XVALO(CONTROL$("", "CIEMECOM"))
     '
10   Screen.MousePointer = 1
     Call SELECHO("PROVED1")
     NC% = XVALO(VALACT1$("PROVED1"))
     If NC% < 1 Then Exit Sub
     '
100  Call BLANARCH("!SELCUECO")
     MAXDEU& = ULTREG&("CUECORR")
     RFD$ = RECFILT$("CUECORR", 4, MKI$(NC%))
     '
     For U& = 1 To Len(RFD$) Step 4
        REG& = CVS(Mid$(RFD$, U&, 4))
        x$ = REGLEIDO$("CUECORR", REG&)
        If CVI(Mid$(x$, 5, 2)) <> NC% Then
            x$ = REGLEIDO$("CUECORR", 1)
            Call GRAREG("CUECORR", x$, 1)
            Call CIERRARCH("CUECORR")
            GoTo 100
        End If
        '
        If CVI(Left$(x$, 2)) > FECHALIM% Then
           TIDOC$ = UCase$(Mid$(x$, 7, 2))
           If TIDOC$ = "OP" Then
              IMPO# = Abs(CVD(Mid$(x$, 115, 8))) + Abs(CVD(Mid$(x$, 123, 8)))
              If IMPO# >= 0.005 Then
                 IMPOS$ = NUMSTRI$(IMPO#, 14, 2, 1)
                 y$ = MKS$(REG&) + Mid$(x$, 7, 18) + IMPOS$
                 Call REGAPP("!SELCUECO", y$)
              End If
           End If
        End If
     Next U&
     '
     If ULTREG&("!SELCUECO") < 1 Then
        Call MENSERR(24, "La Cuenta Elegida no Registra\Ordenes de Pago que Puedan Anularse.")
        GoTo 10
     End If
     '
     Call CIERRARCH("!SELCUECO")
     Call FRESHECHO("!SELCUECO")
     '
     Call SELECHO("!SELCUECO")
     If VALACT1$("!SELCUECO") = "" Then GoTo 10
     REGCUEC& = XVALO(VALACT1$("!SELCUECO"))
     If REGCUEC& < 1 Then GoTo 10
     '
     XX0$ = REGLEIDO$("CUECORR", REGCUEC&)
     IMPO# = Abs(CVD(Mid$(XX0$, 115, 8))) + Abs(CVD(Mid$(XX0$, 123, 8)))
     IMPOS$ = Left$("$ " + TRIM$(NUMSTRI$(IMPO#, 14, 2, 1)) + Space$(18), 18)
     x$ = TRIM$(Mid$(XX0$, 7, 18))
     y$ = IMPOS$
     '
     If COMALTER%("Documento: " + x$ + "\  Importe: " + y$ + "\\Abandona\Confirma Anulacion") <> 2 Then GoTo 10
     '
     ' anulacion
     ' *********
     '
3011 Call BLOQARCH("CUECORR")
     Call BLOQARCH("PROVED2")
     Call BLOQARCH("CHEQUEMI")
     Call BLOQARCH("DETPAGAS")
     Screen.MousePointer = 11
     '
     FF1% = CVI(Left$(XX0$, 2))
3012 NF$ = Mid$(XX0$, 7, 18)
     Call REPLA(NF$, "anul", 6, 4)
     NORPA& = XVALO(Mid$(NF$, 11, 8))
     PROV% = CVI(Mid$(XX0$, 5, 2))
3013 XD# = CVD(Mid$(XX0$, 115, 8))
3014 XH# = CVD(Mid$(XX0$, 123, 8))
3015 SS8X$ = Mid$(XX0$, 139, 60)
     IM# = CVD(Mid$(SS8X$, 3, 8)) + XD# - XH#
     Call REPLA(SS8X$, MKI$(FF1%), 1, 2)
     Call REPLA(SS8X$, MKD$(IM#), 3, 8)
     Call REPLA(XX0$, NF$, 7, 18)
     Call REPLA(XX0$, MKD$(0), 115, 8)
     Call REPLA(XX0$, MKD$(0), 123, 8)
     Call REPLA(XX0$, SS8X$, 139, 60)
     Call GRAREG("CUECORR", XX0$, REGCUEC&)
     '
     NC1% = (-1) * NC%
     RC& = REGICLI&(NC1%)
     DX$ = REGLEIDO$("PROVED2", RC&)
     SALDO# = CVD(Mid$(DX$, 109, 8)) + XD# - XH#
     Call REPLA(DX$, MKD$(SALDO#), 109, 8)
     Call GRAREG("PROVED2", DX$, RC&)
     '
     FIN& = ULTREG&("CHEQUEMI")
     NOP& = NORPA&
     '
     For U& = 1 To FIN&
        Call TRACE(20, U&, FIN&)
        z$ = REGLEIDO$("CHEQUEMI", U&)
        If CVS(Mid$(z$, 41, 4)) = NOP& Then
          If CVI(Mid$(z$, 47, 2)) = FF1% Then
            If CVI(Mid$(z$, 1, 2)) = PROV% Then
              Call REPLA(z$, MKD$(0), 33, 8)
              Call GRAREG("CHEQUEMI", z$, U&)
            End If
          End If
        End If
     Next U&
     Call CIERRARCH("CHEQUEMI")
     '
     FIN& = ULTREG&("DETPAGAS")
     '
     For U& = 1 To FIN&
        Call TRACE(20, U&, FIN&)
        z$ = REGLEIDO$("DETPAGAS", U&)
        If CVS(Mid$(z$, 3, 4)) = NOP& Then
          If CVI(Mid$(z$, 7, 2)) = FF1% Then
            If CVI(Left$(z$, 2)) = PROV% Then
              Call REPLA(z$, MKD$(0), 9, 8)
              Call GRAREG("DETPAGAS", z$, U&)
            End If
          End If
        End If
     Next U&
     Call CIERRARCH("DETPAGAS")
     '
     FIN& = ULTREG&("SURETIB")
     '
     For U& = 1 To FIN&
        Call TRACE(20, U&, FIN&)
        z$ = REGLEIDO$("SURETIB", U&)
        If CVS(Mid$(z$, 7, 4)) = NOP& Then
          If CVI(Mid$(z$, 1, 2)) = FF1% Then
            Call REPLA(z$, MKD$(0), 79, 8)
            Call REPLA(z$, MKD$(0), 87, 8)
            Call GRAREG("SURETIB", z$, U&)
          End If
        End If
     Next U&
     Call CIERRARCH("SURETIB")
     '
     If TICUEPRO%((-1) * NC%) = 0 Then
        TTXX$ = "Anulación Completa.\  \Desea Reconstruir la Composición de Deuda\de la Cuenta '" + TRIM$(RASOCLI$((-1) * NC%)) + "' ?"
        If COMALTER%(TTXX$ + "\\Reconstruir\Continuar") = 1 Then
           Call REAPLICREPRO(NC%)
        End If
     End If
     '
     Call CIERRARCH("CUECORR")
     Call CIERRARCH("PROVED2")
     '
     FIN& = ULTREG&("BADACHE")
     NOP& = NORPA&
     '
     For U& = 1 To FIN&
        Call TRACE(20, U&, FIN&)
        TTS$ = REGLEIDO$("BADACHE", U&)
        If Mid$(TTS$, 135, 18) = "Orden de Pago Nro." Then
          If XVALO(Mid$(TTS$, 135 + 18, 8)) = NOP& Then
            If CVI(Mid$(TTS$, 133, 2)) = FF1% Then
              Call REPLA(TTS$, Space$(30), 103, 30)
              Call REPLA(TTS$, MKI$(0), 133, 2)
              Call REPLA(TTS$, Space$(120), 135, 120)
              Call GRAREG("BADACHE", TTS$, U&)
            End If
          End If
        End If
     Next U&
     Call CIERRARCH("BADACHE")
     '
     GoTo 10
     '
End Sub
'

Sub REAPLICREPRO(NCX%)
    '
    If TICUEPRO%((-1) * NCX%) <> 0 Then
       Call MENSERR(24, "Proceso Imposible - Cuenta de Gastos")
       GoTo 999
    End If
    '
    ATENTI% = 1
    Dim DOCU$(8192), IMPO#(8192), REGDO&(8192), FFDOC%(8192)
    '
    Call BLOQARCH("CUECORR")
    '
100 Screen.MousePointer = 11
    '
    MAXDIAL& = ULTREG&("CUECORR")
    RFF$ = RECFILT$("CUECORR", 4, MKI$(NCX%))
    '
    NUDOCU% = 0
    FIN& = Len(RFF$)
    For U& = 1 To Len(RFF$) Step 4
        Call TRACE(20, U&, FIN&)
        REG& = CVS(Mid$(RFF$, U&, 4))
        x$ = REGLEIDO$("CUECORR", REG&)
        If CVI(Mid$(x$, 5, 2)) <> NCX% Then
            x$ = REGLEIDO$("CUECORR", 1)
            Call GRAREG("CUECORR", x$, 1)
            Call CIERRARCH("CUECORR")
            GoTo 100
        End If
        NUDOCU% = NUDOCU% + 1
        If NUDOCU% > 8192 Then
           Call MENSERR(24, "Imposible Continuar Proceso.\Demasiados Movimientos en la Cuenta Elegida.\Consulte a su Soporte de Sistemas FLEXOFT.")
           Call CIERRARCH("*.*")
           Call FINAL("")
        End If
        IMPO#(NUDOCU%) = CVD(Mid$(x$, 123, 8)) - CVD(Mid$(x$, 115, 8))
        DOCU$(NUDOCU%) = Mid$(x$, 7, 2) + Mid$(x$, 12, 13)
        REGDO&(NUDOCU%) = REG&
        FFDOC%(NUDOCU%) = CVI(Mid$(x$, 1, 2))
    Next U&
    '
    FIN& = NUDOCU%
    For KKI% = 1 To NUDOCU%
       U& = KKI%
       Call TRACE(20, U&, FIN&)
       For KKJ% = KKI% + 1 To NUDOCU%
         If FFDOC%(KKI%) > FFDOC%(KKJ%) Then
           DOCX$ = DOCU$(KKI%)
           FFX% = FFDOC%(KKI%)
           IMX# = IMPO#(KKI%)
           RGX& = REGDO&(KKI%)
           DOCU$(KKI%) = DOCU$(KKJ%)
           FFDOC%(KKI%) = FFDOC%(KKJ%)
           IMPO#(KKI%) = IMPO#(KKJ%)
           REGDO&(KKI%) = REGDO&(KKJ%)
           DOCU$(KKJ%) = DOCX$
           FFDOC%(KKJ%) = FFX%
           IMPO#(KKJ%) = IMX#
           REGDO&(KKJ%) = RGX&
         End If
       Next KKJ%
    Next KKI%
    '
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
    FIN& = ULTREG&("APLIPAG")
    For U& = 1 To FIN&
       Call TRACE(20, U&, FIN&)
       x$ = REGLEIDO$("APLIPAG", U&)
       If CVI(Mid$(x$, 3, 2)) = NCX% Then
          IAPLI# = CVD(Mid$(x$, 23, 8))
          RECI$ = Mid$(x$, 5, 18)
          RECI$ = Left$(RECI$, 2) + Mid$(RECI$, 6)
          FACU$ = Mid$(x$, 55, 18)
          FACU$ = Left$(FACU$, 2) + Mid$(FACU$, 6)
          '
          IRECIMA# = 0: IFACMAX# = 0
          For KKI% = 1 To NUDOCU%
             If DOCU$(KKI%) = RECI$ Then
               If IMPO#(KKI%) < 0 Then
                 IRECIMA# = IRECIMA# - IMPO#(KKI%)
               End If
             End If
             If DOCU$(KKI%) = FACU$ Then
               If IMPO#(KKI%) > 0 Then
                 IFACMAX# = IFACMAX# + IMPO#(KKI%)
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
                If DOCU$(KKI%) = RECI$ Then
                   If IMPO#(KKI%) < 0 Then
                      IAP1# = IRECIMA#
                      If IAP1# > Abs(IMPO#(KKI%)) Then IAP1# = Abs(IMPO#(KKI%))
                      IMPO#(KKI%) = IMPO#(KKI%) + IAP1#
                      IRECIMA# = IRECIMA# - IAP1#
                   End If
                End If
                If DOCU$(KKI%) = FACU$ Then
                   If IMPO#(KKI%) > 0 Then
                      IAP2# = IFACMAX#
                      If IAP2# > Abs(IMPO#(KKI%)) Then IAP2# = Abs(IMPO#(KKI%))
                      IMPO#(KKI%) = IMPO#(KKI%) - IAP2#
                      IFACMAX# = IFACMAX# - IAP2#
                   End If
                End If
             Next KKI%
          End If
          '
        End If
    Next U&
    '
    FIN& = NUDOCU%
    For KKI% = 1 To NUDOCU%
       U& = KKI%
       Call TRACE(20, U&, FIN&)
       For KKJ% = KKI% + 1 To NUDOCU%
         If FFDOC%(KKI%) > FFDOC%(KKJ%) Then
           DOCX$ = DOCU$(KKI%)
           FFX% = FFDOC%(KKI%)
           IMX# = IMPO#(KKI%)
           RGX& = REGDO&(KKI%)
           DOCU$(KKI%) = DOCU$(KKJ%)
           FFDOC%(KKI%) = FFDOC%(KKJ%)
           IMPO#(KKI%) = IMPO#(KKJ%)
           REGDO&(KKI%) = REGDO&(KKJ%)
           DOCU$(KKJ%) = DOCX$
           FFDOC%(KKJ%) = FFX%
           IMPO#(KKJ%) = IMX#
           REGDO&(KKJ%) = RGX&
         End If
       Next KKJ%
    Next KKI%
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
    Screen.MousePointer = 1
    TTXX$ = "Se ha Completado la Reconstrucción\de la Composición de la Deuda de\   \'" + TRIM$(RASOCLI$((-1) * NCX%))
    If COMALTER%(TTXX$ + "'\   \Desea Grabar la Nueva Composición de la Deuda ?\\Cancelar\Grabar") <> 2 Then
       GoTo 999
    End If
    '
    Screen.MousePointer = 11
    HOII% = HOY(HO$)
    FIN& = NUDOCU%
    For KKI% = 1 To NUDOCU%
       U& = KKI%
       Call TRACE(20, U&, FIN&)
       If REGDO&(KKI%) > 0 Then
          x$ = REGLEIDO$("CUECORR", REGDO&(KKI%))
          If Abs(IMPO#(KKI%)) < 0.005 Then
               Call REPLA(x$, String$(60, 0), 139, 60)
               GoTo 890
             Else
               FEMI% = CVI(Left$(x$, 2))
               SIG% = Sgn(IMPO#(KKI%))
               ITOT# = Abs(IMPO#(KKI%))
               VENUE$ = ""
               For KKK% = 139 To 198 Step 10
                  FEVE% = CVI(Mid$(x$, KKK%, 2))
                  IPAX# = Abs(CVD(Mid$(x$, KKK% + 2, 8)))
                  If FEVE% < 1 Then FEVE% = 0: IPAX# = 0
                  If IPAX# > ITOT# Then IPAX# = ITOT#
                  If FEVE% > 0 Then
                     If IPAX# >= 0.005 Then
                        VENUE$ = VENUE$ + MKI$(FEVE%) + MKD$(IPAX# * SIG%)
                        ITOT# = ITOT# - IPAX#
                     End If
                  End If
               Next KKK%
               If ITOT# >= 0.005 Then
                  If Len(VENUE$) > 50 Then
                     IPAX# = ITOT# * SIG% + CVD(Mid$(VENUE$, 3, 8))
                     Call REPLA(VENUE$, MKD$(IPAX#), 3, 8)
                    Else
                     FEVE% = HOII%: If SIG% < 0 Then FEVE% = FEMI%
                     VENUE$ = VENUE$ + MKI$(FEVE) + MKD$(ITOT# * SIG%)
                  End If
               End If
               VENUE$ = VENUE$ + String$(60, 0)
               Call REPLA(x$, VENUE$, 139, 60)
          End If
890       Call GRAREG("CUECORR", x$, REGDO&(KKI%))
       End If
    Next KKI%
    '
999 Call CIERRARCH("CUECORR")
    Call LIBARCH("CUECORR")
    Screen.MousePointer = 1
    ATENTI% = 0
    '
End Sub

Sub IMPUPRO1()
    '
    On Error Resume Next
    FORPASTA.Hide
    FORFACOM.Hide
    FORFAC00.Hide
    MENUACRE.Hide
    FIMPUPRO.Show
    On Error GoTo 0
    '
    ' 1.- Ingresar Numero de Cliente
    '     **************************
    '
10  FIMPUPRO.List2.Clear
    FIMPUPRO.Label1.Caption = ""
    FIMPUPRO.Label2.Caption = ""
    FIMPUPRO.Label3.Caption = ""
    FIMPUPRO.Label6.Caption = ""
    FIMPUPRO.Label12.Caption = ""
    FIMPUPRO.Label15.Caption = ""
    FIMPUPRO.Label22.Caption = ""
    FIMPUPRO.Label23.Caption = ""
    FIMPUPRO.Label24.Caption = ""
    FIMPUPRO.Label25.Caption = ""
    FIMPUPRO.Label8.Caption = ""
    FIMPUPRO.Label10.Caption = ""
    FIMPUPRO.Label14.Caption = ""
    FIMPUPRO.Label18.Caption = ""
    FIMPUPRO.Label20.Caption = ""
    FIMPUPRO.Command5.Enabled = False
    FIMPUPRO.Refresh
    Call FRESHECHO("!SELDOCAP")
    '
    If MODORPA% = 0 Then
        MODORPA% = 1       ' por emision
        If CONTROL$("ORPASTA", "MODORDE") = "VENCIM" Then
           MODORPA% = 2
        End If
    End If
    '
    FIMPUPRO.Refresh
    Call SELECHO("ACREDOR")
    NC% = XVALO(VALACT1$("ACREDOR"))
    If NC% < 1 Then
        GoTo 999
    End If
    '
    FIMPUPRO.Label1.Caption = Str$(NC%)
    FIMPUPRO.Label2.Caption = RASOCLI$((-1) * NC%)
    FIMPUPRO.Label3.Caption = DOMICLI$((-1) * NC%)
    FIMPUPRO.Label6.Caption = CPOSCLI$((-1) * NC%)
    FIMPUPRO.Label12.Caption = LOCACLI$((-1) * NC%)
    FIMPUPRO.Refresh
    '
    Screen.MousePointer = 11
    '
100 FIMPUPRO.List2.Clear
    FIMPUPRO.List2.Refresh
    Call BLANARCH("!SELDOCAP")
    '
    MAXDEU& = ULTREG&("CUECORR")
    RFD$ = RECFILT$("CUECORR", 4, MKI$(NC%))
    '
    RESCUPRO.LINDICE.Clear
    For U& = 1 To Len(RFD$) Step 4
        REG& = CVS(Mid$(RFD$, U&, 4))
        x$ = REGLEIDO$("CUECORR", REG&)
        If CVI(Mid$(x$, 5, 2)) <> NC% Then
            x$ = REGLEIDO$("CUECORR", 1)
            Call GRAREG("CUECORR", x$, 1)
            Call CIERRARCH("CUECORR")
            GoTo 100
        End If
        '
        FEMI = CVI(Left$(x$, 2))
        IMPO# = CVD(Mid$(x$, 115, 8)) - CVD(Mid$(x$, 123, 8))
        SALDEU$ = Mid$(x$, 139, 60)
        For KKI% = 1 To 6
           FEVE = CVI(Mid$(SALDEU$, 10 * KKI% - 9, 2))
           SALDO# = CVD(Mid$(SALDEU$, 10 * KKI% - 7, 8))
           If Abs(SALDO#) >= 0.005 Then
              If MODORPA% = 1 Then
                  FF% = FEMI      ' orden emision
                Else
                  FF% = FEVE      ' orden vencimiento
              End If
              z$ = CSTRING$(MKI$(FF%), 1, 6, 0, 2) + CSTRING$(MKS$(REG&), 4, 8, 0, 2) + CSTRING$(MKI$(KKI%), 1, 4, 0, 2)
              RESCUPRO.LINDICE.AddItem z$
           End If
        Next KKI%
    Next U&
    '
    SUMADEU# = 0: SUMAFAC# = 0: KQ& = 0
    For U& = 1 To RESCUPRO.LINDICE.ListCount
        RESCUPRO.LINDICE.ListIndex = U& - 1
        REG& = XVALO(Mid$(RESCUPRO.LINDICE.TEXT, 7, 8))
        KKI% = XVALO(Mid$(RESCUPRO.LINDICE.TEXT, 15, 4))
        If REG& > 0 Then
          If REG& <= MAXDEU& Then
            If KKI% > 0 Then
              If KKI% <= 6 Then
                x$ = REGLEIDO$("CUECORR", REG&)
                TIX$ = Mid$(x$, 7, 2)
                SG% = (-1)
                TIPOS$ = "Factura": If TIX$ = "AS" Then TIPOS$ = "Ajuste de Saldo"
                                    If TIX$ = "TF" Then TIPOS$ = "Ticket-Factura"
                                    If TIX$ = "RH" Then TIPOS$ = "Rec.Honorarios"
                                    If TIX$ = "ND" Then TIPOS$ = "Nota de Debito"
                                    If TIX$ = "CP" Then TIPOS$ = "Crédito Propio"
                                    If TIX$ = "CO" Then TIPOS$ = "Comprobante"
                                    If TIX$ = "OP" Then TIPOS$ = "Orden de Pago": SG% = 1
                                    If TIX$ = "FF" Then TIPOS$ = "O/P Fondo Fijo": SG% = 1
                                    If TIX$ = "NC" Then TIPOS$ = "Nota de Credito": SG% = 1
                                    If TIX$ = "DP" Then TIPOS$ = "Debito Propio": SG% = 1
                                    If TIX$ = "CC" Then TIPOS$ = "Comp.Cuenta": SG% = 1
                                    '
                IMPO# = CVD(Mid$(x$, 115, 8)) - CVD(Mid$(x$, 123, 8))
                SALDO# = (Abs(CVD(Mid$(x$, 139 + 10 * KKI% - 8, 8)))) * Sgn(IMPO#)
                '
                IMPTEX$ = CSTRING$(MKD$(Abs(IMPO#)), 3, 13, 2, 2)
                SALTEX$ = CSTRING$(MKD$(Abs(SALDO#)), 3, 13, 2, 2)
                REGTEX$ = CSTRING$(MKS$(REG&), 3, 8, 0, 2)
                IRETEX$ = CSTRING$(MKI$(KKI%), 1, 2, 0, 2)
                '
                NROX$ = Mid$(x$, 7, 18)
                FEMI = CVI(Mid$(x$, 1, 2))
                '
                If SALDO# < (-0.005) Then
                    '
                    SUMAFAC# = SUMAFAC# + IMPO#
                    IMPO# = Abs(IMPO#)
                    SUMADEU# = SUMADEU# + SALDO#
                    '
                    FEVE = CVI(Mid$(x$, 139 + 10 * KKI% - 10, 2))
                    '
                    z$ = Space$(128)
                    Call REPLA(z$, NROX$, 1, 18)
                    Call REPLA(z$, FECHATEX$(FEMI), 23, 8)
                    Call REPLA(z$, IMPTEX$, 32, 13)
                    Call REPLA(z$, FECHATEX$(FEVE), 47, 8)
                    Call REPLA(z$, SALTEX$, 57, 13)
                    Call REPLA(z$, REGTEX$, 119, 8)
                    Call REPLA(z$, IRETEX$, 127, 2)
                    '
                    FIMPUPRO.List2.AddItem z$
                    '
                  ElseIf SALDO# >= 0.005 Then
                    '
                    z$ = Space$(60) + MKS$(0)
                    Call REPLA(z$, NROX$, 1, 18)
                    Call REPLA(z$, FECHATEX$(FEMI), 19, 8)
                    Call REPLA(z$, IMPTEX$, 27, 14)
                    Call REPLA(z$, SALTEX$, 41, 14)
                    Call REPLA(z$, MKS$(REG&), 61, 4)
                    Call REPLA(z$, MKI$(KKI%), 59, 2)
                    '
                    Call REGAPP("!SELDOCAP", z$)
                    KQ& = KQ& + 1
                    '
                End If
              End If
            End If
          End If
        End If
    Next U&
    '
    Call CIERRARCH("!SELDOCAP")
    '
    FIMPUPRO.List2.ListIndex = -1
    FIMPUPRO.List2.Refresh
    FIMPUPRO.Label8.Caption = CSTRING$(MKD$(Abs(SUMADEU#)), 4, 14, 2, 2)
    FIMPUPRO.Label8.Refresh
    FIMPUPRO.Label10.Caption = ""
    FIMPUPRO.Label10.Refresh
    FIMPUPRO.List2.Visible = True
    On Error Resume Next
    FIMPUPRO.List2.SetFocus
    If FIMPUPRO.List2.ListCount > 0 Then
        VACON% = 1
        FIMPUPRO.List2.ListIndex = 0
        VACON% = 0
    End If
    On Error GoTo 0
    '
    Screen.MousePointer = 1
    If KQ& < 1 Then
        Call COMUNI("No Aparecen Créditos no Imputados\para la Cuenta Elegida.")
        GoTo 10
    End If
    '
    If FIMPUPRO.List2.ListCount < 1 Then
        Call COMUNI("No Aparecen Obligaciones no Canceladas\en la Cuenta Elegida.")
        GoTo 10
    End If
    '
    RGSEL$ = REGSEL$("!SELDOCAP")
    If Len(RGSEL$) <= 4 Then
        GoTo 10
    End If
    '
    FIMPUPRO.Label25.Caption = Left$(RGSEL$, 18)
    FIMPUPRO.Label24.Caption = Mid$(RGSEL$, 19, 8)
    FIMPUPRO.Label23.Caption = Mid$(RGSEL$, 27, 14)
    FIMPUPRO.Label22.Caption = Mid$(RGSEL$, 41, 14)
    FIMPUPRO.Label15.Caption = CVS(Mid$(RGSEL$, 61, 4))
    FIMPUPRO.Label16.Caption = CVI(Mid$(RGSEL$, 59, 2))
    '
    FIMPUPRO.Label14.Caption = Mid$(RGSEL$, 41, 14)
    FIMPUPRO.Label20.Caption = Mid$(RGSEL$, 41, 14)
    FIMPUPRO.Command5.Enabled = True
    '
    Exit Sub
    '
999 Call CIERRARCH("*.*")
    On Error Resume Next
    FIMPUPRO.Hide
    FORPASTA.Show
    FORFACOM.Show
    MENUACRE.Show
    On Error GoTo 0
    '
End Sub
'

Sub IMPUPRO2()
    '
    CANCE# = CDbl(XVALO(FIMPUPRO.Label18.Caption))
    If CANCE# < 0.005 Then
        Exit Sub
    End If
    '
    REGCUEC& = XVALO(FIMPUPRO.Label15.Caption)
    ORDCUEC% = XVALO(FIMPUPRO.Label16.Caption)
    If REGCUEC& > 0 Then
      If REGCUEC& <= ULTREG&("CUECORR") Then
        If ORDCUEC% > 0 Then
          If ORDCUEC% <= 6 Then
            x$ = REGLEIDO$("CUECORR", REGCUEC&)
            SALPE# = Abs(CVD(Mid$(x$, 139 + 10 * ORDCUEC% - 8, 8))) - Abs(CANCE#)
            Call REPLA(x$, MKD$((-1) * SALPE#), 139 + 10 * ORDCUEC% - 8, 8)
            Call GRAREG("CUECORR", x$, REGCUEC&)
          End If
        End If
      End If
    End If
    '
    VACON% = 1
    For U& = 1 To FIMPUPRO.List2.ListCount
        FIMPUPRO.List2.ListIndex = U& - 1
        x$ = FIMPUPRO.List2.TEXT
        CANCE# = CDbl(XVALO(Mid$(x$, 70, 14)))
        If CANCE# >= 0.005 Then
            NREG& = XVALO(Mid$(x$, 121, 6))
            If NREG& > 0 Then
                ORD% = XVALO(Mid$(x$, 127, 2))
                If ORD% > 0 Then
                   If ORD% <= 6 Then
                      '
                      SS$ = REGLEIDO$("CUECORR", NREG&)
                      DOCORI$ = Mid$(SS$, 7, 18)
                      ITOTORI# = Abs(CVD(Mid$(SS$, 115, 8)) - CVD(Mid$(SS$, 123, 8)))
                      IVAORI# = 0
                      SALDO$ = Mid$(SS$, 139, 60)
                      SALDEU# = Abs(CVD(Mid$(SALDO$, 10 * ORD% - 7, 8)))
                      SALDEU# = SALDEU# - CANCE#
                      Call REPLA$(SALDO$, MKD$(SALDEU#), 10 * ORD% - 7, 8)
                      '
                      If Abs(SALDEU#) < 0.005 Then
                        Call REPLA(SALDO$, String$(10, 0), 10 * ORD% - 9, 10)
                      End If
                      '
                      Call REPLA(SS$, SALDO$, 139, 60)
                      Call GRAREG("CUECORR", SS$, NREG&)
                      '
                      NCI% = XVALO(FIMPUPRO.Label1.Caption)
                      HOII% = HOY(HO$)
                      RefApli$ = "Aplicacion Manual " + HO$
                      FEDOAPLI% = FECHANUM(TRIM$(FIMPUPRO.Label24.Caption))
                      DOCAPLI$ = TRIM$(FIMPUPRO.Label25.Caption)
                      '
                      z$ = REGBLAN$("APLIPAG")
                      Call REPLA(z$, MKI$(FEDOAPLI%), 1, 2)
                      Call REPLA(z$, MKI$(NCI%), 3, 2)
                      Call REPLA(z$, DOCAPLI$, 5, 18)
                      Call REPLA(z$, MKD$(CANCE#), 23, 8)
                      Call REPLA(z$, RefApli$, 31, 24)
                      Call REPLA(z$, DOCORI$, 55, 18)
                      Call REPLA(z$, MKD$(ITOTORI#), 73, 8)
                      Call REPLA(z$, MKD$(IVAORI#), 81, 8)
                      '
                      Call REGAPP("APLIPAG", z$)
                      '
                   End If
                End If
            End If
        End If
    Next U&
    VACON% = 0
    '
    Call CIERRARCH("APLIPAG")
    Call CIERRARCH("*.*")
    '
End Sub
'







Sub AJUSALP()
    '
    Dim FVIK%(10), SD#(10), IMC#(10)
    HOII% = HOY(HO$)
    FAJUI% = HOII%
    '
5   VDEF$ = MKI$(NC%) + MKI$(FAJUI%)
    Screen.MousePointer = 1
    CLSS$ = OBJPLA$("SELPROAJU", VDEF$)
    If CLSS$ = "" Or CLSS$ = "<ESC>" Then GoTo 99
    NC% = CVI(CLSS$)
    If NC% < 1 Then GoTo 5
    CLI$ = RASOCLI$((-1) * NC%)
    FAJUI% = CVI(Mid$(CLSS$, 3, 2))
    If FAJUI% < 1 Or FAJUI% > HOII% Then GoTo 5
    '
    FECHALIM% = XVALO(CONTROL$("", "CIEMECOM"))
    If FAJUI% <= FECHALIM% Then
        Call MENSERR(24, "Fecha no Válida.\  \Ese Período Mensual Está Cerrado.")
        GoTo 5
    End If
    '
    Screen.MousePointer = 11
    '
    SARRASFE# = 0: SALFINAL# = 0: SALDEUTO# = 0
    NUSAL& = ULTREG&("CUECORR")
    For I& = 1 To NUSAL&
       Call TRACE(20, I&, NUSAL&)
       SS0$ = REGLEIDO$("CUECORR", I&)
       SS1$ = Left$(SS0$, 2)
       SS4$ = Mid$(SS0$, 5, 2)
       FF% = CVI(SS1$)
       '
       If CVI(SS4$) = NC% Then
          SD1$ = Mid$(SS0$, 115, 8)
          SH1$ = Mid$(SS0$, 123, 8)
          SALFINAL# = SALFINAL# + (CVD(SH1$) - CVD(SD1$))
          If FF% <= FAJUI% Then
             SARRASFE# = SARRASFE# + (CVD(SH1$) - CVD(SD1$))
          End If
          '
          SS8$ = Mid$(SS0$, 139, 60)
          For V% = 1 To 6
             FVI% = CVI(Mid$(SS8$, 10 * V% - 9, 2))
             If FVI% > 0 Then
                SALDEU# = CVD(Mid$(SS8$, 10 * V% - 7, 8))
                If Abs(SALDEU#) >= 0.005 Then
                   SALDEUTO# = SALDEUTO# + SALDEU#
                End If
             End If
          Next V%
          '
       End If
    Next I&
    '
100 SALAJU# = SARRASFE#
    VDEF$ = CLSS$ + MKD$(SARRASFE#) + MKD$(SALAJU#) + MKI$(0)
102 CLSS$ = OBJPLA$("CARAJUPRO", VDEF$)
    If CLSS$ = "" Or CLSS$ = "<ESC>" Then GoTo 5
    SALAJU# = CVD(Mid$(CLSS$, 13, 8))
    ICONI% = CVI(Mid$(CLSS$, 21, 2))
    If ICONI% < 1 Or ECOARCH$("CUECON", MKI$(ICONI%)) = "" Then
      Call MENSERR(24, "Falta Imputación Contable Ajuste")
      VDEF$ = CLSS$
      GoTo 102
    End If
    If ICONI% = CUECOCLI%((-1) * NC%) Then
      Call MENSERR(24, "Cuenta Contable no Válida")
      VDEF$ = CLSS$
      GoTo 102
    End If
    '
    DIFESAL# = SALAJU# - SARRASFE#
    DIFEDEU# = SALFINAL# - SALDEUTO# + DIFESAL#
    '
    Call BLOQARCH("CUECORR")
    Call BLOQARCH("PROVED2")
    If Abs(DIFESAL#) >= 0.005 Or Abs(DIFEDEU#) >= 0.005 Then
       TOFAC#(1) = DIFESAL#
       TOFAC#(0) = (-1) * DIFESAL#
       FVIK%(1) = HOII%: SD#(1) = DIFEDEU#
       Call CUECORR(FAJUI%, (-1) * NC%, 0, 1, 0, 0&, "AJ.SALDO " + HO$ + "-" + Time$, "$", TOFAC#(), FVIK%(), SD#(), GRU%(), IMC#(), 1, FAJUI%)
    End If
    Call CIERRARCH("CUECORR")
    '
    URECOBRA& = ULTREG&("CUECORR")
    '
    If InStr(UCase$(App.EXEName), "GES") < 1 Then
      GoTo 200
    End If
    '
    ' grabar AJUSTE en CAJABANC
    FEX = FAJUI%
    FETX$ = FECHATEX$(FEX)
    NRECIB$ = "AS-" + Mid$(FETX$, 7, 2) + Mid$(FETX$, 4, 2) + Mid$(FETX$, 1, 2) + "-"
    ORMA% = (-1)
    RFF$ = RECFILT$("CAJABANC", 3, MKI$(FAJUI%))
    For I& = 1 To Len(RFF$) Step 4
      RECAJ& = CVS(Mid$(RFF$, I&, 4))
      XY$ = REGLEIDO$("CAJABANC", RECAJ&)
      If Left$(XY$, 10) = NRECIB$ Then
        ORMB% = XVALO(Mid$(XY$, 11, 2))
        If ORMB% > ORMA% Then ORMA% = ORMB%
      End If
    Next I&
    ORMX$ = TRIM$(Str$(ORMA% + 1))
    While Len(ORMX$) < 2: ORMX$ = "0" + ORMX$: Wend
    NRECIB$ = NRECIB$ + Left$(ORMX$, 2)
    '
    z$ = REGBLAN$("CAJABANC")
    Call REPLA(z$, NRECIB$, 1, 12)
    Call REPLA(z$, RASOCLI$((-1) * NC%), 13, 48)
    Call REPLA(z$, MKI$(FAJUI%), 61, 2)
    Call REPLA(z$, MKI$(FAJUI%), 63, 2)
    Call REPLA(z$, MKI$(CUECOCLI%((-1) * NC%)), 65, 2)
    Call REPLA(z$, MKI$((-1) * NC%), 67, 2)
    Call REPLA(z$, "AS", 69, 2)
    Call REPLA(z$, "Ajuste Saldo " + NRECIB$, 71, 40)
    Call REPLA(z$, MKI$(FAJUI%), 111, 2)
    Call REPLA(z$, MKI$(0), 115, 2)
    Call REPLA(z$, MKS$(URECOBRA&), 117, 4)
    Call REPLA(z$, Chr$(0), 121, 1)
    Call REPLA(z$, MKD$(DIFESAL#), 145, 8)
    If DIFESAL# < 0 Then
        Call REPLA(z$, MKD$(Abs(DIFESAL#)), 153, 8)
      Else
        Call REPLA(z$, MKD$(Abs(DIFESAL#)), 161, 8)
    End If
    Call REPLA(z$, RASOCLI$((-1) * NC%), 177, 32)
    Call REPLA(z$, "CUECORR", 245, 8)
    Call REPLA(z$, MKS$(URECOBRA&), 253, 4)
    Call REPLA(z$, IDACTU$, 257, 6)  ' FECHA - HORA - USUARIO - TERMINAL
    Call REGAPP("CAJABANC", z$)
    '
    Call REPLA(z$, "AJ.Saldo " + RASOCLI$((-1) * NC%), 13, 48)
    Call REPLA(z$, MKI$(ICONI%), 65, 2)
    Call REPLA(z$, String$(16, 0), 153, 16)
    If DIFESAL# < 0 Then
        Call REPLA(z$, MKD$(Abs(DIFESAL#)), 161, 8)
      Else
        Call REPLA(z$, MKD$(Abs(DIFESAL#)), 153, 8)
    End If
    Call REPLA(z$, "Aj.Saldo " + RASOCLI$((-1) * NC%), 177, 32)
    Call REGAPP("CAJABANC", z$)
    Call CIERRARCH("CAJABANC")
    '
200 SALFINAL# = SALFINAL# + DIFESAL#
    RGX& = REGICLI&((-1) * NC%)
    If RGX& > 0 Then
       If RGX& <= ULTREG&("PROVED2") Then
          DX$ = REGLEIDO$("PROVED2", RGX&)
          DX$ = REPLACE$(DX$, MKD$(SALFINAL#), 109, 8)
          Call GRAREG("PROVED2", DX$, RGX&)
       End If
    End If
    '
    Call CIERRARCH("CUECORR")
    Call CIERRARCH("PROVED2")
    GoTo 5
    '
99  Call CIERRARCH("*.*")
    '
End Sub
'


