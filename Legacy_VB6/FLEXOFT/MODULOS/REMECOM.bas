Attribute VB_Name = "REMECOM"
Public CARGAMEN%
'
Dim CORRESP$, ABREMES$, DES%, HAS%, DESFE$, HASFE$
Dim TICUE%(32767), SALINI#(32767), SUMFAC#(32767), SUMCOB#(32767)
Dim SUMAFAC#(32)

Dim TIS$(10)
Dim COPRO$(128)
Dim CUEV%(8), IMXX#(8)
Dim CUEIVA1%, CUEADU1%, CUEACRE1%, YACUE%
Dim CUEIVA$, CUEADU$, CUEACRE$
Dim CUEASI%(512), IMASI#(512), CASDIAR%
Dim GRUCOD%(1024), GRUCLI%(1024), CARUCO%
Dim GRUCOVE%(256), CAGRUCOVE%
Dim GRUPO%(1024), NUGRUPO%
Dim GRU%(50), IPAR#(50)
Dim CUECOD%(32768)
Dim GRUCOC%(256), CAGRUCOCLI%
Dim GRCOCLI%(32768), CUEMA%(32768), RACLI$(32768)
'
Dim NCV%(8192), TIJ%(8192), NROV&(8192), FFII%(8192) ' dimensionamiento
Dim PRIMOSTO&(8192), ULTMOSTO&(8192)
Dim IFAC#(8192), IDET#(8192), SIFAC%(8192)
Dim NRECI&(8192), REGFA$(8192)
Dim X0 As String * 96
'
Dim PORCO(256), TOCOMI#(256), TOCOB#(256)
Dim SUPED#(256), COMPED#(256), SUCOB#(256), COMCOB#(256)
'
Dim TETASA$(10), VALTASA(10), CUEIVAX%(10), INEGRAVA#(10), IGRAIVA#(10)
Dim COPROVIN$(64), DEPROVIN$(64), CUEPROVIN%(64), CAPROVI%
Dim CUEPERIVA%, CUEPERGAN%
Public YALIQUI%
'
Dim PERBRUY#(32), COPROVY$(32), DEPROVY$(32)
'
Dim VECREIVA$(256), UVECRE%
'
Sub CALRECOM()
     '
     REMECOMP.Frame8.Visible = True
     REMECOMP.Frame8.Refresh
     For XI% = 0 To 7: REMECOMP.Check1(XI%).Value = 0: Next XI%
     '
     CCCXXX$ = CUECONTA$(-1)   ' para forzar que refresque el vector
     '
10   REMECOMP.Check1(0).Value = 2
     Call VALICUEIVACOMP
     Screen.MousePointer = 11
     If YACUE% <> 1 Then
        REMECOMP.Frame8.Visible = False
        On Error Resume Next
        REMECOMP.Text2(0).SetFocus
        On Error GoTo 0
        Screen.MousePointer = 1
        Exit Sub
     End If
     '
     Screen.MousePointer = 11
     For U& = 1 To ULTREG&("PROVINB")
       If U& < 65 Then
         X$ = REGLEIDO$("PROVINB", U&)
         CAPROVI% = U&
         COPROVIN$(U&) = Left$(X$, 1)
         DEPROVIN$(U&) = Mid$(X$, 2, 25)
         CUEPROVIN%(U&) = CVI(Mid$(X$, 27, 2))
       End If
     Next U&
     '
     URECORR& = ULTREG&("CUECORR")
     URECOSS& = ULTREG&("CUECOSS")
     For U& = URECOSS& + 1 To URECORR&
       Call TRACE(20, U&, URECORR)
       U1& = U&
       X$ = REGLEIDO$("CUECORR", U1&)
       If U1& = U& Then
         X$ = X$ + String$(212, 0)
         Call GRAREG("CUECOSS", X$, U&)
       End If
     Next U&
     '
     For IXX% = 1 To 8
       TXXX$ = TRIM$(Str$(IXX% - 1))
       If IXX% < 5 Then
          TASA = Val(CONTROL$("", "TASIVA-" + TXXX$))
          If TASA > 0.005 Then
            VALTASA(IXX%) = TASA
            TETASA$(IXX%) = TRIM$(CSTRING$(MKS$(TASA), 3, 6, 2, 2) + "%")
          End If
       End If
       CUENTAI$ = CONTROL$("", "CUEIVA-" + TXXX$)
       If IXX% < 5 Then CUEIVAX%(IXX%) = CUEINT%(AJUSTI$(CUENTAI$, 12))
       If IXX% = 5 Then CUEIVAX%(6) = CUEINT%(AJUSTI$(CUENTAI$, 12))
       If IXX% = 6 Then CUEPERIVA% = CUEINT%(AJUSTI$(CUENTAI$, 12))
       If IXX% = 7 Then CUEPERGAN% = CUEINT%(AJUSTI$(CUENTAI$, 12))
     Next IXX%
     TETASA$(5) = "Exento": VALTASA(5) = 0
     TETASA$(6) = "Aduana": VALTASA(6) = 0
     TETASA$(7) = "Monotr": VALTASA(7) = 0
     TETASA$(8) = "I.Int.": VALTASA(8) = 0
     '
     TETASA$(1) = "19/21%"
     TETASA$(2) = "9/10,5%"
     '
     For IXX% = 1 To 8
       REMECOMP.TASIVA(IXX%).Caption = TETASA(IXX%)
       REMECOMP.IMPOGRA(IXX%).Caption = "0.00"
       REMECOMP.IMPOIVA(IXX%).Caption = "0.00"
       REMECOMP.TAREAL(IXX%).Caption = "0.00"
     Next IXX%
     REMECOMP.TASIVA(0).Caption = "Total"
     REMECOMP.PERIVA.Caption = "0.00"
     REMECOMP.PERGAN.Caption = "0.00"
     REMECOMP.PERIBRU.Caption = "0.00"
     '
     REMECOMP.Check1(0).Value = 1
     '
20   Call ARMAMESCOMP
     '
     Call BACKUPCOM(1, HAYERRO%)
     If HAYERRO% > 0 Then
        REMECOMP.Frame8.Visible = False
        REMECOMP.Text2(0).SetFocus
        Screen.MousePointer = 1
        Exit Sub
     End If
     '
30   REMECOMP.Check1(1).Value = 2
     If YACUE% <> 1 Then
        REMECOMP.Frame8.Visible = False
        REMECOMP.Text2(0).SetFocus
        Screen.MousePointer = 1
        Exit Sub
     End If
     FIN& = ULTREG&("PROVED1")
     For U& = 1 To FIN&
        Call TRACE(20, U&, FIN&)
        PP1$ = REGLEIDO$("PROVED1", U&)
        NC% = CVI(Left$(PP1$, 2))
        If NC% > 0 Then
           PP2$ = REGLEIDO$("PROVED2", U&)
           TICUE%(NC%) = Asc(Mid$(PP2$, 119, 1))
           If TICUE%(NC%) < 0 Or TICUE%(NC%) > 1 Then
              Call MENSERR(24, "Tipo de Cuenta no Válido\para Cuenta " + TRIM$(Str$(NC%)) + " " + TRIM$(Mid$(PP2$, 3, 30)) + ".")
              TICUE%(NC%) = 0
           End If
        End If
     Next U&
     '
     NUGRUPO% = 0: V& = 0
     FIN& = ULTREG&("CUECON")
     For U& = 1 To FIN&
        PP1$ = REGLEIDO$("CUECON", U&)
        If CVI(Left$(PP1$, 2)) > 0 Then
           NUGRUPO% = NUGRUPO% + 1
           GRUPO%(NUGRUPO%) = CVI(Left$(PP1$, 2))
           V& = V& + 1
           Call GRAREG("CUECON", PP1$, V&)
        End If
     Next U&
     Call SETULTREG("CUECON", V&)
     Call CIERRARCH("CUECON")
     '
     REMECOMP.Check1(1).Value = 1
     '
     REMECOMP.Check1(2).Value = 2
     On Error Resume Next
     For KI% = 0 To 31
         REMECOMP.Text1(KI%).TEXT = ""
         REMECOMP.Text1(KI%).Refresh
         SUMAFAC#(KI%) = 0
     Next KI%
     On Error GoTo 0
     '
     Call BLANARCH("SUIVACO")
     Call BLANARCH("LIVACOM")
     Call BLANARCH("ANAICOM")
     Call BLANARCH("SUDICOM")
     '
     Call BLOQARCH("SUIVACO")
     Call BLOQARCH("LIVACOM")
     Call BLOQARCH("ANAICOM")
     Call BLOQARCH("SUDICOM")
     Call BLOQARCH("REIVACOM")
     '
     REMECOMP.LISORT.Clear
     CASDIAR% = 0
     '
     Screen.MousePointer = 11
     FIN& = ULTREG&("CUECORR")
     J% = 0: PRIREG& = 0: ULTIRE& = 0
     PUNREG$ = ""
     For U& = 1 To 32767
        SALINI#(U&) = 0
        SUMFAC#(U&) = 0
        SUMCOB#(U&) = 0
     Next U&
     '
Call COPYSTRU("CUECORR", "CUECOR1")
Call COPYSTRU("CUECOSS", "CUECOS1")
     '
     For IX& = 1 To FIN&
        Call TRACE(20, IX&, FIN&)
35      SS$ = REGLEIDO$("CUECORR", IX&)
        TS$ = REGLEIDO$("CUECOSS", IX&)
        '
' RECUPERACIÓN DE REGISTROS DE BELLMOR FIN DE SETIEMBRE 2006
If IX& = 2350 Or IX& = 2351 Or IX& = 2352 Or IX& = 2353 Or IX& = 2354 Or IX& = 5166 Then
  If EXISTE%(LUDAT$ + "CUECOSS1.DAT") > 0 Then
    Call COPYSTRU("CUECOSS", "CUECOSS1")
    If IX& <= ULTREG&("CUECOSS1") Then
      TS$ = REGLEIDO$("CUECOSS1", IX&)
      SS$ = Left$(TS$, 300)
      Call GRAREG("CUECORR", SS$, IX&)
      Call GRAREG("CUECOSS", TS$, IX&)
    End If
  End If
End If
' FIN RECUPERACION BELLMOR
        '
        If Left$(TS$, 300) <> SS$ Then
           Call REPLA(TS$, SS$, 1, 300)
           Call GRAREG("CUECOSS", TS$, IX&)
        End If
        '
        SS1$ = Left$(SS$, 2)
        SS4$ = Mid$(SS$, 5, 2)
        SS6$ = Mid$(SS$, 7, 18)
        SS7$ = Mid$(SS$, 51, 64)
        SD1$ = Mid$(SS$, 115, 8)
        SH1$ = Mid$(SS$, 123, 8)
        Y6$ = Mid$(SS$, 199, 80)
        TT7$ = Mid$(TS$, 301, 192)
        '
        NC% = CVI(Mid$(SS$, 5, 2))
        FEMI% = CVI(SS1$)
        '
TI$ = Left$(SS6$, 2)
  If TI$ = "AS" Then
    If FEMI% >= DES% Then
      If FEMI% <= DES% + 31 Then
        IAJU# = Abs(CVD(SD1$) - CVD(SH1$))
        IIMPU# = Abs(CVD(Mid$(Y6$, 3, 8)))
        If CVI(Left$(Y6$, 2)) < 1 Or Abs(IAJU# - IIMPU#) >= 0.005 Then
40        VDEF$ = SS4$ + MKI$(FEMI%) + MKD$(Abs(IAJU#)) + MKI$(0)
          OBX$ = OBJPLA$("ICOCOM", VDEF$)
          If OBX$ = "" Then
            Screen.MousePointer = 1
            Exit Sub
          End If
          GRUI% = CVI(Mid$(OBX$, 13, 4))
          If ECOARCH$("CUECON", MKI$(GRUI%)) = "" Then GoTo 40
          Y6$ = MKI$(GRUI%) + MKD$(Abs(IAJU#)) + String$(70, 0)
          Call REPLA(SS$, Y6$, 199, 80)
          Call GRAREG("CUECORR", SS$, IX&)
          Call REPLA(TS$, Y6$, 199, 80)
          Call GRAREG("CUECOSS", TS$, IX&)
          GoTo 35
        End If
      End If
    End If
  End If
  '
        If TICUE%(NC%) <> 0 Then
           If FEMI% >= DES% Then
             If FEMI% <= DES% + 31 Then
               If PRIREG& = 0 Then PRIREG& = IX&
               If IX& > ULTIRE& Then ULTIRE& = IX&
             End If
           End If
           GoTo 1009 ' saltea cuentas de gastos
        End If
        '
        TOT# = CVD(SH1$) - CVD(SD1$)
        '
        If FEMI% < DES% Then
           SUMAFAC#(0) = SUMAFAC#(0) + TOT#
           SUMAFAC#(7) = SUMAFAC#(7) + TOT#
           REMECOMP.Text1(0).TEXT = CSTRING(MKD$(SUMAFAC#(0)), 4, 11, 2, 2)
           REMECOMP.Text1(0).Refresh
           REMECOMP.Text1(7).TEXT = CSTRING(MKD$(SUMAFAC#(7)), 4, 11, 2, 2)
           REMECOMP.Text1(7).Refresh
           SALINI#(NC%) = SALINI#(NC%) + TOT#
        End If
        '
        If FEMI% >= DES% Then
           If FEMI% <= DES% + 31 Then
              If PRIREG& = 0 Then
                 PRIREG& = IX&
                 REMECOMP.Check1(2).Value = 1
                 REMECOMP.Check1(3).Value = 2
              End If
              If IX& > ULTIRE& Then ULTIRE& = IX&
              NUFAC& = IX&
              '
              If UCase$(Mid$(SS$, 12, 4)) = "ANUL" Then
                 Call REPLA(SS$, String$(228, 0), 51, 228)
                 Call GRAREG("CUECORR", SS$, IX&)
                 Call REPLA(TS$, String$(228, 0), 51, 228)
                 Call REPLA(TS$, String$(212, 0), 301, 212)
                 Call GRAREG("CUECOSS", TS$, IX&)
                 GoTo 1009
              End If
              '
              TI$ = Left$(SS6$, 2)
              '
              ICAMPO% = 1: If TOT# < 0 Then ICAMPO% = 3
              If TI$ = "ND" Then ICAMPO% = 2
              If TI$ = "OP" Then ICAMPO% = 5
              If TI$ = "FF" Then ICAMPO% = 5
              If TI$ = "AS" Then ICAMPO% = 6
              If TI$ = "AJ" Then ICAMPO% = 6
              If TI$ = "CO" Then ICAMPO% = 1
              '
              SUMAFAC#(ICAMPO%) = SUMAFAC#(ICAMPO%) + TOT#
              If ICAMPO% < 4 Then
                 SUMAFAC#(4) = SUMAFAC#(4) + TOT#
              End If
              SUMAFAC#(7) = SUMAFAC#(7) + TOT#
              '
              If ICAMPO% <> 5 Then
                   SUMFAC#(NC%) = SUMFAC#(NC%) + TOT#
                 Else
                   SUMCOB#(NC%) = SUMCOB#(NC%) - TOT#
              End If
              '
              For ICAMPO% = 0 To 7
                 REMECOMP.Text1(ICAMPO%).TEXT = CSTRING(MKD$(SUMAFAC#(ICAMPO%)), 4, 11, 2, 2)
                 REMECOMP.Text1(ICAMPO%).Refresh
              Next ICAMPO%
              '
              If TI$ <> "OP" Then
                If TI$ <> "FF" Then
                  PUNREG$ = PUNREG$ + MKS$(IX&)
                End If
              End If
              '
              If TI$ <> "VC" Then
               If TI$ <> "FC" Then
                If TI$ <> "TF" Then
                 If TI$ <> "RH" Then
                  If TI$ <> "ND" Then
                   If TI$ <> "CO" Then
                    If TI$ <> "NC" Then
                     If TI$ <> "DP" Then
                      If TI$ <> "DC" Then
                       If TI$ <> "AS" Then
                        GoTo 1009         ' next IX&
                       End If
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
              NC1% = (-1) * CVI(SS4$)
              CLI$ = RASOCLI$(NC1%)
              CUIT$ = CUITCLI$(NC1%)
              '
              NF$ = SS6$
              TI$ = Left$(SS6$, 2)
              SIG% = 1: If TI$ = "NC" Then SIG% = (-1)
                        If TI$ = "DP" Then SIG% = (-1)
                        If TI$ = "DC" Then SIG% = (-1)
                        If TI$ = "AS" Then
                        SIG% = Sgn(CVD(SD1$) - CVD(SH1$))
                        End If
                        '
              IMGRA# = Abs(CVD(Mid$(SS7$, 1, 8)))
              IMEXE# = Abs(CVD(Mid$(SS7$, 9, 8)))
              IMONO# = 0: If PIVACLI%(NC1%) = 6 Then IMONO# = IMEXE#: IMEXE# = 0
              IMINT# = Abs(CVD(Mid$(SS7$, 17, 8)))
              IMIVA# = Abs(CVD(Mid$(SS7$, 25, 8)))
              imadu# = Abs(CVD(Mid$(SS7$, 33, 8)))
              RTIVA# = Abs(CVD(Mid$(SS7$, 41, 8)))
              RTGAN# = Abs(CVD(Mid$(SS7$, 49, 8)))
              RTIBR# = Abs(CVD(Mid$(SS7$, 57, 8)))
              TOFAC1# = 0: For KKI% = 1 To 64 Step 8
                             TOFAC1# = TOFAC1# + (Int(100 * Abs(CVD(Mid$(SS7$, KKI%, 8))) + 0.5)) / 100
                           Next KKI%
                           '
              If Abs(TOFAC1# - Abs(TOT#)) >= 0.005 Then
                 If Left$(NF$, 2) = "AS" Then
                   SS7$ = String$(64, 0)
                   Call REPLA(SS7$, MKD$(TOT#), 9, 8)
                   Call REPLA(SS$, SS7$, 51, 64)
                   Call GRAREG("CUECORR", SS$, IX&)
                   Call REPLA(TS$, SS7$, 51, 64)
                   Call REPLA(TS$, MKD$(Abs(TOT#)) + String$(184, 0), 301, 192)
                   Call GRAREG("CUECOSS", TS$, IX&)
                  Else
                   TTXX$ = "Se ha Detectado una Inconsistencia(1)\en el Comprobante Número " + TRIM$(NF$) + "\de " + TRIM$(CLI$) + "."
                   TTXX$ = TTXX$ + "\  \Este Comprobante Debe Anularse y Registrarse Nuevamente."
                   Call MENSERR(24, TTXX$)
                 End If
              End If
              '
              If TI$ <> "CO" Then
                If TI$ <> "AS" Then
                  CA$ = Mid$(NF$, 4, 1)
                  If CA$ = "C" Or CA$ = "B" Then
                    If PIVACLI%(NC1%) = 1 Then
                      Call REPLA(SS$, "A", 10, 1)
                      Call REPLA(NF$, "A", 4, 1)
                      Call GRAREG("CUECORR", SS$, IX&)
                      Call REPLA(TS$, "A", 10, 1)
                      Call GRAREG("CUECOSS", TS$, IX&)
                    End If
                  End If
                  '
                  CA$ = Mid$(NF$, 4, 1)
                  If IMIVA# + imadu# < 0.005 Then
                    If CA$ = "A" Then
                      Call COMUNI("Atención: Factura 'A' " + NF$ + "\Emitida por " + CLI$ + "\no Discrimina I.V.A.")
                      'Call REPLA(SS$, "B", 10, 1)
                      'Call REPLA(NF$, "B", 4, 1)
                      'Call GRAREG("CUECORR", SS$, IX&)
                      'Call REPLA(TS$, "B", 10, 1)
                      'Call GRAREG("CUECOSS", TS$, IX&)
                    End If
                  End If
                End If
              End If
              '
              TOIVA# = IMIVA# + IVADU#
              TONET# = IMGRA# + IMEXE# + IMONO#
              If TOIVA# < 0.005 Then
                 Call REPLA(SS$, MKD$(0) + MKD$(TONET#), 51, 16)
                 Call REPLA(TS$, MKD$(0) + MKD$(TONET#), 51, 16)
                 Call GRAREG("CUECORR", SS$, IX&)
                 Call GRAREG("CUECOSS", TS$, IX&)
                 IMEXE# = IMGRA# + IMEXE#: IMGRA# = 0
              End If
              '
              TT7$ = Mid$(TS$, 301, 192)
              If TT7$ = String$(192, 0) Then
                 For KKI% = 1 To 8
                   Call REPLA(TT7$, TETASA$(KKI%), 24 * KKI% - 7, 8)
                 Next KKI%
                 TASARE = 0
                 If IMGRA# >= 0.005 Then
                    TASARE = 100 * IMIVA# / IMGRA#
                 End If
                 If PIVACLI%(NC1%) = 6 Then
                   If Abs(IMEXE#) > 0 Then
                     IMONO# = IMONO# + IMEXE#
                     IMEXE# = 0
                   End If
                 End If
                 Call REPLA(TT7$, MKD$(IMEXE#), 24 * 5 - 23, 8)
                 Call REPLA(TT7$, MKD$(IMONO#), 24 * 7 - 23, 8)
                 Call REPLA(TT7$, MKD$(IMINT#), 24 * 8 - 23, 8)
                 '
                 For KKI% = 1 To 4
                    If Abs(TASARE - VALTASA(KKI%)) < 0.1 Then
                      Call REPLA(TT7$, MKD$(IMGRA#), 24 * KKI% - 23, 8)
                      Call REPLA(TT7$, MKD$(IMIVA#), 24 * KKI% - 15, 8)
                      GoTo 123
                    End If
                 Next KKI%
                 Call REPLA(TT7$, MKD$(IMGRA#), 24 * 6 - 23, 8)
                 Call REPLA(TT7$, MKD$(IMIVA#), 24 * 6 - 15, 8)
                 '
123              Call REPLA(TS$, TT7$, 301, 192)
                 Call GRAREG("CUECOSS", TS$, IX&)
                 '
              End If
              '
111           SUMA1# = 0: SUMA3# = 0
                          For KKI% = 1 To 5
                            SUMA1# = SUMA1# + CVD(Mid$(SS7$, 8 * KKI% - 7, 8))
                            If KKI% <= 3 Then SUMA3# = SUMA3# + CVD(Mid$(SS7$, 8 * KKI% - 7, 8))
                          Next KKI%
              SUMA2# = 0: SUMA5# = 0
                          For KKI% = 1 To 8
                            SUMA2# = SUMA2# + CVD(Mid$(TT7$, 24 * KKI% - 23, 8))
                            SUMA2# = SUMA2# + CVD(Mid$(TT7$, 24 * KKI% - 15, 8))
                            SUMA5# = SUMA5# + CVD(Mid$(TT7$, 24 * KKI% - 23, 8))
                          Next KKI%
                          
              SUMA4# = 0: For KKI% = 1 To 8
                            GRUI% = CVI(Mid$(Y6$, 10 * KKI% - 9, 2))
12444                       If GRUI% > 0 Then
                              For K1% = 1 To NUGRUPO%
                                 If GRUPO%(K1%) = GRUI% Then
                                    SUMA4# = SUMA4# + CVD(Mid$(Y6$, 10 * KKI% - 7, 8))
                                    GoTo 124
                                 End If
                              Next K1%
' esto se puso para atajar el caso de que no figure el grupo contable en la tabla actual
Call MENSERR(24, "Grupo " + TRIM$(Str$(GRUI%)) + " Inexistente en Tabla de Grupos Contables.")
GRUI% = GRUPO%(1)
GoTo 12444
                            End If
124                       Next KKI%
                          '
              If Abs(Abs(SUMA1#) - SUMA2#) >= 0.005 Then
                 'GoSub RECUMO: If RECUI% = 1 Then GoTo 35
                 MXX$ = "2"
                 GoTo 130
              End If
              '
125           If Abs(Abs(SUMA3#) - SUMA4#) >= 0.005 Then
                 'GoSub RECUMO: If RECUI% = 1 Then GoTo 35
                 MXX$ = "3"
                 GoTo 130
              End If
              '
              If Abs(SUMA5# - Abs(SUMA3#)) >= 0.005 Then
                 'GoSub RECUMO: If RECUI% = 1 Then GoTo 35
                 MXX$ = "4"
                 GoTo 130
              End If
              GoTo 140
              '
130           TTXX$ = "Se ha Detectado una Inconsistencia(" + MXX$ + ")\en el Comprobante Número " + TRIM$(NF$) + "\de " + TRIM$(CLI$) + "(" + TRIM$(Str$(Abs(NC1%))) + ")."
              TTXX$ = TTXX$ + "\  \Este Comprobante Debe Anularse y Registrarse Nuevamente."
              Call MENSERR(24, TTXX$)
              '
140           If TI$ = "AS" Then GoTo 1007
              For KKI% = 1 To 8
                INEGRAVA#(KKI%) = SIG% * Abs(CVD(Mid$(TT7$, 24 * KKI% - 23, 8)))
                VALOR# = INEGRAVA#(KKI%) + Val(REMECOMP.IMPOGRA(KKI%).Caption)
                REMECOMP.IMPOGRA(KKI%).Caption = TRIM$(CSTRING$(MKD$(VALOR#), 4, 12, 2, 2))
              Next KKI%
              '
              For KKI% = 1 To 8
                IGRAIVA#(KKI%) = SIG% * Abs(CVD(Mid$(TT7$, 24 * KKI% - 15, 8)))
                VALOR# = IGRAIVA#(KKI%) + Val(REMECOMP.IMPOIVA(KKI%).Caption)
                REMECOMP.IMPOIVA(KKI%).Caption = TRIM$(CSTRING$(MKD$(VALOR#), 4, 12, 2, 2))
              Next KKI%
              '
              PERCEPIVA# = SIG% * Abs(CVD(Mid$(SS7$, 41, 8)))
              VALOR# = PERCEPIVA# + Val(REMECOMP.PERIVA.Caption)
              REMECOMP.PERIVA.Caption = TRIM$(CSTRING$(MKD$(VALOR#), 4, 12, 2, 2))
              '
              PERCEPGAN# = SIG% * Abs(CVD(Mid$(SS7$, 49, 8)))
              VALOR# = PERCEPGAN# + Val(REMECOMP.PERGAN.Caption)
              REMECOMP.PERGAN.Caption = TRIM$(CSTRING$(MKD$(VALOR#), 4, 12, 2, 2))
              '
              PERCEPIBRU# = SIG% * Abs(CVD(Mid$(SS7$, 57, 8)))
              VALOR# = PERCEPIBRU# + Val(REMECOMP.PERIBRU.Caption)
              REMECOMP.PERIBRU.Caption = TRIM$(CSTRING$(MKD$(VALOR#), 4, 12, 2, 2))
              '
              SUMAOTROS# = INEGRAVA#(4) + INEGRAVA#(6) + IGRAIVA#(4) + IGRAIVA#(5) + IGRAIVA#(7) + PERCEPGAN# + PERCEPIBRU#
              '
9999          A0$ = REGBLAN$("SUIVACO")
              Call REPLA(A0$, MKI$(FEMI%), 1, 2)
              Call REPLA(A0$, NF$, 3, 18)
              Call REPLA(A0$, CLI$, 21, 30)
              Call REPLA(A0$, CUIT$, 51, 30)
              Call REPLA(A0$, MKD$(INEGRAVA#(1)), 65, 8)
              Call REPLA(A0$, MKD$(INEGRAVA#(2)), 73, 8)
              Call REPLA(A0$, MKD$(INEGRAVA#(3)), 81, 8)
              Call REPLA(A0$, MKD$(INEGRAVA#(5)), 89, 8)
              Call REPLA(A0$, MKD$(INEGRAVA#(7)), 97, 8)
              Call REPLA(A0$, MKD$(IGRAIVA#(1)), 105, 8)
              Call REPLA(A0$, MKD$(IGRAIVA#(2)), 113, 8)
              Call REPLA(A0$, MKD$(IGRAIVA#(3)), 121, 8)
              Call REPLA(A0$, MKD$(IGRAIVA#(6)), 129, 8)
              Call REPLA(A0$, MKD$(PERCEPIVA#), 137, 8)
              Call REPLA(A0$, MKD$(SUMAOTROS#), 145, 8)
              Call REPLA(A0$, MKD$(TOFAC1# * SIG%), 153, 8)
              '
              Call REGAPP("SUIVACO", A0$)
              '
1007          RESO$ = Space$(16)
              Call REPLA(RESO$, CSTRING$(MKI$(FEMI%), 1, 6, 0, 2), 1, 6)
              Call REPLA(RESO$, CSTRING$(MKS$(IX&), 3, 8, 0, 2), 9, 8)
              REMECOMP.LISORT.AddItem RESO$
              '
           End If
        End If
        '
1009 Next IX&
     '
     Call CIERRARCH("SUIVACO")
     '
     For UI1% = 1 To REMECOMP.LISORT.ListCount
       RESO$ = REMECOMP.LISORT.List(UI1% - 1)
       RECUE& = Val(Mid$(RESO$, 9, 8))
       TS$ = REGLEIDO$("CUECOSS", RECUE&)
       Call LIVACOM(TS$)
     Next UI1%
     Call CIERRARCH("LIVACOM")
     '
     JJK& = 0
     For UI1% = 1 To UVECRE%
       JJK& = JJK& + 1
       Call GRAREG("REIVACOM", VECREIVA$(UI1%), JJK&)
     Next UI1%
     Call SETULTREG("REIVACOM", JJK&)
     Call CIERRARCH("REIVACOM")
     '
     REMECOMP.Check1(3).Value = 1
     '
     REMECOMP.Check1(4) = 1
     '
     REMECOMP.Check1(5) = 2
     TTXX$ = Chr$(1) + Chr$(2) + MKI$(HAS%)
     V& = 0: FIN& = ULTREG&("ASIDIAR")
     For U& = 1 To FIN&
        Call TRACE(20, U&, FIN&)
        X$ = REGLEIDO$("ASIDIAR", U&)
        If Left$(X$, 4) <> TTXX$ Then
           V& = V& + 1
           Call GRAREG("ASIDIAR", X$, V&)
        End If
     Next U&
     '
     For K% = 1 To CASDIAR%
        If CUEASI%(K%) > 0 Then
           If Abs(IMASI#(K%)) >= 0.005 Then
              X$ = REGBLAN$("ASIDIAR")
              Call REPLA(X$, TTXX$, 1, 4)
              Call REPLA(X$, MKI$(CUEASI%(K%)), 5, 2)
              Call REPLA(X$, MKD$(IMASI#(K%)), 39, 8)
              V& = V& + 1
              Call GRAREG("ASIDIAR", X$, V&)
           End If
        End If
     Next K%
     '
     Call SETULTREG&("ASIDIAR", V&)
     Call CIERRARCH("ASIDIAR")
     REMECOMP.Check1(5) = 1
     '
     REMECOMP.Check1(6) = 2
     '
     Call BLANARCH("IFACOMP")
     Call BLANARCH("DETFACC")
     Call BLANARCH("DETDEBC")
     Call BLANARCH("DETCREC")
     Call BLANARCH("DETOTRC")
     Call BLANARCH("DETPAGC")
     Call BLANARCH("DETPAGG")
     Call BLANARCH("DETFOFI")
     '
     If PRIREG& > 0 Then
        FIN& = ULTIRE& - PRIREG& + 1
        REMECOMP.List1.Clear
        For IX& = PRIREG& To ULTIRE&
           Call TRACE(20, IX& - PRIREG& + 1, FIN&)
           SS$ = REGLEIDO$("CUECORR", IX&)
           FEMI% = CVI(Left$(SS$, 2))
           Z$ = Space$(16)
           Call REPLA(Z$, CSTRING$(MKI$(10000 + FEMI%), 1, 5, 0, 2), 1, 6)
           Call REPLA(Z$, Str$(IX&), 9, 6)
           REMECOMP.List1.AddItem Z$
        Next IX&
        '
        FIN& = REMECOMP.List1.ListCount
        For IZ& = 1 To FIN&
           Call TRACE(20, IZ&, FIN&)
           Z$ = REMECOMP.List1.List(IZ& - 1)
           IX& = Val(Mid$(Z$, 9, 8))
155        SS$ = REGLEIDO$("CUECORR", IX&)
           SS1$ = Left$(SS$, 2)
           SS4$ = Mid$(SS$, 5, 2)
           SS6$ = Mid$(SS$, 7, 18)
           SS7$ = Mid$(SS$, 51, 64)
           SD1$ = Mid$(SS$, 115, 8)
           SH1$ = Mid$(SS$, 123, 8)
           TOT# = CVD(SH1$) - CVD(SD1$)
           '
           FEMI% = CVI(SS1$)
           If FEMI% >= DES% Then
              If FEMI% <= DES% + 31 Then
                 NC% = CVI(SS4$)
                 TICUPRO% = TICUE%(NC%)
                 TI$ = Left$(SS6$, 2)
                 NF$ = SS6$
                 SIG% = 1: If TI$ = "NC" Then SIG% = (-1)
                           If TI$ = "DP" Then SIG% = (-1)
                           If TI$ = "DC" Then SIG% = (-1)
                           '
                 NETO# = Abs(CVD(Mid$(SS7$, 1, 8))) + Abs(CVD(Mid$(SS7$, 9, 8)))
                 IMIVA# = Abs(CVD(Mid$(SS7$, 25, 8)))
                 IVADU# = Abs(CVD(Mid$(SS7$, 33, 8)))
                 RIVA# = Abs(CVD(Mid$(SS7$, 41, 8)))
                 OTRO# = Abs(CVD(Mid$(SS7$, 17, 8))) + Abs(CVD(Mid$(SS7$, 49, 8))) + Abs(CVD(Mid$(SS7$, 57, 8)))
                 TOFAC1# = (Int(100 * NETO# + 0.5) + Int(100 * IMIVA# + 0.5) + Int(100 * IVADU# + 0.5) + Int(100 * RIVA# + 0.5) + Int(100 * OTRO# + 0.5)) / 100
                 '
                 If Mid$(NF$, 6, 4) = "anul" Then
                   NETO# = 0: IMIVA# = 0: IVADU# = 0: RIVA# = 0: OTRO# = 0: TOFAC1# = 0
                 End If
                 '
                 If Abs(Abs(TOT#) - Abs(TOFAC1#)) >= 0.005 Then
                   If Left$(NF$, 2) = "AS" Then
                       Call REPLA(SS7$, MKD$(Abs(TOT#)) + String$(56, 0), 1, 64)
                       Call REPLA(SS$, SS7$, 51, 64)
                       Call GRAREG("CUECORR", SS$, IX&)
                       GoTo 155
                     Else
                       TTXX$ = "Se ha Detectado una Inconsistencia(5)\en el Comprobante Número " + TRIM$(NF$) + "\de " + TRIM$(RASOCLI$((-1) * NC%)) + "."
                       TTXX$ = TTXX$ + "\  \Este Comprobante Debe Anularse y Registrarse Nuevamente."
                       Call MENSERR(24, TTXX$)
                   End If
                 End If
                 '
                 DFC$ = REGBLAN$("DETFACC")
                 Call REPLA(DFC$, TI$, 1, 2)
                 Call REPLA(DFC$, MKI$(FEMI%), 3, 2)
                 Call REPLA(DFC$, NF$, 5, 18)
                 Call REPLA(DFC$, MKI$(NC%), 23, 2)
                 Call REPLA(DFC$, MKD$(NETO#), 25, 8)
                 Call REPLA(DFC$, MKD$(IMIVA#), 33, 8)
                 Call REPLA(DFC$, MKD$(IVADU#), 41, 8)
                 Call REPLA(DFC$, MKD$(RIVA#), 49, 8)
                 Call REPLA(DFC$, MKD$(OTRO#), 57, 8)
                 Call REPLA(DFC$, MKD$(TOFAC1#), 65, 8)
                 '
                 DFD$ = REGBLAN$("IFACOMP")
                 Call REPLA(DFD$, NF$, 1, 18)
                 Call REPLA(DFD$, RASOCLI$((-1) * NC%), 21, 30)
                 Call REPLA(DFD$, MKD$(NETO#), 51, 8)
                 Call REPLA(DFD$, MKI$(NC%), 59, 2)
                 '
                 If TICUPRO% = 0 Then
                    If TI$ = "VC" Or TI$ = "FC" Or TI$ = "TF" Or TI$ = "RH" Or TI$ = "CO" Then
                        Call REGAPP("DETFACC", DFC$)
                        Call REGAPP("IFACOMP", DFD$)
                      ElseIf TI$ = "ND" Then
                        Call REGAPP("DETDEBC", DFC$)
                        Call REGAPP("IFACOMP", DFD$)
                     ElseIf TI$ = "NC" Or TI$ = "DP" Or TI$ = "NC" Then
                        Call REGAPP("DETCREC", DFC$)
                        Call REGAPP("IFACOMP", DFD$)
                      ElseIf TI$ = "OP" Or TI$ = "FF" Then
                        Call REGAPP("DETPAGC", DFC$)
                      Else
                        SIG% = Sgn(TOT#)
                        Call REPLA(DFC$, MKD$(NETO# * SIG%), 25, 8)
                        Call REPLA(DFC$, MKD$(IMIVA# * SIG%), 33, 8)
                        Call REPLA(DFC$, MKD$(IVADU# * SIG%), 41, 8)
                        Call REPLA(DFC$, MKD$(RIVA# * SIG%), 49, 8)
                        Call REPLA(DFC$, MKD$(OTRO# * SIG%), 57, 8)
                        Call REPLA(DFC$, MKD$(TOFAC1# * SIG%), 65, 8)
                        Call REGAPP("DETOTRC", DFC$)
                    End If
                    '
                  Else
                    If TI$ = "OP" Or TI$ = "FF" Then
                        Call REGAPP("DETPAGG", DFC$)
                    End If
                 End If
                 '
                 If TI$ = "FF" Then
                    Call REGAPP("DETFOFI", DFC$)
                 End If
                 '
              End If
           End If
           '
        Next IZ&
        '
     End If
     '
     REMECOMP.Check1(6) = 1
     '
     REMECOMP.Check1(7) = 2
     Call BLANARCH("RESUCCP")
     FIN& = ULTREG&("PROVED1")
     For U& = 1 To ULTREG&("PROVED1")
        Call TRACE(20, U&, FIN&)
        X$ = REGLEIDO$("PROVED1", U&)
        NC% = CVI(Left$(X$, 2))
        If NC% > 0 Then
           If Abs(SALINI#(NC%)) + Abs(SUMFAC#(NC%)) + Abs(SUMCOB#(NC%)) >= 0.005 Then
              Y$ = REGBLAN$("RESUCCP")
              Call REPLA(Y$, MKI$(NC%), 1, 2)
              Call REPLA(Y$, MKD$(SALINI#(NC%)), 3, 8)
              Call REPLA(Y$, MKD$(SUMFAC#(NC%)), 11, 8)
              Call REPLA(Y$, MKD$(SUMCOB#(NC%)), 19, 8)
              Call REPLA(Y$, MKD$(SALINI#(NC%) + SUMFAC#(NC%) - SUMCOB#(NC%)), 27, 8)
              Call REGAPP("RESUCCP", Y$)
           End If
        End If
     Next U&
     Call CIERRARCH("RESUCCP")
     REMECOMP.Check1(7) = 1
     '
     '
     CORRESP$ = AJUSTI$(REMECOMP.MESAN.TEXT, 12)
     FGENE% = HOY(HO$): If INCOMPLE% > 0 Then FGENE% = 0
     HGENE% = HORANUM(Time$): If INCOMPLE% > 0 Then HGENE% = 0
     '
     Call CIERRARCH("*.*")
     Call LIBARCH("*.*")
     '
     Call GRARGEN("SUIVACO", TRIM$(CORRESP$))
     '
     Call HEADERS("SUIVACO", "")
     Call HEADERS("SUIVACO", "Corresponde a: " + CORRESP$)
     Call HEADERS("LIVACOM", "")
     Call HEADERS("LIVACOM", "Corresponde a: " + CORRESP$)
     Call HEADERS("SUDICOM", "")
     Call HEADERS("SUDICOM", "Corresponde a: " + CORRESP$)
     Call HEADERS("ANAICOM", "")
     Call HEADERS("ANAICOM", "Corresponde a: " + CORRESP$)
     Call HEADERS("REIVACOM", "")
     Call HEADERS("REIVACOM", "Corresponde a: " + CORRESP$)
     Call CIERRARCH("*.*")
        '
     Call BACKUPCOM(2, HAYERRO%)
     If HAYERRO% <> 0 Then
        Call MENSERR(24, "Error al Respaldar Archivos Generados")
        GoTo 999
     End If
     '
     REBLA$ = String$(300, 0)
     Call REPLA(REBLA$, CORRESP$, 1, 12)
     Call REPLA(REBLA$, MKI$(FGENE%), 13, 2)
     Call REPLA(REBLA$, MKI$(HGENE%), 15, 2)
     '
     On Error Resume Next
     For KI% = 0 To 7
         Call REPLA(REBLA$, MKD$(Val(REMECOMP.Text1(KI%).TEXT)), 45 + 8 * KI%, 8)
         Call REPLA(REBLA$, MKD$(Val(REMECOMP.IMPOGRA(KI%).Caption)), 45 + 64 + 8 * KI%, 8)
         Call REPLA(REBLA$, MKD$(Val(REMECOMP.IMPOIVA(KI%).Caption)), 45 + 128 + 8 * KI%, 8)
     Next KI%
     Call REPLA(REBLA$, MKD$(Val(REMECOMP.PERIVA.Caption)), 45 + 192 + 8 * 0, 8)
     Call REPLA(REBLA$, MKD$(Val(REMECOMP.PERGAN.Caption)), 45 + 192 + 8 * 1, 8)
     Call REPLA(REBLA$, MKD$(Val(REMECOMP.PERIBRU.Caption)), 45 + 192 + 8 * 2, 8)
     Call REPLA(REBLA$, MKD$(Val(REMECOMP.IMPOGRA(8).Caption)), 45 + 192 + 8 * 3, 8)
     On Error GoTo 0
     '
     NX% = FILEOPEN%(LU$ + "IRESCOM.DAT", 0, 300)
     Dim REBLB As String * 300
     '
     For KI% = 1 To LOF(NX%) / 300
        Get #NX%, KI%, REBLB$
        If Left$(REBLB$, 12) = CORRESP$ Then
            REBLB$ = REBLA$
            Put #NX%, KI%, REBLB$
            GoTo 159
        End If
     Next KI%
     '
     KI% = 1 + LOF(NX%) / 300
     REBLB$ = REBLA$
     Put #NX%, KI%, REBLB$
     '
159  Close #NX%
     '
     FEX = CVI(Mid$(REBLA$, 13, 2))
     REMECOMP.Text4(0).TEXT = FECHATEX$(FEX)
     If FEX < 1 Then REMEVENT.Text4(0).TEXT = "00/00/00"
     REMECOMP.Text4(0).Refresh
     HOX% = CVI(Mid$(REBLA$, 15, 2))
     REMECOMP.Text4(1) = HORATEX$(HOX%)
     REMECOMP.Text4(1).Refresh
     '
     Call COMUNI("Cálculo Completo")
     '
999  Call CIERRARCH("*.*")
     Call LIBARCH("*.*")
     Screen.MousePointer = 1
     Exit Sub
     '

RECUMO:
'RECUI% = 0
'EPAC$ = TRIM$(UCase$(EMPREAC$))
'If InStr(EPAC$, "DOSIVA") < 1 Then GoTo 900
''
'If TRIM$(Mid$(SS$, 7, 18)) = ULTIREC$ Then GoTo 900
'ULTIREC$ = TRIM$(Mid$(SS$, 7, 18))
''
'For IX9& = IX& - 100 To FIN&
'  TS9$ = REGLEIDO$("CUECOSS", IX9&)
'  TT79$ = Mid$(TS9$, 301, 192)
'  SUMA29# = 0
'  For KKI9% = 1 To 8
'    SUMA29# = SUMA29# + CVD(Mid$(TT79$, 24 * KKI9% - 23, 8))
'    SUMA29# = SUMA29# + CVD(Mid$(TT79$, 24 * KKI9% - 15, 8))
'  Next KKI9%
'  If Abs(Abs(SUMA1#) - SUMA29#) < 0.005 Then
'    TS8$ = REGLEIDO$("CUECOSS", IX&)
'    SS8$ = REGLEIDO$("CUECORR", IX&)
'    Call REPLA(TS8$, SS8$, 1, 300)
'    Call REPLA(TS8$, TT79$, 301, 192)
'    Call GRAREG("CUECOSS", TS8$, IX&)
'    'MsgBox "Recuperado 1" & IX&
'    RECUI% = 1: GoTo 900
'  End If
'Next IX9&
''
'For IX9& = ULTREG&("CUECOS1") To 1 Step -1
'  TS9$ = REGLEIDO$("CUECOS1", IX9&)
'  If Left$(TS9$, 24) = Left$(SS$, 24) Then
'    TT79$ = Mid$(TS9$, 301, 192)
'    SUMA29# = 0:
'    For KKI9% = 1 To 8
'      SUMA29# = SUMA29# + CVD(Mid$(TT79$, 24 * KKI9% - 23, 8))
'      SUMA29# = SUMA29# + CVD(Mid$(TT79$, 24 * KKI9% - 15, 8))
'    Next KKI9%
'    If Abs(Abs(SUMA1#) - SUMA29#) < 0.005 Then
'      TS8$ = REGLEIDO$("CUECOSS", IX&)
'      SS8$ = REGLEIDO$("CUECORR", IX&)
'      Call REPLA(TS8$, SS8$, 1, 300)
'      Call REPLA(TS8$, TT79$, 301, 192)
'      Call GRAREG("CUECOSS", TS8$, IX&)
'      'MsgBox "Recuperado 2" & IX&
'      RECUI% = 1: GoTo 900
'    End If
'  End If
'Next IX9&
900 Return
'
End Sub
'

Sub CARGARESCOM()
        '
        REBLA$ = Space$(12) + String$(288, 0)
        CORRESP$ = AJUSTI$(REMECOMP.MESAN.TEXT, 12)
        NX% = FILEOPEN%(LU$ + "IRESCOM.DAT", 0, 300)
        Dim REBLB As String * 300
        '
        For KI% = 1 To LOF(NX%) / 300
            Get #NX%, KI%, REBLB$
            If Left$(REBLB$, 12) = CORRESP$ Then
                REBLA$ = REBLB$
            End If
        Next KI%
        '
        FEX = CVI(Mid$(REBLA$, 13, 2))
        REMECOMP.Text4(0).TEXT = FECHATEX$(FEX)
        If FEX < 1 Then REMECOMP.Text4(0).TEXT = "00/00/00"
        REMECOMP.Text4(0).Refresh
        HOX% = CVI(Mid$(REBLA$, 15, 2))
        REMECOMP.Text4(1) = HORATEX$(HOX%)
        REMECOMP.Text4(1).Refresh
        '
        For KI% = 0 To 7
            SUMAFAC#(KI%) = CVD(Mid$(REBLA$, 45 + 8 * KI%, 8))
            IGRA1# = CVD(Mid$(REBLA$, 45 + 64 + 8 * KI%, 8))
            IIVA1# = CVD(Mid$(REBLA$, 45 + 128 + 8 * KI%, 8))
            On Error Resume Next
            REMECOMP.Text1(KI%).TEXT = CSTRING$(MKD$(SUMAFAC#(KI%)), 4, 11, 2, 2)
            REMECOMP.Text1(KI%).Refresh
            REMECOMP.IMPOGRA(KI%).Caption = TRIM$(CSTRING$(MKD$(IGRA1#), 4, 11, 2, 2))
            REMECOMP.IMPOGRA(KI%).Refresh
            REMECOMP.IMPOIVA(KI%).Caption = TRIM$(CSTRING$(MKD$(IIVA1#), 4, 11, 2, 2))
            REMECOMP.IMPOIVA(KI%).Refresh
            On Error GoTo 0
        Next KI%
        PERIV1# = CVD(Mid$(REBLA$, 45 + 192 + 8 * 0, 8))
        REMECOMP.PERIVA.Caption = TRIM$(CSTRING$(MKD$(PERIV1#), 4, 11, 2, 2))
        REMECOMP.PERIVA.Refresh
        PERGA1# = CVD(Mid$(REBLA$, 45 + 192 + 8 * 1, 8))
        REMECOMP.PERGAN.Caption = TRIM$(CSTRING$(MKD$(PERGA1#), 4, 11, 2, 2))
        REMECOMP.PERGAN.Refresh
        PERIB1# = CVD(Mid$(REBLA$, 45 + 192 + 8 * 2, 8))
        REMECOMP.PERIBRU.Caption = TRIM$(CSTRING$(MKD$(PERIB1#), 4, 11, 2, 2))
        REMECOMP.PERIBRU.Refresh
        IINT1# = CVD(Mid$(REBLA$, 45 + 192 + 8 * 3, 8))
        REMECOMP.IMPOGRA(8).Caption = TRIM$(CSTRING$(MKD$(IINT1#), 4, 11, 2, 2))
        REMECOMP.IMPOGRA(8).Refresh
        '
        Close #NX%
        Call FRESHALL
        '
End Sub


Sub VALICUEIVACOMP()
     '
     Screen.MousePointer = 11
     YACUE% = 0
     '
     LU$ = LUDAT$
     TDAT$ = "AS,VC,FC,ND,RB,DC,NC,AS"
     For i% = 0 To 7
        TIS$(i%) = Mid$(TDAT$, 3 * i% + 1, 2)
     Next i%
     '
     CUEIVA$ = AJUSTI$(REMECOMP.Text2(0).TEXT, 12)
     CUEADU$ = AJUSTI$(REMECOMP.Text2(1).TEXT, 12)
     CUEACRE$ = AJUSTI$(REMECOMP.Text2(2).TEXT, 12)
     '
     CUEIVA1% = 0: CUEADU1% = 0: CUEACRE1% = 0
     For U& = 1 To ULTREG&("ECUECON")
        X$ = REGLEIDO$("ECUECON", U&)
        If Left$(X$, 12) = CUEIVA$ Then CUEIVA1% = CVI(Mid$(X$, 41, 2)): DECIVA$ = Mid$(X$, 13, 28)
        If Left$(X$, 12) = CUEADU$ Then CUEADU1% = CVI(Mid$(X$, 41, 2)): DECADU$ = Mid$(X$, 13, 28)
        If Left$(X$, 12) = CUEACRE$ Then CUEACRE1% = CVI(Mid$(X$, 41, 2)): DECACRE$ = Mid$(X$, 13, 28)
     Next U&
     '
     If CUEIVA1% < 1 Or CUEADU1% < 1 Or CUEACRE1% < 1 Then
         Call MENSERR(24, "Faltan Cuentas Contables Correspondientes\a I.V.A. Crédito Fiscal y Aduana.")
         Exit Sub
     End If
     '
     YA1% = 0: YA2% = 0
     For U& = 1 To ULTREG&("CUECON")
        X$ = REGLEIDO$("CUECON", U&)
        If CVI(Left$(X$, 2)) = CUEIVA1% Then YA1% = 1
        If CVI(Left$(X$, 2)) = CUEADU1% Then YA2% = 1
        If CVI(Left$(X$, 2)) = CUEACRE1% Then YA3% = 1
     Next U&
     '
     If YA1% < 1 Then
        X$ = REGBLAN$("CUECON")
        Call REPLA(X$, MKI$(CUEIVA1%), 1, 2)
        Call REPLA(X$, DECIVA$, 3, 29)
        Call REPLA(X$, CUEIVA$, 33, 12)
        Call REPLA(X$, CUEIVA$, 45, 12)
        Call REGAPP("CUECON", X$)
     End If
     '
     If YA2% < 1 Then
        X$ = REGBLAN$("CUECON")
        Call REPLA(X$, MKI$(CUEADU1%), 1, 2)
        Call REPLA(X$, DECADU$, 3, 29)
        Call REPLA(X$, CUEADU$, 33, 12)
        Call REPLA(X$, CUEADU$, 45, 12)
        Call REGAPP("CUECON", X$)
     End If
     '
     If YA3% < 1 Then
        If CUEACRE1% > 0 Then
           X$ = REGBLAN$("CUECON")
           Call REPLA(X$, MKI$(CUEACRE1%), 1, 2)
           Call REPLA(X$, DECACRE$, 3, 29)
           Call REPLA(X$, CUEACRE$, 33, 12)
           Call REPLA(X$, CUEACRE$, 45, 12)
           Call REGAPP("CUECON", X$)
        End If
     End If
     '
     Call GRACONTROL("", "CUEIVA-0", CUEIVA$)
     Call GRACONTROL("", "CUEIVA-4", CUEADU$)
     Call GRACONTROL("SUIVACO", "CUEACRE", CUEACRE$)
     '
     Call CIERRARCH("*.*")
     YACUE% = 1
     '
End Sub
'

Function CUECONTA$(CUE%)
   '
   Static CUECON%(2048), CUECONS$(2048), CANCUEN%
   '
   If CUE% < 0 Then
     Call CIERRARCH("CUECON")
     CANCUEN% = 0
   End If
   '
   If CANCUEN% < 1 Then
     CANCUEN% = 0
     For U& = 1 To ULTREG&("CUECON")
        X$ = REGLEIDO$("CUECON", U&)
        If CVI(Left$(X$, 2)) > 0 Then
           If CANCUEN% < 2048 Then
              CANCUEN% = CANCUEN% + 1
              CUECON%(CANCUEN%) = CVI(Left$(X$, 2))
              CUECONS$(CANCUEN%) = Mid$(X$, 45, 12)
           End If
        End If
     Next U&
   End If
   '
   For IU% = 1 To CANCUEN%
      If CUECON%(IU%) = CUE% Then
         CUECONTA$ = CUECONS$(IU%)
         Exit Function
      End If
   Next IU%
   CUECONTA$ = Space$(12)
   '
End Function
'

Sub ARMAMESCOMP()
     '
     CORRESP$ = TRIM$(REMECOMP.MESAN.TEXT)
     CORRESQ$ = MESANO$(CORRESP$, DES%, HAS%)
     FEX = DES%: DESFE$ = FECHATEX$(FEX)
     FEX = HAS%: HASFE$ = FECHATEX$(FEX)
     '
     CORRESQ$ = TRIM$(CORRESP$)
     ABREMES$ = Left$(CORRESQ$, 3) + Mid$(DESFE$, 7, 2)
     '
End Sub
'

Sub CONASIDIACOM()
    '
    MESA$ = TRIM$(REMECOMP.MESAN.TEXT)
    If MESA$ = "" Then
        MESAN.SetFocus
        Exit Sub
    End If
    '
    MESB$ = MESANO$(MESA$, DES%, HAS%)
    If MESB$ <> MESA$ Then
        REMEVENT.MESAN.SetFocus
        Exit Sub
    End If
    '
    Screen.MousePointer = 11
    REMECOMP.Frame8.Visible = True
    REMECOMP.Frame8.Refresh
    J& = 0
    Call BLOQARCH("!ASIDIAFI")
    Call BLANARCH("!ASIDIAFI")
    TTXXY$ = Chr$(1) + Chr$(2) + MKI$(HAS%)
    FIN& = ULTREG&("ASIDIAR")
    '
    IL0& = 0
    For IL& = 1 To FIN&
        Call TRACE(20, IL& - IL0&, FIN& - IL0&)
        X$ = REGLEIDO$("ASIDIAR", IL&)
        If Left$(X$, 4) = TTXXY$ Then
            If IL0& = 0 Then IL0& = IL& - 1
            CUENU& = CVI(Mid$(X$, 5, 2))
            IDEB# = CVD(Mid$(X$, 39, 8))
            If Abs(IDEB#) >= 0.005 Then
                IHAB# = 0
                If IDEB# < 0 Then
                    IHAB# = Abs(IDEB#)
                    IDEB# = 0
                End If
                CUES$ = Space$(12)
                If CUENU& > 0 Then
                    If CUENU& <= ULTREG&("CUENTAS") Then
                        CUES$ = Left$(REGLEIDO$("CUENTAS", CUENU&), 12)
                    End If
                End If
                DENOCUEX$ = ECOARCH$("ECUECON", CUES$)
                '
                Z$ = REGBLAN$("!ASIDIAFI")
                Call REPLA(Z$, CUES$, 1, 12)
                Call REPLA(Z$, DENOCUEX$, 13, 30)
                Call REPLA(Z$, MKD$(IDEB#), 43, 8)
                Call REPLA(Z$, MKD$(IHAB#), 51, 8)
                '
                J& = J& + 1
                Call GRAREG("!ASIDIAFI", Z$, J&)
            End If
        End If
    Next IL&
    Call SETULTREG("!ASIDIAFI", J&)
    '
    For K1& = 1 To J&
        Call TRACE(20, K1&, J&)
        For K2& = K1& + 1 To J&
            X1$ = REGLEIDO$("!ASIDIAFI", K1&)
            X2$ = REGLEIDO$("!ASIDIAFI", K2&)
            If Left$(X1$, 12) > Left$(X2$, 12) Then
                X0$ = X1$
                X1$ = X2$
                X2$ = X0$
                Call GRAREG("!ASIDIAFI", X1$, K1&)
                Call GRAREG("!ASIDIAFI", X2$, K2&)
            End If
        Next K2&
    Next K1&
    Call CIERRARCH("!ASIDIAFI")
    '
    REMECOMP.Frame8.Visible = False
    Screen.MousePointer = 1
    'Call FINAL("*ASIDIAFI")
    VTB% = VENTABU%("ASIDIAFI/TITUPAN=Asiento Resumen Mensual de Compras - " + CORRESP$)
    '
End Sub
'

Sub BACKUPCOM(modo%, HAYERRO%)
    '
    ' MODO% = 1 ... existencia y acceso a directorio
    '
    Screen.MousePointer = 11
    Dim X As String * 2
    Dim ARESPA$(24)
    CARESPA% = 16
    '
    HAYERRO% = 0
    Call ARMAMESCOMP
    If Len(ABREMES$) <> 5 Then
        Call COMUNI("Período '" + ABREMES$ + "' no Procesable.\  \Consulte a su Soporte de Sistemas.")
        HAYERRO% = 1
        Exit Sub
    End If
    '
    ULOG$ = Left$(CurDir, 2)
    LUBAK$ = LUDAT$ + ABREMES$
    If Mid$(LUBAK$, 2, 1) <> ":" Then
       LUBAK$ = ULOG$ + LUBAK$
    End If
    EXTE$ = Mid$(DESFE$, 8, 1) + Mid$(DESFE$, 4, 2)
    '
    ARESPA$(0) = "ANAICOM"
    ARESPA$(1) = "IFACOMP"
    ARESPA$(2) = "RESUCCP"
    ARESPA$(3) = "DETFACC"
    ARESPA$(4) = "DETDEBC"
    ARESPA$(5) = "DETCREC"
    ARESPA$(6) = "DETPAGC"
    ARESPA$(7) = "DETOTRC"
    ARESPA$(8) = "DETPAGG"
    ARESPA$(9) = "DETFOFI"
    ARESPA$(10) = "SUDICOM"
    ARESPA$(11) = "SUIVACO"
    ARESPA$(12) = "LIVACOM"
    ARESPA$(13) = "REIVACOM"
    'ARESPA$(14) = "REVECLI"
    'ARESPA$(15) = "REVEPIE"
    'ARESPA$(16) = "RECOBCLI"
    '
40  VUELTA% = 1
41  ARXX$ = LUBAK$
    BRXX$ = CurDir
    On Error GoTo 43
    ChDir ARXX$
    ChDir BRXX$
    On Error GoTo 0
    GoTo 50
    '
43  ERNU$ = TRIM$(Str$(Err.Number))
    Resume 44
44  On Error GoTo 0
    If VUELTA% = 1 Then
        On Error Resume Next
        MkDir ARXX$
        On Error GoTo 0
        VUELTA% = 2
        GoTo 41
    End If
    '
    CRXX$ = REPLACAR(ARXX$, "\", "/")
    MENSA$ = "Se ha Producido un Error (" + ERNU$ + ") en la Apertura\"
    MENSA$ = MENSA$ + "para RESPALDO de la Siguiente Carpeta:\  \"
    MENSA$ = MENSA$ + CRXX$ + "\  \"
    '
    MENSA$ = MENSA$ + "Investigue las Siguientes Causas Posibles:\  \"
    MENSA$ = MENSA$ + "   1.- Problemas de Red - No hay Acceso a U.L. '" + Left$(ARXX$, 2) + "'.\"
    MENSA$ = MENSA$ + "   2.- Usuario de Red sin Derechos\"
    MENSA$ = MENSA$ + "       de Creación de Directorios.\"
    MENSA$ = MENSA$ + "   3.- Espacio en disco Insuficiente.\  \"
    MENSA$ = MENSA$ + "   \Consulte a su Soporte de Sistemas.\"
    '
    c$ = FORMESS$(MENSA$)
    RESPU% = MsgBox(c$, 53, "Error de Apertura de Carpeta")
    If RESPU% = 4 Then
        GoTo 40
    End If
    HAYERRO% = 1
    Exit Sub
    '
50  For i% = 0 To CARESPA%
        If TRIM$(ARESPA$(i%)) <> "" Then
            NX% = FILEOPEN(LUBAK$ + "\" + ARESPA$(i%) + "." + EXTE$, 0, 2)
            If NX% < 1 Then
                HAYERRO% = 1
                Exit Sub
            End If
            Get #NX%, 1, X$
            Put #NX%, 1, X$
            Close #NX%
        End If
    Next i%
    '
    '
    If modo% = 2 Then
        LUDA$ = LUDAT$
        If Mid$(LUDA$, 2, 1) <> ":" Then
            LUDA$ = Left$(CurDir, 2) + LUDA$
        End If
        '
        Call CIERRARCH("*.*")
        For i% = 0 To CARESPA%
           If TRIM$(ARESPA$(i%)) <> "" Then
              '
              ORIGI$ = LUDA$ + ARESPA$(i%) + ".DAT"
              If EXISTE%(ORIGI$) = 1 Then
                 DESTI$ = LUBAK$ + "\" + ARESPA$(i%) + "." + EXTE$
                 '
                 If COPYFILE%(ORIGI$, DESTI$) <> 1 Then
                    HAYERRO% = 1
                    Exit Sub
                 End If
                 '
              End If
           End If
        Next i%
    End If
    '
    '
    If modo% = 3 Then
        LUDA$ = LUDAT$
        If Mid$(LUDA$, 2, 1) <> ":" Then
            LUDA$ = Left$(CurDir, 2) + LUDA$
        End If
        C1$ = TRIM$(REMECOMP.MESAN.TEXT)
        C2$ = TRIM$(REMECOMP.Text4(0).TEXT)
        C3$ = TRIM$(REMECOMP.Text4(1).TEXT)
        '
        Call CIERRARCH("*.*")
        If TRIM$(CORRAR$("SUIVACO")) = TRIM$(C1$) Then
            GoTo 80
        End If
        '
        For i% = 0 To CARESPA%
            If TRIM$(ARESPA$(i%)) <> "" Then
                '
                ORIGI$ = LUBAK$ + "\" + ARESPA$(i%) + "." + EXTE$
                DESTI$ = LUDA$ + ARESPA$(i%) + ".DAT"
                '
                If COPYFILE%(ORIGI$, DESTI$) <> 1 Then
                    HAYERRO% = 1
                    Exit Sub
                End If
                '
            End If
        Next i%
        '
'        Call HEADERS("DETACCX", "")
'        Call HEADERS("DETACCX", "Corresponde a: " + TRIM$(C1$) + " - Generado " + C2$ + " - " + C3$ + " hs.")
'        Call HEADERS("RESUCCX", "")
'        Call HEADERS("RESUCCX", "Corresponde a: " + TRIM$(C1$) + " - Generado " + C2$ + " - " + C3$ + " hs.")
        Call GRARGEN("SUIVACO", TRIM$(C1$))
        '
80      For i% = 0 To CARESPA%
            If TRIM$(ARESPA$(i%)) <> "" Then
                '
                Call HEADERS(ARESPA$(i%), "")
                TEHEA$ = "Corresponde a: " + TRIM$(C1$)
                If Left$(UCase$(TRIM$(ARESPA$(i%))), 2) <> "SU" Then
                   TEHEA$ = TEHEA$ + " - Generado " + C2$ + " - " + C3$ + " hs."
                End If
                Call HEADERS(ARESPA$(i%), TEHEA$)
                '
            End If
        Next i%
        '
        ' SUPRIME FECHA GENERACION EN SUBDIARIOS
        CORRESPX$ = "Corresponde a: " + TRIM$(C1$)
        Call HEADERS("SUIVACO", "")
        Call HEADERS("SUIVACO", "Corresponde a: " + CORRESPX$)
        Call HEADERS("LIVACOM", "")
        Call HEADERS("LIVACOM", "Corresponde a: " + CORRESPX$)
        Call HEADERS("SUDICOM", "")
        Call HEADERS("SUDICOM", "Corresponde a: " + CORRESPX$)
        Call HEADERS("ANAICOM", "")
        Call HEADERS("ANAICOM", "Corresponde a: " + CORRESPX$)
        Call HEADERS("REIVACOM", "")
        Call HEADERS("REIVACOM", "Corresponde a: " + CORRESPX$)
        '
    End If
    '
99  Screen.MousePointer = 1
    '
End Sub
'

Sub CALRECOPRO(DES%, HAS%)
        '
Stop
        Call BLANARCH("REVECLI")
        Call BLANARCH("REVEPIE")
        IREVE$ = ""
        '
        Screen.MousePointer = 11
        '
        SUMATO# = 0
        FIN& = ULTREG&("MOVISTO")
        '
        For IL& = 1 To FIN&
              Call TRACE(20, IL&, FIN&)
              MS$ = REGLEIDO$("MOVISTO", IL&)
              FF% = CVI(Left$(MS$, 2))
              If FF% >= DES% And FF% <= HAS% Then
                 COMO$ = UCase$(Mid$(MS$, 33, 2))
                 If TRIM$(COMO$) = "" Then
                    COMO$ = UCase$(Mid$(MS$, 23, 2))
                 End If
                 If COMO$ = "RT" Or COMO$ = "FC" Or COMO$ = "TF" Or COMO$ = "RH" Or COMO$ = "NC" Or COMO$ = "ND" Then
                    NC% = CVI(Mid$(MS$, 84, 2))
                    NRO1& = Val(Mid$(MS$, 36, 7))
                    If NRO1& < 1 Then
                        NRO1& = Val(Mid$(MS$, 26, 7))
                    End If
                    CI% = CVI(Mid$(MS$, 3, 2))
                    CAN# = CVD(Mid$(MS$, 96, 8)) - CVD(Mid$(MS$, 88, 8))
                    PRE# = CVD(Mid$(MS$, 73, 8))
                    GoSub 13700
                 End If
              End If
        Next IL&
        '
        For IL& = 1 To ULTREG&("MATVEN")
              MS$ = REGLEIDO$("MATVEN", IL&)
              FF% = CVI(Mid$(MS$, 7, 2))
              If FF% >= DES% And FF% <= HAS% Then
                 COMO$ = "FC"
                 NC% = CVI(Left$(MS$, 2))
                 NRO1& = CVS(Mid$(MS$, 3, 4))
                 CI% = CVI(Mid$(MS$, 9, 2))
                 CAN# = CVS(Mid$(MS$, 19, 4))
                 PRE# = CVD(Mid$(MS$, 11, 8))
                 GoSub 13700
              End If
        Next IL&
        '
        Call CIERRARCH("REVECLI")
        Call CIERRARCH("REVEPIE")
        '
        FEX = DES%: PERIO$ = FECHATEX$(FEX)
        FEX = HAS%: PERIO$ = PERIO$ + " - " + FECHATEX$(FEX)
        Call HEADERS("REVECLI", "")
        Call HEADERS("REVECLI", "Periodo: " + PERIO$)
        Call HEADERS("REVEPIE", "")
        Call HEADERS("REVEPIE", "Periodo: " + PERIO$)
        '
        Screen.MousePointer = 1
        MOX% = ALTERNA%("Listado General por Cliente\Listado por Codigo de Producto")
        On MOX% GoTo 201, 202
        Exit Sub
        '
201     Call FINAL("*REVECLI"): Exit Sub
202     Call FINAL("*REVEPIE"): Exit Sub
        '
13700   TOT# = Abs(Int(100 * PRE# * CAN# + 0.5)) / 100
        If NC% <> 0 Then
           If NC% < 0 Or NC% > 9999 Then NC% = 9999
           '
           If CI% > 0 Then
               BUSX$ = MKI$(NC%) + MKI$(CI%)
               For K& = 1 To (Len(IREVE$)) / 4  '  To ULTREG&("REVECLI")
                   If Mid$(IREVE$, 4 * K& - 3, 4) = BUSX$ Then
                       RVCL$ = REGLEIDO$("REVECLI", K&)
                       CARVCL = CAN# + CVS(Mid$(RVCL$, 5, 4))
                       VARVCL# = TOT# + CVD(Mid$(RVCL$, 9, 8))
                       Call REPLA(RVCL$, MKS$(CARVCL) + MKD$(VARVCL#), 5, 12)
                       Call GRAREG("REVECLI", RVCL$, K&)
                       '
                       PVCL$ = REGLEIDO$("REVEPIE", K&)
                       Call REPLA(PVCL$, MKS$(CARVCL), 5, 4)
                       Call GRAREG("REVEPIE", PVCL$, K&)
                       '
                       GoTo 13799
                   End If
               Next K&
               '
               CARVCL = CAN#
               RVCL$ = MKI$(NC%) + MKI$(CI%) + MKS$(CARVCL) + MKD$(TOT#) + MKD$(0)
               Call REGAPP("REVECLI", RVCL$)
               '
               PVCL$ = MKI$(CI%) + MKI$(NC%) + MKS$(CARVCL) + MKD$(0) + String$(4, 0)
               Call REGAPP("REVEPIE", PVCL$)
               '
               IREVE$ = IREVE$ + MKI$(NC%) + MKI$(CI%)
               '
           End If
        End If
        '
13799   Return
        '
End Sub
'

Sub CALRESUCO(DES%, HAS%)
    '
Stop
    Screen.MousePointer = 11
    CARESUVE& = 0
    TDAT$ = "AS,VC,FC,ND,RB,DC,NC,AS"
    For i% = 0 To 7
        TIS$(i%) = Mid$(TDAT$, 3 * i% + 1, 2)
    Next i%
    '
    FIN& = ULTREG&("FACTUR")
    '
    For IL& = 1 To FIN&
        Call TRACE(20, IL&, FIN&)
        X$ = REGLEIDO$("FACTUR", IL&)
        NC% = CVI(Left$(X$, 2))
        FF% = CVI(Mid$(X$, 7, 2))
        '
        If FF% >= DES% And FF% <= HAS% Then
           TI% = Asc(Mid$(X$, 13, 1)): If TI% < 0 Or TI% > 7 Then TI% = 2 ' default factura
           VA% = Asc(Mid$(X$, 14, 1)): If VA% < 1 Or VA% > 3 Then VA% = 1 ' default factura
           TIXX$ = TIS$(TI%)
           NRO& = CVS(Mid$(X$, 3, 4))
           '
           NETO# = (Int(100 * (CVD(Mid$(X$, 15, 8)) + CVD(Mid$(X$, 57, 8))) + 0.5)) / 100
           '
           SIG% = 1
           If TI% > 3 Then
               NETO# = Abs(NETO#)
               SIG% = (-1)
           End If
           '
           If TI% > 0 Then
               If TI% <= 6 Then
                   If NC% >= 0 Then
                         CLI$ = RASOCLI$(NC%)
                       ElseIf NC% < 0 Then
                         RECUI& = Abs(NC%)
                         If RECUI& <= ULTREG&("CUITS") Then
                             RCUX$ = REGLEIDO$("CUITS", RECUI&)
                             CLI$ = Left$(RCUX$, 30)
                         End If
                   End If
                   '
                   YR$ = REGBLAN$("RESUVEN")
                   Call REPLA$(YR$, MKI$(NC%), 1, 2)
                   Call REPLA$(YR$, CLI$, 3, 30)
                   Call REPLA(YR$, TIXX$ + AJUSTD$(Str$(NRO&), 6), 33, 8)
                   Call REPLA(YR$, MKI$(FF%), 41, 2)
                   Call REPLA(YR$, MKD$(NETO# * SIG%), 43, 8)
                   CARESUVE& = CARESUVE& + 1
                   Call GRAREG("RESUVEN", YR$, CARESUVE&)
               End If
           End If
        End If
        '
    Next IL&
    '
    Call SETULTREG&("RESUVEN", CARESUVE&)
    Call CIERRARCH("RESUVEN")
    '
    FEX = DES%: PERIO$ = FECHATEX$(FEX)
    FEX = HAS%: PERIO$ = PERIO$ + " - " + FECHATEX$(FEX)
    Call HEADERS("RESUVEN", "")
    Call HEADERS("RESUVEN", "Periodo: " + PERIO$)
    '
    Screen.MousePointer = 1
    Call FINAL("*RESUSOR")
    '
End Sub

Sub LIVACOM(TS$)
   '
   Static REBLA$, REBLB$
   '
   If ULTREG&("LIVACOM") < 1 Then
      UVECRE% = 0
   End If
   '
   If REBLA$ = "" Then REBLA$ = REGBLAN$("LIVACOM")
   If REBLB$ = "" Then REBLB$ = REGBLAN$("REIVACOM")
   RECUE0$ = String$(128, 0)
   '
   FEAL% = HOY(HOS$)
   HOAL% = Int(HORANUM(Time$))
   USUA% = USNBR% Mod 100
   DALTA$ = Chr$(0) + MKI$(FEAL%) + MKI$(HOAL%) + Chr$(USUA%)
   NROIN& = 0
   '
   FC% = CVI(Mid$(TS$, 1, 2))
   Call REPLA(RECUE0$, MKI$(FC%), 3, 2)
   FR% = CVI(Mid$(TS$, 279, 2))
   Call REPLA(RECUE0$, MKI$(FR%), 5, 2)
   '
   FEX = FC%: FFTT$ = FECHATEX$(FEX)
   CMA$ = Chr$(Val(Mid$(FFTT$, 4, 2))) + Chr$(Val(Mid$(FFTT$, 7, 2)))
   Call REPLA(RECUE0$, CMA$, 1, 2)
   '
   NROIN& = NROIN& + 1
   Call REPLA(RECUE0$, MKS$(NROIN&), 7, 4)
   '
   NUCO$ = Mid$(TS$, 7, 18)
   TIMOVI$ = Left$(NUCO$, 2)
   Call REPLA(RECUE0$, NUCO$, 11, 18)
   Call REPLA(RECUE0$, MKI$(BINVAL%(NUCO$)), 29, 2)
   '
   NC% = CVI(Mid$(TS$, 5, 2))
   RACLIX$ = RASOCLI$((-1) * NC%)
   CUCLIX$ = CUITCLI$((-1) * NC%)
   If Mid$(TS$, 295, 6) <> String$(6, 0) Then
      If Left$(RACLIX$, 6) = "VARIOS" Then
         RACLIX$ = Mid$(TS$, 25, 24)
         CUCLIY$ = "": COCUX$ = Mid$(TS$, 295, 6)
         For KKII% = 1 To 6
           KKJJ$ = TRIM$(Str$(Asc(Mid$(COCUX$, KKII%, 1))))
           While Len(KKJJ$) < 2: KKJJ$ = "0" + KKJJ$: Wend
           CUCLIY$ = CUCLIY$ + Left$(KKJJ$, 2)
           If Len(CUCLIY$) = 2 Then CUCLIY$ = CUCLIY$ + "-"
           If Len(CUCLIY$) = 11 Then CUCLIY$ = CUCLIY$ + "-"
         Next KKII%
         CUCLIX$ = Left$(CUCLIY$, 13)
      End If
   End If
   '
   Call REPLA(RECUE0$, MKI$(NC%), 31, 2)
   Call REPLA(RECUE0$, RACLIX$, 33, 30)
   Call REPLA(RECUE0$, CUCLIX$, 63, 16)
   Call REPLA(RECUE0$, POSIVAC$((-1) * NC%), 79, 12)
   '
   SUIMPUTA# = CVD(Mid$(TS$, 123, 8)) - CVD(Mid$(TS$, 115, 8))
   Call REPLA(RECUE0$, MKD$(SUIMPUTA#), 91, 8)
   SIG% = 1: If SUIMPUTA# < 0 Then SIG% = (-1)
   '
   Call REPLA(RECUE0$, DALTA$, 99, 6)
   Call REPLA(RECUE0$, String$(6, 0) + Space$(16), 105, 22)
   Call REPLA(RECUE0$, String$(2, 0), 127, 2)
   '
   PERIVA# = SIG% * Abs(CVD(Mid$(TS$, 91, 8)))
   PERGAN# = SIG% * Abs(CVD(Mid$(TS$, 99, 8)))
   'AQUI ESTA PERCEPCION I.BRUTOS FACTURAS COMPRAS
   PERBRU# = SIG% * Abs(CVD(Mid$(TS$, 107, 8)))
   COPROV$ = Mid$(TS$, 493, 1)
   DEPROV$ = Mid$(TS$, 494, 19)
   '
   For KUY% = 0 To 32
     PERBRUY#(KUY%) = 0
     COPROVY$(KUY%) = ""
     DEPROVY$(KUY%) = ""
   Next KUY%
   '
   KUY% = 0: UINPER% = 0
   If Abs(PERBRU#) > 0 Then
     If COPROV$ <> "*" Then
          KUY% = 0
          PERBRUY#(0) = PERBRU#
          COPROVY$(0) = COPROV$
          DEPROVY$(0) = DEPROV$
        Else
          RFG$ = RECFILT$("PERCIBFA", 2, MKI$(NC%))
          For UKY& = 1 To Len(RFG$) Step 4
            RGY& = CVS(Mid$(RFG$, UKY&, 4))
            RGZ$ = REGLEIDO$("PERCIBFA", RGY&)
            If TRIM$(Mid$(RGZ$, 5, 18)) = TRIM$(NUCO$) Then
              If CVI(Left$(RGZ$, 2)) = FC% Then
                COJU$ = TRIM$(Mid$(RGZ$, 97, 1))
                IMPE# = Abs(CVD(Mid$(RGZ$, 89, 8)))
                If COJU$ <> "" Then
                  If IMPE# >= 0.005 Then
                    PERBRUY#(KUY%) = SIG% * IMPE#
                    COPROVY$(KUY%) = COJU$
                    DEPROVY$(KUY%) = Mid$(RGZ$, 98, 19)
                    UINPER% = KUY%
                    KUY% = KUY% + 1
                  End If
                End If
              End If
            End If
          Next UKY&
     End If
   End If
   '
   JI% = 0: YAPER% = 0
   For KKI% = 1 To 8
     IGRA# = SIG% * Abs(CVD(Mid$(TS$, 300 + 24 * KKI% - 23, 8)))
     VIVA# = SIG% * Abs(CVD(Mid$(TS$, 300 + 24 * KKI% - 15, 8)))
     TASANO$ = Mid$(TS$, 300 + 24 * KKI% - 7, 8)
     COTASA% = KKI%
     '
     If Abs(IGRA#) >= 0.005 Or Abs(VIVA#) >= 0.005 Then
       RECUE$ = REGBLAN$("LIVACOM")
       Call REPLA(RECUE$, RECUE0$, 1, 128)
       'If JI% < 1 Then
         If YAPER% < 1 Then
           Call REPLA(RECUE$, MKD$(PERIVA#), 161, 8)
           Call REPLA(RECUE$, MKD$(PERGAN#), 169, 8)
         End If
         'Call REPLA(RECUE$, MKD$(PERBRU#), 193, 8)
         'Call REPLA(RECUE$, MKI$(CUEIPROVIN%(COPROV$)), 201, 2)
         'Call REPLA(RECUE$, COPROV$, 203, 1)
         'Call REPLA(RECUE$, DEPROV$, 205, 24)
         'Call REPLA(RECUE$, CUEIBRUPRO$(COPROV$), 229, 12)
         Call REPLA(RECUE$, MKD$(PERBRUY#(JI%)), 193, 8)
         Call REPLA(RECUE$, MKI$(CUEIPROVIN%(COPROVY$(JI%))), 201, 2)
         Call REPLA(RECUE$, COPROVY$(JI%), 203, 1)
         Call REPLA(RECUE$, DEPROVY$(JI%), 205, 24)
         Call REPLA(RECUE$, CUEIBRUPRO$(COPROVY$(JI%)), 229, 12)
         YAPER% = 1
       'End If
       Call REPLA(RECUE$, MKD$(IGRA#), 137, 8)
       Call REPLA(RECUE$, MKD$(VIVA#), 145, 8)
       TAREA# = 0: If Abs(IGRA#) > 0.005 Then TAREA# = 100 * Abs(VIVA#) / Abs(IGRA#)
       If TAREA# > 99.9 Then TAREA# = 0
       '
       If TAREA# > 18.5 And TAREA# < 19.5 Then TASANO$ = "  19,0 %"
       If TAREA# > 20.5 And TAREA# < 21.5 Then TASANO$ = "  21,0 %"
       If TAREA# > 9 And TAREA# <= 10 Then TASANO$ = "   9,5 %"
       If TAREA# > 10 And TAREA# <= 11 Then TASANO$ = "  10,5 %"
       '
       Call REPLA(RECUE$, TASANO$, 129, 8)
       Call REPLA(RECUE$, MKD$(TAREA#), 153, 8)
       If TIMOVI$ <> "AS" Then
         Call REGAPP("LIVACOM", RECUE$)
         GoSub 80         ' CARGAR RESUMEN MENSUAL
       End If
       Call REPLA(RECUE0$, MKD$(0), 91, 8)
       JI% = JI% + 1
     End If
   Next KKI%
   '
   If YAPER% = 0 Then
     If PERIVA# >= 0.005 Or PERGAN# >= 0.005 Or PERBRU# > 0.005 Then
       RECUE$ = REGBLAN$("LIVACOM")
       Call REPLA(RECUE$, RECUE0$, 1, 128)
       Call REPLA(RECUE$, MKD$(PERIVA#), 161, 8)
       Call REPLA(RECUE$, MKD$(PERGAN#), 169, 8)
       'Call REPLA(RECUE$, MKD$(PERBRU#), 193, 8)
       'Call REPLA(RECUE$, MKI$(CUEIPROVIN%(COPROV$)), 201, 2)
       'Call REPLA(RECUE$, COPROV$, 203, 1)
       'Call REPLA(RECUE$, DEPROV$, 205, 24)
       'Call REPLA(RECUE$, CUEIBRUPRO$(COPROV$), 229, 12)
       Call REPLA(RECUE$, MKD$(PERBRUY#(JI%)), 193, 8)
       Call REPLA(RECUE$, MKI$(CUEIPROVIN%(COPROVY$(JI%))), 201, 2)
       Call REPLA(RECUE$, COPROVY$(JI%), 203, 1)
       Call REPLA(RECUE$, DEPROVY$(JI%), 205, 24)
       Call REPLA(RECUE$, CUEIBRUPRO$(COPROVY$(JI%)), 229, 12)
       YAPER% = 1
       If TIMOVI$ <> "AS" Then
         Call REGAPP("LIVACOM", RECUE$)
         GoSub 80         ' CARGAR RESUMEN MENSUAL
       End If
       Call REPLA(RECUE0$, MKD$(0), 91, 8)
       JI% = JI% + 1
     End If
   End If
   '
   While JI% <= UINPER%
     RECUE$ = REGBLAN$("LIVACOM")
     Call REPLA(RECUE$, RECUE0$, 1, 128)
     Call REPLA(RECUE$, MKD$(0), 161, 8)
     Call REPLA(RECUE$, MKD$(0), 169, 8)
       'Call REPLA(RECUE$, MKD$(PERBRU#), 193, 8)
       'Call REPLA(RECUE$, MKI$(CUEIPROVIN%(COPROV$)), 201, 2)
       'Call REPLA(RECUE$, COPROV$, 203, 1)
       'Call REPLA(RECUE$, DEPROV$, 205, 24)
       'Call REPLA(RECUE$, CUEIBRUPRO$(COPROV$), 229, 12)
     Call REPLA(RECUE$, MKD$(PERBRUY#(JI%)), 193, 8)
     Call REPLA(RECUE$, MKI$(CUEIPROVIN%(COPROVY$(JI%))), 201, 2)
     Call REPLA(RECUE$, COPROVY$(JI%), 203, 1)
     Call REPLA(RECUE$, DEPROVY$(JI%), 205, 24)
     Call REPLA(RECUE$, CUEIBRUPRO$(COPROVY$(JI%)), 229, 12)
     YAPER% = 1
     If TIMOVI$ <> "AS" Then
       Call REGAPP("LIVACOM", RECUE$)
       GoSub 80         ' CARGAR RESUMEN MENSUAL
     End If
     Call REPLA(RECUE0$, MKD$(0), 91, 8)
     JI% = JI% + 1
   Wend
   '
   If TIMOVI$ <> "AS" Then
     Call REGAPP("LIVACOM", REBLA$)
   End If
   '
   SS$ = Left$(TS$, 300)
   SS1$ = Left$(SS$, 2)
   SS4$ = Mid$(SS$, 5, 2)
   SS6$ = Mid$(SS$, 7, 18)
   SS7$ = Mid$(SS$, 51, 64)
   SD1$ = Mid$(SS$, 115, 8)
   SH1$ = Mid$(SS$, 123, 8)
   Y6$ = Mid$(SS$, 199, 80)
   TT7$ = Mid$(TS$, 301, 192)
   '
   TI$ = UCase$(Left$(SS6$, 2))
   '
   FEMI% = CVI(SS1$)
   NC% = CVI(SS4$)
   NF$ = SS6$
   '
   For J% = 0 To 50: GRU%(J%) = 0: IPAR#(J%) = 0: Next J%
   For J% = 1 To 8
      IPAR#(J%) = 0
      GRU%(J%) = CVI(Mid$(Y6$, 10 * J% - 9, 2))
      If GRU%(J%) > 0 Then
         IPAR#(J%) = CVD(Mid$(Y6$, 10 * J% - 7, 8))
12444    For K1% = 1 To NUGRUPO%
           If GRUPO%(K1%) = GRU%(J%) Then
              GoTo 49
           End If
         Next K1%
' esto se puso para atajar el caso de que no figure el grupo contable en la tabla actual
Call MENSERR(24, "Grupo " + TRIM$(Str$(GRU%(J%))) + " Inexistente en Tabla de Grupos Contables.")
GRU%(J%) = GRUPO%(1)
GoTo 12444
      End If
      '
      GRU%(J%) = 0: IPAR#(J%) = 0
49 Next J%
    '
   GRU%(0) = CUECOCLI%((-1) * NC%)
   GRUIII% = GRU%(0)
   CCCS$ = CUECONTA$(GRUIII%)
   If TRIM$(CCCS$) = "" Then
      GRU%(0) = CUEACRE1%
   End If
   IPAR#(0) = (-1) * (Abs(SUIMPUTA#)) * SIG%
   '
   For KKI% = 1 To 8
     VIVA# = Abs(CVD(Mid$(TS$, 300 + 24 * KKI% - 15, 8)))
     If VIVA# >= 0.005 Then
       GRU%(KKI% + 8) = CUEIVAX%(KKI%)
       IPAR#(KKI% + 8) = VIVA#
     End If
   Next KKI%
   '
   If Abs(PERIVA#) >= 0.005 Then
     GRU%(17) = CUEPERIVA%
     IPAR#(17) = Abs(PERIVA#)
   End If
   '
   If Abs(PERGAN#) >= 0.005 Then
     GRU%(18) = CUEPERGAN%
     IPAR#(18) = Abs(PERGAN#)
   End If
   '
   If Abs(PERBRU#) >= 0.005 Then
     For KIY% = 0 To UINPER%
       GRU%(19 + KIY%) = CUEIPROVIN%(COPROVY$(KIY%))
       IPAR#(19 + KIY%) = Abs(PERBRUY#(KIY%))
       If GRU%(19 + KIY%) < 1 Then
         For KKI% = 1 To CAPROVI%
           If TRIM$(UCase$(DEPROVY$(KIY%))) = TRIM$(UCase$(Left$(DEPROVIN$(KKI%), 19))) Then
             GRU%(19 + KIY%) = CUEPROVIN%(KKI%)
             GoTo 52
           End If
         Next KKI%
       End If
52     If GRU%(19 + KIY%) < 1 Then
         Call MENSERR(24, "Percepción Ingresos Brutos no Procesable\en Documento '" + TRIM$(NF$) + "'\de " + TRIM$(RASOCLI$((-1) * NC%)))
       End If
     Next KIY%
   End If
   '
   For J% = 1 To 50
      IPAR#(J%) = (Abs(IPAR#(J%))) * SIG%
   Next J%
   '
   IREN% = 0
   For J% = 0 To 50
     If GRU%(J%) > 0 Then
       If Abs(IPAR#(J%)) >= 0.005 Then
         '
         If J% > 0 Then
           A0$ = REGBLAN$("ANAICOM")
           Call REPLA(A0$, MKI$(GRU%(J%)), 1, 2)
           Call REPLA(A0$, MKI$(NC%), 3, 2)
           Call REPLA(A0$, NF$, 5, 18)
           Call REPLA(A0$, MKI$(FEMI%), 23, 2)
           Call REPLA(A0$, MKD$(IPAR#(J%)), 25, 8)
           Call REPLA(A0$, MKD$(0), 33, 8)
           Call REGAPP("ANAICOM", A0$)
           '
           A0$ = REGBLAN$("SUDICOM")
           Call REPLA(A0$, MKI$(FEMI%), 1, 2)
           Call REPLA(A0$, NF$, 3, 18)
           Call REPLA(A0$, RASOCLI$((-1) * NC%), 21, 30)
           Call REPLA(A0$, CUITCLI$((-1) * NC%), 51, 16)
           '
           IREN% = IREN% + 1
           If IREN% = 1 Then
             Call REPLA(A0$, MKD$(Abs(SUIMPUTA#) * SIG%), 67, 8)
           End If
           '
           GRUIII% = GRU%(J%)
           CCCS$ = CUECONTA$(GRUIII%)
'If GRU%(J%) = 400 Then
'AAA = BBB
'MsgBox CCCS$ & "   " & IPAR#(J%) & "   " & J%
'End If
           Call REPLA(A0$, CCCS$, 75, 12)
           If IPAR#(J%) >= 0 Then
                Call REPLA(A0$, MKD$(IPAR#(J%)), 95, 8)
              Else
                Call REPLA(A0$, MKD$(IPAR#(J%)), 87, 8)
           End If
           '
           If Left$(UCase$(NF$), 2) <> "AS" Then
              Call REGAPP("SUDICOM", A0$)
           End If
           '
         End If
         '
         For K% = 1 To CASDIAR%
           If GRU%(J%) = CUEASI%(K%) Then
             IMASI#(K%) = IMASI#(K%) + IPAR#(J%)
'If CUEASI%(k%) = 400 Then
'AAA = BBB
'End If
              GoTo 59
           End If
         Next K%
         CASDIAR% = CASDIAR% + 1
         CUEASI%(CASDIAR%) = GRU%(J%)
         IMASI#(CASDIAR%) = IPAR#(J%)
'If CUEASI%(CASDIAR%) = 400 Then
'AAA = BBB
'End If
         '
       End If
     End If
59 Next J%
   Exit Sub
   '
   '
80 ' CARGA RESUMEN MENSUAL DE IVA - COMPRAS
   CORUBRO% = 1: DERUBRO$ = "COMPRA"
   COCONCE% = 8: DECONCE$ = "COMPROBANTES VARIOS"
   '
   If SIG% < 0 Then
       COCONCE% = 9: DECONCE$ = "NOTAS DE CREDITO"
     Else
       If Left$(NUCO$, 2) = "FC" Then COCONCE% = 1: DECONCE$ = "FACTURAS"
       If Left$(NUCO$, 2) = "ND" Then COCONCE% = 2: DECONCE$ = "NOTAS DE DEBITO"
       If Left$(NUCO$, 2) = "TF" Then COCONCE% = 3: DECONCE$ = "TICKET-FACTURA"
       If Left$(NUCO$, 2) = "RH" Then COCONCE% = 4: DECONCE$ = "RECIBO HONORARIOS"
   End If
   '
   COCATE% = COTASA%: DECATE$ = TRIM$(Mid$(RECUE$, 129, 8)) + " - " + Mid$(RECUE0$, 79, 12)
   '
   BASEIMPO# = CVD(Mid$(RECUE$, 137, 8))
   VALOIMPU# = CVD(Mid$(RECUE$, 145, 8))
   GoSub 82
   '
   CORUBRO% = 2: DERUBRO$ = "PERCEPCION"        ' PERCEPCIONES
   For COCO% = 1 To 3
      COCONCE% = 10 + COCO%
      BASEIMPO# = 0
      COCATE% = 0: DECATE$ = ""
      VALOIMPU# = CVD(Mid$(RECUE$, 153 + 8 * COCO%, 8))
      If COCO% = 1 Then
         DECONCE$ = "PERCEPCION I.V.A."
        ElseIf COCO% = 2 Then
         DECONCE$ = "PERCEPCION GANANCIAS"
        ElseIf COCO% = 3 Then
         DECONCE$ = "PERCEPC.ING.BRUTOS"
         DECATE$ = Mid$(RECUE$, 205, 24)
         VALOIMPU# = CVD(Mid$(RECUE$, 193, 8))
      End If
      GoSub 82
   Next COCO%
   '
Return
   '
   '
82 If (Abs(BASEIMPO#) + Abs(VALOIMPU#)) >= 0.005 Then
      ABUS$ = Chr$(CORUBRO%) + AJUSTI$(DERUBRO$, 15)
      ABUS$ = ABUS$ + Chr$(COCONCE%) + AJUSTI$(DECONCE$, 23)
      ABUS$ = ABUS$ + Chr$(COCATE%) + AJUSTI$(DECATE$, 23)
      '
      For UVE% = 1 To UVECRE%
        If Left$(VECREIVA$(UVE%), 64) = ABUS$ Then
           VCC$ = VECREIVA$(UVE%)
           GoTo 84
        End If
      Next UVE%
      UVECRE% = UVECRE% + 1
      UVE% = UVECRE%
      VCC$ = REBLB$
      Call REPLA(VCC$, ABUS$, 1, 64)
      '
84    BASEIMPO# = BASEIMPO# + CVD(Mid$(VCC$, 65, 8))
      VALOIMPU# = VALOIMPU# + CVD(Mid$(VCC$, 81, 8))
      Call REPLA(VCC$, MKD$(BASEIMPO#), 65, 8)
      Call REPLA(VCC$, MKD$(VALOIMPU#), 81, 8)
      Call REPLA(VCC$, MKD$(BASEIMPO# + VALOIMPU#), 97, 8)
      VECREIVA$(UVE%) = VCC$
   End If
   '
Return
   
End Sub

