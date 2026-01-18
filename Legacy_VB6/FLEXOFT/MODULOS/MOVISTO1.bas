Attribute VB_Name = "MOVISTO1"
Public CAITPE%, CIPE%(2048), CADESPA(2048), REPE&(2048)
'
Sub MOVISTO(FECHA%, CI%, LOTE$, CMOV$, DOPRI$, DOSEC$, REFE$, VUNI#, MONE$, NC%, DEPO%, CANTI)
   '
   Static CTX%, REBLAN$, RELIBR$, UREMOVI&, DECONSUM$
   '
   If CANTI < 0 Then
     If DECONSUM$ = "" Then
       DECONSUM$ = "MANUA"
       If CONTROL$("", "DECONSUM") = "AUTOM" Then
         DECONSUM$ = "AUTOM"
       End If
     End If
   End If
   '
   SMP = Screen.MousePointer
   UREMAS& = ULTREG&("MASTER")
   '
   If ULTREG&("MOVISTO") < UREMOVI& Then
      Call CIERRARCH("*.*")
      Call MENSERR(24, "Imposible Registrar este Movimiento.\  \Cierre la Aplicación y Vuelva a Intentarlo.")
      Call FINAL("")
   End If
   '
   User% = USNBR% Mod 100
   If User% < 1 Then User% = 1
   Call CARDEPOS
   '
   If CTX% < 1 Then
      REBLAN$ = REGBLAN$("MOVISTO")
      RELIBR$ = String$(128, 45)
      Call COPYSTRU("MOVISTO", "BKMOVIST")
      CTX% = 1
   End If
   '
   ' ********** ARMAR REGISTRO DE MOVIMIENTO **********
9  DPRX$ = DOPRI$
10 PPXX% = InStr(DPRX$, "-0")
   If PPXX% > 0 Then
      DPRX$ = Left$(DPRX$, PPXX%) + Mid$(DPRX$, PPXX% + 2)
      GoTo 10
   End If
   DOPRIX$ = AJUSTI$(DPRX$, 10)
   '
   DPRX$ = DOSEC$
20 PPXX% = InStr(DPRX$, "-0")
   If PPXX% > 0 Then
      DPRX$ = Left$(DPRX$, PPXX%) + Mid$(DPRX$, PPXX% + 2)
      GoTo 20
   End If
   DOSECX$ = AJUSTI$(DPRX$, 10)
   '
   NCLIE% = NC%: NPROV% = 0
   If NC% < 0 Then
      NCLIE% = 0: NPROV% = Abs(NC%)
   End If
   '
   ENTRADA# = CDbl(CANTI): SALIDA# = 0
   If CANTI < 0 Then
      ENTRADA# = 0: SALIDA# = CDbl(Abs(CANTI))
   End If
   CANACU = SALDETOT + CANTI
   '
   DepoMovi% = DEPO%
   If DECONSUM$ = "AUTOM" Then
     If CANTI < 0 Then
       DEPII% = DEPO%
       DepoMovi% = DEPOSIBLE%(CI%, CANTI, DEPII%)
     End If
   End If
   '
   RECORD$ = REBLAN$
   Call REPLA(RECORD$, MKI$(FECHA%), 1, 2)
   Call REPLA(RECORD$, MKI$(CI%), 3, 2)
   Call REPLA(RECORD$, LOTE$, 5, 16)
   Call REPLA(RECORD$, CMOV$, 21, 2)
   If CMOV$ = "*R" Then
        Call REPLA(RECORD$, "RP", 21, 2)      ' reparación
      ElseIf CMOV$ = "*D" Then
        Call REPLA(RECORD$, "DR", 21, 2)      ' DESPACHO reparación
   End If
   Call REPLA(RECORD$, DOPRIX$, 23, 10)
   Call REPLA(RECORD$, DOSECX$, 33, 10)
   Call REPLA(RECORD$, REFE$, 43, 30)
   Call REPLA(RECORD$, MKD$(VUNI#), 73, 8)
   Call REPLA(RECORD$, MONE$, 81, 2)
   Call REPLA(RECORD$, Chr$(DepoMovi%), 83, 1)
   Call REPLA(RECORD$, MKI$(NCLIE%), 84, 2)
   Call REPLA(RECORD$, MKI$(NPROV%), 86, 2)
   Call REPLA(RECORD$, MKD$(ENTRADA#), 88, 8)
   Call REPLA(RECORD$, MKD$(SALIDA#), 96, 8)
   Call REPLA(RECORD$, String$(8, 0), 104, 8)
   Call REPLA(RECORD$, String$(6, 0), 112, 6)
   Call REPLA(RECORD$, MKS$(CANACU), 118, 4)
   '
   If UCase$(CMOV$) = "PP" Or CMOV$ = "DM" Then
     CAVERI$ = DOSECX$
     CAVERJ$ = MKS$(Val(Left$(CAVERI$, 6))) + MKI$(Val(Mid$(CAVERI$, 8, 3)))
     Call REPLA(RECORD$, CAVERJ$, 112, 6)
   End If
   '
   HO% = HOY(HOS$)
   HP = 100 * Val(Left$(Time$, 2)) + Val(Mid$(Time$, 4, 2))
   '
   Call REPLA(RECORD$, MKI$(HO%), 122, 2)
   Call REPLA(RECORD$, MKI$(Int(HP)), 124, 2)
   Call REPLA(RECORD$, Chr$(User%), 126, 1)
   Call REPLA(RECORD$, String$(2, 0), 127, 2)
   '
   Dim BUGRA As String * 140
   If Left$(CMOV$, 1) <> "*" Then
         Call REGAPP("MOVISTO", RECORD$)
         REMOVIST& = ULTREG&("MOVISTO")
         If REMOVIST& > UREMOVI& Then UREMOVI& = REMOVIST&
         '
         NKX% = FreeFile
         Open LUCOM$ + "\MOVISTO.LOC" For Random Access Read Write As #NKX% Len = 140
         RECONO& = 1 + LOF(NKX%) / 140
         BUGRA$ = "MOVISTO " + MKS$(REMOVIST&) + RECORD$
         Put #NKX%, RECONO&, BUGRA$
         Close #NKX%
         '
         If CI% > 0 Then FLAGMOVI%(CI%) = 1
         '
     Else               ' reparaciones
         Call COPYSTRU("MOVISTO", "MOVIREPA")
         Call REGAPP("MOVIREPA", RECORD$)
         REMOVIRE& = ULTREG&("MOVIREPA")
         Call CIERRARCH("MOVIREPA")
         '
         NKX% = FreeFile
         Open LUCOM$ + "\MOVISTO.LOC" For Random Access Read Write As #NKX% Len = 140
         RECONO& = 1 + LOF(NKX%) / 140
         BUGRA$ = "MOVIREPA" + MKS$(REMOVIRE&) + RECORD$
         Put #NKX%, RECONO&, BUGRA$
         Close #NKX%
         '
   End If
   '
   Screen.MousePointer = SMP
   '
End Sub
'
Sub MOVISTO_VIE(FECHA%, CI%, LOTE$, CMOV$, DOPRI$, DOSEC$, REFE$, VUNI#, MONE$, NC%, DEPO%, CANTI)
   '
   Static CTX%, REBLAN$, RELIBR$, UREMOVI&, DECONSUM$
   '
   If CANTI < 0 Then
     If DECONSUM$ = "" Then
       DECONSUM$ = "MANUA"
       If CONTROL$("", "DECONSUM") = "AUTOM" Then
         DECONSUM$ = "AUTOM"
       End If
     End If
   End If
   '
   SMP = Screen.MousePointer
   UREMAS& = ULTREG&("MASTER")
   '
   If ULTREG&("MOVISTO") < UREMOVI& Then
      Call CIERRARCH("*.*")
      Call MENSERR(24, "Imposible Registrar este Movimiento.\  \Cierre la Aplicación y Vuelva a Intentarlo.")
      Call FINAL("")
   End If
   '
   User% = USNBR% Mod 100
   If User% < 1 Then User% = 1
   Call CARDEPOS
   If CTX% < 1 Then
      REBLAN$ = REGBLAN$("MOVISTO")
      RELIBR$ = String$(128, 45)
      Call COPYSTRU("MOVISTO", "BKMOVIST")
      CTX% = 1
   End If
   '
   '
   If NOCLOSE% > 0 Then GoTo 9
   '
   '
   CTSX% = 0
   If ULTREG&("SALSTOCK") < UREMAS& - 256 Then Screen.MousePointer = 11
   While ULTREG&("SALSTOCK") < UREMAS&
      X$ = String$(39, 0) + Chr$(1)
      Call REGAPP("SALSTOCK", X$)
      CTSX% = 1
   Wend
   '
   If ULTREG&("SALDEPOS") < UREMAS& - 256 Then Screen.MousePointer = 11
   While ULTREG&("SALDEPOS") < UREMAS&
      Y$ = REGLEIDO$("SALSTOCK", 1 + ULTREG&("SALDEPOS"))
      X$ = String$(256, 0)
      Call REPLA(X$, Y$, 1, 40)
      Call REPLA(X$, MKI$(HOII%), 253, 2)
      Call REPLA(X$, MKI$(Int(HORANUM(Time$))), 255, 2)
      Call REGAPP("SALDEPOS", X$)
      CTSX% = 1
   Wend
   '
   If CTSX% > 0 Then
      Call CIERRARCH("SALSTOCK")   ' para liberar bloqueos
      Call CIERRARCH("SALDEPOS")   ' para liberar bloqueos
   End If
   '
   Call VERISAL(CI%)               ' verificar saldos cruzados
   '
   SALDETOT = 0
   REGSTO& = CI% + 1
   If CI% > 0 Then
      If REGSTO& <= UREMAS& Then
         SDP$ = REGLEIDO$("SALDEPOS", REGSTO&)
         For KI% = 1 To NUSU%
            SALDETOT = SALDETOT + CVS(Mid$(SDP$, 4 * KI% - 3, 4))
         Next KI%
      End If
   End If
   '
   ' ********** ARMAR REGISTRO DE MOVIMIENTO **********
   
9  DPRX$ = DOPRI$
10 PPXX% = InStr(DPRX$, "-0")
   If PPXX% > 0 Then
      DPRX$ = Left$(DPRX$, PPXX%) + Mid$(DPRX$, PPXX% + 2)
      GoTo 10
   End If
   DOPRIX$ = AJUSTI$(DPRX$, 10)
   '
   DPRX$ = DOSEC$
20 PPXX% = InStr(DPRX$, "-0")
   If PPXX% > 0 Then
      DPRX$ = Left$(DPRX$, PPXX%) + Mid$(DPRX$, PPXX% + 2)
      GoTo 20
   End If
   DOSECX$ = AJUSTI$(DPRX$, 10)
   '
   NCLIE% = NC%: NPROV% = 0
   If NC% < 0 Then
      NCLIE% = 0: NPROV% = Abs(NC%)
   End If
   '
   ENTRADA# = CDbl(CANTI): SALIDA# = 0
   If CANTI < 0 Then
      ENTRADA# = 0: SALIDA# = CDbl(Abs(CANTI))
   End If
   CANACU = SALDETOT + CANTI
   '
   DepoMovi% = DEPO%
   If DECONSUM$ = "AUTOM" Then
     If CANTI < 0 Then
       DEPII% = DEPO%
       DepoMovi% = DEPOSIBLE%(CI%, CANTI, DEPII%)
     End If
   End If
   '
   RECORD$ = REBLAN$
   Call REPLA(RECORD$, MKI$(FECHA%), 1, 2)
   Call REPLA(RECORD$, MKI$(CI%), 3, 2)
   Call REPLA(RECORD$, LOTE$, 5, 16)
   Call REPLA(RECORD$, CMOV$, 21, 2)
   If CMOV$ = "*R" Then
        Call REPLA(RECORD$, "RP", 21, 2)      ' reparación
      ElseIf CMOV$ = "*D" Then
        Call REPLA(RECORD$, "DR", 21, 2)      ' DESPACHO reparación
   End If
   Call REPLA(RECORD$, DOPRIX$, 23, 10)
   Call REPLA(RECORD$, DOSECX$, 33, 10)
   Call REPLA(RECORD$, REFE$, 43, 30)
   Call REPLA(RECORD$, MKD$(VUNI#), 73, 8)
   Call REPLA(RECORD$, MONE$, 81, 2)
   Call REPLA(RECORD$, Chr$(DepoMovi%), 83, 1)
   Call REPLA(RECORD$, MKI$(NCLIE%), 84, 2)
   Call REPLA(RECORD$, MKI$(NPROV%), 86, 2)
   Call REPLA(RECORD$, MKD$(ENTRADA#), 88, 8)
   Call REPLA(RECORD$, MKD$(SALIDA#), 96, 8)
   Call REPLA(RECORD$, String$(8, 0), 104, 8)
   Call REPLA(RECORD$, String$(6, 0), 112, 6)
   Call REPLA(RECORD$, MKS$(CANACU), 118, 4)
   '
   If UCase$(CMOV$) = "PP" Or CMOV$ = "DM" Then
     CAVERI$ = DOSECX$
     CAVERJ$ = MKS$(Val(Left$(CAVERI$, 6))) + MKI$(Val(Mid$(CAVERI$, 8, 3)))
     Call REPLA(RECORD$, CAVERJ$, 112, 6)
   End If
   '
   HO% = HOY(HOS$)
   HP = 100 * Val(Left$(Time$, 2)) + Val(Mid$(Time$, 4, 2))
   '
   Call REPLA(RECORD$, MKI$(HO%), 122, 2)
   Call REPLA(RECORD$, MKI$(Int(HP)), 124, 2)
   Call REPLA(RECORD$, Chr$(User%), 126, 1)
   Call REPLA(RECORD$, String$(2, 0), 127, 2)
   '
   Dim BUGRA As String * 140
   If Left$(CMOV$, 1) <> "*" Then
         Call REGAPP("MOVISTO", RECORD$)
         REMOVIST& = ULTREG&("MOVISTO")
         If REMOVIST& > UREMOVI& Then UREMOVI& = REMOVIST&
         '
         NKX% = FreeFile
         Open LUCOM$ + "\MOVISTO.LOC" For Random Access Read Write As #NKX% Len = 140
         RECONO& = 1 + LOF(NKX%) / 140
         BUGRA$ = "MOVISTO " + MKS$(REMOVIST&) + RECORD$
         Put #NKX%, RECONO&, BUGRA$
         Close #NKX%
         '
         If App.EXEName = "FACTUMAI" Or App.EXEName = "FACTUDIP" Then
            INIBK& = 1 + ULTREG&("BKMOVIST")
            For UU& = INIBK& To REMOVIST&
               Call GRAREG("BKMOVIST", RELIBR$, UU&)
            Next UU&
            Call GRAREG("BKMOVIST", RECORD$, REMOVIST&)
         End If
         '
         SUCX% = IDEPOS%(DEPO%)
         If CI% > 0 Then
            If CI% + 1 <= ULTREG&("SALSTOCK") Then
               RGG& = CI% + 1: PX% = 4 * SUCX% - 3
               If PX% > 0 Then
                  If PX% <= 37 Then
                     STRI$ = REGLEIDO$("SALSTOCK", RGG&)
                     VALO# = CDbl(CANTI)
                     Call RECUPD(STRI$, PX%, 4, VALO#)
                     Call GRAREG("SALSTOCK", STRI$, RGG&)
                  End If
                  If PX% <= 245 Then
                     STRI$ = REGLEIDO$("SALDEPOS", RGG&)
                     VALO# = CDbl(CANTI)
                     Call RECUPD(STRI$, PX%, 4, VALO#)
                     Call REPLA(STRI$, MKS$(REMOVIST&), 249, 4)
                     Call REPLA(STRI$, MKI$(HOII%), 253, 2)
                     Call REPLA(STRI$, MKI$(Int(HORANUM(Time$))), 255, 2)
                     Call GRAREG("SALDEPOS", STRI$, RGG&)
                  End If
               End If
            End If
         End If
     Else               ' reparaciones
         Call COPYSTRU("MOVISTO", "MOVIREPA")
         Call REGAPP("MOVIREPA", RECORD$)
         REMOVIRE& = ULTREG&("MOVIREPA")
         Call CIERRARCH("MOVIREPA")
         '
         NKX% = FreeFile
         Open LUCOM$ + "\MOVISTO.LOC" For Random Access Read Write As #NKX% Len = 140
         RECONO& = 1 + LOF(NKX%) / 140
         BUGRA$ = "MOVIREPA" + MKS$(REMOVIRE&) + RECORD$
         Put #NKX%, RECONO&, BUGRA$
         Close #NKX%
         '
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
     DPNG% = Val(CONTROL$("", "DEPONEGA"))
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
   

Sub VERISAL(CIXI%)
   '
   Static HOII%
   If HOII% < 1 Then HOII% = HOY(HOS$)
   '
   SMP = Screen.MousePointer
   Call CARDEPOS
   UREMAS& = ULTREG&("MASTER")
   '
   CTSX% = 0
   If ULTREG&("SALSTOCK") < UREMAS& - 256 Then Screen.MousePointer = 11
   While ULTREG&("SALSTOCK") < UREMAS&
      X$ = String$(39, 0) + Chr$(1)
      Call REGAPP("SALSTOCK", X$)
      CTSX% = 1
   Wend
   '
   If ULTREG&("SALDEPOS") < UREMAS& - 256 Then Screen.MousePointer = 11
   While ULTREG&("SALDEPOS") < UREMAS&
      X$ = String$(256, 0)
      Call REPLA(X$, MKI$(HOII%), 253, 2)
      Call REPLA(X$, MKI$(Int(HORANUM(Time$))), 255, 2)
      Call REGAPP("SALDEPOS", X$)
      CTSX% = 1
   Wend
   '
   If CTSX% > 0 Then
      Call CIERRARCH("SALSTOCK")   ' para liberar bloqueos
      Call CIERRARCH("SALDEPOS")   ' para liberar bloqueos
   End If
   '
   REGSTO& = CIXI% + 1
   If CIXI% < 1 Or REGSTO& > UREMAS& Then GoTo 99
   '
   SST$ = REGLEIDO$("SALSTOCK", REGSTO&)
   SDP$ = REGLEIDO$("SALDEPOS", REGSTO&)
   If Left$(SDP$, 40) <> SST$ Then GoTo 10
   '
   UREM& = CVS(Mid$(SDP$, 249, 4))
   If UREM& < 1 Or UREM& >= ULTREG&("MOVISTO") Then GoTo 10
   '
   SALDETOT = 0
   For KI% = 1 To NUSU%
      SALDETOT = SALDETOT + CVS(Mid$(SDP$, 4 * KI% - 3, 4))
   Next KI%
   '
   XX$ = REGLEIDO$("MOVISTO", UREM&)
   If CVI(Mid$(XX$, 3, 2)) <> CIXI% Then
      GoTo 10
   End If
   '
   If Abs(CVS(Mid$(XX$, 118, 4)) - SALDETOT) >= 0.05 Then
      GoTo 10
   End If
   GoTo 99
   '
10 Call VESARRAS(CIXI%)
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
               Call MOVISTO(FECHA%, CI1%, LOTE$, CMOV$, DOPRI$, DOSEC$, REFE$, VUNI#, MONE$, NC%, DP1%, CANTI)
               '
               COD$ = LTrim$(RTrim$(CODEXT$(CI1%)))
               REFE$ = "Vale Mat.de Dep." + TRIM$(Str$(DP1%))
               CANTI = CAN
               Call MOVISTO(FECHA%, CI2%, LOTE$, CMOV$, DOPRI$, DOSEC$, REFE$, VUNI#, MONE$, NC%, DP2%, CANTI)
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
   J& = 0
   For U& = 1 To ULTREG&("VAREMAT")
     X$ = REGLEIDO$("VAREMAT", U&)
     If X$ <> String$(32, 0) Then
       J& = J& + 1
       Call GRAREG("VAREMAT", X$, J&)
     End If
   Next U&
   Call SETULTREG("VAREMAT", J&)
   Call CIERRARCH("VAREMAT")
   '
   OKEY% = 1
   If J& > 0 Then
     OKEY% = (-1)
   End If
   '
End Sub
'

Sub ENTRECONSI(NPROVE%)
        '
        Dim CANPREASI()
        '
        LLAMACOMPRA% = 0
        If NPROVE% > 0 Then LLAMACOMPRA% = 1
        '
        FORIPRE$ = TRIM$(CONTROL$("", "FOREMICO"))
        If FORIPRE$ <> "" Then
            LU$ = LUDAT$
            If EXISTE%(LU$ + FORIPRE$ + ".FRM") <> 1 Then
                RMCC$ = REPLACAR$(LU$ + FORIPRE$ + ".FRM", "\", "/")
                Call COMUNI("Error de Configuración:\  \No Existe Formulario de Impresión\de Remito en Consignación:\  \'" + RMCC$ + "'.")
                GoTo 9999
            End If
        End If
        '
        DEPOCON% = Val(CONTROL$("", "DEPOCONS"))
        DEPORIG% = DEPOCON%
        DEPOSI$ = ECOARCH$("TADEPOSI", Chr$(DEPORIG%))
        '
        'Call BLIMESS("Revisando Archivo de Movimientos de Stock")
1100    Screen.MousePointer = 11
        CMO$ = TRIM$(CONTROL$("", "CODRECON"))
        If CMO$ = "" Then CMO$ = "RC"
        FIN& = ULTREG&("MOVISTO")
        For U& = FIN& To 1 Step -1
            Call TRACE(20, FIN& - U& + 1, FIN&)
            XX$ = REGLEIDO$("MOVISTO", U&)
            If UCase$(Mid$(XX$, 21, 2)) = CMO$ Then
                UNRO& = Val(Mid$(XX$, 26, 7))
                GoTo 1002
            End If
        Next U&
        UNRO& = 0
        '
1002    REMI$ = CMO$ + "-" + TRIM$(Str$(UNRO& + 1))
        Call FRESHALL
        '
1105    Call CIERRARCH("*.*")
        ReDim CANPREASI(NUMAS%)
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
        DEPORIG% = Asc(Mid$(PEDY$, 121, 1))
        If DEPOCON% > 0 Then
          If DEPORIG% < 1 Or ECOARCH$("TADEPOSI", Chr$(DEPORIG%)) = "" Then
            Call MENSERR(24, "Depósito de Salida no Válido")
            DEPORIG% = Val(CONTROL$("", "DEPOCONS"))
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
        For I& = 1 To ULTREG&("MACONSIG")
            X$ = REGLEIDO$("MACONSIG", I&)
            If CVI(Mid$(X$, 3, 2)) = NC% Then
                CAPENE# = CVD(Mid$(X$, 55, 8)) - CVD(Mid$(X$, 63, 8))
                If CAPENE# > 0 Then
                    NUOC& = Val(Mid$(X$, 24, 6))
                    CI% = CVI(Mid$(X$, 33, 2))
                    STODISPO# = CDbl(STOCKACT(CI%) - CANPREASI(CI%))
                    If STODISPO# < 0 Then STODISPO# = 0
                    If STODISPO# < CAPENE# Then CAPENE# = STODISPO#
                    '
                    Y$ = REGBLAN$("!BOLREDET")
                    Call REPLA(Y$, MKI$(CI%), 83, 2)
                    Call REPLA(Y$, UNIMED$(CI%), 85, 4)
                    Call REPLA(Y$, MKD$(CAPENE#), 89, 8)
                    Call REPLA(Y$, MKS$(NUOC&), 151, 4)
                    Call REPLA(Y$, MKS$(I&), 189, 4)
                    HAYOC% = 1
                    Call REGAPP("!BOLREDET", Y$)
                    '
                    'acumula lo pre-asignado por caso de items que aparecen varias veces
                    If CI% > 0 Then
                      If CI% <= NUMAS% Then
                        CANPREASI(CI%) = CANPREASI(CI%) + CAPENE#
                      End If
                    End If
                    '
                End If
            End If
        Next I&
        '
        If HAYOC% = 0 Then
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
                GoTo 1114
              End If
            End If
          End If
        Next YX&
        Call MENSERR(24, "Consignación Cancelada !!!\Faltan Cantidades o Códigos.")
        GoTo 1113
        '
1114    If TRIM$(FORIPRE$) <> "" Then
           DTP$ = OBJPLA$("TRANSPOR", DATRANSPO$)
           If DTP$ = "" Then GoTo 1113         ' CANCELADO
           DATRANSPO$ = DTP$
        End If
        '
11141   CONTA% = ULTREG&("!BOLREDET")
        '
1115    If CONTA% > 0 Then
          '
          Call CIERRARCH("MOVISTO")
          Call BLOQARCH("MOVISTO")
          Call BLOQARCH("MACONSIG")
          Screen.MousePointer = 11  'Call BLIMESS("Registrando Recepcion ...")
          '
          AINDI% = 0
          For YX& = 1 To ULTREG&("!BOLREDET")
             YY$ = REGLEIDO$("!BOLREDET", YX&)
             CI% = CVI(Mid$(YY$, 83, 2))
             CAN# = CVD(Mid$(YY$, 89, 8))
             LOTE$ = Mid$(YY$, 69, 12)
             If CI% > 0 Then
               If CI% <= NUMAS% Then
                 If CAN# > 0.005 Then
                    NUOC& = CVS(Mid$(YY$, 151, 4))
                    REGOC& = CVS(Mid$(YY$, 189, 4))
                    DOSEC$ = ""
                    If REGOC& > 0 Then
                       If REGOC& <= ULTREG&("MACONSIG") Then
                          ZZ$ = REGLEIDO$("MACONSIG", REGOC&)
                          If CVI(Mid$(ZZ$, 3, 2)) = NC% Then
                             If CVI(Mid$(ZZ$, 33, 2)) = CI% Then
                                DOSEC$ = Mid$(ZZ$, 21, 12)
                                TOENTRE# = CVD(Mid$(ZZ$, 63, 8)) + CAN#
                                Call REPLA(ZZ$, MKD$(TOENTRE#), 63, 8)
                                If FF% > CVI(Mid$(ZZ$, 71, 2)) Then
                                    Call REPLA(ZZ$, MKI$(FF%), 71, 2)
                                End If
                                Call GRAREG("MACONSIG", ZZ$, REGOC&)
                                GoTo 1120
                             End If
                          End If
                       End If
                    End If
                    '
                    ZZ$ = REGBLAN$("MACONSIG")
                    Call REPLA(ZZ$, MKI$(FF%), 1, 2)
                    Call REPLA(ZZ$, MKI$(NC%), 3, 2)
                       NOCOM$ = TRIM$(Str$(NUOC&))
                       While Len(NOCOM$) < 6: NOCOM$ = "0" + NOCOM$: Wend
                       NOCOM$ = "OC-" + Left$(NOCOM$, 6) + "-00"
                    Call REPLA(ZZ$, NOCOM$, 21, 12)
                    Call REPLA(ZZ$, MKI$(CI%), 33, 2)
                    Call REPLA(ZZ$, UNIMED$(CI%), 51, 4)
                    Call REPLA(ZZ$, MKD$(CAN#), 55, 8)
                    Call REPLA(ZZ$, MKD$(CAN#), 63, 8)
                    Call REPLA(ZZ$, MKI$(FF%), 71, 2)
                    Call REGAPP("MACONSIG", ZZ$)
                    '
1120                CANTI = (-1) * CAN#
                    Call MOVISTO(FF%, CI%, LOTE$, CMO$, REMI$, DOSEC$, RASOCLI$(NC1%), 0, "$", NC1%, DEPORIG%, CANTI)
                 End If
               End If
             End If
             '
          Next YX&
          Call CIERRARCH("MOVISTO")
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
'

Sub DESCACONSI(NCX%, CIX%, CANX#)
    '
    CAREN# = 0
    CAPEN# = CANX#
    URECO& = ULTREG&("MACONSIG")
    RGCS$ = RECFILT$("MACONSIG", 2, MKI$(NC%))
    For I% = 1 To Len(RGCS$) Step 4
      REGCO& = CVS(Mid$(RGCS$, I%, 4))
      If REGCO& > 0 Then
        If REGCO& <= URECO& Then
          XCO$ = REGLEIDO$("MACONSIG", REGCO&)
          If CVI(Mid$(XCO$, 3, 2)) = NCX% Then
            If CVI(Mid$(XCO$, 33, 2)) = CIX% Then
               CACON# = CVD(Mid$(XCO$, 63, 8)) - CVD(Mid$(XCO$, 111, 8))
               If CACON# < 0 Then CACON# = 0
               CAREN# = CAPEN#: If CAREN# > CACON# Then CAREN# = CACON#
               CAPEN# = CAPEN# - CAREN#
               Call REPLA(XCO$, MKD$(CAREN# + CVD(Mid$(XCO$, 111, 8))), 111, 8)
               Call REPLA(XCO$, MKD$(CVD(Mid$(XCO$, 63, 8)) - CVD(Mid$(XCO$, 111, 8))), 119, 8)
               Call GRAREG("MACONSIG", XCO$, REGCO&)
            End If
          End If
        End If
      End If
    Next I%
    '
    If CAPEN# >= 0.05 Then
      CODDX$ = CODEXT$(CIX%)
      CAPENDE = CAPEN#
      Call LEEEXPLO(CODDX$, CAPENDE)
      For J% = 1 To CAIT%
        CIY% = CIJ%(J%)
        CANY# = CAPEN# * CDbl(USOI(J%))
        '
        For I% = 1 To Len(RGCS$) Step 4
          REGCO& = CVS(Mid$(RGCS$, I%, 4))
          If REGCO& > 0 Then
            If REGCO& <= URECO& Then
              XCO$ = REGLEIDO$("MACONSIG", REGCO&)
              If CVI(Mid$(XCO$, 3, 2)) = NCX% Then
                If CVI(Mid$(XCO$, 33, 2)) = CIY% Then
                  CACON# = CVD(Mid$(XCO$, 63, 8)) - CVD(Mid$(XCO$, 111, 8))
                  If CACON# < 0 Then CACON# = 0
                  CAREN# = CANY#: If CAREN# > CACON# Then CAREN# = CACON#
                  Call REPLA(XCO$, MKD$(CAREN# + CVD(Mid$(XCO$, 111, 8))), 111, 8)
                  Call REPLA(XCO$, MKD$(CVD(Mid$(XCO$, 63, 8)) - CVD(Mid$(XCO$, 111, 8))), 119, 8)
                  Call GRAREG("MACONSIG", XCO$, REGCO&)
                  CANY# = CANY# - CAREN#
                End If
              End If
            End If
          End If
        Next I%
        '
      Next J%
    End If
    '
    Call CIERRARCH("MACONSIG")
    '
End Sub
'
Sub DEVOLCONSI(RECONSI&, NCX%, CIX%, CANX#)
    '
    CAREN# = 0
    CAPEN# = CANX#
    URECO& = ULTREG&("MACONSIG")
    '
    REGCO& = Abs(RECONSI&)
    If REGCO& > 0 Then
      If REGCO& <= URECO& Then
        XCO$ = REGLEIDO$("MACONSIG", REGCO&)
        If CVI(Mid$(XCO$, 3, 2)) = NCX% Then
          If CVI(Mid$(XCO$, 33, 2)) = CIX% Then
            CACON# = CVD(Mid$(XCO$, 63, 8)) - CVD(Mid$(XCO$, 111, 8))
            If CACON# < 0 Then CACON# = 0
            CAREN# = CAPEN#: If CAREN# > CACON# Then CAREN# = CACON#
            CAPEN# = CAPEN# - CAREN#
            Call REPLA(XCO$, MKD$(CAREN# + CVD(Mid$(XCO$, 111, 8))), 111, 8)
            Call REPLA(XCO$, MKD$(CVD(Mid$(XCO$, 63, 8)) - CVD(Mid$(XCO$, 111, 8))), 119, 8)
            Call GRAREG("MACONSIG", XCO$, REGCO&)
          End If
        End If
      End If
    End If
    '
    Call CIERRARCH("MACONSIG")
    '
End Sub
'

Sub NUREMIFA(NROPRO&)
    '
    Screen.MousePointer = 11
    '
    CMO$ = TRIM$(CONTROL$("", "CODREFAC"))
    If CMO$ = "" Then CMO$ = "RT"
    '
    FENUE$ = CONTROL$("", "FENUREFA")
    If FENUE$ = "" Then FENUE$ = "18/04/03"
    FLIM% = FECHANUM(FENUE$)
    If FLIM% > HOY(HO$) Then FLIM% = HOY(HO$)
    '
    CAMOLE% = 0
    NROMA& = 0
    RETOLE& = 0
    EPAC$ = TRIM$(UCase$(EMPREAC$))
    
    For U& = ULTREG&("MOVISTO") To 1 Step -1
        XX$ = REGLEIDO$("MOVISTO", U&)
        RETOLE& = RETOLE& + 1
        If RETOLE& > 2048 Then
           If CAMOLE% >= 1 Then
              GoTo 82
           End If
        End If
        '
        If UCase$(Mid$(XX$, 21, 2)) = CMO$ Then
            If InStr(EPAC$, "NEUMANN") > 0 Then
              If Abs(Val(Mid$(XX$, 26, 7)) - Val(Mid$(XX$, 36, 7))) < 0.1 Then
                GoTo 80
              End If
            End If
            CAMOLE% = CAMOLE% + 1
            If CAMOLE% > 256 Then GoTo 82
            '
            FEMO% = CVI(Left$(XX$, 2))
            If FEMO% >= FLIM% Then
                NROLE& = Val(Mid$(XX$, 26, 7))
                If NROLE& = NROPRO& Then
                    NROPRO& = 0
                End If
                If NROLE& > NROMA& Then NROMA& = NROLE&
            End If
        End If
80  Next U&
    '
82  If EXISTE%(LUDAT$ + "MOVIREPA.RFS") < 1 Then
      Call COPYSTRU("MOVISTO", "MOVIREPA")
    End If
    '
    CAMOLE% = 0
    RETOLE& = 0
    For U& = ULTREG&("MOVIREPA") To 1 Step -1
        XX$ = REGLEIDO$("MOVIREPA", U&)
        RETOLE& = RETOLE& + 1
        If RETOLE& > 2048 Then
           If CAMOLE% >= 1 Then
              GoTo 85
           End If
        End If
        '
        If UCase$(Mid$(XX$, 21, 2)) = "RP" Then
            CAMOLE% = CAMOLE% + 1
            If CAMOLE% > 256 Then GoTo 85
            '
            FEMO% = CVI(Left$(XX$, 2))
            If FEMO% >= FLIM% Then
                NROLE& = Val(Mid$(XX$, 26, 7))
                If NROLE& = NROPRO& Then
                    NROPRO& = 0
                End If
                If NROLE& > NROMA& Then NROMA& = NROLE&
           End If
        End If
84  Next U&
    '
85  If NROPRO& < 1 Then NROPRO& = 1 + NROMA&
    Screen.MousePointer = 1
    '
End Sub
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

Sub CARCOMPO(CODICONJ$, NORFABRI$, CACONJU#, VACONTROL%, MODOXX%)
    '
    ' VACONTROL% = 1  presenta componentes para edicion
    '              0  no presenta componentes
    '             -1  devuelto como señal de cancelacion
    '
    If RECLEN%("MATEROF") <> 32 Then
       Call MENSERR(24, "Longitud de Registro Incorrecta\en Archivo 'MATEROF'.\Re-Setear a 32 Posiciones\y Refrescar Ventana 'MATEROF'.")
       On Error Resume Next
       Kill "MATEROF.PRF"
       On Error GoTo 0
       Call FINAL("")
    End If
    '
    Screen.MousePointer = 11
    Call SETULTREG("!MATEROF", 0)
    Call BLANARCH("!MATEROF")
    CAIT% = 0: COEF = 1
    Dim RERES&(1024)
    On Error GoTo 0
    '
    If Left$(NORFABRI$, 2) = "OR" Then
      CIXI% = 0
      VERISARRAS% = 1
      GoTo 50
    End If
    '
    VERISARRAS% = 0
    If CONTROL$("", "CHECOMPO") = "SI" Then
      VERISARRAS% = 1
    End If
    '
    CIXI% = CODINT%(CODICONJ$)
    If CIXI% < 1 Then
       Call MENSERR(24, "Error de Invocación de Rutina 'CARCOMPO'.\Código de Artículo '" + TRIM$(CODICONJ$) + "' Inexistente\Consulte a su Soporte de Sistemas.")
       Call FINAL("")
    End If
    CODPRO$ = CODICONJ$
    '
50  CAIT% = 0
    If TRIM$(NORFABRI$) <> "" Then
         TITUPA$ = "Consumo según Reservas"
         GoTo 100                ' lee de reservas
       ElseIf CODINT%(CODICONJ$) > 0 Then
         GoTo 200                ' lee de formula standard
       Else
         Call MENSERR(24, "Error de Invocación de Rutina 'CARCOMPO'.\Consulte a su Soporte de Sistemas.")
         Call FINAL("")
    End If
    '
100 NORFA$ = AJUSTI$(NORFABRI$, 12)
    NFX1$ = Mid$(NORFA$, 4)
    PPXX% = InStr(NFX1$, "-"): If PPXX% < 1 Then PPXX% = 1 + Len(NFX1$)
    NUORFA& = Val(Left$(NFX1$, PPXX% - 1))
    '
    RFF$ = RECFILT$("RESERVA", 3, NORFA$)
    For UU& = 1 To Len(RFF$) Step 4
       I& = CVS(Mid$(RFF$, UU&, 4))
'    For I& = 1 To ULTREG&("RESERVA")
       YY$ = REGLEIDO$("RESERVA", I&)
       If Mid$(YY$, 3, 12) = NORFA$ Then
         If Mid$(YY$, 19, 24) <> String$(24, 0) Then
           CAIT% = CAIT% + 1
           CIJ%(CAIT%) = CVI(Left$(YY$, 2))
           USOI(CAIT%) = CVS(Mid$(YY$, 75, 4))
           RERES&(CAIT%) = I&
         End If
       End If
'    Next I&
    Next UU&
    If CAIT% > 0 Or Left$(NORFABRI$, 2) = "OR" Then GoTo 300
    '
110 ' no encontro reservas - determina coeficiente s/pedidos
    NFX1$ = Mid$(NORFA$, 4)
    PPXX% = InStr(NFX1$, "-"): If PPXX% < 1 Then PPXX% = 1 + Len(NFX1$)
    NUORFA& = Val(Left$(NFX1$, PPXX% - 1))
    SUBOR% = Val(Mid$(NFX1$, PPXX% + 1))
    '
    QQ% = 0
    EPAC$ = TRIM$(UCase$(EMPREAC$))
    If InStr(EPAC$, "NOVOESP") < 1 Then GoTo 200
    '
    ABUSQUE$ = "PEDDETP"
    '
120 For III& = 1 To ULTREG&(ABUSQUE$)
       Y$ = REGLEIDO$(ABUSQUE$, III&)
       If Abs(CVS(Left$(Y$, 4)) - NUORFA&) < 0.1 Then
          QQ% = QQ% + 1
          If QQ% = SUBOR% Then
             If CVI(Mid$(Y$, 33, 2)) = CIXI% Then
                If CVS(Mid$(Y$, 59, 4)) > 1 Then
                   If CVS(Mid$(Y$, 63, 4)) > 1 Then
                      If CVS(Mid$(Y$, 67, 4)) > 1 Then
                         COEF = (CVS(Mid$(Y$, 59, 4))) / 1000
                         COEF = COEF * (CVS(Mid$(Y$, 63, 4))) / 1000
                         COEF = COEF * (CVS(Mid$(Y$, 67, 4))) / 1000
                         GoTo 200
                      End If
                   End If
                End If
             End If
          End If
       End If
    Next III&
    '
    EPAC$ = EMPREAC$
    If InStr(EPAC$, "NOVOESP") > 0 Then
       If ABUSQUE$ = "PEDDETP" Then
          ABUSQUE$ = "PEDDETA"
          GoTo 120
       End If
    End If
    '
200 ' arma composicion segun formula standard
    TITUPA$ = "Consumo por Fórmula Standard"
    Call LEEEXPLOMUL(CODPRO$, 1, 1)
    If CAIT% < 1 Then             ' producto sin estructura
        OPXX% = COMALTER%("No hay Fórmula de Consumo Standard\para este Producto (" + TRIM$(CODPRO$) + ").\\Cargar Fórmula\Continuar")
        If OPXX% = 1 Then
           Call CIERRARCH("*.*")
           PUNTCI$ = ""       ' para forzar reapertura de archivo
           Call CONECRUN("FORMULAS/" + CODPRO$, "Fórmulas y Estructuras de Producto/WAIT")
           Call FRESHALL
           GoTo 200
        End If
        TITUPA$ = "Consumo Informado"
    End If

300 FIN& = CAIT%
    Call SETULTREG("!MATEROF", 0)
    Call BLANARCH("!MATEROF")
    For I1% = 1 To CAIT%
       UI1& = I1%
       Call TRACE(20, UI1&, FIN&)
       If CIJ%(I1%) > 0 Then
          If CIJ%(I1%) <= NUMAS% Then
             If Abs(USOI(I1%)) > 0 Then
                '
                If VACONTROL% = 1 Then
                  If VERISARRAS% = 1 Then
                    Call VERISAL(CIJ%(I1%))
                    Call VESARRAS(CIJ%(I1%))
                  End If
                End If
                '
                ZZ$ = REGBLAN$("MATEROF")
                Call REPLA(ZZ$, MKI$(CIJ%(I1%)), 1, 2)
                Call REPLA(ZZ$, UNIMED$(CIJ%(I1%)), 3, 2)
                NEBRU = CACONJU# * USOI(I1%) * COEF
                Call REPLA(ZZ$, MKS$(NEBRU), 5, 4)
                Call REPLA(ZZ$, MKD$(CDbl(USOI(I1%) * COEF)), 9, 8)
                Call REPLA(ZZ$, MKS$(RERES&(I1%)), 17, 4)
                STDIS = STOCKDIS(CIJ%(I1%))
                If MODOXX% = 2 Then STDIS = STDIS + STORESOF(CIJ%(I1%), NUORFA&)
                Call REPLA(ZZ$, MKS$(STDIS), 21, 4)
                FALTA = NEBRU - STDIS: If FALTA < 0 Then FALTA = 0
                Call REPLA(ZZ$, MKS$(FALTA), 25, 4)
                Call REGAPP("!MATEROF", ZZ$)
                '
             End If
          End If
       End If
    Next I1%
    Call CIERRARCH("!MATEROF")
    Screen.MousePointer = 1
    
    If VACONTROL% = 1 Then
       TITUPA$ = TITUPA$ + ": "
       If TRIM$(NORFABRI$) <> "" Then TITUPA$ = TITUPA$ + TRIM$(NORFABRI$)
       If TRIM$(CODPRO$) <> "" Then TITUPA$ = TITUPA$ + " - " + TRIM$(CODPRO$) + " - " + TRIM$(DESCRIT$(CIXI%)) + " - Cant.=" + Str$(Int(CACONJU# + 0.5))
       '
       VTB% = VENTABU%("MATEROF/TITUPAN=" + TITUPA$)
       '
       If VTB% < 0 Then
            VACONTROL% = (-1)
          Else
            If Abs(CACONJU#) >= 0.05 Then
               For I& = 1 To ULTREG&("!MATEROF")
                  ZZ$ = REGLEIDO$("!MATEROF", I&)
                  CANUE# = CDbl(CVS(Mid$(ZZ$, 5, 4)))
                  USNUE# = CANUE# / CACONJU#
                  Call REPLA(ZZ$, MKD$(USNUE#), 9, 8)
                  Call GRAREG("!MATEROF", ZZ$, I&)
               Next I&
            End If
       End If
    End If
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
   FORIPRE$ = TRIM$(CONTROL$("", "REMIRECH"))
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
   LU$ = LUDAT$
   If EXISTE%(LU$ + "COKARDEX.RFS") < 1 Then
      Call COPYSTRU("MOVISTO", "COKARDEX")
   End If
   '
   CODD$ = CODL$
10 If TRIM$(CODL$) = "" Then
      Call SELECHO("MASTER")
      CODD$ = TRIM$(VALACT1$("MASTER"))
      If CODD$ = "" Then Exit Sub
   End If
   '
   CIDD% = CODINT%(CODD$)
   UREMO& = ULTREG&("MOVISTO")
   JJ& = 0
   If CIDD% > 0 Then
     If CIDD% <= NUMAS% Then
       RFF$ = RECFILT$("MOVISTO", 2, MKI$(CIDD%))
       If Len(RFF$) >= 4 Then
            Call BLOQARCH("COKARDEX")
            For U& = 1 To Len(RFF$) Step 4
              REMO& = CVS(Mid$(RFF$, U&, 4))
              If REMO& > 0 Then
                If REMO& <= UREMO& Then
                  X$ = REGLEIDO$("MOVISTO", REMO&)
                  Y$ = X$
                  JJ& = JJ& + 1
                  Call GRAREG("COKARDEX", Y$, JJ&)
                End If
              End If
            Next U&
            Call SETULTREG("COKARDEX", JJ&)
            Call CIERRARCH("COKARDEX")
            Call CONECRUN("*COKARDEX/" + CODD$, "Ficha Kardex Indirecta" + "/WAIT")
          Else
            Call MENSERR(24, "Codigo: " + TRIM$(CODD$) + "\sin Movimientos.")
       End If
     End If
   End If
   '
   If TRIM$(CODL$) = "" Then GoTo 10
   '
End Sub

Sub VALEMAT()

        DEPORIG% = Val(CONTROL$("", "DEPOCEN"))
        DEPOSI$ = ECOARCH$("TADEPOSI", Chr$(DEPORIG%))
        '
        'Call BLIMESS("Revisando Archivo de Movimientos de Stock")
1100    Screen.MousePointer = 11
        CMO$ = TRIM$(CONTROL$("", "CODCONSU"))
        If CMO$ = "" Then CMO$ = "VC"
        FIN& = ULTREG&("MOVISTO")
        For U& = FIN& To 1 Step -1
            Call TRACE(20, FIN& - U& + 1, FIN&)
            XX$ = REGLEIDO$("MOVISTO", U&)
            If UCase$(Mid$(XX$, 21, 2)) = CMO$ Then
                UNRO& = Val(Mid$(XX$, 26, 7))
                GoTo 1002
            End If
        Next U&
        UNRO& = 0
        '
1002    NROVAL& = UNRO& + 1
        Call FRESHALL
        '
1105    Call CIERRARCH("*.*")
        '
1111    HOII% = HOY(HOS$)
        VDEF$ = MKS$(NROVAL&) + MKI$(HOII%) + Chr$(DEPORIG%) + Chr$(0) + Space$(32)
        PEDY$ = OBJPLA$("VALEMATER", VDEF$)
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
1112    Call SETULTREG("!BOLREDET", 0)
        Call CIERRARCH("*.*")
        '
        Screen.MousePointer = 11
        '
        FFF$ = REPLACAR$(FECHATEX$(FF%), "/", "-")
1113    VTB% = VENTABU%("BOLREDET/NC/TITUPAN=Vale de Materiales - " + TRIM$(DEPOSI$) + " - " + FFF$)
        If VTB% < 0 Or ULTREG&("!BOLREDET") < 1 Then
            GoTo 1105
        End If
        '
        Call SETULTREG("!OBSERVOF", 0)
        Call CIERRARCH("!OBSERVOF")
        VTB% = VENTABU%("OBSERVOF/TITUPAN=Vale de Materiales - " + TRIM$(DEPOSI$) + " - " + FFF$)
        '
11141   CONTA% = ULTREG&("!BOLREDET")
        '
1115    If CONTA% > 0 Then
          '
          Call CIERRARCH("MOVISTO")
          Call BLOQARCH("MOVISTO")
          Screen.MousePointer = 11  'Call BLIMESS("Registrando Recepcion ...")
          '
          AINDI% = 0
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
1120                CANTI = (-1) * CAN#
                    Call MOVISTO(FF%, CI%, LOTE$, CMO$, REMI$, DOSEC$, REFE$, 0, "$", 0, DEPORIG%, CANTI)
                 End If
               End If
             End If
             '
          Next YX&
          Call CIERRARCH("MOVISTO")
          '
          FORIPRE$ = TRIM$(CONTROL$("", "FORCONSU"))
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
            '
          End If
          '
1199    End If
        Screen.MousePointer = 1
        GoTo 1100
        '
9999    Screen.MousePointer = 1
        Call CIERRARCH("*.*")
        '
End Sub

Sub DESCRAP()
        '
        EPAC$ = TRIM$(UCase$(EMPREAC$))
        DEPORIG% = Val(CONTROL$("", "DEPOCEN"))
        If InStr(EPAC$, "DOSIVA") > 0 Then
          DEPORIG% = 91
        End If
        DEPOSI$ = ECOARCH$("TADEPOSI", Chr$(DEPORIG%))
        '
        'Call BLIMESS("Revisando Archivo de Movimientos de Stock")
1100    Screen.MousePointer = 11
        CMO$ = TRIM$(CONTROL$("", "CODSCRAP"))
        If CMO$ = "" Then CMO$ = "SC"
        FIN& = ULTREG&("MOVISTO")
        For U& = FIN& To 1 Step -1
            Call TRACE(20, FIN& - U& + 1, FIN&)
            XX$ = REGLEIDO$("MOVISTO", U&)
            If UCase$(Mid$(XX$, 21, 2)) = CMO$ Then
                UNRO& = Val(Mid$(XX$, 26, 7))
                GoTo 1002
            End If
        Next U&
        UNRO& = 0
        '
1002    NROVAL& = UNRO& + 1
        Call FRESHALL
        '
1105    Call CIERRARCH("*.*")
        '
1111    HOII% = HOY(HOS$)
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
1112    Call SETULTREG("!BOLREDET", 0)
        Call CIERRARCH("*.*")
        '
        Screen.MousePointer = 11
        '
        FFF$ = REPLACAR$(FECHATEX$(FF%), "/", "-")
'
1113    VTB% = VENTABU%("BOLREDET/NC/TITUPAN=Descarte de Scrap - " + TRIM$(DEPOSI$) + " - " + FFF$)
        If VTB% < 0 Or ULTREG&("!BOLREDET") < 1 Then
            GoTo 1105
        End If
        '
        Call SETULTREG("!OBSERVOF", 0)
        Call CIERRARCH("!OBSERVOF")
        VTB% = VENTABU%("OBSERVOF/TITUPAN=Descarte de Scrap - " + TRIM$(DEPOSI$) + " - " + FFF$)
        '
11141   CONTA% = ULTREG&("!BOLREDET")
        '
1115    If CONTA% > 0 Then
          '
          Call CIERRARCH("MOVISTO")
          Call BLOQARCH("MOVISTO")
          Screen.MousePointer = 11  'Call BLIMESS("Registrando Recepcion ...")
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
1120                CANTI = (-1) * CAN#
                    Call MOVISTO(FF%, CI%, LOTE$, CMO$, REMI$, DOSEC$, REFE$, 0, "$", 0, DEPORIG%, CANTI)
                 End If
               End If
             End If
             '
          Next YX&
          Call CIERRARCH("MOVISTO")
          '
          FORIPRE$ = TRIM$(CONTROL$("", "FORSCRAP"))
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
1199    End If
        Screen.MousePointer = 1
        GoTo 1100
        '
9999    Screen.MousePointer = 1
        Call CIERRARCH("*.*")
        '
End Sub

Sub VERIDESPRE(REDES&, HUBORECU%)
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
   If DEPEX% < 1 Or DETER% < 1 Then
     DEPEX% = Int(Val(CONTROL$("", "DEPOEXPE")))
     DETER% = Int(Val(CONTROL$("", "DEPOTERM")))
     If DEPEX% < 1 Or DETER% < 1 Or DEPEX% = DETER% Then
       Call MENSERR(24, "Proceso Imposible.\  \Falta Configurar Depósitos\de Expedición\y Productos Terminados.")
       Call FINAL("")
     End If
   End If
   '
   If REDES& < 1 Then
     GoSub CARYENANT
     Exit Sub
   End If
   '
   HUBORECU% = 0
   X$ = REGLEIDO$("DESPAPED", REDES&)
   STATX% = Asc(Mid$(X$, 27, 1))
   REPARX% = 0: If Mid$(X$, 51, 1) = "R" Then REPARX% = 1
   '
   NPX& = CVS(Left$(X$, 4))
   FEMIPED% = CVI(Mid$(X$, 5, 2))
   NC% = CVI(Mid$(X$, 7, 2))
   CI% = CVI(Mid$(X$, 33, 2))
   CAPEDID = CVS(Mid$(X$, 75, 4))  ' CANTIDAD PEDIDA
   If CI% > 0 And CI% <= NUMAS% Then
     ACUPRE = 0
     ACUYEN = 0
     CAVERI$ = Mid$(X$, 113, 10)
     CAVERJ$ = MKS$(Val(Left$(CAVERI$, 6))) + MKI$(Val(Mid$(CAVERI$, 8, 3)))
     '
     RFF$ = RECFILT$("MOVISTO", 2, MKI$(CI%)) ' SELECCIONA MOVIMIENTOS DEL ITEM
     FIN& = Len(RFF$)
     For U& = 1 To FIN& Step 4
       REMO& = CVS(Mid$(RFF$, U&, 4))
       MM$ = REGLEIDO$("MOVISTO", REMO&)
       If Asc(Mid$(MM$, 83, 1)) = DEPEX% Then      ' MOVIMIENTO EN DEPODITO DE EXPEDICION
         IDPEDIT$ = Mid$(MM$, 112, 6)              ' LEE PEDIDO E ITEM
         If IDPEDIT$ = CAVERJ$ Then
           TIMO$ = UCase$(Mid$(MM$, 21, 2))
           If TIMO$ = "PP" Then
               ACUPRE = ACUPRE + CVD(Mid$(MM$, 88, 8)) - CVD(Mid$(MM$, 96, 8))
             ElseIf TIMO$ = "DM" Then
               ACUYEN = ACUYEN + CVD(Mid$(MM$, 96, 8)) - CVD(Mid$(MM$, 88, 8))
           End If
         End If
       End If
     Next U&
     '
     If FEMIPED% < FECHANUM("07/11/05") Then
       ' IR A BUSCAR DESPACHOS REALIZADOS CON ANTERIORIDAD
       ' DEVOLVIENDO SOBRE CAYENANT
       NPEDXX& = NPX&
       GoSub CARYENANT
       ACUYEN = ACUYEN + CAYENANT
     End If
     '
     '
     ' CORRECCION DEL 28/02/06 - SI DESPACHADO ES > PREPARADO
     ' ENTONCES PREPARADO = DESPACHADO
     If ACUYEN > ACUPRE Then
       ACUPRE = ACUYEN
     End If
     '
     ' GRABA LOS VALORES ACUMULADOS DE PREPARACION Y DESPACHO
     If STATX% <> 1 Then
       If Abs(ACUPRE - CVS(Mid$(X$, 79, 4))) > 0.05 Then HUBORECU% = 1
       If Abs(ACUYEN - CVS(Mid$(X$, 95, 4))) > 0.05 Then HUBORECU% = 1
     End If
     '
     Call REPLA(X$, MKS$(ACUPRE), 79, 4)
     Call REPLA(X$, MKS$(ACUYEN), 95, 4)
     If STATX% <= 3 Then
       STATNUE% = 11
       If ACUPRE >= CAPEDID - 0.05 Then
         If ACUYEN >= CAPEDID - 0.05 Then
           STATNUE% = 3
         End If
       End If
       Call REPLA(X$, Chr$(STATNUE%), 27, 1)
     End If
     Call GRAREG("DESPAPED", X$, REDES&)
     '
   End If
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
   RFKL$ = RECFILT$("MOVISTO", 2, MKI$(CI%))
   '
   For KKU& = 1 To Len(RFKL$) Step 4
     REMOVI& = CVS(Mid$(RFKL$, KKU&, 4))
     XMS$ = REGLEIDO$("MOVISTO", REMOVI&)
     If TRIM$(Mid$(XMS$, 23, 10)) = ABUPED$ Then
       If CVI(Left$(XMS$, 2)) < FECHANUM("07/11/05") Then
         CADESPAR = CVD(Mid$(XMS$, 96, 8))
         If CADESPAR >= 0.005 Then
           CAYADES(KKF%) = CAYADES(KKF%) + CADESPAR
         End If
       End If
     End If
   Next KKU&
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

Sub ATRANSFE()
   '
   ' GENERA LOS ARCHIVOS DE TRANSFERENCIA EN EL FORMATO ESPECIFICADO POR ADSUR
   '
   Call DESHASFECHA(DES%, HAS%, PERIO$)
   If PERIO$ <> "" Then
      Screen.MousePointer = 11
      '
      ' GRABA TABLA PARA COMPRAS
      NX1% = FILEOPEN%("C:\FLEXOFT\TRANCOMP.TXT", 2, 0)
      FIN& = ULTREG&("OCOMDETA")
      JJ1& = 0: NOCOA& = 0
      For U& = 1 To FIN&
         Call TRACE(20, U&, FIN&)
         XXX$ = REGLEIDO$("OCOMDETA", U&)
         FEMISI = CVI(Mid$(XXX$, 5, 2))
         If FEMISI >= DES% Then
           If FEMISI <= HAS% Then
             If Asc(Mid$(XXX$, 124, 1)) < 9 Then ' EXCLUYE LAS ANULADAS
                MONEMI% = Asc(Mid$(XXX$, 122, 1))
                If MONEMI% <= 1 Then MONEX$ = "$   "
                If MONEMI% = 2 Then MONEX$ = "U$S "
                If MONEMI% > 2 Then MONEX$ = Left$(ECOARCH$("TAMONED", Chr$(MONEMI%)), 3)
                '
                NUOCO& = CVS(Mid$(XXX$, 1, 4))
                   If NUOCO& <> NOCOA& Then NORIT% = 0
                   NOCOA& = NUOCO&
                NORIT% = NORIT% + 1
                NCX% = (-1) * CVI(Mid$(XXX$, 7, 2))
                RAPRO$ = RASOCLI$(NCX%)
                CUITP$ = CUITCLI$(NCX%)
                  CIXI% = CVI(Mid$(XXX$, 9, 2))
                CODAR$ = CODEXT$(CIXI%)
                DECOR$ = DESCRIT$(CIXI%)
                  ' AQUI TEXTO LIBRE
                  RFF$ = RECFILT$("TXOCOLIB", 1, MKS$(NUOCO&))
                  If Len(RFF$) >= 4 Then
                     For KKL& = 1 To Len(RFF$) Step 4
                        RETX& = CVS(Mid$(RFF$, KKL&))
                        XYX$ = REGLEIDO$("TXOCOLIB", RETX&)
                        DECOR$ = TRIM$(Mid$(XYX$, 5, 64))
                        GoTo 15
                     Next KKL&
                  End If
15                UMEDI$ = Mid$(XXX$, 59, 4)
                CANSOL = CVS(Mid$(XXX$, 55, 4))
                FESENT = CVI(Mid$(XXX$, 95, 2))
                MONEV$ = MONEX$
                PRUNIT = CVD(Mid$(XXX$, 79, 8))
                '
                TTXX$ = Space$(220)
                Call REPLA(TTXX$, FORMATNUM$(NUOCO&, "F6.0"), 1, 6)
                Call REPLA(TTXX$, FORMATNUM$(NORIT%, "I4"), 7, 4)
                Call REPLA(TTXX$, RAPRO$, 13, 48)
                Call REPLA(TTXX$, CUITP$, 61, 16)
                Call REPLA(TTXX$, FECHATEX$(FEMISI), 77, 8)
                Call REPLA(TTXX$, CODAR$, 87, 16)
                Call REPLA(TTXX$, DECOR$, 103, 64)
                Call REPLA(TTXX$, UMEDI$, 167, 4)
                Call REPLA(TTXX$, FORMATNUM$(CANSOL, "F16.4"), 171, 16)
                Call REPLA(TTXX$, FECHATEX(FESENT), 189, 8)
                Call REPLA(TTXX$, MONEV$, 199, 4)
                Call REPLA(TTXX$, FORMATNUM$(PRUNIT, "F16.4"), 203, 16)
                '
                JJ1& = JJ1& + 1
                Print #NX1%, TTXX$
             End If
           End If
         End If
      Next U&
      '
      Close #NX1%
      Call CIERRARCH("OPENDETA")
      Screen.MousePointer = 1
      '
      ' GRABA TABLA PARA RECEPCIONES
      '
      NX1% = FILEOPEN%("C:\FLEXOFT\TRANRECE.TXT", 2, 0)
      FIN& = ULTREG&("BOLRECEP")
      JJ2& = 0: NOCOA& = 0
      For U& = 1 To FIN&
         Call TRACE(20, U&, FIN&)
         XXX$ = REGLEIDO$("BOLRECEP", U&)
         '
         FERECE = CVI(Mid$(XXX$, 81, 2))
         If FERECE >= DES% Then
           If FERECE <= HAS% Then
             MONEX$ = TRIM$(Mid$(XXX$, 155, 4))
             If MONEX$ = "" Then MONEX$ = "$   "
             '
             NUBORE& = CVS(Mid$(XXX$, 1, 4))
                If NUBORE& <> NOCOA& Then NORIT% = 0
                NOCOA& = NUBORE&
             NORIT% = NORIT% + 1
             NCX% = (-1) * CVI(Mid$(XXX$, 105, 2))
             RAPRO$ = RASOCLI$(NCX%)
             CUITP$ = CUITCLI$(NCX%)
               CIXI% = CVI(Mid$(XXX$, 83, 2))
             CODAR$ = CODEXT$(CIXI%)
             DECOR$ = DESCRIT$(CIXI%)
               ' AQUI TEXTO LIBRE
               RFF$ = RECFILT$("TXOCOLIB", 3, MKS$(NUBORE&))
               If Len(RFF$) >= 4 Then
                  For KKL& = 1 To Len(RFF$) Step 4
                     RETX& = CVS(Mid$(RFF$, KKL&))
                     XYX$ = REGLEIDO$("TXOCOLIB", RETX&)
                     DECOR$ = TRIM$(Mid$(XYX$, 5, 64))
                     GoTo 25
                  Next KKL&
               End If
25           UMEDI$ = Mid$(XXX$, 85, 4)
             REMI$ = TRIM$(Mid$(XXX$, 119, 16))
             NFAC$ = TRIM$(Mid$(XXX$, 135, 16))
             NUOCO& = CVS(Mid$(XXX$, 151, 4))
             CANRECI = CVD(Mid$(XXX$, 89, 8))
             CANAPRO = CVD(Mid$(XXX$, 163, 8))
             CANRECH = CVD(Mid$(XXX$, 171, 8))
             CANFALT = CVD(Mid$(XXX$, 179, 8))
             MONEV$ = MONEX$
             PRUNIT = CVS(Mid$(XXX$, 159, 4))
             '
             TTXX$ = Space$(295)
             Call REPLA(TTXX$, FORMATNUM$(NUBORE&, "F6.0"), 1, 6)
             Call REPLA(TTXX$, FORMATNUM$(NORIT%, "I4"), 7, 4)
             Call REPLA(TTXX$, RAPRO$, 13, 48)
             Call REPLA(TTXX$, CUITP$, 61, 16)
             Call REPLA(TTXX$, FECHATEX$(FERECE), 77, 8)
             Call REPLA(TTXX$, REMI$, 87, 16)
             Call REPLA(TTXX$, NFAC$, 103, 16)
             Call REPLA(TTXX$, FORMATNUM$(NUOCO&, "F6.0"), 119, 6)
             Call REPLA(TTXX$, CODAR$, 127, 16)
             Call REPLA(TTXX$, DECOR$, 143, 64)
             Call REPLA(TTXX$, UMEDI$, 207, 4)
             Call REPLA(TTXX$, FORMATNUM$(CANRECI, "F16.4"), 211, 16)
             Call REPLA(TTXX$, FORMATNUM$(CANAPRO, "F16.4"), 227, 16)
             Call REPLA(TTXX$, FORMATNUM$(CANRECH, "F16.4"), 243, 16)
             Call REPLA(TTXX$, FORMATNUM$(CANFALT, "F16.4"), 259, 16)
             Call REPLA(TTXX$, MONEV$, 275, 4)
             Call REPLA(TTXX$, FORMATNUM$(PRUNIT, "F16.4"), 279, 16)
             '
             JJ2& = JJ2& + 1
             Print #NX1%, TTXX$
           End If
         End If
      Next U&
      '
      Close #NX1%
      Call CIERRARCH("BOLRECEP")
      Screen.MousePointer = 1
      '
40    TTXX$ = "Archivos de Transferencia 'TRANCOMP.TXT' y 'TRANRECE.TXT'\Generados Exitosamente en 'C:/FLEXOFT'\con " + TRIM$(Str$(JJ1&)) + " / " + TRIM$(Str$(JJ2&)) + " Registros."
      TTXX$ = TTXX$ + "\   \¿ Deser Visualizar los Archivos Generados ?\\Si, Compras\Si, Recepción\No, Salir"
      OPX% = COMALTER%(TTXX$)
      If OPX% < 1 Or OPX% > 2 Then Exit Sub
      '
      ATRAN$ = "TRANCOMP": If OPX% = 2 Then ATRAN$ = "TRANRECE"
      If EXISTE%("C:\FLEXOFT\WORDPAD.EXE") = 1 Then
             Shell "C:\FLEXOFT\WORDPAD.EXE " + "C:\FLEXOFT\" + ATRAN$ + ".TXT", vbMaximizedFocus
             'PROLLAMA$ = TRIM$(Text2) + " - Wordpad"
             GoTo 40
           Else
             MENSA$ = "Imposible Mostrar Archivo Generado."
             MENSA$ = MENSA$ + "\  \Falta Utilitario 'WORDPAD.EXE'"
             MENSA$ = MENSA$ + "\en Directorio 'C:/FLEXOFT'."
             MENSA$ = MENSA$ + "\   \Instale este Programa Utilitario\y Re-Genere el Listado."
             Call COMUNI(MENSA$)
      End If
   End If
End Sub
