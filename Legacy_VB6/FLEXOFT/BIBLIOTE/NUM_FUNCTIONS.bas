Attribute VB_Name = "NUM_FUNCTIONS"
Function NUMSTRI$(VALO#, LCC%, DC%, SEMIL%)
    '
    FMTX$ = String$(DC%, "0")
    If FMTX$ <> "" Then FMTX$ = "0." + FMTX$
    While Len(FMTX$) < LCC%
      FMTX$ = "#" + FMTX$
    Wend
    NMST$ = Format(VALO#, FMTX$)
    If Len(NMST$) > LCC% Then NMST$ = String$(LCC%, "*")
    NUMSTRI$ = AJUSTD$(REPLACAR$(NMST$, ",", "."), LCC%)
    If InStr(NUMSTRI$, "*") < 1 Then
      If XVALO(NUMSTRI$) = 0 Then
        NUMSTRI$ = Space$(LCC%)
      End If
    End If
    Exit Function
    '
End Function
Function Solonumeros(KEYAS)
EJEMPLO:
'Private Sub TXTFTECNICACAMB_ANCHO_CM_KeyPress(KeyAscii As Integer)
'   KeyAscii = Solonumeros(KeyAscii)
'End Sub
'SOLO NUMEROS,PUNTO, COMA, ESPACIO, RETROCESO
   Select Case KEYAS
       Case Is = 2, 3, 8, 32, 44, 46, 48, 49, 50, 51, 52, 53, 54, 55, 56, 57, 127
         Solonumeros = KEYAS
       Case Else
        Solonumeros = 0
    End Select

End Function

Function NUMSTRI0$(VALO#, LCC%, DC%, SEMIL%)
   '
   NST0$ = TRIM$(NUMSTRI$(VALO#, LCC%, DC%, SEMIL%))
   While Len(NST0$) < LCC%
      NST0$ = "0" + NST0$
   Wend
   NUMSTRI0$ = NST0$
   '
End Function

Function FORMATNUM$(valor, Formato$)
   '
   CENTRA% = 0
   FORMATX$ = Formato$
   PPXX% = InStr(Formato$, "^")
   If PPXX% > 0 Then
      FORMATX$ = Left$(Formato$, PPXX% - 1) + Mid$(Formato$, PPXX% + 1)
      CENTRA% = 1
   End If
   '
   FTM$ = ""
   VTP% = VarType(valor)
   '
   If VTP% = 2 Then
      If Left$(UCase$(FORMATX$), 1) = "C" Then
         CIXI5% = valor
         LOCAMPO% = XVALO(Mid$(FORMATX$, 2))
         FTM$ = AJUSTI$(CODEXT$(CIXI5%), LOCAMPO%)
         GoTo 80
      End If
   End If
   '
   If Left$(UCase$(Formato$), 1) = "H" Then
      If InStr(CStr(valor), ":") > 0 Then
           VTEX$ = CStr(valor)
           VNUM = HORANUM(VTEX$)
        Else
           VNUM = valor
           'If Len(CStr(VNUM)) <= 2 Then VNUM = VALOR & "00"
      End If
      LONGI = XVALO(Mid$(Formato$, 2))
      FTM$ = Left$(HORATEX$(VNUM), LONGI)
      GoTo 80
   End If
   '
   If VTP% >= 2 And VTP% <= 6 Then
        VALODO# = valor
        VX$ = MKD$(VALODO#)
        If Left$(UCase$(FORMATX$), 1) = "X" Then
           VX$ = FORMATNUM$(VALODO#, "F18.8")
        End If
        FTM$ = TEXTBIN$(VX$, FORMATX$)
      ElseIf VTP% = 7 Then     ' FECHA Y HORA
        AFTM% = XVALO(Mid$(FORMATX$, 2))
        If AFTM% > 8 Then
             On Error Resume Next
             FTM$ = AJUSTI$(Format(valor, "dd/mm/yy - hh:mm:ss"), AFTM% - 1) + " "
             If Err Then
                On Error GoTo 0
                GoTo 25
             End If
             On Error GoTo 0
          Else
25           VX$ = MKI$(CVINT(valor))
             FTM$ = TEXTBIN$(VX$, FORMATX$)
        End If
      ElseIf VTP% = 8 Then
        If UCase$(Left$(FORMATX$, 1)) = "A" Or UCase$(Left$(FORMATX$, 1)) = "X" Then
            VX$ = valor
          ElseIf UCase$(Left$(FORMATX$, 1)) = "D" Then
            If IsNumeric(valor) = False Then
                 FETX$ = valor
                 VALORII% = FECHANUM(FETX$)
               Else
                 On Error Resume Next
                 VALORII% = valor
                 On Error GoTo 0
            End If
            VX$ = MKI$(VALORII%)
          Else
            VALODO# = XVALO(valor)
            VX$ = MKD$(VALODO#)
        End If
        FTM$ = TEXTBIN$(VX$, FORMATX$)
   End If
   '
80 FORMATNUM$ = FTM$
   If CENTRA% = 1 Then
      FTM$ = TRIM$(FTM$)
      LOCAMPO% = XVALO(Mid$(FORMATX$, 2))
      While Len(FTM$) < LOCAMPO%
         FTM$ = " " + FTM$ + " "
      Wend
      FTM$ = Left$(FTM$, LOCAMPO%)
   End If
   '
End Function

Function FORMATSTAN$(valor, Formato$, CIXI%)
   '
   PPXX% = InStr(Formato$, ".")
   If PPXX% < 1 Then PPXX% = 1 + Len(Formato$)
   FORXX$ = Left$(Formato$, PPXX% - 1) + "." + CNTDC$
   '
   FORSTA$ = FORMATNUM$(valor, FORXX$)
   LOFOR% = Len(FORSTA$)
   CND% = XVALO(CNTDC$(CIXI%))
   PPXX% = InStr(FORSTA$, ".")
   If PPXX% > 0 Then
     If CND < 1 Then
         FORSTA$ = Left$(FORSTA$, PPXX% - 1)
       Else
         FORSTA$ = Left$(FORSTA$, PPXX% + CND)
     End If
     While Len(FORSTA$) < LOFOR%
        FORSTA$ = FORSTA$ + " "
     Wend
   End If
   '
   FORMATSTAN$ = FORSTA$
   '
End Function
'
Function TMASCARA$(FPPX$)
    If Left$(FPPX$, 1) = "F" Then
         valor = "1.23456789"
         VATEX$ = FORMATNUM$(valor, FPPX$)
         LOREF% = Len(VATEX$)
         VATEX$ = TRIM$(VATEX$)
         VINI% = 1
         While Len(VATEX$) < LOREF%
            VINI% = VINI% + 1 Mod 10
            VATEX$ = TRIM$(Str$(VINI%)) + VATEX$
         Wend
      ElseIf Left$(FPPX$, 1) = "D" Then
         If FPPX$ = "D" Then FPPX$ = "D8"
         If Val(Mid$(FPPX$, 2)) < 20 Then
             valor = FECHANUM("09/09/09")
             VATEX$ = FORMATNUM$(valor, FPPX$)
           Else
             valor = FECHANUM("29/12/09")
             VATEX$ = FETEXTO$(Int(valor))
         End If
         LOREF% = XVALO(Mid$(FPPX$, 2))
         If LOREF% < 1 Then LOREF% = 8
         VATEX$ = TRIM$(Left$(VATEX$, LOREF%))
      ElseIf Left$(FPPX$, 1) = "H" Or Left$(FPPX$, 1) = "T" Then
         If FPPX$ = "H" Or FPPX$ = "T" Then FPPX$ = "T5"
         valor = "22:22:22"
         VATEX$ = FORMATNUM$(valor, FPPX$)
         LOREF% = XVALO(Mid$(FPPX$, 2))
         If LOREF% < 1 Then LOREF% = 5
         VATEX$ = TRIM$(Left$(VATEX$, LOREF%))
      ElseIf Left$(FPPX$, 1) = "A" Then
         valor = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
         LOREF% = XVALO(Mid$(FPPX$, 2))
         While Len(valor) < LOREF%: valor = valor + "ABCDEFGHIJKLMNOPQRSTUVWXYZ": Wend
         VATEX$ = FORMATNUM$(valor, FPPX$)
      ElseIf Left$(FPPX$, 1) = "C" Then
         valor = "0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ"
         LOREF% = XVALO(Mid$(FPPX$, 2))
         VATEX$ = FORMATNUM$(valor, "A" + Mid$(FPPX$, 2))
    End If
    TMASCARA$ = VATEX$
End Function
'
Function TEXTBIN$(VX$, Formato$)

FRX$ = UCase$(TRIM$(Formato$))
If FRX$ = "" Then
   VD$ = VX$
   GoTo 90
End If
LETRA$ = Left$(FRX$, 1)
'
If LETRA$ = "I" Then
     TV% = 2
  ElseIf LETRA$ = "D" Then
     TV% = 5
  ElseIf LETRA$ = "A" Then
     TV% = 6
  ElseIf LETRA$ = "C" Then
     TV% = 7
  ElseIf LETRA$ = "F" Or LETRA$ = "G" Or LETRA$ = "$" Then
     TV% = 4
  ElseIf LETRA$ = "H" Or LETRA$ = "T" Then
     TV% = 8
  ElseIf LETRA$ = "X" Then
     Moneda$ = DEMONECO$(MONEMI%)
     VD$ = TEXTFORM$(RTrim(VX$), Formato$, Moneda$)
     GoTo 90
  Else
5    Call MENSERR(24, "Error Invocacion Funcion 'TXN' - Formato '" + Formato$ + "' no Procesable")
     Close: End ' Call FINAL("")
End If
'
FRX$ = Mid$(FRX$, 2)
LCC% = Int(XVALO(FRX$))
If LCC% < 1 Then LCC% = 1 '  GoTo 5                                ' formato no procesable


Lx% = Len(VX$)
If Lx% >= 8 Then
     VALO# = CVD(Left$(VX$, 8))
  ElseIf Lx% >= 4 Then
     VALO# = CDbl(CVS(Left$(VX$, 4)))
  ElseIf Lx% >= 2 Then
     VALO# = CDbl(CVI(Left$(VX$, 2)))
  ElseIf Lx% >= 1 Then
     VALO# = CDbl(Asc(Left$(VX$, 1)))
  Else
     If TV% <> 6 Then Call MENSERR(24, "Error Invocacion Funcion 'TXN' - String Nulo")
End If

On TV% GoTo 40, 40, 40, 40, 50, 60, 70, 80
'
40 LD% = 10 * (XVALO(FRX$) - Int(XVALO(FRX$)))
   VD$ = NUMSTRI$(VALO#, LCC%, LD%, 0)
   If LETRA$ = "$" Or LETRA$ = "G" Then
     If TRIM$(VD$) <> "" Then
         VDX$ = REPLACAR$(TRIM$(VD$), ".", ",") + ","
         While (InStr(VDX$, ",") Mod 3) <> 1
           VDX$ = " " + VDX$
         Wend
         VDY$ = ""
42       VDY$ = VDY$ + Left$(VDX$, 3)
         VDX$ = Mid$(VDX$, 4)
         If Left$(VDX$, 1) <> "," And TRIM$(VDY$) <> "-" Then
           VDY$ = VDY$ + "."
           GoTo 42
         ElseIf TRIM$(VDY$) = "-" Then
           GoTo 42
         End If
         VDY$ = VDY$ + VDX$
         If Right$(VDY$, 1) = "," Then
           VDY$ = Left$(VDY$, Len(VDY$) - 1)
         End If
         '
         VD$ = VDY$
         ' sI ES FORMATO "$" AGREGA EL SIMBOLO DE LA MONEDA
         SIMBO$ = ""
         If LETRA$ = "$" Then SIMBO$ = SIMBOLPES$(MONEMI%)
         If InStr(FRX$, "<") > 0 Then
             VD$ = TRIM$(SIMBO$ + " " + VDY$)
           Else
             While Len(VDY$) < LCC% - Len(SIMBO$)
               VDY$ = " " + VDY$
             Wend
             VD$ = SIMBO$ + VDY$
         End If
     End If
     '
   End If
45 If InStr(FRX$, "<") > 0 Then
       VD$ = AJUSTI$(TRIM$(VD$), LCC%)
   End If
   GoTo 90
   '
50 FF = VALO#
   VD$ = AJUSTI$(FECHATEX$(FF), LCC%)
   If LCC% >= 20 Then
      VD$ = AJUSTI$(FETEXTO$(Int(FF)), LCC%)
   End If
   GoTo 90
   '
60 VD$ = AJUSTI$(TRIM$(VX$), LCC%)
   If InStr(FRX$, ">") > 0 Then
       VD$ = TRIM$(VX$)
       While Len(VD$) < LCC%: VD$ = " " + VD$: Wend
       VD$ = Right$(VD$, LCC%)
     ElseIf InStr(FRX$, "^") > 0 Then
       VD$ = TRIM$(VD$)
       While Len(VD$) < LCC%: VD$ = " " + VD$ + " ": Wend
       VD$ = Left$(VD$, LCC%)
   End If
   '
   GoTo 90
70 On Error Resume Next
   CI% = Int(VALO# + 0.5)
   On Error GoTo 0
   VD$ = AJUSTI$(CODEXT$(CI%), LCC%)
   GoTo 90
   '
80 HH% = Int(VALO# + 0.5)
   VD$ = AJUSTI$(HORATEX$(HH%), LCC%)
   GoTo 90
   '
90 If FRX$ = "0" Then VD$ = ""
   TEXTBIN$ = VD$
   End Function
'

Function VALIMI(A$, MINI, MAXI)
    '
    valor = XVALO(A$)
    If valor < MINI Or valor > MAXI Then
        valor = MINI
    End If
    VALIMI = valor
    '
End Function
'

Function FORMDOC$(NDOC&, TI%, VA%, NUSUCUVE%)
'
TISS$ = "FC": If TI% = 0 Then TISS$ = "AS"
         If TI% = 1 Then TISS$ = "VC"
         If TI% = 3 Then TISS$ = "ND"
         If TI% = 4 Then TISS$ = "RB"
         If TI% = 5 Then TISS$ = "DC"
         If TI% = 6 Then TISS$ = "NC"
         If TI% = 7 Then TISS$ = "AS"
         '
 VASS$ = "A": If VA% = 2 Then VASS$ = "B"
         If VA% = 3 Then VASS$ = "C"
         '
NDOCSS$ = TRIM$(Str$(NDOC&))
         While Len(NDOCSS$) < 8
            NDOCSS$ = "0" + NDOCSS$
         Wend
         '
SUCUI% = NUSUCUVE%: If SUCUI% < 1 Then SUCUI% = 1
NSUCUS$ = TRIM$(Str$(SUCUI%))
         While Len(NSUCUS$) < 4
            NSUCUS$ = "0" + NSUCUS$
         Wend
         '
FORMDOC$ = TISS$ + "-" + NSUCUS$ + "-" + NDOCSS$ + "-" + VASS$
'
End Function
'
Function COEFDESCU(DESC$)
    
    COEFD = 1
    DESCY$ = TRIM$(REPLACAR$(DESC$, "%", " ")) + "+"
    DESCX$ = TRIM$(REPLACAR$(DESCY$, ",", "."))
    '
10  PPXX% = InStr(DESCX$, "+")
    DESC1 = XVALO(Left$(DESCX$, PPXX% - 1))
    DESCX$ = TRIM$(Mid$(DESCX$, PPXX% + 1))
    COEFD = COEFD * (1 - DESC1 / 100)
    If DESCX$ <> "" Then GoTo 10
    '
99  COEFDESCU = 100 * (1 - COEFD)

End Function

Function REDONVAL#(valor, CADECI%)
    CADEX% = CADECI%
    If CADEX% < 0 Then CADEX% = 0
    COEFXX = 10 ^ CADEX%
    REDONVAL# = (Int(COEFXX * valor + 0.5)) / COEFXX
End Function

Function ROUND#(valor, Optional NDECI%)
   If NDECI% < 1 Then NDECI% = 2
   VALORX = XVALO(valor)
   If IsNull(valor) Or valor = "" Then VALORX = 0
   SIG1% = Sgn(VALORX)
   VALO1# = Abs(VALORX)
   FACTO& = 10 ^ NDECI%
   ROUND# = SIG1% * (Int(FACTO * VALO1# + 0.5)) / FACTO&
End Function

Function MAXIVAL(V1, V2)
   'Ataja valor nulo
   If IsNumeric(V1) = True Or IsNumeric(V2) = True Then
      V1 = XVALO(V1)
      V2 = XVALO(V2)
    Else
      V1 = SAFETEXT(V1)
      V2 = SAFETEXT(V2)
   End If
   'Comparacion de valores
   If V1 > V2 Then
      MAXIVAL = V1
    Else
      MAXIVAL = V2
   End If
   '
End Function
Function MINIVAL(V1, V2)
   'Ataja valor nulo
   If IsNumeric(V1) = True Or IsNumeric(V2) = True Then
      V1 = XVALO(V1)
      V2 = XVALO(V2)
    Else
      V1 = SAFETEXT(V1)
      V2 = SAFETEXT(V2)
   End If
   'Comparacion de valores
   If V1 < V2 Then
      MINIVAL = V1
    Else
      MINIVAL = V2
   End If
   '
End Function

Function RESUCALC#(FORMUTRAB$)
      '
      If IsDate(Left(FORMUTRAB$, 8)) Then
        If Mid$(FORMUTRAB$, 3, 1) = "/" And Mid$(FORMUTRAB$, 6, 1) = "/" Then
          FORMUTRAB$ = REPLACAR(FORMUTRAB$, "/", "|")
        End If
      End If
      '
      FORTRABA$ = REPLACAR$(FORMUTRAB$, ",", ".")
      PPXXX% = InStr(FORTRABA$, "--"): If PPXXX% > 0 Then FORTRABA$ = Left$(FORTRABA$, PPXXX% - 1) + "+" + Mid$(FORTRABA$, PPXXX% + 2)
      PPXXX% = InStr(FORTRABA$, "- -"): If PPXXX% > 0 Then FORTRABA$ = Left$(FORTRABA$, PPXXX% - 1) + "+" + Mid$(FORTRABA$, PPXXX% + 3)
      PPXXX% = InStr(FORTRABA$, "-  -"): If PPXXX% > 0 Then FORTRABA$ = Left$(FORTRABA$, PPXXX% - 1) + "+" + Mid$(FORTRABA$, PPXXX% + 4)
      '
16220 POPE% = InStr(FORTRABA$, "^"): If POPE% > 1 Then OPE% = 5: GoTo 16280
16222 POPE% = InStr(FORTRABA$, "x"): If POPE% > 1 Then OPE% = 1: GoTo 16280
16225 POPE% = InStr(FORTRABA$, "X"): If POPE% > 1 Then OPE% = 1: GoTo 16280
16230 POPE% = InStr(FORTRABA$, "*"): If POPE% > 1 Then OPE% = 1: GoTo 16280
16235 POPE% = InStr(FORTRABA$, "/"): If POPE% > 1 Then OPE% = 2: GoTo 16280
16245 POPEMAS% = InStr(FORTRABA$, "+")
      POPEMEN% = InStr(FORTRABA$, "-")
      POPE% = POPEMAS%: If POPEMEN% > 1 And POPEMEN% < POPEMAS% Then POPE% = POPEMEN%
      'CUANDO FORTRABA$ TIENE SOLO EL SIMBOLO MENOS, A LA VARIABLE POPE% NO SE LE ASIGNA LA POSICION
      If POPE% < 1 And POPEMEN% > 0 Then POPE% = POPEMEN%
      If POPE% > 1 Then
         If Mid$(FORTRABA$, POPE%, 1) = "+" Then OPE% = 3: GoTo 16280
         If Mid$(FORTRABA$, POPE%, 1) = "-" Then OPE% = 4: GoTo 16280
      End If
16255 'POPE% = InStr(FORTRABA$, "-"): If POPE% > 1 Then OPE% = 4: GoTo 16280
      '
16265 ' NO QUEDAN OPERACIONES ARITMETICAS POR CONSIDERAR
16267 If Abs(RI2#) < 0.005 Then
        If CALFIN% = 0 Then
           IZQUIE$ = "":  DERIE$ = "":           VALO1# = 0
           ARGU2$ = FORTRABA$
           OPE% = 3:           CALFIN% = 1
           GoTo 16460
         End If
      End If
      SIG1% = 1: If RI2# > 0 And Left$(FORTRABA$, 1) = "-" Then SIG1% = (-1)
      RESUCALC# = RI2# * SIG1%
      Exit Function
      '
16280 'If Left$(FORTRABA$, 1) = "-" Then FORTRABA$ = "0" + FORTRABA$
      PPX1% = InStr(FORTRABA$, "-(-"): If PPX1% > 0 Then FORTRABA$ = Left$(FORTRABA$, PPX1% - 1) + "+(" + Mid$(FORTRABA$, PPX1% + 3)
      PPX1% = InStr(FORTRABA$, "--"): If PPX1% > 0 Then FORTRABA$ = Left$(FORTRABA$, PPX1% - 1) + "+" + Mid$(FORTRABA$, PPX1% + 2)
      '
      ARG1$ = "": ARG2$ = ""
16285 For pp% = POPE% - 1 To 1 Step -1
16295    CR$ = Mid$(FORTRABA$, pp%, 1)
16300    If CR$ = "^" Or CR$ = "x" Or CR$ = "X" Or CR$ = "*" Or CR$ = "/" Or CR$ = "+" Or CR$ = "-" Then GoTo 16320
16305 Next pp%
16310 pp% = 0
      '
16320 ARG1$ = Mid$(FORTRABA$, pp% + 1, POPE% - pp% - 1)
      IZQUIE$ = ""
      If pp% > 0 Then IZQUIE$ = Left$(FORTRABA$, pp%)
      If IZQUIE$ = "-" Then ARG1$ = IZQUIE$ + ARG1$: IZQUIE$ = ""
      '
16330 For pp% = POPE% + 1 To Len(FORTRABA$)
16335    CR$ = Mid$(FORTRABA$, pp%, 1)
16340    If CR$ = "^" Or CR$ = "x" Or CR$ = "X" Or CR$ = "*" Or CR$ = "/" Or CR$ = "+" Or CR$ = "-" Then GoTo 16360
16345 Next pp%
16350 pp% = 1 + Len(FORTRABA$)
      '
16360 ARG2$ = Mid$(FORTRABA$, POPE% + 1, pp% - POPE% - 1): DERIE$ = Mid$(FORTRABA$, pp%)
16365 '
16370 ARGU1$ = "": For pp% = 1 To Len(ARG1$): CR1% = Asc(Mid$(ARG1$, pp%)): If (CR1% > 47 Or CR1% = 46) Then ARGU1$ = ARGU1$ + Chr$(CR1%)
16375 Next pp%
      If Left$(ARG1$, 1) = "-" Then ARGU1$ = "-" + ARGU1$
16380 '
16385 ARGU2$ = "": For pp% = 1 To Len(ARG2$): CR1% = Asc(Mid$(ARG2$, pp%)): If (CR1% > 47 Or CR1% = 46) Then ARGU2$ = ARGU2$ + Chr$(CR1%)
16390 Next pp%
16395 '
16400 ESHORA% = 0: ESFECHA% = 0
      VALO1# = Val(ARGU1$)
         If InStr(ARGU1$, ":") > 0 Then VALO1# = HORANUM(ARGU1$): ESHORA% = 1
         If InStr(ARGU1$, "|") > 0 Then
            ARGU1$ = REPLACAR(ARGU1$, "|", "/")
            VALO1# = FECHANUM(ARGU1$): ESFECHA% = 1
         End If
16460 VALO2# = Val(ARGU2$)
         If InStr(ARGU2$, ":") > 0 Then VALO2# = HORANUM(ARGU2$): ESHORA% = 1
         If InStr(ARGU2$, "|") > 0 Then
            ARGU2$ = REPLACAR(ARGU2$, "|", "/")
            VALO2# = FECHANUM(ARGU2$): ESFECHA% = 1
         End If
16515 '
16520 RI2# = 0
16525 If OPE% = 1 Then RI2# = VALO1# * VALO2#: GoTo 16550
16530 If OPE% = 2 Then
         ' If Abs(VALO2#) >= 0.005 Then
         On Error Resume Next
         RI2# = VALO1# / VALO2#
         If Err Then RI2# = 0
         On Error GoTo 0
         GoTo 16550
         ' End If
      End If
16535 If OPE% = 3 Then RI2# = VALO1# + VALO2#: GoTo 16550
16540 If OPE% = 4 Then
          RI2# = VALO1# - VALO2#
          If ESHORA% = 1 Then
             RII = 3600 * DIFHORA(VALO2#, VALO1#)
             RIJ$ = HORMINSE$(RII)
             RI2# = HORANUM(RIJ$)
          End If
          If ESFECHA% = 1 Then
             RI2# = DIENTRE%(VALO2#, VALO1#)
             If VALO2# > VALO1# Then
                RI2# = (-1) * DIENTRE(VALO1#, VALO2#)
             End If
          End If
          GoTo 16550
      End If
16542 If OPE% = 5 Then RI2# = VALO1# ^ VALO2#: GoTo 16550
16545 '
16550 RI3$ = TRIM$(FORMATNUM$(RI2#, "F20.9"))
      If RI3$ = "" Then RI3$ = "0"
16555 FORTRABA$ = IZQUIE$ + RI3$ + DERIE$
16560 GoTo 16220
      '
End Function

Function PORCENT_ENTRE_2_VALORES(VALORMAYOR, VALORMENOR)
   '
   PORCENT_ENTRE_2_VALORES = 1
   If VALORMENOR = 0 Then Exit Function
   DIFER = VALORMAYOR - VALORMENOR
   PORCENT = DIFER / VALORMENOR
   PORCENT = PORCENT * 100
   PORCENT_ENTRE_2_VALORES = PORCENT
   '
End Function

Function ToleranciaEnPorcentaje(ValorReferencia As Double, ValorResultado As Double) As Single
   '
   If ValorReferencia < 0.00005 Then                           ' SI EL PRECIO DE LISTA ES CERO NO HACE NADA
        If ValorResultado > 0.00005 Then
            PORCENTAJE = 100
        Else
            PORCENTAJE = 0
        End If
        GoTo 99
   End If
   Diferencia = ValorReferencia - ValorResultado
   PORCENTAJE = Diferencia / ValorReferencia
   PORCENTAJE = PORCENTAJE * 100
99 ToleranciaEnPorcentaje = PORCENTAJE
   '
End Function

Function ImporteConDescuento#(IMPORTEBRUTO#, DESC$)
   
   'DESC$ = EL CASILLERO DE DESCUENTO EN FORMATO TEXTO
   'IMPORTEBRUTO# = EL CASILLERO DEL IMPORTE CON EL CUAL SE VA A CALCULAR EL DESCUENTO
   
   'DESC$ = TRIM$(Pordescu.TEXT)
   PODES = COEFDESCU(DESC$)
   TTT1# = XVALO(IMPORTEBRUTO# * PODES / 100)
   IMPONET# = XVALO(IMPORTEBRUTO#) - TTT1#
   ImporteConDescuento# = IMPONET#
End Function
Function PrecioOriginalPreDescuento(ValorConDesc#, PORCENDESCU)
    'ESTA FUNCION DEVUELVE el precio original después del descuento
    'EJEMPLO
    'ValorConDesc# = 80
    'PORCENDESCU = 20 %
    'LA FUNCION DEVUELVE PrecioOriginalPreDescuento = 100
    'ValorConDesc# es el pecio ya con el descuento generado
    'PORCENDESCU  es el porcentaje de descuento que se le hizo para llegar al ValorConDesc#
    If Abs(PORCENDESCU) > 0 Then PrecioOriginalPreDescuento = ValorConDesc#
    RESULTADO1 = 1 - (PORCENDESCU / 100)
    PrecioOriginalPreDescuento = ValorConDesc# / RESULTADO1
End Function
 Function VECTOR_INICIALIZADO(VECTOR$()) As Boolean
   On Error GoTo Err_Inicializar
   If UBound(VECTOR$) > 0 Then
     VECTOR_INICIALIZADO = True
    'daria el error pero va a la etiqueta Err_Inicializar
   End If
   On Error GoTo 0
   Exit Function
   
Err_Inicializar:
  If Err.Number > 0 Then
    VECTOR_INICIALIZADO = False
    On Error GoTo 0
  End If
End Function

Function CALCULA_IMPO_REPA#(TOTREPA#, DTO1, DTO2)
    DESC1$ = SAFETEXT$(DTO1)
    DESC2$ = SAFETEXT$(DTO2)
   '
   TTT1 = TOTREPA# * COEFDESCU(DESC1$) / 100
   '
   'ACTUALIZO EL SUBTOTAL PRIMERO
     '
   TTT2 = (TOTREPA# - TTT1) * COEFDESCU(DESC2$) / 100 'CALCULO EL SEGUNDO DESCUENTO CON EL PRIMERO APLIDADO
   '********************************************************************************
   'NOTA: TTT1 ES UN IMPORTE DE DESCUENTO Y TTT2 ES UN IMPORTE TRATADO COMO RECARGO.
   '********************************************************************************
   CALCULA_IMPO_REPA# = TOTREPA# - TTT1 + TTT2

End Function

Function ResultadoValorSegúnPorcentaje(PORCENTAJE, IMPORTE)
   ' DEVUELVE EL PORCENTAJE DEL VALOR IMPORTE SEGUN PORCENTAJE
   'EJ. SI SE PASA % 20 E IMPORTE 200 DEVULEVE = 40
   ResultadoValorSegúnPorcentaje = IMPORTE
   If PORCENTAJE > 0 And IMPORTE > 0 Then
       ResultadoValorSegúnPorcentaje = (IMPORTE * PORCENTAJE) / 100
   End If
End Function

