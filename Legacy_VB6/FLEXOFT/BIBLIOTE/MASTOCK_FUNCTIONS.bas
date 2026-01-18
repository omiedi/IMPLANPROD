Attribute VB_Name = "MASTOCK_FUNCTIONS"
Public NUSU%, COSU%(64), SUCU$(64), ADEPO$(64), DISTO%(64), CATEO#(64), CAFIN#(64)
'
Sub APERBASDAT(NBASE$)
   TIBASE$ = "DAT"
End Sub
'
Function NUINV%()
'
X$ = REGLEIDO$("INVERT", 1&)
NUINV% = CVI(Mid$(X$, 17, 2))
End Function
' ***************************************************************************

Function NUMAS%()
   '
   Static TIANT, NMSX%
   If Timer < TIANT + 3 Then
       NUMAS% = NMSX%
     Else
       X$ = REGLEIDO$("MASTER", 1&)
       NUMAS% = CVI(Left$(X$, 2))
       NMSX% = NUMAS%
       TIANT = Timer
   End If
   '
End Function
' ***************************************************************************
'
Function PRICODACT%()
   
   PCC% = 0
   If ULTREG&("ECOMAST") > 0 Then
      XX$ = Left$(REGLEIDO$("ECOMAST", 1), 16)
      PCC% = Abs(CODINT%(XX$))
   End If
   PRICODACT% = PCC%
   '
End Function
'
Function ULTCODACT%()
   '
   UCC% = 0
   If ULTREG&("ECOMAST") > 0 Then
      AAA& = ULTREG&("ECOMAST")
10    XX$ = Left$(REGLEIDO$("ECOMAST", AAA&), 16)
      If TRIM$(XX$) = "" Then
        If AAA& > 1 Then
          AAA& = AAA& - 1
          GoTo 10
        End If
      End If
      UCC% = Abs(CODINT%(XX$))
      If TRIM$(CODEXT$(UCC%)) = "" Then
        If AAA& > 1 Then
          AAA& = AAA& - 1
          GoTo 10
        End If
      End If
   End If
   ULTCODACT% = UCC%
   '
End Function

Function CODINT%(COD$)
'
CODINT% = 0
NV& = NUINV%: LI& = 1: LS& = NV& + 2
'
APROX% = 0
If Left$(COD$, 1) = "~" Then
    COD$ = Mid$(COD$, 2, 16)
    APROX% = 1
End If
'
ABRE$ = AJUSTI$(COD$, 16)
LOCON% = Len(ABRE$)
'
1510 E& = Int((LS& - LI&) / 2): If E& = 0 Then GoTo 1580
     L& = LI& + E&: If L& < 2 Or L& > NV& + 1 Then GoTo 1580
     I$ = REGLEIDO$("INVERT", L&): I1$ = Left$(I$, LOCON%)
                                   If I1$ < ABRE$ Then LI& = L&: GoTo 1510
                                   If I1$ > ABRE$ Then LS& = L&: GoTo 1510
     CODINT% = CVI(Mid$(I$, 17, 2))
     '
1580 If APROX% = 1 Then
1581     If Mid$(I$, 16, 1) = Chr$(96) Or Left$(I$, 16) < ABRE$ Then
             If L& < NUINV% + 1 Then
                 L& = L& + 1
                 I$ = REGLEIDO$("INVERT", L&)
                 GoTo 1581
             End If
         End If
         '
         If Len(I$) >= 18 Then
             CODINT% = CVI(Mid$(I$, 17, 2))
         End If
     End If
     '
     If CODINT% = 0 Then
         ' codigo de baja devuelve CODINT% negativo
         ABRE$ = AJUSTI$(COD$, 15) + Chr$(96)
1590     E& = Int((LS& - LI&) / 2): If E& = 0 Then GoTo 1599
         L& = LI& + E&: If L& < 2 Or L& > NV& + 1 Then GoTo 1599
         I$ = REGLEIDO$("INVERT", L&): I1$ = Left$(I$, LOCON%)
                                       If I1$ < ABRE$ Then LI& = L&: GoTo 1590
                                       If I1$ > ABRE$ Then LS& = L&: GoTo 1590
         CODINT% = (-1) * CVI(Mid$(I$, 17, 2))
     End If
     '
1599 '
End Function
' ***************************************************************************

Function REGINV&(COD$)
     '
     REGINV& = 0
     NV& = NUINV%: LI& = 1
                   LS& = NV& + 2
                   '
     ABRE$ = AJUSTI$(COD$, 15)
     LOCON% = Len(ABRE$)
     '
1610 E& = Int((LS& - LI&) / 2)
     If E& = 0 Then Exit Function
     '
     L& = LI& + E&
     If L& < 2 Or L& > NV& + 1 Then Exit Function
     '
     I$ = Left$(REGLEIDO$("INVERT", L&), 15)
     If I$ < ABRE$ Then LI& = L&: GoTo 1610
     If I$ > ABRE$ Then LS& = L&: GoTo 1610
     '
     REGINV& = L&
     '
End Function
' ***************************************************************************

Function CODEXT$(CI%)
'
CODEXT$ = ""
If CI% > 0 Then
   If CI% <= NUMAS% Then
       REG& = CI% + 1
       XXX$ = REGLEIDO$("MASTER", REG&)
       CODEXT$ = Mid$(XXX$, 77, 16)
   End If
End If
End Function
'
' ***************************************************************************
'
Function DESCRIT$(CI%)
'
DESCRIT$ = ""
If CI% > 0 Then
   If CI% <= NUMAS% Then
       REG& = CI% + 1
       XXX$ = REGLEIDO$("MASTER", REG&)
       DESCRIT$ = TRIM$(Mid$(XXX$, 5, 64)) + " (" + UNIMED$(CI%) + ")"
   End If
End If
End Function
'

' ***************************************************************************
'
Function DESCRIT0$(CI%)
'
DESCRIT0$ = ""
If CI% > 0 Then
   If CI% <= NUMAS% Then
       REG& = CI% + 1
       XXX$ = REGLEIDO$("MASTER", REG&)
       DESCRIT0$ = TRIM$(Mid$(XXX$, 5, 64))
   End If
End If
End Function
'

' ***************************************************************************
'
Function UNIMED$(CI%)
'
UNIMED$ = ""
If CI% > 0 Then
   If CI% <= NUMAS% Then
       REG& = CI% + 1
       XXX$ = REGLEIDO$("MASTER", REG&)
       UNIMED$ = Mid$(XXX$, 109, 2)
       If TRIM$(UNIMED$) = "" Then
           Call REPLA(XXX$, "U.", 109, 2)
           Call GRAREG("MASTER", XXX$, REG&)
           UNIMED$ = "U."
       End If
   End If
End If
End Function
'
Function UNICOM$(CI%)
   '
   UNICO$ = ""
   RFF$ = RECFILT$("OCOMDETA", 5, MKI$(CI%))
   If Len(RFF$) >= 4 Then
     U& = Len(RFF$) - 3
     REGO& = CVS(Mid$(RFF$, U&, 4))
     YX$ = REGLEIDO$("OCOMDETA", REGO&)
     UNICO$ = TRIM$(Mid$(YX$, 59, 4))
   End If
   '
   If ECOARCH$("UNIMED", UNICO$) = "" Then UNICO$ = ""
   If UNICO$ = "" Then UNICO$ = UNIMED$(CI%)
   '
   UNICOM$ = UNICO$
   '
End Function
' ***************************************************************************
'
Function REVICOD$(CI%)
'
If CI% = CIANT% Then GoTo 10
'
REVICOD$ = ""
If CI% > 0 Then
   If CI% <= NUMAS% Then
       REG& = CI% + 1: CIANT% = CI%
       XXX$ = REGLEIDO$("MASTER", REG&)
10     REVICOD$ = TRIM$(Mid$(XXX$, 117, 4))
   End If
End If
End Function
'
' ***************************************************************************
'
Function FEREVI%(CI%)
'
If CI% = CIANT% Then GoTo 10
'
FEREVI% = 0
If CI% > 0 Then
   If CI% <= NUMAS% Then
       REG& = CI% + 1: CIANT% = CI%
       XXX$ = REGLEIDO$("MASTER", REG&)
10     FEREVI% = CVI(Mid$(XXX$, 121, 2))
   End If
End If
End Function
'
' ***************************************************************************
'
Function STATUIT%(CI%)
'
STATUIT% = 1
If CI% > 0 Then
  If CI% <= NUMAS% Then
    REG& = CI% + 1
    XXX$ = REGLEIDO$("MASTER", REG&)
    If Mid$(XXX$, 92, 1) = "`" Then
      STATUIT% = (-1)
    End If
  End If
End If
End Function
'
' ***************************************************************************
'
Function COLOR$(CI%)
'
COLOR$ = ""
If CI% > 0 Then
   If CI% <= NUMAS% Then
       REG& = CI% + 1
       XXX$ = REGLEIDO$("MASTER", REG&)
       CC% = Asc(Mid$(XXX$, 96, 1))
       COLOR$ = LCOLOR$(CC%)
   End If
End If
End Function
'
' ***************************************************************************
'
Function TALLE$(CI%)
'
TALLE$ = ""
If CI% > 0 Then
   If CI% <= NUMAS% Then
       REG& = CI% + 1
       XXX$ = REGLEIDO$("MASTER", REG&)
       CC% = Asc(Mid$(XXX$, 97, 1))
       TALLE$ = LTALLE$(CC%)
   End If
End If
End Function
'
' ***************************************************************************
'
Function TIPOHIL$(CI%)
'
TIPOHIL$ = TALLE$(CI%)
End Function
'
' ***************************************************************************
'
Function PESUNI(CI%)
PESUNI = 0
If CI% > 0 Then
   If CI% <= NUMAS% Then
       REG& = CI% + 1
       XXX$ = REGLEIDO$("MASTER", REG&)
       PP = CVS(Mid$(XXX$, 99, 4))
       If PP > 0 Then
          PESUNI = PP
       End If
   End If
End If
End Function
'
' ***************************************************************************
'
Function TITULO$(CI%)
TITULO$ = ""
If CI% > 0 Then
   If CI% <= NUMAS% Then
       REG& = CI% + 1
       XXX$ = REGLEIDO$("MASTER", REG&)
       CP = CVS(Mid$(XXX$, 99, 4))
       If CP < 0 Then
          LCLR$ = TRIM$(CSTRING$(MKS$(Abs(CP)), 3, 8, 2, 2))
          TITULO$ = REPLACAR$(LCLR$, ".", "/")
       End If
   End If
End If
End Function
'
' ***************************************************************************
'
Function COSUNI(CI%)
'
COSUNI = 0
If CI% > 0 Then
   If CI% <= NUMAS% Then
       REG& = CI% + 1
       XXX$ = REGLEIDO$("MASTER", REG&)
       COSUNI = CVS(Mid$(XXX$, 111, 4))
   End If
End If
End Function
'
' ***************************************************************************
'
Function MONECOSTO%(CI%)
'
Static FICATO#, CAMONED%
'
If FileDateTime(LUDAT$ + "TAMONED.DAT") <> FICATO# Then
  CAMONED% = ULTREG&("TAMONED")
End If
'
If CI% = CIANT% Then GoTo 10
'
MONECOSTO% = 1
If CI% > 0 Then
  If CI% <= NUMAS% Then
    REG& = CI% + 1: CIANT% = CI%
    XXX$ = REGLEIDO$("MASTER", REG&)
10  MONEI% = Asc(Mid$(XXX$, 95, 1))
    If MONEI% > 0 And MONEI% <= CAMONED% Then
      MONECOSTO% = MONEI%
    End If
  End If
End If
'
End Function
'
' ***************************************************************************
'
Function PREBASE(CI%)
'
PREBASE = COSUNI(CI%)
End Function
'
' ***************************************************************************
'
Function FECOSTO%(CI%)
'
FECOSTO% = 0
If CI% > 0 Then
   If CI% <= NUMAS% Then
       REG& = CI% + 1
       XXX$ = REGLEIDO$("MASTER", REG&)
       FECOSTO% = CVI(Mid$(XXX$, 115, 2))
   End If
End If
End Function
'
' ***************************************************************************
'
Function PORCEIVA(CILOCAL%)
'
Static CTX%, PRIVA(15), REMA As String * 128
If CTX% < 1 Then
   CTX% = 1
   For U& = 1 To 2
      X$ = REGLEIDO$("GRUPIVA", U&)
      PRIVA(U& - 1) = CVS(Mid$(X$, 13, 4))
   Next U&
   Call CIERRARCH("GRUPIVA")
End If
'
GRIVA% = 0
If CILOCAL% > 0 Then
   If CILOCAL% <= NUMAS% Then
       NXX% = FILENBR%("MASTER")
       REG& = CILOCAL% + 1
       Get #NXX%, REG&, REMA$
       GRIVA% = Asc(Mid$(REMA$, 3, 2))
   End If
End If
If GRIVA% < 0 Or GRIVA% > 15 Then GRIVA% = 0
'
PORCEIVA = PRIVA(GRIVA%)
'
End Function
'
' ***************************************************************************
'
Function STOMIN(CI%)
'
STOMIN = 0
If CI% > 0 Then
   If CI% <= NUMAS% Then
       REG& = CI% + 1
       XXX$ = REGLEIDO$("MASTER", REG&)
       STOMI = CVI(Mid$(XXX$, 105, 2))
       If STOMI < 0 Then STOMI = 100 * Abs(STOMI)
       STOMIN = STOMI
   End If
End If
'
End Function
'
' ***************************************************************************
'
Function LOTESTAN(CI%)
'
If CI% > 0 Then
   If CI% <= NUMAS% Then
       REG& = CI% + 1
       XXX$ = REGLEIDO$("MASTER", REG&)
       LT = CVI(Mid$(XXX$, 107, 2))
       If LT < 0 Then LT = 100 * Abs(LT)
   End If
End If
'
If LT < 1 Then LT = 1
LOTESTAN = LT
'
End Function
'
' ***************************************************************************
'
Function LOREPOMA(CI%)
'
If CI% > 0 Then
   If CI% <= NUMAS% Then
       REG& = CI% + 1
       XXX$ = REGLEIDO$("MASTER", REG&)
       LR = CVI(Mid$(XXX$, 103, 2))
       If LR < 0 Then LR = 100 * Abs(LR)
       LT = CVI(Mid$(XXX$, 107, 2))
       If LT < 0 Then LT = 100 * Abs(LT)
   End If
End If
'
If LR < 1 Then LR = 1
If LR < LT Then LR = LT
LOREPOMA = LR
'
End Function
'
' ***************************************************************************
'
Function TIMATE%(CI%)
'
If CI% > 0 Then
   If CI% <= NUMAS% Then
       REG& = CI% + 1
       XXX$ = REGLEIDO$("MASTER", REG&)
       TMT% = Asc(Mid$(XXX$, 93, 1))
   End If
End If
'
TIMATE% = TMT%
'
End Function
'
' ***************************************************************************
'
Function GRUCOVE%(CI%)
'
GCV% = 0
If CI% > 0 Then
   If CI% <= NUMAS% Then
       REG& = CI% + 1
       XXX$ = REGLEIDO$("MASTER", REG&)
       GCV% = CVI(Mid$(XXX$, 125, 2))
   End If
End If
'
GRUCOVE% = GCV%
'
End Function
'
' ***************************************************************************
'
Function DEPOPRE%(CI%)
'
DPP% = 0
TMT% = TIMATE%(CI%)
For U& = 1 To ULTREG&("TABTIMA")
   X$ = REGLEIDO$("TABTIMA", U&)
   If Asc(Left$(X$, 1)) = TMT% Then
      DPP% = Asc(Mid$(X$, 17, 1))
      GoTo 10
   End If
Next U&
'
10 If DPP% > 0 Then
     If ECOARCH$("ECODEP", Chr$(DPP%)) = "" Then DPP% = 0
     If DPP% = 99 Then DPP% = 0
End If
'
DEPOPRE% = DPP%
'
End Function
'
' ***************************************************************************
'
Function DEPOEXPE%()
    '
    Static CTXX%, DPXXP%
    If CTXX% = 0 Then
      CTXX% = 1
      DPXXP% = 0
      'If ULTREG&("SUCREMOT") > 0 Then
        DPXX% = Val(CONTROL$("", "DEPOEXPE"))
        If DPXX% > 0 Then
          If DPXX% < 99 Then
            If ECOARCH$("TADEPOSI", Chr$(DPXX%)) <> "" Then
              DPXXP% = DPXX%
            End If
          End If
        End If
      'End If
    End If
    '
    DEPOEXPE% = DPXXP%
    '
End Function
'
' ***************************************************************************
'
Function DESCOMPO%(CI%)
DESCOMPO% = 0
If CI% > 0 Then
   If CI% <= NUMAS% Then
       REG& = CI% + 1
       XXX$ = REGLEIDO$("MASTER", REG&)
       DESCOMPO% = Asc(Mid$(XXX$, 98, 1))
       If DESCOMPO% > 1 Then
          DESCOMPO% = 0
       End If
   End If
End If
End Function
'
' ***************************************************************************
'
Function STOCKACT(CI%)
   '
   If NUSU% < 1 Then Call CARDEPOS
   '
   ST = 0
   If CI% > 0 Then
     If CI% <= NUMAS% Then
       REG& = CI% + 1
       XXX$ = REGLEIDO$("SALDEPOS", REG&)
       For U% = 1 To 4 * NUSU% Step 4
         ST = ST + (1 + DISTO%((U% + 3) / 4)) * CVS(Mid$(XXX$, U%, 4))
       Next U%
     End If
   End If
   '
   STOCKACT = ST
   '
End Function
'
Function STODEPOS(CI%, DEPO%)
'
Static NUSU%, COSU%(64)
ST = 0
'
If NUSU% = 0 Then
   For U& = 1 To ULTREG&("SUCURSAL")
      If U& <= 62 Then
         SS$ = REGLEIDO$("SUCURSAL", U&)
         If Asc(Left$(SS$, 1)) > 32 Then
            NUSU% = U&
            COSU%(U&) = Asc(Mid$(SS$, 64, 1))
         End If
      End If
   Next U&
End If
'
If DEPO% < 1 Then
   ST = STOCKACT(CI%)
   GoTo 99
End If
'
REG& = CI% + 1
If REG& > 0 Then
   If REG& <= NUMAS% + 1 Then
      SALCA$ = REGLEIDO$("SALSTOCK", REG&) + Mid$(REGLEIDO$("SALDEPOS", REG&), 41)
      For U& = 1 To NUSU%
         If COSU%(U&) = DEPO% Then
            ST = CVS(Mid$(SALCA$, 4 * U& - 3, 4))
            GoTo 99
         End If
      Next U&
   End If
End If
'
99 STODEPOS = ST
   '
End Function
'
Function STOCKASI(CI%)
    '
10  ASII = 0
    UREOD& = ULTREG&("MAVENDID")
    RGXX$ = RECFILT$("MAVENDID", 20, MKI$(CI%))
    '
    For U& = 1 To Len(RGXX$) Step 4
      REG& = CVS(Mid$(RGXX$, U&, 4))
      If REG& > 0 Then
        If REG& <= UREOD& Then
          X$ = REGLEIDO$("MAVENDID", REG&)
          If Asc(Mid$(X$, 63, 1)) <= 1 Then  ' STATUS 1
            CASIG = CVS(Mid$(X$, 177, 4)) - CVS(Mid$(X$, 225, 4))
            If CASIG < 0 Then CASIG = 0
            ASII = ASII + CASIG
          End If
        End If
      End If
    Next U&
    '
    STOCKASI = ASII
    '
End Function
'
Function STOCKPED(CI%)
    '
    Static ABUSQUE$
    If ABUSQUE$ = "" Then
      ABUSQUE$ = "PEDDETP"
    End If
    '
10  ASII = 0
    RGXX$ = RECFILT$(ABUSQUE$, 12, MKI$(CI%))
    '
    For I& = 1 To Len(RGXX$) Step 4
       REG& = CVS(Mid$(RGXX$, I&, 4))
       X$ = REGLEIDO$(ABUSQUE$, REG&)
       '
       If ABUSQUE$ = "PEDDETP" Then
         REGPEDA& = CVS(Mid$(X$, 109, 4))
         If Left$(REGLEIDO$("PEDDETA", REGPEDA&), 108) <> Left$(X$, 108) Then
           ABUSQUE$ = "PEDDETA"
           GoTo 10
         End If
       End If
       '
       If Asc(Mid$(X$, 27, 1)) <= 2 Then
         CASIG = CVS(Mid$(X$, 79, 4)) - CVS(Mid$(X$, 95, 4))
         If CASIG < 0 Then CASIG = 0
         '
         ASII = ASII + CASIG
       End If
    Next I&
    '
    STOCKPED = ASII
    '
End Function
'
Function STOCKRES(CI%)
    '
10  ASII = 0
    RGXX$ = RECFILT$("RESERVA", 1, MKI$(CI%))
    '
20  For I& = 1 To Len(RGXX$) Step 4
        REG& = CVS(Mid$(RGXX$, I&, 4))
        X$ = REGLEIDO$("RESERVA", REG&)
        CASIG = CVD(Mid$(X$, 19, 8)) - CVD(Mid$(X$, 27, 8))
        ASII = ASII + CASIG
    Next I&
    '
    STOCKRES = ASII
    '
End Function
'
Function STORESOF(CI%, NORFAI&)
    '
10  ASII = 0
    RGXX$ = RECFILT$("RESERVA", 1, MKI$(CI%))
    '
    For I& = 1 To Len(RGXX$) Step 4
        REG& = CVS(Mid$(RGXX$, I&, 4))
        X$ = REGLEIDO$("RESERVA", REG&)
        If Val(Mid$(X$, 6, 6)) = NORFAI& Then
           CASIG = CVD(Mid$(X$, 19, 8)) - CVD(Mid$(X$, 27, 8))
           ASII = ASII + CASIG
        End If
        '
    Next I&
    '
    STORESOF = ASII
    '
End Function
'
Function STOCKDIS(CIKI%)
    '
'AAA = STOCKACT(CIKI%)
'BBB = STOCKPED(CIKI%)
'CCC = STOCKRES(CIKI%)
    STOCKDIS = STOCKACT(CIKI%) - STOCKPED(CIKI%) - STOCKRES(CIKI%)
    '
End Function
'

Function STOCKDISNEVEN(CI%)
    '
    STOCKDISNEVEN = STODEPOS(CI%, DEPOEXPE%) - STOCKASI(CI%) - STOCKRES(CI%) - PEDNOASI(CI%)
    '
End Function
'

Function STOCKDISNE(CI%)
    '
    STOCKDISNE = STOCKDIS(CI%) - PEDNOASI(CI%)
    '
End Function
'
Function STORFAB(CI%)
   '
   STORF = 0
   RFF$ = RECFILT$("ORDEFABR", 5, MKI$(CI%))
   For UK& = 1 To Len(RFF$) Step 4
      U& = CVS(Mid$(RFF$, UK&, 4))
      X$ = REGLEIDO$("ORDEFABR", U&)
      STAT% = Asc(Mid$(X$, 108, 1))
      If STAT% <= 2 Then
        CACOM = CVD(Mid$(X$, 89, 8))
        CAREC = CVD(Mid$(X$, 119, 8))
        CAPEN = CACOM - CAREC
        If CAPEN > 0 Then
          STORF = STORF + CAPEN
        End If
      End If
   Next UK&
   '
   STORFAB = STORF
   '
End Function
'
Function UBISTO$(CI%)
    '
20  UBX$ = ""
    RGXX$ = RECFILT$("UBISTOCK", 1, MKI$(CI%))
    For U& = 1 To Len(RGXX$) Step 4
        REG& = CVS(Mid$(RGXX$, U&, 4))
        X$ = REGLEIDO$("UBISTOCK", REG&)
        If CVI(Left$(X$, 2)) <> CI% Then
           X$ = REGLEIDO$("UBISTOCK", 1)
           Call GRAREG("UBISTOCK", X$, 1)
           Call CIERRARCH("UBISTOCK")
           GoTo 20
        End If
        '
        If TRIM$(Mid$(X$, 3, 24)) <> "" Then
           If UBX$ <> "" Then UBX$ = UBX$ + ", "
           UBX$ = UBX$ + TRIM$(Mid$(X$, 3, 24))
        End If
    Next U&
    '
    UBISTO$ = UBX$
    '
End Function
'
Function SEPRODUC(CI%)
    '
20  SPROD = 0
    RGXX$ = RECFILT$("ORDEFABR", 5, MKI$(CI%))
    For U& = 1 To Len(RGXX$) Step 4
        REG& = CVS(Mid$(RGXX$, U&, 4))
        X$ = REGLEIDO$("ORDEFABR", REG&)
        If CVI(Mid$(X$, 83, 2)) <> CI% Then
           X$ = REGLEIDO$("ORDEFABR", 1)
           Call GRAREG("ORDEFABR", X$, 1)
           Call CIERRARCH("ORDEFABR")
           GoTo 20
        End If
        '
        SPROD1 = CVS(Mid$(X$, 113, 4))
        If SPROD1 > 0 Then SPROD = SPROD1
    Next U&
    '
    SEPRODUC = SPROD
    '
End Function
'
Function PEDNOASI(CI%)
    '
10  ASII = 0
    RGXX$ = RECFILT$("ODESEMI", 6, MKI$(CI%))
    '
    For I& = 1 To Len(RGXX$) Step 4
        REG& = CVS(Mid$(RGXX$, I&, 4))
        X$ = REGLEIDO$("ODESEMI", REG&)
        CIP% = CVI(Mid$(X$, 11, 2))
        If CIP% <> CI% Then        ' fuerza reindexacion
            MPANTE = Screen.MousePointer
            Screen.MousePointer = 11
            X$ = REGLEIDO$("ODESEMI", 1)
            Call GRAREG("ODESEMI", X$, 1)
            Call CIERRARCH("ODESEMI")
            Screen.MousePointer = MPANTE
            GoTo 10
        End If
        '
        ASII = ASII + CVS(Mid$(X$, 17, 4))
        '
    Next I&
    '
    PEDNOASI = ASII
    '
End Function
'

Function CCOLOR%(TCX$)
    '
    TC$ = TRIM$(UCase$(TCX$))
    CC% = 0
    '
    If TC$ <> "" Then
       For U& = 1 To ULTREG&("TACOLOR")
           XX$ = TRIM$(Left$(REGLEIDO$("TACOLOR", U&), 31))
           If XX$ = TC$ Then
              CC% = U&
              GoTo 99
           End If
       Next U&
       '
       If ULTREG&("TACOLOR") >= 255 Then
            Call MENSERR(24, "Saturación de Tabla de Colores.\Imposible Agregar Color '" + TC$ + "' a Tabla.")
            CC% = 0
          Else
            CC% = 1 + ULTREG&("TACOLOR")
            Call REGAPP("TACOLOR", TC$)
            Call CIERRARCH("TACOLOR")
            Call FRESHECHO("TACOLOR")
       End If
    End If
    '
99  CCOLOR% = CC%
    '
End Function
'

Function LCOLOR$(CC%)
   '
   LCLR$ = ""
   If CC% > 0 Then
      If CC% <= ULTREG&("TACOLOR") Then
         U& = CC%
         LCLR$ = TRIM$(Left$(REGLEIDO$("TACOLOR", U&), 31))
      End If
   End If
   '
99 LCOLOR$ = LCLR$
   '
End Function
'

Function CTALLE%(TCX$)
    '
    TC$ = TRIM$(UCase$(TCX$))
    CC% = 0
    '
    If TC$ <> "" Then
       For U& = 1 To ULTREG&("TATALLES")
           XX$ = TRIM$(Left$(REGLEIDO$("TATALLES", U&), 15))
           If XX$ = TC$ Then
              CC% = U&
              GoTo 99
           End If
       Next U&
       '
       If ULTREG&("TATALLES") >= 255 Then
            Call MENSERR(24, "Saturación de Tabla de Talles.\Imposible Agregar Talle '" + TC$ + "' a Tabla.")
            CC% = 0
          Else
            CC% = 1 + ULTREG&("TATALLES")
            Call REGAPP("TATALLES", TC$)
            Call CIERRARCH("TATALLES")
            Call FRESHECHO("TATALLES")
       End If
    End If
    '
99  CTALLE% = CC%
    '
End Function
'
Function LTALLE$(CC%)
   '
   LCLR$ = ""
   If CC% > 0 Then
      If CC% <= ULTREG&("TATALLES") Then
         U& = CC%
         LCLR$ = TRIM$(Left$(REGLEIDO$("TATALLES", U&), 15))
      End If
   End If
   '
99 LTALLE$ = LCLR$
   '
End Function
'
Function LTITPE$(CP)
   '
   LCLR$ = ""
   If CP > 0 Then
        LCLR$ = TRIM$(CSTRING$(MKS$(CP), 3, 10, 5, 2))
      ElseIf CP < 0 Then
        LCLR$ = TRIM$(CSTRING$(MKS$(Abs(CP)), 3, 8, 2, 2))
        LCLR$ = REPLACAR$(LCLR$, ".", "/")
   End If
   '
99 LTITPE$ = LCLR$
   '
End Function
'
Function PRELISTA#(APREC$, CIIXXI%)
   '
   Static APRECA$, UREPRE&
   '
   If APREC$ <> APRECA$ Then
     APRECA$ = APREC$
     UREPRE& = ULTREG&(APREC$)
   End If
   '
   PRELISTA# = 0: VUELTA% = 0
   If CIIXXI% > 0 Then
     If CIIXXI% <= NUMAS% Then
       AP$ = APREC$
10     RFFF$ = RECFILT$(AP$, 1, MKI$(CIIXXI%))
       If Len(RFFF$) >= 4 Then
         REPRE& = CVS(Left$(RFFF$, 4))
         If REPRE& > 0 Then
           If REPRE& <= UREPRE& Then
             TTZ$ = REGLEIDO$(AP$, REPRE&)
             If CVI(Left$(TTZ$, 2)) = CIIXXI% Then
                 PRELISTA# = CVD(Mid$(TTZ$, 62, 8))
               Else
                 If VUELTA% < 1 Then
                   VUELTA% = VUELTA% + 1
                   TTZ$ = REGLEIDO$(AP$, 1)
                   Call GRAREG(AP$, TTZ$, 1)
                   Call CIERRARCH(AP$)
                   GoTo 10
                 End If
             End If
           End If
         End If
       End If
     End If
   End If
   '
End Function
'
' ***************************************************************************
'
Function NUORIG$(APREC$, CI%)
'
AP$ = APREC$
NUORIG$ = ""
RFFX$ = RECFILT$(AP$, 1, MKI$(CI%))
If Len(RFFX$) >= 4 Then
  RELIP& = CVS(Left$(RFFX$, 4))
  If RELIP& > 0 Then
    TTZ$ = REGLEIDO$(AP$, RELIP&)
    If Len(TTZ$) >= 61 Then
      NUORIG$ = TRIM$(Mid$(TTZ$, 46, 16))
    End If
  End If
End If
'
End Function
'
' ***************************************************************************
'
Function NUOCOM$(NC%, CI%)
'
Static FINOC&, FINAD&, CTX%
If CTX% < 1 Then
    FINOC& = ULTREG&("OCOTERM")
    FINAD& = ULTREG&("DATADI")
    CTX% = 1
End If
'
For V& = 1 To FINOC&
   Y$ = REGLEIDO$("OCOTERM", V&)
   If CVI(Left$(Y$, 2)) = NC% Then
      If CVI(Mid$(Y$, 3, 2)) = CI% Then
         NUOCOM$ = TRIM$(Mid$(Y$, 5, 16))
         Exit Function
      End If
   End If
Next V&
'
For V& = 1 To FINAD&
   Y$ = REGLEIDO$("DATADI", V&)
   If CVI(Left$(Y$, 2)) = CI% Then
      If CVI(Mid$(Y$, 3, 2)) = NC% Then
         NUOCOM$ = TRIM$(Mid$(Y$, 5, 12))
         Exit Function
      End If
   End If
Next V&
'
NUOCOM$ = ""
'
End Function
'
' ***************************************************************************
'
Function DESPACHO$(CI%)
'
VUELTA% = 0
'
10 RXX$ = RECFILT$("DESPAIMP", 1, MKI$(CI%))
'
For VV% = 1 To Len(RXX$) Step 4
    V& = CVS(Mid$(RXX$, VV%, 4))
    Y$ = REGLEIDO$("DESPAIMP", V&)
    If CVI(Left$(Y$, 2)) = CI% Then
       If CVD(Mid$(Y$, 41, 8)) >= 0.5 Then
          FEX = CVI(Mid$(Y$, 3, 2))
          DX$ = FECHATEX$(FEX)
          DX$ = TRIM$(Mid$(Y$, 17, 16)) + " - " + DX$ + " - Aduana " + TRIM$(Mid$(Y$, 5, 12))
          DESPACHO$ = DX$
          Exit Function
       End If
     Else
       If VUELTA% = 0 Then
           VUELTA% = 1
           If ULTREG&("DESPAIMP") > 0 Then
               Y$ = REGLEIDO$("DESPAIMP", 1)
               Call GRAREG("DESPAIMP", Y$, 1)
               Call CIERRARCH("DESPAIMP")
           End If
           GoTo 10
       End If
       '
       Call MENSERR(24, "Error de Indexación de B.D. de Despachos.\Consulte con su Soporte de Sistemas FLEXOFT.")
       DESPACHO$ = ""
       Exit Function
       '
    End If
Next VV%
'
DESPACHO$ = ""
'
End Function
'
' ***************************************************************************
'
Sub ACTECOMAS()
    '
    NVI% = ULTREG&("INVERT") - 1
    II$ = REGLEIDO$("INVERT", 1&)
    Call REPLA(II$, MKI$(NVI%), 17, 2)
    Call GRAREG("INVERT", II$, 1&)
    '
    NMS% = ULTREG&("MASTER") - 1
    MM$ = REGLEIDO$("MASTER", 1&)
    Call REPLA(MM$, MKI$(NMS%), 17, 2)
    Call GRAREG("MASTER", MM$, 1&)
    '
    FI% = FILENBR%("INVERT")
        Dim III As String * 18
    FM% = FILENBR%("MASTER")
        Dim MMM As String * 128
        '
    FIECO% = FILENBR%("ECOMAST")
        Dim EL As String * 64
    FICIA% = FILENBR%("CIACTI")
        Dim CL As String * 2
        '
    Call BLOQARCH("ECOMAST")
    Call BLOQARCH("CIACTI")
    '
    J& = 0
    For I& = 1 To NVI%
        Get #FI%, I& + 1, III$
        CODA$ = Left$(III$, 16)
        CODI% = CVI(Mid$(III$, 17, 2))
        '
        If Mid$(CODA$, 16, 1) <> Chr$(96) Then
          If CODI% > 0 Then
            If CODI% <= NMS% Then
              Get #FM%, CODI% + 1, MMM$
              DEXX$ = Mid$(MMM$, 5, 46)
              '
              EX$ = String$(64, 0)
              Call REPLA(EX$, CODA$, 1, 16)
              Call REPLA(EX$, DEXX$, 17, 46)
              Call REPLA(EX$, MKI$(CODI%), 63, 2)
              '
              EL$ = EX$
              CL$ = MKI$(CODI%)
              '
              J& = J& + 1
              Put #FIECO%, J&, EL$
              Put #FICIA%, J& + 1, CL$
            End If
          End If
        End If
    Next I&
    '
    Get #FICIA%, 1, CL$
    JJ% = J&: CL$ = MKI$(JJ%)
    Put #FICIA%, 1, CL$
    '
    For I& = J& + 1 To ULTREG&("ECOMAST")
        EL$ = String$(64, 0)
        Put #FIECO%, I&, EL$
    Next I&
    '
    For I& = J& + 2 To ULTREG&("CIACTI")
        CL$ = String$(2, 0)
        Put #FICIA%, I&, CL$
    Next I&
    '
    Call CIERRARCH("MASTER")
    Call CIERRARCH("INVERT")
    Call CIERRARCH("SALSTOCK")
    Call CIERRARCH("ECOMAST")
    Call CIERRARCH("CIACTI")
    Call LIBARCH("ECOMAST")
    Call LIBARCH("CIACTI")
    '
End Sub
'
Sub CARDEPOS()
   '
   If NUSU% < 1 Then
      SMP = Screen.MousePointer
      Screen.MousePointer = 11
      FIN& = ULTREG&("SUCURSAL")
      For U& = 1 To FIN&
         If U& <= 62 Then
            X$ = REGLEIDO$("SUCURSAL", U&)
            NUSU% = U&
            If Len(X$) >= 64 Then
               COSU%(U&) = Asc(Mid$(X$, 64, 1))
               SUCU$(U&) = Left$(X$, 20)
               ADEPO$(U&) = Mid$(X$, 21, 8)
               If TRIM$(ADEPO$(U&)) = "" Then ADEPO$(U&) = Left$(X$, 8)
               CATEO#(U&) = 0
               CAFIN#(U&) = 0
               DISTO%(U&) = 0
               If UCase$(Mid$(X$, 62, 2) = "NO") Then
                 DISTO%(U&) = (-1)
               End If
            End If
         End If
      Next U&
      '
      If NUSU% < 1 Then
         NUSU% = 1
         COSU%(1) = 1
         SUCU$(1) = "Deposito Central"
         ADEPO$(1) = "Central"
         CATEO#(1) = 0
         CAFIN#(1) = 0
      End If
      Screen.MousePointer = SMP
   End If
   '
End Sub


Sub CARTACLAS()
   '
10 X$ = REGSEL$("TACRICLA")
   If X$ = "" Then Exit Sub
   '
   CLASI% = Asc(Left$(X$, 1))
   CLASX$ = TRIM$(Mid$(X$, 2, 31))
   If CLASI% < 1 Or CLASI% > 8 Or CLASX$ = "" Then Exit Sub
   '
   CLASX$ = REPLACAR$(CLASX$, "/", "-")
   ALISX$ = TRIM$(Str$(CLASI%))
   While Len(ALISX$) < 3: ALISX$ = "0" + ALISX$: Wend
   ACLASI$ = "CLASI" + ALISX$
   Call COPYSTRU("CLASISTO", ACLASI$)
   '
   Call TRALOCAL(ACLASI$, "CLASISTO")
   VTB% = VENTABU%("CLASISTO/NC/TITUPAN=Clasificación Técnica: " + CLASX$)
   If VTB% >= 0 Then
      For U& = 1 To ULTREG&("!CLASISTO")
         X$ = REGLEIDO$("!CLASISTO", U&)
         Call REPLA(X$, Chr$(U&), 1, 1)
         Call GRAREG(ACLASI$, X$, U&)
      Next U&
      Call SETULTREG(ACLASI$, U&)
      Call CIERRARCH(ACLASI$)
   End If
   Call CIERRARCH("*.*")
   GoTo 10
   '
End Sub

Sub CARCRITEC()
   Call TRALOCAL("TACRICLA", "")
   VTB% = VENTABU%("TACRICLA/NC/TITUPAN=Criterios Clasificación Técnica")
   If VTB% >= 0 Then
      For U& = 1 To ULTREG&("!TACRICLA")
         X$ = REGLEIDO$("!TACRICLA", U&)
         Call REPLA(X$, Chr$(U&), 1, 1)
         Call GRAREG("TACRICLA", X$, U&)
      Next U&
   End If
   If ULTREG&("TACRICLA") > 8 Then
      Call SETULTREG("TACRICLA", 8)
   End If
   Call CIERRARCH("TACRICLA")
End Sub

Sub CLASINTER()
   '
   Screen.MousePointer = 11
   J& = 0: FIN& = NUINV%
   '
   For I& = 1 To FIN&
      Call TRACE(20, I&, FIN&)
      X$ = REGLEIDO$("INVERT", I& + 1)
      If Mid$(X$, 16, 1) <> "`" Then
         CI% = CVI(Mid$(X$, 17, 2))
         If CI% > 0 Then
            If CI% <= NUMAS% Then
               REGMA& = CI% + 1
               M$ = REGLEIDO$("MASTER", REGMA&)
               CLASI$ = Mid$(M$, 69, 8)
               J& = J& + 1
               Y$ = MKI$(CI%) + CLASI$
               Call GRAREG("!CLASIMAS", Y$, J&)
            End If
         End If
      End If
   Next I&
   '
   Call SETULTREG("!CLASIMAS", J&)
   Call CIERRARCH("!CLASIMAS")
   '
   ARX$ = ""
   For U% = 1 To ULTREG&("TACRICLA")
      U1$ = TRIM$(Str$(U% + 2))
      U2$ = TRIM$(Str$(U%)): While Len(U2$) < 3: U2$ = "0" + U2$: Wend
      ARX$ = ARX$ + "/ARCHECO(" + U1$ + ")=CLASI" + U2$
   Next U%
   Screen.MousePointer = 1
   '
   VTB% = VENTABU%("CLASIMAS/NC" + ARX$)
   If VTB% < 0 Then
      Call CIERRARCH("*.*")
      Exit Sub
   End If
   '
   Screen.MousePointer = 11
   FIN& = ULTREG&("!CLASIMAS")
   For I& = 1 To FIN&
      Call TRACE(20, I&, FIN&)
      Y$ = REGLEIDO$("!CLASIMAS", I&)
      CI% = CVI(Left$(Y$, 2))
      If CI% > 0 Then
         If CI% <= NUMAS% Then
            CLASI$ = Mid$(Y$, 3, 8)
            REGMA& = CI% + 1
            If REGMA& > 1 Then
               MX$ = REGLEIDO$("MASTER", REGMA&)
               Call REPLA(MX$, CLASI$, 69, 8)
               Call GRAREG("MASTER", MX$, REGMA&)
            End If
         End If
      End If
   Next I&
   '
   Call CIERRARCH("!CLASIMAS")
   Call CIERRARCH("MASTER")
   Screen.MousePointer = 1
   '
End Sub
'

Sub CLASIRANGO()
   '
   CLR$ = OBJPLA$("CLASIRANGO", CLR$)
   Call CIERRARCH("*.*")
   If CLR$ = "" Then Exit Sub
   CLASI% = Asc(Left$(CLR$, 1))
   If CLASI% < 1 Or CLASI% > ULTREG&("TACRICLA") Then Exit Sub
   '
   ALIX$ = TRIM$(Str$(CLASI%))
   While Len(ALIX$) < 3: ALIX$ = "0" + ALIX$: Wend
   ACLASI$ = "CLASI" + ALIX$
   LU$ = LUDAT$
   Call CIERRARCH(ACLASI$)
   Call CIERRARCH("CLASISTO")
   If COPYFILE%(LU$ + ACLASI$ + ".DAT", LU$ + "CLASISTO.DAT") <> 1 Then
      Exit Sub
   End If
   '
   CLR$ = OBJPLA$("CLASIRANGO-2", CLR$)
   If CLR$ = "" Then Exit Sub
   VACLA% = Asc(Mid$(CLR$, 6, 1))
   If VACLA% < 1 Or VACLA% > ULTREG&(ACLASI$) Then Exit Sub
   '
   Screen.MousePointer = 11
   cod1$ = CODEXT$(CVI(Mid$(CLR$, 2, 2))): If TRIM$(cod1$) = "" Then Exit Sub
   cod2$ = CODEXT$(CVI(Mid$(CLR$, 4, 2))): If TRIM$(cod2$) = "" Then Exit Sub
   '
   J& = 0: FIN& = NUINV%
   '
   For I& = 1 To FIN&
      Call TRACE(20, I&, FIN&)
      X$ = REGLEIDO$("INVERT", I& + 1)
      If Mid$(X$, 16, 1) <> "`" Then
         CI% = CVI(Mid$(X$, 17, 2))
         If CI% > 0 Then
            If CI% <= NUMAS% Then
               REGMA& = CI% + 1
               If REGMA& > 1 Then
                  MX$ = REGLEIDO$("MASTER", REGMA&)
                  If Mid$(MX$, 77, 16) >= cod1$ Then
                    If Mid$(MX$, 77, 16) <= cod2$ Then
                      Call REPLA(MX$, Chr$(VACLA%), 68 + CLASI%, 1)
                      Call GRAREG("MASTER", MX$, REGMA&)
                    End If
                  End If
               End If
            End If
         End If
      End If
   Next I&
   '
   Call CIERRARCH("MASTER")
   Screen.MousePointer = 1
   '
End Sub

'
Function GRUCOVEI%(CI%)
'
If CI% = CIANT% Then GoTo 10
'
GCV% = 0
If CI% > 0 Then
   If CI% <= NUMAS% Then
       REG& = CI% + 1: CIANT% = CI%
       XXX$ = REGLEIDO$("MASTER", REG&)
10     GCV% = CVI(Mid$(XXX$, 125, 2))
   End If
End If
'
GRUCOVEI% = GCV%
'
End Function
'
' ***************************************************************************
'
Sub MUESTRADATCOD(CODDX$)
   CICDX% = CODINT%(CODDX$)
   If CICDX% > 0 Then
      If CICDX% <= NUMAS% Then
         REMASX& = CICDX% + 1
         RMCDX$ = REGLEIDO$("MASTER", REMASX&) + String$(52, 0)
         Call REPLA(RMCDX$, MKS$(STOCKACT(CICDX%)), 129, 4)
         Call REPLA(RMCDX$, MKS$(STOMIN(CICDX%)), 133, 4)
         Call REPLA(RMCDX$, MKS$(LOTESTAN(CICDX%)), 137, 4)
         Call REPLA(RMCDX$, MKS$(LOREPOMA(CICDX%)), 141, 4)
         Call REPLA(RMCDX$, TALLE$(CICDX%), 145, 16)
         Call REPLA(RMCDX$, COLOR$(CICDX%), 161, 16)
         Call REPLA(RMCDX$, MKS$(PESUNI(CICDX%)), 177, 4)
         If DERACCE%("VERCOSTO/NOMESS", "Visualizar Costos y Precios") < 1 Then
            Call REPLA(RMCDX$, MKS$(99999999999#), 111, 4)
         End If
         OBXXX$ = OBJPLA$("MUESTRADATCOD", RMCDX$)
      End If
   End If
End Sub

Sub MUESTRADATPRO(NPROV$)
   NPY& = Val(NPROV$)
   If NPY& > 0 Then
     If NPY& <= 32767 Then
       NPX% = NPY&
       If ECOARCH$("PROVED1", MKI$(NPX%)) <> "" Then
         RMCDX$ = String$(1024, 0)
         NC1% = (-1) * NPX%
         Call REPLA(RMCDX$, MKI$(NPX%), 1, 2)
         Call REPLA(RMCDX$, RASOCLI$(NC1%), 3, 62)
         Call REPLA(RMCDX$, TEL2CLI$(NC1%), 65, 32)
         Call REPLA(RMCDX$, DOMICLI$(NC1%), 97, 64)
         Call REPLA(RMCDX$, CPOSCLI$(NC1%), 161, 16)
         Call REPLA(RMCDX$, LOCACLI$(NC1%), 177, 48)
         Call REPLA(RMCDX$, TEL1CLI$(NC1%), 257, 24)
         Call REPLA(RMCDX$, PAISCLI$(NC1%), 225, 32)
         Call REPLA(RMCDX$, PROVCLI$(NC1%), 529, 1)
         Call REPLA(RMCDX$, FAXCLI$(NC1%), 281, 24)
         Call REPLA(RMCDX$, EMAILCLI$(NC1%), 305, 48)
         Call REPLA(RMCDX$, Chr$(PIVACLI%(NC1%)), 530, 1)
         Call REPLA(RMCDX$, CUITCLI$(NC1%), 513, 16)
         Call REPLA(RMCDX$, Chr$(TICUEPRO%(NC1%)), 531, 1)
         Call REPLA(RMCDX$, Chr$(CPAGCLI%(NC1%)), 532, 1)
         OBXXX$ = OBJPLA$("MUESTRADATPRO", RMCDX$)
       End If
     End If
   End If
End Sub


Function LEERPLANO%(CIXXI%, ARCHIPLANO$)
   '
   Static URE&, RUTARGRA$
   '
   If URE& = 0 Then
     URE& = ULTREG&("INDIPLAN")
     If URE& < 1 Then
       URE& = (-1)
     End If
     '
     RUTARGRA$ = TRIM$(CONTROL$("", "RUTARGRA"))
     If RUTARGRA$ = "" Then RUTARGRA$ = TRIM$(LUDAT$)
     If Right$(RUTARGRA$, 1) <> "\" Then
       RUTARGRA$ = RUTARGRA$ + "\"
     End If
   End If
   '
   LPX% = 0: ARCHIPLANO$ = ""
   ' BUSCAR ARCHIVO DEL PLANO DE CIIX%
   APLANO$ = "": Index = 0
   If CIXXI% > 0 Then
     If CIXXI% <= NUMAS% Then
       Screen.MousePointer = 11
       RNC$ = RECFILT$("INDIPLAN", 1, MKI$(CIXXI%))
       If Len(RNC$) >= 4 Then
         REPLANO& = CVS(RNC$)
         If REPLANO& > 0 Then
           If REPLANO& <= URE& Then
             RPLA$ = REGLEIDO$("INDIPLAN", REPLANO&)
             APLANO$ = TRIM$(Mid$(RPLA$, 65, 48))
             If APLANO$ <> "" Then
               FORMALTER.Image3.Picture = LoadPicture("")
               COPLANO$ = RUTARGRA$ + APLANO$
               If COPLANO$ <> "" Then
                 If EXISTE%(COPLANO$) > 0 Then
                   FORMALTER.Image3.Picture = LoadPicture(COPLANO$)
                   LPX% = 1
                   ARCHIPLANO$ = COPLANO$
                 End If
               End If
             End If
           End If
         End If
       End If
     End If
   End If
   '
   LEERPLANO% = LPX%
   '
End Function

Function DESCRIOP$(CIXI%, NOPE%)
   '
   NPLX$ = ""
   RNC$ = RECFILT$("SECOPER", 1, MKI$(CIXI%))
   URN& = ULTREG&("SECOPER")
   '
   For U& = 1 To Len(RNC$) Step 4
     RENOTA& = CVS(Mid$(RNC$, U&, 4))
     If RENOTA& > 0 Then
       If RENOTA& <= URN& Then
         TTXX$ = REGLEIDO$("SECOPER", RENOTA&)
         If CVI(Mid$(TTXX$, 3, 2)) = NOPE% Then
           NPLX$ = TRIM$(Mid$(TTXX$, 5, 48))
           GoTo 29
         End If
       End If
     End If
   Next U&
   '
29 DESCRIOP$ = NPLX$
   '
End Function
'
Function OBSEROP$(CIXI%, NOPE%)
   '
   NPLX$ = ""
   RNC$ = RECFILT$("OBSEOPER", 1, MKI$(CIXI%))
   URN& = ULTREG&("OBSEOPER")
   '
   For U& = 1 To Len(RNC$) Step 4
     RENOTA& = CVS(Mid$(RNC$, U&, 4))
     If RENOTA& > 0 Then
       If RENOTA& <= URN& Then
         TTZZ$ = REGLEIDO$("OBSEOPER", RENOTA&)
         NOPY% = CVI(Mid$(TTZZ$, 3, 2))
         If NOPY% = NOPE% Then
           NPLX$ = NPLX$ + Mid$(TTZZ$, 5)
         End If
       End If
     End If
   Next U&
   '
29 OBSEROP$ = TRIM$(NPLX$)
   '
End Function

Function MAQUISTAN$(CIXI%, NOPE%)
   '
   NPLX$ = ""
   RNC$ = RECFILT$("SECOPER", 1, MKI$(CIXI%))
   URN& = ULTREG&("SECOPER")
   '
   For U& = 1 To Len(RNC$) Step 4
     RENOTA& = CVS(Mid$(RNC$, U&, 4))
     If RENOTA& > 0 Then
       If RENOTA& <= URN& Then
         TTXX$ = REGLEIDO$("SECOPER", RENOTA&)
         If CVI(Mid$(TTXX$, 3, 2)) = NOPE% Then
           NPLX$ = Mid$(TTXX$, 53, 6)
           GoTo 29
         End If
       End If
     End If
   Next U&
   '
29 MAQUISTAN$ = NPLX$
   '
End Function
'
Function TIESTAN(CIXI%, NOPE%, CANTID#, modo%)
   '
   ' MODO% = 0 ==> incluye tiempo de preparación
   ' MODO% > 0 ==> no considera tiempo de preparación
   '
   TSTA = 0
   RNC$ = RECFILT$("SECOPER", 1, MKI$(CIXI%))
   URN& = ULTREG&("SECOPER")
   '
   For U& = 1 To Len(RNC$) Step 4
     RENOTA& = CVS(Mid$(RNC$, U&, 4))
     If RENOTA& > 0 Then
       If RENOTA& <= URN& Then
         TTXX$ = REGLEIDO$("SECOPER", RENOTA&)
         If CVI(Mid$(TTXX$, 3, 2)) = NOPE% Then
           '
           CALOTES = 1
           CAOPS = CVS(Mid$(TTXX$, 83, 4))
           If CAOPS < 0.01 Then CAOPS = 1
           '
           MINPREP = 60 * CVS(Mid$(TTXX$, 59, 4)) * CALOTES
           If modo% > 0 Then MINPREP = 0
           '
           MINFIJO = 60 * CVS(Mid$(TTXX$, 63, 4)) * CALOTES
           MINVARI = CANTID# * CVS(Mid$(TTXX$, 67, 4))
           MINDEMO = (MINFIJO + MINVARI) * CVS(Mid$(TTXX$, 75, 4)) / 100
           MINLIMP = 60 * CVS(Mid$(TTXX$, 79, 4)) * CALOTES
           MINLOTE = MINPREP + MINFIJO + MINVARI + MINDEMO + MINLIMP
           '
           TSTA = CAOPS * MINLOTE
           '
         End If
       End If
     End If
   Next U&
   '
   TIESTAN = TSTA
   '
End Function
'
Function NUHOPROF$(CIXI%, NOPE%)
   '
   NPLX$ = ""
   RNC$ = RECFILT$("DATADOPS", 1, MKI$(CIXI%))
   URN& = ULTREG&("DATADOPS")
   '
   For U& = 1 To Len(RNC$) Step 4
     RENOTA& = CVS(Mid$(RNC$, U&, 4))
     If RENOTA& > 0 Then
       If RENOTA& <= URN& Then
         TTXX$ = REGLEIDO$("DATADOPS", RENOTA&)
         If CVI(Mid$(TTXX$, 3, 2)) = NOPE% Then
           NPLX$ = TRIM$(Mid$(TTXX$, 5, 16))
           GoTo 29
         End If
       End If
     End If
   Next U&
   '
29 NUHOPROF$ = NPLX$

End Function
'
Function RVHOPROF$(CIXI%, NOPE%)
   '
   NPLX$ = ""
   RNC$ = RECFILT$("DATADOPS", 1, MKI$(CIXI%))
   URN& = ULTREG&("DATADOPS")
   '
   For U& = 1 To Len(RNC$) Step 4
     RENOTA& = CVS(Mid$(RNC$, U&, 4))
     If RENOTA& > 0 Then
       If RENOTA& <= URN& Then
         TTXX$ = REGLEIDO$("DATADOPS", RENOTA&)
         If CVI(Mid$(TTXX$, 3, 2)) = NOPE% Then
           NPLX$ = TRIM$(Mid$(TTXX$, 21, 4))
           GoTo 29
         End If
       End If
     End If
   Next U&
   '
29 RVHOPROF$ = NPLX$

End Function
'
Function NUPLANO$(CIXI%)
   '
   NPLX$ = ""
   RNC$ = RECFILT$("INDIPLAN", 1, MKI$(CIXI%))
   If Len(RNC$) >= 4 Then
     REPLANO& = CVS(RNC$)
     If REPLANO& > 0 Then
       If REPLANO& <= ULTREG&("INDIPLAN") Then
         RPLA$ = REGLEIDO$("INDIPLAN", REPLANO&)
         NPLX$ = TRIM$(Mid$(RPLA$, 3, 16))
       End If
     End If
   End If
   '
   NUPLANO$ = NPLX$
   '
End Function

Function RVPLANO$(CIXI%)
   '
   NPLX$ = ""
   RNC$ = RECFILT$("INDIPLAN", 1, MKI$(CIXI%))
   If Len(RNC$) >= 4 Then
     REPLANO& = CVS(RNC$)
     If REPLANO& > 0 Then
       If REPLANO& <= ULTREG&("INDIPLAN") Then
         RPLA$ = REGLEIDO$("INDIPLAN", REPLANO&)
         NPLX$ = TRIM$(Mid$(RPLA$, 19, 4))
       End If
     End If
   End If
   '
   RVPLANO$ = NPLX$
   '
End Function

Function FEPLANO$(CIXI%)
   '
   NPLX$ = ""
   RNC$ = RECFILT$("INDIPLAN", 1, MKI$(CIXI%))
   If Len(RNC$) >= 4 Then
     REPLANO& = CVS(RNC$)
     If REPLANO& > 0 Then
       If REPLANO& <= ULTREG&("INDIPLAN") Then
         RPLA$ = REGLEIDO$("INDIPLAN", REPLANO&)
         FEX = CVI(Mid$(RPLA$, 23, 2))
         NPLX$ = FECHATEX$(FEX)
       End If
     End If
   End If
   '
   FEPLANO$ = NPLX$
   '
End Function

Function PZPLANO$(CIXI%)
   '
   NPLX$ = ""
   RNC$ = RECFILT$("INDIPLAN", 1, MKI$(CIXI%))
   If Len(RNC$) >= 4 Then
     REPLANO& = CVS(RNC$)
     If REPLANO& > 0 Then
       If REPLANO& <= ULTREG&("INDIPLAN") Then
         RPLA$ = REGLEIDO$("INDIPLAN", REPLANO&)
         NPLX$ = TRIM$(Mid$(RPLA$, 53, 12))
       End If
     End If
   End If
   '
   PZPLANO$ = NPLX$
   '
End Function

Function ESPECIFI$(CIXI%)
   '
   NPLX$ = ""
   RNC$ = RECFILT$("ESPECIFI", 1, MKI$(CIXI%))
   URN& = ULTREG&("ESPECIFI")
   '
   For U& = 1 To Len(RNC$) Step 4
     RENOTA& = CVS(Mid$(RNC$, U&, 4))
     If RENOTA& > 0 Then
       If RENOTA& <= URN& Then
         TTYY$ = RTrim$(Mid$(REGLEIDO$("ESPECIFI", RENOTA&), 3))
         If NPLX$ <> "" Then
           NPLX$ = NPLX$ + Chr$(13) + Chr$(10)
         End If
         NPLX$ = NPLX$ + RTrim$(TTYY$)
       End If
     End If
   Next U&
   '
   ESPECIFI$ = NPLX$
   '
End Function

Function NUHOREI$(CIXI%)
   '
   NPLX$ = ""
   RNC$ = RECFILT$("DATADINS", 1, MKI$(CIXI%))
   URN& = ULTREG&("DATADINS")
   '
   If Len(RNC$) >= 4 Then
     RENOTA& = CVS(Left$(RNC$, 4))
     If RENOTA& > 0 Then
       If RENOTA& <= URN& Then
         TTXX$ = REGLEIDO$("DATADINS", RENOTA&)
         NPLX$ = Mid$(TTXX$, 5, 16)
       End If
     End If
   End If
   '
   NUHOREI$ = NPLX$
   '
End Function

Function RVHOREI$(CIXI%)
   '
   NPLX$ = ""
   RNC$ = RECFILT$("DATADINS", 1, MKI$(CIXI%))
   URN& = ULTREG&("DATADINS")
   '
   If Len(RNC$) >= 4 Then
     RENOTA& = CVS(Left$(RNC$, 4))
     If RENOTA& > 0 Then
       If RENOTA& <= URN& Then
         TTXX$ = REGLEIDO$("DATADINS", RENOTA&)
         NPLX$ = Mid$(TTXX$, 21, 4)
       End If
     End If
   End If
   '
   RVHOREI$ = NPLX$
   '
End Function

Function NUPLIEN$(CIXI%)
   '
   NPLX$ = ""
   RNC$ = RECFILT$("DATADINS", 1, MKI$(CIXI%))
   URN& = ULTREG&("DATADINS")
   '
   If Len(RNC$) >= 4 Then
     RENOTA& = CVS(Left$(RNC$, 4))
     If RENOTA& > 0 Then
       If RENOTA& <= URN& Then
         TTXX$ = REGLEIDO$("DATADINS", RENOTA&)
         NPLX$ = Mid$(TTXX$, 25, 16)
       End If
     End If
   End If
   '
   NUPLIEN$ = NPLX$
   '
End Function

Function RVPLIEN$(CIXI%)
   '
   NPLX$ = ""
   RNC$ = RECFILT$("DATADINS", 1, MKI$(CIXI%))
   URN& = ULTREG&("DATADINS")
   '
   If Len(RNC$) >= 4 Then
     RENOTA& = CVS(Left$(RNC$, 4))
     If RENOTA& > 0 Then
       If RENOTA& <= URN& Then
         TTXX$ = REGLEIDO$("DATADINS", RENOTA&)
         NPLX$ = Mid$(TTXX$, 41, 4)
       End If
     End If
   End If
   '
   RVPLIEN$ = NPLX$
   '
End Function

Function NUHOJES$(CIXI%)
   '
   NPLX$ = ""
   RNC$ = RECFILT$("INDIPLAN", 1, MKI$(CIXI%))
   If Len(RNC$) >= 4 Then
     REPLANO& = CVS(RNC$)
     If REPLANO& > 0 Then
       If REPLANO& <= ULTREG&("INDIPLAN") Then
         RPLA$ = REGLEIDO$("INDIPLAN", REPLANO&)
         NPLX$ = TRIM$(Mid$(RPLA$, 33, 16))
       End If
     End If
   End If
   '
   NUHOJES$ = NPLX$
   '
End Function

Function RVHOJES$(CIXI%)
   '
   NPLX$ = ""
   RNC$ = RECFILT$("INDIPLAN", 1, MKI$(CIXI%))
   If Len(RNC$) >= 4 Then
     REPLANO& = CVS(RNC$)
     If REPLANO& > 0 Then
       If REPLANO& <= ULTREG&("INDIPLAN") Then
         RPLA$ = REGLEIDO$("INDIPLAN", REPLANO&)
         NPLX$ = TRIM$(Mid$(RPLA$, 49, 4))
       End If
     End If
   End If
   '
   RVHOJES$ = NPLX$
   '
End Function

Function DEGRUVEN$(CL%)
   DX$ = "Varios"
   If CL% > 0 Then DX$ = ECOARCH$("GRUCOVEN", MKI$(CL%))
   If DX$ = "" Then Beep: CL% = 0: DX$ = "Varios"
   DEGRUVEN$ = DX$
End Function
'
Sub CLASIAUTDOSI()   ' clasificacion automatica dosivac
   '
   Screen.MousePointer = 11
   Dim ACTI%(32767)
   '
   NM% = NUMAS%
   FIN& = ULTREG&("ESTRUNUE")
   For REG& = 1 To FIN&
        '
        Call TRACE(20, REG&, FIN&)
        X$ = REGLEIDO$("ESTRUNUE", REG&)
        CI% = CVI(Left$(X$, 2) + String$(2, 0))
        If CI% > 0 Then
            If CI% <= NM% Then
               CODYY$ = CODEXT$(CI%)
               If Mid$(CODYY$, 16, 1) <> Chr$(96) Then
                  If ACTI%(CI%) = 0 Then
                     ACTI%(CI%) = 1
                  End If
               End If
            End If
        End If
        '
        DPS$ = Mid$(X$, 3, 120)
        If ULTREG&("ESTRUDOS") >= REG& Then
             DPS$ = DPS$ + REGLEIDO$("ESTRUDOS", REG&)
        End If
        '
        For K% = 1 To Len(DPS$) Step 8
            DPSX$ = Mid$(DPS$, K%, 8)
            CI% = CVI(Left$(DPSX$, 2))
            If CI% <= 0 Then GoTo 99
            If CI% <= NUMAS% Then
                ACTI%(CI%) = (-1)
            End If
        Next K%
        '
99 Next REG&
   '
   FIN& = NM%
   For REG& = 1 To FIN&
     Call TRACE(20, REG&, FIN&)
     CIXI% = REG&
     If ACTI%(CIXI%) = 1 Then
       If TIMATE%(CIXI%) < 1 Then
         XX$ = REGLEIDO$("MASTER", REG& + 1)
         Call REPLA(XX$, Chr$(5), 93, 1)
         Call GRAREG("MASTER", XX$, REG& + 1)
       End If
     End If
   Next REG&
   Call CIERRARCH("MASTER")
   '
   Screen.MousePointer = 1
   '
End Sub

Function CODORIG$(CIXXX%, NCXXX%)
   '
   ASELE$ = "REPLACOC"
   If NCXXX% < 0 Then ASELE$ = "REPLACOP"
   '
   CDXXX$ = CODEXT$(CIXXX%)
   UREP& = ULTREG&(ASELE$)
   RFF$ = RECFILT$(ASELE$, 3, MKI$(CIXXX%))
   '
   For KKI% = 1 To Len(RFF$) Step 4
     REREP& = CVS(Mid$(RFF$, KKI%, 4))
     If REREP& > 0 Then
       If REREP& <= UREP& Then
         XXYY$ = REGLEIDO$(ASELE$, REREP&)
         If CVI(Mid$(XXYY$, 1, 2)) = Abs(NCXXX%) Then
           CDYYY$ = TRIM$(Mid$(XXYY$, 5, 28))
           If CDYYY$ <> "" Then
             CDXXX$ = CDYYY$
             GoTo 99
           End If
         End If
       End If
     End If
   Next KKI%
   '
99 CODORIG$ = CDXXX$
   '
End Function

Function LOTESEL$(CIXI%)
   '
   Dim TLOTE$(256), CALOTE#(256)
   LSX$ = ""
   '
   If CIXI% > 0 Then
     CANLOT% = 0
     RFF$ = RECFILT$("MOVISTO", 2, MKI$(CIXI%))
     For U& = 1 To Len(RFF$) Step 4
       REMO& = CVS(Mid$(RFF$, U&, 4))
       X$ = REGLEIDO$("MOVISTO", REMO&)
       ENTRA# = CVD(Mid$(X$, 88, 8))
       If ENTRA# >= 0.05 Then
         IDENLO$ = TRIM$(Mid$(X$, 5, 16))
         If IDENLO$ = "" Then IDENLO$ = TRIM$(Mid$(X$, 23, 10))
         If IDENLO$ = "" Then IDENLO$ = TRIM$(Mid$(X$, 33, 10))
         IDENLO$ = AJUSTI$(IDENLO$, 16)
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
         FEX = CVI(Left$(X$, 2))
         CANLOT% = CANLOT% + 1
         TTXX$ = AJUSTI$(IDENLO$, 40)
         Call REPLA(TTXX$, FECHATEX$(FEX), 19, 8)
         Call REPLA(TTXX$, CSTRING$(MKD$(ENTRA#), 4, 12, 1, 2), 27, 12)
         TLOTE$(CANLOT%) = TTXX$
         CALOTE#(KKI%) = ENTRA#
         '
       End If
19   Next U&
     '
     For U& = 1 To Len(RFF$) Step 4
       REMO& = CVS(Mid$(RFF$, U&, 4))
       X$ = REGLEIDO$("MOVISTO", REMO&)
       SALID# = CVD(Mid$(X$, 96, 8))
       If SALID# >= 0.05 Then
         IDENLO$ = AJUSTI$(Mid$(X$, 5, 16), 16)
         IDENLOX$ = TRIM$(IDENLO$)
         '
35       For KKI% = 1 To CANLOT%
           If Left$(TLOTE$(KKI%), 16) = IDENLO$ Or IDENLOX$ = "" Then
             If SALID# <= CALOTE#(KKI%) Then
                 CALOTE#(KKI%) = CALOTE#(KKI%) - SALID#
                 SALID# = 0
               Else
                 SALID# = SALID# - CALOTE#(KKI%)
                 CALOTE#(KKI%) = 0
             End If
           End If
         Next KKI%
         If SALID# >= 0.05 Then
           IDENLOX$ = ""
           GoTo 35
         End If
         '
       End If
     Next U&
     '
     REBLA$ = REGBLAN$("STOLOTE"): JJ& = 0
     For KKI% = 1 To CANLOT%
       If CALOTE#(KKI%) > 0.05 Then
         ZZ$ = REBLA$
         Call REPLA(ZZ$, TLOTE$(KKI%), 1, 38)
         Call REPLA(ZZ$, MKD$(CALOTE#(KKI%)), 41, 8)
         JJ& = JJ& + 1
         Call GRAREG("!STOLOTE", ZZ$, JJ&)
       End If
     Next KKI%
     Call SETULTREG("!STOLOTE", JJ&)
     Call CIERRARCH("!STOLOTE")
     '
     Call FRESHECHO("!STOLOTE")
     Call SELECHO("!STOLOTE/Id.Lote                       F.Alta                Q.Orig                      Saldo")
     LSX$ = VALACT1$("!STOLOTE")
   End If
   '
   LOTESEL$ = LSX$
   '
End Function

Sub DESHASCOD(cod1$, cod2$, CIKX1%, CIKX2%)
   '
10 cod1$ = "": cod2$ = "": CIKX1% = 0: CIKX2% = 0
   XX$ = Left$(REGLEIDO$("ECOMAST", 1), 16): CI1% = Abs(CODINT%(XX$))
   XX$ = Left$(REGLEIDO$("ECOMAST", ULTREG&("ECOMAST")), 16): CI2% = Abs(CODINT%(XX$))
   VDEF$ = MKI$(CI1%) + MKI$(CI2%)
   X$ = OBJPLA$("DESHASCOD", MKI$(CI1%) + MKI$(CI2%))
   If Len(X$) < 4 Then
     Exit Sub
   End If
   '
   Screen.MousePointer = 11
   '
   CIKX1% = CVI(Left$(X$, 2))
   cod1$ = CODEXT$(CIKX1%): If TRIM$(cod1$) = "" Then Exit Sub
   CIKX2% = CVI(Mid$(X$, 3, 2))
   cod2$ = CODEXT$(CIKX2%): If TRIM$(cod2$) = "" Then Exit Sub
   If CIKX1% < 1 Or CIKX2% < 1 Then GoTo 10
   If cod2$ < cod1$ Then GoTo 10
   '
End Sub

Function PROVHABI%(CI%)
   '
   PROVHABI% = 0        ' FABRICACION PROPIA
   If CI% > 0 Then
     If CI% <= NUMAS% Then
       REG& = CI% + 1
       XXX$ = REGLEIDO$("MASTER", REG&)
10     NPRO% = CVI(Mid$(XXX$, 123, 2))
       If NPRO% > 0 And ECOARCH$("ACREDOR", MKI$(NPRO%)) <> "" Then
         PROVHABI% = NPRO%
       End If
     End If
   End If
   '
End Function
'
' ***************************************************************************
'
Function COSCALCU(CI%, MODOCO%)
  '
  COU = 0
  RFG$ = RECFILT$("RECOSCAL", 1, MKI$(CI%))
  If Len(RFG$) >= 4 Then
    REGB& = CVS(RFG$)
    If REGB& > 0 Then
      Q$ = REGLEIDO$("RECOSCAL", REGB&)
      COU = CVD(Mid$(Q$, 89, 8))
    End If
  End If
  '
  MODOCO% = 1
  If COU < 0.0005 Then
    COU = COSUNI(CI%) * TICAMONE(MONECOSTO%(CI%)) * (1 + PORNACIO(CI%) / 100)
    MODOCO% = 2
  End If
  '
  COSCALCU = COU
  '
End Function
'
' ***************************************************************************
'
Function PORNACIO(CI%)
  '
  If CI% = CIANT% Then GoTo 10
  '
  PORNACIO = 0
  If CI% > 0 Then
    If CI% <= NUMAS% Then
      REG& = CI% + 1: CIANT% = CI%
      XXX$ = REGLEIDO$("MASTER", REG&)
10    PORNACIO = CVI(Mid$(XXX$, 127, 2)) / 100
      If MONECOSTO%(CI%) <= 1 Then PORNACIO = 0
    End If
  End If
  '
End Function
'
Function MONELISTA%(ILISTA%) ' se traslado esta funcion desde el moduo glo_variabs.bas
   '
   Static CTXX%
   Static MNL%(256)
   Static FICATO#
   '
   ILIS% = ILISTA%: If ILIS% > 255 Then ILIS% = 0
   If EXISTE%(LUDAT$ + "LISTAS.DAT") > 0 Then
     If FileDateTime(LUDAT$ + "LISTAS.DAT") <> FICATO# Then
       CTXX% = 0
       FICATO# = FileDateTime(LUDAT$ + "LISTAS.DAT")
     End If
   End If
   '
   If CTXX% = 0 Then
     CTXX% = 1
     For U& = 1 To ULTREG&("LISTAS")
       XX$ = REGLEIDO$("LISTAS", U&)
       CTG% = Asc(Left$(XX$, 1))
       MNL%(CTG%) = Asc(Mid$(XX$, 50, 1))
     Next U&
     MNL%(0) = 1
   End If
   '
   MONELISTA% = MNL%(ILIS%)
   '
End Function
'
Function DECILISTA%(ILISTA%)           ' DECIMALES DE LISTA DE PRECIOS
   '
   Static CTXX%
   Static DCL%(256)
   Static FICATO#
   '
   ILIS% = ILISTA%: If ILIS% > 255 Then ILIS% = 0
   If EXISTE%(LUDAT$ + "LISTAS.DAT") > 0 Then
     If FileDateTime(LUDAT$ + "LISTAS.DAT") <> FICATO# Then
       CTXX% = 0
       FICATO# = FileDateTime(LUDAT$ + "LISTAS.DAT")
     End If
   End If
   '
   If CTXX% = 0 Then
     CTXX% = 1
     For U& = 1 To ULTREG&("LISTAS")
       XX$ = REGLEIDO$("LISTAS", U&)
       CTG% = Asc(Left$(XX$, 1))
       CANDECI% = Asc(Mid$(XX$, 49, 1))
         If CANDECI% < 2 Then CANDECI% = 2
         If CANDECI% > 5 Then CANDECI% = 5
      DCL%(CTG%) = CANDECI%
     Next U&
     DCL%(0) = 2
   End If
   '
   DECILISTA% = DCL%(ILIS%)
   '
End Function
'
Function DENOLISTA$(ILISTA%) ' se traslado esta funcion desde el moduo glo_variabs.bas
   '
   ILIS% = ILISTA%: If ILIS% > 255 Then ILIS% = 0
   DNL$ = "Precios Concertados (Pesos)"
   If ILIS% > 0 Then
     DNL$ = TRIM$(ECOARCH$("LISTAS", Chr$(ILIS%)))
     MNL% = MONELISTA%(ILIS%)
     If MNL% = 2 Then
         DNL$ = DNL$ + " (Dólares)"
       Else
         DNL$ = DNL$ + " (Pesos)"
     End If
   End If
   DENOLISTA$ = DNL$
   '
End Function


