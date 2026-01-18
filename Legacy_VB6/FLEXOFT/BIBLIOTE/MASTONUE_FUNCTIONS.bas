Attribute VB_Name = "MASTONUE_FUNCTIONS"
Public PUNCODIG As String
Public BUFERCOD As String * 2048
Public NUSU%, COSU%(64), SUCU$(64), ADEPO$(64), DISTO%(64), CATEO#(64), CAFIN#(64)
Dim CIANT%, XXX$
Public PHABIT%(32767)
Public HAYACMAS%
Public CONTROLINTEG%
'
Sub APERBASDAT(NBASE$)
   TIBASE$ = "DAT"
End Sub

Function DECILISTA%(NLX%)
    DECILISTA% = 4
End Function
'
Function DENOLISTA$(NLX%)
    DENOLISTA$ = "LISTA NUMERO " & NLX%
End Function

Function MONELISTA%(NLX%)
    MONELISTA% = 1
End Function

Function NUINV%()
'
X$ = REGLEIDO$("INVERT", 1&)
NUINV% = CVI(Mid$(X$, 17, 2))
End Function
' ***************************************************************************

Function NUMAS%()
   '
   Static FEHOANT#
   On Error Resume Next
   FEHOMAS# = FileDateTime(LUDAT$ + "MASTER.DAT")
   On Error GoTo 0
   If FEHOMAS# <> FEHOANT# Or FEHOANT# < 0.01 Then
      Call CIERRARCH("MASTER")
      FEHOANT# = FEHOMAS#
      '
      If ULTREG&("MASTER") < 1 Then
         MSY$ = String$(128, 0)
         FLIM% = 32767
         Mid$(MSY$, 124, 2) = MKI$(FLIM%)
         Mid$(MSY$, 126, 2) = MKI$(9876 - FLIM%)
         PPX% = 23
         For I% = 124 To 127
           PPX% = PPX% + (I% - 123) * Asc(Mid$(MSY$, I%, 1))
         Next I%
         PPX% = PPX% Mod 256
         Mid$(MSY$, 128, 1) = Chr$(PPX%)
         Call GRAREG("MASTER", MSY$, 1)
         Call CIERRARCH("MASTER")
         Call BAJALDISCO
      End If
      '
   End If
   '
   X$ = REGLEIDO$("MASTER", 1&)
   NUMAS% = CVI(Left$(X$, 2))
   '
End Function
' ***************************************************************************

Function CODINT%(CODEXTE$)
   '
   Static FEHOANT#
   '
   CODINT% = 0
   If EXISTE%(LUDAT$ + "PUNCODIG.DAT") < 1 Then
      Call REPUNCODIG
   End If
   '
   FEHOPUN# = FileDateTime(LUDAT$ + "PUNCODIG.DAT")
   If FEHOPUN# <> FEHOANT# Then
      FEHOANT# = FEHOPUN#
      PUNCODIG$ = ""
   End If
   '
   If PUNCODIG$ = "" Then
      Screen.MousePointer = 11
      PUNCODIG$ = LOADFILE$(LUDAT$ + "PUNCODIG.DAT")
      NMS& = NUMAS%
      LOPUCO& = 16 * NMS&
      If Len(PUNCODIG$) < LOPUCO& Then
        Call REPUNCODIG
      End If
      PUNCODIG$ = Left$(PUNCODIG$, LOPUCO&)
      Screen.MousePointer = 1
   End If
   '
   CODDXX$ = AJUSTI$(CODEXTE$, 16)
10 PPXL& = InStr(PUNCODIG$, CODDXX$)
   If PPXL& Mod 16 = 1 Then
      CODINT% = 1 + Int(PPXL& / 16)
      If Mid$(REGLEIDO$("MASTER", CODINT% + 1&), 77, 15) <> Left$(CODDXX$, 15) Then
         Call REPUNCODIG
         GoTo 10
      End If
   End If
   '
End Function

'
' ***************************************************************************
'
Function CODEXT$(CI%)
'
If CI% = CIANT% Then GoTo 10
'
CODEXT$ = ""
If CI% > 0 Then
   If CI% <= NUMAS% Then
       REG& = CI% + 1: CIANT% = CI%
       XXX$ = REGLEIDO$("MASTER", REG&)
10     CODEXT$ = Mid$(XXX$, 77, 16)
   End If
End If
End Function
'
' ***************************************************************************
'
Function COTIPIF$(CI%)
  RECO$ = FILTERGETRECORD$("COEQTIP", 1, MKI$(CI%))
  COTIPIF$ = TRIM$(Mid$(RECO$, 9, 24))
End Function
' ***************************************************************************
'
Function DESCRIT$(CI%)
'
If CI% = CIANT% Then GoTo 10
'
DESCRIT$ = ""
If CI% > 0 Then
   If CI% <= NUMAS% Then
       REG& = CI% + 1: CIANT% = CI%
       XXX$ = REGLEIDO$("MASTER", REG&)
10     DESCRIT$ = TRIM$(Mid$(XXX$, 5, 64)) + " (" + UNIMED$(CI%) + ")"
   End If
End If
End Function
'

' ***************************************************************************
'
Function DESCRIT0$(CI%)
'
If CI% = CIANT% Then GoTo 10
'
DESCRIT0$ = ""
If CI% > 0 Then
   If CI% <= NUMAS% Then
       REG& = CI% + 1: CIANT% = CI%
       XXX$ = REGLEIDO$("MASTER", REG&)
10     DESCRIT0$ = TRIM$(Mid$(XXX$, 5, 64))
   End If
End If
End Function

'
' ***************************************************************************
'
Function UNIMED$(CI%)
'
If CI% = CIANT% Then GoTo 10
'
UNIMED$ = ""
If CI% > 0 Then
   If CI% <= NUMAS% Then
       REG& = CI% + 1: CIANT% = CI%
       XXX$ = REGLEIDO$("MASTER", REG&)
10     UNIMED$ = Mid$(XXX$, 109, 2)
       If TRIM$(UNIMED$) = "" Then
           Call REPLA(XXX$, "U.", 109, 2)
           If REG& > 0 Then Call GRAREG("MASTER", XXX$, REG&)
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
Function FEREFOR%(CI%)   ' FECHA REVISION FORMULA
  '
  FRF% = 0
  If CI% > 0 Then
    XX$ = FILTERGETRECORD$("REVIFORM", 1, MKI$(CI%))
    FRF% = CVI(Mid$(XX$, 3, 2))
  End If
  FEREFOR% = FRF%
  '
End Function
'
' ***************************************************************************
'
Function USREFOR$(CI%)    ' USUARIO RESPONSABLE REVISION
  '
  URF$ = ""
  If CI% > 0 Then
    XX$ = FILTERGETRECORD$("REVIFORM", 1, MKI$(CI%))
    URF$ = TRIM$(Mid$(XX$, 5, 28))
  End If
  USREFOR$ = URF$
  '
End Function
'
' ***************************************************************************
'
Function STATUIT%(CI%)
'
STATUIT% = 1
If CI% = CIANT% Then GoTo 10
'
If CI% > 0 Then
  If CI% <= NUMAS% Then
    REG& = CI% + 1: CIANT% = CI%
    XXX$ = REGLEIDO$("MASTER", REG&)
10  If Mid$(XXX$, 92, 1) = "`" Then
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
If CI% = CIANT% Then GoTo 10
'
COLOR$ = ""
If CI% > 0 Then
   If CI% <= NUMAS% Then
       REG& = CI% + 1: CIANT% = CI%
       XXX$ = REGLEIDO$("MASTER", REG&)
10     CC% = Asc(Mid$(XXX$, 96, 1))
       COLOR$ = LCOLOR$(CC%)
   End If
End If
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
' ***************************************************************************
'
Function TALLE$(CI%)
'
If CI% = CIANT% Then GoTo 10
'
TALLE$ = ""
If CI% > 0 Then
   If CI% <= NUMAS% Then
       REG& = CI% + 1: CIANT% = CI%
       XXX$ = REGLEIDO$("MASTER", REG&)
10     CC% = Asc(Mid$(XXX$, 97, 1))
       TALLE$ = LTALLE$(CC%)
   End If
End If
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
' ***************************************************************************
'
Function LTITPE$(CP)
   '
   LCLR$ = ""
   If CP > 0 Then
        LCLR$ = CSTRING$(MKS$(CP), 3, 10, 5, 2)
      ElseIf CP < 0 Then
        LCLR$ = TRIM$(CSTRING$(MKS$(Abs(CP)), 3, 8, 2, 2))
        LCLR$ = REPLACAR$(LCLR$, ".", "/")
   End If
   '
99 LTITPE$ = LCLR$
   '
End Function
'
Function TIPOHIL$(CI%)
'
TIPOHIL$ = TALLE$(CI%)
End Function
'
' ***************************************************************************
'
Function PESUNI(CI%)
'
If CI% = CIANT% Then GoTo 10
'
PESUNI = 0
If CI% > 0 Then
   If CI% <= NUMAS% Then
       REG& = CI% + 1: CIANT% = CI%
       XXX$ = REGLEIDO$("MASTER", REG&)
10     PP = CVS(Mid$(XXX$, 99, 4))
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
'
If CI% = CIANT% Then GoTo 10
'
TITULO$ = ""
If CI% > 0 Then
   If CI% <= NUMAS% Then
       REG& = CI% + 1: CIANT% = CI%
       XXX$ = REGLEIDO$("MASTER", REG&)
10     CP = CVS(Mid$(XXX$, 99, 4))
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
  If CI% = CIANT% Then GoTo 10
  '
  COSUNI = 0
  If CI% > 0 Then
    If CI% <= NUMAS% Then
      REG& = CI% + 1: CIANT% = CI%
      XXX$ = REGLEIDO$("MASTER", REG&)
10    COSUNI = CVS(Mid$(XXX$, 111, 4))
    End If
  End If
  '
End Function
'
Function COSPESOS(CI%)
    COSPESOS = COSUNI(CI%) * TICAMONE(MONECOSTO%(CI%)) * (1 + PORNACIO(CI%) / 100)
End Function
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
Function COEFNACIO(CI%)
  '
  COEFNACIO = 1 + PORNACIO(CI%) / 100
  '
End Function
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
' ***************************************************************************
'
Function PRECALCU(CI%)
  '
  COU = 0
  RFG$ = RECFILT$("LICOSCAL", 1, MKI$(CI%))
  If Len(RFG$) >= 4 Then
    REGB& = CVS(RFG$)
    If REGB& > 0 Then
      Q$ = REGLEIDO$("LICOSCAL", REGB&)
      COU = CVD(Mid$(Q$, 97, 8))
    End If
  End If
  '
  PRECALCU = COU
  '
End Function
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

Function COSMATCAL(CI%, MONECO%)
   '
   ' determina el costo calculado de suministros (M.Prima,
   ' Piezas y Trabajos de Terceros en la misma forma que
   ' el programa ANACOS01. Si MONECO%=1 en pesos
   '                       Si MONECO%=2 en dolares
   '
   Static DP%(32767)
   Static FESTRU#, NVI%, NMM%
   '
   SUCOS = 0
   CICON% = CI%
   If CICON% < 1 Or CICON% > NUMAS% Then GoTo 90
   '
   ' IDENTIFICANDO LOS ITEMS SIN EXPLOSION
   '
   FESTRUC# = FileDateTime(LUDAT$ + "ESTRUNUE.DAT")
   If FESTRUC# <> FESTRU# Then
     '
     NVI% = NUINV%
     NMM% = NUMAS%
     '
     FESTRU# = FESTRUC#
     FIN& = ULTREG&("ESTRUNUE")
     For I& = 1 To FIN&
       Call TRACE(20, I&, FIN&)
       R0$ = REGLEIDO$("ESTRUNUE", I&)
       CI1% = CVI(Left$(R0$, 2))
       If CI1% > 0 Then
         If CI1% <= NMM% Then
           If CVS(Mid$(R0$, 5, 4)) > 0.00005 Then
             DP%(CI1%) = 1
           End If
         End If
       End If
     Next I&
   End If
   '
   Call EXPLOMUL(CICON%)
   '
   FIN& = CONTAMUL
   For KKI% = 1 To CONTAMUL
     COSUDOL = COSUNI(CIK%(KKI%))
     If DP%(CIK%(KKI%)) < 1 Then   ' NO TIENE ESTRUCTURA
        '
        MONEI% = MONECOSTO%(CIK%(KKI%))
        COSUN = COSUDOL * TICAMONE(MONEI%) * (1 + PORNACIO(CIK%(KKI%)) / 100)
        SUCOS = SUCOS + CANT(KKI%) * COSUN
        '
      Else
        '
        If COSUDOL > 0 Then
          PHB% = PROVHABI%(CIK%(KKI%))
          If PHB% > 0 Then
            If PROVINT%(PHB%) = 0 Then
              MONEI% = MONECOSTO%(CIK%(KKI%))
              COSUN = COSUDOL * TICAMONE(MONEI%) * (1 + PORNACIO(CIK%(KKI%)) / 100)
              SUCOS = SUCOS + CANT(KKI%) * COSUN
            End If
          End If
        End If
        '
     End If
   Next KKI%
   '
   If MONECO% = 2 Then
      TCM = TICAMONE(2)
      If TCM < 1 Then TCM = 1
      SUCOS = SUCOS / TCM
   End If
   '
90 COSMATCAL = SUCOS
   '
End Function

'
' ***************************************************************************
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
'
Function PROVHABI%(CI%)
'
PROVHABI% = 0        ' FABRICACION PROPIA
If CI% > 0 Then
  If CI% <= NUMAS% Then
    '
    If PHABIT%(CI%) < 0 Then Exit Function
    If PHABIT%(CI%) > 0 Then
      PROVHABI% = PHABIT%(CI%)
      Exit Function
    End If
    '
    REG& = CI% + 1: CIANT% = CI%
    XXX$ = REGLEIDO$("MASTER", REG&)
10  NPRO% = CVI(Mid$(XXX$, 123, 2))
    PHABIT%(CI%) = (-1)
    If NPRO% > 0 And ECOARCH$("ACREDOR", MKI$(NPRO%)) <> "" Then
      PHABIT%(CI%) = NPRO%
      PROVHABI% = NPRO%
    End If
  End If
End If
'
End Function
'
' ***************************************************************************
'
Function PROVINT%(NPRX%)
   '
   Static PVIT%(32767), CPVI%
   '
   If CPVI% < 1 Then
     CPVI% = 1
     For UI& = 0 To 32767
       PVIT%(UI&) = 0
     Next UI&
     '
     For U& = 1 To ULTREG&("TAPROINT")
       X$ = REGLEIDO$("TAPROINT", U&)
       NPPX% = CVI(Left$(X$, 2))
       If NPPX% > 0 Then
         PVIT%(NPPX%) = 1
       End If
     Next U&
     Call CIERRARCH("TAPROINT")
   End If
   '
   PROVINT% = 0
   If NPRX% > 0 Then
     PROVINT% = PVIT%(NPRX%)
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
If CI% = CIANT% Then GoTo 10
'
FECOSTO% = 0
If CI% > 0 Then
   If CI% <= NUMAS% Then
       REG& = CI% + 1: CIANT% = CI%
       XXX$ = REGLEIDO$("MASTER", REG&)
10     FECOSTO% = CVI(Mid$(XXX$, 115, 2))
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
       'NXX% = FILENBR%("MASTER")
       REG& = CILOCAL% + 1
       'Get #NXX%, REG&, REMA$
       REMA$ = REGLEIDO$("MASTER", REG&)
       GRIVA% = Asc(Mid$(REMA$, 3, 1))
   End If
End If
If GRIVA% < 0 Or GRIVA% > 1 Then GRIVA% = 0
'
PORCEIVA = PRIVA(GRIVA%)
'
End Function
'
' ***************************************************************************
'
Function STOMIN(CI%)
'
If CI% = CIANT% Then GoTo 10
'
STOMIN = 0
If CI% > 0 Then
   If CI% <= NUMAS% Then
       REG& = CI% + 1: CIANT% = CI%
       XXX$ = REGLEIDO$("MASTER", REG&)
10     STOMI = CVI(Mid$(XXX$, 105, 2))
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
If CI% = CIANT% Then GoTo 10
'
If CI% > 0 Then
   If CI% <= NUMAS% Then
       REG& = CI% + 1: CIANT% = CI%
       XXX$ = REGLEIDO$("MASTER", REG&)
10     LT = CVI(Mid$(XXX$, 107, 2))
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
If CI% = CIANT% Then GoTo 10
'
If CI% > 0 Then
   If CI% <= NUMAS% Then
       REG& = CI% + 1: CIANT% = CI%
       XXX$ = REGLEIDO$("MASTER", REG&)
10     LR = CVI(Mid$(XXX$, 103, 2))
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
If CI% = CIANT% Then GoTo 10
'
If CI% > 0 Then
   If CI% <= NUMAS% Then
       REG& = CI% + 1: CIANT% = CI%
       XXX$ = REGLEIDO$("MASTER", REG&)
10     TMT% = Asc(Mid$(XXX$, 93, 1))
   End If
End If
'
TIMATE% = TMT%
'
End Function
'
' ***************************************************************************
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

Sub REPUNCODIG()
   '
   Screen.MousePointer = 11
   '
   BFC$ = ""
   NMSX& = NUMAS%
   LOPUCO& = 16 * NMSX&
   BFC$ = String$(LOPUCO&, 0)
   For U& = 1 To NUMAS%
      X$ = REGLEIDO$("MASTER", U& + 1)
      BFD$ = AJUSTI$(Mid$(X$, 77, 16), 16)
      Mid$(BFC$, 16 * U& - 15, 16) = BFD$
   Next U&
   '
   LOPUCO1& = 2048 * (1 + Int(LOPUCO& / 2048))
   BFC1$ = Left$(BFC$ + String$(4096, 0), LOPUCO1&)
   Call SAVEFILE(LUDAT$ + "PUNCODIG.DAT", BFC1$)
   PUNCODIG$ = BFC$
   '
   Screen.MousePointer = 1
   '
End Sub
'
Sub GENECOMAS()
   '
   Static YAGENEC%
   If YAGENEC% > 0 Then Exit Sub
   If HAYACMAS% < 1 Then Exit Sub
   '
   Screen.MousePointer = 11
   '
   FIN& = ULTREG&("MASTER")
   If FIN& < 2 Then
     Screen.MousePointer = 1
     Exit Sub
   End If
   '
   JI& = 1  ' PARA INVERT
   REBLI$ = Space$(16) + MKI$(0)
   JE& = 0  ' PARA ECOMAST
   REBLE$ = Space$(62) + MKI$(0)
   '
   XI$ = REGLEIDO$("INVERT", 1)
   Call REPLA(XI$, String$(16, 0), 1, 16)
   Call GRAREG("INVERT", XI$, 1)
   '
   For U& = 2 To FIN&
     Call TRACE(20, U&, FIN&)
     XM$ = REGLEIDO$("MASTER", U&)
     CII% = U& - 1
     If CVI(Left$(XM$, 2)) <> CII% Then
       Call REPLA(XM$, MKI$(CII%), 1, 2)
       Call GRAREG("MASTER", XM$, U&)
     End If
     '
     CODE$ = Mid$(XM$, 77, 16)
     DEDE$ = Mid$(XM$, 5, 64)
     '
     XI$ = REBLI$
     Call REPLA(XI$, CODE$, 1, 16)
     Call REPLA(XI$, MKI$(CII%), 17, 2)
     JI& = JI& + 1
     Call GRAREG("INVERT", XI$, JI&)
     '
     If Mid$(CODE$, 16, 1) <> Chr$(96) Then
       XE$ = REBLE$
       Call REPLA(XE$, CODE$, 1, 16)
       Call REPLA(XE$, DEDE$, 17, 46)
       Call REPLA(XE$, MKI$(CII%), 63, 2)
       JE& = JE& + 1
       Call GRAREG("ECOMAST", XE$, JE&)
       '
     End If
   Next U&
   '
   NMS% = FIN& - 1
   XM$ = REGLEIDO$("MASTER", 1)
   Call REPLA(XM$, MKI$(NMS%), 1, 2)
   Call GRAREG("MASTER", XM$, 1)
   Call CIERRARCH("MASTER")
   '
   NVI% = JI& - 1
   XI$ = REGLEIDO$("INVERT", 1)
   Call REPLA(XI$, MKI$(NVI%), 17, 2)
   Call GRAREG("INVERT", XI$, 1)
   Call SETULTREG("INVERT", JI&)
   Call CIERRARCH("INVERT")
   Call FILESORT("INVERT", "", 1, 1, "ASC")
   '
   Call SETULTREG("ECOMAST", JE&)
   Call CIERRARCH("ECOMAST")
   Call FILESORT("ECOMAST", "", 1, 1, "ASC")
   '
   Call REPUNCODIG
   YAGENEC% = 1
   Screen.MousePointer = 1
   '
End Sub

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
               If UCase$(Mid$(X$, 62, 2) = "NO") Then DISTO%(U&) = (-1)
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
'
Function STOCKACT(CI%)
   '
   Static NMSX%
   If NMSX% < 1 Then
      NMSX% = NUMAS%
   End If
   '
   CIKI% = CI%
   SALTOTA = 0: SULMOVI = 0
   If CIKI% > 0 Then
      If CIKI% <= NMSX% Then
         '
         RGG$ = RECFILT$("MOVISTO", 2, MKI$(CIKI%))
         If Len(RGG$) >= 4 Then
            UREMOVI& = CVS(Mid$(RGG$, Len(RGG$) - 3))
            MVJ$ = REGLEIDO$("MOVISTO", UREMOVI&)
            SULMOVI = CVS(Mid$(MVJ$, 118, 4))
         End If
         '
         RESALST& = CIKI% + 1
         SSTX$ = REGLEIDO$("SALDEPOS", RESALST&)
         For KKI% = 1 To 62
            SALTOTA = SALTOTA + CVS(Mid$(SSTX$, 4 * KKI% - 3, 4))
         Next KKI%
         UREMOSA& = CVS(Mid$(SSTX$, 249, 4))
         '
         If Abs(SALTOTA - SULMOVI) >= 0.05 Or UREMOVI& <> UREMOSA& Then
            Call VESARRAS(CIKI%)
            ' DESPUES DE RECUPERAR LOS SALDOS CON VESARRAS NO LO MANDO A VERIFICAR
            ' NUEVAMENTE LOS SALDOS CRUZADOS POR RIESGO DE COLGADA - E.P.B. 1/05/2007
            SSTX$ = REGLEIDO$("SALDEPOS", RESALST&)
            For KKI% = 1 To 62
               SALTOTA = SALTOTA + CVS(Mid$(SSTX$, 4 * KKI% - 3, 4))
            Next KKI%
         End If
         '
      End If
   End If
   '
90 STOCKACT = SALTOTA
   '
End Function
'
Function STOCKACT_VIE(CI%)     ' DEHABILITADA DESDE 01 MAYO 2007
   '
   ' NO VA MOVISVIE
   Static CTX%
   If NUSU% < 1 Then Call CARDEPOS
   '
   If CTX% < 1 Then
     Dim CANACU(64), YAREVIS%()
     URE& = ULTREG&("MOVISTO")
     NMSX& = NUMAS%
     HOII% = HOY(HOS$)
     ReDim Preserve YAREVIS%(NMSX&)
     For K& = 0 To NMSX&: YAREVIS%(K&) = 0: Next K&
     HUBORECU% = 0
   End If
   '
   CIKI% = CI%
   If CIKI% > 0 Then
     If CIKI% <= NMSX& Then
       If YAREVIS%(CIKI%) < 1 Then
         For KI% = 0 To 64: CANACU(KI%) = 0: Next KI%
         '
         RGG$ = RECFILT$("MOVISTO", 2, MKI$(CIKI%))
         If Len(RGG$) < 4 Then GoTo 40 ' NO HAY MOVIMIENTOS
         '
         UREMOVI& = CVS(Mid$(RGG$, Len(RGG$) - 3))
         MVJ$ = REGLEIDO$("MOVISTO", UREMOVI&)
         If CVI(Mid$(MVJ$, 127, 2)) > 0 Then
            YAREVIS%(CIKI%) = 1
            GoTo 60
         End If
         '
         ' RECUPERAR SALDOS DE STOCK DEL ITEM CIKI%
         For KJ& = 1 To Len(RGG$) Step 4
           REMOVJ& = CVS(Mid$(RGG$, KJ&, 4))
           If REMOVJ& > 0 Then
             If REMOVJ& <= URE& Then
               MVJ$ = REGLEIDO$("MOVISTO", REMOVJ&)
               CANTU = (Int(10 * (CVD(Mid$(MVJ$, 88, 8)) - CVD(Mid$(MVJ$, 96, 8))) + 0.5)) / 10
               DEP% = Asc(Mid$(MVJ$, 83, 1))
               '
               COLUX% = 1
               For KI% = 1 To NUSU%
                 If DEP% = COSU%(KI%) Then
                   COLUX% = KI%
                   Exit For
                 End If
               Next KI%
               '
               CANACU(COLUX%) = CANACU(COLUX%) + CANTU
               '
             End If
           End If
         Next KJ&
         '
         CTSX% = 0
         For UKL& = ULTREG&("SALSTOCK") + 1 To NMSX& + 1
           Call GRAREG("SALSTOCK", String$(39, 0) + Chr$(1), UKL&)
           CTSX% = 1
         Next UKL&
         For UKL& = ULTREG&("SALDEPOS") + 1 To FIN&
           Call GRAREG("SALDEPOS", String$(255, 0) + Chr$(1), UKL&)
           CTSX% = 1
         Next UKL&
         If CTSX% > 0 Then
           Call CIERRARCH("SALSTOCK")
           Call CIERRARCH("SALDEPOS")
         End If
         '
40       RESALST& = CIKI% + 1
         SSTX$ = REGLEIDO$("SALDEPOS", RESALST&)
         For KI% = 1 To 62
           If Abs(CVS(Mid$(SSTX$, 4 * KI% - 3, 4)) - CANACU(KI%)) > 0.005 Then
             SSTANT = 0
             For KKI% = 1 To 62
               SSTANT = SSTANT + CVS(Mid$(SSTX$, 4 * KKI% - 3, 4))
             Next KKI%
             '
             Call REPLA(SSTX$, String$(248, 0), 1, 248)
             SSTACT = 0
             For KII% = 1 To 62
               Call REPLA(SSTX$, MKS$(CANACU(KII%)), 4 * KII% - 3, 4)
               SSTACT = SSTACT + CANACU(KKI%)
             Next KII%
             Call GRAREG("SALDEPOS", SSTX$, RESALST&)
             Call GRAREG("SALSTOCK", Left$(SSTX$, 40), RESALST&)
             Call CIERRARCH("SALDEPOS")
             Call CIERRARCH("SALSTOCK")
             '
'             ' AQUI GRABAR RECUPERACION DE SALDO DE STOCK
'             RBLACT$ = REGBLAN$("RECUSSTK")
'             Call REPLA(RBLACT$, MKI$(HOII%), 1, 2)
'             Call REPLA(RBLACT$, MKI$(CIKI%), 3, 2)
'             Call REPLA(RBLACT$, MKS$(SSTANT), 5, 4)
'             Call REPLA(RBLACT$, MKS$(SSTACT), 9, 4)
'             Call REPLA(RBLACT$, MKS$(SSTACT - SSTANT), 13, 4)
'             Call REPLA(RBLACT$, UCase$(App.EXEName), 17, 8)
'             Call REGAPP("RECUSSTK", RBLACT$)
'             Call CIERRARCH("RECUSSTK")
'             '
             Exit For
           End If
         Next KI%
         '
         For KJ& = 1 To Len(RGG$) Step 4
           REMOVJ& = CVS(Mid$(RGG$, KJ&, 4))
           If REMOVJ& > 0 Then
             If REMOVJ& <= URE& Then
               MVJ$ = REGLEIDO$("MOVISTO", REMOVJ&)
               If CVI(Mid$(MVJ$, 127, 2)) < 1 Then
                 Call REPLA(MVJ$, MKI$(HOII%), 127, 2)
                 Call GRAREG("MOVISTO", MVJ$, REMOVJ&)
                 HUBORECU% = 1
               End If
             End If
           End If
         Next KJ&
         '
         YAREVIS%(CIKI%) = 1
         '
       End If
     End If
   End If
   '
   If HUBORECU% > 0 Then
     Call CIERRARCH("MOVISTO")
     HUBORECU% = 0
   End If
   '
60 ST = 0
   If CI% > 0 Then
     If CI% <= NUMAS% Then
       REG& = CI% + 1
       XXZ$ = REGLEIDO$("SALDEPOS", REG&)
       For U% = 1 To 4 * NUSU% Step 4
         ST = ST + CVS(Mid$(XXZ$, U%, 4)) * (1 + DISTO%((U% + 3) / 4))
       Next U%
     End If
   End If
   '
   STOCKACT_VIE = ST
   '
End Function
'
Function STODEPFECH(CIKI%, DEPX%, FECHI%)
   '
   Static NMSX%
   '
   If NMSX% < 1 Then NMSX% = NUMAS%
   '
   CANACU = 0
   If CIKI% > 0 Then
      If CIKI% <= NMSX% Then
         '
         URE& = ULTREG&("MOVISTO")
         RGM$ = RECFILT$("MOVISTO", 2, MKI$(CIKI%))
         '
         For K& = 1 To Len(RGM$) Step 4
            REG& = CVS(Mid$(RGM$, K&, 4))
            If REG& <= URE& Then
               MVS$ = REGLEIDO$("MOVISTO", REG&)
               If CVI(Left$(MVS$, 2)) <= FECHI% Then
                 If Asc(Mid$(MVS$, 83, 1)) = DEPX% Or DEPX% = 0 Then
                   CANTI# = (Int(10 * (CVD(Mid$(MVS$, 88, 8)) - CVD(Mid$(MVS$, 96, 8))) + 0.5)) / 10
                   CANACU = CANACU + CANTI#
                 End If
               End If
            End If
         Next K&
         '
      End If
   End If
   '
90 STODEPFECH = CANACU
   '
End Function

Function STOCKACT_NVIG(CI%)    ' NO VIGENTE DESDE 01/02/07
   '
   ' NO VA MOVISVIE
   Static CTX%
   If NUSU% < 1 Then Call CARDEPOS
   '
   If CTX% < 1 Then
     Dim CANACU(64), YAREVIS%()
     URE& = ULTREG&("MOVISTO")
     NMSX& = NUMAS%
     HOII% = HOY(HOS$)
     ReDim Preserve YAREVIS%(NMSX&)
     For K& = 0 To NMSX&: YAREVIS%(K&) = 0: Next K&
     HUBORECU% = 0
     '
     RGM$ = RECFILT$("MOVISTO", 21, MKI$(0))
     PRIMONOCO& = 0
     For K& = 1 To Len(RGM$) Step 4
       REMOVI& = CVS(Mid$(RGM$, K&, 4))
       If K& = 1 Then PRIMONOCO& = REMOVI&
       If REMOVI& > 0 Then
         If REMOVI& <= URE& Then
           MVS$ = REGLEIDO$("MOVISTO", REMOVI&)
           CIKI% = CVI(Mid$(MVS$, 3, 2))
           '
           If CIKI% < 1 Then                             ' MATERIAL LIBRE
             If CVI(Mid$(MVS$, 127, 2)) < 1 Then
               Call REPLA(MVS$, MKI$(HOII%), 127, 2)
               Call GRAREG("MOVISTO", MVS$, REMOVI&)
               HUBORECU% = 1
             End If
           End If
           '
           If CIKI% > 0 Then
             If CIKI% <= NMSX& Then
               If YAREVIS%(CIKI%) < 1 Then
                 For KI% = 0 To 64: CANACU(KI%) = 0: Next KI%
                 '
                 RGG$ = RECFILT$("MOVISTO", 2, MKI$(CIKI%))
                 For KJ& = 1 To Len(RGG$) Step 4
                   REMOVJ& = CVS(Mid$(RGG$, KJ&, 4))
                   If REMOVJ& > 0 Then
                     If REMOVJ& <= URE& Then
                       MVJ$ = REGLEIDO$("MOVISTO", REMOVJ&)
                       CANTU = (Int(10 * (CVD(Mid$(MVJ$, 88, 8)) - CVD(Mid$(MVJ$, 96, 8))) + 0.5)) / 10
                       DEP% = Asc(Mid$(MVJ$, 83, 1))
                       '
                       COLUX% = 1
                       For KI% = 1 To NUSU%
                         If DEP% = COSU%(KI%) Then
                           COLUX% = KI%
                           Exit For
                         End If
                       Next KI%
                       '
                       CANACU(COLUX%) = CANACU(COLUX%) + CANTU
                       '
                     End If
                   End If
                 Next KJ&
                 '
                 CTSX% = 0
                 For UKL& = ULTREG&("SALSTOCK") + 1 To NMSX& + 1
                   Call GRAREG("SALSTOCK", String$(39, 0) + Chr$(1), UKL&)
                   CTSX% = 1
                 Next UKL&
                 For UKL& = ULTREG&("SALDEPOS") + 1 To FIN&
                   Call GRAREG("SALDEPOS", String$(255, 0) + Chr$(1), UKL&)
                   CTSX% = 1
                 Next UKL&
                 If CTSX% > 0 Then
                   Call CIERRARCH("SALSTOCK")
                   Call CIERRARCH("SALDEPOS")
                 End If
                 '
                 RESALST& = CIKI% + 1
                 SSTX$ = REGLEIDO$("SALDEPOS", RESALST&)
                 For KI% = 1 To 62
                   If Abs(CVS(Mid$(SSTX$, 4 * KI% - 3, 4)) - CANACU(KI%)) > 0.005 Then
                     SSTANT = 0
                     For KKI% = 1 To 62
                       SSTANT = SSTANT + CVS(Mid$(SSTX$, 4 * KKI% - 3, 4))
                     Next KKI%
                     '
                     Call REPLA(SSTX$, String$(248, 0), 1, 248)
                     SSTACT = 0
                     For KII% = 1 To 62
                       Call REPLA(SSTX$, MKS$(CANACU(KII%)), 4 * KII% - 3, 4)
                       SSTACT = SSTACT + CANACU(KKI%)
                     Next KII%
                     Call GRAREG("SALDEPOS", SSTX$, RESALST&)
                     Call GRAREG("SALSTOCK", Left$(SSTX$, 40), RESALST&)
                     Call CIERRARCH("SALDEPOS")
                     Call CIERRARCH("SALSTOCK")
                     '
                     ' AQUI GRABAR RECUPERACION DE SALDO DE STOCK
                     RBLACT$ = REGBLAN$("RECUSSTK")
                     Call REPLA(RBLACT$, MKI$(HOII%), 1, 2)
                     Call REPLA(RBLACT$, MKI$(CIKI%), 3, 2)
                     Call REPLA(RBLACT$, MKS$(SSTANT), 5, 4)
                     Call REPLA(RBLACT$, MKS$(SSTACT), 9, 4)
                     Call REPLA(RBLACT$, MKS$(SSTACT - SSTANT), 13, 4)
                     Call REPLA(RBLACT$, UCase$(App.EXEName), 17, 8)
                     Call REGAPP("RECUSSTK", RBLACT$)
                     Call CIERRARCH("RECUSSTK")
                     '
                     Exit For
                   End If
                 Next KI%
                 '
                 If PRIMONOCO& > 0 Then
                   RGG$ = RECFILT$("MOVISTO", 2, MKI$(CIKI%))
                   For KJ& = 1 To Len(RGG$) Step 4
                     REMOVJ& = CVS(Mid$(RGG$, KJ&, 4))
                     If REMOVJ& >= PRIMONOCO& Then
                       If REMOVJ& <= URE& Then
                         MVJ$ = REGLEIDO$("MOVISTO", REMOVJ&)
                         If CVI(Mid$(MVJ$, 127, 2)) < 1 Then
                           Call REPLA(MVJ$, MKI$(HOII%), 127, 2)
                           Call GRAREG("MOVISTO", MVJ$, REMOVJ&)
                           HUBORECU% = 1
                         End If
                       End If
                     End If
                   Next KJ&
                 End If
                 '
                 YAREVIS%(CIKI%) = 1
                 '
               End If
             End If
           End If
         End If
       End If
     Next K&
     '
     If HUBORECU% > 0 Then
       Call CIERRARCH("MOVISTO")
     End If
     '
     CTX% = 1
     '
   End If
   '
   '
   ST = 0
   If CI% > 0 Then
     If CI% <= NUMAS% Then
       REG& = CI% + 1
       XXZ$ = REGLEIDO$("SALDEPOS", REG&)
       For U% = 1 To 4 * NUSU% Step 4
         ST = ST + CVS(Mid$(XXZ$, U%, 4)) * (1 + DISTO%((U% + 3) / 4))
       Next U%
     End If
   End If
   '
   STOCKACT_NVIG = ST
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
Function STOCKRES(CI%)
    '
10  ASII = 0
    RGXX$ = RECFILT$("RESERVA", 1, MKI$(CI%))
    '
    For I& = 1 To Len(RGXX$) Step 4
        REG& = CVS(Mid$(RGXX$, I&, 4))
        X$ = REGLEIDO$("RESERVA", REG&)
        CASIG = CVD(Mid$(X$, 19, 8)) - CVD(Mid$(X$, 27, 8))
        If CASIG > 0 Then
          ASII = ASII + CASIG
        End If
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
        '
        If CVI(Left$(X$, 2)) <> CI% Then
           X$ = REGLEIDO$("RESERVA", 1)
           Call GRAREG("RESERVA", X$, 1)
           Call CIERRARCH("RESERVA")
           GoTo 10
        End If
        '
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
Function STOCKDIS(CI%)
    '
    STOCKDIS = STOCKACT(CI%) - STOCKASI(CI%) - STOCKRES(CI%)
    '
End Function
'
Function STOPREP(CI%)
   STP = 0
   RFF$ = RECFILT$("PEDIPREP", 14, MKI$(CI%))
   For KU& = 1 To Len(RFF$) Step 4
     REPRE& = CVS(Mid$(RFF$, KU&, 4))
     XX$ = REGLEIDO$("PEDIPREP", REPRE&)
     '
     NROPED& = CVS(Mid$(XX$, 17, 4))
     CANPREP = CVS(Mid$(XX$, 91, 4))
     '
     CAPEPE = 0
     RFG$ = RECFILT$("DESPAPED", 1, MKS$(NROPED&))
     For KX& = 1 To Len(RFG$) Step 4
       REPEDA& = CVS(Mid$(RFG$, KX&, 4))
       X$ = REGLEIDO$("DESPAPED", REPEDA&)
       If CVI(Mid$(X$, 33, 2)) = CI% Then
         If Asc(Mid$(X$, 27, 1)) < 2 Then
           If Mid$(X$, 51, 1) <> "R" Then ' no es reparacion
             CAPEN = CVS(Mid$(X$, 75, 4))
             If CAPEN < 0 Then CAPEN = 0
             CAPEPE = CAPEPE + CAPEN
           End If
         End If
       End If
     Next KX&
     If CAPEPE < CANPREP Then CANPREP = CAPEPE
     '
     STP = STP + CANPREP
   Next KU&
   STOPREP = STP
End Function
'
Function STOCKDISNEVEN(CI%)
    '
    'AAA = DEPOEXPE%
    'BBB = STODEPOS(CI%, DEPOEXPE%)
    'CCC = STOCKASI(CI%)
    'DDD = STOCKRES(CI%)
    'EEE = PEDNOASI(CI%)
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
   If STORF < 0 Then STORF = 0
   STORFAB = STORF
   '
End Function

Function STOCOPEN(CI%)
     '
     CAPEN = 0
18   RFF$ = RECFILT$("OCOMDETA", 5, MKI$(CI%))
     For UK& = 1 To Len(RFF$) Step 4
        U& = CVS(Mid$(RFF$, UK&, 4))
        X$ = REGLEIDO$("OCOMDETA", U&)
        CIIX% = CVI(Mid$(X$, 9, 2))
        STAT% = Asc(Mid$(X$, 124, 1))
        If STAT% < 8 Then      ' no esta anulada NI CUMPLIDA
          NUORSE& = CVS(Mid$(X$, 117, 4))
          If NUORSE& < 1 Then
            COEFCAN = CVS(Mid$(X$, 63, 4))
            If COEFCAN < 0.005 Then COEFCAN = 1
            CACOM = CVS(Mid$(X$, 55, 4)) * COEFCAN
            CAREC = CVS(Mid$(X$, 99, 4)) * COEFCAN
            CAPEN = CACOM - CAREC
          End If
        End If
     Next UK&
     '
     If CAPEN < 0 Then CAPEN = 0
     STOCOPEN = CAPEN

End Function

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
Function PRELISTA#(APREC$, CIIXXI%)
   '
   Static APRECA$, UREPRE&, FEHOANT#
   '
   FEHOPRE# = FileDateTime(LUDAT$ + APREC$ + ".DAT")
   If APREC$ <> APRECA$ Or FEHOPRE# <> FEHOANT# Then
     APRECA$ = APREC$
     FEHOANT# = FEHOPRE#
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
           If REPRE& <= ULTREG&(APREC$) Then '  UREPRE& Then
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
   If CONVERPRECIO% = 1 Then
     NLIS% = Val(Mid$(APREC$, 6, 3))
     PRELISTA# = PRELISTA# * TICAMONE(MONELISTA%(NLIS%)) / TICAMONE(MONEMI%)
   End If
   '
End Function
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
Function DEPOEXPE%()
    '
    Static CTXX%, DPXXP%
    If CTXX% = 0 Then
      CTXX% = 1
      DPXXP% = 0
      DPXX% = Val(CONTROL$("", "DEPOEXPE"))
      If DPXX% > 0 Then
        If DPXX% < 99 Then
          If ECOARCH$("TADEPOSI", Chr$(DPXX%)) <> "" Then
            DPXXP% = DPXX%
          End If
        End If
      End If
    End If
    '
    DEPOEXPE% = DPXXP%
    '
End Function
'
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

Sub PRIDOGRAF(CIXI%, IGRA%)
   '
   Static URE&, RUTARGRA$
   '
   If IGRA% < 0 Or IGRA% > 3 Then
      Exit Sub
   End If
   Screen.MousePointer = 11
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
   PRIFORGEN$ = Mid$(TRIM$(CONTROL$("", "FORGRAFI")), 2)
   CAIMAGEN% = 0
   '
   If PRIFORGEN$ <> "" Then
     If CARDOGRAF%(CIXI%, IGRA%, AGRAFIC$) = 1 Then
       HIMAGEN = FORMALTER.Image3.Picture.Width
       VIMAGEN = FORMALTER.Image3.Picture.Height
       If HIMAGEN > VIMAGEN Then
            PRIFORM$ = "H" + PRIFORGEN$
          Else
            PRIFORM$ = "V" + PRIFORGEN$
       End If
       '
       If EXISTE%(LUDAT$ + PRIFORM$ + ".FRM") Then
         If CIXI% > 0 Then
           If CIXI% <= NUMAS% Then
             '
             CAPARDOC% = 0
             CACOLTAB1% = 0
             CAITAB1% = 0
             CACOLTAB2% = 0
             CAITAB2% = 0
             '
             CODPARAM$(1) = "DES-MAT": DOCPARAM$(1) = TRIM$(CODEXT$(CIXI%)) + " - " + DESCRIT0$(CIXI%)
             CAPARDOC% = 1
             '
             COPLANO$ = AGRAFIC$
             If COPLANO$ <> "" Then
               If EXISTE%(COPLANO$) > 0 Then
                 CAIMAGEN% = 1
                 CODIMAGEN$(1) = "GRAFIPIE"
                 ARCHIMAGEN$(1) = COPLANO$
               End If
             End If
           End If
         End If
       End If
       Call PRINTDOC("@" + PRIFORM$)
     End If
   End If
   Screen.MousePointer = 1
   '
End Sub
'
Function CARDOGRAF%(CIXI%, IGRA%, AGRAFIC$)
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
   LPX% = 0: AGRAFIC$ = ""
   If IGRA% < 0 Or IGRA% > 3 Then
      Exit Function
   End If
   '
   ' BUSCAR ARCHIVO DEL PLANO DE CIIX%
   '
   APLANO$ = "": Index = 0
   If CIXI% > 0 Then
     If CIXI% <= NUMAS% Then
       Screen.MousePointer = 11
       RNC$ = RECFILT$("INDIPLAN", 1, MKI$(CIXI%))
       If Len(RNC$) >= 4 Then
         REPLANO& = CVS(RNC$)
         If REPLANO& > 0 Then
           If REPLANO& <= URE& Then
             RPLA$ = REGLEIDO$("INDIPLAN", REPLANO&)
             APLANO$ = TRIM$(Mid$(RPLA$, 65 + 48 * IGRA%, 48))
             If APLANO$ <> "" Then
               FORMALTER.Image3.Picture = LoadPicture("")
               COPLANO$ = RUTARGRA$ + APLANO$
               If COPLANO$ <> "" Then
                 If EXISTE%(COPLANO$) > 0 Then
                   FORMALTER.Image3.Picture = LoadPicture(COPLANO$)
                   LPX% = 1
                   AGRAFIC$ = COPLANO$
                 End If
               End If
             End If
           End If
         End If
       End If
     End If
   End If
   '
   CARDOGRAF% = LPX%
   '
End Function

Function DEGRUVEN$(CL%)
   DX$ = "Varios"
   If CL% > 0 Then DX$ = ECOARCH$("GRUCOVEN", MKI$(CL%))
   If DX$ = "" Then Beep: CL% = 0: DX$ = "Varios"
   DEGRUVEN$ = DX$
End Function
'
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
Function TIESTAN(CIXI%, NOPE%, CANTID#, MODO%)
   '
   ' MODO% = 0 ==> incluye tiempo de preparación
   ' MODO% > 0 ==> no considera tiempo de preparación
   '
   ' con NOPE% = 0 da el tiempo total de todas las operaciones
   '
   TSTA = 0
   LOTESTA = LOTESTAN(CIXI%): If LOTESTA < 1 Then LOTESTA = 1
   RNC$ = RECFILT$("SECOPER", 1, MKI$(CIXI%))
   URN& = ULTREG&("SECOPER")
   '
   For U& = 1 To Len(RNC$) Step 4
     RENOTA& = CVS(Mid$(RNC$, U&, 4))
     If RENOTA& > 0 Then
       If RENOTA& <= URN& Then
         TTXX$ = REGLEIDO$("SECOPER", RENOTA&)
         If NOPE% = 0 Or CVI(Mid$(TTXX$, 3, 2)) = NOPE% Then
           '
           'CALOTES = 1
           CALOTES = CANTID# / LOTESTA
           CAOPS = CVS(Mid$(TTXX$, 83, 4))
           If CAOPS < 0.01 Then CAOPS = 1
           '
           MINPREP = 60 * CVS(Mid$(TTXX$, 59, 4)) * CALOTES
           If MODO% > 0 Then MINPREP = 0
           '
           MINFIJO = 60 * CVS(Mid$(TTXX$, 63, 4)) * CALOTES
           MINVARI = CANTID# * CVS(Mid$(TTXX$, 67, 4))
           MINDEMO = (MINFIJO + MINVARI) * CVS(Mid$(TTXX$, 75, 4)) / 100
           MINLIMP = 60 * CVS(Mid$(TTXX$, 79, 4)) * CALOTES
           MINLOTE = MINPREP + MINFIJO + MINVARI + MINDEMO + MINLIMP
           '
           TSTA = TSTA + CAOPS * MINLOTE
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
Function COSTAN(CIXI%, NOPE%, CANTID#, MODO%)
   '
   ' incluye tiempo de preparación
   ' MODO% = 1 --> Solo Mano de Obra
   ' MODO% = 2 --> Solo Tiempo de maquina
   ' MODO% = 0 --> Ambos Juntos
   ' con NOPE% = 0 da el COSTO total de todas las operaciones
   '
   COSMOD = 0: COSMAQ = 0
   RNC$ = RECFILT$("SECOPER", 1, MKI$(CIXI%))
   URN& = ULTREG&("SECOPER")
   '
   For U& = 1 To Len(RNC$) Step 4
     RENOTA& = CVS(Mid$(RNC$, U&, 4))
     If RENOTA& > 0 Then
       If RENOTA& <= URN& Then
         TTXX$ = REGLEIDO$("SECOPER", RENOTA&)
         If NOPE% = 0 Or CVI(Mid$(TTXX$, 3, 2)) = NOPE% Then
           '
           CALOTES = 1
           CAOPS = CVS(Mid$(TTXX$, 83, 4))
           If CAOPS < 0.01 Then CAOPS = 1
           CATOP% = Asc(Mid$(TTXX$, 87, 1))
           EQUIPO$ = TRIM$(Mid$(TTXX$, 53, 6))
           '
           MINPREP = 60 * CVS(Mid$(TTXX$, 59, 4)) * CALOTES
           If MODO% > 0 Then MINPREP = 0
           '
           MINFIJO = 60 * CVS(Mid$(TTXX$, 63, 4)) * CALOTES
           MINVARI = CANTID# * CVS(Mid$(TTXX$, 67, 4))
           MINDEMO = (MINFIJO + MINVARI) * CVS(Mid$(TTXX$, 75, 4)) / 100
           MINLIMP = 60 * CVS(Mid$(TTXX$, 79, 4)) * CALOTES
           MINLOTE = MINPREP + MINFIJO + MINVARI + MINDEMO + MINLIMP
           '
           HORAPROD = MINLOTE / 60
           COSMOD = COSMOD + CAOPS * HORAPROD * VALHORA(CATOP%)
           COSMAQ = COSMAQ + HORAPROD * CUHORMAQ(EQUIPO$)
           '
         End If
       End If
     End If
   Next U&
   '
   If MODO% = 1 Then
       COSTAN = COSMOD
     ElseIf MODO% = 2 Then
       COSTAN = COSMAQ
     Else
       COSTAN = COSMOD + COSMAQ
   End If
   '
End Function
'
Function COSOPTER(CIXI%, NOPE%, CANTID#, MODO%)
   '
   ' MODO% = 1 --> EXCLUYE TRAT TERMICOS
   ' MODO% = 2 --> Solo TRAT TERMICOS
   ' MODO% = 0 --> Ambos Juntos
   ' con NOPE% = 0 da el COSTO total de todas las operaciones
   '
   RNC$ = RECFILT$("COSOPTER", 1, MKI$(CIXI%))
   URN& = ULTREG&("COSOPTER")
   '
   For U& = 1 To Len(RNC$) Step 4
     RENOTA& = CVS(Mid$(RNC$, U&, 4))
     If RENOTA& > 0 Then
       If RENOTA& <= URN& Then
         TTXX$ = REGLEIDO$("COSOPTER", RENOTA&)
         If NOPE% = 0 Or CVI(Mid$(TTXX$, 3, 2)) = NOPE% Then
           '
           NOPEI% = CVI(Mid$(TTXX$, 3, 2))
           PREUN = CVD(Mid$(TTXX$, 7, 8))
           MONEM% = CVI(Mid$(TTXX$, 15, 2))
           If TIESTAN(CIXI%, NOPEI%, 1, 0) < 1 Then
             DESOP$ = UCase$(DESCRIOP$(CIXI%, NOPEI%))
             If InStr(DESOP$, "TERM") > 0 Then
                 COSTT = COSTT + PREUN * TICAMONE(MONEM%)
               Else
                 COS3S = COS3S + PREUN * TICAMONE(MONEM%)
             End If
           End If
         End If
       End If
     End If
   Next U&
   '
   If MODO% = 1 Then
       COSOPTER = COS3S
     ElseIf MODO% = 2 Then
       COSOPTER = COSTT
     Else
       COSOPTER = COS3S + COSTT
   End If
   '
End Function
'
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

Function IAJUINF#(FEBAS%, FEACT%)
   '
   Static FIDATA#, AJUIN$
   Static CAPER%, DESPE%(512), HASPE%(512), IAJU#(512)
   '
   If AJUIN$ = "" Or FIDATIM#("FLEXCRL") <> FIDATA# Then
     FIDATA# = FIDATIM#("FLEXCRL")
     AJUIN$ = CONTROL$("ANACOSTO", "AJUINFLA")
     If AJUIN$ <> "SI" Then AJUIN$ = "NO"
   End If
   '
   If AJUIN$ <> "SI" Then
     IAJC# = 1
     GoTo 99
   End If
   '
   If CAPER% < 1 Then
     CAPER% = ULTREG&("INDIAJU")
     For U& = 1 To CAPER%
       XX$ = REGLEIDO$("INDIAJU", U&)
       DESPE%(U&) = CVI(Mid$(XX$, 17, 2))
       HASPE%(U&) = CVI(Mid$(XX$, 19, 2))
       IAJU#(U&) = CVD(Mid$(XX$, 21, 8))
     Next U&
     If CAPER% < 1 Then
       CAPER% = 1
       DESPE%(1) = (-32767)
       HASPE%(1) = 32767
       IAJU#(1) = 1
     End If
   End If
   '
10 For KKI% = 1 To CAPER%
     If HASPE%(KKI%) >= FEBAS% Then
       IAJ1# = IAJU#(KKI%)
       GoTo 20
     End If
   Next KKI%
   IAJ1# = IAJU#(CAPER%)
   '
20 For KKI% = 1 To CAPER%
     If HASPE%(KKI%) >= FEACT% Then
       IAJ2# = IAJU#(KKI%)
       GoTo 30
     End If
   Next KKI%
   IAJ2# = IAJU#(CAPER%)
   '
30 IAJC# = 1
   If IAJ1# >= 1 Then
     IAJC# = IAJ2# / IAJ1#
   End If
   '
99 IAJUINF# = IAJC#
   '
End Function

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

Function POSPLANO$(CIXI%)
   '
   NPLX$ = ""
   RNC$ = RECFILT$("INDIPLAN", 1, MKI$(CIXI%))
   If Len(RNC$) >= 4 Then
     REPLANO& = CVS(RNC$)
     If REPLANO& > 0 Then
       If REPLANO& <= ULTREG&("INDIPLAN") Then
         RPLA$ = REGLEIDO$("INDIPLAN", REPLANO&)
         NPLX$ = TRIM$(Mid$(RPLA$, 25, 8))
       End If
     End If
   End If
   '
   POSPLANO$ = NPLX$
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
   Dim TLOTE$(), CALOTE#()
   MAXLOTS% = 0
   LSX$ = ""
   '
   If CIXI% > 0 Then
     CANLOT% = 0
     RFF$ = RECFILT$("MOVISTO", 2, MKI$(CIXI%))
     For U& = 1 To Len(RFF$) Step 4
       REMO& = CVS(Mid$(RFF$, U&, 4))
       X$ = REGLEIDO$("MOVISTO", REMO&)
       If Mid$(X$, 21, 2) <> "KU" Then
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
            If CANLOT% > MAXLOTS% Then
              MAXLOTS% = MAXLOTS% + 128
              ReDim Preserve TLOTE$(MAXLOTS%), CALOTE#(MAXLOTS%)
            End If
            TTXX$ = AJUSTI$(IDENLO$, 40)
            Call REPLA(TTXX$, FECHATEX$(FEX), 19, 8)
            Call REPLA(TTXX$, CSTRING$(MKD$(ENTRA#), 4, 12, 1, 2), 27, 12)
            TLOTE$(CANLOT%) = TTXX$
            CALOTE#(KKI%) = ENTRA#
            '
          End If
        End If
19   Next U&
     '
     If EXISTE%(LUDAT$ + "MOVISVIE.DAT") > 0 Then
        RGG$ = RECFILT$("MOVISVIE", 2, MKI$(CIXI%))
        For U& = 1 To Len(RGG$) Step 4
          REMO& = CVS(Mid$(RGG$, U&, 4))
          X$ = REGLEIDO$("MOVISVIE", REMO&)
          If Mid$(X$, 21, 2) <> "KU" Then
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
                   GoTo 191
                 End If
               Next KKI%
               FEX = CVI(Left$(X$, 2))
               CANLOT% = CANLOT% + 1
               If CANLOT% > MAXLOTS% Then
                 MAXLOTS% = MAXLOTS% + 128
                 ReDim Preserve TLOTE$(MAXLOTS%), CALOTE#(MAXLOTS%)
               End If
               TTXX$ = AJUSTI$(IDENLO$, 40)
               Call REPLA(TTXX$, FECHATEX$(FEX), 19, 8)
               Call REPLA(TTXX$, CSTRING$(MKD$(ENTRA#), 4, 12, 1, 2), 27, 12)
               TLOTE$(CANLOT%) = TTXX$
               CALOTE#(KKI%) = ENTRA#
               '
             End If
           End If
191     Next U&
     End If
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
     For U& = 1 To Len(RGG$) Step 4
       REMO& = CVS(Mid$(RGG$, U&, 4))
       X$ = REGLEIDO$("MOVISVIE", REMO&)
       SALID# = CVD(Mid$(X$, 96, 8))
       If SALID# >= 0.05 Then
         IDENLO$ = AJUSTI$(Mid$(X$, 5, 16), 16)
         IDENLOX$ = TRIM$(IDENLO$)
         '
135       For KKI% = 1 To CANLOT%
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
           GoTo 135
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

Function GRUCOFAB%(CIJK%)
  '
  ' GRUCOFAB%=1 --> Compra
  ' GRUCOFAB%=2 --> Fabricacion
  '
  Static GRF(32767) As Integer
  '
  If GRF%(CIJK%) < 1 Then
    GRF%(CIJK%) = 1
    Call AREXPLO(CIJK%, 1, 1)
    '
    If CAIT% > 0 Then
      GRF%(CIJK%) = 2
      PBL% = PROVHABI%(CIJK%)
      If PBL% > 0 Then
        If PROVINT%(PBL%) = 0 Then
          GRF%(CIJK%) = 1
        End If
      End If
    End If
  End If
  '
  GRUCOFAB% = GRF%(CIJK%)
  '
End Function

Sub AGRECOD(CODINUE$, DESCRINUE$, TIMA%, UNID$, OKX%)
   '
   Screen.MousePointer = 11
   '
   NMSI% = NUMAS% + 1
   '
   OBX$ = REGBLAN$("MASTER")
   Call REPLA(OBX$, MKI$(NMSI%), 1, 2)
   Call REPLA(OBX$, CODINUE$, 77, 16)
   Call REPLA(OBX$, DESCRINUE$, 5, 64)
   Call REPLA(OBX$, Chr$(TIMA%), 93, 1)
   Call REPLA(OBX$, UNID$, 109, 2)
   '
   PUNCODIG$ = PUNCODIG$ + AJUSTI$(CODINUE$, 16)
   PPCC$ = PUNCODIG$
   JI% = 1
   While Len(PPCC$) > 2048
      PPCC$ = Mid$(PPCC$, 2049)
      JI% = JI% + 1
   Wend
   While Len(PPCC$) < 2048
      PPCC$ = PPCC$ + Chr$(0)
   Wend
   '
   NX% = FILEOPEN%(LUDAT$ + "PUNCODIG.DAT", 0, 2048)
   BUFERCOD$ = PPCC$
   Put #NX%, JI%, BUFERCOD$
   Close #NX%
   '
   REMAS& = NUMAS% + 2
   Call GRAREG("MASTER", OBX$, REMAS&)
   X$ = REGLEIDO$("MASTER", 1&)
   Call REPLA(X$, MKI$(NMSI%), 1, 2)
   Call GRAREG("MASTER", X$, 1)
   Call CIERRARCH("MASTER")
   '
   EL1$ = REGBLAN$("ECOMAST")
   Call REPLA(EL1$, CODINUE$, 1, 16)
   Call REPLA(EL1$, Mid$(OBX$, 5, 64), 17, 46)
   Call REPLA(EL1$, MKI$(NMSI%), 63, 2)
   '
   URECO& = ULTREG&("ECOMAST")
   While URECO& > 0
     XXX$ = REGLEIDO$("ECOMAST", URECO&)
     If TRIM$(XXX$) <> "" Then GoTo 25
     URECO& = URECO& - 1
   Wend
   '
25 LOTOTE& = 64 * URECO&
   Call CIERRARCH("ECOMAST")
   FIECO% = FILEOPEN%(LUDAT$ + "ECOMAST.DAT", 0, 2048)
   Dim EL As String * 2048
   TTXX$ = String$(LOTOTE&, 0)
   '
   For U& = 1 To 1 + LOF(FIECO%) / 2048
      Get #FIECO%, U&, EL$
      On Error Resume Next
      Mid$(TTXX$, 2048 * (U& - 1) + 1, 2048) = EL$
      On Error GoTo 0
   Next U&
   TTXX$ = Left$(TTXX$, LOTOTE&)
   '
   For U& = 1 To Len(TTXX$) Step 64
      If Mid$(TTXX$, U&, 16) > CODINUE$ Then
         TTXX$ = Left$(TTXX$, U& - 1) + EL1$ + Mid$(TTXX$, U&)
         GoTo 60
      End If
   Next U&
   TTXX$ = TTXX$ + EL1$
   '
60 JJ& = 0: LOTEGRA& = Len(TTXX$)
   TTXX$ = TTXX$ + String$(2048, 0)
   For U& = 1 To LOTEGRA& Step 2048
      EL$ = Mid$(TTXX$ + String$(2048, 0), U&, 2048)
      If Len(EL$) < 2048 Then
         EL$ = Left$(EL$ + String$(2048, 0), 2048)
      End If
      JJ& = JJ& + 1
      Put #FIECO%, JJ&, EL$
   Next U&
   Close #FIECO%
   '
   NVII& = 1 + NUINV%
   EL1$ = REGBLAN$("INVERT")
   Call REPLA(EL1$, CODINUE$, 1, 16)
   Call REPLA(EL1$, MKI$(NMSI%), 17, 2)
   LOTOTE& = 18 * ULTREG&("INVERT")
   If LOTOTE& < 18 * NVII& Then LOTOTE& = 18 * NVII&
   FIECO% = FILEOPEN%(LUDAT$ + "INVERT.DAT", 0, 2048)
   TTXX$ = String$(LOTOTE&, 0)
   '
   For U& = 1 To 1 + LOF(FIECO%) / 2048
      Get #FIECO%, U&, EL$
      On Error Resume Next
      Mid$(TTXX$, 2048 * (U& - 1) + 1, 2048) = EL$
      On Error GoTo 0
   Next U&
   TTXX$ = Left$(TTXX$, LOTOTE&)
   While Len(TTXX$) < LOTOTE&: TTXX$ = TTXX$ + Chr$(0): Wend
   '
   NVIJ% = NVII&
   Mid$(TTXX$, 17, 2) = MKI$(Int(NVIJ%))
   For U& = 19 To Len(TTXX$) Step 18
      If Mid$(TTXX$, U&, 16) > CODINUE$ Then
         TTXX$ = Left$(TTXX$, U& - 1) + EL1$ + Mid$(TTXX$, U&)
         GoTo 70
      End If
   Next U&
   TTXX$ = TTXX$ + EL1$
   '
70 TTXX$ = TTXX$
   JJ& = 0
   For U& = 1 To Len(TTXX$) Step 2048
      EL$ = Mid$(TTXX$ + String$(2048, 0), U&, 2048)
      If Len(EL$) < 2048 Then
         EL$ = Left$(EL$ + String$(2048, 0), 2048)
      End If
      JJ& = JJ& + 1
      Put #FIECO%, JJ&, EL$
   Next U&
   Close #FIECO%
   '
   Call CIERRARCH("ECOMAST")
   Call FRESHECHO("MASTER")
   Call BAJALDISCO
   Call REPUNCODIG
   OKX% = 1
   '
99 Screen.MousePointer = 1

End Sub

Function DESCRILA$(CI%)

    DESCRILARGA$ = FILTERGETRECORD("DESCRILA", 1, MKI(CI%))
    DESCRILA$ = TRIM$(Mid$(DESCRILARGA$, 3, 254))

End Function


