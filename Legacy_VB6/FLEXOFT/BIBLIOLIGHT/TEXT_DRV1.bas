Attribute VB_Name = "TEXT_DRV1"
Function AJUSTI$(A$, B%)
    '
    TX$ = Left$(Trim$(A$) + Space$(B%), B%)
    AJUSTI$ = TX$
    '
End Function
'
Function AJUSTD$(A$, B%)
    '
    TX$ = Trim$(A$)
    While Len(TX$) < B%
        TX$ = " " + TX$
    Wend
    AJUSTD$ = TX$
    '
End Function
'
Function AJUCEN$(A$, B%)
    '
    TX$ = Trim$(A$)
    While Len(TX$) < B%
        TX$ = " " + TX$ + " "
    Wend
    AJUCEN$ = Left$(TX$, B%)
    '
End Function
'
Function SUBTEX$(AA$, BB$, CC$)
  '
  STX$ = ""
  AA1$ = UCase$(AA$): BB1$ = UCase$(BB$)
  PPXX% = InStr(AA1$, BB1$)
  If PPXX% > 0 Then
    AA1$ = Mid$(AA$, PPXX% + Len(BB$))
    CC1$ = Trim$(UCase$(CC$))
    PPXX% = InStr(UCase$(AA1$), CC1$)
    If PPXX% < 1 Or CC1$ = "" Then PPXX% = 1 + Len(AA1$)
    STX$ = Left$(AA1$, PPXX% - 1)
  End If
  SUBTEX$ = Trim$(STX$)
  '
End Function
'
Function TRAPASS2$(TTXYZ$)
    '
    ' DEVUELVE UN STRING CON EL PASSWOERD ENCRIPTADO PARA SU GRABACION
    TT1$ = TTXYZ$
GoTo 90      ' POR AHORA ESTÁ DESHABILITADO
    PINI% = InStr(UCase$(TT1$), "USER ID=")
    If PINI% > 0 Then
       PFIN% = InStr(PINI% + 1, UCase$(TT1$), ";")
       If PFIN% < PINI% Then PFIN% = 1 + Len(TT1$)
       IZQUI$ = Left$(TT1$, PINI% + 7)
       DERE$ = Mid$(TT1$, PFIN%)
       RPLX$ = Mid$(TT1$, PINI% + 8, PFIN% - PINI% - 8)
       TT1$ = IZQUI$ + "@#$" + ENCRIPTA$(RPLX$) + DERE$
    End If
    '
    PINI% = InStr(UCase$(TT1$), "PWD=")
    If PINI% > 0 Then
       PFIN% = InStr(PINI% + 1, UCase$(TT1$), ";")
       If PFIN% < PINI% Then PFIN% = 1 + Len(TT1$)
       IZQUI$ = Left$(TT1$, PINI% + 3)
       DERE$ = Mid$(TT1$, PFIN%)
       RPLX$ = Mid$(TT1$, PINI% + 4, PFIN% - PINI% - 4)
       TT1$ = IZQUI$ + "@#$" + ENCRIPTA$(RPLX$) + DERE$
    End If
    '
90  TRAPASS2 = TT1$
    '
End Function
'
Function TRAPASS1$(TTXYZ$)
    '
    ' DESENCRIPTA EL PASSWORD PARA EL STRING DE CONEXION
    TT1$ = TTXYZ$
GoTo 90    ' POR AHORA ESTA DESHABILITADO
    PINI% = InStr(UCase$(TT1$), "USER ID=")
    If PINI% > 0 Then
       If Mid$(TT1$, PINI% + 8, 3) = "@#$" Then
          PFIN% = InStr(PINI% + 1, UCase$(TT1$), ";")
          If PFIN% < PINI% Then PFIN% = 1 + Len(TT1$)
          IZQUI$ = Left$(TT1$, PINI% + 7)
          DERE$ = Mid$(TT1$, PFIN%)
          RPLX$ = Mid$(TT1$, PINI% + 11, PFIN% - PINI% - 11)
          TT1$ = IZQUI$ + DESENCRI$(RPLX$) + DERE$
       End If
    End If
    '
    PINI% = InStr(UCase$(TT1$), "PWD=")
    If PINI% > 0 Then
       If Mid$(TT1$, PINI% + 4, 3) = "@#$" Then
          PFIN% = InStr(PINI% + 1, UCase$(TT1$), ";")
          If PFIN% < PINI% Then PFIN% = 1 + Len(TT1$)
          IZQUI$ = Left$(TT1$, PINI% + 3)
          DERE$ = Mid$(TT1$, PFIN%)
          RPLX$ = Mid$(TT1$, PINI% + 7, PFIN% - PINI% - 7)
          TT1$ = IZQUI$ + DESENCRI$(RPLX$) + DERE$
       End If
    End If
    '
90  TRAPASS1 = TT1$
    '
End Function

Sub REPLA(TX$, SBTX$, PINI%, LON%)
    '
    ST$ = SBTX$
    PX% = PINI%: If PX% < 1 Then PX% = 1
    LX% = LON%: If LX% = 0 Then LX% = Len(ST$)
    '
    While Len(TX$) < PX% + LX% - 1
        TX$ = TX$ + " "
    Wend
    '
    ST$ = Left$(ST$ + Space$(LX%), LX%)
    Mid$(TX$, PX%, LX%) = ST$
    '
End Sub
'
Function REPLACAR$(A$, B$, c$)
    a1$ = A$
    '
    If Len(B$) > 0 Then
10        XL% = InStr(a1$, B$)
          If XL% > 0 Then
              a1$ = Left$(a1$, XL% - 1) + c$ + Mid$(a1$, XL% + Len(B$))
              GoTo 10
          End If
    End If
    '
    REPLACAR$ = a1$
    a1$ = ""
    '
End Function
'
Function ReplaSkey$(A$)
   '
   Static TCORCH$
   '
   If TCORCH$ = "" Then
     TCORCH$ = "+^%~()[]{}"
   End If
   '
   B$ = ""
   For UI% = 1 To Len(A$)
     c$ = Mid$(A$, UI%, 1)
     If InStr(TCORCH$, c$) > 0 Then
       c$ = Chr$(123) + c$ + Chr$(125) ' encierra entre llaves
     End If
     B$ = B$ + c$
   Next UI%
   '
   ReplaSkey$ = B$
   '
End Function
'
Function FORMESS$(A$)
    '
    FORMESS$ = REPLACAR$(A$, "\", Chr$(10) + Chr$(13))
    '
End Function
'
Function ENCRIPTA$(DE$)
    '
    DE1$ = ""
    For i% = 1 To Len(DE$)
        XL% = Asc(Mid$(DE$, i%))
        Y% = (XL% + 23 * i% + 17) Mod 256
        DE1$ = DE1$ + Chr$(Y%)
    Next i%
    ENCRIPTA$ = DE1$
    '
End Function
'
Function DESENCRI$(DE$)
    '
    DE1$ = ""
    For i% = 1 To Len(DE$)
        XL% = Asc(Mid$(DE$, i%))
        Y% = (XL% - 23 * i% - 17)
        While Y% < 0
            Y% = Y% + 256
        Wend
        DE1$ = DE1$ + Chr$(Y%)
    Next i%
    DESENCRI$ = DE1$
    '
End Function
'
Function ENCRINUM$(DE$)
  '
  DE1$ = "."
  VANTE0% = 0: VANTE1% = 1
  For U& = 1 To Len(DE$)
    VK2% = (VANTE0% + VANTE1%) Mod 63
    VK1% = VK2% + Asc(Mid$(DE$, U&, 1))
    DE1$ = DE1$ + Trim$(Str$(VK1%)) + "."
    VANTE0% = VANTE1%: VANTE1% = VK2%
  Next U&
  ENCRINUM$ = DE1$
  '
End Function
'
Function DESENUM$(DE$)
  '
  DE3$ = ""
  DE1$ = Mid$(DE$, 2)
  VANTE0% = 0: VANTE1% = 1
  While DE1$ <> ""
    PPXX% = InStr(DE1$, "."): If PPXX% < 1 Then PPXX% = 1 + Len(DE1$)
    DE2$ = Left$(DE1$, PPXX% - 1): DE1$ = Mid$(DE1$, PPXX% + 1)
    VK2% = (VANTE0% + VANTE1%) Mod 63
    On Error Resume Next
    DE3$ = DE3$ + Chr$(Val(DE2$) - VK2%)
    On Error GoTo 0
    VANTE0% = VANTE1%: VANTE1% = VK2%
  Wend
  DESENUM$ = DE3$
  '
End Function
'
Function RANDSTRING$(LON%)
    '
    RST$ = ""
    Randomize Timer
    For KL% = 1 To Abs(LON%)
10    PPX% = Int(Rnd * 255)
      If LON% < 0 Then
        If PPX% < 48 Or PPX% > 57 Then GoTo 10
      End If
      RST$ = RST$ + Chr$(PPX%)
    Next KL%
    '
    RANDSTRING$ = RST$
    '
End Function
'
Function CONTRACLA$(CLAX$)
   LQ0% = Len(CLAX$): RESPUES$ = ""
   For M% = 1 To LQ0%
      POSI% = ((LQ0% - 1) * M%) Mod LQ0%
      While POSI% < 1: POSI% = POSI% + LQ0%: Wend
      VORI% = Val(Mid$(CLAX$, POSI%, 1))
      VTRA% = (VORI% + 3 * M%) Mod 10
      RESPUES$ = RESPUES$ + Mid$(Str$(VTRA%), 2)
   Next M%
   '
   CONTRACLA$ = RESPUES$
End Function

Function FETEX$(FEXI)
   '
   FEXJ = FEXI
   FETEI& = Int(FEXJ / 13 / 32) + 80
   FEXJ = FEXJ - 32 * 13 * (FETEI& - 80)
   While FETEI& > 99: FETEI& = FETEI& - 100: Wend
   FETEX$ = Trim$(Str$(FETEI&)): While Len(FETEX$) < 2: FETEX$ = "0" + FETEX$: Wend: FETEX$ = "/" + FETEX$
   '
   FETEI& = Int(FEXJ / 32)
   FETEX$ = Trim$(Str$(FETEI&)) + FETEX$: While Len(FETEX$) < 5: FETEX$ = "0" + FETEX$: Wend: FETEX$ = "/" + FETEX$
   '
   FETEI& = FEXJ - 32 * FETEI&
   FETEX$ = Trim$(Str$(FETEI&)) + FETEX$: While Len(FETEX$) < 8: FETEX$ = "0" + FETEX$: Wend
   '
End Function

Function BUSCAVALOR$(TTXX$, ABUSQUE$)
    '
    PPXX% = InStr(TTXX$, ABUSQUE$)
    If PPXX% < 1 Then PPXX% = 1 + Len(TTXX$)
    PPXX% = PPXX% + Len(ABUSQUE$)
    '
    PPYY% = InStr(PPXX% + 1, TTXX$, Chr$(124))
    If PPYY% < 1 Then PPYY% = 1 + Len(TTXX$)
    If PPYY% <= PPXX% Then PPYY% = PPXX% + 1
    '
    BUSCAVALOR$ = Trim$(Mid$(TTXX$, PPXX%, PPYY% - PPXX%))
    '
End Function
'
'Function SAFETEXT$(TXTXT)
'    SAFETEXT$ = ""
'    On Error Resume Next
'    SAFETEXT$ = TXTXT
'    On Error GoTo 0
'End Function
'
