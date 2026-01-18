Attribute VB_Name = "TEXT_DRV"
Function SUBTEX$(AA$, BB$, CC$)
  '
  STX$ = ""
  AA1$ = UCase$(AA$): BB1$ = UCase$(BB$)
  PPXX% = InStr(AA1$, BB1$)
  If PPXX% > 0 Then
    AA1$ = Mid$(AA$, PPXX% + Len(BB$))
    CC1$ = TRIM$(UCase$(CC$))
    PPXX% = InStr(UCase$(AA1$), CC1$)
    If PPXX% < 1 Or CC1$ = "" Then PPXX% = 1 + Len(AA1$)
    STX$ = Left$(AA1$, PPXX% - 1)
  End If
  SUBTEX$ = LTrim$(RTrim$(STX$))
  '
End Function
'
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
    BUSCAVALOR$ = LTrim(RTrim(Mid$(TTXX$, PPXX%, PPYY% - PPXX%)))
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
    If LX% > 0 Then
      Mid$(TX$, PX%, LX%) = ST$
    End If
    '
End Sub
'
Function REPLACAR$(A$, B$, c$)
    a1$ = A$
    '
    If Len(B$) > 0 Then
10        XL& = InStr(a1$, B$)   ' Reemplazado XL% por XL& para evitar desbordamiento - EPB 16/01/11
          If XL& > 0 Then
              a1$ = Left$(a1$, XL& - 1) + c$ + Mid$(a1$, XL& + Len(B$))
              GoTo 10
          End If
    End If
    '
    REPLACAR$ = a1$
    a1$ = ""
    '
End Function
'
Function SOLONUM$(A$)
    '
    SLNX$ = ""
    For KKII% = 1 To Len(A$)
       XL% = Asc(Mid$(A$, KKII%, 1))
       If XL% >= 48 Then
         If XL% <= 57 Then
           SLNX$ = SLNX$ + Chr$(XL%)
         End If
       End If
    Next KKII%
    '
    SOLONUM$ = SLNX$
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
    DE1$ = DE1$ + TRIM$(str$(VK1%)) + "."
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
    DE3$ = DE3$ + Chr$(XVALO(DE2$) - VK2%)
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
   For l% = 1 To LQ0%
      POSI% = ((LQ0% - 1) * l%) Mod LQ0%
      While POSI% < 1: POSI% = POSI% + LQ0%: Wend
      VORI% = XVALO(Mid$(CLAX$, POSI%, 1))
      VTRA% = (VORI% + 3 * l%) Mod 10
      RESPUES$ = RESPUES$ + Mid$(str$(VTRA%), 2)
   Next l%
   '
   CONTRACLA$ = RESPUES$
End Function

Function FETEX$(FEXI)
   '
   FEXJ = FEXI
   FETEI& = Int(FEXJ / 13 / 32) + 80
   FEXJ = FEXJ - 32 * 13 * (FETEI& - 80)
   While FETEI& > 99: FETEI& = FETEI& - 100: Wend
   FETEX$ = TRIM$(str$(FETEI&)): While Len(FETEX$) < 2: FETEX$ = "0" + FETEX$: Wend: FETEX$ = "/" + FETEX$
   '
   FETEI& = Int(FEXJ / 32)
   FETEX$ = TRIM$(str$(FETEI&)) + FETEX$: While Len(FETEX$) < 5: FETEX$ = "0" + FETEX$: Wend: FETEX$ = "/" + FETEX$
   '
   FETEI& = FEXJ - 32 * FETEI&
   FETEX$ = TRIM$(str$(FETEI&)) + FETEX$: While Len(FETEX$) < 8: FETEX$ = "0" + FETEX$: Wend
   '
End Function

Function TRIM$(AKKA)    ' ANTES  ERA DE A$ PERO DABA ERROR AL PASARLE UN NULL
   '
   ' ESTO ATAJA EL CASO DE NULL
   B$ = ""
   On Error Resume Next
   B$ = LTrim(RTrim(AKKA))
   On Error GoTo 0
   '
   If Len(B$) < 32767 Then
     For U% = 1 To Len(B$)
       AAXC% = Asc(Mid$(B$, U%, 1))
       If AAXC% < 32 Then Mid$(B$, U%, 1) = " "
     Next U%
   End If
   c$ = LTrim$(RTrim$(B$))
   '
   TRIM$ = c$
   '
End Function
'
Function RSPACE$(A$)   ' REEMPLAZA TODOS LOS CARACTERES < 32 POR ESPACIOS
   B$ = A$
   For U% = 1 To Len(B$)
      If Asc(Mid$(B$, U%, 1)) < 32 Then Mid$(B$, U%, 1) = " "
   Next U%
   RSPACE$ = B$
End Function

Function AJUSTI$(A$, B%)
    '
    TX$ = Left$(TRIM$(A$) + Space$(B%), B%)
    AJUSTI$ = TX$
    '
End Function
'
Function AJUSTD$(A$, B%)
    '
    TX$ = TRIM$(A$)
    While Len(TX$) < B%
        TX$ = " " + TX$
    Wend
    If B% < 1 Then TX$ = ""
    While Len(TX$) > B%
        TX$ = Mid$(TX$, 2)
    Wend
    AJUSTD$ = TX$
    '
End Function
'
Function AJUCEN$(A$, B%)
    '
    TX$ = TRIM$(A$)
    While Len(TX$) < B%
        TX$ = " " + TX$ + " "
    Wend
    AJUCEN$ = Left$(TX$, B%)
    '
End Function
'
Function CONVERTEX$(ORI$)
    TX1$ = Space$(2 * Len(ORI$))
    PROPOS& = 1
    For KKU& = 1 To Len(ORI$)
       LETRAI% = Asc(Mid$(ORI$, KKU&))
       ILETRA% = LETRAI% Mod 90
       KLETRA% = LETRAI% \ 90
       TX2$ = Chr$(ILETRA% + 36)
       If KLETRA% = 1 Then
            TX2$ = Chr$(33) + TX2$
          ElseIf KLETRA% = 2 Then
            TX2$ = Chr$(34) + TX2$
       End If
       Mid$(TX1$, PROPOS&, Len(TX2$)) = TX2$
       PROPOS& = PROPOS& + Len(TX2$)
    Next KKU&
    CONVERTEX$ = Left$(TX1$, PROPOS& - 1)
End Function

Function RECUTEX$(TTRA$)
    LOTRA& = Len(TTRA$)
    TX1$ = String$(LOTRA&, 0)
    PROPOS& = 1: KKU& = 1
    '
10  CARAI% = Asc(Mid$(TTRA$, PROPOS&))
    'Call TRACE(20, PROPOS&, LOTRA&)
    If CARAI% = 33 Then
         PROPOS& = PROPOS& + 1
         CARAI% = 90 + Asc(Mid$(TTRA$, PROPOS&))
       ElseIf CARAI% = 34 Then
         PROPOS& = PROPOS& + 1
         CARAI% = 180 + Asc(Mid$(TTRA$, PROPOS&))
    End If
    ILETRA% = CARAI% - 36
    Mid$(TX1$, KKU&, 1) = Chr$(ILETRA%)
    KKU& = KKU& + 1
    PROPOS& = PROPOS& + 1
    If PROPOS& <= LOTRA& Then GoTo 10
    '
    RECUTEX$ = Left$(TX1$, KKU& - 1)
End Function

Sub REPLATEXT(TXTX$, ARGUB$, TXT$)
 
    If InStr(TXT$, ARGUB$) > 0 Then
        Exit Sub
    End If
    '
10  PPXX% = InStr(TXTX$, ARGUB$)
    If PPXX% > 0 Then
        IZQUI$ = Left$(TXTX$, PPXX% - 1)
        LONGI% = Len(ARGUB$)
        DERE$ = Mid$(TXTX$, PPXX% + LONGI%)
        TXTX$ = IZQUI$ + TXT$ + DERE$
        GoTo 10
    End If
        
End Sub

