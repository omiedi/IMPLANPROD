VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TeFun"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Public Function SUBTEX$(AA$, BB$, CC$)
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
Public Function REPLACAR$(A$, B$, C$)
    A1$ = A$
    '
    If Len(B$) > 0 Then
10        XL% = InStr(A1$, B$)
          If XL% > 0 Then
              A1$ = Left$(A1$, XL% - 1) + C$ + Mid$(A1$, XL% + Len(B$))
              GoTo 10
          End If
    End If
    '
    REPLACAR$ = A1$
    A1$ = ""
    '
End Function
'
Public Function ReplaSkey$(A$)
   '
   Static TCORCH$
   '
   If TCORCH$ = "" Then
     TCORCH$ = "+^%~()[]{}"
   End If
   '
   B$ = ""
   For UI% = 1 To Len(A$)
     C$ = Mid$(A$, UI%, 1)
     If InStr(TCORCH$, C$) > 0 Then
       C$ = Chr$(123) + C$ + Chr$(125) ' encierra entre llaves
     End If
     B$ = B$ + C$
   Next UI%
   '
   ReplaSkey$ = B$
   '
End Function
'
Public Function FORMESS$(A$)
    '
    FORMESS$ = REPLACAR$(A$, "\", Chr$(10) + Chr$(13))
    '
End Function
'
Public Function ENCRIPTA$(DE$)
    '
    DE1$ = ""
    For I% = 1 To Len(DE$)
        XL% = Asc(Mid$(DE$, I%))
        Y% = (XL% + 23 * I% + 17) Mod 256
        DE1$ = DE1$ + Chr$(Y%)
    Next I%
    ENCRIPTA$ = DE1$
    '
End Function
'
Public Function DESENCRI$(DE$)
    '
    DE1$ = ""
    For I% = 1 To Len(DE$)
        XL% = Asc(Mid$(DE$, I%))
        Y% = (XL% - 23 * I% - 17)
        While Y% < 0
            Y% = Y% + 256
        Wend
        DE1$ = DE1$ + Chr$(Y%)
    Next I%
    DESENCRI$ = DE1$
    '
End Function
'
Public Function ENCRINUM$(DE$)
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
Public Function DESENUM$(DE$)
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
Public Function RANDSTRING$(LON%)
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
Public Function CONTRACLA$(CLAX$)
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

Public Function FETEX$(FEXI)
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



