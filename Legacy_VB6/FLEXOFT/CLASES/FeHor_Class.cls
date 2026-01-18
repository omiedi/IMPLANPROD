VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FeHor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Dim DIAS%(12)
Dim IDS%(12)
Dim DSX$(6)
Dim DIASE As String
'

Public Function FECHATEX$(FEX)
    '
    Dim FS$(3), FEXX(3)
    FEXU = FEX
    '
    If FEX = 0 Then
        F0$ = "        "
        GoTo 435
    End If
    '
    FEXX(3) = Int(FEX / 13 / 32) + 80
    FEX = FEX - 32 * 13 * (FEXX(3) - 80)
    FEXX(2) = Int(FEX / 32)
    FEXX(1) = FEX - 32 * FEXX(2)
    '
    While FEXX(3) > 99
       FEXX(3) = FEXX(3) - 100                                ' rutina a¤o 2000
    Wend
    '
    F0$ = ""
    For II% = 1 To 3
        F$ = Str$(FEXX(II%))
        F$ = Right$(F$, Len(F$) - 1)
        FS$(II%) = Right$("00" + F$, 2)
        F0$ = F0$ + FS$(II%)
        If II% < 3 Then F0$ = F0$ + "/"
    Next II%
    '
435 FECHATEX$ = F0$
    FEX = FEXU
    '
End Function
'

Public Function FECHANUM(FETX$)
    '
    Dim FE1(3)
    AX% = 0
    'FETY$ = REPLACAR$(FETX$, ".", "/")
    'FETX$ = REPLACAR$(FETY$, "-", "/")
    If Mid$(FETX$, 3, 1) <> "/" Then FETX$ = Left$(FETX$, 2) + "/" + Mid$(FETX$, 3)
    If Mid$(FETX$, 6, 1) <> "/" Then FETX$ = Left$(FETX$, 5) + "/" + Mid$(FETX$, 6)
    '
    If ((Len(FETX$) <> 8) Or (Mid$(FETX$, 3, 1) <> "/") Or (Mid$(FETX$, 6, 1) <> "/")) Then
        AX% = 1: GoTo 610
    End If
    '
    For II% = 1 To 3
        FE1(II%) = Val(Mid$(FETX$, 3 * II% - 2, 2))
        If II% <> 3 Then If FE1(II%) < 1 Then GoTo 600
    Next II%
    '
    If FE1(3) < 60 Then FE1(3) = 100 + FE1(3)                 ' rutina a¤o 2000
    '
    If ((FE1(1) > 31) Or (FE1(2) > 12)) Then GoTo 600
    FE = FE1(1) + 32 * FE1(2) + 32 * 13 * (FE1(3) - 80)
    GoTo 610

600 AX% = 2
610 If AX% > 0 Then FE = 0
    If FE < (-32767) Then FE = (-32767)
    If FE > 32767 Then FE = 32767
    '
    FECHANUM = FE
    '
End Function
'


