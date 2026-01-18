Attribute VB_Name = "GANTT1"
Public CAPERIO%
Public APERIO
Public DESPERIO(32)
Public XINI(32)
Public YINI(512)
Public YACARGANTT%
Dim XMAXI
Public CODEQ$(512), STATEQ%(512), YEQU(512)
Public CAEQ%, CAEQVIS%, PRIEQVIS%, PRIOFVIS%
Public CONTROCLICK%
Public FEIPER
Dim MATIE$(1024)
Public NORFB$(1024), CIOFA%(1024)
Dim CAOFA%, CACODACT%
Public XINI1(2048), XFIN1(2048), YINI1(2048), NUOFA$(2048), cabarr%
'
Sub DIBULINVER()
   CAPERIO% = 7
   If FORMUGANTT.Option1(0) = True Then
        CAPERIO% = 2
      ElseIf FORMUGANTT.Option1(2) = True Then
        CAPERIO% = 15
      ElseIf FORMUGANTT.Option1(3) = True Then
        CAPERIO% = 31
   End If
   APERIO = (FORMUGANTT.Picture1.Width - FORMUGANTT.Line2.X1 - FORMUGANTT.VScroll1.Width - 80) / (CAPERIO%)
   FORMUGANTT.DrawWidth = 2
   X1 = FORMUGANTT.Line2.X1
   Y2 = FORMUGANTT.Picture1.Height
   For II% = 1 To CAPERIO% + 1
      XINI(II%) = X1
      X1 = X1 + APERIO
      If II% = CAPERIO% Then XMAXI = X1 - 20
      FORMUGANTT.Picture1.Line (X1, 0)-(X1, Y2), QBColor(0)
   Next II%
   FORMUGANTT.Picture1.Refresh
End Sub
'

Sub CALCUFECHA()
   FEBA = FEIPER
   For II% = 1 To CAPERIO% + 1
      DESPERIO(II%) = FEBA
10    FEBA = DIPOST(FEBA)
      If FEBA = DESPERIO(II%) Then GoTo 10
   Next II%
End Sub
'

Sub DIBUFECHA()
   FORMUGANTT.Picture1.Font = FORMUGANTT.TEPRUEBA.Font
   FORMUGANTT.TEPRUEBA.FontBold = True
   If FORMUGANTT.Option1(2).Value = True Then
      FORMUGANTT.TEPRUEBA.FontBold = False
   End If
   FORMUGANTT.Picture1.FontSize = FORMUGANTT.TEPRUEBA.FontSize
   FORMUGANTT.Picture1.FontBold = FORMUGANTT.TEPRUEBA.FontBold
   FORMAFE% = 0: If CAPERIO% > 15 Then FORMAFE% = 1
   For II% = 1 To CAPERIO%
      TTXX$ = FECHATEX$(DESPERIO(II%))
      If CAPERIO% > 7 Then
         TTXX$ = Left$(TTXX$, 5)
      End If
      If CAPERIO% > 15 Then
         TTYY$ = Mid$(TTXX$, 4, 2)
         TTXX$ = Left$(TTXX$, 2)
         If Left$(TTYY$, 1) = "0" Then TTYY$ = Mid$(TTYY$, 2)
      End If
      FORMUGANTT.TEPRUEBA.Caption = TTXX$
      FORMUGANTT.Picture1.CurrentX = FORMUGANTT.Line2.X1 + APERIO * II% - APERIO / 2 - FORMUGANTT.TEPRUEBA.Width / 2
      FORMUGANTT.Picture1.CurrentY = 120 - 60 * FORMAFE%
      FORMUGANTT.Picture1.Print TTXX$
      If FORMAFE% = 1 Then
         FORMUGANTT.TEPRUEBA.Caption = "---"
         FORMUGANTT.Picture1.CurrentX = FORMUGANTT.Line2.X1 + APERIO * II% - APERIO / 2 - FORMUGANTT.TEPRUEBA.Width / 2
         FORMUGANTT.Picture1.CurrentY = 165
         FORMUGANTT.Picture1.Print "---"
         FORMUGANTT.TEPRUEBA.Caption = TTYY$
         FORMUGANTT.Picture1.CurrentX = FORMUGANTT.Line2.X1 + APERIO * II% - APERIO / 2 - FORMUGANTT.TEPRUEBA.Width / 2
         FORMUGANTT.Picture1.CurrentY = 300
         FORMUGANTT.Picture1.Print TTYY$
        Else
         TTYY$ = DSEM$(DESPERIO(II%))
         If CAPERIO% > 7 Then TTYY$ = Left$(TTYY$, 2)
         TTYY$ = Left$(TTYY$, 1) + LCase$(Mid$(TTYY$, 2))
         TTYY$ = "(" + TTYY$ + ")"
         FORMUGANTT.TEPRUEBA.Caption = TTYY$
         FORMUGANTT.Picture1.CurrentX = FORMUGANTT.Line2.X1 + APERIO * II% - APERIO / 2 - FORMUGANTT.TEPRUEBA.Width / 2
         FORMUGANTT.Picture1.CurrentY = 300
         FORMUGANTT.Picture1.Print TTYY$
      End If
   Next II%
   FORMUGANTT.TEPRUEBA.FontBold = True
End Sub
'

Sub DIBUEQUIP()
   Static MAXSCROLL%
   '
   FORMUGANTT.Picture1.Font = FORMUGANTT.TEPRUEBA.Font
   FORMUGANTT.Picture1.FontSize = FORMUGANTT.TEPRUEBA.FontSize
   FORMUGANTT.Picture1.FontBold = FORMUGANTT.TEPRUEBA.FontBold
   INCREY = FORMUGANTT.TEPRUEBA.Height * 1.2
   FORMUGANTT.Picture1.CurrentY = FORMUGANTT.Line1.Y1 - INCREY / 2
   '
   If PRIEQVIS% < 1 Then PRIEQVIS% = 1
   For U& = PRIEQVIS% To CAEQ%
      FORMUGANTT.Picture1.CurrentY = FORMUGANTT.Picture1.CurrentY + INCREY
      YINI(U&) = FORMUGANTT.Picture1.CurrentY
      If YINI(U&) <= FORMUGANTT.Picture1.Height - FORMUGANTT.HScroll1.Height - INCREY Then
         FORMUGANTT.Picture1.CurrentX = 80
         FORMUGANTT.Picture1.Print CODEQ$(U&);
      End If
   Next U&
   '
20 If CAEQVIS% = 0 Then
      CAEQVIS% = (FORMUGANTT.Picture1.Height - FORMUGANTT.Line1.Y1 - FORMUGANTT.HScroll1.Height - INCREY) / INCREY
   End If
   '
   If MAXSCROLL% <= 0 Then
      FORMUGANTT.VScroll1.Max = 0
      If CAEQVIS% < CAEQ% Then
         FORMUGANTT.VScroll1.Max = CAEQ% - CAEQVIS%
      End If
      MAXSCROLL% = FORMUGANTT.VScroll1.Max
   End If
   FORMUGANTT.VScroll1.Max = MAXSCROLL%
   FORMUGANTT.VScroll1.Value = PRIEQVIS% - 1
   '
End Sub
'

Sub DIBUBARRA(FEINI%, HOINI%, FEFIN%, HOFIN%, EQUI$, COLO$, PREPAR%)
   '
   Static DHR0%, TUNOR
   If TUNOR < 1 Then
      'DHR0% = 2400: TUNOR = 0
      'For U& = 1 To ULTREG&("TURGEN")
      '   X$ = REGLEIDO$("TURGEN", U&)
      '   DSX% = Asc(Left$(X$, 1))
      '   If DSX% > 0 Then
      '      DHR% = CVI(Mid$(X$, 2, 2))
      '      HHR% = CVI(Mid$(X$, 4, 2))
      '      If HHR% > DHR% Then
      '         If DHR% < DHR0% Then DHR0% = DHR%
      '         TT1 = 60 * Int(DHR% / 100) + (DHR% Mod 100)
      '         TT2 = 60 * Int(HHR% / 100) + (HHR% Mod 100)
      '         If (TT2 - TT1) > TUNOR Then TUNOR = TT2 - TT1
      '      End If
      '   End If
      'Next U&
      'DHR0% = 60 * Int(DHR0% / 100) + (DHR0% Mod 100)
   End If
   TUNOR = 24 * 60: DHR0% = 0
   '
   If FEINI% > DESPERIO(CAPERIO% + 1) Then
      Exit Sub
   End If
   '
   If FEFIN% < DESPERIO(1) Then
      Exit Sub
   End If
   '
   If FEINI% < DESPERIO(1) Then
      X1 = XINI(1)
      GoTo 10
   End If
   '
   For II% = 1 To CAPERIO%
      If FEINI% = DESPERIO(II%) Then
         MINU1% = 60 * Int(HOINI% / 100) + (HOINI% Mod 100)
         X1 = XINI(II%) + APERIO * (MINU1% - DHR0%) / TUNOR
         GoTo 10
      End If
   Next II%
   Exit Sub
   '
10 If FEFIN% >= DESPERIO(CAPERIO% + 1) Then
      X2 = XINI(CAPERIO% + 1)
      GoTo 20
   End If
   '
   For II% = 1 To CAPERIO%
      If FEFIN% = DESPERIO(II%) Then
         MINU2% = 60 * Int(HOFIN% / 100) + (HOFIN% Mod 100)
         X2 = XINI(II%) + APERIO * (MINU2% - DHR0%) / TUNOR
         GoTo 20
      End If
   Next II%
   '
20 For II% = PRIEQVIS% To CAEQ%
      If CODEQ$(II%) = EQUI$ Then
         Y1 = YINI(II%)
         GoTo 30
      End If
   Next II%
   Exit Sub
   '
30 If Y1 < 100 Then Exit Sub
   If X2 > XMAXI Then X2 = XMAXI
   ROJO% = Asc(Left$(COLO$, 1))
   VERDE% = Asc(Mid$(COLO$, 2, 1))
   AZUL% = Asc(Mid$(COLO$, 3, 1))
   FORMUGANTT.Picture1.Line (X1, Y1)-(X2, Y1 + 145), RGB(ROJO%, VERDE%, AZUL%), BF
   FORMUGANTT.Picture1.Line (X1 + 12, Y1 + 145)-(X2, Y1 + 157), QBColor(0), BF
   FORMUGANTT.Picture1.Line (X2 - 12, Y1 + 25)-(X2, Y1 + 157), QBColor(0), BF
   '
   cabarr% = cabarr% + 1
   XINI1(cabarr%) = X1
   XFIN1(cabarr%) = X2
   YINI1(cabarr%) = Y1
   '
End Sub
'

Sub CARLIORF()
   '
   FORMUGANTT.List1.Clear
   CAOFA% = 0
   CAEQ% = 0
   For U& = 1 To ULTREG&("PADMAQ")
      ZZ$ = REGLEIDO$("PADMAQ", U&)
      If TRIM$(Left$(ZZ$, 6)) <> "" Then
         CAEQ% = CAEQ% + 1
         CODEQ$(CAEQ%) = Left$(ZZ$, 6)
         STATEQ%(CAEQ%) = 0
      End If
   Next U&
   '
   For U& = 1 To ULTREG&("OPERFAP")
      '
      XX$ = REGLEIDO$("OPERFAP", U&)
      NOFA$ = TRIM$(Left$(XX$, 12))
      MAQUI$ = Mid$(XX$, 73, 6) + Mid$(XX$, 203, 18)
      'ITIEM$ = Mid$(XX$, 135, 8)
      ITIEM$ = Mid$(XX$, 135, 12)
      CIIX% = CVI(Mid$(XX$, 13, 2))
      '
      YACODI% = 0
      For K% = 1 To CAOFA%
         If CIOFA%(K%) = CIIX% Then YACODI% = 1
         If NORFB$(K%) = NOFA$ Then GoTo 10
      Next K%
      CAOFA% = CAOFA% + 1
      K% = CAOFA%
      NORFB$(CAOFA%) = NOFA$
      CIOFA%(CAOFA%) = CIIX%
      MATIE$(CAOFA%) = ""
      TTXX$ = Space$(32)
      Call REPLA(TTXX$, NOFA$, 1, 24)
      Call REPLA(TTXX$, Str$(CAOFA%), 25, 8)
      FORMUGANTT.List1.AddItem TTXX$
      If YACODI% = 0 Then
         FORMUGANTT.List3.AddItem CODEXT$(CIIX%)
      End If
      '
10    MATIE$(K%) = MATIE$(K%) + MAQUI$ + ITIEM$
      '
      For K% = 1 To 24 Step 6
         EQUIPO$ = Mid$(MAQUI$, K%, 6)
         If TRIM$(EQUIPO$) <> "" Then
            For J% = 1 To CAEQ%
               If CODEQ$(J%) = EQUIPO$ Then
                  STATEQ%(J%) = 1
                  GoTo 20
               End If
            Next J%
            Call MENSERR(24, "Equipo '" + TRIM$(EQUIPO$) + "'no Figura en Tablas\en O/F Número " + TRIM$(NOFA$) + "para Código " + TRIM$(CODEXT$(CIIX%)) + ".")
         End If
20    Next K%
      '
   Next U&
   '
   J% = 0
   For I% = 1 To CAEQ%
      If STATEQ%(I%) = 1 Then
         J% = J% + 1
         CODEQ$(J%) = CODEQ$(I%)
         STATEQ%(J%) = 1
      End If
   Next I%
   CAEQ% = J%
   '
End Sub
'

Sub PREGANTT()
   '
   CONTROCLICK = (-1)
   'TOPE = FORMUGANTT.List1.TopIndex
   'ACTI = FORMUGANTT.List1.ListIndex
   cabarr% = 0
   For U& = 1 To FORMUGANTT.List1.ListCount
      NOFA$ = TRIM$(Left$(FORMUGANTT.List1.List(U& - 1), 24))
      'NOFA$ = NORFB$(U&)
      COLO$ = Chr$(0) + Chr$(230) + Chr$(64)
      '
      If FORMUGANTT.List1.Selected(U& - 1) = True Then
         COLO$ = Chr$(255) + Chr$(0) + Chr$(0)
      End If
      '
      If FORMUGANTT.List1.ListIndex = U& - 1 Then
         COLO$ = Chr$(128) + Chr$(0) + Chr$(128)
      End If
      '
      COLOPRE$ = Chr$(100) + Chr$(100) + Chr$(0)
      V& = Val(Mid$(FORMUGANTT.List1.List(U& - 1), 25))
      'For J% = 1 To Len(MATIE$(V&)) Step 32
      For J% = 1 To Len(MATIE$(V&)) Step 36
         MAQUI$ = Mid$(MATIE$(V&), J%, 24)
         'FEHOR$ = Mid$(MATIE$(V&), J% + 24, 8)
         FEHOR$ = Mid$(MATIE$(V&), J% + 24, 12)
         FEINI% = CVI(Mid$(FEHOR$, 1, 2))
         HOINI% = CVI(Mid$(FEHOR$, 3, 2))
         FEFIN% = CVI(Mid$(FEHOR$, 5, 2))
         HOFIN% = CVI(Mid$(FEHOR$, 7, 2))
           FEFINPREP% = CVI(Mid$(FEHOR$, 9, 2))
           HOFINPREP% = CVI(Mid$(FEHOR$, 11, 2))
         If FEINI% > FEFIN% Then
            FEINI% = 0
            HOINI% = 0
            FEFIN% = 0
            HOFIN% = 0
              FEFINPREP% = 0
              HOFINPREP% = 0
         End If
         For K% = 1 To 24 Step 6
            EQUI$ = Mid$(MAQUI$, K%, 6)
            If TRIM$(EQUI$) <> "" Then
               CBX% = cabarr%
               Call DIBUBARRA(FEINI%, HOINI%, FEFIN%, HOFIN%, EQUI$, COLO$, 0)
               Call DIBUBARRA(FEINI%, HOINI%, FEFINPREP%, HOFINPREP%, EQUI$, COLOPRE$, 1)
               If cabarr% > CBX% Then
                  NUOFA$(cabarr%) = NOFA$
               End If
            End If
         Next K%
      Next J%
      '
   Next U&
   '
   FORMUGANTT.HScroll1.LargeChange = CAPERIO%
   CONTROCLICK% = 0
   YACARGANTT% = 1
   '
End Sub
'

Function RANCOL$()
   Static ROJO%, VERDE%, AZUL%
   ICOLOR% = 160
   '
   ROJO% = ROJO% + ICOLOR%
5  If ROJO% > 255 Then
      ROJO% = ROJO% - 256
      VERDE% = VERDE% + ICOLOR%
   End If
   If VERDE% > 255 Then
      VERDE% = VERDE% - 256
      AZUL% = AZUL% + ICOLOR%
   End If
   If AZUL% > 255 Then
      AZUL% = AZUL% - 256
      ROJO% = ROJO% + ICOLOR%
      GoTo 5
   End If
   '
   RANCOL$ = Chr$(ROJO%) + Chr$(VERDE%) + Chr$(AZUL%)
End Function
'

Sub MUEDATORF(NORFA$)
   '
   FORMUGANTT.Label2.Caption = ""
   FORMUGANTT.Label4.Caption = ""
   FORMUGANTT.Label6.Caption = ""
   FORMUGANTT.Label8.Caption = ""
   FORMUGANTT.Label10.Caption = ""
   FORMUGANTT.Label11.Caption = ""
   FORMUGANTT.Label13.Caption = ""
   FORMUGANTT.Label15.Caption = ""
   FORMUGANTT.List2.Clear
   '
   ABUSCA$ = "ORDEFABP"
5  For U& = ULTREG&(ABUSCA$) To 1 Step -1
      XX$ = REGLEIDO$(ABUSCA$, U&)
      If TRIM$(Mid$(XX$, 69, 12)) = NORFA$ Then
         FORMUGANTT.Label2.Caption = NORFA$
         CIIX% = CVI(Mid$(XX$, 83, 2))
         FORMUGANTT.Label10.Caption = CODEXT$(CIIX%)
         FORMUGANTT.Label11.Caption = DESCRIT$(CIIX%)
         FORMUGANTT.Label13.Caption = CVD(Mid$(XX$, 89, 8))
         FORMUGANTT.Label6.Caption = Mid$(XX$, 97, 8)
         FEX = CVI(Mid$(XX$, 105, 2))
         FORMUGANTT.Label8.Caption = FECHATEX$(FEX)
         FORMUGANTT.Label4.Caption = Asc(Mid$(XX$, 108, 1))
         FORMUGANTT.Label15.Caption = CVD(Mid$(XX$, 89, 8)) - CVD(Mid$(XX$, 119, 8))
         GoTo 10
      End If
   Next U&
   '
   If ABUSCA$ = "ORDEFABP" Then
      ABUSCA$ = "ORDEFABR"
      GoTo 5
   End If
   '
10 For U& = 1 To ULTREG&("OPERFAP")
      XX$ = REGLEIDO$("OPERFAP", U&)
      If TRIM$(Left$(XX$, 12)) = NORFA$ Then
         TTXX$ = Space$(128)
         Call REPLA(TTXX$, CSTRING$(Mid$(XX$, 15, 2), 1, 4, 0, 2), 1, 4)
         Call REPLA(TTXX$, Mid$(XX$, 17, 48), 6, 26)
         Call REPLA(TTXX$, Mid$(XX$, 73, 6), 33, 6)
         'Call REPLA(TTXX$, CSTRING$(Mid$(XX$, 79, 8), 4, 10, 1, 2), 44, 10)
         'Call REPLA(TTXX$, CSTRING$(Mid$(XX$, 87, 8), 4, 10, 1, 2), 54, 10)
         SALDO# = CVD(Mid$(XX$, 79, 8)) - CVD(Mid$(XX$, 87, 8))
         If SALDO# < 0 Then SALDO# = 0
         Call REPLA(TTXX$, CSTRING$(MKD$(SALDO#), 4, 10, 1, 2), 39, 10)
         Call REPLA(TTXX$, Left$(HORMINSE$(60 * CVS(Mid$(XX$, 131, 4))), 6), 50, 6)
         FEX1 = CVI(Mid$(XX$, 135, 2))
         HORA1% = CVI(Mid$(XX$, 137, 2))
         FEX2 = CVI(Mid$(XX$, 139, 2))
         HORA2% = CVI(Mid$(XX$, 141, 2))
         If FEX1 > FEX2 Then
            FEX1 = 0: FEX2 = 0
            HORA1% = 0: HORA2% = 0
            Call REPLA(XX$, String$(8, 0), 135, 8)
            Call GRAREG("OPERFAP", XX$, U&)
         End If
         Call REPLA(TTXX$, Left$(FECHATEX$(FEX1), 5), 58, 5)
         Call REPLA(TTXX$, Left$(HORATEXNOR$(HORA1%), 5), 64, 5)
         Call REPLA(TTXX$, Left$(FECHATEX$(FEX2), 5), 72, 5)
         Call REPLA(TTXX$, Left$(HORATEXNOR$(HORA2%), 5), 78, 5)
         FORMUGANTT.List2.AddItem TTXX$
      End If
   Next U&
   FORMUGANTT.List2.Refresh
End Sub

Function COLOMAQ$(MAQUI$)
    '
    CLM$ = ""
    If TRIM$(MAQUI$) = "" Then GoTo 99
    If TRIM$(ECOARCH$("PADMAQ", TRIM$(MAQUI$))) = "" Then GoTo 99
    '
    For U& = 1 To ULTREG("PADMAQ")
       X$ = REGLEIDO$("PADMAQ", U&)
       If TRIM$(Left$(X$, 6)) = TRIM$(MAQUI$) Then
          CLM$ = Mid$(X$, 97, 3)
       End If
    Next U&
    '
99  COLOMAQ$ = CLM$
    '
End Function
