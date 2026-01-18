Attribute VB_Name = "COMDEMO"
Public MINIDEM%
Public YAPRELICEN%
Public IIMAGEN%, PAUSA%
Public COMACON$
Dim ATRIBU$
Dim LOARGUA&
Dim XFINAL, YFINAL
Dim ORDEGUARDA%
Public DEMOPATH$
Dim ORIFLE%, TIPFLE%, COLOFLE, XFLECH%, YFLECH%
Public FEDICION As String
'

Sub SHOWIMAGEN(ADEMO$, LDEMO$, INSTRU$, PAUSI%, PAUSF%, MABRE%)
   '
   
   If MINIDEM% = 1 Then
     BLACKSCREEN.Show
     GoTo 10
   End If
   '
   If DEMOFLE1.Visible = True Then
     MENUDEMO.Hide
     BLACKSCREEN.Show
     DEMOFLE1.Hide
     DoEvents
   End If
   '
10 If EXISTIMAGE%(DEMOPATH$ + "\SCREENS\" + ADEMO$) < 1 Then Exit Sub
   FORDEMO.Image1.Picture = LoadPicture("")
   For KKI% = 0 To 9
     FORDEMO.TRAMO1(KKI%).Visible = False
     FORDEMO.TRAMO2(KKI%).Visible = False
     FORDEMO.TRAMO3(KKI%).Visible = False
     FORDEMO.SOMBRA1(KKI%).Visible = False
     FORDEMO.SOMBRA2(KKI%).Visible = False
     FORDEMO.SOMBRA3(KKI%).Visible = False
     FORDEMO.PUNTA(KKI%).Visible = False
     FORDEMO.PUNTB(KKI%).Visible = False
   Next KKI%
   '
   For KKI% = 0 To 9
     FORDEMO.Picture2(KKI%).Visible = False
     FORDEMO.TRAMO1(KKI%).Visible = False
     FORDEMO.TRAMO2(KKI%).Visible = False
     FORDEMO.TRAMO3(KKI%).Visible = False
     FORDEMO.SOMBRA1(KKI%).Visible = False
     FORDEMO.SOMBRA2(KKI%).Visible = False
     FORDEMO.SOMBRA3(KKI%).Visible = False
     FORDEMO.PUNTA(KKI%).Visible = False
     FORDEMO.PUNTB(KKI%).Visible = False
   Next KKI%
   '
   On Error GoTo 30
   FORDEMO.Image1.Picture = LoadPicture(DEMOPATH$ + "\SCREENS\" + ADEMO$)
   FORDEMO.Width = FORDEMO.Image1.Width
   FORDEMO.Label1.Width = FORDEMO.Width - FORDEMO.Picture1.Width + 20
   FORDEMO.Height = FORDEMO.Image1.Height
     APANTA = Screen.Width
     'MsgBox APANTA
     HPANTA = Screen.Height
     'MsgBox HPANTA
   FORDEMO.Left = (APANTA - FORDEMO.Width) / 2
   FORDEMO.Top = (0.967 * HPANTA - FORDEMO.Height) / 2
   FORDEMO.Picture1.Left = FORDEMO.Width - FORDEMO.Picture1.Width - 30
   FORDEMO.Picture1.Refresh
   FORDEMO.Label1 = LDEMO$
   '
   FORDEMO.Show
   DoEvents
   On Error GoTo 0
   '
22 DoEvents
   PAUSAX = 0.5
   TINI = Timer
25 While Timer < TINI + PAUSAX + PAUSI% * 100 / FORDEMO.HScroll1.Value
     If COMACON$ <> "" Then
       Exit Sub
     End If
     DoEvents
   Wend
   '
   Call SECULEYEN(INSTRU$)
   '
32 DoEvents
   TINI = Timer
35 While Timer < TINI + PAUSF% * 100 / FORDEMO.HScroll1.Value
     If COMACON$ <> "" Then
       Exit Sub
     End If
     DoEvents
   Wend
   On Error GoTo 0
   Exit Sub
   '
30 On Error GoTo 0
   FORDEMO.Picture1.Left = FORDEMO.Width - FORDEMO.Picture1.Width - 30
   FORDEMO.Picture1.Refresh
   DoEvents
   MsgBox "Imposible Abrir '" + DEMOPATH$ + "\SCREENS\" + ADEMO$ + "'."
   DoEvents
   '
End Sub

Function EXISTIMAGE%(AIMAGEN$)
   '
   EXI% = 0
   On Error Resume Next
   FILEN = FileLen(AIMAGEN$)
   On Error GoTo 0
   If FILEN > 0 Then
     EXI% = 1
   End If
   '
   EXISTIMAGE% = EXI%
   '
End Function
'
Sub SHOWTEXTO()
   '
   Static MODESPLI%
   MODESPLI% = (MODESPLI% + 1) Mod 4
   '
   If Trim$(INFODEM.Label3.Caption) = "" Then
     Exit Sub
   End If
   '
   DEMOTO = 7: DEMOLE = 2
   TTOTAL = DEMOTO - DEMOLE - 3: If TTOTAL < 1 Then TTOTAL = 1
   '
   DEMOTO = DEMOTO * 100 / FORDEMO.HScroll1.Value
   DEMOLE = DEMOLE * 100 / FORDEMO.HScroll1.Value
   TTOTAL = TTOTAL * 100 / FORDEMO.HScroll1.Value
   DIFLE% = 0
   '
5  On MODESPLI% + 1 GoTo 10, 20, 30, 40
   '
10 ' COMPLETO DESDE INFERIOR IZQUIERDO
  'XACTUAL = XFINAL
   If YACAR% = 0 Then
      'XACTUAL = FORDEMO.Image1.Left
      XACTUAL = FORDEMO.Left
      'YACTUAL = FORDEMO.Image1.Top + FORDEMO.Image1.Height
      YACTUAL = FORDEMO.Top + FORDEMO.Height
      INFODEM.Top = YACTUAL
      INFODEM.Left = XACTUAL
      '
      'INCREY = (FORDEMO.Image1.Height - YFINAL) / (TTOTAL * 1000 / 20)
      INCREY = (YACTUAL - YFINAL) / (TTOTAL * 1000 / 20)
      INCREX = (XFINAL - XACTUAL) / (TTOTAL * 1000 / 20)
      INFODEM.Picture2.Visible = True
      INFODEM.Label3.Visible = True
      INFODEM.Show
      YACAR% = 1
   End If
   '
   If YACTUAL > YFINAL Then YACTUAL = YACTUAL - INCREY: MUEVE% = 1
   If XACTUAL < XFINAL Then XACTUAL = XACTUAL + INCREX: MUEVE% = 1
   If MUEVE% = 1 Then
     INFODEM.Move XACTUAL, YACTUAL
   End If
   GoTo 90
   '
    '
  'XACTUAL = XFINAL
20 If YACAR% = 0 Then
      XACTUAL = XFINAL
      YACTUAL = YFINAL
      INFODEM.Top = YACTUAL
      INFODEM.Left = XACTUAL
      '
      ALTUFIN = INFODEM.Height
      ANCHFIN = INFODEM.Width
      INCREY = (INFODEM.Height - 100) / (TTOTAL * 1000 / 20)
      INCREX = (INFODEM.Width) / (TTOTAL * 1000 / 20)
      INFODEM.Width = 0
      INFODEM.Height = 0
      INFODEM.Picture2.Visible = True
      INFODEM.Label3.Visible = True
      INFODEM.Show
      YACAR% = 1
   End If
   '
   If INFODEM.Height < ALTUFIN Then INFODEM.Height = INFODEM.Height + INCREY: MUEVE% = 1
   If INFODEM.Width < ANCHFIN Then INFODEM.Width = INFODEM.Width + INCREX: MUEVE% = 1
   GoTo 90
   '
30 If YACAR% = 0 Then
      XACTUAL = XFINAL
      YACTUAL = YFINAL
      INFODEM.Top = YACTUAL
      INFODEM.Left = XACTUAL
      '
      ALTUFIN = INFODEM.Height
      INCREY = (INFODEM.ScaleHeight) / (TTOTAL * 1000 / 20)
      INFODEM.Height = 0
      INFODEM.Picture2.Visible = True
      INFODEM.Label3.Visible = True
      INFODEM.Show
      YACAR% = 1
   End If
   '
   If INFODEM.Height < ALTUFIN Then INFODEM.Height = INFODEM.Height + INCREY: MUEVE% = 1
   GoTo 90
   '
40 If YACAR% = 0 Then
      XACTUAL = XFINAL
      YACTUAL = YFINAL + INFODEM.ScaleHeight
      INFODEM.Top = YACTUAL
      INFODEM.Left = XACTUAL
      '
      ALTUFIN = INFODEM.Height
      INCREY = (INFODEM.Height) / (TTOTAL * 1000 / 20)
      INFODEM.Height = 0
      INFODEM.Picture2.Visible = True
      INFODEM.Label3.Visible = True
      INFODEM.Show
      YACAR% = 1
   End If
   '
   If INFODEM.Height < ALTUFIN Then
      MUEVE% = 1
      YACTUAL = YACTUAL - INCREY
      If MUEVE% = 1 Then
        INFODEM.Move XACTUAL, YACTUAL
        INFODEM.Height = INFODEM.Height + INCREY
      End If
   End If
   GoTo 90
   '
90 If MUEVE% = 0 Then
     LLEGO% = LLEGO% + 1
     If DIFLE% = 0 Then
       Call DIBUFLECHA
       Call RELEYENDA
       INFODEM.Hide
       DIFLE% = 1
     End If
   End If
   '
   If LLEGO% >= 50 Then
       TINI = Timer: While Timer < TINI + 5 * 100 / FORDEMO.HScroll1.Value: Wend
       'Call RELEYENDA
       'INFODEM.Hide
       Exit Sub
     Else
       TINI = Timer
       While Timer < TINI + 0.02
         If COMACON$ <> "" Then
           Exit Sub
         End If
         DoEvents
       Wend
       MUEVE% = 0
       GoTo 5
   End If
   '
End Sub

Sub SECULEYEN(ADEMO$)
   '
   AGUARDA$ = UCase$(ADEMO$)
   PPXX% = InStr(AGUARDA$, ".")
   If PPXX% < 1 Then PPXX% = 1 + Len(AGUARDA$)
   AGUARDA$ = Left$(AGUARDA$, PPXX% - 1) + ".LB"
   '
   For ORLEYE% = 0 To 9
10   ALEYEN$ = AGUARDA$ + Trim$(Str$(ORLEYE%))
     LOARGUA& = 0
     On Error Resume Next
     LOARGUA& = FileLen(DEMOPATH$ + "\LABELS\" + ALEYEN$)
     On Error GoTo 0
     If LOARGUA& > 0 Then
        ORDEGUARDA% = ORLEYE%
        Call CARGALEYEN(ALEYEN$)
        If COMACON$ <> "" Then
          If COMACON$ = "NEXT" Then COMACON$ = "": GoTo 19
          If COMACON$ = "PREV" And ORLEYE% > 0 Then COMACON$ = "": ORLEYE% = ORLEYE% - 1: GoTo 10
          Exit Sub
        End If
     End If
19 Next ORLEYE%
   '
End Sub

Sub CARGALEYEN(ALEYEN$)
   '
   N1% = FreeFile
   Open DEMOPATH$ + "\LABELS\" + ALEYEN$ For Binary Access Read Shared As #N1%
   TEGRABA$ = Input(LOARGUA&, #N1%)
   Close #N1%
   '
   PPXX% = InStr(TEGRABA$, "|")
   If PPXX% < 1 Then PPXX% = 1 + Len(TEGRABA$)
   '
   TELEYEN$ = Left$(TEGRABA$, PPXX% - 1)
   INFODEM.Label3.Caption = TELEYEN$
   '
   ATRIBU$ = Mid$(TEGRABA$, PPXX% + 1)
   '
   TITUA$ = Trim$(ATRILE$("TITULE"))
   TITUE$ = "": For KKI% = 1 To Len(TITUA$)
                  CARA% = Asc(Mid$(TITUA$, KKI%, 1))
                  If CARA% >= 32 Then TITUE$ = TITUE$ + Chr$(CARA%)
                Next KKI%
   INFODEM.Caption = TITUE$
   XFINAL = Val(ATRILE$("XFINAL")) + FORDEMO.Left
   YFINAL = Val(ATRILE$("YFINAL")) + FORDEMO.Top
   On Error Resume Next
   INFODEM.TEPRUEBA.FontName = ATRILE$("FONAME")
   INFODEM.TEPRUEBA.FontSize = ATRILE$("FOSIZE")
   INFODEM.TEPRUEBA.FontBold = ATRILE$("FOBOLD")
   INFODEM.TEPRUEBA.FontItalic = ATRILE$("FOITAL")
   INFODEM.TEPRUEBA.FontUnderline = ATRILE$("FUNDER")
   '
   INFODEM.Label3.FontName = ATRILE$("FONAME")
   INFODEM.Label3.FontSize = ATRILE$("FOSIZE")
   INFODEM.Label3.FontBold = ATRILE$("FOBOLD")
   INFODEM.Label3.FontItalic = ATRILE$("FOITAL")
   INFODEM.Label3.FontUnderline = ATRILE$("FUNDER")
   '
   ROJO = Val(ATRILE$("LEROJO"))
   VERDE = Val(ATRILE$("LEVERD"))
   AZUL = Val(ATRILE$("LEAZUL"))
   INFODEM.Label3.ForeColor = RGB(ROJO, VERDE, AZUL)
   '
   ROJO = Val(ATRILE$("FOROJO"))
   VERDE = Val(ATRILE$("FOVERD"))
   AZUL = Val(ATRILE$("FOAZUL"))
   INFODEM.Picture2.BackColor = RGB(ROJO, VERDE, AZUL)
   '
   ALINE% = Val(ATRILE$("ALINEA"))
   INFODEM.Label3.Alignment = ALINE%
   '
   AMAXI = 0
   TTXX$ = TELEYEN
10 PPXX% = InStr(TTXX$, Chr$(10) + Chr$(13))
   If PPXX% < 1 Then PPXX% = 1 + Len(TTXX$)
   TEPRU$ = Left$(TTXX$, PPXX% - 1)
   TTXX$ = Mid$(TTXX$, PPXX% + 2)
   INFODEM.TEPRUEBA.Caption = TEPRU$
   ATEPRU = INFODEM.TEPRUEBA.Width
   If ATEPRU > AMAXI Then AMAXI = ATEPRU
   If Len(Trim$(TTXX$)) > 0 Then GoTo 10
   '
   INFODEM.Picture2.Left = 0
   INFODEM.Picture2.Top = 0
   INFODEM.Label3.Left = 50
   INFODEM.Label3.Top = 50
   INFODEM.Label3.Width = AMAXI
   INFODEM.Picture2.Width = INFODEM.Label3.Width + 150
   INFODEM.Picture2.Height = INFODEM.Label3.Height + 150
   INFODEM.Width = INFODEM.Picture2.Width + 60
   INFODEM.Height = INFODEM.Picture2.Height + 300
   '
   'ROJO(2) = Val(ATRILE$("PUROJO"))
   'VERDE(2) = Val(ATRILE$("PUVERD"))
   'AZUL(2) = Val(ATRILE$("PUAZUL"))
   'If Option2(2).Value = True Then
   '  HScroll1.Value = ROJO(2)
   '  HScroll2.Value = VERDE(2)
   '  HScroll3.Value = AZUL(2)
   'End If
   '
   'ORIFLECHA% = Val(ATRILE$("ORIFLE"))
   '
   ROJOI = Val(ATRILE$("PUROJO"))
   VERDEI = Val(ATRILE$("PUVERD"))
   AZULI = Val(ATRILE$("PUAZUL"))
   COLOFLE = RGB(ROJOI, VERDEI, AZULI)
   ORIFLE% = Val(ATRILE$("ORIFLE"))
   TIPFLE% = Val(ATRILE$("TIPFLE"))
   XFLECH% = Val(ATRILE$("XFLECH"))
   YFLECH% = Val(ATRILE$("YFLECH"))
   '
   On Error GoTo 0
   Call SHOWTEXTO
   '
   If COMACON$ <> "" Then
     Exit Sub
   End If
   '
   Screen.MousePointer = 1
   '
End Sub

Sub SECUDEMO(SECUEN$)
   '
   Screen.MousePointer = 1
   '
   Call CARLISECUEN(SECUEN$)
   If SECUENDE.List1.ListCount < 1 Then
     Screen.MousePointer = 1
     Exit Sub
   End If
   '
   ULTIMAGEN% = SECUENDE.List1.ListCount
5  IIMAGEN% = 0
   '
10 If IIMAGEN% < 1 Then CAPINU% = 0
   IIMAGEN% = IIMAGEN% + 1
15 If IIMAGEN% < 1 Then GoTo 10
   If IIMAGEN% > ULTIMAGEN% Then
      NOSECUEN$ = "Secuencia " + Trim$(SECUEN$)
      If MINIDEM% = 1 Then GoTo 17
      N1% = FreeFile
      Dim BUFLE As String * 128
      Open DEMOPATH$ + "\SCRIPTS\SECUDEMO.DAT" For Random Access Read Shared As #N1% Len = 128
      For KKI% = 1 To LOF(N1%) / 128
        Get #N1%, KKI%, BUFLE$
        If Trim$(UCase$(Left$(BUFLE$, 10))) = UCase$(Trim$(SECUEN$)) Then
          NOSECUEN$ = Trim$(UCase$(Mid$(BUFLE$, 11)))
          Close #N1%
          GoTo 17
        End If
      Next KKI%
      Close #N1%
      '
17    FINSECUEN.Label6.Caption = NOSECUEN$
      FINSECUEN.Label1.Caption = FINSECUEN.Label7.Caption + " " + Trim$(FEDICION) + "."
      If MINIDEM% = 1 Then
         FINSECUEN.Label6.Caption = "Mini-Demostración FLEXOFT"
         FINSECUEN.Command2.Caption = "Terminar"
      End If
      '
      FORDEMO.Hide
      'If MINIDEM% <> 1 Then
         BLACKSCREEN.Show
      'End If
      FINSECUEN.Show 1
      If FINSECUEN.Label9.Caption = "1" Then GoTo 5 'vuelve a empezar
      If MINIDEM% = 1 Then
         Close: End
         Exit Sub
      End If
      FORDEMO.Hide
      BLACKSCREEN.Hide
      DEMOFLE1.Show
      MENUDEMO.Show
      Exit Sub
   End If
   '
   SECUENDE.List1.ListIndex = IIMAGEN% - 1
   '
   TTXX$ = SECUENDE.List1.List(IIMAGEN% - 1)
   ARIMAGE$ = Trim$(Mid$(TTXX$, 91, 16))
   LEYENDA$ = Trim$(Mid$(TTXX$, 11, 80)) + " [" + Trim$(Str$(IIMAGEN%)) + "]"
   INSTRU$ = Trim$(Left$(TTXX$, 8))
   PAUSI% = Val(Mid$(TTXX$, 107, 4))
   PAUSF% = Val(Mid$(TTXX$, 111, 4))
   MABRE% = Val(Mid$(TTXX$, 115, 4))
   '
   If InStr(ARIMAGE$, ".TXT") > 0 Then
     '
     CAPINU% = 0
     For KKII% = 1 To IIMAGEN%
       TTYZ$ = SECUENDE.List1.List(KKII% - 1)
       ARIMYZ$ = Trim$(Mid$(TTYZ$, 91, 16))
       If InStr(ARIMYZ$, ".TXT") > 0 Then
         CAPINU% = CAPINU% + 1
       End If
     Next KKII%
     '
     ENCAPITU.Label1.Caption = "Capítulo " + Trim$(Str$(CAPINU%)) + ":"
     LEYENDA$ = Trim$(Mid$(TTXX$, 11, 80))
18   For KKI% = 1 To Len(LEYENDA$)
       If Mid$(LEYENDA$, KKI%, 1) = "\" Then
         LEYENDA$ = Left$(LEYENDA$, KKI% - 1) + Chr$(10) + Chr$(13) + Mid$(LEYENDA$, KKI% + 1)
         GoTo 18
       End If
     Next KKI%
     ENCAPITU.Label2.Caption = LEYENDA$
     If PAUSF% < 3 Then PAUSF% = 3
     ENCAPITU.Label5.Caption = Str$(Timer + (PAUSI% + PAUSF%) * 100 / FORDEMO.HScroll1.Value)
     FORDEMO.Hide
     ENCAPITU.Show 1
     GoTo 10
   End If
   '
20 Call SHOWIMAGEN(ARIMAGE$, LEYENDA$, INSTRU$, PAUSI%, PAUSF%, MABRE%)
   '
22 If COMACON$ <> "" Then
     If COMACON$ = "HOME" Then
        COMACON$ = ""
        IIMAGEN% = 0
        GoTo 10
     ElseIf COMACON$ = "NEXT" Then
        COMACON$ = ""
        GoTo 10
     ElseIf COMACON$ = "PREV" Then
        COMACON$ = ""
        IIMAGEN% = IIMAGEN% - 1
        GoTo 15
     ElseIf COMACON$ = "POINT" Then
        COMACON$ = ""
        SECUENDE.List1.ListIndex = IIMAGEN% - 1
        SECUENDE.Show 1
        IIMAGEN% = SECUENDE.List1.ListIndex + 1
        If IIMAGEN% < 1 Then
          IIMAGEN% = ULTIMAGEN%
          GoTo 10
        End If
        GoTo 15
     ElseIf COMACON$ = "WAIT" Then
        COMACON$ = ""
        FORDEMO.Image2(3).ToolTipText = "En Pausa - Reanuda Ejecución"
        While COMACON$ = ""
          DoEvents
        Wend
        FORDEMO.Image2(3).ToolTipText = "Pausa de Ejecución"
        If COMACON$ = "WAIT" Then
          COMACON$ = ""
          GoTo 20
        End If
        GoTo 22
     ElseIf COMACON$ = "END" Then
        COMACON$ = ""
        IIMAGEN% = ULTIMAGEN%
        GoTo 10
     End If
   End If
   GoTo 10
   '
End Sub


Function CVI(A$)
    '
    X = 0
    CVI = 0
    If Len(A$) >= 2 Then
       X = 256& * Asc(Mid$(A$, 2)) + Asc(Left$(A$, 1))
       If X > 32767 Then X = X - 65536
    End If
    If X <= 32767 Then
       If X >= (-32768) Then
           CVI = X
       End If
    End If
    '
End Function
'

Function CVS(A$)
    '
    CVS = 0
    If Len(A$) < 4 Then Exit Function
    '
    Dim AA(4)
    Dim CCC As Single
    For K% = 1 To 4: AA(K%) = Asc(Mid$(A$, K%)): Next K%
    SIGNO = Sgn(128 - AA(3))
    If SIGNO = 0 Then SIGNO = (-1)
    If AA(4) = 0 Then SIGNO = 0
    '
    A41 = Int((AA(4) - 129) / 2): A42 = (AA(4) - 129) - A41
    N4 = 2 ^ A42
    N4 = N4 * 2 ^ A41
    N4 = N4 * SIGNO
    '
    A31 = AA(3) Mod 128
    N3 = A31 / 128: N3 = N3 * N4
    N2 = AA(2) / 128 / 256: N2 = N2 * N4
    N1 = AA(1) / 128 / 256 / 256: N1 = N1 * N4
    '
    CCC = N4 + N3 + N2 + N1
    CVS = CCC
    '
End Function
'

Function CVD(A$)
    '
    CVD = 0
    If Len(A$) < 8 Then Exit Function
    '
    Dim AA(8)
    Dim CCC As Double
    For K% = 1 To 8: AA(K%) = Asc(Mid$(A$, K%)): Next K%
    SIGNO = Sgn(128 - AA(7))
    If SIGNO = 0 Then SIGNO = (-1)
    If AA(8) = 0 Then SIGNO = 0
    '
    A41 = Int((AA(8) - 129) / 2): A42 = (AA(8) - 129) - A41
    N8# = 2 ^ A42
    N8# = N8# * 2 ^ A41
    N8# = N8# * SIGNO
    '
    A31 = AA(7) Mod 128
    N7# = A31 / 128: N7# = N7# * N8#
    N6# = AA(6) / 128 / 256: N6# = N6# * N8#
    N5# = AA(5) / 128 / 256 / 256: N5# = N5# * N8#
    N4# = AA(4) / 128 / 256 / 256 / 256: N4# = N4# * N8#
    N3# = AA(3) / 128 / 256 / 256 / 256 / 256: N3# = N3# * N8#
    N2# = AA(2) / 128 / 256 / 256 / 256 / 256 / 256: N2# = N2# * N8#
    N1# = AA(1) / 128 / 256 / 256 / 256 / 256 / 256 / 256: N1# = N1# * N8#
    '
    CCC# = N8# + N7# + N6# + N5# + N4# + N3# + N2# + N1#
    CVD = CCC#
    '
End Function
'
Function ATRILE$(TEBUS$)
   '
   ATRX$ = ""
   PPXX% = InStr(ATRIBU$, TEBUS$ + "=")
   If PPXX% > 0 Then
     PPYY% = InStr(PPXX% + 1, ATRIBU$, "\")
     If PPYY% <= PPXX% Then PPYY% = 1 + Len(ATRIBU$)
     ATRIPAR$ = Mid$(ATRIBU$, PPXX%, PPYY% - PPXX%)
     ATRX$ = Mid$(ATRIPAR$, 2 + Len(TEBUS$))
   End If
   '
   ATRILE$ = ATRX$
   '
End Function

Sub RELEYENDA()
   '
   FORDEMO.Label2(ORDEGUARDA%).Caption = INFODEM.Caption
   FORDEMO.Picture2(ORDEGUARDA%).Height = INFODEM.Height '+ 60
   FORDEMO.Picture2(ORDEGUARDA%).Width = INFODEM.Width '+ 30
   FORDEMO.Picture2(ORDEGUARDA%).Top = INFODEM.Top + 10 - FORDEMO.Top
   FORDEMO.Picture2(ORDEGUARDA%).Left = INFODEM.Left + 10 - FORDEMO.Left
   FORDEMO.Label3(ORDEGUARDA%).Width = INFODEM.Label3.Width
   FORDEMO.Label3(ORDEGUARDA%).Left = 50
   FORDEMO.Label3(ORDEGUARDA%).Font = INFODEM.Label3.Font
   FORDEMO.Label3(ORDEGUARDA%).FontSize = INFODEM.Label3.FontSize
   FORDEMO.Label3(ORDEGUARDA%).FontBold = INFODEM.Label3.FontBold
   FORDEMO.Label3(ORDEGUARDA%).FontItalic = INFODEM.Label3.FontItalic
   FORDEMO.Label3(ORDEGUARDA%).FontUnderline = INFODEM.Label3.FontUnderline
   FORDEMO.Label3(ORDEGUARDA%).Alignment = INFODEM.Label3.Alignment
   FORDEMO.Label3(ORDEGUARDA%).Caption = INFODEM.Label3.Caption
   FORDEMO.Label3(ORDEGUARDA%).ForeColor = INFODEM.Label3.ForeColor
   FORDEMO.Line3(ORDEGUARDA%).Y1 = FORDEMO.Picture2(ORDEGUARDA%).ScaleHeight - 20
   FORDEMO.Line3(ORDEGUARDA%).Y2 = FORDEMO.Picture2(ORDEGUARDA%).ScaleHeight - 20
   FORDEMO.Line4(ORDEGUARDA%).X1 = FORDEMO.Picture2(ORDEGUARDA%).ScaleWidth - 20
   FORDEMO.Line4(ORDEGUARDA%).X2 = FORDEMO.Picture2(ORDEGUARDA%).ScaleWidth - 20
   FORDEMO.Picture2(ORDEGUARDA%).BackColor = INFODEM.Picture2.BackColor
   FORDEMO.Picture2(ORDEGUARDA%).Visible = True
   '
End Sub

Sub CARLISECUEN(SECU$)
   '
   'SECUENDE.LIST1.Visible = True
   'SECUENDE.LIST1.Width = 11500
   'SECUENDE.LIST1.Height = 6000
   SECUENDE.List1.Clear
   '
   SECUEN$ = Trim$(UCase$(SECU$))
   If SECUEN$ = "" Then Exit Sub
   ASECU$ = SECUEN$ + ".SEC"
   LOSECU& = 0
   On Error Resume Next
   LOSECU& = FileLen(DEMOPATH$ + "\SCRIPTS\" + ASECU$)
   On Error GoTo 0
   If LOSECU& < 128 Then
     MsgBox "Secuencia '" + ASECU$ + "' no Instalada."
     Exit Sub
   End If
   '
   N1% = FreeFile
   Dim BUFLE1 As String * 256
   Open DEMOPATH$ + "\SCRIPTS\" + ASECU$ For Random Access Read Shared As #N1% Len = 256
   '
   UFECHA# = 0
   For II% = 1 To LOF(N1%) / 256
      Get #N1%, II%, BUFLE1$
      INSTRU$ = Trim$(UCase$(Left$(BUFLE1$, 8)))
      ARIMA$ = Trim$(UCase$(Mid$(BUFLE1$, 9, 16)))
      LEYEN$ = Trim$(Mid$(BUFLE1$, 25, 80))
      PAUSI% = CVI(Mid$(BUFLE1$, 105, 2))
      PAUSF% = CVI(Mid$(BUFLE1$, 107, 2))
      MABRE% = CVI(Mid$(BUFLE1$, 109, 2))
      '
      If Trim$(INSTRU$) <> "" Then
        TTXX$ = Space$(256)
        Mid$(TTXX$, 1, 8) = INSTRU$
        Mid$(TTXX$, 11, 80) = LEYEN$
        Mid$(TTXX$, 91, 16) = ARIMA$
        Mid$(TTXX$, 107, 4) = Trim$(Str$(PAUSI%))
        Mid$(TTXX$, 111, 4) = Trim$(Str$(PAUSF%))
        Mid$(TTXX$, 115, 4) = Trim$(Str$(MABRE%))
        SECUENDE.List1.AddItem TTXX$
      End If
      On Error Resume Next
      AFECHA# = FileDateTime(DEMOPATH$ + "\SCREENS\" + ARIMA$)
      On Error GoTo 0
      If AFECHA# < Now Then
        If AFECHA# > UFECHA# Then
          UFECHA# = AFECHA#
        End If
      End If
   Next II%
   '
   Close #N1%
   FEDICION$ = Format(UFECHA#, "dd/mm/yyyy") '
   '
10 AGREGA% = 0
   For JJ% = 1 To SECUENDE.List1.ListCount
     TTYY$ = SECUENDE.List1.List(JJ% - 1)
     ARCHIX$ = UCase$(Trim$(Mid$(TTYY$, 91, 16)))
     If InStr(ARCHIX$, ".SEC") > 1 Then
       SECUENDE.List1.RemoveItem JJ% - 1
       POINSE% = JJ% - 1
       Open DEMOPATH$ + "\SCRIPTS\" + ARCHIX$ For Random Access Read Shared As #N1% Len = 256
       '
       For II% = 1 To LOF(N1%) / 256
         Get #N1%, II%, BUFLE1$
         INSTRU$ = Trim$(UCase$(Left$(BUFLE1$, 8)))
         ARIMA$ = Trim$(UCase$(Mid$(BUFLE1$, 9, 16)))
         LEYEN$ = Trim$(Mid$(BUFLE1$, 25, 80))
         PAUSI% = CVI(Mid$(BUFLE1$, 105, 2))
         PAUSF% = CVI(Mid$(BUFLE1$, 107, 2))
         MABRE% = CVI(Mid$(BUFLE1$, 109, 2))
         '
         If Trim$(INSTRU$) <> "" Then
           TTXX$ = Space$(256)
           Mid$(TTXX$, 1, 8) = INSTRU$
           Mid$(TTXX$, 11, 80) = LEYEN$
           Mid$(TTXX$, 91, 16) = ARIMA$
           Mid$(TTXX$, 107, 4) = Trim$(Str$(PAUSI%))
           Mid$(TTXX$, 111, 4) = Trim$(Str$(PAUSF%))
           Mid$(TTXX$, 115, 4) = Trim$(Str$(MABRE%))
           SECUENDE.List1.AddItem TTXX$, POINSE%
           POINSE% = POINSE% + 1
           AGREGA% = 1
         End If
       Next II%
       Close #N1%
     End If
     If AGREGA% > 0 Then GoTo 10
     '
   Next JJ%
   '
End Sub
'
Sub DIBUFLECHA()
   '
   FORDEMO.TRAMO1(ORDEGUARDA%).Visible = False
   FORDEMO.TRAMO2(ORDEGUARDA%).Visible = False
   FORDEMO.TRAMO3(ORDEGUARDA%).Visible = False
   FORDEMO.PUNTA(ORDEGUARDA%).Visible = False
   FORDEMO.PUNTB(ORDEGUARDA%).Visible = False
   '
   FORDEMO.SOMBRA1(ORDEGUARDA%).Visible = False
   FORDEMO.SOMBRA2(ORDEGUARDA%).Visible = False
   FORDEMO.SOMBRA3(ORDEGUARDA%).Visible = False
   
   BIZQ = INFODEM.Left - FORDEMO.Left
   BSUP = INFODEM.Top - FORDEMO.Top
   ALTU = INFODEM.Height
   ANCH = INFODEM.Width
   APAR = 500
   ASOMV = 0: ASOMH = 0
   LARFLE = 180
   '
   FORDEMO.TRAMO1(ORDEGUARDA%).BorderColor = COLOFLE
   FORDEMO.TRAMO2(ORDEGUARDA%).BorderColor = COLOFLE
   FORDEMO.TRAMO3(ORDEGUARDA%).BorderColor = COLOFLE
   FORDEMO.PUNTA(ORDEGUARDA%).BorderColor = RGB(0, 0, 0) 'colofle
   FORDEMO.PUNTB(ORDEGUARDA%).BorderColor = RGB(0, 0, 0) 'COLOFLE
   FORDEMO.TRAMO1(ORDEGUARDA%).BorderWidth = 2
   FORDEMO.TRAMO2(ORDEGUARDA%).BorderWidth = 2
   FORDEMO.TRAMO3(ORDEGUARDA%).BorderWidth = 2
   FORDEMO.PUNTA(ORDEGUARDA%).BorderWidth = 2
   FORDEMO.PUNTB(ORDEGUARDA%).BorderWidth = 2
   '
   FORDEMO.SOMBRA1(ORDEGUARDA%).BorderColor = RGB(0, 0, 0) 'COLOFLE
   FORDEMO.SOMBRA2(ORDEGUARDA%).BorderColor = RGB(0, 0, 0)
   FORDEMO.SOMBRA3(ORDEGUARDA%).BorderColor = RGB(0, 0, 0)
   FORDEMO.SOMBRA1(ORDEGUARDA%).BorderWidth = 4
   FORDEMO.SOMBRA2(ORDEGUARDA%).BorderWidth = 4
   FORDEMO.SOMBRA3(ORDEGUARDA%).BorderWidth = 4
   '
   If TIPFLE% < 1 Or TIPFLE% > 3 Then TIPFLE% = 3
   '
   If TIPFLE% = 1 Then
      If ORIFLE% > 0 Then
        Select Case ORIFLE%
          Case 1
            FORDEMO.TRAMO3(ORDEGUARDA%).X1 = BIZQ + ANCH - 200
            FORDEMO.TRAMO3(ORDEGUARDA%).Y1 = BSUP + ALTU / 2
          Case 2
            FORDEMO.TRAMO3(ORDEGUARDA%).X1 = BIZQ + ANCH / 2
            FORDEMO.TRAMO3(ORDEGUARDA%).Y1 = BSUP + ALTU - 200
          Case 3
            FORDEMO.TRAMO3(ORDEGUARDA%).X1 = BIZQ + 200
            FORDEMO.TRAMO3(ORDEGUARDA%).Y1 = BSUP + ALTU / 2
          Case 4
            FORDEMO.TRAMO3(ORDEGUARDA%).X1 = BIZQ + ANCH / 2
            FORDEMO.TRAMO3(ORDEGUARDA%).Y1 = BSUP + 200
        End Select
        FORDEMO.TRAMO3(ORDEGUARDA%).X2 = XFLECH%
        FORDEMO.TRAMO3(ORDEGUARDA%).Y2 = YFLECH%
        FORDEMO.TRAMO3(ORDEGUARDA%).Visible = True
      End If
   End If
   '
   If TIPFLE% = 2 Then
      If ORIFLE% > 0 Then
        Select Case ORIFLE%
          Case 1
            FORDEMO.TRAMO2(ORDEGUARDA%).X1 = BIZQ + ANCH - 200
            FORDEMO.TRAMO2(ORDEGUARDA%).Y1 = BSUP + ALTU / 2
            FORDEMO.TRAMO2(ORDEGUARDA%).X2 = XFLECH%
            FORDEMO.TRAMO2(ORDEGUARDA%).Y2 = FORDEMO.TRAMO2(ORDEGUARDA%).Y1
          Case 2
            FORDEMO.TRAMO2(ORDEGUARDA%).X1 = BIZQ + ANCH / 2
            FORDEMO.TRAMO2(ORDEGUARDA%).Y1 = BSUP + ALTU - 200
            FORDEMO.TRAMO2(ORDEGUARDA%).X2 = FORDEMO.TRAMO2(ORDEGUARDA%).X1
            FORDEMO.TRAMO2(ORDEGUARDA%).Y2 = YFLECH%
          Case 3
            FORDEMO.TRAMO2(ORDEGUARDA%).X1 = BIZQ + 200
            FORDEMO.TRAMO2(ORDEGUARDA%).Y1 = BSUP + ALTU / 2
            FORDEMO.TRAMO2(ORDEGUARDA%).X2 = XFLECH%
            FORDEMO.TRAMO2(ORDEGUARDA%).Y2 = FORDEMO.TRAMO2(ORDEGUARDA%).Y1
          Case 4
            FORDEMO.TRAMO2(ORDEGUARDA%).X1 = BIZQ + ANCH / 2
            FORDEMO.TRAMO2(ORDEGUARDA%).Y1 = BSUP + 200
            FORDEMO.TRAMO2(ORDEGUARDA%).X2 = FORDEMO.TRAMO2(ORDEGUARDA%).X1
            FORDEMO.TRAMO2(ORDEGUARDA%).Y2 = YFLECH%
        End Select
        FORDEMO.TRAMO3(ORDEGUARDA%).X1 = FORDEMO.TRAMO2(ORDEGUARDA%).X2
        FORDEMO.TRAMO3(ORDEGUARDA%).Y1 = FORDEMO.TRAMO2(ORDEGUARDA%).Y2
        FORDEMO.TRAMO3(ORDEGUARDA%).X2 = XFLECH%
        FORDEMO.TRAMO3(ORDEGUARDA%).Y2 = YFLECH%
        FORDEMO.TRAMO2(ORDEGUARDA%).Visible = True
        FORDEMO.TRAMO3(ORDEGUARDA%).Visible = True
      End If
   End If
   '
   If TIPFLE% = 3 Then
      If ORIFLE% > 0 Then
        Select Case ORIFLE%
          Case 1
            FORDEMO.TRAMO1(ORDEGUARDA%).X1 = BIZQ + ANCH - 200
            FORDEMO.TRAMO1(ORDEGUARDA%).Y1 = BSUP + ALTU / 2
            FORDEMO.TRAMO1(ORDEGUARDA%).X2 = BIZQ + ANCH + APAR
            FORDEMO.TRAMO1(ORDEGUARDA%).Y2 = FORDEMO.TRAMO1(ORDEGUARDA%).Y1
            FORDEMO.TRAMO3(ORDEGUARDA%).X1 = FORDEMO.TRAMO1(ORDEGUARDA%).X2
            FORDEMO.TRAMO3(ORDEGUARDA%).Y1 = YFLECH%
          Case 2
            FORDEMO.TRAMO1(ORDEGUARDA%).X1 = BIZQ + ANCH / 2
            FORDEMO.TRAMO1(ORDEGUARDA%).Y1 = BSUP + ALTU - 200
            FORDEMO.TRAMO1(ORDEGUARDA%).X2 = FORDEMO.TRAMO1(ORDEGUARDA%).X1
            FORDEMO.TRAMO1(ORDEGUARDA%).Y2 = BSUP + ALTU + APAR
            FORDEMO.TRAMO3(ORDEGUARDA%).X1 = XFLECH%
            FORDEMO.TRAMO3(ORDEGUARDA%).Y1 = FORDEMO.TRAMO1(ORDEGUARDA%).Y2
          Case 3
            FORDEMO.TRAMO1(ORDEGUARDA%).X1 = BIZQ + 200
            FORDEMO.TRAMO1(ORDEGUARDA%).Y1 = BSUP + ALTU / 2
            FORDEMO.TRAMO1(ORDEGUARDA%).X2 = BIZQ - APAR
            FORDEMO.TRAMO1(ORDEGUARDA%).Y2 = FORDEMO.TRAMO1(ORDEGUARDA%).Y1
            FORDEMO.TRAMO3(ORDEGUARDA%).X1 = FORDEMO.TRAMO1(ORDEGUARDA%).X2
            FORDEMO.TRAMO3(ORDEGUARDA%).Y1 = YFLECH%
         Case 4
            FORDEMO.TRAMO1(ORDEGUARDA%).X1 = BIZQ + ANCH / 2
            FORDEMO.TRAMO1(ORDEGUARDA%).Y1 = BSUP + 200
            FORDEMO.TRAMO1(ORDEGUARDA%).X2 = FORDEMO.TRAMO1(ORDEGUARDA%).X1
            FORDEMO.TRAMO1(ORDEGUARDA%).Y2 = BSUP - APAR
            FORDEMO.TRAMO3(ORDEGUARDA%).X1 = XFLECH%
            FORDEMO.TRAMO3(ORDEGUARDA%).Y1 = FORDEMO.TRAMO1(ORDEGUARDA%).Y2
        End Select
        FORDEMO.TRAMO3(ORDEGUARDA%).X2 = XFLECH%
        FORDEMO.TRAMO3(ORDEGUARDA%).Y2 = YFLECH%
        FORDEMO.TRAMO2(ORDEGUARDA%).X1 = FORDEMO.TRAMO1(ORDEGUARDA%).X2
        FORDEMO.TRAMO2(ORDEGUARDA%).Y1 = FORDEMO.TRAMO1(ORDEGUARDA%).Y2
        FORDEMO.TRAMO2(ORDEGUARDA%).X2 = FORDEMO.TRAMO3(ORDEGUARDA%).X1
        FORDEMO.TRAMO2(ORDEGUARDA%).Y2 = FORDEMO.TRAMO3(ORDEGUARDA%).Y1
        FORDEMO.TRAMO1(ORDEGUARDA%).Visible = True
        FORDEMO.TRAMO2(ORDEGUARDA%).Visible = True
        FORDEMO.TRAMO3(ORDEGUARDA%).Visible = True
      End If
   End If
   '
   XORI% = FORDEMO.TRAMO3(ORDEGUARDA%).X1
   YORI% = FORDEMO.TRAMO3(ORDEGUARDA%).Y1
   ORIENTA = 90
   If Abs(XORI% - XFLECH%) > 10 Then
     ORIENTA = 360 * Atn((YFLECH% - YORI%) / (XORI% - XFLECH%)) / 6.283
   End If
   If XORI% > XFLECH% Then ORIENTA = ORIENTA + 180
   If ORIENTA = 90 Then
     If YORI% < YFLECH% Then ORIENTA = ORIENTA + 180
   End If
   FORDEMO.TRAMO3(ORDEGUARDA%).X2 = XFLECH% - 70 * Cos(6.28 * ORIENTA / 360)
   FORDEMO.TRAMO3(ORDEGUARDA%).Y2 = YFLECH% + 70 * Sin(6.28 * ORIENTA / 360)
   '
   If FORDEMO.TRAMO1(ORDEGUARDA%).Visible = True Then
     FORDEMO.SOMBRA1(ORDEGUARDA%).X1 = FORDEMO.TRAMO1(ORDEGUARDA%).X1 + ASOMH
     FORDEMO.SOMBRA1(ORDEGUARDA%).Y1 = FORDEMO.TRAMO1(ORDEGUARDA%).Y1 + ASOMV
     FORDEMO.SOMBRA1(ORDEGUARDA%).X2 = FORDEMO.TRAMO1(ORDEGUARDA%).X2 + ASOMH
     FORDEMO.SOMBRA1(ORDEGUARDA%).Y2 = FORDEMO.TRAMO1(ORDEGUARDA%).Y2 + ASOMV
     FORDEMO.SOMBRA1(ORDEGUARDA%).Visible = True
   End If
   '
   If FORDEMO.TRAMO2(ORDEGUARDA%).Visible = True Then
     FORDEMO.SOMBRA2(ORDEGUARDA%).X1 = FORDEMO.TRAMO2(ORDEGUARDA%).X1 + ASOMH
     FORDEMO.SOMBRA2(ORDEGUARDA%).Y1 = FORDEMO.TRAMO2(ORDEGUARDA%).Y1 + ASOMV
     FORDEMO.SOMBRA2(ORDEGUARDA%).X2 = FORDEMO.TRAMO2(ORDEGUARDA%).X2 + ASOMH
     FORDEMO.SOMBRA2(ORDEGUARDA%).Y2 = FORDEMO.TRAMO2(ORDEGUARDA%).Y2 + ASOMV
     FORDEMO.SOMBRA2(ORDEGUARDA).Visible = True
   End If
   '
   If FORDEMO.TRAMO3(ORDEGUARDA%).Visible = True Then
     FORDEMO.SOMBRA3(ORDEGUARDA%).X1 = FORDEMO.TRAMO3(ORDEGUARDA%).X1 + ASOMH
     FORDEMO.SOMBRA3(ORDEGUARDA%).Y1 = FORDEMO.TRAMO3(ORDEGUARDA%).Y1 + ASOMV
     FORDEMO.SOMBRA3(ORDEGUARDA%).X2 = FORDEMO.TRAMO3(ORDEGUARDA%).X2 + ASOMH
     FORDEMO.SOMBRA3(ORDEGUARDA%).Y2 = FORDEMO.TRAMO3(ORDEGUARDA%).Y2 + ASOMV
     FORDEMO.SOMBRA3(ORDEGUARDA).Visible = True
     '
     XORI% = FORDEMO.TRAMO3(ORDEGUARDA%).X1
     YORI% = FORDEMO.TRAMO3(ORDEGUARDA%).Y1
     ORIENTA = 90
     If Abs(XORI% - XFLECH%) > 10 Then
       ORIENTA = 360 * Atn((YFLECH% - YORI%) / (XORI% - XFLECH%)) / 6.283
     End If
     If XORI% > XFLECH% Then ORIENTA = ORIENTA + 180
     If ORIENTA = 90 Then
       If YORI% < YFLECH% Then ORIENTA = ORIENTA + 180
     End If
     '
     FORDEMO.PUNTA(ORDEGUARDA%).X1 = XFLECH%
     FORDEMO.PUNTA(ORDEGUARDA%).Y1 = YFLECH%
     FORDEMO.PUNTA(ORDEGUARDA%).X2 = XFLECH% - LARFLE * Cos(6.283 * (ORIENTA - 25) / 360)
     FORDEMO.PUNTA(ORDEGUARDA%).Y2 = YFLECH% + LARFLE * Sin(6.283 * (ORIENTA - 25) / 360)
     FORDEMO.PUNTA(ORDEGUARDA%).Visible = True
     '
     FORDEMO.PUNTB(ORDEGUARDA%).X1 = XFLECH%
     FORDEMO.PUNTB(ORDEGUARDA%).Y1 = YFLECH%
     FORDEMO.PUNTB(ORDEGUARDA%).X2 = XFLECH% - LARFLE * Cos(6.283 * (ORIENTA + 25) / 360)
     FORDEMO.PUNTB(ORDEGUARDA%).Y2 = YFLECH% + LARFLE * Sin(6.283 * (ORIENTA + 25) / 360)
     FORDEMO.PUNTB(ORDEGUARDA%).Visible = True
     '
   End If
   '
End Sub
'
Sub CENTRAFORM(NOFORMU As Form)
   '
   If Screen.Height > NOFORMU.Height + 600 Then
     NOFORMU.Top = (Screen.Height - NOFORMU.Height) / 2
   End If
   If Screen.Width > NOFORMU.Width Then
     NOFORMU.Left = (Screen.Width - NOFORMU.Width) / 2
   End If
   DoEvents
   '
End Sub

