Attribute VB_Name = "SELECH2001"
Public TASEL As String
Public ITABLA%
Dim ARCHU$
Public ARCHA$(15)
Public ASELEC$(15)
Public ABUSCA$(15)
Public DENARCH$(15)
Public ALIS(15)
Public IACTI&(15)
Dim UACTU#(15)
'
Public VD1$(15), VD2$(15)
Public ACOL1%(15), ACOL2%(15)
Public TIVA1%(15), TIVA2%(15)
'
Public TV1%(15), TV2%(15)
Public LF1%(15), LF2%(15)
Public PR1%(15), PR2%(15)
Public AC1%(15), AC2%(15)
Public JT1%(15), JT2%(15)
Public LD1%(15), LD2%(15)
Public AVEN%(15)
Public TEBUS$(15)
'
Public REGTOPE&(15), REGACTI&(15)
Public CONTROBARRA%

Dim IHOO As Single
Dim IARCHI As Integer
'
Dim TTXX As String * 32767
Dim TABLA_ARCHIVO As String
Dim TABLA_TITULO As String
Dim TABLA_INDICE As Integer
'
Sub SELECHO(ARCHITITU$)
    '
    Call SEPARTITU(ARCHITITU$)  ' separa archivo y titulo
    Call OCULTABLAS             ' oculta todas las tablas
    '
    Call IARTABLA(TABLA_ARCHIVO)
    FORSELECHO.Caption = TABLA_TITULO
    '
    FORSELECHO.LINVER.Y1 = 0
    FORSELECHO.LINVER.Y2 = FORSELECHO.MARCOTA(ITABLA%).Height + 20
    FORSELECHO.LINVER.X1 = FORSELECHO.MARCOTA(ITABLA%).Width + 25
    FORSELECHO.LINVER.X2 = FORSELECHO.LINVER.X1
    FORSELECHO.SCROLLBA.Top = 0
    FORSELECHO.SCROLLBA.Left = FORSELECHO.MARCOTA(ITABLA%).Width + 56
    FORSELECHO.SCROLLBA.Height = FORSELECHO.MARCOTA(ITABLA%).Height
    TASELEC$ = ""
    FORSELECHO.Height = FORSELECHO.MARCOTA(ITABLA%).Height + 1000
    FORSELECHO.Width = FORSELECHO.SCROLLBA.Left + FORSELECHO.SCROLLBA.Width + 112
    '
    FORSELECHO.LINHOR.Y1 = FORSELECHO.MARCOTA(ITABLA%).Height + 20
    FORSELECHO.LINHOR.Y2 = FORSELECHO.LINHOR.Y1
    FORSELECHO.LINHOR.X1 = 0
    FORSELECHO.LINHOR.X2 = FORSELECHO.Width
    '
    CONTROBARRA% = (-1)
    FORSELECHO.SCROLLBA.Min = 1
    MASCROLL& = ULTREG&(ARCHU$) - 10: If MASCROLL& < 1 Then MASCROLL& = 1
    FORSELECHO.SCROLLBA.Max = MASCROLL&
    If REGTOPE&(ITABLA%) > 0 Then
      If REGTOPE&(ITABLA%) <= MASCROLL& Then
        FORSELECHO.SCROLLBA.Value = REGTOPE&(ITABLA%)
      End If
    End If
    FORSELECHO.SCROLLBA.LargeChange = 11
    CONTROBARRA = 0
    '
    FORSELECHO.MARCOTA(ITABLA%).Visible = True
    Call CARGATABLA
    FORSELECHO.TABLA(ITABLA%).Visible = True
    '
    FORSELECHO.Text1.TEXT = ABUSCA$(ITABLA%)
    If TRIM$(FORSELECHO.Text1.TEXT) <> "" Then
       FORSELECHO.TASEL(ITABLA%).Visible = True
       FORSELECHO.TABLA(ITABLA%).Visible = False
       FORSELECHO.SCROLLBA.Min = 1
       MASCROLL& = FORSELECHO.TASEL(ITABLA%).ListCount - 10
       If MASCROLL& < 1 Then MASCROLL& = 1
       FORSELECHO.SCROLLBA.Max = MASCROLL&
       If FORSELECHO.TASEL(ITABLA%).TopIndex >= 0 Then
         If FORSELECHO.TASEL(ITABLA%).TopIndex < MASCROLL& Then
           FORSELECHO.SCROLLBA.Value = FORSELECHO.TASEL(ITABLA%).TopIndex + 1
         End If
       End If
       Call CARTEBUS     ' ESTABA SUPRIMIDO HASTA 30/11/2001
       Call CARTASEL     ' ESTABA SUPRIMIDO HASTA 30/11/2001
    End If
    '
    Screen.MousePointer = 1
    FORSELECHO.Show 1
   Exit Sub
    '
    'ITABLA% = IARCHI%(ARCHI$)
    ARCHU$ = ARCHA$(ITABLA%)
    If ULTREG&(ARCHU$) < 1 Then
        Call MENSERR(24, "Tabla de Validación Vacía.")
        Exit Sub
    End If
    '
    If TRIM$(TTY$) = "" Then
        TTY$ = TITUARCH$(ARCHU$)
    End If
    '
130 FORMUSEL.Caption = TTY$
    FORMUSEL.Width = ALIS(ITABLA%) + 80
    FORMUSEL.TABLA(ITABLA%).Height = FORMUSEL.Text1.Top
    FORMUSEL.TABLA(ITABLA%).Width = ALIS(ITABLA%)
    If FORMUSEL.TABLA(ITABLA%).ListCount > 0 Then
       'If InStr(ARCHU$, "FECHA") < 1 Then
       '    FORMUSEL.TABLA(ITABLA%).ListIndex = 0
       '    If IACTI&(ITABLA%) >= 0 Then
       '        If IACTI&(ITABLA%) < FORMUSEL.TABLA(ITABLA%).ListCount Then
       '            FORMUSEL.TABLA(ITABLA%).ListIndex = IACTI&(ITABLA%)
       '        End If
       '    End If
       '  Else
       '    FORMUSEL.TABLA(ITABLA%).ListIndex = IHOO
       'End If
       'TOPELI& = FORMUSEL.TABLA(ITABLA%).ListIndex - 5
       'If TOPELI& < 0 Then TOPELI& = 0
       'FORMUSEL.TABLA(ITABLA%).TopIndex = TOPELI&
    End If
    '
    FORMUSEL.TABLA(ITABLA%).Visible = True
    FORMUSEL.Text1.TEXT = ASELEC$(ITABLA%)
    FORMUSEL.Refresh
    '
    FORMUSEL.TASEL(ITABLA%).Height = FORMUSEL.Text1.Top
    FORMUSEL.TASEL(ITABLA%).Width = ALIS(ITABLA%)
    FORMUSEL.TASEL(ITABLA%).BackColor = FORMUSEL.TASEL(0).BackColor
    '
    VD1$(ITABLA%) = ""
    VD2$(ITABLA%) = ""
    Screen.MousePointer = 1
    '
    SMP = Screen.MousePointer
    Screen.MousePointer = 1
    FORMUSEL.Show 1
    Screen.MousePointer = SMP
    Call FRESHALL
    '
End Sub


Function VALACT1$(ARCHI$)
    '
    'ITABLA% = IARCHX%(ARCHI$)
    VALACT1$ = VD1$(ITABLA%)
    '
End Function
Function VALACT2$(ARCHI$)
    '
    'ITABLA% = IARCHX%(ARCHI$)
    VALACT2$ = VD2$(ITABLA%)
    '
End Function

Sub FRESHECHO(ARCHI$)
    '
    'IAA% = IARCHX%(ARCHI$)
    'FORSELECHO.TABLA(ITABLA%).Clear
    'UACTU#(ITABLA%) = 0
    'Call CARTEBUS
    'Call CARTASEL
    '
End Sub

Function REGSEL$(ARCHITITU$)
    '
    XXX$ = "": YYY& = 0
    PPXX% = InStr(ARCHITITU$, "/")
    If PPXX% < 2 Then PPXX% = 1 + Len(ARCHITITU$)
    ARCHI$ = Left$(ARCHITITU$, PPXX% - 1)
    '
    Call SELECHO(ARCHITITU$)
    '
    If VALACT1$(ARCHI$) <> "" Then
        XXX$ = REGACT$(ARCHI$)
        YYY& = REGNUM&(ARCHI$)
    End If
    '
    REGSEL$ = XXX$ + MKS$(YYY&)
    '
End Function

Sub SEPARTITU(ARTTX$)
   '
   PPXX% = InStr(ARTTX$, "/")
   If PPXX% < 2 Then PPXX% = 1 + Len(ARTTX$)
   TABLA_ARCHIVO$ = Left$(ARTTX$, PPXX% - 1)
   TABLA_TITULO$ = Mid$(ARTTX$, PPXX% + 1)
   If TRIM$(TABLA_TITULO$) = "" Then
      TABLA_TITULO$ = TRIM$(TITUARCH$(TABLA_ARCHIVO$))
   End If
   '
End Sub

Sub OCULTABLAS()
   '
   For II = 0 To 15
      FORSELECHO.MARCOTA(II).Visible = False
   Next II
   '
End Sub

Sub IARTABLA(ARTTX$)
   '
   Static CARCHI%
   Static VECTORARCH$
   '
   ARCHU$ = TRIM$(UCase$(TABLA_ARCHIVO$))
   If ARCHU$ = "" Or Len(ARCHU$) > 8 Then
      Call MENSERR(24, "Error de Invocación de Rutina 'SELECHO' - Consulte Soporte de Sistemas.")
      Exit Sub
   End If
   '
   If ARCHU$ = "MASTER" Then
       ARCHU$ = "ECOMAST"
     ElseIf ARCHU$ = "FECHA" Then
       ARCHU$ = "SELFECHA"
   End If
   '
   If Left$(ARCHU$, 1) = "*" Then
        FRXX$ = Mid$(ARCHI$, 2) + ".PRF"
        Dim YYY As String * 128
        N% = FreeFile
        Open FRXX$ For Random Access Read Write Shared As #N% Len = 128
        Get #N%, 1, YYY$
        ARCHU$ = Mid$(YYY$, 57, 8)
        Close #N%: N% = 0
      ElseIf Left$(ARCHU$, 1) = "^" Then
        ARCHU$ = Mid$(ARCHU$, 2)
   End If
   '
   ARCHU10$ = AJUSTI$("@@" + ARCHU$, 10)
   PPXX% = InStr(VECTORARCH$, ARCHU10$)
   If (PPXX% Mod 10) = 1 Then
      IARCHI% = (PPXX% - 1) / 10
      ITABLA% = IARCHI%
      GoTo 99
   End If
   '
   If CARCHI% >= 16 Then
      Call MENSERR(24, "Demasiadas Tablas Abiertas.\Consulte al Soporte de Sistemas.")
      Call FINAL("")
   End If
   '
   CARCHI% = CARCHI% + 1
   ARCHA$(CARCHI% - 1) = ARCHU$
   VECTORARCH$ = VECTORARCH$ + ARCHU10$
   IARCHI% = CARCHI% - 1
   ITABLA% = IARCHI%
   '
10 If ARCHU$ = "SELFECHA" Then
     Call GENTAFECH
   End If
   '     TV1%(ITABLA%) = 5: TV2%(ITABLA%) = 6
   '     LF1%(ITABLA%) = 2: LF2%(ITABLA%) = 14
   '     PR1%(ITABLA%) = 1: PR2%(ITABLA%) = 3
   '   Else
        TV1%(ITABLA%) = TIPFIELD%(ARCHU$, 1): TV2%(ITABLA%) = TIPFIELD%(ARCHU$, 2)
        LF1%(ITABLA%) = LENFIELD%(ARCHU$, 1): LF2%(ITABLA%) = LENFIELD%(ARCHU$, 2)
        PR1%(ITABLA%) = POSFIELD%(ARCHU$, 1): PR2%(ITABLA%) = POSFIELD%(ARCHU$, 2)
   'End If
   '
   AC1%(ITABLA%) = LF1%(ITABLA%)
   JT1%(ITABLA%) = 1
   LD1%(ITABLA%) = 0
   If TV1%(ITABLA%) = 1 Or TV1%(ITABLA%) = 2 Then
              JT1%(ITABLA%) = 2
              If LF1%(ITABLA%) < 2 Then AC1%(ITABLA%) = 3 Else AC1%(ITABLA%) = 5
          ElseIf TV1%(ITABLA%) = 3 Or TV1%(ITABLA%) = 4 Then
              JT1%(ITABLA%) = 2: LD1%(ITABLA%) = 1
              If LF1%(ITABLA%) < 8 Then AC1%(ITABLA%) = 7 Else AC1%(ITABLA%) = 12
              If LD1%(ITABLA%) > 0 Then AC1%(ITABLA%) = AC1%(ITABLA%) + 2
          ElseIf TV1%(ITABLA%) = 5 Then
              AC1%(ITABLA%) = 8
          ElseIf TV1%(ITABLA%) = 7 Then
              AC1%(ITABLA%) = 16
    End If
    '
    ACOL1%(ITABLA%) = AC1%(ITABLA%)
    TIVA1%(ITABLA%) = TV1%(ITABLA%)
    '
    AC2%(ITABLA%) = LF2%(ITABLA%)
    JT2%(ITABLA%) = 1
    LD2%(ITABLA%) = 0
    If TV2%(ITABLA%) = 1 Or TV2%(ITABLA%) = 2 Then
              JT2%(ITABLA%) = 2
              If LF2%(ITABLA%) < 2 Then AC2%(ITABLA%) = 3 Else AC2%(ITABLA%) = 5
          ElseIf TV2%(ITABLA%) = 3 Or TV2%(ITABLA%) = 4 Then
              JT2%(ITABLA%) = 2
              LD2%(ITABLA%) = 2
              If LF2%(ITABLA%) < 8 Then AC2%(ITABLA%) = 7 Else AC2%(ITABLA%) = 12
              If LD2%(ITABLA%) > 0 Then AC2%(ITABLA%) = AC2%(ITABLA%) + 2
          ElseIf TV2%(ITABLA%) = 5 Then
              AC2%(ITABLA%) = 8
          ElseIf TV2%(ITABLA%) = 7 Then
              AC2%(ITABLA%) = 16
    End If
    '
    ACOL2%(ITABLA%) = AC2%(ITABLA%)
    '
    If AC1%(ITABLA%) + AC2%(ITABLA%) > 70 Then AC2%(ITABLA%) = 70 - AC1%(ITABLA%)
    AVEN%(ITABLA%) = AC1%(ITABLA%) + 2 + AC2%(ITABLA%) + 4
    If AVEN%(ITABLA%) > 72 Then AVEN%(ITABLA%) = 72
    FORSELECHO.TEPRUEBA.Caption = String$(AVEN%(ITABLA%), "*")
    ALIS(ITABLA%) = FORSELECHO.TEPRUEBA.Width
    '
99  FORSELECHO.MARCOTA(ITABLA%).Top = 0
    FORSELECHO.MARCOTA(ITABLA%).Left = 0
    FORSELECHO.MARCOTA(ITABLA%).Height = 11 * FORSELECHO.TEPRUEBA.Height
    FORSELECHO.MARCOTA(ITABLA).Width = ALIS(ITABLA)
    '
    FORSELECHO.TABLA(ITABLA).Top = 0
    FORSELECHO.TABLA(ITABLA%).Left = 0
    FORSELECHO.TABLA(ITABLA%).Height = 12 * FORSELECHO.TEPRUEBA.Height
    FORSELECHO.TABLA(ITABLA).Width = ALIS(ITABLA)
    '
    FORSELECHO.TASEL(ITABLA).Top = 0
    FORSELECHO.TASEL(ITABLA%).Left = 0
    FORSELECHO.TASEL(ITABLA%).Height = FORSELECHO.Text1.Top
    FORSELECHO.TASEL(ITABLA%).Width = ALIS(ITABLA%) + 300
    FORSELECHO.TASEL(ITABLA%).BackColor = FORSELECHO.TASEL(0).BackColor
    '
    FORSELECHO.BOTINSE.Top = FORSELECHO.Text1.Top - 10
    FORSELECHO.BOTERA.Top = FORSELECHO.Text1.Top - 10
    FORSELECHO.BOTCANCE.Top = FORSELECHO.Text1.Top - 10
    FORSELECHO.BOTCANCE.Left = ALIS(ITABLA%) - FORSELECHO.BOTCANCE.Width
    If ALIS(ITABLA%) < 3000 Then
        FORSELECHO.BOTERA.Width = 0
        FORSELECHO.BOTINSE.Width = 0
        FORSELECHO.BOTERA.Visible = False
        FORSELECHO.BOTINSE.Visible = False
    End If
    FORSELECHO.BOTERA.Left = FORSELECHO.BOTCANCE.Left - FORSELECHO.BOTERA.Width
    FORSELECHO.BOTINSE.Left = FORSELECHO.BOTERA.Left - FORSELECHO.BOTINSE.Width
    FORSELECHO.BOTINSE.Enabled = False
    FORSELECHO.BOTERA.Enabled = False
    '
    FORSELECHO.Text1.Width = FORSELECHO.BOTINSE.Left - FORSELECHO.Text1.Left - 100
    '
End Sub

Sub CARGATABLA()
   '
   REGINI& = REGTOPE&(ITABLA%)
   If REGINI& < 1 Then REGINI& = 1
   If REGACTI&(ITABLA%) < 1 Then REGACTI&(ITABLA%) = 1
   FORSELECHO.TABLA(ITABLA%).Clear
   For U& = REGINI& To REGINI& + 10
      If U& <= ULTREG&(ARCHU$) Then
          XXX$ = REGLEIDO$(ARCHU$, U&)
          XXZ$ = Left$(XXX$, LF1%(ITABLA%))
          TX1$ = CSTRING$(XXZ$, TV1%(ITABLA%), AC1%(ITABLA%), LD1%(ITABLA%), JT1%(ITABLA%))
          XXZ$ = Mid$(XXX$, PR2%(ITABLA%), LF2%(ITABLA%))
          TX2$ = CSTRING$(XXZ$, TV2%(ITABLA%), AC2%(ITABLA%), LD2%(ITABLA%), JT2%(ITABLA%))
          TX12$ = TX1$ + "  " + TX2$
          If TRIM$(TX12$) <> "" Then
              FORSELECHO.TABLA(ITABLA%).AddItem TX12$
          End If
      End If
   Next U&
   '
   If REGACTI&(ITABLA%) >= REGINI& Then
      If REGACTI&(ITABLA%) < REGINI& + FORSELECHO.TABLA(ITABLA%).ListCount Then
         FORSELECHO.TABLA(ITABLA%).ListIndex = REGACTI&(ITABLA%) - REGINI&
      End If
   End If
   FORSELECHO.TABLA(ITABLA%).Refresh
   '
   If FORSELECHO.TABLA(ITABLA%).ListCount > 0 Then
     If REGTOPE&(ITABLA%) < REGINI& Then
       REGTOPE&(ITABLA%) = REGINI&
     End If
   End If
   '
End Sub
'
Sub CARTEBUS()
   '
   On Error Resume Next
   TCX# = CDbl(FileDateTime(LUDAT$ + ARCHA$(ITABLA%) + ".DAT"))
   On Error GoTo 0
   If TCX# <> UACTU#(ITABLA%) Then TEBUS$(ITABLA%) = ""
   '
   If TEBUS$(ITABLA%) = "" Or UACTU#(ITABLA%) < 1 Then
     '
     Screen.MousePointer = 11
     FORSELECHO.TASEL(ITABLA%).Clear
     FORSELECHO.TABLA(ITABLA%).Visible = False
     FORSELECHO.TASEL(ITABLA%).Visible = True
     '
     NX% = FreeFile
     ARCHU$ = ARCHA$(ITABLA%)
     ADEUDO% = 0: If ARCHU$ = "DEUDOR1" Then ADEUDO% = 1
     On Error Resume Next
       TCX# = CDbl(FileDateTime(LUDAT$ + ARCHU$ + ".DAT"))
       UACTU#(ITABLA%) = TCX#
       If ADEUDO% = 0 Then TCY# = CDbl(FileDateTime(LUDAT$ + ARCHU$ + ".SEL"))
     On Error GoTo 0
     EXTE$ = ".DAT": If TCY# >= TCX# Then EXTE$ = ".SEL"
     '
' NUEVA RUTINA DE CARGA
     ALISTA% = AC1%(ITABLA%) + AC2%(ITABLA%) + 2
     URE& = ULTREG&(ARCHU$)
     Open LUDAT$ + ARCHU$ + EXTE$ For Binary Access Read Write Shared As #NX% Len = 1
     ALISLE& = URE& * ALISTA%: If EXTE$ = ".DAT" Then ALISLE& = LOF(NX%)
     VELEC$ = String$(ALISLE&, 0)
     Get #NX%, 1, VELEC$
     Close #NX%
     If EXTE$ = ".SEL" Then
       TEBUS$(ITABLA%) = VELEC$
       GoTo 99
     End If
     ' VIEJA RUTINA DE CARGA
     'Open LUDAT$ + ARCHU$ + EXTE$ For Random Access Read Write Shared As #NX% Len = 32767
     'CARELE& = 1 + LOF(NX%) / 32767
     'VELEC$ = String$(CARELE& * 32767, 0)
     'POSIN& = 1
     'For KK& = 1 To CARELE&
     '  Get #NX%, KK&, TTXX$
     '  Mid$(VELEC$, POSIN&, 32767) = TTXX$
     '  POSIN& = POSIN& + 32767
     'Next KK&
     'Close #NX%
     
     'LOFI = RECLEN%(ARCHU$)
     'ALISTA% = AC1%(ITABLA%) + AC2%(ITABLA%) + 2
     'URE& = ULTREG&(ARCHU$)
     'If EXTE$ = ".SEL" Then
     '   TEBUS$(ITABLA%) = Left$(VELEC$, URE& * ALISTA%)
     '   GoTo 99
     'End If
     '
     LOFI = RECLEN%(ARCHU$)
     TEBUS$(ITABLA%) = Space$(URE& * ALISTA%)
     POSIN& = 1
     '
     For U& = 1 To Len(VELEC$) Step LOFI
       If U& <= URE& * LOFI Then
         XXX$ = Mid$(VELEC$, U&, LOFI)
         XXZ$ = Left$(XXX$, LF1%(ITABLA%))
           If ADEUDO% = 1 Then If CVI(XXZ$) < 1 Then GoTo 19
         TX1$ = CSTRING$(XXZ$, TV1%(ITABLA%), AC1%(ITABLA%), LD1%(ITABLA%), JT1%(ITABLA%))
         XXZ$ = Mid$(XXX$, PR2%(ITABLA%), LF2%(ITABLA%))
         TX2$ = CSTRING$(XXZ$, TV2%(ITABLA%), AC2%(ITABLA%), LD2%(ITABLA%), JT2%(ITABLA%))
         Mid$(TEBUS$(ITABLA%), POSIN&, ALISTA%) = TX1$ + "  " + TX2$
         POSIN& = POSIN& + ALISTA%
       End If
19   Next U&
     '
     NX% = FreeFile
     Open LUDAT$ + ARCHU$ + ".SEL" For Binary As #NX%
     Put #NX%, , TEBUS$(ITABLA%)
     Close #NX%
     '
99   Screen.MousePointer = 1
     '
   End If
   '
End Sub
'
Sub GENTAFECH()
   '
   On Error Resume Next
     TCX# = CDbl(FileDateTime(LUDAT$ + "SELFECHA.DAT"))
     If FileLen(LUDAT$ + "SELFECHA.DAT") < 1 Then TCX# = 0
     TCY# = CDbl(Date)
   On Error GoTo 0
   '
   If Int(TCX#) <> Int(TCY#) Then
      Screen.MousePointer = 11
      HOOI% = HOY(HO$)
      ANX = Val(Mid$(HO$, 7, 2)) - 1: If ANX < 0 Then ANX = ANX + 100
      ANXS$ = TRIM$(Str$(ANX)): While Len(ANXS$) < 2: ANXS$ = "0" + ANXS$: Wend
      DESDE$ = Left$(HO$, 6) + ANXS$
      ANX = Val(Mid$(HO$, 7, 2)) + 1: If ANX > 99 Then ANX = ANX - 100
      ANXS$ = TRIM$(Str$(ANX)): While Len(ANXS$) < 2: ANXS$ = "0" + ANXS$: Wend
      HASTA$ = Left$(HO$, 6) + ANXS$
      '
      JJ& = 0
      HASI% = FECHANUM(HASTA$)
      FEXI% = DIPOST(FECHANUM(DESDE$))
      While FEXI% <= HASI%
        FEX = FEXI%
        TX2$ = DSEM$(FEX)
        TX12$ = MKI$(FEXI%) + TX2$
        If TRIM$(TX12$) <> "" Then
          JJ& = JJ& + 1
          Call GRAREG("SELFECHA", TX12$, JJ&)
        End If
        FEY = DIPOST(FEX)
        FEXI% = FEY
      Wend
      Call SETULTREG("SELFECHA", JJ&)
      Call CIERRARCH("SELFECHA")
      Screen.MousePointer = 1
   End If
   '
   Screen.MousePointer = 11
   HOOI% = HOY(HO$)
   For U& = 1 To ULTREG&("SELFECHA")
     TX12$ = REGLEIDO$("SELFECHA", U&)
     If CVI(Left$(TX12$, 2)) = HOOI% Then
       REGACTI&(ITABLA%) = U&
       GoTo 99
     End If
   Next U&
   '
99 REGTOPE&(ITABLA%) = 1
   If REGACTI&(ITABLA%) > 5 Then REGTOPE&(ITABLA%) = REGACTI&(ITABLA%) - 5
   Screen.MousePointer = 1
   '
End Sub
'
Sub ENTRATECLA(Keycode)
   '
   If Keycode = 27 Then
      If TRIM$(FORSELECHO.Text1.TEXT) <> "" Then
        FORSELECHO.Text1.TEXT = ""
        Exit Sub
       Else
        VD1$(ITABLA%) = ""
        VD2$(ITABLA%) = ""
        FORSELECHO.Hide
      End If
      Exit Sub
      '
    ElseIf Keycode = 8 Then
      POMOUSE = FORSELECHO.Text1.SelStart
      If FORSELECHO.Text1.TEXT <> "" Then
        FORSELECHO.Text1.TEXT = Left$(FORSELECHO.Text1.TEXT, Len(FORSELECHO.Text1.TEXT) - 1)
      End If
      If POMOUSE > 1 Then FORSELECHO.Text1.SelStart = POMOUSE
      Exit Sub
      '
    ElseIf Keycode = 37 Or Keycode = 39 Then
      FORSELECHO.Text1.SetFocus
      Exit Sub
      '
    ElseIf Keycode >= 42 Then
      If Keycode < 120 Then
        FORSELECHO.Text1.TEXT = FORSELECHO.Text1.TEXT + Chr$(Keycode)
        FORSELECHO.Text1.SelStart = Len(FORSELECHO.Text1.TEXT)
        FORSELECHO.Text1.SetFocus
        Exit Sub
      End If
   End If
   '
   If FORSELECHO.TASEL(ITABLA%).Visible = True Then
      If Keycode = 13 Then
        If FORSELECHO.TASEL(ITABLA%).ListIndex >= 0 Then
          VD1$(ITABLA%) = Left$(FORSELECHO.TASEL(ITABLA%).TEXT, AC1%(ITABLA%))
          'RETA& = Val(Mid$(FORSELECHO.TASEL(ITABLA%).TEXT, 2 + AC1%(ITABLA%) + AC2%(ITABLA%)))
   REGG& = Val(Mid$(FORSELECHO.TASEL(ITABLA%).TEXT, 3 + ACOL1%(Index) + ACOL2%(Index)))
   Call SETREG(ARCHU$, REGG&)
   IACTI&(Index) = REGG& - 1
   RETA& = REGG&
          FORSELECHO.Hide
          Exit Sub
        End If
       ElseIf Keycode = 38 Then
        'If FORSELECHO.TASEL(ITABLA%).LISTINDEX > 0 Then
        '  FORSELECHO.TASEL(ITABLA%).LISTINDEX = FORSELECHO.TASEL(ITABLA%).LISTINDEX - 1
        'End If
       ElseIf Keycode = 40 Then
        'If FORSELECHO.TASEL(ITABLA%).LISTINDEX < FORSELECHO.TASEL(ITABLA%).ListCount - 1 Then
        '  FORSELECHO.TASEL(ITABLA%).LISTINDEX = FORSELECHO.TASEL(ITABLA%).LISTINDEX + 1
        'End If
      End If
      FORSELECHO.TASEL(ITABLA%).SetFocus
      Exit Sub
   End If
   '
   If FORSELECHO.TABLA(ITABLA%).Visible = True Then
     FORSELECHO.TABLA(ITABLA%).SetFocus
     If Keycode = 13 Then
        If FORSELECHO.TABLA(ITABLA%).ListIndex >= 0 Then
          VD1$(ITABLA%) = Left$(FORSELECHO.TABLA(ITABLA%).TEXT, AC1%(ITABLA%))
   REGG& = REGACTI&(ITABLA%)
   Call SETREG(ARCHU$, REGG&)
   IACTI&(Index) = REGG& - 1
   RETA& = REGG&
          FORSELECHO.Hide
        End If
        Exit Sub
     ElseIf Keycode = 38 Then
       If FORSELECHO.TABLA(ITABLA%).ListIndex = 0 Then
         If REGACTI&(ITABLA%) > 1 Then
           REGTOPE(ITABLA%) = REGTOPE(ITABLA%) - 1
           Call CARGATABLA
         End If
       End If
     ElseIf Keycode = 40 Then
       If FORSELECHO.TABLA(ITABLA%).ListIndex = 10 Then
         If REGTOPE(ITABLA%) < FORSELECHO.SCROLLBA.Max Then
           REGTOPE(ITABLA%) = REGTOPE(ITABLA%) + 1
           Call CARGATABLA
         End If
       End If
     ElseIf Keycode = 33 Then
       REGACTI(ITABLA%) = REGACTI(ITABLA%) - 11
       REGTOPE(ITABLA%) = REGTOPE(ITABLA%) - 11
       If REGTOPE(ITABLA%) < 1 Then REGTOPE(ITABLA%) = 1
       If REGACTI(ITABLA%) < REGTOPE(ITABLA%) Then REGACTI(ITABLA%) = REGTOPE(ITABLA%)
       Call CARGATABLA
     ElseIf Keycode = 34 Then
       REGACTI(ITABLA%) = REGACTI(ITABLA%) + 11
       REGTOPE(ITABLA%) = REGTOPE(ITABLA%) + 11
       If REGACTI(ITABLA%) > FORSELECHO.SCROLLBA.Max Then REGACTI(ITABLA%) = FORSELECHO.SCROLLBA.Max
       If REGTOPE(ITABLA%) > REGACTI(ITABLA%) Then REGTOPE(ITABLA%) = REGACTI(ITABLA%)
       Call CARGATABLA
     ElseIf Keycode = 36 Then
       REGACTI(ITABLA%) = 1
       REGTOPE(ITABLA%) = 1
       Call CARGATABLA
     ElseIf Keycode = 35 Then
       REGACTI(ITABLA%) = ULTREG&(ARCHU$)  ' FORSELECHO.SCROLLBA.Max
       REGTOPE(ITABLA%) = REGACTI(ITABLA%) - 10
       If REGTOPE(ITABLA%) < 1 Then REGTOPE(ITABLA%) = 1
       Call CARGATABLA
    End If
    '
  End If
  '
End Sub

Sub CARTASEL()
   '
   FORSELECHO.TASEL(ITABLA%).Clear
   ABUX$ = TRIM$(FORSELECHO.Text1.TEXT)
   ABUS$ = UCase$(ABUX$)
   ABUT$ = LCase$(ABUX$)
   LOREBU% = AC1%(ITABLA%) + AC2%(ITABLA%) + 2
   If ARCHA$(ITABLA%) = "ECOMAST" Then LOREBU% = 64
   TINI = Timer
   If ABUS$ <> "" Then
      On Error GoTo 28
      PINI& = 1
24    PPXY& = InStr(PINI&, TEBUS$(ITABLA%), ABUS$)
      PPXY1& = InStr(PINI&, TEBUS$(ITABLA%), ABUS$)
      PPXY2& = InStr(PINI&, TEBUS$(ITABLA%), ABUT$)
      If PPXY1& > 0 Then If (PPXY1& < PPXY& Or PPXY& < 1) Then PPXY& = PPXY1&
      If PPXY2& > 0 Then If (PPXY2& < PPXY& Or PPXY& < 1) Then PPXY& = PPXY2&
      If PPXY& >= PINI& Then
         RETA& = Int(PPXY& / LOREBU%)
         TTTT$ = Mid$(TEBUS$(ITABLA%), LOREBU% * RETA& + 1, LOREBU%) + Space$(6) + Str$(RETA& + 1)
         FORSELECHO.TASEL(ITABLA%).AddItem TTTT$
         PINI& = LOREBU% * RETA& + 1 + LOREBU%
         If Timer - TINI < 0.2 Then GoTo 24
         FORSELECHO.TASEL(ITABLA%).AddItem " . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . ."
28    End If
      On Error GoTo 0
      FORSELECHO.TABLA(ITABLA%).Visible = False
      If FORSELECHO.TASEL(ITABLA%).ListCount > 0 Then
         FORSELECHO.TASEL(ITABLA%).ListIndex = 0
      End If
      FORSELECHO.TASEL(ITABLA%).Visible = True
      'FORSELECHO.SCROLLBA.Value = FORSELECHO.SCROLLBA.Min
      '
     Else
      '
      FORSELECHO.TASEL(ITABLA%).Visible = False
      FORSELECHO.TABLA(ITABLA%).Visible = True
      '
   End If
   '
End Sub
