Attribute VB_Name = "SELECHO1"
Public TASEL As String
Public ITABLA%
Public ARCHA$(15)
Public ASELEC$(15)
Public DENARCH$(15)
Public ALIS(15)
Public IACTI&(15)
'
Public VD1$(15), VD2$(15)
Public ACOL1%(15), ACOL2%(15)
Public TIVA1%(15), TIVA2%(15)

Dim IHOO As Single
'

Sub SELECHO(ARCHITITU$)
    '
    PPXX% = InStr(ARCHITITU$, "/")
    If PPXX% < 2 Then PPXX% = 1 + Len(ARCHITITU$)
    ARCHI$ = Left$(ARCHITITU$, PPXX% - 1)
    TTY$ = Mid$(ARCHITITU$, PPXX% + 1)
       '
    TASELEC$ = ""
    '
    For II = 0 To 15
        FORMUSEL.TABLA(II).Visible = False
        FORMUSEL.TASEL(II).Visible = False
    Next II
    '
    ITABLA% = IARCHI%(ARCHI$)
    ARCHU$ = ARCHA$(ITABLA%)
    '
120 If ARCHU$ = "FECHA" Or ARCHU$ = "SELFECHA" Then
        TTY$ = "Fecha"
        GoTo 130
    End If
    '
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
    FORMUSEL.BOTINSE.Top = FORMUSEL.Text1.Top - 10
    FORMUSEL.BOTERA.Top = FORMUSEL.Text1.Top - 10
    FORMUSEL.BOTCANCE.Top = FORMUSEL.Text1.Top - 10
    FORMUSEL.BOTCANCE.Left = ALIS(ITABLA%) - FORMUSEL.BOTCANCE.Width
    If ALIS(ITABLA%) < 3000 Then
        FORMUSEL.BOTERA.Width = 0
        FORMUSEL.BOTINSE.Width = 0
        FORMUSEL.BOTERA.Visible = False
        FORMUSEL.BOTINSE.Visible = False
    End If
    FORMUSEL.BOTERA.Left = FORMUSEL.BOTCANCE.Left - FORMUSEL.BOTERA.Width
    FORMUSEL.BOTINSE.Left = FORMUSEL.BOTERA.Left - FORMUSEL.BOTINSE.Width
    FORMUSEL.BOTINSE.Enabled = False
    FORMUSEL.BOTERA.Enabled = False
    '
    FORMUSEL.Text1.Width = FORMUSEL.BOTINSE.Left - FORMUSEL.Text1.Left - 100
    If FORMUSEL.TABLA(ITABLA%).ListCount > 0 Then
       If InStr(ARCHU$, "FECHA") < 1 Then
           FORMUSEL.TABLA(ITABLA%).ListIndex = 0
           If IACTI&(ITABLA%) >= 0 Then
               If IACTI&(ITABLA%) < FORMUSEL.TABLA(ITABLA%).ListCount Then
                   FORMUSEL.TABLA(ITABLA%).ListIndex = IACTI&(ITABLA%)
               End If
           End If
         Else
           FORMUSEL.TABLA(ITABLA%).ListIndex = IHOO
       End If
       TOPELI& = FORMUSEL.TABLA(ITABLA%).ListIndex - 5
       If TOPELI& < 0 Then TOPELI& = 0
       FORMUSEL.TABLA(ITABLA%).TopIndex = TOPELI&
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

Function IARCHI%(ARCHI$)
    '
    Static CARCHI%
    '
    ARCHU$ = TRIM$(UCase$(ARCHI$))
    If ARCHU$ = "" Then
        Call MENSERR(24, "Error de Invocación de Rutina 'SELECHO' - Consulte Soporte de Sistemas.")
        Exit Function
    End If
    '
    If ARCHU$ = "MASTER" Then
        ARCHU$ = "ECOMAST"
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
    For UX% = 0 To CARCHI% - 1
        If ARCHA$(UX%) = ARCHU$ Then
            IARCHI% = UX%
            ITABLA% = UX%
            If FORMUSEL.TABLA(ITABLA%).ListCount > 0 Then
                 Exit Function
               Else
                 Screen.MousePointer = 11
                 GoTo 100
            End If
        End If
    Next UX%
    '
    If CARCHI% >= 16 Then
        Call MENSERR(24, "Demasiadas Tablas Abiertas.\Consulte al Soporte de Sistemas.")
        Call FINAL("")
    End If
    '
    Screen.MousePointer = 11
    CARCHI% = CARCHI% + 1
    ARCHA$(CARCHI% - 1) = ARCHU$
    IARCHI% = CARCHI% - 1
    ITABLA% = IARCHI%
    '
100 If ARCHU$ = "FECHA" Or ARCHU$ = "SELFECHA" Then
          TV1% = 5: TV2% = 6
          LF1% = 2: LF2% = 14
          PR1% = 1: PR2% = 3
        Else
          TV1% = TIPFIELD%(ARCHU$, 1): TV2% = TIPFIELD%(ARCHU$, 2)
          LF1% = LENFIELD%(ARCHU$, 1): LF2% = LENFIELD%(ARCHU$, 2)
          PR1% = POSFIELD%(ARCHU$, 1): PR2% = POSFIELD%(ARCHU$, 2)
    End If
    '
    AC1% = LF1%
    JT1% = 1
    LD1% = 0
    If TV1% = 1 Or TV1% = 2 Then
              JT1% = 2
              If LF1% < 2 Then AC1% = 3 Else AC1% = 5
          ElseIf TV1% = 3 Or TV1% = 4 Then
              JT1% = 2: LD1% = 1
              If LF1% < 8 Then AC1% = 7 Else AC1% = 12
              If LD1% > 0 Then AC1% = AC1% + 2
          ElseIf TV1% = 5 Then
              AC1% = 8
          ElseIf TV1% = 7 Then
              AC1% = 16
    End If
    '
    ACOL1%(ITABLA%) = AC1%
    TIVA1%(ITABLA%) = TV1%
    '
    AC2% = LF2%
    JT2% = 1
    LD2% = 0
    If TV2% = 1 Or TV2% = 2 Then
              JT2% = 2
              If LF2% < 2 Then AC2% = 3 Else AC2% = 5
          ElseIf TV2% = 3 Or TV2% = 4 Then
              JT2% = 2
              LD2% = 2
              If LF2% < 8 Then AC2% = 7 Else AC2% = 12
              If LD2% > 0 Then AC2% = AC2% + 2
          ElseIf TV2% = 5 Then
              AC2% = 8
          ElseIf TV2% = 7 Then
              AC2% = 16
    End If
    '
    ACOL2%(ITABLA%) = AC2%
    '
    If AC1% + AC2% > 70 Then AC2% = 70 - AC1%
    AVEN% = AC1% + 2 + AC2% + 4
    If AVEN% > 72 Then AVEN% = 72
    FORMUSEL.TEPRUEBA.Caption = String$(AVEN%, "*")
    ALIS(ITABLA%) = FORMUSEL.TEPRUEBA.Width
    '
    FORMUSEL.TABLA(ITABLA%).Clear
    '
    If ARCHU$ = "FECHA" Or ARCHU$ = "SELFECHA" Then
        HOOI% = HOY(HO$)
        ANX = Val(Mid$(HO$, 7, 2)) - 1: If ANX < 0 Then ANX = ANX + 100
        ANXS$ = TRIM$(Str$(ANX)): While Len(ANXS$) < 2: ANXS$ = "0" + ANXS$: Wend
        Desde$ = Left$(HO$, 6) + ANXS$
        ANX = Val(Mid$(HO$, 7, 2)) + 1: If ANX > 99 Then ANX = ANX - 100
        ANXS$ = TRIM$(Str$(ANX)): While Len(ANXS$) < 2: ANXS$ = "0" + ANXS$: Wend
        HASTA$ = Left$(HO$, 6) + ANXS$
        '
        HAS = FECHANUM(HASTA$)
        FEX = FECHANUM(Desde$)
        While FEX <= HAS
           TX1$ = FECHATEX$(FEX)
           TX2$ = DSEM$(FEX)
           TX12$ = TX1$ + "  " + TX2$
           If TRIM$(TX12$) <> "" Then
              FORMUSEL.TABLA(ITABLA%).AddItem TX12$
              If FEX = HOOI% Then
                 FORMUSEL.TABLA(ITABLA%).ListIndex = FORMUSEL.TABLA(ITABLA%).ListCount - 1
                 IHOO = FORMUSEL.TABLA(ITABLA%).ListIndex
              End If
           End If
           FEY = DIPOST(FEX)
           FEX = FEY
        Wend
        Screen.MousePointer = 1
        Exit Function
    End If
    '
    If TRIM$(UCase$(ARCHU$)) = "ECOMAST" Then
       '
       LOTO& = ULTREG&("ECOMAST")
       Dim ALXX As String * 1024
       NNXX% = FILEOPEN%(LUDAT$ + "ECOMAST.DAT", 0, 1024)
       For U& = 1 To 1 + LOF(NNXX%) / 1024
          Get #NNXX%, U&, ALXX$
          For U1% = 1 To 16
             If 16 * (U& - 1) + U1% <= LOTO& Then
                ALYY$ = Mid$(ALXX$, 64 * U1% - 63, 64)
                TX12$ = Left$(ALYY$, 16) + "  " + Mid$(ALYY$, 17, 46)
                If TRIM$(TX12$) <> "" Then
                   FORMUSEL.TABLA(ITABLA%).AddItem TX12$
                End If
             End If
          Next U1%
       Next U&
       Close #NNXX%
       '
      Else
       '
       For U& = 1 To ULTREG&(ARCHU$)
          XXX$ = REGLEIDO$(ARCHU$, U&)
          XXZ$ = Left$(XXX$, LF1%)
          TX1$ = CSTRING$(XXZ$, TV1%, AC1%, LD1%, JT1%)
          XXZ$ = Mid$(XXX$, PR2%, LF2%)
          TX2$ = CSTRING$(XXZ$, TV2%, AC2%, LD2%, JT2%)
          TX12$ = TX1$ + "  " + TX2$
          If TRIM$(TX12$) <> "" Then
              FORMUSEL.TABLA(ITABLA%).AddItem TX12$
          End If
       Next U&
       '
    End If
    Screen.MousePointer = 1
    '
End Function

Function VALACT1$(ARCHI$)
    '
    ITABLA% = IARCHI%(ARCHI$)
    VALACT1$ = VD1$(ITABLA%)
    '
End Function
Function VALACT2$(ARCHI$)
    '
    ITABLA% = IARCHI%(ARCHI$)
    VALACT2$ = VD2$(ITABLA%)
    '
End Function

Sub FRESHECHO(ARCHI$)
    '
    IAA% = IARCHI%(ARCHI$)
    FORMUSEL.TABLA(IAA%).Clear
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



