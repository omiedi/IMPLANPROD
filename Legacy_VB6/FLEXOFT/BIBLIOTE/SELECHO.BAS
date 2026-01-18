Attribute VB_Name = "SELECHO"
Sub SELECHO(ARCHI$, VDI%, VDD#, VDS$)
    '
    Static ARCHA$(7)
    Static DENARCH$(7)
    Static CARCHI%
    '
    ARCHU$ = TRIM$(UCase$(ARCHI$))
    If ARCHU$ = "" Then
        Call MENSERR(24, "Error de Invocación de Rutina 'SELECHO' - Consulte Soporte de Sistemas.")
        VDI% = 0: VDD# = 0: VDS$ = ""
        Exit Sub
    End If
    '
    If ARCHU$ = "MASTER" Then
        ARCHU$ = "ECOMAST"
    End If
    '
    For UX% = 0 To CARCHI% - 1
        If ARCHA$(UX%) = ARCHU$ Then
            ITBLA% = UX%
            GoTo 120
        End If
    Next UX%
    '
    If CARCHI% < 8 Then
        CARCHI% = CARCHI% + 1
        ARCHA$(CARCHI% - 1) = ARCHU$
        TV1% = TIPFIELD%(ARCHU$, 1): TV2% = TIPFIELD%(ARCHU$, 2)
        LF1% = LENFIELD%(ARCHU$, 1): LF2% = LENFIELD%(ARCHU$, 2)
        PR1% = POSFIELD%(ARCHU$, 1): PR2% = POSFIELD%(ARCHU$, 2)
        ITABLA% = CARCHI% - 1
        '
        AC1% = LF1%
        If TV1% = 1 Or TV1% = 2 Then
              If LF1% < 2 Then AC1% = 4 Else AC1% = 6
            ElseIf TV1% = 3 Or TV1% = 4 Then
              If LF1% < 8 Then AC1% = 7 Else AC1% = 12
              If LD1% > 0 Then AC1% = AC1% + 2
            ElseIf TV1% = 5 Then
              AC1% = 8
            ElseIf TV1% = 7 Then
              AC1% = 16
        End If
        '
        AC2% = LF2%
        If TV2% = 1 Or TV2% = 2 Then
              If LF2% < 2 Then AC2% = 4 Else AC2% = 6
            ElseIf TV2% = 3 Or TV2% = 4 Then
              If LF2% < 8 Then AC2% = 7 Else AC2% = 12
              If LD2% > 0 Then AC2% = AC2% + 2
            ElseIf TV2% = 5 Then
              AC2% = 8
            ElseIf TV2% = 7 Then
              AC2% = 16
        End If
        '
        If AC1% + AC2% > 70 Then AC2% = 70 - AC1%
        AVEN% = AC1% + 2 + AC2%
        If AVEN% > 72 Then AVEN% = 72
        '
        FORMUSEL.TABLA(ITABLA%).Clear
        '
        For U& = 1 To ULTREG&(ARCHU$)
           XXX$ = REGLEIDO$(ARCHU$, U&)
           XXZ$ = Left$(XXX$, LF1%)
           TX1$ = CSTRING$(XXZ$, TV1%, AC1%, LD1%, JT1%)
           XXZ$ = Mid$(XXX$, PR2%, LF2%)
           TX2$ = CSTRING$(XXZ$, TV1%, AC1%, LD1%, JT1%)
           FORMUSEL.TABLA(ITABLA%).AddItem TX1$ + "  " + TX2$
        Next U&
        '
      Else
        Call MENSERR("Demasiadas Tablas Abiertas - Consulte al Soporte de Sistemas.")
        Exit Sub
    End If
    '
120 If ULTREG&(ARCHU$) < 1 Then
        Call MENSERR("Tabla de Validación Vacía.")
        Exit Sub
    End If
    '
    
End Sub

