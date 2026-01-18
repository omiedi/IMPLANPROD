Attribute VB_Name = "MENSAJNU"
Public MODACDIR$(32)
Public MODALTER%
Public ALTERX%
Public TIMENSA%    ' 0-COMALTER ; 1-COMUNI ; 2-MENSERR
'

Sub FRESHALL()
    '
    If UCase$(App.EXEName) = "FAC2002" Then Exit Sub
    If UCase$(App.EXEName) = "FAC2001" Or UCase$(App.EXEName) = "RFAC2001" Then
      For TTT& = 1 To Forms.Count
        If Forms(TTT& - 1).Visible = True Then
          On Error Resume Next
          For U& = 0 To Forms(TTT& - 1).Controls.Count - 1
            Forms(TTT& - 1).Controls(U&).Refresh
          Next U&
          On Error GoTo 0
        End If
      Next TTT&
    End If
    DoEvents
    '
End Sub
'
Sub FRESHPAR(TOPE, IZQI, ALTU, ANCH)
    '
    For TTT& = 1 To Forms.Count
      If Forms(TTT& - 1).Visible = True Then
        On Error Resume Next
        For U& = 0 To Forms(TTT& - 1).Controls.Count - 1
          TOP1 = Forms(TTT& - 1).Controls(U&).Left
          IZQ1 = Forms(TTT& - 1).Controls(U&).Left
          ALT1 = Forms(TTT& - 1).Controls(U&).Left
          ACH1 = Forms(TTT& - 1).Controls(U&).Left
          If TOP1 < TOPE + ALTU And (TOP1 + ALT1) > TOPE Then
            If IZQ1 < (IZQI + ANCH) And (IZQ1 + ACH1) > IZQI Then
              Forms(TTT& - 1).Controls(U&).Refresh
            End If
          End If
        Next U&
        On Error GoTo 0
      End If
    Next TTT&
    '
End Sub
'

Sub ACCDIR(MODU$)
    '
    LOMAX = 0
    CAOP% = 0
    MODX$ = UCase$(AJUSTI$(MODU$, 8))
    Dim OPX As String * 48
    N1% = FreeFile
    Open "ACCDIR.DRV" For Random Access Read Shared As #N1% Len = 48
    For I& = 1 To LOF(N1%) / 48
       Get #N1%, I&, OPX$
       If UCase$(Left$(OPX$, 8)) = MODX$ Then
          OPXY$ = TRIM$(Mid$(OPX$, 9, 32))
          MDXY$ = UCase$(TRIM$(Mid$(OPX$, 41)))
          If OPXY$ <> "" Then
             If MDXY$ <> "" Then
               FORMALTER.TEPRUEBA.Caption = OPXY$
               LONOP = FORMALTER.TEPRUEBA.Width
               If LONOP > LOMAX Then LOMAX = LONOP
               CAOP% = CAOP% + 1
               FORMALTER.BOTALTER(CAOP% - 1).Caption = OPXY$
               MODACDIR$(CAOP%) = MDXY$
             End If
          End If
       End If
    Next I&
    Close #N1%: N1% = 0
    '
    If CAOP% < 1 Then
        Call MENSERR(24, "No hay Accesos Directos\para esta Aplicación.")
        Exit Sub
    End If
    '
    FORMALTER.TEPRUEBA.Caption = "Terminar"
    LONOP = FORMALTER.TEPRUEBA.Width
    If LONOP > LOMAX Then LOMAX = LONOP
    CAOP% = CAOP% + 1
    FORMALTER.BOTALTER(CAOP% - 1).Caption = "Terminar"
    MODACDIR$(CAOP%) = ""
    '
    BOSUP = 210
    FORMALTER.Width = 1.5 * (250 + LOMAX)
    For I& = 0 To CAOP% - 1
        If I& = CAOP% - 1 Then
           BOSUP = BOSUP + 0.5 * FORMALTER.BOTALTER(I&).Height
           FORMALTER.command1.Top = BOSUP
        End If
        FORMALTER.BOTALTER(I&).Left = 0.25 * (150 + LOMAX)
        FORMALTER.BOTALTER(I&).Width = 250 + LOMAX
        FORMALTER.BOTALTER(I&).Top = BOSUP
        BOSUP = BOSUP + 1.6 * FORMALTER.BOTALTER(I&).Height
        FORMALTER.BOTALTER(I&).Visible = True
    Next I&
    '
    ALTU = FORMALTER.Height - FORMALTER.ScaleHeight
    ALTU = ALTU + BOSUP
    FORMALTER.Height = ALTU
    FORMALTER.Image1.Visible = False
    FORMALTER.Image2.Visible = False
    FORMALTER.command1.Visible = False
    FORMALTER.Frame1.Left = 0
    FORMALTER.Frame1.Top = 0
    FORMALTER.Frame1.Width = FORMALTER.ScaleWidth
    FORMALTER.Frame1.Height = FORMALTER.ScaleHeight
    FORMALTER.Top = Screen.Height / 2 - FORMALTER.Height / 2
    FORMALTER.Left = Screen.Width / 2 - FORMALTER.Width / 2
    '
    MODALTER% = 1
    FORMALTER.Refresh
    FORMALTER.command1.Visible = False
    FORMALTER.Show 1
    Call FRESHALL
    '
End Sub
'

Function ALTERNA%(MODU$)
    '
    LOMAX = 0
    CAOP% = 0
    '
    For I& = 0 To 31
        FORMALTER.BOTALTER(I&).Visible = False
        FORMALTER.LEYEN(I&).Visible = False
    Next I&
    '
    TEXT$ = LTrim$(MODU$)
    If Left$(TEXT$, 1) = "&" Then
        TEXT$ = Mid$(TEXT$, 2)
        POSII% = InStr(TEXT$, "&")
        If POSII% > 0 Then
            POSI$ = Left$(TEXT$, POSII% - 1)
            TEXT$ = Mid$(TEXT$, POSII% + 1)
        End If
    End If
    '
    RG% = Int(Val(POSI$))
    CLII% = 100 * (Val(POSI$) - RG%)
    '
    NOCIE% = 0: If InStr(POSI$, "/NC") > 0 Then NOCIE% = 1
    PCOLOR% = InStr(POSI$, "/COLOR=")
    'COLOVEN$ = COLSTAN$(5): If PCOLOR% > 0 Then COLOVEN$ = Mid$(POSI$, PCOLOR% + 7, 3)
    'If MONICOL% <> 1 Then COLOVEN$ = "B-N"
    POSEL% = InStr(POSI$, "/SX"): If POSEL% > 0 Then SELEC% = Val(Mid$(POSI$, POSEL% + 3))
    POSEL% = InStr(POSI$, "/AX"): If POSEL% > 0 Then ACTIV% = Val(Mid$(POSI$, POSEL% + 3))
    '
    If RG% < 2 Or RG% > 23 Then RG% = 13
    If CLII% < 1 Or CLII% > 80 Then CLII% = 40
    '
    If InStr(TEXT$, "\") > 0 Then GoTo 100         ' modo deslizante
    '
    TEXTX$ = TEXT$
    DEFO% = 1
    DEFOS$ = "SI"
    If Left$(TEXT$, 1) = "-" Then
        TEXTX$ = Mid$(TEXT$, 2)
        DEFO% = 0
        DEFOS$ = "NO"
    End If
    '
    ALTE% = DEFO%
    TEXTX$ = TEXTX$ + " ?  (S/N) "
    BPT% = 36: If DEFO% = 0 Then BPT% = BPT% + 256
90  ATT% = MsgBox(TEXTX$, BPT%, "Opción")
    If ATT% = 6 Then
        ALTERNA% = 1
        Call FRESHALL
        Exit Function
      ElseIf ATT% = 7 Then
        ALTERNA% = 0
        Call FRESHALL
        Exit Function
      Else
      GoTo 90
    End If
    '
100 TEXTX$ = TEXT$
    For I& = 1 To 12
        PX% = InStr(TEXTX$, "\")
        If PX% < 1 Then PX% = 1 + Len(TEXTX$)
        OPXY$ = Left$(TEXTX$, PX% - 1)
        Call SUPRACENTO(OPXY$)
        TEXTX$ = Mid$(TEXTX$, PX% + 1)
        If TRIM$(OPXY$) = "- o -" Then OPXY$ = "-"
        If OPXY$ <> "" Then
            FORMALTER.TEPRUEBA.Caption = OPXY$
            LONOP = FORMALTER.TEPRUEBA.Width
            If LONOP > LOMAX Then LOMAX = LONOP
            CAOP% = CAOP% + 1
            FORMALTER.BOTALTER(CAOP% - 1).Caption = OPXY$
        End If
        If Len(TEXTX$) < 1 Then GoTo 110
    Next I&
    '
110 If CAOP% < 1 Then
        Call MENSERR(24, "Error de Invocación de Función 'ALTERNA' - Consulte.")
        ALTERNA% = 0
        Call FRESHALL
        Exit Function
    End If
    '
    FORMALTER.Caption = "Opciones Alternativas"
    FORMALTER.TEPRUEBA.Caption = "Cancelar"
    LONOP = FORMALTER.TEPRUEBA.Width
    If LONOP > LOMAX Then LOMAX = LONOP
    CAOP% = CAOP% + 1
    FORMALTER.BOTALTER(CAOP% - 1).Caption = "Cancelar"
    MODACDIR$(CAOP%) = ""
    '
    BOSUP = 210
    FORMALTER.Width = 1.5 * (250 + LOMAX)
    For I& = 0 To CAOP% - 1
        '
        If I& = CAOP% - 1 Then
           BOSUP = BOSUP + 0.5 * FORMALTER.BOTALTER(I&).Height
           FORMALTER.command1.Top = BOSUP
        End If
        If TRIM$(FORMALTER.BOTALTER(I&).Caption) = "-" Then
           BOSUP = BOSUP + 0.5 * FORMALTER.BOTALTER(I&).Height
           GoTo 129
        End If
        '
        FORMALTER.BOTALTER(I&).Left = 0.25 * (150 + LOMAX)
        FORMALTER.BOTALTER(I&).Width = 250 + LOMAX
        FORMALTER.BOTALTER(I&).Top = BOSUP
        BOSUP = BOSUP + 1.6 * FORMALTER.BOTALTER(I&).Height
        FORMALTER.BOTALTER(I&).Visible = True
        '
129 Next I&
    '
    ALTU = FORMALTER.Height - FORMALTER.ScaleHeight
    ALTU = ALTU + BOSUP
    FORMALTER.Height = ALTU
    '
    FORMALTER.Top = Screen.Height * RG% / 25 - FORMALTER.Height / 2
    FORMALTER.Left = Screen.Width * CLII% / 80 - FORMALTER.Width / 2
    '
    FORMALTER.Image1.Visible = False
    FORMALTER.Image2.Visible = False
    FORMALTER.Frame1.Left = 0
    FORMALTER.Frame1.Top = 0
    FORMALTER.Frame1.Width = FORMALTER.ScaleWidth
    FORMALTER.Frame1.Height = FORMALTER.ScaleHeight
    '
    MODALTER% = 2
    FORMALTER.Refresh
    SMP = Screen.MousePointer
    Screen.MousePointer = 1
    ALTERX% = 0
    FORMALTER.command1.Visible = False
    FORMALTER.Show 1
    Screen.MousePointer = SMP
    Call FRESHALL
    ALTERNA% = ALTERX%
    '
End Function
'

Function COMALTER%(TEX$)
    '
    FORMALTNU.COMALT = TEX$
    FORMALTNU.Show 1
End Function
'

Sub TRACE(RG%, I&, FIN&)
    '
    Static FINA&, COEF, TANT, TANTUI, VISI%, VERTI%, TINI, IFORMU&, IBARRA&, IMARCO&
    '
    If FIN& <> FINA& Then
        FINA& = FIN&
        VERTI% = 0
        If FIN& > 0 Then
            '
            IFORMU& = -1: IBARRA& = -1: IMARCO& = (-1)
            For TTT& = 1 To Forms.Count
               If Forms(TTT& - 1).Visible = True Then
                  If Forms(TTT& - 1).Enabled = True Then
                    For U& = 0 To Forms(TTT& - 1).Controls.Count - 1
                        NM$ = Forms(TTT& - 1).Controls(U&).Name
                        If NM$ = "BARRATRAZA" Then
                            IFORMU& = TTT& - 1
                            IBARRA& = U&
                          ElseIf NM$ = "MARCOBARRA" Then
                            IFORMU& = TTT& - 1
                            IMARCO& = U&
                        End If
                    Next U&
                  End If
               End If
               If IFORMU& >= 0 Then GoTo 10
            Next TTT&
            Exit Sub
            '
10          If IFORMU& < 0 Or IMARCO& < 0 Or IBARRA& < 0 Then
                Exit Sub
            End If
            '
            COEF = Forms(IFORMU&).Controls(IMARCO&).Width / FIN&
            If Forms(IFORMU&).Controls(IMARCO&).Height > Forms(IFORMU&).Controls(IMARCO&).Width Then
               VERTI% = 1
               COEF = Forms(IFORMU&).Controls(IMARCO&).Height / FIN&
            End If
            TINI = Timer
        End If
        On Error Resume Next
        Forms(IFORMU&).Controls(IBARRA&).Left = 0
        Forms(IFORMU&).Controls(IBARRA&).Top = 0
        Forms(IFORMU&).Controls(IBARRA&).Width = Forms(IFORMU&).Controls(IMARCO&).Width
        Forms(IFORMU&).Controls(IBARRA&).Height = Forms(IFORMU&).Controls(IMARCO&).Height
        On Error GoTo 0
    End If
    '
    If Abs(Timer - TANTUI) < 0.25 Then
       Exit Sub
    End If
    TANTUI = Timer
    '
    If RG% = 1 Then
      If Int(Timer) = TANT Then
        If I& < FIN& Then
          Exit Sub
        End If
      End If
      TANT = Int(Timer)
    End If
    '
    On Error Resume Next
    If IFORMU& >= 0 Then
       If IMARCO& >= 0 Then
          If IBARRA& >= 0 Then
             If VERTI% = 0 Then
                  Forms(IFORMU&).Controls(IBARRA&).Left = COEF * I&
                Else
                  Forms(IFORMU&).Controls(IBARRA&).Top = COEF * (-I&)
             End If
             '
             If I& >= FIN& Then
                FINA& = 0
                Forms(IFORMU&).Controls(IBARRA&).Left = 0
                Forms(IFORMU&).Controls(IBARRA&).Top = 0
             End If
             '
             Forms(IFORMU&).Controls(IMARCO&).Refresh
             Forms(IFORMU&).Controls(IBARRA&).Refresh
          End If
       End If
    End If
    On Error GoTo 0
    '
End Sub
'
