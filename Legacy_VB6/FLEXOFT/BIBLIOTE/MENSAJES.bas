Attribute VB_Name = "MENSAJES"
Public MODALTER%
Public ALTERX%
Public TIMENSA%    ' 0-COMALTER ; 1-COMUNI ; 2-MENSERR
'
Public TEXALT$, CANALT%, VECALT$(64)
'
Sub FRESHFORM(CPTX$)
    On Error Resume Next
    AppActivate CPTX$
    SendKeys Chr$(31), True
    On Error GoTo 0
    OPENFORMSB = DoEvents
End Sub
'
Sub FRESHALL()
    '
    If UCase$(App.EXEName) = "FAC2002" Then Exit Sub
    If UCase$(App.EXEName) = "FAC2001" Then GoTo 10
    If UCase$(App.EXEName) = "CONSTO04" Then GoTo 10
    If UCase$(App.EXEName) = "RFAC2001" Then Exit Sub
    DoEvents
Exit Sub
    If App.EXEName = "AMASTO01" Then Exit Sub
    '
10  For TTT& = Forms.Count To 1 Step -1
      If Forms(TTT& - 1).Visible = True Then
        'Forms(TTT& - 1).Refresh
        'NOMEFORM$ = UCase$(Forms(TTT& - 1).Name)
        'If InStr(NOMEFORM$, "FORVENTA") > 0 Or InStr(NOMEFORM$, "ORDEFAB") > 0 Then
            On Error Resume Next
            For U& = 0 To Forms(TTT& - 1).Controls.Count - 1
                Forms(TTT& - 1).Controls(U&).Refresh
            Next U&
            On Error GoTo 0
        'End If
      End If
    Next TTT&
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
Sub ADDACCDIR(LABALTER$, CONNECT$, Optional ENTORNO$)
    '
    If ENTORNO$ <> "" Then
       EPAC$ = TRIM$(UCase$(EMPREAC$))
       If InStr(EPAC$, UCase$(ENTORNO$)) < 1 Then Exit Sub
    End If
    '
    If InStr(CONNECT$, ".") > 0 Then
         If EXISTE%(App.Path + "\" + CONNECT$) > 0 Then GoTo 70
       Else
         For kkj% = 1 To Forms.Count
           If UCase$(Forms(kkj% - 1).Name) = UCase$(CONNECT$) Then GoTo 70
         Next kkj%
    End If
    Exit Sub
    '
70  If TEXTALT$ = "" Then
        CANALT% = 0
      Else
        TEXALT$ = TEXALT$ + "\"
    End If
    TEXALT$ = TEXALT$ + LABALTER$
    CANALT% = CANALT% + 1
    VECALT$(CANALT%) = CONNECT$
    '
End Sub

Sub ACCDIR(MODU$)
    '
    LOMAX = 0
    CAOP% = 0
    MODX$ = UCase$(AJUSTI$(MODU$, 8))
    '
    On Error Resume Next
    FORMALTER.TEPRUEBA.Font = "Ms Sans Serif"
    FORMALTER.TEPRUEBA.FontSize = 8
    FORMALTER.TEPRUEBA.FontBold = True
    On Error GoTo 0
    '
    Dim OPX As String * 48
    n1% = FreeFile
    Open "ACCDIR.DRV" For Random Access Read Shared As #n1% Len = 48
    For i& = 1 To LOF(n1%) / 48
       Get #n1%, i&, OPX$
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
               FORMALTER.MODACDIR(CAOP% - 1).Caption = MDXY$
             End If
          End If
       End If
    Next i&
    Close #n1%: n1% = 0
    '
    If CAOP% < 1 Then
        Call MENSERR(24, "No hay Accesos Directos\para esta Aplicación (" + TRIM$(MODX$) + ").")
        Exit Sub
    End If
    '
    FORMALTER.TEPRUEBA.Caption = "Terminar " + TRIM$(MODX$)
    LONOP = FORMALTER.TEPRUEBA.Width
    If LONOP > LOMAX Then LOMAX = LONOP
    CAOP% = CAOP% + 1
    FORMALTER.BOTALTER(CAOP% - 1).Caption = "Terminar " + TRIM$(MODX$)
    FORMALTER.MODACDIR(CAOP% - 1).Caption = ""
    '
    BOSUP = 210
    FORMALTER.Width = 1.5 * (250 + LOMAX)
    For i& = 0 To CAOP% - 1
        If i& = CAOP% - 1 Then
           BOSUP = BOSUP + 0.5 * FORMALTER.BOTALTER(i&).Height
           FORMALTER.Command1.Top = BOSUP
        End If
        FORMALTER.BOTALTER(i&).Left = 0.25 * (150 + LOMAX)
        FORMALTER.BOTALTER(i&).Width = 250 + LOMAX
        FORMALTER.BOTALTER(i&).Top = BOSUP
        BOSUP = BOSUP + 1.6 * FORMALTER.BOTALTER(i&).Height
        FORMALTER.BOTALTER(i&).Visible = True
    Next i&
    '
    ALTU = FORMALTER.Height - FORMALTER.ScaleHeight
    ALTU = ALTU + BOSUP
    FORMALTER.Height = ALTU
    FORMALTER.Picture3.Visible = False
    FORMALTER.Picture4.Visible = False
    FORMALTER.Command1.Visible = False
    FORMALTER.Frame1.Left = 0
    FORMALTER.Frame1.Top = 0
    FORMALTER.Frame1.Width = FORMALTER.ScaleWidth
    FORMALTER.Frame1.Height = FORMALTER.ScaleHeight
    FORMALTER.Top = Screen.Height / 2 - FORMALTER.Height / 2
    FORMALTER.Left = Screen.Width / 2 - FORMALTER.Width / 2
    '
    MODALTER% = 1
    FORMALTER.Refresh
    FORMALTER.Command1.Visible = False
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
    For KKI% = 0 To 31
        On Error Resume Next
        FORMALTER.BOTALTER(KKI%).Visible = False
        FORMALTER.LEYEN(KKI%) = ""
        FORMALTER.LEYEN(KKI%).Visible = False
        On Error GoTo 0
    Next KKI%
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
    RG% = Int(XVALO(POSI$))
    CLII% = 100 * (XVALO(POSI$) - RG%)
    '
    NOCIE% = 0: If InStr(POSI$, "/NC") > 0 Then NOCIE% = 1
    PCOLOR% = InStr(POSI$, "/COLOR=")
    'COLOVEN$ = COLSTAN$(5): If PCOLOR% > 0 Then COLOVEN$ = Mid$(POSI$, PCOLOR% + 7, 3)
    'If MONICOL% <> 1 Then COLOVEN$ = "B-N"
    POSEL% = InStr(POSI$, "/SX"): If POSEL% > 0 Then SELEC% = XVALO(Mid$(POSI$, POSEL% + 3))
    POSEL% = InStr(POSI$, "/AX"): If POSEL% > 0 Then ACTIV% = XVALO(Mid$(POSI$, POSEL% + 3))
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
    For i& = 1 To 12
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
    Next i&
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
    FORMALTER.MODACDIR(CAOP% - 1).Caption = ""
    '
    BOSUP = 460
    FORMALTER.Width = 1.7 * (250 + LOMAX)
    For i& = 0 To CAOP% - 1
        '
        If i& = CAOP% - 1 Then
           BOSUP = BOSUP + 0.5 * FORMALTER.BOTALTER(i&).Height
           FORMALTER.Command1.Top = BOSUP
        End If
        If TRIM$(FORMALTER.BOTALTER(i&).Caption) = "-" Then
           BOSUP = BOSUP + 0.5 * FORMALTER.BOTALTER(i&).Height
           GoTo 129
        End If
        '
        FORMALTER.BOTALTER(i&).Left = 0.35 * (150 + LOMAX)
        FORMALTER.BOTALTER(i&).Width = 250 + LOMAX
        FORMALTER.BOTALTER(i&).Top = BOSUP
        BOSUP = BOSUP + 1.5 * FORMALTER.BOTALTER(i&).Height
        FORMALTER.BOTALTER(i&).Visible = True
        '
129 Next i&
    '
    ALTU = FORMALTER.Height - FORMALTER.ScaleHeight
    ALTU = ALTU + BOSUP
    FORMALTER.Height = ALTU
    '
    FORMALTER.Top = Screen.Height * RG% / 25 - FORMALTER.Height / 2
    FORMALTER.Left = Screen.Width * CLII% / 80 - FORMALTER.Width / 2
    '
    FORMALTER.Picture3.Visible = False
    FORMALTER.Picture4.Visible = False
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
    FORMALTER.Command1.Visible = False
    '
    Call APLICACIONSKINS(FORMALTER)
    '
    FORMALTER.Show 1
    Screen.MousePointer = SMP
    Call FRESHALL
    ALTERNA% = ALTERX%
    '
End Function
'

Function COMALTER%(TEX$)
    '
    Dim REMENSA$(32)
    '
    TEXT$ = TEX$
    POSI$ = ""
    '
    TEXT$ = LTrim$(TEXT$)
    If Left$(TEXT$, 1) = "&" Then
        TEXT$ = Mid$(TEXT$, 2)
        POSK% = InStr(TEXT$, "&")
        If POSK% > 0 Then
            POSI$ = Left$(TEXT$, POSK% - 1)
            TEXT$ = Mid$(TEXT$, POSK% + 1)
        End If
    End If
    '
    RGQ% = Int(XVALO(POSI$))
    If RGQ% < 2 Or RGQ% > 23 Then RGQ% = 13
    '
    CLQ% = 100 * (XVALO(POSI$) - RGQ%)
    If CLQ% < 1 Or CLQ% > 80 Then CLQ% = 40
    '
    NOCIE% = 0: If InStr(POSI$, "/NC") > 0 Then NOCIE% = 1
    '
    Call SUPRACENTO(TEXT$)
    POSEP% = InStr(TEXT$, "\\")
    If POSEP% < 1 Then
        MsgBox "Opciones Indefinidas en Rutina 'COMALTER'"
        TIMENSA% = 0
        Exit Function
    End If
    '
    OPCI$ = Mid$(TEXT$, POSEP% + 2)
    TEXT$ = Left$(TEXT$, POSEP% - 1)
    '
    If Len(OPCI$) < 1 Then
        MsgBox "No hay Opciones en Rutina 'COMALTER'"
        TIMENSA% = 0
        Exit Function
    End If
    '
    FORMALTER.Frame1.Top = 0
    FORMALTER.Frame1.Left = FORMALTER.Picture3.Left + FORMALTER.Picture3.Width + 100 ' 0
    FORMALTER.Frame1.Caption = ""
    For i% = 0 To 31
        On Error Resume Next
        FORMALTER.BOTALTER(i%).Visible = False
        FORMALTER.LEYEN(i%) = ""
        FORMALTER.LEYEN(i%).Visible = False
        On Error GoTo 0
    Next i%
    '
    TEXTX$ = TEXT$
    CADAT% = 0: AFORMU = 0
    BOSUP = 120
    For i% = 1 To 16
         PX% = InStr(TEXTX$, "\")
         If PX% < 1 Then PX% = 1 + Len(TEXTX$)
         Alter$ = Left$(TEXTX$, PX% - 1)
         TEXTX$ = Mid$(TEXTX$, PX% + 1)
         If Len(Alter$) > 0 Then
             CADAT% = CADAT% + 1
             FORMALTER.LEYEN(CADAT% - 1) = "   " + Alter$ + "   "
             REMENSA$(CADAT%) = Alter$
             If FORMALTER.LEYEN(CADAT% - 1).Width > AFORMU Then
                 AFORMU = FORMALTER.LEYEN(CADAT% - 1).Width
             End If
             FORMALTER.LEYEN(CADAT% - 1).Top = BOSUP
             FORMALTER.LEYEN(CADAT% - 1).Left = 0
             FORMALTER.LEYEN(CADAT% - 1).Visible = True
             BOSUP = BOSUP + FORMALTER.LEYEN(CADAT% - 1).Height
         End If
         If Len(TEXTX$) < 1 Then GoTo 105
    Next i%
    '
105 TEXTX$ = OPCI$
    REMENSA$(CADAT% + 1) = OPCI$
    CAOP% = 0
    ABOTO = 0
    '
    On Error Resume Next
    FORMALTER.TEPRUEBA.Font = "Ms Sans Serif"
    FORMALTER.TEPRUEBA.FontSize = 8
    FORMALTER.TEPRUEBA.FontBold = True
    On Error GoTo 0
    '
    For i% = 1 To 10
         PX% = InStr(TEXTX$, "\")
         If PX% < 1 Then PX% = 1 + Len(TEXTX$)
         OPX$ = Left$(TEXTX$, PX% - 1)
         TEXTX$ = Mid$(TEXTX$, PX% + 1)
         If Len(OPX$) > 0 Then
             CAOP% = CAOP% + 1
             On Error Resume Next
             FORMALTER.BOTALTER(CAOP% - 1).Caption = OPX$
             FORMALTER.TEPRUEBA.Caption = "  " + OPX$ + "  "
             FORMALTER.BOTALTER(CAOP% - 1).Width = FORMALTER.TEPRUEBA.Width
             ABOTO = ABOTO + FORMALTER.BOTALTER(CAOP% - 1).Width
             On Error GoTo 0
         End If
         If Len(TEXTX$) < 1 Then GoTo 107
106 Next i%
    '
107 If ABOTO + 150 * (CAOP% + 1) > AFORMU Then
        AFORMU = ABOTO + 150 * (CAOP% + 1)
    End If
    INTERBOT = (AFORMU - ABOTO) / (CAOP% + 1)
    '
    BOSUP = BOSUP + 240
    BIZQ = INTERBOT
    For i% = 1 To CAOP%
        On Error Resume Next
        FORMALTER.BOTALTER(i% - 1).Left = BIZQ
        FORMALTER.BOTALTER(i% - 1).Top = BOSUP
        FORMALTER.BOTALTER(i% - 1).Visible = True
        BIZQ = BIZQ + FORMALTER.BOTALTER(i% - 1).Width + INTERBOT
        On Error GoTo 0
    Next i%
    '
    FORMALTER.Command1.Top = BOSUP
    '
    For i% = 1 To CADAT%
       While FORMALTER.LEYEN(i% - 1).Width < AFORMU
          FORMALTER.LEYEN(i% - 1).Caption = " " + FORMALTER.LEYEN(i% - 1).Caption + " "
       Wend
       FORMALTER.LEYEN(CADAT% - 1).Visible = True
    Next i%
    '
    ALTUTIT = FORMALTER.Height - FORMALTER.ScaleHeight
    ANBORDE = (FORMALTER.Width - FORMALTER.ScaleWidth) / 2
    FORMALTER.Width = AFORMU + 2 * ANBORDE
    On Error Resume Next
    FORMALTER.Height = BOSUP + FORMALTER.BOTALTER(0).Height + 120 + ALTUTIT
    On Error GoTo 0
    FORMALTER.Frame1.Width = FORMALTER.Width - 2 * ANBORDE
    FORMALTER.Frame1.Height = FORMALTER.ScaleHeight
    '
    FORMALTER.Caption = "Opciones Alternativas"
    FORMALTER.Picture3.Visible = True
    FORMALTER.Picture4.Visible = False
    If CAOP% < 2 Or TIMENSA% = 1 Then
        FORMALTER.Caption = "Información Crítica"
        FORMALTER.Width = 900 + FORMALTER.Width + FORMALTER.Command1.Width + 100
        FORMALTER.Frame1.Left = 900
    End If
    If TIMENSA% = 2 Then
        FORMALTER.Caption = "Operación Incorrecta"
        FORMALTER.Picture3.Visible = False
        FORMALTER.Picture4.Visible = True
    End If
    '
    FORMALTER.Top = Screen.Height * RGQ% / 25 - FORMALTER.Height / 2
    FORMALTER.Left = Screen.Width * CLQ% / 80 - FORMALTER.Width / 2
    MODALTER% = 2
    ALTERX% = 0
    '
    If FORMALTER.Width < FORMALTER.Frame1.Left + FORMALTER.Frame1.Width + 350 Then
       FORMALTER.Width = FORMALTER.Frame1.Left + FORMALTER.Frame1.Width + 350
    End If
    '
    Call APLICACIONSKINS(FORMALTER)
    '
    FORMALTER.Refresh
    SMP = Screen.MousePointer
    Screen.MousePointer = 1
    TII% = Int(HORANUM(Time$))
    On Error Resume Next
    FORMALTER.Show 1
    On Error GoTo 0
    Screen.MousePointer = SMP
    Call FRESHALL
    COMALTER% = ALTERX%
    '
    If VERCD% = 1 Then
      If TRIM$(LUDAT$) = "" Then
        Exit Function
      End If
    End If
    '
    If UCase$(App.EXEName) = "FLEXOFT" Then
       If EXISTE%(LUDAT$ + "MENSAJES.RFS") < 1 Then
          Exit Function
       End If
    End If
    '
    HOI% = HOY(HO$)
    TIF% = Int(HORANUM(Time$))
    TIGRA = Timer
    MNX$ = REGBLAN$("MENSAJES")
    Call REPLA(MNX$, MKS$(TIGRA), 1, 4)
    Call REPLA(MNX$, MKI$(HOI%), 5, 2)
    Call REPLA(MNX$, MKI$(TII%), 7, 2)
    Call REPLA(MNX$, MKI$(TIF%), 9, 2)
    Call REPLA(MNX$, MKI$(USNBR% Mod 100), 11, 2)
    Call REPLA(MNX$, NOMTER$, 13, 30)
    Call REPLA(MNX$, UCase$(App.EXEName), 43, 16)
    Call REPLA(MNX$, PROCON$, 59, 16)
    Call REPLA(MNX$, MKI$(ALTERX%), 157, 2)
    '
    For i% = 1 To CADAT%
       If TRIM$(REMENSA$(i%)) <> "" Then
          Call REPLA(MNX$, "." + AJUCEN$(REMENSA$(i%), 81), 75, 82)
          Call REGAPP("MENSAJES", MNX$)
          If FORMALTER.Visible = False Then
             FORMALTER.LEYEN(i% - 1).Caption = ""
          End If
       End If
    Next i%
    '
    If CAOP% > 1 Then
       Call REPLA(MNX$, "." + AJUCEN$(REMENSA$(CADAT% + 1), 81), 75, 82)
       Call REGAPP("MENSAJES", MNX$)
    End If
    '
    Call CIERRARCH("MENSAJES")
    TIMENSA% = 0
    '
End Function
'

Sub TRACE(RG%, i&, FIN&)
    '
    Static FINA&, COEF, TANT, TANTUI, VISI%, VERTI%, TINI, IFORMU&, IBARRA&, IMARCO&
    '
    If i& > FIN& Or FIN& < 1 Then
      Exit Sub
    End If
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
        If i& < FIN& Then
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
                  Forms(IFORMU&).Controls(IBARRA&).Left = COEF * i&
                Else
                  Forms(IFORMU&).Controls(IBARRA&).Top = COEF * (-i&)
             End If
             Forms(IFORMU&).Controls(IBARRA&).ToolTipText = i& & " / " & FIN&
             '
             If i& >= FIN& Then
                FINA& = 0
                Forms(IFORMU&).Controls(IBARRA&).Left = 0
                Forms(IFORMU&).Controls(IBARRA&).Top = 0
                Forms(IFORMU&).Controls(IBARRA&).ToolTipText = ""
             End If
             '
             Forms(IFORMU&).Controls(IMARCO&).Refresh
             Forms(IFORMU&).Controls(IBARRA&).Refresh
             DoEvents
          End If
       End If
    End If
    On Error GoTo 0
    '
End Sub
'

Sub LogParaSeguimientoDeFlujo(RutaNombreArchivo As String, TEXTO As String, Optional CierraArchivo As Boolean = False)

    Static NARCHI1%
    
    If NARCHI1% = 0 Then NARCHI1% = FILEOPEN%(RutaNombreArchivo, 2, 128)
    Print #NARCHI1%, TEXTO
    
    If CierraArchivo = True Then
        Close #NARCHI1%
        NARCHI1% = 0
    End If
    
End Sub
