Attribute VB_Name = "FLEX_FUNCTIONS"

Sub SELMENU(APLI$, RGL%)
    '
    Dim SS0 As String * 64
    Static N2%, CANIV%, LOPRINIV%, LOULNIV%, NIVBAS%, CONTA%, INCRE%, MENUAN$
    Dim COLVEN$(5), OPX$(32), REGX%(32), BINF%(8)
    '
    If APLI$ <> MENUAN$ Then         ' cambio de menu de aplicacion
        MENUAN$ = APLI$
        If N2% > 0 Then
            Close #N2%: N2% = 0
        End If
        NIVBAS% = 0
        N2% = 0
    End If
    '
    If N2% > 0 Then
        LOMENU = 0
        On Error Resume Next
        LOMENU = LOF(N2%)
        On Error GoTo 0
        If LOMENU < 1 Then N2% = 0
    End If
    '
10  If N2% < 1 Then
        If TRIM$(APLI$) = "" Then
            Call MENSERR(24, "Aplicacion no Definida")
            Close
            End
        End If
        '
        N2% = FreeFile
        Open APLI$ + ".MNU" For Random Access Read As #N2% Len = 64
        '
        CANIV% = 0: LOPRINIV% = 0: LOULNIV% = 0: CONTA% = 0
        For i% = 1 To LOF(N2%) / 64
           Get #N2%, i%, SS0$
           SS1$ = Left$(SS0$, 32)
           SS2$ = Mid$(SS0$, 33, 8)
           SS3$ = Mid$(SS0$, 41, 2)
           SS4$ = Mid$(SS0$, 43, 22)
           If CVI(SS3$) = 1 Then If Len(TRIM$(SS1$)) > LOPRINIV% Then LOPRINIV% = Len(TRIM$(SS1$))
           If CVI(SS3$) > CANIV% Then CANIV% = CVI(SS3$): LOULNIV% = Len(TRIM$(SS1$))
           If CVI(SS3$) = CANIV% Then If Len(TRIM$(SS1$)) > LOULNIV% Then LOULNIV% = Len(TRIM$(SS1$))
           CONTA% = i%
        Next i%
        '
        If CANIV% > 1 Then INCRE% = (144 - LOPRINIV% - LOULNIV%) / (2 * (CANIV% - 1))
        If INCRE% < 4 Then INCRE% = 4
        '
    End If
    '

    If RGL% = (-1) Then
       If NIVBAS% > 0 Then
           NIVBAS% = NIVBAS% - 1
           GoTo 310
         Else
           Exit Sub
        End If
    End If
    '
    If RGL% < 1 Or RGL% > CONTA% Then RGL% = 1
    '
    Get #N2%, RGL%, SS0$: SS1$ = Left$(SS0$, 32)
                          SS2$ = Mid$(SS0$, 33, 8)
                          SS3$ = Mid$(SS0$, 41, 2)
                          SS4$ = Mid$(SS0$, 43, 22)
    NIVBAS% = CVI(SS3$)
    FLEMENUS.Caption = TRIM$(SS1$)
    EPAC$ = TRIM$(EMPREAC$)
    If EPAC$ <> "" Then
        FLEMENUS.Caption = FLEMENUS.Caption + "  -  " + EPAC$
    End If
    '
    FLEMENUS.OPMENU(NIVBAS%).Clear
    '
    For i% = RGL% + 1 To CONTA%
        '
        Get #N2%, i%, SS0$: SS1$ = Left$(SS0$, 32)
                            SS2$ = Mid$(SS0$, 33, 8)
                            SS3$ = Mid$(SS0$, 41, 2)
                            SS4$ = Mid$(SS0$, 43, 22)
        If CVI(SS3$) < NIVBAS% + 1 Then
               GoTo 300
            ElseIf CVI(SS3$) = NIVBAS% + 1 Then
               If TRIM$(SS1$) <> "" Then
                   OPX1$ = LCase$(AJUSTI$(SS1$, 64) + SS2$ + Str$(i%))
                   Mid$(OPX1$, 1, 1) = UCase$(Left$(OPX1$, 1))
                   FLEMENUS.OPMENU(NIVBAS%).AddItem OPX1$
               End If
        End If
    Next i%
    '
300 J% = FLEMENUS.OPMENU(NIVBAS%).ListCount
    If J% < 1 Then
        Call MENSERR(24, "No hay Opciones para este Nivel")
        X% = 0
        Call FUERASYS
    End If
    '
    AREN% = 26 * FLEMENUS.OPMENU(NIVBAS%).FontSize
    If NIVBAS% < 1 Then
          TOPE% = 300
        Else
          TOPEANT% = FLEMENUS.OPMENU(NIVBAS% - 1).Top
          BASEANT% = TOPEANT% + AREN% * FLEMENUS.OPMENU(NIVBAS% - 1).ListCount
          TOPE% = TOPEANT% + 1.6 * AREN%
    End If
    '
301 Base% = TOPE% + AREN% * FLEMENUS.OPMENU(NIVBAS%).ListCount
    If Base% < BASEANT% + AREN% Then
        TOPE% = TOPE% + AREN%
        GoTo 301
    End If
    '
    FLEMENUS.OPMENU(NIVBAS%).Top = TOPE%
    '
310 For KK% = 0 To 4
       FLEMENUS.OPMENU(KK%).Enabled = False
       If KK% > NIVBAS% Then
           FLEMENUS.OPMENU(KK%).Visible = False
       End If
    Next KK%
    '
    'For KK% = 0 To 4
        'FLEMENUS.OPMENU(KK%).BackColor = &H80000016
        'If KK% = NIVBAS% Then
        '    FLEMENUS.OPMENU(KK%).BackColor = &H80000005
        'End If
    'Next KK%
    '
    jj% = FLEMENUS.OPMENU(NIVBAS%).ListCount
    FLEMENUS.OPMENU(NIVBAS%).Height = 50 + jj% * 26 * FLEMENUS.OPMENU(NIVBAS%).FontSize
    FLEMENUS.OPMENU(NIVBAS%).Visible = True
    FLEMENUS.OPMENU(NIVBAS%).Enabled = True
    FLEMENUS.OPMENU(NIVBAS%).ListIndex = 0
    Exit Sub
    '
End Sub

Sub MENUSEL()
    '
    FLESCRIT.MousePointer = 11
    '
    CAPLI% = 0
    Dim APLI$(64)
    '
    If (USNBR% Mod 100) <> 1 Then
       USR% = USNBR% Mod 100
        For U& = 1 To ULTREG&("MENUSER")
            X$ = REGLEIDO$("MENUSER", U&)
            If Asc(Left$(X$, 1)) = NROI% Then
                If CAPLI% < 32 Then
                    CAPLI% = CAPLI% + 1
                    APLI$(CAPLI%) = TRIM$(Mid$(X$, 2, 8))
                End If
             End If
        Next U&
        GoTo 130
    End If
    '
    For U = 1 To FLESCRIT.File1.ListCount
        FLESCRIT.File1.ListIndex = U - 1
        X$ = UCase$(FLESCRIT.File1.FileName)
        PX% = InStr(X$, ".MNU")
        If PX% > 1 Then
            APLX$ = TRIM$(UCase$(Left$(X$, PX% - 1)))
            If APLX$ <> "" Then
                If APLX$ <> "INDINOVE" Then
                    If CAPLI% < 64 Then
                        CAPLI% = CAPLI% + 1
                        APLI$(CAPLI%) = APLX$
                    End If
                End If
            End If
        End If
    Next U

'    DRX$ = CurDir
'    NX% = FreeFile
'    Open "C:\FLEXOFT\DOSCOMM.BAT" For Output As #NX%
'    Print #NX%, "DIR " + DRX$ + "\*.MNU >C:\FLEXOFT\DIRECTO"
'    Print #NX%, "EXIT"
'    Close #NX%
'    Shell "C:\FLEXOFT\DOSCOMM.PIF", 1
    '
130 Dim SS0 As String * 64
    J% = 0
    '
    For i% = 1 To CAPLI%
         N2% = FreeFile
         Open APLI$(i%) + ".MNU" For Random Access Read As #N2% Len = 64
         If LOF(N2%) >= 64 Then
             Get #N2%, 1, SS0$
             DAPLI$ = Left$(SS0$, 32)
             If TRIM$(DAPLI$) <> "" Then
                 DAPLI$ = LCase$(AJUSTI$(DAPLI$, 32))
                 Mid$(DAPLI$, 1, 1) = UCase$(Left$(DAPLI$, 1))
                 DAPLI$ = DAPLI$ + Space$(24) + AJUSTI$(APLI$(i%), 8)
                 ' FLESCRIT.List1.AddItem DAPLI$
             End If
         End If
         '
         Close #N2%: N2% = 0
    Next i%
    '
    'If FLESCRIT.List1.ListCount < 1 Then
    '      Call MENSERR(24, "No se Encuentran Archivos de Menúes")
    '      Call FUERASYS
    'End If
    '
    'J% = FLESCRIT.List1.ListCount
    'AREN% = 26 * FLESCRIT.List1.FontSize
    'BASE% = FLESCRIT.List1.Top + FLESCRIT.List1.Height
    'TOPE% = BASE% - AREN% * FLESCRIT.List1.ListCount
    'If TOPE% < 3000 Then TOPE% = 3000
    '
    'FLESCRIT.List1.Top = TOPE%
    'FLESCRIT.List1.Height = BASE% - TOPE%
    '
    FLESCRIT.MousePointer = 1
    '
End Sub
'

