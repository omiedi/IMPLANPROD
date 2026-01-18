Attribute VB_Name = "FORMEDIT_FUNCTIONS"
Public TIFORMED%
'

Sub GRAENTRA()
  '
  If TIFORMED% = 1 Then
    If MODEDI% = 1 Or MODEDI% = 2 Or MODEDI% = 4 Then               ' editar registro
       XXX$ = REGLEIDO(ARCHILOC$, FORVENTA.LITAB.ListIndex + 1)
       If XXX$ = "" Then
          XXX$ = REGBLAN$(ARCHILOC$)
       End If
       '
       For II% = 1 To NCL%
         If PE%(II%) = 1 Or MODEDI% <> 4 Then
           VNUEVO$ = FORMEDIT.ENTRA(II% - 1).Text
           If JT%(II%) = 1 Then
             VNUEVO$ = TRIM$(VNUEVO$)
               'If MG%(II%) > 0 Then
               '    VNUEVO$ = Mid$(VNUEVO$, 1 + MG%(II%))
               'End If
           End If
           If TVL%(CFT%(II%)) = 5 Then
               VNUEVO$ = TRIM$(VNUEVO$)
               If FECHANUM(TRIM$(VNUEVO$)) = 0 Then
                   VNUEVO$ = Space$(8)
               End If
           End If
           If TVL%(CFT%(II%)) = 1 And TRIM$(VNUEVO$) = "0" Then VNUEVO$ = ""
           XXZ$ = CBIN$(VNUEVO$, TVL%(CFT%(II%)), LF%(CFT%(II%)))
           '
           If TRIM$(AE$(II%)) = "MASTER" Or TRIM$(AE$(II%)) = "ECOMAST" Or Left$(AE$(II%), 5) = "LIPRE" Then
             If TRIM$(VNUEVO$) = "" Or CODINT%(VNUEVO$) < 1 Then
               Call MENSERR(24, "Codigo no Válido")
               Call EDLISVEN
               Exit Sub
             End If
           End If
           '
           If TRIM$(AE$(II%)) <> "" Then
              If TRIM$(VNUEVO$) <> "" Then
                 If ECOARCH$(AE$(II%), XXZ$) = "" Then
                    Call MENSERR(24, "Ingreso no Válido\Según Tabla '" + TRIM$(AE$(II%)) + "'.")
                    Call EDLISVEN
                    Exit Sub
                 End If
              End If
           End If
           '
           Mid$(XXX$, PR%(CFT%(II%)), LF%(CFT%(II%))) = XXZ$
         End If
       Next II%
       Call GRAREG(ARCHILOC$, XXX$, FORVENTA.LITAB.ListIndex + 1)
       '
       RELIB$ = Space$(ATOT% + 2)
       For K2% = 1 To NCL%
           If CFT%(K2%) <> 1 Then
             If RTT%(CFT%(K2%)) < 1 Then
               If LF%(CFT%(K2%)) > 0 Then
                   XXZ$ = Mid$(XXX$, PR%(CFT%(K2%)), LF%(CFT%(K2%)))
                   ACLX% = ACL%(K2%) - MG%(K2%)
                   TXT$ = CSTRING$(XXZ$, TVL%(CFT%(K2%)), ACLX%, LDL%(CFT%(K2%)), JT%(K2%))
                   If JT%(K2%) = 2 Then
                       TXT$ = TXT$ + Space$(MG%(K2%))
                     Else
                       TXT$ = Space$(MG%(K2%)) + TXT$
                   End If
                   If CL%(K2%) > 0 Then Mid$(RELIB$, CL%(K2%) - OFS%, LC%(K2%)) = TXT$
                   '
                   If TRIM$(AE$(K2%)) <> "" Then
                      XXE$ = ECOARCH$(AE$(K2%), XXZ$)
                      COLOG% = K2%
                      CLE% = CL%(CE%(COLOG%))
                      LCE% = LC%(CE%(COLOG%))
                      JTE% = JT%(CE%(COLOG%))
                      TXT$ = CSTRING$(XXE$, 6, LCE%, 0, JTE%)
                      If CLE% > 0 Then Mid$(RELIB$, CLE% - OFS%, LCE%) = TXT$
                   End If
               End If
             End If
           End If
       Next K2%
       '
       ILISTA& = FORVENTA.LITAB.ListIndex
       FORVENTA.LITAB.RemoveItem ILISTA&
       FORVENTA.LITAB.AddItem RELIB$, ILISTA&
       Call VERSCROLLVEN
       If ILISTA& < FORVENTA.LITAB.ListCount Then
            FORVENTA.LITAB.ListIndex = ILISTA&
          Else
            FORVENTA.LITAB.ListIndex = FORVENTA.LITAB.ListCount - 1
       End If
       '
    End If
    '
  Else
    '
    If MODEDI% = 1 Or MODEDI% = 2 Or MODEDI% = 4 Then  ' editar registro
       If TIFORTAB% < 1 Then
           XXX$ = REGLEIDO(ARCHILOC$, WEDITAB.LITAB.ListIndex + 1)
         Else
           XXX$ = REGLEIDO(ARCHILOC$, XFORTAB.LITAB.ListIndex + 1)
       End If
       If XXX$ = "" Then
          XXX$ = REGBLAN$(ARCHILOC$)
       End If
       '
       For II% = 1 To NCL%
         If PR%(CFT%(II%)) > 0 Then
           If PE%(II%) = 1 Or MODEDI% <> 4 Then
             VNUEVO$ = TRIM$(FORMEDIT.ENTRA(II% - 1).Text)
             If TVL%(CFT%(II%)) = 5 Then
               If FECHANUM(TRIM$(VNUEVO$)) = 0 Then
                  VNUEVO$ = Space$(8)
               End If
             End If
             XXZ$ = CBIN$(VNUEVO$, TVL%(CFT%(II%)), LF%(CFT%(II%)))
             Mid$(XXX$, PR%(CFT%(II%)), LF%(CFT%(II%))) = XXZ$
           End If
         End If
       Next II%
       If TIFORTAB% <> 1 Then
            Call GRAREG(ARCHILOC$, XXX$, WEDITAB.LITAB.ListIndex + 1)
          Else
            Call GRAREG(ARCHILOC$, XXX$, XFORTAB.LITAB.ListIndex + 1)
       End If
       '
       RELIB$ = Space$(ATOT% + 2)
       For K2% = 1 To NCL%
           If CFT%(K2%) <> 1 Then
             If RTT%(CFT%(K2%)) < 1 Then
               If LF%(CFT%(K2%)) > 0 Then
                   XXZ$ = Mid$(XXX$, PR%(CFT%(K2%)), LF%(CFT%(K2%)))
                   ACLX% = ACL%(K2%) - MG%(K2%)
                   TXT$ = CSTRING$(XXZ$, TVL%(CFT%(K2%)), ACLX%, LDL%(CFT%(K2%)), JT%(K2%))
                   If JT%(K2%) = 2 Then
                       TXT$ = TXT$ + Space$(MG%(K2%))
                     Else
                       TXT$ = Space$(MG%(K2%)) + TXT$
                   End If
                   If CL%(K2%) > 0 Then Mid$(RELIB$, CL%(K2%), LC%(K2%)) = TXT$
                   '
                   If TRIM$(AE$(K2%)) <> "" Then
                      XXE$ = ECOARCH$(AE$(K2%), XXZ$)
                      COLOG% = K2%
                      CLE% = CL%(CE%(COLOG%))
                      LCE% = LC%(CE%(COLOG%))
                      JTE% = JT%(CE%(COLOG%))
                      TXT$ = CSTRING$(XXE$, 6, LCE%, 0, JTE%)
                      If CLE% > 0 Then Mid$(RELIB$, CLE%, LCE%) = TXT$
                   End If
               End If
             End If
           End If
       Next K2%
       '
       If TIFORTAB% <> 1 Then
            ILISTA& = WEDITAB.LITAB.ListIndex
            WEDITAB.LITAB.RemoveItem ILISTA&
            WEDITAB.LITAB.AddItem RELIB$, ILISTA&
            Call VERISCROLL
            If ILISTA& < WEDITAB.LITAB.ListCount Then
                 WEDITAB.LITAB.ListIndex = ILISTA&
               Else
                 WEDITAB.LITAB.ListIndex = WEDITAB.LITAB.ListCount - 1
            End If
         Else
            ILISTA& = XFORTAB.LITAB.ListIndex
            XFORTAB.LITAB.RemoveItem ILISTA&
            XFORTAB.LITAB.AddItem RELIB$, ILISTA&
            Call VERISCROLL
            If ILISTA& < XFORTAB.LITAB.ListCount Then
                 XFORTAB.LITAB.ListIndex = ILISTA&
               Else
                 XFORTAB.LITAB.ListIndex = XFORTAB.LITAB.ListCount - 1
            End If
       End If
       '
    End If
    '
  End If
  CONTROLGRABA% = 1
  '
End Sub
'

Sub CANCELENTRA()
  If TIFORMED% = 1 Then
    ILISTA& = FORVENTA.LITAB.ListIndex
    RELIB$ = FORVENTA.LITAB.Text
    If TRIM$(RELIB$) = "" Then
         FORVENTA.LITAB.RemoveItem ILISTA&
         For U& = ILISTA& + 2 To ULTREG&(ARCHILOC$)
            XXX$ = REGLEIDO$(ARCHILOC$, U&)
            Call GRAREG(ARCHILOC$, XXX$, U& - 1)
         Next U&
         Call GRAREG(ARCHILOC$, String$(RECLEN%(ARCHILOC$), 0), FORVENTA.LITAB.ListCount + 1)
         Call CIERRARCH(ARCHILOC$)
    End If
    '
    If ILISTA& < FORVENTA.LITAB.ListCount Then
          FORVENTA.LITAB.ListIndex = ILISTA&
        Else
          FORVENTA.LITAB.ListIndex = FORVENTA.LITAB.ListCount - 1
          ILISTA& = FORVENTA.LITAB.ListIndex
    End If
    '
    MODEDI% = 0
    YAPRE% = 0
    FORMEDIT.Hide
    Call FRESHALL
    '
  Else
    '
    If TIFORTAB% <> 1 Then
         ILISTA& = WEDITAB.LITAB.ListIndex
         RELIB$ = WEDITAB.LITAB.Text
       Else
         ILISTA& = XFORTAB.LITAB.ListIndex
         RELIB$ = XFORTAB.LITAB.Text
    End If
    If TRIM$(RELIB$) = "" Then
         If TIFORTAB% <> 1 Then
              WEDITAB.LITAB.RemoveItem ILISTA&
              REGRARE& = WEDITAB.LITAB.ListCount + 1
            Else
              XFORTAB.LITAB.RemoveItem ILISTA&
              REGRARE& = XFORTAB.LITAB.ListCount + 1
         End If
         For U& = ILISTA& + 2 To ULTREG&(ARCHILOC$)
            XXX$ = REGLEIDO$(ARCHILOC$, U&)
            Call GRAREG(ARCHILOC$, XXX$, U& - 1)
         Next U&
         Call GRAREG(ARCHILOC$, String$(RECLEN%(ARCHILOC$), 0), REGRARE&)
         Call CIERRARCH(ARCHILOC$)
    End If
    '
    If TIFORTAB% <> 1 Then
         If ILISTA& < WEDITAB.LITAB.ListCount Then
               WEDITAB.LITAB.ListIndex = ILISTA&
             Else
               WEDITAB.LITAB.ListIndex = WEDITAB.LITAB.ListCount - 1
               ILISTA& = WEDITAB.LITAB.ListIndex
         End If
      Else
         If ILISTA& < XFORTAB.LITAB.ListCount Then
               XFORTAB.LITAB.ListIndex = ILISTA&
             Else
               XFORTAB.LITAB.ListIndex = XFORTAB.LITAB.ListCount - 1
               ILISTA& = XFORTAB.LITAB.ListIndex
         End If
    End If
    '
    MODEDI% = 0
    YAPRE% = 0
    FORMEDIT.Hide
    Call FRESHALL
    '
  End If
  '
End Sub
'

Sub SETNEXTFOC(IENTRA%)
    '
    NUCOLU% = IENTRA% + 1
    For COLUI% = NUCOLU% To NCL%
       If PE%(COLUI%) = 1 Or (MODEDI% <> 4 And PE%(COLUI%) <> 3) Then
          FORMEDIT.ENTRA(COLUI% - 1).SetFocus
          Exit Sub
       End If
    Next COLUI%
    '
    FORMEDIT.Command1.SetFocus
    '
End Sub
'

Sub VERISCROLL()
    '
    Static LIANT&
    If TIFORTAB% <> 1 Then
         If WEDITAB.LITAB.ListCount <> LIANT& Then
             LIANT& = WEDITAB.LITAB.ListCount
             If WEDITAB.LITAB.ListCount > CARELI% Then
                 WEDITAB.VScroll1.Visible = False
               Else
                 WEDITAB.VScroll1.Visible = True
             End If
         End If
      Else
         If XFORTAB.LITAB.ListCount <> LIANT& Then
             LIANT& = XFORTAB.LITAB.ListCount
             If XFORTAB.LITAB.ListCount > CARELI% Then
                 XFORTAB.VScroll1.Visible = False
               Else
                 XFORTAB.VScroll1.Visible = True
             End If
         End If
    End If
    '
End Sub
'

Sub VERSCROLLVEN()
    '
    Static LIANT&
    If FORVENTA.LITAB.ListCount <> LIANT& Then
        LIANT& = FORVENTA.LITAB.ListCount
        If FORVENTA.LITAB.ListCount > CARELI% Then
            FORVENTA.VScroll1.Visible = False
          Else
            FORVENTA.VScroll1.Visible = True
        End If
    End If
    '
End Sub

Sub EDLISVEN()
    '
    If FORVENTA.LITAB.ListIndex < 0 Then
        Exit Sub
    End If
    '
    If FORVENTA.LITAB.ListIndex >= FORVENTA.LITAB.ListCount Then
        Exit Sub
    End If
    '
    MODEDI% = 4
    NUREGIS& = FORVENTA.LITAB.ListIndex + 1
    REACTIVO$ = REGLEIDO$(ARCHILOC$, NUREGIS&)
    '
    Call PRENTRAVEN
    '
End Sub
'
Sub BULISVEN()
    '
    Call OPNOIN("")
    Exit Sub
    '
    If FORVENTA.LITAB.ListCount < 1 Then
        Exit Sub
    End If
    MODEDI% = 5
    Call PRENTRAVEN
    '
End Sub
'
Sub PRENTRAVEN()
    '
    Static FORMANT$, POSEDI%
    Screen.MousePointer = 11
    ACAREN = FORVENTA.LITAB.Height / CARELI%
    RENGLOAC = FORVENTA.LITAB.ListIndex - FORVENTA.LITAB.TopIndex + 1
    If RENGLOAC < 1 Or RENGLOAC > CARELI% Then
        Screen.MousePointer = 1
        Exit Sub
    End If
    YAPRE% = 1
    BORDESUPE = FORVENTA.Top + FORVENTA.LITAB.Top + (RENGLOAC - 1) * ACAREN
    If RENGLOAC > 1 Then BORDESUPE = BORDESUPE + 70
    If TRIM$(FORVENTA.Caption) <> "" Then
        BORDESUPE = BORDESUPE + ACAREN
    End If
    KPLA% = 0: PRICA% = 0
    '
    If FORMA$ = FORMANT$ Then
        FORMEDIT.Top = BORDESUPE
        For II% = 1 To NCL%
           FORMEDIT.ENTRA(II% - 1).Text = ""
        Next II%
        GoTo 90
    End If
    FORMANT$ = FORMA$
    '
    BLANCO = &H80000005
    For II% = 0 To 31
        '
        FORMEDIT.ENTRA(II%).FontSize = 7
        FORMEDIT.ENTRA(II%).FontBold = True
        FORMEDIT.ENTRA(II%).Text = ""
        FORMEDIT.ENTRA(II%).Height = 285
        FORMEDIT.ENTRA(II%).Top = 0
        FORMEDIT.ENTRA(II%).BackColor = BLANCO
        FORMEDIT.ENTRA(II%).Visible = False
        FORMEDIT.BOTSEL(II%).Visible = False
        '
    Next II%
    '
    ALTUCON% = 285
    FORMEDIT.Top = BORDESUPE
    FORMEDIT.Left = FORVENTA.Left
    FORMEDIT.Width = FORVENTA.Width
    FORMEDIT.Height = ALTUCON% * 1.2
    '
    If MODEDI% = 5 Then
        ' FORMEDIT.Top = FORVENTA.Top + FORVENTA.FRAMETOT.Top
        FORMEDIT.Top = FORVENTA.Top + FORVENTA.LITAB.Top + (CARELI% + 1) * ACAREN + (FORVENTA.FRAMETOT.Height - FORMEDIT.Height) / 2
    End If
    '
    FORMEDIT.Command1.Left = FORVENTA.FRAMEBOT2.Left
    FORMEDIT.Command1.Width = FORVENTA.BOTAGRE.Width * 0.6
    FORMEDIT.Command2.Left = FORMEDIT.Command1.Left + FORMEDIT.Command1.Width
    FORMEDIT.Command2.Width = FORVENTA.BOTAGRE.Width * 0.4
    FORMEDIT.Command1.Height = ALTUCON%
    FORMEDIT.Command2.Height = ALTUCON%
    '
    POSEDI% = 0
    For II% = 1 To NCL%
       '
       If CL%(II%) > 0 Then
          'FORMEDIT.ENTRA(II% - 1).Left = COEFH * (CL%(II%) - 1) - FORMEDIT.Left
          FORMEDIT.ENTRA(II% - 1).Left = COEFH * (CL%(II%) - OFS% - 1) ' - FORMEDIT.Left
          FORMEDIT.ENTRA(II% - 1).Top = 0
          FORMEDIT.ENTRA(II% - 1).Width = COEFH * (ACL%(II%) + 0.9)
          FORMEDIT.ENTRA(II% - 1).Visible = True
          FORMEDIT.ENTRA(II% - 1).FontName = FORVENTA.LITAB.FontName
          FORMEDIT.ENTRA(II% - 1).FontSize = FORVENTA.LITAB.FontSize
          FORMEDIT.ENTRA(II% - 1).FontBold = FORVENTA.LITAB.FontBold
          If PE%(II%) <> 1 Then
            If PE%(II%) = 3 Or MODEDI% = 4 Then
              FORMEDIT.ENTRA(II% - 1).BackColor = &H80000013
            End If
          End If
       End If
       '
       If FORMEDIT.ENTRA(II% - 1).BackColor = BLANCO Then
           POSEDI% = 1
       End If
       '
    Next II%
    '
90  If MODEDI% = 4 Then                     ' editar registro
        RELIB$ = FORVENTA.LITAB.Text
        XXX$ = REGLEIDO(ARCHILOC$, FORVENTA.LITAB.ListIndex + 1)
        For II% = 1 To NCL%
           If CL%(II%) > 0 Then
              FORMEDIT.ENTRA(II% - 1).Text = Mid$(RELIB$, CL%(II%) - OFS%, ACL%(II%))
              If TVL%(CFT%(II%)) <= 4 Then
                  If Left$(TRIM$(FORMEDIT.ENTRA(II% - 1)), 1) = "*" Then
                      XXZ$ = Mid$(XXX$, PR%(CFT%(II%)), LF%(CFT%(II%)))
                      FORMEDIT.ENTRA(II% - 1).Text = TRIM$(CSTRING$(XXZ$, TVL%(CFT%(II%)), 15, LDL%(CFT%(II%)), 1))
                  End If
              End If
           End If
        Next II%
    End If
    '
    Screen.MousePointer = 1
    If POSEDI% = 1 Or MODEDI% < 4 Then
        FORMEDIT.Refresh
        On Error Resume Next
        COLUAC% = 0
        SMP = Screen.MousePointer
        Screen.MousePointer = 1
        FORMEDIT.Show 1
        Screen.MousePointer = SMP
        Call FRESHALL
        On Error GoTo 0
      Else
        FORVENTA.BOTEDIT.Enabled = False
    End If
    '
End Sub
'

