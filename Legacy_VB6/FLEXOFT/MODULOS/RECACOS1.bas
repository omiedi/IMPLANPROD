Attribute VB_Name = "RECACOS1"
 Public CALETRA%
 
 Sub RECACOS()
     '
     Screen.MousePointer = 11
     QQQ% = 1024
     Dim DP%(16384), DQ%(16384)
     Dim COMAQ$(1024), CUHOR(1024)
     '
     CAMAQ% = ULTREG&("PADMAQ")
     For UI& = 1 To ULTREG&("PADMAQ")
        R0$ = REGLEIDO$("PADMAQ", UI&)
        COMAQ$(UI&) = TRIM$(Left$(R0$, 6))
        CUHOR(UI&) = CVS(Mid$(R0$, 93, 4))
     Next UI&
     Call CIERRARCH("PADMAQ")
     '
62   HOI% = HOY(HO$)
AAA = CODINT%("12710001")
     NVI% = NUINV%
     NMM% = NUMAS%
     ' IDENTIFICANDO LOS ITEMS SIN EXPLOSION
     '
     FIN& = ULTREG&("ESTRUNUE")
     For I& = 1 To FIN&
        Call TRACE(20, I&, FIN&)
        R0$ = REGLEIDO$("ESTRUNUE", I&)
        R2$ = Mid$(R0$, 3, 120)
        CI% = CVI(Left$(R0$, 2))
        If CI% > 0 Then
            If CI% <= NMM% Then
                For K% = 1 To 120 Step 8
                    DPSX$ = Mid$(R2$, K%, 8)
                    If CVI(Left$(DPSX$, 2)) < 1 Or CVI(Left$(DPSX$, 2)) > NMM% Then GoTo 105
                    If CVS(Mid$(DPSX$, 3)) >= 0.0005 Then
                        DP%(CI%) = 1
                        DQ%(CI%) = 1
                        GoTo 105
                    End If
                Next K%
            End If
        End If
105  Next I&
     '
110  N9% = FILEOPEN%(LU$ + "COCALCU.DAT", 0, 128)
     'OPEN "R",#N9%,LU$+"COCALCU.DAT",128
     'FIELD #N9%,2 AS Q1$,2 AS Q3$,2 AS Q4$,4 AS Q11$,4 AS Q12$,4 AS Q13$,4 AS Q21$,4 AS Q22$,4 AS Q23$,4 AS Q31$,4 AS Q32$,4 AS Q33$,4 AS Q41$,4 AS Q42$,4 AS Q43$,74 AS Q99$
     Dim Q0 As String * 128
     N19% = FILEOPEN%(LU$ + "MODIREC.DAT", 0, 128)
     Dim Q9 As String * 128
     For II% = 1 To NMM%
        Get #N9%, II%, Q0$
        LSet Q0$ = String$(128, 0)
        Put #N9%, II%, Q0$
        Get #N19%, II%, Q9$
        LSet Q9$ = String$(127, 0) + Chr$(1)
        Put #N19%, II%, Q9$
     Next II%
115  CANA& = 0

     ' CARGANDO COSTOS DE SERVICIOS SUMINISTRADOS POR TERCEROS
     N6% = FILEOPEN(LU$ + "LPREPRO.DAT", 0, 48)
     Dim P0 As String * 48
     FIN& = ULTREG&("LPREPRO")
     For I& = 1 To FIN&
        Call TRACE(20, I&, FIN&)
        KI% = I&
        P0$ = REGLEIDO$("LPREPRO", I&)
125     CI% = CVI(Mid$(P0$, 3, 2))
        COU = CVS(Mid$(P0$, 37, 4))
        If CI% > 0 Then
            If CI% <= NMM% Then
                Get #N9%, CI%, Q0$
                RQ0$ = Q0$
                V31 = CVS(Mid$(Q0$, 31, 4)) + COU
                V33 = CVS(Mid$(Q0$, 39, 4)) + COU
                V41 = CVS(Mid$(Q0$, 43, 4)) + COU
                V43 = CVS(Mid$(Q0$, 51, 4)) + COU
                Call REPLA(RQ0$, MKS$(V31), 31, 4)
                Call REPLA(RQ0$, MKS$(V33), 39, 4)
                Call REPLA(RQ0$, MKS$(V41), 43, 4)
                Call REPLA(RQ0$, MKS$(V43), 51, 4)
                '
                Q99$ = Mid$(RQ0$, 55, 74)
                POIN$ = ""
                For U% = 1 To 74 Step 2
                    U1% = CVI(Mid$(Q99$, U%, 2))
                    If U1% > 0 Then
                        POIN$ = POIN$ + MKI$(U1%)
                    End If
                Next U%
                POIN$ = POIN$ + MKI$(KI%) + String$(72, 0)
                Call REPLA(RQ0$, POIN$, 55, 74)
                '
                LSet Q0$ = RQ0$
                Put #N9%, CI%, Q0$
            End If
        End If
     Next I&
     '
     'CARGANDO COSTOS DE MANO DE OBRA DIRECTA
     FIN& = ULTREG&("SECOPER")
     For I& = 1 To FIN&
        Call TRACE(20, I&, FIN&)
        KI% = I&
        SCP$ = REGLEIDO$("SECOPER", I&)
        CI% = CVI(Left$(SCP$, 2))
        MAQUI$ = TRIM$(Mid$(SCP$, 53, 6))
        CUOTHOR = 0
        'If TRIM$(MAQUI$) <> "" Then
           For KK% = 1 To CAMAQ%
              If COMAQ$(KK%) = MAQUI$ Then
                 CUOTHOR = CUHOR(KK%)
              End If
           Next KK%
        'End If
        '
        LOTESTA = LOTESTAN(CI%)
        If LOTESTA < 1 Then LOTESTA = 1
        CALOTES = 1
        MINPREP = 60 * CVS(Mid$(SCP$, 59, 4)) * CALOTES
        MINFIJO = 60 * CVS(Mid$(SCP$, 63, 4)) * CALOTES
        MINVARI = LOTESTA * CVS(Mid$(SCP$, 67, 4))
        MINDEMO = (MINFIJO + MINVARI) * CVS(Mid$(SCP$, 75, 4)) / 100
        MINLIMP = 60 * CVS(Mid$(SCP$, 79, 4)) * CALOTES
        MINLOTE = MINPREP + MINFIJO + MINVARI + MINDEMO + MINLIMP
        HORAPROD = MINLOTE / 60
        CAOPS = CVS(Mid$(SCP$, 83, 4))
        CATOP% = Asc(Mid$(SCP$, 87, 1))
        CUOTOPE = VALHORA(CATOP%)
        COU = (CUOTHOR + CUOTOPE * CAOPS) * HORAPROD / LOTESTA
        '
        If CI% > 0 Then
            If CI% <= NMM% Then
                Get #N19%, CI%, Q9$
                RQ0$ = Q9$
                V61 = CVS(Mid$(Q9$, 3, 4)) + COU
                V63 = CVS(Mid$(Q9$, 11, 4)) + COU
                Call REPLA(RQ0$, MKS$(V61), 3, 4)
                Call REPLA(RQ0$, MKS$(V63), 11, 4)
                LSet Q9$ = RQ0$
                Put #N19%, CI%, Q9$
            End If
        End If
     Next I&
     '
     ' TRANSFIRIENDO COSTOS DE REPOSICION DE ITEMS SIN EXPLOSION
     FIN& = NUMAS%
     Screen.MousePointer = 11
     For I& = 1 To FIN&
        Call TRACE(20, I&, FIN&)
        M0$ = REGLEIDO$("MASTER", I& + 1)
        CI% = CVI(Left$(M0$, 2))
        UNI$ = Mid$(M0$, 109, 2)
        Get #N9%, CI%, Q0$
        RQ0$ = Q0$
        Call REPLA(RQ0$, MKI$(CI%), 1, 2)
        Call REPLA(RQ0$, UNI$, 3, 2)
        Call REPLA(RQ0$, "$ ", 5, 2)
        LSet Q0$ = RQ0$
        Put #N9%, CI%, Q0$
        '
        If DP%(CI%) = 0 Then
            COU = CVS(Mid$(M0$, 111, 4))
            Get #N19%, CI%, Q9$
            If CVS(Mid$(Q9$, 3, 4)) > 0.000005 Then
               COU = 0 ' si tiene M.Obra no puede tener Costo de Materiales
            End If
            MONE$ = "$"
            If Asc(Mid$(M0$, 95, 1)) = 2 Then MONE$ = "U$"
            If Asc(Mid$(M0$, 95, 1)) = 3 Then MONE$ = "DM"
            '
            Call REPLA(RQ0$, MONE$, 5, 2)
            If UNI$ = "U." Then
                 Call REPLA(RQ0$, MKS$(COU), 19, 4)
                 Call REPLA(RQ0$, MKS$(COU), 27, 4)
               Else
                 Call REPLA(RQ0$, MKS$(COU), 7, 4)
                 Call REPLA(RQ0$, MKS$(COU), 15, 4)
            End If
            Call REPLA(RQ0$, MKS$(COU), 43, 4)
            Call REPLA(RQ0$, MKS$(COU), 51, 4)
            LSet Q0$ = RQ0$
            Put #N9%, CI%, Q0$

            If COU > 0.00005 Then
                CANA& = CANA& + 1
                RW0$ = String$(96, 0)
                Call REPLA(RW0$, MKI$(CI%), 1, 2)
                CIXI% = CI%
                Call REPLA(RW0$, CODEXT$(CIXI%), 3, 16)
                CIXI% = CI%
                Call REPLA(RW0$, DESCRIT$(CIXI%), 19, 48)
                Call REPLA(RW0$, MONE$, 67, 2)
                Call REPLA(RW0$, MKS$(COU), 69, 4)
                Call REPLA(RW0$, MKS$(1), 73, 4)
                Call REPLA(RW0$, MKS$(COU), 77, 4)
                Call REPLA(RW0$, MKS$(0), 81, 4)
                Call REPLA(RW0$, MKS$(0), 85, 4)
                Call REPLA(RW0$, MKS$(COU), 89, 4)
                Call REPLA(RW0$, MKS$(0), 93, 4)
                '
                Call GRAREG("ANACOS", RW0$, CANA&)
            End If

130         If CVS(Mid$(Q0$, 31, 4)) > 0 Then
                POIN$ = Mid$(Q0$, 55, 74)
                For U% = 1 To Len(POIN$) Step 2
                    POINI% = CVI(Mid$(POIN$, U%, 2))
                    If POINI% > 0 Then
                        CIXI% = CI%
                        Get #N6%, POINI%, P0$
                        DE$ = TRIM$(Mid$(P0$, 5, 30))
                        MN$ = Mid$(P0$, 35, 2)
                        COU = CVS(Mid$(P0$, 37, 4))
                        CANA& = CANA& + 1
                        '
                        RW0$ = String$(96, 0)
                        Call REPLA(RW0$, MKI$(CI%), 1, 2)
                        CIXI% = CI%
                        Call REPLA(RW0$, CODEXT$(CIXI%), 3, 16)
                        CIXI% = CI%
                        Call REPLA(RW0$, DE$, 19, 48)
                        Call REPLA(RW0$, MN$, 67, 2)
                        Call REPLA(RW0$, MKS$(COU), 69, 4)
                        Call REPLA(RW0$, MKS$(1), 73, 4)
                        Call REPLA(RW0$, MKS$(0), 77, 4)
                        Call REPLA(RW0$, MKS$(0), 81, 4)
                        Call REPLA(RW0$, MKS$(COU), 85, 4)
                        Call REPLA(RW0$, MKS$(COU), 89, 4)
                        Call REPLA(RW0$, MKS$(0), 93, 4)
                        Call GRAREG("ANACOS", RW0$, CANA&)
                        '
                    End If
                Next U%
            End If
        End If
     Next I&
     ' RE-CALCULANDO COSTOS STANDARD DE ITEMS SEGUN EXPLOSION

140  RECA% = 0
     FIN& = ULTREG&("ESTRUNUE")
     Screen.MousePointer = 11
     For I& = 1 To FIN&
         Call TRACE(20, I&, FIN&)
         R0$ = Left$(REGLEIDO$("ESTRUNUE", I&), 122)
         If I& <= ULTREG&("ESTRUDOS") Then
              R0$ = R0$ + REGLEIDO$("ESTRUDOS", I&)
            Else
              R0$ = R0$ + String$(256, 0)
         End If
         CI% = CVI(Left$(R0$, 2))
         DPS$ = Mid$(R0$, 3)
         If CI% < 1 Or CI% > NUMAS% Then GoTo 199

142      If DP%(CI%) = 0 Then GoTo 199                          ' ya fue calculado

         CIXI% = CI%
         CABAS = 1: SUMP = 0: SUMF = 0: SUS3 = 0: SUMO = 0
If CI% = 1 Then
AAA = BBB
End If
145      For K% = 1 To 376 Step 8
             DPSX$ = Mid$(DPS$, K%, 8)
             CI1% = CVI(DPSX$)
             If CI1% < 1 Or CI1% > NUMAS% Then GoTo 149    ' terminar
             '
             If CI1% = CI% Then
                If COMALTER%("Error de Recursividad en la Fórmula\Correspondiente a " + TRIM$(CODEXT$(CI%)) + "\" + DESCRIT0$(CI%) + ".\\Cancelar\Continuar") <> 2 Then
                   Call GRACONTROL("COSTOS", "FECALCOS", "00/00/00")
                   Call GRACONTROL("COSTOS", "HOCALCOS", "00:00:00")
                   FORCOSTO.Label34.Caption = " " + HOS$
                   FORCOSTO.Label34.Refresh
                   FORCOSTO.Label35.Caption = " " + Time$
                   FORCOSTO.Label35.Refresh
                   GoTo 3800
                End If
             End If
             '
             If DP%(CI1%) > 0 Then GoTo 199                ' aun no fue calculado
             CANT7 = CABAS * CVS(Mid$(DPSX$, 3))
             Get #N9%, CI1%, Q0$
             Get #N19%, CI1%, Q9$
             SUMP = SUMP + CANT7 * CVS(Mid$(Q0$, 15, 4))
             SUMF = SUMF + CANT7 * CVS(Mid$(Q0$, 27, 4))
             SUS3 = SUS3 + CANT7 * CVS(Mid$(Q0$, 39, 4))
             MOOX = CVS(Mid$(Q9$, 11, 4))
             SUMO = SUMO + CANT7 * CVS(Mid$(Q9$, 11, 4))
'BBBB$ = CODEXT$(CI1%)
         Next K%
         '
149      Get #N9%, CI%, Q0$
         RQ0$ = Q0$
         Get #N19%, CI%, Q9$
         RQ9$ = Q9$
         '
         V12 = CVS(Mid$(Q0$, 11, 4)) + SUMP
         V13 = CVS(Mid$(Q0$, 15, 4)) + SUMP
         V22 = CVS(Mid$(Q0$, 23, 4)) + SUMF
         V23 = CVS(Mid$(Q0$, 27, 4)) + SUMF
         V32 = CVS(Mid$(Q0$, 35, 4)) + SUS3
         V33 = CVS(Mid$(Q0$, 39, 4)) + SUS3
         V42 = CVS(Mid$(Q0$, 47, 4)) + SUMP + SUMF + SUS3
         V43 = CVS(Mid$(Q0$, 51, 4)) + SUMP + SUMF + SUS3
         V62 = CVS(Mid$(Q9$, 7, 4)) + SUMO
         V63 = CVS(Mid$(Q9$, 11, 4)) + SUMO
         '
         Call REPLA(RQ0$, MKS$(V12), 11, 4)
         Call REPLA(RQ0$, MKS$(V13), 15, 4)
         Call REPLA(RQ0$, MKS$(V22), 23, 4)
         Call REPLA(RQ0$, MKS$(V23), 27, 4)
         Call REPLA(RQ0$, MKS$(V32), 35, 4)
         Call REPLA(RQ0$, MKS$(V33), 39, 4)
         Call REPLA(RQ0$, MKS$(V42), 47, 4)
         Call REPLA(RQ0$, MKS$(V43), 51, 4)
         '
         Call REPLA(RQ9$, MKS$(V62), 7, 4)
         Call REPLA(RQ9$, MKS$(V63), 11, 4)
         '
         LSet Q0$ = RQ0$
         Put #N9%, CI%, Q0$
         '
         LSet Q9$ = RQ9$
         Put #N19%, CI%, Q9$
         QBAS = 1: CIX0% = CI%
         Call EXPLOMUL(CI%)
         Screen.MousePointer = 11
         For J% = 1 To CONTAMUL%
             CII% = CIK%(J%)
             If CII% > 0 Then
                If CII% <= NUMAS% Then
                   Get #N9%, CII%, Q0$
                   If DQ%(CII%) = 0 Then
                      COU = CVS(Mid$(Q0$, 7, 4)) + CVS(Mid$(Q0$, 19, 4))
                      CANA& = CANA& + 1
                      RW0$ = String$(96, 0)
                      Call REPLA(RW0$, MKI$(CIX0%), 1, 2)
                      CIXI% = CII%
                      Call REPLA(RW0$, CODEXT$(CIXI%), 3, 16)
                      CIXI% = CII%
                      Call REPLA(RW0$, DESCRIT$(CIXI%), 19, 48)
                      Call REPLA(RW0$, "$ ", 67, 2)
                      Call REPLA(RW0$, MKS$(COU), 69, 4)
                      Call REPLA(RW0$, MKS$(CANT(J%)), 73, 4)
                      Call REPLA(RW0$, MKS$(COU * CANT(J%)), 77, 4)
                      Call REPLA(RW0$, MKS$(0), 81, 4)
                      Call REPLA(RW0$, MKS$(0), 85, 4)
                      Call REPLA(RW0$, MKS$(COU * CANT(J%)), 89, 4)
                      Call REPLA(RW0$, MKS$(0), 93, 4)
                      Call GRAREG("ANACOS", RW0$, CANA&)
                   End If
                   '
                   If CVS(Mid$(Q0$, 31, 4)) > 0 Then
                      POIN$ = Mid$(Q0$, 55, 74)
                      For U% = 1 To Len(POIN$) Step 2
                         POINI% = CVI(Mid$(POIN$, U%, 2))
                         If POINI% > 0 Then
                            CIXI% = CI%
                            Get #N6%, POINI%, P0$
                            DE$ = TRIM$(Mid$(P0$, 5, 30))
                            MN$ = Mid$(P0$, 35, 2)
                            COU = CVS(Mid$(P0$, 37, 4))
                            CANA& = CANA& + 1
                            '
                            RW0$ = String$(96, 0)
                            Call REPLA(RW0$, MKI$(CIX0%), 1, 2)
                            CIXI% = CII%
                            Call REPLA(RW0$, CODEXT$(CIXI%), 3, 16)
                            CIXI% = CII%
                            Call REPLA(RW0$, DE$, 19, 48)
                            Call REPLA(RW0$, MN$, 67, 2)
                            Call REPLA(RW0$, MKS$(COU), 69, 4)
                            Call REPLA(RW0$, MKS$(CANT(J%)), 73, 4)
                            Call REPLA(RW0$, MKS$(0), 77, 4)
                            Call REPLA(RW0$, MKS$(0), 81, 4)
                            Call REPLA(RW0$, MKS$(COU * CANT(J%)), 85, 4)
                            Call REPLA(RW0$, MKS$(COU * CANT(J%)), 89, 4)
                            Call REPLA(RW0$, MKS$(0), 93, 4)
                            Call GRAREG("ANACOS", RW0$, CANA&)
                            '
                         End If
                      Next U%
                   End If
                End If
             End If
         Next J%
         '
         DP%(CIX0%) = 0
         RECA% = 1
         '
199  Next I&
     If RECA% > 0 Then GoTo 140

210  JJ& = 0
     For I& = 1 To NUINV%
        I0$ = REGLEIDO$("INVERT", I& + 1)
        COD$ = Left$(I0$, 16): CI% = CVI(Mid$(I0$, 17, 2))
        If Asc(Mid$(COD$, 16, 1)) <> 96 Then
            If CI% > 0 Then
                If CI% <= NUMAS% Then
                    Get #N9%, CI%, Q0$
                    Z$ = Left$(Q0$, 64)
                    JJ& = JJ& + 1
                    Call GRAREG("LICOCAL", Z$, JJ&)
                End If
            End If
        End If
     Next I&
     '
     Call SETULTREG("LICOCAL", JJ&)
     Call SETULTREG("ANACOS", CANA&)
     '
     FEX = HOY(HOS$)
     Call GRACONTROL("COSTOS", "FECALCOS", HOS$)
     Call GRACONTROL("COSTOS", "HOCALCOS", Time$)
     '
     FORCOSTO.Label34.Caption = " " + HOS$
     FORCOSTO.Label34.Refresh
     FORCOSTO.Label35.Caption = " " + Time$
     FORCOSTO.Label35.Refresh
     '
3800 Call CIERRARCH("*.*")
     Close #N9%
     Close #N19%
3900 Screen.MousePointer = 1
     '
End Sub

Sub MUECOS(CI%)
   RFF$ = RECFILT$("COCALCU", 1, MKI$(CI%))
   If Len(RFF$) >= 4 Then
      I& = CVS(RFF$)
   'For I& = 1 To ULTREG&("COCALCU")
      Q0$ = REGLEIDO$("COCALCU", I&)
      Q9$ = REGLEIDO$("MODIREC", I&)
      If CVI(Left$(Q0$, 2)) = CI% Then
         FORCOSTO.Label5.Caption = TRIM$(CSTRING$(Mid$(Q0$, 7, 4), 3, 12, 4, 2))
         FORCOSTO.Label5.Refresh
         FORCOSTO.Label20.Caption = TRIM$(CSTRING$(Mid$(Q0$, 11, 4), 3, 12, 4, 2))
         FORCOSTO.Label20.Refresh
         FORCOSTO.Label7.Caption = TRIM$(CSTRING$(Mid$(Q0$, 19, 4), 3, 12, 4, 2))
         FORCOSTO.Label7.Refresh
         FORCOSTO.Label19.Caption = TRIM$(CSTRING$(Mid$(Q0$, 23, 4), 3, 12, 4, 2))
         FORCOSTO.Label19.Refresh
         FORCOSTO.Label9.Caption = TRIM$(CSTRING$(Mid$(Q0$, 31, 4), 3, 12, 4, 2))
         FORCOSTO.Label9.Refresh
         FORCOSTO.Label18.Caption = TRIM$(CSTRING$(Mid$(Q0$, 35, 4), 3, 12, 4, 2))
         FORCOSTO.Label18.Refresh
         FORCOSTO.Label17.Caption = TRIM$(CSTRING$(Mid$(Q9$, 3, 4), 3, 12, 4, 2))
         FORCOSTO.Label17.Refresh
         FORCOSTO.Label23.Caption = TRIM$(CSTRING$(Mid$(Q9$, 7, 4), 3, 12, 4, 2))
         FORCOSTO.Label23.Refresh
         '
         SUMA1 = 0
         SUMA1 = SUMA1 + Val(FORCOSTO.Label5.Caption)
         SUMA1 = SUMA1 + Val(FORCOSTO.Label7.Caption)
         SUMA1 = SUMA1 + Val(FORCOSTO.Label9.Caption)
         SUMA1 = SUMA1 + Val(FORCOSTO.Label20.Caption)
         SUMA1 = SUMA1 + Val(FORCOSTO.Label19.Caption)
         SUMA1 = SUMA1 + Val(FORCOSTO.Label18.Caption)
         FORCOSTO.Label10.Caption = TRIM$(CSTRING$(MKS$(SUMA1), 3, 12, 4, 2))
         FORCOSTO.Label10.Refresh
         '
         SUMA2 = 0
         SUMA2 = SUMA2 + Val(FORCOSTO.Label17.Caption)
         SUMA2 = SUMA2 + Val(FORCOSTO.Label23.Caption)
         FORCOSTO.Label11.Caption = TRIM$(CSTRING$(MKS$(SUMA2), 3, 12, 4, 2))
         FORCOSTO.Label11.Refresh
         '
         SUMA3 = SUMA1
         SUMA3 = SUMA3 + Val(FORCOSTO.Label11.Caption)
         FORCOSTO.Label27.Caption = TRIM$(CSTRING$(MKS$(SUMA3), 3, 12, 4, 2))
         FORCOSTO.Label27.Refresh
         '
         SUMA4 = SUMA3 * (Val(FORCOSTO.Text4.Text)) / 100
         FORCOSTO.Label29.Caption = TRIM$(CSTRING$(MKS$(SUMA4), 3, 12, 4, 2))
         FORCOSTO.Label29.Refresh
         '
         SUMA5 = SUMA3 + SUMA4
         FORCOSTO.Label32.Caption = TRIM$(CSTRING$(MKS$(SUMA5), 3, 12, 4, 2))
         FORCOSTO.Label32.Refresh
         '
      End If
   End If
   'Next I&
End Sub

Sub ACTCOSREPMAT() ' Actualiza Costos de Reposicion por Articulo
   '
   Dim DP%(16384)
   HOI% = HOY(HO$)
   '
   Screen.MousePointer = 11
   '
   FIN& = ULTREG&("ESTRUNUE")
   For I& = 1 To FIN&
      Call TRACE(20, I&, FIN&)
      XX$ = REGLEIDO$("ESTRUNUE", I&)
      CIIX% = CVI(Left$(XX$, 2))
      R2$ = Mid$(XX$, 3, 120)
      If CIIX% > 0 Then
         If CIIX% <= NUMAS% Then
            For K% = 1 To 120 Step 8
               DPSX$ = Mid$(R2$, K%, 8)
               If CVI(Left$(DPSX$, 2)) > 0 Then
                  If CVI(Left$(DPSX$, 2)) <= NUMAS% Then
                     If Abs(CVS(Mid$(DPSX$, 3, 4))) >= 0.000005 Then
                        DP%(CIIX%) = 1
                        GoTo 15
                     End If
                  End If
               End If
            Next K%
         End If
      End If
15 Next I&
   '
   FIN& = ULTREG&("CACORE")
   For U& = 1 To FIN&
      Call TRACE(20, U&, FIN&)
      XX$ = REGLEIDO$("CACORE", U&)
      CIIX% = CVI(Left$(XX$, 2))
      If CIIX% > 0 Then
         If CIIX% <= NUMAS% Then
            COU = CVS(Mid$(XX$, 7, 4))
            FECO% = CVI(Mid$(XX$, 11, 2))
            PROVE% = CVI(Mid$(XX$, 13, 2))
            If COU > 0 Then
               REGMA& = CIIX% + 1
               YY$ = REGLEIDO$("MASTER", REGMA&)
               If FECO% >= CVI(Mid$(YY$, 115, 2)) Then
                  Call REPLA(YY$, MKS$(COU), 111, 4)
                  Call REPLA(YY$, MKI$(FECO%), 115, 2)
                  If PROVE% > 0 Then
                     Call REPLA(YY$, MKI$(PROVE%), 123, 2)
                  End If
                  Call GRAREG("MASTER", YY$, REGMA&)
               End If
            End If
         End If
      End If
   Next U&
   Call CIERRARCH("CACORE")
   '
20 FIN& = NUINV%
   J& = 0
   For I& = 1 To FIN&
      Call TRACE(20, I&, FIN&)
      XX$ = REGLEIDO$("INVERT", I& + 1)
      CODX$ = Left$(XX$, 16)
      If Mid$(CODX$, 16, 1) <> Chr$(96) Then                  ' items de baja
         CIIX% = CVI(Mid$(XX$, 17, 2))
         If CIIX% > 0 Then
            If CIIX% <= NUMAS% Then
               REGMA& = CIIX% + 1
               YY$ = REGLEIDO$("MASTER", REGMA&)
               UNI$ = Mid$(YY$, 109, 2)
               PROVE% = CVI(Mid$(YY$, 123, 2))
               COU = CVS(Mid$(YY$, 111, 4))
               MONE$ = "$": If Asc(Mid$(YY$, 101, 1)) = 2 Then MONE$ = "U$"
                            If Asc(Mid$(YY$, 101, 1)) = 3 Then MONE$ = "DM"
               FECO% = CVI(Mid$(YY$, 115, 2))
               '
               MARFAB$ = "  "
               If DP%(CIIX%) > 0 Then MARFAB$ = "FP"
               '
               ZZ$ = REGBLAN$("CACORE")
               Call REPLA(ZZ$, MKI$(CIIX%), 1, 2)
               Call REPLA(ZZ$, UNI$, 3, 2)
               Call REPLA(ZZ$, MONE$, 5, 2)
               Call REPLA(ZZ$, MKS$(COU), 7, 4)
               Call REPLA(ZZ$, MKI$(FECO%), 11, 2)
               Call REPLA(ZZ$, MKI$(PROVE%), 13, 2)
               Call REPLA(ZZ$, MARFAB$, 15, 2)
               '
               J& = J& + 1
               Call GRAREG("!CACORE", ZZ$, J&)
               '
            End If
         End If
      End If
   Next I&
   '
   Call SETULTREG("!CACORE", J&)
   Call CIERRARCH("!CACORE")
   Screen.MousePointer = 1
   '
50 VTB% = VENTABU%("VCACORE")
   If VTB% >= 0 Then
        Call GRACOSREPMAT
      Else
        If COMALTER%("Abandona sin Grabar los Cambios ?\\Guardar Cambios\Abandonar") <> 2 Then GoTo 50
   End If
   '
   Call CIERRARCH("*.*")
   '
End Sub

Sub GRACOSREPMAT()
   '
   HOI% = HOY(HO$)
   '
   Screen.MousePointer = 11
   FIN& = ULTREG&("!CACORE")
   For U& = 1 To FIN&
      Call TRACE(20, U&, FIN&)
      XX$ = REGLEIDO$("!CACORE", U&)
      Call GRAREG("CACORE", XX$, U&)
      CIIX% = CVI(Left$(XX$, 2))
      If CIIX% > 0 Then
         If CIIX% <= NUMAS% Then
            COU = CVS(Mid$(XX$, 7, 4))
            FECO% = CVI(Mid$(XX$, 11, 2))
            PROVE% = CVI(Mid$(XX$, 13, 2))
            If COU > 0 Then
               If FECO% < 1 Then FECO% = HOI%
               REGMA& = CIIX% + 1
               YY$ = REGLEIDO$("MASTER", REGMA&)
               If FECO% >= CVI(Mid$(YY$, 115, 2)) Then
                  Call REPLA(YY$, MKS$(COU), 111, 4)
                  Call REPLA(YY$, MKI$(FECO%), 115, 2)
                  Call REPLA(YY$, MKI$(PROVE%), 123, 2)
                  Call GRAREG("MASTER", YY$, REGMA&)
               End If
            End If
         End If
      End If
   Next U&
   '
   Call CIERRARCH("CACORE")
   Screen.MousePointer = 1
   '
End Sub

Sub CANACOSMOD(CIXX%)
   '
   Static URESEC&, CAMAQ%
   Static COMAQ$(1024), CUHOR(1024)
   Static ENCALI$, LINSEP$
   '
   If ENCALI$ = "" Then
      ENCALI$ = "Codigo           N-O Operacion                            Tie-Unid C.Oper      $/Hora     Ops   Cat   $/Hora      $/Unid     Cant    Tot.Oper    Tot.Comp               "
      LINSEP$ = String$(154, "-")
   End If
   '
   If URESEC& < 1 Then URESEC& = ULTREG&("SECOPER")
   If CAMAQ% < 1 Then
     CAMAQ% = ULTREG&("PADMAQ")
     For UU& = 1 To ULTREG&("PADMAQ")
        R0$ = REGLEIDO$("PADMAQ", UU&)
        COMAQ$(UU&) = TRIM$(Left$(R0$, 6))
        CUHOR(UU&) = CVS(Mid$(R0$, 93, 4))
     Next UU&
     Call CIERRARCH("PADMAQ")
   End If
   '
   Dim CICOMPO%(4096), CANCOMPO(4096)
   CACOM% = 0
   JJ& = 0
   '
   FISAL% = FILEOPEN%("C:\FLEXOFT\ANCOSMOD.TXT", 2, 0)
   '
   Call EXPLOMUL(CIXX%)
   For UI% = 1 To CONTAMUL%
      For UJ% = 1 To CACOM%
         If CICOMPO%(UJ%) = CIK%(UI%) Then
            CANCOMPO(UJ%) = CANCOMPO(UJ%) + CANT(UI%)
            GoTo 9
         End If
      Next UJ%
      CACOM% = CACOM% + 1
      CICOMPO%(CACOM%) = CIK%(UI%)
      CANCOMPO(CACOM%) = CANT(UI%)
9  Next UI%
   '
   SUMATO = 0
   Print #FISAL%, ENCALI$
   Print #FISAL%, LINSEP$
   For UI% = 1 To CACOM%
     RFF$ = RECFILT$("SECOPER", 1, MKI$(CICOMPO%(UI%)))
     For KK% = 1 To Len(RFF$) Step 4
       RESEC& = CVS(Mid$(RFF$, KK%, 4))
       If RESEC& > 0 Then
         If RESEC& <= URESEC& Then
           SCP$ = REGLEIDO$("SECOPER", RESEC&)
           NOPE% = CVI(Mid$(SCP$, 3, 2))
           DOPE$ = Mid$(SCP$, 5, 48)
           MAQUI$ = TRIM$(Mid$(SCP$, 53, 6))
           '
           CUOTMAQ = 0
           For KJ% = 1 To CAMAQ%
              If COMAQ$(KJ%) = MAQUI$ Then
                 CUOTMAQ = CUHOR(KJ%)
              End If
           Next KJ%
           '
           LOTESTA = LOTESTAN(CICOMPO%(UI%))
           If LOTESTA < 1 Then LOTESTA = 1
           CALOTES = 1
           MINPREP = 60 * CVS(Mid$(SCP$, 59, 4)) * CALOTES
           MINFIJO = 60 * CVS(Mid$(SCP$, 63, 4)) * CALOTES
           MINVARI = LOTESTA * CVS(Mid$(SCP$, 67, 4))
           MINDEMO = (MINFIJO + MINVARI) * CVS(Mid$(SCP$, 75, 4)) / 100
           MINLIMP = 60 * CVS(Mid$(SCP$, 79, 4)) * CALOTES
           MINLOTE = MINPREP + MINFIJO + MINVARI + MINDEMO + MINLIMP
           TIETOTLOTE$ = HORMINSE$(MINLOTE * 60) ' en segundos
           TIEUNIPIEZ$ = HORMINSE$(MINLOTE * 60 / LOTESTA) ' en segundos
           CAOPS = CVS(Mid$(SCP$, 83, 4))
           CATOP% = Asc(Mid$(SCP$, 87, 1))
           CUOTOPE = VALHORA(CATOP%)
           CATOPS$ = CSTRING$(MKI$(CATOP%), 1, 6, 0, 2)
           COSMODUNI = (CUOTMAQ + CUOTOPE * CAOPS) * (MINLOTE / 60) / LOTESTA
           COSMOD = COSMODUNI * CANCOMPO(UI%)
           '
           Z$ = Space$(160)
           If KK% = 1 Then
              Call REPLA(Z$, CODEXT$(CICOMPO%(UI%)), 1, 16)
              SUCOSMOD = 0
           End If
           Call REPLA(Z$, CSTRING$(MKI$(NOPE), 1, 4, 0, 2), 17, 4)
           Call REPLA(Z$, DOPE$, 22, 35)
           'Call REPLA(Z$, CSTRING$(MKS$(LOTESTA), 3, 8, 0, 2), 61, 8)
           'Call REPLA(Z$, AJUSTD$(TIETOTLOTE$, 10), 69, 10)
           Call REPLA(Z$, AJUSTD$(TIEUNIPIEZ$, 10), 58, 10)
           Call REPLA(Z$, MAQUI$, 70, 6)
           Call REPLA(Z$, CSTRING$(MKS$(CUOTMAQ), 3, 10, 2, 2), 76, 10)
           Call REPLA(Z$, CSTRING$(MKS$(CAOPS), 3, 8, 2, 2), 86, 8)
           Call REPLA(Z$, AJUSTD$(CATOPS$, 6), 94, 6)
           Call REPLA(Z$, CSTRING$(MKS$(CUOTOPE), 3, 10, 4, 2), 99, 10)
           Call REPLA(Z$, CSTRING$(MKS$(COSMODUNI), 3, 12, 5, 2), 109, 12)
           Call REPLA(Z$, CSTRING$(MKS$(CANCOMPO(UI%)), 3, 9, 3, 2), 121, 9)
           Call REPLA(Z$, CSTRING$(MKS$(COSMOD), 3, 12, 5, 2), 130, 12)
           SUCOSMOD = SUCOSMOD + COSMOD
           SUMATO = SUMATO + COSMOD
           ATXX% = 0
           If KK% + 4 > Len(RFF$) Then
              Call REPLA(Z$, CSTRING$(MKS$(SUCOSMOD), 3, 12, 5, 2), 142, 12)
              SUCOSMOD = 0
              ATXX% = 1
           End If
           Print #FISAL%, Z$
           If ATXX% = 1 Then Print #FISAL%, LINSEP$
         End If
       End If
     Next KK%
   Next UI%
   '
   Z$ = Space$(160)
   Call REPLA(Z$, "S U B T O T A L   Mano de Obra", 1, 35)
   Call REPLA(Z$, CSTRING$(MKS$(SUMATO), 3, 12, 5, 2), 130, 12)
   Print #FISAL%, Z$
   Print #FISAL%, LINSEP$
   Close #FISAL%
   '
   '
'   If EXISTE%("C:\FLEXOFT\WORDPAD.EXE") = 1 Then
'        Shell "C:\FLEXOFT\WORDPAD.EXE " + "C:\FLEXOFT\ANCOSMOD.TXT", vbMaximizedFocus
'      Else
'        MENSA$ = "Imposible Mostrar Archivo Generado."
'        MENSA$ = MENSA$ + "\  \Falta Utilitario 'WORDPAD.EXE'"
'        MENSA$ = MENSA$ + "\en Directorio 'C:/FLEXOFT'."
'        MENSA$ = MENSA$ + "\   \Instale este Programa Utilitario\y Re-Genere el Listado."
'        Call COMUNI(MENSA$)
'   End If
   '
End Sub

Sub CANACOSMAT(CIXX%)
   '
'Exit Sub
   Static ENCALI$, LINSEP$
   '
   Screen.MousePointer = 11
   If ENCALI$ = "" Then
      ENCALI$ = "Codigo          Descripcion                                    Operacion                             $/Unid      Cant     Tot.Mat    Tot.Serv    Tot.Comp"
      LINSEP$ = String$(154, "-")
   End If
   '
   UREANA& = ULTREG&("ANACOS")
   Dim RECOSMAT$(4096)
   CACOM% = 0
   JJ& = 0
   '
   FISAL% = FILEOPEN%("C:\FLEXOFT\ANCOSMAT.TXT", 2, 0)
   RFF$ = RECFILT$("ANACOS", 1, MKI$(CIXX%))
   For UI% = 1 To Len(RFF$) Step 4
     REANA& = CVS(Mid$(RFF$, UI%, 4))
     If REANA& > 0 Then
       If REANA < UREANA& Then
         XX$ = REGLEIDO$("ANACOS", REANA&)
         CODD$ = Mid$(XX$, 3, 16)
         DECO$ = DESCRIT$(CODINT%(CODD$))
         DEOP$ = ""
         VUNI = CVS(Mid$(XX$, 69, 4))
         CANX = CVS(Mid$(XX$, 73, 4))
         CMAT = CVS(Mid$(XX$, 77, 4))
         CSER = CVS(Mid$(XX$, 85, 4))
         If CSER > 0.0005 Then
           DEOP$ = Mid$(XX$, 19, 48)
         End If
       End If
     End If
     '
     TEBUS$ = Space$(107)
     Call REPLA(TEBUS$, CODD$, 1, 16)
     Call REPLA(TEBUS$, DECO$, 17, 47)
     Call REPLA(TEBUS$, DEOP$, 64, 32)
     Call REPLA(TEBUS$, CSTRING$(MKS$(VUNI), 3, 12, 4, 2), 96, 12)
     '
     For UJ% = 1 To CACOM%
       If Left$(RECOSMAT$(UJ%), 107) = TEBUS$ Then
          CANX = CANX + Val(Mid$(RECOSMAT$(UJ%), 108, 10))
          CMAT = CMAT + Val(Mid$(RECOSMAT$(UJ%), 118, 12))
          CSER = CSER + Val(Mid$(RECOSMAT$(UJ%), 130, 12))
          TENCU$ = RECOSMAT$(UJ%)
          Call REPLA(TENCU$, CSTRING$(MKS$(CANX), 3, 10, 3, 2), 108, 10)
          Call REPLA(TENCU$, CSTRING$(MKS$(CMAT), 3, 12, 4, 2), 118, 12)
          Call REPLA(TENCU$, CSTRING$(MKS$(CSER), 3, 12, 4, 2), 130, 12)
          RECOSMAT(UJ%) = TENCU$
          GoTo 9
        End If
     Next UJ%
     CACOM% = CACOM% + 1
     TENCU$ = Space$(154)
     Call REPLA(TENCU$, TEBUS$, 1, 107)
     Call REPLA(TENCU$, CSTRING$(MKS$(CANX), 3, 10, 3, 2), 108, 10)
     Call REPLA(TENCU$, CSTRING$(MKS$(CMAT), 3, 12, 4, 2), 118, 12)
     Call REPLA(TENCU$, CSTRING$(MKS$(CSER), 3, 12, 4, 2), 130, 12)
     RECOSMAT(CACOM%) = TENCU$
9  Next UI%
   '
   For UI% = 1 To CACOM%
     For UJ% = UI% + 1 To CACOM%
       If Left$(RECOSMAT$(UI%), 16) > Left$(RECOSMAT$(UJ%), 16) Then
         TENCU$ = RECOSMAT$(UJ%)
         RECOSMAT$(UJ%) = RECOSMAT$(UI%)
         RECOSMAT$(UI%) = TENCU$
       End If
     Next UJ%
   Next UI%
   '
   SUMAPAR = 0: SUMATO = 0: SUMATER = 0: SUSERVI = 0
   Print #FISAL%, ENCALI$
   Print #FISAL%, LINSEP$
   For UI% = 1 To CACOM%
     Z$ = RECOSMAT$(UI%)
     SUMAPAR = SUMAPAR + Val(Mid$(Z$, 118, 12)) + Val(Mid$(Z$, 130, 12))
     SUMATO = SUMATO + Val(Mid$(Z$, 118, 12)) + Val(Mid$(Z$, 130, 12))
     SUMATER = SUMATER + Val(Mid$(Z$, 118, 12))
     SUSERVI = SUSERVI + Val(Mid$(Z$, 130, 12))
     If UI% = CACOM% Or Left$(RECOSMAT$(UI%), 16) <> Left$(RECOSMAT$(UI% + 1), 16) Then
       Call REPLA(Z$, CSTRING$(MKS$(SUMAPAR), 3, 12, 4, 2), 142, 12)
       SUMAPAR = 0
       ATXX% = 1
     End If
     Print #FISAL%, Z$
     If ATXX% = 1 Then Print #FISAL%, LINSEP$
   Next UI%
   '
   Z$ = Space$(154)
   Call REPLA(Z$, "S U B T O T A L   Materiales y Servicios", 1, 40)
   Call REPLA(Z$, CSTRING$(MKS$(SUMATER), 3, 12, 4, 2), 118, 12)
   Call REPLA(Z$, CSTRING$(MKS$(SUSERVI), 3, 12, 4, 2), 130, 12)
   Call REPLA(Z$, CSTRING$(MKS$(SUMATO), 3, 12, 4, 2), 142, 12)
   Print #FISAL%, Z$
   Print #FISAL%, LINSEP$
   Close #FISAL%
   '
   Screen.MousePointer = 1
   '
End Sub

Sub PRIANCOSMOD(CIXX%)
   '
   TOMOBRA$ = Space$(12): TOMATER$ = Space$(12)
   If EXISTE%("C:\FLEXOFT\ANCOSMOD.TXT") = 1 Then
      FISAL% = FILEOPEN%("C:\FLEXOFT\ANCOSMOD.TXT", 1, 0)
      On Error GoTo 0
      '
      JJ& = 0
      Call HEADERS("ANCOSMOD", "")
      Call HEADERS("ANCOSMOD", "Codigo: " + TRIM$(CODEXT$(CIXX%)) + " - " + DESCRIT$(CIXX%))
      '
      JJ& = JJ& + 1: Call GRAREG("ANCOSMOD", "1.- MANO DE OBRA DIRECTA", JJ&)
      JJ& = JJ& + 1: Call GRAREG("ANCOSMOD", ".   ********************", JJ&)
      JJ& = JJ& + 1: Call GRAREG("ANCOSMOD", " ", JJ&)
      JJ& = JJ& + 1: Call GRAREG("ANCOSMOD", " ", JJ&)
      '
      While Not EOF(FISAL%)
         Line Input #FISAL%, Z$
         If Left$(Z$, 1) = " " Then Call REPLA(Z$, ".", 1, 1)
         If Mid$(Z$, 93, 1) = " " Then Call REPLA(Z$, ".", 93, 1)
         JJ& = JJ& + 1
         Call GRAREG("ANCOSMOD", RTrim$(Z$), JJ&)
         If Left$(Z$, 15) = "S U B T O T A L" Then
            TOMOBX# = Val(Mid$(Z$, 130, 12))
            TOMOBRA$ = CSTRING$(MKD$(TOMOBX#), 3, 12, 4, 2)
         End If
      Wend
      Close #FISAL%
      '
      JJ& = JJ& + 1: Call GRAREG("ANCOSMOD", " ", JJ&)
      JJ& = JJ& + 1: Call GRAREG("ANCOSMOD", " ", JJ&)
      JJ& = JJ& + 1: Call GRAREG("ANCOSMOD", " ", JJ&)
      JJ& = JJ& + 1: Call GRAREG("ANCOSMOD", " ", JJ&)
   End If
   '
   If EXISTE%("C:\FLEXOFT\ANCOSMAT.TXT") = 1 Then
      FISAL% = FILEOPEN%("C:\FLEXOFT\ANCOSMAT.TXT", 1, 0)
      On Error GoTo 0
      '
      JJ& = JJ& + 1: Call GRAREG("ANCOSMOD", "2.- MATERIALES Y SERVICIOS DE TERCEROS", JJ&)
      JJ& = JJ& + 1: Call GRAREG("ANCOSMOD", ".   **********************************", JJ&)
      JJ& = JJ& + 1: Call GRAREG("ANCOSMOD", " ", JJ&)
      JJ& = JJ& + 1: Call GRAREG("ANCOSMOD", " ", JJ&)
      '
      While Not EOF(FISAL%)
         Line Input #FISAL%, Z$
         If Left$(Z$, 1) = " " Then Call REPLA(Z$, ".", 1, 1)
         If Mid$(Z$, 93, 1) = " " Then Call REPLA(Z$, ".", 93, 1)
         JJ& = JJ& + 1
         Call GRAREG("ANCOSMOD", RTrim$(Z$), JJ&)
         If Left$(Z$, 15) = "S U B T O T A L" Then
            TOMATER$ = Mid$(Z$, 142, 12)
         End If
      Wend
      Close #FISAL%
   End If
   '
   JJ& = JJ& + 1: Call GRAREG("ANCOSMOD", " ", JJ&)
   JJ& = JJ& + 1: Call GRAREG("ANCOSMOD", " ", JJ&)
   JJ& = JJ& + 1: Call GRAREG("ANCOSMOD", " ", JJ&)
   JJ& = JJ& + 1: Call GRAREG("ANCOSMOD", " ", JJ&)
   '
   JJ& = JJ& + 1: Call GRAREG("ANCOSMOD", "3.- RESUMEN GENERAL", JJ&)
   JJ& = JJ& + 1: Call GRAREG("ANCOSMOD", ".   ***************", JJ&)
   JJ& = JJ& + 1: Call GRAREG("ANCOSMOD", " ", JJ&)
   '
   TOTACOS# = Val(TOMOBRA$) + Val(TOMATER$)
   TOCOSTO$ = CSTRING$(MKD$(TOTACOS#), 3, 12, 4, 2)
   '
   LINSEP$ = String$(154, "-")
   JJ& = JJ& + 1: Call GRAREG("ANCOSMOD", LINSEP$, JJ&)
   JJ& = JJ& + 1: Call GRAREG("ANCOSMOD", " ", JJ&)
   Z$ = Space$(155)
   FEX = HOY(HOS$)
   Call REPLA(Z$, "T O T A L  G E N E R A L  (" + HOS$ + " - " + Time$ + ")", 1, 50)
   Call REPLA(Z$, "Materiales:", 60, 12)
   Call REPLA(Z$, TRIM$(TOMATER$), 73, 12)
   Call REPLA(Z$, "Fabricacion:", 93, 15)
   Call REPLA(Z$, TRIM$(TOMOBRA$), 108, 12)
   Call REPLA(Z$, "Total:", 136, 12)
   Call REPLA(Z$, TOCOSTO$, 142, 12)
   If Mid$(Z$, 93, 1) = " " Then Call REPLA(Z$, ".", 93, 1)
   JJ& = JJ& + 1: Call GRAREG("ANCOSMOD", Z$, JJ&)
   JJ& = JJ& + 1: Call GRAREG("ANCOSMOD", " ", JJ&)
   JJ& = JJ& + 1: Call GRAREG("ANCOSMOD", LINSEP$, JJ&)
   '
   Call SETULTREG("ANCOSMOD", JJ&)
   Call CIERRARCH("ANCOSMOD")
   Call CONECRUN("*ANCOSMOD", "Impresion de Analisis de Costo M.O.D")
   '
End Sub
