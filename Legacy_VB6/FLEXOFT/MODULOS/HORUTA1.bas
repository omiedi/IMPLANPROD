Attribute VB_Name = "HORUTA1"
Public CABAS, ULTOP%
Dim REG&(128), REGTX$(128)
'
Sub LEESECOR()
    '
    Screen.MousePointer = 11
    HORUTA.List1.Clear
    ULTOP% = 0
10  RFFX$ = RECFILT$("SECOPER", 1, MKI$(CI%))
    For I1& = 1 To Len(RFFX$) Step 4
       I& = CVS(Mid$(RFFX$, I1&, 4))
       X$ = REGLEIDO$("SECOPER", I&)
       If CVI(Left$(X$ + String$(2, 0), 2)) <> CI% Then
           X$ = REGLEIDO$("SECOPER", 1)
           Call GRAREG("SECOPER", X$, 1)
           Call CIERRARCH("SECOPER")
           GoTo 10
       End If
       '
       NOPY% = CVI(Mid$(X$, 3, 2))
       If NOPY% > ULTOP% Then ULTOP% = NOPY%
       NOPEX$ = AJUSTD$(Str$(NOPY%), 4)
       DOPE$ = Mid$(X$, 5, 48)
       REGO& = I&
       REGOX$ = AJUSTD$(Str$(REGO&), 6)
       SECO$ = NOPEX$ + "     " + DOPE$ + " " + REGOX$
       HORUTA.List1.AddItem SECO$
    Next I1&
    '
    If HORUTA.List1.ListCount > 0 Then
        HORUTA.List1.ListIndex = 0
    End If
    Screen.MousePointer = 1
    '
End Sub
'

Sub GRALOTSTAN()
    '
    CABAS = Val(HORUTA.Text4.Text)
    If CABAS < 1 Then CABAS = 1
    If CABAS > 3000000 Then CABAS = 1
    '
    If CI% > 0 Then
      If CI% <= NUMAS% Then
        LT = CABAS
        If LT > 30000 Then
           LT = Int(CABAS / 100 + 0.5) * (-1)
        End If
        '
        LT1% = LT
        REGMAS& = CI% + 1
        XX$ = REGLEIDO$("MASTER", REGMAS&)
        Call REPLA(XX$, MKI$(LT1%), 107, 2)
        Call GRAREG("MASTER", XX$, REGMAS&)
        '
      End If
    End If
    '
End Sub
'

Sub ACTSECOR()
    '
    CABAS = HORUTA.Text4.Text
    If CABAS < 1 Then CABAS = 1
    If CABAS > 3000000 Then CABAS = 1
    REGI& = Val(Mid$(HORUTA.List1.Text, 59, 6))
    NOPY% = Val(Left$(HORUTA.List1.Text, 4))
    DOPE$ = Mid$(HORUTA.List1.Text, 10, 48)
    If REGI& < 1 Then
        VDEF$ = REGBLAN$("SECOPER")
        Call REPLA(VDEF$, MKI$(CI%), 1, 2)
        Call REPLA(VDEF$, MKI$(NOPY%), 3, 2)
        Call REPLA(VDEF$, DOPE$, 5, 48)
      Else
        VDEF$ = REGLEIDO$("SECOPER", REGI&)
    End If
    '
    Call REPLA(VDEF$, MKS$(CABAS), 91, 4)
    TISTA = CABAS * CVS(Mid$(VDEF$, 67, 4))
    Call REPLA(VDEF$, MKS$(TISTA), 67, 4)
    '
    VALNUE$ = OBJPLA$("PROCFABRI", VDEF$)
    MANDACIE% = 0
    If Len(VALNUE$) > 0 Then
      Screen.MousePointer = 11
      TUNI = CVS(Mid$(VALNUE$, 67, 4)) / CABAS
      Call REPLA(VALNUE$, MKS$(TUNI), 67, 4)
      If REGI& > 0 Then
           VALANT$ = REGLEIDO$("SECOPER", REGI&)
           If VALNUE$ <> VALANT$ Then
              Call GRAREG("SECOPER", VALNUE$, REGI&)
           End If
         Else
           Call REGAPP("SECOPER", VALNUE$)
           MANDACIE% = 1
      End If
      If MANDACIE% = 1 Then
         Call CIERRARCH("SECOPER")
      End If
      Screen.MousePointer = 11
      Call LEESECOR
      Screen.MousePointer = 1
    End If
    '
End Sub
'

Sub GRASECOR()
    '
    CI% = CODINT%(COD$)
    If CI% < 1 Then
       Call MENSERR(24, "Opción Grabar no Válida")
       Exit Sub
    End If
    '
    K1% = 0: K2% = 0
    '
    For INDI& = 0 To HORUTA.List1.ListCount - 1
       HORUTA.List1.ListIndex = INDI&
       REGI& = Val(Mid$(HORUTA.List1.Text, 59, 6))
       K1% = K1% + 1
       If REGI& > 0 And REGI& <= ULTREG&("SECOPER") Then
            REGTX$(K1%) = REGLEIDO$("SECOPER", REGI&)
          Else
            REGTX$(K1%) = REGBLAN$("SECOPER")
            NOPY% = Val(Left$(HORUTA.List1.Text, 4))
            DOPE$ = Mid$(HORUTA.List1.Text, 10, 48)
            Call REPLA(REGTX$(K1%), MKI$(CI%), 1, 2)
            Call REPLA(REGTX$(K1%), MKI$(NOPY%), 3, 2)
            Call REPLA(REGTX$(K1%), DOPE$, 5, 48)
       End If
    Next INDI&
    '
    RFFX$ = RECFILT$("SECOPER", 1, MKI$(CI%))
    For K3% = 1 To K1%
       If 4 * K3% <= Len(RFFX$) Then
           RESE& = CVS(Mid$(RFFX$, 4 * K3% - 3, 4))
         Else
           RESE& = 1 + ULTREG&("SECOPER")
       End If
       VALANT$ = REGLEIDO$("SECOPER", RESE&)
       If REGTX$(K3%) <> VALANT$ Then
          Call GRAREG("SECOPER", REGTX$(K3%), RESE&)
       End If
    Next K3%
    For K3% = 4 * (K1% + 1) - 3 To Len(RFFX$) Step 4
       RESE& = CVS(Mid$(RFFX$, K3%, 4))
       RGX$ = REGBLAN$("SECOPER")
       Call GRAREG("SECOPER", RGX$, RESE&)
    Next K3%
    '
'    FIN& = ULTREG&("SECOPER")
'    For INDI& = 1 To FIN&
'       X$ = REGLEIDO$("SECOPER", INDI&)
'       CIX% = CVI(Left$(X$, 2))
'       If CIX% = CI% Or CIX% < 1 Then
'          K2% = K2% + 1
'          RGX$ = REGBLAN$("SECOPER")
'          If K2% <= K1% Then RGX$ = REGTX$(K2%)
'          Call GRAREG("SECOPER", RGX$, INDI&)
'       End If
'    Next INDI&
    '
'    For K3% = K2% + 1 To K1%
'       Call REGAPP("SECOPER", REGTX$(K3%))
'    Next K3%
    '
    Call CIERRARCH("SECOPER")
    '
    ALTE% = COMALTER%("Puede Ahora Re-Generar las Operaciones Programadas\para las Ordenes de Fabricación en Curso.\\Refrescar Operaciones\No Refrescar")
    If ALTE% = 1 Then Call FRESHOPER(CI%)
    '
End Sub
'

Sub COPYSECU()
    '
100 CICJ$ = OBJPLA$("DEACODIGO", CICJ0$)
    If Len(CICJ$) < 4 Then
        Call CIERRARCH("*.*")
        Exit Sub
    End If
    '
    CODORI% = CVI(Left$(CICJ$, 2))
    CODEST% = CVI(Mid$(CICJ$, 3, 2))
    CICJ0$ = CICJ$
    '
    CODOR$ = TRIM$(CODEXT$(CODORI%))
    CODES$ = TRIM$(CODEXT$(CODEST%))
    '
    If CODOR$ = "" Then
        Call MENSERR(24, "Codigo Origen Inexistente")
        GoTo 100
    End If
    '
    If CODES$ = "" Then
        Call MENSERR(24, "Codigo Destino Inexistente")
        GoTo 100
    End If
    '
    Screen.MousePointer = 11
    PRIREORI& = 0
    PRIREDEST& = 0
    '
    RFFX$ = RECFILT$("SECOPER", 1, MKI$(CODORI%))
    If Len(RFFX$) < 4 Then
        Screen.MousePointer = 1
        Call COMUNI("Copia Imposible.\Código Origen sin Operaciones.")
        GoTo 100
    End If
    '
    RFFY$ = RECFILT$("SECOPER", 1, MKI$(CODEST%))
    If Len(RFFY$) >= 4 Then
        Screen.MousePointer = 1
        OPX% = COMALTER%("ATENCION: Código Destino con Operaciones !!!\\Cancelar\Sobre-escribir")
        If OPX% <> 2 Then GoTo 100
    End If
    '
    Screen.MousePointer = 11
    '
    K1% = 0
    For U& = 1 To Len(RFFX$) Step 4
       RESE& = CVS(Mid$(RFFX$, U&, 4))
       If RESE& > 0 Then
          If RESE& <= ULTREG&("SECOPER") Then
             X$ = REGLEIDO$("SECOPER", RESE&)
             Call REPLA(X$, MKI$(CODEST%), 1, 2)
             K1% = K1% + 1
             REGTX$(K1%) = X$
          End If
       End If
    Next U&
    '
    For K3% = 1 To K1%
       If 4 * K3% <= Len(RFFY$) Then
           RESE& = CVS(Mid$(RFFY$, 4 * K3% - 3, 4))
         Else
           RESE& = 1 + ULTREG&("SECOPER")
       End If
       VALANT$ = REGLEIDO$("SECOPER", RESE&)
       If REGTX$(K3%) <> VALANT$ Then
          Call GRAREG("SECOPER", REGTX$(K3%), RESE&)
       End If
    Next K3%
    For K3% = 4 * (K1% + 1) - 3 To Len(RFFY$) Step 4
       RESE& = CVS(Mid$(RFFY$, K3%, 4))
       RGX$ = REGBLAN$("SECOPER")
       Call GRAREG("SECOPER", RGX$, RESE&)
    Next K3%
    '
    Call CIERRARCH("SECOPER")
    '
    CICJ0$ = MKI$(CODORI%) + MKI$(0)
    Screen.MousePointer = 1
    GoTo 100
    '
End Sub
'

Sub LISECOP()
    '
    XX$ = Left$(REGLEIDO$("ECOMAST", 1), 16): CI1% = Abs(CODINT%(XX$))
    XX$ = Left$(REGLEIDO$("ECOMAST", ULTREG&("ECOMAST")), 16): CI2% = Abs(CODINT%(XX$))
    VDEF$ = MKI$(CI1%) + MKI$(CI2%)
    X$ = OBJPLA$("DESHASCOD", MKI$(CI1%) + MKI$(CI2%))
    If Len(X$) < 4 Then
        Exit Sub
    End If
    '
    Screen.MousePointer = 11
    Dim PUNCOD$(32767)
    '
    CI1% = CVI(Left$(X$, 2))
    cod1$ = CODEXT$(CI1%): If TRIM$(cod1$) = "" Then Exit Sub
    CI2% = CVI(Mid$(X$, 3, 2))
    cod2$ = CODEXT$(CI2%): If TRIM$(cod2$) = "" Then Exit Sub
    '
    V& = 0
    FIN& = ULTREG&("SECOPER")
    For U& = 1 To FIN&
        Call TRACE(20, U&, FIN&)
        X$ = REGLEIDO$("SECOPER", U&)
        CI0% = CVI(Left$(X$, 2))
        If CI0% > 0 Then
           If CI0% <= NUMAS% Then
              'CODEX$ = CODEXT$(CI0%)
              'If CODEX$ >= cod1$ Then
              '   If CODEX$ <= cod2$ Then
                    PUNCOD$(CI0%) = PUNCOD$(CI0%) + MKS$(U&)
              '   End If
              'End If
           End If
        End If
    Next U&
    '
    FIN& = NUINV%
    For IU& = 1 To NUINV%
        Call TRACE(20, IU&, FIN&)
        Z$ = REGLEIDO$("INVERT", IU& + 1)
        CODEX$ = Left$(Z$, 16)
        If Mid$(Z$, 16, 1) <> "`" Then
            CODEX$ = TRIM$(CODEX$)
            If CODEX$ >= cod1$ Then
                If CODEX$ <= cod2$ Then
                    CI0% = CVI(Mid$(Z$, 17, 2))
                    If Len(PUNCOD$(CI0%)) >= 4 Then
                       '
                       LOSTAND = LOTESTAN(CI0%)
                       DESCRI$ = TRIM$("Lt.Std." + TRIM$(Str$(LOSTAND)) + " " + TRIM$(UNIMED$(CI0%)))
                       DESCRI$ = " (" + DESCRI$ + ")"
                       LOMADE% = 60 - Len(DESCRI$)
                       If LOMADE% > 0 Then
                           DESCRI$ = Left$(TRIM$(DESCRIT0$(CI0%)), LOMADE%) + DESCRI$
                       End If
                       '
                       Y$ = REGBLAN$("SECOPSEL")
                       Call REPLA(Y$, MKI$(CI0%), 1, 2)
                       Call REPLA(Y$, DESCRI$, 3, 60)
                       V& = V& + 1
                       Call GRAREG("SECOPSEL", Y$, V&)
                       '
                       Call REPLA(Y$, Space$(60), 3, 60)
                       V& = V& + 1
                       Call GRAREG("SECOPSEL", Y$, V&)
                       '
                       For IV% = 1 To Len(PUNCOD$(CI0%)) Step 4
                          U& = CVS(Mid$(PUNCOD$(CI0%), IV%, 4))
                          If U& > 0 Then
                             If U& <= ULTREG&("SECOPER") Then
                                Y1$ = REGLEIDO$("SECOPER", U&)
                                NOPES$ = TRIM$(CSTRING$(Mid$(Y1$, 3, 2), 1, 6, 0, 2))
                                While Len(NOPES$) < 4: NOPES$ = "." + NOPES$: Wend
                                DOPES$ = Mid$(Y1$, 5, 62)
                                TVAR = CVS(Mid$(Y1$, 67, 4)) * LOSTAND
                                Y$ = Y1$
                                Call REPLA(Y$, NOPES$, 3, 6)
                                Call REPLA(Y$, DOPES$, 9, 62)
                                Call REPLA(Y$, MKS$(TVAR), 71, 4)
                                V& = V& + 1
                                Call GRAREG("SECOPSEL", Y$, V&)
                             End If
                          End If
                       Next IV%
                       '
                       Call REPLA(Y$, String$(126, 0), 3, 126)
                       V& = V& + 1
                       Call GRAREG("SECOPSEL", Y$, V&)
                       '
                       Y$ = REGBLAN$("SECOPSEL")
                       V& = V& + 1
                       Call GRAREG("SECOPSEL", Y$, V&)
                       '
                   End If
                End If
            End If
        End If
    Next IU&
    '
    Call SETULTREG("SECOPSEL", V&)
    Call CIERRARCH("SECOPER")
    Call CIERRARCH("SECOPSEL")
    Screen.MousePointer = 1
    Call FINAL("*LISECOP")
    '
End Sub
'

Sub FRESHOPER(CIJK%)
   '
   Dim ORFA$(1024), STATO%(1024), CALOTE#(1024), DESTFAB$(1024)
   CAOF% = 0
   '
Call MENSERR(24, "Imposible Refrescar Operaciones Programadas.")
Exit Sub
   Screen.MousePointer = 11
   For U& = 1 To ULTREG&("OPERFAB")
      YY$ = REGLEIDO$("OPERFAB", U&)
      If CVI(Mid$(YY$, 13, 2)) = CIJK% Then
         NORFA$ = Left$(YY$, 12)
         For V% = CAOF% To 1 Step -1
            If ORFA$(V%) = NORFA$ Then GoTo 15
         Next V%
         CAOF% = CAOF% + 1
         V% = CAOF%
         ORFA$(V%) = NORFA$
         STATO%(V%) = 1
         CALOTE#(V%) = CVD(Mid$(YY$, 79, 8))
         DESTFAB$(V%) = Mid$(YY$, 65, 8)
         '
15       If CVD(Mid$(YY$, 87, 8)) >= 0.05 Then ' realizado
            STATO%(V%) = 2
         End If
      End If
   Next U&
   '
   JJ& = 0
   For U& = 1 To ULTREG&("OPERFAB")
      YY$ = REGLEIDO$("OPERFAB", U&)
      If CVI(Mid$(YY$, 13, 2)) = CIJK% Then
         NORFA$ = Left$(YY$, 12)
         For V% = 1 To CAOF%
            If ORFA$(V%) = NORFA$ Then
               If STATO%(V%) = 2 Then
                    GoTo 25
                  Else
                    GoTo 29
               End If
            End If
         Next V%
      End If
      '
25    JJ& = JJ& + 1
      Call GRAREG("OPERFAB", YY$, JJ&)
29 Next U&
   '
   Call SETULTREG("OPERFAB", JJ&)
   '
   For II% = 1 To CAOF%
      If STATO%(II%) <> 2 Then
         NORFA$ = ORFA$(II%)
         CAPENDE# = CALOTE#(II%)
         DESTI$ = DESTFAB$(II%)
         '
         SECUOPSTA$ = RECFILT$("SECOPER", 1, MKI$(CIJK%))
         If Len(SECUOPSTA$) > 0 Then
           For TTI% = 1 To Len(SECUOPSTA$) Step 4
             TTL& = CVS(Mid$(SECUOPSTA$, TTI%, 4))
             SCP$ = REGLEIDO$("SECOPER", TTL&)
             If CVI(Left$(SCP$, 2)) = CIJK% Then
                NO% = CVI(Mid$(SCP$, 3, 2))
                DOPE$ = Mid$(SCP$, 4, 48)
                EQ1$ = Mid$(SCP$, 53, 6)
                EQAUX$ = Mid$(SCP$, 103, 12)
                CALOTES = 1
                MINPREP = 60 * CVS(Mid$(SCP$, 59, 4)) * CALOTES
                MINFIJO = 60 * CVS(Mid$(SCP$, 63, 4)) * CALOTES
                MINVARI = CAPENDE# * CVS(Mid$(SCP$, 67, 4))
                MINDEMO = (MINFIJO + MINVARI) * CVS(Mid$(SCP$, 75, 4)) / 100
                MINLIMP = 60 * CVS(Mid$(SCP$, 79, 4)) * CALOTES
                MINLOTE = MINPREP + MINFIJO + MINVARI + MINDEMO + MINLIMP
                '
                HORAPROD = MINLOTE / 60
                CAOPS = CVS(Mid$(SCP$, 83, 4))
                '
                ZZ$ = REGBLAN$("OPERFAB")
                Call REPLA(ZZ$, NORFA$, 1, 12)
                Call REPLA(ZZ$, MKI$(CIJK%), 13, 2)
                Call REPLA(ZZ$, MKI$(NO%), 15, 2)
                Call REPLA(ZZ$, DOPE$, 17, 48)
                Call REPLA(ZZ$, DESTI$, 65, 8)
                Call REPLA(ZZ$, EQ1$, 73, 6)
                Call REPLA(ZZ$, MKD$(CAPENDE#), 79, 8)
                Call REPLA(ZZ$, MKD$(CAPENDE#), 95, 8)
                Call REPLA(ZZ$, MKS$(MINPREP), 103, 4)
                Call REPLA(ZZ$, MKS$(MINFIJO), 107, 4)
                Call REPLA(ZZ$, MKS$(MINVARI), 111, 4)
                Call REPLA(ZZ$, MKS$(MINDEMO), 115, 4)
                Call REPLA(ZZ$, MKS$(MINLIMP), 123, 4)
                Call REPLA(ZZ$, MKS$(MINLOTE), 131, 4)
                Call REPLA(ZZ$, String$(16, 0), 135, 16)
                Call REPLA(ZZ$, Space$(48), 151, 48)
                Call REPLA(ZZ$, MKS$(TTL&), 199, 4)
                Call REPLA(ZZ$, EQAUX$, 203, 18)
                Call REPLA(ZZ$, MKS$(CAOPS), 221, 4)
                Call REPLA(ZZ$, String$(32, 0), 225, 32)
                '
                Call REGAPP("OPERFAB", ZZ$)
             End If
           Next TTI%
         End If
         '
         SECUOPSTB$ = RECFILT$("BCOMPLEM", 1, MKI$(CIJK%))
         If Len(SECUOPSTB$) > 0 Then
           For TTI% = 1 To Len(SECUOPSTB$) Step 4
             TTL& = CVS(Mid$(SECUOPSTB$, TTI%, 4))
             SCP$ = REGLEIDO$("BCOMPLEM", TTL&)
             If CVI(Left$(SCP$, 2)) = CIJK% Then
                NO% = 999
                DOPE$ = Mid$(SCP$, 3, 14)
                EQ1$ = Mid$(SCP$, 33, 6)
                EQAUX$ = String$(18, 0)
                LOTESTA = LOSTA: If LOTESTA < 1 Then LOTESTA = 1
                CALOTES = 1
                MINPREP = CVS(Mid$(SCP$, 65, 4))
                MINFIJO = 0
                MINVARI = CVS(Mid$(SCP$, 73, 4))
                MINDEMO = 0
                MINLIMP = 0
                MINLOTE = MINPREP + MINFIJO + MINVARI + MINDEMO + MINLIMP
                '
                If MINLOTE > 0.1 Then
                   HORAPROD = MINLOTE / 60
                   CAOPS = 1
                   '
                   ZZ$ = REGBLAN$("OPERFAB")
                   Call REPLA(ZZ$, NORFA$, 1, 12)
                   Call REPLA(ZZ$, MKI$(CIJK%), 13, 2)
                   Call REPLA(ZZ$, MKI$(NO%), 15, 2)
                   Call REPLA(ZZ$, DOPE$, 17, 48)
                   Call REPLA(ZZ$, DESTI$, 65, 8)
                   Call REPLA(ZZ$, EQ1$, 73, 6)
                   Call REPLA(ZZ$, MKD$(CAPENDE#), 79, 8)
                   Call REPLA(ZZ$, MKD$(CAPENDE#), 95, 8)
                   Call REPLA(ZZ$, MKS$(MINPREP), 103, 4)
                   Call REPLA(ZZ$, MKS$(MINFIJO), 107, 4)
                   Call REPLA(ZZ$, MKS$(MINVARI), 111, 4)
                   Call REPLA(ZZ$, MKS$(MINDEMO), 115, 4)
                   Call REPLA(ZZ$, MKS$(MINLIMP), 123, 4)
                   Call REPLA(ZZ$, MKS$(MINLOTE), 131, 4)
                   Call REPLA(ZZ$, String$(16, 0), 135, 16)
                   Call REPLA(ZZ$, Space$(48), 151, 48)
                   Call REPLA(ZZ$, MKS$(0), 199, 4)
                   Call REPLA(ZZ$, EQAUX$, 203, 18)
                   Call REPLA(ZZ$, MKS$(CAOPS), 221, 4)
                   Call REPLA(ZZ$, String$(32, 0), 225, 32)
                   '
                   Call REGAPP("OPERFAB", ZZ$)
                End If
             End If
           Next TTI%
         End If
         '
      End If
   Next II%
   Call CIERRARCH("OPERFAB")
   '
   Screen.MousePointer = 1
   '
End Sub

