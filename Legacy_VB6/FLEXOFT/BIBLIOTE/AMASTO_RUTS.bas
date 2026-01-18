Attribute VB_Name = "AMASTO_RUTS"
Sub CARTATIMA()
    '
    Call BLANARCH("!TABTIMA")
    For I& = 1 To ULTREG&("TABTIMA")
       X$ = REGLEIDO$("TABTIMA", I&)
       If Asc(Left$(X$, 1)) > 0 Then
           Call REGAPP("!TABTIMA", X$)
       End If
    Next I&
    Call CIERRARCH("!TABTIMA")
    '
10  VTB% = VENTABU%("TABTIMA")
    If VTB% < 0 Then Exit Sub
    J& = 0
    For I& = 1 To ULTREG&("!TABTIMA")
       X$ = REGLEIDO$("!TABTIMA", I&)
       If Asc(Left$(X$, 1)) > 0 Then
          DPP$ = Mid$(X$, 17, 1)
          If ECOARCH$("ECODEP", DPP$) = "" Then
             Call MENSERR(24, "Depósito no Definido")
             GoTo 10
          End If
          If Asc(DPP$) = 99 Then
             Call MENSERR(24, "Depósito 99 no Válido")
             GoTo 10
          End If
          J& = J& + 1
          Call GRAREG("TABTIMA", X$, J&)
          Call GRAREG("TATIMAT", X$, J&)
       End If
    Next I&
    Call SETULTREG("TABTIMA", J&)
    Call CIERRARCH("TABTIMA")
    Call SETULTREG("TATIMAT", J&)
    Call CIERRARCH("TATIMAT")
    Call CIERRARCH("!TABTIMA")
    '
End Sub
'

Sub CARGRUCOVE()
    '
    Call BLANARCH("!GRUCOVEN")
    For I& = 1 To ULTREG&("GRUCOVEN")
        X$ = REGLEIDO$("GRUCOVEN", I&)
        If CVI(Left$(X$, 2)) > 0 Then
            Call REGAPP("!GRUCOVEN", X$)
        End If
    Next I&
    Call CIERRARCH("GRUCOVEN")
    '
    VTB% = VENTABU%("GRUCOVEN/NC/TITUPAN=Grupos Contables")
    If VTB% < 0 Then Exit Sub
    J& = 0
    For I& = 1 To ULTREG&("!GRUCOVEN")
        X$ = REGLEIDO$("!GRUCOVEN", I&)
        If CVI(Left$(X$, 2)) > 0 Then
            J& = J& + 1
            Call GRAREG("GRUCOVEN", X$, J&)
        End If
    Next I&
    Call SETULTREG("GRUCOVEN", J&)
    Call CIERRARCH("GRUCOVEN")
    '
End Sub

Sub CARTACLAS()
   '
10 X$ = REGSEL$("TACRICLA")
   If X$ = "" Then Exit Sub
   '
   CLASI% = Asc(Left$(X$, 1))
   CLASX$ = TRIM$(Mid$(X$, 2, 31))
   If CLASI% < 1 Or CLASI% > 8 Or CLASX$ = "" Then Exit Sub
   '
   CLASX$ = REPLACAR$(CLASX$, "/", "-")
   ALISX$ = TRIM$(Str$(CLASI%))
   While Len(ALISX$) < 3: ALISX$ = "0" + ALISX$: Wend
   ACLASI$ = "CLASI" + ALISX$
   Call COPYSTRU("CLASISTO", ACLASI$)
   '
   Call TRALOCAL(ACLASI$, "CLASISTO")
   VTB% = VENTABU%("CLASISTO/NC/TITUPAN=Clasificación Técnica: " + CLASX$)
   If VTB% >= 0 Then
      For U& = 1 To ULTREG&("!CLASISTO")
         X$ = REGLEIDO$("!CLASISTO", U&)
         Call REPLA(X$, Chr$(U&), 1, 1)
         Call GRAREG(ACLASI$, X$, U&)
      Next U&
      Call SETULTREG(ACLASI$, U&)
      Call CIERRARCH(ACLASI$)
   End If
   Call CIERRARCH("*.*")
   GoTo 10
   '
End Sub

Sub CARCRITEC()
   Call TRALOCAL("TACRICLA", "")
   VTB% = VENTABU%("TACRICLA/NC/TITUPAN=Criterios Clasificación Técnica")
   If VTB% >= 0 Then
      For U& = 1 To ULTREG&("!TACRICLA")
         X$ = REGLEIDO$("!TACRICLA", U&)
         Call REPLA(X$, Chr$(U&), 1, 1)
         Call GRAREG("TACRICLA", X$, U&)
      Next U&
   End If
   If ULTREG&("TACRICLA") > 8 Then
      Call SETULTREG("TACRICLA", 8)
   End If
   Call CIERRARCH("TACRICLA")
End Sub

Sub CLASINTER()
   '
   Screen.MousePointer = 11
   J& = 0: FIN& = NUINV%
   '
   For I& = 1 To FIN&
     Call TRACE(20, I&, FIN&)
     X$ = REGLEIDO$("INVERT", I& + 1)
     If Mid$(X$, 16, 1) <> "`" Then
       CI% = CVI(Mid$(X$, 17, 2))
       If CI% > 0 Then
         If CI% <= NUMAS% Then
           REGMA& = CI% + 1
           M$ = REGLEIDO$("MASTER", REGMA&)
           CLASI$ = Mid$(M$, 69, 8)
           Z$ = Space$(80)
           Call REPLA(Z$, Left$(X$, 15), 1, 15)
           Call REPLA(Z$, Mid$(M$, 5, 64), 17, 32)
           Z1$ = Space$(24)
           For KKI% = 1 To 8
             CQQ% = Asc(Mid$(CLASI$, KKI%, 1))
             If CQQ% > 0 Then
               If CQQ% < 32 Then
                 Z2$ = CSTRING$(Chr$(CQQ%), 1, 2, 0, 2)
                 Call REPLA(Z1$, Z2, 3 * KKI% - 1, 2)
               End If
             End If
           Next KKI%
           Call REPLA(Z$, Z1$, 49, 24)
           J& = J& + 1
           Call GRAREG("CLASISEL", Z$, J&)
         End If
       End If
     End If
   Next I&
   '
   Call SETULTREG("CLASISEL", J&)
   Call CIERRARCH("CLASISEL")
   '
10 FORSELEC.ABUSQUE = TRIM$(CODD1$)
   Call SELECHO("CLASISEL")
   CODD1$ = TRIM$(VALACT1$("CLASISEL"))
   If CODD1$ <> "" Then
     '
     CODD2$ = AJUSTI$(CODD1$, 15)
     CIXI1% = CODINT%(CODD1$)
     If CIXI1% > 0 Then
       If CIXI1% <= NUMAS% Then
         REGMA& = CIXI1% + 1
         M$ = REGLEIDO$("MASTER", REGMA&)
         CLASI$ = Mid$(M$, 69, 8)
         CODESCRI$ = CODD1$ + " - " + DESCRIT0$(CIXI1%)
         Call TRALOCAL("TACRICLA", "")
         For U& = 1 To 8
           VALOX% = Asc(Mid$(CLASI$, U&, 1))
           XX$ = REGLEIDO$("!TACRICLA", U&)
           Call REPLA(XX$, Chr$(VALOX%), 49, 1)
           Call GRAREG("!TACRICLA", XX$, U&)
         Next U&
         '
         CLASIA$ = CLASI$
         VTB% = VENTABU%("CLASITEC/TITUPAN=" + CODESCRI$)
         If VTB% >= 0 Then
           CLASI$ = String$(8, 0)
           For U& = 1 To 8
             XX$ = REGLEIDO$("!TACRICLA", U&)
             VALOX% = Asc(Mid$(XX, 49, 1))
             Mid$(CLASI$, U&, 1) = Chr$(VALOX%)
           Next U&
           If CLASI$ = CLASIA$ Then GoTo 90
           '
           REGMA& = CIXI1% + 1
           M$ = REGLEIDO$("MASTER", REGMA&)
           Call REPLA(M$, CLASI$, 69, 8)
           Call GRAREG("MASTER", M$, REGMA&)
           Call CIERRARCH("MASTER")
           '
           URECLA& = ULTREG&("CLASISEL")
           For U& = 1 To URECLA&
             Z$ = REGLEIDO$("CLASISEL", U&)
             If Left$(Z$, 15) = CODD2$ Then
               Z1$ = Space$(24)
               For KKI% = 1 To 8
                 CQQ% = Asc(Mid$(CLASI$, KKI%, 1))
                 If CQQ% > 0 Then
                   If CQQ% < 32 Then
                     Z2$ = CSTRING$(Chr$(CQQ%), 1, 2, 0, 2)
                     Call REPLA(Z1$, Z2, 3 * KKI% - 1, 2)
                   End If
                 End If
               Next KKI%
               Call REPLA(Z$, Z1$, 49, 24)
               Call GRAREG("CLASISEL", Z$, U&)
               Call CIERRARCH("CLASISEL")
               GoTo 90
             End If
           Next U&
         End If
       End If
     End If
90   GoTo 10
     '
   End If
   '
End Sub
'

Sub CLASIRANGO()
   '
   CLR$ = OBJPLA$("CLASIRANGO", CLR$)
   Call CIERRARCH("*.*")
   If CLR$ = "" Then Exit Sub
   CLASI% = Asc(Left$(CLR$, 1))
   If CLASI% < 1 Or CLASI% > ULTREG&("TACRICLA") Then Exit Sub
   '
   ALIX$ = TRIM$(Str$(CLASI%))
   While Len(ALIX$) < 3: ALIX$ = "0" + ALIX$: Wend
   ACLASI$ = "CLASI" + ALIX$
   LU$ = LUDAT$
   Call CIERRARCH(ACLASI$)
   Call CIERRARCH("CLASISTO")
   If COPYFILE%(LU$ + ACLASI$ + ".DAT", LU$ + "CLASISTO.DAT") <> 1 Then
      Exit Sub
   End If
   '
   CLR$ = OBJPLA$("CLASIRANGO-2", CLR$)
   If CLR$ = "" Then Exit Sub
   VACLA% = Asc(Mid$(CLR$, 6, 1))
   If VACLA% < 1 Or VACLA% > ULTREG&(ACLASI$) Then Exit Sub
   '
   Screen.MousePointer = 11
   cod1$ = CODEXT$(CVI(Mid$(CLR$, 2, 2))): If TRIM$(cod1$) = "" Then Exit Sub
   cod2$ = CODEXT$(CVI(Mid$(CLR$, 4, 2))): If TRIM$(cod2$) = "" Then Exit Sub
   '
   J& = 0: FIN& = NUINV%
   '
   For I& = 1 To FIN&
      Call TRACE(20, I&, FIN&)
      X$ = REGLEIDO$("INVERT", I& + 1)
      If Mid$(X$, 16, 1) <> "`" Then
         CI% = CVI(Mid$(X$, 17, 2))
         If CI% > 0 Then
            If CI% <= NUMAS% Then
               REGMA& = CI% + 1
               If REGMA& > 1 Then
                  MX$ = REGLEIDO$("MASTER", REGMA&)
                  If Mid$(MX$, 77, 16) >= cod1$ Then
                    If Mid$(MX$, 77, 16) <= cod2$ Then
                      Call REPLA(MX$, Chr$(VACLA%), 68 + CLASI%, 1)
                      Call GRAREG("MASTER", MX$, REGMA&)
                    End If
                  End If
               End If
            End If
         End If
      End If
   Next I&
   '
   Call CIERRARCH("MASTER")
   Screen.MousePointer = 1
   '
End Sub

Sub CLASITAB()
   '
   Screen.MousePointer = 11
   J& = 0: FIN& = NUINV%
   '
   For I& = 1 To FIN&
      Call TRACE(20, I&, FIN&)
      X$ = REGLEIDO$("INVERT", I& + 1)
      If Mid$(X$, 16, 1) <> "`" Then
         CI% = CVI(Mid$(X$, 17, 2))
         If CI% > 0 Then
            If CI% <= NUMAS% Then
               REGMA& = CI% + 1
               M$ = REGLEIDO$("MASTER", REGMA&)
               CLASI$ = Mid$(M$, 69, 8)
               J& = J& + 1
               Y$ = MKI$(CI%) + CLASI$
               Call GRAREG("!CLASIMAS", Y$, J&)
            End If
         End If
      End If
   Next I&
   '
   Call SETULTREG("!CLASIMAS", J&)
   Call CIERRARCH("!CLASIMAS")
   '
   ARX$ = ""
   For U% = 1 To ULTREG&("TACRICLA")
      U1$ = TRIM$(Str$(U% + 2))
      U2$ = TRIM$(Str$(U%)): While Len(U2$) < 3: U2$ = "0" + U2$: Wend
      ARX$ = ARX$ + "/ARCHECO(" + U1$ + ")=CLASI" + U2$
   Next U%
   Screen.MousePointer = 1
   '
   VTB% = VENTABU%("CLASIMAS/NC" + ARX$)
   If VTB% < 0 Then
      Call CIERRARCH("*.*")
      Exit Sub
   End If
   '
   Screen.MousePointer = 11
   FIN& = ULTREG&("!CLASIMAS")
   For I& = 1 To FIN&
      Call TRACE(20, I&, FIN&)
      Y$ = REGLEIDO$("!CLASIMAS", I&)
      CI% = CVI(Left$(Y$, 2))
      If CI% > 0 Then
         If CI% <= NUMAS% Then
            CLASI$ = Mid$(Y$, 3, 8)
            REGMA& = CI% + 1
            If REGMA& > 1 Then
               MX$ = REGLEIDO$("MASTER", REGMA&)
               Call REPLA(MX$, CLASI$, 69, 8)
               Call GRAREG("MASTER", MX$, REGMA&)
            End If
         End If
      End If
   Next I&
   '
   Call CIERRARCH("!CLASIMAS")
   Call CIERRARCH("MASTER")
   Screen.MousePointer = 1
   '
End Sub
'

