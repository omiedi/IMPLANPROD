Attribute VB_Name = "PROCINGE"
'
Sub NUEINSTRU()
   '
   HOII% = HOY(HO$)
   VDEF$ = REGBLAN$("INSTRUCI")
   Call REPLA(VDEF$, USERNAME$, 71, 24)
   Call REPLA(VDEF$, MKI$(HOII%), 95, 2)
   INDINS% = 1 + ULTREG&("INSTRUCI")
   Call REPLA(VDEF$, MKI$(INDINS%), 69, 2)
   '
10 Screen.MousePointer = 1
   OBX$ = OBJPLA$("NUEINSTRUC", VDEF$)
   If OBX$ = "" Then
     Exit Sub
   End If
   '
   VDEF$ = OBX$
   CODINS$ = UCase$(TRIM$(Left$(OBX$, 16)))
   If CODINS$ = "" Then
      Call MENSERR(24, "Falta Código de Instrucción")
      GoTo 10
   End If
   Call REPLA(OBX$, CODINS$, 1, 16)
   '
   Screen.MousePointer = 11
   FIN& = ULTREG&("INSTRUCI")
   For UI& = 1 To FIN&
      XX$ = REGLEIDO$("INSTRUCI", UI&)
      If UCase$(TRIM$(Left$(XX$, 16))) = CODINS$ Then
         Call MENSERR(24, "Código Pre-Existente")
         GoTo 10
      End If
   Next UI&
   '
   DESINS$ = TRIM$(Mid$(OBX$, 17, 48))
   If DESINS$ = "" Then
      Call MENSERR(24, "Falta Descripción de la Instrucción")
      GoTo 10
   End If
   '
   REVINS$ = TRIM$(Mid$(OBX$, 65, 4))
   If REVINS$ = "" Then
      Call MENSERR(24, "Falta Revisión de la Instrucción")
      GoTo 10
   End If
   '
   PREPINS$ = TRIM$(Mid$(OBX$, 71, 24))
   If PREPINS$ = "" Then
      Call MENSERR(24, "Falta 'Preparó' de la Instrucción")
      GoTo 10
   End If
   '
   FECHINS% = CVI(Mid$(OBX$, 95, 2))
   If FECHINS% < 1 Then
      Call MENSERR(24, "Falta Fecha de Preparación\de la Instrucción")
      GoTo 10
   End If
   '
   Call CIERRARCH("INSTRUCI")
   Call BLOQARCH("INSTRUCI")
   INDINS% = 1 + ULTREG&("INSTRUCI")
   Call REPLA(OBX$, MKI$(INDINS%), 69, 2)
   Call REGAPP("INSTRUCI", OBX$)
   Call CIERRARCH("INSTRUCI")
   Call BAJALDISCO
   Screen.MousePointer = 1
   '
   INSTRUCCI.Label28 = TRIM$(Str$(INDINS%))
   On Error Resume Next
   INSTRUCCI.Show 1
   On Error GoTo 0
   '
End Sub

