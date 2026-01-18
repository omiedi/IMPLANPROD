Attribute VB_Name = "RUTAMBEST"
Sub PRIPLANO(CIXI%)
   '
   Static URE&, RUTARGRA$
   '
   Screen.MousePointer = 11
   '
   If URE& = 0 Then
     URE& = ULTREG&("INDIPLAN")
     If URE& < 1 Then
       URE& = (-1)
     End If
     '
     RUTARGRA$ = TRIM$(CONTROL$("", "RUTARGRA"))
     If RUTARGRA$ = "" Then RUTARGRA$ = TRIM$(LUDAT$)
     If Right$(RUTARGRA$, 1) <> "\" Then
       RUTARGRA$ = RUTARGRA$ + "\"
     End If
   End If
   '
   PRIFORGEN$ = Mid$(TRIM$(CONTROL$("FORGRAFI", "DOCUGRA0")), 2)
   CAIMINI% = 0
   '
   If PRIFORGEN$ <> "" Then
     If LEERPLANO%(CIXI%, ARCHIPLANO$) = 1 Then
       HIMAGEN = FORMALTER.Image3.Picture.Width
       VIMAGEN = FORMALTER.Image3.Picture.Height
       If HIMAGEN > VIMAGEN Then
            PRIFORM$ = "H" + PRIFORGEN$
          Else
            PRIFORM$ = "V" + PRIFORGEN$
       End If
       '
       If EXISTE%(LUDAT$ + PRIFORM$ + ".FRM") Then
         If CIXI% > 0 Then
           If CIXI% <= NUMAS% Then
             '
             CAPARDOC% = 2
             CODPARAM$(1) = "COD-MAT": DOCPARAM$(1) = CODEXT$(CIXI%)
             CODPARAM$(2) = "DES-MAT": DOCPARAM$(2) = TRIM$(DESCRIT0$(CIXI%))
             '
             CACOLTAB1% = 0
             CAITAB1% = 0
             CACOLTAB2% = 0
             CAITAB2% = 0
             '
             NUMEDOC$ = ""
             DESTIDOC% = 0
             TIPODOC$ = "Plano " + TRIM$(CODEXT$(CIXI%))
             COPLANO$ = ARCHIPLANO$
             If COPLANO$ <> "" Then
               If EXISTE%(COPLANO$) > 0 Then
                 CAIMINI% = 1
                 CODIMAGEN$(1) = "PLANOPIE"
                 ARCHIMAGEN$(1) = COPLANO$
                 Call PRINTDOC("@" + PRIFORM$)
               End If
             End If
           End If
         End If
       End If
     End If
   End If
   Screen.MousePointer = 1
   '
End Sub

Sub PRIHOJPROFAB(CIXI%, NOPE%)
   '
   Static URE&, RUTARGRA$
   '
   If URE& = 0 Then
     URE& = ULTREG&("INDIPLAN")
     If URE& < 1 Then
       URE& = (-1)
     End If
     '
     RUTARGRA$ = TRIM$(CONTROL$("", "RUTARGRA"))
     If RUTARGRA$ = "" Then RUTARGRA$ = TRIM$(LUDAT$)
     If Right$(RUTARGRA$, 1) <> "\" Then
       RUTARGRA$ = RUTARGRA$ + "\"
     End If
   End If
   '
   PRIFORM$ = TRIM$(CONTROL$("FORGRAFI", "DOCUGRA2"))
   CAIMINI% = 0
   '
   If PRIFORM$ <> "" Then
     If EXISTE%(LUDAT$ + PRIFORM$ + ".FRM") Then
       If CIXI% > 0 Then
         If CIXI% <= NUMAS% Then
           Screen.MousePointer = 11
           CODPARAM$(1) = "NOP-OPER"
           CODPARAM$(2) = "FECH-REC"
           CODPARAM$(3) = "DES-OPER"
           CODPARAM$(4) = "COD-MAT"
           CODPARAM$(5) = "DES-MAT"
           CODPARAM$(6) = "NUOPERS"
           CODPARAM$(7) = "N-HPROF"
           CODPARAM$(8) = "FECH-LAN"
           CAPARDOC% = 8
           '
           DOCPARAM$(1) = TRIM$(Str$(NOPE%))
           DOCPARAM$(3) = DESCRIOP$(CIXI%, NOPE%)       ' DESCRIPCION CORTA OPERACION
           DOCPARAM$(4) = CODEXT$(CIXI%)
           DOCPARAM$(5) = DESCRIT0$(CIXI%)
           '
           CODTABU1$(1) = "COD-MPRI"
           CODTABU1$(2) = "DES-MPRI"
           CODTABU1$(3) = "CAN-MPRI"
           CODTABU1$(4) = "TIP-MPRI"
           CACOLTAB1% = 4
           CAITAB1% = 0
           '
           CODTABU2$(1) = "EQUIPO1"
           CODTABU2$(2) = "EQUIPO2"
           CODTABU2$(3) = "EQUIPO3"
           CODTABU2$(4) = "TIE-PREP"
           CODTABU2$(5) = "TIE-PROD"
           CACOLTAB2% = 5
           CAITAB2% = 0
           '
           RNC$ = RECFILT$("DESOPFAB", 1, MKI$(CIXI%))
           URNDES& = ULTREG&("DESOPFAB")
           '
           For U& = 1 To Len(RNC$) Step 4
             RENOTA& = CVS(Mid$(RNC$, U&, 4))
             If RENOTA& > 0 Then
               If RENOTA& <= URNDES& Then
                  TTXX$ = REGLEIDO$("DESOPFAB", RENOTA&)
                  If CVI(Mid$(TTXX$, 3, 2)) = NOPE% Then
                     TTXY$ = Mid$(TTXX$, 5, 72)
                     TTYY$ = REPLACAR$(TTXY$, Chr$(9), "     ")
                     CAITAB1% = CAITAB1% + 1
                     TETABU1$(1, CAITAB1%) = Left$(TTYY$, 5)
                     TETABU1$(2, CAITAB1%) = Mid$(TTYY$, 6, 48)
                     TETABU1$(3, CAITAB1%) = Mid$(TTYY$, 54, 9)
                     TETABU1$(4, CAITAB1%) = Mid$(TTYY$, 63, 9)
                  End If
               End If
             End If
           Next U&
           '
           RNC$ = RECFILT$("DATMECAN", 1, MKI$(CIXI%))
           URNDES& = ULTREG&("DATMECAN")
           '
           For U& = 1 To Len(RNC$) Step 4
             RENOTA& = CVS(Mid$(RNC$, U&, 4))
             If RENOTA& > 0 Then
               If RENOTA& <= URNDES& Then
                  TTXX$ = REGLEIDO$("DATMECAN", RENOTA&)
                  If CVI(Mid$(TTXX$, 3, 2)) = NOPE% Then
                     TTYY$ = Mid$(TTXX$, 5, 50)
                     CAITAB2% = CAITAB2% + 1
                     TETABU2$(1, CAITAB2%) = Left$(TTYY$, 4)
                     TETABU2$(2, CAITAB2%) = Mid$(TTYY$, 6, 8)
                     TETABU2$(3, CAITAB2%) = Mid$(TTYY$, 14, 7)
                     TETABU2$(4, CAITAB2%) = Mid$(TTYY$, 22, 4)
                     TETABU2$(5, CAITAB2%) = Mid$(TTYY$, 28, 22)
                  End If
               End If
             End If
           Next U&
           '
           DESTIDOC% = 0
           '
           RNC$ = RECFILT$("DATADOPS", 1, MKI$(CIXI%))
           URN& = ULTREG&("DATADOPS")
           '
           For U& = 1 To Len(RNC$) Step 4
             RENOTA& = CVS(Mid$(RNC$, U&, 4))
             If RENOTA& > 0 Then
               If RENOTA& <= URN& Then
                 TTXX$ = REGLEIDO$("DATADOPS", RENOTA&)
                 If CVI(Mid$(TTXX$, 3, 2)) = NOPE% Then
                   GoTo 29
                 End If
               End If
             End If
           Next U&
           TTXX$ = REGBLAN$("DATADOPS")
           '
29         DOCPARAM$(7) = TRIM$(Mid$(TTXX$, 5, 16))
           DOCPARAM$(2) = TRIM$(Mid$(TTXX$, 21, 4))
           DOCPARAM$(6) = TRIM$(Mid$(TTXX$, 29, 8))
           FEX = CVI(Mid$(TTXX$, 25, 2)): If FEX < 1 Then FEX = HOY(HOS$)
           DOCPARAM$(8) = FECHATEX$(FEX)
           ACROQUI$ = TRIM$(Mid$(TTXX$, 53, 48))
           If ACROQUI$ <> "" Then
             COPLANO$ = RUTARGRA$ + ACROQUI$
             If COPLANO$ <> "" Then
               If EXISTE%(COPLANO$) > 0 Then
                 CAIMINI% = 1
                 CODIMAGEN$(1) = "CROQLAUT"
                 ARCHIMAGEN$(1) = COPLANO$
               End If
             End If
           End If
           DESTIDOC% = 0
           TIPODOC$ = "Hoja de Procesos " + TRIM$(CODEXT$(CIXI%)) + " - Op." + Str$(NOPE%)
           Call PRINTDOC("@" + PRIFORM$)
         End If
       End If
     End If
   End If


End Sub

Sub PRIHOJRINSP(CIXI%)
   '
   Static URE&, RUTARGRA$
   '
   If URE& = 0 Then
     URE& = ULTREG&("INDIPLAN")
     If URE& < 1 Then
       URE& = (-1)
     End If
     '
     RUTARGRA$ = TRIM$(CONTROL$("", "RUTARGRA"))
     If RUTARGRA$ = "" Then RUTARGRA$ = TRIM$(LUDAT$)
     If Right$(RUTARGRA$, 1) <> "\" Then
       RUTARGRA$ = RUTARGRA$ + "\"
     End If
   End If
   '
   PRIFORM$ = TRIM$(CONTROL$("FORGRAFI", "DOCUGRA4"))
   CAIMINI% = 0
   '
   If PRIFORM$ <> "" Then
     If EXISTE%(LUDAT$ + PRIFORM$ + ".FRM") Then
       If CIXI% > 0 Then
         If CIXI% <= NUMAS% Then
           Screen.MousePointer = 11
           CODPARAM$(1) = "COD-MAT"
           CODPARAM$(2) = "DES-MAT"
           CODPARAM$(3) = "REF-MAT1"   'HRI
           CODPARAM$(4) = "REF-MAT2"   'PIE
           CODPARAM$(5) = "N-PLANO"
           CODPARAM$(6) = "MEDIDA-1"   'HRI
           CODPARAM$(7) = "MEDIDA-2"   'PIE
           CODPARAM$(8) = "MEDIDA-3"   'HRI
           CAPARDOC% = 8
           '
           DOCPARAM$(1) = CODEXT$(CIXI%)
           DOCPARAM$(2) = DESCRIT0$(CIXI%)
           DOCPARAM$(3) = NUHOREI$(CIXI%)
           DOCPARAM$(4) = NUPLIEN$(CIXI%)
           DOCPARAM$(5) = NUPLANO$(CIXI%)
           DOCPARAM$(6) = RVHOREI$(CIXI%)
           DOCPARAM$(7) = RVPLIEN$(CIXI%)
           DOCPARAM$(8) = RVPLANO$(CIXI%)
           '
           CODTABU1$(1) = "COD-MPRI"
           CODTABU1$(2) = "DES-MPRI"
           CODTABU1$(3) = "CAN-MPRI"
           CODTABU1$(4) = "TIP-MPRI"
           CODTABU1$(5) = "COL-MPRI"
           CACOLTAB1% = 5
           CAITAB1% = 0
           '
           CACOLTAB2% = 0
           CAITAB2% = 0
           '
           RNC$ = RECFILT$("DESOPINS", 1, MKI$(CIXI%))
           URNDES& = ULTREG&("DESOPINS")
           '
           For U& = 1 To Len(RNC$) Step 4
             RENOTA& = CVS(Mid$(RNC$, U&, 4))
             If RENOTA& > 0 Then
               If RENOTA& <= URNDES& Then
                  TTXX$ = REGLEIDO$("DESOPINS", RENOTA&)
                  TTXY$ = Mid$(TTXX$, 5, 72)
                  TTYY$ = REPLACAR$(TTXY$, Chr$(9), "     ")
                  CAITAB1% = CAITAB1% + 1
                  TETABU1$(1, CAITAB1%) = Left$(TTYY$, 5)
                  TETABU1$(2, CAITAB1%) = Mid$(TTYY$, 6, 41)
                  TETABU1$(3, CAITAB1%) = Mid$(TTYY$, 47, 11)
                  TETABU1$(4, CAITAB1%) = Mid$(TTYY$, 58, 10)
                  TETABU1$(5, CAITAB1%) = Mid$(TTYY$, 68, 5)
               End If
             End If
           Next U&
           '
           DESTIDOC% = 0
           TIPODOC$ = "Hoja de Registro de Inspeccion"
           NUMEDOC$ = TRIM$(DOCPARAM$(3))
           Call PRINTDOC("@" + PRIFORM$)
         End If
       End If
     End If
   End If
   '
End Sub

Sub PRIHOJESP(CIXI%)
   '
   Static URE&, RUTARGRA$
   Dim TENCABE$(3)
   '
   If URE& = 0 Then
     URE& = ULTREG&("INDIPLAN")
     If URE& < 1 Then
       URE& = (-1)
     End If
     '
     RUTARGRA$ = TRIM$(CONTROL$("", "RUTARGRA"))
     If RUTARGRA$ = "" Then RUTARGRA$ = TRIM$(LUDAT$)
     If Right$(RUTARGRA$, 1) <> "\" Then
       RUTARGRA$ = RUTARGRA$ + "\"
     End If
   End If
   '
   PRIFORM$ = TRIM$(CONTROL$("FORGRAFI", "DOCUGRA1"))
   CAIMINI% = 0
   '
   If PRIFORM$ <> "" Then
     If EXISTE%(LUDAT$ + PRIFORM$ + ".FRM") Then
       If CIXI% > 0 Then
         If CIXI% <= NUMAS% Then
           Screen.MousePointer = 11
           CODPARAM$(1) = "N-ORIGIN"
           CODPARAM$(2) = "FECH-LAN"
           CODPARAM$(3) = "COD-MAT"
           CODPARAM$(4) = "N-PLANO"
           CODPARAM$(5) = "FECH-VEN"   'PIE
           CODPARAM$(6) = "REF-MAT1"
           CODPARAM$(7) = "MEDIDA-1"   'HRI
           CODPARAM$(8) = "REF-MAT2"
           CODPARAM$(9) = "MEDIDA-2"   'PIE
           CODPARAM$(10) = "REF-MAT3"
           CODPARAM$(11) = "MEDIDA-3"   'HRI
           CAPARDOC% = 11
           '
           DOCPARAM$(1) = NUHOJES$(CIXI%)
           DOCPARAM$(2) = RVHOJES$(CIXI%)
           DOCPARAM$(3) = CODEXT$(CIXI%)
           DOCPARAM$(4) = NUPLANO$(CIXI%)
           DOCPARAM$(5) = RVPLANO$(CIXI%)
           '
           CODTABU1$(1) = "DES-MAT"
           CACOLTAB1% = 1
           CAITAB1% = 0
           '
           CACOLTAB2% = 0
           CAITAB2% = 0
           '
           RNC$ = RECFILT$("ENCASPEC", 1, MKI$(CIXI%))
           URN& = ULTREG&("ENCASPEC")
           '
           If Len(RNC$) >= 4 Then
             RENOTA& = CVS(Left$(RNC$, 4))
             If RENOTA& > 0 Then
               If RENOTA& <= URN& Then
                 TTYY$ = REGLEIDO$("ENCASPEC", RENOTA&)
                 For KK% = 0 To 2
                   TTZZ$ = TRIM$(Mid$(TTYY$, 128 * KK% + 129, 128))
                   RETEX$(1) = "": RETEX$(2) = ""
                   Call FRATEXTO(TTZZ$, 64)
                   For KL% = 1 To 2
                     DOCPARAM$(5 + 2 * KK% + KL%) = RETEX$(KL%)
                   Next KL%
                 Next KK%
               End If
             End If
           End If
           '
           RNC$ = RECFILT$("ESPECIFI", 1, MKI$(CIXI%))
           URN& = ULTREG&("ESPECIFI")
           '
           For U& = 1 To Len(RNC$) Step 4
             RENOTA& = CVS(Mid$(RNC$, U&, 4))
             If RENOTA& > 0 Then
               If RENOTA& <= URN& Then
                 TTXY$ = RTrim$(Mid$(REGLEIDO$("ESPECIFI", RENOTA&), 3))
                 TTYY$ = REPLACAR$(TTXY$, Chr$(9), "     ")
                 CAITAB1% = CAITAB1% + 1
                 TETABU1$(1, CAITAB1%) = TTYY$
               End If
             End If
           Next U&
           '
           DESTIDOC% = 0
           TIPODOC$ = "Hoja de Especificaciones " + CODEXT$(CIXI%)
           Call PRINTDOC("@" + PRIFORM$)
         End If
       End If
     End If
   End If
   '
End Sub


