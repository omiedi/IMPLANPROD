Attribute VB_Name = "AFLEX_FUNCTIONS"
Public COTERMI$   ' codigo de terminal EB / FG
'
Function NEXTCOTELNU$()
   '
   If COTERMI$ = "" Then
'ITERMI$ = "TS"
'GoTo 15
      ITERMI$ = IDENTER$
10    COTERMI$ = CONTROL$(ITERMI$, "COTERMI")
      If Len(COTERMI$) <> 2 Then
        OPCOTEL.Show 1
        GoTo 10
      End If
   End If
   '
15 For MXC& = ULTREG&("COTELEF1") To 1 Step -1
      F$ = REGLEIDO$("COTELEF1", MXC&)
      If Mid$(F$, 11, 2) = COTERMI$ Then
         NUCONSU& = 1 + Val(Mid$(F$, 4, 6))
         GoTo 20
      End If
      NUCONSU& = 100001
   Next MXC&
   '
20 NEXTNU$ = TRIM$(Str$(NUCONSU&))
   While Len(NEXTNU$) < 6: NEXTNU$ = "0" + NEXTNU$: Wend
   NEXTCOTELNU$ = "CT-" + NEXTNU$ + "-" + COTERMI$
   '
End Function

Function NEXTORTRA$()
   '
   FEX = HOY(HOS$)
   COTERMJ$ = Mid$(FECHATEX$(FEX), 7, 2)
   OTX$ = "OT-" + COTERMJ$ + "-"
   '
   NUCONSU& = 1
   For MXC& = ULTREG&("CORTRAB1") To 1 Step -1
      F$ = REGLEIDO$("CORTRAB1", MXC&)
      If Mid$(F$, 4, 2) = COTERMJ$ Then
         NUCONSU& = 1 + Val(Mid$(F$, 7, 6))
         GoTo 20
      End If
   Next MXC&
   '
20 NEXTNU$ = TRIM$(Str$(NUCONSU&))
   While Len(NEXTNU$) < 4: NEXTNU$ = "0" + NEXTNU$: Wend
   NEXTORTRA$ = OTX$ + NEXTNU$
   '
End Function

Sub PRICOTEL(NUCOTE$)
   '
   Static CTXX%, RUTARGRA$
   Dim TENCABE$(3)
   '
   If CTXX% = 0 Then
     RUTARGRA$ = TRIM$(CONTROL$("", "RUTARGRA"))
     If RUTARGRA$ = "" Then RUTARGRA$ = TRIM$(LUDAT$)
     If Right$(RUTARGRA$, 1) <> "\" Then
       RUTARGRA$ = RUTARGRA$ + "\"
     End If
     CTXX% = 1
   End If
   '
   PRIFORM$ = TRIM$(CONTROL$("", "FORCOTEL"))
   CAIMAGEN% = 0
   '
   CODPARAM$(1) = "NUME-COM"
   CODPARAM$(2) = "FECH-EMI"
   CODPARAM$(3) = "HORA-EMI"
   CODPARAM$(4) = "RASO-CLI"
   CODPARAM$(5) = "TELE-CLI"
   CODPARAM$(6) = "DOMI-CLI"
   CODPARAM$(7) = "MEDIDA-1"
   CODPARAM$(8) = "MEDIDA-2"
   CODPARAM$(9) = "MEDIDA-3"
   CODPARAM$(10) = "REF-MAT1"
   CODPARAM$(11) = "FECH-LAN"
   CAPARDOC% = 11
   '
   If PRIFORM$ <> "" Then
     If EXISTE%(LUDAT$ + PRIFORM$ + ".FRM") Then
       If ECOARCH$("COTELEF1", NUCOTE$) <> "" Then
         Screen.MousePointer = 11
         CICOTE% = 0: For UI% = 1 To Len(NUCOTE$)
                        CICOTE% = CICOTE% + UI% * Asc(Mid$(NUCOTE$, UI%, 1))
                      Next UI%
                      '
         RNC$ = RECFILT$("COTELEF1", 3, MKI$(CICOTE%))
         URN& = ULTREG&("COTELEF1")
         '
         For U& = 1 To Len(RNC$) Step 4
            RENOTA& = CVS(Mid$(RNC$, U&, 4))
            If RENOTA& > 0 Then
               If RENOTA& <= URN& Then
                  TTXX$ = REGLEIDO$("COTELEF1", RENOTA&)
                  If Left$(TTXX$, 12) = NUCOTE$ Then
                     GoTo 29
                  End If
               End If
            End If
         Next U&
         Call MENSERR(24, "Faltan Datos de Encabezado de Consulta")
         Exit Sub
         '
29       DOCPARAM$(10) = TRIM$(Mid$(TTXX$, 111, 64))
         DOCPARAM$(1) = Left$(TTXX$, 12)
         NC% = CVI(Mid$(TTXX$, 81, 2))
         If NC% > 0 Then
            DOCPARAM$(4) = RASOCLI$(NC%)
            DOCPARAM$(5) = FAXCLI$(NC%)
         End If
         DOCPARAM$(6) = TRIM$(Mid$(TTXX$, 83, 24))
         DOCPARAM$(2) = FECHATEX$(CVI(Mid$(TTXX$, 107, 2)))
         DOCPARAM$(3) = HORATEX$(CVI(Mid$(TTXX$, 109, 2)))
         DOCPARAM$(7) = TRIM$(Mid$(TTXX$, 179, 24))
         DOCPARAM$(8) = TRIM$(Mid$(TTXX$, 203, 24))
         DOCPARAM$(9) = TRIM$(Mid$(TTXX$, 227, 24))
         '
         HOII% = HOY(HOS$)
         DOCPARAM$(11) = HOS$ + " - " + Left$(Time$, 5) + " hs"
         '
         For KK1% = 0 To 32
           For KK2% = 0 To 1024
             TETABU1$(KK1%, KK2%) = ""
             TETABU2$(KK1%, KK2%) = ""
           Next KK2%
         Next KK1%
         '
         CODTABU1$(1) = "REF-MAT3"
         CACOLTAB1% = 1
         CAITAB1% = 0
         '
         CODTABU2$(1) = "REF-MAT2"
         CACOLTAB2% = 1
         CAITAB2% = 0
         '
         RNC$ = RECFILT$("COTELEF2", 2, MKI$(CICOTE%))
         URE& = ULTREG&("COTELEF2")
         '
         For U& = 1 To Len(RNC$) Step 4
           RENOTA& = CVS(Mid$(RNC$, U&, 4))
           If RENOTA& > 0 Then
             If RENOTA& <= URE& Then
               TTXX$ = REGLEIDO$("COTELEF2", RENOTA&)
               If Left$(TTXX$, 12) = NUCOTE$ Then
                 TIRE$ = Mid$(TTXX$, 15, 2)
                 TTYY$ = Mid$(TTXX$, 17, 72)
                 If Left$(TTYY$, 1) = " " Then TTYY$ = "." + TTYY$
                 If TIRE$ = "DI" Then
                      CAITAB1% = CAITAB1% + 1
                      TETABU1$(1, CAITAB1%) = TTYY$
                    ElseIf TIRE$ = "CO" Then
                      CAITAB2% = CAITAB2% + 1
                      TETABU2$(1, CAITAB2%) = TTYY$
                 End If
               End If
             End If
           End If
         Next U&
         '
         DESTIDOC% = 0
         TIPODOC$ = "Consulta Telefonica"
         Call PRINTDOC("@" + PRIFORM$)
       End If
     End If
   End If
   '
End Sub

Sub PRIORTRA(NUCOTE$)
   '
   Static CTXX%, RUTARGRA$
   Dim TENCABE$(3)
   '
   If CTXX% = 0 Then
     RUTARGRA$ = TRIM$(CONTROL$("", "RUTARGRA"))
     If RUTARGRA$ = "" Then RUTARGRA$ = TRIM$(LUDAT$)
     If Right$(RUTARGRA$, 1) <> "\" Then
       RUTARGRA$ = RUTARGRA$ + "\"
     End If
     CTXX% = 1
   End If
   '
   PRIFORM$ = TRIM$(CONTROL$("", "FORTRABA"))
   CAIMAGEN% = 0
   '
   CODPARAM$(1) = "REF-MAT1"
   CODPARAM$(2) = "FECH-EMI"
   CODPARAM$(3) = "DOCUORIG"
   CODPARAM$(4) = "RASO-CLI"
   CODPARAM$(5) = "TALLE"
   CODPARAM$(6) = "COLOR"
   CODPARAM$(7) = "FECH-LAN"
   CODPARAM$(8) = "FECH-REC"
   CODPARAM$(9) = "FECH-VEN"
   CODPARAM$(10) = "OBS-OPER"
   CODPARAM$(11) = "REF-MAT4"
   CODPARAM$(12) = "REF-MAT5"
   CODPARAM$(13) = "REF-MAT6"
   
   CAPARDOC% = 13
   '
   If PRIFORM$ <> "" Then
     If EXISTE%(LUDAT$ + PRIFORM$ + ".FRM") Then
       If ECOARCH$("CORTRAB1", NUCOTE$) <> "" Then
         Screen.MousePointer = 11
         DOCPARAM$(1) = FADMIFLEX.Text9
         DOCPARAM$(2) = FADMIFLEX.Label14
         DOCPARAM$(3) = FADMIFLEX.Text8
         DOCPARAM$(4) = FADMIFLEX.Label8
         
           MOTI$ = "REPARACION"
           If FADMIFLEX.Option1(2).Value = True Then MOTI$ = "MODIFICACION"
           If FADMIFLEX.Option1(3).Value = True Then MOTI$ = "NUEVO DESARR."
           If FADMIFLEX.Option1(9).Value = True Then MOTI$ = "VARIOS"
         DOCPARAM$(5) = MOTI$
           MOTI$ = "ABONO"
           If FADMIFLEX.Option2(2).Value = True Then MOTI$ = "GARANTIA"
           If FADMIFLEX.Option2(3).Value = True Then MOTI$ = "FACTURACION"
           If FADMIFLEX.Option2(9).Value = True Then MOTI$ = "VARIOS"
         DOCPARAM$(6) = MOTI$
         DOCPARAM$(7) = FADMIFLEX.Text26
         DOCPARAM$(8) = FADMIFLEX.Text27
         DOCPARAM$(9) = FADMIFLEX.Text20
         DOCPARAM$(10) = FADMIFLEX.Text16
         '
         DOCPARAM$(11) = FADMIFLEX.Text18
         DOCPARAM$(12) = FADMIFLEX.Text17
         DOCPARAM$(13) = FADMIFLEX.Text19
         '
         CODTABU1$(1) = "REF-MAT2"
         CODTABU1$(2) = "REF-MAT3"
         CODTABU1$(3) = "DES-MAT"
         CODTABU1$(4) = "DES-OPER"
         CACOLTAB1% = 4
         CAITAB1% = 0
         CAITAB2% = 0
         '
         For KK1% = 0 To 32
           For KK2% = 0 To 1024
             TETABU1$(KK1%, KK2%) = ""
             TETABU2$(KK1%, KK2%) = ""
           Next KK2%
         Next KK1%
         '
         For U& = 1 To FADMIFLEX.List1.ListCount
           TTXX$ = TRIM$(Left$(FADMIFLEX.List1.List(U& - 1), 62))
           TTYY$ = Mid$(FADMIFLEX.List1.List(U& - 1), 63)
           If Left$(TTYY$, 1) = " " Then
             Mid$(TTYY$, 1, 1) = "_"
           End If
           CAITAB1% = CAITAB1% + 1
           TETABU1$(1, CAITAB1%) = TTXX$
           TETABU1$(2, CAITAB1%) = TTYY$
         Next U&
         '
         Call FRATEXTO(FADMIFLEX.Text15, 90)
         For KU% = 1 To CARETEX%
           TETABU1$(3, KU%) = RETEX$(KU%)
           If KU% > CAITAB1% Then CAITAB1% = KU%
         Next KU%
         '
         Call FRATEXTO(TRIM$(FADMIFLEX.Text14), 41)
         If CARETEX% > 6 Then CARETEX% = 6
         For KU% = 1 To CARETEX%
           TETABU1$(4, KU%) = RETEX$(KU%)
           If KU% > CAITAB1% Then CAITAB1% = KU%
         Next KU%
         '
         DESTIDOC% = 0
         TIPODOC$ = "Orden de Trabajo"
         NUMEDOC$ = TRIM$(FADMIFLEX.Label15)
         Call PRINTDOC("@" + PRIFORM$)
         '
       End If
     End If
   End If
   '
End Sub


