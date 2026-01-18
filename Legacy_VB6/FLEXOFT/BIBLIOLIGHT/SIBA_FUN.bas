Attribute VB_Name = "SIBA_FUN"
Sub FINAL(MODU$)
    '
    If TRIM$(MODU$) = "" Then
         Dim NOFORMU$(256)
         CAFORMU% = Forms.Count
         For KKI% = 1 To Forms.Count
            NOFORMU$(KKI%) = Forms(KKI% - 1).Name
         Next KKI%
         Call CIERRARCH("*.*")
         Call BLOQEXE("*")
         Call LIBARCH("*.*")
         Close
         WAITFILE$ = LUCOM$ + JOBID$ + ".DRV"
15       EXIWAIT% = EXISTE%(WAITFILE$)
         If EXIWAIT% > 0 Then
            On Error Resume Next
            Kill WAITFILE$
            On Error GoTo 0
            GoTo 15
         End If
         On Error GoTo 0
         On Error Resume Next
10       For KKI% = CAFORMU% To 1 Step -1
            For KKJ% = 1 To Forms.Count
              If Forms(KKJ% - 1).Name = NOFORMU$(KKI%) Then
                Forms(KKJ% - 1).Hide
                Unload Forms(KKJ% - 1)
                DoEvents
              End If
            Next KKJ%
         Next KKI%
         On Error GoTo 0
         Close: End
       ElseIf Left$(MODU$, 1) = "*" Then
         Call CIERRARCH("*.*")
         Call CONECRUN(MODU$, "")
       Else
         Call MENSERR(24, "Error de Conexion\a Módulo " + MODU$)
         Call CIERRARCH("*.*")
         Close: End
    End If
    '
End Sub

Sub CONECRUN(MODU$, APLI$)
      '
      If DERACCE%(MODU$, APLI$) < 1 Then
         Exit Sub
      End If
      '
      If Left$(MODU$, 1) = "*" Then
           CONECTA$ = UCase$(MODU$)
         Else
           CONECTA$ = UCase$(AJUSTI$(MODU$, 8))
      End If
      '
199   TITU$ = AJUSTI$(APLI$, 32)
      If TITU$ = "" Then TITU$ = "MODULO SIN IDENTIFICACION"
      OPERA$ = TITU$
      '
      If Left$(UCase$(MODU$), 3) = "FAC" Then
        HOII% = HOY(HOS$)
        If EXISTE%(LUDAT$ + "VENCICAI.RFS") = 1 Then
          For UU& = 1 To ULTREG&("VENCICAI")
            XYZ$ = REGLEIDO$("VENCICAI", UU&)
            FEVENCAI% = CVI(Mid$(XYZ$, 63, 2))
            If FEVENCAI% > 0 Then
              If FEVENCAI% < HOII% Then
                  DEFORMU$ = TRIM$(Left$(XYZ$, 32))
                  Call MENSERR(24, "Imposible Acceder a Módulo '" + TRIM$(UCase$(MODU$)) + "'\CAI de " + DEFORMU$ + " está Vencido.")
                  Exit Sub
                ElseIf FEVENCAI% < HOII% + 32 Then
                  DEFORMU$ = TRIM$(Left$(XYZ$, 32))
                  Call MENSERR(24, "ATENCION: \CAI de " + DEFORMU$ + "\Próximo a Vencer !!!")
              End If
            End If
          Next UU&
        End If
      End If
      '
      If Left$(MODU$, 8) = "WEDITABU" Then
          '
          DIRA$ = TRIM$(CurDir(""))
          If InStr(DIRA$, " ") > 0 Then
               MsgBox "Esta Aplicación no Puede Ejecutarse en la Carpeta de Instalación " + DIRA$ + ". reinstale el Sistema en \FLEXOFT\APLICA."
               Exit Sub
          End If
          '
          If EXISTE%(LUBAS$ + MODU$ + ".EXE") < 1 Then
              BRXX$ = REPLACAR$(LUBAS$ + MODU$ + ".EXE", "\", "/")
              Call MENSERR(24, "No Existe " + BRXX$)
              Exit Sub
          End If
          '
          DRIVA$ = Left$(DIRA$, 2)
          DIRA$ = Mid$(DIRA$, 3)
          If Mid$(DRIVA$, 2, 1) <> ":" Then
               Call COMUNI("Error de Instalación y/o de Configuración.\Imposible Identificar  la Unidad de Disco.\ \Consulte a su Soporte de Sistemas FLEXOFT.")
               Exit Sub
          End If
          '
          NX% = FILEOPEN%("C:\FLEXOFT\FLEXBAT.BAT", 2, 0)
          Print #NX%, "@ECHO OFF"
          Print #NX%, DRIVA$
          Print #NX%, "CD " + DIRA$
          Print #NX%, LUBAS$ + MODU$
          Print #NX%, "EXIT"
          Close #NX%: NX% = 0
          '     '
          MDX$ = AJUSTI$("*" + Mid$(MODU$, 13, 8), 8): Call GRARAM(21, MDX$)
          APX$ = AJUSTI$(APLI$, 32): Call GRARAM(29, APX$)
          Call BAJRAM
          '
          CCI$(0) = "WEDITABU\" + Mid$(MDX$, 2)
          Call GRATRACO(RAZO$, EMPRE$, SE$, LU$, QWERTY%, CCI$(), EJU%, Menu%)
          '
          LCB$ = LUBAS$
          If EXISTE%("C:\FLEXOFT\FLEXPIF.PIF") = 1 Then
              Shell "C:\FLEXOFT\FLEXPIF.PIF", vbNormalFocus
            Else
              TXT$ = REPLACAR$("C:\FLEXOFT\FLEXPIF.PIF", "\", "/")
              Call MENSERR(24, "No Existe " + TXT$)
          End If
          '
          Exit Sub
          '
      End If
      '
210   If Left$(CONECTA$, 1) = "#" Then
          '
          ATRIBU$ = ""
          PPX% = InStr(MODU$, "/")
          If PPX% > 0 Then
              CONECTA$ = Left$(MODU$, PPX% - 1)
              ATRIBU$ = Mid$(MODU$, PPX% + 1)
          End If
          '
          CONECTA$ = TRIM$(Mid$(CONECTA$, 2))
          AX$ = CONECTA$ + ".EXE"
          If Left$(CONECTA$, 4) = "FLEX" Then AX$ = LUBAS$ + AX$
          If Left$(CONECTA$, 4) = "PARA" Then AX$ = LUBAS$ + AX$
          '
          If EXISTE%(AX$) < 1 Then
             BRXX$ = REPLACAR$(AX$, "\", "/")
             Call OPNOIN(BRXX$)
             Exit Sub
          End If
          '
          If COMALTER%("No se Encontró la Versión para Windows\de la Aplicación Elegida.\  \ Puede ejecutar la Versión D.O.S. o Abandonar\\Ejecutar D.O.S.\Abandonar") <> 1 Then
             Exit Sub
          End If
          '
          DIRA$ = CurDir("")
          If InStr(DIRA$, " ") > 0 Then
             MsgBox "Esta Aplicación no Puede Ejecutarse en la Carpeta de Instalación " + DIRA$ + ". reinstale el Sistema en \FLEXOFT\WINDOWS."
             Exit Sub
          End If
          '
          DRIVA$ = Left$(DIRA$, 2)
          DIRA$ = Mid$(DIRA$, 3)
          If Mid$(DRIVA$, 2, 1) <> ":" Then
             Call COMUNI("Error de Instalación y/o de Configuración.\Imposible Identificar  la Unidad de Disco.\ \Consulte a su Soporte de Sistemas FLEXOFT.")
             Exit Sub
          End If
          '
          INVOCA$ = AX$
          If ATRIBU$ <> "" Then INVOCA$ = INVOCA$ + " " + ATRIBU$
          NX% = FILEOPEN%("C:\FLEXOFT\FLEXBAT.BAT", 2, 0)
          Print #NX%, "@ECHO OFF"
          Print #NX%, DRIVA$
          Print #NX%, "CD " + DIRA$
          Print #NX%, INVOCA$
          Print #NX%, "EXIT"
          Close #NX%: NX% = 0
          '
          MDX$ = AJUSTI$(MODU$, 8): Call GRARAM(21, MDX$)
          APX$ = AJUSTI$(APLI$, 32): Call GRARAM(29, APX$)

          Call BAJRAM
          '
          If CONECTA$ <> "FILTABU" Then
              CCI$(0) = CONECTA$
          End If
          Call GRATRACO(RAZO$, EMPRE$, SE$, LU$, QWERTY%, CCI$(), EJU%, Menu%)
          Call CIERRARCH("*.*")
          Close
          '
          LCB$ = LUBAS$
          If EXISTE%("C:\FLEXOFT\FLEXPIF.PIF") = 1 Then
                Shell "C:\FLEXOFT\FLEXPIF.PIF", vbNormalFocus
              Else
                TXT$ = REPLACAR$("C:\FLEXOFT\FLEXPIF.PIF", "\", "/")
                Call MENSERR(24, "No Existe " + TXT$)
          End If
          Exit Sub
          '
      ElseIf Left$(CONECTA$, 1) = "*" Then
           '
          If EXISTE%(LUBAS$ + "WINTABU.EXE") < 1 Then
              BRXX$ = REPLACAR$(LUBAS$ + "WINTABU.EXE", "\", "/")
              Call MENSERR(24, "No Existe " + BRXX$)
              Exit Sub
          End If
          '
          MDX$ = AJUSTI$(MODU$, 8): Call GRARAM(21, MDX$)
          APX$ = APLI$
          PPX% = InStr(APX$, "/WAIT")
          ESPERA% = 0
          If PPX% > 0 Then
             APX$ = Left$(APX$, PPX% - 1)
             ESPERA% = 1
          End If
          APY$ = AJUSTI$(APX$, 32)
          Call GRARAM(29, APY$)
          Call GRAJOBID
          Call BAJRAM
          '
          YACARWI% = 0
          TINI = 3600 * Val(Left$(Time$, 2)) + 60 * Val(Mid$(Time$, 4, 2)) + Val(Mid$(Time$, 7, 2))
          HOII% = HOY(HOS$)
          '
          If UCase$(App.EXEName) = "WINTABU" Then
             Call BLOQEXE("*")
          End If
If NUEWINTA% = 1 Then
   CONECWIN.Label1.Caption = CONECTA$
   CONECWIN.Show 1
   DoEvents
   On Error Resume Next
   For KKU% = Forms.Count To 1 Step -1
     If Forms(KKU%).Name <> "CONECWIN" Then
       Forms(KKU% - 1).Enabled = True
       If Forms(KKU% - 1).Visible = True Then
         Forms(KKU%).SetFocus
         GoTo 19
       End If
     End If
   Next KKU%
19 On Error GoTo 0
   Exit Sub
End If
          Shell LUBAS$ + "WINTABU " + CONECTA$ + "#" + JOBID$ + "#", 1
          ' Exit Sub
          '
          Screen.MousePointer = 11
          YACARWI% = 0
          Dim WTT As String * 128
240       On Error GoTo 250
          NK% = FreeFile
          Open LUCOM$ + "WINTABU.EXK" For Random Access Read Write As #NK% Len = 128
          Get #NK%, 1, WTT$
          On Error GoTo 0
          Close #NK%: NK% = 0
          If Mid$(WTT$, 69, 8) <> HOS$ Then GoTo 240
          TIMF$ = Mid$(WTT$, 79, 8)
          tinj = 3600 * Val(Left$(TIMF$, 2)) + 60 * Val(Mid$(TIMF$, 4, 2)) + Val(Mid$(TIMF$, 7, 2))
          If tinj < TINI Then GoTo 240
          '
          Call FRESHALL
          Screen.MousePointer = 1
          Exit Sub
          '
250       Resume 251
251       YACARWI% = 1
          GoTo 240
          '
      Else
          MDY$ = MODU$
          ATRIBU$ = ""
          PPX% = InStr(MDY$, "/")
          If PPX% > 0 Then
              MDY$ = Left$(MODU$, PPX% - 1)
              ATRIBU$ = Mid$(MODU$, PPX% + 1)
          End If
          '
          RECARGA% = 0
          MDY$ = TRIM$(MDY$)
          MDX$ = AJUSTI$(MDY$, 8): Call GRARAM(21, MDX$)
          APX$ = APLI$
          PPX% = InStr(APX$, "/WAIT")
          ESPERA% = 0
          If PPX% > 0 Then
             APX$ = Left$(APX$, PPX% - 1)
             ESPERA% = 1
          End If
          APY$ = AJUSTI$(APX$, 32)
          Call GRARAM(29, APY$)
          Call GRAJOBID
          Call BAJRAM
          If UCase$(TRIM$(MDX$)) = "WINTABU" Then
              MDX$ = LUBAS$ + "WINTABU"
          End If
          '
          If ESPERA% = 1 Then
             N1% = FILEOPEN%(LUCOM$ + JOBID$ + ".DRV", 2, 0)
             Print #N1%, "ESPERAR"
             Close #N1%: N1% = 0
          End If
          '
          If REACTIVA%(MDX$) = 1 Then Exit Sub
          '
          On Error GoTo OPNOIN1
          Shell TRIM$(MDX$) + ".EXE " + ATRIBU$ + "#" + JOBID$ + "#", 1
          On Error GoTo 0
          If ESPERA% = 1 Then
             While EXISTE%(LUCOM$ + JOBID$ + ".DRV") > 0
             Wend
          End If
          Exit Sub
          '
      End If
      '
OPNOIN:
      Call COMUNI(MENSA$)
      Exit Sub
      '
OPNOIN1:
      Resume OPNOIN2
OPNOIN2:
      On Error GoTo 0
      TXT$ = "La Opción Elegida  no se Encuentra Disponible\en la Versión FLEXOFT  Actualmente Instalada."
      TXT$ = TXT$ + "\  \Aplicación: " + UCase$(MDX$) + ".EXE" + "\  \"
      TXT$ = TXT$ + "Consulte con su Soporte de Sistemas o su Pro-\veedor de Software para la Adquisición e Insta-\lación de la misma."
      Call COMUNI(TXT$)
      Exit Sub
      '
End Sub
'

Function DERACCE%(MDX$, APL$)
    '
    DRXXX% = 0
    QWR% = USNBR% Mod 100
    If QWR% = 1 Then
        DRXXX% = 2
        GoTo 99                 ' supervisor entra directo
    End If
    '
    'If ULTREG&("FLEXACC") < 1 Then
    '    DRXXX% = 2
    '    GoTo 99                 ' no hay derechos definidos
    'End If
    '
    MDY$ = TRIM$(UCase$(MDX$))
    If MDY$ = "" Then
        DRXXX% = 2
        GoTo 99
    End If
    '
    PPXX% = InStr(MDY$, "/")
    If PPXX% > 1 Then
       ATRICON$ = Mid$(MDY$, PPXX% + 1)
       MDY$ = Left$(MDY$, PPXX% - 1)
    End If
    '
    PPY$ = Left$(MDY$, 1)
    If Asc(PPY$) < 48 Or (Asc(PPY$) > 57 And Asc(PPY$) < 65) Or Asc(PPY$) > 90 Then
        MDY$ = Mid$(MDY$, 2)
    End If
    '
    MDY$ = TRIM$(MDY$)
    If MDY$ = "" Then
        DRXXX% = 2
        GoTo 99
    End If
    '
    For TYU& = 1 To ULTREG&("FLEXACC")
        XTY$ = REGLEIDO$("FLEXACC", TYU&)
        If CVI(Left$(XTY$, 2)) = QWR% Then
            If TRIM$(UCase$(Mid$(XTY$, 3, 8))) = MDY$ Then
                If Mid$(XTY$, 11, 1) = "S" Then DRXXX% = 1
                If Mid$(XTY$, 12, 1) = "S" Then DRXXX% = 2
                Call CIERRARCH("FLEXACC")
                GoTo 99
            End If
        End If
    Next TYU&
    Call CIERRARCH("FLEXACC")
    '
    APM$ = TRIM$(UCase$(APL$))
    If APM$ = "" Then
       If Left$(MDX$, 1) = "*" Then
          MDZ$ = TRIM$(Mid$(MDX$, 2))
          If EXISTE%(MDZ$ + ".PRF") > 0 Then
             Dim YYY As String * 128
             N15% = FILEOPEN%(MDZ$ + ".PRF", 9, 128)
             Get #N15%, 1, YYY$
             APM$ = TRIM$(Left$(YYY$, 48))
             Close #N15%: N15% = 0
          End If
       End If
    End If
    '
    For U& = 1 To ULTREG&("INDIAPLI")
        X$ = REGLEIDO$("INDIAPLI", U&)
        If TRIM$(Left$(X$, 12)) = MDY$ Then
            If APM$ <> "" Then
                Call REPLA(X$, APM$, 13, 36)
                Call GRAREG("INDIAPLI", X$, U&)
            End If
            GoTo 20
        End If
    Next U&
    '
    If APM$ <> "" Then
         X$ = REGBLAN$("INDIAPLI")
         Call REPLA(X$, MDY$, 1, 12)
         Call REPLA(X$, APM$, 13, 36)
         Call REGAPP("INDIAPLI", X$)
       Else
        DRXXX% = 2
        Call CIERRARCH("INDIAPLI")
        GoTo 99
    End If
    '
20  Call CIERRARCH("INDIAPLI")
    If Left$(ATRICON$, 6) = "NOMESS" Then GoTo 99 ' NO PRESENTA MENSAJE
    '
    MENSA$ = "Usuario no Habilitado\"
    MENSA$ = MENSA$ + "para Ejecutar esta Opción.\  \"
    MENSA$ = MENSA$ + "Modulo: " + TRIM$(UCase$(MDY$)) + "\"
    MENSA$ = MENSA$ + "Aplicacion: " + TRIM$(APM$) + "\   \"
    MENSA$ = MENSA$ + "Solicite la Autorizacion Correspondiente\"
    MENSA$ = MENSA$ + "al Supervisor del Sistema FLEXOFT."
    '
    Call COMUNI(MENSA$)
    '
99  DERACCE% = DRXXX%
    '
End Function
'

Function DERACTU%()
    Static DRXXX%
    If DRXXX% = 0 Then
        '
        QWR% = USNBR% Mod 100
        If QWR% = 1 Then
            DRXXX% = 1
            GoTo 99                 ' supervisor entra directo
        End If
        '
        'If ULTREG&("FLEXACC") < 1 Then
        '    DRXXX% = 1
        '    GoTo 99                 ' no hay derechos definidos
        'End If
        '
        MDY$ = TRIM$(UCase$(App.EXEName))
        If MDY$ = "" Then
            DRXXX% = 1
            GoTo 99
        End If
        '
        If MDY$ = "WINTABU" Then
            DRXXX% = 1
            GoTo 99
        End If
        '
        LU$ = LUDAT$
        Dim XYZT As String * 48
        NUY% = FILEOPEN%(LU$ + "INDIAPLI.DAT", 0, 48)
        For U& = 1 To LOF(NUY%) / 48
            Get #NUY%, U&, XYZT$
            If TRIM$(Left$(XYZT$, 12)) = MDY$ Then
                Close #NUY%: NUY% = 0
                GoTo 20
            End If
        Next U&
        Close #NUY%: NUY% = 0
        DRXXX% = 1
        GoTo 99
        '
20      NUY% = FILEOPEN%(LU$ + "FLEXACC.DAT", 0, 48)
        For U& = 1 To LOF(NUY%) / 48
            Get #NUY%, U&, XYZT$
            If CVI(Left$(XYZT$, 2)) = QWR% Then
                If TRIM$(UCase$(Mid$(XYZT$, 3, 8))) = TRIM$(UCase$(MDY$)) Then
                    If Mid$(XYZT$, 12, 1) = "S" Then
                         DRXXX% = 1
                       Else
                         DRXXX% = (-1)
                    End If
                    Close #NUY%: NUY% = 0
                    GoTo 99
                End If
            End If
        Next U&
        Close #NUY%: NUY% = 0
        DRXXX% = (-1)
        '
    End If
    '
99  DERACTU% = DRXXX%
    '
End Function
'

