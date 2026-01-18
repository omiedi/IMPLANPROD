Attribute VB_Name = "MAIL_DRV"
Public MAI_DESTI$
Public MAI_DIREL$
Public MAI_COPIA$
Public MAI_ASUNT$
Public MAI_TEXTO$
Public MAI_REPIT$
Public MAI_ADJUN$(256)
'
Sub GENERAMAIL()
  '
  
  RPTX$ = ""
  RPTY$ = MAI_REPIT$
  PPXX% = InStr(RPTY$, ";")
  If PPXX% > 0 Then
    VINTER$ = TRIM$(Left$(RPTY$, PPXX% - 1))
    RPTY$ = Mid$(RPTY$, PPXX% + 1) '  INTER$ = ""
    If XVALO(VINTER$) > 0 Then
      While Len(VINTER$) < 3: VINTER$ = "0" + VINTER$: Wend
      PPXX% = InStr(RPTY$, ";")
      If PPXX% > 0 Then
        INTER$ = TRIM$(Left$(RPTY$, PPXX% - 1))
        FEFIN$ = Mid$(RPTY$, PPXX% + 1, 8)
        If FECHANUM(FEFIN$) >= HOY(HOS$) Then
          RPTX$ = "FREC" + VINTER$ + INTER$ + ";END" + FEFIN$
        End If
      End If
    End If
  End If
  
  '
  
  If TRIM$(MAI_DIREL$) <> "" Then
    If InStr(MAI_DIREL$, "@") > 1 Then
      ADJX$ = ""
      For U& = 0 To 256
        If TRIM$(MAI_ADJUN$(U&)) <> "" Then
          If ADJX$ <> "" Then ADJX$ = ADJX$ + " & "
          ADJX$ = ADJX$ + MAI_ADJUN$(U&)
          MAI_ADJUN$(U&) = ""
        End If
      Next U&
      '
      TTXX$ = "DESTINO=" + MAI_DESTI$ + Chr$(124)
      TTXX$ = TTXX$ + "DIRELEC=" + MAI_DIREL$ + Chr$(124)
      If TRIM$(MAI_COPIA$) <> "" Then
         TTXX$ = TTXX$ + "DIRECOP=" + MAI_COPIA$ + Chr$(124)
      End If
      TTXX$ = TTXX$ + "ASUNTOM=" + MAI_ASUNT$ + Chr$(124)
         If RPTX$ <> "" Then
           TTXX$ = TTXX$ + "REPEAT=" + RPTX$ + Chr$(124)
         End If
      TTXX$ = TTXX$ + "TEXTOMA=" + MAI_TEXTO$ + Chr$(124)
      TTXX$ = TTXX$ + "ADJUNTO=" + ADJX$ + Chr$(124)
      TTXX$ = TTXX$ + "SOLCONF=SI" + Chr$(124)
      '
      CLA1$ = RANDSTRING$(-6)
      ATRAMAI$ = LUCOM$ + "" + CLA1$ + ".TMP"
      Call SAVEFILE(ATRAMAI$, TTXX$)
      Call SAVEFILE(LUCOM$ + "TEMPMAIL.WKF", TTXX$)
      Call BAJALDISCO
      '
      CLAX$ = CLA1$ + CONTRACLA$(CLA1$)
      RUTEX$ = ""
      If INTERPRE% = 1 Then
         RUTEX$ = "F:\FLEXOFT\NOVEDAD\"
      End If
      On Error Resume Next
      If EXISTE%(LUCOM$ & "FLXMAIL_OUTLOOK.EXE") = 1 Then
            TEXTOMAIL1$ = Chr$(34) & TTXX$ & Chr$(34)
            RESPUESTA = Shell(LUCOM$ & "flxmail_outlook.exe " & TEXTOMAIL1$, vbMinimizedNoFocus)
            Exit Sub
      End If
      Shell RUTEX$ + "FLXMAIL " + CLAX$ + Chr$(124) + LUCOM$, 0 ' EJECUTA MINIMIZADO
      If Err Then
        On Error GoTo 0
        TTXX$ = "Falta Utilitario 'FLXMAIL.EXE' en Carpeta "
        If RUTEX$ = "" Then
             TTXX$ = TTXX$ + App.Path
           Else
             TTXX$ = TTXX$ + RUTEX$
        End If
        MsgBox TTXX$
      End If
      On Error GoTo 0
      '
    End If
  End If
  '
End Sub
