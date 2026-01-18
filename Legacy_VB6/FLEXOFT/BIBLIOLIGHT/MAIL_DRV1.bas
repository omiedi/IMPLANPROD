Attribute VB_Name = "MAIL_DRV1"
Public MAI_AUTO%
Public MAI_TEST%
'
Global MAI_SENDOK%
'
Public MAI_DESTI$
Public MAI_DIREL$
Public MAI_ASUNT$
Public MAI_TEXTO$
Public MAI_REPIT$
Public MAI_FTEXT$
Public MAI_COLEC$
Public MAI_ADJUN$(256)
'
Public MAI_SSMTP$
Public MAI_SNAME$
Public MAI_SMAIL$
Public MAI_UNAME$
Public MAI_PWORD$
'
Public ARENVI$
'
Sub GENERAMAIL(Optional MODO$, Optional CLASEN$)
  '
  TTXX$ = LOADFILE(LUCOM$ + "STPFMAIL.CFG")
  If TRIM$(Mid$(TTXX$, 553, 8)) = "CLIENT" Then
      Exit Sub    ' ES CLIENTE, NO SERVIDOR DE CORREP
  End If
  '
  If TRIM$(MAI_DIREL$) <> "" Then
    If InStr(MAI_DIREL$, "@") > 1 Then
      '
      RPTX$ = ""
      RPTY$ = MAI_REPIT$
      PPXX% = InStr(RPTY$, ";")
      If PPXX% > 0 Then
        VINTER$ = TRIM$(Left$(RPTY$, PPXX% - 1))
        RPTY$ = Mid$(RPTY$, PPXX% + 1) '  INTER$ = ""
        If Val(VINTER$) > 0 Then
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
10    LUCOMAIL$ = LUCOM$
      '
      On Error Resume Next
      MkDir LUCOM$ + "MAILCTRL"
      On Error GoTo 0
      If EXISTE%(LUCOM$ + "MAILCTRL\.") > 0 Then
         LUCOMAIL$ = LUCOM$ + "MAILCTRL\"
      End If
      If EXISTE%("F:\FLEXOFT\MAILCTRL\.") > 0 Then
         LUCOMAIL$ = "F:\FLEXOFT\MAILCTRL\"
      End If
      '
      If CLASEN$ <> "" Then
          CLA1$ = CLASEN$
        Else
          CLA1$ = RANDSTRING$(-6)
          ATRAMAI$ = LUCOMAIL$ + "" + CLA1$ + ".TMP"
          If EXISTE%(ATRAMAI$) > 0 Then GoTo 10 ' PARA QUE NO PISE EL EXISTENTE
      End If
      '
      ADJX$ = ""
      For U& = 0 To 256
        If TRIM$(MAI_ADJUN$(U&)) <> "" Then
          If ADJX$ <> "" Then ADJX$ = ADJX$ + " & "
          ADJX$ = ADJX$ + MAI_ADJUN$(U&)
        End If
      Next U&
      '
      TTXX$ = "DESTINO=" + MAI_DESTI$ + Chr$(124)
      TTXX$ = TTXX$ + "DIRELEC=" + MAI_DIREL$ + Chr$(124)
      TTXX$ = TTXX$ + "ASUNTOM=" + MAI_ASUNT$ + Chr$(124)
         If RPTX$ <> "" Then
           TTXX$ = TTXX$ + "REPEAT=" + RPTX$ + Chr$(124)
         End If
      If TRIM$(MAI_COLEC$) <> "" Then TTXX$ = TTXX$ + "SOLCONF=SI" + Chr$(124)
      '
      ' REVISA Y CONVIERTE TEXTO
      For KKIL& = 1 To Len(MAI_TEXTO$)
         ICARA% = Asc(Mid$(MAI_TEXTO$, KKIL&, 1))
         If ICARA% < 32 Or ICARA% > 122 Then GoTo 25
      Next KKIL&
      GoTo 30              ' NO CONVIERTE EL TEXTO
      '
25    ' GUARDA EL TEXTO COMO ARCHIVO
      AMAITEX$ = LUCOMAIL$ + "" + CLA1$ + ".TMX"
      If CLASEN$ = "" Then Call SAVEFILE(AMAITEX$, MAI_TEXTO$)
      MAI_TEXTO$ = "###" + AMAITEX$
      '
30    TTXX$ = TTXX$ + "TEXTOMA=" + MAI_TEXTO$ + Chr$(124)
      If TRIM$(MAI_FTEXT$) <> "" Then TTXX$ = TTXX$ + "FORTEXT=" + MAI_FTEXT$ + Chr$(124)
      TTXX$ = TTXX$ + "ADJUNTO=" + ADJX$ + Chr$(124)
      '
      If TRIM$(MAI_SSMTP$) <> "" Then TTXX$ = TTXX$ + "SMTPSERVER=" + MAI_SSMTP$ + Chr$(124)
      If TRIM$(MAI_SNAME$) <> "" Then TTXX$ = TTXX$ + "SENDERNAME=" + MAI_SNAME$ + Chr$(124)
      If TRIM$(MAI_SMAIL$) <> "" Then TTXX$ = TTXX$ + "SENDERMAIL=" + MAI_SMAIL$ + Chr$(124)
      If TRIM$(MAI_UNAME$) <> "" Then TTXX$ = TTXX$ + "USERNAME=" + MAI_UNAME$ + Chr$(124)
      If TRIM$(MAI_PWORD$) <> "" Then TTXX$ = TTXX$ + "PASSWORD=" + ENCRIPTA$(MAI_PWORD$) + Chr$(124)
      '
      ATRAMAI$ = LUCOMAIL$ + "" + CLA1$ + ".TMP"
      Call SAVEFILE(ATRAMAI$, TTXX$)
      ARENVI$ = ATRAMAI$
      '
      CLAX$ = CLA1$ + CONTRACLA$(CLA1$)
      CLAX$ = AJUSTI$(CLAX$, 13) + LUCOMAIL$
      If MODO$ <> "" Then CLAX$ = CLAX$ + MODO$
      '
      MODEJ% = 4                                           ' VISIBLE
      If InStr(UCase$(MODO$), "/AUTO") > 0 Then MODEJ% = 0 ' MINIMIZADO
      If EXISTE%(App.Path + "\FLXMWSK.EXE") > 0 Then
           APLIMAI$ = "FLXMWSK"
         ElseIf EXISTE%(App.Path + "\FLXMAIL.EXE") Then
           APLIMAI$ = "FLXMAIL"
         Else
           GoTo 40
      End If
      On Error Resume Next
      Shell App.Path + "\" + APLIMAI$ + " " + CLAX$, MODEJ%                ' EJECUTA MINIMIZADO
      If Err Then
40        MsgBox "Faltan Utilitarios 'FLXMWSK.EXE' o 'FLXMAIL.EXE' en Carpeta " + App.Path + "."
      End If
      On Error GoTo 0
      '
    End If
  End If
  '
End Sub

Sub ENVIAMAIL(DESTI$, DIREL$, ASUNT$, TEXTO$, Optional ADJUN$, Optional MODO$)
  '
  If DIREL$ = "" Or InStr(DIREL$, "@") < 1 Then
     MsgBox "Imposible Enviar e-mail. Dir-Electronica Destino '" + DIREL$ + "' no Válida"
     Exit Sub
  End If
  '
  If ASUNT$ = "" Then ASUNT$ = "Sin Asunto"
  If DESTI$ = "" Then DESTI$ = DIREL$
  '
  If EXISTE%("FLXMAIL.EXE") > 0 Then
      MAI_DESTI$ = DESTI$
      MAI_DIREL$ = DIREL$                 ' "ebauer@flexoft.com.ar"
      MAI_ASUNT$ = ASUNT$
      MAI_TEXTO$ = TEXTO$
      If ADJUN$ <> "" Then
         MAI_ADJUN$(0) = ADJUN$
      End If
      Call GENERAMAIL(MODO$)
    Else
      FLEXLOG3.MAPISession1.DownLoadMail = False
      FLEXLOG3.MAPISession1.NewSession = True
      FLEXLOG3.MAPISession1.SignOn
      FLEXLOG3.MAPIMessages1.SessionID = FLEXLOG3.MAPISession1.SessionID
      '
      FLEXLOG3.MAPIMessages1.Compose
      FLEXLOG3.MAPIMessages1.RecipDisplayName = DESTI$
      FLEXLOG3.MAPIMessages1.RecipAddress = DIREL$
      FLEXLOG3.MAPIMessages1.MsgSubject = ASUNT$
      FLEXLOG3.MAPIMessages1.MsgNoteText = TEXTO$
      '
      While Len(ADJUN$) > 0
        PPXX% = InStr(ADJUN$, "&")
        If PPXX% < 1 Then PPXX% = 1 + Len(ADJUN$)
        ARAD$ = TRIM$(Left$(ADJUN$, PPXX% - 1))
        ADJUN$ = TRIM$(Mid$(ADJUN$, PPXX% + 1))
        ARAD1$ = ARAD$: If InStr(ARAD$, ".") < 1 Then ARAD1$ = ARAD$ + "."
        If ARAD$ <> "" Then
          If EXISTE%(ARAD1$) > 0 Then      ' PARA QUE ADMITA SIN EXTENSIÓN
            FLEXLOG3.MAPIMessages1.AttachmentIndex = FLEXLOG3.MAPIMessages1.AttachmentCount
            FLEXLOG3.MAPIMessages1.AttachmentPathName = ARAD$
          End If
        End If
      Wend
      '
      On Error Resume Next
      FLEXLOG3.MAPIMessages1.Send False
      FLEXLOG3.MAPISession1.SignOff
      On Error GoTo 0
      DoEvents
  End If
  '
End Sub


