Attribute VB_Name = "CTRX_DRV"
Function CONTROVENCIM%(MODULOX$)
   '
   CTRV% = 0
   Call ARCONTROL(OKKX%)
   If OKKX% < 1 Then GoTo 90
   '
   CTRV% = VECONTROL%(MODULOX$)
   '
90 CONTROVENCIM% = CTRV%
   '
End Function

Sub ARCONTROL(OKKX%)
   '
   OKKX% = 0
   If EXISTE%("FLXCTRL.EXE") < 1 Then
      '
      CLAX$ = RANDSTRING$(-6)
      CONTRASE$ = InputBox$("Santo y Seña (" + CLAX$ + ")", "Contraseña para Habilitación Basica del Sistema")
      If CONTRASE$ <> CONTRACLA$(CLAX$) Then Exit Sub
      '
      ABC$ = LOADFILE$("FLOGON.EXE")
      If Len(ABC$) < 129 Then
        MsgBox "Falta Utilitario 'FLOGON.EXE'"
        Exit Sub
      End If
      DEF$ = Left$(ABC$, 129)
      '
      CD$ = Left$(CurDir$, 2)
      SEVOLF$ = Left$(SERIVOL$(CD$) + Space$(12), 12)
      HOII% = HOY(HOS$)
      '
      Call REPLA(DEF$, ENCRIPTA$(SEVOLF$), 1, 12)
      Call REPLA(DEF$, ENCRIPTA$("FLSYSTEM"), 17, 8)
      Call REPLA(DEF$, MKI$(HOII%), 33, 2)      ' fecha instalacion
      Call REPLA(DEF$, MKI$(HOII%), 35, 2)      ' fecha ultimo acceso
      Call REPLA(DEF$, MKI$(HOII% + 33), 37, 2) ' fecha limite utilizacion
      Call REPLA(DEF$, MKI$(0), 39, 2)
      DIVERI% = 0
      For UI% = 1 To 127
        DIVERI% = DIVERI% + Asc(Mid$(DEF$, UI%, 1)) - 123
      Next UI%
      Call REPLA(DEF$, MKI$(DIVERI%), 128, 2)
      Call REPLA(ABC$, DEF$, 1, 129)
      Call SAVEFILE("FLXCTRL.EXE", ABC$)
   End If
   OKKX% = 1
   '
End Sub
'
Function VECONTROL%(MODULOX$)
   '
   VCTRL% = 0
   CD$ = Left$(CurDir$, 2)
   SEVOLF$ = Left$(SERIVOL$(CD$) + Space$(12), 12)
   HOII% = HOY(HOS$)
   '
   ABUX$ = ENCRIPTA$(MODULOX$)
   DEFT$ = LOADFILE$("FLXCTRL.EXE")
   '
   For U& = 1 To Len(DEFT$) Step 129
     DEF$ = Mid$(DEFT$, U&, 129)
     If Mid$(DEF$, 17, Len(ABUX$)) = ABUX$ Then
        GoTo 20
     End If
   Next U&
   '
   While ((Len(DEFT$)) Mod 129) <> 0
     DEFT$ = DEFT$ + Chr$(0)
   Wend
   '
   DEF$ = RANDSTRING$(129)
   Call REPLA(DEF$, ENCRIPTA$(SEVOLF$), 1, 12)
   Call REPLA(DEF$, ENCRIPTA$(MODULOX$), 17, 8)
   Call REPLA(DEF$, MKI$(HOII%), 33, 2)      ' fecha instalacion
   Call REPLA(DEF$, MKI$(HOII%), 35, 2)      ' fecha ultimo acceso
   Call REPLA(DEF$, MKI$(HOII% + 33), 37, 2) ' fecha limite utilizacion
   Call REPLA(DEF$, MKI$(0), 39, 2)
   DIVERI% = 0
   For UI% = 1 To 127
     DIVERI% = DIVERI% + Asc(Mid$(DEF$, UI%, 1)) - 123
   Next UI%
   Call REPLA(DEF$, MKI$(DIVERI%), 128, 2)
   DEFT$ = DEFT$ + DEF$
   Call SAVEFILE("FLXCTRL.EXE", DEFT$)
   '
20 DIVERI% = 0
   For UI% = 1 To 127
     DIVERI% = DIVERI% + Asc(Mid$(DEF$, UI%, 1)) - 123
   Next UI%
   If CVI(Mid$(DEF$, 128, 2)) <> DIVERI% Then
      MsgBox "Fallo en Dígito Verificador"
      GoTo 90 ' copia pirata
   End If
   '
   If DESENCRI$(Left$(DEF$, 12)) <> SEVOLF$ Then
      MsgBox "Error Grave de Instalación." + Chr$(13) + Chr$(10) + "No Corresponde Entorno (SVLF)"
      CLAX$ = RANDSTRING$(-6)
      CONTRASE$ = InputBox$("Santo y Seña (" + CLAX$ + ")", "Contraseña para Re-Habilitación del Sistema en Entorno")
      If CONTRASE$ = CONTRACLA$(CLAX$) Then
        Close: Reset
        On Error Resume Next
        Kill "FLXCTRL.EXE"
        If Err Then
            MsgBox "Imposible Rehabilitar - Consulte !"
          Else
            MsgBox "Re-Habilitación Iniciada Exitosamente." + Chr$(13) + Chr$(10) + "Para Terminar de Rehabilitar Sistema" + Chr$(13) + Chr$(10) + "Cierre y Re-Inicie la Aplicación."
        End If
      End If
      GoTo 90 ' copia pirata
   End If
   '
30 VCTRL% = 1
   GoTo 90
   '
   ' el mismo dia no controla
   If CVI(Mid$(DEF$, 35, 2)) = HOII% Then
     VCTRL% = 1
     GoTo 90    ' coincide fecha ultimo acceso
   End If
   '
   Call REPLA(DEF$, MKI$(HOII%), 35, 2)
   DIVERI% = 0
   For UI% = 1 To 127
     DIVERI% = DIVERI% + Asc(Mid$(DEF$, UI%, 1)) - 123
   Next UI%
   Call REPLA(DEF$, MKI$(DIVERI%), 128, 2)
   '
   ABC$ = LOADFILE$("FLXCTRL.EXE")
   Call REPLA(ABC$, DEF$, 1, 129)
   Call SAVEFILE("FLXCTRL.EXE", ABC$)
   If ABC$ <> LOADFILE$("FLXCTRL.EXE") Then
     MsgBox "Acceso Denegado a Controlador"
     GoTo 90
   End If
   '
   FLUTIL% = CVI(Mid$(DEF$, 37, 2))
   CACCAD% = CVI(Mid$(DEF$, 39, 2))
   If FLUTIL% < HOII% Or CACCAD% > 0 Then
     CACCAD% = CACCAD% + 1
     Call REPLA(DEF$, MKI$(CACCAD%), 39, 2)
     DIVERI% = 0
     For UI% = 1 To 127
       DIVERI% = DIVERI% + Asc(Mid$(DEF$, UI%, 1)) - 123
     Next UI%
     Call REPLA(DEF$, MKI$(DIVERI%), 128, 2)
     '
     ABC$ = LOADFILE$("FLXCTRL.EXE")
     Call REPLA(ABC$, DEF$, 1, 129)
     Call SAVEFILE("FLXCTRL.EXE", ABC$)
     If ABC$ <> LOADFILE$("FLXCTRL.EXE") Then
       MsgBox "Acceso Denegado a Controlador"
       GoTo 90
     End If
   End If
   '
   CACCAD% = CVI(Mid$(DEF$, 39, 2))
   If CACCAD% < 1 Then
     VCTRL% = 1
     GoTo 90
   End If
   '
   If CACCAD% > 0 Then
     ' enviar un mail a flexoft
     DEXX$ = LOADFILE("IDENSYS.DRV")
     If Len(DEXX$) >= 96 Then
       DEYY$ = DESENCRI$(Left$(DEXX$, 94))
       EMPREAC$ = Trim$(Mid$(DEYY$, 13, 30))
     End If
     '
     ASUNTO$ = "PWX-" + Trim$(EMPREAC$) + "-" + MODULOX$ + "-" + Trim$(Str$(FLUTIL%)) + "/" + Trim$(Str$(CACCAD%))
     CLAX$ = RANDSTRING$(-6)
     TEXTOMAI$ = "CLACONTRA=" + CLAX$ + CONTRACLA$(CLAX$) + Chr$(124)
     TEXTOMAI$ = TEXTOMAI$ + "EMPRELICEN=" + Trim$(EMPREAC$) + Chr$(124)
     TEXTOMAI$ = TEXTOMAI$ + "MODULO=" + Trim$(MODULOX$) + Chr$(124)
     TEXTOMAI$ = TEXTOMAI$ + "VENCIMIENTO=" + Trim$(FETEX$(FLUTIL%)) + Chr$(124)
     TEXTOMAI$ = TEXTOMAI$ + "ACCESADIC=" + Trim$(Str$(CACCAD%)) + Chr$(124)
     TEXTOMAI$ = TEXTOMAI$ + "SERIVOL=" + Trim$(SEVOLF$) + Chr$(124)
     TEXTOMAI$ = TEXTOMAI$ + "ENVIADOEL=" + Trim$(FETEX$(HOII%)) + " - " + Time$ + Chr$(124)
     '
     MAI_DESTI$ = "FLEXOFT Soluciones"
     MAI_DIREL$ = "asistec@flexoft.com.ar"
     MAI_ASUNT$ = ASUNTO$
     MAI_TEXTO$ = ENCRINUM$(TEXTOMAI$)
     MAI_ADJUN$(0) = ""
     'Call GENERAMAIL
     VCTRL% = 1
     GoTo 90
     '
     If CACCAD% > 6 Then
       ' PRESENTAR MENSAJE RECORDATORIO DEL VENCIMIENTO
       If CACCAD% > 15 Then
         'SUSPENDER FUNCIONAMIENTO
         VCTRL% = 0
       End If
     End If
   End If
   '
90 VECONTROL% = VCTRL%
   '
End Function

Function HOY(HOS$)
    HO$ = Format$(Now, "dd/mm/yy")
    HOY = Val(Left$(HO$, 2)) + 32 * Val(Mid$(HO$, 4, 2)) + 32 * 13 * (100 + Val(Mid$(HO$, 7, 2)) - 80)
End Function
    '
Function ESUNDEMO%()
   '
   ESUNDEMO% = 0
   On Error Resume Next
   IKXX$ = FLEXLOG.Label99.Caption
   If Err Then Exit Function
   '
   If Mid$(IKXX$, 51, 9) <> "........." Then
     IKYY$ = Mid$(IKXX$, 51, 9)
     '
     ' DIGITO VERIFICADOR DEL SECTOR DE CONTROL
     DIVERI& = 0
     For KKI% = 1 To 8
       KKJ% = KKI%: If KKI% > 4 Then KKJ% = KKI% + 1
       DIVERI& = Val(Mid$(IKYY$, KKJ%, 1)) + (7 * KKJ% ^ 2)
     Next KKI%
     DIVERI& = DIVERI& Mod 10
     If Val(Mid$(IKYY$, 5, 1)) <> DIVERI& Then
       MsgBox "Módulo '" + App.EXEName + "' Indebidamente Alterado (D.Ver)"
       Close: End: Exit Function
     End If
     '
     ESUNDEMO% = 1
   End If
   '
End Function
