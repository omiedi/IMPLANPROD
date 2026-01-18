Attribute VB_Name = "CTRX_DRV1"
Public JOBID$
Public LUBAS$
Public LUDAT$
Public LUCOM$
Public ULODA$
Public ESUNDEMO%
Public FALLOCONEXION%

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
Function FECHATEX$(FEX)
    '
    Dim FS$(3), FEXX(3)
    FEXU = FEX
    '
10  If FEX = 0 Then
        F0$ = "        "
        GoTo 435
    End If
    '
    FEXX(3) = Int(FEX / 13 / 32) + 80
    FEX = FEX - 32 * 13 * (FEXX(3) - 80)
    FEXX(2) = Int(FEX / 32)
    FEXX(1) = FEX - 32 * FEXX(2)
    '
    While FEXX(3) > 99
       FEXX(3) = FEXX(3) - 100                                ' rutina a¤o 2000
    Wend
    '
    F0$ = ""
    For II% = 1 To 3
        F$ = Str$(FEXX(II%))
        F$ = Right$(F$, Len(F$) - 1)
        FS$(II%) = Right$("00" + F$, 2)
        F0$ = F0$ + FS$(II%)
        If II% < 3 Then F0$ = F0$ + "/"
    Next II%
    '
    SGL$ = "20": If Val(Mid$(F0$, 7, 2)) >= 80 Then SGL$ = "19"
    On Error Resume Next
    CDATX = CDate(Left$(F0$, 6) + SGL$ + Mid$(F0$, 7))
    If Err Then
       On Error GoTo 0
       FEXU = FEXU - 1
       FEX = FEXU
       GoTo 10
    End If
    On Error GoTo 0
    '
435 FECHATEX$ = F0$
    FEX = FEXU
    '
End Function
'
Function FECHANUM(FETX$)
    '
    Dim FE1(3)
    AX% = 0
    FETY$ = REPLACAR$(FETX$, ".", "/")
    FETX$ = REPLACAR$(FETY$, "-", "/")
    If Mid$(FETX$, 3, 1) <> "/" Then FETX$ = Left$(FETX$, 2) + "/" + Mid$(FETX$, 3)
    If Mid$(FETX$, 6, 1) <> "/" Then FETX$ = Left$(FETX$, 5) + "/" + Mid$(FETX$, 6)
    '
    If ((Len(FETX$) <> 8) Or (Mid$(FETX$, 3, 1) <> "/") Or (Mid$(FETX$, 6, 1) <> "/")) Then
        AX% = 1: GoTo 610
    End If
    '
    For II% = 1 To 3
        FE1(II%) = Val(Mid$(FETX$, 3 * II% - 2, 2))
        If II% <> 3 Then If FE1(II%) < 1 Then GoTo 600
    Next II%
    '
    If FE1(3) < 60 Then FE1(3) = 100 + FE1(3)                 ' rutina a¤o 2000
    '
    If ((FE1(1) > 31) Or (FE1(2) > 12)) Then GoTo 600
    FE = FE1(1) + 32 * FE1(2) + 32 * 13 * (FE1(3) - 80)
    GoTo 610

600 AX% = 2
610 If AX% > 0 Then FE = 0
    If FE < (-32767) Then FE = (-32767)
    If FE > 32767 Then FE = 32767
    '
    FECHANUM = FE
    '
End Function
'
Function HORATEX$(XL%)
    '
    HORA% = Int(XL% / 100)
    MINU% = XL% Mod 100
    If HORA% < 0 Or HORA% > 24 Or MINU% < 0 Or MINU% > 59 Or XL% > 2400 Then
        HORA% = 0: MINU% = 0
    End If
    '
    HORAS$ = Mid$(Str$(HORA%), 2)
    While Len(HORAS$) < 2: HORAS$ = " " + HORAS$: Wend
    MINUS$ = Mid$(Str$(MINU%), 2)
    While Len(MINUS$) < 2: MINUS$ = "0" + MINUS$: Wend
    '
    HORATEX$ = HORAS$ + ":" + MINUS$
    '
End Function
'
Function EMPRELI$()
    '
    Static EPLI$
    If EPLI$ = "" Then
       EPLI$ = "EMPRESA SIN LICENCIA"
       If EXISTE%("IDENSYS.DRV") > 0 Then
            Dim IDS As String * 96
            N1% = FILEOPEN%("IDENSYS.DRV", 9, 96)
            If LOF(N1%) > 0 Then
                Get #N1%, 1, IDS$
                DEXX$ = DESENCRI$(Left$(IDS$, 94))
                Get #N1%, 2, IDS$
                DF$ = DESENCRI$(Left$(IDS$, 94))
            End If
            Close #N1%
       End If
       EPLI$ = Trim$(UCase$(Mid$(DEXX$, 13, 30)))
       '
       If EPLI$ = "" Then
         CD$ = Left$(App.Path, 2)
         SEVOLF$ = Left$(SERIVOL$(CD$) + Space$(12), 12)
         If Trim$(SEVOLF$) = "CC34-EAB0" Then
              EPLI$ = "EYM DE MARCELO ANDRADE"
            Else
              EPLI$ = "VOLUMEN " + Trim$(SEVOLF$)
         End If
       End If
    End If
    '
    EMPRELI$ = EPLI$
    '
End Function

Function DIENTRE%(A, B)
    '
    Dim DIAS%(16)
    A1 = A: B1 = B
    '
    DIAS%(1) = 31
    DIAS%(2) = 28
    DIAS%(3) = 31
    DIAS%(4) = 30
    DIAS%(5) = 31
    DIAS%(6) = 30
    DIAS%(7) = 31
    DIAS%(8) = 31
    DIAS%(9) = 30
    DIAS%(10) = 31
    DIAS%(11) = 30
    DIAS%(12) = 31
    '
    If A1 > B1 Then
        DIENTRE% = 0
        Exit Function
    End If

    F1$ = FECHATEX$(A1)
    F2$ = FECHATEX$(B1)

    X% = Val(Left$(F2$, 2)) - Val(Left$(F1$, 2))
    Y2% = Val(Mid$(F2$, 4, 2))
    Y3% = Val(Mid$(F2$, 7, 2))
        If Y3% < 50 Then Y3% = 100 + Y3%
    FIN& = 12 * (Y3% - 80) + Y2%
    '
    X2% = Val(Mid$(F1$, 4, 2))
    X3% = Val(Mid$(F1$, 7, 2))
        If X3% < 50 Then X3% = 100 + X3%

    While 12 * (X3% - 80) + X2% < FIN&
        DIASI% = DIAS%(X2%)
        If (X3% Mod 4) = 0 Then
          If X3% <> 0 Then
            If X2% = 2 Then
                DIASI% = DIASI% + 1
            End If
          End If
        End If

        X% = X% + DIASI%
        X2% = X2% + 1
        If X2% > 12 Then
            X2% = 1
            X3% = X3% + 1
        End If
    Wend
    '
    DIENTRE% = X%
    '
End Function
'
Function DERACCE%(MDX$, APL$)
    '
    ' MDX$ = MODULO INVOCADO
    ' APX$ = DESCRIPCION DEL MÓDULO O APLICACION
    '
    Static CTRX%, CAPLI$(), DAPLI$()
    '
    DRXXX% = 0
    MDY$ = Trim$(UCase$(MDX$))
      If MDY$ = "" Then GoSub CARGALISDER: DRXXX% = 2: GoTo 99
    APM$ = Trim$(UCase$(APL$))
    QWR% = USERNU% Mod 100
      If QWR% = 1 Then GoSub CARGALISDER: DRXXX% = 2: GoTo 99  ' supervisor entra directo
    If Trim$(UCase$(MDX$)) = "FLSETUP" Then
      GoSub CARGALISDER: DRXXX% = 2: GoTo 99  ' PARA PODER CAMBIAR CLAVE PERSONAL
    End If
    '
    ATRICON$ = ""
    PPXX% = InStr(MDY$, "/")
    If PPXX% > 1 Then
       ATRICON$ = Mid$(MDY$, PPXX% + 1)
       MDY$ = Left$(MDY$, PPXX% - 1)
    End If
    '
    PPXX% = InStr(MDY$, "(")
    If PPXX% > 1 Then
       MDY$ = Left$(MDY$, PPXX% - 1)
    End If
    '
    PPY$ = Left$(MDY$, 1)
    If Asc(PPY$) < 48 Or (Asc(PPY$) > 57 And Asc(PPY$) < 65) Or Asc(PPY$) > 90 Then
        MDY$ = Mid$(MDY$, 2)
    End If
    '
    MDY$ = Trim$(MDY$)
    If MDY$ = "" Then GoSub CARGALISDER: DRXXX% = 2:               GoTo 99
    '
    ' BUSCAR EL DERECHO EN MDB
    MDY32$ = Left$(MDY$, 32)
    DRXXX% = 0
    Call ABRETACON
    SQLX$ = "SELECT * FROM DER_ACCE "
    SQLX$ = SQLX$ + " WHERE NUM_USER = " & QWR%
    SQLX$ = SQLX$ + " AND COD_APLI = '" & MDY32$ & "'"
    '
    Set PARATEMP = BaDaConsul.OpenRecordset(SQLX$, dbOpenSnapshot)
    With PARATEMP
      On Error Resume Next
      .MoveFirst
      If Err Then
          On Error GoTo 0
        Else
          On Error GoTo 0
          If Val(!DER_READ) > 0 Then DRXXX% = 1
          If Val(!DER_WRITE) > 0 Then DRXXX% = 2
      End If
    End With
25  PARATEMP.Close
    '
    If DRXXX% < 2 Then
20     If Left$(ATRICON$, 6) <> "NOMESS" Then  ' NO PRESENTA MENSAJE
          MENSA$ = "Usuario no Habilitado para Ejecutar esta Opción." + vbCrLf + vbCrLf
          MENSA$ = MENSA$ + "Modulo: " + Trim$(UCase$(MDY$)) + " - " + Trim$(APM$) + vbCrLf + vbCrLf
          MENSA$ = MENSA$ + "Solicite la Autorizacion Correspondiente al Supervisor." + vbCrLf
          MsgBox MENSA$, vbExclamation
       End If
       '
    End If
    '
99  DERACCE% = DRXXX%
    '
Exit Function
'
CARGALISDER:
    '
    If CTRX% = 0 Then
       ' Carga de MDB
       Call ABRETACON
       SQLX$ = "SELECT * FROM INDI_APLI "
       Set PARATEMP = BaDaConsul.OpenRecordset(SQLX$, dbOpenSnapshot)
       '
       JJ& = 0
       ReDim CAPLI$(1), DAPLI$(1)
       With PARATEMP
         Do While Not .EOF
           CAPLI1$ = Trim$(!COD_APLI)
           DAPLI1$ = Trim$(!DES_APLI)
           JJ& = JJ& + 1
           If JJ& >= UBound(CAPLI$) Then
             ReDim Preserve CAPLI$(JJ& + 1), DAPLI$(JJ& + 1)
           End If
           CAPLI$(JJ&) = CAPLI1$
           DAPLI$(JJ&) = DAPLI1$
         .MoveNext
         Loop
       End With
       PARATEMP.Close
       '
    End If
    '
    '
    ' AGREGA A LA LISTA DE DERECHOS DEFINIDOS
    If APM$ <> "" Then
       CANDE& = UBound(CAPLI$)
       For U& = 1 To CANDE&
         If CAPLI$(U&) = MDY$ Then GoTo 120
       Next U&
       '
       ' AGREGA AL VECTOR
       MDY32$ = Left$(MDY$, 32)
       JJ& = JJ& + 1
       If JJ& >= UBound(CAPLI$) Then
         ReDim Preserve CAPLI$(JJ& + 1), DAPLI$(JJ& + 1)
       End If
       CAPLI$(JJ&) = MDY32$
       DAPLI$(JJ&) = APM$
       '
       ' AGREGA A LA TABLA MDB
       Call ABRETACON
       SQLX$ = "SELECT * FROM INDI_APLI WHERE COD_APLI = '" & MDY32$ & "'"
       Set PARATEMP = BaDaConsul.OpenRecordset(SQLX$, dbOpenDynaset)
       With PARATEMP
         On Error Resume Next
         .MoveFirst
         If Err Then
             On Error GoTo 0
             .AddNew
           Else
             On Error GoTo 0
             .Edit
         End If
         !COD_APLI = MDY32$
         !DES_APLI = Left$(APM$, 128)
         .Update
       End With
       PARATEMP.Close
       '
    End If
    '
120 Return
    '
    '
End Function
'
