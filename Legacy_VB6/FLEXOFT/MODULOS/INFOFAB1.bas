Attribute VB_Name = "INFOFAB1"
Dim HO As Integer
'


Sub INFADIA()
     '
     HO% = HOY(HOS$)
     LU$ = LUDAT$
     '
     TITU$ = "INFORMES DE PRODUCCION - GENERALES"
     '
100  VDEF$ = MKI$(HO%) + Space$(6)
101  DIAEQUI$ = OBJPLA$("DIAEQUI", VDEF$)
     If DIAEQUI$ = "" Then Exit Sub
     '
     DIA% = CVI(Left$(DIAEQUI$, 2))
     If DIA% > HO% Then
         Call MENSERR(24, "Fecha Incorrecta")
         GoTo 100
     End If
     '
     MAQUI$ = Mid$(DIAEQUI$, 3, 6)
     If TRIM$(ECOARCH("PADMAQ", MAQUI$)) = "" Then
         Call MENSERR(24, "Equipo no Definido")
         GoTo 100
     End If
     '
     Screen.MousePointer = 11
     RGFSX$ = RECFILT$("OPERFAP", 7, MAQUI$)
     If Len(RGFSX$) < 4 Then
         Call MENSERR(24, "No hay Operaciones Programadas\para el Equipo Elegido.")
         Screen.MousePointer = 1
         GoTo 100
     End If
     '
     J% = 0
     MENSAIN$ = "Inconsistencia en Base de Datos de Ope-\raciones de Fabricación.\  \Los Punteros de  Operaciones Pendientes\no se Corresponden con los Registros de\Operaciones Programadas.\  \Consulte con su Soporte de Sistemas."
     Call FRESHECHO("!OPERINF")
     For K% = 1 To Len(RGFSX$) Step 4
        REGI& = CVS(Mid$(RGFSX$, K%, 4))
        XX$ = REGLEIDO$("OPERFAP", REGI&)
        DES% = CVI(Mid$(XX$, 135, 2))
        HAS% = CVI(Mid$(XX$, 139, 2))
        If DIA% >= DES% Then
           If DIA% <= HAS% Then
              CAPROD# = CVD(Mid$(XX$, 79, 8))
              CAREA# = CVD(Mid$(XX$, 87, 8))
              SALDO# = CAPROD# - CAREA#: If SALDO# < 0 Then SALDO# = 0
              Call REPLA(XX$, MKD$(SALDO#), 95, 8)
              Call GRAREG("OPERFAP", XX$, REGI&)
              '
              INTEN% = 0
110           REGO& = CVS(Mid$(XX$, 221, 4))
              If REGO& < 1 Or REGO& > ULTREG&("OPERFAB") Then
                 Call MENSERR(24, MENSAIN$)
                 Screen.MousePointer = 1
                 Exit Sub
              End If
              '
              ZZ$ = REGLEIDO$("OPERFAB", REGO&)
              If Left$(ZZ$, 86) <> Left$(XX$, 86) Then
If HOY(HOS$) > FECHANUM("10/06/01") Then GoTo 111
                If INTEN% = 0 Then
                  Screen.MousePointer = 11
                  For KKL& = 1 To ULTREG&("OPERFAB")
                    If Left$(REGLEIDO$("OPERFAB", KKL&), 86) = Left$(XX$, 86) Then
                       XX$ = REGLEIDO$("OPERFAP", REGI&)
                       Call REPLA(XX$, MKS$(KKL&), 221, 4)
                       Call GRAREG("OPERFAP", XX$, REGI&)
                       Screen.MousePointer = 1
                       INTEN% = INTEN% + 1
                       GoTo 110
                    End If
                  Next KKL&
                  Screen.MousePointer = 1
                End If
111             Call MENSERR(24, MENSAIN$)
                Screen.MousePointer = 1
                Exit Sub
              End If
              '
              If SALDO# > 0.05 Then
                 NROOF$ = Left$(XX$, 12)
                 CI% = CVI(Mid$(XX$, 13, 2))
                 NOPE% = CVI(Mid$(XX$, 15, 2))
                 DOPE$ = Mid$(XX$, 17, 48)
                 DESTI$ = Mid$(XX$, 65, 8)
                 CAINF# = SALDO#
                 FEOP% = CVI(Mid$(XX$, 139, 2))
                 HOOP% = CVI(Mid$(XX$, 141, 2))
                 MALIB$ = "S"
                 TITOT = CVS(Mid$(XX$, 131, 4))
                 '
                 J% = J% + 1
                 YY$ = MKI$(J%) + Space$(58) + MKS$(REGI&)
                 NOPX$ = AJUSTI$(NROOF$, 13) + AJUSTI$(DESTI$, 9)
                 NOPX$ = NOPX$ + AJUSTD$(Str$(NOPE%), 5) + " - " + AJUSTI$(DOPE$, 38)
                 NOPX$ = Left$(NOPX$, 46)
                 NOPX$ = NOPX$ + NUMSTRI$(SALDO#, 12, 1, 0)
                 Call REPLA(YY$, NOPX$, 3, 58)
                 JJ& = J%
                 Call GRAREG("!OPERINF", YY$, JJ&)
                 '
              End If
           End If
        End If
     Next K%
     '
     If J% < 1 Then
         Call MENSERR(24, "No hay Operaciones Programadas\para esa Fecha y ese Equipo.")
         Screen.MousePointer = 1
         GoTo 100
     End If
     '
     Call SETULTREG("!OPERINF", JJ&)
     Call CIERRARCH("!OPERINF")
     '
     Screen.MousePointer = 1
     '
200  FEX = DIA%
     Call SELECHO("!OPERINF/Operaciones Programadas para el " + FECHATEX$(FEX) + " - " + TRIM$(MAQUI$))
     NORDE& = Val(VALACT1$("!OPERINF"))
     If NORDE& < 1 Then
         GoTo 100
     End If
     '
     XX$ = REGLEIDO$("!OPERINF", NORDE&)
     REGI& = CVS(Mid$(XX$, 61, 4))
     Call CARGAINF(REGI&)
     If REGI& < 1 Then
          GoTo 200
        Else
          VDEF$ = DIAEQUI$
          GoTo 101
     End If
     '
End Sub
'
Sub INFOPER()      ' INFORMES DE OPERACION POR O/F
    '
    HO% = HOY(HOS$)
    LU$ = LUDAT$
    '
    TITU$ = "INFORMES DE OPERACION POR O/F"

100  NROOF$ = ""

     Call FRESHECHO("ECORDEF")
     If ULTREG&("ECORDEF") < 1 Then
         Call MENSERR(24, "No hay Ordenes de Fabricación en Curso.")
         Exit Sub
     End If
     '
1405 Call SELECHO("ECORDEF/Ordenes de Trabajo en Curso")
     Call FRESHALL
     '
1410 NROOF$ = AJUSTI$(VALACT1$("ECORDEF"), 12)
     If TRIM$(NROOF$) = "" Then
         Exit Sub
     End If
     '
     BVOF% = BINVAL%(AJUSTI$(NROOF$, 12))
     RFF$ = RECFILT$("OPERFAB", 6, MKI$(BVOF%))
     JJ& = 0
     For KK% = 1 To Len(RFF$) Step 4
       REOPER& = CVS(Mid$(RFF$, KK%, 4))
       ZZ$ = REGLEIDO$("OPERFAB", REOPER&)
       NORFA$ = Left$(ZZ$, 12)
       If NORFA$ = NROOF$ Then
         CAREA# = CVD(Mid$(ZZ$, 87, 8))
         If Mid$(ZZ$, 177, 6) <> String$(6, 0) Then ' solo las ya informadas
            MAQUI$ = Mid$(ZZ$, 151, 6): If TRIM$(MAQUI$) = "" Then MAQUI$ = Mid$(ZZ$, 73, 6)
            FERE% = CVI(Mid$(ZZ$, 157, 2))
            HORE% = CVI(Mid$(ZZ$, 159, 2))
            LEGOPER$ = Mid$(ZZ$, 165, 6)
            LEGA4% = CVI(Mid$(ZZ$, 171, 2))
            If LEGA4% > 0 Then
              Call MENSERR(24, "Informe no Editable")
              GoTo 1405
            End If
            TITOT = CVS(Mid$(ZZ$, 161, 4))
            HORAS% = Int(TITOT / 60)
            MINUS% = Int(TITOT - 60 * HORAS%)
            SEGUS% = 60 * (TITOT - 60 * HORAS% - MINUS%)
            While SEGUS% > 59: SEGUS% = SEGUS% - 60: MINUS% = MINUS% + 1: Wend
            While MINUS% > 59: MINUS% = MINUS% - 60: HORAS% = HORAS% + 1: Wend
            '
            XX$ = REGBLAN$("OPERFAIN")
            Call REPLA(XX$, Left$(ZZ$, 64), 1, 64)
            Call REPLA(XX$, MKD$(CAREA#), 65, 8)
            Call REPLA(XX$, MAQUI$, 73, 6)
            Call REPLA(XX$, MKI$(FERE%), 79, 2)
            Call REPLA(XX$, MKI$(HORE%), 81, 2)
            Call REPLA(XX$, MKI$(HORAS%) + Chr$(MINUS%) + Chr$(SEGUS%), 84, 4)
            Call REPLA(XX$, LEGOPER$, 118, 6)
            Call REPLA(XX$, MKS$(REOPER&), 125, 4)
            '
            JJ& = JJ& + 1
            Call GRAREG("!OPERFAIN", XX$, JJ&)
         End If
       End If
     Next KK%
     Call SETULTREG("!OPERFAIN", JJ&)
     '
     If JJ& < 1 Then
        Call MENSERR(24, "No Hay Operaciones Informadas\para la " + NROOF$)
        GoTo 1405
     End If
     '
     TITUPA$ = VALACT2$("ECORDEF")
     PPXX% = InStr(TITUPA$, " - ")
     If PPXX% < 1 Then PPXX% = 17
     CODPIE$ = TRIM$(Left$(TITUPA$, PPXX% - 1))
     CIPIE% = CODINT%(CODPIE$)
     If CIPIE% < 1 Then
        Call MENSERR(24, "Imposible Determinar Pieza de esta O/F")
     End If
     '
   ' MODIFICAR PARA PODER SOLO CAMBIAR LA CANTIDAD Y EL TIEMPO EMPLEADO
   ' AL GRABAR QUE RECALCULE EL TIEMPO STANDARD
     '
     VTB% = VENTABU%("INFOPER/TITUPAN=Informe de Produccion: " + TITUPA$)
     '
     If VTB% >= 0 Then
        '
        While Len(RFF$) < 4 * ULTREG&("!OPERFAIN")
           RFF$ = RFF$ + Chr$(0)
        Wend
        '
        For U& = 1 To ULTREG&("!OPERFAIN")
           '
           XX$ = REGLEIDO$("!OPERFAIN", U&)
           HMS$ = Mid$(XX$, 84, 4)
           TITOT = 60 * CVI(Left$(HMS$, 2)) + Asc(Mid$(HMS$, 3, 1)) + Asc(Mid$(HMS$, 4, 1)) / 60
           Call REPLA(XX$, MKS$(TITOT), 84, 4)
           Call REPLA(XX$, NROOF$, 1, 12)
           Call REPLA(XX$, MKI$(CIPIE%), 13, 2)
           '
           REOPER& = CVS(Mid$(XX$, 125, 4))
           '
           If REOPER& > 0 And REOPER& <= ULTREG&("OPERFAB") Then
               ZZ$ = REGLEIDO$("OPERFAB", REOPER&)
               NORFA$ = Left$(ZZ$, 12)
               If NORFA$ <> NROOF$ Then
                   Call CIERRARCH("*.*")
                   Call MENSERR(24, "Imposible Grabar")
                   Call FINAL("")
               End If
               CAREA# = CVD(Mid$(XX$, 65, 8))
               Call REPLA(ZZ$, MKD$(CAREA#), 87, 8)
               MAQUI$ = Mid$(XX$, 73, 6)
               Call REPLA(ZZ$, MAQUI$, 151, 6)
               FERE% = CVI(Mid$(XX$, 79, 2))
               Call REPLA(ZZ$, MKI$(FERE%), 157, 2)
               HORE% = CVI(Mid$(XX$, 81, 2))
               Call REPLA(ZZ$, MKI$(HORE%), 159, 2)
               LEGOPER$ = Mid$(XX$, 118, 6)
               Call REPLA(ZZ$, LEGOPER$, 165, 6)
               Call REPLA(ZZ$, MKS$(TITOT), 161, 4)
               '
               NOPE% = CVI(Mid$(ZZ$, 15, 2))
               SALRE# = CVD(Mid$(ZZ$, 79, 8)) - CAREA#
               If SALRE# < 0 Then SALRE# = 0
               NUOROP% = Asc(Mid$(ZZ$, 67, 1))
               TSTAN = TIESTAN(CIPIE%, NOPE%, CAREA#, NUOROP%)
               '
               Call REPLA(ZZ$, MKD$(SALRE#), 95, 8)
               Call REPLA(ZZ$, MKS$(TSTAN), 173, 4)
               '
               Call GRAREG("OPERFAB", ZZ$, REOPER&)
           End If
           '
        Next U&
     End If
     '
     Call CIERRARCH("OPERFAB")
     GoTo 1405
     '
End Sub
'
Sub CARGAINF(REGI&)
     '
     HOII% = HOY(HOS$)
     AGREGA% = 0
     '
     ZZ$ = REGLEIDO$("OPERFAB", REGI&)
     RANTEINF$ = ZZ$
     '
     NROOF$ = Left$(ZZ$, 12)
     CI% = CVI(Mid$(ZZ$, 13, 2))
     NOPE% = CVI(Mid$(ZZ$, 15, 2))
     DOPE$ = Mid$(ZZ$, 17, 48)
        CAPROD# = CVD(Mid$(ZZ$, 79, 8))
        CAREA# = CVD(Mid$(ZZ$, 87, 8))
        SALDO# = CAPROD# - CAREA#: If SALDO# < 0 Then SALDO# = 0
     CAINF# = SALDO#
     MAQUI$ = Mid$(ZZ$, 73, 6)
     FEOP% = CVI(Mid$(ZZ$, 139, 2))
     HOOP% = CVI(Mid$(ZZ$, 141, 2))
     If FEOP% < 1 Then
        FEOP% = HOII%
        HOOP% = HORANUM(Time$)
     End If
     MALIB$ = "S"
     TITOT = CVS(Mid$(ZZ$, 131, 4))
     HORAS% = Int(TITOT / 60)
     MINUS% = Int(TITOT - 60 * HORAS%)
     SEGUS% = 60 * (TITOT - 60 * HORAS% - MINUS%)
     While SEGUS% > 59: SEGUS% = SEGUS% - 60: MINUS% = MINUS% + 1: Wend
     While MINUS% > 59: MINUS% = MINUS% - 60: HORAS% = HORAS% + 1: Wend
     OBSE$ = Space$(30)
     '
     YY$ = REGBLAN$("INFOFAB")
     Call REPLA(YY$, NROOF$, 1, 12)
     Call REPLA(YY$, MKI$(CI%), 13, 2)
     Call REPLA(YY$, MKI$(NOPE%), 15, 2)
     Call REPLA(YY$, DOPE$, 17, 48)
     Call REPLA(YY$, MKD$(CAINF#), 65, 8)
     Call REPLA(YY$, MAQUI$, 73, 6)
     Call REPLA(YY$, MKI$(FEOP%), 79, 2)
     Call REPLA(YY$, MKI$(HOOP%), 81, 2)
     Call REPLA(YY$, MALIB$, 83, 1)
     Call REPLA(YY$, MKI$(HORAS%) + Chr$(MINUS%) + Chr$(SEGUS%), 84, 4)
     Call REPLA(YY$, OBSE$, 88, 30)
     Call REPLA(YY$, String$(8, 0), 118, 8)
     '
1420 XX$ = OBJPLA$("INFOFAB", YY$)
     If XX$ = "" Then
        REGI& = 0
        GoTo 1498
     End If
     '
     YY$ = XX$
     CAINF# = CVD(Mid$(XX$, 65, 8))
     '
     MAQUI$ = Mid$(XX$, 73, 6)
     If TRIM$(ECOARCH$("PADMAQ", MAQUI$)) = "" Then
         Call MENSERR(24, " Equipo Incorrecto.")
         GoTo 1420
     End If
     '
     FEOP% = CVI(Mid$(XX$, 79, 2))
     HOOP% = CVI(Mid$(XX$, 81, 2))
     If FEOP% > HOII% Then
           Call MENSERR(24, "Fecha Incorrecta.")
           GoTo 1420
         ElseIf FEOP% = HOII% Then
           If HOOP% > HORANUM(Time$) Then
             Call MENSERR(24, "Hora Incorrecta")
             GoTo 1420
           End If
     End If
     '
     If ULTREG&("PADRON") > 0 Then
        If CVI(Mid$(XX$, 118, 2)) < 1 Then
           Call MENSERR(24, "Falta Informar Operario(s)")
           GoTo 1420
        End If
        For KK% = 118 To 125 Step 2
           LEGA% = CVI(Mid$(XX$, KK%, 2))
           If LEGA% > 0 Then
              If ECOARCH$("PADRON", MKI$(LEGA%)) = "" Then
                 Call MENSERR(24, "Legajo Operario no Válido")
                 GoTo 1420
              End If
           End If
        Next KK%
     End If
     '
     MALIB$ = " "
     If CAINF# >= SALDO# - 0.05 Then
         MALIB$ = "S"
     End If
     Call REPLA(XX$, MALIB$, 83, 1)
     '
     TITOT = 60 * CVI(Mid$(XX$, 84, 2)) + Asc(Mid$(XX$, 86, 1)) + Asc(Mid$(XX$, 87, 1)) / 60
     Call REPLA(XX$, MKS$(TITOT), 84, 4)
     TITOT = CVS(Mid$(XX$, 84, 4))
     '
     Call BLOQARCH("OPERFAB")
     '
     FEBA% = HOII%
     HOBA% = HORANUM(Time$)
     USUA% = USNBR% Mod 100
     DBAJA$ = MKI$(FEBA%) + MKI$(HOBA%) + Chr$(USUA%) + Chr$(1) + "Parte Fabric."
     '
     NORFB$ = Left$(XX$, 12)
     CIIK% = CVI(Mid$(XX$, 13, 2))
     NOPE% = CVI(Mid$(XX$, 15, 2))
     CAREAL# = CVD(Mid$(XX$, 65, 8))
     MAQUI$ = Mid$(XX$, 73, 6)
     FERE% = CVI(Mid$(XX$, 79, 2))
     HORE% = CVI(Mid$(XX$, 81, 2))
     TEJEC = CVS(Mid$(XX$, 84, 4))
     Observa$ = Mid$(XX$, 88, 30)
     LEGOPER$ = Mid$(XX$, 118, 8)
     '
     ZZ$ = REGLEIDO$("OPERFAB", REGI&)
     NORFA$ = Left$(ZZ$, 12)
     If ZZ$ <> RANTEINF$ Or NORFA$ <> NORFB$ Or CVI(Mid$(ZZ$, 15, 2)) <> NOPE% Or Mid$(ZZ$, 177, 6) <> String$(6, 0) Then
        Call CIERRARCH("OPERFAB")
        Call MENSERR(24, "Imposible Grabar este Parte de Operación.\  \Causa Probable: Se Ha Registrado\un Parte para esta Operación\desde otra Terminal de la Red.\  \Vuelva a Ingresar el Parte de Operación.")
        GoTo 1499
     End If
     '
     SALRE# = CVD(Mid$(ZZ$, 79, 8)) - CAREAL#
     If SALRE# < 0 Then SALRE# = 0
     AAA = CVI(Mid$(ZZ$, 65, 2))
     '
     NUOROP% = Asc(Mid$(ZZ$, 67, 1))
     If NUOROP% > 31 Then NUOROP% = 0
     TSTAN = TIESTAN(CIIK%, NOPE%, CAREAL#, NUOROP%)
     '
     Call REPLA(ZZ$, MKD$(CAREAL#), 87, 8)
     Call REPLA(ZZ$, MKD$(SALRE#), 95, 8)
     Call REPLA(ZZ$, MAQUI$, 151, 6)
     Call REPLA(ZZ$, MKI$(FERE%), 157, 2)
     Call REPLA(ZZ$, MKI$(HORE%), 159, 2)
     Call REPLA(ZZ$, MKS$(TEJEC), 161, 4)
     Call REPLA(ZZ$, LEGOPER$, 165, 8)
     Call REPLA(ZZ$, MKS$(TSTAN), 173, 4)
     Call REPLA(ZZ$, DBAJA$, 177, 22)
     Call GRAREG("OPERFAB", ZZ$, REGI&)
     '
     If SALRE# > 0.01 Then
       '
       TTXX$ = "Queda un Saldo Pendiente\para esta Operación\\Mantener Pendiente\Cancelar Saldo"
       If COMALTER%(TTXX$) <> 2 Then
         NUOROP% = 1 + Asc(Mid$(ZZ$, 67, 1))
         If NUOROP% > 32 Then NUORORP% = 1
         CIIK% = CVI(Mid$(ZZ$, 13, 2))
         NOPE% = CVI(Mid$(ZZ$, 15, 2))
         DALTA$ = Chr$(NUOROP%) + Mid$(DBAJA$, 2, 5)
         INFOBLA$ = Space$(6) + String$(26, 0) + Space$(16)
         PROGBLA$ = Space$(6) + String$(26, 0)
         MINVARI = TIESTAN(CIIK%, NOPE%, SALRE#, NUOROP%)
         '
         Call REPLA(ZZ$, DALTA$, 67, 6)
         Call REPLA(ZZ$, MKD$(SALRE#), 79, 8)
         Call REPLA(ZZ$, MKD$(0), 87, 8)
         Call REPLA(ZZ$, MKD$(SALRE#), 95, 8)
         Call REPLA(ZZ$, String$(32, 0), 103, 32)
         Call REPLA(ZZ$, MKS$(MINVARI), 111, 4)
         Call REPLA(ZZ$, MKS$(MINVARI), 131, 4)
         Call REPLA(ZZ$, String$(16, 0), 135, 16)
         Call REPLA(ZZ$, INFOBLA$, 151, 48)
         Call REPLA(ZZ$, PROGBLA$, 225, 32)
         Call REGAPP("OPERFAB", ZZ$)
         AGREGA% = 1
       End If
     End If
     '
1498 Call CIERRARCH("OPERFAB")
     '
1499 Call CIERRARCH("*.*")
     Call BAJALDISCO
     '
End Sub
'

Sub GESALPROP(MODO%)
   '
   Dim NORFAX$(4096)
   NROOFA$ = "": NOPEA% = 0
   '
   FIN& = ULTREG&("OPERFAP")
   FIN1& = ULTREG&("OPERFAB")
   W& = 0
   For U& = 1 To FIN&
      XX$ = REGLEIDO$("OPERFAP", U&)
      CAPROD# = CVD(Mid$(XX$, 79, 8))
      CAREA# = CVD(Mid$(XX$, 87, 8))
      SALDO# = CAPROD# - CAREA#
      If SALDO# >= 0.05 Then
         NROOF$ = Left$(XX$, 12)
         CI% = CVI(Mid$(XX$, 13, 2))
         NOPE% = CVI(Mid$(XX$, 15, 2))
         DOPE$ = Mid$(XX$, 17, 48)
         '
         For KX& = 1 To W&
            If NORFAX$(KX&) = NROOF$ Then GoTo 10
         Next KX&
         '
         Z$ = REGBLAN$("SPPROCE")
         Call REPLA(Z$, MKI$(CI%), 1, 2)
         Call REPLA(Z$, NROOF$, 3, 12)
         Call REPLA(Z$, MKI$(CI%), 15, 2)
         Call REPLA(Z$, MKI$(NOPE%), 17, 2)
         Call REPLA(Z$, DOPE$, 19, 48)
         Call REPLA(Z$, MKD$(CAPROD#), 67, 8)
         W& = W& + 1
         Call GRAREG("SPPROCE", Z$, W&)
         NORFAX$(W&) = NROOF$: NOPEA% = NOPE%
         GoTo 19
         '
10       Z$ = REGLEIDO$("SPPROCE", KX&)
         NOPEA% = CVI(Mid$(Z$, 17, 2))
         If NOPE% < NOPEA% Then
            Call REPLA(Z$, MKI$(NOPE%), 17, 2)
            Call REPLA(Z$, DOPE$, 19, 48)
            Call GRAREG("SPPROCE", Z$, KX&)
         End If
      End If
19 Next U&
   '
   Call SETULTREG("SPPROCE", W&)
   Call CIERRARCH("SPPROCE")
   '
   If MODO% = 1 Then
         Call CONECRUN("*SOSPPR1", "Saldos en Proceso por O.F")
       ElseIf MODO% = 2 Then
         Call CONECRUN("*SOSPPR2", "Saldos en Proceso por Producto")
   End If
   '
End Sub
