Attribute VB_Name = "ADMIORF1"
Public CHECOMPO%
Public PRECOLAN%
Public PRECOCIE%
Public FRALOMAX%
Public VALINDOP%
Public INCLUPRI%
Dim APROLAN%
'

Sub ORFNUE(REPIM%)                           ' NUEVA ORDEN DE FABRICACION
    '
    ' REPIM%=1 ==> nueva orden de fabricacion
    ' REPIM%=2 ==> repite impresion de O/FAB
    '
    Static YADEPU%
    HO% = HOY(HOS$)
    LU$ = LUDAT$
    Dim CILOTE%(256), CALOTE#(256), CALOTE1#(256), PUNMAT$(256), PUNOPE$(256)
    Dim CISUMI%(512), CASUMI#(512)
    '
    CANCELPRINT% = 0
    If CONTROL$("ORDEFABR", "SUPRINT") = "SI" Then
       CANCELPRINT% = 1
    End If
    '
    GoSub FORIPRE                      ' formulario de impresion
    GoSub MONUME                       ' modo de numeracion
    GoSub GESUBORDEN                   ' genera sub-ordenes
    GoSub OBSERVACI                    ' incluye observaciones
    '
100 If REPIM% = 2 Then
       Call SELECHO("ORFANUE/Ordenes de Trabajo para Re-Impresión")
       NRO& = Val(VALACT1$("ORFANUE"))
       If NRO& < 1 Then
          INDI% = 0
          Call CIERRARCH("*.*")
          Call LIBARCH("*.*")
          Exit Sub
       End If
    End If
    '
    If MODORFA% = 0 Then
       GoSub MODOGEN                      ' modo directo / MRP / Pedidos
    End If
    '
1000:
    Call CIERRARCH("*.*")
    Call LIBARCH("RESERVA")
    Call FRESHECHO("ORFAEMI")
    '
    If REPIM% = 2 Then
       For U& = ULTREG&("ORFANUE") To 1 Step -1
          X3$ = REGLEIDO$("ORFANUE", U&)
          If CVS(Left$(X3$, 4)) = NRO& Then
             X1$ = X3$: X2$ = X3$: TORFANUE$ = X3$
             CIOF% = CVI(Mid$(X3$, 71, 2)): CI% = CIOF%
             CAN# = CVD(Mid$(X3$, 89, 8))
             LOTE# = CDbl(CVS(Mid$(X3$, 97, 4)))
             FGEN = CVI(Mid$(X3$, 69, 2))
             FLAN% = CVI(Mid$(X3$, 113, 2))
             SPY$ = TRIM$(CSTRING$(Mid$(X3$, 115, 4), 3, 7, 3, 2))
             If Val(SPY$) < 1 Then
                SPY$ = TRIM$(Mid$(X3$, 101, 8))
             End If
             DESTINO$ = SPY$
             DESTI$ = SPY$
             If Asc(Mid$(X3$, 112, 1)) > 1 Then
               Call MENSERR(24, "Imposible Regenerar.\O/F Cumplida, Suspendida o Anulada.")
               GoTo 100
             End If
             GoTo 103
          End If
       Next U&
       Call MENSERR(24, "Orden no Encontrada")
       GoTo 100
    End If
    '
    GoSub ARMAPRIMERA
    X1$ = OBJPLA$(OBJEPRI$, X0$)
    If X1$ = "" Then GoTo 999
    '

    GoSub VERICODI
    If CIOF% < 1 Then GoTo 100
    CIXI% = CIOF%
    SPX$ = String$(4, 0): SPY$ = ""
    If ULTREG&("SECPROD") > 0 Then
101   OBX$ = OBJPLA$("ENTRASECPROD", MKI$(CIXI%) + MKS$(SEPRODUC(CIXI%)))
      If OBX$ = "" Then GoTo 999
      SPX$ = Mid$(OBX$, 3, 4)
      If EXISEPRO% = 1 Then
        If ECOARCH$("SECPROD", SPX$) = "" Then GoTo 101
      End If
      SPY$ = ""
      If CVS(SPX$) > 0 Then
        SPY$ = TRIM$(CSTRING$(SPX$, 3, 8, 3, 2))
      End If
    End If
    '
    GoSub AJUCANORF
    DISTA% = 0
    If EXISTE%(LUDAT$ + "DISTALLE.DAT") > 0 Then
      Call SETULTREG("!DISTALLE", 0)
      RFF$ = RECFILT$("DISTALLE", 1, MKI$(CIOF%))
      If RFF$ <> "" Then
        DISTA% = 1
        Call FILTERFILE("DISTALLE", "", 1, MKI$(CIOF%))
1011        VTB% = VENTABU%("DISTALLE")
        If VTB% < 0 Then GoTo 999
        SUCADIS& = 0
        For KU& = 1 To ULTREG&("!DISTALLE")
          XXZ$ = REGLEIDO$("!DISTALLE", KU&)
          SUCADIS& = SUCADIS& + CVS(Mid$(XXZ$, 9, 4))
        Next KU&
        If SUCADIS& < 1 Then
          Call MENSERR(24, "Falta Ingresar Cantidades")
          GoTo 1011
        End If
        CANORF = SUCADIS&
      End If
    End If
    '
    GoSub ARMASEGUNDA
    '
    '
102 X2$ = OBJPLA$("ORFANUE-2", X1$)
    If X2$ = "" Then GoTo 100
    '
    GoSub VALIDA1                   ' validacion de entrada
    If VALIDA% < 0 Then
         GoTo 100
       ElseIf VALIDA% = 0 Then
         GoTo 102
    End If
    '
    If DISTA% > 0 Then
      If Int(CANORF + 0.5) <> SUCADIS& Then
        Call MENSERR(24, "No Coinciden Cantidades Distribuidas")
        GoTo 1011
      End If
    End If
    '
    If VERITAND% = 1 Then
       If LOECON >= 1 Then
          If CANORF / LOECON <> Int(CANORF / LOECON) Then
             Call MENSERR(24, "La Cantidad Ordenada no Resulta\un Múltiplo Entero de Lote Standard (Tanda).\  \Ajustada a un Múltiplo de Tanda.")
             X1$ = X2$
             CANORF = LOECON * Int(CANORF / LOECON + 0.5)
             Call REPLA(X1$, MKD$(CDbl(CANORF)), 89, 8)
             GoTo 102
          End If
       End If
    End If
    '
103 XACCE$ = ""
    If InStr(UCase$(EMPREAC$), "BORENS") > 0 Then
       GoSub BORENACCE
       If XACCE$ = "" Then GoTo 100
       GoSub BORENCOMPLE
       If VTB% < 0 Then GoTo 100
    End If
    '
    GoSub OBSERVOF                  ' ingresa observaciones
    If VALIDA% < 0 Then
       GoTo 100
    End If
    '
    Call CIERRARCH("ORFANUE")
    Call CIERRARCH("*.*")
    '
    If DISTA% = 1 Then
      LOTEA# = LOTE#
      LOTE# = CAN#
      GoSub GELOTES
      LOTE# = LOTEA#
      GoTo 104
    End If
    '
    GoSub GELOTES       ' genera subordenes por lotes
    GoSub GESUBOR       ' genera subordenes por subconjuntos
    '
104 For II% = 1 To NORDE%
       SECUOPSTA$ = RECFILT$("SECOPER", 1, MKI$(CILOTE%(II%)))
       If Len(SECUOPSTA$) < 4 Then
          CMAT% = COMALTER%("Falta Definir Operaciones para\" + TRIM$(CODEXT$(CILOTE%(II%))) + " - " + TRIM$(DESCRIT0$(CILOTE%(II%))) + "\\Cancelar\Continuar")
          If CMAT% <> 2 Then GoTo 102
       End If
    Next II%
    '
    GoSub GERESMAT      ' genera reserva de componentes
       If VACONTROL% < 0 Then
         If REPIM% = 2 Then
            GoTo 100
          Else
            GoTo 102
         End If
       End If
       '
    Call CIERRARCH("ORFANUE")
    Call CIERRARCH("*.*")
    '
    Call BLOQARCH("ORDEFABR")
    Call BLOQARCH("INDIORF")
    Call BLOQARCH("RESERVA")
    Call BLOQARCH("OPERFAB")
    Call BLOQARCH("OPERFAP")
    Call BLOQARCH("ECORDEP")
    Call BLOQARCH("ECORDEF")
    Call BLOQARCH("ORFAEMI")
    Call BLOQARCH("ORFANUE")
    '
    TERENUME$ = ""
    If REPIM% <> 2 Then
       '
       If CONTROL$("", "NUMEORFA") = "MANUAL" Then
         RFFX$ = RECFILT$("ORFANUE", 1, MKS$(NRO&))
         If RFFX$ <> "" Then
           Call MENSERR(24, "El Numero " + TRIM$(Str$(NRO&)) + " ya Está Utilizado")
           GoTo 105
         End If
         GoTo 107
       End If
       '
105    REGXX& = ULTREG&("ORFANUE")
       If REGXX& < 1 Then
            NROPROX& = 1
          Else
            NROPROX& = 1 + CVS(Left$(REGLEIDO$("ORFANUE", REGXX&), 4))
       End If
       '
       If NROPROX& <> NRO& Then
          TERENUME$ = "FLEXOFT ha Re-Numerado la O/F\como Orden Número '" + TRIM$(Str$(NROPROX&)) + "'."
          NRO& = NROPROX&
       End If
       '
107    GoSub GRAORFNUE
       Call CIERRARCH("ORFANUE")
       GoSub GRAORDEFA
       Call CIERRARCH("ORDEFABR")
       '\\\OT-06-0209-WAR-07/06/06///
       'RUTINA DE ANOTACION
       TTXYZ$ = Space$(76)
       HOII% = HOY(HOS$)
       Call REPLA(TTXYZ$, HOS$, 1, 8)
       Call REPLA(TTXYZ$, "Generación de O/Fabricación", 11, 56)
       Call REPLA(TTXYZ$, USERTER$, 69, 24)
       Call REGAPP("ANOTAORF", MKS$(NRO&) + TTXYZ$)
       Call REGAPP("ANOTAORF", MKS$(NRO&) + Space$(88))
       TTXYZ$ = Space$(76)
       CA12$ = FORMATNUM(CAN#, "F12.1")
       Call REPLA(TTXYZ$, "     - Cantidad =  " & CA12$, 11, 56)
       Call REGAPP("ANOTAORF", MKS$(NRO&) + TTXYZ$)
       Call REGAPP("ANOTAORF", MKS$(NRO&) + Space$(88))
       Call REGAPP("ANOTAORF", MKS$(NRO&) + String$(88, "="))
       Call CIERRARCH("ANOTAORF")
       '\\\OT-06-0209-WAR-FIN///
       GoSub GRARESERV
       '
       If DISTA% > 0 Then
         For KU& = 1 To ULTREG&("!DISTALLE")
           XXZ$ = MKI$(0) + REGLEIDO$("!DISTALLE", KU&)
           Call REPLA(XXZ$, MKS$(NRO&), 1, 4)
           Call REGAPP("DISTORFA", XXZ$)
         Next KU&
         Call CIERRARCH("DISTORFA")
       End If
       '
       GoSub DEPUORFEM
    End If
    '
    GoSub GRASECOP
    '
    Call CIERRARCH("ORFANUE")
    Call CIERRARCH("ORDEFABR")
    Call CIERRARCH("INDIORF")
    Call CIERRARCH("RESERVA")
    Call CIERRARCH("OPERFAB")
    Call CIERRARCH("OPERFAP")
    Call CIERRARCH("ECORDEP")
    Call CIERRARCH("ECORDEF")
    Call CIERRARCH("ORFAEMI")
    '
    GoSub PRINORFAB
    '
    If TERENUME$ <> "" Then
       Call COMUNI(TERENUME$)
    End If
    '
    GoTo 100
    '
    '
' ************************************* subrutinas asociadas
    '
FORIPRE:
    FORIPRE$ = TRIM$(CONTROL$("", "FORMORFA"))
    If FORIPRE$ <> "" Then
        If EXISTE%(LUDAT$ + FORIPRE$ + ".FRM") <> 1 Then
            RMCC$ = REPLACAR$(LU$ + FORIPRE$ + ".FRM", "\", "/")
            Call COMUNI("Error de Configuración:\  \No Existe Formulario de Impresión\de Orden de Fabricación:\  \'" + RMCC$ + "'.")
            FORIPRE$ = ""
        End If
    End If
    '
    FORIPRE2$ = TRIM$(CONTROL$("", "FORMORFB"))
    If FORIPRE2$ <> "" Then
        If EXISTE%(LUDAT$ + FORIPRE2$ + ".FRM") <> 1 Then
            RMCC$ = REPLACAR$(LU$ + FORIPRE2$ + ".FRM", "\", "/")
            Call COMUNI("Error de Configuración:\  \No Existe Formulario de Impresión\de Orden de Fabricación:\  \'" + RMCC$ + "'.")
            FORIPRE2$ = ""
        End If
    End If
    Return
    '
MONUME:
    MONUME$ = "AUTO"
    If CONTROL$("", "NUMEORFA") = "MANUAL" Then
        MONUME$ = "MANUAL"
    End If
    Return
    '
GESUBORDEN:
    GESUBOR% = 0
    If CONTROL$("", "GESUBOR") = "SI" Then
        GESUBOR% = 1
    End If
    '
    VERITAND% = 0
    If CONTROL$("", "VERITAND") = "SI" Then
        VERITAND% = 1
    End If
    '
    EXISEPRO% = 0
    If CONTROL$("", "EXISEPRO") = "SI" Then
        EXISEPRO% = 1
    End If
    '
    Return
    '
OBSERVACI:
    OBSERVOF% = 0
    If CONTROL$("", "OBSERVOF") = "SI" Then
        OBSERVOF% = 1
    End If
    Return
    '
MODOGEN:
    If PRECOLAN% = 0 Then
       PRECOLAN% = (-1)
       If CONTROL$("", "PRECOLAN") = "SI" Then
          PRECOLAN% = 1
       End If
    End If
    '
    If INCLUPRI% = 0 Then
       INCLUPRI% = (-1)
       If CONTROL$("ORFABRI", "MPRIUNIQ") = "INCLUPRI" Then
          INCLUPRI% = 1
       End If
    End If
    '
    If FRALOMAX% = 0 Then
       FRALOMAX% = (-1)
       If CONTROL$("", "FRALOMAX") = "SI" Then
          FRALOMAX% = 1
       End If
    End If
    '
    If VALINDOP% = 0 Then
       VALINDOP% = (-1)
       If CONTROL$("", "VALINDOP") = "SI" Then
          VALINDOP% = 1
       End If
    End If
    '
    MODORFA% = 1
    If REPIM% <> 2 Then
401    MODORFA% = COMALTER%("Puede Generar Ordenes de Fabricación\en Forma Directa, sobre la Base\del Cálculo MRP o de una O/Venta\\Directa\MRP\O/Venta\Cancelar")
       If MODORFA% = 1 Then
             OBJEPRI$ = "ORFANUE-1"
          ElseIf MODORFA% = 2 Then
             OBJEPRI$ = "ORFANUE-9"
          ElseIf MODORFA% = 3 Then
             Call MENSERR(24, "No Disponible en la Presente Versión")
             GoTo 401
             '
             Call GENPEDPEN
             Call SELECHO("INDIPEP/Indice General de Pedidos Pendientes")
             NPX& = Val(VALACT1$("INDIPEP"))
             If NPX& <> 0 Then
               OFABRIC.Label1.Caption = CStr(NPX&)
               FORSELPPE.Show 1
               RFFS$ = OFABRIC.Label1
               RFFA$ = OFABRIC.Label12
               Call CIERRARCH("*.*")
               Exit Sub
             Else
               Call CIERRARCH("*.*")
               Exit Sub
             End If
          Else
             Call CIERRARCH("*.*")
             Exit Sub
       End If
       If MONUME$ = "MANUAL" Then
          OBJEPRI$ = OBJEPRI$ + "M"
       End If
    End If
    Return
    '
ARMAPRIMERA:
    REGXX& = ULTREG&("ORFANUE")
    If REGXX& < 1 Then
        NRO& = 1
      Else
        NRO& = 1 + CVS(Left$(REGLEIDO$("ORFANUE", REGXX&), 4))
    End If
    '
    If CONTROL$("", "NUMEORFA") = "MANUAL" Then
      NRO& = Val(InputBox$("Número de la Próxima O/F", "Numeración de Ordenes de Fabricación", TRIM$(Str$(NRO&))))
      If NRO& < 1 Then Exit Sub
      '
      RFFX$ = RECFILT$("ORFANUE", 1, MKS$(NRO&))
      If RFFX$ <> "" Then
        Call MENSERR(24, "Numero ya Utilizado")
        GoTo ARMAPRIMERA
      End If
    End If
    '
    X0$ = REGBLAN$("ORFANUE")
    Call REPLA(X0$, MKS$(NRO&), 1, 4)
    Call REPLA(X0$, MKI$(HO%), 69, 2)
    Call REPLA(X0$, MKI$(HO%), 109, 2)
    Call REPLA(X0$, MKI$(HO%), 113, 2)
    Return
    '
VERICODI:
    CIOF% = 0: UMED$ = "U.": LOSTA = 1
    CI% = CVI(Mid$(X1$, 71, 2))
    If CI% < 1 Or CI% > NUMAS% Then
         Call MENSERR(24, "Falta Código de Producto.")
       Else
         CIOF% = CI%
         UMED$ = UNIMED$(CI%)
         LOSTA = LOREPOMA(CI%)
    End If
    Return
    '
BORENACCE:
    XACCE$ = ""
    RGACC& = 0
    For U& = 1 To ULTREG&("BOREACCE")
       YCC$ = REGLEIDO$("BOREACCE", U&)
       If CVI(Left$(YCC$, 2)) = CIOF% Then
          RGACC& = U&
          YACCE$ = YCC$
          GoTo 710
       End If
    Next U&
    '
    YACCE$ = String$(128, 0)
    Call REPLA(YACCE$, MKI$(CIOF%), 1, 2)
    Call REPLA(YACCE$, Space$(28), 3, 28)
    Call REPLA(YACCE$, Space$(40), 63, 40)
    Call REPLA(YACCE$, Space$(18), 105, 16)
    '
710 Call REPLA(YACCE$, TALLE$(CIOF%), 3, 12)
    Call REPLA(YACCE$, COLOR$(CIOF%), 15, 16)
    '
    RFF$ = RECFILT$("ESTRUNUE", 1, MKI$(CIOF%))
    If RFF$ <> "" Then
       REGEST& = CVS(Left$(RFF$, 4))
       CIHIL% = CVI(Mid$(REGLEIDO$("ESTRUNUE", REGEST&), 3, 2))
       Call REPLA(YACCE$, TIPOHIL$(CIHIL%), 63, 12)
       Call REPLA(YACCE$, TITULO$(CIHIL%), 75, 12)
       Call REPLA(YACCE$, COLOR$(CIHIL%), 87, 16)
    End If
    '
    XACCE$ = YACCE$
    If REPIM% <> 2 Then
       XACCE$ = OBJPLA$("BORENACCE", YACCE$)
    End If
    '
    If XACCE$ <> "" Then
       PROVHIL$ = TEL2CLI$((-1) * CVI(Mid$(XACCE$, 103, 2)))
       PARTHIL$ = Mid$(XACCE$, 105, 16)
       '
       LARCUTE$ = CSTRING$(Mid$(XACCE$, 31, 4), 3, 8, 1, 1)
       ACHCUTE$ = CSTRING$(Mid$(XACCE$, 35, 4), 3, 8, 1, 1)
       LARMATE$ = CSTRING$(Mid$(XACCE$, 39, 4), 3, 8, 1, 1)
       ACHMATE$ = CSTRING$(Mid$(XACCE$, 43, 4), 3, 8, 1, 1)
       LARCUCO$ = CSTRING$(Mid$(XACCE$, 47, 4), 3, 8, 1, 1)
       ACHCUCO$ = CSTRING$(Mid$(XACCE$, 51, 4), 3, 8, 1, 1)
       LARMACO$ = CSTRING$(Mid$(XACCE$, 55, 4), 3, 8, 1, 1)
       ACHMACO$ = CSTRING$(Mid$(XACCE$, 59, 4), 3, 8, 1, 1)
       '
       If RGACC& < 1 Then RGACC& = 1 + ULTREG&("BOREACCE")
       Call GRAREG("BOREACCE", XACCE$, RGACC&)
    End If
    Call CIERRARCH("BOREACCE")
    '
    Return

BORENCOMPLE:
    Call BLANARCH$("!BCOMPLEM")
    HAYCOMP% = 0
    For U& = 1 To ULTREG&("BCOMPLEM")
       YCC$ = REGLEIDO$("BCOMPLEM", U&)
       If CVI(Left$(YCC$, 2)) = CIOF% Then
          CACOM = CVS(Mid$(YCC$, 49, 4)) * CAN#
          Call REPLA(YCC$, MKS$(CACOM), 53, 4)
          CACOM = CVS(Mid$(YCC$, 57, 4)) * CAN#
          Call REPLA(YCC$, MKS$(CACOM), 61, 4)
          CACOM = CVS(Mid$(YCC$, 69, 4)) * CAN#
          Call REPLA(YCC$, MKS$(CACOM), 73, 4)
          Call REGAPP("!BCOMPLEM", YCC$)
          HAYCOMP% = 1
       End If
    Next U&
    '
    If HAYCOMP% = 0 Then
       YCC$ = REGBLAN$("BCOMPLEM")
       Call REPLA(YCC$, MKI$(CIOF%), 1, 2)
       Call REPLA(YCC$, "Cuerpo", 3, 14): Call REGAPP("!BCOMPLEM", YCC$)
       Call REPLA(YCC$, "Manga", 3, 14): Call REGAPP("!BCOMPLEM", YCC$)
       Call REPLA(YCC$, "Puño", 3, 14): Call REGAPP("!BCOMPLEM", YCC$)
       Call REPLA(YCC$, "Cuello", 3, 14): Call REGAPP("!BCOMPLEM", YCC$)
       Call REPLA(YCC$, "Tiras", 3, 14): Call REGAPP("!BCOMPLEM", YCC$)
       Call REPLA(YCC$, "Capucha", 3, 14): Call REGAPP("!BCOMPLEM", YCC$)
       Call REPLA(YCC$, "Bolsillo", 3, 14): Call REGAPP("!BCOMPLEM", YCC$)
       Call REPLA(YCC$, "Presilla", 3, 14): Call REGAPP("!BCOMPLEM", YCC$)
    End If
    '
    REFCOMP$ = ""
    TITU$ = " - " + TRIM$(CSTRING$(MKD$(CAN#), 3, 8, 0, 2)) + " Prendas " + TRIM$(CODEXT$(CIOF%)) + " - " + DESCRIT0$(CIOF%)
    '
    VTB% = 0
    If REPIM% <> 2 Then
       VTB% = VENTABU%("BCOMPLEM/TITUPAN=Orden de Complementos:" + TITU$)
    End If
    '
    If VTB% >= 0 Then
       If HAYCOMP% > 0 Then
           J& = 0
           For U& = 1 To ULTREG&("BCOMPLEM")
             YCC$ = REGLEIDO$("BCOMPLEM", U&)
             If CVI(Left$(YCC$, 2)) <> CIOF% Then
                J& = J& + 1
                Call GRAREG("BCOMPLEM", YCC$, J&)
             End If
           Next U&
         Else
           J& = ULTREG&("BCOMPLEM")
       End If
       For U& = 1 To ULTREG&("!BCOMPLEM")
          YCC$ = REGLEIDO$("!BCOMPLEM", U&)
          CACOM = CVS(Mid$(YCC$, 53, 4)) / CAN#
          Call REPLA(YCC$, MKS$(CACOM), 49, 4)
          CACOX = CVS(Mid$(YCC$, 61, 4)) / CAN#
          Call REPLA(YCC$, MKS$(CACOX), 57, 4)
          CACOX = CVS(Mid$(YCC$, 73, 4)) / CAN#
          Call REPLA(YCC$, MKS$(CACOX), 69, 4)
          Call GRAREG("!BCOMPLEM", YCC$, U&)
          J& = J& + 1
          Call GRAREG("BCOMPLEM", YCC$, J&)
          '
          If CACOM > 0 Then
             If REFCOMP$ <> "" Then REFCOMP$ = REFCOMP$ + " - "
             REFCOMP$ = REFCOMP$ + TRIM$(Mid$(YCC$, 3, 14))
          End If
       Next U&
       Call SETULTREG("BCOMPLEM", J&)
    End If
    Return
    '
AJUCANORF:
    CAMRP = 0
    CANORF = LOSTA
    LOECON = LOTESTAN(CIOF%)
    If MODORFA% = 2 Then
         CAMRP = CVD(Mid$(REGACT$("ORFAEMI"), 49, 8))
         CANORF = CAMRP
         If CAMRP < LOECON Then
            CANORF = LOECON
            Call MENSERR(24, "La Cantidad Pedida es Inferior\a un Lote Económico.\  \Ajustada a un Lote Standard Económico.")
         End If
      ElseIf MODORFA% = 1 Then
         CANORF = LOECON
    End If
    '
    If LOSTA < CANORF Then
        XLMX$ = MKI$(CI%) + MKS$(CANORF)
10111   XLMY$ = OBJPLA$("CARLOREPOMA", XLMX$)
        If XLMY$ <> "" Then
            LOSTA = CVS(Mid$(XLMY$, 3, 4))
            '
            If LOSTA <= 30000 Then
                 LOMA& = LOSTA
               ElseIf LOSTA <= 3000000 Then
                 LOMA& = (-1) * LOSTA / 100
               Else
                 Call MENSERR(24, "Lote de Reposición fuera de Rango")
                 GoTo 10111
            End If
            LOTEMAI% = LOMA&
            REGMAS& = CIOF% + 1
            If REGMAS& > 1 Then
                XRMX$ = REGLEIDO$("MASTER", REGMAS&)
                Call REPLA(XRMX$, MKI$(LOTEMAI%), 103, 2)
                Call GRAREG("MASTER", XRMX$, REGMAS&)
            End If
        End If
    End If
    Return
    '
ARMASEGUNDA:
    NRO& = CVS(Left$(X1$, 4))
    REOFE& = 0
    FENTRESOL% = HO%
    If MODORFA% = 2 Then
        FENTRESOL% = CVI(Mid$(REGACT$("ORFAEMI"), 57, 2))
        REOFE& = REGNUM&("ORFAEMI")
    End If
    '
    FENTREGA% = HO%
    If MODORFA% = 2 Then FENTREGA% = FENTRESOL% ' Por MRP
    If FENTREGA% < HO% Then
        FEX = FENTREGA%
        Call MENSERR(24, "Imposible Cumplir la Fecha de\Entrega Solicitada del " + FECHATEX$(FEX) + ".\  \Entrega Re-Programada para el " + HOS$ + ".")
        FENTREGA% = HO%
    End If
    '
    DESTINO$ = "STOCK": If TIMATE%(CIOF%) = 4 Then DESTINO$ = "SEMIELA"
    If SPY$ <> "" Then DESTINO$ = SPY$
    '
    Call REPLA(X1$, UMED$, 73, 4)
    Call REPLA(X1$, MKD$(CDbl(CANORF)), 89, 8)
    Call REPLA(X1$, MKS$(LOSTA), 97, 4)
    Call REPLA(X1$, DESTINO$, 101, 8)
    Call REPLA(X1$, MKI$(FENTREGA%), 109, 2)
    Call REPLA(X1$, SPX$, 115, 4)
    Call REPLA(X1$, MKI$(TIMATE%(CIOF%)), 257, 1)
    '
    Return
    '
VALIDA1:
    VALIDA% = (-1)
    X3$ = X2$: X1$ = X2$: X0$ = X2$
    If CVS(Mid$(X3$, 1, 4)) < 1 Then Return
    If CVI(Mid$(X3$, 69, 2)) < 1 Then Return
    If CVI(Mid$(X3$, 71, 2)) < 1 Then Return
    '
    VALIDA% = 0
    If TRIM$(Mid$(X3$, 73, 4)) = "" Then
        Call REPLA(X3$, "UNID", 73, 4)
    End If
    If CVD(Mid$(X3$, 89, 8)) < 1 Then Return
    If CVS(Mid$(X3$, 97, 4)) < 1 Then Return
    If CVI(Mid$(X3$, 109, 2)) < HO% Then
        Call COMUNI("Fecha de Entrega Proyectada\es Anterior a la Fecha Actual")
        Return
    End If
    FENTRE% = CVI(Mid$(X3$, 109, 2))
    '
    CAN# = CVD(Mid$(X3$, 89, 8))
    CANORF = CAN#
    If FRALOMAX% < 0 Then
       If CAN# > CVS(Mid$(X3$, 97, 4)) Then
          CALOTX = CAN#
          Call REPLA(X3$, MKS$(CALOTX), 97, 4)
       End If
    End If
    LOTE# = CDbl(CVS(Mid$(X3$, 97, 4)))
    If CAN# / LOTE# > 6 Then
        Call COMUNI("Cantidades no procesables\superan límite de 6 lotes\por Orden de Fabricación.")
        Return
    End If
    '
    DESTI$ = Mid$(X3$, 101, 8)
    DESTJ$ = ECOARCH$("DESTFAB", DESTI$)
    If TRIM$(DESTJ$) <> "" Then
        DESTI$ = DESTI$ + " - " + DESTJ$
    End If
    '
    If Asc(Mid$(X3$, 111, 1)) > 9 Then
      PRIO% = Asc(Mid$(X3$, 111, 1)) Mod 10
      Call REPLA(X3$, Chr$(PRIO%), 111, 1)
    End If
    PRIO% = Asc(Mid$(X3$, 111, 1))
    FLAN% = CVI(Mid$(X3$, 113, 2))
    Status% = 0
    If FLAN% <= HO% Then
       If HAYDISTO%(CIOF%, NRO&, CANORF) = 1 Then
          Status% = 1
       End If
    End If
    Call REPLA(X3$, Chr$(Status%), 112, 1)
    '
    CI% = CVI(Mid$(X3$, 71, 2))
    DESCRI$ = TRIM$(CODEXT$(CI%)) + " - " + TRIM$(DESCRIT$(CI%))
    Call REPLA(X3$, DESCRI$, 5, 64)
    '
    TORFANUE$ = X3$     ' prepara registro para ORFANUE.DAT
    '
    FGEN = HO%
    STS$ = TRIM$(Str$(Status%)): While Len(STS$) < 3: STS$ = " " + STS$: Wend
    PRS$ = TRIM$(Str$(PRIO%)): While Len(PRS$) < 3: PRS$ = " " + PRS$: Wend
    SECPROD = CVS(SPX$)
    FCIE = 0
    '
    X$ = REGBLAN$("INDIORF")
    Call REPLA(X$, MKS$(NRO&), 1, 4)
    Call REPLA(X$, FECHATEX$(FGEN), 5, 8)
    If Status% > 0 Then Call REPLA(X$, STS$, 13, 3)
    If PRIO% > 0 Then Call REPLA(X$, PRS$, 16, 3)
    Call REPLA(X$, ECOARCH$("SECPROD", MKS$(SECPROD)), 21, 20)
    If FCIE <> 0 Then Call REPLA(X$, FECHATEX$(FCIE), 43, 8)
    '
    TINDIORF$ = X$      ' prepara registro para INDIORF.DAT
    VALIDA% = 1
    Return
    '
OBSERVOF:
    VALIDA% = 0
    Call CIERRARCH("!OBSERVOF")
    Call BLANARCH("!OBSERVOF")
    If REPIM% <> 2 Then
       If OBSERVOF% = 1 Then
          VALIDA% = VENTABU%("OBSERVOF/TITUPAN=Orden de Fabricacion Numero" + Str$(CVS(Left$(X3$, 4))))
       End If
    End If
    Return
    '
GELOTES:
    NORDE% = 0
    CAPEN# = CAN#
    '
    While CAPEN# > 0
        CALOT# = LOTE#
        If CALOT# > CAPEN# Then
            CALOT# = CAPEN#
        End If
        '
        NORDE% = NORDE% + 1
        CILOTE%(NORDE%) = CI%
        CALOTE#(NORDE%) = CALOT#
        CAPEN# = CAPEN# - CALOT#
    Wend
    Return
    '
GESUBOR:
    If GESUBOR% = 1 Then
        CIXX% = CIOF%
        Call EXPLOMUL(CIXX%)
        For II% = 1 To CONTAMUL%
            PAD%(II%) = 1
        Next II%
        PAD%(CONTAMUL%) = 0
        '
        For II% = 1 To CONTAMUL% - 1
            If NIVE%(II% + 1) <= NIVE%(II%) Then
                PAD%(II%) = 0
              Else
                FIN& = ULTREG&("ARTIPRO")
                For J& = 1 To FIN&
                  X$ = REGLEIDO$("ARTIPRO", J&)
                    If CVI(Left$(X$, 2)) = CIK%(II%) Then
                      If CVI(Mid$(X$, 3, 2)) > 0 Then
                        PAD%(II%) = 0
                        GoTo 29
                      End If
                    End If
                Next J&
            End If
29      Next II%
        '
        SELSUBOR.LIMATSEL.Clear
        SELSUBOR.Label2 = "EMITIR"
        For II% = 2 To CONTAMUL%
          If PAD%(II%) = 1 Then
            CI00% = CIK%(II%)
            TTXX$ = AJUSTI$(CODEXT$(CI00%), 16) + AJUSTI$(DESCRIT0$(CI00%), 48) + FORMATNUM$(CAN# * CANT(II%), "F8.1")
            SELSUBOR.LIMATSEL.AddItem TTXX$
          End If
        Next II%
        '
        For U& = 1 To SELSUBOR.LIMATSEL.ListCount
          SELSUBOR.LIMATSEL.Selected(U& - 1) = True
        Next U&
        '
        Screen.MousePointer = 1
        On Error Resume Next
        SELSUBOR.BOTREC.SetFocus
        SELSUBOR.LIMATSEL.TopIndex = 0
        On Error GoTo 0
        SELSUBOR.Show 1
        '
        DoEvents
        If TRIM$(SELSUBOR.Label2 = "") Then
           Call COMUNI("No se Generarán Sub-Ordenes")
           GoTo 291
        End If
        '
        For II% = 1 To SELSUBOR.LIMATSEL.ListCount
          If SELSUBOR.LIMATSEL.Selected(II% - 1) = True Then
            TTXX$ = SELSUBOR.LIMATSEL.List(II% - 1)
            NORDE% = NORDE% + 1
            CILOTE%(NORDE%) = CODINT%(TRIM$(Left$(TTXX$, 15)))
            CALOTE#(NORDE%) = Val(Mid$(TTXX$, 65, 8))
          End If
        Next II%
        '
    End If
    '
291 Return
    '
GERESMAT:
    NBRS$ = TRIM$(Str$(NRO&)): While Len(NBRS$) < 6
                                  NBRS$ = "0" + NBRS$
                               Wend
    NBRS$ = "OF-" + NBRS$ + "-"
    Call BLANARCH("!RESERVA")
    '
    For II% = 1 To NORDE%
       PUNMAT$(II%) = ""
       CODICONJ$ = CODEXT$(CILOTE%(II%))
       NORFABRI$ = ""
       CACONJU# = CALOTE#(II%)
       VACONTROL% = 0: If PRECOLAN% = 1 Then VACONTROL% = 1
       '
       Call CARCOMPO(CODICONJ$, NORFABRI$, CACONJU#, VACONTROL%, 1)
       If VACONTROL% < 0 Then Return         ' cancelacion
       '
       HAYFALTA% = 0
       For U& = 1 To ULTREG&("!MATEROF")
         ZZ$ = REGLEIDO$("!MATEROF", U&)
         If CVS(Mid$(ZZ$, 25, 4)) > 0.05 Then HAYFALTA% = 1
       Next U&
       '
       NORFA$ = NBRS$ + TRIM$(Str$(II%))
       '
       If HAYFALTA% > 0 Then
         If COMALTER%("Hay Faltantes de Materiales\para la Orden de Fabricación " + NORFA$ + "\Codigo: " + TRIM$(CODICONJ$) + ".\\Cancelar\Generar Igual") <> 2 Then
           VACONTROL% = (-1)
           Return
         End If
       End If
       '
       For U& = 1 To ULTREG&("!MATEROF")
          ZZ$ = REGLEIDO$("!MATEROF", U&)
          CIIX% = CVI(Left$(ZZ$, 2))
          CANTX# = CVS(Mid$(ZZ$, 5, 4))
          USOII = CANTX# / CACONJU#
          '
          If CANTX# >= 0.05 Then
             YY$ = REGBLAN$("RESERVA")
             Call REPLA(YY$, MKI$(CIIX%), 1, 2)
             Call REPLA(YY$, NORFA$, 3, 12)
             Call REPLA(YY$, MKI$(CIIX%), 15, 2)
             Call REPLA(YY$, MKI$(HO%), 17, 2)
             Call REPLA(YY$, MKD$(CANTX#), 19, 8)
             Call REPLA(YY$, MKD$(CANTX#), 35, 8)
             Call REPLA(YY$, MKS$(USOII), 75, 4)
             Call REGAPP("!RESERVA", YY$)
             PUNMAT$(II%) = PUNMAT$(II%) + MKS$(ULTREG&("!RESERVA"))
         End If
       Next U&
    Next II%
    '
    Call CIERRARCH("!RESERVA")
    Return
    '
GRAORFNUE:
    Call REPLA(TORFANUE$, MKS$(NRO&), 1, 4)
    Call REPLA(TINDIORF$, MKS$(NRO&), 1, 4)
    Call REGAPP("ORFANUE", TORFANUE$)
    Z$ = RECFILT$("INDIORF", 1, MKS$(NRO&))
    If Len(Z$) >= 4 Then
          REGXX& = CVS(Left$(Z$, 4))
          Call GRAREG("INDIORF", TINDIORF$, REGXX&)
        Else
          Call REGAPP("INDIORF", TINDIORF$)
    End If
    Return
    '
GRAORDEFA:
    NBRS$ = TRIM$(Str$(NRO&)): While Len(NBRS$) < 6
                                  NBRS$ = "0" + NBRS$
                               Wend
    NBRS$ = "OF-" + NBRS$ + "-"
    For II% = 1 To NORDE%
       NORFA$ = NBRS$ + TRIM$(Str$(II%))
       DESCRI$ = TRIM$(CODEXT$(CILOTE%(II%))) + " - " + TRIM$(DESCRIT$(CILOTE%(II%)))
       '
       YY$ = REGBLAN$("ORDEFABR")
       Call REPLA(YY$, MKS$(NRO&), 1, 4)
       Call REPLA(YY$, DESCRI$, 5, 64)
       Call REPLA(YY$, NORFA$, 69, 12)
       Call REPLA(YY$, MKI$(HO%), 81, 2)
       Call REPLA(YY$, MKI$(CILOTE%(II%)), 83, 2)
       Call REPLA(YY$, UNIMED$(CILOTE%(II%)), 85, 4)
       Call REPLA(YY$, MKD$(CALOTE#(II%)), 89, 8)
       Call REPLA(YY$, DESTI$, 97, 8)
       Call REPLA(YY$, MKI$(FENTRE%), 105, 2)
       Call REPLA(YY$, Chr$(PRIO%), 107, 1)
       Call REPLA(YY$, Chr$(Status%), 108, 1)
       Call REPLA(YY$, MKI$(FLAN%), 109, 2)
       Call REPLA(YY$, MKS$(SECPROD), 113, 4)
       '
       Call REGAPP("ORDEFABR", YY$)
       REGORFA& = ULTREG&("ORDEFABR")
       '
       CANS$ = " - (" + TRIM$(Str$(CALOTE#(II%))) + ")"
       DESCRI1$ = Left$(DESCRI$, 52 - Len(CANS$)) + CANS$
       '
       For U& = ULTREG&("ECORDEF") + 1 To REGORFA&
          Call GRAREG("ECORDEF", REGBLAN$("ECORDEF"), U&)
       Next U&
       '
       ECORDEF$ = AJUSTI$(NORFA$, 12) + DESCRI1$
       Call GRAREG("ECORDEF", ECORDEF$, REGORFA&)
       If Status% = 1 Then
          Call REGAPP("ECORDEP", ECORDEF$)
       End If
       '
    Next II%
    Return
    '
GRARESERV:
    LONUO% = Len(NBRS$)
    For U& = 1 To ULTREG&("!RESERVA")
       YY$ = REGLEIDO$("!RESERVA", U&)
       Call REPLA(YY$, NBRS$, 3, LONUO%)
       Call GRAREG("!RESERVA", YY$, U&)
       Call REGAPP("RESERVA", YY$)
    Next U&
    Return
    '
DEPUORFEM:
    'If MODORFA% = 2 Then
    Call BLOQARCH("ORFAEMI")
    '
    For II% = 1 To NORDE%
      CALOTEX# = CALOTE#(II%)
      CIFX% = CILOTE%(II%)
      If CIFX% > 0 Then
        RFOE$ = RECFILT$("ORFAEMI", 1, MKI$(CIFX%))
        For KKU& = 1 To Len(RFOE$) Step 4
          REOFEM& = CVS(Mid$(RFOE$, KKU&, 4))
          X$ = REGLEIDO$("ORFAEMI", REOFEM&)
          CAMENO# = CVD(Mid$(X$, 49, 8))
          If CAMENO# <= CALOTEX# Then
              CALOTEX# = CALOTEX# - CAMENO#
              CAMENO# = 0
            Else
              CAMENO# = CAMENO# - CALOTEX#
              CALOTEX# = 0
          End If
          Call REPLA(X$, MKD$(CAMENO#), 49, 8)
          Call REPLA(X$, FORMATNUM$(CAMENO#, "F10.1"), 29, 10)
          Call GRAREG("ORFAEMI", X$, REOFEM&)
          If CALOTEX# < 0.05 Then Exit For
        Next KKU&
      End If
    Next II%
    '
    JK& = 0: FIN& = ULTREG&("ORFAEMI")
    For REOFEM& = 1 To FIN&
      X$ = REGLEIDO$("ORFAEMI", REOFEM&)
      If CVD(Mid$(X$, 49, 8)) >= 0.05 Then
        JK& = JK& + 1
        Call GRAREG("ORFAEMI", X$, JK&)
      End If
    Next REOFEM&
    Call SETULTREG&("ORFAEMI", JK&)
    Call CIERRARCH("ORFAEMI")
    Call FRESHECHO("ORFAEMI")
    '
    'End If
    Return
    '
GRASECOP:
    Call BLANARCH("!OPERFAB")
    For II% = 1 To NORDE%
       NORFA$ = NBRS$ + TRIM$(Str$(II%))
       CIJK% = CILOTE%(II%)
       CAPENDE# = CALOTE#(II%)
       '
       PUNOPE$(II%) = ""
       SECUOPSTA$ = RECFILT$("SECOPER", 1, MKI$(CIJK%))
       If Len(SECUOPSTA$) > 0 Then
          For TTI% = 1 To Len(SECUOPSTA$) Step 4
             TTL& = CVS(Mid$(SECUOPSTA$, TTI%, 4))
             SCP$ = REGLEIDO$("SECOPER", TTL&)
             If CVI(Left$(SCP$, 2)) = CIJK% Then
                NO% = CVI(Mid$(SCP$, 3, 2))
                DOPE$ = Mid$(SCP$, 5, 48)
                EQ1$ = Mid$(SCP$, 53, 6)
                EQAUX$ = Mid$(SCP$, 103, 18)
                LOTESTA = LOSTA: If LOTESTA < 1 Then LOTESTA = 1
                CALOTES = 1
                MINPREP = 60 * CVS(Mid$(SCP$, 59, 4)) * CALOTES
                MINFIJO = 60 * CVS(Mid$(SCP$, 63, 4)) * CALOTES
                MINVARI = CAPENDE# * CVS(Mid$(SCP$, 67, 4))
                MINDEMO = (MINFIJO + MINVARI) * CVS(Mid$(SCP$, 75, 4)) / 100
                MINLIMP = 60 * CVS(Mid$(SCP$, 79, 4)) * CALOTES
                MINLOTE = MINPREP + MINFIJO + MINVARI + MINDEMO + MINLIMP
                '
                HORAPROD = MINLOTE / 60
                CAOPS = CVS(Mid$(SCP$, 83, 4))
                '
' YA ESTA
                BVOF% = BINVAL%(AJUSTI$(NORFA$, 12))
                NUOROP% = 0
                FEAL% = HOY(HOS$)
                HOAL% = Int(HORANUM(Time$))
                USUA% = USNBR% Mod 100
                DALTA$ = Chr$(NUOROP%) + MKI$(FEAL%) + MKI$(HOAL%) + Chr$(USUA%)
                INFOBLA$ = Space$(6) + String$(26, 0) + Space$(16)
                PROGBLA$ = Space$(6) + String$(26, 0)
                '
                ZZ$ = REGBLAN$("OPERFAB")
                Call REPLA(ZZ$, NORFA$, 1, 12)
                Call REPLA(ZZ$, MKI$(CIJK%), 13, 2)
                Call REPLA(ZZ$, MKI$(NO%), 15, 2)
                Call REPLA(ZZ$, DOPE$, 17, 48)
                Call REPLA(ZZ$, MKI$(BVOF%), 65, 2)
                Call REPLA(ZZ$, DALTA$, 67, 6)
                Call REPLA(ZZ$, EQ1$, 73, 6)
                Call REPLA(ZZ$, MKD$(CAPENDE#), 79, 8)
                Call REPLA(ZZ$, MKD$(CAPENDE#), 95, 8)
                Call REPLA(ZZ$, MKS$(MINPREP), 103, 4)
                Call REPLA(ZZ$, MKS$(MINFIJO), 107, 4)
                Call REPLA(ZZ$, MKS$(MINVARI), 111, 4)
                Call REPLA(ZZ$, MKS$(MINDEMO), 115, 4)
                Call REPLA(ZZ$, MKS$(MINLIMP), 123, 4)
                Call REPLA(ZZ$, MKS$(MINLOTE), 131, 4)
                Call REPLA(ZZ$, String$(16, 0), 135, 16)
                Call REPLA(ZZ$, INFOBLA$, 151, 48)
                Call REPLA(ZZ$, MKS$(TTL&), 199, 4)
                Call REPLA(ZZ$, EQAUX$, 203, 18)
                Call REPLA(ZZ$, MKS$(CAOPS), 221, 4)
                Call REPLA(ZZ$, PROGBLA$, 225, 32)
                '
                Call REGAPP("!OPERFAB", ZZ$)
                PUNOPE$(II%) = PUNOPE$(II%) + MKS$(ULTREG&("!OPERFAB"))
                If REPIM% <> 2 Then
                   If Status% = 1 Then
                      Call REGAPP("OPERFAB", ZZ$)
                   End If
                End If
             End If
          Next TTI%
       End If
       '
       SECUOPSTB$ = RECFILT$("BCOMPLEM", 1, MKI$(CIJK%))
       If Len(SECUOPSTB$) > 0 Then
          For TTI% = 1 To Len(SECUOPSTB$) Step 4
             TTL& = CVS(Mid$(SECUOPSTB$, TTI%, 4))
             SCP$ = REGLEIDO$("BCOMPLEM", TTL&)
             If CVI(Left$(SCP$, 2)) = CIJK% Then
                NO% = 999
                DOPE$ = Mid$(SCP$, 3, 14)
                EQ1$ = Mid$(SCP$, 33, 6)
                EQAUX$ = String$(18, 0)
                LOTESTA = LOSTA: If LOTESTA < 1 Then LOTESTA = 1
                CALOTES = 1
                MINPREP = CVS(Mid$(SCP$, 65, 4))
                MINFIJO = 0
                MINVARI = CVS(Mid$(SCP$, 73, 4))
                MINDEMO = 0
                MINLIMP = 0
                MINLOTE = MINPREP + MINFIJO + MINVARI + MINDEMO + MINLIMP
                '
                If MINLOTE > 0.1 Then
                   HORAPROD = MINLOTE / 60
                   CAOPS = 1
                   '
' YA ESTA
                   BVOF% = BINVAL%(AJUSTI$(NORFA$, 12))
                   NUOROP% = 0
                   FEAL% = HOY(HOS$)
                   HOAL% = Int(HORANUM(Time$))
                   USUA% = USNBR% Mod 100
                   DALTA$ = Chr$(NUOROP%) + MKI$(FEAL%) + MKI$(HOAL%) + Chr$(USUA%)
                   INFOBLA$ = Space$(6) + String$(26, 0) + Space$(16)
                   PROGBLA$ = Space$(6) + String$(26, 0)
                   '
                   ZZ$ = REGBLAN$("OPERFAB")
                   Call REPLA(ZZ$, NORFA$, 1, 12)
                   Call REPLA(ZZ$, MKI$(CIJK%), 13, 2)
                   Call REPLA(ZZ$, MKI$(NO%), 15, 2)
                   Call REPLA(ZZ$, DOPE$, 17, 48)
                   Call REPLA(ZZ$, MKI$(BVOF%), 65, 2)
                   Call REPLA(ZZ$, DALTA$, 67, 6)
                   Call REPLA(ZZ$, EQ1$, 73, 6)
                   Call REPLA(ZZ$, MKD$(CAPENDE#), 79, 8)
                   Call REPLA(ZZ$, MKD$(CAPENDE#), 95, 8)
                   Call REPLA(ZZ$, MKS$(MINPREP), 103, 4)
                   Call REPLA(ZZ$, MKS$(MINFIJO), 107, 4)
                   Call REPLA(ZZ$, MKS$(MINVARI), 111, 4)
                   Call REPLA(ZZ$, MKS$(MINDEMO), 115, 4)
                   Call REPLA(ZZ$, MKS$(MINLIMP), 123, 4)
                   Call REPLA(ZZ$, MKS$(MINLOTE), 131, 4)
                   Call REPLA(ZZ$, String$(16, 0), 135, 16)
                   Call REPLA(ZZ$, INFOBLA$, 151, 48)
                   Call REPLA(ZZ$, MKS$(TTL&), 199, 4)
                   Call REPLA(ZZ$, EQAUX$, 203, 18)
                   Call REPLA(ZZ$, MKS$(CAOPS), 221, 4)
                   Call REPLA(ZZ$, PROGBLA$, 225, 32)
                   '
                   Call REGAPP("!OPERFAB", ZZ$)
                   PUNOPE$(II%) = PUNOPE$(II%) + MKS$(ULTREG&("!OPERFAB"))
                   If REPIM% <> 2 Then
                      If Status% = 1 Then
                         Call REGAPP("OPERFAB", ZZ$)
                      End If
                   End If
                End If
             End If
          Next TTI%
       End If
    Next II%
    Call CIERRARCH("OPERFAB")
    Return
    '
PRINORFAB:
237 If TRIM$(FORIPRE$) = "" Then Return
    '
    CODFOR$ = TRIM$(UCase$(FORIPRE$))
    MARIZW = Val(ATRIFORM$(CODFOR$, "MARIZQMM"))
    MARSUW = Val(ATRIFORM$(CODFOR$, "MARSUPMM"))
    '
    DETDES = Val(ATRIFORM$(CODFOR$, "DETDESMM"))
    DETHAS = Val(ATRIFORM$(CODFOR$, "DETHASMM"))
    INTERLI = Val(ATRIFORM$(CODFOR$, "INTERLIN"))
    If INTERLI < 1 Then
    Call COMUNI("Error de Configuración de Interlinea\en Formulario '" + CODFOR$ + ".FRM'.")
       INTERLI = 4
    End If
    '
    MAXITEMS = Val(ATRIFORM$(CODFOR$, "MAX-ITEM"))
    If MAXITEMS < 1 Then MAXITEMS = 999
    '
    CACOPI% = Val(ATRIFORM$(CODFOR$, "CAN-COP"))
    If CACOPI% < 1 Or CACOPI% > 6 Then
       Call COMUNI("Error de Configuración de Cantidad de Copias\en Formulario '" + CODFOR$ + ".FRM'.")
       CACOPI% = 1
    End If
    '
    PORIAN = Printer.Orientation  ' retiene orientacion
    LARHOAN = Printer.Height
    ANCHOAN = Printer.Width
    ORIENTACT = 1       ' anterior
    NORPAPSI = Printer.PaperSize
    '
    PORI$ = TRIM$(UCase$(ATRIFORM$(CODFOR$, "ORIENTA")))
    If PORI$ = "PAISAJE" Or PORI$ = "2" Then
       ORIENTACT = 2
    End If
    '
    DESCRIOR$ = TRIM$(CODEXT$(CILOTE%(1))) + " - " + TRIM$(DESCRIT0$(CILOTE%(1)))
    '
    FINOR& = NORDE%
    Screen.MousePointer = 11
    NOCONFIRM% = 1
    For II% = 1 To NORDE%
       ILI& = II%
       Call TRACE(20, ILI&, FINOR&)
       For COP% = 1 To CACOPI%
          'Printer.Orientation = ORIENTACT
          Call SETSPOOL("ORIENTATION", Str$(ORIENTACT))
          KKINI% = 1
249       KK1% = 1: NH% = 1
250       ATTX% = 0
          Call SETSPOOL("ORIENTATION", Str$(ORIENTACT))
          Call CARFORWIN
          
          NORFA$ = NBRS$ + TRIM$(Str$(II%))
          NORFP$ = NORFA$: If InStr(EMPREAC$, "BORENS") > 0 Then NORFP$ = Mid$(NORFA$, 4, 6)
          If NH% > 1 Then NORFP$ = NORFP$ + "/" + TRIM$(Str$(NH%))
          Call PRINTFORWIN(NORFP$, "NUME-COM", 0, 0)
          Call PRINTFORWIN(DESCRIOR$, "REF-MAT1", 0, 0)
          If REPIM% = 2 Then
             Call PRINTFORWIN("( COPIA )", "REF-MAT2", 0, 0)
          End If
          '
          FEX = FGEN
          Call PRINTFORWIN(FECHATEX$(FEX), "FECH-EMI", 0, 0)
          Call PRINTFORWIN(Time$, "HORA-EMI", 0, 0)
          Call PRINTFORWIN(USERNAME$, "USUARIO", 0, 0)
          FEX = FLAN%
          Call PRINTFORWIN(FECHATEX$(FEX), "FECH-LAN", 0, 0)
          FEX = FENTRE%
          Call PRINTFORWIN(FECHATEX$(FEX), "FECH-VEN", 0, 0)
          '
          Call PRINTFORWIN(Str$(COP%), "ORD-COPI", 0, 0)
          Call PRINTFORWIN(DESTI$, "DESTINO", 0, 0)
            DESTI2$ = ECOARCH$("DESTFAB", DESTI$)
          Call PRINTFORWIN(DESTI2$, "DEPOSITO", 0, 0)
          Call PRINTFORWIN(TRIM$(Str$(II%)), "HOJA-NUM", 0, 0)
          Call PRINTFORWIN(TRIM$(Str$(NORDE%)), "HOJA-TOT", 0, 0)
          '
          CODICONJU% = 0: POSICONJU$ = ""
          If II% > 1 Then CODICONJU% = CILOTE%(1): POSICONJU$ = "1"
          CODISUCON% = 0: POSISUCON$ = ""
          For KKII% = II% - 1 To 1 Step -1
            CIXII% = CILOTE%(KKII%)
            Call AREXPLO(CIXII%, 1, 1)
            For KKJJ% = 1 To CAIT%
              If CIJ%(KKJJ%) = CILOTE%(II%) Then
                If CIXII% <> CODICONJU% Then
                  CODISUCON% = CIXII%
                  POSISUCON$ = TRIM$(Str$(KKII%))
                  'MsgBox CODEXT$(CILOTE%(II%)) & "   " & CODEXT$(CODISUCON%) & "   " & CODEXT$(CODICONJU%)
                End If
                GoTo 255
              End If
            Next KKJJ%
          Next KKII%
          '
255       Call PRINTFORWIN(TRIM$(CODEXT$(CILOTE%(II%))), "COD-CONJ", 0, 0)
          Call PRINTFORWIN(DESCRIT$(CILOTE%(II%)), "DES-CONJ", 0, 0)
          Call PRINTFORWIN(NUPLANO$(CILOTE%(II%)), "N-PLANO", 0, 0)
          Call PRINTFORWIN(PZPLANO$(CILOTE%(II%)), "Z-PLANO", 0, 0)
          Call PRINTFORWIN(RVPLANO$(CILOTE%(II%)), "R-PLANO", 0, 0)
          Call PRINTFORWIN(FEPLANO$(CILOTE%(II%)), "F-PLANO", 0, 0)
          Call PRINTFORWIN(NUHOREI$(CILOTE%(II%)), "N-HINSP", 0, 0)
          'Call PRINTFORWIN(ESPECIFI$(CILOTE%(II%)), "ESP-MAT", 0, 0)
          Call PRINTFORWIN(UNIMED$(CILOTE%(II%)), "UNI-CONJ", 0, 0)
          Call PRINTFORWIN(TALLE$(CILOTE%(II%)), "TAL-CONJ", 0, 0)
          Call PRINTFORWIN(COLOR$(CILOTE%(II%)), "COL-CONJ", 0, 0)
          Call PRINTFORWIN(REVICOD$(CILOTE%(II%)), "ZONA-VEN", 0, 0)
          FEX = FEREVI%(CILOTE%(II%))
          Call PRINTFORWIN(FECHATEX$(FEX), "FECH-CHE", 0, 0)
          CANS$ = CSTRING$(MKD$(CALOTE#(II%)), 3, 9, 1, 2)
          Call PRINTFORWIN(CANS$, "CAN-CONJ", 0, 0)
          CANS$ = CSTRING$(MKD$(CALOTE#(II%) * PESUNI(CILOTE%(II%))), 3, 9, 1, 2)
          Call PRINTFORWIN(CANS$, "PES-CONJ", 0, 0)
          '
          Call PRINTFORWIN(TRIM$(REFCOMP$), "REF-CONJ", 0, 0)
          '
          LARCUTE$ = CODEXT$(CODICONJU%)
          Call PRINTFORWIN(TRIM$(LARCUTE$), "LAR-CUTE", 0, 0)
          ACHCUTE$ = CODEXT$(CODISUCON%)
          Call PRINTFORWIN(TRIM$(ACHCUTE$), "ACH-CUTE", 0, 0)
          LARMATE$ = NUPLANO$(CODICONJU%)
          Call PRINTFORWIN(TRIM$(LARMATE$), "LAR-MATE", 0, 0)
          ACHMATE$ = NUPLANO$(CODISUCON%)
          Call PRINTFORWIN(TRIM$(ACHMATE$), "ACH-MATE", 0, 0)
          LARCUCO$ = RVPLANO$(CODICONJU%)
          Call PRINTFORWIN(TRIM$(LARCUCO$), "LAR-CUCO", 0, 0)
          ACHCUCO$ = RVPLANO(CODISUCON%)
          Call PRINTFORWIN(TRIM$(ACHCUCO$), "ACH-CUCO", 0, 0)
          LARMACO$ = POSICONJU$
          Call PRINTFORWIN(TRIM$(LARMACO$), "LAR-MACO", 0, 0)
          ACHMACO$ = POSISUCON$
          Call PRINTFORWIN(TRIM$(ACHMACO$), "ACH-MACO", 0, 0)
          '
          YAC = DETDES
          ITNU% = 0
          If DETHAS > DETDES + INTERLI Then
            '
            EPAC$ = TRIM$(UCase$(EMPREAC$))
            If INCLUPRI% = 1 Or InStr(EPAC$, "ARCON") > 0 Then
              If Len(PUNMAT$(II%)) > 4 Then
                Call PRINTFORWIN("Segun Vale de Materiales", "DES-MPRI", 0, YAC)
                GoTo 260
              End If
            End If
            '
            For KK% = KK1% To Len(PUNMAT$(II%)) Step 4
              TTL& = CVS(Mid$(PUNMAT$(II%), KK%, 4))
              If TTL& > 0 Then
                '
                ITNU% = ITNU% + 1
                If ITNU% > MAXITEMS Then GoTo 260
                '
                If YAC + INTERLI > DETHAS Then
                   KK1% = KK%
                   ATTX% = 1
                   NH% = NH% + 1
                   GoTo 260
                End If
                '
                YY$ = REGLEIDO$("!RESERVA", TTL&)
                CIXI% = CVI(Mid$(YY$, 1, 2))
                CANTX# = CVD(Mid$(YY$, 19, 8))
                '
                If CANTX# > 0 Then
                  Call PRINTFORWIN(TRIM$(CODEXT$(CIXI%)), "COD-MPRI", 0, YAC)
                  Call PRINTFORWIN(DESCRIT$(CIXI%), "DES-MPRI", 0, YAC)
                  Call PRINTFORWIN(UNIMED$(CIXI%), "UNI-MPRI", 0, YAC)
                  Call PRINTFORWIN(TIPOHIL$(CIXI%), "TIP-MPRI", 0, YAC)
                  Call PRINTFORWIN(TITULO$(CIXI%), "TIT-MPRI", 0, YAC)
                  Call PRINTFORWIN(COLOR$(CIXI%), "COL-MPRI", 0, YAC)
                  Call PRINTFORWIN(PROVHIL$, "PRO-MPRI", 0, YAC)
                  Call PRINTFORWIN(PARTHIL$, "PAR-MPRI", 0, YAC)
                  CANS$ = CSTRING$(MKS$(CANTX#), 3, 9, 1, 2)
                  Call PRINTFORWIN(CANS$, "CAN-MPRI", 0, YAC)
                  CANS$ = CSTRING$(MKS$(CANTX# / CALOTE#(II%)), 3, 9, 3, 2)
                  Call PRINTFORWIN(CANS$, "USO-MPRI", 0, YAC)
                  Call PRINTFORWIN(UBISTO$(CIXI%), "UBI-MPRI", 0, YAC)
                  YAC = YAC + INTERLI
                End If
              End If
            Next KK%
            '
260         YAC = DETDES
            ITNU% = 0
            TMPRE = 0: TMPRO = 0
            If KKINI% < 1 Then KKINI% = 1
            For KK% = KKINI% To Len(PUNOPE$(II%)) Step 4
              TTL& = CVS(Mid$(PUNOPE$(II%), KK%, 4))
              If TTL& > 0 Then
                '
                If INCLUPRI% = 1 Or InStr(EPAC$, "ARCON") > 0 Then
                  If YAC + INTERLI > DETHAS Then
                    KKINI% = KK%
                    Call SETSPOOL("NEWPAGE", "")
                    GoTo 249
                  End If
                End If
                '
                YY$ = REGLEIDO$("!OPERFAB", TTL&)
                '
                NOPE% = CVI(Mid$(YY$, 15, 2))
                DOPE$ = Mid$(YY$, 17, 48)
                DESOP$ = OBSEROP$(CILOTE%(II%), NOPE%)
                If TRIM$(DESOP$) = "" Then DESOP$ = DOPE$
                EQ1$ = Mid$(YY$, 73, 6)
                EQ1$ = EQ1$ + " - " + ECOARCH$("PADMAQ", EQ1$)
                EQ2$ = Mid$(YY$, 203, 6)
                EQ2$ = EQ2$ + " - " + ECOARCH$("PADMAQ", EQ2$)
                EQ3$ = Mid$(YY$, 209, 6)
                EQ3$ = EQ3$ + " - " + ECOARCH$("PADMAQ", EQ3$)
                MPRE = Int(CVS(Mid$(YY$, 103, 4)) + 0.5)
                MPRO = 0: For KI% = 107 To 123 Step 4
                              MPRO = MPRO + Int(CVS(Mid$(YY$, KI%, 4)) + 0.5)
                           Next KI%
                CAOPS = CVS(Mid$(YY$, 221, 4)): If CAOPS > 99 Or CAOPS < 0 Then CAOPS = 0
                TMPRE = TMPRE + MPRE
                TMPRO = TMPRO + MPRO
                MPREI% = 0: If MPRE < 16384 Then MPREI% = MPRE
                MPROI% = 0: If MPRO < 16384 Then MPROI% = MPRO
                '
                ITNU% = ITNU% + 1
                If ITNU% > MAXITEMS Then GoTo 269
                '
                Call PRINTFORWIN(Str$(NOPE%), "NOP-OPER", 0, YAC)
                Call PRINTFORWIN(TRIM$(DOPE$), "DES-OPER", 0, YAC)
                Call PRINTFORWIN(TRIM$(DESOP$), "OBS-OPER", 0, YAC)
                Call PRINTFORWIN(TRIM$(EQ1$), "EQUIPO1", 0, YAC)
                Call PRINTFORWIN(TRIM$(EQ2$), "EQUIPO2", 0, YAC)
                Call PRINTFORWIN(TRIM$(EQ3$), "EQUIPO3", 0, YAC)
                Call PRINTFORWIN(AJUSTD$(HORAMINU$(MPREI%), 6), "TIE-PREP", 0, YAC)
                Call PRINTFORWIN(AJUSTD$(HORAMINU$(MPROI%), 6), "TIE-PROD", 0, YAC)
                Call PRINTFORWIN(AJUSTD$(HORAMINU$(MPREI% + MPROI%), 6), "TIE-TOTA", 0, YAC)
                Call PRINTFORWIN(Str$(CAOPS), "NUOPERS", 0, YAC)
                Call PRINTFORWIN(NUHOPROF$(CILOTE%(II%), NOPE%), "N-HPROF", 0, YAC)
                '
                YAC = YAC + INTERLI
                '
              End If
269         Next KK%
            '
            TMPREI% = 0: If TMPRE < 16384 Then TMPREI% = TMPRE
            TMPROI% = 0: If TMPRO < 16384 Then TMPROI% = TMPRO
            Call PRINTFORWIN(AJUSTD$(HORAMINU$(TMPREI%), 6), "TOT-PREP", 0, 0)
            Call PRINTFORWIN(AJUSTD$(HORAMINU$(TMPROI%), 6), "TOT-PROD", 0, 0)
            Call PRINTFORWIN(AJUSTD$(HORAMINU$(TMPREI% + TMPROI%), 6), "TOT-TIEM", 0, 0)
            '
          End If
          '
          GoSub PRIESPECIFI
          GoSub PRIPLACONTROL
          GoSub PRINPUEPUNTO
          GoSub PRIMAGENES
          '
          YAC = 0
          For U& = 1 To ULTREG&("!OBSERVOF")
             TXOB$ = TRIM$(REGLEIDO$("!OBSERVOF", U&))
             Call PRINTFORWIN(TXOB$, "OBSERVA", 0, YAC)
             YAC = YAC + INTERLI
          Next U&
          Call CIERRARCH("!OBSERVOF")
          '
          If DISTA% > 0 Then
            GoSub PRIDETALLES
          End If
   
          'Printer.EndDoc
          Call SETSPOOL("ENDDOC", "")
          '
          If ATTX% <> 0 Then GoTo 250
          '
       Next COP%
       '
       Call SETSPOOL("ORIENTATION", Str$(PORIAN))
       Call SAVESPOOL("Orden de Fabricación " + NORFA$, OKX%)
       '
    Next II%
    '
'    Call SETSPOOL("ORIENTATION", Str$(PORIAN))
'    Call SAVESPOOL("Orden de Fabricación " + NORFA$, OKX%)
    '
    On Error Resume Next
      Printer.Orientation = PORIAN
      'Printer.Height = LARHOAN
      'Printer.Width = ANCHOAN
      Printer.PaperSize = NORPAPSI
    On Error GoTo 0
    '
    For II% = 1 To NORDE%
       '
       CIXOFX% = CILOTE%(II%)
       PUNOPFABRI$ = PUNOPE$(II%)
       GoSub DOCUANEXOS
       '
    Next II%
    '
337 If TRIM$(FORIPRE2$) = "" Then GoTo 990
    '
    EPAC$ = TRIM$(UCase$(EMPREAC$))
    If INCLUPRI% = 1 Or InStr(EPAC$, "ARCON") > 0 Then
      If Len(PUNMAT$(1)) <= 4 Then
        GoTo 990
      End If
    End If
    '
    PORIAN = Printer.Orientation
    LARHOAN = Printer.Height
    ANCHOAN = Printer.Width
    CODFOR$ = TRIM$(UCase$(FORIPRE2$))
    MARIZW = Val(ATRIFORM$(CODFOR$, "MARIZQMM"))
    MARSUW = Val(ATRIFORM$(CODFOR$, "MARSUPMM"))
    '
    DETDES = Val(ATRIFORM$(CODFOR$, "DETDESMM"))
    DETHAS = Val(ATRIFORM$(CODFOR$, "DETHASMM"))
    INTERLI = Val(ATRIFORM$(CODFOR$, "INTERLIN"))
    If INTERLI < 1 Then
    Call COMUNI("Error de Configuración de Interlinea\en Formulario '" + CODFOR$ + ".FRM'.")
       INTERLI = 4
    End If
    '
    MAXITEMS = Val(ATRIFORM$(CODFOR$, "MAX-ITEM"))
    If MAXITEMS < 1 Then MAXITEMS = 999
    '
    CACOPI% = Val(ATRIFORM$(CODFOR$, "CAN-COP"))
    If CACOPI% < 1 Or CACOPI% > 6 Then
       Call COMUNI("Error de Configuración de Cantidad de Copias\en Formulario '" + CODFOR$ + ".FRM'.")
       CACOPI% = 1
    End If
    '
    ORIENTACT = 1
    PORI$ = TRIM$(UCase$(ATRIFORM$(CODFOR$, "ORIENTA")))
    If PORI$ = "PAISAJE" Or PORI$ = "2" Then
       ORIENTACT = 2
    End If
    '
    If VALINDOP% = 1 Then GoTo 445 ' caso vales de MAINIERI
    '
345 For II% = 1 To NORDE%
       For COP% = 1 To CACOPI%
          'Printer.Orientation = ORIENTACT
          Call SETSPOOL("ORIENTATION", Str$(ORIENTACT))
          KK1% = 1: NH% = 1
350       ATTX% = 0
          Call SETSPOOL("ORIENTATION", Str$(ORIENTACT))
          Call CARFORWIN
          NORFA$ = NBRS$ + TRIM$(Str$(II%))
          NORFP$ = NORFA$: If InStr(EMPREAC$, "BORENS") > 0 Then NORFP$ = Mid$(NORFA$, 4, 6)
          If NH% > 1 Then NORFP$ = NORFP$ + "/" + TRIM$(Str$(NH%))
          Call PRINTFORWIN(NORFP$, "NUME-COM", 0, 0)
          '
          FEX = FGEN
          Call PRINTFORWIN(FECHATEX$(FEX), "FECH-EMI", 0, 0)
          Call PRINTFORWIN(Time$, "HORA-EMI", 0, 0)
          Call PRINTFORWIN(USERNAME$, "USUARIO", 0, 0)
          FEX = FLAN%
          Call PRINTFORWIN(FECHATEX$(FEX), "FECH-LAN", 0, 0)
          FEX = FENTRE%
          Call PRINTFORWIN(FECHATEX$(FEX), "FECH-VEN", 0, 0)
          '
          Call PRINTFORWIN(Str$(COP%), "ORD-COPI", 0, 0)
          Call PRINTFORWIN(DESTI$, "DESTINO", 0, 0)
          Call PRINTFORWIN(TRIM$(Str$(II%)), "HOJA-NUM", 0, 0)
          Call PRINTFORWIN(TRIM$(Str$(NORDE%)), "HOJA-TOT", 0, 0)
          '
          If REPIM% = 2 Then
             Call PRINTFORWIN("( COPIA )", "REF-MAT2", 0, 0)
          End If
          '
          CODICONJU% = 0: POSICONJU$ = ""
          If II% > 1 Then CODICONJU% = CILOTE%(1): POSICONJU$ = "1"
          CODISUCON% = 0: POSISUCON$ = ""
          For KKII% = II% - 1 To 1 Step -1
            CIXII% = CILOTE%(KKII%)
            Call AREXPLO(CIXII%, 1, 1)
            For KKJJ% = 1 To CAIT%
              If CIJ%(KKJJ%) = CILOTE%(II%) Then
                If CIXII% <> CODICONJU% Then
                  CODISUCON% = CIXII%
                  POSISUCON$ = TRIM$(Str$(KKII%))
                  'MsgBox CODEXT$(CILOTE%(II%)) & "   " & CODEXT$(CODISUCON%) & "   " & CODEXT$(CODICONJU%)
                End If
                GoTo 355
              End If
            Next KKJJ%
          Next KKII%
          '
355       Call PRINTFORWIN(TRIM$(CODEXT$(CILOTE%(II%))), "COD-CONJ", 0, 0)
          Call PRINTFORWIN(DESCRIT$(CILOTE%(II%)), "DES-CONJ", 0, 0)
          Call PRINTFORWIN(NUPLANO$(CILOTE%(II%)), "N-PLANO", 0, 0)
          Call PRINTFORWIN(RVPLANO$(CILOTE%(II%)), "R-PLANO", 0, 0)
          Call PRINTFORWIN(PZPLANO$(CILOTE%(II%)), "Z-PLANO", 0, 0)
          Call PRINTFORWIN(NUHOREI$(CILOTE%(II%)), "N-HINSP", 0, 0)
          'Call PRINTFORWIN(ESPECIFI$(CILOTE%(II%)), "ESP-MAT", 0, 0)
          Call PRINTFORWIN(UNIMED$(CILOTE%(II%)), "UNI-CONJ", 0, 0)
          Call PRINTFORWIN(TALLE$(CILOTE%(II%)), "TAL-CONJ", 0, 0)
          Call PRINTFORWIN(COLOR$(CILOTE%(II%)), "COL-CONJ", 0, 0)
          Call PRINTFORWIN(REVICOD$(CILOTE%(II%)), "ZONA-VEN", 0, 0)
          FEX = FEREVI%(CILOTE%(II%))
          Call PRINTFORWIN(FECHATEX$(FEX), "FECH-CHE", 0, 0)
          CANS$ = CSTRING$(MKD$(CALOTE#(II%)), 3, 9, 1, 2)
          Call PRINTFORWIN(CANS$, "CAN-CONJ", 0, 0)
          CANS$ = CSTRING$(MKD$(CALOTE#(II%) * PESUNI(CILOTE%(II%))), 3, 9, 1, 2)
          Call PRINTFORWIN(CANS$, "PES-CONJ", 0, 0)
          '
          Call PRINTFORWIN(TRIM$(REFCOMP$), "REF-CONJ", 0, 0)
          '
          LARCUTE$ = CODEXT$(CODICONJU%)
          Call PRINTFORWIN(TRIM$(LARCUTE$), "LAR-CUTE", 0, 0)
          ACHCUTE$ = CODEXT$(CODISUCON%)
          Call PRINTFORWIN(TRIM$(ACHCUTE$), "ACH-CUTE", 0, 0)
          LARMATE$ = NUPLANO$(CODICONJU%)
          Call PRINTFORWIN(TRIM$(LARMATE$), "LAR-MATE", 0, 0)
          ACHMATE$ = NUPLANO$(CODISUCON%)
          Call PRINTFORWIN(TRIM$(ACHMATE$), "ACH-MATE", 0, 0)
          LARCUCO$ = RVPLANO$(CODICONJU%)
          Call PRINTFORWIN(TRIM$(LARCUCO$), "LAR-CUCO", 0, 0)
          ACHCUCO$ = RVPLANO(CODISUCON%)
          Call PRINTFORWIN(TRIM$(ACHCUCO$), "ACH-CUCO", 0, 0)
          LARMACO$ = POSICONJU$
          Call PRINTFORWIN(TRIM$(LARMACO$), "LAR-MACO", 0, 0)
          ACHMACO$ = POSISUCON$
          Call PRINTFORWIN(TRIM$(ACHMACO$), "ACH-MACO", 0, 0)
          '
          YAC = DETDES
          ITNU% = 0
          If DETHAS > DETDES + INTERLI Then
            For KK% = KK1% To Len(PUNMAT$(II%)) Step 4
              TTL& = CVS(Mid$(PUNMAT$(II%), KK%, 4))
              If TTL& > 0 Then
                '
                ITNU% = ITNU% + 1
                If ITNU% > MAXITEMS Then GoTo 360
                '
                If YAC + INTERLI > DETHAS Then
                   KK1% = KK%
                   ATTX% = 1
                   NH% = NH% + 1
                   GoTo 360
                End If
                '
                YY$ = REGLEIDO$("!RESERVA", TTL&)
                CIXI% = CVI(Mid$(YY$, 1, 2))
                CANTX# = CVD(Mid$(YY$, 19, 8))
                '
                Call PRINTFORWIN(TRIM$(CODEXT$(CIXI%)), "COD-MPRI", 0, YAC)
                Call PRINTFORWIN(DESCRIT$(CIXI%), "DES-MPRI", 0, YAC)
                Call PRINTFORWIN(UNIMED$(CIXI%), "UNI-MPRI", 0, YAC)
                Call PRINTFORWIN(TIPOHIL$(CIXI%), "TIP-MPRI", 0, YAC)
                Call PRINTFORWIN(TITULO$(CIXI%), "TIT-MPRI", 0, YAC)
                Call PRINTFORWIN(COLOR$(CIXI%), "COL-MPRI", 0, YAC)
                Call PRINTFORWIN(PROVHIL$, "PRO-MPRI", 0, YAC)
                Call PRINTFORWIN(PARTHIL$, "PAR-MPRI", 0, YAC)
                CANS$ = CSTRING$(MKS$(CANTX#), 3, 9, 1, 2)
                Call PRINTFORWIN(CANS$, "CAN-MPRI", 0, YAC)
                CANS$ = CSTRING$(MKS$(CANTX# / CALOTE#(II%)), 3, 9, 3, 2)
                Call PRINTFORWIN(CANS$, "USO-MPRI", 0, YAC)
                Call PRINTFORWIN(UBISTO$(CIXI%), "UBI-MPRI", 0, YAC)
                YAC = YAC + INTERLI
              End If
            Next KK%
            '
360         YAC = DETDES
            ITNU% = 0
            TMPRE = 0: TMPRO = 0
            For KK% = 1 To Len(PUNOPE$(II%)) Step 4
              TTL& = CVS(Mid$(PUNOPE$(II%), KK%, 4))
              If TTL& > 0 Then
               '
               YY$ = REGLEIDO$("!OPERFAB", TTL&)
                '
                NOPE% = CVI(Mid$(YY$, 15, 2))
                DOPE$ = Mid$(YY$, 17, 48)
                EQ1$ = Mid$(YY$, 73, 6)
                EQ1$ = EQ1$ + " - " + ECOARCH$("PADMAQ", EQ1$)
                EQ2$ = Mid$(YY$, 203, 6)
                EQ2$ = EQ2$ + " - " + ECOARCH$("PADMAQ", EQ2$)
                EQ3$ = Mid$(YY$, 209, 6)
                EQ3$ = EQ3$ + " - " + ECOARCH$("PADMAQ", EQ3$)
                MPRE = Int(CVS(Mid$(YY$, 103, 4)) + 0.5)
                MPRO = 0: For KI% = 107 To 123 Step 4
                              MPRO = MPRO + Int(CVS(Mid$(YY$, KI%, 4)) + 0.5)
                           Next KI%
                CAOPS = CVS(Mid$(YY$, 221, 4)): If CAOPS > 99 Or CAOPS < 0 Then CAOPS = 0
                TMPRE = TMPRE + MPRE
                TMPRO = TMPRO + MPRO
                MPREI% = 0: If MPRE < 16384 Then MPREI% = MPRE
                MPROI% = 0: If MPRO < 16384 Then MPROI% = MPRO
                '
                ITNU% = ITNU% + 1
                If ITNU% > MAXITEMS Then GoTo 369
                '
                Call PRINTFORWIN(Str$(NOPE%), "NOP-OPER", 0, YAC)
                Call PRINTFORWIN(TRIM$(DOPE$), "DES-OPER", 0, YAC)
                Call PRINTFORWIN(TRIM$(EQ1$), "EQUIPO1", 0, YAC)
                Call PRINTFORWIN(TRIM$(EQ2$), "EQUIPO2", 0, YAC)
                Call PRINTFORWIN(TRIM$(EQ3$), "EQUIPO3", 0, YAC)
                Call PRINTFORWIN(AJUSTD$(HORAMINU$(MPREI%), 6), "TIE-PREP", 0, YAC)
                Call PRINTFORWIN(AJUSTD$(HORAMINU$(MPROI%), 6), "TIE-PROD", 0, YAC)
                Call PRINTFORWIN(AJUSTD$(HORAMINU$(MPREI% + MPROI%), 6), "TIE-TOTA", 0, YAC)
                Call PRINTFORWIN(Str$(CAOPS), "NUOPERS", 0, YAC)
                '
                YAC = YAC + INTERLI
                '
              End If
369         Next KK%
            '
            TMPREI% = 0: If TMPRE < 16384 Then TMPREI% = TMPRE
            TMPROI% = 0: If TMPRO < 16384 Then TMPROI% = TMPRO
            Call PRINTFORWIN(AJUSTD$(HORAMINU$(TMPREI%), 6), "TOT-PREP", 0, 0)
            Call PRINTFORWIN(AJUSTD$(HORAMINU$(TMPROI%), 6), "TOT-PROD", 0, 0)
            Call PRINTFORWIN(AJUSTD$(HORAMINU$(TMPREI% + TMPROI%), 6), "TOT-TIEM", 0, 0)
            '
          End If
          '
370       YAC = DETDES
          ITNU% = 0
          For TTL& = 1 To ULTREG&("!BCOMPLEM")
             '
             YY$ = REGLEIDO$("!BCOMPLEM", TTL&)
             '
             ITNU% = ITNU% + 1
             If ITNU% > MAXITEMS Then GoTo 379
             '
             TTXX$ = Mid$(YY$, 3, 14)
             Call PRINTFORWIN(TRIM$(TTXX$), "PIE-COMP", 0, YAC)
             TTXX$ = Mid$(YY$, 17, 16)
             Call PRINTFORWIN(TRIM$(TTXX$), "DIB-COMP", 0, YAC)
             TTXX$ = Mid$(YY$, 33, 6)
             Call PRINTFORWIN(TRIM$(TTXX$), "MAQ-COMP", 0, YAC)
             TTXX$ = TRIM$(CSTRING$(Mid$(YY$, 39, 4), 3, 10, 1, 2))
             Call PRINTFORWIN(TRIM$(TTXX$), "LAR-COMP", 0, YAC)
             TTXX$ = TRIM$(CSTRING$(Mid$(YY$, 43, 4), 3, 10, 1, 2))
             Call PRINTFORWIN(TRIM$(TTXX$), "ACH-COMP", 0, YAC)
             TTXX$ = Mid$(YY$, 47, 2)
             Call PRINTFORWIN(TRIM$(TTXX$), "UNI-COMP", 0, YAC)
             TTXX$ = TRIM$(CSTRING$(Mid$(YY$, 53, 4), 3, 10, 1, 2))
             Call PRINTFORWIN(TRIM$(TTXX$), "CAN-COMP", 0, YAC)
             TTXX$ = TRIM$(CSTRING$(Mid$(YY$, 61, 4), 3, 10, 1, 2))
             Call PRINTFORWIN(TRIM$(TTXX$), "PES-COMP", 0, YAC)
             '
             YAC = YAC + INTERLI
379       Next TTL&
          '
          '
          YAC = 0
          For U& = 1 To ULTREG&("!OBSERVOF")
             TXOB$ = TRIM$(REGLEIDO$("!OBSERVOF", U&))
             Call PRINTFORWIN(TXOB$, "OBSERVA", 0, YAC)
             YAC = YAC + INTERLI
          Next U&
          Call CIERRARCH("!OBSERVOF")
          '
          'Printer.EndDoc
          Call SETSPOOL("ENDDOC", "")
          '
          If ATTX% <> 0 Then GoTo 350
          '
       Next COP%
       Call SETSPOOL("ORIENTATION", Str$(PORIAN))
       Call SAVESPOOL("Orden Complem. Fabricación " + NORFA$, OKX%)
    Next II%
    '
    'Call SETSPOOL("ORIENTATION", Str$(PORIAN))
    'Call SAVESPOOL("Orden Complem. Fabricación " + NORFA$, OKX%)
    '
    On Error Resume Next
      Printer.Orientation = PORIAN
      'Printer.Height = LARHOAN
      'Printer.Width = ANCHOAN
      Printer.PaperSize = NORPAPSI
    On Error GoTo 0
    '
    GoTo 990
    
    '
445 ' CASO VALES DE MAINIERI
    PORIAN = Printer.Orientation
    LARHOAN = Printer.Height
    ANCHOAN = Printer.Width
    For II% = 1 To Len(PUNOPE$(1)) / 4
       For COP% = 1 To CACOPI%
          'Printer.Orientation = ORIENTACT
          Call SETSPOOL("ORIENTATION", Str$(ORIENTACT))
          KK1% = 1: NH% = 1
450       ATTX% = 0
          Call SETSPOOL("ORIENTATION", Str$(ORIENTACT))
          Call CARFORWIN
          NORFA$ = NBRS$ + TRIM$(Str$(II%))
          NORFP$ = NORFA$: If InStr(EMPREAC$, "BORENS") > 0 Then NORFP$ = Mid$(NORFA$, 4, 6)
          If NH% > 1 Then NORFP$ = NORFP$ + "/" + TRIM$(Str$(NH%))
          Call PRINTFORWIN(NORFP$, "NUME-COM", 0, 0)
          If REPIM% = 2 Then
             Call PRINTFORWIN("( COPIA )", "REF-MAT2", 0, 0)
          End If
          '
          FEX = FGEN
          Call PRINTFORWIN(FECHATEX$(FEX), "FECH-EMI", 0, 0)
          Call PRINTFORWIN(Time$, "HORA-EMI", 0, 0)
          Call PRINTFORWIN(USERNAME$, "USUARIO", 0, 0)
          FEX = FLAN%
          Call PRINTFORWIN(FECHATEX$(FEX), "FECH-LAN", 0, 0)
          FEX = FENTRE%
          Call PRINTFORWIN(FECHATEX$(FEX), "FECH-VEN", 0, 0)
          '
          Call PRINTFORWIN(Str$(COP%), "ORD-COPI", 0, 0)
          Call PRINTFORWIN(DESTI$, "DESTINO", 0, 0)
          Call PRINTFORWIN(TRIM$(Str$(II%)), "HOJA-NUM", 0, 0)
          Call PRINTFORWIN(TRIM$(Str$(NORDE%)), "HOJA-TOT", 0, 0)
          '
          CODICONJU% = 0: POSICONJU$ = ""
          If II% > 1 Then CODICONJU% = CILOTE%(1): POSICONJU$ = "1"
          CODISUCON% = 0: POSISUCON$ = ""
          For KKII% = II% - 1 To 1 Step -1
            CIXII% = CILOTE%(KKII%)
            Call AREXPLO(CIXII%, 1, 1)
            For KKJJ% = 1 To CAIT%
              If CIJ%(KKJJ%) = CILOTE%(II%) Then
                If CIXII% <> CODICONJU% Then
                  CODISUCON% = CIXII%
                  POSISUCON$ = TRIM$(Str$(KKII%))
                  'MsgBox CODEXT$(CILOTE%(II%)) & "   " & CODEXT$(CODISUCON%) & "   " & CODEXT$(CODICONJU%)
                End If
                GoTo 455
              End If
            Next KKJJ%
          Next KKII%
          '
455       Call PRINTFORWIN(TRIM$(CODEXT$(CILOTE%(1))), "COD-CONJ", 0, 0)
          Call PRINTFORWIN(DESCRIT$(CILOTE%(1)), "DES-CONJ", 0, 0)
          Call PRINTFORWIN(NUPLANO$(CILOTE%(II%)), "N-PLANO", 0, 0)
          Call PRINTFORWIN(RVPLANO$(CILOTE%(II%)), "R-PLANO", 0, 0)
          Call PRINTFORWIN(PZPLANO$(CILOTE%(II%)), "Z-PLANO", 0, 0)
          Call PRINTFORWIN(NUHOREI$(CILOTE%(II%)), "N-HINSP", 0, 0)
          'Call PRINTFORWIN(ESPECIFI$(CILOTE%(II%)), "ESP-MAT", 0, 0)
          Call PRINTFORWIN(UNIMED$(CILOTE%(1)), "UNI-CONJ", 0, 0)
          Call PRINTFORWIN(TALLE$(CILOTE%(1)), "TAL-CONJ", 0, 0)
          Call PRINTFORWIN(COLOR$(CILOTE%(1)), "COL-CONJ", 0, 0)
          Call PRINTFORWIN(REVICOD$(CILOTE%(1)), "ZONA-VEN", 0, 0)
          FEX = FEREVI%(CILOTE%(1))
          Call PRINTFORWIN(FECHATEX$(FEX), "FECH-CHE", 0, 0)
          CANS$ = CSTRING$(MKD$(CALOTE#(1)), 3, 9, 1, 2)
          Call PRINTFORWIN(CANS$, "CAN-CONJ", 0, 0)
          CANS$ = CSTRING$(MKD$(CALOTE#(1) * PESUNI(CILOTE%(1))), 3, 9, 1, 2)
          Call PRINTFORWIN(CANS$, "PES-CONJ", 0, 0)
          '
          Call PRINTFORWIN(TRIM$(REFCOMP$), "REF-CONJ", 0, 0)
          '
          LARCUTE$ = CODEXT$(CODICONJU%)
          Call PRINTFORWIN(TRIM$(LARCUTE$), "LAR-CUTE", 0, 0)
          ACHCUTE$ = CODEXT$(CODISUCON%)
          Call PRINTFORWIN(TRIM$(ACHCUTE$), "ACH-CUTE", 0, 0)
          LARMATE$ = NUPLANO$(CODICONJU%)
          Call PRINTFORWIN(TRIM$(LARMATE$), "LAR-MATE", 0, 0)
          ACHMATE$ = NUPLANO$(CODISUCON%)
          Call PRINTFORWIN(TRIM$(ACHMATE$), "ACH-MATE", 0, 0)
          LARCUCO$ = RVPLANO$(CODICONJU%)
          Call PRINTFORWIN(TRIM$(LARCUCO$), "LAR-CUCO", 0, 0)
          ACHCUCO$ = RVPLANO(CODISUCON%)
          Call PRINTFORWIN(TRIM$(ACHCUCO$), "ACH-CUCO", 0, 0)
          LARMACO$ = POSICONJU$
          Call PRINTFORWIN(TRIM$(LARMACO$), "LAR-MACO", 0, 0)
          ACHMACO$ = POSISUCON$
          Call PRINTFORWIN(TRIM$(ACHMACO$), "ACH-MACO", 0, 0)
          '
460       YAC = DETDES
          ITNU% = 0
          TMPRE = 0: TMPRO = 0
          TTL& = CVS(Mid$(PUNOPE$(1), 4 * II% - 3, 4))
          '
          If TTL& > 0 Then
             '
             YY$ = REGLEIDO$("!OPERFAB", TTL&)
             '
             NOPE% = CVI(Mid$(YY$, 15, 2))
             DOPE$ = Mid$(YY$, 17, 48)
             EQ1$ = Mid$(YY$, 73, 6)
             EQ1$ = EQ1$ + " - " + ECOARCH$("PADMAQ", EQ1$)
             EQ2$ = Mid$(YY$, 203, 6)
             EQ2$ = EQ2$ + " - " + ECOARCH$("PADMAQ", EQ2$)
             EQ3$ = Mid$(YY$, 209, 6)
             EQ3$ = EQ3$ + " - " + ECOARCH$("PADMAQ", EQ3$)
             MPRE = Int(CVS(Mid$(YY$, 103, 4)) + 0.5)
             MPRO = 0: For KI% = 107 To 123 Step 4
                           MPRO = MPRO + Int(CVS(Mid$(YY$, KI%, 4)) + 0.5)
                        Next KI%
             CAOPS = CVS(Mid$(YY$, 221, 4)): If CAOPS > 99 Or CAOPS < 0 Then CAOPS = 0
             TMPRE = TMPRE + MPRE
             TMPRO = TMPRO + MPRO
             MPREI% = 0: If MPRE < 16384 Then MPREI% = MPRE
             MPROI% = 0: If MPRO < 16384 Then MPROI% = MPRO
             '
             Call PRINTFORWIN(Str$(NOPE%), "NOP-OPER", 0, YAC)
             Call PRINTFORWIN(TRIM$(DOPE$), "DES-OPER", 0, YAC)
             Call PRINTFORWIN(TRIM$(EQ1$), "EQUIPO1", 0, YAC)
             Call PRINTFORWIN(TRIM$(EQ2$), "EQUIPO2", 0, YAC)
             Call PRINTFORWIN(TRIM$(EQ3$), "EQUIPO3", 0, YAC)
             Call PRINTFORWIN(AJUSTD$(HORAMINU$(MPREI%), 6), "TIE-PREP", 0, YAC)
             Call PRINTFORWIN(AJUSTD$(HORAMINU$(MPROI%), 6), "TIE-PROD", 0, YAC)
             Call PRINTFORWIN(AJUSTD$(HORAMINU$(MPREI% + MPROI%), 6), "TIE-TOTA", 0, YAC)
             Call PRINTFORWIN(Str$(CAOPS), "NUOPERS", 0, YAC)
             '
             YAC = YAC + INTERLI
             '
          End If
          '
          '
          YAC = 0
          For U& = 1 To ULTREG&("!OBSERVOF")
             TXOB$ = TRIM$(REGLEIDO$("!OBSERVOF", U&))
             Call PRINTFORWIN(TXOB$, "OBSERVA", 0, YAC)
             YAC = YAC + INTERLI
          Next U&
          Call CIERRARCH("!OBSERVOF")
          '
          'Printer.EndDoc
          Call SETSPOOL("ENDDOC", "")
       Next COP%
       '
       Call SETSPOOL("ORIENTATION", Str$(PORIAN))
       Call SAVESPOOL("Vales de Operación O/F " + NORFA$, OKX%)
       '
    Next II%
    '
    Call SETSPOOL("ORIENTATION", Str$(PORIAN))
    Call SAVESPOOL("Vales de Operación O/F " + NORFA$, OKX%)
    '
    On Error Resume Next
      Printer.Orientation = PORIAN
      'Printer.Height = LARHOAN
      'Printer.Width = ANCHOAN
      Printer.PaperSize = NORPAPSI
    On Error GoTo 0
    '
    GoTo 990
    '
    '
    '
990 'Printer.Orientation = PORIAN ' restaura anterior
    'Call SETSPOOL("ORIENTATION", Str$(PORIAN))
    'Call SAVESPOOL("Orden de Fabricación " + NORFA$, OKX%)
    Screen.MousePointer = 1
    'If CANCELPRINT% < 1 Then
      If COMALTER%("Desea Continuar o Repetir\la Impresion de este Documento ?\\Continuar\Repetir Impresion") = 2 Then GoTo 237
    'End If
    '
    DOCUORIG$ = TRIM$(NORFA$)
    NUMEDOC$ = ""
    '
    EPAC$ = TRIM$(UCase$(EMPREAC$))
    If InStr(EPAC$, "ARCON") > 0 Then
       '
       Call CIERRARCH("*.*")
       Call BAJALDISCO
       '
       ANEXOCOM.Caption = "Ordenes de Fabricación - Documentos Anexos"
       YY$ = REGBLAN$("OCOMDETA")
       Call REPLA(YY$, MKI$(CILOTE%(1)), 9, 2)
       Call GRAREG("!OCOMDETA", YY$, 1)
       Call SETULTREG("!OCOMDETA", 1)
       Call CIERRARCH("!OCOMDETA")
       '
       ANEXOCOM.Show 1
       '
    End If
    '
Return
    '
    '
PRIPLACONTROL:
    '
    YAC = DETDES
    PPCX$ = RECFILT$("PLACONTR", 1, MKI$(CIOF%))
    INTERLIX = INTERLI
    '
    If Len(PPCX$) >= 4 Then
       EPAC$ = EMPREAC$
       If InStr(EPAC$, "NEUMANN") > 0 Then
          CARENEC% = Len(PPCX$) / 4
          INTERLI2 = (190 - 152) / (CARENEC% + 1)
          If INTERLI2 < INTERLIX Then INTERLIX = INTERLI2
       End If
    End If
    '
    For KK% = 1 To Len(PPCX$) Step 4
       TTL& = CVS(Mid$(PPCX$, KK%, 4))
       If TTL& > 0 Then
          If TTL& <= ULTREG&("PLACONTR") Then
             YY$ = REGLEIDO$("PLACONTR", TTL&)
             Call PRINTFORWIN(Mid$(YY$, 3, 2), "COT-PLCO", 0, YAC)
             Call PRINTFORWIN(Mid$(YY$, 5, 28), "REF-PLCO", 0, YAC)
             Call PRINTFORWIN(Mid$(YY$, 33, 32), "MED-PLCO", 0, YAC)
             Call PRINTFORWIN(Str$(CVS(Mid$(YY$, 65, 4))), "CAN-PLCO", 0, YAC)
             Call PRINTFORWIN(Mid$(YY$, 69, 12), "FRE-PLCO", 0, YAC)
             Call PRINTFORWIN(Mid$(YY$, 81, 160), "REA-PLCO", 0, YAC)
             '
             YAC = YAC + INTERLIX
          End If
       End If
    Next KK%
    '
    Return
    '
    '
PRINPUEPUNTO:
    '
    YAC = DETDES
    PPCX$ = RECFILT$("PUEPUNTO", 1, MKI$(CIOF%))
    For KK% = 1 To Len(PPCX$) Step 4
       TTL& = CVS(Mid$(PPCX$, KK%, 4))
       If TTL& > 0 Then
          If TTL& <= ULTREG&("PUEPUNTO") Then
             YY$ = REGLEIDO$("PUEPUNTO", TTL&)
             Call PRINTFORWIN(Mid$(YY$, 3, 24), "OPE-PUEP", 0, YAC)
             Call PRINTFORWIN(Mid$(YY$, 27, 24), "MAQ-PUEP", 0, YAC)
             Call PRINTFORWIN(Mid$(YY$, 51, 24), "PAR-PUEP", 0, YAC)
             Call PRINTFORWIN(Mid$(YY$, 75, 24), "HER-PUEP", 0, YAC)
             '
             YAC = YAC + INTERLI
          End If
       End If
    Next KK%
    '
Return
    '
PRIMAGENES:
   '
   If RUTARGRA$ = "" Then
     RUTARGRA$ = TRIM$(CONTROL$("", "RUTARGRA"))
     If RUTARGRA$ = "" Then RUTARGRA$ = TRIM$(LUDAT$)
     If Right$(RUTARGRA$, 1) <> "\" Then
       RUTARGRA$ = RUTARGRA$ + "\"
     End If
   End If
   '
   REBLA$ = REGBLAN$("INDIPLAN")
   RCIP$ = FILTERGETRECORD$("INDIPLAN", 1, MKI$(CIOF%))
   If RCIP$ <> REBLA$ Then
     For KII% = 0 To 3
       COPLANO$ = TRIM$(Mid$(RCIP$, 65 + 48 * KII%, 48))
       If COPLANO$ <> "" Then
         COPLANO$ = RUTARGRA$ + "\" + COPLANO$
         If EXISTE%(COPLANO$) > 0 Then
           CODIMAG$ = "DOCUGRA" + TRIM$(Str$(KII%))
           Call PRINTIMAGE(CODIMAG$, COPLANO$)
         End If
       End If
     Next KII%
   End If
   '
Return
    '
PRIDETALLES:
    YAC = 0 'DETDES
    INTERLIX = INTERLI
    '
    For KU& = 1 To ULTREG&("!DISTALLE")
      XXZ$ = REGLEIDO$("!DISTALLE", KU&)
        COTAX$ = TRIM$(Mid$(XXZ$, 33, 2))
      Call PRINTFORWIN(COTAX$, "PIE-COMP", 0, YAC)
        DETAX$ = TRIM$(Mid$(XXZ$, 35, 16))
      Call PRINTFORWIN(DETAX$, "DIB-COMP", 0, YAC)
        CATAX$ = TRIM$(Str$(CVS(Mid$(XXZ$, 9, 4))))
      Call PRINTFORWIN(CATAX$, "CAN-COMP", 0, YAC)
        OBTAX$ = TRIM$(Mid$(XXZ$, 13, 20))
      Call PRINTFORWIN(OBTAX$, "MAQ-COMP", 0, YAC)
        YAC = YAC + INTERLIX
    Next KU&
    '
Return
'
PRIESPECIFI:
   TTXXX$ = ESPECIFI$(CIOF%)
   Call FRATEXTO(TTXXX$, 128)
   '
   YAC = 0
   INTERLIX = INTERLI
   '
   For KKK% = 1 To CARETEX%
     If YAC < 89 Then
       Call PRINTFORWIN(RETEX$(KKK%), "ESP-MAT", 0, YAC)
     End If
     YAC = YAC + INTERLIX
   Next KKK%
Return
'
DOCUANEXOS:
    '
    DOCUORIG$ = NORFA$
    '
    If CONTROL$("ORDEFABR", "PRIPLANO") = "SI" Then
       Call PRIPLANO(CIXOFX%)
    End If
    '
    If CONTROL$("ORDEFABR", "PRIESPEC") = "SI" Then
       Call PRIHOJESP(CIXOFX%)
    End If
    '
    If CONTROL$("ORDEFABR", "PRIHOPRO") = "SI" Then
       For KK% = 1 To Len(PUNOPFABRI$) Step 4
          TTL& = CVS(Mid$(PUNOPFABRI$, KK%, 4))
          If TTL& > 0 Then
            YY$ = REGLEIDO$("!OPERFAB", TTL&)
            NOPE% = CVI(Mid$(YY$, 15, 2))
            Call PRIHOJPROFAB(CIXOFX%, NOPE%)
          End If
       Next KK%
    End If
    '
    If CONTROL$("ORDEFABR", "PRIHOINS") = "SI" Then
       Call PRIHOJRINSP(CIXOFX%)
    End If
    '
    DOCUORIG$ = ""
    '
Return
    '
    '
999 Call CIERRARCH("*.*")
    Call FRESHALL
    End Sub
    '

Sub ORFLAN()      ' LANZAMIENTO DE ORDEN DE FABRICACION
    '
    HO% = HOY(HOS$)
    LU$ = LUDAT$
    '
    TITU$ = "PROCESO DE LANZAMIENTO DE ORDENES DE FABRICACION"

100  NROOF$ = ""
     '
1400 If INDI% = 0 Then
       '
       Screen.MousePointer = 11
       Call COPYSTRU("ORFANUE", "ORFANUL")
       J& = 0
       '
       For I& = 1 To ULTREG&("ORFANUE")
         XX$ = REGLEIDO$("ORFANUE", I&)
         ST% = Asc(Mid$(XX$, 112, 1))
         '
         If ST% <> 3 Then
            If CVI(Mid$(XX$, 119, 2)) > 0 Then
                Call REPLA(XX$, MKI$(0), 119, 2)
                Call GRAREG("ORFANUE", XX$, I&)
            End If
         End If
         '
         If ST% = 0 Or ST% = 2 Then           ' solo proyectadas y suspendidas
            FEX = CVI(Mid$(XX$, 113, 2))
            Call REPLA(XX$, " (" + FECHATEX$(FEX) + ")          ", 54, 15)
            J& = J& + 1
            Call GRAREG("ORFANUL", XX$, J&)
         End If
       Next I&
       '
       Call SETULTREG("ORFANUL", J&)
       Call CIERRARCH("ORFANUL")
       INDI% = 1
       '
       Screen.MousePointer = 1
       If ULTREG&("ORFANUL") < 1 Then
          Call MENSERR(24, "No hay Ordenes de Fabricación Proyectadas\o Suspendidas que Puedan Lanzarse.")
          Call CIERRARCH("ORFANUL")
          Call CIERRARCH("*.*")
          GoTo 1499
       End If
       '
     End If
     
1405 Call FRESHECHO("ORFANUL")
     If ULTREG&("ORFANUL") < 1 Then
         INDI% = 0
         Call CIERRARCH("*.*")
         GoTo 1499
     End If
     Call SELECHO("ORFANUL/Ordenes de Trabajo para Lanzamiento")
     '
1410 NROOFL& = Val(VALACT1$("ORFANUL"))
     If NROOFL& < 1 Then
         INDI% = 0
         Call CIERRARCH("*.*")
         GoTo 1499
     End If
     '
     For I& = ULTREG&("ORFANUE") To 1 Step -1
       XX$ = REGLEIDO$("ORFANUE", I&)
       If CVS(Left$(XX$, 4)) = NROOFL& Then GoTo 1420
     Next I&
     '
     Call COMUNI("Inconsistencia en Base de Datos\de Ordenes de Fabricación.\  \Consulte al Soporte de Sistemas.")
     INDI% = 0
     Call CIERRARCH("*.*")
     GoTo 1499
     '
1420 YY$ = OBJPLA$("ORFALAN/Lanzamiento de Ordenes de Fabricación", XX$)
     If YY$ = "" Then GoTo 1405
     '
     For I& = J& + 1 To ULTREG&("ORFANUL")
        Call GRAREG("ORFANUL", String$(256, 0), I&)
     Next I&
     '
     FLAN% = CVI(Mid$(YY$, 113, 2))
     If FLAN% < HO% Then FLAN% = HO%
     Call LANORF(NROOFL&, FLAN%)
     '
     If APROLAN% = 1 Then
        J& = 0
        For I& = 1 To ULTREG&("ORFANUL")
           XX$ = REGLEIDO$("ORFANUL", I&)
           If CVS(Left$(XX$, 4)) <> NROOFL& Then
              J& = J& + 1
              Call GRAREG("ORFANUL", XX$, J&)
           End If
        Next I&
        Call SETULTREG("ORFANUL", J&)
     End If
     '
     GoTo 1405
     '
1499 Call ACTECOR(0)
     '
End Sub

Sub ORFSUS()
    '
    HO% = HOY(HOS$)
    LU$ = LUDAT$
    '
    TITU$ = "PROCESO DE SUSPENSION DE ORDENES DE FABRICACION"

100  NROOF$ = ""

     '
     'Call VENTANA(0, 0, 0, 0, "", "CLEAR")
     'Call ENCAPAN(0, "ORDENES DE FABRICACION")
     'Call PANTREC(4, 21, TITU$)
     '
1400 If INDI% = 0 Then
       '
       Screen.MousePointer = 11
       Call COPYSTRU("ORFANUE", "ORFANUL")
       J& = 0
       '
       For I& = 1 To ULTREG&("ORFANUE")
         XX$ = REGLEIDO$("ORFANUE", I&)
         ST% = Asc(Mid$(XX$, 112, 1))
         '
         If ST% <> 3 Then
            If CVI(Mid$(XX$, 119, 2)) > 0 Then
                Call REPLA(XX$, MKI$(0), 119, 2)
                Call GRAREG("ORFANUE", XX$, I&)
            End If
         End If
         '
         If ST% = 1 Then      ' solo lanzadas
            J& = J& + 1
            Call GRAREG("ORFANUL", XX$, J&)
         End If
       Next I&
       '
       If J& < 1 Then
          Call MENSERR(24, "No hay Ordenes de Fabricación\Lanzadas que Puedan Suspenderse.")
          Call CIERRARCH("ORFANUL")
          Screen.MousePointer = 1
          Call CIERRARCH("*.*")
          Exit Sub
       End If
       '
       For I& = J& + 1 To ULTREG&("ORFANUL")
         Call GRAREG("ORFANUL", String$(256, 0), I&)
       Next I&
       '
       Call CIERRARCH("ORFANUL")
       INDI% = 1
       '
       Screen.MousePointer = 1
       '
     End If
     
1405 Call FRESHECHO("ORFANUL")
     If ULTREG&("ORFANUL") < 1 Then INDI% = 0: Exit Sub
     Call SELECHO("ORFANUL/Ordenes de Trabajo para Suspensión")
     '
1410 NROOFL& = Val(VALACT1$("ORFANUL"))
     If NROOFL& < 1 Then
         INDI% = 0
         Exit Sub
     End If
     '
     For I& = ULTREG&("ORFANUE") To 1 Step -1
       XX$ = REGLEIDO$("ORFANUE", I&)
       If CVS(Left$(XX$, 4)) = NROOFL& Then GoTo 1420
     Next I&
     '
     Call COMUNI("Inconsistencia en Base de Datos\de Ordenes de Fabricación.\  \Consulte al Soporte de Sistemas.")
     INDI% = 0
     Exit Sub
     '
1420 YY$ = OBJPLA$("ORFALAN/Suspensión de Ordenes de Fabricación", XX$)
     If YY$ = "" Then GoTo 1405
     '
     Screen.MousePointer = 11
     For J& = I& To 1 Step -1
       XX$ = REGLEIDO$("ORFANUE", J&)
       If CVS(Left$(XX$, 4)) = NROOFL& Then
          Call REPLA(XX$, Chr$(2), 112, 1)
          Call GRAREG("ORFANUE", XX$, J&)
       End If
     Next J&
     '
     J& = 0
     For I& = 1 To ULTREG&("ORFANUL")
        XX$ = REGLEIDO$("ORFANUL", I&)
        If CVS(Left$(XX$, 4)) <> NROOFL& Then
           J& = J& + 1
           Call GRAREG("ORFANUL", XX$, J&)
        End If
     Next I&
     '
     For I& = J& + 1 To ULTREG&("ORFANUL")
        Call GRAREG("ORFANUL", String$(256, 0), I&)
     Next I&
     '
     For J& = 1 To ULTREG&("ORDEFABR")
       XX$ = REGLEIDO$("ORDEFABR", J&)
       If CVS(Left$(XX$, 4)) = NROOFL& Then
          Call REPLA(XX$, Chr$(2), 108, 1)
          Call GRAREG("ORDEFABR", XX$, J&)
       End If
     Next J&
     '
     Screen.MousePointer = 1
     Call CIERRARCH("*.*")
     GoTo 1405
     '

End Sub
'
Sub ORFPRI()
    '
    HO% = HOY(HOS$)
    LU$ = LUDAT$
    '
    TITU$ = "PROCESO DE CAMBIO DE PRIORIDAD DE ORDENES DE FABRICACION"

100  NROOF$ = ""
     '
1400 If INDI% = 0 Then
       '
       Screen.MousePointer = 11
       Call COPYSTRU("ORFANUE", "ORFANUL")
       J& = 0
       '
       For I& = 1 To ULTREG&("ORFANUE")
         XX$ = REGLEIDO$("ORFANUE", I&)
         ST% = Asc(Mid$(XX$, 112, 1))
         '
         If ST% < 3 Then
             Call REPLA(XX$, MKI$(0), 119, 2)
             Call GRAREG("ORFANUE", XX$, I&)
             J& = J& + 1
             Call GRAREG("ORFANUL", XX$, J&)
         End If
         '
       Next I&
       '
       If J& < 1 Then
          Call MENSERR(24, "No hay Ordenes de Fabricación\en Curso para Cambio de Prioridad.")
          Call CIERRARCH("ORFANUL")
          Screen.MousePointer = 1
          Exit Sub
       End If
       '
       For I& = J& + 1 To ULTREG&("ORFANUL")
         Call GRAREG("ORFANUL", String$(256, 0), I&)
       Next I&
       '
       Call CIERRARCH("ORFANUL")
       INDI% = 1
       '
       Screen.MousePointer = 1
       '
     End If
     
1405 Call FRESHECHO("ORFANUL")
     If ULTREG&("ORFANUL") < 1 Then INDI% = 0: Exit Sub
     Call SELECHO("ORFANUL/Ordenes de Trabajo para Cambio de Prioridad")
     '
1410 NROOFL& = Val(VALACT1$("ORFANUL"))
     If NROOFL& < 1 Then
         INDI% = 0
         Exit Sub
     End If
     '
     For I& = ULTREG&("ORFANUE") To 1 Step -1
       XX$ = REGLEIDO$("ORFANUE", I&)
       If CVS(Left$(XX$, 4)) = NROOFL& Then GoTo 1420
     Next I&
     '
     Call COMUNI("Inconsistencia en Base de Datos\de Ordenes de Fabricación.\  \Consulte al Soporte de Sistemas.")
     INDI% = 0
     Exit Sub
     '
1420 YY$ = OBJPLA$("ORFAPRI/Cambio Prioridad o Sección de O/F's", XX$)
     If YY$ = "" Then GoTo 1405
     '
     PRIO% = Asc(Mid$(YY$, 111, 1))
     If PRIO% > 9 Then
       PRIO% = PRIO% Mod 10
       Call REPLA(YY$, Chr$(PRIO%), 111, 1)
     End If
     SPXS$ = Mid$(YY$, 115, 4)
     '
     Screen.MousePointer = 11
     For J& = I& To 1 Step -1
       XX$ = REGLEIDO$("ORFANUE", J&)
       If CVS(Left$(XX$, 4)) = NROOFL& Then
          Call GRAREG("ORFANUE", YY$, J&)
       End If
     Next J&
     '
     For J& = 1 To ULTREG&("ORDEFABR")
       XX$ = REGLEIDO$("ORDEFABR", J&)
       If CVS(Left$(XX$, 4)) = NROOFL& Then
          Call REPLA(XX$, Chr$(PRIO%), 107, 1)
          Call REPLA(XX$, SPXS$, 113, 4)
          Call GRAREG("ORDEFABR", XX$, J&)
       End If
     Next J&
     '
     Screen.MousePointer = 1
     Call CIERRARCH("*.*")
     GoTo 1405
     '

End Sub
'
Sub ORFCIE()
     '
     HO% = HOY(HOS$)
     LU$ = LUDAT$
     '
     If PRECOCIE% = 0 Then
        PRECOCIE% = (-1)
        If CONTROL$("", "PRECOCIE") = "SI" Then
           PRECOCIE% = 1
        End If
     End If
     '
     Call ACTECORLAN    ' LISTA DE SELECCION DE ORDENES LANZADAS
     '
     TITU$ = "PROCESO DE CIERRE DE ORDENES DE FABRICACION"

100  Screen.MousePointer = 1
     NROOF$ = OBJPLA$("SELORFAP", "")
     If TRIM$(NROOF$) = "" Then
         Exit Sub
     End If
     '
     Screen.MousePointer = 11
     NORFA$ = AJUSTI$(NROOF$, 12)
     If TRIM$(ECOARCH$("ECORDEP", NORFA$)) = "" Then
         Call MENSERR(24, "Orden no Registrada")
         GoTo 100
     End If
     '
     For I& = ULTREG&("ORDEFABR") To 1 Step -1
       XX$ = REGLEIDO$("ORDEFABR", I&)
       If Mid$(XX$, 69, 12) = NORFA$ Then
          ROFANT$ = XX$
          CICONJU% = CVI(Mid$(XX$, 83, 2))
          '
          CAPRB# = 0
          RFF$ = RECFILT$("MOVISTO", 2, MKI$(CICONJU%))
          For KKU& = 1 To Len(RFF$) Step 4
             REMOX& = CVS(Mid$(RFF$, KKU&, 4))
             MVX$ = REGLEIDO$("MOVISTO", REMOX&)
             If TRIM$(Mid$(MVX$, 5, 16)) = TRIM$(NORFA$) Then   ' FILTRA POR NUMERO DE LOTE
                CAPRB# = CAPRB# + CVD(Mid$(MVX$, 88, 8))
             End If
          Next KKU&
          '
          CATOT# = CVD(Mid$(XX$, 89, 8))
          CAPRA# = CVD(Mid$(XX$, 119, 8))
          If Abs(CAPRA# - CAPRB#) >= 0.05 Then
             Call COMUNI("Diferencia de Saldo en O/F\Ajustada Automáticamente.")
             CAPRA# = CAPRB#
             Call REPLA(XX$, MKD$(CAPRA#), 119, 8)
             Call GRAREG("ORDEFABR", XX$, I&)
             ROFANT$ = XX$
          End If
          '
          CAPEN# = CATOT# - CAPRA#
          If CAPEN# < 0 Then CAPEN# = 0
          Call REPLA(XX$, MKD$(CAPEN#), 119, 8)
          Call REPLA(XX$, MKD$(0), 127, 8)
          Call REPLA(XX$, MKI$(HO%), 117, 2)
          REGOR& = I&
          GoTo 1420
       End If
     Next I&
     '
     Call COMUNI("Inconsistencia en Base de Datos\de Ordenes de Fabricación.\  \Consulte al Soporte de Sistemas.")
     INDI% = 0
     GoTo 100
     '
1420 YY$ = OBJPLA$("ORFACIE/Cierre Parcial de Ordenes de Fabricación", XX$ + Chr$(0) + Chr$(0))
     If YY$ = "" Then GoTo 100
     '
     NUVALIN& = 1
     For KU& = ULTREG&("MOVISTO") To 1 Step -1
       XYX$ = REGLEIDO$("MOVISTO", KU&)
       If Mid$(XYX$, 21, 2) = "OF" Then
         If Mid$(XYX$, 23, 2) = "OF" Then
           If Mid$(XYX$, 33, 2) = "VI" Then
             NUVALIN& = 1 + Val(Mid$(XYX$, 36, 7))
             GoTo 1421
           End If
         End If
       End If
     Next KU&
     '
1421 FECHA% = CVI(Mid$(YY$, 117, 2))
     FEX = FECHA%
     FECHACIOF$ = FECHATEX$(FEX)
     CI% = CVI(Mid$(YY$, 83, 2))
     CODCON$ = CODEXT$(CI%)
     DESCON$ = DESCRIT$(CI%)
     LOTE$ = NORFA$
     CMOV$ = "OF"
     DOPRI$ = NORFA$
     DOSEC$ = "VI-" + TRIM$(Str$(NUVALIN&))
     REFE$ = "Fabricación Propia"
     VUNI# = 0
     MONE$ = "$"
     NC% = 0
     DEPO% = Asc(Mid$(YY$, 257, 1))
     DP2% = Asc(Mid$(YY$, 258, 1))
     CANTI = CVD(Mid$(YY$, 119, 8))
     SCRAP = CVD(Mid$(YY$, 127, 8))
     '
     If CANTI < 0 Or SCRAP < 0 Then
        Call MENSERR(24, "Cantidades Incorrectas.")
        GoTo 100
     End If
     '
     XX$ = YY$
     If DEPO% < 1 Or TRIM$(ECOARCH$("ECODEP", Chr$(DEPO%))) = "" Then
        Call MENSERR(24, "Falta Depósito de Entrada")
        GoTo 1420
     End If
     '
     If DP2% < 1 Or TRIM$(ECOARCH$("ECODEP", Chr$(DP2%))) = "" Then
        Call MENSERR(24, "Falta Depósito de Consumo")
        GoTo 1420
     End If
     '
     If SCRAP >= 0.05 Then
       EPAC$ = UCase$(EMPREAC$)
       '\\\OT-05-0619-WAR-27/09/05///
       FIMATER& = ULTREG&("!MATEROF")
       If FIMATER& <= 1 Then
         If EXISTE%("RECHLI02.EXE") > 0 Then
           DSC% = 90 'DEPOSITO DE DEVOLUCIONES (***)
           'SE FORZO EL DEPOSITO PARA PASAR POR ALTO LAS VALIDACIONES
           'POSTERIORES SOBRE DSC%; DE TODAS MANERAS CUANDO
           'CUANDO HAY RECHAZO EN LINEA TOMA EL DEPOSITO
           'SELECCIONADO EN ESA OPERACIÓN.
           GoTo 1426
         End If
       End If
       '\\\OT-05-0619-WAR-FIN///
1425   Call SELECHO("TADEPOSI/Depósito Entrada Rechazos")
       DSC% = Val(VALACT1$("TADEPOSI"))
         If DSC% < 1 Then
           If InStr(EPAC$, "DOSIVA") > 0 Then
             Call MENSERR(24, "Debe Indicar Depósito\de Entrada de Rechazos.")
             GoTo 1425
           End If
         End If
     End If
     '
1426 CODICONJ$ = CODEXT$(CI%)
     CODPRO$ = CODICONJ$
     CICONJ% = CI%
     NORFABRI$ = NORFA$
     CACONJU# = CDbl(CANTI + SCRAP)
     VACONTROL% = 0: If PRECOCIE% = 1 Then VACONTROL% = 1
     '
     ' AQUI COMIENZA PRESENTACION DE DATOS SEGUN RESERVAS
     Screen.MousePointer = 11
     '
     VERISARRAS% = 0
     If CONTROL$("", "CHECOMPO") = "SI" Then
       VERISARRAS% = 1
     End If
     '
     NFX1$ = Mid$(NORFA$, 4)
     PPXX% = InStr(NFX1$, "-"): If PPXX% < 1 Then PPXX% = 1 + Len(NFX1$)
     NUORFA& = Val(Left$(NFX1$, PPXX% - 1))
     '
     JJ& = 0
     RFF$ = RECFILT$("RESERVA", 3, NORFA$)
     FINIL& = Len(RFF$)
     For UU& = 1 To FINIL& Step 4
       Call TRACE(20, UU&, FINIL&)
       I& = CVS(Mid$(RFF$, UU&, 4))
       YYR$ = REGLEIDO$("RESERVA", I&)
       If Mid$(YYR$, 3, 12) = NORFA$ Then
         If Mid$(YYR$, 19, 24) <> String$(24, 0) Then
           CAIT% = CAIT% + 1
           CIELE% = CVI(Left$(YYR$, 2))
           If CIELE% > 0 Then
             If CIELE% <= NUMAS% Then
               USOXI = CVS(Mid$(YYR$, 75, 4))
               If Abs(USOXI) > 0 Then
                 If VACONTROL% = 1 Then
                   If VERISARRAS% = 1 Then
                     Call VERISAL(CIELE%)
                     Call VESARRAS(CIELE%)
                   End If
                 End If
                 RERESX& = I&
                 ZZ$ = REGBLAN$("MATEROF")
                 Call REPLA(ZZ$, MKI$(CIELE%), 1, 2)
                 Call REPLA(ZZ$, UNIMED$(CIELE%), 3, 2)
                 NEBRU = CACONJU# * USOXI
                 Call REPLA(ZZ$, MKS$(NEBRU), 5, 4)
                 Call REPLA(ZZ$, MKD$(CDbl(USOXI)), 9, 8)
                 Call REPLA(ZZ$, MKS$(RERESX&), 17, 4)
                 STDIS = STOCKDIS(CIELE%) + STORESOF(CIELE%, NUORFA&)
                 Call REPLA(ZZ$, MKS$(STDIS), 21, 4)
                 FALTA = NEBRU - STDIS: If FALTA < 0 Then FALTA = 0
                 Call REPLA(ZZ$, MKS$(FALTA), 25, 4)
                   JJ& = JJ& + 1
                 Call GRAREG("!MATEROF", ZZ$, JJ&)
                 '
               End If
             End If
           End If
         End If
       End If
     Next UU&
     Call SETULTREG("!MATEROF", JJ&)
     Call CIERRARCH("!MATEROF")
     Screen.MousePointer = 1
     '
     If VACONTROL% = 1 Then
       TITUPA$ = TITUPA$ + ": "
       If TRIM$(NORFABRI$) <> "" Then TITUPA$ = TITUPA$ + TRIM$(NORFABRI$)
       If TRIM$(CODPRO$) <> "" Then TITUPA$ = TITUPA$ + " - " + TRIM$(CODPRO$) + " - " + TRIM$(DESCRIT$(CIXI%)) + " - Cant.=" + Str$(Int(CACONJU# + 0.5))
       '
       VTB% = VENTABU%("MATEROF/TITUPAN=" + TITUPA$)
       '
       If VTB% < 0 Then
         If COMALTER%("Realmente Quiere CANCELAR el Cierre de esta O/F ?\\No, Cerrarla\Si, Cancelar Cierre") <> 2 Then GoTo 1426
         Call CIERRARCH("*.*")
         Call LIBARCH("MOVISTO")
         Call LIBARCH("RESERVA")
         Call LIBARCH("ORDEFABR")
         Call LIBARCH("OPERFAB")
         Call LIBARCH("ECORDEP")
         GoTo 159           ' cancelacion
       End If
       '
       If Abs(CACONJU#) >= 0.05 Then
         For I& = 1 To ULTREG&("!MATEROF")
           ZZ$ = REGLEIDO$("!MATEROF", I&)
           CANUE# = CDbl(CVS(Mid$(ZZ$, 5, 4)))
           USNUE# = CANUE# / CACONJU#
           Call REPLA(ZZ$, MKD$(USNUE#), 9, 8)
           Call GRAREG("!MATEROF", ZZ$, I&)
         Next I&
       End If
     End If
     '
     ' FIN PRESENTACION DE DATOS SEGUN RESERVAS
     '
     CIERRA% = 0
     If CAPRA# + CANTI > CATOT# * 0.99 Then
          CIERRA% = 1
        Else
          MDXX% = COMALTER%("ATENCION: Esta O/F queda con\un Saldo Pendiente.\\Mantener Abierta\Cerrar")
          If MDXX% = 2 Then CIERRA% = 1
     End If
     '
     Call REPLA(YY$, MKD$(CAPRA# + CANTI), 119, 8)
     If CIERRA% = 1 Then
        Call REPLA(YY$, Chr$(3), 108, 1)
     End If
     '
     Screen.MousePointer = 11
     Call CIERRARCH("MOVISTO")
     Call BLOQARCH("ORDEFABR")
     Call BLOQARCH("MOVISTO")
     Call BLOQARCH("RESERVA")
     Call BLOQARCH("OPERFAB")
     Call BLOQARCH("ECORDEP")
     '
     If REGLEIDO$("ORDEFABR", REGOR&) <> ROFANT$ Then
       Call CIERRARCH("MOVISTO")
       Call CIERRARCH("ORDEFABR")
       Call CIERRARCH("RESERVA")
       Call CIERRARCH("OPERFAB")
       Call CIERRARCH("ECORDEP")
       Call MENSERR(24, "Imposible Grabar Cierre de esta Orden de Fabricación.\  \Causa Probable: Se Ha Registrado\un Cierre Parcial para esta Orden\desde otra Terminal de la Red.\  \Vuelva a Ingresar el Cierre de O/F.")
       GoTo 159
     End If
     '
     For U& = 1 To ULTREG&("!MATEROF")
       ZZ1$ = REGLEIDO$("!MATEROF", U&)
       RERES& = CVS(Mid$(ZZ1$, 17, 4))
       If RERES& < 1 Or RERES& > ULTREG&("RESERVA") Then GoTo 143
       If PRECOCIE% = 1 Then GoTo 143
       '
       ZZ2$ = REGLEIDO$("RESERVA", RERES&)
       If Mid$(ZZ2$, 3, 12) = NORFA$ Then
          If CVI(Left$(ZZ2$, 2)) = CVI(Left$(ZZ1$, 2)) Then
             GoTo 143
          End If
       End If
       '
       Call CIERRARCH("MOVISTO")
       Call CIERRARCH("ORDEFABR")
       Call CIERRARCH("RESERVA")
       Call CIERRARCH("OPERFAB")
       Call CIERRARCH("ECORDEP")
       Call MENSERR(24, "Imposible Grabar Cierre de esta Orden de Fabricación.\  \Causa Probable: Se Ha Registrado\un Cierre Parcial para esta Orden\desde otra Terminal de la Red.\  \Vuelva a Ingresar el Cierre de O/F.")
       GoTo 149
       '
143  Next U&
     '
     Call GRAREG("ORDEFABR", YY$, REGOR&)
     '
     If CANTI >= 0.05 Then
       DEPX% = DEPO%: If DEPO% = 99 Then DEPX% = DEPOPRE%(CICONJ%)
       Call MOVISTO(FECHA%, CICONJ%, LOTE$, CMOV$, DOPRI$, DOSEC$, REFE$, VUNI#, MONE$, NC%, DEPX%, CANTI)
     End If
     '
     FIMATER& = ULTREG&("!MATEROF")
     HUBORECHAZOS% = 0
     EPAC$ = TRIM$(UCase$(EMPREAC$))
     For U& = 1 To FIMATER&
       Call TRACE(20, U&, FIMATER&)
       ZZ$ = REGLEIDO$("!MATEROF", U&)
       CI2% = CVI(Left$(ZZ$, 2))
       CAN = (-1) * CVS(Mid$(ZZ$, 5, 4))
       '
       If TRIM$(UCase$(UNIMED$(CI2%))) = "GK" Then
          CAN = CAN / 1000
       End If
       '
       USOUNI# = CVD(Mid$(ZZ$, 9, 8))
       RERES& = CVS(Mid$(ZZ$, 17, 4))
       If CI2% > 0 Then
         If CI2% <= NUMAS% Then
           If Abs(CAN) >= 0.05 Then
             LOTE$ = ""
             CMOV$ = "CF"
             '
             DOPRI$ = NORFA$
             DOSEC$ = "PF-" + TRIM$(Str$(NUORFA&))
             REFE$ = "CONSUMO " + CODICONJ$
             If CAN > 0 Then
                REFE$ = "RECUPERO " + CODICONJ$
             End If
             '
             DEPX% = DP2%: If DP2% = 99 Then DEPX% = DEPOPRE%(CI2%)
             If STODEPOS(CI2%, DEPX%) < (-1) * CAN Then
               If InStr(EPAC$, "YALTRE") > 0 Then
                 Call CARDEPOS
                 For KKII% = 1 To NUSU%
                   If STODEPOS(CI2%, COSU%(KKII%)) >= (-1) * CAN Then
                     DEPX% = COSU%(KKII%)
                     GoTo 144
                   End If
                 Next KKII%
               End If
             End If
             '
144          CACONSUM = CAN
             ' ESTO ES PARA NO CONSIDERAR EN EL CONSUMIDO LO QUE LUEGO
             ' SE PROCESARA COMO RECHAZO EN LINEA
             If SCRAP >= 0.05 Then
               If DSC% > 0 Then
                 If FIMATER& <= 1 Then
                   If EXISTE%("RECHLI02.EXE") > 0 Then
                     CACONSUM = CAN + SCRAP * USOUNI#
                   End If
                 End If
               End If
             End If
             '
             Call MOVISTO(FECHA%, CI2%, LOTE$, CMOV$, DOPRI$, DOSEC$, REFE$, VUNI#, MONE$, NC%, DEPX%, CACONSUM)
             REGMOVI& = ULTREG&("MOVISTO")
           End If
           '
           If SCRAP >= 0.05 Then
             If DSC% > 0 Then
               '
               If FIMATER& <= 1 Then
                 If EXISTE%("RECHLI02.EXE") > 0 Then
                   CARECHAZA = SCRAP * USOUNI#
                   DOPRI12$ = NORFA$
                   Call PROCERECHA(CI2%, FECHA%, CARECHAZA, DOPRI12$) '\\\OT-06-0191-WAR-09/06/06///
                   HUBORECHAZOS% = 1
                   GoTo 146
                 End If
               End If
               '
               LOTE$ = ""
               CMOV$ = "SC"
               '
               DOPRI$ = NORFA$
               DOSEC$ = "PF-" + TRIM$(Str$(NUORFA&))
               REFE$ = "SCRAP " + CODICONJ$
               CANTUSCRA = SCRAP * USOUNI#
               If Abs(CANTUSCRA) >= 0.05 Then
                 DEPX% = DSC%: If DP2% = 99 Then DEPX% = DEPOPRE%(CI2%)
                 Call MOVISTO(FECHA%, CI2%, LOTE$, CMOV$, DOPRI$, DOSEC$, REFE$, VUNI#, MONE$, NC%, DEPX%, CANTUSCRA)
               End If
             End If
           End If
           '
146        If RERES& > 0 Then
             If RERES& <= ULTREG&("RESERVA") Then
               YYR$ = REGLEIDO$("RESERVA", RERES&)
               CACONSU# = CVD(Mid$(YYR$, 27, 8)) + CACONJU# * USOUNI#
               SALRESE# = CVD(Mid$(YYR$, 19, 8)) - CACONSU#
               If SALRESE# * Sgn(USOUNI) < 0 Then SALRESE# = 0
               Call REPLA(YYR$, MKD$(CACONSU#), 27, 8)
               Call REPLA(YYR$, MKD$(SALRESE#), 35, 8)
               Call REPLA(YYR$, MKS$(REGMOVI&), 61, 4)
               Call GRAREG("RESERVA", YYR$, RERES&)
             End If
           End If
           '
         End If
       End If
     Next U&
     '
     Call CIERRARCH("*.*")
     Call LIBARCH("RESERVA")
     Call LIBARCH("ORDEFABR")
     Call LIBARCH("MOVISTO")
     Call LIBARCH("OPERFAB")
     Call LIBARCH("ECORDEP")
     '
     Call CIERRARCH("MOVISTO")       ' EXPLÍCITO PARA QUE RECUPERE SALDOS
     '
     If CIERRA% = 1 Then             ' cancela operaciones
        '
        FEBA% = HOY(HOS$)
        HOBA% = Int(HORANUM(Time$))
        USUA% = USNBR% Mod 100
        DBAJA$ = MKI$(FEBA%) + MKI$(HOBA%) + Chr$(USUA%) + Chr$(9)
        DBAJA$ = DBAJA$ + "Cierre O/F      "
        '
        Call BLOQARCH("OPERFAB")
        Call BLOQARCH("ECORDEP")
        Call BLOQARCH("RESERVA")
        '
        For J& = 1 To ULTREG&("OPERFAB")
          XX$ = REGLEIDO$("OPERFAB", J&)
          NOFA$ = AJUSTI$(Left$(XX$, 12), 12)
          If NOFA$ = NORFA$ Then
             If Mid$(XX$, 177, 6) = String$(6, 0) Then
                Call REPLA(XX$, DBAJA$, 177, 22)
                Call GRAREG("OPERFAB", XX$, J&)
             End If
          End If
        Next J&
        Call CIERRARCH("OPERFAB")
        '
        JJ& = 0                      ' cancelar de ECORDEP
        For J& = 1 To ULTREG&("ECORDEP")
           XX$ = REGLEIDO$("ECORDEP", J&)
           NOFA$ = Left$(XX$, 12)
           If NOFA$ <> NORFA$ Then
              JJ& = JJ& + 1
              Call GRAREG("ECORDEP", XX$, JJ&)
           End If
        Next J&
        Call SETULTREG("ECORDEP", JJ&)
        '
        'JJ& = 0
        RFF$ = RECFILT$("RESERVA", 3, NORFA$)
        For UU& = 1 To Len(RFF$) Step 4
           J& = CVS(Mid$(RFF$, UU&, 4))
           XX$ = REGLEIDO$("RESERVA", J&)
           NOFA$ = Mid$(XX$, 3, 12)
           If NOFA$ = NORFA$ Then
              Call REPLA(XX$, String$(24, 0), 19, 24)
              Call GRAREG("RESERVA", XX$, J&)
           End If
        Next UU&
        Call CIERRARCH("RESERVA")
        '
        Call FRESHECHO("ECORDEP")
        '
     End If
     '
149 Screen.MousePointer = 1
    Call CIERRARCH("*.*")
    Call LIBARCH("RESERVA")
    Call LIBARCH("ORDEFABR")
    Call LIBARCH("MOVISTO")
    Call LIBARCH("OPERFAB")
    Call LIBARCH("ECORDEP")
    '
    '\\\OT-06-0209-WAR-07/06/06///
    'RUTINA DE ANOTACION
    NRO& = Val(Mid$(NORFA$, 4, 6))
    TTXYZ$ = Space$(76)
    HOII% = HOY(HOS$)
    Call REPLA(TTXYZ$, HOS$, 1, 8)
    Call REPLA(TTXYZ$, "Cierre (Parcial) de O/Fabricación", 11, 56)
    Call REPLA(TTXYZ$, USERTER$, 69, 24)
    Call REGAPP("ANOTAORF", MKS$(NRO&) + TTXYZ$)
    Call REGAPP("ANOTAORF", MKS$(NRO&) + Space$(88))
    TTXYZ$ = Space$(76)
    CA12$ = FORMATNUM(CANTI, "F12.1")
    Call REPLA(TTXYZ$, "     - Cantidad Aprobada =  " & CA12$, 11, 56)
    Call REGAPP("ANOTAORF", MKS$(NRO&) + TTXYZ$)
    TTXYZ$ = Space$(76)
    CARE1$ = FORMATNUM(SCRAP, "F12.1")
    Call REPLA(TTXYZ$, "     - Cantidad Rechazada = " & CARE1$, 11, 56)
    Call REGAPP("ANOTAORF", MKS$(NRO&) + TTXYZ$)
    Call REGAPP("ANOTAORF", MKS$(NRO&) + Space$(88))
    Call REGAPP("ANOTAORF", MKS$(NRO&) + String$(88, "="))
    Call CIERRARCH("ANOTAORF")
    '\\\OT-06-0209-WAR-FIN///
    '
    LU$ = LUDAT$
    If EXISTE%(LU$ + "OFAP-ARC.FRM") > 0 Then
       FECHDOC$ = FECHACIOF$
       NUMEDOC$ = TRIM$(Str$(NUVALIN&))
       TIPODOC$ = "Vale Aprobación Fabricación"
       DOCUORIG$ = NORFA$
       '
       CODPARAM$(1) = "COD-MAT": DOCPARAM$(1) = CODCON$
       CODPARAM$(2) = "DES-MAT": DOCPARAM$(2) = DESCON$
       CODPARAM$(3) = "CANTIDAD": DOCPARAM$(3) = TRIM$(Str$(CANTI))
       CAPARDOC% = 3
       '
       DESTIDOC% = 0
       CAITAB1% = 0
       '
       Call CIERRARCH("!OBSERVOF")
       Call BLANARCH("!OBSERVOF")
       VALIDA% = VENTABU%("OBSERVOF/TITUPAN=Cierre O-F Numero " + NORFA$)
       '
       Call PRINTDOC("@OFAP-ARC")   ' Remito de Facturacion
       '
    End If
    '
159 If HUBORECHAZOS% > 0 Then
      Call RECONECTA
      Call FINAL("")
    End If
    '
    GoTo 100
    '
End Sub
'

Sub ORFANU()
    '
    HO% = HOY(HOS$)
    LU$ = LUDAT$
    '
    TITU$ = "PROCESO DE ANULACION DE ORDENES DE FABRICACION"

100  NROOF$ = ""

1400 If INDI% = 0 Then
       '
       Screen.MousePointer = 11
       Call COPYSTRU("ECORDEF", "ECORANU")
       J& = 0: NORFANT$ = ""
       '
       For I& = 1 To ULTREG&("ORDEFABR")
         XX$ = REGLEIDO$("ORDEFABR", I&)
         ST% = Asc(Mid$(XX$, 108, 1))
         '
         If ST% < 3 Then     '    excluye cerradas
            If CVD(Mid$(XX$, 119, 8)) < CVD(Mid$(XX$, 89, 8)) Then
               YY$ = REGBLAN$("ECORANU")
               NORFA$ = Mid$(XX$, 69, 12)
               DORFA$ = Mid$(XX$, 5, 64)
               If NORFA$ <> NORFANT$ Then
                  NORFANT$ = NORFA$
                  Call REPLA(YY$, NORFA$, 1, 12)
                  Call REPLA(YY$, DORFA$, 13, 52)
                  J& = J& + 1
                  Call GRAREG("!ECORANU", YY$, J&)
               End If
            End If
         End If
       Next I&
       '
       If J& < 1 Then
          Call MENSERR(24, "No hay Ordenes de Fabricación\Abiertas que Puedan Anularse.")
          Call CIERRARCH("!ECORANU")
          Screen.MousePointer = 1
          Exit Sub
       End If
       '
       Call SETULTREG("!ECORANU", J&)
       Call CIERRARCH("!ECORANU")
       INDI% = 1
       '
       Screen.MousePointer = 1
       '
     End If
     
1405 Call FRESHECHO("!ECORANU")
     If ULTREG&("!ECORANU") < 1 Then INDI% = 0: Exit Sub
     Call SELECHO("!ECORANU/Ordenes de Trabajo para Anulación")
     '
1410 NORFA$ = AJUSTI$(VALACT1$("!ECORANU"), 12)
     If TRIM$(NORFA$) = "" Then
         INDI% = 0
         Exit Sub
     End If
     '
     NROOFL& = Val(Mid$(NORFA$, 4))
     For I& = ULTREG&("ORFANUE") To 1 Step -1
       XX$ = REGLEIDO$("ORFANUE", I&)
       If CVS(Left$(XX$, 4)) = NROOFL& Then GoTo 1420
     Next I&
     '
     Call COMUNI("Inconsistencia en Base de Datos\de Ordenes de Fabricación.\  \Consulte al Soporte de Sistemas.")
     INDI% = 0
     Exit Sub
     '
1420 YY$ = OBJPLA$("ORFALAN/Anulacion de Ordenes de Fabricación", XX$)
     If YY$ = "" Then GoTo 1405
     '
     Call BLOQARCH("ORFANUE")
     Call BLOQARCH("ORDEFABR")
     Call BLOQARCH("OPERFAB")
     Call BLOQARCH("RESERVA")
     Call BLOQARCH("DISTORFA")
     '
     Screen.MousePointer = 11
     For J& = 1 To ULTREG&("ORFANUE")
       XX$ = REGLEIDO$("ORFANUE", J&)
       If CVS(Left$(XX$, 4)) = NROOFL& Then
          Call REPLA(XX$, Chr$(9), 112, 1)
          Call GRAREG("ORFANUE", XX$, J&)
       End If
     Next J&
     Call CIERRARCH("ORFANUE")
     '
     J1& = 0
     For J& = 1 To ULTREG&("DISTORFA")
       XX$ = REGLEIDO$("DISTORFA", J&)
       If CVS(Left$(XX$, 4)) <> NROOFL& Then
         J1& = J1& + 1
         Call GRAREG("ORFANUE", XX$, J1&)
       End If
     Next J&
     Call SETULTREG("DISTORFA", J1&)
     Call CIERRARCH("DISTORFA")
     '
     For J& = 1 To ULTREG&("ORDEFABR")
       XX$ = REGLEIDO$("ORDEFABR", J&)
       If Mid$(XX$, 69, 12) = NORFA$ Then
          Call REPLA(XX$, Chr$(9), 108, 1)
          Call GRAREG("ORDEFABR", XX$, J&)
       End If
     Next J&
     Call CIERRARCH("ORDEFABR")
     '
     '\\\OT-06-0209-WAR-07/06/06///
     'RUTINA DE ANOTACION
     TTXYZ$ = Space$(76)
     NRO& = Mid$(NORFA$, 4, 6)
     HOII% = HOY(HOS$)
     Call REPLA(TTXYZ$, HOS$, 1, 8)
     Call REPLA(TTXYZ$, "Anulación de O/Fabricación", 11, 56)
     Call REPLA(TTXYZ$, USERTER$, 69, 24)
     Call REGAPP("ANOTAORF", MKS$(NRO&) + TTXYZ$)
     Call REGAPP("ANOTAORF", MKS$(NRO&) + Space$(88))
     Call REGAPP("ANOTAORF", MKS$(NRO&) + String$(88, "="))
     Call CIERRARCH("ANOTAORF")
     '\\\OT-06-0209-WAR-FIN///
     '
     FEBA% = HOY(HOS$)
     HOBA% = Int(HORANUM(Time$))
     USUA% = USNBR% Mod 100
     DBAJA$ = MKI$(FEBA%) + MKI$(HOBA%) + Chr$(USUA%) + Chr$(9)
     DBAJA$ = DBAJA$ + "Anulacion O/F   "
     '
' CANCELAR OPERACIONES PROGRAMADAS
     '
     For J& = 1 To ULTREG&("OPERFAB")
       XX$ = REGLEIDO$("OPERFAB", J&)
       NOFA$ = AJUSTI$(Left$(XX$, 12), 12)
       If NOFA$ = NORFA$ Then
          If Mid$(XX$, 177, 6) = String$(6, 0) Then
             Call REPLA(XX$, DBAJA$, 177, 22)
             Call GRAREG("OPERFAB", XX$, J&)
          End If
       End If
     Next J&
     Call CIERRARCH("OPERFAB")
     '
' CANCELAR RESERVAS
     '
    RFF$ = RECFILT$("RESERVA", 3, NORFA$)
    For UU& = 1 To Len(RFF$) Step 4
       J& = CVS(Mid$(RFF$, UU&, 4))
     'For J& = 1 To ULTREG&("RESERVA")
        XX$ = REGLEIDO$("RESERVA", J&)
        NOFA$ = Mid$(XX$, 3, 12)
        If NOFA$ = NORFA$ Then
          Call REPLA(XX$, String$(24, 0), 19, 24)
          Call GRAREG("RESERVA", XX$, J&)
        End If
     Next UU&
     Call CIERRARCH("RESERVA")
     '
     JJ& = 0
     NORFB$ = TRIM$(NORFA$)
     For J& = 1 To ULTREG&("!ECORANU")
       XX$ = REGLEIDO$("!ECORANU", J&)
       If TRIM$(Left$(XX$, 12)) <> NORFB$ Then
          JJ& = JJ& + 1
          Call GRAREG("!ECORANU", XX$, JJ&)
       End If
     Next J&
     Call SETULTREG("!ECORANU", JJ&)
     '
     Screen.MousePointer = 1
     Call CIERRARCH("*.*")
     GoTo 1405
     '
End Sub
'

Sub LANORF(NROOFL&, FLAN%)
     '
     Screen.MousePointer = 11
     APROLAN% = 0
     '
     FIN& = ULTREG&("ORFANUE")
     For J& = FIN& To 1 Step -1
       XX$ = REGLEIDO$("ORFANUE", J&)
       If CVS(Left$(XX$, 4)) = NROOFL& Then
          cix% = CVI(Mid$(XX$, 71, 2))
          CANORF = CVD(Mid$(XX$, 89, 8))
          If HAYDISTO%(cix%, NROOFL&, CANORF) <> 1 Then GoTo 1499
          Call REPLA(XX$, Chr$(1), 112, 1)
          Call REPLA(XX$, MKI$(FLAN%), 113, 2)
          Call GRAREG("ORFANUE", XX$, J&)
          GoTo 1510
       End If
     Next J&
     '
     Call MENSERR(24, "Lanzamiento Imposible.\Orden no Encontrada.")
     Exit Sub
     '
1510 SECUOR$ = ""
     For J& = 1 To ULTREG&("ORDEFABR")
       XX$ = REGLEIDO$("ORDEFABR", J&)
       If CVS(Left$(XX$, 4)) = NROOFL& Then
          Call REPLA(XX$, Chr$(1), 108, 1)
          Call REPLA(XX$, MKI$(FLAN%), 109, 2)
          Call GRAREG("ORDEFABR", XX$, J&)
          SECUOR$ = SECUOR$ + MKS$(J&)
       End If
     Next J&
     '
     Call ACTECOR(0)
     '
     For UUU% = 1 To Len(SECUOR$) Step 4
        REGORDE& = CVS(Mid$(SECUOR$, UUU%, 4))
        If REGORDE& > 0 Then
           If REGORDE& <= ULTREG&("ORDEFABR") Then
              XX$ = REGLEIDO$("ORDEFABR", REGORDE&)
              IDENLO$ = Mid$(XX$, 69, 12)
              CI% = CVI(Mid$(XX$, 83, 2))
              CAPENDE# = CVD(Mid$(XX$, 89, 8))
              DESTI$ = Mid$(XX$, 97, 8)
              REFE$ = Mid$(XX$, 5, 64)
              '
              Call BLANARCH("!OPERFAB")
              
              If CI% > 0 Then
                 If CI% <= NUMAS% Then
                    '
                    ' ELIMINAR OPERACIONES ANTERIORES
                    Screen.MousePointer = 11
                    For J& = 1 To ULTREG&("OPERFAB")
                       YY$ = REGLEIDO$("OPERFAB", J&)
                       NORFA$ = Left$(YY$, 12)
                       If NORFA$ = IDENLO$ Then
                          Call GRAREG("OPERFAB", String$(256, 0), J&)
                       End If
                    Next J&
                    Call CIERRARCH("OPERFAB")
                    '
                    CIJK% = CI%
                    NORFA$ = IDENLO$
                    GoSub GRASECOP
                    '
1480                NROFLLA$ = IDENLO$
                    Call GENRESER(NROFLLA$)
                    '
                 End If
              End If
           End If
        End If
1490 Next UUU%
     APROLAN% = 1
     '
1499 Screen.MousePointer = 1
     Call CIERRARCH("*.*")

Exit Sub
'
GRASECOP:
    Call BLANARCH("!OPERFAB")
    '
    PUNOPE$ = ""
    SECUOPSTA$ = RECFILT$("SECOPER", 1, MKI$(CIJK%))
    If Len(SECUOPSTA$) > 0 Then
       For TTI% = 1 To Len(SECUOPSTA$) Step 4
          TTL& = CVS(Mid$(SECUOPSTA$, TTI%, 4))
          SCP$ = REGLEIDO$("SECOPER", TTL&)
          If CVI(Left$(SCP$, 2)) = CIJK% Then
             NO% = CVI(Mid$(SCP$, 3, 2))
             DOPE$ = Mid$(SCP$, 5, 48)
             EQ1$ = Mid$(SCP$, 53, 6)
             EQAUX$ = Mid$(SCP$, 103, 18)
             LOTESTA = LOSTA: If LOTESTA < 1 Then LOTESTA = 1
             CALOTES = 1
             MINPREP = 60 * CVS(Mid$(SCP$, 59, 4)) * CALOTES
             MINFIJO = 60 * CVS(Mid$(SCP$, 63, 4)) * CALOTES
             MINVARI = CAPENDE# * CVS(Mid$(SCP$, 67, 4))
             MINDEMO = (MINFIJO + MINVARI) * CVS(Mid$(SCP$, 75, 4)) / 100
             MINLIMP = 60 * CVS(Mid$(SCP$, 79, 4)) * CALOTES
             MINLOTE = MINPREP + MINFIJO + MINVARI + MINDEMO + MINLIMP
             '
             HORAPROD = MINLOTE / 60
             CAOPS = CVS(Mid$(SCP$, 83, 4))
             '
' YA ESTA
             BVOF% = BINVAL%(AJUSTI$(NORFA$, 12))
             NUOROP% = 0
             FEAL% = HOY(HOS$)
             HOAL% = Int(HORANUM(Time$))
             USUA% = USNBR% Mod 100
             DALTA$ = Chr$(NUOROP%) + MKI$(FEAL%) + MKI$(HOAL%) + Chr$(USUA%)
             INFOBLA$ = Space$(6) + String$(26, 0) + Space$(16)
             PROGBLA$ = Space$(6) + String$(26, 0)
             '
             ZZ$ = REGBLAN$("OPERFAB")
             Call REPLA(ZZ$, NORFA$, 1, 12)
             Call REPLA(ZZ$, MKI$(CIJK%), 13, 2)
             Call REPLA(ZZ$, MKI$(NO%), 15, 2)
             Call REPLA(ZZ$, DOPE$, 17, 48)
             Call REPLA(ZZ$, MKI$(BVOF%), 65, 2)
             Call REPLA(ZZ$, DALTA$, 67, 6)
             Call REPLA(ZZ$, EQ1$, 73, 6)
             Call REPLA(ZZ$, MKD$(CAPENDE#), 79, 8)
             Call REPLA(ZZ$, MKD$(CAPENDE#), 95, 8)
             Call REPLA(ZZ$, MKS$(MINPREP), 103, 4)
             Call REPLA(ZZ$, MKS$(MINFIJO), 107, 4)
             Call REPLA(ZZ$, MKS$(MINVARI), 111, 4)
             Call REPLA(ZZ$, MKS$(MINDEMO), 115, 4)
             Call REPLA(ZZ$, MKS$(MINLIMP), 123, 4)
             Call REPLA(ZZ$, MKS$(MINLOTE), 131, 4)
             Call REPLA(ZZ$, String$(16, 0), 135, 16)
             Call REPLA(ZZ$, INFOBLA$, 151, 48)
             Call REPLA(ZZ$, MKS$(TTL&), 199, 4)
             Call REPLA(ZZ$, EQAUX$, 203, 18)
             Call REPLA(ZZ$, MKS$(CAOPS), 221, 4)
             Call REPLA(ZZ$, PROGBLA$, 225, 32)
             '
             Call REGAPP("!OPERFAB", ZZ$)
             PUNOPE$ = PUNOPE$ + MKS$(ULTREG&("!OPERFAB"))
             Call REGAPP("OPERFAB", ZZ$)
          End If
       Next TTI%
    End If
    '
    SECUOPSTB$ = RECFILT$("BCOMPLEM", 1, MKI$(CIJK%))
    If Len(SECUOPSTB$) > 0 Then
       For TTI% = 1 To Len(SECUOPSTB$) Step 4
          TTL& = CVS(Mid$(SECUOPSTB$, TTI%, 4))
          SCP$ = REGLEIDO$("BCOMPLEM", TTL&)
          If CVI(Left$(SCP$, 2)) = CIJK% Then
             NO% = 999
             DOPE$ = Mid$(SCP$, 3, 14)
             EQ1$ = Mid$(SCP$, 33, 6)
             EQAUX$ = String$(18, 0)
             LOTESTA = LOSTA: If LOTESTA < 1 Then LOTESTA = 1
             CALOTES = 1
             MINPREP = CVS(Mid$(SCP$, 65, 4))
             MINFIJO = 0
             MINVARI = CVS(Mid$(SCP$, 73, 4))
             MINDEMO = 0
             MINLIMP = 0
             MINLOTE = MINPREP + MINFIJO + MINVARI + MINDEMO + MINLIMP
             '
             If MINLOTE > 0.1 Then
                HORAPROD = MINLOTE / 60
                CAOPS = 1
                '
' YA ESTA
                BVOF% = BINVAL%(AJUSTI$(NORFA$, 12))
                NUOROP% = 0
                FEAL% = HOY(HOS$)
                HOAL% = Int(HORANUM(Time$))
                USUA% = USNBR% Mod 100
                DALTA$ = Chr$(NUOROP%) + MKI$(FEAL%) + MKI$(HOAL%) + Chr$(USUA%)
                INFOBLA$ = Space$(6) + String$(26, 0) + Space$(16)
                PROGBLA$ = Space$(6) + String$(26, 0)
                '
                ZZ$ = REGBLAN$("OPERFAB")
                Call REPLA(ZZ$, NORFA$, 1, 12)
                Call REPLA(ZZ$, MKI$(CIJK%), 13, 2)
                Call REPLA(ZZ$, MKI$(NO%), 15, 2)
                Call REPLA(ZZ$, DOPE$, 17, 48)
                Call REPLA(ZZ$, MKI$(BVOF%), 65, 2)
                Call REPLA(ZZ$, DALTA$, 67, 6)
                Call REPLA(ZZ$, EQ1$, 73, 6)
                Call REPLA(ZZ$, MKD$(CAPENDE#), 79, 8)
                Call REPLA(ZZ$, MKD$(CAPENDE#), 95, 8)
                Call REPLA(ZZ$, MKS$(MINPREP), 103, 4)
                Call REPLA(ZZ$, MKS$(MINFIJO), 107, 4)
                Call REPLA(ZZ$, MKS$(MINVARI), 111, 4)
                Call REPLA(ZZ$, MKS$(MINDEMO), 115, 4)
                Call REPLA(ZZ$, MKS$(MINLIMP), 123, 4)
                Call REPLA(ZZ$, MKS$(MINLOTE), 131, 4)
                Call REPLA(ZZ$, String$(16, 0), 135, 16)
                Call REPLA(ZZ$, INFOBLA$, 151, 48)
                Call REPLA(ZZ$, MKS$(TTL&), 199, 4)
                Call REPLA(ZZ$, EQAUX$, 203, 18)
                Call REPLA(ZZ$, MKS$(CAOPS), 221, 4)
                Call REPLA(ZZ$, PROGBLA$, 225, 32)
                '
                Call REGAPP("!OPERFAB", ZZ$)
                PUNOPE$ = PUNOPE$ + MKS$(ULTREG&("!OPERFAB"))
                Call REGAPP("OPERFAB", ZZ$)
             End If
          End If
       Next TTI%
    End If
    '
    Call CIERRARCH("OPERFAB")
Return
'
End Sub


Sub ORFTINT()
     '
     HO% = HOY(HOS$)
     LU$ = LUDAT$
     '
     Call ACTECOR(0)
     '
     TITU$ = "ORDEN DE DISPOSICION DE TINTORERIA"

100  Screen.MousePointer = 1
     OBX1$ = OBJPLA$("SELORTINTO", "")
     If TRIM$(OBX1$) = "" Then
         Exit Sub
     End If
     '
     NROOF$ = Left$(OBX1$, 12)
     PROVEE% = CVI(Mid$(OBX1$, 13, 2))
     '
     Screen.MousePointer = 11
     NORFA$ = AJUSTI$(NROOF$, 12)
     If TRIM$(ECOARCH$("ECORDEP", NORFA$)) = "" Then
         Call MENSERR(24, "Orden no Registrada como Pendiente")
         GoTo 100
     End If
     '
     If PROVEE% < 1 Or ECOARCH$("PROVED1", MKI$(PROVEE%)) = "" Then
         Call MENSERR(24, "Proveedor no Figura en Padrón")
         GoTo 100
     End If
     '
     For I& = ULTREG&("ORDEFABR") To 1 Step -1
       XX$ = REGLEIDO$("ORDEFABR", I&)
       If Mid$(XX$, 69, 12) = NORFA$ Then
          CAPREN# = CVD(Mid$(XX$, 89, 8))
          If CVD(Mid$(XX$, 119, 8)) > CAPREN# Then
              CAPREN# = CVD(Mid$(XX$, 119, 8))
          End If
          CIBAS% = CVI(Mid$(XX$, 83, 2))
          CODCON$ = CODEXT$(CIBAS%)
          DECON$ = DESCRIT0$(CIBAS%)
          REGOR& = I&
          GoTo 1420
       End If
     Next I&
     '
     Call COMUNI("Inconsistencia en Base de Datos\de Ordenes de Fabricación.\  \Consulte al Soporte de Sistemas.")
     INDI% = 0
     GoTo 100
     '
1420 Screen.MousePointer = 11
     NUOF& = Val(Mid$(NORFA$, 4))
     Call BLANARCH("!DISPOTIN")
     For U& = 1 To ULTREG&("ESTRUNUE")
        XX$ = REGLEIDO$("ESTRUNUE", U&)
        CIPRE% = CVI(Left$(XX$, 2))
        R2$ = Mid$(XX$, 3, 120) + REGLEIDO$("ESTRUDOS", U&)
        For K% = 1 To Len(R2$) Step 8
           DPSX$ = Mid$(R2$, K%, 2)
           If CVI(DPSX$) = CIBAS% Then
              YY$ = REGBLAN$("DISPOTIN")
              Call REPLA(YY$, MKS$(NUOF&), 1, 4)
              Call REPLA(YY$, MKI$(CIBAS%), 5, 2)
              Call REPLA(YY$, MKI$(CIPRE%), 7, 2)
              Call REPLA(YY$, COLOR$(CIPRE%), 17, 32)
              Call REPLA(YY$, TALLE$(CIPRE%), 49, 16)
              Call REPLA(YY$, MKI$(HO%), 129, 2)
              Call REPLA(YY$, MKI$(PROVEE%), 149, 2)
              Call REGAPP("!DISPOTIN", YY$)
              GoTo 1424
           End If
        Next K%
1424 Next U&
     '
     CADISOTPRO = 0
     For U& = 1 To ULTREG&("DISPOTIN")
        XX$ = REGLEIDO$("DISPOTIN", U&)
        If CVS(Left$(XX$, 4)) = NUOF& Then
           CAPAR = CVS(Mid$(XX$, 9, 4))
           If CVI(Mid$(XX$, 149, 2)) <> PROVEE% Or CVI(Mid$(XX$, 131, 2)) > 0 Then
              CADISOTPRO = CADISOTPRO + CAPAR
              GoTo 1427
           End If
           CIPRE% = CVI(Mid$(XX$, 7, 2))
           For V& = 1 To ULTREG&("!DISPOTIN")
              YY$ = REGLEIDO$("!DISPOTIN", V&)
              If Left$(YY$, 8) = Left$(XX$, 8) Then
                 Call GRAREG("!DISPOTIN", XX$, V&)
              End If
           Next V&
        End If
1427 Next U&
     '
     Call CIERRARCH("!DISPOTIN")
     Screen.MousePointer = 1
     '
1430 TITU$ = " - " + TRIM$(CSTRING$(MKD$(CAPREN#), 3, 8, 0, 2)) + " Prendas " + TRIM$(CODCON$) + " - " + DECON$
     VTB% = VENTABU%("DISPOTIN/TITUPAN=DISPOSICION PARA TINTORERIA" + TITU$)
     If VTB% < 0 Then GoTo 100
     '
     CADISPO = 0
     For U& = 1 To ULTREG&("!DISPOTIN")
        XX$ = REGLEIDO$("!DISPOTIN", U&)
        CAPAR = CVS(Mid$(XX$, 9, 4))
        CIPRE% = CVI(Mid$(XX$, 7, 2))
        KILOS = CAPAR * PESUNI(CIPRE%)
        Call REPLA(XX$, MKS$(KILOS), 13, 4)
        Call GRAREG("!DISPOTIN", XX$, U&)
        CADISPO = CADISPO + CAPAR
     Next U&
     '
     If CADISPO + CADISOTPRO > CAPREN# + 0.5 Then
        OPXX% = COMALTER%("ATENCION: El Total Dispuesto (" + TRIM$(Str$(Int(CADISPO + CADISOTPRO + 0.5))) + ")\Excede la Cantidad Según O/F (" + TRIM$(Str$(Int(CAPREN# + 0.5))) + ")\\Corregir\Continuar")
        If OPXX% <> 2 Then GoTo 1430
     End If
     '
     J& = 0
     For U& = 1 To ULTREG&("DISPOTIN")
        XX$ = REGLEIDO$("DISPOTIN", U&)
        If CVS(Left$(XX$, 4)) <> NUOF& Or CVI(Mid$(XX$, 149, 2)) <> PROVEE% Or CVI(Mid$(XX$, 131, 2)) > 0 Then
           J& = J& + 1
           Call GRAREG("DISPOTIN", XX$, J&)
        End If
     Next U&
     '
     For U& = 1 To ULTREG&("!DISPOTIN")
        XX$ = REGLEIDO$("!DISPOTIN", U&)
        If CVS(Mid$(XX$, 9, 4)) > 0 Then
           J& = J& + 1
           Call GRAREG("DISPOTIN", XX$, J&)
        End If
     Next U&
     '
     Call CIERRARCH("*.*")
     Call COMUNI("Disposición Completa")
     GoTo 100
     '
End Sub
'

Sub REPRINORFA()
    '
    Call ORFNUE(2)

End Sub
'

Sub REMITINTO()
     '
     Static DATRANSPO$
     Dim CPREN%(1024), CADIS(1024)
     Dim CCRU%(512), CACRU(512), REGMAC&(512)
     '
     HO% = HOY(HOS$)
     LU$ = LUDAT$
     '
50   REGXX& = ULTREG&("OCOMPRA")
     NRO& = 1
     If REGXX& > 0 Then
        NRO& = 1 + CVS(Left$(REGLEIDO$("OCOMPRA", REGXX&), 4))
     End If
     If NRO& <= REGXX& Then
        Screen.MousePointer = 11
        For U& = 1 To REGXX&
           X0$ = REGLEIDO$("OCOMPRA", U&)
           If CVS(Left$(X0$, 4)) >= NRO& Then
              NRO& = 1 + CVS(Left$(X0$, 4))
           End If
        Next U&
        Screen.MousePointer = 1
     End If
     '
100  Screen.MousePointer = 1
     EBOL$ = OBJPLA$("SELPROVE", "")
     If EBOL$ = "" Then
         Exit Sub
     End If
     '
     PROVEE% = CVI(EBOL$)
     If PROVEE% < 1 Or ECOARCH$("PROVED1", MKI$(PROVEE%)) = "" Then
        Call MENSERR(24, "Proveedor no Figura en Padrón")
        GoTo 100
     End If
     '
     Screen.MousePointer = 11
     '
     FECHDOC$ = HOS$
     NUMEDOC$ = TRIM$(Str$(NRO&))
     While Len(NUMEDOC$) < 6
        NUMEDOC$ = "0" + NUMEDOC$
     Wend
     TIPODOC$ = "Orden de Disposicion"
     '
     CODTABU1$(1) = "COD-MAT"
     CODTABU1$(2) = "TALLE"
     CODTABU1$(3) = "COLOR"
     CODTABU1$(4) = "CANTIDAD"
     CODTABU1$(5) = "CAKILOS"
     CACOLTAB1% = 5
     '
     DESTIDOC% = (-1) * PROVEE%
     CAITAB1% = 0
     CAITCRU% = 0
     '
     For U& = 1 To ULTREG&("DISPOTIN")
        XX$ = REGLEIDO$("DISPOTIN", U&)
        If CVI(Mid$(XX$, 149, 2)) = PROVEE% Then
          If CVI(Mid$(XX$, 131, 2)) <= 0 Then
            CAPORCOL = CVS(Mid$(XX$, 9, 4))
            If CAPORCOL > 0 Then
              CICRU% = CVI(Mid$(XX$, 5, 2))
              CIPRE% = CVI(Mid$(XX$, 7, 2))
              TALLE1$ = Mid$(XX$, 49, 16)
              COCOL1$ = Mid$(XX$, 17, 32)
              CACOL1$ = Str$(CAPORCOL)
              KICOL1$ = Str$(CVS(Mid$(XX$, 13, 4)))
              '
              Call REPLA(XX$, MKI$(HO%), 131, 2)
              Call REPLA(XX$, MKS$(NRO&), 133, 4)
              Call GRAREG("DISPOTIN", XX$, U&)
              '
              CAITAB1% = CAITAB1% + 1
              TETABU1$(1, CAITAB1%) = CODEXT$(CIPRE%)
              TETABU1$(2, CAITAB1%) = TALLE1$
              TETABU1$(3, CAITAB1%) = COCOL1$
              TETABU1$(4, CAITAB1%) = CACOL1$
              TETABU1$(5, CAITAB1%) = KICOL1$
              '
              CPREN%(CAITAB1%) = CIPRE%
              CADIS(CAITAB1%) = CAPORCOL
              '
              For J% = 1 To CAITCRU%
                 If CCRU%(J%) = CICRU% Then
                    CACRU(J%) = CACRU(J%) + CAPORCOL
                    GoTo 109
                 End If
              Next J%
              CAITCRU% = CAITCRU% + 1
              CCRU%(CAITCRU%) = CICRU%
              CACRU(CAITCRU%) = CAPORCOL
              '
            End If
          End If
        End If
109  Next U&
     '
     If CAITAB1% < 1 Then
        Call MENSERR(24, "No hay Disposiciones Pendientes\para " + TRIM$(RASOCLI$((-1) * PROVEE%)) + ".")
        GoTo 100
     End If
     '
     Call PRINTDOC("FORDITIN")   ' Orden de Disposicion
     '
    'FENTRE% = CVI(Mid$(X1$, 9, 2))
    'If FENTRE% < FEMI% Then
    '    Call MENSERR(24, "Fecha de Entrega Incoherente")
    '    Call REPLA(X0$, MKI$(FEMI%), 9, 2)
    '    GoTo 101
    'End If
    '
     X1$ = REGBLAN$("OCOMPRA")
     Call REPLA(X1$, MKS$(NRO&), 1, 4)
     Call REPLA(X1$, MKI$(HO%), 5, 2)
     Call REPLA(X1$, MKI$(PROVEE%), 7, 2)
     Call REPLA(X1$, MKI$(HO%), 9, 2) ' fecha de entrega)
     Call REPLA(X1$, Chr$(2), 57, 1)
     Call REPLA(X1$, Chr$(0), 62, 1)  ' abierta
     '
     For KKI% = 1 To CAITAB1%
        Call REPLA(X1$, MKI$(CPREN%(KKI%)), 11, 2)
        Call REPLA(X1$, MKS$(CADIS(KKI%)), 13, 4)
        Call REGAPP("OCOMPRA", X1$)
     Next KKI%
     '
     NOCOM$ = TRIM$(Str$(NRO&))
     While Len(NOCOM$) < 6: NOCOM$ = "0" + NOCOM$: Wend
     NOCOM$ = "OC-" + Left$(NOCOM$, 6) + "-00"
     '
     For I% = 1 To CAITCRU%
        YY$ = REGBLAN$("MACONSIG")
        Call REPLA(YY$, MKI$(HO%), 1, 2)
        Call REPLA(YY$, MKI$(PROVEE%), 3, 2)
        Call REPLA(YY$, NOCOM$, 21, 12)
        '
        CII% = CCRU%(I%)
        CANTI# = CDbl(CACRU(I%))
        REGMAC&(I%) = 0
        '
        If TRIM$(CODEXT$(CII%)) <> "" Then
           If CANTI# > 0.05 Then
              Call REPLA(YY$, MKI$(CII%), 33, 2)
              Call REPLA(YY$, UNIMED$(CII%), 51, 4)
              Call REPLA(YY$, MKD$(CANTI#), 55, 8)
              Call REGAPP("MACONSIG", YY$)
              REGMAC&(I%) = ULTREG&("MACONSIG")
           End If
        End If
     Next I%
     '
' REMITO EN CONSIGNACION ENTRECONSI
     CMO$ = TRIM$(CONTROL$("", "CODRECON"))
     If CMO$ = "" Then CMO$ = "RC"
     FIN& = ULTREG&("MOVISTO")
     For U& = FIN& To 1 Step -1
        Call TRACE(20, FIN& - U& + 1, FIN&)
        XX$ = REGLEIDO$("MOVISTO", U&)
        If UCase$(Mid$(XX$, 21, 2)) = CMO$ Then
           UNRO& = Val(Mid$(XX$, 26, 7))
              GoTo 210
        End If
     Next U&
     UNRO& = 0
     '
210  FECHDOC$ = HOS$
     NUMEDOC$ = TRIM$(Str$(1 + UNRO&))
     While Len(NUMEDOC$) < 6
        NUMEDOC$ = "0" + NUMEDOC$
     Wend
     NUMEDOC$ = CMO$ + "-" + NUMEDOC$
     REMI$ = NUMEDOC$
     TIPODOC$ = "Remito de Traslado"
     '
     DTP$ = OBJPLA$("TRANSPOR", DATRANSPO$)
     DATRANSPO$ = DTP$
     '
     CODPARAM$(1) = "TRANSPOR": DOCPARAM$(1) = TRIM$(Mid$(DTP$, 1, 32))
     CODPARAM$(2) = "DOMI-TRA": DOCPARAM$(2) = TRIM$(Mid$(DTP$, 33, 32))
     CODPARAM$(3) = "CUIT-TRA": DOCPARAM$(3) = TRIM$(Mid$(DTP$, 65, 16))
     CODPARAM$(4) = "PATE-TRA": DOCPARAM$(4) = TRIM$(Mid$(DTP$, 81, 16))
     CODPARAM$(5) = "COND-TRA": DOCPARAM$(5) = TRIM$(Mid$(DTP$, 97, 32))
     CODPARAM$(6) = "DOCUCOND": DOCPARAM$(6) = TRIM$(Mid$(DTP$, 129, 24))
     CAPARDOC% = 6
     '
     CODTABU1$(1) = "COD-MAT"
     CODTABU1$(2) = "DES-MAT"
     CODTABU1$(3) = "CANTIDAD"
     CODTABU1$(4) = "CAKILOS"
     CACOLTAB1% = 4
     '
     DESTIDOC% = (-1) * PROVEE%
     CAITAB1% = 0
     '
     For U& = 1 To CAITCRU%
        '
        CAITAB1% = CAITAB1% + 1
        TETABU1$(1, CAITAB1%) = CODEXT$(CCRU%(U&))
        TETABU1$(2, CAITAB1%) = DESCRIT$(CCRU%(U&))
        TETABU1$(3, CAITAB1%) = Str$(CACRU(U&))
        TETABU1$(4, CAITAB1%) = Str$(CACRU(U&) * PESUNI(CCRU%(U&)))
        '
        If REGMAC&(U&) > 0 Then
           ZZ$ = REGLEIDO$("MACONSIG", REGMAC&(U&))
           DOSEC$ = Mid$(ZZ$, 21, 12)
           TOENTRE# = CVD(Mid$(ZZ$, 63, 8)) + CACRU(U&)
           Call REPLA(ZZ$, MKD$(TOENTRE#), 63, 8)
           If HO% > CVI(Mid$(ZZ$, 71, 2)) Then
              Call REPLA(ZZ$, MKI$(HO%), 71, 2)
           End If
           Call GRAREG("MACONSIG", ZZ$, REGMAC&(U&))
           '
           LOTE$ = NOCOM$
           Call MOVISTO(HO%, CCRU%(U&), LOTE$, CMO$, REMI$, DOSEC$, RASOCLI$((-1) * PROVEE%), 0, "$", (-1) * PROVEE%, 0, (-1) * CACRU(U&))
           '
        End If
219  Next U&
     Call CIERRARCH("MOVISTO")
     '
     Call PRINTDOC("FOREMICO")   ' Remito en Consignacion
     Call CIERRARCH("*.*")
     GoTo 50
     '
End Sub

Function HAYDISTO%(cix%, NORFAI&, CACONJU)
     '
     HDS% = 0
     If CHECOMPO% = 0 Then
        CHECOMPO% = (-1)
        If CONTROL$("", "CHECOMPO") = "SI" Then
          If CONTROL$("", "PRECOLAN") <> "SI" Then
            CHECOMPO% = 1
            ' PRESENTA MENSAJE DE ADVERTENCIA SI
            '   VERIFICA COMPONENTES Y NO LOS PRESENTA AL LANZAMIENTO
          End If
        End If
     End If
     '
     If CHECOMPO% <> 1 Then
        HAYDISTO% = 1
        Exit Function
     End If
     '
     If cix% < 1 Or cix% > NUMAS% Then
        HAYDISTO% = 1
        Exit Function
     End If
     '
     COD$ = CODEXT$(cix%)
     CODCON$ = COD$
     Call LEEEXPLOMUL(COD$, 1, 1)
     For II% = 1 To CAIT%
        CI2% = CIJ%(II%)
        CAN = CACONJU * USOI(II%)
        STDIS = STOCKDIS(CI2%)
        STROF = STORESOF(CI2%, NORFAI&)
        If CAN > STDIS + STROF Then
           MENSA$ = "El Stock Disponible (" + TRIM$(CODEXT$(CI2%)) + ") es Insuficiente\para Lanzar Esta Orden de Fabricación."
           If COMALTER%(MENSA$ + "\\Lanzar Igual\Cancelar") = 1 Then
              HDS% = 1
           End If
           GoTo 99
        End If
     Next II%
     HDS% = 1
     '
99   HAYDISTO% = HDS%
     '
End Function
'

Sub GENRESER(NROFLLA$)
     '
10   NROOF$ = NROFLLA$
     HO% = HOY(HOS$)
     '
     If PRECOLAN% = 0 Then
        PRECOLAN% = (-1)
        If CONTROL$("", "PRECOLAN") = "SI" Then
           PRECOLAN% = 1
        End If
     End If
     '
     If TRIM$(NROOF$) <> "" Then
        NORFA$ = AJUSTI$(NROOF$, 12)
        If TRIM$(ECOARCH$("ECORDEP", NORFA$)) <> "" Then
            GoTo 120
        End If
     End If
     NROOF$ = ""
     '
100  Screen.MousePointer = 1
     NROOF$ = OBJPLA$("SELORFAP", "")
     If TRIM$(NROOF$) = "" Then
         Exit Sub
     End If
     '
110  Screen.MousePointer = 11
     NORFA$ = AJUSTI$(NROOF$, 12)
     If TRIM$(ECOARCH$("ECORDEP", NORFA$)) = "" Then
         Call MENSERR(24, "Orden no Registrada")
         GoTo 100
     End If
     '
120  For I& = ULTREG&("ORDEFABR") To 1 Step -1
       XX$ = REGLEIDO$("ORDEFABR", I&)
       If Mid$(XX$, 69, 12) = NORFA$ Then
          CATOT# = CVD(Mid$(XX$, 89, 8))
          CAPRA# = CVD(Mid$(XX$, 119, 8))
          CAPEN# = CATOT# - CAPRA#
          If CAPEN# < 0 Then CAPEN# = 0
          Call REPLA(XX$, MKD$(CAPEN#), 119, 8)
          Call REPLA(XX$, MKD$(0), 127, 8)
          Call REPLA(XX$, MKI$(HO%), 117, 2)
          REGOR& = I&
          GoTo 1410
       End If
     Next I&
     '
     Call COMUNI("Inconsistencia en Base de Datos\de Ordenes de Fabricación.\  \Consulte al Soporte de Sistemas.")
     INDI% = 0
     GoTo 100
     '
1410 If CAPEN# < 0.1 Then
         Call MENSERR(24, "No hay Pendiente para la O/F '" + TRIM$(NORFA$) + "'.")
     End If
     '
     CIXI% = CVI(Mid$(XX$, 83, 2))
     CICONJ% = CI%
     CODPRO$ = CODEXT$(CIXI%)
     CODICONJ$ = CODPRO$
     NORFABRI$ = NORFA$
     CACONJU# = CAPEN#
     VACONTROL% = 0: If PRECOLAN% = 1 Then VACONTROL% = 1
     '
1412 Call CARCOMPO(CODICONJ$, NORFABRI$, CACONJU#, VACONTROL%, 2)
     If VACONTROL% < 0 Then
        GoTo 1499
     End If
     '
1415 Call BLOQARCH("RESERVA")
     JJ& = ULTREG&("RESERVA")
     RFF$ = RECFILT$("RESERVA", 3, NORFA$)
     For UU& = 1 To Len(RFF$) Step 4
        I& = CVS(Mid$(RFF$, UU&, 4))
        YY$ = REGLEIDO$("RESERVA", I&)
        If Mid$(YY$, 3, 12) = NORFA$ Then
            Call REPLA(YY$, String$(96, 0), 1, 96)
            Call GRAREG("RESERVA", YY$, I&)
        End If
     Next UU&
     '
     For U& = 1 To ULTREG&("!MATEROF")
        ZZ$ = REGLEIDO$("!MATEROF", U&)
        CIIX% = CVI(Left$(ZZ$, 2))
        CANTX# = CVS(Mid$(ZZ$, 5, 4))
        USOII = 0: If Abs(CAPEN#) >= 0.05 Then USOII = CANTX# / CAPEN#
        '
        If Abs(CANTX#) >= 0.05 Then
           YY$ = REGBLAN$("RESERVA")
           Call REPLA(YY$, MKI$(CIIX%), 1, 2)
           Call REPLA(YY$, NORFA$, 3, 12)
           Call REPLA(YY$, MKI$(CIIX%), 15, 2)
           Call REPLA(YY$, MKI$(HO%), 17, 2)
           Call REPLA(YY$, MKD$(CANTX#), 19, 8)
           Call REPLA(YY$, MKD$(CANTX#), 35, 8)
           Call REPLA(YY$, MKS$(USOII), 75, 4)
           JJ& = JJ& + 1
           Call GRAREG("RESERVA", YY$, JJ&)
        End If
     Next U&
     '
     Call CIERRARCH("RESERVA")
     Call LIBARCH("RESERVA")
     '
1499 If TRIM$(NROFLLA$) = "" Then
        GoTo 100
     End If
     '
End Sub
'

Sub REGENRESER()
   '
   If DERACCE%("REGERESE", "Re-Generación de Reservas según O/F") < 2 Then
      Exit Sub
   End If
   '
   Dim NOFACTI$(4096)
   CANOF% = 0: J& = 0
   '
   Screen.MousePointer = 11
   FIN& = ULTREG&("ORDEFABR")
   For I& = 1 To FIN&
      Call TRACE(20, I&, FIN&)
      XX$ = REGLEIDO$("ORDEFABR", I&)
      ST% = Asc(Mid$(XX$, 108, 1))
      If ST% < 3 Then
         CATOT# = CVD(Mid$(XX$, 89, 8))
         CAPRA# = CVD(Mid$(XX$, 119, 8))
         CAPEN# = CATOT# - CAPRA#
         If CAPEN# > 0.05 Then
            NOFANT$ = TRIM$(Mid$(XX$, 69, 12))
            If TRIM$(NOFANT$) <> "" Then
               CANOF% = CANOF% + 1
               NOFACTI$(CANOF%) = NOFANT$
               X$ = REGLEIDO$("ECORDEF", I&)
               J& = J& + 1
               Call GRAREG("ECORDEP", X$, J&)
            End If
         End If
      End If
   Next I&
   Call SETULTREG("RESERVA", 0)
   Call CIERRARCH("RESERVA")
   Call SETULTREG("ECORDEP", J&)
   Call CIERRARCH("ECORDEP")
   '
   FIN& = CANOF%
   J& = 0
   '
   For I& = 1 To FIN&
      Call TRACE(20, I&, FIN&)
      '
      For I1& = I& + 1 To FIN&                   ' ORDENAR
         If NOFACTI$(I&) > NOFACTI$(I1&) Then
            NOFAC$ = NOFACTI$(I&)
            NOFACTI$(I&) = NOFACTI$(I1&)
            NOFACTI$(I1&) = NOFAC$
         End If
      Next I1&
      '
      Call GENRESER(NOFACTI$(I&))
   Next I&
   Call CIERRARCH("RESERVA")
   Call CIERRARCH("*.*")
   '
   Call ACTECOR(0)
   Call CIERRARCH("*.*")
   '
   Screen.MousePointer = 1
   '
End Sub
'

Sub CHECKRES()
   '
   If DERACCE%("CHECKRES", "Control de Reservas según O/Fs en Curso") < 2 Then
      Exit Sub
   End If
   '
   If COMALTER%("Para Ejecutar Este Proceso\TODOS LOS USUARIOS deben SALIR DEL SISTEMA\\Cancelar\Continuar") <> 2 Then
      Exit Sub
   End If
   '
   Call BLOQARCH("RESERVA")
   Dim NOFACTI$(4096)
   CANOF% = 0: J& = 0
   '
   Screen.MousePointer = 11
   FIN& = ULTREG&("ORDEFABR")
   For I& = 1 To FIN&
      Call TRACE(20, I&, FIN&)
      XX$ = REGLEIDO$("ORDEFABR", I&)
      ST% = Asc(Mid$(XX$, 108, 1))
      If ST% < 3 Then
         CATOT# = CVD(Mid$(XX$, 89, 8))
         CAPRA# = CVD(Mid$(XX$, 119, 8))
         CAPEN# = CATOT# - CAPRA#
         If CAPEN# > 0.05 Then
            NOFANT$ = TRIM$(Mid$(XX$, 69, 12))
            If TRIM$(NOFANT$) <> "" Then
               CANOF% = CANOF% + 1
               NOFACTI$(CANOF%) = NOFANT$
               X$ = REGLEIDO$("ECORDEF", I&)
               J& = J& + 1
               Call GRAREG("ECORDEP", X$, J&)
            End If
         End If
      End If
   Next I&
   '
   Call SETULTREG("ECORDEP", J&)
   Call CIERRARCH("ECORDEP")
   '
   FIN& = ULTREG&("RESERVA")
   J& = 0: NOFAN$ = ""
   '
   For I& = 1 To FIN&
      Call TRACE(20, I&, FIN&)
      X$ = REGLEIDO$("RESERVA", I&)
      NOFARES$ = TRIM$(Mid$(X$, 3, 12))
      '
      For IOF% = 1 To CANOF%
        If NOFACTI$(IOF%) = NOFARES$ Then
          J& = J& + 1
          Call GRAREG("RESERVA", X$, J&)
          GoTo 19
        End If
      Next IOF%
      '
      If NOFARES$ <> NOFAN$ Then
        NOFAN$ = NOFARES$
        Call COMUNI("Canceladas Reservas para " + NOFARES$)
      End If
      '
19 Next I&
   Call SETULTREG("RESERVA", J&)
   Call CIERRARCH("RESERVA")
   '
   Call ACTECOR(0)
   Call CIERRARCH("*.*")
   '
   Screen.MousePointer = 1
   '
End Sub
'
Sub ANURESER(NROFLLA$)
     '
     NROOF$ = NROFLLA$
     HO% = HOY(HOS$)
     '
     If TRIM$(NROOF$) = "" Then
       If DERACCE%("ANURESE", "Anulación de Reservas de Materiales") < 2 Then
         Exit Sub
       End If
     End If
     '
     If TRIM$(NROOF$) <> "" Then
        NORFA$ = AJUSTI$(NROOF$, 12)
        If TRIM$(ECOARCH$("ECORDEF", NORFA$)) <> "" Then
            GoTo 120
        End If
     End If
     NROOF$ = ""
     '
100  Screen.MousePointer = 1
     Call SELECHO("ECORDEF/Anulacion de Reservas segun O/F")
     NROOF$ = TRIM$(VALACT1$("ECORDEF"))
     If TRIM$(NROOF$) = "" Then
         Exit Sub
     End If
     '
110  NORFA$ = AJUSTI$(NROOF$, 12)
     If TRIM$(ECOARCH$("ECORDEF", NORFA$)) = "" Then
         Call MENSERR(24, "Orden no Registrada")
         GoTo 100
     End If
     '
     If COMALTER%("Confirma la Anulación de los Materiales Reservados para '" + NORFA$ + "' ?\\No, Conservar Reservas\Si, Anular Reservas") = 2 Then
        Call BLOQARCH("RESERVA")
120     Screen.MousePointer = 11
        RFF$ = RECFILT$("RESERVA", 3, NORFA$)
        For KKUU& = 1 To Len(RFF$) Step 4
           I& = CVS(Mid$(RFF$, KKUU&, 4))
           YY$ = REGLEIDO$("RESERVA", I&)
           If TRIM$(Mid$(YY$, 3, 12)) <> "" Then
              If Mid$(YY$, 3, 12) = NORFA$ Then
                 Call REPLA(YY$, String$(24, 0), 19, 24)
                 Call GRAREG("RESERVA", YY$, I&)
              End If
           End If
        Next KKUU&
        Call CIERRARCH("RESERVA")
        Screen.MousePointer = 1
     End If
     '
     If TRIM$(NROFLLA$) = "" Then
        GoTo 100
     End If
     '
End Sub
'

Sub ACTECOR(MODOI%)
    '
    Screen.MousePointer = 11
    Call COPYSTRU("ECORDEF", "ECORDEP")
    J& = 0
    '
    FIN& = ULTREG&("ORDEFABR")
    For I& = 1 To FIN&
        Call TRACE(20, I&, FIN&)
        F$ = REGLEIDO$("ORDEFABR", I&)
        CANOF# = CVD(Mid$(F$, 89, 8))
        DESCRI$ = TRIM$(Mid$(F$, 5, 64))
        NORFA$ = Mid$(F$, 69, 12)
        CANS$ = " - (" + TRIM$(Str$(CANOF#)) + ")"
        DESCRI1$ = Left$(DESCRI$, 52 - Len(CANS$)) + CANS$
        '
        ECORDEF$ = AJUSTI$(NORFA$, 12) + DESCRI1$
        Call GRAREG("ECORDEF", ECORDEF$, I&)
        '
        If Asc(Mid$(F$, 108, 1)) = 1 Then ' esta lanzada
            X$ = REGLEIDO$("ECORDEF", I&)
            J& = J& + 1
            Call GRAREG("ECORDEP", X$, J&)
        End If
    Next I&
    '
    Call SETULTREG("ECORDEP", J&)
    Call CIERRARCH("ECORDEP")
    Call CIERRARCH("ECORDEF")
    Screen.MousePointer = 1
    '
    If MODOI% = 1 Then Call FINAL("*OPERFAB")
    '
End Sub

Sub ACTECORLAN()   ' GENERA LISTA DE SELECCION DE ORDENES LANZADAS
    '
    Screen.MousePointer = 11
    Call COPYSTRU("ECORDEF", "ECORDEP")
    J& = 0
    '
    RFF$ = RECFILT$("ORDEFABR", 12, Chr$(1))
    'FIN& = ULTREG&("ORDEFABR")
    For KU& = 1 To Len(RFF$) Step 4
        'Call TRACE(20, I&, FIN&)
        I& = CVS(Mid$(RFF$, KU&, 4))
        F$ = REGLEIDO$("ORDEFABR", I&)
        CANOF# = CVD(Mid$(F$, 89, 8))
        DESCRI$ = TRIM$(Mid$(F$, 5, 64))
        NORFA$ = Mid$(F$, 69, 12)
        CANS$ = " - (" + TRIM$(Str$(CANOF#)) + ")"
        DESCRI1$ = Left$(DESCRI$, 52 - Len(CANS$)) + CANS$
        '
        'ECORDEF$ = AJUSTI$(NORFA$, 12) + DESCRI1$
        'Call GRAREG("ECORDEF", ECORDEF$, I&)
        '
        'If Asc(Mid$(F$, 108, 1)) = 1 Then ' esta lanzada
            X$ = REGLEIDO$("ECORDEF", I&)
            J& = J& + 1
            Call GRAREG("ECORDEP", X$, J&)
        'End If
    Next KU&
    '
    Call SETULTREG("ECORDEP", J&)
    Call CIERRARCH("ECORDEP")
    'Call CIERRARCH("ECORDEF")
    Screen.MousePointer = 1
    '
    'If MODOI% = 1 Then Call FINAL("*OPERFAB")
    '
End Sub


Sub TOTEQUIP()
    '
    Dim EQQ$(512), TPREP(512), TOPER(512)
    CAEQUI& = ULTREG&("PADMAQ")
    If CAEQUI& > 512 Then
        Call COMUNI("Demasiados Equipos.\Consulte Soporte de Sistemas.")
        Call FINAL("")
    End If
    '
    Screen.MousePointer = 11
    For I& = 1 To CAEQUI&
        XX$ = REGLEIDO$("PADMAQ", I&)
        EQQ$(I&) = Left$(XX$, 6)
    Next I&
    '
    ULTRE& = ULTREG&("OPERFAP")
    
        For TTL& = 1 To ULTRE&
           '
           SCP$ = REGLEIDO$("OPERFAP", TTL&)
           NORFA$ = TRIM$(Mid$(SCP$, 4, 9))
           PPXX% = InStr(NORFA$, "-"): If PPXX% < 1 Then PPXX% = 1 + Len(NORFA$)
           NOFAX = Val(Left$(NORFA$, PPXX%))
           If NOFAX > 0 Then
              If NOFAX <= 999999 Then
                 EQ1$ = Mid$(SCP$, 73, 6)
                 EQ2$ = Mid$(SCP$, 203, 6)
                 EQ3$ = Mid$(SCP$, 209, 6)
                 EQ4$ = Mid$(SCP$, 215, 6)
                 '
                 MINUPREP = CVS(Mid$(SCP$, 103, 4))
                 MINLOTE = 0: For UU% = 107 To 130 Step 4
                                MINLOTE = MINLOTE + CVS(Mid$(SCP$, UU%, 4))
                              Next UU%
                 CAOPS = 1    '  CVS(MID$(SCP$,83,4))
                 '
                 For J% = 1 To CAEQUI&
                   If TRIM$(EQQ$(J%)) <> "" Then
                     If EQQ$(J%) = EQ1$ Then
                        TPREP(J%) = TPREP(J%) + MINUPREP
                        TOPER(J%) = TOPER(J%) + MINLOTE
                       ElseIf EQQ$(J%) = EQ2$ Then
                        TPREP(J%) = TPREP(J%) + MINUPREP
                        TOPER(J%) = TOPER(J%) + MINLOTE
                       ElseIf EQQ$(J%) = EQ3$ Then
                        TPREP(J%) = TPREP(J%) + MINUPREP
                        TOPER(J%) = TOPER(J%) + MINLOTE
                       ElseIf EQQ$(J%) = EQ4$ Then
                        TPREP(J%) = TPREP(J%) + MINUPREP
                        TOPER(J%) = TOPER(J%) + MINLOTE
                     End If
                   End If
                 Next J%
              End If
           End If
        Next TTL&
        '
        For I& = 1 To CAEQUI&
            XX$ = REGLEIDO$("PADMAQ", I&)
            Call REPLA(XX$, MKS$(60 * TPREP(I&)), 81, 4)
            Call REPLA(XX$, MKS$(60 * TOPER(I&)), 85, 4)
            Call REPLA(XX$, MKS$(60 * (TPREP(I&) + TOPER(I&))), 89, 4)
            Call GRAREG("PADMAQ", XX$, I&)
        Next I&
        '
        Call CIERRARCH("PADMAQ")
        Screen.MousePointer = 1
        '
        Call FRESHALL
        Call FINAL("*TOTIEPRO")
        '
End Sub

Sub VERIDEST()
    '
    STO1% = 0: STO2% = 0
    For I& = 1 To ULTREG&("DESTFAB")
       X$ = REGLEIDO$("DESTFAB", I&)
       CODEST$ = TRIM$(UCase$(Left$(X$, 8)))
       If CODEST$ = "STOCK" Then STO1% = 1
       If CODEST$ = "SEMIELA" Then STO2% = 1
    Next I&
    '
    If STO1% = 0 Then
       X$ = REGBLAN$("DESTFAB")
       Call REPLA(X$, "STOCK", 1, 8)
       Call REPLA(X$, "Stock de Producto Terminado", 9, 32)
       Call REGAPP("DESTFAB", X$)
    End If
    '
    If STO2% = 0 Then
       X$ = REGBLAN$("DESTFAB")
       Call REPLA(X$, "SEMIELA", 1, 8)
       Call REPLA(X$, "Stock de Semielaborados", 9, 32)
       Call REGAPP("DESTFAB", X$)
    End If
    '
    Call CIERRARCH("DESTFAB")
    '
End Sub

Sub REVIORFLAN()
   '
   Screen.MousePointer = 11
   Dim NUORF(16384), REORF&(16384), STORF%(16384)
   J& = 0
   FIN& = ULTREG&("ORFANUE")
   For U& = 1 To FIN&
      Call TRACE(20, U&, FIN&)
      XX$ = REGLEIDO$("ORFANUE", U&)
      If Asc(Mid$(XX$, 112, 1)) = 1 Then
         J& = J& + 1
         NUORF(J&) = CVS(Left$(XX$, 4))
         REORF&(J&) = U&
      End If
   Next U&
   '
   FIN& = ULTREG&("ORDEFABR")
   For U& = 1 To FIN&
      Call TRACE(20, U&, FIN&)
      XX$ = REGLEIDO$("ORDEFABR", U&)
      If Asc(Mid$(XX$, 108, 1)) = 1 Then
         NOFA = CVS(Left$(XX$, 4))
         For K% = 1 To J&
            If NUORF(K%) = NOFA Then
               STORF%(K%) = 1
            End If
         Next K%
      End If
   Next U&
   '
   For U& = 1 To J&
      Call TRACE(20, U&, J&)
      If STORF%(U&) <> 1 Then
         XX$ = REGLEIDO$("ORFANUE", REORF&(U&))
         Call REPLA(XX$, Chr$(3), 112, 1)
         Call GRAREG("ORFANUE", XX$, REORF&(U&))
      End If
   Next U&
   '
   Call CIERRARCH("*.*")
   Screen.MousePointer = 1
   '
End Sub

Sub REVIBOPER()   ' revisión de base de datos de operaciones programadas
   '
   PUNORFA$ = ""
   If TRIM$(CONTROL$("", "REBOPER")) <> "" Then
      Exit Sub
   End If
   '
   TTXX$ = ""
   TTXX$ = TTXX$ + "FLEXOFT Debe Ahora Re-organizar\"
   TTXX$ = TTXX$ + "la Base de Datos de Operaciones Programadas.\  \"
   TTXX$ = TTXX$ + "Para Ejecutar este Proceso de Re-organización\"
   TTXX$ = TTXX$ + "TODOS los Usuarios deben SALIR DE FLEXOFT.\  \"
   TTXX$ = TTXX$ + "Haga Click sobre 'Continuar' una Vez que Haya\"
   TTXX$ = TTXX$ + "Verificado esta Situación.\"
   If COMALTER%(TTXX$ + "\\Continuar\Cancelar") <> 1 Then
      Call CIERRARCH("*.*")
      Call FINAL("")
   End If
   '
   Screen.MousePointer = 11
   Call BLOQARCH("ORDEFABR")
   Call BLOQARCH("ORFANUE")
   Call BLOQARCH("OPERFAB")
   Call BLOQARCH("INFOFAB")
   '
' reconstruye operaciones programadas segun standard
   '
   NUOROP% = 0
   FEAL% = HOY(HOS$)
   HOAL% = Int(HORANUM(Time$))
   USUA% = USNBR% Mod 100
   DALTA$ = Chr$(NUOROP%) + MKI$(FEAL%) + MKI$(HOAL%) + Chr$(USUA%)
   INFOBLA$ = Space$(6) + String$(26, 0) + Space$(16)
   PROGBLA$ = Space$(6) + String$(26, 0)
   '
   JJ& = 0
   FIN& = ULTREG&("ORDEFABR")
   For U& = 1 To FIN&
      Call TRACE(20, U&, FIN&)
      XX$ = REGLEIDO$("ORDEFABR", U&)
      NORFA$ = Mid$(XX$, 69, 12)
      CIJK% = CVI(Mid$(XX$, 83, 2))
      CAORFA# = CVD(Mid$(XX$, 89, 8))
      BVOF% = BINVAL%(AJUSTI$(NORFA$, 12))
      '
      SECUOPSTA$ = RECFILT$("SECOPER", 1, MKI$(CIJK%))
      If Len(SECUOPSTA$) > 0 Then
         For TTI% = 1 To Len(SECUOPSTA$) Step 4
            TTL& = CVS(Mid$(SECUOPSTA$, TTI%, 4))
            SCP$ = REGLEIDO$("SECOPER", TTL&)
            If CVI(Left$(SCP$, 2)) = CIJK% Then
               NO% = CVI(Mid$(SCP$, 3, 2))
               DOPE$ = Mid$(SCP$, 5, 48)
               EQ1$ = Mid$(SCP$, 53, 6)
               EQAUX$ = Mid$(SCP$, 103, 18)
               LOTESTA = LOSTA: If LOTESTA < 1 Then LOTESTA = 1
               CALOTES = 1
               MINPREP = 60 * CVS(Mid$(SCP$, 59, 4)) * CALOTES
               MINFIJO = 60 * CVS(Mid$(SCP$, 63, 4)) * CALOTES
               MINVARI = CAORFA# * CVS(Mid$(SCP$, 67, 4))
               MINDEMO = (MINFIJO + MINVARI) * CVS(Mid$(SCP$, 75, 4)) / 100
               MINLIMP = 60 * CVS(Mid$(SCP$, 79, 4)) * CALOTES
               MINLOTE = MINPREP + MINFIJO + MINVARI + MINDEMO + MINLIMP
               '
               HORAPROD = MINLOTE / 60
               CAOPS = CVS(Mid$(SCP$, 83, 4))
               '
               BVOF% = BINVAL%(AJUSTI$(NORFA$, 12))
               '
               ZZ$ = REGBLAN$("OPERFAB")
               Call REPLA(ZZ$, NORFA$, 1, 12)
               Call REPLA(ZZ$, MKI$(CIJK%), 13, 2)
               Call REPLA(ZZ$, MKI$(NO%), 15, 2)
               Call REPLA(ZZ$, DOPE$, 17, 48)
               Call REPLA(ZZ$, MKI$(BVOF%), 65, 2)
               Call REPLA(ZZ$, DALTA$, 67, 6)
               Call REPLA(ZZ$, EQ1$, 73, 6)
               Call REPLA(ZZ$, MKD$(CAORFA#), 79, 8)
               Call REPLA(ZZ$, MKD$(CAORFA#), 95, 8)
               Call REPLA(ZZ$, MKS$(MINPREP), 103, 4)
               Call REPLA(ZZ$, MKS$(MINFIJO), 107, 4)
               Call REPLA(ZZ$, MKS$(MINVARI), 111, 4)
               Call REPLA(ZZ$, MKS$(MINDEMO), 115, 4)
               Call REPLA(ZZ$, MKS$(MINLIMP), 123, 4)
               Call REPLA(ZZ$, MKS$(MINLOTE), 131, 4)
               Call REPLA(ZZ$, String$(16, 0), 135, 16)
               Call REPLA(ZZ$, INFOBLA$, 151, 48)
               Call REPLA(ZZ$, MKS$(TTL&), 199, 4)
               Call REPLA(ZZ$, EQAUX$, 203, 18)
               Call REPLA(ZZ$, MKS$(CAOPS), 221, 4)
               Call REPLA(ZZ$, PROGBLA$, 225, 32)
               '
               JJ& = JJ& + 1
               Call GRAREG("OPERFAB", ZZ$, JJ&)
               PUNORFA$ = PUNORFA$ + AJUSTI$(NORFA$, 12)
            End If
         Next TTI%
      End If
   Next U&
   Call SETULTREG("OPERFAB", JJ&)
   Call CIERRARCH("*.*")
   '
' blanquear operaciones por informes de produccion s/INFOFAB
   '
   FEBA% = HOY(HOS$)
   HOBA% = Int(HORANUM(Time$))
   USUA% = USNBR% Mod 100
   DBAJA$ = MKI$(FEBA%) + MKI$(HOBA%) + Chr$(USUA%) + Chr$(1) + "Parte Fabric."
   '
   FLIMI% = 0
   EPAC$ = UCase$(EMPREAC$)
   If InStr(EPAC$, "MAINIERI") > 0 Or InStr(EPAC$, "FASS") > 0 Then
      FLIMI% = FECHANUM("01/12/00")
   End If
   '
   FIN& = ULTREG&("INFOFAB")
   For U& = 1 To FIN&
      Call TRACE(20, U&, FIN&)
      XX$ = REGLEIDO$("INFOFAB", U&)
      NORFB$ = Left$(XX$, 12)
      CIIK% = CVI(Mid$(XX$, 13, 2))
      NOPE% = CVI(Mid$(XX$, 15, 2))
      CAREAL# = CVD(Mid$(XX$, 65, 8))
      MAQUI$ = Mid$(XX$, 73, 6)
      FERE% = CVI(Mid$(XX$, 79, 2))
      If FERE% < FLIMI% Then GoTo 29
      HORE% = CVI(Mid$(XX$, 81, 2))
      TEJEC = CVS(Mid$(XX$, 84, 4))
      OBSERVA$ = Mid$(XX$, 88, 30)
      LEGOPER$ = Mid$(XX$, 118, 8)
      For KIK% = 1 To 8 Step 2
        If CVI(Mid$(LEGOPER$, KIK%, 2)) = 8224 Then
          Mid$(LEGOPER$, KIK%, 2) = MKI$(0)
        End If
      Next KIK%
      '
      If CVI(Left$(LEGOPER$, 2)) < 1 Then
        If TRIM$(OBSERVA$) <> "" Then
           NOMBROPER$ = TRIM$(OBSERVA$)
           LEGA% = 0
           For TTB& = 1 To ULTREG&("PADRON")
             QQ$ = REGLEIDO$("PADRON", TTB&)
             If InStr(Mid$(QQ$, 3, 30), NOMBROPER$) > 0 Then
               LEGA% = CVI(Left$(QQ$, 2))
               GoTo 19
             End If
           Next TTB&
19         LEGOPER$ = MKI$(LEGA%) + String$(6, 0)
        End If
      End If
      '
      RFF$ = "": ABUS$ = AJUSTI$(NORFB$, 12)
      For KJ& = 1 To Len(PUNORFA$) Step 12
        If Mid$(PUNORFA$, KJ&, 12) = ABUS$ Then
          RFF$ = RFF$ + MKS$((KJ& + 11&) / 12)
        End If
      Next KJ&
      '
      For KK% = 1 To Len(RFF$) Step 4
        REOPER& = CVS(Mid$(RFF$, KK%, 4))
        ZZ$ = REGLEIDO$("OPERFAB", REOPER&)
        NORFA$ = Left$(ZZ$, 12)
        If NORFA$ = NORFB$ Then
          If CVI(Mid$(ZZ$, 15, 2)) = NOPE% Then
            If Mid$(ZZ$, 177, 6) = String$(6, 0) Then ' marca de pendiente
              '
              SALRE# = CVD(Mid$(ZZ$, 79, 8)) - CAREAL#
              If SALRE# < 0 Then SALRE# = 0
              NUOROP% = Asc(Mid$(ZZ$, 67, 1))
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
              '
              Call GRAREG("OPERFAB", ZZ$, REOPER&)
              '
              If SALRE# > 0.01 Then
                 '
                 NUOROP% = 1 + Asc(Mid$(ZZ$, 67, 1))
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
                 '
                 Call REGAPP("OPERFAB", ZZ$)
                 NORX$ = Left$(ZZ$, 12)
                 PUNORFA$ = PUNORFA$ + NORX$
                 '
              End If
              GoTo 29
            End If
          End If
        End If
      Next KK%
      '
      ' AGREGA REGISTRO A OPERACIONES PROGRAMADAS
      '
      NORFB$ = Left$(XX$, 12)
      CIIK% = CVI(Mid$(XX$, 13, 2))
      NOPE% = CVI(Mid$(XX$, 15, 2))
      DOPE$ = DESCRIOP$(CIIK%, NOPE%)
      CAREAL# = CVD(Mid$(XX$, 65, 8))
      MAQUI$ = Mid$(XX$, 73, 6)
      FERE% = CVI(Mid$(XX$, 79, 2))
      HORE% = CVI(Mid$(XX$, 81, 2))
      TEJEC = CVS(Mid$(XX$, 84, 4))
      OBSERVA$ = Mid$(XX$, 88, 30)
      LEGOPER$ = Mid$(XX$, 118, 8)
      For KIK% = 1 To 8 Step 2
        If CVI(Mid$(LEGOPER$, KIK%, 2)) = 8224 Then
          Mid$(LEGOPER$, KIK%, 2) = MKI$(0)
        End If
      Next KIK%
      '
      NUOROP% = 0
      FEAL% = HOY(HOS$)
      HOAL% = Int(HORANUM(Time$))
      USUA% = USNBR% Mod 100
      DALTA$ = Chr$(NUOROP%) + MKI$(FEAL%) + MKI$(HOAL%) + Chr$(USUA%)
      INFOBLA$ = Space$(6) + String$(26, 0) + Space$(16)
      PROGBLA$ = Space$(6) + String$(26, 0)
      BVOF% = BINVAL%(AJUSTI$(NORFB$, 12))
      EQAUX$ = ""
      MINVARI = TIESTAN(CIIK%, NOPE%, CAREAL#, NUOROP%)
      '
      ZZ$ = REGBLAN$("OPERFAB")
      Call REPLA(ZZ$, NORFB$, 1, 12)
      Call REPLA(ZZ$, MKI$(CIIK%), 13, 2)
      Call REPLA(ZZ$, MKI$(NOPE%), 15, 2)
      Call REPLA(ZZ$, DOPE$, 17, 48)
      Call REPLA(ZZ$, MKI$(BVOF%), 65, 2)
      Call REPLA(ZZ$, DALTA$, 67, 6)
      Call REPLA(ZZ$, MAQUI$, 73, 6)
      Call REPLA(ZZ$, MKD$(CAREAL#), 79, 8)
      Call REPLA(ZZ$, MKD$(CAREAL#), 95, 8)
      Call REPLA(ZZ$, String$(32, 0), 103, 32)
      Call REPLA(ZZ$, MKS$(MINVARI), 111, 4)
      Call REPLA(ZZ$, MKS$(MINVARI), 131, 4)
      Call REPLA(ZZ$, String$(16, 0), 135, 16)
      Call REPLA(ZZ$, INFOBLA$, 151, 48)
      Call REPLA(ZZ$, MKS$(TTL&), 199, 4)
      Call REPLA(ZZ$, EQAUX$, 203, 18)
      Call REPLA(ZZ$, MKS$(1), 221, 4)
      Call REPLA(ZZ$, PROGBLA$, 225, 32)
      '
      TSTAN = TIESTAN(CIIK%, NOPE%, CAREAL#, NUOROP%)
      '
      Call REPLA(ZZ$, MKD$(CAREAL#), 87, 8)
      Call REPLA(ZZ$, MKD$(0), 95, 8)
      Call REPLA(ZZ$, MAQUI$, 151, 6)
      Call REPLA(ZZ$, MKI$(FERE%), 157, 2)
      Call REPLA(ZZ$, MKI$(HORE%), 159, 2)
      Call REPLA(ZZ$, MKS$(TEJEC), 161, 4)
      Call REPLA(ZZ$, LEGOPER$, 165, 8)
      Call REPLA(ZZ$, MKS$(TSTAN), 173, 4)
      Call REPLA(ZZ$, DBAJA$, 177, 22)
      '
      Call REGAPP("OPERFAB", ZZ$)
      PUNORFA$ = PUNORFA$ + AJUSTI$(NORFB$, 12)
      '
29 Next U&
   '
   Call CIERRARCH("*.*")
   '
'anular operaciones segun ordenes cerradas o anuladas
   '
   PUNORFA$ = ""
   FIN& = ULTREG&("ORDEFABR")
   For U& = 1 To FIN&
      Call TRACE(20, U&, FIN&)
      XX$ = REGLEIDO$("ORDEFABR", U&)
      NORFA$ = Mid$(XX$, 69, 12)
      PUNORFA$ = PUNORFA$ + AJUSTI$(NORFA$, 12)
   Next U&
   '
   FIN& = ULTREG&("OPERFAB")
   For U& = 1 To FIN&
      ZZ$ = REGLEIDO$("OPERFAB", U&)
      If Mid$(ZZ$, 177, 6) = String$(6, 0) Then ' marca de pendiente
        NORFA$ = AJUSTI$(Left$(ZZ$, 12), 12)
        For KJ& = 1 To Len(PUNORFA$) Step 12
          If Mid$(PUNORFA$, KJ&, 12) = NORFA$ Then
            REORFA& = (KJ& + 11) / 12
            GoTo 39
          End If
        Next KJ&
        REORFA& = 0
        STATU% = 9
        '
39      If REORFA& > 0 Then
          XX$ = REGLEIDO$("ORDEFABR", REORFA&)
          STATU% = Asc(Mid$(XX$, 108, 1))
        End If
        '
        If STATU% = 3 Then
            FEBA% = HOY(HOS$)
            HOBA% = Int(HORANUM(Time$))
            USUA% = USNBR% Mod 100
            DBAJA$ = MKI$(FEBA%) + MKI$(HOBA%) + Chr$(USUA%) + Chr$(1) + "Cierre O/F"
            Call REPLA(ZZ$, DBAJA$, 177, 22)
            Call GRAREG("OPERFAB", ZZ$, U&)
          ElseIf STATU% > 3 Then
            FEBA% = HOY(HOS$)
            HOBA% = Int(HORANUM(Time$))
            USUA% = USNBR% Mod 100
            DBAJA$ = MKI$(FEBA%) + MKI$(HOBA%) + Chr$(USUA%) + Chr$(1) + "Anulacion O/F"
            Call REPLA(ZZ$, DBAJA$, 177, 22)
            Call GRAREG("OPERFAB", ZZ$, U&)
        End If
        '
      End If
   Next U&
   '
   Call CIERRARCH("OPERFAB")
   Call CIERRARCH("*.*")
   Reset
   '
   HOII% = HOY(HOS$)
   Call GRACONTROL("", "REBOPER", HOS$)
   Screen.MousePointer = 1
   '
End Sub

Sub GENPEDPEN()
   '
   JJ& = 0
   UREPRE& = ULTREG&("PEDDETA")
   Screen.MousePointer = 11
   JJ& = 0
   NROPEDA& = 0: CAPED% = 0
   FINF = 999999: FSUP = 0
   RFF$ = RECFILT$("PEDDETA", 7, Chr$(1)) ' SOLO PENDIENTES DE ENTREGA Y NO DESPACHADOS
   For UU& = 1 To Len(RFF$) Step 4
     REPRESU& = CVS(Mid$(RFF$, UU&, 4))
     X$ = REGLEIDO$("PEDDETA", REPRESU&)
     If Asc(Mid$(X$, 27, 1)) <> 9 Then ' no esta anulado
       CAN = CVS(Mid$(X$, 79, 4))
       CADESPACH = CVS(Mid$(X$, 95, 4))
       CAPENDENT = CAN - CADESPACH
       '
       If Asc(Mid$(X$, 27, 1)) < 3 Then
         If CAPENDENT < 0.05 Then
           Call REPLA(X$, Chr$(3), 27, 1)
           Call GRAREG("PEDDETA", X$, REPRESU&)
         End If
       End If
       '
       If CAPENDENT >= 0.05 Then
         NRODES& = CVS(Left$(X$, 4))
         If NRODES& <> NRODESA& Then
           PRIREDETA& = REPRESU&
           NRODESA& = NRODES&
           TTXX$ = MKS$(NRODES&)
           NOCO$ = TRIM$(Mid$(X$, 9, 16))
           If NOCO$ <> "" Then
             TTXX$ = TTXX$ + "O.C." + NOCO$ + " - "
           End If
           FEX = CVI(Mid$(X$, 5, 2))
           TTXX$ = TTXX$ + FECHATEX$(FEX) + " - "
           NCLIE% = CVI(Mid$(X$, 7, 2))
           TTXX$ = TTXX$ + TRIM$(RASOCLI$(NCLIE%)) + Space$(52)
           REGENCA& = CVS(Mid$(X$, 109, 4))
           JJ& = JJ& + 1
           Call GRAREG("INDIPEP", Left$(TTXX$, 52) + MKS$(REGENCA&), JJ&)
         End If
       End If
     End If
     '
   Next UU&
   Call SETULTREG("INDIPEP", JJ&)
   Call CIERRARCH("INDIPEP")
   Call FRESHECHO("INDIPEP")
   '
   Screen.MousePointer = 1
   '
End Sub
'
Sub PROCERECHA(CODI1%, FERECH%, CARECHA, DOCORX$) '\\\OT-06-0191-WAR-09/06/06///
  '
  If CARECHA <= 0 Then
    Call MENSERR(24, "Cantidad de Rechazo no Válida.")
    Exit Sub
  End If
  '
    'CARGO EL DEFAULT DEL PASO 2
    VDEF$ = Space$(256)
    FIN1& = ULTREG&("RECHLIN")
    ULTRECH& = FIN1& + 1
    Call REPLA(VDEF$, MKS$(ULTRECH&), 199, 4)
    PROHACOD% = PROVHABI%(CODI1%)
    Call REPLA(VDEF$, MKI$(CODI1%), 1, 2)
    Call REPLA(VDEF$, MKI$(FERECH%), 3, 2)
    Call REPLA(VDEF$, MKI$(PROHACOD%), 5, 2)
    Call REPLA(VDEF$, Chr$(0), 12, 1)
    Call REPLA(VDEF$, Chr$(0), 11, 1)
    Call REPLA(VDEF$, Chr$(0), 16, 1)
    Call REPLA(VDEF$, DOCORX$, 17, 12) '\\\OT-06-0191-WAR-09/06/06///
    '
15  OBJ$ = OBJPLA$("DETARECHA", VDEF$)
    PA2$ = OBJ$
    If TRIM$(PA2$) <> "" Then
      'VALIDACIONES
      CODI2% = CVI(Mid$(PA2$, 1, 2))
      FEMI1$ = FECHATEX$(CVI(Mid$(PA2$, 3, 2)))
      NPRO% = CVI(Mid$(PA2$, 5, 2))
      DESTI% = Asc(Mid$(PA2$, 11, 1))
      SECT& = Asc(Mid$(PA2$, 12, 1)) 'AHORA DEPOSITO
      MOTI% = Asc(Mid$(PA2$, 16, 1))
      CAN3 = CARECHA
      REFE1$ = TRIM$(Mid$(PA2$, 17, 48))
      '
      JJ& = 0
      For U& = 65 To 193 Step 43
        TTXX$ = Mid$(PA2$, U&, 43)
        JJ& = JJ& + 1
        Call GRAREG("!OBSERVOF", TTXX$, JJ&)
      Next U&
      Call SETULTREG("!OBSERVOF", JJ&)
      '
      If TRIM$(ECOARCH$("PROVED1", MKI$(NPRO%))) = "" Then
        If CARPROVE% > 0 Then
          Call MENSERR(24, "Proveedor no Válido.")
          VDEF$ = PA2$
          GoTo 15
        End If
      End If
      If TRIM$(ECOARCH$("TADESTI", Chr(DESTI%))) = "" Then
        Call MENSERR(24, "Destino de Rechazo no Válido.")
        VDEF$ = PA2$
        GoTo 15
      End If
      If TRIM$(ECOARCH$("TADEPOSI", Chr$(SECT&))) = "" Then
        Call MENSERR(24, "Depósito no Válido.")
        VDEF$ = PA2$
        GoTo 15
      End If
      If TRIM$(ECOARCH$("TAMOTIV", Chr(MOTI%))) = "" Then
        Call MENSERR(24, "Motivo de Rechazo no Válido.")
        VDEF$ = PA2$
        GoTo 15
      End If
      If TRIM$(REFE1$) = "" Then
        Call MENSERR(24, "Referencia de Rechazo no Válida.")
        VDEF$ = PA2$
        GoTo 15
      End If
      '
      ZZZ$ = FILTERGETRECORD$("TADESTI", 1, Chr$(DESTI%))
      CARPROVE% = 0
      If TRIM$(UCase$(Mid$(ZZZ$, 53, 2))) = "SI" Then
        CARPROVE% = 1
      End If
      '
      'CALCULO DE PRECIO E IMPORTE
      PREUNI1# = ULTIPRE#(NPRO%, CODI2%)
      IMPORT1# = PREUNI1# * CAN3
      If CARPROVE% < 1 Then
        NPRO% = 0: PREUNI1# = 0: IMPORT1# = 0
      End If
      'REGISTRO DE RECHAZO EN BDD
      FIN1& = ULTREG&("RECHLIN")
      ULTRECH& = FIN1& + 1
      Y$ = Space(128)
      Call REPLA(Y$, MKS$(ULTRECH&), 1, 4)
      Call REPLA(Y$, MKI$(CODI2%), 5, 2)
      Call REPLA(Y$, MKI$(FECHANUM(FEMI1$)), 7, 2)
      Call REPLA(Y$, MKI$(NPRO%), 9, 2)
      Call REPLA(Y$, MKS$(CAN3), 11, 4)
      Call REPLA(Y$, Chr$(DESTI%), 15, 1)
      Call REPLA(Y$, Chr$(MOTI%), 16, 1)
      Call REPLA(Y$, USERNAME$, 17, 32)
      Call REPLA(Y$, REFE1$, 49, 48) 'FALTA REFERENCIA
      Call REPLA(Y$, MKD$(PREUNI1#), 97, 8) 'FALTA VALOR UNITARIO
      Call REPLA(Y$, MKD$(IMPORT1#), 105, 8) 'FALTA IMPORTE
      Call REPLA(Y$, MKI$(0), 113, 2) 'FALTA PROCESADO FECHA
      Call REPLA(Y$, MKI$(0), 115, 2) 'FALTA NUMERO USUARIO PROCESO
      Call REGAPP("RECHLIN", Y$)
      Call CIERRARCH("RECHLIN")
      'AGREGO A ARCHIVO INDICE DE RECHAZO
      REFERE2$ = FEMI1$
         If CARPROVE% > 0 Then
              REFERE2$ = REFERE2$ + " - " + RASOCLI$(NPRO% * (-1))
            Else
              REFERE2$ = REFERE2$ + " - Cargo Interno"
         End If
      ZZ$ = Space(64)
      Call REPLA(ZZ$, MKS$(ULTRECH&), 1, 4)
      Call REPLA(ZZ$, REFERE2$, 5, 60)
      Call REGAPP("INDIRECH", ZZ$)
      Call CIERRARCH("INDIRECH")
      'ACA HAGO LOS MOVIMIENTOS DE STOCK SEGUN DEPOSITO
      'XXX$ = FILTERGETRECORD$("SECPROD", 1, MKS$(SECT&))
      DEPORIG% = Val(SECT&) 'Asc(Mid$(XXX$, 81, 1)) 'DEP. ORIGEN
      ZZZ$ = FILTERGETRECORD$("TADESTI", 1, Chr$(DESTI%))
      DEPODES% = Asc(Mid$(ZZZ$, 32, 1)) 'DEP. DESTINO
      '
      Screen.MousePointer = 11
      NRO& = 1
      For U& = ULTREG&("MOVISTO") To 1 Step -1
        XX$ = REGLEIDO$("MOVISTO", U&)
        If UCase$(Mid$(XX$, 21, 2)) = "TR" Then
          NRO& = 1 + Val(Mid$(XX$, 26, 7))
          GoTo 110
        End If
      Next U&
      '
110   Screen.MousePointer = 1
      REMI$ = "TR." + TRIM$(Str$(NRO&))
      '
      FEMOVI% = FECHANUM(FEMI1$)
      CI% = CODI2%
      LOTE$ = ""
      DOSEC$ = "RECH-" + TRIM$(Str$(NRO&))
      REFMO$ = REFE1$
      CANTI = (-1) * CAN3
      '
      Call MOVISTO(FEMOVI%, CI%, LOTE$, "TR", REMI$, DOSEC$, REFMO$, 0, "$", 0, DEPORIG%, CANTI)
      Call MOVISTO(FEMOVI%, CI%, LOTE$, "TR", REMI$, DOSEC$, REFMO$, 0, "$", 0, DEPODES%, (-1) * CANTI)
      Call CIERRARCH("MOVISTO")
      '
      'SI DESTINO = DEVOLUCION EMITO UN REMITO DE DEVOLUCION
'      If DESTI% = 1 Then
'        RECHA# = Val(CAN3)
'        NURECH$ = "RL-" & TRIM$(Str$(ULTRECH&)) 'RL = RECHAZO EN LINEA
'        Call REMIDEV(NPRO%, DEPORIG%, CODI2%, RECHA#, NURECH$, HOII%, PREUNI1#)
'      End If

      'PREPARACION PARA IMPRESIÓN
      'Label2 = FEMI1$
      CI% = CODI2%
      'Label4 = FORMATNUM$(CAN3, "F11.1")
'      Label5 = FORMATNUM$(DESTI%, "F9.0")
'      Label7 = FORMATNUM$(NPRO%, "F9.0")
'      Label8 = FORMATNUM$(MOTI%, "F9.0")
'      Label9 = REFE1$
      NUDO$ = FORMATNUM$(ULTRECH&, "F11")
'      Label15 = FORMATNUM$(PREUNI1#, "F11.4")
'      Label16 = FORMATNUM$(IMPORT1#, "F11.2")
      '
      Call PRIRECHALI(1, FEMI1$, NUDO$, CI%, CAN3, DESTI%, NPRO%, MOTI%, REFMO$, PREUNI1#, IMPORT1#) 'IMPRESION LOCAL
      '
      If CARPROVE% < 1 Then GoTo 25 ' NO IMPRIME EN TESORERIA Y MAIL PROVEEDOR
      '
      JJ& = 0
      For U& = 65 To 193 Step 43
        TTXX$ = Mid$(PA2$, U&, 43)
        JJ& = JJ& + 1
        Call GRAREG("!OBSERVOF", TTXX$, JJ&)
      Next U&
      Call SETULTREG("!OBSERVOF", JJ&)
      '
      Call PRIRECHALI(2, FEMI1$, NUDO$, CI%, CAN3, DESTI%, NPRO%, MOTI%, REFMO$, PREUNI1#, IMPORT1#) 'IMPRESION REMOTA
      '
      JJ& = 0
      For U& = 65 To 193 Step 43
        TTXX$ = Mid$(PA2$, U&, 43)
        JJ& = JJ& + 1
        Call GRAREG("!OBSERVOF", TTXX$, JJ&)
      Next U&
      Call SETULTREG("!OBSERVOF", JJ&)
      '
      EMAI$ = TRIM$(EMAILCLI$((-1) * NPRO%))
      If EMAI$ = "" Or InStr(EMAI$, "@") < 1 Then
          Call MENSERR(24, "Imposible Enviar por Correo Electrónico.\  \Dirección de Correo del Proveedor es\Inexistente o no Válida.")
        Else
          SENDBYMAIL$ = EMAI$
          Call PRIRECHALI(1, FEMI1$, NUDO$, CI%, CAN3, DESTI%, NPRO%, MOTI%, REFMO$, PREUNI1#, IMPORT1#)
      End If
      '
25  End If
    '
End Sub
Function ULTIPRE#(NPRO%, CIIX%)
  '
  'DEVUELVE EL ULTIMO PRECIO DE COMPRA DEL ARTICULO CIIX%
  'AL PROVEEDOR NPRO%
  '
  UC$ = ULTICOM$(NPRO%, CIIX%)
  MONULCO% = Asc(Mid$(UC$, 122, 1)): If MONULCO% < 1 Then MONULCO% = 1
  ULTIPRE# = CVD(Mid$(UC$, 67, 8)) * TICAMONE(MONULCO%)
  '
End Function

Function ULTICOM$(NPRO%, CIIX%)
   '
   ' trae el registro de OCOMDETA correspondiente a la
   ' última compra no anulada del artículo y proveedor
   ' con precio registrado mayor que cero
   '
   FULCO% = 0: UCO$ = REGBLAN$("OCOMDETA")
   '
   RFF$ = RECFILT$("OCOMDETA", 5, MKI$(CIIX%))
   If Len(RFF$) >= 4 Then
     For U& = 1 To Len(RFF$) Step 4
       REGO& = CVS(Mid$(RFF$, U&, 4))
       YX$ = REGLEIDO$("OCOMDETA", REGO&)
       If CVI(Mid$(YX$, 7, 2)) = NPRO% Then
         If CVI(Mid$(YX$, 9, 2)) = CIIX% Then
           PREUNRE# = CVD(Mid$(YX$, 67, 8))
           If PREUNRE# > 0.00005 Or FULCO% = 0 Then
             If CVI(Mid$(YX$, 5, 2)) >= FULCO% Then
               UCO$ = YX$
               FULCO% = CVI(Mid$(YX$, 5, 2))
             End If
           End If
         End If
       End If
     Next U&
   End If
   '
   ULTICOM$ = UCO$
   '
End Function

Sub PRIRECHALI(DONDE%, FETE$, NUDO$, CI%, CAN3, DESTI%, NPRO%, MOTI%, REFE1$, PREUNI1#, IMPORT1#)
    ' IMPRESION DEL RECHAZO EN LÍNEA
    'DONDE%=1 IMPRESORA LOCAL (IMPRESORA DEFAULT)
    'DONDE%=2 IMPRESORA REMOTA (IMPRESORA CUENTAS A PAGAR)
    '
    FORIPRE$ = CONTROL$("", "RECHALIN")
    '
    If TRIM$(FORIPRE$) <> "" Then
      '
      Printer.TrackDefault = False
      If DONDE% = 2 Then 'IMPRESION REMOTA
        PORTUFAC$ = TRIM$(CONTROL$("", "RUIMPREC"))
        PRINTERPORT$ = PORTUFAC$
      End If
      '
      Screen.MousePointer = 11
      'FECHEMII% = FECHANUM(Label2)
      FECHEMII% = FECHANUM(FETE$)
      FECHDOC$ = FECHATEX$(FECHEMII%)
      'NUMEDOC$ = TRIM$(Label10)
      NUMEDOC$ = TRIM$(NUDO$)
      While Len(NUMEDOC$) < 6
        NUMEDOC$ = "0" + NUMEDOC$
      Wend
      TIPODOC$ = "Rechazo en Línea"
      '
      CODPARAM$(1) = "COD-MAT"
      CODPARAM$(2) = "DES-MAT"
      CODPARAM$(3) = "CANTIDAD"
      CODPARAM$(4) = "DESTINO"
      CODPARAM$(5) = "RASO-CLI"
      CODPARAM$(6) = "REF-MAT1"
      CODPARAM$(7) = "PRE-NETO"
      CODPARAM$(8) = "NET-ITEM"
      CODPARAM$(9) = "REF-MAT2"
      CODPARAM$(10) = "USUARIO"
      CODPARAM$(11) = "REF-MAT3"
      CODPARAM$(12) = "DOCU-APL"
      CAPARDOC% = 12
      '
'      CI% = Val(Label3)
'      CAN3 = Val(Label4)
'      DESTI% = Val(Label5)
'      NPRO% = Val(Label7)
'      MOTI% = Val(Label8)
'      PREUNI1# = Val(Label10)
'      REFE1$ = Label9
'      PREUNI1# = Val(Label15)
'      IMPORT1# = Val(Label16)
      '
      DOCPARAM$(1) = CODEXT$(CI%)
      DOCPARAM$(2) = DESCRIT0$(CI%)
      DOCPARAM$(3) = FORMATNUM$(CAN3, "F9.2")
      DOCPARAM$(4) = TRIM$(ECOARCH$("TADESTI", Chr(DESTI%)))
      DOCPARAM$(5) = ""
        If NPRO% > 0 Then DOCPARAM$(5) = RASOCLI$(NPRO% * (-1))
      DOCPARAM$(6) = TRIM$(ECOARCH$("TAMOTIV", Chr(MOTI%)))
      DOCPARAM$(7) = FORMATNUM$(PREUNI1#, "F11.4") 'PRUEBA PARA VALOR UNITARIO
      DOCPARAM$(8) = FORMATNUM$(IMPORT1#, "F11.2") 'PRUEBA PARA IMPORTE
      DOCPARAM$(9) = REFE1$ 'PRUEBA DE REFERENCIA
      DOCPARAM$(10) = ""
      DOCPARAM$(11) = "" 'NADA EN E TESORERIA
      DOCPARAM$(12) = "" 'NADA EN  N/D
      '
      DESTIDOC% = 0
      CAITAB1% = 0
      '
      Call PRINTDOC("RECHALIN")   ' RECHAZO EN LÍNEA
      If DONDE% = 2 Then 'IMPRESION REMOTA
        PRINTERPORT$ = ""
        Call SETPRINPORT("") 'RESETEA EL CAMBIO DE PUERTO DE IMPRESIÓN.
      End If
      '
    End If
  '
End Sub
