Attribute VB_Name = "CANECO07"
Dim LOST$(10)
Dim LOSTN(10)
Dim CII1%(1024), CANT1(1024), PAD1%(1024)
Dim INIPER%(53), CAPER(53), DEPER$(53)
'
Dim LOREC As Integer
Dim CACOL As Integer
Dim COSEMACT As Integer
Dim HOO As Integer
Dim HO As String
Dim DMX%(12)
'
Dim DISPACT(16384)
Dim LOREPOS(16384)
Dim LINDIVI(16384)
Dim DPI%(16384)
Dim PROVARTI%(16384)
'
Dim SST As String * 40
Dim M0 As String * 128
Dim RA As String * 128
Dim RB As String * 256
Dim NESTRU As Long
Dim PUNTCI As String
Dim PUNTRG As String
'
Dim SEMANAL As Integer
Dim PERIO As String
Dim PLADISTA As Integer
Dim REPOMIN As Single
'
Dim CTX As Integer
Dim CACO As Long
Dim PUNTE$(32767), PUNRE$(32767), PUNFE$(32767)
'
Dim ACTICOD%(32767)
Dim FILPEDIPEN%
'
Dim PROVINT%(32767)
Dim NIVEGEN%(32767)
Public OPCIPLAN%         ' 0=solo entregas / 1=entregas o fabricacion
Public MODOPLAN%         ' 0=programa entregas / 1=plan maestro de fabricacion
'
Sub PARACAL()

'   ************* determinacion parametros de calculo ******************
    BALAREI% = 0
    'If CantRegistros("BALAREQMO") > 0 Then BALAREI% = 1
    '
    CT$ = UCase$(Control$("CANECOM", "PERIODO"))
    SEMANAL% = 0: FORCANEC07.Option5.Value = True
    If InStr(CT$, "SEMAN") > 0 Then
        SEMANAL% = 1
        FORCANEC07.Option4.Value = True
    End If
    '
    CT$ = UCase$(Control$("CANECOM", "PLADISTA"))
    PLADISTA% = Val(CT$)
    If PLADISTA% < 0 Then PLADISTA% = 0
    If PLADISTA% > 6 Then PLADISTA% = 6
    FORCANEC07.Text1.TEXT = TRIM$(Str$(PLADISTA%))
    '
    CT$ = UCase$(Control$("CANECOM", "LOST-TER"))
    LOSTN(3) = Val(CT$): If LOSTN(3) < 0 Then LOSTN(3) = 0
    LOST$(3) = Mid$(Str$(LOSTN(3)), 2)
    FORCANEC07.Text2.TEXT = LOST$(3)
    '
    CT$ = UCase$(Control$("CANECOM", "LOST-MON"))
    LOSTN(2) = Val(CT$): If LOSTN(2) < 0 Then LOSTN(2) = 0
    LOST$(2) = Mid$(Str$(LOSTN(2)), 2)
    FORCANEC07.Text3.TEXT = LOST$(2)
    '
    CT$ = UCase$(Control$("CANECOM", "LOST-MEC"))
    LOSTN(1) = Val(CT$): If LOSTN(1) < 0 Then LOSTN(1) = 0
    LOST$(1) = Mid$(Str$(LOSTN(1)), 2)
    FORCANEC07.Text4.TEXT = LOST$(1)
    '
    CT$ = UCase$(Control$("CANECOM", "LOST-COM"))
    LOSTN(0) = Val(CT$): If LOSTN(0) < 0 Then LOSTN(0) = 0
    LOST$(0) = Mid$(Str$(LOSTN(0)), 2)
    FORCANEC07.Text5.TEXT = LOST$(0)
    '
    CT$ = UCase$(Control$("CANECOM", "REPOMIN"))
    REPOMIN = Val(CT$): If REPOMIN < 1 Then REPOMIN = 1
    REPOMINS$ = Mid$(Str$(REPOMIN), 2)
    '
    FORCANEC07.Check1.Value = 0
    FORCANEC07.Check2.Value = 0
    FORCANEC07.Check3.Value = 0
    FORCANEC07.Check4.Value = 0
    FORCANEC07.Check5.Value = 0
    FORCANEC07.Check6.Value = 0
    FORCANEC07.Check7.Value = 0
    FORCANEC07.Label1 = ""
    '
    If BALAREI% > 0 Then
       FORCANEC07.Option1.Value = True
       FORCANEC07.Check8.Value = 1
       FORCANEC07.Check9.Value = 1
       FORCANEC07.Check10.Value = 1
    End If
       
    CTX% = 0
    '
End Sub
'

Sub GRAPARCAL()
    '
    SEMANAL% = 0: PERIO$ = "MENSUAL"
    If FORCANEC07.Option4.Value = True Then
        SEMANAL% = 1: PERIO$ = "SEMANAL"
    End If
    Call GRACONTROL("CANECOM", "PERIODO", PERIO$)
    '
    PLADISTA% = Val(FORCANEC07.Text1.TEXT)
    Call GRACONTROL("CANECOM", "PLADISTA", TRIM$(Str$(PLADISTA%)))
    '
    LOST$(3) = TRIM$(FORCANEC07.Text2.TEXT): LOSTN(3) = Val(LOST$(3))
    LOST$(2) = TRIM$(FORCANEC07.Text3.TEXT): LOSTN(2) = Val(LOST$(2))
    LOST$(1) = TRIM$(FORCANEC07.Text4.TEXT): LOSTN(1) = Val(LOST$(1))
    LOST$(0) = TRIM$(FORCANEC07.Text5.TEXT): LOSTN(0) = Val(LOST$(0))
    REPOMINS$ = "1": REPOMIN = 1
    '
    Call GRACONTROL("CANECOM", "LOST-TER", LOST$(3))
    Call GRACONTROL("CANECOM", "LOST-MON", LOST$(2))
    Call GRACONTROL("CANECOM", "LOST-MEC", LOST$(1))
    Call GRACONTROL("CANECOM", "LOST-COM", LOST$(0))
    Call GRACONTROL("CANECOM", "REPOMIN", REPOMINS$)
    '
End Sub
'
Sub SETCALCU()
    '
    Screen.MousePointer = 11
    '
    HOO% = HOY(HO$)
    '
    LOREC% = 72  ' *************  caso programacion mensual
    CACOL% = 18  '                cantidad de periodos
    COSEMACT% = HOO%            ' comienzo semana actual
    '
    J% = 1
    IPERIO$ = "01" + Mid$(HO$, 3, 6)
    INIPER%(1) = FECHANUM(IPERIO$)
    DIBAS = INIPER%(1)
        '
        While J% < CACOL%
95         DIBAS = DIBAS + 1
           FF = DIBAS: DIA = Val(FECHATEX$(FF))
                       If DIA <> 1 Then GoTo 95
           FF = DIBAS: MES = Val(Mid$(FECHATEX$(FF), 4, 2))
                       If MES < 1 Then GoTo 95
                       If MES > 12 Then GoTo 95
           J% = J% + 1
           INIPER%(J%) = DIBAS
        Wend
        INIPER%(1) = HOO%
        '
    If SEMANAL% = 1 Then  ' ***** caso programacion semanal
        '
        LOREC% = 208
        CACOL% = 52
        '
        DMX%(1) = 31
        DMX%(2) = 28
        DMX%(3) = 31
        DMX%(4) = 30
        DMX%(5) = 31
        DMX%(6) = 30
        DMX%(7) = 31
        DMX%(8) = 31
        DMX%(9) = 30
        DMX%(10) = 31
        DMX%(11) = 30
        DMX%(12) = 31
        '
        FFX = HOO%
        SEMANACT% = SEMANU%(FFX, A, B)
        COSEMACT% = A
        '
        FFX = HOO%
        SEMAHOO% = SEMANU%(FFX, A, B)
        DIBAS = A
        SMNA% = SEMAHO%
        J% = 1
        INIPER%(1) = DIBAS
'FEX = INIPER%(1): FFSS$ = FECHATEX$(FEX)
        While J% < CACOL% + 1
98         DIBAS = DIBAS + 1
           FF = DIBAS: DIAI% = Val(Left$(FECHATEX$(FF), 2))
                       If DIAI% < 1 Then GoTo 98
           FF = DIBAS: MESI% = Val(Mid$(FECHATEX$(FF), 4, 2))
                       If MESI% < 1 Then GoTo 98
                       If MESI% > 12 Then GoTo 98
99         MESX% = MESI%
           DIAMES% = DMX%(MESX%)
           If DIAMES% = 28 Then
               If (Val(Mid$(FECHATEX$(FF), 7, 2)) Mod 4) = 0 Then
                   DIAMES% = 29
               End If
           End If
           '
           If DIAI% > DIAMES% Then GoTo 98

           FF = DIBAS
           SMN% = SEMANU%(FF, A, B)
           If SMN% = SMNA% Then GoTo 98
           '
           FF = DIBAS
           FEX = FF
           If DIASEM%(FF) <> 1 Then GoTo 98

           SMNA% = SMN%
           J% = J% + 1: INIPER%(J%) = DIBAS
'FEX = INIPER%(J%): FFSS$ = FECHATEX$(FEX)
           DIBAS = DIBAS + 6
        Wend
    End If
    '
    Screen.MousePointer = 1
    '
End Sub
'
Sub COBERFA(FFI%, CI%, REFE$, Tipo%, CAN)
      '
      Static MESU$(12)
      Static CTX%
      Static COMIEMACT%, COMIACT$, INIPROXE%
      '
      If CTX% = 0 Then
         '
         CACO& = 0
         '
         MESU$(1) = "ENERO"
         MESU$(2) = "FEBRERO"
         MESU$(3) = "MARZO"
         MESU$(4) = "ABRIL"
         MESU$(5) = "MAYO"
         MESU$(6) = "JUNIO"
         MESU$(7) = "JULIO"
         MESU$(8) = "AGOSTO"
         MESU$(9) = "SETIEMBRE"
         MESU$(10) = "OCTUBRE"
         MESU$(11) = "NOVIEMBRE"
         MESU$(12) = "DICIEMBRE"
         '
         For NMM% = 1 To NUMAS%
            PUNTE$(NMM%) = ""
            PUNRE$(NMM%) = ""
            PUNFE$(NMM%) = ""
         Next NMM%
         '
         HOII% = HOY(HOS$)
         COMIACT$ = "01" + Mid$(HOS$, 3)
         COMIEMACT% = FECHANUM(COMIACT$)
         INIPROXE% = COMIEMACT%
         CTX% = 1
      End If
      '
      If ACTICOD%(CI%) < 1 Then Exit Sub
      If Abs(CAN) < 0.000005 Then Exit Sub
      '
      If SEMANAL% = 1 Then
         FHOY = HOY(HOS$)
         SMXYZ$ = SEMANU%(FHOY, DELX, ALX)
         INIPROXE% = DELX                         ' PRIMER DÍA DE LA SEMANA ACTUAL
         '
         FFX = FFI%
         If FFX < INIPROXE% Then FFX = INIPROXE%  ' LO FUERZA A LA SEMANA ACTUAL
         SMN% = SEMANU%(FFX, A, B)
         MESANIO$ = "S." + Mid$(Str$(SMN%), 2) + " - " + FECHATEX$(A)
         MESANIOT% = FFI%
         '
         FF$ = FECHATEX$(FFX)
         ANI% = Val(Mid$(FF$, 7, 2))
         MES% = Val(Mid$(FF$, 4, 2))
         ANIOX% = ANI%: If ANIOX% > 15 Then ANIOX% = 0
         MESANXX& = 30 * (54 * ANIOX% + SMN%) + Tipo%
         If MESANXX& > 32767 Then MESANXX& = 32767
         If MESANXX& < (-32767) Then MESANXX& = (-32767)
         MESANIOI% = MESANXX&
         '
         GoTo 25902
      End If
      '
      ' ESTA PARTE ES PARA PROGRAMACIÓN MENSUAL
      FFX = FFI%
      If FFX < INIPROXE% Then FFX = INIPROXE%
      FF$ = FECHATEX$(FFX)
      ANI% = Val(Mid$(FF$, 7, 2))
      ANIO% = Val(Mid$(FF$, 7, 2)) - 90
      If ANIO% < 0 Then ANIO% = ANIO% + 100
      MES% = Val(Mid$(FF$, 4, 2))
      MESANIO$ = MESU$(MES%) + " " + Mid$(FF$, 7, 2)
      MESANIOT% = FECHANUM("01" + Mid$(FF$, 3))
      ANIOX% = ANI%: If ANIOX% > 20 Then ANIOX% = 0
      MESANIOI% = 100 * (13 * ANIOX% + MES%) + Tipo%
      '
25902 ' ESTA PARTE ES PARA ACUMULAR SI COINCIDE FECHAM CODIGO Y REFERENCIA
      CARC% = 0
      REF1$ = TRIM$(REFE$)
      For KKII% = 1 To Len(REF1$)
        CARC% = CARC% + Asc(Mid$(REF1$, KKII%, 1)) - 128
      Next KKII%
      '
      FUX% = Len(PUNTE$(CI%)) / 4
      For UX% = 1 To FUX%
        If CVI(Mid$(PUNRE$(CI%), 2 * UX% - 1, 2)) = CARC% Then
          If CVI(Mid$(PUNFE$(CI%), 2 * UX% - 1, 2)) = FFI% Then
            '
            REG& = CVI(Mid$(PUNTE$(CI%), 4 * UX% - 3, 2))
            FF0$ = REGLEIDO$("COBERFA", REG&)
            F1$ = Left$(FF0$, 2)
            F2$ = Mid$(FF0$, 3, 2)
            F4$ = Mid$(FF0$, 5, 16)
            F5$ = Mid$(FF0$, 21, 30)
            F0$ = Mid$(FF0$, 51, 2)
            F6$ = Mid$(FF0$, 53, 4)
            F7$ = Mid$(FF0$, 57, 4)
            '
            If CVI(F1$) = FFI% Then
              If CVI(F2$) = CI% Then
                If TRIM$(F4$) = TRIM$(MESANIO$) Then
                  If TRIM$(F5$) = TRIM$(REFE$) Then
                    If CVI(F0$) = Tipo% Then
                      If Sgn(CAN) = Sgn(CVS(F6$) - CVS(F7$)) Then
                        If CAN >= 0 Then
                            Call REPLA(FF0$, MKS$(CVS(F6$) + CAN), 53, 4)
                            F6$ = Mid$(FF0$, 53, 4)
                          ElseIf CAN < 0 Then
                            Call REPLA(FF0$, MKS$(CVS(F7$) - CAN), 57, 4)
                            F7$ = Mid$(FF0$, 57, 4)
                        End If
                        Call GRAREG("COBERFA", FF0$, REG&)
                        Exit Sub
                      End If
                    End If
                  End If
                End If
              End If
            End If
          End If
        End If
      Next UX%
      '
      ' ARMA LA REFERENCIA CON LA FECHA ORIGINAL PROGRAMADA EN EL TEXTO
      FFI1% = FFI%
      REFEI1$ = REFE$
      If FFI% < INIPROXE% Then
         FEX = FFI1%
         REFEI1$ = TRIM$(REFE$) + " (" + Left$(FECHATEX$(FEX), 5) + ")"
         FFI1% = INIPROXE%
         FEX = INIPROXE%
         AA1$ = FECHATEX$(FEX)
      End If
      '
      ' GRABA EL MOVIMIENTO PROYECTADO
      FF0$ = REGBLAN$("COBERFA")
      Call REPLA(FF0$, MKI$(FFI1%), 1, 2)
      Call REPLA(FF0$, MKI$(CI%), 3, 2)
      Call REPLA(FF0$, MESANIO$, 5, 16)
      Call REPLA(FF0$, REFEI1$, 21, 30)
      Call REPLA(FF0$, MKI$(Tipo%), 51, 2)
      If CAN >= 0 Then
            Call REPLA(FF0$, MKS$(CAN) + MKS$(0) + MKS$(0), 53, 12)
          Else
            Call REPLA(FF0$, MKS$(0) + MKS$(Abs(CAN)) + MKS$(0), 53, 12)
      End If
      Call REPLA(FF0$, MKI$(FFI1%), 61, 2)
      CACO& = CACO& + 1
      RECOBE& = CACO&
      Call GRAREG("COBERFA", FF0$, RECOBE&)
      '
      PUNTE$(CI%) = PUNTE$(CI%) + MKS$(CACO&)
      PUNRE$(CI%) = PUNRE$(CI%) + MKI$(CARC%)
      PUNFE$(CI%) = PUNFE$(CI%) + MKI$(FFI1%)
      '
      ' ARMA LA LISTA DE SELECCIÓN DE PERÍODOS PROGRAMADOS
      If TRIM$(MESANIO$) <> "" Then
        For KKK% = 1 To 52   '52 To 1 Step -1
          If INIPER%(KKK%) > 0 Then
            If MESANIOT% < INIPER%(KKK% + 1) Then
              If DEPER$(KKK%) = "" Then
                DEPER$(KKK%) = MESANIO$
              End If
              Exit Sub
            End If
          End If
        Next KKK%
      End If
      '
End Sub



Sub PRODISP(FF%, CI%, Tipo%, CAN)

Static SEMANAL%, CTX%, CAPER%, N9%, N98%, F1$, F2$, F4$, F5$, F7$, F8$, F9$, F10$, G1$, G2$, LU$
Static MES$()

If CTX% = 0 Then
      LU$ = LUDAT$
'      CT$ = UCase$(CTRLFAC$("PERIODO"))
      SEMANAL% = 0: If InStr(CT$, "SEMAN") > 0 Then SEMANAL% = 1

35901 N9% = FreeFile
'      OPEN "R",#N9%,LU$+"PERIPRO.DAT",32
'      FIELD #N9%,6 AS G1$,26 AS G2$
      If LOF(N9%) > 0 Then
         Close #N9%
         Kill LU$ + "PERIPRO.DAT"
         GoTo 35901
      End If
      CAPER% = 0

 '     Dim MES$(12)
'35902 Data ENERO, FEBRERO, MARZO, ABRIL, MAYO, JUNIO, JULIO, AGOSTO, SETIEMBRE, OCTUBRE, NOVIEMBRE, DICIEMBRE
'      RESTORE 35902
'      For M% = 1 To 12
'        READ MES$(M%)
'      Next M%

      CTX% = 1
End If

If SEMANAL% = 1 Then
'    FF = FF%: FF$ = FECHATEX$(FF)
'    FF = FF%: SMN% = SEMANU%(FF, A, B)
'    AN$ = Mid$(FF$, 7, 2): If Left$(AN$, 1) = "9" Then AN$ = " " + AN$ Else AN$ = "1" + AN$
'    SMN$ = Mid$(Str$(SMN%), 2): While Len(SMN$) < 2: SMN$ = "0" + SMN$: Wend
'    CODPER$ = AN$ + "-" + SMN$
    PERIO$ = "S." + Mid$(Str$(SMN%), 2) + " - "
'    A$ = FECHATEX$(A): B$ = FECHATEX$(B)
    If Mid$(A$, 3) = Mid$(B$, 3) Then A$ = Left$(A$, 2)
    If Mid$(A$, 6) = Mid$(B$, 6) Then A$ = Left$(A$, 5)
    If Left$(A$, 1) = "0" Then A$ = Mid$(A$, 2)
    PERIO$ = PERIO$ + A$ + " al " + B$
    GoTo 35905
End If

'    FF = FF%: FF$ = FECHATEX$(FF)
'    AN$ = Mid$(FF$, 7, 2): If Left$(AN$, 1) = "9" Then AN$ = " " + AN$ Else AN$ = "1" + AN$
'    Me = Mid$(FF$, 4, 2): MES% = Val(Me)
    CODPER$ = AN$ + "-" + Mexx
'    PERIO$ = MES$(MES%) + " " + AN$

35905 '

For G% = CAPER% To 1 Step -1
'    GET #N9%,G%
    If G1$ = CODPER$ Then GoTo 35990
Next G%

For G% = CAPER% To 1 Step -1
'    GET #N9%,G%
    If G1$ <= CODPER$ Then GoTo 35910
'    PUT #N9%,G%+1
Next G%
G% = 0

35910 '
G% = G% + 1
LSet G1$ = CODPER$
LSet G2$ = PERIO$
'PUT #N9%,G%
CAPER% = CAPER% + 1

35990 '
End Sub

Sub CARPROSEL()
   '
   Call COPYSTRU("PROVED1", "PROVEPLA")
   ASELPRO$ = "PROVEPLA"
   Screen.MousePointer = 11
   PPXXZ$ = String$(32767, 0)
   For U& = 1 To ULTREG("PLACOMP")
     XX$ = REGLEIDO$("PLACOMP", U&)
     NPRX% = CVI(Mid$(XX$, 33, 2))
     If NPRX% > 0 Then
       Mid$(PPXXZ$, NPRX%, 1) = Chr$(1)
     End If
   Next U&
   '
   JJ& = 0
   For U& = 1 To ULTREG&("ACREDOR")
     XX$ = REGLEIDO$("ACREDOR", U&)
     NPRX% = CVI(Left$(XX$, 2))
     If NPRX% > 0 Then
       If Asc(Mid$(PPXXZ$, NPRX%, 1)) > 0 Then
         JJ& = JJ& + 1
         Call GRAREG("PROVEPLA", XX$, JJ&)
       End If
     End If
   Next U&
   '
   Call SETULTREG("PROVEPLA", JJ&)
   Call CIERRARCH("PROVEPLA")
   '
   Screen.MousePointer = 1
   '
End Sub

'
'
'
Sub CANECOM()         ' ESTE ES EL VIGENTE A PARTIR DE 17 NOV 2007
                      ' SE INTRODUJO FORZAR EL PROCESAMIENTO EN ORDEN DE NIVELES
                      ' DE EXPLOSIÓN POR EL PROBLEMA DE ACM DE AGOLPAR TODA LA
                      ' PROGRAMACION EN LA PRIMERA SEMANA.
                      '
     Dim BAGRE() As String
     Dim CLICONPEDID%(32767)    'Este Vector registra los Clientes con Pedidos Clientes.
     Dim HAYPLAFAB%(32767)
     '
     EPAC$ = TRIM$(UCase$(EMPREAC$))
     UDIM& = 2048
     ReDim Preserve BAGRE$(UDIM&)
     LU$ = LUDAT$
     HOO% = HOY(HO$)
     CACO& = 0
     '
     SUPRICO% = 0: If FORCANEC07.Check8.Value = 1 Then SUPRICO% = 1
     SUPRIFA% = 0: If FORCANEC07.Check9.Value = 1 Then SUPRIFA% = 1
     SUPRIPE% = 0: If FORCANEC07.Check10.Value = 1 Then SUPRIPE% = 1
     SUPRILECON% = 0: If FORCANEC07.Check11.Value = 1 Then SUPRILECON% = 1
     ' verifica compras pendientes de aprobación
     Call ABRECOMPRAS
     If SUPRICO% > 0 Then GoTo 90
'     '
     SQLX$ = " SELECT Nume_Ocom, Refe_Ocom,ToBru_Ocom FROM OCOMENCA"
     SQLX$ = SQLX$ + " WHERE Stat_Ocom < 0 "
     '
     Dim VEROCTEMP As ADODB.Recordset
     Set VEROCTEMP = New ADODB.Recordset
     VEROCTEMP.CursorType = adOpenDynamic
     VEROCTEMP.LockType = adLockReadOnly
     If InStr(EPAC$, "GABAL") > 0 Then
          VEROCTEMP.Open SQLX$, FLEXCONN, , , adCmdText
        Else
          VEROCTEMP.Open FILTSQL$(SQLX$), FLEXCONN, , , adCmdText
     End If
     '
     With VEROCTEMP
        On Error Resume Next
        .MoveLast
        If Err Then GoTo 90    ' no hay pendientes de aprobación
        On Error GoTo 0
        TTXX$ = "Existen Ordenes de Compra Pendientes de Aprobación.\   \Estas Compras NO Serán Tenidas en Cuenta\en el Proceso de Cálculo MRP.\Puede Optar entre Cancelar Cálculo o Continuar\\Cancelar Cálculo\Continuar"
        If COMALTER%(TTXX$) <> 2 Then GoTo 9999
     End With
     '
90   On Error GoTo 0
     Screen.MousePointer = 11
     '
100  ' ************************* ABRIENDO ARCHIVOS MAESTROS
     '
     FORCANEC07.Check1.Value = 2
     Screen.MousePointer = 11
     NMS& = NUMAS%
     '               NVI% = NUINV%
     '               N1% = FILENBR%("MASTER")
     '               N1% = FILEOPEN%(LUDAT$ + "MASTER.DAT", 0, 128)
     '               N9% = FILENBR%("SALSTOCK")
     '               N9% = FILEOPEN%(LUDAT$ + "SALSTOCK.DAT", 0, 40)
     '
     '               GESUBOR% = 0
     '               If CONTROL$("", "GESUBOR") = "SI" Then
     '                 GESUBOR% = 1
     '               End If
     '
     'GoSub 5300                           ' estrunue - estrudos
     FORCANEC07.Check1.Value = 1
     '
     'If NESTRU& < 1 Then
     '   Call MENSERR(24, "Proceso Imposible.\ \Faltan Definir Fórmulas y\Estructuras de Producto.")
     '   Call FINAL("")
     'End If
     '
     PLADISTA% = Val(FORCANEC07.Text1.TEXT)
     If PLADISTA% > 6 Then PLADISTA% = 6
     '
'    **** DEPURA PROGRAMA DE ENTREGAS POR ENTREGAS REALIZADAS
     '
     APROGRA$ = "PROGENT"
     If MODOPLAN% = 1 Then APROGRA$ = "PROGFAB"
     '
     ' ESTO QUE SIGUE ES PARA NO ALTERAR EL PROGRAMA ORIGINAL DE ENTREGAS
     ' SE CONFECCIONA UNA IMAGEN EN PROGCALC Y SE TRABAJA SOBRE ESTA
     Call COPYSTRU(APROGRA$, "PROGCALC")
     JJL1& = 0
     For i& = 1 To ULTREG&(APROGRA$)
       ZZ$ = REGLEIDO$(APROGRA$, i&)
       JJL1& = JJL1& + 1
       Call GRAREG("PROGCALC", ZZ$, JJL1&)
     Next i&
     Call SETULTREG("PROGCALC", JJL1&)
     Call CIERRARCH("PROGCALC")
     Call CIERRARCH(APROGRA$)
     APROGRA$ = "PROGCALC"
     ' FIN CONFECCION DE IMAGEN PROGENT.
     '
     UREPRO& = ULTREG&(APROGRA$)
     JJ& = 0
     For i& = 1 To UREPRO&
       Call TRACE(20, i&, FIN&)
       ZZ$ = REGLEIDO$(APROGRA$, i&)
       If CVS(Mid$(ZZ$, 7, 4)) > 0.05 Then
         Call REPLA(ZZ$, MKS$(0), 11, 4)
         Call REPLA(ZZ$, MKS$(0), 19, 4)
         JJ& = JJ& + 1
         Call GRAREG(APROGRA$, ZZ$, JJ&)
       End If
       If MODOPLAN% = 1 Then
         ' marca items programados para no considerar stock como disponibilidad
         CIXI% = CVI(Mid$(ZZ$, 1, 2))
         If CIXI% > 0 Then
           If CIXI% <= NUMAS% Then
             CAPROBRU = CVS(Mid$(ZZ$, 7, 4))
             If CAPROBRU > 0.05 Then
               HAYPLAFAB%(CIXI%) = 1
             End If
           End If
         End If
         '
       End If
     Next i&
     Call SETULTREG(APROGRA$, JJ&)
     Call CIERRARCH(APROGRA$)
     If JJ& < 1 Then
        On Error Resume Next
        Kill LUDAT$ + APROGRA$ + ".DAT"
        On Error GoTo 0
     End If
     '
' LEER ENTREGAS DEL MES PARA NETEAR PROGRAMA DE ENTREGAS
     If MODOPLAN% = 1 Then GoTo 65   ' SI POR FABRICACION NO NETEA
     '
     CMA% = COMESACT%
     Call ABRESTOCKS
     CMAX = FETEXCOR1$(CMA%)
     'SQLX$ = " SELECT * FROM MOVISTO "
     SQLX$ = " SELECT * FROM MOVISTO WITH(NOLOCK)"
     SQLX$ = SQLX$ + "WHERE FechMov >= '" & CMAX & "' "
     SQLX$ = SQLX$ + "AND (CodiMovi = 'RT' Or CodiMovi = 'DV')"
     '
     Dim MoviTemp As ADODB.Recordset
     Set MoviTemp = New ADODB.Recordset
     MoviTemp.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockReadOnly, adCmdText
     With MoviTemp
         Do While Not .EOF
           CIXI% = !COD_INTE
           If CIXI% > 0 Then
             If CIXI% <= NUMAS% Then
             AAA = CDbl(!FechMov)
               CAENTREG = !CANTSALID - !CantIngre
               '
               RFF$ = RECFILT$(APROGRA$, 1, MKI$(CIXI%))
               For KK1& = 1 To Len(RFF$) Step 4
                 REPRO& = CVS(Mid$(RFF$, KK1&, 4))
                 ZZ$ = REGLEIDO$(APROGRA$, REPRO&)
                 ACUENTRE = CVS(Mid$(ZZ$, 11, 4))
                 ENTREMA = CVS(Mid$(ZZ$, 7, 4)) - ACUENTRE
                 If CAENTREG <= ENTREMA Then
                    Call REPLA(ZZ$, MKS$(CAENTREG + ACUENTRE), 11, 4)
                    CAENTREG = 0
                  Else
                    Call REPLA(ZZ$, MKS$(ENTREMA + ACUENTRE), 11, 4)
                    CAENTREG = CAENTREG - ENTREMA
                 End If
                 Call GRAREG(APROGRA$, ZZ$, REPRO&)
                 If CAENTREG < 0.05 Then GoTo 102
               Next KK1&
             End If
           End If
           '
102      .MoveNext
         Loop
     End With
     '
'    **** AGREGA PEDIDOS PENDIENTES DE ENTREGA
'    **** DESCARTANDO L0S CORRESPONDIENTES A REPARACIONES
      '
65    If MODOPLAN% = 1 Then
        For CIXI% = 1 To NUMAS%: ACTICOD%(CIXI%) = 1: Next CIXI%
        GoTo 75
      End If
      '
      FIN& = ULTREG&(APROGRA$)
      RFF$ = ""
      For U& = 1 To FIN&
        X$ = REGLEIDO$(APROGRA$, U&)
        RFF$ = RFF$ + Left$(X$, 2)
      Next U&
      '
     Call ABREPEDCLI
     Condi$ = "(Status >= 0 AND Status < 8) "      ' para que no este anulada ni cumplida ni pendiente de aprobacion
     Condi$ = Condi$ + "AND CantEnt < (CanPed - 0.005) "
     Condi$ = Condi$ + " AND VenRep = 'V' "
     '
     SQLX$ = " SELECT * FROM PEDDETA WITH(NOLOCK)"
     'SQLX$ = SQLX$ + "WHERE CodiInt = " & CI% & " "
     SQLX$ = SQLX$ + "WHERE " & Condi$
     SQLX$ = SQLX$ + " ORDER BY FeSolEnt DESC "
     '
     FINPE& = CantRegistros("PEDDETA", Condi$)
     Dim PEDTEMP As ADODB.Recordset
     If InStr(EPAC$, "GABAL") > 0 Then
          Set PEDTEMP = FLEXCONN.Execute(SQLX$)
        Else
          Set PEDTEMP = FLEXCONN.Execute(FILTSQL$(SQLX$))
     End If
     '\\\OT-05-0631-WAR-20/09/05///
     INICLICON% = 0
     If UCase$(Control$("CANECOM", "FILPECLI")) <> "SI" Then INICLICON% = 1
     ' SI NO FILTRA POR CLIENTES CON PEDIDOS, ENTONCES TOMA TODOS
     For JJ1% = 1 To 32766 'PONGO EN 0 EL VECTOR
       CLICONPEDID%(JJ1%) = INICLICON%
     Next JJ1%
     FILPEDIPEN% = 0
     If INICLICON% = 1 Then
        For CIXI% = 1 To NUMAS%: ACTICOD%(CIXI%) = 1: Next CIXI%
        GoTo 50
     End If
     '
     ' CARGA UN VECTOR CON LOS CLIENTES CON PEDIDOS PENDIENTES
     FILPEDIPEN% = 1: KU& = 0
     For CIXI% = 1 To NUMAS%: ACTICOD%(CIXI%) = 0: Next CIXI%
     With PEDTEMP
       On Error Resume Next
       .MoveFirst
       If Err < 1 Then
         Do While Not .EOF
          KU& = KU& + 1
          Call TRACE(20, KU&, FINPE&)
          NCLIE1 = !NROCLIE
          If Val(NCLIE1) > 0 Then
            If Val(NCLIE1) <= 32767 Then
              CLICONPEDID%(NCLIE1) = 1 'MARCO LOS CLIENTES CON PEDIDOS
            End If
          End If
          .MoveNext
         Loop
       End If
     End With
     '
     For JL& = 1 To 32767
      If CLICONPEDID%(JL&) = 1 Then
        NCLIE2$ = FORMATNUM$(JL&, "F6.0")
        RASOCLIE2$ = rasocli$(JL&)
        Z$ = Space(128)
        Call REPLA(Z$, NCLIE2$, 1, 6)
        Call REPLA(Z$, RASOCLIE2$, 15, 64)
        With CLICOPE07
          .LICOPEN.AddItem Z$
        End With
      End If
     Next JL&
     Screen.MousePointer = 1
     If CLICOPE07.LICOPEN.ListCount > 0 Then
       CLICOPE07.Show 1
       Call FRESHALL
       If CLICOPE07.Label2.Caption <> "OK" Then
         Call COMUNI("Cancelado !!!")
         GoTo 9999
       End If
       '
       ' ACA VUELVE A PONER EN '0' EL VECTOR DE CLIENTES
       For JL& = 1 To 32767
         CLICONPEDID%(JL&) = 0
       Next JL&
       '
       Call COMUNI("El Cálculo MRP-II sólo Considerará Aquellos Items de Stock\que Sean Componentes de Alguno de los Productos de los Pedidos\Pendientes de Entrega de los Clientes Seleccionados.\   \Todos los Restantes Artículos se Suprimen, Independientemente de su\Situación de Stocks Actual o Mínimo y Ordenes de Compra o Fabricación en Curso.")
       ' ACA RE-CARGA EL VECTOR PONIENDO EN '1' LOS CLIENTES SELECCIONADOS
       CLISELE$ = ""
       For JL& = 1 To CLICOPE07.LICOPEN.ListCount
         If CLICOPE07.LICOPEN.Selected(JL& - 1) = True Then
           ICLIE1& = Val(Left$(CLICOPE07.LICOPEN.List(JL& - 1), 6))
           If ICLIE1& > 0 Then
             If ICLIE1& <= 32767 Then
               CLICONPEDID%(ICLIE1&) = 1
               If CLISELE$ <> "" Then
                  CLISELE$ = CLISELE$ + " + "
               End If
               NCXX% = ICLIE1&
               CLISELE$ = CLISELE$ + TRIM$(ABRECLI$(NCXX%)) + "(" + TRIM$(Str$(NCXX%)) + ")"
             End If
           End If
         End If
       Next JL&
     End If
     Call GRACONTROL("CANECOM", "CLISELE", CLISELE$)
     '
50   Screen.MousePointer = 11
     '\\\OT-05-0631-WAR-FIN///
     KU& = 0
     With PEDTEMP
       On Error Resume Next
       .MoveFirst
       If Err < 1 Then
         Do While Not .EOF
           '
           KU& = KU& + 1
           Call TRACE(20, KU&, FINPE&)
           ' AQUI DESCARTA LOS PEDIDOS DE CLIENTES NO MARCADOS
           NCLIE1 = !NROCLIE
           If Val(NCLIE1) > 0 Then
             If Val(NCLIE1) <= 32767 Then
               If CLICONPEDID%(NCLIE1) <> 1 Then
                 GoTo 103
               End If
             End If
           End If
           '
           CIXI% = !CODIINT
           Call EXPLOMUL(CIXI%, 1)
           For JJ1% = 1 To CONTAMUL%
             CIX2% = CIK%(JJ1%)
             ACTICOD%(CIX2%) = 2
           Next JJ1%
           '
           NCX1% = NCLIE1
           CAPEN = !CanPed - !CANTENT
           If SUPRIPE% = 1 Then CAPEN = 0
           If CAPEN > 0 Then
             FENTRESOL% = CVINT%(!FeSolEnt)
             For KK1& = 1 To Len(RFF$) Step 2    ' - 3 To 1 Step -4
               If CVI(Mid$(RFF$, KK1, 2)) = CIXI% Then
                 REPRO& = (KK1& + 1) / 2 ' CVS(Mid$(RFF$, KK1&, 4))
                 ZZ$ = REGLEIDO$(APROGRA$, REPRO&)
                 FEPROG% = CVI(Mid$(ZZ$, 3, 2))
                 If COMIENFE%(FEPROG%) <= COMIENFE%(FENTRESOL%) Then
                   CAPEPROMAX = CVS(Mid$(ZZ$, 7, 4)) - CVS(Mid$(ZZ$, 11, 4)) - CVS(Mid$(ZZ$, 19, 4))
                   If CAPEPROMAX < 0 Then CAPEPROMAX = 0
                   '         PROGRAMA BRUTO - ENTREGAS - PEDIDOS PROGRAMADOS
                   If CAPEPROMAX >= CAPEN Then
                       CAPEDID = CVS(Mid$(ZZ$, 19, 4)) + CAPEN
                       Call REPLA(ZZ$, MKS$(CAPEDID), 19, 4)
                       Call GRAREG(APROGRA$, ZZ$, REPRO&)
                       CAPEN = 0
                       GoTo 103
                     Else
                       CAPEDID = CVS(Mid$(ZZ$, 19, 4)) + CAPEPROMAX
                       Call REPLA(ZZ$, MKS$(CAPEDID), 19, 4)
                       Call GRAREG(APROGRA$, ZZ$, REPRO&)
                       CAPEN = CAPEN - CAPEPROMAX
                   End If
                 End If
               End If
             Next KK1&
             '
             If CAPEN > 0.05 Then
               FENTRESO% = FENTRESOL%
               If FENTRESO% < HOO% Then FENTRESO% = HOO%
               ZZ$ = REGBLAN$(APROGRA$)
               Call REPLA(ZZ$, MKI$(CIXI%), 1, 2)
               Call REPLA(ZZ$, MKI$(FENTRESO%), 3, 2)
               Call REPLA(ZZ$, MKI$(NCX1%), 5, 2)
               Call REPLA(ZZ$, MKS$(CAPEN), 19, 4)
               Call REGAPP(APROGRA$, ZZ$)
               RFF$ = RFF$ + MKI$(CIXI%)
                 'Call CIERRARCH(APROGRA$)
             End If
           End If
           '
103      .MoveNext
         Loop
         '
       End If
     End With
75   Call CIERRARCH(APROGRA$)
     '
' ************************* pone ACTICOD% = 0 para los items de baja
     Call ABREMASTER
     SQLX$ = " SELECT Codint, Status FROM MASTER WITH(NOLOCK) "
     SQLX$ = SQLX$ + "WHERE Status < 0 "          ' SELECCIONA LOS DE BAJA
     Dim IDEBAJA As ADODB.Recordset
     Set IDEBAJA = FLEXCONN.Execute(FILTSQL$(SQLX$))
     '
     NMS1% = NUMAS%
     With IDEBAJA
       Do While Not .EOF
         CIXI% = !CODINT
         If CIXI% > 0 Then
            If CIXI% <= NMS1% Then
               ACTICOD%(CIXI%) = 0
            End If
         End If
       .MoveNext
       Loop
     End With
     IDEBAJA.Close
     Set IDEBAJA = Nothing
     '
' ************** lee existencias, stock minimo y lote reposicion *******************
     FORCANEC07.Check2.Value = 2
     Screen.MousePointer = 11
     '
     ' determina items fanricados con dpi%()=1
     FIN& = NUMAS%
     For i& = 1 To FIN&
       Call TRACE(20, i&, FIN&)
       CI% = i&
       Call AREXPLO(CI%, 1, 3)
       If CAIT% > 0 Then DPI%(CI%) = 1
105  Next i&
     '
     Call CALNIVEGEN        ' DETERMINA EL NIVEL DE EXPLOSION GENERAL DE CADA ITEM
     '
     HAYNE% = 0
     MODO% = 0
     MODOCAX$ = FORCANEC07.Option3.Caption
     If FORCANEC07.Option1.Value = True Then
         MODO% = 1
         MODOCAX$ = FORCANEC07.Option1.Caption
       ElseIf FORCANEC07.Option2.Value = True Then
         MODO% = 2
         MODOCAX$ = FORCANEC07.Option2.Caption
       ElseIf FORCANEC07.Option6.Value = True Then
         MODO% = 3
         MODOCAX$ = FORCANEC07.Option6.Caption
     End If
     Call GRACONTROL("CANECOM", "MODOCAL", MODOCAX$)
     '
     FFXY = HOY(HO$)
     SMXX% = SEMANU%(FFXY, AA, BB)
     COSEMACT% = AA
     '
     QLN& = 0
     Call CIERRARCH("COBERFA")
     Call CIERRARCH("ORFAEMI")
     Call CIERRARCH("OCOMEMI")
     '
     For CI% = 1 To NMS&
       CIIX& = CI%
       Call TRACE(20, CIIX&, NMS&)
       SALSTO = 0
       If STATUIT%(CI%) >= 0 Then
         SALSTO = STOCKACT(CI%)
         '
         If MODO% = 1 Then SALSTO = 0
         If MODO% = 2 Then If SALSTO < 0 Then SALSTO = 0

         If SALSTO < (-0.05) Then
            HAYNE% = HAYNE% + 1
            CI1% = CI%
            TTXX$ = REGBLAN$("LINEGAT")
            Call REPLA(TTXX$, MKI$(CI%), 1, 2)
            'Call REPLA(TTXX$, DESCRIT$(CI%), 3, 30)
            
            Call REPLA(TTXX$, CSTRING$(MKS$(SALSTO), 4, 10, 1, 2), 33, 10)
            Call REPLA(TTXX$, Unimed$(CI%), 45, 4)
            '
            QLN& = QLN& + 1
            Call GRAREG("LINEGAT", TTXX$, QLN&)
            '
         End If
         '
         STMIN = STOMIN(CI%): If MODO% = 1 Or MODO% = 3 Then STMIN = 0
         LINDIVI(CI%) = LINDIVIS(CI%): If LINDIVI(CI%) < 1 Then LINDIVI(CI%) = 1
         REPOS = LOTESTAN(CI%)      ' SUJETO A SUPRIME LOTE ECONOMICO
         DP% = DPI%(CI%): If DP% < 0 Or DP% > 3 Then DP% = 0

         If REPOS < 0.5 Then REPOS = LOSTN(DP%)
         If REPOS < 0.5 Then REPOS = REPOMIN
         If SUPRILECON% > 0 Then REPOS = 0  ' SUPRIME LOTE ECONOMICO
         '
         DISPACT(CI%) = SALSTO - STMIN
         If HAYPLAFAB%(CI%) > 0 Then DISPACT(CI%) = 0
         LOREPOS(CI%) = REPOS
         '
         If HAYPLAFAB%(CI%) < 1 Then
            If Abs(SALSTO) > 0.05 Then
              Call COBERFA(COSEMACT%, CI%, "Stock al " + HO$, 10, SALSTO)
            End If
            '
            If Abs(STMIN) > 0.05 Then
              Call COBERFA(COSEMACT%, CI%, "Stock Minimo del Item", 20, (-1) * STMIN)
            End If
         End If
       End If
     Next CI%
     '
     Call SETULTREG("LINEGAT", QLN&)
     Call CIERRARCH("LINEGAT")
     '
     If HAYNE% > 0 Then
        Call FILESORT("LINEGAT", "", 1, 1, "ASC")
        A% = COMALTER%("ATENCION: Se Detectaron Existencias Negativas\          en" + Str$(HAYNE%) + " Items de Stock.\\Revisar\Cancelar\Continuar")
        On A% GoTo 107, 106, 108
106     Call CIERRARCH("*.*")
        Screen.MousePointer = 1
        Exit Sub
107     Call FRESHECHO("LINEGAT")
        Call SELECHO("LINEGAT")
        '
        If ALTERNA%("Continuar Proceso") <> 1 Then
          Call CIERRARCH("*.*")
          Screen.MousePointer = 1
          Exit Sub
        End If
     End If
     '
108  FORCANEC07.Check2.Value = 1
     FORCANEC07.Refresh
     FORCANEC07.Refresh
     '
     '
     ' revisa las ordenes de compra pendientes de recepcion
     Call ABRECOMPRAS
     If SUPRICO% = 1 Then GoTo 325
     '
     'SQLX$ = " SELECT * FROM OCOMDETA "
     SQLX$ = " SELECT * FROM OCOMDETA WITH(NOLOCK) "
     SQLX$ = SQLX$ + "WHERE Cod_Inte > 0 "    ' SOLO MATERIAL CODIFICADO
     SQLX$ = SQLX$ + "AND Stat_Ocom < 8 "   ' para que no este anulada ni cumplida
     SQLX$ = SQLX$ + "AND  Stat_Ocom <> -1 " 'Para que no  hayan O/C's sin aprobar
     SQLX$ = SQLX$ + "AND Cant_Rec < (Cant_Ocom - 0.005) "
     '
     Dim OCOTEMP As ADODB.Recordset
     If InStr(EPAC$, "GABAL") > 0 Then
          Set OCOTEMP = FLEXCONN.Execute(SQLX$)   ' PARA QUE UNIFIQUE LAS COMPRAS
        Else
          Set OCOTEMP = FLEXCONN.Execute(FILTSQL$(SQLX$))
     End If
     With OCOTEMP
         Do While Not .EOF
           NUORSE& = XVALO(!N_OrServ)
           If NUORSE& < 1 Then
             COEFCAN = XVALO(!Coef_Unids)
             If COEFCAN < 0.005 Then COEFCAN = 1
             CATOT# = XVALO(!Cant_Ocom) * COEFCAN
             CAPRA# = XVALO(!Cant_Rec) * COEFCAN
             CAPEN = CATOT# - CAPRA#
             If CAPEN > 0 Then
               CIBX% = XVALO(!COD_INTE)
               If STATUIT%(CIBX%) >= 0 Then
                 NC% = XVALO(!Npro_Ocom)
                 NROPED& = XVALO(!NUME_OCOM)
                 FENTRE% = CVINT(!Fentre_Sol)
                 COMPROB$ = "OC." + TRIM$(Str$(NROPED&))
                 CICX% = CIBX%
                 Call COBERFA(FENTRE%, CIBX%, COMPROB$ + " " + TEL2CLI$((-1) * NC%), 13, CAPEN)
                 If HAYPLAFAB%(CICX%) < 1 Then
                    DISPACT(CICX%) = DISPACT(CICX%) + CAPEN
                 End If
               End If
             End If
           End If
         .MoveNext
         Loop
       'End If
     End With
     '
     ' revisa las boletas de recepción pendientes de aprobacion
325  Call ABRESTOCKS
     If MODO% = 1 Then GoTo 326            ' CERO STOCKS Y MINIMOS
     '
     'SQLX$ = " SELECT * FROM BOLRECEP "
     SQLX$ = " SELECT * FROM BOLRECEP WITH(NOLOCK) "
     SQLX$ = SQLX$ + "WHERE Cod_Inte > 0 "    ' SOLO MATERIAL CODIFICADO
     SQLX$ = SQLX$ + "AND Status = 0 "   ' para que no este anulada ni cumplida
     '
     Dim COBRECEPTMP As ADODB.Recordset
     If InStr(EPAC$, "GABAL") > 0 Then
          Set COBRECEPTMP = FLEXCONN.Execute(SQLX$)
        Else
          Set COBRECEPTMP = FLEXCONN.Execute(FILTSQL$(SQLX$))
     End If
     With COBRECEPTMP
         Do While Not .EOF
           CATOT# = XVALO(!CANTNOM)
           CAPRA# = XVALO(!CanToApro)
           CAPEN = CATOT# - CAPRA#
           If CAPEN > 0 Then
             CIBX% = XVALO(!COD_INTE)
             If STATUIT%(CIBX%) >= 0 Then
               NC% = XVALO(!NUMPROV)
               NROPED& = XVALO(!NUMEBORE)
               FENTRE% = CVINT(!FechRecep)
               COMPROB$ = "BR." + TRIM$(Str$(NROPED&))
               CICX% = CIBX%
               Call COBERFA(FENTRE%, CIBX%, COMPROB$ + " " + TEL2CLI$((-1) * NC%), 11, CAPEN)
               If HAYPLAFAB%(CICX%) < 1 Then
                  DISPACT(CICX%) = DISPACT(CICX%) + CAPEN
               End If
             End If
           End If
         .MoveNext
         Loop
     End With
     '
' ******** agrega ordenes de fabricacion lanzadas y en curso ******
     '
326  Call ABREORFAB
     If SUPRIFA% = 1 Then GoTo 327
     '
     'DEMORA 1:20
     'SQLX$ = " SELECT * FROM ORDEFABR "
     'SQLX$ = SQLX$ + "WHERE StatuOrf < 8 "   ' para que no este anulada ni cumplida
     'SQLX$ = SQLX$ + "AND CanApro < (CanProy - 0.005) "
     
     'DEMORA 0:45
     SQLX$ = "SELECT * FROM ORDEFABR INNER JOIN MASTER WITH(NOLOCK) "
     SQLX$ = SQLX$ + " ON ORDEFABR.COD_INTE = MASTER.CODINT "
     SQLX$ = SQLX$ + "WHERE ORDEFABR.StatuOrf < 8 "   ' para que no este anulada ni cumplida
     SQLX$ = SQLX$ + "AND MASTER.Status > 0 "
     SQLX$ = SQLX$ + "AND ORDEFABR.CanApro < (ORDEFABR.CanProy - 0.005) AND MASTER.STATUS > 0"
     FIN& = CantRegistros("ORDEFABR", "CanApro < (CanProy - 0.005)", "NOMESS")
     i& = 0
     '
     Dim OFATEMP As ADODB.Recordset
     Set OFATEMP = FLEXCONN.Execute(FILTSQL$(SQLX$))
     With OFATEMP
         Do While Not .EOF
           i& = i& + 1
           Call TRACE(24, i&, FIN&)
           CI% = XVALO(!COD_INTE)
'           If STATUIT%(CI%) >= 0 Then
              If MODOPLAN% <> 1 Or HAYPLAFAB%(CI%) < 1 Then 'descarta ordenes de fabricación productos terminados
                CACOM = XVALO(!CanProy)
                CAREC = XVALO(!CanApro)
                CAPEN = CACOM - CAREC
                If CAPEN > 0 Then
                  NROPED& = XVALO(!NUMEoRFA)
                  FEX = CVINT(!FechEnSol)
                  REFOR$ = SAFETEXT$(!Referen)
                  '
                  STATU% = XVALO(!StatuOrf)
                  If STATU% <= 2 Then ' esta proyectada o lanzada
                    '
                    NORFA$ = SAFETEXT$(!IDSUBOR)
                    NUORFA& = Val(Mid$(NORFA$, 4, 4))
                    SUBOR% = Val(Mid$(NORFA$, 9, 2))
                    FENTRE% = CVINT(!FechProyEntre) 'CVI(Mid$(F$, 147, 2))
                    If FENTRE% < HOO% Then FENTRE% = CVINT(!FechEnSol)   ' CVI(Mid$(F$, 105, 2))
                    If FENTRE% < HOO% Then FENTRE% = HOO%
                    FELANZ% = CVINT(!FechLan)   'CVI(Mid$(F$, 109, 2))
                    If FELANZ% < HOO% Then FELANZ% = HOO%
                    Stat$ = " (proyectada)"
                    If STATU% = 1 Then Stat$ = " (lanzada)"
                    If STATU% = 2 Then Stat$ = " (suspendida)"
                    '
                    DISPACT(CI%) = DISPACT(CI%) + CAPEN
                    '
                    CIBX% = CI%
                    'COD$ = CODEXT$(CI%) ' SE COMENTA PORQUE NO SE UTILIZA
                    COEF = 1
                    '
                    Call COBERFA(FENTRE%, CIBX%, NORFA$ + Stat$, 12, CAPEN)
                    '
                  End If
                End If
              End If
'          End If
         .MoveNext
         Loop
     End With
     '
327  Screen.MousePointer = 11
35   Call ABRESERVA
     If SUPRIFA% = 1 Then GoTo 328
     '
     'DEMORA 9 MIN
     'SQLX$ = " SELECT * FROM RESERVA "
     'SQLX$ = SQLX$ + "WHERE CantCons < (CantRes - 0.005) "
     '
     'DEMORA 1 MIN
     SQLX$ = "SELECT * FROM RESERVA INNER JOIN MASTER WITH(NOLOCK) "
     SQLX$ = SQLX$ + " ON RESERVA.COD_INTE = MASTER.CODINT "
     SQLX$ = SQLX$ + "WHERE RESERVA.CantCons < (RESERVA.CantRes - 0.005) "
     FIN& = CantRegistros("RESERVA", "CantCons < (CantRes - 0.005)", "NOMESS")
     i& = 0
     '
     Dim RESERTEMP As ADODB.Recordset
     If InStr(EPAC$, "GABAL") > 0 Then
          Set RESERTEMP = FLEXCONN.Execute(SQLX$)
        Else
          Set RESERTEMP = FLEXCONN.Execute(FILTSQL$(SQLX$))
     End If
     With RESERTEMP
         Do While Not .EOF
           i& = i& + 1
           Call TRACE(24, i&, FIN&)
           CI1% = XVALO(!COD_INTE)
           If CI1% > 0 Then
             TORE# = XVALO(!CANTRES)
             CACON# = XVALO(!CANTCONS)
             '
'            If TRIM$(UCase$(UNIMED$(CI1%))) = "GK" Then
             If TRIM$(UCase$(SAFETEXT$(!UMEDIDA))) = "GK" Then
               TORE# = TORE# / 1000
               CACON# = CACON# / 1000
             End If
             '
             SIGNO% = Sgn(TORE#)
             SARE = TORE# - CACON#
             If SARE * SIGNO% < 0 Then
               SARE = 0
             End If
             CAPEN = SARE
             If CAPEN > 0 Then
               NORFA$ = SAFETEXT$(!IDSUBOR)
               DISPACT(CI1%) = DISPACT(CI1%) - CAPEN
               FELANZ% = CVINT(!FechRes)
               If InStr(1, EMPREAC$, "FERVI") > 0 Then
                    CIXICONJ% = XVALO(VALOBADA("ORDEFABR", "COD_INTE", "IDSUBOR='" & NORFA$ & "'"))
                    FELANZ% = CVINT(VALOBADA("ORDEFABR", "FechLan", "IDSUBOR='" & NORFA$ & "'"))
                    PLADISPO% = XVALO(VALOBADA("FORMULAS", "PLADISPO", "codiconj=" & CIXICONJ% & " and " & "codielem=" & CI1%))
                    If PLADISPO% < 1 Then PLADISPO% = PLADISTA%
                    FF% = FELANZ%: FELANZ% = DIASANTES(FF%, PLADISPO% * 7)
               End If
               ' CONSIDERA EL LEADTIME EN DIAS DEL MAESTRO DE ARTÍCULOS.
               ' ES GENERAL Y NO ESTA RELACIONADO CON NINGUN ARTÍCULO PADRE COMO EL PLAZO DE DISPOSICION
               ' CUANDO EXISTE, PASA POR ENCIMA DEL PLAZO DE DISPOSICION
               LEADTIME% = XVALO(VALOBADA("MASTER", "LETICOM", "Codint=" & CI1%))
               If LEADTIME% > 0 Then
                   RESTO = LEADTIME% Mod 7
                   LTI_SEMANAS% = Int(LEADTIME% / 7)
                   If RESTO > 0 Then LTI_SEMANAS% = LTI_SEMANAS% + 1
                   PLADISPO% = LTI_SEMANAS%
                   FELANZ% = CVINT(!FechRes)
                   FF% = FELANZ%: FELANZ% = DIASANTES(FF%, PLADISPO% * 7)
               End If
               Call COBERFA(FELANZ%, CI1%, "Res.p/" + NORFA$, 22, (-1) * CAPEN)
             End If
           End If
         .MoveNext
         Loop
     End With
     '
     '
328  ' ******** inicializa necesidades secundarias de compra y fabricacion ********
110  FORCANEC07.Check3.Value = 2
     Screen.MousePointer = 11
     '
111  N13% = FreeFile
     On Error Resume Next
     Open LU$ + "NEFACOM.DAT" For Random Access Read Write Shared As #N13% Len = LOREC%
     If Err Then
        ERRU% = Err.Number
        On Error GoTo 0
        TXTX$ = "Error no Recuperable " & ERRU% & " al Ejecutar 'Open " & LU$ + "NEFACOM.DAT'."
        MsgBox TXTX$
        Call FINAL("")
     End If
     On Error GoTo 0
     '
     Dim PDS As String * 208
     Dim PDM As String * 72
     '
     If LOF(N13%) > 0 Then
         Close #N13%
         On Error Resume Next
         Kill LU$ + "NEFACOM.DAT"
         If Err Then
            ERRU% = Err.Number
            On Error GoTo 0
            TXTX$ = "Error no Recuperable " & ERRU% & " al Borrar 'Kill " & LU$ + "NEFACOM.DAT'."
            MsgBox TXTX$
            Call FINAL("")
         End If
         On Error GoTo 0
         GoTo 111
     End If
     '
'     MsgBox "Sin Error de Acceso a Archivo. Pulse 'Aceptar'"
     '
     For i& = 1 To NMS&
         CIIX& = i&
         Call TRACE(20, CIIX&, NMS&)
         If LOREC% = 208 Then
              Get #N13%, i&, PDS$
            Else
              Get #N13%, i&, PDM$
         End If
         PDS$ = String$(208, 0)
         PDM$ = String$(72, 0)
         If LOREC% = 208 Then
              Put #N13%, i&, PDS$
            Else
              Put #N13%, i&, PDM$
         End If
     Next i&

     FORCANEC07.Check3.Value = 1

' **************** necesidades brutas segun entregas programadas **************
120  N9% = FreeFile
     Open LU$ + "NEBRUNE.DAT" For Random Access Read Write Shared As #N9% Len = 48
     Dim B00 As String * 48
     Dim B01 As String * 48
     X1& = 1: X2& = 0

     If LOF(N9%) > 0 Then
        Close #N9%
        Kill LU$ + "NEBRUNE.DAT"
        GoTo 120
     End If

     FORCANEC07.Check4.Value = 2
     Screen.MousePointer = 11
     '
121  FIN& = NUMAS%
     For i& = 1 To FIN&
       CIIX& = i&
       Call TRACE(20, CIIX&, FIN&)
       CI% = i&
       If STATUIT%(CI%) >= 0 Then
         '
         If DISPACT(CI%) < 0 Then
           CANTU = Abs(DISPACT(CI%))
           FF% = HOO%
           '
           X2& = X2& + 1
           B00$ = String$(48, 0)
           Call REPLA(B00$, MKI$(CI%), 1, 2)
           Call REPLA(B00$, MKS$(CANTU), 3, 4)
           Call REPLA(B00$, MKI$(FF%), 7, 2)
           Call REPLA(B00$, "PROGRAMA POR MINIMOS", 9, 29)
           Call REPLA(B00$, MKS$(0), 38, 4)
           Call REPLA(B00$, MKS$(0), 42, 4)
           Call REPLA(B00$, MKS$(0), 46, 2)
           Call REPLA(B00$, " ", 48, 1)
           '
           Put #N9%, X2&, B00$
           DISPACT(CI%) = 0
           '
         End If
       End If
122  Next i&
     '
     N4% = FreeFile
     Open LU$ + APROGRA$ + ".DAT" For Random Access Read Write Shared As #N4% Len = 24
     Dim P00 As String * 24
     FIN& = LOF(N4%) / 24
     For i& = 1 To FIN&
       CIIX& = i&
       Call TRACE(20, CIIX&, FIN&)
       Get #N4%, i&, P00$
       P1$ = Left$(P00$, 2)
       CI% = CVI(P1$)
       '
       If CI% > 0 Then
         If CI% <= NUMAS% Then
           If STATUIT%(CI%) >= 0 Then
             '
             P3$ = Mid$(P00$, 3, 2)
             P6$ = Mid$(P00$, 7, 4)
             P7$ = Mid$(P00$, 11, 4)
             '
             If CVI(P3$) < COMESACT% Then
                CIX% = CVI(P1$)
                Call MENSERR(24, "Error: Codigo " + TRIM$(CODEXT$(CIX%)) + " Mal Fecha Entrega Programada")
                GoTo 129
             End If
             '
             PLAVEN = CVS(P6$)
             ENTREG = CVS(P7$)
             PEDIPE = CVS(Mid$(P00$, 19, 4))
             CANTU = PLAVEN - ENTREG
             '
             PEDINOPRO = PEDIPE + ENTREG - PLAVEN
             If PEDINOPRO < 0 Then PEDINOPRO = 0
             '
             CI% = CVI(P1$)
             FF% = CVI(P3$)
             '
             If CANTU > 0 Then
               If CI% > 0 Then
                  If CI% <= NMS& Then
                     DPI%(CI%) = 3
                  End If
               End If
             End If
             '
             If CANTU > 0 Then
               '
               X2& = X2& + 1
               B00$ = String$(48, 0)
               Call REPLA(B00$, MKI$(CI%), 1, 2)
               Call REPLA(B00$, MKS$(CANTU), 3, 4)
               Call REPLA(B00$, MKI$(FF%), 7, 2)
               Call REPLA(B00$, "PROGRAMA ENTREGAS AL " + HO$, 9, 29)
               If MODOPLAN% = 1 Then
                 Call REPLA(B00$, "PLAN DE FABRICACION AL " + HO$, 9, 29)
               End If
               If FF% < INIPER%(1) Then
                   Call REPLA(B00$, MKI$(INIPER%(1)), 7, 2)
                   Call REPLA(B00$, "ATRASO DE ENTREGA AL " + HO$, 9, 29)
               End If
               Call REPLA(B00$, MKS$(0), 38, 4)
               Call REPLA(B00$, MKS$(0), 42, 4)
               Call REPLA(B00$, MKS$(0), 46, 2)
               Call REPLA(B00$, " ", 48, 1)
               '
               Put #N9%, X2&, B00$
               '
             End If
             '
             If MODOPLAN% = 0 Then
               If PEDINOPRO > 0 Then
                 '
                 X2& = X2& + 1
                 B00$ = String$(48, 0)
                 Call REPLA(B00$, MKI$(CI%), 1, 2)
                 Call REPLA(B00$, MKS$(PEDINOPRO), 3, 4)
                 Call REPLA(B00$, MKI$(FF%), 7, 2)
                 Call REPLA(B00$, "PEDIDOS NO PROGRAMADOS AL " + HO$, 9, 29)
                 Call REPLA(B00$, MKS$(0), 38, 4)
                 Call REPLA(B00$, MKS$(0), 42, 4)
                 Call REPLA(B00$, MKS$(0), 46, 2)
                 Call REPLA(B00$, " ", 48, 1)
                 '
                 Put #N9%, X2&, B00$
                 '
               End If
             End If
             '
           End If
         End If
       End If
129  Next i&
     '
     FORCANEC07.Check4.Value = 1
     '
     ' ORDENA POR FECHA DE NECESIDAD
     For KKI1& = 1 To X2&
1291    Get #N9%, KKI1&, B00$
        KKB1$ = B00$
        For KKI2& = KKI1& + 1 To X2&
          Get #N9%, KKI2&, B00$
          KKB2$ = B00$
          If CVI(Mid$(KKB2$, 7, 2)) < CVI(Mid$(KKB1$, 7, 2)) Then
            B00$ = KKB2$
            Put #N9%, KKI1&, B00$
            B00$ = KKB1$
            Put #N9%, KKI2&, B00$
            GoTo 1291
          End If
        Next KKI2&
     Next KKI1&

130  CICLO% = CICLO% + 1
     CICLOMA% = 9
     If InStr(UCase$(EMPREAC$), "YALTRE") > 0 Then CICLOMA% = 15
     '
     If CICLO% > CICLOMA% Then GoTo 150
     FORCANEC07.Label1 = TRIM$(Str$(CICLO%))
     FORCANEC07.Label1.Refresh
     '
' ******** CALCULANDO NECESIDADES NETAS SEGUN DISPONIBILIDAD
     '
     FORCANEC07.Check5.Value = 2
     Screen.MousePointer = 11
     For i& = X1& To X2&
       '
       CIIX& = i&
       Call TRACE(20, CIIX&, X2&)
       '
       Get #N9%, i&, B00$
       B1$ = Left$(B00$, 2)
       B3$ = Mid$(B00$, 3, 4)
       B4$ = Mid$(B00$, 7, 2)
       '
       CI% = CVI(B1$)
       If CI% > 0 Then
         If CI% <= NMS& Then
           If STATUIT%(CI%) >= 0 Then
             '
             ' ESTO SE AGREGO PARA FORZAR EL PROCESO POR NIVELES
             If NIVEGEN%(CI%) <> CICLO% - 1 Then
                GoTo 139
             End If
             ' **************************************************
             '
             If DISPACT(CI%) >= CVS(B3$) Then
               DISPACT(CI%) = DISPACT(CI%) - CVS(B3$)
               Call REPLA(B00$, B3$, 38, 4)
               Call REPLA(B00$, MKS$(0), 42, 4)
               Put #N9%, i&, B00$
               GoTo 139
             End If
             '
             CANEC = CVS(B3$) - DISPACT(CI%)
             FENEBRU% = CVI(B4$): CUANDO% = 1
             '
             For UX% = 1 To CACOL%
               If INIPER%(UX%) <= FENEBRU% Then CUANDO% = UX%
             Next UX%
             '
             If LOREC% = 208 Then
                 Get #N13%, CI%, PDS$
                 PDX$ = PDS$
               Else
                 Get #N13%, CI%, PDM$
                 PDX$ = PDM$
             End If
             '
             For UX% = CUANDO% To 1 Step -1
               PROANT = CVS(Mid$(PDX$, 4 * UX% - 3, 4))
               If PROANT > 0 Then GoTo 135
             Next UX%
             UX% = CUANDO%
135          If PROANT >= LOREPOS(CI%) Then UX% = CUANDO%
             '
             PROANT = CVS(Mid$(PDX$, 4 * UX% - 3, 4))
             ' INTERVENCION PARA CONSIDERAR LOTE INDIVISIBLE - 28/01/08
             CACORRE = PROANT + CANEC
             If LINDIVI(CI%) > 0 Then
                CACORRE = LINDIVI(CI%)
                While CACORRE < PROANT + CANEC: CACORRE = CACORRE + LINDIVI(CI%): Wend
             End If
             CANADI = CACORRE - (PROANT + CANEC)
             'Mid$(PDX$, 4 * UX% - 3, 4) = MKS$(PROANT + CANEC)
             Mid$(PDX$, 4 * UX% - 3, 4) = MKS$(CACORRE)
             '
             If LOREC% = 208 Then
                 PDS$ = PDX$
                 Put #N13%, CI%, PDS$
               Else
                 PDM$ = PDX$
                 Put #N13%, CI%, PDM$
             End If
             '
             FENENE% = INIPER%(UX%)
             '
             Call REPLA(B00$, MKS$(DISPACT(CI%)), 38, 4)
             DISPACT(CI%) = CANADI
             Call REPLA(B00$, MKS$(CANEC + CANADI), 42, 4)
             '
             Call REPLA(B00$, MKI$(FENENE%), 46, 2)
             Put #N9%, i&, B00$
             '
           End If
         End If
       End If
139  Next i&
     '
140  ' Call MENSAJE(14, "GENERANDO NECESIDADES SECUNDARIAS SEGUN EXPLOSION")

     X3& = X2&
     IAGRE& = 0
     For i& = X1& To X2&
       CIIX& = i&
       UKI& = i& - X1& + 1
       Call TRACE(20, UKI&, X2& - X1& + 1)
       '
       Get #N9%, i&, B00$
       '
       ' ESTO SE AGREGO PARA FORZAR EL PROCESO POR NIVELES
       KKII% = CVI(Left$(B00$, 2))
       If NIVEGEN%(KKII%) <> CICLO% - 1 Then
          IAGRE& = IAGRE& + 1
          If IAGRE& >= UDIM& Then
              UDIM& = UDIM& + 1024
              ReDim Preserve BAGRE$(UDIM&)
          End If
          BAGRE$(IAGRE&) = B00$
          B00$ = String$(48, 0)
          Put #N9%, i&, B00$
          GoTo 1236
       End If
       ' **************************************************
       
       Call REPLA(B00$, "C", 48, 1)
       Put #N9%, i&, B00$
       '
       B1$ = Left$(B00$, 2)
       B3$ = Mid$(B00$, 3, 4)
       B4$ = Mid$(B00$, 7, 2)
       B7$ = Mid$(B00$, 42, 4)
       B8$ = Mid$(B00$, 46, 2)
       '
       CI% = CVI(B1$)
       If STATUIT%(CI%) >= 0 Then
         CABAS = CVS(B7$)
         FeNec% = CVI(B8$): CUANEC% = 1
                         For UX% = 1 To CACOL%
                            If INIPER%(UX%) <= FeNec% Then CUANEC% = UX%
                         Next UX%
         If CABAS > 0 Then
1234        CIBAS% = CI%
           CII% = CIBAS%
           CODBAS$ = TRIM$(CODEXT$(CII%))
           ' CONSIDERA EL LEADTIME EN DIAS DEL MAESTRO DE ARTÍCULOS.
           Call LEEEXPLOMUL(CODBAS$, CABAS, 1, , 1)
           If CAIT% > 0 Then
             Call REPLA(B00$, "F", 48, 1)
             Put #N9%, i&, B00$
             For KJ% = 1 To CAIT%    ' J%
               CI% = CIJ%(KJ%)     '  CII1%(KJ%)
               If CI% < 1 Or CI% > NUMAS% Then
                  Call MENSERR(24, "Codigo Componente no Válido\en Estructura de '" + CODBAS$ + "'")
                  GoTo 1235
               End If
               '
               If STATUIT%(CI%) >= 0 Then
                 CANTU = CABAS * USOI(KJ%) 'CANT1(KJ%)
                 PLADISPO% = PDI%(KJ%)
                 If PLADISPO% < 1 Then PLADISPO% = PLADISTA%
                 CLX% = CUANEC% - PLADISPO%  'PAD1%(KJ%)
                 If CLX% < 1 Then CLX% = 1
                 If CANTU > 0 Then
                   '
                   For KX1& = 1 To IAGRE&
                     If CVI(Left$(BAGRE$(KX1&), 2)) = CIJ%(KJ%) Then
                       If CVI(Mid$(BAGRE$(KX1&), 7, 2)) = INIPER%(CLX%) Then
                         If TRIM$(Mid$(BAGRE$(KX1&), 9, 28)) = TRIM$("Res.p/" + CODBAS$) Then
                           CANTU22 = CANTU + CVS(Mid$(BAGRE$(KX1&), 3, 4))
                           B00$ = BAGRE$(KX1&)
                           Call REPLA(B00$, MKS$(CANTU22), 3, 4)
                           BAGRE$(KX1&) = B00$
                           GoTo 1235
                         End If
                       End If
                     End If
                   Next KX1&
                   '
                   B00$ = String$(48, 0)
                   Call REPLA(B00$, MKI$(CIJ%(KJ%)), 1, 2)
                   Call REPLA(B00$, MKS$(CANTU), 3, 4)
                   Call REPLA(B00$, MKI$(INIPER%(CLX%)), 7, 2)
                   Call REPLA(B00$, "Res.p/" + CODBAS$, 9, 28)
                   Call REPLA(B00$, " ", 48, 1)
                   IAGRE& = IAGRE& + 1
                   If IAGRE& >= UDIM& Then
                     UDIM& = UDIM& + 1024
                     ReDim Preserve BAGRE$(UDIM&)
                   End If
                   BAGRE$(IAGRE&) = B00$
                 End If
               End If
1235         Next KJ%
           End If
         End If
       End If
1236 Next i&
     '
1237 For JAGRE& = 1 To IAGRE&
       X3& = X3& + 1
       B00$ = BAGRE$(JAGRE&)
       Put #N9%, X3&, B00$
     Next JAGRE&
     '
     X1& = X2& + 1
     X2& = X3&: If X2& < X1& Then GoTo 150 ' terminar
     GoTo 130
 
150  FORCANEC07.Check5.Value = 1
     FORCANEC07.Check6.Value = 2
     Screen.MousePointer = 11
     '
     For i& = 1 To NMS&
       '
       CIIX& = i&
       Call TRACE(20, CIIX&, NMS&)
       If LOREC% = 208 Then
           Get #N13%, i&, PDS$
           PDX$ = PDS$
         Else
           Get #N13%, i&, PDM$
           PDX$ = PDM$
       End If
       '
       For J% = 1 To CACOL%
         CAFAB = CVS(Mid$(PDX$, 4 * J% - 3, 4))
         If CAFAB > 0 Then
           REFE$ = "O/F a emitir": Tipo% = 15
'           If DPI%(I&) = 0 Then REFE$ = "Tentativo de Compras": TIPO% = 14
'           UII% = I&
            UII% = i&
            If DPI%(i&) = 0 Or PROVHABI%(UII%) > 0 Then REFE$ = "Tentativo de Compras": Tipo% = 14
           Call COBERFA(INIPER%(J%), UII%, REFE$, Tipo%, CAFAB)
         End If
       Next J%
     Next i&
     '
     For i& = 1 To X3&
       CIIX& = i&
       Call TRACE(20, CIIX&, X3&)
       Get #N9%, i&, B00$
       B1$ = Left$(B00$, 2)
       B3$ = Mid$(B00$, 3, 4)
       B4$ = Mid$(B00$, 7, 2)
       B5$ = Mid$(B00$, 9, 29)
       B7$ = Mid$(B00$, 42, 4)
       B8$ = Mid$(B00$, 46, 2)
       '
       CI% = CVI(B1$)
       If CI% > 0 Then
         If STATUIT%(CI%) >= 0 Then
           CANEC = (-1) * CVS(B3$)
           FFX% = CVI(B4$)
           REFE$ = B5$
           Tipo% = 22
           If Left$(REFE$, 20) <> "PROGRAMA POR MINIMOS" Then
             If Left$(REFE$, 4) = "PEDI" Then Tipo% = 21
             Call COBERFA(FFX%, CI%, REFE$, Tipo%, CANEC)
           End If
         End If
       End If
     Next i&
     '
     FORCANEC07.Check6.Value = 1
     '
3900 Call CIERRARCH("*.*")
     If N4% > 0 Then Close #N4%: N4% = 0
     If N5% > 0 Then Close #N5%: N5% = 0
     If N7% > 0 Then Close #N7%: N7% = 0
     If N9% > 0 Then Close #N9%: N9% = 0
     If N13% > 0 Then Close #N13%: N13% = 0
     '
     KKLL& = 0
     For KKK% = 1 To 52
       If TRIM$(DEPER$(KKK%)) <> "" Then
         KKLL& = KKLL& + 1
         Call GRAREG("PERPROMA", DEPER$(KKK%) + Space$(16), KKLL&)
       End If
     Next KKK%
     Call SETULTREG("PERPROMA", KKLL&)
     Call CIERRARCH("PERPROMA")
     '
     For UI& = 0 To 32767
       PROVINT%(UI&) = 0
     Next UI&
     '
     For U& = 1 To ULTREG&("TAPROINT")
       X$ = REGLEIDO$("TAPROINT", U&)
       NP% = CVI(Left$(X$, 2))
       If NP% > 0 Then
         PROVINT%(NP%) = 1
       End If
     Next U&
     Call CIERRARCH("TAPROINT")
     '
     JC& = 0
     JF& = 0
     '
     FORCANEC07.Check7.Value = 2
     Screen.MousePointer = 11
     '
     FIN& = ULTREG&("ARTIPRO")
     For III& = 1 To FIN&
       Call TRACE(20, III&, FIN&)
       X$ = REGLEIDO$("ARTIPRO", III&)
       CI% = CVI(Left$(X$, 2))
       If CI% > 0 Then
         If CI% <= NUMAS% Then
           PROVEX$ = Mid$(X$, 3, 2)
           If ECOARCH$("PROVED1", PROVEX$) <> "" Then
             PROVARTI%(CI%) = CVI(PROVEX$)
           End If
         End If
       End If
     Next III&
     '
     URECO& = CACO&
     FIN& = URECO&
     For III& = 1 To FIN&
       '
       Call TRACE(20, III&, FIN&)
       X$ = REGLEIDO$("COBERFA", III&)
       TIMO% = CVI(Mid$(X$, 51, 2))
       CI% = CVI(Mid$(X$, 3, 2))
       FF% = CVI(Mid$(X$, 1, 2))
       PERIO$ = Mid$(X$, 5, 16)
       CANTX = CVS(Mid$(X$, 53, 4))
       ' FFI% = CVI(Mid$(X$, 61, 2))
       '
       If TIMO% = 15 Then
         PHB% = PROVHABI%(CI%) ' PROVEEDOR HABITUAL
         If PHB% > 0 Then
           If PROVINT%(PHB%) = 0 Then ' NO ES PROVEEDOR INTERNO
             TIMO% = 14
           End If
         End If
       End If
       '
       Call REPLA(X$, MKI$(TIMO%), 51, 2)
       Call REPLA(X$, MKS$(0), 61, 4)
       '
       Call GRAREG("COBERFA", X$, III&)
       '
       If DESCOMPO%(CI%) = 1 Then
         GoTo 3930
       End If
       '
       If TIMO% = 14 Then
         '
         MONEI% = MONECOSTO%(CI%)
         COSUDOL = COSUNI(CI%)
         COSUN = COSUDOL * TICAMONE(MONEI%)
         PHB% = PROVHABI%(CI%) ' PROVEEDOR HABITUAL
         If PHB% > 0 Then
           If PROVINT%(PHB%) <> 0 Then ' ES PROVEEDOR INTERNO
             COSUN = 0
           End If
         End If
         '
         'Se graba en PLACOMP
         Z$ = REGBLAN$("PLACOMP")
         Call REPLA(Z$, MKI$(CI%), 1, 2)
         Call REPLA(Z$, MKI$(FF%), 3, 2)
         Call REPLA(Z$, PERIO$, 5, 16)
         Call REPLA(Z$, MKS$(CANTX), 21, 4)
         Call REPLA(Z$, MKS$(0), 25, 4) ' tot.progra.por item
         Call REPLA(Z$, MKS$(0), 29, 4) ' filler
         Call REPLA(Z$, MKI$(PHB%), 33, 2)
         Call REPLA(Z$, MKS$(COSUN), 35, 4)
         Call REPLA(Z$, MKD$(COSUN * CANTX), 39, 8)
         '
         JC& = JC& + 1
         Call GRAREG("PLACOMP", Z$, JC&)
         '
         FF1$ = REGBLAN$("OCOMEMI")
         Call REPLA(FF1$, MKI$(CI%), 1, 2)
         Call REPLA(FF1$, DESCRIT$(CI%), 3, 26)
         Call REPLA(FF1$, CSTRING$(MKS$(CANTX), 4, 10, 1, 2), 29, 10)
         FEX = FF%
         Call REPLA(FF1$, FECHATEX$(FEX), 41, 8)
         Call REPLA(FF1$, MKD$(CDbl(CANTX)), 49, 8)
         Call REPLA(FF1$, MKI$(FF%), 57, 2)
         Call REPLA(FF1$, MKI$(CI%), 59, 2)
         Call GRAREG("OCOMEMI", FF1$, JC&)
         '
       ElseIf TIMO% = 15 Then
         '
         Z$ = REGBLAN$("PLAFABR")
         Call REPLA(Z$, MKI$(CI%), 1, 2)
         Call REPLA(Z$, MKI$(FF%), 3, 2)
         Call REPLA(Z$, PERIO$, 5, 16)
         Call REPLA(Z$, MKS$(CANTX), 21, 4)
         Call REPLA(Z$, MKS$(0), 29, 4)  ' FILLER
         JF& = JF& + 1
         Call GRAREG("PLAFABR", Z$, JF&)
         '
         FF1$ = REGBLAN$("ORFAEMI")
         Call REPLA(FF1$, MKI$(CI%), 1, 2)
         Call REPLA(FF1$, DESCRIT$(CI%), 3, 26)
         Call REPLA(FF1$, CSTRING$(MKS$(CANTX), 4, 10, 1, 2), 29, 10)
         FEX = FF%
         Call REPLA(FF1$, FECHATEX$(FEX), 41, 8)
         Call REPLA(FF1$, MKD$(CDbl(CANTX)), 49, 8)
         Call REPLA(FF1$, MKI$(FF%), 57, 2)
         Call REPLA(FF1$, MKI$(CI%), 59, 2)
         Call GRAREG("ORFAEMI", FF1$, JF&)
         '
       End If
       '
3930 Next III&
     FORCANEC07.Check7.Value = 1
     '
     Screen.MousePointer = 1
     URECO& = CACO&
     Call SETULTREG("COBERFA", URECO&)
     Call CIERRARCH("COBERFA")
     Call SETULTREG("ORFAEMI", JF&)
     Call CIERRARCH("ORFAEMI")
     Call SETULTREG("OCOMEMI", JC&)
     Call CIERRARCH("OCOMEMI")
     Call SETULTREG("PLACOMP", JC&)
     Call CIERRARCH("PLACOMP")
     Call SETULTREG("PLAFABR", JF&)
     Call CIERRARCH("PLAFABR")
     Call CIERRARCH("PERPROMA")
     '
     FECHOR$ = HOS$ + "  " + Left$(Time$, 5)
     Call GRACONTROL("CANECOM", "FECHOCAL", FECHOR$)
     If UCase$(Left$(App.EXEName, 5)) = "FLEXO" Then
         PROMAT00.Label7.Caption = Left$(FECHOR$, 8): PROMAT00.Label7.Refresh
         PROMAT00.Label95.Caption = Mid$(FECHOR$, 11, 5): PROMAT00.Label95.Refresh
       Else
         PROMAT07.Label7.Caption = Left$(FECHOR$, 8): PROMAT07.Label7.Refresh
         PROMAT07.Label95.Caption = Mid$(FECHOR$, 11, 5): PROMAT07.Label95.Refresh
     End If
     Call CIERRARCH("*.*")
     Call BAJALDISCO
     '
     Call FILESORT("COBERFA", "", 6, 6, "ASC")
     Call FILESORT("COBERFA", "", 1, 2, "ASC")
     Call FILESORT("PLACOMP", "", 1, 3, "ASC")
     Call FILESORT("PLAFABR", "", 1, 3, "ASC")
     Call FILESORT("ORFAEMI", "", 4, 5, "ASC")
     Call FILESORT("OCOMEMI", "", 4, 5, "ASC")
     '
     CIAX% = 0
     DEPOSTAKK% = Val(TRIM$(Control$("PROMATER", "DEPOSTOK")))
     '
     Call BORRAREGISTROS("OCOEMI", "COD_INTE <> 0", REG&)
     Dim Ocoemi As ADODB.Recordset
     Set Ocoemi = New ADODB.Recordset
     SQLX1$ = "SELECT * FROM OCOEMI WHERE COD_INTE = 0"
25   On Error Resume Next
     Ocoemi.Open SQLX1$, FLEXCONN, adOpenDynamic, adLockPessimistic
     If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
        On Error GoTo 0
        Call CapturaErrorOpen
         GoTo 25
     End If
     On Error GoTo 0

     '
     For U& = 1 To ULTREG&("PLACOMP")
       XX$ = REGLEIDO$("PLACOMP", U&)
       CII% = CVI(Left$(XX$, 2))
       If CII% <> CIAX% Then
         CIAX% = CII%
         Stakk = STODEPOS(CII%, DEPOSTAKK%): If Stakk < 0 Then Stakk = 0
         Call REPLA(XX$, MKS$(Stakk), 29, 4)
         Call REPLA(XX$, MKS$(STOMIN(CII%)), 55, 4)
         Call REPLA(XX$, MKS$(LOTESTAN(CII%)), 59, 4)
       End If
       '
       LT% = LEAD_TIME%(CII%)
       FeNec = CVI(Mid$(XX$, 3, 2))
       FANT% = DIANTE(FeNec, LT%)
       'Si la fecha de solicitud ya ha pasado, entonces el sistema pone la del dia del calculo
       If FANT% < HOY(HO$) Then
          FANT% = HOY(HO$)
       End If
       Call REPLA(XX$, MKI$(FANT%), 63, 2)
       '
       'Se graba en OCOEMI
       CI% = CVI(Mid$(XX$, 1, 2))
       FF% = CVI(Mid$(XX$, 3, 2))
       PERIO$ = Mid$(XX$, 5, 16)
       CANTX = CVS(Mid$(XX$, 21, 4))
       PHB% = CVI(Mid$(XX$, 33, 2))
       COSUN = CVS(Mid(XX$, 35, 4))
       'Se Graba OCOEMI
       With Ocoemi
           .AddNew
           !COD_INTE = CI%
           !CANTIDAD = CANTX
           !Fentre_Sol = FECHATEX$(FF%)
           !Npro_Ocom = PHB%
           !Raso_prov = Left(rasocli$(PHB% * -1), 48)
           !Uni_Mas = Unimed$(CI%)
           !FechRecep = FECHATEX$(FANT%)
           !COSUNI = COSUN
           !PERIODO = PERIO$
           !CodiEmpr = CodiEmp%
           !StokAct = Stakk
           !STOMIN = STOMIN(CI%)
           !LOREPOS = LOTESTAN(CI%)
           .Update
         End With
       '
'AQUI PONER FECHA LIMITE EMISION O/COMPRA
'FECHA DE NECESIDAD EN 3,2
'EN 63,2 PONER MKI$(DIANTE(FECHA DE NECESIDAD, LEAD TIME))
       Call GRAREG("PLACOMP", XX$, U&)
     Next U&
     Set Ocoemi = Nothing
     Call CIERRARCH("PLACOMP")
     Call CARPROSEL
     '
     CIAX% = 0
     'ACA LE GRABA EL STOCKACTUAL POR DEPOSITO - MINIMO - LOTESTAN.
     For U& = 1 To ULTREG&("PLAFABR")
       XX$ = REGLEIDO$("PLAFABR", U&)
       CII% = CVI(Left$(XX$, 2))
       If CII% <> CIAX% Then
         CIAX% = CII%
         Stakk = STODEPOS(CII%, DEPOSTAKK%): If Stakk < 0 Then Stakk = 0
         Call REPLA(XX$, MKS$(Stakk), 29, 4)
         Call REPLA(XX$, MKS$(STOMIN(CII%)), 55, 4)
         Call REPLA(XX$, MKS$(LOTESTAN(CII%)), 59, 4)
         Call GRAREG("PLAFABR", XX$, U&)
       End If
     Next U&
     '
     Call CIERRARCH("PLAFABR")
     '
999  Call CIERRARCH("*.*")

     '
     'SE GUARDA EN PLACOMP
     If InStr(UCase$(EMPREAC$), "GABAL") > 0 Or InStr(UCase$(EMPREAC$), "LAFE") > 0 Then

         'Inicio Autonomia en Calculo de MRP
        Dim ACONSUMO()
        NMSI% = NUMAS%
        ReDim ACONSUMO(NMSI%)
        
        HOII% = HOY(HO$)
        FEDESDE$ = FETEXCOR1$(FECHANUM("01/01/10"))
        FEHASTA$ = FETEXCOR1$(HOII%)
        DIAS2010 = DIENTRE(FECHANUM("31/12/09"), HOII)
        '
        If DIENTRE(FECHANUM("01/01/10"), HOII%) > 182 Then
           FINI% = DIANTE(HOII, 183)
           FEDESDE$ = FETEXCOR1$(FINI%)
           DIAS2010 = 182
        End If
        '
        SQLX$ = "SELECT COD_INTE, SUM(CANTSALID) AS SALIDA, SUM(CANTINGRE) AS INGRESO FROM MOVISTO WHERE "
        SQLX$ = SQLX$ + " (FECHMOV >= '" & FEDESDE$ & "' AND FECHMOV <= '" & FEHASTA$ & "') "
        SQLX$ = SQLX$ + " AND (CODIMOVI ='RT' OR CODIMOVI ='CF' "
        SQLX$ = SQLX$ + " OR CODIMOVI ='DV' "
        SQLX$ = SQLX$ + " OR CODIMOVI ='RC') GROUP BY COD_INTE "
        '
        Dim articulos As ADODB.Recordset
        Set articulos = New ADODB.Recordset
        Set articulos = FLEXCONN.Execute(FILTSQL$(SQLX$))
        '
        'SE GUARDA EL CONSUMO + LA CUENTA
        With articulos
           Do While Not (.EOF)
              CI% = XVALO(!COD_INTE)
              ACONSUMO(CI%) = XVALO(!SALIDA) - XVALO(!INGRESO)
           .MoveNext
           Loop
        End With
        articulos.Close
        '
        If DIAS2010 < 182 Then
          For CI% = 1 To NUMAS%
             Condi$ = "CODINT = " & CI%
             CON# = XVALO(VALOBADA("MASTER", "CONSUPROM", Condi$))
             If DIAS2010 >= 182 Then CON# = 0
             CUENTA# = CON# / 30 * (182 - DIAS2010)
             ACONSUMO(CI%) = ACONSUMO(CI%) + CUENTA#
          Next CI%
        End If

         For U& = 1 To ULTREG&("PLACOMP")
             REGISTRO$ = REGLEIDO("PLACOMP", U&)
             CI% = CVI(Mid(REGISTRO$, 1, 2))
             CONSUDIARIO# = ACONSUMO(CI%) / 182
             STOACT# = CVS(Mid(REGISTRO$, 29, 4))
             On Error Resume Next
             AUTONOMIA% = STOACT# / CONSUDIARIO#     'STOCKACT(CI%) / CONSUDIARIO#
             If AUTONOMIA% < 0 Then AUTONOMIA% = 0
             If Err > 0 Then
                AUTONOMIA% = 999
             End If
             On Error GoTo 0
             Call REPLA(REGISTRO$, MKI$(AUTONOMIA%), 63, 2)
             Call GRAREG("PLACOMP", REGISTRO$, U&)
         Next U&
    
    '     SE GUARDA EN PLAFABR
         For U& = 1 To ULTREG&("PLAFABR")
             REGISTRO$ = REGLEIDO("PLAFABR", U&)
             CI% = CVI(Mid(REGISTRO$, 1, 2))
             CONSUDIARIO# = ACONSUMO(CI%) / 182
             STOACT# = CVS(Mid(REGISTRO$, 29, 4))
             On Error Resume Next
             AUTONOMIA% = STOACT# / CONSUDIARIO#     'STOCKACT(CI%) / CONSUDIARIO#
             If AUTONOMIA% < 0 Then AUTONOMIA% = 0
             If Err > 0 Then
               AUTONOMIA% = 999
             End If
             On Error GoTo 0
             Call REPLA(REGISTRO$, MKI$(AUTONOMIA%), 63, 2)
             Call GRAREG("PLAFABR", REGISTRO$, U&)
         Next U&
     '
     End If
     Call GRABA_COBERFA_a_SQL
     
     Call BAJALDISCO
     finix = Now
     Call COMUNI("Re-Cálculo Completo.")
     '
9999 FORCANEC07.Hide
     '
End Sub

Sub CANECOM_vie()         ' ESTE ES EL VIGENTE A PARTIR DE 17 NOV 2007
                      ' SE INTRODUJO FORZAR EL PROCESAMIENTO EN ORDEN DE NIVELES
                      ' DE EXPLOSIÓN POR EL PROBLEMA DE ACM DE AGOLPAR TODA LA
                      ' PROGRAMACION EN LA PRIMERA SEMANA.
                      '
     iniciox = Now
     Dim BAGRE() As String
     Dim CLICONPEDID%(32767)    'Este Vector registra los Clientes con Pedidos Clientes.
     Dim HAYPLAFAB%(32767)
     '
     EPAC$ = TRIM$(UCase$(EMPREAC$))
     UDIM& = 2048
     ReDim Preserve BAGRE$(UDIM&)
     LU$ = LUDAT$
     HOO% = HOY(HO$)
     CACO& = 0
     '
     SUPRICO% = 0: If FORCANEC07.Check8.Value = 1 Then SUPRICO% = 1
     SUPRIFA% = 0: If FORCANEC07.Check9.Value = 1 Then SUPRIFA% = 1
     SUPRIPE% = 0: If FORCANEC07.Check10.Value = 1 Then SUPRIPE% = 1
     SUPRILECON% = 0: If FORCANEC07.Check11.Value = 1 Then SUPRILECON% = 1
     ' verifica compras pendientes de aprobación
     Call ABRECOMPRAS
     If SUPRICO% > 0 Then GoTo 90
'     '
     SQLX$ = " SELECT Nume_Ocom, Refe_Ocom,ToBru_Ocom FROM OCOMENCA"
     SQLX$ = SQLX$ + " WHERE Stat_Ocom < 0 "
     '
     Dim VEROCTEMP As ADODB.Recordset
     Set VEROCTEMP = New ADODB.Recordset
     VEROCTEMP.CursorType = adOpenDynamic
     VEROCTEMP.LockType = adLockReadOnly
     If InStr(EPAC$, "GABAL") > 0 Then
          VEROCTEMP.Open SQLX$, FLEXCONN, , , adCmdText
        Else
          VEROCTEMP.Open FILTSQL$(SQLX$), FLEXCONN, , , adCmdText
     End If
     '
     With VEROCTEMP
        On Error Resume Next
        .MoveLast
        If Err Then GoTo 90    ' no hay pendientes de aprobación
        On Error GoTo 0
        TTXX$ = "Existen Ordenes de Compra Pendientes de Aprobación.\   \Estas Compras NO Serán Tenidas en Cuenta\en el Proceso de Cálculo MRP.\Puede Optar entre Cancelar Cálculo o Continuar\\Cancelar Cálculo\Continuar"
        If COMALTER%(TTXX$) <> 2 Then GoTo 9999
     End With
     '
90   On Error GoTo 0
     Screen.MousePointer = 11
     '
100  ' ************************* ABRIENDO ARCHIVOS MAESTROS
     '
     FORCANEC07.Check1.Value = 2
     Screen.MousePointer = 11
     NMS& = NUMAS%
     '               NVI% = NUINV%
     '               N1% = FILENBR%("MASTER")
     '               N1% = FILEOPEN%(LUDAT$ + "MASTER.DAT", 0, 128)
     '               N9% = FILENBR%("SALSTOCK")
     '               N9% = FILEOPEN%(LUDAT$ + "SALSTOCK.DAT", 0, 40)
     '
     '               GESUBOR% = 0
     '               If CONTROL$("", "GESUBOR") = "SI" Then
     '                 GESUBOR% = 1
     '               End If
     '
     'GoSub 5300                           ' estrunue - estrudos
     FORCANEC07.Check1.Value = 1
     '
     'If NESTRU& < 1 Then
     '   Call MENSERR(24, "Proceso Imposible.\ \Faltan Definir Fórmulas y\Estructuras de Producto.")
     '   Call FINAL("")
     'End If
     '
     PLADISTA% = Val(FORCANEC07.Text1.TEXT)
     If PLADISTA% > 6 Then PLADISTA% = 6
     '
'    **** DEPURA PROGRAMA DE ENTREGAS POR ENTREGAS REALIZADAS
     '
     APROGRA$ = "PROGENT"
     If MODOPLAN% = 1 Then APROGRA$ = "PROGFAB"
     '
     ' ESTO QUE SIGUE ES PARA NO ALTERAR EL PROGRAMA ORIGINAL DE ENTREGAS
     ' SE CONFECCIONA UNA IMAGEN EN PROGCALC Y SE TRABAJA SOBRE ESTA
     Call COPYSTRU(APROGRA$, "PROGCALC")
     JJL1& = 0
     For i& = 1 To ULTREG&(APROGRA$)
       ZZ$ = REGLEIDO$(APROGRA$, i&)
       JJL1& = JJL1& + 1
       Call GRAREG("PROGCALC", ZZ$, JJL1&)
     Next i&
     Call SETULTREG("PROGCALC", JJL1&)
     Call CIERRARCH("PROGCALC")
     Call CIERRARCH(APROGRA$)
     APROGRA$ = "PROGCALC"
     ' FIN CONFECCION DE IMAGEN PROGENT.
     '
     UREPRO& = ULTREG&(APROGRA$)
     JJ& = 0
     For i& = 1 To UREPRO&
       Call TRACE(20, i&, FIN&)
       ZZ$ = REGLEIDO$(APROGRA$, i&)
       If CVS(Mid$(ZZ$, 7, 4)) > 0.05 Then
         Call REPLA(ZZ$, MKS$(0), 11, 4)
         Call REPLA(ZZ$, MKS$(0), 19, 4)
         JJ& = JJ& + 1
         Call GRAREG(APROGRA$, ZZ$, JJ&)
       End If
       If MODOPLAN% = 1 Then
         ' marca items programados para no considerar stock como disponibilidad
         CIXI% = CVI(Mid$(ZZ$, 1, 2))
         If CIXI% > 0 Then
           If CIXI% <= NUMAS% Then
             CAPROBRU = CVS(Mid$(ZZ$, 7, 4))
             If CAPROBRU > 0.05 Then
               HAYPLAFAB%(CIXI%) = 1
             End If
           End If
         End If
         '
       End If
     Next i&
     Call SETULTREG(APROGRA$, JJ&)
     Call CIERRARCH(APROGRA$)
     If JJ& < 1 Then
        On Error Resume Next
        Kill LUDAT$ + APROGRA$ + ".DAT"
        On Error GoTo 0
     End If
     '
' LEER ENTREGAS DEL MES PARA NETEAR PROGRAMA DE ENTREGAS
     If MODOPLAN% = 1 Then GoTo 65   ' SI POR FABRICACION NO NETEA
     '
     CMA% = COMESACT%
     Call ABRESTOCKS
     CMAX = FETEXCOR1$(CMA%)
     'SQLX$ = " SELECT * FROM MOVISTO "
     SQLX$ = " SELECT * FROM MOVISTO WITH(NOLOCK)"
     SQLX$ = SQLX$ + "WHERE FechMov >= '" & CMAX & "' "
     SQLX$ = SQLX$ + "AND (CodiMovi = 'RT' Or CodiMovi = 'DV')"
     '
     Dim MoviTemp As ADODB.Recordset
     Set MoviTemp = New ADODB.Recordset
     MoviTemp.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockReadOnly, adCmdText
     With MoviTemp
         Do While Not .EOF
           CIXI% = !COD_INTE
           If CIXI% > 0 Then
             If CIXI% <= NUMAS% Then
             AAA = CDbl(!FechMov)
               CAENTREG = !CANTSALID - !CantIngre
               '
               RFF$ = RECFILT$(APROGRA$, 1, MKI$(CIXI%))
               For KK1& = 1 To Len(RFF$) Step 4
                 REPRO& = CVS(Mid$(RFF$, KK1&, 4))
                 ZZ$ = REGLEIDO$(APROGRA$, REPRO&)
                 ACUENTRE = CVS(Mid$(ZZ$, 11, 4))
                 ENTREMA = CVS(Mid$(ZZ$, 7, 4)) - ACUENTRE
                 If CAENTREG <= ENTREMA Then
                    Call REPLA(ZZ$, MKS$(CAENTREG + ACUENTRE), 11, 4)
                    CAENTREG = 0
                  Else
                    Call REPLA(ZZ$, MKS$(ENTREMA + ACUENTRE), 11, 4)
                    CAENTREG = CAENTREG - ENTREMA
                 End If
                 Call GRAREG(APROGRA$, ZZ$, REPRO&)
                 If CAENTREG < 0.05 Then GoTo 102
               Next KK1&
             End If
           End If
           '
102      .MoveNext
         Loop
     End With
     '
'    **** AGREGA PEDIDOS PENDIENTES DE ENTREGA
'    **** DESCARTANDO L0S CORRESPONDIENTES A REPARACIONES
      '
65    If MODOPLAN% = 1 Then
        For CIXI% = 1 To NUMAS%: ACTICOD%(CIXI%) = 1: Next CIXI%
        GoTo 75
      End If
      '
      FIN& = ULTREG&(APROGRA$)
      RFF$ = ""
      For U& = 1 To FIN&
        X$ = REGLEIDO$(APROGRA$, U&)
        RFF$ = RFF$ + Left$(X$, 2)
      Next U&
      '
     Call ABREPEDCLI
     Condi$ = "(Status >= 0 AND Status < 8) "      ' para que no este anulada ni cumplida ni pendiente de aprobacion
     Condi$ = Condi$ + "AND CantEnt < (CanPed - 0.005) "
     Condi$ = Condi$ + " AND VenRep = 'V' "
     '
     SQLX$ = " SELECT * FROM PEDDETA WITH(NOLOCK)"
     'SQLX$ = SQLX$ + "WHERE CodiInt = " & CI% & " "
     SQLX$ = SQLX$ + "WHERE " & Condi$
     SQLX$ = SQLX$ + " ORDER BY FeSolEnt DESC "
     '
     FINPE& = CantRegistros("PEDDETA", Condi$)
     Dim PEDTEMP As ADODB.Recordset
     If InStr(EPAC$, "GABAL") > 0 Then
          Set PEDTEMP = FLEXCONN.Execute(SQLX$)
        Else
          Set PEDTEMP = FLEXCONN.Execute(FILTSQL$(SQLX$))
     End If
     '\\\OT-05-0631-WAR-20/09/05///
     INICLICON% = 0
     If UCase$(Control$("CANECOM", "FILPECLI")) <> "SI" Then INICLICON% = 1
     ' SI NO FILTRA POR CLIENTES CON PEDIDOS, ENTONCES TOMA TODOS
     For JJ1% = 1 To 32766 'PONGO EN 0 EL VECTOR
       CLICONPEDID%(JJ1%) = INICLICON%
     Next JJ1%
     FILPEDIPEN% = 0
     If INICLICON% = 1 Then
        For CIXI% = 1 To NUMAS%: ACTICOD%(CIXI%) = 1: Next CIXI%
        GoTo 50
     End If
     '
     ' CARGA UN VECTOR CON LOS CLIENTES CON PEDIDOS PENDIENTES
     FILPEDIPEN% = 1: KU& = 0
     For CIXI% = 1 To NUMAS%: ACTICOD%(CIXI%) = 0: Next CIXI%
     With PEDTEMP
       On Error Resume Next
       .MoveFirst
       If Err < 1 Then
         Do While Not .EOF
          KU& = KU& + 1
          Call TRACE(20, KU&, FINPE&)
          NCLIE1 = !NROCLIE
          If Val(NCLIE1) > 0 Then
            If Val(NCLIE1) <= 32767 Then
              CLICONPEDID%(NCLIE1) = 1 'MARCO LOS CLIENTES CON PEDIDOS
            End If
          End If
          .MoveNext
         Loop
       End If
     End With
     '
     For JL& = 1 To 32767
      If CLICONPEDID%(JL&) = 1 Then
        NCLIE2$ = FORMATNUM$(JL&, "F6.0")
        RASOCLIE2$ = rasocli$(JL&)
        Z$ = Space(128)
        Call REPLA(Z$, NCLIE2$, 1, 6)
        Call REPLA(Z$, RASOCLIE2$, 15, 64)
        With CLICOPE07
          .LICOPEN.AddItem Z$
        End With
      End If
     Next JL&
     Screen.MousePointer = 1
     If CLICOPE07.LICOPEN.ListCount > 0 Then
       CLICOPE07.Show 1
       Call FRESHALL
       If CLICOPE07.Label2.Caption <> "OK" Then
         Call COMUNI("Cancelado !!!")
         GoTo 9999
       End If
       '
       ' ACA VUELVE A PONER EN '0' EL VECTOR DE CLIENTES
       For JL& = 1 To 32767
         CLICONPEDID%(JL&) = 0
       Next JL&
       '
       Call COMUNI("El Cálculo MRP-II sólo Considerará Aquellos Items de Stock\que Sean Componentes de Alguno de los Productos de los Pedidos\Pendientes de Entrega de los Clientes Seleccionados.\   \Todos los Restantes Artículos se Suprimen, Independientemente de su\Situación de Stocks Actual o Mínimo y Ordenes de Compra o Fabricación en Curso.")
       ' ACA RE-CARGA EL VECTOR PONIENDO EN '1' LOS CLIENTES SELECCIONADOS
       CLISELE$ = ""
       For JL& = 1 To CLICOPE07.LICOPEN.ListCount
         If CLICOPE07.LICOPEN.Selected(JL& - 1) = True Then
           ICLIE1& = Val(Left$(CLICOPE07.LICOPEN.List(JL& - 1), 6))
           If ICLIE1& > 0 Then
             If ICLIE1& <= 32767 Then
               CLICONPEDID%(ICLIE1&) = 1
               If CLISELE$ <> "" Then
                  CLISELE$ = CLISELE$ + " + "
               End If
               NCXX% = ICLIE1&
               CLISELE$ = CLISELE$ + TRIM$(ABRECLI$(NCXX%)) + "(" + TRIM$(Str$(NCXX%)) + ")"
             End If
           End If
         End If
       Next JL&
     End If
     Call GRACONTROL("CANECOM", "CLISELE", CLISELE$)
     '
50   Screen.MousePointer = 11
     '\\\OT-05-0631-WAR-FIN///
     KU& = 0
     With PEDTEMP
       On Error Resume Next
       .MoveFirst
       If Err < 1 Then
         Do While Not .EOF
           '
           KU& = KU& + 1
           Call TRACE(20, KU&, FINPE&)
           ' AQUI DESCARTA LOS PEDIDOS DE CLIENTES NO MARCADOS
           NCLIE1 = !NROCLIE
           If Val(NCLIE1) > 0 Then
             If Val(NCLIE1) <= 32767 Then
               If CLICONPEDID%(NCLIE1) <> 1 Then
                 GoTo 103
               End If
             End If
           End If
           '
           CIXI% = !CODIINT
           Call EXPLOMUL(CIXI%, 1)
           For JJ1% = 1 To CONTAMUL%
             CIX2% = CIK%(JJ1%)
             ACTICOD%(CIX2%) = 2
           Next JJ1%
           '
           NCX1% = NCLIE1
           CAPEN = !CanPed - !CANTENT
           If SUPRIPE% = 1 Then CAPEN = 0
           If CAPEN > 0 Then
             FENTRESOL% = CVINT%(!FeSolEnt)
             For KK1& = 1 To Len(RFF$) Step 2    ' - 3 To 1 Step -4
               If CVI(Mid$(RFF$, KK1, 2)) = CIXI% Then
                 REPRO& = (KK1& + 1) / 2 ' CVS(Mid$(RFF$, KK1&, 4))
                 ZZ$ = REGLEIDO$(APROGRA$, REPRO&)
                 FEPROG% = CVI(Mid$(ZZ$, 3, 2))
                 If COMIENFE%(FEPROG%) <= COMIENFE%(FENTRESOL%) Then
                   CAPEPROMAX = CVS(Mid$(ZZ$, 7, 4)) - CVS(Mid$(ZZ$, 11, 4)) - CVS(Mid$(ZZ$, 19, 4))
                   If CAPEPROMAX < 0 Then CAPEPROMAX = 0
                   '         PROGRAMA BRUTO - ENTREGAS - PEDIDOS PROGRAMADOS
                   If CAPEPROMAX >= CAPEN Then
                       CAPEDID = CVS(Mid$(ZZ$, 19, 4)) + CAPEN
                       Call REPLA(ZZ$, MKS$(CAPEDID), 19, 4)
                       Call GRAREG(APROGRA$, ZZ$, REPRO&)
                       CAPEN = 0
                       GoTo 103
                     Else
                       CAPEDID = CVS(Mid$(ZZ$, 19, 4)) + CAPEPROMAX
                       Call REPLA(ZZ$, MKS$(CAPEDID), 19, 4)
                       Call GRAREG(APROGRA$, ZZ$, REPRO&)
                       CAPEN = CAPEN - CAPEPROMAX
                   End If
                 End If
               End If
             Next KK1&
             '
             If CAPEN > 0.05 Then
               FENTRESO% = FENTRESOL%
               If FENTRESO% < HOO% Then FENTRESO% = HOO%
               ZZ$ = REGBLAN$(APROGRA$)
               Call REPLA(ZZ$, MKI$(CIXI%), 1, 2)
               Call REPLA(ZZ$, MKI$(FENTRESO%), 3, 2)
               Call REPLA(ZZ$, MKI$(NCX1%), 5, 2)
               Call REPLA(ZZ$, MKS$(CAPEN), 19, 4)
               Call REGAPP(APROGRA$, ZZ$)
               RFF$ = RFF$ + MKI$(CIXI%)
                 'Call CIERRARCH(APROGRA$)
             End If
           End If
           '
103      .MoveNext
         Loop
         '
       End If
     End With
75   Call CIERRARCH(APROGRA$)
     '
' ************************* pone ACTICOD% = 0 para los items de baja
     Call ABREMASTER
     SQLX$ = " SELECT Codint, Status FROM MASTER WITH(NOLOCK) "
     SQLX$ = SQLX$ + "WHERE Status < 0 "          ' SELECCIONA LOS DE BAJA
     Dim IDEBAJA As ADODB.Recordset
     Set IDEBAJA = FLEXCONN.Execute(FILTSQL$(SQLX$))
     '
     NMS1% = NUMAS%
     With IDEBAJA
       Do While Not .EOF
         CIXI% = !CODINT
         If CIXI% > 0 Then
            If CIXI% <= NMS1% Then
               ACTICOD%(CIXI%) = 0
            End If
         End If
       .MoveNext
       Loop
     End With
     IDEBAJA.Close
     Set IDEBAJA = Nothing
     '
' ************** lee existencias, stock minimo y lote reposicion *******************
     FORCANEC07.Check2.Value = 2
     Screen.MousePointer = 11
     '
     ' determina items fanricados con dpi%()=1
     FIN& = NUMAS%
     For i& = 1 To FIN&
       Call TRACE(20, i&, FIN&)
       CI% = i&
       Call AREXPLO(CI%, 1, 3)
       If CAIT% > 0 Then DPI%(CI%) = 1
105  Next i&
     '
     Call CALNIVEGEN        ' DETERMINA EL NIVEL DE EXPLOSION GENERAL DE CADA ITEM
     '
     HAYNE% = 0
     MODO% = 0
     MODOCAX$ = FORCANEC07.Option3.Caption
     If FORCANEC07.Option1.Value = True Then
         MODO% = 1
         MODOCAX$ = FORCANEC07.Option1.Caption
       ElseIf FORCANEC07.Option2.Value = True Then
         MODO% = 2
         MODOCAX$ = FORCANEC07.Option2.Caption
     End If
     Call GRACONTROL("CANECOM", "MODOCAL", MODOCAX$)
     '
     FFXY = HOY(HO$)
     SMXX% = SEMANU%(FFXY, AA, BB)
     COSEMACT% = AA
     '
     QLN& = 0
     Call CIERRARCH("COBERFA")
     Call CIERRARCH("ORFAEMI")
     Call CIERRARCH("OCOMEMI")
     '
     For CI% = 1 To NMS&
       CIIX& = CI%
       Call TRACE(20, CIIX&, NMS&)
       SALSTO = 0
       If STATUIT%(CI%) >= 0 Then
         SALSTO = STOCKACT(CI%)
         '
         If MODO% = 1 Then SALSTO = 0
         If MODO% = 2 Then If SALSTO < 0 Then SALSTO = 0

         If SALSTO < (-0.05) Then
            HAYNE% = HAYNE% + 1
            CI1% = CI%
            TTXX$ = REGBLAN$("LINEGAT")
            Call REPLA(TTXX$, MKI$(CI%), 1, 2)
            Call REPLA(TTXX$, DESCRIT$(CI%), 3, 30)
            Call REPLA(TTXX$, CSTRING$(MKS$(SALSTO), 4, 10, 1, 2), 33, 10)
            Call REPLA(TTXX$, Unimed$(CI%), 45, 4)
            '
            QLN& = QLN& + 1
            Call GRAREG("LINEGAT", TTXX$, QLN&)
            '
         End If
         '
         STMIN = STOMIN(CI%): If MODO% = 1 Then STMIN = 0
         LINDIVI(CI%) = LINDIVIS(CI%): If LINDIVI(CI%) < 1 Then LINDIVI(CI%) = 1
         REPOS = LOTESTAN(CI%)      ' SUJETO A SUPRIME LOTE ECONOMICO
         DP% = DPI%(CI%): If DP% < 0 Or DP% > 3 Then DP% = 0

         If REPOS < 0.5 Then REPOS = LOSTN(DP%)
         If REPOS < 0.5 Then REPOS = REPOMIN
         If SUPRILECON% > 0 Then REPOS = 0  ' SUPRIME LOTE ECONOMICO
         '
         DISPACT(CI%) = SALSTO - STMIN
         If HAYPLAFAB%(CI%) > 0 Then DISPACT(CI%) = 0
         LOREPOS(CI%) = REPOS
         '
         If HAYPLAFAB%(CI%) < 1 Then
            If Abs(SALSTO) > 0.05 Then
              Call COBERFA(COSEMACT%, CI%, "Stock al " + HO$, 10, SALSTO)
            End If
            '
            If Abs(STMIN) > 0.05 Then
              Call COBERFA(COSEMACT%, CI%, "Stock Minimo del Item", 20, (-1) * STMIN)
            End If
         End If
       End If
     Next CI%
     '
     Call SETULTREG("LINEGAT", QLN&)
     Call CIERRARCH("LINEGAT")
     '
     If HAYNE% > 0 Then
        Call FILESORT("LINEGAT", "", 1, 1, "ASC")
        A% = COMALTER%("ATENCION: Se Detectaron Existencias Negativas\          en" + Str$(HAYNE%) + " Items de Stock.\\Revisar\Cancelar\Continuar")
        On A% GoTo 107, 106, 108
106     Call CIERRARCH("*.*")
        Screen.MousePointer = 1
        Exit Sub
107     Call FRESHECHO("LINEGAT")
        Call SELECHO("LINEGAT")
        '
        If ALTERNA%("Continuar Proceso") <> 1 Then
          Call CIERRARCH("*.*")
          Screen.MousePointer = 1
          Exit Sub
        End If
     End If
     '
108  FORCANEC07.Check2.Value = 1
     FORCANEC07.Refresh
     FORCANEC07.Refresh
     '
     '
     ' revisa las ordenes de compra pendientes de recepcion
     Call ABRECOMPRAS
     If SUPRICO% = 1 Then GoTo 325
     '
     'SQLX$ = " SELECT * FROM OCOMDETA "
     SQLX$ = " SELECT * FROM OCOMDETA WITH(NOLOCK) "
     SQLX$ = SQLX$ + "WHERE Cod_Inte > 0 "    ' SOLO MATERIAL CODIFICADO
     SQLX$ = SQLX$ + "AND Stat_Ocom < 8 "   ' para que no este anulada ni cumplida
     SQLX$ = SQLX$ + "AND  Stat_Ocom <> -1 " 'Para que no  hayan O/C's sin aprobar
     SQLX$ = SQLX$ + "AND Cant_Rec < (Cant_Ocom - 0.005) "
     '
     Dim OCOTEMP As ADODB.Recordset
     If InStr(EPAC$, "GABAL") > 0 Then
          Set OCOTEMP = FLEXCONN.Execute(SQLX$)   ' PARA QUE UNIFIQUE LAS COMPRAS
        Else
          Set OCOTEMP = FLEXCONN.Execute(FILTSQL$(SQLX$))
     End If
     With OCOTEMP
         Do While Not .EOF
           NUORSE& = XVALO(!N_OrServ)
           If NUORSE& < 1 Then
             COEFCAN = XVALO(!Coef_Unids)
             If COEFCAN < 0.005 Then COEFCAN = 1
             CATOT# = XVALO(!Cant_Ocom) * COEFCAN
             CAPRA# = XVALO(!Cant_Rec) * COEFCAN
             CAPEN = CATOT# - CAPRA#
             If CAPEN > 0 Then
               CIBX% = XVALO(!COD_INTE)
               If STATUIT%(CIBX%) >= 0 Then
                 NC% = XVALO(!Npro_Ocom)
                 NROPED& = XVALO(!NUME_OCOM)
                 FENTRE% = CVINT(!Fentre_Sol)
                 COMPROB$ = "OC." + TRIM$(Str$(NROPED&))
                 CICX% = CIBX%
                 Call COBERFA(FENTRE%, CIBX%, COMPROB$ + " " + TEL2CLI$((-1) * NC%), 13, CAPEN)
                 If HAYPLAFAB%(CICX%) < 1 Then
                    DISPACT(CICX%) = DISPACT(CICX%) + CAPEN
                 End If
               End If
             End If
           End If
         .MoveNext
         Loop
       'End If
     End With
     '
     ' revisa las boletas de recepción pendientes de aprobacion
325  Call ABRESTOCKS
     If MODO% = 1 Then GoTo 326            ' CERO STOCKS Y MINIMOS
     '
     'SQLX$ = " SELECT * FROM BOLRECEP "
     SQLX$ = " SELECT * FROM BOLRECEP WITH(NOLOCK) "
     SQLX$ = SQLX$ + "WHERE Cod_Inte > 0 "    ' SOLO MATERIAL CODIFICADO
     SQLX$ = SQLX$ + "AND Status = 0 "   ' para que no este anulada ni cumplida
     '
     Dim COBRECEPTMP As ADODB.Recordset
     If InStr(EPAC$, "GABAL") > 0 Then
          Set COBRECEPTMP = FLEXCONN.Execute(SQLX$)
        Else
          Set COBRECEPTMP = FLEXCONN.Execute(FILTSQL$(SQLX$))
     End If
     With COBRECEPTMP
         Do While Not .EOF
           CATOT# = XVALO(!CANTNOM)
           CAPRA# = XVALO(!CanToApro)
           CAPEN = CATOT# - CAPRA#
           If CAPEN > 0 Then
             CIBX% = XVALO(!COD_INTE)
             If STATUIT%(CIBX%) >= 0 Then
               NC% = XVALO(!NUMPROV)
               NROPED& = XVALO(!NUMEBORE)
               FENTRE% = CVINT(!FechRecep)
               COMPROB$ = "BR." + TRIM$(Str$(NROPED&))
               CICX% = CIBX%
               Call COBERFA(FENTRE%, CIBX%, COMPROB$ + " " + TEL2CLI$((-1) * NC%), 11, CAPEN)
               If HAYPLAFAB%(CICX%) < 1 Then
                  DISPACT(CICX%) = DISPACT(CICX%) + CAPEN
               End If
             End If
           End If
         .MoveNext
         Loop
     End With
     '
' ******** agrega ordenes de fabricacion lanzadas y en curso ******
     '
326  Call ABREORFAB
     If SUPRIFA% = 1 Then GoTo 327
     '
     'DEMORA 1:20
     'SQLX$ = " SELECT * FROM ORDEFABR "
     'SQLX$ = SQLX$ + "WHERE StatuOrf < 8 "   ' para que no este anulada ni cumplida
     'SQLX$ = SQLX$ + "AND CanApro < (CanProy - 0.005) "
     
     'DEMORA 0:45
     SQLX$ = "SELECT * FROM ORDEFABR INNER JOIN MASTER WITH(NOLOCK) "
     SQLX$ = SQLX$ + " ON ORDEFABR.COD_INTE = MASTER.CODINT "
     SQLX$ = SQLX$ + "WHERE ORDEFABR.StatuOrf < 8 "   ' para que no este anulada ni cumplida
     SQLX$ = SQLX$ + "AND MASTER.Status > 0 "
     SQLX$ = SQLX$ + "AND ORDEFABR.CanApro < (ORDEFABR.CanProy - 0.005) AND MASTER.STATUS > 0"
     FIN& = CantRegistros("ORDEFABR", "CanApro < (CanProy - 0.005)", "NOMESS")
     i& = 0
     '
     Dim OFATEMP As ADODB.Recordset
     Set OFATEMP = FLEXCONN.Execute(FILTSQL$(SQLX$))
     With OFATEMP
         Do While Not .EOF
           i& = i& + 1
           Call TRACE(24, i&, FIN&)
           CI% = XVALO(!COD_INTE)
'           If STATUIT%(CI%) >= 0 Then
              If MODOPLAN% <> 1 Or HAYPLAFAB%(CI%) < 1 Then 'descarta ordenes de fabricación productos terminados
                CACOM = XVALO(!CanProy)
                CAREC = XVALO(!CanApro)
                CAPEN = CACOM - CAREC
                If CAPEN > 0 Then
                  NROPED& = XVALO(!NUMEoRFA)
                  FEX = CVINT(!FechEnSol)
                  REFOR$ = SAFETEXT$(!Referen)
                  '
                  STATU% = XVALO(!StatuOrf)
                  If STATU% <= 2 Then ' esta proyectada o lanzada
                    '
                    NORFA$ = SAFETEXT$(!IDSUBOR)
                    NUORFA& = Val(Mid$(NORFA$, 4, 4))
                    SUBOR% = Val(Mid$(NORFA$, 9, 2))
                    FENTRE% = CVINT(!FechProyEntre) 'CVI(Mid$(F$, 147, 2))
                    If FENTRE% < HOO% Then FENTRE% = CVINT(!FechEnSol)   ' CVI(Mid$(F$, 105, 2))
                    If FENTRE% < HOO% Then FENTRE% = HOO%
                    FELANZ% = CVINT(!FechLan)   'CVI(Mid$(F$, 109, 2))
                    If FELANZ% < HOO% Then FELANZ% = HOO%
                    Stat$ = " (proyectada)"
                    If STATU% = 1 Then Stat$ = " (lanzada)"
                    If STATU% = 2 Then Stat$ = " (suspendida)"
                    '
                    DISPACT(CI%) = DISPACT(CI%) + CAPEN
                    '
                    CIBX% = CI%
                    'COD$ = CODEXT$(CI%) ' SE COMENTA PORQUE NO SE UTILIZA
                    COEF = 1
                    '
                    Call COBERFA(FENTRE%, CIBX%, NORFA$ + Stat$, 12, CAPEN)
                    '
                  End If
                End If
              End If
'          End If
         .MoveNext
         Loop
     End With
     '
327  Screen.MousePointer = 11
35   Call ABRESERVA
     If SUPRIFA% = 1 Then GoTo 328
     '
     'DEMORA 9 MIN
     'SQLX$ = " SELECT * FROM RESERVA "
     'SQLX$ = SQLX$ + "WHERE CantCons < (CantRes - 0.005) "
     '
     'DEMORA 1 MIN
     SQLX$ = "SELECT * FROM RESERVA INNER JOIN MASTER WITH(NOLOCK) "
     SQLX$ = SQLX$ + " ON RESERVA.COD_INTE = MASTER.CODINT "
     SQLX$ = SQLX$ + "WHERE RESERVA.CantCons < (RESERVA.CantRes - 0.005) "
     FIN& = CantRegistros("RESERVA", "CantCons < (CantRes - 0.005)", "NOMESS")
     i& = 0
     '
     Dim RESERTEMP As ADODB.Recordset
     If InStr(EPAC$, "GABAL") > 0 Then
          Set RESERTEMP = FLEXCONN.Execute(SQLX$)
        Else
          Set RESERTEMP = FLEXCONN.Execute(FILTSQL$(SQLX$))
     End If
     With RESERTEMP
         Do While Not .EOF
           i& = i& + 1
           Call TRACE(24, i&, FIN&)
           CI1% = XVALO(!COD_INTE)
           If CI1% > 0 Then
             TORE# = XVALO(!CANTRES)
             CACON# = XVALO(!CANTCONS)
             '
'            If TRIM$(UCase$(UNIMED$(CI1%))) = "GK" Then
             If TRIM$(UCase$(SAFETEXT$(!UMEDIDA))) = "GK" Then
               TORE# = TORE# / 1000
               CACON# = CACON# / 1000
             End If
             '
             SIGNO% = Sgn(TORE#)
             SARE = TORE# - CACON#
             If SARE * SIGNO% < 0 Then
               SARE = 0
             End If
             CAPEN = SARE
             If CAPEN > 0 Then
               NORFA$ = SAFETEXT$(!IDSUBOR)
               DISPACT(CI1%) = DISPACT(CI1%) - CAPEN
               FELANZ% = CVINT(!FechRes)
               If InStr(1, EMPREAC$, "FERVI") > 0 Then
                    CIXICONJ% = XVALO(VALOBADA("ORDEFABR", "COD_INTE", "IDSUBOR='" & NORFA$ & "'"))
                    FELANZ% = CVINT(VALOBADA("ORDEFABR", "FechLan", "IDSUBOR='" & NORFA$ & "'"))
                    PLADISPO% = XVALO(VALOBADA("FORMULAS", "PLADISPO", "codiconj=" & CIXICONJ% & " and " & "codielem=" & CI1%))
                    If PLADISPO% < 1 Then PLADISPO% = PLADISTA%
                    FF% = FELANZ%: FELANZ% = DIASANTES(FF%, PLADISPO% * 7)
               End If
               ' CONSIDERA EL LEADTIME EN DIAS DEL MAESTRO DE ARTÍCULOS.
               ' ES GENERAL Y NO ESTA RELACIONADO CON NINGUN ARTÍCULO PADRE COMO EL PLAZO DE DISPOSICION
               ' CUANDO EXISTE, PASA POR ENCIMA DEL PLAZO DE DISPOSICION
               LEADTIME% = XVALO(VALOBADA("MASTER", "LETICOM", "Codint=" & CIXICONJ%))
               If LEADTIME% > 0 Then
                   LTI_SEMANAS% = LEADTIME% Mod 7
                   If LTI_SEMANAS% > 0 Then LTI_SEMANAS% = LTI_SEMANAS% + 1
                   PLADISPO% = LTI_SEMANAS%
                   FELANZ% = CVINT(!FechRes)
                   FF% = FELANZ%: FELANZ% = DIASANTES(FF%, PLADISPO% * 7)
               End If
               Call COBERFA(FELANZ%, CI1%, "Res.p/" + NORFA$, 22, (-1) * CAPEN)
             End If
           End If
         .MoveNext
         Loop
     End With
     '
     '
328  ' ******** inicializa necesidades secundarias de compra y fabricacion ********
110  FORCANEC07.Check3.Value = 2
     Screen.MousePointer = 11
     '
111  N13% = FreeFile
     On Error Resume Next
     Open LU$ + "NEFACOM.DAT" For Random Access Read Write Shared As #N13% Len = LOREC%
     If Err Then
        ERRU% = Err.Number
        On Error GoTo 0
        TXTX$ = "Error no Recuperable " & ERRU% & " al Ejecutar 'Open " & LU$ + "NEFACOM.DAT'."
        MsgBox TXTX$
        Call FINAL("")
     End If
     On Error GoTo 0
     '
     Dim PDS As String * 208
     Dim PDM As String * 72
     '
     If LOF(N13%) > 0 Then
         Close #N13%
         On Error Resume Next
         Kill LU$ + "NEFACOM.DAT"
         If Err Then
            ERRU% = Err.Number
            On Error GoTo 0
            TXTX$ = "Error no Recuperable " & ERRU% & " al Borrar 'Kill " & LU$ + "NEFACOM.DAT'."
            MsgBox TXTX$
            Call FINAL("")
         End If
         On Error GoTo 0
         GoTo 111
     End If
     '
'     MsgBox "Sin Error de Acceso a Archivo. Pulse 'Aceptar'"
     '
     For i& = 1 To NMS&
         CIIX& = i&
         Call TRACE(20, CIIX&, NMS&)
         If LOREC% = 208 Then
              Get #N13%, i&, PDS$
            Else
              Get #N13%, i&, PDM$
         End If
         PDS$ = String$(208, 0)
         PDM$ = String$(72, 0)
         If LOREC% = 208 Then
              Put #N13%, i&, PDS$
            Else
              Put #N13%, i&, PDM$
         End If
     Next i&

     FORCANEC07.Check3.Value = 1

' **************** necesidades brutas segun entregas programadas **************
120  N9% = FreeFile
     Open LU$ + "NEBRUNE.DAT" For Random Access Read Write Shared As #N9% Len = 48
     Dim B00 As String * 48
     Dim B01 As String * 48
     X1& = 1: X2& = 0

     If LOF(N9%) > 0 Then
        Close #N9%
        Kill LU$ + "NEBRUNE.DAT"
        GoTo 120
     End If

     FORCANEC07.Check4.Value = 2
     Screen.MousePointer = 11
     '
121  FIN& = NUMAS%
     For i& = 1 To FIN&
       CIIX& = i&
       Call TRACE(20, CIIX&, FIN&)
       CI% = i&
       If STATUIT%(CI%) >= 0 Then
         '
         If DISPACT(CI%) < 0 Then
           CANTU = Abs(DISPACT(CI%))
           FF% = HOO%
           '
           X2& = X2& + 1
           B00$ = String$(48, 0)
           Call REPLA(B00$, MKI$(CI%), 1, 2)
           Call REPLA(B00$, MKS$(CANTU), 3, 4)
           Call REPLA(B00$, MKI$(FF%), 7, 2)
           Call REPLA(B00$, "PROGRAMA POR MINIMOS", 9, 29)
           Call REPLA(B00$, MKS$(0), 38, 4)
           Call REPLA(B00$, MKS$(0), 42, 4)
           Call REPLA(B00$, MKS$(0), 46, 2)
           Call REPLA(B00$, " ", 48, 1)
           '
           Put #N9%, X2&, B00$
           DISPACT(CI%) = 0
           '
         End If
       End If
122  Next i&
     '
     N4% = FreeFile
     Open LU$ + APROGRA$ + ".DAT" For Random Access Read Write Shared As #N4% Len = 24
     Dim P00 As String * 24
     FIN& = LOF(N4%) / 24
     For i& = 1 To FIN&
       CIIX& = i&
       Call TRACE(20, CIIX&, FIN&)
       Get #N4%, i&, P00$
       P1$ = Left$(P00$, 2)
       CI% = CVI(P1$)
       '
       If CI% > 0 Then
         If CI% <= NUMAS% Then
           If STATUIT%(CI%) >= 0 Then
             '
             P3$ = Mid$(P00$, 3, 2)
             P6$ = Mid$(P00$, 7, 4)
             P7$ = Mid$(P00$, 11, 4)
             '
             If CVI(P3$) < COMESACT% Then
                CIX% = CVI(P1$)
                Call MENSERR(24, "Error: Codigo " + TRIM$(CODEXT$(CIX%)) + " Mal Fecha Entrega Programada")
                GoTo 129
             End If
             '
             PLAVEN = CVS(P6$)
             ENTREG = CVS(P7$)
             PEDIPE = CVS(Mid$(P00$, 19, 4))
             CANTU = PLAVEN - ENTREG
             '
             PEDINOPRO = PEDIPE + ENTREG - PLAVEN
             If PEDINOPRO < 0 Then PEDINOPRO = 0
             '
             CI% = CVI(P1$)
             FF% = CVI(P3$)
             '
             If CANTU > 0 Then
               If CI% > 0 Then
                  If CI% <= NMS& Then
                     DPI%(CI%) = 3
                  End If
               End If
             End If
             '
             If CANTU > 0 Then
               '
               X2& = X2& + 1
               B00$ = String$(48, 0)
               Call REPLA(B00$, MKI$(CI%), 1, 2)
               Call REPLA(B00$, MKS$(CANTU), 3, 4)
               Call REPLA(B00$, MKI$(FF%), 7, 2)
               Call REPLA(B00$, "PROGRAMA ENTREGAS AL " + HO$, 9, 29)
               If MODOPLAN% = 1 Then
                 Call REPLA(B00$, "PLAN DE FABRICACION AL " + HO$, 9, 29)
               End If
               If FF% < INIPER%(1) Then
                   Call REPLA(B00$, MKI$(INIPER%(1)), 7, 2)
                   Call REPLA(B00$, "ATRASO DE ENTREGA AL " + HO$, 9, 29)
               End If
               Call REPLA(B00$, MKS$(0), 38, 4)
               Call REPLA(B00$, MKS$(0), 42, 4)
               Call REPLA(B00$, MKS$(0), 46, 2)
               Call REPLA(B00$, " ", 48, 1)
               '
               Put #N9%, X2&, B00$
               '
             End If
             '
             If MODOPLAN% = 0 Then
               If PEDINOPRO > 0 Then
                 '
                 X2& = X2& + 1
                 B00$ = String$(48, 0)
                 Call REPLA(B00$, MKI$(CI%), 1, 2)
                 Call REPLA(B00$, MKS$(PEDINOPRO), 3, 4)
                 Call REPLA(B00$, MKI$(FF%), 7, 2)
                 Call REPLA(B00$, "PEDIDOS NO PROGRAMADOS AL " + HO$, 9, 29)
                 Call REPLA(B00$, MKS$(0), 38, 4)
                 Call REPLA(B00$, MKS$(0), 42, 4)
                 Call REPLA(B00$, MKS$(0), 46, 2)
                 Call REPLA(B00$, " ", 48, 1)
                 '
                 Put #N9%, X2&, B00$
                 '
               End If
             End If
             '
           End If
         End If
       End If
129  Next i&
     '
     FORCANEC07.Check4.Value = 1
     '
     ' ORDENA POR FECHA DE NECESIDAD
     For KKI1& = 1 To X2&
1291    Get #N9%, KKI1&, B00$
        KKB1$ = B00$
        For KKI2& = KKI1& + 1 To X2&
          Get #N9%, KKI2&, B00$
          KKB2$ = B00$
          If CVI(Mid$(KKB2$, 7, 2)) < CVI(Mid$(KKB1$, 7, 2)) Then
            B00$ = KKB2$
            Put #N9%, KKI1&, B00$
            B00$ = KKB1$
            Put #N9%, KKI2&, B00$
            GoTo 1291
          End If
        Next KKI2&
     Next KKI1&

130  CICLO% = CICLO% + 1
     CICLOMA% = 9
     If InStr(UCase$(EMPREAC$), "YALTRE") > 0 Then CICLOMA% = 15
     '
     If CICLO% > CICLOMA% Then GoTo 150
     FORCANEC07.Label1 = TRIM$(Str$(CICLO%))
     FORCANEC07.Label1.Refresh
     '
' ******** CALCULANDO NECESIDADES NETAS SEGUN DISPONIBILIDAD
     '
     FORCANEC07.Check5.Value = 2
     Screen.MousePointer = 11
     For i& = X1& To X2&
       '
       CIIX& = i&
       Call TRACE(20, CIIX&, X2&)
       '
       Get #N9%, i&, B00$
       B1$ = Left$(B00$, 2)
       B3$ = Mid$(B00$, 3, 4)
       B4$ = Mid$(B00$, 7, 2)
       '
       CI% = CVI(B1$)
       If CI% > 0 Then
         If CI% <= NMS& Then
           If STATUIT%(CI%) >= 0 Then
             '
             ' ESTO SE AGREGO PARA FORZAR EL PROCESO POR NIVELES
             If NIVEGEN%(CI%) <> CICLO% - 1 Then
                GoTo 139
             End If
             ' **************************************************
             '
             If DISPACT(CI%) >= CVS(B3$) Then
               DISPACT(CI%) = DISPACT(CI%) - CVS(B3$)
               Call REPLA(B00$, B3$, 38, 4)
               Call REPLA(B00$, MKS$(0), 42, 4)
               Put #N9%, i&, B00$
               GoTo 139
             End If
             '
             CANEC = CVS(B3$) - DISPACT(CI%)
             FENEBRU% = CVI(B4$): CUANDO% = 1
             '
             For UX% = 1 To CACOL%
               If INIPER%(UX%) <= FENEBRU% Then CUANDO% = UX%
             Next UX%
             '
             If LOREC% = 208 Then
                 Get #N13%, CI%, PDS$
                 PDX$ = PDS$
               Else
                 Get #N13%, CI%, PDM$
                 PDX$ = PDM$
             End If
             '
             For UX% = CUANDO% To 1 Step -1
               PROANT = CVS(Mid$(PDX$, 4 * UX% - 3, 4))
               If PROANT > 0 Then GoTo 135
             Next UX%
             UX% = CUANDO%
135          If PROANT >= LOREPOS(CI%) Then UX% = CUANDO%
             '
             PROANT = CVS(Mid$(PDX$, 4 * UX% - 3, 4))
             ' INTERVENCION PARA CONSIDERAR LOTE INDIVISIBLE - 28/01/08
             CACORRE = PROANT + CANEC
             If LINDIVI(CI%) > 0 Then
                CACORRE = LINDIVI(CI%)
                While CACORRE < PROANT + CANEC: CACORRE = CACORRE + LINDIVI(CI%): Wend
             End If
             CANADI = CACORRE - (PROANT + CANEC)
             'Mid$(PDX$, 4 * UX% - 3, 4) = MKS$(PROANT + CANEC)
             Mid$(PDX$, 4 * UX% - 3, 4) = MKS$(CACORRE)
             '
             If LOREC% = 208 Then
                 PDS$ = PDX$
                 Put #N13%, CI%, PDS$
               Else
                 PDM$ = PDX$
                 Put #N13%, CI%, PDM$
             End If
             '
             FENENE% = INIPER%(UX%)
             '
             Call REPLA(B00$, MKS$(DISPACT(CI%)), 38, 4)
             DISPACT(CI%) = CANADI
             Call REPLA(B00$, MKS$(CANEC + CANADI), 42, 4)
             '
             Call REPLA(B00$, MKI$(FENENE%), 46, 2)
             Put #N9%, i&, B00$
             '
           End If
         End If
       End If
139  Next i&
     '
140  ' Call MENSAJE(14, "GENERANDO NECESIDADES SECUNDARIAS SEGUN EXPLOSION")

     X3& = X2&
     IAGRE& = 0
     For i& = X1& To X2&
       CIIX& = i&
       UKI& = i& - X1& + 1
       Call TRACE(20, UKI&, X2& - X1& + 1)
       '
       Get #N9%, i&, B00$
       '
       ' ESTO SE AGREGO PARA FORZAR EL PROCESO POR NIVELES
       KKII% = CVI(Left$(B00$, 2))
       If NIVEGEN%(KKII%) <> CICLO% - 1 Then
          IAGRE& = IAGRE& + 1
          If IAGRE& >= UDIM& Then
              UDIM& = UDIM& + 1024
              ReDim Preserve BAGRE$(UDIM&)
          End If
          BAGRE$(IAGRE&) = B00$
          B00$ = String$(48, 0)
          Put #N9%, i&, B00$
          GoTo 1236
       End If
       ' **************************************************
       
       Call REPLA(B00$, "C", 48, 1)
       Put #N9%, i&, B00$
       '
       B1$ = Left$(B00$, 2)
       B3$ = Mid$(B00$, 3, 4)
       B4$ = Mid$(B00$, 7, 2)
       B7$ = Mid$(B00$, 42, 4)
       B8$ = Mid$(B00$, 46, 2)
       '
       CI% = CVI(B1$)
       If STATUIT%(CI%) >= 0 Then
         CABAS = CVS(B7$)
         FeNec% = CVI(B8$): CUANEC% = 1
                         For UX% = 1 To CACOL%
                            If INIPER%(UX%) <= FeNec% Then CUANEC% = UX%
                         Next UX%
         If CABAS > 0 Then
1234        CIBAS% = CI%
           CII% = CIBAS%
           CODBAS$ = TRIM$(CODEXT$(CII%))
           ' CONSIDERA EL LEADTIME EN DIAS DEL MAESTRO DE ARTÍCULOS.
           Call LEEEXPLOMUL(CODBAS$, CABAS, 1, , 1)
           If CAIT% > 0 Then
             Call REPLA(B00$, "F", 48, 1)
             Put #N9%, i&, B00$
             For KJ% = 1 To CAIT%    ' J%
               CI% = CIJ%(KJ%)     '  CII1%(KJ%)
               If CI% < 1 Or CI% > NUMAS% Then
                  Call MENSERR(24, "Codigo Componente no Válido\en Estructura de '" + CODBAS$ + "'")
                  GoTo 1235
               End If
               '
               If STATUIT%(CI%) >= 0 Then
                 If CI% = 2706 Then
                 A = B
                 End If
                 CANTU = CABAS * USOI(KJ%) 'CANT1(KJ%)
                 PLADISPO% = PDI%(KJ%)
                 If PLADISPO% < 1 Then PLADISPO% = PLADISTA%
                 CLX% = CUANEC% - PLADISPO%  'PAD1%(KJ%)
                 If CLX% < 1 Then CLX% = 1
                 If CANTU > 0 Then
                   '
                   For KX1& = 1 To IAGRE&
                     If CVI(Left$(BAGRE$(KX1&), 2)) = CIJ%(KJ%) Then
                       If CVI(Mid$(BAGRE$(KX1&), 7, 2)) = INIPER%(CLX%) Then
                         If TRIM$(Mid$(BAGRE$(KX1&), 9, 28)) = TRIM$("Res.p/" + CODBAS$) Then
                           CANTU22 = CANTU + CVS(Mid$(BAGRE$(KX1&), 3, 4))
                           B00$ = BAGRE$(KX1&)
                           Call REPLA(B00$, MKS$(CANTU22), 3, 4)
                           BAGRE$(KX1&) = B00$
                           GoTo 1235
                         End If
                       End If
                     End If
                   Next KX1&
                   '
                   B00$ = String$(48, 0)
                   Call REPLA(B00$, MKI$(CIJ%(KJ%)), 1, 2)
                   Call REPLA(B00$, MKS$(CANTU), 3, 4)
                   Call REPLA(B00$, MKI$(INIPER%(CLX%)), 7, 2)
                   Call REPLA(B00$, "Res.p/" + CODBAS$, 9, 28)
                   Call REPLA(B00$, " ", 48, 1)
                   IAGRE& = IAGRE& + 1
                   If IAGRE& >= UDIM& Then
                     UDIM& = UDIM& + 1024
                     ReDim Preserve BAGRE$(UDIM&)
                   End If
                   BAGRE$(IAGRE&) = B00$
                 End If
               End If
1235         Next KJ%
           End If
         End If
       End If
1236 Next i&
     '
1237 For JAGRE& = 1 To IAGRE&
       X3& = X3& + 1
       B00$ = BAGRE$(JAGRE&)
       Put #N9%, X3&, B00$
     Next JAGRE&
     '
     X1& = X2& + 1
     X2& = X3&: If X2& < X1& Then GoTo 150 ' terminar
     GoTo 130
 
150  FORCANEC07.Check5.Value = 1
     FORCANEC07.Check6.Value = 2
     Screen.MousePointer = 11
     '
     For i& = 1 To NMS&
       '
       CIIX& = i&
       Call TRACE(20, CIIX&, NMS&)
       If LOREC% = 208 Then
           Get #N13%, i&, PDS$
           PDX$ = PDS$
         Else
           Get #N13%, i&, PDM$
           PDX$ = PDM$
       End If
       '
       For J% = 1 To CACOL%
         CAFAB = CVS(Mid$(PDX$, 4 * J% - 3, 4))
         If CAFAB > 0 Then
           REFE$ = "O/F a emitir": Tipo% = 15
'           If DPI%(I&) = 0 Then REFE$ = "Tentativo de Compras": TIPO% = 14
'           UII% = I&
            UII% = i&
            If DPI%(i&) = 0 Or PROVHABI%(UII%) > 0 Then REFE$ = "Tentativo de Compras": Tipo% = 14
           Call COBERFA(INIPER%(J%), UII%, REFE$, Tipo%, CAFAB)
         End If
       Next J%
     Next i&
     '
     For i& = 1 To X3&
       CIIX& = i&
       Call TRACE(20, CIIX&, X3&)
       Get #N9%, i&, B00$
       B1$ = Left$(B00$, 2)
       B3$ = Mid$(B00$, 3, 4)
       B4$ = Mid$(B00$, 7, 2)
       B5$ = Mid$(B00$, 9, 29)
       B7$ = Mid$(B00$, 42, 4)
       B8$ = Mid$(B00$, 46, 2)
       '
       CI% = CVI(B1$)
       If CI% > 0 Then
         If STATUIT%(CI%) >= 0 Then
           CANEC = (-1) * CVS(B3$)
           FFX% = CVI(B4$)
           REFE$ = B5$
           Tipo% = 22
           If Left$(REFE$, 20) <> "PROGRAMA POR MINIMOS" Then
             If Left$(REFE$, 4) = "PEDI" Then Tipo% = 21
             Call COBERFA(FFX%, CI%, REFE$, Tipo%, CANEC)
           End If
         End If
       End If
     Next i&
     '
     FORCANEC07.Check6.Value = 1
     '
3900 Call CIERRARCH("*.*")
     If N4% > 0 Then Close #N4%: N4% = 0
     If N5% > 0 Then Close #N5%: N5% = 0
     If N7% > 0 Then Close #N7%: N7% = 0
     If N9% > 0 Then Close #N9%: N9% = 0
     If N13% > 0 Then Close #N13%: N13% = 0
     '
     KKLL& = 0
     For KKK% = 1 To 52
       If TRIM$(DEPER$(KKK%)) <> "" Then
         KKLL& = KKLL& + 1
         Call GRAREG("PERPROMA", DEPER$(KKK%) + Space$(16), KKLL&)
       End If
     Next KKK%
     Call SETULTREG("PERPROMA", KKLL&)
     Call CIERRARCH("PERPROMA")
     '
     For UI& = 0 To 32767
       PROVINT%(UI&) = 0
     Next UI&
     '
     For U& = 1 To ULTREG&("TAPROINT")
       X$ = REGLEIDO$("TAPROINT", U&)
       NP% = CVI(Left$(X$, 2))
       If NP% > 0 Then
         PROVINT%(NP%) = 1
       End If
     Next U&
     Call CIERRARCH("TAPROINT")
     '
     JC& = 0
     JF& = 0
     '
     FORCANEC07.Check7.Value = 2
     Screen.MousePointer = 11
     '
     FIN& = ULTREG&("ARTIPRO")
     For III& = 1 To FIN&
       Call TRACE(20, III&, FIN&)
       X$ = REGLEIDO$("ARTIPRO", III&)
       CI% = CVI(Left$(X$, 2))
       If CI% > 0 Then
         If CI% <= NUMAS% Then
           PROVEX$ = Mid$(X$, 3, 2)
           If ECOARCH$("PROVED1", PROVEX$) <> "" Then
             PROVARTI%(CI%) = CVI(PROVEX$)
           End If
         End If
       End If
     Next III&
     '
     URECO& = CACO&
     FIN& = URECO&
     For III& = 1 To FIN&
       '
       Call TRACE(20, III&, FIN&)
       X$ = REGLEIDO$("COBERFA", III&)
       TIMO% = CVI(Mid$(X$, 51, 2))
       CI% = CVI(Mid$(X$, 3, 2))
       FF% = CVI(Mid$(X$, 1, 2))
       PERIO$ = Mid$(X$, 5, 16)
       CANTX = CVS(Mid$(X$, 53, 4))
       ' FFI% = CVI(Mid$(X$, 61, 2))
       '
       If TIMO% = 15 Then
         PHB% = PROVHABI%(CI%) ' PROVEEDOR HABITUAL
         If PHB% > 0 Then
           If PROVINT%(PHB%) = 0 Then ' NO ES PROVEEDOR INTERNO
             TIMO% = 14
           End If
         End If
       End If
       '
       Call REPLA(X$, MKI$(TIMO%), 51, 2)
       Call REPLA(X$, MKS$(0), 61, 4)
       '
       Call GRAREG("COBERFA", X$, III&)
       '
       If DESCOMPO%(CI%) = 1 Then
         GoTo 3930
       End If
       '
       If TIMO% = 14 Then
         '
         MONEI% = MONECOSTO%(CI%)
         COSUDOL = COSUNI(CI%)
         COSUN = COSUDOL * TICAMONE(MONEI%)
         PHB% = PROVHABI%(CI%) ' PROVEEDOR HABITUAL
         If PHB% > 0 Then
           If PROVINT%(PHB%) <> 0 Then ' ES PROVEEDOR INTERNO
             COSUN = 0
           End If
         End If
         '
         'Se graba en PLACOMP
         Z$ = REGBLAN$("PLACOMP")
         Call REPLA(Z$, MKI$(CI%), 1, 2)
         Call REPLA(Z$, MKI$(FF%), 3, 2)
         Call REPLA(Z$, PERIO$, 5, 16)
         Call REPLA(Z$, MKS$(CANTX), 21, 4)
         Call REPLA(Z$, MKS$(0), 25, 4) ' tot.progra.por item
         Call REPLA(Z$, MKS$(0), 29, 4) ' filler
         Call REPLA(Z$, MKI$(PHB%), 33, 2)
         Call REPLA(Z$, MKS$(COSUN), 35, 4)
         Call REPLA(Z$, MKD$(COSUN * CANTX), 39, 8)
         '
         JC& = JC& + 1
         Call GRAREG("PLACOMP", Z$, JC&)
         '
         FF1$ = REGBLAN$("OCOMEMI")
         Call REPLA(FF1$, MKI$(CI%), 1, 2)
         Call REPLA(FF1$, DESCRIT$(CI%), 3, 26)
         Call REPLA(FF1$, CSTRING$(MKS$(CANTX), 4, 10, 1, 2), 29, 10)
         FEX = FF%
         Call REPLA(FF1$, FECHATEX$(FEX), 41, 8)
         Call REPLA(FF1$, MKD$(CDbl(CANTX)), 49, 8)
         Call REPLA(FF1$, MKI$(FF%), 57, 2)
         Call REPLA(FF1$, MKI$(CI%), 59, 2)
         Call GRAREG("OCOMEMI", FF1$, JC&)
         '
       ElseIf TIMO% = 15 Then
         '
         Z$ = REGBLAN$("PLAFABR")
         Call REPLA(Z$, MKI$(CI%), 1, 2)
         Call REPLA(Z$, MKI$(FF%), 3, 2)
         Call REPLA(Z$, PERIO$, 5, 16)
         Call REPLA(Z$, MKS$(CANTX), 21, 4)
         Call REPLA(Z$, MKS$(0), 29, 4)  ' FILLER
         JF& = JF& + 1
         Call GRAREG("PLAFABR", Z$, JF&)
         '
         FF1$ = REGBLAN$("ORFAEMI")
         Call REPLA(FF1$, MKI$(CI%), 1, 2)
         Call REPLA(FF1$, DESCRIT$(CI%), 3, 26)
         Call REPLA(FF1$, CSTRING$(MKS$(CANTX), 4, 10, 1, 2), 29, 10)
         FEX = FF%
         Call REPLA(FF1$, FECHATEX$(FEX), 41, 8)
         Call REPLA(FF1$, MKD$(CDbl(CANTX)), 49, 8)
         Call REPLA(FF1$, MKI$(FF%), 57, 2)
         Call REPLA(FF1$, MKI$(CI%), 59, 2)
         Call GRAREG("ORFAEMI", FF1$, JF&)
         '
       End If
       '
3930 Next III&
     FORCANEC07.Check7.Value = 1
     '
     Screen.MousePointer = 1
     URECO& = CACO&
     Call SETULTREG("COBERFA", URECO&)
     Call CIERRARCH("COBERFA")
     Call SETULTREG("ORFAEMI", JF&)
     Call CIERRARCH("ORFAEMI")
     Call SETULTREG("OCOMEMI", JC&)
     Call CIERRARCH("OCOMEMI")
     Call SETULTREG("PLACOMP", JC&)
     Call CIERRARCH("PLACOMP")
     Call SETULTREG("PLAFABR", JF&)
     Call CIERRARCH("PLAFABR")
     Call CIERRARCH("PERPROMA")
     '
     FECHOR$ = HOS$ + "  " + Left$(Time$, 5)
     Call GRACONTROL("CANECOM", "FECHOCAL", FECHOR$)
     If UCase$(Left$(App.EXEName, 5)) = "FLEXO" Then
         PROMAT00.Label7.Caption = Left$(FECHOR$, 8): PROMAT00.Label7.Refresh
         PROMAT00.Label95.Caption = Mid$(FECHOR$, 11, 5): PROMAT00.Label95.Refresh
       Else
         PROMAT07.Label7.Caption = Left$(FECHOR$, 8): PROMAT07.Label7.Refresh
         PROMAT07.Label95.Caption = Mid$(FECHOR$, 11, 5): PROMAT07.Label95.Refresh
     End If
     Call CIERRARCH("*.*")
     Call BAJALDISCO
     '
     Call FILESORT("COBERFA", "", 6, 6, "ASC")
     Call FILESORT("COBERFA", "", 1, 2, "ASC")
     Call FILESORT("PLACOMP", "", 1, 3, "ASC")
     Call FILESORT("PLAFABR", "", 1, 3, "ASC")
     Call FILESORT("ORFAEMI", "", 4, 5, "ASC")
     Call FILESORT("OCOMEMI", "", 4, 5, "ASC")
     '
     CIAX% = 0
     DEPOSTAKK% = Val(TRIM$(Control$("PROMATER", "DEPOSTOK")))
     '
     Call BORRAREGISTROS("OCOEMI", "COD_INTE <> 0", REG&)
     Dim Ocoemi As ADODB.Recordset
     Set Ocoemi = New ADODB.Recordset
     SQLX1$ = "SELECT * FROM OCOEMI WHERE COD_INTE = 0"
25   On Error Resume Next
     Ocoemi.Open SQLX1$, FLEXCONN, adOpenDynamic, adLockPessimistic
     If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
        On Error GoTo 0
        Call CapturaErrorOpen
         GoTo 25
     End If
     On Error GoTo 0

     '
     For U& = 1 To ULTREG&("PLACOMP")
       XX$ = REGLEIDO$("PLACOMP", U&)
       CII% = CVI(Left$(XX$, 2))
       If CII% <> CIAX% Then
         CIAX% = CII%
         Stakk = STODEPOS(CII%, DEPOSTAKK%): If Stakk < 0 Then Stakk = 0
         Call REPLA(XX$, MKS$(Stakk), 29, 4)
         Call REPLA(XX$, MKS$(STOMIN(CII%)), 55, 4)
         Call REPLA(XX$, MKS$(LOTESTAN(CII%)), 59, 4)
       End If
       '
       LT% = LEAD_TIME%(CII%)
       FeNec = CVI(Mid$(XX$, 3, 2))
       FANT% = DIANTE(FeNec, LT%)
       'Si la fecha de solicitud ya ha pasado, entonces el sistema pone la del dia del calculo
       If FANT% < HOY(HO$) Then
          FANT% = HOY(HO$)
       End If
       Call REPLA(XX$, MKI$(FANT%), 63, 2)
       '
       'Se graba en OCOEMI
       CI% = CVI(Mid$(XX$, 1, 2))
       FF% = CVI(Mid$(XX$, 3, 2))
       PERIO$ = Mid$(XX$, 5, 16)
       CANTX = CVS(Mid$(XX$, 21, 4))
       PHB% = CVI(Mid$(XX$, 33, 2))
       COSUN = CVS(Mid(XX$, 35, 4))
       'Se Graba OCOEMI
       With Ocoemi
           .AddNew
           !COD_INTE = CI%
           !CANTIDAD = CANTX
           !Fentre_Sol = FECHATEX$(FF%)
           !Npro_Ocom = PHB%
           !Raso_prov = Left(rasocli$(PHB% * -1), 48)
           !Uni_Mas = Unimed$(CI%)
           !FechRecep = FECHATEX$(FANT%)
           !COSUNI = COSUN
           !PERIODO = PERIO$
           !CodiEmpr = CodiEmp%
           !StokAct = Stakk
           !STOMIN = STOMIN(CI%)
           !LOREPOS = LOTESTAN(CI%)
           .Update
         End With
       '
'AQUI PONER FECHA LIMITE EMISION O/COMPRA
'FECHA DE NECESIDAD EN 3,2
'EN 63,2 PONER MKI$(DIANTE(FECHA DE NECESIDAD, LEAD TIME))
       Call GRAREG("PLACOMP", XX$, U&)
     Next U&
     Set Ocoemi = Nothing
     Call CIERRARCH("PLACOMP")
     Call CARPROSEL
     '
     CIAX% = 0
     'ACA LE GRABA EL STOCKACTUAL POR DEPOSITO - MINIMO - LOTESTAN.
     For U& = 1 To ULTREG&("PLAFABR")
       XX$ = REGLEIDO$("PLAFABR", U&)
       CII% = CVI(Left$(XX$, 2))
       If CII% <> CIAX% Then
         CIAX% = CII%
         Stakk = STODEPOS(CII%, DEPOSTAKK%): If Stakk < 0 Then Stakk = 0
         Call REPLA(XX$, MKS$(Stakk), 29, 4)
         Call REPLA(XX$, MKS$(STOMIN(CII%)), 55, 4)
         Call REPLA(XX$, MKS$(LOTESTAN(CII%)), 59, 4)
         Call GRAREG("PLAFABR", XX$, U&)
       End If
     Next U&
     '
     Call CIERRARCH("PLAFABR")
     '
999  Call CIERRARCH("*.*")

      'Inicio Autonomia en Calculo de MRP
     Dim ACONSUMO()
     NMSI% = NUMAS%
     ReDim ACONSUMO(NMSI%)
     
     HOII% = HOY(HO$)
     FEDESDE$ = FETEXCOR1$(FECHANUM("01/01/10"))
     FEHASTA$ = FETEXCOR1$(HOII%)
     DIAS2010 = DIENTRE(FECHANUM("31/12/09"), HOII)
     '
     If DIENTRE(FECHANUM("01/01/10"), HOII%) > 182 Then
        FINI% = DIANTE(HOII, 183)
        FEDESDE$ = FETEXCOR1$(FINI%)
        DIAS2010 = 182
     End If
     '
     SQLX$ = "SELECT COD_INTE, SUM(CANTSALID) AS SALIDA, SUM(CANTINGRE) AS INGRESO FROM MOVISTO WHERE "
     SQLX$ = SQLX$ + " (FECHMOV >= '" & FEDESDE$ & "' AND FECHMOV <= '" & FEHASTA$ & "') "
     SQLX$ = SQLX$ + " AND (CODIMOVI ='RT' OR CODIMOVI ='CF' "
     SQLX$ = SQLX$ + " OR CODIMOVI ='DV' "
     SQLX$ = SQLX$ + " OR CODIMOVI ='RC') GROUP BY COD_INTE "
     '
     Dim articulos As ADODB.Recordset
     Set articulos = New ADODB.Recordset
     Set articulos = FLEXCONN.Execute(FILTSQL$(SQLX$))
     '
     'SE GUARDA EL CONSUMO + LA CUENTA
     With articulos
        Do While Not (.EOF)
           CI% = XVALO(!COD_INTE)
           ACONSUMO(CI%) = XVALO(!SALIDA) - XVALO(!INGRESO)
        .MoveNext
        Loop
     End With
     articulos.Close
     '
     If DIAS2010 < 182 Then
       For CI% = 1 To NUMAS%
          Condi$ = "CODINT = " & CI%
          CON# = XVALO(VALOBADA("MASTER", "CONSUPROM", Condi$))
          If DIAS2010 >= 182 Then CON# = 0
          CUENTA# = CON# / 30 * (182 - DIAS2010)
          ACONSUMO(CI%) = ACONSUMO(CI%) + CUENTA#
       Next CI%
     End If
     '
     'SE GUARDA EN PLACOMP
     If InStr(UCase$(EMPREAC$), "GABAL") > 0 Or InStr(UCase$(EMPREAC$), "LAFE") > 0 Then

         For U& = 1 To ULTREG&("PLACOMP")
             REGISTRO$ = REGLEIDO("PLACOMP", U&)
             CI% = CVI(Mid(REGISTRO$, 1, 2))
             CONSUDIARIO# = ACONSUMO(CI%) / 182
             STOACT# = CVS(Mid(REGISTRO$, 29, 4))
             On Error Resume Next
             AUTONOMIA% = STOACT# / CONSUDIARIO#     'STOCKACT(CI%) / CONSUDIARIO#
             If AUTONOMIA% < 0 Then AUTONOMIA% = 0
             If Err > 0 Then
                AUTONOMIA% = 999
             End If
             On Error GoTo 0
             Call REPLA(REGISTRO$, MKI$(AUTONOMIA%), 63, 2)
             Call GRAREG("PLACOMP", REGISTRO$, U&)
         Next U&
    
    '     SE GUARDA EN PLAFABR
         For U& = 1 To ULTREG&("PLAFABR")
             REGISTRO$ = REGLEIDO("PLAFABR", U&)
             CI% = CVI(Mid(REGISTRO$, 1, 2))
             CONSUDIARIO# = ACONSUMO(CI%) / 182
             STOACT# = CVS(Mid(REGISTRO$, 29, 4))
             On Error Resume Next
             AUTONOMIA% = STOACT# / CONSUDIARIO#     'STOCKACT(CI%) / CONSUDIARIO#
             If AUTONOMIA% < 0 Then AUTONOMIA% = 0
             If Err > 0 Then
               AUTONOMIA% = 999
             End If
             On Error GoTo 0
             Call REPLA(REGISTRO$, MKI$(AUTONOMIA%), 63, 2)
             Call GRAREG("PLAFABR", REGISTRO$, U&)
         Next U&
     '
     End If
     
     'DESDE AQUI PASO PLAFABR A SQL
'     Call GRABA_PLAFABR_PLACOMP
     'HASTA AQUI PASO PLAFABR A SQL
     
     Call BAJALDISCO
     finix = Now
     Call COMUNI("Re-Cálculo Completo.")
     '
9999 FORCANEC07.Hide
     '
End Sub

Sub GRABA_COBERFA_a_SQL()
'        Call CREATABLA_PLAFABR
'        Call CREATABLA_OCOEMIHIS
        Call CREATABLA_COBERFA
        Call VectArticulos.cargarvector
        Call VectProveedor.cargarvector
        '
'        'GRABO DE DAT PLAFABR EN SQL TABLA PLAFABR
'        '*******************************************
'        SQLX$ = "SELECT TOP 1 * FROM PLAFABR WHERE ID >= 0 ORDER BY ID DESC "
'        Dim RS As New ADODB.Recordset
'        RS.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
'        With RS
'          On Error Resume Next
'          IDENT& = XVALO(!ID) + 1
'          FEREAL = Now
'          On Error GoTo 0
'          FIN& = ULTREG&("PLAFABR")
'          For U& = 1 To FIN&
'            Call TRACE(20, U&, FIN&)
'            TXTX$ = REGLEIDO("PLAFABR", U&)
'            CI1% = CVI(Mid$(TXTX$, 1, 2))
'            DESC$ = VectArticulos.Descripcion(CI1%)
'            FECHNE% = CVI(Mid$(TXTX$, 3, 2))
'            PERIO_PROG$ = Mid$(TXTX$, 5, 16)
'            CANT_PR = CVS(Mid$(TXTX$, 21, 4))
'            OF_PEN = CVS(Mid$(TXTX$, 25, 4))
'            OF_A_EMIT = CVS(Mid$(TXTX$, 29, 4))
'            NROPR% = CVI(Mid$(TXTX$, 33, 2)) '
'            RAZOPR$ = VectProveedor.nombre(NROPR%)
'            Pre_Un = CVS(Mid$(TXTX$, 35, 4)) '
'            IMP_TOT = CVD(Mid$(TXTX$, 39, 8))
'
'            .AddNew
'            !ID = IDENT&
'            !Cod_Inte = CI1%
'            !Descripcion = VectArticulos.Descripcion(CI1%)
'            !FECHNECE = CVDAT(FECHNE%)
'            !PERIO_PROGRAM = PERIO_PROG$
'            !CANT_PROG = ROUND(CANT_PR, 1)
'            !OF_PEND = ROUND(OF_PEN, 1)
'            !OF_A_EMITIR = ROUND(OF_A_EMIT, 1)
'            !NROPROV = NROPR%
'            !RAZOPROV = RAZOPR$
'            !Pre_Unit = ROUND(Pre_Un, 4)
'            !IMP_TOTAL = ROUND(IMP_TOT, 2)
'            !FECALCULO = FEREAL
'            !CODIMOVI = "OF"
'            !NUSER = USNBR%
'            .Update
'          Next U&
'          RS.Close
'          Set RS = Nothing
'          '
'        End With
'        '
'        'GRABO DE OCOEMI EN SQL TABLA OCOEMIHIS
'        '*******************************************
'        SQLX$ = "SELECT TOP 1 * FROM OCOEMIHIS WHERE ID >= 0 ORDER BY ID DESC "
'        Dim RS1 As New ADODB.Recordset
'        RS1.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
'        With RS1
'            On Error Resume Next
'            IDENT& = XVALO(!ID) + 1
'            FEREAL = Now
'
'            On Error GoTo 0
'
'            SQLX$ = "SELECT * FROM OCOEMI  with(NOLOCK) ORDER BY COD_INTE"
'            Set RS2 = READSET(SQLX$)
'            Do While Not RS2.EOF
'                 .AddNew
'                  ORDP = ORDP + 1
'                 !ID = IDENT&
'                 !FECALCULO = FEREAL
'                 !Raso_prov = RS2!Raso_prov
'                     CI1% = RS2!Cod_Inte
'                 !Cod_Inte = CI1%
'                 !Descripcion = VectArticulos.Descripcion(CI1%)
'                 !Fentre_Sol = RS2!Fentre_Sol
'                 !Npro_Ocom = RS2!Npro_Ocom
'                 !Uni_Mas = RS2!Uni_Mas
'                 !FechRecep = RS2!FechRecep
'                 !COSUNI = RS2!COSUNI
'                 !PERIODO = RS2!PERIODO
'                 !CodiEmpr = RS2!CodiEmpr
'                 !StokAct = RS2!StokAct
'                 !STOMIN = RS2!STOMIN
'                 !LOREPOS = RS2!LOREPOS
'                 !FECALCULO = FEREAL
'                 !CODIMOVI = "OC"
'                 !NUSER = USNBR%
'                 !CANTIDAD = RS2!CANTIDAD
'                 .Update
'                 RS2.MoveNext
'            Loop
'            On Error Resume Next
'            RS2.Close
'            Set RS2 = Nothing
'
'            RS1.Close
'            Set RS1 = Nothing
'
'       End With
       '
        'GRABO DE OCOEMI EN SQL TABLA COBERFA
        '*******************************************
        SQLX$ = "SELECT TOP 1 * FROM COBERFA WHERE ID >= 0 ORDER BY ID DESC "
        Dim RS3 As New ADODB.Recordset
        RS3.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
        With RS3
            On Error Resume Next
            Ident& = XVALO(!ID) + 1
            FEREAL = Now
            ORDP = 0
            On Error GoTo 0
            FIN& = ULTREG&("COBERFA")
            For U& = 1 To ULTREG&("COBERFA")
                Call TRACE(20, U&, FIN&)
                TXTX$ = REGLEIDO("COBERFA", U&)
                CI1% = CVI(Mid$(TXTX$, 3, 2))
                ORDP = ORDP + 1 'ORDEN DE PRESENTACION
                
                ORD = CVI(Mid$(TXTX$, 1, 2))
                CODI$ = VectArticulos.Codigo(CI1%)
                DESC$ = VectArticulos.Descripcion(CI1%)
                MESAN$ = TRIM$(Mid$(TXTX$, 5, 16))
                FECHNECI = CVDAT(CVINT(Right$(MESAN$, 8)))
                
                DETRE$ = Mid$(TXTX$, 21, 30)
                CMO% = CVI(Mid$(TXTX$, 51, 2)) '
                DISP = CVS(Mid$(TXTX$, 53, 4))  '
                COMPROMET = CVS(Mid$(TXTX$, 57, 4))
                .AddNew
                !ID = Ident&
                !COD_INTE = CI1%
                !FECALCULO = FEREAL
                !Codigo = CODI$
                !Descripcion = DESC$
                !ORDENANOMES = ORD
                !MESANO = MESAN$
                !DETREFE = DETRE$
                !TIPOMOVI = CMO%
                !CANTIDISP = DISP
                !CANTICOMPROM = COMPROMET
                !SALDODISP = 0
                !NUSER = USNBR%
                !ORDENPRES = ORDP ' ORDEN DE PRESENTACION
                !IDYFECAL = SAFETEXT$(Ident&) & " - " & SAFETEXT$(FEREAL)
                !FecNecesi = FECHNECI
                .Update
            Next U&
            RS3.Close
            Set RS3 = Nothing
           End With
           '
           'ACTULIZO LOS SALDOS DISPONIBLES POR QUE NECESITO ORDENARLOS
           SQLX$ = "SELECT CANTIDISP,CANTICOMPROM, SALDODISP,COD_INTE,ID FROM COBERFA   WHERE ID = " & Ident&
           SQLX$ = SQLX$ + " ORDER BY COD_INTE, ORDENPRES"
           Dim RS4 As New ADODB.Recordset
           RS4.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
           With RS4
             Do While Not .EOF
                CI1% = XVALO(!COD_INTE)
                If CI1% <> CI1_ANT% Then SALDO = 0
                    QDISP = XVALO(!CANTIDISP)
                    QRES = XVALO(!CANTICOMPROM)
                    SALDO = (SALDO + QDISP) - QRES
                    CI1_ANT% = CI1%
                    !SALDODISP = SALDO
                    .Update
                    CI1_ANT% = CI1%
                .MoveNext
             Loop
           End With
           RS4.Close
           Set RS4 = Nothing
End Sub
     
'
Sub CALNIVEGEN()               ' DETERMINA LOS NIVELES DE EXPLOSIÓN GENERALES
     '                         ' PARA TODOS LOS ARTÍCULOS. ASÍ LOS QUE NO SEAN
     '                         ' COMPONENTE DE NINGUNO SERÁN NIVEL CERO; LUEGO
     '                         ' LOS DE NIVEL 1, ETC.
     NMS% = NUMAS%
     For KI% = 1 To NMS%
        NIVEGEN%(KI%) = 0
     Next KI%
     '
     NIVBAS% = 0
     '
10   HUBOEXPLO% = 0
     FIN& = NMS%
     For KI% = 1 To NMS%
       KIU& = KI%
       Call TRACE(20, KIU&, FIN&)
       If NIVEGEN%(KI%) = NIVBAS% Then
         Call AREXPLO(KI%, 1, 3)
         If CAIT% > 0 Then
           For KKJ% = 1 To CAIT%
             If USOI(KKJ%) >= 0.00005 Then
               CIKJ1% = CIJ%(KKJ%)
               If CIKJ1% <> KI% Then
                 If CIKJ1% > 0 Then
                   NIVEGEN%(CIKJ1%) = NIVBAS% + 1
                   HUBOEXPLO% = 1
                 End If
               End If
             End If
           Next KKJ%
         End If
       End If
105  Next KI%
     '
     If HUBOEXPLO% > 0 Then
       NIVBAS% = NIVBAS% + 1
       If NIVBAS% > 12 Then
          Call MENSERR(24, "Imposible Continuar.\Aparente Problema de Recursividad.")
          Call FINAL("")
       End If
       GoTo 10
     End If
     '
End Sub
