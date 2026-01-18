Attribute VB_Name = "CANECO04"
Dim LOST$(10)
Dim LOSTN(10)
Dim CII1%(1024), CANT1(1024), PAD1%(1024)
Dim INIPER%(52), CAPER(52), DEPER$(52)
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

    CT$ = UCase$(CONTROL$("CANECOM", "PERIODO"))
    SEMANAL% = 0: FORCANEC04.Option5.Value = True
    If InStr(CT$, "SEMAN") > 0 Then
        SEMANAL% = 1
        FORCANEC04.Option4.Value = True
    End If
    '
    CT$ = UCase$(CONTROL$("CANECOM", "PLADISTA"))
    PLADISTA% = Val(CT$)
    If PLADISTA% < 0 Then PLADISTA% = 0
    If PLADISTA% > 6 Then PLADISTA% = 6
    FORCANEC04.Text1.TEXT = TRIM$(Str$(PLADISTA%))
    '
    CT$ = UCase$(CONTROL$("CANECOM", "LOST-TER"))
    LOSTN(3) = Val(CT$): If LOSTN(3) < 0 Then LOSTN(3) = 0
    LOST$(3) = Mid$(Str$(LOSTN(3)), 2)
    FORCANEC04.Text2.TEXT = LOST$(3)
    '
    CT$ = UCase$(CONTROL$("CANECOM", "LOST-MON"))
    LOSTN(2) = Val(CT$): If LOSTN(2) < 0 Then LOSTN(2) = 0
    LOST$(2) = Mid$(Str$(LOSTN(2)), 2)
    FORCANEC04.Text3.TEXT = LOST$(2)
    '
    CT$ = UCase$(CONTROL$("CANECOM", "LOST-MEC"))
    LOSTN(1) = Val(CT$): If LOSTN(1) < 0 Then LOSTN(1) = 0
    LOST$(1) = Mid$(Str$(LOSTN(1)), 2)
    FORCANEC04.Text4.TEXT = LOST$(1)
    '
    CT$ = UCase$(CONTROL$("CANECOM", "LOST-COM"))
    LOSTN(0) = Val(CT$): If LOSTN(0) < 0 Then LOSTN(0) = 0
    LOST$(0) = Mid$(Str$(LOSTN(0)), 2)
    FORCANEC04.Text5.TEXT = LOST$(0)
    '
    CT$ = UCase$(CONTROL$("CANECOM", "REPOMIN"))
    REPOMIN = Val(CT$): If REPOMIN < 1 Then REPOMIN = 1
    REPOMINS$ = Mid$(Str$(REPOMIN), 2)
    '
    FORCANEC04.Check1.Value = 0
    FORCANEC04.Check2.Value = 0
    FORCANEC04.Check3.Value = 0
    FORCANEC04.Check4.Value = 0
    FORCANEC04.Check5.Value = 0
    FORCANEC04.Check6.Value = 0
    FORCANEC04.Check7.Value = 0
    FORCANEC04.Label1 = ""
    '
    CTX% = 0
    '
End Sub
'

Sub GRAPARCAL()
    '
    SEMANAL% = 0: PERIO$ = "MENSUAL"
    If FORCANEC04.Option4.Value = True Then
        SEMANAL% = 1: PERIO$ = "SEMANAL"
    End If
    Call GRACONTROL("CANECOM", "PERIODO", PERIO$)
    '
    PLADISTA% = Val(FORCANEC04.Text1.TEXT)
    Call GRACONTROL("CANECOM", "PLADISTA", TRIM$(Str$(PLADISTA%)))
    '
    LOST$(3) = TRIM$(FORCANEC04.Text2.TEXT): LOSTN(3) = Val(LOST$(3))
    LOST$(2) = TRIM$(FORCANEC04.Text3.TEXT): LOSTN(2) = Val(LOST$(2))
    LOST$(1) = TRIM$(FORCANEC04.Text4.TEXT): LOSTN(1) = Val(LOST$(1))
    LOST$(0) = TRIM$(FORCANEC04.Text5.TEXT): LOSTN(0) = Val(LOST$(0))
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
        While J% < CACOL%
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
Sub CANECOM_VIE()   ' ESTE ERA EL VIGENTE HASTA EL 16 NOV 2007 Y FUE
                    ' DISCONTINUADO Y REEMPLAZADO POR LA SUB CANECOM
     '
     Dim BAGRE() As String
     '
     Dim CLICONPEDID%(32767)    'Este Vector registra los Clientes con Pedidos Clientes.
     '
     UDIM& = 2048
     ReDim Preserve BAGRE$(UDIM&)
     LU$ = LUDAT$
     HOO% = HOY(HO$)
     CACO& = 0
     Screen.MousePointer = 11
     '
100  ' ************************* ABRIENDO ARCHIVOS MAESTROS
     '
     FORCANEC04.Check1.Value = 2
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
     FORCANEC04.Check1.Value = 1
     '
     'If NESTRU& < 1 Then
     '   Call MENSERR(24, "Proceso Imposible.\ \Faltan Definir Fórmulas y\Estructuras de Producto.")
     '   Call FINAL("")
     'End If
     '
     PLADISTA% = Val(FORCANEC04.Text1.TEXT)
     If PLADISTA% > 6 Then PLADISTA% = 6
     '
'    **** DEPURA PROGRAMA DE ENTREGAS POR ENTREGAS REALIZADAS
     '
     APROGRA$ = "PROGENT"
     'If MODOPLAN% = 1 Then APROGRA$ = "PROGFAB"
     '
     ' ESTO QUE SIGUE ES PARA NO ALTERAR EL PROGRAMA ORIGINAL DE ENTREGAS
     ' SE CONFECCIONA UNA IMAGEN EN PROGCALC Y SE TRABAJA SOBRE ESTA
     Call COPYSTRU(APROGRA$, "PROGCALC")
     JJL1& = 0
     For I& = 1 To ULTREG&(APROGRA$)
       ZZ$ = REGLEIDO$(APROGRA$, I&)
       JJL1& = JJL1& + 1
       Call GRAREG("PROGCALC", ZZ$, JJL1&)
     Next I&
     Call SETULTREG("PROGCALC", JJL1&)
     Call CIERRARCH("PROGCALC")
     Call CIERRARCH(APROGRA$)
     APROGRA$ = "PROGCALC"
     ' FIN CONFECCION DE IMAGEN PROGENT.
     '
     UREPRO& = ULTREG&(APROGRA$)
     JJ& = 0
     For I& = 1 To UREPRO&
       ZZ$ = REGLEIDO$(APROGRA$, I&)
       If CVS(Mid$(ZZ$, 7, 4)) > 0.05 Then
         Call REPLA(ZZ$, MKS$(0), 11, 4)
         Call REPLA(ZZ$, MKS$(0), 19, 4)
         JJ& = JJ& + 1
         Call GRAREG(APROGRA$, ZZ$, JJ&)
       End If
     Next I&
     Call SETULTREG(APROGRA$, JJ&)
     Call CIERRARCH(APROGRA$)
     If JJ& < 1 Then
        On Error Resume Next
        Kill LUDAT$ + APROGRA$ + ".DAT"
        On Error GoTo 0
     End If
     '
     '
' LEER ENTREGAS DEL MES PARA NETEAR PROGRAMA DE ENTREGAS
     CMA% = COMESACT%
     Call ABRESTOCKS
     CMAX = CDbl(CVDAT(CMA%))
     SQLX$ = " SELECT * FROM MOVISTO "
     SQLX$ = SQLX$ + "WHERE CDBL(FechMov) >= " & CMAX & " "
     SQLX$ = SQLX$ + "AND (CodiMovi = 'RT' Or CodiMovi = 'DV')"
     '
     Set MoviTemp = Stocks.OpenRecordset(SQLX$, dbOpenSnapshot)
     With MoviTemp
       On Error Resume Next
       .MoveFirst
       If Err < 1 Then
         Do While Not .EOF
           CIXI% = !COD_INTE
           If CIXI% > 0 Then
             If CIXI% <= NUMAS% Then
             AAA = CDbl(!FechMov)
               CAENTREG = !CantSalid - !CantIngre
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
       End If
     End With
     '
'    **** AGREGA PEDIDOS PENDIENTES DE ENTREGA
'    **** DESCARTANDO L0S CORRESPONDIENTES A REPARACIONES
     '
'    If MODOPLAN% <> 1 Then
      FIN& = ULTREG&(APROGRA$)
      RFF$ = ""
      For U& = 1 To FIN&
        X$ = REGLEIDO$(APROGRA$, U&)
        RFF$ = RFF$ + Left$(X$, 2)
      Next U&
      '
     Call ABREPEDCLI
     SQLX$ = " SELECT * FROM PEDDETA "
     'SQLX$ = SQLX$ + "WHERE CodiInt = " & CI% & " "
     SQLX$ = SQLX$ + "WHERE (Status >= 0 AND Status < 8) "   ' para que no este anulada ni cumplida ni pendiente de aprobacion
     SQLX$ = SQLX$ + "AND CantEnt < (CanPed - 0.005) "
     SQLX$ = SQLX$ + "AND VenRep = 'V' "
     SQLX$ = SQLX$ + "ORDER BY FeSolEnt DESC "
     '
     Set PEDTEMP = Ventas.OpenRecordset(SQLX$, dbOpenSnapshot)
     '\\\OT-05-0631-WAR-20/09/05///
     INICLICON% = 0
     If UCase$(CONTROL$("CANECOM", "FILPECLI")) <> "SI" Then INICLICON% = 1
     ' SI NO FILTRA POR CLIENTES CON PEDIDOS, ENTONCES TOMA TODOS
     For JJ1% = 1 To 32767 'PONGO EN 0 EL VECTOR
       CLICONPEDID%(JJ1%) = INICLICON%
     Next JJ1%
     FILPEDIPEN% = 0
     If INICLICON% = 1 Then
        For CIXI% = 1 To NUMAS%: ACTICOD%(CIXI%) = 1: Next CIXI%
        GoTo 50
     End If
     '
     ' CARGA UN VECTOR CON LOS CLIENTES CON PEDIDOS PENDIENTES
     FILPEDIPEN% = 1
     For CIXI% = 1 To NUMAS%: ACTICOD%(CIXI%) = 0: Next CIXI%
     With PEDTEMP
       On Error Resume Next
       .MoveFirst
       If Err < 1 Then
         Do While Not .EOF
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
        RASOCLIE2$ = RASOCLI$(JL&)
        Z$ = Space(128)
        Call REPLA(Z$, NCLIE2$, 1, 6)
        Call REPLA(Z$, RASOCLIE2$, 15, 64)
        With CLICOPE04
          .LICOPEN.AddItem Z$
        End With
      End If
     Next JL&
     Screen.MousePointer = 1
     If CLICOPE04.LICOPEN.ListCount > 0 Then
       CLICOPE04.Show 1
       Call FRESHALL
       If CLICOPE04.Label2.Caption <> "OK" Then
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
       For JL& = 1 To CLICOPE04.LICOPEN.ListCount
         If CLICOPE04.LICOPEN.Selected(JL& - 1) = True Then
           ICLIE1& = Val(Left$(CLICOPE04.LICOPEN.List(JL& - 1), 6))
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
     With PEDTEMP
       On Error Resume Next
       .MoveFirst
       If Err < 1 Then
         Do While Not .EOF
           '
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
           CIXI% = !CodiInt
           Call EXPLOMUL(CIXI%)
           For JJ1% = 1 To CONTAMUL%
             CIX2% = CIK%(JJ1%)
             ACTICOD%(CIX2%) = 2
           Next JJ1%
           '
           NCX1% = NCLIE1
           CAPEN = !CanPed - !CantEnt
           If CAPEN > 0 Then
             FENTRESOL% = CVINT%(!FeSolEnt)
             For KK1& = 1 To Len(RFF$) Step 2    ' - 3 To 1 Step -4
               If CVI(Mid$(RFF$, KK1, 2)) = CIXI% Then
                 REPRO& = (KK1& + 1) / 2 ' CVS(Mid$(RFF$, KK1&, 4))
                 ZZ$ = REGLEIDO$(APROGRA$, REPRO&)
                 FEPROG% = CVI(Mid$(ZZ$, 3, 2))
                 If COMIENFE%(FEPROG%) <= COMIENFE%(FENTRESOL%) Then
                   CAPEPROMAX = CVS(Mid$(ZZ$, 7, 4)) - CVS(Mid$(ZZ$, 11, 4)) - CVS(Mid$(ZZ$, 19, 4))
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
     Call CIERRARCH(APROGRA$)
     '
' hasta aqui esta corregido para entregas / fabricacion
' Y PARA DESCARTAR PEDIDOS POR REPARACIONES
'
' ************** lee existencias, stock minimo y lote reposicion *******************

     FORCANEC04.Check2.Value = 2
     Screen.MousePointer = 11
     '
     ' determina items fanricados con dpi%()=1
     FIN& = NUMAS%
     For I& = 1 To FIN&
       Call TRACE(20, I&, FIN&)
       CI% = I&
       Call AREXPLO(CI%, 1, 3)
       If CAIT% > 0 Then DPI%(CI%) = 1
105  Next I&
     '
     HAYNE% = 0
     modo% = 0
     MODOCAX$ = FORCANEC04.Option3.Caption
     If FORCANEC04.Option1.Value = True Then
         modo% = 1
         MODOCAX$ = FORCANEC04.Option1.Caption
       ElseIf FORCANEC04.Option2.Value = True Then
         modo% = 2
         MODOCAX$ = FORCANEC04.Option2.Caption
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
         If modo% = 1 Then SALSTO = 0
         If modo% = 2 Then If SALSTO < 0 Then SALSTO = 0

         If SALSTO < (-0.05) Then
            HAYNE% = HAYNE% + 1
            CI1% = CI%
            TTXX$ = REGBLAN$("LINEGAT")
            Call REPLA(TTXX$, MKI$(CI%), 1, 2)
            Call REPLA(TTXX$, DESCRIT$(CI%), 3, 30)
            Call REPLA(TTXX$, CSTRING$(MKS$(SALSTO), 4, 10, 1, 2), 33, 10)
            Call REPLA(TTXX$, UNIMED$(CI%), 45, 4)
            '
            QLN& = QLN& + 1
            Call GRAREG("LINEGAT", TTXX$, QLN&)
            '
         End If
         '
         STMIN = STOMIN(CI%): If modo% = 1 Then STMIN = 0
         REPOS = LOTESTAN(CI%)
         DP% = DPI%(CI%): If DP% < 0 Or DP% > 3 Then DP% = 0

         If REPOS < 0.5 Then REPOS = LOSTN(DP%)
         If REPOS < 0.5 Then REPOS = REPOMIN

         DISPACT(CI%) = SALSTO - STMIN
         LOREPOS(CI%) = REPOS
         '
         If Abs(SALSTO) > 0.05 Then
           Call COBERFA(COSEMACT%, CI%, "Stock al " + HO$, 10, SALSTO)
         End If
         '
         If Abs(STMIN) > 0.05 Then
           Call COBERFA(COSEMACT%, CI%, "Stock Minimo del Item", 20, (-1) * STMIN)
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
108  FORCANEC04.Check2.Value = 1
     FORCANEC04.Refresh
     FORCANEC04.Refresh
     '
     '
     ' revisa las ordenes de compra pendientes de recepcion
     Call ABRECOMPRAS
     SQLX$ = " SELECT * FROM OCOMDETA "
     SQLX$ = SQLX$ + "WHERE Cod_Inte > 0 "    ' SOLO MATERIAL CODIFICADO
     SQLX$ = SQLX$ + "AND Stat_Ocom < 8 "   ' para que no este anulada ni cumplida
     SQLX$ = SQLX$ + "AND  Stat_Ocom <> -1 " 'Para que no  hayan O/C's sin aprobar
     SQLX$ = SQLX$ + "AND Cant_Rec < (Cant_Ocom - 0.005) "
     '
     Set OcoTemp = Compras.OpenRecordset(SQLX$, dbOpenSnapshot)
     With OcoTemp
       On Error Resume Next
       .MoveFirst
       If Err < 1 Then
         Do While Not .EOF
           NUORSE& = !N_OrServ
           If NUORSE& < 1 Then
             COEFCAN = !Coef_Unids
             If COEFCAN < 0.005 Then COEFCAN = 1
             CATOT# = !Cant_Ocom * COEFCAN
             CAPRA# = !Cant_Rec * COEFCAN
             CAPEN = CATOT# - CAPRA#
             If CAPEN > 0 Then
               CIBX% = !COD_INTE
               If STATUIT%(CIBX%) >= 0 Then
                 NC% = !Npro_Ocom
                 NroPed& = !Nume_Ocom
                 FENTRE% = CVINT(!Fentre_Sol)
                 COMPROB$ = "OC." + TRIM$(Str$(NroPed&))
                 CICX% = CIBX%
                 Call COBERFA(FENTRE%, CIBX%, COMPROB$ + " " + TEL2CLI$((-1) * NC%), 13, CAPEN)
                 DISPACT(CICX%) = DISPACT(CICX%) + CAPEN
               End If
             End If
           End If
         .MoveNext
         Loop
       End If
     End With
     '
     '
     ' revisa las boletas de recepción pendientes de aprobacion
     Call ABRESTOCKS
     SQLX$ = " SELECT * FROM BOLRECEP "
     SQLX$ = SQLX$ + "WHERE Cod_Inte > 0 "    ' SOLO MATERIAL CODIFICADO
     SQLX$ = SQLX$ + "AND Status = 0 "   ' para que no este anulada ni cumplida
     '
     Set COBRECEPTMP = Stocks.OpenRecordset(SQLX$, dbOpenSnapshot)
     With COBRECEPTMP
       On Error Resume Next
       .MoveFirst
       If Err < 1 Then
         Do While Not .EOF
           CATOT# = !CANTNOM
           CAPRA# = !CanToApro
           CAPEN = CATOT# - CAPRA#
           If CAPEN > 0 Then
             CIBX% = !COD_INTE
             If STATUIT%(CIBX%) >= 0 Then
               NC% = !NumProv
               NroPed& = !NumeBoRe
               FENTRE% = CVINT(!FechRecep)
               COMPROB$ = "BR." + TRIM$(Str$(NroPed&))
               CICX% = CIBX%
               Call COBERFA(FENTRE%, CIBX%, COMPROB$ + " " + TEL2CLI$((-1) * NC%), 11, CAPEN)
               DISPACT(CICX%) = DISPACT(CICX%) + CAPEN
             End If
           End If
         .MoveNext
         Loop
       End If
     End With
     '
' ******** agrega ordenes de fabricacion lanzadas y en curso ******
     '
     Call ABREORFAB
     SQLX$ = " SELECT * FROM ORDEFABR "
     'SQLX$ = SQLX$ + "WHERE Cod_Inte = " & CI% & " "
     SQLX$ = SQLX$ + "WHERE StatuOrf < 8 "   ' para que no este anulada ni cumplida
     SQLX$ = SQLX$ + "AND CanApro < (CanProy - 0.005) "
     'SQLX$ = SQLX$ + "ORDER BY NumeOrFa ASC "
     '
     Set OfaTemp = Ofabric.OpenRecordset(SQLX$, dbOpenSnapshot)
     With OfaTemp
       On Error Resume Next
       .MoveFirst
       If Err < 1 Then
         Do While Not .EOF
           CI% = !COD_INTE
           If STATUIT%(CI%) >= 0 Then
             CACOM = !CanProy
             CAREC = !CanApro
             CAPEN = CACOM - CAREC
             If CAPEN > 0 Then
               NroPed& = !NUMEoRFA
               FEX = CVINT(!FechEnSol)
               REFOR$ = !Referen
               '
               STATU% = !StatuOrf
               If STATU% <= 2 Then ' esta proyectada o lanzada
                 '
                 NORFA$ = !IdSubOr
                 NUORFA& = Val(Mid$(NORFA$, 4, 4))
                 SUBOR% = Val(Mid$(NORFA$, 9, 2))
                 FENTRE% = CVINT(!FechProyEntre) 'CVI(Mid$(F$, 147, 2))
                 If FENTRE% < HOO% Then FENTRE% = !FechEnSol   ' CVI(Mid$(F$, 105, 2))
                 If FENTRE% < HOO% Then FENTRE% = HOO%
                 FELANZ% = CVINT(!FechLan)   'CVI(Mid$(F$, 109, 2))
                 If FELANZ% < HOO% Then FELANZ% = HOO%
                 STAT$ = " (proyectada)"
                 If STATU% = 1 Then STAT$ = " (lanzada)"
                 If STATU% = 2 Then STAT$ = " (suspendida)"
                 '
                 DISPACT(CI%) = DISPACT(CI%) + CAPEN
                 '
                 CIBX% = CI%
                 COD$ = CODEXT$(CI%)
                 COEF = 1
                 '
                 Call COBERFA(FENTRE%, CIBX%, NORFA$ + STAT$, 12, CAPEN)
                 '
               End If
             End If
           End If
         .MoveNext
         Loop
       End If
     End With
     '
     Screen.MousePointer = 11
35   Call ABRESERVA
     SQLX$ = " SELECT * FROM RESERVA "
     SQLX$ = SQLX$ + "WHERE CantCons < (CantRes - 0.005) "
     '
     Set ReserTemp = Stocks.OpenRecordset(SQLX$, dbOpenSnapshot)
     With ReserTemp
       On Error Resume Next
       .MoveFirst
       If Err < 1 Then
         Do While Not .EOF
           CI1% = !COD_INTE
           If CI1% > 0 Then
             TORE# = !CantRes
             CACON# = !CantCons
             '
             If TRIM$(UCase$(UNIMED$(CI1%))) = "GK" Then
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
               NORFA$ = !IdSubOr
               DISPACT(CI1%) = DISPACT(CI1%) - CAPEN
               FELANZ% = CVINT(!FechRes)
               Call COBERFA(FELANZ%, CI1%, "Res.p/" + NORFA$, 22, (-1) * CAPEN)
             End If
           End If
         .MoveNext
         Loop
       End If
     End With
     '
     '
' ******** inicializa necesidades secundarias de compra y fabricacion ********

110  FORCANEC04.Check3.Value = 2
     Screen.MousePointer = 11
     '
111  N13% = FreeFile
     Open LU$ + "NEFACOM.DAT" For Random Access Read Write Shared As #N13% Len = LOREC%
     Dim PDS As String * 208
     Dim PDM As String * 72
     '
     If LOF(N13%) > 0 Then
         Close #N13%
         Kill LU$ + "NEFACOM.DAT"
         GoTo 111
     End If
     '
     For I& = 1 To NMS&
         CIIX& = I&
         Call TRACE(20, CIIX&, NMS&)
         If LOREC% = 208 Then
              Get #N13%, I&, PDS$
            Else
              Get #N13%, I&, PDM$
         End If
         PDS$ = String$(208, 0)
         PDM$ = String$(72, 0)
         If LOREC% = 208 Then
              Put #N13%, I&, PDS$
            Else
              Put #N13%, I&, PDM$
         End If
     Next I&

     FORCANEC04.Check3.Value = 1


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

     FORCANEC04.Check4.Value = 2
     Screen.MousePointer = 11
     '
121  FIN& = NUMAS%
     For I& = 1 To FIN&
       CIIX& = I&
       Call TRACE(20, CIIX&, FIN&)
       CI% = I&
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
122  Next I&
     '
     N4% = FreeFile
     Open LU$ + APROGRA$ + ".DAT" For Random Access Read Write Shared As #N4% Len = 24
     Dim P00 As String * 24
     FIN& = LOF(N4%) / 24
     For I& = 1 To FIN&
       CIIX& = I&
       Call TRACE(20, CIIX&, FIN&)
       Get #N4%, I&, P00$
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
                cix% = CVI(P1$)
                Call MENSERR(24, "Error: Codigo " + TRIM$(CODEXT$(cix%)) + " Mal Fecha Entrega Programada")
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
129  Next I&
     '
     FORCANEC04.Check4.Value = 1
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
     FORCANEC04.Label1 = TRIM$(Str$(CICLO%))
     FORCANEC04.Label1.Refresh
     '
' ******** CALCULANDO NECESIDADES NETAS SEGUN DISPONIBILIDAD
     '
     FORCANEC04.Check5.Value = 2
     Screen.MousePointer = 11
     For I& = X1& To X2&
       '
       CIIX& = I&
       Call TRACE(20, CIIX&, X2&)
       '
       Get #N9%, I&, B00$
       B1$ = Left$(B00$, 2)
       B3$ = Mid$(B00$, 3, 4)
       B4$ = Mid$(B00$, 7, 2)
       '
       CI% = CVI(B1$)
       If CI% > 0 Then
         If CI% <= NMS& Then
           If STATUIT%(CI%) >= 0 Then
             If DISPACT(CI%) >= CVS(B3$) Then
               DISPACT(CI%) = DISPACT(CI%) - CVS(B3$)
               Call REPLA(B00$, B3$, 38, 4)
               Call REPLA(B00$, MKS$(0), 42, 4)
               Put #N9%, I&, B00$
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
             Mid$(PDX$, 4 * UX% - 3, 4) = MKS$(PROANT + CANEC)
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
             DISPACT(CI%) = 0
             Call REPLA(B00$, MKS$(CANEC), 42, 4)
             Call REPLA(B00$, MKI$(FENENE%), 46, 2)
             Put #N9%, I&, B00$
             '
           End If
         End If
       End If
139  Next I&
     '
140  ' Call MENSAJE(14, "GENERANDO NECESIDADES SECUNDARIAS SEGUN EXPLOSION")

     X3& = X2&
     IAGRE& = 0
     For I& = X1& To X2&
       CIIX& = I&
       UKI& = I& - X1& + 1
       Call TRACE(20, UKI&, X2& - X1& + 1)
       Get #N9%, I&, B00$
       Call REPLA(B00$, "C", 48, 1)
       Put #N9%, I&, B00$
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
         FENEC% = CVI(B8$): CUANEC% = 1
                         For UX% = 1 To CACOL%
                            If INIPER%(UX%) <= FENEC% Then CUANEC% = UX%
                         Next UX%
         If CABAS > 0 Then
1234        CIBAS% = CI%
           CII% = CIBAS%
           CODBAS$ = TRIM$(CODEXT$(CII%))
           Call LEEEXPLOMUL(CODBAS$, CABAS, 1)
           If CAIT% > 0 Then
             Call REPLA(B00$, "F", 48, 1)
             Put #N9%, I&, B00$
             For KJ% = 1 To CAIT%    ' J%
               CI% = CIJ%(KJ%)     '  CII1%(KJ%)
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
     Next I&
     '
     For JAGRE& = 1 To IAGRE&
       X3& = X3& + 1
       B00$ = BAGRE$(JAGRE&)
       Put #N9%, X3&, B00$
     Next JAGRE&
     '
     X1& = X2& + 1
     X2& = X3&: If X2& < X1& Then GoTo 150 ' terminar
     GoTo 130
 
150  FORCANEC04.Check5.Value = 1
     FORCANEC04.Check6.Value = 2
     Screen.MousePointer = 11
     '
     For I& = 1 To NMS&
       '
       CIIX& = I&
       Call TRACE(20, CIIX&, NMS&)
       If LOREC% = 208 Then
           Get #N13%, I&, PDS$
           PDX$ = PDS$
         Else
           Get #N13%, I&, PDM$
           PDX$ = PDM$
       End If
       '
       For J% = 1 To CACOL%
         CAFAB = CVS(Mid$(PDX$, 4 * J% - 3, 4))
         If CAFAB > 0 Then
           REFE$ = "O/F a emitir": TIPO% = 15
           If DPI%(I&) = 0 Then REFE$ = "Tentativo de Compras": TIPO% = 14
           UII% = I&
           Call COBERFA(INIPER%(J%), UII%, REFE$, TIPO%, CAFAB)
         End If
       Next J%
     Next I&
     '
     For I& = 1 To X3&
       CIIX& = I&
       Call TRACE(20, CIIX&, X3&)
       Get #N9%, I&, B00$
       B1$ = Left$(B00$, 2)
       B3$ = Mid$(B00$, 3, 4)
       B4$ = Mid$(B00$, 7, 2)
       B5$ = Mid$(B00$, 9, 29)
       B7$ = Mid$(B00$, 42, 4)
       B8$ = Mid$(B00$, 46, 2)
       '
       CI% = CVI(B1$)
       If STATUIT%(CI%) >= 0 Then
         CANEC = (-1) * CVS(B3$)
         FFX% = CVI(B4$)
         REFE$ = B5$
         TIPO% = 22
         If Left$(REFE$, 20) <> "PROGRAMA POR MINIMOS" Then
           If Left$(REFE$, 4) = "PEDI" Then TIPO% = 21
           Call COBERFA(FFX%, CI%, REFE$, TIPO%, CANEC)
         End If
       End If
     Next I&
     '
     FORCANEC04.Check6.Value = 1
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
     FORCANEC04.Check7.Value = 2
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
     FORCANEC04.Check7.Value = 1
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
         PROMATER.Label7.Caption = Left$(FECHOR$, 8): PROMATER.Label7.Refresh
         PROMATER.Label95.Caption = Mid$(FECHOR$, 11, 5): PROMATER.Label95.Refresh
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
     DEPOSTAKK% = Val(TRIM$(CONTROL$("PROMATER", "DEPOSTOK")))
     '
     For U& = 1 To ULTREG&("PLACOMP")
       XX$ = REGLEIDO$("PLACOMP", U&)
       CII% = CVI(Left$(XX$, 2))
       If CII% <> CIAX% Then
         CIAX% = CII%
         STAKK = STODEPOS(CII%, DEPOSTAKK%): If STAKK < 0 Then STAKK = 0
         Call REPLA(XX$, MKS$(STAKK), 29, 4)
         Call REPLA(XX$, MKS$(STOMIN(CII%)), 55, 4)
         Call REPLA(XX$, MKS$(LOTESTAN(CII%)), 59, 4)
         Call GRAREG("PLACOMP", XX$, U&)
       End If
     Next U&
     Call CIERRARCH("PLACOMP")
     Call CARPROSEL
     '
     CIAX% = 0
     For U& = 1 To ULTREG&("PLAFABR")
       XX$ = REGLEIDO$("PLAFABR", U&)
       CII% = CVI(Left$(XX$, 2))
       If CII% <> CIAX% Then
         CIAX% = CII%
         STAKK = STODEPOS(CII%, DEPOSTAKK%): If STAKK < 0 Then STAKK = 0
         Call REPLA(XX$, MKS$(STAKK), 29, 4)
         Call REPLA(XX$, MKS$(STOMIN(CII%)), 55, 4)
         Call REPLA(XX$, MKS$(LOTESTAN(CII%)), 59, 4)
         Call GRAREG("PLAFABR", XX$, U&)
       End If
     Next U&
     Call CIERRARCH("PLAFABR")
     '
999  Call CIERRARCH("*.*")
     Call BAJALDISCO
     Call COMUNI("Re-Cálculo Completo.")
     '
9999 FORCANEC04.Hide
     '
End Sub
'
Sub COBERFA(FFI%, CI%, REFE$, TIPO%, CAN)
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
         MESANXX& = 30 * (54 * ANIOX% + SMN%) + TIPO%
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
      MESANIOI% = 100 * (13 * ANIOX% + MES%) + TIPO%
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
                    If CVI(F0$) = TIPO% Then
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
      Call REPLA(FF0$, MKI$(TIPO%), 51, 2)
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



Sub PRODISP(FF%, CI%, TIPO%, CAN)

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

Sub CANECOM_REV()    ' NUEVO AL 13 DE AGOSTO
     'ESTE ESTÁ PENDIENTE DE REVISIÓN
     '
     Dim BAGRE() As String
     Dim CLICONPEDID%(32767)    'Este Vector registra los Clientes con Pedidos Clientes.
     '
     UDIM& = 2048
     ReDim Preserve BAGRE$(UDIM&)
     LU$ = LUDAT$
     HOO% = HOY(HO$)
     CACO& = 0
     Screen.MousePointer = 11
     '
100  ' ************************* ABRIENDO ARCHIVOS MAESTROS
     '
     FORCANEC04.Check1.Value = 2
     Screen.MousePointer = 11
     NMS& = NUMAS%
     FORCANEC04.Check1.Value = 1
     '
     PLADISTA% = Val(FORCANEC04.Text1.TEXT)
     If PLADISTA% > 6 Then PLADISTA% = 6
     '
     APROGRA$ = "PROGENT"
     'If MODOPLAN% = 1 Then APROGRA$ = "PROGFAB"
     '
     ' ESTO QUE SIGUE ES PARA NO ALTERAR EL PROGRAMA ORIGINAL DE ENTREGAS
     ' SE CONFECCIONA UNA IMAGEN EN PROGCALC Y SE TRABAJA SOBRE ESTA
     Call COPYSTRU(APROGRA$, "PROGCALC")
     JJL1& = 0
     For I& = 1 To ULTREG&(APROGRA$)
       ZZ$ = REGLEIDO$(APROGRA$, I&)
       JJL1& = JJL1& + 1
       Call GRAREG("PROGCALC", ZZ$, JJL1&)
     Next I&
     Call SETULTREG("PROGCALC", JJL1&)
     Call CIERRARCH("PROGCALC")
     Call CIERRARCH(APROGRA$)
     APROGRA$ = "PROGCALC"
     ' FIN CONFECCION DE IMAGEN PROGENT.
     '
     UREPRO& = ULTREG&(APROGRA$)
     JJ& = 0
     For I& = 1 To UREPRO&
       ZZ$ = REGLEIDO$(APROGRA$, I&)
       If CVS(Mid$(ZZ$, 7, 4)) > 0.05 Then
         Call REPLA(ZZ$, MKS$(0), 11, 4)
         Call REPLA(ZZ$, MKS$(0), 19, 4)
         JJ& = JJ& + 1
         Call GRAREG(APROGRA$, ZZ$, JJ&)
       End If
     Next I&
     Call SETULTREG(APROGRA$, JJ&)
     Call CIERRARCH(APROGRA$)
     If JJ& < 1 Then
        On Error Resume Next
        Kill LUDAT$ + APROGRA$ + ".DAT"
        On Error GoTo 0
     End If
     '
' LEER ENTREGAS DEL MES PARA NETEAR PROGRAMA DE ENTREGAS
     CMA% = COMESACT%
     Call ABRESTOCKS
     CMAX = CDbl(CVDAT(CMA%))
     SQLX$ = " SELECT * FROM MOVISTO "
     SQLX$ = SQLX$ + "WHERE CDBL(FechMov) >= " & CMAX & " "
     SQLX$ = SQLX$ + "AND (CodiMovi = 'RT' Or CodiMovi = 'DV')"
     '
     Set MoviTemp = Stocks.OpenRecordset(SQLX$, dbOpenSnapshot)
     With MoviTemp
       On Error Resume Next
       .MoveFirst
       If Err < 1 Then
         Do While Not .EOF
           CIXI% = !COD_INTE
           If CIXI% > 0 Then
             If CIXI% <= NUMAS% Then
             AAA = CDbl(!FechMov)
               CAENTREG = !CantSalid - !CantIngre
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
       End If
     End With
     '
'    **** DESCUENTA LOS PEDIDOS PENDIENTES DE ENTREGA
'    **** DESCARTANDO L0S CORRESPONDIENTES A REPARACIONES
     '
'    If MODOPLAN% <> 1 Then
      FIN& = ULTREG&(APROGRA$)
      RFF$ = ""
      For U& = 1 To FIN&
        X$ = REGLEIDO$(APROGRA$, U&)
        RFF$ = RFF$ + Left$(X$, 2)
      Next U&
      '
     Call ABREPEDCLI
     SQLX$ = " SELECT * FROM PEDDETA "
     'SQLX$ = SQLX$ + "WHERE CodiInt = " & CI% & " "
     SQLX$ = SQLX$ + "WHERE (Status >= 0 AND Status < 8) "   ' para que no este anulada ni cumplida ni pendiente de aprobacion
     SQLX$ = SQLX$ + "AND CantEnt < (CanPed - 0.005) "
     SQLX$ = SQLX$ + "AND VenRep = 'V' "
     SQLX$ = SQLX$ + "ORDER BY FeSolEnt DESC "
     '
     Set PEDTEMP = Ventas.OpenRecordset(SQLX$, dbOpenSnapshot)
     '\\\OT-05-0631-WAR-20/09/05///
     INICLICON% = 0
     If UCase$(CONTROL$("CANECOM", "FILPECLI")) <> "SI" Then INICLICON% = 1
     ' SI NO FILTRA POR CLIENTES CON PEDIDOS, ENTONCES TOMA TODOS
     For JJ1% = 1 To 32767 'PONGO EN 0 EL VECTOR
       CLICONPEDID%(JJ1%) = INICLICON%
     Next JJ1%
     FILPEDIPEN% = 0
     If INICLICON% = 1 Then
        For CIXI% = 1 To NUMAS%: ACTICOD%(CIXI%) = 1: Next CIXI%
        GoTo 50
     End If
     '
     ' CARGA UN VECTOR CON LOS CLIENTES CON PEDIDOS PENDIENTES
     FILPEDIPEN% = 1
     For CIXI% = 1 To NUMAS%: ACTICOD%(CIXI%) = 0: Next CIXI%
     With PEDTEMP
       On Error Resume Next
       .MoveFirst
       If Err < 1 Then
         Do While Not .EOF
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
        RASOCLIE2$ = RASOCLI$(JL&)
        Z$ = Space(128)
        Call REPLA(Z$, NCLIE2$, 1, 6)
        Call REPLA(Z$, RASOCLIE2$, 15, 64)
        With CLICOPE04
          .LICOPEN.AddItem Z$
        End With
      End If
     Next JL&
     Screen.MousePointer = 1
     If CLICOPE04.LICOPEN.ListCount > 0 Then
       CLICOPE04.Show 1
       Call FRESHALL
       If CLICOPE04.Label2.Caption <> "OK" Then
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
       For JL& = 1 To CLICOPE04.LICOPEN.ListCount
         If CLICOPE04.LICOPEN.Selected(JL& - 1) = True Then
           ICLIE1& = Val(Left$(CLICOPE04.LICOPEN.List(JL& - 1), 6))
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
     With PEDTEMP
       On Error Resume Next
       .MoveFirst
       If Err < 1 Then
         Do While Not .EOF
           '
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
           CIXI% = !CodiInt
           Call EXPLOMUL(CIXI%)
           For JJ1% = 1 To CONTAMUL%
             CIX2% = CIK%(JJ1%)
             ACTICOD%(CIX2%) = 2
           Next JJ1%
           '
           NCX1% = NCLIE1
           CAPEN = !CanPed - !CantEnt
           If CAPEN > 0 Then
             FENTRESOL% = CVINT%(!FeSolEnt)
             For KK1& = 1 To Len(RFF$) Step 2    ' - 3 To 1 Step -4
               If CVI(Mid$(RFF$, KK1, 2)) = CIXI% Then
                 REPRO& = (KK1& + 1) / 2 ' CVS(Mid$(RFF$, KK1&, 4))
                 ZZ$ = REGLEIDO$(APROGRA$, REPRO&)
                 FEPROG% = CVI(Mid$(ZZ$, 3, 2))
                 If COMIENFE%(FEPROG%) <= COMIENFE%(FENTRESOL%) Then
                   CAPEPROMAX = CVS(Mid$(ZZ$, 7, 4)) - CVS(Mid$(ZZ$, 11, 4)) - CVS(Mid$(ZZ$, 19, 4))
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
'             If CAPEN > 0.05 Then
'               FENTRESO% = FENTRESOL%
'               If FENTRESO% < HOO% Then FENTRESO% = HOO%
'               ZZ$ = REGBLAN$(APROGRA$)
'               Call REPLA(ZZ$, MKI$(CIXI%), 1, 2)
'               Call REPLA(ZZ$, MKI$(FENTRESO%), 3, 2)
'               Call REPLA(ZZ$, MKI$(NCX1%), 5, 2)
'               Call REPLA(ZZ$, MKS$(CAPEN), 19, 4)
'               Call REGAPP(APROGRA$, ZZ$)
'               RFF$ = RFF$ + MKI$(CIXI%)
'                 'Call CIERRARCH(APROGRA$)
'             End If
           End If
           '
103      .MoveNext
         Loop
         '
       End If
     End With
     Call CIERRARCH(APROGRA$)
     '
' hasta aqui esta corregido para entregas / fabricacion
' Y PARA DESCARTAR PEDIDOS POR REPARACIONES
'
' ************** lee existencias, stock minimo y lote reposicion *******************

     FORCANEC04.Check2.Value = 2
     Screen.MousePointer = 11
     '
     ' determina items fanricados con dpi%()=1
     FIN& = NUMAS%
     For I& = 1 To FIN&
       Call TRACE(20, I&, FIN&)
       CI% = I&
       Call AREXPLO(CI%, 1, 3)
       If CAIT% > 0 Then DPI%(CI%) = 1
105  Next I&
     '
     HAYNE% = 0
     modo% = 0
     MODOCAX$ = FORCANEC04.Option3.Caption
     If FORCANEC04.Option1.Value = True Then
         modo% = 1
         MODOCAX$ = FORCANEC04.Option1.Caption
       ElseIf FORCANEC04.Option2.Value = True Then
         modo% = 2
         MODOCAX$ = FORCANEC04.Option2.Caption
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
         If modo% = 1 Then SALSTO = 0
         If modo% = 2 Then If SALSTO < 0 Then SALSTO = 0

         If SALSTO < (-0.05) Then
            HAYNE% = HAYNE% + 1
            CI1% = CI%
            TTXX$ = REGBLAN$("LINEGAT")
            Call REPLA(TTXX$, MKI$(CI%), 1, 2)
            Call REPLA(TTXX$, DESCRIT$(CI%), 3, 30)
            Call REPLA(TTXX$, CSTRING$(MKS$(SALSTO), 4, 10, 1, 2), 33, 10)
            Call REPLA(TTXX$, UNIMED$(CI%), 45, 4)
            '
            QLN& = QLN& + 1
            Call GRAREG("LINEGAT", TTXX$, QLN&)
            '
         End If
         '
         STMIN = STOMIN(CI%): If modo% = 1 Then STMIN = 0
         REPOS = LOTESTAN(CI%)
         DP% = DPI%(CI%): If DP% < 0 Or DP% > 3 Then DP% = 0

         If REPOS < 0.5 Then REPOS = LOSTN(DP%)
         If REPOS < 0.5 Then REPOS = REPOMIN

         DISPACT(CI%) = SALSTO - STMIN
         LOREPOS(CI%) = REPOS
         '
         If Abs(SALSTO) > 0.05 Then
           Call COBERFA(COSEMACT%, CI%, "Stock al " + HO$, 10, SALSTO)
         End If
         '
         If Abs(STMIN) > 0.05 Then
           Call COBERFA(COSEMACT%, CI%, "Stock Minimo del Item", 20, (-1) * STMIN)
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
108  FORCANEC04.Check2.Value = 1
     FORCANEC04.Refresh
     FORCANEC04.Refresh
     '
     '
     ' revisa las ordenes de compra pendientes de recepcion
     Call ABRECOMPRAS
     SQLX$ = " SELECT * FROM OCOMDETA "
     SQLX$ = SQLX$ + "WHERE Cod_Inte > 0 "    ' SOLO MATERIAL CODIFICADO
     SQLX$ = SQLX$ + "AND Stat_Ocom < 8 "   ' para que no este anulada ni cumplida
     SQLX$ = SQLX$ + "AND  Stat_Ocom <> -1 " 'Para que no  hayan O/C's sin aprobar
     SQLX$ = SQLX$ + "AND Cant_Rec < (Cant_Ocom - 0.005) "
     '
     Set OcoTemp = Compras.OpenRecordset(SQLX$, dbOpenSnapshot)
     With OcoTemp
       On Error Resume Next
       .MoveFirst
       If Err < 1 Then
         Do While Not .EOF
           NUORSE& = !N_OrServ
           If NUORSE& < 1 Then
             COEFCAN = !Coef_Unids
             If COEFCAN < 0.005 Then COEFCAN = 1
             CATOT# = !Cant_Ocom * COEFCAN
             CAPRA# = !Cant_Rec * COEFCAN
             CAPEN = CATOT# - CAPRA#
             If CAPEN > 0 Then
               CIBX% = !COD_INTE
               If STATUIT%(CIBX%) >= 0 Then
                 NC% = !Npro_Ocom
                 NroPed& = !Nume_Ocom
                 FENTRE% = CVINT(!Fentre_Sol)
                 COMPROB$ = "OC." + TRIM$(Str$(NroPed&))
                 CICX% = CIBX%
                 Call COBERFA(FENTRE%, CIBX%, COMPROB$ + " " + TEL2CLI$((-1) * NC%), 13, CAPEN)
                 DISPACT(CICX%) = DISPACT(CICX%) + CAPEN
               End If
             End If
           End If
         .MoveNext
         Loop
       End If
     End With
     '
     '
     ' revisa las boletas de recepción pendientes de aprobacion
     Call ABRESTOCKS
     SQLX$ = " SELECT * FROM BOLRECEP "
     SQLX$ = SQLX$ + "WHERE Cod_Inte > 0 "    ' SOLO MATERIAL CODIFICADO
     SQLX$ = SQLX$ + "AND Status = 0 "   ' para que no este anulada ni cumplida
     '
     Set COBRECEPTMP = Stocks.OpenRecordset(SQLX$, dbOpenSnapshot)
     With COBRECEPTMP
       On Error Resume Next
       .MoveFirst
       If Err < 1 Then
         Do While Not .EOF
           CATOT# = !CANTNOM
           CAPRA# = !CanToApro
           CAPEN = CATOT# - CAPRA#
           If CAPEN > 0 Then
             CIBX% = !COD_INTE
             If STATUIT%(CIBX%) >= 0 Then
               NC% = !NumProv
               NroPed& = !NumeBoRe
               FENTRE% = CVINT(!FechRecep)
               COMPROB$ = "BR." + TRIM$(Str$(NroPed&))
               CICX% = CIBX%
               Call COBERFA(FENTRE%, CIBX%, COMPROB$ + " " + TEL2CLI$((-1) * NC%), 11, CAPEN)
               DISPACT(CICX%) = DISPACT(CICX%) + CAPEN
             End If
           End If
         .MoveNext
         Loop
       End If
     End With
     '
' ******** agrega ordenes de fabricacion lanzadas y en curso ******
     '
     Call ABREORFAB
     SQLX$ = " SELECT * FROM ORDEFABR "
     'SQLX$ = SQLX$ + "WHERE Cod_Inte = " & CI% & " "
     SQLX$ = SQLX$ + "WHERE StatuOrf < 8 "   ' para que no este anulada ni cumplida
     SQLX$ = SQLX$ + "AND CanApro < (CanProy - 0.005) "
     'SQLX$ = SQLX$ + "ORDER BY NumeOrFa ASC "
     '
     Set OfaTemp = Ofabric.OpenRecordset(SQLX$, dbOpenSnapshot)
     With OfaTemp
       On Error Resume Next
       .MoveFirst
       If Err < 1 Then
         Do While Not .EOF
           CI% = !COD_INTE
           If STATUIT%(CI%) >= 0 Then
             CACOM = !CanProy
             CAREC = !CanApro
             CAPEN = CACOM - CAREC
             If CAPEN > 0 Then
               NroPed& = !NUMEoRFA
               FEX = CVINT(!FechEnSol)
               REFOR$ = !Referen
               '
               STATU% = !StatuOrf
               If STATU% <= 2 Then ' esta proyectada o lanzada
                 '
                 NORFA$ = !IdSubOr
                 NUORFA& = Val(Mid$(NORFA$, 4, 4))
                 SUBOR% = Val(Mid$(NORFA$, 9, 2))
                 FENTRE% = CVINT(!FechProyEntre) 'CVI(Mid$(F$, 147, 2))
                 If FENTRE% < HOO% Then FENTRE% = !FechEnSol   ' CVI(Mid$(F$, 105, 2))
                 If FENTRE% < HOO% Then FENTRE% = HOO%
                 FELANZ% = CVINT(!FechLan)   'CVI(Mid$(F$, 109, 2))
                 If FELANZ% < HOO% Then FELANZ% = HOO%
                 STAT$ = " (proyectada)"
                 If STATU% = 1 Then STAT$ = " (lanzada)"
                 If STATU% = 2 Then STAT$ = " (suspendida)"
                 '
                 DISPACT(CI%) = DISPACT(CI%) + CAPEN
                 '
                 CIBX% = CI%
                 COD$ = CODEXT$(CI%)
                 COEF = 1
                 '
                 Call COBERFA(FENTRE%, CIBX%, NORFA$ + STAT$, 12, CAPEN)
                 '
               End If
             End If
           End If
         .MoveNext
         Loop
       End If
     End With
     '
     Screen.MousePointer = 11
35   Call ABRESERVA
     SQLX$ = " SELECT * FROM RESERVA "
     SQLX$ = SQLX$ + "WHERE CantCons < (CantRes - 0.005) "
     '
     Set ReserTemp = Stocks.OpenRecordset(SQLX$, dbOpenSnapshot)
     With ReserTemp
       On Error Resume Next
       .MoveFirst
       If Err < 1 Then
         Do While Not .EOF
           CI1% = !COD_INTE
           If CI1% > 0 Then
             TORE# = !CantRes
             CACON# = !CantCons
             '
             If TRIM$(UCase$(UNIMED$(CI1%))) = "GK" Then
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
               NORFA$ = !IdSubOr
               DISPACT(CI1%) = DISPACT(CI1%) - CAPEN
               FELANZ% = CVINT(!FechRes)
               Call COBERFA(FELANZ%, CI1%, "Res.p/" + NORFA$, 22, (-1) * CAPEN)
             End If
           End If
         .MoveNext
         Loop
       End If
     End With
     '
     '
' ******** inicializa necesidades secundarias de compra y fabricacion ********

110  FORCANEC04.Check3.Value = 2
     Screen.MousePointer = 11
     '
111  N13% = FreeFile
     Open LU$ + "NEFACOM.DAT" For Random Access Read Write Shared As #N13% Len = LOREC%
     Dim PDS As String * 208
     Dim PDM As String * 72
     '
     If LOF(N13%) > 0 Then
         Close #N13%
         Kill LU$ + "NEFACOM.DAT"
         GoTo 111
     End If
     '
     For I& = 1 To NMS&
         CIIX& = I&
         Call TRACE(20, CIIX&, NMS&)
         If LOREC% = 208 Then
              Get #N13%, I&, PDS$
            Else
              Get #N13%, I&, PDM$
         End If
         PDS$ = String$(208, 0)
         PDM$ = String$(72, 0)
         If LOREC% = 208 Then
              Put #N13%, I&, PDS$
            Else
              Put #N13%, I&, PDM$
         End If
     Next I&

     FORCANEC04.Check3.Value = 1


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

     FORCANEC04.Check4.Value = 2
     Screen.MousePointer = 11
     '
121  FIN& = NUMAS%
     For I& = 1 To FIN&
       CIIX& = I&
       Call TRACE(20, CIIX&, FIN&)
       CI% = I&
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
122  Next I&
     '
'     If MODOPLAN% = 0 Then
        Call ABREPEDCLI
        SQLX$ = " SELECT * FROM PEDDETA "
        SQLX$ = SQLX$ + "WHERE (Status >= 0 AND Status < 8) "   ' para que no este anulada ni cumplida ni pendiente de aprobacion
        SQLX$ = SQLX$ + "AND CantEnt < (CanPed - 0.005) "
        SQLX$ = SQLX$ + "AND VenRep = 'V' "
        SQLX$ = SQLX$ + "ORDER BY FeSolEnt DESC "
        '
        Set PEDTEMP = Ventas.OpenRecordset(SQLX$, dbOpenSnapshot)
        '
        With PEDTEMP
          On Error Resume Next
          .MoveFirst
          If Err < 1 Then
            Do While Not .EOF
              '
              ' AQUI DESCARTA LOS PEDIDOS DE CLIENTES NO MARCADOS
              NCLIE1 = !NROCLIE
              If Val(NCLIE1) > 0 Then
                If Val(NCLIE1) <= 32767 Then
                  If CLICONPEDID%(NCLIE1) <> 1 Then
                    GoTo 123
                  End If
                End If
              End If
              '
              CIXI% = !CodiInt
              NCX1% = NCLIE1
              CAPEN = !CanPed - !CantEnt
              If CAPEN > 0 Then
                FENTRESOL% = CVINT%(!FeSolEnt)
                X2& = X2& + 1
                B00$ = String$(48, 0)
                Call REPLA(B00$, MKI$(CIXI%), 1, 2)
                Call REPLA(B00$, MKS$(CAPEN), 3, 4)
                Call REPLA(B00$, MKI$(FENTRESOL%), 7, 2)
                  TTXX$ = "O/V " + TRIM$(Str$(!NroPed)) + " - " + RASOCLI$(NCX1%)
                Call REPLA(B00$, TTXX$, 9, 29)
                Call REPLA(B00$, MKS$(0), 38, 4)
                Call REPLA(B00$, MKS$(0), 42, 4)
                Call REPLA(B00$, MKS$(0), 46, 2)
                Call REPLA(B00$, " ", 48, 1)
                '
                Put #N9%, X2&, B00$
                '
              End If
123         .MoveNext
            Loop
          End If
        End With
'     End If
     
     N4% = FreeFile
     Open LU$ + APROGRA$ + ".DAT" For Random Access Read Write Shared As #N4% Len = 24
     Dim P00 As String * 24
     FIN& = LOF(N4%) / 24
     For I& = 1 To FIN&
       CIIX& = I&
       Call TRACE(20, CIIX&, FIN&)
       Get #N4%, I&, P00$
       P1$ = Left$(P00$, 2)
       CI% = CVI(P1$)
       '
       If CI% > 0 Then
         If CI% <= NUMAS% Then
           If STATUIT%(CI%) >= 0 Then
             '
             FEPROGRA% = CVI(Mid$(P00$, 3, 2))
             PLAVEN = CVS(Mid$(P00$, 7, 4))
             ENTREG = CVS(Mid$(P00$, 11, 4))
             PEDIPE = CVS(Mid$(P00$, 19, 4))
             '
             If FEPROGRA% < COMESACT% Then
                cix% = CVI(P1$)
                Call MENSERR(24, "Error: Codigo " + TRIM$(CODEXT$(cix%)) + " Mal Fecha Entrega Programada")
                GoTo 129
             End If
             '
             SALPROGRA = PLAVEN - ENTREG - PEDIPE
             If SALPROGRA >= 0.05 Then
                CI% = CVI(P1$)
                FF% = FEPROGRA%
                '
                If SALPROGRA > 0 Then
                  If CI% > 0 Then
                    If CI% <= NMS& Then
                      DPI%(CI%) = 3
                    End If
                  End If
                End If
                '
                X2& = X2& + 1
                B00$ = String$(48, 0)
                Call REPLA(B00$, MKI$(CI%), 1, 2)
                Call REPLA(B00$, MKS$(CANTU), 3, 4)
                Call REPLA(B00$, MKI$(FF%), 7, 2)
                Call REPLA(B00$, "PROGRAMA NETO ENTREGAS AL " + HO$, 9, 29)
                If MODOPLAN% = 1 Then
                   Call REPLA(B00$, "PLAN NETO FABRICACION AL " + HO$, 9, 29)
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
           End If
         End If
       End If
129  Next I&
     '
     FORCANEC04.Check4.Value = 1
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
     FORCANEC04.Label1 = TRIM$(Str$(CICLO%))
     FORCANEC04.Label1.Refresh
     '
' ******** CALCULANDO NECESIDADES NETAS SEGUN DISPONIBILIDAD
     '
     FORCANEC04.Check5.Value = 2
     Screen.MousePointer = 11
     For I& = X1& To X2&
       '
       CIIX& = I&
       Call TRACE(20, CIIX&, X2&)
       '
       Get #N9%, I&, B00$
       B1$ = Left$(B00$, 2)
       B3$ = Mid$(B00$, 3, 4)
       B4$ = Mid$(B00$, 7, 2)
       '
       CI% = CVI(B1$)
       If CI% > 0 Then
         If CI% <= NMS& Then
           If STATUIT%(CI%) >= 0 Then
             If DISPACT(CI%) >= CVS(B3$) Then
               DISPACT(CI%) = DISPACT(CI%) - CVS(B3$)
               Call REPLA(B00$, B3$, 38, 4)
               Call REPLA(B00$, MKS$(0), 42, 4)
               Put #N9%, I&, B00$
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
             Mid$(PDX$, 4 * UX% - 3, 4) = MKS$(PROANT + CANEC)
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
             DISPACT(CI%) = 0
             Call REPLA(B00$, MKS$(CANEC), 42, 4)
             Call REPLA(B00$, MKI$(FENENE%), 46, 2)
             Put #N9%, I&, B00$
             '
           End If
         End If
       End If
139  Next I&
     '
140  ' Call MENSAJE(14, "GENERANDO NECESIDADES SECUNDARIAS SEGUN EXPLOSION")

     X3& = X2&
     IAGRE& = 0
     For I& = X1& To X2&
       CIIX& = I&
       UKI& = I& - X1& + 1
       Call TRACE(20, UKI&, X2& - X1& + 1)
       Get #N9%, I&, B00$
       Call REPLA(B00$, "C", 48, 1)
       Put #N9%, I&, B00$
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
         FENEC% = CVI(B8$): CUANEC% = 1
                         For UX% = 1 To CACOL%
                            If INIPER%(UX%) <= FENEC% Then CUANEC% = UX%
                         Next UX%
         If CABAS > 0 Then
1234        CIBAS% = CI%
           CII% = CIBAS%
           CODBAS$ = TRIM$(CODEXT$(CII%))
           Call LEEEXPLOMUL(CODBAS$, CABAS, 1)
           If CAIT% > 0 Then
             Call REPLA(B00$, "F", 48, 1)
             Put #N9%, I&, B00$
             For KJ% = 1 To CAIT%    ' J%
               CI% = CIJ%(KJ%)     '  CII1%(KJ%)
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
     Next I&
     '
     For JAGRE& = 1 To IAGRE&
       X3& = X3& + 1
       B00$ = BAGRE$(JAGRE&)
       Put #N9%, X3&, B00$
     Next JAGRE&
     '
     X1& = X2& + 1
     X2& = X3&: If X2& < X1& Then GoTo 150 ' terminar
     GoTo 130
 
150  FORCANEC04.Check5.Value = 1
     FORCANEC04.Check6.Value = 2
     Screen.MousePointer = 11
     '
     For I& = 1 To NMS&
       '
       CIIX& = I&
       Call TRACE(20, CIIX&, NMS&)
       If LOREC% = 208 Then
           Get #N13%, I&, PDS$
           PDX$ = PDS$
         Else
           Get #N13%, I&, PDM$
           PDX$ = PDM$
       End If
       '
       For J% = 1 To CACOL%
         CAFAB = CVS(Mid$(PDX$, 4 * J% - 3, 4))
         If CAFAB > 0 Then
           REFE$ = "O/F a emitir": TIPO% = 15
           If DPI%(I&) = 0 Then REFE$ = "Tentativo de Compras": TIPO% = 14
           UII% = I&
           Call COBERFA(INIPER%(J%), UII%, REFE$, TIPO%, CAFAB)
         End If
       Next J%
     Next I&
     '
     For I& = 1 To X3&
       CIIX& = I&
       Call TRACE(20, CIIX&, X3&)
       Get #N9%, I&, B00$
       B1$ = Left$(B00$, 2)
       B3$ = Mid$(B00$, 3, 4)
       B4$ = Mid$(B00$, 7, 2)
       B5$ = Mid$(B00$, 9, 29)
       B7$ = Mid$(B00$, 42, 4)
       B8$ = Mid$(B00$, 46, 2)
       '
       CI% = CVI(B1$)
       If STATUIT%(CI%) >= 0 Then
         CANEC = (-1) * CVS(B3$)
         FFX% = CVI(B4$)
         REFE$ = B5$
         TIPO% = 22
         If Left$(REFE$, 20) <> "PROGRAMA POR MINIMOS" Then
           If Left$(REFE$, 4) = "PEDI" Then TIPO% = 21
           Call COBERFA(FFX%, CI%, REFE$, TIPO%, CANEC)
         End If
       End If
     Next I&
     '
     FORCANEC04.Check6.Value = 1
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
     FORCANEC04.Check7.Value = 2
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
     FORCANEC04.Check7.Value = 1
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
         PROMATER.Label7.Caption = Left$(FECHOR$, 8): PROMATER.Label7.Refresh
         PROMATER.Label95.Caption = Mid$(FECHOR$, 11, 5): PROMATER.Label95.Refresh
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
     DEPOSTAKK% = Val(TRIM$(CONTROL$("PROMATER", "DEPOSTOK")))
     '
     For U& = 1 To ULTREG&("PLACOMP")
       XX$ = REGLEIDO$("PLACOMP", U&)
       CII% = CVI(Left$(XX$, 2))
       If CII% <> CIAX% Then
         CIAX% = CII%
         STAKK = STODEPOS(CII%, DEPOSTAKK%): If STAKK < 0 Then STAKK = 0
         Call REPLA(XX$, MKS$(STAKK), 29, 4)
         Call REPLA(XX$, MKS$(STOMIN(CII%)), 55, 4)
         Call REPLA(XX$, MKS$(LOTESTAN(CII%)), 59, 4)
         Call GRAREG("PLACOMP", XX$, U&)
       End If
     Next U&
     Call CIERRARCH("PLACOMP")
     Call CARPROSEL
     '
     CIAX% = 0
     For U& = 1 To ULTREG&("PLAFABR")
       XX$ = REGLEIDO$("PLAFABR", U&)
       CII% = CVI(Left$(XX$, 2))
       If CII% <> CIAX% Then
         CIAX% = CII%
         STAKK = STODEPOS(CII%, DEPOSTAKK%): If STAKK < 0 Then STAKK = 0
         Call REPLA(XX$, MKS$(STAKK), 29, 4)
         Call REPLA(XX$, MKS$(STOMIN(CII%)), 55, 4)
         Call REPLA(XX$, MKS$(LOTESTAN(CII%)), 59, 4)
         Call GRAREG("PLAFABR", XX$, U&)
       End If
     Next U&
     Call CIERRARCH("PLAFABR")
     '
999  Call CIERRARCH("*.*")
     Call BAJALDISCO
     Call COMUNI("Re-Cálculo Completo.")
     '
9999 FORCANEC04.Hide
     '

End Sub

Sub CANECOM()         ' ESTE ES EL VIGENTE A PARTIR DE 17 NOV 2007
                      ' SE INTRODUJO FORZAR EL PROCESAMIENTO EN ORDEN DE NIVELES
                      ' DE EXPLOSIÓN POR EL PROBLEMA DE ACM DE AGOLPAR TODA LA
                      ' PROGRAMACION EN LA PRIMERA SEMANA.
                      '
     Dim BAGRE() As String
     Dim CLICONPEDID%(32767)    'Este Vector registra los Clientes con Pedidos Clientes.
     '
     UDIM& = 2048
     ReDim Preserve BAGRE$(UDIM&)
     LU$ = LUDAT$
     HOO% = HOY(HO$)
     CACO& = 0
     '
     ' verifica compras pendientes de aprobación
     Call APERBASDAT("COMPRAS")
     SQLX$ = " SELECT Nume_Ocom, Refe_Ocom,ToBru_Ocom FROM OCOMENCA"
     SQLX$ = SQLX$ + " WHERE Stat_Ocom < 0 "
     SQLX$ = SQLX$ + " ORDER BY Nume_Ocom  ASC "
     Set VEROCTEMP = Compras.OpenRecordset(SQLX$, dbOpenSnapshot)
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
     FORCANEC04.Check1.Value = 2
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
     FORCANEC04.Check1.Value = 1
     '
     'If NESTRU& < 1 Then
     '   Call MENSERR(24, "Proceso Imposible.\ \Faltan Definir Fórmulas y\Estructuras de Producto.")
     '   Call FINAL("")
     'End If
     '
     PLADISTA% = Val(FORCANEC04.Text1.TEXT)
     If PLADISTA% > 6 Then PLADISTA% = 6
     '
'    **** DEPURA PROGRAMA DE ENTREGAS POR ENTREGAS REALIZADAS
     '
     APROGRA$ = "PROGENT"
     'If MODOPLAN% = 1 Then APROGRA$ = "PROGFAB"
     '
     ' ESTO QUE SIGUE ES PARA NO ALTERAR EL PROGRAMA ORIGINAL DE ENTREGAS
     ' SE CONFECCIONA UNA IMAGEN EN PROGCALC Y SE TRABAJA SOBRE ESTA
     Call COPYSTRU(APROGRA$, "PROGCALC")
     JJL1& = 0
     For I& = 1 To ULTREG&(APROGRA$)
       ZZ$ = REGLEIDO$(APROGRA$, I&)
       JJL1& = JJL1& + 1
       Call GRAREG("PROGCALC", ZZ$, JJL1&)
     Next I&
     Call SETULTREG("PROGCALC", JJL1&)
     Call CIERRARCH("PROGCALC")
     Call CIERRARCH(APROGRA$)
     APROGRA$ = "PROGCALC"
     ' FIN CONFECCION DE IMAGEN PROGENT.
     '
     UREPRO& = ULTREG&(APROGRA$)
     JJ& = 0
     For I& = 1 To UREPRO&
       ZZ$ = REGLEIDO$(APROGRA$, I&)
       If CVS(Mid$(ZZ$, 7, 4)) > 0.05 Then
         Call REPLA(ZZ$, MKS$(0), 11, 4)
         Call REPLA(ZZ$, MKS$(0), 19, 4)
         JJ& = JJ& + 1
         Call GRAREG(APROGRA$, ZZ$, JJ&)
       End If
     Next I&
     Call SETULTREG(APROGRA$, JJ&)
     Call CIERRARCH(APROGRA$)
     If JJ& < 1 Then
        On Error Resume Next
        Kill LUDAT$ + APROGRA$ + ".DAT"
        On Error GoTo 0
     End If
     '
     '
' LEER ENTREGAS DEL MES PARA NETEAR PROGRAMA DE ENTREGAS
     CMA% = COMESACT%
     Call ABRESTOCKS
     CMAX = CDbl(CVDAT(CMA%))
     SQLX$ = " SELECT * FROM MOVISTO "
     SQLX$ = SQLX$ + "WHERE CDBL(FechMov) >= " & CMAX & " "
     SQLX$ = SQLX$ + "AND (CodiMovi = 'RT' Or CodiMovi = 'DV')"
     '
     Set MoviTemp = Stocks.OpenRecordset(SQLX$, dbOpenSnapshot)
     With MoviTemp
       On Error Resume Next
       .MoveFirst
       If Err < 1 Then
         Do While Not .EOF
           CIXI% = !COD_INTE
           If CIXI% > 0 Then
             If CIXI% <= NUMAS% Then
             AAA = CDbl(!FechMov)
               CAENTREG = !CantSalid - !CantIngre
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
       End If
     End With
     '
'    **** AGREGA PEDIDOS PENDIENTES DE ENTREGA
'    **** DESCARTANDO L0S CORRESPONDIENTES A REPARACIONES
     '
'    If MODOPLAN% <> 1 Then
      FIN& = ULTREG&(APROGRA$)
      RFF$ = ""
      For U& = 1 To FIN&
        X$ = REGLEIDO$(APROGRA$, U&)
        RFF$ = RFF$ + Left$(X$, 2)
      Next U&
      '
     Call ABREPEDCLI
     SQLX$ = " SELECT * FROM PEDDETA "
     'SQLX$ = SQLX$ + "WHERE CodiInt = " & CI% & " "
     SQLX$ = SQLX$ + "WHERE (Status >= 0 AND Status < 8) "   ' para que no este anulada ni cumplida ni pendiente de aprobacion
     SQLX$ = SQLX$ + "AND CantEnt < (CanPed - 0.005) "
     SQLX$ = SQLX$ + "AND VenRep = 'V' "
     SQLX$ = SQLX$ + "ORDER BY FeSolEnt DESC "
     '
     Set PEDTEMP = Ventas.OpenRecordset(SQLX$, dbOpenSnapshot)
     '\\\OT-05-0631-WAR-20/09/05///
     INICLICON% = 0
     If UCase$(CONTROL$("CANECOM", "FILPECLI")) <> "SI" Then INICLICON% = 1
     ' SI NO FILTRA POR CLIENTES CON PEDIDOS, ENTONCES TOMA TODOS
     For JJ1% = 1 To 32767 'PONGO EN 0 EL VECTOR
       CLICONPEDID%(JJ1%) = INICLICON%
     Next JJ1%
     FILPEDIPEN% = 0
     If INICLICON% = 1 Then
        For CIXI% = 1 To NUMAS%: ACTICOD%(CIXI%) = 1: Next CIXI%
        GoTo 50
     End If
     '
     ' CARGA UN VECTOR CON LOS CLIENTES CON PEDIDOS PENDIENTES
     FILPEDIPEN% = 1
     For CIXI% = 1 To NUMAS%: ACTICOD%(CIXI%) = 0: Next CIXI%
     With PEDTEMP
       On Error Resume Next
       .MoveFirst
       If Err < 1 Then
         Do While Not .EOF
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
        RASOCLIE2$ = RASOCLI$(JL&)
        Z$ = Space(128)
        Call REPLA(Z$, NCLIE2$, 1, 6)
        Call REPLA(Z$, RASOCLIE2$, 15, 64)
        With CLICOPE04
          .LICOPEN.AddItem Z$
        End With
      End If
     Next JL&
     Screen.MousePointer = 1
     If CLICOPE04.LICOPEN.ListCount > 0 Then
       CLICOPE04.Show 1
       Call FRESHALL
       If CLICOPE04.Label2.Caption <> "OK" Then
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
       For JL& = 1 To CLICOPE04.LICOPEN.ListCount
         If CLICOPE04.LICOPEN.Selected(JL& - 1) = True Then
           ICLIE1& = Val(Left$(CLICOPE04.LICOPEN.List(JL& - 1), 6))
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
     With PEDTEMP
       On Error Resume Next
       .MoveFirst
       If Err < 1 Then
         Do While Not .EOF
           '
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
           CIXI% = !CodiInt
           Call EXPLOMUL(CIXI%)
           For JJ1% = 1 To CONTAMUL%
             CIX2% = CIK%(JJ1%)
             ACTICOD%(CIX2%) = 2
           Next JJ1%
           '
           NCX1% = NCLIE1
           CAPEN = !CanPed - !CantEnt
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
     Call CIERRARCH(APROGRA$)
     '
' hasta aqui esta corregido para entregas / fabricacion
' Y PARA DESCARTAR PEDIDOS POR REPARACIONES
'
' ************** lee existencias, stock minimo y lote reposicion *******************

     FORCANEC04.Check2.Value = 2
     Screen.MousePointer = 11
     '
     ' determina items fanricados con dpi%()=1
     FIN& = NUMAS%
     For I& = 1 To FIN&
       Call TRACE(20, I&, FIN&)
       CI% = I&
       Call AREXPLO(CI%, 1, 3)
       If CAIT% > 0 Then DPI%(CI%) = 1
105  Next I&
     '
     Call CALNIVEGEN        ' DETERMINA EL NIVEL DE EXPLOSION GENERAL DE CADA ITEM
     '
     HAYNE% = 0
     modo% = 0
     MODOCAX$ = FORCANEC04.Option3.Caption
     If FORCANEC04.Option1.Value = True Then
         modo% = 1
         MODOCAX$ = FORCANEC04.Option1.Caption
       ElseIf FORCANEC04.Option2.Value = True Then
         modo% = 2
         MODOCAX$ = FORCANEC04.Option2.Caption
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
         If modo% = 1 Then SALSTO = 0
         If modo% = 2 Then If SALSTO < 0 Then SALSTO = 0

         If SALSTO < (-0.05) Then
            HAYNE% = HAYNE% + 1
            CI1% = CI%
            TTXX$ = REGBLAN$("LINEGAT")
            Call REPLA(TTXX$, MKI$(CI%), 1, 2)
            Call REPLA(TTXX$, DESCRIT$(CI%), 3, 30)
            Call REPLA(TTXX$, CSTRING$(MKS$(SALSTO), 4, 10, 1, 2), 33, 10)
            Call REPLA(TTXX$, UNIMED$(CI%), 45, 4)
            '
            QLN& = QLN& + 1
            Call GRAREG("LINEGAT", TTXX$, QLN&)
            '
         End If
         '
         STMIN = STOMIN(CI%): If modo% = 1 Then STMIN = 0
         LINDIVI(CI%) = LINDIVIS(CI%): If LINDIVI(CI%) < 1 Then LINDIVI(CI%) = 1
         REPOS = LOTESTAN(CI%)
         DP% = DPI%(CI%): If DP% < 0 Or DP% > 3 Then DP% = 0

         If REPOS < 0.5 Then REPOS = LOSTN(DP%)
         If REPOS < 0.5 Then REPOS = REPOMIN

         DISPACT(CI%) = SALSTO - STMIN
         LOREPOS(CI%) = REPOS
         '
         If Abs(SALSTO) > 0.05 Then
           Call COBERFA(COSEMACT%, CI%, "Stock al " + HO$, 10, SALSTO)
         End If
         '
         If Abs(STMIN) > 0.05 Then
           Call COBERFA(COSEMACT%, CI%, "Stock Minimo del Item", 20, (-1) * STMIN)
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
108  FORCANEC04.Check2.Value = 1
     FORCANEC04.Refresh
     FORCANEC04.Refresh
     '
     '
     ' revisa las ordenes de compra pendientes de recepcion
     Call ABRECOMPRAS
     SQLX$ = " SELECT * FROM OCOMDETA "
     SQLX$ = SQLX$ + "WHERE Cod_Inte > 0 "    ' SOLO MATERIAL CODIFICADO
     SQLX$ = SQLX$ + "AND Stat_Ocom < 8 "   ' para que no este anulada ni cumplida
     SQLX$ = SQLX$ + "AND  Stat_Ocom <> -1 " 'Para que no  hayan O/C's sin aprobar
     SQLX$ = SQLX$ + "AND Cant_Rec < (Cant_Ocom - 0.005) "
     '
     Set OcoTemp = Compras.OpenRecordset(SQLX$, dbOpenSnapshot)
     With OcoTemp
       On Error Resume Next
       .MoveFirst
       If Err < 1 Then
         Do While Not .EOF
           NUORSE& = !N_OrServ
           If NUORSE& < 1 Then
             COEFCAN = !Coef_Unids
             If COEFCAN < 0.005 Then COEFCAN = 1
             CATOT# = !Cant_Ocom * COEFCAN
             CAPRA# = !Cant_Rec * COEFCAN
             CAPEN = CATOT# - CAPRA#
             If CAPEN > 0 Then
               CIBX% = !COD_INTE
               If STATUIT%(CIBX%) >= 0 Then
                 NC% = !Npro_Ocom
                 NroPed& = !Nume_Ocom
                 FENTRE% = CVINT(!Fentre_Sol)
                 COMPROB$ = "OC." + TRIM$(Str$(NroPed&))
                 CICX% = CIBX%
                 Call COBERFA(FENTRE%, CIBX%, COMPROB$ + " " + TEL2CLI$((-1) * NC%), 13, CAPEN)
                 DISPACT(CICX%) = DISPACT(CICX%) + CAPEN
               End If
             End If
           End If
         .MoveNext
         Loop
       End If
     End With
     '
     ' revisa las boletas de recepción pendientes de aprobacion
     Call ABRESTOCKS
     SQLX$ = " SELECT * FROM BOLRECEP "
     SQLX$ = SQLX$ + "WHERE Cod_Inte > 0 "    ' SOLO MATERIAL CODIFICADO
     SQLX$ = SQLX$ + "AND Status = 0 "   ' para que no este anulada ni cumplida
     '
     Set COBRECEPTMP = Stocks.OpenRecordset(SQLX$, dbOpenSnapshot)
     With COBRECEPTMP
       On Error Resume Next
       .MoveFirst
       If Err < 1 Then
         Do While Not .EOF
           CATOT# = !CANTNOM
           CAPRA# = !CanToApro
           CAPEN = CATOT# - CAPRA#
           If CAPEN > 0 Then
             CIBX% = !COD_INTE
             If STATUIT%(CIBX%) >= 0 Then
               NC% = !NumProv
               NroPed& = !NumeBoRe
               FENTRE% = CVINT(!FechRecep)
               COMPROB$ = "BR." + TRIM$(Str$(NroPed&))
               CICX% = CIBX%
               Call COBERFA(FENTRE%, CIBX%, COMPROB$ + " " + TEL2CLI$((-1) * NC%), 11, CAPEN)
               DISPACT(CICX%) = DISPACT(CICX%) + CAPEN
             End If
           End If
         .MoveNext
         Loop
       End If
     End With
     '
' ******** agrega ordenes de fabricacion lanzadas y en curso ******
     '
     Call ABREORFAB
     SQLX$ = " SELECT * FROM ORDEFABR "
     'SQLX$ = SQLX$ + "WHERE Cod_Inte = " & CI% & " "
     SQLX$ = SQLX$ + "WHERE StatuOrf < 8 "   ' para que no este anulada ni cumplida
     SQLX$ = SQLX$ + "AND CanApro < (CanProy - 0.005) "
     'SQLX$ = SQLX$ + "ORDER BY NumeOrFa ASC "
     '
     Set OfaTemp = Ofabric.OpenRecordset(SQLX$, dbOpenSnapshot)
     With OfaTemp
       On Error Resume Next
       .MoveFirst
       If Err < 1 Then
         Do While Not .EOF
           CI% = !COD_INTE
           If STATUIT%(CI%) >= 0 Then
             CACOM = !CanProy
             CAREC = !CanApro
             CAPEN = CACOM - CAREC
             If CAPEN > 0 Then
               NroPed& = !NUMEoRFA
               FEX = CVINT(!FechEnSol)
               REFOR$ = !Referen
               '
               STATU% = !StatuOrf
               If STATU% <= 2 Then ' esta proyectada o lanzada
                 '
                 NORFA$ = !IdSubOr
                 NUORFA& = Val(Mid$(NORFA$, 4, 4))
                 SUBOR% = Val(Mid$(NORFA$, 9, 2))
                 FENTRE% = CVINT(!FechProyEntre) 'CVI(Mid$(F$, 147, 2))
                 If FENTRE% < HOO% Then FENTRE% = !FechEnSol   ' CVI(Mid$(F$, 105, 2))
                 If FENTRE% < HOO% Then FENTRE% = HOO%
                 FELANZ% = CVINT(!FechLan)   'CVI(Mid$(F$, 109, 2))
                 If FELANZ% < HOO% Then FELANZ% = HOO%
                 STAT$ = " (proyectada)"
                 If STATU% = 1 Then STAT$ = " (lanzada)"
                 If STATU% = 2 Then STAT$ = " (suspendida)"
                 '
                 DISPACT(CI%) = DISPACT(CI%) + CAPEN
                 '
                 CIBX% = CI%
                 COD$ = CODEXT$(CI%)
                 COEF = 1
                 '
                 Call COBERFA(FENTRE%, CIBX%, NORFA$ + STAT$, 12, CAPEN)
                 '
               End If
             End If
           End If
         .MoveNext
         Loop
       End If
     End With
     '
     Screen.MousePointer = 11
35   Call ABRESERVA
     SQLX$ = " SELECT * FROM RESERVA "
     SQLX$ = SQLX$ + "WHERE CantCons < (CantRes - 0.005) "
     '
     Set ReserTemp = Stocks.OpenRecordset(SQLX$, dbOpenSnapshot)
     With ReserTemp
       On Error Resume Next
       .MoveFirst
       If Err < 1 Then
         Do While Not .EOF
           CI1% = !COD_INTE
           If CI1% > 0 Then
             TORE# = !CantRes
             CACON# = !CantCons
             '
             If TRIM$(UCase$(UNIMED$(CI1%))) = "GK" Then
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
               NORFA$ = !IdSubOr
               DISPACT(CI1%) = DISPACT(CI1%) - CAPEN
               FELANZ% = CVINT(!FechRes)
               Call COBERFA(FELANZ%, CI1%, "Res.p/" + NORFA$, 22, (-1) * CAPEN)
             End If
           End If
         .MoveNext
         Loop
       End If
     End With
     '
     '
' ******** inicializa necesidades secundarias de compra y fabricacion ********
110  FORCANEC04.Check3.Value = 2
     Screen.MousePointer = 11
     '
111  N13% = FreeFile
     Open LU$ + "NEFACOM.DAT" For Random Access Read Write Shared As #N13% Len = LOREC%
     Dim PDS As String * 208
     Dim PDM As String * 72
     '
     If LOF(N13%) > 0 Then
         Close #N13%
         Kill LU$ + "NEFACOM.DAT"
         GoTo 111
     End If
     '
     For I& = 1 To NMS&
         CIIX& = I&
         Call TRACE(20, CIIX&, NMS&)
         If LOREC% = 208 Then
              Get #N13%, I&, PDS$
            Else
              Get #N13%, I&, PDM$
         End If
         PDS$ = String$(208, 0)
         PDM$ = String$(72, 0)
         If LOREC% = 208 Then
              Put #N13%, I&, PDS$
            Else
              Put #N13%, I&, PDM$
         End If
     Next I&

     FORCANEC04.Check3.Value = 1

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

     FORCANEC04.Check4.Value = 2
     Screen.MousePointer = 11
     '
121  FIN& = NUMAS%
     For I& = 1 To FIN&
       CIIX& = I&
       Call TRACE(20, CIIX&, FIN&)
       CI% = I&
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
122  Next I&
     '
     N4% = FreeFile
     Open LU$ + APROGRA$ + ".DAT" For Random Access Read Write Shared As #N4% Len = 24
     Dim P00 As String * 24
     FIN& = LOF(N4%) / 24
     For I& = 1 To FIN&
       CIIX& = I&
       Call TRACE(20, CIIX&, FIN&)
       Get #N4%, I&, P00$
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
                cix% = CVI(P1$)
                Call MENSERR(24, "Error: Codigo " + TRIM$(CODEXT$(cix%)) + " Mal Fecha Entrega Programada")
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
129  Next I&
     '
     FORCANEC04.Check4.Value = 1
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
     FORCANEC04.Label1 = TRIM$(Str$(CICLO%))
     FORCANEC04.Label1.Refresh
     '
' ******** CALCULANDO NECESIDADES NETAS SEGUN DISPONIBILIDAD
     '
     FORCANEC04.Check5.Value = 2
     Screen.MousePointer = 11
     For I& = X1& To X2&
       '
       CIIX& = I&
       Call TRACE(20, CIIX&, X2&)
       '
       Get #N9%, I&, B00$
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
               Put #N9%, I&, B00$
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
             'DISPACT(CI%) = 0
             DISPACT(CI%) = CANADI
             'Call REPLA(B00$, MKS$(CANEC), 42, 4)
             Call REPLA(B00$, MKS$(CANEC + CANADI), 42, 4)
             '
             Call REPLA(B00$, MKI$(FENENE%), 46, 2)
             Put #N9%, I&, B00$
             '
           End If
         End If
       End If
139  Next I&
     '
140  ' Call MENSAJE(14, "GENERANDO NECESIDADES SECUNDARIAS SEGUN EXPLOSION")

     X3& = X2&
     IAGRE& = 0
     For I& = X1& To X2&
       CIIX& = I&
       UKI& = I& - X1& + 1
       Call TRACE(20, UKI&, X2& - X1& + 1)
       '
       Get #N9%, I&, B00$
       '
       ' ESTO SE AGREGO PARA FORZAR EL PROCESO POR NIVELES
       KKII% = CVI(Left$(B00$, 2))
       If NIVEGEN%(KKII%) <> CICLO% - 1 Then
          IAGRE& = IAGRE& + 1
          BAGRE$(IAGRE&) = B00$
          B00$ = String$(48, 0)
          Put #N9%, I&, B00$
          GoTo 1236
       End If
       ' **************************************************
       
       Call REPLA(B00$, "C", 48, 1)
       Put #N9%, I&, B00$
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
         FENEC% = CVI(B8$): CUANEC% = 1
                         For UX% = 1 To CACOL%
                            If INIPER%(UX%) <= FENEC% Then CUANEC% = UX%
                         Next UX%
         If CABAS > 0 Then
1234        CIBAS% = CI%
           CII% = CIBAS%
           CODBAS$ = TRIM$(CODEXT$(CII%))
           Call LEEEXPLOMUL(CODBAS$, CABAS, 1)
           If CAIT% > 0 Then
             Call REPLA(B00$, "F", 48, 1)
             Put #N9%, I&, B00$
             For KJ% = 1 To CAIT%    ' J%
               CI% = CIJ%(KJ%)     '  CII1%(KJ%)
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
1236 Next I&
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
 
150  FORCANEC04.Check5.Value = 1
     FORCANEC04.Check6.Value = 2
     Screen.MousePointer = 11
     '
     For I& = 1 To NMS&
       '
       CIIX& = I&
       Call TRACE(20, CIIX&, NMS&)
       If LOREC% = 208 Then
           Get #N13%, I&, PDS$
           PDX$ = PDS$
         Else
           Get #N13%, I&, PDM$
           PDX$ = PDM$
       End If
       '
       For J% = 1 To CACOL%
         CAFAB = CVS(Mid$(PDX$, 4 * J% - 3, 4))
         If CAFAB > 0 Then
           REFE$ = "O/F a emitir": TIPO% = 15
           If DPI%(I&) = 0 Then REFE$ = "Tentativo de Compras": TIPO% = 14
           UII% = I&
           Call COBERFA(INIPER%(J%), UII%, REFE$, TIPO%, CAFAB)
         End If
       Next J%
     Next I&
     '
     For I& = 1 To X3&
       CIIX& = I&
       Call TRACE(20, CIIX&, X3&)
       Get #N9%, I&, B00$
       B1$ = Left$(B00$, 2)
       B3$ = Mid$(B00$, 3, 4)
       B4$ = Mid$(B00$, 7, 2)
       B5$ = Mid$(B00$, 9, 29)
       B7$ = Mid$(B00$, 42, 4)
       B8$ = Mid$(B00$, 46, 2)
       '
       CI% = CVI(B1$)
       If STATUIT%(CI%) >= 0 Then
         CANEC = (-1) * CVS(B3$)
         FFX% = CVI(B4$)
         REFE$ = B5$
         TIPO% = 22
         If Left$(REFE$, 20) <> "PROGRAMA POR MINIMOS" Then
           If Left$(REFE$, 4) = "PEDI" Then TIPO% = 21
           Call COBERFA(FFX%, CI%, REFE$, TIPO%, CANEC)
         End If
       End If
     Next I&
     '
     FORCANEC04.Check6.Value = 1
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
     FORCANEC04.Check7.Value = 2
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
     FORCANEC04.Check7.Value = 1
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
         PROMATER.Label7.Caption = Left$(FECHOR$, 8): PROMATER.Label7.Refresh
         PROMATER.Label95.Caption = Mid$(FECHOR$, 11, 5): PROMATER.Label95.Refresh
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
     DEPOSTAKK% = Val(TRIM$(CONTROL$("PROMATER", "DEPOSTOK")))
     '
     For U& = 1 To ULTREG&("PLACOMP")
       XX$ = REGLEIDO$("PLACOMP", U&)
       CII% = CVI(Left$(XX$, 2))
       If CII% <> CIAX% Then
         CIAX% = CII%
         STAKK = STODEPOS(CII%, DEPOSTAKK%): If STAKK < 0 Then STAKK = 0
         Call REPLA(XX$, MKS$(STAKK), 29, 4)
         Call REPLA(XX$, MKS$(STOMIN(CII%)), 55, 4)
         Call REPLA(XX$, MKS$(LOTESTAN(CII%)), 59, 4)
         Call GRAREG("PLACOMP", XX$, U&)
       End If
     Next U&
     Call CIERRARCH("PLACOMP")
     Call CARPROSEL
     '
     CIAX% = 0
     For U& = 1 To ULTREG&("PLAFABR")
       XX$ = REGLEIDO$("PLAFABR", U&)
       CII% = CVI(Left$(XX$, 2))
       If CII% <> CIAX% Then
         CIAX% = CII%
         STAKK = STODEPOS(CII%, DEPOSTAKK%): If STAKK < 0 Then STAKK = 0
         Call REPLA(XX$, MKS$(STAKK), 29, 4)
         Call REPLA(XX$, MKS$(STOMIN(CII%)), 55, 4)
         Call REPLA(XX$, MKS$(LOTESTAN(CII%)), 59, 4)
         Call GRAREG("PLAFABR", XX$, U&)
       End If
     Next U&
     Call CIERRARCH("PLAFABR")
     '
999  Call CIERRARCH("*.*")
     Call BAJALDISCO
     Call COMUNI("Re-Cálculo Completo.")
     '
9999 FORCANEC04.Hide
     '
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
     For KI% = 1 To NMS%
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
