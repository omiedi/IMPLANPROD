Attribute VB_Name = "RUTSUEL"
'Variables Publicas del Empleado
Public LEGA_LEG%, APNO_LEG$, GRUP_LEG%, COPC_LEG$, ECIV_LEG$
Public CUIL_LEG$, DOCU_LEG$, CATE_LEG$, DCAT_LEG$, TARE_LEG$, SECC_LEG
Public STAT_LEG%, FNAC_LEG%, FEIN_LEG%, FANT_LEG%, FEGR_LEG%
Public CAJU_LEG$, NAFI_LEG$, AFAM_LEG$, DOMI_LEG$, LOCA_LEG$, NCIO_LEG$
Public AFJP_LEG$, TCON_LEG$
'
'Variables Publicas de la Liquidación
Public NRO_LIQ%, DESCRI_LIQ$, MESACU_LIQ%, ANOACU_LIQ%
Public GRUP_LIQ%, TIPO_LIQ%, ASIFLI_LIQ%, FEPA_LIQ%
Public LUPA_LIQ$, PEULDE_LIQ$, FEDE_LIQ%, BANCO_LIQ$
Public PERIO_LIQ$, MESIMPU_LIQ%, ANOIMPU_LIQ%
Public GRULI%, TILIQ%, ASIFLI%
'
'Variables Publicas de Conceptos
Public NUCON%, NURET%, NUASI%, NUREM%, NUPAR%
Public CON$(100), FORCONCE$(100), FORCON%(100), ESTCON%(100), TILIQUI%(100), CARCON%(100), SURETLE%(100), RCLCON%(100), QLCON%(100)
Public RET$(100), FORETLE$(100), FORRET%(100), ESTRET%(100), CARRET%(100), RCLRET%(100), QLRET%(100)
Public ASI$(32), VALASI(32), ESTASI%(32)
'
'Variables Publicas de la Empresa
Public RASO_EMP$, LOCA_EMP$, CAJU_EMP$, FANT_EMP$
Public TELE_EMP$, NUIM_EMP$, DOMI_EMP$, DAT1_EMP$, DAT2_EMP$
'
Public YACARFO%
'
Sub CARGEMPRESA() 'Carga Empresa
  For J& = 1 To ULTREG("DATEMPRE")
    X$ = REGLEIDO$("DATEMPRE", J&)
    Select Case J&
      Case 1
       RASO_EMP$ = TRIM$(Mid$(X$, 1, 30))
       LOCA_EMP$ = TRIM$(Mid$(X$, 31, 24))
       CAJU_EMP$ = TRIM$(Mid$(X$, 55, 16))
      Case 2
       FANT_EMP$ = TRIM$(Mid$(X$, 1, 30))
       TELE_EMP$ = TRIM$(Mid$(X$, 31, 24))
       NUIM_EMP$ = TRIM$(Mid$(X$, 55, 16))
      Case 3
       DOMI_EMP$ = TRIM$(Mid$(X$, 1, 30))
       DAT1_EMP$ = TRIM$(Mid$(X$, 31, 24))
       DAT2_EMP$ = TRIM$(Mid$(X$, 55, 16))
    End Select
  Next J&
  '
End Sub
Sub CARGLIQ(NULIQ1%) 'Carga Liquidación
  '
  NRO_LIQ% = NULIQ1%
  NULI1& = NULIQ1%
  XX$ = REGLEIDO$("ENCLIQUI", NULI1&)
  DESCRI_LIQ$ = Mid$(XX$, 3, 30)
  GRUP_LIQ% = Asc(Mid$(XX$, 33, 1))
  MESACU_LIQ% = Asc(Mid$(XX$, 34, 1))
  ANOACU_LIQ% = Asc(Mid$(XX$, 35, 1))
  TIPO_LIQ% = Asc(Mid$(XX$, 36, 1))
  ASIFLI_LIQ% = Asc(Mid$(XX$, 37, 1))
  FEPA_LIQ% = CVI(Mid$(XX$, 38, 2))
  LUPA_LIQ$ = TRIM$(Mid$(XX$, 40, 20))
  PEULDE_LIQ$ = TRIM$(Mid$(XX$, 60, 16))
  FEDE_LIQ% = CVI(Mid$(XX$, 76, 2))
  BANCO_LIQ$ = TRIM$(Mid$(XX$, 78, 16))
  PERIO_LIQ$ = TRIM$(Mid$(XX$, 94, 12))
  MESIMPU_LIQ% = Asc(Mid$(XX$, 108, 1))
  ANOIMPU_LIQ% = Asc(Mid$(XX$, 109, 1))
  '
End Sub
'
Sub CARGLIQACT()    'Carga Liquidación ACTIVA
  '
  YY$ = REGLEIDO$("ULTILIQ", 1)
  NRO_LIQ% = CVI(Mid$(YY$, 104, 2))
  XX$ = MKI$(NRO_LIQ%) + YY$
  '
  DESCRI_LIQ$ = Mid$(XX$, 3, 30)
  GRUP_LIQ% = Asc(Mid$(XX$, 33, 1))
  MESACU_LIQ% = Asc(Mid$(XX$, 34, 1))
  ANOACU_LIQ% = Asc(Mid$(XX$, 35, 1))
  TIPO_LIQ% = Asc(Mid$(XX$, 36, 1))
  ASIFLI_LIQ% = Asc(Mid$(XX$, 37, 1))
  FEPA_LIQ% = CVI(Mid$(XX$, 38, 2))
  LUPA_LIQ$ = TRIM$(Mid$(XX$, 40, 20))
  PEULDE_LIQ$ = TRIM$(Mid$(XX$, 60, 16))
  FEDE_LIQ% = CVI(Mid$(XX$, 76, 2))
  BANCO_LIQ$ = TRIM$(Mid$(XX$, 78, 16))
  PERIO_LIQ$ = TRIM$(Mid$(XX$, 94, 12))
  MESIMPU_LIQ% = Asc(Mid$(XX$, 108, 1))
  ANOIMPU_LIQ% = Asc(Mid$(XX$, 109, 1))
  '
End Sub
'
Sub CARGEMP(LEGA%) 'Carga Datos del Empleado
  '
  XX$ = FILTERGETRECORD$("PADREMP", 1, MKI$(LEGA%))
  '
  LEGA_LEG% = LEGA%
  APNO_LEG$ = TRIM$(Mid$(XX$, 3, 30))   'APELLIDO Y NOMBRE
  DOCU_LEG$ = Mid$(XX$, 33, 13)         'DOCUMENTO LEGAJO
  FEIN_LEG% = CVI(Mid$(XX$, 46, 2))     'FECHA DE INGRESO
  STAT_LEG% = Asc(Mid$(XX$, 48, 1))     'STATUS
  GRUP_LEG% = Asc(Mid$(XX$, 49, 1))     'GRUPO DE EMPLEADO
  SECC_LEG = CVS(Mid$(XX$, 50, 4))     'SECTOR
  CATE_LEG$ = TRIM$(Mid$(XX$, 54, 9))   'CATEGORIA
  DCAT_LEG$ = TRIM$(ECOARCH$("CATEGO", CATE_LEG$)) ' DESCRIPCION DE LA CATEGORIA
  FNAC_LEG% = CVI(Mid$(XX$, 63, 2))     'FECHA DE NACIMIENTO
  DOMI_LEG$ = TRIM$(Mid$(XX$, 65, 30))  'DOMICILIO
  LOCA_LEG$ = TRIM$(Mid$(XX$, 95, 30))  'LOCALIDAD
  NCIO_LEG$ = TRIM$(Mid$(XX$, 125, 10)) 'NACIONALIDAD
    COECIV% = XVALO(TRIM$(Mid$(XX$, 135, 1)))
  ECIV_LEG$ = TRIM$(ECOARCH("ESTCIV", MKI$(COECIV%)))
  FANT_LEG% = CVI(Mid$(XX$, 145, 2))    'F.ANTIGUEDAD
  FEGR_LEG% = CVI(Mid$(XX$, 147, 2))    'F.EGRESO
  SEXO_LEG$ = TRIM$(Mid$(XX$, 149, 1))  'SEXO
  TARE_LEG$ = Mid$(XX$, 161, 20)        'TAREA
  CAJU_LEG$ = Mid$(XX$, 181, 16)        'CAJA DE JUBILACION
  NAFI_LEG$ = Mid$(XX$, 197, 12)        'NUMERO DE AFILIADO
  CUIL_LEG$ = Mid$(XX$, 209, 13)        'NUMERO DE CUIL
  AFJP_LEG$ = Mid$(XX$, 247, 10)        'NOMBRE DE AFJP$
  TCON_LEG$ = TRIM$(Mid$(FILTERGETRECORD$("TICONTRA", 1, MKI$(LEGA%)), 3))
  '
  RFF$ = RECFILT$("PADREMP", 1, MKI$(LEGA%))
  J1& = CVS(RFF$)
  X1$ = REGLEIDO$("ACTIVA", J1&)
  COPC_LEG$ = Left$(X1$, 96) 'CONCEPTOS OPCIONALES
  AFAM_LEG$ = Mid$(X1$, 97, 32) 'ASIGNACIONES FAMILIARES
  '
End Sub

Sub CARGAFORMULAS()
     '
     If YACARFO% < 1 Then
       GoSub 5000     ' CARGA CONCEPTOS
       GoSub 5050     ' CARGA RETENCIONES
       GoSub 5100     ' CARGA ASIGNACIONES
       GoSub 5150     ' CARGA REMUNERACIONES BASICAS
       YACARFO% = 1
     End If
     Exit Sub
     '
5000 If N1% > 0 Then Close #N1%: N1% = 0
     N1% = FILEOPEN%(LU$ + "CONCEP.DAT", 0, 128)
     Dim CCC0 As String * 128
5007 'If NUCON% > 0 Then GoTo 5032
5010 For I% = 1 To 512 - 1
5015   Get #N1%, I%, CCC0$
       C1$ = Left$(CCC0$, 30)
       C2$ = Mid$(CCC0$, 31, 8)
       C3$ = Mid$(CCC0, 39, 10)
       C4$ = Mid$(CCC0$, 49, 72)
       C5$ = Mid$(CCC0$, 121, 8)
       If Asc(C1$) = 0 Then GoTo 5031
5016   NUCON% = I%
       CON$(I%) = TRIM$(C1$)
       FORMU$ = TRIM$(C4$)
       COD$ = C2$
5017   EST% = Asc(COD$): If EST% < 0 Or EST% > 2 Then EST% = 0
       If Left$(CON$(I%), 3) = "..." Then EST% = 0
5018   TILI% = Asc(Mid$(COD$, 2)): If TILI% < 1 Or TILI% > 2 Then TILI% = 1
5019   FORCA% = Asc(Mid$(COD$, 3)): If FORCA% < 0 Or FORCA% > 4 Then FORCA% = 0
5020   CARA% = Asc(Mid$(COD$, 4)): If CARA% < 0 Or CARA% > 3 Then CARA% = 0
5021   SURET% = Asc(Mid$(COD$, 5)): If SURET% < 0 Or SURET% > 1 Then SURET% = 1
5022   QUILI% = Asc(Mid$(COD$, 6)): If QUILI% < 0 Or QUILI% > 2 Then QUILI% = 0
5023   RCLX% = Asc(Mid$(COD$, 7)): If RCLX% < 0 Or RCLX% > 1 Then RCLX% = 1
5024   If FORMU$ = "" Then FORMU$ = "0"
5027   '
5028   FORCONCE$(I%) = FORMU$
       ESTCON%(I%) = EST%
       TILIQUI%(I%) = TILI%
       CARCON%(I%) = CARA%
       FORCON%(I%) = FORCA%
       SURETLE%(I%) = SURET%
       QLCON%(I%) = QUILI%
       RCLCON%(I%) = RCLX%: If FORCA% = 1 Then RCLCON%(I%) = 0
5030 Next I%
     '
5031 '
5032 If N1% > 0 Then Close #N1%: N1% = 0
     Return
     ' **********************************************************************
     '
5050 If N1% > 0 Then Close #N1%: N1% = 0
     N1% = FILEOPEN%(LU$ + "RETLEG.DAT", 0, 128)
     Dim CCC1 As String * 128
5057 'If NURET% > 0 Then GoTo 5084
5060 For I% = 1 To 512 - 1
5065   Get #N1%, I%, CCC1$
       C1$ = Left$(CCC1$, 30)
       C2$ = Mid$(CCC1$, 31, 8)
       C3$ = Mid$(CCC1$, 39, 10)
       C4$ = Mid$(CCC1$, 49, 72)
       C5$ = Mid$(CCC1$, 121, 8)
       If Asc(C1$) = 0 Then GoTo 5084
       '
5067   NURET% = I%
       RET$(I%) = TRIM$(C1$)
       FORMU$ = TRIM$(C4$)
       COD$ = C2$
5068   EST% = Asc(COD$): If EST% < 0 Or EST% > 2 Then EST% = 0
       If Left$(RET$(I%), 3) = "..." Then EST% = 0
5069   CARA% = Asc(Mid$(COD$, 3)): If CARA% < 0 Or CARA% > 3 Then CARA% = 0
5070   FORCA% = Asc(Mid$(COD$, 4)): If FORCA% < 0 Or FORCA% > 1 Then FORCA% = 0
5071   QUILI% = Asc(Mid$(COD$, 6)): If QUILI% < 0 Or QUILI% > 2 Then QUILI% = 0
5072   RCLX% = Asc(Mid$(COD$, 7)): If RCLX% < 0 Or RCLX% > 1 Then RCLX% = 1
5076   If FORMU$ = "" Then FORMU$ = "0"
5077   '
5078   FORETLE$(I%) = FORMU$
       ESTRET%(I%) = EST%
       CARRET%(I%) = CARA%
       FORRET%(I%) = FORCA%
       QLRET%(I%) = QUILI%
       RCLRET%(I%) = RCLX%
       If FORCA% = 1 Then
         If RCLX% <> 2 Then
           RCLRET%(I%) = 0
         End If
       End If
5079   If I% = 16 Then If InStr(C1$, "ANTIC") > 0 Then RCLRET%(I%) = 2
5080   If I% = 17 Then If InStr(C1$, "PREST") > 0 Then RCLRET%(I%) = 2
5082 Next I%
     '
5084 If N1% > 0 Then Close #N1%: N1% = 0
     Return
     '************************************************************************
     '
5100 If N1% > 0 Then Close #N1%: N1% = 0
     N1% = FILEOPEN%(LU$ + "ASIGNA.DAT", 0, 48)
     Dim CCC2 As String * 48
5105 'If NUASI% > 0 Then GoTo 5135
5110 For I% = 1 To 32
5115   Get #N1%, I%, CCC2$
       C1$ = Left$(CCC2$, 30)
       C2$ = Mid$(CCC2$, 31, 8)
       C3$ = Mid$(CCC2$, 39, 4)
       C4$ = Mid$(CCC2$, 43, 2)
       C5$ = Mid$(CCC2$, 45, 4)
       If Asc(C1$) < 1 Then GoTo 5135
       '
5120   NUASI% = I%
       ASI$(I%) = C1$
       VALASI(I%) = CVS(C3$)
       ESTASI%(I%) = Asc(C4$)
       If Left$(ASI$(I%), 3) = "..." Then ESTASI%(I%) = 0
5125 Next I%
5130 '
5135 Return
     '************************************************************************
     '
     
5150 If N4% > 0 Then Close #N4%: N4% = 0
     N4% = FILEOPEN(LU$ + "NOMREM.DAT", 0, 128)
     Dim RRR0 As String * 128
5155 For I% = 1 To 12
5157   Get #N4%, I%, RRR0$
       R1$ = Left$(RRR0$, 30)
       R2$ = Mid$(RRR0$, 31, 8)
       R3$ = Mid$(RRR0$, 39, 1)
       R4$ = Mid$(RRR0$, 40, 3)
       If Asc(R1$) < 1 Then GoTo 5165
5160   NUREM% = I%
5162 Next I%
5163 '
5165 If N4% > 0 Then Close #N4%: N4% = 0
5167 '
5170 If N4% > 0 Then Close #N4%: N4% = 0
     N4% = FILEOPEN%(LU$ + "PARAME.DAT", 0, 64)
     Dim RRR1 As String * 64
5175 For I% = 1 To 20
5177   Get #N4%, I%, RRR1$
       R1$ = Left$(RRR1$, 30)
       R2$ = Mid$(RRR1$, 31, 8)
       R3$ = Mid$(RRR1$, 39, 1)
       R4$ = Mid$(RRR1$, 40, 3)
       If Asc(R1$) < 1 Then GoTo 5185
5180   NUPAR% = I%
5182 Next I%
5183 '
5185 If N4% > 0 Then Close #N4%: N4% = 0
     Return
     '************************************************************************

End Sub

Sub LIQUIORDI()
     '
10 ' dimensionamiento de vectores
11   Dim CTB$(16), DEX$(16), CAN(16), IMPO#(16), SIG%(16)
12   Dim CTA%(128), CANTJ(128), IM#(128), INFO%(128)
13   Dim FERI%(20)
     Dim GRU$(32), FRELIQ%(32)
14   Dim CON$(512), FORCON$(512), FORCONI%(512), ESTCON%(512), TILJ%(512), CARCON%(512), SURETI%(512), RCLCON%(512), QLCON%(512)
15   Dim RET$(512), FORRET$(512), FORRETI%(512), ESTRET%(512), CARRET%(512), RCLRET%(512), QLRET%(512)
16   Dim ASI$(32), VALASI(32), ESTASI%(32)
17   Dim SB(512), QT(512), RMD#(512), RT#(512), PII(512), RI#(512)
18   Dim LEGAI%(4196), ACMEN#(4196), RTGAN#(4196), PULIQ$(4196)
     '
     Dim VALESC(32)
     
     RELIMI# = 1000 ' LIMITE PARA EL CAMBIO PORCENTAJE ANSAL

29   '
50   QWERTY% = USNBR%: LU$ = LUDAT$
52   '
60   HOII% = HOY(HO$)
65   RTAN% = 16: RTPR% = 17     ' ANTICIPO ES RETENCION NRO 16, PRESTAMO 17
66   RFER% = 10                 ' PARA LODIGIANI FERIADO ES CONCEPTO NRO 10
68   RTGN% = 18                 ' PARA LODIGIANI GANANCIAS ES RETENCION NRO 18
69   '
70 ' caracteres de control
71   S1$ = "PROCESO DE LIQUIDACION"
     If EMPRE$ <> "" Then S1$ = S1$ + " - " + EMPRE$
72   S3$ = "Un momento, por favor ..."
77   'WIDTH "SCRN:",255
78   VI% = 1: PH% = 1
79   PRIMOD% = 0
90   '
      '
150   'GoSub 31000: VDSS$ = "6\6\23\0\8\29\4\8\12\12\<Codigo\<Descripcion\^%\^Cant.\^Haberes\^Deducc.\": RLB% = 11
151   'GoSub 15900: Call TOOLBAR: GoSub 15950: Call TOOLBAR
160   '
200   GoSub 5900: GoSub 5930     ' LEER DATOS LIQUIDACION ACTIVA Y CALEND. FERIADOS
201   GoSub 5400                 ' LEER GRUPOS DE EMPLEADOS
202   GoSub 5000                 ' LEER CONCEPTOS DEFINIDOS
203   GoSub 5050                 ' LEER RETENCIONES DEFINIDAS
204   GoSub 5100                 ' LEER ASIGNACIONES FAMILIARES
        Call CIERRARCH("ECONCEP")
206   GoSub 5150                 ' LEER NUREM% - NRO. DE SALARIOS BASICOS
      '
500   GoSub 5200                 ' ABRIR LIQUIDACION ACTIVA EN #N6%
510   HORE% = 0
      BARRE% = 0
      If GRULI% > 0 Then
        If MAXDIA% = 0 Then
          BARRE% = 1
          EPL% = 0
        End If
      End If
      '
530   If BARRE% = 1 Then GoSub 5600                       ' ABRIR RELOJ
600   If MAXDIA% = 0 Then
          GoTo 1000
        Else
          GoSub 8300
      End If
Stop


610 ''LOCATE 12, 1: Print LI0$
611 ''LOCATE 13, 1: Print LI0$
612 ''LOCATE 14, 1: Print LI0$
615 If EJU% < 0 Then GoTo 2000
620 '
1000 'Procesamiento de liquidacion activa
1001 GoSub 19000: CONTA% = 0
1002 GoSub 5500: GoSub 5550 ' ABRIR PADRON EN #N5% Y SALARIOS BASICOS EN #N7%
1004 GoSub 5000: GoSub 5200                         ' ABRIR ARCHIVOS
1005 ''Call VENTANA(0, 0, 0, 0, "", "CLEAR")
     GoSub 15000
1006 '
1007 ''Call BLIMESS(" ESPERE  !!! ")
1008 GoSub 17000  ' AGREGAR CONCEPTOS MANDATORIOS DE CALCULO
1010 '
1012 FEX = COMEJ%: CMJ$ = FECHATEX$(FEX)
     FEX = FF%: FFS$ = FECHATEX$(FEX): COLMES% = Val(Mid$(FFS$, 4, 2)) - Val(Mid$(CMJ$, 4, 2)) + 2: If COLMES% < 2 Then COLMES% = COLMES% + 12
1014 If NRO = 1 Then COLMES% = 1
1015 If LEGA% > ULTLEGA% Then GoSub 18500: GoTo 1050
1020 If INTERCA% = 0 Then GoSub 18300: GoTo 1050    ' ACTUALIZAR MODIFICACIONES
1030 If INTERCA% = 1 Then GoSub 18600: GoTo 1050    ' INTERCALAR ASIENTO
1040 '
1050 If BARRE% = 1 Then GoSub 5200: GoSub 19050: GoTo 1005
1080 '
1090 GoSub 5200: GoSub 19050: GoTo 1000
1095 '
1100                         ' INDEXA ARCHIVO DE DIARIO
1110 GoTo 3900
1120 '
2000 ''Call BLIMESS(" ESPERE  !!! ")
2001 GoSub 19000: CONTA% = 0
2002 GoSub 5500: GoSub 5550 ' ABRIR PADRON EN #N5% Y SALARIOS BASICOS EN #N7%
2005 EPL% = 0
2010 GoSub 20000
2015 GoSub 19000
2020 GoSub 8000
2025 GoSub 8100
2030 GoSub 17000
2035 If LEGA% > ULTLEGA% Then GoSub 18500: GoTo 2050
2040 If INTERCA% = 0 Then GoSub 18300: GoTo 2050    ' ACTUALIZAR MODIFICACIONES
2045 If INTERCA% = 1 Then GoSub 18600: GoTo 2050    ' INTERCALAR ASIENTO
2047 '
2050 GoSub 5200: GoTo 2010
2080 '
2092 GoTo 3900
2095 '
3900 Call FINAL("")
3920 '
5000 If N1% > 0 Then Close #N1%: N1% = 0
     N1% = FILEOPEN%(LU$ + "CONCEP.DAT", 0, 128)
     Dim CCC0 As String * 128
5007 If NUCON% > 0 Then GoTo 5032
5010 For I% = 1 To 512 - 1
5015   Get #N1%, I%, CCC0$
       C1$ = Left$(CCC0$, 30)
       C2$ = Mid$(CCC0$, 31, 8)
       C3$ = Mid$(CCC0, 39, 10)
       C4$ = Mid$(CCC0$, 49, 72)
       C5$ = Mid$(CCC0$, 121, 8)
       If Asc(C1$) = 0 Then GoTo 5031
5016   NUCON% = I%
       CON$(I%) = TRIM$(C1$)
       FORMU$ = TRIM$(C4$)
       COD$ = C2$
5017   EST% = Asc(COD$): If EST% < 0 Or EST% > 2 Then EST% = 0
       If Left$(CON$(I%), 3) = "..." Then EST% = 0
5018   TILI% = Asc(Mid$(COD$, 2)): If TILI% < 1 Or TILI% > 2 Then TILI% = 1
5019   CARA% = Asc(Mid$(COD$, 3)): If CARA% < 0 Or CARA% > 3 Then CARA% = 0
5020   FORCA% = Asc(Mid$(COD$, 4)): If FORCA% < 0 Or FORCA% > 4 Then FORCA% = 0
5021   SURET% = Asc(Mid$(COD$, 5)): If SURET% < 0 Or SURET% > 1 Then SURET% = 1
5022   RCLX% = Asc(Mid$(COD$, 7)): If RCLX% < 0 Or RCLX% > 1 Then RCLX% = 1
5023   QUILI% = Asc(Mid$(COD$, 6)): If QUILI% < 0 Or QUILI% > 2 Then QUILI% = 0
5024   If FORMU$ = "" Then FORMU$ = "0"
5027   '
5028   FORCON$(I%) = FORMU$
       ESTCON%(I%) = EST%
       TILJ%(I%) = TILI%
       CARCON%(I%) = CARA%
       FORCONI%(I%) = FORCA%
       SURETI%(I%) = SURET%
       QLCON%(I%) = QUILI%
       RCLCON%(I%) = RCL%: If FORCA% = 1 Then RCLCON%(I%) = 0
5030 Next I%
     '
5031 '
5032 If N1% > 0 Then Close #N1%: N1% = 0
     Return
     '
     '
5050 If N1% > 0 Then Close #N1%: N1% = 0
     N1% = FILEOPEN%(LU$ + "RETLEG.DAT", 0, 128)
     Dim CCC1 As String * 128
5057 If NURET% > 0 Then GoTo 5084
5060 For I% = 1 To 512 - 1
5065   Get #N1%, I%, CCC1$
       C1$ = Left$(CCC1$, 30)
       C2$ = Mid$(CCC1$, 31, 8)
       C3$ = Mid$(CCC1$, 39, 10)
       C4$ = Mid$(CCC1$, 49, 72)
       C5$ = Mid$(CCC1$, 121, 8)
       If Asc(C1$) = 0 Then GoTo 5083
       '
5067   NURET% = I%
       RET$(I%) = TRIM$(C1$)
       FORMU$ = TRIM$(C4$)
       COD$ = C2$
5068   EST% = Asc(COD$): If EST% < 0 Or EST% > 2 Then EST% = 0
       If Left$(RET$(I%), 3) = "..." Then EST% = 0
5069   CARA% = Asc(Mid$(COD$, 3)): If CARA% < 0 Or CARA% > 3 Then CARA% = 0
5070   FORCA% = Asc(Mid$(COD$, 4)): If FORCA% < 0 Or FORCA% > 1 Then FORCA% = 0
5071   QUILI% = Asc(Mid$(COD$, 6)): If QUILI% < 0 Or QUILI% > 2 Then QUILI% = 0
5072   RCLX% = Asc(Mid$(COD$, 7)): If RCLX% < 0 Or RCLX% > 1 Then RCLX% = 1
5076   If FORMU$ = "" Then FORMU$ = "0"
5077   '
5078   FORRET$(I%) = FORMU$
       ESTRET%(I%) = EST%
       CARRET%(I%) = CARA%
       FORRETI%(I%) = FORCA%
       QLRET%(I%) = QUILI%
       RCLRET%(I%) = RCL%
       If FORCA% = 1 Then
         If RCLX% <> 2 Then
           RCLRET%(I%) = 0
         End If
       End If
5079   If I% = 16 Then If InStr(C1$, "ANTIC") > 0 Then RCLRET%(I%) = 2
5080   If I% = 17 Then If InStr(C1$, "PREST") > 0 Then RCLRET%(I%) = 2
5082 Next I%
     '
5083 '
5084 If N1% > 0 Then Close #N1%: N1% = 0
     Return
     '
     '
5100 If N1% > 0 Then Close #N1%: N1% = 0
     N1% = FILEOPEN%(LU$ + "ASIGNA.DAT", 0, 48)
     Dim CCC2 As String * 48
5105 If NUASI% > 0 Then GoTo 5135
5110 For I% = 1 To 32
5115   Get #N1%, I%, CCC2$
       C1$ = Left$(CCC2$, 30)
       C2$ = Mid$(CCC2$, 31, 8)
       C3$ = Mid$(CCC2$, 39, 4)
       C4$ = Mid$(CCC2$, 43, 2)
       C5$ = Mid$(CCC2$, 45, 4)
       If Asc(C1$) < 1 Then GoTo 5132
       '
5120   NUASI% = I%
       ASI$(I%) = C1$
       VALASI(I%) = CVS(C3$)
       ESTASI%(I%) = Asc(C4$)
       If Left$(ASI$(I%), 3) = "..." Then ESTASI%(I%) = 0
5125 Next I%
5130 '
5132 '
5135 Return
     '
     '
     
5150 If N4% > 0 Then Close #N4%: N4% = 0
     N4% = FILEOPEN(LU$ + "NOMREM.DAT", 0, 128)
     Dim RRR0 As String * 128
5155 For I% = 1 To 12
5157   Get #N4%, I%, RRR0$
       R1$ = Left$(RRR0$, 30)
       R2$ = Mid$(RRR0$, 31, 8)
       R3$ = Mid$(RRR0$, 39, 1)
       R4$ = Mid$(RRR0$, 40, 3)
       If Asc(R1$) < 1 Then GoTo 5165
5160   NUREM% = I%
5162 Next I%
5163 '
5165 If N4% > 0 Then Close #N4%: N4% = 0
5167 '
5170 If N4% > 0 Then Close #N4%: N4% = 0
     N4% = FILEOPEN%(LU$ + "PARAME.DAT", 0, 64)
     Dim RRR1 As String * 64
5175 For I% = 1 To 20
5177   Get #N4%, I%, RRR1$
       R1$ = Left$(RRR1$, 30)
       R2$ = Mid$(RRR1$, 31, 8)
       R3$ = Mid$(RRR1$, 39, 1)
       R4$ = Mid$(RRR1$, 40, 3)
       If Asc(R1$) < 1 Then GoTo 5185
5180   NUPAR% = I%
5182 Next I%
5183 '
5185 If N4% > 0 Then Close #N4%: N4% = 0
     Return
     '
     '
5200 If N6% > 0 Then Close #N6%: N6% = 0
     N6% = FILEOPEN(LU$ + "DETLIQUI.DAT", 0, 16)
     Dim DDD0 As String * 16
5205 ''FIELD #N6%,2 AS D1$,2 AS D2$,4 AS D3$,8 AS D4$
5210 For I% = LOF(N6%) / 16 To 1 Step -1
5215   Get #N6%, I%, DDD0$
       D1$ = Left$(DDD0$, 2)
       D2$ = Mid$(DDD0$, 3, 2)
       D3$ = Mid$(DDD0$, 5, 4)
       D4$ = Mid$(DDD0$, 9, 8)
       If CVI(D1$) > 0 Then
         MAXDIA% = I%
         ULTLEGA% = CVI(D1$)
         GoTo 5250
       End If
5220 Next I%
5225 MAXDIA% = 0
     ULTLEGA% = 0
     '
     ' ANTICIPOS Y PRESTAMOS QUEDA PARA DESPUES
5250 ''If N13% > 0 Then Close #N13% Else N13% = FreeFile
      ''  OPEN "R",#N13%,LU$+"ANTIPEN.DAT",16
5251 ''FIELD #N13%,2 AS V1$,2 AS V2$,8 AS V3$,1 AS V4$,3 AS LIBRE$
5253 ''CANANT% = 0: For I% = 1 To LOF(N13%) / 16
5258 ''GET #N13%,I%:IF CVI(V1$)=0 THEN 5270
5262 ''CANANT% = I%
5266 ''Next I%
5268 '
5270 ''If N14% > 0 Then Close #N14% Else N14% = FreeFile
     ''   OPEN "R",#N14%,LU$+"PRESTAMO.DAT",32
5271 ''FIELD #N14%,2 AS W1$,2 AS W2$,8 AS W3$,2 AS W4$,8 AS W5$,2 AS W6$,1 AS W7$,7 AS W8$
5273 ''CAPRES% = 0
5277 ''For I% = 1 To LOF(N14%) / 32
5278 ''GET #N14%,I%
5280 ''If CVI(W1$) = 0 Then GoTo 5288
5282 ''CAPRES% = I%
5284 ''Next I%
5286 '
5288 Return
     '
     '
5400 If N1% > 0 Then Close #N1% Else N1% = FreeFile
     ''   OPEN "R",#N1%,LU$+"DATGREMP.DAT",64
5405 ''FIELD #N1%,25 AS C1$,5 AS C2$,30 AS C3$,1 AS C4$,1 AS C5$
5410 For I% = 1 To 10
5415 ''GET #N1%,I%:IF ASC(C1$)=0 THEN 5432
5416 NUGRU% = I%: GRU$(I%) = C1$: FRELIQ%(I%) = Asc(C5$)
5417 Next I%
5420 '
5432 If N1% > 0 Then Close #N1%: N1% = 0
        Return
5433 '
5500 If N5% > 0 Then Close #N5%: N5% = 0
     N5% = FILEOPEN(LU$ + "PADRON.DAT", 0, 64)
     Dim F000 As String * 64
5505 ''FIELD #N5%,2 AS F1$,30 AS F2$,13 AS F3$,2 AS F4$,1 AS F5$,1 AS F6$,4 AS F7$,9 AS F8$,2 AS F9$
5507 If NUEMPL% > 0 Then GoTo 5530
5508 '
5510 For I% = 1 To LOF(N5%) / 64
5515    Get #N5%, I%, FFF0$
        F1$ = Left$(FFF0$, 2)
        If CVI(F1$) = 0 Then GoTo 5527
        If I% > 4196 Then
          Call MENSERR(24, "Vector de Empleados fuera de Rango - Consulte")
          GoTo 5527
        End If
5517    NUEMPL% = I%
        LEGAI%(I%) = CVI(F1$)
5520 Next I%
     '
5527 GoSub 5960
     GoSub 7600
     If N10% > 0 Then Close #N10%: N10% = 0
     '
5530 If N2% > 0 Then Close #N2%: N2 = 0
     N2% = FILEOPEN(LU$ + "PADAUX.DAT", 0, 96)
     Dim PDX0 As String * 96
5532 ''FIELD #N2%,30 AS F11$,30 AS F12$,10 AS F13$,10 AS F14$,2 AS F15$,2 AS F16$,1 AS F17$,1 AS F18$,10 AS F19$
5535 Return
     '
     '
5550 If N7% > 0 Then Close #N7%: N7% = 0
     N7% = FILEOPEN(LU$ + "REMBAS.DAT", 0, 128)
     Dim RBB0 As String * 128
5555 ''FIELD #N7%,48 AS RB$,80 AS PI$
5557 '
5560 If N9% > 0 Then Close #N9%: N9% = 0
     N9% = FILEOPEN(LU$ + "ACTIVA.DAT", 0, 128)
5565 ''FIELD #N9%,96 AS AC1$,32 AS AC2$
     Dim ACC0 As String * 128
5570 Return
     '
     '
5600 '' las horas quedan para mas adelante
     HORE% = 0
5602 ''If N11% > 0 Then Close #N11% Else N11% = FreeFile
     ''   OPEN "R",#N11%,LU$+"HORELOJ.DAT",16
5605 ''FIELD #N11%,2 AS H1$,2 AS H3$,4 AS H5$
5610 ''If LOF(N11%) < 16 Then
 ''If N11% > 0 Then Close #N11%: N11% = 0
 ''Kill LU$ + "HORELOJ.DAT"
 ''GoTo 5690
     ''End If
5615 ''HORE% = LOF(N11%) / 16
5690 Return
5695 '
5900 If N1% > 0 Then Close #N1%: N1% = 0
     N1% = FILEOPEN%(LU$ + "ULTILIQ.DAT", 0, 128)
5905 'FIELD #N1%,30 AS C1$,1 AS C2$,1 AS C3$,1 AS C4$,1 AS C5$,1 AS C6$,2 AS C7$,20 AS C8$,16 AS C9$,2 AS C10$,16 AS C11$,12 AS C12$,2 AS C13$,1 AS C14$,1 AS C15$,2 AS C16$,19 AS C17$
     Dim C0 As String * 128
5910 Get #N1%, 1, C0$: GoSub 25001
     LIQ$ = C1$: GRULI% = Asc(C2$): TILIQ% = Asc(C5$): QUASI% = Asc(C6$): MES% = Asc(C3$): ANJ% = Asc(C4$): MARCA% = CVI(C13$): LSUE% = CVI(C16$)
     If GRULI% = 0 Then DES% = 10: HAS% = 0: GoTo 5919
5919 If N1% > 0 Then Close #N1%: N1% = 0
     Return
     '
5930 If N1% > 0 Then Close #N1%: N1% = 0
     N1% = FILEOPEN%(LU$ + "FERIADOS.DAT", 0, 64)
     Dim FRD0 As String * 64
5937 CAFER% = 0
5940 For I% = 1 To 20
5945   Get #N1%, I%, FRD0$
         F1$ = Left$(FRD0$, 30)
         F2$ = Mid$(FRD0$, 31, 2)
         F3$ = Mid$(FRD0$, 33, 28)
         F4$ = Mid$(FRD0$, 61, 4)
         If Asc(F1$) = 0 Then GoTo 5955
5947     If CVI(F2$) >= DES% Then
           If CVI(F2$) <= HAS% Then
             CAFER% = CAFER% + 1
           End If
         End If
5950 Next I%
5955 If N1% > 0 Then Close #N1%: N1% = 0
     Return
     '
     '
5960 MESI$ = Mid$(Str$(MES%), 2): While Len(MESI$) < 2: MESI$ = "0" + MESI$: Wend
5961 ANI$ = Mid$(Str$(ANJ%), 2): While Len(ANI$) < 2: ANI$ = "0" + ANI$: Wend
5962 Archi$ = "DTL" + MESI$ + ANI$ + ".DAT"
5970 If N10% > 0 Then Close #N10%: N10% = 0
     N10% = FILEOPEN(LU$ + Archi$, 0, 24)
     Dim JJJ0 As String * 24
5971      ''FIELD #N10%,2 AS J0$,2 AS J1$,2 AS J2$,4 AS J3$,8 AS J4$,6 AS LBR$
5977 '
5980 Return
     '
     '
7600 For III% = 1 To LOF(N10%) / 24
7602 ''GET #N10%,III%
7603 If CVI(J0$) = 0 Then GoTo 7697
7604 REMGRA# = 0: RETGAN# = 0: If CVI(J0$) = MARCA% Then GoTo 7695 ' ES LIQUIDACION ACTIVA
7605 LGJ% = CVI(J1$): CODI% = CVI(J2$): IMPORT# = CVD(J4$)
     If (CODI% < 1 Or CODI% > 199) Then GoTo 7609
     For IJK% = 1 To NUEMPL%
       If LGJ% = LEGAI%(IJK%) Then GoTo 7606
     Next IJK%
     GoTo 7695    ' NO ENCONTRO LEGAJO
     '
7606 CODIG% = CODI% Mod 100
     PULIQ$(IJK%) = PULIQ$(IJK%) + MKI$(III%)
     '
     If SURETI%(CODIG%) = 0 Then GoTo 7695 ' EXENTO
7607 REMGRA# = IMPORT#: GoTo 7615
7608 '
7609 If CODI% <> 200 + RTGN% Then GoTo 7695
7610 RETGAN# = IMPORT#
7612 '
7615 For IJK% = 1 To NUEMPL%
7620 If LGJ% = LEGAI%(IJK%) Then ACMEN#(IJK%) = ACMEN#(IJK%) + REMGRA#: RTGAN#(IJK%) = RTGAN#(IJK%) + RETGAN#: GoTo 7695
7625 Next IJK%
7630 '
7695 Next III%
7696 '
7697 Return
7698 '
8000 ENCU% = 0: CONTANT% = 0: ERRO% = 0: LI% = 0: LS% = MAXDIA% + 1
8010 E% = Int((LS% - LI%) / 2): If E% < 1 Then GoTo 8059
8020 L% = LI% + E%: If L% < 1 Or L% > MAXDIA% Then GoTo 8059
8024 ''GET #N6%,L%
8025 If CVI(D1$) < LEGA% Then LI% = L%: GoTo 8010
8026 If CVI(D1$) > LEGA% Then LS% = L%: GoTo 8010
8027 '
8028 For K1% = L% To 1 Step -1
8029 ''GET #N6%,K1%:IF CVI(D1$)<>LEGA% THEN 8033
8030 ENCU% = 1: REGINI% = K1%
8031 Next K1%
8032 '
8033 For K1% = 1 To 256
8034 ''GET #N6%,REGINI%+K1%-1
8036 If CVI(D1$) <> LEGA% Then GoTo 8054
8038 CONTA% = K1%: CTA%(K1%) = CVI(D2$): CANTJ(K1%) = CVS(D3$): IM#(K1%) = CVD(D4$): INFO%(K1%) = 0
8039 CODCON% = CTA%(K1%): If (CODCON% > 0 And CODCON% < 100) Then RMD#(CODCON%) = IM#(K1%): QT(CODCON%) = CANTJ(K1%) Else If (CODCON% > 200 And CODCON% < 300) Then RT#(CODCON% - 200) = IM#(K1%)
8040 Next K1%
8052 '
8054 If ENCU% = 0 Then CONTA% = 0: GoTo 8059
8056 If REGINI% = 1 Then GoTo 8058
8057 ''GET #N6%,REGINI%-1:IF CVI(D1$)>=LEGA% THEN ERRO%=1:GOTO 8059
8058 ''GET #N6%,REGINI%+CONTA%:IF CVI(D1$)<>0 THEN IF CVI(D1$)<=LEGA% THEN ERRO%=1
8059 If ERRO% > 0 Then
  Call MENSERR(24, "ERROR DE CONSISTENCIA DE ARCHIVOS  !!!")
  Close: End
     End If
8060 CONTANT% = CONTA%
8061 '
8062 INTERCA% = 0: If ENCU% = 0 Then If LEGA% < ULTLEGA% Then INTERCA% = 1
8065 '
8087 Return
8088 '
8100 ''GET #N2%,EMPL%:GET #N7%,EMPL%:GET #N9%,EMPL%
8102 ACTI$ = AC1$: ASFAM$ = AC2$
8105 For VX% = 1 To NUREM%: SB(VX%) = 0: Next VX%
8106 For VX% = 1 To NUPAR%: PII(VX%) = 0: Next VX%
8110 For VX% = 1 To NUREM%
8112 SB(VX%) = CVS(Mid$(RB$, 4 * VX% - 3))
8114 Next VX%
8115 For VX% = 1 To NUPAR%
8117 PII(VX%) = CVS(Mid$(PI$, 4 * VX% - 3))
8120 Next VX%
8122 AE1 = Asc(F18$)
8123 QM# = ACMEN#(EMPL%): RG1# = RTGAN#(EMPL%)
8125 Return
8130 '
8200 CONTA% = 0
8203 For CONCE1% = 1 To NUCON%
8206 If ESTCON%(CONCE1%) < 1 Then GoTo 8230 ' DESCARTO SUSPENDIDOS
8209 If FORCONI%(CONCE1%) < 1 Then GoTo 8230 ' DESCARTO CONCEPTOS DE CALCULO
8210 If QLCON%(CONCE1%) <> 0 Then If QLCON%(CONCE1%) <> TILIQ% Then GoTo 8230
8212 If CARCON%(CONCE1%) = 3 Or CARCON%(CONCE1%) = FRELI% Then GoTo 8220 ' MANDATORIO
8214 For X% = 1 To 96 Step 2
8215 X1% = CVI(Mid$(ACTI$, X%)): If X1% = 0 Or X1% > CONCE1% Then GoTo 8218
8216 If X1% = CONCE1% Then GoTo 8220           ' PREASIGNADO
8217 Next X%
8218 GoTo 8230
8219 '
8220 CONTA% = CONTA% + 1
8223 CTA%(CONTA%) = CONCE1%: CANTJ(CONTA%) = 0: IM#(CONTA%) = 0: RMD#(CONCE1%) = 0: INFO%(CONTA%) = 0
8224 If CTA%(CONTA%) = RFER% Then CANTJ(CONTA%) = 9 * CAFER%: FORMORI$ = FORCON$(RFER%): QT(RFER%) = CANTJ(CONTA%): GoSub 16000: IM#(CONTA%) = RID#: RMD#(RFER%) = RID#
8226 '
8230 Next CONCE1%
8233 '
8235 For X% = 1 To 96 Step 2
8236 X1% = CVI(Mid$(ACTI$, X%)): If X1% = 0 Then GoTo 8240
8237 If X1% < 0 Then CONTA% = CONTA% + 1: CTA%(CONTA%) = X1%: CANTJ(CONTA%) = 0: IM#(CONTA%) = 0: INFO%(CONTA%) = 0
8238 Next X%
8239 '
8240 For CONCE1% = 1 To NURET%
8246 If ESTRET%(CONCE1%) < 1 Then GoTo 8270 ' DESCARTO SUSPENDIDOS
8249 If FORRETI%(CONCE1%) <> 1 Then GoTo 8270 ' DESCARTO CONCEPTOS DE CALCULO
8250 If QLRET%(CONCE1%) <> 0 Then If QLRET%(CONCE1%) <> TILIQ% Then GoTo 8270
8252 If CARRET%(CONCE1%) = 3 Or CARRET%(CONCE1%) = FRELI% Then GoTo 8260 ' MANDATORIO
8254 For X% = 1 To 96 Step 2
8255 X1% = CVI(Mid$(ACTI$, X%)): If X1% = 0 Or X1% > 200 + CONCE1% Then GoTo 8258
8256 If X1% = 200 + CONCE1% Then GoTo 8260         ' PREASIGNADO
8257 Next X%
8258 GoTo 8270
8259 '
8260 CONTA% = CONTA% + 1
8263 CTA%(CONTA%) = 200 + CONCE1%: CANTJ(CONTA%) = 0: IM#(CONTA%) = 0: INFO%(CONTA%) = 0
8266 '
8270 Next CONCE1%
8273 '
8295 CAUTO% = CONTA%: Return
8297 '
8300 EJU% = ALTERNA%("CORRECCION INTERACTIVA\RECALCULO AUTOMATICO")
8367 If EJU% = 0 Then EJU% = 1
     If EJU% = 2 Then EJU% = (-1)
8370 GoTo 8380
8375 '
8380 Return
8385 '
8500 ''Call BLIMESS(" ESPERE  !!! ")
8501 For H% = 1 To HORE%
8502 ''   GET #N11%,H%
8505    If CVI(H1$) <> LEGA% Then GoTo 8560
8507    CONCE1% = CVI(H3$): CANTI = CVS(H5$)
8510    For HI1% = 1 To CONTA%
8515        If CTA%(HI1%) <> CONCE1% Then GoTo 8555
8517        CANTJ(HI1%) = CANTI
8520          If CONCE1% < 1 Or CONCE1% > 99 Then GoTo 8555
8522            FORMORI$ = FORCON$(CONCE1%)
8525            QT(CONCE1%) = CANTJ(HI1%)
8530            GoSub 16000
8532        IM#(HI1%) = RID#
8555    Next HI1%
8560 Next H%
8570 ''Call BLIMESS("-")
8590 Return
8595 '

9000 ATX% = 0: RV% = 0: If VI% > 0 Then RV% = 1
9001 If LKI% < LCII% Then LKI% = LCII%
9002 GoSub 9200                                      ' PRESENTA VALOR POR DEFECTO
9005 GoSub 9300: On CT% GoTo 9020, 9020, 9020, 9010, 9020 ' ENTRADA DE VARIABLE
9010 GoSub 9400: If AP% = 0 Then GoTo 9000           ' VALIDACION ENTRADA
9015 GoSub 9500                                      ' ASIGNACION ENTRADA
9020 RV% = 0: GoSub 9200                             ' PRESENTA VALOR ACEPTADO
9025 GoTo 9899
9030 '
9100 ''LOCATE RG%, CLX% 'POSICIONAR
9105 Return
9110 '
9200 ''Call CURMOUSE(0)
     ''If RV% = 1 Then Call SETCOL(1, 2) Else Call SETCOL(1, 1)
9201 GoSub 9100
9202 On TV% GoSub 9205, 9205, 9218, 9218, 9288, 9280, 9280 ' IMPRIMIR VALOR POR DEFECTO
9203 ''Call CURMOUSE(1)
     Return
9204 '
9205 If JTI% = 2 Then GoTo 9213
9206 VDSS$ = Str$(VD%): If Left$(VDSS$, 1) = " " Then VDSS$ = Mid$(VDSS$, 2)
9207 If Len(VDSS$) > LCII% Then VDSS$ = String$(LCII%, "*")
9208 If VD% = 0 Then VDSS$ = " "
9209 ' IF VD%>=0 THEN PRINT CHR$(27)+"[1D";
9210 ''Print Left$(VDSS$ + Space$(LCII%), LCII%)
9211 GoTo 9215
9212 '
9213 MA$ = String$(LCII%, "#")
     If RV% = 0 Then
       If VD% = 0 Then
         ''Print Space$(LCII%)
         GoTo 9215
       End If
     End If
9214 If Len(MA$) > LCII% Then
       ''Print String$(LCII%, "*")
       Else
       ''Print USING; MA$; VD%
     End If
9215 Return
9216 '
9218 If RV% = 0 Then
       If Abs(VDDD#) < 0.5 / 10 ^ LD% Then
         ''Print Space$(LCII%)
         GoTo 9277
       End If
     End If
9219 If LD% = 0 Then MA$ = "#": GoTo 9222          ' PRESENTAR VARIABLE DECIMAL
9220 MA$ = "#." + String$(LD%, "#")
9222 If JTI% = 2 Then GoTo 9242
9224 RM = 10
9226 If Abs(VDDD#) < RM Then GoTo 9234
9228 If Left$(MA$, 3) = "###" Then If SC% <> (-1) Then MA$ = "," + MA$
9230 MA$ = "#" + MA$: RM = 10 * RM
9232 GoTo 9226
9234 If VDDD# < 0 Then MA$ = "#" + MA$
9236 While Len(MA$) < LCII%: MA$ = MA$ + " ": Wend
9238 GoTo 9274
9240 '
9242 If Len(MA$) > LCII% Then GoTo 9274
9244 MA$ = String$(LCII% - Len(MA$), "#") + MA$
9246 PP% = InStr(MA$, "."): If PP% = 0 Then PP% = LCII% + 1
9248 For PX% = PP% - 4 To 1 Step -4
9250 If SC% <> (-1) Then Mid$(MA$, PX%, 1) = ","
9252 Next PX%
9254 RM = 1
9256 For PX% = 1 To Len(MA$)
9258 PY% = Asc(Mid$(MA$, PX%, 1))
9260 If PY% = 46 Then GoTo 9268
9262 If PY% = 35 Then RM = 10 * RM
9264 Next PX%
9266 '
9268 If VDDD# < 0 Then RM = RM / 10
9269 If Abs(VDDD#) < RM Then GoTo 9274
9270 If SC% <> (-1) Then SC% = -1: GoTo 9200
9271 If Abs(VDDD#) >= RM Then MA$ = String$(LCII% + 1, 35)
9272 '
9274 If Left$(MA$, 1) = "," Then MA$ = "#" + Mid$(MA$, 2)
9275 If Len(MA$) > LCII% Then
       ''Print String$(LCII%, "*")
       Else
       ''Print USING; MA$; VDDD#
     End If
9276 VDDD# = (Int(10 ^ LD% * VDDD# + 0.5)) / 10 ^ LD%
9277 Return
9278 '
9280 ''Print Left$(VDSS$ + Space$(LCII%), LCII%)
9281 Return
9282 '
9288 ' IF VD%=0 THEN VD%=HO%
9289 FEX = VD%: FE$ = FECHATEX$(FEX)             ' PRESENTAR VARIABLE FECHA
9290 VDSS$ = FE$
9292 GoTo 9280
9294 Return
9296 '
9300 GoSub 9100: ENTRA$ = "": LON% = 0           ' ENTRADA VARIABLE
9301 While INKEY$ <> "": Wend
9302 If PH% < 1 Or Left$(Time$, 5) = TANT$ Or LON% > 0 Then GoTo 9305
9303 If RV% = 1 Then
       ''Call SETCOL(1, 2)
       Else
       ''Call SETCOL(1, 1)
     End If
9304 GoSub 9100
9305 If LON% < 1 Then
 ''Call TRAPRIN
 If RV% = 1 Then
 ''Call SETCOL(1, 2)
 Else
 ''Call SETCOL(1, 1)
 End If
 GoSub 9100
     End If

     A$ = INKEY$
     If A$ = "" Then
  If LON% = 0 Then A$ = COMOUSE$
  If A$ = "" Then GoTo 9302
  GoSub 9900
      If RGI1% = RG% Then
   If CL1% >= CLX% And CL1% < CLX% + LCII% Then
       A$ = Chr$(0) + Chr$(60): GoTo 9306
   End If
      End If
  GoSub 9910
  GoSub 9100
     End If

9306 AAII% = Asc(A$): If AAII% <> 0 Then GoTo 9308
9307 A$ = Mid$(A$, 2)
 If A$ = Chr$(59) Then VDC% = 74: VDSS$ = ENTRA$: CT% = 5: GoTo 9899

 If A$ = Chr$(60) Then
     If TV% <> 5 Then
  VDC% = 60
  VDSS$ = ENTRA$
  CT% = 0
  GoTo 9899
     Else
  ''EX$ = FECHATEX$(SELFECH(RG%, CL%))
  If EX$ <> "  /  /  " Then ENTRA$ = EX$
  GoTo 9334
     End If
 End If

 If Len(ENTRA$) = 0 Then
     AAII% = Asc(A$)
     CT% = 5
     VDC% = AAII%
     GoTo 9340
   Else
     ''Print Chr$(7);
     GoTo 9305
 End If
9308 If AAII% = 8 Then GoTo 9322 Else If AAII% = 9 Then GoTo 9600 ' PASA A MODO EDICION
9309 If AAII% = 63 Then VDSS$ = ENTRA$: CT% = 0: GoTo 9899
9310 If AAII% = 27 Then ENTRA$ = Chr$(27): GoTo 9338
9311 If AAII% = 13 Then GoTo 9334
9312 If AAII% = 42 Or AAII% = 43 Then AAII% = 76: CT% = 5: VDC% = AAII%: GoTo 9340
9313 If AAII% < 32 Then GoTo 9302
9314 If TV% > 4 Then If LON% < 255 Then If LON% >= LCII% Then If LON% < LKI% Then ENTRA$ = ENTRA$ + A$: LON% = LON% + 1: GoSub 9350: GoTo 9318
9315 If LON% >= LKI% Then
       ''Print Chr$(7);
       ATX% = 1
     End If
9316 If LON% >= LCII% Then GoTo 9302
9317 ''Print A$;: LON% = LON% + 1: ENTRA$ = ENTRA$ + A$
9318 GoTo 9302
9320 '
9322 If LON% < 1 Then ENTRA$ = Chr$(8): GoTo 9338
9324 If LON% <= LCII% Then
       ''LOCATE RG%, CLX% + LON% - 1
       ''Print " " + Chr$(29);
     End If
9326 ENTRA$ = Left$(ENTRA$, Len(ENTRA$) - 1)
9328 LON% = LON% - 1
9329 If LON% >= LCII% - 1 Then GoSub 9350
9330 GoTo 9302
9332 '
9334 If ENTRA$ = "" Then GoTo 9338
9336 If LCII% > Len(ENTRA$) Then
       ''Print Space$(LCII% - Len(ENTRA$))
     End If
9338 If ENTRA$ = Chr$(27) Then CT% = 1 Else If ENTRA$ = Chr$(8) Then CT% = 2 Else If ENTRA$ = "" Then CT% = 3 Else CT% = 4
9340 If CT% = 5 And VDC% = 60 Then VDSS$ = "": CT% = 0: GoTo 9899
9341 Return
9342 '
9350 PPX% = Len(ENTRA$) - LCII% + 1: If PPX% < 1 Then PPX% = 1
9351 GoSub 9100
     ''Print Mid$(ENTRA$, PPX%);
9352 Return
9354 '
9400 On TV% GoSub 9406, 9422, 9440, 9456, 9474, 9486, 9494 ' CONTROL ENTRADA
9401 If AP% = 0 Then
       GoSub 9100
       ''Print Space$(LCII%)
     End If
9402 Return
9404 '
9406 AP% = 1                                    ' VALIDAR VARIABL. ENTERA +Y-
9408 For I% = 1 To Len(ENTRA$)
9410 B% = Asc(Mid$(ENTRA$, I%))
9412 If B% < 48 Or B% > 57 Then AP% = 0: GoTo 9418
9414 Next I%
9416 If Val(ENTRA$) > 32767 Then AP% = 0
9418 Return
9420 '
9422 AP% = 1                                    ' VALIDAR VARIBL. ENTERA -
9424 For I% = 1 To Len(ENTRA$)
9426 B% = Asc(Mid$(ENTRA$, I%))
9428 If B% = 45 Then If I% = 1 Then GoTo 9432
9430 If B% < 48 Or B% > 57 Then AP% = 0: GoTo 9436
9432 Next I%
9434 If Val(ENTRA$) > 32767 Or Val(ENTRA$) < -32768! Then AP% = 0
9436 Return
9438 '
9440 AP% = 1: PUN% = 0                          ' VALIDAR VARIABL. DECIMAL +Y-
9442 For I% = 1 To Len(ENTRA$)
9444 B% = Asc(Mid$(ENTRA$, I%))
9446 If B% = 46 Then If PUN% = 0 Then PUN% = 1: GoTo 9450
9448 If B% < 48 Or B% > 57 Then AP% = 0: GoTo 9452
9450 Next I%
9452 Return
9454 '
9456 AP% = 1: PUN% = 0                          ' VALIDAR VARIABL. DECIMAL -
9458 For I% = 1 To Len(ENTRA$)
9460 B% = Asc(Mid$(ENTRA$, I%))
9462 If B% = 45 Then If I% = 1 Then GoTo 9468
9464 If B% = 46 Then If PUN% = 0 Then PUN% = 1: GoTo 9468
9466 If B% < 48 Or B% > 57 Then AP% = 0: GoTo 9470
9468 Next I%
9470 Return
9472 '
9474 AP% = 1                                    ' VALIDAR FECHA
9476 ENTRA1$ = ENTRA$ + Mid$(VDSS$, Len(ENTRA$) + 1)
9477 II% = InStr(ENTRA1$, "."): If II% > 0 Then Mid$(ENTRA1$, II%, 1) = "/": GoTo 9477
9478 II% = InStr(ENTRA1$, "-"): If II% > 0 Then Mid$(ENTRA1$, II%, 1) = "/": GoTo 9478
9479 FE$ = ENTRA1$: FEX = FECHANUM(FE$)
9480 ' IF FEX=0 THEN AP%=0
9482 Return
9484 '
9486 AP% = 1                                    ' VALIDAR ALFABETICOS
9488 For I% = 1 To Len(ENTRA$)
9489 B% = Asc(Mid$(ENTRA$, I%))
9490 If B% < 32 Then AP% = 0: GoTo 9492
9491 Next I%
9492 Return
9493 '
9494 AP% = 1: CI% = Val(ENTRA$): If CI% < 1 Then AP% = 0
9495 Return
9496 '
9500 On TV% GoTo 9502, 9502, 9506, 9506, 9504, 9508, 9507 ' ASIGNACION ENTRADA
9502 VD% = Val(ENTRA$): GoTo 9510
9504 VD% = FEX: GoTo 9510
9506 VDDD# = Val(ENTRA$): GoTo 9510
9507 VD% = CI%
9508 VDSS$ = ENTRA$: If Len(VDSS$) < LCII% Then VDSS$ = VDSS$ + Space$(LCII% - Len(VDSS$))
9510 Return
9512 '
9600 If Len(ENTRA$) > 0 Then GoTo 9623
9602 On TV% GoTo 9606, 9606, 9610, 9610, 9615, 9622, 9622 ' TRANSFORMAR EN STRING
9606 VDSS$ = Str$(VD%): GoTo 9620                        ' VAR. ENTERA
9610 VDSS$ = Str$((Int(10 ^ LD% * VDDD# + 0.5)) / 10 ^ LD%): GoTo 9620 ' VAR. DECIMAL
9615 FEX = VD%: FE$ = FECHATEX$(FEX): VDSS$ = FE$: GoTo 9622 ' VAR.FECHA
9620 If Val(VDSS$) >= 0 Then VDSS$ = Mid$(VDSS$, 2)          ' SUPR.LEAD.BLANKS
9622 ENTRA$ = VDSS$: ENTRANT$ = VDSS$: GoSub 9700          ' SUPR.BLANKS FIN
9623 '
9625 PX% = 1: PY% = 1: GoSub 9710: PPXX$ = ENTRA$        ' PRESENTAR VALOR
9627 '
9631 If PY% > LCII% Then PY% = LCII%
9632 PPXX$ = Mid$(ENTRA$, PX%, 1): If PPXX$ = "" Then PPXX$ = " "
9633 ''LOCATE RG%, CLX% + PY% - 1: Call SETCOL(1, 1): Print PPXX$ + Chr$(29);: Call SETCOL(1, 2) 'POSICIONAR
9634 '
9635 A$ = INKEY$ + COMOUSE$
     If Len(A$) = 0 Then GoTo 9635

     ''LOCATE RG%, CLX% + PY% - 1                         ' POSICIONAR
     ''Call SETCOL(1, 2)

9636 AAII% = Asc(A$): If AAII% <> 0 Then GoTo 9660
9637 AAII% = Asc(Mid$(A$, 2))
9638 '
9639 ''If AAII% = 71 Then Call SETCOL(1, 2): Print PX$;: GoTo 9625
9640 ''If AAII% = 79 Then Call SETCOL(1, 2): Print PX$;: PX% = 1 + Len(ENTRA$): GoSub 9720: GoTo 9631
9642 If AAII% <> 75 Or PX% <= 1 Then GoTo 9645
9643 ''Call SETCOL(1, 2): Print PX$;: PX% = PX% - 1: PY% = PY% - 1: If PY% >= 1 Then GoTo 9631 Else PY% = 1: GoSub 9710: GoTo 9631
9645 If AAII% <> 77 Then GoTo 9654
9646 ''If PX% > Len(ENTRA$) Then Print Chr$(7);: GoTo 9635
9647 ''Call SETCOL(1, 2): Print PX$;
9648 PX% = PX% + 1: PY% = PY% + 1: If PY% <= LCII% Then GoTo 9631 Else PY% = LCII%: GoSub 9710: GoTo 9631
9654 ''If AAII% <> 83 Then Print Chr$(7);: GoTo 9635
9655 ENTRA$ = Left$(ENTRA$, PX% - 1) + Mid$(ENTRA$, PX% + 1): '' Print Mid$(ENTRA$ + " ", PX%, LCII% - PY% + 1)
9656 GoTo 9631
9657 '
9660 If AAII% <> 8 Then GoTo 9665
9661 ''If PX% <= 1 Then Print Chr$(7);: GoTo 9635
9662 PX% = PX% - 1: PY% = PY% - 1: ENTRA$ = Left$(ENTRA$, PX% - 1) + Mid$(ENTRA$, PX% + 1): GoSub 9710: GoTo 9631
9663 ''Print Chr$(29);: GoTo 9655
9664 '
9665 If AAII% <> 9 Then GoTo 9670
9666 ''If PX% + 9 > Len(ENTRA$) Then Print Chr$(7);: GoTo 9631
9667 ''Call SETCOL(1, 2): Print PX$;: PX% = PX% + 10: PY% = PY% + 10: If PY% <= LCII% Then GoTo 9631 Else PY% = LCII%: GoSub 9710: GoTo 9631
9668 '
9670 If AAII% = 27 Then ENTRA$ = ENTRANT$: GoTo 9690
9671 If AAII% = 13 Then GoTo 9690
9672 If AAII% = 43 Then If TV% < 6 Then GoTo 9690
9673 If AAII% < 32 Then GoTo 9635
9674 '
9675 ''If Len(ENTRA$) >= LKI% Then Print Chr$(7);: ATX% = 1: GoTo 9635
9676 ENTRA$ = Left$(ENTRA$, PX% - 1) + A$ + Mid$(ENTRA$, PX%): PPXX$ = A$
9677 ''If PY% < LCII% - 1 Then Print Mid$(ENTRA$ + " ", PX%, LCII% - PY% + 1);
9678 GoTo 9648
9690 '
9692 GoTo 9334
9695 '
9700 For PX% = Len(ENTRA$) To 1 Step -1
9701 If Asc(Mid$(ENTRA$, PX%)) > 32 Then ENTRA$ = Left$(ENTRA$, PX%): GoTo 9704
9702 Next PX%
9703 ENTRA$ = ""
9704 Return
9705 '
9710 GoSub 9100: If PY% < 1 Then PY% = 1
9711 ''Call SETCOL(1, 2)
9712 ENTRB$ = Mid$(ENTRA$, PX%, LCII%)
9713 ''Print Left$(ENTRB$ + Space$(LCII%), LCII% - PY% + 1)
9714 GoSub 9100
9715 '
9720 If PX% > LKI% Then PX% = LKI%
9721 PY% = PX%: If PY% > LCII% Then PY% = LCII% Else If PY% < 1 Then PY% = 1
9722 GoSub 9100
9723 ''Call SETCOL(1, 2): If PX% = PY% Then ENTRB$ = ENTRA$: GoTo 9725
9724 ENTRB$ = Mid$(ENTRA$, PX% - LCII% + 1, LCII%)
9725 ''Print Left$(ENTRB$ + Space$(LCII%), LCII%);
9726 Return
9727 '
9899 GoTo 9999

9900 CL1% = 0: RGI1% = 0
     If Len(A$) = 4 Then
 If Mid$(A$, 1, 1) = "x" Then
    CL1% = Asc(Mid$(A$, 2, 1))
    If Mid$(A$, 3, 1) = "y" Then
       RGI1% = Asc(Mid$(A$, 4, 1))
    End If
 End If
     End If
     Return

9910 If (CL1% > 0 And CL1% <= 80 And RGI1% > 0 And RGI1% <= 23) Then
 RV% = 0: GoSub 9200
 CT% = 5
 VDC% = (-1) * (100 * RGI1% + CL1%)
 VDSS$ = ""
 GoTo 9899
     End If
     Return

9999 Return
  
14000 If POSEDI% < 0 Then VDSS$ = CUE$
14005 While Right$(VDSS$, 1) = " ": VDSS$ = Left$(VDSS$, Len(VDSS$) - 1): Wend
14010 While Len(VDSS$) < 3: VDSS$ = "0" + VDSS$: Wend: COD$ = VDSS$: CUE$ = Left$(VDSS$ + Space$(8), 8)
14015 Return
14020 '
14100 TICO% = 0: ENCUE% = 0: AJU% = 0: POSEDI% = 1
14105 While Right$(COD$, 1) = " ": COD$ = Left$(COD$, Len(COD$) - 1): Wend
14110 For I% = 1 To Len(COD$)
14115   B% = Asc(Mid$(COD$, I%)): If B% < 48 Or B% > 57 Then GoTo 14390 ' DESTAJO
14120 Next I%
14125 CODCON = Val(COD$): If CODCON < 1 Or CODCON > 599 Then GoTo 14390 ' DESTAJO
14130 '
14145 TICO% = 1 + Int(CODCON / 200): CODCON% = CODCON Mod 200 '1=REMUN, 2=RETEN, 3=ASFAM
14150 AJU% = Int(CODCON% / 100): CODCON% = CODCON% Mod 100 '0=NORMAL, 1=AJUSTE
14155 CDI% = 200 * (TICO% - 1) + 100 * AJU% + CODCON%
14160 On TICO% GoTo 14200, 14230, 14260
14165 '
14200 If CODCON% < 1 Or CODCON% > NUCON% Then GoTo 14390 ' CONCEPTO NO DEFINIDO
14205 ENCUE% = 1: DENO$ = Left$(CON$(CODCON%), 28)
14210 MCAL% = FORCONI%(CODCON%): FORMU$ = FORCON$(CODCON%)
14215 GoTo 14390
14220 '
14230 If CODCON% < 1 Or CODCON% > NURET% Then GoTo 14390 ' RETENCION NO DEFINIDA
14235 ENCUE% = 1: DENO$ = Left$(RET$(CODCON%), 28)
14240 MCAL% = FORRETI%(CODCON%): FORMU$ = FORRET$(CODCON%): If RCLRET%(CODCON%) = 2 Then POSEDI% = -1
14245 GoTo 14390
14250 '
14260 If CODCON% < 1 Or CODCON% > NUASI% Then GoTo 14390 ' ASIGNACION NO DEFINIDA
14265 ENCUE% = 1: DENO$ = Left$(ASI$(CODCON%), 28)
14270 MCAL% = 3
14275 GoTo 14390
14280 '
14390 If AJU% > 0 Then DENO$ = "AJUSTE " + Left$(DENO$, 21)
14392 Return
14395 '
14500 REPE% = 0: If AJU% > 0 Or TICO% = 4 Then GoTo 14520
14505 For XCV% = 1 To CONTA%
14510 If CTA%(XCV%) = CDI% Then REPE% = 1: GoTo 14520
14515 Next XCV%
14520 Return
14525 '
15000 If BARRE% > 0 Then GoSub 20000: GoTo 15008
15002 RG% = 5: CLX% = 11: LCII% = 5: LKI% = 5: TV% = 1: JTI% = 1: VD% = LEGA%: GoSub 9000
15003 On CT% GoTo 1100, 15002, 15005, 15005, 15002
      Call SELECHO("PADRON")  '', VD%, VDDD#, VDSS$)
15005 ''GoSub 15850: If ENCU% < 1 Then Print Chr$(7);: GoTo 15002 ' buscar si es un legajo existente
15006 '
15007 LEGA% = VD%
15008 ''LOCATE 5, 10: Print LEGA%;
      ''LOCATE 5, 19: Print NOM$
      ''LOCATE 5, 59: Print Left$(GRU$(GRU%), 21)
15009 If ESEMP% > 0 Then If ESEMP% < 3 Then GoTo 15014
15010 ''Print Chr$(7)
      ''LOCATE 5, 59: Call SETCOL(1, 3): Print "SUSPENDIDO !!!": Call SETCOL(1, 1)
15011 GoTo 15002
15012 '
15014 GoSub 19000: YANTI% = 0      ' LIMPIAR VECTORES
15015 '
15016 GoSub 8000                   ' BUSCAR LIQUIDACION ANTERIOR
15017 GoSub 8100                   ' LEER SALARIOS BASICOS Y ACTI$
15018 CAUTO% = 0: If CONTANT% = 0 Then GoSub 8200 ' LEER CONC. MANDATORIOS A INFORMAR
15020 If HORE% > 0 Then GoSub 8500 ' LEER HORAS DE HORELOJ
15023 '
15025 INFE% = 1: RGL% = 1: ''GoSub 15600       ' PRESENTAR DATOS
15035 RGL% = 1: RG% = 9
15040 '
15100 CUE$ = CTB$(RG% - 8): If RGL% <= CONTA% Then COD$ = CUE$: GoSub 14100
15101 If RGL% <= CAUTO% Then GoTo 15130
15102 COLOG% = 1: CLX% = 2: LCII% = 8: LKI% = 8: JTI% = 1: TV% = 6: VDSS$ = CUE$
      GoSub 9000
15105 ''On CT% GoTo 15343, 15100, 15106, 15106, 15400
      VD% = Val(VDSS$): Call SELECHO("ECONCEP") '', VD%, VDDD#, VDSS$)
      VDSS$ = AJUSTI$(Str$(VD%), 8)
15106 GoSub 14000
15110 '
15115 ''GoSub 14100: If ENCUE% < 1 Then Print Chr$(7);: GoTo 15100 ' COD. NO REGISTR.
15116 ''GoSub 14500: If REPE% > 0 Then Print Chr$(7);: GoTo 15100 ' COD. REPETIDO
15117 If CDI% <> 200 + RTPR% Then GoTo 15122
15118 ''FLG% = 1: PRESTA# = 0: NETO# = SUMD# - SUMH#: GoSub 17407: If PRESTA# <= 0.005 Then Print Chr$(7);: GoTo 15100
15119 IM#(RGL%) = PRESTA#: IMPO#(RG% - 8) = PRESTA#: CLX% = 68: LCII% = 12: LKI% = 12: LD% = 2: JTI% = 2: TV% = 3: VDDD# = PRESTA#: GoSub 9200
15120 '
15122 ''LOCATE RG%, 2: Print CUE$
      ''LOCATE RG%, 12: Print Left$(DENO$, 30)
      If RGL% > CONTA% Then GoTo 15125
15123 GoSub 18100: If (CDI% > 0 And CDI% < 100) Then QT(CDI%) = CANTJ(RGL%): RMD#(CDI%) = IM#(RGL%) Else If (CDI% > 200 And CDI% < 300) Then RT#(CDI% - 200) = Abs(IM#(RGL%))
15124 ''GoSub 15500
15125 CTB$(RG% - 8) = CUE$: DEX$(RG% - 8) = DENO$: CTA%(RGL%) = CDI%
15127 '
15130 If CONTANT% > 0 Then If RGL% <= CONTA% Then GoTo 15150
15131 If MCAL% >= 2 Then GoTo 15150            ' INFORME HORAS O CANTIDAD
15132 ''If MCAL% = 1 Or AJU% > 0 Then GoTo 15200 ' INFORME IMPORTE
15133 '
15135 FORMORI$ = FORMU$: QT0 = CANTJ(RGL%): GoSub 16000        ' CALCULO DE IMPORTE
15136 IMPO#(RG% - 8) = RID#: IM#(RGL%) = RID#
15137 If AJU% > 0 Then GoTo 15140
15138 If TICO% = 1 Then RMD#(CODCON%) = RID# Else If TICO% = 2 Then RT#(CODCON%) = RID#
15140 CLX% = 55: If CTA%(RGL%) > 200 Then If CTA%(RGL%) < 400 Then CLX% = 68
15142 LCII% = 12: LKI% = 12: LD% = 2: JTI% = 2: TV% = 3: VDDD# = RID#: GoSub 9200
15143 ''GoTo 15300
15145 '
15150 COD$ = CTB$(RG% - 8): GoSub 14100
15152 CANTUSI = CAN(RG% - 8): SCA% = SC%: CLX% = 46
15154 COLOG% = 2: LCII% = 8: LKI% = 8: LD% = 2: JTI% = 2: TV% = 4: VDDD# = CDbl(CANTUSI)
      GoSub 9000
15156 ''On CT% GoTo 15343, 15100, 15158, 15158, 15400
      ''Print Chr$(7);: GoTo 15152
15157 '
''15158 If POSEDI% < 0 Then VDDD# = CDbl(CANTUSI): GoSub 9100: Print USING; "#####.##"; CANTUSI
''15159 If VDDD# < 0.005 Then GoSub 9100: Print Space$(LCII%);: CAN(RG% - 8) = 0
''15160 CAN(RG% - 8) = VDDD#: CANTJ(RGL%) = CAN(RG% - 8)
''15161 If TICO% = 1 Then If AJU% < 1 Then QT(CODCON%) = CAN(RG% - 8)
''15162 If TICO% = 3 Then RID# = CDbl(CAN(RG% - 8) * VALASI(CODCON%)): GoTo 15170
''15163 If TICO% = 4 Then RID# = CDbl(CAN(RG% - 8) * VU): GoSub 15250: GoTo 15170
''15164 '
''15165 RID# = 0: If AJU% = 0 Then FORMORI$ = FORMU$: QT0 = CANTJ(RGL%): GoSub 16000: GoTo 15170 ' CALCULO DE IMPORTE
''15166 If RGL% <= CONTA% Then RID# = IM#(RGL%)
''15170 IMPO#(RG% - 8) = RID#: IM#(RGL%) = RID#: If AJU% > 0 Then GoTo 15172
''15171 If TICO% = 1 Then RMD#(CODCON%) = RID# Else If TICO% = 2 Then RT#(CODCON%) = RID#
''15172 If RGL% <= CONTA% Then If RGL% > CAUTO% Then GoSub 15500: GoTo 15200
''15173 If AJU% > 0 Then GoSub 15500: GoTo 15200
''15174 CLX% = 55: If CTA%(RGL%) > 200 Then If CTA%(RGL%) < 400 Then CLX% = 68
''15175 LCII% = 12: LKI% = 12: LD% = 2: JTI% = 2: TV% = 3: VDDD# = RID#: GoSub 9200
''15180 GoTo 15300
''15198 '
''15200 COD$ = CTB$(RG% - 8): GoSub 14100: IMPO# = IMPO#(RG% - 8): SCA% = SC%
''15202 CLX% = 55: If CDI% > 200 Then If CDI% < 400 Then CLX% = 68
''15204 COLOG% = 3: LCII% = 12: LKI% = 12: LD% = 2: JTI% = 2: TV% = 4: VDDD# = IMPO#
''      GoSub 9000
''15206 On CT% GoTo 15343, 15207, 15208, 15208, 15400
''      Print Chr$(7);: GoTo 15202
''15207 If MCAL% < 2 Then GoTo 15100 Else GoTo 15150
''15208 If POSEDI% < 0 Then VDDD# = IMPO#: GoSub 9100: Print USING; MASPA$; IMPO#
''15209 If AJU% > 0 Then GoTo 15211
''15210 If TICO% = 1 Then RMD#(CODCON%) = VDDD# Else If TICO% = 2 Then RT#(CODCON%) = VDDD#
''15211 INFO%(RGL%) = 0: If VDDD# <> IMPO# Then INFO%(RGL%) = 1
''15212 IMPO#(RG% - 8) = VDDD#: IM#(RGL%) = VDDD#: If (Abs(VDDD#) > 0.005 Or RGL% <= CONTA%) Then GoTo 15300 Else GoTo 15150
''15214 GoTo 15300
''15215 '
''15250 CLX% = 41: LCII% = 4: LKI% = 4: LD% = 0: JTI% = 2: TV% = 1: VD% = 100
''      GoSub 9000
''15252 On CT% GoTo 15343, 15250, 15255, 15255, 15250
''      Print Chr$(7);: GoTo 15250
''15255 RID# = CDbl(CAN(RG% - 8) * VU * VD% / 100)
''15260 Return
''15265 '
''15300 If RGL% > CONTA% Then CONTA% = RGL%
''15302 If SC% <> SCA% Then GoSub 15600: SCA% = SC%
''15303 GoSub 15500
''15305 If RGL% < CONTA% Or POSEDI% >= 0 Then RGL% = RGL% + 1: RG% = RGL% - INFE% + 9
''15310 If RG% <= 9 + RLB% Then GoTo 15335
''15315 INFE% = RGL% - 5: SUPE% = RGL% + 4
''15320 If INFE% < 1 Then INFE% = 1
''15325 If SUPE% > CONTA% Then SUPE% = CONTA%
''15330 GoSub 15600
''15335 RG% = RGL% - INFE% + 9
''15340 GoTo 15100
''15342 '
''15343 CONTOT% = CONTA%: Return
''15344 '
''15400 If VDC% < 71 Or VDC% > 83 Then GoTo 15490  ' VOLVER SIN CAMBIOS
''15401 On VDC% - 70 GoTo 15410, 15414, 15420, 15482, 15426, 15485, 15430, 15490, 15435, 15440, 15448, 15460, 15470
''15409 '
''15410 RGL% = 1: If RGL% >= INFE% Then GoTo 15490 ' IR A PRIMER REGISTRO
''15411 INFE% = 1: SUPE% = INFE% + RLB%: If SUPE% > CONTA% Then SUPE% = CONTA%
''15412 GoSub 15600: GoTo 15490
''15413 '
''15414 RGL% = RGL% - 1: If RGL% < 1 Then RGL% = 1 ' REGISTRO ANTERIOR
''15415 If RGL% >= INFE% Then GoTo 15490
''15416 INFE% = RGL% - 5: If INFE% < 1 Then INFE% = 1
''15417 SUPE% = INFE% + RLB%: If SUPE% > CONTA% Then SUPE% = CONTA%
''15418 GoSub 15600: GoTo 15490
''15419 '
''15420 RGL% = RGL% - RLB%: If RGL% < 1 Then RGL% = 1 ' PAGINA ANTERIOR
''15421 If RGL% >= INFE% Then GoTo 15490
''15422 INFE% = RGL% - 5: If INFE% < 1 Then INFE% = 1
''15423 SUPE% = INFE% + RLB%: If SUPE% > CONTA% Then SUPE% = CONTA%
''15424 GoSub 15600: GoTo 15490
''15425 '
''15426 If RGL% > CONTA% Or COLOG% = 1 Then GoTo 15490      ' CURSOR IZQUIERDA
''15427 If (COLOG% = 2 Or MCAL% >= 2) Then COLOG% = COLOG% - 1: GoTo 15490
''15428 COLOG% = 1: GoTo 15490
''15429 '
''15430 If (RGL% > CONTA% Or COLOG% = 3) Then GoTo 15490                  ' CURSOR DERECHA
''15431 If (COLOG% = 2 Or MCAL% >= 2) Then COLOG% = COLOG% + 1: GoTo 15490
''15432 COLOG% = 3: GoTo 15490
''15434 '
''15435 RGL% = CONTA%: If RGL% < INFE% + RLB% Then GoTo 15438 ' IR A ULTIMO REGISTRO
''15436 SUPE% = CONTA%: INFE% = CONTA% - 5: If INFE% < 1 Then INFE% = 1
''15437 GoSub 15600                              ' PRESENTAR DATOS
''15438 If POSEDI% >= 0 Then RGL% = CONTA% + 1: COLOG% = 1: GoTo 15490
''15439 '
''15440 If RGL% <= CONTA% Then RGL% = RGL% + 1 ' REGISTRO SIGUIENTE
''15441 If RGL% > CONTA% Then COLOG% = 1
''15442 If RGL% <= INFE% + RLB% Then GoTo 15446
''15443 INFE% = RGL% - 5: If INFE% < 1 Then INFE% = 1
''15444 SUPE% = INFE% + RLB%: If SUPE% > CONTA% Then SUPE% = CONTA%
''15445 GoSub 15600
''15446 GoTo 15490
''15447 '
''15448 RGL% = RGL% + RLB%: If RGL% > CONTA% Then RGL% = CONTA% ' PAGINA SIGUIENTE
''15449 If RGL% <= INFE% + RLB% Then GoTo 15490
''15450 INFE% = RGL% - 5: If INFE% < 1 Then INFE% = 1
''15451 SUPE% = INFE% + RLB%: If SUPE% > CONTA% Then SUPE% = CONTA%
''15452 GoSub 15600
''15453 GoTo 15490
''15454 '
''15460 GoTo 15490                         ' INSERCION DESHABILITADA
''15465 '
''15470 If RGL% > CONTA% Or POSEDI% < 0 Then GoTo 15475  ' BORRADO
''15471 Call BLIMESS(" ESPERE  !!! ")
''15472 GoSub 18100: GoSub 18200
''15473 GoSub 15600
''15474 GoSub 15500
''15475 If RGL% > CONTA% Then COLOG% = 1
''15476 Call BLIMESS("-")
''15477 GoTo 15490
''15479 '
''15482 GoTo 15490
''15483 GoTo 15490
''15484 '
''15485 If LSUE% > 0 Then Print Chr$(7);: GoTo 15488
''      Call BLIMESS(" RE-CALCULANDO ")
''15486 GoSub 17001
''      Call BLIMESS("-")
''15487 If RGL% > 1 + RLB% Then RGL% = 1
''15488 GoTo 15411
''15489 '
''15490 RG% = RGL% - INFE% + 9
''15491 If Abs(CTA%(RGL%)) < 1 Then COLOG% = 1
''15492 On COLOG% GoTo 15100, 15150, 15200
''15493 GoTo 15100
''15494 '
''15500 SUMD# = 0: SUMH# = 0: RMD#(0) = 0: DJ0# = 0: K2% = CONTA%: If RGL% > K2% Then If CONTA% > 0 Then If Abs(CTA%(RGL%)) > 0 Then K2% = RGL%
''15501 MASPA$ = "#,###,###.##": RAMAX# = 9999999#: If SC% = (-1) Then MASPA$ = "#########.##": RAMAX# = 999999999#
''15502 '
''15503 SCA% = SC%
''15505 For K1% = 1 To K2%
''15506 If IM#(K1%) > 9999999# Or IM#(K1%) < -999999# Then SC% = (-1)
''15507 If CTA%(K1%) < 0 Then DJ0# = DJ0# + (Int(100 * IM#(K1%) + 0.5)) / 100
''15508 If (CTA%(K1%) <= 200 Or CTA%(K1%) >= 400) Then SUMD# = SUMD# + (Int(100 * IM#(K1%) + 0.5)) / 100 Else SUMH# = SUMH# + (Int(100 * IM#(K1%) + 0.5)) / 100
''15509 '
''15511 If CTA%(K1%) < 1 Or CTA%(K1%) > 199 Then GoTo 15515
''15512 CTX% = CTA%(K1%) Mod 100
''15513 If CTX% > 0 Then If CTX% <= NUCON% Then If SURETI%(CTX%) > 0 Then RMD#(0) = RMD#(0) + (Int(100 * IM#(K1%) + 0.5)) / 100
''15515 Next K1%
''15516 If SUMD# > 9999999# Or SUMH# > 9999999# Then SC% = (-1)
''15519 If SC% <> SCA% Then GoSub 15600
''15522 LOCATE 22, 55: If SUMD# <= RAMAX# Then Print USING; MASPA$; SUMD# Else Print "************"
''15525 LOCATE 22, 68: If SUMH# <= RAMAX# Then Print USING; MASPA$; SUMH# Else Print "************"
''15530 '
''15533 If CONTOT% < CONTA% Then CONTOT% = CONTA%
''15534 Return
''15535 '
''15600 If INFE% < 1 Then INFE% = 1
''15601 For K1% = INFE% To INFE% + RLB%
''15605 GoSub 15700              ' LEER DE ARCHIVO Y TRANSFORMAR
''15610 GoSub 15650
''15615 If INFE% <= 1 Then If K1% > CONTA% Then GoTo 15630
''15625 Next K1%
''15630 GoSub 15500
''15635 Return
''15640 '
''15650 NR1% = K1% - INFE% + 9: NR1$ = Mid$(Str$(NR1%), 2): RGF% = NR1% - 8: MASPA$ = "#,###,###.##": RAMAX# = 9999999#: If SC% = (-1) Then MASPA$ = "#########.##": RAMAX# = 999999999#
''15655 LOCATE NR1%, 2: Print CUE$: CTB$(RGF%) = CUE$
''15656 DEX$(RGF%) = DENO$: LOCATE NR1%, 12: Print DENO$
''15662 LOCATE NR1%, 46: CAN(RGF%) = CANTUSI: If Abs(CANTUSI) < 0.005 Then Print Space$(8) Else If (CANTUSI >= (-999.99) And CANTUSI <= 9999.99) Then Print USING; "#,###.##"; CANTUSI Else Print "********"
''15663 Col% = 55: COZ% = 68: SIG% = 1: If TICO% = 2 Then Col% = 68: COZ% = 55: SIG% = -1
''15665 LOCATE NR1%, Col%
'' IMPO#(RGF%) = IMPO#: SIG%(RGF%) = SIG%: If Abs(IMPO#) < 0.005 Then Print Space$(12) Else If (IMPO# >= (-1) * RAMAX# / 10 And IMPO# <= RAMAX#) Then Print USING; MASPA$; IMPO# Else Print "************"
''15666 LOCATE NR1%, COZ%: Print "            "
''15670 Return
''15675 '
15700 ''If K1% > CONTA% Then CUE% = 0: CUE$ = Space$(8): DENO$ = Space$(28): CANTUSI = 0:: IMPO# = 0: GoTo 15725
''15710 CUE% = CTA%(K1%)
''15715 IMPO# = IM#(K1%): CANTUSI = CANTJ(K1%)
''      GoSub 15760
''15720 If IMPO# > 999999999# Then SC% = (-1)
15725 Return
''15730 '
''15760 If Abs(CUE%) < 1 Then CUE$ = Space$(8): DENO$ = Space$(28): GoTo 15762
''15761 If CUE% < 0 Then GoTo 15790 '************** CONCEPTO DE DESTAJO
''15762 TICO% = 1 + Int(CUE% / 200): CODCON% = CUE% Mod 200
''15763 AJU% = Int(CODCON% / 100): CODCON% = CODCON% Mod 100
''15764 On TICO% GoTo 15770, 15775, 15780
''15765 '
''15770 If CODCON% < 1 Or CODCON% > NUCON% Then GoTo 15790
''15771 DENO$ = CON$(CODCON%): If AJU% > 0 Then DENO$ = "AJUSTE " + DENO$
''15772 DENO$ = Left$(DENO$, 28)
''      RMD#(CODCON%) = IMPO#
''15773 GoTo 15785
''15774 '
''15775 If CODCON% < 1 Or CODCON% > NURET% Then GoTo 15790
''15776 DENO$ = RET$(CODCON%): If AJU% > 0 Then DENO$ = "AJUSTE " + DENO$
''15777 DENO$ = Left$(DENO$, 28)
''      RT#(CODCON%) = IMPO#
''15778 GoTo 15785
''15779 '
''15780 If CODCON% < 1 Or CODCON% > NUASI% Then GoTo 15790
''15781 DENO$ = ASI$(CODCON%): If AJU% > 0 Then DENO$ = "AJUSTE " + DENO$
''15782 DENO$ = Left$(DENO$, 28)
''15783 GoTo 15785
''15784 '
''15785 CUE$ = Mid$(Str$(CUE%), 2): While Len(CUE$) < 3: CUE$ = "0" + CUE$: Wend: CUE$ = Left$(CUE$ + Space$(8), 8)
''15786 GoTo 15798
''15787 '
''15790 CUE$ = Space$(8): DENO$ = Space$(28)
''15791 TICO% = 4: AJU% = 0: CODCON% = Abs(CUE%)
''15792 IF CODCON%>0 THEN IF CODCON%<=NUPIE% THEN:GET #N8%,CODCON%:CUE$=P1$:DENO$=LEFT$(P2$,28)
''15793 '
''15798 Return
''15799 '
''15850 ENCU% = 0
''15852 For EMPL% = 1 To NUEMPL%
''15855 If LEGAI%(EMPL%) = VD% Then ENCU% = 1: GoTo 15870
''15860 Next EMPL%
''15862 GoTo 15880
''15865 '
''15870 GET #N5%,EMPL%
''15872 NOM$ = F2$: ESEMP% = Asc(F5$): GRU% = Asc(F6$): FRELI% = 0: If GRU% <= NUGRU% Then FRELI% = FRELIQ%(GRU%)
''15875 '
''15880 Return
''15885 '
''15900 NCLX% = Val(VDSS$): XCLX% = InStr(VDSS$, "\"): If XCLX% = 0 Then XCLX% = Len(VDSS$)
''15901 VDSS$ = Mid$(VDSS$, XCLX% + 1)
''15902 RGI% = Val(VDSS$): XCLX% = InStr(VDSS$, "\"): If XCLX% = 0 Then XCLX% = Len(VDSS$)
''15903 VDSS$ = Mid$(VDSS$, XCLX% + 1)
''15904 RGZ% = Val(VDSS$): XCLX% = InStr(VDSS$, "\"): If XCLX% = 0 Then XCLX% = Len(VDSS$)
''15905 VDSS$ = Mid$(VDSS$, XCLX% + 1)
''15906 OF$ = Space$(Val(VDSS$)): XCLX% = InStr(VDSS$, "\"): If XCLX% = 0 Then XCLX% = Len(VDSS$)
''15907 VDSS$ = Mid$(VDSS$, XCLX% + 1)
''15908 For KCLX% = 1 To NCL%
''15909 ACL%(KCL%) = Val(VDSS$)
''15910 XCLX% = InStr(VDSS$, "\"): If XCLX% = 0 Then XCLX% = Len(VDSS$)
''15911 VDSS$ = Mid$(VDSS$, XCLX% + 1)
''15912 Next KCL%
''15913 For KCLX% = 1 To NCL%
''15914 XCLX% = InStr(VDSS$, "\"): If XCLX% = 0 Then XCLX% = Len(VDSS$) + 1
''15915 ECL$(KCL%) = Left$(VDSS$, XCLX% - 1)
''15916 VDSS$ = Mid$(VDSS$, XCLX% + 1)
''15917 If Left$(ECL$(KCL%), 1) <> "<" Then GoTo 15920
''15918 ECL$(KCL%) = Left$(" " + Mid$(ECL$(KCL%), 2) + Space$(ACL%(KCL%)), ACL%(KCL%))
''15919 GoTo 15926
''15920 If Left$(ECL$(KCL%), 1) <> ">" Then GoTo 15923
''15921 ECL$(KCL%) = Right$(Space$(ACL%(KCL%)) + Mid$(ECL$(KCL%), 2) + " ", ACL%(KCL%))
''15922 GoTo 15926
''15923 If Left$(ECL$(KCL%), 1) = "^" Then ECL$(KCL%) = Mid$(ECL$(KCL%), 2)
''15924 While Len(ECL$(KCL%)) < ACL%(KCL%): ECL$(KCL%) = " " + ECL$(KCL%) + " ": Wend
''15925 ECL$(KCL%) = Left$(ECL$(KCL%), ACL%(KCL%))
''15926 Next KCL%
''15927 ECX$ = OF$ + Chr$(186): LI0$ = OF$ + Chr$(186): LI1$ = OF$ + Chr$(201): LI2$ = OF$ + Chr$(204): LI3$ = OF$ + Chr$(200)
''15928 For KCLX% = 1 To NCL%
''15929 ECX$ = ECX$ + ECL$(KCL%): If KCLX% < NCLX% Then ECX$ = ECX$ + Chr$(179) Else ECX$ = ECX$ + Chr$(186)
''15930 LI0$ = LI0$ + Space$(ACL%(KCL%)): If KCLX% < NCLX% Then LI0$ = LI0$ + Chr$(179) Else LI0$ = LI0$ + Chr$(186)
''15931 LI1$ = LI1$ + String$(ACL%(KCL%), 205): If KCLX% < NCLX% Then LI1$ = LI1$ + Chr$(209) Else LI1$ = LI1$ + Chr$(187)
''15932 LI2$ = LI2$ + String$(ACL%(KCL%), 205): If KCLX% < NCLX% Then LI2$ = LI2$ + Chr$(216) Else LI2$ = LI2$ + Chr$(185)
''15933 LI3$ = LI3$ + String$(ACL%(KCL%), 205): If KCLX% < NCLX% Then LI3$ = LI3$ + Chr$(207) Else LI3$ = LI3$ + Chr$(188)
''15934 Next KCL%
''15936 LOCATE RGI%, 1: Print LI1$
''15937 LOCATE RGI% + 1, 1: Print ECX$
''15938 LOCATE RGI% + 2, 1: Print LI2$
''15939 RGX% = RGI% + 3
''      While RGX% < RGZ%
''  LOCATE RGX%, 1: Print LI0$
''  RGX% = RGX% + 1
''      Wend
''15940 LOCATE RGZ%, 1
''      Print LI3$
''15941 Return
''15950 LOCATE 4, 1: Print Chr$(201) + String$(15, 205) + Chr$(203) + String$(32, 205) + Chr$(203) + String$(29, 205) + Chr$(187)
''15951 LOCATE 5, 1: Print Chr$(186) + " Legajo:" + Space$(7) + Chr$(186) + Space$(32) + Chr$(186) + " Grupo:" + Space$(22) + Chr$(186)
''15952 LOCATE 6, 1: Print Chr$(204)
''      LOCATE 6, 10: Print Chr$(209)
''      LOCATE 6, 17: Print Chr$(202)
''      LOCATE 6, 40: Print Chr$(209)
''15953 LOCATE 6, 45: Print Chr$(209)
''      LOCATE 6, 50: Print Chr$(202)
''      LOCATE 6, 54: Print Chr$(209)
''      LOCATE 6, 67: Print Chr$(209)
''      LOCATE 6, 80: Print Chr$(185)
''15954 LOCATE 21, 1: Print LI2$
''15956 LOCATE 22, 12: Print "Tope Retencion:"
''15958 RG% = 22: CLX% = 28: LCII% = 10: LKI% = 10: LD% = 2: TV% = 3: JTI% = 2: VDDD# = 0
''      GoSub 9000
''      On CT% GoTo 3900, 15959, 15960, 15960, 15959
''15959 Print Chr$(7);: GoTo 15958
''15960 TOPERET# = VDDD#
''15962 LOCATE 22, 1: Print LI0$
''15964 LOCATE 22, 12: Print "Total"
''15965 Return
''15966 '
16000 If TOPERET# > 0.5 Then If RMD#(0) > TOPERET# Then RMD#(0) = TOPERET#
16001 FORMUTRAB$ = FORMORI$: RIII% = 0: RID# = 0: If MCAL% = 4 Then GoSub 16600: GoTo 16050
16002 LIMI% = 0: If InStr(FORMUTRAB$, ">") = 0 Then GoTo 16005
16003 LIMI% = InStr(FORMUTRAB$, ">"): FORMU2$ = Mid$(FORMUTRAB$, LIMI% + 1): FORMUTRAB$ = Left$(FORMUTRAB$, LIMI% - 1)
16004 '
16005 PINI% = InStr(FORMUTRAB$, "@"): If PINI% < 1 Then GoTo 16025
16010 GoSub 16060              ' RESOLUCION SUMATORIA
16015 GoTo 16005
16020 '
16025 PINI% = InStr(FORMUTRAB$, "("): If PINI% < 1 Then GoTo 16045
16030 GoSub 16155              ' RESOLUCION PARENTESIS
16035 GoTo 16025
16040 '
16045 CALFIN% = 0: GoSub 16220 ' RESOLUCION FORMULA ALGEBRAICA
16046 If LIMI% = 0 Then GoTo 16050
16047 If LIMI% > 0 Then RIBAS# = RID#: FORMUTRAB$ = FORMU2$: RID# = 0: RIII% = 0: LIMI% = -1: GoTo 16005
16048 If RIBAS# > RID# Then RID# = RIBAS#
16049 '
16050 Return
16055 '
16060 IZQUI$ = "": If PINI% > 1 Then IZQUI$ = Left$(FORMUTRAB$, PINI% - 1): FORMUTRAB$ = Mid$(FORMUTRAB$, PINI%)
16065 VARI$ = Mid$(FORMUTRAB$, 2, 2): FORMUTRAB$ = Mid$(FORMUTRAB$, 4)
16070 LINF% = Val(Left$(FORMUTRAB$, 2)): FORMUTRAB$ = Mid$(FORMUTRAB$, 3)
16075 If Left$(FORMUTRAB$, 1) = "." Then FORMUTRAB$ = Mid$(FORMUTRAB$, 2): GoTo 16075
16080 LSUP% = Val(Left$(FORMUTRAB$, 2))
16085 If LSUP% < 10 Then DERE$ = Mid$(FORMUTRAB$, 2) Else DERE$ = Mid$(FORMUTRAB$, 3)
16090 RID# = 0
16095 For RANGO% = LINF% To LSUP%
16100 VALOR# = 0
16105 If VARI$ = "SB" Then VALOR# = SB(RANGO%): GoTo 16125
16110 If VARI$ = "PI" Then VALOR# = PII(RANGO%): GoTo 16125
16115 If VARI$ = "QT" Then VALOR# = QT(RANGO%): GoTo 16125
16120 If VARI$ = "RM" Then VALOR# = RMD#(RANGO%): GoTo 16125
16125 RID# = RID# + (Int(100 * VALOR# + 0.5)) / 100
16130 Next RANGO%
16135 RIII% = RIII% + 1: RI#(RIII%) = RID#: RISS$ = "RI" + Mid$(Str$(RIII%), 2)
16140 FORMUTRAB$ = IZQUI$ + RISS$ + DERE$
16145 Return
16150 '
16155 IZQUI$ = "": If PINI% > 1 Then IZQUI$ = Left$(FORMUTRAB$, PINI% - 1)
16160 FORMUTRAB$ = Mid$(FORMUTRAB$, PINI% + 1)
16165 PINI% = InStr(FORMUTRAB$, "("): If PINI% > 0 Then IZQUI$ = IZQUI$ + "(" + Left$(FORMUTRAB$, PINI% - 1): FORMUTRAB$ = Mid$(FORMUTRAB$, PINI% + 1): GoTo 16165
16170 PFIN% = InStr(FORMUTRAB$, ")"): If PFIN% < 1 Then PFIN% = 1 + Len(FORMUTRAB$)
16175 DERE$ = Mid$(FORMUTRAB$, PFIN% + 1)
16180 FORMUTRAB$ = Left$(FORMUTRAB$, PFIN% - 1)
16185 '
16190 CALFIN% = 1: GoSub 16220       ' RESOLUCION FORMULA SIN PARENTESIS
16195 '
16200 RIII% = RIII% + 1: RI#(RIII%) = RID#: RISS$ = "RI" + Mid$(Str$(RIII%), 2)
16205 FORMUTRAB$ = IZQUI$ + RISS$ + DERE$
16210 Return
16215 '
16220 POPE% = InStr(FORMUTRAB$, "^"): If POPE% > 0 Then OPE% = 5: GoTo 16280
16222 POPE% = InStr(FORMUTRAB$, "x"): If POPE% > 0 Then OPE% = 1: GoTo 16280
16225 POPE% = InStr(FORMUTRAB$, "X"): If POPE% > 0 Then OPE% = 1: GoTo 16280
16230 POPE% = InStr(FORMUTRAB$, "*"): If POPE% > 0 Then OPE% = 1: GoTo 16280
16235 POPE% = InStr(FORMUTRAB$, "/"): If POPE% > 0 Then OPE% = 2: GoTo 16280
16240 '
16245 POPE% = InStr(FORMUTRAB$, "+"): If POPE% > 0 Then OPE% = 3: GoTo 16280
16250 '
16255 POPE% = InStr(FORMUTRAB$, "-"): If POPE% > 0 Then OPE% = 4: GoTo 16280
16260 '
16265 ' NO QUEDAN OPERACIONES ARITMETICAS POR CONSIDERAR
16267 If Abs(RID#) < 0.005 Then If CALFIN% = 0 Then IZQUIE$ = "": DERIE$ = "": VALO1# = 0: ARGU2$ = FORMUTRAB$: OPE% = 3: CALFIN% = 1: GoTo 16460
16270 Return
16275 '
16280 ARG1$ = "": ARG2$ = ""
16285 For PP% = POPE% - 1 To 1 Step -1
16290 '
16295 CR$ = Mid$(FORMUTRAB$, PP%, 1)
16300 If CR$ = "^" Or CR$ = "x" Or CR$ = "X" Or CR$ = "*" Or CR$ = "/" Or CR$ = "+" Or CR$ = "-" Then GoTo 16320
16305 Next PP%
16310 PP% = 0
16315 '
16320 ARG1$ = Mid$(FORMUTRAB$, PP% + 1, POPE% - PP% - 1): IZQUIE$ = "": If PP% > 0 Then IZQUIE$ = Left$(FORMUTRAB$, PP%)
16325 '
16330 For PP% = POPE% + 1 To Len(FORMUTRAB$)
16335 CR$ = Mid$(FORMUTRAB$, PP%, 1)
16340 If CR$ = "^" Or CR$ = "x" Or CR$ = "X" Or CR$ = "*" Or CR$ = "/" Or CR$ = "+" Or CR$ = "-" Then GoTo 16360
16345 Next PP%
16350 PP% = 1 + Len(FORMUTRAB$)
16355 '
16360 ARG2$ = Mid$(FORMUTRAB$, POPE% + 1, PP% - POPE% - 1): DERIE$ = Mid$(FORMUTRAB$, PP%)
16365 '
16370 ARGU1$ = "": For PP% = 1 To Len(ARG1$): CRII% = Asc(Mid$(ARG1$, PP%)): If (CRII% > 47 Or CRII% = 46) Then ARGU1$ = ARGU1$ + Chr$(CRII%)
16375 Next PP%
16380 '
16385 ARGU2$ = "": For PP% = 1 To Len(ARG2$): CRII% = Asc(Mid$(ARG2$, PP%)): If (CRII% > 47 Or CRII% = 46) Then ARGU2$ = ARGU2$ + Chr$(CRII%)
16390 Next PP%
16395 '
16400 VALO1# = 0
16405 If Left$(ARGU1$, 2) = "AE" Then VALO1# = AE1: GoTo 16460
16406 If Left$(ARGU1$, 2) = "QM" Then VALO1# = QM#: GoTo 16460
16407 If Left$(ARGU1$, 2) = "RG" Then VALO1# = RG1#: GoTo 16460
16408 If Left$(ARGU1$, 2) = "AC" Then VALO1# = AC: GoTo 16460
16410 If Left$(ARGU1$, 2) = "DJ" Then VALO1# = DJ0#: GoTo 16460
16415 INDI% = Val(Mid$(ARGU1$, 3)): INDI% = 0: If INDI% >= 0 Then If INDI% < 100 Then INDI% = Int(INDI)
16420 If Left$(ARGU1$, 2) = "RI" Then VALO1# = RI#(INDI%): GoTo 16460
16425 If Left$(ARGU1$, 2) = "SB" Then VALO1# = SB(INDI%): GoTo 16460
16430 If Left$(ARGU1$, 2) = "QT" Then VALO1# = QT(INDI%): GoTo 16460
16435 If Left$(ARGU1$, 2) = "RM" Then VALO1# = RMD#(INDI%): GoTo 16460
16440 If Left$(ARGU1$, 2) = "PI" Then VALO1# = PII(INDI%): GoTo 16460
      If Left$(ARGU1$, 2) = "AM" Then GoSub 16800: VALO1# = VALOACU#: GoTo 16460 ' ACUMULADO MENSUAL DEL CODIGO
16445 '
16450 VALO1# = Val(ARGU1$)
16455 '
16460 VALO2# = 0
16465 If Left$(ARGU2$, 2) = "AE" Then VALO2# = AE1: GoTo 16520
16466 If Left$(ARGU2$, 2) = "QM" Then VALO2# = QM#: GoTo 16520
16467 If Left$(ARGU2$, 2) = "RG" Then VALO2# = RG1#: GoTo 16520
16468 If Left$(ARGU2$, 2) = "AC" Then VALO2# = AC: GoTo 16520
16470 If Left$(ARGU2$, 2) = "DJ" Then VALO2# = DJ0#: GoTo 16520
16475 INDI% = Val(Mid$(ARGU2$, 3)): INDI% = 0: If INDI% >= 0 Then If INDI% < 100 Then INDI% = Int(INDI)
16480 If Left$(ARGU2$, 2) = "RI" Then VALO2# = RI#(INDI%): GoTo 16520
16485 If Left$(ARGU2$, 2) = "SB" Then VALO2# = SB(INDI%): GoTo 16520
16490 If Left$(ARGU2$, 2) = "QT" Then VALO2# = QT(INDI%): GoTo 16520
16495 If Left$(ARGU2$, 2) = "RM" Then VALO2# = RMD#(INDI%): GoTo 16520
16500 If Left$(ARGU2$, 2) = "PI" Then VALO2# = PII(INDI%): GoTo 16520
      If Left$(ARGU2$, 2) = "AM" Then GoSub 16800: VALO2# = VALOACU#: GoTo 16520 ' ACUMULADO MENSUAL DEL CODIGO
16505 '
16510 VALO2# = Val(ARGU2$)
16515 '
16520 RID# = 0
16525 If OPE% = 1 Then RID# = VALO1# * VALO2#: GoTo 16550
16530 If OPE% = 2 Then If Abs(VALO2#) >= 0.005 Then RID# = VALO1# / VALO2#: GoTo 16550
16535 If OPE% = 3 Then RID# = VALO1# + VALO2#: GoTo 16550
16540 If OPE% = 4 Then RID# = VALO1# - VALO2#: GoTo 16550
16542 If OPE% = 5 Then RID# = VALO1# ^ VALO2#: GoTo 16550
16545 '
16550 RIII% = RIII% + 1: RI#(RIII%) = RID#: RISS$ = "RI" + Mid$(Str$(RIII%), 2)
16555 FORMUTRAB$ = IZQUIE$ + RISS$ + DERIE$
16560 GoTo 16220
16565 '
16600 For F% = 0 To 10: VALESC(F%) = 0: Next F%
16602 F% = 0
16605 VALESC(F%) = Val(FORMUTRAB$)
16607 FII1% = InStr(FORMUTRAB$, "-"): If FII1% < 1 Then GoTo 16615
16610 FORMUTRAB$ = Mid$(FORMUTRAB$, FII1% + 1): If F% < 10 Then F% = F% + 1: GoTo 16605
16612 '
16615 FMAX% = F%
16617 If QT0 < 0 Or QT0 > FMAX% Then RID# = 0: GoTo 16625
16620 RID# = CDbl(VALESC(QT0))
16625 Return
16627 '

16800 If N10% < 1 Then GoSub 5960
      UREDTLQ% = LOF(N10%) / 24
      VALOACU# = 0
      For IJKI% = 1 To Len(PULIQ$(EMPL%)) Step 2
 REDTLQ% = CVI(Mid$(PULIQ$(EMPL%), IJKI%, 2))
 If REDTLQ% < 1 Or REDTLQ% > UREDTLQ% Then GoTo 16819
 ''GET #N10%,REDTLQ%
 If CVI(J0$) = 0 Then GoTo 16819
 If CVI(J0$) = MARCA% Then GoTo 16819 ' ES LIQUIDACION ACTIVA
 LGJ8% = CVI(J1$): If LGJ8% <> LEGA% Then GoTo 16819
 CODI8% = CVI(J2$): If (CODI8% < 1 Or CODI8% > 199) Then GoTo 16819
 If CODI8% Mod 100 <> INDI% Then GoTo 16819
 IMPORT# = CVD(J4$)
 VALOACU# = VALOACU# + IMPORT#
16819 Next IJKI%
      '
16829 Return
      '
17000 If GRULI% <> 0 Then If GRUII% <> GRULI% Then If Abs(SUMD#) < 0.5 Then GoTo 17140
17001 If CONTANT% > 0 Then If EJU% > 0 Then GoTo 17032
17002 If LSUE% > 0 Then GoTo 17200
17003 For V% = 1 To NUCON%
17005 If ESTCON%(V%) < 1 Then GoTo 17030
17007 If FORCONI%(V%) <> 0 Then GoTo 17030
17008 If QLCON%(V%) <> 0 Then If QLCON%(V%) <> TILIQ% Then GoTo 17030
17010 If CARCON%(V%) = 3 Or CARCON%(V%) = FRELI% Then GoTo 17017
17011 For X% = 1 To 96 Step 2
17012 X1% = CVI(Mid$(ACTI$, X%)): If X1% = 0 Or X1% > V% Then GoTo 17015
17013 If X1% = V% Then GoTo 17017           ' PREASIGNADO
17014 Next X%
17015 GoTo 17030
17016 '
17017 For V1% = 1 To CONTA%
17018   If CTA%(V1%) = V% Then
   If IM#(V1%) > 0.005 Then GoTo 17030 ' DESCARTAR SI YA FUE INFORMADO
   FORMORI$ = FORCON$(V%): MCAL% = FORCONI%(V%): QT0 = CANTJ(V1%)
   GoSub 16000
   IM#(V1%) = RID#: RMD#(V%) = RID#
   GoTo 17030
 End If
17019 Next V1%
17020 FORMORI$ = FORCON$(V%): MCAL% = FORCONI%(V%): QT0 = CANTJ(V1%)
      GoSub 16000
17025 CONTA% = CONTA% + 1: CTA%(CONTA%) = V%: CANTJ(CONTA%) = 0: IM#(CONTA%) = RID#
17027 QT(V%) = 0: RMD#(V%) = RID#
17030 Next V%
17031 '
17032 For V% = 1 To CONTA%
17033 If CTA%(V%) < 1 Or CTA%(V%) > 99 Then GoTo 17038
17034 If INFO%(V%) > 0 Then GoTo 17038        ' IMPORTE INFORMADO
17035 If RCLCON%(CTA%(V%)) < 1 Then If (EJU% > 0 Or FORCONI%(CTA%(V%)) = 1) Then GoTo 17038
17036 FORMORI$ = FORCON$(CTA%(V%)): MCAL% = FORCONI%(CTA%(V%)): QT0 = CANTJ(V%): GoSub 16000
17037 IM#(V%) = RID#: RMD#(CTA%(V%)) = RID#
17038 Next V%
17039 '
17040 ''GoSub 15500
17045 '
17050 If ESEMP% <> 1 Then GoTo 17100
17051 If CONTANT% > 0 Then If EJU% > 0 Then GoTo 17082
17052 For V% = 1 To NURET%
17055 If ESTRET%(V%) < 1 Then GoTo 17080
17057 If FORRETI%(V%) = 1 Then GoTo 17080   ' IMPORTE INFORMADO
17058 If QLRET%(V%) <> 0 Then If QLRET%(V%) <> TILIQ% Then GoTo 17080
17060 If CARRET%(V%) = 3 Or CARRET%(V%) = FRELI% Then GoTo 17067
17061 For X% = 1 To 96 Step 2
17062 X1% = CVI(Mid$(ACTI$, X%)): If X1% = 0 Or X1% > V% + 200 Then GoTo 17065
17063 If X1% = V% + 200 Then GoTo 17067         ' PREASIGNADO
17064 Next X%
17065 GoTo 17080
17066 '
17067 For V1% = 1 To CONTA%
17068 If CTA%(V1%) = V% + 200 Then GoTo 17080 ' DESCARTAR SI YA FUE INFORMADO
17069 Next V1%
17070 FORMORI$ = FORRET$(V%): MCAL% = FORRETI%(V%): QT0 = CANTJ(V1%): If RCLRET%(V%) <> 2 Then GoSub 16000 Else GoTo 17080
17075 CONTA% = CONTA% + 1: CTA%(CONTA%) = 200 + V%: CANTJ(CONTA%) = 0: IM#(CONTA%) = RID#
17077 RT#(V%) = RID#
17080 Next V%
17081 '
17082 For V% = 1 To CONTA%
17083 If CTA%(V%) < 201 Or CTA%(V%) > 299 Or CTA%(V%) = RTAN% + 200 Then GoTo 17088
17084 If INFO%(V%) > 0 Then GoTo 17088        ' IMPORTE INFORMADO
17085 If RCLRET%(CTA%(V%) - 200) <> 1 Then If (EJU% > 0 Or FORRETI%(CTA%(V%) - 200) = 1) Then GoTo 17088
17086 FORMORI$ = FORRET$(CTA%(V%) - 200): MCAL% = FORRETI%(CTA%(V%) - 200): QT0 = CANTJ(V%)
      '
      If RMD#(0) > RELIMI# Then   ' CAMBIA PORCENTAJE ANSAL
 '
 PPXXYY% = InStr(FORMORI$, "0.0255")
 If PPXXYY% > 0 Then
   Mid$(FORMORI$, PPXXYY%, 6) = "0.0240"
 End If
 PPXXYY% = InStr(FORMORI$, "0.0045")
 If PPXXYY% > 0 Then
   Mid$(FORMORI$, PPXXYY%, 6) = "0.0060"
 End If
 '
 PPXXYY% = InStr(FORMORI$, "0.0270")
 If PPXXYY% > 0 Then
   Mid$(FORMORI$, PPXXYY%, 6) = "0.0255"
 End If
 PPXXYY% = InStr(FORMORI$, "0.0030")
 If PPXXYY% > 0 Then
   Mid$(FORMORI$, PPXXYY%, 6) = "0.0045"
 End If
 '
      End If
      '
      GoSub 16000      ' PRODUCE EL RECALCULO
      '
17087 IM#(V%) = RID#: RT#(CTA%(V%) - 200) = RID#
17088 Next V%
17089 '
17100 If QUASI% = 0 Or ESEMP% <> 1 Then GoTo 17140
17101 If CONTANT% > 0 Then If EJU% > 0 Then GoTo 17140
17102 For V% = 1 To NUASI%
17105 If ESTASI%(V%) < 1 Then GoTo 17130
17112 X1% = Asc(Mid$(ASFAM$, V%)): If X1% < 1 Then GoTo 17130
17116 '
17117 For V1% = 1 To CONTA%: If CTA%(V1%) <> V% + 400 Then GoTo 17119
17118 If EJU% > 0 Then GoTo 17130 Else IM#(V1%) = CDbl(X1% * VALASI(V%)): GoTo 17130
17119 Next V1%
17120 RID# = CDbl(X1% * VALASI(V%))
17125 CONTA% = CONTA% + 1: CTA%(CONTA%) = 400 + V%: CANTJ(CONTA%) = X1%: IM#(CONTA%) = RID#
17130 Next V%
17131 '
17140 ''GoSub 15500: GoSub 17400: GoSub 15500  ' ABSORBER PRESTAMOS Y ANTICIPOS
17145 '
17156 For V% = 1 To CONTA%
17157 For V1% = V% + 1 To CONTA%
17158 If CTA%(V%) <= CTA%(V1%) Then GoTo 17164
17159 ''SWAP CTA%(V%), CTA%(V1%)
17160 ''SWAP CANTJ(V%), CANTJ(V1%)
17161 ''SWAP IM#(V%), IM#(V1%)
17162 ''SWAP INFO%(V%), INFO%(V1%)
17164 Next V1%
17166 Next V%
17168 '
17200 Return
17205 '
17400 FLG% = 0: PRESTA# = 0: NETO# = SUMD# - SUMH#: If RCLRET%(17) <> 2 Then GoTo 17500
17402 GoSub 17450: If QUI% <> 0 And TILIQ% <> QUI% Then GoTo 17500 ' PRESTAMOS QUE NO SE DESCUENTAN EN QUI% QUINCENA
17403 For JJJ% = 1 To CONTA%
17404 If CTA%(JJJ%) = 200 + RTPR% Then If IM#(JJJ%) > 0 Then GoTo 17500 ' DESCUENTA UNA SOLA VEZ
17405 Next JJJ%
17406 '
17407 If LSUE% > 0 Then GoTo 17595
17408 For JJJ% = 1 To CAPRES%
17409 ''GET #N14%,JJJ%:IF LEGA%<>CVI(W1$) THEN 17430
17410 CUOPEN% = CVI(W6$): If CUOPEN% < 1 Or CVI(W4$) < 1 Or CVD(W5$) < 1 Then GoTo 17430
17412 BASACT# = SB(1) + SB(2) + SB(4) + SB(5)
17415 PRESTA# = (CVD(W3$) / CVI(W4$)) * (BASACT# / CVD(W5$))
17420 If PRESTA# > NETO# Then PRESTA# = 0: GoTo 17500
17422 ''LSET W6$=MKI$(CUOPEN%-1):PUT #N14%,JJJ%:IF FLG%<>0 THEN 17595
17423 If PRESTA# >= 0.005 Then CONTA% = CONTA% + 1: CTA%(CONTA%) = 200 + RTPR%: CANTJ(CONTA%) = 0: IM#(CONTA%) = PRESTA#: INFO%(CONTA%) = 0
17425 ' GOTO 17500
17430 Next JJJ%
17432 GoTo 17500
17435 '
17450 ' BUSCO EN PRESTAMO LA QUINCENA QUE CORRESPONDE DESCONTAR EN QUI%
17452 For PW% = 1 To CAPRES%
17454 ''GET #N14%,PW%
17456 If CVI(W6$) = 0 Then GoTo 17460
17457 If CVI(W1$) <> LEGA% Then GoTo 17460
17458 QUI% = Asc(W7$): GoTo 17462
17460 Next PW%
17462 Return
17464 '
17500 If YANTI% > 0 Then GoTo 17595
17501 ANTIC# = 0: NETO# = SUMD# - SUMH# - PRESTA#: If RCLRET%(16) <> 2 Then GoTo 17595
17502 If LSUE% > 0 Then GoTo 17595
17505 For JJJ% = 1 To CANANT%
17510 ''GET #N13%,JJJ%
17525 If LEGA% <> CVI(V2$) Then GoTo 17560
17527 If V4$ = Chr$(1) Then GoTo 17560 ' SALTEA LOS DESCONTADOS
17530 IMPOR# = CVD(V3$)
17535 AUXIMP# = ANTIC# + (Int(100 * IMPOR# + 0.5)) / 100
17540 If NETO# < AUXIMP# Then GoTo 17560
17545 ANTIC# = ANTIC# + (Int(100 * IMPOR# + 0.5)) / 100
17550 LSet V4$ = Chr$(1)
17555 ''PUT #N13%,JJJ%
17560 Next JJJ%
17565 '
17570 YANTI% = 1: If ANTIC# >= 0.005 Then CONTA% = CONTA% + 1: CTA%(CONTA%) = 200 + RTAN%: CANTJ(CONTA%) = 0: IM#(CONTA%) = ANTIC#: INFO%(CONTA%) = 0
17595 FLG% = 0: Return
17598 '
18100 If CTA%(RGL%) > 0 Then If CTA%(RGL%) < 100 Then QT(CTA%(RGL%)) = 0: RMD#(CTA%(RGL%)) = 0
18105 If CTA%(RGL%) > 200 Then If CTA%(RGL%) < 300 Then RT#(CTA%(RGL%) - 200) = 0
18110 Return
18115 '
18200 For RGL1% = RGL% + 1 To CONTA%
18205 CTA%(RGL1% - 1) = CTA%(RGL1%): IM#(RGL1% - 1) = IM#(RGL1%): CANTJ(RGL1% - 1) = CANTJ(RGL1%): INFO%(RGL1% - 1) = INFO%(RGL1%)
18210 Next RGL1%
18215 CONTA% = CONTA% - 1: If RGL% <= CAUTO% Then CAUTO% = CAUTO% - 1
18222 GoSub 18250
18223 CTA%(CONTA% + 1) = 0: IM#(CONTA% + 1) = 0: CANTJ(CONTA% + 1) = 0: INFO%(CONTA% + 1) = 0
18225 Return
18230 '
18250 RGFII% = RGL% - INFE% + 1
18255 For K1% = RGFII% To RLB% + 1
18260 CAN(K1%) = CAN(K1% + 1): CTB$(K1%) = CTB$(K1% + 1): DEX$(K1%) = DEX$(K1% + 1): IMPO#(K1%) = IMPO#(K1% + 1): SIG%(K1%) = SIG%(K1% + 1)
18265 Next K1%
18270 K1% = INFE% + RLB%: GoSub 15700: CAN(RLB% + 1) = CANTUSI: CTB$(RLB% + 1) = CUE$: DEX$(RLB% + 1) = DENO$: IMPO#(RLB% + 1) = IMPOD#
18275 Return
18280 '
18300 ''CAMBENC% = 0: CAMBDIA% = 0: If LSUE% > 0 Then Print Chr$(7);: GoTo 18490
18301 ''Call BLIMESS("  ESPERE  !!!  ")
      GoSub 18800
18314 '
18316 If CONTA% <> CONTANT% Then CAMBDIA% = 1: GoTo 18330
18318 For I% = 1 To CONTA%
18320 ''GET #N6%,REGINI%+I%-1
18322 If CVI(D2$) <> CTA%(I%) Then CAMBDIA% = 1: GoTo 18330
18323 If CVS(D3$) <> CANTJ(I%) Then CAMBDIA% = 1: GoTo 18330
18324 If CVD(D4$) <> IM#(I%) Then CAMBDIA% = 1: GoTo 18330
18326 Next I%
18328 '
18330 ''Call BLIMESS("-")
      If CAMBDIA% = 0 Then GoTo 18490
18335 '
18390 If CONTA% = CONTANT% Then GoTo 18440
18394 If CONTA% > CONTANT% Then GoTo 18420
18396 DIFE% = CONTANT% - CONTA%
18398 For I% = REGINI% + CONTANT% To MAXDIA%
18400 ''GET #N6%,I%:PUT #N6%,I%-DIFE%
18402 Next I%
18404 '
18406 LONGI% = LOF(N6%) / 16
18407 For I% = MAXDIA% - DIFE% + 1 To LONGI%
18408 ''LSET D1$=MKI$(0):PUT #N6%,I%
18410 Next I%
18412 GoTo 18440
18414 '
18420 DIFE% = CONTA% - CONTANT%
18422 For I% = MAXDIA% To REGINI% + CONTANT% Step -1
18424 ''GET #N6%,I%:PUT #N6%,I%+DIFE%
18426 Next I%
18428 '
18430 LONGI% = LOF(N6%) / 16
18431 For I% = MAXDIA% + DIFE% + 1 To LONGI%
18432 ''LSET D1$=MKI$(0):PUT #N6%,I%
18434 Next I%
18436 '
18440 For I% = 1 To CONTA%
18442 ''GET #N6%,REGINI%+I%-1
18444 LSet D1$ = MKI$(LEGA%)
18446 LSet D2$ = MKI$(CTA%(I%))
18448 LSet D3$ = MKS$(CANTJ(I%))
18450 LSet D4$ = MKD$(IM#(I%))
18458 ''PUT #N6%,REGINI%+I%-1
18462 Next I%
18464 '
18490 If N6% > 0 Then Close #N6%: N6% = 0
      If N13% > 0 Then Close #N13%: N13% = 0
      If N14% > 0 Then Close #N14%: N14% = 0
18492 ''Call BLIMESS("-")
18493 Return
18495 '
18496 '
18500 GoSub 18800
      If CONTA% < 1 Or LSUE% > 0 Then
 If N6% > 0 Then Close #N6%: N6% = 0
 If N13% > 0 Then Close #N13%: N13% = 0
 If N14% > 0 Then Close #N14%: N14% = 0
 GoTo 18596
      End If
18501 ''Call BLIMESS(" GRABANDO  !!! ")
18548 '
18552 For I% = 1 To CONTA%
18554 ''GET #N6%,I%+MAXDIA%
18557 LSet D1$ = MKI$(LEGA%)
18560 LSet D2$ = MKI$(CTA%(I%))
18563 LSet D3$ = MKS$(CANTJ(I%))
18566 LSet D4$ = MKD$(IM#(I%))
18578 ''PUT #N6%,I%+MAXDIA%
18581 Next I%
18584 '
18587 LONGI% = LOF(N6%) / 16
18588 For I% = MAXDIA% + CONTA% + 1 To LONGI%
18590 ''LSET D1$=MKI$(0):PUT #N6%,I%
18593 Next I%
18595 If N6% > 0 Then Close #N6%: N6% = 0
      If N13% > 0 Then Close #N13%: N13% = 0
      If N14% > 0 Then Close #N14%: N14% = 0
18596 ''Call BLIMESS("-")
18597 Return
18599 '
18600 ''If LSUE% > 0 Then Print Chr$(7);: GoTo 18694
18605 ''Call BLIMESS(" GRABANDO  !!! ")
      GoSub 18800
      If CONTA% < 1 Then
 If N6% > 0 Then Close #N6%: N6% = 0
 If N13% > 0 Then Close #N13%: N13% = 0
 If N14% > 0 Then Close #N14%: N14% = 0
 GoTo 18694
      End If
18650 For I% = MAXDIA% To 1 Step -1
18652 ''GET #N6%,I%:IF CVI(D1$)<=LEGA% THEN 18662
18654 ''PUT #N6%,I%+CONTA%
18656 Next I%
18658 I% = 0
18660 '
18662 REGINI% = I%
18664 For I% = 1 To CONTA%
18666 ''GET #N6%,I%+REGINI%
18667 LSet D1$ = MKI$(LEGA%)
18668 LSet D2$ = MKI$(CTA%(I%))
18669 LSet D3$ = MKS$(CANTJ(I%))
18670 LSet D4$ = MKD$(IM#(I%))
18682 ''PUT #N6%,I%+REGINI%
18686 Next I%
18688 '
18690 LONGI% = LOF(N6%) / 16
18691 For I% = MAXDIA% + CONTA% + 1 To LONGI%
18692 ''LSET D1$=MKI$(0):PUT #N6%,I%
18693 Next I%
18694 If N6% > 0 Then Close #N6%: N6% = 0
      If N13% > 0 Then Close #N13%: N13% = 0
      If N14% > 0 Then Close #N14%: N14% = 0

18695 ''Call BLIMESS("-")
18696 Return
18699 '
18800 KQ% = 0
18803 For I% = 1 To CONTA%
18806 If Abs(CTA%(I%)) > 0 Then If (Abs(CANTJ(I%)) >= 0.005 Or Abs(IM#(I%)) >= 0.005) Then KQ% = KQ% + 1: CTA%(KQ%) = CTA%(I%): CANTJ(KQ%) = CANTJ(I%): IM#(KQ%) = IM#(I%)
18809 Next I%
18812 CONTA% = KQ%
18815 '
18820 For I% = 1 To CONTA%
18821 If CTA%(I%) < 0 Or CTA%(I%) = 200 + RTAN% Then GoTo 18847 ' ADMITE DUPLICACION DE DESTAJO Y ANTICIPOS
18822 CTGD% = CTA%(I%) Mod 100
18823 If (CTDG% Mod 2) = 1 Then GoTo 18847 ' ADMITE DUPLICACION DE AJUSTES
18824 For I1% = 1 To I% - 1
18825 If CTA%(I1%) <> CTA%(I%) Then GoTo 18845
18827 For I2% = I% + 1 To CONTA%
18830 CTA%(I2% - 1) = CTA%(I2%): CANTJ(I2% - 1) = CANTJ(I2%): IM#(I2% - 1) = IM#(I2%)
18833 Next I2%
18835 CONTA% = CONTA% - 1
18840 GoTo 18820
18842 '
18845 Next I1%
18847 Next I%
18848 '
18850 Return
18855 '
19000 For X% = 1 To CONTA%
        CTA%(X%) = 0
        IM#(X%) = 0
        CANTJ(X%) = 0
        INFO%(X%) = 0
      Next X%
19002 For X% = 1 To 16
        CTB$(X%) = ""
        DEX$(X%) = ""
        CAN(X%) = 0
        IMPO#(X%) = 0
        SIG%(X%) = 0
      Next X%
19004 For X% = 1 To NUCON%
        QT(X%) = 0
        RMD#(X%) = 0
      Next X%
19006 For X% = 1 To NURET%
        RT#(X%) = 0
      Next X%
19010 CONTA% = 0
19015 Return
      '
      '
19050 For X% = INFE% To CONTOT% + 1
19052 If X% - INFE% > RLB% Then GoTo 19067
19055 NR1% = X% - INFE% + 9
19056 ''LOCATE NR1%, 2: Print Space$(79);
19057 ''LOCATE NR1%, 10: Print Chr$(179)
19058 ''LOCATE NR1%, 40: Print Chr$(179)
19059 ''LOCATE NR1%, 45: Print Chr$(179)
19060 ''LOCATE NR1%, 54: Print Chr$(179)
19061 ''LOCATE NR1%, 67: Print Chr$(179)
19062 ''LOCATE NR1%, 80: Print Chr$(186)
19065 Next X%
19067 ''LOCATE 5, 10: Print "      "
      ''LOCATE 5, 19: Print Space$(30)
      ''LOCATE 5, 59: Print Space$(21)
19068 ''LOCATE 22, 55: Print "            "
19069 ''LOCATE 22, 68: Print "            "
19072 Return
19075 '
20000 EPL% = EPL% + 1
20005 If EPL% > NUEMPL% Then GoTo 3900
20010 ''GET #N5%,EPL%:ESEMP%=ASC(F5$):GRU%=ASC(F6$)
20015 If ESEMP% < 1 Or ESEMP% > 2 Or GRUII% <> GRULI% Then GoTo 20000
20020 EMPL% = EPL%: LEGA% = CVI(F1$): NOM$ = F2$: FRELI% = 0: If GRUII% <= NUGRU% Then FRELI% = FRELIQ%(GRUII%)
20022 ''LOCATE 5, 10: Print LEGA%;
      ''LOCATE 5, 19: Print NOM$
      ''LOCATE 5, 59: Print Left$(GRU$(GRU%), 21)
20025 Return
20030 '

25001 C1$ = Mid$(C0, 1, 30)
      C2$ = Mid$(C0, 31, 1)
      C3$ = Mid$(C0, 32, 1)
      C4$ = Mid$(C0, 33, 1)
      C5$ = Mid$(C0, 34, 1)
      C6$ = Mid$(C0, 35, 1)
      C7$ = Mid$(C0, 36, 2)
      C8$ = Mid$(C0, 38, 20)
      C9$ = Mid$(C0, 58, 16)
      C10$ = Mid$(C0, 74, 2)
      C11$ = Mid$(C0, 76, 16)
      C12$ = Mid$(C0, 92, 12)
      C13$ = Mid$(C0, 104, 2)
      C14$ = Mid$(C0, 106, 1)
      C15$ = Mid$(C0, 107, 1)
      C16$ = Mid$(C0, 108, 1)
      C17$ = Mid$(C0, 109, 19)
Return
      '

30300 FE$ = FECHATEX$(FEX)
30440 Return
30450 '
30500 FEX = FECHANUM(FE$)
      Return

'30500 A%=0
'30510 IF ((LEN(FE$)=8)AND(MID$(FE$,3,1)="/")AND(MID$(FE$,6,1)="/")) THEN
'            GOTO 30530
'      END IF
'30520 A%=1:RETURN
'30530 FOR II%=1 TO 3
'30540 FEX(II%)=VAL(MID$(FE$,3*II%-2,2))
'30550 IF FEX(II%)<1 THEN
'            GOTO 30600
'      END IF
'30560 NEXT II%
'30570 IF ((FEX(1)>31)OR(FEX(2)>12)) THEN
'            GOTO 30600
'      END IF
'30580 FEX=FEX(1)+32*FEX(2)+32*13*(FEX(3)-80)
'30590 RETURN
'30600 A%=2:RETURN

31000 ''Cls
      ''Call SETCOL(1, 1)
      ''Call ENCAPAN(0, S1$)
      Return
   
64000 Resume Next

65000 CODERR = Err: LINERR = Erl
      ''Call CAPERR(CODERR, LINERR)
      If LINERR = 1 Then Resume Next
      Resume 0
      Close: End


End Sub

Sub LIBROSUEL()
     '
Dim AJCON(100), AJRET(100), AJASI(100)
Dim CO%(100), IMCO#(100), CACO(100), ACTHOR%(100)
Dim HOR%(100)
Dim L$(10)
Dim NOM1$(24), PARE1$(24), NACI1$(24), ESTU1$(24)
     Dim ACTI%(32767)
     Dim CON3(100), AJCON3(100)
     Dim RET3(100), AJRET3(100)
     Dim ASI3(100), AJASI3(100)
     '
     STPIMP05.Show 1
     If STPIMP05.Label1 <> "OK" Then
        Exit Sub
     End If
     NUHOJ% = Val(STPIMP05.Text6) - 1
     '
     Call SETPRINTER(132)
     '
     Call CARGAFORMULAS       ' CARGAR CONCEPTOS y RETENCIONES
     Call CARGLIQACT          ' datos de liquidacion activa
     ' PRESENTAR DATOS DE LIQ.ACTIVA Y OPCION CONTINUAR O CANCELAR
     '
     For U& = 1 To ULTREG&("DETLIQUI")
       XX$ = REGLEIDO$("DETLIQUI", U&)
       LGX% = CVI(Left$(XX$, 2))
       If LGX% > 0 Then
         ACTI%(LGX%) = 1
       End If
     Next U&
     '
     'JJ& = 0
     'REBLA$ = REGBLAN$("LINFHOR")
     '
1000 For KU& = 1 To ULTREG&("PADREMP")
       PP0$ = REGLEIDO$("PADREMP", KU&)
       LEGAX% = CVI(Left$(PP0$, 2))
       If ACTI%(LEGAX%) > 0 Then
         Call CARGEMP(LEGAX%)
1020     GRUPO% = GRUP_LEG%
1025     NOM$ = AJUSTI$(APNO_LEG$, 30)
         LEGA% = LEGAX%
         SECCI$ = FORMATNUM$(SECC_LEG, "F7.3")
         CATE$ = AJUSTI$(DCAT_LEG$, 10)
         INGRE$ = FECHATEX$(FEIN_LEG%)
         AFIL$ = CUIL_LEG$
         ECIV$ = AJUSTI$(ECIV_LEG$, 10)
         FEGRE$ = FECHATEX$(FEGR_LEG%)
         
'         DOCU$
'         INGRE$ = FE$
'         TARE$
'         CAJA$
'         DOMI$
'         LO$
'         NACIO$
'         ECIV$
         '
1050     TOTAL# = 0: TOHA# = 0: TODE# = 0: TOASI# = 0: SUMHOR = 0: TOEXTR# = 0: HEXTR = 0: NURENGLO% = 0
1055     For I% = 1 To NUCON%: CON3(I%) = 0: AJCON3(I%) = 0: Next I%
1056     For I% = 1 To NURET%: RET3(I%) = 0: AJRET3(I%) = 0: Next I%
1057     For I% = 1 To NUASI%: ASI3(I%) = 0: AJASI3(I%) = 0: Next I%
         CTPAR% = 0
1058     '
         RFF$ = RECFILT$("DETLIQUI", 1, MKI$(LEGAX%))
         For KKL& = 1 To Len(RFF$) Step 4
           REGI& = CVS(Mid$(RFF$, KKL&, 4))
           JJ$ = REGLEIDO$("DETLIQUI", REGI&)
1062       LEGAL% = CVI(Left$(JJ$, 2))
           CO1% = CVI(Mid$(JJ$, 3, 2))
           CACO1 = CVS(Mid$(JJ$, 5, 4))
           IMCO1# = (Int(100 * CVD(Mid$(JJ$, 9, 8)) + 0.5)) / 100
1065       If CO1% < 1 Then GoTo 1137
1076       R1% = CO1% Mod 100: RR% = Int(CO1% / 200)
1079       '
1080       NURENGLO% = NURENGLO% + 1
1100       CO%(NURENGLO%) = CO1%
1107       IMCO#(NURENGLO%) = IMCO1#
1108       CACO(NURENGLO%) = CACO1
1110       ACTHOR%(NURENGLO%) = 0
1116       RR% = Int(CO1% / 200)
1120       On RR% + 1 GoTo 1125, 1130, 1135
1125       TOHA# = TOHA# + (Int(100 * IMCO1# + 0.5)) / 100
1126       If HOR%(R1%) = 1 Then
             ACTHOR%(NURENGLO%) = 1
           End If
1128       GoTo 1137
           '
1130       TODE# = TODE# + (Int(100 * IMCO1# + 0.5)) / 100
           GoTo 1137
           '
1135       TOASI# = TOASI# + (Int(100 * IMCO1# + 0.5)) / 100
           CTPAR% = 1
           '
1137     Next KKL&
         '
1150     NUPAR% = 0: RENGLADI% = 0
         If CTPAR% > 0 Then GoSub 8000 ' CARGAR PARIENTES
1152     If NUPAR% > 0 Then RENGLADI% = 4 + NUPAR%
1155     '
1300     If Abs(TOHA#) < 0.05 And Abs(TODE#) < 0.05 And Abs(TOASI#) < 0.05 Then
            GoTo 1900
         End If
         '
1302     PCURRY = XVALO(Printer.CurrentY)
         If PCURRY > 0 Then
           INTERLIXI = PCURRY / NR%
           RENMAX% = (Printer.ScaleHeight - 750) * NR% / PCURRY
           If NR% + NURENGLO% + RENGLADI% > RENMAX% Then
             Printer.NewPage
             NR% = 0
           End If
         End If
         '
1305     If NR% = 0 Then
            GoSub 25000
         End If
1310     GoSub 25700
         GoSub 25100
1315     GoSub 25600
1320     GoSub 25400
1350     '
1900
1901     'GoSub 26200
1903     'NR% = 0:
         GoSub 25800
1905 '
       End If
1910 Next KU&                               ' PROXIMO REGISTRO
1990 '
1991 '                                  IMPRIMO TOTALES
1992 Printer.Print
1993 Printer.Print "+--------------------------------------------------------------------------------------------------------------------------------+"
1995 Printer.Print "| TOTALES GENERALES" + Space$(40);
'1996 Printer.Print USING; "###,###,###,###.##"; HABG#;: Printer.Print Space$(16);
1996 Printer.Print FORMATNUM$(HABG#, "F18.2") + Space$(16);
'1997 Printer.Print USING; "###,###,###,###.##"; ODEG#;: Printer.Print "                  |"
1997 Printer.Print FORMATNUM$(ODEG#, "F18.2") + "                  |"
1998 Printer.Print "|                  " + Space$(42);
1999 Printer.Print Space$(15);
'     Printer.Print USING; "###,###,###,###.##"; ASIG#;: Printer.Print Space$(16);
     Printer.Print FORMATNUM$(ASIG#, "F18.2") + Space$(16);
'2000 Printer.Print USING; "###,###,###,###.##"; NETG#;: Printer.Print " |"
2000 Printer.Print FORMATNUM$(NETG#, "F18.2") + " |"
2002 Printer.Print "+--------------------------------------------------------------------------------------------------------------------------------+"
     NR% = NR% + 5
3000 On Error GoTo 3008
3002 Printer.Print CCI$(2);: NR% = 0  ' CALL EJECT(NR%)
     Printer.Print CCI$(13);
3006 GoTo 3010
3007 '
3008 Resume 3010
3010 On Error GoTo 0
3900 LU$ = lua$: Close
3910 Call FINAL("")
3920 '
'4500 GoSub 5990
'4510 MES$(1) = "Enero": MES$(2) = "Febrero": MES$(3) = "Marzo": MES$(4) = "Abril"
'4511 MES$(5) = "Mayo": MES$(6) = "Junio": MES$(7) = "Julio": MES$(8) = "Agosto"
'4512 MES$(9) = "Setiembre": MES$(10) = "Octubre": MES$(11) = "Noviembre": MES$(12) = "Diciembre"
'4513 MES$(13) = "Aguinaldo": MES$(14) = "Vacaciones"
'4515 Return
'4520 '
'5620 If N6% > 0 Then Close #N6% Else N6% = FreeFile
'        OPEN "R",#N6%,LU$ +"FAMCAR.DAT",50
'5625 FIELD #N6%,2 AS F1$,25 AS F2$,2 AS F3$,10 AS F4$,11 AS F5$
'5630 Return
'5635 '
'5650 If N6% > 0 Then Close #N6% Else N6% = FreeFile
'        OPEN "R",#N6%,LU$+"FORMREC.DAT",96
'5652 FIELD #N6%,1 AS F1$,4 AS F2$,1 AS F3$,4 AS F4$,4 AS F5$,4 AS F6$,4 AS F7$,1 AS F8$,30 AS F9$,30 AS F10$,1 AS F11$,12 AS LIBRE$
'5658 GET #N6%,1
'5660 REDON% = Asc(F8$)
'5662 If N6% > 0 Then Close #N6%: N6% = 0
'5665 Return
'5668 '
'5700 If N8% > 0 Then Close #N8% Else N8% = FreeFile
'        OPEN "R",#N8%,LU$+"DETLIQUI.DAT",16
'5705 FIELD #N8%,2 AS J1$,2 AS J2$,4 AS J3$,8 AS J4$
'5710 Return
'5715 '
'5750 If N7% > 0 Then Close #N7% Else N7% = FreeFile
'        OPEN "R",#N7%,LU$+"BILLETES.DAT",40
'5754 FIELD #N7%,40 AS C1$
'5758 GET #N7%,1
'5760    For I% = 1 To 10
'5765    VALBI(I%) = CVSMBF(Mid$(C1$, 4 * I% - 3, 4))
'5770    Next I%
'5775 If N7% > 0 Then Close #N7%: N7% = 0
'5777 VALMIN# = CDbl(VALBI(1))
'5778 For I% = 1 To 10
'5780 If VALBI(I%) <= 0 Then
'            GoTo 5785
'     End If
'5782 If VALBI(I%) < VALMIN# Then
'            VALMIN# = CDbl(VALBI(I%))
'     End If
'5785 Next I%
'5786 If VALMIN# < 0.01 Then
'            VALMIN# = 1
'     End If
'5788 Return
'5790 '
'5850 If N1% > 0 Then Close #N1% Else N1% = FreeFile
'        OPEN "R",#N1%,LU$+"DATEMPRE.DAT",128
'5851 FIELD #N1%,30 AS D1$,24 AS D2$,16 AS D3$,58 AS LIBRE$
'5858 '
'5859 GET #N1%,1:RAZON$=D1$:LOCA$=D2$:CAJU$=D3$
'5860 GET #N1%,2:FANT$=D1$:TELE$=D2$:NUIN$=D3$
'5861 GET #N1%,3:DOMIC$=TRIM$(D1$):DATO1$=D2$:DATO2$=D3$
'5862 If InStr(RAZON$, "LODIGIA") > 0 Then
'            FAE% = 1
'     End If
'5889 If N1% > 0 Then Close #N1%: N1% = 0
'        Return
'5899 '
'5900 If N11% > 0 Then Close #N11% Else N11% = FreeFile
'        OPEN "R",#N11%,LU$+"ULTILIQ.DAT",128
'5905 FIELD #N11%,30 AS U1$,1 AS U2$,1 AS U3$,1 AS U4$,1 AS U5$,1 AS U6$,2 AS U7$,20 AS U8$,16 AS U9$,2 AS U10$,16 AS U11$,12 AS U12$,2 AS U13$,1 AS U14$,1 AS U15$,2 AS U16$,19 AS U17$
'5906 Return
'5909 '
'5990 If N10% > 0 Then Close #N10% Else N10% = FreeFile
'        OPEN "R",#N10%,LU$+"DATGREMP.DAT",64
'5991 FIELD #N10%,30 AS K1$,30 AS K2$,1 AS K3$,3 AS K4$
'5992 I% = 0
'5993 I% = I% + 1
'5994 GET #N10%,I%
'        If Asc(K1$) > 0 Then
'            GRU$(I%) = K1$
'            AGR%(I%) = Asc(K3$)
'            GoTo 5993
'        End If
'5995 NUGRU% = I% - 1
'5996 If N10% > 0 Then Close #N10%: N10% = 0
'        Return
'5997 '
'6100 GoSub 5900
'6105 GET #N11%,1
'6110 MES% = Asc(U3$)
'        If MES% > 0 Then
'            If MES% < 15 Then MES$ = MES$(MES%)
'        Else
'            MES% = 0
'        End If
'6113 MES$ = Mid$(Str$(MES%), 2)
'6115 AN% = Asc(U4$): AN$ = TRIM$(Str$(AN%)): While Len(AN$) < 2: AN$ = "0" + AN$: Wend
'6117 QUIN% = Asc(U4$)
'6119 GRU% = Asc(U2$): GRU$ = GRU$(GRU%)
'6120 CORRESP$ = U1$
'6121 DES$ = CORRESP$
'6122 For GI% = 30 To 1 Step -1
'6123 If Mid$(DES$, GI%, 1) > " " Then
'            GoTo 6126
'     End If
'6124 Next GI%
'6125 GI% = 0
'6126 CORRESP$ = Left$(DES$, GI%)
'6134 LPAG$ = U8$
'6135 FEPA% = CVI(U7$): FEX = FEPA%: GoSub 30300: FEPA$ = FE$
'6136 FEDEP% = CVI(U10$): FEX = FEDEP%: GoSub 30300: FEDEP$ = FE$
'6142 UDEP$ = U9$
'6148 BAN$ = U11$: PERLI$ = U12$: LSUE% = CVI(U16$)
'6149 '
'6150 If N11% > 0 Then Close #N11%: N11% = 0
'        Return
'6199 '
'6300 QUASI% = 0
'6305 If MES% > 12 Then
'            GoTo 6315
'     End If
'6310 If QUIN% = 2 Or GRU% <> 2 Then
'            QUASI% = 1
'     End If
'6315 Return
'6320 '
8000  JV% = 0
      RFG$ = RECFILT$("FAMCAR", 1, MKI$(LEGAX%))
      For KKUL& = 1 To Len(RFG$) Step 4
        IVL& = CVS(Mid$(RFG$, KKUL&, 4))
        F00$ = REGLEIDO$("FAMCAR", IVL&)
8050    JV% = JV% + 1
8060    NOM1$(JV%) = Mid$(F00$, 3, 25)
8070    FEX = CVI(Mid$(F00$, 28, 2)): NACI1$(JV%) = FECHATEX$(FEX)
8080    PARE1$(JV%) = Mid$(F00$, 30, 10)
8090    ESTU1$(JV%) = Mid$(F00$, 40, 11)
      Next KKUL&
      '
8130 NUPAR% = JV%
8140 Return
8150 '
'9100 LOCATE RG%, CL%
'     Return
'
'9200 Call PRESENTA(RG%, VI%, CL%, LC%, LK%, TV%, LD%, JT%, VD%, VD#, VD$, CT%)
'     Return
'
'9900 Print TRXPRIN$(Chr$(27) + "[" + Mid$(Str$(20), 2) + ";" + Mid$(Str$(Int((69 - Len(ERRO$)) / 2)), 2) + "H");: Print TRXPRIN$("ERROR: " + ERRO$ + " !!!")
'9910 For I1% = 1 To 4: For KX% = 1 To 25000: Next KX%: Next I1%
'9915 Print TRXPRIN$(Chr$(27) + "[20;1H" + Chr$(186) + Space$(78) + Chr$(186))
'9925 Return
'9930 '
'11850 I% = 0: LEGANT% = 0
'11855 I% = I% + 1
'11860 GET #N8%,I%
'11865 LEGA% = CVI(J1$)
'        If LEGA% = 0 Then
'            JX% = I% - 1
'            GoTo 11899
'        End If
'11870 If LEGA% = LEGANT% Then
'            GoTo 11855
'      End If
'11875 LEGANT% = LEGA%
'11880 For J% = J1% To NUEMPL%
'11885 If LEGA% <> EMPL%(J%) Then
'            GoTo 11895
'      End If
'11890 If PUNTAJU%(J%) <= 0 Then
'            PUNTAJU%(J%) = I%
'            J1% = J%
'            GoTo 11898
'      End If
'11895 Next J%
'11897 If J1% > 1 Then
'            J1% = 1
'            GoTo 11880
'      End If
'11898 GoTo 11855
'11899 Return
'11900 '
'24000 Print TRXPRIN$(Chr$(27) + "[5;1H" + Chr$(201) + String$(78, 205) + Chr$(187))
'24010 Print TRXPRIN$(Chr$(27) + "[6;1H" + Chr$(186) + "           LIBRO DE SUELDOS Y JORNALES (Ley 20.744 - Art. 52 t.o.)            " + Chr$(186))
'24020 Print TRXPRIN$(Chr$(27) + "[7;1H" + Chr$(204) + String$(78, 205) + Chr$(185))
'24030 For NR% = 8 To 18
'24040 NR$ = Mid$(Str$(NR%), 2)
'24050 Print TRXPRIN$(Chr$(27) + "[" + NR$ + ";1H" + Chr$(186) + Space$(78) + Chr$(186))
'24060 Next NR%
'24070 Print TRXPRIN$(Chr$(27) + "[19;1H" + Chr$(204) + String$(78, 205) + Chr$(185))
'24080 Print TRXPRIN$(Chr$(27) + "[20;1H" + Chr$(186) + Space$(78) + Chr$(186))
'24090 Print TRXPRIN$(Chr$(27) + "[21;1H" + Chr$(200) + String$(78, 205) + Chr$(188))
'24095 Return
'24096 '
'
'
'
'
'

25000 On Error GoTo 25070
      GoSub 25800
25001 L$(1) = RAZON$
25003 NUHOJ% = NUHOJ% + 1
25004 Printer.Print CCI$(14) + L$(1) + CCI$(15)
      'IF CCI$(19)=CHR$(80) THEN
      Printer.Print CCI$(12);
      'END IF
'25008 If LSUE% = 0 Then
'     GoSub 5900
'     GET #N11%,1
'     LSet U16$ = MKI$(HO%)
'     PUT #N11%,1
'     If N11% > 0 Then Close #N11%: N11% = 0
'     LSUE% = HO%
'      End If
      '
      FEPA$ = FECHATEX$(FEPA_LIQ%) 'Fecha de Pago
      CORRESP$ = PERIO_LIQ$ 'Periodo Liquidado
      Call CARGEMPRESA
      DATO1$ = DAT1_EMP$ 'ACTIVIDAD PRINCIPAL
      DATO2$ = DAT2_EMP$ 'CUIT DE LA EMPRESA
      '
25010 Printer.Print DOMIC$;
      Printer.Print Space$(114 - Len(DOMIC$)) + "Hoja Nro.: ";
      Printer.Print FORMATNUM$(NUHOJ%, "F5.0")
25015 Printer.Print LOCA$
25030 Printer.Print: Printer.Print Space$(23) + "L I B R O   D E   S U E L D O S   Y   J O R N A L E S   (Ley 20.744 - Art. 52 t.o.)"
25031 Printer.Print Space$(23) + "***********************************************************************************"
25040 Printer.Print
      Printer.Print "Actividad principal: " + DATO1$ + Space$(62) + "Fecha de pago: " + FEPA$
      TXTX$ = Left$("           C.U.I.T.: " + TRIM$(DATO2$) + Space$(50), 50)
25042 Printer.Print TXTX$ + Space$(49 - Len(CORRESP$)) + "              Liq. corresp. a: " + CORRESP$
25050 Printer.Print
25060 Printer.Print: GoTo 25075
25070 Resume 25075
25075 On Error GoTo 0: Return
25080 '
25100 On Error GoTo 25155
25105 For TY% = 1 To NURENGLO%
25106 If Abs(IMCO#(TY%)) < 0.005 Then
            GoTo 25150
      End If
25108 COI% = (CO%(TY%) Mod 100)
25109 If CO%(TY%) >= 100 Then
            'Printer.Print USING; "|######  | "; CO%(TY%);
            Printer.Print "|";
            Printer.Print FORMATNUM$(CO%(TY%), "F6.0");
            Printer.Print "  | ";
            GoTo 25112
      End If
25110 COAUX% = CO%(TY%) + 1000
      COI1$ = Mid$(Str$(COAUX%), 3, 3)
25111 Printer.Print "|   " + COI1$ + "  | ";
25112 RR% = Int(CO%(TY%) / 200)
        AJU$ = ""
        If (CO%(TY%) Mod 200) > 100 Then
            AJU$ = "AJUSTE "
        End If
25113 On RR% + 1 GoTo 25115, 25117, 25119
25115 Printer.Print AJUSTI$(AJU$ + CON$(COI%), 30);: GoTo 25121
25117 Printer.Print AJUSTI$(AJU$ + RET$(COI%), 30);: GoTo 25121
25119 Printer.Print AJUSTI$(AJU$ + ASI$(COI%), 30);
25121 If CACO(TY%) < 0.1 Then
            Printer.Print Space$(8) + "|" + Space$(11);
            GoTo 25135
      End If
25124 If RR% <> 2 Then
            'MA4$ = "        |######.##  "
            MA4$ = "        |" + FORMATNUM$(CACO(TY%) * Sgn(IMCO#(TY%)), "F9.2") + "  "
      End If
25125 If RR% = 2 Then
            'MA4$ = "        |######     "
            MA4$ = "        |" + FORMATNUM$(CACO(TY%) * Sgn(IMCO#(TY%)), "F6.0") + "     "
      End If
'25126 Printer.Print USING; MA4$; CACO(TY%) * Sgn(IMCO#(TY%));
25126 Printer.Print MA4$;
        If ACTHOR%(TY%) > 0 Then
            SUMHOR = SUMHOR + CACO(TY%) * Sgn(IMCO#(TY%))
        End If
25128 CACO(TY%) = 0
25135 On RR% + 1 GoTo 25140, 25142, 25144
'25140 Printer.Print USING; MA1$; IMCO#(TY%): NR% = NR% + 1: GoTo 25150
25140 Printer.Print "| " + FORMATNUM$(IMCO#(TY%), "F14.2") + " |                |                |                |"
        NR% = NR% + 1: GoTo 25150
'25142 Printer.Print USING; MA2$; IMCO#(TY%): NR% = NR% + 1: GoTo 25150
25142 Printer.Print "|                |                | " + FORMATNUM$(IMCO#(TY%), "F14.2") + " |                |"
        NR% = NR% + 1: GoTo 25150
'25144 Printer.Print USING; MA3$; IMCO#(TY%): NR% = NR% + 1
25144 Printer.Print "|                | " + FORMATNUM$(IMCO#(TY%), "F14.2") + " |                |                |"
        NR% = NR% + 1
25150 Next TY%
25151 GoTo 25160
25152 '
25155 Resume 25160
25160 On Error GoTo 0: Return
25165 '
25400 On Error GoTo 25425
25402 Printer.Print "+--------+---------------------------------------+-----------+----------------+----------------+----------------+----------------+": NR% = NR% + 1
25405 Printer.Print "|        | TOTALES                               |";
'25407 Printer.Print USING; "##,###.##  |"; SUMHOR;
25407 Printer.Print FORMATNUM(SUMHOR, "F9.2") + "  |";
'25410 Printer.Print USING; " ###,###,###.## |"; TOHA#; TOASI#; TODE#; TOHA# + TOASI# - TODE#: NR% = NR% + 1
25410 Printer.Print FORMATNUM$(TOHA#, "F15.2") + " |";
      Printer.Print FORMATNUM$(TOASI#, "F15.2") + " |";
      Printer.Print FORMATNUM$(TODE#, "F15.2") + " |";
      Printer.Print FORMATNUM$(TOHA# + TOASI# - TODE#, "F15.2") + " |"
25416 HABG# = HABG# + TOHA#: ASIG# = ASIG# + TOASI#: ODEG# = ODEG# + TODE#: NETG# = NETG# + (TOHA# + TOASI# - TODE#)
25417 If NUPAR% > 0 Then
     Printer.Print "+--------+---------------------------------+-----+-----------+---------+------+----------------+------+---------+----------------+": NR% = NR% + 1
     GoSub 25450
  Else
     Printer.Print "+--------+---------------------------------------+-----------+----------------+----------------+----------------+----------------+": NR% = NR% + 1
      End If
25420 Printer.Print: NR% = NR% + 1
25421 GoTo 25430
25422 '
25425 Resume 25430
25430 On Error GoTo 0: Return
25435 '
25450 Printer.Print "| " + "Familiares : " + NOM1$(1) + "   | " + "Parentesco : " + PARE1$(1) + "   | ";
25454 Printer.Print "Fecha Nac./Cas. : " + NACI1$(1) + "   | " + "   Docum.: " + ESTU1$(1) + "   |"
25458 NR% = NR% + 1
        If NUPAR% = 1 Then
            GoTo 25476
        End If
25460 For II% = 2 To NUPAR%
25464 Printer.Print "|" + Space$(14) + NOM1$(II%);
25466 Printer.Print "   | " + Space$(13) + PARE1$(II%);
25468 Printer.Print "   | " + Space$(18) + NACI1$(II%);
25470 Printer.Print "   | " + Space$(11) + ESTU1$(II%) + "   | "
25472 NR% = NR% + 1
25474 Next II%
25476 Printer.Print "+------------------------------------------+---------------------------+------------------------------+--------------------------+": NR% = NR% + 1
25478 Return
25480 '
25500 On Error GoTo 25510
25502 If NR% < 6 Then
            GoTo 25515
      End If
25505 Printer.Print CCI$(2);     ' CALL EJECT(NR%)
      NR% = 0
25506 GoTo 25515
25507 '
25510 Resume 25515
25515 On Error GoTo 0: Return
25516 '
25600 On Error GoTo 25635
25605 VALNET# = TOHA# + TOASI# - TODE#
25610 If REDON% = 0 Then
            GoTo 25640
      End If
'25612 MARCA% = 0: GoSub 27400
25613 If AJRED# = 0 Then
            GoTo 25640
      End If
25614 Printer.Print "|" + Space$(8) + "| " + "REDONDEO" + Space$(30) + "|" + Space$(11) + "|";
25615 If MARCA% = 2 Then
            Printer.Print Space$(16) + "|" + Space$(16) + "|";
            Printer.Print USING; " ###,###,###.## "; AJRED#;
            Printer.Print "|" + Space$(16) + "|"
      End If
25620 If MARCA% = 1 Then
            Printer.Print USING; " ###,###,###.## "; AJRED#;
            Printer.Print "|" + Space$(16) + "|" + Space$(16) + "|"
      End If
25625 VALNET# = TOHA# + TOASI# - TODE#
25626 GoTo 25640
25630 '
25635 Resume 25640
25640 On Error GoTo 0
      Return
25641 '
25700 On Error GoTo 25745
25702 Printer.Print "+----------------+-----------------------------------------+--------------------------+------------------------------+-----------+"
25705 Printer.Print "| Legajo No.     | " + NOM$ + "          | Fecha Ingreso: " + INGRE$ + "  | Seccion: " + SECCI$ + Space$(13) + "| Est.Civil |"
25707 Printer.Print "|         ";
      Printer.Print FORMATNUM$(LEGA%, "F6.0");
      Printer.Print " ";
      CUIL$ = Left$("C.U.I.L.: " + TRIM$(AFIL$) + Space$(40), 40)
25710 Printer.Print "| " + CUIL$ + "| Fecha Egreso: " + FEGRE$ + "   | Tarea / Categoria: " + CATE$ + "| " + ECIV$ + "|"
25725 Printer.Print "+--------+-------+-------------------------------+---------+-+----------------+-------+--------+----------------+----+-----------+"
25730 Printer.Print "| Codigo | Detalle                               |   Cant.   | Remuneraciones |  Asig. Famil.  |   Retenciones  |   Total Neto   |"
25735 Printer.Print "+--------+---------------------------------------+-----------+----------------+----------------+----------------+----------------+"
25740 NR% = NR% + 6
25741 GoTo 25750
25744 '
25745 Resume 25750
25750 On Error GoTo 0: Return
25752 '
25800 On Error GoTo 25825
25802 Printer.Print CCI$(1) + CCI$(4);
'25805 WIDTH "LPT1:",255
'25810 SP$ = SETPRIN$(796): Printer.Print SP$;
25812 MA1$ = "| ###,###,###.## |                |                |                |"
25813 MA2$ = "|                |                | ###,###,###.## |                |"
25814 MA3$ = "|                | ###,###,###.## |                |                |"
25820 'IF CCI$(19)=CHR$(80) THEN
     Printer.Print CCI$(12);
      'END IF
25821 '
'25822 GoTo 25830
25823 '
25825 'Resume 25830
'25830 On Error GoTo 65000: Return
25840 Return
'

End Sub

