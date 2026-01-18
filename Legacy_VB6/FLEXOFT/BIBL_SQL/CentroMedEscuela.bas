Attribute VB_Name = "CentroMedEscuela"

Sub GRABACUOTA(NC, DESCRICURSO$, Tipo%, LETRA$, NUMPVTA%, DOCUMENTO$, IMPORTE#, MODO$, RET_NUMFACTURALAR$, RET_NUMFACTUCOR$, FEINICIO%)
    'GRABA REGISTRO POR EL TOTAL DE LA FACTURA 1
    '
    TICOMX$ = "FC"
    If Tipo% = 3 Then TICOMX$ = "ND"
    If Tipo% = 6 Then TICOMX$ = "NC"
    FFI% = HOY(HO$)
    '
    NPVTA& = NUMPVTA%
    PVTA$ = FORMATOCON0$(NPVTA&, 4)
    NFCZ$ = FORMATOCON0$(TRIM$(Str(ASIGCURSO.SIG_NUMERO_FACTURA&(2, "B"))), 8)
    NFCX$ = TICOMX$ + "-" + Mid$(NFCZ$, 3) + "-" + LETRA$
    NFCY$ = TICOMX$ + "-" + PVTA$ + "-" + NFCZ$ + "-" + LETRA$
    RET_NUMFACTUCOR$ = NFCX$
    RET_NUMFACTURALAR$ = NFCY$
    OPMOVIX$ = UCase(Left$(NFCY$, 2))
    '
    SIGX% = 1
    If OPMOVIX$ = "NC" Then SIGX% = (-1)  ' SI ES NOTA DE CREDITO
    '
    CoComCor$ = NFCX$
    CoComLar$ = NFCY$
    NumeClie = Int(GRUDIVCLI(NC))
    NumeProv% = 0
    FechEmi = CVDAT(FEINICIO%)
    Referen$ = rasocli$(NC)
    FeContab = CVDAT(FEINICIO%)
    FeAcredi = CVDAT(FFI%)
    CuentaCo% = CUECOCLI%(NC)
    TiMovi$ = "FVEN"
    OpMovi$ = OPMOVIX$
    VaMovi$ = LETRA$
    DOCUMEN$ = "Factura:": If OPMOVIX$ = "NC" Then DOCUMEN$ = "N.Crédito:": If OPMOVIX$ = "ND" Then DOCUMEN$ = "N.Débito:"
    Descrip$ = DOCUMEN$ + " " + LETRA$ + " (Mat.Codif.)" + " " + PVTA$ + "-" + TRIM$(Str(XVALO(Mid$(NFCY$, 9, 8))))
    NumeTal = NUMPVTA%
    NumeCom = XVALO(Mid$(NFCY$, 9, 8))
    StatInit = 0
    ImpoDeb = IMPORTE# * SIGX%
    ValoDola = IMPORTE# 'TOTADOLAR# * SIGX% ' VER COMO SE VA A GRABAR SI ES A VALOR DOLAR ACTUALIZADO
    Observa$ = DESCRICURSO$
    NumeroCae$ = ""
    IdentiCae$ = ""
    ObservaFactu$ = DESCRICURSO$ & " Correspondiente a " & DOCUMENTO$ & " " & METEXTO$(FEINICIO%)  ' MATRICULA ,CUOTA O DERECHO DE EXAMEN"
    Call GRACABAN
    '
    'GRABO EL MOVIMIENTO TIPO VENTA POR EL IMPORTE TOTAL POR QUE AQUI NO HAY IVA
    TOTANETX# = IMPORTE#
    If TOTANETX# >= 0.005 Then
      PPVI% = 0
      If PPVI% < 0 Or PPVI% > 7 Then PPVI% = 1
      CUENTAX$ = CONTROL$("SVD451S", "CUEVE-" + TRIM$(Str$(PPVI%)))
      QQ1% = CUECOINT%(CUENTAX$)
      CuentaCo% = QQ1%
      TiMovi$ = "VENTA"
      OpMovi$ = OPMOVIX$
      VaMovi$ = LETRA$
      Descrip$ = "Mercaderia Vendida " + NFCX$
      StatInit = 0
      ImpoDeb = TOTANETX# * SIGX% * (-1)
        NENETODOL# = TOTANETX# '(NENETO# / TICAMONE(2))
        ValoDola = NENETODOL# * SIGX% * (-1)
      Call GRACABAN
    End If

    VARIDOC% = 1
    '
    If LETRA$ = "B" Then VARIDOC% = 2
    If LETRA$ = "E" Then VARIDOC% = 3
    Tipo% = 2
    If OPMOVIX$ = "ND" Then Tipo% = 3
    If OPMOVIX$ = "NC" Then Tipo% = 6
    COPAGI% = 1
    
    If TRIM$(DOCUMENTO$) = "MATRICULA" Then COPAGI% = 0
    
   '
    Call ASIGCURSO.GRADEUFA(Tipo%, VARIDOC%, NumeCom, NC, FEINICIO%, COPAGI%, IMPORTE#)

'SELECT * FROM CURSOS  WHERE CURSOS_ID  = 5
'SELECT * FROM MASTER WHERE Codint = 38
End Sub

Function PRESENTEALDIA%(NC, CI1, FECHA%)
     CONDI$ = "NUME_CLI = " & NC & " AND CODINT = " & CI1 & " AND FECHLIS = '" & FETEXCOR1$(FECHA%) & "'"
     'PRESENTE1 = XVALO(VALOBADA("PRESENTES", "PRESENTE", CONDI$, "NOMESS"))
     PRESENTEALDIA% = XVALO(VALOBADA("PRESENTES", "PRESENTE", CONDI$, "NOMESS"))
End Function
Function OBSERVACION_PRESENTES$(NC, CI1, FECHA%)
     CONDI$ = "NUME_CLI = " & NC & " AND CODINT = " & CI1 & " AND FECHLIS = '" & FETEXCOR1$(FECHA%) & "'"
     OBSERVACION_PRESENTES$ = VALOBADA("PRESENTES", "OBSERVACIONES", CONDI$, "NOMESS")
End Function
Function AULA$(IDAULA&)
       CONDI$ = "LISTA = 'AULAS' And CAMPO1 = '" & FORMATOCON0$(IDAULA&, 4) & "'"
       AULA$ = VALOBADA("DATASQL", "CAMPO2", CONDI$, "MOMES")
End Function
Function NameProfesor$(NC)
     CONDI$ = "NUME_CLI = " & NC & " "
     NameProfesor$ = VALOBADA("PROFESORES", "Raso_Cli", CONDI$, "NOMESS")
End Function

