Attribute VB_Name = "FORM_MAG_VIE"

Sub ALTANUEFORM()
        '
1000    CCC$ = OBJPLA$("CODEFORMU", CCC$)
        CODFOR$ = AJUSTI$(Left$(CCC$, 8), 8)
        If TRIM$(CODFOR$) = "" Then Exit Sub
        For UU% = 1 To Len(TRIM$(CODFOR$))
           VV% = Asc(Mid$(CODFOR$, UU%))
           If VV% < 48 Or VV% > 90 Or (VV% > 57 And VV% < 65) Then
              If VV% <> 45 Then
                 Call COMUNI("Código no Válido")
                 GoTo 1000
              End If
           End If
        Next UU%
        '
        DESFOR$ = Mid$(CCC$, 9, 48)
        If TRIM$(DESFOR$) = "" Then
            Beep
            GoTo 1000
        End If
        '
        If EXISTE%(LU$ + TRIM$(CODFOR$) + ".FRM") > 0 Then
              X% = COMALTER%("Ya Existe Formulario\'" + TRIM$(CODFOR$) + ".FRM'\\Otro Código\Sobre-Escribir")
              If X% <> 2 Then GoTo 1000
              Kill LU$ + TRIM$(CODFOR$) + ".FRM"
            Else
              Call REGAPP("INDIFRM", CCC$)
              Call CIERRARCH("INDIFRM")
        End If
        '
        If ULTREG&("INDIFRM") > 1 Then
           If ALTERNA%("Copia Formato Existente") = 1 Then
1060          Call SELECHO("INDIFRM")
              CCC$ = VALACT1$("INDIFRM")
              If Len(CCC$) <= 4 Then Exit Sub
              '
              CODFOR1$ = AJUSTI$(Left$(CCC$, 8), 8)
              If CODFOR1$ = CODFOR$ Then
                  Beep
                  GoTo 1060
              End If
              '
              FileCopy LU$ + TRIM$(CODFOR1$) + ".FRM ", LU$ + TRIM$(CODFOR$) + ".FRM"
           End If
        End If
        '
End Sub
'
'Sub DEFLAYOUTW()

    '
'    If TRIM$(CODFOR$) = "" Then
'        Call MENSERR(24, "Falta Seleccionar Formulario")
'        PRINFORM.Text1.SetFocus
'        Exit Sub
'    End If
    '
'    NX% = FILEOPEN%(LU$ + CODFOR$ + ".FRM", 0, 128)
'    Dim RF As String * 128
'    Call CIERRARCH("!CARFORMU")
'    Call BLANARCH("!CARFORMU")
    '
    'Call REGAPP("!CARFORMU", "TI-LETRATipo de letra (PICA - ELITE - COMPRIM)")
    'Call REGAPP("!CARFORMU", "LIN-PULGLineas por Pulgada (6 - 8)")
    'Call REGAPP("!CARFORMU", "CAN-COL Cantidad de Columnas sin Margen (h.256)")
    'Call REGAPP("!CARFORMU", "CAN-REN Cantidad de Renglones (inc.Margen Sup)")
'    Call REGAPP("!CARFORMU", "MARSUPMMMargen Superior (milimetros)")
'    Call REGAPP("!CARFORMU", "MARIZQMMMargen Izquierdo (milimetros)")
'    Call REGAPP("!CARFORMU", "MARGENINMargen Intermedio (espacios)")
'    Call REGAPP("!CARFORMU", "LARGOJMMLargo de Hoja (milimetros)")
'    Call REGAPP("!CARFORMU", "CAN-COP Cantidad de Copias (hasta 6)")
'    Call REGAPP("!CARFORMU", "IMP-DES Primera Copia con Importes (Factura)")
'    Call REGAPP("!CARFORMU", "IMP-HAS Ultima Copia con Importes (Factura)")
'    Call REGAPP("!CARFORMU", "DETDESMMBorde Superior de Detalle (milimetros)")
'    Call REGAPP("!CARFORMU", "DETHASMMBorde Inferior de Detalle (milimetros)")
'    Call REGAPP("!CARFORMU", "INTERLINInterlinea de Detalle (milimetros)")
'    Call REGAPP("!CARFORMU", "ORIENTA Orientacion (Retrato / Paisaje)")
'    Call REGAPP("!CARFORMU", ".")
'    Call REGAPP("!CARFORMU", "MON-PRI Moneda Principal (PESOS - DOLARES)")
'    Call REGAPP("!CARFORMU", "MON-CEN Moneda Secundaria (CENTAVOS)")
'    Call REGAPP("!CARFORMU", "LOIMPLETLongitud Importe en Letras")
'    Call REGAPP("!CARFORMU", "PREFIFACPrefijo de Facturación")
'    Call REGAPP("!CARFORMU", "MAX-ITEMMaximo de Items por Documento")
'    '
'    For U& = 1 To 128
'        CODVAR$(U&) = ""
'        REGFOR%(U&) = 0
'    Next U&
'    '
'    CAREL& = ULTREG&("!CARFORMU")
'    For U& = 1 To CAREL&
'        CODVAR$(U&) = Left$(REGLEIDO$("!CARFORMU", U&), 8)
'    Next U&
'    '
'    For I% = 1 To LOF(NX%) / 128
'        Get #NX%, I%, RF$
'        CODVARI$ = Left$(RF$, 8)
'        For J& = 1 To CAREL&
'            If CODVAR$(J&) = CODVARI$ Then
'                REGFOR%(J&) = I%
'                NYY$ = RF$
'                Call GRAREG("!CARFORMU", NYY$, J&)
'            End If
'        Next J&
'    Next I%
'    '
'    VTB% = VENTABU%("CONFORMU/TITUPAN=" + DESFOR$)
'    If VTB% >= 0 Then
'        For U& = 1 To ULTREG&("!CARFORMU")
'           XXX$ = REGLEIDO$("!CARFORMU", U&)
'           If REGFOR%(U&) < 1 Then
'               REGFOR%(U&) = 1 + LOF(NX%) / 128
'           End If
'           RF$ = XXX$
'           Put #NX%, REGFOR%(U&), RF$
'        Next U&
'    End If
'    '
'    Close #NX%: NX% = 0
'    '
'End Sub
'

'Sub DEFLEYENW()
'    '
'    If TRIM$(CODFOR$) = "" Then
'        Call MENSERR(24, "Falta Seleccionar Formulario")
'        PRINFORM.Text1.SetFocus
'        Exit Sub
'    End If
'    '
'
'    NX% = FILEOPEN%(LU$ + CODFOR$ + ".FRM", 0, 128)
'    Dim RF As String * 128
'    Call CIERRARCH("!CARFORMW")
'    Call BLANARCH("!CARFORMW")
'    For I% = 1 To LOF(NX%) / 128
'        Get #NX%, I%, RF$
'        If TRIM$(Left$(RF$, 8)) = "TEXTWIN" Then
'            RFF$ = RF$
'            Call REGAPP("!CARFORMW", RFF$)
'        End If
'    Next I%
'    '
'    VTB% = VENTABU%("CARLEYEW")
'    If VTB% >= 0 Then
'        J% = 0
'        For I% = 1 To LOF(NX%) / 128
'            Get #NX%, I%, RF$
'            If TRIM$(Left$(RF$, 8)) <> "TEXTWIN" Then
'                J% = J% + 1
'                Put #NX%, J%, RF$
'            End If
'        Next I%
'        '
'        For U& = 1 To ULTREG("!CARFORMW")
'            RFF$ = REGLEIDO$("!CARFORMW", U&)
'            Call REPLA(RFF$, "TEXTWIN", 1, 8)
'            RF$ = RFF$
'            J% = J% + 1
'            Put #NX%, J%, RF$
'        Next U&
'        '
'        For I% = J% + 1 To LOF(NX%) / 128
'            RF$ = String$(128, 0)
'            J% = J% + 1
'            Put #NX%, J%, RF$
'        Next I%
'        '
'    End If
'    Close #NX%: NX% = 0
'    '
'End Sub
'

'Sub DEFCAJASW()
'    '
'    If TRIM$(CODFOR$) = "" Then
'        Call MENSERR(24, "Falta Seleccionar Formulario")
'        PRINFORM.Text1.SetFocus
'        Exit Sub
'    End If
'    '
'    NX% = FILEOPEN%(LU$ + CODFOR$ + ".FRM", 0, 128)
'    Dim RF As String * 128
'    Call CIERRARCH("!CARFORMW")
'    Call BLANARCH("!CARFORMW")
'    For I% = 1 To LOF(NX%) / 128
'        Get #NX%, I%, RF$
'        If TRIM$(Left$(RF$, 8)) = "CAJAWIN" Then
'            RFF$ = RF$
'            Call REGAPP("!CARFORMW", RFF$)
'        End If
'    Next I%
'    '
'    VTB% = VENTABU%("CARCAJAW")
'    If VTB% >= 0 Then
'        J% = 0
'        For I% = 1 To LOF(NX%) / 128
'            Get #NX%, I%, RF$
'            If TRIM$(Left$(RF$, 8)) <> "CAJAWIN" Then
'                J% = J% + 1
'                Put #NX%, J%, RF$
'            End If
'        Next I%
'        '
'        For U& = 1 To ULTREG("!CARFORMW")
'            RFF$ = REGLEIDO$("!CARFORMW", U&)
'            Call REPLA(RFF$, "CAJAWIN", 1, 8)
'            RF$ = RFF$
'            J% = J% + 1
'            Put #NX%, J%, RF$
'        Next U&
'        '
'        For I% = J% + 1 To LOF(NX%) / 128
'            RF$ = String$(128, 0)
'            J% = J% + 1
'            Put #NX%, J%, RF$
'        Next I%
'        '
'    End If
'    Close #NX%: NX% = 0
'    '
'End Sub
'

'Sub DEFLINEAW()
'    '
'    If TRIM$(CODFOR$) = "" Then
'        Call MENSERR(24, "Falta Seleccionar Formulario")
'        PRINFORM.Text1.SetFocus
'        Exit Sub
'    End If
'    '
'    NX% = FILEOPEN%(LU$ + CODFOR$ + ".FRM", 0, 128)
'    Dim RF As String * 128
'    Call CIERRARCH("!CARFORMW")
'    Call BLANARCH("!CARFORMW")
'    For I% = 1 To LOF(NX%) / 128
'        Get #NX%, I%, RF$
'        If TRIM$(Left$(RF$, 8)) = "LINEWIN" Then
'            RFF$ = RF$
'            Call REGAPP("!CARFORMW", RFF$)
'        End If
'    Next I%
'    '
'    VTB% = VENTABU%("CARLINEW")
'    If VTB% >= 0 Then
'        J% = 0
'        For I% = 1 To LOF(NX%) / 128
'            Get #NX%, I%, RF$
'            If TRIM$(Left$(RF$, 8)) <> "LINEWIN" Then
'                J% = J% + 1
'                Put #NX%, J%, RF$
'            End If
'        Next I%
'        '
'        For U& = 1 To ULTREG("!CARFORMW")
'            RFF$ = REGLEIDO$("!CARFORMW", U&)
'            Call REPLA(RFF$, "LINEWIN", 1, 8)
'            RF$ = RFF$
'            J% = J% + 1
'            Put #NX%, J%, RF$
'        Next U&
'        '
'        For I% = J% + 1 To LOF(NX%) / 128
'            RF$ = String$(128, 0)
'            J% = J% + 1
'            Put #NX%, J%, RF$
'        Next I%
'        '
'    End If
'    Close #NX%: NX% = 0
'    '
'End Sub
'
'Sub DEFIMAGEW()
'    '
'    If TRIM$(CODFOR$) = "" Then
'        Call MENSERR(24, "Falta Seleccionar Formulario")
'        PRINFORM.Text1.SetFocus
'        Exit Sub
'    End If
'    '
'    NX% = FILEOPEN%(LU$ + CODFOR$ + ".FRM", 0, 128)
'    Dim RF As String * 128
'    Call CIERRARCH("!CARFORMW")
'    Call BLANARCH("!CARFORMW")
'    For I% = 1 To LOF(NX%) / 128
'        Get #NX%, I%, RF$
'        If TRIM$(Left$(RF$, 8)) = "IMAGEWIN" Then
'            RFF$ = RF$
'            Call REGAPP("!CARFORMW", RFF$)
'        End If
'    Next I%
'    '
'    VTB% = VENTABU%("CARIMAGW")
'    If VTB% >= 0 Then
'        J% = 0
'        For I% = 1 To LOF(NX%) / 128
'            Get #NX%, I%, RF$
'            If TRIM$(Left$(RF$, 8)) <> "IMAGEWIN" Then
'                J% = J% + 1
'                Put #NX%, J%, RF$
'            End If
'        Next I%
'        '
'        For U& = 1 To ULTREG("!CARFORMW")
'            RFF$ = REGLEIDO$("!CARFORMW", U&)
'            Call REPLA(RFF$, "IMAGEWIN", 1, 8)
'            RF$ = RFF$
'            J% = J% + 1
'            Put #NX%, J%, RF$
'        Next U&
'        '
'        For I% = J% + 1 To LOF(NX%) / 128
'            RF$ = String$(128, 0)
'            J% = J% + 1
'            Put #NX%, J%, RF$
'        Next I%
'        '
'    End If
'    Close #NX%: NX% = 0
'    '
'End Sub
'

Sub DEFPARAMW()
    '
    If TRIM$(CODFOR$) = "" Then
        Call MENSERR(24, "Falta Seleccionar Formulario")
        PRINFORM.Text1.SetFocus
        Exit Sub
    End If
    '
    Call BLANARCH("ECOPARAM")
    Call REGAPP("ECOPARAM", "RASO-EMPRazon Social de la Empresa")
    Call REGAPP("ECOPARAM", "NOM-FANTNombre de Fantasia")
    Call REGAPP("ECOPARAM", "DOMI-EMPDomicilio")
    Call REGAPP("ECOPARAM", "CPOS-EMPCodigo Postal")
    Call REGAPP("ECOPARAM", "LOCA-EMPLocalidad")
    Call REGAPP("ECOPARAM", "TELE-EMPTelefonos")
    Call REGAPP("ECOPARAM", "CUIT-EMPNumero de C.U.I.T.")
    Call REGAPP("ECOPARAM", "IBRU-EMPNumero de Ing.Brutos")
    Call REGAPP("ECOPARAM", "CAJU-EMPNumero de Caja Jubilacion")
    Call REGAPP("ECOPARAM", ".")
    '
    Call REGAPP("ECOPARAM", "TIPO-COMTipo de Documento")
    Call REGAPP("ECOPARAM", "NUME-COMNumero de Documento")
    Call REGAPP("ECOPARAM", "DOCUORIGDocumento de origen")
    Call REGAPP("ECOPARAM", "FECH-EMIFecha de Emision")
    Call REGAPP("ECOPARAM", "HORA-EMIHora de Emision")
    Call REGAPP("ECOPARAM", "USUARIO Nombre Usuario FLEXOFT")
    Call REGAPP("ECOPARAM", "FECH-LANFecha de Lanzamiento (Orden de Fabricacion)")
    Call REGAPP("ECOPARAM", "FECH-VENFecha de Vencimiento / Entrega")
    Call REGAPP("ECOPARAM", "HOJA-NUMNumero de Orden de Hoja Actual")
    Call REGAPP("ECOPARAM", "HOJA-TOTTotal de Hojas del Documento")
    Call REGAPP("ECOPARAM", "ORD-COPIOrden de Copia (Original / Duplicado / etc.)")
    Call REGAPP("ECOPARAM", "DESTINO Destino (Orden de Fabricacion)")
    Call REGAPP("ECOPARAM", "LIST-PRETexto Lista de Precios")
    Call REGAPP("ECOPARAM", "COND-PAGTexto Condicion de Pago")
    Call REGAPP("ECOPARAM", "VENDEDORTexto Vendedor")
    Call REGAPP("ECOPARAM", "ZONA-VENTexto Zona de Ventas")
    Call REGAPP("ECOPARAM", "DEPOSITODeposito de Entrada / Salida")
    Call REGAPP("ECOPARAM", "OBSERVA Observaciones - Comentarios")
    Call REGAPP("ECOPARAM", ".")
    '
    Call REGAPP("ECOPARAM", "COD-CONJCodigo de Conjunto (Orden de Fabricacion)")
    Call REGAPP("ECOPARAM", "DES-CONJDescripcion de Conjunto (Orden de Fabricacion)")
    Call REGAPP("ECOPARAM", "REF-CONJReferencia de Complementos (Orden de Fabricacion)")
    Call REGAPP("ECOPARAM", "UNI-CONJUnidad del Conjunto (Orden de Fabricacion)")
    Call REGAPP("ECOPARAM", "COL-CONJColor de Conjunto (Orden de Fabricacion)")
    Call REGAPP("ECOPARAM", "TAL-CONJTalle de Conjunto (Orden de Fabricacion)")
    Call REGAPP("ECOPARAM", "CAN-CONJCantidad de Conjunto (Orden de Fabricacion)")
    Call REGAPP("ECOPARAM", "PES-CONJPeso en Kilos de Conjunto (Orden de Fabricacion)")
    Call REGAPP("ECOPARAM", ".")
    '
    Call REGAPP("ECOPARAM", "COD-MPRICodigo de Materia Prima (Orden de Fabricacion)")
    Call REGAPP("ECOPARAM", "DES-MPRIDescripcion M.Prima (Orden de Fabricacion)")
    Call REGAPP("ECOPARAM", "UNI-MPRIUnidad de Mat. Prima (Orden de Fabricacion)")
    Call REGAPP("ECOPARAM", "USO-MPRIUso Unitario de M.Prima (Orden de Fabricacion)")
    Call REGAPP("ECOPARAM", "CAN-MPRICantidad de M.Prima (Orden de Fabricacion)")
    Call REGAPP("ECOPARAM", "TIP-MPRITipo de Material (Orden de Fabricacion)")
    Call REGAPP("ECOPARAM", "TIT-MPRITitulo del Hilado (Orden de Fabricacion)")
    Call REGAPP("ECOPARAM", "COL-MPRIColor de Material (Orden de Fabricacion)")
    Call REGAPP("ECOPARAM", "PRO-MPRIProveedor del Material (Orden de Fabricacion)")
    Call REGAPP("ECOPARAM", "PAR-MPRILote-Partida del Material (Orden de Fabricacion)")
    Call REGAPP("ECOPARAM", "UBI-MPRIUbicación del Material (Orden de Fabricacion)")
    Call REGAPP("ECOPARAM", ".")
    '
    Call REGAPP("ECOPARAM", "OPE-PUEPOperacion (Puesta a Punto)")
    Call REGAPP("ECOPARAM", "MAQ-PUEPMaquina (Puesta a Punto)")
    Call REGAPP("ECOPARAM", "PAR-PUEPParámetro (Puesta a Punto)")
    Call REGAPP("ECOPARAM", "HER-PUEPHerramental (Puesta a Punto)")
    Call REGAPP("ECOPARAM", ".")
    '
    Call REGAPP("ECOPARAM", "COT-PLCOCota (Plan de Control)")
    Call REGAPP("ECOPARAM", "REF-PLCOReferencia (Plan de Control)")
    Call REGAPP("ECOPARAM", "MED-PLCOMedio de Control (Plan de Control)")
    Call REGAPP("ECOPARAM", "CAN-PLCOCantidad (Plan de Control)")
    Call REGAPP("ECOPARAM", "FRE-PLCOFrecuencia (Plan de Control)")
    Call REGAPP("ECOPARAM", "REA-PLCOPlan de Reacción (Plan de Control)")
    Call REGAPP("ECOPARAM", ".")
    '
    Call REGAPP("ECOPARAM", "LAR-CUTELargo Cuerpo Tejeduria (Orden de Fabricacion)")
    Call REGAPP("ECOPARAM", "ACH-CUTEAncho Cuerpo Tejeduria (Orden de Fabricacion)")
    Call REGAPP("ECOPARAM", "LAR-MATELargo Manga Tejeduria (Orden de Fabricacion)")
    Call REGAPP("ECOPARAM", "ACH-MATEAncho Manga Tejeduria (Orden de Fabricacion)")
    Call REGAPP("ECOPARAM", "LAR-CUCOLargo Cuerpo Corte (Orden de Fabricacion)")
    Call REGAPP("ECOPARAM", "ACH-CUCOAncho Cuerpo Corte (Orden de Fabricacion)")
    Call REGAPP("ECOPARAM", "LAR-MACOLargo Manga Corte (Orden de Fabricacion)")
    Call REGAPP("ECOPARAM", "ACH-MACOAncho Manga Corte (Orden de Fabricacion)")
    Call REGAPP("ECOPARAM", ".")
    '
    Call REGAPP("ECOPARAM", "PIE-COMPPieza Complemento (Orden de Fabricacion)")
    Call REGAPP("ECOPARAM", "DIB-COMPDibujo Complemento (Orden de Fabricacion)")
    Call REGAPP("ECOPARAM", "MAQ-COMPMáquina Complemento (Orden de Fabricacion)")
    Call REGAPP("ECOPARAM", "LAR-COMPLargo(cm) Complemento (Orden de Fabricacion)")
    Call REGAPP("ECOPARAM", "ACH-COMPAncho(cm) Complemento (Orden de Fabricacion)")
    Call REGAPP("ECOPARAM", "UNI-COMPUnidad Complemento (Orden de Fabricacion)")
    Call REGAPP("ECOPARAM", "CAN-COMPCantidad Complemento (Orden de Fabricacion)")
    Call REGAPP("ECOPARAM", "PES-COMPPeso(kg) Complemento (Orden de Fabricacion)")
    Call REGAPP("ECOPARAM", ".")
    '
    Call REGAPP("ECOPARAM", "NOP-OPERNumero de Operacion (Orden de Fabricacion)")
    Call REGAPP("ECOPARAM", "DES-OPERDescripcion Corta de Operacion (O.Fabricacion)")
    Call REGAPP("ECOPARAM", "OBS-OPERDescripcion Larga de Operacion (O.Fabricacion)")
    Call REGAPP("ECOPARAM", "EQUIPO1 Maquina / Equipo Nro.1 (Orden de Fabricacion)")
    Call REGAPP("ECOPARAM", "EQUIPO2 Maquina / Equipo Nro.2 (Orden de Fabricacion)")
    Call REGAPP("ECOPARAM", "EQUIPO3 Maquina / Equipo Nro.3 (Orden de Fabricacion)")
    Call REGAPP("ECOPARAM", "TIE-PREPHoras de Preparacion (Orden de Fabricacion)")
    Call REGAPP("ECOPARAM", "TIE-PRODHoras de Produccion (Orden de Fabricacion)")
    Call REGAPP("ECOPARAM", "TIE-TOTATotal de Horas (Orden de Fabricacion)")
    Call REGAPP("ECOPARAM", "NUOPERS Numero de Operarios (Orden de Fabricacion)")
    Call REGAPP("ECOPARAM", ".")
    '
    Call REGAPP("ECOPARAM", "TOT-PREPSuma Horas Preparacion (Orden de Fabricacion)")
    Call REGAPP("ECOPARAM", "TOT-PRODSuma Horas Produccion (Orden de Fabricacion)")
    Call REGAPP("ECOPARAM", "TOT-TIEMSuma Total de Horas (Orden de Fabricacion)")
    Call REGAPP("ECOPARAM", ".")
    '
    Call REGAPP("ECOPARAM", "NRO-REMINumero de Remito o Resumen de Entregas")
    Call REGAPP("ECOPARAM", "NRO-OCOMNumero de Orden de Compra")
    Call REGAPP("ECOPARAM", "NRO-AENTNumero de Aut.Entregas o Pedido")
    Call REGAPP("ECOPARAM", "IMP-LET1Importe en Letras - Primer Renglon")
    Call REGAPP("ECOPARAM", "IMP-LET2.                 - Segundo Renglon")
    Call REGAPP("ECOPARAM", "IMP-LET3.                 - Tercer Renglon")
    Call REGAPP("ECOPARAM", ".")
    Call REGAPP("ECOPARAM", "TRANSPORTransportista")
    Call REGAPP("ECOPARAM", "DOMI-TRADomicilio del Transportista")
    Call REGAPP("ECOPARAM", "LOCA-TRALocalidad del Transportista")
    Call REGAPP("ECOPARAM", "CUIT-TRANumero de C.U.I.T.")
    Call REGAPP("ECOPARAM", "PATE-TRANumero de Patente")
    Call REGAPP("ECOPARAM", "COND-TRAConductor")
    Call REGAPP("ECOPARAM", "DOCUCONDDocumento del Conductor")
    Call REGAPP("ECOPARAM", ".")
    Call REGAPP("ECOPARAM", "VALO-DECValor Declarado (Remito)")
    Call REGAPP("ECOPARAM", "NUBULTOSNumero de Bultos (Remito)")
    Call REGAPP("ECOPARAM", ".")
    '
    Call REGAPP("ECOPARAM", "RASO-CLIRazon Social de Cliente / Proveedor")
    Call REGAPP("ECOPARAM", "DOMI-CLIDomicilio de Cliente / Proveedor")
    Call REGAPP("ECOPARAM", "CPOS-CLICodigo Postal Destinatario")
    Call REGAPP("ECOPARAM", "LOCA-CLILocalidad Destinatario")
    Call REGAPP("ECOPARAM", "PCIA-CLIProvincia Destinatario")
    Call REGAPP("ECOPARAM", "PAIS-CLIPaís Destinatario")
    Call REGAPP("ECOPARAM", "TELE-CLINúmero de Telefono Destinatario")
    Call REGAPP("ECOPARAM", "FAX-CLI Número de FAX Destinatario")
    Call REGAPP("ECOPARAM", "CUIT-CLINumero de C.U.I.T. Destinatario")
    Call REGAPP("ECOPARAM", "IBRU-CLINumero de Ingresos Brutos")
    Call REGAPP("ECOPARAM", "CUEN-CLINumero de Cuenta Cliente / Proveedor")
    Call REGAPP("ECOPARAM", "PROV-CLINumero de Proveedor")
    Call REGAPP("ECOPARAM", "PIVA-CLIPosicion de IVA Cliente / Proveedor")
    Call REGAPP("ECOPARAM", "CIVA-CLICategoria de IVA ( 'A' - 'B' - 'C' )")
    Call REGAPP("ECOPARAM", ".")
    '
    Call REGAPP("ECOPARAM", "ORD-ITEMNumero de Orden del Item")
    Call REGAPP("ECOPARAM", "CANTIDADCantidad del Item")
    Call REGAPP("ECOPARAM", "CAKILOS Kilaje Total del Item")
    Call REGAPP("ECOPARAM", "COD-MAT C¢digo de Articulo")
    Call REGAPP("ECOPARAM", "DES-MAT Descripcion del Material")
    Call REGAPP("ECOPARAM", "ESP-MAT Especificacion del Material")
    Call REGAPP("ECOPARAM", "UNIMED  Unidad de medida")
    Call REGAPP("ECOPARAM", "REF-MAT1Referencia - Especificacion - Despacho a Plaza")
    Call REGAPP("ECOPARAM", "REF-MAT2Referencia - 2")
    Call REGAPP("ECOPARAM", "REF-MAT3Referencia - 3")
    Call REGAPP("ECOPARAM", "NRO-LOTEIdentificador de Lote")
    Call REGAPP("ECOPARAM", "TALLE   Talle")
    Call REGAPP("ECOPARAM", "COLOR   Color")
    Call REGAPP("ECOPARAM", "MEDIDA-1Medida 1 - Largo")
    Call REGAPP("ECOPARAM", "MEDIDA-2Medida 2 - Ancho")
    Call REGAPP("ECOPARAM", "MEDIDA-3Medida 3 - Alto")
    Call REGAPP("ECOPARAM", "N-PLANO Numero de Plano")
    Call REGAPP("ECOPARAM", "R-PLANO Revision de Plano")
    Call REGAPP("ECOPARAM", "F-PLANO Fecha de Revision de Plano")
    Call REGAPP("ECOPARAM", "Z-PLANO Plazo de Plano")
    Call REGAPP("ECOPARAM", "N-HPROF Numero de Hoja de Procesos de Fabricacion")
    Call REGAPP("ECOPARAM", "N-HINSP Numero de Hoja de Inspeccion")
    Call REGAPP("ECOPARAM", "N-ORIGINNumero Original de Pieza (s/lista de precios")
    Call REGAPP("ECOPARAM", "N-OCOMPRNumero de Orden de Compra")
    Call REGAPP("ECOPARAM", "N-AUTENTNumero de Autorizacion de Entregas")
    Call REGAPP("ECOPARAM", "N-REMITONumero de Remito")
    Call REGAPP("ECOPARAM", "N-BOLRECNumero de Comprobante de Recepcion")
    Call REGAPP("ECOPARAM", "FECH-RECFecha de Recepcion / Entrega")
    Call REGAPP("ECOPARAM", "PRE-LISTPrecio Unitario de Lista")
    Call REGAPP("ECOPARAM", "PORDESCUPorcentaje de Descuento")
    Call REGAPP("ECOPARAM", "PRE-NETOPrecio Unitario Neto")
    Call REGAPP("ECOPARAM", "NET-ITEMImporte Neto DEL ITEM")
    Call REGAPP("ECOPARAM", "ALICUIVAAlicuota - Porcentaje de IVA")
    Call REGAPP("ECOPARAM", "IVA-ITEMImporte de IVA del Item")
    Call REGAPP("ECOPARAM", "BRU-ITEMImporte Bruto del Item")
    Call REGAPP("ECOPARAM", ".")
    '
    Call REGAPP("ECOPARAM", "ENCA-DETEncabezado de Detalle de Facturacion")
    Call REGAPP("ECOPARAM", "ESPE-DETEspecificacion Final de Detalle")
    Call REGAPP("ECOPARAM", ".")
    '
    Call REGAPP("ECOPARAM", "IMP-NETOTotal Neto Factura")
    Call REGAPP("ECOPARAM", "POR-DESCPorcentaje de Descuento")
    Call REGAPP("ECOPARAM", "IMP-DESCImporte de Descuento")
    Call REGAPP("ECOPARAM", "IMPNEDESImporte Neto de Descuento")
    Call REGAPP("ECOPARAM", "POR-IVA Porcentaje de I.V.A. - Tasa Básica")
    Call REGAPP("ECOPARAM", "POR-IVA1 Porcentaje de I.V.A. - Tasa Básica")
    Call REGAPP("ECOPARAM", "IMP-IVA Importe de I.V.A. - Tasa Básica")
    Call REGAPP("ECOPARAM", "IMP-IVA1Importe de I.V.A. - Tasa Secundaria")
    Call REGAPP("ECOPARAM", "IMP-RNI Importe de I.V.A. Resp.No Inscripto")
    Call REGAPP("ECOPARAM", "ITO-IVA Importe Total de I.V.A.")
    Call REGAPP("ECOPARAM", "IMP-VARSImpuestos Varios")
    Call REGAPP("ECOPARAM", "PERC-IVAImporte Percepciones I.V.A.")
    Call REGAPP("ECOPARAM", "IMP-TOTAImporte Total")
    Call REGAPP("ECOPARAM", "TOTA-DOLImporte Total en Dolares")
    Call REGAPP("ECOPARAM", ".")
    Call REGAPP("ECOPARAM", "LIN-CIE Linea de Cierre de Detalle")
    Call REGAPP("ECOPARAM", "TOTA-ITSTotal de Items")
    Call REGAPP("ECOPARAM", "LETOT-ITLeyenda 'Total de Items'")
    Call REGAPP("ECOPARAM", ".")
    '
    Call REGAPP("ECOPARAM", "TIPO-VALTipo de Valor (Recibo)")
    Call REGAPP("ECOPARAM", "NUME-CHENumero de Cheque")
    Call REGAPP("ECOPARAM", "BANC-CHEBanco Emisor del Cheque")
    Call REGAPP("ECOPARAM", "IMPO-CHEImporte del Cheque")
    Call REGAPP("ECOPARAM", "FECH-CHEFecha Vencimiento de Cheque")
    Call REGAPP("ECOPARAM", "DOCU-APLDocumento de Aplicacion")
    Call REGAPP("ECOPARAM", "IMPO-PENSaldo Pendiente Documento")
    Call REGAPP("ECOPARAM", "IMPO-CANImporte Parcial Aplicado")
    Call REGAPP("ECOPARAM", ".")
    '
    Call REGAPP("ECOPARAM", "TOTA-VALTotal Valores Recibidos (Recibo)")
    Call REGAPP("ECOPARAM", "TOTA-CANTotal Cancelado (Recibo)")
    Call REGAPP("ECOPARAM", ".")
    '
    Call REGAPP("ECOPARAM", "REF-ASIEReferencia (Asiento Contable)")
    Call REGAPP("ECOPARAM", "NRO-PASENúmero de Pase (As.Contable)")
    Call REGAPP("ECOPARAM", "COD-CUENCódigo de Cuenta (As.Contable)")
    Call REGAPP("ECOPARAM", "DEN-CUENDenominacion de la Cuenta")
    Call REGAPP("ECOPARAM", "DET-PASEDetalle - Referencia del Pase")
    Call REGAPP("ECOPARAM", "IMP-DEBEImporte Debe (As.Contable)")
    Call REGAPP("ECOPARAM", "IMP-HABEImporte Haber (As.Contable)")
    Call REGAPP("ECOPARAM", "TOT-DEBETotal Debe (Asiento Contable)")
    Call REGAPP("ECOPARAM", "TOT-HABETotal Haber (Asiento Contable)")
    Call REGAPP("ECOPARAM", ".")
    '
    Call REGAPP("ECOPARAM", "CANT-MINStock Mínimo (Situación de Stocks)")
    Call REGAPP("ECOPARAM", "CANT-RESStock Comprometido (Situación de Stocks)")
    Call REGAPP("ECOPARAM", "CANT-INGStock a Ingresar (Situación de Stocks)")
    Call REGAPP("ECOPARAM", "CANT-DISStock Disponible (Situación de Stocks)")
    Call REGAPP("ECOPARAM", "CANT-FALStock Faltante (Situación de Stocks)")
    Call REGAPP("ECOPARAM", ".")
    '
    Call CIERRARCH("ECOPARAM")
    '
100 Call SELECHO("ECOPARAM")
    CODVARI$ = TRIM$(VALACT1$("ECOPARAM"))
    If CODVARI$ = "." Then GoTo 100
    If CODVARI$ = "" Then
        Call CIERRARCH("*.*")
        Exit Sub
    End If
    DEVARI$ = VALACT2$("ECOPARAM")
    '
    NX% = FILEOPEN%(LU$ + CODFOR$ + ".FRM", 0, 128)
    Dim RF As String * 128
    For I% = 1 To LOF(NX%) / 128
        Get #NX%, I%, RF$
        If TRIM$(Left$(RF$, 8)) = CODVARI$ Then
            REGCO& = I%
            RFF$ = RF$
            GoTo 200
        End If
    Next I%
    RFF$ = String$(128, 0)
    Call REPLA(RFF$, CODVARI$, 1, 8)
    REGCO& = 1 + LOF(NX%) / 128
    '
200 Call REPLA(RFF$, DEVARI$, 9, 48)
    Call REPLA(RFF$, "w", 126, 1)
    '
    RFG$ = OBJPLA$("DEFPARAMW", RFF$)
    If RFG$ = "" Then
        Close #NX%: NX% = 0
        GoTo 100
    End If
    '
    RF$ = RFG$
    Put #NX%, REGCO&, RF$
    Close #NX%: NX% = 0
    GoTo 100
    '
End Sub
'
Sub SHOWCOMMVIE(COPRI%, ATRIP$)
   '
   If PAGINACTIVA% > CAPAGINAS% Then
      CAPAGINAS% = PAGINACTIVA%
   End If
10 On Error Resume Next
   Select Case COPRI%
   Case 1
     'Printer.Orientation = Val(TRIM$(ATRIP$))
   Case 2
     'Printer.NewPage
     'If PAGINACTIVA% = 0 Then
     '   EDITFORM.Picture1(0).Refresh
     '   EDITFORM.VScroll1.Min = 0
     '   EDITFORM.VScroll1.Value = 0
     '   EDITFORM.VScroll1.Max = 1
     '   If ALTUPAGINA > EDITFORM.Frame1.Height Then
     '      EDITFORM.VScroll1.Max = (ALTUPAGINA - EDITFORM.Frame1.Height) / 100
     '   End If
     '   EDITFORM.Picture1(0).Visible = True
     '   EDITFORM.Picture1(0).Refresh
     'End If
     PAGINACTIVA% = PAGINACTIVA% + 1
   Case 3
     'Printer.EndDoc
     'If PAGINACTIVA% = 0 Then
     '   EDITFORM.Picture1(0).Refresh
     '   EDITFORM.VScroll1.Min = 0
     '   EDITFORM.VScroll1.Value = 0
     '   EDITFORM.VScroll1.Max = 1
     '   If ALTUPAGINA > EDITFORM.Frame1.Height Then
     '      EDITFORM.VScroll1.Max = (ALTUPAGINA - EDITFORM.Frame1.Height) / 100
     '   End If
     '   EDITFORM.Picture1(0).Visible = True
     '   EDITFORM.Picture1(0).Refresh
     'End If
     PAGINACTIVA% = PAGINACTIVA% + 1
   Case 4
     ALTUPAGINA = Val(TRIM$(ATRIP$))
     For KKI% = 0 To CAPAGINAS%
       If EDITFORM.Picture1(KKI%).Height < ALTUPAGINA Then
         EDITFORM.Picture1(KKI%).Height = ALTUPAGINA
       End If
     Next KKI%
   Case 5
     EDITFORM.Picture1(PAGINACTIVA%).Width = Val(TRIM$(ATRIP$))
   Case 6
     EDITFORM.Picture1(PAGINACTIVA%).ScaleMode = Val(TRIM$(ATRIP$))
   Case 7
     EDITFORM.Picture1(PAGINACTIVA%).ForeColor = Val(TRIM$(ATRIP$))
   Case 8
     EDITFORM.Picture1(PAGINACTIVA%).DrawMode = Val(TRIM$(ATRIP$))
   Case 9
     EDITFORM.Picture1(PAGINACTIVA%).DrawWidth = Val(TRIM$(ATRIP$))
   Case 10
     EDITFORM.Picture1(PAGINACTIVA%).DrawStyle = Val(TRIM$(ATRIP$))
   Case 11
     EDITFORM.Picture1(PAGINACTIVA%).FillStyle = Val(TRIM$(ATRIP$))
   Case 12
     AA11 = Val(Mid$(ATRIP$, 1, 16))
     BB11 = Val(Mid$(ATRIP$, 17, 16))
     AA22 = Val(Mid$(ATRIP$, 33, 16))
     BB22 = Val(Mid$(ATRIP$, 49, 16))
     CCOLO = Val(Mid$(ATRIP$, 65, 16))
     If BB11 + 1000 > ALTUPAGINA Then ALTUPAGINA = BB11 + 1000
     If BB22 + 1000 > ALTUPAGINA Then ALTUPAGINA = BB22 + 1000
     For KKI% = 0 To CAPAGINAS%
       If EDITFORM.Picture1(KKI%).Height < ALTUPAGINA Then
         EDITFORM.Picture1(KKI%).Height = ALTUPAGINA
       End If
     Next KKI%
     If AA11 > EDITFORM.Picture1(PAGINACTIVA%).Width Then EDITFORM.Picture1(PAGINACTIVA%).Width = AA11
     If AA22 > EDITFORM.Picture1(PAGINACTIVA%).Width Then EDITFORM.Picture1(PAGINACTIVA%).Width = AA22
     BBLO$ = UCase$(TRIM$(Mid$(ATRIP$, 81, 16)))
     If BBLO$ = "B" Then
        EDITFORM.Picture1(PAGINACTIVA%).Line (AA11, BB11 + TOPEPAGINA)-(AA22, BB22 + TOPEPAGINA), CCOLO, B
       Else
        EDITFORM.Picture1(PAGINACTIVA%).Line (AA11, BB11 + TOPEPAGINA)-(AA22, BB22 + TOPEPAGINA), CCOLO
     End If
   Case 13
     EDITFORM.Picture1(PAGINACTIVA%).FontName = TRIM$(ATRIP$)
   Case 14
     EDITFORM.Picture1(PAGINACTIVA%).FontItalic = Val(TRIM$(ATRIP$))
   Case 15
     EDITFORM.Picture1(PAGINACTIVA%).FontSize = Val(TRIM$(ATRIP$))
   Case 16
     EDITFORM.Picture1(PAGINACTIVA%).FontBold = Val(TRIM$(ATRIP$))
   Case 17
     If Val(TRIM$(ATRIP$)) > EDITFORM.Picture1(PAGINACTIVA%).Width Then EDITFORM.Picture1(PAGINACTIVA%).Width = Val(TRIM$(ATRIP$))
     EDITFORM.Picture1(PAGINACTIVA%).CurrentX = Val(TRIM$(ATRIP$))
   Case 18
     If Val(TRIM$(ATRIP$)) + 1000 > ALTUPAGINA Then ALTUPAGINA = Val(TRIM$(ATRIP$)) + 1000
     For KKI% = 0 To CAPAGINAS%
       If EDITFORM.Picture1(KKI%).Height < ALTUPAGINA Then
         EDITFORM.Picture1(KKI%).Height = ALTUPAGINA
       End If
     Next KKI%
     EDITFORM.Picture1(PAGINACTIVA%).CurrentY = Val(TRIM$(ATRIP$)) + TOPEPAGINA
   Case 19
     ATRIP$ = RTrim$(ATRIP$)
     If Right$(ATRIP$, 1) = ";" Then
         ATRIP$ = Left$(ATRIP$, Len(ATRIP$) - 1)
         EDITFORM.Picture1(PAGINACTIVA%).Print ATRIP$;
        Else
         EDITFORM.Picture1(PAGINACTIVA%).Print ATRIP$
     End If
   Case 20
     EDITFORM.Picture1(PAGINACTIVA%).FontUnderline = Val(TRIM$(ATRIP$))
   Case 21
     Call INCLUDEIMAGE(ATRIP$, "SHOW")
   End Select
   On Error GoTo 0
   '
End Sub
'
Sub CARINDIFORM()
    '
    Screen.MousePointer = 11
    Dim RF As String * 128
    PRINFORM.File1.Path = LUDAT$
    For U& = 0 To PRINFORM.File1.ListCount - 1
        PRINFORM.File1.ListIndex = U&
        ARXYZ$ = UCase$(PRINFORM.File1.filename)
        PPXX% = InStr(ARXYZ$, ".FRM")
        If PPXX% > 1 Then
           ARXYZ$ = TRIM$(Left$(ARXYZ$, PPXX% - 1))
           If ARXYZ$ <> "" Then
               NX% = FILEOPEN%(LU$ + ARXYZ$ + ".FRM", 0, 128)
               TITUFOR$ = ""
               For I% = 1 To LOF(NX%) / 128
                   Get #NX%, I%, RF$
                   If TRIM$(Left$(RF$, 8)) = "TITUFORM" Then
                       TITUFOR$ = Mid$(RF$, 9, 48)
                       GoTo 10
                   End If
               Next I%
10             Close #NX%: NX% = 0
               '
               For V& = 1 To ULTREG&("INDIFRM")
                   INF$ = REGLEIDO$("INDIFRM", V&)
                   If Left$(INF$, 8) = AJUSTI$(ARXYZ$, 8) Then
                       If TITUFOR$ <> "" Then
                           Call REPLA(INF$, TITUFOR$, 9, 48)
                           Call GRAREG("INDIFRM", INF$, V&)
                       End If
                       GoTo 19
                   End If
               Next V&
               '
               INF$ = REGBLAN$("INDIFRM")
               Call REPLA(INF$, ARXYZ$, 1, 8)
               Call REPLA(INF$, TITUFOR$, 9, 48)
               Call REGAPP("INDIFRM", INF$)
               '
           End If
        End If
19  Next U&
    '
    Call CIERRARCH("INDIFRM")
    Call FRESHECHO("INDIFRM")
    Screen.MousePointer = 1
    '
End Sub
'

Sub GRATITUFORM()
    '
    Screen.MousePointer = 11
    If TRIM$(PRINFORM.Text1.TEXT) = TRIM$(CODFOR$) Then
        If TRIM$(CODFOR$) <> "" Then
            ARXYZ$ = TRIM$(CODFOR$)
            TITUFOR$ = PRINFORM.Text2.TEXT
            NX% = FILEOPEN%(LU$ + ARXYZ$ + ".FRM", 0, 128)
            Dim RF As String * 128
            '
            For I% = 1 To LOF(NX%) / 128
                Get #NX%, I%, RF$
                If TRIM$(Left$(RF$, 8)) = "TITUFORM" Then
                    RFY$ = RF$
                    Call REPLA(RFY$, TITUFOR$, 9, 48)
                    LSet RF$ = RFY$
                    Put #NX%, I%, RF$
                    GoTo 10
                End If
            Next I%
            '
            I% = 1 + LOF(NX%) / 128
            RFY$ = Space$(56) + String$(72, 0)
            Call REPLA(RFY$, "TITUFORM", 1, 8)
            Call REPLA(RFY$, TITUFOR$, 9, 48)
            LSet RF$ = RFY$
            Put #NX%, I%, RF$
            '
10          Close #NX%: NX% = 0
            '
            For V& = 1 To ULTREG&("INDIFRM")
                   INF$ = REGLEIDO$("INDIFRM", V&)
                   If Left$(INF$, 8) = AJUSTI$(ARXYZ$, 8) Then
                       If TITUFOR$ <> "" Then
                           Call REPLA(INF$, TITUFOR$, 9, 48)
                           Call GRAREG("INDIFRM", INF$, V&)
                       End If
                       GoTo 19
                   End If
            Next V&
            '
            INF$ = REGBLAN$("INDIFRM")
            Call REPLA(INF$, ARXYZ$, 1, 8)
            Call REPLA(INF$, TITUFOR$, 9, 48)
            Call REGAPP("INDIFRM", INF$)
            '
        End If
    End If
    '
19  Call CIERRARCH("INDIFRM")
    Call FRESHECHO("INDIFRM")
    Screen.MousePointer = 1
    '
End Sub
'
