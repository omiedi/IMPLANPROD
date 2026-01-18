VERSION 5.00
Object = "{D9AF33E0-7C55-11D5-9151-0000E856BC17}#1.0#0"; "fiscal010724.ocx"
Begin VB.Form FORHAS320 
   Caption         =   "CONTROLADOR HASAR320"
   ClientHeight    =   2415
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3600
   LinkTopic       =   "Form1"
   ScaleHeight     =   2415
   ScaleWidth      =   3600
   StartUpPosition =   3  'Windows Default
   Begin FiscalPrinterLibCtl.HASAR HASAR1 
      Left            =   1560
      OleObjectBlob   =   "FORHAS320.frx":0000
      Top             =   945
   End
End
Attribute VB_Name = "FORHAS320"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Sub REPORTES()
        '
        Call MUESTRACOMANDO("Puerto", 1)
        If MUECOMHAS% < 2 Then HASAR1.Puerto = 1
        '
        Call MUESTRACOMANDO("Modelo", MODELO_P320)
        If MUECOMHAS% < 2 Then HASAR1.Modelo = MODELO_P320          '//Este OCX no es 100% para 715F
        
        On Error Resume Next
        '
        Call MUESTRACOMANDO("Comenzar")
        If MUECOMHAS% < 2 Then HASAR1.Comenzar
        '
        Call MUESTRACOMANDO("PrecioBase", False)
        If MUECOMHAS% < 2 Then HASAR1.PrecioBase = False
        '
        Call MUESTRACOMANDO("TratarDeCancelarTodo")
        If MUECOMHAS% < 2 Then HASAR1.TratarDeCancelarTodo
        '
        On Error GoTo 0
        '
       OPXX% = ALTERNA%("Reporte Z\Reporte X")
       If OPXX% = 1 Then
           Call reporteZ
         ElseIf OPXX% = 2 Then
           Call reporteX
       End If
       On Error Resume Next
       '
       Call MUESTRACOMANDO("Finalizar")
       If MUECOMHAS% < 2 Then HASAR1.Finalizar
       On Error GoTo 0

End Sub

Sub reporteZ()
On Error GoTo impresora_apag

Procesar:

    Call MUESTRACOMANDO("reporteZ")
    If MUECOMHAS% < 2 Then HASAR1.reporteZ
    Exit Sub

impresora_apag:
    If Abs(Err.Number) = 2147220937 Then
        Call MENSERR(24, "El Puerto se Encuentra Abierto.\Cierre el Módulo y Vuelva a Intentarlo")
        Exit Sub
    End If
    If MsgBox("Error Impresora:" & Err.Description, vbRetryCancel, "Errores") = vbRetry Then
        Resume Procesar
    End If
End Sub

Sub reporteX()
On Error GoTo impresora_apag

Procesar:
    '
    Call MUESTRACOMANDO("reporteX")
    If MUECOMHAS% < 2 Then HASAR1.reporteX
    Exit Sub
    '
impresora_apag:

    If Abs(Err.Number) = 2147220937 Then
        Call MENSERR(24, "El Puerto se Encuentra Abierto.\Cierre el Módulo y Vuelva a Intentarlo")
        Exit Sub
    End If
    If MsgBox("Error Impresora:" & Err.Description, vbRetryCancel, "Errores") = vbRetry Then
        Resume Procesar
    End If
    
End Sub

Sub INICIA()
    Call MUESTRACOMANDO("Puerto", 1)
    If MUECOMHAS% < 2 Then HASAR1.Puerto = 1
    '
    Call MUESTRACOMANDO("Modelo", MODELO_P320)
    If MUECOMHAS% < 2 Then HASAR1.Modelo = MODELO_P320             ' //Este OCX no es 100% para 715F
    '
    On Error Resume Next
    '
    Call MUESTRACOMANDO("Comenzar")
    If MUECOMHAS% < 2 Then HASAR1.Comenzar
    '
    If Err Then
      On Error GoTo 0
      TTXX$ = "Imposible Conectar con la Impresora Fiscal.\Revise el Encendido y las Conexiones\o Re-Configure para Impresora Común.\  \En Caso de Continuar, los Comandos se Presentarán\en Pantalla, pero NO se Enviarán a la Impresora."
      If COMALTER%(TTXX$ + "\\Configurar\Continuar") <> 2 Then
         Call FORFACPRINT.Command5_Click
         Exit Sub
      End If
      MUECOMHAS% = 2
    End If
    '
    Call MUESTRACOMANDO("PrecioBase", False)
    If MUECOMHAS% < 2 Then HASAR1.PrecioBase = False
    '
    Call MUESTRACOMANDO("TratarDeCancelarTodo")
    If MUECOMHAS% < 2 Then HASAR1.TratarDeCancelarTodo
    '
    On Error GoTo 0
    '
End Sub

Sub PRITIFACA()
    '
    Dim COMANDO As String
    Dim FS As String
    FS$ = Chr$(28)                                            '// Separador de campos del comando
    '
    On Error GoTo impresora_apag
    '
Procesar:
    '
    NCLIE% = XVALO(FORFACPRINT.Text1) ' VALIDACION DE QUE TENGA CUIT EL CLIENTE SI NO MUESTRA MENSAJE
    CUITT$ = CUITCLI$(NCLIE%)
    If VALICUIT(CUITT$) < 1 Then
      Call MENSERR(24, "Esta impresora fiscal requiere para los comprobantes del tipo 'A' que el CUIT\del contribuyente sea válido. El cliente actual posee un CUIT inválido,\por lo que el comprobante ha quedado registrado en su cuenta corriente\pero el mismo no se podrá imprimir. ")
    End If
    '
    TIDOCUX = TIPO_CUIT
    CUITT$ = SOLONUM$(CUITT$)
    '
    PVCL% = PIVACLI%(NCLIE%)
    RESPONIVA = RESPONSABLE_INSCRIPTO
    If PVCL% < 1 Then RESPONIVA = CONSUMIDOR_FINAL
    If PVCL% = 2 Then RESPONIVA = RESPONSABLE_NO_INSCRIPTO
    If PVCL% = 3 Then RESPONIVA = NO_RESPONSABLE
    If PVCL% = 4 Then RESPONIVA = RESPONSABLE_EXENTO
    If PVCL% = 5 Then RESPONIVA = NO_CATEGORIZADO
    If PVCL% = 6 Then RESPONIVA = MONOTRIBUTO

'    TTXX$ = Chr$(98) & FS & STRINOM$(NCLIE%)
'        CUITT$ = SOLONUM$(CUITCLI$(NCLIE%))
'    TTXX$ = TTXX$ + FS + CUITT$
'    TTXX$ = TTXX$ + FS + "I"    'RESPONSABLE INSCRIPTO--> VER MANUAL (PUBLTICK) PAG. 75
'    TTXX$ = TTXX$ + FS + "C"    'C.U.I.T.--> VER MANUAL (PUBLTICK) PAG. 75
'    TTXX$ = TTXX$ + FS + STRIDOMI$(NCLIE%)
'    TTXX$ = Left$(TTXX$, 85)  ' SI NO DA ERROR
'    Call MUESTRACOMANDO("Enviar", TTXX$)
'    If MUECOMHAS% < 2 Then HASAR1.Enviar TTXX$         ' IMPRIME EL ENCABEZADO DEL TICKET
'            ' ATENCION: Unicamente en la Factura A se utiliza el método ENVIAR, porque
'            ' es el que permite incluir el domicilio del cliente. En las "B" se usa
'            ' método DATOSCLIENTE
'            '
    ' ENCABEZADO - probamos aquí  si funciona
    On Error Resume Next
    HASAR1.Encabezado(1) = "VIAMONTE 2255"
    HASAR1.Encabezado(2) = "C1056ABI - CAPITAL FEDERAL"
      TELE$ = "Tel: 5411-4953-6114  /  6111"
    HASAR1.Encabezado(3&) = TELE$
      CTE$ = TRIM$(FORFACPRINT.Text1) & " - " & TRIM$(FORFACPRINT.Label26)
    HASAR1.Encabezado(4&) = "Cliente: " & CTE$
    HASAR1.Encabezado(5&) = ""
    On Error GoTo 0
    '
    Call MUESTRACOMANDO("DatosCliente", STRINOM$(NCLIE%), CUITT$, TIDOCUX, RESPONIVA)
    If MUECOMHAS% < 2 Then HASAR1.DatosCliente STRINOM$(NCLIE%), CUITT$, TIDOCUX, RESPONIVA, STRIDOMI$(NCLIE%)
    '
    Call MUESTRACOMANDO("AbrirComprobanteFiscal", FACTURA_A)
    If MUECOMHAS% < 2 Then HASAR1.AbrirComprobanteFiscal FACTURA_A
    '
    Call VERIFICAELNUMERO
    '
    FIN& = ULTREG&("!CARDEFAC")
    For U& = 1 To FIN&
        XXX$ = REGLEIDO$("!CARDEFAC", U&)
        CI% = CVI(Mid$(XXX$, 5, 2))
        CODIGO$ = REPLACEN$(CODEXT$(CI%))
        DESCR$ = REPLACEN$(TRIM$(Mid$(XXX$, 7, 64)))
        CANTIDA1# = CVS(Left$(XXX$, 4))
        PREUNI# = CVS(Mid$(XXX$, 429, 4))
        ALICUIVA# = CVS(Mid$(XXX$, 441, 4))
        PRECONIVA# = (PREUNI * (1 + ALICUIVA# / 100))
        ACUMUIMPBRU = ACUMUIMPBRU + PRECONIVA# * CANTIDA1#
        Call MUESTRACOMANDO("ImprimirItem", STRICODES$(CI%, DESCR$), CANTIDA1#, PRECONIVA#, ALICUIVA#, 0)
        If MUECOMHAS% < 2 Then HASAR1.ImprimirItem STRICODES$(CI%, DESCR$), CANTIDA1#, PRECONIVA#, ALICUIVA#, 0
    Next U&
    '
    IDESCUENTO = ACUMUIMPBRU * (XVALO(FORFACPRINT.PORDESCU) / 100)
    'IDESCUENTO = XVALO(FORFACPRINT.Total(1)) * (1 + ALIVA / 100)
    Call MUESTRACOMANDO("DescuentoGeneral", "Bonificacion", IDESCUENTO, True)
    If MUECOMHAS% < 2 Then HASAR1.DescuentoGeneral "Bonificacion", IDESCUENTO, True           'INDICA UN IMPORTE, NO PORCENTAJE: TRUE = DESCUENTO, FALSE = INCREMENTO
    '
    Call MUESTRACOMANDO("Subtotal", True)
    If MUECOMHAS% < 2 Then HASAR1.Subtotal True
    '
    ' PIE DE PAGINA
    On Error Resume Next
      CPAG$ = REPLACEN$(TRIM$(FORFACPRINT.Eco21(1)))
      LETRACOM$ = REPLACEN$(TRIM$(FORFACPRINT.Label28))
    HASAR1.Encabezado(11) = "C-Pago: " & CPAG$ & " - Letra = " & LETRACOM$
    Call FRATEXTO(FORFINFAC.Text4, 32)
    For KKII% = 1 To 3
        If KKII% <= 3 Then
          OBS$ = RETEX$(KKII%)
          HASAR1.Encabezado(KKII% + 11) = REPLACEN$(OBS$)
        End If
    Next KKII%
    On Error GoTo 0
    '
    Call MUESTRACOMANDO("CerrarComprobanteFiscal")
    If MUECOMHAS% < 2 Then HASAR1.CerrarComprobanteFiscal
    '
    Exit Sub

impresora_apag:

    If MsgBox("Error de Impresora:" & Err.Description, vbRetryCancel, "Errores") = vbRetry Then
        Resume Procesar
    End If

End Sub

Sub PRITIFACB()
    '
    Dim COMANDO As String
    'Dim FS As String
    'FS = Chr$(28)                                            '// Separador de campos del comando
    '
    On Error GoTo impresora_apag
           
Procesar:
    '
    On Error Resume Next
    NCX% = XVALO(FORFACPRINT.Text1)
    CUITT$ = SOLONUM$(CUITCLI$(NCX%))
    On Error GoTo 0
    TIDOCUX = TIPO_CUIT
    If PIVACLI%(NCX%) < 1 Then TIDOCUX = TIPO_DNI
    '
    PVCL% = PIVACLI%(NCX%)
    RESPONIVA = RESPONSABLE_INSCRIPTO
    If PVCL% < 1 Then RESPONIVA = CONSUMIDOR_FINAL
    If PVCL% = 2 Then RESPONIVA = RESPONSABLE_NO_INSCRIPTO
    If PVCL% = 3 Then RESPONIVA = NO_RESPONSABLE
    If PVCL% = 4 Then RESPONIVA = RESPONSABLE_EXENTO
    If PVCL% = 5 Then RESPONIVA = NO_CATEGORIZADO
    If PVCL% = 6 Then RESPONIVA = MONOTRIBUTO
    '
    If PVCL% < 1 Then  ' VALIDA QUE POSEA DNI CUANDO ES CONSUMIDOR FINAL
      TOTALFACTU# = XVALO(FORFACPRINT.Total(6))
      If (TOTALFACTU# >= 1000) And CUITT$ = "" Then
       Call MENSERR(24, "Las Facturas Mayores a $1000 sobre Consumidores Finales Requieren DNI")
       Call FINAL("") ' SI NO CUMPLE CANCELA TODO
      End If
    End If
    '
    ' ENCABEZADO
    On Error Resume Next
    HASAR1.Encabezado(1) = "VIAMONTE 2255"
    HASAR1.Encabezado(2) = "C1056ABI - CAPITAL FEDERAL"
      TELE$ = "Tel: 5411-4953-6114/6111"
    HASAR1.Encabezado(3&) = TELE$
      CTE$ = TRIM$(FORFACPRINT.Text1) & " - " & TRIM$(FORFACPRINT.Label26)
    HASAR1.Encabezado(4&) = "Cliente: " & CTE$
    HASAR1.Encabezado(5&) = ""
    On Error GoTo 0
    '
    Call MUESTRACOMANDO("DatosCliente", STRINOM$(NCX%), CUITT$, TIDOCUX, RESPONIVA)
    If MUECOMHAS% < 2 Then HASAR1.DatosCliente STRINOM$(NCX%), CUITT$, TIDOCUX, RESPONIVA, STRIDOMI$(NCX%)
    'IGUALMENTE SOLO IMPRIME LA LEYENDA CONSUMIDOR FINAL
    '
    Call MUESTRACOMANDO("AbrirComprobanteFiscal", FACTURA_B)
    If MUECOMHAS% < 2 Then HASAR1.AbrirComprobanteFiscal FACTURA_B
    '
    Call VERIFICAELNUMERO
    '
    FIN& = ULTREG&("!CARDEFAC")
    For U& = 1 To FIN&
        XXX$ = REGLEIDO$("!CARDEFAC", U&)
        CI% = CVI(Mid$(XXX$, 5, 2))
        CODIGO$ = REPLACEN$(CODEXT$(CI%))
        DESCR$ = REPLACEN$(TRIM$(Mid$(XXX$, 7, 64)))
        CANTIDA1# = ROUND(CVS(Left$(XXX$, 4)), 1)
        PREUNI# = CVS(Mid$(XXX$, 429, 4))
        ALICUIVA# = ROUND(CVS(Mid$(XXX$, 441, 4)), 2)
        PRECONIVA# = ROUND(PREUNI# * (1 + ALICUIVA# / 100))
        ACUMUIMPBRU = ACUMUIMPBRU + ROUND(PRECONIVA# * CANTIDA1#)
        '
        Call MUESTRACOMANDO("ImprimirItem", STRICODES$(CI%, DESCR$), CANTIDA1#, PRECONIVA#, ALICUIVA#, 0)
        If MUECOMHAS% < 2 Then HASAR1.ImprimirItem STRICODES$(CI%, DESCR$), CANTIDA1#, PRECONIVA#, ALICUIVA#, 0
        '
    Next U&
    '
    IDESCUENTO = ACUMUIMPBRU * (XVALO(FORFACPRINT.PORDESCU) / 100)
    'IDESCUENTO = XVALO(FORFACPRINT.Total(1)) * (1 + ALIVA / 100)
    '
    Call MUESTRACOMANDO("DescuentoGeneral", "Bonificacion", IDESCUENTO, True)
    If MUECOMHAS% < 2 And IDESCUENTO > 0 Then HASAR1.DescuentoGeneral "Bonificacion", IDESCUENTO, True        'INDICA UN IMPORTE, NO PORCENTAJE: TRUE = DESCUENTO, FALSE = INCREMENTO
    '
    Call MUESTRACOMANDO("SubTotal", True)
    If MUECOMHAS% < 2 Then HASAR1.Subtotal True
    '
    ' PIE DE PAGINA
    On Error Resume Next
      CPAG$ = TRIM$(FORFACPRINT.Eco21(1))
      LETRACOM$ = REPLACEN$(TRIM$(FORFACPRINT.Label28))
    HASAR1.Encabezado(11) = "C-Pago: " & CPAG$ & " - Letra = " & LETRACOM$
    Call FRATEXTO(FORFINFAC.Text4, 32)
    For KKII% = 1 To 3
        If KKII% <= 3 Then
          OBS$ = RETEX$(KKII%)
          HASAR1.Encabezado(KKII% + 11) = REPLACEN$(OBS$)
        End If
    Next KKII%
    On Error GoTo 0
    '
    Call MUESTRACOMANDO("CerrarComprobanteFiscal")
    If MUECOMHAS% < 2 Then HASAR1.CerrarComprobanteFiscal
    '
    Exit Sub

impresora_apag:

    If MsgBox("Error Impresora:" & Err.Description, vbRetryCancel, "Errores") = vbRetry Then
        Resume Procesar
    End If

End Sub

Sub PRINOCREA()
   
   Dim COMANDO As String
   Dim FS As String
   FS$ = Chr$(28)                                            '// Separador de campos del comando
   '
   On Error GoTo impresora_apag
       
Procesar:
    ' **************************************************
    ' LA IMPRESION DE LA NOTA DE CREDITO DEBE REALIZARSE
    ' MEDIANTE EL METODO ENVIAR PORQUE LA OCX NO SOPORTA
    ' ESTE TIPO DE COMPROBANTE
    ' **************************************************
    NCLIE% = XVALO(FORFACPRINT.Text1) ' VALIDACION DE QUE TENGA CUIT EL CLIENTE SI NO MUESTRA MENSAJE
    CUITT$ = CUITCLI$(NCLIE%)
    If VALICUIT(CUITT$) < 1 Then
      Call MENSERR(24, "Esta impresora fiscal requiere para los comprobantes del tipo 'A' que el CUIT\del contribuyente sea válido. El cliente actual posee un CUIT inválido,\por lo que el comprobante ha quedado registrado en su cuenta corriente\pero el mismo no se podrá imprimir. ")
    End If
    '
'    CUITT$ = SOLONUM$(CUITT$)
'    TIDOCUX = TIPO_CUIT
'
'    PVCL% = PIVACLI%(NCLIE%)
'    RESPONIVA = RESPONSABLE_INSCRIPTO
'    If PVCL% < 1 Then RESPONIVA = CONSUMIDOR_FINAL
'    If PVCL% = 2 Then RESPONIVA = RESPONSABLE_NO_INSCRIPTO
'    If PVCL% = 3 Then RESPONIVA = NO_RESPONSABLE
'    If PVCL% = 4 Then RESPONIVA = RESPONSABLE_EXENTO
'    If PVCL% = 5 Then RESPONIVA = NO_CATEGORIZADO
'    If PVCL% = 6 Then RESPONIVA = MONOTRIBUTO

    COMANDO = Chr$(147) & FS & "1" & FS & "0000-0000000" ' comprobante original
    '
    Call MUESTRACOMANDO("Enviar", COMANDO)
    If MUECOMHAS% < 2 Then HASAR1.Enviar COMANDO
    '
    TTXX$ = Chr$(98) & FS & STRINOM$(NCLIE%)
    CUITT$ = SOLONUM$(CUITCLI$(NCLIE%))
    TTXX$ = TTXX$ + FS + CUITT$
    TTXX$ = TTXX$ + FS + "I"    'RESPONSABLE INSCRIPTO--> VER MANUAL (PUBLTICK) PAG. 75
    TTXX$ = TTXX$ + FS + "C"    'C.U.I.T.--> VER MANUAL (PUBLTICK) PAG. 75
    TTXX$ = TTXX$ + FS + STRIDOMI$(NCLIE%)
    TTXX$ = Left$(TTXX$, 85)  ' SI NO DA ERROR
    '
    Call MUESTRACOMANDO("Enviar", TTXX$)
    If MUECOMHAS% < 2 Then HASAR1.Enviar TTXX$               ' IMPRIME EL ENCABEZADO DEL TICKET
    '
    COMANDO = Chr$(128) & FS & "R" & FS & "T"           ' DEFINE EL TIPO DE COMPROBANTE
    '
    Call MUESTRACOMANDO("Enviar", COMANDO)
    If MUECOMHAS% < 2 Then HASAR1.Enviar COMANDO
    '
'    Call MUESTRACOMANDO("DatosCliente", STRINOM$(NCLIE%), CUITT$, TIDOCUX, RESPONIVA)
'    If MUECOMHAS% < 2 Then HASAR1.DatosCliente STRINOM$(NCLIE%), CUITT$, TIDOCUX, RESPONIVA, STRIDOMI$(NCLIE%)
'    '
'    Call MUESTRACOMANDO("AbrirComprobanteFiscal", NOTA_CREDITO_A)
'    If MUECOMHAS% < 2 Then HASAR1.AbrirComprobanteFiscal NOTA_CREDITO_A

    Call VERIFICAELNUMERO
    '
    ' ENCABEZADO
    On Error Resume Next
      TELE$ = "5411-4953-6114/6111"
    HASAR1.Encabezado(3) = "Tel: " & TELE$
      CTE$ = TRIM$(FORFACPRINT.Text1) & " - " & REPLACEN$(TRIM$(FORFACPRINT.Label26))
    HASAR1.Encabezado(4) = "Cliente: " & CTE$
    HASAR1.Encabezado(5) = ""
    On Error GoTo 0
    '
    FIN& = ULTREG&("!CARDEFAC")
    For U& = 1 To FIN&
        XXX$ = REGLEIDO$("!CARDEFAC", U&)
        CI% = CVI(Mid$(XXX$, 5, 2))
        DESCR$ = REPLACEN$(TRIM$(Mid$(XXX$, 7, 50)))
        CANTIDA1 = CVS(Left$(XXX$, 4))
        PREUNI = CVS(Mid$(XXX$, 429, 4))
        ALICUIVA = CVS(Mid$(XXX$, 441, 4))
        PRECONIVA = (PREUNI * (1 + ALICUIVA / 100))
        ACUMUIMPBRU = ACUMUIMPBRU + PRECONIVA * CANTIDA1
        '
        Call MUESTRACOMANDO("ImprimirItem", DESCR$, CANTIDA1, PRECONIVA, ALICUIVA, 0)
        If MUECOMHAS% < 2 Then HASAR1.ImprimirItem DESCR$, CANTIDA1, PRECONIVA, ALICUIVA, 0
        '
    Next U&
    '
    IDESCUENTO = ACUMUIMPBRU * (XVALO(FORFACPRINT.PORDESCU) / 100)
    'IDESCUENTO = XVALO(FORFACPRINT.Total(1)) * (1 + ALIVA / 100)
    '
    Call MUESTRACOMANDO("DescuentoGeneral", "Bonificacion", IDESCUENTO, True)
    If MUECOMHAS% < 2 Then HASAR1.DescuentoGeneral "Bonificacion", IDESCUENTO, True                'INDICA UN IMPORTE, NO PORCENTAJE: TRUE = DESCUENTO, FALSE = INCREMENTO
    '
    Call MUESTRACOMANDO("Subtotal", True)
    If MUECOMHAS% < 2 Then HASAR1.Subtotal True
    '
    ' PIE DE PAGINA
    On Error Resume Next
      CPAG$ = TRIM$(FORFACPRINT.Eco21(1))
    HASAR1.Encabezado(11) = "C-Pago: " & REPLACEN$(CPAG$)
    Call FRATEXTO(FORFINFAC.Text4, 32)
    For KKII% = 1 To 3
        If KKII% <= 3 Then
          OBS$ = RETEX$(KKII%)
          HASAR1.Encabezado(KKII% + 11) = REPLACEN$(OBS$)
        End If
    Next KKII%
    On Error GoTo 0
    '
    Call MUESTRACOMANDO("Enviar", Chr$(129))
    If MUECOMHAS% < 2 Then HASAR1.Enviar Chr$(129)
    '
    Exit Sub

impresora_apag:

    If MsgBox("Error Impresora:" & Err.Description, vbRetryCancel, "Errores") = vbRetry Then
        Resume Procesar
    End If

End Sub

Sub PRINOCREB()

   Dim COMANDO As String
   Dim FS As String
   FS$ = Chr$(28)                                            '// Separador de campos del comando
   '
   On Error GoTo impresora_apag
       
Procesar:
    ' **************************************************
    ' LA IMPRESION DE LA NOTA DE CREDITO DEBE REALIZARSE
    ' MEDIANTE EL METODO ENVIAR PORQUE LA OCX NO SOPORTA
    ' ESTE TIPO DE COMPROBANTE
    ' **************************************************
    NCLIE% = XVALO(FORFACPRINT.Text1) ' VALIDACION DE QUE TENGA CUIT EL CLIENTE SI NO MUESTRA MENSAJE
    CUITT$ = CUITCLI$(NCLIE%)
    If VALICUIT(CUITT$) < 1 Then
      Call MENSERR(24, "Esta impresora fiscal requiere para los comprobantes del tipo 'B' que el CUIT\del contribuyente sea válido. El cliente actual posee un CUIT inválido,\por lo que el comprobante ha quedado registrado en su cuenta corriente\pero el mismo no se podrá imprimir. ")
    End If
    '
'    CUITT$ = SOLONUM$(CUITT$)
'    TIDOCUX = TIPO_CUIT
'    If PIVACLI%(NCX%) < 1 Then TIDOCUX = TIPO_DNI
'    '
'    PVCL% = PIVACLI%(NCX%)
'    RESPONIVA = RESPONSABLE_INSCRIPTO
'    If PVCL% < 1 Then RESPONIVA = CONSUMIDOR_FINAL
'    If PVCL% = 2 Then RESPONIVA = RESPONSABLE_NO_INSCRIPTO
'    If PVCL% = 3 Then RESPONIVA = NO_RESPONSABLE
'    If PVCL% = 4 Then RESPONIVA = RESPONSABLE_EXENTO
'    If PVCL% = 5 Then RESPONIVA = NO_CATEGORIZADO
'    If PVCL% = 6 Then RESPONIVA = MONOTRIBUTO

    COMANDO = Chr$(147) & FS & "1" & FS & "0000-0000000" ' comprobante original
    '
    Call MUESTRACOMANDO("Enviar", COMANDO)
    If MUECOMHAS% < 2 Then HASAR1.Enviar COMANDO
    '
    TTXX$ = Chr$(98) & FS & STRINOM$(NCLIE%)
    CUITT$ = SOLONUM$(CUITCLI$(NCLIE%))
    TTXX$ = TTXX$ + FS + CUITT$
    TTXX$ = TTXX$ + FS + "C"            'C.U.I.T.--> VER MANUAL (PUBLTICK) PAG. 75
    TTXX$ = TTXX$ + FS + "2"
    TTXX$ = TTXX$ + FS + STRIDOMI$(NCLIE%)
    TTXX$ = Left$(TTXX$, 85)            ' SI NO DA ERROR
    '
    Call MUESTRACOMANDO("Enviar", TTXX$)
    If MUECOMHAS% < 2 Then HASAR1.Enviar TTXX$              ' IMPRIME EL ENCABEZADO DEL TICKET
    'IGUALMENTE SOLO IMPRIME LA LEYENDA CONSUMIDOR FINAL COMO ENCABEZADO
    '
    COMANDO = Chr$(128) & FS & "S" & FS & "T"       ' DEFINE EL TIPO DE COMPROBANTE
    '
    Call MUESTRACOMANDO("Enviar", COMANDO)
    If MUECOMHAS% < 2 Then HASAR1.Enviar COMANDO
    '
'    Call MUESTRACOMANDO("DatosCliente", STRINOM$(NCX%), CUITT$, TIDOCUX, RESPONIVA)
'    If MUECOMHAS% < 2 Then HASAR1.DatosCliente STRINOM$(NCX%), CUITT$, TIDOCUX, RESPONIVA, STRIDOMI$(NCX%)
'    '
'    Call MUESTRACOMANDO("AbrirComprobanteFiscal", NOTA_CREDITO_B)
'    If MUECOMHAS% < 2 Then HASAR1.AbrirComprobanteFiscal NOTA_CREDITO_B

    Call VERIFICAELNUMERO
    '
    ' ENCABEZADO
    On Error Resume Next
      TELE$ = "5411-4953-6114/6111"
    HASAR1.Encabezado(3) = "Tel: " & REPLACEN$(TELE$)
      CTE$ = TRIM$(FORFACPRINT.Text1) & " - " & REPLACEN$(TRIM$(FORFACPRINT.Label26))
    HASAR1.Encabezado(4) = "Cliente: " & CTE$
    HASAR1.Encabezado(5) = ""
    On Error GoTo 0
    '
    FIN& = ULTREG&("!CARDEFAC")
    For U& = 1 To FIN&
        XXX$ = REGLEIDO$("!CARDEFAC", U&)
        CI% = CVI(Mid$(XXX$, 5, 2))
        DESCR$ = REPLACEN$(TRIM$(Mid$(XXX$, 7, 50)))
        CANTIDA1 = CVS(Left$(XXX$, 4))
        PREUNI = CVS(Mid$(XXX$, 429, 4))
        ALICUIVA = CVS(Mid$(XXX$, 441, 4))
        PRECONIVA = (PREUNI * (1 + ALICUIVA / 100))
        ACUMUIMPBRU = ACUMUIMPBRU + PRECONIVA * CANTIDA1
        '
        Call MUESTRACOMANDO("ImprimirItem", DESCR$, CANTIDA1, PRECONIVA, ALICUIVA, 0)
        If MUECOMHAS% < 2 Then HASAR1.ImprimirItem DESCR$, CANTIDA1, PRECONIVA, ALICUIVA, 0
        '
    Next U&
    '
    IDESCUENTO = ACUMUIMPBRU * (XVALO(FORFACPRINT.PORDESCU) / 100)
    'IDESCUENTO = XVALO(FORFACPRINT.Total(1)) * (1 + ALIVA / 100)
    '
    Call MUESTRACOMANDO("DescuentoGeneral", "Bonificacion", IDESCUENTO, True)
    If MUECOMHAS% < 2 Then HASAR1.DescuentoGeneral "Bonificacion", IDESCUENTO, True     'INDICA UN IMPORTE, NO PORCENTAJE: TRUE = DESCUENTO, FALSE = INCREMENTO
    '
    Call MUESTRACOMANDO("Subtotal", True)
    If MUECOMHAS% < 2 Then HASAR1.Subtotal True
    '
    ' PIE DE PAGINA
    On Error Resume Next
      CPAG$ = TRIM$(FORFACPRINT.Eco21(1))
    HASAR1.Encabezado(11) = "C-Pago: " & REPLACEN$(CPAG$)
    Call FRATEXTO(FORFINFAC.Text4, 32)
    For KKII% = 1 To 3
        If KKII% <= 3 Then
          OBS$ = RETEX$(KKII%)
          HASAR1.Encabezado(KKII% + 11) = REPLACEN$(OBS$)
        End If
    Next KKII%
    On Error GoTo 0
    '
    Call MUESTRACOMANDO("Enviar", Chr$(129))
    If MUECOMHAS% < 2 Then HASAR1.Enviar Chr$(129)
    '
99  Exit Sub

impresora_apag:

    If MsgBox("Error Impresora:" & Err.Description, vbRetryCancel, "Errores") = vbRetry Then
        Resume Procesar
    End If


End Sub

Sub PRINODEBA()
  '
   Dim FS As String
   FS$ = Chr$(28)                                            '// Separador de campos del comando
   Dim COMANDO As String
   '
   On Error GoTo impresora_apag
       
Procesar:
    '
    COMANDO = Chr$(147) & FS & "1" & FS & "0000-00000000" ' comprobante original
    '
    Call MUESTRACOMANDO("Enviar", COMANDO)
    If MUECOMHAS% < 2 Then HASAR1.Enviar COMANDO
    '
    NCX% = XVALO(FORFACPRINT.Text1)
    TTXX$ = Chr$(98) & FS & STRINOM$(NCX%)
      CUITT$ = SOLONUM$(CUITCLI$(NCX%))
    TTXX$ = TTXX$ + FS + CUITT$
    TTXX$ = TTXX$ + FS + "I"    'RESPONSABLE INSCRIPTO--> VER MANUAL (PUBLTICK) PAG. 75
    TTXX$ = TTXX$ + FS + "C"    'C.U.I.T.--> VER MANUAL (PUBLTICK) PAG. 75
    TTXX$ = TTXX$ + FS + STRIDOMI$(NCX%)
    TTXX$ = Left$(TTXX$, 85)  ' SI NO DA ERROR
    '
    Call MUESTRACOMANDO("Enviar", TTXX$)
    If MUECOMHAS% < 2 Then HASAR1.Enviar TTXX$              ' IMPRIME EL ENCABEZADO DEL TICKET
    '
    COMANDO = Chr$(64) & FS & "D" & FS & "S"   ' DEFINE EL TIPO DE COMPROBANTE
    '                                           ' (LA FACTURA NO LO TIENE)
    Call MUESTRACOMANDO("Enviar", COMANDO)
    If MUECOMHAS% < 2 Then HASAR1.Enviar COMANDO
    '
    Call VERIFICAELNUMERO
    '
    FIN& = ULTREG&("!CARDEFAC")
    For U& = 1 To FIN&
        XXX$ = REGLEIDO$("!CARDEFAC", U&)
        CI% = CVI(Mid$(XXX$, 5, 2))
        DESCR$ = REPLACEN$(TRIM$(Mid$(XXX$, 7, 50)))
        CANTIDA1 = CVS(Left$(XXX$, 4))
        PREUNI = CVS(Mid$(XXX$, 429, 4))
        ALICUIVA = CVS(Mid$(XXX$, 441, 4))
        PRECONIVA = (PREUNI * (1 + ALICUIVA / 100))
        ACUMUIMPBRU = ACUMUIMPBRU + PRECONIVA * CANTIDA1
        '
        Call MUESTRACOMANDO("ImprimirItem", DESCR$, CANTIDA1, PRECONIVA, ALICUIVA, 0)
        If MUECOMHAS% < 2 Then HASAR1.ImprimirItem DESCR$, CANTIDA1, PRECONIVA, ALICUIVA, 0
       '
    Next U&
    '
    IDESCUENTO = ACUMUIMPBRU * (XVALO(FORFACPRINT.PORDESCU) / 100)
    'IDESCUENTO = XVALO(FORFACPRINT.Total(1)) * (1 + ALIVA / 100)
    '
    Call MUESTRACOMANDO("DescuentoGeneral", "Bonificacion", IDESCUENTO, True)
    If MUECOMHAS% < 2 Then HASAR1.DescuentoGeneral "Bonificacion", IDESCUENTO, True     'INDICA UN IMPORTE, NO PORCENTAJE: TRUE = DESCUENTO, FALSE = INCREMENTO
    '
    Call MUESTRACOMANDO("Subtotal", True)
    If MUECOMHAS% < 2 Then HASAR1.Subtotal True
    '
    Call MUESTRACOMANDO("CerrarComprobanteFiscal")
    If MUECOMHAS% < 2 Then HASAR1.CerrarComprobanteFiscal
    '
    Call MUESTRACOMANDO("Enviar", Chr$(129))
    If MUECOMHAS% < 2 Then HASAR1.Enviar Chr$(129)
    '
    Exit Sub

impresora_apag:

    If MsgBox("Error Impresora:" & Err.Description, vbRetryCancel, "Errores") = vbRetry Then
        Resume Procesar
    End If

End Sub

Sub PRINODEBB()
  '
   Dim COMANDA As String
   Dim FS As String
   FS$ = Chr$(28)                           '// Separador de campos del comando
   '
   On Error GoTo impresora_apag
       
Procesar:
    '
    COMANDO = Chr$(147) & FS & "1" & FS & "0000-0000000" ' comprobante original
    '
    Call MUESTRACOMANDO("Enviar", "COMANDO")
    If MUECOMHAS% < 2 Then HASAR1.Enviar COMANDO
    '
    COMANDO = Chr$(64) & FS & "E" & FS & "S"            ' DEFINE EL TIPO DE COMPROBANTE
    '
    Call MUESTRACOMANDO("Enviar", "COMANDO")
    If MUECOMHAS% < 2 Then HASAR1.Enviar COMANDO
    '
    Call VERIFICAELNUMERO
    '
    FIN& = ULTREG&("!CARDEFAC")
    For U& = 1 To FIN&
        XXX$ = REGLEIDO$("!CARDEFAC", U&)
        CI% = CVI(Mid$(XXX$, 5, 2))
        DESCR$ = REPLACEN$(TRIM$(Mid$(XXX$, 7, 50)))
        CANTIDA1 = CVS(Left$(XXX$, 4))
        PREUNI = CVS(Mid$(XXX$, 429, 4))
        ALICUIVA = CVS(Mid$(XXX$, 441, 4))
        PRECONIVA = (PREUNI * (1 + ALICUIVA / 100))
        ACUMUIMPBRU = ACUMUIMPBRU + PRECONIVA * CANTIDA1
        '
        Call MUESTRACOMANDO("ImprimirItem", DESCR$, CANTIDA1, PRECONIVA, ALICUIVA, 0)
        If MUECOMHAS% < 2 Then HASAR1.ImprimirItem DESCR$, CANTIDA1, PRECONIVA, ALICUIVA, 0
        '
    Next U&
    '
    IDESCUENTO = ACUMUIMPBRU * (XVALO(FORFACPRINT.PORDESCU) / 100)
    'IDESCUENTO = XVALO(FORFACPRINT.Total(1)) * (1 + ALIVA / 100)
    '
    Call MUESTRACOMANDO("DescuentoGeneral", "Bonificacion", IDESCUENTO, True)
    If MUECOMHAS% < 2 Then HASAR1.DescuentoGeneral "Bonificacion", IDESCUENTO, True     'INDICA UN IMPORTE, NO PORCENTAJE: TRUE = DESCUENTO, FALSE = INCREMENTO
    '
    Call MUESTRACOMANDO("Subtotal", True)
    If MUECOMHAS% < 2 Then HASAR1.Subtotal True
    '
    Call MUESTRACOMANDO("CerrarComprobanteFiscal")
    If MUECOMHAS% < 2 Then HASAR1.CerrarComprobanteFiscal
    '
    Call MUESTRACOMANDO("Enviar", Chr$(129))
    If MUECOMHAS% < 2 Then HASAR1.Enviar Chr$(129)
    '
    Exit Sub
    '
impresora_apag:

    If MsgBox("Error Impresora:" & Err.Description, vbRetryCancel, "Errores") = vbRetry Then
        Resume Procesar
    End If

End Sub

Sub MUESTRACOMANDO(COMANDO$, Optional TXT$, Optional V1, Optional V2, Optional V3, Optional V4)
    '
    TXT1$ = TXT$
    
10  For KKI% = 1 To Len(TXT1$)
      KKJ% = Asc(Mid$(TXT1$, KKI%))
      If KKJ% < 32 Or KKJ% > 128 Then
         TXT1$ = Left$(TXT1$, KKI% - 1) + "(" + TRIM$(Str$(KKJ%)) + ")" + Mid$(TXT1$, KKI% + 1)
         GoTo 10
      End If
    Next KKI%
    SCANT = Screen.MousePointer
    Screen.MousePointer = 1
    If MUECOMHAS% > 0 Then
       With VERICONFIS
         .Label2 = COMANDO$
         .Label7 = TXT1$
         If XVALO(V1) <> 448 Then .Label8 = CStr(V1) Else .Label8 = "Vacío"
         If CDbl(V2) <> 448 Then .Label9 = CStr(V2) Else .Label9 = "Vacío"
         If CDbl(V3) <> 448 Then .Label10 = CStr(V3) Else .Label10 = "Vacío"
         If CDbl(V4) <> 448 Then .Label11 = CStr(V4) Else .Label11 = "Vacío"
         .Show 1
         Call FRESHALL
       End With
    End If
    Screen.MousePointer = SCA
    '
End Sub

Sub VERIFICAELNUMERO()
    '
    Exit Sub ' SE INUTILIZA ESTA SUB POR QUE CON EL USO DE LA RUTINA NUFACTURA YA TOMA EL NUMERO DE LA IMPRESORA FISCAL
    '
    On Error Resume Next
    PRONUMERO$ = HASAR1.Respuesta(3)
    If Err Then
       On Error GoTo 0
       ' CALL MENSERR(24,"Imposible Determinar el Número\del Próximo Comprobante Fiscal.\   \Cierre la Aplicación y Re-Intente.")
       If COMALTER%("Imposible Determinar el Número\del Próximo Comprobante Fiscal.\\Abortar\Continuar") <> 2 Then
          On Error Resume Next
          HASAR1.TratarDeCancelarTodo
          On Error GoTo 0
          Call FINAL("")
       End If
    End If
    On Error GoTo 0
    '
    If XVALO(PRONUMERO$) <> XVALO(FORFACPRINT.Text11) Then
       'Call MENSERR(24, "El número del Próximo Comprobante Fiscal '" + PRONUMERO$ + "'\Difiere del Seteado en la Base de Datos FLEXOFT.\   \Debe Aparear los Números para Continuar.\   \Cierre la Aplicación y Re-Intente.")
       If COMALTER%("El Número del Próximo Comprobante Fiscal '" + PRONUMERO$ + "'\Difiere del Seteado en la Base de Datos FLEXOFT.\\Abortar\Continuar") <> 2 Then
          On Error Resume Next
          HASAR1.TratarDeCancelarTodo
          On Error GoTo 0
          Call FINAL("")
       End If
    End If
    '
End Sub
'
Sub PRITIREMI()
    '
    Dim COMANDO As String
    '
    On Error GoTo impresora_apag

Procesar:
    '
    On Error Resume Next
    NCX% = XVALO(FORFACPRINT.Text1)
    CUITT$ = SOLONUM$(CUITCLI$(NCX%))
    On Error GoTo 0
    TIDOCUX = TIPO_CUIT
    '
    PVCL% = PIVACLI%(NCX%)
    RESPONIVA = RESPONSABLE_INSCRIPTO
    If PVCL% < 1 Then RESPONIVA = CONSUMIDOR_FINAL: TIDOCUX = TIPO_DNI
    If PVCL% = 2 Then RESPONIVA = RESPONSABLE_NO_INSCRIPTO
    If PVCL% = 3 Then RESPONIVA = NO_RESPONSABLE: TIDOCUX = TIPO_DNI
    If PVCL% = 4 Then RESPONIVA = RESPONSABLE_EXENTO
    If PVCL% = 5 Then RESPONIVA = NO_CATEGORIZADO: TIDOCUX = TIPO_DNI
    If PVCL% = 6 Then RESPONIVA = MONOTRIBUTO
    '
    ' ENCABEZADO
    On Error Resume Next
    HASAR1.Encabezado(1) = "VIAMONTE 2255"
    HASAR1.Encabezado(2) = "C1056ABI - CAPITAL FEDERAL"
      TELE$ = "Tel: 5411-4953-6114/6111"
    HASAR1.Encabezado(3&) = TELE$
      CTE$ = TRIM$(FORFACPRINT.Text1) & " - " & TRIM$(FORFACPRINT.Label26)
    HASAR1.Encabezado(4&) = "Cliente: " & CTE$
    HASAR1.Encabezado(5&) = ""
    On Error GoTo 0
    '
    ' *** DATOS DEL CLIENTE ****
    Call MUESTRACOMANDO("DatosCliente", STRINOM$(NCX%), CUITT$, TIDOCUX, RESPONIVA)
    CODPROVI$ = PROVCLI$(NCX%): PROVINCIA$ = REPLACEN$(TRIM$(ECOARCH$("PROVINC", CODPROVI$)))
    If MUECOMHAS% < 2 Then HASAR1.DatosCliente STRINOM$(NCX%), CUITT$, TIDOCUX, RESPONIVA, STRIDOMI$(NCX%)
    '
    ' *** ABRE COMPROBANTE NO HOMOLOGADO ***
    Call MUESTRACOMANDO("AbrirComprobanteNoFiscalHomologado", REMITO)
    If MUECOMHAS% < 2 Then HASAR1.AbrirComprobanteNoFiscalHomologado REMITO
    '
    Call VERIFICANUMEROREMITO
    '
    FIN& = ULTREG&("!CARDEFAC")
    For U& = 1 To FIN&
        XXX$ = REGLEIDO$("!CARDEFAC", U&)
        CI% = CVI(Mid$(XXX$, 5, 2))
        CODIGO$ = REPLACEN$(TRIM$(CODEXT$(CI%)))
        DESCR$ = REPLACEN$(TRIM$(Mid$(XXX$, 7, 64)))
        CANTIDA1# = ROUND(CVS(Left$(XXX$, 4)), 1)
        NUMSER$ = TRIM$(Mid$(XXX$, 139, 64))
        '
        Call MUESTRACOMANDO("ImprimirItemEnRemito", DESCR$, CANTIDA1#, 0, 0, 0)
        If MUECOMHAS% < 2 Then
          HASAR1.ImprimirItemEnRemito STRICODES$(CI%, DESCR$), CANTIDA1#
          If NUMSER$ <> "" Then
            Call MUESTRACOMANDO("ImprimirItemEnRemito", NUMSER$, 0, 0, 0, 0)
            HASAR1.ImprimirItemEnRemito "S/N: " & NUMSER$, 0
          End If
        End If
        '
    Next U&
    '
    Call MUESTRACOMANDO("CerrarComprobanteNoFiscalHomologado")
    If MUECOMHAS% < 2 Then HASAR1.CerrarComprobanteNoFiscalHomologado
    Call MUESTRACOMANDO("ReimprimirComprobante")
    If MUECOMHAS% < 2 Then HASAR1.ReimprimirComprobante  ' EMITE SEGUNDA COPIA
    '
    Exit Sub

impresora_apag:

    If MsgBox("Error Impresora:" & Err.Description, vbRetryCancel, "Errores") = vbRetry Then
        Resume Procesar
    End If
End Sub

Sub VERIFICANUMEROREMITO()
    '
    Exit Sub ' SE INUTILIZA ESTA SUB POR QUE CON EL USO DE LA RUTINA NUFACTURA YA TOMA EL NUMERO DE LA IMPRESORA FISCAL
    '
    On Error Resume Next
    PRONUMERO$ = HASAR1.Respuesta(3)
    If Err Then
       On Error GoTo 0
       ' CALL MENSERR(24,"Imposible Determinar el Número\del Próximo Comprobante Fiscal.\   \Cierre la Aplicación y Re-Intente.")
       If COMALTER%("Imposible Determinar el Número\del Próximo Comprobante Fiscal.\\Abortar\Continuar") <> 2 Then
          On Error Resume Next
          HASAR1.TratarDeCancelarTodo
          On Error GoTo 0
          Call FINAL("")
       End If
    End If
    On Error GoTo 0
    '
    If XVALO(PRONUMERO$) <> XVALO(FORFACPRINT.Text10) Then
       'Call MENSERR(24, "El número del Próximo Comprobante Fiscal '" + PRONUMERO$ + "'\Difiere del Seteado en la Base de Datos FLEXOFT.\   \Debe Aparear los Números para Continuar.\   \Cierre la Aplicación y Re-Intente.")
       If COMALTER%("El Número del Próximo Comprobante Fiscal '" + PRONUMERO$ + "'\Difiere del Seteado en la Base de Datos FLEXOFT.\\Abortar\Continuar") <> 2 Then
          On Error Resume Next
          HASAR1.TratarDeCancelarTodo
          On Error GoTo 0
          Call FINAL("")
       End If
    End If
    '
End Sub

Sub TOMANUMDELCONFIS_MAL(TDC$, LTDC$, PUNTODEVENTA$, NUMECOMP$)
    '
    Dim FS As String
    FS$ = Chr$(28)                                            '// Separador de campos del comando
    '
    Call INICIA
    On Error GoTo impresora_apag
    Call HASAR1.ObtenerDatosDeInicializacion
    YY$ = HASAR1.Respuesta(0)
    PUNTODEVENTA$ = TRIM$(Mid$(YY$, 92, 4))
    '
    Select Case TDC$   ' TIPO DE COMPROBANTE
        Case "FC"
            If LTDC$ = "A" Then                  ' FACTURA 'A'
                     '
                     TTXX$ = Chr$(98) & FS & "Cliente: NOMBREDELCLIENTE"
                     TTXX$ = TTXX$ + FS + "23109234559"         ' PARA QUE SEA UN NUMERO VALIDO
                     TTXX$ = TTXX$ + FS + "I"                   ' RESPONSABLE INSCRIPTO--> VER MANUAL (PUBLTICK) PAG. 75
                     TTXX$ = TTXX$ + FS + "C"                   ' C.U.I.T.--> VER MANUAL (PUBLTICK) PAG. 75
                     TTXX$ = TTXX$ + FS + "Domicilio: CALLE, CP NNNN"
                     TTXX$ = Left$(TTXX$, 85)                   ' SI NO DA ERROR
                  HASAR1.Enviar TTXX$                           ' IMPRIME EL ENCABEZADO DEL TICKET
                     '
                  HASAR1.AbrirComprobanteFiscal FACTURA_A
                     '
                  NUMECOMP$ = TRIM$(HASAR1.Respuesta(3))
                  HASAR1.TratarDeCancelarTodo
                     '
               ElseIf LTDC$ = "B" Then            ' FACTURA 'B'
                     '
                  HASAR1.DatosCliente "Cliente: NOMBREDELCLIENTE", "23109234559", TIPO_CUIT, MONOTRIBUTO, "Domicilio: DOMICILIO"
                     '
                  HASAR1.AbrirComprobanteFiscal FACTURA_B
                     '
                  NUMECOMP$ = TRIM$(HASAR1.Respuesta(3))
                  HASAR1.TratarDeCancelarTodo
                  '
               Else
                  GoTo impresora_apag
            End If
            '
        Case "ND"
            If LTDC$ = "A" Then                    ' NOTA DE DEBITO 'A'
                     '
                     COMANDO$ = Chr$(147) & FS & "1" & FS & "0000-00000000" ' comprobante original
                  HASAR1.Enviar COMANDO$
                     '
                     TTXX$ = Chr$(98) & FS & "Cliente: NOMBREDELCLIENTE"
                     TTXX$ = TTXX$ + FS + "23109234559"         ' PARA QUE SEA UN NUMERO VALIDO
                     TTXX$ = TTXX$ + FS + "I"                   ' RESPONSABLE INSCRIPTO--> VER MANUAL (PUBLTICK) PAG. 75
                     TTXX$ = TTXX$ + FS + "C"                   ' C.U.I.T.--> VER MANUAL (PUBLTICK) PAG. 75
                     TTXX$ = TTXX$ + FS + "Domicilio: CALLE, CP NNNN"
                     TTXX$ = Left$(TTXX$, 85)                   ' SI NO DA ERROR
                  HASAR1.Enviar TTXX$                           ' IMPRIME EL ENCABEZADO DEL TICKET
                     '
                     COMANDO$ = Chr$(64) & FS & "D" & FS & "S"  ' DEFINE EL TIPO DE COMPROBANTE
                  HASAR1.Enviar COMANDO$
                     '
                  NUMECOMP$ = TRIM$(HASAR1.Respuesta(3))
                  HASAR1.TratarDeCancelarTodo
                     '
                ElseIf LTDC$ = "B" Then            ' NOTA DE DEBITO 'B'
                     '
                     COMANDO$ = Chr$(147) & FS & "1" & FS & "0000-0000000" ' comprobante original
                  HASAR1.Enviar COMANDO
                     '
                     COMANDO$ = Chr$(64) & FS & "E" & FS & "S"            ' DEFINE EL TIPO DE COMPROBANTE
                  HASAR1.Enviar COMANDO
                     '
                  NUMECOMP$ = TRIM$(HASAR1.Respuesta(3))
                  HASAR1.TratarDeCancelarTodo
                  '
               Else
                  GoTo impresora_apag
            End If
            '
        Case "NC"
            If LTDC$ = "A" Then                    ' NOTA DE CREDITO 'A'
                     '
                     COMANDO = Chr$(147) & FS & "1" & FS & "0000-0000000" ' comprobante original
                  HASAR1.Enviar COMANDO
                     '
                     TTXX$ = Chr$(98) & FS & "Cliente: NOMBREDELCLIENTE"
                     TTXX$ = TTXX$ + FS + "23109234559"         ' PARA QUE SEA UN NUMERO VALIDO
                     TTXX$ = TTXX$ + FS + "I"                   ' RESPONSABLE INSCRIPTO--> VER MANUAL (PUBLTICK) PAG. 75
                     TTXX$ = TTXX$ + FS + "C"                   ' C.U.I.T.--> VER MANUAL (PUBLTICK) PAG. 75
                     TTXX$ = TTXX$ + FS + "Domicilio: CALLE, CP NNNN"
                     TTXX$ = Left$(TTXX$, 85)                   ' SI NO DA ERROR
                  HASAR1.Enviar TTXX$                           ' IMPRIME EL ENCABEZADO DEL TICKET
                     '
                     COMANDO = Chr$(128) & FS & "R" & FS & "T"   ' DEFINE EL TIPO DE COMPROBANTE
                  HASAR1.Enviar COMANDO
                     '
                  NUMECOMP$ = TRIM$(HASAR1.Respuesta(3))
                  HASAR1.TratarDeCancelarTodo
                     '
                ElseIf LTDC$ = "B" Then            ' NOTA DE CREDITO 'B'
                     '
                     COMANDO = Chr$(147) & FS & "1" & FS & "0000-0000000" ' comprobante original
                  HASAR1.Enviar COMANDO
                     '
                     TTXX$ = Chr$(98) & FS & "Cliente: NOMBREDELCLIENTE"
                     TTXX$ = TTXX$ + FS + "23109234559"         ' PARA QUE SEA UN NUMERO VALIDO
                     TTXX$ = TTXX$ + FS + "C"                   ' C.U.I.T.--> VER MANUAL (PUBLTICK) PAG. 75
                     TTXX$ = TTXX$ + FS + "2"
                     TTXX$ = TTXX$ + FS + "Domicilio: CALLE, CP NNNN"
                     TTXX$ = Left$(TTXX$, 85)                   ' SI NO DA ERROR
                  HASAR1.Enviar TTXX$              ' IMPRIME EL ENCABEZADO DEL TICKET
                     '
                     COMANDO = Chr$(128) & FS & "S" & FS & "T"       ' DEFINE EL TIPO DE COMPROBANTE
                  HASAR1.Enviar COMANDO
                  '
                Else
                  GoTo impresora_apag
            End If
            '
        Case Else
           Call MENSERR(24, "Comprobante Tipo '" + TDC$ + "' no Procesable - Consulte")
           Call FINAL("")
    End Select
    '
    ' AQUI CAE Y TOMA EL NUMERO DEL PROXIMO COMPROBANTE
    NUMECOMP$ = TRIM$(HASAR1.Respuesta(3))
    HASAR1.TratarDeCancelarTodo
    On Error GoTo 0
    Exit Sub
    '
impresora_apag:
    '
    Call MENSERR(24, "Error Impresora Fiscal: " & Err.Description & "\  \Imposible Obtener Datos de Punto de Venta\y Número del Próximo Comprobante\de la Impresora Fiscal.\   \Cierre y Re-Ingrese al Módulo de Facturación.")
    Call CIERRARCH("*.*")
    Call FINAL("")
    '
End Sub

Sub TOMANUMDELCONFIS(TDC$, LTDC$, PUNTODEVENTA$, NUMECOMP$)
    '
    Dim FS As String
    FS$ = Chr$(28)                                            '// Separador de campos del comando
    '
    If LTDC$ = "E" Then                                       ' SI ES EXPORTACION SALE
       If TDC$ <> "RT" Then
          Exit Sub
       End If
    End If
    '
    On Error GoTo impresora_apag
    Call HASAR1.ObtenerDatosDeInicializacion
    YY$ = HASAR1.Respuesta(0)
    PUNTODEVENTA$ = TRIM$(Mid$(YY$, 92, 4))
    '
    Select Case TDC$   ' TIPO DE COMPROBANTE
        Case "FC"
            If LTDC$ = "A" Then                    ' FACTURA 'A'
                  PARAM% = 5
               ElseIf LTDC$ = "B" Then             ' FACTURA 'B'
                  PARAM% = 3
               Else
                  GoTo impresora_apag
            End If
            '
        Case "ND"
            If LTDC$ = "A" Then                    ' NOTA DE DEBITO 'A'
                   PARAM% = 5
               ElseIf LTDC$ = "B" Then             ' FACTURA 'B'
                   PARAM% = 3
               Else
                  GoTo impresora_apag
            End If
        Case "NC"
            If LTDC$ = "A" Then                    ' NOTA DE CREDITO 'A'
                   PARAM% = 8
               ElseIf LTDC$ = "B" Then             ' FACTURA 'B'
                   PARAM% = 7
               Else
                  GoTo impresora_apag
            End If
            '
        Case "RT"                                   ' REMITO
            PARAM% = 9
            '
       Case Else
           Call MENSERR(24, "Comprobante Tipo '" + TDC$ + "' no Procesable - Consulte")
           Call FINAL("")
    End Select
    '
    ' AQUI CAE Y TOMA EL NUMERO DEL PROXIMO COMPROBANTE
    HASAR1.PedidoDeStatus
    NUMECOMP$ = TRIM$(FORMATNUM$(1 + XVALO(TRIM$(HASAR1.Respuesta(PARAM%))), "F8.0"))
    While Len(NUMECOMP$) < 8: NUMECOMP$ = "0" + NUMECOMP$: Wend
    HASAR1.TratarDeCancelarTodo
    '
    On Error GoTo 0
    Exit Sub
    '
impresora_apag:
    '
    Call MENSERR(24, "Error Impresora Fiscal: " & Err.Description & "\  \Imposible Obtener Datos de Punto de Venta\y Número del Próximo Comprobante\de la Impresora Fiscal.\   \Cierre y Re-Ingrese al Módulo de Facturación.")
    If MUECOMHAS% > 0 Then
       PUNTODEVENTA$ = "9999"
       NUMECOMP$ = "00000001"
       Exit Sub
    End If
    Call CIERRARCH("*.*")
    Call FINAL("")
    '
End Sub

Function STRIDOMI$(NC%)
    '
    DATODOMI$ = ""
    
    DOMICLIE$ = REPLACEN$(DOMICLI$(NC%))
    COPOS$ = REPLACEN$(CPOSCLI$(NC%))
    LOCALL$ = REPLACEN$(LOCACLI$(NC))
    '
    DATODOMI$ = "Domicilio: " + DOMICLIE$ + " CP " + COPOS$ + " " + LOCALL$
    '
    STRIDOMI$ = Left$(DATODOMI$, 50)
    '
End Function

Function STRINOM$(NC%)
    '
    DATO$ = ""
    '
    RASCLIE$ = REPLACEN$(RASOCLI(NC%))
    NN = Len(TRIM(FORMATNUM(NC%, "I5.0")))
    DATO$ = "Cliente: " + Mid$(RASCLIE$, 1, (39 - NN)) + "(" & NC% & ")"
    STRINOM$ = DATO$
    '
End Function

Function STRICODES$(CI%, DESCR0$)
    '
    DATOCODI$ = ""
    CODI$ = TRIM$(REPLACEN$(CODEXT$(CI%)))
    If CODI$ <> "" Then
         DESCRIPP$ = REPLACEN$(DESCRIT$(CI%))
       Else
         DESCRIPP$ = DESCR0$
    End If
    DATOCODI$ = Left$(TRIM$(CODI$ + " " + DESCRIPP$), 50)
    '
    STRICODES$ = DATOCODI$
    '
End Function

