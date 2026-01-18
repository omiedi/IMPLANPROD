VERSION 5.00
Object = "{D9AF33E0-7C55-11D5-9151-0000E856BC17}#1.0#0"; "fiscal010724.ocx"
Begin VB.Form FORHASAR 
   Caption         =   "CONTROLADOR HASAR"
   ClientHeight    =   2415
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3600
   LinkTopic       =   "Form1"
   ScaleHeight     =   2415
   ScaleWidth      =   3600
   StartUpPosition =   3  'Windows Default
   Begin FiscalPrinterLibCtl.HASAR HASAR1 
      Left            =   1575
      OleObjectBlob   =   "FORHASAR.frx":0000
      Top             =   945
   End
End
Attribute VB_Name = "FORHASAR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Sub REPORTES()
        '
        Call MUESTRACOMANDO("Puerto", 1)
        If MUECOMHAS% < 2 Then HASAR1.Puerto = 1
        '
        Call MUESTRACOMANDO("Modelo", MODELO_615)
        If MUECOMHAS% < 2 Then HASAR1.Modelo = MODELO_615          '//Este OCX no es 100% para 715F
        
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
    Call MUESTRACOMANDO("Modelo", MODELO_615)
    If MUECOMHAS% < 2 Then HASAR1.Modelo = MODELO_615              ' //Este OCX no es 100% para 715F
    '
    On Error Resume Next
    '
    Call MUESTRACOMANDO("Comenzar")
    If MUECOMHAS% < 2 Then HASAR1.Comenzar
    '
    If Err Then
      On Error GoTo 0
      TTXX$ = "Imposible Conectar con el Controlador Fiscal.\Revise el Encendido y las Conexiones\o Re-Configure para Impresora Común.\  \En Caso de Continuar, los Comandos se Presentarán\en Pantalla, pero NO se Enviarán al Controlador."
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
      Call MENSERR(24, "Esta Controladora fiscal requiere para los comprobantes del tipo 'A' que el CUIT\del contribuyente sea válido. El cliente actual posee un CUIT inválido,\por lo que el comprobante ha quedado registrado en su cuenta corriente\pero el mismo no se podrá imprimir. ")
    End If

    TTXX$ = Chr$(98) & FS & "Cliente: " + TRIM$(FORFACPRINT.Text2)
        NCX% = XVALO(FORFACPRINT.Text1)
        CUITT$ = REPLACAR$(CUITCLI$(NCX%), "-", "")
    TTXX$ = TTXX$ + FS + CUITT$
    TTXX$ = TTXX$ + FS + "I"    'RESPONSABLE INSCRIPTO--> VER MANUAL (PUBLTICK) PAG. 75
    TTXX$ = TTXX$ + FS + "C"    'C.U.I.T.--> VER MANUAL (PUBLTICK) PAG. 75
    TTXX$ = TTXX$ + FS + "Domicilio: " + TRIM$(FORFACPRINT.Text3) + "," + "CP " + TRIM$(FORFACPRINT.Text4)
    TTXX$ = Left$(TTXX$, 85)  ' SI NO DA ERROR
    Call MUESTRACOMANDO("Enviar", TTXX$)
    If MUECOMHAS% < 2 Then HASAR1.Enviar TTXX$         ' IMPRIME EL ENCABEZADO DEL TICKET
            ' ATENCION: Unicamente en la Factura A se utiliza el método ENVIAR, porque
            ' es el que permite incluir el domicilio del cliente. En las "B" se usa
            ' método DATOSCLIENTE
    '
    Call MUESTRACOMANDO("AbrirComprobanteFiscal", TICKET_FACTURA_A)
    If MUECOMHAS% < 2 Then HASAR1.AbrirComprobanteFiscal TICKET_FACTURA_A
    '
    FIN& = ULTREG&("!CARDEFAC")
    For U& = 1 To FIN&
        XXX$ = REGLEIDO$("!CARDEFAC", U&)
        CI% = CVI(Mid$(XXX$, 5, 2))
        DESCR$ = TRIM$(Mid$(XXX$, 7, 64))
        CANTIDA1# = CVS(Left$(XXX$, 4))
        PREUNI# = CVS(Mid$(XXX$, 429, 4))
        ALICUIVA# = CVS(Mid$(XXX$, 441, 4))
        PRECONIVA# = (PREUNI * (1 + ALICUIVA# / 100))
        Call MUESTRACOMANDO("ImprimirItem", DESCR$, CANTIDA1#, PRECONIVA#, ALICUIVA#, 0)
        If MUECOMHAS% < 2 Then HASAR1.ImprimirItem DESCR$, CANTIDA1#, PRECONIVA#, ALICUIVA#, 0
    Next U&
    '
    IDESCUENTO = XVALO(FORFACPRINT.Total(1)) * (1 + ALIVA / 100)
    Call MUESTRACOMANDO("DescuentoGeneral", "Bonificacion", IDESCUENTO, True)
    If MUECOMHAS% < 2 Then HASAR1.DescuentoGeneral "Bonificacion", IDESCUENTO, True           'INDICA UN IMPORTE, NO PORCENTAJE: TRUE = DESCUENTO, FALSE = INCREMENTO
    '
    Call MUESTRACOMANDO("Subtotal", True)
    If MUECOMHAS% < 2 Then HASAR1.Subtotal True
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
    Dim COMANDO As String
    Dim FS As String
    FS = Chr$(28)                                            '// Separador de campos del comando
    '
    On Error GoTo impresora_apag
           
Procesar:
    '
    NCX% = XVALO(FORFACPRINT.Text1)
    CUITT$ = REPLACAR$(CUITCLI$(NCX%), "-", "")
    '
    Call MUESTRACOMANDO("AbrirComprobanteFiscal", TICKET_FACTURA_B)
    If MUECOMHAS% < 2 Then HASAR1.AbrirComprobanteFiscal TICKET_FACTURA_B
    '
    Call MUESTRACOMANDO("DatosCliente", "Cliente: " + TRIM$(FORFACPRINT.Text2), CUITT$, TIPO_NINGUNO, CONSUMIDOR_FINAL)
    If MUECOMHAS% < 2 Then HASAR1.DatosCliente "Cliente: " + TRIM$(FORFACPRINT.Text2), CUITT$, TIPO_NINGUNO, CONSUMIDOR_FINAL
    'IGUALMENTE SOLO IMPRIME LA LEYENDA CONSUMIDOR FINAL
    '
    FIN& = ULTREG&("!CARDEFAC")
    For U& = 1 To FIN&
        XXX$ = REGLEIDO$("!CARDEFAC", U&)
        CI% = CVI(Mid$(XXX$, 5, 2))
        DESCR$ = TRIM$(Mid$(XXX$, 7, 64))
        CANTIDA1# = CVS(Left$(XXX$, 4))
        PREUNI# = CVS(Mid$(XXX$, 429, 4))
        ALICUIVA# = CVS(Mid$(XXX$, 441, 4))
        PRECONIVA# = (PREUNI# * (1 + ALICUIVA# / 100))
        '
        Call MUESTRACOMANDO("ImprimirItem", DESCR$, CANTIDA1#, PRECONIVA#, ALICUIVA#, 0)
        If MUECOMHAS% < 2 Then HASAR1.ImprimirItem DESCR$, CANTIDA1#, PRECONIVA#, ALICUIVA#, 0
        '
    Next U&
    '
    IDESCUENTO = XVALO(FORFACPRINT.Total(1)) * (1 + ALIVA / 100)
    '
    Call MUESTRACOMANDO("DescuentoGeneral", "Bonificacion", IDESCUENTO, True)
    If MUECOMHAS% < 2 Then HASAR1.DescuentoGeneral "Bonificacion", IDESCUENTO, True        'INDICA UN IMPORTE, NO PORCENTAJE: TRUE = DESCUENTO, FALSE = INCREMENTO
    '
    Call MUESTRACOMANDO("SubTotal", True)
    If MUECOMHAS% < 2 Then HASAR1.Subtotal True
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
    '
    NCLIE% = XVALO(FORFACPRINT.Text1) ' VALIDACION DE QUE TENGA CUIT EL CLIENTE SI NO MUESTRA MENSAJE
    CUITT$ = CUITCLI$(NCLIE%)
    If VALICUIT(CUITT$) < 1 Then
      Call MENSERR(24, "Esta controladora fiscal requiere para los comprobantes del tipo 'A' que el CUIT\del contribuyente sea válido. El cliente actual posee un CUIT inválido,\por lo que el comprobante ha quedado registrado en su cuenta corriente\pero el mismo no se podrá imprimir. ")
    End If

    COMANDO = Chr$(147) & FS & "1" & FS & "0000-0000000" ' comprobante original
    '
    Call MUESTRACOMANDO("Enviar", COMANDO)
    If MUECOMHAS% < 2 Then HASAR1.Enviar COMANDO
    '
    TTXX$ = Chr$(98) & FS & "Cliente: " + TRIM$(FORFACPRINT.Text2)
    NCX% = XVALO(FORFACPRINT.Text1)
    CUITT$ = REPLACAR$(CUITCLI$(NCX%), "-", "")
    TTXX$ = TTXX$ + FS + CUITT$
    TTXX$ = TTXX$ + FS + "I"    'RESPONSABLE INSCRIPTO--> VER MANUAL (PUBLTICK) PAG. 75
    TTXX$ = TTXX$ + FS + "C"    'C.U.I.T.--> VER MANUAL (PUBLTICK) PAG. 75
    TTXX$ = TTXX$ + FS + "Domicilio: " + TRIM$(FORFACPRINT.Text3) + "," + "CP " + TRIM$(FORFACPRINT.Text4)
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
    FIN& = ULTREG&("!CARDEFAC")
    For U& = 1 To FIN&
        XXX$ = REGLEIDO$("!CARDEFAC", U&)
        CI% = CVI(Mid$(XXX$, 5, 2))
        DESCR$ = TRIM$(Mid$(XXX$, 7, 64))
        CANTIDA1 = CVS(Left$(XXX$, 4))
        PREUNI = CVS(Mid$(XXX$, 429, 4))
        ALICUIVA = CVS(Mid$(XXX$, 441, 4))
        PRECONIVA = (PREUNI * (1 + ALICUIVA / 100))
        '
        Call MUESTRACOMANDO("ImprimirItem", DESCR$, CANTIDA1, PRECONIVA, ALICUIVA, 0)
        If MUECOMHAS% < 2 Then HASAR1.ImprimirItem DESCR$, CANTIDA1, PRECONIVA, ALICUIVA, 0
        '
    Next U&
    '
    IDESCUENTO = XVALO(FORFACPRINT.Total(1)) * (1 + ALIVA / 100)
    '
    Call MUESTRACOMANDO("DescuentoGeneral", "Bonificacion", IDESCUENTO, True)
    If MUECOMHAS% < 2 Then HASAR1.DescuentoGeneral "Bonificacion", IDESCUENTO, True                'INDICA UN IMPORTE, NO PORCENTAJE: TRUE = DESCUENTO, FALSE = INCREMENTO
    '
    Call MUESTRACOMANDO("Subtotal", True)
    If MUECOMHAS% < 2 Then HASAR1.Subtotal True
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
    '
    NCLIE% = XVALO(FORFACPRINT.Text1) ' VALIDACION DE QUE TENGA CUIT EL CLIENTE SI NO MUESTRA MENSAJE
    CUITT$ = CUITCLI$(NCLIE%)
    If VALICUIT(CUITT$) < 1 Then
      Call MENSERR(24, "Esta controladora fiscal requiere para los comprobantes del tipo 'B' que el CUIT\del contribuyente sea válido. El cliente actual posee un CUIT inválido,\por lo que el comprobante ha quedado registrado en su cuenta corriente\pero el mismo no se podrá imprimir. ")
    End If

    COMANDO = Chr$(147) & FS & "1" & FS & "0000-0000000" ' comprobante original
    '
    Call MUESTRACOMANDO("Enviar", COMANDO)
    If MUECOMHAS% < 2 Then HASAR1.Enviar COMANDO
    '
    TTXX$ = Chr$(98) & FS & "Cliente: " + TRIM$(FORFACPRINT.Text2)
    NCX% = XVALO(FORFACPRINT.Text1)
    CUITT$ = REPLACAR$(CUITCLI$(NCX%), "-", "")
    TTXX$ = TTXX$ + FS + CUITT$
    TTXX$ = TTXX$ + FS + "C"            'C.U.I.T.--> VER MANUAL (PUBLTICK) PAG. 75
    TTXX$ = TTXX$ + FS + "2"
    TTXX$ = TTXX$ + FS + "Domicilio: " + TRIM$(FORFACPRINT.Text3) + "," + "CP " + TRIM$(FORFACPRINT.Text4)
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
    FIN& = ULTREG&("!CARDEFAC")
    For U& = 1 To FIN&
        XXX$ = REGLEIDO$("!CARDEFAC", U&)
        CI% = CVI(Mid$(XXX$, 5, 2))
        DESCR$ = TRIM$(Mid$(XXX$, 7, 64))
        CANTIDA1 = CVS(Left$(XXX$, 4))
        PREUNI = CVS(Mid$(XXX$, 429, 4))
        ALICUIVA = CVS(Mid$(XXX$, 441, 4))
        PRECONIVA = (PREUNI * (1 + ALICUIVA / 100))
        '
        Call MUESTRACOMANDO("ImprimirItem", DESCR$, CANTIDA1, PRECONIVA, ALICUIVA, 0)
        If MUECOMHAS% < 2 Then HASAR1.ImprimirItem DESCR$, CANTIDA1, PRECONIVA, ALICUIVA, 0
        '
    Next U&
    '
    IDESCUENTO = XVALO(FORFACPRINT.Total(1)) * (1 + ALIVA / 100)
    '
    Call MUESTRACOMANDO("DescuentoGeneral", "Bonificacion", IDESCUENTO, True)
    If MUECOMHAS% < 2 Then HASAR1.DescuentoGeneral "Bonificacion", IDESCUENTO, True     'INDICA UN IMPORTE, NO PORCENTAJE: TRUE = DESCUENTO, FALSE = INCREMENTO
    '
    Call MUESTRACOMANDO("Subtotal", True)
    If MUECOMHAS% < 2 Then HASAR1.Subtotal True
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
    TTXX$ = Chr$(98) & FS & "Cliente: " + TRIM$(FORFACPRINT.Text2)
      NCX% = XVALO(FORFACPRINT.Text1)
      CUITT$ = REPLACAR$(CUITCLI$(NCX%), "-", "")
    TTXX$ = TTXX$ + FS + CUITT$
    TTXX$ = TTXX$ + FS + "I"    'RESPONSABLE INSCRIPTO--> VER MANUAL (PUBLTICK) PAG. 75
    TTXX$ = TTXX$ + FS + "C"    'C.U.I.T.--> VER MANUAL (PUBLTICK) PAG. 75
    TTXX$ = TTXX$ + FS + "Domicilio: " + TRIM$(FORFACPRINT.Text3) + "," + "CP " + TRIM$(FORFACPRINT.Text4)
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
    FIN& = ULTREG&("!CARDEFAC")
    For U& = 1 To FIN&
        XXX$ = REGLEIDO$("!CARDEFAC", U&)
        CI% = CVI(Mid$(XXX$, 5, 2))
        DESCR$ = TRIM$(Mid$(XXX$, 7, 64))
        CANTIDA1 = CVS(Left$(XXX$, 4))
        PREUNI = CVS(Mid$(XXX$, 429, 4))
        ALICUIVA = CVS(Mid$(XXX$, 441, 4))
        PRECONIVA = (PREUNI * (1 + ALICUIVA / 100))
        '
        Call MUESTRACOMANDO("ImprimirItem", DESCR$, CANTIDA1, PRECONIVA, ALICUIVA, 0)
        If MUECOMHAS% < 2 Then HASAR1.ImprimirItem DESCR$, CANTIDA1, PRECONIVA, ALICUIVA, 0
       '
    Next U&
    '
    IDESCUENTO = XVALO(FORFACPRINT.Total(1)) * (1 + ALIVA / 100)
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
    FIN& = ULTREG&("!CARDEFAC")
    For U& = 1 To FIN&
        XXX$ = REGLEIDO$("!CARDEFAC", U&)
        CI% = CVI(Mid$(XXX$, 5, 2))
        DESCR$ = TRIM$(Mid$(XXX$, 7, 64))
        CANTIDA1 = CVS(Left$(XXX$, 4))
        PREUNI = CVS(Mid$(XXX$, 429, 4))
        ALICUIVA = CVS(Mid$(XXX$, 441, 4))
        PRECONIVA = (PREUNI * (1 + ALICUIVA / 100))
        '
        Call MUESTRACOMANDO("ImprimirItem", DESCR$, CANTIDA1, PRECONIVA, ALICUIVA, 0)
        If MUECOMHAS% < 2 Then HASAR1.ImprimirItem DESCR$, CANTIDA1, PRECONIVA, ALICUIVA, 0
        '
    Next U&
    '
    IDESCUENTO = XVALO(FORFACPRINT.Total(1)) * (1 + ALIVA / 100)
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
         .Label7 = TXT1$ 'se agrego un xvalo a la siguiente linea por que en los casos que el cliente no tiene cuit viene vacio y da un error al querer convertir una cadena vacia en la funcion CDBL
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

