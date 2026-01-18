Attribute VB_Name = "FORM_MANAGER"
Public REFRESHFORM%    ' EJECUTO BAJALDISCO - RE-ABRIR FORMULARIO
Public REFRESHFORMDOS% ' EJECUTO BAJALDISCO - RE-ABRIR FORMULARIO
Public SETUPSTRI$
Public PRINTERPORT$, PORTUFAC$, PORTUREM$
Dim FORMACT$
Public CODFOR$
Dim COEFH, COEFV
Dim CARE%, ATOT%
Public MARSUP%
Public MARIZW, MARSUW
Public DETDESMM, DETHASMM
Public INTERLI, ORIENTA%, RENULIS%
Public CACOPI%
'
Public VECPRICO$(), CAREVEC&
'
Public YAC
'
Dim CODVAR$(128), REGFOR%(128)
Dim CAFUEN As Integer
Dim Fuente$(1024)
Public ABOPRIN%
Public CANCELPRINT%
Public NOCONFIRM%
Public HOJASUEL%
Public MATRIXPUN%
Public FUERADEAREA%
Public FUENODISPO%
'
Public NUMEDOC$, FECHDOC$, TIPODOC$, DOCUORIG$, DESCRIDOC$, DEDOCOR$, CODOCUM$
Public CAPARDOC%, DOCPARAM$(256), CODPARAM$(256)
Public CAIMAGEN%, CAIMINI%, CODIMAGEN$(16), ARCHIMAGEN$(16)
Public CACOLTAB1%, CAITAB1%, TETABU1$(32, 1024), CODTABU1$(32)
Public CACOLTAB2%, CAITAB2%, TETABU2$(32, 1024), CODTABU2$(32)
Public DESTIDOC%   ' Destinatario del Documento
Public DESTIDOCNUE   ' Destinatario del Documento
Public CACOPDOC%   ' CANTIDAD DE COPIAS
'
Public ALTUPAGINA, TOPEPAGINA, ANCHOPAGI
Public PAGINACTIVA%, CAPAGINAS%
'
Public CODPAR$(2048), VALPAR$(2048), CAVALPAR%

Declare Function GetProfileString Lib "KERNEL32" Alias "GetProfileStringA" (ByVal lpAppName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long) As Long
Declare Function WriteProfileString Lib "KERNEL32" Alias "WriteProfileStringA" (ByVal lpszSection As String, ByVal _
lpszKeyName As String, ByVal lpszString As String) As Long

Public Const HWND_BROADCAST As Long = &HFFFF
Public Const WM_WININICHANGE As Long = &H1A



'
Sub SETSPOOL(COMANDO$, ATRIBU$)
   '
   CAREVEC& = CAREVEC& + 1
   UBU& = 0
   On Error Resume Next
   UBU& = UBound(VECPRICO$)
   On Error GoTo 0
   If CAREVEC& > UBU& Then
      ReDim Preserve VECPRICO$(UBU& + 128)
   End If
   '
   COPRIN$ = Space$(256)
   Call REPLA(COPRIN$, COMANDO$, 1, 16)
   Call REPLA(COPRIN$, ATRIBU$, 17, 240)
   'FORMALTER.PSPOOL.AddItem COPRIN$
   VECPRICO$(CAREVEC&) = COPRIN$
   '
End Sub
'
Sub SPOOLINE(H1, V1, H2, V2, CCOLO, BBXX$)
   ATRIBU$ = Space$(112)
   Call REPLA(ATRIBU$, str$(H1), 1, 16)
   Call REPLA(ATRIBU$, str$(V1), 17, 16)
   Call REPLA(ATRIBU$, str$(H2), 33, 16)
   Call REPLA(ATRIBU$, str$(V2), 49, 16)
   Call REPLA(ATRIBU$, str$(CCOLO), 65, 16)
   Call REPLA(ATRIBU$, BBXX$, 81, 16)
   Call SETSPOOL("LINE", ATRIBU$)
End Sub
'
Sub SPOOLIMAGE(IMAGEFILE$, PHX1, PVX1, ACHX, ALTX)
   '
   IMAGFI$ = TRIM$(IMAGEFILE$)
   If EXISTE%(IMAGFI$) < 1 Then Exit Sub
   '
   Dim IMAGFIDATE As Date
   IMAGFIDATE = FileDateTime(IMAGFI$)
   LARGIMAG& = FileLen(IMAGFI$)
   FECHIMAG# = CDbl(IMAGFIDATE)
   '
   NOMBIMAG$ = IMAGFI$
10 PPXX% = InStr(NOMBIMAG$, "\")
   If PPXX% > 0 Then
      NOMBIMAG$ = Mid$(NOMBIMAG$, PPXX% + 1)
      GoTo 10
   End If
   If TRIM$(NOMBIMAG$) = "" Then Exit Sub
   If Len(NOMBIMAG$) > 38 Then
      Call MENSERR(24, "Imposible Imprimir Imagen Contenida en Archivo\  \'" + NOMBIMAG$ + "'\  \Nombre de Archivo Demasiado Extenso (max = 38 car)")
      Exit Sub
   End If
   '
   ABUS$ = AJUSTI$(NOMBIMAG$, 38) + MKS$(LARGIMAG&) + MKD$(FECHIMAG#)
   BBUS% = BINVAL%(ABUS$)
   RFF$ = RECFILT$("PIMGLST", 1, MKI$(BBUS%))
   For KKI% = 1 To Len(RFF$) Step 4
     RELE& = CVS(Mid$(RFF$, KKI%, 4))
     If RELE& > 0 Then
       XX$ = REGLEIDO$("PIMGLST", RELE&)
       If Mid$(XX$, 3, 50) = ABUS$ Then
         IMAGENUMBER& = CVS(Mid$(XX$, 53, 4))
         GoTo 50
       End If
     End If
   Next KKI%
   '
   Call BLOQARCH("PIMGLST")
   Call BLOQARCH("PIMGSPL")
   '
   LASTIMAGENUMBER& = 0
   URELIS& = ULTREG&("PIMGLST")
   If URELIS& > 0 Then
     XX$ = REGLEIDO$("PIMGLST", URELIS&)
     On Error Resume Next
     LASTIMAGENUMBER& = CVS(Mid$(XX$, 51, 4))
     On Error GoTo 0
     If LASTIMAGENUMBER& < URELIS& Then LASTIMAGENUMBER& = URELIS&
   End If
   IMAGENUMBER& = 1 + LASTIMAGENUMBER&
   FirstRecord& = 1 + ULTREG&("PIMGSPL")
   '
   IMGHEADER$ = ABUS$
   BBUS% = BINVAL%(ABUS$)
   '
   PPXX% = InStr(IMAGFI$, ".")
   EXTAR$ = ".BMP"
   If PPXX% > 0 Then EXTAR$ = Mid$(IMAGFI$, PPXX%)
   FileCopy IMAGFI$, LUCOM$ + "IMAGEN" + EXTAR$
   'NX% = FILEOPEN%(IMAGFI$, 0, 508)
   NX% = FILEOPEN%(LUCOM$ + "IMAGEN" + EXTAR$, 0, 508)
   Dim BUFI As String * 508
   CARELE& = 1 + LOF(NX%) / 508
   jj& = FirstRecord& - 1
   For RELE& = 1 To CARELE&
     Get #NX%, RELE&, BUFI$
     ZZ$ = String$(512, 0)
     Call REPLA(ZZ$, MKS$(IMAGENUMBER&), 1, 4)
     Call REPLA(ZZ$, BUFI$, 5, 508)
     jj& = jj& + 1
     LastRecord& = jj&
     Call GRAREG("PIMGSPL", ZZ$, jj&)
   Next RELE&
   Call CIERRARCH("PIMGSPL")
   Close #NX%: NX% = 0
   '
   ZZ$ = String$(128, 0)
   Call REPLA(ZZ$, MKI$(BBUS%), 1, 2)
   Call REPLA(ZZ$, IMGHEADER$, 3, 50)
   Call REPLA(ZZ$, MKS$(IMAGENUMBER&), 53, 4)
   Call REPLA(ZZ$, MKS$(FirstRecord&), 57, 4)
   Call REPLA(ZZ$, MKS$(LastRecord&), 61, 4)
   Call REGAPP("PIMGLST", ZZ$)
   Call CIERRARCH("PIMGLST")
   '
50 ATRIBU$ = Space$(112)
   Call REPLA(ATRIBU$, str$(BBUS%), 1, 16)
   Call REPLA(ATRIBU$, str$(IMAGENUMBER&), 17, 16)
   Call REPLA(ATRIBU$, str$(PHX1), 33, 16)
   Call REPLA(ATRIBU$, str$(PVX1), 49, 16)
   Call REPLA(ATRIBU$, str$(ACHX), 65, 16)
   Call REPLA(ATRIBU$, str$(ALTX), 81, 16)
   Call SETSPOOL("PAINTPICTURE", ATRIBU$)
   '
End Sub
'
'
Sub ENVIAPORMAIL()
   '

   HOII% = HOY(HOS$)
   ASUNTO$ = DESCRIDOC$
   
   '
   TEXTOMAI$ = "Estamos enviando " + DESCRIDOC$ + CRLF$
   TEXTOMAI$ = TEXTOMAI$ + "como documento adjunto en formato PDF." + CRLF$
   TEXTOMAI$ = TEXTOMAI$ + CRLF$
   '
   'PARA QUE EL MAIL SALGA COMO ORDEN DE COMPRA EN EL ASUNTO
      EPAC$ = TRIM$(UCase$(EMPREAC$))

   If InStr(EPAC$, "AHP") > 0 And UCase$(Left$(DESCRIDOC$, 16)) = "ORDEN DE COMPRA " Then
       POS1% = InStr(DESCRIDOC$, " - ")
       DESCRIDOC1$ = Mid$(DESCRIDOC$, 1, POS1%)
       ASUNTOX$ = "AMERICAN HEAVY PARTS - " + DESCRIDOC1$
       ASUNTO$ = ASUNTOX$
       
       TEXTOMAI$ = "Estamos enviando " + DESCRIDOC1$ + CRLF$
       TEXTOMAI$ = TEXTOMAI$ + "como documento adjunto en formato PDF." + CRLF$
       TEXTOMAI$ = TEXTOMAI$ + CRLF$
   End If
   '
   
   'AGREGADO PARA DOSIVAC
   If InStr(EPAC$, "DOSIVA") > 0 And UCase$(Left$(DESCRIDOC$, 11)) = "PRESUPUESTO" Then
      TEXTOMAI$ = TEXTOMAI$ + "Si su presupuesto está hecho al cliente 6666 (Presupuestos)," + CRLF$
      TEXTOMAI$ = TEXTOMAI$ + "y de aceptar dicho presupuesto debería enviarnos" + CRLF$
      TEXTOMAI$ = TEXTOMAI$ + "datos para facturar (nombre de transporte en capital)" + CRLF$
      TEXTOMAI$ = TEXTOMAI$ + "recordando que el pago es anticipado."
      TEXTOMAI$ = TEXTOMAI$ + "" + CRLF$
   End If
   '
   'AGREGADO PARA AHP
   If InStr(EPAC$, "AHP") > 0 And UCase$(Left$(DESCRIDOC$, 10)) = "COTIZACION" Then
       
      'DESCRIDOC1$ = REPLACAR$(UCase$(DESCRIDOC$), "PRESUPUESTO ", "COTIZACION ")
      POS1% = InStr(DESCRIDOC$, " - ")
      DESCRIDOC1$ = Mid$(DESCRIDOC$, 1, POS1%)
      ASUNTOX$ = "AMERICAN HEAVY PARTS - " + DESCRIDOC1$
      If TRIM$(FRMPRESUAHP.TXTREFERENCIA) <> "" Then ASUNTOX$ = ASUNTOX$ + " - " & TRIM$(FRMPRESUAHP.TXTREFERENCIA)
      ASUNTO$ = ASUNTOX$
      TEXTOMAI$ = "Estamos enviando " + DESCRIDOC1$ + CRLF$
      TEXTOMAI$ = TEXTOMAI$ + "como documento adjunto en formato PDF." + CRLF$
      TEXTOMAI$ = TEXTOMAI$ + CRLF
      TEXTOMAI$ = TEXTOMAI$ + CONTROL("PRESPAHP", "TEXTMAIL")
      TEXTOMAI$ = TEXTOMAI$ + CRLF
      
   End If
   '
   'AGREGADO PARA MEDITEA
   'SI HAY QUE REALIZAR ALGUN CAMBIO EN EL TEXTO VER TEXTO DE ORDEN DE COMPRA QUE SE CARGA DESDE EL SETUP DE COMPRAS
   ' Y VER CODIGO MAS ABAJO  (BUSCAR: 'TEXTO DE MAIL PARA ORDEN DE COMPRA FIJO MAS TEXTO VARIALBE QUE SE CARGA EN SETUP DE COMPRAS)
   If InStr(EPAC$, "MEDITEA") > 0 And UCase$(Left$(DESCRIDOC$, 14)) = "NOTA DE PEDIDO" Then
      'ENCABEZADO
      TEXTOMAIENC$ = "POR FAVOR CONFIRMAR LA RECEPCION DE LA ORDEN DE COMPRA" + CRLF$
      TEXTOMAIENC$ = TEXTOMAIENC$ + String$(64, "*") + CRLF$ + CRLF$
      'CUERPO
      TEXTOMAIMED$ = "Meditea Electromédica SRL" + CRLF$ + CRLF$
      TEXTOMAIMED$ = TEXTOMAIMED$ + "Viamonte 2255 - C1056ABI" + CRLF$
      TEXTOMAIMED$ = TEXTOMAIMED$ + "Buenos Aires - Argentina" + CRLF$
      TEXTOMAIMED$ = TEXTOMAIMED$ + "TEL: (54-11) 4953-6114 " + CRLF$
      TEXTOMAIMED$ = TEXTOMAIMED$ + "TEL / FAX (54-11) 4953-8273 " + CRLF$
      TEXTOMAIMED$ = TEXTOMAIMED$ + "Site: www.meditea.com " + CRLF$
      '
      TEXTOMAI$ = TEXTOMAIENC$
   End If
   '
   'AGREGADO PARA GABAL
   If InStr(EPAC$, "GABAL") > 0 And UCase$(Left$(DESCRIDOC$, 17)) = "PEDIDO DE CLIENTE" Then
      'ENCABEZADO
      TEXTOMAIENC$ = "Estamos enviando la orden de venta nro. " + TRIM$(Mid$(DESCRIDOC$, 18)) + " como documento adjunto" + CRLF$
      TEXTOMAIENC$ = TEXTOMAIENC$ + "en formato PDF." + CRLF$ + CRLF$
'      TEXTOMAIENC$ = TEXTOMAIENC$ + "Solicitamos envíen el detalle del pago recordando que el plazo máximo de diferimiento de" + CRLF$
'      TEXTOMAIENC$ = TEXTOMAIENC$ + "valores es de 30 días, en el detalle debe constar:" + CRLF$ + CRLF$
'      '
'      TEXTOMAIENC$ = TEXTOMAIENC$ + "Nro. de CUIT del firmante" + CRLF$
'      TEXTOMAIENC$ = TEXTOMAIENC$ + "Razón Social del Firmante" + CRLF$
'      TEXTOMAIENC$ = TEXTOMAIENC$ + "Nro. de Cheque" + CRLF$
'      TEXTOMAIENC$ = TEXTOMAIENC$ + "Banco" + CRLF$
'      TEXTOMAIENC$ = TEXTOMAIENC$ + "Importe" + CRLF$
'      TEXTOMAIENC$ = TEXTOMAIENC$ + "Fecha diferida" + CRLF$ + CRLF$
'      '
'      TEXTOMAIENC$ = TEXTOMAIENC$ + "Si Ud. es nuevo cliente o no realizo operaciones durante el último año las operaciones deberán" + CRLF$
'      TEXTOMAIENC$ = TEXTOMAIENC$ + "realizarse únicamente de contado efectivo por medio de transferencia bancaria" + CRLF$ + CRLF$
'      '
'      TEXTOMAIENC$ = TEXTOMAIENC$ + "BANCO SANTANDER RIO" + CRLF$
'      TEXTOMAIENC$ = TEXTOMAIENC$ + "CBU 07201949 20000001174290" + CRLF$
'      TEXTOMAIENC$ = TEXTOMAIENC$ + "CTA CORRIENTE PESOS N° 194-11742/9" + CRLF$
'      TEXTOMAIENC$ = TEXTOMAIENC$ + "LAFE SA" + CRLF$
'      TEXTOMAIENC$ = TEXTOMAIENC$ + "CUIT  33-64415725-9" + CRLF$
      '
      TEXTOMAI$ = TEXTOMAIENC$
   End If
   '
   TEXTOMAI$ = TEXTOMAI$ + "" + CRLF$
   TEXTOMAI$ = TEXTOMAI$ + "El documento puede ser visualizado mediante ACROBAT Reader." + CRLF$
   TEXTOMAI$ = TEXTOMAI$ + "En caso de no estar instalado este utilitario, puede descargarlo" + CRLF$
   TEXTOMAI$ = TEXTOMAI$ + "de www.adobe.es/products/acrobat/readstep2.html" + CRLF$
   TEXTOMAI$ = TEXTOMAI$ + "" + CRLF$
'   TEXTOMAI$ = TEXTOMAI$ + "Para Auto-Presentar, puede bajar el utilitario FLX_MSGR.EXE" + CRLF$
'   TEXTOMAI$ = TEXTOMAI$ + "de www.flexoft.com.ar." + CRLF$
'   TEXTOMAI$ = TEXTOMAI$ + "" + CRLF$
   '
   If InStr(EPAC$, "GABAL") > 0 And UCase$(Left$(DESCRIDOC$, 17)) = "PEDIDO DE CLIENTE" Then
      'ENCABEZADO
      TEXTOMAIENC$ = "Saludos cordiales" + CRLF$ + CRLF$
      TEXTOMAIENC$ = TEXTOMAIENC$ + "Dto. Comercial" + CRLF$ + CRLF$
      TEXTOMAIENC$ = TEXTOMAIENC$ + "LAFE S.A. / GB GABAL" + CRLF$
      TEXTOMAIENC$ = TEXTOMAIENC$ + "Pasaje Jorge Newbery 1137 Avellaneda" + CRLF$
      TEXTOMAIENC$ = TEXTOMAIENC$ + "Buenos Aires - Argentina" + CRLF$
      TEXTOMAIENC$ = TEXTOMAIENC$ + "www.gbgabal.com / pedidos@gbgabal.com" + CRLF$
      TEXTOMAIENC$ = TEXTOMAIENC$ + "+ 54 11 4139-7010 Ext 11"
      '
      TEXTOMAI$ = TEXTOMAI$ + TEXTOMAIENC$
   End If
   '
   'TEXTO DE MAIL PARA ORDEN DE COMPRA FIJO MAS TEXTO VARIALBE QUE SE CARGA EN SETUP DE COMPRAS
   If UCase$(Left$(DESCRIDOC$, 14)) = "NOTA DE PEDIDO" And InStr(EPAC$, "MEDITEA") < 1 Then ' CASO DE ORDEN DE COMPRA
     Dim FORMULAACTIVO As Form 'PARA QUE TOME EL FORMULARIO ACTIVO ORIGINAL O MODIFICACION
     Set FORMULAACTIVO = FORCOMP07
     If FOMODOC07.Visible = True Then Set FORMULAACTIVO = FOMODOC07
     TXTCOMPRA$ = "Nota de Pedido: " & TRIM$(FORMULAACTIVO.Text11) & "-" & Right("00" & XVALO(FORMULAACTIVO.Label16), 2) & vbCrLf
     If FOMODOC07.Visible = True Then
        TXTCOMPRA$ = TXTCOMPRA$ + "Proveedor: " & rasocli$(-1 * XVALO(FORMULAACTIVO.LTEXT1)) & vbCrLf
     Else
        TXTCOMPRA$ = TXTCOMPRA$ + "Proveedor: " & rasocli$(-1 * XVALO(FORMULAACTIVO.Text1)) & vbCrLf
     End If
     '
     TXTCOMPRA$ = TXTCOMPRA$ + "Fecha de Emisión: " & FORMULAACTIVO.Text9.TEXT & vbCrLf & vbCrLf
     TXTCOMPRA$ = TXTCOMPRA$ & TEXTO_MAIL_COMPRA$ ' FUNCION QUE DEVUELVE EL TEXTO DE MAIL GRABADO EN SETUP DE COMPRAS
     TEXTOMAI$ = TXTCOMPRA$
   End If
   ' FIN **  TEXTO DE MAIL PARA ORDEN DE COMPRA FIJO MAS TEXTO VARIALBE QUE SE CARGA EN SETUP DE COMPRAS
   '
   MAI_DESTI$ = TRIM$(SENDBYMAIL$)
   MAI_DIREL$ = MAI_DESTI$                 ' "ebauer@flexoft.com.ar"
   MAI_ASUNT$ = ASUNTO$
   MAI_TEXTO$ = TEXTOMAI$
   MAI_ADJUN$(0) = LUCOM$ + "" + TRIM$(EDITFORM.Label1) + ".PDF"
   EDITFORM.Label2 = ""
   '
   Call GENERAMAIL
   ATENTI% = 1
   '
90 For Each PRX1 In Printers
     If PRX1.DeviceName = IMPREDET$ Then
       On Error Resume Next
       Set Printer = PRX1
       On Error GoTo 0
     End If
   Next
   '
End Sub
Function TEXTO_MAIL_COMPRA$()
      '
      '' FUNCION QUE DEVUELVE EL TEXTO DE MAIL GRABADO EN SETUP DE COMPRAS
      '
      TEXTO_MAIL_COMPRA$ = ""
      '
      Call COPYSTRU("OBSERVOF", "TXMAICOM")
      For U& = 1 To ULTREG&("TXMAICOM")
           TX$ = REGLEIDO$("TXMAICOM", U&)
           TEXTOBS$ = TEXTOBS$ & TRIM$(TX$) & vbCrLf
      Next U&
      '
      If TRIM$(TEXTOBS$) <> "" Then
         TEXTO_MAIL_COMPRA$ = TEXTOBS$
      End If
      '
End Function

'
Sub SETPRINPORT(NEWPORT$)
   '
   Static IMPREDETER$, CALIPREDET
   '
   NP$ = TRIM$(NEWPORT$)
   NOMESS% = 0
   PPXX% = InStr(NP$, "/NOMESS")
   If PPXX% > 0 Then
     NP$ = Left$(NP$, PPXX% - 1)
     NOMESS% = 1
   End If
   Dim PRX As Object
   '
   If IMPREDETER$ = "" Then
      On Error Resume Next
      IMPREDETER$ = Printer.DeviceName
      On Error GoTo 0
   End If
   '
   If NP$ = "RESTORE" Then
      'On Error Resume Next
      'Printer.PrintQuality = CALIPREDET
      'On Error GoTo 0
      If TRIM$(IMPREDETER$) <> "" Then
        For Each PRX In Printers
          If UCase$(PRX.Port) = UCase$(IMPREDETER$) Or UCase(PRX.DeviceName) = UCase(IMPREDETER$) Then
            On Error Resume Next
            Set Printer = PRX
            CALIPREDET = Printer.PrintQuality
            Printer.PrintQuality = (-1)         ' draft
            On Error GoTo 0
            Exit Sub
          End If
        Next PRX
      End If
      Exit Sub
   End If
   '
   If NP$ <> "" Then
      For Each PRX In Printers
        If UCase$(PRX.Port) = UCase$(NP$) Or UCase$(PRX.DeviceName) = UCase$(NP$) Then
          On Error Resume Next
          Set Printer = PRX
          CALIPREDET = Printer.PrintQuality
          Printer.PrintQuality = (-1)         ' draft
          On Error GoTo 0
          GoTo 99
        End If
      Next PRX
      '
      If NOMESS% = 0 Then
          On Error Resume Next
          NP3$ = Printer.DeviceName
          If Err Then
             Call COMUNI("Problemas Para Obtener Información\de la Impresora Predeterminada.\   \Resultados Impredecibles.")
             On Error GoTo 0
             GoTo 90
          End If
          NP1$ = REPLACAR$(NP$, "\", "/")
          NP2$ = REPLACAR$(Printer.DeviceName, "\", "/")
          Call COMUNI("Puerto de Impresión '" + NP1$ + "' no Encontrado.\  \El Comprobante Será Emitido por la\Impresora Predeterminada '" + NP2$ + "'")
          GoTo 99
          'Call FINAL("")
        Else
90        NEWPORT$ = ""
      End If
   End If
   '
99 End Sub
'
Sub SETPRINPORT2(NEWPORT$)
    '
    Static IMPREDETER$, CALIPREDET
    Dim objPrinter As Printer
    '
    NP$ = TRIM$(NEWPORT$)
    NOMESS% = 0
    PPXX% = InStr(NP$, "/NOMESS")
    If PPXX% > 0 Then
        NP$ = Left$(NP$, PPXX% - 1)
        NOMESS% = 1
    End If
    Dim PRX As Object
    '
    If IMPREDETER$ = "" Then
        Set objPrinter = ObtenerImpresoraPredeterminada()
        IMPREDETER$ = objPrinter.DeviceName
    End If
    '
    If NP$ = "RESTORE" Then
        If TRIM$(IMPREDETER$) <> "" Then
            Configurado% = EstablecerImpresoraPredeterminada(IMPREDETER$)
        End If
        Exit Sub
    End If
    '
    If NP$ <> "" Then
        Configurado% = EstablecerImpresoraPredeterminada(NP$)
        If Configurado% = 1 Then GoTo 99
        '
        If NOMESS% = 0 Then
             On Error Resume Next
             NP3$ = Printer.DeviceName
             If Err Then
                Call COMUNI("Problemas Para Obtener Información\de la Impresora Predeterminada.\   \Resultados Impredecibles.")
                On Error GoTo 0
                GoTo 90
             End If
             NP1$ = REPLACAR$(NP$, "\", "/")
             NP2$ = REPLACAR$(Printer.DeviceName, "\", "/")
             Call COMUNI("Puerto de Impresión '" + NP1$ + "' no Encontrado.\  \El Comprobante Será Emitido por la\Impresora Predeterminada '" + NP2$ + "'")
             GoTo 99
         Else
90           NEWPORT$ = ""
       End If
    End If
    '
99 End Sub

Function EstablecerImpresoraPredeterminada(IMPRESORA As String) As Integer

    For Each Prt In Printers
    If InStr(Prt.DeviceName, IMPRESORA$) > 0 Then Set Printer = Prt: Exit For
    Next
    
    On Error GoTo ErrorSeteoImpresora
    R = WriteProfileString("windows", "Device", Printer.DeviceName + "," + Printer.DriverName + "," + Printer.Port)
    
    ' LA SIGUIENTE LINEA SIRVE PARA QUE WINDOWS LE AVISE AL RESTO DE LAS APLICACIONES QUE HA CAMBIADO
    ' LA IMPRESORA PREDETERMINADA. SIN EMBARGO, EN WINDOWS 10 EJECUTAR ESTA LINEA DEMORA MUCHISIMO TIEMPO, Y ADEMAS
    ' NO DEVUELVE EL CONTROL A LA APLICACION QUE LA EJECUTO HASTA QUE FINALIZA POR COMPLETO.
    ' COMO SU USO ES CAMBIAR LA IMPRESORA PARA IMPRIMIR ALGO MOMENTANEO Y LUEGO VOLVER A LA PREDETERMINADA, SE ANULA
    ' DICHA LINEA YA QUE EN DEFINITIVA TODO VUELVA A QUEDAR COMO ESTABA INICIALMENTE.
    
    ' L = SendMessage(HWND_BROADCAST, WM_WININICHANGE, 0, ByVal "windows")
    
    On Error GoTo 0
    EstablecerImpresoraPredeterminada = 1
    
    Exit Function

ErrorSeteoImpresora:
    EstablecerImpresoraPredeterminada = 0
    Exit Function
    
End Function

Function ObtenerImpresoraPredeterminada() As Printer

    Dim strBuffer As String * 254
    Dim iRetValue As Long
    Dim strDefaultPrinterInfo As String
    Dim tblDefaultPrinterInfo() As String
    
    iRetValue = GetProfileString("windows", "device", ",,,", strBuffer, 254)
    strDefaultPrinterInfo = Left(strBuffer, InStr(strBuffer, Chr(0)) - 1)
    tblDefaultPrinterInfo = Split(strDefaultPrinterInfo, ",")
    
    For Each objPrinter In Printers
        If objPrinter.DeviceName = tblDefaultPrinterInfo(0) Then
        Exit For
        End If
    Next
    
    'Si no se encuentra devuelve Nothing
    If objPrinter.DeviceName <> tblDefaultPrinterInfo(0) Then
        Set objPrinter = Nothing
    End If
    
    Set ObtenerImpresoraPredeterminada = objPrinter
    
End Function

Sub PRINTCOMMAND(COPRI%, ATRIP$)
   '
   Static LASTX, LASTY, LASTFONTNAME$, CTXYZ%
   '
   Select Case COPRI%
   Case 1
     On Error Resume Next
        '   MsgBox "SCALEMODE: " & Printer.ScaleMode
        '   MsgBox "SCALEWIDTH: " & Printer.ScaleWidth
        '   MsgBox "SCALELEFT: " & Printer.ScaleLeft
        '   MsgBox "SCALEHEIGHT: " & Printer.ScaleHeight
        '   MsgBox "SCALETOP: " & Printer.ScaleTop
        '   MsgBox "PAPERSIZE: " & Printer.PaperSize
     Printer.Orientation = XVALO(TRIM$(ATRIP$))
        '   MsgBox "SCALEMODE: " & Printer.ScaleMode
        '   MsgBox "SCALEWIDTH: " & Printer.ScaleWidth
        '   MsgBox "SCALELEFT: " & Printer.ScaleLeft
        '   MsgBox "SCALEHEIGHT: " & Printer.ScaleHeight
        '   MsgBox "SCALETOP: " & Printer.ScaleTop
        '   MsgBox "PAPERSIZE: " & Printer.PaperSize
     On Error GoTo 0
   Case 2
     If MATRIXPUN% = 1 Then
        Call VUELCAIMPRE(1)
        If HOJASUEL% = 1 Then
          Call COMUNI("Coloque Formulario y Pulse Aceptar")
        End If
        Exit Sub
     End If
     If HOJASUEL% = 1 Then
       Printer.EndDoc
       Call COMUNI("Coloque Formulario y Pulse Aceptar")
       Exit Sub
     End If
     On Error GoTo 90
     Printer.NewPage
     On Error GoTo 0
   Case 3
     If MATRIXPUN% = 1 Then
        Call VUELCAIMPRE(2)
        Exit Sub
     End If
     On Error GoTo 90
     Printer.EndDoc
     On Error GoTo 0
   Case 4
     On Error Resume Next
     'Printer.Height = XVALO(TRIM$(ATRIP$))
     On Error GoTo 0
   Case 5
     On Error Resume Next
     'Printer.Width = XVALO(TRIM$(ATRIP$))
     On Error GoTo 0
   Case 6
     Printer.ScaleMode = XVALO(TRIM$(ATRIP$))
   Case 7
     Printer.ForeColor = XVALO(TRIM$(ATRIP$))
   Case 8
     Printer.DrawMode = XVALO(TRIM$(ATRIP$))
   Case 9
     Printer.DrawWidth = XVALO(TRIM$(ATRIP$))
   Case 10
     Printer.DrawStyle = XVALO(TRIM$(ATRIP$))
   Case 11
     Printer.FillStyle = XVALO(TRIM$(ATRIP$))
   Case 12
     AA11& = XVALO(Mid$(ATRIP$, 1, 16))
     BB11& = XVALO(Mid$(ATRIP$, 17, 16))
     AA22& = XVALO(Mid$(ATRIP$, 33, 16))
     BB22& = XVALO(Mid$(ATRIP$, 49, 16))
     CCOLO& = XVALO(Mid$(ATRIP$, 65, 16))
     BBLO$ = UCase$(TRIM$(Mid$(ATRIP$, 81, 16)))
     If BBLO$ = "B" Then
        Printer.Line (AA11&, BB11&)-(AA22&, BB22&), CCOLO&, B
       Else
        Printer.Line (AA11&, BB11&)-(AA22&, BB22&), CCOLO&
     End If
   Case 13
     Printer.FontName = TRIM$(ATRIP$)
     LASTFONTNAME$ = TRIM$(ATRIP$)
   Case 14
     Printer.FontItalic = XVALO(TRIM$(ATRIP$))
   Case 15
     Printer.FontSize = XVALO(TRIM$(ATRIP$))
   Case 16
     Printer.FontBold = XVALO(TRIM$(ATRIP$))
   Case 17
     If XVALO(TRIM$(ATRIP$)) >= Printer.ScaleWidth - 10 Then
          Printer.CurrentX = Printer.ScaleWidth - 10
        Else
          Printer.CurrentX = Int(XVALO(TRIM$(ATRIP$)) + 0.5)
     End If
     LASTX = Int(XVALO(TRIM$(ATRIP$)) + 0.5)
   Case 18
     If XVALO(TRIM$(ATRIP$)) >= Printer.ScaleHeight - 10 Then
          Printer.CurrentY = Printer.ScaleHeight - 10
        Else
          Printer.CurrentY = Int(XVALO(TRIM$(ATRIP$)) + 0.5)
     End If
     LASTY = Int(XVALO(TRIM$(ATRIP$)) + 0.5)
   Case 19
     ATRIP$ = RTrim$(ATRIP$)
     If TRIM$(ATRIP$) = "" Or TRIM$(ATRIP$) = ";" Then Exit Sub
     '
     PRUVUELTA% = 0
191  If Printer.CurrentX <> LASTX Or Printer.CurrentY <> LASTY Then
       If PRUVUELTA% < 3 Then
         PRUVUELTA% = PRUVUELTA% + 1
         Printer.CurrentX = LASTX
         Printer.CurrentY = LASTY
         GoTo 191
       End If
       MsgBox "Imposible Imprimir '" + TRIM$(ATRIP$) + "' en Coordenadas " + TRIM$(str$(LASTX)) + "," + TRIM$(str$(LASTY))
       Exit Sub
     End If
     '
     PRUVUELTA% = 0
192  If Printer.FontName <> LASTFONTNAME$ Then
       If PRUVUELTA% < 3 Then
         PRUVUELTA% = PRUVUELTA% + 1
         Printer.FontName = LASTFONTNAME$
         GoTo 192
       End If
       FUENODISPO% = 1
     End If
     '
     If MATRIXPUN% = 1 Then
        Call CARLIMPRE(Printer.CurrentX, Printer.CurrentY, ATRIP$)
        Exit Sub
     End If
     ATRIXX$ = ATRIP$
     If Right$(ATRIP$, 1) = ";" Then ATRIXX$ = Left$(ATRIP$, Len(ATRIP$) - 1)
     ATRIXY$ = ATRIXX$
     If Printer.CurrentX + Printer.TextWidth(RTrim$(ATRIXX$)) >= Printer.ScaleWidth Then
193    If Len(ATRIXX$) > 0 Then
         If Printer.CurrentX + Printer.TextWidth(RTrim$(ATRIXX$)) >= Printer.ScaleWidth Then
           ATRIXX$ = Left$(ATRIXX$, Len(ATRIXX$) - 1)
           FUERADEAREA% = 1
           GoTo 193
         End If
       End If
       If Len(ATRIXX$) < 1 Then
         FUERADEAREA% = 1
         'MsgBox "Texto '" + TRIM$(ATRIXY$) + "' Fuera de Area de Impresion Horizontal - Suprimido"
         GoTo 99
       End If
       If Right$(ATRIP$, 1) = ";" Then ATRIXX$ = ATRIXX$ + ";"
       ATRIP$ = ATRIXX$
       '
     End If
     If Printer.CurrentY + Printer.TextHeight(RTrim$(ATRIXX$)) >= Printer.ScaleHeight Then
       FUERADEAREA% = 1
       GoTo 99
     End If
     '
195  If Right$(ATRIP$, 1) = ";" Then
         ATRIP$ = Left$(ATRIP$, Len(ATRIP$) - 1)
         Printer.Print RTrim$(ATRIP$);
        Else
         Printer.Print RTrim$(ATRIP$)
     End If
     '
   Case 20
     Printer.FontUnderline = XVALO(TRIM$(ATRIP$))
   Case 21
     Call INCLUDEIMAGE(ATRIP$, "PRINT")
   End Select
   GoTo 99
   '
90 ERRNU = Err
   Resume 91
91 On Error GoTo 0
   Call MENSERR(24, "Error (" & ERRNU & ") de Impresion\   \Revise Impresora Desconectada o Fuera de Linea,\Atasco de Papel o Saturación de Cola de Impresión.")
   '
99 End Sub
'

Sub CARLIMPRE(PPX, PPY, ATRIP$)
   TTXX$ = CSTRING$(MKS$(PPY), 3, 8, 0, 2) + CSTRING$(MKS$(PPX), 3, 8, 0, 2) + ATRIP$
   For KKL& = 1 To FORMALTER.List1.ListCount
     If FORMALTER.List1.List(KKL& - 1) > TTXX$ Then
       FORMALTER.List1.AddItem TTXX$, KKL& - 1
       Exit Sub
     End If
   Next KKL&
   FORMALTER.List1.AddItem TTXX$
End Sub

Sub VUELCAIMPRE(MODOVUEL%)
   '
   ' MODOVUEL%=1 --> NEWPAGE
   ' MODOVUEL%=2 --> ENDDOC
   '
   EPAC$ = TRIM$(UCase$(EMPREAC$))
   If InStr(EPAC$, "DOSIVA") > 0 Then
      On Error Resume Next
      Printer.PrintQuality = (-1)         ' draft
      Printer.FontName = "Draft 12cpi"
      Printer.FontBold = False
      Printer.FontItalic = False
      Printer.FontUnderline = False
      Printer.FontSize = 10
      Printer.PrintQuality = (-1)
      On Error GoTo 0
   End If
   '
   HOII% = HOY(HOS$)
   NXX% = FreeFile
   Open LUCOM$ + "CTRLPRIN.XXX" For Append As #NXX%
   Print #NXX%, "Antes de " + TIPODOC$ + " " + NUMEDOC$
   Print #NXX%, HOS$ + " - " + Time$
   Print #NXX%, "Letra: "; Printer.FontName
   Print #NXX%, "Negrita: "; Printer.FontBold
   Print #NXX%, "Italica: "; Printer.FontItalic
   Print #NXX%, "Subrayado: "; Printer.FontUnderline
   Print #NXX%, "Tamano: "; Printer.FontSize
   Print #NXX%, "Calidad: "; Printer.PrintQuality
   Print #NXX%, "   "
   '
   For KKL& = 1 To FORMALTER.List1.ListCount
     TTXX$ = FORMALTER.List1.List(KKL& - 1)
     PPY = XVALO(Left$(TTXX$, 8))
     PPX = XVALO(Mid$(TTXX$, 9, 8))
     ATRIP$ = Mid$(TTXX$, 17)
     '
 If PPX + Printer.TextWidth(RTrim$(ATRIP$)) >= Printer.ScaleWidth Then
    FUERADEAREA% = 1
    GoTo 99
 End If
 If PPY + Printer.TextHeight(RTrim$(ATRIP$)) >= Printer.ScaleHeight Then
    FUERADEAREA% = 1
    GoTo 99
 End If

     Printer.CurrentY = PPY
     Printer.CurrentX = PPX
     If Right$(ATRIP$, 1) = ";" Then
         ATRIP$ = Left$(ATRIP$, Len(ATRIP$) - 1)
         Printer.Print RTrim$(ATRIP$);
        Else
         Printer.Print RTrim$(ATRIP$)
     End If
99 Next KKL&
   If MODOVUEL% = 1 Then
        Printer.NewPage
      Else
        Printer.EndDoc
   End If
   FORMALTER.List1.Clear
   '
   Print #NXX%, "Despues de " + TIPODOC$ + " " + NUMEDOC$
   Print #NXX%, HOS$ + " - " + Time$
   Print #NXX%, "Letra: "; Printer.FontName
   Print #NXX%, "Negrita: "; Printer.FontBold
   Print #NXX%, "Italica: "; Printer.FontItalic
   Print #NXX%, "Subrayado: "; Printer.FontUnderline
   Print #NXX%, "Tamano: "; Printer.FontSize
   Print #NXX%, "Calidad: "; Printer.PrintQuality
   Print #NXX%, "................................................."
   Close #NXX%
   '
End Sub
'
Sub SHOWCOMMAND(COPRI%, ATRIP$)
   '
   If PAGINACTIVA% > CAPAGINAS% Then
      CAPAGINAS% = PAGINACTIVA%
   End If
10 On Error Resume Next
   Select Case COPRI%
   Case 1
     'Printer.Orientation = XVALO(TRIM$(ATRIP$))
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
     Call CREACONTROL
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
     Call CREACONTROL
   Case 4
     'ALTUPAGINA = XVALO(TRIM$(ATRIP$))
     'For KKI% = 0 To CAPAGINAS%
     '  If EDITFORM.Picture1(KKI%).Height < ALTUPAGINA Then
     '    EDITFORM.Picture1(KKI%).Height = ALTUPAGINA
     '  End If
     'Next KKI%
   Case 5
     EDITFORM.Picture1(PAGINACTIVA%).Width = XVALO(TRIM$(ATRIP$))
   Case 6
     EDITFORM.Picture1(PAGINACTIVA%).ScaleMode = XVALO(TRIM$(ATRIP$))
   Case 7
     EDITFORM.Picture1(PAGINACTIVA%).ForeColor = XVALO(TRIM$(ATRIP$))
   Case 8
     EDITFORM.Picture1(PAGINACTIVA%).DrawMode = XVALO(TRIM$(ATRIP$))
   Case 9
     EDITFORM.Picture1(PAGINACTIVA%).DrawWidth = XVALO(TRIM$(ATRIP$))
   Case 10
     EDITFORM.Picture1(PAGINACTIVA%).DrawStyle = XVALO(TRIM$(ATRIP$))
   Case 11
     EDITFORM.Picture1(PAGINACTIVA%).FillStyle = XVALO(TRIM$(ATRIP$))
   Case 12
     AA11 = XVALO(Mid$(ATRIP$, 1, 16))
     BB11 = XVALO(Mid$(ATRIP$, 17, 16))
     AA22 = XVALO(Mid$(ATRIP$, 33, 16))
     BB22 = XVALO(Mid$(ATRIP$, 49, 16))
     CCOLO = XVALO(Mid$(ATRIP$, 65, 16))
     If BB11 + 200 > ALTUPAGINA Then ALTUPAGINA = BB11 + 200
     If BB22 + 200 > ALTUPAGINA Then ALTUPAGINA = BB22 + 200
     For KKI% = 0 To CAPAGINAS%
       If EDITFORM.Picture1(KKI%).Height < ALTUPAGINA - 1 Then
         EDITFORM.Picture1(KKI%).Height = ALTUPAGINA
       End If
     Next KKI%
     If AA11 > EDITFORM.Picture1(PAGINACTIVA%).Width Then
        EDITFORM.Picture1(PAGINACTIVA%).Width = AA11
     End If
     If AA22 > EDITFORM.Picture1(PAGINACTIVA%).Width Then
        EDITFORM.Picture1(PAGINACTIVA%).Width = AA22
     End If
     BBLO$ = UCase$(TRIM$(Mid$(ATRIP$, 81, 16)))
     If BBLO$ = "B" Then
        EDITFORM.Picture1(PAGINACTIVA%).Line (AA11, BB11 + TOPEPAGINA)-(AA22, BB22 + TOPEPAGINA), CCOLO, B
       Else
        EDITFORM.Picture1(PAGINACTIVA%).Line (AA11, BB11 + TOPEPAGINA)-(AA22, BB22 + TOPEPAGINA), CCOLO
     End If
     MayorAnchoPicture = AA11
     If MayorAnchoPicture > ANCHOPAGI Then ANCHOPAGI = MayorAnchoPicture
     MayorAnchoPicture = AA22
     If MayorAnchoPicture > ANCHOPAGI Then ANCHOPAGI = MayorAnchoPicture
   Case 13
     EDITFORM.Picture1(PAGINACTIVA%).FontName = TRIM$(ATRIP$)
   Case 14
     EDITFORM.Picture1(PAGINACTIVA%).FontItalic = XVALO(TRIM$(ATRIP$))
   Case 15
     EDITFORM.Picture1(PAGINACTIVA%).FontSize = XVALO(TRIM$(ATRIP$))
   Case 16
     EDITFORM.Picture1(PAGINACTIVA%).FontBold = XVALO(TRIM$(ATRIP$))
   Case 17
     If XVALO(TRIM$(ATRIP$)) > EDITFORM.Picture1(PAGINACTIVA%).Width Then
        EDITFORM.Picture1(PAGNACTIVA%).Width = XVALO(TRIM$(ATRIP$))
     End If
     EDITFORM.Picture1(PAGINACTIVA%).CurrentX = XVALO(TRIM$(ATRIP$))
     'MayorAnchoPicture = EDITFORM.Picture1(PAGINACTIVA%).CurrentX
     'If MayorAnchoPicture > ANCHOPAGI Then ANCHOPAGI = MayorAnchoPicture
   Case 18
     If XVALO(TRIM$(ATRIP$)) + 200 > ALTUPAGINA Then ALTUPAGINA = XVALO(TRIM$(ATRIP$)) + 1000
     For KKI% = 0 To CAPAGINAS%
       If EDITFORM.Picture1(KKI%).Height < ALTUPAGINA - 1 Then
         EDITFORM.Picture1(KKI%).Height = ALTUPAGINA
       End If
     Next KKI%
     EDITFORM.Picture1(PAGINACTIVA%).CurrentY = XVALO(TRIM$(ATRIP$)) + TOPEPAGINA
   Case 19
     ATRIP$ = RTrim$(ATRIP$)
     If Right$(ATRIP$, 1) = ";" Then
         ATRIP$ = Left$(ATRIP$, Len(ATRIP$) - 1)
         EDITFORM.Picture1(PAGINACTIVA%).Print ATRIP$;
        Else
         EDITFORM.Picture1(PAGINACTIVA%).Print ATRIP$
     End If
     a = EDITFORM.Picture1(PAGINACTIVA%).ScaleMode
     MayorAnchoPicture = EDITFORM.Picture1(PAGINACTIVA%).TextWidth(TRIM$(ATRIP$))
     If MayorAnchoPicture > ANCHOPAGI Then ANCHOPAGI = MayorAnchoPicture
   Case 20
     EDITFORM.Picture1(PAGINACTIVA%).FontUnderline = XVALO(TRIM$(ATRIP$))
   Case 21
     REDUFORMU = 1
     Call INCLUDEIMAGE(ATRIP$, "SHOW")
   End Select
   On Error GoTo 0
   '
   If ANCHOPAGI > 14500 Then
   a = b
   End If
End Sub
'
Sub CREACONTROL()
   NUPIC% = EDITFORM.Picture1.Count
   If PAGINACTIVA% >= NUPIC% - 1 Then
     Load EDITFORM.Picture1(NUPIC%)
     Load EDITFORM.Picture2(NUPIC%)
   End If
End Sub
'
Sub INCLUDEIMAGE(ATRIP$, MODOPRE$)
   '
   BBUS% = XVALO(Mid$(ATRIP$, 1, 16))
   IMAGENUMBER& = XVALO(Mid$(ATRIP$, 17, 16))
   PHX1 = XVALO(Mid$(ATRIP$, 33, 16))
   PVX1 = XVALO(Mid$(ATRIP$, 49, 16))
   ACHX = XVALO(Mid$(ATRIP$, 65, 16))
   ALTX = XVALO(Mid$(ATRIP$, 81, 16))
   '
   RFF$ = RECFILT$("PIMGLST", 1, MKI$(BBUS%))
   For KKI% = 1 To Len(RFF$) Step 4
     RELI& = CVS(Mid$(RFF$, KKI%, 4))
     If RELI& > 0 Then
       ZZ$ = REGLEIDO$("PIMGLST", RELI&)
       If CVS(Mid$(ZZ$, 53, 4)) = IMAGENUMBER& Then GoTo 10
     End If
   Next KKI%
   If Len(RFF$) < 4 Then
      Call MENSERR(24, "Imagen no Encontrada - Imposible Incluir.")
      Exit Sub
   End If
   '
10 NOMBIMAG$ = TRIM$(Mid$(ZZ$, 3, 38))
   LARGIMAG& = CVS(Mid$(ZZ$, 41, 4))
   FECHIMAG# = CVD(Mid$(ZZ$, 45, 8))
   FirstRecord& = CVS(Mid$(ZZ$, 57, 4))
   LastRecord& = CVS(Mid$(ZZ$, 61, 4))
   '
   If FirstRecord& < 1 Or LastRecord& < FirstRecord& Or LastRecord& > ULTREG&("PIMGSPL") Then
      Call MENSERR(24, "Imagen no Encontrada - Imposible Incluir.")
      Exit Sub
   End If
   '
   If FirstRecord& > 1 Then
      XX$ = REGLEIDO$("PIMGSPL", FirstRecord& - 1)
      If CVS(Left$(XX$, 4)) = IMAGENUMBER& Then
        Call MENSERR(24, "Imagen no Encontrada - Imposible Incluir.")
        Exit Sub
      End If
   End If
   '
   If LastRecord& < ULTREG&("PIMGSPL") Then
      XX$ = REGLEIDO$("PIMGSPL", LastRecord& + 1)
      If CVS(Left$(XX$, 4)) = IMAGENUMBER& Then
        Call MENSERR(24, "Imagen no Encontrada - Imposible Incluir.")
        Exit Sub
      End If
   End If
   '
   IMAGSTRI$ = String$(508 * (1 + LastRecord& - FirstRecord&), 0)
   PINSERT& = 1
   For U& = FirstRecord& To LastRecord&
      XX$ = REGLEIDO$("PIMGSPL", U&)
      If CVS(Left$(XX$, 4)) <> IMAGENUMBER& Then
        Call MENSERR(24, "Imagen no Encontrada - Imposible Incluir.")
        Exit Sub
      End If
      Mid$(IMAGSTRI$, PINSERT&, 508) = Mid$(XX$, 5, 508)
      PINSERT& = PINSERT& + 508
   Next U&
   '
   If Len(IMAGSTRI$) < LARGIMAG& Then
      Call MENSERR(24, "Imagen no Encontrada - Imposible Incluir.")
      Exit Sub
   End If
   IMAGSTRI$ = Left$(IMAGSTRI$, LARGIMAG&)
   '
   NX% = FreeFile
   Open LUCOM$ + NOMBIMAG$ For Binary As #NX%
   Put #NX%, , IMAGSTRI$
   Close #NX%
   '
   FORMALTER.Image3.Picture = LoadPicture("")
   FORMALTER.Image3.Picture = LoadPicture(LUCOM$ + NOMBIMAG$)
   '
   HIMAGEN = FORMALTER.Image3.Picture.Width
   VIMAGEN = FORMALTER.Image3.Picture.Height
   '
   If MODOPRE$ = "PRINT" Then
       Printer.PaintPicture FORMALTER.Image3.Picture, PHX1, PVX1, ACHX, ALTX
     ElseIf MODOPRE$ = "SHOW" Then
       EDITFORM.Picture1(PAGINACTIVA%).PaintPicture FORMALTER.Image3.Picture, PHX1, PVX1 + TOPEPAGINA, ACHX, ALTX
   End If
   '
End Sub
'
Sub FORMMGR(Form$, RG%, CLII%, VLX$, cmd$)
   '
   ' FORM$  =  Nombre de formulario
   ' RG%    =  Renglon
   ' CLII%    =  Columna
   ' VLX$   =  Valor o texto a colocar
   ' CMD$   =  Comando a ejecutar
   '
   Static NX%, NY%, ATOT%
   Static FXXY As String * 128
   Static FYY As String * 256
   '
   If REFRESHFORM% > 0 Then    ' se ejecuto BAJALDISCO
      NX% = 0
      NY% = 0
      FORMACT$ = ""
      REFRESHFORM% = 0
   End If
   If cmd$ = "CLOSE" Then
      '
      If NX% > 0 Then
         Close #NX%: NX% = 0
      End If
      If NY% > 0 Then
         Close #NY%: NY% = 0
      End If
      FORMACT$ = ""
      Exit Sub
      '
   End If
   '
   If Form$ <> "" Then
      If UCase$(TRIM$(Form$)) <> FORMACT$ Then
         FORMACT$ = UCase$(TRIM$(Form$))
         FORMUDRV$ = LUDAT$ + FORMACT$ + ".FRM"
         If NX% > 0 Then Close #NX%: NX% = 0
         NX% = FILEOPEN%(FORMUDRV$, 0, 128)
         '
         ATOT% = 256: CARE% = 128       ' ancho total y cantidad de renglones
         For k% = 1 To LOF(NX%) / 128
            Get #NX%, k%, FXXY$
            If Left$(FXXY$, 8) = "CAN-COL " Then
                 ATOT% = XVALO(TRIM$(Mid$(FXXY$, 89, 8)))
              ElseIf Left$(FXXY$, 8) = "CAN-REN " Then
                 CARE% = XVALO(TRIM$(Mid$(FXXY$, 89, 8)))
            End If
         Next k%
      End If
   End If
   '
   COMA$ = UCase$(TRIM$(cmd$))
   '
   If Left$(COMA$, 5) = "CLEAR" Then           ' genera formulario en blanco
      LUX$ = LUCOM$
      If NY% > 0 Then Close #NY%: NY% = 0
      NY% = FILEOPEN%(LUX$ + FORMACT$, 0, 256)
      '
      For i% = 1 To CARE%
         LSet FYY$ = Space$(256)
         Put #NY%, i%, FYY
      Next i%
      '
      Exit Sub
      '
   ElseIf Left$(COMA$, 5) = "PRINT" Then
      Screen.MousePointer = 11
      UTI$ = LUBAS$
      If Mid$(UTI$, 2, 1) <> ":" Then
          UTI$ = Left$(CurDir, 2) + UTI$
      End If
      LUX$ = LUCOM$
      ATRIBU$ = "/" + LUX$ + FORMACT$
      ATRIBU$ = ATRIBU$ + "/CAREN=" + TRIM$(str$(CARE%))
      ATRIBU$ = ATRIBU$ + "/SETUP=" + SETUPSTRI$
      ATRIBU$ = ATRIBU$ + "/MARSU=" + TRIM$(str$(MARSUP%))
      '
      NZ% = FILEOPEN%(LUCOM$ + "COMBAT.$$$", 2, 0)
      Print #NZ%, "FINBAT"
      Close #NZ%: NZ% = 0
      '
      On Error Resume Next
      Kill LUCOM$ + "FINBAT.$$$"
      On Error GoTo 0
      '
      NZ% = FILEOPEN%(LUCOM$ + "DOSCOMM.BAT", 2, 0)
      Print #NZ%, UTI$ + "PRINTFRM.EXE " + ATRIBU$
      Close #NZ%: NZ% = 0
      '
      Shell LUCOM$ + "DOSCOMM.PIF"
      TINI = Timer
      MENSA$ = "Se ha producido un error no recuperable" + Chr$(10) + Chr$(13)
      MENSA$ = MENSA$ + "en la impresion de este documento." + Chr$(10) + Chr$(13) + Chr$(10) + Chr$(13)
      MENSA$ = MENSA$ + "Consulte a su Soporte de Sistemas FLEXOFT"
      '
149   If Timer > TINI + 30 Then
          Screen.MousePointer = 1
          MsgBox MENSA$, 4096 + 48, "Control de Impresion"
          On Error Resume Next
          Kill LUCOM$ + "COMBAT.$$$"
          Kill LUCOM$ + "FINBAT.$$$"
          On Error GoTo 0
          ABOPRIN% = 1
          Exit Sub
      End If
      '
      If EXISTE%(LUCOM$ + "FINBAT.$$$") < 1 Then GoTo 149
      Screen.MousePointer = 1
      ABOPRIN% = 0
      Exit Sub
      '
   ElseIf Left$(COMA$, 4) = "FSET" Then
      If RG% > 0 Then
         If RG% <= CARE% Then
            If CLII% > 0 Then
               If CLII% <= ATOT% Then
                  Get #NY%, RG%, FYY$
                  TXTT$ = FYY$
                  Mid$(TXTT$, CLII%, Len(VLX$)) = VLX$
                  LSet FYY$ = TXTT$
                  Put #NY%, RG%, FYY$
               End If
            End If
         End If
      End If
      Exit Sub
      '
   ElseIf Left$(COMA$, 5) = "LEYEN" Then
      For k% = 1 To LOF(NX%) / 128
        Get #NX%, k%, FXXY$
        If Left$(FXXY$, 4) = "LEYE" Then
          VLX$ = TRIM$(Mid$(FXXY$, 9, 80))
          If VLX$ <> "" Then
            For t% = 97 To 128 Step 8
               RG% = Asc(Mid$(FXXY$, t%, 1))
               If RG% > 0 Then
                  If RG% <= CARE% Then
                     CLII% = Asc(Mid$(FXXY$, t% + 1, 1))
                     If CLII% > 0 Then
                        If CLII% <= ATOT% Then
                           Formato$ = TRIM$(Mid$(FXXY$, t% + 2, 6))
                           If Formato$ <> "" Then
                              TX$ = TEXTFORM$(VLX$, Formato$)
                              Get #NY%, RG%, FYY$
                              TXTT$ = FYY$
                              Mid$(TXTT$, CLII%, Len(TX$)) = TX$
                              FYY$ = TXTT$
                              Put #NY%, RG%, FYY$
                           End If
                        End If
                     End If
                  End If
               End If
            Next t%
          End If
        End If
      Next k%
      '
      Exit Sub
      '
   ElseIf Left$(COMA$, 1) = "@" Then
      CD$ = AJUSTI$(Mid$(COMA$, 2), 8)
      For k% = 1 To LOF(NX%) / 128
         Get #NX%, k%, FXXY$
         If Left$(FXXY$, 8) = CD$ Then
            RG0% = RG%
            UREGUT% = 0
            For t% = 97 To 128 Step 8
               RG% = RG0%
               RGX% = Asc(Mid$(FXXY$, t%, 1))
               If RG% + RGX% > 0 Then
                  If RG% + RGX% <= CARE% Then
                     CLII% = Asc(Mid$(FXXY$, t% + 1, 1))
                     If CLII% > 0 Then
                        If CLII% <= ATOT% Then
                           Formato$ = TRIM$(Mid$(FXXY$, t% + 2, 6))
                           If Formato$ <> "" Then
                              ILETRA% = 0
                              '
                              If CD$ = "IMP-LET1" Then
                                 ILETRA% = 1
                                 VLX1$ = VLX$
                                 LOTEX% = XVALO(Mid$(Formato$, 2, 4))
                                 Call FRATEX(VLX$, NRX1$, NRX2$, NRX3$, LOTEX%)
                                 VLX$ = NRX1$ + String$(LOTEX%, "*")
                                 RL% = 1
                              End If
                              '
                              If CD$ = "ORD-COPI" Then
                                 ORDECOP% = ORDECOP% + 1
                                 VLX$ = ""
                                 If ORDECOP% <= 10 Then
                                    VLX$ = ORCOP$(ORDECOP%)
                                 End If
                              End If
                              '
1000                          TX$ = TEXTFORM$(VLX$, Formato$)
                              RG% = RG% + RGX%
                              Get #NY%, RG%, FYY$
                              TXTT$ = FYY$
                              Mid$(TXTT$, CLII%, Len(TX$)) = TX$
                              LSet FYY$ = TXTT$
                              Put #NY%, RG%, FYY$
                              If RG% > UREGUT% Then UREGUT% = RG%
                              '
                              If ILETRA% = 1 Then
                                 If RL% = 1 Then
                                     RL% = 2
                                     VLX$ = NRX2$ + String$(LOTEX%, "*")
                                     RGX% = 1
                                     GoTo 1000
                                   ElseIf RL% = 2 Then
                                     RL% = 3
                                     VLX$ = NRX3$ + String$(LOTEX%, "*")
                                     RGX% = 1
                                     GoTo 1000
                                 End If
                                 VLX$ = VLX1$
                              End If
                           End If
                        End If
                     End If
                  End If
               End If
            Next t%
            If UREGUT% > RG% Then RG% = UREGUT%
            Exit Sub
         End If
      Next k%
      Exit Sub
      '
   ElseIf Left$(COMA$, 1) = "?" Then
      CD$ = TRIM$(Mid$(COMA$, 2))
      PPXY% = InStr(CD$, "/"): If PPXY% < 1 Then PPXY% = 1 + Len(CD$)
      CD2$ = TRIM$(Mid$(CD$, 1 + PPXY%))
      CD$ = AJUSTI$(Left$(CD$, PPXY% - 1), 8)
      '
      If NX% > 0 Then
         For k% = 1 To LOF(NX%) / 128
            Get #NX%, k%, FXXY$
            If Left$(FXXY$, 8) = CD$ Then
               VLX$ = TRIM$(Mid$(FXXY$, 89, 8))
               If CD2$ = "POSINHOR" Then
                   VLX$ = TRIM$(str$(CVS(Mid$(FXXY$, 57, 4))))
                 ElseIf CD2$ = "POSINVER" Then
                   VLX$ = TRIM$(str$(CVS(Mid$(FXXY$, 61, 4))))
                 ElseIf CD2$ = "FORMATO" Then
                   VLX$ = TRIM$(Mid$(FXXY$, 75, 6))
                 ElseIf CD2$ = "FUENTE" Then
                   VLX$ = TRIM$(Mid$(FXXY$, 87, 32))
                 ElseIf CD2$ = "TAFUEN" Then
                   VLX$ = TRIM$(str$(CVS(Mid$(FXXY$, 119, 4))))
               End If
               Exit Sub
            End If
         Next k%
      End If
      VLX$ = ""
      Exit Sub
      '
   Else
      '
      Call COMUNI("Comando '" + COMA$ + "' no Procesable\en Rutina 'FORMMGR'\para Formulario '" + FORMACT$ + "'.")
      '
   End If
   '
999 End Sub
   '
'*****************************************************************************
   '
Sub SETFORM(Form$, VLX$, CTRX$)
   '
   cmd$ = "@" + TRIM$(CTRX$)
   RG% = 0: CLII% = 0
   Call FORMMGR(Form$, RG%, CLII%, VLX$, cmd$)
   ' sobre RG% y CLII% devuelve renglon y columna utilizado
   '
End Sub
'
'*****************************************************************************
'
Sub CARLEYEN(Form$)
   '
   cmd$ = "LEYEN"
   RG% = 0: CLII% = 0
   Call FORMMGR(Form$, RG%, CLII%, VLX$, cmd$)
   '
End Sub
'
'*****************************************************************************
'
Sub APPFORM(VLX$, CTRX$)
   '
   TEXT$ = VLX$
   CODVARI$ = CTRX$
   X0 = 0
   Y0 = YAC
   If TRIM$(TEXT$) <> "" Then
       Call PRINTFORWIN(TEXT$, CODVARI$, X0, Y0)
       YACNUE = Printer.CurrentY / 56.7 - MARSUW
       If YACNUE > YAC Then YAC = YACNUE
   End If
   '
End Sub
'
'*****************************************************************************
'
Sub CLEARFORMPRIN(Form$)
   '
   Call FORMMGR(Form$, RG%, CLII%, VLX$, "CLEAR")
   '
End Sub
'
'*****************************************************************************
'
Sub CLOSEFORMPRIN(Form$)
   '
   Call FORMMGR(Form$, RG%, CLII%, VLX$, "CLOSE")
   '
End Sub
'
'*****************************************************************************
'
Sub PRINTFORM(Form$)
   '
   Call FORMMGR(Form$, RG%, CLII%, VLX$, "PRINT")
   '
End Sub
'
'
'*****************************************************************************
'
Sub CARSETUP(AAA$)
   '
   If AAA$ = "NEW" Then
        SETUPSTRI$ = ""
      Else
        For UU% = 1 To Len(AAA$)
           UV$ = TRIM$(str$(Asc(Mid$(AAA$, UU%, 1))))
           While Len(UV$) < 3
              UV$ = "0" + UV$
           Wend
           SETUPSTRI$ = SETUPSTRI$ + UV$
        Next UU%
   End If
   '
End Sub
'
'*****************************************************************************
'
Function ATRIFORM$(FORMUL$, CTRX$)
   '
   If EXISTE%(LUDAT$ + FORMUL$ + ".FRM") > 0 Then
        FORMOX$ = FORMUL$
      Else
        'If TRIM$(FORMUL$) = "" Then Call GRACONTROL("", "", "") SE COMENTO PARA VER SI QUITABA EL ERROR EN EL ACTUREGISTRO
        FORMOX$ = TRIM$(CONTROL$("", FORMUL$))                    'DEBENIDO DEL GRACONTROL
   End If
   '
   cmd$ = "?" + TRIM$(CTRX$)
   Call FORMMGR(FORMOX$, RG%, CLII%, VLX$, cmd$)
   ATRIFORM$ = VLX$
   '
End Function
'
'*****************************************************************************
'
Sub PRINTDOC(DOCTYPE$, Optional MODO$, Optional RET_DOCUNU&)
   '
   RET_DOCUNU& = 0
   Dim ORCOP$(9)
   ORCOP$(1) = "ORIGINAL"
   ORCOP$(2) = "DUPLICADO"
   ORCOP$(3) = "TRIPLICADO"
   ORCOP$(4) = "CUADRUPLICADO"
   ORCOP$(5) = "QUINTUPLICADO"
   ORCOP$(6) = "SEXTUPLICADO"
   '
   If DOCTYPE$ = "" Then
      Call MENSERR(24, "Error de Configuración:\  \Falta Definir Tipo de Documento 'DOCTYPE'.")
      Screen.MousePointer = 1
      Exit Sub
   End If
   '
   If Left$(DOCTYPE$, 1) = "@" Then
        FORIPRE$ = TRIM$(Mid$(DOCTYPE$, 2))
      Else
        FORIPRE$ = TRIM$(CONTROL$("", DOCTYPE$))
   End If
   '
   If FORIPRE$ = "" Then
        Call COMUNI("Error de Configuración:\  \Falta Definir Formulario de Impresión\Correspondiente a Documento " + TRIM$(DOCTYPE$) + ".")
        Screen.MousePointer = 1
        Exit Sub
      Else
        If EXISTE%(LUDAT$ + FORIPRE$ + ".FRM") <> 1 Then
           RMCC$ = REPLACAR$(LUDAT$ + FORIPRE$ + ".FRM", "\", "/")
           Call COMUNI("Error de Configuración:\  \No Existe Formulario de Impresión\Correspondiente a " + TRIM$(DOCTYPE$) + ":\  \'" + RMCC$ + "'.")
           Screen.MousePointer = 1
           Exit Sub
        End If
   End If
   '
   If App.EXEName = "FACTUMAI" Then
      Screen.MousePointer = 1
      COLOFORM.Label1.Caption = NUMEDOC$
      COLOFORM.Label2.Caption = TIPODOC$
      COLOFORM.Show 1
   End If
   '
   Screen.MousePointer = 11   ' Generando Documento
   '
   If TRIM$(PRINTERPORT$) <> "" Then
      Call SETPRINPORT(PRINTERPORT$)
   End If
   '
   CODFOR$ = TRIM$(UCase$(FORIPRE$))
   MARIZW = XVALO(ATRIFORM$(CODFOR$, "MARIZQMM"))
   MARSUW = XVALO(ATRIFORM$(CODFOR$, "MARSUPMM"))
   '
   LARHOW = XVALO(ATRIFORM$(CODFOR$, "LARGOJMM"))
   If LARHOW < 80 Then LARHOW = 25.4 * 12
   '
   INTERX$ = TRIM$(ATRIFORM$(CODFOR$, "INTERLIN"))
   '
   MULTIREN% = 1
   PPXX% = InStr(INTERX$, "/")
   If PPXX% > 0 Then
      MULTIREN% = XVALO(Mid$(INTERX$, PPXX% + 1))
      INTERX$ = Left$(INTERX$, PPXX% - 1)
   End If
   If MULTIREN% < 1 Then MULTIREN% = 1
   INTERLI = XVALO(INTERX$)
   If INTERLI < 1 Then
      'Call COMUNI("Error de Configuración de Interlinea\en Formulario '" + CODFOR$ + ".FRM'.")
      INTERLI = 4
   End If
   '
   DETDESMM = XVALO(ATRIFORM$(CODFOR$, "DETDESMM"))
   DETHASMM = XVALO(ATRIFORM$(CODFOR$, "DETHASMM"))
   If DETHASMM < DETDESMM + 4 + INTERLI Then
      DETHASMM = DETDESMM + 4 * INTERLI
   End If
   '
   OPCIM$ = TRIM$(UCase$(ATRIFORM$(CODFOR$, "OPCIMPRE")))
   If OPCIM$ = "SI" Then
     CANCELPRINT% = 0
     If COMALTER%("Opciones de Impresión:\\Imprime Comprobante\Suprime Impresión") = 2 Then
       CANCELPRINT% = 1
     End If
   End If
   '
   'CANHOJ& = 1
   'If CAITAB1% > 0 Then
   '   On Error Resume Next
   '   CANHOJ& = 1 + Int(MULTIREN% * CAITAB1% / ((DETHASMM - DETDESMM) / INTERLI))
   '   On Error GoTo 0
   'End If
   'If CANHOJ& < 1 Then CANHOJ& = 1
   '
   CACOPX$ = TRIM$(ATRIFORM$(CODFOR$, "CAN-COP"))
   If Left$(CACOPX$, 1) = "#" Then
5      CACOPI% = XVALO(InputBox$("Cantidad de Copias", "Ingrese Cantidad de Copias", Mid$(CACOPX$, 2)))
       If CACOPI% < 1 Or CACOPI% > 9 Then GoTo 5
     Else
       CACOPI% = XVALO(CACOPX$)
       If CACOPDOC% > 0 Then CACOPI% = CACOPDOC%
       If CACOPI% < 1 Or CACOPI% > 9 Then
         'Call COMUNI("Error de Configuración de Cantidad de Copias\en Formulario '" + CODFOR$ + ".FRM'.")
         CACOPI% = 1
       End If
   End If
   '
   If CACOPI% > 6 Then
     For KKII% = 2 To 9
       ORCOP$(KKII%) = "COPIA"
     Next KKII%
   End If
   '
   ORIENTA% = 1
   PORI$ = TRIM$(UCase$(ATRIFORM$(CODFOR$, "ORIENTA")))
   If PORI$ = "PAISAJE" Or PORI$ = "2" Then
      ORIENTA% = 2
   End If
   '
   PORIAN = Printer.Orientation  ' retiene orientacion
   'Printer.Orientation = 1
   On Error Resume Next
   Printer.ScaleMode = 1
   '
   Call SETSPOOL("SCALEMODE", "1")
   Call SETSPOOL("ORIENTATION", "1")
   Printer.Orientation = 1
   LARHOAN = Printer.Height
   If LARHOAN < 10 Then LARHOAN = 12 * 25.4 * 56.7
   ANCHOAN = Printer.Width
   If ANCHOAN < 10 Then ANCHOAN = 203 * 56.7
   On Error GoTo 0
   '
10 'Printer.Orientation = 1
   If ORIENTA% = 2 Then
        'Printer.Orientation = 2
        Call SETSPOOL("ORIENTATION", "2")
        'Call SETSPOOL("HEIGHT", Str$(ANCHOAN))
        'Call SETSPOOL("WIDTH", Str$(LARHOW * 56.7))
        'Printer.Width = ANCHOAN
      Else
        Call SETSPOOL("ORIENTATION", "1")
        'Call SETSPOOL("HEIGHT", Str$(LARHOW * 56.7))
        'Call SETSPOOL("WIDTH", Str$(ANCHOAN))
   End If
   '
   For COP% = 1 To CACOPI%
      '
      INIDET& = 1
      NUHOJ% = 0: NUHOK% = 0
      CANHOJ& = 1
      '
20    Call CARFORWIN         ' Carga Formulario en Blanco
      Call LIMPIAVALPAR
      '
      Call PRINTFORWIN(TIPODOC$, "TIPO-COM", 0, 0)
      Call PRINTFORWIN(NUMEDOC$, "NUME-COM", 0, 0)
      Call PRINTFORWIN(DOCUORIG$, "DOCUORIG", 0, 0)
      Call PRINTFORWIN(FECHDOC$, "FECH-EMI", 0, 0)
        AHORAX$ = REPLACE$(Format$(Now, "hh:mm:ss     dd/mm/yyyy"), "del", 10, 3)
      Call PRINTFORWIN(AHORAX$, "HORA-EMI", 0, 0)
      Call PRINTFORWIN(USERNAME$, "USUARIO", 0, 0)
      '
      Call PRINTFORWIN(ORCOP$(COP%), "ORD-COPI", 0, 0)
      '
      If Abs(DESTIDOC%) > 0 Then DESTIDOCNUE = DESTIDOC%
      '
      If Abs(DESTIDOCNUE) > 0 Then
         Call PRINTFORWIN(rasocli$(DESTIDOCNUE), "RASO-CLI", 0, 0)
         Call PRINTFORWIN(DOMICLI$(DESTIDOCNUE), "DOMI-CLI", 0, 0)
           ATRIX$ = TRIM$(ATRIFORM$(CODFOR$, "CPOS-CLI/FORMATO"))
           If ATRIX$ <> "" Then
               'IMPRIME SEPARADO COD-POSTAL Y LOCALIDAD
               Call PRINTFORWIN(CPOSCLI$(DESTIDOCNUE), "CPOS-CLI", 0, 0)
               Call PRINTFORWIN(LOCACLI$(DESTIDOCNUE), "LOCA-CLI", 0, 0)
             Else
               'iMPRIME CODIGO POSTAL Y LOCALIDAD TODO JUNTO
               Call PRINTFORWIN(CPLOCLI$(DESTIDOCNUE), "LOCA-CLI", 0, 0)
           End If
         Call PRINTFORWIN(DEPROVI$(PROVCLI$(DESTIDOCNUE)), "PCIA-CLI", 0, 0)
         Call PRINTFORWIN(PAISCLI$(DESTIDOCNUE), "PAIS-CLI", 0, 0)
         Call PRINTFORWIN(TEL1CLI$(DESTIDOCNUE), "TELE-CLI", 0, 0)
         Call PRINTFORWIN(FAXCLI$(DESTIDOCNUE), "FAX-CLI", 0, 0)
         Call PRINTFORWIN(EMAILCLI$(DESTIDOCNUE), "MAIL-CLI", 0, 0)
         Call PRINTFORWIN(CuitCli$(DESTIDOCNUE), "CUIT-CLI", 0, 0)
         Call PRINTFORWIN(NIBRCLI$(DESTIDOCNUE), "IBRU-CLI", 0, 0)
         Call PRINTFORWIN(POSIVAL$(DESTIDOCNUE), "PIVA-CLI", 0, 0)
         Call PRINTFORWIN(NPROCLI$(DESTIDOCNUE), "PROV-CLI", 0, 0)
           FORCUCLI$ = TRIM$(UCase$(ATRIFORM$(CODFOR$, "CUEN-CLI/FORMATO")))
           DEDOC$ = str$(Abs(DESTIDOCNUE))
           If Left$(FORCUCLI$, 1) = "A" Then
             DEDOC1$ = TRIM$(CODICLI$(DESTIDOCNUE))
             If DEDOC1$ <> "" Then DEDOC$ = DEDOC1$
           End If
         Call PRINTFORWIN(DEDOC$, "CUEN-CLI", 0, 0)
         '
         If DESTIDOCNUE > 0 Then
            Call PRINTFORWIN(NTRANCLI$(DESTIDOCNUE), "NTRA-CLI", 0, 0)
            Call PRINTFORWIN(RTRANCLI$(DESTIDOCNUE), "RTRA-CLI", 0, 0)
            Call PRINTFORWIN(DTRANCLI$(DESTIDOCNUE), "DTRA-CLI", 0, 0)
            Call PRINTFORWIN(LTRANCLI$(DESTIDOCNUE), "LTRA-CLI", 0, 0)
         End If
      End If
      '
      ' MONEDA DE EMISION
      Call PRINTFORWIN(DEMONECO$(MONEMI%), "MONEDEMI", 0, 0)
      '
      CIACTI% = 0
      For KKI% = 1 To CAPARDOC%          '
         Call PRINTFORWIN(DOCPARAM$(KKI%), CODPARAM$(KKI%), 0, 0)
         ' ESTO ES PARA RETENER EL VALOR PARA LAS IMAGENES
         If UCase$(CODPARAM$(KKI%)) = "COD-MAT" Then CIACTI% = CODINT%(DOCPARAM$(KKI%))
      Next KKI%
      '
      ' AQUI AGREGAR LOS PARAMETROS TOMADOS DE LA BASE DE DATOS
      XX$ = LOADFILE$(LUDAT$ + CODFOR$ + ".FRQ") 'ASIGNO A XX$ LA LECTURA DEL ARCHIVO
      For U% = 1 To Len(XX$) Step 256 'LO RECORRO (256 ES LA LONGITUD DEL REGISTRO)
        YY$ = Mid$(XX$, U%, 256)
        CODVARI$ = TRIM$(Left(YY$, 16))
        TABLASE$ = TRIM$(Mid$(YY$, 17, 24))
        If TABLASE$ <> "" Then
           CAMPOSE$ = TRIM$(Mid$(YY$, 41, 24))
           If CAMPOSE$ <> "" Then
             CONDI$ = TRIM$(Mid$(YY$, 65))
31           TBUS$ = "¿"
             PPXX1% = InStr(CONDI$, TBUS$)
             If PPXX1% > 0 Then
               TBUS$ = "?"
               PPXX2% = InStr(CONDI$, TBUS$)
               IZQUI$ = Left$(CONDI$, PPXX1% - 1)
               DERE$ = Mid$(CONDI$, PPXX2% + 1)
               AREPLA$ = ""
               If PPXX2% > PPXX1% + 1 Then
                  AREPLA$ = Mid$(CONDI$, PPXX1% + 1, PPXX2% - PPXX1% - 1)
               End If
               If AREPLA$ <> "" Then
                  For KKII% = 1 To CAVALPAR%
                     If CODPAR$(KKII%) = AREPLA$ Then
                        AREPLA$ = VALPAR$(KKII%)
                        GoTo 32
                     End If
                  Next KKII%
               End If
32             CONDI$ = IZQUI$ + AREPLA$ + DERE$
               GoTo 31
             End If
             '
             TTXX$ = SAFETEXT$(VALOBADA(TABLASE$, CAMPOSE$, CONDI$))
             Call PRINTFORWIN$(TTXX$, CODVARI$, 0, 0)
             ' ESTO ES PARA RETENER EL VALOR PARA LAS IMAGENES
             If UCase$(CODVARI$) = "COD-MAT" Then CIACTI% = CODINT%(TTXX$)
           End If
        End If
      Next U%
      '
      For KKI% = 1 To CAIMINI% + CAIMAGEN%
         CODIMA$ = UCase$(CODIMAGEN$(KKI%))
         If Left$(CODIMA$, 7) = "DOCUGRA" Then
           ' ESTA PARTE ES PARA INCRUSTAR LOS 4 DOCUMENTOS GRAFICOS
           If CIACTI% > 0 Then
             NIMA% = XVALO(Mid$(CODIMA$, 8, 1))
             If NIMA% <= 3 Then
               FGRC$ = FILTERGETRECORD$("INDIPLAN", 1, MKI$(CIACTI%))
               ARIMA$ = TRIM$(Mid$(FGRC$, 65 + 48 * NIMA%, 48))
                 RUTARGRA$ = TRIM$(CONTROL$("", "RUTARGRA"))
                 If RUTARGRA$ = "" Then RUTARGRA$ = TRIM$(LUDAT$)
                 If Right$(RUTARGRA$, 1) <> "\" Then RUTARGRA$ = RUTARGRA$ + "\"
               ARCHIMAGEN$(KKI%) = RUTARGRA$ + ARIMA$
             End If
           End If
         End If
         Call PRINTIMAGE(CODIMAGEN$(KKI%), ARCHIMAGEN$(KKI%))
      Next KKI%
      '
      YAC = 0
      For YX& = 1 To CAITAB2%
         For KKI% = 1 To CACOLTAB2%
            TTTXXX$ = TETABU2$(KKI%, YX&)
            Call PRINTFORWIN(TTTXXX$, CODTABU2$(KKI%), 0, YAC)
         Next KKI%
         YAC = YAC + INTERLI
      Next YX&
      '
      YAC = 0: INTEROBS = INTERLI
      If INTERLI > 6 Then INTEROBS = INTERLI / 2
      For UKU& = 1 To ULTREG&("!OBSERVOF")
         TXOB$ = TRIM$(REGLEIDO$("!OBSERVOF", UKU&))
         Call PRINTFORWIN(TXOB$, "OBSERVA", 0, YAC)
         YAC = YAC + INTEROBS
      Next UKU&
      Call CIERRARCH("!OBSERVOF")
      '
      YAC = XVALO(ATRIFORM$(CODFOR$, "DETDESMM"))
      '
      For YX& = INIDET& To CAITAB1%
         '
         If YAC > DETHASMM Then
            '
            NUHOJ% = NUHOJ% + 1
            If NUHOJ% = 1 Then
               On Error Resume Next
               CANHOJ& = 1 + Int((CAITAB1% / (YX& - 1)) - 0.031) ' CAMBIE 0.05 POR 0.031
               'se dio el caso que con 193 items imprimia 6 copias de 7 caso en ahp
               On Error GoTo 0
            End If
            Call PRINTFORWIN(TRIM$(str$(NUHOJ%)), "HOJA-NUM", 0, 0)
            Call PRINTFORWIN(TRIM$(str$(CANHOJ&)), "HOJA-TOT", 0, 0)
            '
            INIDET& = YX&
            'Printer.NewPage
            Call SETSPOOL("NEWPAGE", "")
            GoTo 20
         End If
         '
         If TETABU1$(1, YX&) = String$(32, "-") Then
            Call PRILINSEP
            GoTo 89
         End If
         '
         YACMAX = 0
         For KKI% = 1 To CACOLTAB1%
            TTTXXX$ = TETABU1$(KKI%, YX&)
            If Left$(TTTXXX$, 1) = "@" Then
               YAC = YAC - INTERLI
               TTTXXX$ = Mid$(TTTXXX$, 2)
            End If
            Call PRINTFORWIN(TTTXXX$, CODTABU1$(KKI%), 0, YAC)
            On Error Resume Next
            If Printer.CurrentY / 56.7 - MARSUW + INTERLI > YACMAX Then
               YACMAX = Printer.CurrentY / 56.7 - MARSUW + INTERLI
            End If
            '
            On Error GoTo 0
         Next KKI%
         '
         If YACMAX < YAC Then YACMAX = YAC + INTERLI
         YAC = YACMAX

89    Next YX&
      '
      NUHOJ% = NUHOJ% + 1
      If NUHOJ% = 1 Then
         On Error Resume Next
         'CANHOJ& = 1 + Int((CAITAB1% / (YX& - 1)) - 0.05)
         CANHOJ& = 1 + Int((CAITAB1% / (YX& - 1)) - 0.031) ' CAMBIE 0.05 POR 0.031
         On Error GoTo 0
      End If
      Call PRINTFORWIN(TRIM$(str$(NUHOJ%)), "HOJA-NUM", 0, 0)
      Call PRINTFORWIN(TRIM$(str$(CANHOJ&)), "HOJA-TOT", 0, 0)
      '
      If COP% < CACOPI% Then
         'Printer.NewPage
         Call SETSPOOL("NEWPAGE", "")
      End If
      '
   Next COP%
   '
   'Printer.EndDoc
99 Call SETSPOOL("ENDDOC", "")
   '
   If Len(TIPODOC$) >= 60 Then
      DESCRIDOC$ = TRIM$(TIPODOC$)
      GoTo 98
   End If
   '
   DESCRIDOC$ = TRIM$(TIPODOC$) + " Nro." + TRIM$(NUMEDOC$)
   If Len(TIPODOC$) > 8 Then
     DESCRIDOC$ = TRIM$(TRIM$(TIPODOC$) + " " + NUMEDOC$)
   End If
   DEDOCOR$ = REPLACAR$(DESCRIDOC$, " ", "_")
   If Abs(DESTIDOCNUE) > 0 Then
      DESCRIDOC$ = DESCRIDOC$ + " - " + TRIM$(rasocli$(DESTIDOCNUE))
   End If
   If TRIM$(FECHDOC$) <> "" Then
      DESCRIDOC$ = DESCRIDOC$ + " - " + FECHDOC$
   End If
   '
98 Call SAVESPOOL(DESCRIDOC$, OKX%, RET_DOCUNU&)
   '
   Call CIERRARCH("PDOCLST")
   Call CIERRARCH("PDOCSPL")
   Call CIERRARCH("PIMGLST")
   Call CIERRARCH("PIMGSPL")
   '
   Screen.MousePointer = 1
   If NOCONFIRM% < 1 Then
     If CANCELPRINT% < 1 Then
       If TRIM$(NUMEDOC$) <> "" Then
         If OKX% < 1 Then
           Call MENSERR(24, "Problemas de Impresión o de Generación de Archivo (1).")
           GoTo 49
         End If
         TTXYZ$ = "Se ha Emitido por Impresora"
         If SENDBYMAIL$ <> "" Then TTXYZ$ = "Se ha Emitido en Formato PDF para Envio por Mail"
         If MATRIXPUN% = 1 Then TTXYZ$ = TTXYZ$ + " (MP)"
         TTXYZ$ = TTXYZ$ + "\" + TIPODOC$ + " Numero '" + NUMEDOC$ + "'"
         If COMALTER%(TTXYZ$ + "\\Repetir Impresion\Continuar") = 1 Then GoTo 10
       End If
     End If
   End If
   '
49 On Error Resume Next
   Printer.Orientation = 1
   'Call SETSPOOL("ORIENTATION", "1")
   'Printer.Height = LARHOAN
   'Call SETSPOOL("HEIGHT", Str$(LARHOAN))
   'Printer.Width = ANCHOAN
   'Call SETSPOOL("WIDTH", Str$(ANCHOAN))
   Printer.Orientation = PORIAN
   'Call SETSPOOL("ORIENTATION", Str$(PORIAN))
   '
   If TRIM$(PRINTERPORT$) <> "" Then
      Call SETPRINPORT("RESTORE")
   End If
   '
   If InStr(UCase$(MODO$), "NOCANOBS") < 1 Then
      Call SETULTREG("!OBSERVOF", 0)
   End If
   Call CIERRARCH("!OBSERVOF")
   '
End Sub

Sub PRINTDOCNUE(DOCTYPE$, Optional MODO$)
   '
   Dim ORCOP$(9)
   ORCOP$(1) = "ORIGINAL"
   ORCOP$(2) = "DUPLICADO"
   ORCOP$(3) = "TRIPLICADO"
   ORCOP$(4) = "CUADRUPLICADO"
   ORCOP$(5) = "QUINTUPLICADO"
   ORCOP$(6) = "SEXTUPLICADO"
   '
   If DOCTYPE$ = "" Then
      Call MENSERR(24, "Error de Configuración:\  \Falta Definir Tipo de Documento 'DOCTYPE'.")
      Screen.MousePointer = 1
      Exit Sub
   End If
   '
   If Left$(DOCTYPE$, 1) = "@" Then
        FORIPRE$ = TRIM$(Mid$(DOCTYPE$, 2))
      Else
        FORIPRE$ = TRIM$(CONTROL$("", DOCTYPE$))
   End If
   '
   If FORIPRE$ = "" Then
        Call COMUNI("Error de Configuración:\  \Falta Definir Formulario de Impresión\Correspondiente a Documento " + TRIM$(DOCTYPE$) + ".")
        Screen.MousePointer = 1
        Exit Sub
      Else
        If EXISTE%(LUDAT$ + FORIPRE$ + ".FRM") <> 1 Then
           RMCC$ = REPLACAR$(LUDAT$ + FORIPRE$ + ".FRM", "\", "/")
           Call COMUNI("Error de Configuración:\  \No Existe Formulario de Impresión\Correspondiente a " + TRIM$(DOCTYPE$) + ":\  \'" + RMCC$ + "'.")
           Screen.MousePointer = 1
           Exit Sub
        End If
   End If
   '
   If App.EXEName = "FACTUMAI" Then
      Screen.MousePointer = 1
      COLOFORM.Label1.Caption = NUMEDOC$
      COLOFORM.Label2.Caption = TIPODOC$
      COLOFORM.Show 1
   End If
   '
   Screen.MousePointer = 11   ' Generando Documento
   '
   If TRIM$(PRINTERPORT$) <> "" Then
      Call SETPRINPORT(PRINTERPORT$)
   End If
   '
   CODFOR$ = TRIM$(UCase$(FORIPRE$))
   MARIZW = XVALO(ATRIFORM$(CODFOR$, "MARIZQMM"))
   MARSUW = XVALO(ATRIFORM$(CODFOR$, "MARSUPMM"))
   '
   LARHOW = XVALO(ATRIFORM$(CODFOR$, "LARGOJMM"))
   If LARHOW < 80 Then LARHOW = 25.4 * 12
   '
   INTERX$ = TRIM$(ATRIFORM$(CODFOR$, "INTERLIN"))
   '
   MULTIREN% = 1
   PPXX% = InStr(INTERX$, "/")
   If PPXX% > 0 Then
      MULTIREN% = XVALO(Mid$(INTERX$, PPXX% + 1))
      INTERX$ = Left$(INTERX$, PPXX% - 1)
   End If
   If MULTIREN% < 1 Then MULTIREN% = 1
   INTERLI = XVALO(INTERX$)
   If INTERLI < 1 Then
      'Call COMUNI("Error de Configuración de Interlinea\en Formulario '" + CODFOR$ + ".FRM'.")
      INTERLI = 4
   End If
   '
   DETDESMM = XVALO(ATRIFORM$(CODFOR$, "DETDESMM"))
   DETHASMM = XVALO(ATRIFORM$(CODFOR$, "DETHASMM"))
   If DETHASMM < DETDESMM + 4 + INTERLI Then
      DETHASMM = DETDESMM + 4 * INTERLI
   End If
   '
   OPCIM$ = TRIM$(UCase$(ATRIFORM$(CODFOR$, "OPCIMPRE")))
   If OPCIM$ = "SI" Then
     CANCELPRINT% = 0
     If COMALTER%("Opciones de Impresión:\\Imprime Comprobante\Suprime Impresión") = 2 Then
       CANCELPRINT% = 1
     End If
   End If
   '
   'CANHOJ& = 1
   'If CAITAB1% > 0 Then
   '   On Error Resume Next
   '   CANHOJ& = 1 + Int(MULTIREN% * CAITAB1% / ((DETHASMM - DETDESMM) / INTERLI))
   '   On Error GoTo 0
   'End If
   'If CANHOJ& < 1 Then CANHOJ& = 1
   '
   CACOPX$ = TRIM$(ATRIFORM$(CODFOR$, "CAN-COP"))
   If Left$(CACOPX$, 1) = "#" Then
5      CACOPI% = XVALO(InputBox$("Cantidad de Copias", "Ingrese Cantidad de Copias", Mid$(CACOPX$, 2)))
       If CACOPI% < 1 Or CACOPI% > 9 Then GoTo 5
     Else
       CACOPI% = XVALO(CACOPX$)
       If CACOPDOC% > 0 Then CACOPI% = CACOPDOC%
       If CACOPI% < 1 Or CACOPI% > 9 Then
         'Call COMUNI("Error de Configuración de Cantidad de Copias\en Formulario '" + CODFOR$ + ".FRM'.")
         CACOPI% = 1
       End If
   End If
   '
   If CACOPI% > 6 Then
     For KKII% = 2 To 9
       ORCOP$(KKII%) = "COPIA"
     Next KKII%
   End If
   '
   ORIENTA% = 1
   PORI$ = TRIM$(UCase$(ATRIFORM$(CODFOR$, "ORIENTA")))
   If PORI$ = "PAISAJE" Or PORI$ = "2" Then
      ORIENTA% = 2
   End If
   '
   PORIAN = Printer.Orientation  ' retiene orientacion
   'Printer.Orientation = 1
   On Error Resume Next
   Printer.ScaleMode = 1
   '
   Call SETSPOOL("SCALEMODE", "1")
   Call SETSPOOL("ORIENTATION", "1")
   Printer.Orientation = 1
   LARHOAN = Printer.Height
   If LARHOAN < 10 Then LARHOAN = 12 * 25.4 * 56.7
   ANCHOAN = Printer.Width
   If ANCHOAN < 10 Then ANCHOAN = 203 * 56.7
   On Error GoTo 0
   '
10 'Printer.Orientation = 1
   If ORIENTA% = 2 Then
        'Printer.Orientation = 2
        Call SETSPOOL("ORIENTATION", "2")
        'Call SETSPOOL("HEIGHT", Str$(ANCHOAN))
        'Call SETSPOOL("WIDTH", Str$(LARHOW * 56.7))
        'Printer.Width = ANCHOAN
      Else
        Call SETSPOOL("ORIENTATION", "1")
        'Call SETSPOOL("HEIGHT", Str$(LARHOW * 56.7))
        'Call SETSPOOL("WIDTH", Str$(ANCHOAN))
   End If
   '
   For COP% = 1 To CACOPI%
      '
      INIDET& = 1
      NUHOJ% = 0: NUHOK% = 0
      CANHOJ& = 1
      '
20    Call CARFORWIN         ' Carga Formulario en Blanco
      Call LIMPIAVALPAR
      '
      Call PRINTFORWIN(TIPODOC$, "TIPO-COM", 0, 0)
      Call PRINTFORWIN(NUMEDOC$, "NUME-COM", 0, 0)
      Call PRINTFORWIN(DOCUORIG$, "DOCUORIG", 0, 0)
      Call PRINTFORWIN(FECHDOC$, "FECH-EMI", 0, 0)
        AHORAX$ = REPLACE$(Format$(Now, "hh:mm:ss     dd/mm/yyyy"), "del", 10, 3)
      Call PRINTFORWIN(AHORAX$, "HORA-EMI", 0, 0)
      Call PRINTFORWIN(USERNAME$, "USUARIO", 0, 0)
      '
      Call PRINTFORWIN(ORCOP$(COP%), "ORD-COPI", 0, 0)
      '
      If Abs(DESTIDOC%) > 0 Then
         Call PRINTFORWIN(rasocli$(DESTIDOC%), "RASO-CLI", 0, 0)
         Call PRINTFORWIN(DOMICLI$(DESTIDOC%), "DOMI-CLI", 0, 0)
           ATRIX$ = TRIM$(ATRIFORM$(CODFOR$, "CPOS-CLI/FORMATO"))
           If ATRIX$ <> "" Then
               'IMPRIME SEPARADO COD-POSTAL Y LOCALIDAD
               Call PRINTFORWIN(CPOSCLI$(DESTIDOC%), "CPOS-CLI", 0, 0)
               Call PRINTFORWIN(LOCACLI$(DESTIDOC%), "LOCA-CLI", 0, 0)
             Else
               'iMPRIME CODIGO POSTAL Y LOCALIDAD TODO JUNTO
               Call PRINTFORWIN(CPLOCLI$(DESTIDOC%), "LOCA-CLI", 0, 0)
           End If
         Call PRINTFORWIN(DEPROVI$(PROVCLI$(DESTIDOC%)), "PCIA-CLI", 0, 0)
         Call PRINTFORWIN(PAISCLI$(DESTIDOC%), "PAIS-CLI", 0, 0)
         Call PRINTFORWIN(TEL1CLI$(DESTIDOC%), "TELE-CLI", 0, 0)
         Call PRINTFORWIN(FAXCLI$(DESTIDOC%), "FAX-CLI", 0, 0)
         Call PRINTFORWIN(EMAILCLI$(DESTIDOC%), "MAIL-CLI", 0, 0)
         Call PRINTFORWIN(CuitCli$(DESTIDOC%), "CUIT-CLI", 0, 0)
         Call PRINTFORWIN(NIBRCLI$(DESTIDOC%), "IBRU-CLI", 0, 0)
         Call PRINTFORWIN(POSIVAL$(DESTIDOC%), "PIVA-CLI", 0, 0)
         Call PRINTFORWIN(NPROCLI$(DESTIDOC%), "PROV-CLI", 0, 0)
           FORCUCLI$ = TRIM$(UCase$(ATRIFORM$(CODFOR$, "CUEN-CLI/FORMATO")))
           DEDOC$ = str$(Abs(DESTIDOC%))
           If Left$(FORCUCLI$, 1) = "A" Then
             DEDOC1$ = TRIM$(CODICLI$(DESTIDOC%))
             If DEDOC1$ <> "" Then DEDOC$ = DEDOC1$
           End If
         Call PRINTFORWIN(DEDOC$, "CUEN-CLI", 0, 0)
         '
         If DESTIDOC% > 0 Then
            Call PRINTFORWIN(NTRANCLI$(DESTIDOC%), "NTRA-CLI", 0, 0)
            Call PRINTFORWIN(RTRANCLI$(DESTIDOC%), "RTRA-CLI", 0, 0)
            Call PRINTFORWIN(DTRANCLI$(DESTIDOC%), "DTRA-CLI", 0, 0)
            Call PRINTFORWIN(LTRANCLI$(DESTIDOC%), "LTRA-CLI", 0, 0)
         End If
      End If
      '
      ' MONEDA DE EMISION
      Call PRINTFORWIN(DEMONECO$(MONEMI%), "MONEDEMI", 0, 0)
      '
      CIACTI% = 0
      For KKI% = 1 To CAPARDOC%          '
         Call PRINTFORWIN(DOCPARAM$(KKI%), CODPARAM$(KKI%), 0, 0)
         ' ESTO ES PARA RETENER EL VALOR PARA LAS IMAGENES
         If UCase$(CODPARAM$(KKI%)) = "COD-MAT" Then CIACTI% = CODINT%(DOCPARAM$(KKI%))
      Next KKI%
      '
      ' AQUI AGREGAR LOS PARAMETROS TOMADOS DE LA BASE DE DATOS
      XX$ = LOADFILE$(LUDAT$ + CODFOR$ + ".FRQ") 'ASIGNO A XX$ LA LECTURA DEL ARCHIVO
      For U% = 1 To Len(XX$) Step 256 'LO RECORRO (256 ES LA LONGITUD DEL REGISTRO)
        YY$ = Mid$(XX$, U%, 256)
        CODVARI$ = TRIM$(Left(YY$, 16))
        TABLASE$ = TRIM$(Mid$(YY$, 17, 24))
        If TABLASE$ <> "" Then
           CAMPOSE$ = TRIM$(Mid$(YY$, 41, 24))
           If CAMPOSE$ <> "" Then
             CONDI$ = TRIM$(Mid$(YY$, 65))
31           TBUS$ = "¿"
             PPXX1% = InStr(CONDI$, TBUS$)
             If PPXX1% > 0 Then
               TBUS$ = "?"
               PPXX2% = InStr(CONDI$, TBUS$)
               IZQUI$ = Left$(CONDI$, PPXX1% - 1)
               DERE$ = Mid$(CONDI$, PPXX2% + 1)
               AREPLA$ = ""
               If PPXX2% > PPXX1% + 1 Then
                  AREPLA$ = Mid$(CONDI$, PPXX1% + 1, PPXX2% - PPXX1% - 1)
               End If
               If AREPLA$ <> "" Then
                  For KKII% = 1 To CAVALPAR%
                     If CODPAR$(KKII%) = AREPLA$ Then
                        AREPLA$ = VALPAR$(KKII%)
                        GoTo 32
                     End If
                  Next KKII%
               End If
32             CONDI$ = IZQUI$ + AREPLA$ + DERE$
               GoTo 31
             End If
             '
             TTXX$ = SAFETEXT$(VALOBADA(TABLASE$, CAMPOSE$, CONDI$))
             Call PRINTFORWIN$(TTXX$, CODVARI$, 0, 0)
             ' ESTO ES PARA RETENER EL VALOR PARA LAS IMAGENES
             If UCase$(CODVARI$) = "COD-MAT" Then CIACTI% = CODINT%(TTXX$)
           End If
        End If
      Next U%
      '
      For KKI% = 1 To CAIMINI% + CAIMAGEN%
         CODIMA$ = UCase$(CODIMAGEN$(KKI%))
         If Left$(CODIMA$, 7) = "DOCUGRA" Then
           ' ESTA PARTE ES PARA INCRUSTAR LOS 4 DOCUMENTOS GRAFICOS
           If CIACTI% > 0 Then
             NIMA% = XVALO(Mid$(CODIMA$, 8, 1))
             If NIMA% <= 3 Then
               FGRC$ = FILTERGETRECORD$("INDIPLAN", 1, MKI$(CIACTI%))
               ARIMA$ = TRIM$(Mid$(FGRC$, 65 + 48 * NIMA%, 48))
                 RUTARGRA$ = TRIM$(CONTROL$("", "RUTARGRA"))
                 If RUTARGRA$ = "" Then RUTARGRA$ = TRIM$(LUDAT$)
                 If Right$(RUTARGRA$, 1) <> "\" Then RUTARGRA$ = RUTARGRA$ + "\"
               ARCHIMAGEN$(KKI%) = RUTARGRA$ + ARIMA$
             End If
           End If
         End If
         Call PRINTIMAGE(CODIMAGEN$(KKI%), ARCHIMAGEN$(KKI%))
      Next KKI%
      '
      YAC = 0
      For YX& = 1 To CAITAB2%
         For KKI% = 1 To CACOLTAB2%
            TTTXXX$ = TETABU2$(KKI%, YX&)
            Call PRINTFORWIN(TTTXXX$, CODTABU2$(KKI%), 0, YAC)
         Next KKI%
         YAC = YAC + INTERLI
      Next YX&
      '
      YAC = 0: INTEROBS = INTERLI
      If INTERLI > 6 Then INTEROBS = INTERLI / 2
      For UKU& = 1 To ULTREG&("!OBSERVOF")
         TXOB$ = TRIM$(REGLEIDO$("!OBSERVOF", UKU&))
         Call PRINTFORWIN(TXOB$, "OBSERVA", 0, YAC)
         YAC = YAC + INTEROBS
      Next UKU&
      Call CIERRARCH("!OBSERVOF")
      '
      YAC = XVALO(ATRIFORM$(CODFOR$, "DETDESMM"))
      '
      For YX& = INIDET& To CAITAB1%
         '
         If YAC > DETHASMM Then
            '
            NUHOJ% = NUHOJ% + 1
            If NUHOJ% = 1 Then
               On Error Resume Next
               'CANHOJ& = 1 + Int((CAITAB1% / (YX& - 1)) - 0.05)
               CANHOJ& = 1 + Int((CAITAB1% / (YX& - 1)) - 0.031) ' CAMBIE 0.05 POR 0.031
               On Error GoTo 0
            End If
            Call PRINTFORWIN(TRIM$(str$(NUHOJ%)), "HOJA-NUM", 0, 0)
            Call PRINTFORWIN(TRIM$(str$(CANHOJ&)), "HOJA-TOT", 0, 0)
            '
            INIDET& = YX&
            'Printer.NewPage
            Call SETSPOOL("NEWPAGE", "")
            GoTo 20
         End If
         '
         If TETABU1$(1, YX&) = String$(32, "-") Then
            Call PRILINSEP
            GoTo 89
         End If
         '
         YACMAX = 0
         For KKI% = 1 To CACOLTAB1%
            TTTXXX$ = TETABU1$(KKI%, YX&)
            If Left$(TTTXXX$, 1) = "@" Then
               YAC = YAC - INTERLI
               TTTXXX$ = Mid$(TTTXXX$, 2)
            End If
            Call PRINTFORWIN(TTTXXX$, CODTABU1$(KKI%), 0, YAC)
            On Error Resume Next
            If Printer.CurrentY / 56.7 - MARSUW + INTERLI > YACMAX Then
               YACMAX = Printer.CurrentY / 56.7 - MARSUW + INTERLI
            End If
            '
            On Error GoTo 0
         Next KKI%
         '
         If YACMAX < YAC Then YACMAX = YAC + INTERLI
         YAC = YACMAX

89    Next YX&
      '
      NUHOJ% = NUHOJ% + 1
      If NUHOJ% = 1 Then
         On Error Resume Next
         CANHOJ& = 1 + Int((CAITAB1% / (YX& - 1)) - 0.05)
         CANHOJ& = 1 + Int((CAITAB1% / (YX& - 1)) - 0.031) ' CAMBIE 0.05 POR 0.031
         On Error GoTo 0
      End If
      Call PRINTFORWIN(TRIM$(str$(NUHOJ%)), "HOJA-NUM", 0, 0)
      Call PRINTFORWIN(TRIM$(str$(CANHOJ&)), "HOJA-TOT", 0, 0)
      '
      If COP% < CACOPI% Then
         'Printer.NewPage
         Call SETSPOOL("NEWPAGE", "")
      End If
      '
   Next COP%
   '
   'Printer.EndDoc
99 Call SETSPOOL("ENDDOC", "")
   '
   If Len(TIPODOC$) >= 60 Then
      DESCRIDOC$ = TRIM$(TIPODOC$)
      GoTo 98
   End If
   '
   DESCRIDOC$ = TRIM$(TIPODOC$) + " Nro." + TRIM$(NUMEDOC$)
   If Len(TIPODOC$) > 8 Then
     DESCRIDOC$ = TRIM$(TRIM$(TIPODOC$) + " " + NUMEDOC$)
   End If
   DEDOCOR$ = REPLACAR$(DESCRIDOC$, " ", "_")
   If Abs(DESTIDOC%) > 0 Then
      DESCRIDOC$ = DESCRIDOC$ + " - " + TRIM$(rasocli$(DESTIDOC%))
   End If
   If TRIM$(FECHDOC$) <> "" Then
      DESCRIDOC$ = DESCRIDOC$ + " - " + FECHDOC$
   End If
   '
98 If InStr(UCase$(MODO$), "NOPRINT") < 1 Then
    Call SAVESPOOL(DESCRIDOC$, OKX%)
    '
    Call CIERRARCH("PDOCLST")
    Call CIERRARCH("PDOCSPL")
    Call CIERRARCH("PIMGLST")
    Call CIERRARCH("PIMGSPL")
    '
    Screen.MousePointer = 1
    If NOCONFIRM% < 1 Then
      If CANCELPRINT% < 1 Then
        If TRIM$(NUMEDOC$) <> "" Then
          If OKX% < 1 Then
            Call MENSERR(24, "Problemas de Impresión o de Generación de Archivo (1).")
            GoTo 49
          End If
          TTXYZ$ = "Se ha Emitido por Impresora"
          If SENDBYMAIL$ <> "" Then TTXYZ$ = "Se ha Emitido en Formato PDF para Envio por Mail"
          If MATRIXPUN% = 1 Then TTXYZ$ = TTXYZ$ + " (MP)"
          TTXYZ$ = TTXYZ$ + "\" + TIPODOC$ + " Numero '" + NUMEDOC$ + "'"
          If COMALTER%(TTXYZ$ + "\\Repetir Impresion\Continuar") = 1 Then GoTo 10
        End If
      End If
    End If
   End If
   '
49 On Error Resume Next
   Printer.Orientation = 1
   'Call SETSPOOL("ORIENTATION", "1")
   'Printer.Height = LARHOAN
   'Call SETSPOOL("HEIGHT", Str$(LARHOAN))
   'Printer.Width = ANCHOAN
   'Call SETSPOOL("WIDTH", Str$(ANCHOAN))
   Printer.Orientation = PORIAN
   'Call SETSPOOL("ORIENTATION", Str$(PORIAN))
   '
   If TRIM$(PRINTERPORT$) <> "" Then
      Call SETPRINPORT("RESTORE")
   End If
   '
   If InStr(UCase$(MODO$), "NOCANOBS") < 1 Then
      Call SETULTREG("!OBSERVOF", 0)
   End If
   Call CIERRARCH("!OBSERVOF")
   '
End Sub
'
Sub PRILINSEP()
   '
   LINSEP$ = TRIM$(ATRIFORM$(CODFOR$, "LIN-SEP"))
   If LINSEP$ = "" Then Exit Sub
   '
   GUION% = InStr(LINSEP$, "-")
   If GUION% < 1 Then Exit Sub
   '
   YAC = YAC + INTERLI / 4
   HH11 = (TRIM$(Mid$(LINSEP$, 1, GUION% - 1)) + MARIZW) * COEFH
   HH22 = (TRIM$(Mid$(LINSEP$, GUION% + 1)) + MARIZW) * COEFH
   VV11 = (YAC + MARSUW) * COEFV
   VV22 = (YAC + MARSUW) * COEFV
   Call SPOOLINE(HH11, VV11, HH22, VV22, QBColor(0), "")
   YAC = YAC + INTERLI / 4
   '
End Sub
'
Sub CARLINSEP(Optional tipolineaseparacion$)
   '
   If UCase$(TRIM$(tipolineaseparacion$)) = "INVISIBLE" Then
      tipolineaseparacion$ = " "
   Else
      tipolineaseparacion$ = "-"
   End If
   
   LINSEP$ = TRIM$(ATRIFORM$(CODFOR$, "LIN-SEP"))
   If LINSEP$ = "" Then Exit Sub
   '
   GUION% = InStr(LINSEP$, "-")
   If GUION% < 1 Then Exit Sub
   '
   CAITAB1% = CAITAB1% + 1
   TETABU1$(1, CAITAB1%) = String$(32, tipolineaseparacion$)
   '
End Sub

Sub PRINTIMAGE(IMAGECOD$, IMAGEFILE$)
   '
   If TRIM$(CODFOR$) = "" Then Exit Sub
   If EXISTE%(IMAGEFILE$) < 1 Then Exit Sub
   '
   CODVARI$ = TRIM$(UCase$(IMAGECOD$))
   If CODVARI$ = "" Then Exit Sub
   '
   Dim RFP As String * 128
   LU$ = LUDAT$
   NXP% = FILEOPEN%(LU$ + CODFOR$ + ".FRM", 0, 128)
   For i% = 1 To LOF(NXP%) / 128
      Get #NXP%, i%, RFP$
      If Left$(RFP$, 8) = "IMAGEWIN" Then
         If TRIM$(Mid$(RFP$, 9, 16)) = CODVARI$ Then
            GoTo 20
         End If
      End If
   Next i%
   Exit Sub
   '
20 ph1 = 56.7 * (CVS(Mid$(RFP$, 57, 4)) + MARIZW + X0)
   pv1 = 56.7 * (CVS(Mid$(RFP$, 61, 4)) + MARSUW + Y0)
   '
   PH2 = 56.7 * (CVS(Mid$(RFP$, 65, 4)) + MARIZW + X0)
   PV2 = 56.7 * (CVS(Mid$(RFP$, 69, 4)) + MARSUW + Y0)
   '
   FORMALTER.Image3.Picture = LoadPicture("")
   On Error Resume Next
   FORMALTER.Image3.Picture = LoadPicture(IMAGEFILE$)
   If Err Then
      On Error GoTo 0
      IMAFI$ = REPLACAR$(IMAGEFILE$, "\", "/")
      Call MENSERR(24, "Imposible Abrir Archivo Gráfico\'" + IMAFI$ + "'\en Formulario '" + CODFOR$ + "'.")
      Exit Sub
   End If
   On Error GoTo 0
   '
   HIMAGEN = FORMALTER.Image3.Picture.Width
   VIMAGEN = FORMALTER.Image3.Picture.Height
   '
   HCUADRO = PH2 - ph1
   VCUADRO = PV2 - pv1
   '
   REDUH = HIMAGEN / HCUADRO
   REDUV = VIMAGEN / VCUADRO
   REDUC = REDUH: If REDUV > REDUH Then REDUC = REDUV
   '
   altucua = VCUADRO * REDUV / REDUC
   anchcua = HCUADRO * REDUH / REDUC
   borizq = (HCUADRO - anchcua) / 2
   borsup = (VCUADRO - altucua) / 2
   '
   'Printer.PaintPicture FORMALTER.Image1.Picture, PH1 + BORIZQ, PV1 + BORSUP, ANCHCUA, ALTUCUA
   Call SPOOLIMAGE(IMAGEFILE$, ph1 + borizq, pv1 + borsup, anchcua, altucua)
   '
End Sub

Sub PRINTEXT(ESPECIFIX$, NPARA%, PARAMEXT$, Optional FORIPRE$)

       FORTELIB$ = TRIM$(UCase$(ATRIFORM$(FORIPRE$, PARAMEXT$ + "/FORMATO")))
       ATELIB% = XVALO(Mid$(FORTELIB$, 2))
       '
       XX$ = ESPECIFIX$
       Call FRATEXTO(XX$, ATELIB%)
       For jj& = 1 To CARETEX%
            CAITAB1% = CAITAB1% + 1
            TETABU1$(NPARA%, CAITAB1%) = RETEX$(jj&)
       Next jj&

End Sub

Sub AGREGATEXTO(REFER$, NPARA%)
    '
    CAITAB1% = CAITAB1% + 1
    TETABU1$(NPARA%, CAITAB1%) = REFER$
    '
End Sub
'
'*****************************************************************************
'
Sub FRATEX(TX$, NRX1$, NRX2$, NRX3$, LOTEX%)
   '
   Dim PAL$(24), NUMERO$(3)
   DE$ = TX$
   KK% = 1
   '
11550 For KM% = 1 To Len(DE$)
11560 PAL$(KK%) = PAL$(KK%) + Mid$(DE$, KM%, 1)
11570 If Mid$(DE$, KM%, 1) = " " Then KK% = KK% + 1
      If KK% > 24 Then GoTo 11590
11580 Next KM%
      '
11590 KN% = 1
11600 LON% = 0
11610 For KL% = 1 To KK%
11620    If LON% + Len(PAL$(KL%)) > LOTEX% Then
            If KN% < 3 Then
               KN% = KN% + 1
               LON% = 0
            End If
         End If
11630    NUMERO$(KN%) = NUMERO$(KN%) + PAL$(KL%)
         LON% = LON% + Len(PAL$(KL%))
11655 Next KL%
      '
11660 NRX1$ = NUMERO$(1)
      NRX2$ = NUMERO$(2)
      NRX3$ = NUMERO$(3)
      '
End Sub
'
Sub PRINTFORWIN(TEXT$, CODVARI$, X0, Y0)
    '
    Static CODFORA$
    Static OFSECOPIA%
    Static FUENPOROMIS$
    '
    Static RFXP$(), LNXP&
    '
    If TRIM$(TEXT$) = "" Then
      Exit Sub
    End If
    '
    If TRIM$(CODFOR$) = "" Then
        Call MENSERR(24, "Falta Seleccionar Formulario")
        PRINFORM.Text1.SetFocus
        Exit Sub
    End If
    '
    If CODFOR$ <> CODFORA$ Then
      AFORMX$ = LOADFILE$(LUDAT$ + CODFOR$ + ".FRM")
      LNXP& = 1 + Len(AFORMX$) / 128
      ReDim RFXP$(LNXP&)
      For KU& = 1 To LNXP&
        RFXP$(KU&) = Mid$(AFORMX$, 128 * KU& - 127, 128)
      Next KU&
      CODFORA$ = CODFOR$
    End If
    '
    If FUENPOROMIS$ = "" Then
      FUENPOROMIS$ = ATRIFORM$("", "FUEN-DEF")
    End If
    '
    Call SETSPOOL("CURRENTX", str$(0))
    Call SETSPOOL("CURRENTY", str$(0))
    '
    If TRIM$(TEXT$) = "" Then
       GoTo 900
    End If
    '
    On Error Resume Next
    If CAFUEN% = 0 Then
      For i% = 1 To Printer.FontCount
        FUEX$ = TRIM$(Printer.Fonts(i% - 1))
        If FUEX$ <> "" Then
          If CAFUEN% < 1024 Then
            CAFUEN% = CAFUEN% + 1
            Fuente$(CAFUEN%) = FUEX$
          End If
        End If
      Next i%
    End If
    On Error GoTo 0
    '
    For i% = 1 To LNXP&
      RFP$ = RFXP$(i%)
      If TRIM$(Left$(RFP$, 8)) = CODVARI$ Then
        GoTo 200
      End If
    Next i%
    Exit Sub
    '
200 ' AQUI ACUMULA LOS ULTIMOS VALORES DE CADA PARAMETRO
    For KKII% = 1 To CAVALPAR%
       If CODPAR$(KKII%) = UCase$(CODVARI$) Then
          VALPAR$(KKII%) = TEXT$
          GoTo 202
       End If
    Next KKII%
    CAVALPAR% = CAVALPAR% + 1
    CODPAR$(CAVALPAR%) = UCase$(CODVARI$)
    VALPAR$(CAVALPAR%) = TEXT$
    '
202 FUENX$ = TRIM$(Mid$(RFP$, 87, 32))
    FORMA1$ = TRIM$(Mid$(RFP$, 75, 6))
    FORMA2$ = TRIM$(Mid$(RFP$, 81, 6))
    If TRIM$(FORMA1$ + FORMA2$) = "" Then
      GoTo 900
    End If
    '
    If FUENX$ = "" Then FUENX$ = FUENPOROMIS$
    '
    If FUENX$ = "" Then
      Call MENSERR(24, "Fuente No Definida para\Parámetro '" + TRIM$(CODVARI$) + "'\en Formulario '" + TRIM$(CODFOR$) + "'.")
      FUEREPLA$ = "Courier"
      GoTo 203
    End If
    '
    For i% = 1 To CAFUEN%
      If Fuente$(i%) = FUENX$ Then GoTo 205
    Next i%
    '
    ' ESTO SE AGREGO PARA QUE RECONOZCA LA FUENTE CODE 39 DE SATO EN MEDITEA
    If UCase$(Left$(FUENX$, 4)) = "CODE" Then
       CAFUEN% = CAFUEN% + 1
       Fuente$(CAFUEN%) = FUENX$
    End If
    '
    If FUEREPLA$ = "" Then
      FUEREPLA$ = "Courier"
      TTXX$ = "Fuente '" + FUENX$ + "' no Disponible\en el Presente Modelo de Impresora\para Parámetro '" + TRIM$(CODVARI$) + "'\en Formulario '" + TRIM$(CODFOR$) + "'."
      TTXX$ = TTXX$ + "\   \FLEXOFT Reemplazará las Fuentes\no Disponibles por '" + FUEREPLA$ + "'."
      Call COMUNI(TTXX$)
    End If
    '
203 FUENX$ = FUEREPLA$
    '
205 On Error GoTo 999
    Call SETSPOOL("FONTNAME", FUENX$)

    If CVS(Mid$(RFP$, 119, 4)) < 1 Then
      Call MENSERR(24, "Tamaño de Fuente Inexistente\para Parámetro '" + TRIM$(CODVARI$) + "'\en Formulario '" + TRIM$(CODFOR$) + "'.")
      Exit Sub
    End If
    Call SETSPOOL("FONTSIZE", str$(CVS(Mid$(RFP$, 119, 4))))
    On Error GoTo 0
    '
    If Asc(Mid$(RFP$, 123, 1)) > 32 Then
      Call SETSPOOL("FONTBOLD", "1")
    End If
    If Asc(Mid$(RFP$, 124, 1)) > 32 Then
      Call SETSPOOL("FONTITALIC", "1")
    End If
    If Asc(Mid$(RFP$, 125, 1)) > 32 Then
      Call SETSPOOL("FONTUNDERLINE", "1")
    End If
    '
    ph1 = 56.7 * (CVS(Mid$(RFP$, 57, 4)) + MARIZW + X0)
    pv1 = 56.7 * (CVS(Mid$(RFP$, 61, 4)) + MARSUW + Y0)
    FORMA1$ = TRIM$(Mid$(RFP$, 75, 6))
    '
    PH2 = 56.7 * (CVS(Mid$(RFP$, 65, 4)) + MARIZW + X0)
    PV2 = 56.7 * (CVS(Mid$(RFP$, 69, 4)) + MARSUW + Y0)
    FORMA2$ = TRIM$(Mid$(RFP$, 81, 6))
    '
    If FORMA1$ <> "" Then
      If UCase$(Left$(TRIM$(FORMA1$), 1) = "Y") Then
        PPXXZZ% = InStr(FORMA1$, "/")
        If PPXXZZ% < 3 Then GoTo 880 ' NO PROCESABLE
        ESPACAR% = XVALO(Mid$(FORMA1$, 2, PPXXZZ% - 2))
        MILIDES = XVALO(Mid$(FORMA1$, PPXXZZ% + 1)) / 10
        If ESPACAR < 1 Or MILIDES < 0.1 Then GoTo 880
        '
        TTXX$ = TEXT$
        Do While TRIM$(TTXX$) <> ""
          TTYY$ = Left$(TTXX$, ESPACAR%)
          TTXX$ = Mid$(TTXX$, ESPACAR% + 1)
          Call SETSPOOL("CURRENTX", str$(ph1))
          Call SETSPOOL("CURRENTY", str$(pv1))
          Call SETSPOOL("PRINT", TTYY$ + ";")
          ph1 = ph1 + 56.7 * MILIDES
        Loop
        GoTo 880
      End If
      '
      TTXX$ = TEXTFORM$(TEXT$, FORMA1$)
      If Left$(TTXX$, 1) = "." Then TTXX$ = " " + Mid$(TTXX$, 2)
      '
      ' ESTO SE AGREGO PARA QUE EN MEDITEA CON LA FUENTE CODE 39 PONGA ASTERISCO ADELANTE Y ATRAS - EPB 14/2/12
      If UCase$(Left$(FUENX$, 4)) = "CODE" Then
         TTXX$ = "*" + TRIM$(TTXX$) + "*" ' SE AGREGA TRIM PARA QUITAR BLANCOS AL FINAL
      End If
      '
      If UCase$(Left$(TRIM$(FORMA1$), 1) = "X") Then
        PPXXZZ% = InStr(FORMA1$, "/"): If PPXXZZ% < 1 Then PPXXZZ% = 1 + Len(FORMA1$)
        CAREMA% = XVALO(Mid$(FORMA1$, 1 + PPXXZZ%)): If CAREMA% < 1 Then CAREMA% = 1
        INTERLIX = INTERLI: If INTERLIX / CAREMA% < 3 Then INTERLIX = INTERLI * (1 + CAREMA%)
        For CAREI% = 1 To CARETEX%
          If CAREI% <= CAREMA% Then
            Call SETSPOOL("CURRENTX", str$(ph1))
            Call SETSPOOL("CURRENTY", str$(pv1))
            Call SETSPOOL("PRINT", RETEX$(CAREI%) + ";")
            pv1 = pv1 + INTERLIX / (CAREMA% + 1) * 56.7
          End If
        Next CAREI%
        GoTo 880
      End If
      If OFSECOPIA% > 0 Then
        TTXY$ = TTXX$
        While Len(TTXX$) < OFSECOPIA%
          TTXX$ = TTXX$ + " "
        Wend
        TTXX$ = RTrim$(TTXX$ + TTXY$)
      End If
      If TRIM$(TTXX$) <> "" Then
        Call SETSPOOL("CURRENTX", str$(ph1))
        Call SETSPOOL("CURRENTY", str$(pv1))
        Call SETSPOOL("PRINT", "" + ";")
        Call SETSPOOL("CURRENTX", str$(ph1))
        Call SETSPOOL("CURRENTY", str$(pv1))
        If TRIM$(TTXX$) <> "" Then
           Call SETSPOOL("PRINT", TTXX$ + ";")
        End If
      End If
    End If
    '
880 If OFSECOPIA% < 1 Then
      If FORMA2$ <> "" Then
        TTXX$ = TEXTFORM$(TEXT$, FORMA2$)
        If UCase$(Left$(TRIM$(FORMA2$), 1) = "X") Then
          PPXXZZ% = InStr(FORMA2$, "/"): If PPXXZZ% < 1 Then PPXXZZ% = 1 + Len(FORMA2$)
          CAREMA% = XVALO(Mid$(FORMA2$, 1 + PPXXZZ%)): If CAREMA% < 1 Then CAREMA% = 1
          INTERLIX = INTERLI: If INTERLIX / CAREMA% < 3 Then INTERLIX = INTERLI * (1 + CAREMA%)
          For CAREI% = 1 To CARETEX%
            If CAREI% <= CAREMA% Then
              Call SETSPOOL("CURRENTX", str$(PH2))
              Call SETSPOOL("CURRENTY", str$(PV2))
              Call SETSPOOL("PRINT", RETEX$(CAREI%) + ";")
              PV2 = PV2 + INTERLIX / (CAREMA% + 1) * 56.7
            End If
          Next CAREI%
          GoTo 890
        End If
        If TRIM$(TTXX$) <> "" Then
          Call SETSPOOL("CURRENTX", str$(PH2))
          Call SETSPOOL("CURRENTY", str$(PV2))
          Call SETSPOOL("PRINT", "" + ";")
          Call SETSPOOL("CURRENTX", str$(PH2))
          Call SETSPOOL("CURRENTY", str$(PV2))
          If TRIM$(TTXX$) <> "" Then
            Call SETSPOOL("PRINT", TTXX$ + ";")
          End If
        End If
      End If
890 End If
    '
900 Call SETSPOOL("FONTBOLD", "0")
    Call SETSPOOL("FONTITALIC", "0")
    Call SETSPOOL("FONTUNDERLINE", "0")
    '
999 On Error GoTo 0
    '
End Sub

Sub PRINTFORWINVIE(TEXT$, CODVARI$, X0, Y0)
    '
    Static NXP%
    Static RFP As String * 128
    Static CODFORA$, LNXPA&
    Static OFSECOPIA%
    Static FUENPOROMIS$
    '
    If FUENPOROMIS$ = "" Then
      FUENPOROMIS$ = ATRIFORM$("", "FUEN-DEF")
    End If
    '
    On Error Resume Next
    LNXP& = LOF(NXP%)
    On Error GoTo 0
    If LNXP& < 1 Or LNXP& <> LNXPA& Or CODFOR$ <> CODFORA$ Then
      If NXP% > 0 Then Close #NXP%
      NXP% = 0
    End If
    LNXPA& = LNXP&
    '
    'Printer.CurrentX = 0
    Call SETSPOOL("CURRENTX", str$(0))
    'Printer.CurrentY = 0
    Call SETSPOOL("CURRENTY", str$(0))
    '
    If TRIM$(TEXT$) = "" Then
        GoTo 900
        ' Exit Sub
    End If
    '
    If TRIM$(CODFOR$) = "" Then
        Call MENSERR(24, "Falta Seleccionar Formulario")
        PRINFORM.Text1.SetFocus
        Exit Sub
    End If
    '
    On Error Resume Next
    If CAFUEN% = 0 Then
        For i% = 1 To Printer.FontCount
            FUEX$ = TRIM$(Printer.Fonts(i% - 1))
            If FUEX$ <> "" Then
                If CAFUEN% < 1024 Then
                    CAFUEN% = CAFUEN% + 1
                    Fuente$(CAFUEN%) = FUEX$
                End If
            End If
        Next i%
    End If
    On Error GoTo 0
    '
    If CODFOR$ <> CODFORA$ Then
        If NXP% > 0 Then
            Close #NXP%: NXP% = 0
            NXP% = 0
            OFSECOPIA% = 0
        End If
        CODFORA$ = CODFOR$
    End If
    '
    If NXP% = 0 Then
        NXP% = FILEOPEN%(LU$ + CODFOR$ + ".FRM", 0, 128)
        For i% = 1 To LOF(NXP%) / 128
           Get #NXP%, i%, RFP$
           If TRIM$(Left$(RFP$, 8)) = "MARGENIN" Then
              OFSECOPIA% = XVALO(Mid$(RFP$, 89, 8))
           End If
       Next i%
    End If
    '
    For i% = 1 To LOF(NXP%) / 128
        Get #NXP%, i%, RFP$
        If TRIM$(Left$(RFP$, 8)) = CODVARI$ Then
            GoTo 200
        End If
    Next i%
    Exit Sub
    '
200 FUENX$ = TRIM$(Mid$(RFP$, 87, 32))
    FORMA1$ = TRIM$(Mid$(RFP$, 75, 6))
    FORMA2$ = TRIM$(Mid$(RFP$, 81, 6))
    If TRIM$(FORMA1$ + FORMA2$) = "" Then
        GoTo 900
    End If
    '
    If FUENX$ = "" Then FUENX$ = FUENPOROMIS$
    '
    If FUENX$ = "" Then
        Call MENSERR(24, "Fuente No Definida para\Parámetro '" + TRIM$(CODVARI$) + "'\en Formulario '" + TRIM$(CODFOR$) + "'.")
        FUEREPLA$ = "Courier"
        GoTo 203
    End If
    '
    For i% = 1 To CAFUEN%
        If Fuente$(i%) = FUENX$ Then GoTo 205
    Next i%
    If FUEREPLA$ = "" Then
        FUEREPLA$ = "Courier"
        TTXX$ = "Fuente '" + FUENX$ + "' no Disponible\en el Presente Modelo de Impresora\para Parámetro '" + TRIM$(CODVARI$) + "'\en Formulario '" + TRIM$(CODFOR$) + "'."
        TTXX$ = TTXX$ + "\   \FLEXOFT Reemplazará las Fuentes\no Disponibles por '" + FUEREPLA$ + "'."
        Call COMUNI(TTXX$)
    End If
203 FUENX$ = FUEREPLA$
    '
205 On Error GoTo 999
    'Printer.FontName = FUENX$
    Call SETSPOOL("FONTNAME", FUENX$)

    If CVS(Mid$(RFP$, 119, 4)) < 1 Then
        Call MENSERR(24, "Tamaño de Fuente Inexistente\para Parámetro '" + TRIM$(CODVARI$) + "'\en Formulario '" + TRIM$(CODFOR$) + "'.")
        Exit Sub
    End If
    'Printer.FontSize = CVS(Mid$(RFP$, 119, 4))
    Call SETSPOOL("FONTSIZE", str$(CVS(Mid$(RFP$, 119, 4))))
    On Error GoTo 0
    '
    If Asc(Mid$(RFP$, 123, 1)) > 32 Then
       'Printer.FontBold = True
       Call SETSPOOL("FONTBOLD", "1")
    End If
    If Asc(Mid$(RFP$, 124, 1)) > 32 Then
       'Printer.FontItalic = True
       Call SETSPOOL("FONTITALIC", "1")
    End If
    If Asc(Mid$(RFP$, 125, 1)) > 32 Then
       'Printer.FontUnderline = True
       Call SETSPOOL("FONTUNDERLINE", "1")
    End If
    '
    ph1 = 56.7 * (CVS(Mid$(RFP$, 57, 4)) + MARIZW + X0)
    pv1 = 56.7 * (CVS(Mid$(RFP$, 61, 4)) + MARSUW + Y0)
    FORMA1$ = TRIM$(Mid$(RFP$, 75, 6))
    '
    PH2 = 56.7 * (CVS(Mid$(RFP$, 65, 4)) + MARIZW + X0)
    PV2 = 56.7 * (CVS(Mid$(RFP$, 69, 4)) + MARSUW + Y0)
    FORMA2$ = TRIM$(Mid$(RFP$, 81, 6))
    '
    If FORMA1$ <> "" Then
        TTXX$ = TEXTFORM$(TEXT$, FORMA1$)
        If Left$(TTXX$, 1) = "." Then TTXX$ = " " + Mid$(TTXX$, 2)
        If UCase$(Left$(TRIM$(FORMA1$), 1) = "X") Then
           PPXXZZ% = InStr(FORMA1$, "/"): If PPXXZZ% < 1 Then PPXXZZ% = 1 + Len(FORMA1$)
           CAREMA% = XVALO(Mid$(FORMA1$, 1 + PPXXZZ%)): If CAREMA% < 1 Then CAREMA% = 1
           INTERLIX = INTERLI: If INTERLIX / CAREMA% < 3 Then INTERLIX = INTERLI * (1 + CAREMA%)
           For CAREI% = 1 To CARETEX%
             If CAREI% <= CAREMA% Then
               'Printer.CurrentX = PH1
               Call SETSPOOL("CURRENTX", str$(ph1))
               'Printer.CurrentY = PV1
               Call SETSPOOL("CURRENTY", str$(pv1))
               'Printer.Print RETEX$(CAREI%);
               Call SETSPOOL("PRINT", RETEX$(CAREI%) + ";")
               pv1 = pv1 + INTERLIX / (CAREMA% + 1) * 56.7
             End If
           Next CAREI%
           GoTo 880
        End If
        If OFSECOPIA% > 0 Then
           TTXY$ = TTXX$
           While Len(TTXX$) < OFSECOPIA%
              TTXX$ = TTXX$ + " "
           Wend
           TTXX$ = RTrim$(TTXX$ + TTXY$)
        End If
        If TRIM$(TTXX$) <> "" Then
            Call SETSPOOL("CURRENTX", str$(ph1))
            Call SETSPOOL("CURRENTY", str$(pv1))
            Call SETSPOOL("PRINT", "" + ";")
            Call SETSPOOL("CURRENTX", str$(ph1))
            Call SETSPOOL("CURRENTY", str$(pv1))
            If TRIM$(TTXX$) <> "" Then
               Call SETSPOOL("PRINT", TTXX$ + ";")
            End If
        End If
    End If
    '
880 If OFSECOPIA% < 1 Then
       If FORMA2$ <> "" Then
         TTXX$ = TEXTFORM$(TEXT$, FORMA2$)
         If UCase$(Left$(TRIM$(FORMA2$), 1) = "X") Then
           PPXXZZ% = InStr(FORMA2$, "/"): If PPXXZZ% < 1 Then PPXXZZ% = 1 + Len(FORMA2$)
           CAREMA% = XVALO(Mid$(FORMA2$, 1 + PPXXZZ%)): If CAREMA% < 1 Then CAREMA% = 1
           INTERLIX = INTERLI: If INTERLIX / CAREMA% < 3 Then INTERLIX = INTERLI * (1 + CAREMA%)
           For CAREI% = 1 To CARETEX%
             If CAREI% <= CAREMA% Then
               'Printer.CurrentX = PH1
               Call SETSPOOL("CURRENTX", str$(ph1))
               'Printer.CurrentY = PV1
               Call SETSPOOL("CURRENTY", str$(pv1))
               'Printer.Print RETEX$(CAREI%);
               Call SETSPOOL("PRINT", RETEX$(CAREI%) + ";")
               pv1 = pv1 + INTERLIX / (CAREMA% + 1) * 56.7
             End If
           Next CAREI%
           GoTo 890
         End If
         If TRIM$(TTXX$) <> "" Then
              'Printer.CurrentX = PH2 ' - Printer.TextHeight(TTXX$)
              Call SETSPOOL("CURRENTX", str$(PH2))
              'Printer.CurrentY = PV2
              Call SETSPOOL("CURRENTY", str$(PV2))
              'Printer.Print "";
              Call SETSPOOL("PRINT", "" + ";")
              'Printer.CurrentX = PH2 ' - Printer.TextHeight(TTXX$)
              Call SETSPOOL("CURRENTX", str$(PH2))
              'Printer.CurrentY = PV2
              Call SETSPOOL("CURRENTY", str$(PV2))
              If TRIM$(TTXX$) <> "" Then
                 'Printer.Print TTXX$;
                 Call SETSPOOL("PRINT", TTXX$ + ";")
              End If
         End If
       End If
890 End If
    '
900 'Printer.FontBold = False
    Call SETSPOOL("FONTBOLD", "0")
    'Printer.FontItalic = False
    Call SETSPOOL("FONTITALIC", "0")
    'Printer.FontUnderline = False
    Call SETSPOOL("FONTUNDERLINE", "0")
    '
999 On Error GoTo 0
    '
End Sub

Function NUMTEX$(valor#)
Dim UND$(19), DECA$(9), CENTO$(19): GoSub 11690

11000 NUMERO$ = "": valor# = (Int(100 * valor# + 0.5)) / 100
11010 ENT# = Int(valor#)
11020 DECI# = Int(100 * (valor# - ENT#) + 0.5)
11030 MILLO# = Int(ENT# / 1000000!)
11040 MILI# = Int((ENT# - MILLO# * 1000000!) / 1000)
11050 UND1# = Int(ENT# - MILLO# * 1000000! - MILI# * 1000)
11060 BLOQUE# = MILLO#
11070 If BLOQUE# = 0 Then GoTo 11130
11080 GoSub 11270
11090 If BLOQUE# > 1 Then GoTo 11120
11100 NUMERO$ = BLOQ$ + " MILLON "
11110 GoTo 11130
11120 NUMERO$ = BLOQ$ + " MILLONES "
11130 BLOQUE# = MILI#
11140 If BLOQUE# = 0 Then GoTo 11170
11150 GoSub 11270
11160 NUMERO$ = NUMERO$ + BLOQ$ + " MIL "
11170 BLOQUE# = UND1#
11180 If BLOQUE# = 0 Then GoTo 11210
11190 GoSub 11270
11200 NUMERO$ = NUMERO$ + BLOQ$
11205 If Right$(NUMERO$, 2) = "UN" Then NUMERO$ = NUMERO$ + "O"
11207 MPRI$ = ATRIFORM$("", "MON-PRI")
11210 If TRIM$(MPRI$) <> "" Then
          NUMERO$ = NUMERO$ + " " + MPRI$
      End If
      NUMERO$ = NUMERO$ + " CON "
11220 BLOQUE# = DECI#: GoSub 11270
11225 MCEN$ = ATRIFORM$("", "MON-CEN")
11230 NUMERO$ = NUMERO$ + BLOQ$
      If TRIM$(MCEN$) <> "" Then
          NUMERO$ = NUMERO$ + " " + MCEN$
      End If
11240 GoTo 40000
11260 '
11270 BLOQ$ = ""
11280 B1 = Int(BLOQUE# / 100)
11290 B2 = Int((BLOQUE# - B1 * 100) / 10)
11300 B3 = Int(BLOQUE# - B1 * 100 - B2 * 10)
11310 If B2 >= 2 Then GoTo 11340
11320 B3 = B2 * 10 + B3
11330 B2 = 0
11340 If B1 = 0 Then GoTo 11370
11345 '
11350 If B1 > 19 Then                ' PARA PREVENIR ERROR 9 POR FUERA DE RANGO
         NUMERO$ = CStr(valor#)
         GoTo 40000
      End If
      BLOQ$ = CENTO$(B1)
11360 If ((B2 <> 0) Or (B3 <> 0)) And (B1 = 1) Then BLOQ$ = BLOQ$ + "TO "
11370 If B2 = 0 Then GoTo 11420
11380 BLOQ$ = BLOQ$ + DECA$(B2)
11390 If (B3 = 0) And (B2 = 2) Then BLOQ$ = BLOQ$ + "E"
11400 If (B3 <> 0) And (B2 = 2) Then BLOQ$ = BLOQ$ + "I"
11410 If (B3 <> 0) And (B2 > 2) Then BLOQ$ = BLOQ$ + " Y "
11420 If B3 = 0 Then GoTo 11440
11430 BLOQ$ = BLOQ$ + UND$(B3)
11440 If ((B1 = 0) And (B2 = 0) And (B3 = 0)) Then If BLOQUE# = DECI# Then BLOQ$ = "CERO"
11450 Return
11460 '
11690 DTUNI$ = " UN, DOS, TRES, CUATRO, CINCO, SEIS, SIETE, OCHO, NUEVE, DIEZ, ONCE, DOCE, TRECE, CATORCE, QUINCE, DIECISEIS, DIECISIETE, DIECIOCHO, DIECINUEVE"
11700 DTDEC$ = ", VEINT, TREINTA, CUARENTA, CINCUENTA, SESENTA, SETENTA, OCHENTA, NOVENTA"
11710 DTCEN$ = " CIEN, DOSCIENTOS , TRESCIENTOS , CUATROCIENTOS , QUINIENTOS , SEISCIENTOS , SETECIENTOS , OCHOCIENTOS , NOVECIENTOS "
11720 'RESTORE 11690
11730 For i = 1 To 19
        PPXX% = InStr(DTUNI$, ","): If PPXX% < 1 Then PPXX% = 1 + Len(DTUNI$)
        UND$(i) = TRIM$(Left$(DTUNI$, PPXX% - 1))
        DTUNI$ = Mid$(DTUNI$, PPXX% + 1)
      Next i
11740 For i = 1 To 9
        PPXX% = InStr(DTDEC$, ","): If PPXX% < 1 Then PPXX% = 1 + Len(DTDEC$)
        DECA$(i) = TRIM$(Left$(DTDEC$, PPXX% - 1))
        DTDEC$ = Mid$(DTDEC$, PPXX% + 1)
      Next i
11750 For i = 1 To 9
        PPXX% = InStr(DTCEN$, ","): If PPXX% < 1 Then PPXX% = 1 + Len(DTCEN$)
        CENTO$(i) = TRIM$(Left$(DTCEN$, PPXX% - 1))
        If i > 1 Then CENTO$(i) = CENTO$(i) + " "
        DTCEN$ = Mid$(DTCEN$, PPXX% + 1)
      Next i
11760 Return
11770 '
40000 NUMTEX$ = NUMERO$
40010 End Function

Sub CARFORWIN()
    '
    If TRIM$(CODFOR$) = "" Then
        Call MENSERR(24, "Falta Seleccionar Formulario")
        PRINFORM.Text1.SetFocus
        Exit Sub
    End If
    '
    On Error Resume Next
    If CAFUEN% = 0 Then
        For i% = 1 To Printer.FontCount
            FUEX$ = TRIM$(Printer.Fonts(i% - 1))
            If FUEX$ <> "" Then
                If CAFUEN% < 1024 Then
                    CAFUEN% = CAFUEN% + 1
                    Fuente$(CAFUEN%) = FUEX$
                End If
            End If
        Next i%
    End If
    On Error GoTo 0
    '
    LU$ = LUDAT$
    'Printer.ScaleMode = 1
    Call SETSPOOL("SCALEMODE", "1")
    COEFH = 56.7
    COEFV = 56.7
    On Error Resume Next
    ANCHOMAX = Printer.ScaleWidth
    On Error GoTo 0
    If ANCHOMAX < 10 Then ANCHOMAX = 203 * COEFH
    'Printer.ColorMode = 2   ' blanco y negro
    On Error Resume Next
    'Printer.ForeColor = QBColor(0)
    Call SETSPOOL("FORECOLOR", str$(QBColor(0)))
    'Printer.DrawMode = 1    ' linea negra
    Call SETSPOOL("DRAWMODE", "13")
    '
    NX% = FILEOPEN%(LU$ + CODFOR$ + ".FRM", 0, 128)
    Dim RF As String * 128
    '
' dibuja cajas *************************************
    For i% = 1 To LOF(NX%) / 128
        Get #NX%, i%, RF$
        If TRIM$(Left$(RF$, 8)) = "CAJAWIN" Then
            '
            'Printer.DrawWidth = 3
            Call SETSPOOL("DRAWWIDTH", "3")
            If Asc(Mid$(RF$, 73, 1)) > 0 Then
                'Printer.DrawWidth = Asc(Mid$(RF$, 73, 1))
                Call SETSPOOL("DRAWWIDTH", str$(Asc(Mid$(RF$, 73, 1))))
            End If
            '
            'Printer.DrawStyle = 0
            Call SETSPOOL("DRAWSTYLE", "0")
            If Asc(Mid$(RF$, 74, 1)) < 5 Then
                'Printer.DrawStyle = Asc(Mid$(RF$, 74, 1))
                Call SETSPOOL("DRAWSTYLE", str$(Asc(Mid$(RF$, 74, 1))))
            End If
            '
            'Printer.FillStyle = 1
            Call SETSPOOL("FILLSTYLE", "1")
            '
            HXX1 = (CVS(Mid$(RF$, 57, 4)) + MARIZW) * COEFH
            VXX1 = (CVS(Mid$(RF$, 61, 4)) + MARSUW) * COEFV
            HXX2 = (CVS(Mid$(RF$, 65, 4)) + MARIZW) * COEFH
            VXX2 = (CVS(Mid$(RF$, 69, 4)) + MARSUW) * COEFV
            'Printer.Line (H1, V1)-(H2, V2), QBColor(0), B
            Call SPOOLINE(HXX1, VXX1, HXX2, VXX2, QBColor(0), "B")
            '
        End If
    Next i%
    '
' dibuja lineas *************************************
    For i% = 1 To LOF(NX%) / 128
        Get #NX%, i%, RF$
        If TRIM$(Left$(RF$, 8)) = "LINEWIN" Then
            '
            'Printer.DrawWidth = 1
            Call SETSPOOL("DRAWWIDTH", "1")
            If Asc(Mid$(RF$, 73, 1)) > 0 Then
                'Printer.DrawWidth = Asc(Mid$(RF$, 73, 1))
                Call SETSPOOL("DRAWWIDTH", str$(Asc(Mid$(RF$, 73, 1))))
            End If
            '
            'Printer.DrawStyle = 0
            Call SETSPOOL("DRAWSTYLE", "0")
            If Asc(Mid$(RF$, 74, 1)) < 5 Then
                'Printer.DrawStyle = Asc(Mid$(RF$, 74, 1))
                Call SETSPOOL("DRAWSTYLE", str$(Asc(Mid$(RF$, 74, 1))))
            End If
            '
            HXX1 = (CVS(Mid$(RF$, 57, 4)) + MARIZW) * COEFH
            VXX1 = (CVS(Mid$(RF$, 61, 4)) + MARSUW) * COEFV
            HXX2 = (CVS(Mid$(RF$, 65, 4)) + MARIZW) * COEFH
            VXX2 = (CVS(Mid$(RF$, 69, 4)) + MARSUW) * COEFV
            'Printer.Line (H1, V1)-(H2, V2), QBColor(0)
            Call SPOOLINE(HXX1, VXX1, HXX2, VXX2, QBColor(0), "")
            '
        End If
    Next i%
    '
'Exit Sub
' escribe leyendas fijas ***************************
    For i% = 1 To LOF(NX%) / 128
        Get #NX%, i%, RF$
        If TRIM$(Left$(RF$, 8)) = "TEXTWIN" Then
            TXTX$ = TRIM$(Mid$(RF$, 9, 48))
            If TXTX$ = "" Then
                GoTo 109
            End If
            '
            If Left$(TXTX$, 1) = "?" Then
              CODICOND$ = ""
              TXTX$ = TRIM$(Mid$(TXTX$, 2))
              PPXX% = InStr(TXTX$, "?")
              If PPXX% < 1 Then GoTo 101
              '
              CODICOND$ = TRIM$(UCase$(Left$(TXTX$, PPXX% - 1)))
              TXTX$ = TRIM$(Mid$(TXTX$, PPXX% + 1))
              If CODICOND$ <> "" Then
                If CODICOND$ <> UCase$(CODOCUM$) Then
                  GoTo 109
                End If
              End If
            End If
            '
            'Printer.DrawStyle = 0
101         Call SETSPOOL("DRAWSTYLE", "0")
            'Printer.ForeColor = QBColor(0)
            Call SETSPOOL("FORECOLOR", str$(QBColor(0)))
'Call SETSPOOL("FORECOLOR", Str$(RGB(0, 0, 255)))
            FUENX$ = TRIM$(Mid$(RF$, 87, 32))
            If FUENX$ = "" Then
                Call COMUNI("Fuente No Definida para\Leyenda '" + TRIM$(Mid$(RF$, 9, 48)) + "'\en Formulario '" + TRIM$(CODFOR$) + "'.")
                FUEREPLA$ = "Courier"
                GoTo 103
            End If
            '
            For I9% = 1 To CAFUEN%
                If Fuente$(I9%) = FUENX$ Then GoTo 105
            Next I9%
            If FUEREPLA$ = "" Then
                FUEREPLA$ = "Courier"
                TTXX$ = "Fuente '" + FUENX$ + "' no Disponible\en el Presente Modelo de Impresora\para Leyenda '" + TRIM$(Mid$(RF$, 9, 48)) + "'\en Formulario '" + TRIM$(CODFOR$) + "'."
                TTXX$ = TTXX$ + "\   \FLEXOFT Reemplazará las Fuentes\no Disponibles por '" + FUEREPLA$ + "'."
                Call COMUNI(TTXX$)
            End If
103         FUENX$ = FUEREPLA$
            '
105         'Printer.FontName = FUENX$
            Call SETSPOOL("FONTNAME", FUENX$)
            'Printer.FontItalic = False
            Call SETSPOOL("FONTITALIC", "0")
            TAFUENTE = CVS(Mid$(RF$, 119, 4))
            If TAFUENTE > 2 Then
                'Printer.FontSize = TAFUENTE
                Call SETSPOOL("FONTSIZE", str$(TAFUENTE))
            End If
            If Asc(Mid$(RF$, 123, 1)) > 32 Then
                'Printer.FontBold = True
                Call SETSPOOL("FONTBOLD", "1")
              Else
                'Printer.FontBold = False
                Call SETSPOOL("FONTBOLD", "0")
            End If
            'Printer.CurrentX = (CVS(Mid$(RF$, 57, 4)) + MARIZW) * COEFH
            Call SETSPOOL("CURRENTX", str$((CVS(Mid$(RF$, 57, 4)) + MARIZW) * COEFH))
            'Printer.CurrentY = (CVS(Mid$(RF$, 61, 4)) + MARSUW) * COEFV
            Call SETSPOOL("CURRENTY", str$((CVS(Mid$(RF$, 61, 4)) + MARSUW) * COEFV))
            'Printer.Print TRIM$(Mid$(RF$, 9, 48))
            Call SETSPOOL("PRINT", TXTX$)
            '
        End If
109 Next i%
    '
    If ATRIFORM("", "CAR-OBSE") = "SI" Then
       Call CIERRARCH("!OBSERVOF")
       Call SETULTREG("!OBSERVOF", 0)
       FORMATOBS$ = TRIM$(UCase$(ATRIFORM$(CODFOR$, "OBSERVA/FORMATO")))
       ANCHOMA = XVALO(Mid$(FORMATOBS$, 2))
       PPXYZ% = InStr(FORMATOBS$, "/"): If PPXYZ% < 1 Then PPXYZ% = 1 + Len(FORMATOBS$)
       CARENMA& = XVALO(Mid$(FORMATOBS$, PPXYZ% + 1))
       If CARENMA& < 1 Then CARENMA& = 4
       If ANCHOMA > 0 Then
         If CARENMA& > 0 Then
           On Error Resume Next
           ENTRAOBS.TEPRUEBA = String$(ANCHOMA, "*")
           If Err Then
             On Error GoTo 0
             GoTo 114
           End If
           ENTRAOBS.Caption = UCase$(TIPODOC$) + " - " + NUMEDOC$
           ENTRAOBS.Label1 = ""
           ENTRAOBS.Text28.Width = 270 + ENTRAOBS.TEPRUEBA.Width
           ENTRAOBS.Label24.Width = ENTRAOBS.Text28.Width
           ENTRAOBS.Picture2.Width = ENTRAOBS.Text28.Width
           ENTRAOBS.Picture5.Left = ENTRAOBS.Text28.Width + 2 * ENTRAOBS.Text28.Left + 120
           ENTRAOBS.Width = ENTRAOBS.Picture5.Left + 1100
           ENTRAOBS.Command5.Left = ENTRAOBS.Text28.Left + ENTRAOBS.Text28.Width - ENTRAOBS.Command5.Width
           ENTRAOBS.Command6.Left = ENTRAOBS.Command5.Left - ENTRAOBS.Command6.Width
           ENTRAOBS.Command8.Left = ENTRAOBS.Command6.Left - ENTRAOBS.Command8.Width
           ENTRAOBS.Command17.Left = ENTRAOBS.Command8.Left - ENTRAOBS.Command17.Width
           ENTRAOBS.Command7.Width = ENTRAOBS.Command17.Left - ENTRAOBS.Command7.Left
           Screen.MousePointer = 1
           ENTRAOBS.Show 1
           If ENTRAOBS.Label1.Caption <> "OK" Then
             ENTRAOBS.Text28 = ""
           End If
           ANCHOMAXU% = ANCHOMA
           Call FRATEXTO(ENTRAOBS.Text28, ANCHOMAXU%)
           For KU& = 1 To CARETEX%
             Call GRAREG("!OBSERVOF", RETEX$(KU&), KU&)
           Next KU&
           Call SETULTREG("!OBSERVOF", KU& - 1)
           If CARENMA& < KU& Then
             Call SETULTREG("!OBSERVOF", CARENMA&)
           End If
           GoTo 115
         End If
       End If
114    VALIDA% = VENTABU%("OBSERVOF/TITUPAN=" + UCase$(TIPODOC$) + " - " + NUMEDOC$)
    End If
    '
115 RUTARGRA$ = TRIM$(CONTROL$("", "RUTARGRA"))
    If RUTARGRA$ = "" Then RUTARGRA$ = TRIM$(LUDAT$)
    If Right$(RUTARGRA$, 1) <> "\" Then
       RUTARGRA$ = RUTARGRA$ + "\"
    End If
    '
    ISOLOGO$ = TRIM$(CONTROL$("", "ISOLOGO"))
    If ISOLOGO$ <> "" Then
       ARCHIMAG$ = RUTARGRA$ + ISOLOGO$
       If EXISTE%(ARCHIMAG$) = 1 Then
         Call PRINTIMAGE("ISOLOGO", ARCHIMAG$)
       End If
    End If
    '
    CAIMAGEN% = 0    ' AGREGADO EL 14/04/2009 PORQUE LA IMPRIMIR VARIAS COPIAS NO PONIA IMAGEN A CERO EN CADA UNA
    For i% = 1 To LOF(NX%) / 128
      Get #NX%, i%, RF$
      If TRIM$(Left$(RF$, 8)) = "IMAGEWIN" Then
        ARIMA$ = TRIM$(Mid$(RF$, 9, 48))
        If InStr(ARIMA$, ".") > 0 Then
            ARCHIMAG$ = RUTARGRA$ + ARIMA$
            If EXISTE%(ARCHIMAG$) = 1 Then
              CAIMAGEN% = CAIMAGEN% + 1
              CODIMAGEN$(CAIMINI% + CAIMAGEN%) = ARIMA$
              ARCHIMAGEN$(CAIMINI% + CAIMAGEN%) = ARCHIMAG$
            End If
          ElseIf ARIMA$ <> "" Then
            CAIMAGEN% = CAIMAGEN% + 1
            CODIMAGEN$(CAIMINI% + CAIMAGEN%) = ARIMA$
            ARCHIMAGEN$(CAIMINI% + CAIMAGEN%) = ""
        End If
      End If
    Next i%
    '
    Close #NX%: NX% = 0
    On Error GoTo 0
    '
End Sub
'
Sub LIMPIAVALPAR()
    '
    For KKII% = 1 To CAVALPAR%
       CODPAR$(KKII%) = ""
       VALPAR$(KKII%) = ""
    Next KKII%
    CAVALPAR% = 0
    '
End Sub

Sub LIMPIATETAB()
    '
    For KKI% = 0 To CACOLTAB1%
      For KKJ% = 0 To CAITAB1%
        TETABU1$(KKI%, KKJ%) = ""
      Next KKJ%
    Next KKI%
    '
    For KKI% = 0 To CACOLTAB2%
      For KKJ% = 0 To CAITAB2%
        TETABU2$(KKI%, KKJ%) = ""
      Next KKJ%
    Next KKI%
    '
End Sub

Sub DIBUIMAG(DESTINIMAG As Object, ARIMAG$, HH11, HH22, VV11, VV22)
    '
    On Error Resume Next
    FORMALTER.Image3.Picture = LoadPicture("")
    FORMALTER.Image3.Picture = LoadPicture(ARIMAG$)
    HIMAGEN = FORMALTER.Image3.Picture.Width / 56.7
    VIMAGEN = FORMALTER.Image3.Picture.Height / 56.7
    '
    HCUADRO = HH22 - HH11
    VCUADRO = VV22 - VV11
    '
    REDUH = HIMAGEN / HCUADRO
    REDUV = VIMAGEN / VCUADRO
    REDUC = REDUH: If REDUV > REDUH Then REDUC = REDUV
    VV0 = VV11 + ((VV22 - VV11) - (VCUADRO * REDUV / REDUC)) / 2
    HH0 = HH11 + ((HH22 - HH11) - (HCUADRO * REDUH / REDUC)) / 2
    '
    DESTINIMAG.PaintPicture FORMALTER.Image3.Picture, 56.7 * HH0, 56.7 * VV0, 56.7 * HIMAGEN / REDUC, 56.7 * VIMAGEN / REDUC
    On Error GoTo 0
    '
End Sub

Sub refresca_botones(FRM As Form)
   Dim CONTROLX As CONTROL 'REFRESCA TODOS LOS CONTROLES
   Dim TIPX As Boolean
   For Each CONTROLX In FRM.Controls ' RECORRO LOS CONTROLES DEL FORM ACTIVO
         On Error Resume Next
          If CONTROLX.Visible = True Then
             CONTROLX.Visible = False
             CONTROLX.Visible = True
          End If
          On Error GoTo 0
20 Next
End Sub

