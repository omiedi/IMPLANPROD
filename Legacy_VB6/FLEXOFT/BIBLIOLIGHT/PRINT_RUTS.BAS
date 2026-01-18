Attribute VB_Name = "PRINT_RUTS"
Dim FUFIJA$(512)
Dim PPZZ As String * 128

Dim FUENTESEL
Dim FUENPRUE
Dim NADAFUEN
'

Sub SETPRINTER(ATOT%)
    '
    If Printer.CurrentX > 0 Then
        Printer.NewPage
    End If
    '
10  NOMIMPRE$ = TRIM$(Printer.DeviceName)
    If NOMIMPRE$ = "" Then
        Call MENSERR(24, "Falta Seleccionar Impresora\en la Configuración de Windows")
        Call FINAL("")
    End If
    '
    YAESTA% = 0
    CAFUFIJ% = 0
    NX% = FILEOPEN(LUCOM$ + "PRINTERS.DRV", 0, 128)
    FIN& = LOF(NX%) / 128
    For IPRIN& = 1 To FIN&
        Get #NX%, IPRIN&, PPZZ$
        PPXX$ = PPZZ$
        If TRIM$(Left$(PPXX$, 48)) = Left$(NOMIMPRE$, 48) Then
            YAESTA% = 1
            If UCase$(Mid$(PPXX$, 125, 2)) = "SI" Then
                CAFUFIJ% = CAFUFIJ% + 1
                FUFIJA$(CAFUFIJ%) = PPZZ$
            End If
        End If
    Next IPRIN&
    '
    Close #NX%
    If YAESTA% < 1 Then
        OPX% = COMALTER%("Se ha Detectado un Modelo de Impresora\Desconocido para el Sistema Aplicativo\  \'" + NOMIMPRE$ + "'\  \\Configurar\Cancelar")
        If OPX% <> 1 Then Call FINAL("")
        Call SETUPRINTER
        GoTo 10
    End If
    '
    If CAFUFIJ% < 1 Then
        OPX% = COMALTER%("No hay Fuentes Utilizables para Impresora\  \'" + NOMIMPRE$ + "'\  \\Configurar\Cancelar")
        If OPX% <> 1 Then Call FINAL("")
        Call SETUPRINTER
        GoTo 10
    End If
    '
    'Printer.Orientation = 1
    Printer.ScaleMode = 1
    ANCHOPAP = Int(Printer.ScaleWidth / 56.7)
    LARGOPAP = Int(Printer.ScaleHeight / 56.7)
    MARGENI = 0
    MARGENS = 0
    ORIENTA% = 1
    CACOPIAS% = 1
    RENULIS% = 0
    '
    Printer.ScaleMode = 1
    ANCHOJA = 56.7 * (ANCHOPAP - MARGENI)
    If ORIENTA% = 2 Then ANCHOJA = 56.7 * (LARGOPAP - MARGENI)
    '
    ATEXMAX = 0
    TAFUESEL = 0
    TALEBASI = 24
    '
100     TEPRUEBA$ = String$(ATOT%, "*")
    Printer.FontBold = False
    Printer.FontItalic = False
    For I% = 1 To CAFUFIJ%
        '
        Printer.FontName = TRIM$(Mid$(FUFIJA$(I%), 65, 48))
        Printer.FontBold = False
        Printer.FontItalic = False
        '
        FUENPRUE = TRIM$(Mid$(FUFIJA$(I%), 65, 48))
        If Printer.FontName <> FUENPRUE Then
            GoTo 109   ' modificado 20/11/01 para que busque la fuente
                       ' disponible
            Call COMUNI("Se ha Detectado un Error\en la Configuración Local\de la Impresora Predeterminada.\  \Imposible Emitir el Listado Solicitado.")
            Call CIERRARCH("*.*")
            Close
            Call FILEKILL(LUCOM$ + "PRINTERS.DRV")
            Call FINAL("")
        End If
        '
        FUENTEMIN = CVS(Mid$(FUFIJA$(I%), 113, 4))
        FUENTEMAX = CVS(Mid$(FUFIJA$(I%), 117, 4))
        '
        TALEBA = FUENTEMAX + 0.4
        While TALEBA > FUENTEMIN
           TALEBA = TALEBA - 0.4
           Printer.FontSize = TALEBA
           If Printer.FontSize > TALEBA Then GoTo 105
           '
           On Error Resume Next
           ATEXTO = Printer.TextWidth(TEPRUEBA$)
           If Err Then
             On Error GoTo 0
             GoTo 105
           End If
           On Error GoTo 0
           If ATEXTO <= ANCHOJA Then
              If ATEXTO > ATEXMAX Then
                FUENTESEL = FUENPRUE    ' Printer.FontName
                TAFUESEL = Printer.FontSize ' TALEBASI
                ATEXMAX = ATEXTO
              End If
              GoTo 109
           End If
           TALEBA = Printer.FontSize
105         Wend
    '
109     Next I%
    '
    If FUENTESEL = NADAFUEN Or TAFUESEL < 1 Then
        Call MENSERR(24, "No hay Fuentes Adecuadas\para la Impresion de Este Listado")
        Call FINAL("")
    End If
    '
    Printer.FontName = FUENTESEL
    Printer.FontSize = TAFUESEL
   '
    If FUENTESEL = "" Or TAFUESEL < 0.4 Then
        Call MENSERR(24, "No puede Establecerse\el Formato de Impresion.\Consulte al Soporte de Sistemas.")
        Call FINAL("")
    End If
    '
110     Printer.FontName = FUENTESEL
    Printer.FontSize = TAFUESEL
    LARGOJA = Val(Printer.ScaleHeight)
    TEPRUEBB$ = TEPRUEBA$ ' : If LX810 = 1 Then TEPRUEBB$ = TEPRUEBA$ + "*****"
    ANCHOJB = Val(Printer.TextWidth(TEPRUEBB$))
    TALEBASI = Printer.FontSize
    ALTUREN = Printer.TextHeight(TEPRUEBA$)
    Printer.FontBold = False
    Printer.FontItalic = False
    MIZMAX = ANCHOJA - ANCHOJB
    If MARGENI > 0 Then
        MIZMAX = 56.7 * MARGENI
    End If
    MIZQ$ = ""
    While Printer.TextWidth(MIZQ$ + "*") <= MIZMAX
       MIZQ$ = MIZQ$ + " "
    Wend
    MIZQ$ = Mid$(MIZQ$, 2)
    '
End Sub
'
Sub SETUPRINTER()
    '
    NOMIMPRE$ = TRIM$(Printer.DeviceName)
    If NOMIMPRE$ = "" Then
        Call MENSERR(24, "Falta Seleccionar Impresora\en la Configuración de Windows")
        Call FINAL("")
    End If
    '
    YAESTA% = 0
    Call BLANARCH("!PRINTERS")
    NOMI$ = AJUSTI$(NOMIMPRE$, 48)
    NX% = FILEOPEN(LUCOM$ + "PRINTERS.DRV", 0, 128)
    For IPRIN& = 1 To LOF(NX%) / 128
        Get #NX%, IPRIN&, PPZZ$
        PPXX$ = PPZZ$
        If TRIM$(Left$(PPXX$, 48)) = NOMIMPRE$ Then
            Call REGAPP("!PRINTERS", PPXX$)
            YAESTA% = 1
            GoTo 10
        End If
    Next IPRIN&
    '
10  Close #NX%
    If YAESTA% = 1 Then
        PPXX$ = Left$(PPXX$, 52)
        'GoTo 50
    End If
    '
    Printer.ScaleMode = 1
    ANCHOJMM = Int(Printer.ScaleWidth / 56.7)
    Call REPLA(PPXX$, MKS$(ANCHOJMM), 49, 4)
    '
    CAFUFIJ% = 0
    For I% = 1 To Printer.FontCount
        Printer.FontName = Printer.Fonts(I% - 1)
        A1 = 0: A2 = 0
        On Error Resume Next
        A1 = Printer.TextWidth(String$(80, "l"))
        A2 = Printer.TextWidth(String$(80, "M"))
        On Error GoTo 0
        If A1 > 0 And A2 > 0 Then
          If A1 = A2 Then
            If CAFUFIJ% < 128 Then
              CAFUFIJ% = CAFUFIJ% + 1
              FUFIJA$(CAFUFIJ%) = Printer.Fonts(I% - 1)
            End If
          End If
        End If
    Next I%
    '
    If CAFUFIJ% < 1 Then
       Call MENSERR(24, "No se Encontraron Fuentes de Ancho Fijo\para la Impresora Seleccionada.\  \No es Posible Emitir Listado por Impresora.\  \Consulte a su Soporte de Sistemas.")
       Call FINAL("")
    End If
    '
    Call BLANARCH("!PRINTERS")
    TALEBASI = 24
    TEPRUEBA$ = String$(128, "*")
    Printer.FontBold = False
    Printer.FontItalic = False
    For I% = 1 To CAFUFIJ%
        Printer.FontName = FUFIJA$(I%)
        If Printer.FontName <> FUFIJA$(I%) Then GoTo 19
        '
        Printer.FontSize = 24
        FUENTEMAX = Printer.FontSize
        Printer.FontSize = 2
        FUENTEMIN = Printer.FontSize
        '
14      If FUENTEMIN <= FUENTEMAX Then
            XX$ = REGBLAN$("!PRINTERS")
            Call REPLA(XX$, NOMIMPRE$, 1, 48)
            Call REPLA(XX$, MKS$(ANCHOJMM), 49, 4)
            Call REPLA(XX$, FUFIJA$(I%), 65, 48)
            Call REPLA(XX$, MKS$(FUENTEMIN), 113, 4)
            Call REPLA(XX$, MKS$(FUENTEMAX), 117, 4)
            Call REPLA(XX$, "SI", 125, 2)
            Call REGAPP("!PRINTERS", XX$)
        End If
        '
19  Next I%
    Call CIERRARCH("!PRINTERS")
    '
    PPXX$ = AJUSTI$(NOMIMPRE$, 48) + MKS$(ANCHOJMM) + String$(12, 0)
    '
50  PPYY$ = OBJPLA$("DATIMPRE", PPXX$)
    If PPYY$ = "" Then
        Call CIERRARCH("PRINTERS")
        GoTo 99
    End If
    '
    ANCHOJMM = CVS(Mid$(PPYY$, 49, 8))
    If ANCHOJMM > Int(Printer.ScaleWidth / 56.7) Or ANCHOJMM < 48 Then
        PPXX$ = PPYY$
        GoTo 50
    End If
    '
    VTB% = VENTABU%("PRINSET/" + AJUSTI$(NOMIMPRE$, 48))
    If VTB% >= 0 Then
        NX% = FILEOPEN%(LUCOM$ + "PRINTERS.DRV", 0, 128)
        NOMI$ = AJUSTI$(NOMIMPRE$, 48)
        JJ& = 0
        For IPRIN& = 1 To LOF(NX%) / 128
            Get #NX%, IPRIN&, PPZZ$
            If Left$(PPZZ$, 48) <> NOMI$ Then
               JJ& = JJ& + 1
               Put #NX%, JJ&, PPZZ$
            End If
        Next IPRIN&
        For IPRIN& = 1 To ULTREG&("!PRINTERS")
            PPXX$ = REGLEIDO$("!PRINTERS", IPRIN&)
            Call REPLA(PPXX$, NOMI$, 1, 48)
            Call REPLA(PPXX$, MKS$(ANCHOJMM), 49, 4)
            If UCase$(Mid$(PPXX$, 125, 2)) <> "SI" Then
                Call REPLA(PPXX$, "  ", 125, 2)
            End If
            PPZZ$ = PPXX$
            JJ& = JJ& + 1
            Put #NX%, JJ&, PPZZ$
        Next IPRIN&
        For IPRIN& = JJ& + 1 To LOF(NX%) / 128
            PPZZ$ = String$(128, 0)
            Put #NX%, IPRIN&, PPZZ$
        Next IPRIN&
        Close #NX%
    End If
    Call CIERRARCH("!PRINTERS")
    '
99  End Sub
'

