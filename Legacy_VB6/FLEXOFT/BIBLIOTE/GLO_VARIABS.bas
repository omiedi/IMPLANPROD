Attribute VB_Name = "GLO_VARIABS"
Private Declare Sub GetSafeArrayPointer Lib "msvbvm60.dll" Alias "GetMem4" (pArray() As Any, ret As Long)

'Declaramos un Array dinámico
Dim UnArray() As Long
' Public FEHOR                       ' ESTO ES PARA ENLAZAR LA BIBLIOTECA FECHORAS
Public ACOLUM%(256), ALTUREMAX       ' ALTURA DE RENGLON AL IMPRIMIR IMAGENES O MULTILINEA
Public UTIMETEC
'
Public MIGRATODO%
Public REPRIFAC%        ' SE PONE EN 1 AL RE-IMPRIMIR UNA FACTURA YA EMITIDA
Public MUECOMHAS%       ' 1 / 2 Presenta parametros enviados al controlador fiscal
'
Public TIBASE$
Public RUTIACT$         ' RUTINA ACTIVA PARA CAPTURA DE ERROR
Public NUEWINTA%        ' SI ES = 1 ACTIVA WINTAB02
'
Public JOBID$           ' string identificador de tarea|
Public VERDEMO%         ' version de demostracion +/- 1
Public VERCD%           ' version  en CD 0/1
Public VENCIABO%        ' fecha vencimiento abono
Public VERANTER$        ' version anterior del modulo
Public APLIXYZ$
'
Public USNBR%           ' numero de usuario
Public USERNAME$        ' nombre de usuario activo
'
Public MENUPRIN$        ' menu de distribucion activo
Public MODINVO$
Public DENOMOD$
Public MODMENU$
Public APLINVO$, FILTX$, SUBAPLI$, PARAPLI$, LLAMOCONFILTROS%
Public ESFINAL%
'
Public EJER%            ' ejercicio actual / anterior
Public COMEJ%           ' fecha comienzo ejercicio activo
Public FINEJ%           ' fecha fin ejercicio activo
Public DENOEJ$
Public NEJER$
Public SE$              ' subempresa
Public EXTE$            ' extension 'DAT' o 'OLD"
Public LANGUA%
'
Public EMPREAC$, UNINEGO%
Public CodiEmp%
Public CodiSuc%
'
Public LUDAT$           ' U.Logica de datos
Public LUACC$           ' U.Logica de B.Datos Access
Public LUBAS$           ' U.Logica sistema basico
Public LUCOM$           ' U.Logica datos locales
Public LUBOOT$          ' U.Logica de boot
'
Public IDENTER$         ' identificador de terminal
Public NOMTER$          ' nombre de terminal
Public USERTER$
'
Public FLAGMOVI%(32767)
'
Public UCOD$(128), ICOMPO%(128), IUCOD%, SUBENI%
Public LLAMACOSTO$, ORILLAMA$, REGILLAMA%, TOPELLAMA%
Public DECICALCOS%
Public RECALCOSTO%, FINCALCOS%, ENPRORECA%
'
Public ATRA$
Public CCI$(32)
Public PROLLAMA$
'
Public LIPRACTI%
Public ALIVARNI
Public TICAMBIO
Public TAIVA(15)
Public MONEMI%, CONVERPRECIO%
Public VH(256)
Public TICACH(256)
'
Public Ci%, DE$, INCLUALIAS%
'
Public FORMA$, FORMALI$, ARCHITRA$, ARCHILOC$
Public COMMA$, COMMAX$, PROPRIN$, ATRIBU$
Public CARELI%
Public MODEDI%
Public TIFORTAB%
Public HDR&, SUPRIFE%, RFT%, PRX%, NCL%, NCLF%, TCL%, RG%, TNC%
Public ACL%(48), ECL$(48), ECLF$(48), cl%(48), lC%(48), LK%(48), MG%(48), JT%(48), TVL%(48), PR%(48), LF%(48), LDL%(48), PE%(48), ae$(48), CE%(48), CL0%(48), TT%(48)
Public APIX%(48)     ' ANCHO DE COLUMNA EN PIXELS
Public RTT%(48), CLT%(48), LCF%(48), RGFL&(48), CLF%(48), LKF%(48), LDF%(48), JTF%(48), CEF%(48), AEF$(48), LEF%(48), CFT%(48), CFF%(48), CRF%(48)
Public TTX#(64), TXX#(64)
Public CLNC%(48)
'
Public YAFIL%, ORDE%, NUEFORM%
Public LU$, DRX$
Public LCB$
Public HOO%, HOS$
Public GRABACT%
Public Archi$, PROCE$, LOR%, LOR1%, N15%
Public CAMAX&
Public PEOF%, EOFS$
Public s1$, sali$
Public COD$, RGI&, RGF&, OFS%, BLP%, VIV%, IDL%, IDL0%, REHOJX%, INTERLIX, CURRY, APPI%
Public BSUW&, BIZW&
Public APQ%
Public ATOT%, ATABU%, CLII%, KCL%, KCL1%, NFT%, RG0%
Public ENCA$, ENCACO$, COLU%
Public FCF%, RGFI%, RGFS%, HREN%
Public AMARCO%, HMARCO%, ALETRA%, HCATOT%
Public PX%, APLX$, CAPRO%
Public REACTIVO$, NUREGIS&
'
Public COEFH
'
Public VTB%
Public ORDECOP%, ORCOP$(10)
'
Public VACON%
Public PRINTERCONTROL%
Public CONTROLGRABA%
Public FUEREPLA$
Public SPOOLSTRING$
Public ORIPRINTER$
'
Public ATENTI%
Public MODIFIX%
Public UPDATCLI%
Public UPDATPRO%
Public APLIAUTO%
'
Public IDENSEDE$
Public DESTITRA$
Public ORITRA$(3)
'
Public HORIMAGEN, VERTIMAGEN, INDIMAGEN%

Public CUEVENTA$(18), COCUEVEN%(18)

Public CARFAPEN%
Public MODORECI%
Public FECHAVENCI(10), IMPOVENCI#(10), CAPLAVENCI%
Public TOTAFAC#(3, 10)

Public NUEVODOC%
Public DOCUACTIVO$
Public RESTAT$
'
Public PRISEMIN%
'
Public APLICA$
Public RETEX$(4096), CARETEX%
'
Public COMPLECALCU%
Public XMUOFA, YMUOFA
Public CLIPROACTIVO
'
Public AQUIMANTEN%
Public IMPREDET$
Public SENDBYMAIL$
'
Public COLTELIB%
'
Public NOCLOSE%
Public NOFRESHPRE%    ' CUANDO ESTÁ EN 0 (CERO), AL INVOCAR EN FORVENTB UN
'                     ' ARTÍCULO REFRESCA EL PRECIO SEGÚN LISTA ACTIVA
'                     ' SI ESTÁ EN 1 (CASO DE MODIFICAR PEDIDOS) NO LO REFRESCA
'
Public FORFACPRINT As Form  'VARIABLE PARA DECLARAR SI EL FORMULARIO VA A SER FORFACCS O FORFASQL EN EL FORMULARIO FORFACSQL O EN FORFACCS
Public FORFINFAC As Form
'
Public CODCARPE$      ' CODIGO DE CARPETA DE EXPORT / IMPORT

Public FORINSTRUC As Form  'VARIABLE PARA DECLARAR SI EL FORMULARIO INSTRUCCI
Public CANTICOPIAS& 'VARIABLE PARA USAR COMO CONTROL DE CANTIDAD DE COPIAS IMPRIMIR EN LA LLAMDA AL STDFORM

Public INDEXTEX%
Public UTILIZA_SKIN%
Public TXTRAZA$
Public REFRESCARPED%
Public RESP_ZELE_VECT(64)
Public CI_VECTOR_TIVA(256)
Public TASAIVAX_CI(32767)
Public VALIDA_TASAIVA$ 'TIENE EL VALOR DEL CONTROL SI SE VALIDA O NO LA TASA DE IVA POR COMBINACIONES.

Public CONSIDERAR_MOVIVIE% 'VARIABLE PARA QUE EL GENERADOR DE REPORTES CONSIDERE LA TABLA MOVIVIE Y MUESTRE TAMBIEN LOS MOVIMIENTOS DE ESTA
Public PINTAFILACLICK% 'SE UTILIZA EN FRMZELECHO PARA PINTAR  O NO LA FILA AL HACER CLICK
Public TXTRAZA_ARMADOR$ ' SE UTILIZA PARA TRAZAR LA RUTINA GRABAARMADOR
Public GRILLA_IMPRESION  As Object
Public VectorToolTipText(128) 'SE UTILIZA PARA FRM GRILLAGRAL LE ASIGNA UN TOOLTIPTEXT DIFERENTE SEGUN COLUMNA A LA GRILLA

' *** PARA UTILIZAR EN WEBCAM
' *******************************

' declaraciones Api, constantes, variables
Public Const ws_child = &H40000000
Public Const ws_visible = &H10000000
Public Const WM_USER = 1024
Public Const wm_cap_driver_connect = WM_USER + 10
Public Const wm_cap_set_preview = WM_USER + 50
Public Const WM_CAP_SET_PREVIEWRATE = WM_USER + 52
Public Const WM_CAP_DRIVER_DISCONNECT = WM_USER + 11
Public Const WM_CAP_DLG_VIDEOFORMAT = WM_USER + 41
Public Const WM_CAP_DLG_VIDEOCONFIG = WM_USER + 42
Public Const WM_CAP_SET_SCALE = WM_USER + 53

'Public Const WM_USER As Long = &H400
Public Const WM_CAP_START As Long = WM_USER
Public Const WM_CAP_EDIT_COPY As Long = WM_CAP_START + 30
Public Const WM_CAP_SET_VIDEOFORMAT As Long = WM_USER + 45
  
'Api para crear la ventana de captura
Public Declare Function capCreateCaptureWindow Lib "avicap32.dll" _
    Alias "capCreateCaptureWindowA" ( _
    ByVal lpszWindowName As String, _
    ByVal dwStyle As Long, _
    ByVal X As Long, _
    ByVal Y As Long, _
    ByVal nWidth As Long, _
    ByVal nHeight As Long, _
    ByVal hwndParent As Long, _
    ByVal nID As Long) As Long
  
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" ( _
    ByVal hwnd As Long, _
    ByVal wMsg As Long, _
    ByVal wParam As Long, _
    lParam As Any) As Long
  
Public Declare Function DestroyWindow Lib "user32" (ByVal hndw As Long) As Boolean
Public Declare Function SendMessageAsBitMap Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByRef lParam As BITMAPINFO) As Long

Public Type BITMAPINFOHEADER
    biSize As Long
    biWidth As Long
    biHeight As Long
    biPlanes As Integer
    biBitCount As Integer
    biCompression As Long
    biSizeImage As Long
    biXPelsPerMeter As Long
    biYPelsPerMeter As Long
    biClrUsed As Long
    biClrImportant As Long
End Type

Public Type BITMAPINFO
    bmiHeader As BITMAPINFOHEADER
    bmiColors() As Integer
End Type
    
Public hwdc As Long
Public startcap As Integer

' *** FIN PARA WEBCAM
' ***********************

' *** MODIFICAR TAMAÑO DE FOTOS
' *****************************

Public Const IMAGE_BITMAP = 0
Public Const LR_COPYRETURNORG = &H4
Public Const CF_BITMAP = 2
Public Declare Function SetClipboardData Lib "user32" (ByVal wFormat As Long, ByVal hMem As Long) As Long
Public Declare Function CopyImage Lib "user32" (ByVal handle As Long, ByVal imageType As Long, ByVal NewWidth As Long, ByVal NewHeight As Long, ByVal lFlags As Long) As Long
Public Declare Function EmptyClipboard Lib "user32" () As Long
Public Declare Function CloseClipboard Lib "user32" () As Long
Public Declare Function OpenClipboard Lib "user32" (ByVal hwnd As Long) As Long

' *** FIN MODIFICAR TAMAÑO DE FOTOS
' *********************************

' LO QUE SIGUE ES PARA TRAER AL FRENTE UN FORMULARIO. LAS ULTIMAS DOS LINEAS SON EL EJEMPLO DE USO.
Public Const HWND_TOPMOST = -1
Public Const HWND_NOTOPMOST = -2
Public Const SWP_NOSIZE = &H1
Public Const SWP_NOMOVE = &H2
Public Const SWP_NOACTIVATE = &H10
Public Const SWP_SHOWWINDOW = &H40
Declare Sub SetWindowPos Lib "user32" ( _
    ByVal hwnd As Long, _
    ByVal hWndInsertAfter As Long, _
    ByVal X As Long, _
    ByVal Y As Long, _
    ByVal cx As Long, _
    ByVal cy As Long, _
    ByVal wFlags As Long)
'SetWindowPos Me.hWnd, HWND_TOPMOST, 0, 0, 0, 0, _
'                 SWP_NOACTIVATE Or SWP_SHOWWINDOW Or SWP_NOMOVE Or SWP_NOSIZE

Public Type AtributosImpresionEtiqueta
    LOTE As String
    FECHA As Integer
    CodigoInterno As Integer
    Largo As Double
    Peso As Double
    NumeroProveedor As Long
    Remito As String
    Cantidad As Double
    textoComentario As String
    cantidadBultos As Integer
    IncluyeQR As Boolean
    ModoPrint As String
End Type

Type ItemPicking
    OrdenFabricacion As String
    Partida As String
    CodigoInterno As Integer
    Pallet As Long
    Bulto As Long
    Cantidad As Single
End Type

Function CAPTURAERRORSKIN%(FRM As Form)
   
   RUTXX$ = LUBAS$
   If INTERPRE% > 0 Then RUTXX$ = LUDAT$

   ARCHIVOSKIN$ = TRIM$(Left$(LOADFILE$(RUTXX$ + "FLEX_SKIN.CFG"), 512))
   On Error Resume Next
     A = FRM.SkinLabel999(0)
     FRM.Skin1.LoadSkin ARCHIVOSKIN$
     If Err > 0 Then CAPTURAERRORSKIN% = 1
   On Error GoTo 0

End Function

'
Function CRLF$()
   CRLF$ = Chr$(13) + Chr$(10)
End Function
'
Function SWAP(VX1, VX2)
   VX3 = VX2
   VX2 = VX1
   VX1 = VX3
End Function
'
Function HABILAPLI%(APLICA$)
   HABILAPLI% = 0
Exit Function
   FEVEN% = FECHANUM("17/10/10")
   If HOY(HO$) > FEVEN% Then
     HABILAPLI% = (-1)
     SEVOLF$ = SERIVOL$(Left$(App.Path, 2))
     TTXX$ = "La Aplicación Invocada:"
     If APLICA$ = "SHELLEXEC" Then
         TTXX$ = TTXX$ + "\   \'Presentación de POE's y Conexion con Documentos Vinculados'"
         TTXX$ = TTXX$ + "\   \Excede el Alcance de la Presente Instalación (" + SEVOLF$ + ")."
         TTXX$ = TTXX$ + "\   \Para Conocer esta Aplicación y Acceder al Licenciamiento Correspondiente,"
         TTXX$ = TTXX$ + "\Debe Contactar a FLEXOFT S.A.:"
         TTXX$ = TTXX$ + "\   \info@flexoft.com.ar - 5424-1284 / 4522-2048 / 4524-0796"
         TTXX$ = TTXX$ + "\   \Muchas gracias por su Amable Atencion."
         Call MENSERR(24, TTXX$)
       ElseIf APLICA$ = "GETPICTURE" Then
         TTXX$ = TTXX$ + "\   \'Captura de Imagenes en Línea mediante Camara o Scanner'"
         TTXX$ = TTXX$ + "\   \Excede el Alcance de la Presente Instalación (" + SEVOLF$ + ")."
         TTXX$ = TTXX$ + "\   \Para Conocer esta Aplicación y Acceder al Licenciamiento Correspondiente,"
         TTXX$ = TTXX$ + "\Debe Contactar a FLEXOFT S.A.:"
         TTXX$ = TTXX$ + "\   \info@flexoft.com.ar - 5424-1284 / 4522-2048 / 4524-0796"
         TTXX$ = TTXX$ + "\   \Muchas gracias por su Amable Atencion."
         Call MENSERR(24, TTXX$)
     End If
   End If
End Function
'
Function VERILICEN%(CAPLI$, Optional DAPLI$)
   '
   Static SEVOF$
   VRL% = 0
   LUAPP$ = App.Path + "\"
   If TRIM$(CAPLI$) = "" Then
      Call MENSERR(24, "Error de Invocacion de Funcion VERILICEN - Consulte.")
      Exit Function
   End If
   '
   If SEVOF$ = "" Then SEVOF$ = SERIVOL$(Left$(App.Path, 2))
   '
   If EXISTE%(LUAPP$ + CAPLI$ + ".HBL") > 0 Then
      XXX$ = LOADFILE$(LUAPP$ + CAPLI$ + ".HBL")
      TORI$ = Left$(ENCRIPTA$(AJUSTI$(SEVOF$, 16) + AJUSTI$(CAPLI$, 112)), 126)
      If Left$(XXX$, 126) = TORI$ Then
         VRL% = 1: GoTo 90
      End If
   End If
   '
   TTXX$ = "La Aplicación Invocada: "
      DECOAP$ = CAPLI$:   If TRIM$(DAPLI$) <> "" Then DECOAP$ = CAPLI$ + " - " + DAPLI$
   TTXX$ = TTXX$ + "\   \'" + DECOAP$ + "'"
   TTXX$ = TTXX$ + "\   \No se Encuentra Licenciada para la Presente Instalación (" + SEVOF$ + ")."
   TTXX$ = TTXX$ + "\   \Para Conocer esta Aplicación y Acceder al Licenciamiento Correspondiente,"
   TTXX$ = TTXX$ + "\Debe Contactar a FLEXOFT S.A.:"
   TTXX$ = TTXX$ + "\   \info@flexoft.com.ar - 5424-1284 / 4522-2048 / 4524-0796"
   TTXX$ = TTXX$ + "\   \Muchas gracias por su Amable Atencion."
   Call MENSERR(24, TTXX$)
   '
90 VERILICEN% = VRL%
   '
End Function
'
Sub CARDATCENCALC()
    '
    OPPX% = COMALTER%("Elija Opción de Actualización:\\Tasas de IVA\Vencimientos de CAI")
    If OPPX% < 1 Or OPPX% > 2 Then Exit Sub
    If OPPX% = 2 Then GoTo 20
    '
    LU$ = LUDAT$
    Dim FF0 As String * 96
    '
    n1% = FILEOPEN%(LU$ + "FORMFAC.DAT", 0, 96)

    If LOF(n1%) >= 192 Then
        Get #n1%, 2, FF0$
        VDEF$ = FF0$
      Else
        VDEF$ = String$(96, 0)
    End If
    '
    X$ = OBJPLA$("CARDATCENCALC", VDEF$)
    If X$ <> "" Then
        '
        ALIVARNI = CVS(Mid$(X$, 7, 4))
        '
        If LOF(n1%) < 92 Then
            FF0$ = String$(96, 0)
            Put #n1%, 1, FF0$
        End If
        FF0$ = X$
        Put #n1%, 2, FF0$
    End If
    '
    Close #n1%: n1% = 0
    Exit Sub
    '
20  Call TRALOCAL("VENCICAI", "")
    VTB% = VENTABU%("VENCICAI/NC/TITUPAN=Tabla de Vencimientos de CAI")
    If VTB% >= 0 Then
      JJ& = 0
      For U& = 1 To ULTREG&("!VENCICAI")
        X$ = REGLEIDO$("!VENCICAI", U&)
        If TRIM$(Left$(X$, 32)) <> "" Then
          JJ& = JJ& + 1
          Call GRAREG("VENCICAI", X$, JJ&)
        End If
      Next U&
      Call SETULTREG("VENCICAI", JJ&)
    End If
    Call CIERRARCH("VENCICAI")
    Call CIERRARCH("!VENCICAI")
    '
End Sub
'

Function ALIVA()
    '
    Static ALIVAA
    '
    If ALIVAA < 1 Or TICAMBIO < 0.1 Then
10     If ULTREG&("FORMFAC") >= 2 Then
            FFS$ = REGLEIDO$("FORMFAC", 2&)
            ALIVAA = CVS(Mid$(FFS$, 2, 4))
            ALI2 = CVS(Mid$(FFS$, 7, 4))
            ALIVARNI = ALI2
            TC1 = CVS(Mid$(FFS$, 11, 4))
            TICAMBIO = TC1: If TICAMBIO < 1 Then TICAMBIO = 1
            '
            TFOR% = Asc(Mid$(FFS$, 6, 1))
            CAFOR% = 2: If TFOR% > 10 Then CAFOR% = (TFOR% Mod 10): TFOR% = Int(TFOR% / 10)
            MPRI$ = TRIM$(Mid$(FFS$, 24, 30))
            MCEN$ = TRIM$(Mid$(FFS$, 54, 30))
            '
            If TFOR% < 1 Or TFOR% > 2 Then TFOR% = 1
         Else
            Call CARDATCENCALC
            GoTo 10
       End If
    End If
    '
    ALIVA = ALIVAA
    '
End Function

Function VERNUM$()     ' revision del 23/09/03
   '
   YY$ = ""
   On Error Resume Next
   IKXX$ = FORSELEC.Label99.Caption
   If Err Then GoTo 3
   '
   If Mid$(IKXX$, 51, 9) <> "........." Then
     IKYY$ = Mid$(IKXX$, 51, 9)
     '
     VERDEMO% = 1
     '
     ' DIGITO VERIFICADOR DEL SECTOR DE CONTROL
     DIVERI& = 0
     For KKI% = 1 To 8
       KKJ% = KKI%: If KKI% > 4 Then KKJ% = KKI% + 1
       DIVERI& = XVALO(Mid$(IKYY$, KKJ%, 1)) + (7 * KKJ% ^ 2)
     Next KKI%
     DIVERI& = DIVERI& Mod 10
     If XVALO(Mid$(IKYY$, 5, 1)) <> DIVERI& Then
       MsgBox "Módulo '" + App.EXEName + "' Indebidamente Alterado (D.Ver)"
       Close: End: Exit Function
     End If
     '
     ' FECHA DEL MóDULO
     On Error Resume Next
     FEMODUI$ = TRIM$(Str$(CVINT(FileDateTime(App.EXEName + ".EXE"))))
     While Len(FEMODUI$) > 4: FEMODUI$ = Mid$(FEMODUI$, 2): Wend
     FINAJ$ = AJUSTI$(TRIM$(UCase$(App.EXEName)) + ".EXE", 8)
     FEMODUJ$ = ""
     For KKI% = 1 To 4
       CARAJ$ = TRIM$(Str$((XVALO(Mid$(FEMODUI$, KKI%, 1)) + Asc(Mid$(FINAJ$, KKI%, 1))) Mod 10))
       FEMODUJ$ = FEMODUJ$ + CARAJ$
     Next KKI%
     If Left$(IKYY$, 4) <> FEMODUJ$ Then
       MsgBox "Módulo '" + App.EXEName + "' Indebidamente Alterado (F.Gen)"
       Close: End: Exit Function
     End If
     FEGENER = 10000 + XVALO(Left$(FEMODUI$, 4))
     '
     ' VERIFICACION FECHA DE VENCIMIENTO
     FEVENCJ$ = Mid$(IKYY$, 6, 4)
     FEVENCX$ = ""
     For KKI% = 1 To 4
       CARAL& = XVALO(Mid$(FEVENCJ$, KKI%, 1)) - Asc(Mid$(FINAJ$, KKI% + 4, 1))
       While CARAL& < 0: CARAL& = CARAL& + 10: Wend
       FEVENCX$ = FEVENCX$ + TRIM$(Str$(CARAL&))
     Next KKI%
     FeVencI = 10000 + XVALO(FEVENCX$)
     If DIENTRE%(FEGENER, FeVencI) > 180 Then
       FeVencI = FEGENER
       While DIENTRE(FEGENER, FeVencI) < 180
         FeVencI = DIPOST(FeVencI)
       Wend
     End If
     '
     'HACE CAER FICHA DE BLOQUEO
     HOII% = HOY(HOS$)
     If HOII% > FeVencI Then
       On Error Resume Next
       FileCopy App.EXEName + ".EXE", "C:\FLEXOFT\" + App.EXEName + ".WKX"
       ABC$ = LOADFILE$("FLOGON.EXE")
       If Len(ABC$) < 129 Then ABC$ = RANDSTRING$(28345)
       If EXISTE%("C:\WINDOWS\SYSTEM\.") > 0 Then Call SAVEFILE("C:\WINDOWS\SYSTEM\C0MCTL32.0CX", ABC$)
       If EXISTE%("C:\WINDOWS\SYSTEM32\.") > 0 Then Call SAVEFILE("C:\WINDOWS\SYSTEM32\C0MCTL32.0CX", ABC$)
       If EXISTE%("C:\WINNT\SYSTEM\.") > 0 Then Call SAVEFILE("C:\WINNT\SYSTEM\C0MCTL32.0CX", ABC$)
       If EXISTE%("C:\WINNT\SYSTEM32\.") > 0 Then Call SAVEFILE("C:\WINNT\SYSTEM32\C0MCTL32.0CX", ABC$)
     End If
     '
     XK1% = EXISTE%("C:\FLEXOFT\" + App.EXEName + ".WKX")
     XK1% = XK1% + EXISTE%("C:\WINDOWS\SYSTEM\C0MCTL32.0CX")
     XK1% = XK1% + EXISTE%("C:\WINDOWS\SYSTEM\C0MCTL32.0CX")
     XK1% = XK1% + EXISTE%("C:\WINDOWS\SYSTEM\C0MCTL32.0CX")
     XK1% = XK1% + EXISTE%("C:\WINDOWS\SYSTEM\C0MCTL32.0CX")
     '
     If XK1% > 0 Then
        TTXX$ = "Plazo de  Validez  de  Demostración Vencido" + Chr$(13) + Chr$(10)
        TTXX$ = TTXX$ + Chr$(13) + Chr$(10)
        TTXX$ = TTXX$ + "Para Obtener su Licencia de Uso, Contacte a" + Chr$(13) + Chr$(10)
        TTXX$ = TTXX$ + "FLEXOFT  Soluciones Informaticas Integrales" + Chr$(13) + Chr$(10)
        TTXX$ = TTXX$ + "info@flexoft.com.ar - Tel:  54-11-4524-1284" + Chr$(13) + Chr$(10)
        TTXX$ = TTXX$ + "              Buenos Aires - Argentina"
        MsgBox TTXX$
        Close: End: Exit Function
     End If
     '
     MsgBox "Módulo '" + App.EXEName + "' - Versión Demo"
     '
   End If
   '
   On Error GoTo 0
   YY$ = Mid$(IKXX$, 41, 10)
   If YY$ <> ".........." Then
     If YY$ <> "" Then
       GoTo 90
     End If
   End If
   '
3  On Error GoTo 0
   YY$ = ""
   If EXISTE%(App.EXEName + ".EXE") = 1 Then
     YY$ = Format$(FileDateTime(App.EXEName + ".EXE"), "yy.m.d")
     If Left$(YY$, 1) = "0" Then YY$ = Mid$(YY$, 2)
     YY$ = "V." + YY$
   End If
   '
90 VERNUM$ = YY$
   '
End Function
'
Function FECHMODU%()
   FMD% = HOY(HOS$)
   If EXISTE%(App.EXEName + ".EXE") = 1 Then
     YY$ = Format$(FileDateTime(App.EXEName + ".EXE"), "dd/mm/yy")
     FMD% = FECHANUM(YY$)
   End If
   FECHMODU% = FMD%
End Function
'
Sub CARUTRA()
   '
   IDENSEDE$ = Control$("", "IDENSEDE")
   DESTITRA$ = Control$("", "DESTITRA")
   '
   v& = (-1)
   For U& = 0 To 3
     ORITRA$(U&) = ""
     ORITRX$ = Control$("", "ORITRA-" + TRIM$(Str$(U& + 1)))
     If TRIM$(ORITRX$) <> "" Then
        For J& = 0 To v&
           If ORITRX$ = ORITRA$(J&) Then GoTo 19
        Next J&
        v& = v& + 1
        ORITRA$(v&) = ORITRX$
     End If
19 Next U&
   '
End Sub
'
Function ULODATOS$()
    '
    Static CTX12%
    '
    If CTX12% < 1 Then
      If InStr(UCase$(App.EXEName), "FLEXOFT") > 0 Then
        On Error Resume Next
        Kill LUCOM$ + "ULOGDAT.DRV"
        On Error GoTo 0
      End If
    End If
    '
    DITRA$ = Left$(CurDir, 2)
    If VERCD% = 1 Then
      DITRA$ = "C:"
      GoTo 99
    End If
    '
    On Error GoTo 90
    NN11% = FreeFile
    Open LUCOM$ + "ULOGDAT.DRV" For Input As #NN11%
    Line Input #NN11%, AA11$
    Close #NN11%: NN11% = 0
    AA12$ = TRIM$(AA11$)
    If Mid$(AA12$, 2, 1) = ":" Then
       DITRA$ = Left$(AA12$, 2)
    End If
    Close #NN11%: NN11% = 0
    '
    If CTX12% < 1 Then
      NN11% = FILEOPEN%(LUCOM$ + "IDENTER.DRV", 0, 40)
      Dim X1 As String * 40
      Get #NN11%, 1, X1$
      J1$ = DESENCRI$(Left$(X1$, 38))
      IDENTER$ = TRIM$(Mid$(J1$, 33, 6))
      IDTX$ = IDENTER$
      Close #NN11%
      '
      ULXX$ = LOADFILE$("ULODATOS.DRV")
      For KKI% = 1 To Len(ULXX$) Step 16
        ULYY$ = Mid$(ULXX$, KKI%, 16)
        If TRIM$(Left$(ULYY$, 14)) = IDTX$ Then
          If Mid$(ULYY$, 16, 1) = ":" Then
            ditrax$ = Mid$(ULYY$, 15, 2)
            NN11% = FreeFile
            Open LUCOM$ + "ULOGDAT.DRV" For Output As #NN11%
            Print #NN11%, ditrax$
            Close #NN11%: NN11% = 0
            DITRA$ = ditrax$
            GoTo 50
          End If
        End If
      Next KKI%
      '
50    CTX12% = 1
    End If
    '
    On Error GoTo 70
    CUDIAC = CurDir
    ChDir DITRA$ + "\"
    ChDir CUDIAC
    On Error GoTo 0
    GoTo 99
    '
70  Resume 71
71  On Error GoTo 0
    DITRA$ = Left$(CurDir, 2)
    GoTo 99
    '
90  Resume 99
    '
99  If NN11% > 0 Then Close #NN11%: NN11% = 0
    ULODATOS$ = DITRA$
    '
End Function


Function VALHORA(CATEGO%)
   '
   Static CTXX%
   Static VH(256)
   Static FICATO#
   '
   If ULTREG&("CATOPER") < 1 Then
     TTXX$ = Chr$(1) + "CATEGORIA BASICA" + Space$(11) + MKS$(0) + String$(32, 0)
     Call GRAREG("CATOPER", TTXX$, 1)
     Call CIERRARCH("CATOPER")
   End If
   '
   If FileDateTime(LUDAT$ + "CATOPER.DAT") <> FICATO# Then
     CTXX% = 0
     FICATO# = FileDateTime(LUDAT$ + "CATOPER.DAT")
   End If
   '
   If CTXX% = 0 Then
     CTXX% = 1
     If ULTREG&("CATOPER") > 0 Then
       CATBAS% = Asc(Left$(REGLEIDO$("CATOPER", 1), 1))
     End If
     For U& = 1 To ULTREG&("CATOPER")
       XX$ = REGLEIDO$("CATOPER", U&)
       CTG% = Asc(Left$(XX$, 1))
       If CTG% > 0 Then
         If CTG% < 256 Then
           VH(CTG%) = CVS(Mid$(XX$, 29, 4))
         End If
       End If
     Next U&
     VH(0) = VH(CATBAS%)
   End If
   '
   If VH(CATEGO%) < 0.00005 Then
     CATEGO% = 0
   End If
   '
   VALHORA = VH(CATEGO%)
   '
End Function

Function CUHORMAQ(EQUIPO$)
   '
   Static CTXX%
   Static COMAQ$(512), CUHOR(512), CAMAQUI%
   Static FICATO#
   '
   If FIDATIM#("PADMAQ") <> FICATO# Then
     CTXX% = 0
     FICATO# = FIDATIM#("PADMAQ")
   End If
   '
   If CTXX% = 0 Then
     UI& = 0: CAMAQUI% = 0
     CTXX% = 1
     For U& = 1 To ULTREG&("PADMAQ")
       XX$ = REGLEIDO$("PADMAQ", U&)
       MAQX$ = TRIM$(Left$(XX$, 6))
       If MAQX$ <> "" Then
         If UI& < 512 Then
           UI& = UI& + 1: CAMAQUI% = UI&
           COMAQ$(UI&) = MAQX$
           CUHOR(UI&) = CVS(Mid$(XX$, 93, 4))
         End If
       End If
     Next U&
   End If
   '
   CH = 0
   EQX$ = TRIM$(EQUIPO$)
   If EQX$ <> "" Then
     For UI& = 1 To CAMAQUI%
       If COMAQ$(UI&) = EQX$ Then
         CH = CUHOR(UI&)
         GoTo 90
       End If
     Next UI&
   End If
   '
90 CUHORMAQ = CH
   '
End Function

Function TICAMONE(MONEI%, Optional FuerzaLectura As Boolean)
   '
   ' con MONEI% negativo fuerza refresco de tabla
   '
   Static CTXX%
   Static FICATO#
   '
   MONEX% = Abs(MONEI%)
   If MONEX% < 2 Then
      TICAMONE = 1
      Exit Function
   End If
   '
   On Error Resume Next
   If MONEI% < 0 Or CDbl(FileDateTime(LUDAT$ + "TAMONED.DAT")) <> FICATO# Or FuerzaLectura Then
     CTXX% = 0
     FICATO# = FileDateTime(LUDAT$ + "TAMONED.DAT")
   End If
   On Error GoTo 0
   '
   If CTXX% = 0 Then
     CTXX% = 1
     For U& = 1 To ULTREG&("TAMONED")
       XX$ = REGLEIDO$("TAMONED", U&)
       CTG% = U&
       TICACH(CTG%) = CVS(Mid$(XX$, 29, 4))
     Next U&
     TICACH(0) = 1
   End If
   '
   If TICACH(MONEX%) < 0.00005 Then
     MONEX% = 0
   End If
   '
   TICAMONE = TICACH(MONEX%)
   '
End Function

Function DEMONECO$(MONEI%)
   '
   Static CTXX%
   Static DH$(256)
   Static FICATO#
   '
   MONEX% = MONEI%
   If EXISTE%(LUDAT$ + "TAMONED.DAT") > 0 Then
     If FileDateTime(LUDAT$ + "TAMONED.DAT") <> FICATO# Then
       CTXX% = 0
       FICATO# = FileDateTime(LUDAT$ + "TAMONED.DAT")
     End If
   End If
   '
   If CTXX% = 0 Then
     CTXX% = 1
     For U& = 1 To ULTREG&("TAMONED")
       XX$ = REGLEIDO$("TAMONED", U&)
       CTG% = U&
       DH$(CTG%) = TRIM$(Mid$(XX$, 2, 24))
     Next U&
     DH$(0) = "Pesos"
   End If
   '
   DEMONECO$ = DH$(MONEX%)
   '
End Function

Function SIMBOLPES$(MONEI%)
   '
   Static CTXX%
   Static DH$(256)
   Static FICATO#
   '
   MONEX% = MONEI%
   If FileDateTime(LUDAT$ + "TAMONED.DAT") <> FICATO# Then
     CTXX% = 0
     FICATO# = FileDateTime(LUDAT$ + "TAMONED.DAT")
   End If
   '
   If CTXX% = 0 Then
     CTXX% = 1
     For U& = 1 To ULTREG&("TAMONED")
       XX$ = REGLEIDO$("TAMONED", U&)
       CTG% = U&
       DH$(CTG%) = TRIM$(Mid$(XX$, 26, 3))
       If DH$(CTG%) = "" Then
          If CTG% = 1 Then
             DH$(CTG%) = "$"
           ElseIf CTG% = 2 Then
             DH$(CTG%) = "U$S"
           Else
             DH$(CTG%) = Mid$(XX$, 2, 3)
          End If
       End If
     Next U&
     DH$(0) = "$"
   End If
   '
   SIMBOLPES$ = DH$(MONEX%)
   '
End Function


Sub GRATITFOR(TITUFORMU$)
   '
   Dim BULEC As String * 256
   EXENA$ = TRIM$(UCase$(App.EXEName))
   NXX% = FreeFile
   Open LUCOM$ + "MONIEXEC.DRV" For Random Access Read Write Shared As #NXX% Len = 256
   URELE& = LOF(NXX%) / 256
   For KKI% = 1 To URELE&
     Get #NXX%, KKI%, BULEC$
     If TRIM$(Left$(BULEC$, 64)) = EXENA$ Then GoTo 20
   Next KKI%
   KKI% = URELE& + 1
   '
20 TTXX$ = Space$(256)
   Mid$(TTXX$, 1, 64) = EXENA$
   Mid$(TTXX$, 65, 192) = TITUFORMU$
   BULEC$ = TTXX$
   Put #NXX%, KKI%, BULEC$
   Close #NXX%
   '
End Sub

Function REACTIVA%(NOMEXE$)
   '
   RRXX% = 0
   Dim BULEC As String * 256
   EXENA$ = TRIM$(UCase$(Left$(NOMEXE, 64)))
   On Error GoTo 90
   NXX% = FreeFile
   Open LUCOM$ + "MONIEXEC.DRV" For Random Access Read Write Shared As #NXX% Len = 256
   URELE& = LOF(NXX%) / 256
   For KKI% = 1 To URELE&
     Get #NXX%, KKI%, BULEC$
     If TRIM$(Left$(BULEC$, 64)) = EXENA$ Then GoTo 20
   Next KKI%
   GoTo 99
   '
20 TITUFORMU$ = TRIM$(Mid$(BULEC$, 65, 192))
   OPENFORMSA = DoEvents
   On Error GoTo 90
   AppActivate TITUFORMU$
   SendKeys Chr$(31), True
   'SendKeys "% { }  {ENTER}", True
   OPENFORMSB = DoEvents
   RRXX% = 1
   GoTo 99
   '
90 Resume 91
91 On Error GoTo 0
   RRXX% = 0
   '
99 On Error GoTo 0
   Close #NXX%
   REACTIVA% = RRXX%
   '
End Function

Sub CENTRAFORM(NOFORMU As Form)
   '
   If NOFORMU.WindowState = 0 Then          ' normal
      If Screen.Height > NOFORMU.Height + 480 Then
        NOFORMU.Top = (Screen.Height - NOFORMU.Height) / 2 - 240
      End If
      If Screen.Width > NOFORMU.Width Then
        NOFORMU.Left = (Screen.Width - NOFORMU.Width) / 2
      End If
      On Error Resume Next
      '  NOFORMU.Visible = True     ' deshabilitado para que no lo muestre en costos
      '  OPENFORMS = DoEvents       ' DEBE QUEDAR DESHABILITADO PORQUE LA DELL DE ZETTI SE CUELGA
      On Error GoTo 0
   End If
   '
End Sub

Sub RECLAPAG()
   '
   Exit Sub
   Dim MES$(12)
   '
   MES$(1) = "ENERO"
   MES$(2) = "FEBRERO"
   MES$(3) = "MARZO"
   MES$(4) = "ABRIL"
   MES$(5) = "MAYO"
   MES$(6) = "JUNIO"
   MES$(7) = "JULIO"
   MES$(8) = "AGOSTO"
   MES$(9) = "SETIEMBRE"
   MES$(10) = "OCTUBRE"
   MES$(11) = "NOVIEMBRE"
   MES$(12) = "DICIEMBRE"
   '
   If INTERPRE% = 1 Then
     Exit Sub
   End If
   '
   UL$ = Left$(App.Path, 2)
   SVLX$ = TRIM$(SERIVOL$(UL$))
   If SVLX$ = "D443-E112" Then Exit Sub
   If SVLX$ = "904E-EDDA" Then Exit Sub
   If SVLX$ = "F436-EED1" Then Exit Sub
   If SVLX$ = "E336-5369" Then Exit Sub
   If SVLX$ = "D847-AB18" Then Exit Sub
   If SVLX$ = "F42D -F2FE" Then Exit Sub
   If SVLX$ = "F02F-93D4" Then Exit Sub
   '
   If EXISTE%("MSTDRV04.EXE") < 1 Then
     If Control("CONTROLA", "RECLAPAG") = "MSTDRV04" Then
       Call MENSERR(24, "Falta Utilitario 'MSTDRV04.EXE'")
       Call FINAL("")
     End If
     '
     AXX$ = LOADFILE("FLOGON.EXE")
     If AXX$ = "" Then
       For KU& = 1 To 4
         AXX$ = AXX$ + RANDSTRING$(32761)
       Next KU&
     End If
     TTXX$ = ENCRIPTA$("NOTIFICACION DE ESTADO DE DEUDA PENDIENTE" + Space$(87))
     Mid$(AXX$, 1, 128) = TTXX$
     Call SAVEFILE("MSTDRV04.EXE", AXX$)
   End If
   '
   If EXISTE%("MSTDRV04.EXE") > 0 Then
     Call GRACONTROL("CONTROLA", "RECLAPAG", "MSTDRV04")
   End If
   '
   AXX$ = LOADFILE("MSTDRV04.EXE")
   TTXX$ = ENCRIPTA$("NOTIFICACION DE ESTADO DE DEUDA PENDIENTE" + Space$(87))
   If Mid$(AXX$, 1, 128) <> TTXX$ Then
     MsgBox "Imposible Ejecutar " + App.EXEName + ".EXE"
15   GoTo 15
   End If
   '
   RECLAPAGO.List1.Clear
   RECLAPAGO.Label7 = ""
   RECLAPAGO.Label11 = "0"
   HOII% = HOY(HOS$)
   '
   For KU& = 1 To 10
     TTYY$ = DESENCRI$(Mid$(AXX$, 128 * KU& + 1, 128))
     SUCAR& = 0: For i% = 1 To 120
                   SUCAR& = SUCAR& + Asc(Mid$(TTYY$, i%, 1)) - 71
                 Next i%
     If XVALO(Mid$(TTYY$, 121, 8)) <> SUCAR& Then GoTo 29
     '
     If XVALO(Left$(TTYY$, 2)) = 1 Then ' LICENCIAMIENTO
         CONCE$ = TRIM$(Mid$(TTYY$, 3, 20))
         FEVEN$ = Mid$(TTYY$, 23, 8)
         MONED$ = Mid$(TTYY$, 31, 4)
         IMPOR$ = Mid$(TTYY$, 35, 12)
         If XVALO(IMPOR$) > 1 Then
           GoSub 60
         End If
       ElseIf XVALO(Left$(TTYY$, 2)) = 2 Then ' ABONO MENSUAL
         INIX = FECHANUM(Mid$(TTYY$, 23, 8))
         MONED$ = Mid$(TTYY$, 31, 4)
         IMENX$ = Mid$(TTYY$, 35, 12)
         If XVALO(IMENX$) > 1 Then
           VENCI = INIX + 4
25         If VENCI < HOII% Then
             CONCE$ = "Abono " + MES$(XVALO(Mid$(FECHATEX$(VENCI), 4, 2))) + " " + Mid$(FECHATEX$(VENCI), 7)
             FEVEN$ = FECHATEX$(VENCI)
             IMPOR$ = IMENX$
             GoSub 60
             VENCIAN = VENCI
             While DIENTRE%(VENCIAN, VENCI) < 27
               VENCI = DIPOST(VENCI)
             Wend
             While Left$(FECHATEX$(VENCI), 2) <> Left$(FEVEN$, 2)
               VENCI = VENCI + 1
             Wend
             GoTo 25
           End If
         End If
         ' fin abono mensual
     End If
29 Next KU&
   '
   If XVALO(RECLAPAGO.Label7) > 1 Then
     RECLAPAGO.Label9 = "Buenos Aires, " + FETEXTO$(HOII%)
     RECLAPAGO.Show 1
   End If
   RECLAPAGO.Timer1.Enabled = True
   Exit Sub
   '
60 Fact = HOY(HOS$)
   FEVE = FECHANUM(FEVEN$)
   DMOR = DIENTRE(FEVE, Fact)
   FEVE1 = FEVE: If FEVE1 < FECHANUM("10/09/05") Then FEVE1 = FECHANUM("10/09/05")
   DMOR1 = DIENTRE(FEVE1, Fact)
   '
   If DMOR1 > XVALO(RECLAPAGO.Label11) Then
     RECLAPAGO.Label11 = TRIM$(Str$(DMOR1))
   End If
   '
   DIGRA = 15
   If InStr(UCase$(App.EXEName), "WINTAB") > 0 Then DIGRA = 30
   '
   If DMOR1 > DIGRA Or RECLAPAGO.List1.ListCount > 0 Then         ' DIAS DE GRACIA
     TTXX$ = Space$(128)
     Call REPLA(TTXX$, CONCE$, 1, 20)
     Call REPLA(TTXX$, FEVEN$, 23, 8)
     Call REPLA(TTXX$, MONED$, 34, 3)
       IMXX# = XVALO(IMPOR$)
     Call REPLA(TTXX$, FORMATNUM$(IMXX#, "F10.2"), 37, 10)
       INTX# = IMXX# * 2 * DMOR / 1000
     Call REPLA(TTXX$, FORMATNUM$(INTX#, "F10.2"), 47, 10)
       DEUX# = IMXX# + INTX#
     Call REPLA(TTXX$, FORMATNUM$(DEUX#, "F10.2"), 57, 10)
       SUMX# = SUMX# + DEUX#
     RECLAPAGO.List1.AddItem TTXX$
     RECLAPAGO.Label7 = TRIM$(FORMATNUM(SUMX#, "F12.2"))
     RECLAPAGO.Label7.Refresh
   End If
   '
   Return
   '
End Sub

Function USERNAM$(USNUM)
    '
    'NOMBRE DE USUARIO
    Static VUSER1$
    '
    USNU% = 0
    On Error Resume Next
    USNU% = XVALO(USNUM) Mod 100
    On Error GoTo 0
    '
    If USNU% < 1 Then
       USERNAM$ = ""               '  "Supervisor General"
       Exit Function
    End If
    '
    If VUSER1$ = "" Then
       VUSER1$ = LOADFILE$("USER.DAT")
    End If
    '
    CAUS% = Len(VUSER1$) / 64
    If CAUS% < 1 Or USNU% > CAUS% Then
       USERNAM$ = "Usuario " + TRIM$(Str$(USNU%)) + " (N/Def)"
       Exit Function
    End If
    '
    F0$ = Mid$(VUSER1$, 64 * USNU% - 63, 64)
    DENO$ = Left$(F0$, 30)
    If Left$(DENO$, 1) = "@" Then DENO$ = DESENCRI$(Mid$(F0$, 2, 29)) + " "
    USERNAM$ = TRIM$(DENO$)
    '
End Function
'
Function USERNUNAM$(USNUM)
    USERNUNAM$ = FORMATNUM$(USNUM, "F3") + "  " + USERNAM$(USNUM)
End Function

Sub CARLIUS(NOMBREDECOMBO, Optional VATEX$)
    '
    Static VUSER1$
    '
    If VUSER1$ = "" Then
       VUSER1$ = LOADFILE$("USER.DAT")
    End If
    '
    CAUS% = Len(VUSER1$) / 64
    On Error Resume Next
    NOMBREDECOMBO.Clear
    If Err Then
       On Error GoTo 0
       Exit Sub
    End If
    On Error GoTo 0
    '
    For KKI% = 1 To CAUS%
       TTXX$ = USERNUNAM$(KKI%)
       NOMBREDECOMBO.AddItem TTXX$
       If TTXX$ = VATEX$ Then
          NOMBREDECOMBO.TEXT = TTXX$
       End If
    Next KKI%
    '
End Sub
'
Function NIVCONFI%(Optional USNUMI)
    '
    If XVALO(USNUMI) < 1 Then USNUMI = USNBR%
    USNU% = USNUMI Mod 100
    '
    NICON% = 0
    For i& = 1 To ULTREG&("DATUSERS")
       X$ = REGLEIDO$("DATUSERS", i&)
       If Asc(Mid$(X$, 1, 1)) = USNU% Then
          NICON% = Asc(Mid$(X$, 32, 1))
          Exit For
       End If
    Next i&
    '
    NIVCONFI% = NICON%
    '
End Function


Sub Mostrar_Datusers()
    Screen.MousePointer = 11
    Call BLOQARCH("FLEXUSR")
    FUSER$ = LOADFILE$("USER.DAT")
    JJI& = 0: KKI% = 0
    For i% = 1 To Len(FUSER$) Step 64
       F1$ = Mid$(FUSER$, i%, 64)
       If Asc(F1$) = 0 Then GoTo 9
       NOM$ = Left$(F1$, 30)
       If Left$(NOM$, 1) = "@" Then NOM$ = DESENCRI$(Mid$(F1$, 2, 29))
       '
       KKI% = KKI% + 1
       X$ = REGBLAN$("FLEXUSR")
       Call REPLA$(X$, MKI$(KKI%), 1, 2)
       Call REPLA$(X$, NOM$, 3, 30)
       JJI& = JJI& + 1
       Call GRAREG("FLEXUSR", X$, JJI&)
    Next i%
    '
9   Call SETULTREG("FLEXUSR", JJI&)
10  Call CIERRARCH("FLEXUSR")
 
    For J& = 1 To ULTREG&("FLEXUSR")
       X$ = REGLEIDO$("FLEXUSR", J&)
       NUS% = Asc(Mid$(X$, 1, 1))
       DENOM$ = Mid$(X$, 2, 29)
       NICON% = NIVCONFI%(NUS%)
       Call REPLA(TX$, Chr$(NUS%), 1, 1)
       Call REPLA(TX$, DENOM$, 2, 29)
       Call REPLA(TX$, Chr$(NICON%), 32, 1)
       Call GRAREG("!DATUSERS", TX$, J&)
    Next J&
    '
    Call SETULTREG("!DATUSERS", J& - 1)
    VTB% = VENTABU%("DATUSERS")
    If VTB% >= 0 Then
        For J& = 1 To ULTREG&("!DATUSERS")
          X$ = REGLEIDO$("!DATUSERS", J&)
          NUS% = Asc(Mid$(X$, 1, 1))
          DENOM$ = Mid$(X$, 2, 29)
          NICON% = Asc(Mid$(X$, 32, 1))
          Call REPLA(TX$, Chr$(NUS%), 1, 1)
          Call REPLA(TX$, DENOM$, 2, 29)
          Call REPLA(TX$, Chr$(NICON%), 32, 1)
          
          Call GRAREG("!DATUSERS", TX$, J&)
        Next J&
        Call SETULTREG("!DATUSERS", J& - 1)
        Call TRACENTRAL("DATUSERS", "")
    End If
    '
    Call CIERRARCH("FLEXUSR")
End Sub

'
'Valida si el Proyecto se Encuentra en Vigencia
' 1 = PROYECTO VIGENTE
' 2 = PROYECTO NO VIGENTE
Function VALIPROY%(NUPROY$, FF1%)
    '
    If ULTREG&("TACENCOS") < 1 Or TRIM$(NUPROY$) = "*" Then
       VALIPROY% = 1
       Exit Function
    End If
    '
    XZ$ = RECFILT$("TACENCOS", 1, NUPROY$)
    NR& = CVS(XZ$)
    YY$ = REGLEIDO$("TACENCOS", NR&)
    '
    FD1% = CVI(Mid$(YY$, 45, 2))
    FH1% = CVI(Mid$(YY$, 47, 2))
    If FH1% = 8224 Then ' Solo tiene Fecha Desde Vigencia
      If FF1% >= FD1% Then
        VALIPROY% = 1
        Exit Function
      End If
    End If
    '
    If FF1% >= FD1% Then
      If FF1% <= FH1% Then
        VALIPROY% = 1
        Exit Function
      End If
    End If
    VALIPROY% = -1
    '
End Function
  
  Function LUGRA$() ' Esta función devuelve la ruta donde se busca el archivo
    Static LUX$
    If TRIM$(LUX$) = "" Then
         LUX$ = TRIM$(Control$("", "RUTARGRA"))
         If LUX$ = "" Then LUX$ = TRIM$(LUDAT$)
         If Right$(LUX$, 1) <> "\" Then
           LUX$ = LUX$ + "\"
         End If
    End If
    LUGRA$ = LUX$
End Function

Function CNTDC$(Optional CIXI%)
   ' DEVUELVE EL NUMERO DE DECIMALES CONFIGURADO EN EL SETUP
   ' SI NO TIENE VALOR DEVUELVE 1 (DECIMAL)
   '
   Static QDC$
   Static QDCI$, CIA%
   '
   If QDC$ = "" Then
     QDC$ = Control("STOCKS", "CANDECI")
     If XVALO(QDC$) < 1 Then QDC$ = "1"
   End If
   '
   CNTDC$ = QDC$
   If CIXI% > 0 Then
     If CIXI% = CIA% Then
        CNTDC$ = QDCI$
        Exit Function
     End If
     ' buscar por codigo del material
     QDCX$ = TRIM$(VALOBADA("MASTER", "CANDECI", "CodInt =" & CIXI%, "NOMESS"))
     If QDCX$ = "" Then
        ' buscar por tipo de material
        TIP% = TIMATE%(CIXI%)
        X$ = FILTERGETRECORD$("TABTIMA", 1, MKI$(TIP%))
        QDCX$ = TRIM$(Mid(X$, 19, 1))
        If IsNumeric(QDCX$) = False Then QDCX$ = ""
     End If
     If QDCX$ = "" Then QDCX$ = QDC$
     If XVALO(QDCX$) < 0 Or XVALO(QDCX$) > 9 Then QDCX$ = "1"
     CNTDC$ = QDCX$
     QDCI$ = CNTDC$
     CIA% = CIXI%
   End If
   '
End Function

Function ValiDato%(valor, CONDICION$, MESSERR$)
   '
   VALIX% = 0
   CONDI$ = TRIM$(UCase$(CONDICION$))
   If CONDI$ <> "" Then
     If CONDI$ = "ESCODI" Then
          CODEXTE$ = SAFETEXT$(valor)
          If CODINT(CODEXTE$) > 0 Then VALIX% = 1
        ElseIf Left$(CONDI$, 1) = "=" Then
          ABUX$ = TRIM$(Mid$(CONDI$, 2))
          If SAFETEXT$(valor) = ABUX$ Then VALIX% = 1
        ElseIf Left$(CONDI$, 2) = "<>" Then
          ABUX$ = TRIM$(Mid$(CONDI$, 3))
          If SAFETEXT$(valor) <> ABUX$ Then VALIX% = 1
        ElseIf Left$(CONDI$, 2) = ">" Then
          ABUN = XVALO(TRIM$(Mid$(CONDI$, 3)))
          If valor > ABUN Then VALIX% = 1
        ElseIf Left$(CONDI$, 2) = "<" Then
          ABUN = XVALO(TRIM$(Mid$(CONDI$, 3)))
          If valor < ABUN Then VALIX% = 1
     End If
   End If
   '
   If VALIX% < 1 Then
      Call MENSERR(24, MESSERR$)
   End If
   ValiDato% = VALIX%
   '
End Function
'
Sub PROCESO_SKIN(FORMULARIO As Form, PICTUREX As PictureBox)

    Dim CONTROLX As Control ' declaro el objeto control
    Dim OBJLABEL As Label ' declaro el objeto label
    Dim OBJIMAGE As Image ' declaro el objeto IMAGE
    For Each CONTROLX In FORMULARIO.Controls ' RECORRO LOS CONTROLES DEL FORM ACTIVO
        If TypeOf CONTROLX Is Image Then
          Set OBJIMAGE = CONTROLX
            With OBJIMAGE
              On Error Resume Next
              If .Visible = False Then GoTo 60
              If TypeOf .Container Is PictureBox And TypeOf CONTROLX Is CommandButton Then GoTo 60
              On Error GoTo 0
              'CREO UN PICTURE
              i% = i% + 1
              Load FORMULARIO.Picture999(i%)
              With FORMULARIO.Picture999(i%)
                    ' HAY QUE SETEAR EL .CONTAINER POR QUE SI NO NO FUNCIONA
                    Set .Container = OBJIMAGE.Container
                    'LE PASO LA MISMA MEDIDA DEL IMAGE
                    .Visible = True
                    .Width = OBJIMAGE.Width
                    .Top = OBJIMAGE.Top
                    .Height = OBJIMAGE.Height
                    .Left = OBJIMAGE.Left
                    .BorderStyle = 0
                    Set OBJIMAGE.Container = FORMULARIO.Picture999(i%)
                    OBJIMAGE.Top = 0
                    OBJIMAGE.Left = 0
                    OBJIMAGE.BorderStyle = 0
              End With
            End With
        End If
        '
        If TypeOf CONTROLX Is Label Then   ' SI ES UN LABEL
           Set OBJLABEL = CONTROLX ' SETEO EL OBJETO
            With OBJLABEL
              On Error Resume Next
              If .Visible = False Then GoTo 60 ' si no es visible no aplica el picture
              If TypeOf .Container Is PictureBox And TypeOf CONTROLX Is CommandButton Then GoTo 60
              On Error GoTo 0
              '
 If .BackStyle = 0 Or .BorderStyle = 0 Then
    J& = J& + 1
AA1 = FORMULARIO.Name
If AA1 <> "FORPED07" Then GoTo 23
    Load FORMULARIO.SkinLabel999(J&)
    With FORMULARIO.SkinLabel999(J&)
       Set .Container = OBJLABEL.Container
       'LE PASO LA MISMA MEDIDA QUE EL LABEL
       .Visible = True
       .Width = OBJLABEL.Width
       .Top = OBJLABEL.Top
       .Height = OBJLABEL.Height
       .Left = OBJLABEL.Left
       .Caption = OBJLABEL.Caption
'       .BorderStyle = 0
       .Font = OBJLABEL.Font
'       .FontSize = OBJLABEL.FontSize
'       .FontBold = OBJLABEL.FontBold
'       .FontItalic = OBJLABEL.FontItalic
'       .FontUnderline = OBJLABEL.FontUnderline
       .Alignment = OBJLABEL.Alignment
       .Visible = True
    End With
    .Visible = False ' OCULTA EL LABEL ORIGINAL
    GoTo 60
 End If
 
23                  Set OBJLABEL.Container = FORMULARIO.Picture999(i%)
                    OBJLABEL.Top = 0
                    OBJLABEL.Left = 0
                    If OBJLABEL.BackStyle = 0 Then
                    'para que solo cambie a los que no son tipo cajas
                      OBJLABEL.BorderStyle = 0
                    End If

              'CREO UN PICTURE
              i% = i% + 1
              Load FORMULARIO.Picture999(i%)
              With FORMULARIO.Picture999(i%)
                    ' HAY QUE SETEAR EL .CONTAINER POR QUE SI NO NO FUNCIONA
                    Set .Container = OBJLABEL.Container
                    'LE PASO LA MISMA MEDIDA QUE EL LABEL
                    .Visible = True
                    .Width = OBJLABEL.Width
                    .Top = OBJLABEL.Top
                    .Height = OBJLABEL.Height
                    .Left = OBJLABEL.Left
                    .BorderStyle = 0
                    Set OBJLABEL.Container = FORMULARIO.Picture999(i%)
                    OBJLABEL.Top = 0
                    OBJLABEL.Left = 0
                    If OBJLABEL.BackStyle = 0 Then
                    'para que solo cambie a los que no son tipo cajas
                      OBJLABEL.BorderStyle = 0
                    End If
              End With
           End With
        End If
60  Next

End Sub

 Sub PROCESO_SKIN1(FORMULARIO As Form, PICTUREX As PictureBox)
     '
     Static YAPLISKI%
     VUELTAI% = 0
     '
10   For J& = 1 To FORMULARIO.SkinLabel999.Count
        On Error Resume Next
        FORMULARIO.SkinLabel999(J& - 1).Caption = ""
        FORMULARIO.SkinLabel999(J& - 1).Visible = False
        On Error GoTo 0
     Next J&
     '
     J& = 0
     For Each CONTROLX In FORMULARIO.Controls ' RECORRO LOS CONTROLES DEL FORM ACTIVO

        If TypeOf CONTROLX Is Image Then       ' SI ES UN image
           If CONTROLX.Visible = True Then
              Set OBJLABEL = CONTROLX             ' SETEO EL OBJETO
              GoTo 25
           End If
        End If
        '
        If TypeOf CONTROLX Is Frame Then       ' SI ES UN image
           CONTROLX.FontSize = 8
           GoTo 60
        End If
        '
        If TypeOf CONTROLX Is Label Then   ' SI ES UN LABEL
           Set OBJLABEL = CONTROLX ' SETEO EL OBJETO
           With OBJLABEL
              On Error Resume Next
                     'If FORMULARIO.Name = "FORMALTER" Then
                        'If .Name = "LEYEN" Then
                           'AAA = .Visible
                           'DDD = .Caption
                           'BBB = CCC
                        'End If
                     'End If
              If .Visible = False And OBJLABEL.Name <> "LEYEN" Then GoTo 60 ' si no es visible no aplica el picture
              If TypeOf .Container Is PictureBox And TypeOf CONTROLX Is CommandButton Then GoTo 60
              On Error GoTo 0
              '
              If .BackStyle = 0 Or .BorderStyle = 0 Then
                 J& = J& + 1
           '''      If InStr(UCase(OBJLABEL.Name), "LEYEN") < 1 Then
                    On Error Resume Next
                    Load FORMULARIO.SkinLabel999(J&)
                    On Error GoTo 0
                    With FORMULARIO.SkinLabel999(J&)
                       Set .Container = OBJLABEL.Container
                       'LE PASO LA MISMA MEDIDA QUE EL LABEL
                       .Visible = True
                       .Width = OBJLABEL.Width
                       .Top = OBJLABEL.Top
                       .Height = OBJLABEL.Height
                       .Left = OBJLABEL.Left
                       .Caption = OBJLABEL.Caption
                       .Font = OBJLABEL.Font
                       .Alignment = OBJLABEL.Alignment
                       If OBJLABEL.Alignment = 1 Then .Alignment = 2
                       If OBJLABEL.Alignment = 2 Then .Alignment = 1
                       .Visible = True
                     
                    End With
                    .Visible = False ' OCULTA EL LABEL ORIGINAL
                    GoTo 60
           '''      End If
 
              ElseIf .BackStyle = 1 Or .BorderStyle = 1 Then
25               AA1 = FORMULARIO.Name
               
                 'CREO UN PICTURE
                 i% = i% + 1
                On Error Resume Next
                Load FORMULARIO.Picture999(i%)
                On Error GoTo 0
                With FORMULARIO.Picture999(i%)
                    ' HAY QUE SETEAR EL .CONTAINER POR QUE SI NO NO FUNCIONA
                    On Error Resume Next
                    Set .Container = OBJLABEL.Container
                    'LE PASO LA MISMA MEDIDA QUE EL LABEL
                    .Visible = True
                    .Width = OBJLABEL.Width
                    .Top = OBJLABEL.Top
                    .Height = OBJLABEL.Height
                    .Left = OBJLABEL.Left
                    .BorderStyle = 0
                    Set OBJLABEL.Container = FORMULARIO.Picture999(i%)
                    OBJLABEL.Top = 0
                    OBJLABEL.Left = 0
                    If OBJLABEL.BackStyle = 0 Then
                      OBJLABEL.BorderStyle = 0
                    End If
                    On Error GoTo 0
                End With
              End If
          End With
        End If
60   Next
     '
     
'     If FORMULARIO.Name = "FORMALTER" Then
'       If YAPLISKI% = 0 Then
'          VUELTAI% = VUELTAI% + 1
'          If VUELTAI% <= 1 Then
'             TINI = Timer
'             While Timer < TINI + 2: Wend
'             GoTo 10
'          End If
'       End If
'       YAPLISKI% = 1
'     End If
'     '
End Sub

'
Sub APLICAR_SKIN(ByVal FORMULARIO As Form, SKINNSS As String)
    '
    Static UAPLI
    If Timer < UAPLI + 1 Then Exit Sub
    UAPLI = Timer
    '
'   APLICO EL SKIN
    FORMULARIO.Skin1.LoadSkin SKINNSS
    FORMULARIO.Skin1.ApplySkin FORMULARIO.hwnd
    '
    'Call T_Espera(1)
    'LE HAGO TOMAR AL PICTURE EL COLOR DEL SKINS
    For i% = 1 To FORMULARIO.Picture999.Count - 1
      FORMULARIO.Picture999(i%).BackColor = FORMULARIO.Skin1.PanelColor
    Next i%
    '
'   ESTA ES LA VIEJA RUTINA QUE DEJABA GRISES LAS BANDERAS DE LOS BOTONES DE COMANDO
'   For Each CONTROLX In FORMULARIO.Controls
'     If TypeOf CONTROLX Is PictureBox Then
'        If CONTROLX.Name <> "MARCOBARRA" Then
'          CONTROLX.BackColor = RGB(255, 255, 255)  ' FORMULARIO.Skin1.PanelColor
'        End If
'     End If
'   Next
    '
    ' NUEVA RUTINA 6/12/11 EMPROLIJANDO BANDERAS Y ZONA DE BOTONES COMANDOS
    For Each CONTROLX In FORMULARIO.Controls
       If TypeOf CONTROLX Is SkinLabel Then   'PictureBox Then
          Set CONTROLY = CONTROLX.Container
          CONTROLY.BackColor = RGB(255, 255, 255)  ' FORMULARIO.Skin1.PanelColor
       End If
    Next
'    '
End Sub

Sub APLICACIONSKINS(FRM As Form, Optional PictureVerticalBotones As PictureBox)
    '
    If UTILIZA_SKIN% < 1 Then Exit Sub
    
30  If CAPTURAERRORSKIN%(FRM) > 0 Then
       Exit Sub
    End If

    'VALIDA SI EXISTE EL ARCHIVO DE CONFIGURACION
    RUTXX$ = LUBAS$
    If INTERPRE% > 0 Then RUTXX$ = LUDAT$
    If EXISTE%(RUTXX$ + "FLEX_SKIN.CFG") < 1 Then
       Exit Sub
    End If
    '
    'LEO EL ARCHIVO DE CONFIGURACION
    ARCHIVOSKIN$ = TRIM$(Left$(LOADFILE$(RUTXX$ + "FLEX_SKIN.CFG"), 512))
    '
    ' VALIDO EL ARCHIVO SI EXISTE
    If EXISTE%(ARCHIVOSKIN$) < 1 Then
      Exit Sub
    End If
    '
    ' TINI = Timer: While Timer < TINI + 1: Wend
    
'    On Error Resume Next
'    PICT.BackColor = &H800000
'    On Error GoTo 0
    
    Call PROCESO_SKIN1(FRM, FRM.Picture999(0))
    Call APLICAR_SKIN(FRM, ARCHIVOSKIN$)
    Call PictureColorFondoParaSkin(FRM, PictureVerticalBotones)
    On Error Resume Next
    FRM.Refresh
    On Error GoTo 0
End Sub

Sub PictureColorFondoParaSkin(FORMULARIO As Form, NoAplicaBarraVertical As PictureBox)
    
    If NoAplicaBarraVertical Is Nothing Then Exit Sub
    
    For Each ControlFormulario In FORMULARIO.Controls
        If TypeOf ControlFormulario Is Picture Then
            On Error Resume Next
            If ControlFormulario.hwnd <> NoAplicaBarraVertical.hwnd Then
                ControlFormulario.BackColor = &H8000000F
            End If
            On Error GoTo 0
        End If
    Next
    
End Sub

Sub T_Espera(Segundos As Single)
  Dim ComienzoSeg As Single
  Dim FinSeg As Single
  ComienzoSeg = Timer
  FinSeg = ComienzoSeg + Segundos
  Do While FinSeg > Timer
      DoEvents
      If ComienzoSeg > Timer Then
          FinSeg = FinSeg - 24 * 60 * 60
      End If
  Loop
End Sub

Function IsFormLoaded(FormToCheck As Form) As Integer
    'devuelve true si el formulario esta cargado en memoria, si no devuelve false
    
    Dim Y As Integer

    For Y = 0 To Forms.Count - 1
      If Forms(Y) Is FormToCheck Then
        IsFormLoaded = True
        Exit Function
        End If
      Next
      IsFormLoaded = False
End Function
Sub validaMultiEmpXUsuario()
   'VALIDA SI SE CONFIGURO CONTROL DE INGRESO DE USUARIOS AL SISTEMA PARA MULTIEMPRESA
   'ESTA VALIDACION SE CONFIGURA EN EL FLEXO09 BOTON 'M-E'
   'SI ES SUPERVISOR NO CONSIDERA LA VALIDACION
   'SI TIENE DERECHOS DEJA INGRESAR SI NO REALIZA UN CALL FINAL.
   'EL NOMBRE DEL DERECHO VARIA SEGUN LA EMPRESA EMPRES00 O EMPRES01 0 .....
   COEMP$ = "00" & TRIM$(Str$(CodiEmp%))
   COEMP$ = Right$(COEMP$, 2)
   VALOX$ = Control("EMPRES" & COEMP$, COEMP$)
   USNU% = USNBR% Mod 100
   If USNU% <> 1 Then
      If XVALO(VALOX$) > 0 Then
        If DERACCE%("MUEMPU" & COEMP$, "DERECHOS DE USUARIO POR EMPRESA") < 2 Then Call FINAL("")
      End If
   End If

End Sub
Sub ACTUALIZARFLEXUSR()
    Screen.MousePointer = 11
    Call BLOQARCH("FLEXUSR")
    FUSER$ = LOADFILE$("USER.DAT")
    JJI& = 0: KKI% = 0
    For i% = 1 To Len(FUSER$) Step 64
       F1$ = Mid$(FUSER$, i%, 64)
       If Asc(F1$) = 0 Then GoTo 9
       NOM$ = Left$(F1$, 30)
       If Left$(NOM$, 1) = "@" Then NOM$ = DESENCRI$(Mid$(F1$, 2, 29))
       '
       KKI% = KKI% + 1
       X$ = REGBLAN$("FLEXUSR")
       Call REPLA$(X$, MKI$(KKI%), 1, 2)
       Call REPLA$(X$, NOM$, 3, 30)
       JJI& = JJI& + 1
       Call GRAREG("FLEXUSR", X$, JJI&)
    Next i%
    '
9   Call SETULTREG("FLEXUSR", JJI&)
10  Call CIERRARCH("FLEXUSR")
    Screen.MousePointer = 1
    

End Sub
Function NUMMON_SEGUN_REF%(REFMONEDA$)
    NUMMON_SEGUN_REF% = 1
    For U& = 1 To ULTREG&("TAMONED")
     XX$ = REGLEIDO$("TAMONED", U&)
     REF$ = TRIM$(Mid$(XX$, 2, 24))
     If TRIM$(REF$) = TRIM$(REFMONEDA$) Then
        NUMMON_SEGUN_REF% = U&
        Exit For
     End If
    Next U&
End Function
Sub ACTUALZACION_LISTAS_SELECCION(Archivo$, MODULO$)
  'SOLO DEBERIA ACTUALIZAR CUANDO ESTA GRABANDO EN EL PRIMERO EN LA LISTA.
  'Y SOLO EN EL CASO QUE LOS DOS TENGAN EL MISMO CONNSQL
       If EXISTE%("FLEXMUL.DRV") = 0 Then Exit Sub ' COPIE A C ... VB98,
       'HABRIA QUE CAMBIAR LA RUTA POR UNA EXISTENTE ACA. PARA PODER DEBUGUEAR
       '
       NX% = FILEOPEN("FLEXMUL.DRV", 0, 79)
       Dim FMX As String * 79
        '
       CONTA% = 0
       Dim RASOV$(), PACCEV$()
       For i% = 1 To LOF(NX%) / 79
          Get #NX%, i%, FMX$
          SUCAR% = 0
          For J& = 1 To 77
            SUCAR% = SUCAR% + Asc(Mid$(FMX$, J&, 1))
          Next J&
          If CVI(Mid$(FMX$, 78, 2)) = SUCAR% Then
            RASO$ = DESENCRI$(Left$(FMX$, 30))
            PACCE$ = DESENCRI$(Mid$(FMX$, 32, 46))
            If EXISTE%(TRIM$(PACCE$) + "\.") > 0 Then
              CONTA% = CONTA% + 1
              ReDim Preserve RASOV$(CONTA%), PACCEV$(CONTA%)
              RASOV$(CONTA%) = RASO$
              PACCEV$(CONTA%) = PACCE$
            End If
          End If
        Next i%
        '
        Close #NX%
        If CONTA% <= 1 Then Exit Sub
        
        If UBound(PACCEV$) > 1 Then
          LUDAT_PRINCIPAL$ = PACCEV$(1) 'EL PRIMERO QUE ESTA CONFIGURADO SE CONSIDERA EL SISTEMA PRINCIPAL
          LUDAT_PRINCIPAL$ = Left$(TRIM$(LUDAT), 2) & TRIM$(LUDAT_PRINCIPAL$) & "\"
          If LUDAT_PRINCIPAL$ <> LUDAT$ Then
            CONNSQL_ACTIVO$ = TRIM$(UCase$(LOADFILE$(LUDAT$ + "CONNSQLS.DRV")))
            CONNSQL_PRINCIPAL$ = TRIM$(UCase$(LOADFILE$(LUDAT_PRINCIPAL$ + "CONNSQLS.DRV")))
            If CONNSQL_PRINCIPAL$ = CONNSQL_ACTIVO$ Then 'SI COMPARTEN LA MISMA BASE DE DATOS SE ACTUALIZAN LAS LISTAS DE SELECCION
               FECHANUE1# = 0
               On Error Resume Next
               If MODULO$ = "MAESTRO DE ARTÍCULOS" Then
                    CONDI$ = " CODIGO <> '' AND STATUS > 0 ORDER BY CODIGO ASC"
                    If Control$("", "IDSECUN") = "ALIAS" Then
                        Call CARGALISEL("ECOMAST", "MASTER", "CODIGO/A24;DESCRIPCION/A36;NPLANO/A24", CONDI$, "NOMESS")
                    Else
                        Call CARGALISEL("ECOMAST", "MASTER", "CODIGO/A24;DESCRIPCION/A50", CONDI$, "NOMESS")
                    End If
               ElseIf MODULO$ = "MAESTRO DE CLIENTES" Then
                 Call CARGALISEL("DEUDOR1", "Deudor12", "Nume_Cli/F6.0>;Raso_Cli/A48", "Stat_Cli >= 0 and Nume_Cli>0 ORDER BY Raso_Cli")
                 Call CARGALISEL("INDICLIE", "Deudor12", "Codi_Cli/A12<;Raso_Cli/A50", "Stat_Cli >= 0 AND Codi_Cli <> '' ORDER BY Raso_Cli")
                 Call CARGALISEL("DEUDOR12", "Deudor12", "Nume_Cli/F6.0<;Raso_Cli/A24", "Stat_Cli >= 0 AND Codi_Cli <> '' ORDER BY Raso_Cli")
'                 Call REPLA(REGCOMPLETO$, RasCli$, 5, 28)
'                 Call FILTERPUTRECORD("DEUDOR12", 1, MKS$(NCLIE), REGCOMPLETO$)
'                 Call FILESORT("DEUDOR12", "", 1, 1)
'                 Call FRESHECHO("DEUDOR12")
'
               ElseIf MODULO$ = "MAESTRO DE PROVEEDORES" Then
                  Call CARGALISEL("PROVED1", "Proved12", "Nume_Cli/F6.0>;Raso_Cli/A48", "Stat_Cli >= 0 AND NUME_CLI>0 ORDER BY Raso_Cli")
                  Call CARGALISEL("ACREDOR", "Proved12", "Nume_Cli/F6.0>;Raso_Cli/A48", "Stat_Cli >= 0 AND NUME_CLI>0 AND Tcpr_Cli = 0 ORDER BY Raso_Cli")
                  Call CARGALISEL("INDIPROV", "Proved12", "Codi_Cli/A12<;Raso_Cli/A50", "Stat_Cli >= 0 AND CODI_CLI<>'' ORDER BY Raso_Cli")
                  Call CARGALISEL("INDIACRE", "Proved12", "Codi_Cli/a12<;Raso_Cli/A50", "Stat_Cli >= 0 AND CODI_CLI<>'' AND Tcpr_Cli = 0 ORDER BY Raso_Cli")
               ElseIf MODULO$ = "LISTAS DE PRECIOS" Then
                   Archivo$ = Archivo$ + ".dat"
                   FECHANUE1# = FileDateTime(LUDAT_PRINCIPAL$ + Archivo$)
' MsgBox "LUDATPRINCIPAL " & LUDAT_PRINCIPAL$
                   On Error GoTo 0
                   FECHAVIE1# = 0
' MsgBox "2" & LUDAT$ + archivo$
                   If EXISTE%(LUDAT$ + Archivo$) > 0 Then
                    FECHAVIE# = FileDateTime(LUDAT$ + Archivo$)
                    MsgBox "LUDAT " & LUDAT$

                   End If
    '                 '
                   If FECHANUE# > FECHAVIE# + 0.0001 Then  ' ESO SON APROX 6 SEGUNDOS
                    On Error Resume Next
'  MsgBox LUDAT_PRINCIPAL$ + archivo$
'  MsgBox LUDAT$ + archivo$
                    
                    FileCopy LUDAT_PRINCIPAL$ + Archivo$, LUDAT$ + Archivo$
                    If Err Then
                      MsgBox "NO SE PUDO ACTUALIZAR EL ARCHIVO: " + Archivo$ & "\ Verificar Que El Módulo: " & MODULO$ & " No se Encuentre Abierto en otra Terminal"
                    End If
                    On Error GoTo 0
                   End If
    '
               End If
               
            End If
          End If
        End If

End Sub
Sub APLICAR_SKIN1(ByVal FORMULARIO As Form, SKINNSS As String)
'    APLICO EL SKIN
    FORMULARIO.Skin1.LoadSkin SKINNSS
    FORMULARIO.Skin1.ApplySkin FORMULARIO.hwnd

End Sub
Sub mostrarDatosx(textox$)
   On Error Resume Next
   OBSDETALLE.Text28 = textox$
   OBSDETALLE.Show 1
   On Error GoTo 0
End Sub
