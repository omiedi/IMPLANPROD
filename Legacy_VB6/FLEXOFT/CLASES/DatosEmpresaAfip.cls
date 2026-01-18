VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DatosEmpresaAfip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private ConfiguracionOk As String

Private o_cuit As String
Private o_numeroDocumento As String
Private o_condicionIva As Integer
Private o_razonSocial As String
Private o_domicilioFiscal As String
Private o_codigoPostalDomicilioFiscal As String
Private o_localidadDomicilioFiscal As String
Private o_provinciaDomicilioFiscal As String
Private o_estado As Integer

Private o_provincias(24) As String

Private cuitRepresentada As String

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
    
Property Get Cuit() As String
    Cuit = o_cuit
End Property

Property Get numeroDocumento() As String
    numeroDocumento = o_numeroDocumento
End Property

Property Get condicionIva() As Integer
    condicionIva = o_condicionIva
End Property

Property Get razonSocial() As String
    razonSocial = o_razonSocial
End Property

Property Get domicilioFiscal() As String
    domicilioFiscal = o_domicilioFiscal
End Property

Property Get codigoPostalDomicilioFiscal() As String
    codigoPostalDomicilioFiscal = o_codigoPostalDomicilioFiscal
End Property

Property Get localidadDomicilioFiscal() As String
    localidadDomicilioFiscal = o_localidadDomicilioFiscal
End Property

Property Get provinciaDomicilioFiscal() As String
    provinciaDomicilioFiscal = o_provinciaDomicilioFiscal
End Property

Property Get ESTADO() As Integer
    ESTADO = o_estado
End Property

Private Sub Class_Initialize()

    o_provincias(0) = "CIUDAD AUTONOMA DE BUENOS AIRES"
    o_provincias(1) = "BUENOS AIRES"
    o_provincias(2) = "CATAMARCA"
    o_provincias(3) = "CORDOBA"
    o_provincias(4) = "CORRIENTES"
    o_provincias(5) = "ENTRE RIOS"
    o_provincias(6) = "JUJUY"
    o_provincias(7) = "MENDOZA"
    o_provincias(8) = "LA RIOJA"
    o_provincias(9) = "SALTA"
    o_provincias(10) = "SAN JUAN"
    o_provincias(11) = "SAN LUIS"
    o_provincias(12) = "SANTA FE"
    o_provincias(13) = "SANTIAGO DEL ESTERO"
    o_provincias(14) = "TUCUMAN"
    o_provincias(16) = "CHACO"
    o_provincias(17) = "CHUBUT"
    o_provincias(18) = "FORMOSA"
    o_provincias(19) = "MISIONES"
    o_provincias(20) = "NEUQUEN"
    o_provincias(21) = "LA PAMPA"
    o_provincias(22) = "RIO NEGRO"
    o_provincias(23) = "SANTA CRUZ"
    o_provincias(24) = "TIERRA DEL FUEGO"
    
    ConfiguracionOk = 1
    CarpetaSistema$ = App.Path + "\"
    If EXISTE%(CarpetaSistema$ & "ToolsAfip.exe") = 1 Then
        Call FileCopy(CarpetaSistema$ & "ToolsAfip.exe", LUCOM$ & "ToolsAfip.exe")
    Else
        ConfiguracionOk = 0
    End If
    
    If EXISTE%(CarpetaSistema$ & "ToolsAfip.exe.Config") = 1 Then
        Call FileCopy(CarpetaSistema$ & "ToolsAfip.exe.Config", LUCOM$ & "ToolsAfip.exe.Config")
    Else
        ConfiguracionOk = 0
    End If
    
    
    If EXISTE%(CarpetaSistema$ & "ClienteLoginCms_CS.dll") = 1 Then
        Call FileCopy(CarpetaSistema$ & "ClienteLoginCms_CS.dll", LUCOM$ & "ClienteLoginCms_CS.dll")
    Else
        ConfiguracionOk = 0
    End If
        
    If EXISTE%(CarpetaSistema$ & "DatosContribuyente.dll") = 1 Then
        Call FileCopy(CarpetaSistema$ & "DatosContribuyente.dll", LUCOM$ & "DatosContribuyente.dll")
    Else
        ConfiguracionOk = 0
    End If
        
    If EXISTE%(LUDAT$ & "micertificado.pfx") = 1 Then
        Call FileCopy(LUDAT$ & "micertificado.pfx", LUCOM$ & "micertificado.pfx")
    Else
        ConfiguracionOk = 0
    End If
    
    N1% = FILEOPEN(LUDAT$ + "SVD147.EXE", 0, 41)
    Dim X As String * 41
    Get #N1%, 6, X
    cuitRepresentada = TRIM$(REPLACAR$(TRIM$(DESENCRI$(X)), "-", ""))
    Close #N1%
        
End Sub

Public Function ObtenerDatosEmpresaAfip(cuitDni As String, Optional response As String) As Boolean

    Dim SegundosTranscurridos As Long
    Dim InicioLecturaAfip As Date
    Dim intentosCredencial As Integer
    
    Screen.MousePointer = 11
    
    resultadoCredencial% = ObtenerCredencial()
    If resultadoCredencial% < 1 Then
        If resultadoCredencial% = 0 Then response = "Demora al Obtener Datos desde Afip"
        If resultadoCredencial% < 0 Then response = "Fallo al Eliminar Constancia"
        GoTo 65
    End If
        
    cuitBuscada$ = TRIM$(REPLACAR$(cuitDni, "-", ""))
    Token$ = TRIM$(CONTROL$("ALTAAFIP", "CRE_TKN"))
    Sign$ = TRIM$(CONTROL$("ALTAAFIP", "CRE_SGN"))
    InicioLecturaAfip = Now
    On Error Resume Next
    Kill LUCOM$ & "cuit_" & cuitBuscada$ & ".txt"
    On Error GoTo 0
    Respuesta = Shell(LUCOM$ & "TOOLSAFIP.EXE -dc " & Token$ & " " & Sign$ & " " & cuitRepresentada & " " & cuitBuscada$, vbMinimizedNoFocus)
    
10  SegundosTranscurridos = DateDiff("s", InicioLecturaAfip, Now)
    If SegundosTranscurridos > 30 Then
        response = "Tiempo de Espera Agotado. ReintenEte la Consulta."
        GoTo 65
    End If
    If EXISTE%(LUCOM$ & "CUIT_" & cuitBuscada$ & ".TXT") = 0 Then GoTo 10
        
    NroArchivo% = FILEOPEN%(LUCOM$ & "CUIT_" & cuitBuscada$ & ".TXT", 1, 128) 'ABRE EL ARCHIVO DE LECTURA
    Do While Not EOF(NroArchivo%) ' RECORRE
        Line Input #NroArchivo%, Linea$ ' LEE CADA REGISTRO

        If Mid$(Linea$, 1, 8) = "empresa=" Then
            o_razonSocial = TRIM$(Left$(Mid$(Linea$, 9), 128))
            If Mid$(o_razonSocial, 1, 15) = "SIN INFORMACION" Then GoTo 60     ' BANDERA DE QUE NO PUDO OBTENER INFORMACION DE AFIP
        ElseIf Mid$(Linea$, 1, 10) = "domicilio=" Then
            o_domicilioFiscal = TRIM$(Left$(Mid$(Linea$, 11), 64))
        ElseIf Mid$(Linea$, 1, 13) = "codigopostal=" Then
            o_codigoPostalDomicilioFiscal = Left$(Mid$(Linea$, 14), 16)
        ElseIf Mid$(Linea$, 1, 10) = "localidad=" Then
            o_localidadDomicilioFiscal = Left$(Mid$(Linea$, 11), 48)
        ElseIf Mid$(Linea$, 1, 10) = "provincia=" Then
            nombreProvincia$ = TRIM$(Mid$(Linea$, 11))
            o_provinciaDomicilioFiscal = codProvinciaFlx$(nombreProvincia$)
        ElseIf Mid$(Linea$, 1, 12) = "posicioniva=" Then
            o_condicionIva = XVALO(Mid$(Linea$, 13))
        ElseIf Mid$(Linea$, 1, 7) = "estado=" Then
            o_estado = XVALO(Mid$(Linea$, 8))
        ElseIf Mid$(Linea$, 1, 14) = "estadomensaje=" Then
            response = TRIM(Mid$(Linea$, 15)) & vbCrLf
        ElseIf Mid$(TRIM$(Linea$), 1, 6) = "<error" Or Mid$(TRIM$(Linea$), 1, 7) = "</error" Then
            response = response + TRIM(Linea$) & vbCrLf
        End If
    Loop
    o_cuit = cuitBuscada$
    
25  Close #NroArchivo%
    ObtenerDatosEmpresaAfip = True
    If o_estado > 0 Then GoTo 65
    Screen.MousePointer = 1
    GoTo 99

60  Close #NroArchivo%
65  ObtenerDatosEmpresaAfip = False
    Screen.MousePointer = 1
    
99  Exit Function
        
End Function

Private Function codProvinciaFlx$(nombreProvincia As String)

    codProvinciaFlx$ = ""
    nombreProvinciaAuxiliar$ = UCase(nombreProvincia)
    
5   For ireg& = 1 To ULTREG&("PROVINC")
        resgistroProvincia$ = REGLEIDO$("PROVINC", ireg&)
        nombreP$ = TRIM$(Mid$(resgistroProvincia$, 2, 25))
        
        If InStr(1, UCase(nombreP$), nombreProvinciaAuxiliar$) > 0 Then
            codProvinciaFlx$ = TRIM$(Mid$(resgistroProvincia$, 1, 1))
            Exit For
        End If
    Next ireg&
    
    If InStr(1, nombreProvinciaAuxiliar$, "CIUDAD AUTONOMA") > 0 And codProvinciaFlx$ = "" Then
        nombreProvinciaAuxiliar$ = "CAPITAL FEDERAL"
        GoTo 5
    End If

End Function

Sub DescargarConstanciaInscripcion(cuitDni As String)

    Dim sUrl As String
    Const SW_NORMAL = 5
    
    cuitDni = TRIM$(REPLACAR$(cuitDni, "-", ""))
    
    sUrl = "https://soa.afip.gob.ar/sr-padron/v1/constancia/" & cuitDni
    ShellExecute 0, "Open", sUrl, "", "", SW_NORMAL

End Sub

Function ObtenerIdCondicionIvaAfip(cuitDni As String) As Integer
        
    ObtenerIdCondicionIvaAfip = -1
    If ObtenerDatosEmpresaAfip(cuitDni) Then
        ObtenerIdCondicionIvaAfip = o_condicionIva
    End If

End Function

Function ObtenerCredencial() As Integer
    
    Const SW_MINIMIZE = 6
    Dim intentosSolicitud As Integer
5   ObtenerCredencial = 1
    FechaExpira$ = TRIM$(CONTROL$("ALTAAFIP", "CRE_EXP"))
    If FechaExpira$ <> "" Then
        ANIO% = CInt(Mid$(FechaExpira$, 7, 2))
        If InStr(1, UCase(FechaExpira$), "A.M.") > 0 Or InStr(1, UCase(FechaExpira$), "A. M.") > 0 Then
            If ANIO% = 20 Then ' FORMATO DE AÑO aaaa
                FechaExpira$ = Mid$(FechaExpira$, 1, 16)
            Else               ' FORMATO DE AÑO aa
                FechaExpira$ = Mid$(FechaExpira$, 1, 14)
            End If
        ElseIf InStr(1, UCase(FechaExpira$), "P.M.") > 0 Or InStr(1, UCase(FechaExpira$), "P. M.") > 0 Then
            If ANIO% = 20 Then ' FORMATO DE AÑO aaaa
                HORA% = XVALO(Mid$(FechaExpira$, 12, 2))
                If HORA% < 12 Then HORA% = HORA% + 12
                FechaExpira$ = Mid$(FechaExpira$, 1, 11) & " " & CStr(HORA%) & Mid$(FechaExpira$, 14, 3)
            Else               ' FORMATO DE AÑO aa
                HORA% = XVALO(Mid$(FechaExpira$, 10, 2))
                If HORA% < 12 Then HORA% = HORA% + 12
                FechaExpira$ = Mid$(FechaExpira$, 1, 9) & " " & CStr(HORA%) & Mid$(FechaExpira$, 12, 3)
            End If
        End If
        FechaExpiraUltimaCredencial = CDate(FechaExpira$)
    Else
        FechaExpiraUltimaCredencial = CDate("01/01/2000")
    End If
    If FechaExpiraUltimaCredencial < Now Then
        If EXISTE%(LUCOM$ & "wsaa_constancia.txt") > 0 Then
            NroArchivo% = FILEOPEN%(LUCOM$ & "wsaa_constancia.txt", 1, 128) 'ABRE EL ARCHIVO DE LECTURA
            Do While Not EOF(NroArchivo%) ' RECORRE
                Line Input #NroArchivo%, Linea$ ' LEE CADA REGISTRO
                If Mid$(Linea$, 1, 6) = " Token" Then
                    Call GRACONTROL("ALTAAFIP", "CRE_TKN", Mid$(Linea$, 9))
                ElseIf Mid$(Linea$, 1, 5) = " Sign" Then
                    Call GRACONTROL("ALTAAFIP", "CRE_SGN", Mid$(Linea$, 8))
                ElseIf Mid$(Linea$, 1, 15) = " ExpirationTime" Then
                    Call GRACONTROL("ALTAAFIP", "CRE_EXP", Mid$(Linea$, 17))
                End If
            Loop
            Close #NroArchivo%
            On Error GoTo 0
            
            On Error GoTo EliminaConstancia
            Kill LUCOM$ & "wsaa_constancia.txt"
            On Error GoTo 0
        Else
            If intentosSolicitud = 0 Then Shell LUCOM$ & "toolsafip.exe -tk -s ws_sr_constancia_inscripcion -c " & LUCOM$ & "micertificado.pfx -w https://wsaa.afip.gov.ar/ws/services/LoginCms?wsdl -v on", vbMinimizedNoFocus
            intentosSolicitud = intentosSolicitud + 1
            If intentosSolicitud > 5 Then
                ObtenerCredencial = 0: Exit Function
            End If
            Call T_Espera(2)
            GoTo 5
        End If
    End If
    On Error GoTo 0
    
99: Exit Function

EliminaConstancia:
    ObtenerCredencial = -1

End Function

