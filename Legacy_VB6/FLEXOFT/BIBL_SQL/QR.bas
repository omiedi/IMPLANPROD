Attribute VB_Name = "QR"
Function IMAGENCODBARRA$(Ruta$, NOMB_ARCHIVO$, TEXTOQR$, Optional TIPOBARRA$, Optional ANCHOX_mm%, Optional ALTOX_mm%)
    'EJEMPLO
    'RUTA$ = "C:\Flexoft\"
    'NOMB_ARCHIVO$ = "ImagenCodBarraOF.bmp"
    'TEXTOQR$ = "TRAP-0000-0001 OF-00105-01 150"
    'DEVUELVE LA RUTA DEL ARCHIVO DONDE SE GENERO EL MISMO.
    'PARA QUE FUNCIONE HAY QUE INSTALAR EL EJECUTABLE EN WINDOWS Y LOS ARCHIVO 'MessagingToolkit.QRCode.DLL Y zxing.dll ' TAMBIEN EN WINDOWS
    'RUTAEJECUT_VBNET$ = "C:\Program Files (x86)\Microsoft Visual Studio\VB98\IMPLANET\QR_Generador\QR_Generador\bin\Debug\"
    
    '***** RECORDAR **** EN Diseñador de frm EL OBJETO GRAPH-n en el casillero Texto (Cinta Amarilla) debe ir el texto @CAPTURA({COMPAR-0})
    
    'ATENCION PASAR EL ALTO Y ANCHO EN MM EL MISMO SE MULTIPLICA POR 3.3 PARA LLEVARLO A LA UNIDAD DE MEDIDA DEL VB.NET
    '
    If ANCHOX_mm% > 0 Then ANCHOX_mm% = ANCHOX_mm% * 3.3 Else ANCHOX_mm% = 0
    If ALTOX_mm% > 0 Then ALTOX_mm% = ALTOX_mm% * 3.3 Else ALTOX_mm% = 0
    If TRIM$(TIPOBARRA$) = "" Then TIPOBARRA$ = "QR_CODE" ' POR DEFECTO QR
    '
    RUTAEJECUT_VBNET$ = App.Path + "\"
    'XXX = Shell(RUTAEJECUT_VBNET$ + "QR_Generador.exe " & Ruta$ & ";" & NOMB_ARCHIVO$ & ";" & TEXTOQR$, 4) 'VER QUE SE SEPARA POR ;
    XXX = Shell(RUTAEJECUT_VBNET$ + "GENERAR BARRAS.exe " & Ruta$ & ";" & NOMB_ARCHIVO$ & ";" & TEXTOQR$ & ";" & TIPOBARRA$ & ";" & ANCHOX_mm% & ";" & ALTOX_mm%, 4)  'VER QUE SE SEPARA POR ;
    IMAGENCODBARRA$ = Ruta$ & NOMB_ARCHIVO$
    Call T_Espera(1)
End Function
Sub GRABAR_QR(RUTARCHIVOX$, DOCU_ORIGEN$)
    User% = USNBR% Mod 100
    If User% < 1 Then User% = 1

    CONDI$ = "DOCUMENTOORIGEN='" & DOCU_ORIGEN$ & "'  AND PATH_IMAG = '" & RUTARCHIVOX$ & "' AND NUMEROUSUARIO = " & User%
    If CantRegistros&("ARCHIVOS_QR", CONDI$, "NOMESS") > 0 Then
        Call ACTUREGISTRO("ARCHIVOS_QR", "PATH_IMAG", CONDI$, RUTARCHIVOX$, REGAF&)
    Else
        SQLX$ = "INSERT INTO ARCHIVOS_QR "
        SQLX$ = SQLX$ + " (NUMEROUSUARIO,PATH_IMAG,DOCUMENTOORIGEN)"
        SQLX$ = SQLX$ + " VALUES(" & User%
        SQLX$ = SQLX$ + " ,'" & RUTARCHIVOX$ & "'"
        SQLX$ = SQLX$ + " ,'" & DOCU_ORIGEN$ & "'"
        SQLX$ = SQLX$ + ")"
        '
        FLEXCONN.Execute (SQLX$)
        '
    End If
End Sub
Sub IMPRIMIR_ETIQUETAS_EYM(MSF As msFlexGrid, NORFA$, QR_OF%, ORDX%, SERIEX%)

        If QR_OF% > 0 Then
            NOF& = XVALO(VALOBADA("ORDEFABR", "NumeOrFa", "IDSUBOR = '" & NORFA$ & "'", "NOMESS"))
            PRINTETI.mnuMenu3.Caption = "Revertir Selección"
            PRINTETI.Label10(0) = "O.Fabricación"
            Call CREATABLA_ARCHIVOS_QR
        '    PRINTETI.mnuMenu3.Visible = True
            PRINTETI.Option5.Value = True
            PRINTETI.Text5 = SAFETEXT$(NOF&)
            PRINTETI.Show 1
        End If
        If (ORDX% + SERIEX%) > 0 Then
            FORIMPRE$ = Control$("", "STIKAPR2")
            FORIMPRE_SERIE$ = Control("", "STIKAPR3")
            If FORIMPRE$ = "" Then
              Call COMUNI("FALTA FORMULARIO DE IMPRESIÓN")
              Exit Sub
            End If
        Else
            Exit Sub
        End If
        '
        PARAMX1$ = "-"
        PARAMX2$ = "-"
        PARAMX3$ = "-"
        'PARAMX4$ = "-"
        k& = 0
        If ORDX% > 0 Then
            For U& = 1 To MSF.Rows - 1
                MARCA$ = MSF.TextMatrix(U&, 0)
                If MARCA$ = "*" Then
                      NUMSERIDES$ = MSF.TextMatrix(U&, 3)
                      ORDEX$ = Right(NUMSERIDES$, 3)
                      k& = k& + 1
                      Select Case k&
                         Case 1: PARAMX1$ = AJUCEN$(ORDEX$, 8)
                         Case 2: PARAMX2$ = AJUCEN$(ORDEX$, 8)
                         Case 3: PARAMX3$ = AJUCEN$(ORDEX$, 8)
                         'Case 4: PARAMX4$ = AJUCEN$(ORDEX$, 8)
                      End Select
                      If k& Mod 3 = 0 Then
                         GoSub IMPRIMIE_ETIQ
                         PARAMX1$ = "-"
                         PARAMX2$ = "-"
                         PARAMX3$ = "-"
                         'PARAMX4$ = "-"
                         k& = 0
                      End If
                End If
            Next U&
            If k& > 0 Then GoSub IMPRIMIE_ETIQ
            
        End If
        '
        k& = 0
        PARAMX1$ = "-"
        PARAMX2$ = "-"
        PARAMX3$ = "-"
        'PARAMX4$ = "-"
        ESETIQUETANSERIE% = 0
        If SERIEX% > 0 Then
            ESETIQUETANSERIE% = 1
            For U& = 1 To MSF.Rows - 1
                MARCA$ = MSF.TextMatrix(U&, 0)
                If MARCA$ = "*" Then
                      NUMSERIDES$ = MSF.TextMatrix(U&, 3)
                      k& = k& + 1
                      Select Case k&
                         Case 1: PARAMX1$ = AJUCEN$(NUMSERIDES$, 8)
                         Case 2: PARAMX2$ = AJUCEN$(NUMSERIDES$, 8)
                         Case 3: PARAMX3$ = AJUCEN$(NUMSERIDES$, 8)
                         'Case 4: PARAMX4$ = AJUCEN$(NUMSERIDES$, 8)
                      End Select
                      If k& Mod 3 = 0 Then
                         GoSub IMPRIMIE_ETIQ
                         PARAMX1$ = "-"
                         PARAMX2$ = "-"
                         PARAMX3$ = "-"
                         'PARAMX4$ = "-"
                         k& = 0
                      End If
                End If
            Next U&
            If k& > 0 Then GoSub IMPRIMIE_ETIQ
             
        End If
        '
        GoTo 90
         '
IMPRIMIE_ETIQ:
        If ESETIQUETANSERIE% > 0 Then FORIMPRE$ = FORIMPRE_SERIE$
        
        PUERIMPRE$ = Control("*", "PRIETIQ1")
        If Control(IDENTER$, "PRIETIQ1") <> "" Then PUERIMPRE$ = Control(IDENTER$, "PRIETIQ1")
        Call SETPRINPORT(PUERIMPRE$)
        FILTX$ = PARAMX1$ & ";" & PARAMX2$ & ";" & PARAMX3$ & ";" & PARAMX4$
        
        Call STDFORM(FORIMPRE$, FILTX$, "PRINT")

        Return
        
      
90      Call SETPRINPORT("RESTORE")

End Sub

Sub IMPRIMIR_ETIQUETAQR(FORIMPRE_SQR$, FILTX$, TEXTOQR$, NAME_ARCH_IMAG$, CANT_ETIQ&, IDENTIFICADOR$, IMPRESORA$, ANCHOX%, ALTOX%)
        'FORIMPRE_SQR$ = NOMBRE DEL FRM DE IMPRESION SQR
        'FILTX$ = EL FILTRO NECESARIO PARA EL .SQR (PARAMETROS)
        'TEXTOQR$ = CADENA DE TEXTO CON EL CUAL SE VA A GENERAR EL QR (TEXTOQR$ = NSERIE$)
        'NAME_ARCH_IMAG$ = NOMBRE DEL ARCHIVO QUE SE GENERA CON LA IMAGEN QR CON EXTENSION JPG (IMAG$ = "IMAGNSERIE_QR.JPG")
        'IDENTIFICADOR$= CON QUE NOMBRE SE VA A GUARDAR EN LA BASE DE DATOS PARA QUE LUEGO EN EL SQR SE BUSQUE LA IMAGEN (IDENTIFICADOR$ = "QR_NSERIE")
        'IMPRESORA$ = IMPRESORA POR DONDE SE VA A IMPRMIR
        'ANCHOX% = ANCHO CON EL CUAL SE VA A GENERAR LA IMAGEN
        'ALTOX% = ALTO CON EL IGUAL SE VA GENERAR LA IMAGEN
        
        '
        Call SETPRINPORT("RESTORE")
        If IMPRESORA$ <> "" Then
           Call SETPRINPORT(IMPRESORA$)
        End If

        On Error Resume Next
        Kill "C:\FLEXOFT\" & NAME_ARCH_IMAG$
        On Error GoTo 0
        'RUTINA QUE LLAMA A APLICACION DE .NET PARA GENERAR LA IMAGEN DEL CODIGO QR
        QRX$ = TRIM$(IMAGENCODBARRA$("C:\FLEXOFT\", NAME_ARCH_IMAG$, TEXTOQR$, IMPRESORA$, ANCHOX%, ALTOX%))
        Call T_Espera(2)
        Call GRABAR_QR("C:\FLEXOFT\" & NAME_ARCH_IMAG$, IDENTIFICADOR$)

        User% = USNBR% Mod 100
        If User% < 1 Then User% = 1
        FILTX$ = FILTX$ & ";" & IDENTIFICADOR$ & ";" & TRIM$(Str$(User%))
        Call STDFORM(FORIMPRE_SQR$, FILTX$, "PRINT", "", 0, CANT_ETIQ&)
        Call SETPRINPORT("RESTORE")

End Sub

Public Sub RecortaImagenQRFactura()
    
'    Dim reintento As Byte
'
'    If EXISTE%(App.Path + "\" + "RecortaImagen.exe") = 1 Then
'10      If EXISTE%(LUCOM$ & "IMAGFACTU_QR.JPG") = 1 Then
'            XXX = Shell(App.Path + "\" + "RecortaImagen.exe", 4)
'        ElseIf reintento < 3 Then
'            Call T_Espera(1 / 2)
'            reitento = reintento + 1
'            GoTo 10
'        End If
'    End If
    
    If EXISTE%(App.Path + "\" + "RecortaImagen.exe") = 1 Then
        XXX = Shell(App.Path + "\" + "RecortaImagen.exe", 4)
    End If
    
End Sub
