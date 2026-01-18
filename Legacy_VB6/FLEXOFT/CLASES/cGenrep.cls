VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cGenrep"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public RstTemporal As ADODB.Recordset
Private gCodReport As String
Private gDesReport As String
Private gTablaCantidadCampos As Integer
Private gListaCampos() As String
Private gListaFiltros() As Byte
Private gListaFiltrosTitulo() As String
Private gListaFormatoCampos() As String
Private gListaTituloCampo() As String
Private gListaTotalizaCampos() As Byte
Private gListaEcoCampos() As String
Private gListaCorteControl() As String
Private gListaFunction() As String
Private gListaValobada() As String
Private gListaCampoOculto() As Byte
Private gStatus As Boolean

Public Function GeneraReporte(CodigoReporte As String, DescripcionReporte As String, CamposConFormato As String, ConsultaSql As String) As Boolean

    Set RstTemporal = GeneraReporteConTablaTemporal(CodigoReporte, DescripcionReporte, CamposConFormato)
    
    Dim CABATEMP As ADODB.Recordset
    Set CABATEMP = READSET(ConsultaSql$)
    '
    Screen.MousePointer = 11
    '
    With CABATEMP
      On Error Resume Next
      .MoveLast
      If Err > 0 Then
        On Error GoTo 0
        Call MENSERR(24, "Sin Movimientos")
        GoTo 99
      End If
      '
      On Error GoTo 0
      .MoveFirst
      Do While Not .EOF
        RstTemporal.AddNew
            
        IdCampo% = -1: CantidaValobadas% = 0
        For i% = 1 To RstTemporal.Fields.Count
            
            If gListaValobada(i%) = "" Then
                IdCampo% = IdCampo% + 1
            Else
                CantidaValobadas% = CantidaValobadas% + 1
            End If
            
            If Mid$(gListaFormatoCampos(i%), 1, 1) = "A" Then   ' DATO ALFANUMERICO. CONSIDERA LONGITUD MAXIMA DEL CAMPO
                If gListaFunction(i%) = "" Then
                    If gListaValobada(i%) = "" Then
                        RstTemporal.Fields(i% - 1) = Mid$(CABATEMP.Fields(IdCampo%), 1, XVALO(Mid$(gListaFormatoCampos(i%), 2)))
                    End If
                Else
                    If InStr(1, gListaFunction(i%), "TICAMONE") > 0 Then
                        Moneda% = XVALO(Mid$(gListaFunction(i%), 10))
                        RstTemporal.Fields(i% - 1) = TICAMONE(Moneda%)
                    End If
                End If
            Else
                If gListaFunction(i%) = "" Then
                    If gListaValobada(i%) = "" Then
                        RstTemporal.Fields(i% - 1) = CABATEMP.Fields(IdCampo%)
                    End If
                Else
                    If InStr(1, gListaFunction(i%), "TICAMONE") > 0 Then
                        ParametroColumna% = XVALO(Mid$(gListaFunction(i%), 10))
                        Moneda% = CABATEMP.Fields(ParametroColumna% - CantidaValobadas% - 1) ' OBVIA LOS VALOBADAS QUE NO SE REGITRAN COMO CAMPOS EN LA TABLA
                        RstTemporal.Fields(i% - 1) = TICAMONE(Moneda%)
                    End If
                    If InStr(1, gListaFunction(i%), "CONVDESCACOEF") > 0 Then
                        RstTemporal.Fields(i% - 1) = 1 - XVALO(COEFDESCU(CABATEMP.Fields(IdCampo%))) / 100
                    End If
                End If
            End If
        Next i%
        
        RstTemporal.Update
        
        .MoveNext
      Loop
    End With
    RstTemporal.Close
    Set RstTemporal = Nothing
    GeneraReporte = True
    
99  CABATEMP.Close
    Set CABATEMP = Nothing
    
    Screen.MousePointer = 1
        
End Function

Public Function GeneraReporteConTablaTemporal(CodigoReporte As String, DescripcionReporte As String, CamposConFormato As String) As ADODB.Recordset

    ReDim gListaCampos(0), gListaFiltros(0), gListaFiltrosTitulo(0), gListaFormatoCampos(0), gListaTituloCampo(0), gListaTotalizaCampos(0), gListaEcoCampos(0), gListaCorteControl(0), gListaFunction(0), gListaValobada(0), gListaCampoOculto(0)
    
    Dim ContadorCamposEcos%
    
    ' OBTIENE DATOS PARA CREAR EL NUEVO REPORTE
    gCodReport = CodigoReporte
    gDesReport = DescripcionReporte
    gStatus = False     ' ESTADO DE LA CREACION DE LA TABLA Y EL REPORTE
    If TRIM$(CamposConFormato) <> "" Then
        CamposCF$ = CamposConFormato
        n% = InStr(1, CamposCF$, ";")
        While n% > 0
        
            UnCampo$ = Mid$(CamposCF$, 1, n% - 1)
            m% = InStr(1, UnCampo$, "/")
            If m% > 0 Then
                FormatoCampo$ = Mid$(UnCampo$, m% + 1)
            Else
                GoTo Err_FaltaFormatoEnCampo
            End If
            i% = UBound(gListaCampos)
            ReDim Preserve gListaCampos(i% + 1), gListaFiltros(i% + 1), gListaFiltrosTitulo(i% + 1), gListaFormatoCampos(i% + 1), gListaTituloCampo(i% + 1), gListaTotalizaCampos(i% + 1), gListaEcoCampos(i% + 1), gListaCorteControl(i% + 1), gListaFunction(i% + 1), gListaValobada(i% + 1), gListaCampoOculto(i% + 1)
            
            r% = InStr(1, UnCampo$, "(")
            If r% <= 0 Then r% = m%
            NombreDelCampo$ = Mid$(UnCampo$, 1, r% - 1)
            If r% < m% Then
                SetupCampo$ = Mid$(UnCampo$, r% + 1, m% - r% - 2)
                t% = InStr(1, SetupCampo$, "|")
                While t% > 0
                    UnSetup$ = Mid$(SetupCampo$, 1, t% - 1)
                    s% = InStr(1, UnSetup$, "=")
                    If s% > 0 Then CodigoSetup$ = TRIM$(Mid$(UnSetup$, 1, s% - 1))
                
                    Select Case UCase(CodigoSetup$)
                        ' ENCABEZADO
                        Case "T": gListaTituloCampo(i% + 1) = TRIM$(Mid$(UnSetup$, s% + 1))
                        ' TOTALIZA COLUMNA
                        Case "TT": gListaTotalizaCampos(i% + 1) = 1
                        ' FILTRO
                        Case "FILTRO": gListaFiltros(i% + 1) = 1: gListaFiltrosTitulo(i% + 1) = TRIM$(Mid$(SetupCampo$, s% + 1))
                        ' ECO
                        Case "ECO": gListaEcoCampos(i% + 1) = TRIM$(Mid$(SetupCampo$, s% + 1))
                        ' CORTE DE CONTROL
                        Case "CC": gListaCorteControl(i% + 1) = TRIM$(Mid$(SetupCampo$, s% + 1))
                        ' OCULTA CAMPO
                        Case "HIDE": gListaCampoOculto(i% + 1) = 1
                        ' FUNCIONES
                        Case "FUNCTION": gListaFunction(i% + 1) = TRIM$(Mid$(SetupCampo$, s% + 1))
                        ' VALOBADA
                        Case "VALOBADA": gListaValobada(i% + 1) = TRIM$(Mid$(SetupCampo$, s% + 1))
                    End Select
                    
                    SetupCampo$ = TRIM$(Mid$(SetupCampo$, t% + 1))
                    t% = InStr(1, SetupCampo$, "|")
                Wend
                
                If SetupCampo$ <> "" Then
                    s% = InStr(1, SetupCampo$, "=")
                    If s% > 0 Then CodigoSetup$ = TRIM$(Mid$(SetupCampo$, 1, s% - 1))
                
                    Select Case CodigoSetup$
                        ' ENCABEZADO
                        Case "T": gListaTituloCampo(i% + 1) = TRIM$(Mid$(SetupCampo$, s% + 1))
                        ' TOTALIZA COLUMNA
                        Case "TT": gListaTotalizaCampos(i% + 1) = 1
                        ' FILTRO
                        Case "FILTRO": gListaFiltros(i% + 1) = 1: gListaFiltrosTitulo(i% + 1) = TRIM$(Mid$(SetupCampo$, s% + 1))
                        ' ECO
                        Case "ECO": gListaEcoCampos(i% + 1) = TRIM$(Mid$(SetupCampo$, s% + 1))
                        ' CORTE DE CONTROL
                        Case "CC": gListaCorteControl(i% + 1) = TRIM$(Mid$(SetupCampo$, s% + 1))
                        ' OCULTA CAMPO
                        Case "HIDE": gListaCampoOculto(i% + 1) = 1
                        ' FUNCIONES
                        Case "FUNCTION": gListaFunction(i% + 1) = TRIM$(Mid$(SetupCampo$, s% + 1))
                        ' VALOBADA
                        Case "VALOBADA": gListaValobada(i% + 1) = TRIM$(Mid$(SetupCampo$, s% + 1))
                    End Select
                
                End If
            End If
            gListaCampos(i% + 1) = TRIM$(NombreDelCampo$)
            gListaFormatoCampos(i% + 1) = TRIM$(FormatoCampo$)
            If gListaTituloCampo(i% + 1) = "" Then gListaTituloCampo(i% + 1) = gListaCampos(i% + 1)
            
            CamposCF$ = TRIM$(Mid$(CamposCF$, n% + 1))
            n% = InStr(1, CamposCF$, ";")
            If n% = 0 And CamposCF$ <> "" Then      ' ULTIMO CAMPO
                CamposCF$ = CamposCF$ & " "
                n% = Len(CamposCF$)
            End If
            
        Wend
        
    End If
    
    gTablaCantidadCampos = UBound(gListaCampos)
    If gTablaCantidadCampos > 0 Then
        
        ' CREA EL REPORTE
        IdTerm& = IdTerminal()
        CadenaReporte$ = "CODQUERY = " & CodigoReporte & "|"
        CadenaReporte$ = CadenaReporte$ & "DESQUERY = " & DescripcionReporte & "|"
        CadenaReporte$ = CadenaReporte$ & "TIPQUERY =    0|"
        CadenaReporte$ = CadenaReporte$ & "SELTABLE = TEMP_" & CodigoReporte & CStr(IdTerm&) & "|"
        
        For i% = 1 To gTablaCantidadCampos
            If gListaFiltros(i%) = 0 Then
                ' AGREGA LAS COLUMNAS
                UnCampo$ = Mid$(gListaCampos(i%) & Space(64), 1, 64)
                FormatoCampo$ = Mid$(gListaFormatoCampos(i%) & Space(8), 1, 8)
                CadenaReporte$ = CadenaReporte$ & "SELFIELD = " & UnCampo$ & FormatoCampo$ & CStr(gListaCampoOculto(i%)) & " 0" & CStr(gListaTotalizaCampos(i%)) & "    " & AJUSTI(gListaTituloCampo(i%), 64)
                ' AGREGA LOS ECOS
                If gListaEcoCampos(i%) <> "" Then
                    SeparaFormato% = InStr(1, gListaEcoCampos(i%), "-")
                    If SeparaFormato% > 0 Then
                        FormatoEco$ = TRIM$(Mid$(gListaEcoCampos(i%), SeparaFormato% + 1))
                        CampoEco$ = TRIM$(Mid$(gListaEcoCampos(i%), 1, SeparaFormato% - 1))
                        CadenaReporte$ = CadenaReporte$ & AJUSTI(CampoEco$, 16) & AJUSTI(FormatoEco$, 8) & "|"
                        ContadorCamposEcos% = ContadorCamposEcos% + 1
                    End If
                End If
                ' AGREGAS LOS CORTES DE CONTROL
                If gListaCorteControl(i%) <> "" Then
                    SeparaFormato% = InStr(1, gListaCorteControl(i%), "-")
                    If SeparaFormato% > 0 Then
                        FormatoCorte$ = TRIM$(Mid$(gListaCorteControl(i%), SeparaFormato% + 1))
                        ColumnaCorte% = XVALO(Mid$(gListaCorteControl(i%), 1, SeparaFormato% - 1))
                        If ContadorCamposEcos% > 0 Then ColumnaCorte% = ColumnaCorte% + ContadorCamposEcos% ' LOS ECOS SON CONSIDERADOS EN LOS CORTES DE CONTROL
                        CampoCorte$ = "SUTOT/" & CStr(ColumnaCorte%)
                    Else
                        FormatoCorte$ = FormatoCampo$   ' POR DEFECTO ESTABLECE EL FORMATO DEL CAMPO QUE ESTA TOTALIZANDO
                        ColumnaCorte% = XVALO(gListaCorteControl(i%))
                        If ContadorCamposEcos% > 0 Then
                            CamposEcosAnteriores% = 0
                            For r% = 1 To i%
                                If TRIM$(gListaEcoCampos(i%)) <> "" Then CamposEcosAnteriores% = CamposEcosAnteriores% + 1
                            Next r%
                            ColumnaCorte% = ColumnaCorte% + CamposEcosAnteriores% ' LOS ECOS SON CONSIDERADOS EN LOS CORTES DE CONTROL
                        End If
                        CampoCorte$ = "SUTOT/" & CStr(ColumnaCorte%)
                    End If
                    CadenaReporte$ = CadenaReporte$ & AJUSTI(CampoCorte$, 16) & AJUSTI(FormatoCorte$, 8) & "|"
                End If
                ' AGREGA VALOBADA
                If gListaValobada(i%) <> "" Then
                
                    SetupValobada$ = gListaValobada(i%)
                    Separa% = InStr(1, SetupValobada$, ",")
                    TablaValobada$ = TRIM$(Mid$(SetupValobada$, 1, Separa% - 1))
                    
                    SetupValobada$ = Mid$(SetupValobada$, Separa% + 1)
                    Separa% = InStr(1, SetupValobada$, ",")
                    CampoValobada$ = TRIM$(Mid$(SetupValobada$, 1, Separa% - 1))
                    
                    SetupValobada$ = Mid$(SetupValobada$, Separa% + 1)
                    CondicionValobada$ = TRIM$(SetupValobada$)
                    
                    If gListaEcoCampos(i%) = "" And gListaCorteControl(i%) = "" Then
                        CadenaReporte$ = CadenaReporte$ & AJUSTI(" ", 16) & AJUSTI(" ", 8)
                    Else
                        CadenaReporte$ = Mid$(CadenaReporte$, 1, Len(CadenaReporte$) - 1)   ' LE QUITA EL ULTIMO CARACTER (|)
                    End If
                    
                    CadenaReporte$ = CadenaReporte$ & AJUSTI(TablaValobada$, 64) & AJUSTI(CampoValobada$, 64) & AJUSTI(CondicionValobada$, 512) & "|"
                   
                End If
                CadenaReporte$ = CadenaReporte$ & "|"
            End If
        Next i%

        ' AGREGA LOS FILTROS
        For i% = 1 To gTablaCantidadCampos
            If gListaFiltros(i%) = 1 Then
                UnCampo$ = Mid$(gListaCampos(i%) & Space(64), 1, 25)
                TituloFiltro$ = Right(Space(25) & "¿" & REPLACAR$(gListaFiltrosTitulo(i%), "-", "/") & "?", 25)
                CadenaReporte$ = CadenaReporte$ & "CONDITIO = " & UnCampo$ & " = " & TituloFiltro$ & "|"
            End If
        Next i%
        
        ' AGREGA O ACTUALIZA EL REPORTE
        Call NuevoReporte(CodigoReporte, DescripcionReporte, 0, CadenaReporte$)
        
        
        ' BORRA TABLA TEMPORAL
        On Error GoTo Err_ImposibleBorrarTabla
        TablaTemporal$ = TRIM$(CodigoReporte) & CStr(IdTerm&)
        FLEXCONN.Execute "DROP TABLE TEMP_" & TablaTemporal$
        On Error GoTo 0
           
        ' CREA TABLA TEMPORAL
        QueryCreate$ = "CREATE TABLE TEMP_" & CodigoReporte & CStr(IdTerm&) & "("
        For i% = 1 To gTablaCantidadCampos
            UnCampo$ = gListaCampos(i%)
            FormatoCampo$ = gListaFormatoCampos(i%)
            
            CodigoFormato$ = Mid$(FormatoCampo$, 1, 1)
            Tamano$ = TRIM$(Mid$(FormatoCampo$, 2))
            Select Case CodigoFormato$
            
                Case "A": TipoSql$ = "NVARCHAR(" & Tamano & ")"
                Case "I": TipoSql$ = "INT"
                Case "C": TipoSql$ = "INT"                          ' UTILIZADO PARA EL CODIGO INTERNO DE UN ARTICULO
                Case "L": TipoSql$ = "BIGINT"
                Case "F": TipoSql$ = "FLOAT"
                Case "G": TipoSql$ = "FLOAT"
                Case "Z": TipoSql$ = "DECIMAL"
                Case "M": TipoSql$ = "MONEY"
                Case "D": TipoSql$ = "SMALLDATETIME"
                Case "T": TipoSql$ = "TIME"
            
            End Select
            
            QueryCreate$ = QueryCreate$ & UnCampo$ & " " & TipoSql$ & ","
        
        Next i%

        QueryCreate$ = TRIM$(QueryCreate$)
        QueryCreate$ = Mid$(QueryCreate$, 1, Len(QueryCreate$) - 1)
        QueryCreate$ = QueryCreate$ & ")"
        
        On Error GoTo Err_CreaTablaTemporal
        FLEXCONN.Execute QueryCreate$
        On Error GoTo 0
        
        
        ' OBTIENE UN RECORDSET DE ESCRITURA DE LA TABLA TEMPORAL
        On Error GoTo Err_CreaObjetoRecordset
        ConsultaSql$ = "SELECT * FROM TEMP_" & CodigoReporte$ & CStr(IdTerm&)
        Set RstTemporal = New ADODB.Recordset
        RstTemporal.Open ConsultaSql$, FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
        On Error GoTo 0
        
        
        ' RETORNA EL RECORDSET A LA RUTINA LLAMADORA
        Set GeneraReporteConTablaTemporal = RstTemporal
        gStatus = True
    End If

99  Exit Function



Err_FaltaFormatoEnCampo:
    Set GeneraReporteConTablaTemporal = Nothing
    MsgBox "Falta Formato de Campo"
    GoTo 99

Err_ImposibleBorrarTabla:
    Set GeneraReporteConTablaTemporal = Nothing
    Resume Next
        
Err_CreaTablaTemporal:
    Set GeneraReporteConTablaTemporal = Nothing
    MsgBox Err.Description
    GoTo 99
        
Err_CreaObjetoRecordset:
    Set GeneraReporteConTablaTemporal = Nothing
    MsgBox Err.Description
    GoTo 99

End Function

Private Sub NuevoReporte(CodigoReporte As String, DescripcionReporte As String, TipoReporte As Integer, Query As String)

    ' AGREGA EL NUEVO REPORTE A LA TABLA SQL_QUERIES
    Dim RST As New ADODB.Recordset
    ConsultaSql$ = "SELECT * FROM SQL_QUERIES WHERE COD_QUERY = '" & CodigoReporte & "'"
    RST.Open ConsultaSql$, FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
    With RST
        If .EOF Then
            .AddNew
            !COD_QUERY = CodigoReporte
            !CodiEmpr = CodiEmp%
        Else
            ' SI YA EXISTE, ACTUALIZA SOLO SI EL REPORTE ES DISTINTO
            If TRIM$(!PAR_QUERY) = TRIM$(Query) Then GoTo 80
        End If
        !DES_QUERY = DescripcionReporte
        !TIP_QUERY = TipoReporte
        !PAR_QUERY = Query
        .Update
    End With
80  RST.Close
    Set RST = Nothing
        
End Sub

Public Sub MostrarReporte()

    If gStatus Then
        ' SE DESCONECTA DE LA TABLA TEMPORAL Y ELIMINA EL OBJETO
        On Error Resume Next
        RstTemporal.Close
        Set RstTemporal = Nothing
        On Error GoTo 0
        '
        ' MUESTRA EL REPORTE
        MODULO$ = "?" + gCodReport
        Descripcion$ = gDesReport
        Call CONECRUN(MODULO$, Descripcion$)
    End If

End Sub


Private Function IdTerminal() As Long

    ' OBTIENE EL ID DE LA TERMINAL ALMACENADO EN EL ARCHIVO IDENTER.DRV
    Dim RegistroEncriptado As String * 40
    
    n% = FILEOPEN%(LUCOM$ + "IDENTER.DRV", 0, 40)
    Get #n%, 1, RegistroEncriptado$
    RegistroDesEncriptado$ = DESENCRI$(Left$(RegistroEncriptado$, 38))
    
    IdTerminal = XVALO(Mid$(RegistroDesEncriptado$, 33, 6))
    
End Function


Private Sub Class_Terminate()

    ' SE DESCONECTA DE LA TABLA TEMPORAL Y ELIMINA EL OBJETO
    On Error Resume Next
    RstTemporal.Close
    Set RstTemporal = Nothing
    On Error GoTo 0

End Sub
