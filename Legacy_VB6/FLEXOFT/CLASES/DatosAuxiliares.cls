VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DatosAuxiliares"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Private Sub Class_Initialize()

    Call CrearTablasAuxiliares
    
End Sub

    
Public Sub CrearTabla(nombre As String, Descripcion As String, Campos() As eDatosAuxiliaresCampos)

    Call ABRECONEXION
    Call CREATABLA_ENCAPRF
    
    If Descripcion = "" Then Descripcion = nombre
    CONDI$ = " LISTA =   '" & nombre & "'"
    EXILISTA% = CantRegistros&("ENCAPRFUN", CONDI$, "NOMESS")
    If EXILISTA% > 0 Then
      FE% = 16465
      FEBD% = CVINT(VALOBADA("ENCAPRFUN", "FECHEMISI", CONDI$, "NOMESS"))
      GENERAREST% = 0
      If FEBD% < FE% Then
        IDLISTAX& = XVALO(VALOBADA("ENCAPRFUN", "ID_LISTA", CONDI$, "NOMESS"))
        Call BORRAREGISTROS("ENCAPRFUN", "ID_LISTA = " & IDLISTAX&, REGAF&, "NOMESS")
        Call BORRAREGISTROS("ENCAPRF", "ID_LISTA = " & IDLISTAX&, REGAF&, "NOMESS")
        GENERAREST% = 1
      End If
    Else
      GENERAREST% = 1
      CONDI$ = " ID_LISTA  > 0 "
      IDLISTAX& = 1 + XVALO(VALOBADA("ENCAPRFUN", "MAX(ID_LISTA)", CONDI$, "NOMESS"))
    End If
    If GENERAREST% > 0 Then
      SQLX$ = "INSERT INTO  ENCAPRFUN (ID_LISTA ,LISTA, Descripcion, ADDREGISTROS, USAFORMULA, BORRAR, INSERTAR, NOTILDE "
      SQLX$ = SQLX$ + ", CANTICAMPOS, CodiEmpr, USARECORD, FECHEMISI, TITURECORD, ConsultaSql)"
      SQLX$ = SQLX$ + "VALUES ( " & IDLISTAX&
      SQLX$ = SQLX$ + ",'" + nombre + "'"
      SQLX$ = SQLX$ + ",'" + Descripcion + "'"
      SQLX$ = SQLX$ + ",1"
      SQLX$ = SQLX$ + ",1"
      SQLX$ = SQLX$ + " ,1"
      SQLX$ = SQLX$ + " ,1"
      SQLX$ = SQLX$ + " ,0"
      SQLX$ = SQLX$ + " ," + CStr(UBound(Campos))
      SQLX$ = SQLX$ + " ,0"
      SQLX$ = SQLX$ + " ,0"
      SQLX$ = SQLX$ + " ,'" + Format(Now, "dd/mm/yyyy") + "'"
      SQLX$ = SQLX$ + " ,''"
      SQLX$ = SQLX$ + " ,'SELECT "
      For i% = 0 To UBound(Campos) - 1
        If i% > 0 Then SQLX$ = SQLX$ + ", "
        SQLX$ = SQLX$ + "CAMPO" + CStr(i% + 1)
      Next i%
      SQLX$ = SQLX$ + " FROM DATASQL WITH(NOLOCK)'"
      SQLX$ = SQLX$ + ")"
      FLEXCONN.Execute (SQLX$)
    
      Dim A As eDatosAuxiliaresCampos
      For i% = 0 To UBound(Campos) - 1
      
        SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
        SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
        SQLX$ = SQLX$ + " ," + CStr(CODIEMPR%)
        SQLX$ = SQLX$ + " ," + CStr(Campos(i%).MarcaBorrado)
        SQLX$ = SQLX$ + " , '" + Campos(i%).Encabezado + "'"
        SQLX$ = SQLX$ + " , '" + Campos(i%).Formula + "'"
        SQLX$ = SQLX$ + " , '" + Campos(i%).CampoEditable + "'"
        SQLX$ = SQLX$ + " , '" + Campos(i%).ColumnaCampoEco + "'"
        SQLX$ = SQLX$ + " , '" + Campos(i%).Total + "'"
        SQLX$ = SQLX$ + " , '" + Campos(i%).DejarLibre + "'"
        SQLX$ = SQLX$ + " , '" + Campos(i%).Formato + "'"
        SQLX$ = SQLX$ + " )"
        FLEXCONN.Execute (SQLX$)
        
      Next i%
    
    End If
    '|FIN ARMAINFAGRU

End Sub

Public Function ObtenerIdTablaAuxiliar(NombreTabla As String) As Long
    
    ObtenerIdTablaAuxiliar = XVALO(VALOBADA("ENCAPRFUN", "ID_LISTA", "LISTA = '" & NombreTabla & "'", "NOMESS"))
    
End Function

Public Function ListaSelDataSql(nombre As String, Optional TITULO As String, Optional CamposUsuario As String, Optional Distinct As Boolean = False) As Boolean

    ' OBTIENE ID DE LA TABLA, DESCRIPCION Y CONSULTA PARA DATASQL
    Dim RstENCAPRFUN As ADODB.Recordset
    ConsultaSql$ = "SELECT ID_LISTA, LISTA, DESCRIPCION, CONSULTASQL FROM ENCAPRFUN WITH(NOLOCK) WHERE LISTA='" & nombre & "'"
    Set RstENCAPRFUN = READSET(ConsultaSql$)
    If RstENCAPRFUN.EOF And RstENCAPRFUN.BOF Then
        ListaSeleccionDataSql = False
        Exit Function
    End If
    ID_Lista& = RstENCAPRFUN!ID_Lista
    Descripcion$ = RstENCAPRFUN!Descripcion
    ConsultaSqlListaSeleccion$ = RstENCAPRFUN!ConsultaSql
    RstENCAPRFUN.Close
    Set RstENCAPRFUN = Nothing
    
    If TRIM$(CamposUsuario) <> "" Then
        Campos$ = CamposUsuario
        Call TEXTOLOTES$(Campos$, ";")
        CantidadDeCampos% = UBound(CADENATEX)
        ConsultaSqlListaSeleccion$ = "SELECT "
        For ICAMPO = 1 To UBound(CADENATEX)
            NOMBRECAMPO$ = Mid$(CADENATEX(ICAMPO), 1, InStr(1, CADENATEX(ICAMPO), "/") - 1)
            ConsultaSqlListaSeleccion$ = ConsultaSqlListaSeleccion$ & NOMBRECAMPO$ & ","
        Next ICAMPO
        ConsultaSqlListaSeleccion$ = Mid$(ConsultaSqlListaSeleccion$, 1, Len(ConsultaSqlListaSeleccion$) - 1)
        ConsultaSqlListaSeleccion$ = ConsultaSqlListaSeleccion$ & " FROM DATASQL WITH(NOLOCK)"
        GoTo 50
    End If
    
    ' OBTIENE LISTA DE CAMPOS PARA ARMAR GRILLA
    Dim RstENCAPRF As ADODB.Recordset
    ConsultaSql$ = "SELECT ENCABEZADO FROM ENCAPRF WITH(NOLOCK) WHERE ID_LISTA=" & CStr(ID_Lista&)
    Set RstENCAPRF = READSET(ConsultaSql$)
    If RstENCAPRF.EOF And RstENCAPRF.BOF Then
        ListaSeleccionDataSql = False
        Exit Function
    End If
    CantidadDeCampos% = 0
    With RstENCAPRF
        Do While Not .EOF
            Campos$ = Campos$ & !Encabezado & ";"
            CantidadDeCampos% = CantidadDeCampos% + 1
            .MoveNext
        Loop
    End With
    Campos$ = Mid$(Campos$, 1, Len(Campos$) - 1)
    RstENCAPRF.Close
    Set RstENCAPRF = Nothing

    ' ARMA ZELECHO CON LOS DATOS DE LA TABLA
50  Call CARGA_ENCABEZADO(FRMZELECHO.MSF2, Campos$, FRMZELECHO)
    Call CARGA_ENCABEZADO(FRMZELECHO.MSF1, Campos$, FRMZELECHO)
    FRMZELECHO.Caption = Descripcion$
    If TITULO <> "" Then FRMZELECHO.Caption = TITULO
    
    ' CARGA GRILLA CON LOS DATOS DE LA TABLA
    Dim Rst As ADODB.Recordset
    If Distinct = True Then
        ConsultaSqlListaSeleccion$ = Mid$(ConsultaSqlListaSeleccion$, 1, 6) & " distinct " & Mid$(ConsultaSqlListaSeleccion$, 7)
    End If
    ConsultaSql$ = ConsultaSqlListaSeleccion$ & " WHERE ID_LISTA=" & CStr(ID_Lista&)
    Set Rst = READSET(ConsultaSql$)
    With Rst
        Do While Not .EOF
            FRMZELECHO.MSF2.Rows = FRMZELECHO.MSF2.Rows + 1
            For i% = 0 To CantidadDeCampos% - 1
                FRMZELECHO.MSF2.TextMatrix(FRMZELECHO.MSF2.Rows - 1, i% + 1) = .Fields(i%)
            Next i%
            .MoveNext
        Loop
    End With
    On Error Resume Next
    Rst.Close
    Set Rst = Nothing
    On Error GoTo 0
    
    ' DETERMINA ANCHO A PARTIR DEL ANCHO DE LA GRILLA
    For i = 0 To FRMZELECHO.MSF2.Cols - 1
        AnchoColumnas# = AnchoColumnas# + FRMZELECHO.MSF2.ColWidth(i)
    Next i
    AnchoAnterior# = FRMZELECHO.Width
    FRMZELECHO.Width = AnchoColumnas# + 200
    
    ' DESPLIEGA LISTA
    FRMZELECHO.Show 1
    FRMZELECHO.Width = AnchoAnterior#
    
    Unload FRMZELECHO
    
    ListaSelDataSql = True

End Function

Public Function ListaSeleccionTablaAuxiliar(nombre As String, Optional TITULO As String, Optional Campos As String, Optional Distinct As Boolean = False) As Boolean

    ListaSeleccionTablaAuxiliar = ListaSelDataSql(nombre, TITULO, Campos, Distinct)
    
End Function

Function EcoTablaAuxiliar(nombre As String, Argumento As String, NumeroCampoBusca As Integer, NumeroCampoEco As Integer) As String
    
    If NumeroCampoBusca = 0 Then NumeroCampoBusca = 1
    CampoBusca$ = "CAMPO" & CStr(NumeroCampoBusca)
    If NumeroCampoEco = 0 Then NumeroCampoEco = 2
    CampoEco$ = "CAMPO" & CStr(NumeroCampoEco)
    ID_Lista& = ObtenerIdTablaAuxiliar(nombre)
    If ID_Lista& > 0 Then
        ECO$ = VALOBADA("DATASQL", CampoEco$, "ID_LISTA=" & CStr(ID_Lista&) & " AND ltrim(rtrim(" & CampoBusca$ & "))='" & Argumento & "'")
        EcoTablaAuxiliar = ECO$
    End If
    
End Function

Public Sub ListarTodasLasTablasAuxiliares()

  Call CARGA_ENCABEZADO(FRMZELECHO.MSF2, "ID LISTA/F12;NOMBRE LISTA/A32;DESCRIPCION/A48", FRMZELECHO)
  Call CARGA_ENCABEZADO(FRMZELECHO.MSF1, "ID LISTA/F12;NOMBRE LISTA/A32;DESCRIPCION/A48", FRMZELECHO)
  FRMZELECHO.Caption = "LISTA DE ESTRUCTURAS PRF Y RFS"

  Dim obj As New RECORXET
  Set rs1 = obj.RS_ENCAPRFLISTAS()
  Call Carga_MSF_Recordset(rs1, "ID LISTA/F12;NOMBRE LISTA/A32;DESCRIPCION/A48", FRMZELECHO.MSF2)
  Set obj = Nothing
  rs1.Close
  Set rs1 = Nothing
  EPAC$ = UCase$(TRIM$(EMPREAC$))

  FRMZELECHO.Width = 8500
  FRMZELECHO.Show 1
  IDLIST& = XVALO(RESP_ZELE_VECT(1))

  Call ABM_TABLA(IDLIST&)

End Sub

Public Sub MostrarTablaAuxilar(nombre As String)

    IDLIST& = ObtenerIdTablaAuxiliar(nombre)
    If IDLIST& > 0 Then Call ABM_TABLA(IDLIST&)
    
End Sub


Public Function CrearTablasAuxiliares()

    Dim ListaCampos() As eDatosAuxiliaresCampos
    Dim CAMPO As eDatosAuxiliaresCampos
    
    Call GENERAR_ESTRUCTURAS_SQL
    
    ' STOCK DE SEGURIDAD POR RANGO
    ' ****************************
    Set CAMPO = New eDatosAuxiliaresCampos
    CAMPO.Encabezado = "Rango/A16"
    CAMPO.Formula = ""
    CAMPO.CampoEditable = "S"
    CAMPO.ColumnaCampoEco = "codrango;descripcion;codigorango;codigorango_id>0;group by codrango, descripcion order by codrango"
    CAMPO.Total = ""
    CAMPO.DejarLibre = ""
    CAMPO.MarcaBorrado = 0
    n = n + 1
    ReDim Preserve ListaCampos(n)
    Set ListaCampos(n - 1) = CAMPO
        
    Set CAMPO = New eDatosAuxiliaresCampos
    CAMPO.Encabezado = "Denominacion/A48"
    CAMPO.Formula = ""
    CAMPO.CampoEditable = "N"
    CAMPO.ColumnaCampoEco = ""
    CAMPO.Total = ""
    CAMPO.DejarLibre = ""
    CAMPO.MarcaBorrado = 0
    n = n + 1
    ReDim Preserve ListaCampos(n)
    Set ListaCampos(n - 1) = CAMPO
    
    Set CAMPO = New eDatosAuxiliaresCampos
    CAMPO.Encabezado = "Stock Seg/F14.0"
    CAMPO.Formula = ""
    CAMPO.CampoEditable = "S"
    CAMPO.ColumnaCampoEco = ""
    CAMPO.Total = ""
    CAMPO.DejarLibre = ""
    CAMPO.MarcaBorrado = 0
    n = n + 1
    ReDim Preserve ListaCampos(n)
    Set ListaCampos(n - 1) = CAMPO
    
    Call CrearTabla("STOSEGRANGO", "Stock de Seguridad por Rango", ListaCampos)
    
    ' COMPENSACION CLIENTE-PROVEEDOR
    ' ******************************
    Set CAMPO = New eDatosAuxiliaresCampos
    CAMPO.Encabezado = "Ord/I10"
    CAMPO.Formula = ""
    CAMPO.CampoEditable = "N"
    CAMPO.ColumnaCampoEco = ""
    CAMPO.Total = ""
    CAMPO.DejarLibre = ""
    CAMPO.MarcaBorrado = 0
    n = 1
    ReDim Preserve ListaCampos(n)
    Set ListaCampos(n - 1) = CAMPO
    
    Set CAMPO = New eDatosAuxiliaresCampos
    CAMPO.Encabezado = "Cta/I10"
    CAMPO.Formula = ""
    CAMPO.CampoEditable = "S"
    CAMPO.ColumnaCampoEco = "nume_cli;raso_cli;deudor12;stat_cli>0;order by raso_cli"
    CAMPO.Total = ""
    CAMPO.DejarLibre = ""
    CAMPO.MarcaBorrado = 0
    n = n + 1
    ReDim Preserve ListaCampos(n)
    Set ListaCampos(n - 1) = CAMPO
        
    Set CAMPO = New eDatosAuxiliaresCampos
    CAMPO.Encabezado = "Cliente/A32"
    CAMPO.Formula = ""
    CAMPO.CampoEditable = "N"
    CAMPO.ColumnaCampoEco = ""
    CAMPO.Total = ""
    CAMPO.DejarLibre = ""
    CAMPO.MarcaBorrado = 0
    n = n + 1
    ReDim Preserve ListaCampos(n)
    Set ListaCampos(n - 1) = CAMPO
    
    Set CAMPO = New eDatosAuxiliaresCampos
    CAMPO.Encabezado = "Cta/I10"
    CAMPO.Formula = ""
    CAMPO.CampoEditable = "S"
    CAMPO.ColumnaCampoEco = "nume_cli;raso_cli;proved12;stat_cli>0;order by raso_cli"
    CAMPO.Total = ""
    CAMPO.DejarLibre = ""
    CAMPO.MarcaBorrado = 0
    n = n + 1
    ReDim Preserve ListaCampos(n)
    Set ListaCampos(n - 1) = CAMPO
        
    Set CAMPO = New eDatosAuxiliaresCampos
    CAMPO.Encabezado = "Proveedor/A32"
    CAMPO.Formula = ""
    CAMPO.CampoEditable = "N"
    CAMPO.ColumnaCampoEco = ""
    CAMPO.Total = ""
    CAMPO.DejarLibre = ""
    CAMPO.MarcaBorrado = 0
    n = n + 1
    ReDim Preserve ListaCampos(n)
    Set ListaCampos(n - 1) = CAMPO
    
    Call CrearTabla("COMPECLIPRO", "", ListaCampos)
    
    ' COMPENSACION DE IMPUESTOS
    ' *************************
    Set CAMPO = New eDatosAuxiliaresCampos
    CAMPO.Encabezado = "Ord/I10"
    CAMPO.Formula = ""
    CAMPO.CampoEditable = "N"
    CAMPO.ColumnaCampoEco = ""
    CAMPO.Total = ""
    CAMPO.DejarLibre = ""
    CAMPO.MarcaBorrado = 0
    n = 1
    ReDim Preserve ListaCampos(n)
    Set ListaCampos(n - 1) = CAMPO
    
    Set CAMPO = New eDatosAuxiliaresCampos
    CAMPO.Encabezado = "Impuesto/A24"
    CAMPO.Formula = ""
    CAMPO.CampoEditable = "N"
    CAMPO.ColumnaCampoEco = ""
    CAMPO.Total = ""
    CAMPO.DejarLibre = ""
    CAMPO.MarcaBorrado = 0
    n = n + 1
    ReDim Preserve ListaCampos(n)
    Set ListaCampos(n - 1) = CAMPO
        
    Set CAMPO = New eDatosAuxiliaresCampos
    CAMPO.Encabezado = "C.Contable/A24"
    CAMPO.Formula = ""
    CAMPO.CampoEditable = "S"
    CAMPO.ColumnaCampoEco = "codexcue;denocue;plancuen;statcue>0;order by codexcue"
    CAMPO.Total = ""
    CAMPO.DejarLibre = ""
    CAMPO.MarcaBorrado = 0
    n = n + 1
    ReDim Preserve ListaCampos(n)
    Set ListaCampos(n - 1) = CAMPO
    
    Set CAMPO = New eDatosAuxiliaresCampos
    CAMPO.Encabezado = "Denominación/A24"
    CAMPO.Formula = ""
    CAMPO.CampoEditable = "N"
    CAMPO.ColumnaCampoEco = ""
    CAMPO.Total = ""
    CAMPO.DejarLibre = ""
    CAMPO.MarcaBorrado = 0
    n = n + 1
    ReDim Preserve ListaCampos(n)
    Set ListaCampos(n - 1) = CAMPO
    
    Call CrearTabla("COMCLPRIMPU", "Compensación de Impuestos", ListaCampos)
    
    ' TABLA DE TARAS
    ' **************
    Set CAMPO = New eDatosAuxiliaresCampos
    CAMPO.Encabezado = "Ord/I10"
    CAMPO.Formula = ""
    CAMPO.CampoEditable = "N"
    CAMPO.ColumnaCampoEco = ""
    CAMPO.Total = ""
    CAMPO.DejarLibre = ""
    CAMPO.MarcaBorrado = 0
    n = 1
    ReDim Preserve ListaCampos(n)
    Set ListaCampos(n - 1) = CAMPO
    
    Set CAMPO = New eDatosAuxiliaresCampos
    CAMPO.Encabezado = "Descripcion/A32"
    CAMPO.Formula = ""
    CAMPO.CampoEditable = "S"
    CAMPO.ColumnaCampoEco = ""
    CAMPO.Total = ""
    CAMPO.DejarLibre = ""
    CAMPO.MarcaBorrado = 0
    n = n + 1
    ReDim Preserve ListaCampos(n)
    Set ListaCampos(n - 1) = CAMPO
        
    Set CAMPO = New eDatosAuxiliaresCampos
    CAMPO.Encabezado = "Peso/F12.2"
    CAMPO.Formula = ""
    CAMPO.CampoEditable = "S"
    CAMPO.ColumnaCampoEco = ""
    CAMPO.Total = ""
    CAMPO.DejarLibre = ""
    CAMPO.MarcaBorrado = 0
    n = n + 1
    ReDim Preserve ListaCampos(n)
    Set ListaCampos(n - 1) = CAMPO
    
    Call CrearTabla("TABLATARAS", "Tabla de Taras", ListaCampos)

    ' TABLA DE TARAS POR TIPO DE RECEPCION
    ' ************************************
    Set CAMPO = New eDatosAuxiliaresCampos
    CAMPO.Encabezado = "Ord/I10"
    CAMPO.Formula = ""
    CAMPO.CampoEditable = "N"
    CAMPO.ColumnaCampoEco = ""
    CAMPO.Total = ""
    CAMPO.DejarLibre = ""
    CAMPO.MarcaBorrado = 0
    n = 1
    ReDim Preserve ListaCampos(n)
    Set ListaCampos(n - 1) = CAMPO
    
    Set CAMPO = New eDatosAuxiliaresCampos
    CAMPO.Encabezado = "C-R/I5"
    CAMPO.Formula = ""
    CAMPO.CampoEditable = "S"
    CAMPO.ColumnaCampoEco = "codigo;denominacion;#TPTiposRecepcion;codigo>0;order by codigo"
    CAMPO.Total = ""
    CAMPO.DejarLibre = ""
    CAMPO.MarcaBorrado = 0
    n = 1
    ReDim Preserve ListaCampos(n)
    Set ListaCampos(n - 1) = CAMPO
    
    Set CAMPO = New eDatosAuxiliaresCampos
    CAMPO.Encabezado = "Denominacion/A32"
    CAMPO.Formula = ""
    CAMPO.CampoEditable = "N"
    CAMPO.ColumnaCampoEco = ""
    CAMPO.Total = ""
    CAMPO.DejarLibre = ""
    CAMPO.MarcaBorrado = 0
    n = n + 1
    ReDim Preserve ListaCampos(n)
    Set ListaCampos(n - 1) = CAMPO
        
    Set CAMPO = New eDatosAuxiliaresCampos
    CAMPO.Encabezado = "C-T/I5"
    CAMPO.Formula = ""
    CAMPO.CampoEditable = "S"
    IDLIST& = ObtenerIdTablaAuxiliar("TABLATARAS")
    CAMPO.ColumnaCampoEco = "campo1;campo2;datasql;id_lista=" & CStr(IDLIST&) & ";order by cast(campo1 as int)"
    CAMPO.Total = ""
    CAMPO.DejarLibre = ""
    CAMPO.MarcaBorrado = 0
    n = n + 1
    ReDim Preserve ListaCampos(n)
    Set ListaCampos(n - 1) = CAMPO
    
    Set CAMPO = New eDatosAuxiliaresCampos
    CAMPO.Encabezado = "Descripcion/A32"
    CAMPO.Formula = ""
    CAMPO.CampoEditable = "N"
    CAMPO.ColumnaCampoEco = ""
    CAMPO.Total = ""
    CAMPO.DejarLibre = ""
    CAMPO.MarcaBorrado = 0
    n = n + 1
    ReDim Preserve ListaCampos(n)
    Set ListaCampos(n - 1) = CAMPO
    
    Call CrearTabla("TARATIRECE", "Tara por Tipo Recepción", ListaCampos)


    ' TABLA DE ALTAS POR FAMILIA
    ' **************************
    Set CAMPO = New eDatosAuxiliaresCampos
    CAMPO.Encabezado = "Ord/I10"
    CAMPO.Formula = ""
    CAMPO.CampoEditable = "N"
    CAMPO.ColumnaCampoEco = ""
    CAMPO.Total = ""
    CAMPO.DejarLibre = ""
    CAMPO.MarcaBorrado = 0
    n = 1
    ReDim Preserve ListaCampos(n)
    Set ListaCampos(n - 1) = CAMPO
    
    Set CAMPO = New eDatosAuxiliaresCampos
    CAMPO.Encabezado = "Familia/A48"
    CAMPO.Formula = ""
    CAMPO.CampoEditable = "S"
    CAMPO.ColumnaCampoEco = ""
    CAMPO.Total = ""
    CAMPO.DejarLibre = ""
    CAMPO.MarcaBorrado = 0
    n = n + 1
    ReDim Preserve ListaCampos(n)
    Set ListaCampos(n - 1) = CAMPO
    
    Set CAMPO = New eDatosAuxiliaresCampos
    CAMPO.Encabezado = "Cod.Fam./A8"
    CAMPO.Formula = ""
    CAMPO.CampoEditable = "S"
    CAMPO.ColumnaCampoEco = ""
    CAMPO.Total = ""
    CAMPO.DejarLibre = ""
    CAMPO.MarcaBorrado = 0
    n = n + 1
    ReDim Preserve ListaCampos(n)
    Set ListaCampos(n - 1) = CAMPO
   
    For i = 1 To 11
        Set CAMPO = New eDatosAuxiliaresCampos
        CAMPO.Encabezado = "Grupo-" & CStr(i) & "/A16"
        CAMPO.Formula = ""
        CAMPO.CampoEditable = "S"
        CAMPO.ColumnaCampoEco = "lista;descripcion|NO_REPLICA_ECO;encaprfun;marborra=0 or marborra is null"
        CAMPO.Total = ""
        CAMPO.DejarLibre = ""
        CAMPO.MarcaBorrado = 0
        n = n + 1
        ReDim Preserve ListaCampos(n)
        Set ListaCampos(n - 1) = CAMPO
    Next i
        
    Call CrearTabla("TAFORALFA", "Tabla Formato para Alta por Familia", ListaCampos)

    Call CrearTablasAuxiliares_Familias_RETEPLAST
    
    ' TABLA DE ATRIBUTOS COMERCIALES DE PEDIDOS
    ' *****************************************
    Set CAMPO = New eDatosAuxiliaresCampos
    CAMPO.Encabezado = "Ord/I10"
    CAMPO.Formula = ""
    CAMPO.CampoEditable = "N"
    CAMPO.ColumnaCampoEco = ""
    CAMPO.Total = ""
    CAMPO.DejarLibre = ""
    CAMPO.MarcaBorrado = 0
    n = 1
    ReDim Preserve ListaCampos(n)
    Set ListaCampos(n - 1) = CAMPO
    
    Set CAMPO = New eDatosAuxiliaresCampos
    CAMPO.Encabezado = "Denominacion/A48"
    CAMPO.Formula = ""
    CAMPO.CampoEditable = "S"
    CAMPO.ColumnaCampoEco = ""
    CAMPO.Total = ""
    CAMPO.DejarLibre = ""
    CAMPO.MarcaBorrado = 0
    n = n + 1
    ReDim Preserve ListaCampos(n)
    Set ListaCampos(n - 1) = CAMPO
    
    Call CrearTabla("ATRICOPE", "Atributos Comerciales de Pedidos", ListaCampos)

    
End Function

Private Sub CrearTablasAuxiliares_Familias_RETEPLAST()

    Dim ListaCampos() As eDatosAuxiliaresCampos
    Dim CAMPO As eDatosAuxiliaresCampos
    
    ' TIPOS
    ' *****
    Set CAMPO = New eDatosAuxiliaresCampos
    CAMPO.Encabezado = "Ord/I10"
    CAMPO.Formula = ""
    CAMPO.CampoEditable = "N"
    CAMPO.ColumnaCampoEco = ""
    CAMPO.Total = ""
    CAMPO.DejarLibre = ""
    CAMPO.MarcaBorrado = 0
    n = 1
    ReDim Preserve ListaCampos(n)
    Set ListaCampos(n - 1) = CAMPO
    
    Set CAMPO = New eDatosAuxiliaresCampos
    CAMPO.Encabezado = "Codigo/A10"
    CAMPO.Formula = ""
    CAMPO.CampoEditable = "S"
    CAMPO.ColumnaCampoEco = ""
    CAMPO.Total = ""
    CAMPO.DejarLibre = ""
    CAMPO.MarcaBorrado = 0
    n = n + 1
    ReDim Preserve ListaCampos(n)
    Set ListaCampos(n - 1) = CAMPO

    Set CAMPO = New eDatosAuxiliaresCampos
    CAMPO.Encabezado = "Descripción/A48"
    CAMPO.Formula = ""
    CAMPO.CampoEditable = "S"
    CAMPO.ColumnaCampoEco = ""
    CAMPO.Total = ""
    CAMPO.DejarLibre = ""
    CAMPO.MarcaBorrado = 0
    n = n + 1
    ReDim Preserve ListaCampos(n)
    Set ListaCampos(n - 1) = CAMPO

    Call CrearTabla("FAM_TIPO", "TIPOS", ListaCampos)
    
    ' PRODUCTO
    ' ********
    Set CAMPO = New eDatosAuxiliaresCampos
    CAMPO.Encabezado = "Ord/I10"
    CAMPO.Formula = ""
    CAMPO.CampoEditable = "N"
    CAMPO.ColumnaCampoEco = ""
    CAMPO.Total = ""
    CAMPO.DejarLibre = ""
    CAMPO.MarcaBorrado = 0
    n = 1
    ReDim Preserve ListaCampos(n)
    Set ListaCampos(n - 1) = CAMPO
    
    Set CAMPO = New eDatosAuxiliaresCampos
    CAMPO.Encabezado = "Codigo/A10"
    CAMPO.Formula = ""
    CAMPO.CampoEditable = "S"
    CAMPO.ColumnaCampoEco = ""
    CAMPO.Total = ""
    CAMPO.DejarLibre = ""
    CAMPO.MarcaBorrado = 0
    n = n + 1
    ReDim Preserve ListaCampos(n)
    Set ListaCampos(n - 1) = CAMPO

    Set CAMPO = New eDatosAuxiliaresCampos
    CAMPO.Encabezado = "Descripción/A48"
    CAMPO.Formula = ""
    CAMPO.CampoEditable = "S"
    CAMPO.ColumnaCampoEco = ""
    CAMPO.Total = ""
    CAMPO.DejarLibre = ""
    CAMPO.MarcaBorrado = 0
    n = n + 1
    ReDim Preserve ListaCampos(n)
    Set ListaCampos(n - 1) = CAMPO

    Call CrearTabla("FAM_PRODUCTO", "PRODUCTO", ListaCampos)
    
    ' DETALLE
    ' *******
    Set CAMPO = New eDatosAuxiliaresCampos
    CAMPO.Encabezado = "Ord/I10"
    CAMPO.Formula = ""
    CAMPO.CampoEditable = "N"
    CAMPO.ColumnaCampoEco = ""
    CAMPO.Total = ""
    CAMPO.DejarLibre = ""
    CAMPO.MarcaBorrado = 0
    n = 1
    ReDim Preserve ListaCampos(n)
    Set ListaCampos(n - 1) = CAMPO
    
    Set CAMPO = New eDatosAuxiliaresCampos
    CAMPO.Encabezado = "Codigo/A10"
    CAMPO.Formula = ""
    CAMPO.CampoEditable = "S"
    CAMPO.ColumnaCampoEco = ""
    CAMPO.Total = ""
    CAMPO.DejarLibre = ""
    CAMPO.MarcaBorrado = 0
    n = n + 1
    ReDim Preserve ListaCampos(n)
    Set ListaCampos(n - 1) = CAMPO

    Set CAMPO = New eDatosAuxiliaresCampos
    CAMPO.Encabezado = "Descripción/A48"
    CAMPO.Formula = ""
    CAMPO.CampoEditable = "S"
    CAMPO.ColumnaCampoEco = ""
    CAMPO.Total = ""
    CAMPO.DejarLibre = ""
    CAMPO.MarcaBorrado = 0
    n = n + 1
    ReDim Preserve ListaCampos(n)
    Set ListaCampos(n - 1) = CAMPO

    Call CrearTabla("FAM_DETA", "DETALLE", ListaCampos)
    
    ' COLOR
    ' *****
    Set CAMPO = New eDatosAuxiliaresCampos
    CAMPO.Encabezado = "Ord/I10"
    CAMPO.Formula = ""
    CAMPO.CampoEditable = "N"
    CAMPO.ColumnaCampoEco = ""
    CAMPO.Total = ""
    CAMPO.DejarLibre = ""
    CAMPO.MarcaBorrado = 0
    n = 1
    ReDim Preserve ListaCampos(n)
    Set ListaCampos(n - 1) = CAMPO
    
    Set CAMPO = New eDatosAuxiliaresCampos
    CAMPO.Encabezado = "Codigo/A10"
    CAMPO.Formula = ""
    CAMPO.CampoEditable = "S"
    CAMPO.ColumnaCampoEco = ""
    CAMPO.Total = ""
    CAMPO.DejarLibre = ""
    CAMPO.MarcaBorrado = 0
    n = n + 1
    ReDim Preserve ListaCampos(n)
    Set ListaCampos(n - 1) = CAMPO

    Set CAMPO = New eDatosAuxiliaresCampos
    CAMPO.Encabezado = "Descripción/A48"
    CAMPO.Formula = ""
    CAMPO.CampoEditable = "S"
    CAMPO.ColumnaCampoEco = ""
    CAMPO.Total = ""
    CAMPO.DejarLibre = ""
    CAMPO.MarcaBorrado = 0
    n = n + 1
    ReDim Preserve ListaCampos(n)
    Set ListaCampos(n - 1) = CAMPO

    Call CrearTabla("FAM_COLOR", "COLOR", ListaCampos)
    
    ' ESTADO
    ' ******
    Set CAMPO = New eDatosAuxiliaresCampos
    CAMPO.Encabezado = "Ord/I10"
    CAMPO.Formula = ""
    CAMPO.CampoEditable = "N"
    CAMPO.ColumnaCampoEco = ""
    CAMPO.Total = ""
    CAMPO.DejarLibre = ""
    CAMPO.MarcaBorrado = 0
    n = 1
    ReDim Preserve ListaCampos(n)
    Set ListaCampos(n - 1) = CAMPO
    
    Set CAMPO = New eDatosAuxiliaresCampos
    CAMPO.Encabezado = "Codigo/A10"
    CAMPO.Formula = ""
    CAMPO.CampoEditable = "S"
    CAMPO.ColumnaCampoEco = ""
    CAMPO.Total = ""
    CAMPO.DejarLibre = ""
    CAMPO.MarcaBorrado = 0
    n = n + 1
    ReDim Preserve ListaCampos(n)
    Set ListaCampos(n - 1) = CAMPO

    Set CAMPO = New eDatosAuxiliaresCampos
    CAMPO.Encabezado = "Descripción/A48"
    CAMPO.Formula = ""
    CAMPO.CampoEditable = "S"
    CAMPO.ColumnaCampoEco = ""
    CAMPO.Total = ""
    CAMPO.DejarLibre = ""
    CAMPO.MarcaBorrado = 0
    n = n + 1
    ReDim Preserve ListaCampos(n)
    Set ListaCampos(n - 1) = CAMPO

    Call CrearTabla("FAM_ESTADO", "ESTADO", ListaCampos)
    
    ' PROCEDENCIA
    ' ***********
    Set CAMPO = New eDatosAuxiliaresCampos
    CAMPO.Encabezado = "Ord/I10"
    CAMPO.Formula = ""
    CAMPO.CampoEditable = "N"
    CAMPO.ColumnaCampoEco = ""
    CAMPO.Total = ""
    CAMPO.DejarLibre = ""
    CAMPO.MarcaBorrado = 0
    n = 1
    ReDim Preserve ListaCampos(n)
    Set ListaCampos(n - 1) = CAMPO
    
    Set CAMPO = New eDatosAuxiliaresCampos
    CAMPO.Encabezado = "Codigo/A10"
    CAMPO.Formula = ""
    CAMPO.CampoEditable = "S"
    CAMPO.ColumnaCampoEco = ""
    CAMPO.Total = ""
    CAMPO.DejarLibre = ""
    CAMPO.MarcaBorrado = 0
    n = n + 1
    ReDim Preserve ListaCampos(n)
    Set ListaCampos(n - 1) = CAMPO

    Set CAMPO = New eDatosAuxiliaresCampos
    CAMPO.Encabezado = "Descripción/A48"
    CAMPO.Formula = ""
    CAMPO.CampoEditable = "S"
    CAMPO.ColumnaCampoEco = ""
    CAMPO.Total = ""
    CAMPO.DejarLibre = ""
    CAMPO.MarcaBorrado = 0
    n = n + 1
    ReDim Preserve ListaCampos(n)
    Set ListaCampos(n - 1) = CAMPO

    Call CrearTabla("FAM_PROCE", "PROCEDENCIA", ListaCampos)
    
    ' CONTAMINACION
    ' *************
    Set CAMPO = New eDatosAuxiliaresCampos
    CAMPO.Encabezado = "Ord/I10"
    CAMPO.Formula = ""
    CAMPO.CampoEditable = "N"
    CAMPO.ColumnaCampoEco = ""
    CAMPO.Total = ""
    CAMPO.DejarLibre = ""
    CAMPO.MarcaBorrado = 0
    n = 1
    ReDim Preserve ListaCampos(n)
    Set ListaCampos(n - 1) = CAMPO
    
    Set CAMPO = New eDatosAuxiliaresCampos
    CAMPO.Encabezado = "Codigo/A10"
    CAMPO.Formula = ""
    CAMPO.CampoEditable = "S"
    CAMPO.ColumnaCampoEco = ""
    CAMPO.Total = ""
    CAMPO.DejarLibre = ""
    CAMPO.MarcaBorrado = 0
    n = n + 1
    ReDim Preserve ListaCampos(n)
    Set ListaCampos(n - 1) = CAMPO

    Set CAMPO = New eDatosAuxiliaresCampos
    CAMPO.Encabezado = "Descripción/A48"
    CAMPO.Formula = ""
    CAMPO.CampoEditable = "S"
    CAMPO.ColumnaCampoEco = ""
    CAMPO.Total = ""
    CAMPO.DejarLibre = ""
    CAMPO.MarcaBorrado = 0
    n = n + 1
    ReDim Preserve ListaCampos(n)
    Set ListaCampos(n - 1) = CAMPO

    Call CrearTabla("FAM_CONTAM", "CONTAMINACION", ListaCampos)
    
    ' MATERIAL
    ' ********
    Set CAMPO = New eDatosAuxiliaresCampos
    CAMPO.Encabezado = "Ord/I10"
    CAMPO.Formula = ""
    CAMPO.CampoEditable = "N"
    CAMPO.ColumnaCampoEco = ""
    CAMPO.Total = ""
    CAMPO.DejarLibre = ""
    CAMPO.MarcaBorrado = 0
    n = 1
    ReDim Preserve ListaCampos(n)
    Set ListaCampos(n - 1) = CAMPO
    
    Set CAMPO = New eDatosAuxiliaresCampos
    CAMPO.Encabezado = "Codigo/A10"
    CAMPO.Formula = ""
    CAMPO.CampoEditable = "S"
    CAMPO.ColumnaCampoEco = ""
    CAMPO.Total = ""
    CAMPO.DejarLibre = ""
    CAMPO.MarcaBorrado = 0
    n = n + 1
    ReDim Preserve ListaCampos(n)
    Set ListaCampos(n - 1) = CAMPO

    Set CAMPO = New eDatosAuxiliaresCampos
    CAMPO.Encabezado = "Descripción/A48"
    CAMPO.Formula = ""
    CAMPO.CampoEditable = "S"
    CAMPO.ColumnaCampoEco = ""
    CAMPO.Total = ""
    CAMPO.DejarLibre = ""
    CAMPO.MarcaBorrado = 0
    n = n + 1
    ReDim Preserve ListaCampos(n)
    Set ListaCampos(n - 1) = CAMPO

    Call CrearTabla("FAM_MATERIAL", "MATERIAL", ListaCampos)
    
    ' PROCESO
    ' *******
    Set CAMPO = New eDatosAuxiliaresCampos
    CAMPO.Encabezado = "Ord/I10"
    CAMPO.Formula = ""
    CAMPO.CampoEditable = "N"
    CAMPO.ColumnaCampoEco = ""
    CAMPO.Total = ""
    CAMPO.DejarLibre = ""
    CAMPO.MarcaBorrado = 0
    n = 1
    ReDim Preserve ListaCampos(n)
    Set ListaCampos(n - 1) = CAMPO
    
    Set CAMPO = New eDatosAuxiliaresCampos
    CAMPO.Encabezado = "Codigo/A10"
    CAMPO.Formula = ""
    CAMPO.CampoEditable = "S"
    CAMPO.ColumnaCampoEco = ""
    CAMPO.Total = ""
    CAMPO.DejarLibre = ""
    CAMPO.MarcaBorrado = 0
    n = n + 1
    ReDim Preserve ListaCampos(n)
    Set ListaCampos(n - 1) = CAMPO

    Set CAMPO = New eDatosAuxiliaresCampos
    CAMPO.Encabezado = "Descripción/A48"
    CAMPO.Formula = ""
    CAMPO.CampoEditable = "S"
    CAMPO.ColumnaCampoEco = ""
    CAMPO.Total = ""
    CAMPO.DejarLibre = ""
    CAMPO.MarcaBorrado = 0
    n = n + 1
    ReDim Preserve ListaCampos(n)
    Set ListaCampos(n - 1) = CAMPO

    Call CrearTabla("FAM_PROCESO", "PROCESO", ListaCampos)
    
    ' FLUIDEZ
    ' *******
    Set CAMPO = New eDatosAuxiliaresCampos
    CAMPO.Encabezado = "Ord/I10"
    CAMPO.Formula = ""
    CAMPO.CampoEditable = "N"
    CAMPO.ColumnaCampoEco = ""
    CAMPO.Total = ""
    CAMPO.DejarLibre = ""
    CAMPO.MarcaBorrado = 0
    n = 1
    ReDim Preserve ListaCampos(n)
    Set ListaCampos(n - 1) = CAMPO
    
    Set CAMPO = New eDatosAuxiliaresCampos
    CAMPO.Encabezado = "Codigo/A10"
    CAMPO.Formula = ""
    CAMPO.CampoEditable = "S"
    CAMPO.ColumnaCampoEco = ""
    CAMPO.Total = ""
    CAMPO.DejarLibre = ""
    CAMPO.MarcaBorrado = 0
    n = n + 1
    ReDim Preserve ListaCampos(n)
    Set ListaCampos(n - 1) = CAMPO

    Set CAMPO = New eDatosAuxiliaresCampos
    CAMPO.Encabezado = "Descripción/A48"
    CAMPO.Formula = ""
    CAMPO.CampoEditable = "S"
    CAMPO.ColumnaCampoEco = ""
    CAMPO.Total = ""
    CAMPO.DejarLibre = ""
    CAMPO.MarcaBorrado = 0
    n = n + 1
    ReDim Preserve ListaCampos(n)
    Set ListaCampos(n - 1) = CAMPO

    Call CrearTabla("FAM_FLUIDEZ", "FLUIDEZ", ListaCampos)
    
    ' CARGA
    ' *****
    Set CAMPO = New eDatosAuxiliaresCampos
    CAMPO.Encabezado = "Ord/I10"
    CAMPO.Formula = ""
    CAMPO.CampoEditable = "N"
    CAMPO.ColumnaCampoEco = ""
    CAMPO.Total = ""
    CAMPO.DejarLibre = ""
    CAMPO.MarcaBorrado = 0
    n = 1
    ReDim Preserve ListaCampos(n)
    Set ListaCampos(n - 1) = CAMPO
    
    Set CAMPO = New eDatosAuxiliaresCampos
    CAMPO.Encabezado = "Codigo/A10"
    CAMPO.Formula = ""
    CAMPO.CampoEditable = "S"
    CAMPO.ColumnaCampoEco = ""
    CAMPO.Total = ""
    CAMPO.DejarLibre = ""
    CAMPO.MarcaBorrado = 0
    n = n + 1
    ReDim Preserve ListaCampos(n)
    Set ListaCampos(n - 1) = CAMPO

    Set CAMPO = New eDatosAuxiliaresCampos
    CAMPO.Encabezado = "Descripción/A48"
    CAMPO.Formula = ""
    CAMPO.CampoEditable = "S"
    CAMPO.ColumnaCampoEco = ""
    CAMPO.Total = ""
    CAMPO.DejarLibre = ""
    CAMPO.MarcaBorrado = 0
    n = n + 1
    ReDim Preserve ListaCampos(n)
    Set ListaCampos(n - 1) = CAMPO

    Call CrearTabla("FAM_CARGA", "CARGA", ListaCampos)
    
    ' CARGATIPO
    ' *********
    Set CAMPO = New eDatosAuxiliaresCampos
    CAMPO.Encabezado = "Ord/I10"
    CAMPO.Formula = ""
    CAMPO.CampoEditable = "N"
    CAMPO.ColumnaCampoEco = ""
    CAMPO.Total = ""
    CAMPO.DejarLibre = ""
    CAMPO.MarcaBorrado = 0
    n = 1
    ReDim Preserve ListaCampos(n)
    Set ListaCampos(n - 1) = CAMPO
    
    Set CAMPO = New eDatosAuxiliaresCampos
    CAMPO.Encabezado = "Codigo/A10"
    CAMPO.Formula = ""
    CAMPO.CampoEditable = "S"
    CAMPO.ColumnaCampoEco = ""
    CAMPO.Total = ""
    CAMPO.DejarLibre = ""
    CAMPO.MarcaBorrado = 0
    n = n + 1
    ReDim Preserve ListaCampos(n)
    Set ListaCampos(n - 1) = CAMPO

    Set CAMPO = New eDatosAuxiliaresCampos
    CAMPO.Encabezado = "Descripción/A48"
    CAMPO.Formula = ""
    CAMPO.CampoEditable = "S"
    CAMPO.ColumnaCampoEco = ""
    CAMPO.Total = ""
    CAMPO.DejarLibre = ""
    CAMPO.MarcaBorrado = 0
    n = n + 1
    ReDim Preserve ListaCampos(n)
    Set ListaCampos(n - 1) = CAMPO

    Call CrearTabla("FAM_CARGATIPO", "CARGATIPO", ListaCampos)
    
    ' ADIT-PROP
    ' *********
    Set CAMPO = New eDatosAuxiliaresCampos
    CAMPO.Encabezado = "Ord/I10"
    CAMPO.Formula = ""
    CAMPO.CampoEditable = "N"
    CAMPO.ColumnaCampoEco = ""
    CAMPO.Total = ""
    CAMPO.DejarLibre = ""
    CAMPO.MarcaBorrado = 0
    n = 1
    ReDim Preserve ListaCampos(n)
    Set ListaCampos(n - 1) = CAMPO
    
    Set CAMPO = New eDatosAuxiliaresCampos
    CAMPO.Encabezado = "Codigo/A10"
    CAMPO.Formula = ""
    CAMPO.CampoEditable = "S"
    CAMPO.ColumnaCampoEco = ""
    CAMPO.Total = ""
    CAMPO.DejarLibre = ""
    CAMPO.MarcaBorrado = 0
    n = n + 1
    ReDim Preserve ListaCampos(n)
    Set ListaCampos(n - 1) = CAMPO

    Set CAMPO = New eDatosAuxiliaresCampos
    CAMPO.Encabezado = "Descripción/A48"
    CAMPO.Formula = ""
    CAMPO.CampoEditable = "S"
    CAMPO.ColumnaCampoEco = ""
    CAMPO.Total = ""
    CAMPO.DejarLibre = ""
    CAMPO.MarcaBorrado = 0
    n = n + 1
    ReDim Preserve ListaCampos(n)
    Set ListaCampos(n - 1) = CAMPO

    Call CrearTabla("FAM_ADIT", "ADIT-PROP", ListaCampos)
    
    ' CATEGORIA
    ' *********
    Set CAMPO = New eDatosAuxiliaresCampos
    CAMPO.Encabezado = "Ord/I10"
    CAMPO.Formula = ""
    CAMPO.CampoEditable = "N"
    CAMPO.ColumnaCampoEco = ""
    CAMPO.Total = ""
    CAMPO.DejarLibre = ""
    CAMPO.MarcaBorrado = 0
    n = 1
    ReDim Preserve ListaCampos(n)
    Set ListaCampos(n - 1) = CAMPO
    
    Set CAMPO = New eDatosAuxiliaresCampos
    CAMPO.Encabezado = "Codigo/A10"
    CAMPO.Formula = ""
    CAMPO.CampoEditable = "S"
    CAMPO.ColumnaCampoEco = ""
    CAMPO.Total = ""
    CAMPO.DejarLibre = ""
    CAMPO.MarcaBorrado = 0
    n = n + 1
    ReDim Preserve ListaCampos(n)
    Set ListaCampos(n - 1) = CAMPO

    Set CAMPO = New eDatosAuxiliaresCampos
    CAMPO.Encabezado = "Descripción/A48"
    CAMPO.Formula = ""
    CAMPO.CampoEditable = "S"
    CAMPO.ColumnaCampoEco = ""
    CAMPO.Total = ""
    CAMPO.DejarLibre = ""
    CAMPO.MarcaBorrado = 0
    n = n + 1
    ReDim Preserve ListaCampos(n)
    Set ListaCampos(n - 1) = CAMPO

    Call CrearTabla("FAM_CATEGORIA", "CATEGORIA", ListaCampos)
    
    ' COLORTONO
    ' *********
    Set CAMPO = New eDatosAuxiliaresCampos
    CAMPO.Encabezado = "Ord/I10"
    CAMPO.Formula = ""
    CAMPO.CampoEditable = "N"
    CAMPO.ColumnaCampoEco = ""
    CAMPO.Total = ""
    CAMPO.DejarLibre = ""
    CAMPO.MarcaBorrado = 0
    n = 1
    ReDim Preserve ListaCampos(n)
    Set ListaCampos(n - 1) = CAMPO
    
    Set CAMPO = New eDatosAuxiliaresCampos
    CAMPO.Encabezado = "Codigo/A10"
    CAMPO.Formula = ""
    CAMPO.CampoEditable = "S"
    CAMPO.ColumnaCampoEco = ""
    CAMPO.Total = ""
    CAMPO.DejarLibre = ""
    CAMPO.MarcaBorrado = 0
    n = n + 1
    ReDim Preserve ListaCampos(n)
    Set ListaCampos(n - 1) = CAMPO

    Set CAMPO = New eDatosAuxiliaresCampos
    CAMPO.Encabezado = "Descripción/A48"
    CAMPO.Formula = ""
    CAMPO.CampoEditable = "S"
    CAMPO.ColumnaCampoEco = ""
    CAMPO.Total = ""
    CAMPO.DejarLibre = ""
    CAMPO.MarcaBorrado = 0
    n = n + 1
    ReDim Preserve ListaCampos(n)
    Set ListaCampos(n - 1) = CAMPO

    Call CrearTabla("FAM_COLORTONO", "COLOR TONO", ListaCampos)
    
    ' TABLA DE FAMILIA DE COMPONENTES
    ' *******************************
    Set CAMPO = New eDatosAuxiliaresCampos
    CAMPO.Encabezado = "Codigo Fam/A16"
    CAMPO.Formula = ""
    CAMPO.CampoEditable = "S"
    CAMPO.ColumnaCampoEco = "codigo;descripcion;master;status>0;order by codigo"
    CAMPO.Total = ""
    CAMPO.DejarLibre = ""
    CAMPO.MarcaBorrado = 0
    n = 1
    ReDim Preserve ListaCampos(n)
    Set ListaCampos(n - 1) = CAMPO
    
    Set CAMPO = New eDatosAuxiliaresCampos
    CAMPO.Encabezado = "Familia/A48"
    CAMPO.Formula = ""
    CAMPO.CampoEditable = "N"
    CAMPO.ColumnaCampoEco = ""
    CAMPO.Total = ""
    CAMPO.DejarLibre = ""
    CAMPO.MarcaBorrado = 0
    n = n + 1
    ReDim Preserve ListaCampos(n)
    Set ListaCampos(n - 1) = CAMPO

    Set CAMPO = New eDatosAuxiliaresCampos
    CAMPO.Encabezado = "Codigo/A16"
    CAMPO.Formula = ""
    CAMPO.CampoEditable = "S"
    CAMPO.ColumnaCampoEco = "codigo;descripcion;master;status>0;order by codigo"
    CAMPO.Total = ""
    CAMPO.DejarLibre = ""
    CAMPO.MarcaBorrado = 0
    n = n + 1
    ReDim Preserve ListaCampos(n)
    Set ListaCampos(n - 1) = CAMPO
    
    Set CAMPO = New eDatosAuxiliaresCampos
    CAMPO.Encabezado = "Descripcion/A48"
    CAMPO.Formula = ""
    CAMPO.CampoEditable = "N"
    CAMPO.ColumnaCampoEco = ""
    CAMPO.Total = ""
    CAMPO.DejarLibre = ""
    CAMPO.MarcaBorrado = 0
    n = n + 1
    ReDim Preserve ListaCampos(n)
    Set ListaCampos(n - 1) = CAMPO
    
    Call CrearTabla("FAM_FAMCOM", "FAM. DE COMPONENTES", ListaCampos)
    
End Sub


