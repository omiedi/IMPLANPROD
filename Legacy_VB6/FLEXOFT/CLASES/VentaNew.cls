VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "VentaNew"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
    Private Type Formulax
        Formula As String
        ColEcoCampo As String
        Condicion1 As String
        Condicion2 As String
        TxtUbicadoGrillaTotal As String
        TOTALES As String
        DejarLibre As String
        Enumerar As String
        EnumerarRepUltimo As String
        RepiteDato As String
        VentVectorToolTipText As String
    End Type
    'Declaramos el vector
    
    Private FormulaGrid(32) As Formulax
    Private ColEcoCampoGrid(32) As Formulax
    Private Condicion1Grid(32) As Formulax
    Private Condicion2Grid(32) As Formulax

    Private TxtUbicadoGrillaTotales(128) As Formulax
    Private TOTALESGrid(128) As Formulax
    Private DejarLibreGrid(128) As Formulax
    Private EnumerarGrid(128) As Formulax
    Private EnumerarRepUltimoGrid(128) As Formulax
    Private RepiteDatoGrid(128) As Formulax
    Private VentVectorToolTipTextGrid(64) As Formulax
    Private Type VENTANA
    
        Campox As String
        UtilizaFormula As Integer
        FormulaAyuda As String
        Inicializar As Integer
        CampEditables As String
        AgregaRegistro As String
        TITULO As String
        HabilitaBorrar As Integer
        HabilitaInsertar As Integer
        NoMostrarBotonTilde As Integer
        PermitirFilasVacias As Integer
        ValidarDuplicadoFila1 As String
    End Type
    Private VentaNew  As VENTANA
    
    Private Type Dato
        Columna1 As String
        Columna2 As Double
        Columna3 As Date
        Columna4 As String
        Columna5 As Double
        Columna6 As Date
        Columna7 As String
        Columna8 As Double
        Columna9 As Date
        Columna10 As String
        Columna11 As Double
        Columna12 As Date
        Columna13 As String
        Columna14 As Double
        Columna15 As Date
        Columna16 As String
        Columna17 As Double
        Columna18 As Date
    End Type
 Private Dato  As VENTANA
    
    Public Property Get Campox() As String
       Campox = VentaNew.Campox
    End Property
    Public Property Let Campox(valor As String)
       VentaNew.Campox = valor
    End Property
    '
    Public Property Get TITULO() As String
       TITULO = VentaNew.TITULO
    End Property
    Public Property Let TITULO(valor As String)
       VentaNew.TITULO = valor
    End Property
    '
    
    Public Property Get AgregaRegistro() As String
       AgregaRegistro = VentaNew.AgregaRegistro
    End Property
    Public Property Let AgregaRegistro(valor As String)
       VentaNew.AgregaRegistro = valor
    End Property
    Public Property Get ValidarDuplicadoFila1() As String
       ValidarDuplicadoFila1 = VentaNew.ValidarDuplicadoFila1
    End Property
    Public Property Let ValidarDuplicadoFila1(valor As String)
       VentaNew.ValidarDuplicadoFila1 = valor
    End Property
    '
    Public Property Get UtilizaFormula() As Integer
       UtilizaFormula = VentaNew.UtilizaFormula
    End Property
    Public Property Let UtilizaFormula(valor As Integer)
       VentaNew.UtilizaFormula = valor
    End Property
    
    Public Property Get Inicializar() As Integer
       Inicializar = VentaNew.Inicializar
    End Property
    Public Property Let Inicializar(valor As Integer)
       VentaNew.Inicializar = valor
    End Property
    '
    Public Property Get CampEditables() As String
       CampEditables = VentaNew.CampEditables
    End Property
    Public Property Let CampEditables(valor As String)
       VentaNew.CampEditables = valor
    End Property
    '
    Public Property Get Formula(i As Integer) As String
       Formula = FormulaGrid(i).Formula 'vector
       ' = "Indice=Columna(donde se muestra el resultado)| 4;*;6 texto de columnas y operadores de la formula separado por ';'"
    End Property
    Public Property Let Formula(i As Integer, ByVal VALOR1 As String)
       FormulaGrid(i).Formula = VALOR1  'vector
    End Property
    '
    Public Property Get ColEcoCampo(i As Integer) As String
       ColEcoCampo = ColEcoCampoGrid(i).ColEcoCampo 'vector
       ' = "Indice=Columna(donde se muestra el resultado)| 4;*;6 texto de columnas y operadores de la formula separado por ';'"
    End Property
    Public Property Let ColEcoCampo(i As Integer, ByVal VALOR1 As String)
       ColEcoCampoGrid(i).ColEcoCampo = VALOR1  'vector
    End Property
    '
    Public Property Get TxtUbicadoGrillaTotal(i As Integer) As String
       TxtUbicadoGrillaTotal = TxtUbicadoGrillaTotales(i).TxtUbicadoGrillaTotal 'vector
       ' = "Indice=Columna(donde se muestra el resultado)| 4;*;6 texto de columnas y operadores de la formula separado por ';'"
    End Property
    Public Property Let TxtUbicadoGrillaTotal(i As Integer, ByVal VALOR1 As String)
       TxtUbicadoGrillaTotales(i).TxtUbicadoGrillaTotal = VALOR1  'vector
    End Property
    
     Public Property Get HabilitaBorrar() As Integer
       HabilitaBorrar = VentaNew.HabilitaBorrar
    End Property
    Public Property Let HabilitaBorrar(valor As Integer)
       '0 NO MUESTRA EL BOTON
       '1 DESHABILITA BOTON BORRAR
       VentaNew.HabilitaBorrar = valor
    End Property
     Public Property Get HabilitaInsertar() As Integer
       HabilitaInsertar = VentaNew.HabilitaInsertar
    End Property
    Public Property Let HabilitaInsertar(valor As Integer)
       '0 NO MUESTRA EL BOTON
       '1 DESHABILITA BOTON BORRAR
       VentaNew.HabilitaInsertar = valor
    End Property
     Public Property Get NoMostrarBotonTilde() As Integer
       NoMostrarBotonTilde = VentaNew.NoMostrarBotonTilde
    End Property
    Public Property Let NoMostrarBotonTilde(valor As Integer)
       '0 VISIBLE
       '1 NO VISIBLE
       VentaNew.NoMostrarBotonTilde = valor
    End Property
     Public Property Get PermitirFilasVacias() As Integer
       PermitirFilasVacias = VentaNew.PermitirFilasVacias
    End Property
    Public Property Let PermitirFilasVacias(valor As Integer)
       '0 las FILAS VACIAS SE BORRAN
       '1 LAS FILAS VACIAS NO SE BORRAN
       VentaNew.PermitirFilasVacias = valor
    End Property
    
    Public Property Get TOTALES(i As Integer) As String
       TOTALES = TOTALESGrid(i).TOTALES 'vector
       ' = "Indice=Columna(donde se muestra el resultado)| 4;*;6 texto de columnas y operadores de la formula separado por ';'"
    End Property
    Public Property Let TOTALES(i As Integer, ByVal VALOR1 As String)
       TOTALESGrid(i).TOTALES = VALOR1    'vector
    End Property
    Public Property Get DejarLibre(i As Integer) As String
       DejarLibre = DejarLibreGrid(i).DejarLibre 'vector
       ' = "Indice=Columna(donde se muestra el resultado)| 4;*;6 texto de columnas y operadores de la formula separado por ';'"
    End Property
    Public Property Let DejarLibre(i As Integer, ByVal VALOR1 As String)
       DejarLibreGrid(i).DejarLibre = VALOR1  'vector
    End Property
    
Private Sub Class_Initialize()
'    Dim FormulaGrid(32) As Formulax
'    Dim ColEcoCampoGrid(32) As Formulax
'    Dim TxtUbicadoGrillaTotales(32) As Formulax
End Sub
    '
    Public Property Get Condicion1(i As Integer) As String
       Condicion1 = Condicion1Grid(i).Condicion1 'vector
       ' = "Indice=Columna(donde se muestra el resultado)| 4;*;6 texto de columnas y operadores de la formula separado por ';'"
    End Property
    Public Property Let Condicion1(i As Integer, ByVal VALOR1 As String)
       Condicion1Grid(i).Condicion1 = VALOR1  'vector
    End Property
    Public Property Get Condicion2(i As Integer) As String
       Condicion2 = Condicion2Grid(i).Condicion2 'vector
       ' = "Indice=Columna(donde se muestra el resultado)| 4;*;6 texto de columnas y operadores de la formula separado por ';'"
    End Property
    Public Property Let Condicion2(i As Integer, ByVal VALOR1 As String)
       Condicion2Grid(i).Condicion2 = VALOR1  'vector
    End Property
    Public Property Get Enumerar(i As Integer) As String
       Enumerar = EnumerarGrid(i).Enumerar 'vector
    End Property
    Public Property Let Enumerar(i As Integer, ByVal VALOR1 As String)
       EnumerarGrid(i).Enumerar = VALOR1  'vector
    End Property
    Public Property Get RepiteDato(i As Integer) As String
       RepiteDato = RepiteDatoGrid(i).RepiteDato 'vector
    End Property
    Public Property Let RepiteDato(i As Integer, ByVal VALOR1 As String)
       RepiteDatoGrid(i).RepiteDato = VALOR1  'vector
    End Property
    Public Property Get EnumerarRepUltimo(i As Integer) As String
       EnumerarRepUltimo = EnumerarRepUltimoGrid(i).EnumerarRepUltimo 'vector
    End Property
    Public Property Let EnumerarRepUltimo(i As Integer, ByVal VALOR1 As String)
       EnumerarRepUltimoGrid(i).EnumerarRepUltimo = VALOR1  'vector
    End Property
    Public Property Get VentVectorToolTipText(i As Integer) As String
       VentVectorToolTipText = VentVectorToolTipTextGrid(i).VentVectorToolTipText 'vector
  'SE LE DEBERA INDICAR EN LA LLAMADA EL TOOLTIPTEXT CORRESPONDIENTE A CADA COLULMNA,
  'CUANDO EL CURSOR PASA POR ENCIMA DE LOS TITULOS DE COLUMNA PRESENTARA EL MISMO

    End Property

    Public Property Let VentVectorToolTipText(i As Integer, ByVal VALOR1 As String)
  'SE LE DEBERA INDICAR EN LA LLAMADA EL TOOLTIPTEXT CORRESPONDIENTE A CADA COLULMNA,
  'CUANDO EL CURSOR PASA POR ENCIMA DE LOS TITULOS DE COLUMNA PRESENTARA EL MISMO
       VentVectorToolTipTextGrid(i).VentVectorToolTipText = VALOR1  'vector
    End Property




 Private Sub ProcessFlexGrid(flexGrid As msFlexGrid, ByRef dataCollection() As Dato, LISTADECAMPOS$)
    Dim rowIndex As Integer
    Dim colIndex As Integer
    Dim currentColumn As Integer
    ''
    Dim NOCAMPO$(), FORMX$(), TipoDato%()
    ReDim NOCAMPO$(1), FORMX$(1), TipoDato%(1)

    While LISTADECAMPOS$ <> ""
       PPXX% = InStr(LSTX$, ";")
       PPXX2% = InStr(LSTX$, ",")
       If PPXX2% > 0 And PPXX2% < PPXX% Then PPXX% = PPXX2%
       If PPXX% < 1 Then PPXX% = 1 + Len(LSTX$)
       '
       CANCA% = CANCA% + 1
       If CANCA% >= UBound(NOCAMPO$) Then
          ReDim Preserve NOCAMPO$(CANCA%), FORMX$(CANCA%), TipoDato%(CANCA%)
       End If
       '
       NOCX$ = Left$(LSTX$, PPXX% - 1)
       LSTX$ = Mid$(LSTX$, PPXX% + 1)
       '
       PPXX% = InStr(NOCX$, "/")
       If PPXX% < 1 Then PPXX% = 1 + Len(NOCX$)
       NOCAMPO$(CANCA%) = Left$(NOCX$, PPXX% - 1)
       FORMX$(CANCA%) = Mid$(NOCX$, PPXX% + 1)
       If FORMX$(CANCA%) = "" Then FORMX$(CANCA%) = "A16"
       
       For i% = 0 To CANCA%
          LETR$ = UCase$(Left$(FORMX$(i%), 1))
          If LETR$ = "A" Then TipoDato%(CANCA%) = 1
          If LETR$ = "I" Or LETR$ = "G" Or LETR$ = "F" Then TipoDato%(CANCA%) = 2
          If LETR$ = "D" Then TipoDato%(CANCA%) = 3
       Next i%
    
    Wend
    ''
    For rowIndex = 1 To flexGrid.Rows - 1 ' Recorrer filas del MSFlexGrid
        currentColumn = 1 ' Reiniciar el índice de columna en cada fila
        
        ' Leer los datos de cada columna de la fila actual del MSFlexGrid
        For colIndex = 0 To flexGrid.Cols - 1
            Dim cellData As Variant
            cellData = flexGrid.TextMatrix(rowIndex, colIndex)
            
            ' Validar y guardar los datos en la colección según el tipo de dato
            Select Case colIndex
                Case 1 ' Columna1 (String)
                    dataCollection(rowIndex).Columna1 = CStr(cellData)
                Case 2 ' Columna2 (Double)
                    dataCollection(rowIndex).Columna2 = CDbl(cellData)
                Case 3 ' Columna3 (Date)
                    If IsDate(cellData) Then
                        dataCollection(rowIndex).Columna3 = CDate(cellData)
                    End If
                Case 4 ' Columna4 (String)
                    dataCollection(rowIndex).Columna4 = CStr(cellData)
                Case 5 ' Columna5 (Double)
                    dataCollection(rowIndex).Columna5 = CDbl(cellData)
                Case 6 ' Columna6 (Date)
                    If IsDate(cellData) Then
                        dataCollection(rowIndex).Columna6 = CDate(cellData)
                    End If
                Case 7 ' Columna7 (String)
                    dataCollection(rowIndex).Columna7 = CStr(cellData)
                Case 8 ' Columna8 (Double)
                    dataCollection(rowIndex).Columna8 = CDbl(cellData)
                Case 9 ' Columna9 (Date)
                    If IsDate(cellData) Then
                        dataCollection(rowIndex).Columna9 = CDate(cellData)
                    End If
                Case 10 ' Columna10 (String)
                    dataCollection(rowIndex).Columna10 = CStr(cellData)
                Case 11 ' Columna11 (Double)
                    dataCollection(rowIndex).Columna11 = CDbl(cellData)
                Case 12 ' Columna12 (Date)
                    If IsDate(cellData) Then
                        dataCollection(rowIndex).Columna12 = CDate(cellData)
                    End If
                Case 13 ' Columna13 (String)
                    dataCollection(rowIndex).Columna13 = CStr(cellData)
                Case 14 ' Columna14 (Double)
                    dataCollection(rowIndex).Columna14 = CDbl(cellData)
                Case 15 ' Columna15 (Date)
                    If IsDate(cellData) Then
                        dataCollection(rowIndex).Columna15 = CDate(cellData)
                    End If
                Case 16 ' Columna16 (String)
                    dataCollection(rowIndex).Columna16 = CStr(cellData)
                Case 17 ' Columna17 (Double)
                    dataCollection(rowIndex).Columna17 = CDbl(cellData)
                Case 18 ' Columna18 (Date)
                    If IsDate(cellData) Then
                        dataCollection(rowIndex).Columna18 = CDate(cellData)
                    End If
            End Select
            
            currentColumn = currentColumn + 1 ' Avanzar a la siguiente columna en la colección
        Next colIndex
    Next rowIndex
End Sub

Public Sub TestProcessFlexGrid()
    Dim flexGrid As msFlexGrid
    Set flexGrid = YourFlexGrid ' Reemplaza "YourFlexGrid" con el nombre de tu MSFlexGrid
    
    Dim dataCollection() As Dato
    ReDim dataCollection(1 To flexGrid.Rows - 1) ' Redimensionar la colección según el número de filas
    
    ProcessFlexGrid flexGrid, dataCollection, LISTADECAMPOS$
    
    ' Mostrar los datos guardados en la colección (solo como ejemplo)
    Dim rowIndex As Integer
    For rowIndex = 1 To UBound(dataCollection)
        ' Imprimir los datos de cada columna en la fila actual
        Debug.Print "Columna1: " & dataCollection(rowIndex).Columna1
        Debug.Print "Columna2: " & dataCollection(rowIndex).Columna2
        Debug.Print "Columna3: " & dataCollection(rowIndex).Columna3
        ' ...
        ' Imprimir los datos de las demás columnas
        
        Debug.Print "-----------------------"
    Next rowIndex
End Sub

