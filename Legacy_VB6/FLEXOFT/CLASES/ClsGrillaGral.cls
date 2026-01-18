VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsGrillaGral"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
    Private Type Formulax
        formula As String
        ColEcoCampo As String
        Condicion1 As String
        Condicion2 As String
        TxtUbicadoGrillaTotal As String
        TOTALES As String
        DejarLibre As String
        Enumerar As String
        EnumerarRepUltimo As String
        RepiteDato As String
    End Type
    'Declaramos el vector
    
    Private FormulaGrid(32) As Formulax
    Private ColEcoCampoGrid(32) As Formulax
    Private Condicion1Grid(32) As Formulax
    Private Condicion2Grid(32) As Formulax

    Private TxtUbicadoGrillaTotales(32) As Formulax
    Private TOTALESGrid(32) As Formulax
    Private DejarLibreGrid(32) As Formulax
    Private EnumerarGrid(32) As Formulax
    Private EnumerarRepUltimoGrid(32) As Formulax
    Private RepiteDatoGrid(32) As Formulax
    
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
    End Type
    Private VentaNew  As VENTANA
    Public Property Get Campox() As String
       Campox = VentaNew.Campox
    End Property
    Public Property Let Campox(VALOR As String)
       VentaNew.Campox = VALOR
    End Property
    '
    Public Property Get TITULO() As String
       TITULO = VentaNew.TITULO
    End Property
    Public Property Let TITULO(VALOR As String)
       VentaNew.TITULO = VALOR
    End Property
    '
    
    Public Property Get AgregaRegistro() As String
       AgregaRegistro = VentaNew.AgregaRegistro
    End Property
    Public Property Let AgregaRegistro(VALOR As String)
       VentaNew.AgregaRegistro = VALOR
    End Property
    '
    Public Property Get UtilizaFormula() As Integer
       UtilizaFormula = VentaNew.UtilizaFormula
    End Property
    Public Property Let UtilizaFormula(VALOR As Integer)
       VentaNew.UtilizaFormula = VALOR
    End Property
    
    Public Property Get Inicializar() As Integer
       Inicializar = VentaNew.Inicializar
    End Property
    Public Property Let Inicializar(VALOR As Integer)
       VentaNew.Inicializar = VALOR
    End Property
    '
    Public Property Get CampEditables() As String
       CampEditables = VentaNew.CampEditables
    End Property
    Public Property Let CampEditables(VALOR As String)
       VentaNew.CampEditables = VALOR
    End Property
    '
    Public Property Get formula(i As Integer) As String
       formula = FormulaGrid(i).formula 'vector
       ' = "Indice=Columna(donde se muestra el resultado)| 4;*;6 texto de columnas y operadores de la formula separado por ';'"
    End Property
    Public Property Let formula(i As Integer, ByVal VALOR1 As String)
       FormulaGrid(i).formula = VALOR1  'vector
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
    Public Property Let HabilitaBorrar(VALOR As Integer)
       '0 NO MUESTRA EL BOTON
       '1 DESHABILITA BOTON BORRAR
       VentaNew.HabilitaBorrar = VALOR
    End Property
     Public Property Get HabilitaInsertar() As Integer
       HabilitaInsertar = VentaNew.HabilitaInsertar
    End Property
    Public Property Let HabilitaInsertar(VALOR As Integer)
       '0 NO MUESTRA EL BOTON
       '1 DESHABILITA BOTON BORRAR
       VentaNew.HabilitaInsertar = VALOR
    End Property
     Public Property Get NoMostrarBotonTilde() As Integer
       NoMostrarBotonTilde = VentaNew.NoMostrarBotonTilde
    End Property
    Public Property Let NoMostrarBotonTilde(VALOR As Integer)
       '0 VISIBLE
       '1 NO VISIBLE
       VentaNew.NoMostrarBotonTilde = VALOR
    End Property
     Public Property Get PermitirFilasVacias() As Integer
       PermitirFilasVacias = VentaNew.PermitirFilasVacias
    End Property
    Public Property Let PermitirFilasVacias(VALOR As Integer)
       '0 las FILAS VACIAS SE BORRAN
       '1 LAS FILAS VACIAS NO SE BORRAN
       VentaNew.PermitirFilasVacias = VALOR
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

