VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Zelecho"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
    Private Type Zelex
        Campox As String
        Inicializar As Integer
        Titulo As String
        Condix As String
        Tablax As String
        SelectCompleta As String
        CamposCFormato As String
        NombreZelechoArmado As String
        Anchox As Long
        SelectCount As String
        CantiReg  As Long
        CampoOrdenamiento As String
'        FORMUX As Form
    End Type

    Private Zelecho  As Zelex
    Public Property Get Campox() As String
       Campox = Zelecho.Campox
    End Property
    Public Property Let Campox(VALOR As String)
       Zelecho.Campox = VALOR
    End Property
    '
    Public Property Get Titulo() As String
       Titulo = Zelecho.Titulo
    End Property
    Public Property Let Titulo(VALOR As String)
       Zelecho.Titulo = VALOR
    End Property
    '
    Public Property Get Inicializar() As Integer
       Inicializar = Zelecho.Inicializar
    End Property
    Public Property Let Inicializar(VALOR As Integer)
       Zelecho.Inicializar = VALOR
    End Property

    Public Property Get Condix() As String
       Condix = Zelecho.Condix
    End Property
    Public Property Let Condix(VALOR As String)
       Zelecho.Condix = VALOR
    End Property

    Public Property Get Tablax() As String
       Tablax = Zelecho.Tablax
    End Property
    Public Property Let Tablax(VALOR As String)
       Zelecho.Tablax = VALOR
    End Property

    Public Property Get SelectCompleta() As String
       SelectCompleta = Zelecho.SelectCompleta
    End Property
    Public Property Let SelectCompleta(VALOR As String)
       Zelecho.SelectCompleta = VALOR
    End Property

    Public Property Get CamposCFormato() As String
       CamposCFormato = Zelecho.CamposCFormato
    End Property
    Public Property Let CamposCFormato(VALOR As String)
       Zelecho.CamposCFormato = VALOR
    End Property

    Public Property Get NombreZelechoArmado() As String
       NombreZelechoArmado = Zelecho.NombreZelechoArmado
    End Property
    Public Property Let NombreZelechoArmado(VALOR As String)
       Zelecho.NombreZelechoArmado = VALOR
    End Property

    Public Property Get Anchox() As Long
       Anchox = Zelecho.Anchox
    End Property
    Public Property Let Anchox(VALOR As Long)
       Zelecho.Anchox = VALOR
    End Property
    
    Public Property Get SelectCount() As String
       SelectCount = Zelecho.SelectCount
    End Property
    Public Property Let SelectCount(VALOR As String)
       Zelecho.SelectCount = VALOR
    End Property

    Public Property Get CantiReg() As Long
       CantiReg = Zelecho.CantiReg
    End Property
    Public Property Let CantiReg(VALOR As Long)
       Zelecho.CantiReg = VALOR
    End Property
    Public Property Get CampoOrdenamiento() As String
       CampoOrdenamiento = Zelecho.CampoOrdenamiento
    End Property
    Public Property Let CampoOrdenamiento(VALOR As String)
       Zelecho.CampoOrdenamiento = VALOR
    End Property

'    Public Property Get FORMUX() As Form
'       FORMUX = Zelecho.FORMUX
'    End Property
'    Public Property Let FORMUX(VALOR As Form)
'       Zelecho.FORMUX = VALOR
'    End Property
