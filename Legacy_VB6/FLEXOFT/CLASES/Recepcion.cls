VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Recepcion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Function Nuevo(Encabezado As eRecepcion, Detalle() As eRecepcionDetalle) As Long

    Nuevo = -1
    
    On Error GoTo ControlError
    ConsultaSql$ = "select max(numebore) ultimo from bolrecep"
    Set Rst = FLEXCONN.Execute(ConsultaSql$)
    ProximaRecepcion& = XVALO(Rst!ultimo) + 1
    Rst.Close
    Set Rst = Nothing
    
    Dim RstRecepcion As New ADODB.Recordset
    Dim ItemDetalle As eRecepcionDetalle
    ConsultaSql$ = "select * from bolrecep where numebore=0"
    RstRecepcion.Open FILTSQL$(ConsultaSql$), FLEXCONN, adOpenKeyset, adLockPessimistic
    
    With RstRecepcion
        NumeroRecepcion& = ProximaRecepcion&
        NumeroProveedor& = Encabezado.NumeroProveedor
        CodigoEmpresa% = CODIEMPR%
        FechaRecepcion% = HOY(HO$)
        RemitoProveedor$ = Encabezado.RemitoProveedor
        FacturaProveedor$ = Encabezado.FacturaProveedor
        PatenteTransporte$ = Encabezado.PatenteTransporte
        Deposito% = Encabezado.Deposito
        UbicacionFisica$ = Encabezado.UbicacionFisica
        NumConvig$ = Encabezado.NumConvig
        Status% = Encabezado.Status
        
        For i = 0 To UBound(Detalle) - 1
            
            Set ItemDetalle = Detalle(i)
            
            .AddNew
            
            !NUMEBORE = NumeroRecepcion
            !CODIEMPR = CodigoEmpresa%
            !idenlote = "BR-" & CStr(NumeroRecepcion) & "-" & CStr(i + 1)
            !fechrecep = CVDAT(FechaRecepcion%)
            !FEYHORECEP = Now
            !cod_inte = ItemDetalle.CodigoInterno
            !cod_exte = CODEXT$(ItemDetalle.CodigoInterno)
            !descripitem = DESCRIT0$(ItemDetalle.CodigoInterno)
            !Descrip = DESCRIT0$(ItemDetalle.CodigoInterno)
            !unidad = Unimed(ItemDetalle.CodigoInterno)
            !cantnom = ItemDetalle.Cantidad
            !M_Largo = ItemDetalle.M_Largo
            !M_Ancho = ItemDetalle.M_Ancho
            !depoent = Deposito%
            !ubicfis = UbicacionFisica$
            !numprov = NumeroProveedor&
            !raso_prov = rasocli$(NumeroProveedor&)
            !lotecol_prov = ItemDetalle.LoteProveedor
            !numrem_prov = RemitoProveedor$
            !numfac_prov = FacturaProveedor$
            !NumConvig = NumConvig$
            !cantoapro = ItemDetalle.CantidadAprobada
            !cantorech = ItemDetalle.CantidadRechazada
            !cantofalta = ItemDetalle.CantidadFaltante
            !Status = Status%
            !Peso = ItemDetalle.Cantidad
            !PatenteTransporte = PatenteTransporte$
            !KgBalanzapublica = ItemDetalle.KgBalanzapublica
            !CantidadBultos = ItemDetalle.CantidadBultos
            !NUMOC = 0
            !nroitemoc = 0
            !fechlibe = CVDAT(FechaRecepcion%)
            !numeordser = 0
            !Observa = ""
            !usuapro = USERNAME$
            !fehoapro = CVDAT(FechaRecepcion%)
            
            .Update
        Next i
    End With
    On Error GoTo 0
    
    On Error Resume Next
    RstRecepcion.Close
    Set RstRecepcion = Nothing
    On Error GoTo 0

    Nuevo = ProximaRecepcion&
    
    Exit Function

ControlError:

    
End Function
