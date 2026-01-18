VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Pedidos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False


Private ConexionBBDD As ADODB.Connection


'Public Function Nuevo(Encabezado As PedidoEncabezado, Detalle() As PedidoDetalle) As Long
'
'  ' PEDENCA
'
'  Nuevo = True
'
'  If Encabezado.NUMERO < 0 Then
'    ValorReferencia& = Abs(Encabezado.NUMERO)
'    Encabezado.NUMERO = ProximoNumero(ValorReferencia&)
'  End If
'
'  Dim PedEnca As ADODB.Recordset
'
'  On Error GoTo ControlErrorPedido
'  'FLEXCONN.BeginTrans
'
'  Set PedEnca = New ADODB.Recordset
'  PedEnca.CursorType = adOpenKeyset
'  PedEnca.LockType = adLockPessimistic
'  ConsultaSql$ = "SELECT * FROM PEDENCA WHERE NROPED=" & CStr(Encabezado.NUMERO)
'  PedEnca.Open FILTSQL$(ConsultaSql$), FLEXCONN, , , adCmdText
'  With PedEnca
'    .AddNew
'    !CODIEMPR = CodiEmp%
'    !NROPED = Encabezado.NUMERO
'    !tipoped = Encabezado.TipoPedido
'    !REFERENCIA = Left$(rasocli$(Encabezado.NumeroCliente), 35) & " - " & FECHATEX$(Encabezado.FechaEmision)
'    !FECHEMIS = CVDAT(Encabezado.FechaEmision)
'    !NROCLIE = Encabezado.NumeroCliente
'    !RasoClie = rasocli$(Encabezado.NumeroCliente)
'    !NROOCOM = Encabezado.ORDENCOMPRA
'    !DOMIENT = Encabezado.DomicilioEntrega
'    !LOCAENT = Encabezado.LocalidadEntrega
'    !TranspEnt = Encabezado.TransporteEntrega
'    !NTranspo = Encabezado.NumeroTransporte
'    !NSucent = Encabezado.NumeroSucursalEntrega
'    !Observa = ""
'    !fesolent = CVDAT(Encabezado.FechaSolicitudEntrega)
'    !Status = 1
'    !ListPre = Encabezado.ListaPrecio
'    !CondPag = Encabezado.CondicionPago
'    !VENDED = Encabezado.Vendedor
'    !TIPOINGRE = 0
'    !OV_Motivo = 0
'    !OV_AUTORIZA2 = ""
'    !OV_AUTORIZA1 = ""
'    !PorComVend = Encabezado.PorcentajeComisionVendedor
'    !Rec_ConPa = ""
'    !TDescApli = ""
'    !Dto_Ret_Cte = ""
'    !VenRep = Encabezado.VentaReparacion
'    !FLETE_TEX = ""
'    !CODIEMPR = CodiEmp%
'    !FechUltCambio = CVDAT(FECHANUM("01/01/80"))
'
'    .Update
'  End With
'20 PedEnca.Close
'   Set PedEnca = Nothing
'  ' PEDDETA
'
'  Dim PedDeta As ADODB.Recordset
'  Dim d As PedidoDetalle
'
'  Set PedDeta = New ADODB.Recordset
'  PedDeta.CursorType = adOpenKeyset
'  PedDeta.LockType = adLockPessimistic
'  ConsultaSql$ = "SELECT * FROM PEDDETA WHERE NROPED=0"
'  PedDeta.Open FILTSQL$(ConsultaSql$), FLEXCONN, , , adCmdText
'  NumeroOrden = 1
'  With PedDeta
'    For i& = 0 To UBound(Detalle) - 1
'        .AddNew
'        !CODIEMPR = CodiEmp%
'        !NROPED = Encabezado.NUMERO
'        !FECHEMIS = CVDAT(Encabezado.FechaEmision)
'        !NROCLIE = Encabezado.NumeroCliente
'        !RasoClie = rasocli$(Encabezado.NumeroCliente)
'        !NROOCOM = Encabezado.ORDENCOMPRA
'        !fesolent = CVDAT(Encabezado.FechaSolicitudEntrega)
'        !Status = 1
'        !ListPre = 0
'        !CondPag = 0
'        !VENDED = 0
'        !PorComVend = 0
'        !NuOrItem = NumeroOrden
'        !codiint = Detalle(i&).CodigoInterno
'        !codiext = CODEXT$(Detalle(i&).CodigoInterno)
'        If Detalle(i&).Descripcion <> "" Then
'            !Descrip = Detalle(i&).Descripcion
'        Else
'            !Descrip = DESCRIT0$(Detalle(i&).CodigoInterno)
'        End If
'        !SubCod = ""
'        !tcolor = Detalle(i&).COLOR
'        !ttalle = Detalle(i&).TALLE
'        !tlargo = Detalle(i&).Largo
'        !tancho = Detalle(i&).ANCHO
'        !tespesor = Detalle(i&).Espesor
'        !VenRep = Encabezado.VentaReparacion
'        !MoneEmis = 1
'        !TIPOCAM = 0
'        !lista_de_item = 0
'        !PreList = Detalle(i&).PrecioLista
'        !CanAsig = Detalle(i&).CantidadAsignada
'        !canped = Detalle(i&).CantidadPedida
'        !PreUnid = Detalle(i&).PrecioUnitario
'        !ImpoTot = Detalle(i&).ImporteTotal
'        !CANTENT = Detalle(i&).CantidadEntregada
'        !CanSaldo = Detalle(i&).CantidadSaldo
'        !feultent = CVDAT(FECHANUM("01/01/80"))
'        !NROREMI = ""
'        !TPedItem = ""
'        !ImpoAnti = 0
'        !FApro_Nope = CVDAT(FECHANUM("01/01/80"))
'        !acuprepa = 0
'        !ACUDESPA = 0
'        !tipoped = Encabezado.TipoPedido
'        !finvigen = CVDAT(FECHANUM("01/01/80"))
'        !StaDespa = 0
'        !CBase = ""
'        !DescLista = 0
'        !DescCombi = ""
'        !DescRubro = 0
'        !DescTotal = 0
'        !NUMEROSERIE = ""
'        !idenlote = Detalle(i&).LOTE
'        !Num_Prov = 0
'        !SIN_CARGO = 0
'        !ObservaItem = "" ' Detalle(i&).Observacion
'        !ARMADOR = 0
'        !FechUltCambio = CVDAT(FECHANUM("01/01/80"))
'
'        .Update
'
'    Next i&
'  End With
'  PedDeta.Close
'  Set PedDeta = Nothing
'
'  'FLEXCONN.CommitTrans
'
'  On Error GoTo 0
'  Nuevo = Encabezado.NUMERO
'
'  Exit Function
'
'ControlErrorPedido:
'
'  On Error GoTo 0
'  Nuevo = False
'  'FLEXCONN.RollbackTrans
'
'  On Error Resume Next
'  PedEnca.Close
'  Set PedEnca = Nothing
'  PedDeta.Close
'  Set PedDeta = Nothing
'  On Error GoTo 0
'
'End Function
'
Public Function ProximoNumero(ValorReferencia As Long) As Long

    ConsultaSql$ = "select max(nroped) as UltimoNumero from pedenca with(nolock)"
    If ValorReferencia > 0 Then
        ConsultaSql$ = ConsultaSql$ & " where nroped >= " & CStr(ValorReferencia)
    End If

    Dim Rst As ADODB.Recordset
    Set Rst = READSET(ConsultaSql$)
    If IsNull(Rst!UltimoNumero) Then
        ProximoNumero = ValorReferencia
    Else
        ProximoNumero = XVALO(Rst!UltimoNumero) + 1
    End If

End Function

Public Function ListaPedidosPendientes(Optional Cliente As Long) As Long

    ConsultaSql$ = "select pe.nroped, pe.fechemis, pe.referencia from pedenca pe with(nolock)" _
    & " inner join (select nroped from peddeta with(nolock) where status<3 and (canped-cantent)>0.05 group by nroped) pd" _
    & " on (pe.nroped=pd.nroped)" _
    & " where pe.status<3"
    If Cliente > 0 Then
        ConsultaSql$ = ConsultaSql$ & " and pe.nroclie=" & CStr(Cliente)
    End If
    ConsultaSql$ = ConsultaSql$ & " order by pe.nroped"
    Campos$ = "Pedido/I10;Emisión/D8;Referencia/A36"
    Call CARGA_ENCABEZADO(FRMZELECHO.msf2, Campos$, FRMZELECHO)
    Call CARGA_ENCABEZADO(FRMZELECHO.MSF1, Campos$, FRMZELECHO)
    FRMZELECHO.Caption = "Lista de Pedidos Pendientes"
    
    Dim Rst As ADODB.Recordset
    Set Rst = READSET(ConsultaSql$)
    Call Carga_MSF_Recordset(Rst, Campos$, FRMZELECHO.msf2, 1)
    On Error Resume Next
    Rst.Close
    Set Rst = Nothing
    On Error GoTo 0
    
    FRMZELECHO.Width = 8000
    FRMZELECHO.Show 1
    
    Pedido& = XVALO(RESP_ZELE_VECT(1))
    ListaPedidosPendientes = Pedido&
    
End Function


