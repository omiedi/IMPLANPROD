VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ListaSeleccion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public TITULO As String
Public EncabezadoDeColumnas As String
Public ConsultaSql As String
Public Grilla As MSFlexGrid
Public AnchoLista As Long
Public ResaltaFilaElegida As Boolean
Public SeleccionAprobada As Boolean

Private Sub Crear()

    TituloLista$ = TITULO
    Campos$ = EncabezadoDeColumnas
    Call CARGA_ENCABEZADO(FRMZELECHO.msf2, Campos$, FRMZELECHO)
    Call CARGA_ENCABEZADO(FRMZELECHO.MSF1, Campos$, FRMZELECHO)
    FRMZELECHO.Caption = TituloLista$
    PINTAFILACLICK% = 0
    If ResaltaFilaElegida = True Then PINTAFILACLICK% = 1
    
    If ConsultaSql$ <> "" Then
        Dim Rst As ADODB.Recordset
        Set Rst = FLEXCONN.Execute(ConsultaSql)
        Call Carga_MSF_Recordset(Rst, Campos$, FRMZELECHO.msf2, 1)
        On Error Resume Next
        Rst.Close
        Set Rst = Nothing
        On Error GoTo 0
    Else
        If Grilla.Rows > 0 Then
            For fi = 1 To Grilla.Rows - 1
                If Grilla.TextMatrix(fi, 0) = "*" Then
                    FRMZELECHO.msf2.Rows = FRMZELECHO.msf2.Rows + 1
                    For co = 1 To Grilla.Cols - 1
                        FRMZELECHO.msf2.TextMatrix(FRMZELECHO.msf2.Rows - 1, co) = Grilla.TextMatrix(fi, co)
                    Next co
                End If
            Next fi
        End If
    End If
    
    If AnchoDeLista > 0 Then FRMZELECHO.Width = AnchoDeLista
    
End Sub

Public Sub Mostrar(Optional RecargaDatos As Boolean = True)

    If RecargaDatos = True Then Call Crear
    If ResaltaFilaElegida = True Then
        SeleccionAprobada = False
        FRMZELECHO.mnuImprimir.Caption = "Aprobar selección"
        FRMZELECHO.mnuImprimir.Visible = True
    End If
    
    If AnchoLista > 0 Then FRMZELECHO.Width = AnchoLista
    FRMZELECHO.Show 1
    
    If FRMZELECHO.IMPRIMIR = 1 Then SeleccionAprobada = True
    PINTAFILACLICK% = 0
    FRMZELECHO.mnuImprimir.Caption = "Imprimir"
    FRMZELECHO.mnuImprimir.Visible = False
    
    
End Sub

Public Function FilaElegida(IdCampo As Integer) As String

    FilaElegida = RESP_ZELE_VECT(IdCampo)

End Function

Public Function EsFilaSeleccionada(Fila As Long) As Boolean

    EsFilaSeleccionada = False
    If FRMZELECHO.msf2.TextMatrix(Fila, 0) = "*" Then
        EsFilaSeleccionada = True
        For i& = 1 To FRMZELECHO.msf2.Cols - 1
          RESP_ZELE_VECT(i&) = FRMZELECHO.msf2.TextMatrix(Fila, i&)
        Next i&
    End If
    
End Function

Public Function TotalFilasDeLaLista() As Long

    TotalFilasDeLaLista = FRMZELECHO.msf2.Rows - 1
    
End Function
