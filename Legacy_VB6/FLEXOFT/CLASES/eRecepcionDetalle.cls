VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "eRecepcionDetalle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public CodigoInterno As Long
Public Lote As String
Public LoteProveedor As String
Public Cantidad As Double
Public M_Largo As Double
Public M_Ancho As Double
Public CantidadAprobada As Double
Public CantidadRechazada As Double
Public CantidadFaltante As Double
Public Peso As Double
Public KgBalanzapublica As Double
Public CantidadBultos As Integer

