VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ARTICULOS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False



Private Type arti
    Codigo As String
    Descripcion As String
    Status As Long
    PesoUn As Double
    ESPECIFICACION As String
    Codigo_Ncm As String
    Unimed As String
End Type
'Declaramos el vector
Private articulos() As arti
Public Property Get Codigo(i As Variant) As String
   Codigo = articulos(i).Codigo
End Property
Public Property Get Descripcion(i As Variant) As String
   Descripcion = articulos(i).Descripcion
End Property
Public Property Get Status(i As Variant) As Integer
   Status = articulos(i).Status
End Property
Public Property Get PesoUn(i As Variant) As Double
   PesoUn = articulos(i).PesoUn
End Property
Public Property Get ESPECIFICACION(i As Variant) As String
   ESPECIFICACION = articulos(i).ESPECIFICACION
End Property
Public Property Get Codigo_Ncm(i As Variant) As String
   Codigo_Ncm = articulos(i).Codigo_Ncm
End Property
Public Property Get Unimed(i As Variant) As String
   Unimed = articulos(i).Unimed
End Property

Public Sub cargarvector()
  inicio1 = Now
  Erase articulos()
  SQLX$ = "select CODINT,Codigo,Status,Descripcion,pesuni,ESPECIFICACION,Codigo_Ncm,UMEDIDA"
  SQLX$ = SQLX$ + " from MASTER with(nolock) WHERE CODINT > 0 order by CODINT asc "
  Set RS = READSET(SQLX$)
  Do While Not RS.EOF
     CIX = XVALO(RS!CODINT)
     ReDim Preserve articulos(CIX)
     articulos(CIX).Codigo = SAFETEXT$(RS!Codigo)
     articulos(CIX).Descripcion = SAFETEXT$(RS!Descripcion)
     articulos(CIX).Status = XVALO(RS!Status)
     articulos(CIX).PesoUn = CDbl(XVALO(RS!PESUNI))
     articulos(CIX).ESPECIFICACION = SAFETEXT$(RS!ESPECIFICACION)
     articulos(CIX).Codigo_Ncm = SAFETEXT$(RS!Codigo_Ncm)
     articulos(CIX).Unimed = SAFETEXT$(RS!UMEDIDA)
     RS.MoveNext
  Loop
  RS.Close
  Set RS = Nothing
  salida1 = Now
99
End Sub


