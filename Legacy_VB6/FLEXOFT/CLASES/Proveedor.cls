VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Proveedor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False


Private Type agenda
    nombre As String
    Domicilio As String
    localidad As String
End Type
Private Type CTAXDF
    CTA As Integer
End Type
'Declaramos el vector
Private personas() As agenda
Private CTAXDFTO() As CTAXDF


Public Property Get nombre(i As Variant) As String
   nombre = personas(i).nombre
End Property
Public Property Get Domicilio(i As Variant) As String
   Domicilio = personas(i).Domicilio
End Property
Public Property Get localidad(i As Variant) As String
   localidad = personas(i).localidad
End Property
Public Property Get CTA(J As Variant) As Integer
   CTA = CTAXDFTO(J).CTA
End Property

Public Sub cargarvector()
  inicio1 = Now
  Erase personas()
  sqlx$ = "select Raso_Cli,nume_cli,Domi_Cli,loca_Cli from proved12 with(nolock) order by Nume_Cli asc "
  Set RS = READSET(sqlx$)
  Do While Not RS.EOF
     NC = XVALO(RS!nume_cli)
     ReDim Preserve personas(NC)
     personas(NC).nombre = SAFETEXT$(RS!raso_cli)
     personas(NC).Domicilio = SAFETEXT$(RS!domi_cli)
     personas(NC).localidad = SAFETEXT$(RS!loca_Cli)
     RS.MoveNext
  Loop
  RS.Close
  Set RS = Nothing
  salida1 = Now

End Sub
Public Sub CARGAVECTORCTAXDFTO()
   Erase CTAXDFTO()
   ReDim Preserve CTAXDFTO(32767)
   sqlx$ = "SELECT VALCONTROL,CLABUS FROM FLEXCRL WITH(NOLOCK) "
   sqlx$ = sqlx$ + " WHERE PROBUS = 'CTA X DEFECTO CARGA DETALLE' AND CLABUS <> ''  "
   Set RS = READSET(sqlx$)
   With RS
     Do While Not .EOF
       NPROV = XVALO(!CLABUS)
       CTAXDFTO(NPROV).CTA = XVALO(!VALCONTROL)
       .MoveNext
     Loop
   End With

End Sub

