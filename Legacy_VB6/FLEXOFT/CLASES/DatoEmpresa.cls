VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DatoEmpresa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Type agenda
    NombreEmpresa As String
    razonSocial As String
    CalleNumero As String
    Localidad As String
    CodPostal As String
    Telefono As String
    Cuit As String
    IngBruto As String
    CajaPrevisional As String
    LetProvincia As String
    
End Type

Private personas As agenda

Public Property Get NombreEmpresa() As String
   NombreEmpresa = personas.NombreEmpresa
End Property
Public Property Get razonSocial() As String
   razonSocial = personas.razonSocial
End Property
Public Property Get CalleNumero() As String
   CalleNumero = personas.CalleNumero
End Property
Public Property Get Localidad() As String
   Localidad = personas.Localidad
End Property
Public Property Get CodPostal() As String
   CodPostal = personas.CodPostal
End Property
Public Property Get Telefono() As String
   Telefono = personas.Telefono
End Property
Public Property Get Cuit() As String
   Cuit = personas.Cuit
End Property
Public Property Get IngBruto() As String
   IngBruto = personas.IngBruto
End Property
Public Property Get CajaPrevisional() As String
   CajaPrevisional = personas.CajaPrevisional
End Property
Public Property Get LetProvincia() As String
   LetProvincia = personas.LetProvincia
End Property

Private Sub Class_Initialize()
    LU$ = LUDAT$
    N1% = FILEOPEN(LU$ + "SVD147.EXE", 0, 41)
    Dim X As String * 41
    '
    For kk% = 1 To 8
        '
        If kk% > LOF(N1%) / 41 Then
              X = ""
           Else
              Get #N1%, kk%, X
              
              Select Case kk%
                Case 1: personas.NombreEmpresa = DESENCRI$(X)
                Case 2: personas.razonSocial = DESENCRI$(X)
                Case 3: personas.CalleNumero = DESENCRI$(X)
                Case 4: personas.Localidad = DESENCRI$(X)
                Case 5: personas.Telefono = DESENCRI$(X)
                Case 6: personas.Cuit = DESENCRI$(X)
                Case 7: personas.IngBruto = DESENCRI$(X)
                Case 8: personas.CajaPrevisional = DESENCRI$(X)
              End Select
        End If
        '
        personas.CodPostal = CONTROL("FSTCOPOS", "CODIPOST")
        personas.LetProvincia = CONTROL("FSTCOPOS", "LETRAPRO")
    Next kk%

End Sub
