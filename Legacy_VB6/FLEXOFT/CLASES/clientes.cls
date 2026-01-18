VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clientes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"


Private Type agenda
    LpreCli As Integer
    nombre As String
    Domicilio As String
    localidad As String
    mailcliente As String
    TeleCli As String
    CeluCli As String
    NotaCli As String
    CtacCli As String
    CuitCli As String
End Type
'Declaramos el vector
Private personas() As agenda
Public Property Get nombre(i As Variant) As String
   On Error Resume Next
   nombre = personas(i).nombre
   On Error GoTo 0
End Property
Public Property Get Domicilio(i As Variant) As String
   On Error Resume Next
   Domicilio = personas(i).Domicilio
   On Error GoTo 0
End Property
Public Property Get localidad(i As Variant) As String
   On Error Resume Next
   localidad = personas(i).localidad
   On Error GoTo 0
End Property
Public Property Get mailcliente(i As Variant) As String
   On Error Resume Next
   mailcliente = personas(i).mailcliente
   On Error GoTo 0
End Property
Public Property Get LpreCli(i As Variant) As String
   On Error Resume Next
   LpreCli = personas(i).LpreCli
   On Error GoTo 0
End Property
Public Property Get TeleCli(i As Variant) As String
   On Error Resume Next
   TeleCli = personas(i).TeleCli
   On Error GoTo 0
End Property
Public Property Get CeluCli(i As Variant) As String
   On Error Resume Next
   CeluCli = personas(i).CeluCli
   On Error GoTo 0
End Property
Public Property Get NotaCli(i As Variant) As String
   On Error Resume Next
   NotaCli = personas(i).NotaCli
   On Error GoTo 0
End Property
Public Property Get CtacCli(i As Variant) As String
   On Error Resume Next
   CtacCli = personas(i).CtacCli
   On Error GoTo 0
End Property
Public Property Get CuitCli(i As Variant) As String
   On Error Resume Next
   CuitCli = personas(i).CuitCli
   On Error GoTo 0
End Property



Public Sub cargarvector()
  inicio1 = Now
  Erase personas()
  SQLX$ = "select Raso_Cli,nume_cli,Domi_Cli,loca_Cli,Mail_Cli,Tele_Cli,CELU_CLI,nota_Cli,CTAC_CLI,Lpre_Cli,Cuit_Cli"
  SQLX$ = SQLX$ + " from deudor12 with(nolock) order by Nume_Cli asc "
  Set RS = READSET(SQLX$)
  Do While Not RS.EOF
     NC = XVALO(RS!nume_cli)
     ReDim Preserve personas(NC)
     personas(NC).nombre = SAFETEXT$(RS!raso_cli)
     personas(NC).Domicilio = SAFETEXT$(RS!domi_cli)
     personas(NC).localidad = SAFETEXT$(RS!loca_Cli)
     personas(NC).mailcliente = SAFETEXT$(RS!Mail_Cli)
     '
     personas(NC).TeleCli = SAFETEXT$(RS!Tele_Cli)
     personas(NC).CeluCli = SAFETEXT$(RS!CELU_CLI)
     personas(NC).NotaCli = SAFETEXT$(RS!nota_Cli)
     personas(NC).CtacCli = SAFETEXT$(RS!CTAC_CLI)
     personas(NC).LpreCli = XVALO(RS!Lpre_Cli)
     personas(NC).CuitCli = SAFETEXT$(RS!Cuit_Cli)
     '
     RS.MoveNext
  Loop
  RS.Close
  Set RS = Nothing
  salida1 = Now
End Sub
