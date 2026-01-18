VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form FRMEXPOTOTVS 
   BackColor       =   &H00E2D3C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Exportación de Datos"
   ClientHeight    =   2820
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8580
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2820
   ScaleWidth      =   8580
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   8040
      Top             =   240
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox MARCOBARRA 
      BackColor       =   &H002B9973&
      Height          =   200
      Left            =   2040
      ScaleHeight     =   135
      ScaleWidth      =   4545
      TabIndex        =   6
      Top             =   2160
      Width           =   4605
      Begin VB.Frame BARRATRAZA 
         BackColor       =   &H0048CC9D&
         BorderStyle     =   0  'None
         Height          =   500
         Left            =   -120
         TabIndex        =   7
         Top             =   0
         Width           =   12000
      End
   End
   Begin VB.CommandButton btnEquipos 
      Caption         =   "Equipos"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5880
      TabIndex        =   5
      Top             =   1320
      Width           =   2055
   End
   Begin VB.CommandButton btnListasDePrecios 
      Caption         =   "Listas de Precios"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3240
      TabIndex        =   4
      Top             =   1320
      Width           =   2055
   End
   Begin VB.CommandButton btnEstructuras 
      Caption         =   "Estructura de Productos"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   600
      TabIndex        =   3
      Top             =   1320
      Width           =   2055
   End
   Begin VB.CommandButton btnProductos 
      Caption         =   "Productos"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5880
      TabIndex        =   2
      Top             =   480
      Width           =   2055
   End
   Begin VB.CommandButton btnProveedores 
      Caption         =   "Proveedores"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3240
      TabIndex        =   1
      Top             =   480
      Width           =   2055
   End
   Begin VB.CommandButton btnClientes 
      Caption         =   "Clientes"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   600
      TabIndex        =   0
      Top             =   480
      Width           =   2055
   End
   Begin VB.Label lblTablaExportando 
      Alignment       =   2  'Center
      BackColor       =   &H8000000C&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   240
      Left            =   2040
      TabIndex        =   9
      Top             =   120
      Visible         =   0   'False
      Width           =   4575
   End
   Begin VB.Label lblTextoAvance 
      Alignment       =   2  'Center
      BackColor       =   &H00E2D3C0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   2040
      TabIndex        =   8
      Top             =   2505
      Width           =   4575
   End
End
Attribute VB_Name = "FRMEXPOTOTVS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Sub EjemploCreacionDBF()

Dim db As Database
Dim RecDBF As Recordset
Dim MyTableDef As TableDef
Dim MyField As Field

Set db = Workspaces(0).OpenDatabase("C:\FLEXOFT", False, False, "dBASE III;DATABASE=PRUEBA.DBF")
Set MyTableDef = db.CreateTableDef("Test1")

' 1 TEXT Character 6
' 2 date date 8
' 3 DOUBLE Numeric 20 5
' 4 BOOL Logical 1
' 5 INTEGER Numeric 6
' 6 LONG Numeric 11
' 7 SINGLE Numeric 20 5

'add fields to the table definition
Set MyField = MyTableDef.CreateField("Text", dbText, 6)
MyTableDef.Fields.Append MyField
Set MyField = MyTableDef.CreateField("date", dbDate, 8)
MyTableDef.Fields.Append MyField
Set MyField = MyTableDef.CreateField("Double", dbDouble, 8)
MyTableDef.Fields.Append MyField
Set MyField = MyTableDef.CreateField("Bool", dbBoolean, 1)
MyTableDef.Fields.Append MyField
Set MyField = MyTableDef.CreateField("Integer", dbInteger, 8)
MyTableDef.Fields.Append MyField
Set MyField = MyTableDef.CreateField("Long", dbLong, 8)
MyTableDef.Fields.Append MyField
Set MyField = MyTableDef.CreateField("Single", dbSingle, 8)
MyTableDef.Fields.Append MyField

'add the table to the database
db.TableDefs.Append MyTableDef

db.Close
Set db = Nothing

End Sub

Private Sub btnClientes_Click()

    Dim db As Database
    Dim RecordsetDBF As Recordset
    Dim MyTable As TableDef
    Dim MyField As Field
    Dim NombreBaseDatos As String
    Dim MensajeError As String
    Dim RutaArchivo As String
    Dim ArchivoSalida As String
    Dim RutaUsuario As String
       
    ArchivoSalida = "SA1"
    
    CommonDialog1.CancelError = True ' SELECCION DE ARCHIVO A IMPORTAR Y GRABACION
    CommonDialog1.DialogTitle = "Exportacion de Datos"
    CommonDialog1.FileName = ArchivoSalida & ".DBF"
    CommonDialog1.Filter = "Archivos DBF(*.DBF)|*.DBF|"
    CommonDialog1.FilterIndex = 1
    On Error GoTo 99
    CommonDialog1.ShowOpen
    On Error GoTo 0
    RutaUsuario = TRIM$(CommonDialog1.FileName)
    If RutaUsuario = "" Or InStr(1, RutaUsuario, ArchivoSalida & ".DBF", vbTextCompare) < 1 Then
        MensajeError = "Falta Carpeta de Archivo o Se Ha Modificado el Nombre del Archivo"
        GoTo 98
    End If
    
    RutaArchivo = Mid(RutaUsuario, 1, Len(RutaUsuario) - Len("\" & ArchivoSalida & ".DBF"))

    
    ' ELIMINA EL ARCHIVO EXISTENTE
    On Error Resume Next
    Kill RutaUsuario
    If Err > 0 And Err <> 53 Then
        MensajeError = "Se Produjo un Error al Eliminar el Archivo '" & REPLACAR$(RutaArchivo, "\", "/") & "/" & ArchivoSalida & ".DBF'. \El Proceso No Continuará."
        GoTo 98
    End If
    On Error GoTo 0
    
    Screen.MousePointer = 11
    
    lblTablaExportando = "Exportando Clientes"
    Timer1.Enabled = True
    
    ' CREAR EL ARCHIVO DBF
    On Error GoTo 98
    MensajeError = ""
    Set db = Workspaces(0).OpenDatabase(RutaArchivo, False, False, "dBASE IV;DATABASE=" & ArchivoSalida & ".DBF")
    
    ' CREA LA TABLA Y SUS CAMPOS
    Set MyTable = db.CreateTableDef(ArchivoSalida)
    
    Call NuevoCampo(MyTable, "A1_FILIAL", dbText, 4)
    Call NuevoCampo(MyTable, "A1_COD", dbText, 13)
    Call NuevoCampo(MyTable, "A1_LOJA", dbText, 2)
    Call NuevoCampo(MyTable, "A1_NOME", dbText, 50)
    Call NuevoCampo(MyTable, "A1_PESSOA", dbText, 1)
    Call NuevoCampo(MyTable, "A1_NREDUZ", dbText, 50)
    Call NuevoCampo(MyTable, "A1_END", dbText, 50)
    Call NuevoCampo(MyTable, "A1_TIPO", dbText, 1)
    Call NuevoCampo(MyTable, "A1_EST", dbText, 2)
    Call NuevoCampo(MyTable, "A1_MUN", dbText, 15)
    Call NuevoCampo(MyTable, "A1_TEL", dbText, 25)
    Call NuevoCampo(MyTable, "A1_TELEX", dbText, 30)
    Call NuevoCampo(MyTable, "A1_FAX", dbText, 15)
    Call NuevoCampo(MyTable, "A1_ENDCOB", dbText, 40)
    Call NuevoCampo(MyTable, "A1_CGC", dbText, 14)
    Call NuevoCampo(MyTable, "A1_MOEDALC", dbInteger)
    Call NuevoCampo(MyTable, "A1_PERCIVA", dbText, 1)
    Call NuevoCampo(MyTable, "A1_TIPCLI", dbText, 1)
    Call NuevoCampo(MyTable, "A1_CONDPAG", dbText, 5)
    Call NuevoCampo(MyTable, "A1_AGENPER", dbText, 1)
    Call NuevoCampo(MyTable, "A1_RECCOFI", dbText, 1)
    Call NuevoCampo(MyTable, "A1_RECCSLL", dbText, 1)
    Call NuevoCampo(MyTable, "A1_RECPIS", dbText, 1)
    Call NuevoCampo(MyTable, "A1_CONTAB", dbText, 15)
    Call NuevoCampo(MyTable, "A1_B2B", dbText, 1)
    Call NuevoCampo(MyTable, "A1_MSBLQL", dbText, 1)
    Call NuevoCampo(MyTable, "A1_RG1415", dbText, 1)
    Call NuevoCampo(MyTable, "A1_AGREGAN", dbText, 2)
    Call NuevoCampo(MyTable, "A1_INSCGAN", dbText, 1)
    Call NuevoCampo(MyTable, "A1_ABATIMP", dbText, 1)
    Call NuevoCampo(MyTable, "A1_SITUACA", dbText, 1)
    Call NuevoCampo(MyTable, "A1_IRBAX", dbText, 1)
    
    ' AGREGA LA TABLA
    db.TableDefs.Append MyTable
    On Error GoTo 0
    
    ' LEE LA TABLA SQL
    Call ABRECONEXION
    Dim dbClientes As ADODB.Recordset
    Dim PosIva As Integer
    Dim PosicionIva As String
    Dim Personeria As String
    Dim PagaIva As String
    Dim ClienteSuspendido As String
    Dim CuitCliente As String
    Dim CuentaMadre As String
    Dim CuentaMadreDefault As String

    On Error GoTo 98
    MensajeError = "Error al Leer Tabla de Articulos"
    Set dbClientes = READSET("SELECT * FROM DEUDOR12 WITH(NOLOCK) WHERE NUME_CLI > 0 ")
    FIN& = CantRegistros("DEUDOR12", "NUME_CLI>0")
    On Error GoTo 0
    
    ' ABRE LA TABLA DBF EN MODO ESCRITURA
    On Error GoTo 98
    MensajeError = "Error al Leer Archivo DBF Para Escritura"
    Set RecordsetDBF = db.OpenRecordset(ArchivoSalida, dbOpenDynaset)
    On Error GoTo 0
       
    On Error GoTo 98
    MensajeError = ""
    CuentaMadreDefault = TRIM$(CONTROL$("AMACLI", "ECUECON"))
    I& = 1
    With dbClientes
        Do While Not .EOF
        
            I& = I& + 1
            Call TRACE(24, I&, FIN&)
            
            ' AGREGAR LOS REGISTROS DESDE LA TABLA SQL A LA TABLA DBF
            RecordsetDBF.AddNew
            RecordsetDBF!A1_FILIAL = ""
            RecordsetDBF!A1_COD = CStr(XVALO(!NUME_CLI))
            RecordsetDBF!A1_LOJA = ""
            RecordsetDBF!A1_NOME = Left$(SAFETEXT$(!RASO_CLI), 50)
                Personeria = "J"
                PosIva = XVALO(!PIVA_CLI)
                If PosIva = 0 And PosIva = 6 Then Personeria = "F"
            RecordsetDBF!A1_PESSOA = Personeria
            RecordsetDBF!A1_NREDUZ = Left$(SAFETEXT$(!FANT_CLI), 50)
            RecordsetDBF!A1_END = Left$(SAFETEXT$(!DOMI_CLI), 50)
                PosicionIva = "I"   ' RESP. INSCRIPTO
                Select Case PosIva
                    Case 0: PosicionIva = "F"       ' CONSUMIDOR FINAL
                    Case 2: PosicionIva = "N"       ' RESP. NO INSCRIPTO
                    Case 4: PosicionIva = "X"       ' EXENTO
                    Case 5, 7: PosicionIva = "E"    ' EXTERIOR
                    Case 6: PosicionIva = "M"       ' MONOTRIBUTO
                End Select
            RecordsetDBF!A1_TIPO = PosicionIva
            RecordsetDBF!A1_EST = Left$(SAFETEXT$(!PROV_CLI), 2)
            RecordsetDBF!A1_MUN = Left$(SAFETEXT$(!LOCA_CLI), 15)
            RecordsetDBF!A1_TEL = Left$(SAFETEXT$(!TELE_CLI), 25)
            RecordsetDBF!A1_TELEX = ""
            RecordsetDBF!A1_FAX = Left$(SAFETEXT$(!FAXI_CLI), 15)
            RecordsetDBF!A1_ENDCOB = ""
                CuitCliente = REPLACAR(SAFETEXT$(!CUIT_CLI), "-", "")
            RecordsetDBF!A1_CGC = Left$(CuitCliente, 14)
                MonedaCliente = 1
                If PosIva = 5 Then MonedaCliente = 2
            RecordsetDBF!A1_MOEDALC = 0
                PagaIva = "S"
                If PosIva = 4 Then PagaIva = "N"
            RecordsetDBF!A1_PERCIVA = PagaIva
            RecordsetDBF!A1_TIPCLI = 0
            RecordsetDBF!A1_CONDPAG = XVALO(!CPAG_CLI)
            RecordsetDBF!A1_AGENPER = "N"   ' NO
            RecordsetDBF!A1_RECCOFI = "" ' "N"   ' NO
            RecordsetDBF!A1_RECCSLL = "" ' "N"   ' NO
            RecordsetDBF!A1_RECPIS = "" ' "N"    ' NO
                CuentaMadre = TRIM$(CODCUE$(XVALO(!CUMA_CLI)))
                If CuentaMadre = "" Then CuentaMadre = CuentaMadreDefault
            RecordsetDBF!A1_CONTAB = CuentaMadre
            RecordsetDBF!A1_B2B = "" ' "2"       ' NO
            RecordsetDBF!A1_MSBLQL = "" ' "2"    ' NO
            RecordsetDBF!A1_RG1415 = "" ' "2"    ' NO
            RecordsetDBF!A1_AGREGAN = "" ' ""
            RecordsetDBF!A1_INSCGAN = "" ' "2"   ' NO
            RecordsetDBF!A1_ABATIMP = "" ' "3"   ' NO EFECTUA RETENCION
            RecordsetDBF!A1_SITUACA = "1"   ' NORMAL
            RecordsetDBF!A1_IRBAX = "" ' "2"     ' NO
            RecordsetDBF.Update
            
            .MoveNext
        Loop
    End With
    On Error GoTo 0
    
    ' CIERRA LOS OBJETOS DE ACCESO A SQL Y AL ARCHIVO DBF
    On Error GoTo 98
    MensajeError = "Error al Cerrar Objetos"
    
    dbClientes.Close
    Set dbClientes = Nothing
    RecordsetDBF.Close
    Set RecordsetDBF = Nothing
    db.Close
    Set db = Nothing
    
    On Error GoTo 0

    Screen.MousePointer = 1
    Timer1.Enabled = False: lblTablaExportando = ""
    Call COMUNI("Exportación de Clientes Finalizada")
    Exit Sub

    ' CAPTURA DE ERROR
    
98  MensajeVisual = Err.Description
    If TRIM$(MensajeError) = "" Then MensajeError = MensajeVisual
    On Error GoTo 0
    
    On Error Resume Next
    dbClientes.Close
    Set dbClientes = Nothing
    RecordsetDBF.Close
    Set RecordsetDBF = Nothing
    db.Close
    Set db = Nothing
    On Error GoTo 0
    
    Screen.MousePointer = 1
    Timer1.Enabled = False: lblTablaExportando = ""
    Call COMUNI(MensajeError)

99  Screen.MousePointer = 1
    On Error GoTo 0
    
    On Error Resume Next
    dbClientes.Close
    Set dbClientes = Nothing
    RecordsetDBF.Close
    Set RecordsetDBF = Nothing
    db.Close
    Set db = Nothing
    On Error GoTo 0
    
    Timer1.Enabled = False: lblTablaExportando = ""
    Exit Sub
End Sub

Private Sub btnEquipos_Click()

    Dim db1 As Database
    Dim db2 As Database
    Dim RecordsetDBF As Recordset
    Dim MyTable As TableDef
    Dim MyField As Field
    Dim NombreBaseDatos As String
    Dim MensajeError As String
    Dim RutaArchivo As String
    Dim ArchivoSalida1 As String
    Dim ArchivoSalida2 As String
    Dim RutaUsuario As String
    
    ArchivoSalida1 = "AA3"
    
    CommonDialog1.CancelError = True ' SELECCION DE ARCHIVO A IMPORTAR Y GRABACION
    CommonDialog1.DialogTitle = "Exportacion de Datos"
    CommonDialog1.FileName = ArchivoSalida1 & ".DBF"
    CommonDialog1.Filter = "Archivos DBF(*.DBF)|*.DBF|"
    CommonDialog1.FilterIndex = 1
    On Error GoTo 99
    CommonDialog1.ShowOpen
    On Error GoTo 0
    RutaUsuario = TRIM$(CommonDialog1.FileName)
    If RutaUsuario = "" Or InStr(1, RutaUsuario, ArchivoSalida1 & ".DBF", vbTextCompare) < 1 Then
        MensajeError = "Falta Carpeta de Archivo o Se Ha Modificado el Nombre del Archivo"
        GoTo 98
    End If
    
    RutaArchivo = Mid(RutaUsuario, 1, Len(RutaUsuario) - Len("\" & ArchivoSalida1 & ".DBF"))
    
    ArchivoSalida2 = "AA4"

    ' *** ELIMINA EL ARCHIVO EXISTENTE
    On Error Resume Next
    Kill RutaArchivo & "\" & ArchivoSalida1 & ".DBF"
    If Err > 0 And Err <> 53 Then
        MensajeError = "Se Produjo un Error al Eliminar el Archivo '" & REPLACAR$(RutaArchivo, "\", "/") & "/" & ArchivoSalida1 & ".DBF .\El Proceso No Continuará."
        GoTo 98
    End If
    
    Kill RutaArchivo & "\" & ArchivoSalida2 & ".DBF"
    If Err > 0 And Err <> 53 Then
        MensajeError = "Se Produjo un Error al Eliminar el Archivo '" & REPLACAR$(RutaArchivo, "\", "/") & "/" & ArchivoSalida2 & ".DBF .\El Proceso No Continuará."
        GoTo 98
    End If
    On Error GoTo 0
    
    Screen.MousePointer = 11
    
    lblTablaExportando = "Exportando Equipos"
    Timer1.Enabled = True
    
    ' *** CREAR EL ARCHIVO DBF
    On Error GoTo 98
    MensajeError = ""
    Set db1 = Workspaces(0).OpenDatabase(RutaArchivo, False, False, "dBASE IV;DATABASE=" & ArchivoSalida1 & ".DBF")
    
    ' *** CREA LA TABLA Y SUS CAMPOS CABECERA
    Set MyTable = db1.CreateTableDef(ArchivoSalida1)
    
    Call NuevoCampo(MyTable, "AA3_FILIAL", dbText, 4)
    Call NuevoCampo(MyTable, "AA3_CODCLI", dbText, 13)
    Call NuevoCampo(MyTable, "AA3_LOJA", dbText, 2)
    Call NuevoCampo(MyTable, "AA3_NOMCLI", dbText, 30)
    Call NuevoCampo(MyTable, "AA3_CODPRO", dbText, 15)
    Call NuevoCampo(MyTable, "AA3_DESPRO", dbText, 30)
    Call NuevoCampo(MyTable, "AA3_NUMSER", dbText, 20)
    Call NuevoCampo(MyTable, "AA3_DTVEND", dbDate)
    Call NuevoCampo(MyTable, "AA3_DTINST", dbDate, 4)
    Call NuevoCampo(MyTable, "AA3_DTGAR", dbDate, 4)
    Call NuevoCampo(MyTable, "AA3_NFVEND", dbText, 12)
    Call NuevoCampo(MyTable, "AA3_STATUS", dbText, 2)
    
    ' AGREGA TABLA A LA BD
    db1.TableDefs.Append MyTable
    On Error GoTo 0
    
    On Error GoTo 98
    MensajeError = ""
    Set db2 = Workspaces(0).OpenDatabase(RutaArchivo, False, False, "dBASE IV;DATABASE=" & ArchivoSalida2 & ".DBF")
    
    ' *** CREA LA TABLA Y SUS CAMPOS DETALLE
    Set MyTable = db2.CreateTableDef(ArchivoSalida2)
    
    Call NuevoCampo(MyTable, "AA4_FILIAL", dbText, 4)
    Call NuevoCampo(MyTable, "AA4_CODCLI", dbText, 13)
    Call NuevoCampo(MyTable, "AA4_LOJA", dbText, 2)
    Call NuevoCampo(MyTable, "AA4_CODPRO", dbText, 15)
    Call NuevoCampo(MyTable, "AA4_NUMSER", dbText, 20)
    Call NuevoCampo(MyTable, "AA4_PRODAC", dbText, 15)
    Call NuevoCampo(MyTable, "AA4_DESCPR", dbText, 30)
    Call NuevoCampo(MyTable, "AA4_NSERAC", dbText, 20)
    Call NuevoCampo(MyTable, "AA4_DTGAR", dbDate)
    
    ' AGREGA TABLA A LA BD
    db2.TableDefs.Append MyTable
    On Error GoTo 0
    
    
    ' *** EXPORTACION DE DATOS
    
    Call ABRECONEXION
    
    ' *** CABECERA DE EQUIPOS ********************************
    
    Dim dbListaCabecera As ADODB.Recordset
    Dim NumeroCliente As Long
    Dim FechaX As Integer
    Dim NumeroFactura As String
    
    ' OBTIENE LA CABECERA DE EQUIPOS EN TABLA SQL
    On Error GoTo 98
    MensajeError = "Error al Leer Tabla de Cabecera de Equipos"
    Set dbListaCabecera = READSET("SELECT * FROM TANUMSE WITH(NOLOCK) WHERE NUMEROSERIE <> '' ORDER BY FECHALTA")
    FIN& = CantRegistros("TANUMSE", "NUMEROSERIE <> ''")
    On Error GoTo 0
    
    ' ABRE LA BD DBF EN MODO ESCRITURA
    On Error GoTo 98
    MensajeError = "Error al Leer Archivo DBF Para Escritura"
    Set RecordsetDBF = db1.OpenRecordset(ArchivoSalida1, dbOpenDynaset)
    On Error GoTo 0

    ' AGREGAR LOS REGISTROS DESDE LA TABLA SQL A LA TABLA DBF
    On Error GoTo 98
    MensajeError = ""
    I& = 1
    With dbListaCabecera
        Do While Not .EOF
        
            I& = I& + 1
            Call TRACE(24, I&, FIN&)
            
            RecordsetDBF.AddNew
            RecordsetDBF!AA3_FILIAL = ""
                NumeroCliente = XVALO(!NUMECLI)
            RecordsetDBF!AA3_CODCLI = CStr(NumeroCliente)
            RecordsetDBF!AA3_LOJA = ""
            RecordsetDBF!AA3_NOMCLI = Left$(RASOCLI$(NumeroCliente), 30)
            RecordsetDBF!AA3_CODPRO = Left$(SAFETEXT$(!Cod_Exte), 15)
            RecordsetDBF!AA3_DESPRO = Left$(SAFETEXT$(!Descrip), 30)
            RecordsetDBF!AA3_NUMSER = SAFETEXT$(!NUMEROSERIE)
                FechaX = CVINT(!FechEntrega)
            RecordsetDBF!AA3_DTVEND = CVDAT(FechaX)
            RecordsetDBF!AA3_DTINST = CVDAT(FechaX)
                FechaX = CVINT(!FEVENGAR)
            RecordsetDBF!AA3_DTGAR = CVDAT(FechaX)
                NumeroFactura = SAFETEXT$(!NUMFACTURA)
            RecordsetDBF!AA3_NFVEND = Mid$(NumeroFactura, 9, 8)
            RecordsetDBF!AA3_STATUS = ""
            RecordsetDBF.Update
            
            .MoveNext
        Loop
    End With
    
    On Error GoTo 0
    
    ' CIERRA LOS OBJETOS DE ACCESO A SQL Y AL ARCHIVO DBF
    On Error GoTo 98
    MensajeError = "Error al Cerrar Objetos"
    
    dbListaCabecera.Close
    Set dbListaCabecera = Nothing
    RecordsetDBF.Close
    Set RecordsetDBF = Nothing
    db1.Close
    Set db1 = Nothing
    
    On Error GoTo 0
    
    ' *** DETALLE DE EQUIPOS ********************************
    
    Dim dbListaDetalle As ADODB.Recordset
    Dim FechaGarantia As Integer
    Dim NumCliente As Long
    Dim CodigoEquipo As String
    Dim NumSerie As String
    Dim OrdenFabricacion As String
    
    ' OBTIENE LA CABECERA DE PRECIOS EN TABLA SQL
    On Error GoTo 98
    MensajeError = "Error al Leer Tabla de Detalle de Equipos"
    Set dbListaDetalle = READSET("SELECT NUMEROSERIE, NUMECLI, COD_EXTE, FEVENGAR FROM TANUMSE WITH(NOLOCK) WHERE NUMEROSERIE <> '' ORDER BY FECHALTA")
    FIN& = CantRegistros("TANUMSE", "NUMEROSERIE <> ''")
    On Error GoTo 0
    
    ' ABRE LA BD DBF EN MODO ESCRITURA
    On Error GoTo 98
    MensajeError = "Error al Leer Archivo DBF Para Escritura"
    Set RecordsetDBF = db2.OpenRecordset(ArchivoSalida2, dbOpenDynaset)
    On Error GoTo 0

    ' AGREGAR LOS REGISTROS DESDE LA TABLA SQL A LA TABLA DBF
    On Error GoTo 98
    MensajeError = ""
    I& = 1
    With dbListaDetalle
        Do While Not .EOF
        
            I& = I& + 1
            Call TRACE(24, I&, FIN&)
            
            NumSerie = SAFETEXT$(!NUMEROSERIE)
            OrdenFabricacion = SAFETEXT$(VALOBADA("MOVISTO", "DOPRILAR", "IDENLOTE='" & NumSerie & "' AND CODIMOVI='OF'"))
            
            If TRIM$(OrdenFabricacion) <> "" Then
                
                NumCliente = XVALO(!NUMECLI)
                CodigoEquipo = SAFETEXT$(!Cod_Exte)
                FechaGarantia = CVINT(!FEVENGAR)
            
                Dim dbComponentesTrazables As ADODB.Recordset
                Set dbComponentesTrazables = READSET("SELECT COD_EXTE, DESCRIP, IDENLOTE FROM MOVISTO WITH(NOLOCK) WHERE DOSECLAR='" & OrdenFabricacion & "' AND (CODIMOVI='TR' OR CODIMOVI='CF') AND DEPOMOVI=50 AND CANTSALID > 0")
                
                With dbComponentesTrazables
                    Do While Not .EOF
                        RecordsetDBF.AddNew
                        RecordsetDBF!AA4_FILIAL = ""
                        RecordsetDBF!AA4_CODCLI = NumCliente
                        RecordsetDBF!AA4_LOJA = ""
                        RecordsetDBF!AA4_CODPRO = CodigoEquipo
                        RecordsetDBF!AA4_NUMSER = NumSerie
                        RecordsetDBF!AA4_PRODAC = SAFETEXT$(!Cod_Exte)
                        RecordsetDBF!AA4_DESCPR = Left$(SAFETEXT$(!Descrip), 30)
                        RecordsetDBF!AA4_NSERAC = SAFETEXT$(!IDENLOTE)
                        RecordsetDBF!AA4_DTGAR = CVDAT(FechaGarantia)
                        RecordsetDBF.Update
                        
                        JJ& = JJ& + 1
                        .MoveNext
                    Loop
                End With
            
                On Error Resume Next
                dbComponentesTrazables.Close
                Set dbComponentesTrazables = Nothing
                On Error GoTo 0
            End If
            
            .MoveNext
        Loop
    End With
    
    On Error GoTo 0
    
    ' CIERRA LOS OBJETOS DE ACCESO A SQL Y AL ARCHIVO DBF
    On Error GoTo 98
    MensajeError = "Error al Cerrar Objetos"
    
    dbListaDetalle.Close
    Set dbListaDetalle = Nothing
    RecordsetDBF.Close
    Set RecordsetDBF = Nothing
    db2.Close
    Set db2 = Nothing
    
    On Error GoTo 0
    
    ' ***********************************
    
    Screen.MousePointer = 1
    Timer1.Enabled = False: lblTablaExportando = ""
    Call COMUNI("Exportación de EQUIPOS Finalizada")
    Exit Sub

    ' *** CAPTURA DE ERRORES
    
98  MensajeVisual = Err.Description
    If TRIM$(MensajeError) = "" Then MensajeError = MensajeVisual
    On Error GoTo 0
    
    On Error Resume Next
    dbListaCabecera.Close
    Set dbListaCabecera = Nothing
    RecordsetDBF.Close
    Set RecordsetDBF = Nothing
    db1.Close
    Set db1 = Nothing
    dbListaDetalle.Close
    Set dbListaDetalle = Nothing
    db2.Close
    Set db2 = Nothing
    On Error GoTo 0
    
    Screen.MousePointer = 1
    Timer1.Enabled = False: lblTablaExportando = ""
    Call COMUNI(MensajeError)

99  Screen.MousePointer = 1
    On Error GoTo 0
    
    On Error Resume Next
    dbListaCabecera.Close
    Set dbListaCabecera = Nothing
    RecordsetDBF.Close
    Set RecordsetDBF = Nothing
    db1.Close
    Set db1 = Nothing
    dbListaDetalle.Close
    Set dbListaDetalle = Nothing
    db2.Close
    Set db2 = Nothing
    On Error GoTo 0
    
    Timer1.Enabled = False: lblTablaExportando = ""
    Exit Sub

End Sub

Private Sub btnEstructuras_Click()

    Dim db As Database
    Dim RecordsetDBF As Recordset
    Dim MyTable As TableDef
    Dim MyField As Field
    Dim NombreBaseDatos As String
    Dim MensajeError As String
    Dim RutaArchivo As String
    Dim ArchivoSalida As String
    Dim RutaUsuario As String
    
    ArchivoSalida = "SG1"
    
    CommonDialog1.CancelError = True ' SELECCION DE ARCHIVO A IMPORTAR Y GRABACION
    CommonDialog1.DialogTitle = "Exportacion de Datos"
    CommonDialog1.FileName = ArchivoSalida & ".DBF"
    CommonDialog1.Filter = "Archivos DBF(*.DBF)|*.DBF|"
    CommonDialog1.FilterIndex = 1
    On Error GoTo 99
    CommonDialog1.ShowOpen
    On Error GoTo 0
    RutaUsuario = TRIM$(CommonDialog1.FileName)
    If RutaUsuario = "" Or InStr(1, RutaUsuario, ArchivoSalida & ".DBF", vbTextCompare) < 1 Then
        MensajeError = "Falta Carpeta de Archivo o Se Ha Modificado el Nombre del Archivo"
        GoTo 98
    End If
    
    RutaArchivo = Mid(RutaUsuario, 1, Len(RutaUsuario) - Len("\" & ArchivoSalida & ".DBF"))
    
    ' *** ELIMINA EL ARCHIVO EXISTENTE
    On Error Resume Next
    Kill RutaArchivo & "\" & ArchivoSalida & ".DBF"
    If Err > 0 And Err <> 53 Then
        MensajeError = "Se Produjo un Error al Eliminar el Archivo '" & REPLACAR$(RutaArchivo, "\", "/") & "/" & ArchivoSalida & ".DBF' .\El Proceso No Continuará."
        GoTo 98
    End If
    On Error GoTo 0
    
    Screen.MousePointer = 11
    
    lblTablaExportando = "Exportando Estructuras de Producto"
    Timer1.Enabled = True
    
    ' *** CREAR EL ARCHIVO DBF
    On Error GoTo 98
    MensajeError = ""
    Set db = Workspaces(0).OpenDatabase(RutaArchivo, False, False, "dBASE IV;DATABASE=" & ArchivoSalida & ".DBF")
    
    ' *** CREA LA TABLA Y SUS CAMPOS
    Set MyTable = db.CreateTableDef(ArchivoSalida)
    
    Call NuevoCampo(MyTable, "G1_FILIAL", dbText, 4)
    Call NuevoCampo(MyTable, "G1_COD", dbText, 15)
    Call NuevoCampo(MyTable, "G1_COMP", dbText, 15)
    Call NuevoCampo(MyTable, "G1_DESC", dbText, 30)
    Call NuevoCampo(MyTable, "G1_QUANT", dbDouble, 8)
    Call NuevoCampo(MyTable, "G1_INI", dbDate, 8)
    Call NuevoCampo(MyTable, "G1_FIM", dbDate, 8)
    Call NuevoCampo(MyTable, "G1_FIXVAR", dbText, 1)
    Call NuevoCampo(MyTable, "G1_VLCOMPE", dbText, 1)
    
    ' AGREGA TABLA A LA BD
    db.TableDefs.Append MyTable
    On Error GoTo 0
    
    ' *** EXPORTACION DE DATOS
    
    Call ABRECONEXION
    Dim dbEstructuraArticulos As ADODB.Recordset
    Dim FechaRevisionFormula As Integer
    
    ' OBTIENE LOS ARTICULOS EN TABLA SQL
    On Error GoTo 98
    MensajeError = "Error al Leer Tabla de Articulos"
    Set dbEstructuraArticulos = READSET("SELECT * FROM FORMULAS WITH(NOLOCK) WHERE CODICONJ > 0 ORDER BY CODICONJ, CODIELEM ASC")
    FIN& = CantRegistros("FORMULAS", "CODICONJ > 0")
    On Error GoTo 0
    
    ' ABRE LA BD DBF EN MODO ESCRITURA
    On Error GoTo 98
    MensajeError = "Error al Leer Archivo DBF Para Escritura"
    Set RecordsetDBF = db.OpenRecordset(ArchivoSalida, dbOpenDynaset)
    On Error GoTo 0

    ' AGREGAR LOS REGISTROS DESDE LA TABLA SQL A LA TABLA DBF
    On Error GoTo 98
    MensajeError = ""
    I& = 1
    With dbEstructuraArticulos
        Do While Not .EOF
        
            I& = I& + 1
            Call TRACE(24, I&, FIN&)
            
            RecordsetDBF.AddNew
            RecordsetDBF!G1_FILIAL = ""
            RecordsetDBF!G1_COD = Left$(SAFETEXT$(!CODXCONJ), 15)
            RecordsetDBF!G1_COMP = Left$(SAFETEXT$(!CODXELEM), 15)
            RecordsetDBF!G1_DESC = Left$(DESCRIT0$(!CODIELEM), 30)
            RecordsetDBF!G1_QUANT = XVALO(!USOBRUTO)
                FechaRevisionFormula = CVINT(VALOBADA("MASTER", "Revi_Formu", "CODINT=" & CStr(!CODICONJ)))
            RecordsetDBF!G1_INI = CVDAT(FechaRevisionFormula)
            RecordsetDBF!G1_FIM = CVDAT(FechaRevisionFormula)
            RecordsetDBF!G1_FIXVAR = ""
            RecordsetDBF!G1_VLCOMPE = ""
            RecordsetDBF.Update
            
            .MoveNext
        Loop
    End With
    
    On Error GoTo 0
    
    ' CIERRA LOS OBJETOS DE ACCESO A SQL Y AL ARCHIVO DBF
    On Error GoTo 98
    MensajeError = "Error al Cerrar Objetos"
    
    dbEstructuraArticulos.Close
    Set dbEstructuraArticulos = Nothing
    RecordsetDBF.Close
    Set RecordsetDBF = Nothing
    db.Close
    Set db = Nothing
    
    On Error GoTo 0
    Screen.MousePointer = 1
    Timer1.Enabled = False: lblTablaExportando = ""
    Call COMUNI("Exportación de ESTRUCTURAS Finalizada")
    Exit Sub

    ' *** CAPTURA DE ERRORES
    
98  MensajeVisual = Err.Description
    If TRIM$(MensajeError) = "" Then MensajeError = MensajeVisual
    On Error GoTo 0
    
    On Error Resume Next
    dbEstructuraArticulos.Close
    Set dbEstructuraArticulos = Nothing
    RecordsetDBF.Close
    Set RecordsetDBF = Nothing
    db.Close
    Set db = Nothing
    On Error GoTo 0
    
    Screen.MousePointer = 1
    Timer1.Enabled = False: lblTablaExportando = ""
    Call COMUNI(MensajeError)

99  Screen.MousePointer = 1
    On Error GoTo 0
    
    On Error Resume Next
    dbEstructuraArticulos.Close
    Set dbEstructuraArticulos = Nothing
    RecordsetDBF.Close
    Set RecordsetDBF = Nothing
    db.Close
    Set db = Nothing
    On Error GoTo 0
    
    Timer1.Enabled = False: lblTablaExportando = ""
    Exit Sub
End Sub

Private Sub btnListasDePrecios_Click()

    Dim db1 As Database
    Dim db2 As Database
    Dim RecordsetDBF As Recordset
    Dim MyTable As TableDef
    Dim MyField As Field
    Dim NombreBaseDatos As String
    Dim MensajeError As String
    Dim RutaArchivo As String
    Dim ArchivoSalida1 As String
    Dim ArchivoSalida2 As String
    Dim RutaUsuario As String
    
    ArchivoSalida1 = "DA0"
    
    CommonDialog1.CancelError = True ' SELECCION DE ARCHIVO A IMPORTAR Y GRABACION
    CommonDialog1.DialogTitle = "Exportacion de Datos"
    CommonDialog1.FileName = ArchivoSalida1 & ".DBF"
    CommonDialog1.Filter = "Archivos DBF(*.DBF)|*.DBF|"
    CommonDialog1.FilterIndex = 1
    On Error GoTo 99
    CommonDialog1.ShowOpen
    On Error GoTo 0
    RutaUsuario = TRIM$(CommonDialog1.FileName)
    If RutaUsuario = "" Or InStr(1, RutaUsuario, ArchivoSalida1 & ".DBF", vbTextCompare) < 1 Then
        MensajeError = "Falta Carpeta de Archivo o Se Ha Modificado el Nombre del Archivo"
        GoTo 98
    End If
    
    RutaArchivo = Mid(RutaUsuario, 1, Len(RutaUsuario) - Len("\" & ArchivoSalida1 & ".DBF"))
    
    ArchivoSalida2 = "DA1"
    
    ' *** ELIMINA EL ARCHIVO EXISTENTE
    On Error Resume Next
    Kill RutaArchivo & "\" & ArchivoSalida1 & ".DBF"
    If Err > 0 And Err <> 53 Then
        MensajeError = "Se Produjo un Error al Eliminar el Archivo '" & REPLACAR$(RutaArchivo, "\", "/") & "/" & ArchivoSalida1 & ".DBF' .\El Proceso No Continuará."
        GoTo 98
    End If
    
    Kill RutaArchivo & "\" & ArchivoSalida2 & ".DBF"
    If Err > 0 And Err <> 53 Then
        MensajeError = "Se Produjo un Error al Eliminar el Archivo '" & REPLACAR$(RutaArchivo, "\", "/") & "/" & ArchivoSalida2 & ".DBF' .\El Proceso No Continuará."
        GoTo 98
    End If
    On Error GoTo 0
    
    Screen.MousePointer = 11
    
    lblTablaExportando = "Exportando Listas de Precios"
    Timer1.Enabled = True
    
    ' *** CREAR EL ARCHIVO DBF
    On Error GoTo 98
    MensajeError = ""
    Set db1 = Workspaces(0).OpenDatabase(RutaArchivo, False, False, "dBASE IV;DATABASE=" & ArchivoSalida1 & ".DBF")
    
    ' *** CREA LA TABLA Y SUS CAMPOS CABECERA
    Set MyTable = db1.CreateTableDef(ArchivoSalida1)
    
    Call NuevoCampo(MyTable, "DA0_CODTAB", dbText, 2)
    Call NuevoCampo(MyTable, "DA0_DESCRI", dbText, 30)
    Call NuevoCampo(MyTable, "DA0_DATDE", dbDate, 8)
    Call NuevoCampo(MyTable, "DA0_HORADE", dbText, 5)
    Call NuevoCampo(MyTable, "DA0_DATATE", dbDate, 8)
    Call NuevoCampo(MyTable, "DA0_HORATE", dbText, 5)
    Call NuevoCampo(MyTable, "DA0_ATIVO", dbText, 1)
    Call NuevoCampo(MyTable, "DA0_FILIAL", dbText, 4)
    
    ' AGREGA TABLA A LA BD
    db1.TableDefs.Append MyTable
    On Error GoTo 0
    
    On Error GoTo 98
    MensajeError = ""
    Set db2 = Workspaces(0).OpenDatabase(RutaArchivo, False, False, "dBASE IV;DATABASE=" & ArchivoSalida2 & ".DBF")
    
    ' *** CREA LA TABLA Y SUS CAMPOS DETALLE
    Set MyTable = db2.CreateTableDef(ArchivoSalida2)
    
    Call NuevoCampo(MyTable, "DA1_FILIAL", dbText, 4)
    Call NuevoCampo(MyTable, "DA1_ITEM", dbText, 4)
    Call NuevoCampo(MyTable, "DA1_CODTAB", dbText, 3)
    Call NuevoCampo(MyTable, "DA1_DESTAB", dbText, 30)
    Call NuevoCampo(MyTable, "DA1_CODPRO", dbText, 27)
    Call NuevoCampo(MyTable, "DA1_DESCRI", dbText, 30)
    Call NuevoCampo(MyTable, "DA1_PRCVEN", dbDouble)
    Call NuevoCampo(MyTable, "DA1_ATIVO", dbText, 1)
    Call NuevoCampo(MyTable, "DA1_TPOPER", dbText, 1)
    Call NuevoCampo(MyTable, "DA1_QTDLOT", dbDouble)
    Call NuevoCampo(MyTable, "DA1_INDLOT", dbText, 20)
    Call NuevoCampo(MyTable, "DA1_MOEDA", dbInteger)
    Call NuevoCampo(MyTable, "DA1_DATVIG", dbDate, 8)
    
    ' AGREGA TABLA A LA BD
    db2.TableDefs.Append MyTable
    On Error GoTo 0
    
    
    ' *** EXPORTACION DE DATOS
    
    Call ABRECONEXION
    
    ' *** CABECERA DE LISTA DE PRECIOS ********************************
    
    Dim dbListaCabecera As ADODB.Recordset
    Dim FechaVigencia As Integer
    
    ' OBTIENE LA CABECERA DE PRECIOS EN TABLA SQL
    On Error GoTo 98
    MensajeError = "Error al Leer Tabla de Cabecera de Precios"
    Set dbListaCabecera = READSET("SELECT * FROM LISTAPRE WITH(NOLOCK) WHERE COD_LISTA > 0")
    FIN& = CantRegistros("LISTAPRE", "COD_LISTA > 0")
    On Error GoTo 0
    
    ' ABRE LA BD DBF EN MODO ESCRITURA
    On Error GoTo 98
    MensajeError = "Error al Leer Archivo DBF Para Escritura"
    Set RecordsetDBF = db1.OpenRecordset(ArchivoSalida1, dbOpenDynaset)
    On Error GoTo 0

    ' AGREGAR LOS REGISTROS DESDE LA TABLA SQL A LA TABLA DBF
    On Error GoTo 98
    MensajeError = ""
    I& = 1
    With dbListaCabecera
        Do While Not .EOF
        
            I& = I& + 1
            Call TRACE(24, I&, FIN&)
            
            RecordsetDBF.AddNew
            RecordsetDBF!DA0_CODTAB = SAFETEXT$(!COD_LISTA)
            RecordsetDBF!DA0_DESCRI = Left$(SAFETEXT$(!DENOM_LISTA), 30)
                FechaVigencia = CVINT(!VIGENCIA)
            RecordsetDBF!DA0_DATDE = CVDAT(FechaVigencia)
            RecordsetDBF!DA0_HORADE = ""
            RecordsetDBF!DA0_DATATE = CVDAT(FechaVigencia)
            RecordsetDBF!DA0_HORATE = ""
            RecordsetDBF!DA0_ATIVO = "1"
            RecordsetDBF!DA0_FILIAL = ""
            RecordsetDBF.Update
            
            .MoveNext
        Loop
    End With
    
    On Error GoTo 0
    
    ' CIERRA LOS OBJETOS DE ACCESO A SQL Y AL ARCHIVO DBF
    On Error GoTo 98
    MensajeError = "Error al Cerrar Objetos"
    
    dbListaCabecera.Close
    Set dbListaCabecera = Nothing
    RecordsetDBF.Close
    Set RecordsetDBF = Nothing
    db1.Close
    Set db1 = Nothing
    
    On Error GoTo 0
    
    ' *** DETALLE DE LISTA DE PRECIOS ********************************
    
    Dim dbListaDetalle As ADODB.Recordset
    Dim DenominacionLista As String
    Dim CodigoArticulo As Integer
    Dim Activo As String
    Dim EstadoItem As Integer
    Dim CodigoMoneda As Integer
    
    ' OBTIENE LA CABECERA DE PRECIOS EN TABLA SQL
    On Error GoTo 98
    MensajeError = "Error al Leer Tabla de Detalle de Precios"
    Set dbListaDetalle = READSET("SELECT * FROM PREVENTA WITH(NOLOCK) WHERE COD_LISTA > 0 ORDER BY COD_LISTA")
    FIN& = CantRegistros("PREVENTA", "COD_LISTA > 0")
    On Error GoTo 0
    
    ' ABRE LA BD DBF EN MODO ESCRITURA
    On Error GoTo 98
    MensajeError = "Error al Leer Archivo DBF Para Escritura"
    Set RecordsetDBF = db2.OpenRecordset(ArchivoSalida2, dbOpenDynaset)
    On Error GoTo 0

    ' AGREGAR LOS REGISTROS DESDE LA TABLA SQL A LA TABLA DBF
    On Error GoTo 98
    MensajeError = ""
    I& = 1
    With dbListaDetalle
        Do While Not .EOF
        
            I& = I& + 1
            Call TRACE(24, I&, FIN&)
            
            RecordsetDBF.AddNew
            RecordsetDBF!DA1_FILIAL = ""
            RecordsetDBF!DA1_ITEM = ""
            RecordsetDBF!DA1_CODTAB = CStr(XVALO(!COD_LISTA))
                DenominacionLista = SAFETEXT$(VALOBADA("LISTAPRE", "DENOM_LISTA", "COD_LISTA=" & SAFETEXT$(!COD_LISTA)))
            RecordsetDBF!DA1_DESTAB = Left$(DenominacionLista, 30)
                CodigoArticulo = XVALO(!COD_PIEZA)
            RecordsetDBF!DA1_CODPRO = Left$(CODEXT$(CodigoArticulo), 27)
            RecordsetDBF!DA1_DESCRI = Left$(DESCRIT0$(CodigoArticulo), 30)
            RecordsetDBF!DA1_PRCVEN = XVALO(!PRECIO_UNITARIO)
                Activo = "1"
                EstadoItem = XVALO(VALOBADA("MASTER", "STATUS", "CODINT=" & CStr(CodigoArticulo)))
                If EstadoItem < 0 Then Activo = "2"
            RecordsetDBF!DA1_ATIVO = Activo
            RecordsetDBF!DA1_TPOPER = ""
            RecordsetDBF!DA1_QTDLOT = 0
            RecordsetDBF!DA1_INDLOT = ""
                CodigoMoneda = XVALO(VALOBADA("LISTAPRE", "MONEDA", "COD_LISTA=" & CStr(XVALO(!COD_LISTA))))
            RecordsetDBF!DA1_MOEDA = CodigoMoneda
                FechaVigencia = CVINT(VALOBADA("LISTAPRE", "VIGENCIA", "COD_LISTA=" & CStr(XVALO(!COD_LISTA))))
            RecordsetDBF!DA1_DATVIG = CVDAT(FechaVigencia)
            RecordsetDBF.Update
            
            .MoveNext
        Loop
    End With
    
    On Error GoTo 0
    
    ' CIERRA LOS OBJETOS DE ACCESO A SQL Y AL ARCHIVO DBF
    On Error GoTo 98
    MensajeError = "Error al Cerrar Objetos"
    
    dbListaDetalle.Close
    Set dbListaDetalle = Nothing
    RecordsetDBF.Close
    Set RecordsetDBF = Nothing
    db2.Close
    Set db2 = Nothing
    
    On Error GoTo 0
    
    ' ***********************************
    
    Screen.MousePointer = 1
    Timer1.Enabled = False: lblTablaExportando = ""
    Call COMUNI("Exportación de LISTAS DE PRECIOS Finalizada")
    Exit Sub

    ' *** CAPTURA DE ERRORES
    
98  MensajeVisual = Err.Description
    If TRIM$(MensajeError) = "" Then MensajeError = MensajeVisual
    On Error GoTo 0
    
    On Error Resume Next
    dbListaCabecera.Close
    Set dbListaCabecera = Nothing
    RecordsetDBF.Close
    Set RecordsetDBF = Nothing
    db1.Close
    Set db1 = Nothing
    dbListaDetalle.Close
    Set dbListaDetalle = Nothing
    db2.Close
    Set db2 = Nothing
    On Error GoTo 0
    
    Screen.MousePointer = 1
    Timer1.Enabled = False: lblTablaExportando = ""
    Call COMUNI(MensajeError)

99  Screen.MousePointer = 1
    On Error GoTo 0
    
    On Error Resume Next
    dbListaCabecera.Close
    Set dbListaCabecera = Nothing
    RecordsetDBF.Close
    Set RecordsetDBF = Nothing
    db1.Close
    Set db1 = Nothing
    dbListaDetalle.Close
    Set dbListaDetalle = Nothing
    db2.Close
    Set db2 = Nothing
    On Error GoTo 0
    
    Timer1.Enabled = False: lblTablaExportando = ""
    Exit Sub
    
End Sub

Private Sub btnProductos_Click()

    Dim db As Database
    Dim RecordsetDBF As Recordset
    Dim MyTable As TableDef
    Dim MyField As Field
    Dim NombreBaseDatos As String
    Dim MensajeError As String
    Dim RutaArchivo As String
    Dim ArchivoSalida As String
    Dim RutaUsuario As String
    
    ArchivoSalida = "SB1"
    
    CommonDialog1.CancelError = True ' SELECCION DE ARCHIVO A IMPORTAR Y GRABACION
    CommonDialog1.DialogTitle = "Exportacion de Datos"
    CommonDialog1.FileName = ArchivoSalida & ".DBF"
    CommonDialog1.Filter = "Archivos DBF(*.DBF)|*.DBF|"
    CommonDialog1.FilterIndex = 1
    On Error GoTo 99
    CommonDialog1.ShowOpen
    On Error GoTo 0
    RutaUsuario = TRIM$(CommonDialog1.FileName)
    If RutaUsuario = "" Or InStr(1, RutaUsuario, ArchivoSalida & ".DBF", vbTextCompare) < 1 Then
        MensajeError = "Falta Carpeta de Archivo o Se Ha Modificado el Nombre del Archivo"
        GoTo 98
    End If
    
    RutaArchivo = Mid(RutaUsuario, 1, Len(RutaUsuario) - Len("\" & ArchivoSalida & ".DBF"))
    
    ' *** ELIMINA EL ARCHIVO EXISTENTE
    On Error Resume Next
    Kill RutaArchivo & "\" & ArchivoSalida & ".DBF"
    If Err > 0 And Err <> 53 Then
        MensajeError = "Se Produjo un Error al Eliminar el Archivo '" & REPLACAR$(RutaArchivo, "\", "/") & "/" & ArchivoSalida & ".DBF' .\El Proceso No Continuará."
        GoTo 98
    End If
    On Error GoTo 0
    
    Screen.MousePointer = 11
    
    lblTablaExportando = "Exportando Productos"
    Timer1.Enabled = True
    
    ' *** CREAR EL ARCHIVO DBF
    On Error GoTo 98
    MensajeError = ""
    Set db = Workspaces(0).OpenDatabase(RutaArchivo, False, False, "dBASE IV;DATABASE=" & ArchivoSalida & ".DBF")
    
    ' *** CREA LA TABLA Y SUS CAMPOS
    Set MyTable = db.CreateTableDef(ArchivoSalida)
    
    Call NuevoCampo(MyTable, "B1_FILIAL", dbText, 4)
    Call NuevoCampo(MyTable, "B1_COD", dbText, 15)
    Call NuevoCampo(MyTable, "B1_DESC", dbText, 30)
    Call NuevoCampo(MyTable, "B1_TIPO", dbText, 2)
    Call NuevoCampo(MyTable, "B1_CODITE", dbText, 27)
    Call NuevoCampo(MyTable, "B1_UM", dbText, 2)
    Call NuevoCampo(MyTable, "B1_LOCPAD", dbText, 2)
    Call NuevoCampo(MyTable, "B1_TE", dbText, 3)
    Call NuevoCampo(MyTable, "B1_TS", dbText, 3)
    Call NuevoCampo(MyTable, "B1_TIPCONV", dbText, 1)
    Call NuevoCampo(MyTable, "B1_PRV1", dbDouble, 8)
    Call NuevoCampo(MyTable, "B1_CUSTD", dbDouble, 8)
    Call NuevoCampo(MyTable, "B1_ESTSEG", dbDouble, 8)
    Call NuevoCampo(MyTable, "B1_PE", dbInteger, 8)
    Call NuevoCampo(MyTable, "B1_TIPE", dbText, 1)
    Call NuevoCampo(MyTable, "B1_CONTA", dbText, 20)
    Call NuevoCampo(MyTable, "B1_CC", dbText, 9)
    Call NuevoCampo(MyTable, "B1_APROPRI", dbText, 1)
    Call NuevoCampo(MyTable, "B1_TIPODEC", dbText, 1)
    Call NuevoCampo(MyTable, "B1_FANTASM", dbText, 1)
    Call NuevoCampo(MyTable, "B1_RASTRO", dbText, 1)
    Call NuevoCampo(MyTable, "B1_UREV", dbDate, 8)
    Call NuevoCampo(MyTable, "B1_DATREF", dbDate, 8)
    Call NuevoCampo(MyTable, "B1_CODBAR", dbText, 15)
    Call NuevoCampo(MyTable, "B1_LOCALIZ", dbText, 1)
    Call NuevoCampo(MyTable, "B1_CONTRAT", dbText, 1)
    Call NuevoCampo(MyTable, "B1_IMPORT", dbText, 1)
    Call NuevoCampo(MyTable, "B1_ANUENTE", dbText, 1)
    Call NuevoCampo(MyTable, "B1_TIPOCQ", dbText, 1)
    Call NuevoCampo(MyTable, "B1_SOLICIT", dbText, 1)
    Call NuevoCampo(MyTable, "B1_AGREGCU", dbText, 1)
    Call NuevoCampo(MyTable, "B1_DESPIMP", dbText, 1)
    Call NuevoCampo(MyTable, "B1_INSS", dbText, 1)
    Call NuevoCampo(MyTable, "B1_CONCGAN", dbText, 2)
    Call NuevoCampo(MyTable, "B1_CONCIVA", dbText, 3)
    Call NuevoCampo(MyTable, "B1_FLAGSUG", dbText, 1)
    Call NuevoCampo(MyTable, "B1_CLASSVE", dbText, 1)
    Call NuevoCampo(MyTable, "B1_MIDIA", dbText, 1)
    Call NuevoCampo(MyTable, "B1_QTMIDIA", dbDouble, 8)
    Call NuevoCampo(MyTable, "B1_ATIVO", dbText, 1)
    Call NuevoCampo(MyTable, "B1_CPOTENC", dbText, 1)
    Call NuevoCampo(MyTable, "B1_NALCIVA", dbText, 1)
    Call NuevoCampo(MyTable, "B1_USAFEFO", dbText, 1)
    Call NuevoCampo(MyTable, "B1_PIS", dbText, 1)
    Call NuevoCampo(MyTable, "B1_MSBLQL", dbText, 1)
    Call NuevoCampo(MyTable, "B1_PRODSBP", dbText, 1)
    Call NuevoCampo(MyTable, "B1_COFINS", dbText, 1)
    Call NuevoCampo(MyTable, "B1_CSLL", dbText, 1)
    Call NuevoCampo(MyTable, "B1_GARANT", dbText, 1)
    Call NuevoCampo(MyTable, "B1_ULTSERI", dbText, 12)
    Call NuevoCampo(MyTable, "B1_TRAZAB", dbText, 1)
    Call NuevoCampo(MyTable, "B1_XFICHA", dbText, 64)
    Call NuevoCampo(MyTable, "B1_DIRFCHA", dbText, 250)
    
    ' AGREGA TABLA A LA BD
    db.TableDefs.Append MyTable
    On Error GoTo 0
    
    ' *** EXPORTACION DE DATOS
    
    Call ABRECONEXION
    Dim dbArticulos As ADODB.Recordset
    Dim CuentaVentasRespInscripto As String
    Dim EstadoActivo As String
    Dim Garantia As String
    Dim UltimoNumeroSerie As String
    Dim Traza As String
    Dim FechaX As Integer
    Dim CodigoArticulo As String
    
    ' OBTIENE LOS ARTICULOS EN TABLA SQL
    On Error GoTo 98
    MensajeError = "Error al Leer Tabla de Articulos"
    Set dbArticulos = READSET("SELECT * FROM MASTER WITH(NOLOCK) WHERE CODINT > 0")
    FIN& = CantRegistros("MASTER", "Codint>0")
    On Error GoTo 0
    
    ' ABRE LA BD DBF EN MODO ESCRITURA
    On Error GoTo 98
    MensajeError = "Error al Leer Archivo DBF Para Escritura"
    Set RecordsetDBF = db.OpenRecordset(ArchivoSalida, dbOpenDynaset)
    On Error GoTo 0

    ' AGREGAR LOS REGISTROS DESDE LA TABLA SQL A LA TABLA DBF
    On Error GoTo 98
    MensajeError = ""
    I& = 1
    CuentaVentasRespInscripto = TRIM$(CONTROL$("SVD451S", "CUEVE-1"))
    With dbArticulos
        Do While Not .EOF
        
            I& = I& + 1
            Call TRACE(24, I&, FIN&)
            
            RecordsetDBF.AddNew
            RecordsetDBF!B1_FILIAL = ""
                CodigoArticulo = Left$(SAFETEXT$(!CODIGO), 15)
            RecordsetDBF!B1_COD = CodigoArticulo
            RecordsetDBF!B1_DESC = Left$(SAFETEXT$(!Descripcion), 30)
            RecordsetDBF!B1_TIPO = Left$(SAFETEXT$(!TIMATE), 2)
            RecordsetDBF!B1_CODITE = Left$(SAFETEXT$(!CODIGO), 15)
            RecordsetDBF!B1_UM = Left$(SAFETEXT$(!UMEDIDA), 2)
            RecordsetDBF!B1_LOCPAD = Left$(SAFETEXT$(!DEPOENTRA), 2)
            RecordsetDBF!B1_TE = ""
            RecordsetDBF!B1_TS = ""
            RecordsetDBF!B1_TIPCONV = ""
            RecordsetDBF!B1_PRV1 = PRELISTA("LIPRE001", XVALO(!CODINT))
            RecordsetDBF!B1_CUSTD = XVALO(!COSUNI)
            RecordsetDBF!B1_ESTSEG = XVALO(!STOMIN)
            RecordsetDBF!B1_PE = 0
            RecordsetDBF!B1_TIPE = "" ' "D"
            RecordsetDBF!B1_CONTA = CuentaVentasRespInscripto
            RecordsetDBF!B1_CC = ""
            RecordsetDBF!B1_APROPRI = ""
            RecordsetDBF!B1_TIPODEC = ""
            RecordsetDBF!B1_FANTASM = ""
                FechaX = CVINT(!FEREVIS)
            RecordsetDBF!B1_UREV = CVDAT(FechaX)
                FechaX = CVINT(!FECOSTO)
            RecordsetDBF!B1_DATREF = CVDAT(FechaX)
            RecordsetDBF!B1_CODBAR = ""
            RecordsetDBF!B1_LOCALIZ = ""
            RecordsetDBF!B1_CONTRAT = ""
            RecordsetDBF!B1_IMPORT = ""
            RecordsetDBF!B1_ANUENTE = ""
            RecordsetDBF!B1_TIPOCQ = ""
            RecordsetDBF!B1_SOLICIT = ""
            RecordsetDBF!B1_AGREGCU = "" ' "2"
            RecordsetDBF!B1_DESPIMP = "" ' "N"
            RecordsetDBF!B1_INSS = "" ' "N"
            RecordsetDBF!B1_CONCGAN = ""
            RecordsetDBF!B1_CONCIVA = "" ' "2"   ' NO
            RecordsetDBF!B1_FLAGSUG = "" ' "1"   ' NORMAL
            RecordsetDBF!B1_CLASSVE = "" ' "1"   ' NORMAL
            RecordsetDBF!B1_MIDIA = "" ' "2"     ' NO
            RecordsetDBF!B1_QTMIDIA = "0"
                EstadoActivo = "S"
                If XVALO(!Status) < 0 Then EstadoActivo = "N"
            RecordsetDBF!B1_ATIVO = EstadoActivo
            RecordsetDBF!B1_CPOTENC = "" ' "2"   ' NO
            RecordsetDBF!B1_NALCIVA = "" ' "2"   ' NO
            RecordsetDBF!B1_USAFEFO = "" ' "1"   ' SI
            RecordsetDBF!B1_PIS = "" ' "2"       ' NO
            RecordsetDBF!B1_MSBLQL = "" ' "2"    ' NO
            RecordsetDBF!B1_PRODSBP = "" ' "P"   ' NO
            RecordsetDBF!B1_COFINS = "" ' "2"    ' NO
            RecordsetDBF!B1_CSLL = "" ' "2"      ' NO
                Garantia = "S"
                If XVALO(!Trazable) <> 2 Then Garantia = "N"
            RecordsetDBF!B1_GARANT = Garantia
                UltimoNumeroSerie = ""
                If XVALO(!Trazable) = 2 Then
                    UltimoNumeroSerie = SAFETEXT$(VALOBADA("TANUMSE", "NUMEROSERIE", "COD_INTE='" & CStr(XVALO(!CODINT)) & "' ORDER BY FECHALTA DESC, NUMEROSERIE ASC "))
                End If
            RecordsetDBF!B1_ULTSERI = Left$(UltimoNumeroSerie, 12)
                Traza = "N"
                If XVALO(!Trazable) > 0 Then Traza = "S"
            RecordsetDBF!B1_TRAZAB = Traza
                ArchivoTecnico = SAFETEXT$(!ARCHIPDF)
                If TRIM$(ArchivoTecnico) <> "" Then
                    For IX = Len(ArchivoTecnico) To 1 Step -1
                        If Mid$(ArchivoTecnico, IX, 1) = "\" Then
                            ArchivoTecnico = Mid$(ArchivoTecnico, IX + 1)
                            Exit For
                        End If
                    Next IX
                End If
            RecordsetDBF!B1_XFICHA = Left$(ArchivoTecnico, 64)
            RecordsetDBF!B1_DIRFCHA = Left$(SAFETEXT$(!ESPECIFICACION), 250)
            RecordsetDBF.Update
            
            .MoveNext
        Loop
    End With
    
    On Error GoTo 0
    
    ' CIERRA LOS OBJETOS DE ACCESO A SQL Y AL ARCHIVO DBF
    On Error GoTo 98
    MensajeError = "Error al Cerrar Objetos"
    
    dbArticulos.Close
    Set dbArticulos = Nothing
    RecordsetDBF.Close
    Set RecordsetDBF = Nothing
    db.Close
    Set db = Nothing
    
    On Error GoTo 0
    Screen.MousePointer = 1
    Timer1.Enabled = False: lblTablaExportando = ""

    Call COMUNI("Exportación de PRODUCTOS Finalizada")
    Exit Sub

    ' *** CAPTURA DE ERRORES
    
98  MensajeVisual = Err.Description
    If TRIM$(MensajeError) = "" Then MensajeError = "(" & CodigoArticulo & ") - " & MensajeVisual
    On Error GoTo 0
    
    On Error Resume Next
    dbArticulos.Close
    Set dbArticulos = Nothing
    RecordsetDBF.Close
    Set RecordsetDBF = Nothing
    db.Close
    Set db = Nothing
    On Error GoTo 0
    
    Screen.MousePointer = 1
    Timer1.Enabled = False: lblTablaExportando = ""
    Call COMUNI(MensajeError)

99  Screen.MousePointer = 1
    On Error GoTo 0
    
    On Error Resume Next
    dbArticulos.Close
    Set dbArticulos = Nothing
    RecordsetDBF.Close
    Set RecordsetDBF = Nothing
    db.Close
    Set db = Nothing
    On Error GoTo 0
    
    Timer1.Enabled = False: lblTablaExportando = ""
    Exit Sub

End Sub

Private Sub btnProveedores_Click()

    Dim db As Database
    Dim RecordsetDBF As Recordset
    Dim MyTable As TableDef
    Dim MyField As Field
    Dim NombreBaseDatos As String
    Dim MensajeError As String
    Dim RutaArchivo As String
    Dim ArchivoSalida As String
    Dim RutaUsuario As String
    
    ArchivoSalida = "SA2"
    
    CommonDialog1.CancelError = True ' SELECCION DE ARCHIVO A IMPORTAR Y GRABACION
    CommonDialog1.DialogTitle = "Exportacion de Datos"
    CommonDialog1.FileName = ArchivoSalida & ".DBF"
    CommonDialog1.Filter = "Archivos DBF(*.DBF)|*.DBF|"
    CommonDialog1.FilterIndex = 1
    On Error GoTo 99
    CommonDialog1.ShowOpen
    On Error GoTo 0
    RutaUsuario = TRIM$(CommonDialog1.FileName)
    If RutaUsuario = "" Or InStr(1, RutaUsuario, ArchivoSalida & ".DBF", vbTextCompare) < 1 Then
        MensajeError = "Falta Carpeta de Archivo o Se Ha Modificado el Nombre del Archivo"
        GoTo 98
    End If
    
    RutaArchivo = Mid(RutaUsuario, 1, Len(RutaUsuario) - Len("\" & ArchivoSalida & ".DBF"))
    
    ' ELIMINA EL ARCHIVO EXISTENTE
    On Error Resume Next
    Kill RutaArchivo & "\" & ArchivoSalida & ".DBF"
    If Err > 0 And Err <> 53 Then
        MensajeError = "Se Produjo un Error al Eliminar el Archivo '" & REPLACAR$(RutaArchivo, "\", "/") & "/" & ArchivoSalida & ".DBF' .\El Proceso No Continuará."
        GoTo 98
    End If
    On Error GoTo 0
    
    Screen.MousePointer = 11
    
    lblTablaExportando = "Exportando Proveedores"
    Timer1.Enabled = True
    
    On Error GoTo 98
    MensajeError = "Error al Crear Archivo DBF"
    ' CREAR EL ARCHIVO DBF
    Set db = Workspaces(0).OpenDatabase(RutaArchivo, False, False, "dBASE IV;DATABASE=" & ArchivoSalida & ".DBF")
    
    ' CREA LA TABLA Y SUS CAMPOS
    Set MyTable = db.CreateTableDef(ArchivoSalida)
    
    Call NuevoCampo(MyTable, "A2_FILIAL", dbText, 4)
    Call NuevoCampo(MyTable, "A2_COD", dbText, 13)
    Call NuevoCampo(MyTable, "A2_LOJA", dbText, 2)
    Call NuevoCampo(MyTable, "A2_NOME", dbText, 40)
    Call NuevoCampo(MyTable, "A2_NREDUZ", dbText, 20)
    Call NuevoCampo(MyTable, "A2_END", dbText, 40)
    Call NuevoCampo(MyTable, "A2_NR_END", dbText, 6)
    Call NuevoCampo(MyTable, "A2_BAIRRO", dbText, 20)
    Call NuevoCampo(MyTable, "A2_EST", dbText, 2)
    Call NuevoCampo(MyTable, "A2_MUN", dbText, 15)
    Call NuevoCampo(MyTable, "A2_CEP", dbText, 8)
    Call NuevoCampo(MyTable, "A2_TIPO", dbText, 1)
    Call NuevoCampo(MyTable, "A2_CGC", dbText, 14)
    Call NuevoCampo(MyTable, "A2_AGENRET", dbText, 1)
    Call NuevoCampo(MyTable, "A2_TEL", dbText, 50)
    Call NuevoCampo(MyTable, "A2_MCOMPRA", dbDouble, 8)
    Call NuevoCampo(MyTable, "A2_MSALDO", dbDouble, 8)
    Call NuevoCampo(MyTable, "A2_PRICOM", dbDate, 8)
    Call NuevoCampo(MyTable, "A2_ULTCOM", dbDate, 8)
    Call NuevoCampo(MyTable, "A2_SALDUP", dbDouble, 8)
    Call NuevoCampo(MyTable, "A2_SALDUPM", dbDouble, 8)
    Call NuevoCampo(MyTable, "A2_VINCULA", dbText, 1)
    Call NuevoCampo(MyTable, "A2_ID_REPR", dbText, 1)
    Call NuevoCampo(MyTable, "A2_EMAIL", dbText, 30)
    Call NuevoCampo(MyTable, "A2_PORIVA", dbDouble, 8)
    Call NuevoCampo(MyTable, "A2_PORGAN", dbDouble, 8)
    Call NuevoCampo(MyTable, "A2_PERCIVA", dbText, 1)
    Call NuevoCampo(MyTable, "A2_INSCGAN", dbText, 1)
    Call NuevoCampo(MyTable, "A2_PERCIB", dbText, 1)
    Call NuevoCampo(MyTable, "A2_AGREGAN", dbText, 2)
    Call NuevoCampo(MyTable, "A2_RETIB", dbText, 1)
    Call NuevoCampo(MyTable, "A2_CONTAB", dbText, 15)
    Call NuevoCampo(MyTable, "A2_B2B", dbText, 1)
    Call NuevoCampo(MyTable, "A2_PLCRRES", dbText, 1)
    Call NuevoCampo(MyTable, "A2_PLFIL", dbText, 1)
    Call NuevoCampo(MyTable, "A2_MSBLQL", dbText, 1)
    Call NuevoCampo(MyTable, "A2_RETISI", dbText, 1)
    Call NuevoCampo(MyTable, "A2_ISICM", dbText, 1)
    Call NuevoCampo(MyTable, "A2_RG1415", dbText, 1)
    Call NuevoCampo(MyTable, "A2_OBRICAI", dbText, 1)
    Call NuevoCampo(MyTable, "A2_MJURIDI", dbText, 1)
    Call NuevoCampo(MyTable, "A2_IMPIP", dbText, 1)
    Call NuevoCampo(MyTable, "A2_SITUACA", dbText, 1)
    Call NuevoCampo(MyTable, "A2_ENDOSSO", dbText, 1)
    Call NuevoCampo(MyTable, "A2_ARMADOR", dbText, 1)
    Call NuevoCampo(MyTable, "A2_AFIP", dbText, 6)
    
    ' AGREGA LA TABLA
    db.TableDefs.Append MyTable
    
    On Error GoTo 0
    
    ' LEE LA TABLA SQL
    Call ABRECONEXION
    Dim dbClientes As ADODB.Recordset
    Dim PosIva As Integer
    Dim PosicionIva As String
    Dim Personeria As String
    Dim PagaIva As String
    Dim ClienteSuspendido As String
    Dim CuitProveedor As String
    Dim CuentaMadre As String
    Dim CuentaMadreDefault As String
    
    On Error GoTo 98
    MensajeError = "Error al Leer Tabla de Proveedores"
    Set dbClientes = READSET("SELECT * FROM PROVED12 WITH(NOLOCK) WHERE NUME_CLI > 0")
    FIN& = CantRegistros("PROVED12", "NUME_CLI>0")
    On Error GoTo 0
    
    On Error GoTo 98
    MensajeError = "Error al Leer Archivo DBF Para Escritura"
    ' ABRE LA TABLA DBF EN MODO ESCRITURA
    Set RecordsetDBF = db.OpenRecordset(ArchivoSalida, dbOpenDynaset)
    On Error GoTo 0

    On Error GoTo 98
    MensajeError = ""
    I& = 1
    CuentaMadreDefault = TRIM$(CONTROL$("AMAPRO", "ECUECON"))
    With dbClientes
        Do While Not .EOF
        
            I& = I& + 1
            Call TRACE(24, I&, FIN&)
                    
            ' AGREGAR LOS REGISTROS DESDE LA TABLA SQL A LA TABLA DBF
            RecordsetDBF.AddNew
            RecordsetDBF!A2_FILIAL = ""
            RecordsetDBF!A2_COD = CStr(XVALO(!NUME_CLI))
            RecordsetDBF!A2_LOJA = ""
            RecordsetDBF!A2_NOME = Left$(SAFETEXT$(!RASO_CLI), 40)
            RecordsetDBF!A2_NREDUZ = Left$(SAFETEXT$(!FANT_CLI), 20)
            RecordsetDBF!A2_END = Left$(SAFETEXT$(!DOMI_CLI), 40)
            RecordsetDBF!A2_NR_END = ""
            RecordsetDBF!A2_BAIRRO = Left$(SAFETEXT$(!LOCA_CLI), 20)
            RecordsetDBF!A2_EST = Left$(SAFETEXT$(!PROV_CLI), 2)
            RecordsetDBF!A2_MUN = Left$(SAFETEXT$(!LOCA_CLI), 15)
            RecordsetDBF!A2_CEP = Left$(SAFETEXT$(!CPOS_CLI), 8)
                PosicionIva = "I"   ' RESPONSABLE INSCRIPTO
                Select Case PosIva
                    Case 0: PosicionIva = "F"       ' CONSUMIDOR FINAL
                    Case 2: PosicionIva = "N"       ' RESP. NO INSCRIPTO
                    Case 4: PosicionIva = "X"       ' EXENTO
                    Case 5, 7: PosicionIva = "E"    ' EXTERIOR
                    Case 6: PosicionIva = "M"       ' MONOTRIBUTO
                End Select
            RecordsetDBF!A2_TIPO = PosicionIva
                CuitProveedor = REPLACAR(SAFETEXT$(!CUIT_CLI), "-", "")
            RecordsetDBF!A2_CGC = Left$(CuitProveedor, 14)
            RecordsetDBF!A2_AGENRET = "N"   ' NO
            RecordsetDBF!A2_TEL = Left$(SAFETEXT$(!TELE_CLI), 25)
            RecordsetDBF!A2_MCOMPRA = 0
            RecordsetDBF!A2_MSALDO = 0
            RecordsetDBF!A2_PRICOM = CVDAT(FECHANUM("01/01/1980"))
            RecordsetDBF!A2_ULTCOM = CVDAT(FECHANUM("01/01/1980"))
            RecordsetDBF!A2_SALDUP = 0
            RecordsetDBF!A2_SALDUPM = 0
            RecordsetDBF!A2_VINCULA = 1     ' SIN VINCULO
            RecordsetDBF!A2_ID_REPR = 2     ' NO
            RecordsetDBF!A2_EMAIL = Left$(SAFETEXT$(!MAIL_CLI), 30)
            RecordsetDBF!A2_PORIVA = 0
            RecordsetDBF!A2_PORGAN = 0
                PagaIva = "S"               ' SI
                If PosIva = 4 Then PagaIva = "N"    ' NO
            RecordsetDBF!A2_PERCIVA = PagaIva
                InscriptoGanancias = "S"    ' SI
                If PosIva = 4 Or PosIva = 0 Then InscriptoGanancias = "N"   ' NO
            RecordsetDBF!A2_INSCGAN = InscriptoGanancias
            RecordsetDBF!A2_PERCIB = "N"    ' NO
            RecordsetDBF!A2_AGREGAN = ""
            RecordsetDBF!A2_RETIB = "N"     ' NO
                CuentaMadre = TRIM$(CODCUE$(XVALO(!CUMA_CLI)))
                If CuentaMadre = "" Then CuentaMadre = CuentaMadreDefault
            RecordsetDBF!A2_CONTAB = CODCUE$(XVALO(!CUMA_CLI))
            RecordsetDBF!A2_B2B = "" ' "2"       ' NO
            RecordsetDBF!A2_PLCRRES = "" ' "N"   ' NO
            RecordsetDBF!A2_PLFIL = "" ' "N"     ' NO
            RecordsetDBF!A2_MSBLQL = "" ' "2"    ' NO
            RecordsetDBF!A2_RETISI = "" ' "2"    ' NO
            RecordsetDBF!A2_ISICM = "" ' "2"     ' NO
            RecordsetDBF!A2_RG1415 = "" ' "2"    ' NO
            RecordsetDBF!A2_OBRICAI = "" ' "2"   ' NO
            RecordsetDBF!A2_MJURIDI = "" ' "2"   ' NO
            RecordsetDBF!A2_IMPIP = "" ' "3"     ' NO IMPRIME
            RecordsetDBF!A2_SITUACA = "1"   ' NORMAL
            RecordsetDBF!A2_ENDOSSO = "" ' "2"   ' NO
            RecordsetDBF!A2_ARMADOR = "" ' "2"   ' NO
            RecordsetDBF!A2_AFIP = "" ' "1"      ' SI
            RecordsetDBF.Update
            
            .MoveNext
        Loop
    End With
    
    On Error GoTo 0
    
    ' CIERRA LOS OBJETOS DE ACCESO A SQL Y AL ARCHIVO DBF
    On Error GoTo 98
    MensajeError = "Error al Cerrar Objetos"
    
    dbClientes.Close
    Set dbClientes = Nothing
    RecordsetDBF.Close
    Set RecordsetDBF = Nothing
    db.Close
    Set db = Nothing
    
    On Error GoTo 0
    Screen.MousePointer = 1
    Timer1.Enabled = False: lblTablaExportando = ""
    Call COMUNI("Exportación de PROVEEDORES Finalizada")
    Exit Sub

    ' CAPTURA DE ERROR
    
98  MensajeVisual = Err.Description
    If TRIM$(MensajeError) = "" Then MensajeError = MensajeVisual
    On Error GoTo 0
    
    On Error Resume Next
    dbClientes.Close
    Set dbClientes = Nothing
    RecordsetDBF.Close
    Set RecordsetDBF = Nothing
    db.Close
    Set db = Nothing
    On Error GoTo 0
    
    Screen.MousePointer = 1
    Timer1.Enabled = False: lblTablaExportando = ""
    Call COMUNI(MensajeError)

99  Screen.MousePointer = 1
    On Error GoTo 0
    
    On Error Resume Next
    dbClientes.Close
    Set dbClientes = Nothing
    RecordsetDBF.Close
    Set RecordsetDBF = Nothing
    db.Close
    Set db = Nothing
    On Error GoTo 0
    
    Timer1.Enabled = False: lblTablaExportando = ""
    Exit Sub

End Sub


Private Sub Form_Load()

   Call VERJOBID(OKEY%)
   If OKEY% < 1 Then Call FINAL("")

End Sub

Private Sub Form_Unload(Cancel As Integer)
    
   Call CIERRARCH("*.*")
   Call BAJALDISCO
   Call FINAL("")

End Sub
Private Sub NuevoCampo(ByRef TABLA As TableDef, NombreCampo As String, TipoDato As Integer, Optional LongitudCampo As Integer)
 
    Dim MyField As Field
    
    Set MyField = TABLA.CreateField(NombreCampo, TipoDato, LongitudCampo)
    TABLA.Fields.Append MyField
    
End Sub

Private Sub Timer1_Timer()
    
    If lblTextoAvance = "" Then
        lblTextoAvance = UCase(lblTablaExportando)
    Else
        lblTextoAvance = ""
    End If
    
End Sub


