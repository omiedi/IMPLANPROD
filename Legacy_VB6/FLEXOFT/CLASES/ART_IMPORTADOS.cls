VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ART_IMPORTADOS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'===================================================================================================================
'========================================= DECLARACION DE PROPIEDADES ==============================================
'=========================================          (ENTIDAD)         ==============================================
'===================================================================================================================

Private Type art_impo
    Aduana As String
    Despacho As String
    FECHA As Date
    AÑO As String
    EsImportado As Integer
End Type
Private Art_Importados As art_impo

    Public Property Get Aduana() As String
       Aduana = Art_Importados.Aduana
    End Property
    
    Public Property Let Aduana(valor As String)
        Art_Importados.Aduana = valor
    End Property
    Public Property Get AÑO() As String
       AÑO = Art_Importados.AÑO
    End Property
    
    Public Property Let AÑO(valor As String)
        Art_Importados.AÑO = valor
    End Property
    
    Public Property Get Despacho() As String
       Despacho = Art_Importados.Despacho
    End Property
    
    Public Property Let Despacho(valor As String)
        Art_Importados.Despacho = valor
    End Property
    Public Property Get FECHA() As Date
       FECHA = Art_Importados.FECHA
    End Property
    
    Public Property Let FECHA(valor As Date)
        Art_Importados.FECHA = valor
    End Property
    Public Property Get EsImportado() As String
       EsImportado = Art_Importados.EsImportado
    End Property
    
    Public Property Let EsImportado(valor As String)
        Art_Importados.EsImportado = valor
    End Property

'===================================================================================================================
'========================================= FIN DECLARACION DE PROPIEDADES ==========================================
'===================================================================================================================



'===================================================================================================================
'========================================= *****    CAPA DE DATOS   ****  ==========================================
'===================================================================================================================

    
    Sub GrabarDatosAduanaDespFecha(LOTE As String, NUBORE As Long)

        Dim OBJCMD As ADODB.Command 'El objeto
        Set OBJCMD = New ADODB.Command
        OBJCMD.CommandType = adCmdStoredProc
        OBJCMD.CommandText = "GRABADATOSINGDESPACHO_V1" 'NOMBRE DEL PROCEDIMIENTO ALMACENADO
        Set OBJCMD.ActiveConnection = FLEXCONN
        Dim valor_NuevoID As Integer
        'Los parametros
        OBJCMD.Parameters(1) = CodiEmpr% '@CodiEmpr smallint
        OBJCMD.Parameters(2) = FETEXCOR1$(CVINT(Art_Importados.FECHA)) '@FechMov smalldatetime
        OBJCMD.Parameters(3) = Art_Importados.Despacho
        OBJCMD.Parameters(4) = Art_Importados.Aduana
        OBJCMD.Parameters(5) = LOTE
        OBJCMD.Parameters(6) = NUBORE
        '
        OBJCMD.Execute
        Set OBJCMD = Nothing
    End Sub
    
    Sub READ_DATOS_IMPO(NBORE, Optional STRINGBORE$)
      'devuelve el detalle de una recepcion MAS LOS DATOS DE DESPACHO SE CREO PARA MEDITEA
      'LIMPIO LOS ATRIBUTOS POR SI NO ENTRA EN EL RECORDSET
      Me.Despacho = ""
      Me.Aduana = ""
      Me.FECHA = Date
      NUBORETEX$ = TRIM$(STRINGBORE$)
      Dim RS_FUN As ADODB.Recordset
      Set RS_FUN = New ADODB.Recordset
      SQLX$ = "SELECT TOP 1 D.FECHA,D.DESPACHO,D.ADUANA "
      SQLX$ = SQLX$ + " FROM DESPAIMPO D WITH(NOLOCK)"
      If NUBORETEX$ <> "" Then
         SQLX$ = SQLX$ + " WHERE D.IDENLOTE = '" & NUBORETEX$ & "'"
      Else
         SQLX$ = SQLX$ + " WHERE D.NUMEBORE = " & NBORE
      End If
      
      Set RS_FUN = READSET(SQLX$)
      If RS_FUN.EOF <> True Then
        Me.Despacho = TRIM$(RS_FUN!Despacho)
        Me.Aduana = TRIM$(RS_FUN!Aduana)
        Me.FECHA = TRIM$(RS_FUN!FECHA)
        Me.AÑO = Year(RS_FUN!FECHA)
      End If
      Set RS_FUN = Nothing
    End Sub
    Sub VeriImportado(CODI$)
        Me.EsImportado = 0
        SQLX$ = "SELECT COUNT(*) AS CANTI FROM DATASQL WITH(NOLOCK) WHERE LISTA = 'Equipos_Importados' AND CAMPO1 = '" & CODI$ & "'"
        Set rs = READSET(SQLX$)
        
        If XVALO(rs!CANTI) > 0 Then
           Me.EsImportado = 1
        Else
           Me.EsImportado = 0
        End If
    End Sub
    Sub GrabarItemImportado(CODI$)
        Static ID_LISTAX
        If ID_LISTAX < 1 Then
          ID_LISTAX = XVALO(VALOBADA("DATASQL", "ID_LISTA", "LISTA = 'Equipos_Importados'", "NOMESS"))
        End If
        CI% = CODINT%(CODI$)
        DESCRIX$ = DESCRIT0$(CI%)
        SQLX$ = "INSERT INTO DATASQL ("
        SQLX$ = SQLX$ + "CODIEMPR,ID_LISTA,  LISTA,CAMPO1,CAMPO2)"
        SQLX$ = SQLX$ + " VALUES(" & CodiEmp%
        SQLX$ = SQLX$ + " ," & ID_LISTAX
        SQLX$ = SQLX$ + " ,'Equipos_Importados'"
        SQLX$ = SQLX$ + " ,'" & CODI$ & "'"
        SQLX$ = SQLX$ + " ,'" & DESCRIX$ & "'"
        SQLX$ = SQLX$ + ")"
        FLEXCONN.Execute (SQLX$)
    End Sub

'===================================================================================================================
'========================================= *****  FIN CAPA DE DATOS   **** =========================================
'===================================================================================================================

