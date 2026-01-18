VERSION 5.00
Begin VB.Form ABMNUMSE 
   Caption         =   "ABM de Números de Serie"
   ClientHeight    =   5220
   ClientLeft      =   225
   ClientTop       =   555
   ClientWidth     =   8340
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5220
   ScaleWidth      =   8340
   StartUpPosition =   2  'CenterScreen
   Begin VB.Menu mnuTransacciones 
      Caption         =   "Transacciones"
      Begin VB.Menu mnuNumerosDeSerie 
         Caption         =   "Numeros de Serie"
         Begin VB.Menu mnuNumerosSerieRecibidos 
            Caption         =   "Materiales Recibidos"
         End
         Begin VB.Menu mnuNumerosSerieFabricados 
            Caption         =   "Materiales Fabricados"
         End
      End
      Begin VB.Menu mnuModificarNumeroSerie 
         Caption         =   "Modificar Datos de Número de Serie"
      End
      Begin VB.Menu mnuBajas 
         Caption         =   "Bajas"
         Begin VB.Menu mnuBajaNumeroSerie 
            Caption         =   "Baja de Número de Serie"
         End
         Begin VB.Menu mnuRevertirBajaNumeroSerie 
            Caption         =   "Revertir Baja de Número de Serie"
         End
      End
      Begin VB.Menu mnuAnulaciones 
         Caption         =   "Anulaciones"
         Begin VB.Menu mnuAnularNumeroSerie 
            Caption         =   "Anular Número de Serie"
         End
         Begin VB.Menu mnuRevertirAnulacionNumeroSerie 
            Caption         =   "Revertir Anulación de Número de Serie"
         End
      End
      Begin VB.Menu mnuSalir 
         Caption         =   "Salir"
      End
   End
   Begin VB.Menu mnuConsultasListados 
      Caption         =   "Consultas y Listados"
      Begin VB.Menu mnuAbrirFormularioCalidad 
         Caption         =   "Abrir Formulario de Calidad"
      End
      Begin VB.Menu linea2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuReportesVarios 
         Caption         =   "Reportes Varios"
      End
   End
   Begin VB.Menu mnuVer 
      Caption         =   "Ver"
      Begin VB.Menu mnuReimprimirFormularioDeCalidad 
         Caption         =   "Re-imprimir Formulario de Calidad"
      End
   End
   Begin VB.Menu mnuUtilidades 
      Caption         =   "Utilidades"
   End
   Begin VB.Menu mnuConfiguracion 
      Caption         =   "Configuración"
      Begin VB.Menu mnuAsociarReporte 
         Caption         =   "Vincular Reportes al Formulario"
      End
      Begin VB.Menu linea1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFamilias 
         Caption         =   "Tabla de Familias"
      End
      Begin VB.Menu mnuFormularioCalidadPorFamilia 
         Caption         =   "Formularios de Calidad por Familia"
      End
   End
   Begin VB.Menu mnuAccesosDirectos 
      Caption         =   "Accesos Directos"
      Begin VB.Menu mnuMaestroItemsDeStock 
         Caption         =   "Maestro de Items de Stock"
      End
   End
End
Attribute VB_Name = "ABMNUMSE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Type DatosNumeroSerie
    NumeroSerie As Long
    Codigo As String
    Descripcion As String
    Proveedor As String
    NumeroRecepcion As Long
    Colada As Long
End Type
Private Sub BajaAnulaNumeroSerie(MODO As Integer)

    ' LISTA DE MATERIALES RECIBIDOS CON NUMERO DE SERIE
    
10  TituloLista$ = "MATERIALES RECIBIDOS CON NUMERO DE SERIE"
    CAMPOS$ = "N-Serie/A10;Código/A24;Descripción/A36;Recepción/A10;Colada/A24;Fam/A3"
    Call CARGA_ENCABEZADO(FRMZELECHO.msf2, CAMPOS$, FRMZELECHO)
    Call CARGA_ENCABEZADO(FRMZELECHO.MSF1, CAMPOS$, FRMZELECHO)
    FRMZELECHO.Caption = TituloLista$
    
    SQLX$ = "SELECT * FROM TRAZARECE WITH(NOLOCK) WHERE NUMEROSERIE > 0 AND STATUS < 8"
    Set RS = READSET(SQLX$)
    J& = 0
    With RS
        If Not (.EOF And .BOF) Then
            FIN& = RS.RecordCount
            FRMZELECHO.msf2.Rows = FIN& + 1 ' ESTABLECE DE ANTEMANO LA CANTIDAD DE FILAS A UTILIZAR
        End If
        Do While Not .EOF
          IDLOTE$ = !IDENLOTE
          CodigoExterno$ = !COD_EXTE
          Descripcion$ = !Descrip
          NumeroSerie$ = CStr(!NumeroSerie)
          NumeroColada$ = CStr(!NumeroColada)
          Familia$ = !CFAMILIA
          
          J& = J& + 1
          Call TRACE(24, J&, FIN&)
          FRMZELECHO.msf2.TextMatrix(J&, 1) = NumeroSerie$
          FRMZELECHO.msf2.TextMatrix(J&, 2) = CodigoExterno$
          FRMZELECHO.msf2.TextMatrix(J&, 3) = Descripcion$
          FRMZELECHO.msf2.TextMatrix(J&, 4) = IDLOTE$
          FRMZELECHO.msf2.TextMatrix(J&, 5) = NumeroColada$
          FRMZELECHO.msf2.TextMatrix(J&, 6) = Familia$
          .MoveNext
        Loop
    End With
    On Error Resume Next
    RS.Close
    Set RS = Nothing
    On Error GoTo 0
    
    FRMZELECHO.Width = 11000
60  FRMZELECHO.Show 1
    
    NSERIE& = XVALO(TRIM(RESP_ZELE_VECT(1)))
    If NSERIE& = 0 Then Exit Sub
    CodigoExterno$ = TRIM(RESP_ZELE_VECT(2))
    CodigoInterno% = CODINT%(CodigoExterno$)
    IDLOTE$ = TRIM(RESP_ZELE_VECT(4))
        
    OBX$ = OBJPLA$("OBSNUMSE", VDEF$)
    MotivoBajaAnulacion$ = Mid$(OBX$, 1, 48) & Mid$(OBX$, 49, 48) & Mid$(OBX$, 97, 48)
    If MODO% = 1 Then
        ' STATUS = 8 => DADO DE BAJA
        Call ACTUREGISTRO("TRAZARECE", "STATUS", "COD_INTE=" & CStr(CodigoInterno%) & " AND IDENLOTE='" & IDLOTE$ & "' AND NUMEROSERIE=" & CStr(NSERIE&), 8, REAF&)
        Call ACTUREGISTRO("TRAZARECE", "MOTIVOBAJA", "COD_INTE=" & CStr(CodigoInterno%) & " AND IDENLOTE='" & IDLOTE$ & "' AND NUMEROSERIE=" & CStr(NSERIE&), MotivoBajaAnulacion$, REAF&)
        
        Call COMUNI("Número de Serie Dado de Baja")
    ElseIf MODO% = 2 Then
        ' STATUS = 9 => ANULADO
        Call ACTUREGISTRO("TRAZARECE", "STATUS", "COD_INTE=" & CStr(CodigoInterno%) & " AND IDENLOTE='" & IDLOTE$ & "' AND NUMEROSERIE=" & CStr(NSERIE&), 9, REAF&)
        Call ACTUREGISTRO("TRAZARECE", "MOTIVOBAJA", "COD_INTE=" & CStr(CodigoInterno%) & " AND IDENLOTE='" & IDLOTE$ & "' AND NUMEROSERIE=" & CStr(NSERIE&), MotivoBajaAnulacion$, REAF&)
        
        Call COMUNI("Número de Serie Anulado")
    End If

End Sub

Private Sub RevertirBajaAnulacionDeNumeroDeSerie(MODO As Integer)

10  TituloLista$ = "MATERIALES CON NUMERO DE SERIE DE BAJA"
    If MODO = 2 Then TituloLista$ = "MATERIALES CON NUMERO DE SERIE ANULADOS"
    CAMPOS$ = "N-Serie/A10;Código/A24;Descripción/A36;Recepción/A10;Colada/A24;Fam/A3"
    Call CARGA_ENCABEZADO(FRMZELECHO.msf2, CAMPOS$, FRMZELECHO)
    Call CARGA_ENCABEZADO(FRMZELECHO.MSF1, CAMPOS$, FRMZELECHO)
    FRMZELECHO.Caption = TituloLista$
    
    SQLX$ = "SELECT * FROM TRAZARECE WITH(NOLOCK) WHERE NUMEROSERIE > 0 AND STATUS = 8"
    If MODO = 2 Then SQLX$ = "SELECT * FROM TRAZARECE WITH(NOLOCK) WHERE NUMEROSERIE > 0 AND STATUS = 9"
    Set RS = READSET(SQLX$)
    J& = 0
    With RS
        If Not (.EOF And .BOF) Then
            FIN& = RS.RecordCount
            FRMZELECHO.msf2.Rows = FIN& + 1 ' ESTABLECE DE ANTEMANO LA CANTIDAD DE FILAS A UTILIZAR
        End If
        Do While Not .EOF
          IDLOTE$ = !IDENLOTE
          CodigoExterno$ = !COD_EXTE
          Descripcion$ = !Descrip
          NumeroSerie$ = CStr(!NumeroSerie)
          NumeroColada$ = CStr(!NumeroColada)
          Familia$ = !CFAMILIA
          
          J& = J& + 1
          Call TRACE(24, J&, FIN&)
          FRMZELECHO.msf2.TextMatrix(J&, 1) = NumeroSerie$
          FRMZELECHO.msf2.TextMatrix(J&, 2) = CodigoExterno$
          FRMZELECHO.msf2.TextMatrix(J&, 3) = Descripcion$
          FRMZELECHO.msf2.TextMatrix(J&, 4) = IDLOTE$
          FRMZELECHO.msf2.TextMatrix(J&, 5) = NumeroColada$
          FRMZELECHO.msf2.TextMatrix(J&, 6) = Familia$
          .MoveNext
        Loop
    End With
    On Error Resume Next
    RS.Close
    Set RS = Nothing
    On Error GoTo 0
    
    FRMZELECHO.Width = 11000
60  FRMZELECHO.Show 1
    
    NSERIE& = XVALO(TRIM(RESP_ZELE_VECT(1)))
    If NSERIE& = 0 Then Exit Sub
    CodigoExterno$ = TRIM(RESP_ZELE_VECT(2))
    CodigoInterno% = CODINT%(CodigoExterno$)
    IDLOTE$ = TRIM(RESP_ZELE_VECT(4))
    
    Call ACTUREGISTRO("TRAZARECE", "STATUS", "COD_INTE=" & CStr(CodigoInterno%) & " AND IDENLOTE='" & IDLOTE$ & "' AND NUMEROSERIE=" & CStr(NSERIE&), 0, REAF&)
    Call ACTUREGISTRO("TRAZARECE", "MOTIVOBAJA", "COD_INTE=" & CStr(CodigoInterno%) & " AND IDENLOTE='" & IDLOTE$ & "' AND NUMEROSERIE=" & CStr(NSERIE&), "", REAF&)

    Call COMUNI("Movimiento Revertido Correctamente")

End Sub

Private Sub Form_Load()
    Show
    Refresh
    '
    Call VERJOBID(OKEY%)
    If OKEY% < 1 Then Close: End
    '
    EPAC$ = TRIM$(EMPREAC$)
    If EPAC$ <> "" Then
        Caption = Caption + "  -  " + EPAC$
    End If
    Call GRATITFOR(Caption)
    '
    Call ABRECONEXION
    Call InstalaFuenteAnchoFijo
    Call CREATABLA_STOCKS
    Call GENERAR_ESTRUCTURAS_SQL

End Sub

Private Sub Form_Unload(Cancel As Integer)

    Call CIERRARCH("*.*")
    Call FINAL("")
    
End Sub



Private Sub mnuAnularNumeroSerie_Click()

    Call BajaAnulaNumeroSerie(2)

End Sub

Private Sub mnuAsociarReporte_Click()

    Call CONECTA_CONSUL_PRE(TRIM$(UCase$(Name)))

End Sub

Private Sub mnuBajaNumeroSerie_Click()

    Call BajaAnulaNumeroSerie(1)
    
End Sub

Private Sub mnuFamilias_Click()

   Call TRALOCAL("TAFAMIL", "")

10  VTB% = VENTABU("TAFAMIL")
    If VTB% >= 0 Then
      For U& = 1 To ULTREG&("!TAFAMIL")
        X$ = REGLEIDO$("!TAFAMIL", U&)
        CD$ = TRIM$(UCase$(Mid$(X$, 1, 2)))
        Call REPLA(X$, CD$, 1, 2)
        Call GRAREG("!TAFAMIL", X$, U&)
        '
        ' Valida que el codigo no sea un valor nulo o menor a dos caracteres
        If Len(CD$) < 2 Then
          Call MENSERR(24, "Código de Familia Menor a dos Caracteres")
          GoTo 10
        End If
        ' Valida que la familia posea descripcion
        DESC$ = TRIM$(Mid$(X$, 3, 27))
        If Len(DESC$) < 1 Then
          Call MENSERR(24, "Codigo de Familia sin Descripción")
          GoTo 10
        End If
        ' Valida que la familia posea descripcion
        For U2& = 1 To U& - 1
          Y$ = REGLEIDO$("!TAFAMIL", U2&)
          If TRIM$(Mid$(X$, 1, 2)) = TRIM$(Mid$(Y$, 1, 2)) Then
            Call MENSERR(24, "Código de Familia Repetido")
            GoTo 10
          End If
        Next U2&
      Next U&
      Call TRACENTRAL("TAFAMIL", "")
      
    End If
    
End Sub

Private Sub mnuFormularioCalidadPorFamilia_Click()

  Call GENERAR_ESTRUCTURAS_SQL
  IDLIST& = ObtenerIdTablaAuxiliar("FOCALFA")
  
  Call ABM_TABLA(IDLIST&)

End Sub

Private Sub mnuAbrirFormularioCalidad_Click()

    Dim DatosNumeroSerie As DatosNumeroSerie
    
    Call SELECHO("TAFAMIL")
    Familia$ = TRIM$(VALACT1$("TAFAMIL"))
    
    If Familia$ = "" Then Exit Sub
    DescripcionFamilia$ = VALACT2$("TAFAMIL")
    
    Call ImprimeFormulariosDeCalidad(Familia$, DatosNumeroSerie)
    
End Sub

Private Sub mnuMaestroItemsDeStock_Click()

    Call CONECRUN("AMASTO07", "Maestro de Items de Stock")

End Sub

Private Sub mnuModificarNumeroSerie_Click()

10  TituloLista$ = "MATERIALES RECIBIDOS CON NUMERO DE SERIE"
    CAMPOS$ = "N-Serie/I10;Código/A24;Descripción/A36;CodInterno/I0"
    Call CARGA_ENCABEZADO(FRMZELECHO.msf2, CAMPOS$, FRMZELECHO)
    Call CARGA_ENCABEZADO(FRMZELECHO.MSF1, CAMPOS$, FRMZELECHO)
    FRMZELECHO.Caption = TituloLista$
    
    Dim rst As ADODB.Recordset
    ConsultaSql$ = "SELECT NumeroSerie, COD_INTE, COD_EXTE, DESCRIP FROM TRAZARECE WITH(NOLOCK) WHERE NUMEROSERIE > 0 AND STATUS < 8"
    Set rst = READSET(ConsultaSql$)
    
    With rst
        Do While Not .EOF
            FRMZELECHO.msf2.Rows = FRMZELECHO.msf2.Rows + 1
            FRMZELECHO.msf2.TextMatrix(FRMZELECHO.msf2.Rows - 1, 1) = !NumeroSerie
            FRMZELECHO.msf2.TextMatrix(FRMZELECHO.msf2.Rows - 1, 2) = !COD_EXTE
            FRMZELECHO.msf2.TextMatrix(FRMZELECHO.msf2.Rows - 1, 3) = !Descrip
            FRMZELECHO.msf2.TextMatrix(FRMZELECHO.msf2.Rows - 1, 4) = !COD_INTE
            
            .MoveNext
        Loop
    End With
    On Error Resume Next
    rst.Close
    Set rst = Nothing
    On Error GoTo 0
    
    FRMZELECHO.Width = 8500
    FRMZELECHO.Show 1
    NumeroSerie& = XVALO(RESP_ZELE_VECT(1))
    If NumeroSerie& = 0 Then Exit Sub
    CodigoInterno% = XVALO(RESP_ZELE_VECT(4))

    Dim rstNumeroSerie As ADODB.Recordset
    ConsultaSql$ = "SELECT * FROM TRAZARECE WITH(NOLOCK) WHERE NUMEROSERIE =" & CStr(NumeroSerie&) & " AND COD_INTE=" & CStr(CodigoInterno%)
    Set rstNumeroSerie = READSET(ConsultaSql$)
    With rstNumeroSerie
        CodigoExterno$ = !COD_EXTE
        Descripcion$ = !Descrip
        CantidadAprobada& = 1
        IdRecepcion$ = !IDENLOTE
        CodigoInternoDestino% = !CodInte_DestinoConsumo
        CodigoExternoDestino$ = !CodExte_DestinoConsumo
        NumeroSerie& = !NumeroSerie
        NumeroColada& = !NumeroColada
        Material$ = !Material
        Observaciones$ = !Observaciones
        Modelo$ = !Modelo
        MARCA$ = !MARCA
        NumeroSerieOrigen$ = !NumeroSerieOrigen
        ClienteDestino% = !ClienteDestino
        Seteo$ = !Seteo
        Alambre$ = !Alambre
    End With
    NumeroRecepcion& = XVALO(VALOBADA("BOLRECEP", "NUMEBORE", "IDENLOTE='" & IdRecepcion$ & "'"))
    ProveedorRecepcion$ = VALOBADA("BOLRECEP", "RASO_PROV", "IDENLOTE='" & IdRecepcion$ & "'")
    DescripcionDestino$ = DESCRIT0$(CodigoInternoDestino%)
    
    VDEF$ = Space(260)
    Call REPLA(VDEF$, CodigoExterno$, 1, 24)
    Call REPLA(VDEF$, Descripcion$, 25, 35)
    Call REPLA(VDEF$, MKS$(XVALO(CantidadAprobada&)), 60, 4)
    Call REPLA(VDEF$, MKS$(NumeroRecepcion&), 64, 32)
    Call REPLA(VDEF$, ProveedorRecepcion$, 68, 32)
    Call REPLA(VDEF$, CodigoExternoDestino$, 100, 24)
    Call REPLA(VDEF$, DescripcionDestino$, 124, 32)
    Call REPLA(VDEF$, DescripcionFamilia$, 156, 24)
    Call REPLA(VDEF$, MKS$(NumeroSerie&), 180, 4)
    
    Call REPLA(VDEF$, MKS$(NumeroColada&), 184, 4)
    Call REPLA(VDEF$, Material$, 188, 24)
    Call REPLA(VDEF$, Observaciones$, 212, 48)
    Call REPLA(VDEF$, Modelo$, 260, 24)
    Call REPLA(VDEF$, MARCA$, 284, 24)
    Call REPLA(VDEF$, NumeroSerieOrigen$, 308, 24)
    Call REPLA(VDEF$, MKS$(ClienteDestino%), 332, 4)
    Call REPLA(VDEF$, Seteo$, 336, 24)
    Call REPLA(VDEF$, Alambre$, 360, 24)
        
    OBX$ = OBJPLA$("MODNUMSER", VDEF$)
    If OBX$ = "" Then GoTo 10
    
    ' REGISTRA NUMERO DE SERIE
    
    NumeroColada& = CVS(Mid$(OBX$, 184, 4))
    Material$ = TRIM$(Mid$(OBX$, 188, 24))
    Observaciones$ = TRIM$(Mid$(OBX$, 212, 48))
    Modelo$ = TRIM$(Mid$(OBX$, 260, 24))
    MARCA$ = TRIM$(Mid$(OBX$, 284, 24))
    NumeroSerieOrigen$ = TRIM$(Mid$(OBX$, 308, 24))
    ClienteDestino% = CVS(Mid$(OBX$, 332, 4))
    Seteo$ = TRIM$(Mid$(OBX$, 336, 24))
    Alambre$ = TRIM$(Mid$(OBX$, 360, 24))
    
    Dim RstTrazarece As ADODB.Recordset
    Set RstTrazarece = New ADODB.Recordset
    ConsultaSql$ = "SELECT * FROM TRAZARECE WHERE COD_INTE=" & CStr(CodigoInterno%) & " AND NUMEROSERIE=" & CStr(NumeroSerie&)
    RstTrazarece.Open ConsultaSql$, FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
    With RstTrazarece

            !NumeroColada = NumeroColada&
            !CodInte_DestinoConsumo = CODINT%(CodigoExternoDestino$)
            !CodExte_DestinoConsumo = CodigoExternoDestino$
            !Material = Material$
            !Observaciones = Observaciones$
            !Modelo = Modelo$
            !MARCA = MARCA$
            !NumeroSerieOrigen = NumeroSerieOrigen$
            !ClienteDestino = ClienteDestino%
            !Seteo = Seteo$
            !Alambre = Alambre$
            !Status = 0
            !ORDEN = i&
            .Update
            
    End With
    On Error Resume Next
    RstTrazarece.Close
    Set RstTrazarece = Nothing
    On Error GoTo 0
    
    Call COMUNI("Número de Serie Actualizado")
    
    GoTo 10

End Sub

Private Sub mnuNumerosSerieRecibidos_Click()

    ' PRESENTA LISTA DE RECEPCIONES
    '
10  TituloLista$ = "Recepciones de materiales"
    If TRIM$(FRMZELECHO.Caption) = TituloLista$ Then GoTo 50
    
    Call CARGA_ENCABEZADO(FRMZELECHO.msf2, "Recepcion/A12;Fecha/D8;Remito/A20;Proveedor/A32", FRMZELECHO)
    Call CARGA_ENCABEZADO(FRMZELECHO.MSF1, "Recepcion/A12;Fecha/D8;Remito/A20;Proveedor/A32", FRMZELECHO)
    FRMZELECHO.Caption = TituloLista$
    
    Dim obj As New RECORXET
    Set RS1 = obj.RS_BOLRECEPEN()
    Call Carga_MSF_Recordset(RS1, "Recepcion/A12;Fecha/D8;Remito/A20;Proveedor/A32", FRMZELECHO.msf2, 1)
    On Error Resume Next
    Set obj = Nothing
    RS1.Close
    Set RS1 = Nothing
    On Error GoTo 0
        
    FRMZELECHO.Width = 10000
50  FRMZELECHO.Show 1
    
    NumeroRecepcion$ = TRIM(RESP_ZELE_VECT(1))
    
    If NumeroRecepcion$ = "" Then Exit Sub
    
    ' PRESENTA LISTA DE ITEMS DE LA RECEPCION ELEGIDA
    '
    TituloLista$ = "LISTA DE ITEMS DE RECEPCION '" & NumeroRecepcion$ & "'"
    CAMPOS$ = "Codigo/A24;Descripcin/A48;Aprobado/F12.1;Lote/A16"
    Call CARGA_ENCABEZADO(FRMZELECHO.msf2, CAMPOS$, FRMZELECHO)
    Call CARGA_ENCABEZADO(FRMZELECHO.MSF1, CAMPOS$, FRMZELECHO)
    FRMZELECHO.Caption = TituloLista$
    
    SQLX$ = "SELECT IDENLOTE, COD_INTE, CANTOAPRO FROM BOLRECEP WITH(NOLOCK) WHERE NUMEBORE=" & NumeroRecepcion$
    Set RS = READSET(SQLX$)
    J& = 0
    With RS
        If Not (.EOF And .BOF) Then
            FIN& = RS.RecordCount
            FRMZELECHO.msf2.Rows = FIN& + 1 ' ESTABLECE DE ANTEMANO LA CANTIDAD DE FILAS A UTILIZAR
        End If
        Do While Not .EOF
          IDLOTE$ = SAFETEXT$(!IDENLOTE)
          CodigoInterno% = XVALO(!COD_INTE)
          CodigoExterno$ = CODEXT$(CodigoInterno%)
          Descripcion$ = DESCRIT0$(CodigoInterno%)
          CantidadAprobada& = XVALO(!CANTOAPRO)
          
          J& = J& + 1
          Call TRACE(24, J&, FIN&)
          FRMZELECHO.msf2.TextMatrix(J&, 1) = CodigoExterno$
          FRMZELECHO.msf2.TextMatrix(J&, 2) = Descripcion$
          FRMZELECHO.msf2.TextMatrix(J&, 3) = CantidadAprobada&
          FRMZELECHO.msf2.TextMatrix(J&, 4) = IDLOTE$
          .MoveNext
        Loop
    End With
    On Error Resume Next
    RS.Close
    Set RS = Nothing
    On Error GoTo 0
    
    FRMZELECHO.Width = 10000
60  FRMZELECHO.Show 1
    
    IdLoteElegido$ = TRIM(RESP_ZELE_VECT(4))
    If IdLoteElegido$ = "" Then GoTo 10
    
    CodigoExterno$ = TRIM(RESP_ZELE_VECT(1))
    CodigoInterno% = CODINT%(CodigoExterno$)
    Descripcion$ = TRIM(RESP_ZELE_VECT(2))
    CantidadAprobada& = 1
    
    
    ' EL USUARIO ELIGE EL CODIGO DE DESTINO (EN EL QUE SE VA A TRANSFORMAR PRODUCTO DE LA FABRICACION)

    CodigoExternoDestino$ = ""
    CodigoExternoDestino$ = ListaSeleccionArticulos("", False)
    If CodigoExternoDestino$ = "" Then
        GoTo 10
    End If
    DescripcionDestino$ = DESCRIT0$(CODINT%(CodigoExternoDestino$))
     
     
    ' BUSCA LA FAMILIA DEL CODIGO DE DESTINO. EL NUMERO DE SERIE DEL CODIGO RECEPCIONADO SE OBTIENE A PARTIR DE LA FAMILIA DEL CODIGO DE DESTINO ELEGIDO POR EL USUARIO.
    
    CodigoInternoDestino% = CODINT(CodigoExternoDestino$)
    Familia$ = TRIM$(CODFAMIL$(CodigoInternoDestino%))
    If Familia$ = "" Then
        RESPUESTA% = COMALTER%("El Artículo '" & CodigoExternoDestino$ & "' No Pertenece a una Familia.\ \¿Desea Asignar en este Momento?\\Si, Asignar\No, Continuar")
        If RESPUESTA% = 2 Then GoTo 10
        
        Call SELECHO("TAFAMIL")
        Familia$ = TRIM$(VALACT1$("TAFAMIL"))
        If Familia$ = "" Then GoTo 10
        Call ACTUREGISTRO("MASTER", "FAMILIA", "CODINT=" & CStr(CodigoInternoDestino%), Familia$, REAF&)
    End If
    DescripcionFamilia$ = ECOARCH$("TAFAMIL", Familia$)
    
    
    ' CARGA DATOS ASOCIADOS AL NUMERO DE SERIE
    
70  ProveedorRecepcion$ = SAFETEXT$(VALOBADA("BOLRECEP", "RASO_PROV", "NUMEBORE=" & NumeroRecepcion$))
75  UltimoNumeroSerie& = XVALO(VALOBADA("TRAZARECE", "MAX(NUMEROSERIE)", "CFAMILIA='" & Familia$ & "'")) + 1
    VDEF$ = Space(260)
    Call REPLA(VDEF$, CodigoExterno$, 1, 24)
    Call REPLA(VDEF$, Descripcion$, 25, 35)
    Call REPLA(VDEF$, MKS$(XVALO(CantidadAprobada&)), 60, 4)
    Call REPLA(VDEF$, MKS$(XVALO(NumeroRecepcion$)), 64, 32)
    Call REPLA(VDEF$, ProveedorRecepcion$, 68, 32)
    Call REPLA(VDEF$, CodigoExternoDestino$, 100, 24)
    Call REPLA(VDEF$, DescripcionDestino$, 124, 32)
    Call REPLA(VDEF$, DescripcionFamilia$, 156, 24)
    Call REPLA(VDEF$, MKS$(UltimoNumeroSerie&), 180, 4)
    Call REPLA(VDEF$, MKS$(0), 184, 4)
        
    OBX$ = OBJPLA$("NUMSERMATRE", VDEF$)
    If OBX$ = "" Then GoTo 10
    
    
    ' REGISTRA NUMERO DE SERIE
    
    NumeroSerie& = CVS(Mid$(OBX$, 180, 4))
    If CantRegistros&("TRAZARECE", "CFAMILIA='" & Familia$ & "' AND NUMEROSERIE=" & CStr(NumeroSerie&) & " AND STATUS < 9") > 0 Then  ' STATUS = 9 => ANULADO
        Call MENSERR(24, "Número de Serie Ocupado")
        GoTo 75
    End If
    NumeroColada& = CVS(Mid$(OBX$, 184, 4))
    Material$ = TRIM$(Mid$(OBX$, 188, 24))
    Observaciones$ = TRIM$(Mid$(OBX$, 212, 48))
    Modelo$ = TRIM$(Mid$(OBX$, 260, 24))
    MARCA$ = TRIM$(Mid$(OBX$, 284, 24))
    NumeroSerieOrigen$ = TRIM$(Mid$(OBX$, 308, 24))
    ClienteDestino% = CVS(Mid$(OBX$, 332, 4))
    Seteo$ = TRIM$(Mid$(OBX$, 336, 24))
    Alambre$ = TRIM$(Mid$(OBX$, 360, 24))
    
    Dim RstTrazarece As ADODB.Recordset
    Set RstTrazarece = New ADODB.Recordset
    ConsultaSql$ = "SELECT * FROM TRAZARECE WHERE IDENLOTE=''"
    RstTrazarece.Open ConsultaSql$, FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
    With RstTrazarece

            .AddNew
            !COD_INTE = CodigoInterno%
            !COD_EXTE = CodigoExterno$
            !Descrip = Descripcion$
            !IDENLOTE = IdLoteElegido$
            !NumeroSerie = NumeroSerie&
            !NumeroColada = NumeroColada&
            !CFAMILIA = Familia$
            !ORIGEN = ""
            !DESTINO = ""
            !DESTINOCONSUMO = ""
            !CodInte_DestinoConsumo = CODINT%(CodigoExternoDestino$)
            !CodExte_DestinoConsumo = CodigoExternoDestino$
            !Material = Material$
            !Observaciones = Observaciones$
            !Modelo = Modelo$
            !MARCA = MARCA$
            !NumeroSerieOrigen = NumeroSerieOrigen$
            !ClienteDestino = ClienteDestino%
            !Seteo = Seteo$
            !Alambre = Alambre$
            !Status = 0
            !ORDEN = i&
            .Update
            
    End With
    On Error Resume Next
    RstTrazarece.Close
    Set RstTrazarece = Nothing
    On Error GoTo 0
   
    ' IMPRIME FORMULARIO DE CALIDAD ASOCIADO A LA FAMILIA
    
    Dim DatosNumeroSerie As DatosNumeroSerie
    
    ' SE COMPLETA CON LOS DATOS VARIABLES DEL FORMULARIO DE CALIDAD (HOJA DEL ARCHIVO EXCEL)
    With DatosNumeroSerie
        .NumeroSerie = NumeroSerie&
        .Codigo = Mid$(CodigoExterno$, 1, Len(CodigoExterno$) - 2)
        .Descripcion = Descripcion$
        .Proveedor = ProveedorRecepcion$
        .NumeroRecepcion = CLng(NumeroRecepcion$)
        .Colada = NumeroColada&
    End With
    
    Call ImprimeFormulariosDeCalidad(Familia$, DatosNumeroSerie)
    
    GoTo 10

End Sub


Private Sub ImprimeFormulariosDeCalidad(Familia As String, DatosNumeroSerie As DatosNumeroSerie)
       
    If Familia = "" Then Exit Sub
    DescripcionFamilia$ = ECOARCH$("TAFAMIL", Familia)
    
    ' LISTA DE FORMULARIOS DE CALIDAD ASOCIADOS A LA FAMILIA
    ' (CADA FAMILIA TIENE UN ARCHIVO EXCEL, Y DENTRO DE ESTE, TANTAS HOJAS COMO FORMULARIOS)
    
    CAMPOS$ = "Cod/F3;Familia/A24;Formulario/A24;Ruta/A0"
    Call CARGA_ENCABEZADO(FRMZELECHO.msf2, CAMPOS$, FRMZELECHO)
    Call CARGA_ENCABEZADO(FRMZELECHO.MSF1, CAMPOS$, FRMZELECHO)
    FRMZELECHO.Caption = "LISTA DE FORMULARIOS POR FAMILIA"

    Dim rst As ADODB.Recordset
    ConsultaSql$ = "SELECT CAMPO1, CAMPO3, CAMPO4 FROM DATASQL WITH(NOLOCK) WHERE CAMPO2 = '" & Familia$ & "'"
    Set rst = READSET(ConsultaSql$)
    
    With rst
        Do While Not .EOF
            FRMZELECHO.msf2.Rows = FRMZELECHO.msf2.Rows + 1
            FRMZELECHO.msf2.TextMatrix(FRMZELECHO.msf2.Rows - 1, 1) = XVALO(!CAMPO1)
            FRMZELECHO.msf2.TextMatrix(FRMZELECHO.msf2.Rows - 1, 2) = DescripcionFamilia$
            FRMZELECHO.msf2.TextMatrix(FRMZELECHO.msf2.Rows - 1, 3) = !CAMPO4
            FRMZELECHO.msf2.TextMatrix(FRMZELECHO.msf2.Rows - 1, 4) = !CAMPO3
            
            .MoveNext
        Loop
    End With
    On Error Resume Next
    rst.Close
    Set rst = Nothing
    On Error GoTo 0
    
    FRMZELECHO.Width = 8500
    FRMZELECHO.Show 1
    NombreHoja$ = RESP_ZELE_VECT(3)
    If NombreHoja$ = "" Then Exit Sub
    RutaFormularioCalidad$ = RESP_ZELE_VECT(4)
    
    ' ABRE EL ARCHIVO EXCEL ASOCIADO A LA FAMILIA Y SE UBICA EN LA HOJA ELEGIDA (FORMULARIO DE CALIDAD)
    ' SI LA VARIABLE DATOSNUMEROSERIE VIENE CON INFORMACION, SE GRABAN ESOS DATOS EN LA HOJA (FORMULARIO) EXCEL Y LUEGO SE IMPRIME.
    
    Dim ObjetoExcel As Excel.Application
    Dim Libro As Excel.Workbook
    Dim Hoja As Excel.Worksheet

    Set ObjetoExcel = CreateObject("Excel.Application")
    With ObjetoExcel
        ' ARCHIVO EXCEL / ABRE EL LIBRO
        Set Libro = ObjetoExcel.Workbooks.Open(RutaFormularioCalidad)
        If DatosNumeroSerie.NumeroSerie > 0 Then
            Libro.Worksheets(NombreHoja).Cells(3, 26) = DatosNumeroSerie.NumeroSerie
            Libro.Worksheets(NombreHoja).Cells(3, 27) = DatosNumeroSerie.Codigo
            Libro.Worksheets(NombreHoja).Cells(3, 28) = DatosNumeroSerie.Descripcion
            Libro.Worksheets(NombreHoja).Cells(3, 29) = DatosNumeroSerie.Proveedor
            Libro.Worksheets(NombreHoja).Cells(3, 30) = DatosNumeroSerie.NumeroRecepcion
            Libro.Worksheets(NombreHoja).Cells(3, 31) = DatosNumeroSerie.Colada
            ObjetoExcel.ActiveWorkbook.Save
        End If
        Set Hoja = Libro.Sheets(NombreHoja)
        Hoja.PrintOut Copies:=1, Collate:=True
    End With

    ' CIERRA HOJA, LIBRO Y EXCEL
99  On Error Resume Next
    Set Hoja = Nothing
    Libro.Close: Set Libro = Nothing
    ObjetoExcel.Close: Set ObjetoExcel = Nothing
    On Error GoTo 0

End Sub


Private Sub mnuReimprimirFormularioDeCalidad_Click()

    ' LISTA DE MATERIALES RECIBIDOS CON NUMERO DE SERIE
    
10  TituloLista$ = "MATERIALES RECIBIDOS CON NUMERO DE SERIE"
    CAMPOS$ = "N-Serie/A10;Código/A24;Descripción/A36;Recepción/A10;Colada/A24;Fam/A3"
    Call CARGA_ENCABEZADO(FRMZELECHO.msf2, CAMPOS$, FRMZELECHO)
    Call CARGA_ENCABEZADO(FRMZELECHO.MSF1, CAMPOS$, FRMZELECHO)
    FRMZELECHO.Caption = TituloLista$
    
    SQLX$ = "SELECT * FROM TRAZARECE WITH(NOLOCK) WHERE NUMEROSERIE > 0 AND STATUS < 8"
    Set RS = READSET(SQLX$)
    J& = 0
    With RS
        If Not (.EOF And .BOF) Then
            FIN& = RS.RecordCount
            FRMZELECHO.msf2.Rows = FIN& + 1 ' ESTABLECE DE ANTEMANO LA CANTIDAD DE FILAS A UTILIZAR
        End If
        Do While Not .EOF
          IDLOTE$ = !IDENLOTE
          CodigoExterno$ = !COD_EXTE
          Descripcion$ = !Descrip
          NumeroSerie$ = CStr(!NumeroSerie)
          NumeroColada$ = CStr(!NumeroColada)
          Familia$ = !CFAMILIA
          
          J& = J& + 1
          Call TRACE(24, J&, FIN&)
          FRMZELECHO.msf2.TextMatrix(J&, 1) = NumeroSerie$
          FRMZELECHO.msf2.TextMatrix(J&, 2) = CodigoExterno$
          FRMZELECHO.msf2.TextMatrix(J&, 3) = Descripcion$
          FRMZELECHO.msf2.TextMatrix(J&, 4) = IDLOTE$
          FRMZELECHO.msf2.TextMatrix(J&, 5) = NumeroColada$
          FRMZELECHO.msf2.TextMatrix(J&, 6) = Familia$
          .MoveNext
        Loop
    End With
    On Error Resume Next
    RS.Close
    Set RS = Nothing
    On Error GoTo 0
    
    FRMZELECHO.Width = 11000
60  FRMZELECHO.Show 1
    
    NumeroSerie$ = TRIM(RESP_ZELE_VECT(1))
    If NumeroSerie$ = "" Then Exit Sub
    CodigoExterno$ = TRIM(RESP_ZELE_VECT(2))
    Descripcion$ = TRIM(RESP_ZELE_VECT(3))
    IDLOTE$ = TRIM(RESP_ZELE_VECT(4))
    NumeroColada$ = TRIM(RESP_ZELE_VECT(5))
    Familia$ = TRIM(RESP_ZELE_VECT(6))
    
    ' IMPRIME FORMULARIO DE CALIDAD ASOCIADO A LA FAMILIA
    
    Dim DatosNumeroSerie As DatosNumeroSerie
    
    ' SE COMPLETA CON LOS DATOS VARIABLES DEL FORMULARIO DE CALIDAD (HOJA DEL ARCHIVO EXCEL)
    ProveedorRecepcion$ = SAFETEXT$(VALOBADA("BOLRECEP", "RASO_PROV", "IDENLOTE='" & IDLOTE$ & "'"))
    NumeroRecepcion& = XVALO(VALOBADA("BOLRECEP", "NUMEBORE", "IDENLOTE='" & IDLOTE$ & "'"))
    With DatosNumeroSerie
        .NumeroSerie = CLng(NumeroSerie$)
        .Codigo = Mid$(CodigoExterno$, 1, Len(CodigoExterno$) - 2)
        .Descripcion = Descripcion$
        .Proveedor = ProveedorRecepcion$
        .NumeroRecepcion = NumeroRecepcion&
        .Colada = CLng(NumeroColada$)
    End With
    
    Call ImprimeFormulariosDeCalidad(Familia$, DatosNumeroSerie)
    
    GoTo 10

End Sub

Private Sub mnuReportesVarios_Click()

    Call ACONSUL_PRE(TRIM$(UCase$(Name)))

End Sub

Private Sub mnuRevertirAnulacionNumeroSerie_Click()

    Call RevertirBajaAnulacionDeNumeroDeSerie(2)
    
End Sub

Private Sub mnuRevertirBajaNumeroSerie_Click()

    Call RevertirBajaAnulacionDeNumeroDeSerie(1)
    
End Sub

Private Sub mnuSalir_Click()

    Call CIERRARCH("*.*")
    Call FINAL("")
    
End Sub

