VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   11340
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   12465
   LinkTopic       =   "Form1"
   ScaleHeight     =   11340
   ScaleWidth      =   12465
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   525
      Left            =   6600
      MultiLine       =   -1  'True
      TabIndex        =   52
      Top             =   6120
      Width           =   5655
   End
   Begin VB.Frame Frame5 
      Caption         =   "Base de datos"
      Height          =   3735
      Left            =   7200
      TabIndex        =   43
      Top             =   1920
      Width           =   4815
      Begin VB.CommandButton Command6 
         Caption         =   "Agregar Campos"
         Height          =   375
         Left            =   2520
         TabIndex        =   51
         Top             =   3240
         Width           =   1815
      End
      Begin VB.ListBox listaTablas 
         Height          =   2595
         Left            =   120
         TabIndex        =   23
         Top             =   480
         Width           =   2055
      End
      Begin VB.ListBox listaCampos 
         Height          =   2595
         Left            =   2520
         TabIndex        =   24
         Top             =   480
         Width           =   2055
      End
      Begin VB.Label Label7 
         Caption         =   "Tablas"
         Height          =   255
         Left            =   120
         TabIndex        =   45
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label8 
         Caption         =   "Campos"
         Height          =   255
         Left            =   2520
         TabIndex        =   44
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Generar codigo"
      Height          =   375
      Left            =   10920
      TabIndex        =   26
      Top             =   10320
      Width           =   1335
   End
   Begin VB.Frame Frame3 
      Caption         =   "Configuracion de la columna"
      Height          =   4935
      Left            =   240
      TabIndex        =   30
      Top             =   5760
      Width           =   6255
      Begin VB.TextBox txtCampoNombre 
         Height          =   285
         Left            =   1080
         TabIndex        =   6
         Top             =   360
         Width           =   3255
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Limpiar datos"
         Height          =   375
         Left            =   4680
         TabIndex        =   22
         Top             =   3960
         Width           =   1335
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Quitar campo"
         Height          =   375
         Left            =   4680
         TabIndex        =   21
         Top             =   1320
         Width           =   1335
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Nuevo campo"
         Height          =   375
         Left            =   4680
         TabIndex        =   19
         Top             =   360
         Width           =   1335
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Guardar cambios"
         Height          =   375
         Left            =   4680
         TabIndex        =   20
         Top             =   840
         Width           =   1335
      End
      Begin VB.Frame Frame2 
         Caption         =   "Tipo filtro"
         Height          =   1095
         Left            =   120
         TabIndex        =   39
         Top             =   720
         Width           =   4215
         Begin VB.TextBox Text3 
            Height          =   285
            Left            =   3360
            TabIndex        =   8
            Top             =   360
            Width           =   615
         End
         Begin VB.TextBox txtFiltroTitulo 
            Height          =   285
            Left            =   600
            TabIndex        =   7
            Top             =   360
            Width           =   1695
         End
         Begin VB.TextBox txtFiltroEco 
            Height          =   285
            Left            =   600
            TabIndex        =   9
            Top             =   720
            Width           =   1695
         End
         Begin VB.Label Label19 
            Caption         =   "Formato"
            Height          =   255
            Left            =   2520
            TabIndex        =   50
            Top             =   360
            Width           =   975
         End
         Begin VB.Label Label5 
            Caption         =   "Titulo"
            Height          =   255
            Left            =   120
            TabIndex        =   41
            Top             =   360
            Width           =   615
         End
         Begin VB.Label Label11 
            Caption         =   "ECO"
            Height          =   255
            Left            =   120
            TabIndex        =   40
            Top             =   720
            Width           =   615
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Tipo tabulado"
         Height          =   2895
         Left            =   120
         TabIndex        =   31
         Top             =   1920
         Width           =   4215
         Begin VB.TextBox txtCampoEcoFormato 
            Height          =   285
            Left            =   2880
            TabIndex        =   13
            Top             =   840
            Width           =   615
         End
         Begin VB.TextBox txtCampoFormato 
            Height          =   285
            Left            =   3360
            TabIndex        =   11
            Top             =   480
            Width           =   615
         End
         Begin VB.TextBox txtCampoTitulo 
            Height          =   285
            Left            =   960
            TabIndex        =   10
            Top             =   480
            Width           =   1455
         End
         Begin VB.TextBox txtCampoEco 
            Height          =   285
            Left            =   960
            TabIndex        =   12
            Top             =   840
            Width           =   1095
         End
         Begin VB.TextBox txtCampoFuncion 
            Height          =   285
            Left            =   960
            TabIndex        =   14
            Top             =   1200
            Width           =   1455
         End
         Begin VB.TextBox txtCampoValobada 
            Height          =   285
            Left            =   960
            TabIndex        =   15
            Top             =   1560
            Width           =   1455
         End
         Begin VB.CheckBox chkCampoCControl 
            Height          =   195
            Left            =   1440
            TabIndex        =   16
            Top             =   2040
            Width           =   255
         End
         Begin VB.CheckBox chkCampoTotal 
            Height          =   195
            Left            =   1440
            TabIndex        =   18
            Top             =   2400
            Width           =   255
         End
         Begin VB.TextBox txtCampoCCCol 
            Height          =   285
            Left            =   3480
            TabIndex        =   17
            Top             =   2040
            Width           =   495
         End
         Begin VB.Label Label17 
            Caption         =   "Formato"
            Height          =   255
            Left            =   2160
            TabIndex        =   48
            Top             =   840
            Width           =   975
         End
         Begin VB.Label Label10 
            Caption         =   "Formato"
            Height          =   255
            Left            =   2520
            TabIndex        =   47
            Top             =   480
            Width           =   975
         End
         Begin VB.Label Label4 
            Caption         =   "Titulo"
            Height          =   255
            Left            =   120
            TabIndex        =   38
            Top             =   480
            Width           =   975
         End
         Begin VB.Label Label6 
            Caption         =   "ECO"
            Height          =   255
            Left            =   120
            TabIndex        =   37
            Top             =   840
            Width           =   975
         End
         Begin VB.Label Label12 
            Caption         =   "Función"
            Height          =   255
            Left            =   120
            TabIndex        =   36
            Top             =   1200
            Width           =   975
         End
         Begin VB.Label Label13 
            Caption         =   "Valobada"
            Height          =   255
            Left            =   120
            TabIndex        =   35
            Top             =   1560
            Width           =   975
         End
         Begin VB.Label Label14 
            Caption         =   "Corte de Control"
            Height          =   255
            Left            =   120
            TabIndex        =   34
            Top             =   2040
            Width           =   1215
         End
         Begin VB.Label Label15 
            Caption         =   "Totalizador"
            Height          =   255
            Left            =   120
            TabIndex        =   33
            Top             =   2400
            Width           =   1215
         End
         Begin VB.Label Label16 
            Caption         =   "Nº Columna de corte"
            Height          =   255
            Left            =   1800
            TabIndex        =   32
            Top             =   2040
            Width           =   1575
         End
      End
      Begin VB.Label Label18 
         Caption         =   "Campo"
         Height          =   255
         Left            =   240
         TabIndex        =   49
         Top             =   360
         Width           =   615
      End
   End
   Begin VB.TextBox txtCodigoVB 
      Height          =   3135
      Left            =   6600
      MultiLine       =   -1  'True
      TabIndex        =   25
      Top             =   7080
      Width           =   5655
   End
   Begin MSFlexGridLib.MSFlexGrid mfgReporteColumnas 
      Height          =   2775
      Left            =   240
      TabIndex        =   5
      Top             =   2880
      Width           =   6855
      _ExtentX        =   12091
      _ExtentY        =   4895
      _Version        =   393216
   End
   Begin VB.TextBox txtReporteDescripcion 
      Height          =   285
      Left            =   1320
      TabIndex        =   4
      Top             =   1680
      Width           =   5535
   End
   Begin VB.TextBox txtReporteNombre 
      Height          =   285
      Left            =   1320
      TabIndex        =   3
      Top             =   1200
      Width           =   2535
   End
   Begin VB.Frame Frame1 
      Caption         =   "Tipo de reporte"
      Height          =   735
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   3255
      Begin VB.OptionButton optPorCodigo 
         Caption         =   "Carga por codigo"
         Height          =   255
         Left            =   1440
         TabIndex        =   2
         Top             =   360
         Width           =   1575
      End
      Begin VB.OptionButton optDirecto 
         Caption         =   "Directo"
         Height          =   255
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   855
      End
   End
   Begin VB.Label Label20 
      Caption         =   "Condicion Sql (Solo modo Directo)"
      Height          =   255
      Left            =   6600
      TabIndex        =   53
      Top             =   5880
      Width           =   3135
   End
   Begin VB.Label TEPRUEBA 
      AutoSize        =   -1  'True
      Caption         =   "TEPRUEBA"
      Height          =   195
      Left            =   11400
      TabIndex        =   46
      Top             =   120
      Visible         =   0   'False
      Width           =   870
   End
   Begin VB.Label Label9 
      Caption         =   "Codigo VB"
      Height          =   255
      Left            =   6600
      TabIndex        =   42
      Top             =   6840
      Width           =   2055
   End
   Begin VB.Label Label3 
      Caption         =   "Lista de columnas"
      Height          =   255
      Left            =   240
      TabIndex        =   29
      Top             =   2520
      Width           =   1455
   End
   Begin VB.Label Label2 
      Caption         =   "Descripcion"
      Height          =   255
      Left            =   240
      TabIndex        =   28
      Top             =   1800
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Nombre"
      Height          =   255
      Left            =   240
      TabIndex        =   27
      Top             =   1320
      Width           =   735
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

    NumeroColumna& = mfgReporteColumnas.Row
    ActualizarDatosColumna (NumeroColumna&)

End Sub

Private Sub Command2_Click()

    txtCodigoVB = ""
    
    If TRIM$(txtReporteNombre) = "" Or TRIM$(txtReporteDescripcion) = "" Then
        Call COMUNI("Falta completar nombre o descripcion del reporte")
        Exit Sub
    End If
    
    If mfgReporteColumnas.Rows < 2 Then
        Call COMUNI("Falta ingresar las columnas del reporte")
        Exit Sub
    End If
    
    If (optDirecto.Value = True And TRIM$(txtCondicionSql) = "") Then
        Call COMUNI("Falta completar la condicion sql")
        Exit Sub
    End If

    CodigoVB$ = ""
    Comillas = Chr$(34)
    NuevaLineaDeCodigo CodigoVB$, "Dim oReporte As cGenrep"
    NuevaLineaDeCodigo CodigoVB$, ""
    NuevaLineaDeCodigo CodigoVB$, "Dim RST As ADODB.Recordset"
    NuevaLineaDeCodigo CodigoVB$, ""
    
    NuevaLineaDeCodigo CodigoVB$, "Set oReporte = New cGenrep"
    NuevaLineaDeCodigo CodigoVB$, ""
    Parametros$ = ""
    TotalCampos& = mfgReporteColumnas.Rows - 1
    For f& = 1 To TotalCampos&
            
            CampoNombre$ = TRIM$(mfgReporteColumnas.TextMatrix(f&, 9))
            Tipo$ = TRIM$(mfgReporteColumnas.TextMatrix(f&, 2))
            If Tipo$ = "FILTRO" Then
                
                FilTitulo$ = TRIM$(mfgReporteColumnas.TextMatrix(f&, 1))
                If FilTitulo$ = "" Then FilTitulo$ = CampoNombre$
                FilFormato$ = TRIM$(mfgReporteColumnas.TextMatrix(f&, 3))
                FilEco$ = TRIM$(mfgReporteColumnas.TextMatrix(f&, 4))
                
                SeparadorDeCampo$ = "": If Parametros$ <> "" Then SeparadorDeCampo$ = ";"
                Parametros$ = Parametros$ & SeparadorDeCampo$ & CampoNombre$ & "(FILTRO=" & FilTitulo$ & "-" & FilEco$ & ")/" & FilFormato$
                
            ElseIf Tipo$ = "COLUMNA" Then
            
                CampoTitulo$ = TRIM$(mfgReporteColumnas.TextMatrix(f&, 1))
                If CampoTitulo$ = "" Then CampoTitulo$ = CampoNombre$
                CampoFormato$ = TRIM$(mfgReporteColumnas.TextMatrix(f&, 3))
                CampoEco$ = TRIM$(mfgReporteColumnas.TextMatrix(f&, 4))
                
                n = InStr(1, mfgReporteColumnas.TextMatrix(f&, 7), "/")
                CampoCC% = 0
                If n > 0 Then
                    CampoCC% = XVALO(Mid$(mfgReporteColumnas.TextMatrix(f&, 7), 1, n - 1))
                    CampoCCCol% = XVALO(Mid$(mfgReporteColumnas.TextMatrix(f&, 7), n + 1))
                End If
                CampoTotal% = XVALO(mfgReporteColumnas.TextMatrix(f&, 8))
                
                TextoLinea$ = CampoNombre$ & "(T=" & CampoTitulo$
                If CampoCC% = 1 Then TextoLinea$ = TextoLinea$ & ";CC=" & CStr(CampoCCCol%)
                If CampoTotal% = 1 Then TextoLinea$ = TextoLinea$ & ";TT=SI"
                TextoLinea$ = TextoLinea$ & ")/" & CampoFormato$
                
                SeparadorDeCampo$ = "": If Parametros$ <> "" Then SeparadorDeCampo$ = ";"
                Parametros$ = Parametros$ & SeparadorDeCampo$ & TextoLinea$
                
            End If
            
        Next f&
    
        If optPorCodigo.Value = True Then
            
            TextoLinea$ = "Set RST = oReporte.GeneraReporteConTablaTemporal(" & Comillas & txtReporteNombre & Comillas & "," & Comillas & txtReporteDescripcion & Comillas & ","
            LineaCodigo$ = TextoLinea$ & Comillas & Parametros$ & Comillas & ")"
            NuevaLineaDeCodigo CodigoVB$, LineaCodigo$
            NuevaLineaDeCodigo CodigoVB$, ""
                 
            NuevaLineaDeCodigo CodigoVB$, "With RST"
            NuevaLineaDeCodigo CodigoVB$, "     .AddNew"
            For f& = 1 To TotalCampos&
                NuevaLineaDeCodigo CodigoVB$, "     .Fields(" & CStr(f&) & ") = "
            Next f&
            NuevaLineaDeCodigo CodigoVB$, "     .Update"
            NuevaLineaDeCodigo CodigoVB$, "End With"
        
        ElseIf optDirecto.Value = True Then
        
            CondicionSql$ = txtCondicionSql
            TextoLinea$ = "Resultado = oReporte.GeneraReporte(" & Comillas & txtReporteNombre & Comillas & "," & Comillas & txtReporteDescripcion & Comillas & ","
            LineaCodigo$ = TextoLinea$ & Comillas & Parametros$ & Comillas & "," & Comillas & CondicionSql$ & Comillas & ")"
            NuevaLineaDeCodigo CodigoVB$, LineaCodigo$
            NuevaLineaDeCodigo CodigoVB$, ""
            
        End If
        
        NuevaLineaDeCodigo CodigoVB$, ""
        NuevaLineaDeCodigo CodigoVB$, "oReporte.MostrarReporte"
        
        txtCodigoVB = CodigoVB$
    
End Sub

Sub NuevaLineaDeCodigo(Linea$, Texto$)

    Linea$ = Linea$ & TRIM$(Texto$) & vbCrLf
    
End Sub

Private Sub Command3_Click()

    ActualizarDatosColumna (0)
    
End Sub


Private Sub Command5_Click()

    InicializarVista
    
End Sub

Private Sub Command6_Click()

    If listaCampos.ListCount < 1 Then Exit Sub
    
    If mfgReporteColumnas.Rows > 1 Then
        opcion% = COMALTER%("La lista de columnas tiene información.\ \¿Desea sobreescribir la misma\con todos los campos de la tabla activa?\\No, Cancelar\Si, Continuar")
        If opcion% = 1 Then Exit Sub
    End If
    
    mfgReporteColumnas.Rows = 1
    For i& = 1 To listaCampos.ListCount - 1
        mfgReporteColumnas.Rows = mfgReporteColumnas.Rows + 1
        f& = mfgReporteColumnas.Rows - 1
        
        mfgReporteColumnas.TextMatrix(f&, 1) = listaCampos.List(i&)
        mfgReporteColumnas.TextMatrix(f&, 2) = "COLUMNA"
        mfgReporteColumnas.TextMatrix(f&, 9) = listaCampos.List(i&)
    Next i&
    
End Sub

Private Sub Form_Load()
  
    Call ABRECONEXION

    Call MUESTRATABLAS
    
    Call CARGA_ENCABEZADO(mfgReporteColumnas, "Titulo/A20;Tipo/A5;Formato/A8;ECO/A15;Funcion/A15;Valobada/A15;CC/A2;Total/A5;Campo/A25", Me)
    
End Sub


Sub MUESTRATABLAS()
    listaTablas.Clear: listaCampos.Clear
    Dim AUX As ADODB.Recordset
    Set AUX = New ADODB.Recordset
22  SQLX$ = "Select name from sysobjects WHERE LEFT(NAME,3) <> 'dt_' and LEFT(NAME,3) <> 'pk_' and LEFT(NAME,3) <> 'sys'"
    Set AUX = FLEXCONN.Execute(FILTSQL$(SQLX$))
    Do While Not AUX.EOF
       If InStr(UCase$(AUX!Name), "DTPROPERT") > 0 Then GoTo 29
       listaTablas.AddItem UCase$(AUX!Name)
29  AUX.MoveNext
    Loop
End Sub


Private Sub listaTablas_DblClick()

    If listaTablas.ListIndex < 0 Then Exit Sub
    
    Dim AUX As ADODB.Recordset
    Set AUX = New ADODB.Recordset
    listaCampos.Clear
    TAPRIN$ = TRIM$(listaTablas.List(listaTablas.ListIndex))
    
    On Error GoTo controlError
    Set AUX = FLEXCONN.Execute(FILTSQL$("select * from " & TAPRIN$))
    For i& = 1 To AUX.Fields.Count
        If UCase$(Left$(AUX.Fields(i& - 1).Name, 8)) <> "VALOBADA" Then
           listaCampos.AddItem PREFPRIN$ + AUX.Fields(i& - 1).Name
        End If
    Next i&
    On Error GoTo 0
    Exit Sub
    
controlError:
    Exit Sub

End Sub


Sub InicializarVista()

    txtCampoNombre = ""
    txtFiltroTitulo = ""
    txtFiltroEco = ""
    txtCampoFormato = ""
    txtCampoTitulo = ""
    txtCampoEco = ""
    txtCampoEcoFormato = ""
    txtCampoFuncion = ""
    txtCampoValobada = ""
    txtCampoCCCol = ""
    
    chkCampoCControl.Value = 0
    chkCampoTotal.Value = 0

End Sub


Private Sub mfgReporteColumnas_DblClick()

    InicializarVista
    f = mfgReporteColumnas.Row
    CampoNombre$ = mfgReporteColumnas.TextMatrix(f, 9)
    txtCampoNombre = CampoNombre$
    Tipo = mfgReporteColumnas.TextMatrix(f, 2)
    
    If Tipo = "FILTRO" Then
        txtFiltroTitulo = mfgReporteColumnas.TextMatrix(f, 1)
        txtFiltroEco = mfgReporteColumnas.TextMatrix(f, 4)
    ElseIf Tipo = "COLUMNA" Then
        txtCampoTitulo = mfgReporteColumnas.TextMatrix(f, 1)
        txtCampoFormato = mfgReporteColumnas.TextMatrix(f, 3)
        
        TextoEco$ = TRIM$(mfgReporteColumnas.TextMatrix(f, 4))
        If TextoEco$ <> "" Then
            n = InStr(1, TextoEco$, "/")
            txtCampoEco = Mid$(mfgReporteColumnas.TextMatrix(f, 4), 1, n - 1)
            txtCampoEcoFormato = Mid$(mfgReporteColumnas.TextMatrix(f, 4), n + 1)
        End If
        txtCampoFuncion = mfgReporteColumnas.TextMatrix(f, 5)
        txtCampoValobada = mfgReporteColumnas.TextMatrix(f, 6)
        
        TextoCCCol = TRIM$(mfgReporteColumnas.TextMatrix(f, 7))
        If TextoCCCol <> "" Then
            n = InStr(1, TextoCCCol, "/")
            chkCampoCControl.Value = XVALO(Mid$(mfgReporteColumnas.TextMatrix(f, 7), 1, n - 1))
            txtCampoCCCol = Mid$(mfgReporteColumnas.TextMatrix(f, 7), n + 1)
        End If
        chkCampoTotal.Value = XVALO(mfgReporteColumnas.TextMatrix(f, 8))
    End If
End Sub

Function ValidaCampoNombre(Texto As String) As Boolean

    ValidaCampoNombre = True
    Textoi$ = Texto
    For i& = 1 To Len(Textoi$)
        n% = Asc(Mid$(Textoi$, i&, 1))
        If Not ((n% >= 48 And n% <= 57) Or (n% >= 65 And n% <= 90) Or (n% >= 97 And n% <= 122)) Then
            ValidaCampoNombre = False
            Exit Function
        End If
    Next i&

End Function

Sub ActualizarDatosColumna(NumeroColumna As Long)

    ' VALIDACIONES
    If TRIM$(txtCampoNombre) = "" Then
        Call COMUNI("Falta ingresar nombre del campo.")
        Exit Sub
    End If
    
    If TRIM$(txtFiltroTitulo) = "" And TRIM$(txtCampoTitulo) = "" Then
        Call COMUNI("Falta ingresar titulo de la columna.")
        Exit Sub
    End If
    
    If TRIM$(txtFiltroTitulo) <> "" And TRIM$(txtCampoTitulo) <> "" Then
        Call COMUNI("La columna no puede ser simultaneamente\del tipo filtro y tabulada.")
        Exit Sub
    End If
    
    If chkCampoCControl.Value = 1 Then
        CCCol% = XVALO(txtCampoCCCol)
        If CCCol% = 0 Or CCCol% > mfgReporteColumnas.Rows - 1 Then
            Call COMUNI("Falta ingresar numero de columna para el corte de control\o el numero de columna no es valido.")
            Exit Sub
        End If
    End If
    
    ' EN MODO EDICION, LIMPIA LAS CELDAS
    If NumeroColumna > 0 Then
        For c& = 1 To mfgReporteColumnas.Cols - 1
            mfgReporteColumnas.TextMatrix(NumeroColumna, c&) = ""
        Next c&
    End If
    
    ' MODO FILTRO
    If TRIM$(txtFiltroTitulo) <> "" And TRIM$(txtCampoTitulo) = "" Then
    
        If TRIM$(txtCampoNombre) = "" Then txtCampoNombre = txtFiltroTitulo
        If ValidaCampoNombre(txtCampoNombre) = False Then
            Call COMUNI("Caracteres no admisibles para el nombre del campo")
            Exit Sub
        End If
        If NumeroColumna > 0 Then
            mfgReporteColumnas.Row = NumeroColumna
            f = NumeroColumna
        Else
            mfgReporteColumnas.Rows = mfgReporteColumnas.Rows + 1
            f = mfgReporteColumnas.Rows - 1
        End If
        mfgReporteColumnas.TextMatrix(f, 1) = TRIM$(txtFiltroTitulo)
        mfgReporteColumnas.TextMatrix(f, 2) = "FILTRO"
        mfgReporteColumnas.TextMatrix(f, 3) = TRIM$(txtCampoFormato)
        mfgReporteColumnas.TextMatrix(f, 4) = TRIM$(txtFiltroEco)
        mfgReporteColumnas.TextMatrix(f, 9) = TRIM$(txtCampoNombre)
        
        Call InicializarVista
    End If
    
    
    ' MODO COLUMNA
    If TRIM$(txtCampoTitulo) <> "" And TRIM$(txtFiltroTitulo) = "" Then
    
        If TRIM$(txtCampoNombre) = "" Then txtCampoNombre = txtCampoTitulo
        If ValidaCampoNombre(txtCampoNombre) = False Then
            Call COMUNI("Caracteres no admisibles para el nombre del campo")
            Exit Sub
        End If
        If NumeroColumna > 0 Then
            mfgReporteColumnas.Row = NumeroColumna
            f = NumeroColumna
        Else
            mfgReporteColumnas.Rows = mfgReporteColumnas.Rows + 1
            f = mfgReporteColumnas.Rows - 1
        End If
        mfgReporteColumnas.TextMatrix(f, 1) = TRIM$(txtCampoTitulo)
        mfgReporteColumnas.TextMatrix(f, 2) = "COLUMNA"
        mfgReporteColumnas.TextMatrix(f, 3) = TRIM$(txtCampoFormato)
        If txtCampoEco <> "" Then mfgReporteColumnas.TextMatrix(f, 4) = TRIM$(txtCampoEco) & "/" & TRIM$(txtCampoEcoFormato)
        mfgReporteColumnas.TextMatrix(f, 5) = TRIM$(txtCampoFuncion)
        mfgReporteColumnas.TextMatrix(f, 6) = TRIM$(txtCampoValobada)
        If chkCampoCControl.Value = 1 Then mfgReporteColumnas.TextMatrix(f, 7) = TRIM$(chkCampoCControl) & "/" & TRIM$(txtCampoCCCol)
        mfgReporteColumnas.TextMatrix(f, 8) = TRIM$(chkCampoTotal)
        mfgReporteColumnas.TextMatrix(f, 9) = TRIM$(txtCampoNombre)
        
        InicializarVista
    End If
End Sub
