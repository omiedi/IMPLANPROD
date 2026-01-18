VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form PLANICOS_TB 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00E2D3C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Planilla de Costos de Producción"
   ClientHeight    =   9285
   ClientLeft      =   45
   ClientTop       =   600
   ClientWidth     =   11355
   ClipControls    =   0   'False
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9285
   ScaleWidth      =   11355
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   3960
      TabIndex        =   1
      Top             =   240
      Width           =   175
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   10680
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSFlexGridLib.MSFlexGrid GRID 
      Height          =   8445
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   11115
      _ExtentX        =   19606
      _ExtentY        =   14896
      _Version        =   393216
      Rows            =   1
      Cols            =   4
      FixedRows       =   0
      FixedCols       =   0
      BackColor       =   16777215
      BackColorFixed  =   16501405
      BackColorBkg    =   16777215
      FillStyle       =   1
      SelectionMode   =   1
      AllowUserResizing=   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   1320
      TabIndex        =   3
      Top             =   240
      Width           =   2655
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E2D3C0&
      Caption         =   "Mes y Año"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   -120
      TabIndex        =   2
      Top             =   310
      Width           =   1215
   End
   Begin VB.Menu Inicio 
      Caption         =   "Inicio"
      Begin VB.Menu Cerrar 
         Caption         =   "Cerrar"
      End
   End
   Begin VB.Menu ConsultasListados 
      Caption         =   "Consultas y Listados"
      Begin VB.Menu ExportarExcel 
         Caption         =   "Exportar a Excel"
      End
   End
End
Attribute VB_Name = "PLANICOS_TB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim FrmLeft, FrmTop, FrmWidth, FrmHeight As Long
Dim GridTop, GridLeft, GridWidth, GridHeight As Long
Dim FRAMETop, FRAMELeft, FRAMEWidth, FRAMEHeight As Long
Dim Com1Top, Com2Top, Com2Left As Long
Dim Com1LEFT As Long
Public APROBO%
Dim LecturaDatosOriginales As Byte

Dim UltimaFilaDelReporte As Long




Private Sub Cerrar_Click()
    
    Call CIERRARCH("*.*")
    Call FINAL("")

End Sub

Private Sub Command1_Click()

    Call SELECHO("SELMESA")
    Periodo$ = TRIM$(VALACT1$("SELMESA"))
    
    If Periodo$ = "" Then
        Label2 = ""
    Else
        Label2 = Periodo$
    End If
    
    
End Sub

Private Sub GeneraPlanillaCostosProduccion()

    Dim ArticulosFabricados() As Long
    Dim CostoArticulosFabricados() As Single

    Periodo$ = Label2
    If Periodo$ = "" Then Exit Sub
    
    ' LIMPIA LA GRILLA Y LA DEJA CON VALORES POR DEFAULT
    ' **************************************************
    GRID.Clear: GRID.Rows = 0: GRID.Cols = 0
    GRID.Rows = 2: GRID.Cols = 4
    

    ' OBTIENE DESDE Y HASTA DEL MES ELEGIDO
    ' *************************************
    Call MESANO$(Periodo$, DES%, HAS%)
    desde = FETEXCOR1(DES%)
    Hasta = FETEXCOR1(HAS%)
    
        
    ' CANTIDAD DE UNIDADES FABRICADAS CON COSTO POR UNIDAD Y TOTAL
    ' ************************************************************
    On Error GoTo ErrMaterialFabricado
    
    ReDim ArticulosFabricados(NUMAS)
    ReDim CostoArticulosFabricados(NUMAS)

    SQLX$ = "SELECT cod_inte, SUM(cantingre) AS CantidadFabricada FROM MOVISTO INNER JOIN MASTER WITH(NOLOCK)"
    SQLX$ = SQLX$ & " ON MOVISTO.cod_inte = MASTER.codint"
    SQLX$ = SQLX$ & " WHERE MASTER.timate=5 AND MOVISTO.codimovi='OF'"
    SQLX$ = SQLX$ & " AND MOVISTO.fechmov >= '" & desde & "' AND MOVISTO.fechmov <= '" & Hasta & "'"
    SQLX$ = SQLX$ + " GROUP BY MOVISTO.cod_inte"
    Dim MaterialFabricado As ADODB.Recordset
    Set MaterialFabricado = READSET(SQLX$)
    Dim CantidadTotalFabricada As Long
    Dim CostoTotalFabricado As Double
    Dim UnidadesConFuelle As Long
    Dim CantidadDeUnidadesPorOmnibus As Long
    
    With MaterialFabricado
        Do While Not .EOF
            CI% = XVALO(!cod_inte)
            '
            ArticulosFabricados(CI%) = XVALO(!CantidadFabricada)
            CantidadTotalFabricada = CantidadTotalFabricada + ArticulosFabricados(CI%)
            CostoArticulosFabricados(CI%) = XVALO(VALOBADA("HISTOCOSTOS", "costostandard", "periodo='" & Periodo$ & "' AND cod_inte=" & CStr(CI%)))
            CostoTotalFabricado = CostoTotalFabricado + (ArticulosFabricados(CI%) * CostoArticulosFabricados(CI%))
            
            ' VERIFICA SI SE TRATA DE UN OMNIBUS CON DOBLE CUERPO, YA QUE ES CONSIDERADO COMO DOS UNIDADES
            CantidadDeUnidadesPorOmnibus = XVALO(VALOBADA("MASTER", "Talle", "Codint=" & CStr(CI%)))
            If CantidadDeUnidadesPorOmnibus > 1 Then UnidadesConFuelle = UnidadesConFuelle + XVALO(!CantidadFabricada)
            .MoveNext
        Loop
    End With
    
    If CantidadTotalFabricada > 0 Then
        CostoPorUnidad = CostoTotalFabricado / CantidadTotalFabricada
    End If
    
    MaterialFabricado.Close
    Set MaterialFabricado = Nothing
    On Error GoTo 0
   
    
    ' DETALLE DE GASTOS
    ' *****************
    On Error GoTo ErrCargaDetalle
    
    'GRID.Rows = GRID.Rows + 1
    GRID.TextMatrix(GRID.Rows - 1, 0) = "Costo de Materiales Para 1 Unidad": GRID.ColWidth(0) = GRID.ColWidth(0) + 5000
        Call DisenioDeCelda(GRID, GRID.Rows - 1, 0, "", 0, True, False, False, False, vbWhite, &H808080)
    GRID.TextMatrix(GRID.Rows - 1, 1) = "$  " & TRIM$(FORMATNUM(CostoPorUnidad, "G14.2")): GRID.ColWidth(1) = GRID.ColWidth(1) + 500: GRID.ColAlignment(1) = 7
        Call DisenioDeCelda(GRID, GRID.Rows - 1, 1, "", 0, True, False, False, False, vbWhite, &H808080)

    ' EL PRIMER RENGLON DEL DETALLE SE CREA A PARTIR DE LOS COSTOS DE LAS UNIDADES FABRICADAS,
    ' CALCULADO EN LA CONSULTA SQL DE MAS ARRIBA
    GRID.Rows = GRID.Rows + 2
    GRID.TextMatrix(GRID.Rows - 1, 0) = "DESCRIPCION": Call DisenioDeCelda(GRID, GRID.Rows - 1, 0, "", 0, True, False, False, False, vbWhite, &H808080)
    GRID.TextMatrix(GRID.Rows - 1, 1) = "TOTAL": Call DisenioDeCelda(GRID, GRID.Rows - 1, 1, "", 0, True, False, False, False, vbWhite, &H808080)
    GRID.Rows = GRID.Rows + 1
    GRID.TextMatrix(GRID.Rows - 1, 0) = "Costo de materiales para unidades"
    GRID.TextMatrix(GRID.Rows - 1, 1) = "$  " & TRIM$(FORMATNUM(CostoTotalFabricado, "G14.2"))
        RangoFila$ = CStr(GRID.Rows - 1) & ":" & CStr(GRID.Rows - 1)
        Call DisenioDeCelda(GRID, RangoFila$, "0:1", "", 0, False, False, False, False, 0, &HEEEEDD)
        Call DisenioDeCelda(GRID, GRID.Rows - 1, 1, "", 0, True, False, False, False, 0, 0)
    
    SQLX$ = "SELECT cuecontai, SUM(idebecomp) AS idebe, SUM(ihabecomp) AS ihabe FROM CAJABANC WITH(NOLOCK)"
    SQLX$ = SQLX$ & " WHERE nuprove > 0 AND (tipomovim='compra' OR (tipomovim='opag' AND varimovim='cgast'))"
    SQLX$ = SQLX$ & " AND fechconta >= '" & desde & "' AND fechconta <= '" & Hasta & "'"
    SQLX$ = SQLX$ & " GROUP BY cuecontai"
    
    Dim ImporteGasto As Double
    Dim CuentaContable As Integer
    Dim TotalDeGastos As Double
    Dim GastosDelPeriodo As ADODB.Recordset
    Set GastosDelPeriodo = READSET(SQLX$)
    
    With GastosDelPeriodo
        Do While Not .EOF
            ImporteGasto = XVALO(!idebe) - XVALO(!ihabe)
            CuentaContable = XVALO(!cuecontai)
            TotalDeGastos = TotalDeGastos + ImporteGasto
            
            GRID.Rows = GRID.Rows + 1
            GRID.TextMatrix(GRID.Rows - 1, 0) = DENOCUE$(CuentaContable)
            GRID.TextMatrix(GRID.Rows - 1, 1) = "$  " & TRIM$(FORMATNUM$(ImporteGasto, "G14.2"))
                RangoFila$ = CStr(GRID.Rows - 1) & ":" & CStr(GRID.Rows - 1)
                Call DisenioDeCelda(GRID, RangoFila$, "0:1", "", 0, False, False, False, False, 0, &HEEEEDD)
                Call DisenioDeCelda(GRID, GRID.Rows - 1, 1, "", 0, True, False, False, False, 0, 0)
                
            .MoveNext
        Loop
    End With
    GastosDelPeriodo.Close
    Set GastosDelPeriodo = Nothing
    
    GRID.Rows = GRID.Rows + 1
    RangoFila$ = CStr(GRID.Rows - 1) & ":" & CStr(GRID.Rows - 1)
    Call DisenioDeCelda(GRID, RangoFila$, "0:1", "", 0, False, False, False, False, 0, &HEEEEDD)
    GRID.Rows = GRID.Rows + 1
    GRID.TextMatrix(GRID.Rows - 1, 0) = "Gasto Mensual Estimado de " & Periodo$
        Call DisenioDeCelda(GRID, GRID.Rows - 1, 0, "", 0, True, True, False, False, vbWhite, &H808080)
    GRID.TextMatrix(GRID.Rows - 1, 1) = "$  " & TRIM$(FORMATNUM$(CostoTotalFabricado + TotalDeGastos, "G14.2"))
        Call DisenioDeCelda(GRID, GRID.Rows - 1, 1, "", 0, True, False, False, False, vbWhite, &H808080)
    
    On Error GoTo 0
       
    
    ' PIE DE REPORTE
    ' **************
    GRID.Rows = GRID.Rows + 4
    GRID.TextMatrix(GRID.Rows - 1, 0) = "Cantidad de Unidades Fabricadas"
    GRID.TextMatrix(GRID.Rows - 1, 1) = FORMATNUM(CantidadTotalFabricada, "I10")
        RangoFila$ = CStr(GRID.Rows - 1) & ":" & CStr(GRID.Rows - 1)
        Call DisenioDeCelda(GRID, RangoFila$, "0:1", "", 0, True, False, False, False, 0, &HE0E0E0)
    GRID.Rows = GRID.Rows + 1
    If UnidadesConFuelle > 0 Then
        GRID.TextMatrix(GRID.Rows - 1, 0) = "Cantidad de Unidades Fabricadas (c/Articulados)"
        GRID.TextMatrix(GRID.Rows - 1, 1) = FORMATNUM(CantidadTotalFabricada + UnidadesConFuelle, "I10")
            RangoFila$ = CStr(GRID.Rows - 1) & ":" & CStr(GRID.Rows - 1)
            Call DisenioDeCelda(GRID, RangoFila$, "0:1", "", 0, True, False, False, False, 0, &HE0E0E0)
        GRID.Rows = GRID.Rows + 1
    End If
    
    GRID.Rows = GRID.Rows + 1
    GRID.TextMatrix(GRID.Rows - 1, 0) = "Descripción"
    GRID.TextMatrix(GRID.Rows - 1, 1) = "Unidades"
    GRID.TextMatrix(GRID.Rows - 1, 2) = "Costo": GRID.ColWidth(2) = GRID.ColWidth(2) + 500: GRID.ColAlignment(2) = 7
    GRID.TextMatrix(GRID.Rows - 1, 3) = "Costo Total": GRID.ColWidth(3) = GRID.ColWidth(3) + 500: GRID.ColAlignment(3) = 7
        RangoFila$ = CStr(GRID.Rows - 1) & ":" & CStr(GRID.Rows - 1)
        Call DisenioDeCelda(GRID, RangoFila$, "0:3", "", 0, True, False, False, False, 0, &HE0E0E0)
    For ix& = 1 To NUMAS
        If ArticulosFabricados(ix&) > 0.05 Then
            GRID.Rows = GRID.Rows + 1
            GRID.TextMatrix(GRID.Rows - 1, 0) = DESCRIT0$(CInt(ix&))
            GRID.TextMatrix(GRID.Rows - 1, 1) = FORMATNUM$(ArticulosFabricados(ix&), "I10")
            GRID.TextMatrix(GRID.Rows - 1, 2) = "$  " & TRIM$(FORMATNUM$(CostoArticulosFabricados(ix&), "G14.2"))
            GRID.TextMatrix(GRID.Rows - 1, 3) = "$  " & TRIM$(FORMATNUM$(ArticulosFabricados(ix&) * CostoArticulosFabricados(ix&), "G14.2"))
            RangoFila$ = CStr(GRID.Rows - 1) & ":" & CStr(GRID.Rows - 1)
            Call DisenioDeCelda(GRID, RangoFila$, "0:3", "", 0, False, False, False, False, 0, &HC0E0DD)
        End If
    Next ix&
    
    GRID.Rows = GRID.Rows + 1
    RangoFila$ = CStr(GRID.Rows - 1) & ":" & CStr(GRID.Rows - 1)
    Call DisenioDeCelda(GRID, RangoFila$, "0:3", "", 0, False, False, False, False, 0, &HC0E0DD)
    GRID.Rows = GRID.Rows + 1
    GRID.TextMatrix(GRID.Rows - 1, 0) = "Costo Total Fabricado"
    GRID.TextMatrix(GRID.Rows - 1, 3) = "$  " & TRIM$(FORMATNUM$(CostoTotalFabricado, "G14.2"))
        RangoFila$ = CStr(GRID.Rows - 1) & ":" & CStr(GRID.Rows - 1)
        Call DisenioDeCelda(GRID, RangoFila$, "0:3", "", 0, True, False, False, False, 0, &HE0E0E0)
    

    ' COSTO TOTAL POR UNIDAD / MODELO
    ' *******************************
    
    Dim TotalGastosPorUnidad As Double
    Dim CostoUnidad As Double
    
    GRID.Rows = GRID.Rows + 2
    If CantidadTotalFabricada > 0 Then
        TotalGastosPorUnidad = TotalDeGastos / (CantidadTotalFabricada + UnidadesConFuelle)
    End If
    
    For ix& = 1 To NUMAS
        If ArticulosFabricados(ix&) > 0.05 Then
            GRID.Rows = GRID.Rows + 1
            CostoUnidad = (TotalGastosPorUnidad + CostoArticulosFabricados(ix&))
            GRID.TextMatrix(GRID.Rows - 1, 0) = "Costo de un " & DESCRIT0$(CInt(ix&))
            GRID.TextMatrix(GRID.Rows - 1, 1) = "$  " & TRIM$(FORMATNUM$(CostoUnidad, "G14.2"))
            RangoFila$ = CStr(GRID.Rows - 1) & ":" & CStr(GRID.Rows - 1)
            Call DisenioDeCelda(GRID, RangoFila$, "0:1", "", 0, False, False, False, False, 0, &HC0E0DD)
        End If
    Next ix&
    
    UltimaFilaDelReporte = GRID.Rows - 1
    
99: Exit Sub

ErrMaterialFabricado:
    Call COMUNI("Error al Obtener Cantidad de Materiales Fabricados\ \" & Err.Description)
    On Error GoTo 0
    On Error Resume Next
    MaterialFabricado.Close
    Set MaterialFabricado = Nothing
    On Error GoTo 0
    GoTo CierreProlijo
    
ErrCargaDetalle:
    Call COMUNI("Error al Cargar Detalle de Gastos\ \" & Err.Description)
    On Error GoTo 0
    GoTo CierreProlijo
    
CierreProlijo:
    GRID.Rows = 0: GRID.Cols = 0
    GRID.Rows = 30: GRID.Cols = 4
    GRID.ColWidth(0) = GRID.ColWidth(0) + 5000: GRID.ColWidth(1) = GRID.ColWidth(1) + 500
    GRID.ColWidth(2) = GRID.ColWidth(1): GRID.ColWidth(3) = GRID.ColWidth(1)
    GoTo 99

End Sub




Private Sub ExportarExcel_Click()

    If GRID.Rows > 1 Then
    
        Dim NombreDelArchivo As String
        
        CommonDialog1.CancelError = True
        CommonDialog1.DialogTitle = "Exportación de Planilla de Costos"
        CommonDialog1.FILTER = "Hojas de Cálculo Excel(*.XLS)|*.XLS|"
        CommonDialog1.FilterIndex = 1
        On Error Resume Next
        CommonDialog1.ShowOpen
        On Error GoTo 0
        '
        Openforms = DoEvents
        NombreDelArchivo = UCase$(CommonDialog1.FileName)
        
        If TRIM$(NombreDelArchivo) <> "" Then
            If UCase(Mid$(NombreDelArchivo, Len(TRIM$(NombreDelArchivo)) - 3)) <> ".XLS" Then
                NombreDelArchivo = NombreDelArchivo & ".xls"
            End If
            Call GeneraExcelDeGrilla(NombreDelArchivo, GRID, "COSTO POR UNIDAD DE UNA CARROCERIA '" & TRIM$(Label2) & "'")
        End If

    End If
    
End Sub

Private Sub Form_Load()

    Call VERJOBID(OKEY%)
    
    If OKEY% < 1 Then Call FINAL("")
    '
    EPAC$ = TRIM$(UCase$(EMPREAC$))
    If EPAC$ <> "" Then
       Caption = Caption + "  -  " + EPAC$
    End If
    Call GRATITFOR(Caption)
    '
    Call ABRECONEXION
    '
    If Left$(APLINVO$, 8) = "SEGUIPRO" Then
       SEGUIPRO_TB.Show 1
       Call FINAL("")
    End If
    '
    GRID.Rows = 0: GRID.Cols = 0
    GRID.Rows = 1: GRID.Cols = 4
    GRID.ColWidth(0) = GRID.ColWidth(0) + 5000: GRID.ColWidth(1) = GRID.ColWidth(1) + 500
    GRID.ColWidth(2) = GRID.ColWidth(1): GRID.ColWidth(3) = GRID.ColWidth(1)
      
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call CIERRARCH("*.*")
    Call FINAL("")
End Sub

Private Sub Label2_Change()

    Call GeneraPlanillaCostosProduccion
    
End Sub


Function RangoFormatoExcel$(Fila1 As Long, Columna1 As Long, Optional Fila2 As Long, Optional Columna2 As Long)
    
    Dim LetraFila As String
    
    If Fila2 = 0 Then Fila2 = Fila1
    If Columna2 = 0 Then Columna2 = Columna1
    
    LetraFila = Chr$(65 + (Fila1 - 1))
    RangoFormatoExcel$ = LetraFila & CStr(Columna1)
    
    LetraFila = Chr$(65 + (Fila2 - 1))
    RangoFormatoExcel$ = RangoFormatoExcel$ & ":" & LetraFila & CStr(Columna2)

End Function

