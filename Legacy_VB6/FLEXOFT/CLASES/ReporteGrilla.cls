VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ReporteGrilla"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public TITULO As String
Public EncabezadoDeColumnas As String
Public ConsultaSql As String
Public Grilla As MSFlexGrid
Public MENU1 As VB.menu
Public MENU2 As VB.menu
Public MenuEdicion As Boolean
Public FiltrosHabilitados As Boolean
Public CodigoReporte As String

Private Sub Class_Initialize()

    Set MENU1 = GRILLAGRAL.mnuMenu1
    Set MENU2 = GRILLAGRAL.mnuMenu2
    MenuEdicion = False
    FiltrosHabilitados = True
    
End Sub

Private Sub Crear()
    
    If ConsultaSql$ <> "" Then
        Dim Rst As ADODB.Recordset
        Set Rst = FLEXCONN.Execute(ConsultaSql)
        Call Carga_MSF_Recordset(Rst, EncabezadoDeColumnas, GRILLAGRAL.GRID, 0)
        On Error Resume Next
        Rst.Close
        Set Rst = Nothing
        On Error GoTo 0
    Else
        If Grilla.Rows > 0 Then
            For fi = 1 To Grilla.Rows - 1
                If Mid$(TRIM$(Grilla.TextMatrix(fi, 0)), 1, 1) = "*" Then
                    GRILLAGRAL.GRID.Rows = GRILLAGRAL.GRID.Rows + 1
                    ' PASO LOS DATOS DE LA FILA, COLUMNA POR COLUMNA
                    For co = 1 To Grilla.Cols - 1
                        GRILLAGRAL.GRID.TextMatrix(GRILLAGRAL.GRID.Rows - 1, co) = Grilla.TextMatrix(fi, co)
                    Next co
                    
                    ' AJUSTA ALTURA DE LA FILA QUE SE SETEA EN LA COLUMNA 0 JUNTO CON LA MARCA *-n
                    n = XVALO(Mid$(TRIM$(Grilla.TextMatrix(fi, 0)), 3))
                    If n > 0 Then
                        GRILLAGRAL.GRID.RowHeight(GRILLAGRAL.GRID.Rows - 1) = GRILLAGRAL.GRID.RowHeight(GRILLAGRAL.GRID.Rows - 1) * n
                    End If
                End If
            Next fi
        End If
    End If
    
    GRILLAGRAL.GRID.WordWrap = True
    GRILLAGRAL.MSF_2.WordWrap = True
    
End Sub

Public Sub Mostrar(Optional CargaExterna As Boolean = False)

    TituloLista$ = TITULO
    Campos$ = EncabezadoDeColumnas
    Call CARGA_ENCABEZADO(GRILLAGRAL.GRID, Campos$, GRILLAGRAL, 1)
    Call CARGA_ENCABEZADO(GRILLAGRAL.MSF_2, Campos$, GRILLAGRAL, 1)
    GRILLAGRAL.Caption = TituloLista$

    If CargaExterna = False Then Call Crear
    
    GRILLAGRAL.MENU1 = False
    If MENU1.Visible Or MENU2.Visible Then
        GRILLAGRAL.MENU1 = True
        GRILLAGRAL.mnuMenu1.Visible = True
        GRILLAGRAL.mnuMenu1.Caption = MENU1.Caption
    End If
    
    GRILLAGRAL.MENU2 = False
    If MenuEdicion = True Then
        GRILLAGRAL.MnuEdición = True
        GRILLAGRAL.mnuMenu2.Caption = MENU2.Caption
    End If
    
    GRILLAGRAL.Frame1.Visible = True
    If FiltrosHabilitados = False Then GRILLAGRAL.Frame1.Visible = False
    GRILLAGRAL.SinEfectoClickEnGrilla% = 1
    
    GRILLAGRAL.Show 1
    
    If GRILLAGRAL.MENU1 > 0 Then
        MenuElegido$ = GRILLAGRAL.mnuMenu1.Caption
    ElseIf GRILLAGRAL.MENU2 > 0 Then
        MenuElegido$ = GRILLAGRAL.mnuMenu2.Caption
    End If
        
    Select Case MenuElegido$
        Case "Imprimir":
                Dim MSF As MSFlexGrid

                Set MSF = GRILLAGRAL.GRID
                If GRILLAGRAL.MSF_2.Rows > 1 Then Set MSF = GRILLAGRAL.MSF_2
                Call CARGA_TABLA_IMPRE(MSF, EncabezadoDeColumnas, CodigoReporte, GRILLAGRAL.Caption, "TABLA_IMPRE", RET_NAMECONS$, 1)
                Call FINAL("?" & RET_NAMECONS$)
                GRILLAGRAL.MENU1 = 0
        
        Case "Exportar"
    End Select
        
End Sub


