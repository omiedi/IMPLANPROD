Attribute VB_Name = "COMMDEMO"
Public PAUSA As Integer
Public INVOCA$
Public YAEXTRA%
Public VERIPAN%
Public CAPAN%
Public CTX%
Public EXTRA%
Dim PANTA$(256)
'

Sub SETCOMMA(A$)
    '
    NX% = FreeFile
    Open "COMMDEMO" For Output As #NX%
    Print #NX%, A$
    Close #NX%
    '
End Sub
'

Function COMMDEM$()
    '
    On Error GoTo SEGUIR
    NX% = FreeFile
    Open "COMMDEMO" For Input As #NX%
    Line Input #NX%, A$
    Close #NX%
    COMMDEM$ = LTrim$(RTrim$(UCase$(A$)))
    '
    Open "COMMDEMO" For Output As #NX%
    Print #NX%, ""
    Close #NX%
    '
    Exit Function
    '
SEGUIR:
    COMMDEM$ = ""
    Exit Function
    '
End Function
'

Sub MOREINF()
    Call OPNODIS
End Sub
'
Sub OPNODIS()
    '
    MsgBox "No hay mayor información disponible" + Chr$(10) + Chr$(13) + "para  la presente  pantalla FLEXOFT."
    '
End Sub
'

Sub PRESDEMO(CTY%)
    '
10  If CTY% > CAPAN% Then
       WINDEMO1.Timer3.Enabled = False
       WINDEMO1.Command1.Visible = True
       WINDEMO1.Show
       Exit Sub
    End If
    '
    On Error GoTo ERRO
    Shell PANTA$(CTY%), 1
    On Error GoTo 0
    '
    Exit Sub
    '
ERRO:
    MsgBox "Imposible Presentar Pantalla " + PANTA$(CTY%)
    CTY% = CTY% + 1
    Resume 10
    '
End Sub
'

Sub EXTRACT()
    '
    Static YAEX%
    If YAEX% = 0 Then
        YAEX% = 1
        '
        Screen.MousePointer = 11
        Call SETCOMMA("****")
        '
        NX% = FreeFile
        Open "COMMDEMX" For Output As #NX%
        Print #NX%, "NEXT"
        Close #NX%
        '
        YAEY% = 0
        On Error GoTo 99
        Open "DEMOZIP.ZIP" For Input Shared As #NX%
        Close #NX%
        '
10      On Error GoTo 0
        Open "FUNZIP.BAT" For Output As #NX%

        If YAEY% = 0 Then
            Print #NX%, "PKUNZIP -O DEMOZIP"
            Print #NX%, "DEL DEMOZIP.ZIP"
        End If
        Print #NX%, "COPY COMMDEMX COMMDEMO"
        Close #NX%
        '
        Shell "FUNZIP.BAT"
        YAEXTRA% = 1
        Exit Sub
        '
99      YAEY% = 1
        Resume 10
        '
    End If
    '
RECUERR:
    Resume Next
    Exit Sub
    '
End Sub
'

Sub SECUDEMO(SECU$)
    '
    CTX% = 0
    CAPAN% = 0
    WINDEMO1.LISECU.Clear
    '
    If SECU$ = "INTRODU" Then
        '
        ADDLIS "DEMO001"   ' Mensaje de Bienvenida
        ADDLIS "DEMO002"   ' Licencia de distribucion
        ADDLIS "PRES003"   ' Presentacion de escritorio
        '
    ElseIf SECU$ = "ARCHIMA" Then
        '
        ADDLIS "AMAS001"   ' Escritorio Actualizacion Archivos Maestros
        ADDLIS "AMAS002"   ' Escritorio Maestro de Items de Stock
        ADDLIS "DEMO011"   ' Formulario Modificacion Item de Stock
        ADDLIS "AMAS003"   ' Modificacion Item con Lista de Validacion
        ADDLIS "DEMO010"   ' Busqueda selectiva por Descripcion
        ADDLIS "DEMO012"   ' Formulario Sustitucion de Codigo
        ADDLIS "DEMO013"   ' Formulario Estructura de Producto
        ADDLIS "DEMO014"   ' Formulario Estructura Subconjunto
        ADDLIS "AMAS004"   ' Escritorio Hojas de Ruta y Operaciones
        ADDLIS "AMAS005"   ' Secuencia de Operaciones por Articulo
        ADDLIS "AMAS006"   ' Detalle de Tiempos por Operacion
        ADDLIS "AMAS007"   ' Ayuda FLEXOFT en Linea
        ADDLIS "AMAS008"   ' Maquinas y Equipos de Produccion
        ADDLIS "DEMO006"   ' Escritorio Actualizacion de Clientes
        ADDLIS "DEMO007"   ' Formulario Modificacion de Clientes
        ADDLIS "DEMO008"   ' Formulario Tablas de Validacion Clientes
        ADDLIS "DEMO027"   ' Pantalla Final de Demostracion
        '
    ElseIf SECU$ = "STOCK" Then
        '
        ADDLIS "AMAS002"   ' Escritorio Maestro de Items de Stock
        ADDLIS "DEMO011"   ' Formulario Modificacion Item de Stock
        ADDLIS "AMAS003"   ' Modificacion Item con Lista de Validacion
        ADDLIS "DEMO010"   ' Busqueda selectiva por Descripcion
        ADDLIS "DEMO012"   ' Formulario Sustitucion de Codigo
        ADDLIS "STOK001"   ' Escritorio Movimientos de Stock
        ADDLIS "DEMO016"   ' Formulario Consulta de Existencias
        ADDLIS "DEMO017"   ' Ficha Kardex por Pieza
        ADDLIS "STOK002"   ' Generacion de Inventarios
        ADDLIS "STOK003"   ' Consulta de Inventarios por pantalla
        ADDLIS "DEMO019"   ' Menu de control de Sub-Minimos
        ADDLIS "STOK004"   ' Lista de Items Sub-Minimo
        
        ADDLIS "DEMO027"   ' Pantalla Final de Demostracion
        '
    Else
        WINDEMO1.Show
    End If
    '
    CAPAN% = WINDEMO1.LISECU.ListCount
    If CAPAN% > 0 Then
        For II% = 1 To CAPAN%
            WINDEMO1.LISECU.ListIndex = II% - 1
            PANTA$(II%) = WINDEMO1.LISECU.Text
        Next II%
        '
        If EXTRA% <> 1 Then
           Call SETCOMMA("NEXT")
           WINDEMO1.Timer3.Enabled = True
        End If
        '
    End If
    '
End Sub
'

Sub ADDLIS(TXX$)
    '
    WINDEMO1.LISECU.AddItem TXX$
    '
End Sub
