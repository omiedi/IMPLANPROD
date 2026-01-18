Attribute VB_Name = "FUNCIONESMSF"
'Public canticolumnas%
'Public sCol As Integer, Orden As SortSettings

'declaracion para llevar al final el flexgrid
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" _
    (ByVal hwnd As Long, ByVal wMsg As Long, _
    ByVal wParam As Long, lParam As Any) As Long

' declaraciones del api para mover con la ruedita del mouse el msflexgrid
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" ( _
    ByVal hwnd As Long, _
    ByVal nIndex As Long, _
    ByVal dwNewLong As Long) As Long

Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" ( _
    ByVal lpPrevWndFunc As Long, _
    ByVal hwnd As Long, _
    ByVal Msg As Long, _
    ByVal wParam As Long, _
    ByVal lParam As Long) As Long


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

' Constantes
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Private Const GWL_WNDPROC = (-4)
Private Const WM_MOUSEWHEEL = &H20A
Private Const WM_VSCROLL As Integer = &H115


Dim PrevProc As Long

Sub AGREGAR_NombreFilas(MSF As msFlexGrid, NOMBREFILAS$, aPartiFila%, COLUMNA%)

       Dim strArray() As String
       EnFila% = aPartiFila%
       strArray = Split(NOMBREFILAS$, ";")
       For i% = 0 To UBound(strArray)
           LEIDO$ = strArray(i%)
           If EnFila% >= MSF.Rows Then MSF.Rows = EnFila% + 1
           MSF.TextMatrix(EnFila%, COLUMNA%) = LEIDO$
           If LEIDO$ = "DETALLE DE INGRESOS" Then
              MSF.TextMatrix(EnFila%, MSF.Cols - 1) = 0
              EnFila% = EnFila% + 1
              If EnFila% >= MSF.Rows Then MSF.Rows = EnFila% + 1
              MSF.TextMatrix(EnFila%, MSF.Cols - 1) = 1
           ElseIf LEIDO$ = "TOTAL INGRESOS" Then
              MSF.TextMatrix(EnFila%, MSF.Cols - 1) = 991
              EnFila% = EnFila% + 1
              If EnFila% >= MSF.Rows Then MSF.Rows = EnFila% + 1
              MSF.TextMatrix(EnFila%, MSF.Cols - 1) = 990
              EnFila% = EnFila% + 1
              If EnFila% >= MSF.Rows Then MSF.Rows = EnFila% + 1
              EnFila% = EnFila% + 1
              If EnFila% >= MSF.Rows Then MSF.Rows = EnFila% + 1
              MSF.TextMatrix(EnFila%, MSF.Cols - 1) = 992
           ElseIf LEIDO$ = "DETALLE DE EGRESOS" Then
              MSF.TextMatrix(EnFila%, MSF.Cols - 1) = 998
              EnFila% = EnFila% + 1
              If EnFila% >= MSF.Rows Then MSF.Rows = EnFila% + 1
              MSF.TextMatrix(EnFila%, MSF.Cols - 1) = 999
           ElseIf LEIDO$ = "TOTAL EGRESOS" Then
              MSF.TextMatrix(EnFila%, MSF.Cols - 1) = 3001
              EnFila% = EnFila% + 1
              If EnFila% >= MSF.Rows Then MSF.Rows = EnFila% + 1

              MSF.TextMatrix(EnFila%, MSF.Cols - 1) = 3000
              EnFila% = EnFila% + 1
              If EnFila% >= MSF.Rows Then MSF.Rows = EnFila% + 1
              EnFila% = EnFila% + 1
              If EnFila% >= MSF.Rows Then MSF.Rows = EnFila% + 1
              MSF.TextMatrix(EnFila%, MSF.Cols - 1) = 3002
              
           End If
           EnFila% = EnFila% + 1
           
       Next i%
       '
       NUCOLU% = 1: VENTABNEW.TEPRUEBA = String$(24, "A"): MSF.ColWidth(NUCOLU% - 1) = VENTABNEW.TEPRUEBA.Width: MSF.ColAlignment(NUCOLU% - 1) = 1: MSF.TextMatrix(0, NUCOLU% - 1) = ""
       '
       'RECORRO LA GRILLA Y ELIMINO LOS ITEMS QUE NO TENGAN INFORMACION PARA QUE NO ARRUINE EL ORDENAMIENTO
       R& = 0
       While R& < MSF.Rows - 1
          R& = R& + 1
          If TRIM$(MSF.TextMatrix(R&, MSF.Cols - 1)) = "" Then 'SI NO TIENE DATO DE ORDENAMIENTO ES POR QUE NO TIENE DATOS EN FILA
             Call cmdEliminaItem(MSF, R&)
             R& = R& - 1
          End If
       Wend
       'FIN 'RECORRO LA GRILLA Y ELIMINO LOS ITEMS QUE NO TENGAN INFORMACION

End Sub

Function CALCULARFORMULAS(TEXTOFINAL$)
      TEXTOFINAL2$ = TEXTOFINAL$
      POS1% = InStr(TEXTOFINAL2$, ")")
      If POS1% > 0 Then
            '(10+(20*30)+40+50)
            For L& = 1 To Len(TEXTOFINAL2$)
              P1% = InStr(TEXTOFINAL2$, "(")
              If P1% > 0 Then
                 TEXTOFINAL2$ = Mid$(TEXTOFINAL2$, P1% + 1)
              Else
                 On Error Resume Next
                 P2% = InStr(TEXTOFINAL2$, ")")
                 TEXTOFINAL2$ = Mid$(TEXTOFINAL2$, 1, P2% - 1)
                 If Err > 0 Then Call COMUNI("Error de Sintaxis en Celda " & UBICAX$): Exit Function
                 On Error GoTo 0

                 BLOQUE_A_CALCULAR$ = TEXTOFINAL2$
                 BLOQUE_A_CALCULAR$ = CALCULABLOQUE1(BLOQUE_A_CALCULAR$)
                 CALCULARFORMULAS = XVALO(CALCULABLOQUE1(BLOQUE_A_CALCULAR$))
                 TEXTOFINAL$ = REPLACAR$(TEXTOFINAL$, "(" & TEXTOFINAL2$ & ")", TRIM$(Str$(CALCULARFORMULAS)))
                 TEXTOFINAL2$ = TEXTOFINAL$
                 If InStr(TEXTOFINAL2$, "(") < 1 Then
                   'SI YA NO POSEE MAS PARENTESIS PUEDE SER QUE QUEDE UNA OPERACION POR ESO LO PASO DEVUELTA AL VALORCELDA
                   BLOQUE_A_CALCULAR$ = TEXTOFINAL2$
                   CALCULARFORMULAS = XVALO(CALCULABLOQUE1(BLOQUE_A_CALCULAR$))
                   Exit For
                 End If
              End If
            Next L&
            
      Else
        BLOQUE_A_CALCULAR$ = TEXTOFINAL2$
        CALCULARFORMULAS = XVALO(CALCULABLOQUE1(BLOQUE_A_CALCULAR$))
      End If
End Function

Sub ColorearMismoColorQuitarAzterizco(MSF As msFlexGrid, COLOR As ColorConstants)
    For i& = 1 To MSF.Rows - 1
       For J% = 0 To MSF.Cols - 1
         If J% = 0 Then
            MSF.TextMatrix(i&, J%) = ""
         Else
            
            MSF.row = i&
            MSF.col = J%
          
            MSF.CellBackColor = COLOR
         End If
       Next J%
    Next i&
End Sub

Public Sub HookForm(obj As Object)
    PrevProc = SetWindowLong(obj.hwnd, GWL_WNDPROC, AddressOf WindowProc)
End Sub

Sub LeoYescriboExcel(RUTA_ARCH$, Columnas_a_Leer$)

    '****** LEO ESCRIBO EN EXCEL ******
    
    'este proceso se realizo para meditea ellos enviaron los excel con las notas de credito
    'que estaban aplicadas y dejando tantas filas como huecos de las nc que no tenian informacion de la aplicacion
    'la idea es leer el excel y donde encuentra una celda vacia en la columna donde esta la nota de credito
    'entonces le suma 1 a la ultima encontrada y trae la info de la tabla cajabanc y la imprime en el mismo excel
    ', asi hasta que encuentre otra nota de credito existente en el excel, y repite este ciclo
    
    NUMEROHOJA% = XVALO(InputBox$("NUMERO DE HOJA", "Ingrese el N° de Hoja a Leer"))
    CANREGISTR& = XVALO(InputBox$("CANTIDAD DE REGISTROS A LEER ", "Ingrese HASTA QUE FILA SE VAN A LEER LOS DATOS"))
    
    If CANREGISTR& < 1 Or NUMEROHOJA% < 1 Then
       Call COMUNI("Faltan Ingresar Datos de Hoja o Cantidad de Registros")
       Screen.MousePointer = 1
       Exit Sub
    End If
    
    Dim ObjetoExcel As Excel.Application
    Dim Libro As Excel.Workbook
    Dim Hoja As Excel.Worksheet
    
    Set ObjetoExcel = CreateObject("Excel.Application")
    Set Libro = ObjetoExcel.Workbooks.Open(RUTA_ARCH$)

    With ObjetoExcel
      While REG& < CANREGISTR&
         REG& = REG& + 1
        ' ARCHIVO EXCEL / ABRE EL LIBRO
         Col_a_Leer% = XVALO(Columnas_a_Leer$)
         DOCUMENTO$ = TRIM$(Libro.Worksheets(NUMEROHOJA%).Cells(REG&, Col_a_Leer%))
         If DOCUMENTO$ = "" And DOCUMENTO_ANT$ <> "" Then
            NUEVODOCUMENTO$ = Left$(DOCUMENTO_ANT$, 8) & FORMATOCON0(ULTIMODOCULEIDO& + 1, 8) & Right(DOCUMENTO_ANT$, 2)
            CONDI$ = "CODICOM_LAR = '" & NUEVODOCUMENTO$ & "' AND NUCLIEN > 0 AND TIPOMOVIM = 'FVEN' AND OPCIMOVIM ='NC'"
            SQLX$ = "SELECT FECHEMISI,FECHCONTA,CODICOM_LAR,IHABECOMP,CUIT_CLI,NUCLIEN FROM CAJABANC WITH(NOLOCK) WHERE " & CONDI$
            Set rs = READSET(SQLX$)
            FECHAEMI = SAFETEXT$(rs!FECHEMISI)
            FECONTA = SAFETEXT$(rs!fechconta)
            ImpoNeto = XVALO(rs!IHABECOMP)
            NC = XVALO(rs!NUCLIEN)
            CuitClien$ = SAFETEXT$(rs!Cuit_Cli)
            rs.Close: Set rs = Nothing
            
            If NC > 0 Then
                Libro.Worksheets(NUMEROHOJA%).Cells(REG&, 1) = FECHAEMI
                Libro.Worksheets(NUMEROHOJA%).Cells(REG&, 2) = FECONTA
                Libro.Worksheets(NUMEROHOJA%).Cells(REG&, 3) = NUEVODOCUMENTO$
                Libro.Worksheets(NUMEROHOJA%).Cells(REG&, 4) = Format(ImpoNeto, "##,####.##")
                
                Libro.Worksheets(NUMEROHOJA%).Cells(REG&, 9) = NC
                Libro.Worksheets(NUMEROHOJA%).Cells(REG&, 10) = rasocli$(NC)
                Libro.Worksheets(NUMEROHOJA%).Cells(REG&, 11) = CuitClien$
            End If
            DOCUMENTO$ = NUEVODOCUMENTO$

         End If
         
         DOCUMENTO_ANT$ = DOCUMENTO$
         ULTIMODOCULEIDO& = XVALO(Mid$(DOCUMENTO_ANT$, 9, 8))
         
         
         'HOJA.PrintOut Copies:=1, Collate:=True
         ' HOJA.PrintOut Copies:=1, Collate:=True, ActivePrinter:=IMPRESORA$
         'HOJA.PrintOut Copies:=1, Collate:=True 'asi sale por la impresora predeterminada
      Wend
    End With
    
    ObjetoExcel.ActiveWorkbook.Save
    'Set HOJA = Libro.Sheets(NUMEROHOJA%)
         
         
    'HOJA.PrintOut Copies:=1, Collate:=True
    ' HOJA.PrintOut Copies:=1, Collate:=True, ActivePrinter:=IMPRESORA$
    'HOJA.PrintOut Copies:=1, Collate:=True 'asi sale por la impresora predeterminada
    'CIERRA HOJA, LIBRO Y EXCEL
99  On Error Resume Next
    Set Hoja = Nothing
    Libro.Close: Set Libro = Nothing
    ObjetoExcel.Close: Set ObjetoExcel = Nothing
    On Error GoTo 0
    
    Screen.MousePointer = 1
    Call COMUNI("Proceso Terminado Para Hoja N°: " & NUMEROHOJA% & " Con: " & CANREGISTR& & " Leidos")
End Sub

Function NRODECOLUMNA_SEGUN_TEXTO%(MSF As msFlexGrid, TEXTO$, Optional CantidadCaracteresaComparar%)
   'BUSCA EL NRO DE COLUMNA SEGUN EL TEXTO PARAMETRIZADO SOLO BUSCA EN EL ENCABEZADO USADO PARA LA CARGA DE FECHAS
   TEXTO$ = (TRIM$(UCase$(TEXTO$)))
'   If CantidadCaracteresaComparar% > 0 Then
'      CAN% = CantidadCaracteresaComparar%
'      TEXTO$ = Mid$(TEXTO$, 1, CAN%)
'   End If
   
   For i% = 1 To MSF.Cols - 1
      TXLEIDO$ = UCase$(TRIM$(MSF.TextMatrix(0, i%)))
      If CantidadCaracteresaComparar% > 0 Then
        CAN% = CantidadCaracteresaComparar%
        TXLEIDO$ = Mid$(TXLEIDO$, 1, CAN%)
      End If

10    If TXLEIDO$ = TEXTO$ Then
         NRODECOLUMNA_SEGUN_TEXTO% = i%
         Exit For
      End If
      If Left$(TXLEIDO$, 3) = "SET" Then TXLEIDO$ = REPLACAR(TXLEIDO$, "T", "P"): GoTo 10
      
   Next i%
End Function

Public Sub UnHookForm(obj As Object)
    SetWindowLong obj.hwnd, GWL_WNDPROC, PrevProc
End Sub

' Procedimiento qie intercepta los mensajes de windows, en este caso para _
  interceptar el uso del Scroll del mouse
  '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function WindowProc(ByVal hwnd As Long, _
                           ByVal uMsg As Long, _
                           ByVal wParam As Long, _
                           ByVal lParam As Long) As Long
                           
'    NOTA: SI EL FORMULARIO DONDE ESTA EL MSF QUE LLAMA A ESTA FUNCION SE CIERRA CON UN HIDE CUANDO SE VUELVE A ABRIR DA ERROR
'    SOLUCION VALIDAR SI EL MISMO ESTA EN MEMORIA ANTES DE ASIGNAR LA FUNCION, SI NO ESTA EN MEMORIA APLICARLA CASO CONTRARIO NO
'    EJEMPLO CON UN FRMZELECHO QUE DABA ERROR (ESTE CODIGO DEBE APLICARSE AL PRINCIPIO DE LA RUTINA ANTES DE NOMBRAR AL FRMZELECHO
'    POR EJEMPLO PARA APLICARLE ALGUNA PROPIEDAD O CARGAR LA GRILLA, POR QUE SI ES ASI YA LO PONE EN MEMORIA Y NO VA A FUNCIONAR.
'    If IsFormLoaded(FRMZELECHO) Then
'        ESTAENMEMORIA = 1
'    Else
'        Call HookForm(FRMZELECHO.msf2) 'ESTA ES LA LLAMDA PARA QUE EL MSF SE MUEVA CON LA RUEDITA DEL MOUSSE.
'        Call HookForm(FRMZELECHO.MSF1) '
'    End If
                         

    WindowProc = CallWindowProc(PrevProc, hwnd, uMsg, wParam, lParam)
    
    If uMsg = WM_MOUSEWHEEL Then
        If wParam < 0 Then
            ' envia mediante SendMessage el comando para mover el Scroll hacia abajo
            SendMessage hwnd, WM_VSCROLL, ByVal 1, ByVal 0
        Else
            ' Mueve el scroll hacia arriba
            SendMessage hwnd, WM_VSCROLL, ByVal 0, ByVal 0
        End If
    End If
End Function

    
'Sub CARGA_GRILLA_SHOWREP(Form1 As Form, MSF As MSHFlexGrid, Campos$, VALORES$, INICIALIZO%, Optional Negrita%)
'   If INICIALIZO% < 1 Then
'      Call ENCABEZADO_GENREP(MSF, Campos$, Form1)
'      INICIALIZO% = 1
'   End If
'
'   If TRIM$(VALORES$) = "" Then Exit Sub
'   MSF.Rows = MSF.Rows + 1
'   Call TEXTOLOTES$(VALORES$, "|")
'   For j& = 1 To UBound(CADENATEX$)
'      MSF.TextMatrix(MSF.Rows - 1, j& - 1) = CADENATEX$(j&)
'      MSF.Row = MSF.Rows - 1
'      MSF.COL = j& - 1
'      If Negrita% > 0 Then MSF.CellFontBold = True
'   Next j&
'End Sub
Sub CARGA_GRILLA_SHOWREP2(Form1 As Form, MSF As MSHFlexGrid, Campos$, VALORES$, INICIALIZO%, Optional Negrita%)
   If INICIALIZO% < 1 Then
      Call ENCABEZADO_GENREP2(MSF, Campos$, Form1)
      INICIALIZO% = 1
   End If

   If TRIM$(VALORES$) = "" Then Exit Sub
   MSF.Rows = MSF.Rows + 1
   Call TEXTOLOTES$(VALORES$, "|")
   For J& = 1 To UBound(CADENATEX$)
      MSF.TextMatrix(MSF.Rows - 1, J& - 1) = CADENATEX$(J&)
      MSF.row = MSF.Rows - 1
      MSF.col = J& - 1
      If Negrita% > 0 Then MSF.CellFontBold = True
   Next J&
End Sub

Sub CARGA_GRILLA_GENERICO(Form1 As Form, MSF As msFlexGrid, Campos$, VALORES$, INICIALIZO%)
   On Error GoTo ERROR_TRAZAR
   If INICIALIZO% < 1 Then
      Call CARGA_ENCABEZADO(MSF, Campos$, Form1)
      INICIALIZO% = 1
   End If
   If TRIM$(VALORES$) = "" Then Exit Sub
   MSF.Rows = MSF.Rows + 1
   Call TEXTOLOTES$(VALORES$, "|")
   For J& = 1 To UBound(CADENATEX$)
      MENSAX$ = "REG: " & MSF.Rows & " COL: " & J&
      MSF.TextMatrix(MSF.Rows - 1, J&) = CADENATEX$(J&)
   Next J&
Exit Sub
ERROR_TRAZAR:
   MsgBox MENSAX$ + " ERROR DENTRO DE CARGA_GRILLA_GENERICO  " & " " & Err.Description
   Call FINAL("")
End Sub

Sub EstiloCeldaSegunGrilla(Grilla As msFlexGrid, Hoja As Excel.Worksheet, FilaGrid As Long, ColumnaGrid As Long, FILAEXCEL As Long, ColumnaExcel As Long)

    Grilla.row = FilaGrid: Grilla.col = ColumnaGrid
    FuenteNombre = Grilla.CellFontName
    Tamanio = Grilla.CellFontSize
    Negrita = Grilla.CellFontBold
    Italica = Grilla.CellFontItalic
    Subrayado = Grilla.CellFontUnderline
    Alinear = Grilla.ColAlignment(ColumnaGrid)
    ColorFuente = Grilla.CellForeColor
    ColorFondo = Grilla.CellBackColor
    '
    With Hoja.Cells(FILAEXCEL, ColumnaExcel)
        .Font.Name = FuenteNombre
        .Font.Size = Tamanio
        .Font.Bold = Negrita
        .Font.Underline = Subrayado
        .Font.Italic = Italica
            Call LongToRGB(ColorFuente, rd%, gr%, BL%)
        .Font.COLOR = RGB(rd%, gr%, BL%)
            rd% = 255: gr% = 255: BL% = 255
            If ColorFondo <> 0 Then Call LongToRGB(ColorFondo, rd%, gr%, BL%)
        If Not (rd% = 255 And gr% = 255 And BL% = 255) Then
            .Interior.COLOR = RGB(rd%, gr%, BL%)
        End If
    End With
    
End Sub

'Sub ENCABEZADO_GENREP(MSF1 As MSHFlexGrid, LISTADECAMPOS$, Form1 As Form)
'    '
'
'    Dim NOCAMPO$(), FORMX$()
'    ReDim NOCAMPO$(1), FORMX$(1)
'    '
'    ANCHO_TOTAL = 0
'    On Error Resume Next
'    If Form1.Label8.Count > 1 Then
'      Erase Form1.Label8
'    End If
'    NUPIC% = 0
'    MODOSE$ = UCase$(SAFETEXT$(MODOSEL$))
'    LSTX$ = TRIM$(LISTADECAMPOS$)
'    CANCA% = 0
'    ATOTA% = 1
'    While LSTX$ <> ""
'       PPXX% = InStr(LSTX$, ";")
'       PPXX2% = InStr(LSTX$, ",")
'       If PPXX2% > 0 And PPXX2% < PPXX% Then PPXX% = PPXX2%
'       If PPXX% < 1 Then PPXX% = 1 + Len(LSTX$)
'       '
'       CANCA% = CANCA% + 1
'       If CANCA% >= UBound(NOCAMPO$) Then
'          ReDim Preserve NOCAMPO$(CANCA%), FORMX$(CANCA%)
'       End If
'       NOCX$ = Left$(LSTX$, PPXX% - 1)
'       LSTX$ = Mid$(LSTX$, PPXX% + 1)
'       '
'       PPXX% = InStr(NOCX$, "/")
'       If PPXX% < 1 Then PPXX% = 1 + Len(NOCX$)
'       NOCAMPO$(CANCA%) = Left$(NOCX$, PPXX% - 1)
'       FORMX$(CANCA%) = Mid$(NOCX$, PPXX% + 1)
'       '
'
'       If FORMX$(CANCA%) = "" Then FORMX$(CANCA%) = "A16"
'       '
'       ATOTA% = ATOTA% + XVALO(Mid$(FORMX$(CANCA%), 2)) + 2
'       '
'
'       If CANCA% = 1 Then ACOL1% = XVALO(Mid$(FORMX$(CANCA%), 2))
'       '
'    Wend
'
'    MSF1.Cols = CANCA%
'    MSF1.Rows = 1
'    For i% = 1 To CANCA%
'       'PRIMER COLUMNA ASIGNA POR DEFECTO ANCHO = 2
'       If i% = 0 Then
'         If UCase$(Form1.Name) = "SHOWREP07" Then GoTo 80
'
'         If ANCHO_COLUMNA_0% > 0 Then
'           ANCHO% = ANCHO_COLUMNA_0% 'SI  SE PASA EL ANCHO DE LA COLUMNA 0 COMO PARAMETRO TOMA ESTE
'         Else
'           ANCHO% = 3  'EN CASO CONTRARIO ASIGNA VALOR POR DEFECTO
'         End If
'         NUCOLU% = i%:   Form1.Teprueba = String$(ANCHO%, "A"): MSF1.ColWidth(NUCOLU%) = Form1.Teprueba.Width: MSF1.ColAlignment(NUCOLU%) = ALINEA%: MSF1.TextMatrix(0, NUCOLU%) = NOCAMPO$(i%)
'         GoTo 50
'       End If
'
'30     LETRX$ = "A"
'       ANCHO% = XVALO(Mid$(FORMX$(i%), 2))
'       If XVALO(Mid$(FORMX$(i%), 2)) > 0 Then
'          ANCHO% = ANCHO% + 2
'       ElseIf LETRX$ = "Z" Then
'
'       Else
'          ANCHO% = 0 'COLUMNAS OCULTAS PASADAS CON F0.0
'       End If
'       LETR$ = UCase$(Left$(FORMX$(i%), 1))
'       ALINEA% = 1
'
'       If LETR$ = "I" Or LETR$ = "G" Or LETR$ = "F" Then
'          ALINEA% = 6
'       ElseIf LETR$ = "Z" Then
'          ALINEA% = 3
'       End If
'       '
'       If ANCHO% > 0 Then
'          NUCOLU% = i% - 1: Form1.Teprueba = String$(ANCHO%, "A"): MSF1.ColWidth(NUCOLU%) = Form1.Teprueba.Width: MSF1.ColAlignment(NUCOLU%) = ALINEA%: MSF1.TextMatrix(0, NUCOLU%) = NOCAMPO$(i%)
'       Else ' PARA QUE NO SE VEAN LAS COLUMNAS OCULTAS
'          NUCOLU% = i% - 1: Form1.Teprueba = String$(ANCHO%, "A"): MSF1.ColWidth(NUCOLU%) = 0: MSF1.ColAlignment(NUCOLU%) = ALINEA%: MSF1.TextMatrix(0, NUCOLU%) = NOCAMPO$(i%)
'       End If
'       FORCAMPO$(NUCOLU% + 1) = LETR$
'       FORCOLUM$(NUCOLU% + 1) = FORMX$(NUCOLU% + 1)
'       MSF1.COL = NUCOLU% - 1
'       MSF1.CellFontBold = True
'
'       ANCHO_TOTAL = ANCHO_TOTAL + Form1.Teprueba.Width
'
'       'A PARTIR DE AQUI CODIGO PARA CARGAR ENCABEZADO DE GENREP07
'
'50     If UCase$(Form1.Name) = "SHOWREP07" Then
'          Form1.Label5.FontName = Form1.GRID.CellFontName
'          Form1.Label5.FontSize = Form1.GRID.CellFontSize
'          Form1.Label5.FontBold = Form1.GRID.CellFontBold
'          ' Label5.Caption = TEXTBIN$("1.234567", FORMAECO$(I&)) + "*"
'          Form1.Label5.Caption = String$(ANCHO%, "A")    ' PROVISORIAMENTE A VER SI FUNCIONA
'          Form1.GRID.COL = i%
'          Form1.GRID.ColWidth(i% - 1) = SHOWREP07.Label5.Width
'25        If i% >= Form1.Label8.Count Then
'             NUPIC% = Form1.Label8.Count
'             Load Form1.Label8(NUPIC%)
'             GoTo 25
'          End If
'          Form1.Label8(i%) = Str$(Len(Form1.Label5) - 1)
'          If i% = 0 Then Form1.Label8(i%) = 0
'       End If
'80  Next i%
'    '
'    If NUPIC% = i% - 1 Then
'      If UCase$(Form1.Name) = "SHOWREP07" Then Load Form1.Label8(NUPIC% + 1): Load Form1.Label8(NUPIC% + 2)
'    End If
'
'End Sub
Sub ENCABEZADO_GENREP2(MSF1 As MSHFlexGrid, LISTADECAMPOS$, Form1 As Form)
    '

    Dim NOCAMPO$(), FORMX$()
    ReDim NOCAMPO$(1), FORMX$(1)
    '
    ANCHO_TOTAL = 0
    On Error Resume Next
    If Form1.Label8.Count > 1 Then
      Erase Form1.Label8
    End If
    NUPIC% = 0
    MODOSE$ = UCase$(SAFETEXT$(MODOSEL$))
    LSTX$ = TRIM$(LISTADECAMPOS$)
    CANCA% = 0
    ATOTA% = 1
    While LSTX$ <> ""
       PPXX% = InStr(LSTX$, ";")
       PPXX2% = InStr(LSTX$, ",")
       If PPXX2% > 0 And PPXX2% < PPXX% Then PPXX% = PPXX2%
       If PPXX% < 1 Then PPXX% = 1 + Len(LSTX$)
       '
       CANCA% = CANCA% + 1
       If CANCA% >= UBound(NOCAMPO$) Then
          ReDim Preserve NOCAMPO$(CANCA%), FORMX$(CANCA%)
       End If
       NOCX$ = Left$(LSTX$, PPXX% - 1)
       LSTX$ = Mid$(LSTX$, PPXX% + 1)
       '
       PPXX% = InStr(NOCX$, "/")
       If PPXX% < 1 Then PPXX% = 1 + Len(NOCX$)
       NOCAMPO$(CANCA%) = Left$(NOCX$, PPXX% - 1)
       FORMX$(CANCA%) = Mid$(NOCX$, PPXX% + 1)
       '

       If FORMX$(CANCA%) = "" Then FORMX$(CANCA%) = "A16"
       '
       ATOTA% = ATOTA% + XVALO(Mid$(FORMX$(CANCA%), 2)) + 2
       '

       If CANCA% = 1 Then ACOL1% = XVALO(Mid$(FORMX$(CANCA%), 2))
       '
    Wend

    MSF1.Cols = CANCA%
    MSF1.Rows = 1
    For i% = 1 To CANCA%
       'PRIMER COLUMNA ASIGNA POR DEFECTO ANCHO = 2
       If i% = 0 Then
         If UCase$(Form1.Name) = "SHOWREP07" Then GoTo 80

         If ANCHO_COLUMNA_0% > 0 Then
           ANCHO% = ANCHO_COLUMNA_0% 'SI  SE PASA EL ANCHO DE LA COLUMNA 0 COMO PARAMETRO TOMA ESTE
         Else
           ANCHO% = 3  'EN CASO CONTRARIO ASIGNA VALOR POR DEFECTO
         End If
         NUCOLU% = i%:   Form1.TEPRUEBA = String$(ANCHO%, "A"): MSF1.ColWidth(NUCOLU%) = Form1.TEPRUEBA.Width: MSF1.ColAlignment(NUCOLU%) = ALINEA%: MSF1.TextMatrix(0, NUCOLU%) = NOCAMPO$(i%)
         GoTo 50
       End If

30     LETRX$ = "A"
       ANCHO% = XVALO(Mid$(FORMX$(i%), 2))
       If XVALO(Mid$(FORMX$(i%), 2)) > 0 Then
          ANCHO% = ANCHO% + 2
       ElseIf LETRX$ = "Z" Then

       Else
          ANCHO% = 0 'COLUMNAS OCULTAS PASADAS CON F0.0
       End If
       LETR$ = UCase$(Left$(FORMX$(i%), 1))
       ALINEA% = 1

       If LETR$ = "I" Or LETR$ = "G" Or LETR$ = "F" Then
          ALINEA% = 6
       ElseIf LETR$ = "Z" Then
          ALINEA% = 3
       End If
       '
       If ANCHO% > 0 Then
          NUCOLU% = i% - 1: Form1.TEPRUEBA = String$(ANCHO%, "A"): MSF1.ColWidth(NUCOLU%) = Form1.TEPRUEBA.Width: MSF1.ColAlignment(NUCOLU%) = ALINEA%: MSF1.TextMatrix(0, NUCOLU%) = NOCAMPO$(i%)
       Else ' PARA QUE NO SE VEAN LAS COLUMNAS OCULTAS
          NUCOLU% = i% - 1: Form1.TEPRUEBA = String$(ANCHO%, "A"): MSF1.ColWidth(NUCOLU%) = 0: MSF1.ColAlignment(NUCOLU%) = ALINEA%: MSF1.TextMatrix(0, NUCOLU%) = NOCAMPO$(i%)
       End If
       FORCAMPO$(NUCOLU% + 1) = LETR$
       FORCOLUM$(NUCOLU% + 1) = FORMX$(NUCOLU% + 1)
       MSF1.col = NUCOLU% - 1
       MSF1.CellFontBold = True

       ANCHO_TOTAL = ANCHO_TOTAL + Form1.TEPRUEBA.Width

       'A PARTIR DE AQUI CODIGO PARA CARGAR ENCABEZADO DE GENREP07

50     If UCase$(Form1.Name) = "SHOWREP07" Then
          Form1.Label5.FontName = Form1.GRID.CellFontName
          Form1.Label5.FontSize = Form1.GRID.CellFontSize
          Form1.Label5.FontBold = Form1.GRID.CellFontBold
          ' Label5.Caption = TEXTBIN$("1.234567", FORMAECO$(I&)) + "*"
          Form1.Label5.Caption = String$(ANCHO%, "A")    ' PROVISORIAMENTE A VER SI FUNCIONA
          Form1.GRID.col = i%
          Form1.GRID.ColWidth(i% - 1) = SHOWREP07.Label5.Width
25        If i% >= Form1.Label8.Count Then
             NUPIC% = Form1.Label8.Count
             Load Form1.Label8(NUPIC%)
             GoTo 25
          End If
          Form1.Label8(i%) = Str$(Len(Form1.Label5) - 1)
          If i% = 0 Then Form1.Label8(i%) = 0
       End If
80  Next i%
    '
    If NUPIC% = i% - 1 Then
      If UCase$(Form1.Name) = "SHOWREP07" Then Load Form1.Label8(NUPIC% + 1): Load Form1.Label8(NUPIC% + 2)
    End If

End Sub
'
'
 'COMO ORDENAR UN FLEXGRID
 
''''''MSF.COL = XX 'xx Columan por la q quieres ordenar
''''''MSF.Sort = 1 'Aplicar orden al grid
''''''Valores del orden
''''''0 Ninguna. No se efectúa la ordenación.
''''''1 Genérica ascendente. Realiza una ordenación ascendente basada en la estimación de si el texto es una cadena o un número.
''''''2 Genérica descendente. Realiza una ordenación descendente basada en la estimación de si el texto es una cadena o un número.
''''''3 Numérica ascendente. Realiza una ordenación ascendente que convierte las cadenas a números.
''''''4 Numérica descendente. Realiza una ordenación descendente que convierte las cadenas a números.
''''''5 Cadena ascendente. Realiza una ordenación ascendente sin distinguir mayúsculas y minúsculas en la comparación de cadenas.
''''''6 Cadena descendente. Realiza una ordenación descendente sin distinguir mayúsculas y minúsculas en la comparación de cadenas.
''''''7 Cadena ascendente. Realiza una ordenación ascendente distinguiendo mayúsculas y minúsculas en la comparación de cadenas.
''''''8 Cadena descendente. Realiza una ordenación descendente distinguiendo mayúsculas y minúsculas en la comparación de cadenas.
''''''9 Personalizada. Utiliza el evento Compare para comparar filas.
''''''---------------------------------------------------------------------------------------------------------------------------------

    



Function Sumar(msFlexGrid As msFlexGrid, COLUMNA As Integer) As Currency
                  
   
   On Error GoTo Error_Function
  
    With msFlexGrid
        Dim Total As Currency
        Dim i As Long
           
        If COLUMNA > .Cols Then
           MsgBox "Columna no válida", vbExclamation
           Exit Function
        End If
           
        ' recorrer  las filas de la grilla
        For i = 1 To .Rows - 1
            ' comprobar que el dato es de tipo numérico con la función IsNumeric de vb
            If IsNumeric(.TextMatrix(i, COLUMNA)) Then
                ' Sumar, obteniendo el valor de la celda con TextMatrix
                Total = Total + .TextMatrix(i, COLUMNA)
            End If
        Next
           
        ' retornar el total de la suma a la función
        Sumar = Total
    End With
       
Exit Function
Error_Function:
  
    MsgBox Err.Description, vbCritical, "error al sumar"
  
       
End Function

Sub OrdenarPorFecha(GRID As msFlexGrid, colFecha As Integer)
    'esta rutina ordena las fechas cronologicamene de una columna y luego ordena el reto de las celdas del mismo registro
    Dim i As Integer, J As Integer
    Dim N As Integer, m As Integer
    Dim tempFecha As Date
    Dim tempValue As Variant
    
    N = GRID.Rows - 1
    m = GRID.Cols - 1
    
    For i = 1 To N - 1
        For J = i + 1 To N
            If CDate(GRID.TextMatrix(J, colFecha)) < CDate(GRID.TextMatrix(i, colFecha)) Then
                ' Intercambiar valores de fecha y valores de otras columnas
                tempFecha = CDate(GRID.TextMatrix(i, colFecha))
                GRID.TextMatrix(i, colFecha) = GRID.TextMatrix(J, colFecha)
                GRID.TextMatrix(J, colFecha) = Format(tempFecha, "dd/mm/yy") ' Formatear fecha de nuevo
                
                ' Intercambiar valores de otras columnas
                For k = 1 To m
                    If k <> colFecha Then
                        tempValue = GRID.TextMatrix(i, k)
                        GRID.TextMatrix(i, k) = GRID.TextMatrix(J, k)
                        GRID.TextMatrix(J, k) = tempValue
                    End If
                Next k
            End If
        Next J
    Next i
End Sub


Sub AgregarMesAñoAColenFinalSegunColumnaFecha(MSF As msFlexGrid, columnaFecha%)
    Dim i As Integer
    Dim J As Integer
    Dim dt As Date
    Dim sMonthYear As String
    
    'Agregar una columna al final de la grilla
    MSF.Cols = MSF.Cols + 1
    
    'Recorrer todas las filas de la grilla
    For i = 1 To MSF.Rows - 1
        'Comprobar si el valor del campo es una fecha
        If IsDate(MSF.TextMatrix(i, columnaFecha%)) Then
            'Obtener la fecha del campo
            dt = CDate(MSF.TextMatrix(i, columnaFecha%))
            'Obtener el mes y año de la fecha
            sMonthYear = Format(dt, "mmmm yyyy")
            'Escribir el valor en la nueva columna
            MSF.TextMatrix(i, MSF.Cols - 1) = sMonthYear
        End If
    Next i
End Sub
Sub COLOREAR_GRILLASQL(MSF As msFlexGrid, COLOR As Variant, Optional COLUM1&)

  ' COLUM1& = 0 'COMIENZA A PINTAR A PARTIR DE LA COLUMNA 1
  ' COLUM1& = 1 'COMIENZA A PINTAR A PARTIR DE LA COLUMNA 0
   
   If COLUM1& = 1 Then
      COLUM1& = 0
   Else
      COLUM1& = 1
   End If
   
   With MSF
   .Redraw = False
   For iRow% = 2 To .Rows - 1 Step 2 ' RECORRE LOS RENGLONES
        .row = iRow%                 ' ASIGNA EL NUMERO DE REGISTRO POSICIONADO
        For iCol% = COLUM1& To .Cols - 1 ' DENTRO DEL RENGLON RECORRE CADA COLUMNA
          .col = iCol%
          
          .CellBackColor = COLOR
         
        Next iCol%
   Next iRow%
   .Redraw = True
   End With
End Sub
Sub COLOREAR_GRILLASQL2(MSF As MSHFlexGrid, COLOR As Variant, Optional COLUM1&)

  ' COLUM1& = 0 'COMIENZA A PINTAR A PARTIR DE LA COLUMNA 1
  ' COLUM1& = 1 'COMIENZA A PINTAR A PARTIR DE LA COLUMNA 0

   If COLUM1& = 1 Then
      COLUM1& = 0
   Else
      COLUM1& = 1
   End If

   With MSF
   .Redraw = False
   For iRow% = 2 To .Rows - 1 Step 2 ' RECORRE LOS RENGLONES
        .row = iRow%                 ' ASIGNA EL NUMERO DE REGISTRO POSICIONADO
        For iCol% = COLUM1& To .Cols - 1 ' DENTRO DEL RENGLON RECORRE CADA COLUMNA
          .col = iCol%

          .CellBackColor = COLOR

        Next iCol%
   Next iRow%
   .Redraw = True
   End With
End Sub

Sub COLOREAR_GRILLASQL_TODAS_MISMO_COLOR(MSF As msFlexGrid, COLOR As Variant, Optional COLUM1&)

  ' COLUM1& = 0 'COMIENZA A PINTAR A PARTIR DE LA COLUMNA 1
  ' COLUM1& = 1 'COMIENZA A PINTAR A PARTIR DE LA COLUMNA 0
   
   If COLUM1& = 1 Then
      COLUM1& = 0
   Else
      COLUM1& = 1
   End If
   
   With MSF
   .Redraw = False
   For iRow% = 1 To .Rows - 1  ' RECORRE LOS RENGLONES
        .row = iRow%                 ' ASIGNA EL NUMERO DE REGISTRO POSICIONADO
        For iCol% = COLUM1& To .Cols - 1 ' DENTRO DEL RENGLON RECORRE CADA COLUMNA
          .col = iCol%
          .CellBackColor = COLOR
        Next iCol%
   Next iRow%
   .Redraw = True
   End With
End Sub
Sub COLOREAR_GRILLASQL_VERTICAL(MSF As msFlexGrid, COLOR As Variant, Optional COLUM1&, Optional INTERVALO%)

  ' COLUM1& = 0 'COMIENZA A PINTAR A PARTIR DE LA COLUMNA 1
  ' COLUM1& = 1 'COMIENZA A PINTAR A PARTIR DE LA COLUMNA 0
   
   If COLUM1& = 1 Then
      COLUM1& = 0
   Else
      COLUM1& = 1
   End If
   
   With MSF
   .Redraw = False
   For iRow% = 1 To .Rows - 1  ' RECORRE LOS RENGLONES
        .row = iRow%                 ' ASIGNA EL NUMERO DE REGISTRO POSICIONADO
        For iCol% = COLUM1& To .Cols - 1  ' DENTRO DEL RENGLON RECORRE CADA COLUMNA
           If INTERVALO% > 0 Then
              VUELTA% = VUELTA% + 1
              If INTERVALO% + 1 <= VUELTA% Then
                
                 .col = iCol%
                 .CellBackColor = COLOR
                 '
                 PASO% = PASO% + 1
                 If PASO% = 2 Then
                    VUELTA% = 0
                    PASO% = 0
                    
                 End If
                    
              End If
              GoTo 20
           End If
          .col = iCol%
'          .CellBackColor = COLOR
20      Next iCol%
   Next iRow%
   .Redraw = True
   End With
End Sub

Sub COLOREAR_GRILLASQL_CadenaDeColumnas(MSF As msFlexGrid, COLOR As Variant, Columnas$)
    'A Columnas$ SE LE PASA UN STRING CON LOS NUMEROS DE COLUMNAS A COLOREAR ("4,5;8;9,11")
    Dim strArray() As String
    strArray = Split(Columnas$, ";")
    CANTVECTOR = UBound(strArray)
    With MSF
     For i% = 0 To CANTVECTOR
       COLUMNANUMERO% = XVALO(strArray(i%))
       .Redraw = False
       For iRow% = 1 To .Rows - 1  ' RECORRE LOS RENGLONES
         .row = iRow%                 ' ASIGNA EL NUMERO DE REGISTRO POSICIONADO
         .col = COLUMNANUMERO%
         .CellBackColor = COLOR
       Next iRow%
       .Redraw = True
    Next i%
   End With

End Sub


Sub AJUSTAR_COLUMNAS(MSF As msFlexGrid)
  'ANCHOINICIAL& = TextWidth(MSF.TextMatrix(0, U&))
  'AA = TextWidth(MSF.TextMatrix(0, 1))
  Dim ANCHO&()
  ReDim ANCHO&(MSF.Cols)
  For i& = 0 To MSF.Rows - 1
  
    For U& = 0 To MSF.Cols - 1
    
       ANCHO&(U&) = Len(TRIM(MSF.TextMatrix(i&, U&)))
       If ANCHO(U&) * 190 > MSF.ColWidth(U&) Then
         MSF.ColWidth(U&) = ANCHO(U&) * 150
       End If
    Next U&
  
  
  Next i&


End Sub
Sub CARGAMANUAL(Form As Form, SQLR$)
     'Form.MSF.Rows = 2
     Static ConsultaManual
     ConsultaManual = True
     Set AUX = BASSSE.Execute(SQLR$)
     With AUX
      For i& = 0 To .Fields.Count - 1 ' CARGA ENCABEZADOS
          Form.MSF(0).Cols = .Fields.Count + 1
         CAMP$ = .Fields(i&).Name
         Form.MSF(0).TextMatrix(0, i& + 1) = CAMP$
      Next i&
      Form.MSF(0).Rows = 1
      ROWX& = 1
      Set AUX = BASSSE.Execute(SQLR$)
      Do While Not .EOF
        Form.MSF(0).Rows = Form.MSF(0).Rows + 1

        For J& = 0 To .Fields.Count - 1

         Form.MSF(0).TextMatrix(ROWX&, J& + 1) = SAFETEXT$(.Fields(J&).Value)

        Next J&
        .MoveNext
        ROWX& = ROWX& + 1
       Loop
     End With
     'Call Command4_Click ' boton de flecha atras
     
    
End Sub
'
Sub Carga_MSF_Recordset(rs, Campos$, MSF As msFlexGrid, Optional EsSelecho%, Optional DejarFilaLibre%)
    Dim NOCAMPO$(), FORMX$()
    ReDim NOCAMPO$(1), FORMX$(1)
    LSTX$ = Campos$
    While LSTX$ <> ""
       PPXX% = InStr(LSTX$, ";")
       PPXX2% = InStr(LSTX$, ",")
       If PPXX2% > 0 And PPXX2% < PPXX% Then PPXX% = PPXX2%
       If PPXX% < 1 Then PPXX% = 1 + Len(LSTX$)
       '
       CANCA% = CANCA% + 1
       If CANCA% >= UBound(NOCAMPO$) Then
          ReDim Preserve NOCAMPO$(CANCA%), FORMX$(CANCA%)
       End If
       NOCX$ = Left$(LSTX$, PPXX% - 1)
       LSTX$ = Mid$(LSTX$, PPXX% + 1)
       '
       PPXX% = InStr(NOCX$, "/")
       If PPXX% < 1 Then PPXX% = 1 + Len(NOCX$)
       NOCAMPO$(CANCA%) = Left$(NOCX$, PPXX% - 1)
       FORMX$(CANCA%) = UCase$(Mid$(NOCX$, PPXX% + 1))
       If FORMX$(CANCA%) = "" Then FORMX$(CANCA%) = "A16"
       '
    Wend
    
      MSF.Rows = 100
      ROWX& = 1
      With rs
      Do While Not .EOF
        'MSF.Rows = MSF.Rows + 1
        DEJALIBRE% = 0
        For J& = 0 To .Fields.Count - 1
          If J& + 1 > MSF.Cols - 1 Then MSF.Cols = J& + 2
          If EsSelecho% < 1 Then
            If VENTANA.DejarLibre(J& + 1) = "S" Then
              DEJALIBRE% = DEJALIBRE% + 1 ' PARA QUE DEJE LIBRE COLUMNAS SIN LLENAR DEL RECORDSET
              'INTERVENCION PARA QUE CARGUE EL ECO DIRECTAMENTE SI ESTA CONFIGURADO VENTANA.ColEcoCampo(J& + 0)
              'SE UTILIZO EN ORDEN DE FABRIC > FRM PROCEFRM > mnuppimpresora_Click
              If VENTANA.ColEcoCampo(J& + 0) <> "" Then 'si tiene campo eco lo asigna directamente.
                 If TRIM$(MSF.TextMatrix(ROWX&, J&)) <> "" Then
                    Dim strArray() As String
                    On Error Resume Next
                    strArray = Split(VENTANA.ColEcoCampo(J& + 0), ";")
                    CAMPO1$ = strArray(0)
                    campo2$ = strArray(1)
                    TABLAX1$ = strArray(2)
                    valor = TRIM$(MSF.TextMatrix(ROWX&, J&))
                    CONDI$ = CAMPO1$ & " = '" & valor & "'"
                    MSF.TextMatrix(ROWX&, J& + DEJALIBRE%) = VALOBADA(TABLAX1$, campo2$, CONDI$, "NOMESS")
                    On Error GoTo 0
                 End If
              End If
              'FIN INTERVENCION PARA QUE CARGUE EL ECO DIRECTAMENTE SEGUN CONFIGURACION DE CLASE VENTANA
              MSF.Cols = MSF.Cols + 1
            End If
          End If
          LEIDO$ = SAFETEXT$(.Fields(J&).Value)
          A = .Fields(J&).Name
          FORMA1$ = UCase$(Left$(FORMX$(J& + 1), 1))
          FORMACOMP$ = FORMX$(J& + 1)
          If Left$(FORMACOMP$, 1) = "F" Then
             MSF.TextMatrix(ROWX&, J& + 1 + DEJALIBRE%) = FORMATNUM$(XVALO(LEIDO$), "" & FORMACOMP$ & "")
          ElseIf FORMA1$ = "D" Then
             MSF.TextMatrix(ROWX&, J& + 1 + DEJALIBRE%) = FECHATEX$(CVINT(LEIDO$))
          Else
             MSF.TextMatrix(ROWX&, J& + 1 + DEJALIBRE%) = LEIDO$
          End If
        Next J&
        .MoveNext
        ROWX& = ROWX& + 1
        If DejarFilaLibre% > 0 Then
          ROWX& = ROWX& + DejarFilaLibre%
        End If
        If ROWX& + 1 = MSF.Rows Then
           MSF.Rows = MSF.Rows + 100
        End If

       Loop
     End With
     MSF.Rows = ROWX&
     If ROWX& < 2 Then MSF.Rows = 2
 End Sub
 Sub LIMPIAR_MSF_SIN_BORRAR_REGISTROS(MSF As msFlexGrid, Optional APARTIR_COLUM%)
    If APARTIR_COLUM% < 1 Then APARTIR_COLUM% = 1
    For i& = 1 To MSF.Rows - 1
      For J& = APARTIR_COLUM% To MSF.Cols - 1
        MSF.TextMatrix(i&, J&) = ""
      Next J&
    Next i&
End Sub
 Sub LIMPIAR_MSF_1_COLUMNA_SIN_BORRAR_REGISTROS(MSF As msFlexGrid, COLUM%)
    For i& = 1 To MSF.Rows - 1
      For J& = APARTIR_COLUM% To MSF.Cols - 1
        If J& = COLUM% Then
          MSF.TextMatrix(i&, J&) = ""
        End If
      Next J&
    Next i&
End Sub
Sub POSICIONAR_TXTMOVIL(TXTMOVIL As TextBox, MSF As msFlexGrid, REG&, col&)
  TXTMOVIL.ToolTipText = REG& & "," & col&
  TXTMOVIL.TEXT = ""
  With TXTMOVIL
    TXTMOVIL = MSF.TextMatrix(REG&, col&)
    .ZOrder 0
'    If MSF.CellWidth <= 0 Then Exit Sub
    .Width = MSF.CellWidth
    .Left = MSF.CellLeft + MSF.Left
    .Top = MSF.CellTop + MSF.Top
    .Visible = True
    On Error Resume Next
    .SetFocus
    On Error GoTo 0
  End With
End Sub
Sub COLOREAR_GRILLA_SOLAFILA(MSF As msFlexGrid, COLOR As Variant, Optional REGISTRO, Optional COLUMNA%, Optional NO_ASTERISCO%, Optional NOBLANQUEAR%)
   'RECORRE LA FILA SELECCIONADA Y PINTA CADA CELDA
   'SI LA CELDA ES DIFERENTE AL COLOR PASADO COMO PARAMETRO LA PINTA , SI NO LA PONE EN BLANCO
   
   
   With MSF
    If XVALO(REGISTRO) > 0 Then
      iRow% = XVALO(REGISTRO)
    Else
      iRow% = MSF.row
    End If
    .row = iRow%                 ' ASIGNA EL NUMERO DE REGISTRO POSICIONADO
    INICOL% = 1: If COLUMNA% > 0 Then INICOL% = COLUMNA%
    For iCol% = INICOL% To .Cols - 1 ' DENTRO DEL RENGLON RECORRE CADA COLUMNA
        .col = iCol%
        If .CellBackColor <> COLOR And .CellBackColor <> vbYellow Then
          .CellBackColor = COLOR
          If iCol% = 1 Then
            .col = 0
            If NO_ASTERISCO% < 1 Then .TEXT = "*"
          End If
        Else
          
          If NOBLANQUEAR% < 1 Then .CellBackColor = 0 ' SI ES OTRO COLOR PINTA DE BLANCO
          If iCol% = 1 Then
            .col = 0
            If NO_ASTERISCO% < 1 Then .TEXT = ""
          End If
        End If
    Next iCol%
   End With

End Sub

Sub COLOREAR_GRILLA_SOLOCOLUMNA(MSF As msFlexGrid, COLOR As Variant, Optional Fila%, Optional COLUMNA%)
   'RECORRE LA COLUMNA SELECCIONADA Y PINTA CADA CELDA
   
   If COLUMNA% < 1 Or COLUMNA% > MSF.Cols - 1 Then Exit Sub
   If Fila% < 1 Or Fila% > MSF.Rows - 1 Then Exit Sub
      
   With MSF
        .col = COLUMNA%
        For iRow% = Fila% To .Rows - 1
            .row = iRow%
            .CellBackColor = COLOR
        Next iRow%
   End With

End Sub
Sub COLOREAR_GRILLA_SOLAFILA_2(MSF As msFlexGrid, COLOR As Variant, Optional REGISTRO, Optional COLUMNA%, Optional NO_ASTERISCO%, Optional NOBLANQUEAR%)
   'RECORRE LA FILA SELECCIONADA Y PINTA CADA CELDA
   'SI LA CELDA ES DIFERENTE AL COLOR PASADO COMO PARAMETRO LA PINTA , SI NO LA PONE EN BLANCO
   'ESTA ES UNA SEGUNDA VERSION DONDE SI SE MANDA  NO BLANQUEAR = 1 TAMPOCO BORRE EL ASTERIZCO
   'SE GENERO ESTA NUEVA VERSION POR MIEDO A QUE CAMBIE EN OTROS CASOS
   
   With MSF
    If XVALO(REGISTRO) > 0 Then
      iRow% = XVALO(REGISTRO)
    Else
      iRow% = MSF.row
    End If
    .row = iRow%                 ' ASIGNA EL NUMERO DE REGISTRO POSICIONADO
    For iCol% = 1 To .Cols - 1 ' DENTRO DEL RENGLON RECORRE CADA COLUMNA
        .col = iCol%
        If .CellBackColor <> COLOR And CellBackColor <> vbYellow Then
          .CellBackColor = COLOR
          If iCol% = 1 Then
            .col = 0
            If NO_ASTERISCO% < 1 Then .TEXT = "*"
          End If
        Else
          
          If NOBLANQUEAR% < 1 Then .CellBackColor = 0 ' SI ES OTRO COLOR PINTA DE BLANCO
          If iCol% = 1 Then
            .col = 0
            If NO_ASTERISCO% > 0 Then .TEXT = ""
          End If
        End If
    Next iCol%
   End With

End Sub
  Public Sub llevar_al_final_flexgrid(MSF As msFlexGrid)
End Sub

Sub POSICIONAR_ULTIMAFILA(MSF As msFlexGrid)
     With MSF
        If .Rows > 1 Then
          .row = MSF.Rows - 1
          .TopRow = MSF.row
          .RowSel = MSF.row
          .col = 0
          .ColSel = .Cols - 1
        End If
     End With

'       On Error Resume Next
'        MSF.Row = MSF.Rows - 1
'        VISIROWS = (MSF.Height / MSF.RowHeight(J&)) - 4
'        If Err > 0 Then
'          VISIROWS = 1
'        End If
'        On Error GoTo 0
'        '
'        If VISIROWS > 0 Then
'           If MSF.TopRow < MSF.Row - VISIROWS Then
'              MSF.TopRow = MSF.Row - VISIROWS
'           End If
'        End If
End Sub


Public Function borrafila(ndo, Form1 As Form, MSF As msFlexGrid)

        Dim i, J As Integer
        For i = ndo To MSF.Rows - 2
          For J = 0 To MSF.Cols - 1
            MSF.TextMatrix(i, J) = MSF.TextMatrix(i + 1, J)
          Next
        Next
        MSF.Rows = MSF.Rows - 1
End Function


'''Function Ordenar_Columna_FlexGrid( _
'''                FlexGrid As MSFlexGrid, _
'''                ByVal columna As Integer) As SortSettings
''''ORDENA LA COLUMNA DE UN FLEXGRID ESTA COMENTADA PARA QUE SIEMPRE ORDENE DE MANERA ASCENDENTE
'''    FlexGrid.Visible = False
'''    FlexGrid.Refresh
'''
''''    If sCol <> columna Then
'''        Orden = 1
''''    ElseIf Orden = 1 Then
''''        Orden = 2
''''    Else
''''        Orden = 1
''''    End If
'''
'''    'Ordena la columna del FlexGRid según _
'''     el valor de la variable Orden
'''    FlexGrid.Sort = Orden
'''
'''    sCol = columna
'''
'''    'Devuelve a la función el tipo de orden
'''    Ordenar_Columna_FlexGrid = Orden
'''
'''    FlexGrid.Visible = True
''''*************************
''''ESTO ES PARA HACER LA LLAMADA CUANDO SE HACE DOBLE CLIC EN EL ENCABEZADO
''''Y ORDENA DE ACUERDO A LA COLUMNA QUE SE P RESIONA
'''''Private Sub MSF1_MouseUp( _
'''''    Button As Integer, _
'''''    Shift As Integer, _
'''''    x As Single, _
'''''    y As Single)
'''''
'''''    With MSF1
'''''    A = .Row
'''''        If .MouseRow <> 0 Then
'''''           Exit Sub
'''''        End If
'''''
'''''        'La función retorna el tipo de orden, es decir esta declarada como As SortSettings
'''''        If Ordenar_Columna_FlexGrid(MSF1, .MouseCol) = flexSortGenericAscending Then
'''''           Me.Caption = " FlexGrid en orden: Ascendente "
'''''        Else
'''''           Me.Caption = " FlexGrid en orden: Descendente "
'''''        End If
'''''    End With
'''''
'''''End Sub
'''
'''End Function

Sub SELECCION_COLUMNA_0(MSF As msFlexGrid, REGISTRO&, COLUMNA&)
   i& = REGISTRO&
   J& = COLUMNA&
   '
   'VALIDO
   If i& = 0 Or i& > MSF.Rows Then Exit Sub
   '
   'RESTAURO AL COLOR ORIGINAL LA COLUMNA 0
   MSF.col = 0
   For h& = 0 To MSF.Rows - 1
     MSF.row = h&
     MSF.CellBackColor = &H8000000F
   Next h&
   '
   'PINTO DE AMARILLO LA SELECCIONADA
   MSF.row = i&
   MSF.CellBackColor = &H80FFFF   ' LA PINTA DE AMARILLA

End Sub
Sub CARGALISTA_LIST_GRID(NOMBREDELISTA, TABLA$, LISTADECAMPOS$, CONDICION$, Optional MODOSEL$, Optional Encabezado$, Optional COLUMNA_AGRUPA%)
    '
    Dim NOCAMPO$(), FORMX$()
    ReDim NOCAMPO$(1), FORMX$(1)
    '
    MODOSE$ = TRIM$(UCase$(SAFETEXT$(MODOSEL$)))
    '
    On Error Resume Next
    NOMBREDELISTA.Clear
    If Err Then
       On Error GoTo 0
       Call MENSERR(24, "Imposible Limpiar Lista '" & NOMBREDELISTA & "'.")
       Call FINAL("")
    End If
    On Error GoTo 0
    '
    TIPODELISTA = "GRID"
    On Error Resume Next
    CANCOL& = NOMBREDELISTA.Cols
    If Err Then
       On Error GoTo 0
       TIPODELISTA = "LIST"
    End If
    On Error GoTo 0
    '
    If TIPODELISTA = "GRID" Then
       NOMBREDELISTA.Rows = 2
       ALTUREN = NOMBREDELISTA.RowHeight(0) + NOMBREDELISTA.RowHeight(1)
       While ALTUREN < NOMBREDELISTA.Height
          NOMBREDELISTA.Rows = NOMBREDELISTA.Rows + 1
          ALTUREN = ALTUREN + NOMBREDELISTA.RowHeight(NOMBREDELISTA.Rows - 1)
       Wend
       If TRIM$(Encabezado$) <> "" Then
          NOMBREDELISTA.FormatString = Encabezado$
       End If
       
    End If
    '
    LSTX$ = TRIM$(LISTADECAMPOS$)
    CANCA% = 0
    SQLX$ = ""
    While LSTX$ <> ""
       PPXX% = InStr(LSTX$, ";")
       PPXX2% = InStr(LSTX$, ",")
       If PPXX2% > 0 And PPXX2% < PPXX% Then PPXX% = PPXX2%
       If PPXX% < 1 Then PPXX% = 1 + Len(LSTX$)
       '
       CANCA% = CANCA% + 1
       If CANCA% >= UBound(NOCAMPO$) Then
          ReDim Preserve NOCAMPO$(CANCA%), FORMX$(CANCA%)
       End If
       NOCX$ = Left$(LSTX$, PPXX% - 1)
       LSTX$ = Mid$(LSTX$, PPXX% + 1)
       '
       PPXX% = InStr(NOCX$, "/")
       If PPXX% < 1 Then PPXX% = 1 + Len(NOCX$)
       NOCAMPO$(CANCA%) = Left$(NOCX$, PPXX% - 1)
       FORMX$(CANCA%) = Mid$(NOCX$, PPXX% + 1)
       If FORMX$(CANCA%) = "" Then FORMX$(CANCA%) = "A16"
       '
       If TIPODELISTA = "GRID" Then
          NOMBREDELISTA.Cols = CANCA%
       End If
       '
       If SQLX$ <> "" Then SQLX$ = SQLX$ + ", "
       SQLX$ = SQLX$ + NOCAMPO$(CANCA%) + " AS CAMPO" + TRIM$(Str$(CANCA%))
    Wend
    '
    TABLA1$ = TABLA$: TABLA2$ = "": IJO$ = ""
    PPXX% = InStr(TABLA$, "+")
    If PPXX% > 0 Then
       TABLA1$ = TRIM$(Left$(TABLA$, PPXX% - 1))
       TABLA$ = TRIM$(Mid$(TABLA$, PPXX% + 1))
       PPXX% = InStr(TABLA$, "/")
       If PPXX% > 0 Then
         TABLA2$ = TRIM$(Left$(TABLA$, PPXX% - 1))
         IJO$ = TRIM$(Mid$(TABLA$, PPXX% + 1))
         TABLA$ = TABLA1$ + " INNER JOIN " + TABLA2$ + " ON " + IJO$
       End If
    End If
    
    If InStr(MODOSE$, "DISTINCT") > 0 Then SQLX$ = "DISTINCT " + SQLX$
    SQLX$ = "SELECT " + SQLX$ + " FROM " + TABLA$
    If TRIM$(CONDICION$) <> "" Then
       SQLX$ = SQLX$ + " WHERE " & CONDICION$
    End If
    On Error Resume Next
    Set TABTMP = FLEXCONN.Execute(SQLX$)
    If Err Then
       On Error GoTo 0
       MsgBox "Imposible Ejecutar Consulta SQL: " + vbCrLf + vbCrLf & SQLX$, vbExclamation
       Call FINAL("")
    End If
    '
    With TABTMP
       TTXX$ = ""
       J% = 0: Fila& = 0
       Do While Not .EOF
           TTXX$ = "": COLUI% = 0
           '
           If TIPODELISTA = "GRID" Then
              Fila& = Fila& + 1
37            If Fila& > NOMBREDELISTA.Rows - 1 Then
                 NOMBREDELISTA.Rows = NOMBREDELISTA.Rows + 1
                 GoTo 37
              End If
           End If
           '
           For U& = 1 To .Fields.Count
              If U& > 2048 Then Exit For
              valor = .Fields(U& - 1).Value
              FORMATX$ = FORMX$(U&) ' FORMATO
              TECO$ = ""
              PPXX% = InStr(FORMATX$, "+")
              If PPXX% > 0 Then
                 TECO$ = TRIM$(Mid$(FORMATX$, PPXX% + 1))
                 FORMATX$ = Left$(FORMATX$, PPXX% - 1)
              End If
              '
              ENCAB$ = ""
              PPXX% = InStr(FORMATX$, "/")
              If PPXX% > 0 Then
                 ENCAB$ = Mid$(FORMATX$, PPXX% + 1)
                 FORMATX$ = Left$(FORMATX$, PPXX% - 1)
              End If
              TTYY$ = FORMATNUM$(valor, FORMATX$)
              If TRIM$(TTYY$) = "" Then
                  TTYY$ = Space$(XVALO(Mid$(FORMATX$, 2)))
              End If
              If TIPODELISTA = "LIST" Then
                   If TTXX$ <> "" Then TTXX$ = TTXX$ + " "
                   TTXX$ = TTXX$ + TTYY$
              ElseIf TIPODELISTA = "GRID" Then
                   COLUI% = COLUI% + 1
17                 If COLUI% > NOMBREDELISTA.Cols Then
                      NOMBREDELISTA.Cols = NOMBREDELISTA.Cols + 1
                      GoTo 17
                   End If
                   If Fila& = 1 Then
                      FORMALTER.TEPRUEBA.Font = NOMBREDELISTA.Font
                      FORMALTER.TEPRUEBA.FontSize = NOMBREDELISTA.FontSize
                      FORMALTER.TEPRUEBA.Caption = "l" + TMASCARA$(FORMATX$)
                      NOMBREDELISTA.ColWidth(COLUI% - 1) = FORMALTER.TEPRUEBA.Width
                      NOMBREDELISTA.COLALIGNEMENT(COLUI% - 1) = 7
                      If Left$(FORMATX$, 1) = "A" Or Left$(FORMATX$, 1) = "C" Or Left$(FORMATX$, 1) = "D" Then
                         NOMBREDELISTA.COLALIGNEMENT(COLUI% - 1) = 1
                      End If
                      NOMBREDELISTA.FontBold = True
                      NOMBREDELISTA.TextMatrix(0, COLUI% - 1) = ENCAB$
                      NOMBREDELISTA.FontBold = False
                   End If
                   If Left$(FORMATX$, 1) = "A" Then NOMBREDELISTA.COLALIGNEMENT(COLUI% - 1) = 1: TTYY$ = valor   ' PARA QUE EN LA GRILLA LO PONGA COMPLETO
                   
'                   TTYY$ = "HOLA OMAR"
                   NOMBREDELISTA.TextMatrix(Fila&, COLUI% - 1) = TTYY$
              End If
              '
              'PARA QUE AGRUPE POR LA COLUMNA PASADA COMO PARAMETRO
              If COLUMNA_AGRUPA% > 0 Then
                NOMBREDELISTA.MergeCells = flexMergeRestrictColumns
                NOMBREDELISTA.MergeCol(COLUMNA_AGRUPA%) = True
              End If
              '
              If TECO$ <> "" Then
                 PPXX% = InStr(TECO$, "/")
                 If PPXX% < 1 Then PPXX% = 1 + Len(TECO$)
                 FORMATX$ = Mid$(TECO$, PPXX% + 1)
                 TECO$ = TRIM$(UCase$(Left$(TECO$, PPXX% - 1)))
                 '
                 ENCAB$ = ""
                 PPXX% = InStr(FORMATX$, "/")
                 If PPXX% > 0 Then
                    ENCAB$ = Mid$(FORMATX$, PPXX% + 1)
                    FORMATX$ = Left$(FORMATX$, PPXX% - 1)
                 End If
                 If FORMATX$ = "" Then FORMATX$ = "A16"
                 '
                 VALOECO = Null
                 If Left$(TECO$, 5) = "DEUDO" Or Left$(TECO$, 5) = "CLIEN" Then
                   VALOECO = rasocli$(XVALO(valor))
                 ElseIf Left$(TECO$, 5) = "PROVE" Or Left$(TECO$, 5) = "ACRED" Then
                   VALOECO = rasocli$((-1) * XVALO(valor))
                 ElseIf Left$(TECO$, 5) = "MASTE" Then
                   TXVALO$ = valor
                     If Len(TXVALO$) > 4 Then
                          CIXI% = CODINT%(TXVALO$)
                        Else
                          CIXI% = XVALO(valor)
                     End If
                   VALOECO = DESCRIT0$(CIXI%)
                 ElseIf EXISTE%(LUDAT$ + TECO$ + ".RFS") Then
                   LFF% = LENFIELD%(TECO$, 1)
                   ABUS$ = MKS$(XVALO(valor))
                   On Error Resume Next
                   If LFF% = 2 Then ABUS$ = MKI$(Int(XVALO(valor)))
                   If LFF% = 1 Then ABUS$ = Chr$(Int(XVALO(valor)))
                   On Error GoTo 0
                   VALOECO = ECOARCH$(TECO$, ABUS$)
                 End If
                 '
                 TTYY$ = FORMATNUM$(VALOECO, FORMATX$)
                 If TRIM$(TTYY$) = "" Then
                     TTYY$ = Space$(XVALO(Mid$(FORMATX$, 2)))
                 End If
                 If TIPODELISTA = "LIST" Then
                      If TTXX$ <> "" Then TTXX$ = TTXX$ + " "
                      TTXX$ = TTXX$ + TTYY$
                    ElseIf TIPODELISTA = "GRID" Then
                      COLUI% = COLUI% + 1
27                    If COLUI% > NOMBREDELISTA.Cols Then
                         NOMBREDELISTA.Cols = NOMBREDELISTA.Cols + 1
                         GoTo 27
                      End If
                      If Fila& = 1 Then
                         FORMALTER.TEPRUEBA.Font = NOMBREDELISTA.Font
                         FORMALTER.TEPRUEBA.FontSize = NOMBREDELISTA.FontSize
                         FORMALTER.TEPRUEBA.Caption = "l" + TMASCARA$(FORMATX$)
                         NOMBREDELISTA.ColWidth(COLUI% - 1) = FORMALTER.TEPRUEBA.Width
                         NOMBREDELISTA.COLALIGNEMENT(COLUI% - 1) = 7
                         If Left$(FORMATX$, 1) = "A" Or Left$(FORMATX$, 1) = "C" Or Left$(FORMATX$, 1) = "D" Then
                            NOMBREDELISTA.COLALIGNEMENT(COLUI% - 1) = 1
                         End If
                         NOMBREDELISTA.FontBold = True
                         NOMBREDELISTA.TextMatrix(0, COLUI% - 1 + 1) = ENCAB$
                         NOMBREDELISTA.FontBold = False
                      End If
                      If Left$(FORMATX$, 1) = "A" Then TTYY$ = VALOECO       ' PARA QUE EN LA GRILLA LO PONGA COMPLETO
                      NOMBREDELISTA.TextMatrix(Fila&, COLUI% - 1) = TTYY$
                 End If
                 '
              End If
           Next U&
           '
           If TIPODELISTA = "LIST" Then
               NOMBREDELISTA.AddItem TTXX$
           End If
           '
       .MoveNext
       Loop
    End With
    '
    If TIPODELISTA = GRID Then
    End If
    '
    TABTMP.Close
End Sub

Sub COLOREAR_GRILLA_SOLO_UNA_AFILA(MSF As msFlexGrid, COLOR As Variant, Optional REGISTRO)
   'RECORRE LA FILA SELECCIONADA Y PINTA CADA CELDA
   
   With MSF
    If XVALO(REGISTRO) > 0 Then
      iRow% = XVALO(REGISTRO)
    Else
      iRow% = MSF.row
    End If
    
    .row = iRow%                 ' ASIGNA EL NUMERO DE REGISTRO POSICIONADO
    For iCol% = 1 To .Cols - 1 ' DENTRO DEL RENGLON RECORRE CADA COLUMNA
        .col = iCol%
          .CellBackColor = COLOR
    Next iCol%
   End With
End Sub
Sub COLOREAR_TEXTO_SOLO_UNA_AFILA(MSF As msFlexGrid, COLOR As Variant, Optional REGISTRO, Optional COLUM%)
   'RECORRE LA FILA SELECCIONADA Y PINTA CADA CELDA
   
   With MSF
    If XVALO(REGISTRO) > 0 Then
      iRow% = XVALO(REGISTRO)
    Else
      iRow% = MSF.row
    End If
    
    .row = iRow%                 ' ASIGNA EL NUMERO DE REGISTRO POSICIONADO
    For iCol% = 1 To .Cols - 1 ' DENTRO DEL RENGLON RECORRE CADA COLUMNA
        .col = iCol%
        If COLUM% > 0 Then     'SI SE LE PASA UN NUMERO DE COLUMNA SOLO LO CAMBIE DICHA COLUMNA
          If .col = COLUM% Then
            .col = iCol%
            .CellForeColor = COLOR
          End If
        Else
          .CellForeColor = COLOR
        End If
    Next iCol%
   End With
End Sub

Function VALIDA_COD_GRID%(Dato$, MSFXX As msFlexGrid, COLUM%, Optional MENSAXE$)
   'VALIDA SI EN LA COLUMNA PASADA COMO PARAMETRO EL DATO EXISTE, SI ES ASI PRESENTA UN MENSAJE TAMBIEN PARAMETRIZADO
   
   DATO1$ = TRIM$(UCase$(Dato$))
   For i& = 1 To MSFXX.Rows - 1
     LECTURA_GRID$ = TRIM$(UCase$(MSFXX.TextMatrix(i&, COLUM%)))
     If LECTURA_GRID$ = DATO1$ Then
        If TRIM$(MENSAXE$) = "" Then
          TXT$ = "Código Existente en Esta Lista " + vbCrLf
          TXT$ = TXT$ + "Edite Realizando Doble Click Sobre El Mismo"
        Else
          TXT$ = MENSAXE$
        End If
        If TRIM$(MENSAXE$) <> "" Then
         Call COMUNI(TXT$)
        End If
        MSFXX.row = i&
        If MSFXX.TopRow > MSFXX.row Then MSFXX.TopRow = MSFXX.row
         On Error Resume Next
         VISIROWS = (MSFXX.Height / MSFXX.RowHeight(J&)) - 4
         If Err > 0 Then
          VISIROWS = 1
         End If
         On Error GoTo 0
         If VISIROWS > 0 Then
             If MSFXX.TopRow < MSFXX.row - VISIROWS Then
               MSFXX.TopRow = MSFXX.row - VISIROWS
             End If
         End If
         On Error Resume Next
         MSFXX.SetFocus
         On Error GoTo 0
         If MSFXX.row = 0 Then Exit Function
         '
         VALIDA_COD_GRID = i&
         Exit Function

         '
     End If

     
   Next i&
   VALIDA_COD_GRID% = 0
End Function
Sub CARGA_ENCABEZADO_VIE(MSF1 As msFlexGrid, LISTADECAMPOS$, Form1 As Form)
    '
    Dim NOCAMPO$(), FORMX$()
    ReDim NOCAMPO$(1), FORMX$(1)
    '
    MODOSE$ = UCase$(SAFETEXT$(MODOSEL$))
    LSTX$ = TRIM$(LISTADECAMPOS$)
    CANCA% = 0
    ATOTA% = 1
    While LSTX$ <> ""
       PPXX% = InStr(LSTX$, ";")
       PPXX2% = InStr(LSTX$, ",")
       If PPXX2% > 0 And PPXX2% < PPXX% Then PPXX% = PPXX2%
       If PPXX% < 1 Then PPXX% = 1 + Len(LSTX$)
       '
       CANCA% = CANCA% + 1
       If CANCA% >= UBound(NOCAMPO$) Then
          ReDim Preserve NOCAMPO$(CANCA%), FORMX$(CANCA%)
       End If
       NOCX$ = Left$(LSTX$, PPXX% - 1)
       LSTX$ = Mid$(LSTX$, PPXX% + 1)
       '
       PPXX% = InStr(NOCX$, "/")
       If PPXX% < 1 Then PPXX% = 1 + Len(NOCX$)
       NOCAMPO$(CANCA%) = Left$(NOCX$, PPXX% - 1)
       FORMX$(CANCA%) = Mid$(NOCX$, PPXX% + 1)
       If FORMX$(CANCA%) = "" Then FORMX$(CANCA%) = "A16"
       '
       ATOTA% = ATOTA% + XVALO(Mid$(FORMX$(CANCA%), 2)) + 2
       '
'       TECO$ = ""
'       PPXX% = InStr(FORMX$(CANCA%), "+")
'       If PPXX% > 0 Then
'          TECO$ = UCase$(TRIM$(Mid$(FORMX$(CANCA%), PPXX% + 1)))
'          PPXX% = InStr(TECO$, "/")
'          If PPXX% > 0 Then
'             FORMATX$ = TRIM$(Mid$(TECO$, PPXX% + 1))
'             ATOTA% = ATOTA% + Len(TMASCARA$(FORMATX$))
'          End If
'       End If
       
       If CANCA% = 1 Then ACOL1% = XVALO(Mid$(FORMX$(CANCA%), 2))
       '
    Wend

    MSF1.Cols = CANCA% + 1
    MSF1.Rows = 1
    For i% = 0 To CANCA%
       'PRIMER COLUMNA ASIGNA POR DEFECTO ANCHO = 2
       If i% = 0 Then
         ANCHO% = 2
         NUCOLU% = i%:   Form1.TEPRUEBA = String$(ANCHO%, "A"): MSF1.ColWidth(NUCOLU%) = Form1.TEPRUEBA.Width: MSF1.ColAlignment(NUCOLU%) = ALINEA%: MSF1.TextMatrix(0, NUCOLU%) = NOCAMPO$(i%)
         GoTo 50
       End If
       
       ANCHO% = XVALO(Mid$(FORMX$(i%), 2)) + 2
       LETR$ = UCase$(Left$(FORMX$(i%), 1))
       ALINEA% = 1
       If LETR$ = "I" Or LETR$ = "G" Or LETR$ = "F" Then ALINEA% = 6
       NUCOLU% = i%:   Form1.TEPRUEBA = String$(ANCHO%, "A"): MSF1.ColWidth(NUCOLU%) = Form1.TEPRUEBA.Width: MSF1.ColAlignment(NUCOLU%) = ALINEA%: MSF1.TextMatrix(0, NUCOLU%) = NOCAMPO$(i%)
50  Next i%
End Sub

'

Sub CARGA_ENCABEZADO(MSF1 As msFlexGrid, LISTADECAMPOS$, Form1 As Form, Optional FiltroCombo%, Optional NOREPIT%, Optional ANCHO_TOTAL, Optional ANCHO_COLUMNA_0%, Optional PINTAFILACLICK%, Optional ApartirdeColum%)
    '
    
    Dim NOCAMPO$(), FORMX$()
    ReDim NOCAMPO$(1), FORMX$(1)
    On Error Resume Next
    Form1.CMBFILTRO.Clear
    Form1.CMBSERIE1.Clear
    Form1.CMBSERIE2.Clear
    Form1.CMBSERIE3.Clear
    Form1.CMBLABEL.Clear
    On Error GoTo 0
    '
    ANCHO_TOTAL = 0
    On Error Resume Next
    If Form1.Label8.Count > 1 Then
      Erase Form1.Label8
    End If
    
    If FiltroCombo% > 0 Then
       Form1.CMBFILTRO.AddItem ""
       Form1.CMBSERIE1.AddItem ""
       Form1.CMBSERIE2.AddItem ""
       Form1.CMBSERIE3.AddItem ""
       Form1.CMBLABEL.AddItem ""
    End If
    On Error GoTo 0
    NUPIC% = 0
    MODOSE$ = UCase$(SAFETEXT$(MODOSEL$))
    LSTX$ = TRIM$(LISTADECAMPOS$)
    CANCA% = 0 + ApartirdeColum%
    ATOTA% = 1
    While LSTX$ <> ""
       PPXX% = InStr(LSTX$, ";")
       PPXX2% = InStr(LSTX$, ",")
       If PPXX2% > 0 And PPXX2% < PPXX% Then PPXX% = PPXX2%
       If PPXX% < 1 Then PPXX% = 1 + Len(LSTX$)
       '
       CANCA% = CANCA% + 1
       If CANCA% >= UBound(NOCAMPO$) Then
          ReDim Preserve NOCAMPO$(CANCA%), FORMX$(CANCA%)
       End If
       NOCX$ = Left$(LSTX$, PPXX% - 1)
       LSTX$ = Mid$(LSTX$, PPXX% + 1)
       '
       PPXX% = InStr(NOCX$, "/")
       If PPXX% < 1 Then PPXX% = 1 + Len(NOCX$)
       NOCAMPO$(CANCA%) = Left$(NOCX$, PPXX% - 1)
       FORMX$(CANCA%) = Mid$(NOCX$, PPXX% + 1)
       '
       'AGREGO EL LOS CAMPOS AL FILTRO y para los graficos
       If FiltroCombo% > 0 Then
          If XVALO(Mid$(FORMX$(CANCA%), 2)) > 0 Then
             Form1.CMBFILTRO.AddItem NOCAMPO$(CANCA%)
             On Error Resume Next 'CODIGO AGREGADO PARA GENERAR GRAFICOS
             LETR1$ = UCase$(Left$(FORMX$(CANCA%), 1))
             If LETR1$ = "I" Or LETR1$ = "G" Or LETR1$ = "F" Then
                Call REPLA(TX$, (NOCAMPO$(CANCA%)), 1, 48): Call REPLA(TX$, TRIM$(CStr(CANCA%)), 50, 4): Call REPLA(TX$, "F", 56, 4)
                Form1.CMBSERIE1.AddItem TX$ 'SI ES NUMERICO CARGA PARA LAS SERIES
                Form1.CMBSERIE2.AddItem TX$
                Form1.CMBSERIE3.AddItem TX$
                Form1.CMBLABEL.AddItem TX$ 'en el eje lo pongo tambien, estoy parobando que esten todos en el eje
             Else
                Call REPLA(TX$, (NOCAMPO$(CANCA%)), 1, 48): Call REPLA(TX$, TRIM$(CStr(CANCA%)), 50, 4): Call REPLA(TX$, LETR1$, 56, 4)
                Form1.CMBLABEL.AddItem TX$ 'SI ES TEXTO VA AL COMOB DE LABEL
             End If
             On Error GoTo 0
          End If
       End If

       If FORMX$(CANCA%) = "" Then FORMX$(CANCA%) = "A16"
       '
       ATOTA% = ATOTA% + XVALO(Mid$(FORMX$(CANCA%), 2)) + 2
       '
'       TECO$ = ""
'       PPXX% = InStr(FORMX$(CANCA%), "+")
'       If PPXX% > 0 Then
'          TECO$ = UCase$(TRIM$(Mid$(FORMX$(CANCA%), PPXX% + 1)))
'          PPXX% = InStr(TECO$, "/")
'          If PPXX% > 0 Then
'             FORMATX$ = TRIM$(Mid$(TECO$, PPXX% + 1))
'             ATOTA% = ATOTA% + Len(TMASCARA$(FORMATX$))
'          End If
'       End If
       
       If CANCA% = 1 Then ACOL1% = XVALO(Mid$(FORMX$(CANCA%), 2))
       '
    Wend

    MSF1.Cols = CANCA% + 1
    MSF1.Rows = 1
    For i% = 0 To CANCA%
       'PRIMER COLUMNA ASIGNA POR DEFECTO ANCHO = 2

       If i% = 0 Then
         If UCase$(Form1.Name) = "SHOWREP07" Then GoTo 80

         If ANCHO_COLUMNA_0% > 0 Then
           ANCHO% = ANCHO_COLUMNA_0% 'SI  SE PASA EL ANCHO DE LA COLUMNA 0 COMO PARAMETRO TOMA ESTE
         Else
           ANCHO% = 3  'EN CASO CONTRARIO ASIGNA VALOR POR DEFECTO
         End If
         NUCOLU% = i%:   Form1.TEPRUEBA = String$(ANCHO%, "A"): MSF1.ColWidth(NUCOLU%) = Form1.TEPRUEBA.Width: MSF1.ColAlignment(NUCOLU%) = ALINEA%: MSF1.TextMatrix(0, NUCOLU%) = NOCAMPO$(i%)
         GoTo 50
       End If
       If i% = 1 Then
          i% = i% + ApartirdeColum%
       End If
30     LETRX$ = "A"
       ANCHO% = XVALO(Mid$(FORMX$(i%), 2))
       If XVALO(Mid$(FORMX$(i%), 2)) > 0 Then
          ANCHO% = ANCHO% + 2
       ElseIf LETRX$ = "Z" Then
            
       Else
          ANCHO% = 0 'COLUMNAS OCULTAS PASADAS CON F0.0
       End If
       LETR$ = UCase$(Left$(FORMX$(i%), 1))
       ALINEA% = 1
       If LETR$ = "I" Or LETR$ = "G" Or LETR$ = "F" Then
          ALINEA% = 6
       ElseIf LETR$ = "Z" Then
          ALINEA% = 3
       End If
       '
       If ANCHO% > 0 Then
          NUCOLU% = i%:   Form1.TEPRUEBA = String$(ANCHO%, "A"): MSF1.ColWidth(NUCOLU%) = Form1.TEPRUEBA.Width: MSF1.ColAlignment(NUCOLU%) = ALINEA%: MSF1.TextMatrix(0, NUCOLU%) = NOCAMPO$(i%)
       Else ' PARA QUE NO SE VEAN LAS COLUMNAS OCULTAS
          NUCOLU% = i%:   Form1.TEPRUEBA = String$(ANCHO%, "A"): MSF1.ColWidth(NUCOLU%) = 0: MSF1.ColAlignment(NUCOLU%) = ALINEA%: MSF1.TextMatrix(0, NUCOLU%) = NOCAMPO$(i%)
       End If
       MSF1.col = NUCOLU% - 1
       MSF1.CellFontBold = True
       
       ANCHO_TOTAL = ANCHO_TOTAL + Form1.TEPRUEBA.Width
       
       'A PARTIR DE AQUI CODIGO PARA CARGAR ENCABEZADO DE GENREP07

50     If UCase$(Form1.Name) = "SHOWREP07" Then

          Form1.Label5.FontName = Form1.GRID.CellFontName
          Form1.Label5.FontSize = Form1.GRID.CellFontSize
          Form1.Label5.FontBold = Form1.GRID.CellFontBold
          ' Label5.Caption = TEXTBIN$("1.234567", FORMAECO$(I&)) + "*"
          Form1.Label5.Caption = String$(ANCHO%, "A")    ' PROVISORIAMENTE A VER SI FUNCIONA
          Form1.GRID.col = i%
          Form1.GRID.ColWidth(JI& - 1) = SHOWREP07.Label5.Width
25        If i% >= Form1.Label8.Count Then
             NUPIC% = Form1.Label8.Count
             Load Form1.Label8(NUPIC%)
             GoTo 25
          End If
          Form1.Label8(i%) = Str$(Len(Form1.Label5) - 1)
          If i% = 0 Then Form1.Label8(i%) = 0
       End If
80  Next i%
    '
    If NUPIC% = i% - 1 Then
      If UCase$(Form1.Name) = "SHOWREP07" Then Load Form1.Label8(NUPIC% + 1): Load Form1.Label8(NUPIC% + 2)
    End If
    If NOREPIT% > 0 Then
       Call Duplicados1(Form1.CMBFILTRO)
     End If
End Sub
Sub CARGA_ENCABEZADOMSFH(MSF1 As MSHFlexGrid, LISTADECAMPOS$, Form1 As Form, Optional FiltroCombo%, Optional NOREPIT%, Optional ANCHO_TOTAL, Optional ANCHO_COLUMNA_0%, Optional PINTAFILACLICK%, Optional ApartirdeColum%)
    '
    
    Dim NOCAMPO$(), FORMX$()
    ReDim NOCAMPO$(1), FORMX$(1)
    On Error Resume Next
    Form1.CMBFILTRO.Clear
    Form1.CMBSERIE1.Clear
    Form1.CMBSERIE2.Clear
    Form1.CMBSERIE3.Clear
    Form1.CMBLABEL.Clear
    On Error GoTo 0
    '
    ANCHO_TOTAL = 0
    On Error Resume Next
    If Form1.Label8.Count > 1 Then
      Erase Form1.Label8
    End If
    
    If FiltroCombo% > 0 Then
       Form1.CMBFILTRO.AddItem ""
       Form1.CMBSERIE1.AddItem ""
       Form1.CMBSERIE2.AddItem ""
       Form1.CMBSERIE3.AddItem ""
       Form1.CMBLABEL.AddItem ""
    End If
    On Error GoTo 0
    NUPIC% = 0
    MODOSE$ = UCase$(SAFETEXT$(MODOSEL$))
    LSTX$ = TRIM$(LISTADECAMPOS$)
    CANCA% = 0 + ApartirdeColum%
    ATOTA% = 1
    While LSTX$ <> ""
       PPXX% = InStr(LSTX$, ";")
       PPXX2% = InStr(LSTX$, ",")
       If PPXX2% > 0 And PPXX2% < PPXX% Then PPXX% = PPXX2%
       If PPXX% < 1 Then PPXX% = 1 + Len(LSTX$)
       '
       CANCA% = CANCA% + 1
       If CANCA% >= UBound(NOCAMPO$) Then
          ReDim Preserve NOCAMPO$(CANCA%), FORMX$(CANCA%)
       End If
       NOCX$ = Left$(LSTX$, PPXX% - 1)
       LSTX$ = Mid$(LSTX$, PPXX% + 1)
       '
       PPXX% = InStr(NOCX$, "/")
       If PPXX% < 1 Then PPXX% = 1 + Len(NOCX$)
       NOCAMPO$(CANCA%) = Left$(NOCX$, PPXX% - 1)
       FORMX$(CANCA%) = Mid$(NOCX$, PPXX% + 1)
       '
       'AGREGO EL LOS CAMPOS AL FILTRO y para los graficos
       If FiltroCombo% > 0 Then
          If XVALO(Mid$(FORMX$(CANCA%), 2)) > 0 Then
             Form1.CMBFILTRO.AddItem NOCAMPO$(CANCA%)
             On Error Resume Next 'CODIGO AGREGADO PARA GENERAR GRAFICOS
             LETR1$ = UCase$(Left$(FORMX$(CANCA%), 1))
             If LETR1$ = "I" Or LETR1$ = "G" Or LETR1$ = "F" Then
                Call REPLA(TX$, (NOCAMPO$(CANCA%)), 1, 48): Call REPLA(TX$, TRIM$(CStr(CANCA%)), 50, 4): Call REPLA(TX$, "F", 56, 4)
                Form1.CMBSERIE1.AddItem TX$ 'SI ES NUMERICO CARGA PARA LAS SERIES
                Form1.CMBSERIE2.AddItem TX$
                Form1.CMBSERIE3.AddItem TX$
                Form1.CMBLABEL.AddItem TX$ 'en el eje lo pongo tambien, estoy parobando que esten todos en el eje
             Else
                Call REPLA(TX$, (NOCAMPO$(CANCA%)), 1, 48): Call REPLA(TX$, TRIM$(CStr(CANCA%)), 50, 4): Call REPLA(TX$, LETR1$, 56, 4)
                Form1.CMBLABEL.AddItem TX$ 'SI ES TEXTO VA AL COMOB DE LABEL
             End If
             On Error GoTo 0
          End If
       End If

       If FORMX$(CANCA%) = "" Then FORMX$(CANCA%) = "A16"
       '
       ATOTA% = ATOTA% + XVALO(Mid$(FORMX$(CANCA%), 2)) + 2
       '
'       TECO$ = ""
'       PPXX% = InStr(FORMX$(CANCA%), "+")
'       If PPXX% > 0 Then
'          TECO$ = UCase$(TRIM$(Mid$(FORMX$(CANCA%), PPXX% + 1)))
'          PPXX% = InStr(TECO$, "/")
'          If PPXX% > 0 Then
'             FORMATX$ = TRIM$(Mid$(TECO$, PPXX% + 1))
'             ATOTA% = ATOTA% + Len(TMASCARA$(FORMATX$))
'          End If
'       End If
       
       If CANCA% = 1 Then ACOL1% = XVALO(Mid$(FORMX$(CANCA%), 2))
       '
    Wend

    MSF1.Cols = CANCA% + 1
    MSF1.Rows = 1
    For i% = 0 To CANCA%
       'PRIMER COLUMNA ASIGNA POR DEFECTO ANCHO = 2

       If i% = 0 Then
         If UCase$(Form1.Name) = "SHOWREP07" Then GoTo 80

         If ANCHO_COLUMNA_0% > 0 Then
           ANCHO% = ANCHO_COLUMNA_0% 'SI  SE PASA EL ANCHO DE LA COLUMNA 0 COMO PARAMETRO TOMA ESTE
         Else
           ANCHO% = 3  'EN CASO CONTRARIO ASIGNA VALOR POR DEFECTO
         End If
         NUCOLU% = i%:   Form1.TEPRUEBA = String$(ANCHO%, "A"): MSF1.ColWidth(NUCOLU%) = Form1.TEPRUEBA.Width: MSF1.ColAlignment(NUCOLU%) = ALINEA%: MSF1.TextMatrix(0, NUCOLU%) = NOCAMPO$(i%)
         GoTo 50
       End If
       If i% = 1 Then
          i% = i% + ApartirdeColum%
       End If
30     LETRX$ = "A"
       ANCHO% = XVALO(Mid$(FORMX$(i%), 2))
       If XVALO(Mid$(FORMX$(i%), 2)) > 0 Then
          ANCHO% = ANCHO% + 2
       ElseIf LETRX$ = "Z" Then
            
       Else
          ANCHO% = 0 'COLUMNAS OCULTAS PASADAS CON F0.0
       End If
       LETR$ = UCase$(Left$(FORMX$(i%), 1))
       ALINEA% = 1
       If LETR$ = "I" Or LETR$ = "G" Or LETR$ = "F" Then
          ALINEA% = 6
       ElseIf LETR$ = "Z" Then
          ALINEA% = 3
       End If
       '
       If ANCHO% > 0 Then
          NUCOLU% = i%:   Form1.TEPRUEBA = String$(ANCHO%, "A"): MSF1.ColWidth(NUCOLU%) = Form1.TEPRUEBA.Width: MSF1.ColAlignment(NUCOLU%) = ALINEA%: MSF1.TextMatrix(0, NUCOLU%) = NOCAMPO$(i%)
       Else ' PARA QUE NO SE VEAN LAS COLUMNAS OCULTAS
          NUCOLU% = i%:   Form1.TEPRUEBA = String$(ANCHO%, "A"): MSF1.ColWidth(NUCOLU%) = 0: MSF1.ColAlignment(NUCOLU%) = ALINEA%: MSF1.TextMatrix(0, NUCOLU%) = NOCAMPO$(i%)
       End If
       MSF1.col = NUCOLU% - 1
       MSF1.CellFontBold = True
       
       ANCHO_TOTAL = ANCHO_TOTAL + Form1.TEPRUEBA.Width
       
       'A PARTIR DE AQUI CODIGO PARA CARGAR ENCABEZADO DE GENREP07

50     If UCase$(Form1.Name) = "SHOWREP07" Then

          Form1.Label5.FontName = Form1.GRID.CellFontName
          Form1.Label5.FontSize = Form1.GRID.CellFontSize
          Form1.Label5.FontBold = Form1.GRID.CellFontBold
          ' Label5.Caption = TEXTBIN$("1.234567", FORMAECO$(I&)) + "*"
          Form1.Label5.Caption = String$(ANCHO%, "A")    ' PROVISORIAMENTE A VER SI FUNCIONA
          Form1.GRID.col = i%
          Form1.GRID.ColWidth(JI& - 1) = SHOWREP07.Label5.Width
25        If i% >= Form1.Label8.Count Then
             NUPIC% = Form1.Label8.Count
             Load Form1.Label8(NUPIC%)
             GoTo 25
          End If
          Form1.Label8(i%) = Str$(Len(Form1.Label5) - 1)
          If i% = 0 Then Form1.Label8(i%) = 0
       End If
80  Next i%
    '
    If NUPIC% = i% - 1 Then
      If UCase$(Form1.Name) = "SHOWREP07" Then Load Form1.Label8(NUPIC% + 1): Load Form1.Label8(NUPIC% + 2)
    End If
    If NOREPIT% > 0 Then
       Call Duplicados1(Form1.CMBFILTRO)
     End If
End Sub

Sub CARGA_ENCABEZADO_MESES(MSF As msFlexGrid, Form1 As Form, FEDESA, FEHASA, Formato$, ApartirdeColum%, RET_CANTMESES%, Optional RET_NUMPRIMMES%, Optional RET_NUMULTIMOMES%, Optional NOMBREFILAS$)
    'ESTA FUNCION AGREGA COMO ENCABEZADO LOS MESES Y AÑOS SEGUN UN RANGO DE FECHA
    'ENE. 2016, FEB. 2016 ......................
    
    FEDES = FECHATEX$(FEDESA)
    FEHAS = FECHATEX$(FEHASA)
    CANTIMESES% = 1 + DateDiff("M", FEDES, FEHAS)  ' CANTIDAD DE MESES COMPRENIDOS EN EL RANGO
    If CANTIMESES% < 1 Then Call COMUNI("Error de Ingreso de Rango de Fechas")
    RET_CANTMESES% = CANTIMESES% 'RETORNO DE LA RUTINA
    '
    Dim MES$(12)
    MES$(1) = "ENERO"
    MES$(2) = "FEBRERO"
    MES$(3) = "MARZO"
    MES$(4) = "ABRIL"
    MES$(5) = "MAYO"
    MES$(6) = "JUNIO"
    MES$(7) = "JULIO"
    MES$(8) = "AGOSTO"
    MES$(9) = "SETIEMBRE"
    MES$(10) = "OCTUBRE"
    MES$(11) = "NOVIEMBRE"
    MES$(12) = "DICIEMBRE"
    
    MESDES% = Month(FEDES)
    MESHAS% = Month(FEHAS)
    ANO = Year(FEDES)
    ANOHAS = Year(FEHAS)
    J% = MESDES%
    ANCHO% = XVALO(Mid$(Formato$, 2))
    
    LETR$ = UCase$(Left$(Formato$, 1))
    ALINEA% = 1
    If LETR$ = "I" Or LETR$ = "G" Or LETR$ = "F" Then
          ALINEA% = 6
    ElseIf LETR$ = "Z" Then
          ALINEA% = 3
    End If

    For i& = 0 To CANTIMESES% - 1
       MSF.Cols = ApartirdeColum% + i& + 1
       NUCOLU% = i& + ApartirdeColum%:   Form1.TEPRUEBA = String$(ANCHO%, "A"): MSF.ColWidth(NUCOLU%) = Form1.TEPRUEBA.Width: MSF.ColAlignment(NUCOLU%) = ALINEA%
       APARTIR$ = ""
       If i& = 0 Then
          APARTIR$ = Left$(FECHATEX$(FEDESA), 2)
          APARTIR$ = "(" & APARTIR$ & ")"
       End If
       
       MSF.TextMatrix(0, NUCOLU%) = Left$(MES$(J%), 3) & ". " & ANO & APARTIR$
       If J% > MESHAS% And ANO >= ANOHAS Then GoTo 80    'CONTROLA SI LLEGA AL FINAL
       If J% = 12 Then J% = 0: ANO = ANO + 1
       J% = J% + 1
    Next i&
80  HASTAX$ = Left$(FECHATEX$(FEHASA), 2)
    HASTAX$ = "(" & HASTAX$ & ")"
    MSF.TextMatrix(0, NUCOLU%) = MSF.TextMatrix(0, NUCOLU%) & HASTAX$
    'AGREGO UNA COLUMNA MAS AL FINAL PARA GUARDAR EL ORDEN
    MSF.Cols = ApartirdeColum% + i& + 1
    MSF.ColWidth(ApartirdeColum% + i&) = 0 'LA HAGO INVISIBLE

End Sub

Sub CARGA_ENCABEZADO_SEMANA(MSF As msFlexGrid, Form1 As Form, FEDESA, FEHASA, Formato$, ApartirdeColum%, Optional RET_CANTISEM%, Optional RET_NUMPRIMSEMANA%, Optional RET_NUMULTIMSEMANA%, Optional NOMBREFILAS$)
    'ESTA FUNCION AGREGA COMO ENCABEZADO LAS SEMANAS JUNTO CON EL MES SEGUN UN RANGO DE FECHA
    '01/03 al 07/03 ......................
    
    FEDES = FECHATEX$(FEDESA)
    FEHAS = FECHATEX$(FEHASA)
    CANTISEM% = 1 + DateDiff("W", FEDES, FEHAS)  ' CANTIDAD DE MESES COMPRENIDOS EN EL RANGO
    If CANTISEM% < 1 Then Call COMUNI("Error de Ingreso de Rango de Fechas")
    RET_CANTISEM% = CANTISEM% 'RETORNO DE LA RUTINA
    '
    Dim MES$(12)
    MES$(1) = "ENERO"
    MES$(2) = "FEBRERO"
    MES$(3) = "MARZO"
    MES$(4) = "ABRIL"
    MES$(5) = "MAYO"
    MES$(6) = "JUNIO"
    MES$(7) = "JULIO"
    MES$(8) = "AGOSTO"
    MES$(9) = "SETIEMBRE"
    MES$(10) = "OCTUBRE"
    MES$(11) = "NOVIEMBRE"
    MES$(12) = "DICIEMBRE"
    
    SEMDES% = DatePart("ww", FEDES)
    SEMHAS% = DatePart("ww", FEHAS)
    
    ANO = Year(FEDES)
    ANOHAS = Year(FEHAS)
    
    J% = MESDES%
    ANCHO% = XVALO(Mid$(Formato$, 2))
'    Function RangoSemanaSegunFecha$(FechaInicio%, RET_ULTIMODIASEMANA, RET_NUMSEMANA, Optional INICIADIAHABIL%, Optional FECHAFINRANGO%)

    LETR$ = UCase$(Left$(Formato$, 1))
    ALINEA% = 1
    If LETR$ = "I" Or LETR$ = "G" Or LETR$ = "F" Then
          ALINEA% = 6
    ElseIf LETR$ = "Z" Then
          ALINEA% = 3
    End If
    i& = 0
    Des% = FEDESA
    RET_FEFIN = FEHASA
    While FEHASA > Des%
       MSF.Cols = ApartirdeColum% + i& + 1
       NUCOLU% = i& + ApartirdeColum%:   Form1.TEPRUEBA = String$(ANCHO%, "A"): MSF.ColWidth(NUCOLU%) = Form1.TEPRUEBA.Width: MSF.ColAlignment(NUCOLU%) = ALINEA%
       MSF.TextMatrix(0, NUCOLU%) = RangoSemanaSegunFecha$(Des%, RET1, RET2, 1, Int(FEHASA))
       Des% = RET1
       i& = i& + 1
   Wend
    MSF.Cols = ApartirdeColum% + i& + 1
    MSF.ColWidth(ApartirdeColum% + i&) = 0 'LA HAGO INVISIBLE

End Sub

Sub CARGA_ENCABEZADO_DIAS(MSF As msFlexGrid, Form1 As Form, FEDESA, FEHASA, Formato$, ApartirdeColum%, RET_CANTDIAS%, Optional RET_NUMPRIMDIA%, Optional RET_NUMULTIMODIA%, Optional NOMBREFILAS$)
    'ESTA FUNCION AGREGA COMO ENCABEZADO LOS MESES Y AÑOS SEGUN UN RANGO DE FECHA
    'ENE. 2016, FEB. 2016 ......................
    
    FEDES = FECHATEX$(FEDESA)
    FEHAS = FECHATEX$(FEHASA)
    CANTIMESES% = 1 + DateDiff("M", FEDES, FEHAS)  ' CANTIDAD DE MESES COMPRENIDOS EN EL RANGO
    If CANTIMESES% < 1 Then Call COMUNI("Error de Ingreso de Rango de Fechas")
    
    CantiDias% = 1 + DateDiff("D", FEDES, FEHAS)  ' CANTIDAD DE DIAS COMPRENIDOS EN EL RANGO
    If CantiDias% < 1 Then Call COMUNI("Error de Ingreso de Rango de Fechas")
    RET_CANTDIAS% = CantiDias% 'RETORNO DE LA RUTINA
    
    '
    Dim MES$(12)
    MES$(1) = "ENERO"
    MES$(2) = "FEBRERO"
    MES$(3) = "MARZO"
    MES$(4) = "ABRIL"
    MES$(5) = "MAYO"
    MES$(6) = "JUNIO"
    MES$(7) = "JULIO"
    MES$(8) = "AGOSTO"
    MES$(9) = "SETIEMBRE"
    MES$(10) = "OCTUBRE"
    MES$(11) = "NOVIEMBRE"
    MES$(12) = "DICIEMBRE"
    
    MESDES% = Month(FEDES)
    MESHAS% = Month(FEHAS)
    ANO = Year(FEDES)
    ANOHAS = Year(FEHAS)
    
    J% = 0
    ANCHO% = XVALO(Mid$(Formato$, 2))
    
    LETR$ = UCase$(Left$(Formato$, 1))
    ALINEA% = 1
    If LETR$ = "I" Or LETR$ = "G" Or LETR$ = "F" Then
          ALINEA% = 6
    ElseIf LETR$ = "Z" Then
          ALINEA% = 3
    End If
    FEXX% = FEDESA
    For i& = 0 To CantiDias% - 1
       If esFinSemana(FEXX%) = False Then
         MSF.Cols = ApartirdeColum% + J% + 1
         NUCOLU% = J% + ApartirdeColum%:   Form1.TEPRUEBA = String$(ANCHO%, "A"): MSF.ColWidth(NUCOLU%) = Form1.TEPRUEBA.Width: MSF.ColAlignment(NUCOLU%) = ALINEA%
         'MSF.TextMatrix(0, NUCOLU%) = Left$(FECHATEX$(FEXX%), 2) ' Left$(MES$(J%), 3)
         MSF.TextMatrix(0, NUCOLU%) = FECHATEX$(FEXX%) ' Left$(MES$(J%), 3)

'         If J% > MESHAS% And ANO >= ANOHAS Then Exit For    'CONTROLA SI LLEGA AL FINAL
'         If J% = 12 Then J% = 0: ANO = ANO + 1
         J% = J% + 1
         
         
       End If
       FEXX% = DIASPOST(FEXX%, 1)
    Next i&
    MSF.Cols = ApartirdeColum% + J% + 1
    MSF.ColWidth(ApartirdeColum% + J%) = 0 'LA HAGO INVISIBLE
    

End Sub


 Function Duplicados1(COM As comboBox)
    Dim i As Integer, X As Integer
    i% = 0
    While i% < COM.ListCount - 1
       i% = i% + 1
       X = i%
       While X < COM.ListCount - 1
        If TRIM$((COM.List(i%))) = TRIM$(COM.List(X)) And X <> i% Then
           COM.RemoveItem (X)
           X = X - 1
        End If
        X = X + 1
       Wend
   Wend
End Function

Sub AgregarSinRepetirValorComboBox(ByVal comboBox As comboBox, ByVal newItem As String)
    Dim i As Integer
    Dim itemExists As Boolean
    
    ' Verificar si el nuevo elemento ya existe en el ComboBox
    For i = 0 To comboBox.ListCount - 1
        If comboBox.List(i) = newItem Then
            itemExists = True
            Exit For
        End If
    Next i
    
    ' Agregar el nuevo elemento al ComboBox si no existe
    If Not itemExists Then
        comboBox.AddItem newItem
    End If

End Sub
 Function DuplicadosList1(COM As ListBox)
    Dim i As Integer, X As Integer
    While i% < COM.ListCount - 1
       i% = i% + 1
       While X < COM.ListCount - 1
        X = X + 1
        If (COM.List(i%) = COM.List(X)) And X <> i% Then
           COM.RemoveItem (i%)
        End If
       Wend
   Wend
End Function

Sub SELECCIONARFILA(MSF As msFlexGrid, REG&)

   MSF.row = REG&: MSF.col = 0
   MSF.RowSel = REG&: MSF.ColSel = MSF.Cols - 1

End Sub
   
Sub INSERTYDESPLAZAR(Fila&, MSF As msFlexGrid, DIRECCION%)
    'DESPLAZA TODA LA FILA HACIA ARRIBA HACIA ABAJO DE UNA POSICION
    'DIRECCION -1 ARRIBA      1 ABAJO
    'EJEMPLO DE LLAMADA: 'REG& = GRID.RowSel
                         'Call INSERTYDESPLAZAR(REG&, GRID, 1)

    If DIRECCION% > 0 Then
       If Fila& >= MSF.Rows - 1 Then Exit Sub
    Else
       If Fila& <= 1 Then Exit Sub
    End If
    UR& = MSF.Rows - 1
    For i& = 1 To MSF.Cols - 1
         DATOAMOVER$ = MSF.TextMatrix(Fila&, i&)
         DATOPISADO$ = MSF.TextMatrix(Fila& + DIRECCION%, i&)
         MSF.TextMatrix(Fila& + DIRECCION%, i&) = DATOAMOVER$
         MSF.TextMatrix(Fila&, i&) = DATOPISADO$
         Call COLOREAR_GRILLA_SOLAFILA(MSF, vbWhite, Fila&, 1, 1)
    Next i&
    Call UbicarFocoTop(MSF, Fila& + DIRECCION%)

End Sub


Sub UbicarFocoTop(MSF As msFlexGrid, REG&)
      'SE POSICIONA SEGUN EL REGISTRO PASADO PARA QUEDAR A LA VISTA EN UNA GRILLA CON MUCHOS DATOS.
      'SE USA EN LA SELECCION CON FLECHA HACIA ARRIBA Y ABAJO
      MSF.row = REG&
      If MSF.TopRow > MSF.row Then MSF.TopRow = MSF.row
      On Error Resume Next
      VISIROWS = (MSF.Height / MSF.RowHeight(REG&)) - 4
      If Err > 0 Then
         VISIROWS = 1
      End If
      On Error GoTo 0
      If VISIROWS > 0 Then
            If MSF.TopRow < MSF.row - VISIROWS Then
              MSF.TopRow = MSF.row - VISIROWS
            End If
      End If
      Call SELECCIONARFILA(MSF, REG&)

End Sub
Sub ORDENAR_MSF(MSF As msFlexGrid, COLUMNA_A_ORDENAR, Optional MODO_ASC_DESC$, Optional ESFECHA%)
       'ASCENDENTE = 1 (ASC)
       'DESCENDENTE = 2 (DESC)
       If UCase$(MODO_ASC_DESC$) = "DESC" Then
         Ordenamiento% = 2
       Else
         Ordenamiento% = 1
       End If
       '
       'SI ES CAMPO FECHA LO PASA A NUMERO A UNA NUEVA COLUMNA TEMPORAL A MODO NUMERICO, ORDENA Y LUEGO BORRA LA COLUMNA TEMPORAL
       If ESFECHA% > 0 Then
         MSF.Cols = MSF.Cols + 1
         COLUTEMP = MSF.Cols - 1
         For i& = 1 To MSF.Rows - 1
            MSF.TextMatrix(i&, COLUTEMP) = FECHANUM(MSF.TextMatrix(i&, COLUMNA_A_ORDENAR))
         Next i&
         MSF.col = COLUTEMP
         MSF.Sort = Ordenamiento%

         MSF.Cols = MSF.Cols - 1 'BORRA LA COLUMNA TEMPORAL
         Exit Sub 'SI ES FECHA TERMINA ACA
       End If
       
       MSF.col = COLUMNA_A_ORDENAR
       MSF.Sort = Ordenamiento%
        
End Sub
Sub COLOREAR_GRILLA_SOLACELDA(MSF As msFlexGrid, COLOR As Variant, REGISTRO, COLUMNA%, Optional TEXTX%)
   'RECORRE LA FILA SELECCIONADA Y PINTA CADA CELDA
   'SI LA CELDA ES DIFERENTE AL COLOR PASADO COMO PARAMETRO LA PINTA , SI NO LA PONE EN BLANCO
   If REGISTRO < 1 Then Exit Sub
   VALIDACOLUM% = 0
   On Error Resume Next
   With MSF
    If XVALO(REGISTRO) > 0 Then
      iRow% = XVALO(REGISTRO)
    Else
      iRow% = MSF.row
    End If
    
    .row = iRow%                ' ASIGNA EL NUMERO DE REGISTRO POSICIONADO
    If COLUMNA% > 0 Then
        .col = COLUMNA%
       .CellBackColor = COLOR
       Exit Sub
    End If
    For iCol% = 0 To .Cols - 1 ' DENTRO DEL RENGLON RECORRE CADA COLUMNA
        .col = iCol%
        If iCol% = COLUMNA% Then
          If TEXTX% > 0 Then
            .CellForeColor = COLOR
          Else
           .CellBackColor = COLOR
          End If
          Exit For
        End If
    Next iCol%
   End With
   On Error GoTo 0
End Sub
Sub COLOREAR_FUENTE_SOLO_UNA_CELDA(MSF As msFlexGrid, COLOR As Variant, REG&, col%)
   MSF.row = REG&
   MSF.col = col%
   MSF.CellBackColor = COLOR
   
End Sub
Sub NEGRITA_SOLO_UNA_CELDA(MSF As msFlexGrid, REG&, col%)
   MSF.row = REG&
   MSF.col = col%
   MSF.CellFontBold = True

End Sub
Sub NEGRITA_SOLO_UNA_FILA(MSF As msFlexGrid, REG&, Optional TamFuente%)
   MSF.row = REG&
   TamañoFuenteOriginal = MSF.Font.Size
   For i% = 1 To MSF.Cols - 1
     MSF.col = i%
     MSF.CellFontBold = True
     If TamFuente% > 0 Then
        MSF.Font.Size = TamFuente%
     End If
   Next i%
   MSF.Font.Size = TamañoFuenteOriginal
End Sub

Sub numerargrilla(MSF As msFlexGrid, col&, Optional DESCENDENTE%, Optional A_PARTIRDELULTIMO%, Optional DESDECERO%)
    If A_PARTIRDELULTIMO% > 0 Then
       'SI A_PARTIRDELULTIMO% ES > 0 VA A MODIFICAR SOLO LAS CELDAS QUE ESTEN SIN VALOR REPITIENDO EL INMEDIATO
       'ANTERIOR LEIDO O SEA QUE SI HAY 5 FILAS Y DICHA COLUMNA ESTA VACIA VAN A QUEDAR TODAS NUMERADAS EN 1
       'SI EN CAMBIO EN LA TERCER FILA SE CAMBIA A 2 QUEDARAN LAS 2 PRIMERAS EN 1 Y LAS 3 SIGUIENTES EN 2
       For i& = 1 To MSF.Rows - 1
         ValLeido& = XVALO(MSF.TextMatrix(i&, col&))
         If ValLeido& < 1 Then
            VALLEIDO_ANT& = VALLEIDO_ANT&: If VALLEIDO_ANT& = 0 Then VALLEIDO_ANT& = 1
            
            MSF.TextMatrix(i&, col&) = VALLEIDO_ANT&
'            VALLEIDO& = VALLEIDO_ANT&
         End If
         VALLEIDO_ANT& = ValLeido&
       Next i&
       Exit Sub
    End If
    '
    DESCUENTO% = 0
    If DESDECERO% > 0 Then DESCUENTO% = 1
    If DESCENDENTE% > 0 Then
       For i& = MSF.Rows - 1 To 1 Step -1
          J& = J& + 1
          MSF.TextMatrix(J&, col&) = i&
       Next i&
    Else
        For i& = 1 To MSF.Rows - 1
           MSF.TextMatrix(i&, col&) = i& - DESCUENTO%
        Next i&
     End If
End Sub
Function SUMA_COLUMNA_SELECCION(MSF As msFlexGrid, col&, Optional MARCA$)
   SUMA_COLUMNA_SELECCION = 0
   For i& = 1 To MSF.Rows - 1
     If MARCA$ <> "" Then
        If MSF.TextMatrix(i&, 0) = MARCA$ Then
          SUMA_COLUMNA_SELECCION = SUMA_COLUMNA_SELECCION + XVALO(MSF.TextMatrix(i&, col&))
        End If
     Else
          SUMA_COLUMNA_SELECCION = SUMA_COLUMNA_SELECCION + XVALO(MSF.TextMatrix(i&, col&))
     End If
   Next i&
   SUMA_COLUMNA_SELECCION = XVALO(SUMA_COLUMNA_SELECCION)
End Function
Function TOTALFILAS&(MSF As msFlexGrid)
   TOTALFILAS& = MSF.Rows - 1
End Function
   Sub CopiarMsfAOtroMsf(MSF As msFlexGrid, MSF2 As msFlexGrid, Optional A_PARTIRDE%, Optional MSFWidth, Optional FRM As Form)
      'SELECCIONA .COL Y ROW ES EL DESDE
      'COLSEL Y ROWSEL ES EL HASTA
      
      If XVALO(MSFWidth) > 0 Then
        MSF2.Width = MSFWidth
      End If
      
      With MSF
         .col = 1
         .row = 1
         .ColSel = .Cols - 1
         .RowSel = .Rows - 1
         On Error Resume Next
         .SetFocus
         On Error GoTo 0
      End With
        
      With MSF2
          
        .Rows = 1
          
        ' si el rango es de una sola columna sale
        If MSF.ColSel = MSF.col Then
            .Rows = 1
            .Cols = MSF.Cols
            Exit Sub
        End If
          
        'Asigna la cantidad de filas seleccionadas del rango
        .Rows = Abs(MSF.RowSel - MSF.row) + 2
        .Cols = MSF.Cols
        .FixedCols = MSF.FixedCols
        .FixedRows = MSF.FixedRows
                  
        .row = 1
        .col = A_PARTIRDE%
        .RowSel = MSF.Rows - 1
        .ColSel = MSF.Cols - 1
        ' copia los datos al otro flexgrid
        .Clip = MSF.Clip
    End With
End Sub

Sub INSERTARFILA(MSF As msFlexGrid, Fila&, Form1 As Form, Campos$)
   '
   If Fila& < 1 Then Exit Sub
   
   Set FRMINSERTAR.FORMULARIOX = Form1
   Set FRMINSERTAR.MSF = MSF
   '
   Call CARGA_ENCABEZADO(FRMINSERTAR.MSF1, Campos$, FRMINSERTAR)
   '
   'CARGA LOS DATOS DEL REGISTRO SELECCIONADO
   For h& = 1 To MSF.Cols - 1
      DATO1$ = MSF.TextMatrix(Fila&, h&)
      FRMINSERTAR.MSF1.Rows = 2
      FRMINSERTAR.MSF1.TextMatrix(1, h&) = TRIM$(DATO1$)
   Next h&
   FRMINSERTAR.MSF1.TextMatrix(1, 0) = TRIM$(Fila&)
   FRMINSERTAR.Top = Form1.Top + (Form1.Height - FRMINSERTAR.Height)
   FRMINSERTAR.Left = Form1.Left
   '
   FRMINSERTAR.Show 1
End Sub

Sub MODIFICARMSF(MSF As msFlexGrid, Fila&, Form1 As Form, Campos$)
   '
   If Fila& < 1 Then Exit Sub
   
   Set FRMMODEDI.FORMULARIOX = Form1
   Set FRMMODEDI.MSF = MSF
   '
   'Call CARGA_ENCABEZADO(FRMMODEDI.MSF1, CAMPOS$, FRMMODEDI)
   '
   'CARGA LOS DATOS DEL REGISTRO SELECCIONADO
   For h& = 1 To MSF.Cols - 1
      DATO1$ = MSF.TextMatrix(Fila&, h&)
      FRMMODEDI.MSF1.Rows = 2
      FRMMODEDI.MSF1.TextMatrix(1, h&) = TRIM$(DATO1$)
   Next h&
   FRMMODEDI.MSF1.TextMatrix(1, 0) = TRIM$(Fila&)
   FRMMODEDI.Top = Form1.Top + (Form1.Height - FRMMODEDI.Height)
   FRMMODEDI.Left = Form1.Left
   FRMMODEDI.Width = Form1.Width
   FRMMODEDI.Command1.Left = Form1.Width - ((FRMMODEDI.Command1.Width * 2) + 1000)
   FRMMODEDI.cmdcancelar.Left = Form1.Width - ((FRMMODEDI.cmdcancelar.Width) + 850)
   '
   FRMMODEDI.Show 1
End Sub

Function SUMACOLUMNA#(MSF As msFlexGrid, COLUMNA&, APARTIRDEFILA&)
   SUMACOLUMNA# = 0
   For J& = APARTIRDEFILA& To MSF.Rows - 1
     SUMACOLUMNA# = SUMACOLUMNA# + XVALO(MSF.TextMatrix(J&, COLUMNA&))
   Next J&
   
End Function
Function COLUMNA%(CAMPO_ENCA$, MSF As msFlexGrid)

   COLUMNA% = -1
   For i% = 1 To MSF.Cols - 1
     CAMPO_LEIDO$ = TRIM$(MSF.TextMatrix(0, i%))
     If CAMPO_LEIDO$ = CAMPO_ENCA$ Then
        COLUMNA% = i%
        Exit For
     End If
   Next i%
   
End Function
Sub ColorearRangoCeldasExcel(Rango$, COLOR%, OExcel As Object, XLB As Object, XLSH As Object, COLORESCLAROS%)
   'SOLO PONE LOS COLORES MAS CLAROS
   'VALIDACION PARA EL CASO EN EL QUE EL COLOR SE PASE P.EJ. EN UN FOR Y LA VARIABLE SEA INCEREMENTAL CUANDOSE PASA COMIENZA DE 1
10 If COLOR% > 48 Then
     COLOR% = COLOR% - 48
     If COLOR% > 48 Then
        COLOR% = COLOR% - 48
        GoTo 10
     End If
   End If
   '
   If COLORESCLAROS% > 0 Then
    Select Case COLOR%
       Case Is < 2
         COLOR% = 2
       Case Is < 8
         COLOR% = 8
       Case Is < 15
         COLOR% = 15
       Case Is < 19
         COLOR% = 19
       Case Is < 24
         COLOR% = 24
       Case Is < 34
         COLOR% = 34
       Case Is < 36
         COLOR% = 35
       Case Is < 37
         COLOR% = 36
       Case Is < 40
         COLOR% = 40
       Case Is < 44
         COLOR% = 44
       Case Is < 46
         COLOR% = 45
       Case Is < 56
         COLOR% = 48
         
       Case Else
          COLOR% = 8
    End Select
   End If
   'PALETA DE COLORES: http://dmcritchie.mvps.org/excel/colors.htm
   If Val(OExcel.Application.Version) >= 8 Then
      Set OExcel = XLB.ActiveSheet
      OExcel.Range(Rango$).Interior.ColorIndex = COLOR%
   Else
       XLSH.Range(Rango$).Select
       XLSH.Range(Rango$).Interior.ColorIndex = COLOR%
   End If

End Sub

Sub GeneraExcelDeGrilla(NombreDelArchivo As String, Grilla As Object, TITULO As String, Optional PrimerColumnaSeleccion As Boolean, Optional FormatoColumnas As String, Optional MensajeFin As Boolean = True)
   
    Dim XLAPP As Excel.Application
    Dim XLBOOK As Excel.Workbook
    Dim XLSHEET As Excel.Worksheet
    '
    On Error GoTo ErrInstanciaExcel
    '
    Dim CantDecimales%(64)
    Set XLAPP = New Excel.Application
    Set XLBOOK = XLAPP.Workbooks.Add
    Set XLSHEET = XLBOOK.Worksheets.Add
    '
    Dim Fila, COLUMNA, ColumnasVisibles As Long
    Dim DatoCelda As String
    Dim Encabezado As Integer
    Dim MensajeError As String

    If Grilla.Rows < 1 Then Exit Sub
    
    Screen.MousePointer = 11
    
    EncabezadoFila = 0: TITULO = TRIM$(TITULO)
    If TITULO <> "" Then
        XLSHEET.Cells(1, 1) = TITULO
        With XLSHEET.Cells(1, 1)
            .Font.Name = "Arial"
            .Font.Size = 14
            .Font.Bold = True
            .Font.Underline = True
            .Font.Italic = True
        End With
        XLSHEET.Range("A1:Z1").Merge
        XLSHEET.Range("A1:Z1").HorizontalAlignment = xlLeft
        Encabezado = 2
    End If
    'CARGA EN UN VECTOR LA ALINEACION QUE TIENE LA COLUMNA
    'VALIDA CON ESTE DATO SI UN DATO NUMERICO SOLO SE TRATA COMO TAL SI ESTA ALINEADO A DERECHA, CASO
    'CONTRARIO LO TOMA COMO TEXTO.(CASO AHP CODIGO EXT. TIENE FORMATO NUMERICO CON 0 ADELANTE Y LOS MISMOS
    'LOS BORRABA EJ. 000100 LO MOSTRABA EN EL EXCEL COMO 100
    Dim FORMATOCOL()
    ReDim Preserve FORMATOCOL(Grilla.Cols - 1)
    For i& = 0 To Grilla.Cols - 1
       Grilla.col = i&
       FORMATOCOL(i&) = Grilla.ColAlignment(i&)
    Next i&
    
    Call TEXTOLOTES$(FormatoColumnas, ";")
    For i = 0 To UBound(CADENATEX)
        LEIDOX$ = UCase(Mid$(Mid$(CADENATEX(i), InStr(1, CADENATEX(i), "/") + 1), 1))
        CantDecimales%(i) = 0
        If Left$(LEIDOX$, 1) = "F" And InStr(1, LEIDOX$, ".") Then
             CantDecimales%(i) = XVALO(Mid$(Mid$(LEIDOX$, InStr(1, LEIDOX$, ".") + 1), 1))
        End If
        CADENATEX(i) = UCase(Mid$(Mid$(CADENATEX(i), InStr(1, CADENATEX(i), "/") + 1), 1, 1))
        
    Next i
    '
    For Fila = 0 To Grilla.Rows - 1
       ColumnasVisibles = 0
       For COLUMNA = 0 To Grilla.Cols - 1
            CantDecimal% = XVALO(CantDecimales%(COLUMNA))
            If Grilla.ColWidth(COLUMNA) = 0 Then GoTo 10 ' OBVIA LAS COLUMNAS OCULTAS
            ColumnasVisibles = ColumnasVisibles + 1
            '
            DatoCelda = TRIM$(Grilla.TextMatrix(Fila, COLUMNA))
            '
            If COLUMNA = 0 And PrimerColumnaSeleccion = True Then
                If DatoCelda = "*" Then GoTo 10 ' OBVIA LA COLUMNA CON EL ASTERISCO
                If DatoCelda = "" Then GoTo 15 ' OBVIA LA FILA. NO FUE SELECCIONADA.
            End If
            
            Grilla.row = Fila: Grilla.col = COLUMNA
            FuenteNombre = Grilla.CellFontName
            Tamanio = Grilla.CellFontSize
            Negrita = Grilla.CellFontBold
            Italica = Grilla.CellFontItalic
            Subrayado = Grilla.CellFontUnderline
            Alinear = Grilla.ColAlignment(COLUMNA)
            ColorFuente = XVALO(Grilla.CellForeColor)
            If ColorFuente < 0 Then ColorFuente = 0      ' NEGRO POR DEFAULT
            If ColorFuente > 2147483647 Then ColorFuente = 0
            ColorFondo = XVALO(Grilla.CellBackColor)
            If ColorFondo < 0 Then ColorFondo = 16777215 ' BLANCO POR DEFAULT
            If ColorFondo > 2147483647 Then ColorFondo = 16777215
            '
            If IsDate(DatoCelda) And Len(DatoCelda) > 6 Then
              If Mid$(DatoCelda, 3, 1) = "/" And Mid$(DatoCelda, 6, 1) = "/" Then
                 DatoCelda = Format$(DatoCelda, "yyyy-mm-dd")
              End If
            End If
            '
            With XLSHEET.Cells(Encabezado + Fila + 1, ColumnasVisibles)
                .Font.Name = FuenteNombre
                .Font.Size = Tamanio
                .Font.Bold = Negrita
                .Font.Underline = Subrayado
                .Font.Italic = Italica
                    Call LongToRGB(ColorFuente, rd%, gr%, BL%)
                .Font.COLOR = RGB(rd%, gr%, BL%)
                    rd% = 255: gr% = 255: BL% = 255
                    If ColorFondo <> 0 Then
                        Call LongToRGB(ColorFondo, rd%, gr%, BL%)
                    End If
                If Not (rd% = 255 And gr% = 255 And BL% = 255) Then
                    .Interior.COLOR = RGB(rd%, gr%, BL%)
                End If
                
                If IsNumeric(DatoCelda) And (Alinear = 6 Or Alinear = 7 Or Alinear = 8) Then 'INGRESA AQUI SI ES NUMERO Y EN LA GRILLA ESTA ALINEADO A DERECHA
                    Formato$ = ""
                    If UBound(CADENATEX) > 0 Then Formato$ = CADENATEX(COLUMNA)
                    ' LA EXISTENCIA DEL PUNTO Y DE LA COMA SIGNIFICA QUE TIENE SEPARADOR DECIMAL Y DE MILES
                    If (PosicionComa% > 0 And PosicionPunto% > 0) Or Formato$ = "G" Then
                        DatoCelda = REPLACAR$(DatoCelda, ".", "") ' QUITO EL SEPARADOR DE MILES
                        DatoCelda = REPLACAR$(DatoCelda, ",", ".") ' REEMPLAZA LA COMA POR EL PUNTO PARA LA PARTE DECIMAL
                        PosicionPunto% = InStr(1, DatoCelda, ".")
                    ElseIf PosicionComa% > 0 And Val(DatoCelda$) < 1000 Then ' VALORES MENORES A 1000 QUE NO TIENEN SEPARADOR DE MILES
                        DatoCelda = REPLACAR$(DatoCelda, ",", ".") ' REEMPLAZA LA COMA POR EL PUNTO PARA LA PARTE DECIMAL
                        PosicionPunto% = InStr(1, DatoCelda, ".")
                    End If
                    If PosicionPunto% > 0 Then
                        LongitudDecimal% = Len((Mid$(DatoCelda, PosicionPunto% + 1)))
                        Do While Len(FormatoDecimal$) < LongitudDecimal%
                            FormatoDecimal$ = FormatoDecimal$ & "0"
                        Loop
                    End If
                    'agregado por que el codigo anterior no funcionaba correctamente si se concatenaba con variables
                    
                    If CantDecimal% = 1 Then .NumberFormat = "#,##0.0"
                    If CantDecimal% = 2 Then .NumberFormat = "#,##0.00"
                    If CantDecimal% = 3 Then .NumberFormat = "#,##0.000"
                    If CantDecimal% >= 4 Then .NumberFormat = "#,##0.0000"
                    
                    If LongitudDecimal% > 0 Then .NumberFormat = "#,##0." & FormatoDecimal$
'                    If CantDecimales%(ColumnasVisibles) > 0 Then XLSHEET.Cells(Encabezado + Fila + 1, ColumnasVisibles).NumberFormat = "#,##0." & FORMATOCON0$(0, CantDecimales%(ColumnasVisibles))
'                    XXXX$ = "#.##0," & FORMATOCON0$(0, CantDecimales%(ColumnasVisibles))
                    'DatoCelda = REPLACAR$(DatoCelda, ".", "")
                    'DatoCelda = REPLACAR$(DatoCelda, ",", ".")
                    XLSHEET.Cells(Encabezado + Fila + 1, ColumnasVisibles) = DatoCelda
                ElseIf IsDate(DatoCelda) And ((Mid$(DatoCelda, 5, 1) = "-" And Mid$(DatoCelda, 8, 1) = "-") Or (Mid$(DatoCelda, 5, 1) = "/" And Mid$(DatoCelda, 8, 1) = "/")) Then
'                    .NumberFormat = "d/m/yyyy;@" 'esto cambie para envaplast 130824
'                    XLSHEET.Cells(Encabezado + Fila + 1, ColumnasVisibles) = DatoCelda
                     XLSHEET.Cells(Encabezado + Fila + 1, ColumnasVisibles).NumberFormat = "dd/mm/yyyy"
                     XLSHEET.Cells(Encabezado + Fila + 1, ColumnasVisibles).Value = CDate(DatoCelda)
                    
                ElseIf InStr(1, DatoCelda, "$") > 0 Then
                    .NumberFormat = "$#,##0.00"
                    XLSHEET.Cells(Encabezado + Fila + 1, ColumnasVisibles) = DatoCelda
                
                Else
                    .NumberFormat = "@"
                    XLSHEET.Cells(Encabezado + Fila + 1, ColumnasVisibles) = DatoCelda
                
                End If
                
                Select Case Alinear
                    Case 1: .HorizontalAlignment = xlLeft
                    Case 4: .HorizontalAlignment = xlCenter
                    Case 6, 7, 8: .HorizontalAlignment = xlRight:
                    Case Else: .HorizontalAlignment = xlLeft
                End Select
                .VerticalAlignment = xlCenter
                
                .Borders.LineStyle = xlContinuous
            
            End With
10     Next COLUMNA
15  Next Fila



'   LetraColumna$ = ExcelConvierteColumnaNumeroALetra(Grilla.Cols)
'   XLSHEET.Range("A" & CStr(Encabezado + 1) & ":" & LetraColumna$ & CStr(Encabezado + 1)).Borders.LineStyle = xlContinuous
'   XLSHEET.Range("A" & CStr(Encabezado + 1) & ":" & LetraColumna$ & CStr(Encabezado + 1)).Borders.Weight = xlMedium
'
'   XLSHEET.Range(LetraColumna$ & CStr(Encabezado + 1) & ":" & LetraColumna$ & CStr(Encabezado + Grilla.Rows)).Borders.LineStyle = xlContinuous
'   XLSHEET.Range(LetraColumna$ & CStr(Encabezado + 1) & ":" & LetraColumna$ & CStr(Encabezado + Grilla.Rows)).Borders.Weight = xlMedium
'
'   XLSHEET.Range("A" & CStr(Encabezado + Grilla.Rows) & ":" & LetraColumna$ & CStr(Encabezado + Grilla.Rows)).Borders.LineStyle = xlContinuous
'   XLSHEET.Range("A" & CStr(Encabezado + Grilla.Rows) & ":" & LetraColumna$ & CStr(Encabezado + Grilla.Rows)).Borders.Weight = xlMedium
'
'   XLSHEET.Range("A" & CStr(Encabezado + 1) & ":A" & CStr(Encabezado + Grilla.Rows)).Borders.LineStyle = xlContinuous
'   XLSHEET.Range("A" & CStr(Encabezado + 1) & ":A" & CStr(Encabezado + Grilla.Rows)).Borders.Weight = xlMedium
  
   XLSHEET.Columns.AutoFit
   '
   Screen.MousePointer = 1
   '
   XLSHEET.SaveAs NombreDelArchivo
   On Error GoTo 0
   '
   XLAPP.Quit      ' cierra EXCEL que se esta ejecutando minimizado
   Set XLAPP = Nothing     ') libera memoria
   Set XLBOOK = Nothing    ') utilizada por
   Set XLSHEET = Nothing   ') las variables objeto
   
   Screen.MousePointer = 1
   If MensajeFin = True Then Call COMUNI("Planilla Excel Creada Exitosamente:\ \" + REPLACAR$(NombreDelArchivo, "\", "/"))
   '
   
99: Exit Sub
   
ErrInstanciaExcel:
   '
   MensajeError = TRIM$(Err.Description)
   On Error GoTo 0
   XLAPP.Quit      ' cierra EXCEL que se esta ejecutando minimizado
   Set XLAPP = Nothing     ') libera memoria
   Set XLBOOK = Nothing    ') utilizada por
   Set XLSHEET = Nothing   ') las variables objeto
   
   Screen.MousePointer = 1
   Call MENSERR(24, "Imposible Crear Planilla Excel.\ \" & MensajeError)
   Exit Sub
   '
End Sub

Sub GeneraExcelDeGrillaTabulada(NombreDelArchivo As String, Grilla As Object, TITULO As String, Optional PrimerColumnaSeleccion As Boolean, Optional FormTraza As Form)
   
    Dim XLAPP As Excel.Application
    Dim XLBOOK As Excel.Workbook
    Dim XLSHEET As Excel.Worksheet
    '
    On Error GoTo ErrInstanciaExcel
    '
    Set XLAPP = New Excel.Application
    Set XLBOOK = XLAPP.Workbooks.Add
    Set XLSHEET = XLBOOK.Worksheets.Add
    '
    Dim Fila, COLUMNA, ColumnasVisibles As Long
    Dim DatoCelda As String
    Dim Encabezado As Integer
    Dim MensajeError As String

    If Grilla.Rows < 1 Then Exit Sub
    
    Screen.MousePointer = 11
    
    EncabezadoFila = 0: TITULO = TRIM$(TITULO)
    If TITULO <> "" Then
        XLSHEET.Cells(1, 1) = TITULO
        With XLSHEET.Cells(1, 1)
            .Font.Name = "Arial"
            .Font.Size = 14
            .Font.Bold = True
            .Font.Underline = True
            .Font.Italic = True
        End With
        XLSHEET.Range("A1:Z1").Merge
        XLSHEET.Range("A1:Z1").HorizontalAlignment = xlLeft
        Encabezado = 2
    End If
    'CARGA EN UN VECTOR LA ALINEACION QUE TIENE LA COLUMNA
    'VALIDA CON ESTE DATO SI UN DATO NUMERICO SOLO SE TRATA COMO TAL SI ESTA ALINEADO A DERECHA, CASO
    'CONTRARIO LO TOMA COMO TEXTO.(CASO AHP CODIGO EXT. TIENE FORMATO NUMERICO CON 0 ADELANTE Y LOS MISMOS
    'LOS BORRABA EJ. 000100 LO MOSTRABA EN EL EXCEL COMO 100
    Dim FORMATOCOL()
    ReDim Preserve FORMATOCOL(Grilla.Cols - 1)
    For i& = 0 To Grilla.Cols - 1
       Grilla.col = i&
       FORMATOCOL(i&) = Grilla.ColAlignment(i&)
    Next i&
    '
    FIN& = Grilla.Rows - 1
    For Fila = 0 To Grilla.Rows - 1
       ColumnasVisibles = 0
       If Not (FormTraza Is Nothing) Then Call FormTraza.Trazador(CLng(Fila), FIN&)
       For COLUMNA = 0 To Grilla.Cols - 1
            If Grilla.ColWidth(COLUMNA) = 0 Then GoTo 10 ' OBVIA LAS COLUMNAS OCULTAS
            ColumnasVisibles = ColumnasVisibles + 1
            '
            DatoCelda = TRIM$(Grilla.TextMatrix(Fila, COLUMNA))
            '
            If COLUMNA = 0 And PrimerColumnaSeleccion = True Then
                If DatoCelda = "*" Then GoTo 10 ' OBVIA LA COLUMNA CON EL ASTERISCO
                If DatoCelda = "" Then GoTo 15 ' OBVIA LA FILA. NO FUE SELECCIONADA.
            End If
            
            Grilla.row = Fila: Grilla.col = COLUMNA
            If Fila < 2 Then
                FuenteNombre = Grilla.CellFontName
                Tamanio = Grilla.CellFontSize
                Negrita = Grilla.CellFontBold
                Italica = Grilla.CellFontItalic
                Subrayado = Grilla.CellFontUnderline
                Alinear = Grilla.ColAlignment(COLUMNA)
            End If
            
            ColorFuente = XVALO(Grilla.CellForeColor)
            If ColorFuente < 0 Then ColorFuente = 0      ' NEGRO POR DEFAULT
            If ColorFuente > 2147483647 Then ColorFuente = 0
            ColorFondo = XVALO(Grilla.CellBackColor)
            If ColorFondo < 0 Then ColorFondo = 16777215 ' BLANCO POR DEFAULT
            If ColorFondo > 2147483647 Then ColorFondo = 16777215
            
            With XLSHEET.Cells(Encabezado + Fila + 1, ColumnasVisibles)
                If Fila < 2 Then
                    '
                    If IsDate(DatoCelda) And Len(DatoCelda) > 6 Then
                      If Mid$(DatoCelda, 3, 1) = "/" And Mid$(DatoCelda, 6, 1) = "/" Then
                         DatoCelda = Format$(DatoCelda, "yyyy-mm-dd")
                      End If
                    End If
                    '
                    .Font.Name = FuenteNombre
                    .Font.Size = Tamanio
                    .Font.Bold = Negrita
                    .Font.Underline = Subrayado
                    .Font.Italic = Italica
                        Call LongToRGB(ColorFuente, rd%, gr%, BL%)
                    .Font.COLOR = RGB(rd%, gr%, BL%)
                        rd% = 255: gr% = 255: BL% = 255
                        If ColorFondo <> 0 Then
                            Call LongToRGB(ColorFondo, rd%, gr%, BL%)
                        End If
                    If Not (rd% = 255 And gr% = 255 And BL% = 255) Then
                        .Interior.COLOR = RGB(rd%, gr%, BL%)
                    End If
                
                    If IsNumeric(DatoCelda) And (Alinear = 6 Or Alinear = 7 Or Alinear = 8) Then 'INGRESA AQUI SI ES NUMERO Y EN LA GRILLA ESTA ALINEADO A DERECHA
                        PosicionComa% = InStr(1, DatoCelda, ",")
                        PosicionPunto% = InStr(1, DatoCelda, ".")
                        ' LA EXISTENCIA DEL PUNTO Y DE LA COMA SIGNIFICA QUE TIENE SEPARADOR DECIMAL Y DE MILES
                        If PosicionComa% > 0 And PosicionPunto% > 0 Then
                            DatoCelda = REPLACAR$(DatoCelda, ".", "") ' QUITO EL SEPARADOR DE MILES
                            DatoCelda = REPLACAR$(DatoCelda, ",", ".") ' REEMPLAZA LA COMA POR EL PUNTO PARA LA PARTE DECIMAL
                            PosicionPunto% = InStr(1, DatoCelda, ".")
                            
                            If PosicionPunto% > 0 Then
                                LongitudDecimal% = Len((Mid$(DatoCelda, PosicionPunto% + 1)))
                                FormatoDecimal$ = ""
                                Do While Len(FormatoDecimal$) < LongitudDecimal%
                                    FormatoDecimal$ = FormatoDecimal$ & "0"
                                Loop
                            End If
                        End If
    
                        .NumberFormat = "#,##0"
                        If LongitudDecimal% > 0 Then .NumberFormat = "#,##0." & FormatoDecimal$
    
                        XLSHEET.Cells(Encabezado + Fila + 1, ColumnasVisibles) = DatoCelda
                        
                    ElseIf IsDate(DatoCelda) And ((Mid$(DatoCelda, 5, 1) = "-" And Mid$(DatoCelda, 8, 1) = "-") Or (Mid$(DatoCelda, 5, 1) = "/" And Mid$(DatoCelda, 8, 1) = "/")) Then
                        .NumberFormat = "d/m/yyyy;@"
                        XLSHEET.Cells(Encabezado + Fila + 1, ColumnasVisibles) = DatoCelda
                        
                    ElseIf InStr(1, DatoCelda, "$") > 0 Then
                        .NumberFormat = "$#,##0.00"
                        XLSHEET.Cells(Encabezado + Fila + 1, ColumnasVisibles) = DatoCelda
                    
                    Else
                        .NumberFormat = "@"
                        XLSHEET.Cells(Encabezado + Fila + 1, ColumnasVisibles) = DatoCelda
                    
                    End If
                    
                    Select Case Alinear
                        Case 1: .HorizontalAlignment = xlLeft
                        Case 4: .HorizontalAlignment = xlCenter
                        Case 6, 7, 8: .HorizontalAlignment = xlRight:
                        Case Else: .HorizontalAlignment = xlLeft
                    End Select
                    .VerticalAlignment = xlCenter
                
                Else
                    
                    If IsNumeric(DatoCelda) And (Alinear = 6 Or Alinear = 7 Or Alinear = 8) Then 'INGRESA AQUI SI ES NUMERO Y EN LA GRILLA ESTA ALINEADO A DERECHA
                        PosicionComa% = InStr(1, DatoCelda, ",")
                        PosicionPunto% = InStr(1, DatoCelda, ".")
                        ' LA EXISTENCIA DEL PUNTO Y DE LA COMA SIGNIFICA QUE TIENE SEPARADOR DECIMAL Y DE MILES
                        If PosicionComa% > 0 And PosicionPunto% > 0 Then
                            DatoCelda = REPLACAR$(DatoCelda, ".", "") ' QUITO EL SEPARADOR DE MILES
                        End If
                    End If
                    '
                    If IsDate(DatoCelda) And Len(DatoCelda) > 6 Then
                      If Mid$(DatoCelda, 3, 1) = "/" And Mid$(DatoCelda, 6, 1) = "/" Then
                         DatoCelda = Format$(DatoCelda, "yyyy-mm-dd")
                      End If
                    End If
                    '
                        Call LongToRGB(ColorFuente, rd%, gr%, BL%)
                    .Font.COLOR = RGB(rd%, gr%, BL%)
                        rd% = 255: gr% = 255: BL% = 255
                        If ColorFondo <> 0 Then
                            Call LongToRGB(ColorFondo, rd%, gr%, BL%)
                        End If
                    If Not (rd% = 255 And gr% = 255 And BL% = 255) Then
                        .Interior.COLOR = RGB(rd%, gr%, BL%)
                    End If
                    
                    XLSHEET.Cells(Encabezado + Fila + 1, ColumnasVisibles) = DatoCelda
                End If
                
                .Borders.LineStyle = xlContinuous
            
            End With
10     Next COLUMNA
15  Next Fila
'   LetraColumna$ = ExcelConvierteColumnaNumeroALetra(Grilla.Cols)
'   XLSHEET.Range("A" & CStr(Encabezado + 1) & ":" & LetraColumna$ & CStr(Encabezado + 1)).Borders.LineStyle = xlContinuous
'   XLSHEET.Range("A" & CStr(Encabezado + 1) & ":" & LetraColumna$ & CStr(Encabezado + 1)).Borders.Weight = xlMedium
'
'   XLSHEET.Range(LetraColumna$ & CStr(Encabezado + 1) & ":" & LetraColumna$ & CStr(Encabezado + Grilla.Rows)).Borders.LineStyle = xlContinuous
'   XLSHEET.Range(LetraColumna$ & CStr(Encabezado + 1) & ":" & LetraColumna$ & CStr(Encabezado + Grilla.Rows)).Borders.Weight = xlMedium
'
'   XLSHEET.Range("A" & CStr(Encabezado + Grilla.Rows) & ":" & LetraColumna$ & CStr(Encabezado + Grilla.Rows)).Borders.LineStyle = xlContinuous
'   XLSHEET.Range("A" & CStr(Encabezado + Grilla.Rows) & ":" & LetraColumna$ & CStr(Encabezado + Grilla.Rows)).Borders.Weight = xlMedium
'
'   XLSHEET.Range("A" & CStr(Encabezado + 1) & ":A" & CStr(Encabezado + Grilla.Rows)).Borders.LineStyle = xlContinuous
'   XLSHEET.Range("A" & CStr(Encabezado + 1) & ":A" & CStr(Encabezado + Grilla.Rows)).Borders.Weight = xlMedium
  
   XLSHEET.Columns.AutoFit
   '
   Screen.MousePointer = 1
   '
   XLSHEET.SaveAs NombreDelArchivo
   On Error GoTo 0
   '
   XLAPP.Quit      ' cierra EXCEL que se esta ejecutando minimizado
   Set XLAPP = Nothing     ') libera memoria
   Set XLBOOK = Nothing    ') utilizada por
   Set XLSHEET = Nothing   ') las variables objeto
   
   Screen.MousePointer = 1
   Call COMUNI("Planilla Excel Creada Exitosamente:\ \" + REPLACAR$(NombreDelArchivo, "\", "/"))
   '
   
99: Exit Sub
   
ErrInstanciaExcel:
   '
   MensajeError = TRIM$(Err.Description)
   On Error GoTo 0
   XLAPP.Quit      ' cierra EXCEL que se esta ejecutando minimizado
   Set XLAPP = Nothing     ') libera memoria
   Set XLBOOK = Nothing    ') utilizada por
   Set XLSHEET = Nothing   ') las variables objeto
   
   Screen.MousePointer = 1
   Call MENSERR(24, "Imposible Crear Planilla Excel.\ \" & MensajeError)
   Exit Sub
   '
End Sub


Sub EDITARFILA(MSFORIGIEN As msFlexGrid, MSFDESTINO As msFlexGrid, REG&)
    'CARGA DE UN MSF DE ORIGEN EN UN MSF DESTINO SOLO LA CELDA PASA EN REG&
    MSFDESTINO.Rows = 1: MSFDESTINO.Rows = 2
    MSFDESTINO.Cols = MSFORIGIEN.Cols
    For i& = 1 To MSFORIGIEN.Cols - 1
    
       MSFDESTINO.TextMatrix(1, i&) = MSFORIGIEN.TextMatrix(REG&, i&)
    Next i&
End Sub

Sub GeneraExcelDeGrillaSinDisenio(NombreDelArchivo As String, Grilla As Object, TITULO As String, Optional PrimerColumnaSeleccion As Boolean)
   
    Dim XLAPP As Excel.Application
    Dim XLBOOK As Excel.Workbook
    Dim XLSHEET As Excel.Worksheet
    '
    On Error GoTo ErrInstanciaExcel
    '
    Set XLAPP = New Excel.Application
    Set XLBOOK = XLAPP.Workbooks.Add
    Set XLSHEET = XLBOOK.Worksheets.Add
    '
    Dim Fila, COLUMNA, ColumnasVisibles, FILAEXCEL As Long
    Dim DatoCelda As String
    Dim Encabezado As Integer
    Dim MensajeError As String

    If Grilla.Rows < 1 Then Exit Sub
    
    Screen.MousePointer = 11
    
    EncabezadoFila = 0: TITULO = TRIM$(TITULO)
    If TITULO <> "" Then
        XLSHEET.Cells(1, 1) = TITULO
        With XLSHEET.Cells(1, 1)
            .Font.Name = "Arial"
            .Font.Size = 14
            .Font.Bold = True
            .Font.Underline = True
            .Font.Italic = True
        End With
        XLSHEET.Range("A1:Z1").Merge
        XLSHEET.Range("A1:Z1").HorizontalAlignment = xlLeft
        Encabezado = 2
    End If
    
    For Fila = 0 To Grilla.Rows - 1
       ColumnasVisibles = 0
       FILAEXCEL = FILAEXCEL + 1
       For COLUMNA = 0 To Grilla.Cols - 1
       
            If Grilla.ColWidth(COLUMNA) = 0 Then GoTo 10 ' OBVIA LAS COLUMNAS OCULTAS
            ColumnasVisibles = ColumnasVisibles + 1
            '
            DatoCelda = TRIM$(Grilla.TextMatrix(Fila, COLUMNA))
            '
            If Fila > 0 And COLUMNA = 0 And PrimerColumnaSeleccion = True Then
                If DatoCelda = "*" Then GoTo 10 ' OBVIA LA COLUMNA CON EL ASTERISCO
                If DatoCelda = "" Then
                    FILAEXCEL = FILAEXCEL - 1
                    GoTo 15 ' OBVIA LA FILA. NO FUE SELECCIONADA.
                End If
            End If

            
            Grilla.row = Fila: Grilla.col = COLUMNA
            Alinear = Grilla.ColAlignment(COLUMNA)
            '
            If IsDate(DatoCelda) And Len(DatoCelda) > 6 Then
              If Mid$(DatoCelda, 3, 1) = "/" And Mid$(DatoCelda, 6, 1) = "/" Then
                 DatoCelda = Format$(DatoCelda, "yyyy-mm-dd")
              End If
            End If
            '
            With XLSHEET.Cells(Encabezado + FILAEXCEL + 1, ColumnasVisibles)
                If IsNumeric(DatoCelda) Then
                    PosicionComa% = InStr(1, DatoCelda, ",")
                    PosicionPunto% = InStr(1, DatoCelda, ".")
                    LongitudDecimal% = 0: FormatoDecimal$ = ""
                    ' LA EXISTENCIA DEL PUNTO Y DE LA COMA SIGNIFICA QUE TIENE SEPARADOR DECIMAL Y DE MILES
                    If PosicionComa% > 0 And PosicionPunto% > 0 Then
                        DatoCelda = REPLACAR$(DatoCelda, ".", "") ' QUITO EL SEPARADOR DE MILES
                        DatoCelda = REPLACAR$(DatoCelda, ",", ".") ' REEMPLAZA LA COMA POR EL PUNTO PARA LA PARTE DECIMAL
                        PosicionPunto% = PosicionComa%
                    End If
                    If PosicionPunto% > 0 Then
                        LongitudDecimal% = Len((Mid$(DatoCelda, PosicionPunto% + 1)))
                        Do While Len(FormatoDecimal$) < LongitudDecimal%
                            FormatoDecimal$ = FormatoDecimal$ & "0"
                        Loop
                    End If
                    .NumberFormat = "###0"
                    If LongitudDecimal% > 0 Then .NumberFormat = "#,##0." & FormatoDecimal$
                    XLSHEET.Cells(Encabezado + FILAEXCEL + 1, ColumnasVisibles) = DatoCelda
                    
                ElseIf IsDate(DatoCelda) And ((Mid$(DatoCelda, 5, 1) = "-" And Mid$(DatoCelda, 8, 1) = "-") Or (Mid$(DatoCelda, 5, 1) = "/" And Mid$(DatoCelda, 8, 1) = "/")) Then
                    .NumberFormat = "d/m/yyyy;@"
                    XLSHEET.Cells(Encabezado + FILAEXCEL + 1, ColumnasVisibles) = DatoCelda
                    
                ElseIf InStr(1, DatoCelda, "$") > 0 Then
                    .NumberFormat = "$#,##0.00"
                    XLSHEET.Cells(Encabezado + FILAEXCEL + 1, ColumnasVisibles) = DatoCelda
                
                Else
                    .NumberFormat = "@"
                    XLSHEET.Cells(Encabezado + FILAEXCEL + 1, ColumnasVisibles) = DatoCelda
                
                End If
                
                Select Case Alinear
                    Case 1: .HorizontalAlignment = xlLeft
                    Case 4: .HorizontalAlignment = xlCenter
                    Case 6, 7, 8: .HorizontalAlignment = xlRight:
                    Case Else: .HorizontalAlignment = xlLeft
                End Select
                .VerticalAlignment = xlCenter
            
            End With
10     Next COLUMNA
15  Next Fila

   XLSHEET.Columns.AutoFit
   '
   Screen.MousePointer = 1
   '
   XLSHEET.SaveAs NombreDelArchivo
   On Error GoTo 0
   '
   XLAPP.Quit      ' cierra EXCEL que se esta ejecutando minimizado
   Set XLAPP = Nothing     ') libera memoria
   Set XLBOOK = Nothing    ') utilizada por
   Set XLSHEET = Nothing   ') las variables objeto
   
   Screen.MousePointer = 1
   Call COMUNI("Planilla Excel Creada Exitosamente:\ \" + REPLACAR$(NombreDelArchivo, "\", "/"))
   '
   
99: Exit Sub
   
ErrInstanciaExcel:
   '
   MensajeError = TRIM$(Err.Description)
   On Error GoTo 0
   XLAPP.Quit      ' cierra EXCEL que se esta ejecutando minimizado
   Set XLAPP = Nothing     ') libera memoria
   Set XLBOOK = Nothing    ') utilizada por
   Set XLSHEET = Nothing   ') las variables objeto
   
   Screen.MousePointer = 1
   Call MENSERR(24, "Imposible Crear Planilla Excel.\ \" & MensajeError)
   Exit Sub
   '
End Sub



Sub DisenioDeCelda(Grilla As msFlexGrid, Fila As Variant, COLUMNA As Variant, FUENTE As String, Tamanio As Integer, Negrita As Boolean, Italica As Boolean, Subrayado As Boolean, Alinear As Integer, ColorFuente As Long, ByRef ColorFondo As Long)
   
    Dim filaInicio, FilaFin As Integer
    Dim columnaInicio, ColumnaFin As Integer
        
    If InStr(1, Fila, ":") > 0 And InStr(1, COLUMNA, ":") > 0 Then
        posicion = InStr(1, Fila, ":")
        filaInicio = Val(Mid$(Fila, 1, posicion - 1))
        FilaFin = Val(Mid$(Fila, posicion + 1))
        
        posicion = InStr(1, COLUMNA, ":")
        columnaInicio = Val(Mid$(COLUMNA, 1, posicion - 1))
        ColumnaFin = Val(Mid$(COLUMNA, posicion + 1))
    Else
        If (Fila < 0 Or Fila > Grilla.Rows - 1) Or (COLUMNA < 0 And COLUMNA > Grilla.Cols - 1) Then Exit Sub
        filaInicio = Fila: FilaFin = Fila
        columnaInicio = COLUMNA: ColumnaFin = COLUMNA
    End If
    
    For Y = filaInicio To FilaFin
        Grilla.row = Y
        For X = columnaInicio To ColumnaFin
            Grilla.col = X
            With Grilla
                If FUENTE <> "" Then
                    .CellFontName = FUENTE
                End If
                
                If Tamanio > 0 Then
                    .CellFontSize = Tamanio
                End If
                
                If Negrita = True Then
                    .CellFontBold = True
                End If
                
                If Italica = True Then
                    .CellFontItalic = True
                End If
                
                If Subrayado = True Then
                    .CellFontUnderline = True
                End If
                
                '0 El contenido de la celda se alinea a la izquierda, arriba.
                '1 Predeterminado para las cadenas. El contenido de la celda se alinea a la izquierda, centro.
                '2 El contenido de la celda se alinea a la izquierda, abajo.
                '3 El contenido de la celda se alinea al centro y arriba.
                '4 El contenido de la celda se alinea al centro, centro.
                '5 El contenido de la celda se alinea al centro, abajo.
                '6 El contenido de la celda se alinea a la derecha, arriba.
                '7 Predeterminado para los números. El contenido de la celda se alinea a la derecha, centro.
                '8 El contenido de la celda se alinea a la derecha, abajo.
                '9 El contenido de la celda tiene la alineación general. Esta corresponde a "izquierda, centro" para las cadenas y "derecha, centro" para los números
                
                If Alinear > 0 Then
                    .CellAlignment = Alinear
                End If
                
                If ColorFuente > 0 Then
                    .CellForeColor = ColorFuente
                End If
                
                If ColorFondo > 0 Then
                    .CellBackColor = ColorFondo
                End If
                
            End With
        Next X
    Next Y
    
End Sub
Public Sub LongToRGB(ByVal lngColor As Long, intRed As Integer, intGreen As Integer, intBlue As Integer)

    intRed = lngColor Mod &H100
    lngColor = lngColor \ &H100
    intGreen = lngColor Mod &H100
    lngColor = lngColor \ &H100
    intBlue = lngColor Mod &H100
    
End Sub
Sub cmdEliminaItem(MSF As msFlexGrid, REG&)
    If REG& <= 0 Then
        MsgBox "Debe Seleccionar una fila"
    ElseIf MSF.Rows = 2 Then
        MSF.Rows = 1
    Else
      Call MSF.RemoveItem(REG&)
    End If

End Sub
Sub ELIMINADUPLICADOS(MSF As msFlexGrid, col&)
    'COMPARA QUE EN UNA DETERMINADA COLUMNA NO SE REPITA SI ES ASI LO ELIMINA DIRECTAMENTE
     While MSF.Rows - 1 > i&
       i& = i& + 1
       LEIDO$ = MSF.TextMatrix(i&, col&)
       While MSF.Rows - 1 > J&
          J& = J& + 1
          ACOMPARAR$ = MSF.TextMatrix(J&, col&)
          If TRIM$(UCase$(LEIDO$)) = TRIM$(UCase$(ACOMPARAR$)) And J& <> i& Then
             Call cmdEliminaItem(MSF, J&)
             J& = J& - 1
          End If
       Wend
       J& = 0
     Wend
    
End Sub

Function STRINGFILA$(MSF As msFlexGrid, REG&, Optional ApartirdeColum%, Optional AGREGAFORMATO$)
'DEVUELVE UN STRING CON TODOS LOS DATOS SEPARADOS POR PUNTO Y COMA
'SI ApartirdeColum% > 0 RECORRE A PARTIR DE ESTA COLUMNA
'SI AGREGAFORMATO$ <> "" LE AGREGA UN FORMATO POR DEFECTO A TODOS IGUALES.
   STRINGFILA$ = ""
   CANTREGIS& = MSF.Rows
   If CANTREGIS& <= 1 Then Exit Function
   For i& = ApartirdeColum% To MSF.Cols - 1
     If TRIM$(AGREGAFORMATO$) <> "" Then
       STRINGFILA$ = STRINGFILA$ + MSF.TextMatrix(REG&, i&) & AGREGAFORMATO$ & ";"
     Else
       STRINGFILA$ = STRINGFILA$ + MSF.TextMatrix(REG&, i&) + ";"
     End If
   Next i&
End Function
Function STRINGFILA_STRING$(MSF As msFlexGrid, REG&, Optional ApartirdeColum%, Optional AGREGAFORMATO$)
'DEVUELVE UN STRING CON TODOS LOS DATOS CONCATENADOS DE LA FILA
'CREADA PARA SABER SI LA FILA NO TIENE NINGUN DATO EN NINGUNA CELDA
   STRINGFILA_STRING$ = ""
   CANTREGIS& = MSF.Rows
   If CANTREGIS& <= 1 Then Exit Function
   For i& = ApartirdeColum% To MSF.Cols - 1
       STRINGFILA_STRING$ = STRINGFILA_STRING$ + TRIM$(MSF.TextMatrix(REG&, i&))
   Next i&
End Function

Function SUMA_CANT_SEGUNDATO(MSF As msFlexGrid, DATO1$, COLUMDATO&, COLUMNUMERICO&)
    'EJEMPLO
    'PASO UN CODIGO EN DATO$ Y EL NUMERO DE COLUMNA DONDE SE UBICA
    'RECORRE LA GRILLA Y SI ENCUENTRA EL DATO$  SUMA LA VALOR DE LA COLUMNA COLUMNUMERICO&
    VALORSUMADO = 0
    If COLUMDATO& < 1 Then
       Call COMUNI("FALTA NUMERO DE COLUMNA DATO")
       Exit Function
    End If
    If COLUMNUMERICO& < 1 Then
       Call COMUNI("FALTA NUMERO DE COLUMNA VALOR A SUMAR")
       Exit Function
    End If
    
    For i& = 1 To MSF.Rows - 1
       If TRIM$(DATO1$) = TRIM$(MSF.TextMatrix(i&, COLUMDATO&)) Then
          VALORSUMADO = VALORSUMADO + XVALO(MSF.TextMatrix(i&, COLUMNUMERICO&))
       End If
    Next i&
    SUMA_CANT_SEGUNDATO = VALORSUMADO
End Function
Function DEVUELVE_COLOR_SEGUN_NUMERO%(COLOR%, COLORESCLAROS%)
   'SOLO PONE LOS COLORES MAS CLAROS
   'VALIDACION PARA EL CASO EN EL QUE EL COLOR SE PASE P.EJ. EN UN FOR Y LA VARIABLE SEA INCEREMENTAL CUANDOSE PASA COMIENZA DE 1
10 If COLOR% > 48 Then
     COLOR% = COLOR% - 48
     If COLOR% > 48 Then
        COLOR% = COLOR% - 48
        GoTo 10
     End If
   End If
   '
   If COLORESCLAROS% > 0 Then
    Select Case COLOR%
       Case Is < 2
         COLOR% = 2
       Case Is < 8
         COLOR% = 8
       Case Is < 15
         COLOR% = 15
       Case Is < 19
         COLOR% = 19
       Case Is < 24
         COLOR% = 24
       Case Is < 34
         COLOR% = 34
       Case Is < 36
         COLOR% = 35
       Case Is < 37
         COLOR% = 36
       Case Is < 40
         COLOR% = 40
       Case Is < 44
         COLOR% = 44
       Case Is < 46
         COLOR% = 45
       Case Is < 56
         COLOR% = 48
         
       Case Else
          COLOR% = 8
    End Select
   End If
   DEVUELVE_COLOR_SEGUN_NUMERO% = COLOR%
End Function

Sub PASAR_LIST_a_Grilla(List As ListBox, MSF As msFlexGrid, POSICIONESLIST$)
    'EJEMPLO "1,12;15,30;44,9;53,9; 62,10;74,9;84,24"
    Dim COMIENZO%()
    Dim FINAL1%()
    If InStr(POSICIONESLIST$, ";") < 1 Then Exit Sub
    Call TEXTOLOTES$(POSICIONESLIST$, ";")
    If UBound(CADENATEX$) < 1 Then Exit Sub
    For i& = 1 To UBound(CADENATEX$)
       Dato$ = CADENATEX$(i&)
       If InStr(Dato$, ",") > 0 Then
          POS1% = InStr(Dato$, ",")
          ReDim Preserve COMIENZO%(i&), FINAL1%(i&)
          COMIENZO%(i&) = XVALO(Mid$(Dato$, 1, POS1% - 1))
          FINAL1%(i&) = XVALO(Mid$(Dato$, POS1% + 1))
       End If
    Next i&
    '
    For i& = 0 To List.ListCount - 1
       TX$ = List.List(i&)
        REG& = MSF.Rows
        MSF.Rows = MSF.Rows + 1
       For J& = 1 To UBound(CADENATEX$)
         MSF.TextMatrix(REG&, J&) = TRIM$(Mid$(TX$, COMIENZO(J&), FINAL1%(J&)))
       Next J&
    Next i&
    '
End Sub

Function GENERA_FORMATOS$(Campos$)
    '"Código/A16;Descripción/A20;U-M/A3;Cantidad/f10.1;F.Solic./D8;Costo/F10.2;Importe/F12.2;Formula/F10.2;%/F7;Venta/F10.2;NUMORDEN/F0,NUMCOMPRA/F0;Entregar/D8"
    If TRIM$(Campos$) <> "" Then
        Erase CADENATEX$
        Call TEXTOLOTES$(Campos$, "/")
        GENERA_FORMATOS$ = ""
        For i& = 2 To UBound(CADENATEX$)
           POS1% = InStr(CADENATEX$(i&), ";")
           If POS1% > 0 Then
             GENERA_FORMATOS$ = GENERA_FORMATOS$ & Mid$(CADENATEX$(i&), 1, POS1% - 1) & ";"
           End If
        Next i&
        GENERA_FORMATOS$ = Left$(GENERA_FORMATOS$, Len(GENERA_FORMATOS$) - 1)
    End If
End Function
Sub AGREGAFILA(MSF As msFlexGrid, FORMATOS$, FilaDondeSeAgrega&)
    'GUARDA LOS FORMATOS DE CADA COLUMNA EN LA GRILLA OCULTA EN POSICION FilaDondeSeAgrega&
    Call TEXTOLOTES$(FORMATOS$, ";")
    On Error Resume Next
    If MSF.Rows - 1 < FilaDondeSeAgrega& Then MSF.Rows = FilaDondeSeAgrega& + 1
    If UBound(CADENATEX$) > 0 Then
        For i& = 1 To UBound(CADENATEX$)
          MSF.TextMatrix(FilaDondeSeAgrega&, i&) = CADENATEX$(i&)
        Next i&
    End If
    On Error GoTo 0
End Sub
Function FormatoSegunCampo$(MSF As msFlexGrid, Fila&, COLUMNA&)
     FormatoSegunCampo$ = TRIM$(MSF.TextMatrix(Fila&, COLUMNA&))
End Function
Function GENERA_FORMATOS_CON_BARRA$(Campos$)
    '"Código/A16;Descripción/A20;U-M/A3;Cantidad/f10.1;F.Solic./D8;Costo/F10.2;Importe/F12.2;Formula/F10.2;%/F7;Venta/F10.2;NUMORDEN/F0,NUMCOMPRA/F0;Entregar/D8"
    Call TEXTOLOTES$(Campos$, "/")
    GENERA_FORMATOS_CON_BARRA$ = ""
    For i& = 2 To UBound(CADENATEX$)
       POS1% = InStr(CADENATEX$(i&), ";")
       If POS1% > 0 Then
         ESPACIOX% = XVALO(Mid$(CADENATEX$(i&), 2)): If ESPACIOX% < 1 Then ESPACIOX% = 1
         GENERA_FORMATOS_CON_BARRA$ = GENERA_FORMATOS_CON_BARRA$ & Space(ESPACIOX%) & "/" & Mid$(CADENATEX$(i&), 1, POS1% - 1) & ";"
       Else
         GENERA_FORMATOS_CON_BARRA$ = GENERA_FORMATOS_CON_BARRA$ & Space(ESPACIOX%) & "/" & CADENATEX$(i&) & ";"
       
       End If
    Next i&
    GENERA_FORMATOS_CON_BARRA$ = Left$(GENERA_FORMATOS_CON_BARRA$, Len(GENERA_FORMATOS_CON_BARRA$) - 1)
End Function

Sub COLOREAR_CELDA_EDITAB_BLANCO(MSF As msFlexGrid, CELDASEDITABLES$)
    COLU_EDIT$ = TRIM$(CELDASEDITABLES$)
    If COLU_EDIT$ <> "" Then 'SI ES EDITABLE ENTRA ACA
      Call TEXTOLOTES$(COLU_EDIT$, ";")
      On Error Resume Next
      MSF.Redraw = False
      For i& = 1 To UBound(CADENATEX$) 'RECORRO EL VECTOR QUE TIENE CUALES SON LAS CELDAS EDITABLES
         COLUEDIT& = XVALO(CADENATEX$(i&))
         For F& = 1 To MSF.Rows - 1
          For COLUX% = 1 To MSF.Cols - 1
             If COLUEDIT& = COLUX% Then
               MSF.row = F&: MSF.col = COLUX%
               MSF.CellBackColor = vbWhite
             End If
          Next COLUX%
         Next F&
      Next i&
      MSF.Redraw = True
    End If
End Sub
Sub Asterizco_En_Celda(MSF As msFlexGrid, REGISTRO, Optional COLUMNA%, Optional COLOR)
   'PINTA DE UN COLOR LA CELDA 0 Y MARCA CON UN ASTERIZCO, CUANDO VUELVE A HACER CLIC
   'SI ESTA CON ASTERIZCO LO LIMPIA Y LO VUELVE AL COLOR ORIGINAL BackColorFixed
   With MSF
        .row = XVALO(REGISTRO)               ' ASIGNA EL NUMERO DE REGISTRO POSICIONADO
        If COLUMNA% < 1 Then COLUMNA% = 0
        .col = COLUMNA%
        If .TEXT = "*" Then
           .TEXT = ""
           If COLOR <> "" Then .CellBackColor = BackColorFixed
        Else
           .TEXT = "*"
           If COLOR <> "" Then .CellBackColor = COLOR
        End If
   End With

End Sub

Sub ACTUALIZA_FILA(MSF As msFlexGrid, REG&)
   For i& = 1 To MSF.Cols - 1
      MSF.TextMatrix(REG&, i&) = REPROGRID.TXTCAJA(i&)
   Next i&
   MSF.TopRow = REG&
   MSF.row = REG&
   MSF.ColSel = 1
   On Error Resume Next
   MSF.SetFocus
   On Error GoTo 0
   
End Sub
Sub ORDENARCOLUFECHA(MSF As msFlexGrid, col%)
    For i& = 1 To MSF.Rows - 1
      MSF.TextMatrix(i&, col%) = FECHALETRA$(FECHANUM(MSF.TextMatrix(i&, col%)))
    Next i&
End Sub
Function FECHALETRA$(NUM)

    'PASA LOS NUMEROS A LETRAS PARA UN ORDENAMIENTO MAS SEGURO
    NUME$ = TRIM$(Str(NUM)) 'NUM ES LA FECHA EN VALOR NUMERICO
    FLTRA$ = ""
    For i% = 1 To Len(NUME$)
       NU& = XVALO(Mid(NUME$, i%, 1)) + 65 '(A PARTIR DEL 65 EMPIEZAN LAS LETRAS)
       FLTRA$ = FLTRA$ + Chr(NU&)
    Next i%
    FECHALETRA$ = FLTRA$
End Function
Sub PASARTEXTOAFECHA(MSF As msFlexGrid, col%)
  'PASA A FECHA LAS LETRAS TRANSFORMADA CON LA FUNCION FECHALETRA$ SE USA PARA ORDENAR LAS FECHAS EN UNA GRILLA
  For J& = 1 To MSF.Rows - 1
    TXTLET$ = MSF.TextMatrix(J&, col%)
    FLTRA$ = ""
    For i& = 1 To Len(TXTLET$)
       LET1$ = Mid$(TXTLET$, 1, 1)
       FLTRA$ = FLTRA$ & XVALO(Asc(LET1$)) - 65
       If Len(TXTLET$) > 0 Then TXTLET$ = Mid(TXTLET$, 2)
    Next i&
    FENU = XVALO(FLTRA$)
    MSF.TextMatrix(J&, col%) = FECHATEX$(FENU)
  Next J&
End Sub

Sub INSERTFILAVACIA(MSF As msFlexGrid, REG&)
    With MSF
        'AGREGA UNA FILA
        .Rows = .Rows + 1
        'RECORRE DESDE LA ANTEULTIMA FILA Y VA COPIANDO EN LA SIGUIENTE HASTA LLEGAR A LA FILA SELECCIONADA
        For i& = .Rows - 1 To REG& Step -1
'           If I& = .Rows - 1 Then .Rows = .Rows + 1
           For J& = 1 To .Cols - 1
               .TextMatrix(i&, J&) = MSF.TextMatrix(i& - 1, J&)
           Next J&
        Next i&
        'BLANQUEA LA FILA SELECCIONADA
        For J& = 1 To MSF.Cols - 1
            .TextMatrix(REG&, J&) = ""
        Next
        

    End With

End Sub
Sub CargarComboSegunGrid(MSF As msFlexGrid, COMBO As comboBox, Optional PrimerafilaVacia%, Optional Campos$)
    'carga un combo segun el encabezado de una grilla.
    COMBO.Clear
    If Campos$ <> "" Then
       FORMATOX$ = GENERA_FORMATOS$(Campos$)
       Call TEXTOLOTES$(FORMATOX$, ";")
    End If
    If PrimerafilaVacia% > 0 Then COMBO.AddItem "" 'deja la primera   fila vacia
    For i& = 1 To MSF.Cols - 1
       On Error Resume Next
       If XVALO(Mid$(CADENATEX$(i&), 2)) > 0 Then
         NAMECOL$ = TRIM$(MSF.TextMatrix(0, i&))
         COMBO.AddItem NAMECOL$
       End If
       On Error GoTo 0
    Next i&

End Sub
Sub INICIALIZAR_CHECKLIST_ENGRILLA(MSF As msFlexGrid, COLUM%, Optional REG&, Optional ESTADO%)
       'RUTINA QUE AGREGA A UNA COLUMNA UN CHECK LIST
       MSF.col = COLUM%
       VALORESTADO% = 168 ' FALSE POR DEFECTO
       If ESTADO% > 0 Then VALORESTADO% = 254 'TRUE  SI SE LE PASA VALOR 1
       
       If REG& > 0 Then 'SI ES PARA UN SOLO REGISTRO
         MSF.row = REG&
         MSF.CellFontName = "Wingdings"
         MSF.CellFontSize = 14
         MSF.CellAlignment = flexAlignCenterCenter
         '
         MSF.TextMatrix(REG&, COLUM%) = Chr(VALORESTADO%)
         Exit Sub
       End If
       '
       'SI NO VIENE AQUI Y RECORRE Y A TODA LA COLUMNA LE APLICA EL CASILLERO SEGUN ESTADO PASADO POR PARAMETRO
       For i& = 1 To MSF.Rows - 1
         MSF.row = i&
         MSF.CellFontName = "Wingdings"
         MSF.CellFontSize = 14
         MSF.CellAlignment = flexAlignCenterCenter
         MSF.TextMatrix(i&, COLUM%) = Chr(VALORESTADO%)
       Next i&

End Sub
      
Function MENORHORA(MSF As msFlexGrid, COL1%)
     'DEVUELVE EL MENOR VALOR DE HORA DE UNA GRILLA SEGUN COLUMNA PASADA COMO PARAMETRO
     'EL FORMATO QUE LEE DEBE ESTAR TIPO HORATEX 10:00:00
     'VA A DEVOLVER 1000 SI SE QUIERE SE DEBERA CAMBIAR A HORATEX
     For i& = 1 To MSF.Rows - 1
        HORALEIDA = HORANUM(MSF.TextMatrix(i&, COL1%))
        If i& = 1 Then HoraMenor = HORALEIDA
        If HORALEIDA < HORALEIDA_ANT Then HoraMenor = HORALEIDA
        HORALEIDA_ANT = HORALEIDA
     Next i&
     
     MENORHORA = HoraMenor
End Function
Function MAYORHORA(MSF As msFlexGrid, COL1%)
     'DEVUELVE EL MAYOR VALOR DE HORA DE UNA GRILLA SEGUN COLUMNA PASADA COMO PARAMETRO
     'EL FORMATO QUE LEE DEBE ESTAR TIPO HORATEX 10:00:00
     'VA A DEVOLVER 1000 SI SE QUIERE SE DEBERA CAMBIAR A HORATEX
     For i& = 1 To MSF.Rows - 1
        HORALEIDA = HORANUM(MSF.TextMatrix(i&, COL1%))
        If i& = 1 Then HoraMayor = HORALEIDA
        If HORALEIDA > 0 Then
          If HORALEIDA > HoraMayor Then HoraMayor = HORALEIDA
        End If
     Next i&
     MAYORHORA = HoraMayor
     
End Function

Function VALIDAR_DUPLICADO(MSF As msFlexGrid, COLUM%, Optional NO_VALIDAR_CERO%)

    'ESTA FUNCION VALIDA SEGUN COLUMNA PARAMETRIZADA SI EL VALOR SE REPITE
    'ANTE LA PRIMER REPETICION DEVUELVE EL NUMERO DE FILA DONDE SE DETECTA LA REPETICION
    Dim i As Long, X As Long
    While i < MSF.Rows - 1
       i = i + 1
       X = 0
        While X < MSF.Rows - 1
        X = X + 1
        'SE AGREGO ESTO PARA QUE SI EL PARAMETRO NO_VALIDAR_CERO% ES 1 NO VALIDE LA REPETICION DE DATOS CON 0 '22/03/18'
        If NO_VALIDAR_CERO% > 0 Then
           If XVALO(MSF.TextMatrix(i, COLUM%)) = 0 Then
              GoTo 30
           End If
        End If
        'FIN AGREGADO
        If MSF.TextMatrix(i, COLUM%) = MSF.TextMatrix(X, COLUM%) And X <> i Then
             VALIDAR_DUPLICADO = X
             Exit Function
        End If
       Wend
30 Wend

End Function

Function VALIDAR_ECO(MSF As msFlexGrid, COLUECO%, MOSTRARMENSAJE%)
   'ESTA FUNCION VALIDA SI EL DATO QUE ESTA EN LA COLUMANA PARAMETRIZADA  ESTA VACIA PRESENTA UN MENSAJE
   'DEVOLVIENDO EN VALIDAR_ECO EL NUMERO DE FILA , USAR PARA CAMPOS ECO DONDE SI NO TIENE ECO EL ID NO DEBERIA SSER VALIDO
   MOSTRARMENSAJE% = 0
   For i& = 1 To MSF.Rows - 1
     If TRIM$(MSF.TextMatrix(i&, COLUECO%)) = "" Then
        If MOSTRARMENSAJE% > 0 Then Call COMUNI("Fila " & i& & " Con Campo Eco No Válido")
        VALIDAR_ECO = i&
        Exit For
     End If
   Next i&
   
End Function

Function VALIDARHORAS&(MSF As msFlexGrid, COL_DESDE%, COL_HASTA%)
    'VALIDA EN UNA GRILLA QUE TENGA INGRESO DE HORA DESDE Y HORA HASTA
    'DEVUELVE EL NUMERO DE FILA EN VALIDARHORAS& DONDE ES SUPERIOR LA FECHA DESDE QUE LA FECHA HASTA EJ.: 09 A 08
    For i& = 1 To MSF.Rows - 1
       HORDES = HORANUM(MSF.TextMatrix(i&, COL_DESDE%))
       HORHAS = HORANUM(MSF.TextMatrix(i&, COL_HASTA%))
       If HORHAS < HORDES Then
          VALIDARHORAS& = i&
       End If
    Next i&
    A = HORHAS - HORDES
    DIF = HORATEX(A)
End Function

Function FilaVacia%(MSF As msFlexGrid, Optional ELIMINARVACIAS%, Optional CONSIDERA0%, Optional HASTACOLUMNA%)
  'DEVUELVE 1 SI LA FILA ESTA VACIA O SEA QUE NO TIENE NINNGUN VALOR INGRESADO EN NINGUNA CELDA DE LAS VISIBLES EN LA GRILLA.
  'For i& = 1 To MSF.Rows - 1
  XEROX$ = ""
  If HASTACOLUMNA% < 1 Then HASTACOLUMNA% = MSF.Cols - 1
  If CONSIDERA0% < 1 Then XEROX$ = "0"
  Do While i& < MSF.Rows - 1
    i& = i& + 1
    For J& = 1 To HASTACOLUMNA%
      VALORING$ = TRIM$(MSF.TextMatrix(i&, J&))
      
      If VALORING$ = "" Or VALORING$ = XEROX$ Then
        FilaVacia% = 1
      Else
        FilaVacia% = 0
        Exit For
      End If
    Next J&
    If ELIMINARVACIAS% > 0 And FilaVacia% = 1 Then
         Call cmdEliminaItem(MSF, i&)
         i& = i& - 1
    End If

  Loop
End Function
Function VALORCELDA1#(UBICAX$, GRID2, Optional PARAMREG&)
      'recordar que esta rutina lee o posiciones o numeros como valor
      col& = columnax&(UBICAX$, FILAX&, PARAMREG&)
      REG& = FILAX&
      If PARAMREG& > 0 Then REG& = PARAMREG&
      COL1% = col&
      '=EXCEL;(12/11)*(12/11)*13
      CALCULOX$ = VENTANA.Formula(COL1%)
      CELDA$ = CALCULOX$
      CELDA$ = REPLACAR$(CELDA$, " ", "")
      CELDA$ = UCase$(CELDA$)
'      CELDA$ = "=A1+(A2-(A3/A11))"

      If Left$(CELDA$, 1) <> "=" Then
         VALORCELDA1# = XVALO(CELDA$)
      Else
        If Left$(UCase$(TRIM$(CELDA$)), 7) = "=EXCEL;" Then
           CELDA$ = Mid$(CELDA$, 8)
        ElseIf Left$(CELDA$, 1) = "=" Then
           CELDA$ = Mid$(CELDA$, 2)
        End If
        '=suma(E3;E4;E5;E6;E7;E8)
        '=(E3+E4(E5*E6)+ (E75E8))+
        UBICAX1$ = ""
        For i% = 1 To Len(CELDA$)
          CARACTERX$ = Mid$(CELDA$, i%, 1)
          
          Select Case CARACTERX$
            Case "=" 'NO HACE NADA - NO SE ESCRIBE
            Case "(", "+", "-", "*", "/", ")", ";"
              If CARACTERX$ = ";" Then CARACTERX$ = "+"
              ULTIMO_CARACTER$ = CARACTERX$
              ACUMUCARACTERES$ = ACUMUCARACTERES$ + CARACTERX$
              If UBICAX1$ <> "" Then
               VALORINDIVIDUAL$ = UBICAX1$
               If Asc(Left$(UBICAX1$, 1)) >= 65 And Asc(Left$(UBICAX1$, 1)) <= 90 Then ' SI NO ES LETRA (NO ES UNA CELDA SI NO VALOR)
                 VALORINDIVIDUAL$ = LEOCELDA1$(UBICAX1$, REG&)
               End If
                 If Left$(VALORINDIVIDUAL$, 1) = "-" Then
                    VALORINDIVIDUAL$ = REPLACAR(VALORINDIVIDUAL$, "-", "|") 'el caracter 124 reemplaza al signo - de los negativos
                 End If                                                     'para que no se detecto dos simbolos juntos y romper el funcionamiento programado
                 TEXTOFINAL$ = TEXTOFINAL$ + VALORINDIVIDUAL$
              End If
              UBICAX1$ = ""
              TEXTOFINAL$ = TEXTOFINAL$ & CARACTERX$
            Case Else
              UBICAX1$ = UBICAX1$ + CARACTERX$
              ACUMUCARACTERES$ = ""
            End Select
        Next i%
      End If
      'PARA LA ULTIMA LECTURA VA A PASAR POR ACA
      If TRIM$(UBICAX1$) <> "" Then
        VALORINDIVIDUAL$ = UBICAX1$
        If Asc(Left$(UBICAX1$, 1)) >= 65 And Asc(Left$(UBICAX1$, 1)) <= 90 Then  ' SI NO ES LETRA (NO ES UNA CELDA SI NO VALOR)
           VALORINDIVIDUAL$ = LEOCELDA1$(UBICAX1$, REG&)
        End If
        If Left$(VALORINDIVIDUAL$, 1) = "-" Then
         VALORINDIVIDUAL$ = REPLACAR(VALORINDIVIDUAL$, "-", "|")
        End If
      End If
      TEXTOFINAL$ = TEXTOFINAL$ + VALORINDIVIDUAL$
      'AHORA VOY CALCULANDO CONSIDERANDO LOS PARENTESIS
      'VOY BUSCANDO EL PARENTESIS MAS INTERNO Y DE AHI PARA AFUERA
        '=(E3+E4(E5*E6)+ (E75E8))
      TEXTOFINAL2$ = TEXTOFINAL$
      POS1% = InStr(TEXTOFINAL2$, ")")
      If POS1% > 0 Then
            '(10+(20*30)+40+50)
            For L& = 1 To Len(TEXTOFINAL2$)
              P1% = InStr(TEXTOFINAL2$, "(")
              If P1% > 0 Then
                 TEXTOFINAL2$ = Mid$(TEXTOFINAL2$, P1% + 1)
              Else
                 On Error Resume Next
                 P2% = InStr(TEXTOFINAL2$, ")")
                 TEXTOFINAL2$ = Mid$(TEXTOFINAL2$, 1, P2% - 1)
                 If Err > 0 Then Call COMUNI("Error de Sintaxis en Celda " & UBICAX$): Exit Function
                 On Error GoTo 0

                 BLOQUE_A_CALCULAR$ = TEXTOFINAL2$
                 BLOQUE_A_CALCULAR$ = CALCULABLOQUE1(BLOQUE_A_CALCULAR$)
                 VALORCELDA1# = XVALO(CALCULABLOQUE1(BLOQUE_A_CALCULAR$))
                 TEXTOFINAL$ = REPLACAR$(TEXTOFINAL$, "(" & TEXTOFINAL2$ & ")", TRIM$(Str$(VALORCELDA1#)))
                 TEXTOFINAL2$ = TEXTOFINAL$
                 If InStr(TEXTOFINAL2$, "(") < 1 Then
                   'SI YA NO POSEE MAS PARENTESIS PUEDE SER QUE QUEDE UNA OPERACION POR ESO LO PASO DEVUELTA AL VALORCELDA
                   BLOQUE_A_CALCULAR$ = TEXTOFINAL2$
                   VALORCELDA1# = XVALO(CALCULABLOQUE1(BLOQUE_A_CALCULAR$))
                   Exit For
                 End If
              End If
            Next L&
            
      Else
        BLOQUE_A_CALCULAR$ = TEXTOFINAL2$
        VALORCELDA1# = XVALO(CALCULABLOQUE1(BLOQUE_A_CALCULAR$))
      End If
End Function
Function columnax&(POS_CELDA_LEIDA$, FILAX&, PARAMREG&)
   'LA VARIABLE FILAX& LA DEVUELVE COMO PARAMETRO DE SALIDA
   For J% = 1 To Len(POS_CELDA_LEIDA$)  ' RECORRO A6
         LETRAX$ = Mid$(POS_CELDA_LEIDA$, J%, 1) 'A
         If Not IsNumeric(LETRAX$) Then
             If J% = 0 Then
               COLU1& = COLU1& + (Asc(LETRAX$) - 64)
               MULTIP& = COLU1&
             Else
               'EN ESTA VUELTA SE CONSIDERA LAS LETRAS DOBLES EJEMPLO AB
               On Error Resume Next
               COLU1& = COLU1& + (Asc(LETRAX$) - 64) + (25 * MULTIP&) ' LE SUMO 25 POR QUE ES EL TOTAL DE LETRAS - 1
               MULTIP& = COLU1&
               On Error GoTo 0
             End If
             FINLEN% = J%
          End If
      Next J%
      columnax& = COLU1&
      FILAX& = XVALO(Mid$(POS_CELDA_LEIDA$, FINLEN% + 1))
      If PARAMREG& > 0 Then FILAX& = PARAMREG&
      
End Function
Function UBICACIONX1$(REG&, col&)
            LETRAX$ = Chr$(col& + 64)
            NUMBERX$ = SAFETEXT$(REG&)
            UBICACIONX1$ = LETRAX$ & NUMBERX$
End Function

Function CALCULABLOQUE1(BLOQUEPARACALCULARORIG$)
    BLOQUEPARACALCULAR$ = BLOQUEPARACALCULARORIG$
    'ACA YA VOY A TENER LOS VALORES
    BLOQUEPARACALCULAR$ = TRIM$(REPLACAR$(BLOQUEPARACALCULAR$, " ", "")) 'QUITO LOS ESPACIOS
    'A PARTIR DE ACA CODIGO NUEVO
    'SEPARO DIVISION
'    For I& = 1 To Len(BLOQUEPARACALCULAR$)
'       '10/5*6*10.658+25.6359/-11.55
'       POS1% = InStr(LECTURAPREVIA$, "/")
'       If POS1% > 0 Then
'          LECTURAPREVIA$ = Mid$(BLOQUEPARACALCULAR$, 1, POS1% - 1)
'          IF
'
'
'    Next I&
    'FIN CODIGO NUEVO
    
    Dim ULTIMONEGATIVO$()
    '
    'PRIMERO EJECUTA LAS MULTIPLICACIONES Y LUEGO LAS DIVISIONES Y VA REEMPLAZANDO POR EJEMPLO (10*10) POR 100
40      POS1% = InStr(BLOQUEPARACALCULAR$, "/")
    If POS1% > 0 Then
       LECTURAPREVIA$ = Left$(BLOQUEPARACALCULAR$, POS1% - 1)
       'LECTURAPREVIA$ = REPLACAR$(LECTURAPREVIA$, "-", "|")
       While InStr(LECTURAPREVIA$, "/") > 0
       P1% = InStr(LECTURAPREVIA$, "/"): If P1% > 0 Then LECTURAPREVIA$ = Mid$(LECTURAPREVIA$, P1% + 1)
       Wend
       R& = 0
       While InStr(LECTURAPREVIA$, "-") > 0
       R& = R& + 1
       P1% = InStr(LECTURAPREVIA$, "-"): If P1% > 0 Then LECTURAPREVIA$ = Mid$(LECTURAPREVIA$, P1% + 1)
       ReDim ULTIMONEGATIVO$(R&)
       ULTIMONEGATIVO$(R&) = LECTURAPREVIA$
       Wend
       If P1% > 0 Then SIMBOPREVIO$ = Mid$(ULTIMONEGATIVO$(R& - 1), P1%): If SIMBOPREVIO$ = "-" Then LECTURAPREVIA$ = "-"
       While InStr(LECTURAPREVIA$, "+") > 0
       P1% = InStr(LECTURAPREVIA$, "+"): If P1% > 0 Then LECTURAPREVIA$ = Mid$(LECTURAPREVIA$, P1% + 1)
       Wend
       While InStr(LECTURAPREVIA$, "*") > 0
       P1% = InStr(LECTURAPREVIA$, "*"): If P1% > 0 Then LECTURAPREVIA$ = Mid$(LECTURAPREVIA$, P1% + 1)
       Wend
       'LECTURAPREVIA$ = REPLACAR$(LECTURAPREVIA$, "|", "-") 'el caracter 124 reemplaza al signo - de los negativos
       '
       LECTURAPOST$ = Mid$(BLOQUEPARACALCULAR$, POS1% + 1)
       P1% = InStr(LECTURAPOST$, "/"): If P1% > 1 Then LECTURAPOST$ = Mid$(LECTURAPOST$, 1, P1% - 1) 'CAMBIE DE 0 A 1 POR SI TIENE VALOR NEGATIVO
       P1% = InStr(LECTURAPOST$, "-"): If P1% > 1 Then LECTURAPOST$ = Mid$(LECTURAPOST$, 1, P1% - 1)
       P1% = InStr(LECTURAPOST$, "+"): If P1% > 1 Then LECTURAPOST$ = Mid$(LECTURAPOST$, 1, P1% - 1)
       P1% = InStr(LECTURAPOST$, "*"): If P1% > 1 Then LECTURAPOST$ = Mid$(LECTURAPOST$, 1, P1% - 1)
       'LECTURAPOST$ = REPLACAR$(LECTURAPOST$, "|", "-") 'el caracter 124 reemplaza al signo - de los negativos
A = BLOQUEPARACALCULAR$
       Resultado = 0 'POR SI EL DIVISOR ES 0
       LECTURAPOST1$ = REPLACAR$(LECTURAPOST$, "|", "-"): LECTURAPREVIA1$ = REPLACAR$(LECTURAPREVIA$, "|", "-")
       If XVALO(LECTURAPOST1$) <> 0 Then
        Resultado = XVALO(LECTURAPREVIA1$) / XVALO(LECTURAPOST1$)
        REEMPLARESULTADO$ = TRIM$(FORMATNUM$(Resultado, "F16.6")): If REEMPLARESULTADO$ = "" Then REEMPLARESULTADO$ = "0"
        BLOQUECALCULADO$ = SAFETEXT$(LECTURAPREVIA$) & "/" & SAFETEXT$(LECTURAPOST$)
        'BLOQUEPARACALCULAR$ = REPLACAR$(BLOQUEPARACALCULAR$, "|", "-")
        BLOQUEPARACALCULAR$ = TRIM$(REPLACAR$(BLOQUEPARACALCULAR$, BLOQUECALCULADO$, REEMPLARESULTADO$))
        GoTo 40
       End If
    End If

20  POS1% = InStr(BLOQUEPARACALCULAR$, "*")
    If POS1% > 0 Then
       '10/5*6*10.658+25.6359-11.55
       LECTURAPREVIA$ = Left$(BLOQUEPARACALCULAR$, POS1% - 1)
       While InStr(LECTURAPREVIA$, "/") > 0
       P1% = InStr(LECTURAPREVIA$, "/"): If P1% > 0 Then LECTURAPREVIA$ = Mid$(LECTURAPREVIA$, P1% + 1)
       Wend
       While InStr(LECTURAPREVIA$, "-") > 0
       P1% = InStr(LECTURAPREVIA$, "-"): If P1% > 0 Then LECTURAPREVIA$ = Mid$(LECTURAPREVIA$, P1% + 1)
       Wend
       While InStr(LECTURAPREVIA$, "+") > 0
       P1% = InStr(LECTURAPREVIA$, "+"): If P1% > 0 Then LECTURAPREVIA$ = Mid$(LECTURAPREVIA$, P1% + 1)
       Wend
       While InStr(LECTURAPREVIA$, "*") > 0
       P1% = InStr(LECTURAPREVIA$, "*"): If P1% > 0 Then LECTURAPREVIA$ = Mid$(LECTURAPREVIA$, P1% + 1)
       Wend
       'LECTURAPREVIA$ = REPLACAR$(LECTURAPREVIA$, "|", "-") 'el caracter 124 reemplaza al signo - de los negativos
       '
       
       LECTURAPOST$ = Mid$(BLOQUEPARACALCULAR$, POS1% + 1)
       P1% = InStr(LECTURAPOST$, "/"): If P1% > 1 Then LECTURAPOST$ = Mid$(LECTURAPOST$, 1, P1% - 1)
       P1% = InStr(LECTURAPOST$, "-"): If P1% > 1 Then LECTURAPOST$ = Mid$(LECTURAPOST$, 1, P1% - 1)
       P1% = InStr(LECTURAPOST$, "+"): If P1% > 1 Then LECTURAPOST$ = Mid$(LECTURAPOST$, 1, P1% - 1)
       P1% = InStr(LECTURAPOST$, "*"): If P1% > 1 Then LECTURAPOST$ = Mid$(LECTURAPOST$, 1, P1% - 1)
       'LECTURAPOST$ = REPLACAR$(LECTURAPOST$, "|", "-") 'el caracter 124 reemplaza al signo - de los negativos
       'ASIGNO A OTRA VARIABLE PARA NO TENER PROBLEMAS CUANDO REALIZO EL REEMPLAZO EN LA FORMULA - CON EL TEMA DEL CARACTER 124
       LECTURAPOST1$ = REPLACAR$(LECTURAPOST$, "|", "-"): LECTURAPREVIA1$ = REPLACAR$(LECTURAPREVIA$, "|", "-")
       Resultado = XVALO(LECTURAPREVIA1$) * XVALO(LECTURAPOST1$)
       REEMPLARESULTADO$ = TRIM$(FORMATNUM$(Resultado, "F16.6")): If REEMPLARESULTADO$ = "" Then REEMPLARESULTADO$ = "0"
       BLOQUECALCULADO$ = SAFETEXT$(LECTURAPREVIA$) & "*" & SAFETEXT$(LECTURAPOST$)
       BLOQUEPARACALCULAR$ = REPLACAR$(BLOQUEPARACALCULAR$, "|", "-")
       BLOQUEPARACALCULAR$ = TRIM$(REPLACAR$(BLOQUEPARACALCULAR$, BLOQUECALCULADO$, REEMPLARESULTADO$))
       GoTo 20
    End If
    '
    'AHORA SOLO QUEDAN OPERACIONES CON + O -
    ACUVALOR = 0: VALO$ = "": SIMBOLO$ = ""
    For i& = 1 To Len(BLOQUEPARACALCULAR$)
       CARACTEREX$ = Mid$(BLOQUEPARACALCULAR$, i&, 1)
       CARACTEREX_Y_SIGUIENTE$ = Mid$(BLOQUEPARACALCULAR$, i&, 2)
       If CARACTEREX_Y_SIGUIENTE$ = "--" Then
         CARACTEREX$ = "+"
       End If
       Select Case CARACTEREX$
          Case "+", "-"
            ULTIMO_CARACTER$ = CARACTEREX$
            VALO$ = REPLACAR$(VALO$, "|", "-") 'el caracter 124 reemplaza al signo - de los negativos
            VALORX = XVALO(VALO$)
            If SIMB_ANT$ = "+" Then
               ACUVALOR = ACUVALOR + VALORX
            ElseIf SIMB_ANT$ = "-" Then
               VALO$ = REPLACAR$(VALO$, "|", "-") 'el caracter 124 reemplaza al signo - de los negativos
               ACUVALOR = ACUVALOR - VALORX
            Else
               'ACA DEBERIA ENTRAR SOLO LA PRIIMERA VEZ
               VALO$ = REPLACAR$(VALO$, "|", "-")
               VALORX = XVALO(VALO$)
               ACUVALOR = ACUVALOR + VALORX
            End If
            VALO$ = ""
            SIMB_ANT$ = CARACTEREX$
            If ULTIMO_CARACTER$ = "|" Then SIMB_ANT$ = "-"
          Case Else
            VALO$ = VALO$ & CARACTEREX$
       End Select
       If CARACTEREX_Y_SIGUIENTE$ = "--" Then: i& = i& + 1  ' 'CASO QUE RESTA UN VALOR NEGATIVO

65  Next i&
    If ULTIMO_CARACTER$ = "-" Then
      VALO$ = REPLACAR$(VALO$, "|", "-")
      ACUVALOR = ACUVALOR - XVALO(VALO$)
    ElseIf ULTIMO_CARACTER$ = "+" Then
      VALO$ = REPLACAR$(VALO$, "|", "-")
      ACUVALOR = ACUVALOR + XVALO(VALO$)
    Else
      BLOQUEPARACALCULAR$ = REPLACAR$(BLOQUEPARACALCULAR$, "|", "-") 'el caracter 124 reemplaza al signo - de los negativos
      ACUVALOR = ACUVALOR + XVALO(BLOQUEPARACALCULAR$)
    End If
    CALCULABLOQUE1 = XVALO(ACUVALOR)
End Function
Function CALCULABLOQUE_RESP(BLOQUE$)

         MsgBox "OJO ESTA RUTINA FUE COMENTADA CONSULTAR CON OMAR"
'        CALCULABLOQUE1 = BLOQUE$
'        ACUMUCARACTERES$ = ""
'        If InStr(BLOQUE$, "-") < 1 Then Exit Function
'        For i% = 1 To Len(BLOQUE$)
'          CARACTERX$ = Mid$(BLOQUE$, i%, 1)
'          TEXTX$ = ""
'          Select Case CARACTERX$
'            Case "(", "+", "-", "*", "/", ")", ";"
'              CARACTERYSIGUIENTE$ = Mid$(BLOQUE$, i%, 2)
'              Select Case CARACTERYSIGUIENTE$
'                 Case "*-", "/-", "+-", "--"
'                     CARACTERX$ = REPLACAR$(CARACTERX$, "*", "?@")
'                     CARACTERX$ = REPLACAR$(CARACTERX$, "/", "¿@")
'                     CARACTERX$ = REPLACAR$(CARACTERX$, "+", "!@")
'                     CARACTERX$ = REPLACAR$(CARACTERX$, "-", "¡@")
'                     i% = i% + 1
'              End Select
'            Case Else
'            End Select
'            ACUMUCARACTERES$ = ACUMUCARACTERES$ + CARACTERX$
'
'        Next i%
'        ACUMUCARACTERES$ = REPLACAR$(ACUMUCARACTERES$, "?@", "*|")
'        ACUMUCARACTERES$ = REPLACAR$(ACUMUCARACTERES$, "¿@", "/|")
'        ACUMUCARACTERES$ = REPLACAR$(ACUMUCARACTERES$, "!@", "+|")
'        ACUMUCARACTERES$ = REPLACAR$(ACUMUCARACTERES$, "¡@", "-|")
'
'
'        CALCULABLOQUE1 = ACUMUCARACTERES$

End Function
'Function CALCULABLOQUE1(BLOQUEPARACALCULARORIG$)
'    BLOQUEPARACALCULAR$ = BLOQUEPARACALCULARORIG$
'    'ACA YA VOY A TENER LOS VALORES
'    BLOQUEPARACALCULAR$ = TRIM$(REPLACAR$(BLOQUEPARACALCULAR$, " ", "")) 'QUITO LOS ESPACIOS
'    'A PARTIR DE ACA CODIGO NUEVO
'    'SEPARO DIVISION
''    For I& = 1 To Len(BLOQUEPARACALCULAR$)
''       '10/5*6*10.658+25.6359/-11.55
''       POS1% = InStr(LECTURAPREVIA$, "/")
''       If POS1% > 0 Then
''          LECTURAPREVIA$ = Mid$(BLOQUEPARACALCULAR$, 1, POS1% - 1)
''          IF
''
''
''    Next I&
'    'FIN CODIGO NUEVO
'
'    Dim ULTIMONEGATIVO$()
'    '
'    'PRIMERO EJECUTA LAS MULTIPLICACIONES Y LUEGO LAS DIVISIONES Y VA REEMPLAZANDO POR EJEMPLO (10*10) POR 100
'40      POS1% = InStr(BLOQUEPARACALCULAR$, "/")
'    If POS1% > 0 Then
'       LECTURAPREVIA$ = Left$(BLOQUEPARACALCULAR$, POS1% - 1)
'       'LECTURAPREVIA$ = REPLACAR$(LECTURAPREVIA$, "-", "|")
'       While InStr(LECTURAPREVIA$, "/") > 0
'       P1% = InStr(LECTURAPREVIA$, "/"): If P1% > 0 Then LECTURAPREVIA$ = Mid$(LECTURAPREVIA$, P1% + 1)
'       Wend
'       r& = 0
'       While InStr(LECTURAPREVIA$, "-") > 0
'       r& = r& + 1
'       P1% = InStr(LECTURAPREVIA$, "-"): If P1% > 0 Then LECTURAPREVIA$ = Mid$(LECTURAPREVIA$, P1% + 1)
'       ReDim ULTIMONEGATIVO$(r&)
'       ULTIMONEGATIVO$(r&) = LECTURAPREVIA$
'       Wend
'       If P1% > 0 Then SIMBOPREVIO$ = Mid$(ULTIMONEGATIVO$(r& - 1), P1%): If SIMBOPREVIO$ = "-" Then LECTURAPREVIA$ = "-"
'       While InStr(LECTURAPREVIA$, "+") > 0
'       P1% = InStr(LECTURAPREVIA$, "+"): If P1% > 0 Then LECTURAPREVIA$ = Mid$(LECTURAPREVIA$, P1% + 1)
'       Wend
'       While InStr(LECTURAPREVIA$, "*") > 0
'       P1% = InStr(LECTURAPREVIA$, "*"): If P1% > 0 Then LECTURAPREVIA$ = Mid$(LECTURAPREVIA$, P1% + 1)
'       Wend
'       'LECTURAPREVIA$ = REPLACAR$(LECTURAPREVIA$, "|", "-") 'el caracter 124 reemplaza al signo - de los negativos
'       '
'       LECTURAPOST$ = Mid$(BLOQUEPARACALCULAR$, POS1% + 1)
'       P1% = InStr(LECTURAPOST$, "/"): If P1% > 1 Then LECTURAPOST$ = Mid$(LECTURAPOST$, 1, P1% - 1) 'CAMBIE DE 0 A 1 POR SI TIENE VALOR NEGATIVO
'       P1% = InStr(LECTURAPOST$, "-"): If P1% > 1 Then LECTURAPOST$ = Mid$(LECTURAPOST$, 1, P1% - 1)
'       P1% = InStr(LECTURAPOST$, "+"): If P1% > 1 Then LECTURAPOST$ = Mid$(LECTURAPOST$, 1, P1% - 1)
'       P1% = InStr(LECTURAPOST$, "*"): If P1% > 1 Then LECTURAPOST$ = Mid$(LECTURAPOST$, 1, P1% - 1)
'       'LECTURAPOST$ = REPLACAR$(LECTURAPOST$, "|", "-") 'el caracter 124 reemplaza al signo - de los negativos
'A = BLOQUEPARACALCULAR$
'       Resultado = 0 'POR SI EL DIVISOR ES 0
'       LECTURAPOST1$ = REPLACAR$(LECTURAPOST$, "|", "-"): LECTURAPREVIA1$ = REPLACAR$(LECTURAPREVIA$, "|", "-")
'       If XVALO(LECTURAPOST1$) <> 0 Then
'        Resultado = XVALO(LECTURAPREVIA1$) / XVALO(LECTURAPOST1$)
'        REEMPLARESULTADO$ = TRIM$(FORMATNUM$(Resultado, "F16.6")): If REEMPLARESULTADO$ = "" Then REEMPLARESULTADO$ = "0"
'        BLOQUECALCULADO$ = SAFETEXT$(LECTURAPREVIA$) & "/" & SAFETEXT$(LECTURAPOST$)
'        'BLOQUEPARACALCULAR$ = REPLACAR$(BLOQUEPARACALCULAR$, "|", "-")
'        BLOQUEPARACALCULAR$ = TRIM$(REPLACAR$(BLOQUEPARACALCULAR$, BLOQUECALCULADO$, REEMPLARESULTADO$))
'        GoTo 40
'       End If
'    End If
'
'20  POS1% = InStr(BLOQUEPARACALCULAR$, "*")
'    If POS1% > 0 Then
'       '10/5*6*10.658+25.6359-11.55
'       LECTURAPREVIA$ = Left$(BLOQUEPARACALCULAR$, POS1% - 1)
'       While InStr(LECTURAPREVIA$, "/") > 0
'       P1% = InStr(LECTURAPREVIA$, "/"): If P1% > 0 Then LECTURAPREVIA$ = Mid$(LECTURAPREVIA$, P1% + 1)
'       Wend
'       While InStr(LECTURAPREVIA$, "-") > 0
'       P1% = InStr(LECTURAPREVIA$, "-"): If P1% > 0 Then LECTURAPREVIA$ = Mid$(LECTURAPREVIA$, P1% + 1)
'       Wend
'       While InStr(LECTURAPREVIA$, "+") > 0
'       P1% = InStr(LECTURAPREVIA$, "+"): If P1% > 0 Then LECTURAPREVIA$ = Mid$(LECTURAPREVIA$, P1% + 1)
'       Wend
'       While InStr(LECTURAPREVIA$, "*") > 0
'       P1% = InStr(LECTURAPREVIA$, "*"): If P1% > 0 Then LECTURAPREVIA$ = Mid$(LECTURAPREVIA$, P1% + 1)
'       Wend
'       'LECTURAPREVIA$ = REPLACAR$(LECTURAPREVIA$, "|", "-") 'el caracter 124 reemplaza al signo - de los negativos
'       '
'
'       LECTURAPOST$ = Mid$(BLOQUEPARACALCULAR$, POS1% + 1)
'       P1% = InStr(LECTURAPOST$, "/"): If P1% > 1 Then LECTURAPOST$ = Mid$(LECTURAPOST$, 1, P1% - 1)
'       P1% = InStr(LECTURAPOST$, "-"): If P1% > 1 Then LECTURAPOST$ = Mid$(LECTURAPOST$, 1, P1% - 1)
'       P1% = InStr(LECTURAPOST$, "+"): If P1% > 1 Then LECTURAPOST$ = Mid$(LECTURAPOST$, 1, P1% - 1)
'       P1% = InStr(LECTURAPOST$, "*"): If P1% > 1 Then LECTURAPOST$ = Mid$(LECTURAPOST$, 1, P1% - 1)
'       'LECTURAPOST$ = REPLACAR$(LECTURAPOST$, "|", "-") 'el caracter 124 reemplaza al signo - de los negativos
'       'ASIGNO A OTRA VARIABLE PARA NO TENER PROBLEMAS CUANDO REALIZO EL REEMPLAZO EN LA FORMULA - CON EL TEMA DEL CARACTER 124
'       LECTURAPOST1$ = REPLACAR$(LECTURAPOST$, "|", "-"): LECTURAPREVIA1$ = REPLACAR$(LECTURAPREVIA$, "|", "-")
'       Resultado = XVALO(LECTURAPREVIA1$) * XVALO(LECTURAPOST1$)
'       REEMPLARESULTADO$ = TRIM$(FORMATNUM$(Resultado, "F16.6")): If REEMPLARESULTADO$ = "" Then REEMPLARESULTADO$ = "0"
'       BLOQUECALCULADO$ = SAFETEXT$(LECTURAPREVIA$) & "*" & SAFETEXT$(LECTURAPOST$)
'       BLOQUEPARACALCULAR$ = REPLACAR$(BLOQUEPARACALCULAR$, "|", "-")
'       BLOQUEPARACALCULAR$ = TRIM$(REPLACAR$(BLOQUEPARACALCULAR$, BLOQUECALCULADO$, REEMPLARESULTADO$))
'       GoTo 20
'    End If
'    '
'    'AHORA SOLO QUEDAN OPERACIONES CON + O -
'    ACUVALOR = 0: VALO$ = "": SIMBOLO$ = ""
'    For i& = 1 To Len(BLOQUEPARACALCULAR$)
'       CARACTEREX$ = Mid$(BLOQUEPARACALCULAR$, i&, 1)
'       CARACTEREX_Y_SIGUIENTE$ = Mid$(BLOQUEPARACALCULAR$, i&, 2)
'       If CARACTEREX_Y_SIGUIENTE$ = "--" Then
'         CARACTEREX$ = "+"
'       End If
'       Select Case CARACTEREX$
'          Case "+", "-"
'            ULTIMO_CARACTER$ = CARACTEREX$
'            VALO$ = REPLACAR$(VALO$, "|", "-") 'el caracter 124 reemplaza al signo - de los negativos
'            VALORX = XVALO(VALO$)
'            If SIMB_ANT$ = "+" Then
'               ACUVALOR = ACUVALOR + VALORX
'            ElseIf SIMB_ANT$ = "-" Then
'               VALO$ = REPLACAR$(VALO$, "|", "-") 'el caracter 124 reemplaza al signo - de los negativos
'               ACUVALOR = ACUVALOR - VALORX
'            Else
'               'ACA DEBERIA ENTRAR SOLO LA PRIIMERA VEZ
'               VALO$ = REPLACAR$(VALO$, "|", "-")
'               VALORX = XVALO(VALO$)
'               ACUVALOR = ACUVALOR + VALORX
'            End If
'            VALO$ = ""
'            SIMB_ANT$ = CARACTEREX$
'            If ULTIMO_CARACTER$ = "|" Then SIMB_ANT$ = "-"
'          Case Else
'            VALO$ = VALO$ & CARACTEREX$
'       End Select
'       If CARACTEREX_Y_SIGUIENTE$ = "--" Then: i& = i& + 1  ' 'CASO QUE RESTA UN VALOR NEGATIVO
'
'65  Next i&
'    If ULTIMO_CARACTER$ = "-" Then
'      VALO$ = REPLACAR$(VALO$, "|", "-")
'      ACUVALOR = ACUVALOR - XVALO(VALO$)
'    ElseIf ULTIMO_CARACTER$ = "+" Then
'      VALO$ = REPLACAR$(VALO$, "|", "-")
'      ACUVALOR = ACUVALOR + XVALO(VALO$)
'    Else
'      BLOQUEPARACALCULAR$ = REPLACAR$(BLOQUEPARACALCULAR$, "|", "-") 'el caracter 124 reemplaza al signo - de los negativos
'      ACUVALOR = ACUVALOR + XVALO(BLOQUEPARACALCULAR$)
'    End If
'    CALCULABLOQUE1 = XVALO(ACUVALOR)
'End Function
Function LEOCELDA1$(UBICAX$, Optional PARAMETROREG&)
      'EL PARAMETRO PARAMETROREG& DEFINE EL REGISTRO  A LEER
      'POR QUE CUANDO SE GENERA LA FORMULA SE CARGA
      POS_CELDA_LEIDA$ = TRIM$(UBICAX$)
      col& = columnax&(POS_CELDA_LEIDA$, FILAX&, PARAMETROREG&)
      REG& = FILAX&
      'If PARAMETROREG& > 0 Then REG& = PARAMETROREG&
      On Error Resume Next
      LEOCELDA1$ = TRIM$(VENTABNEW.MSF.TextMatrix(REG&, col&))
      On Error GoTo 0
End Function

Sub eliminar_Repetidos(MSF As msFlexGrid)
    'ELIMINA LAS FILAS REPETIDAS DE UNA GRILLA
    Dim estaFila, unaFila As String
    Dim i, J, m, N As Integer
    
    estaFila = vbNullString
    unaFila = vbNullString
    
    With MSF
        m = 1 ' Apunta a la primera Fila
        N = .Rows - 1 ' Cantidad de filas del MSF
        
        Do Until m > N
            estaFila = vbNullString  'cadena con la Fila Actual
            
            For i = 0 To .Cols - 1
                estaFila = estaFila + .TextMatrix(m, i) ' Relleno la cadena con los datos de toda la fila
            Next
            
            J = 1
            Do Until J > N
                unaFila = vbNullString ' Fila a comparar
                
                For i = 0 To .Cols - 1
                    unaFila = unaFila + .TextMatrix(J, i) ' la relleno con los datos de toda la fila
                Next
                
                If estaFila = unaFila And J <> m Then
                    .RemoveItem (J) ' Si son iguales remuevo unaFila solamente
                    N = N - 1 ' tamaño actual del HierarchicalFlexgrid, al remover disminuye en uno
                End If
                
                J = J + 1
            Loop
            m = m + 1
        Loop
        
        'Chequeo si la última y la penúltima fila son iguales, porque no se eliminan si lo son
        For i = 0 To .Cols - 1
            estaFila = estaFila + .TextMatrix(.Rows - 1, i)
            unaFila = unaFila + .TextMatrix(.Rows - 2, i)
        Next
        
        If unaFila = estaFila Then .RemoveItem (.Rows - 2)
    End With
End Sub

Function ValorMaximoColumna(MSF As msFlexGrid, COLUM%)
   For i& = 1 To MSF.Rows - 1
     VALOR_LEIDO = XVALO(MSF.TextMatrix(i&, COLUM%))
     If VALOR_LEIDO > VALORMAX Then VALORMAX = VALOR_LEIDO
   Next i&
   ValorMaximoColumna = VALORMAX
End Function


Sub PEGARFILAEDITADA(MSFORIGIEN As msFlexGrid, MSFDESTINO As msFlexGrid, REG&)
    'PEGA LA FILA EDITADA EN UNA GRILLA EN EL REGISTRO PARAMETRIZADO SE USA CUANDO VUELVE DE LA FUNCION EDITAR FILA
    For i& = 1 To MSFORIGIEN.Cols - 1
       MSFDESTINO.TextMatrix(REG&, i&) = MSFORIGIEN.TextMatrix(1, i&)
    Next i&
End Sub
Sub AGREGAR_FILA_OTRO_MSF(MSFORIGIEN As msFlexGrid, MSFDESTINO As msFlexGrid, REG&)
    'PEGA LA FILA EN EL SISGUIENTE REGISTRO
    REGNUEVO = MSFDESTINO.Rows
    MSFDESTINO.Rows = MSFDESTINO.Rows + 1
    For i& = 1 To MSFORIGIEN.Cols - 1
       MSFDESTINO.TextMatrix(REGNUEVO, i&) = MSFORIGIEN.TextMatrix(REG&, i&)
    Next i&
End Sub

'Sub EDITARFILA(MSFORIGIEN As MSFlexGrid, MSFDESTINO As MSFlexGrid, REG&)
'    'CARGA DE UN MSF DE ORIGEN EN UN MSF DESTINO SOLO LA CELDA PASA EN REG&
'    MSFDESTINO.Rows = 1: MSFDESTINO.Rows = 2
'    MSFDESTINO.Cols = MSFORIGIEN.Cols
'    For i& = 1 To MSFORIGIEN.Cols - 1
'       MSFDESTINO.TextMatrix(1, i&) = MSFORIGIEN.TextMatrix(REG&, i&)
'    Next i&
'End Sub
Function UltimaFilaVacia%(MSF As msFlexGrid, Optional ELIMINARVACIAS%, Optional CONSIDERA0%, Optional HASTACOLUMNA%)
  'DEVUELVE 1 SI LA ultima FILA ESTA VACIA O SEA QUE NO TIENE NINNGUN VALOR INGRESADO EN NINGUNA CELDA DE LAS VISIBLES EN LA GRILLA.
  If MSF.Rows <= 1 Then
    UltimaFilaVacia% = 1
    Exit Function
  End If
  XEROX$ = ""
  If HASTACOLUMNA% < 1 Then HASTACOLUMNA% = MSF.Cols - 1
  If CONSIDERA0% < 1 Then XEROX$ = "0"
  Do While i& < MSF.Rows - 1
    i& = MSF.Rows - 1 'considera solo la ultima fila
    For J& = 1 To HASTACOLUMNA%
      VALORING$ = TRIM$(MSF.TextMatrix(i&, J&))
      
      If VALORING$ = "" Or VALORING$ = XEROX$ Then
        UltimaFilaVacia% = 1
      Else
        UltimaFilaVacia% = 0
        Exit For
      End If
    Next J&
    If ELIMINARVACIAS% > 0 And UltimaFilaVacia% = 1 Then
         Call cmdEliminaItem(MSF, i&)
         i& = i& - 1
    End If

  Loop
End Function
Sub BORRARFILASVACIAS(MSF As msFlexGrid, Optional ELIMINARVACIAS%, Optional CONSIDERA0%)
  'DEVUELVE 1 SI LA ultima FILA ESTA VACIA O SEA QUE NO TIENE NINNGUN VALOR INGRESADO EN NINGUNA CELDA DE LAS VISIBLES EN LA GRILLA.
  HASTACOLUMNA% = MSF.Cols - 1
  i& = 0
  Do While i& < MSF.Rows - 1
    i& = i& + 1 'considera solo la ultima fila
    For J& = 0 To HASTACOLUMNA%
      VALORING$ = TRIM$(MSF.TextMatrix(i&, J&))
      
      If VALORING$ = "" Then
        FilaVacia1% = 1
      Else
        FilaVacia1% = 0
        UltimaFilaConDato& = i&
        Exit For
      End If
    Next J&
    If ELIMINARVACIAS% > 0 And FilaVacia1% = 1 Then
         Call cmdEliminaItem(MSF, i&)
         i& = i& - 1
    End If

  Loop
End Sub

Function ANCHOVENTABNEW&(CamposConFormato$)
   FORMATOSCONBARRA$ = GENERA_FORMATOS_CON_BARRA$(CamposConFormato$)
   Call TEXTOLOTES(FORMATOSCONBARRA$, ";")
   On Error Resume Next
   ANCHOVENTABNEW& = 0
   For i% = 1 To UBound(CADENATEX$)
      ANCHOVENTABNEW& = ANCHOVENTABNEW& + Int(XVALO(Mid$(TRIM$(CADENATEX$(i%)), 3)))
   Next i%
   ANCHOVENTABNEW& = ANCHOVENTABNEW& * 1.1
   On Error GoTo 0
End Function
Function Linea_Caracteres(CamposConFormato$, MSF As msFlexGrid, Fila&, CARACTER$)
   'ESTA FUNCION SE LE PASA COMO ARGUMENTOS LOS
   'CamposConFormato$ = CAMPOS CON EL FORMATO NORMAL (Codigo/A16,Descripcion/A32;.....)
   'EL NUMERO DE FILA EN LA CUAL SE VA A ESCRIBIR
   'Y EL CARACTER (EJEMPLO: "=" )
   'EN BASE A ESTO VA A GENERAR UNA REPETICIION SEGUN LA LONGITUD DEL CAMPO REPITIENDO EL CARACTER TANTAS VECES HASTA COMPLETAR LA CELDA
   FORMATOSCONBARRA$ = GENERA_FORMATOS_CON_BARRA$(CamposConFormato$)
   Call TEXTOLOTES(FORMATOSCONBARRA$, ";")
   On Error Resume Next
   For i% = 1 To UBound(CADENATEX$)
      MSF.TextMatrix(Fila&, i%) = String$(XVALO(Mid$(TRIM$(CADENATEX$(i%)), 3)), CARACTER$)
   Next i%
   On Error GoTo 0
End Function

 Sub Carga_MSF_RecordsetConEco(rs, Campos$, MSF As msFlexGrid, Optional EsSelecho%)
 
    Dim NOCAMPO$(), FORMX$()
    ReDim NOCAMPO$(1), FORMX$(1), ECO$(1)
    LSTX$ = Campos$
    While LSTX$ <> ""
       PPXX% = InStr(LSTX$, ";")
       PPXX2% = InStr(LSTX$, ",")
       If PPXX2% > 0 And PPXX2% < PPXX% Then PPXX% = PPXX2%
       If PPXX% < 1 Then PPXX% = 1 + Len(LSTX$)
       '
       CANCA% = CANCA% + 1
       If CANCA% >= UBound(NOCAMPO$) Then
          ReDim Preserve NOCAMPO$(CANCA%), FORMX$(CANCA%), ECO$(CANCA%)
       End If
       NOCX$ = Left$(LSTX$, PPXX% - 1)
       LSTX$ = Mid$(LSTX$, PPXX% + 1)
       '
       ' FORMATO (DAT|TADEPOSI|BYTE) **** MODO=DAT ó SQL|TABLA|FORMATO BUSQUEDA COLUMNA ANTERIOR
       ECOTXT$ = "": ECOFI% = 0
       ECOIN% = InStr(1, NOCX$, "(")
       If ECOIN% > 0 Then ECOFI% = InStr(1, NOCX$, ")")
       If ECOFI% > 0 Then ECOTXT$ = Mid$(NOCX$, ECOIN% + 1, ECOFI% - ECOIN% - 1)
       If ECOIN% > 0 Then NOCX$ = Mid$(NOCX$, 1, ECOIN% - 1) & Mid$(NOCX$, ECOFI% + 1)
       '
       PPXX% = InStr(NOCX$, "/")
       If PPXX% < 1 Then PPXX% = 1 + Len(NOCX$)
       NOCAMPO$(CANCA%) = Left$(NOCX$, PPXX% - 1)
       FORMX$(CANCA%) = UCase$(Mid$(NOCX$, PPXX% + 1))
       If FORMX$(CANCA%) = "" Then FORMX$(CANCA%) = "A16"
       ECO$(CANCA%) = ECOTXT$
       '
    Wend
    
      MSF.Rows = 100
      ROWX& = 1
      With rs
      Do While Not .EOF
        'MSF.Rows = MSF.Rows + 1
        DEJALIBRE% = 0
        For J& = 0 To .Fields.Count - 1
          If J& + 1 > MSF.Cols - 1 Then MSF.Cols = J& + 2
          If EsSelecho% < 1 Then
            If VENTANA.DejarLibre(J& + 1) = "S" Then
              DEJALIBRE% = DEJALIBRE% + 1 ' PARA QUE DEJE LIBRE COLUMNAS SIN LLENAR DEL RECORDSET
              'INTERVENCION PARA QUE CARGUE EL ECO DIRECTAMENTE SI ESTA CONFIGURADO VENTANA.ColEcoCampo(J& + 0)
              'SE UTILIZO EN ORDEN DE FABRIC > FRM PROCEFRM > mnuppimpresora_Click
              If VENTANA.ColEcoCampo(J& + 0) <> "" Then 'si tiene campo eco lo asigna directamente.
                 If TRIM$(MSF.TextMatrix(ROWX&, J&)) <> "" Then
                    Dim strArray() As String
                    On Error Resume Next
                    strArray = Split(VENTANA.ColEcoCampo(J& + 0), ";")
                    CAMPO1$ = strArray(0)
                    campo2$ = strArray(1)
                    TABLAX1$ = strArray(2)
                    valor = TRIM$(MSF.TextMatrix(ROWX&, J&))
                    CONDI$ = CAMPO1$ & " = '" & valor & "'"
                    MSF.TextMatrix(ROWX&, J& + DEJALIBRE%) = VALOBADA(TABLAX1$, campo2$, CONDI$, "NOMESS")
                    On Error GoTo 0
                 End If
              End If
              'FIN INTERVENCION PARA QUE CARGUE EL ECO DIRECTAMENTE SEGUN CONFIGURACION DE CLASE VENTANA
              MSF.Cols = MSF.Cols + 1
            End If
          End If
          LEIDO$ = SAFETEXT$(.Fields(J&).Value)
          
          ' FORMATO (DAT|TADEPOSI|BYTE)
          If TRIM$(ECO(J& + 1)) <> "" Then
                separa1% = InStr(1, ECO(J& + 1), "|")
                separa2% = InStr(separa1% + 1, ECO(J& + 1), "|")
                If separa1% > 0 Then
                    MODO$ = Mid$(ECO(J& + 1), 1, separa1% - 1)
                    TABLA$ = Mid$(ECO(J& + 1), separa1% + 1, separa2% - separa1% - 1)
                    Formato$ = Mid$(ECO(J& + 1), separa2% + 1)
                End If
                If MODO$ = "DAT" Then
                    If Formato$ = "BYTE" Then
                        ArgBus$ = Chr$(CInt(MSF.TextMatrix(ROWX&, J&)))
                    ElseIf Formato$ = "INT" Then ArgBus$ = MKI$(XVALO(MSF.TextMatrix(ROWX&, J&)))
                    ElseIf Formato$ = "SINGLE" Then ArgBus$ = MKS$(XVALO(MSF.TextMatrix(ROWX&, J&)))
                    ElseIf Formato$ = "DOUBLE" Then ArgBus$ = MKD$(XVALO(MSF.TextMatrix(ROWX&, J&)))
                    Else
                        ArgBus$ = TRIM$(MSF.TextMatrix(ROWX&, J&))
                    End If
                End If
                LEIDO$ = ECOARCH$(TABLA$, ArgBus$)
          End If
          
          A = .Fields(J&).Name
          FORMA1$ = UCase$(Left$(FORMX$(J& + 1), 1))
          FORMACOMP$ = FORMX$(J& + 1)
          If Left$(FORMACOMP$, 1) = "F" Then
             MSF.TextMatrix(ROWX&, J& + 1 + DEJALIBRE%) = FORMATNUM$(XVALO(LEIDO$), "" & FORMACOMP$ & "")
          ElseIf FORMA1$ = "D" Then
             MSF.TextMatrix(ROWX&, J& + 1 + DEJALIBRE%) = FECHATEX$(CVINT(LEIDO$))
          Else
             MSF.TextMatrix(ROWX&, J& + 1 + DEJALIBRE%) = LEIDO$
          End If
        Next J&
        .MoveNext
        ROWX& = ROWX& + 1
        If ROWX& + 1 = MSF.Rows Then
           MSF.Rows = MSF.Rows + 100
        End If

       Loop
     End With
     MSF.Rows = ROWX&
     If ROWX& < 2 Then MSF.Rows = 2
 End Sub

Public Sub ResaltarFila(ByVal poGrilla As msFlexGrid, REG&)  'PINTA LA FILA SELECCIONADA DE UN FLEXGRID
    Dim i As Long
    With poGrilla
        If .Rows = 0 Then Exit Sub
        .col = 0
        .row = REG&
        .ColSel = .Cols - 1: .RowSel = .row
        .TopRow = REG&
        .SetFocus
    End With
End Sub
Function CANTFILAS_SELECCIONADAS&(MSF As msFlexGrid, CARACTER$)
   CANTFILAS_SELECCIONADAS& = 0
   For i& = 1 To MSF.Rows - 1
      If MSF.TextMatrix(i&, 0) = CARACTER$ Then
         CANTFILAS_SELECCIONADAS& = CANTFILAS_SELECCIONADAS& + 1
      End If
   Next i&
End Function

Sub Param_Print_Excel(RUTA_ARCH$, NombreHoja$, MSF As msFlexGrid, IMPRESORA$, Columnas_a_Leer$, A_Partir_Fila%, En_Columna%, Optional datofijo$, Optional SIVACIO%, Optional NOIMPRIMIR%, Optional ARCHIVOIMAGEN$, Optional TOPX, Optional LEFTX, Optional HEIGHTX, Optional WIDTHX)
    'Columnas_a_Leer$ = NUMEROS DE COLUMNAS QUE VA A LEER DEL MSF SEPARADOS PRO PUNTO Y COMA "1;3;5"
    'A_Partir_Fila% = A PARTIR DE QUE FILA DEL EXCEL EMPIEZA A ESCRIBIR LOS DATOS
    'En_Columna%= EN QUE COLUMNA DEL EXCEL VA A ESCRIBIR LOS DATOS
    'SI Columnas_a_Leer$ ES MAS DE UNA LAS ESCRIBE EN EL EXCEL EN LA COLUMNA DE AL LADO
    
    'EJEMPLO SI Columnas_a_Leer$ = "1;4" Y En_Columna% = 5
    'VA A ESCRIBIR LA INFORMACION EN EL EXCEL _
    EN LA COLUMNA 5 Y 6
    'SI SE PASA UN VALOR COMO DATO FIJO Y DENTRO DE LAS COLUMNAS A LEER INCLUYE EL 0 CUANDO LEA ESTA ASIGNA EL VALORFIJO
    
    Dim ObjetoExcel As Excel.Application
    Dim Libro As Excel.Workbook
    Dim Hoja As Excel.Worksheet
    Dim shp As Object
    Dim marco As Object
    Dim RutaQr As String
                           '******  HABRIA QUE VER DE TERMINARLO MODIFICANDO TEXTO FIJO DEL JPG POR LA VARIABLE ARCHIVOIMAGEN$  ***********************
    Call TEXTOLOTES(Columnas_a_Leer$, ";")
    
    Set ObjetoExcel = CreateObject("Excel.Application")
    With ObjetoExcel
        ' ARCHIVO EXCEL / ABRE EL LIBRO
        Set Libro = ObjetoExcel.Workbooks.Open(RUTA_ARCH$)
        Set Hoja = Libro.Sheets(NombreHoja$)
        On Error Resume Next
        Set marco = Hoja.Shapes("marcoqr")   ' marcoqr es el nombre del marco en Excel (es una forma geometrica, que se le asigna nombre arriba a la izquierda en el excel)
        On Error GoTo 0
        'ES UN TRATAMIENTO ESPECIAL PARA EYM Y LO NUEVO ES QUE BORRA LA IMAGEN PARA QUE  NO SE CARGUE EL EXCEL CON IMAGENES TRAS OTRA
        If TRIM$(ARCHIVOIMAGEN$) = "QR_ETIEMBALAJE.JPG" Then 'SI ES IMPRESION DE EMBALAJE (DESDE LOTEXSQR>ARCHIVO>IMPRIMIR>ETIQUETA EMBALAJE C15)
             If marco Is Nothing Then ' LA HOJA TIENE UN MARCON CON EL NOMBRE MARCOQR (UNA FIGURA)
               MsgBox "No se encontró el marco MARCOQR.", vbExclamation
               GoTo 99
             Else
               'Eliminar la imagen anterior PERO TOMA LA UBICACION DEL MARCO
               RutaQr = "C:\FLEXOFT\QR_ETIEMBALAJE.JPG"
               On Error Resume Next
               Hoja.Shapes("QR_ETIEMBALAJE.JPG").Delete 'Eliminar la imagen anterior
               On Error GoTo 0

                Set shp = Hoja.Shapes.AddPicture( _
                FileName:=RutaQr, _
                LinkToFile:=False, _
                SaveWithDocument:=True, _
                Left:=marco.Left + 5, _
                Top:=marco.Top + 5, _
                Width:=marco.Width - 10, _
                Height:=marco.Height - 10)

               ' Nombrar la imagen (para poder reemplazarla luego)
               'shp.Name = "QR_ETIEMBALAJE.JPG"
               shp.Name = ARCHIVOIMAGEN$


            End If
        Else
'         If TRIM$(ARCHIVOIMAGEN$) = "QR_ETIEMBALAJE.JPG" Then
'            Libro.Worksheets(NombreHoja$).Shapes.AddPicture "C:\FLEXOFT\QR_ETIEMBALAJE.JPG", False, True, LEFTX, TOPX, WIDTHX, HEIGHTX
'         Else
            If TRIM$(ARCHIVOIMAGEN$) <> "" Then
'                 Libro.Worksheets(NombreHoja$).Shapes.AddPicture "C:\FLEXOFT\QR_ETIQPTERM.JPG", False, True, LEFTX, TOPX, WIDTHX, HEIGHTX
                 Libro.Worksheets(NombreHoja$).Shapes.AddPicture "C:\FLEXOFT\" & ARCHIVOIMAGEN$, False, True, LEFTX, TOPX, WIDTHX, HEIGHTX

            End If
         End If
         
'        End If
         k& = 0
         For i& = 1 To MSF.Rows - 1
            For J& = 1 To UBound(CADENATEX$)
               Col_a_Leer% = CADENATEX$(J&)
               Dato$ = TRIM$(MSF.TextMatrix(i&, Col_a_Leer%))
               If Dato$ = "" And SIVACIO% < 1 Then
                  Dato$ = "-----"
               End If
               If Left$(Dato$, 3) = "BR-" Then Dato$ = REPLACAR(Dato$, "BR-", "PR-")
               
               If Col_a_Leer% = 0 Then
                  If datofijo$ <> "" And STRINGFILA_STRING$(MSF, i&) <> "" Then
                     Dato$ = datofijo$
                  Else
                      Dato$ = "--"
                  End If
               End If

               Libro.Worksheets(NombreHoja).Cells(A_Partir_Fila% + k&, En_Columna% + J& - 1) = Dato$
            Next J&
            k& = k& + 1
         Next i&
         
         ObjetoExcel.ActiveWorkbook.Save
         Set Hoja = Libro.Sheets(NombreHoja$)
         If NOIMPRIMIR% > 0 Then GoTo 99
         On Error Resume Next
         Hoja.PrintOut Copies:=1, Collate:=True, ActivePrinter:=IMPRESORA$
         'HOJA.PrintOut Copies:=1, Collate:=True 'asi sale por la impresora predeterminada
        On Error GoTo 0
    End With

    'CIERRA HOJA, LIBRO Y EXCEL
99  On Error Resume Next
    Set Hoja = Nothing
    Libro.Close: Set Libro = Nothing
    ObjetoExcel.Close: Set ObjetoExcel = Nothing
    Set shp = Nothing
    Set marco = Nothing

    On Error GoTo 0


End Sub


Sub VOLVER_ABLANCO_TODA_LA_GRILLA(MSF As msFlexGrid)
    For i& = 1 To MSF.Rows - 1
          Call COLOREAR_GRILLA_SOLAFILA(MSF, vbWhite, i&, , 1)
    Next i&
End Sub

 Sub BuscarTexto(MSF As msFlexGrid, FILABUSCAR&, ColumnaaBuscar%, Argumento$, FRM As Form)
'EJEMPLO DE LLAMADA DE LA CAJA DE TEXTO TXTBUSCADOR
'OJO FILABUSCAR& DEBE ESTAR COMO PUBLICA A NIVEL DE FORMULARIO DE DONDES SE LLAMA

'' Private Sub TXTBUSCADOR_Change()
''   If TRIM$(TXTBUSCADOR) = "" Then
''      Exit Sub
''    Else
''       FILABUSCAR& = 1
''    End If
''    '
''    Call BuscarTexto(MSF3, FILABUSCAR&, 1, TXTBUSCADOR.TEXT, Me)
''
''    '
''End Sub
'=======================================================================
'CODIGO PARA CONTINUAR LA BUSQUEDA CON F3
'Private Sub MSF3_KeyDown(KeyCode As Integer, Shift As Integer)
'   If KeyCode = 114 Then ' BUSQUEDA  POR TEXTO CON F3
'       If TRIM$(Me.TXTBUSCADOR) <> "" Then
'          Call BuscarTexto(MSF3, FILABUSCAR&, 1, TXTBUSCADOR, Me)
'       End If
'   End If
'End Sub
    If FILABUSCAR& > 0 Then
      FILAINI& = FILABUSCAR& + 1
      If FilaInicial& >= MSF.Rows - 1 Then FILABUSCAR& = 1
    End If
    '
    'RECORRO EL MSF QUE TIENE TODO E IMPRIMO EN msf  LOS QUE MACHEAN
    For i& = FILAINI& To MSF.Rows - 1
        If (InStr(UCase$(MSF.TextMatrix(i&, ColumnaaBuscar%)), UCase$(Argumento$))) > 0 Then
          Call SELECCIONARFILA(MSF, i&): FRM.FILABUSCAR& = i&
          Call UbicarFocoTop(MSF, i&)
          Exit For
        End If
    Next i&
    If i& >= MSF.Rows Then
       Call COMUNI("Finalizó la Busqueda\Acepte Mensaje y Repita")
       FILABUSCAR& = 1
    End If
    
End Sub

Sub SOLO_ABRIR_Y_CERRAR_EXCEL(RUTA_ARCH$)
        Dim ObjetoExcel As Excel.Application
        Dim Libro As Excel.Workbook
        Dim Hoja As Excel.Worksheet
        
        Set ObjetoExcel = CreateObject("Excel.Application")
        
        ' ARCHIVO EXCEL / ABRE EL LIBRO
        Set Libro = ObjetoExcel.Workbooks.Open(RUTA_ARCH$)
'        ObjetoExcel.Visible = True
        Call T_Espera(10)
'        ObjetoExcel.Visible = False
        ObjetoExcel.ActiveWorkbook.Save
        'CIERRA HOJA, LIBRO Y EXCEL
        On Error Resume Next
        Set Hoja = Nothing
        Libro.Close: Set Libro = Nothing
        ObjetoExcel.Close: Set ObjetoExcel = Nothing
        On Error GoTo 0
End Sub
Sub GeneraTablaAuxiliarCodigoDescripcion(NOMBRELISTA$, FECHA_FE%, TITULO$)
    FE% = HOY(HO$)
        CONDI$ = " LISTA =   '" & NOMBRELISTA$ & "'"
        EXILISTA% = CantRegistros&("ENCAPRFUN", CONDI$, "NOMESS")
        If EXILISTA% > 0 Then
          FE% = FECHA_FE%
          FEBD% = CVINT(VALOBADA("ENCAPRFUN", "FECHEMISI", CONDI$, "NOMESS"))
          GENERAREST% = 0
          If FEBD% < FE% Then
            IDLISTAX& = XVALO(VALOBADA("ENCAPRFUN", "ID_LISTA", CONDI$, "NOMESS"))
            Call BORRAREGISTROS("ENCAPRFUN", "ID_LISTA = " & IDLISTAX&, REGAF&, "NOMESS")
            Call BORRAREGISTROS("ENCAPRF", "ID_LISTA = " & IDLISTAX&, REGAF&, "NOMESS")
            GENERAREST% = 1
          End If
        Else
          GENERAREST% = 1
          CONDI$ = " ID_LISTA  > 0 "
          IDLISTAX& = 1 + XVALO(VALOBADA("ENCAPRFUN", "MAX(ID_LISTA)", CONDI$, "NOMESS"))
        End If
        If GENERAREST% > 0 Then
          SQLX$ = "INSERT INTO  ENCAPRFUN (ID_LISTA ,LISTA, Descripcion, ADDREGISTROS, USAFORMULA, BORRAR, INSERTAR, NOTILDE "
          SQLX$ = SQLX$ + ", CANTICAMPOS, CodiEmpr, USARECORD, FECHEMISI, TITURECORD, ConsultaSql)"
          SQLX$ = SQLX$ + "VALUES ( " & IDLISTAX&
          SQLX$ = SQLX$ + ",'" & NOMBRELISTA$ & "'"
          SQLX$ = SQLX$ + ",'" & TITULO$ & "'"
          SQLX$ = SQLX$ + ",1"
          SQLX$ = SQLX$ + ",1"
          SQLX$ = SQLX$ + " ,1"
          SQLX$ = SQLX$ + " ,0"
          SQLX$ = SQLX$ + " ,0"
          SQLX$ = SQLX$ + " ,7"
          SQLX$ = SQLX$ + " ,0"
          SQLX$ = SQLX$ + " ,0"
          SQLX$ = SQLX$ + " ,'07/03/23'"
          SQLX$ = SQLX$ + " ,''"
          SQLX$ = SQLX$ + " ,'SELECT CAMPO1,CAMPO2 FROM DATASQL WITH(NOLOCK)'"
          SQLX$ = SQLX$ + ")"
          FLEXCONN.Execute (SQLX$)
            SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
            SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
            SQLX$ = SQLX$ + " ,0"
            SQLX$ = SQLX$ + " ,0"
            SQLX$ = SQLX$ + " , 'CODIGO/A12'"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " , 'S'"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " )"
            FLEXCONN.Execute (SQLX$)
        
        '
            SQLX$ = "INSERT INTO  ENCAPRF (ID_LISTA ,CodiEmpr, MARBORRA, Encabezado, formula, CAMPEDITAB, COLCAMECO, TXTTOTAL,DEJARLIBRE,FORMATO) "
            SQLX$ = SQLX$ + " VALUES ( " & IDLISTAX&
            SQLX$ = SQLX$ + " ,0"
            SQLX$ = SQLX$ + " ,0"
            SQLX$ = SQLX$ + " , 'DESCRIPCIÓN/A128'"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " , 'S'"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " , ''"
            SQLX$ = SQLX$ + " )"
            FLEXCONN.Execute (SQLX$)
  

        '
        End If
End Sub
Sub IntercaleFilasVacias(ByRef GRID As msFlexGrid)
    ' Obtenemos la posición de la fila seleccionada actualmente
    Dim currentRow As Long
    currentRow = GRID.row

    ' Desactivamos la redibujado automático para evitar un parpadeo
    GRID.Redraw = False

    ' Recorremos todas las filas de la grilla
    Dim i As Long
    For i = GRID.Rows - 1 To 0 Step -1
        ' Agregamos una nueva fila vacía después de la fila actual
        GRID.Rows = GRID.Rows + 1
        GRID.row = i + 1

        ' Desplazamos las filas hacia abajo para hacer espacio para la nueva fila vacía
        Dim J As Long
        For J = GRID.Rows - 1 To i + 2 Step -1
            GRID.row = J
            For c = 0 To GRID.Cols - 1
               GRID.TextMatrix(J, c) = GRID.TextMatrix(J - 1, c)
            Next c
            ' ... Repetir para cada columna de la grilla que se quiera conservar
        Next J

        ' Establecemos los valores de la nueva fila vacía
        If GRID.Rows - 1 < (i + 2) Then
        GRID.Rows = 1 + i + 2
        End If
        GRID.row = i + 2
        Dim k As Long
        For k = 0 To GRID.Cols - 1
            GRID.TextMatrix(GRID.row, k) = ""
        Next k
    Next i

    ' Restauramos la posición de la fila seleccionada original
    GRID.row = currentRow

    ' Reactivamos el redibujado automático
    GRID.Redraw = True
End Sub

Sub EliminarFilasVacias(ByRef GRID As msFlexGrid)
    ' Obtenemos la posición de la fila seleccionada actualmente
    Dim currentRow As Long
    currentRow = GRID.row

    ' Desactivamos la redibujado automático para evitar un parpadeo
    GRID.Redraw = False

    ' Recorremos todas las filas de la grilla
    Dim i As Long
    For i = GRID.Rows - 1 To 0 Step -1
        Dim isEmpty As Boolean
        isEmpty = True

        ' Verificamos si la fila actual está vacía
        Dim J As Long
        For J = 0 To GRID.Cols - 1
            If Len(GRID.TextMatrix(i, J)) > 0 Then
                isEmpty = False
                Exit For
            End If
        Next J

        ' Si la fila está vacía, la eliminamos
        If isEmpty Then
            GRID.RemoveItem i
        End If
    Next i

    ' Restauramos la posición de la fila seleccionada original
    If currentRow >= GRID.Rows Then
        currentRow = GRID.Rows - 1
    End If
    GRID.row = currentRow

    ' Reactivamos el redibujado automático
    GRID.Redraw = True
End Sub
Sub POSICIONAR_FILA(MSF As msFlexGrid, Fila As Long)
     With MSF
        If .Rows - 1 > 1 Then
            For i = Fila To 2 Step -1
                .TopRow = i
                If .RowIsVisible(Fila) = False Then
                    .TopRow = i + 2
                    .row = Fila: .col = 0
                    .RowSel = Fila: .ColSel = .Cols - 1
                    .Refresh
                    Encontro% = 1
                    Exit For
                End If
            Next i
            If Encontro% = 0 Then
                .TopRow = 1
                .row = 1: .col = 0
                .RowSel = 1: .ColSel = .Cols - 1
                .Refresh
            End If
        End If
     End With
End Sub
Sub CARGA_GRILLA_SHOWREP(Form1 As Form, MSF As MSHFlexGrid, Campos$, VALORES$, INICIALIZO%, Optional Negrita%)
   If INICIALIZO% < 1 Then
      Call ENCABEZADO_GENREP(MSF, Campos$, Form1)
      INICIALIZO% = 1
   End If
   
   If TRIM$(VALORES$) = "" Then Exit Sub
   MSF.Rows = MSF.Rows + 1
   Call TEXTOLOTES$(VALORES$, "|")
   For J& = 1 To UBound(CADENATEX$)
      MSF.TextMatrix(MSF.Rows - 1, J& - 1) = CADENATEX$(J&)
      MSF.row = MSF.Rows - 1
      MSF.col = J& - 1
      If Negrita% > 0 Then MSF.CellFontBold = True
   Next J&
End Sub
Sub ENCABEZADO_GENREP(MSF1 As MSHFlexGrid, LISTADECAMPOS$, Form1 As Form)
    '
    
    Dim NOCAMPO$(), FORMX$()
    ReDim NOCAMPO$(1), FORMX$(1)
    '
    ANCHO_TOTAL = 0
    On Error Resume Next
    If Form1.Label8.Count > 1 Then
      Erase Form1.Label8
    End If
    NUPIC% = 0
    MODOSE$ = UCase$(SAFETEXT$(MODOSEL$))
    LSTX$ = TRIM$(LISTADECAMPOS$)
    CANCA% = 0
    ATOTA% = 1
    While LSTX$ <> ""
       PPXX% = InStr(LSTX$, ";")
       PPXX2% = InStr(LSTX$, ",")
       If PPXX2% > 0 And PPXX2% < PPXX% Then PPXX% = PPXX2%
       If PPXX% < 1 Then PPXX% = 1 + Len(LSTX$)
       '
       CANCA% = CANCA% + 1
       If CANCA% >= UBound(NOCAMPO$) Then
          ReDim Preserve NOCAMPO$(CANCA%), FORMX$(CANCA%)
       End If
       NOCX$ = Left$(LSTX$, PPXX% - 1)
       LSTX$ = Mid$(LSTX$, PPXX% + 1)
       '
       PPXX% = InStr(NOCX$, "/")
       If PPXX% < 1 Then PPXX% = 1 + Len(NOCX$)
       NOCAMPO$(CANCA%) = Left$(NOCX$, PPXX% - 1)
       FORMX$(CANCA%) = Mid$(NOCX$, PPXX% + 1)
       '

       If FORMX$(CANCA%) = "" Then FORMX$(CANCA%) = "A16"
       '
       ATOTA% = ATOTA% + XVALO(Mid$(FORMX$(CANCA%), 2)) + 2
       '
       
       If CANCA% = 1 Then ACOL1% = XVALO(Mid$(FORMX$(CANCA%), 2))
       '
    Wend

    MSF1.Cols = CANCA%
    MSF1.Rows = 1
    For i% = 1 To CANCA%
       'PRIMER COLUMNA ASIGNA POR DEFECTO ANCHO = 2
       If i% = 0 Then
         If UCase$(Form1.Name) = "SHOWREP07" Then GoTo 80

         If ANCHO_COLUMNA_0% > 0 Then
           ANCHO% = ANCHO_COLUMNA_0% 'SI  SE PASA EL ANCHO DE LA COLUMNA 0 COMO PARAMETRO TOMA ESTE
         Else
           ANCHO% = 3  'EN CASO CONTRARIO ASIGNA VALOR POR DEFECTO
         End If
         NUCOLU% = i%:   Form1.TEPRUEBA = String$(ANCHO%, "A"): MSF1.ColWidth(NUCOLU%) = Form1.TEPRUEBA.Width: MSF1.ColAlignment(NUCOLU%) = ALINEA%: MSF1.TextMatrix(0, NUCOLU%) = NOCAMPO$(i%)
         GoTo 50
       End If
       
30     LETRX$ = "A"
       ANCHO% = XVALO(Mid$(FORMX$(i%), 2))
       If XVALO(Mid$(FORMX$(i%), 2)) > 0 Then
          ANCHO% = ANCHO% + 2
       ElseIf LETRX$ = "Z" Then
            
       Else
          ANCHO% = 0 'COLUMNAS OCULTAS PASADAS CON F0.0
       End If
       LETR$ = UCase$(Left$(FORMX$(i%), 1))
       ALINEA% = 1
       
       If LETR$ = "I" Or LETR$ = "G" Or LETR$ = "F" Then
          ALINEA% = 6
       ElseIf LETR$ = "Z" Then
          ALINEA% = 3
       End If
       '
       If ANCHO% > 0 Then
          NUCOLU% = i% - 1: Form1.TEPRUEBA = String$(ANCHO%, "A"): MSF1.ColWidth(NUCOLU%) = Form1.TEPRUEBA.Width: MSF1.ColAlignment(NUCOLU%) = ALINEA%: MSF1.TextMatrix(0, NUCOLU%) = NOCAMPO$(i%)
       Else ' PARA QUE NO SE VEAN LAS COLUMNAS OCULTAS
          NUCOLU% = i% - 1: Form1.TEPRUEBA = String$(ANCHO%, "A"): MSF1.ColWidth(NUCOLU%) = 0: MSF1.ColAlignment(NUCOLU%) = ALINEA%: MSF1.TextMatrix(0, NUCOLU%) = NOCAMPO$(i%)
       End If
       FORCAMPO$(NUCOLU% + 1) = LETR$
       FORCOLUM$(NUCOLU% + 1) = FORMX$(NUCOLU% + 1)
       MSF1.col = NUCOLU% - 1
       MSF1.CellFontBold = True
       
       ANCHO_TOTAL = ANCHO_TOTAL + Form1.TEPRUEBA.Width
       
       'A PARTIR DE AQUI CODIGO PARA CARGAR ENCABEZADO DE GENREP07

50     If UCase$(Form1.Name) = "SHOWREP07" Then
          Form1.Label5.FontName = Form1.GRID.CellFontName
          Form1.Label5.FontSize = Form1.GRID.CellFontSize
          Form1.Label5.FontBold = Form1.GRID.CellFontBold
          ' Label5.Caption = TEXTBIN$("1.234567", FORMAECO$(I&)) + "*"
          Form1.Label5.Caption = String$(ANCHO%, "A")    ' PROVISORIAMENTE A VER SI FUNCIONA
          Form1.GRID.col = i%
          Form1.GRID.ColWidth(i% - 1) = SHOWREP07.Label5.Width
25        If i% >= Form1.Label8.Count Then
             NUPIC% = Form1.Label8.Count
             Load Form1.Label8(NUPIC%)
             GoTo 25
          End If
          Form1.Label8(i%) = Str$(Len(Form1.Label5) - 1)
          If i% = 0 Then Form1.Label8(i%) = 0
       End If
80  Next i%
    '
    If NUPIC% = i% - 1 Then
      If UCase$(Form1.Name) = "SHOWREP07" Then Load Form1.Label8(NUPIC% + 1): Load Form1.Label8(NUPIC% + 2)
    End If
    
End Sub

Sub OrdenarListBox(ByRef lstBox As ListBox)
    Dim i As Integer, J As Integer
    Dim temp As Variant
    
    ' Implementación del algoritmo de ordenamiento de burbuja
    For i = 0 To lstBox.ListCount - 2
        For J = i + 1 To lstBox.ListCount - 1
            If lstBox.List(i) > lstBox.List(J) Then
                ' Intercambiar los elementos si están desordenados
                temp = lstBox.List(i)
                lstBox.List(i) = lstBox.List(J)
                lstBox.List(J) = temp
            End If
        Next J
    Next i
End Sub

Function SUMA_HORIZONTAL(MSF As msFlexGrid, REG&, DESDE_COL%, HASTA_COL%)
    SUMA_HORIZONTAL = 0
    'considerar que si es solo una columna desde y hasta va a ser la misma. si son 2 desde va a ser 1 y  hasta 2
    For col% = DESDE_COL% To HASTA_COL%
        SUMA_HORIZONTAL = SUMA_HORIZONTAL + XVALO(MSF.TextMatrix(REG&, col%))
    Next col%
End Function
 Sub OrdenarGrid(ByRef flexGrid As msFlexGrid, ByVal COLUMNA As Integer, Optional A_PartirDeRegistro&)
    Dim i As Integer
    Dim J As Integer
    Dim temp As String
    If A_PartirDeRegistro& > 0 Then
       Desde& = A_PartirDeRegistro&
    Else
       Desde& = 0
    End If

    For i = Desde& To flexGrid.Rows - 2
        For J = i + 1 To flexGrid.Rows - 1
            If flexGrid.TextMatrix(J, COLUMNA) < flexGrid.TextMatrix(i, COLUMNA) Then
                ' Intercambiar filas
                For col = 0 To flexGrid.Cols - 1
                    temp = flexGrid.TextMatrix(i, col)
                    flexGrid.TextMatrix(i, col) = flexGrid.TextMatrix(J, col)
                    flexGrid.TextMatrix(J, col) = temp
                Next col
            End If
        Next J
    Next i
End Sub
Sub AgregarFilaVaciaCuandoCambiaElDato(MSF As msFlexGrid, col%, Optional A_PartirDeFila%)
    Dim filaActual As Integer
    Dim valorActual As String
    Dim valorAnterior As String
    Dim filaInsertar As Integer
    If MSF.Rows <= 2 Then Exit Sub
    ' Inicializar valores iniciales
    If A_PartirDeFila% > 0 Then
      filaActual = A_PartirDeFila%
    Else
      filaActual = 1
    End If
    valorAnterior = MSF.TextMatrix(filaActual, col%)
    
    ' Recorrer las filas del MSFlexGrid
    Do While filaActual <= MSF.Rows - 1
        valorActual = MSF.TextMatrix(filaActual, col%)
        
        If valorActual <> valorAnterior Then
            filaInsertar = filaActual
            ' Insertar fila vacía
            MSF.Rows = MSF.Rows + 1
            ' Mover filas hacia abajo para abrir espacio
            For i = MSF.Rows - 1 To filaInsertar + 1 Step -1
                For J = 0 To MSF.Cols - 1
                    MSF.TextMatrix(i, J) = MSF.TextMatrix(i - 1, J)
                Next J
            Next i
            ' Limpiar la fila insertada
            For J = 0 To MSF.Cols - 1
                MSF.TextMatrix(filaInsertar, J) = ""
            Next J
        End If
        
        valorAnterior = valorActual
        filaActual = filaActual + 1
    Loop

End Sub
 Sub AdaptarAltoDeFilaSegunTexto(LBL As Label, MSF As msFlexGrid, ByVal row As Long, ByVal col As Long)
'    CODIGO GENERADO EN FORM DESPACBRRA EN PROYECTO PEDIDO07
'
'    HAY QUE PONER UN LABEL EN EL FORMULARIO Y PASAR ESTE COMO PARAMETRO COMO EN EL EJ. LABEL22 (AUTOZIZE = TRUE,WORDWRAP = TRUE)
'    EL LABEL DEBE TENER LA MISMA FUENTE Y TAMAÑO QUE LA GRILLA,
'    UTILIZAR EL FRATEXTO POR QUE ACOMODA CON GOLPES DE ENTER PARA QUE QUEDE COMO MULTILINEA
'    Call FRATEXTO("LISANDRO DE LA TORRE 3343 DEPARTAMENTO A EN PLANTA BAJA EN EL BARRIO DE RECOLETA ", (MSF.ColWidth(1) / 100))
'    For JJ& = 1 To CARETEX%
'      TXTX$ = TXTX$ + RETEX$(JJ&) + vbCrLf
'    Next JJ&
'    Label22 = TXTX$
'    MSF.TextMatrix(MSF.Row, 1) = TXTX$
'    Label22.Width = MSF.ColWidth(1)
'    Call AdaptarAltoDeFilaSegunTexto(Me.Label22, Me.MSF, MSF.Row, 1)

    tamañoOriginal = LBL.Font.Size
    tamañoSegunGrilla = MSF.Font.Size
    LBL.Font.Size = tamañoSegunGrilla
    ' Capturar el texto ingresado en la celda
    Dim TEXTO As String
    TEXTO = MSF.TextMatrix(row, col)
    
    ' Establecer el ancho actual de la columna
    Dim anchoColumna As Single
    anchoColumna = MSF.ColWidth(col)
    
    ' Establecer el alto actual de la fila
    Dim altoFila As Single
    altoFilaActual = MSF.RowHeight(row)
    altoFilaNecesario = LBL.Height
    
    'NUNCA VA A MENOS SI ES MENOR NO ENTRA
    If altoFilaActual < altoFilaNecesario Then
       
       MSF.RowHeight(row) = altoFilaNecesario - (altoFilaNecesario * 0.1)
    End If
    

End Sub



 Sub PintarFilasRepetidas(ByVal GRID As msFlexGrid, ByVal COLUMNA As Integer)
    'recorre toda la grilla, y pinta y si alguno se repite lo pinta de amarillo, no pinta el original solo la/s repeticion/es
    Dim dicRepetidos As Object
    Dim dicLeidos As Object
    Set dicRepetidos = CreateObject("Scripting.Dictionary")
    Set dicLeidos = CreateObject("Scripting.Dictionary")
    
    Dim Fila As Integer
    Dim valor As String
    'en dicLeidos va guardando el dato de la celda corresondiente a la columna leida
    'si ya existe en el diccionario lo mete dentro de la lista dicRepetidos
    'la llave va a ser el nro de pedido (valor) y le agrega la fila separado por coma
    For Fila = 1 To GRID.Rows - 1
        valor = GRID.TextMatrix(Fila, COLUMNA)
        If dicLeidos.Exists(valor) Then
            dicRepetidos(valor) = dicRepetidos(valor) & "," & Fila
        Else
            dicLeidos.Add valor, CStr(Fila)
        End If
    Next Fila
    
    Dim key As Variant
    For Each key In dicRepetidos
        Dim filas() As String
        filas = Split(dicRepetidos(key), ",")
        
        Dim i As Integer
        For i = 1 To UBound(filas)
            GRID.row = XVALO(filas(i))
            GRID.col = 0
            GRID.ColSel = GRID.Cols - 1
            Call COLOREAR_GRILLA_SOLAFILA(GRID, vbYellow, GRID.row, 0, 0, 1)      ' BLANQUEO LA FILA
'            GRID.CellBackColor = RGB(255, 0, 0) ' Pintar la fila de rojo
        Next i
    Next key
    
    Set dicRepetidos = Nothing
    Set dicLeidos = Nothing

End Sub
Function FilaMaximaConContenidoDentroDeUnRangodeColumnas&(RutaArc$, NombreHoja$, RangoCol$)
    'DEVUELVE DEL RANGO PARAMETRIZADO CUAL ES LA FILA MAXIMA CON CONTENIDO
    'EJ. SI SE LE PASA B:H RECORRE LAS COLUMAS B, C D,...H Y DE ESTAS SE QUEDA CON EL VALOR MAXIMO DE FILA CON CONTENIDO
    'RangoCol$ = A:J no debe pasarse de Z
    Dim XLAPP As Object
    Dim xlWorkbook As Object
    Dim xlWorksheet As Object
    'Dim lastRowArray(1 To 10) As Long ' Arreglo para almacenar los valores de la última fila con contenido
    Dim maxVal As Double
    
    ' Crear una instancia de Excel
    Set XLAPP = CreateObject("Excel.Application")
    
    ' Abrir el archivo de Excel
    Set xlWorkbook = XLAPP.Workbooks.Open(RutaArc$)
    
    ' Seleccionar la hoja de Excel
    Set xlWorksheet = xlWorkbook.Worksheets(NombreHoja$)
    
    ' Obtener la última fila con datos en cada columna del rango A a J
    valMin% = Asc(Left(RangoCol$, 1)) - 64
    valMax% = Asc(Right(RangoCol$, 1)) - 64
    Dim lastRowArray() As Long ' Arreglo para almacenar los valores de la última fila con contenido
    For i = valMin% To valMax%
        If i = valMin% Then ReDim Preserve lastRowArray(valMax%)
        lastRowArray(i) = xlWorksheet.Cells(xlWorksheet.Rows.Count, Chr(64 + i)).End(-4162).row
    Next i
    
    ' Obtener el valor máximo utilizando la función Max
    maxVal = Application.Max(lastRowArray)
    FilaMaximaConContenidoDentroDeUnRangodeColumnas& = maxVal
    
    ' Desactiva las alertas de Excel para evitar que pregunte al reemplazar el archivo
    XLAPP.DisplayAlerts = False

    ' Guardar y cerrar el archivo de Excel
    xlWorkbook.Save
    
    ' Activa las alertas de Excel
    XLAPP.DisplayAlerts = True
    
    xlWorkbook.Close
    
    ' Liberar la memoria utilizada por Excel
    Set xlWorksheet = Nothing
    Set xlWorkbook = Nothing
    XLAPP.Quit
    Set XLAPP = Nothing
End Function

Sub BorrarContenidoExcelXColumnas(RUTA_ARCH$, NombreHoja$, RangoDeColumnas$, aPartirDeFilaBorrarDatos%)
        'EJEMPLO DE COMO PASAR PARAMETRO DE RANGOS: "C:J" BORRARIA DE LA A1 HASTA LA COLUMNA J , TODOS LOS DATOS
        Dim XLAPP As Object
        Dim xlWorkbook As Object
        Dim xlWorksheet As Object
        Dim lastRow As Long
        
        ' Crear una instancia de Excel
        Set XLAPP = CreateObject("Excel.Application")
        ' Abrir el archivo de Excel
        Set xlWorkbook = XLAPP.Workbooks.Open(RUTA_ARCH$)
        
        ' Seleccionar la hoja de Excel
        Set xlWorksheet = xlWorkbook.Worksheets(NombreHoja$)
        
        ' Obtener la última fila con datos en la columna
        'lastRow = xlWorksheet.Cells(xlWorksheet.Rows.Count, "D").End(-4162).Row
        lastRow = FilaMaximaConContenidoDentroDeUnRangodeColumnas&(RUTA_ARCH$, NombreHoja$, RangoDeColumnas$)
        ' Limpiar el rango de columnas A a C hasta la última fila
        'xlWorksheet.Range("A1:C" & lastRow).ClearContents
        xlWorksheet.Range(RangoDeColumnas$ & lastRow).ClearContents
        
        
        ' Desactiva las alertas de Excel para evitar que pregunte al reemplazar el archivo
         XLAPP.DisplayAlerts = False
         
         ' Guardar y cerrar el archivo de Excel
         xlWorkbook.Save
         ' Activa las alertas de Excel nuevamente
         XLAPP.DisplayAlerts = True
        
        xlWorkbook.Close
        
        ' Liberar la memoria utilizada por Excel
        Set xlWorksheet = Nothing
        Set xlWorkbook = Nothing
        XLAPP.Quit
        Set XLAPP = Nothing
End Sub

 Sub ExportarAExcel_MSF_IA(flexGrid As msFlexGrid, RUTA_ARCH$, NombreHoja$, IMPRESORA$, filaInicio As Integer, columnaInicio As Integer, Optional areaImpresion$, Optional CantSumarFilasImpre%, Optional RANGOENCABEZADO$)
'    Dim excelApp As New Excel.Application
'    Dim excelWorkbook As Excel.Workbook
'    Dim excelWorksheet As Excel.Worksheet
'Ej. areaImpresion$ = "A1:L"
'EJ, RANGOENCABEZADO$ = "$1:$4" ESL DESDE QUE FILA HASTA QUE FILA SE VA A REPETIR LOS ENCABEZADOS AL CAMBIAR DE HOJA.
    Dim ObjetoExcel As Excel.Application
    Dim Libro As Excel.Workbook
    Dim Hoja As Excel.Worksheet
    
    ' Crea una nueva instancia de Excel y abre un libro
'    Set excelWorkbook = excelApp.Workbooks.Add
'    Set excelWorksheet = excelWorkbook.Worksheets(1)
    Set ObjetoExcel = CreateObject("Excel.Application")
    With ObjetoExcel
    Set Libro = ObjetoExcel.Workbooks.Open(RUTA_ARCH$)
    ' Copia los datos del MSFlexGrid al rango de celdas en Excel
    
    Dim J As Integer
    FIN& = flexGrid.Rows - 1
    For i& = 0 To flexGrid.Rows - 1
        Call TRACE(20, i&, FIN&)
        ColumnasOcultas = 0
        For J = 1 To flexGrid.Cols - 1
            anchoColumna = flexGrid.ColWidth(J)
            If Val(anchoColumna) < 1 Then
               ColumnasOcultas = ColumnasOcultas + 1
               GoTo 30 'para que las celdas ocultas no se impriman
            End If
'            excelWorksheet.Cells(i + filaInicio, j + columnaInicio).Value = flexGrid.TextMatrix(i, j)
'            excelWorksheet.Cells(i + filaInicio, j + columnaInicio).WrapText = True ' Establece WrapText en True para ajustar el alto de la fila
'            excelWorksheet.Cells(i + filaInicio, j + columnaInicio).HorizontalAlignment = Excel.Constants.xlCenter ' Alinea el texto en el centro horizontalmente
'            excelWorksheet.Cells(i + filaInicio, j + columnaInicio).VerticalAlignment = Excel.Constants.xlCenter ' Alinea el texto en el centro verticalmente
            Libro.Worksheets(NombreHoja).Cells(i + filaInicio, J - ColumnasOcultas + columnaInicio).Value = flexGrid.TextMatrix(i, J)
            Libro.Worksheets(NombreHoja).Cells(i + filaInicio, J - ColumnasOcultas + columnaInicio).WrapText = True ' Establece WrapText en True para ajustar el alto de la fila
            Libro.Worksheets(NombreHoja).Cells(i + filaInicio, J - ColumnasOcultas + columnaInicio).HorizontalAlignment = Excel.Constants.xlCenter ' Alinea el texto en el centro horizontalmente
            Libro.Worksheets(NombreHoja).Cells(i + filaInicio, J - ColumnasOcultas + columnaInicio).VerticalAlignment = Excel.Constants.xlCenter ' Alinea el texto en el centro verticalmente

30        Next J
    Next i&
'    Dim lastColumn As Long
'    lastColumn = Libro.Worksheets(NombreHoja).Cells(4, Libro.Worksheets(NombreHoja).Columns.Count).End(xlToLeft).Column
'
'    Libro.Worksheets(NombreHoja).PageSetup.PrintArea = "A1:" & ColumnLetter(lastColumn) & "1"
'    Libro.Worksheets(NombreHoja).PageSetup.FitToPagesWide = 1
'    Libro.Worksheets(NombreHoja).Columns.AutoFit
    'esto es para que marque el final del area de impresion', parea el caso que no se sabe el final p.e. una tabla c/n datos
    If areaImpresion$ <> "" Then
       'si ingresa aqui configura el area de impres, y si CantSumarFilasImpre% tiene valor le suma
       'para que amplie el area n filas mas.
       'Ej. areaImpresion$ = "A1:L"
       Libro.Worksheets(NombreHoja).PageSetup.PrintArea = areaImpresion$ & i + CantSumarFilasImpre%
       Set xlPageSetup = Libro.Worksheets(NombreHoja).PageSetup
       ' Establecer las filas que se deben repetir en cada página
       
       'xlPageSetup.PrintTitleRows = "$1:$4" ' Reemplaza 1 y 3 con los números de fila que deseas fijar
        xlPageSetup.PrintTitleRows = RANGOENCABEZADO$
    End If
    ' Desactiva las alertas de Excel para evitar que pregunte al reemplazar el archivo
         ObjetoExcel.DisplayAlerts = False
         ObjetoExcel.ActiveWorkbook.Save
         ' Activa las alertas de Excel nuevamente
         ObjetoExcel.DisplayAlerts = True
         Set Hoja = Libro.Sheets(NombreHoja$)
         If NOIMPRIMIR% > 0 Then GoTo 99
         On Error Resume Next
         Hoja.PrintOut Copies:=1, Collate:=True, ActivePrinter:=IMPRESORA$
         'HOJA.PrintOut Copies:=1, Collate:=True 'asi sale por la impresora predeterminada
        On Error GoTo 0
End With
    ' Guarda el libro de Excel y cierra la aplicación Excel
'    excelWorkbook.SaveAs "C:\Users\Omar\Desktop\Libro2.xlsx" ' Reemplaza "Ruta\Archivo.xlsx" con la ruta y el nombre de archivo deseados
'    excelWorkbook.Close
'    excelApp.Quit
'
'    ' Libera los objetos de Excel de la memoria
'    Set excelWorksheet = Nothing
'    Set excelWorkbook = Nothing
'    Set excelApp = Nothing
99  On Error Resume Next
    Set Hoja = Nothing
    Libro.Close: Set Libro = Nothing
    ObjetoExcel.Close: Set ObjetoExcel = Nothing
    On Error GoTo 0

End Sub


Function ColumnLetter(ColumnNumber As Long) As String
    Dim N As Long
    Dim c As Byte
    Dim s As String
    N = ColumnNumber
    Do
        c = ((N - 1) Mod 26)
        s = Chr(c + 65) & s
        N = (N - c) \ 26
    Loop While N > 0
    ColumnLetter = s
End Function

Sub APlicarEnBloqueRepetido(MSF As msFlexGrid, TEXTO$, NregImpreso&, ColumnaALeer%, ColumnaParaAplicar%, Optional COLOR)
  'esta funcion se usa en EYM DONDE EN UNA GRILLA EN UNA COLUMNA TIENEN ENUMERADOR PERO EN ALGUNOS CASOS
  'EL MISMO SE REPITE POR EJEMPLO VA DEL 1 AL 20 PERO EL 8 SE REPITE 2 VECES  Y LUEGO SALTA AL 9
  'Y LA NECESIDAD ES QUE EL RESULTADO SE REPITA ES DECIR SI EN  CUALQUIERA DE LOS REGISTRO REPETIDO VA UN RESULTADO
  'EL MISMO SE DEBE REPLICAR EN SU PAR
  'TEXTO$ = AL TEXTO QUE SE VA A IMPRIMIR
  'NregImpreso& VIENE COMO PARAMETRO QUE  INDICA CUAL ES EL NRO QUE  SE REPITE
  'ColumnaALeer% = ES DE DONDE SE VA A LEER Y COMPARAR EL NRO REPETIDO
  'ColumnaParaAplicar% = DONDE SE VA A INTRODUCIR EL TEXTO  CUANDO SE CUMPLA LA REPETICION DEL NregImpreso&
  For i& = 1 To MSF.Rows - 1
    DatoLeido& = XVALO(MSF.TextMatrix(i&, ColumnaALeer%))
    If DatoLeido& = NregImpreso& Then
       MSF.TextMatrix(i&, ColumnaParaAplicar%) = TEXTO$
       If COLOR <> "" Then
         Call COLOREAR_TEXTO_SOLO_UNA_AFILA(MSF, COLOR, i, ColumnaParaAplicar%)
       End If
    End If
  Next i&
End Sub
Sub ImprimirEnExcel(MSFlexGrid1 As msFlexGrid)
    Dim excelApp As Object
    Dim excelWorkbook As Object
    Dim excelWorksheet As Object
    Dim flexGrid As msFlexGrid
    Dim i As Integer
    Dim J As Integer
    
    ' Asignar el nombre de tu MSFlexGrid aquí
    Set flexGrid = MSFlexGrid1
    
    ' Crear una instancia de Excel
    Set excelApp = CreateObject("Excel.Application")
    excelApp.Visible = True
    
    ' Agregar un nuevo libro de trabajo
    Set excelWorkbook = excelApp.Workbooks.Add
    Set excelWorksheet = excelWorkbook.Worksheets(1)
    
    ' Copiar los datos desde el MSFlexGrid a Excel
    For i = 0 To flexGrid.Rows - 1
        For J = 0 To flexGrid.Cols - 1
            excelWorksheet.Cells(i + 1, J + 1).Value = flexGrid.TextMatrix(i, J)
            excelWorksheet.Cells(i + 1, J + 1).WrapText = True
        Next J
    Next i
    
    ' Ajustar automáticamente la altura de las celdas
    excelWorksheet.Cells.EntireRow.AutoFit
    
    ' Aplicar formatos a las celdas de fecha y números
    For i = 0 To flexGrid.Rows - 1
        For J = 0 To flexGrid.Cols - 1
            If IsDate(flexGrid.TextMatrix(i, J)) And InStr(flexGrid.TextMatrix(i, J), "/") > 0 Then
                excelWorksheet.Cells(i + 1, J + 1).NumberFormat = "dd/mm/yyyy"
                excelWorksheet.Cells(i + 1, J + 1).Value = CDate(flexGrid.TextMatrix(i, J))
            ElseIf IsNumeric(flexGrid.TextMatrix(i, J)) Then
                If InStr(flexGrid.TextMatrix(i, J), ".") > 0 Then
                  excelWorksheet.Cells(i + 1, J + 1).NumberFormat = "#,##0.00"
'                Else
'                  excelWorksheet.Cells(i + 1, J + 1).NumberFormat = "#,##0"
                End If
            End If
        Next J
    Next i
    
   With excelWorkbook.Sheets("Hoja1") ' Reemplaza "NombreDeTuHoja" con el nombre de tu hoja en Excel
    .Rows(1).Interior.COLOR = RGB(128, 128, 128)  ' Cambia el color de fondo de la primera fila a amarillo (puedes ajustar los valores RGB según tu preferencia)
   End With




    
    ' Liberar recursos de Excel
    Set excelWorksheet = Nothing
    Set excelWorkbook = Nothing
    Set excelApp = Nothing
End Sub
