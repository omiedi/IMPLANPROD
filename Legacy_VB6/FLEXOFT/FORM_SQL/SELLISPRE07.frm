VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form SELLISPRE07 
   BackColor       =   &H00CDDADA&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Listado de Precios de Venta vs Costos"
   ClientHeight    =   3735
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4665
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3735
   ScaleWidth      =   4665
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox MARCOBARRA 
      Appearance      =   0  'Flat
      BackColor       =   &H002B9973&
      ForeColor       =   &H80000008&
      Height          =   200
      Left            =   120
      ScaleHeight     =   165
      ScaleWidth      =   2655
      TabIndex        =   3
      Top             =   3360
      Width           =   2685
      Begin VB.Frame BARRATRAZA 
         BackColor       =   &H0048CC9D&
         BorderStyle     =   0  'None
         Height          =   500
         Left            =   0
         TabIndex        =   4
         Top             =   0
         Width           =   12000
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   3720
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Generar Reporte"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Left            =   2960
      TabIndex        =   2
      Top             =   3240
      Width           =   1600
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2580
      ItemData        =   "SELLISPRE07.frx":0000
      Left            =   120
      List            =   "SELLISPRE07.frx":0002
      Style           =   1  'Checkbox
      TabIndex        =   0
      Top             =   480
      Width           =   4455
   End
   Begin VB.Label Label1 
      BackColor       =   &H00CDDADA&
      Caption         =   "Seleccione las Listas de Precios:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   160
      Width           =   3375
   End
End
Attribute VB_Name = "SELLISPRE07"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command8_Click()
    '
    Dim XLAPP As Excel.Application
    Dim XLBOOK As Excel.Workbook
    Dim XLSHEET As Excel.Worksheet
    Dim listaselec(255) As Integer
    '
    ' VERIFICA QUE HAYAN LISTAS ELEGIDAS Y LAS GUARDA PARA MAS ADELANTE
    cantlistas& = 0
    For i& = 0 To List1.ListCount - 1
        If List1.Selected(i&) Then
            cantlistas& = cantlistas& + 1
            listaselec(cantlistas&) = i& + 1
        End If
    Next i&
    '
    If cantlistas& = 0 Then
        Call COMUNI("Imposible Continuar. Ninguna de las Listas Fue Seleccionada.")
        Exit Sub
    End If
    '
    ' ARMA CONDICION PARA LA CONSULTA SQL PARA LA OBTENCION DE LOS CODIGOS DEL REPORTE
    CONDI$ = ""
    For i& = 1 To cantlistas&
        nroli& = listaselec(i&)
        CONDI$ = CONDI$ & "PREVENTA.COD_LISTA=" & Str$(nroli&) & " OR "
    Next i&
    CONDI$ = Mid$(CONDI$, 1, Len(CONDI$) - 4)
    ' OBTIENE LOS CODIGOS DE LAS LISTAS DE PRECIOS PARA ARMADO DEL REPORTE,
    ' UTILIZANDO EL DISTINCT PARA QUE NO HAYAN REPETICIONES
    'SQLX$ = "SELECT DISTINCT COD_PIEZA FROM PREVENTA WITH(NOLOCK) WHERE " & CONDI$
    SQLX$ = "SELECT DISTINCT PREVENTA.COD_PIEZA, MASTER.CODIGO, MASTER.DESCRIPCION FROM PREVENTA WITH(NOLOCK) INNER JOIN MASTER WITH(NOLOCK) ON PREVENTA.COD_PIEZA=MASTER.CODINT WHERE " & CONDI$ & " ORDER BY MASTER.CODIGO"
    Set RST = READSET(SQLX$)
    '
    'SQLX$ = "SELECT COUNT (DISTINCT COD_PIEZA) FROM PREVENTA WITH(NOLOCK) WHERE " & CONDI$
    SQLX$ = "SELECT COUNT (DISTINCT COD_PIEZA) FROM PREVENTA WITH(NOLOCK) WHERE " & CONDI$
    Dim cantidadregistros As ADODB.Recordset
    Set cantidadregistros = READSET(SQLX$)
    cantireg& = cantidadregistros.Fields(0)
    '
    With RST
        If .EOF And .BOF Then
            Call COMUNI("Las Listas de Precio Seleccionadas No Poseen Información")
            Exit Sub
        End If
        
        'ELIGE UBICACION PARA EL ARCHIVO EXCEL
        CommonDialog1.CancelError = True
        CommonDialog1.DialogTitle = "Exportar a Excel"
        CommonDialog1.FILTER = "Excel Files(*.XLS)|*.XLS|"
        CommonDialog1.FilterIndex = 1
        On Error GoTo 90
        CommonDialog1.ShowOpen
        On Error GoTo 0
        NOMBREDEARCHIVO$ = UCase$(CommonDialog1.FileName)
        If TRIM$(NOMBREDEARCHIVO$) = "" Then
            Call COMUNI("No Ha Elegido una Ubicación o la Misma es Incorrecta")
            Exit Sub
        End If
        '
        On Error GoTo 91
        Set XLAPP = New Excel.Application
        Set XLBOOK = XLAPP.Workbooks.Add
        Set XLSHEET = XLBOOK.Worksheets.Add
        On Error GoTo 0
        '
        'ARMA ENCABEZADO
        XLSHEET.Cells(1, 1) = "Código"
        XLSHEET.Cells(1, 2) = "Descripción"
        XLSHEET.Cells(1, 3) = "Costo"
        COLUMNAX1& = 3
        For i& = 1 To cantlistas&
            COLUMNAX1& = COLUMNAX1& + 1
            XLSHEET.Cells(1, COLUMNAX1&) = "Lista " & TRIM$(Str$(listaselec(i&)))
'            COLUMNAX1& = COLUMNAX1& + 1
'            xlsheet.Cells(1, COLUMNAX1&) = "Margen $"
            COLUMNAX1& = COLUMNAX1& + 1
            XLSHEET.Cells(1, COLUMNAX1&) = "Margen %"
        Next i&
        '
        Screen.MousePointer = 11
        FILA& = 1: FIN& = cantireg&: II& = 0
        Do While Not .EOF
            II& = II& + 1
            Call TRACE(24, II&, FIN&)
            CIXI% = !COD_PIEZA
            FILA& = FILA& + 1
            XLSHEET.Cells(FILA&, 1) = "'" & CODEXT$(CIXI%)
            XLSHEET.Cells(FILA&, 2) = "'" & DESCRIT0$(CIXI%)
                'Call ANACOSTO07.CARCOSTO_NUEVO(CIXI%)
                'costomp# = XVALO(ANACOSTO07.Label85.Caption)
                YXY$ = FILTERGETRECORD$("RECOSCAL", 1, MKI$(CIXI%))
                costomp# = CVD(Mid$(YXY$, 33, 8))
            XLSHEET.Cells(FILA&, 3) = ROUND#(costomp#, 4)
            COLUMNAX1& = 3
            For i& = 1 To cantlistas&
                'listas$ = TRIM$(Str$(listaselec(I&)))
                'Do While Len(listas$) < 3: listas$ = "0" & listas$: Loop
                'precio# = PRELISTA#("LIPRE" & listas$, CIXI%)
                precio# = XVALO(VALOBADA("preventa", "precio_unitario", "cod_lista=" & listaselec(i&) & " and cod_pieza=" & CStr(CIXI%)))
                    COLUMNAX1& = COLUMNAX1& + 1
                If precio# > 0 Then
                    XLSHEET.Cells(FILA&, COLUMNAX1&) = ROUND#(precio#, 4)
'                        COLUMNAX1& = COLUMNAX1& + 1
'                    xlsheet.Cells(FILA&, COLUMNAX1&) = ROUND#(precio# - costomp#, 4)
                        COLUMNAX1& = COLUMNAX1& + 1
                    XLSHEET.Cells(FILA&, COLUMNAX1&) = ROUND#(((precio# - costomp#) / precio#) * 100)
                Else
                    COLUMNAX1& = COLUMNAX1& + 1
                End If
            Next i&
            .MoveNext
        Loop
    End With
    On Error GoTo 92
    XLSHEET.SaveAs NOMBREDEARCHIVO$
    XLAPP.Visible = False
    XLAPP.Quit      ' cierra EXCEL que se esta ejecutando minimizado
    Set XLAPP = Nothing     ') libera memoria
    Set XLBOOK = Nothing    ') utilizada por
    Set XLSHEET = Nothing   ') las variables objeto
    On Error GoTo 0
    '
    Screen.MousePointer = 1
    Call COMUNI("Planilla Excel Creada Exitosamente:\ \" + REPLACAR$(NOMBREDEARCHIVO$, "\", "/"))
    '
    Exit Sub
    '
90: Call COMUNI("Error al Detectar la Ubicación Elegida.\Intente Nuevamente o con Otra Ubicación")
    Exit Sub
91: Call COMUNI("Error al Crear Archivo Excel")
    Exit Sub
92: Call COMUNI("Error al Grabar Archivo Excel")
    Exit Sub
    '
End Sub

Private Sub Form_Load()
    '
    ' CARGA LISTA PARA EL ARMADO DEL REPORTE
    '
    For i& = 1 To ULTREG("LISTAS")
        XX$ = REGLEIDO$("LISTAS", i&)
        NROLISTA% = Asc(Mid$(XX$, 1, 1))
        DESCRIPLISTA$ = TRIM$(Mid$(XX$, 2, 46))
        '
        NROLIS$ = Str$(NROLISTA%)
        Do While Len(NROLIS$) < 3: NROLIS$ = " " & NROLIS$: Loop
        TX$ = NROLIS$ & " - " & DESCRIPLISTA$
        '
        List1.AddItem TX$
    Next i&
    '
End Sub
