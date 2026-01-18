VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form CARHORAR 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   9225
   ClientLeft      =   45
   ClientTop       =   585
   ClientWidth     =   15675
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9225
   ScaleWidth      =   15675
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command4 
      Caption         =   "Carga Horaria de Producción"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   765
      Left            =   13680
      Picture         =   "CARHORAR.frx":0000
      TabIndex        =   9
      ToolTipText     =   "Cargar Detalle de Piezas y Carga Horaria de Producción"
      Top             =   2000
      Width           =   1905
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Exportar a Excel"
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
      Left            =   11760
      TabIndex        =   5
      Top             =   8400
      Visible         =   0   'False
      Width           =   1600
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   7680
      Top             =   3960
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox Picture999 
      Height          =   225
      Index           =   0
      Left            =   7680
      ScaleHeight     =   165
      ScaleWidth      =   990
      TabIndex        =   0
      Top             =   7200
      Visible         =   0   'False
      Width           =   1050
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   2625
      OleObjectBlob   =   "CARHORAR.frx":030A
      Top             =   0
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   225
      Index           =   0
      Left            =   120
      OleObjectBlob   =   "CARHORAR.frx":053E
      TabIndex        =   1
      Top             =   240
      Visible         =   0   'False
      Width           =   1485
   End
   Begin MSFlexGridLib.MSFlexGrid MSF1 
      Height          =   3375
      Left            =   120
      TabIndex        =   6
      Top             =   5760
      Width           =   15555
      _ExtentX        =   27437
      _ExtentY        =   5953
      _Version        =   393216
      Cols            =   4
      BackColor       =   16777215
      BackColorFixed  =   16501405
      BackColorBkg    =   16777215
      WordWrap        =   -1  'True
      FocusRect       =   2
      AllowUserResizing=   1
      FormatString    =   "|>CAMPO1|>CAMPO2|>CAMPO3"
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
   Begin MSFlexGridLib.MSFlexGrid MSF0 
      Height          =   2415
      Left            =   120
      TabIndex        =   7
      ToolTipText     =   "Click Sobre Columna 0 Para Seleccionar Proyecto"
      Top             =   360
      Width           =   13515
      _ExtentX        =   23839
      _ExtentY        =   4260
      _Version        =   393216
      Cols            =   4
      BackColor       =   16777215
      BackColorFixed  =   16501405
      BackColorBkg    =   16777215
      WordWrap        =   -1  'True
      FocusRect       =   2
      AllowUserResizing=   1
      FormatString    =   "|>CAMPO1|>CAMPO2|>CAMPO3"
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
   Begin MSFlexGridLib.MSFlexGrid MSF2 
      Height          =   2295
      Left            =   120
      TabIndex        =   8
      Top             =   3120
      Width           =   15555
      _ExtentX        =   27437
      _ExtentY        =   4048
      _Version        =   393216
      Cols            =   4
      BackColor       =   16777215
      BackColorFixed  =   16501405
      BackColorBkg    =   16777215
      WordWrap        =   -1  'True
      FocusRect       =   2
      AllowUserResizing=   1
      FormatString    =   "|>CAMPO1|>CAMPO2|>CAMPO3"
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
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "LISTA DE PROYECTOS"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   12
      Top             =   0
      Width           =   15615
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "DETALLE DE PIEZAS (Según Proyectos Seleccionados)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   11
      Top             =   2760
      Width           =   15375
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "CARGA HORARIA DE  PRODUCCIÓN"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   10
      Top             =   5400
      Width           =   15375
   End
   Begin VB.Label Label7 
      BorderStyle     =   1  'Fixed Single
      Height          =   225
      Left            =   11040
      TabIndex        =   4
      Top             =   1080
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Label5"
      Height          =   240
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Label TEPRUEBA 
      AutoSize        =   -1  'True
      Caption         =   "TEPRUEBA"
      Height          =   300
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Menu mnuarchivo 
      Caption         =   "Archivo"
      Begin VB.Menu exportarExcel 
         Caption         =   "Exportar a Excel"
      End
   End
End
Attribute VB_Name = "CARHORAR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim FrmLeft, FrmTop, FrmWidth, FrmHeight As Long
Dim MSF1Top, MSF1Left, MSF1Width, MSF1Height As Long
Dim FRAMETop, FRAMELeft, FRAMEWidth, FRAMEHeight As Long
Dim Com1Top, Com2Top, Com2Left As Long
Dim Com1LEFT As Long

Dim LecturaDatosOriginales As Byte


Sub AGREGAR_PROMEDIO(MSF As MSFlexGrid, HORIZ_1_VERT_2%, APARTIRDE&)
   If HORIZ_1_VERT_2% = 1 Then
      COLUM& = MSF.Cols
      MSF.Cols = MSF.Cols + 1
      For j& = APARTIRDE& To MSF.Rows - 1
        For I& = 1 To MSF.COL - 1
         VALOR = XVALO(MSF.TextMatrix(j&, I&))
        Next I&
        MSF.TextMatrix(j&, COLUM&) = VALOR / (I& - 1)
      Next j&
   Else
      REG& = MSF.Rows
      MSF.Rows = MSF.Rows + 1
      For I& = APARTIRDE& To MSF.COL - 1
        For j& = 1 To MSF.Rows - 1
         VALOR = XVALO(MSF.TextMatrix(I&, j&))
        Next j&
        MSF.TextMatrix(REG&, I&) = VALOR / (j& - 1)
      Next I&

   End If

End Sub

Private Sub BOTEXIT_Click()
  Unload Me
End Sub

Private Sub CMBFILTRO_Change()
   TXTBUSCAR_Change
End Sub

Private Sub Command1_Click()
'   Command1.Enabled = False
'   '
'
'   CommonDialog1.CancelError = True
'   CommonDialog1.DialogTitle = "Exportar a Excel"
'       ' cancelado, si no da error - CommonDialog1.Flags = &H2& + &H4& + &H8 + &H800 + &H2000 + &H8000 'cdlOFNCreatePrompt
'   CommonDialog1.FILTER = "Excel Files(*.XLS)|*.XLS|"
'   CommonDialog1.FilterIndex = 1
'   On Error GoTo 99
'   CommonDialog1.ShowOpen
'   On Error GoTo 0
'   '
'   DoEvents
'   Screen.MousePointer = 11
'   AREXPO$ = UCase$(CommonDialog1.FileName)
'   If TRIM$(AREXPO$) <> "" Then
'     If MSF_2.Visible = True Then
'      Call GENEXCEL(AREXPO$, MUESTRADATOS.MSF_2)
'     Else
'      Call GENEXCEL(AREXPO$, MUESTRADATOS.MSF1)
'     End If
'   End If
'   '
'
'99 Command1.Enabled = True

End Sub

Sub GENEXCEL_PRUEBA_GRAFICO(NOMBREDEARCHIVO$, MSF1 As MSFlexGrid)
   Dim XLAPP As Excel.Application
   Dim XLBOOK As Excel.Workbook
   Dim XLSHEET As Excel.Worksheet
   Dim oChart As Object     ' Excel Chart

   '
   Dim FECHA As Date
   Dim VALOR As Double
   Dim TEXTO As String
   '
   On Error GoTo 20
   Set XLAPP = New Excel.Application
   Set XLBOOK = XLAPP.Workbooks.Add
   Set XLSHEET = XLBOOK.Worksheets.Add
   On Error GoTo 0
   '
   TIPODATO% = 0: COLUM& = 0
   ReDim aTemp(MSF1.Rows To MSF1.Rows, 3 To MSF1.Cols - 1)

   For I& = 1 To MSF1.Rows
       RENG1& = I&: COLUM& = 0
       For COLU1& = 1 To MSF1.Cols
          If MSF1.ColWidth(COLU1& - 1) = 0 Then GoTo 10 ' OBVIA LAS COLUMNAS OCULTAS
          COLUM& = COLUM& + 1
          '
          VALORX = MSF1.TextMatrix(I& - 1, COLU1 - 1)
          If Len(VALORX) = 8 And Mid$(VALORX, 3, 1) = "/" And Mid$(VALORX, 6, 1) = "/" Then
             FEXU$ = VALORX
             VALORX = FETEXCOR1$(Int(FECHANUM(FEXU$)))
             TIPODATO% = 1
          End If
          ' ESTO ES PARA QUE EXPORTE COMO CERO Y NO ESPACIOS
          FFCC$ = UCase$(Left$(FORCAMPO$(COLU1&), 1))
          If FFCC$ = "I" Or FFCC$ = "F" Or FFCC$ = "G" Or FFCC$ = "$" Then
             If TRIM$(CStr(VALORX)) = "" Then
                VALORX = 0
             End If
             TIPODATO% = 2
          End If
          If FFCC$ = "C" Or FFCC$ = "A" Then VALORX = "'" & VALORX
          '

          XLSHEET.Cells(RENG1&, COLUM&) = VALORX
          If RENG1& = MSF1.Rows And COLUM& > 3 And COLUM& < MSF1.Cols Then
             aTemp(6, COLUM&) = VALORX
          End If
          If TIPODATO% = 1 Then XLSHEET.Cells(RENG1&, COLUM&).NumberFormat = "dd/mm/yyyy"
          'If TIPODATO% = 2 Then XLSHEET.Cells(RENG1&, COLU1&).NumberFormat = "#,##0.00"
          If TIPODATO% = 0 Then XLSHEET.Cells(RENG1&, COLUM&).NumberFormat = "@"
          TIPODATO% = 0
          VALORX_ANT = VALORX
10     Next COLU1&
   Next I&
   '
   
    Dim iRow As Integer      ' Index variable for the current Row
    Dim iCol As Integer      ' Index variable for the current Row
    
       Set oChart = XLSHEET.ChartObjects.Add(450, 100, 300, 200).Chart
       'oChart.SetSourceData Source:=XLSHEET.Range("A1").Resize(cNumRows, cNumCols)
      ' oChart.SetSourceData Source:=XLSHEET.Range("$D6:$O6")
       oChart.SetSourceData Source:=XLSHEET.Range("$D$1:$O$1").Resize(RENG1&, COLUM& - 4)
    ' Make Excel Visible:
    XLAPP.Visible = True

    XLAPP.UserControl = True

   Close #N1%
   On Error GoTo 30
   XLSHEET.SaveAs NOMBREDEARCHIVO$
   On Error GoTo 0
   '
   XLAPP.Quit      ' cierra EXCEL que se esta ejecutando minimizado
   Set XLAPP = Nothing     ') libera memoria
   Set XLBOOK = Nothing    ') utilizada por
   Set XLSHEET = Nothing   ') las variables objeto
   '
   Screen.MousePointer = 1
   Call COMUNI("Planilla Excel Creada Exitosamente:\ \" + REPLACAR$(NOMBREDEARCHIVO$, "\", "/"))
   '
Exit Sub
   '
   '
20 Resume 21
21 On Error GoTo 0
   MsgBox "Utilitario EXCEL no Instalado"
   On Error GoTo 0
   Exit Sub
   '
30 Resume 31
31 On Error GoTo 0
   Screen.MousePointer = 1
   '
   ' Cierra el objeto Book, el cual es utilizado al crear
   ' una nueva hoja en el archivo Excel.
   ' El valor False es para que no pregunte si quiere
   ' guardar el archivo.
   XLBOOK.Close (False)
   XLAPP.Quit ' Cierra Excel
   Set XLAPP = Nothing
   Set XLBOOK = Nothing
   Set XLSHEET = Nothing
   '
   Call MENSERR(24, "Imposible Crear Planilla Excel.")
   Exit Sub
   '
   Call RECONECTA
End Sub



Private Sub Label13_Change()
End Sub

Private Sub Label9_Change()
End Sub

Private Sub Command3_Click()

End Sub

Private Sub Command4_Click()
    Command4.Enabled = False
    'VALIDO QUE HAYA SELECCIONADO
    SELECCION% = 0
    For I& = 1 To MSF0.Rows - 1
      If MSF0.TextMatrix(I&, 0) = "*" Then
         SELECCION% = 1
      End If
    Next I&
    
    If SELECCION% < 1 Then
         Call COMUNI("Debe Seleccionar Al Menos un Proyecto")
         GoTo 99
    End If
    '
    'CARGO LOS PROYECTOS
    MSF2.Cols = 4 'PARA NO PASARSE DE C OLUMUNAS SI SE REALIZAN VARIOS CLICK
    VALCOLUM% = MSF0.Cols - 1
    For I& = 1 To MSF0.Rows - 1
      If MSF0.TextMatrix(I&, 0) = "*" Then
         Call LISTARPROYECTO(XVALO(MSF0.TextMatrix(I&, 4)), VALCOLUM%)
         VALCOLUM% = VALCOLUM% + 1
      End If
    Next I&
    '
    'CARGO LOS ITEMS DE LOS PROYECTOS
    MSF1.Cols = 6
    VALCOLUM% = MSF1.Cols - 2
    MSF1.Rows = 1 'INICIALIZO EL MSF
    For I& = 1 To MSF0.Rows - 1
      If MSF0.TextMatrix(I&, 0) = "*" Then
         Call CARGARHORASMAQUINA(XVALO(MSF0.TextMatrix(I&, 4)), VALCOLUM%)
         VALCOLUM% = VALCOLUM% + 1
      End If
    Next I&
    '
    'CARGO EL PROMEDIO DE LAS PIEZAS
    JK& = MSF2.Cols
    For I& = 1 To MSF2.Rows - 1
       For j& = 3 To JK& - 1
          VALOR = VALOR + XVALO(MSF2.TextMatrix(I&, j&))
          DIVI% = DIVI% + 1
       Next j&
      MSF2.Cols = JK& + 1
       
       MSF2.TextMatrix(I&, JK&) = FORMATNUM$(VALOR / DIVI%, "F10.0")
       VALOR = 0: DIVI% = 0
    Next I&
    MSF2.ColWidth(JK&) = 1500
    MSF2.ColAlignment(JK&) = flexAlignRightTop
    MSF2.TextMatrix(0, JK&) = "PROMEDIO"
    '
    'CARGA LOS PROMEDIOS DE LA CARGA HORARIA
    JK& = MSF1.Cols
    VALORA = 0
    For I& = 1 To MSF1.Rows - 1
       For j& = 5 To JK& - 1
          VALORA = VALORA + XVALO(MSF1.TextMatrix(I&, j&))
          DIVI% = DIVI% + 1
       Next j&
      MSF1.Cols = JK& + 1
 
       MSF1.TextMatrix(I&, JK&) = FORMATNUM$((XVALO(VALORA / DIVI%)), "F10.0")
       VALORA = 0: DIVI% = 0
    Next I&
    '
    'ALINEA Y AGREGA EL TITULO
    MSF1.ColWidth(JK&) = 1500
    MSF1.ColAlignment(JK&) = flexAlignRightTop
    MSF1.TextMatrix(0, JK&) = "PROMEDIO"
    '
99  Command4.Enabled = True
End Sub

Private Sub exportarExcel_Click()
   '
   
   CommonDialog1.CancelError = True
   CommonDialog1.DialogTitle = "Exportar a Excel"
       ' cancelado, si no da error - CommonDialog1.Flags = &H2& + &H4& + &H8 + &H800 + &H2000 + &H8000 'cdlOFNCreatePrompt
   CommonDialog1.FILTER = "Excel Files(*.XLS)|*.XLS|"
   CommonDialog1.FilterIndex = 1
   On Error GoTo 99
   CommonDialog1.ShowOpen
   On Error GoTo 0
   '
   DoEvents
   Screen.MousePointer = 11
   AREXPO$ = UCase$(CommonDialog1.FileName)
   If TRIM$(AREXPO$) <> "" Then
      Call GENEXCEL2MSF(AREXPO$, MSF1, MSF2)
   End If
   '
99

End Sub

Private Sub Form_Load()

'PRESENTAR LA LISTA DE LOS PROYECTOS
   Screen.MousePointer = 11
   CAMPOS$ = "REFERENCIA/A44;Fec.Calc./D10;Hora Calc./A10;Indice/F0"
   Call CARGA_ENCABEZADO(CARHORAR.MSF0, CAMPOS$, CARHORAR)
   
   j& = 0
   For U& = 1 To ULTREG&("PROFABRI")
     XX$ = REGLEIDO$("PROFABRI", U&)
     DENOPRO$ = TRIM$(Mid$(XX$, 5, 48))
     MODOCA$ = Mid$(XX$, 63, 2)
     FECA% = CVI(Mid$(XX$, 65, 2))
     HOCA% = CVI(Mid$(XX$, 67, 2))
     If DENOPRO$ <> "" Then
       j& = j& + 1
       MSF0.Rows = j& + 1
       Me.MSF0.TextMatrix(j&, 1) = DENOPRO$
       Me.MSF0.TextMatrix(j&, 2) = FECHATEX$(FECA%)
       Me.MSF0.TextMatrix(j&, 3) = HORATEX$(HOCA%)
       Me.MSF0.TextMatrix(j&, 4) = Str$(U&)
     End If
   Next U&
   Screen.MousePointer = 1
   
   CAMPOS$ = "CÓDIGO/A18;PIEZA/A24;CANTIDAD/F11.1"
   Call CARGA_ENCABEZADO(CARHORAR.MSF2, CAMPOS$, CARHORAR)
   
   CAMPOS$ = "CODIGO/f0;MÁQUINA/A16;CODI-OPER/F0;OPERADOR/A18;TIEMPO/F8.0"
   Call CARGA_ENCABEZADO(CARHORAR.MSF1, CAMPOS$, CARHORAR)
   
   'PARA AGRUPAR CELDAS
   MSF1.Rows = 2
   MSF1.MergeCells = flexMergeRestrictColumns
   MSF1.MergeCol(2) = True

'VER COMO SE CALCULA LOS TIEMPOS DE LA GRILLA SUPERIOR
'PRESENTAR LA INFO EN LA GRILLA
'A = CATOPREP%(CI1%, NOPE%) EL NOMBRE DEL OPERARIO.
'MAQUISTAN$(CIXI%, NOPE%, Optional ALTEROP%)

''''   CNP% = 0
''''   SQLX$ = "SELECT count(*) as CANOPE from SecOper"
''''   SQLX$ = SQLX$ + " WHERE CodIConj=" & CIXI% & ""
''''   SQLX$ = SQLX$ + " AND StatOper = 1"
''''   Dim SECOPTEMP As ADODB.Recordset
''''   Set SECOPTEMP = FLEXCONN.Execute(FILTSQL$(SQLX$))
''''   '
''''   On Error Resume Next
''''   CNP% = SECOPTEMP!CANOPE    'RecordCount
''''   On Error GoTo 0
''''   Set SECOPTEMP = Nothing
''''   '
''''90 CANOPS% = CNP%

End Sub

Sub LISTARPROYECTO(IPROGRA&, COLUM%)
   If IPROGRA& < 1 Then
        Call MENSERR(24, "Imposible Abrir Programa '0'")
        Exit Sub
      Else
        If MSF2.Cols - 2 <= COLUM% + 1 Then
           MSF2.Cols = COLUM%
        End If

        'ESTE ARCHIVO CONTIENE LOS PROYECTOS CARGADOS
        'RECORRO LOS ITEM
        SCMA = Screen.MousePointer
        Screen.MousePointer = 11
'        COLUM% = MSF2.Cols
'        MSF2.Cols = COLUM% + 1
        RFF$ = RECFILT$("DEPROFAB", 1, MKS$(IPROGRA&))
        XA$ = FILTERGETRECORD$("PROFABRI", 1, MKS$(IPROGRA&))
        DENOPRO$ = TRIM$(Mid$(XA$, 5, 48))
        
        MSF2.TextMatrix(0, COLUM% - 1) = DENOPRO$
        j& = 0
        For U& = 1 To Len(RFF$) Step 4
          REFI& = CVS(Mid$(RFF$, U&, 4))
          XX$ = REGLEIDO$("DEPROFAB", REFI&)
          CIXI% = CVI(Mid$(XX$, 5, 2))
          CAXX = CVS(Mid$(XX$, 7, 4))
          CODI$ = CODEXT$(CIXI%)
          REG& = CONTROLMSF(Me.MSF2, 1, CODI$)  'CONTROLO SI EL CODIGO YA EXISTE EN EL LIST SOLO AGREGO LA CANTIDAD EN LA COLUMNA CORRESPONDIENTE
          If REG& < 1 Then
             'SI EL CODIGO NO ESTA EN LA GRILLA - AGREGA
             MSF2.Rows = MSF2.Rows + 1
             MSF2.TextMatrix(MSF2.Rows - 1, 1) = CODI$
             MSF2.TextMatrix(MSF2.Rows - 1, 2) = DESCRIT0$(CIXI%)
             MSF2.TextMatrix(MSF2.Rows - 1, COLUM% - 1) = FORMATNUM$(CAXX, "F10.1")
          Else
'             SI ESTA EN LA GRILLA LE PONE LA CANTIDAD A LA COLUMNA CORRESPONDIENTE EN REG&
             MSF2.TextMatrix(REG&, COLUM% - 1) = FORMATNUM$(CAXX, "F10.1")
          End If
        Next U&
        '
        'ALINEO LAS COLUMNAS AGREGADAS.
        For I& = 3 To MSF2.Cols - 1
          MSF2.ColWidth(I&) = 1300
          MSF2.ColAlignment(I&) = flexAlignRightTop
        Next I&

        Screen.MousePointer = SCMA
   End If
End Sub
'





Private Sub MSF0_Click()
   I& = MSF0.MouseRow
   j& = MSF0.MouseCol
   If I& = 0 Or I& > MSF0.Rows Then Exit Sub
   '
   MSF0.Row = I&
   MSF0.COL = j&
   '
   COL% = MSF1.MouseCol
   If COL% = 0 Then
     On Error Resume Next
     Call COLOREAR_GRILLA_SOLAFILA(CARHORAR.MSF0, &HC0C0C0, I&)
     On Error GoTo 0
   End If

   
End Sub
Sub CARGARHORASMAQUINA(IPROGRA&, COLUM%)
        If IPROGRA& < 1 Then
          Call MENSERR(24, "Imposible Abrir Programa '0'")
          Exit Sub
        End If

        Dim CODMAQ%()
        Dim DESCRIPCIOX$()
        Dim TOTIX()
        Dim CANPROGAM()
        '
        'NOMBRE DEL PROYECTO
        XA$ = FILTERGETRECORD$("PROFABRI", 1, MKS$(IPROGRA&))
        DENOPRO$ = TRIM$(Mid$(XA$, 5, 48))

'        If MSF1.Cols < COLUM% Then
'           MSF1.Cols = COLUM%
'           MSF1.ColWidth(COLUM% - 1) = 11000
'        End If
        '
        'LEO EL ARCHIVO DONDE ESTA EL DETALLE DEL PROYECTO (IPROGRA&)
        ReDim Preserve CANPROGAM(1)
        CONDI$ = ""
        RFF$ = RECFILT$("DEPROFAB", 1, MKS$(IPROGRA&))
        For U& = 1 To Len(RFF$) Step 4
          REFI& = CVS(Mid$(RFF$, U&, 4))
          XX$ = REGLEIDO$("DEPROFAB", REFI&)
          CI1% = CVI(Mid$(XX$, 5, 2))
          CAXX = CVS(Mid$(XX$, 7, 4))
          CANTVECTOR = UBound(CANPROGAM)
          'TOMO PARA LA CONDICION QUE CODIGOS ESTAN DENTRO DEL PROYECTO EN CONDI$
          If CANTVECTOR < CI1% Then ReDim Preserve CANPROGAM(CI1%)
          CANPROGAM(CI1%) = CAXX 'CARGO LA CANTIDAD EN UN VECTOR  PARA LEERLA MAS ABAJO

          If U& = 1 Then
               CONDI$ = CONDI$ + " AND ( CODICONJ = " & CI1%
          Else
               CONDI$ = CONDI$ + " OR CODICONJ = " & CI1%
          End If
          
        Next U&

        If U& > 1 Then CONDI$ = CONDI$ + ")"
        'FIN DE CONDI$
        '
        'FILTRA LOS ARCHIVOS
        Call TRALOFIL("RECALMAQ", 1, MKS$(IPROGRA&))
'
        'CARGO EN UN VECTOR LAS MAQUINAS
        For U& = 1 To ULTREG&("!RECALMAQ")
          TTXX$ = Mid$(REGLEIDO$("!RECALMAQ", U&), 5)
          ReDim Preserve CODMAQ%(U&), DESCRIPCIOX$(U&), TOTIX(U&)
          CODMAQ%(U&) = XVALO(Mid$(TTXX$, 1, 6))
          DESCRIPCIOX$(U&) = TRIM$(ECOARCH$("PADMAQ", SAFETEXT$(CODMAQ%(U&))))
          TOTIX(U&) = XVALO(Mid$(TTXX$, 34, 2)) + 60 * XVALO(Mid$(TTXX$, 31, 2)) + 3600 * XVALO(Mid$(TTXX$, 26, 4))
'          ACUTOTIX = ACUTOTIX + TOTIX(U&)
        Next U&
        JK& = MSF1.Rows - 1
        For j& = 1 To U& - 1
          COMAQ% = CODMAQ%(j&)
          DESCRIOX$ = DESCRIPCIOX$(j&)
          TOTX = TOTIX(j&)
          SQLX$ = "SELECT  * FROM SecOper  WITH(NOLOCK) "
          SQLX$ = SQLX$ + " WHERE CODMAQ0 = " & COMAQ%
          SQLX$ = SQLX$ + CONDI$
          SQLX$ = SQLX$ & " AND StatOper > 0 "
          SQLX$ = SQLX$ + " AND (AltOper = 0 OR AltOper IS NULL ) "
          SQLX$ = SQLX$ + " ORDER BY NUMEOPER ASC"
          Set RS = READSET(SQLX$)
          With RS
            Do While Not .EOF
              CI1% = XVALO(!CODICONJ)
              CATEG% = XVALO(!CATEGOPS)
                 LOTESTA = LOTESTAN(CI1%)
                 If LOTESTA < 1 Then LOTESTA = 1
                 CICLOHE = XVALO(!PieCiclo)
                    If CICLOHE < 1 Then CICLOHE = 1
                 MINPREP = 60 * XVALO(!HorsPrep)
                 MINFIJO = 60 * XVALO(!HorsFijo)
                 MINVARI = LOTESTA * XVALO(!MinuStan)
                 MINHERR = LOTESTA * 60 * XVALO(!HorsMHer) / CICLOHE
                 PIECICL = XVALO(!PieCiclo)
                 MINLIMP = 60 * XVALO(!HorsLimp)
                 DEMODIS = XVALO(!PorDemor)
                 MINDEMO = (MINFIJO + MINVARI + MINHERR) * DEMODIS / 100
                 SOLAPAI% = XVALO(!CoefSolp)
                 MINLOTE = MINPREP + MINFIJO + MINVARI + MINHERR + MINDEMO
                 CANPROG = CANPROGAM(CI1%)
                 LOTESTA = LOTESTAN(CI1%)
                 If LOTESTA <= 0 Then LOTESTA = 1
                 HORAPROD = CANPROG * MINLOTE / 60 / LOTESTA
                 REG& = CONTROLMSF(Me.MSF1, 1, COMAQ%, 3, CATEG%) 'CONTROLO SI LA MAQUINA Y EL OPERARIO YA ESTA EN LA LISTA LO AGREGO EN LA MISMA FILA COLUMNA CORRESPONDIENTE
                 If MSF1.Cols - 2 < COLUM% Then
                    MSF1.Cols = COLUM% + 2
                 End If
                 If REG& < 1 Then
                     JK& = JK& + 1
                     MSF1.Rows = JK& + 1
                     MSF1.TextMatrix(0, COLUM% + 1) = DENOPRO$
                     MSF1.TextMatrix(JK&, 1) = COMAQ%
                     MSF1.TextMatrix(JK&, 2) = DESCRIOX$
                     MSF1.TextMatrix(JK&, 3) = CATEG%
                     MSF1.TextMatrix(JK&, 4) = TRIM$(ECOARCH$("CATOPER", Chr$(CATEG%)))
'                     MSF1.TextMatrix(JK&, COLUM% + 1) = HORMINSE$(3600 * HORAPROD)
                     MSF1.TextMatrix(JK&, COLUM% + 1) = TRIM$(FORMATNUM$(HORAPROD, "F10.0"))
                    ' MSF1.TextMatrix(jK&, 5) = CATEG%
                     CATEG2% = XVALO(!CATEGOPS2)
                     If CATEG2% > 0 Then
                        JK& = JK& + 1
                        MSF1.Rows = JK& + 1
                        MSF1.TextMatrix(JK&, 1) = COMAQ%
                        MSF1.TextMatrix(JK&, 2) = DESCRIOX$
                        MSF1.TextMatrix(JK&, 3) = CATEG2%
                        MSF1.TextMatrix(JK&, 4) = TRIM$(ECOARCH$("CATOPER", Chr$(CATEG2%)))
                        'MSF1.TextMatrix(JK&, COLUM% + 1) = HORMINSE$(3600 * HORAPROD)
                        MSF1.TextMatrix(JK&, COLUM% + 1) = TRIM$(FORMATNUM$(HORAPROD, "F10.0"))
                        'MSF1.TextMatrix(jK&, 5) = CATEG%
                     End If
                 Else
                     ' SI NO AGREGA EL DATO A LA COLUMNA EN LA FILA CORRESPONDIENTE
                     MSF1.TextMatrix(0, COLUM% + 1) = DENOPRO$
                     MSF1.TextMatrix(REG&, COLUM% + 1) = TRIM$(FORMATNUM$(HORAPROD, "F10.0"))
                 End If
40            .MoveNext
            Loop
          End With
        Next j&
        
        For I& = 5 To MSF1.Cols - 1
          MSF1.ColWidth(I&) = 1500
          MSF1.ColAlignment(I&) = flexAlignRightTop
        Next I&
End Sub




''''''   Dim COEQUI$(1024), TIEQUI(1024)
''''''   CANEQ% = 0
''''''   Dim COCATO%(1024), TICATO(1024)
''''''   CANCA% = 0
''''''   '
''''''   MSF1.Rows = 1
''''''   REBLA$ = REGBLAN$("CAPANECE")
''''''   JJ& = 0
''''''   For I& = 1 To MSF2.Rows - 1
''''''     CIKI% = CODINT%(MSF2.TextMatrix(I&, 1))
''''''     '
''''''     If CIKI% > 0 Then
''''''         CANPROG = XVALO(MSF2.TextMatrix(I&, 3))
''''''         Call EXPLOMUL(CIKI%)
''''''         '
''''''         FIN& = CONTAMUL%
'''''''         If CASUBCON% < 1 Then
'''''''           FIN& = 1
'''''''         End If
''''''         '
''''''         URE& = ULTREG&("SECOPER")
''''''         SUTIE = 0: SUMOD = 0: SUMAQ = 0
''''''         ICX% = 0
''''''         '
''''''         For UI& = 1 To FIN&
''''''           '
''''''           CIXI% = CIK%(UI&)
''''''           CODEX$ = CODEXT$(CIXI%)
''''''           CACOM = CANT(UI&)
''''''           '
''''''           ULTOP% = 0
''''''           Call ABREMASTER
''''''           SQLX$ = "SELECT * FROM Secoper"
''''''           SQLX$ = SQLX$ & " WHERE CodIConj = " & CIXI%
''''''           SQLX$ = SQLX$ & " AND StatOper >0 "
''''''           SQLX$ = SQLX$ + "ORDER BY NumeOper ASC"
''''''           Dim SECOPTEMP As ADODB.Recordset
''''''           Set SECOPTEMP = FLEXCONN.Execute(FILTSQL$(SQLX$))
''''''           '
''''''           REBLA$ = REGBLAN$("SECOPER") + String$(64, 0)
''''''           IOP% = 0
''''''           On Error Resume Next
''''''           With SECOPTEMP
''''''             .MoveFirst
''''''             If Err Then GoTo 25
''''''             '
''''''             Do Until .EOF = True
''''''               If IsNull(!AltOper) = True Or !AltOper < 1 Then
''''''                 NOPY% = XVALO(!NUMEOPER)
''''''                 DOPE$ = SAFETEXT(!DESCOPER)
''''''                 LOTESTA = LOTESTAN(CIXI%)
''''''                    If LOTESTA < 1 Then LOTESTA = 1
''''''                 CICLOHE = XVALO(!PieCiclo)
''''''                    If CICLOHE < 1 Then CICLOHE = 1
''''''                 MINPREP = 60 * XVALO(!HorsPrep)
''''''                 MINFIJO = 60 * XVALO(!HorsFijo)
''''''                 MINVARI = LOTESTA * XVALO(!MinuStan)
''''''                 MINHERR = LOTESTA * 60 * XVALO(!HorsMHer) / CICLOHE
''''''                 PIECICL = XVALO(!PieCiclo)
''''''                 MINLIMP = 60 * XVALO(!HorsLimp)
''''''                 DEMODIS = XVALO(!PorDemor)
''''''                 MINDEMO = (MINFIJO + MINVARI + MINHERR) * DEMODIS / 100
''''''                 SOLAPAI% = XVALO(!CoefSolp)
''''''                 MINLOTE = MINPREP + MINFIJO + MINVARI + MINHERR + MINDEMO
''''''                 HORAPROD = CANPROG * CACOM * MINLOTE / 60 / LOTESTA
''''''                 '
''''''                 CAOPS = XVALO(!CantiOps)
''''''                 CATEGO% = XVALO(!CategOps)
''''''                 If CATEGO% < 1 Then CATEGO% = 1
''''''                 EQUIPO$ = AJUSTI$(SAFETEXT$(!CodMaq0), 6)
''''''                 EQUIP1$ = AJUSTI$(SAFETEXT$(!CodMaq1), 6)
''''''                 EQUIP2$ = AJUSTI$(SAFETEXT$(!CodMaq2), 6)
''''''                 EQUIP3$ = AJUSTI$(SAFETEXT$(!CodMaq3), 6)
''''''                 '
''''''                 If HORAPROD > 0.01 Then
''''''                   EQUIPO$ = EQUIPO$ + EQUIP1$ + EQUIP2$ + EQUIP3$
''''''                   For KU% = 1 To 24 Step 6
''''''                     EQUI$ = TRIM$(Mid$(EQUIPO$, KU%, 6))
''''''                     If EQUI$ <> "" Then
''''''                       RECAP$ = REBLA$
''''''                       Call REPLA(RECAP$, EQUI$, 1, 6)
''''''                       Call REPLA(RECAP$, MKI$(CIXI%), 7, 2)
''''''                       Call REPLA(RECAP$, MKI$(NOPY%), 9, 2)
''''''                       Call REPLA(RECAP$, DOPE$, 11, 48)
''''''                       Call REPLA(RECAP$, MKS$(CANPROG * CACOM), 59, 4)
''''''                       Call REPLA(RECAP$, MKI$(CIKI%), 67, 2)
''''''                       Call REPLA(RECAP$, MKS$(HORAPROD), 69, 4)
''''''                       Call REPLA(RECAP$, MKS$(HORAPROD), 77, 4)
''''''                       JJ& = JJ& + 1
''''''                       Call GRAREG("CAPANECE", RECAP$, JJ&)
''''''                       '
''''''                       For TUI% = 1 To CANEQ%
''''''                         If COEQUI$(TUI%) = EQUI$ Then GoTo 20
''''''                       Next TUI%
''''''                       CANEQ% = CANEQ% + 1
''''''                       COEQUI$(CANEQ%) = EQUI$
''''''                       TIEQUI(CANEQ%) = 0
''''''                       TUI% = CANEQ%
''''''                       CATOPERA% = CATOPREP%(CIXI%, NOPY%)
''''''                       '
''''''20                     TIEQUI(TUI%) = TIEQUI(TUI%) + HORAPROD
''''''                       '
''''''                     End If
''''''                   Next KU%
''''''                   '
''''''                   For TUI% = 1 To CANCA%
''''''                     If COCATO%(TUI%) = CATEGO% Then GoTo 30
''''''                   Next TUI%
''''''                   CANCA% = CANCA% + 1
''''''                   COCATO%(CANCA%) = CATEGO%
''''''                   TICATO(CANCA%) = 0
''''''                   TUI% = CANCA%
''''''                   '
''''''30                 TICATO(TUI%) = TICATO(TUI%) + HORAPROD * CAOPS
''''''                   '
''''''                 End If
''''''               End If
''''''             .MoveNext
''''''             Loop
''''''           End With
''''''
''''''25       Next UI&
''''''         SECOPTEMP.Close
''''''         Set SECOPTEMP = Nothing
''''''
''''''       End If
''''''   Next I&
''''''   List1.ListIndex = (-1)
''''''   '
''''''   Call SETULTREG("CAPANECE", JJ&)
''''''   Call FILESORT("CAPANECE", "", 1, 3, "ASC")
''''''   '
''''''   '
''''''   JK& = 0
''''''   SUTIE = 0: J4& = 0
''''''   For TUI% = 1 To CANEQ%
''''''     If TRIM$(COEQUI$(TUI%)) <> "" Then
''''''       If TIEQUI(TUI%) > 0 Then
''''''         J4& = J4& + 1
''''''         TTXX$ = Space$(128)
''''''         Call REPLA(TTXX$, COEQUI$(TUI%), 1, 6)
''''''          MSF1.Rows = J4& + 1
''''''          MSF1.TextMatrix(J4&, 1) = COEQUI$(TUI%)
''''''         Call REPLA(TTXX$, ECOARCH$("PADMAQ", COEQUI$(TUI%)), 9, 15)
''''''          MSF1.TextMatrix(J4&, 2) = ECOARCH$("PADMAQ", COEQUI$(TUI%))
''''''           SUTIE = SUTIE + TIEQUI(TUI%)
''''''           HMS$ = HORMINSE$(3600 * TIEQUI(TUI%))
''''''           While Len(HMS$) < 12: HMS$ = " " + HMS$: Wend
''''''         Call REPLA(TTXX$, HMS, 24, 12)
''''''          MSF1.TextMatrix(J4&, 4) = HMS
'''''''         List4.AddItem (TTXX$)
''''''         Call GRAREG("!RECALMAQ", MKS$(IPROGRA&) + TTXX$, J4&)
''''''       End If
''''''     End If
''''''   Next TUI%
'''''''   List4.Refresh
''''''   Label2 = HORMINSE$(3600 * SUTIE)
''''''   Call SETULTREG("!RECALMAQ", J4&)
''''''
''''''''''A = CATOPREP%(CI1%, NOPE%) 'EL NOMBRE DEL OPERARIO.
''''''''''MAQUISTAN$(CIXI%, NOPE%, Optional ALTEROP%)
''''''''''
''''''''''   SQLX$ = "SELECT Catoper_Prep from SecOper"
''''''''''   SQLX$ = SQLX$ + " WHERE CodIConj=" & CIXI% & ""
''''''''''   SQLX$ = SQLX$ + " AND StatOper = 1"
''''''''''   Dim SECOPTEMP As ADODB.Recordset
''''''''''   Set SECOPTEMP = FLEXCONN.Execute(FILTSQL$(SQLX$))
''''''''''   '
''''''''''   On Error Resume Next
''''''''''   CNP% = SECOPTEMP!CANOPE    'RecordCount
''''''''''   On Error GoTo 0
''''''''''   Set SECOPTEMP = Nothing
''''''''''   '
''''''''''90 CANOPS% = CNP%
''''''
Private Sub MSF1_Click()
    Static MODO  As Boolean
    If (MSF1.MouseRow = 0) Then
          
        ' Ordena en forma ascendente
        If MODO Then
            MSF1.COL = MSF1.MouseCol
            MSF1.Sort = 2
            MODO = False
        ' Ordena en forma descendente
        Else
            MSF1.COL = MSF1.MouseCol
            MSF1.Sort = 1
            MODO = True
        End If
    End If
End Sub
Function CONTROLMSF(MSF As MSFlexGrid, COLUM%, DATO1, Optional COLUM2%, Optional DATO2)
   CONTROLMSF = 0
   'CONTROLA SI EXISTE EL DATO EN LA COLUMNA PASADA COMO PARAMETRO EN ESE CASO DEVUELVE EL NUMERO DE FILA
   For I& = 1 To MSF.Rows - 1
      If TRIM$(DATO1) = TRIM$(MSF.TextMatrix(I&, COLUM%)) Then
         If COLUM2% > 0 Then
            If TRIM$(DATO2) = TRIM$(MSF.TextMatrix(I&, COLUM2%)) Then
               CONTROLMSF = I&
            End If
         Else
           CONTROLMSF = I&
         End If
      End If
   Next I&
End Function

Private Sub MSF1_DblClick()
   I& = MSF1.MouseRow
   j& = MSF1.MouseCol
   If I& = 0 Or I& > MSF1.Rows Then Exit Sub
   '
   MSF1.Row = I&
   MSF1.COL = j&
   '
   '
End Sub
Private Sub MSF_2_Click()
    Static MODO  As Boolean
    If (MSF_2.MouseRow = 0) Then
          
        ' Ordena en forma ascendente
        If MODO Then
            MSF_2.COL = MSF_2.MouseCol
            MSF_2.Sort = 2
            MODO = False
        ' Ordena en forma descendente
        Else
            MSF_2.COL = MSF_2.MouseCol
            MSF_2.Sort = 1
            MODO = True
        End If
    End If
End Sub

Private Sub msf2_Click()
    Static MODO  As Boolean
    If (MSF2.MouseRow = 0) Then
          
        ' Ordena en forma ascendente
        If MODO Then
            MSF2.COL = MSF2.MouseCol
            MSF2.Sort = 2
            MODO = False
        ' Ordena en forma descendente
        Else
            MSF2.COL = MSF2.MouseCol
            MSF2.Sort = 1
            MODO = True
        End If
    End If

End Sub




Private Sub MSF_2_DblClick()
   I& = MSF_2.MouseRow
   j& = MSF_2.MouseCol
   If I& = 0 Or I& > MSF_2.Rows Then Exit Sub
   '
   MSF_2.Row = I&
   MSF_2.COL = j&
   '
   '
End Sub



Private Sub TXTBUSCAR_Change()
    '
    If TRIM$(TXTBUSCAR) = "" Then
      MSF_2.Visible = False
      MSF1.Visible = True
      Exit Sub
    End If
    '
    MSF_2.Rows = 2
    For I& = 1 To MSF_2.Cols - 1
       MSF_2.TextMatrix(1, I&) = ""
    Next I&
    REG& = 0
    MSF_2.Visible = True
    MSF_2.ZOrder 0
    MSF_2.Cols = MSF1.Cols
    MSF_2.Redraw = False
    MSF1.Visible = False
    '
    'RECORRO EL MSF QUE TIENE TODO E IMPRIMO EN msf  LOS QUE MACHEAN
    For I& = 1 To MSF1.Rows - 1
      For j& = 1 To MSF1.Cols - 1
        If (InStr(UCase$(MSF1.TextMatrix(I&, j&)), UCase$(TXTBUSCAR))) Then
          If (TRIM$(CMBFILTRO.TEXT)) = "" Or TRIM$(UCase$(CMBFILTRO)) = TRIM$(UCase$(MSF1.TextMatrix(0, j&))) Then
            REG& = REG& + 1
            MSF_2.Rows = REG& + 1
            For H& = 1 To MSF_2.Cols - 1
              MSF_2.TextMatrix(REG&, H&) = MSF1.TextMatrix(I&, H&)
            Next H&
          End If
          If (TRIM$(CMBFILTRO.TEXT)) = "" Then Exit For
        End If
      Next j&
    Next I&
    '
    MSF_2.Redraw = True
    
    '
End Sub

Private Sub TXTBUSCAR_KeyPress(KeyAscii As Integer)
   If KeyAscii = 27 Then
      TXTBUSCAR.TEXT = ""
   End If
End Sub
Sub GENEXCEL2MSF(NOMBREDEARCHIVO$, MSF1 As MSFlexGrid, MSF2 As MSFlexGrid)
   Dim XLAPP As Excel.Application
   Dim XLBOOK As Excel.Workbook
   Dim XLSHEET As Excel.Worksheet
   '
   Dim FECHA As Date
   Dim VALOR As Double
   Dim TEXTO As String
   '
   On Error GoTo 20
   Set XLAPP = New Excel.Application
   Set XLBOOK = XLAPP.Workbooks.Add
   Set XLSHEET = XLBOOK.Worksheets.Add
   Dim objExcel As Object, OBJTEMP As Object
   Set objExcel = CreateObject("Excel.Application")
   objExcel.Workbooks.Add
   On Error GoTo 0
   '
   TIPODATO% = 0: COLUM& = 0
   Set MSF = MSF1
   VUELTA% = 0
   XLSHEET.Cells(1, 4) = "CARGA HORARIA DE PRODUCCIÓN"
   XLSHEET.Range("B1:G1").Select
   Selection.Merge

   Call ColorearRangoCeldasExcel("B1:G1", 43, objExcel, XLBOOK, XLSHEET, 1)

5  For I& = 1 To MSF.Rows
       COLUM& = 0
       For COLU1& = 1 To MSF.Cols
          If MSF.ColWidth(COLU1& - 1) = 0 Then GoTo 10 ' OBVIA LAS COLUMNAS OCULTAS
          COLUM& = COLUM& + 1
          '
          VALORX = MSF.TextMatrix(I& - 1, COLU1 - 1)
          If Len(VALORX) = 8 And Mid$(VALORX, 3, 1) = "/" And Mid$(VALORX, 6, 1) = "/" Then
             FEXU$ = VALORX
             VALORX = FETEXCOR1$(Int(FECHANUM(FEXU$)))
             TIPODATO% = 1
          End If
          ' ESTO ES PARA QUE EXPORTE COMO CERO Y NO ESPACIOS
          FFCC$ = UCase$(Left$(FORCAMPO$(COLU1&), 1))
          If FFCC$ = "I" Or FFCC$ = "F" Or FFCC$ = "G" Or FFCC$ = "$" Then
             If TRIM$(CStr(VALORX)) = "" Then
                VALORX = 0
             End If
             TIPODATO% = 2
          End If
          If FFCC$ = "C" Or FFCC$ = "A" Then VALORX = "'" & VALORX
          '
          '
          XLSHEET.Cells(RENG1& + I& + 1, COLUM&) = VALORX
          XLSHEET.Cells(RENG1& + I& + 1, COLUM&).EntireColumn.AutoFit
          XLSHEET.Cells(RENG1& + I& + 1, COLUM&).EntireRow.AutoFit
         
          If TIPODATO% = 1 Then XLSHEET.Cells(RENG1& + I& + 1, COLUM&).NumberFormat = "dd/mm/yyyy"
          'If TIPODATO% = 2 Then XLSHEET.Cells(RENG1&, COLU1&).NumberFormat = "#,##0.00"
          If TIPODATO% = 0 Then XLSHEET.Cells(RENG1& + I& + 1, COLUM&).NumberFormat = "@"
          TIPODATO% = 0
          If COLUM& = 2 Then
            If VALORX_ANT = VALORX Then
               'AA = "B1:E1"
               FL$ = SAFETEXT$(RENG1& + I&)
               If RANG$ = "" Then
                 'LA PRIMERA COINCIDENCIA GUARDA LA POSICION ANTERIOR
                 RANG$ = "B" & FL$ & ":G" & SAFETEXT$(XVALO(FL$) + 1)
               Else
                 'DESPUES SIEMPRE VA A INGRESAR ACA Y SELECCIONA DESDE LA POSICIION INICIAL HASTA QUE SE SIGA REPITIENDO
                 RANG$ = Left$(RANG$, 4) & SAFETEXT$(XVALO(FL$) + 1)
               End If
               If VUELTA% < 1 Then
                 COLORX% = COLORX% + 1
                 Call ColorearRangoCeldasExcel(RANG$, COLORX%, objExcel, XLBOOK, XLSHEET, 1)
               End If
'               XLSHEET.Range(RANG$).Select
'               Selection.Merge
            Else

               RANG$ = ""
               If VUELTA% < 1 Then
                 RANX$ = "B" & SAFETEXT$(RENG1& + I& + 1) & ":G" & SAFETEXT$(RENG1& + I& + 1)
                 COLORX% = COLORX% + 1
                 Call ColorearRangoCeldasExcel(RANX$, COLORX%, objExcel, XLBOOK, XLSHEET, 1)
               End If
            End If
            VALORX_ANT = VALORX

          End If
10     Next COLU1&
   Next I&
   '
   If VUELTA% < 1 Then
     Set MSF = MSF2
     VUELTA% = 1
     XLSHEET.Cells(I& + 2, 4) = "CANTIDAD DE PIEZAS"
     COLORX% = COLORX% + 1
     RANX$ = "B" & SAFETEXT$(I& + 2) & ":G" & SAFETEXT$(I& + 2)
     Call ColorearRangoCeldasExcel(RANX$, COLORX%, objExcel, XLBOOK, XLSHEET, 1)
     RENG1& = I& + 2
     GoTo 5
   End If
   Close #N1%
   On Error GoTo 30
   XLSHEET.SaveAs NOMBREDEARCHIVO$
   On Error GoTo 0
   '
   XLAPP.Quit      ' cierra EXCEL que se esta ejecutando minimizado
   Set XLAPP = Nothing     ') libera memoria
   Set XLBOOK = Nothing    ') utilizada por
   Set XLSHEET = Nothing   ') las variables objeto
   '
   Screen.MousePointer = 1
   Call COMUNI("Planilla Excel Creada Exitosamente:\ \" + REPLACAR$(NOMBREDEARCHIVO$, "\", "/"))
   '
Exit Sub
   '
   '
20 Resume 21
21 On Error GoTo 0
   MsgBox "Utilitario EXCEL no Instalado"
   On Error GoTo 0
   Exit Sub
   '
30 Resume 31
31 On Error GoTo 0
   Screen.MousePointer = 1
   '
   ' Cierra el objeto Book, el cual es utilizado al crear
   ' una nueva hoja en el archivo Excel.
   ' El valor False es para que no pregunte si quiere
   ' guardar el archivo.
   XLBOOK.Close (False)
   XLAPP.Quit ' Cierra Excel
   Set XLAPP = Nothing
   Set XLBOOK = Nothing
   Set XLSHEET = Nothing
   '
   Call MENSERR(24, "Imposible Crear Planilla Excel.")
   Exit Sub
   '
End Sub

