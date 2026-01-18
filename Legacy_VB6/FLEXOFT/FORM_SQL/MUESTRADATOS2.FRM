VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.OCX"
Begin VB.Form MUESTRADATOS2 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Detalle de Composición de Deuda"
   ClientHeight    =   7875
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   15735
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7875
   ScaleWidth      =   15735
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   600
      Left            =   240
      TabIndex        =   9
      Top             =   7200
      Width           =   7455
      Begin VB.TextBox TXTBUSCAR 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   720
         TabIndex        =   11
         Top             =   160
         Width           =   2535
      End
      Begin VB.ComboBox CMBFILTRO 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   4200
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   160
         Width           =   3135
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00D8BD05&
         BackStyle       =   0  'Transparent
         Caption         =   "Filtrar:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   0
         Left            =   120
         TabIndex        =   13
         Top             =   160
         Width           =   585
      End
      Begin VB.Label Label71 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00D8BD05&
         BackStyle       =   0  'Transparent
         Caption         =   "Campo:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   3360
         TabIndex        =   12
         Top             =   160
         Width           =   705
      End
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
      Left            =   13920
      TabIndex        =   8
      Top             =   7320
      Width           =   1600
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Maximizar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Left            =   11760
      Picture         =   "MUESTRADATOS2.frx":0000
      TabIndex        =   7
      ToolTipText     =   "Emite Listado"
      Top             =   7320
      Width           =   2025
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
      Left            =   1440
      ScaleHeight     =   165
      ScaleWidth      =   1110
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   1170
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   2625
      OleObjectBlob   =   "MUESTRADATOS2.frx":030A
      Top             =   0
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   225
      Index           =   0
      Left            =   120
      OleObjectBlob   =   "MUESTRADATOS2.frx":053E
      TabIndex        =   1
      Top             =   240
      Visible         =   0   'False
      Width           =   1485
   End
   Begin MSFlexGridLib.MSFlexGrid GRID 
      Height          =   7000
      Left            =   240
      TabIndex        =   2
      ToolTipText     =   "Un Click Sobre Encabezado Para Ordenar/Doble Click Para Pasar a Grilla Inferior"
      Top             =   120
      Width           =   15300
      _ExtentX        =   26988
      _ExtentY        =   12356
      _Version        =   393216
      Rows            =   26
      Cols            =   4
      BackColor       =   16777215
      BackColorFixed  =   16501405
      BackColorBkg    =   16777215
      FillStyle       =   1
      GridLines       =   0
      SelectionMode   =   1
      AllowUserResizing=   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSFlexGridLib.MSFlexGrid MSF_2 
      Height          =   7000
      Left            =   240
      TabIndex        =   4
      ToolTipText     =   "Un Click Sobre Encabezado Para Ordenar/Doble Click Para Pasar a Grilla Inferior/Click Boton Derecho Imagen Digital"
      Top             =   120
      Visible         =   0   'False
      Width           =   15300
      _ExtentX        =   26988
      _ExtentY        =   12356
      _Version        =   393216
      Cols            =   4
      BackColor       =   16777215
      BackColorFixed  =   8421631
      BackColorBkg    =   12632256
      FillStyle       =   1
      SelectionMode   =   1
      AllowUserResizing=   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label CENTROCOSTO 
      Height          =   375
      Left            =   8520
      TabIndex        =   15
      Top             =   7320
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label ANO 
      Height          =   375
      Left            =   7680
      TabIndex        =   14
      Top             =   7320
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label Label7 
      BorderStyle     =   1  'Fixed Single
      Height          =   225
      Left            =   8520
      TabIndex        =   6
      Top             =   120
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Label5"
      Height          =   240
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Label TEPRUEBA 
      AutoSize        =   -1  'True
      Caption         =   "TEPRUEBA"
      Height          =   300
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Visible         =   0   'False
      Width           =   975
   End
End
Attribute VB_Name = "MUESTRADATOS2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim FrmLeft, FrmTop, FrmWidth, FrmHeight As Long
Dim GridTop, GridLeft, GridWidth, GridHeight As Long
Dim FRAMETop, FRAMELeft, FRAMEWidth, FRAMEHeight As Long
Dim Com1Top, Com2Top, Com2Left As Long
Dim Com1LEFT As Long

Dim LecturaDatosOriginales As Byte


Private Sub BOTEXIT_Click()
  Unload Me
End Sub

Private Sub CMBFILTRO_Change()
   TXTBUSCAR_Change
End Sub

Private Sub Command1_Click()
   Command1.Enabled = False
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
     If MSF_2.Visible = True Then
      Call GENEXCEL(AREXPO$, MUESTRADATOS2.MSF_2)
     Else
      Call GENEXCEL(AREXPO$, MUESTRADATOS2.GRID)
     End If
   End If
   '

99 Command1.Enabled = True

End Sub
Private Sub Command2_Click()
    '
    If Command2.Caption = "Maximizar" Then
        If LecturaDatosOriginales = 0 Then
            ' GUARDA DIMENSIONES Y POSICION DEL FORMULARIO, GRID Y BOTONES
            FrmLeft = Me.Left: FrmTop = Me.Top: FrmWidth = Me.Width: FrmHeight = Me.Height
            GridTop = GRID.Top: GridLeft = GRID.Left: GridWidth = GRID.Width: GridHeight = GRID.Height
            Com1Top = Command1.Top: Com1Top = Command1.Top: Com1LEFT = Command1.Left
            Com2Top = Command2.Top: Com2Top = Command2.Top: Com2Left = Command2.Left

            FRAMETop = Frame1.Top: FRAMELeft = Frame1.Left: FRAMEWidth = Frame1.Width: FRAMEHeight = Frame1.Height
            LecturaDatosOriginales = 1
        End If
    
        QCOL% = GRID.Cols
        ANCHOCOL& = 0
        For COLUM% = 1 To QCOL% - 1
            ANCHOCOL& = ANCHOCOL& + GRID.ColWidth(COLUM%)
        Next COLUM%
        ANCHOSCR& = Screen.Width
        If ANCHOCOL& > ANCHOSCR& Then
            Me.Left = 0: Me.Width = ANCHOSCR&
            GRID.Width = ANCHOSCR& - 150
        Else
            LEFTFORM& = Int((ANCHOSCR& - ANCHOCOL&) / 2)
            Me.Left = LEFTFORM&: Me.Width = ANCHOCOL&
            GRID.Width = ANCHOCOL& - 150
        End If
        '
        QFILAS% = GRID.Rows
        altoFila& = 0
        For Fila% = 1 To QFILAS% - 1
            altoFila& = altoFila& + GRID.RowHeight(Fila%)
        Next Fila%
        ALTOSCR& = Screen.Height
        If altoFila& + 1600 > ALTOSCR& Then
            Me.Top = 0: Me.Height = ALTOSCR& - 600
            GRID.Height = ALTOSCR& - 1600
        Else
            TOPFORM& = Int((ALTOSCR& - (altoFila& + 1600)) / 2)
            Me.Top = TOPFORM&: Me.Height = altoFila& + 1600
            GRID.Height = altoFila& + 600
        End If
        '
        BORDEINF& = GRID.Height + 100
        Command2.Top = BORDEINF&
        Command1.Top = BORDEINF&
        Frame1.Top = BORDEINF&
        Command1.Left = Me.Width - 1800
        Command2.Caption = "Minimizar"
        
    Else
        Me.Left = FrmLeft: Me.Top = FrmTop: Me.Width = FrmWidth: Me.Height = FrmHeight
        GRID.Top = GridTop: GRID.Left = GridLeft: GRID.Width = GridWidth: GRID.Height = GridHeight
        Frame1.Top = FRAMETop: Frame1.Left = FRAMELeft: Frame1.Width = FRAMEWidth: Frame1.Height = FRAMEHeight
        Command2.Top = Com2Top: Command2.Top = Com2Top: Command2.Left = Com2Left
        Command1.Top = Com1Top: Command1.Top = Com1Top: Command1.Left = Com1LEFT
        Command2.Caption = "Maximizar"
    End If
    MSF_2.Height = GRID.Height
    MSF_2.Width = GRID.Width
    MSF_2.Top = GRID.Top
    MSF_2.Left = GRID.Left
    
    '
End Sub
Sub GENEXCEL_PRUEBA_GRAFICO(NombreDeArchivo$, GRID As msFlexGrid)
   Dim XLAPP As Excel.Application
   Dim XLBOOK As Excel.Workbook
   Dim XLSHEET As Excel.Worksheet
   Dim oChart As Object     ' Excel Chart

   '
   Dim FECHA As Date
   Dim valor As Double
   Dim TEXTO As String
   '
   On Error GoTo 20
   Set XLAPP = New Excel.Application
   Set XLBOOK = XLAPP.Workbooks.Add
   Set XLSHEET = XLBOOK.Worksheets.Add
   On Error GoTo 0
   '
   TipoDato% = 0: COLUM& = 0
   ReDim aTemp(GRID.Rows To GRID.Rows, 3 To GRID.Cols - 1)

   For i& = 1 To GRID.Rows
       RENG1& = i&: COLUM& = 0
       For COLU1& = 1 To GRID.Cols
          If GRID.ColWidth(COLU1& - 1) = 0 Then GoTo 10 ' OBVIA LAS COLUMNAS OCULTAS
          COLUM& = COLUM& + 1
          '
          VALORX = GRID.TextMatrix(i& - 1, COLU1 - 1)
          If Len(VALORX) = 8 And Mid$(VALORX, 3, 1) = "/" And Mid$(VALORX, 6, 1) = "/" Then
             FEXU$ = VALORX
             VALORX = FETEXCOR1$(Int(FECHANUM(FEXU$)))
             TipoDato% = 1
          End If
          ' ESTO ES PARA QUE EXPORTE COMO CERO Y NO ESPACIOS
          FFCC$ = UCase$(Left$(FORCAMPO$(COLU1&), 1))
          If FFCC$ = "I" Or FFCC$ = "F" Or FFCC$ = "G" Or FFCC$ = "$" Then
             If TRIM$(CStr(VALORX)) = "" Then
                VALORX = 0
             End If
             TipoDato% = 2
          End If
          If FFCC$ = "C" Or FFCC$ = "A" Then VALORX = "'" & VALORX
          '

          XLSHEET.Cells(RENG1&, COLUM&) = VALORX
          If RENG1& = GRID.Rows And COLUM& > 3 And COLUM& < GRID.Cols Then
             aTemp(6, COLUM&) = VALORX
          End If
          If TipoDato% = 1 Then XLSHEET.Cells(RENG1&, COLUM&).NumberFormat = "dd/mm/yyyy"
          'If TIPODATO% = 2 Then XLSHEET.Cells(RENG1&, COLU1&).NumberFormat = "#,##0.00"
          If TipoDato% = 0 Then XLSHEET.Cells(RENG1&, COLUM&).NumberFormat = "@"
          TipoDato% = 0
          VALORX_ANT = VALORX
10     Next COLU1&
   Next i&
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
   XLSHEET.SaveAs NombreDeArchivo$
   On Error GoTo 0
   '
   XLAPP.Quit      ' cierra EXCEL que se esta ejecutando minimizado
   Set XLAPP = Nothing     ') libera memoria
   Set XLBOOK = Nothing    ') utilizada por
   Set XLSHEET = Nothing   ') las variables objeto
   '
   
'    Dim oXL As Object        ' Excel application
'    Dim oBook As Object      ' Excel workbook
'    Dim oSheet As Object     ' Excel Worksheet
    

    
'    ReDim aTemp(1 To cNumRows, 1 To cNumCols)
'
'    'Start Excel and create a new workbook
'    Set oXL = CreateObject("Excel.application")
'    Set oBook = oXL.Workbooks.Add
'    Set oSheet = oBook.Worksheets.Item(1)
'
'    ' Insert Random data into Cells for the two Series:
'    Randomize Now()
'    For iRow = 1 To cNumRows
'       For iCol = 1 To cNumCols
'          aTemp(iRow, iCol) = Int(Rnd * 50) + 1
'       Next iCol
'    Next iRow
   ' oSheet.Range("A1").Resize(cNumRows, cNumCols).Value = aTemp
    
    'Add a chart object to the first worksheet
        
           
   
   
   
   
   
   
   
   
   Screen.MousePointer = 1
   Call COMUNI("Planilla Excel Creada Exitosamente:\ \" + REPLACAR$(NombreDeArchivo$, "\", "/"))
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



Private Sub Command3_Click()

End Sub

Private Sub Form_Load()
'   GRID.Cols = 5
'   For i = 1 To 100
'      GRID.Rows = i + 1
'      GRID.TextMatrix(i, 1) = "AAAA" & i
'      GRID.TextMatrix(i, 2) = "BBBB" & i
'      GRID.TextMatrix(i, 3) = "CCCC" & i
'      GRID.TextMatrix(i, 4) = "DDDD" & i
'   Next i
   
'   'EN LA FUNCION DE CARGAR LOS ENCABEZADOS HABRIA QUE ASIGNAR LOS ANCHOS DE LABEL 5 Y ASIGNARSELO A LABEL8
'   For U& = 1 To GRID.Cols
'          JI& = JI& + 1
'          ' PROVISORIAMENTE A VER SI FUNCIONA
'          'AQUI AGREGAR A LABEL5 LA CANTIDAD DE CARACTERES
'          ANCH& = 20
'          Label5.Caption = FORMATNUM$("9999999.9999999", "A25" + "*")
'          GRID.ColWidth(JI& - 1) = Label5.Width
'25        If JI& >= Label8.Count Then
'             NUPIC% = Label8.Count
'             Load Label8(NUPIC%)
'             GoTo 25
'          End If
'          Label8(U&) = Label5.Width
'   Next U&
'   '
'   Label7 = 100
   '
End Sub

Private Sub Label13_Change()
   Call ACTUTOTA2
End Sub

Private Sub GRID_DblClick()
    REG& = GRID.MouseRow
    
    If (REG& > 0) Then
       Call T_Espera(1 / 5)
       COLUM& = GRID.MouseCol
       If COLUM& >= GRID.Cols - 2 Then Exit Sub
       If REG& >= GRID.Rows - 2 Then Exit Sub
       If COLUM& <= 2 Then Exit Sub 'para que aplique solo a los meses
       CUENTACONT$ = GRID.TextMatrix(REG&, 1)
       CODCUENTA% = CUECOINT%(CUENTACONT$)
       CODIEMPRE% = XVALO(GRID.TextMatrix(REG&, GRID.Cols - 1))
       TEXTOMES$ = GRID.TextMatrix(0, COLUM&)
       NUMES% = XVALO(MESNUM$(UCase$(TEXTOMES$)))
       POS1% = InStr(TRIM$(TEXTOMES$), " ")
       MES$ = Left$(TEXTOMES$, POS1% - 1)
       ANO1$ = TRIM$(Right(TEXTOMES$, Len(TEXTOMES$) - Len(MES$)))
       
       MESFILTRO$ = "01/" & MESNUM$(UCase$(TEXTOMES$)) & "/" & "2024"
       ANOX_ACTUAL% = Year(MESFILTRO$)
       If CODCUENTA% < 1 Then Exit Sub
       
       Call CARGA_ENCABEZADO(MUESTRADATOS3.GRID, "Cta/F6;R.Social/A32;Fecha/D8;Documento/A22;Importe/F14.2;Medios de Pago/A100", MUESTRADATOS3, 1, 1)
      '
      'CARGO ENCABEZADO DE GRILLA DE FILTRO
       MUESTRADATOS3.MSF_2.Rows = 3
      'MUESTRADATOS3.MSF_2.FixedRows = 2
       Call CARGA_ENCABEZADO(MUESTRADATOS3.MSF_2, "Cta/F6;R.Social/A32;Fecha/D8;Documento/A22;Importe/F14.2;Medios de Pago/A100", MUESTRADATOS3, 1, 1)
   
       MES1% = GRID.MouseCol - 2
       ANOSEL% = XVALO(ANO.Caption)
       CENTCOS = XVALO(CENTROCOSTO.Caption)
       i% = 0
       SQLX$ = "SELECT  FECHEMISI,NUPROVE,cuecontai ,CenCosto, IDebeComp, IHabeComp,OpciMovim,CODICOM_LAR, "
       SQLX$ = SQLX$ + " MONTH(FECHEMISI) as mes, YEAR(FECHEMISI) AS ANO "
       SQLX$ = SQLX$ + " From CajaBanc C WITH(NOLOCK)" ' LE AGREGO C COMO SI FUERA UN JOIN POR QUE LA CONDI_PUBLIC VIEN DE UN JOIN CON ESTE IDENTIFICADOR DE TABLA
       SQLX$ = SQLX$ + " Where NuProve > 0 And CenCosto > 0 And CUECONTAI = " & CODCUENTA%
       SQLX$ = SQLX$ + " And Year(FECHEMISI) = " & ANOX_ACTUAL%
       SQLX$ = SQLX$ + " And MONTH(FECHEMISI) = " & NUMES%
       If TRIM$(STPCOS07.CONDI_PUBLIC$) <> "" Then SQLX$ = SQLX$ + " And " & STPCOS07.CONDI_PUBLIC$
       SQLX$ = SQLX$ + " And CodiEmpr = " & CODIEMPRE%
       '
       SQLX$ = SQLX$ + " ORDER BY FECHEMISI, cuecontai   "
       Dim RS As New ADODB.Recordset
       Set RS = FLEXCONN.Execute(SQLX$)
       With RS
          Do While Not .EOF
            FE% = CVINT(!FECHEMISI)
            NPROV = XVALO(!NUPROVE)
            RASOPROV$ = rasocli$(-1 * NPROV)
            IMPORTE1# = XVALO(!Idebecomp) - XVALO(!IHABECOMP)
            TIPODOCU$ = SAFETEXT$(!OpciMovim)
            NUMEDOCU$ = SAFETEXT$(!codicom_lar)
            i% = i% + 1
            'MUESTRADATOS3.GRID.Rows = I%
            MUESTRADATOS3.GRID.Rows = MUESTRADATOS3.GRID.Rows + 1

            MUESTRADATOS3.GRID.TextMatrix(i%, 1) = NPROV
            MUESTRADATOS3.GRID.TextMatrix(i%, 2) = RASOPROV$
            MUESTRADATOS3.GRID.TextMatrix(i%, 3) = FECHATEX$(FE%)
            MUESTRADATOS3.GRID.TextMatrix(i%, 4) = NUMEDOCU$
            MUESTRADATOS3.GRID.TextMatrix(i%, 5) = FORMATNUM$(IMPORTE1#, "F12.2")
            MUESTRADATOS3.GRID.TextMatrix(i%, 6) = MediosDePagosImputados$(NPROV, TIPODOCU$, NUMEDOCU$)
            .MoveNext
          Loop
       End With
       RS.Close
       Set RS = Nothing
       MUESTRADATOS3.Left = MUESTRADATOS2.Left + 500
       MUESTRADATOS3.Top = MUESTRADATOS2.Top + 500
       MUESTRADATOS3.Caption = "DETALLE DE CUENTA DE GASTOS: " & TRIM$(DENOCUE$(CODCUENTA%))
       MUESTRADATOS3.Show 1
    End If
    
End Sub
Function MediosDePagosImputados$(NPROV, TIPODOCU$, NUMEDOCU$)
      '
      DescripcionMPago$ = ""
      If UCase$(Left$(NUMEDOCU$, 3)) = "OP-" Then GoTo 80 ' SI ES UNA ORDEN DE PAGO MUESTRA DIRECTAMENTE LA OP
      '
      'FACTURA DE PROVEEDOR CANCELADA EN EL MOMENTO.
      'VALIDAR SI FUE DE PAGO CONTADO (QUE SE LIQUIDA EN EL MOMOENTO DE CARGAR LA FACTURA)
      CONDI$ = "NUPROVE = " & NPROV
      CONDI$ = CONDI$ + " AND CODICOM_LAR = '" & NUMEDOCU$ & "'"
      CONDI$ = CONDI$ + " AND VARIMOVIM = 'CTDO'"
      CanceladaDeContado% = CantRegistros&("CAJABANC", CONDI$, "NOMESS")
      If CanceladaDeContado% > 0 Then
         CONDI$ = "NUPROVE = " & NPROV
         CONDI$ = CONDI$ + " AND CODICOM_LAR = '" & NUMEDOCU$ & "'"
         CONDI$ = CONDI$ + " AND TIPOMOVIM = 'CHEMI'"
         DescripcionMPago$ = TRIM$(VALOBADA("CAJABANC", "DescriMov", CONDI, "NOMESS"))
         FechaEmicion$ = Left$(TRIM$(VALOBADA("CAJABANC", "FECHEMISI", CONDI, "NOMESS")), 10)
         FE% = CVINT(FechaEmicion$)
         DescripcionMPago$ = DescripcionMPago$ + " - " + FECHATEX$(FE%)
         GoTo 85
      End If
      
      'FACTURA DE PROVEEDOR CON ORDEN DE PAGO
      JJ& = 0
      MEDIOPAGO$ = "": DESCR1$ = ""
      Dim Aplipago As ADODB.Recordset
      strsql = "Select  DISTINCT CoDocAp_Lar from Aplipago where NuPROVE = " & NPROV
      strsql = strsql & " And  CoDocOr_LAR = '" & TRIM(NUMEDOCU$) & "'"
      Set Aplipago = FLEXCONN.Execute(strsql)
      With Aplipago
        Do While Not (.EOF)
           'OBTENGO LA ORDEN DE PAGO IMPUTADA
           DOCU$ = SAFETEXT$(!codocap_lar)
           CONDI$ = "CODICOM_LAR = '" & DOCU$ & "' AND   TIPOMOVIM = 'CHENTR' AND OPCIMOVIM = 'OP'"
           SQLX$ = "SELECT * FROM CAJABANC WITH(NOLOCK) WHERE " & CONDI$
           Set RS = READSET(SQLX$)
           'RECORRO EL CAJABANC PARA ESCRIBIR EN MEDIOPAGO LOS DIFERENTES MEDIOS DE PAGO QUE TENGA LA ORDEN DE PAGO.
           With RS
            Do While Not .EOF
              Descr$ = TRIM$(SAFETEXT$(!DESCRIMOV))
              'QUITO LOS ESPACIOS Y ME VOY QUEDANDO CON LOS DATOS SEPARADOS CON '-'
              Call TEXTOLOTES$(Descr$, " ")
              For i% = 1 To UBound(CADENATEX$)
                a1$ = CADENATEX$(i%)
                DESCR1$ = DESCR1$ + TRIM$(a1$)
                If TRIM$(a1$) <> "" Then DESCR1$ = DESCR1$ + " - "
              Next i%
              MEDIOPAGO$ = MEDIOPAGO$ + DESCR1$
              MEDIOPAGO$ = MEDIOPAGO$ + " ; "
              DESCR1$ = ""
              .MoveNext
            Loop
           End With
           .MoveNext
        Loop   '
      End With
      Aplipago.Close
      Set Aplipago = Nothing
      GoTo 90
      '
      'ORDENES DE PAGO - RECORRE DIRECTAMENTE EL CAJABANC
80    DOCU$ = NUMEDOCU$: MEDIOPAGO$ = "": DESCR1$ = ""
      CONDI$ = "CODICOM_LAR = '" & DOCU$ & "' AND   TIPOMOVIM = 'CHENTR' AND OPCIMOVIM = 'OP'"
      SQLX$ = "SELECT * FROM CAJABANC WITH(NOLOCK) WHERE " & CONDI$
      Set RS = READSET(SQLX$)
      With RS
       Do While Not .EOF
        Descr$ = TRIM$(SAFETEXT$(!DESCRIMOV))
        Call TEXTOLOTES$(Descr$, " ")
        'QUITO LOS ESPACIOS Y ME VOY QUEDANDO CON LOS DATOS SEPARADOS CON '-'
        For i% = 1 To UBound(CADENATEX$)
          a1$ = CADENATEX$(i%)
          DESCR1$ = DESCR1$ + TRIM$(a1$)
          If TRIM$(a1$) <> "" Then DESCR1$ = DESCR1$ + " - "
        Next i%
        MEDIOPAGO$ = MEDIOPAGO$ + DESCR1$
        MEDIOPAGO$ = MEDIOPAGO$ + " ; "
        DESCR1$ = ""
        .MoveNext
       Loop
       End With
'
85    If DescripcionMPago$ <> "" Then
         MediosDePagosImputados$ = DescripcionMPago$
      Else
        MediosDePagosImputados$ = Left$(MEDIOPAGO$, Len(MEDIOPAGO$) - 5)
90      If Len(MEDIOPAGO$) > 3 Then
           MediosDePagosImputados$ = Left$(MEDIOPAGO$, Len(MEDIOPAGO$) - 5)
        End If
      End If

End Function


Private Sub Label5_Click()
   Call ACTUTOTA2
End Sub


Private Sub Label9_Change()
    Call ACTUTOTA1
End Sub
Sub ACTUTOTA1()
   Static COTIHOY#
   If COTIHOY# = 0 Then
      COTIHOY# = TICAMONE(2)
   End If
   If COTIHOY# = 0 Then COTIHOY# = 1
   Label14 = FORMATNUM$((XVALO(Label9) * COTIHOY#) + XVALO(Label8), "F12.2")
   Label17 = FORMATNUM$((XVALO(Label13) * COTIHOY#) + XVALO(Label5), "F12.2")
End Sub
Sub ACTUTOTA2()
   Static COTIHOY#
   If COTIHOY# = 0 Then
      COTIHOY# = TICAMONE(2)
   End If
   If COTIHOY# = 0 Then COTIHOY# = 1
   Label17 = FORMATNUM$((XVALO(Label13) * COTIHOY#) - XVALO(Label5), "F12.2")
   Label14 = FORMATNUM$((XVALO(Label9) * COTIHOY#) + XVALO(Label8), "F12.2")
End Sub

Private Sub GRID_Click()
    
    If InStr(Me.Caption, "Informe de Gastos - Centro de Costos") > 0 Then Exit Sub
    Static MODO  As Boolean
    If (GRID.MouseRow = 0) Then
          
        ' Ordena en forma ascendente
        If MODO Then
            GRID.COL = GRID.MouseCol
            GRID.Sort = 2
            MODO = False
        ' Ordena en forma descendente
        Else
            GRID.COL = GRID.MouseCol
            GRID.Sort = 1
            MODO = True
        End If
    End If
End Sub
Private Sub MSF_2_Click()

    If InStr(Me.Caption, "Informe de Gastos - Centro de Costos") > 0 Then Exit Sub
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
Private Sub GRID_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If Button = 2 Then
    '
    i& = GRID.MouseRow
    J& = GRID.MouseCol
    If i& = 0 Or i& > GRID.Rows Then Exit Sub
    '
    On Error Resume Next
    GRID.Row = i&
    GRID.COL = J&
    On Error Resume Next

    NPX& = TRIM$(GRID.TextMatrix(i&, 1))
    If NPX& < 1 Then
        Exit Sub
    End If

    If Me.Caption = "Pedidos" Then
        '
        NUDOCU$ = TRIM$(Str$(NPX&))
        TIDOCUM$ = "Pedido de Cliente"
        While Len(NUDOCU$) < 6: NUDOCU$ = "0" + NUDOCU$: Wend
        NUDOCU$ = NUDOCU$ + " "
        Screen.MousePointer = 11
        '
        'BUSCA EL DOCUMENTO EN LA TABLA SQL CON LA NUEVA FUNCION
        DOCUNU& = DOCID&(TIDOCUM$, Str$(NPX&), NUDOCU$)
'        If DOCUNU& < 1 Then
'             Screen.MousePointer = 1
'             Call COMUNI("No se Encontró el Documento Impreso\Correspondiente a la " & DOCUMENTO_ELEGIDO$ & " .\  \Si se Trata de un Pedido Pendiente de Entrega\Puede Refrescarlo con la Opción\'Modificación de Pedido en Curso'.")
'             On Error Resume Next
'             OPPED07.Command46.SetFocus
'             OPPED07.Show 1
'             If Err Then
'              On Error GoTo 0
'             End If
'             Exit Sub
'        End If
        Call MUESTRADOC(DOCUNU&, "", TIDOCUM$, Str$(NPX&), NUDOCU$)
    Else
        TIDOCUM$ = "Nota de Pedido"
        NUDOCU$ = TRIM$(Str$(NPX&))
        DOCUNU& = 0
        '
        Call MUESTRADOC(DOCUNU&, "", TIDOCUM$, NUDOCU$)
    End If
    Screen.MousePointer = 1
  End If
End Sub

Private Sub MSF_2_DblClick()
    If (MSF_2.MouseRow > 0) Then
       Call T_Espera(1 / 2)
       CUENTACONT$ = MSF_2.TextMatrix(MSF_2.MouseRow, 1)
       CODCUENTA% = CUECOINT%(CUENTACONT$)
       CODIEMPRE% = XVALO(MSF_2.TextMatrix(MSF_2.MouseRow, 16))
       If CODCUENTA% < 1 Then Exit Sub
       
       Call CARGA_ENCABEZADO(MUESTRADATOS3.GRID, "Cta/F6;R.Social/A32;Fecha/D8;Documento/A22;Importe/F14.2", MUESTRADATOS3, 1, 1)
      '
      'CARGO ENCABEZADO DE GRILLA DE FILTRO
       MUESTRADATOS3.MSF_2.Rows = 3
      'MUESTRADATOS3.MSF_2.FixedRows = 2
       Call CARGA_ENCABEZADO(MUESTRADATOS3.GRID, "Cta/F6;R.Social/A32;Fecha/D8;Documento/A22;Importe/F14.2", MUESTRADATOS3, 1, 1)
   
       MES1% = MSF_2.MouseCol - 2
       ANOSEL% = XVALO(ANO.Caption)
       CENTCOS = XVALO(CENTROCOSTO.Caption)
       i% = 0
       SQLX$ = "SELECT  FECHEMISI,NUPROVE,cuecontai ,CenCosto, IDebeComp, IHabeComp,OpciMovim,CODICOM_LAR, "
       SQLX$ = SQLX$ + " MONTH(FECHEMISI) as mes, YEAR(FECHEMISI) AS ANO "
       SQLX$ = SQLX$ + " From CajaBanc WITH(NOLOCK)"
       SQLX$ = SQLX$ + " Where NuProve > 0 And CenCosto > 0 And CUECONTAI = " & CODCUENTA%
       SQLX$ = SQLX$ + " And Year(FECHEMISI) = " & ANOSEL%
       SQLX$ = SQLX$ + " And MONTH(FECHEMISI) = " & MES1%
       SQLX$ = SQLX$ + " And CENCOSTO = " & CENTCOS
       SQLX$ = SQLX$ + " And CodiEmpr = " & CODIEMPRE%
       '
       SQLX$ = SQLX$ + " ORDER BY FECHEMISI, cuecontai   "
       Dim RS As New ADODB.Recordset
       Set RS = FLEXCONN.Execute(SQLX$)
       With RS
          Do While Not .EOF
            FE% = CVINT(!FECHEMISI)
            NPROV = XVALO(!NUPROVE)
            RASOPROV$ = rasocli$(-1 * NPROV)
            IMPORTE1# = XVALO(!Idebecomp) - XVALO(!IHABECOMP)
            TIPODOCU$ = SAFETEXT$(!OpciMovim)
            NUMEDOCU$ = SAFETEXT$(!codicom_lar)
            i% = i% + 1
            'MUESTRADATOS3.MSF_2.Rows = I%
            MUESTRADATOS3.GRID.Rows = MUESTRADATOS3.GRID.Rows + 1

            MUESTRADATOS3.GRID.TextMatrix(i%, 1) = NPROV
            MUESTRADATOS3.GRID.TextMatrix(i%, 2) = RASOPROV$
            MUESTRADATOS3.GRID.TextMatrix(i%, 3) = FECHATEX$(FE%)
            MUESTRADATOS3.GRID.TextMatrix(i%, 4) = NUMEDOCU$
            MUESTRADATOS3.GRID.TextMatrix(i%, 5) = FORMATNUM$(IMPORTE1#, "F12.2")
            

            .MoveNext
          Loop
       End With
       RS.Close
       Set RS = Nothing
       MUESTRADATOS3.Left = MUESTRADATOS2.Left + 500
       MUESTRADATOS3.Top = MUESTRADATOS2.Top + 500
       MUESTRADATOS3.Show 1
    End If

End Sub

Private Sub MSF_2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If Button = 2 Then
    '
    i& = MSF_2.MouseRow
    J& = MSF_2.MouseCol
    If i& = 0 Or i& > MSF_2.Rows Then Exit Sub
    '
    On Error Resume Next
    MSF_2.Row = i&
    MSF_2.COL = J&
    On Error Resume Next

    NPX& = TRIM$(MSF_2.TextMatrix(i&, 1))
    If NPX& < 1 Then
        Exit Sub
    End If

    If Me.Caption = "Pedidos" Then
        '
        NUDOCU$ = TRIM$(Str$(NPX&))
        TIDOCUM$ = "Pedido de Cliente"
        While Len(NUDOCU$) < 6: NUDOCU$ = "0" + NUDOCU$: Wend
        NUDOCU$ = NUDOCU$ + " "
        Screen.MousePointer = 11
        '
        'BUSCA EL DOCUMENTO EN LA TABLA SQL CON LA NUEVA FUNCION
        DOCUNU& = DOCID&(TIDOCUM$, Str$(NPX&), NUDOCU$)
'        If DOCUNU& < 1 Then
'             Screen.MousePointer = 1
'             Call COMUNI("No se Encontró el Documento Impreso\Correspondiente a la " & DOCUMENTO_ELEGIDO$ & " .\  \Si se Trata de un Pedido Pendiente de Entrega\Puede Refrescarlo con la Opción\'Modificación de Pedido en Curso'.")
'             On Error Resume Next
'             OPPED07.Command46.SetFocus
'             OPPED07.Show 1
'             If Err Then
'              On Error GoTo 0
'             End If
'             Exit Sub
'        End If
        Call MUESTRADOC(DOCUNU&, "", TIDOCUM$, Str$(NPX&), NUDOCU$)
    Else
        TIDOCUM$ = "Nota de Pedido"
        NUDOCU$ = TRIM$(Str$(NPX&))
        DOCUNU& = 0
        '
        Call MUESTRADOC(DOCUNU&, "", TIDOCUM$, NUDOCU$)
    End If
    Screen.MousePointer = 1
  End If
End Sub

Private Sub MSF2_Click()
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

Private Sub MSF2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If Button = 2 Then
    '
    i& = MSF2.MouseRow
    J& = MSF2.MouseCol
    If i& = 0 Or i& > MSF2.Rows Then Exit Sub
    '
    On Error Resume Next
    MSF2.Row = i&
    MSF2.COL = J&
    On Error Resume Next

    NPX& = TRIM$(MSF2.TextMatrix(i&, 1))
    If NPX& < 1 Then
        Exit Sub
    End If

    If Me.Caption = "Pedidos" Then
        '
        NUDOCU$ = TRIM$(Str$(NPX&))
        TIDOCUM$ = "Pedido de Cliente"
        While Len(NUDOCU$) < 6: NUDOCU$ = "0" + NUDOCU$: Wend
        NUDOCU$ = NUDOCU$ + " "
        Screen.MousePointer = 11
        '
        'BUSCA EL DOCUMENTO EN LA TABLA SQL CON LA NUEVA FUNCION
        DOCUNU& = DOCID&(TIDOCUM$, Str$(NPX&), NUDOCU$)
'        If DOCUNU& < 1 Then
'             Screen.MousePointer = 1
'             Call COMUNI("No se Encontró el Documento Impreso\Correspondiente a la " & DOCUMENTO_ELEGIDO$ & " .\  \Si se Trata de un Pedido Pendiente de Entrega\Puede Refrescarlo con la Opción\'Modificación de Pedido en Curso'.")
'             On Error Resume Next
'             OPPED07.Command46.SetFocus
'             OPPED07.Show 1
'             If Err Then
'              On Error GoTo 0
'             End If
'             Exit Sub
'        End If
        Call MUESTRADOC(DOCUNU&, "", TIDOCUM$, Str$(NPX&), NUDOCU$)
    Else
        TIDOCUM$ = "Nota de Pedido"
        NUDOCU$ = TRIM$(Str$(NPX&))
        DOCUNU& = 0
        '
        Call MUESTRADOC(DOCUNU&, "", TIDOCUM$, NUDOCU$)
    End If
    Screen.MousePointer = 1
  End If
End Sub


Private Sub TXTBUSCAR_Change()
    '
    If TRIM$(TXTBUSCAR) = "" Then
      MSF_2.Visible = False
      GRID.Visible = True
      Exit Sub
    End If
    '
    MSF_2.Rows = 2
    For i& = 1 To MSF_2.Cols - 1
       MSF_2.TextMatrix(1, i&) = ""
    Next i&
    REG& = 0
    MSF_2.Visible = True
    MSF_2.ZOrder 0
    MSF_2.Cols = GRID.Cols
    MSF_2.Redraw = False
    GRID.Visible = False
    '
    'RECORRO EL MSF QUE TIENE TODO E IMPRIMO EN msf  LOS QUE MACHEAN
    For i& = 1 To GRID.Rows - 1
      For J& = 1 To GRID.Cols - 1
        If (InStr(UCase$(GRID.TextMatrix(i&, J&)), UCase$(TXTBUSCAR))) Then
          If (TRIM$(CMBFILTRO.TEXT)) = "" Or TRIM$(UCase$(CMBFILTRO)) = TRIM$(UCase$(GRID.TextMatrix(0, J&))) Then
            REG& = REG& + 1
            MSF_2.Rows = REG& + 1
            For H& = 1 To MSF_2.Cols - 1
              MSF_2.TextMatrix(REG&, H&) = GRID.TextMatrix(i&, H&)
            Next H&
          End If
          If (TRIM$(CMBFILTRO.TEXT)) = "" Then Exit For
        End If
      Next J&
    Next i&
    '
    MSF_2.Redraw = True
    
    '
End Sub

Private Sub TXTBUSCAR_KeyPress(KeyAscii As Integer)
   If KeyAscii = 27 Then
      TXTBUSCAR.TEXT = ""
   End If
End Sub
'
Sub GENEXCEL(NombreDeArchivo$, GRID As msFlexGrid)
   Dim XLAPP As Excel.Application
   Dim XLBOOK As Excel.Workbook
   Dim XLSHEET As Excel.Worksheet
   '
   Dim FECHA As Date
   Dim valor As Double
   Dim TEXTO As String
   '
   On Error GoTo 20
   Set XLAPP = New Excel.Application
   Set XLBOOK = XLAPP.Workbooks.Add
   Set XLSHEET = XLBOOK.Worksheets.Add
   On Error GoTo 0
   '
   TipoDato% = 0: COLUM& = 0
   For i& = 1 To GRID.Rows
       RENG1& = i&: COLUM& = 0: CampoConEco% = 0: ColumnaFormato = 0
       For COLU1& = 1 To GRID.Cols
          If GRID.ColWidth(COLU1& - 1) = 0 Then GoTo 10 ' OBVIA LAS COLUMNAS OCULTAS
          COLUM& = COLUM& + 1
          ColumnaFormato = ColumnaFormato + 1
          '
          VALORX = GRID.TextMatrix(i& - 1, COLU1 - 1)
          If Len(VALORX) = 8 And Mid$(VALORX, 3, 1) = "/" And Mid$(VALORX, 6, 1) = "/" Then
             FEXU$ = VALORX
             VALORX = FETEXCOR1$(Int(FECHANUM(FEXU$)))
             TipoDato% = 1
          End If
          ' ESTO ES PARA QUE EXPORTE COMO CERO Y NO ESPACIOS
          If CampoConEco% = 0 Then
            FFCC$ = UCase$(Left$(FORCAMPO$(ColumnaFormato), 1))
          Else
            ColumnaFormato = ColumnaFormato - 1
            FFCC$ = UCase$(Left$(FORCAMPO$(ColumnaFormato), 1))
            CampoConEco% = 0
          End If
          If FFCC$ = "I" Or FFCC$ = "F" Or FFCC$ = "G" Or FFCC$ = "$" Then
             If TRIM$(CStr(VALORX)) = "" Then
                VALORX = 0
             End If
             TipoDato% = 2
          End If
          If FFCC$ = "C" Or FFCC$ = "A" Then VALORX = "'" & VALORX
          '
          GRID.Row = i& - 1: GRID.COL = 1
          If GRID.CellFontUnderline = True Then
              XLSHEET.Cells(RENG1&, COLUM&).Font.Underline = True
          End If

          XLSHEET.Cells(RENG1&, COLUM&) = VALORX
          If TipoDato% = 1 Then XLSHEET.Cells(RENG1&, COLUM&).NumberFormat = "dd/mm/yyyy"
          If TipoDato% = 2 Then XLSHEET.Cells(RENG1&, COLU1&).NumberFormat = "#,##0.00"
          'If TIPODATO% = 0 Then XLSHEET.Cells(RENG1&, COLUM&).NumberFormat = "@"
          TipoDato% = 0
          '
          If TRIM$(TABLAECO(COLU1&)) <> "" Then CampoConEco% = 1
          '
10     Next COLU1&
   Next i&
   '
   Close #N1%
   On Error GoTo 30
   XLSHEET.SaveAs NombreDeArchivo$
   On Error GoTo 0
   '
   XLAPP.Quit      ' cierra EXCEL que se esta ejecutando minimizado
   Set XLAPP = Nothing     ') libera memoria
   Set XLBOOK = Nothing    ') utilizada por
   Set XLSHEET = Nothing   ') las variables objeto
   '
   Screen.MousePointer = 1
   Call COMUNI("Planilla Excel Creada Exitosamente:\ \" + REPLACAR$(NombreDeArchivo$, "\", "/"))
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
