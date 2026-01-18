VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.OCX"
Begin VB.Form MUESTRADATOS3 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Detalle de Composición de Deuda"
   ClientHeight    =   7875
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   18150
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7875
   ScaleWidth      =   18150
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Height          =   600
      Left            =   120
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
      Left            =   16440
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
      Left            =   14280
      Picture         =   "MUESTRADATOS3.frx":0000
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
      OleObjectBlob   =   "MUESTRADATOS3.frx":030A
      Top             =   0
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   225
      Index           =   0
      Left            =   120
      OleObjectBlob   =   "MUESTRADATOS3.frx":053E
      TabIndex        =   1
      Top             =   240
      Visible         =   0   'False
      Width           =   1485
   End
   Begin MSFlexGridLib.MSFlexGrid GRID 
      Height          =   7005
      Left            =   120
      TabIndex        =   2
      ToolTipText     =   "Un Click Sobre Encabezado Para Ordenar/Doble Click Para Pasar a Grilla Inferior/Click Boton Derecho Imagen Digital"
      Top             =   120
      Width           =   17940
      _ExtentX        =   31644
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
Attribute VB_Name = "MUESTRADATOS3"
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
      Call GENEXCEL(AREXPO$, MUESTRADATOS.MSF_2)
     Else
      Call GENEXCEL(AREXPO$, MUESTRADATOS.GRID)
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

Private Sub Label13_Change()
End Sub

Private Sub Label9_Change()
End Sub

Private Sub GRID_Click()
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

Private Sub GRID_DblClick()
   i& = GRID.MouseRow
   J& = GRID.MouseCol
   If i& < 1 Or i& > GRID.Rows Then Exit Sub
   '
   GRID.Row = i&
   GRID.COL = J&
   '
   OPX% = COMALTER%("Opciones de Información\\Ver Movimientos Contables\Ver Orden de Pago Imputada")
   If OPX% < 1 Or OPX% > 2 Then Exit Sub
   NUDOCU1$ = TRIM$(GRID.TextMatrix(GRID.Row, 4))
   FE% = FECHANUM(GRID.TextMatrix(GRID.Row, 3))
   NUPROV% = XVALO(GRID.TextMatrix(GRID.Row, 1))
   If OPX% = 1 Then
      Call MUEFACO1(NUDOCU1$, FE%, NUPROV%)
      Exit Sub
   Else
      '
      If UCase$(Left$(NUDOCU1$, 3)) = "OP-" Then GoTo 80 ' SI ES UNA ORDEN DE PAGO MUESTRA DIRECTAMENTE LA OP
      '
      JJ& = 0
      DAPLI$ = NUDOCU1$
      Dim Aplipago As ADODB.Recordset
      strsql = "Select DISTINCT CoDocAp_Lar, RefeApli, ImpoApli from Aplipago where NuPROVE = " & NUPROV%
      strsql = strsql & "And  CoDocOr_LAR = '" & TRIM(DAPLI$) & "'"
      Set Aplipago = FLEXCONN.Execute(strsql)
      With Aplipago
         If Not (.EOF And .BOF) Then
            .MoveFirst
            Y$ = REGBLAN$("!SELDOCAP")
            Do While Not (.EOF)
               NROY$ = !codocap_lar
               REFY$ = Left(!refeapli, 18)
               If TRIM$(REFY$) = "" Then REFY$ = "." + REFY$
               IMPO# = !IMPOAPLI
               IMPOS$ = FORMATNUM$(IMPO#, "F13.2")
               '
               'CALCULO EL TOTAL
               IMPO# = 0
               Dim MIRS1 As ADODB.Recordset
               strsql1 = "Select DISTINCT ImpoApli from Aplipago"
               strsql1 = strsql1 & " where NuPROVE = " & NC%
               strsql1 = strsql1 & " And  CoDocAP_lar = '" & NROY$ & "' "
               Set MIRS1 = FLEXCONN.Execute(strsql1)
               With MIRS1
               If Not (.EOF And .BOF) Then
                  .MoveFirst
                  Do While Not (.EOF)
                   IMPO# = IMPO# + XVALO(!IMPOAPLI)
                   .MoveNext
                  Loop
               End If
               End With
               MIRS1.Close
               Set MIRS1 = Nothing
               '
               Y$ = Space$(60)
               Call REPLA(Y$, NROY$, 1, 18)
               Call REPLA(Y$, REFY$, 19, 24)
               Call REPLA(Y$, IMPOS$, 44, 13)
               JJ& = JJ& + 1
               Call GRAREG("!SELDOCAP", Y$, JJ&)
            .MoveNext
            Loop
         End If
         '
      End With
      Aplipago.Close
      Set Aplipago = Nothing
      '
      Call SETULTREG("!SELDOCAP", JJ&)
      Call FRESHECHO("!SELDOCAP")
      '
      Call SELECHO("!SELDOCAP")
      NUDOCU1$ = TRIM$(VALACT1$("!SELDOCAP"))
      If NUDOCU1$ = "" Then Exit Sub
80     TIDOCUM$ = "Orden de Pago"
       NPX& = XVALO(Mid$(NUDOCU1$, 11, 8))
       NUDOCU$ = "" + TRIM$(Str$(NPX&)) ' + " -"
       Screen.MousePointer = 11
        '
        DOCUNU& = DOCID&(TIDOCUM$, Str$(NPX&), NUDOCU$)
        If DOCUNU& < 1 Then
              Screen.MousePointer = 1
              Call COMUNI("Imposible Mostrar Documento")
              Exit Sub
        End If
        Call MUESTRADOC(DOCUNU&, "", TIDOCUM$, Str$(NPX&), NUDOCU$)
   End If
   '
End Sub

Sub MUEFACO1(TIDOCUM$, FE%, NUPROV%)
   '
   TIPODOCU$ = TRIM$(Left$(TIDOCUM$, 4)) 'FC-A
   NUDOCOR$ = TRIM$(TIDOCUM$) 'NUMERO DE DOC COMP
   NP1& = NUPROV%
   FECDOC1 = FETEXCOR1$(FE%) 'FECHA DEL DOC.
   '
   Call APERBASDAT("CABAN")
   Dim SHOFACOTMP As New ADODB.Recordset
   SQLZ$ = " WHERE FechEmisi = '" & FECDOC1 & "' "
   SQLZ$ = SQLZ$ + " AND CodiCom_Lar= '" & NUDOCOR$ & "' "
   SQLZ$ = SQLZ$ + " AND NuProve =  " & NP1& & " "
   SQLX$ = " SELECT * FROM CAJABANC " + SQLZ$
   Set SHOFACOTMP = FLEXCONN.Execute(SQLX$)
   '
   With SHOFACOTMP
     On Error Resume Next
     .MoveFirst
     If Err Then
        Call COMUNI("No se ha Podido Encontrar el Detalle del Movimiento")
        On Error GoTo 0
        Exit Sub
     End If
     On Error GoTo 0
     SHOWFACO.Caption = NUDOCOR$ + " del " + FECHATEX$(FE%) ' TIPODOCU$ + " " + !Codicom_lar + " del " + FECHDOCU$
     SHOWFACO.LIST1.Clear
     SHOWFACO.Label5 = ""
     SHOWFACO.Label7 = ""
     SHOWFACO.Label9 = ""
     SHOWFACO.Label11 = ""
     SHOWFACO.Label14 = ""
     '\\\OT-08-988-29/10/08-FL
     SHOWFACO.Label19 = ""
     '\\\OT-08-988-FIN-FL
     SHOWFACO.Label22 = SQLZ$
     SHOWFACO.Picture2.Visible = False
     NICLA% = 0
     '
     SUDEBE1# = 0: SUHABE1# = 0
     Do While .EOF = False
        QQQ1$ = CODCUE$(!CUECONTAI)
        QQQ2$ = "": If IsNull(!CenCosto) = False Then QQQ2$ = !CenCosto
        REFE1$ = !DESCRIMOV
        IDEBE1# = !Idebecomp
        IHABE1# = !IHABECOMP
        IDEBE2$ = FORMATNUM$(IDEBE1#, "F12.2")
        IHABE2$ = FORMATNUM$(IHABE1#, "F12.2")
        SUDEBE1# = SUDEBE1# + XVALO(IDEBE2$)
        SUHABE1# = SUHABE1# + XVALO(IHABE2$)
        ADDLIS1$ = AJUSTI$(QQQ1$, 14) + AJUSTI$(QQQ2$, 10) + AJUSTI$(REFE1$, 40) & IDEBE2$ & IHABE2$
        If (Abs(IDEBE1#) + Abs(IHABE1#)) >= 0.005 Then
          SHOWFACO.LIST1.AddItem ADDLIS1$
        End If
        If !Niv_Clasi > NICLA% Then NICLA% = !Niv_Clasi
        '
        If UCase$(TRIM$(!TipoMovim)) = "FCOM" Then
          MEMI% = 1: TICA = 1
          On Error Resume Next
            MEMI% = !Mone_Emis
            TICA = !Tipo_Cam
              If (TICA < 0.1 Or IsNull(TICA) = True Or isEmpty(TICA) = True) Then TICA = 1
          On Error GoTo 0
          SHOWFACO.Label5 = DEMONECO$(MEMI%)
          IDOLA# = (!IHABECOMP - !Idebecomp) / TICA
          SHOWFACO.Label7 = TRIM$(FORMATNUM$(IDOLA#, "F12.2"))
          IACTU# = IDOLA# * TICAMONE(MEMI%)
          SHOWFACO.Label9 = TRIM$(FORMATNUM$(IACTU#, "F12.2"))
          DICAM# = IACTU# - (!IHABECOMP - !Idebecomp)
          SHOWFACO.Label11 = TRIM$(FORMATNUM$(DICAM#, "F12.2"))
          SHOWFACO.Label12 = TRIM$(FORMATNUM$(TICA, "F10.4"))
          SHOWFACO.Label8.ToolTipText = "T.C.: " + TRIM$(FORMATNUM$(TICAMONE(MEMI%), "F10.4"))
          SHOWFACO.Label9.ToolTipText = "T.C.: " + TRIM$(FORMATNUM$(TICAMONE(MEMI%), "F10.4"))
          SHOWFACO.Label14 = !OBSERGEN
          '\\\OT-08-988-29/10/08-FL
          'SE CAPTURA UN POSIBLE ERROR PRODUCTO QUE EL CAMPO IDENALTA CONTENGA INFORMACION INVALIDA
          On Error Resume Next
          USNBRX% = Asc(Mid$(!IdenAlta, 5, 1))
          On Error GoTo 0
          SHOWFACO.Label19 = TRIM$(USERNAM$(USNBRX%))
          '\\\OT-08-988-FIN-FL
        End If
        '
     .MoveNext
     Loop
   End With
   SHOFACOTMP.Close
   Set SHOFACOTMP = Nothing
   '
   If NICLA% > 0 Then
      SHOWFACO.Label21 = "<" + TRIM$(Str$(NICLA%))
      SHOWFACO.Picture2.Visible = True
   End If
   '
   SHOWFACO.Label2 = TRIM$(FORMATNUM$(SUDEBE1#, "F12.2"))
   SHOWFACO.Label3 = TRIM$(FORMATNUM$(SUHABE1#, "F12.2"))
   SHOWFACO.Show 1
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




Private Sub MSF_2_DblClick()
   i& = MSF_2.MouseRow
   J& = MSF_2.MouseCol
   If i& = 0 Or i& > MSF_2.Rows Then Exit Sub
   '
   MSF_2.Row = i&
   MSF_2.COL = J&
   '
   OPX% = COMALTER%("Opciones de Información\\Ver Movimientos Contables\Ver Orden de Pago Imputada")
   If OPX% < 1 Or OPX% > 2 Then Exit Sub
   NUDOCU1$ = TRIM$(MSF_2.TextMatrix(MSF_2.Row, 4))
   FE% = FECHANUM(MSF_2.TextMatrix(MSF_2.Row, 3))
   NUPROV% = XVALO(MSF_2.TextMatrix(MSF_2.Row, 1))
   If OPX% = 1 Then
      Call MUEFACO1(NUDOCU1$, FE%, NUPROV%)
      Exit Sub
   Else
      '
      JJ& = 0
      DAPLI$ = NUDOCU1$
      'jandy sql esto deberia ser select sum()....group by
      Dim Aplipago As ADODB.Recordset
      strsql = "Select DISTINCT CoDocAp_Lar, RefeApli, ImpoApli from Aplipago where NuPROVE = " & NUPROV%
      strsql = strsql & "And  CoDocOr_LAR = '" & TRIM(DAPLI$) & "'"
      Set Aplipago = FLEXCONN.Execute(strsql)
      With Aplipago
         If Not (.EOF And .BOF) Then
            .MoveFirst
            Y$ = REGBLAN$("!SELDOCAP")
            Do While Not (.EOF)
               NROY$ = !codocap_lar
               REFY$ = Left(!refeapli, 18)
               If TRIM$(REFY$) = "" Then REFY$ = "." + REFY$
               IMPO# = !IMPOAPLI
               IMPOS$ = FORMATNUM$(IMPO#, "F13.2")
               '
               'CALCULO EL TOTAL
               IMPO# = 0
               Dim MIRS1 As ADODB.Recordset
               strsql1 = "Select DISTINCT ImpoApli from Aplipago"
               strsql1 = strsql1 & " where NuPROVE = " & NC%
               strsql1 = strsql1 & " And  CoDocAP_lar = '" & NROY$ & "' "
               Set MIRS1 = FLEXCONN.Execute(strsql1)
               With MIRS1
               If Not (.EOF And .BOF) Then
                  .MoveFirst
                  Do While Not (.EOF)
                   IMPO# = IMPO# + XVALO(!IMPOAPLI)
                   .MoveNext
                  Loop
               End If
               End With
               MIRS1.Close
               Set MIRS1 = Nothing
               '
               Y$ = Space$(60)
               Call REPLA(Y$, NROY$, 1, 18)
               Call REPLA(Y$, REFY$, 19, 24)
               Call REPLA(Y$, IMPOS$, 44, 13)
               JJ& = JJ& + 1
               Call GRAREG("!SELDOCAP", Y$, JJ&)
            .MoveNext
            Loop
         End If
         '
      End With
      Aplipago.Close
      Set Aplipago = Nothing
      '
      Call SETULTREG("!SELDOCAP", JJ&)
      Call FRESHECHO("!SELDOCAP")
      '
      Call SELECHO("!SELDOCAP")
      NUDOCU1$ = TRIM$(VALACT1$("!SELDOCAP"))
      If NUDOCU1$ = "" Then Exit Sub
       TIDOCUM$ = "Orden de Pago"
       NPX& = XVALO(Mid$(NUDOCU1$, 11, 8))
       NUDOCU$ = " " + TRIM$(Str$(NPX&)) + " -"
       Screen.MousePointer = 11
        '
        DOCUNU& = DOCID&(TIDOCUM$, Str$(NPX&), NUDOCU$)
        If DOCUNU& < 1 Then
              Screen.MousePointer = 1
              Call COMUNI("Imposible Mostrar Documento")
              Exit Sub
        End If
        Call MUESTRADOC(DOCUNU&, "", TIDOCUM$, Str$(NPX&), NUDOCU$)
   End If
   '
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

          '
          XLSHEET.Cells(RENG1&, COLUM&) = VALORX
          If TipoDato% = 1 Then XLSHEET.Cells(RENG1&, COLUM&).NumberFormat = "dd/mm/yyyy"
          'If TIPODATO% = 2 Then XLSHEET.Cells(RENG1&, COLU1&).NumberFormat = "#,##0.00"
          If TipoDato% = 0 Then XLSHEET.Cells(RENG1&, COLUM&).NumberFormat = "@"
          TipoDato% = 0
          VALORX_ANT = VALORX
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

