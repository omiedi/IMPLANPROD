VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form MUESTRA17 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   7875
   ClientLeft      =   45
   ClientTop       =   585
   ClientWidth     =   16095
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7875
   ScaleWidth      =   16095
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command4 
      Caption         =   "Salir"
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
      Left            =   14760
      TabIndex        =   15
      ToolTipText     =   "Asigna Cuit a los Comprobantes Seleccionados"
      Top             =   7320
      Width           =   1125
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Grabar Solicitudes de Facturación"
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
      Left            =   11640
      TabIndex        =   14
      Top             =   7320
      Width           =   2925
   End
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
      Caption         =   "Exp. a Excel"
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
      Left            =   7800
      TabIndex        =   8
      Top             =   7320
      Width           =   1125
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
      Left            =   7920
      Picture         =   "MUESTRA17.frx":0000
      TabIndex        =   7
      ToolTipText     =   "Emite Listado"
      Top             =   7320
      Visible         =   0   'False
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
      OleObjectBlob   =   "MUESTRA17.frx":030A
      Top             =   0
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   225
      Index           =   0
      Left            =   120
      OleObjectBlob   =   "MUESTRA17.frx":053E
      TabIndex        =   1
      Top             =   240
      Visible         =   0   'False
      Width           =   1485
   End
   Begin MSFlexGridLib.MSFlexGrid GRID 
      Height          =   7125
      Left            =   120
      TabIndex        =   2
      ToolTipText     =   "Un Click Sobre Encabezado Para Ordenar/Click en el Detalle Para Seleccionar"
      Top             =   120
      Width           =   15795
      _ExtentX        =   27861
      _ExtentY        =   12568
      _Version        =   393216
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
      Height          =   7125
      Left            =   120
      TabIndex        =   4
      ToolTipText     =   "Un Click Sobre Encabezado Para Ordenar/Click  Sobre el Detalle Para Seleccionar"
      Top             =   120
      Visible         =   0   'False
      Width           =   15795
      _ExtentX        =   27861
      _ExtentY        =   12568
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
   Begin VB.Menu mnuarchivo 
      Caption         =   "Archivo"
   End
   Begin VB.Menu mnuedicion 
      Caption         =   "Edición"
      Begin VB.Menu revertirseleccion 
         Caption         =   "Revertir Selección"
      End
   End
End
Attribute VB_Name = "MUESTRA17"
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
Public HaySeleccion% '0 no hay ningun items seleccionado, 1 hay un item o mas seleccionados
Public MSFACTIVO As MSFlexGrid
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
      Call GENEXCEL(AREXPO$, MUESTRA17.MSF_2)
     Else
      Call GENEXCEL(AREXPO$, MUESTRA17.GRID)
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
        ALTOFILA& = 0
        For FILA% = 1 To QFILAS% - 1
            ALTOFILA& = ALTOFILA& + GRID.RowHeight(FILA%)
        Next FILA%
        ALTOSCR& = Screen.Height
        If ALTOFILA& + 1600 > ALTOSCR& Then
            Me.Top = 0: Me.Height = ALTOSCR& - 600
            GRID.Height = ALTOSCR& - 1600
        Else
            TOPFORM& = Int((ALTOSCR& - (ALTOFILA& + 1600)) / 2)
            Me.Top = TOPFORM&: Me.Height = ALTOFILA& + 1600
            GRID.Height = ALTOFILA& + 600
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
Sub GENEXCEL(NOMBREDEARCHIVO$, GRID As MSFlexGrid)
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
   TIPODATO% = 0: COLUM& = 0
   For i& = 1 To GRID.Rows
       RENG1& = i&: COLUM& = 0
       For COLU1& = 1 To GRID.Cols
          If GRID.ColWidth(COLU1& - 1) = 0 Then GoTo 10 ' OBVIA LAS COLUMNAS OCULTAS
          COLUM& = COLUM& + 1
          '
          VALORX = GRID.TextMatrix(i& - 1, COLU1 - 1)
          If Len(VALORX) = 8 And Mid$(VALORX, 3, 1) = "/" And Mid$(VALORX, 6, 1) = "/" Then
             FEXU$ = VALORX
             'VALORX = FETEXCOR1$(Int(FECHANUM(FEXU$)))
             VALORX = FETEXCOR2$(Int(FECHANUM(FEXU$)))
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
          'PARA QUE NO REPITA ENCABEZADO
'          If RENG1& = 1 Then
'            If VALORX = VALORX_ANT And TRIM$(VALORX) <> "" Then
'              If COLUM& <= 17 Then
'               LETRAX1$ = Chr$(64 + COLUM& - 1)
'               LETRAX2$ = Chr$(64 + COLUM&)
'               RANXO$ = LETRAX1$ & TRIM$(Str$(RENG1&)) & ":" & LETRAX2$ & TRIM$(Str$(RENG1&))
'               Range(RANXO$).Select
'
'               XLSHEET.Cells.HorizontalAlignment = xlCenter
'               XLSHEET.Cells.HorizontalAlignment = xlCenter
'               XLSHEET.Cells.VerticalAlignment = xlCenter
'               XLSHEET.Cells.WrapText = False
'               XLSHEET.Cells.Orientation = 0
'               XLSHEET.Cells.AddIndent = False
'               XLSHEET.Cells.IndentLevel = 0
'               XLSHEET.Cells.ShrinkToFit = False
'               XLSHEET.Cells.ReadingOrder = xlContext
'               XLSHEET.Cells.MergeCells = True
'               GoTo 10
'             End If
'            End If
'          End If
          '
          XLSHEET.Cells(RENG1&, COLUM&) = VALORX
          If TIPODATO% = 1 Then XLSHEET.Cells(RENG1&, COLUM&).NumberFormat = "dd/mm/yyyy"
          'If TIPODATO% = 2 Then XLSHEET.Cells(RENG1&, COLU1&).NumberFormat = "#,##0.00"
          If TIPODATO% = 0 Then XLSHEET.Cells(RENG1&, COLUM&).NumberFormat = "@"
          TIPODATO% = 0
          VALORX_ANT = VALORX
10     Next COLU1&
   Next i&
   '
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
Private Sub Command3_Click()
         Command3.Enabled = False
        'SETEO EL MSF QUE ESTE VISIBLE
        If GRID.Visible = True Then
          Set MSFACTIVO = GRID
        Else
          Set MSFACTIVO = MSF_2
        End If
        '
        Seleccionados% = 0
        MSFACTIVO.Row = 0
        MSFACTIVO.COL = 2
        MSFACTIVO.Sort = 2
        HaySeleccion% = 0
        For J& = 1 To MSFACTIVO.Rows - 1
           If MSFACTIVO.TextMatrix(J&, 0) = "*" Then
              FE% = CVINT(MUESTRA17.GRID.TextMatrix(J&, 1))
              NC = XVALO(MUESTRA17.GRID.TextMatrix(J&, 2))
              RASYDOMI$ = TRIM$(MUESTRA17.GRID.TextMatrix(J&, 3))
              CI1% = XVALO(MUESTRA17.GRID.TextMatrix(J&, 4))
              DESCRITRABAJO$ = MUESTRA17.GRID.TextMatrix(J&, 5)
              OPERARIO% = XVALO(MUESTRA17.GRID.TextMatrix(J&, 6))
              NAMEOPERARIO$ = TRIM$(MUESTRA17.GRID.TextMatrix(J&, 7))
              CANCUOTAS = XVALO(MUESTRA17.GRID.TextMatrix(J&, 8))
              ID = XVALO(MUESTRA17.GRID.TextMatrix(J&, 9))
              Call GENERAPEDIDO(FE%, NC, RASYDOMI$, CI1%, DESCRITRABAJO$, OPERARIO%, NAMEOPERARIO$, CANCUOTAS, ID, FECHAREALIZO, RETORNO_PEDIDO)
              
              Seleccionados% = 1
           End If
        Next J&
        '
        If Seleccionados% > 0 Then Call Agenda17.Command10_Click 'boton filtrar para actualizar
99      Command3.Enabled = True
        Hide
End Sub

Sub GENERAPEDIDO(FE%, NC, RASYDOMI$, CIXI%, DESCRITRABAJO$, OPERARIO%, NAMEOPERARIO$, CANCUOTAS, ID, FECHAREALIZO, RETORNO_PEDIDO)

   OKX% = 0
   RETORNO_PEDIDO = 0
   Call ABREPEDCLI
   '
   Screen.MousePointer = 11
    APREC$ = "LIPRE" + "001"
    PRLST# = PRELISTA#(APREC$, CIXI%)
'    FEMI% = HOY(HO$)
    FEMI% = FE%
    FEX$ = FECHATEX$(FEMI%)
    i% = 0
    
    For i% = 1 To CANCUOTAS
       Set RS = New ADODB.Recordset
       SQLX$ = "SELECT * FROM  PEDENCA  "
       SQLX$ = SQLX$ + " WHERE NroClie = -1"
'      Set RS = Ventas.OpenRecordset(SQLX$, dbOpenDynaset)
25     On Error Resume Next
       RS.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
       If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
              On Error GoTo 0
              Call CapturaErrorOpen
              GoTo 25
           End If
           On Error GoTo 0
       With RS
         .AddNew
         NroPedi1& = NUPEDIDO& ' FUNCION
         !NROPED = NroPedi1&
            If i% = 1 Then
                MEX2$ = Mid$(FEX$, 4, 2)
                ANOX$ = Right(FEX$, 2)
            Else
                
                MEX1% = 1 + XVALO(MEX2$)
                MEX2$ = TRIM$(CStr(MEX1%))
                If MEX1% > 12 Then
                  MEX1% = 1
                  MEX2$ = TRIM$(CStr(MEX1%))
                  ano% = 1 + XVALO(ANOX$)
                  ANOX$ = TRIM$(CStr(ano%))
                  '
                End If
            End If
            ANOX$ = "00" + ANOX$: ANOX$ = Right(ANOX$, 2)
            MEX2$ = "00" + MEX2$: MEX2$ = Right$(MEX2$, 2)
            FEX1$ = "01/" + MEX2$ + "/" + ANOX$
            'MESYAÑOTX$ = METEXTO$(FECHANUM(FEX1$)) 'SI QUIERE QUE VAYA EL MES QUE CORRESPONDE LA CUOTA
            MESYAÑOTX$ = METEXTO$(FE%) 'SI QUIERE QUE VAYA EL  MES DE REALIZACION
         !FeSolEnt = CVDAT(CVINT(FEX1$))
         If CANCUOTAS > 1 Then
    '        TXTREFE$ = "Cuota de " & METEXTO$(CVINT(FEX1$)) & " " & I% & " de " & CANCUOTAS
    '         TXTREFE$ = "Cuota " & I% & " de " & CANCUOTAS
             TXTREFE$ = DESCRITRABAJO$ & " | " & MESYAÑOTX$ & " (" & i% & "/" & CANCUOTAS & ")"
         Else
    '        TXTREFE$ = "Pago de " & METEXTO$(CVINT(FEX1$))
             TXTREFE$ = DESCRITRABAJO$ & " | " & MESYAÑOTX$
         End If
         !REFERENCIA = Left$(TXTREFE$, 128)
         !FECHEMIS = CVDAT(FEMI%)
         !NROCLIE = NC
         !RaSoClie = rasocli$(NC)
         !NROOCOM = ""
         !DOMIENT = ""
         !LOCAENT = ""
         !TranspEnt = ""
         !NTranspo = 0
         !NSucent = 0
         !Observa = ""
    
    '     !FESOLENT = Now
           '\\\OT-05-0236-WAR-02/05/05///
         !Status = 0              ' Status  de Aprobación
           '\\\OT-05-0236-FIN///
         !ListPre = 1
         !CondPag = 0
         !VENDED = 0
         !TDescApli = ""
         .Update
       End With
       '
       RS.Close
       Set RS = Nothing
       
       ANOTA$ = ""
       TTXYZ$ = Space$(76)
       HOII% = HOY(HOS$)
       Call REPLA(TTXYZ$, HOS$, 1, 8)
       Call REPLA(TTXYZ$, "Generación de Orden de Venta:", 11, 56)
       Call REPLA(TTXYZ$, USERTER$, 69, 24)
       ANOTA$ = ANOTA$ + TTXYZ$ + Chr$(10) + Chr$(13)
       IKITEM% = 0
       
       '
    '   FENTRESOL% , CIXI%, CODEXT$(CIXI%), DESCRIT0$(CIXI%), PreList, PREUNID, ImpoTot, CANTID
       SQLX$ = "SELECT * FROM  PEDDETA  "
       SQLX$ = SQLX$ + " WHERE NroClie = -1"
'       Set RS1 = Ventas.OpenRecordset(SQLX$, dbOpenDynaset)
       Set RS1 = New ADODB.Recordset
30     On Error Resume Next
       RS1.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
       If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
              On Error GoTo 0
              Call CapturaErrorOpen
              GoTo 30
       End If
       On Error GoTo 0
       With RS1
             .AddNew
               IKITEM% = IKITEM% + 1
               !NROPED = NroPedi1&
               !FECHEMIS = CVDAT(FEMI%)
               !NROCLIE = NC
               !RaSoClie = rasocli$(NC)
               !NROOCOM = ""
               
    '            If I% = 1 Then
    '                MEX2$ = Mid$(FEX$, 4, 2)
    '                ANOX$ = Right(FEX$, 2)
    '
    '            Else
    '
    '                MEX1% = 1 + XVALO(MEX2$)
    '                MEX2$ = TRIM$(CStr(MEX1%))
    '                If MEX1% > 12 Then
    '                  MEX1% = 1
    '                  MEX2$ = TRIM$(CStr(MEX1%))
    '                  ano% = 1 + XVALO(ANOX$)
    '                  ANOX$ = TRIM$(CStr(ano%))
    '                  '
    '                End If
    '            End If
    '            ANOX$ = "00" + ANOX$: ANOX$ = Right(ANOX$, 2)
    '            MEX2$ = "00" + MEX2$: MEX2$ = Right$(MEX2$, 2)
    '            FEX1$ = "01/" + MEX2$ + "/" + ANOX$
                !FeSolEnt = CVDAT(CVINT(FEX1$)) ' LO GENERA GUANDO GRABA EL PEDENCA
                 'AQUI RESOLVER LA FECHA DE ENTREGA
    '             FENTRESOL% = CVI(Mid$(X$, 63, 2))
    '           !FESOLENT = CVDAT(FENTRESOL%)
               !Status = 0              ' Status  de Aprobación
               !ListPre = 0
               !CondPag = 0
               !VENDED = 0
               !PorComVend = 0
               !NuOrItem = IKITEM%
               !CODIINT = CIXI%
               !CODIEXT = CODEXT$(CIXI%)
               !SubCod = ""
               !Descrip = DESCRIT0$(CIXI%)
               !TLargo = 0
               !TAncho = 0
               !TEspesor = 0
               !CBase = ""
               !TTALLE = ""
               !TCOLOR = ""
               !VenRep = ""   ' Venta o Reparacion
               !MoneEmis = 1
               PREUNI# = PRLST# / CANCUOTAS
               !TIPOCAM = 1
               !PreList = PREUNI#
               AA = FORMATNUM$(PREUNI#, "F12.2")
               !DescLista = 0
               !DescCombi = ""
               !DescRubro = 0
               !DescTotal = 0
               !CanAsig = 0
               !CanPed = 1
               !PreUnid = PREUNI#
               !ImpoTot = PREUNI#
               
               !CANTENT = 0
               !CanSaldo = 1
               !FeUltEnt = CVDAT(0)
               !NroRemi = ""
               !TPedItem = ""
               '
               'VER ESTO ???????
               TTXYZ$ = Space$(76)
               Call REPLA(TTXYZ$, "- " + CODEXT$(CIXI%), 11, 18)
               Call REPLA(TTXYZ$, DESCRIT0$(CIXI%), 29, 38)
               Call REPLA(TTXYZ$, CSTRING$(MKS$(1), 3, 10, 1, 2), 57, 10)
               ANOTA$ = ANOTA$ + TTXYZ$ + Chr$(10) + Chr$(13)
               TTXYZ$ = Space$(76)
               Call REPLA(TTXYZ$, "Cuota de " & METEXTO$(FEMI%) & " " & i% & " de " & CANCUOTAS, 11, 56)
               ANOTA$ = ANOTA$ + TTXYZ$ + Chr$(10) + Chr$(13)
               '
             .Update
           End With
           '
           RS1.Close
           Set RS1 = Nothing
       '
            ANOTA$ = ANOTA$ + Space$(10) + String$(56, "-") + Chr$(10) + Chr$(13)
    '        Call FRATEXTO(Text13.TEXT, 32)
    '        For KKI% = 1 To CARETEX%
    '          TTXX$ = Space$(92)
    '          Call REPLA(TTXX$, RETEX$(KKI%), 26, 32)
    '          If KKI% = 1 Then Call REPLA(TTXX$, "Observaciones:", 11, 15)
    '          ANOTA$ = ANOTA$ + TTXX$ + Chr$(10) + Chr$(13)
    '        Next KKI%
    '        ANOTA$ = ANOTA$ + String$(88, "=") + Chr$(10) + Chr$(13)
            '
            Call APERBASDAT("PEDCLI")
            SQLX$ = "SELECT ANOTAPED FROM PEDENCA WITH(NOLOCK)"
            SQLX$ = SQLX$ + " WHERE NroPed = " & NroPedi1& & ""
            Set PEDDDENCA = READSET(SQLX$)
            ANOTB$ = SAFETEXT$(PEDDDENCA!AnotaPed)
            ANOTA$ = ANOTB$ + ANOTA$
            '\\\OT-08-0912-OD-30/09/08/// PARA EVITAR QUE SE GUARDEN DATOS CON COMILLAS SIMPLES
            ANOTA$ = REPLACAR$(ANOTA$, "'", "")
            '\\\OT-08-0912-OD-FIN///
            SQLX$ = "UPDATE PEDENCA SET AnotaPed=" & "'" & ANOTA$ & "' WHERE NroPed=" & NroPedi1&
            FLEXCONN.Execute (SQLX$)
       '
            SQLX$ = "SELECT PEDIDO,FACTURADO FROM  AGENDA  "
            SQLX$ = SQLX$ + " WHERE ID = " & XVALO(ID)
            'Set RS2 = Badaclie.OpenRecordset(SQLX$, dbOpenDynaset)
            Set RS2 = New ADODB.Recordset
35          On Error Resume Next
            RS2.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
            If Err Then  'VALIDACION DE ERROR POR SI GENERA ERROR POR TIEMPO DE ESPERA
              On Error GoTo 0
              Call CapturaErrorOpen
              GoTo 35
            End If
            On Error GoTo 0
            With RS2
               If .EOF = False Then
'                 .Edit
                 !FACTURADO = "3"
                 !PEDIDO = NUPEDIDO&
                 .Update
               End If
            End With
            RS2.Close
            Set RS2 = Nothing
       Next i%
   Call LIBARCH("*.*")
   Call BAJALDISCO
   '
   OKX% = 1

   '
98 Call LIBARCH("PEDENCA")
   '
   '
   Screen.MousePointer = 1
   '
End Sub





Private Sub Command4_Click()
   APROBO% = 0
   Unload Me
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  APROBO% = 0
End Sub

   '
Function NUPEDIDO&()
   '
   NPP& = 1
   Call ABREPEDCLI
   '
   Set PEDENCATEMP = New ADODB.Recordset
   SQLX$ = "SELECT max(NroPed) as NUMAXI FROM Pedenca "
   Set PEDENCATEMP = FLEXCONN.Execute(SQLX$)
   '
   NPP& = 1 + XVALO(PEDENCATEMP!NUMAXI)
   PEDENCATEMP.Close


   '
   NUPEDIDO& = NPP&
   '
End Function
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
    Else
        Call COLOREAR_GRILLA_SOLAFILA(GRID, vbYellow, GRID.Row, 0, 0, 0)
    End If
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
    Else
       Call COLOREAR_GRILLA_SOLAFILA(MSF_2, vbYellow, MSF_2.Row, 0, 0, 0)
    End If
End Sub

Private Sub revertirseleccion_Click()
        'SETEO EL MSF QUE ESTE VISIBLE
        If GRID.Visible = True Then
          Set MSFACTIVO = GRID
        Else
          Set MSFACTIVO = MSF_2
        End If
        '
        HaySeleccion% = 0
        Screen.MousePointer = 11
        For i& = 1 To MSFACTIVO.Rows - 1
           Call COLOREAR_GRILLA_SOLAFILA(MSFACTIVO, vbYellow, i&, 0, 0, 0)
        Next i&
        Screen.MousePointer = 1
        '
End Sub

'Private Sub MSF2_Click()
'    Static MODO  As Boolean
'    If (MSF2.MouseRow = 0) Then
'
'        ' Ordena en forma ascendente
'        If MODO Then
'            MSF2.COL = MSF2.MouseCol
'            MSF2.Sort = 2
'            MODO = False
'        ' Ordena en forma descendente
'        Else
'            MSF2.COL = MSF2.MouseCol
'            MSF2.Sort = 1
'            MODO = True
'        End If
'    End If
'
'End Sub



Private Sub TXTBUSCAR_Change()
    '
    If TRIM$(TXTBUSCAR) = "" Then
      MSF_2.Visible = False
      GRID.Visible = True
      Exit Sub
    End If
    '
    MSF_2.Rows = 1
'    For I& = 1 To MSF_2.Cols - 1
'       MSF_2.TextMatrix(0, I&) = ""
'    Next I&
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
            Call COLOREAR_GRILLA_SOLAFILA(MSF_2, vbWhite, REG&, 0, 0, 0)
            For H& = 1 To MSF_2.Cols - 1
              MSF_2.TextMatrix(REG&, 0) = ""
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
