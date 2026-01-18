VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form MUESTRADATOS12 
   BackColor       =   &H00E2D3C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   " "
   ClientHeight    =   7845
   ClientLeft      =   45
   ClientTop       =   585
   ClientWidth     =   15585
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   7845
   ScaleWidth      =   15585
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   WhatsThisHelp   =   -1  'True
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   30
      Left            =   4920
      TabIndex        =   14
      Top             =   2760
      Width           =   30
      _ExtentX        =   53
      _ExtentY        =   53
      _Version        =   393216
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Exporta a Excel"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   9360
      TabIndex        =   5
      Top             =   7320
      Width           =   1725
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Actualizar Costo s/Variación Inflación"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   13560
      TabIndex        =   13
      ToolTipText     =   "Asigna Cuit a los Comprobantes Seleccionados"
      Top             =   7320
      Width           =   1845
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Actualizar Costo s/Variación del Dólar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   11400
      TabIndex        =   12
      ToolTipText     =   "Asigna Cuit a los Comprobantes Seleccionados"
      Top             =   7320
      Width           =   1845
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Asignar C.Contable"
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
      Left            =   9075
      TabIndex        =   11
      ToolTipText     =   "Asigna Cuit a los Comprobantes Seleccionados"
      Top             =   7320
      Visible         =   0   'False
      Width           =   1125
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E2D3C0&
      Height          =   600
      Left            =   120
      TabIndex        =   6
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
         TabIndex        =   8
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
         TabIndex        =   7
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
         TabIndex        =   10
         Top             =   240
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
         Left            =   3480
         TabIndex        =   9
         Top             =   240
         Width           =   705
      End
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
      Picture         =   "MUESTRADATOS12.frx":0000
      TabIndex        =   4
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
   Begin MSFlexGridLib.MSFlexGrid GRID 
      Height          =   7245
      Left            =   120
      TabIndex        =   0
      ToolTipText     =   "Un Click Sobre Encabezado Para Ordenar/Doble Click Para Pasar a Grilla Inferior/Click Boton Derecho Imagen Digital"
      Top             =   0
      Width           =   15435
      _ExtentX        =   27226
      _ExtentY        =   12779
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
      TabIndex        =   2
      ToolTipText     =   "Un Click Sobre Encabezado Para Ordenar/Doble Click Para Pasar a Grilla Inferior/Click Boton Derecho Imagen Digital"
      Top             =   120
      Visible         =   0   'False
      Width           =   15435
      _ExtentX        =   27226
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
         Size            =   9
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
      TabIndex        =   3
      Top             =   120
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label TEPRUEBA 
      AutoSize        =   -1  'True
      Caption         =   "TEPRUEBA"
      Height          =   300
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Menu mnuArchivo 
      Caption         =   "Archivo"
      Begin VB.Menu Imprimir 
         Caption         =   "Imprimir"
      End
   End
   Begin VB.Menu mnuEdicion 
      Caption         =   "Edición"
      Begin VB.Menu mnuseleccion 
         Caption         =   "Seleccionar Todo / Revertir Selección"
      End
   End
End
Attribute VB_Name = "MUESTRADATOS12"
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
Public EstructuraDeCampos   As String

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
     Dim Grilla As MSFlexGrid
     If MSF_2.Visible = True Then
        Set Grilla = MUESTRADATOS12.MSF_2
     Else
        Set Grilla = MUESTRADATOS12.GRID
     End If
     EligioFila% = 0
     For I& = 1 To Grilla.Rows - 1
        If TRIM$(Grilla.TextMatrix(I&, 0)) = "*" Then
            EligioFila% = 1: Exit For
        End If
     Next I&
     If EligioFila% = 0 Then mnuseleccion_Click
     Call GeneraExcelDeGrillaSinDisenio(AREXPO$, Grilla, "ANALISIS DE COSTOS DESACTUALIZADOS", True)
     If EligioFila% = 0 Then mnuseleccion_Click
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
Private Sub Command4_Click()

   APROBO% = 0
   FechaHoy = CVDAT(HOY(HO$))
   
   Dim Grilla As MSFlexGrid
   If MSF_2.Rows > 1 Then Set Grilla = MSF_2
   
   For IFila& = 1 To Grilla.Rows - 1
        
        If Grilla.TextMatrix(IFila&, 0) = "*" Then
            Codigo$ = TRIM$(Grilla.TextMatrix(IFila&, 1))
            CostoNuevo$ = TRIM$(FORMATNUM$(XVALO(Grilla.TextMatrix(IFila&, 7)), "F14.4"))
            Call ACTUREGISTRO("MASTER", "COSUNI", "CODIGO='" & Codigo & "'", CostoNuevo$, REGAF&)
            Call ACTUREGISTRO("MASTER", "FECOSTO", "CODIGO='" & Codigo & "'", FechaHoy, REGAF&)
        End If
   
   Next IFila&
   
   Unload Me
   
End Sub

Private Sub Command5_Click()

   APROBO% = 0
   FechaHoy = CVDAT(HOY(HO$))
   
   Dim Grilla As MSFlexGrid
   If MSF_2.Rows > 1 Then Set Grilla = MSF_2
   
   For IFila& = 1 To Grilla.Rows - 1
        
        If Grilla.TextMatrix(IFila&, 0) = "*" Then
            Codigo$ = TRIM$(Grilla.TextMatrix(IFila&, 1))
            CostoNuevo$ = TRIM$(FORMATNUM$(XVALO(Grilla.TextMatrix(IFila&, 9)), "F14.4"))
            Call ACTUREGISTRO("MASTER", "COSUNI", "CODIGO='" & Codigo & "'", CostoNuevo$, REGAF&)
            Call ACTUREGISTRO("MASTER", "FECOSTO", "CODIGO='" & Codigo & "'", FechaHoy, REGAG&)
        End If
   
   Next IFila&
   
   Unload Me
   
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  APROBO% = 0
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    GRID.Height = Me.Height - 1600
    MSF_2.Height = Me.Height - 1600
    GRID.Width = Me.Width - 235
    MSF_2.Width = Me.Width - 235
    On Error GoTo 0

End Sub

Private Sub GRID_Click()

    Static MODO  As Boolean
    COL& = GRID.MouseCol
    If (GRID.MouseRow = 0) Then
        ' Ordena en forma ascendente
        If MODO Then
            GRID.COL = GRID.MouseCol
            If GRID.COL = 3 Then GRID.COL = 12
            GRID.Sort = 2
            MODO = False
        ' Ordena en forma descendente
        Else
            GRID.COL = GRID.MouseCol
            If GRID.COL = 3 Then GRID.COL = 12
            GRID.Sort = 1
            MODO = True
        End If
    Else
        Call COLOREAR_GRILLA_SOLAFILA(GRID, vbYellow, GRID.Row, 0, 0, 0)
    End If
End Sub

Private Sub GRID_DblClick()

    If GRID.Row = 1 Then Exit Sub
    
    CAMPOS$ = "Código/A16;Descripción/A36;Uso Unitario/F12.4"
    Call CARGA_ENCABEZADO(FRMZELECHO.MSF1, CAMPOS$, FRMZELECHO)
    Call CARGA_ENCABEZADO(FRMZELECHO.msf2, CAMPOS$, FRMZELECHO)
    FRMZELECHO.Caption = TRIM$(GRID.TextMatrix(GRID.Row, 1)) & " - " & TRIM$(GRID.TextMatrix(GRID.Row, 2))
    
    CodigoConjunto% = CODINT(TRIM$(GRID.TextMatrix(GRID.Row, 1)))
    ConsultaSql$ = "SELECT MASTER.CODIGO, MASTER.DESCRIPCION, FORMULAS.USOBRUTO"
    ConsultaSql$ = ConsultaSql$ + " FROM FORMULAS WITH(NOLOCK) "
    ConsultaSql$ = ConsultaSql$ + " INNER JOIN MASTER WITH(NOLOCK) "
    ConsultaSql$ = ConsultaSql$ + " ON FORMULAS.CODICONJ = MASTER.CODINT "
    ConsultaSql$ = ConsultaSql$ + " WHERE MASTER.STATUS > 0 AND FORMULAS.CODIELEM = " & CStr(CodigoConjunto%)
    ConsultaSql$ = ConsultaSql$ + " ORDER BY FORMULAS.NuorIt ASC"
    
    Set RST = READSET(ConsultaSql$)
    With RST
        Do While Not .EOF
            FILA& = FILA& + 1
            FRMZELECHO.msf2.Rows = FRMZELECHO.msf2.Rows + 1
            FRMZELECHO.msf2.TextMatrix(FILA&, 1) = !Codigo
            FRMZELECHO.msf2.TextMatrix(FILA&, 2) = !Descripcion
            FRMZELECHO.msf2.TextMatrix(FILA&, 3) = !USOBRUTO
            
            .MoveNext
        Loop
   End With
   FRMZELECHO.msf2.Redraw = True
   
   On Error Resume Next
   RST.Close
   Set RST = Nothing
   On Error GoTo 0
   
   FRMZELECHO.Width = 8500
   FRMZELECHO.Show 1

End Sub

Private Sub Imprimir_Click()

    Dim Grilla As MSFlexGrid
    Set Grilla = GRID
    If MSF_2.Rows > 1 Then Set Grilla = MSF_2
    Call CARGA_TABLA_IMPRE(Grilla, EstructuraDeCampos, "LICOSDES", "LISTADO DE COSTOS DESACTUALIZADOS", "TABLA_IMPRE", RET_NAMECONS$, 1)
    Call FINAL("?" & RET_NAMECONS$)

End Sub

Private Sub mnuseleccion_Click()
    Dim MSF As MSFlexGrid
    If MSF_2.Visible = True Then
      Set MSF = MSF_2
    Else
      Set MSF = GRID
    End If

   FIN& = MSF.Rows - 1
   MSF.Redraw = False
   For I& = 1 To FIN&
       Call TRACE(20, I&, FIN&)
       Call COLOREAR_GRILLA_SOLAFILA(MSF, vbYellow, I&)
   Next I&
   MSF.Redraw = True

End Sub

Private Sub MSF_2_Click()
    Static MODO  As Boolean
    COL& = GRID.MouseCol
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
        Call COLOREAR_GRILLA_SOLAFILA(MSF_2, vbYellow, MSF_2.Row, 0, 0, 0)
    End If

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
    MSF_2.Rows = 2
    For I& = 1 To MSF_2.Cols - 1
       MSF_2.TextMatrix(1, I&) = ""
    Next I&
    REG& = 0
    MSF_2.Visible = True
    MSF_2.ZOrder 0
    MSF_2.Cols = GRID.Cols
    MSF_2.Redraw = False
    GRID.Visible = False
    '
    'RECORRO EL MSF QUE TIENE TODO E IMPRIMO EN msf  LOS QUE MACHEAN
    For I& = 1 To GRID.Rows - 1
      For J& = 1 To GRID.Cols - 1
        If (InStr(UCase$(GRID.TextMatrix(I&, J&)), UCase$(TXTBUSCAR))) Then
          If (TRIM$(CMBFILTRO.TEXT)) = "" Or TRIM$(UCase$(CMBFILTRO)) = TRIM$(UCase$(GRID.TextMatrix(0, J&))) Then
            REG& = REG& + 1
            MSF_2.Rows = REG& + 1
            For H& = 1 To MSF_2.Cols - 1
              MSF_2.TextMatrix(REG&, H&) = GRID.TextMatrix(I&, H&)
            Next H&
          End If
          If (TRIM$(CMBFILTRO.TEXT)) = "" Then Exit For
        End If
      Next J&
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
