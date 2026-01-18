VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form INFO_GRILLA09 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Valorización"
   ClientHeight    =   6525
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11565
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6525
   ScaleWidth      =   11565
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Maximizar"
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
      Left            =   30
      TabIndex        =   7
      Top             =   6120
      Width           =   1600
   End
   Begin MSFlexGridLib.MSFlexGrid MSF2 
      Height          =   5715
      Left            =   2880
      TabIndex        =   6
      ToolTipText     =   "Escape para Ocultar"
      Top             =   0
      Visible         =   0   'False
      Width           =   6275
      _ExtentX        =   11060
      _ExtentY        =   10081
      _Version        =   393216
      Cols            =   4
      BackColor       =   16777215
      BackColorFixed  =   12632319
      BackColorBkg    =   16777215
      AllowUserResizing=   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.PictureBox Picture999 
      Height          =   435
      Index           =   0
      Left            =   105
      ScaleHeight     =   375
      ScaleWidth      =   1110
      TabIndex        =   4
      Top             =   735
      Visible         =   0   'False
      Width           =   1170
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   5460
      Top             =   2415
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command8 
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
      Left            =   9870
      TabIndex        =   3
      Top             =   6120
      Width           =   1600
   End
   Begin MSFlexGridLib.MSFlexGrid MSF 
      Height          =   6045
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11505
      _ExtentX        =   20294
      _ExtentY        =   10663
      _Version        =   393216
      Cols            =   4
      BackColor       =   16777215
      BackColorFixed  =   16501405
      BackColorBkg    =   16777215
      AllowUserResizing=   1
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   420
      OleObjectBlob   =   "INFO_GRILLA09.frx":0000
      Top             =   0
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   435
      Index           =   0
      Left            =   0
      OleObjectBlob   =   "INFO_GRILLA09.frx":0234
      TabIndex        =   5
      Top             =   315
      Visible         =   0   'False
      Width           =   1590
   End
   Begin VB.Label Label1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   105
      TabIndex        =   2
      Top             =   105
      Visible         =   0   'False
      Width           =   3480
   End
   Begin VB.Label TEPRUEBA 
      AutoSize        =   -1  'True
      Caption         =   "TEPRUEBA"
      Height          =   300
      Left            =   600
      TabIndex        =   1
      Top             =   7200
      Visible         =   0   'False
      Width           =   975
   End
End
Attribute VB_Name = "INFO_GRILLA09"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim FrmLeft, FrmTop, FrmWidth, FrmHeight As Long
Dim GridTop, GridLeft, GridWidth, GridHeight As Long
Dim Com1Top, Com8Top, Com8Left As Long
Dim LecturaDatosOriginales As Byte

Private Sub Command1_Click()
    '
    If Command1.Caption = "Maximizar" Then
        If LecturaDatosOriginales = 0 Then
            ' GUARDA DIMENSIONES Y POSICION DEL FORMULARIO, GRID Y BOTONES
            FrmLeft = Me.Left: FrmTop = Me.Top: FrmWidth = Me.Width: FrmHeight = Me.Height
            GridTop = MSF.Top: GridLeft = MSF.Left: GridWidth = MSF.Width: GridHeight = MSF.Height
            Com1Top = Command1.Top: Com8Top = Command8.Top: Com8Left = Command8.Left
            LecturaDatosOriginales = 1
        End If
    
        QCOL% = MSF.Cols
        ANCHOCOL& = 0
        For COLUM% = 1 To QCOL% - 1
            ANCHOCOL& = ANCHOCOL& + MSF.ColWidth(COLUM%)
        Next COLUM%
        ANCHOSCR& = Screen.Width
        If ANCHOCOL& > ANCHOSCR& Then
            Me.Left = 0: Me.Width = ANCHOSCR&
            MSF.Width = ANCHOSCR& - 150
        Else
            LEFTFORM& = Int((ANCHOSCR& - ANCHOCOL&) / 2)
            Me.Left = LEFTFORM&: Me.Width = ANCHOCOL&
            MSF.Width = ANCHOCOL& - 150
        End If
        '
        QFILAS% = MSF.Rows
        ALTOFILA& = 0
        For FILA% = 1 To QFILAS% - 1
            ALTOFILA& = ALTOFILA& + MSF.RowHeight(FILA%)
        Next FILA%
        ALTOSCR& = Screen.Height
        If ALTOFILA& + 1600 > ALTOSCR& Then
            Me.Top = 0: Me.Height = ALTOSCR& - 600
            MSF.Height = ALTOSCR& - 1600
        Else
            TOPFORM& = Int((ALTOSCR& - (ALTOFILA& + 1600)) / 2)
            Me.Top = TOPFORM&: Me.Height = ALTOFILA& + 1600
            MSF.Height = ALTOFILA& + 600
        End If
        '
        BORDEINF& = MSF.Height + 100
        Command1.Top = BORDEINF&
        Command8.Top = BORDEINF&
        Command8.Left = Me.Width - 1800
        Command1.Caption = "Minimizar"
    Else
        Me.Left = FrmLeft: Me.Top = FrmTop: Me.Width = FrmWidth: Me.Height = FrmHeight
        MSF.Top = GridTop: MSF.Left = GridLeft: MSF.Width = GridWidth: MSF.Height = GridHeight
        Command1.Top = Com1Top: Command8.Top = Com8Top: Command8.Left = Com8Left
        Command1.Caption = "Maximizar"
    End If
    '
End Sub

Private Sub Command8_Click()
   Command8.Enabled = False
   
   'valido que haya informacion en el list
   If MSF.Rows <= 1 Then Exit Sub
   '
   'presenta cuadro de dialogo
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
   ARCHIEX$ = UCase$(CommonDialog1.FileName)
   
   'si selecciona archivo genera excel
   If TRIM$(ARCHIEX$) <> "" Then
      Call GENERAXCEL(ARCHIEX$)
   End If
   
99 Command8.Enabled = True

End Sub
Sub GENERAXCEL(NOMBREDEARCHIVO$)
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
   For i& = 1 To MSF.Rows
       RENG1& = i&
       For COLU1& = 1 To MSF.Cols
          VALORX = MSF.TextMatrix(i& - 1, COLU1 - 1)
          If IsNumeric(VALORX) = True Then
             If InStr(VALORX, ".") > 0 Then
                VALO$ = VALORX
                VALO$ = REPLACAR(VALO$, ".", "")
                VALORX = CDbl(VALO$)
             End If
          End If

          '
          XLSHEET.Cells(RENG1&, COLU1&) = VALORX
       Next COLU1&
   Next i&
   '
   On Error GoTo 30
   'XLSHEET.SaveAs NOMBREDEARCHIVO$

   XLAPP.Visible = True
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



Private Sub Form_KeyPress(KeyAscii As Integer)
   If KeyAscii = 27 Then msf2.Visible = False
End Sub

Private Sub Form_Load()
    Call APLICACIONSKINS(Me)
End Sub

Private Sub MSF_DblClick()
   
   GRAFICOB.MSChart1.Repaint = True
   GRAFICOB.MSChart1.RowCount = MSF.Cols + 1
   GRAFICOB.MSChart1.Plot.Axis(1).AxisTitle.TEXT = "Días de Cobranza"
   JJ% = 0
   i& = MSF.Row
   EMPRESA$ = TRIM$(MSF.TextMatrix(i&, 1))
   CPAGO$ = TRIM$(MSF.TextMatrix(i&, 2))
   GRAFICOB.MSChart1.Title = "Estadística de Cobranza de '" & EMPRESA$ & "'" & vbCrLf & "Condición de Pago '" & CPAGO$ & "'"

   For KKI% = 4 To MSF.Cols - 4 Step 3
        JJ% = JJ% + 1
        GRAFICOB.MSChart1.Row = JJ%
        DATO$ = MSF.TextMatrix(0, KKI%)
        DATO1$ = MSF.TextMatrix(i&, KKI%)
        GRAFICOB.MSChart1.RowLabel = DATO$
        GRAFICOB.MSChart1.Data = DATO1$
   Next KKI%

   GRAFICOB.MSChart1.RowCount = JJ%
   GRAFICOB.Show 1
   Exit Sub
   
   ' **********************************************************************************
    
   'ASIGNO POSICION SELECCIONADA A VARIABLES
   i& = MSF.Row
   J& = MSF.COL
   '
   If i& < 1 Then Exit Sub
   '
   If J& <> 1 Then Exit Sub
   
   msf2.Visible = True

   msf2.Rows = MSF.Cols
   msf2.Cols = 2
   'VALIDO
   If i& = 0 Or i& > MSF.Rows Then Exit Sub
   If J& > MSF.Cols Then Exit Sub
   '
   NUCOLU% = 1: TEPRUEBA = "123456789012345678902345678901234": msf2.ColWidth(NUCOLU% - 1) = TEPRUEBA.Width: msf2.ColAlignment(NUCOLU% - 1) = 1: msf2.TextMatrix(0, NUCOLU% - 1) = ""
   NUCOLU% = 2: TEPRUEBA = "123456789012345678902345678901234": msf2.ColWidth(NUCOLU% - 1) = TEPRUEBA.Width: msf2.ColAlignment(NUCOLU% - 1) = 1: msf2.TextMatrix(0, NUCOLU% - 1) = "Cliente"
   '
   'CARGA EL ENCABEZADO A MODO VERTICAL
   For H& = 0 To MSF.Cols - 1
      DATO1$ = MSF.TextMatrix(0, H&)
      msf2.TextMatrix(H&, 0) = TRIM$(DATO1$)
   Next H&
   '
   'CARGA LOS DATOS DEL REGISTRO SELECCIONADO
   For H& = 0 To MSF.Cols - 1
      DATO1$ = MSF.TextMatrix(i&, H&)
      msf2.TextMatrix(H&, 1) = TRIM$(DATO1$)
   Next H&
   If msf2.Rows < 24 Then msf2.Rows = 24
   Call COLOREAR_GRILLASQL(msf2, &HEFF1F1)
End Sub


