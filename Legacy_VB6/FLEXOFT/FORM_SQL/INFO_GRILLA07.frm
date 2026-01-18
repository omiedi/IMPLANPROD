VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form INFO_GRILLA07 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Valorización"
   ClientHeight    =   6570
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13515
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6570
   ScaleWidth      =   13515
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00EFF1F1&
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   2520
      ScaleHeight     =   1995
      ScaleWidth      =   5055
      TabIndex        =   9
      Top             =   5160
      Visible         =   0   'False
      Width           =   5120
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ESTA ES LA PRIMERA LINEA EN EL CASO QUE VQAYA SIGUIENDO ESCRIBIENDO SE DEBERIA MANTENER SIEMPRE DEL M"
         BeginProperty Font 
            Name            =   "Lucida Console"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Left            =   0
         TabIndex        =   10
         Top             =   0
         Visible         =   0   'False
         Width           =   5100
         WordWrap        =   -1  'True
      End
   End
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
      Left            =   40
      TabIndex        =   8
      Top             =   6120
      Width           =   1600
   End
   Begin VB.CommandButton Command7 
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
      Left            =   9675
      Picture         =   "INFO_GRILLA07.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Impresión de Grilla"
      Top             =   6120
      UseMaskColor    =   -1  'True
      Visible         =   0   'False
      Width           =   1845
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
      Left            =   11580
      TabIndex        =   3
      Top             =   6120
      Width           =   1845
   End
   Begin MSFlexGridLib.MSFlexGrid MSF 
      Height          =   6045
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   13425
      _ExtentX        =   23680
      _ExtentY        =   10663
      _Version        =   393216
      Cols            =   4
      RowHeightMin    =   300
      BackColor       =   16777215
      BackColorFixed  =   16501405
      BackColorBkg    =   16777215
      WordWrap        =   -1  'True
      AllowUserResizing=   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   420
      OleObjectBlob   =   "INFO_GRILLA07.frx":066A
      Top             =   0
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   435
      Index           =   0
      Left            =   0
      OleObjectBlob   =   "INFO_GRILLA07.frx":089E
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
Attribute VB_Name = "INFO_GRILLA07"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public CI1X%
Dim FrmLeft, FrmTop, FrmWidth, FrmHeight As Long
Dim GridTop, GridLeft, GridWidth, GridHeight As Long
Dim Com1Top, Com8Top, Com8Left As Long
Dim Com7Top, Com7Left As Long
Dim LecturaDatosOriginales As Byte

Private Sub Command1_Click()
    '
    If Command1.Caption = "Maximizar" Then
        If LecturaDatosOriginales = 0 Then
            ' GUARDA DIMENSIONES Y POSICION DEL FORMULARIO, GRID Y BOTONES
            FrmLeft = Me.Left: FrmTop = Me.Top: FrmWidth = Me.Width: FrmHeight = Me.Height
            GridTop = MSF.Top: GridLeft = MSF.Left: GridWidth = MSF.Width: GridHeight = MSF.Height
            Com1Top = Command1.Top: Com8Top = Command8.Top: Com8Left = Command8.Left
            Com7Top = Command7.Top: Com7Left = Command7.Left
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
        altoFila& = 0
        For Fila% = 1 To QFILAS% - 1
            altoFila& = altoFila& + MSF.RowHeight(Fila%)
        Next Fila%
        ALTOSCR& = Screen.Height
        If altoFila& + 1600 > ALTOSCR& Then
            Me.Top = 0: Me.Height = ALTOSCR& - 600
            MSF.Height = ALTOSCR& - 1600
        Else
            TOPFORM& = Int((ALTOSCR& - (altoFila& + 1600)) / 2)
            Me.Top = TOPFORM&: Me.Height = altoFila& + 1600
            MSF.Height = altoFila& + 600
        End If
        '
        BORDEINF& = MSF.Height + 100
        Command1.Top = BORDEINF&
        Command7.Top = BORDEINF&
        Command8.Top = BORDEINF&
        Command8.Left = Me.Width - 1800
        Command7.Left = Me.Width - 5000

        Command1.Caption = "Minimizar"
    Else
        Me.Left = FrmLeft: Me.Top = FrmTop: Me.Width = FrmWidth: Me.Height = FrmHeight
        MSF.Top = GridTop: MSF.Left = GridLeft: MSF.Width = GridWidth: MSF.Height = GridHeight
        Command1.Top = Com1Top: Command8.Top = Com8Top: Command8.Left = Com8Left
        Command7.Top = Com7Top: Command7.Left = Com7Left
        
        Command1.Caption = "Maximizar"
    End If
    '
End Sub

Private Sub Command7_Click()

   'FECHAS DESDE Y HASTA TOMADAS DEL CAPTION
   '***************************************
   POS1% = 7 + InStr(INFO_GRILLA07.Caption, "Desde:")
   FE1$ = Mid$(INFO_GRILLA07.Caption, POS1%, 8)
   DESDEX% = FECHANUM(FE1$)
   '
   POS2% = 7 + InStr(INFO_GRILLA07.Caption, "Hasta:")
   FE2$ = Mid$(INFO_GRILLA07.Caption, POS2%, 8)
   '
   HASTAX% = FECHANUM(FE2$)
   On Error Resume Next
   Call ANACOSI09.PRINTINFOGRILLA07(CI1X%, DESDEX%, HASTAX%)
   If Err Then
      On Error GoTo 0
      Call AMASTOK07.PRINTINFOGRILLA07(CI1X%, DESDEX%, HASTAX%)
   End If
   On Error GoTo 0
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
Sub GENERAXCEL(NombreDeArchivo$)
   Dim XLAPP As Excel.Application
   Dim XLBOOK As Excel.Workbook
   Dim XLSHEET As Excel.Worksheet
   '
   Dim FECHA As Date
   Dim valor As Double
   Dim Texto As String
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
                VALO$ = REPLACAR(VALO$, ".", ",")
                VALORX = CDbl(VALO$)
             End If
          End If

          '
          XLSHEET.Cells(RENG1&, COLU1&) = VALORX
       Next COLU1&
   Next i&
   '
   On Error GoTo 30
   XLSHEET.SaveAs NombreDeArchivo$

   XLAPP.Visible = True
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



Private Sub Form_KeyPress(KeyAscii As Integer)
   If KeyAscii = 27 Then MSF2.Visible = False
End Sub

Private Sub Form_Load()
    
    Call APLICACIONSKINS(Me)
    LecturaDatosOriginales = 0
        
End Sub

Private Sub Label7_Change()
   Label7.Width = 5100
End Sub

Private Sub MSF_DblClick()
   'ASIGNO POSICION SELECCIONADA A VARIABLES
   i& = MSF.Row
   J& = MSF.COL
   '
   If i& < 1 Then Exit Sub
   '
   If J& <> 1 Then Exit Sub
   
   MSF2.Visible = True

   MSF2.Rows = MSF.Cols
   MSF2.Cols = 2
   'VALIDO
   If i& = 0 Or i& > MSF.Rows Then Exit Sub
   If J& > MSF.Cols Then Exit Sub
   '
   NUCOLU% = 1: TEPRUEBA = "123456789012345678902345678901234": MSF2.ColWidth(NUCOLU% - 1) = TEPRUEBA.Width: MSF2.ColAlignment(NUCOLU% - 1) = 1: MSF2.TextMatrix(0, NUCOLU% - 1) = ""
   NUCOLU% = 2: TEPRUEBA = "123456789012345678902345678901234": MSF2.ColWidth(NUCOLU% - 1) = TEPRUEBA.Width: MSF2.ColAlignment(NUCOLU% - 1) = 1: MSF2.TextMatrix(0, NUCOLU% - 1) = "Cliente"
   '
   'CARGA EL ENCABEZADO A MODO VERTICAL
   For h& = 0 To MSF.Cols - 1
      DATO1$ = MSF.TextMatrix(0, h&)
      MSF2.TextMatrix(h&, 0) = TRIM$(DATO1$)
   Next h&
   '
   'CARGA LOS DATOS DEL REGISTRO SELECCIONADO
   For h& = 0 To MSF.Cols - 1
      DATO1$ = MSF.TextMatrix(i&, h&)
      MSF2.TextMatrix(h&, 1) = TRIM$(DATO1$)
   Next h&
   If MSF2.Rows < 24 Then MSF2.Rows = 24
   Call COLOREAR_GRILLASQL(MSF2, &HEFF1F1)
End Sub
Private Sub MSF_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
      strContenidoCelda$ = MSF.TextMatrix(MSF.MouseRow, MSF.MouseCol)
      If MsgBox(strContenidoCelda$, vbYesNo, "Copiar Contenido de Celda Seleccionada") = vbYes Then
       Clipboard.Clear
       Clipboard.SetText strContenidoCelda
      End If
    End If
End Sub

