VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form GRILLAGRAL 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   7830
   ClientLeft      =   45
   ClientTop       =   600
   ClientWidth     =   16605
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   7830
   ScaleWidth      =   16605
   StartUpPosition =   2  'CenterScreen
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
      TabIndex        =   32
      Top             =   7560
      Visible         =   0   'False
      Width           =   1600
   End
   Begin VB.ListBox Lista 
      Height          =   2010
      Left            =   240
      TabIndex        =   17
      Top             =   4080
      Visible         =   0   'False
      Width           =   6255
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid2 
      Height          =   30
      Left            =   2040
      TabIndex        =   15
      Top             =   1800
      Width           =   30
      _ExtentX        =   53
      _ExtentY        =   53
      _Version        =   393216
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   30
      Left            =   6720
      TabIndex        =   14
      Top             =   2640
      Width           =   30
      _ExtentX        =   53
      _ExtentY        =   53
      _Version        =   393216
   End
   Begin VB.Frame Frame2 
      Height          =   600
      Left            =   7560
      TabIndex        =   13
      Top             =   7200
      Width           =   9015
      Begin VB.OptionButton Option3 
         Caption         =   "Torta"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   6840
         TabIndex        =   31
         Top             =   120
         Width           =   850
      End
      Begin VB.CommandButton Command3 
         Caption         =   "E"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   8580
         Picture         =   "GRILLAGRAL.frx":0000
         TabIndex        =   30
         ToolTipText     =   "Enumerar Grilla"
         Top             =   240
         Width           =   345
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Lineas"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5910
         TabIndex        =   29
         Top             =   320
         Width           =   850
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Barras"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5910
         TabIndex        =   28
         Top             =   120
         Width           =   850
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Gráfico"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   7740
         Picture         =   "GRILLAGRAL.frx":030A
         TabIndex        =   27
         Top             =   240
         Width           =   800
      End
      Begin VB.ComboBox CMBSERIE3 
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
         Left            =   4440
         Style           =   2  'Dropdown List
         TabIndex        =   22
         Top             =   240
         Width           =   1455
      End
      Begin VB.ComboBox CMBSERIE2 
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
         Left            =   3000
         Style           =   2  'Dropdown List
         TabIndex        =   21
         Top             =   240
         Width           =   1455
      End
      Begin VB.ComboBox CMBSERIE1 
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
         Left            =   1560
         Style           =   2  'Dropdown List
         TabIndex        =   20
         Top             =   240
         Width           =   1455
      End
      Begin VB.TextBox TXTSERIE2 
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
         Left            =   8040
         TabIndex        =   19
         Top             =   120
         Visible         =   0   'False
         Width           =   500
      End
      Begin VB.ComboBox CMBLABEL 
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
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   18
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label71 
         BackColor       =   &H00D8BD05&
         BackStyle       =   0  'Transparent
         Caption         =   "Serie 3:"
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
         Index           =   4
         Left            =   4440
         TabIndex        =   26
         Top             =   60
         Width           =   1290
      End
      Begin VB.Label Label71 
         BackColor       =   &H00D8BD05&
         BackStyle       =   0  'Transparent
         Caption         =   "Serie 2:"
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
         Index           =   3
         Left            =   3000
         TabIndex        =   25
         Top             =   60
         Width           =   1290
      End
      Begin VB.Label Label71 
         BackColor       =   &H00D8BD05&
         BackStyle       =   0  'Transparent
         Caption         =   "Serie 1:"
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
         Index           =   2
         Left            =   1680
         TabIndex        =   24
         Top             =   60
         Width           =   1290
      End
      Begin VB.Label Label71 
         BackColor       =   &H00D8BD05&
         BackStyle       =   0  'Transparent
         Caption         =   "Campo Eje:"
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
         Index           =   1
         Left            =   165
         TabIndex        =   23
         Top             =   60
         Width           =   1290
      End
   End
   Begin VB.Frame Frame1 
      Height          =   600
      Left            =   80
      TabIndex        =   9
      Top             =   7200
      Width           =   7455
      Begin VB.CommandButton Command4 
         Caption         =   "Inc."
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
         Left            =   3240
         TabIndex        =   16
         ToolTipText     =   "(Inc.:valor ingresado exista en la celda)(I.a.D.:valor ingresado coincida de izquierda a derecha en la celda)"
         Top             =   300
         Width           =   525
      End
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
         Left            =   600
         TabIndex        =   0
         Top             =   285
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
         Left            =   4575
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   240
         Width           =   2775
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
         Left            =   0
         TabIndex        =   11
         Top             =   285
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
         Index           =   0
         Left            =   3765
         TabIndex        =   10
         Top             =   285
         Width           =   705
      End
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
      Left            =   1320
      ScaleHeight     =   165
      ScaleWidth      =   1110
      TabIndex        =   2
      Top             =   0
      Visible         =   0   'False
      Width           =   1170
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   2625
      OleObjectBlob   =   "GRILLAGRAL.frx":0614
      Top             =   0
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   225
      Index           =   0
      Left            =   120
      OleObjectBlob   =   "GRILLAGRAL.frx":0848
      TabIndex        =   3
      Top             =   240
      Visible         =   0   'False
      Width           =   1485
   End
   Begin MSFlexGridLib.MSFlexGrid GRID 
      Height          =   7005
      Left            =   120
      TabIndex        =   4
      Top             =   0
      Width           =   16455
      _ExtentX        =   29025
      _ExtentY        =   12356
      _Version        =   393216
      Cols            =   4
      BackColor       =   16777215
      BackColorFixed  =   16501405
      BackColorBkg    =   16777215
      FillStyle       =   1
      SelectionMode   =   1
      AllowUserResizing=   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSFlexGridLib.MSFlexGrid MSF_2 
      Height          =   7125
      Left            =   60
      TabIndex        =   6
      Top             =   120
      Visible         =   0   'False
      Width           =   15620
      _ExtentX        =   27543
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
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "TEPRUEBA"
      Height          =   300
      Left            =   120
      TabIndex        =   12
      Top             =   0
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label Label7 
      BorderStyle     =   1  'Fixed Single
      Height          =   225
      Left            =   8520
      TabIndex        =   8
      Top             =   120
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Label5"
      Height          =   240
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Label TEPRUEBA 
      AutoSize        =   -1  'True
      Caption         =   "TEPRUEBA"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   135
      Left            =   0
      TabIndex        =   5
      Top             =   900
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.Menu mnuArchivo 
      Caption         =   "Archivo"
      Begin VB.Menu mnuMenu1 
         Caption         =   "MENU1"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuMenu2 
         Caption         =   "MENU2"
         Visible         =   0   'False
      End
      Begin VB.Menu s1 
         Caption         =   "-"
      End
   End
   Begin VB.Menu MnuEdición 
      Caption         =   "Edición"
      Begin VB.Menu mnuMenu3 
         Caption         =   "Revertir Selección"
      End
      Begin VB.Menu s2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuIntFilaBlanco 
         Caption         =   "Intercalar Fila en Blanco"
      End
      Begin VB.Menu mnuEliminarFilaBlanco 
         Caption         =   "Eliminar Filas en Blanco"
      End
   End
   Begin VB.Menu mnuExportarExc 
      Caption         =   "Exportar"
      Begin VB.Menu mnuExportaExcelGrillaGral 
         Caption         =   "Exportar a Excel"
      End
   End
End
Attribute VB_Name = "GRILLAGRAL"
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
Public MENU1%
Public MENU2%
Public MENU3%
Public NOPINTAR%
Public SeleccionaFilaDobleClick%
Public SinEfectoClickEnGrilla%
Public NoOrdenarColumnas% '1 no ordena

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long




Private Sub BOTEXIT_Click()
  Unload Me
End Sub

Private Sub CMBFILTRO_Click()
   TXTBUSCAR_Change
End Sub

Private Sub CMBFILTRO_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
        KeyAscii = 0        ' Para que no "pite"
        SendKeys "{tab}"    ' Envía una pulsación TAB
   End If

End Sub


Private Sub Command1_Click()
   Command1.Enabled = False
   '
   
'   CommonDialog1.CancelError = True
'   CommonDialog1.DialogTitle = "Exportar a Excel"
'       ' cancelado, si no da error - CommonDialog1.Flags = &H2& + &H4& + &H8 + &H800 + &H2000 + &H8000 'cdlOFNCreatePrompt
'   CommonDialog1.Filter = "Excel Files(*.XLS)|*.XLS|"
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
'      Call GENEXCEL(AREXPO$, MSF_2)
'     Else
'      Call GENEXCEL(AREXPO$, GRID)
'     End If
'   End If
   '

99 Command1.Enabled = True

End Sub

'Private Sub Command2_Click()
'    '
'
'    Dim MSF As MSFlexGrid
'    Set MSF = GRID
'    If TRIM$(TXTBUSCAR.TEXT) <> "" Then
'       Set MSF = MSF_2
'    End If
'    If XVALO(Me.TXTCANFILA) < 1 Then
'           Call COMUNI("Atención!! Falta Cantidad de Filas")
'       GoTo 99
'    End If
'    If XVALO(Me.TXTCAMPODESCRIPTIVO) < 1 Then
'           Call COMUNI("Atención!! Falta Campo Descriptivo")
'       GoTo 99
'    End If
'    If XVALO(Me.TXTSERIE1) < 1 Then
'       Call COMUNI("Atención!! Falta Serie (1)")
'       GoTo 99
'    End If
'
'    If MSF.Rows - 1 < XVALO(Me.TXTCANFILA) Then
'       Call COMUNI("Atención!! La Cantidad De Filas No Puede Ser Mayor a Las Expuestas en La Grilla")
'       GoTo 99
'    End If
'    If MSF.Cols < XVALO(Me.TXTSERIE1) Then
'       Call COMUNI("Atención!! La Cantidad De Series (1) No Puede Ser Mayor a Las Columnas Expuestas en La Grilla")
'       GoTo 99
'    End If
'    If MSF.Cols < XVALO(Me.TXTSERIE2) Then
'       Call COMUNI("Atención!! La Cantidad De Series (2) No Puede Ser Mayor a Las Columnas Expuestas en La Grilla")
'       GoTo 99
'    End If
'
'
'    Set FRMGRAFICO.GRID = GRID
'    If TRIM$(TXTBUSCAR.TEXT) <> "" Then
'      Set FRMGRAFICO.GRID = MSF_2
'    End If
'
'    FRMGRAFICO.CANTIFILASGRID = XVALO(Me.TXTCANFILA)
'    FRMGRAFICO.CAMPOLABEL = XVALO(Me.TXTCAMPODESCRIPTIVO)
'    FRMGRAFICO.SERIE1 = XVALO(Me.TXTSERIE1)
'    FRMGRAFICO.SERIE2 = XVALO(Me.TXTSERIE2)
'    FRMGRAFICO.Show 1
'99    Exit Sub
'    If Command2.Caption = "Maximizar" Then
'        If LecturaDatosOriginales = 0 Then
'            ' GUARDA DIMENSIONES Y POSICION DEL FORMULARIO, GRID Y BOTONES
'            FrmLeft = Me.Left: FrmTop = Me.Top: FrmWidth = Me.Width: FrmHeight = Me.Height
'            GridTop = GRID.Top: GridLeft = GRID.Left: GridWidth = GRID.Width: GridHeight = GRID.Height
'            Com1Top = Command1.Top: Com1Top = Command1.Top: Com1LEFT = Command1.Left
'            Com2Top = Command2.Top: Com2Top = Command2.Top: Com2Left = Command2.Left
'
'            FRAMETop = Frame1.Top: FRAMELeft = Frame1.Left: FRAMEWidth = Frame1.Width: FRAMEHeight = Frame1.Height
'            LecturaDatosOriginales = 1
'        End If
'
'        QCOL% = GRID.Cols
'        ANCHOCOL& = 0
'        For COLUM% = 1 To QCOL% - 1
'            ANCHOCOL& = ANCHOCOL& + GRID.ColWidth(COLUM%)
'        Next COLUM%
'        ANCHOSCR& = Screen.Width
'        If ANCHOCOL& > ANCHOSCR& Then
'            Me.Left = 0: Me.Width = ANCHOSCR&
'            GRID.Width = ANCHOSCR& - 150
'        Else
'            LEFTFORM& = Int((ANCHOSCR& - ANCHOCOL&) / 2)
'            Me.Left = LEFTFORM&: Me.Width = ANCHOCOL&
'            GRID.Width = ANCHOCOL& - 150
'        End If
'        '
'        QFILAS% = GRID.Rows
'        ALTOFILA& = 0
'        For Fila% = 1 To QFILAS% - 1
'            ALTOFILA& = ALTOFILA& + GRID.RowHeight(Fila%)
'        Next Fila%
'        ALTOSCR& = Screen.Height
'        If ALTOFILA& + 1600 > ALTOSCR& Then
'            Me.Top = 0: Me.Height = ALTOSCR& - 600
'            GRID.Height = ALTOSCR& - 1600
'        Else
'            TOPFORM& = Int((ALTOSCR& - (ALTOFILA& + 1600)) / 2)
'            Me.Top = TOPFORM&: Me.Height = ALTOFILA& + 1600
'            GRID.Height = ALTOFILA& + 600
'        End If
'        '
'        BORDEINF& = GRID.Height + 100
'        Command2.Top = BORDEINF&
'        Command1.Top = BORDEINF&
'        Frame1.Top = BORDEINF&
'        Command1.Left = Me.Width - 1800
'        Command2.Caption = "Minimizar"
'
'    Else
'        Me.Left = FrmLeft: Me.Top = FrmTop: Me.Width = FrmWidth: Me.Height = FrmHeight
'        GRID.Top = GridTop: GRID.Left = GridLeft: GRID.Width = GridWidth: GRID.Height = GridHeight
'        Frame1.Top = FRAMETop: Frame1.Left = FRAMELeft: Frame1.Width = FRAMEWidth: Frame1.Height = FRAMEHeight
'        Command2.Top = Com2Top: Command2.Top = Com2Top: Command2.Left = Com2Left
'        Command1.Top = Com1Top: Command1.Top = Com1Top: Command1.Left = Com1LEFT
'        Command2.Caption = "Maximizar"
'    End If
'    MSF_2.Height = GRID.Height
'    MSF_2.Width = GRID.Width
'    MSF_2.Top = GRID.Top
'    MSF_2.Left = GRID.Left
'
'    '
'End Sub
Sub GENEXCEL(NOMBREDEARCHIVO$, GRID As msFlexGrid)
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






Private Sub Command2_Click()
    Command2.Enabled = False
    'FALTARIA CONSIDERAR SI ESTA FILTADO
    If TRIM$(Me.CMBLABEL.TEXT) = "" Then
       Call COMUNI("Falta Seleccionar Campo de Eje")
       GoTo 99
    End If
    If TRIM$(Me.CMBSERIE1.TEXT) = "" Then
       Call COMUNI("Falta Seleccionar Campo de Valor Para Gráfico (El Primero es Mandatario)")
       GoTo 99
    End If
     'validar que los campos seleccionados tengan nombres diferentes.
    'validar si el campo es fecha que haga el lproceso y cambie el nro de columna para el eje
    Dim MSF As msFlexGrid
    Set MSF = GRID
    If TRIM(TXTBUSCAR.TEXT) <> "" Then Set MSF = MSF_2
    
    etiquetaNombre$ = UCase$(TRIM$(Mid$(Me.CMBLABEL.TEXT, 1, 48)))
    etiquetaNumeroColumna% = XVALO(TRIM$(Mid$(Me.CMBLABEL.TEXT, 50, 4)))
    etiquetaTipoDato$ = UCase$(TRIM$(Mid$(Me.CMBLABEL.TEXT, 56, 4)))
    '
    Serie1$ = Me.CMBSERIE1.TEXT
    serie1Nombre$ = UCase$(TRIM$(Mid$(Me.CMBSERIE1.TEXT, 1, 48)))
    serie1NumeroColumna% = XVALO(TRIM$(Mid$(Me.CMBSERIE1.TEXT, 50, 4)))
    serie1TipoDato$ = UCase$(TRIM$(Mid$(Me.CMBSERIE1.TEXT, 56, 4)))
    '
    Serie2$ = Me.CMBSERIE2.TEXT
    serie2Nombre$ = UCase$(TRIM$(Mid$(Me.CMBSERIE2.TEXT, 1, 48)))
    serie2NumeroColumna% = XVALO(TRIM$(Mid$(Me.CMBSERIE2.TEXT, 50, 4)))
    serie2TipoDato$ = UCase$(TRIM$(Mid$(Me.CMBSERIE2.TEXT, 56, 4)))
    
    Serie3$ = Me.CMBSERIE3.TEXT
    serie3Nombre$ = UCase$(TRIM$(Mid$(Me.CMBSERIE3.TEXT, 1, 48)))
    serie3NumeroColumna% = XVALO(TRIM$(Mid$(Me.CMBSERIE3.TEXT, 50, 4)))
    serie3TipoDato$ = UCase$(TRIM$(Mid$(Me.CMBSERIE3.TEXT, 56, 4)))
    '
    If etiquetaTipoDato$ = "D" Then 'CASO DE EJE TIPO FECHA
        Call OrdenarPorFecha(MSF, etiquetaNumeroColumna%)
        Call AgregarMesAñoAColenFinalSegunColumnaFecha(MSF, etiquetaNumeroColumna%)
        etiquetaNumeroColumna% = MSF.Cols - 1 'le cambio para que lea como eje la ultima columna que tiene el mes y año
    End If
    
    'Variables para almacenar los datos de la Grid
    Dim Codigo As String
    Dim Descripcion As String
    Dim cantidadComprada As Long
    Dim cantidadVendida As Long
    Dim cantidadEnTransito As Long
    
    'Diccionarios para almacenar los datos acumulados por descripción
    Dim comprasPorDescripcion As Object
    Dim ventasPorDescripcion As Object
    Dim transitoPorDescripcion As Object
    Set comprasPorDescripcion = CreateObject("Scripting.Dictionary")
    Set ventasPorDescripcion = CreateObject("Scripting.Dictionary")
    Set transitoPorDescripcion = CreateObject("Scripting.Dictionary")
    
    'Recorrer la Grid y acumular los datos por descripción
    For i = 1 To MSF.Rows - 1
        Codigo = MSF.TextMatrix(i, 0)
        Descripcion = TRIM$(MSF.TextMatrix(i, etiquetaNumeroColumna%))
        cantidadComprada = XVALO(MSF.TextMatrix(i, serie1NumeroColumna%))
        cantidadVendida = XVALO(MSF.TextMatrix(i, serie2NumeroColumna%))
        cantidadEnTransito = XVALO(MSF.TextMatrix(i, serie3NumeroColumna%))
        
        If Not comprasPorDescripcion.Exists(Descripcion) Then
            comprasPorDescripcion.Add Descripcion, cantidadComprada
            ventasPorDescripcion.Add Descripcion, cantidadVendida
            transitoPorDescripcion.Add Descripcion, cantidadEnTransito
        Else
            comprasPorDescripcion(Descripcion) = comprasPorDescripcion(Descripcion) + cantidadComprada
            ventasPorDescripcion(Descripcion) = ventasPorDescripcion(Descripcion) + cantidadVendida
            transitoPorDescripcion(Descripcion) = transitoPorDescripcion(Descripcion) + cantidadEnTransito
        End If
    Next i
    
    Dim descripcionActual As Variant
    EtiquetasBase$ = ""
    For Each descripcionActual In comprasPorDescripcion.Keys
      If EtiquetasBase$ = "" Then
        EtiquetasBase$ = Chr(34) & descripcionActual & Chr(34)
      Else
        EtiquetasBase$ = EtiquetasBase$ & "," & Chr(34) & descripcionActual & Chr(34)
      End If
    Next descripcionActual
    EtiquetasBase$ = "labels: [" & EtiquetasBase$ & "],"
    '
    Serie1$ = ""
    For Each descripcionActual In comprasPorDescripcion.Keys
      If Serie1$ = "" Then
        Serie1$ = Chr(34) & comprasPorDescripcion(descripcionActual) & Chr(34)
      Else
        Serie1$ = Serie1$ & "," & Chr(34) & comprasPorDescripcion(descripcionActual) & Chr(34)
      End If
    Next descripcionActual
    Serie1$ = "data: [" & Serie1$ & "],"
    '
    Serie2$ = ""
    For Each descripcionActual In ventasPorDescripcion.Keys
      If Serie2$ = "" Then
        Serie2$ = Chr(34) & ventasPorDescripcion(descripcionActual) & Chr(34)
      Else
        Serie2$ = Serie2$ & "," & Chr(34) & ventasPorDescripcion(descripcionActual) & Chr(34)
      End If
    Next descripcionActual
    Serie2$ = "data: [" & Serie2$ & "],"
    '
    Serie3$ = ""
    For Each descripcionActual In transitoPorDescripcion.Keys
      If Serie3$ = "" Then
        Serie3$ = Serie3$ & Chr(34) & transitoPorDescripcion(descripcionActual) & Chr(34)
      Else
        Serie3$ = Serie3$ & "," & Chr(34) & transitoPorDescripcion(descripcionActual) & Chr(34)
      End If
    Next descripcionActual
    Serie3$ = "data: [" & Serie3$ & "],"
    archivoHTML% = FILEOPEN%(LUDAT$ + "graficoGrilla.html", 2, 256) ' ABRE EL ARCHIVO DE ESCRITURA

tipoGrafico$ = "bar"                               'barras
If Option2.Value = True Then tipoGrafico$ = "line" 'lineas
If Option3.Value = True Then tipoGrafico$ = "pie"  'grafico de torta

Print #archivoHTML, "<!DOCTYPE html>"
Print #archivoHTML, "<html>"
Print #archivoHTML, "<head>"
Print #archivoHTML, "    <title>" & Me.Caption & "</title>"
Print #archivoHTML, "    <script src=" & Chr(34) & "https://cdn.jsdelivr.net/npm/chart.js" & Chr(34) & "></script>"
Print #archivoHTML, "</head>"
Print #archivoHTML, "<body>"
Print #archivoHTML, "    <canvas id=" & Chr(34) & "miGrafico" & Chr(34) & "></canvas>"
Print #archivoHTML, "    <script>"
Print #archivoHTML, "        var ctx = document.getElementById(" & Chr(34) & "miGrafico" & Chr(34) & ").getContext(" & Chr(34) & "2d" & Chr(34) & ");"
Print #archivoHTML, "        var miGrafico = new Chart(ctx, {"
Print #archivoHTML, "            type: " & Chr(34) & tipoGrafico$ & Chr(34) & ","
Print #archivoHTML, "            data: {"
Print #archivoHTML, "                 " & EtiquetasBase$ & ""
Print #archivoHTML, "                datasets: ["
Print #archivoHTML, "                {"
Print #archivoHTML, "                     label: " & Chr(34) & serie1Nombre$ & Chr(34) & ","
Print #archivoHTML, "                    " & Serie1$
Print #archivoHTML, "                    borderWidth: 3"
Print #archivoHTML, "                }"
If serie2NumeroColumna% > 0 Then
        Print #archivoHTML, "                ,"
        Print #archivoHTML, "                {"
        Print #archivoHTML, "                     label: " & Chr(34) & serie2Nombre$ & Chr(34) & ","
        Print #archivoHTML, "                    " & Serie2$
        Print #archivoHTML, "                    borderWidth: 3"
        Print #archivoHTML, "                }"
End If
If serie3NumeroColumna% > 0 Then
        Print #archivoHTML, "                ,"
        Print #archivoHTML, "                {"
        Print #archivoHTML, "                     label: " & Chr(34) & serie3Nombre$ & Chr(34) & ","
        Print #archivoHTML, "                    " & Serie3$
        Print #archivoHTML, "                    borderWidth: 3"
        Print #archivoHTML, "                }"
End If
Print #archivoHTML, "            ]"
Print #archivoHTML, "            },"
Print #archivoHTML, "            options: {"
Print #archivoHTML, "                scales: {"
Print #archivoHTML, "                    yAxes: [{"
Print #archivoHTML, "                        ticks: {"
Print #archivoHTML, "                            beginAtZero: true"
Print #archivoHTML, "                        }"
Print #archivoHTML, "                    }]"
Print #archivoHTML, "                }"
Print #archivoHTML, "            }"
Print #archivoHTML, "        });"
Print #archivoHTML, "    </script>"
Print #archivoHTML, "</body>"
Print #archivoHTML, "</html>"

If etiquetaTipoDato$ = "D" Then 'CASO DE EJE TIPO FECHA
   MSF.Cols = MSF.Cols - 1
End If
    Close #archivoHTML
    FileName = LUDAT$ + "graficoGrilla.html"
    A$ = " " & rutachrome$ & " " & FileName
    Shell "explorer.exe " & FileName, vbNormalFocus
99  Command2.Enabled = True
End Sub

Private Sub Command4_Click()
    Call TXTBUSCAR_Change
    If Command4.Caption = "Inc." Then
       
       Command4.Caption = "I.a.D."
    Else
        Command4.Caption = "Inc."
    End If
    
End Sub



Private Sub Command5_Click()

End Sub

Private Sub Form_Activate()
MENU1% = 0
MENU2% = 0
MENU3% = 0
SeleccionaFilaDobleClick% = 0
End Sub

Private Sub Form_Load()
    Call APLICACIONSKINS(Me)
    Option1.Value = True
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   If Cancel = 0 Then
     Unload Me
   End If
End Sub

Private Sub Form_Resize()
   On Error Resume Next
   Me.MSF_2.Width = Me.Width - 300
   Me.GRID.Width = Me.Width - 300
   Me.MSF_2.Height = Me.Height - 1600
   Me.GRID.Height = Me.Height - 1600
   Frame1.Top = Me.Height - 1300
   Frame2.Top = Me.Height - 1300
   On Error GoTo 0
End Sub
Private Sub Command3_Click()
  Dim MSF As msFlexGrid
  Set MSF = GRID
  If TRIM(TXTBUSCAR.TEXT) <> "" Then Set MSF = MSF_2
  Call numerargrilla(MSF, 0)

End Sub
Private Sub GRID_DblClick()


    i& = GRID.MouseRow
    J& = GRID.MouseCol
    If i& = 0 Or i& > GRID.Rows Then Exit Sub
    '
    On Error Resume Next
    GRID.row = i&
    GRID.COL = J&
    On Error Resume Next

    REG& = i&
    If REG& < 1 Then
        Exit Sub
    End If
    If SinEfectoClickEnGrilla% > 0 Then Exit Sub
    Call COLOREAR_GRILLA_SOLAFILA(Me.GRID, &HFFFF80, REG&)
    '
30  SeleccionaFilaDobleClick% = 1
    If GRID.Rows <= 1 Then Exit Sub
    GRID.COL = 0
    RESP_ZELE$ = GRID.TEXT
    REG& = GRID.MouseRow
    If REG& < 1 Then Exit Sub
    For i& = 1 To GRID.Cols - 1
     RESP_ZELE_VECT(i&) = GRID.TextMatrix(REG&, i&)
    Next i&
    Call T_Espera(1 / 10) 'le agrego un tiempo de espera antes que desaparezca por que si no toma el foco en otro lugar
    Me.Hide

End Sub

Private Sub GRID_Click()
    Static MODO  As Boolean
    i& = GRID.MouseRow
    J& = GRID.MouseCol

    If (GRID.MouseRow = 0) Then
        If Me.NoOrdenarColumnas% < 1 Then
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
    ElseIf SinEfectoClickEnGrilla% > 0 Then Exit Sub
    ElseIf NOPINTAR% > 0 Then Exit Sub
    End If
    If J& = 0 Then
       Call COLOREAR_GRILLA_SOLAFILA(GRID, vbYellow, GRID.row, 0, 0, 0)
    End If
End Sub

Private Sub GRID_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
 If Button = 2 Then
      strContenidoCelda$ = GRID.TextMatrix(GRID.MouseRow, GRID.MouseCol)
      Clipboard.Clear
      Clipboard.SetText strContenidoCelda
 End If
End Sub

Private Sub GRID_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If GRID.Rows <= 1 Then Exit Sub
  With GRID
    COL = .MouseCol
    ' verifica que el GRIDlexgrid contiene filas
    If .MouseRow >= 0 And COL >= 0 Then
        ' le asigna el dato con TextMatrix
        .ToolTipText = VectorToolTipText(COL)
    Else
        ' No muestra ningún ToolTip
        .ToolTipText = vbNullString
          
    End If
  End With
    
End Sub

Private Sub MC_KeyPress(KeyAscii As Integer)
   If KeyAscii = 27 Then MC.Visible = False
End Sub

Private Sub mnuEliminarFilaBlanco_Click()
    mnuEliminarFilaBlanco.Enabled = False
    Dim MSF As msFlexGrid
    Set MSF = GRILLAGRAL.GRID
    If GRILLAGRAL.MSF_2.Rows > 1 Then Set MSF = GRILLAGRAL.MSF_2
    mnuIntFilaBlanco.Enabled = False
    Screen.MousePointer = 11
    Call EliminarFilasVacias(MSF)
    mnuIntFilaBlanco.Enabled = True
    Screen.MousePointer = 1
End Sub

Private Sub mnuExportaExcelGrillaGral_Click()
    If TRIM$(TXTBUSCAR) = "" Then
       Call ImprimirEnExcel(GRID)
    Else
       Call ImprimirEnExcel(Me.MSF_2)
    End If
End Sub

Private Sub mnuIntFilaBlanco_Click()

    Dim MSF As msFlexGrid
    Set MSF = GRILLAGRAL.GRID
    If GRILLAGRAL.MSF_2.Rows > 1 Then Set MSF = GRILLAGRAL.MSF_2
    mnuIntFilaBlanco.Enabled = False
    Screen.MousePointer = 11
    Call IntercaleFilasVacias(MSF)
    mnuEliminarFilaBlanco.Enabled = True

    Screen.MousePointer = 1
    
End Sub

Private Sub mnuMenu1_Click()
  MENU1% = 1
  Hide
End Sub
Private Sub mnuMenu2_Click()
  MENU2% = 2
  Hide
End Sub
Private Sub mnuMenu3_Click()
   Screen.MousePointer = 11
   Dim MSFF As msFlexGrid
   Set MSFF = Me.GRID
   If TRIM$(TXTBUSCAR.TEXT) <> "" Then Set MSFF = Me.MSF_2
   FIN& = MSFF.Rows - 1
   MSFF.Redraw = False
   For i& = 1 To MSFF.Rows - 1
       Call TRACE(20, i&, FIN&)
       Call COLOREAR_GRILLA_SOLAFILA(MSFF, vbYellow, i&)
   Next i&
   MSFF.Redraw = True
   Screen.MousePointer = 1
End Sub


Private Sub MSF_2_Click()
    Static MODO  As Boolean
    i& = MSF_2.MouseRow
    J& = MSF_2.MouseCol

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
    ElseIf SinEfectoClickEnGrilla% > 0 Then Exit Sub
    ElseIf NOPINTAR% > 0 Then Exit Sub
    End If
    If J& = 0 Then
       Call COLOREAR_GRILLA_SOLAFILA(MSF_2, vbYellow, MSF_2.row, 0, 0, 0)
    End If
30  SeleccionaFilaDobleClick% = 1

    If MSF_2.Rows <= 1 Then Exit Sub
    MSF_2.COL = 0
    RESP_ZELE$ = MSF_2.TEXT
    REG& = MSF_2.MouseRow
    If REG& < 1 Then Exit Sub
    For i& = 1 To MSF_2.Cols - 1
     RESP_ZELE_VECT(i&) = MSF_2.TextMatrix(REG&, i&)
    Next i&
    Call T_Espera(1 / 10) 'le agrego un tiempo de espera antes que desaparezca por que si no toma el foco en otro lugar
    Me.Hide
 
End Sub


'Private Sub msf2_Click()
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



Private Sub MSF_2_DblClick()
    If NOPINTAR% > 0 Then Exit Sub
    i& = MSF_2.MouseRow
    J& = MSF_2.MouseCol
    If i& = 0 Or i& > MSF_2.Rows Then Exit Sub
    '
    On Error Resume Next
    MSF_2.row = i&
    MSF_2.COL = J&
    On Error Resume Next

    REG& = i&
    If REG& < 1 Then
        Exit Sub
    End If
    If SinEfectoClickEnGrilla% > 0 Then Exit Sub
    Call COLOREAR_GRILLA_SOLAFILA(Me.MSF_2, &HFFFF80, REG&)

End Sub

Private Sub TXTBUSCAR_Change()
    If TRIM$(TXTBUSCAR) = "" Then
      MSF_2.Visible = False
      GRID.Visible = True
      MSF_2.Rows = 1
      Call mnuEliminarFilaBlanco_Click
      mnuIntFilaBlanco.Enabled = True
      Exit Sub
    End If
    '
    mnuIntFilaBlanco.Enabled = True 'HABILITO MENU INTERCALAR FILAS EN BLANCO
    
    If Command4.Caption = "Inc." Then 'que el texto se incluya en la celda
       BusquedaIzqDer% = 0
    Else
       BusquedaIzqDer% = 1 'que el texto de la celda comience como el argumento de busqueda
    End If
    
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
    POS1% = Len(TXTBUSCAR)
    'RECORRO EL MSF QUE TIENE TODO E IMPRIMO EN msf  LOS QUE MACHEAN
    For i& = 1 To GRID.Rows - 1
      For J& = 1 To GRID.Cols - 1
         TextoMachea% = 0
         If BusquedaIzqDer% < 1 Then
            If (InStr(UCase$(GRID.TextMatrix(i&, J&)), UCase$(TXTBUSCAR))) Then TextoMachea% = 1
         Else
            If UCase$(Left$(TXTBUSCAR, POS1%)) = Left(UCase$(TRIM$(GRID.TextMatrix(i&, J&))), POS1%) Then TextoMachea% = 1
         End If
         If TextoMachea% > 0 Then
              If (TRIM$(CMBFILTRO.TEXT)) = "" Or TRIM$(UCase$(CMBFILTRO)) = TRIM$(UCase$(GRID.TextMatrix(0, J&))) Then
                REG& = REG& + 1
                MSF_2.Rows = REG& + 1
                For h& = 1 To MSF_2.Cols - 1
                  MSF_2.TextMatrix(REG&, h&) = GRID.TextMatrix(i&, h&)
                Next h&
              End If
              If (TRIM$(CMBFILTRO.TEXT)) = "" Then Exit For
            
         End If
      Next J&
    Next i&
    '
    MSF_2.Redraw = True
    
    '
End Sub

Private Sub TXTSERIE2_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
        KeyAscii = 0        ' Para que no "pite"
        SendKeys "{tab}"    ' Envía una pulsación TAB
   End If

End Sub



Private Sub ObtenerDatos_Click()
    If TRIM$(Me.CMBLABEL.TEXT) = "" Then
       Call COMUNI("Falta Seleccionar Campo Label")
       GoTo 99
    End If
    If TRIM$(Me.CMBSERIE1.TEXT) = "" Then
       Call COMUNI("Falta Seleccionar Campo de Valor Para Gráfico (El Primero es Mandatario, y Ademas Marca el Orden)")
       GoTo 99
    End If
    
    Dim FechaInicio As Date
    Dim FECHAFIN As Date
    Dim valoresPorMes() As Double
    Dim cantidadesPorMes() As Long
    Dim fechaActual As Date
    Dim i As Integer
    Dim indice As Integer
    Dim FECHA As Date
    Dim Cantidad As Long
    Dim cantidadesDeArticulosXGrupo()
    Dim cantidadesPorArticulos1()
    Dim cantidadesPorArticulos2()
    Dim cantidadesPorArticulos3()
    '
    Dim ArticuloLeido1$()
    Dim ArticuloLeido2$()
    Dim ArticuloLeido3$()
    '
    'leo que tipo de datos es la etiqueta (fecha o texto)
    ETIQUETA$ = Me.CMBLABEL.TEXT
    etiquetaNombre$ = UCase$(TRIM$(Mid$(Me.CMBLABEL.TEXT, 1, 48)))
    etiquetaNumeroColumna% = UCase$(TRIM$(Mid$(Me.CMBLABEL.TEXT, 50, 4)))
    etiquetaTipoDato$ = UCase$(TRIM$(Mid$(Me.CMBLABEL.TEXT, 56, 4)))
    '
    Serie1$ = Me.CMBSERIE1.TEXT
    serie1Nombre$ = UCase$(TRIM$(Mid$(Me.CMBSERIE1.TEXT, 1, 48)))
    serie1NumeroColumna% = XVALO(TRIM$(Mid$(Me.CMBSERIE1.TEXT, 50, 4)))
    serie1TipoDato$ = UCase$(TRIM$(Mid$(Me.CMBSERIE1.TEXT, 56, 4)))
    '
    Serie2$ = Me.CMBSERIE2.TEXT
    serie2Nombre$ = UCase$(TRIM$(Mid$(Me.CMBSERIE2.TEXT, 1, 48)))
    serie2NumeroColumna% = XVALO(TRIM$(Mid$(Me.CMBSERIE2.TEXT, 50, 4)))
    serie2TipoDato$ = UCase$(TRIM$(Mid$(Me.CMBSERIE2.TEXT, 56, 4)))
    
    Serie3$ = Me.CMBSERIE3.TEXT
    serie3Nombre$ = UCase$(TRIM$(Mid$(Me.CMBSERIE3.TEXT, 1, 48)))
    serie3NumeroColumna% = XVALO(TRIM$(Mid$(Me.CMBSERIE3.TEXT, 50, 4)))
    serie3TipoDato$ = UCase$(TRIM$(Mid$(Me.CMBSERIE3.TEXT, 56, 4)))
    '_________________________

    
    If etiquetaTipoDato$ = "D" Then
      'ESTA RUTINA ORDENA POR FECHA DE MAYOR A MENOR POR DEFECTO POR QUE SI NO NO FUNCIONA, POR QUE VA RELEVANDO SI CAMBIO DE MES Y PARA ESO NECESITA QUE ESTE ORDENADO
       Call OrdenarPorFecha(Me.GRID, etiquetaNumeroColumna%)
    Else
        GRID.COL = etiquetaNumeroColumna% 'si no ordena por texto
        GRID.Sort = 1
    End If
    
    Call numerargrilla(Me.GRID, 0)
    Me.GRID.Refresh
    Call T_Espera(1)
    ' Obtener la fecha de inicio y la fecha de fin
    If etiquetaTipoDato$ = "D" Then
                FechaInicio = GRID.TextMatrix(1, etiquetaNumeroColumna%)
                FECHAFIN = GRID.TextMatrix(GRID.Rows - 1, etiquetaNumeroColumna%)
                
                ReDim Preserve valoresPorMes(0)
                ReDim Preserve cantidadesPorMes(0)
                
                ' Inicializar el índice y la fecha anterior
                indice = 0
                
                ' Recorrer la Grid y acumular los valores por mes
                For i = 1 To GRID.Rows - 1
                    FECHA = GRID.TextMatrix(i, etiquetaNumeroColumna%)
                    FE_LEIDA% = FECHANUM(GRID.TextMatrix(i, etiquetaNumeroColumna%))
                    cantidad1 = XVALO(GRID.TextMatrix(i, serie1NumeroColumna%))
                    If serie2NumeroColumna% > 0 Then cantidad2 = XVALO(GRID.TextMatrix(i, serie2NumeroColumna%))
                    If serie3NumeroColumna% > 0 Then cantidad3 = XVALO(GRID.TextMatrix(i, serie3NumeroColumna%))
                    
                    MES_LEIDO$ = METEXTO$(FE_LEIDA%)
                    If MES_LEIDO$ <> MES_LEIDO_ANT$ And MES_LEIDO_ANT$ <> "" Then 'COMPARO SI EL MES CAMBIO Y SI YA PASO AL MENOS UNA VEZ
                        indice = indice + 1
                        ReDim Preserve valoresPorMes(indice)
                        ReDim Preserve cantidadesPorMes(indice)
                    End If
                    
                    ' Acumular los valores y las cantidades para el mes actual
                    valoresPorMes(indice) = valoresPorMes(indice) + cantidad1
                    cantidadesPorMes(indice) = cantidadesPorMes(indice) + 1
                    
                    ' Actualizar la fecha anterior
                    MES_LEIDO_ANT$ = MES_LEIDO$
                Next i
                
                ' Mostrar los resultados en la lista
                Lista.Clear
                stringMeses$ = ""
                For i = 0 To UBound(valoresPorMes)
                    fechaActual = DateAdd("m", i, FechaInicio)
                    Lista.AddItem "" & Format(fechaActual, "mmmm yyyy") & ": " & valoresPorMes(i) & " (Cantidad: " & cantidadesPorMes(i) & ")"
                    If TRIM$(stringMeses$) = "" Then
                      stringMeses$ = Chr(34) & Format(fechaActual, "mmmm yyyy") & Chr(34)
                    Else
                      stringMeses$ = stringMeses$ & "," & Chr(34) & Format(fechaActual, "mmmm yyyy") & Chr(34)
                    End If
                Next i
                stringFinal$ = "labels: [" & stringMeses$ & "],"
                '_____________hasta aqui fecha
    Else
                ReDim Preserve cantidadesDeArticulosXGrupo(0)
                ReDim Preserve cantidadesPorArticulos1(0)
                ReDim Preserve cantidadesPorArticulos2(0)
                ReDim Preserve cantidadesPorArticulos3(0)
                ReDim Preserve ArticuloLeido1$(0)
                ReDim Preserve ArticuloLeido2$(0)
                ReDim Preserve ArticuloLeido3$(0)
                
                indice = 0
                
                ' Recorrer la Grid y acumular los valores por mes
                For i = 1 To GRID.Rows - 1
                    textoEtiquetaLeido$ = GRID.TextMatrix(i, etiquetaNumeroColumna%)
                    
                    cantidad1 = XVALO(GRID.TextMatrix(i, serie1NumeroColumna%))
                    If serie2NumeroColumna% > 0 Then cantidad2 = XVALO(GRID.TextMatrix(i, serie2NumeroColumna%))
                    If serie3NumeroColumna% > 0 Then cantidad3 = XVALO(GRID.TextMatrix(i, serie3NumeroColumna%))
                    
                    
                    If textoEtiquetaLeido$ <> textoEtiquetaLeido_ANT$ And textoEtiquetaLeido_ANT$ <> "" Then 'COMPARO SI EL MES CAMBIO Y SI YA PASO AL MENOS UNA VEZ
                        indice = indice + 1
                        ReDim Preserve cantidadesDeArticulosXGrupo(indice)
                        ReDim Preserve cantidadesPorArticulos1(indice)
                        ReDim Preserve cantidadesPorArticulos2(indice)
                        ReDim Preserve cantidadesPorArticulos3(indice)
                        ReDim Preserve ArticuloLeido1$(indice)
                        ReDim Preserve ArticuloLeido2$(indice)
                        ReDim Preserve ArticuloLeido3$(indice)
                    End If
                    
                    ' Acumular los valores y las cantidades para el Articulo actual
                    cantidadesPorArticulos1(indice) = cantidadesPorArticulos1(indice) + cantidad1
                    cantidadesPorArticulos2(indice) = cantidadesPorArticulos2(indice) + cantidad2
                    cantidadesPorArticulos3(indice) = cantidadesPorArticulos3(indice) + cantidad3
                    '
                    ArticuloLeido1$(indice) = ArticuloLeido1$(indice)
                    ArticuloLeido2$(indice) = ArticuloLeido2$(indice)
                    ArticuloLeido3$(indice) = ArticuloLeido3$(indice)
                   
                    cantidadesDeArticulosXGrupo(indice) = cantidadesDeArticulosXGrupo(indice) + 1
                    
                    ' Actualizar la fecha anterior
                    textoEtiquetaLeido_ANT$ = textoEtiquetaLeido$
                Next i
                
                ' Mostrar los resultados en la lista
                Lista.Clear
                stringArticulos$ = ""
                For i = 0 To UBound(cantidadesDeArticulosXGrupo)
                    EtiquetaActual$ = cantidadesDeArticulosXGrupo(i)
                    Lista.AddItem "" & EtiquetaActual$ & ": " & valoresPorMes(i) & " (Cantidad: " & cantidadesPorMes(i) & ")"
                    If TRIM$(stringMeses$) = "" Then
                      stringArticulos$ = Chr(34) & EtiquetaActual$ & Chr(34)
                    Else
                      stringArticulos$ = stringArticulos$ & "," & Chr(34) & EtiquetaActual$ & Chr(34)
                    End If
                Next i
                stringFinal$ = "labels: [" & stringArticulos$ & "],"
               ' data: [5, 10, 7, 12, 8, 4],
    End If
    
    
    
    Lista.ZOrder 0
99
End Sub

Sub ImprimirEnExcel(MSFlexGrid1 As msFlexGrid)
    Dim excelApp As Object
    Dim excelWorkbook As Object
    Dim excelWorksheet As Object
    Dim flexGrid As msFlexGrid
    Dim i As Long
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
            ' Verificar si la columna está oculta
            If flexGrid.ColWidth(J) > 0 Then
                excelWorksheet.Cells(i + 1, J + 1).Value = flexGrid.TextMatrix(i, J)
                excelWorksheet.Cells(i + 1, J + 1).WrapText = True
            End If
        Next J
    Next i
    
    ' Ajustar automáticamente la altura de las celdas
    excelWorksheet.Cells.EntireRow.AutoFit
    
    ' Aplicar formatos a las celdas de fecha y números
    For i = 0 To flexGrid.Rows - 1
        For J = 0 To flexGrid.Cols - 1
            If flexGrid.ColWidth(J) > 0 Then ' Verificar si la columna está oculta
                If IsDate(flexGrid.TextMatrix(i, J)) And InStr(flexGrid.TextMatrix(i, J), "/") > 0 Then
                    excelWorksheet.Cells(i + 1, J + 1).NumberFormat = "dd/mm/yyyy"
                    excelWorksheet.Cells(i + 1, J + 1).Value = CDate(flexGrid.TextMatrix(i, J))
                ElseIf IsNumeric(flexGrid.TextMatrix(i, J)) Then
                    If InStr(flexGrid.TextMatrix(i, J), ".") > 0 Then
                        excelWorksheet.Cells(i + 1, J + 1).NumberFormat = "#,##0.00"
                    End If
                End If
            End If
        Next J
    Next i
    
    ' Liberar recursos de Excel
    Set excelWorksheet = Nothing
    Set excelWorkbook = Nothing
    Set excelApp = Nothing


'Dim excelApp As Object
'    Dim excelWorkbook As Object
'    Dim excelWorksheet As Object
'    Dim flexGrid As msFlexGrid
'    Dim i As Long
'    Dim J As Integer
'
'    ' Asignar el nombre de tu MSFlexGrid aquí
'    Set flexGrid = MSFlexGrid1
'
'    ' Crear una instancia de Excel
'    Set excelApp = CreateObject("Excel.Application")
'    excelApp.Visible = True
'
'    ' Agregar un nuevo libro de trabajo
'    Set excelWorkbook = excelApp.Workbooks.Add
'    Set excelWorksheet = excelWorkbook.Worksheets(1)
'
'    ' Copiar los datos desde el MSFlexGrid a Excel
'    For i = 0 To flexGrid.Rows - 1
'        For J = 0 To flexGrid.Cols - 1
'            excelWorksheet.Cells(i + 1, J + 1).Value = flexGrid.TextMatrix(i, J)
'            excelWorksheet.Cells(i + 1, J + 1).WrapText = True
'        Next J
'    Next i
'
'    ' Ajustar automáticamente la altura de las celdas
'    excelWorksheet.Cells.EntireRow.AutoFit
'
'    ' Aplicar formatos a las celdas de fecha y números
'    For i = 0 To flexGrid.Rows - 1
'        For J = 0 To flexGrid.Cols - 1
'            If IsDate(flexGrid.TextMatrix(i, J)) And InStr(flexGrid.TextMatrix(i, J), "/") > 0 Then
'                excelWorksheet.Cells(i + 1, J + 1).NumberFormat = "dd/mm/yyyy"
'                excelWorksheet.Cells(i + 1, J + 1).Value = CDate(flexGrid.TextMatrix(i, J))
'            ElseIf IsNumeric(flexGrid.TextMatrix(i, J)) Then
'                If InStr(flexGrid.TextMatrix(i, J), ".") > 0 Then
'                  excelWorksheet.Cells(i + 1, J + 1).NumberFormat = "#,##0.00"
''                Else
''                  excelWorksheet.Cells(i + 1, J + 1).NumberFormat = "#,##0"
'                End If
'            End If
'        Next J
'    Next i
'
'
'    ' Liberar recursos de Excel
'    Set excelWorksheet = Nothing
'    Set excelWorkbook = Nothing
'    Set excelApp = Nothing
End Sub
