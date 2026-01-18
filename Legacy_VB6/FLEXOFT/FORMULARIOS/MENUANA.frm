VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form MENUANA 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0E0E0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Análisis Estadístico de Ventas"
   ClientHeight    =   4155
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   4680
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4155
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command38 
      Caption         =   "List"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   690
      Left            =   3675
      Picture         =   "MENUANA.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   18
      ToolTipText     =   "Informe Mensual de Margen Bruto"
      Top             =   2040
      Width           =   855
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0E0E0&
      Caption         =   "Por Producto"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1800
      Left            =   210
      TabIndex        =   11
      Top             =   2205
      Width           =   3270
      Begin VB.PictureBox Picture2 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   630
         ScaleHeight     =   585
         ScaleWidth      =   2265
         TabIndex        =   12
         Top             =   945
         Width           =   2325
         Begin VB.CommandButton Command3 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   600
            Left            =   0
            Picture         =   "MENUANA.frx":030A
            Style           =   1  'Graphical
            TabIndex        =   13
            ToolTipText     =   "Análisis de Entregas en Unidades por Artículo"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Entregas en Unidades"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   540
            Left            =   840
            TabIndex        =   14
            Top             =   105
            Width           =   1380
         End
      End
      Begin VB.PictureBox Picture4 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   315
         ScaleHeight     =   585
         ScaleWidth      =   2265
         TabIndex        =   15
         Top             =   420
         Width           =   2325
         Begin VB.CommandButton Command8 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   600
            Left            =   0
            Picture         =   "MENUANA.frx":0454
            Style           =   1  'Graphical
            TabIndex        =   16
            ToolTipText     =   "Análisis de Ventas en Pesos por Artículo"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "Ventas Totales en Pesos"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   540
            Left            =   840
            TabIndex        =   17
            Top             =   50
            Width           =   1380
         End
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00C0E0E0&
      Caption         =   "Por Cliente"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1800
      Left            =   210
      TabIndex        =   4
      Top             =   210
      Width           =   3270
      Begin VB.PictureBox Picture14 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   630
         ScaleHeight     =   585
         ScaleWidth      =   2265
         TabIndex        =   5
         Top             =   945
         Width           =   2325
         Begin VB.CommandButton Command22 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   600
            Left            =   0
            Picture         =   "MENUANA.frx":075E
            Style           =   1  'Graphical
            TabIndex        =   6
            ToolTipText     =   "Análisis de Entregas por Cliente"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label15 
            BackStyle       =   0  'Transparent
            Caption         =   "Entregas en Unidades"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   540
            Left            =   840
            TabIndex        =   7
            Top             =   105
            Width           =   1590
         End
      End
      Begin VB.PictureBox Picture16 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   315
         ScaleHeight     =   585
         ScaleWidth      =   2265
         TabIndex        =   8
         Top             =   420
         Width           =   2325
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
            Height          =   600
            Left            =   0
            Picture         =   "MENUANA.frx":0A68
            Style           =   1  'Graphical
            TabIndex        =   9
            ToolTipText     =   "Análisis de Ventas por Cliente"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label17 
            BackStyle       =   0  'Transparent
            Caption         =   "Ventas Totales en Pesos"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   540
            Left            =   840
            TabIndex        =   10
            Top             =   105
            Width           =   1485
         End
      End
   End
   Begin VB.PictureBox MARCOBARRA 
      BackColor       =   &H000000FF&
      Height          =   135
      Left            =   3675
      ScaleHeight     =   75
      ScaleWidth      =   795
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   3150
      Width           =   855
      Begin VB.Frame BARRATRAZA 
         BackColor       =   &H00FF0000&
         BorderStyle     =   0  'None
         Height          =   500
         Left            =   0
         TabIndex        =   3
         Top             =   0
         Width           =   12000
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   800
      Left            =   3675
      Picture         =   "MENUANA.frx":0BB2
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Terminar - Salir de la Aplicación"
      Top             =   315
      Width           =   855
   End
   Begin VB.CommandButton Command29 
      Caption         =   "Help"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   800
      Left            =   3675
      Picture         =   "MENUANA.frx":0CFC
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Ayuda Flexoft en Línea"
      Top             =   1155
      Width           =   855
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Image Image2 
      Height          =   510
      Left            =   3100
      Picture         =   "MENUANA.frx":1006
      Top             =   3465
      Width           =   2010
   End
End
Attribute VB_Name = "MENUANA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
    Call CIERRARCH("*.*")
    Call BAJALDISCO
    Hide
End Sub


Private Sub Command22_Click()
   Command22.Enabled = False
   ANAENTCLI.Show 1
   Command22.Enabled = True
End Sub

Private Sub Command3_Click()
   Command3.Enabled = False
   ANAENTCOD.Show 1
   Command3.Enabled = True
End Sub

Private Sub Command38_Click()
    '
    Command38.Enabled = False
    '
    Call COPYSTRU("DVECLART", "MBCLIART")
    '
    If IGESVEN.MESAN = "" Then
        Call COMUNI("Imposible Continuar. Debe Elegir Mes a Consultar")
        GoTo 99
    End If
    '
    ' GENERA INFORMACION A PARTIR DEL ARCHIVO CVECLART FILTRANDO POR EL MES ACTIVO
    JJ& = 0: AA = 0: BB = 0
    For II& = 1 To ULTREG&("!DVECLART")
        XX$ = REGLEIDO$("!DVECLART", II&)
        'CMES% = CVI(Left$(XX$, 2))
        'If CMES% = CODIMES% Then
            Cli% = CVI(Mid$(XX$, 5, 2))
            CIXI% = CVI(Mid$(XX$, 7, 2))
            REMI$ = Mid$(XX$, 25, 24)
            CANTI# = CVD(Mid$(XX$, 51, 8))
            COSTO# = CVD(Mid$(XX$, 193, 8))
            precio# = CVD(Mid$(XX$, 169, 8))
            '
            ' NO CONTEMPLA A FC, ND Y NC DE TEXTO LIBRE
            If Abs(CANTI#) > 0 And CIXI% >= 0 And Not (Mid$(REMI$, 1, 2) = "FC" And CIXI% = 0) Then
                YY$ = REGBLAN("MBCLIART")
                Call REPLA(YY$, MKI$(Cli%), 5, 2)
                Call REPLA(YY$, MKI$(CIXI%), 7, 2)
                Call REPLA(YY$, MKD$(CANTI#), 51, 8)
                Call REPLA(YY$, MKD$(precio#), 145, 8)
                Call REPLA(YY$, MKD$(COSTO#), 185, 8)
                '
                JJ& = JJ& + 1
                Call GRAREG("MBCLIART", YY$, JJ&)
            End If
        'End If
    Next II&
    
    Call SETULTREG("MBCLIART", JJ&)
    
    If JJ& < 1 Then
        Call COMUNI("No Existe Información para el Período Elegido")
        GoTo 99
    End If
    
    Call FILESORT("MBCLIART", "", 4, 4)
    Call FILESORT("MBCLIART", "", 3, 3)
    
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
        GoTo 99
    End If
    '
    On Error GoTo 91
    Set XLAPP = New Excel.Application
    Set XLBOOK = XLAPP.Workbooks.Add
    Set XLSHEET = XLBOOK.Worksheets.Add
    On Error GoTo 0
    '
    'ARMA ENCABEZADO
    XLSHEET.Cells(1, 1) = "Cliente"
    XLSHEET.Cells(1, 2) = "Código"
    XLSHEET.Cells(1, 3) = "Descripción"
    XLSHEET.Cells(1, 4) = "Cantidad"
    XLSHEET.Cells(1, 5) = "Precio"
    XLSHEET.Cells(1, 6) = "Costo Unitario"
    XLSHEET.Cells(1, 7) = "% MB/V"
    XLSHEET.Cells(1, 8) = "Venta"
    XLSHEET.Cells(1, 9) = "Costo"
    XLSHEET.Cells(1, 10) = "Total % MB/V"
    
    ' ARMA EXCEL CON LOS MARGENES POR ARTICULO Y CLIENTES
    COLUMNA& = 1: FILA& = 2
    ACUMUVENTA# = 0: ACUMUCOSTO# = 0: CLIANT% = 0: ACUMUCOSCLI# = 0: ACUMUVENCLI# = 0
    For II& = 1 To ULTREG&("MBCLIART")
        XX$ = REGLEIDO$("MBCLIART", II&)
        CIXI% = CVI(Mid$(XX$, 7, 2))
            CODI$ = CODEXT$(CIXI%)
            DESCRI$ = DESCRIT0$(CIXI%)
        Cli% = CVI(Mid$(XX$, 5, 2))
            RASO$ = RASOCLI$(Cli%)
        '
        If CLIANT% = 0 Then CLIANT% = Cli%
        '
        If CLIANT% <> Cli% Then
            XLSHEET.Cells(FILA&, 1) = "Total " & RASOCLI$(CLIANT%)
            XLSHEET.Cells(FILA&, 8) = ROUND(ACUMUVENCLI#, 2)
            XLSHEET.Cells(FILA&, 9) = ROUND(ACUMUCOSCLI#, 2)
            MBV# = 0
            If Abs(ACUMUVENCLI#) > 0 Then MBV# = ROUND(100 - ((ACUMUCOSCLI# / ACUMUVENCLI#) * 100), 2)
            XLSHEET.Cells(FILA&, 10) = MBV#
            ACUMUCOSCLI# = 0: ACUMUVENCLI# = 0
            FILA& = FILA& + 2
            '
            CLIANT% = Cli%
        End If
        '
        CANTI# = ROUND#(CVD(Mid$(XX$, 51, 8)))
        TOTVENTA# = (CVD(Mid$(XX$, 145, 8))) ' CANTIDAD * PRECIO
        TOTCOSTO# = (CVD(Mid$(XX$, 185, 8))) ' CANTIDAD * COSTO
        precio# = TOTVENTA# / CANTI#
        COSTO# = TOTCOSTO# / CANTI#
        MBV# = 0
        If precio# > 0 Then MBV# = ROUND(100 - ((COSTO# / precio#) * 100), 2)
        '
        XLSHEET.Cells(FILA&, 1) = "'" & RASO$
        XLSHEET.Cells(FILA&, 2) = "'" & CODI$
        XLSHEET.Cells(FILA&, 3) = "'" & DESCRI$
        XLSHEET.Cells(FILA&, 4) = ROUND(CANTI#, 2)
        XLSHEET.Cells(FILA&, 5) = ROUND(precio#, 4)
        XLSHEET.Cells(FILA&, 6) = ROUND(COSTO#, 4)
        XLSHEET.Cells(FILA&, 7) = MBV#
        XLSHEET.Cells(FILA&, 8) = ROUND(TOTVENTA#, 2)
        XLSHEET.Cells(FILA&, 9) = ROUND(TOTCOSTO#, 2)
        '
        ACUMUCOSCLI# = ACUMUCOSCLI# + ROUND#(TOTCOSTO#, 2)
        ACUMUVENCLI# = ACUMUVENCLI# + ROUND#(TOTVENTA#, 2)
        '
        ACUMUCOSTO# = ACUMUCOSTO# + ROUND#(TOTCOSTO#, 4)
        ACUMUVENTA# = ACUMUVENTA# + ROUND#(TOTVENTA#, 4)
        '
        FILA& = FILA& + 1
22  Next II&
    
    XLSHEET.Cells(FILA&, 1) = "Total " & RASOCLI$(CLIANT%)
    XLSHEET.Cells(FILA&, 8) = ROUND(ACUMUVENCLI#)
    XLSHEET.Cells(FILA&, 9) = ROUND(ACUMUCOSCLI#)
    MBV# = 0
    If Abs(ACUMUVENCLI#) > 0 Then MBV# = ROUND(100 - ((ACUMUCOSCLI# / ACUMUVENCLI#) * 100))
    XLSHEET.Cells(FILA&, 10) = MBV#
    FILA& = FILA& + 2
    XLSHEET.Cells(FILA&, 1) = "Total GENERAL"
    XLSHEET.Cells(FILA&, 8) = ROUND#(ACUMUVENTA#)
    XLSHEET.Cells(FILA&, 9) = ROUND#(ACUMUCOSTO#)
    If Abs(ACUMUVENTA#) > 0 Then MBV# = ROUND(100 - ((ACUMUCOSTO# / ACUMUVENTA#) * 100))
    XLSHEET.Cells(FILA&, 10) = MBV#
    
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
    
    GoTo 99
    '
90: Call COMUNI("Error al Detectar la Ubicación Elegida.\Intente Nuevamente o con Otra Ubicación")
    GoTo 99
91: Call COMUNI("Error al Crear Archivo Excel")
    GoTo 99
92: Call COMUNI("Error al Grabar Archivo Excel. Verifique el Archivo No se Encuentre Abierto")
    GoTo 99
    '
99: Command38.Enabled = True
    Exit Sub
End Sub

Private Sub Command7_Click()
   Command7.Enabled = False
   ANAVENCLI.Show 1
   Command7.Enabled = True
End Sub

Private Sub Command8_Click()
   Command8.Enabled = False
   ANAVENCOD.Show 1
   Command8.Enabled = True
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   While WindowState <> 0
     WindowState = 0
     Openforms = DoEvents
   Wend
End Sub

Private Sub Form_Load()
    EPAC$ = TRIM$(EMPREAC$)
    If EPAC$ <> "" Then
        Caption = Caption + "  -  " + EPAC$
    End If
    '
End Sub


Private Sub Form_Resize()
   If WindowState = 1 Then
     Call CIERRARCH("*.*")
     Call BAJALDISCO
   End If
End Sub


Private Sub Form_Unload(Cancel As Integer)
   Call CIERRARCH("*.*")
   Call BAJALDISCO
   Hide
End Sub

