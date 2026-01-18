VERSION 5.00
Object = "{02B5E320-7292-11CF-93D5-0020AF99504A}#1.0#0"; "MSCHART.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form ANAVENCOD 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFC0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Análisis Estadistico - Ventas en Pesos por Producto"
   ClientHeight    =   8355
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   11910
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8355
   ScaleWidth      =   11910
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   1700
      Left            =   105
      ScaleHeight     =   1695
      ScaleWidth      =   2955
      TabIndex        =   28
      Top             =   6510
      Width           =   2955
      Begin VB.ListBox List3 
         BeginProperty Font 
            Name            =   "Courier"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1225
         Left            =   0
         Sorted          =   -1  'True
         TabIndex        =   30
         Top             =   220
         Width           =   3300
      End
      Begin VB.CommandButton Command7 
         Caption         =   "Exportar a Excel"
         Height          =   230
         Left            =   0
         TabIndex        =   29
         Top             =   1475
         Width           =   2950
      End
      Begin VB.Label Label3 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Cliente"
         Height          =   285
         Left            =   0
         TabIndex        =   31
         Top             =   0
         Width           =   2955
      End
   End
   Begin VB.PictureBox Picture2 
      BorderStyle     =   0  'None
      Height          =   1485
      Left            =   3045
      ScaleHeight     =   1485
      ScaleWidth      =   8490
      TabIndex        =   24
      Top             =   6510
      Width           =   8490
      Begin VB.PictureBox Picture3 
         Height          =   1550
         Left            =   0
         ScaleHeight     =   1485
         ScaleWidth      =   0
         TabIndex        =   25
         Top             =   0
         Width           =   60
      End
      Begin VB.ListBox List4 
         BeginProperty Font 
            Name            =   "Courier"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1230
         Left            =   -30
         TabIndex        =   26
         Top             =   220
         Width           =   30000
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Cliente"
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   0
         TabIndex        =   27
         Top             =   0
         Width           =   1020
      End
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   225
      Left            =   3075
      Max             =   10000
      TabIndex        =   23
      Top             =   7980
      Width           =   8470
   End
   Begin VB.PictureBox MARCOBARRA 
      BackColor       =   &H000000FF&
      Height          =   135
      Left            =   11160
      ScaleHeight     =   75
      ScaleWidth      =   585
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   1365
      Width           =   645
      Begin VB.Frame BARRATRAZA 
         BackColor       =   &H00FF0000&
         BorderStyle     =   0  'None
         Height          =   500
         Left            =   0
         TabIndex        =   22
         Top             =   0
         Width           =   12000
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   210
      Top             =   105
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.VScrollBar VScroll2 
      Height          =   1485
      Left            =   11550
      TabIndex        =   20
      Top             =   6510
      Width           =   225
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Print"
      Height          =   645
      Left            =   11130
      Picture         =   "ANAVENCOD.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   17
      ToolTipText     =   "Imprime Ficha"
      Top             =   1575
      Width           =   645
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Exit"
      Height          =   645
      Left            =   11130
      Picture         =   "ANAVENCOD.frx":066A
      Style           =   1  'Graphical
      TabIndex        =   16
      ToolTipText     =   "Cierra y Abandona la Aplicación"
      Top             =   420
      Width           =   645
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Excluidos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1905
      Left            =   8400
      TabIndex        =   14
      Top             =   315
      Width           =   2640
      Begin VB.CommandButton Command6 
         Caption         =   "Incluir Todos"
         Height          =   225
         Left            =   210
         TabIndex        =   19
         Top             =   1510
         Width           =   2220
      End
      Begin VB.ListBox List2 
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1185
         Left            =   210
         Sorted          =   -1  'True
         TabIndex        =   15
         ToolTipText     =   "Doble-Click para Incluir "
         Top             =   350
         Width           =   2220
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Período Mensual Base"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   105
      TabIndex        =   8
      Top             =   315
      Width           =   3060
      Begin VB.CommandButton Command5 
         Caption         =   "."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   280
         Left            =   2730
         TabIndex        =   9
         Top             =   420
         Width           =   175
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Período:"
         Height          =   225
         Left            =   105
         TabIndex        =   11
         Top             =   480
         Width           =   855
      End
      Begin VB.Label MESAN 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1050
         TabIndex        =   10
         ToolTipText     =   "Listado Analisis por Facturacion"
         Top             =   420
         Width           =   1725
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Seleccionados"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1905
      Left            =   3255
      TabIndex        =   6
      Top             =   315
      Width           =   5055
      Begin VB.CommandButton Command4 
         Caption         =   "Quitar Todos"
         Height          =   225
         Left            =   210
         TabIndex        =   18
         Top             =   1510
         Width           =   4605
      End
      Begin VB.ListBox List1 
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1185
         Left            =   210
         Sorted          =   -1  'True
         TabIndex        =   7
         ToolTipText     =   "Doble-Click para Quitar"
         Top             =   350
         Width           =   4635
      End
   End
   Begin VB.Frame Frame9 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Meses"
      Height          =   960
      Left            =   105
      TabIndex        =   1
      Top             =   1260
      Width           =   3060
      Begin VB.VScrollBar VScroll1 
         Height          =   435
         Left            =   2670
         Max             =   48
         Min             =   4
         TabIndex        =   13
         Top             =   315
         Value           =   28
         Width           =   225
      End
      Begin VB.OptionButton Option3 
         BackColor       =   &H00FFFFC0&
         Caption         =   "12"
         Height          =   225
         Index           =   0
         Left            =   210
         TabIndex        =   5
         Top             =   315
         Width           =   600
      End
      Begin VB.OptionButton Option3 
         BackColor       =   &H00FFFFC0&
         Caption         =   "24"
         Height          =   225
         Index           =   1
         Left            =   525
         TabIndex        =   4
         Top             =   580
         Value           =   -1  'True
         Width           =   600
      End
      Begin VB.OptionButton Option3 
         BackColor       =   &H00FFFFC0&
         Caption         =   "36"
         Height          =   225
         Index           =   2
         Left            =   1050
         TabIndex        =   3
         Top             =   315
         Width           =   600
      End
      Begin VB.OptionButton Option3 
         BackColor       =   &H00FFFFC0&
         Caption         =   "48"
         Height          =   225
         Index           =   3
         Left            =   1365
         TabIndex        =   2
         Top             =   580
         Width           =   600
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "24"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   2145
         TabIndex        =   12
         Top             =   315
         Width           =   540
      End
   End
   Begin MSChartLib.MSChart MSChart1 
      Height          =   4005
      Left            =   105
      OleObjectBlob   =   "ANAVENCOD.frx":07B4
      TabIndex        =   0
      Top             =   2415
      Width           =   11700
   End
   Begin VB.Line Line1 
      X1              =   105
      X2              =   11550
      Y1              =   8205
      Y2              =   8205
   End
End
Attribute VB_Name = "ANAVENCOD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim DESDEFE%, HASTAFE%
Dim CORRESP$, ABREMES$, CORRESQ$
Dim DESFE$, HASFE$
Dim CODIMES%, DECODI%, HACODI%
Dim CAMESA%
'
Dim TOVEN(2048, 48)

Private Sub Command2_Click()
   Call CIERRARCH("*.*")
   Hide
End Sub

Private Sub Command3_Click()
   ANCHAN = Width
   ALTAN = Height
   ANCHCAR = MSChart1.Width
   DIFANCH = ScaleWidth - Command3.Left
   DIFALT = ScaleHeight - Picture1.Top
   MSChart1.Width = MSChart1.Width - DIFANCH
   Width = Width - DIFANCH
   Height = Height - DIFALT
   PRINTFORM
   Width = ANCHAN
   Height = ALTAN
   MSChart1.Width = ANCHCAR
End Sub

Private Sub Command4_Click()
   For U& = 1 To List1.ListCount
     List2.AddItem List1.List(U& - 1)
   Next U&
   List1.Clear
   MSChart1.Title = ""
   CAMESA% = Val(Label1)
   Call GENGRAFIC
End Sub

Private Sub Command5_Click()
   Command5.Enabled = False
   Call CARSELMESA
   Call SELECHO("SELMESA")
   MESASEL$ = VALACT1$("SELMESA")
   If TRIM$(MESASEL$) <> "" Then
      MESAN = MESASEL$
      MESAN.Refresh
      Call ARMAMES
      Call PREMESES
   End If
   Command5.Enabled = True
End Sub

Private Sub Command6_Click()
   For U& = 1 To List2.ListCount
     List1.AddItem List2.List(U& - 1)
   Next U&
   List2.Clear
   MSChart1.Title = "TODOS"
   CAMESA% = Val(Label1)
   Call GENGRAFIC
End Sub

Private Sub Command7_Click()
   '
   Command7.Enabled = False
   If List3.ListCount < 1 Then
     Call MENSERR(24, "Sin Datos para Exportar")
     GoTo 99
   End If
   '
   Dim XLAPP As Excel.Application
   Dim XLBOOK As Excel.Workbook
   Dim XLSHEET As Excel.Worksheet
   '
   Dim FECHA As Date
   Dim VALOR As Double
   Dim TEXTO As String
   '
   CommonDialog1.CancelError = True
   On Error GoTo 99
   '
5  Screen.MousePointer = 1
   CommonDialog1.DialogTitle = "Exportar a Planilla Excel"
   '\\\OT-5-779-FG-14/12/05///
   ' deshabilitado porque da error
   'CommonDialog1.Flags = &H2& + &H4& + &H8 + &H800 + &H2000 + &H8000 'cdlOFNCreatePrompt
   '\\\OT-5-779-FG-FIN///
   CommonDialog1.FILTER = "Excel-Files (*.xls)|*.xls|"
   CommonDialog1.FilterIndex = 1
   CommonDialog1.ShowSave
   '
   ' Presentar el nombre del archivo seleccionado
   Screen.MousePointer = 11
   DoEvents
   TTXX$ = UCase$(CommonDialog1.FileName)
   PPXX% = InStr(TTXX$, ".")
   If PPXX% > 0 Then
     If Mid$(TTXX$, PPXX% + 1, 3) <> "XLS" Then
       Call MENSERR(24, "Archivo no Válido - Debe ser '*.XLS'")
       GoTo 5
     End If
     TTXX$ = Left$(TTXX$, PPXX% - 1)
   End If
   NOMBREDEARCHIVO$ = TTXX$
   '
   Screen.MousePointer = 11
   RENGL& = 0: COLUL& = 1
   '
   On Error GoTo 20
   Set XLAPP = Excel.Application
   Set XLBOOK = XLAPP.Workbooks.Add
   Set XLSHEET = XLBOOK.Worksheets.Add
   On Error GoTo 0
   '
   OPENFORMS = DoEvents
   REPLACOMA% = 0
   If Val("9.99") <> 9.99 Then REPLACOMA% = 1
   '
   ' En la primera hoja los artículos seleccionados
   TTXX$ = "Código"
   RENGL& = 1: COLUL& = 1
   XLSHEET.Cells(RENGL&, COLUL&) = Chr$(39) + Left$(TTXX$, 48)
   TTXX$ = "Descripción"
   RENGL& = 1: COLUL& = 2
   XLSHEET.Cells(RENGL&, COLUL&) = Chr$(39) + Left$(TTXX$, 48)
   '
   For U& = 1 To List1.ListCount
     Call TRACE(20, U&, FIN&)
     RENGL& = RENGL& + 1
     TTXX$ = Left$(List1.List(U& - 1), 16)
     COLUL& = 1
     XLSHEET.Cells(RENGL&, COLUL&) = Chr$(39) + TTXX$
     TTXX$ = TRIM$(Mid$(List1.List(U& - 1), 17, 64))
     COLUL& = 2
     XLSHEET.Cells(RENGL&, COLUL&) = Chr$(39) + TTXX$
   Next U&
   '
   ' en la segunda hoja el detalle por cliente
   Set XLSHEET = XLBOOK.Worksheets.Add
   TTXX$ = "Cuenta"
   RENGL& = 1: COLUL& = 1
   XLSHEET.Cells(RENGL&, COLUL&) = Chr$(39) + Left$(TTXX$, 48)
   TTXX$ = "Cliente"
   RENGL& = 1: COLUL& = 2
   XLSHEET.Cells(RENGL&, COLUL&) = Chr$(39) + Left$(TTXX$, 48)
   '
   TTYY$ = Label4
   For CLXX% = 1 To Len(TTYY$) Step 10
     TENU$ = TRIM$(Mid$(TTYY$, CLXX%, 10))
     COLUL& = COLUL& + 1
     XLSHEET.Cells(RENGL&, COLUL&) = Chr$(39) + TENU$
   Next CLXX%
   '
   Screen.MousePointer = 11
   FIN& = List3.ListCount
   For U& = 1 To List3.ListCount
     Call TRACE(20, U&, FIN&)
     TTXX$ = List3.List(U& - 1)
     RENGL& = RENGL& + 1
     CLIE% = Val(Mid$(TTXX$, 121, 8))
     COLUL& = 1
     XLSHEET.Cells(RENGL&, COLUL&) = CLIE%
     COLUL& = 2
     XLSHEET.Cells(RENGL&, COLUL&) = Chr$(39) + RASOCLI(CLIE%)
     '
     TTYY$ = List4.List(U& - 1)
     For CLXX% = 1 To Len(TTYY$) Step 10
       TENU$ = TRIM$(Mid$(TTYY$, CLXX%, 10))
       PPXX% = InStr(TENU$, ".")
       If PPXX% > 0 Then
         If REPLACOMA% = 1 Then
           Mid$(TENU$, PPXX%, 1) = ","
         End If
       End If
       If PPXX% < 1 Then PPXX% = 1 + Len(TENU$)
       LUDECI% = Len(TENU$) - PPXX%: If LUDECI% < 0 Then LUDECI% = 0
       EDIVI& = 10 ^ LUDECI%
       VALOR = (Int(EDIVI& * Val(TENU$) + 0.5)) / EDIVI&
       COLUL& = COLUL& + 1
       XLSHEET.Cells(RENGL&, COLUL&) = VALOR
     Next CLXX%
   Next U&
   '
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
   GoTo 99
   '
   '
20 Resume 21
21 On Error GoTo 0
   Screen.MousePointer = 1
   MsgBox "Utilitario EXCEL no Instalado"
   Close: GoTo 99
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

99 Call CIERRARCH("*.*")
   Screen.MousePointer = 1
   Close
   Call BAJALDISCO
   Command7.Enabled = True
   '
End Sub

Private Sub Form_Load()
    '
    Call CENTRAFORM(Me)
    EPAC$ = TRIM$(EMPREAC$)
    If EPAC$ <> "" Then
        Caption = Caption + "  -  " + EPAC$
    End If
    '
    Call CARSELMESA
    MESAN = Left$(REGLEIDO$("SELMESA", 4), 12)
    Call ARMAMES
    Call PREMESES
    '
    Refresh
    DoEvents
    '
    Screen.MousePointer = 11
    List1.Clear: List2.Clear
    SQLX$ = "SELECT Codint, Descripcion, Codigo FROM MASTER WITH(NOLOCK) WHERE STATUS>0 ORDER BY CODIGO"
    Set RST = READSET(SQLX$)
    With RST
        Do While Not .EOF
          COD$ = SAFETEXT$(!Codigo)
          DES$ = SAFETEXT$(!Descripcion)
          CIXI% = XVALO(!CODINT)
          XX$ = AJUSTI$(COD$, 15) & AJUSTI$(DES$, 47) & Space$(58) & TRIM$(FORMATNUM(XVALO(CIXI%), "I8"))
          List2.AddItem XX$
          .MoveNext
        Loop
    End With
    RST.Close
    Set RST = Nothing
'    For U& = 1 To ULTREG&("ECOMAST")
'      x$ = REGLEIDO$("ECOMAST", U&)
'      List2.AddItem Left$(x$, 62) + Space$(58) + TEXTBIN$(Mid$(x$, 63, 2), "I8")
'    Next U&
    List2.ListIndex = (-1)
    Call ARMATITU
    Screen.MousePointer = 1
    '
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call CIERRARCH("*.*")
End Sub

Private Sub HScroll1_Scroll()
  List4.Left = (-1) * 10 * HScroll1.Value - 30
  Label4.Left = List4.Left
End Sub

Private Sub label1_Change()
   Call PREMESES
   CAMESA% = Val(Label1)
   Call GENGRAFIC
End Sub

Private Sub List1_DblClick()
   TTXX$ = List1.TEXT
   List1.RemoveItem List1.ListIndex
   List2.AddItem TTXX$
   Call ARMATITU
   CAMESA% = Val(Label1)
   Call GENGRAFIC
End Sub

Private Sub List2_DblClick()
   TTXX$ = List2.TEXT
   List2.RemoveItem List2.ListIndex
   List1.AddItem TTXX$
   Call ARMATITU
   CAMESA% = Val(Label1)
   Call GENGRAFIC
End Sub

Private Sub List3_Click()
   '
   ILISTA = List3.ListIndex
   List4.ListIndex = ILISTA
   MSChart1.Plot.Axis(0).AxisTitle.TEXT = TRIM$(Left$(List3.TEXT, 120))
   '
   ICOD% = Val(Mid$(List3.TEXT, 129, 8))
   If ICOD% >= 0 Then
     If ICOD% <= 2048 Then
       For UJ% = 1 To HACODI% - DECODI% + 1
         MSChart1.Row = UJ%
         MSChart1.Column = 1
         MSChart1.Data = TOVEN(ICOD%, UJ%)
       Next UJ%
     End If
   End If
   '
End Sub

Private Sub List4_Click()
   '
   ILISTA = List4.ListIndex
   List3.ListIndex = ILISTA
   MSChart1.Plot.Axis(0).AxisTitle.TEXT = TRIM$(Left$(List3.TEXT, 120))
   '
   ICOD% = Val(Mid$(List3.TEXT, 129, 8))
   If ICOD% >= 0 Then
     If ICOD% <= 2048 Then
       For UJ% = 1 To HACODI% - DECODI% + 1
         MSChart1.Row = UJ%
         MSChart1.Column = 1
         MSChart1.Data = TOVEN(ICOD%, UJ%)
       Next UJ%
     End If
   End If
   '
End Sub

Private Sub List4_Scroll()
   List3.TopIndex = List4.TopIndex
End Sub

Private Sub MESAN_Change()
   CAMESA% = Val(Label1)
   Call GENGRAFIC
End Sub

Private Sub MESAN_DblClick()
   Call CARSELMESA
   Call SELECHO("SELMESA")
   MESASEL$ = VALACT1$("SELMESA")
   If TRIM$(MESASEL$) <> "" Then
      MESAN = MESASEL$
      MESAN.Refresh
      Call ARMAMES
      Call PREMESES
   End If
End Sub

Sub GENGRAFIC()

   Call ARMAMES
   If DESDEFE% < 1 Then Exit Sub
   '
   Screen.MousePointer = 11
   Dim ACCLI%(32767), ACCOD%(32767)
   '
   MSChart1.ColumnCount = 1
   '
   List3.Clear: List4.Clear
   For UKI% = 1 To 32766
     ACCLI%(UKI%) = 0
     ACCOD%(UKI%) = 0
   Next UKI%
   '
   For UKI% = 0 To 2048
     For UKJ% = 0 To 48
       TOVEN(UKI%, UKJ%) = 0
     Next UKJ%
   Next UKI%
   '
   For U& = 1 To List1.ListCount
     ACCOD%(Val(Mid$(List1.List(U& - 1), 121, 8))) = U&
   Next U&
   '
   CAMESA% = Val(Label1)
   HACODI% = CODIMES%
   DECODI% = HACODI% - CAMESA% + 1
   MSChart1.RowCount = HACODI% - DECODI% + 1
   MSChart1.Plot.Axis(0).AxisTitle.TEXT = ""
   For UJ% = 1 To HACODI% - DECODI% + 1
      MSChart1.Row = UJ%
      MSChart1.Column = 1
      MSChart1.Data = 0
   Next UJ%
   '
   Call PREMESES
   MSChart1.Plot.Axis(1).AxisTitle.TEXT = "Pesos"
   '
   TOTAVEN = 0: TOTACOS = 0: CACLIACT% = 0
   FIN& = ULTREG&("RVECLART")
   For U& = 1 To FIN&
      Call TRACE(20, U&, FIN&)
      Z$ = REGLEIDO$("RVECLART", U&)
      CODIPE% = CVI(Left$(Z$, 2))
      If CODIPE% >= DECODI% Then
        If CODIPE% <= HACODI% Then
          CI% = CVI(Mid$(Z$, 7, 2))
          If CI% > 0 Then
            AQQ% = Abs(ACCOD%(CI%))
            If AQQ% > 0 Then
              CLI% = CVI(Mid$(Z$, 5, 2))
              If CLI% < 0 Then CLI% = 0
              ICLI% = ACCLI%(CLI%)
              If ICLI% < 1 Then
                If CACLIACT% >= 2048 Then GoTo 19
                CACLIACT% = CACLIACT% + 1
                ICLI% = CACLIACT%
                ACCLI%(CLI%) = ICLI%
                TTXYZ$ = AJUSTI$(RASOCLI$(CLI%), 128)
                Call REPLA(TTXYZ$, TRIM$(Str$(CLI%)) + Space$(8), 121, 8)
                List3.AddItem TTXYZ$ + Str$(ICLI%)
              End If
              '
              TOVEN(ICLI%, 1 + CODIPE% - DECODI%) = TOVEN(ICLI%, 1 + CODIPE% - DECODI%) + CVD(Mid$(Z$, 169, 8))
            End If
          End If
        End If
      End If
19 Next U&
   '
   For U& = 1 To List3.ListCount
     ICLI% = Val(Mid$(List3.List(U& - 1), 129))
     TTYY$ = ""
     For KI% = 1 To CAMESA%
       If KI% > CAMESA% - 24 Then
         IPAR = TOVEN(ICLI%, KI%)
         TTYY$ = TTYY$ + AJUSTD$(Format$(IPAR, "######0.00"), 10)
       End If
     Next KI%
     List4.AddItem TTYY$
   Next U&
   
   MXX% = List4.ListCount - 5
   If MXX% < 0 Then MXX% = 0
   VScroll2.Value = 0
   VScroll2.Max = MXX%
   '
90 Screen.MousePointer = 1
   '
End Sub

Sub ARMAMES()
   '
   DESDEFE% = 0: HASTAFE% = 0
   CORRESP$ = TRIM$(MESAN)
   CORRESQ$ = MESANO$(CORRESP$, DESDEFE%, HASTAFE%)
   '
   If DESDEFE% < 1 Then
     Call MENSERR(24, "Falta Indicar Período Mensual")
   End If
   '
   FEX = DESDEFE%: DESFE$ = FECHATEX$(FEX)
   FEX = HASTAFE%: HASFE$ = FECHATEX$(FEX)
   '
   CORRESQ$ = TRIM$(CORRESP$)
   ABREMES$ = Left$(CORRESQ$, 3) + Mid$(DESFE$, 7, 2)
   ANIO% = Val(Mid$(DESFE$, 7, 2)): If ANIO% < 80 Then ANIO% = 100 + ANIO%
   MESO% = Val(Mid$(DESFE$, 4, 2))
   CODIMES% = 12 * ANIO% + MESO%
   '
End Sub

Private Sub Option3_Click(Index As Integer)
   If Option3(Index).Value = True Then
     Label1 = Option3(Index).Caption
     VScroll1.Value = 52 - Label1
   End If
End Sub

Private Sub VScroll1_Change()
   Label1 = 52 - VScroll1.Value
   Label1.Refresh
End Sub

Sub PREMESES()
   '
   CAMESA% = Val(Label1)
   HACODI% = CODIMES%
   DECODI% = HACODI% - CAMESA% + 1
   '
   MSChart1.Repaint = True
   MSChart1.RowCount = CAMESA%
   Label4 = "": Label4.Left = 0
   List4.Left = -30
   For KKI% = DECODI% To HACODI%
     MSChart1.Row = 1 + KKI% - DECODI%
     MSChart1.RowLabel = ABREMESAN$(KKI%, 0)
     If KKI% > HACODI% - 24 Then
       Label4 = Label4 + AJUSTD$(ABREMESAN$(KKI%, 0), 10)
     End If
   Next KKI%
   While Label4.Width < HScroll1.Width
     Label4 = Label4 + " "
   Wend
   Label4 = Label4 + " "
   '
   HScroll1.Value = 0
   HScroll1.Max = (Label4.Width - HScroll1.Width) / 10
   '
End Sub

Sub ARMATITU()
   TTXX$ = ""
   For U& = 1 To List1.ListCount
     If TTXX$ <> "" Then TTXX$ = TTXX$ + " + "
     TTXX$ = TTXX$ + TRIM$(Left$(List1.List(U& - 1), 15))
   Next U&
   If Len(TTXX$) > 100 Then
     TTXX$ = Left$(TTXX$, 100) + " ..."
   End If
   MSChart1.Title = TTXX$
End Sub

Private Sub VScroll2_Change()
   List4.TopIndex = VScroll2.Value
End Sub

Private Sub VScroll2_Scroll()
   List4.TopIndex = VScroll2.Value
End Sub
