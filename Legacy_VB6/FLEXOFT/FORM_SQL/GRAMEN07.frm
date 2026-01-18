VERSION 5.00
Begin VB.Form GRAMEN07 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Definicion de Gráficos FLEXOFT"
   ClientHeight    =   5235
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6330
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5235
   ScaleWidth      =   6330
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
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
      Height          =   540
      Left            =   5670
      Picture         =   "GRAMEN07.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Cierra y Abandona la Aplicación"
      Top             =   210
      Width           =   540
   End
   Begin VB.Frame Frame4 
      Caption         =   "Seleccion de Registros"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4425
      Left            =   105
      TabIndex        =   3
      Top             =   630
      Width           =   5340
      Begin VB.ComboBox Combo4 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   11
         Left            =   2730
         TabIndex        =   27
         Text            =   "Combo4"
         Top             =   3990
         Width           =   2340
      End
      Begin VB.ComboBox Combo4 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   10
         Left            =   2730
         TabIndex        =   25
         Text            =   "Combo4"
         Top             =   3675
         Width           =   2340
      End
      Begin VB.ComboBox Combo4 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   9
         Left            =   2730
         TabIndex        =   23
         Text            =   "Combo4"
         Top             =   3360
         Width           =   2340
      End
      Begin VB.ComboBox Combo4 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   8
         Left            =   2730
         TabIndex        =   21
         Text            =   "Combo4"
         Top             =   3045
         Width           =   2340
      End
      Begin VB.ComboBox Combo4 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   7
         Left            =   2730
         TabIndex        =   19
         Text            =   "Combo4"
         Top             =   2730
         Width           =   2340
      End
      Begin VB.ComboBox Combo4 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   6
         Left            =   2730
         TabIndex        =   17
         Text            =   "Combo4"
         Top             =   2415
         Width           =   2340
      End
      Begin VB.ComboBox Combo4 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   5
         Left            =   2730
         TabIndex        =   15
         Text            =   "Combo4"
         Top             =   2100
         Width           =   2340
      End
      Begin VB.ComboBox Combo4 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   4
         Left            =   2730
         TabIndex        =   13
         Text            =   "Combo4"
         Top             =   1785
         Width           =   2340
      End
      Begin VB.ComboBox Combo4 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   3
         Left            =   2730
         TabIndex        =   11
         Text            =   "Combo4"
         Top             =   1470
         Width           =   2340
      End
      Begin VB.ComboBox Combo4 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   2
         Left            =   2730
         TabIndex        =   9
         Text            =   "Combo4"
         Top             =   1155
         Width           =   2340
      End
      Begin VB.ComboBox Combo4 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   1
         Left            =   2730
         TabIndex        =   7
         Text            =   "Combo4"
         Top             =   840
         Width           =   2340
      End
      Begin VB.ComboBox Combo4 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   0
         Left            =   2730
         TabIndex        =   4
         Text            =   "Combo4"
         Top             =   525
         Width           =   2340
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         Caption         =   "12:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   11
         Left            =   735
         TabIndex        =   28
         ToolTipText     =   "Doble Click para Eliminar Registro"
         Top             =   4035
         Width           =   1965
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         Caption         =   "11:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   10
         Left            =   735
         TabIndex        =   26
         ToolTipText     =   "Doble Click para Eliminar Registro"
         Top             =   3720
         Width           =   1965
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         Caption         =   "10:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   9
         Left            =   735
         TabIndex        =   24
         ToolTipText     =   "Doble Click para Eliminar Registro"
         Top             =   3405
         Width           =   1965
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         Caption         =   "9:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   8
         Left            =   735
         TabIndex        =   22
         ToolTipText     =   "Doble Click para Eliminar Registro"
         Top             =   3090
         Width           =   1965
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         Caption         =   "8:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   7
         Left            =   735
         TabIndex        =   20
         ToolTipText     =   "Doble Click para Eliminar Registro"
         Top             =   2775
         Width           =   1965
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         Caption         =   "7:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   6
         Left            =   735
         TabIndex        =   18
         ToolTipText     =   "Doble Click para Eliminar Registro"
         Top             =   2460
         Width           =   1965
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         Caption         =   "6:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   5
         Left            =   735
         TabIndex        =   16
         ToolTipText     =   "Doble Click para Eliminar Registro"
         Top             =   2145
         Width           =   1965
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         Caption         =   "5:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   4
         Left            =   735
         TabIndex        =   14
         ToolTipText     =   "Doble Click para Eliminar Registro"
         Top             =   1830
         Width           =   1965
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         Caption         =   "4:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   3
         Left            =   735
         TabIndex        =   12
         ToolTipText     =   "Doble Click para Eliminar Registro"
         Top             =   1515
         Width           =   1965
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         Caption         =   "3:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   2
         Left            =   735
         TabIndex        =   10
         ToolTipText     =   "Doble Click para Eliminar Registro"
         Top             =   1200
         Width           =   1965
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         Caption         =   "2:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   1
         Left            =   735
         TabIndex        =   8
         ToolTipText     =   "Doble Click para Eliminar Registro"
         Top             =   885
         Width           =   1965
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         Caption         =   "Registro / Valor 1:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   0
         Left            =   735
         TabIndex        =   5
         ToolTipText     =   "Doble Click para Eliminar Registro"
         Top             =   570
         Width           =   1965
      End
   End
   Begin VB.CommandButton EJECUTAR 
      Caption         =   "O.K."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   795
      Left            =   5670
      Picture         =   "GRAMEN07.frx":014A
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Aprueba y Crea la Consulta"
      Top             =   3675
      Width           =   540
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2800
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   210
      Width           =   2650
   End
   Begin VB.Image Image2 
      Height          =   390
      Left            =   5145
      Picture         =   "GRAMEN07.frx":0454
      Top             =   4650
      Width           =   1515
   End
   Begin VB.Label Label3 
      Caption         =   "Tipo / Formato del Gráfico:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   315
      TabIndex        =   1
      ToolTipText     =   "Doble Click para Nueva Consulta"
      Top             =   255
      Width           =   2430
   End
End
Attribute VB_Name = "GRAMEN07"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    '
    Command1.Enabled = False
    '
    If Combo2.TEXT = "" Then
        Call MENSERR(24, "No Hay Reporte Activo")
        GoTo 99
    End If
    '
    TICONSU% = XVALO(Label4)
    If TICONSU% = 0 Then
         FGENREP07.Label1 = Combo2.TEXT
         FGENREP07.Show 1
       ElseIf TICONSU% = 1 Then
         FGENPLA07.Show 1
       ElseIf TICONSU% = 9 Then
         FGENGRA07.Show 1
       Else
         Call MENSERR(24, "Imposible Editar Consulta de Tpo '" + TRIM$(Str$(TICONSU%)) + "'.")
    End If
    '
99  Command1.Enabled = True
    '
End Sub

Private Sub Command2_Click()
   Command2.Enabled = False
   Hide
   Command2.Enabled = True
End Sub

Private Sub Command7_Click()
   Command7.Enabled = False
   CommonDialog1.CancelError = True
   CommonDialog1.DialogTitle = "Archivo de Importación"
       ' cancelado, si no da error - CommonDialog1.Flags = &H2& + &H4& + &H8 + &H800 + &H2000 + &H8000 'cdlOFNCreatePrompt
   CommonDialog1.FILTER = "Controladores de Proceso SQF(*.SQF)|*.SQF|"
   CommonDialog1.FilterIndex = 1
   On Error GoTo 99
   CommonDialog1.ShowOpen
   On Error GoTo 0
   '
   DoEvents
   Screen.MousePointer = 11
   ARIMPO$ = UCase$(CommonDialog1.FileName)
   If TRIM$(ARIMPO$) <> "" Then
       TEXMO$ = LOADFILE$(ARIMPO$)
       TTXX$ = TEXMO$
       '
       CODMO$ = ""
       DESMO$ = ""
       While Len(TTXX$) > 0
          PPXX% = InStr(TTXX$, Chr$(124))
          If PPXX% < 1 Then PPXX% = 1 + Len(TTXX$)
          COMAI$ = Left$(TTXX$, PPXX% - 1)
          TTXX$ = Mid$(TTXX$, PPXX% + 1)
          '
          If Left$(COMAI$, 10) = "CODQUERY =" Then
              CODMO$ = UCase$(TRIM$(Mid$(COMAI$, 11)))
            ElseIf Left$(COMAI$, 10) = "DESQUERY =" Then
              DESMO$ = UCase$(TRIM$(Mid$(COMAI$, 11)))
          End If
19     Wend
       '
       If TRIM$(CODMO$) = "" Then
          Call MENSERR(24, "Imposible Registrar Consulta sin Código")
          GoTo 20
       End If
       '
       If EXISTE_TABLA%("SQL_QUERIES") > 0 Then
          Call GRABA_CONSULTA_SQL(CODMO$, DESMO$, 0, TEXMO$, OK%)
       End If
       
       Call ABRETACON
       STRSQL = "SELECT * FROM SQL_QUERIES WHERE COD_QUERY = '" & CODMO$ & "'"
       Set CONSUSQL = BaDaConsul.OpenRecordset(STRSQL, dbOpenDynaset)
       With CONSUSQL
          If (.BOF And .EOF) Then
            .AddNew   ' YA CONVERTIDO
            On Error Resume Next
            !CodiEmpr = CodiEmp%
            On Error GoTo 0
            GoTo 15
          End If
          If COMALTER%("Ya Existe una Consulta Registrada\Bajo el Código '" + CODMO$ + "'.\\Cancela Guardar\Sobre-Escribir") <> 2 Then GoTo 20
          .Edit
15        !COD_QUERY = CODMO$
          !DES_QUERY = DESMO$
          !PAR_QUERY = TEXMO$
          .Update
       End With
       CONSUSQL.Close
       Combo2.TEXT = CODMO$
       '
20 End If
   '
99 Screen.MousePointer = 1
   Command7.Enabled = True
End Sub

Private Sub EJECUTAR_Click()
    EJECUTAR.Enabled = False
    '
'    ORILE1$ = TRIM$(Combo2.TEXT)
'    ORILE2$ = TRIM$(Combo3.TEXT)
'    ORIVA1$ = TRIM$(Combo4.TEXT)
'    ORIVA2$ = TRIM$(Combo5.TEXT)
'    '
'    LONLE = XVALO(Text3)
'    If LONLE < 1 Then LONLE = XVALO(Mid$(Text3, 2))
'    If LONLE < 1 Then LONLE = 12
'    If LONLE > 32 Then LONLE = 12
'    '
'    MAXICLA% = XVALO(Text5)
'    If MAXICLA% < 1 Then MAXICLA% = 16
'    '
'    If ORILE1$ = "" Then
'       Call MENSERR(24, "Falta Elegir Leyenda Eje Abscisas.")
'       GoTo 99
'    End If
'    '
'    If ORIVA1$ = "" Then
'       Call MENSERR(24, "Falta Elegir Columna de Valores.")
'       GoTo 99
'    End If
'    '
'    COLORI% = XVALO(Left$(ORILE1$, 3))
'    COLORJ% = 0
'    If ORILE2$ <> "" Then COLORJ% = XVALO(Left$(ORILE2$, 3))
'    '
'    COLVAI% = XVALO(Left$(ORIVA1$, 3))
'    COLVAJ% = 0
'    If ORIVA2$ <> "" Then COLVAJ% = XVALO(Left$(ORIVA2$, 3))
'    '
    DIVI = 1: DIVI2 = 1
10  CANCO% = 0
    For i& = 1 To 12
      If TRIM$(Combo4(i& - 1).TEXT) <> "" Then
         CANCO% = CANCO% + 1
      End If
    Next i&
    If CANCO% < 1 Then GoTo 99
    '
    GRABAR07.MSChart1.Title.TEXT = SHOWPLA07.Label3
    GRABAR07.MSChart1.ChartType = VtChChartType2dBar
    If XVALO(Left$(Combo1.TEXT, 3)) = 1 Then
         GRABAR07.MSChart1.ChartType = VtChChartType2dLine
      ElseIf XVALO(Left$(Combo1.TEXT, 3)) = 2 Then
         GRABAR07.MSChart1.ChartType = VtChChartType3dLine
    End If
    GRABAR07.MSChart1.Repaint = True
    GRABAR07.MSChart1.ColumnCount = CANCO%
    CAMESA% = SHOWPLA07.GRID2.Cols - 2
    GRABAR07.MSChart1.RowCount = CAMESA%
    '
    With SHOWPLA07.GRID2
      For KKL% = 1 To CAMESA
         TTXX$ = .TextMatrix(0, KKL% + 1)
         GRABAR07.MSChart1.Row = KKL%
         GRABAR07.MSChart1.RowLabel = TTXX$
      Next KKL%
      '
      IREGI% = 0
      For i& = 1 To 12
        If TRIM$(Combo4(i& - 1).TEXT) <> "" Then
          IREGI% = IREGI% + 1
          KKI$ = Combo4(i& - 1).TEXT
          GRABAR07.MSChart1.Column = IREGI%
          GRABAR07.MSChart1.ColumnLabel = TRIM$(KKI$)
          GRABAR07.MSChart1.ShowLegend = True
          For KKI1% = 1 To .Rows - 1
            KKJ0$ = .TextMatrix(KKI1%, 0): KKJ1$ = .TextMatrix(KKI1%, 1)
            If KKJ0$ = KKI$ Or KKJ1$ = KKI$ Then
              For KKL% = 1 To CAMESA
                VALOR = XVALO(.TextMatrix(KKI1%, KKL% + 1))
                GRABAR07.MSChart1.Row = KKL%
                GRABAR07.MSChart1.Column = IREGI%
                GRABAR07.MSChart1.Data = VALOR / DIVI
                If VALOR / DIVI > 800000 Then DIVI2 = DIVI * 1000
              Next KKL%
            End If
          Next KKI1%
        End If
      Next i&
    End With
    If DIVI2 <> DIVI Then
       DIVI = DIVI2
       GoTo 10
    End If
    '
    GRABAR07.Refresh
    GRABAR07.Show 1
    '
99  EJECUTAR.Enabled = True
End Sub

Private Sub Form_Load()
    '
    Combo1.Clear
    Combo1.AddItem "  0.- Barras Verticales"
    Combo1.AddItem "  1.- Gráfico X-Y"
    Combo1.AddItem "  2.- Gráfico X-Y - 3 D"
    Combo1.AddItem "  3.- Torta"
    '
    For i& = 1 To 12
       Combo4(i& - 1).Clear
    Next i&
    '
    With SHOWPLA07.GRID2
       For KKI% = 1 To .Rows - 1
          KKJ$ = .TextMatrix(KKI%, 0)
          For i& = 1 To 12
             Combo4(i& - 1).AddItem KKJ$
          Next i&
       Next KKI%
    End With
    '
End Sub

Private Sub GUARDAR_Click()
    '
'    GUARDAR.Enabled = False
'    '
'    CODMO$ = TRIM$(Text1)
'    If CODMO$ = "" Then
'        Call MENSERR(24, "Falta Codigo de Consulta")
'        GoTo 99
'    End If
'    '
'    DESMO$ = TRIM$(Text2)
'    If DESMO$ = "" Then
'        Call MENSERR(24, "Falta Descripcion de Consulta")
'        GoTo 99
'    End If
'    '
'    If TRIM$(Combo1.TEXT) = "" Then
'        Call MENSERR(24, "Falta Seleccionar Tipo de Reporte")
'        GoTo 99
'    End If
'    '
'    TISMO% = XVALO(Left$(Combo1.TEXT, 3))
'    If TISMO% <> 0 Then
'      If TISMO% <> 1 Then
'        If TISMO% <> 9 Then
'          Call MENSERR(24, "Tipo de Reporte no Definido")
'          GoTo 99
'        End If
'      End If
'    End If
'    '
'    GRABOCO% = 0
'    Call ABRETACON
'    STRSQL = "SELECT * FROM SQL_QUERIES WHERE COD_QUERY = '" & CODMO$ & "'"
'    Set CONSUSQL = BaDaConsul.OpenRecordset(STRSQL, dbOpenDynaset)
'    With CONSUSQL
'        If (.BOF And .EOF) Then
'          .AddNew
'          GoTo 15
'        End If
'        If COMALTER%("Ya Existe una Consulta Registrada\Bajo El Código '" + CODMO$ + "'.\\Cancela Guardar\Sobre-Escribir") <> 2 Then GoTo 20
'        If COMALTER%("¿ Está Seguro de Sobre-Escribir el Reporte Pr-Existente ?\\No, Cancelo Guardar\Si, Sobre-Escribir") <> 2 Then GoTo 20
'        .Edit
'15      !COD_QUERY = CODMO$
'        !DES_QUERY = DESMO$
'        !TIP_QUERY = TISMO%
'        !PAR_QUERY = ""
'        .Update
'        GRABOCO% = 1
'    End With
'    '
'20  On Error Resume Next
'    CONSUSQL.Close
'    On Error GoTo 0
'    Set CONSUSQL = Nothing
'    '
'    If GRABOCO% = 1 Then
'       DEFTREP07.Combo2.TEXT = CODMO$
'       Unload Me
'    End If
'    '
'99  GUARDAR.Enabled = True
'    '
End Sub

Private Sub Label9_DblClick(Index As Integer)
    For KII% = Index + 1 To Combo4.Count - 1
       Combo4(KII% - 1).TEXT = Combo4(KII%).TEXT
       Combo4(KII%).TEXT = ""
    Next KII%
End Sub
