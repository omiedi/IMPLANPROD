VERSION 5.00
Begin VB.Form DEFGRAF07 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Definicion de Gráficos FLEXOFT"
   ClientHeight    =   5505
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
   ScaleHeight     =   5505
   ScaleWidth      =   6330
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Save"
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
      Picture         =   "DEFGRAF07.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   27
      ToolTipText     =   "Cierra y Abandona la Aplicación"
      Top             =   840
      Width           =   540
   End
   Begin VB.ComboBox Combo6 
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
      Left            =   3150
      TabIndex        =   26
      Top             =   240
      Width           =   2295
   End
   Begin VB.TextBox Text1 
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
      Left            =   480
      MaxLength       =   64
      TabIndex        =   24
      Top             =   630
      Width           =   5000
   End
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
      Picture         =   "DEFGRAF07.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   20
      ToolTipText     =   "Cierra y Abandona la Aplicación"
      Top             =   210
      Width           =   540
   End
   Begin VB.Frame Frame4 
      Caption         =   "Valores"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1380
      Left            =   105
      TabIndex        =   13
      Top             =   3990
      Width           =   5340
      Begin VB.ComboBox Combo5 
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
         Left            =   1155
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   16
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
         Left            =   1155
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Top             =   420
         Width           =   2340
      End
      Begin VB.TextBox Text4 
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
         Left            =   4515
         MaxLength       =   16
         TabIndex        =   14
         Top             =   840
         Width           =   645
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         Caption         =   "Valor 1:"
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
         TabIndex        =   19
         ToolTipText     =   "Doble Click para Nueva Consulta"
         Top             =   465
         Width           =   810
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         Caption         =   "Valor 2:"
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
         TabIndex        =   18
         ToolTipText     =   "Doble Click para Nueva Consulta"
         Top             =   885
         Width           =   810
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Caption         =   "Formato:"
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
         Left            =   3570
         TabIndex        =   17
         ToolTipText     =   "Doble Click para Nueva Consulta"
         Top             =   630
         Width           =   1065
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Eje de Clasificación"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1380
      Left            =   105
      TabIndex        =   6
      Top             =   2520
      Width           =   5340
      Begin VB.TextBox Text5 
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
         Left            =   4515
         MaxLength       =   16
         TabIndex        =   21
         Top             =   840
         Width           =   645
      End
      Begin VB.TextBox Text3 
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
         Left            =   4515
         MaxLength       =   16
         TabIndex        =   11
         Top             =   420
         Width           =   645
      End
      Begin VB.ComboBox Combo3 
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
         Left            =   1155
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   840
         Width           =   2340
      End
      Begin VB.ComboBox Combo2 
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
         Left            =   1155
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   420
         Width           =   2340
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         Caption         =   "Maximo:"
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
         Left            =   3360
         TabIndex        =   22
         ToolTipText     =   "Doble Click para Nueva Consulta"
         Top             =   885
         Width           =   1065
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "Formato:"
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
         Left            =   3360
         TabIndex        =   12
         ToolTipText     =   "Doble Click para Nueva Consulta"
         Top             =   465
         Width           =   1065
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "Clase 2:"
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
         TabIndex        =   10
         ToolTipText     =   "Doble Click para Nueva Consulta"
         Top             =   885
         Width           =   810
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Clase 1:"
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
         TabIndex        =   8
         ToolTipText     =   "Doble Click para Nueva Consulta"
         Top             =   465
         Width           =   810
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
      Picture         =   "DEFGRAF07.frx":0454
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Aprueba y Crea la Consulta"
      Top             =   3990
      Width           =   540
   End
   Begin VB.Frame Frame2 
      Caption         =   "Selección de Valores"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   120
      TabIndex        =   2
      Top             =   1560
      Width           =   5340
      Begin VB.OptionButton Option2 
         Caption         =   "Por Columnas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   2940
         TabIndex        =   4
         Top             =   420
         Value           =   -1  'True
         Width           =   1905
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Por Filas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   735
         TabIndex        =   3
         Top             =   420
         Width           =   2115
      End
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
      Top             =   1170
      Width           =   2650
   End
   Begin VB.Label Label11 
      BorderStyle     =   1  'Fixed Single
      Height          =   225
      Left            =   0
      TabIndex        =   28
      Top             =   0
      Visible         =   0   'False
      Width           =   1380
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      Caption         =   "Descripcion:"
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
      Left            =   240
      TabIndex        =   25
      ToolTipText     =   "Doble Click para Nueva Consulta"
      Top             =   400
      Width           =   1065
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Nombre:"
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
      Left            =   630
      TabIndex        =   23
      ToolTipText     =   "Doble Click para Nueva Consulta"
      Top             =   285
      Width           =   2430
   End
   Begin VB.Image Image2 
      Height          =   390
      Left            =   5145
      Picture         =   "DEFGRAF07.frx":075E
      Top             =   4965
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
      Top             =   1215
      Width           =   2430
   End
End
Attribute VB_Name = "DEFGRAF07"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Sub Combo6_Click()
    CODQUE$ = Label11
    CODGRA$ = Combo6.TEXT
    '
    Call ABRETACON
    CONDI$ = "COD_QUERY = '" & CODQUE$ & "' "
    CONDI$ = CONDI$ & " AND Cod_Graph = '" & CODGRA$ & "'"
    TEXGRA$ = VALOTACO("SQL_GRAFICS", "PAR_GRAPH", CONDI$)
    '
    Text1 = TRIM$(BUSCAVALOR(TEXGRA$, "DESGRAPH ="))
    SELVA$ = TRIM$(BUSCAVALOR(TEXGRA$, "SELVALOR ="))
    If SELVA$ = TRIM$(Option1.Caption) Then
         Option1.Value = True
       Else
         Option2.Value = True: Call Option2_Click
    End If
    '
    Call CliCombo(Combo1, TRIM$(BUSCAVALOR(TEXGRA$, "TIPGRAPH =")))
    Call CliCombo(Combo2, TRIM$(BUSCAVALOR(TEXGRA$, "ECLAS-1 =")))
    Call CliCombo(Combo3, TRIM$(BUSCAVALOR(TEXGRA$, "ECLAS-1 =")))
    Call CliCombo(Combo4, TRIM$(BUSCAVALOR(TEXGRA$, "EJVAL-1 =")))
    Call CliCombo(Combo5, TRIM$(BUSCAVALOR(TEXGRA$, "EJVAL-2 =")))
    '
    Text3 = TRIM$(BUSCAVALOR(TEXGRA$, "FORCLAS ="))
    Text5 = TRIM$(BUSCAVALOR(TEXGRA$, "ECLAMAX ="))
    Text4 = TRIM$(BUSCAVALOR(TEXGRA$, "FORXVAL ="))
    '
End Sub

Private Sub Command1_Click()
    '
    Command1.Enabled = False
    '
    CODQUE$ = TRIM$(Label11)
    If CODQUE$ = "" Then
        Call MENSERR(24, "Falta Consulta Base"): GoTo 99
    End If
    '
    CODGRA$ = TRIM$(Combo6.TEXT)
    If CODGRA$ = "" Then
        Call MENSERR(24, "Falta Codigo de Gráfico."): GoTo 99
    End If
    '
    DESGRA$ = TRIM$(Text1)
    If DESGRA$ = "" Then
        Call MENSERR(24, "Falta Título / Descripcion del Gráfico."): GoTo 99
    End If
    '
    ' ARMA EL TEXTO DE LA CONSULTA
    TEXGRA$ = ARMA_CONSUGRA
    '
10  Call ABRETACON
    STRSQL = "SELECT * FROM SQL_GRAFICS WHERE COD_QUERY = '" & CODQUE$ & "'"
    STRSQL = STRSQL + " AND Cod_Graph = '" & CODGRA$ & "'"
    '
    Set CONSUSQL = BaDaConsul.OpenRecordset(STRSQL, dbOpenDynaset)
    With CONSUSQL
      If (.BOF And .EOF) Then
        .AddNew   ' YA CONVERTIDO
        On Error Resume Next
        !CodiEmpr = CodiEmp%
        On Error GoTo 0
        GoTo 15
      End If
      If COMALTER%("Ya Existe una Consulta Registrada\Bajo Ese Código.\\Cancela Guardar\Sobre-Escribir") <> 2 Then GoTo 20
      .Edit
15    !COD_QUERY = CODQUE$
      !COD_GRAPH = CODGRA$
      !DES_GRAPH = DESGRA$
      !PAR_GRAPH = TEXGRA$
      .Update
    End With
    '
20  On Error Resume Next
    CONSUSQL.Close
    On Error GoTo 0
    Set CONSUSQL = Nothing
    '
99  Command1.Enabled = True
    '
End Sub
'
Function ARMA_CONSUGRA()
    '
    CONSUB$ = ""
    CONSUB$ = "CODQUERY = " + TRIM$(Label11) + Chr$(124)
    CONSUB$ = CONSUB$ + "CODGRAPH = " + TRIM$(Combo6.TEXT) + Chr$(124)
    CONSUB$ = CONSUB$ + "DESGRAPH = " + TRIM$(Text1) + Chr$(124)
    CONSUB$ = CONSUB$ + "TIPGRAPH = " + TRIM$(Combo1.TEXT) + Chr$(124)
    If Option1.Value = True Then
         CONSUB$ = CONSUB$ + "SELVALOR = " + Option1.Caption + Chr$(124)
       Else
         CONSUB$ = CONSUB$ + "SELVALOR = " + Option2.Caption + Chr$(124)
    End If
    CONSUB$ = CONSUB$ + "ECLAS-1 = " + TRIM$(Combo2.TEXT) + Chr$(124)
    CONSUB$ = CONSUB$ + "ECLAS-2 = " + TRIM$(Combo3.TEXT) + Chr$(124)
    CONSUB$ = CONSUB$ + "FORCLAS = " + TRIM$(Text3) + Chr$(124)
    CONSUB$ = CONSUB$ + "ECLAMAX = " + TRIM$(Text5) + Chr$(124)
    '
    CONSUB$ = CONSUB$ + "EJVAL-1 = " + TRIM$(Combo4.TEXT) + Chr$(124)
    CONSUB$ = CONSUB$ + "EJVAL-2 = " + TRIM$(Combo5.TEXT) + Chr$(124)
    CONSUB$ = CONSUB$ + "FORXVAL = " + TRIM$(Text4) + Chr$(124)
    '
    ARMA_CONSUGRA = CONSUB$
    '
End Function
'
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
       'ATENCION *** CUANDO NO SE GRABE MAS EN MDB HAY QUE AGREGAR EL COMALTER QUE SE ENCUENTRA EN LA CONSULTA MDB ARRIBA.
       If EXISTE_TABLA%("SQL_QUERIES") > 0 Then
           Call GRABA_CONSULTA_SQL(CODMO$, DESMO$, 0, TEXMO$, OK%)
       End If

       Combo2.TEXT = CODMO$
       '
20 End If
   '
99 Screen.MousePointer = 1
   Command7.Enabled = True
End Sub

Sub EJECUTAR_Click()
    EJECUTAR.Enabled = False
    '
    ORILE1$ = TRIM$(Combo2.TEXT)
    ORILE2$ = TRIM$(Combo3.TEXT)
    ORIVA1$ = TRIM$(Combo4.TEXT)
    ORIVA2$ = TRIM$(Combo5.TEXT)
    '
    LONLE = XVALO(Text3)
    If LONLE < 1 Then LONLE = XVALO(Mid$(Text3, 2))
    If LONLE < 1 Then LONLE = 12
    If LONLE > 32 Then LONLE = 12
    '
    MAXICLA% = XVALO(Text5)
    If MAXICLA% < 1 Then MAXICLA% = 16
    '
    If ORILE1$ = "" Then
       Call MENSERR(24, "Falta Elegir Leyenda Eje Abscisas.")
       GoTo 99
    End If
    '
    If ORIVA1$ = "" Then
       Call MENSERR(24, "Falta Elegir Columna de Valores.")
       GoTo 99
    End If
    '
    COLORI% = XVALO(Left$(ORILE1$, 3))
    COLORJ% = 0
    If ORILE2$ <> "" Then COLORJ% = XVALO(Left$(ORILE2$, 3))
    '
    COLVAI% = XVALO(Left$(ORIVA1$, 3))
    COLVAJ% = 0
    GRABAR07.MSChart1.ColumnCount = 1
    If ORIVA2$ <> "" Then
       COLVAJ% = XVALO(Left$(ORIVA2$, 3))
       GRABAR07.MSChart1.ColumnCount = 2
    End If
    '
    GRABAR07.MSChart1.Title.TEXT = SHOWREP07.Label3
    If TRIM$(UCase$(Text4)) = "ACUM" Then
       GRABAR07.MSChart1.Title.TEXT = SHOWREP07.Label3 + " - Acumulado"
    End If
    '
    GRABAR07.MSChart1.ChartType = VtChChartType2dBar
    If XVALO(Left$(Combo1.TEXT, 3)) = 1 Then
         GRABAR07.MSChart1.ChartType = VtChChartType2dLine
      ElseIf XVALO(Left$(Combo1.TEXT, 3)) = 2 Then
         GRABAR07.MSChart1.ChartType = VtChChartType3dLine
    End If
    GRABAR07.MSChart1.Repaint = True
    GRABAR07.MSChart1.ShowLegend = False
    With SHOWREP07.GRID
      CAMESA% = XVALO(SHOWREP07.Label7)
      If CAMESA% > MAXICLA% Then CAMESA% = MAXICLA%
      GRABAR07.MSChart1.RowCount = CAMESA%
      '
      COEFI = 1
      For KKI% = 1 To CAMESA
        VALOR = XVALO(.TextMatrix(KKI%, COLVAI% - 1)) / COEFI
        If VALOR > 990000 Then
           COEFI = COEFI * 1000
        End If
      Next KKI%
      '
      VALOANT1 = 0: VALOANT2 = 0
      For KKI% = 1 To CAMESA
        TTXX$ = .TextMatrix(KKI%, COLORI% - 1)
        If COLORJ% > 0 Then
           If TRIM$(TTXX$) <> "" Then TTXX$ = TTXX$ + " - "
           TTXX$ = TTXX$ + .TextMatrix(KKI%, COLORJ% - 1)
        End If
        TTXX$ = TRIM$(Left$(TTXX$, LONLE))
        '
        GRABAR07.MSChart1.Row = KKI%
        GRABAR07.MSChart1.RowLabel = TTXX$
        '
        VALOR = XVALO(.TextMatrix(KKI%, COLVAI% - 1)) / COEFI
        GRABAR07.MSChart1.Column = 1
        GRABAR07.MSChart1.Data = VALOANT1 + VALOR
        If UCase$(Text4) = "ACUM" Then VALOANT1 = VALOANT1 + VALOR
        '
        If COLVAJ% > 0 Then
           VALOR = XVALO(.TextMatrix(KKI%, COLVAJ% - 1)) / COEFI
           GRABAR07.MSChart1.Column = 2
           GRABAR07.MSChart1.Data = VALOANT2 + VALOR
           If UCase$(Text4) = "ACUM" Then VALOANT2 = VALOANT2 + VALOR
        End If
        '
      Next KKI%
    End With
    '
    GRABAR07.Show 1
    '
99  EJECUTAR.Enabled = True
    If SUBAPLI$ <> "" Then
       Unload Me
    End If
    '
End Sub

Private Sub Form_Load()
    '
    Combo1.Clear
    Combo1.AddItem "  0.- Barras Verticales"
    Combo1.AddItem "  1.- Gráfico X-Y"
    Combo1.AddItem "  2.- Gráfico X-Y en 3D"
    '
    Text3 = "12"
    Text5 = "16"
    '
    Call Option2_Click
    '
End Sub

Sub VerConGra()
    '
    Combo6.Clear
    Call ABRETACON
    CODQUE$ = Label11
    If CODQUE$ = "" Then CODQUE$ = SHOWREP07.Label2
    STRSQL = "SELECT * FROM SQL_GRAFICS WHERE COD_QUERY = '" & CODQUE$ & "'"
    On Error Resume Next
    Set MIRS = BaDaConsul.OpenRecordset(STRSQL, dbOpenSnapshot)
    MIRS.MoveFirst
    If Err Then
       On Error GoTo 0
       Exit Sub
    End If
    '
    On Error GoTo 0
    If Not (MIRS.EOF And MIRS.BOF) Then
       Do While Not MIRS.EOF
          Combo6.AddItem (MIRS!COD_GRAPH)
          MIRS.MoveNext
       Loop
    End If
    '
    MIRS.Close
    Set MIRS = Nothing
    BaDaConsul.Close
    '
End Sub

Private Sub Combo2_Click()
   Call Combo2_Change
End Sub

Private Sub Combo2_Change()
    '
    ' CANFI% = 0
    STRSQL$ = "SELECT * FROM SQL_QUERIES with(NOLOCK) where COD_QUERY = '" & Combo2.TEXT & "'"
    Set RS = READSET(STRSQL$)
    If Not (RS.EOF And RS.BOF) Then
       Label9 = RS!DES_QUERY
       TICONSU% = XVALO(RS!TIP_QUERY)
       Label4 = TRIM$(Str$(TICONSU%))
       '
       If TICONSU% = 0 Then Label2 = "Tabla Plana"
       If TICONSU% = 1 Then Label2 = "Estadistica Mensual"
       If TICONSU% = 9 Then Label2 = "Grafico"
       Exit Sub ' SI ENTRO ES QUE ENCONTRO  Y SE VA
    End If

    STRSQL = "SELECT * FROM SQL_QUERIES where COD_QUERY = '" & Combo2.TEXT & "'"
    Set MIRS = BaDaConsul.OpenRecordset(STRSQL, dbOpenSnapshot)
    If Not (MIRS.EOF And MIRS.BOF) Then
       MIRS.MoveFirst
       Label9 = MIRS!DES_QUERY
       TICONSU% = XVALO(MIRS!TIP_QUERY)
       Label4 = TRIM$(Str$(TICONSU%))
       '
       If TICONSU% = 0 Then Label2 = "Tabla Plana"
       If TICONSU% = 1 Then Label2 = "Estadistica Mensual"
       If TICONSU% = 9 Then Label2 = "Grafico"
       
    End If
    '
End Sub

Private Sub Label11_Change()
    Call VerConGra
End Sub

Private Sub Option2_Click()
    '
    Combo2.Clear
    Combo3.Clear
    Combo4.Clear
    Combo5.Clear
    With SHOWREP07.GRID
       For KKI% = 0 To .Cols - 1
          KKJ$ = FORMATNUM$(KKI% + 1, "I3") + ".- "
          Combo2.AddItem KKJ$ + .TextMatrix(0, KKI%)
          Combo3.AddItem KKJ$ + .TextMatrix(0, KKI%)
          Combo4.AddItem KKJ$ + .TextMatrix(0, KKI%)
          Combo5.AddItem KKJ$ + .TextMatrix(0, KKI%)
       Next KKI%
    End With
    '
End Sub


