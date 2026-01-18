VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form DEFTREP07 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Generador de Reportes  FLEXOFT"
   ClientHeight    =   2805
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6660
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
   ScaleHeight     =   2805
   ScaleWidth      =   6660
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "Crear Nueva Consulta"
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
      Left            =   105
      TabIndex        =   3
      ToolTipText     =   "Crea Nuevo Reporte"
      Top             =   2310
      Width           =   3165
   End
   Begin VB.Frame Frame1 
      Caption         =   "Reporte Activo"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2010
      Left            =   105
      TabIndex        =   5
      Top             =   210
      Width           =   6420
      Begin VB.Frame Frame2 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   375
         Left            =   -450
         TabIndex        =   12
         Top             =   420
         Width           =   3855
         Begin VB.Line Line1 
            BorderColor     =   &H00C0C0C0&
            X1              =   450
            X2              =   450
            Y1              =   0
            Y2              =   500
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
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
            Left            =   600
            TabIndex        =   14
            ToolTipText     =   "Doble Click para Nueva Consulta"
            Top             =   150
            Width           =   1065
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            Caption         =   "Consulta:"
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
            Left            =   2640
            TabIndex        =   13
            ToolTipText     =   "Doble Click para Nueva Consulta"
            Top             =   50
            Width           =   1065
         End
         Begin VB.Line Line2 
            BorderColor     =   &H00808080&
            X1              =   3820
            X2              =   3820
            Y1              =   0
            Y2              =   300
         End
      End
      Begin VB.PictureBox Picture1 
         BorderStyle     =   0  'None
         Height          =   1770
         Left            =   5500
         ScaleHeight     =   1770
         ScaleWidth      =   885
         TabIndex        =   10
         Top             =   200
         Width           =   880
         Begin VB.CommandButton Command1 
            Caption         =   "Edit"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   105
            Style           =   1  'Graphical
            TabIndex        =   2
            ToolTipText     =   "Abre Reporte para Edición y Modificación"
            Top             =   210
            Width           =   645
         End
         Begin VB.CommandButton EJECUTAR 
            Caption         =   "Exec."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   585
            Left            =   105
            Picture         =   "DEFTREP07.frx":0000
            Style           =   1  'Graphical
            TabIndex        =   1
            ToolTipText     =   "Ejecuta la Consulta"
            Top             =   645
            Width           =   645
         End
         Begin VB.Image Image2 
            Height          =   390
            Left            =   -300
            Picture         =   "DEFTREP07.frx":030A
            Top             =   1380
            Width           =   1515
         End
      End
      Begin VB.ComboBox Combo2 
         BeginProperty Font 
            Name            =   "Lucida Console"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   -1620
         Sorted          =   -1  'True
         TabIndex        =   0
         Top             =   420
         Width           =   7085
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2580
         TabIndex        =   11
         Top             =   1575
         Width           =   260
      End
      Begin VB.Label Label3 
         Caption         =   "Tipo / Formato  de Reporte:"
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
         Left            =   105
         TabIndex        =   9
         ToolTipText     =   "Doble Click para Nueva Consulta"
         Top             =   1620
         Width           =   2430
      End
      Begin VB.Label Label2 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2940
         TabIndex        =   8
         Top             =   1575
         Width           =   2535
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         Caption         =   "Consulta"
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
         Left            =   1995
         TabIndex        =   7
         ToolTipText     =   "Doble Click para Nueva Consulta"
         Top             =   465
         Width           =   1065
      End
      Begin VB.Label Label9 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   525
         TabIndex        =   6
         Top             =   840
         Width           =   4950
      End
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Importador de Reportes"
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
      Left            =   3360
      TabIndex        =   4
      ToolTipText     =   "Importador de Consultas"
      Top             =   2310
      Width           =   3165
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "DEFTREP07"
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
         FGENREP07.Label1 = Mid$(Combo2.TEXT, 49)
         FGENREP07.Show 1
       ElseIf TICONSU% = 1 Then
         FGENPLA07.Label1 = Mid$(Combo2.TEXT, 49)
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
   NEWREP07.Show 1
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
       TIPMO% = 0
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
            ElseIf Left$(COMAI$, 10) = "TIPQUERY =" Then
              TIPMO% = XVALO(Mid$(COMAI$, 11))
          End If
19     Wend
       '
       If TRIM$(CODMO$) = "" Then
          Call MENSERR(24, "Imposible Registrar Consulta sin Código")
          GoTo 20
       End If
       '
       If EXISTE_TABLA%("SQL_QUERIES") > 0 Then
          Call GRABA_CONSULTA_SQL(CODMO$, DESMO$, TIPMO%, TEXMO$, OK%)
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
          !TIP_QUERY = TIPMO%
          !PAR_QUERY = TEXMO$
          .Update
       End With
       CONSUSQL.Close
       Combo2.TEXT = AJUSTI$(DESMO$, 46) + "  " + CODMO$
       Call Combo2_Change
       '
20 End If
   '
99 Screen.MousePointer = 1
   Command7.Enabled = True
End Sub

Private Sub Form_Load()
    Call ABRECONEXION
    Call VerConsul
End Sub

Sub VerConsul()
    '
    Combo2.Clear
    If EXISTE_TABLA%("SQL_QUERIES") > 0 Then
        SQLX$ = "SELECT COD_QUERY, DES_QUERY FROM SQL_QUERIES WITH(NOLOCK) ORDER BY COD_QUERY ASC"
        Set RS = READSET(SQLX$)
        With RS
          If Not .EOF Then
            TTXX$ = AJUSTI$(SAFETEXT$(RS!DES_QUERY), 46) + "  " + SAFETEXT$(RS!COD_QUERY)
            Combo2.AddItem TTXX$
            RS.MoveNext
          End If
        End With
    End If
    RS.Close
    Set RS = Nothing
    '
    Call ABRETACON
    STRSQL = "SELECT COD_QUERY, DES_QUERY FROM SQL_QUERIES ORDER BY COD_QUERY ASC"
    On Error Resume Next
    Set MIRS = BaDaConsul.OpenRecordset(STRSQL, dbOpenSnapshot)
    MIRS.MoveFirst
    If Err Then Exit Sub
    '
    On Error GoTo 0
    If Not (MIRS.EOF And MIRS.BOF) Then
       Do While Not MIRS.EOF
          TTXX$ = AJUSTI$(SAFETEXT$(MIRS!DES_QUERY), 46) + "  " + SAFETEXT$(MIRS!COD_QUERY)
          Combo2.AddItem TTXX$
          MIRS.MoveNext
       Loop
    End If
    '
End Sub

Private Sub Combo2_Click()
   Call Combo2_Change
End Sub

Private Sub Combo2_Change()
    '
    CANFI% = 0
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
    '
    STRSQL = "SELECT * FROM SQL_QUERIES where COD_QUERY = '" & Mid$(Combo2.TEXT, 49) & "'"
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

Private Sub EJECUTAR_Click()
    '
    EJECUTAR.Enabled = False
    '
    If Combo2.TEXT = "" Then
        Call MENSERR(24, "No Hay Reporte Activo")
        GoTo 99
    End If
        
    Call CARGA_CONSULTA(Mid$(Combo2.TEXT, 49))
    If TIPQUERY% = 0 Then
        SHOWREP07.Label2 = Mid$(Combo2.TEXT, 49)
        SHOWREP07.Label3 = Label9
        SHOWREP07.Show 1
      ElseIf TIPQUERY% = 1 Then
        SHOWPLA07.Label2 = Mid$(Combo2.TEXT, 49)
        SHOWPLA07.Label3 = Label9
        SHOWPLA07.Show 1
    End If
    '
99  EJECUTAR.Enabled = True
    '
End Sub


