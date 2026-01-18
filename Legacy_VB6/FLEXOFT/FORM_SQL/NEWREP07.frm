VERSION 5.00
Begin VB.Form NEWREP07 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Nuevo Reporte  FLEXOFT"
   ClientHeight    =   2385
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6390
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
   ScaleHeight     =   2385
   ScaleWidth      =   6390
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Definicion del Reporte"
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
      TabIndex        =   4
      Top             =   210
      Width           =   6180
      Begin VB.TextBox Text2 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   570
         Left            =   210
         MultiLine       =   -1  'True
         TabIndex        =   1
         Top             =   840
         Width           =   5055
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
         Left            =   2940
         MaxLength       =   16
         TabIndex        =   0
         Top             =   420
         Width           =   2325
      End
      Begin VB.PictureBox Picture1 
         BorderStyle     =   0  'None
         Height          =   1770
         Left            =   5270
         ScaleHeight     =   1770
         ScaleWidth      =   885
         TabIndex        =   8
         Top             =   200
         Width           =   880
         Begin VB.CommandButton GUARDAR 
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
            Height          =   1005
            Left            =   105
            Picture         =   "NEWREP07.frx":0000
            Style           =   1  'Graphical
            TabIndex        =   3
            ToolTipText     =   "Aprueba y Crea la Consulta"
            Top             =   225
            Width           =   645
         End
         Begin VB.Image Image2 
            Height          =   390
            Left            =   -300
            Picture         =   "NEWREP07.frx":030A
            Top             =   1380
            Width           =   1515
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
         Left            =   2625
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   1575
         Width           =   2650
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
         TabIndex        =   7
         ToolTipText     =   "Doble Click para Nueva Consulta"
         Top             =   1620
         Width           =   2430
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
         Left            =   105
         TabIndex        =   6
         ToolTipText     =   "Doble Click para Nueva Consulta"
         Top             =   585
         Width           =   1065
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
         Left            =   1785
         TabIndex        =   5
         ToolTipText     =   "Doble Click para Nueva Consulta"
         Top             =   465
         Width           =   1065
      End
   End
End
Attribute VB_Name = "NEWREP07"
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

Private Sub Form_Load()
    '
    Combo1.Clear
    Combo1.AddItem "  0.- Tabla Plana"
    Combo1.AddItem "  1.- Estadistica Mensual"
    Combo1.AddItem "  9.- Grafico"
    '
End Sub

Sub VerConsul()
    '
    Combo2.Clear
    If EXISTE_TABLA%("SQL_QUERIES") > 0 Then
        SQLX$ = "SELECT COD_QUERY FROM SQL_QUERIES WITH(NOLOCK) ORDER BY COD_QUERY ASC"
        Set RS = READSET(SQLX$)
        With RS
          If Not .EOF Then
            Combo2.AddItem (RS!COD_QUERY)
            MIRS.MoveNext
          End If
        End With
    End If
    RS.Close
    Set RS = Nothing
    
    Call ABRETACON
    STRSQL = "SELECT COD_QUERY FROM SQL_QUERIES ORDER BY COD_QUERY ASC"
    On Error Resume Next
    Set MIRS = BaDaConsul.OpenRecordset(STRSQL, dbOpenSnapshot)
    MIRS.MoveFirst
    If Err Then Exit Sub
    '
    On Error GoTo 0
    If Not (MIRS.EOF And MIRS.BOF) Then
       Do While Not MIRS.EOF
          Combo2.AddItem (MIRS!COD_QUERY)
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

Private Sub GUARDAR_Click()
    '
    GUARDAR.Enabled = False
    '
    CODMO$ = TRIM$(Text1)
    If CODMO$ = "" Then
        Call MENSERR(24, "Falta Codigo de Consulta")
        GoTo 99
    End If
    '
    DESMO$ = TRIM$(Text2)
    If DESMO$ = "" Then
        Call MENSERR(24, "Falta Descripcion de Consulta")
        GoTo 99
    End If
    '
    If TRIM$(Combo1.TEXT) = "" Then
        Call MENSERR(24, "Falta Seleccionar Tipo de Reporte")
        GoTo 99
    End If
    '
    TISMO% = XVALO(Left$(Combo1.TEXT, 3))
    If TISMO% <> 0 Then
      If TISMO% <> 1 Then
        If TISMO% <> 9 Then
          Call MENSERR(24, "Tipo de Reporte no Definido")
          GoTo 99
        End If
      End If
    End If
    '
    GRABOCO% = 0
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
        If COMALTER%("Ya Existe una Consulta Registrada\Bajo El Código '" + CODMO$ + "'.\\Cancela Guardar\Sobre-Escribir") <> 2 Then GoTo 20
        If COMALTER%("¿ Está Seguro de Sobre-Escribir el Reporte Pr-Existente ?\\No, Cancelo Guardar\Si, Sobre-Escribir") <> 2 Then GoTo 20
        .Edit
15      !COD_QUERY = CODMO$
        !DES_QUERY = DESMO$
        !TIP_QUERY = TISMO%
        !PAR_QUERY = ""
        .Update
        GRABOCO% = 1
    End With
    '
20  On Error Resume Next
    CONSUSQL.Close
    On Error GoTo 0
    Set CONSUSQL = Nothing
    '
    'ATENCION *** CUANDO NO SE GRABE MAS EN MDB HAY QUE AGREGAR EL COMALTER QUE SE ENCUENTRA EN LA CONSULTA MDB ARRIBA.
    If EXISTE_TABLA%("SQL_QUERIES") > 0 Then
       Call GRABA_CONSULTA_SQL(CODMO$, DESMO$, TISMO%, "", OK%)
       If OK% >= 0 Then GRABOCO% = 1
    End If

    If GRABOCO% = 1 Then
       Call DEFTREP07.VerConsul
       DEFTREP07.Combo2.TEXT = AJUSTI$(DESMO$, 46) + "  " + CODMO$
       Unload Me
    End If
    '
99  GUARDAR.Enabled = True
    '
End Sub


