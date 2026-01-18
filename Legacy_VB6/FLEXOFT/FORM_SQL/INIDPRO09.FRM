VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form INIDPRO09 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Generador de Procesos  FLEXOFT"
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
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "Crear Nuevo Proceso"
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
      Begin VB.PictureBox Picture1 
         BorderStyle     =   0  'None
         Height          =   1770
         Left            =   5500
         ScaleHeight     =   1770
         ScaleWidth      =   885
         TabIndex        =   11
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
            Picture         =   "INIDPRO09.frx":0000
            Style           =   1  'Graphical
            TabIndex        =   1
            ToolTipText     =   "Ejecuta la Consulta"
            Top             =   645
            Width           =   645
         End
         Begin VB.Image Image2 
            Height          =   390
            Left            =   -300
            Picture         =   "INIDPRO09.frx":030A
            Top             =   1380
            Width           =   1515
         End
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
         Left            =   3150
         Sorted          =   -1  'True
         TabIndex        =   0
         Top             =   420
         Width           =   2325
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
         TabIndex        =   12
         Top             =   1575
         Width           =   260
      End
      Begin VB.Label Label3 
         Caption         =   "Tipo / Formato  de Proceso:"
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
         TabIndex        =   10
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
         TabIndex        =   9
         Top             =   1575
         Width           =   2535
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
         TabIndex        =   8
         ToolTipText     =   "Doble Click para Nueva Consulta"
         Top             =   585
         Width           =   1065
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         Caption         =   "Proceso"
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
      Caption         =   "Importador de Procesos"
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
Attribute VB_Name = "INIDPRO09"
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
         FGENPROC09.Label1 = Combo2.TEXT
         FGENPROC09.Show 1
       Else
         Call MENSERR(24, "Imposible Editar Consulta de Tpo '" + TRIM$(Str$(TICONSU%)) + "'.")
    End If
    '
99  Command1.Enabled = True
    '
End Sub

Private Sub Command2_Click()
   Command2.Enabled = False
   NEWPROC09.Show 1
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
   ARIMPO$ = UCase$(CommonDialog1.filename)
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
       Call ABRETACON
       STRSQL = "SELECT * FROM SQL_PROCEDS WHERE COD_PROCEX = '" & CODMO$ & "'"
       Set CONSUSQL = BaDaConsul.OpenRecordset(STRSQL, dbOpenDynaset)
       With CONSUSQL
          If (.BOF And .EOF) Then
            .AddNew   ' YA CONVERTIDO
            ON ERROR RESUME NEXT
            !CodiEmpr = Codiemp%
            ON ERROR GOTO 0
            GoTo 15
          End If
          If COMALTER%("Ya Existe una Consulta Registrada\Bajo el Código '" + CODMO$ + "'.\\Cancela Guardar\Sobre-Escribir") <> 2 Then GoTo 20
          .Edit
15        !COD_PROCEX = CODMO$
          !DES_PROCEX = DESMO$
          !TIP_PROCEX = TIPMO%
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
    Call VERJOBID(OKEY%)
    If OKEY% < 1 Then
       Call FINAL("")
       Exit Sub
    End If
    '
    Call ABRECONEXION
    Call VerConsul
End Sub

Sub VerConsul()
    '
    Combo2.Clear
    Call ABRETACON
    STRSQL = "SELECT COD_PROCEX FROM SQL_PROCEDS ORDER BY COD_PROCEX ASC"
    On Error Resume Next
    Set MIRS = BaDaConsul.OpenRecordset(STRSQL, dbOpenSnapshot)
    MIRS.MoveFirst
    If Err Then Exit Sub
    '
    On Error GoTo 0
    If Not (MIRS.EOF And MIRS.BOF) Then
       Do While Not MIRS.EOF
          Combo2.AddItem (MIRS!COD_PROCEX)
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
    STRSQL = "SELECT * FROM SQL_PROCEDS where COD_PROCEX = '" & Combo2.TEXT & "'"
    Set MIRS = BaDaConsul.OpenRecordset(STRSQL, dbOpenSnapshot)
    If Not (MIRS.EOF And MIRS.BOF) Then
       MIRS.MoveFirst
       Label9 = MIRS!DES_PROCEX
       TICONSU% = XVALO(MIRS!TIP_PROCEX)
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
        
    Call CARGA_CONSULTA(Combo2.TEXT)
    If TIPQUERY% = 0 Then
        SHOWREP07.Label2 = Combo2.TEXT
        SHOWREP07.Label3 = Label9
        SHOWREP07.Show 1
      ElseIf TIPQUERY% = 1 Then
        SHOWPLA07.Label2 = Combo2.TEXT
        SHOWPLA07.Label3 = Label9
        SHOWPLA07.Show 1
    End If
    '
99  EJECUTAR.Enabled = True
    '
End Sub


Private Sub Form_Unload(Cancel As Integer)
    Cancel = 0
    Call FINAL("")
End Sub
