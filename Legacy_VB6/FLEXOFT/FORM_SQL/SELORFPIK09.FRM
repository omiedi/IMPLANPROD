VERSION 5.00
Begin VB.Form SELORFPIK09 
   BackColor       =   &H00FFFFC0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Selección Lista de Picking"
   ClientHeight    =   3795
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8205
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3795
   ScaleWidth      =   8205
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture14 
      BackColor       =   &H00400000&
      Height          =   2760
      Left            =   7320
      ScaleHeight     =   2700
      ScaleWidth      =   1155
      TabIndex        =   11
      Top             =   0
      Width           =   1215
      Begin VB.CommandButton AGenRep 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   650
         Left            =   100
         Picture         =   "SELORFPIK09.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "Acceso a Consultas Pre-Configuradas"
         Top             =   1200
         Width           =   650
      End
      Begin VB.PictureBox MARCOBARRA 
         BackColor       =   &H000000FF&
         Height          =   135
         Left            =   100
         ScaleHeight     =   75
         ScaleWidth      =   585
         TabIndex        =   13
         Top             =   2000
         Width           =   650
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H00FF0000&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   255
            TabIndex        =   14
            Top             =   0
            Width           =   12000
         End
      End
      Begin VB.CommandButton Command4 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   690
         Left            =   100
         Picture         =   "SELORFPIK09.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Salir"
         Top             =   135
         Width           =   660
      End
      Begin VB.Image Image2 
         Height          =   390
         Left            =   -320
         Picture         =   "SELORFPIK09.frx":0454
         ToolTipText     =   "Acceso Directo a Otras Aplicaciones"
         Top             =   2310
         Width           =   1515
      End
   End
   Begin VB.CommandButton Command2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   780
      Left            =   7320
      Picture         =   "SELORFPIK09.frx":2376
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "Graba - Guardar Modificaciones"
      Top             =   2880
      Width           =   800
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Ordenes de Producción en Curso"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2190
      Left            =   210
      TabIndex        =   8
      Top             =   960
      Width           =   6975
      Begin VB.ListBox List1 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Lucida Console"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1740
         ItemData        =   "SELORFPIK09.frx":2680
         Left            =   105
         List            =   "SELORFPIK09.frx":2682
         MousePointer    =   1  'Arrow
         Sorted          =   -1  'True
         Style           =   1  'Checkbox
         TabIndex        =   9
         Top             =   300
         Width           =   6720
      End
   End
   Begin VB.CommandButton Command1 
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
      Height          =   285
      Left            =   1830
      TabIndex        =   4
      Top             =   3345
      Width           =   175
   End
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
      Height          =   285
      Left            =   6960
      TabIndex        =   0
      Top             =   180
      Width           =   175
   End
   Begin VB.Label Label4 
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
      Left            =   2205
      TabIndex        =   7
      Top             =   3360
      Width           =   4980
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Deposito:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   180
      TabIndex        =   6
      Top             =   3390
      Width           =   870
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
      Left            =   1140
      TabIndex        =   5
      Top             =   3345
      Width           =   690
   End
   Begin VB.Label Label1 
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
      Left            =   240
      TabIndex        =   3
      Top             =   600
      Width           =   6915
   End
   Begin VB.Label Label22 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Código de Producto:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   345
      TabIndex        =   2
      Top             =   210
      Width           =   4665
   End
   Begin VB.Label Label28 
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
      Left            =   5115
      TabIndex        =   1
      Top             =   180
      Width           =   1860
   End
End
Attribute VB_Name = "SELORFPIK09"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Ok As Boolean  'Para determinar si Aprueba o Sale

Private Sub Command1_Click()
  Command1.Enabled = False
  '
  Call SELECHO("TADEPOSI")
  DEPO$ = VALACT1$("TADEPOSI")
    
  If DEPO$ <> "" Then
     Label2 = DEPO$
     Label4 = VALACT2$("TADEPOSI")
  End If
  '
  Command1.Enabled = True
End Sub

Private Sub Command2_Click()
   Command2.Enabled = False
   If List1.SelCount < 1 Then
        Unload Me
        Exit Sub
     ElseIf List1.SelCount > 1 Then
        Call MENSERR(24, "Solo Puede Seleccionarse una Unica O/Produccion.")
        Command2.Enabled = True
        Exit Sub
   End If
   Ok = True
   Command2.Enabled = True
   Hide
End Sub

Private Sub Command4_Click()
   Unload Me
   
End Sub

Private Sub Command5_Click()
    Command5.Enabled = False
    '
    Call SELECHO("MASTER")
    COD$ = VALACT1("MASTER")
    If COD$ <> "" Then
       Label28 = COD$
       Label1 = DESCRIT$(CODINT(COD$))
    End If
    '
    Command5.Enabled = True
End Sub


Private Sub Form_Activate()
Ok = False
End Sub


Sub Label28_Change()
   '
   List1.Clear
   CI% = CODINT(Label28)
   If CI% < 1 Then Exit Sub
   '
   SQLX$ = " SELECT * FROM OrdeFabr WHERE Cod_Inte =" & CI% & " AND StatuOrf = 1 AND (NuOrProd <> '' AND NuOrProd IS NOT NULL)"
   Dim ORDEFTMP As ADODB.Recordset
   Set ORDEFTMP = New ADODB.Recordset
   Set ORDEFTMP = FLEXCONN.Execute(FILTSQL$(SQLX$))
   
   With ORDEFTMP
      If .EOF And .BOF Then
         Call MENSERR(24, "No Hay Ordenes De Producción Activas para Este Producto")
         Exit Sub
      End If
      .MoveFirst
      STRFILA$ = ""
      '
      Do While Not (.EOF)
         SOLFA$ = TRIM$(SAFETEXT$(!IDSUBOR))
         NORPRO$ = TRIM$(SAFETEXT$(!NuOrProd))
         CONDI$ = "IDENREQUE LIKE '" & SOLFA$ & "%' and NULISPICK > 0"
         CANRE& = CantRegistros("BALAREQMO", CONDI$)
         '
         If CANRE& > 0 Then GoTo 19 ' ASÍ SUPRIME LOS QUE NO TIENEN ASIGNACION
            REFE$ = SAFETEXT$(!REFEREN)
            FEGEN$ = FECHATEX$(CVINT(!FechGen))
            Call REPLA(STRFILA$, NORPRO$, 1, 10)
            Call REPLA(STRFILA$, " - " & FEGEN$, 11, 12)
            Call REPLA(STRFILA$, " - " & REFE$, 22, 60)
            List1.AddItem STRFILA$
19    .MoveNext
      Loop
   End With
   ORDEFTMP.Close
   Set ORDEFTMP = Nothing
   '
End Sub

Private Sub AGenRep_Click()
    '
    AGenRep.Enabled = False
    Call ACONSUL_PRE(TRIM$(UCase$(Name)))
    AGenRep.Enabled = True
    '
End Sub

Private Sub AGenRep_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   '
   If Button = 2 Then
      Call CONECTA_CONSUL_PRE(TRIM$(UCase$(Name)))
   End If
   '
End Sub

Public Sub BLANCFORMU()
  Label28 = ""
  Label2 = ""
  Label4 = ""
  Label1 = ""
  List1.Clear
  
End Sub

