VERSION 5.00
Begin VB.Form NECOMAR04 
   BackColor       =   &H00C0E0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Materiales a Comprar segun Calculo de Disponibilidad / MRP"
   ClientHeight    =   5145
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   10170
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5145
   ScaleWidth      =   10170
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture6 
      Height          =   5150
      Left            =   9050
      ScaleHeight     =   5085
      ScaleWidth      =   1080
      TabIndex        =   16
      Top             =   0
      Visible         =   0   'False
      Width           =   1140
      Begin VB.Image Image3 
         Height          =   5150
         Left            =   0
         Picture         =   "NECOMAR04.frx":0000
         Stretch         =   -1  'True
         Top             =   0
         Width           =   1095
      End
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00004000&
      Height          =   6165
      Left            =   9060
      ScaleHeight     =   6105
      ScaleWidth      =   1080
      TabIndex        =   9
      Top             =   -195
      Width           =   1140
      Begin VB.CommandButton BOTEXIT 
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
         Left            =   240
         Picture         =   "NECOMAR04.frx":2398
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Cancela la Pre-Selección"
         Top             =   330
         Width           =   650
      End
      Begin VB.CommandButton BOTREC 
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
         Left            =   240
         Picture         =   "NECOMAR04.frx":26A2
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Aprueba la Pre-Selección"
         Top             =   1005
         Width           =   650
      End
      Begin VB.PictureBox MARCOBARRA 
         BackColor       =   &H000000FF&
         Height          =   135
         Left            =   210
         ScaleHeight     =   75
         ScaleWidth      =   585
         TabIndex        =   11
         Top             =   4365
         Width           =   650
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H00FF0000&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   75
            TabIndex        =   12
            Top             =   75
            Width           =   12000
         End
      End
      Begin VB.CommandButton Command1 
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
         Left            =   240
         Picture         =   "NECOMAR04.frx":29AC
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Refresca - Invierte Selección"
         Top             =   1695
         Width           =   650
      End
      Begin VB.Image Image2 
         Height          =   390
         Left            =   -225
         Picture         =   "NECOMAR04.frx":2AF6
         Top             =   4620
         Width           =   1515
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0E0C0&
      Caption         =   "Proveedores"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   1515
      TabIndex        =   2
      Top             =   4140
      Width           =   6075
      Begin VB.CommandButton Command3 
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
         Left            =   855
         TabIndex        =   8
         ToolTipText     =   "Proveedores"
         Top             =   330
         Width           =   195
      End
      Begin VB.TextBox Text3 
         BackColor       =   &H8000000F&
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
         Left            =   90
         MaxLength       =   5
         TabIndex        =   6
         Text            =   " "
         Top             =   330
         Width           =   750
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
         Height          =   405
         Left            =   5625
         Picture         =   "NECOMAR04.frx":4A18
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Refresca Lista de Selección"
         Top             =   240
         Width           =   315
      End
      Begin VB.CommandButton Command2 
         Caption         =   "!"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   12
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1095
         TabIndex        =   4
         ToolTipText     =   "Proveedores Habituales"
         Top             =   330
         Width           =   195
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H00E0E0E0&
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
         Left            =   1410
         TabIndex        =   3
         TabStop         =   0   'False
         Text            =   " "
         Top             =   330
         Width           =   4065
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0E0C0&
      Caption         =   "Necesidades"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4005
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8970
      Begin VB.ListBox LIMATSEL 
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3630
         ItemData        =   "NECOMAR04.frx":4B62
         Left            =   75
         List            =   "NECOMAR04.frx":4B64
         Style           =   1  'Checkbox
         TabIndex        =   1
         Top             =   225
         Width           =   8820
      End
   End
   Begin VB.Label Label2 
      Height          =   270
      Left            =   120
      TabIndex        =   15
      Top             =   4725
      Visible         =   0   'False
      Width           =   1065
   End
   Begin VB.Label Label1 
      Height          =   285
      Left            =   120
      TabIndex        =   7
      Top             =   4410
      Visible         =   0   'False
      Width           =   1065
   End
   Begin VB.Menu menu 
      Caption         =   "Menú"
      Begin VB.Menu revertirseleccion 
         Caption         =   "Revertir Selección"
      End
      Begin VB.Menu as 
         Caption         =   "-"
      End
      Begin VB.Menu aprobarseleccion 
         Caption         =   "Aprobar Selección"
      End
      Begin VB.Menu a 
         Caption         =   "-"
      End
      Begin VB.Menu salir 
         Caption         =   "Salir"
      End
   End
End
Attribute VB_Name = "NECOMAR04"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub aprobarseleccion_Click()
    Call BOTREC_Click
End Sub

Private Sub BOTEXIT_Click()
    LIMATSEL.Clear
    Label1.Caption = "CANCELO"
    Label2.Caption = ""
    FORCOMP04.Text1.TEXT = ""
    Hide
End Sub

Private Sub BOTREC_Click()
     Hide
     Label2.Caption = "ACEPTO"
End Sub

Private Sub Command1_Click()
    '
    Command1.Enabled = False
    TPI = LIMATSEL.TopIndex
    For U& = 1 To LIMATSEL.ListCount
      If LIMATSEL.Selected(U& - 1) = True Then
          LIMATSEL.Selected(U& - 1) = False
        Else
          LIMATSEL.Selected(U& - 1) = True
      End If
    Next U&
    On Error Resume Next
    LIMATSEL.TopIndex = TPI
    On Error GoTo 0
    DoEvents
    Command1.Enabled = True
    '
End Sub

Private Sub Command2_Click()
     '
     Call SELECHO("PROVHAB/Proveedores Habituales")
        NPR% = XVALO(TRIM$(VALACT1$("PROVHAB")))
        If NPR% > 0 Then
           Text3.TEXT = TRIM$(Str$(NPR%))
        End If
     '
End Sub

Private Sub Command3_Click()
     Call SELECHO("ACREDOR")
     NPR% = XVALO(TRIM$(VALACT1$("ACREDOR")))
     If NPR% > 0 Then
          Text3.TEXT = TRIM$(Str$(NPR%))
     End If
End Sub

Sub Command4_Click()

    Call COPYSTRU("ACREDOR", "PROVHAB")
    Call BLANARCH("PROVHAB")
    '
    CONDI$ = ""
    For I& = 0 To LIMATSEL.ListCount - 1
        If LIMATSEL.Selected(I&) = True Then
            CODIGO$ = Mid$(LIMATSEL.List(I&), 1, 16)
            CONDI1$ = CONDI1$ & "Cod_Inte=" & CODINT%(CODIGO$) & " OR "
        End If
    Next I&
    NN% = Len(CONDI1$)
    If CONDI1$ <> "" Then
        CONDI$ = Mid$(CONDI1$, 1, NN% - 3)
        '
        Call ABRECOMPRAS
        Dim RS As Recordset
        SQLX$ = "SELECT DISTINCT NPRO_OCOM FROM OCOMDETA "
        SQLX$ = SQLX$ + "WHERE " & CONDI$ & " "
        Set RS = compras.OpenRecordset(SQLX$, dbOpenSnapshot)
        '
        With RS
            On Error Resume Next
            .MoveFirst
            On Error GoTo 0
            REG& = 0
            Do While .EOF = False
                REG& = REG& + 1
                NPRO% = XVALO(!NPRO_OCOM)
                XX$ = REGBLAN$("PROVHAB")
                Call REPLA(XX$, MKI$(NPRO%), 1, 2)
                Call REPLA(XX$, RASOCLI$((-1) * NPRO%), 3, 30)
                Call GRAREG("PROVHAB", XX$, REG&)
            .MoveNext
            Loop
        End With
        '
        Call SETULTREG("PROVHAB", REG&)
        Call CIERRARCH("PROVHAB")
        Call FRESHECHO("PROVHAB")
        '
    Else
        Exit Sub
    End If
    '
End Sub

Private Sub Form_Load()
    If CONTROL("NEWASPEC", "CAMBIOAS") = "SI" Then
      Picture6.Visible = True
      Me.BackColor = &HFCD7B6
      Me.Frame1.BackColor = &HFCD7B6
      Me.Frame2.BackColor = &HFCD7B6
      MARCOBARRA.Top = 4800
      Set MARCOBARRA.Container = Me.Picture6
    End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
    LIMATSEL.Clear
    Label1.Caption = "CANCELO"
    Hide
End Sub

Private Sub LIMATSEL_ItemCheck(Item As Integer)
    Call Command4_Click
    Text3 = ""
End Sub

Private Sub revertirseleccion_Click()
    Call Command1_Click
End Sub

Private Sub salir_Click()
    Call BOTEXIT_Click
End Sub

Private Sub Text3_Change()
    '
    If XVALO(Text3.TEXT) <= 32767 Then
       NPR% = (-1) * XVALO(Text3.TEXT)
       Text2.TEXT = RASOCLI$(NPR%)
    End If
    '
End Sub

Private Sub Text3_GotFocus()
      '
      Command2.SetFocus
      '
End Sub
