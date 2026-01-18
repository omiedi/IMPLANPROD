VERSION 5.00
Begin VB.Form CLICOPE04 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0E0FF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Selección de Clientes con Pedidos Pendientes "
   ClientHeight    =   3615
   ClientLeft      =   45
   ClientTop       =   585
   ClientWidth     =   8250
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3615
   ScaleWidth      =   8250
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox LICOPEN 
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2355
      Left            =   120
      Style           =   1  'Checkbox
      TabIndex        =   0
      ToolTipText     =   "Tilde los Clientes a Considerar"
      Top             =   1110
      Width           =   7065
   End
   Begin VB.PictureBox Picture3 
      BackColor       =   &H00E0E0E0&
      Height          =   435
      Left            =   120
      ScaleHeight     =   375
      ScaleWidth      =   7005
      TabIndex        =   8
      Top             =   705
      Width           =   7065
      Begin VB.Label Label22 
         BackStyle       =   0  'Transparent
         Caption         =   "Código        Razón Social"
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404000&
         Height          =   225
         Left            =   220
         TabIndex        =   9
         Top             =   105
         Width           =   10515
      End
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00004000&
      Height          =   5775
      Left            =   7365
      ScaleHeight     =   5715
      ScaleWidth      =   1035
      TabIndex        =   3
      Top             =   0
      Width           =   1095
      Begin VB.PictureBox Picture1 
         Height          =   3600
         Left            =   0
         ScaleHeight     =   3540
         ScaleWidth      =   1005
         TabIndex        =   13
         Top             =   0
         Visible         =   0   'False
         Width           =   1065
         Begin VB.Image Image3 
            Height          =   3600
            Left            =   0
            Picture         =   "CLICOPE04.frx":0000
            Stretch         =   -1  'True
            Top             =   0
            Width           =   1065
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
         Left            =   120
         Picture         =   "CLICOPE04.frx":2398
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Refresca - Invierte Selección"
         Top             =   2100
         Width           =   650
      End
      Begin VB.PictureBox MARCOBARRA 
         BackColor       =   &H000000FF&
         Height          =   135
         Left            =   120
         ScaleHeight     =   75
         ScaleWidth      =   585
         TabIndex        =   5
         Top             =   1740
         Width           =   650
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H00FF0000&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   0
            TabIndex        =   6
            Top             =   0
            Width           =   12000
         End
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
         Left            =   120
         Picture         =   "CLICOPE04.frx":24E2
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Aprueba la Selección"
         Top             =   2775
         Width           =   650
      End
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
         Left            =   120
         Picture         =   "CLICOPE04.frx":27EC
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Cancela la Pre-Selección"
         Top             =   120
         Width           =   650
      End
      Begin VB.Image Image2 
         Height          =   390
         Left            =   -300
         Picture         =   "CLICOPE04.frx":2AF6
         Top             =   1200
         Width           =   1515
      End
   End
   Begin VB.ListBox List1 
      Height          =   645
      Left            =   10200
      Sorted          =   -1  'True
      TabIndex        =   2
      Top             =   5400
      Visible         =   0   'False
      Width           =   1065
   End
   Begin VB.Label Label2 
      Height          =   435
      Left            =   6615
      TabIndex        =   12
      Top             =   210
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Marque (tilde) los clientes cuyos pedidos desea incluir en el proceso de cálculo MRP. "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   750
      Left            =   210
      TabIndex        =   11
      Top             =   105
      Width           =   6945
   End
   Begin VB.Line Line3 
      X1              =   120
      X2              =   7200
      Y1              =   3460
      Y2              =   3460
   End
   Begin VB.Line Line2 
      X1              =   7180
      X2              =   7180
      Y1              =   700
      Y2              =   3450
   End
   Begin VB.Label Label17 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404000&
      Height          =   330
      Left            =   7440
      TabIndex        =   10
      Top             =   1140
      Width           =   1695
   End
   Begin VB.Menu archivo 
      Caption         =   "Archivo"
      Begin VB.Menu aprobar 
         Caption         =   "Aprobar"
      End
      Begin VB.Menu s 
         Caption         =   "-"
      End
      Begin VB.Menu revertseleccion 
         Caption         =   "Revertir Selección"
      End
      Begin VB.Menu ss 
         Caption         =   "-"
      End
      Begin VB.Menu salir 
         Caption         =   "Salir"
      End
   End
End
Attribute VB_Name = "CLICOPE04"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub aprobar_Click()
   Call BOTREC_Click
End Sub

Private Sub BOTEXIT_Click()
   LICOPEN.Clear
   Unload Me
End Sub
'

Private Sub BOTREC_Click()
   Label2 = "OK"
   Hide
End Sub

Private Sub Command1_Click()
   Command1.Enabled = False
   TPI = LICOPEN.TopIndex
   For U& = 1 To LICOPEN.ListCount
     If LICOPEN.Selected(U& - 1) = True Then
         LICOPEN.Selected(U& - 1) = False
       Else
         LICOPEN.Selected(U& - 1) = True
     End If
   Next U&
   On Error Resume Next
   LICOPEN.TopIndex = TPI
   On Error GoTo 0
   DoEvents
   Command1.Enabled = True
End Sub

Private Sub Command9_Click()
   Call SELECHO("SELFECHA")
   VDEVU$ = TRIM$(VALACT1$("SELFECHA"))
   If VDEVU$ <> "" Then
      Label2 = VALACT1$("SELFECHA")
   End If
End Sub

Private Sub Label14_Change()
   NUOCO& = Val(Label14)
   If NUOCO& > 0 Then
   Call ABRECOMPRAS
   SQLX$ = " SELECT Npro_Ocom, Femi_Ocom, AtSr_Ocom "
   SQLX$ = SQLX$ + " FROM OCOMENCA "
   SQLX$ = SQLX$ + " WHERE Nume_Ocom =   " & NUOCO& & "     "
   Set COMPLETEMP = compras.OpenRecordset(SQLX$, dbOpenSnapshot)
   With COMPLETEMP
     On Error Resume Next
     .MoveFirst
     If Err Then
     Call MENSERR(24, "Imposible Presentar O/C + Str$(nuoco&) ")
     Exit Sub
     End If
     NPRO% = !Npro_Ocom
     Label15 = Format(!Femi_Ocom, "dd/mm/yy")
     Label9 = !AtSr_Ocom
   End With
   Set COMPLETEMP = Nothing
   '
   label3 = NPRO%
   NPRO1% = (-1) * NPRO%
   '
   Label4 = RASOCLI(NPRO1%)
   Label6 = DOMICLI(NPRO1%)
   Label7 = CPOSCLI(NPRO1%)
   Label8 = LOCACLI(NPRO1%)
   Label10 = PIVACLI(NPRO1%)
   Label11 = CUITCLI(NPRO1%)
End If
   '
Call ABRECOMPRAS
'
SQLX$ = " SELECT * "
SQLX$ = SQLX$ + " FROM OCOMDETA "
SQLX$ = SQLX$ + " WHERE  Nume_Ocom =   " & NUOCO& & " "
SQLX$ = SQLX$ + " AND Stat_Ocom < 9 "
'
Set ITETEMP = compras.OpenRecordset(SQLX$, dbOpenSnapshot)
'
With ITETEMP
  .MoveFirst
  LICOPEN.Clear
  Do While (.EOF = False)
    TTXX$ = Space$(128)
    If !Stat_Ocom < 9 Then
     CII% = !COD_INTE
    End If
    COEFCAN = !Coef_Unids
    If COEFCAN < 0.005 Then
     COEFCAN = 1
    End If
    CAORI = !Cant_Ocom * COEFCAN
    CAREC = !Cant_Rec * COEFCAN
    CAPEN = CAORI - CAREC: If CAPEN < 0 Then CAPEN = 0
    UNICO$ = !Uni_Mas
    COEFUNI = !Coef_Unids
    STAT% = !Stat_Ocom
    NUORIT% = !NuOrd_Item
    '
    Call REPLA(TTXX$, CODEXT$(CII%), 1, 16)
    Call REPLA(TTXX$, DESCRIT0$(CII%), 18, 30)
    Call REPLA(TTXX$, UNICO$, 49, 4)
    Call REPLA(TTXX$, FORMATNUM$(CAORI, "F10.1"), 54, 10)
    Call REPLA(TTXX$, FORMATNUM$(CAREC, "F10.1"), 64, 10)
    Call REPLA(TTXX$, FORMATNUM$(CAPEN, "F10.1"), 74, 10)
    '
    Call REPLA(TTXX$, Str$(NUORIT%), 121, 8)  ' numero de orden del item
    '
    If STAT% = 8 Then
     Call REPLA(TTXX$, "Cumplido", 76, 10)
    End If
    LICOPEN.AddItem TTXX$
    .MoveNext
  Loop
End With
'
End Sub

Private Sub Form_Load()

    If Control("NEWASPEC", "CAMBIOAS") = "SI" Then
      Me.Picture1.Visible = True
      Me.BackColor = &HFCD7B6
      Me.Label1.BackColor = &HFCD7B6
      Me.Label17.BackColor = &HFCD7B6
      Me.Label2.BackColor = &HFCD7B6
      Me.Label22.BackColor = &HFCD7B6
      Me.MARCOBARRA.Top = 3300
      Set MARCOBARRA.Container = Me.Picture1
      MARCOBARRA.ZOrder 0
    End If

End Sub

Private Sub revertseleccion_Click()
   Call Command1_Click
End Sub

Private Sub salir_Click()
    Call BOTEXIT_Click
End Sub
