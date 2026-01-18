VERSION 5.00
Begin VB.Form FACABON04 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0E0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Selección de Cuentas a Facturar:"
   ClientHeight    =   7290
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   8250
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7290
   ScaleWidth      =   8250
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
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
      Height          =   260
      Left            =   6570
      TabIndex        =   19
      Text            =   " "
      Top             =   105
      Width           =   495
   End
   Begin VB.CommandButton Command3 
      Caption         =   "."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   260
      Left            =   7050
      TabIndex        =   18
      Top             =   105
      Width           =   175
   End
   Begin VB.TextBox Text2 
      Alignment       =   1  'Right Justify
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
      Left            =   6150
      TabIndex        =   15
      Top             =   6825
      Width           =   1020
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
      Height          =   600
      Left            =   840
      MultiLine       =   -1  'True
      TabIndex        =   14
      Top             =   6510
      Width           =   4800
   End
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
      Height          =   4905
      Left            =   120
      Style           =   1  'Checkbox
      TabIndex        =   0
      ToolTipText     =   "Tilde los Clientes a Considerar"
      Top             =   1215
      Width           =   7065
   End
   Begin VB.PictureBox Picture3 
      BackColor       =   &H00E0E0E0&
      Height          =   435
      Left            =   120
      ScaleHeight     =   375
      ScaleWidth      =   7005
      TabIndex        =   8
      Top             =   810
      Width           =   7065
      Begin VB.Label Label22 
         BackStyle       =   0  'Transparent
         Caption         =   "Cuenta    Razón Social"
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
      Height          =   7455
      Left            =   7365
      ScaleHeight     =   7395
      ScaleWidth      =   1035
      TabIndex        =   3
      Top             =   0
      Width           =   1095
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
         Height          =   650
         Left            =   105
         Picture         =   "FACABON04.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   17
         ToolTipText     =   "Borra Marcas de la Última Facturación"
         Top             =   1575
         Width           =   650
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
         Left            =   105
         Picture         =   "FACABON04.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Refresca - Invierte Selección"
         Top             =   2940
         Width           =   650
      End
      Begin VB.PictureBox MARCOBARRA 
         BackColor       =   &H000000FF&
         Height          =   135
         Left            =   120
         ScaleHeight     =   75
         ScaleWidth      =   585
         TabIndex        =   5
         Top             =   5565
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
         Height          =   720
         Left            =   105
         Picture         =   "FACABON04.frx":0454
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Aprueba la Selección"
         Top             =   6405
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
         Picture         =   "FACABON04.frx":075E
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Cancela la Pre-Selección"
         Top             =   120
         Width           =   650
      End
      Begin VB.Image Image2 
         Height          =   390
         Left            =   -315
         Picture         =   "FACABON04.frx":0A68
         Top             =   5880
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
   Begin VB.Label Label23 
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
      Height          =   250
      Left            =   4620
      TabIndex        =   21
      Top             =   420
      Width           =   2565
   End
   Begin VB.Label Label41 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Grupo de Clientes:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   4830
      TabIndex        =   20
      Top             =   150
      Width           =   1710
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Importe:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   5355
      TabIndex        =   16
      Top             =   6510
      Width           =   1170
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Texto:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   -420
      TabIndex        =   13
      Top             =   6510
      Width           =   1170
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   7400
      Y1              =   6315
      Y2              =   6315
   End
   Begin VB.Label Label2 
      Height          =   435
      Left            =   105
      TabIndex        =   12
      Top             =   6720
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Marque (tilde) los Clientes Asociados a Facturar: "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   750
      Left            =   210
      TabIndex        =   11
      Top             =   210
      Width           =   3375
   End
   Begin VB.Line Line3 
      X1              =   120
      X2              =   7180
      Y1              =   6120
      Y2              =   6120
   End
   Begin VB.Line Line2 
      X1              =   7185
      X2              =   7185
      Y1              =   810
      Y2              =   6110
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
End
Attribute VB_Name = "FACABON04"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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

Private Sub Command2_Click()
   If DERACCE%("BORRMAFA", "Borrar Marcas de Facturación") = 2 Then
      If COMALTER%("¿ Está Seguro de Borrar las Marcas de Facturación\e Iniciar un Nuevo Lote Mensual ?\\No, Cancelar Transacción\Si, Re-Inicializar") = 2 Then
        Screen.MousePointer = 11
        Call APERBASDAT("CLIEN")
        SQLX$ = "UPDATE Deudor12 SET Mar_Ya_Fac = 0 "
        SQLX$ = SQLX$ + "WHERE Mar_Ya_Fac > 0 "
        Badaclie.Execute (SQLX$)
        Call GRACONTROL("FACABONO", "LEYTEXTO", "")
        FACABON04.Text1 = ""
        Call GRACONTROL("FACABONO", "IMPORTEM", "")
        FACABON04.Text2 = ""
        Screen.MousePointer = 1
        Call COMUNI("Re-Inicialización Completada")
        Label2 = "FRESH"
        Hide
      End If
   End If
End Sub

Private Sub Command3_Click()
   Call SELECHO("GRUCOCLI")
   VDEVU$ = VALACT1$("GRUCOCLI")
   If TRIM$(VDEVU$) <> "" Then
      Text3.TEXT = VDEVU$
   End If
End Sub

Sub Text3_Change()
    '
    FACABON04.LICOPEN.Clear
    If XVALO(Text3) < 1 Then
         Label23 = "Todos los Clientes"
      ElseIf XVALO(Text3) < 32767 Then
         Label23 = ECOARCH("GRUCOCLI", MKI$(XVALO(Text3)))
      Else
         Call MENSERR(24, "Grupo no Válido")
         Exit Sub
    End If
    '
5   Screen.MousePointer = 11
    Call APERBASDAT("CLIEN")
    SQLX$ = "SELECT * FROM Deudor12 "
    SQLX$ = SQLX$ + "WHERE Stat_Cli > 0 "
    If XVALO(Text3) > 0 Then
       SQLX$ = SQLX$ + "AND GrCo_Cli = " & Text3 & " "
    End If
    SQLX$ = SQLX$ + "AND (Mar_Ya_Fac < 1 OR ISNULL(Mar_Ya_Fac) = TRUE) "
    SQLX$ = SQLX$ + "order by Nume_Cli ASC"
    Set ClienTemp = Badaclie.OpenRecordset(SQLX$, dbOpenSnapshot)
    With ClienTemp
      On Error Resume Next '
       .MoveLast
       If Err Then
         On Error GoTo 0
         Call MENSERR(24, "No hay Clientes Activos Pendientes de Facturación.")
         GoTo 99
       End If
       '
       FIN& = .RecordCount
       .MoveFirst
       Do While Not .EOF
          TTXX$ = FORMATNUM$(!Nume_Cli, "f6.0") + "    " + !Raso_Cli
          FACABON04.LICOPEN.AddItem TTXX$
       .MoveNext
       Loop
    End With
    '
99  Screen.MousePointer = 1
    '
End Sub

Private Sub Text3_DblClick()
   Call Command3_Click
End Sub
