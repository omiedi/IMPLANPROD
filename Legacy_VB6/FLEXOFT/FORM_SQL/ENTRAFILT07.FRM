VERSION 5.00
Begin VB.Form ENTRAFILT07 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Consulta SQL - Filtros"
   ClientHeight    =   5085
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6465
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5085
   ScaleWidth      =   6465
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00000080&
      Height          =   7575
      Left            =   5565
      ScaleHeight     =   7515
      ScaleWidth      =   1425
      TabIndex        =   5
      Top             =   0
      Width           =   1485
      Begin VB.PictureBox MARCOBARRA 
         BackColor       =   &H000000FF&
         Height          =   135
         Left            =   105
         ScaleHeight     =   75
         ScaleWidth      =   585
         TabIndex        =   7
         Top             =   6090
         Width           =   640
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H00FF0000&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   0
            TabIndex        =   8
            Top             =   105
            Width           =   12000
         End
      End
      Begin VB.CommandButton Command11 
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
         Picture         =   "ENTRAFILT07.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Ejecuta Consulta"
         Top             =   105
         Width           =   640
      End
      Begin VB.CommandButton Command1 
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
         Height          =   650
         Left            =   105
         Picture         =   "ENTRAFILT07.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Cierra - Abandona la Aplicación"
         Top             =   105
         Width           =   640
      End
      Begin VB.Image Image2 
         Height          =   390
         Left            =   -290
         Picture         =   "ENTRAFILT07.frx":0454
         Top             =   840
         Width           =   1515
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "..."
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
      Index           =   0
      Left            =   4335
      TabIndex        =   1
      ToolTipText     =   "Lista de Valores Posibles"
      Top             =   945
      Visible         =   0   'False
      Width           =   330
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Index           =   0
      Left            =   1785
      TabIndex        =   0
      Top             =   945
      Visible         =   0   'False
      Width           =   2445
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Label3"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   255
      Left            =   315
      TabIndex        =   4
      Top             =   210
      Width           =   690
   End
   Begin VB.Label Label2 
      BorderStyle     =   1  'Fixed Single
      Height          =   330
      Left            =   3255
      TabIndex        =   3
      Top             =   1470
      Visible         =   0   'False
      Width           =   1485
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
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
      Height          =   195
      Index           =   0
      Left            =   1020
      TabIndex        =   2
      Top             =   945
      Visible         =   0   'False
      Width           =   675
   End
End
Attribute VB_Name = "ENTRAFILT07"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub command1_Click()
   Unload Me
End Sub

Private Sub Command11_Click()
     SHOWREP07.Show 1
End Sub

Private Sub Command2_Click(Index As Integer)
   '
   CAMCONDI$ = CAMPOCOND$(Index + 1)
   If CAMCONDI$ <> "" Then
      Screen.MousePointer = 11
      SQLX$ = "SELECT DISTINCT " & CAMCONDI$ & " AS VALOR FROM " & TABSELE$ & " WITH(NOLOCK)"
      SQLX$ = SQLX$ + " ORDER BY " & CAMCONDI$ & " ASC"
      '
      Combo1(Index).Clear
      Dim VAPOSI As ADODB.Recordset
      Set VAPOSI = New ADODB.Recordset
      Set VAPOSI =  FLEXCONN.EXECUTE (FILTSQL$ (SQLX$))
      With VAPOSI
         Do While Not .EOF
            On Error Resume Next
            Combo1(Index).AddItem !VALOR
            On Error GoTo 0
         .MoveNext
         Loop
      End With
      Screen.MousePointer = 1
   End If
   '
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Cancel = 0
End Sub

Private Sub Label2_Change()
    If TRIM$(Label2.Caption) <> "" Then
       Call FILTROVAR
    End If
End Sub

Sub FILTROVAR()
    '
    CACONFIL% = 0
    '
    If CACONDI% < 1 Then
       Call Command11_Click
       Hide
       Exit Sub
    End If
    '
    PVER = 1000
    JK& = 0
    TAMAXI = 0
    For I& = 1 To CACONDI%
       VAREFE$ = VAREFCOND$(I&)
       VALORCOND$(I&) = VAREFE$
       IZX$ = "": If Left$(VAREFE$, 1) = "'" Then IZX$ = "'": VAREFE$ = Mid$(VAREFE$, 2)
       DZX$ = "": If Right$(VAREFE$, 1) = "'" Then DZX$ = "'": VAREFE$ = Left$(VAREFE$, Len(VAREFE$) - 1)
10     If Left$(VAREFE$, 1) = "¿" Then
          If Right$(VAREFE$, 1) = "?" Then
             LABPRE$ = Mid$(VAREFE$, 2, Len(VAREFE$) - 2)
             JK& = JK& + 1
             '
             If I& > Label1.Count Then
                NUPIC% = Label1.Count
                Load Label1(NUPIC%)
                Load Combo1(NUPIC%)
                Load Command2(NUPIC%)
             End If
             Label1(I& - 1).Top = PVER + 400 * (JK& - 1)
             Combo1(I& - 1).Top = PVER - 50 + 400 * (JK& - 1)
             Command2(I& - 1).Top = PVER - 50 + 400 * (JK& - 1)
             Label1(I& - 1) = LABPRE$
             If Label1(I& - 1).Width > TAMAXI Then TAMAXI = Label1(I& - 1).Width
             Label1(I& - 1).Visible = True
             Combo1(I& - 1).Visible = True
             Command2(I& - 1).Visible = True
          End If
       End If
    Next I&
    '
    CACONFIL% = JK&
    If CACONFIL% < 1 Then
        Call Command11_Click
        Hide
        Exit Sub
    End If
    '
    For I& = 1 To CACONFIL%
       Label1(I& - 1).Left = 500 + TAMAXI - Label1(I& - 1).Width
       Combo1(I& - 1).Left = Label1(I& - 1).Left + Label1(I& - 1).Width + 250
       Command2(I& - 1).Left = Combo1(I& - 1).Left + Combo1(I& - 1).Width + 150
    Next I&
    Picture1.Left = Command2(CACONDI% - 1).Left + Command2(CACONDI% - 1).Width + 250
    Width = Picture1.Left + 1000
    Height = Command2(CACONDI% - 1).Top + 400 + 500
    Image2.Top = Height - Image2.Height - 400
    If Image2.Top - Command1.Top - Command1.Height > Command11.Height Then Command11.Top = Command1.Top + Command1.Height
End Sub

