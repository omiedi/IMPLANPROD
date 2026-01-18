VERSION 5.00
Begin VB.Form DOCPREFE 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0FFC0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Vinculación Documentos Preferidos"
   ClientHeight    =   4650
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   9660
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4650
   ScaleWidth      =   9660
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.Frame Frame4 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Archivo de Conexión"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4320
      Left            =   105
      TabIndex        =   0
      Top             =   210
      Width           =   9465
      Begin VB.CommandButton Command1 
         Height          =   360
         Left            =   8505
         Picture         =   "DOCPREFE.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Agregar a Lista de preferidos"
         Top             =   3750
         Width           =   750
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1050
         TabIndex        =   11
         Top             =   3780
         Width           =   7365
      End
      Begin VB.CommandButton Command3 
         Caption         =   "View"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1000
         Left            =   8505
         Picture         =   "DOCPREFE.frx":014A
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Visualizar Documento Activo"
         Top             =   2550
         Width           =   750
      End
      Begin VB.FileListBox File1 
         Height          =   1845
         Left            =   3780
         TabIndex        =   6
         Top             =   420
         Width           =   5475
      End
      Begin VB.DirListBox Dir1 
         Height          =   1440
         Left            =   315
         TabIndex        =   2
         Top             =   840
         Width           =   3375
      End
      Begin VB.DriveListBox Drive1 
         Height          =   315
         Left            =   1470
         TabIndex        =   1
         Top             =   420
         Width           =   2220
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Conectar Como:"
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
         Left            =   105
         TabIndex        =   10
         Top             =   3675
         Width           =   810
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Unidad:"
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
         Top             =   470
         Width           =   1065
      End
      Begin VB.Label Label1 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
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
         Left            =   1050
         TabIndex        =   7
         Top             =   2580
         Width           =   7365
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Archivo:"
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
         TabIndex        =   5
         Top             =   2625
         Width           =   810
      End
      Begin VB.Label Label11 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   1050
         TabIndex        =   4
         Top             =   2940
         Width           =   7365
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Carpeta:"
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
         TabIndex        =   3
         Top             =   2925
         Width           =   810
      End
   End
End
Attribute VB_Name = "DOCPREFE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub command1_Click()
   XX$ = REGBLAN$("DOCPREFE")
   Call REPLA(XX$, AJUSTI$(Text1, 64), 1, 64)
   Call REPLA(XX$, TRIM$(Label1), 65, 128)
   Call REPLA(XX$, TRIM$(Mid$(Label11, 1)), 193, 128)
   Call REGAPP("DOCPREFE", XX$)
   Call CIERRARCH("DOCPREFE")
   '
   MAILMESS.List2.Clear
   For U& = 1 To ULTREG&("DOCPREFE")
     XX$ = REGLEIDO$("DOCPREFE", U&)
     YY$ = Space$(320)
     Call REPLA(YY$, XX$, 1, 320)
     MAILMESS.List2.AddItem YY$
   Next U&
   MAILMESS.List2.Refresh
   DoEvents
   '
End Sub

Private Sub Command3_Click()
    '
    Command3.Enabled = False
    Call MUESTRADOC3
    Command3.Enabled = True
    '
End Sub

Private Sub Dir1_Change()
   Label11.Caption = Dir1.Path
   File1.Path = Dir1.Path
   Label1 = ""
End Sub

Private Sub Drive1_Change()
   On Error GoTo 20
   Dir1.Path = Drive1.Drive
   GoTo 99
   '
20 Resume 21
21 On Error GoTo 0
   Call MENSERR(24, "Unidad no Disponible.")
   Drive1.Drive = Left$(CurDir$, 2)
   '
99 On Error GoTo 0
End Sub

Private Sub File1_Click()
  Label1 = File1.filename
End Sub

Private Sub File1_DblClick()
   Call MUESTRADOC3
End Sub

Private Sub Form_Load()
    '
    EPAC$ = EMPREAC$
    If EPAC$ <> "" Then
      Caption = Caption + " - " + EMPREAC$
    End If
    '
    Label11.Caption = Dir1.Path
    File1.Path = Dir1.Path
    '
End Sub

Private Sub Form_Resize()
   If WindowState = 1 Then
     Call CIERRARCH("*.*")
     Call BAJALDISCO
   End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call CIERRARCH("*.*")
   Call BAJALDISCO
End Sub

Sub MUESTRADOC3()
    '
    On Error Resume Next
    WORDDOCS.OLE1.SourceDoc = Label11 + "\" + Label1
    TIDOCU$ = "WORD"
    If UCase$(Right$(Label1, 4)) = ".EML" Then
        TIDOCU$ = "MAIL"
      ElseIf UCase$(Right$(Label1, 4)) = ".XLS" Then
        TIDOCU$ = "EXCL"
    End If
    WORDDOCS.OLE1.Action = 1
    WORDDOCS.OLE1.Action = 6
    WORDDOCS.VScroll1.Min = 0
    WORDDOCS.VScroll1.Value = 0
    ABARRA = WORDDOCS.OLE1.Height - WORDDOCS.Picture1.Height
    If ABARRA < 1 Then ABARRA = 1
    WORDDOCS.VScroll1.Max = ABARRA
    WORDDOCS.VScroll1.SmallChange = ABARRA / 100
    WORDDOCS.VScroll1.LargeChange = ABARRA / 10
    WORDDOCS.OLE1.Top = 0
    WORDDOCS.OLE1.DoVerb (-1)
    If TIDOCU$ = "WORD" Then
        SendKeys "%{F10}", True  ' esto para documentos en word
      ElseIf TIDOCU$ = "MAIL" Then
        SendKeys Chr$(13) + "%{F10}", True ' esto para mails
      ElseIf TIDOCU$ = "EXCL" Then
        SendKeys Chr$(13) & "%V" & (DOWN8) & "P", True ' ESTO PARA EXCEL
    End If
    WORDDOCS.Show 1
    On Error GoTo 0
    '
End Sub

