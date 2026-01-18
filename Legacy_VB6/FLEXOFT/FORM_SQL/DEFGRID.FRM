VERSION 5.00
Begin VB.Form DEFGRID 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Objeto Tipo GRID"
   ClientHeight    =   3735
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6975
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3735
   ScaleWidth      =   6975
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Configuración de Consulta"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3270
      Left            =   105
      TabIndex        =   4
      Top             =   315
      Width           =   6060
      Begin VB.Frame Frame2 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Filtros Dinámicos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1800
         Left            =   210
         TabIndex        =   8
         Top             =   1260
         Width           =   5640
         Begin VB.TextBox Text2 
            Height          =   285
            Index           =   3
            Left            =   1050
            ScrollBars      =   2  'Vertical
            TabIndex        =   14
            Top             =   1365
            Width           =   4455
         End
         Begin VB.TextBox Text2 
            Height          =   285
            Index           =   2
            Left            =   1050
            ScrollBars      =   2  'Vertical
            TabIndex        =   13
            Top             =   1050
            Width           =   4455
         End
         Begin VB.TextBox Text2 
            Height          =   285
            Index           =   1
            Left            =   1050
            ScrollBars      =   2  'Vertical
            TabIndex        =   11
            Top             =   735
            Width           =   4455
         End
         Begin VB.TextBox Text2 
            Height          =   285
            Index           =   0
            Left            =   1050
            ScrollBars      =   2  'Vertical
            TabIndex        =   9
            Top             =   420
            Width           =   4455
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "-3:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   450
            Index           =   3
            Left            =   -2100
            TabIndex        =   16
            Top             =   1425
            Width           =   3135
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "-2:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   450
            Index           =   2
            Left            =   -2100
            TabIndex        =   15
            Top             =   1110
            Width           =   3135
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "-1:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   450
            Index           =   1
            Left            =   -2100
            TabIndex        =   12
            Top             =   795
            Width           =   3135
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Filtro-0:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   450
            Index           =   0
            Left            =   -2100
            TabIndex        =   10
            Top             =   480
            Width           =   3135
         End
      End
      Begin VB.CommandButton Command20 
         Caption         =   "."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   5130
         TabIndex        =   6
         Top             =   420
         Width           =   175
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   5460
         Picture         =   "DEFGRID.frx":0000
         Top             =   420
         Width           =   480
      End
      Begin VB.Label Label1 
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
         Height          =   285
         Left            =   2955
         TabIndex        =   17
         Top             =   420
         Width           =   2190
      End
      Begin VB.Label Label2 
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
         Height          =   285
         Left            =   210
         TabIndex        =   7
         Top             =   795
         Width           =   5130
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Codigo de Consulta:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   735
         TabIndex        =   5
         Top             =   480
         Width           =   2190
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
      Height          =   690
      Left            =   6270
      Picture         =   "DEFGRID.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Termina - Cierra la Aplicación"
      Top             =   2900
      Width           =   615
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
      Height          =   585
      Left            =   6270
      Picture         =   "DEFGRID.frx":0614
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Termina - Cierra la Aplicación"
      Top             =   290
      Width           =   615
   End
   Begin VB.Label Label4 
      Height          =   330
      Left            =   210
      TabIndex        =   18
      Top             =   3885
      Width           =   6735
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "N-O:"
      Height          =   195
      Left            =   6405
      TabIndex        =   3
      Top             =   1665
      Width           =   330
   End
   Begin VB.Label Label14 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   6270
      TabIndex        =   2
      Top             =   1875
      Width           =   615
   End
End
Attribute VB_Name = "DEFGRID"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    FILTXI$ = ""
    For JJ& = 0 To 3
       If TRIM$(Text2(JJ&)) <> "" Then
          If FILTXI$ <> "" Then FILTXI$ = FILTXI$ + ";"
          FILTXI$ = FILTXI$ + TRIM$(Text2(JJ&))
       End If
    Next JJ&
    If TRIM$(FILTXI$) <> "" Then
       FILTXI$ = "(" + FILTXI$ + ")"
    End If
    Label4 = TRIM$(Label1) + FILTXI$
    Hide
End Sub

Private Sub Command2_Click()
    Label4 = ""
    Unload Me
End Sub

Private Sub Command20_Click()
    RSEL$ = REGSEL$("!INDIASIE/Consultas y Reportes Preconfigurados")
    If Len(RSEL$) > 4 Then
       Label1 = TRIM$(Left$(RSEL$, 16))
       Label2 = TRIM$(Mid$(RSEL$, 16, 50))
    End If
End Sub

Private Sub Form_Load()
    '
    Call ABRETACON
    JJ& = 0
    '
    STRSQL = "SELECT COD_QUERY, DES_QUERY FROM SQL_QUERIES ORDER BY COD_QUERY Asc "
    On Error Resume Next
    Set MIRS = BaDaConsul.OpenRecordset(STRSQL, dbOpenSnapshot)
    MIRS.MoveFirst
    If Err Then
       On Error GoTo 0
       Call MENSERR(24, "No Hay Consultas Preconfiguradas\para esta Instalacion.")
       MIRS.Close
       Set MIRS = Nothing
       Unload Me
       Exit Sub
    End If
    On Error GoTo 0
    Do While Not MIRS.EOF
       COCU$ = MIRS!COD_QUERY
       DECU$ = MIRS!DES_QUERY
       TTXX$ = AJUSTI$(COCU$, 16) + DECU$
       JJ& = JJ& + 1
       Call GRAREG("!INDIASIE", TTXX$, JJ&)
    MIRS.MoveNext
    Loop
    MIRS.Close: Set MIRS = Nothing
    Call SETULTREG("!INDIASIE", JJ&)
    Call FRESHECHO("!INDIASIE")
    '
End Sub

Private Sub Image1_DblClick()
    FGENREP07.Label1 = TRIM$(Label1.Caption)
    FGENREP07.Show 1
    Call Form_Load
End Sub

Private Sub Label14_Change()
    IOBJ% = XVALO(Label14)
    Label4 = TEXTO(IOBJ%)
End Sub

Private Sub Label4_Change()
    Label1 = "": Label2 = ""
    For U& = 0 To 3: Text2(U&) = "": Next U&
    '
    APLIXI$ = Label4
    FILTXI$ = ""
    If APLIXI$ <> "" Then
      PPXX% = InStr(APLIXI$, "(")
      If PPXX% > 0 Then
         PPYY% = InStr(PPXX%, APLIXI$, ")")
         If PPYY% <= PPXX% Then PPYY% = 1 + Len(APLIXI$)
         FILTXI$ = Mid$(APLIXI$, PPXX% + 1, PPYY% - PPXX% - 1)
         APLIXI$ = Left$(APLIXI$, PPXX% - 1) + Mid$(APLIXI$, PPYY% + 1)
      End If
    End If
    '
    Label1 = APLIXI$
    Label2 = ECOARCH$("!INDIASIE", APLIXI$)
    '
    JJ& = 0
20  PPXX% = InStr(FILTXI$, ";")
    If PPXX% < 1 Then PPXX% = 1 + Len(FILTXI$)
    If PPXX% > 1 Then
       Text2(JJ&) = Left$(FILTXI$, PPXX% - 1)
       FILTXI$ = Mid$(FILTXI$, PPXX% + 1)
       JJ& = JJ& + 1
       If JJ& <= 3 Then GoTo 20
    End If
    '
End Sub

