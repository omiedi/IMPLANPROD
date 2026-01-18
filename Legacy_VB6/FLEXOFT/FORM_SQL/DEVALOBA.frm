VERSION 5.00
Begin VB.Form DEVALOBA 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Parametro Reemplazable "
   ClientHeight    =   6735
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6975
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6735
   ScaleWidth      =   6975
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame3 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Validaciones y Bloqueos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1950
      Left            =   120
      TabIndex        =   24
      Top             =   4680
      Width           =   6750
      Begin VB.Frame Frame8 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Por Tabla de Valores o Condicion (Textos)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   750
         Left            =   120
         TabIndex        =   35
         Top             =   1080
         Width           =   6495
         Begin VB.TextBox Text5 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   240
            ScrollBars      =   2  'Vertical
            TabIndex        =   36
            Top             =   360
            Width           =   6165
         End
      End
      Begin VB.Frame Frame7 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Lock-Field"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   840
         Left            =   4600
         TabIndex        =   33
         Top             =   240
         Width           =   2015
         Begin VB.ComboBox Combo4 
            Height          =   315
            ItemData        =   "DEVALOBA.frx":0000
            Left            =   100
            List            =   "DEVALOBA.frx":0002
            Sorted          =   -1  'True
            TabIndex        =   34
            Text            =   "Combo4"
            Top             =   360
            Width           =   1785
         End
      End
      Begin VB.Frame Frame6 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Valores (Numeros)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   840
         Left            =   2160
         TabIndex        =   28
         Top             =   240
         Width           =   2360
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
            Height          =   240
            Left            =   600
            ScrollBars      =   2  'Vertical
            TabIndex        =   31
            Top             =   500
            Width           =   1605
         End
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
            Height          =   240
            Left            =   600
            ScrollBars      =   2  'Vertical
            TabIndex        =   29
            Top             =   200
            Width           =   1605
         End
         Begin VB.Label Label7 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "<="
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   -480
            TabIndex        =   32
            Top             =   500
            Width           =   1035
         End
         Begin VB.Label Label6 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   ">="
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   -480
            TabIndex        =   30
            Top             =   200
            Width           =   1035
         End
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Ingreso"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   840
         Left            =   120
         TabIndex        =   25
         Top             =   240
         Width           =   1935
         Begin VB.OptionButton Option2 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Opcional"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   360
            TabIndex        =   27
            Top             =   480
            Width           =   1300
         End
         Begin VB.OptionButton Option1 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Obligatorio"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   360
            TabIndex        =   26
            Top             =   240
            Value           =   -1  'True
            Width           =   1300
         End
      End
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Paste"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6240
      TabIndex        =   23
      Top             =   3600
      Width           =   615
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Copy"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6240
      TabIndex        =   22
      Top             =   3240
      Width           =   615
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Origen del Dato Vacío / Fórmula de Cálculo "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   960
      Left            =   105
      TabIndex        =   11
      Top             =   210
      Width           =   6060
      Begin VB.TextBox Text1 
         Height          =   495
         Left            =   840
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   12
         Top             =   315
         Width           =   5080
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Origen/ Fórmula:"
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
         Left            =   -200
         TabIndex        =   13
         Top             =   360
         Width           =   1035
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
      Height          =   480
      Left            =   6240
      Picture         =   "DEVALOBA.frx":0004
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Termina - Cierra la Aplicación"
      Top             =   4080
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
      Height          =   480
      Left            =   6270
      Picture         =   "DEVALOBA.frx":030E
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Termina - Cierra la Aplicación"
      Top             =   290
      Width           =   615
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Consulta / Destino SQL"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3315
      Left            =   105
      TabIndex        =   0
      Top             =   1260
      Width           =   6060
      Begin VB.Frame Frame2 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Construye Condicion"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1080
         Left            =   120
         TabIndex        =   14
         Top             =   2040
         Width           =   5820
         Begin VB.CommandButton Command3 
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
            Height          =   255
            Left            =   5080
            TabIndex        =   21
            Top             =   240
            Width           =   615
         End
         Begin VB.ComboBox Combo1 
            Height          =   315
            Left            =   3600
            TabIndex        =   18
            Text            =   "Combo1"
            Top             =   600
            Width           =   2115
         End
         Begin VB.ComboBox CMBOPERADOR 
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
            ItemData        =   "DEVALOBA.frx":0458
            Left            =   2400
            List            =   "DEVALOBA.frx":045A
            TabIndex        =   17
            TabStop         =   0   'False
            Text            =   "CMBOPERADOR"
            Top             =   600
            Width           =   1080
         End
         Begin VB.ComboBox CMBCAMPOS 
            Height          =   315
            Left            =   240
            Sorted          =   -1  'True
            TabIndex        =   16
            Text            =   "CMBCAMPOS"
            Top             =   600
            Width           =   2055
         End
         Begin VB.Label Label5 
            BackStyle       =   0  'Transparent
            Caption         =   "Valor:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   3480
            TabIndex        =   20
            Top             =   390
            Width           =   1275
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "Condicion:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   2280
            TabIndex        =   19
            Top             =   390
            Width           =   1275
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Campo:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   100
            TabIndex        =   15
            Top             =   390
            Width           =   1275
         End
      End
      Begin VB.TextBox Text7 
         Height          =   810
         Left            =   480
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   3
         Top             =   1080
         Width           =   5445
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         Left            =   480
         Sorted          =   -1  'True
         TabIndex        =   2
         Text            =   "CMBCAMPOS"
         Top             =   525
         Width           =   2025
      End
      Begin VB.ComboBox Combo3 
         Height          =   315
         Left            =   3300
         Sorted          =   -1  'True
         TabIndex        =   1
         Text            =   "CMBCAMPOS"
         Top             =   525
         Width           =   2625
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   "Condicion:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   120
         TabIndex        =   6
         Top             =   885
         Width           =   1275
      End
      Begin VB.Label Label16 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Campo:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   2205
         TabIndex        =   5
         Top             =   315
         Width           =   1275
      End
      Begin VB.Label Label17 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Tabla:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   105
         TabIndex        =   4
         Top             =   315
         Width           =   615
      End
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "N-O:"
      Height          =   195
      Left            =   6345
      TabIndex        =   10
      Top             =   2085
      Width           =   330
   End
   Begin VB.Label Label14 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Height          =   285
      Left            =   6270
      TabIndex        =   9
      Top             =   2295
      Width           =   615
   End
End
Attribute VB_Name = "DEVALOBA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Combo2_Change()
    Call Combo2_Click
End Sub

Private Sub Command1_Click()
    Command1.Enabled = False
    '
    'VALIDA SI EL OBJETO DE LOCK FIELD ES EXISTENTE
    NomObj$ = TRIM$(Combo4.TEXT)
    If NomObj$ <> "" Then
        If VALIOBJ%(NomObj$) < 1 Then
           Call COMUNI("Nombre de Objeto (" & NomObj$ & ") Inexistente")
           Combo4.TEXT = NomObj$
           GoTo 99
        End If
    End If
    '
    IPAR% = XVALO(Label14)
    OBLI% = 1: If Option2.Value = True Then OBLI% = 0
    If IPAR% >= 0 Then
      If IPAR% <= CANOBJ% Then
        FOREPLA$(IPAR%) = Text1
        TASEL$(IPAR%) = TRIM$(Combo2.TEXT)
        CASEL$(IPAR%) = TRIM$(Combo3.TEXT)
        CONDSEL$(IPAR%) = TRIM$(UCase$(Text7))
        OBLIGA%(IPAR%) = OBLI%
        LOCKFI$(IPAR%) = TRIM$(Combo4.TEXT)
        VAMINI$(IPAR%) = Text2
        VAMAXI$(IPAR%) = Text3
        TABVALI$(IPAR%) = Text5
        Command1.Enabled = True
        Label14 = ""
        Hide
      End If
    End If
    '
99  Command1.Enabled = True
End Sub
Function VALIOBJ%(NOMBRE$)
    '
    ' VALIDA QUE EL NOMBRE PASADO COMO PARAMETRO SEA UN OBJETO EXISTENTE
    '
    VALIOBJ% = 0
    'RECORRO LOS OBJETOS Y COMPARO NOMBRE
    For I% = 1 To CANOBJ%
     NOM$ = ONAME$(I%)
     If UCase$(NOM$) = UCase$(NOMBRE$) Then
       VALIOBJ% = 1
       Exit For
     End If
    Next I%
    '
End Function
Private Sub Command2_Click()
    Unload Me
End Sub

Private Sub Command3_Click()
   '
   If TRIM$(CMBCAMPOS.TEXT) <> "" Then
     If TRIM$(CMBOPERADOR) <> "" Then
       If TRIM$(Combo1.TEXT) <> "" Then
         TTXX$ = CMBCAMPOS.TEXT + " " + CMBOPERADOR.TEXT + " " + Combo1.TEXT
         If Text7 <> "" Then Text7 = Text7 + " AND "
         Text7 = Text7 + TTXX$
       End If
     End If
   End If
   '
End Sub

Private Sub Command4_Click()
    RTEXT1 = Text1
    RCOMBO2 = Combo2.TEXT
    RCOMBO3 = Combo3.TEXT
    RTEXT7 = Text7
End Sub

Private Sub Command5_Click()
    Text1 = RTEXT1
    Combo2.TEXT = RCOMBO2
    Combo3.TEXT = RCOMBO3
    Text7 = RTEXT7
End Sub

Private Sub Form_Activate()
    Combo1.Clear
    Combo4.Clear
    CMBCAMPOS.TEXT = ""
    For KKI% = 0 To 3
       Combo1.AddItem "{COMPAR-" & CStr(KKI%) + "}"
    Next KKI%
    '
    For KKI% = 1 To CANOBJ%
       If TIOBJ$(KKI%) = "PARAM" Or TIOBJ$(KKI%) = "TEXT" Then
          Combo1.AddItem "{" + ONAME$(KKI%) + "}"
          If Left$(UCase$(FOREPLA$(KKI%)), 5) = "@PASS" Then
             If KKI% <> XVALO(Label14) Then
                Combo4.AddItem ONAME$(KKI%)
             End If
          End If
       End If
    Next KKI%
    '
    Call Label14_Change        ' PARA REFRESCAR
    '
End Sub

Private Sub Form_Load()
    '
    Call ABRECONEXION
    Call MUESTRATABLAS
    '
    CMBOPERADOR.Clear
    CMBOPERADOR.AddItem " =  "
    CMBOPERADOR.AddItem " <  "
    CMBOPERADOR.AddItem " <= "
    CMBOPERADOR.AddItem " >  "
    CMBOPERADOR.AddItem " >= "
    CMBOPERADOR.AddItem " <> "
    CMBOPERADOR.AddItem "LIKE"
    '
End Sub

Sub MUESTRATABLAS()
    Combo2.Clear: Combo3.Clear
    Dim AUX As ADODB.Recordset
    Set AUX = New ADODB.Recordset
22  SQLX$ = "Select name from sysobjects WHERE LEFT(NAME,3) <> 'dt_' and LEFT(NAME,3) <> 'pk_' and LEFT(NAME,3) <> 'sys'"
    Set AUX = FLEXCONN.Execute(SQLX$)
    Do While Not AUX.EOF
       If InStr(UCase$(AUX!Name), "DTPROPERT") > 0 Then GoTo 29
       Combo2.AddItem UCase$(AUX!Name)
29  AUX.MoveNext
    Loop
End Sub
'
Private Sub Combo2_Click()
    If TRIM$(Combo2.TEXT) <> "" Then
       Combo3.Clear: CMBCAMPOS.Clear
       Dim AUX As ADODB.Recordset
       Set AUX = New ADODB.Recordset
       On Error Resume Next
       Set AUX = FLEXCONN.Execute("select * from " & Combo2.TEXT)
       If Err Then
          On Error GoTo 0
          Call MENSERR(24, "Referencia a Tabla SQL '" + TRIM$(Combo2.TEXT) + "' Inexistente.")
          Exit Sub
       End If
       On Error GoTo 0
       For I& = 1 To AUX.Fields.Count
           Combo3.AddItem AUX.Fields(I& - 1).Name
           CMBCAMPOS.AddItem AUX.Fields(I& - 1).Name
       Next I&
    End If
End Sub
'
Private Sub Label14_Change()
    Option1.Value = True
    IPAR% = XVALO(Label14)
    If IPAR% >= 0 Then
      If IPAR% <= CANOBJ% Then
        Text1 = FOREPLA$(IPAR%)
        Combo2.TEXT = TASEL$(IPAR%)
        Combo3.TEXT = CASEL$(IPAR%)
        Combo4.TEXT = LOCKFI$(IPAR%)
        Text2 = VAMINI$(IPAR%)
        Text3 = VAMAXI$(IPAR%)
        Text5 = TABVALI$(IPAR%)
        Text7 = CONDSEL$(IPAR%)
        If OBLIGA%(IPAR%) < 1 Then Option2.Value = True
      End If
    End If
End Sub

