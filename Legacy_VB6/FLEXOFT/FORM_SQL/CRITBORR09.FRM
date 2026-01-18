VERSION 5.00
Begin VB.Form CRITBORR09 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Criterios de Borrado"
   ClientHeight    =   4725
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7770
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4725
   ScaleWidth      =   7770
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00400000&
      Height          =   1905
      Left            =   6825
      ScaleHeight     =   1845
      ScaleWidth      =   1440
      TabIndex        =   13
      Top             =   0
      Width           =   1500
      Begin VB.PictureBox MARCOBARRA 
         BackColor       =   &H000000FF&
         Height          =   135
         Left            =   120
         ScaleHeight     =   75
         ScaleWidth      =   585
         TabIndex        =   15
         Top             =   1155
         Width           =   650
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H00FF0000&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   0
            TabIndex        =   16
            Top             =   0
            Width           =   12000
         End
      End
      Begin VB.CommandButton Command3 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   105
         Picture         =   "CRITBORR09.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Cerrar y Salir"
         Top             =   315
         Width           =   700
      End
      Begin VB.Image Image4 
         Height          =   390
         Left            =   -315
         Picture         =   "CRITBORR09.frx":014A
         Top             =   1460
         Width           =   1515
      End
   End
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "Quitar Criterio"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   6825
      TabIndex        =   11
      Top             =   2100
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Agregar Criterio"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   6825
      TabIndex        =   10
      Top             =   2730
      Width           =   855
   End
   Begin VB.Frame Frame1 
      Caption         =   "Definicion de Criterio"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1170
      Left            =   105
      TabIndex        =   1
      Top             =   3465
      Width           =   7575
      Begin VB.ComboBox Combo1 
         Height          =   315
         Index           =   1
         Left            =   5880
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   735
         Width           =   1590
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   1
         Left            =   1260
         TabIndex        =   7
         Top             =   735
         Width           =   4215
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Index           =   0
         Left            =   5880
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   360
         Width           =   1590
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Index           =   0
         Left            =   1260
         TabIndex        =   3
         Top             =   360
         Width           =   4215
      End
      Begin VB.Label Label3 
         Caption         =   "Label3"
         Height          =   225
         Left            =   6510
         TabIndex        =   12
         Top             =   105
         Visible         =   0   'False
         Width           =   750
      End
      Begin VB.Label Label2 
         Caption         =   " en"
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
         Index           =   1
         Left            =   5460
         TabIndex        =   8
         Top             =   795
         Width           =   435
      End
      Begin VB.Label Label1 
         Caption         =   "Y tambien"
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
         Index           =   1
         Left            =   300
         TabIndex        =   6
         Top             =   795
         Width           =   1065
      End
      Begin VB.Label Label2 
         Caption         =   " en"
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
         Index           =   0
         Left            =   5460
         TabIndex        =   4
         Top             =   420
         Width           =   435
      End
      Begin VB.Label Label1 
         Caption         =   "Si Aparece"
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
         Index           =   0
         Left            =   210
         TabIndex        =   2
         Top             =   420
         Width           =   1065
      End
   End
   Begin VB.ListBox List1 
      Height          =   2985
      Left            =   105
      TabIndex        =   0
      Top             =   315
      Width           =   6630
   End
End
Attribute VB_Name = "CRITBORR09"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    If TRIM$(Text1(0)) <> "" Then
       If TRIM$(Combo1(0).TEXT) <> "" Then
          NCRIT% = XVALO(Label3)
          If NCRIT% < 1 Then
             NCRIT% = 1 + XVALO(VALOBADA("CRITBORR", "MAX(NCRITBO)", "NCRITBO > 0", "NOMESS"))
             Call CreaRegistro("CRITBORR", "NCRITBO", NCRIT%)
          End If
          '
          CONDI$ = "NCRITBO = " & NCRIT%
          Call ACTUREGISTRO("CRITBORR", "STRING1", CONDI$, Text1(0), RAFE&, "NOMESS")
          Call ACTUREGISTRO("CRITBORR", "UBICA1", CONDI$, Combo1(0), RAFE&, "NOMESS")
          Call ACTUREGISTRO("CRITBORR", "STRING2", CONDI$, Text1(1), RAFE&, "NOMESS")
          Call ACTUREGISTRO("CRITBORR", "UBICA2", CONDI$, Combo1(1), RAFE&, "NOMESS")
          Label3 = ""
          Call CARGALISTA
          '
          Text1(0) = "": Text1(1) = ""
          Combo1(0).Clear:  Combo1(1).Clear
          For KKI% = 1 To 2
             Combo1(KKI% - 1).AddItem "Asunto"
             Combo1(KKI% - 1).AddItem "Remitente"
             Combo1(KKI% - 1).AddItem "Texto"
          Next KKI%
          Hide
          '
       End If
    End If
End Sub

Private Sub Command2_Click()
    '
    If DERACCE%("BORRCRIT", "Borrar Criterios de Borrado") >= 2 Then
       NCRIT% = XVALO(Label3)
       If NCRIT% > 0 Then
          CONDI$ = "NCRITBO = " & NCRIT%
          Call BORRAREGISTROS("CRITBORR", CONDI$, RAFE&, "NOMESS")
          Label3 = ""
          Call CARGALISTA
          '
          Text1(0) = "": Text1(1) = ""
          Combo1(0).Clear:  Combo1(1).Clear
          For KKI% = 1 To 2
             Combo1(KKI% - 1).AddItem "Asunto"
             Combo1(KKI% - 1).AddItem "Remitente"
             Combo1(KKI% - 1).AddItem "Texto"
          Next KKI%
          '
       End If
    End If
    '
End Sub

Private Sub Command3_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    '
    For KKI% = 1 To 2
       Combo1(KKI% - 1).AddItem "Asunto"
       Combo1(KKI% - 1).AddItem "Remitente"
       Combo1(KKI% - 1).AddItem "Texto"
    Next KKI%
    '
    Call CREACAMPO("", "CRITBORR", "NCRITBO", 3, 2, 1)
    Call CREACAMPO("", "CRITBORR", "STRING1", 10, 128)
    Call CREACAMPO("", "CRITBORR", "UBICA1", 10, 32)
    Call CREACAMPO("", "CRITBORR", "STRING2", 10, 128)
    Call CREACAMPO("", "CRITBORR", "UBICA2", 10, 32)
    '
    Call CARGALISTA
    '
End Sub

Sub CARGALISTA()
    List1.Clear
    Dim CRIBO As ADODB.Recordset
    Set CRIBO = New ADODB.Recordset
    Set CRIBO = FLEXCONN.Execute("SELECT * FROM CRITBORR WHERE NCRITBO > 0 ORDER BY NCRITBO")
    With CRIBO
       Do While Not .EOF
         TTXX$ = TRIM$(SAFETEXT$(!STRING1)) + " en " + TRIM$(SAFETEXT$(!UBICA1))
         If TRIM$(SAFETEXT$(!STRING2)) <> "" Then
            If TRIM$(SAFETEXT$(!UBICA2)) <> "" Then
               TTXX$ = TTXX$ + " y tambien " + TRIM$(SAFETEXT$(!STRING2)) + " en " + TRIM$(SAFETEXT$(!UBICA2))
            End If
         End If
         List1.AddItem FORMATNUM$(XVALO(!NCRITBO), "F4") + "  " + TTXX$
       .MoveNext
       Loop
    End With
End Sub

Private Sub List1_DblClick()
    '
    Label3 = TRIM$(Left$(List1.TEXT, 4))
    NCRIT% = XVALO(Label3)
    CONDI$ = "NCRITBO = " & NCRIT%
    Text1(0) = VALOBADA("CRITBORR", "STRING1", CONDI$, "NOMESS")
    Text1(1) = VALOBADA("CRITBORR", "STRING2", CONDI$, "NOMESS")
    On Error Resume Next
       Combo1(0) = VALOBADA("CRITBORR", "UBICA1", CONDI$, "NOMESS")
       Combo1(1) = VALOBADA("CRITBORR", "UBICA2", CONDI$, "NOMESS")
    On Error GoTo 0
    '
End Sub
