VERSION 5.00
Begin VB.Form FORMMENSA 
   BackColor       =   &H00C0E0FF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "A.B.M. de Mensajes"
   ClientHeight    =   8655
   ClientLeft      =   165
   ClientTop       =   5535
   ClientWidth     =   9315
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8655
   ScaleWidth      =   9315
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame5 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Mensaje Viejo"
      Height          =   1905
      Left            =   105
      TabIndex        =   17
      Top             =   6720
      Width           =   8070
      Begin VB.TextBox Text4 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1440
         Left            =   210
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   18
         Top             =   250
         Width           =   7695
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Ingles"
      Height          =   1905
      Left            =   105
      TabIndex        =   11
      Top             =   4725
      Width           =   8070
      Begin VB.TextBox Text3 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1440
         Left            =   210
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   12
         Top             =   250
         Width           =   7695
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Portugués"
      Height          =   1905
      Left            =   105
      TabIndex        =   9
      Top             =   2730
      Width           =   8070
      Begin VB.TextBox Text2 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1440
         Left            =   210
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   10
         Top             =   250
         Width           =   7695
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Español"
      Height          =   1905
      Left            =   105
      TabIndex        =   5
      Top             =   735
      Width           =   8070
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1440
         Left            =   210
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   6
         Top             =   250
         Width           =   7695
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00062A70&
      Height          =   8670
      Left            =   8400
      ScaleHeight     =   8610
      ScaleWidth      =   1215
      TabIndex        =   2
      Top             =   -20
      Width           =   1275
      Begin VB.CommandButton Command4 
         Caption         =   "Probar"
         Height          =   640
         Left            =   105
         Style           =   1  'Graphical
         TabIndex        =   16
         ToolTipText     =   "Guardar Mensaje"
         Top             =   2730
         Width           =   640
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Save"
         Height          =   640
         Left            =   105
         Picture         =   "FORMMENSA.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "Guardar Mensaje"
         Top             =   1890
         Width           =   640
      End
      Begin VB.CommandButton Command1 
         Caption         =   "New"
         Height          =   640
         Left            =   105
         Picture         =   "FORMMENSA.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Nuevo Mensaje"
         Top             =   1050
         Width           =   640
      End
      Begin VB.CommandButton Command3 
         Height          =   640
         Left            =   105
         Picture         =   "FORMMENSA.frx":0454
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Cierra y Abandona la Aplicación"
         Top             =   210
         Width           =   640
      End
      Begin VB.PictureBox MARCOBARRA 
         BackColor       =   &H000000FF&
         Height          =   135
         Left            =   105
         ScaleHeight     =   75
         ScaleWidth      =   585
         TabIndex        =   3
         Top             =   7875
         Width           =   640
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H00FF0000&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   0
            TabIndex        =   4
            Top             =   105
            Width           =   12000
         End
      End
      Begin VB.Image Image2 
         Height          =   390
         Left            =   -315
         Picture         =   "FORMMENSA.frx":059E
         Top             =   8190
         Width           =   1515
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0E0FF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Left            =   5040
      TabIndex        =   0
      Top             =   105
      Width           =   3135
      Begin VB.CommandButton Command9 
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
         Height          =   360
         Left            =   2730
         Picture         =   "FORMMENSA.frx":24C0
         TabIndex        =   8
         Top             =   210
         Width           =   200
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00EEEEDD&
         BackStyle       =   0  'Transparent
         Caption         =   "Código:"
         Height          =   225
         Left            =   630
         TabIndex        =   13
         Top             =   210
         Width           =   900
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
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
         Left            =   1575
         TabIndex        =   1
         Top             =   210
         Width           =   1170
      End
   End
End
Attribute VB_Name = "FORMMENSA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

   Command1.Enabled = False
   
   NUMPRO& = 1 + XVALO(VALOMENSA("MENSAJES", "MAX(NUMMENSA)", "NUMMENSA > 0 ", "NOMESS"))
   Label11 = NUMPRO&
   
   Command1.Enabled = True
End Sub

Private Sub Command2_Click()
    '
    Command2.Enabled = False
    Call ABREMENSAFLEX
    '
    'TOMO EL NUMERO DEL LABEL
    NUMEN& = XVALO(Label11)
    '
    'VALIDACION QUE SE HAYA SELECCIONADO UN CODIGO
    If NUMEN& < 1 Then
       Call MENSERR(24, "Código de Mensaje No Valido")
       GoTo 99
    End If
    '
    'GRABACION , SI NO EXISTE LO GENERA SI NO LO EDITA
    SQLX$ = "SELECT * FROM MENSAJES "
    SQLX$ = SQLX$ + " WHERE NUMMENSA = " & NUMEN&
    '
    Set MENSATMP = BaDaMensaFlex.OpenRecordset(SQLX$, dbOpenDynaset)
    With MENSATMP
          If (.BOF And .EOF) Then
            .AddNew
          Else
            .Edit
          End If
            !NUMMENSA = NUMEN&
            !MENSAESP = REPLACAR(Text1, "\", CRLF$)
            !MENSAPOR = REPLACAR(Text2, "\", CRLF$)
            !MENSAING = REPLACAR(Text3, "\", CRLF$)
            !MensaViejo = REPLACAR(Text4, "\", CRLF$)
            .Update
    End With
    '
    MENSATMP.Close
    Set MENSATMP = Nothing
    
99  Command2.Enabled = True
End Sub

Private Sub Command3_Click()
   
    Call CIERRARCH("*.*")
    Call FINAL("")

End Sub

Private Sub Command4_Click()
    'BOTON PARA REALIZAR PRUEBAS
    '
    TEXTO1$ = InputBox$("Ingrese TEXTO", "TEXTO")

    NUM& = XVALO(Label11)
    
    Call MENSERR(24, "#" & NUM& & Chr(124) & "", TEXTO1$)
    '
End Sub

Private Sub Command9_Click()
    'PRESENTA INDICE DE SELECCION DE MENSAJES
    
    Command9.Enabled = False
    Call ABREMENSAFLEX
    JJ& = 0
    '
    'RECORRE LA TABLA MENSAJES Y CARGA LOS MENSAJES EN EL ARCHIVO
    SQLX$ = "SELECT * FROM MENSAJES "
    SQLX$ = SQLX$ + " WHERE NUMMENSA > 0"
    SQLX$ = SQLX$ + " ORDER BY NUMMENSA"
    '
    Set MENSATMP = BaDaMensaFlex.OpenRecordset(SQLX$, dbOpenSnapshot)
    With MENSATMP
      Do While Not .EOF
       COCU$ = Str(XVALO(!NUMMENSA))
       DESCRI$ = SAFETEXT$(!MENSAESP)
       TTXX$ = AJUSTI$(COCU$, 16) + DESCRI$
       JJ& = JJ& + 1
       Call GRAREG("!INDIASIE", TTXX$, JJ&)
       .MoveNext
      Loop
    End With
    
    Call SETULTREG("!INDIASIE", JJ&)
    Call FRESHECHO("!INDIASIE")

    MENSATMP.Close
    Set MENSATMP = Nothing
    '
    'PRESENTO SELECHO Y SI ELIGE LO ASIGNA AL LABEL
    Call SELECHO("!INDIASIE/Indice de Mensajes")
    CODI$ = VALACT1$("!INDIASIE")
    If CODI$ = "" Then GoTo 99
    Label11 = CODI$
    '
99  Command9.Enabled = True
    '
End Sub

Private Sub Form_Load()

    Call CENTRAFORM(Me)
    '
    Call VERJOBID(OKEY%)
    If OKEY% < 1 Then Call FINAL("")
    '
End Sub

Private Sub Label11_Change()
   
   NUMEN& = XVALO(Label11)
   If NUMEN& < 1 Then Exit Sub
   '
   'MUESTRA LA INFORMACION EN LOS CASILLEROS CORRESPONDIENTES A CADA IDIOMA
   Text1 = REPLACAR(VALOMENSA("MENSAJES", "MENSAESP", "NUMMENSA =" & NUMEN, "NOMESS"), CRLF, "\")
   Text2 = REPLACAR(VALOMENSA("MENSAJES", "MENSAPOR", "NUMMENSA =" & NUMEN, "NOMESS"), CRLF, "\")
   Text3 = REPLACAR(VALOMENSA("MENSAJES", "MENSAING", "NUMMENSA =" & NUMEN, "NOMESS"), CRLF, "\")
   Text4 = REPLACAR(VALOMENSA("MENSAJES", "MensaViejo", "NUMMENSA =" & NUMEN, "NOMESS"), CRLF, "\")
   '
End Sub

