VERSION 5.00
Begin VB.Form FORMUDEP 
   BackColor       =   &H00C0E0FF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Depuracion del Programa de Entregas"
   ClientHeight    =   2670
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   4305
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2670
   ScaleWidth      =   4305
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
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
      Height          =   620
      Left            =   3360
      Picture         =   "FORMUDEP.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Grabar - Guardar Modificaciones"
      Top             =   855
      Width           =   750
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
      Height          =   620
      Left            =   3360
      Picture         =   "FORMUDEP.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Cancelar la Aplicación"
      Top             =   210
      Width           =   750
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
      Height          =   330
      Left            =   2205
      TabIndex        =   3
      Top             =   1155
      Width           =   175
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
      Left            =   1155
      TabIndex        =   2
      Top             =   1155
      Width           =   1065
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Fecha:"
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
      Left            =   210
      TabIndex        =   1
      Top             =   1155
      Width           =   855
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Mediante  este  utilitario  se anulan y eliminan automaticamente las entregas programadas anteriores a la fecha elegida."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   105
      TabIndex        =   0
      Top             =   1785
      Width           =   4635
   End
End
Attribute VB_Name = "FORMUDEP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub BOTEXIT_Click()
    Hide
End Sub

Private Sub BOTREC_Click()
    '
    HOO% = HOY(HO$)
    FLIM% = FECHANUM(Text1.Text)
    If FLIM% < 1 Then
        Call MENSERR(24, "Fecha no Válida")
        Text1.Text = ""
        Text1.SetFocus
        Exit Sub
      ElseIf FLIM% > HOO% Then
        Call MENSERR(24, "Fecha no Válida")
        Text1.Text = ""
        Text1.SetFocus
        Exit Sub
      Else
        FIN& = ULTREG&("PROGENT")
        FIN1& = FIN&: If FIN1& < 1 Then FIN1& = 1
        '
        'Call BLIMESS("Grabando Programa de Entregas")
        J& = 0
        For I& = 1 To FIN&
           Call TRACE(10, I&, FIN&)
           X$ = REGLEIDO$("PROGENT", I&)
           If CVI(Mid$(X$, 3, 2)) >= FLIM% Then
              J& = J& + 1
              Call GRAREG("PROGENT", X$, J&)
           End If
        Next I&
        '
        For I& = J& + 1 To FIN&
           Call GRAREG("PROGENT", String$(24, 0), I&)
        Next I&
        Hide
        '
    End If
    '
End Sub

Private Sub Command1_Click()
    Call SELECHO("SELFECHA")
    Text1.Text = VALACT1$("SELFECHA")
End Sub

Private Sub Form_Activate()
    If TRIM$(Text1.Text) = "" Then
        Text1.Text = FECHATEX$(HOY(HO$))
    End If
    Text1.SetFocus
End Sub

Private Sub Text1_DblClick()
    Call SELECHO("SELFECHA")
    Text1.Text = VALACT1$("SELFECHA")
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
    If KeyAscii% = 13 Then
        BOTREC.SetFocus
    End If
End Sub
