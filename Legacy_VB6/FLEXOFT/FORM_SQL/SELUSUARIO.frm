VERSION 5.00
Begin VB.Form SELUSUARIO 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Ajustar Cuenta de Cliente - Pagos x Impuestos"
   ClientHeight    =   3120
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7770
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3120
   ScaleWidth      =   7770
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Grabar"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   12
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   5640
      TabIndex        =   4
      Top             =   2520
      Width           =   1935
   End
   Begin VB.CommandButton Command6 
      Caption         =   "!"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   1320
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   1680
      Width           =   420
   End
   Begin VB.Label lblNUSER 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   15.75
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   5
      Top             =   1680
      Width           =   1095
   End
   Begin VB.Label lblDescUser 
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   15.75
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1800
      TabIndex        =   2
      Top             =   1680
      Width           =   5775
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Usuario Seleccionado:"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   12
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   1080
      Width           =   3135
   End
   Begin VB.Label Label6 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Selección de Usuario Responsable de Pago"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   15.75
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7575
   End
End
Attribute VB_Name = "SELUSUARIO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
   Hide
End Sub

Private Sub Command6_Click()
    Command6.Enabled = False
    Screen.MousePointer = 11
    Call BLOQARCH("FLEXUSR")
    A = LUDAT$
    FUSER$ = LOADFILE$("USER.DAT")
    JJI& = 0: KKI% = 0
    For I% = 1 To Len(FUSER$) Step 64
       F1$ = Mid$(FUSER$, I%, 64)
       If Asc(F1$) = 0 Then GoTo 9
       NOM$ = Left$(F1$, 30)
       If Left$(NOM$, 1) = "@" Then NOM$ = DESENCRI$(Mid$(F1$, 2, 29))
       '
       KKI% = KKI% + 1
       x$ = REGBLAN$("FLEXUSR")
       Call REPLA$(x$, MKI$(KKI%), 1, 2)
       Call REPLA$(x$, NOM$, 3, 30)
       JJI& = JJI& + 1
       Call GRAREG("FLEXUSR", x$, JJI&)
    Next I%
9   Call SETULTREG("FLEXUSR", JJI&)
10  Call CIERRARCH("FLEXUSR")

    
    Call SELECHO("FLEXUSR")
    Screen.MousePointer = 1
    If VALACT1$("FLEXUSR") = "" Then Exit Sub
    '
    USEL% = XVALO(TRIM$(VALACT1$("FLEXUSR")))
    DEUSER$ = VALACT2$("FLEXUSR")
    lblNUSER = USEL%
    'lblDescUser = DEUSER$
    
    Command6.Enabled = True
End Sub



Private Sub lblNUSER_Change()
    lblDescUser = USERNAM$(XVALO(lblNUSER))
End Sub

