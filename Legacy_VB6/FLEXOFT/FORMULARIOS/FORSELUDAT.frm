VERSION 5.00
Begin VB.Form FORSELUDAT 
   BackColor       =   &H00C0FFC0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Opcion Multiempresa - Seleccion"
   ClientHeight    =   2280
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   6030
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2280
   ScaleWidth      =   6030
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command21 
      Caption         =   "Help"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   5040
      Picture         =   "FORSELUDAT.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Ayuda en Línea"
      Top             =   1200
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "End"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   5040
      Picture         =   "FORSELUDAT.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Terminar Sesión de Trabajo - Cerrar Sistema"
      Top             =   210
      Width           =   855
   End
   Begin VB.ListBox List2 
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1860
      Left            =   210
      TabIndex        =   0
      Top             =   210
      Visible         =   0   'False
      Width           =   4635
   End
End
Attribute VB_Name = "FORSELUDAT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub command1_Click()
    Call FUERASYS
End Sub

Private Sub List2_DblClick()
    '
    REMPRE$ = List2.TEXT
    EMPREAC$ = TRIM$(Left$(REMPRE$, 30))
    LUSELEC$ = TRIM$(Mid$(REMPRE$, 51, 46))
    Call GRARAM(193, AJUSTI$(EMPREAC$, 30))
    '
    Dim FF0 As String * 128
    N1% = FILEOPEN%("FLEXCFG.DAT", 0, 128)
    '
    Get #N1%, 1, FF0$
    FF00$ = FF0$
    Call REPLA(FF00$, LUSELEC$, 51, 45)
    LSet FF0$ = FF00$
    Put #N1%, 1, FF0$
    Close #N1%: N1% = 0
    '
    If Len(LUSELEC$) > 0 Then
      If Right$(LUSELEC$, 1) <> "\" Then
        LUSELEC$ = LUSELEC$ + "\"
      End If
    End If
    '
    LUDAT$ = ULODATOS$ + LUSELEC$
    Call GRARAM(99, AJUSTI$(LUDAT$, 46))
    '
    EPAC$ = TRIM$(EMPREAC$)
    If EPAC$ <> "" Then
      On Error Resume Next
      FLEXOFTA.Caption = "Microsoft Visual Basic (r) - V.5.0 - " + EPAC$
      On Error GoTo 0
    End If
    Hide
    '
End Sub
