VERSION 5.00
Begin VB.Form SELUDAT1 
   BackColor       =   &H00C0E0FF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Opcion Multiempresa - Seleccion"
   ClientHeight    =   1980
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   4515
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1980
   ScaleWidth      =   4515
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
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
      Height          =   750
      Left            =   3600
      Picture         =   "SELUDAT1.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Terminar Sesión de Trabajo - Cerrar Sistema"
      Top             =   210
      Width           =   750
   End
   Begin VB.ListBox List2 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1620
      Left            =   210
      TabIndex        =   0
      Top             =   210
      Visible         =   0   'False
      Width           =   3165
   End
   Begin VB.Line Line2 
      X1              =   3375
      X2              =   3375
      Y1              =   210
      Y2              =   1830
   End
   Begin VB.Line Line1 
      X1              =   210
      X2              =   4515
      Y1              =   1830
      Y2              =   1830
   End
   Begin VB.Image Image3 
      Height          =   510
      Left            =   2970
      Picture         =   "SELUDAT1.frx":014A
      Top             =   1320
      Width           =   2010
   End
End
Attribute VB_Name = "SELUDAT1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub command1_Click()
   Close: End
End Sub

Sub List2_DblClick()
    '
    REMPRE$ = List2.Text
    EMPREAC$ = Left$(Trim$(Left$(REMPRE$, 30)) + Space$(30), 30)
    LUSELEC$ = Trim$(Mid$(REMPRE$, 51, 46))
'MsgBox 1
    'Call GRARAM(193, Left(EMPREAC$ + Space$(30), 30))
    '
    Dim FF0 As String * 128
    N1% = FILEOPEN%("FLEXCFG.DAT", 0, 128)
'MsgBox 2
    '
    Get #N1%, 1, FF0$
    FF00$ = FF0$
    Call REPLA(FF00$, LUSELEC$, 51, 45)
    LSet FF0$ = FF00$
    Put #N1%, 1, FF0$
    Close #N1%: N1% = 0
'MsgBox 3 & LUSELEC$
    '
    RAMSTRI$ = Left$(LOADFILE(LUCOM$ + "FLEXRAM.DRV") + String$(512, 0), 512)
'MsgBox 4
    Call REPLA(RAMSTRI$, EMPREAC$, 193, 30)
'MsgBox 5
    Call SAVEFILE(LUCOM$ + "FLEXRAM.DRV", RAMSTRI$)
'MsgBox 6
    'If Len(LUSELEC$) > 0 Then
    '  If Right$(LUSELEC$, 1) <> "\" Then
    '    LUSELEC$ = LUSELEC$ + "\"
    '  End If
    'End If
    '
    'LUDAT$ = ULODATOS$ + LUSELEC$
    'Call GRARAM(99, AJUSTI$(LUDAT$, 46))
    '
    'EPAC$ = TRIM$(EMPREAC$)
    'If EPAC$ <> "" Then
    '  On Error Resume Next
    '  FLEXOFTA.Caption = "Microsoft Visual Basic (r) - V.5.0 - " + EPAC$
    '  On Error GoTo 0
    'End If
    Hide
    '
End Sub

Private Sub List2_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
       Call List2_DblClick
     ElseIf KeyAscii = 27 Then
       Call command1_Click
   End If
End Sub
