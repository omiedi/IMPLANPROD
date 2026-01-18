VERSION 5.00
Begin VB.Form IFLFRM09 
   BackColor       =   &H00C0E0FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Consulta Documentos Emitidos"
   ClientHeight    =   825
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   6255
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   825
   ScaleWidth      =   6255
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Espere un Momento, por Favor . . . . . . . ."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1320
      TabIndex        =   0
      Top             =   240
      Width           =   5175
   End
End
Attribute VB_Name = "IFLFRM09"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    '
    Call CENTRAFORM(Me)
    '
    Call VERJOBID(OKEY%)
    If OKEY% < 1 Then Call FINAL("")
    '
    Show
    Refresh
    DoEvents
    '
    Screen.MousePointer = 1
    '
    DOCUNU& = XVALO(TRIM$(APLINVO$))
    If DOCUNU& > 0 Then
        Call MUESTRADOC(DOCUNU&, MMDDXX$)
        Call CIERRARCH("*.*")
        Call FINAL("")
    End If
    
5   Call SELECHO("PDOCLST")
    DOCUNU& = XVALO(VALACT1$("PDOCLST"))
    If DOCUNU& < 1 Then
       Call FINAL("")
    End If
    REDOCU$ = REGACT$("PDOCLST")
    MMDDXX$ = ""
    If InStr(REDOCU$, "0007-00000032") > 0 Then MMDDXX$ = "RECUP"
    ORIPRINTER$ = TRIM$(Mid$(REDOCU$, 129))
    '
    Call MUESTRADOC(DOCUNU&, MMDDXX$)
    GoTo 5
    '
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call CIERRARCH("*.*")
   Call FINAL("")
End Sub

