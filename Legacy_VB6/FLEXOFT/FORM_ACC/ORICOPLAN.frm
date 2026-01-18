VERSION 5.00
Begin VB.Form ORICOPLAN 
   BackColor       =   &H00C0D0D0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Origen de Copia"
   ClientHeight    =   1245
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5250
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1245
   ScaleWidth      =   5250
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Height          =   700
      Left            =   4440
      Picture         =   "ORICOPLAN.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   360
      Width           =   615
   End
   Begin VB.ListBox List1 
      Height          =   645
      Left            =   480
      Sorted          =   -1  'True
      TabIndex        =   2
      Top             =   1200
      Visible         =   0   'False
      Width           =   3735
   End
   Begin VB.ComboBox Combo1 
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
      Left            =   480
      Sorted          =   -1  'True
      TabIndex        =   1
      Top             =   720
      Width           =   3735
   End
   Begin VB.Label LUACCORI 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "LUACCORI"
      Height          =   255
      Left            =   480
      TabIndex        =   4
      Top             =   2040
      Visible         =   0   'False
      Width           =   3735
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Copia Plan de Cuentas desde:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   3975
   End
End
Attribute VB_Name = "ORICOPLAN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
  '
  Call CIERRARCH("*.*")
  Call BAJALDISCO
  LUACCORI = ""
  '
  ICOMBO& = Combo1.ListIndex
  LUDA1$ = ULODATOS + TRIM$(Mid$(List1.List(ICOMBO&), 31)) + "."
  LUDA2$ = ULODATOS + TRIM$(Mid$(List1.List(ICOMBO&), 31)) + "\"
  '
  EXISZ$ = ""
  On Error Resume Next
  EXISZ$ = TRIM$(Dir(LUDA1$, 16))
  On Error GoTo 0
  '
  If EXISZ$ = "" Then
     MsgBox "No Existe Carpeta " + LUDA2$
     Exit Sub
  End If
  '
  EMPREORI$ = TRIM$(Left$(List1.List(ICOMBO&), 30))
  LUSELORI$ = TRIM$(Mid$(List1.List(ICOMBO&), 31))
  LUDATORI$ = LUDA2$
  '
  If EMPREORI$ = "" Then
     MsgBox "Debe Seleccionar Empresa Origen"
     Exit Sub
  End If
  '
5 EXISZ$ = ""
  On Error Resume Next
  EXISZ$ = TRIM$(Dir(LUDATORI$ + "BdAccess.", 16))
  On Error GoTo 0
  If EXISZ$ = "" Then
    MsgBox "No Existe Carpeta " + LUDATORI$ + "BdAccess\"
    Exit Sub
  End If
  '
  LUACCORI = LUDATORI$ + "BdAccess\"
  Hide
  '
End Sub

Private Sub Form_Load()
   
    If CONTROL("NEWASPEC", "CAMBIOAS") = "SI" Then
      Me.BackColor = &HFCD7B6
    End If
   '
   Combo1.Clear: List1.Clear
   LUACCORI = ""
   '
   If EXISTE%(App.Path + "\FLEXMUL.DRV") = 1 Then
     LIEMP$ = LOADFILE$(App.Path + "\FLEXMUL.DRV")
     Dim FMX As String * 79
     '
     For i% = 1 To Len(LIEMP$) Step 79
       FMX$ = Mid$(LIEMP$, i%, 79)
       SUCAR% = 0
       For J% = 1 To 77
         SUCAR% = SUCAR% + Asc(Mid$(FMX$, J%, 1))
       Next J%
       If CVI(Mid$(FMX$, 78, 2)) = SUCAR% Then
         RASO$ = DESENCRI$(Left$(FMX$, 30))
         PACCE$ = DESENCRI$(Mid$(FMX$, 32, 46))
         Z$ = Space$(96)
         Call REPLA(Z$, RASO$, 1, 30)
         Call REPLA(Z$, PACCE$, 51, 46)
         Combo1.AddItem TRIM$(Left$(Z$, 30))
         List1.AddItem Z$
       End If
     Next i%
     '
   End If
   '
End Sub

