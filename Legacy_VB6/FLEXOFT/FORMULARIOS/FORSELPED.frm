VERSION 5.00
Begin VB.Form FORSELPED 
   BackColor       =   &H00FFFFC0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Seleccion de Pedidos - Remitos"
   ClientHeight    =   3090
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   4875
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3090
   ScaleWidth      =   4875
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "All"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   720
      Left            =   3990
      Picture         =   "FORSELPED.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Seleccionar-Marcar Todos"
      Top             =   1450
      Width           =   750
   End
   Begin VB.CommandButton BOTREC 
      Caption         =   "Next"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   720
      Left            =   3990
      Picture         =   "FORSELPED.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Grabar - Guardar Modificaciones"
      Top             =   2250
      Width           =   750
   End
   Begin VB.CommandButton BOTEXIT 
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   825
      Left            =   3990
      Picture         =   "FORSELPED.frx":0614
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Cancelar la Aplicación"
      Top             =   105
      Width           =   750
   End
   Begin VB.ListBox LIPEDPE 
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2865
      ItemData        =   "FORSELPED.frx":091E
      Left            =   105
      List            =   "FORSELPED.frx":0920
      Style           =   1  'Checkbox
      TabIndex        =   0
      Top             =   105
      Width           =   3795
   End
End
Attribute VB_Name = "FORSELPED"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub BOTEXIT_Click()
    LIPEDPE.Clear
    Hide
End Sub
'

Private Sub BOTREC_Click()
    Hide
End Sub

Private Sub Command1_Click()
    REACTI& = LIPEDPE.ListIndex
    For U& = 1 To LIPEDPE.ListCount
       LIPEDPE.Selected(U& - 1) = True
    Next U&
    LIPEDPE.ListIndex = REACTI&
End Sub

Private Sub Form_Activate()
    LIPEDPE.SetFocus
End Sub

Private Sub LIPEDPE_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       U& = LIPEDPE.ListIndex
       LIPEDPE.Selected(U&) = True
       If LIPEDPE.ListIndex < LIPEDPE.ListCount - 1 Then
          LIPEDPE.ListIndex = LIPEDPE.ListIndex + 1
       End If
     ElseIf KeyAscii = 27 Then
       U& = LIPEDPE.ListIndex
       LIPEDPE.Selected(U&) = False
       If LIPEDPE.ListIndex < LIPEDPE.ListCount - 1 Then
          LIPEDPE.ListIndex = LIPEDPE.ListIndex + 1
       End If
    End If
End Sub
