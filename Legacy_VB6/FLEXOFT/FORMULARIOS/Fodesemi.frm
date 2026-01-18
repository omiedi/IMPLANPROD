VERSION 5.00
Begin VB.Form FODESEMI 
   BackColor       =   &H00C0E0FF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Seleccion de Pedidos a Despachar"
   ClientHeight    =   2985
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   5415
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2985
   ScaleWidth      =   5415
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
      Left            =   4500
      Picture         =   "Fodesemi.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Seleccionar-Marcar Todos"
      Top             =   1300
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
      Left            =   4500
      Picture         =   "Fodesemi.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Grabar - Guardar Modificaciones"
      Top             =   2130
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
      Left            =   4500
      Picture         =   "Fodesemi.frx":0614
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Cancelar la Aplicación"
      Top             =   105
      Width           =   750
   End
   Begin VB.ListBox LIPEDES 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2760
      ItemData        =   "Fodesemi.frx":091E
      Left            =   120
      List            =   "Fodesemi.frx":0925
      Sorted          =   -1  'True
      Style           =   1  'Checkbox
      TabIndex        =   0
      Top             =   120
      Width           =   4245
   End
End
Attribute VB_Name = "FODESEMI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub BOTEXIT_Click()
    LIPEDES.Clear
    Hide
End Sub
'

Private Sub BOTREC_Click()
    Hide
End Sub

Private Sub Command1_Click()
    REACTI& = LIPEDES.ListIndex
    For U& = 1 To LIPEDES.ListCount
       LIPEDES.Selected(U& - 1) = True
    Next U&
    LIPEDES.ListIndex = REACTI&
End Sub

Private Sub Form_Activate()
    LIPEDES.SetFocus
End Sub

Private Sub LIPEDES_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       U& = LIPEDES.ListIndex
       LIPEDES.Selected(U&) = True
       If LIPEDES.ListIndex < LIPEDES.ListCount - 1 Then
          LIPEDES.ListIndex = LIPEDES.ListIndex + 1
       End If
     ElseIf KeyAscii = 27 Then
       U& = LIPEDES.ListIndex
       LIPEDES.Selected(U&) = False
       If LIPEDES.ListIndex < LIPEDES.ListCount - 1 Then
          LIPEDES.ListIndex = LIPEDES.ListIndex + 1
       End If
    End If
End Sub
