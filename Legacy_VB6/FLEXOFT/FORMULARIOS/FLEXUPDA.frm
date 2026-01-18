VERSION 5.00
Begin VB.Form FLEXUPDA 
   Caption         =   "Form1"
   ClientHeight    =   6360
   ClientLeft      =   60
   ClientTop       =   315
   ClientWidth     =   7035
   LinkTopic       =   "Form1"
   ScaleHeight     =   6360
   ScaleWidth      =   7035
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List2 
      Height          =   2595
      Left            =   210
      TabIndex        =   3
      Top             =   3360
      Width           =   6630
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   435
      Left            =   5565
      TabIndex        =   2
      Top             =   315
      Width           =   1275
   End
   Begin VB.ListBox List1 
      Height          =   645
      Left            =   210
      TabIndex        =   1
      Top             =   2520
      Width           =   5160
   End
   Begin VB.FileListBox File1 
      Height          =   1845
      Left            =   3045
      TabIndex        =   0
      Top             =   315
      Width           =   2325
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label1"
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
      Left            =   5565
      TabIndex        =   4
      Top             =   2835
      Width           =   1275
   End
End
Attribute VB_Name = "FLEXUPDA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
   '
   Dim NOMARCH$(4096), FEARCH#(4096), RUTARCH$(4096)
   List1.Clear
   List1.AddItem "C:\"
   ILISTA& = 0
   CARCHI% = 0
   '
   ULOAC$ = Left$(CurDir, 2)
   DIREINSTA$ = CurDir
   DIRENOVED$ = ULOAC$ + "\FLEXOFT\NOVEDAD"
   DIRERESPA$ = ULOAC$ + "\FLEXOFT\VIEJOS"
   '
   Call CREACARPE(DIRENOVED$)
   Call CREACARPE(DIRERESPA$)
   Stop
   '
End Sub

Sub CREACARPE(ARXX$)
    '
    BRXX$ = CurDir
63  On Error GoTo 61
    ChDir ARXX$
    On Error GoTo 0
    GoTo 70
    '
61  Resume 62
62  On Error GoTo 0
    Screen.MousePointer = 1
    If MsgBox("Crear Carpeta '" + ARXX$ + "' en Disco Rígido.", 1) <> 1 Then
      Close: End
    End If
    Screen.MousePointer = 11
66  On Error GoTo 67
    MkDir ARXX$
    On Error GoTo 0
    GoTo 63
    '
67  Resume 68
68  On Error GoTo 0
    MsgBox "Imposible Crear Carpeta '" + ARXX$ '."
    Close: End
    '
70  ChDir BRXX$
    '
End Sub
