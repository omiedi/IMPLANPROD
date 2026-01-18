VERSION 5.00
Begin VB.Form FLEMENUS 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00800000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   4950
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   9615
   ControlBox      =   0   'False
   Icon            =   "FLEMENUS.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4950
   ScaleWidth      =   9615
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command3 
      Caption         =   "Qwery"
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
      Left            =   8550
      Picture         =   "FLEMENUS.frx":1D2A
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Consultas y Listados"
      Top             =   3045
      Width           =   855
   End
   Begin VB.CommandButton Command4 
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
      Left            =   8550
      Picture         =   "FLEMENUS.frx":1E74
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   1155
      Width           =   855
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Exit"
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
      Left            =   8550
      Picture         =   "FLEMENUS.frx":217E
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   210
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Util"
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
      Left            =   8550
      Picture         =   "FLEMENUS.frx":22C8
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2100
      Width           =   855
   End
   Begin VB.ListBox OPMENU 
      BackColor       =   &H00808000&
      DragMode        =   1  'Automatic
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   330
      Index           =   4
      Left            =   3780
      TabIndex        =   4
      Top             =   3465
      Width           =   4200
   End
   Begin VB.ListBox OPMENU 
      BackColor       =   &H00C0C000&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   330
      Index           =   3
      ItemData        =   "FLEMENUS.frx":25D2
      Left            =   2940
      List            =   "FLEMENUS.frx":25D4
      TabIndex        =   3
      Top             =   2730
      Width           =   4200
   End
   Begin VB.ListBox OPMENU 
      BackColor       =   &H00FFFF00&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   330
      Index           =   2
      ItemData        =   "FLEMENUS.frx":25D6
      Left            =   2100
      List            =   "FLEMENUS.frx":25D8
      TabIndex        =   2
      Top             =   1995
      Width           =   4200
   End
   Begin VB.ListBox OPMENU 
      BackColor       =   &H00FFFF80&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   1
      ItemData        =   "FLEMENUS.frx":25DA
      Left            =   1260
      List            =   "FLEMENUS.frx":25DC
      TabIndex        =   1
      Top             =   1260
      Width           =   4200
   End
   Begin VB.ListBox OPMENU 
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   330
      Index           =   0
      ItemData        =   "FLEMENUS.frx":25DE
      Left            =   420
      List            =   "FLEMENUS.frx":25E5
      TabIndex        =   0
      Top             =   525
      Width           =   4200
   End
   Begin VB.Timer Timer1 
      Left            =   0
      Top             =   0
   End
   Begin VB.Image Image1 
      Height          =   5300
      Index           =   1
      Left            =   0
      Picture         =   "FLEMENUS.frx":25F1
      Stretch         =   -1  'True
      Top             =   0
      Width           =   8280
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   8295
      X2              =   8295
      Y1              =   0
      Y2              =   5880
   End
   Begin VB.Image Image2 
      Height          =   510
      Left            =   8000
      Picture         =   "FLEMENUS.frx":4D737
      Top             =   4450
      Width           =   2010
   End
End
Attribute VB_Name = "FLEMENUS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub command1_Click()
    Form4.Show
End Sub

Private Sub Command2_Click()
    Hide
    'If FLESCRIT.List1.ListCount <= 1 Then
    '     Close
    '     End
    '   Else
         FLESCRIT.Show
    ' End If
End Sub

Private Sub Command3_Click()
    Call GRAJOBID
    Call BAJRAM
    Shell LUBAS$ + "WINTABU " + "#" + JOBID$ + "#", 1
End Sub

Private Sub Command4_Click()
    Call COMUNI("Ayuda en Línea no Disponible")
End Sub

Private Sub Command5_Click()
    '
    REG% = (-1)
    Call SELMENU(MENUPRIN$, REG%)
    '
End Sub

Private Sub Form_Activate()
    '
    EPAC$ = TRIM$(EMPREAC$)
    If EPAC$ <> "" Then
        Caption = Caption + "  -  " + EPAC$
    End If
    '
    FLESCRIT.Hide
    APLI$ = MENUPRIN$
    Call SELMENU(APLI$, REG%)
End Sub

Private Sub Form_Click()
    '
    REG% = (-1)
    Call SELMENU(MENUPRIN$, REG%)
    '
End Sub

Private Sub Form_Load()
    FLESCRIT.Hide
End Sub


Private Sub Image1_Click(Index As Integer)
    '
    REG% = (-1)
    Call SELMENU(MENUPRIN$, REG%)
    '
End Sub

Private Sub OPMENU_DblClick(Index As Integer)
    TTXX$ = OPMENU(Index).TEXT
    OPX$ = Left$(TTXX$, 32)
    MODU$ = TRIM$(Mid$(TTXX$, 65, 8))
    If MODU$ = "" Then
         REG% = Val(Mid$(TTXX$, 73))
         Call SELMENU(MENUPRIN$, REG%)
       Else
         Call CONECRUN(MODU$, OPX$)
    End If
End Sub

