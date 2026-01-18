VERSION 5.00
Begin VB.Form FORM3 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0FFC0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   7620
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   10875
   ControlBox      =   0   'False
   Icon            =   "FLEXOFT3.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7620
   ScaleWidth      =   10875
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command5 
      Caption         =   "Nivel Anterior"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1065
      Left            =   525
      Picture         =   "FLEXOFT3.frx":1D2A
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   5145
      Width           =   1065
   End
   Begin VB.CommandButton Command4 
      Caption         =   " Ayuda en Línea"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1065
      Left            =   9555
      Picture         =   "FLEXOFT3.frx":1E74
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   2625
      Width           =   1065
   End
   Begin VB.CommandButton Command3 
      Caption         =   " Consultas  Listados"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1065
      Left            =   9555
      Picture         =   "FLEXOFT3.frx":217E
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   5145
      Width           =   1065
   End
   Begin VB.CommandButton Command2 
      Caption         =   " Terminar Sesion"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1065
      Left            =   9555
      Picture         =   "FLEXOFT3.frx":22C8
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   210
      Width           =   1065
   End
   Begin VB.CommandButton Command1 
      Caption         =   " Utilitarios FLEXOFT"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1065
      Left            =   9555
      Picture         =   "FLEXOFT3.frx":2412
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   3885
      Width           =   1065
   End
   Begin VB.ListBox OPMENU 
      Appearance      =   0  'Flat
      BackColor       =   &H00008080&
      DragMode        =   1  'Automatic
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   330
      Index           =   4
      Left            =   4305
      TabIndex        =   4
      Top             =   3885
      Width           =   4950
   End
   Begin VB.ListBox OPMENU 
      Appearance      =   0  'Flat
      BackColor       =   &H0000C0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   330
      Index           =   3
      Left            =   3360
      TabIndex        =   3
      Top             =   3045
      Width           =   4950
   End
   Begin VB.ListBox OPMENU 
      Appearance      =   0  'Flat
      BackColor       =   &H0000FFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   330
      Index           =   2
      Left            =   2415
      TabIndex        =   2
      Top             =   2205
      Width           =   4950
   End
   Begin VB.ListBox OPMENU 
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   1
      Left            =   1470
      TabIndex        =   1
      Top             =   1365
      Width           =   4950
   End
   Begin VB.ListBox OPMENU 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   330
      Index           =   0
      Left            =   525
      TabIndex        =   0
      Top             =   525
      Width           =   4950
   End
   Begin VB.Timer Timer1 
      Left            =   0
      Top             =   0
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "    Sistema Modular Integrado de Procesamiento de Datos                              Linea FLEXOFT (r) - Version 9.03"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Left            =   3465
      TabIndex        =   6
      Top             =   6825
      Width           =   7365
   End
   Begin VB.Image Image3 
      Height          =   885
      Left            =   0
      Picture         =   "FLEXOFT3.frx":271C
      Stretch         =   -1  'True
      Top             =   6720
      Width           =   10920
   End
End
Attribute VB_Name = "FORM3"
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

