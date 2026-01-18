VERSION 5.00
Begin VB.Form INITABU 
   BackColor       =   &H00C0FFC0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Procedimientos Tabulados"
   ClientHeight    =   5265
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   6645
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5265
   ScaleWidth      =   6645
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox LIPROCE 
      BackColor       =   &H00C0E0FF&
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   0
      Sorted          =   -1  'True
      TabIndex        =   7
      Top             =   630
      Visible         =   0   'False
      Width           =   2220
   End
   Begin VB.FileListBox File1 
      Height          =   480
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Visible         =   0   'False
      Width           =   2220
   End
   Begin VB.CommandButton Command5 
      Caption         =   "  Fichero   Plano"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1275
      Left            =   4410
      Picture         =   "INITABU.form.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2100
      Width           =   1065
   End
   Begin VB.CommandButton Command2 
      Caption         =   "  Listador  Tabulado"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1275
      Left            =   2835
      Picture         =   "INITABU.form.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2100
      Width           =   1065
   End
   Begin VB.CommandButton Command1 
      Caption         =   " Formulario Tabulado"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1275
      Left            =   1260
      Picture         =   "INITABU.form.frx":0974
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2100
      Width           =   1065
   End
   Begin VB.CommandButton BOTEXIT 
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
      Left            =   5565
      Picture         =   "INITABU.form.frx":0ABE
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   210
      Width           =   855
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFF00&
      BackStyle       =   0  'Transparent
      Caption         =   "    Sistema Modular Integrado de             Procesamiento de Datos           Linea FLEXOFT (r) - Version 9.03"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Left            =   3465
      TabIndex        =   5
      Top             =   4515
      Width           =   3060
   End
   Begin VB.Label LABEL1 
      BackColor       =   &H00FFFF00&
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
      Left            =   3570
      TabIndex        =   0
      Top             =   6930
      Width           =   7365
   End
   Begin VB.Image Image3 
      Height          =   885
      Left            =   0
      Picture         =   "INITABU.form.frx":0C08
      Stretch         =   -1  'True
      Top             =   4410
      Width           =   10920
   End
End
Attribute VB_Name = "INITABU"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub BOTEXIT_Click()
    '
    If LIPROCE.ListCount > 0 Then
          LIPROCE.Visible = False
          LIPROCE.Clear
          Exit Sub
       Else
          Call CIERRARCH("*.*")
          Call FINAL("")
    End If
    '
End Sub

Private Sub Command1_Click()
    TIPRO% = 1
    Call CARLIPRO
End Sub
Private Sub Command2_Click()
    TIPRO% = 2
    Call CARLIPRO
End Sub
Private Sub Command5_Click()
    TIPRO% = 5
    Call CARLIPRO
End Sub

Private Sub Form_Load()
    '
    EPAC$ = TRIM$(EMPREAC$)
    If EPAC$ <> "" Then
        Caption = Caption + "  -  " + EPAC$
    End If
    '
End Sub

Private Sub LIPROCE_DblClick()
    '
    PROPRIN$ = TRIM$(Left$(INITABU.LIPROCE.Text, 8))
    ATRIBU$ = ""
    Hide
    '
End Sub
