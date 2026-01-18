VERSION 5.00
Begin VB.Form FECHYUSER 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Filtro"
   ClientHeight    =   1560
   ClientLeft      =   45
   ClientTop       =   10305
   ClientWidth     =   7200
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1560
   ScaleWidth      =   7200
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton BOTREC 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   6600
      Picture         =   "FECHYUSER.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Graba - Guardar Modificaciones"
      Top             =   920
      Width           =   540
   End
   Begin VB.TextBox TXTDESCRIUSER 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   2220
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   1080
      Width           =   4245
   End
   Begin VB.CommandButton Command3 
      Caption         =   "."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   2025
      TabIndex        =   5
      Top             =   1080
      Width           =   175
   End
   Begin VB.TextBox TXTUSER 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   960
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   1080
      Width           =   1070
   End
   Begin VB.CommandButton Command2 
      Caption         =   "."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   2025
      TabIndex        =   3
      Top             =   600
      Width           =   175
   End
   Begin VB.TextBox TXTFECHA2 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   960
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   600
      Width           =   1070
   End
   Begin VB.CommandButton Command1 
      Caption         =   "."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   2025
      TabIndex        =   1
      Top             =   120
      Width           =   175
   End
   Begin VB.TextBox TXTFECHA1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   960
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   120
      Width           =   1070
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Usuario:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   120
      TabIndex        =   10
      Top             =   1200
      Width           =   750
   End
   Begin VB.Label Label11 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Hasta:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   120
      TabIndex        =   9
      Top             =   720
      Width           =   750
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Desde:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   0
      TabIndex        =   8
      Top             =   240
      Width           =   855
   End
End
Attribute VB_Name = "FECHYUSER"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public APROBADO%

Private Sub BOTREC_Click()
   APROBADO% = 1
   Hide
End Sub

Private Sub Command1_Click()
  Call SELECHO("SELFECHA")
  VDEV$ = VALACT1$("SELFECHA")
  If VDEV$ <> "" Then
     TXTFECHA1 = VDEV$
  End If

End Sub

Private Sub Command2_Click()
  Call SELECHO("SELFECHA")
  VDEV$ = VALACT1$("SELFECHA")
  If VDEV$ <> "" Then
     TXTFECHA2 = VDEV$
  End If

End Sub

Private Sub Command3_Click()
    Command3.Enabled = False
    Call SELECHO("FLEXUSR")
    Screen.MousePointer = 1
    If VALACT1$("FLEXUSR") = "" Then Exit Sub
    '
    USEL% = XVALO(TRIM$(VALACT1$("FLEXUSR")))
    DEUSER$ = VALACT2$("FLEXUSR")
    If USEL% > 0 Then
      Me.TXTUSER.TEXT = USEL%
      Me.TXTDESCRIUSER = DEUSER$
    End If
    Command3.Enabled = True

End Sub

