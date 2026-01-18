VERSION 5.00
Begin VB.Form MENUENSA 
   BackColor       =   &H00C0FFC0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Ensayos de Bombas"
   ClientHeight    =   1845
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   4590
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   1845
   ScaleWidth      =   4590
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Selección"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1485
      Left            =   210
      TabIndex        =   2
      Top             =   210
      Width           =   3270
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Bombas de Vacío"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   525
         TabIndex        =   5
         Top             =   420
         Value           =   -1  'True
         Width           =   2535
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Bombas DVR"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   525
         TabIndex        =   4
         Top             =   735
         Width           =   2430
      End
      Begin VB.OptionButton Option3 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Dosificadoras"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   525
         TabIndex        =   3
         Top             =   1050
         Width           =   2430
      End
   End
   Begin VB.CommandButton command1 
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
      Left            =   3570
      Picture         =   "MENUENSA.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   315
      Width           =   855
   End
   Begin VB.CommandButton Command2 
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
      Left            =   3570
      Picture         =   "MENUENSA.frx":014A
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1050
      Width           =   855
   End
End
Attribute VB_Name = "MENUENSA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Call CIERRARCH("*.*")
    Call FINAL("")
End Sub

Private Sub Command2_Click()
    Command2.Enabled = False
    If Option1.Value = True Then
        If ENSABVAC.Visible = False Then
          ENSABVAC.Show 1
        End If
      ElseIf Option2.Value = True Then
        If PARDIDVR.Visible = False Then
          PARDIDVR.Show 1
        End If
      ElseIf Option3.Value = True Then
        If PARDI015.Visible = False Then
          PARDI015.Show 1
        End If
    End If
    Command2.Enabled = True
End Sub

Private Sub Form_Load()
    Call VERJOBID(OKEY%)
    '
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call CIERRARCH("*.*")
    Call FINAL("")
End Sub

Private Sub Option1_DblClick()
   If ENSABVAC.Visible = False Then
     ENSABVAC.Show 1
   End If
End Sub

Private Sub Option2_DblClick()
   If PARDIDVR.Visible = False Then
     PARDIDVR.Show 1
   End If
End Sub

Private Sub Option3_DblClick()
   If PARDI015.Visible = False Then
     PARDI015.Show 1
   End If
End Sub
