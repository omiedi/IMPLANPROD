VERSION 5.00
Begin VB.Form INIMAS04 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Maestro de Items de Stock"
   ClientHeight    =   3060
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   5130
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3060
   ScaleWidth      =   5130
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      FillStyle       =   2  'Horizontal Line
      Height          =   3105
      Left            =   0
      Picture         =   "INIMAS04.frx":0000
      ScaleHeight     =   3105
      ScaleMode       =   0  'User
      ScaleWidth      =   2725.888
      TabIndex        =   0
      Top             =   0
      Width           =   5370
      Begin VB.PictureBox Picture2 
         Height          =   3135
         Left            =   0
         ScaleHeight     =   3075
         ScaleWidth      =   5100
         TabIndex        =   2
         Top             =   0
         Visible         =   0   'False
         Width           =   5161
         Begin VB.Image Image1 
            Height          =   3045
            Left            =   0
            Picture         =   "INIMAS04.frx":36EB6
            Stretch         =   -1  'True
            Top             =   0
            Width           =   5145
         End
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Maestro de Items de Stock "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   900
         Left            =   2835
         TabIndex        =   1
         Top             =   210
         Width           =   2205
      End
   End
End
Attribute VB_Name = "INIMAS04"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
   '
   'Show
   'Refresh
   
   VERANTER$ = "ANACOS01"
   Call VERJOBID(OKEY%)
   If OKEY% < 1 Then Call FINAL("")
   '
   If CONTROL("NEWASPEC", "CAMBIOAS") = "SI" Then
      Picture2.Visible = True
   End If

   If CONTROL("NEWASPEC", "CAMBIOAS") = "SI" Then
      Picture2.Visible = True
    End If

   EPAC$ = TRIM$(UCase$(EMPREAC$))
   If InStr(EPAC$, "GRISFER") > 0 Then
     If HOY(HOS$) > 9647 Then    ' corresponde 15/02/03
       Call GRACONTROL("ANACOSTO", "LMTBNMNT", "SI")
     End If
   End If
   '
   If TRIM$(CONTROL$("ANACOSTO", "LMTBNMNT")) = "SI" Then
     Call CIERRARCH("*.*")
     Call FINAL("")
   End If
   '
   AMASTOK04.Show
   Hide
   '
End Sub

Private Sub Picture1_Click()

End Sub
