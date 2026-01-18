VERSION 5.00
Begin VB.Form SELSUBOR 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0E0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Selección Sub-Ordenes a Emitir"
   ClientHeight    =   3915
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   10500
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3915
   ScaleWidth      =   10500
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00004000&
      Height          =   4575
      Left            =   9555
      ScaleHeight     =   4515
      ScaleWidth      =   1035
      TabIndex        =   3
      Top             =   0
      Width           =   1095
      Begin VB.CommandButton Command1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   650
         Left            =   120
         Picture         =   "SELSUBOR.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Refresca - Invierte Selección"
         Top             =   1480
         Width           =   650
      End
      Begin VB.PictureBox MARCOBARRA 
         BackColor       =   &H000000FF&
         Height          =   135
         Left            =   120
         ScaleHeight     =   75
         ScaleWidth      =   585
         TabIndex        =   5
         Top             =   3030
         Width           =   650
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H00FF0000&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   0
            TabIndex        =   6
            Top             =   0
            Width           =   12000
         End
      End
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
         Height          =   650
         Left            =   120
         Picture         =   "SELSUBOR.frx":014A
         Style           =   1  'Graphical
         TabIndex        =   0
         ToolTipText     =   "Aprueba la Pre-Selección"
         Top             =   800
         Width           =   650
      End
      Begin VB.CommandButton BOTEXIT 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   650
         Left            =   120
         Picture         =   "SELSUBOR.frx":0454
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Cancela la Pre-Selección"
         Top             =   120
         Width           =   650
      End
      Begin VB.Image Image2 
         Height          =   390
         Left            =   -300
         Picture         =   "SELSUBOR.frx":075E
         Top             =   3320
         Width           =   1515
      End
   End
   Begin VB.ListBox LIMATSEL 
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3630
      Left            =   105
      Style           =   1  'Checkbox
      TabIndex        =   1
      Top             =   105
      Width           =   9255
   End
   Begin VB.ListBox List1 
      Height          =   645
      Left            =   3045
      Sorted          =   -1  'True
      TabIndex        =   2
      Top             =   2205
      Visible         =   0   'False
      Width           =   1065
   End
   Begin VB.Label Label2 
      Caption         =   "Label1"
      Height          =   540
      Left            =   1575
      TabIndex        =   8
      Top             =   4095
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.Line Line2 
      X1              =   9360
      X2              =   9360
      Y1              =   105
      Y2              =   3750
   End
   Begin VB.Line Line1 
      X1              =   105
      X2              =   9345
      Y1              =   3735
      Y2              =   3735
   End
End
Attribute VB_Name = "SELSUBOR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub BOTEXIT_Click()
    LIMATSEL.Clear
    Label2.Caption = ""
    Hide
End Sub
'

Private Sub BOTREC_Click()
   Hide
End Sub

Private Sub Command1_Click()
   Command1.Enabled = False
   TPI = LIMATSEL.TopIndex
   For U& = 1 To LIMATSEL.ListCount
     If LIMATSEL.Selected(U& - 1) = True Then
         LIMATSEL.Selected(U& - 1) = False
       Else
         LIMATSEL.Selected(U& - 1) = True
     End If
   Next U&
   On Error Resume Next
   LIMATSEL.TopIndex = TPI
   On Error GoTo 0
   DoEvents
   Command1.Enabled = True
End Sub

Private Sub Command9_Click()
   Call SELECHO("SELFECHA")
   VDEVU$ = TRIM$(VALACT1$("SELFECHA"))
   If VDEVU$ <> "" Then
      Label2 = VALACT1$("SELFECHA")
   End If
End Sub

