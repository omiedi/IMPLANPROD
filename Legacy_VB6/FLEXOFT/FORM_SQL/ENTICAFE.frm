VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form ENTICAFE 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   " Tipo de Cambio"
   ClientHeight    =   1995
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   4290
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1995
   ScaleWidth      =   4290
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture999 
      Height          =   300
      Index           =   0
      Left            =   1920
      ScaleHeight     =   240
      ScaleWidth      =   1110
      TabIndex        =   8
      Top             =   0
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00800000&
      Height          =   2850
      Left            =   3390
      ScaleHeight     =   2790
      ScaleWidth      =   1110
      TabIndex        =   6
      Top             =   0
      Width           =   1170
      Begin VB.CommandButton Command15 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   105
         Picture         =   "ENTICAFE.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Cancelar Cambio de Cotización"
         Top             =   210
         Width           =   600
      End
      Begin VB.CommandButton Command14 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   105
         Picture         =   "ENTICAFE.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Aprobar Cambio de Cotización"
         Top             =   1365
         Width           =   600
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1920
      Left            =   120
      TabIndex        =   4
      Top             =   0
      Width           =   3210
      Begin VB.CommandButton Command9 
         Caption         =   "."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   2820
         TabIndex        =   11
         Top             =   360
         Width           =   195
      End
      Begin VB.TextBox Text9 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
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
         Left            =   1860
         TabIndex        =   10
         TabStop         =   0   'False
         Text            =   " "
         Top             =   360
         Width           =   960
      End
      Begin VB.TextBox Text4 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1560
         TabIndex        =   0
         Top             =   1365
         Width           =   1380
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004040&
         Height          =   210
         Left            =   1200
         TabIndex        =   12
         Top             =   465
         Width           =   585
      End
      Begin VB.Label Label6 
         BackColor       =   &H00C0FFC0&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   315
         TabIndex        =   7
         Top             =   1050
         Visible         =   0   'False
         Width           =   960
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Cotización:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   345
         TabIndex        =   5
         Top             =   1470
         Width           =   1170
      End
   End
   Begin VB.CommandButton Command7 
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
      Left            =   5600
      Picture         =   "ENTICAFE.frx":0614
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3990
      Width           =   855
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   0
      OleObjectBlob   =   "ENTICAFE.frx":091E
      Top             =   1680
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   435
      Index           =   0
      Left            =   60
      OleObjectBlob   =   "ENTICAFE.frx":0B52
      TabIndex        =   9
      Top             =   0
      Visible         =   0   'False
      Width           =   1485
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   255
      Left            =   0
      TabIndex        =   13
      Top             =   360
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Image Image2 
      Height          =   510
      Left            =   5040
      Picture         =   "ENTICAFE.frx":0BCC
      Top             =   6950
      Width           =   2010
   End
End
Attribute VB_Name = "ENTICAFE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command14_Click()
  '
  If XVALO(Text4) > 0.1 Then
     Call Text4_LostFocus
     Label1.Caption = "OK"
     Hide
  End If
  '\\\OT-07-0086-RG-FIN///
  '
End Sub

Private Sub Command15_Click()
  '
  Unload Me
  '
End Sub

Private Sub Command9_Click()
    Command9.Enabled = False
    Call SELECHO("SELFECHA")
    VDEVV$ = VALACT1$("SELFECHA")
    If VDEVV$ <> "" Then
        Text9.TEXT = VDEVV$
        FF% = FECHANUM(Text9.TEXT)
    End If
    Command9.Enabled = True
End Sub

Private Sub Form_Activate()
   On Error Resume Next
   Text4.SetFocus
   On Error GoTo 0
End Sub

Private Sub Form_Load()
  If CONTROL("NEWASPEC", "CAMBIOAS") = "SI" Then
      Me.BackColor = &HFCD7B6
      Me.Frame4.BackColor = &HFCD7B6
  End If
   
  Call APLICACIONSKINS(Me, Picture1)
  On Error Resume Next
  Text4.SetFocus
  On Error GoTo 0
End Sub

Private Sub Text4_GotFocus()
  Text4.SelStart = 0
  Text4.SelLength = Len(Text4)
End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
     Call Command14_Click
   End If
End Sub

Private Sub Text4_LostFocus()
  '
  '\\\OT-07-0086-RG-06/03/07///
  Text4 = FORMATNUM(XVALO(Text4), "F12.4")
  '\\\OT-07-0086-RG-FIN///
  '
End Sub

Private Sub Text9_Change()
   If FECHANUM(Text9) > 0 Then
      Text4 = FORMATNUM$(COTIFECHA_HISTODOL#(FECHANUM(Text9), 2), "F12.4")
      Call PINTATEXT(Text4)
   End If
End Sub
