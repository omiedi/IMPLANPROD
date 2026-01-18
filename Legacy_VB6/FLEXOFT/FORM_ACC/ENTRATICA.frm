VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form ENTRATICA 
   BackColor       =   &H00E0E0D0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   " Tipo de Cambio"
   ClientHeight    =   1245
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   3045
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
   ScaleHeight     =   1245
   ScaleWidth      =   3045
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame4 
      Caption         =   "Frame1"
      Height          =   615
      Left            =   600
      TabIndex        =   8
      Top             =   3000
      Width           =   2175
   End
   Begin VB.CommandButton Command14 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Modificar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Aprobar Cambio de Cotización"
      Top             =   840
      Width           =   1200
   End
   Begin VB.CommandButton Command15 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Cancelar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1680
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Cancelar Cambio de Cotización"
      Top             =   840
      Width           =   1200
   End
   Begin VB.TextBox Text4 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1250
      TabIndex        =   3
      Top             =   280
      Width           =   1380
   End
   Begin VB.PictureBox Picture999 
      Height          =   540
      Index           =   0
      Left            =   1920
      ScaleHeight     =   480
      ScaleWidth      =   1110
      TabIndex        =   1
      Top             =   -240
      Visible         =   0   'False
      Width           =   1170
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
      Picture         =   "ENTRATICA.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   3990
      Width           =   855
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   0
      OleObjectBlob   =   "ENTRATICA.frx":030A
      Top             =   1680
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   435
      Index           =   0
      Left            =   0
      OleObjectBlob   =   "ENTRATICA.frx":053E
      TabIndex        =   2
      Top             =   0
      Visible         =   0   'False
      Width           =   1485
   End
   Begin VB.Label Label6 
      Caption         =   "Label1"
      Height          =   615
      Left            =   240
      TabIndex        =   7
      Top             =   2160
      Width           =   1095
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Cotización"
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
      Left            =   -50
      TabIndex        =   4
      Top             =   400
      Width           =   1170
   End
   Begin VB.Image Image2 
      Height          =   510
      Left            =   5040
      Picture         =   "ENTRATICA.frx":05B8
      Top             =   6950
      Width           =   2010
   End
End
Attribute VB_Name = "ENTRATICA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command14_Click()
  '
  If XVALO(Text4) > 0.1 Then
     Call Text4_LostFocus
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
   
  UTILIZA_SKIN% = 1
  Call APLICACIONSKINS(Me)
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
  Text4 = FORMATNUM(XVALO(Text4), "F10.4")
  '\\\OT-07-0086-RG-FIN///
  '
End Sub
