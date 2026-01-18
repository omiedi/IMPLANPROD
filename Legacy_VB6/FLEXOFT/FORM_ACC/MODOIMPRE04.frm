VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "ACTSKIN4.OCX"
Begin VB.Form MODOIMPRE04 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Control de Impresión"
   ClientHeight    =   1305
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4410
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1305
   ScaleWidth      =   4410
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture999 
      Height          =   540
      Index           =   0
      Left            =   1680
      ScaleHeight     =   480
      ScaleWidth      =   1110
      TabIndex        =   5
      Top             =   -210
      Visible         =   0   'False
      Width           =   1170
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
      Height          =   875
      Left            =   3780
      Picture         =   "MODOIMPRE04.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Aprobar"
      Top             =   210
      Width           =   540
   End
   Begin VB.Frame Frame10 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Impresión"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   960
      Left            =   105
      TabIndex        =   0
      Top             =   105
      Width           =   3600
      Begin VB.CheckBox Check8 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Impresión Fantasma"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   510
         Left            =   2205
         TabIndex        =   2
         Top             =   315
         Width           =   1200
      End
      Begin VB.CheckBox Check7 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Imprime O/Venta"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   510
         Left            =   240
         TabIndex        =   1
         Top             =   315
         Value           =   1  'Checked
         Width           =   1665
      End
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   0
      OleObjectBlob   =   "MODOIMPRE04.frx":030A
      Top             =   1800
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   435
      Index           =   0
      Left            =   300
      OleObjectBlob   =   "MODOIMPRE04.frx":053E
      TabIndex        =   6
      Top             =   0
      Visible         =   0   'False
      Width           =   1485
   End
   Begin VB.Label Label1 
      Height          =   435
      Left            =   105
      TabIndex        =   3
      Top             =   1890
      Width           =   1380
   End
End
Attribute VB_Name = "MODOIMPRE04"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command14_Click()
    Label1.Caption = "OK"
    Hide
End Sub

Private Sub Check7_GotFocus()
   Check8.SetFocus
End Sub

Private Sub Form_Load()
    '
    If CONTROL("NEWASPEC", "CAMBIOAS") = "SI" Then
      Me.BackColor = &HFCD7B6
    End If

    Call APLICACIONSKINS(Me)
End Sub

