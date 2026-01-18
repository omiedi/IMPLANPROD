VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form FILEORIG 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Origen del Archivo"
   ClientHeight    =   1665
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3375
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   1665
   ScaleWidth      =   3375
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture999 
      Height          =   330
      Index           =   0
      Left            =   1470
      ScaleHeight     =   270
      ScaleWidth      =   1110
      TabIndex        =   4
      Top             =   525
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cancelar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   1785
      TabIndex        =   2
      Top             =   1050
      Width           =   1170
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Aceptar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   420
      TabIndex        =   1
      Top             =   1050
      Width           =   1170
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   420
      TabIndex        =   0
      Top             =   420
      Width           =   2535
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   2625
      OleObjectBlob   =   "FILEORIG.frx":0000
      Top             =   0
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   225
      Index           =   0
      Left            =   0
      OleObjectBlob   =   "FILEORIG.frx":0234
      TabIndex        =   5
      Top             =   0
      Visible         =   0   'False
      Width           =   1485
   End
   Begin VB.Label Label1 
      Height          =   225
      Left            =   105
      TabIndex        =   3
      Top             =   105
      Visible         =   0   'False
      Width           =   1065
   End
End
Attribute VB_Name = "FILEORIG"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
   '
   Label1.Caption = Drive1.Drive
   Hide
   
   '
End Sub

Private Sub Command2_Click()
   '
   Label1.Caption = ""
   Unload Me
   '
End Sub

Private Sub Form_Load()
  '
  UTILIZA_SKIN% = 1
  Call APLICACIONSKINS(Me)
  '03/04/07 (OT-07-0147)
  On Error Resume Next
  Drive1.Drive = "A:"
  On Error GoTo 0
  '03/04/07 (FIN)
  '
End Sub
