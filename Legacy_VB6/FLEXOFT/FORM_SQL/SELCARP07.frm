VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form SELCARP07 
   BackColor       =   &H00DAE4EB&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Carpeta de Exportación"
   ClientHeight    =   2700
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3420
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2700
   ScaleWidth      =   3420
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture999 
      Height          =   540
      Index           =   0
      Left            =   2280
      ScaleHeight     =   480
      ScaleWidth      =   1110
      TabIndex        =   4
      Top             =   1320
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.CommandButton Command1 
      Height          =   645
      Left            =   2640
      Picture         =   "SELCARP07.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1920
      Width           =   645
   End
   Begin VB.DirListBox Dir1 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2010
      Left            =   120
      TabIndex        =   1
      Top             =   525
      Width           =   2430
   End
   Begin VB.DriveListBox Drive1 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   120
      TabIndex        =   0
      Top             =   105
      Width           =   2430
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   750
      Index           =   0
      Left            =   0
      OleObjectBlob   =   "SELCARP07.frx":030A
      TabIndex        =   5
      Top             =   0
      Visible         =   0   'False
      Width           =   1590
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   3600
      OleObjectBlob   =   "SELCARP07.frx":0384
      Top             =   1200
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   120
      Left            =   525
      TabIndex        =   3
      Top             =   2730
      Visible         =   0   'False
      Width           =   2640
   End
End
Attribute VB_Name = "SELCARP07"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
   Label1 = Dir1.Path
   Hide
End Sub

Private Sub Dir1_Change()
   On Error Resume Next
   RUTASAL$ = Dir1.Path
   WLITAB04.Label11.Caption = RUTASAL$
   WLITAB07.Label11.Caption = RUTASAL$
   On Error GoTo 0
End Sub

Private Sub Drive1_Change()
   On Error Resume Next
   Dir1.Path = Drive1.Drive
   On Error GoTo 0
End Sub

Private Sub Form_Load()
      Call APLICACIONSKINS(Me)
End Sub
