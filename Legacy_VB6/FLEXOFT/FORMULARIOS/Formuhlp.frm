VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form FORMUHLP 
   BackColor       =   &H00FFFFC0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Ayuda FLEXOFT en Línea"
   ClientHeight    =   4785
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   7350
   BeginProperty Font 
      Name            =   "Courier New"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4785
   ScaleWidth      =   7350
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture999 
      Height          =   540
      Index           =   0
      Left            =   315
      ScaleHeight     =   480
      ScaleWidth      =   1110
      TabIndex        =   9
      Top             =   0
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.PictureBox MARCOBARRA 
      BackColor       =   &H000000FF&
      Height          =   135
      Left            =   240
      ScaleHeight     =   75
      ScaleWidth      =   2475
      TabIndex        =   7
      Top             =   4440
      Width           =   2535
      Begin VB.Frame BARRATRAZA 
         BackColor       =   &H00FF0000&
         BorderStyle     =   0  'None
         Height          =   500
         Left            =   0
         TabIndex        =   8
         Top             =   0
         Width           =   12000
      End
   End
   Begin VB.ListBox LIAPLI 
      Height          =   1740
      Left            =   2100
      Sorted          =   -1  'True
      TabIndex        =   6
      Top             =   1050
      Visible         =   0   'False
      Width           =   3060
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   540
      Left            =   2940
      TabIndex        =   2
      Top             =   4200
      Width           =   4005
      Begin VB.CommandButton Command3 
         Caption         =   "Print"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   105
         TabIndex        =   5
         Top             =   105
         Width           =   1065
      End
      Begin VB.CommandButton Command1 
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
         Height          =   345
         Left            =   2625
         TabIndex        =   4
         Top             =   105
         Width           =   1065
      End
      Begin VB.CommandButton Command2 
         Caption         =   "More"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1365
         TabIndex        =   3
         Top             =   105
         Width           =   1065
      End
   End
   Begin VB.ListBox LISTA 
      BackColor       =   &H00DFFFFF&
      Height          =   3900
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6945
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   2625
      OleObjectBlob   =   "Formuhlp.frx":0000
      Top             =   0
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   225
      Index           =   0
      Left            =   0
      OleObjectBlob   =   "Formuhlp.frx":0234
      TabIndex        =   10
      Top             =   0
      Visible         =   0   'False
      Width           =   1485
   End
   Begin VB.Label TEPRUEBA 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Height          =   240
      Left            =   0
      TabIndex        =   1
      Top             =   4515
      Visible         =   0   'False
      Width           =   120
   End
End
Attribute VB_Name = "FORMUHLP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    FORMUHLP.Hide
    Call FRESHALL
End Sub

Private Sub Command2_Click()
    LIAPLI.Visible = True
End Sub

Private Sub Command3_Click()
    Call OPNOIN("")
End Sub

Private Sub Form_Activate()
  Call CARGAHELP
End Sub

Private Sub Form_Load()
   '
   Call TRANSLABELS(Name)
   Call APLICACIONSKINS(Me)
   '
End Sub

Private Sub LIAPLI_DblClick()
    LIAPLI.Visible = False
    APLICA$ = Left$(LIAPLI.TEXT, 8)
    Call CARGAHELP
End Sub

