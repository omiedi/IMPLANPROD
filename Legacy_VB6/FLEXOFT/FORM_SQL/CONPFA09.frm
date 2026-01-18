VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form CONPFA09 
   BackColor       =   &H00E0E0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Control de Impresión"
   ClientHeight    =   1296
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   4008
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1296
   ScaleWidth      =   4008
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture999 
      Height          =   435
      Index           =   0
      Left            =   105
      ScaleHeight     =   384
      ScaleWidth      =   1128
      TabIndex        =   5
      Top             =   735
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.CommandButton Command14 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   875
      Left            =   3360
      Picture         =   "CONPFA09.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Aprueba, Graba e Imprime "
      Top             =   220
      Width           =   540
   End
   Begin VB.Frame Frame10 
      BackColor       =   &H00E0E0C0&
      Caption         =   "Impresión"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   960
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3075
      Begin VB.CheckBox Check8 
         BackColor       =   &H00E0E0C0&
         Caption         =   "Impresion Fantasma"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   510
         Left            =   1560
         TabIndex        =   2
         Top             =   315
         Width           =   1400
      End
      Begin VB.CheckBox Check7 
         BackColor       =   &H00E0E0C0&
         Caption         =   "Imprime Informe"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
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
         Width           =   1140
      End
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   420
      OleObjectBlob   =   "CONPFA09.frx":030A
      Top             =   0
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   435
      Index           =   0
      Left            =   0
      OleObjectBlob   =   "CONPFA09.frx":053E
      TabIndex        =   6
      Top             =   315
      Visible         =   0   'False
      Width           =   1590
   End
   Begin VB.Label Label1 
      Height          =   435
      Left            =   105
      TabIndex        =   3
      Top             =   1890
      Width           =   1380
   End
End
Attribute VB_Name = "CONPFA09"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command14_Click()
    Label1.Caption = "OK"
    Hide
End Sub

Private Sub Form_Load()
    Check7.Value = 1
    Check8.Value = 0
    '
    Call APLICACIONSKINS(Me)

End Sub
