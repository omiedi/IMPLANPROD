VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form STMOVI09 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00CDDADA&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Detalle de Estado de  Movimientos"
   ClientHeight    =   4080
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   8220
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
   ScaleHeight     =   4080
   ScaleWidth      =   8220
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Picture999 
      Height          =   330
      Index           =   0
      Left            =   1470
      ScaleHeight     =   270
      ScaleWidth      =   1110
      TabIndex        =   12
      Top             =   525
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Modificar Recepción"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   4320
      Picture         =   "STMOVI09.frx":0000
      TabIndex        =   8
      ToolTipText     =   "Calificación de Proveedores"
      Top             =   3600
      Width           =   2445
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Continuar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   480
      Picture         =   "STMOVI09.frx":030A
      TabIndex        =   7
      ToolTipText     =   "Calificación de Proveedores"
      Top             =   3600
      Width           =   2445
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00CDDADA&
      Caption         =   "DETALLE DE ESTADO DE MOVIMIENTOS"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7100
      Begin VB.Frame Frame1 
         Caption         =   "Frame9"
         Height          =   2780
         Left            =   6690
         TabIndex        =   11
         Top             =   480
         Width           =   15
      End
      Begin VB.Frame Frame9 
         Caption         =   "Frame9"
         Height          =   2790
         Left            =   2040
         TabIndex        =   10
         Top             =   460
         Width           =   15
      End
      Begin VB.ListBox List1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2505
         ItemData        =   "STMOVI09.frx":0614
         Left            =   120
         List            =   "STMOVI09.frx":061B
         TabIndex        =   3
         Top             =   720
         Width           =   6825
      End
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   120
         ScaleHeight     =   255
         ScaleWidth      =   6795
         TabIndex        =   1
         Top             =   450
         Width           =   6825
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Código         Descripción                 "
            BeginProperty Font 
               Name            =   "Fixedsys"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   120
            TabIndex        =   2
            Top             =   0
            Width           =   9795
         End
      End
      Begin VB.Line Line1 
         X1              =   210
         X2              =   5460
         Y1              =   2370
         Y2              =   2370
      End
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      ForeColor       =   &H80000008&
      Height          =   4620
      Left            =   7320
      ScaleHeight     =   4590
      ScaleWidth      =   1125
      TabIndex        =   4
      Top             =   0
      Width           =   1155
      Begin VB.CommandButton command1 
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
         Height          =   690
         Left            =   120
         Picture         =   "STMOVI09.frx":0671
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Salir"
         Top             =   120
         Width           =   650
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
         Height          =   690
         Left            =   120
         Picture         =   "STMOVI09.frx":07BB
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Ayuda FLEXOFT en Línea"
         Top             =   960
         Width           =   650
      End
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   2625
      OleObjectBlob   =   "STMOVI09.frx":0AC5
      Top             =   0
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   225
      Index           =   0
      Left            =   0
      OleObjectBlob   =   "STMOVI09.frx":0CF9
      TabIndex        =   13
      Top             =   0
      Visible         =   0   'False
      Width           =   1485
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      Height          =   495
      Left            =   2640
      TabIndex        =   9
      Top             =   3720
      Visible         =   0   'False
      Width           =   1215
   End
End
Attribute VB_Name = "STMOVI09"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
    Call CIERRARCH("*.*")
    Hide
End Sub
'


Private Sub Command2_Click()
   Command1_Click
End Sub

Private Sub Command3_Click()
   Label2 = "MODIFICAR"
   Command1_Click
End Sub

Private Sub Form_Load()
  '
  UTILIZA_SKIN% = 1
  Call APLICACIONSKINS(Me)
  '
End Sub
