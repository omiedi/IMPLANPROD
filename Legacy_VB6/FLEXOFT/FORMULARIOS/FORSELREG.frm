VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form FORSELREG 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00CDDADA&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Reglamentos de Compra"
   ClientHeight    =   4335
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   8070
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4335
   ScaleWidth      =   8070
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture999 
      Height          =   540
      Index           =   0
      Left            =   2280
      ScaleHeight     =   480
      ScaleWidth      =   1110
      TabIndex        =   9
      Top             =   420
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4575
      Left            =   7200
      ScaleHeight     =   4575
      ScaleWidth      =   1095
      TabIndex        =   3
      Top             =   0
      Width           =   1095
      Begin VB.CommandButton Command1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   650
         Left            =   150
         Picture         =   "FORSELREG.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Refresca - Invierte Selección"
         Top             =   1480
         Width           =   650
      End
      Begin VB.CommandButton BOTREC 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   650
         Left            =   150
         Picture         =   "FORSELREG.frx":014A
         Style           =   1  'Graphical
         TabIndex        =   0
         ToolTipText     =   "Aprueba la Pre-Selección"
         Top             =   800
         Width           =   650
      End
      Begin VB.CommandButton BOTEXIT 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   650
         Left            =   150
         Picture         =   "FORSELREG.frx":0454
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Cancela la Pre-Selección"
         Top             =   120
         Width           =   650
      End
   End
   Begin VB.ListBox LIMATSEL 
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
      Height          =   3600
      Left            =   120
      Style           =   1  'Checkbox
      TabIndex        =   1
      Top             =   550
      Width           =   6945
   End
   Begin VB.ListBox List1 
      Height          =   645
      Left            =   2880
      Sorted          =   -1  'True
      TabIndex        =   2
      Top             =   7800
      Visible         =   0   'False
      Width           =   1065
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   0
      OleObjectBlob   =   "FORSELREG.frx":075E
      Top             =   1680
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   435
      Index           =   0
      Left            =   60
      OleObjectBlob   =   "FORSELREG.frx":0992
      TabIndex        =   10
      Top             =   0
      Visible         =   0   'False
      Width           =   1485
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   5880
      TabIndex        =   8
      Top             =   180
      Width           =   1125
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Documentos Anexos para la Orden de Compra Nº: "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   1440
      TabIndex        =   7
      Top             =   180
      Width           =   4395
   End
   Begin VB.Label Label1 
      Height          =   330
      Left            =   360
      TabIndex        =   6
      Top             =   8040
      Width           =   855
   End
End
Attribute VB_Name = "FORSELREG"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub BOTEXIT_Click()
   Unload Me
End Sub
'
Private Sub BOTREC_Click()
   Label1 = "OK"
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

Private Sub Form_Load()
    UTILIZA_SKIN% = 1
    Call APLICACIONSKINS(Me)
End Sub
