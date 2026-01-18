VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form ORCOPEND07 
   BackColor       =   &H00CDDADA&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   7140
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12525
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7140
   ScaleWidth      =   12525
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture999 
      Height          =   435
      Index           =   1
      Left            =   120
      ScaleHeight     =   375
      ScaleWidth      =   1110
      TabIndex        =   17
      Top             =   720
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.Frame Frame7 
      Caption         =   "Frame14"
      Height          =   6135
      Left            =   11120
      TabIndex        =   15
      Top             =   0
      Width           =   15
   End
   Begin VB.Frame Frame6 
      Caption         =   "Frame14"
      Height          =   6030
      Left            =   10160
      TabIndex        =   14
      Top             =   0
      Width           =   15
   End
   Begin VB.Frame Frame5 
      Caption         =   "Frame14"
      Height          =   6030
      Left            =   9220
      TabIndex        =   13
      Top             =   105
      Width           =   15
   End
   Begin VB.Frame Frame4 
      Caption         =   "Frame14"
      Height          =   6030
      Left            =   8350
      TabIndex        =   12
      Top             =   105
      Width           =   15
   End
   Begin VB.Frame Frame3 
      Caption         =   "Frame14"
      Height          =   6030
      Left            =   5710
      TabIndex        =   11
      Top             =   105
      Width           =   15
   End
   Begin VB.Frame Frame2 
      Caption         =   "Frame14"
      Height          =   6135
      Left            =   3925
      TabIndex        =   10
      Top             =   0
      Width           =   15
   End
   Begin VB.Frame Frame1 
      Caption         =   "Frame14"
      Height          =   6030
      Left            =   2040
      TabIndex        =   9
      Top             =   105
      Width           =   15
   End
   Begin VB.Frame Frame14 
      Caption         =   "Frame14"
      Height          =   5950
      Left            =   1330
      TabIndex        =   8
      Top             =   160
      Width           =   15
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      ForeColor       =   &H80000008&
      Height          =   825
      Left            =   120
      ScaleHeight     =   795
      ScaleWidth      =   12300
      TabIndex        =   4
      Top             =   6240
      Width           =   12330
      Begin VB.PictureBox MARCOBARRA 
         Appearance      =   0  'Flat
         BackColor       =   &H002B9973&
         ForeColor       =   &H80000008&
         Height          =   200
         Left            =   8950
         ScaleHeight     =   165
         ScaleWidth      =   1695
         TabIndex        =   19
         Top             =   480
         Width           =   1725
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H0048CC9D&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   0
            TabIndex        =   20
            Top             =   0
            Width           =   12000
         End
      End
      Begin VB.CommandButton Command4 
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   650
         Left            =   8150
         Picture         =   "ORCOPEND07.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   16
         ToolTipText     =   "Ver por Pantalla Orden de Compra"
         Top             =   105
         Width           =   650
      End
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
         Left            =   105
         Picture         =   "ORCOPEND07.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Cancela la Pre-Selección"
         Top             =   105
         Width           =   650
      End
      Begin VB.CommandButton Command3 
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
         Left            =   11550
         Picture         =   "ORCOPEND07.frx":0614
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Aprueba la Selección"
         Top             =   105
         Width           =   650
      End
      Begin VB.CommandButton Command2 
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
         Left            =   10815
         Picture         =   "ORCOPEND07.frx":091E
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Refresca - Invierte Selección"
         Top             =   105
         Width           =   650
      End
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5700
      ItemData        =   "ORCOPEND07.frx":0A68
      Left            =   120
      List            =   "ORCOPEND07.frx":0A6A
      Style           =   1  'Checkbox
      TabIndex        =   3
      Top             =   420
      Width           =   12285
   End
   Begin VB.PictureBox Picture999 
      Height          =   540
      Index           =   0
      Left            =   2280
      ScaleHeight     =   480
      ScaleWidth      =   1110
      TabIndex        =   1
      Top             =   420
      Visible         =   0   'False
      Width           =   1170
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   0
      OleObjectBlob   =   "ORCOPEND07.frx":0A6C
      Top             =   1680
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   225
      Index           =   0
      Left            =   2310
      OleObjectBlob   =   "ORCOPEND07.frx":0CA0
      TabIndex        =   2
      Top             =   2205
      Visible         =   0   'False
      Width           =   1905
   End
   Begin ACTIVESKINLibCtl.Skin Skin2 
      Left            =   -120
      OleObjectBlob   =   "ORCOPEND07.frx":0D1A
      Top             =   1320
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   435
      Index           =   1
      Left            =   0
      OleObjectBlob   =   "ORCOPEND07.frx":0F4E
      TabIndex        =   18
      Top             =   315
      Visible         =   0   'False
      Width           =   1590
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   $"ORCOPEND07.frx":0FC8
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   105
      Width           =   12285
   End
End
Attribute VB_Name = "ORCOPEND07"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub BOTREC_Click()

End Sub

Private Sub Command1_Click()
    '
    Call CIERRARCH("*.*")
    List1.Clear
    Unload Me
    '
End Sub
Private Sub Command3_Click()
   Hide
End Sub

Sub Command2_Click()
   Command2.Enabled = False
   TPI = List1.TopIndex
   For U& = 1 To List1.ListCount
     If List1.Selected(U& - 1) = True Then
         List1.Selected(U& - 1) = False
       Else
         List1.Selected(U& - 1) = True
     End If
   Next U&
   On Error Resume Next
   List1.TopIndex = TPI
   On Error GoTo 0
   DoEvents
   Command2.Enabled = True
End Sub

Private Sub Command4_Click()
   Command4.Enabled = False
   Call OPLISCO07.Command22_Click
   Command4.Enabled = True
End Sub

Private Sub Form_Load()
    UTILIZA_SKIN% = 1
    Call APLICACIONSKINS(Me)
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   Call Command1_Click
End Sub

