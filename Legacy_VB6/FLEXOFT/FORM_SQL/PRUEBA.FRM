VERSION 5.00
Begin VB.Form PRUEBA 
   BackColor       =   &H00E0E0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Control de Impresión"
   ClientHeight    =   6360
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8490
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6360
   ScaleWidth      =   8490
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   375
      Left            =   6120
      TabIndex        =   6
      Top             =   360
      Width           =   1335
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   4440
      TabIndex        =   5
      Text            =   "50"
      Top             =   2640
      Width           =   615
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   3600
      TabIndex        =   4
      Text            =   "8"
      Top             =   2640
      Width           =   615
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   1320
      TabIndex        =   3
      Text            =   "Arial"
      Top             =   2640
      Width           =   1935
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      FontTransparent =   0   'False
      Height          =   2175
      Left            =   240
      ScaleHeight     =   2115
      ScaleWidth      =   6435
      TabIndex        =   2
      Top             =   3360
      Width           =   6495
   End
   Begin VB.TextBox Text1 
      Height          =   2175
      Left            =   240
      MultiLine       =   -1  'True
      TabIndex        =   1
      Text            =   "PRUEBA.frx":0000
      Top             =   240
      Width           =   5295
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
      Height          =   630
      Left            =   5880
      Picture         =   "PRUEBA.frx":0006
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Impresion de Prueba de Remito"
      Top             =   1680
      Width           =   780
   End
End
Attribute VB_Name = "PRUEBA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Picture1.Print "PEPE"
End Sub

Private Sub Command14_Click()
   TTXX$ = Text1
   ANCHOMM = XVALO(Text4)
   Call FRATEXFONT(TTXX, ANCHOMM, Text2, Text3, 0)
   '
   Picture1.Cls
   Picture1.FontName = Text2
   Picture1.FontSize = Text3
   Picture1.FontBold = False
'   Picture1.CurrentX = 500
'   Picture1.CurrentY = 500
'   Picture1.Print "PEPE"   ' RETEX$(1)
   For KKI% = 1 To CARETEX%
     Picture1.Print RETEX$(KKI%)
   Next KKI%
'   Picture1.Refresh
End Sub
