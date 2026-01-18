VERSION 5.00
Begin VB.Form SITCOB09 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Sit.Cobertura"
   ClientHeight    =   2730
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   2385
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2730
   ScaleWidth      =   2385
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00E0E0E0&
      Height          =   2680
      Left            =   20
      ScaleHeight     =   2625
      ScaleWidth      =   2265
      TabIndex        =   0
      Top             =   0
      Width           =   2325
      Begin VB.Shape Shape1 
         FillColor       =   &H00C00000&
         FillStyle       =   0  'Solid
         Height          =   645
         Left            =   105
         Top             =   1785
         Width           =   950
      End
      Begin VB.Shape Shape2 
         FillColor       =   &H00808000&
         FillStyle       =   0  'Solid
         Height          =   540
         Left            =   105
         Top             =   1260
         Width           =   950
      End
      Begin VB.Shape Shape3 
         FillColor       =   &H0000C000&
         FillStyle       =   0  'Solid
         Height          =   750
         Left            =   105
         Top             =   525
         Width           =   950
      End
      Begin VB.Shape Shape4 
         FillColor       =   &H000000C0&
         FillStyle       =   0  'Solid
         Height          =   645
         Left            =   1220
         Top             =   1785
         Width           =   950
      End
      Begin VB.Shape Shape5 
         FillColor       =   &H000080FF&
         FillStyle       =   0  'Solid
         Height          =   645
         Left            =   1220
         Top             =   1155
         Width           =   950
      End
      Begin VB.Shape Shape6 
         FillColor       =   &H0080C0FF&
         FillStyle       =   0  'Solid
         Height          =   645
         Left            =   1220
         Top             =   525
         Width           =   950
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Caption         =   "Cobertura"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   260
         Left            =   0
         TabIndex        =   1
         Top             =   0
         Width           =   2325
      End
      Begin VB.Line Line5 
         X1              =   0
         X2              =   2310
         Y1              =   270
         Y2              =   270
      End
   End
End
Attribute VB_Name = "SITCOB09"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'
Private Sub Form_Load()

End Sub
