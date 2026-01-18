VERSION 5.00
Begin VB.Form PRUSERVOL 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Pruseri"
   ClientHeight    =   3225
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3225
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Fin"
      Height          =   330
      Left            =   3465
      TabIndex        =   3
      Top             =   1260
      Width           =   915
   End
   Begin VB.CommandButton Command1 
      Caption         =   "O.K"
      Height          =   330
      Left            =   3465
      TabIndex        =   2
      Top             =   630
      Width           =   960
   End
   Begin VB.TextBox Text2 
      Height          =   330
      Left            =   630
      TabIndex        =   1
      Text            =   " "
      Top             =   1260
      Width           =   1905
   End
   Begin VB.TextBox Text1 
      Height          =   330
      Left            =   2100
      TabIndex        =   0
      Text            =   " "
      Top             =   630
      Width           =   435
   End
End
Attribute VB_Name = "PRUSERVOL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    xxx% = Asc(Text1.Text) - 64
    Text2.Text = SERIVOL$("C:")
End Sub
