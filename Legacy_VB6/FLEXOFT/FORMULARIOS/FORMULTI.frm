VERSION 5.00
Begin VB.Form FORMULTI 
   BackColor       =   &H0080FFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "FLEXOFT - Multiempresa"
   ClientHeight    =   1965
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   4680
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form5"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1965
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   540
      Left            =   1155
      TabIndex        =   0
      Top             =   525
      Width           =   2010
   End
End
Attribute VB_Name = "FORMULTI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    FORMULTI.Hide
End Sub
