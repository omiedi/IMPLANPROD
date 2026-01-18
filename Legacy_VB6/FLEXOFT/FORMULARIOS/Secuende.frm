VERSION 5.00
Begin VB.Form SECUENDE 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Secuencia de Demostración"
   ClientHeight    =   3810
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   8685
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3810
   ScaleWidth      =   8685
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox List1 
      BackColor       =   &H00D0FFFF&
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3885
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8730
   End
End
Attribute VB_Name = "SECUENDE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub List1_DblClick()
   Hide
End Sub
