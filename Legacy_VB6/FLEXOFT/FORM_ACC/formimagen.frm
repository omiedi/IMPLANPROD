VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   8220
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   8220
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture16 
      Height          =   8300
      Left            =   0
      ScaleHeight     =   8235
      ScaleWidth      =   1035
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   1100
      Begin VB.Image Image3 
         Height          =   8400
         Left            =   0
         Picture         =   "formimagen.frx":0000
         Stretch         =   -1  'True
         Top             =   240
         Width           =   1095
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
  If CONTROL("NEWASPEC", "CAMBIOAS") = "SI" Then
      Me.BackColor = &H8000000F
      Picture2.Visible = True
      Frame1.BackColor = &H8000000F
      Frame2.BackColor = &H8000000F
      Frame3.BackColor = &H8000000F
      Frame4.BackColor = &H8000000F
      Frame5.BackColor = &H8000000F

      Frame6.BackColor = &H8000000F
      Frame9.BackColor = &H8000000F
      Frame11.BackColor = &H8000000F
      For I% = 0 To 9
         Option1(I%).BackColor = &H8000000F
      Next I%
      Check1.BackColor = &H8000000F
      Option2.BackColor = &H8000000F
      Option3.BackColor = &H8000000F
      Option4.BackColor = &H8000000F
      Option5.BackColor = &H8000000F
      Option6.BackColor = &H8000000F
      Option7.BackColor = &H8000000F
      mnunewaspecto.Checked = True
      
  End If
 'celeste &HFCD7B6
 'gris &H8000000F
End Sub
