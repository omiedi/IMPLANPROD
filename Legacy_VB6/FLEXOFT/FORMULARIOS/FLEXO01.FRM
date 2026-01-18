VERSION 5.00
Begin VB.Form Form2 
   BackColor       =   &H000080FF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Sistema Integrado Flexoft"
   ClientHeight    =   2880
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4935
   DrawMode        =   8  'Xor Pen
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2880
   ScaleWidth      =   4935
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   6000
      MaxLength       =   10
      TabIndex        =   2
      Top             =   3960
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   5040
      Top             =   105
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      AutoRedraw      =   -1  'True
      Height          =   4440
      Left            =   0
      Picture         =   "flexo01.frx":0000
      ScaleHeight     =   4380
      ScaleWidth      =   4875
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   4935
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Sistemas de Fábrica"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   105
         TabIndex        =   1
         Top             =   225
         Visible         =   0   'False
         Width           =   2415
      End
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   0
      Top             =   120
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H000080FF&
      Caption         =   "Ingrese su Clave"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   4080
      TabIndex        =   3
      Top             =   3960
      Visible         =   0   'False
      Width           =   1755
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    Picture1.Visible = True
    Timer1.Enabled = True
    Timer2.Enabled = False
    Text1.Visible = False
    Label2.Visible = False
    '
End Sub

Private Sub Timer1_Timer()
    Static veces
    veces = veces + 1
    '
    If veces >= 15 Then
       Label1.Visible = True
       Timer2.Enabled = True
       Form1.Enabled = True
       Form1.Label1.Visible = True
       Form1.Label2.Visible = True
       Form1.Label3.Visible = True
     
    End If
    '
End Sub

Private Sub Timer2_Timer()
    termina% = 1
    If Form2.Left > Form1.Left + 100 Then
       Form2.Move Form2.Left - 20, Form2.Top
       termina% = 0
    End If
    '
    If Form2.Top > Form1.Top + 380 Then
       Form2.Move Form2.Left, Form2.Top - 14
       termina% = 0
    End If
    '
    If termina% = 1 Then
        'Form1.Command1.Visible = True
        Form1.Enabled = True
        Form1.Frame1.Visible = True
        'FORM3.Show
    End If
End Sub
