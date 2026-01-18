VERSION 5.00
Begin VB.Form OPKARDEX 
   BackColor       =   &H0080FF80&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Fichas de Movimientos de Stock"
   ClientHeight    =   2385
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   5730
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2385
   ScaleWidth      =   5730
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Opciones"
      Height          =   1905
      Left            =   210
      TabIndex        =   2
      Top             =   210
      Width           =   3690
      Begin VB.OptionButton Option3 
         BackColor       =   &H00C0C0FF&
         Caption         =   "Terminar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   420
         TabIndex        =   5
         Top             =   1260
         Width           =   2955
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00C0C0FF&
         Caption         =   "Ficha por Pieza y Deposito"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   420
         TabIndex        =   4
         Top             =   840
         Width           =   2745
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0C0FF&
         Caption         =   "Ficha General por Pieza"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   420
         TabIndex        =   3
         Top             =   420
         Width           =   2430
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Start"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   4620
      Picture         =   "OPKARDEX.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1260
      Width           =   855
   End
   Begin VB.CommandButton Command1 
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
      Height          =   855
      Left            =   4620
      Picture         =   "OPKARDEX.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   210
      Width           =   855
   End
End
Attribute VB_Name = "OPKARDEX"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    OPKARDEX.Hide
End Sub

Private Sub Command2_Click()
    OPCI% = 1
    If OPKARDEX.Option2.Value = True Then OPCI% = 2
    If OPKARDEX.Option3.Value = True Then OPCI% = 3
    If OPCI% = 1 Then
        Call CONECRUN("*KARDEX", "")
      ElseIf OPCI% = 2 Then
        Call CONECRUN("*KARDEXD", "")
      Else
        OPKARDEX.Hide
    End If
    '
End Sub

Private Sub Form_Load()
   OPKARDEX.Option1.Value = True
End Sub

Private Sub Option1_Click()

End Sub
