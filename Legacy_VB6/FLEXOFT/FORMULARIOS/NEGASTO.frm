VERSION 5.00
Begin VB.Form NEGASTO 
   BackColor       =   &H00C0FFFF&
   Caption         =   "Control de Negativos"
   ClientHeight    =   3390
   ClientLeft      =   60
   ClientTop       =   315
   ClientWidth     =   5895
   LinkTopic       =   "Form1"
   ScaleHeight     =   3390
   ScaleWidth      =   5895
   StartUpPosition =   3  'Windows Default
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
      Left            =   4725
      Picture         =   "NEGASTO.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   525
      Width           =   855
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
      Left            =   4725
      Picture         =   "NEGASTO.frx":014A
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1575
      Width           =   855
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H0080C0FF&
      Caption         =   "Opciones"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1905
      Left            =   315
      TabIndex        =   0
      Top             =   525
      Width           =   3690
      Begin VB.OptionButton Option1 
         BackColor       =   &H0080C0FF&
         Caption         =   "Consulta por Pantalla"
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
      Begin VB.OptionButton Option2 
         BackColor       =   &H0080C0FF&
         Caption         =   "Listado por Impresora"
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
         TabIndex        =   2
         Top             =   840
         Width           =   2745
      End
      Begin VB.OptionButton Option3 
         BackColor       =   &H0080C0FF&
         Caption         =   "Revisar existencias"
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
         TabIndex        =   1
         Top             =   1260
         Width           =   2955
      End
   End
End
Attribute VB_Name = "NEGASTO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    NEGASTO.Hide
End Sub

Private Sub Command2_Click()
    If NEGASTO.Option1.Value = True Then
        Call CONECRUN("*NEGASTS", "")
      ElseIf NEGASTO.Option2.Value = True Then
        Call CONECRUN("*NEGASTL", "")
      Else
        Call NEGASTO2
    End If
End Sub

Private Sub Form_Load()
    Screen.MousePointer = 11
    Screen.MousePointer = 1
    NEGASTO.Option3.Value = True
End Sub
