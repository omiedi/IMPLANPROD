VERSION 5.00
Begin VB.Form OPINVE 
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Opciones de Inventarios"
   ClientHeight    =   3105
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   6285
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3105
   ScaleWidth      =   6285
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      BackColor       =   &H0080FF80&
      Caption         =   "Selección"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1485
      Left            =   735
      TabIndex        =   3
      Top             =   1155
      Width           =   3060
      Begin VB.OptionButton Option1 
         Caption         =   "Option1"
         Height          =   225
         Left            =   525
         TabIndex        =   6
         Top             =   420
         Value           =   -1  'True
         Width           =   225
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Option2"
         Height          =   195
         Left            =   525
         TabIndex        =   5
         Top             =   735
         Width           =   225
      End
      Begin VB.OptionButton Option3 
         Caption         =   "Option1"
         Height          =   225
         Left            =   525
         TabIndex        =   4
         Top             =   1050
         Width           =   225
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Inventario General"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   1050
         TabIndex        =   9
         Top             =   420
         Width           =   1590
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Inventario Analítico"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   1050
         TabIndex        =   8
         Top             =   735
         Width           =   1800
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Parcial por Depósito"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   1050
         TabIndex        =   7
         Top             =   1050
         Width           =   1905
      End
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   2730
      TabIndex        =   2
      Text            =   " "
      Top             =   420
      Width           =   1065
   End
   Begin VB.CommandButton command1 
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
      Left            =   4830
      Picture         =   "OPINVE.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   420
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
      Left            =   4830
      Picture         =   "OPINVE.frx":014A
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1785
      Width           =   855
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Fecha de Inventario:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   630
      TabIndex        =   10
      Top             =   525
      Width           =   1905
   End
End
Attribute VB_Name = "OPINVE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    OPINVE.Hide
End Sub

Private Sub Command2_Click()
    If OPINVE.Option1.Value = True Then
        Shell LUBAS$ + "WINTABU INVEGENS", 1
      ElseIf OPINVE.Option2.Value = True Then
        Shell LUBAS$ + "WINTABU INVEDEPS", 1
      ElseIf OPINVE.Option3.Value = True Then
        Shell LUBAS$ + "WINTABU INVANALS", 1
    End If
End Sub

Private Sub Form_Load()
    OPINVE.Text1.Text = FFS$
    OPINVE.Option1.Value = True
End Sub

Private Sub Text1_GotFocus()
    OPINVE.Command2.SetFocus
End Sub
