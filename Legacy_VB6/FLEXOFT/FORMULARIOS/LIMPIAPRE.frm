VERSION 5.00
Begin VB.Form LIMPIAPRE 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0E0FF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Depuración de Listas de Precios"
   ClientHeight    =   2565
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   4575
   Icon            =   "LIMPIAPRE.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2565
   ScaleWidth      =   4575
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command5 
      Caption         =   "O.K."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   740
      Left            =   3675
      Picture         =   "LIMPIAPRE.frx":014A
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Aprueba y Ejecuta"
      Top             =   1155
      Width           =   750
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   740
      Left            =   3675
      Picture         =   "LIMPIAPRE.frx":0454
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Cancela la Aplicación"
      Top             =   315
      Width           =   750
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Acciones"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2220
      Left            =   210
      TabIndex        =   0
      Top             =   210
      Width           =   3375
      Begin VB.CheckBox Check1 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Eliminar Items Duplicados"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Index           =   3
         Left            =   480
         TabIndex        =   6
         Top             =   1080
         Width           =   2745
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Control Precios vs. Costos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Index           =   2
         Left            =   480
         TabIndex        =   5
         Top             =   1680
         Width           =   2775
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Eliminar Artículos sin Stock"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Index           =   1
         Left            =   480
         TabIndex        =   2
         Top             =   720
         Width           =   2745
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Eliminar Items de Baja"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Index           =   0
         Left            =   480
         TabIndex        =   1
         Top             =   360
         Width           =   2535
      End
   End
   Begin VB.Image Image2 
      Height          =   390
      Left            =   3360
      Picture         =   "LIMPIAPRE.frx":059E
      Top             =   2020
      Width           =   1515
   End
End
Attribute VB_Name = "LIMPIAPRE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Check1_Click(Index As Integer)
  '
  Select Case Index
    Case 0
        If Check1(0).Value = 1 Then
           Check1(2).Value = 0
           Check1(3).Value = 0
        End If
    Case 1
        If Check1(1).Value = 1 Then
           Check1(2).Value = 0
           Check1(3).Value = 0
        End If
    Case 2
        If Check1(2).Value = 1 Then
           Check1(0).Value = 0
           Check1(1).Value = 0
           Check1(3).Value = 0
        End If
    Case 3
        If Check1(3).Value = 1 Then
           Check1(0).Value = 0
           Check1(1).Value = 0
           Check1(2).Value = 0
        End If
  End Select
  '
End Sub

Private Sub Command1_Click()
   Command1.Enabled = False
   Check1(0).Value = Unchecked
   Check1(1).Value = Unchecked
   Check1(2).Value = Unchecked
   Check1(3).Value = Unchecked
   Command1.Enabled = True
   Hide
End Sub

Private Sub Command5_Click()
   Hide
End Sub

Private Sub Form_Load()
   Check1(0).Value = Unchecked
   Check1(1).Value = Unchecked
End Sub
