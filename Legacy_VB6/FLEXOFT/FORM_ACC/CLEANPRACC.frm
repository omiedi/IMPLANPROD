VERSION 5.00
Begin VB.Form CLEANPRACC 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00D0E0FF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Depuración de Listas de Precios"
   ClientHeight    =   1905
   ClientLeft      =   45
   ClientTop       =   585
   ClientWidth     =   4725
   Icon            =   "CLEANPRACC.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1905
   ScaleWidth      =   4725
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00000080&
      Height          =   2115
      Left            =   3885
      ScaleHeight     =   2055
      ScaleWidth      =   1740
      TabIndex        =   4
      Top             =   0
      Width           =   1800
      Begin VB.PictureBox Picture16 
         Height          =   1900
         Left            =   0
         ScaleHeight     =   1845
         ScaleWidth      =   1020
         TabIndex        =   7
         Top             =   0
         Visible         =   0   'False
         Width           =   1080
         Begin VB.Image Image3 
            Height          =   1900
            Left            =   0
            Picture         =   "CLEANPRACC.frx":014A
            Stretch         =   -1  'True
            Top             =   0
            Width           =   1095
         End
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
         Height          =   690
         Left            =   105
         Picture         =   "CLEANPRACC.frx":24E2
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Cancela la Aplicación"
         Top             =   290
         Width           =   600
      End
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
         Height          =   690
         Left            =   105
         Picture         =   "CLEANPRACC.frx":262C
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Aprueba y Ejecuta"
         Top             =   1050
         Width           =   600
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00D0E0FF&
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
      Height          =   1550
      Left            =   210
      TabIndex        =   0
      Top             =   210
      Width           =   3480
      Begin VB.CheckBox Check1 
         BackColor       =   &H00D0E0FF&
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
         TabIndex        =   3
         Top             =   1080
         Width           =   2775
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00D0E0FF&
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
         BackColor       =   &H00D0E0FF&
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
   Begin VB.Menu archivo 
      Caption         =   "Archivo"
      Begin VB.Menu grabar 
         Caption         =   "Aceptar"
      End
      Begin VB.Menu s 
         Caption         =   "-"
      End
      Begin VB.Menu salir 
         Caption         =   "Salir"
      End
   End
End
Attribute VB_Name = "CLEANPRACC"
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
        End If
    Case 1
        If Check1(1).Value = 1 Then
           Check1(2).Value = 0
        End If
    Case 2
        If Check1(2).Value = 1 Then
           Check1(0).Value = 0
           Check1(1).Value = 0
        End If
  End Select
  '
End Sub

Private Sub Command1_Click()
   Command1.Enabled = False
   Check1(0).Value = Unchecked
   Check1(1).Value = Unchecked
   Command1.Enabled = True
   Hide
End Sub

Private Sub Command5_Click()
   Hide
End Sub

Private Sub Form_Load()
    If CONTROL("NEWASPEC", "CAMBIOAS") = "SI" Then
      Picture16.Visible = True
      Picture16.ZOrder 0
      Me.BackColor = &HFCD7B6
      Me.Frame1.BackColor = &HFCD7B6
      Me.Check1(0).BackColor = &HFCD7B6
      Me.Check1(1).BackColor = &HFCD7B6
      Me.Check1(2).BackColor = &HFCD7B6
      
      ' &H8000000F
    End If

   Check1(0).Value = Unchecked
   Check1(1).Value = Unchecked
End Sub

Private Sub grabar_Click()
  Call Command5_Click
End Sub

Private Sub salir_Click()
  Call Command1_Click
End Sub
