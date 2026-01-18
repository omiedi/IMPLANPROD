VERSION 5.00
Begin VB.Form HABILAP 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Habilitación / Licenciamiento de Aplicaciones"
   ClientHeight    =   1650
   ClientLeft      =   1755
   ClientTop       =   1710
   ClientWidth     =   7515
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1650
   ScaleWidth      =   7515
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "O.K."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   6600
      Picture         =   "HABILAP.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Apruena y Guarda Controlador de Licencia en Ruta Activa."
      Top             =   840
      Width           =   915
   End
   Begin VB.Frame Frame1 
      Height          =   1455
      Left            =   120
      TabIndex        =   5
      Top             =   65
      Width           =   6375
      Begin VB.TextBox Text2 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   4485
         TabIndex        =   2
         Top             =   465
         Width           =   1620
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1245
         TabIndex        =   1
         Top             =   480
         Width           =   1620
      End
      Begin VB.TextBox Text3 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1245
         TabIndex        =   3
         Top             =   945
         Width           =   4860
      End
      Begin VB.Label lblServer 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Aplicacion:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   180
         TabIndex        =   8
         Top             =   510
         Width           =   960
      End
      Begin VB.Label lblFromName 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nro de Serie:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   3255
         TabIndex        =   7
         Top             =   525
         Width           =   1140
      End
      Begin VB.Label lblFrom 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Descripcion:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   45
         TabIndex        =   6
         Top             =   1005
         Width           =   1080
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00004080&
      Height          =   1080
      Left            =   6600
      ScaleHeight     =   1020
      ScaleWidth      =   1965
      TabIndex        =   0
      Top             =   -360
      Width           =   2025
      Begin VB.PictureBox MARCOBARRA 
         BackColor       =   &H000000FF&
         Height          =   135
         Left            =   120
         ScaleHeight     =   75
         ScaleWidth      =   585
         TabIndex        =   9
         Top             =   390
         Width           =   650
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H00FF0000&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   0
            TabIndex        =   10
            Top             =   105
            Width           =   12000
         End
      End
      Begin VB.Image Image2 
         Height          =   390
         Left            =   -300
         Picture         =   "HABILAP.frx":030A
         Top             =   620
         Width           =   1515
      End
   End
End
Attribute VB_Name = "HABILAP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command2_Click()
    PROCE$ = TRIM$(UCase$(Text1))
    SERNU$ = TRIM$(UCase$(Text2))
    DAPLI$ = Text3
    '
    If PROCE$ <> "" Then
      If SERNU$ <> "" Then
        LUAPP$ = App.Path + "\"
        TTXX$ = ENCRIPTA$(AJUSTI$(SERNU$, 16) + AJUSTI$(PROCE$, 110) + MKI$(32767))
        Call SAVEFILE(LUAPP$ + PROCE$ + ".HBL", TTXX$)
      End If
    End If
    '
End Sub
'
Private Sub Form_Load()
      If CONTROL("NEWASPEC", "CAMBIOAS") = "SI" Then
       Me.BackColor = &HFCD7B6
       Me.Frame1.BackColor = &HFCD7B6
       Me.Image2.Visible = False
    End If

End Sub
