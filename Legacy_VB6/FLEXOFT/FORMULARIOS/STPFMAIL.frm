VERSION 5.00
Begin VB.Form STPFMAIL 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Configuracion de E-Mail por Winsock"
   ClientHeight    =   3540
   ClientLeft      =   1755
   ClientTop       =   1710
   ClientWidth     =   7890
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3540
   ScaleWidth      =   7890
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
      Height          =   1170
      Left            =   7010
      Picture         =   "STPFMAIL.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Guarda Configuración de Correo Electrónico"
      Top             =   2280
      Width           =   800
   End
   Begin VB.Frame Frame1 
      Height          =   3375
      Left            =   120
      TabIndex        =   12
      Top             =   65
      Width           =   6735
      Begin VB.Frame Frame3 
         Caption         =   "Estación"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   3600
         TabIndex        =   19
         Top             =   2640
         Width           =   2895
         Begin VB.OptionButton Option3 
            Caption         =   "Servidor"
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
            Left            =   240
            TabIndex        =   21
            Top             =   280
            Value           =   -1  'True
            Width           =   1095
         End
         Begin VB.OptionButton Option4 
            Caption         =   "Cliente"
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
            Left            =   1680
            TabIndex        =   20
            Top             =   280
            Width           =   1095
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Modo Envio"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   600
         TabIndex        =   18
         Top             =   2640
         Width           =   2775
         Begin VB.OptionButton Option2 
            Caption         =   "Direct"
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
            Left            =   1680
            TabIndex        =   7
            Top             =   280
            Width           =   855
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Preview"
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
            Left            =   240
            TabIndex        =   6
            Top             =   280
            Value           =   -1  'True
            Width           =   1095
         End
      End
      Begin VB.TextBox txtFromName 
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
         Left            =   1600
         TabIndex        =   2
         Top             =   945
         Width           =   4860
      End
      Begin VB.TextBox txtUserName 
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
         Left            =   600
         TabIndex        =   4
         Top             =   2160
         Width           =   3540
      End
      Begin VB.TextBox txtPassword 
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
         IMEMode         =   3  'DISABLE
         Left            =   4560
         TabIndex        =   5
         Top             =   2160
         Width           =   1890
      End
      Begin VB.TextBox txtServer 
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
         Left            =   1600
         TabIndex        =   1
         Top             =   480
         Width           =   4860
      End
      Begin VB.TextBox txtFrom 
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
         Left            =   1600
         TabIndex        =   3
         Top             =   1425
         Width           =   4860
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00C0C0C0&
         X1              =   0
         X2              =   6720
         Y1              =   1840
         Y2              =   1840
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Username:"
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
         Left            =   120
         TabIndex        =   17
         Top             =   1935
         Width           =   975
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Password:"
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
         Left            =   4080
         TabIndex        =   16
         Top             =   1920
         Width           =   1095
      End
      Begin VB.Label lblServer 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "SMTP Server:"
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
         Left            =   300
         TabIndex        =   15
         Top             =   510
         Width           =   1200
      End
      Begin VB.Label lblFromName 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Sender Name:"
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
         Left            =   285
         TabIndex        =   14
         Top             =   1005
         Width           =   1215
      End
      Begin VB.Label lblFrom 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Sender Email:"
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
         Left            =   300
         TabIndex        =   13
         Top             =   1485
         Width           =   1185
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00004080&
      Height          =   2160
      Left            =   6960
      ScaleHeight     =   2100
      ScaleWidth      =   1965
      TabIndex        =   0
      Top             =   0
      Width           =   2025
      Begin VB.PictureBox MARCOBARRA 
         BackColor       =   &H000000FF&
         Height          =   135
         Left            =   120
         ScaleHeight     =   75
         ScaleWidth      =   585
         TabIndex        =   10
         Top             =   1200
         Width           =   650
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H00FF0000&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   0
            TabIndex        =   11
            Top             =   105
            Width           =   12000
         End
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
         Height          =   650
         Left            =   120
         Picture         =   "STPFMAIL.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   120
         Width           =   650
      End
      Begin VB.Image Image2 
         Height          =   390
         Left            =   -300
         Picture         =   "STPFMAIL.frx":0454
         Top             =   1440
         Width           =   1515
      End
   End
End
Attribute VB_Name = "STPFMAIL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Command1_Click()
    '
    Call CIERRARCH("*.*")
    Hide
    '
End Sub

Private Sub Command2_Click()
'
    Command2.Enabled = False
    '
    If txtServer.TEXT = "" Then
        Call MENSERR(24, "Debe Ingresar Nombre del Servidor SMTP")
        GoTo 99
    End If
    
    If txtFromName.TEXT = "" Then
        Call MENSERR(24, "Debe Ingresar Nombre Remitente")
        GoTo 99
    End If
    
    If txtFrom.TEXT = "" Then
        Call MENSERR(24, "Debe Ingresar Mail Remitente")
        GoTo 99
    End If
    
    If txtUserName.TEXT = "" Then
        Call MENSERR(24, "Nombre de Usuario no Valido")
        GoTo 99
    End If
    
    If txtPassword.TEXT = "" Then
        Call MENSERR(24, "Password no Valido")
        GoTo 99
    End If
    
    TTXX$ = AJUSTI$(txtServer, 128)
    TTXX$ = TTXX$ + AJUSTI$(txtFromName, 128)
    TTXX$ = TTXX$ + AJUSTI$(txtFrom, 128)
    TTXX$ = TTXX$ + AJUSTI$(txtUserName, 128)
    TTXX$ = TTXX$ + ENCRIPTA$(AJUSTI$(txtPassword, 32))
    '
    If Option1.Value = True Then
       TTXX$ = TTXX$ + AJUSTI$("PREVIEW", 8)
    Else
        TTXX$ = TTXX$ + AJUSTI$("DIRECT", 8)
    End If
    If Option3.Value = True Then
       TTXX$ = TTXX$ + AJUSTI$("SERVER", 8)
    Else
        TTXX$ = TTXX$ + AJUSTI$("CLIENT", 8)
    End If
    Call SAVEFILE(LUCOM$ + "STPFMAIL.CFG", TTXX$)
    '
    Call MENSERR(24, "Configuración Procesada")
    
99  Command2.Enabled = True
    '
End Sub

Private Sub Form_Load()
    '
    If CONTROL("NEWASPEC", "CAMBIOAS") = "SI" Then
      Me.Image2.Visible = False
      ' &H8000000F
     End If

    TTXX$ = LOADFILE(LUCOM$ + "STPFMAIL.CFG")
    txtServer = TRIM$(Mid$(TTXX$, 1, 128))
    txtFromName = TRIM$(Mid$(TTXX$, 129, 128))
    txtFrom = TRIM$(Mid$(TTXX$, 257, 128))
    txtUserName = TRIM$(Mid$(TTXX$, 385, 128))
    txtPassword = TRIM$(DESENCRI$(Mid$(TTXX$, 513, 32)))
    '
    If TRIM$(Mid$(TTXX$, 545, 8)) = "PREVIEW" Then Option1.Value = True
    If TRIM$(Mid$(TTXX$, 545, 8)) = "DIRECT" Then Option2.Value = True
    '
    Option3.Value = True
    If TRIM$(Mid$(TTXX$, 553, 8)) = "CLIENT" Then Option4.Value = True
    '
End Sub


