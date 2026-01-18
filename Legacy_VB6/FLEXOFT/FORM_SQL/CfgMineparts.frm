VERSION 5.00
Begin VB.Form CfgMineparts 
   BackColor       =   &H009FBE85&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Configuración Especial Modo AHP"
   ClientHeight    =   3840
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   6540
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3840
   ScaleWidth      =   6540
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00400040&
      ForeColor       =   &H00FFFFFF&
      Height          =   8295
      Left            =   5400
      ScaleHeight     =   8235
      ScaleWidth      =   1185
      TabIndex        =   1
      Top             =   -75
      Width           =   1245
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
         Height          =   650
         Left            =   150
         Picture         =   "CfgMineparts.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   165
         Width           =   800
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Guardar"
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
         Left            =   150
         Picture         =   "CfgMineparts.frx":014A
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   2520
         Width           =   800
      End
      Begin VB.CommandButton Command11 
         Caption         =   "Trasla."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   700
         Left            =   0
         Picture         =   "CfgMineparts.frx":0454
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Generar Diskette de Transferencia"
         Top             =   7920
         Visible         =   0   'False
         Width           =   800
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Imprimir"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   700
         Left            =   0
         Picture         =   "CfgMineparts.frx":075E
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Aprueba - Graba e Imprime Remito"
         Top             =   7920
         Width           =   800
      End
      Begin VB.Image Image2 
         Height          =   510
         Left            =   -405
         Picture         =   "CfgMineparts.frx":0DC8
         ToolTipText     =   "Acceso Directo a Otras Aplicaciones"
         Top             =   3360
         Width           =   2010
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H009FBE85&
      Caption         =   "Cliente Mineparts"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3615
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5160
      Begin VB.Frame Frame7 
         BackColor       =   &H009FBE85&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   930
         Index           =   2
         Left            =   120
         TabIndex        =   16
         Top             =   2400
         Width           =   4900
         Begin VB.TextBox Text45 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   2520
            MaxLength       =   25
            TabIndex        =   18
            Top             =   420
            Width           =   855
         End
         Begin VB.Label Label6 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Porcentaje Ganancia AHP"
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
            Left            =   0
            TabIndex        =   17
            Top             =   480
            Width           =   2355
         End
      End
      Begin VB.Frame Frame7 
         BackColor       =   &H009FBE85&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   930
         Index           =   1
         Left            =   160
         TabIndex        =   11
         Top             =   1320
         Width           =   4900
         Begin VB.CommandButton Command3 
            Caption         =   "."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   2040
            TabIndex        =   12
            Top             =   360
            Width           =   175
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Proveedor Mineparts:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   465
            Left            =   -30
            TabIndex        =   15
            Top             =   240
            Width           =   1110
         End
         Begin VB.Label Label2 
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   2280
            TabIndex        =   14
            Top             =   360
            Width           =   2535
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1200
            TabIndex        =   13
            Top             =   360
            Width           =   855
         End
      End
      Begin VB.Frame Frame7 
         BackColor       =   &H009FBE85&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   930
         Index           =   0
         Left            =   160
         TabIndex        =   6
         Top             =   285
         Width           =   4900
         Begin VB.CommandButton Command8 
            Caption         =   "."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   2040
            TabIndex        =   8
            Top             =   360
            Width           =   175
         End
         Begin VB.Label Label9 
            Alignment       =   1  'Right Justify
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1200
            TabIndex        =   10
            Top             =   360
            Width           =   855
         End
         Begin VB.Label Label4 
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   2280
            TabIndex        =   9
            Top             =   360
            Width           =   2535
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Cliente AHP:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   465
            Left            =   225
            TabIndex        =   7
            Top             =   240
            Width           =   765
         End
      End
   End
End
Attribute VB_Name = "CfgMineparts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
    Call CIERRARCH("*.*")
    Call FINAL("")
End Sub



Private Sub Command2_Click()
    '
    Call GRACONTROL("CFGMINEP", "CLTEMINE", Label9)
    Call GRACONTROL("CFGMINEP", "PRFEMINE", Label1)
    Call GRACONTROL("CFGMINEP", "PJGANAHP", TRIM$(FORMATNUM$(XVALO(Text45), "F16.2")))
    Call CIERRARCH("*.*")
    Hide
    '
End Sub

Private Sub Command8_Click()

   NCLIE = ListaSeleccionClientes("", True)
   If NCLIE < 1 Then Exit Sub
   Label9 = CStr(NCLIE)
   

End Sub
Private Sub Command3_Click()

   NP = ListaSeleccionProveedores("", True)
   If NP < 1 Then Exit Sub
   Label1 = CStr(NP)

End Sub


Private Sub Form_Load()
    Label9 = Control("CFGMINEP", "CLTEMINE")
    Label1 = Control("CFGMINEP", "PRFEMINE")
    Text45 = Control("CFGMINEP", "PJGANAHP")
End Sub






Private Sub Label9_Change()

    Label4 = rasocli$(XVALO(Label9))
End Sub
Private Sub Label1_Change()

    Label2 = rasocli$(-1 * XVALO(Label1))
End Sub

