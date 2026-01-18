VERSION 5.00
Begin VB.Form CfgContep 
   BackColor       =   &H009FBE85&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Configuración Especial Modo AHP"
   ClientHeight    =   5160
   ClientLeft      =   45
   ClientTop       =   600
   ClientWidth     =   14370
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
   ScaleHeight     =   5160
   ScaleWidth      =   14370
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00400040&
      ForeColor       =   &H00FFFFFF&
      Height          =   8295
      Left            =   13320
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
         Picture         =   "CfgContep.frx":0000
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
         Picture         =   "CfgContep.frx":014A
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
         Picture         =   "CfgContep.frx":0454
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
         Picture         =   "CfgContep.frx":075E
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Aprueba - Graba e Imprime Remito"
         Top             =   7920
         Width           =   800
      End
      Begin VB.Image Image2 
         Height          =   510
         Left            =   -405
         Picture         =   "CfgContep.frx":0DC8
         ToolTipText     =   "Acceso Directo a Otras Aplicaciones"
         Top             =   3360
         Width           =   2010
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H009FBE85&
      Caption         =   "Configuración Participantes de Compras Internas"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4935
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   13080
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
         TabIndex        =   12
         Top             =   3720
         Width           =   12945
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
            TabIndex        =   14
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
            TabIndex        =   13
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
         Top             =   2640
         Width           =   12810
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
         Height          =   1770
         Index           =   0
         Left            =   160
         TabIndex        =   6
         Top             =   285
         Width           =   12810
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
            Top             =   240
            Width           =   175
         End
         Begin VB.Label Label1 
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
            Height          =   915
            Left            =   1200
            TabIndex        =   16
            ToolTipText     =   "Doble Click para seleccionar ruta"
            Top             =   600
            Width           =   11415
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Ruta:"
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
            Index           =   1
            Left            =   -120
            TabIndex        =   15
            Top             =   720
            Width           =   885
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
            Top             =   240
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
            Left            =   2400
            TabIndex        =   9
            Top             =   240
            Width           =   3135
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Proveedor:"
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
            Index           =   0
            Left            =   105
            TabIndex        =   7
            Top             =   240
            Width           =   1110
         End
      End
   End
   Begin VB.Menu mnuArchivo 
      Caption         =   "Archivo"
   End
   Begin VB.Menu mnuConfig 
      Caption         =   "Configuración"
      Begin VB.Menu mnuTAux 
         Caption         =   "Tablas Auxiliares"
      End
   End
End
Attribute VB_Name = "CfgContep"
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





Private Sub Label1_Click()
    SELCARPETA.CONTROL1 = "COMPINTE"
    SELCARPETA.CONTROL2 = "CARPCTE1"
    SELCARPETA.Show 1
    Label1 = CONTROL("COMPINTE", "CARPCTE1")

End Sub

Private Sub Label9_Change()

    Label4 = rasocli$(XVALO(Label9))
End Sub


