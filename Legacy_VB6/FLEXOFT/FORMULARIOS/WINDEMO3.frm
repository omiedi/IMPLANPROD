VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.1#0"; "COMCTL32.OCX"
Begin VB.Form WINDEMO3 
   BackColor       =   &H00C00000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Sistema FLEXOFT - Demostración Interactiva"
   ClientHeight    =   8655
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   11910
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8655
   ScaleWidth      =   11910
   Begin VB.Frame Frame2 
      BackColor       =   &H00C00000&
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      ForeColor       =   &H8000000E&
      Height          =   8730
      Left            =   10710
      TabIndex        =   0
      Top             =   -105
      Width           =   1485
      Begin ComctlLib.Slider Slider1 
         Height          =   540
         Left            =   420
         TabIndex        =   15
         Top             =   7770
         Width           =   645
         _ExtentX        =   1138
         _ExtentY        =   953
         _Version        =   327680
         BorderStyle     =   1
         SelStart        =   5
         Value           =   5
      End
      Begin VB.CommandButton Command1 
         Height          =   645
         Index           =   0
         Left            =   420
         Picture         =   "WINDEMO3.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Finalizar Demostración"
         Top             =   315
         Width           =   645
      End
      Begin VB.CommandButton Command1 
         Height          =   645
         Index           =   1
         Left            =   420
         Picture         =   "WINDEMO3.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Volver al Comienzo"
         Top             =   1365
         Width           =   645
      End
      Begin VB.CommandButton Command1 
         Height          =   645
         Index           =   2
         Left            =   420
         Picture         =   "WINDEMO3.frx":0614
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Próxima Pantalla"
         Top             =   2415
         Width           =   645
      End
      Begin VB.CommandButton Command1 
         Height          =   645
         Index           =   3
         Left            =   420
         Picture         =   "WINDEMO3.frx":091E
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Pantalla Anterior"
         Top             =   3465
         Width           =   645
      End
      Begin VB.CommandButton Command1 
         Height          =   645
         Index           =   4
         Left            =   420
         Picture         =   "WINDEMO3.frx":0C28
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Suspende Ejecución Automática"
         Top             =   4515
         Width           =   645
      End
      Begin VB.CommandButton Command1 
         Height          =   645
         Index           =   5
         Left            =   420
         Picture         =   "WINDEMO3.frx":0F32
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Presenta Información Adicional"
         Top             =   5565
         Width           =   645
      End
      Begin VB.CommandButton Command1 
         Height          =   645
         Index           =   6
         Left            =   420
         Picture         =   "WINDEMO3.frx":123C
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Imprime Pantalla Actual"
         Top             =   6615
         Width           =   645
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Speed"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   225
         Index           =   3
         Left            =   210
         TabIndex        =   17
         Top             =   8400
         Width           =   1065
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "End"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   225
         Index           =   0
         Left            =   210
         TabIndex        =   14
         Top             =   1005
         Width           =   1065
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
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
         ForeColor       =   &H00FFFFFF&
         Height          =   225
         Index           =   1
         Left            =   210
         TabIndex        =   13
         Top             =   2055
         Width           =   1065
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Next"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   225
         Index           =   7
         Left            =   210
         TabIndex        =   12
         Top             =   3105
         Width           =   1065
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Back"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   225
         Index           =   8
         Left            =   210
         TabIndex        =   11
         Top             =   4155
         Width           =   1065
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Pause"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   225
         Index           =   9
         Left            =   210
         TabIndex        =   10
         Top             =   5205
         Width           =   1065
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "More..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   225
         Index           =   10
         Left            =   210
         TabIndex        =   9
         Top             =   6255
         Width           =   1065
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Print"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   225
         Index           =   11
         Left            =   210
         TabIndex        =   8
         Top             =   7305
         Width           =   1065
      End
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Speed"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Index           =   2
      Left            =   10920
      TabIndex        =   16
      Top             =   8295
      Width           =   1065
   End
End
Attribute VB_Name = "WINDEMO3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click(Index As Integer)
    WINDEMO1.Show
    Hide
End Sub

Private Sub Form_ACTIVATE()
    Call PRESDEM
End Sub
