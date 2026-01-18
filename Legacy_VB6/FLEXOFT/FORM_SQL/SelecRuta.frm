VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form SelecRuta 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Definición Accesos Directos"
   ClientHeight    =   2025
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6870
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2025
   ScaleWidth      =   6870
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog CMD 
      Left            =   4440
      Top             =   1560
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame1 
      Caption         =   "Configuración "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6615
      Begin VB.CommandButton Command1 
         Caption         =   "Grabar"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5160
         TabIndex        =   7
         Top             =   1080
         Width           =   1215
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   765
         Left            =   315
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   2
         Top             =   630
         Width           =   4740
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Examinar"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   5160
         TabIndex        =   1
         Top             =   630
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "Ruta:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   360
         Width           =   3255
      End
      Begin VB.Label Label1 
         Caption         =   "Label1"
         Height          =   375
         Left            =   2880
         TabIndex        =   3
         Top             =   600
         Width           =   15
      End
   End
   Begin VB.Label Label7 
      Height          =   330
      Left            =   840
      TabIndex        =   6
      Top             =   3990
      Width           =   435
   End
   Begin VB.Label Label6 
      Height          =   330
      Left            =   105
      TabIndex        =   5
      Top             =   3885
      Visible         =   0   'False
      Width           =   435
   End
End
Attribute VB_Name = "SelecRuta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Command1_Click()
   Call GRACONTROL("CONFIRUT", "RUTAXMDB", Text1)
   Hide
End Sub





Private Sub Command4_Click()
   Command4.Enabled = False
   Archix$ = SELECCION_Archivo$("", CMD)
   If TRIM$(Archix$) = "" Then GoTo 99
   Text1 = TRIM$(Archix$)
   Screen.MousePointer = 1
99 Command4.Enabled = True

End Sub

Private Sub Form_Load()
   Call APLICACIONSKINS(Me)
   Text1 = CONTROL("CONFIRUT", "RUTAXMDB")
End Sub

