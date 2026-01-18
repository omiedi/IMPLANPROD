VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form CFGSCGPLUS 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0FFC0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Contabilidad General - Configuracion Avanzada"
   ClientHeight    =   4455
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   5430
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
   ScaleHeight     =   4455
   ScaleWidth      =   5430
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Picture999 
      Height          =   330
      Index           =   0
      Left            =   1470
      ScaleHeight     =   270
      ScaleWidth      =   1110
      TabIndex        =   15
      Top             =   525
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.CommandButton Command2 
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
      Height          =   855
      Left            =   4350
      Picture         =   "CFGSCGPLUS.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Emite Reporte Solicitado"
      Top             =   1320
      Width           =   855
   End
   Begin VB.CommandButton command1 
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
      Height          =   855
      Left            =   4350
      Picture         =   "CFGSCGPLUS.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Abandona la Aplicación"
      Top             =   360
      Width           =   855
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Privilegios"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2970
      Left            =   150
      TabIndex        =   3
      Top             =   1320
      Width           =   4005
      Begin VB.Frame Frame3 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Caja y Bancos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   765
         Left            =   240
         TabIndex        =   12
         Top             =   2040
         Width           =   3525
         Begin VB.OptionButton Option7 
            BackColor       =   &H00C0FFC0&
            Caption         =   "Owner"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   600
            TabIndex        =   14
            Top             =   360
            Value           =   -1  'True
            Width           =   1155
         End
         Begin VB.OptionButton Option8 
            BackColor       =   &H00C0FFC0&
            Caption         =   "Guest"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   2280
            TabIndex        =   13
            Top             =   360
            Width           =   1000
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Contabilidad General"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   765
         Left            =   240
         TabIndex        =   9
         Top             =   1200
         Width           =   3525
         Begin VB.OptionButton Option5 
            BackColor       =   &H00C0FFC0&
            Caption         =   "Owner"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   600
            TabIndex        =   11
            Top             =   360
            Value           =   -1  'True
            Width           =   1155
         End
         Begin VB.OptionButton Option6 
            BackColor       =   &H00C0FFC0&
            Caption         =   "Guest"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   2280
            TabIndex        =   10
            Top             =   360
            Width           =   1000
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Plan de Cuentas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   765
         Left            =   240
         TabIndex        =   4
         Top             =   360
         Width           =   3525
         Begin VB.OptionButton Option4 
            BackColor       =   &H00C0FFC0&
            Caption         =   "Guest"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   2280
            TabIndex        =   6
            Top             =   360
            Width           =   1000
         End
         Begin VB.OptionButton Option3 
            BackColor       =   &H00C0FFC0&
            Caption         =   "Owner"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   600
            TabIndex        =   5
            Top             =   360
            Value           =   -1  'True
            Width           =   1155
         End
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Sistema"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   960
      Left            =   150
      TabIndex        =   0
      Top             =   240
      Width           =   4005
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Local"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   480
         TabIndex        =   2
         Top             =   480
         Value           =   -1  'True
         Width           =   975
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Distribuido (2 sedes)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1680
         TabIndex        =   1
         Top             =   480
         Width           =   2175
      End
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   2625
      OleObjectBlob   =   "CFGSCGPLUS.frx":0454
      Top             =   0
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   225
      Index           =   0
      Left            =   0
      OleObjectBlob   =   "CFGSCGPLUS.frx":0688
      TabIndex        =   16
      Top             =   0
      Visible         =   0   'False
      Width           =   1485
   End
   Begin VB.Image Image2 
      Height          =   510
      Left            =   3840
      Picture         =   "CFGSCGPLUS.frx":06E6
      Top             =   3750
      Width           =   2010
   End
End
Attribute VB_Name = "CFGSCGPLUS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
    Call CIERRARCH("*.*")
    Unload Me
End Sub
'

Private Sub Command2_Click()
   Command2.Enabled = False
   '
   If Option1.Value = True Then
     Option3.Value = True
     Option5.Value = True
     Option7.Value = True
   End If
   '
   SSTX$ = "LOCAL": If Option2.Value = True Then SSTX$ = "DISTRIBU"
   Call GRACONTROL("CONTAGEN", "SIDISTRI", SSTX$)
   '
   SSTX$ = "OWNER": If Option4.Value = True Then SSTX$ = "GUEST"
   Call GRACONTROL("PROPSCG", "PLAGRAL", SSTX$)
   '
   SSTX$ = "OWNER": If Option6.Value = True Then SSTX$ = "GUEST"
   Call GRACONTROL("PROPSCG", "CONTAGEN", SSTX$)
   '
   SSTX$ = "OWNER": If Option8.Value = True Then SSTX$ = "GUEST"
   Call GRACONTROL("PROPSCG", "CAJABANC", SSTX$)
   '
   Command2.Enabled = True
End Sub

Private Sub Form_Load()
    '
    Option1.Value = True
    If Control$("CONTAGEN", "SIDISTRI") = "DISTRIBU" Then Option2.Value = True
    '
    Option3.Value = True
    If Control$("PROPSCG", "PLAGRAL") = "GUEST" Then Option4.Value = True
    '
    Option5.Value = True
    If Control$("PROPSCG", "CONTAGEN") = "GUEST" Then Option6.Value = True
    '
    Option7.Value = True
    If Control$("PROPSCG", "CAJABANC") = "GUEST" Then Option8.Value = True
    '
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call CIERRARCH("*.*")
End Sub

'
