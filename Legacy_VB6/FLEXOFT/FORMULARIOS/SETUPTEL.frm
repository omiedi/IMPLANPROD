VERSION 5.00
Begin VB.Form SETUPTEL 
   BackColor       =   &H00D0FFD0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Configuración de Funcionamiento"
   ClientHeight    =   4905
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   8685
   Icon            =   "SETUPTEL.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4905
   ScaleWidth      =   8685
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame2 
      BackColor       =   &H00D0FFD0&
      Caption         =   "Linea Externa"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   960
      Left            =   6090
      TabIndex        =   22
      Top             =   3780
      Width           =   2430
      Begin VB.Frame Frame3 
         BackColor       =   &H00C0EEFF&
         BorderStyle     =   0  'None
         Caption         =   "Frame4"
         Height          =   435
         Left            =   630
         TabIndex        =   24
         Top             =   4305
         Width           =   3690
      End
      Begin VB.TextBox Text4 
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
         Index           =   5
         Left            =   945
         TabIndex        =   23
         Top             =   525
         Width           =   1260
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Acceso:"
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
         Index           =   5
         Left            =   0
         TabIndex        =   25
         Top             =   525
         Width           =   855
      End
   End
   Begin VB.Frame Frame7 
      BackColor       =   &H00D0FFD0&
      Caption         =   "Acceso a Outlook Express"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4530
      Left            =   210
      TabIndex        =   5
      Top             =   210
      Width           =   5685
      Begin VB.FileListBox File1 
         Height          =   3210
         Left            =   2940
         Pattern         =   "*.exe"
         TabIndex        =   11
         Top             =   420
         Width           =   2535
      End
      Begin VB.DriveListBox Drive1 
         Height          =   315
         Left            =   210
         TabIndex        =   8
         Top             =   420
         Width           =   1275
      End
      Begin VB.DirListBox Dir1 
         Height          =   2790
         Left            =   210
         TabIndex        =   7
         Top             =   840
         Width           =   2535
      End
      Begin VB.CommandButton Command4 
         Enabled         =   0   'False
         Height          =   315
         Left            =   2205
         Picture         =   "SETUPTEL.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   420
         Width           =   540
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "Ruta de Acceso:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   210
         TabIndex        =   10
         Top             =   3750
         Width           =   1590
      End
      Begin VB.Label Label11 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   420
         TabIndex        =   9
         Top             =   4070
         Width           =   5055
      End
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Save"
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
      Left            =   6930
      Picture         =   "SETUPTEL.frx":08BC
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Guarda Configuración del Funcionamiento"
      Top             =   315
      Width           =   750
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Default"
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
      Left            =   6090
      Picture         =   "SETUPTEL.frx":0BC6
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Restaura Valores Predeterminados"
      Top             =   315
      Width           =   750
   End
   Begin VB.CommandButton BOTEXIT 
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
      Left            =   7770
      Picture         =   "SETUPTEL.frx":0D10
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Abandona la Configuración"
      Top             =   315
      Width           =   750
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00D0FFD0&
      Caption         =   "Configuración"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2325
      Left            =   6090
      TabIndex        =   0
      Top             =   1365
      Width           =   2430
      Begin VB.TextBox Text4 
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
         Index           =   4
         Left            =   1470
         TabIndex        =   20
         Text            =   "1 "
         Top             =   1890
         Width           =   750
      End
      Begin VB.TextBox Text4 
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
         Index           =   0
         Left            =   1470
         TabIndex        =   15
         Text            =   "2"
         Top             =   420
         Width           =   750
      End
      Begin VB.TextBox Text4 
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
         Index           =   1
         Left            =   1470
         TabIndex        =   14
         Text            =   "9600 "
         Top             =   945
         Width           =   750
      End
      Begin VB.TextBox Text4 
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
         Index           =   2
         Left            =   1470
         TabIndex        =   13
         Text            =   "N "
         Top             =   1260
         Width           =   750
      End
      Begin VB.TextBox Text4 
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
         Index           =   3
         Left            =   1470
         TabIndex        =   12
         Text            =   "8"
         Top             =   1575
         Width           =   750
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00C0EEFF&
         BorderStyle     =   0  'None
         Caption         =   "Frame4"
         Height          =   435
         Left            =   630
         TabIndex        =   3
         Top             =   4305
         Width           =   3690
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Bits Parada:"
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
         Index           =   4
         Left            =   315
         TabIndex        =   21
         Top             =   1890
         Width           =   1065
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Puerto:"
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
         Index           =   0
         Left            =   525
         TabIndex        =   19
         Top             =   420
         Width           =   855
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Baudios:"
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
         Index           =   1
         Left            =   525
         TabIndex        =   18
         Top             =   945
         Width           =   855
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Bits Datos:"
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
         Index           =   2
         Left            =   315
         TabIndex        =   17
         Top             =   1575
         Width           =   1065
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Paridad:"
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
         Index           =   3
         Left            =   210
         TabIndex        =   16
         Top             =   1260
         Width           =   1170
      End
   End
End
Attribute VB_Name = "SETUPTEL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub BOTEXIT_Click()
    Call CIERRARCH("*.*")
    Hide
End Sub


Private Sub Command2_Click()
   Text4(0) = "2"
   Text4(1) = "9600"
   Text4(2) = "N"
   Text4(3) = "8"
   Text4(4) = "1"
   Text4(5) = ""
End Sub

Private Sub Command3_Click()
   SETUPSTRI$ = Space$(48)
   For KKI% = 0 To 4
     Call REPLA(SETUPSTRI$, TRIM$(Text4(KKI%)), 6 * KKI% + 1, 6)
   Next KKI%
   Call GRACONTROL("AGETELEF", "SETUPST", SETUPSTRI$)
   Call GRACONTROL("AGETELEF", "RUT-OUT", Dir1.Path)
   Call GRACONTROL("AGETELEF", "EXE-OUT", File1.filename)
   Call GRACONTROL("AGETELEF", "TOMALIN", Text4(5).Text)
   Call CIERRARCH("*.*")
   Call BAJALDISCO
   Hide
End Sub

Private Sub Dir1_Change()
   File1.Path = Dir1.Path
End Sub

Private Sub Drive1_Change()
   On Error Resume Next
   Dir1.Path = Drive1.Drive
   If Err Then
      Drive1.Drive = Left$(CurDir, 2)
      Exit Sub
   End If
End Sub

Private Sub File1_Click()
   Label11.Caption = Dir1.Path + "\" + File1.filename
   Label11.Refresh
End Sub

Private Sub Form_Load()
   SETUPSTRI$ = CONTROL("AGETELEF", "SETUPST")
   If SETUPSTRI$ = "" Then
     SETUPSTRI$ = "2     9600  N     8     1     "
   End If
   For KKI% = 0 To 4
     Text4(KKI%) = TRIM$(Mid$(SETUPSTRI$, 6 * KKI% + 1, 6))
   Next KKI%
   Text4(5) = TRIM$(CONTROL("AGETELEF", "TOMALIN"))
   On Error Resume Next
   Dir1.Path = CONTROL("AGETELEF", "RUT-OUT")
   If Err Then
     On Error GoTo 0
     Dir1.Path = CurDir
   End If
   On Error GoTo 0
   FISEL$ = CONTROL("AGETELEF", "EXE-OUT")
   For KKI% = 1 To File1.ListCount
     If File1.List(KKI% - 1) = FISEL$ Then
       File1.ListIndex = KKI% - 1
     End If
   Next KKI%
   Label11.Caption = Dir1.Path + "\" + FISEL$
   Label11.Refresh
End Sub

