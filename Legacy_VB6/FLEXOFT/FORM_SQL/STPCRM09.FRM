VERSION 5.00
Begin VB.Form STPCRM09 
   BackColor       =   &H00D0FFD0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Configuración de Funcionamiento"
   ClientHeight    =   7485
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   8685
   Icon            =   "STPCRM09.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7485
   ScaleWidth      =   8685
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command8 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   850
      Left            =   7875
      Picture         =   "STPCRM09.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   32
      TabStop         =   0   'False
      ToolTipText     =   "Configuracion Envio de Mails por FLEXWSK"
      Top             =   3885
      Width           =   600
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00008000&
      Height          =   1335
      Left            =   6120
      ScaleHeight     =   1275
      ScaleWidth      =   2475
      TabIndex        =   28
      Top             =   0
      Width           =   2535
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
         Left            =   1680
         Picture         =   "STPCRM09.frx":0884
         Style           =   1  'Graphical
         TabIndex        =   31
         ToolTipText     =   "Abandona la Configuración"
         Top             =   240
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
         Left            =   120
         Picture         =   "STPCRM09.frx":09CE
         Style           =   1  'Graphical
         TabIndex        =   30
         ToolTipText     =   "Restaura Valores Predeterminados"
         Top             =   240
         Width           =   750
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
         Left            =   900
         Picture         =   "STPCRM09.frx":0B18
         Style           =   1  'Graphical
         TabIndex        =   29
         ToolTipText     =   "Guarda Configuración del Funcionamiento"
         Top             =   240
         Width           =   750
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00D0FFD0&
      Caption         =   "Asunto, Firma y Pie de Mail"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2500
      Left            =   210
      TabIndex        =   22
      Top             =   4830
      Width           =   8310
      Begin VB.TextBox Text20 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   825
         TabIndex        =   25
         Text            =   " "
         Top             =   420
         Width           =   7380
      End
      Begin VB.TextBox Text2 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1530
         Left            =   825
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   24
         Text            =   "STPCRM09.frx":0E22
         Top             =   840
         Width           =   7350
      End
      Begin VB.Frame Frame6 
         BackColor       =   &H00C0EEFF&
         BorderStyle     =   0  'None
         Caption         =   "Frame4"
         Height          =   435
         Left            =   630
         TabIndex        =   23
         Top             =   4305
         Width           =   3690
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Pie:"
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
         Left            =   -225
         TabIndex        =   27
         Top             =   840
         Width           =   960
      End
      Begin VB.Label Label24 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Asunto:"
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
         Left            =   -225
         TabIndex        =   26
         Top             =   450
         Width           =   960
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00D0FFD0&
      Caption         =   "Linea Externa"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   960
      Left            =   6090
      TabIndex        =   19
      Top             =   3780
      Width           =   1695
      Begin VB.Frame Frame3 
         BackColor       =   &H00C0EEFF&
         BorderStyle     =   0  'None
         Caption         =   "Frame4"
         Height          =   435
         Left            =   630
         TabIndex        =   21
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
         Left            =   315
         TabIndex        =   20
         Top             =   525
         Width           =   1155
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
      TabIndex        =   2
      Top             =   210
      Width           =   5685
      Begin VB.FileListBox File1 
         Height          =   3210
         Left            =   2940
         Pattern         =   "*.exe"
         TabIndex        =   8
         Top             =   420
         Width           =   2535
      End
      Begin VB.DriveListBox Drive1 
         Height          =   315
         Left            =   210
         TabIndex        =   5
         Top             =   420
         Width           =   1275
      End
      Begin VB.DirListBox Dir1 
         Height          =   2790
         Left            =   210
         TabIndex        =   4
         Top             =   840
         Width           =   2535
      End
      Begin VB.CommandButton Command4 
         Enabled         =   0   'False
         Height          =   315
         Left            =   2205
         Picture         =   "STPCRM09.frx":0E26
         Style           =   1  'Graphical
         TabIndex        =   3
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
         TabIndex        =   7
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
         TabIndex        =   6
         Top             =   4070
         Width           =   5055
      End
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
         TabIndex        =   17
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
         TabIndex        =   12
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
         TabIndex        =   11
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
         TabIndex        =   10
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
         TabIndex        =   9
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
         TabIndex        =   1
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
         TabIndex        =   18
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
         TabIndex        =   16
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
         TabIndex        =   15
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
         TabIndex        =   14
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
         TabIndex        =   13
         Top             =   1260
         Width           =   1170
      End
   End
End
Attribute VB_Name = "STPCRM09"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub BOTEXIT_Click()
    Call CIERRARCH("*.*")
    Hide
End Sub


Private Sub Command1_Click()
    '
    Command8.Enabled = False
    '
    If DERACCE%("CFGEMAIL", "Configurar Cuenta de Correo Electronico") < 2 Then
     GoTo 99
    End If
    '
    STPFMAIL.Show 1
    '
99  Command8.Enabled = True
    '
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
   Call GRACONTROL_SQL("AGETELEF", "SETUPST", SETUPSTRI$)
   Call GRACONTROL_SQL("AGETELEF", "RUT-OUT", Dir1.Path)
   Call GRACONTROL_SQL("AGETELEF", "EXE-OUT", File1.FileName)
   Call GRACONTROL_SQL("AGETELEF", "TOMALIN", Text4(5).TEXT)
   Call GRACONTROL_SQL("SEGUICOM", "ASUNTO", Text20)
   '
   Call GRACONTROL_SQL("SEGUICOM", "PIEDMAI-" & TRIM$(TRIM$(USNBR%)), Text2)
   '
End Sub


Private Sub Command8_Click()
    '
    Command8.Enabled = False
    '
    If DERACCE%("CFGEMAIL", "Configurar Cuenta de Correo Electronico") < 2 Then
     GoTo 99
    End If
    '
    STPFMAIL.Show 1
    '
99  Command8.Enabled = True
    '
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

Private Sub file1_click()
   Label11.Caption = Dir1.Path + "\" + File1.FileName
   Label11.Refresh
End Sub

Private Sub Form_Load()
   SETUPSTRI$ = CONTROL_SQL("AGETELEF", "SETUPST")
   If SETUPSTRI$ = "" Then
     SETUPSTRI$ = "2     9600  N     8     1     "
   End If
   For KKI% = 0 To 4
     Text4(KKI%) = TRIM$(Mid$(SETUPSTRI$, 6 * KKI% + 1, 6))
   Next KKI%
   Text4(5) = TRIM$(CONTROL_SQL("AGETELEF", "TOMALIN"))
   Dir1.Path = CONTROL_SQL("AGETELEF", "RUT-OUT")
   FISEL$ = CONTROL_SQL("AGETELEF", "EXE-OUT")
   For KKI% = 1 To File1.ListCount
     If File1.List(KKI% - 1) = FISEL$ Then
       File1.ListIndex = KKI% - 1
     End If
   Next KKI%
   Label11.Caption = Dir1.Path + "\" + FISEL$
   Label11.Refresh
   '
   Text20 = CONTROL_SQL$("SEGUICOM", "ASUNTO")
   Text2 = CONTROL_SQL$("SEGUICOM", "PIEDMAI-" & TRIM$(TRIM$(USNBR%)))
   '
End Sub

