VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.1#0"; "COMDLG32.OCX"
Begin VB.Form STPLAFAB 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0FFC0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Planilla de Producción"
   ClientHeight    =   4125
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   5580
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4125
   ScaleWidth      =   5580
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   327680
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Estructura de Tabla"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2535
      Left            =   210
      TabIndex        =   3
      Top             =   1470
      Width           =   4215
      Begin VB.Frame Frame2 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Parámetros"
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
         Index           =   1
         Left            =   210
         TabIndex        =   7
         Top             =   1470
         Width           =   3795
         Begin VB.TextBox Text5 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   3045
            TabIndex        =   15
            Top             =   315
            Width           =   540
         End
         Begin VB.TextBox Text7 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   1575
            TabIndex        =   13
            Top             =   315
            Width           =   435
         End
         Begin VB.Label Label6 
            BackStyle       =   0  'Transparent
            Caption         =   "F.Entrega:"
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
            Left            =   2100
            TabIndex        =   14
            Top             =   420
            Width           =   1380
         End
         Begin VB.Label Label8 
            BackStyle       =   0  'Transparent
            Caption         =   "Columna Código:"
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
            Left            =   105
            TabIndex        =   8
            Top             =   420
            Width           =   1485
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Rango de Datos"
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
         Index           =   0
         Left            =   210
         TabIndex        =   4
         Top             =   420
         Width           =   3795
         Begin VB.TextBox TEXT4 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   3045
            TabIndex        =   12
            Top             =   315
            Width           =   540
         End
         Begin VB.TextBox Text2 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   1575
            TabIndex        =   11
            Top             =   315
            Width           =   435
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   "Hasta:"
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
            Left            =   2415
            TabIndex        =   6
            Top             =   420
            Width           =   1380
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Renglón Desde:"
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
            Left            =   105
            TabIndex        =   5
            Top             =   420
            Width           =   1380
         End
      End
   End
   Begin VB.CommandButton Command29 
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
      Height          =   705
      Left            =   4620
      Picture         =   "STPLAFAB.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Guarda los Datos Ingresados"
      Top             =   945
      Width           =   750
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
      Height          =   705
      Left            =   4620
      Picture         =   "STPLAFAB.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Terminar - Salir de la Aplicación"
      Top             =   210
      Width           =   750
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Ruta y Archivo de Acceso"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1275
      Left            =   210
      TabIndex        =   0
      Top             =   105
      Width           =   4215
      Begin VB.TextBox Text1 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   210
         MultiLine       =   -1  'True
         TabIndex        =   16
         Top             =   420
         Width           =   3375
      End
      Begin VB.CommandButton Command2 
         Height          =   645
         Left            =   3675
         Picture         =   "STPLAFAB.frx":0454
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   420
         Width           =   330
      End
      Begin VB.Label Label9 
         BackColor       =   &H00E0E0E0&
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
         Height          =   640
         Left            =   210
         TabIndex        =   9
         Top             =   420
         Width           =   3375
         WordWrap        =   -1  'True
      End
   End
   Begin VB.Image Image2 
      Height          =   510
      Left            =   3990
      Picture         =   "STPLAFAB.frx":075E
      Top             =   3465
      Width           =   2010
   End
End
Attribute VB_Name = "STPLAFAB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
   Cancel = 0
   Call CIERRARCH("*.*")
   Hide
End Sub

Private Sub Command2_Click()
   CommonDialog1.CancelError = True
   CommonDialog1.DialogTitle = "Conectar Archivo Planilla de Producción"
       'deshabilitado porque si no da error - CommonDialog1.Flags = &H2& + &H4& + &H8 + &H800 + &H2000 + &H8000 'cdlOFNCreatePrompt
   CommonDialog1.FILTER = "Planillas Excel (*.XLS)|*.XLS|"
   CommonDialog1.FilterIndex = 1
   On Error Resume Next
   CommonDialog1.ShowOpen
   On Error GoTo 0
   '
   RUTARCH$ = CommonDialog1.filename
   If RUTARCH$ <> "" Then
     Label9 = RUTARCH$
     Text1 = RUTARCH$
   End If
   '
End Sub

Private Sub Command29_Click()
   Command29.Enabled = False
   '
   Call SAVEFILE("C:\FLEXOFT\PLANIFAB.DRV", Text1)
   Call GRACONTROL("PLANIFAB", "DERENGLO", Text2)
   Call GRACONTROL("PLANIFAB", "HARENGLO", TEXT4)
   Call GRACONTROL("PLANIFAB", "COLCODIG", Text7)
   Call GRACONTROL("PLANIFAB", "COLFENTR", Text5)
   '
   Command29.Enabled = True
End Sub

Private Sub Form_Load()
   '
   EPAC$ = TRIM$(EMPREAC$)
   If EPAC$ <> "" Then
       Caption = Caption + " - " + EPAC$
   End If
   '
   APLAFA$ = TRIM$(LOADFILE$("C:\FLEXOFT\PLANIFAB.DRV"))
   Label9 = APLAFA$
   Text1 = APLAFA$
   '
   Text2 = TRIM$(CONTROL$("PLANIFAB", "DERENGLO"))
   TEXT4 = TRIM$(CONTROL$("PLANIFAB", "HARENGLO"))
   Text7 = TRIM$(CONTROL$("PLANIFAB", "COLCODIG"))
   Text5 = TRIM$(CONTROL$("PLANIFAB", "COLFENTR"))
   '
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Cancel = 0
   Call CIERRARCH("*.*")
   Hide
End Sub

Private Sub Label5_Click()

End Sub

Private Sub Text1_GotFocus()
   Command2.SetFocus
End Sub
