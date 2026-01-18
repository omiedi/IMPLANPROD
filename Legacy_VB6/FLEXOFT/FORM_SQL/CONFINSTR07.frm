VERSION 5.00
Begin VB.Form CONFINSTR07 
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Inventario de Instrumentos de Medicion y Ensayo"
   ClientHeight    =   3525
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   6525
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
   ScaleHeight     =   3525
   ScaleWidth      =   6525
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00008080&
      Height          =   3735
      Left            =   5640
      ScaleHeight     =   3675
      ScaleWidth      =   1155
      TabIndex        =   19
      Top             =   -120
      Width           =   1215
      Begin VB.CommandButton Command7 
         Caption         =   "Grupo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   690
         Left            =   120
         Picture         =   "CONFINSTR07.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   22
         ToolTipText     =   "Grupos de Instrumentos"
         Top             =   1830
         Width           =   650
      End
      Begin VB.CommandButton Command2 
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
         Height          =   690
         Left            =   120
         Picture         =   "CONFINSTR07.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   1140
         Width           =   650
      End
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
         Height          =   690
         Left            =   120
         Picture         =   "CONFINSTR07.frx":0614
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   450
         Width           =   650
      End
      Begin VB.Image Image2 
         Height          =   390
         Left            =   -300
         Picture         =   "CONFINSTR07.frx":075E
         Top             =   3050
         Width           =   1515
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Formularios de Impresion"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3165
      Left            =   210
      TabIndex        =   0
      Top             =   210
      Width           =   5205
      Begin VB.Frame Frame4 
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   0  'None
         Caption         =   "Frame4"
         Height          =   855
         Index           =   2
         Left            =   105
         TabIndex        =   13
         Top             =   2205
         Width           =   4950
         Begin VB.TextBox FORMUNAME 
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
            Left            =   2625
            TabIndex        =   16
            Text            =   " "
            Top             =   105
            Width           =   1275
         End
         Begin VB.CommandButton BOTSELFORM 
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
            Height          =   285
            Index           =   2
            Left            =   3885
            TabIndex        =   15
            Top             =   105
            Width           =   175
         End
         Begin VB.CommandButton BOTOPENFORM 
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
            Index           =   2
            Left            =   4200
            Picture         =   "CONFINSTR07.frx":2680
            Style           =   1  'Graphical
            TabIndex        =   14
            Top             =   105
            Width           =   540
         End
         Begin VB.Label FORMUDES 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Ficha de Calibración:"
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
            Left            =   105
            TabIndex        =   18
            Top             =   165
            Width           =   2430
         End
         Begin VB.Label FORMUECO 
            BackColor       =   &H80000009&
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
            Height          =   285
            Index           =   2
            Left            =   210
            TabIndex        =   17
            Top             =   430
            Width           =   3855
         End
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   0  'None
         Caption         =   "Frame4"
         Height          =   855
         Index           =   1
         Left            =   105
         TabIndex        =   7
         Top             =   1260
         Width           =   4950
         Begin VB.TextBox FORMUNAME 
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
            Left            =   2625
            TabIndex        =   10
            Text            =   " "
            Top             =   105
            Width           =   1275
         End
         Begin VB.CommandButton BOTSELFORM 
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
            Height          =   285
            Index           =   1
            Left            =   3885
            TabIndex        =   9
            Top             =   105
            Width           =   175
         End
         Begin VB.CommandButton BOTOPENFORM 
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
            Index           =   1
            Left            =   4200
            Picture         =   "CONFINSTR07.frx":298A
            Style           =   1  'Graphical
            TabIndex        =   8
            Top             =   105
            Width           =   540
         End
         Begin VB.Label FORMUDES 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Registro de Calibración:"
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
            Left            =   210
            TabIndex        =   12
            Top             =   160
            Width           =   2325
         End
         Begin VB.Label FORMUECO 
            BackColor       =   &H80000009&
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
            Height          =   285
            Index           =   1
            Left            =   210
            TabIndex        =   11
            Top             =   430
            Width           =   3855
         End
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   0  'None
         Caption         =   "Frame4"
         Height          =   855
         Index           =   0
         Left            =   105
         TabIndex        =   1
         Top             =   315
         Width           =   4950
         Begin VB.CommandButton BOTOPENFORM 
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
            Index           =   0
            Left            =   4200
            Picture         =   "CONFINSTR07.frx":2C94
            Style           =   1  'Graphical
            TabIndex        =   4
            Top             =   105
            Width           =   540
         End
         Begin VB.CommandButton BOTSELFORM 
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
            Height          =   285
            Index           =   0
            Left            =   3885
            TabIndex        =   3
            Top             =   105
            Width           =   175
         End
         Begin VB.TextBox FORMUNAME 
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
            Left            =   2625
            TabIndex        =   2
            Text            =   " "
            Top             =   105
            Width           =   1275
         End
         Begin VB.Label FORMUECO 
            BackColor       =   &H80000009&
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
            Height          =   285
            Index           =   0
            Left            =   210
            TabIndex        =   6
            Top             =   430
            Width           =   3855
         End
         Begin VB.Label FORMUDES 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Inventario Instrumentos:"
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
            Left            =   210
            TabIndex        =   5
            Top             =   160
            Width           =   2325
         End
      End
   End
End
Attribute VB_Name = "CONFINSTR07"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub BOTOPENFORM_Click(Index As Integer)
    Call CONECRUN("PLFORMS/" + TRIM$(FORMUNAME(Index).TEXT), "Formularios de Impresion")
    If EXISTE%("PLFORMS.EXE") < 1 Then
        Call FINAL("")
    End If
End Sub

Private Sub BOTSELFORM_Click(Index As Integer)
    Call SELECHO("INDIFRM")
    VDEVU$ = VALACT1$("INDIFRM")
    If VDEVU$ <> "" Then
        FORMUNAME(Index).TEXT = VDEVU$
        FORMUECO(Index).Caption = VALACT2$("INDIFRM")
    End If
End Sub

Private Sub Command1_Click()
    Call CIERRARCH("*.*")
    Hide
End Sub

Private Sub Command2_Click()
    '
    For KK% = 0 To 2
       On Error Resume Next
       Call GRACONTROL("FORINMED", "DOCUMED" + TRIM$(Str$(KK%)), FORMUNAME(KK%).TEXT)
    Next KK%
    '
    On Error GoTo 0
    Call CIERRARCH("*.*")
    Hide
    '
End Sub

Private Sub Command7_Click()
   Command7.Enabled = False
5  OPX% = ALTERNA%("Grupo Existente\Nuevo Grupo de Instrumentos")
   Select Case OPX%
     Case 1
       Call MODIGRUINS
       GoTo 5
     Case 2
       Call ALTAGRUINS
       GoTo 5
   End Select
   Command7.Enabled = True
End Sub

Private Sub Form_Load()
    '
    For KK% = 0 To 2
       On Error Resume Next
       FORMUNAME(KK%).TEXT = CONTROL$("FORINMED", "DOCUMED" + TRIM$(Str$(KK%)))
       FORMUECO(KK%).Caption = ECOARCH$("INDIFRM", FORMUNAME(KK%).TEXT)
    Next KK%
    On Error GoTo 0
    '
End Sub

Private Sub FORMUNAME_DblClick(Index As Integer)
    Call SELECHO("INDIFRM")
    VDEVU$ = VALACT1$("INDIFRM")
    If VDEVU$ <> "" Then
        FORMUNAME(Index).TEXT = VDEVU$
        FORMUECO(Index).Caption = VALACT2$("INDIFRM")
    End If
End Sub

Sub ALTAGRUINS()
   '
   VDEF$ = REGBLAN$("GRUINSTR")
   INDINS% = 1 + ULTREG&("GRUINSTR")
   Call REPLA(VDEF$, MKI$(INDINS%), 1, 2)
   '
10 Screen.MousePointer = 1
   OBX$ = OBJPLA$("NUEGRUINSTRU", VDEF$)
   If OBX$ = "" Then
     Exit Sub
   End If
   '
   VDEF$ = OBX$
   DESINS$ = TRIM$(Mid$(OBX$, 3, 62))
   If DESINS$ = "" Then
      Call MENSERR(24, "Falta Denominación del Grupo")
      GoTo 10
   End If
   '
   Call CIERRARCH("GRUINSTR")
   Call BLOQARCH("GRUINSTR")
   INDINS% = 1 + ULTREG&("GRUINSTR")
   Call REPLA(OBX$, MKI$(INDINS%), 1, 2)
   Call REGAPP("GRUINSTR", OBX$)
   Call CIERRARCH("GRUINSTR")
   Call BAJALDISCO
   Screen.MousePointer = 1
   '
End Sub

Sub MODIGRUINS()
   '
5  XX$ = REGSEL$("GRUINSTR")
   If Len(XX$) <= 4 Then Exit Sub
   VDEF$ = XX$
   INDIN& = REGNUM&("GRUINSTR")
   '
10 Screen.MousePointer = 1
   OBX$ = OBJPLA$("NUEGRUINSTRU", VDEF$)
   If OBX$ = "" Then
     GoTo 5
   End If
   '
   VDEF$ = OBX$
   DESINS$ = TRIM$(Mid$(OBX$, 3, 62))
   If DESINS$ = "" Then
      Call MENSERR(24, "Falta Denominación del Grupo")
      GoTo 10
   End If
   '
   Call CIERRARCH("GRUINSTR")
   Call BLOQARCH("GRUINSTR")
   Call GRAREG("GRUINSTR", OBX$, INDIN&)
   Call CIERRARCH("GRUINSTR")
   Call BAJALDISCO
   Screen.MousePointer = 1
   GoTo 5
   '
End Sub


