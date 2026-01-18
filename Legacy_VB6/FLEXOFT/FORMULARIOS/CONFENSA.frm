VERSION 5.00
Begin VB.Form CONFENSA 
   BackColor       =   &H00FFFFC0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Ensayos de Bombas"
   ClientHeight    =   3600
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   6510
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
   ScaleHeight     =   3600
   ScaleWidth      =   6510
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
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
      Height          =   750
      Left            =   5565
      Picture         =   "CONFENSA.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   315
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
      Height          =   750
      Left            =   5565
      Picture         =   "CONFENSA.frx":014A
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1155
      Width           =   800
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFC0&
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
      Height          =   3270
      Left            =   210
      TabIndex        =   2
      Top             =   210
      Width           =   5205
      Begin VB.Frame Frame4 
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   0  'None
         Caption         =   "Frame4"
         Height          =   855
         Index           =   2
         Left            =   105
         TabIndex        =   15
         Top             =   2310
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
            Height          =   620
            Index           =   2
            Left            =   4200
            Picture         =   "CONFENSA.frx":0454
            Style           =   1  'Graphical
            TabIndex        =   18
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
            Index           =   2
            Left            =   3885
            TabIndex        =   17
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
            Index           =   2
            Left            =   2625
            TabIndex        =   16
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
            Index           =   2
            Left            =   210
            TabIndex        =   20
            Top             =   430
            Width           =   3855
         End
         Begin VB.Label FORMUDES 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Bombas Dosificadoras:"
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
            Left            =   210
            TabIndex        =   19
            Top             =   160
            Width           =   2325
         End
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   0  'None
         Caption         =   "Frame4"
         Height          =   855
         Index           =   1
         Left            =   105
         TabIndex        =   9
         Top             =   1365
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
            TabIndex        =   12
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
            TabIndex        =   11
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
            Height          =   620
            Index           =   1
            Left            =   4200
            Picture         =   "CONFENSA.frx":075E
            Style           =   1  'Graphical
            TabIndex        =   10
            Top             =   105
            Width           =   540
         End
         Begin VB.Label FORMUDES 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Bombas Tipo DVR:"
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
            TabIndex        =   14
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
            TabIndex        =   13
            Top             =   430
            Width           =   3855
         End
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   0  'None
         Caption         =   "Frame4"
         Height          =   855
         Index           =   0
         Left            =   105
         TabIndex        =   3
         Top             =   420
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
            Height          =   620
            Index           =   0
            Left            =   4200
            Picture         =   "CONFENSA.frx":0A68
            Style           =   1  'Graphical
            TabIndex        =   6
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
            TabIndex        =   5
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
            TabIndex        =   4
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
            TabIndex        =   8
            Top             =   430
            Width           =   3855
         End
         Begin VB.Label FORMUDES 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Bombas de Vacío:"
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
            TabIndex        =   7
            Top             =   160
            Width           =   2325
         End
      End
   End
   Begin VB.Image Image2 
      Height          =   555
      Left            =   5355
      Picture         =   "CONFENSA.frx":0D72
      Stretch         =   -1  'True
      Top             =   2910
      Width           =   4725
   End
End
Attribute VB_Name = "CONFENSA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub BOTOPENFORM_Click(Index As Integer)
    Call CONECRUN("PRIFORMS/" + TRIM$(FORMUNAME(Index).TEXT), "Formularios de Impresion")
    If EXISTE%("PRIFORMS.EXE") < 1 Then
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
       Call GRACONTROL("FORENSA", "FORENSA" + TRIM$(Str$(KK%)), FORMUNAME(KK%).TEXT)
    Next KK%
    '
    Call CIERRARCH("*.*")
    Hide
    '
End Sub

Private Sub Form_Load()
    '
    For KK% = 0 To 2
       On Error Resume Next
       FORMUNAME(KK%).TEXT = CONTROL$("FORENSA", "FORENSA" + TRIM$(Str$(KK%)))
       FORMUECO(KK%).Caption = ECOARCH$("INDIFRM", FORMUNAME(KK%).TEXT)
    Next KK%
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

