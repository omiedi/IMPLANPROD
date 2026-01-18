VERSION 5.00
Begin VB.Form GRAFITEM 
   BackColor       =   &H00C0E0FF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Artículos de Stock - Gráficos Asociados "
   ClientHeight    =   7440
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   10200
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7440
   ScaleWidth      =   10200
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command9 
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
      Height          =   330
      Left            =   3570
      Picture         =   "GRAFITEM.frx":0000
      TabIndex        =   34
      Top             =   240
      Width           =   175
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Control"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1200
      Left            =   7245
      TabIndex        =   22
      Top             =   120
      Width           =   2745
      Begin VB.CommandButton Command2 
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
         Left            =   1800
         Picture         =   "GRAFITEM.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   33
         ToolTipText     =   "Ayuda en Línea"
         Top             =   315
         Width           =   750
      End
      Begin VB.CommandButton Command1 
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
         Height          =   750
         Left            =   210
         Picture         =   "GRAFITEM.frx":0454
         Style           =   1  'Graphical
         TabIndex        =   24
         ToolTipText     =   "Guarda los Cambios Realizados"
         Top             =   315
         Width           =   750
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Help"
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
         Left            =   1050
         Picture         =   "GRAFITEM.frx":075E
         Style           =   1  'Graphical
         TabIndex        =   23
         ToolTipText     =   "Ayuda en Línea"
         Top             =   315
         Width           =   750
      End
   End
   Begin VB.Frame MARCOFICHA 
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   0  'None
      Caption         =   "5"
      Height          =   5895
      Index           =   5
      Left            =   0
      TabIndex        =   0
      Top             =   1470
      Width           =   10095
      Begin VB.PictureBox Picture2 
         AutoRedraw      =   -1  'True
         Height          =   645
         Left            =   4620
         ScaleHeight     =   585
         ScaleWidth      =   900
         TabIndex        =   21
         Top             =   2310
         Visible         =   0   'False
         Width           =   960
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Fotografia 2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2650
         Index           =   3
         Left            =   5250
         TabIndex        =   16
         Top             =   3045
         Width           =   4750
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
            Height          =   285
            Index           =   3
            Left            =   3780
            TabIndex        =   19
            Top             =   315
            Width           =   175
         End
         Begin VB.FileListBox File1 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1845
            Index           =   3
            Left            =   840
            TabIndex        =   18
            Top             =   630
            Visible         =   0   'False
            Width           =   3165
         End
         Begin VB.PictureBox Picture1 
            AutoRedraw      =   -1  'True
            Height          =   1695
            Index           =   3
            Left            =   315
            ScaleHeight     =   1635
            ScaleWidth      =   4095
            TabIndex        =   17
            Top             =   735
            Width           =   4150
         End
         Begin VB.Image Image1 
            Height          =   19995
            Index           =   3
            Left            =   0
            Stretch         =   -1  'True
            Top             =   105
            Visible         =   0   'False
            Width           =   19995
         End
         Begin VB.Label Label34 
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
            Height          =   285
            Index           =   3
            Left            =   840
            TabIndex        =   20
            Top             =   315
            Width           =   2955
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Plano"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2650
         Index           =   0
         Left            =   210
         TabIndex        =   11
         Top             =   210
         Width           =   4750
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
            Height          =   285
            Index           =   0
            Left            =   3780
            TabIndex        =   14
            Top             =   315
            Width           =   175
         End
         Begin VB.FileListBox File1 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1845
            Index           =   0
            Left            =   840
            TabIndex        =   13
            Top             =   630
            Visible         =   0   'False
            Width           =   3165
         End
         Begin VB.PictureBox Picture1 
            AutoRedraw      =   -1  'True
            Height          =   1695
            Index           =   0
            Left            =   315
            ScaleHeight     =   1635
            ScaleWidth      =   4095
            TabIndex        =   12
            Top             =   735
            Width           =   4150
         End
         Begin VB.Label Label34 
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
            Height          =   285
            Index           =   0
            Left            =   840
            TabIndex        =   15
            Top             =   315
            Width           =   2955
         End
         Begin VB.Image Image1 
            Height          =   10005
            Index           =   0
            Left            =   0
            Top             =   105
            Visible         =   0   'False
            Width           =   10005
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Croquis"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2650
         Index           =   1
         Left            =   210
         TabIndex        =   6
         Top             =   3045
         Width           =   4785
         Begin VB.FileListBox File1 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1845
            Index           =   1
            Left            =   840
            TabIndex        =   9
            Top             =   630
            Visible         =   0   'False
            Width           =   3165
         End
         Begin VB.PictureBox Picture1 
            AutoRedraw      =   -1  'True
            Height          =   1695
            Index           =   1
            Left            =   315
            ScaleHeight     =   1635
            ScaleWidth      =   4095
            TabIndex        =   8
            Top             =   735
            Width           =   4150
         End
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
            Height          =   285
            Index           =   1
            Left            =   3780
            TabIndex        =   7
            Top             =   315
            Width           =   175
         End
         Begin VB.Label Label34 
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
            Height          =   285
            Index           =   1
            Left            =   840
            TabIndex        =   10
            Top             =   315
            Width           =   2955
         End
         Begin VB.Image Image1 
            Height          =   19995
            Index           =   1
            Left            =   0
            Stretch         =   -1  'True
            Top             =   105
            Visible         =   0   'False
            Width           =   19995
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Fotografia 1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2650
         Index           =   2
         Left            =   5250
         TabIndex        =   1
         Top             =   210
         Width           =   4750
         Begin VB.FileListBox File1 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1845
            Index           =   2
            Left            =   840
            TabIndex        =   4
            Top             =   630
            Visible         =   0   'False
            Width           =   3165
         End
         Begin VB.PictureBox Picture1 
            AutoRedraw      =   -1  'True
            Height          =   1695
            Index           =   2
            Left            =   315
            ScaleHeight     =   1635
            ScaleWidth      =   4095
            TabIndex        =   3
            Top             =   735
            Width           =   4150
         End
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
            Height          =   285
            Index           =   2
            Left            =   3780
            TabIndex        =   2
            Top             =   315
            Width           =   175
         End
         Begin VB.Label Label34 
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
            Height          =   285
            Index           =   2
            Left            =   840
            TabIndex        =   5
            Top             =   315
            Width           =   2955
         End
         Begin VB.Image Image1 
            Height          =   19995
            Index           =   2
            Left            =   0
            Stretch         =   -1  'True
            Top             =   105
            Visible         =   0   'False
            Width           =   19995
         End
      End
   End
   Begin VB.Label Label38 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFF80&
      BackStyle       =   0  'Transparent
      Caption         =   "U-M:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   3465
      TabIndex        =   32
      Top             =   345
      Width           =   1005
   End
   Begin VB.Label Label37 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   4560
      TabIndex        =   31
      Top             =   240
      Width           =   540
   End
   Begin VB.Label Label36 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFF80&
      BackStyle       =   0  'Transparent
      Caption         =   "Interno:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   5145
      TabIndex        =   30
      Top             =   345
      Width           =   1005
   End
   Begin VB.Label Label28 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   6240
      TabIndex        =   29
      Top             =   240
      Width           =   750
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Descripcion:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   0
      TabIndex        =   28
      Top             =   765
      Width           =   1380
   End
   Begin VB.Label Label11 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFF80&
      BackStyle       =   0  'Transparent
      Caption         =   "Codigo:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   270
      TabIndex        =   27
      Top             =   345
      Width           =   1005
   End
   Begin VB.Label Label1 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1470
      TabIndex        =   26
      Top             =   240
      Width           =   2115
   End
   Begin VB.Label Label2 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   1470
      TabIndex        =   25
      Top             =   765
      Width           =   5520
   End
End
Attribute VB_Name = "GRAFITEM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim TOPEINI(3), IZQUINI(3), ANCHINI(3), ALTUINI(3)
Dim YACAR%, MODIFIC%
Dim GRABATODO%

Private Sub Command1_Click()
   GRABATODO% = 1
   Call GRABAPLANOS
   GRABATODO% = 0
End Sub

Private Sub Command2_Click()
   Call GRABAPLANOS
   Call CIERRARCH("*.*")
   Hide
End Sub

Private Sub Command3_Click(Index As Integer)
   RUTARGRA$ = TRIM$(CONTROL$("", "RUTARGRA"))
   If RUTARGRA$ = "" Then RUTARGRA$ = LUDAT$
   File1(Index).Path = RUTARGRA$
   '
   TIARGRA$ = "*." + TRIM$(CONTROL$("", "TIARGRA"))
   If TIARGRA$ = "*." Then TIARGRA$ = "*.*"
   If TIARGRA$ = "*.*.*" Then TIARGRA$ = "*.*"
   File1(Index).Pattern = TIARGRA$
   File1(Index).Visible = True
End Sub

Private Sub Command9_Click()
   '
   Call GRABAPLANOS
   '
   Call SELECHO("MASTER")
   CODISEL$ = TRIM$(VALACT1$("MASTER"))
   If COD$ <> "" Then
     CODXX$ = CODISEL$
     CIXXI% = CODINT%(CODISEL$)
     Label1.Caption = VALACT1$("MASTER")
     Label2.Caption = DESCRIT0$(CIXXI%)
     Label37.Caption = UNIMED$(CIXXI%)
     Label28.Caption = Str$(CIXXI%)
   End If
   '
End Sub

Private Sub File1_DblClick(Index As Integer)
   Label34(Index).Caption = File1(Index).filename
End Sub

Private Sub Form_Load()
    For KK% = 0 To 3
      TOPEINI(KK%) = Picture1(KK%).Top
      IZQUINI(KK%) = Picture1(KK%).Left
      ANCHINI(KK%) = Picture1(KK%).Width
      ALTUINI(KK%) = Picture1(KK%).Height
    Next KK%
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call GRABAPLANOS
   Call CIERRARCH("*.*")
   YACAR% = 0
End Sub

Private Sub Frame1_Click(Index As Integer)
   '
   File1(Index).Visible = False
   Picture1(Index).Visible = True
   Label34(Index).Visible = True
   Refresh
   '
End Sub

'
Private Sub Label1_Change()
   YACAR% = 0
   Call CARGAPLANOS
End Sub

Private Sub Label1_DblClick()
   '
   Call GRABAPLANOS
   '
   Call SELECHO("MASTER")
   CODISEL$ = TRIM$(VALACT1$("MASTER"))
   If COD$ <> "" Then
     CODXX$ = CODISEL$
     CIXXI% = CODINT%(CODISEL$)
     Label1.Caption = VALACT1$("MASTER")
     Label2.Caption = DESCRIT0$(CIXXI%)
     Label37.Caption = UNIMED$(CIXXI%)
     Label28.Caption = Str$(CIXXI%)
   End If
   '
End Sub

Private Sub Label34_Change(Index As Integer)
   File1(Index).Visible = False
   Call MUESTRAPLANO(Index)
   MODIFIC% = 1
End Sub

Private Sub Label34_DblClick(Index As Integer)
   Label34(Index).Caption = ""
End Sub

Private Sub Picture1_Click(Index As Integer)
      
      Screen.MousePointer = 11
      '
      Picture2.Top = 500
      Picture2.Left = 500
      Picture2.Width = 10000
      Picture2.Height = MARCOFICHA(5).Height - 800
      '
      'HIMAGEN = Picture1(Index).Picture.Width
      'VIMAGEN = Picture1(Index).Picture.Height
      HIMAGEN = Image1(Index).Picture.Width
      VIMAGEN = Image1(Index).Picture.Height
      If HIMAGEN < 1 Or VIMAGEN < 1 Then
         GoTo 99
      End If
      '
      HCUADRO = Picture2.Width
      VCUADRO = Picture2.Height
      '
      REDUH = HIMAGEN / HCUADRO
      REDUV = VIMAGEN / VCUADRO
      REDUC = REDUH: If REDUV > REDUH Then REDUC = REDUV
      Picture2.Height = VCUADRO * REDUV / REDUC
      Picture2.Width = HCUADRO * REDUH / REDUC
      HCUADRO = Picture2.Width
      VCUADRO = Picture2.Height
      Picture2.Left = (Width - HCUADRO) / 2
      '
      Picture2.PaintPicture Image1(Index).Picture, 0, 0, 1 * HCUADRO, 1 * VCUADRO
      Picture2.Visible = True
      IMACTI = Index
      '
99    Screen.MousePointer = 1
      '
End Sub

Private Sub Picture2_Click()
   Picture2.Visible = False
End Sub

Sub CARGAPLANOS()
   '
   For KK% = 0 To 3
      Label34(KK%).Caption = ""
      Picture1(KK%).Top = TOPEINI(KK%)
      Picture1(KK%).Left = IZQUINI(KK%)
      Picture1(KK%).Width = ANCHINI(KK%)
      Picture1(KK%).Height = ALTUINI(KK%)
      Picture1(KK%).Picture = LoadPicture("")
   Next KK%
   '
   CODXX$ = Label1.Caption
   CIXXI% = CODINT%(CODXX$)
   If CIXXI% > 0 Then
      If CIXXI% <= NUMAS% Then
         Screen.MousePointer = 11
         RNC$ = RECFILT$("INDIPLAN", 1, MKI$(CIXXI%))
         If Len(RNC$) >= 4 Then
            REPLANO& = CVS(RNC$)
            If REPLANO& > 0 Then
               If REPLANO& <= ULTREG&("INDIPLAN") Then
                  RPLA$ = REGLEIDO$("INDIPLAN", REPLANO&)
                  Label34(0).Caption = TRIM$(Mid$(RPLA$, 65, 48))
                  Label34(1).Caption = TRIM$(Mid$(RPLA$, 113, 48))
                  Label34(2).Caption = TRIM$(Mid$(RPLA$, 161, 48))
                  Label34(3).Caption = TRIM$(Mid$(RPLA$, 209, 48))
               End If
            End If
         End If
         'For KK% = 0 To 3
         '   Call MUESTRAPLANO(KK%)
         'Next KK%
         Screen.MousePointer = 1
      End If
   End If
   YACAR% = 1
   MODIFIC% = 0
   '
End Sub
'
Sub GRABAPLANOS()
   '
   CODXX$ = Label1.Caption
   CIXXI% = CODINT%(CODXX$)
   If CIXXI% < 1 Or CIXXI% > NUMAS% Then Exit Sub
   If YACAR% < 1 Then Exit Sub
   If MODIFIC% < 1 Then Exit Sub
   '
   If GRABATODO% <> 1 Then
     If COMALTER%("Se ha Modificado la Información Gráfica del Artículo\'" + TRIM$(CODEXT$(CIXXI%)) + " - " + TRIM$(Label2.Caption) + "'\\Registrar Cambios\Cancelar") <> 1 Then
       YACAR% = 0
       MODIFIC% = 0
       Exit Sub
     End If
   End If
   '
   Screen.MousePointer = 11
   RNC$ = RECFILT$("INDIPLAN", 1, MKI$(CIXXI%))
   REPLANO& = 0
   If Len(RNC$) >= 4 Then REPLANO& = CVS(RNC$)
   '
   If REPLANO& < 1 Or REPLANO > ULTREG&("INDIPLAN") Then
      REPLANO& = 1 + ULTREG&("INDIPLAN")
   End If
   RPLA$ = REGBLAN$("INDIPLAN")
   Call REPLA(RPLA$, MKI$(CIXXI%), 1, 2)
   Call REPLA(RPLA$, TRIM$(Label34(0).Caption), 65, 48)
   Call REPLA(RPLA$, TRIM$(Label34(1).Caption), 113, 48)
   Call REPLA(RPLA$, TRIM$(Label34(2).Caption), 161, 48)
   Call REPLA(RPLA$, TRIM$(Label34(3).Caption), 209, 48)
   Call GRAREG("INDIPLAN", RPLA$, REPLANO&)
   '
   MODIFIC% = 0
   Screen.MousePointer = 1
   '
End Sub

Sub MUESTRAPLANO(Index)
   '
   Screen.MousePointer = 11
   '
   RUTARGRA$ = TRIM$(CONTROL$("", "RUTARGRA"))
   If RUTARGRA$ = "" Then RUTARGRA$ = TRIM$(LUDAT$)
   If Right$(RUTARGRA$, 1) <> "\" Then
      RUTARGRA$ = RUTARGRA$ + "\"
   End If
   '
   File1(Index).Visible = False
   Picture1(Index).Visible = True
   Label34(Index).Visible = True
   Refresh
   '
   Image1(Index).Picture = LoadPicture("")
   Picture1(Index).Cls
   COPLANO$ = RUTARGRA$ + TRIM$(Label34(Index).Caption)
   If COPLANO$ <> "" Then
      If EXISTE%(COPLANO$) > 0 Then
         On Error GoTo 90
         'Picture1(Index).Visible = False
         Image1(Index).Picture = LoadPicture(COPLANO$)
         Picture1(Index).Top = TOPEINI(Index)
         Picture1(Index).Left = IZQUINI(Index)
         Picture1(Index).Width = ANCHINI(Index)
         Picture1(Index).Height = ALTUINI(Index)
         Picture1(Index).Picture = LoadPicture("")
         Picture1(Index).Refresh
         '
         HIMAGEN = Image1(Index).Picture.Width
         VIMAGEN = Image1(Index).Picture.Height
         
         HCUADRO = Picture1(Index).Width
         VCUADRO = Picture1(Index).Height
         '
         REDUH = HIMAGEN / HCUADRO
         REDUV = VIMAGEN / VCUADRO
         REDUC = REDUH: If REDUV > REDUH Then REDUC = REDUV
         Picture1(Index).Height = VCUADRO * REDUV / REDUC
         Picture1(Index).Width = HCUADRO * REDUH / REDUC
         Picture1(Index).Refresh
         '
         HCUADRO = Picture1(Index).Width
         VCUADRO = Picture1(Index).Height
         Picture1(Index).Left = (Frame1(Index).Width - HCUADRO) / 2
         '
         Picture1(Index).PaintPicture Image1(Index).Picture, 0, 0, HCUADRO, VCUADRO
         Picture1(Index).Visible = True
         Picture1(Index).Refresh
       Else
         'Call MENSERR(24, "No Existe Archivo '" + COPLANO$ + "'.")
      End If
   End If
   GoTo 98
   '
90 On Error GoTo 0
   Call MENSERR(24, "Imagen no Válida")
   '
98 Screen.MousePointer = 1
   '
99 End Sub



