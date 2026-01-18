VERSION 5.00
Begin VB.Form FORPLAN04 
   BackColor       =   &H00C0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Planera Electronica"
   ClientHeight    =   7110
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   8370
   Icon            =   "FORPLAN04.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   7110
   ScaleWidth      =   8370
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00008080&
      Height          =   7155
      Left            =   7485
      ScaleHeight     =   7095
      ScaleWidth      =   1005
      TabIndex        =   29
      Top             =   0
      Width           =   1065
      Begin VB.CommandButton Command2 
         Caption         =   "Command2"
         Height          =   855
         Left            =   105
         TabIndex        =   35
         Top             =   2625
         Width           =   540
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
         Height          =   690
         Left            =   105
         Picture         =   "FORPLAN04.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   34
         ToolTipText     =   "Guarda los Cambios Realizados"
         Top             =   1650
         Width           =   650
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
         Height          =   690
         Left            =   105
         Picture         =   "FORPLAN04.frx":0614
         Style           =   1  'Graphical
         TabIndex        =   33
         ToolTipText     =   "Ayuda en Línea"
         Top             =   960
         Width           =   650
      End
      Begin VB.CommandButton Command3 
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
         Left            =   105
         Picture         =   "FORPLAN04.frx":091E
         Style           =   1  'Graphical
         TabIndex        =   32
         ToolTipText     =   "Ayuda en Línea"
         Top             =   270
         Width           =   650
      End
      Begin VB.CommandButton Command37 
         Caption         =   "Setup"
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
         Left            =   105
         Picture         =   "FORPLAN04.frx":0A68
         Style           =   1  'Graphical
         TabIndex        =   31
         ToolTipText     =   "Configuración - Tablas de Validación"
         Top             =   4410
         Width           =   650
      End
      Begin VB.CommandButton Command39 
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
         Height          =   690
         Left            =   105
         Picture         =   "FORPLAN04.frx":0EAA
         Style           =   1  'Graphical
         TabIndex        =   30
         ToolTipText     =   "Imprime Ficha del Artículo"
         Top             =   5100
         Width           =   650
      End
      Begin VB.Image Image3 
         Height          =   390
         Left            =   -315
         Picture         =   "FORPLAN04.frx":1514
         Top             =   6510
         Width           =   1515
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0E0E0&
      Caption         =   "Fotografia"
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
      Left            =   3780
      TabIndex        =   24
      Top             =   1470
      Width           =   3300
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
         Left            =   210
         Pattern         =   "*.JPG"
         TabIndex        =   25
         Top             =   630
         Visible         =   0   'False
         Width           =   2755
      End
      Begin VB.CommandButton Command53 
         Caption         =   "."
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
         Left            =   2950
         TabIndex        =   27
         Top             =   315
         Width           =   175
      End
      Begin VB.PictureBox Picture1 
         AutoRedraw      =   -1  'True
         Height          =   1695
         Index           =   2
         Left            =   210
         ScaleHeight     =   1635
         ScaleWidth      =   2850
         TabIndex        =   26
         ToolTipText     =   "Click para Maximizar"
         Top             =   735
         Width           =   2905
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
         Left            =   210
         TabIndex        =   28
         Top             =   315
         Width           =   2755
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0E0E0&
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
      Left            =   315
      TabIndex        =   19
      Top             =   4305
      Width           =   3300
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
         Left            =   210
         Pattern         =   "*.JPG"
         TabIndex        =   20
         Top             =   630
         Visible         =   0   'False
         Width           =   2755
      End
      Begin VB.CommandButton Command53 
         Caption         =   "."
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
         Left            =   2950
         TabIndex        =   22
         Top             =   315
         Width           =   175
      End
      Begin VB.PictureBox Picture1 
         AutoRedraw      =   -1  'True
         Height          =   1695
         Index           =   1
         Left            =   210
         ScaleHeight     =   1635
         ScaleWidth      =   2850
         TabIndex        =   21
         ToolTipText     =   "Click para Maximizar"
         Top             =   735
         Width           =   2905
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
         Left            =   210
         TabIndex        =   23
         Top             =   315
         Width           =   2755
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0E0E0&
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
      Left            =   315
      TabIndex        =   14
      Top             =   1470
      Width           =   3300
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
         Left            =   210
         Pattern         =   "*.JPG"
         TabIndex        =   15
         Top             =   630
         Visible         =   0   'False
         Width           =   2755
      End
      Begin VB.PictureBox Picture1 
         AutoRedraw      =   -1  'True
         Height          =   1695
         Index           =   0
         Left            =   210
         ScaleHeight     =   1635
         ScaleWidth      =   2850
         TabIndex        =   17
         ToolTipText     =   "Click para Maximizar"
         Top             =   735
         Width           =   2905
      End
      Begin VB.CommandButton Command53 
         Caption         =   "."
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
         Left            =   2950
         TabIndex        =   16
         Top             =   315
         Width           =   175
      End
      Begin VB.Image Image1 
         Height          =   10005
         Index           =   0
         Left            =   0
         Top             =   105
         Visible         =   0   'False
         Width           =   10005
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
         Left            =   210
         TabIndex        =   18
         Top             =   315
         Width           =   2755
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0E0E0&
      Caption         =   "Fotografia"
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
      Left            =   3780
      TabIndex        =   9
      Top             =   4305
      Width           =   3300
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
         Left            =   210
         Pattern         =   "*.JPG"
         TabIndex        =   10
         Top             =   630
         Visible         =   0   'False
         Width           =   2755
      End
      Begin VB.PictureBox Picture1 
         AutoRedraw      =   -1  'True
         Height          =   1695
         Index           =   3
         Left            =   210
         ScaleHeight     =   1635
         ScaleWidth      =   2850
         TabIndex        =   12
         ToolTipText     =   "Click para Maximizar"
         Top             =   735
         Width           =   2905
      End
      Begin VB.CommandButton Command53 
         Caption         =   "."
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
         Left            =   2950
         TabIndex        =   11
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
         Index           =   3
         Left            =   210
         TabIndex        =   13
         Top             =   315
         Width           =   2755
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
   End
   Begin VB.TextBox Text1 
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
      Left            =   1575
      MaxLength       =   15
      TabIndex        =   1
      Top             =   315
      Width           =   2040
   End
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
      Left            =   3600
      Picture         =   "FORPLAN04.frx":3436
      TabIndex        =   0
      Top             =   315
      Width           =   175
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
      Left            =   1575
      TabIndex        =   8
      Top             =   840
      Width           =   5475
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
      TabIndex        =   7
      Top             =   420
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
      TabIndex        =   6
      Top             =   315
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
      TabIndex        =   5
      Top             =   420
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
      Left            =   6300
      TabIndex        =   4
      Top             =   315
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
      Left            =   105
      TabIndex        =   3
      Top             =   840
      Width           =   1380
   End
   Begin VB.Label Label1 
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
      Left            =   480
      TabIndex        =   2
      Top             =   420
      Width           =   1005
   End
End
Attribute VB_Name = "FORPLAN04"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim TOPEINI(3), IZQUINI(3), ANCHINI(3), ALTUINI(3)

Private Sub Command1_Click()
   Command1.Enabled = False
   Call GRAPLANOS
   Command1.Enabled = True
End Sub

Private Sub Command2_Click()
Pdf1.LOADFILE "C:\Mis Documentos\contab04.pdf"

End Sub

Private Sub Command3_Click()
   Call CIERRARCH("*.*")
   Call FINAL("")
End Sub

Private Sub Command37_Click()
   Command37.Enabled = False
   FORILOGO.Show 1
   Command37.Enabled = True
End Sub

Private Sub Command53_Click(Index As Integer)
   RUTARGRA$ = TRIM$(CONTROL$("", "RUTARGRA"))
   If RUTARGRA$ = "" Then RUTARGRA$ = LUDAT$
   On Error Resume Next
   File1(Index).Path = RUTARGRA$
   If Err Then
     RTG$ = REPLACAR$(RUTARGRA$, "\", "/")
     Call MENSERR(24, "Ruta '" + RTG$ + "' no Encontrada.")
     Exit Sub
   End If
   '
   TIARGRA$ = "*." + TRIM$(CONTROL$("", "TIARGRA"))
   If TIARGRA$ = "*." Then TIARGRA$ = "*.*"
   If TIARGRA$ = "*.*.*" Then TIARGRA$ = "*.*"
   File1(Index).Pattern = TIARGRA$
   File1(Index).Visible = True
   If Err Then
     RTG$ = REPLACAR$(TIARGRA$, "\", "/")
     Call MENSERR(24, "Máscara '" + RTG$ + "' no Válida.")
     Exit Sub
   End If
   On Error GoTo 0
   '
End Sub
'
Private Sub Command9_Click()
   '
   Call SELECHO("MASTER")
   COD$ = TRIM$(VALACT1$("MASTER"))
   If COD$ <> "" Then
     Text1.TEXT = COD$
   End If
   '
End Sub

Private Sub File1_DblClick(Index As Integer)
   Label34(Index).Caption = File1(Index).FileName
   File1(Index).Visible = False
End Sub

Private Sub Picture1_Click(Index As Integer)
      '
      Call Command1_Click
      '
      Screen.MousePointer = 11
      '
      ALTUTITU = SHOWGRAF.Height - SHOWGRAF.ScaleHeight
      ANCHOBOR = SHOWGRAF.Width - SHOWGRAF.ScaleWidth
      SHOWGRAF.Picture2.Top = 0
      SHOWGRAF.Picture2.Left = 0
      SHOWGRAF.Picture2.Width = Screen.Width - ANCHOBOR
      SHOWGRAF.Picture2.Height = Screen.Height - ALTUTITU - 300
      '
      'HIMAGEN = Picture1(Index).Picture.Width
      'VIMAGEN = Picture1(Index).Picture.Height
      HIMAGEN = Image1(Index).Picture.Width
      VIMAGEN = Image1(Index).Picture.Height
      If HIMAGEN < 1 Or VIMAGEN < 1 Then
         GoTo 99
      End If
      '
      HCUADRO = SHOWGRAF.Picture2.Width
      VCUADRO = SHOWGRAF.Picture2.Height
      '
      REDUH = HIMAGEN / HCUADRO
      REDUV = VIMAGEN / VCUADRO
      REDUC = REDUH: If REDUV > REDUH Then REDUC = REDUV
      SHOWGRAF.Picture2.Height = VCUADRO * REDUV / REDUC
      SHOWGRAF.Picture2.Width = HCUADRO * REDUH / REDUC
      HCUADRO = SHOWGRAF.Picture2.Width
      VCUADRO = SHOWGRAF.Picture2.Height
      SHOWGRAF.Height = ALTUTITU + SHOWGRAF.Picture2.Height
      SHOWGRAF.Width = ANCHOBOR + SHOWGRAF.Picture2.Width
      If HORIMAGEN + SHOWGRAF.Width > Screen.Width Then HORIMAGEN = 0
      If HORIMAGEN < 1 Then HORIMAGEN = (Screen.Width - SHOWGRAF.Width) / 2
      SHOWGRAF.Left = HORIMAGEN
      If VERTIMAGEN + SHOWGRAF.Height > Screen.Height Then VERTIMAGEN = 0
      If VERTIMAGEN < 1 Then VERTIMAGEN = (Screen.Height - SHOWGRAF.Height) / 2
      SHOWGRAF.Top = VERTIMAGEN - 300
        If SHOWGRAF.Top < 0 Then SHOWGRAF.Top = 0
      SHOWGRAF.Caption = TRIM$(Text1.TEXT) + " - " + TRIM$(Label2)
      Screen.MousePointer = 1
      '
      INDIMAGEN% = Index
      SHOWGRAF.Picture2.PaintPicture Image1(Index).Picture, 0, 0, 1 * HCUADRO, 1 * VCUADRO
      SHOWGRAF.Picture2.Visible = True
      SHOWGRAF.Show 1
      '
99    Screen.MousePointer = 1
      '
End Sub

Private Sub Form_Load()
    '
    Call CENTRAFORM(Me)
    Call TRANSLABELS(Name)
    DoEvents
    '
    Call VERJOBID(OKEY%)
    If OKEY% < 1 Then Call FINAL("")
    '
    Screen.MousePointer = 11
    '
    EPAC$ = TRIM$(EMPREAC$)
    If EPAC$ <> "" Then
        Caption = Caption + "  -  " + EPAC$
    End If
    Call GRATITFOR(Caption)
    '
    For KK% = 0 To 3
      TOPEINI(KK%) = Picture1(KK%).Top
      IZQUINI(KK%) = Picture1(KK%).Left
      ANCHINI(KK%) = Picture1(KK%).Width
      ALTUINI(KK%) = Picture1(KK%).Height
    Next KK%
    '
    Text1 = APLINVO$
    Screen.MousePointer = 1
    '
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call CIERRARCH("*.*")
   Call FINAL("")
End Sub

Private Sub Label34_Change(Index As Integer)
   File1(Index).Visible = False
   Call MUESTRAPLANO(Index)
End Sub

Private Sub Label34_DblClick(Index As Integer)
   '
   RUTARGRA$ = TRIM$(CONTROL$("", "RUTARGRA"))
   If RUTARGRA$ = "" Then RUTARGRA$ = LUDAT$
   On Error Resume Next
   File1(Index).Path = RUTARGRA$
   If Err Then
     RTG$ = REPLACAR$(RUTARGRA$, "\", "/")
     Call MENSERR(24, "Ruta '" + RTG$ + "' no Encontrada.")
     Exit Sub
   End If
   '
   TIARGRA$ = "*." + TRIM$(CONTROL$("", "TIARGRA"))
   If TIARGRA$ = "*." Then TIARGRA$ = "*.*"
   If TIARGRA$ = "*.*.*" Then TIARGRA$ = "*.*"
   File1(Index).Pattern = TIARGRA$
   File1(Index).Visible = True
   If Err Then
     RTG$ = REPLACAR$(TIARGRA$, "\", "/")
     Call MENSERR(24, "Máscara '" + RTG$ + "' no Válida.")
     Exit Sub
   End If
   On Error GoTo 0
   '
End Sub

Private Sub Text1_Change()
   '
   Call BLANFORMU
   If CODINT%(Text1.TEXT) > 0 Then
       Call CARDATITEM
       Call CARPLANOS
   End If
   '
End Sub

Sub BLANFORMU()
   '
   Label2 = ""
   Label37.Caption = ""
   Label28.Caption = ""
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
End Sub

Sub CARDATITEM()
    '
    CODD$ = Text1.TEXT
    CI% = CODINT%(CODD$)
    If CI% < 1 Then
       If TRIM$(CODD$) <> "" Then
          Call MENSERR(24, "Código Inexistente o no Válido")
       End If
       Exit Sub
    End If
    '
    Screen.MousePointer = 11
    Label2 = DESCRIT0$(CI%)
    Label37.Caption = UNIMED$(CI%)
    Label28.Caption = Str$(CI%)
    Screen.MousePointer = 1
    '
End Sub

Sub CARPLANOS()
   '
   For KK% = 0 To 3
      'Label34(KK%).Caption = ""
      Picture1(KK%).Top = TOPEINI(KK%)
      Picture1(KK%).Left = IZQUINI(KK%)
      Picture1(KK%).Width = ANCHINI(KK%)
      Picture1(KK%).Height = ALTUINI(KK%)
      Picture1(KK%).Picture = LoadPicture("")
   Next KK%
   '
   CODXX$ = Text1.TEXT
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
   '
End Sub
'
Sub GRAPLANOS()
   '
   CODXX$ = Text1.TEXT
   CIXXI% = CODINT%(CODXX$)
   If CIXXI% < 1 Or CIXXI% > NUMAS% Then Exit Sub
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
   Call CIERRARCH("INDIPLAN")
   '
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
   'File1(Index).Visible = False
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


