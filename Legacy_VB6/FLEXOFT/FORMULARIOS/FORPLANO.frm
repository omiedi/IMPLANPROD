VERSION 5.00
Begin VB.Form FORPLANO 
   BackColor       =   &H00C0E0E0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Planera Electronica"
   ClientHeight    =   5745
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   7935
   Icon            =   "FORPLANO.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5745
   ScaleWidth      =   7935
   ShowInTaskbar   =   0   'False
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
      Left            =   7140
      Picture         =   "FORPLANO.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   23
      ToolTipText     =   "Configuración - Tablas de Validación"
      Top             =   4305
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
      Left            =   7140
      Picture         =   "FORPLANO.frx":074C
      Style           =   1  'Graphical
      TabIndex        =   22
      ToolTipText     =   "Ayuda en Línea"
      Top             =   210
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
      Left            =   7140
      Picture         =   "FORPLANO.frx":0896
      Style           =   1  'Graphical
      TabIndex        =   21
      ToolTipText     =   "Ayuda en Línea"
      Top             =   900
      Width           =   650
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
      Left            =   7140
      Picture         =   "FORPLANO.frx":0BA0
      Style           =   1  'Graphical
      TabIndex        =   20
      ToolTipText     =   "Guarda los Cambios Realizados"
      Top             =   1590
      Width           =   650
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0E0E0&
      Caption         =   "Fotografia-1"
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
      Left            =   3675
      TabIndex        =   15
      Top             =   105
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
         TabIndex        =   16
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
         TabIndex        =   18
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
         TabIndex        =   17
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
         TabIndex        =   19
         Top             =   315
         Width           =   2755
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0E0E0&
      Caption         =   "Plano-2"
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
      TabIndex        =   10
      Top             =   2940
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
         TabIndex        =   11
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
         TabIndex        =   13
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
         TabIndex        =   12
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
         TabIndex        =   14
         Top             =   315
         Width           =   2755
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0E0E0&
      Caption         =   "Plano-1"
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
      TabIndex        =   5
      Top             =   105
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
         TabIndex        =   6
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
         TabIndex        =   8
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
         TabIndex        =   7
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
         TabIndex        =   9
         Top             =   315
         Width           =   2755
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0E0E0&
      Caption         =   "Fotografia-2"
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
      Left            =   3675
      TabIndex        =   0
      Top             =   2940
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
         TabIndex        =   1
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
         TabIndex        =   3
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
         Index           =   3
         Left            =   210
         TabIndex        =   4
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
   Begin VB.Image Image3 
      Height          =   390
      Left            =   6720
      Picture         =   "FORPLANO.frx":0EAA
      Top             =   5190
      Width           =   1515
   End
End
Attribute VB_Name = "FORPLANO"
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

Private Sub Command3_Click()
   Call CIERRARCH("*.*")
   Hide
End Sub

Private Sub Command37_Click()
   Command37.Enabled = False
   FORILOGO.Show 1
   Command37.Enabled = True
End Sub

Private Sub Command39_Click()

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
   Label34(Index).Caption = File1(Index).filename
   File1(Index).Visible = False
End Sub

Private Sub Label34_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   If Button = 2 Then
     If COMALTER%("Confirma Desconectar Esta Imagen ?\\No, Conservar\Si,Desconectar") = 2 Then
       Label34(Index) = ""
     End If
   End If
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
      SHOWGRAF.Caption = TRIM$(ALTAMODE.Text3) + " - " + TRIM$(ALTAMODE.Text2)
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
    Screen.MousePointer = 1
    '
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call CIERRARCH("*.*")
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

Sub CARPLANOS()
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
   CODXX$ = AJUSTI$(ALTAMODE.Text3, 16)
   Screen.MousePointer = 11
   RNC$ = RECFILT$("PLANOMOD", 1, CODXX$)
   For U& = 1 To Len(RNC$) Step 4
     REPLANO& = CVS(Mid$(RNC$, U, 4))
     If REPLANO& > 0 Then
       If REPLANO& <= ULTREG&("PLANOMOD") Then
         RPLA$ = Mid$(REGLEIDO$("PLANOMOD", REPLANO&), 17)
         IPLA% = (U& - 1) / 4
         Label34(IPLA%).Caption = RPLA$
       End If
     End If
   Next U&
   Screen.MousePointer = 1
   '
End Sub
'
Sub GRAPLANOS()
   '
   Screen.MousePointer = 11
   CODEMODE$ = AJUSTI$(ALTAMODE.Text3, 16)
   '
   If CODEMODE$ <> "" Then
     If ECOARCH("BADAMODE", CODEMODE$) <> "" Then
       For KKI% = 0 To 3
         REPLAX$ = Space$(128)
         Call REPLA(REPLAX$, CODEMODE, 1, 16)
         Call REPLA(REPLAX$, Label34(KKI%), 17, 112)
         JJ& = KKI% + 1
         Call GRAREG("!PLANOMOD", REPLAX$, JJ&)
       Next KKI%
       Call SETULTREG("!PLANOMOD", 4)
       '
       Call FILTERUPDATE("PLANOMOD", "", 1, CODEMODE$)
     End If
   End If
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


