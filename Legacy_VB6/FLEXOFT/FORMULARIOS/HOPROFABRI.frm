VERSION 5.00
Begin VB.Form HOPROFABRI 
   BackColor       =   &H00FFFFC0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Hoja de Procesos de Fabricación"
   ClientHeight    =   8655
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   11910
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8655
   ScaleWidth      =   11910
   Begin VB.TextBox Text8 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   5775
      MaxLength       =   16
      TabIndex        =   41
      Top             =   2055
      Width           =   2010
   End
   Begin VB.PictureBox Picture4 
      AutoRedraw      =   -1  'True
      Height          =   1590
      Left            =   -210
      ScaleHeight     =   1530
      ScaleWidth      =   2055
      TabIndex        =   23
      Top             =   5355
      Visible         =   0   'False
      Width           =   2115
   End
   Begin VB.TextBox Text7 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   9345
      MaxLength       =   8
      TabIndex        =   11
      Top             =   3885
      Width           =   1380
   End
   Begin VB.TextBox Text6 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   9345
      MaxLength       =   8
      TabIndex        =   10
      Top             =   3360
      Width           =   1380
   End
   Begin VB.TextBox Text5 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   9345
      MaxLength       =   8
      TabIndex        =   9
      Top             =   2835
      Width           =   1380
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   6930
      MaxLength       =   4
      TabIndex        =   2
      Top             =   420
      Width           =   855
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1485
      Left            =   1680
      MultiLine       =   -1  'True
      TabIndex        =   5
      Top             =   2835
      Width           =   6120
   End
   Begin VB.CommandButton Command7 
      Caption         =   " "
      Height          =   260
      Left            =   1680
      TabIndex        =   33
      Top             =   8190
      Width           =   4100
   End
   Begin VB.CommandButton Command26 
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
      Height          =   855
      Left            =   10920
      Picture         =   "HOPROFABRI.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   2650
      Width           =   855
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   225
      Left            =   0
      TabIndex        =   30
      Top             =   5040
      Visible         =   0   'False
      Width           =   1485
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   1170
      Left            =   420
      TabIndex        =   29
      Top             =   5985
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.CommandButton Command12 
      Caption         =   "Selecc Todo"
      Height          =   260
      Left            =   9435
      TabIndex        =   27
      Top             =   8190
      Width           =   1290
   End
   Begin VB.CommandButton Command11 
      Caption         =   "Pegar"
      Height          =   260
      Left            =   8145
      TabIndex        =   26
      Top             =   8190
      Width           =   1275
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Copiar"
      Height          =   260
      Left            =   6975
      TabIndex        =   25
      Top             =   8190
      Width           =   1170
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Cortar"
      Height          =   260
      Left            =   5790
      TabIndex        =   24
      Top             =   8190
      Width           =   1170
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Lay - Out"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2535
      Index           =   0
      Left            =   7980
      TabIndex        =   21
      Top             =   105
      Width           =   2790
      Begin VB.FileListBox File1 
         Height          =   1455
         Index           =   0
         Left            =   210
         TabIndex        =   8
         Top             =   840
         Visible         =   0   'False
         Width           =   2400
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
         Index           =   0
         Left            =   2415
         TabIndex        =   7
         Top             =   315
         Width           =   175
      End
      Begin VB.PictureBox Picture1 
         AutoRedraw      =   -1  'True
         Height          =   1485
         Index           =   0
         Left            =   210
         ScaleHeight     =   1425
         ScaleWidth      =   2340
         TabIndex        =   22
         Top             =   840
         Width           =   2400
      End
      Begin VB.TextBox Text4 
         Height          =   285
         Index           =   0
         Left            =   210
         TabIndex        =   6
         Text            =   " "
         Top             =   315
         Width           =   2400
      End
      Begin VB.Image Image1 
         Height          =   10005
         Index           =   0
         Left            =   -105
         Top             =   1260
         Visible         =   0   'False
         Width           =   10005
      End
   End
   Begin VB.TextBox Text3 
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3480
      Left            =   1680
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   20
      Top             =   4725
      Width           =   9060
   End
   Begin VB.CommandButton Command5 
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
      Height          =   800
      Left            =   10920
      Picture         =   "HOPROFABRI.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   16
      ToolTipText     =   "Imprime Ficha de Ingenieria"
      Top             =   6930
      Width           =   855
   End
   Begin VB.CommandButton Command2 
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
      Height          =   800
      Left            =   10920
      Picture         =   "HOPROFABRI.frx":0AAC
      Style           =   1  'Graphical
      TabIndex        =   14
      ToolTipText     =   "Ayuda FLEXOFT en Línea"
      Top             =   1785
      Width           =   855
   End
   Begin VB.CommandButton Command9 
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
      Height          =   800
      Left            =   10920
      Picture         =   "HOPROFABRI.frx":0DB6
      Style           =   1  'Graphical
      TabIndex        =   12
      ToolTipText     =   "Sale de la Aplicación"
      Top             =   105
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "O.K."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   800
      Left            =   10920
      Picture         =   "HOPROFABRI.frx":0F00
      Style           =   1  'Graphical
      TabIndex        =   13
      ToolTipText     =   "Graba las Modificaciones"
      Top             =   945
      Width           =   855
   End
   Begin VB.Line Line3 
      X1              =   1680
      X2              =   7800
      Y1              =   4320
      Y2              =   4320
   End
   Begin VB.Line Line2 
      X1              =   7800
      X2              =   7800
      Y1              =   2625
      Y2              =   4305
   End
   Begin VB.Line Line1 
      X1              =   10740
      X2              =   10740
      Y1              =   4490
      Y2              =   8430
   End
   Begin VB.Image Image2 
      Height          =   510
      Left            =   10350
      Picture         =   "HOPROFABRI.frx":120A
      Top             =   7930
      Width           =   2010
   End
   Begin VB.Label Label17 
      BackStyle       =   0  'Transparent
      Caption         =   "Nro H.P.F::"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   5355
      TabIndex        =   42
      Top             =   1785
      Width           =   1485
   End
   Begin VB.Label CONTROCAM 
      BackStyle       =   0  'Transparent
      Caption         =   "Label17"
      Height          =   330
      Left            =   105
      TabIndex        =   40
      Top             =   105
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.Label Label16 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Fecha:"
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
      Left            =   7770
      TabIndex        =   39
      Top             =   3990
      Width           =   1485
   End
   Begin VB.Label Label15 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Preparó:"
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
      Left            =   7770
      TabIndex        =   38
      Top             =   3465
      Width           =   1485
   End
   Begin VB.Label Label14 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Pzas/Hora:"
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
      Left            =   7770
      TabIndex        =   37
      Top             =   2940
      Width           =   1485
   End
   Begin VB.Label Label13 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Rev:"
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
      Left            =   6195
      TabIndex        =   36
      Top             =   525
      Width           =   645
   End
   Begin VB.Label Label12 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Mecanizado:"
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
      Left            =   -105
      TabIndex        =   35
      Top             =   2835
      Width           =   1695
   End
   Begin VB.Label Label11 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Fase  TNETO   mm/REV  RPM   Observaciones"
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   280
      Left            =   1680
      TabIndex        =   34
      Top             =   2580
      Width           =   6105
   End
   Begin VB.Label Label10 
      BackColor       =   &H80000016&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   555
      Left            =   1680
      TabIndex        =   4
      Top             =   1785
      Width           =   3480
   End
   Begin VB.Label Label9 
      BackColor       =   &H80000016&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   555
      Left            =   1680
      TabIndex        =   3
      Top             =   1050
      Width           =   6105
   End
   Begin VB.Label Label8 
      BackColor       =   &H80000016&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   390
      Left            =   5355
      TabIndex        =   1
      Top             =   420
      Width           =   645
   End
   Begin VB.Label Label7 
      BackColor       =   &H80000016&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   390
      Left            =   1680
      TabIndex        =   0
      Top             =   420
      Width           =   2430
   End
   Begin VB.Label Label6 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Item  Descripcion de la Operacion                     Maquina  Herram.  "
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   280
      Left            =   1680
      TabIndex        =   32
      Top             =   4470
      Width           =   9045
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Operación:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   105
      TabIndex        =   31
      Top             =   1890
      Width           =   1485
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Op.Nro:"
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
      Left            =   4305
      TabIndex        =   28
      Top             =   525
      Width           =   960
   End
   Begin VB.Label Label1 
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
      Left            =   -105
      TabIndex        =   19
      Top             =   4725
      Width           =   1695
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Código:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   315
      TabIndex        =   18
      Top             =   525
      Width           =   1275
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Descripcion Corta:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   105
      TabIndex        =   17
      Top             =   1050
      Width           =   1485
   End
End
Attribute VB_Name = "HOPROFABRI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public CI%, CABAS, CODE$, CIE%
Dim HIDOCLICK As Integer
Dim TOPEINI(3), IZQUINI(3), ANCHINI(3), ALTUINI(3)
Dim IMACTI
Dim MODIFJ%
'
Private Sub Command1_Click()
    Call GRADESOPER
    Call GRADATMECA
    Call GRADATOPER
    MODIFJ% = 0
End Sub

Private Sub Command11_Click()
   COMIE = Text3.SelStart
   LOSEL = Text3.SelLength
   Text3.TEXT = Left$(Text3.TEXT, COMIE) + Clipboard.GetText + Mid$(Text3.TEXT, COMIE + LOSEL + 1)
End Sub

Private Sub Command12_Click()
   Text3.SelStart = 0
   Text3.SelLength = Len(Text3.TEXT)
   Text3.Refresh
End Sub

Private Sub Command26_Click()
   Command26.Enabled = False
   OPGRAFIC.Show 1
   Command26.Enabled = True
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

Private Sub Command4_Click()
   COMIE = Text3.SelStart
   LOSEL = Text3.SelLength
   Text3.TEXT = Left$(Text3.TEXT, COMIE) + Mid$(Text3.TEXT, COMIE + LOSEL + 1)
End Sub

Private Sub Command5_Click()
   PRIFORM$ = TRIM$(CONTROL$("FORGRAFI", "DOCUGRA2"))
   If PRIFORM$ <> "" Then
     If EXISTE%(LUDAT$ + PRIFORM$ + ".FRM") Then
       CIXI% = CODINT%(Label7.Caption)
       NOPE% = Val(Label8.Caption)
       Call PRIHOJPROFAB(CIXI%, NOPE%)
       Exit Sub
     End If
   End If
   '
   ORIENTANT = Printer.Orientation
   Printer.Orientation = 2
   PRINTFORM
   Printer.Orientation = ORIENTANT
End Sub
'

Private Sub Command8_Click()
   Clipboard.SetText Mid$(Text3.TEXT, Text3.SelStart + 1, Text3.SelLength)
End Sub

Private Sub CONTROCAM_Change()
   Label7.Caption = FICHINGEPRO.Text1.TEXT
   Label9.Caption = FICHINGEPRO.Text2.TEXT
   Label8.Caption = TRIM$(Left$(FICHINGEPRO.List1.TEXT, 6))
   Label10.Caption = TRIM$(Mid$(FICHINGEPRO.List1.TEXT, 7))
   Call CARDESOPER
   Call CARDATMECA
   Call CARDATOPER
   Call MUESTRAPLANO(0)
   MODIFJ% = 0
End Sub

Private Sub File1_DblClick(Index As Integer)
   Text4(Index).TEXT = File1(Index).filename
   File1(Index).Visible = False
   Call MUESTRAPLANO(Index)
End Sub

Private Sub Form_Load()
    '
    EPAC$ = TRIM$(EMPREAC$)
    If EPAC$ <> "" Then
        Caption = Caption + "  -  " + EPAC$
    End If
    '
    For KK% = 0 To 0
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
    INI& = 1 + ULTREG&("ESTRUDOS")
    FIN& = ULTREG&("ESTRUNUE")
    For U& = INI& To FIN&
        RB$ = String$(256, 0)
        Mid$(RB$, 251, 4) = MKS$(0.0001)
        Call GRAREG("ESTRUDOS", RB$, U&)
    Next U&
    '
    Cancel = 0
    Call CIERRARCH("*.*")
    Call FINAL("")
End Sub

Private Sub HScroll1_Change()
     HCUADRO = Picture4.Width
     VCUADRO = Picture4.Height
     Picture4.PaintPicture Image1(IMACTI).Picture, (-1) * HScroll1.Value, (-1) * VScroll1.Value, 1 * HCUADRO, 1 * VCUADRO
End Sub

Private Sub Picture1_Click(Index As Integer)
      
      Screen.MousePointer = 11
      '
      Picture4.Top = 500
      Picture4.Left = 500
      Picture4.Width = 10000
      Picture4.Height = 7000
      '
      'HIMAGEN = Picture1(Index).Picture.Width
      'VIMAGEN = Picture1(Index).Picture.Height
      HIMAGEN = Image1(Index).Picture.Width
      VIMAGEN = Image1(Index).Picture.Height
      If HIMAGEN < 1 Or VIMAGEN < 1 Then
         GoTo 99
      End If
      '
      HCUADRO = Picture4.Width
      VCUADRO = Picture4.Height
      '
      REDUH = HIMAGEN / HCUADRO
      REDUV = VIMAGEN / VCUADRO
      REDUC = REDUH: If REDUV > REDUH Then REDUC = REDUV
      Picture4.Height = VCUADRO * REDUV / REDUC
      Picture4.Width = HCUADRO * REDUH / REDUC
      HCUADRO = Picture4.Width
      VCUADRO = Picture4.Height
      Picture4.Left = (12000 - HCUADRO) / 2
      '
      VScroll1.Top = Picture4.Top
      VScroll1.Left = Picture4.Left + Picture4.Width
      VScroll1.Height = Picture4.Height
      HScroll1.Top = Picture4.Top + Picture4.Height
      HScroll1.Width = Picture4.Width
      HScroll1.Left = Picture4.Left
      HScroll1.Max = HScroll1.Min   'HCUADRO / 2
      VScroll1.Max = VScroll1.Min   'VCUADRO / 2
      HScroll1.Value = 0
      VScroll1.Value = 0
      
      Picture4.PaintPicture Image1(Index).Picture, 0, 0, 1 * HCUADRO, 1 * VCUADRO
      Picture4.Visible = True
      HScroll1.Visible = True
      VScroll1.Visible = True
      IMACTI = Index
      '
99    Screen.MousePointer = 1
      '
End Sub

Private Sub Picture4_Click()
   Picture4.Visible = False
   HScroll1.Visible = False
   VScroll1.Visible = False
End Sub

Private Sub Command2_Click()
    Call HELPONLINE("FICHACAL")
End Sub
'

Private Sub Command9_Click()
    If MODIFJ% > 0 Then
        RESPU = MsgBox("Abandona sin Guardar los Cambios ?", 4371, "Operacion Critica")
        If RESPU <> 6 Then
            Command1.SetFocus
            Exit Sub
        End If
        MODIFJ% = 0
    End If
    '
    Hide
End Sub

Private Sub Text1_Change()
   MODIFJ% = 1
End Sub

Private Sub Text2_Change()
   MODIFJ% = 1
End Sub

Private Sub Text3_Change()
   MODIFJ% = 1
End Sub

Private Sub Text4_Change(Index As Integer)
   MODIFJ% = 1
End Sub

Private Sub Text4_KeyPress(Index As Integer, KeyAscii As Integer)
   '
   If KeyAscii% = 13 Then
      Call MUESTRAPLANO(Index)
   End If
   '
End Sub
'
Sub CARDESOPER()
   '
   COD$ = Label7.Caption
   CI% = CODINT%(Label7.Caption)
   Text3.TEXT = ""
   NOPE% = Val(Label8.Caption)
   '
   If CI% > 0 Then
      If CI% <= NUMAS% Then
         Screen.MousePointer = 11
         RNC$ = RECFILT$("DESOPFAB", 1, MKI$(CI%))
         URN& = ULTREG&("DESOPFAB")
         '
         For U& = 1 To Len(RNC$) Step 4
            RENOTA& = CVS(Mid$(RNC$, U&, 4))
            If RENOTA& > 0 Then
               If RENOTA& <= URN& Then
                  TTXX$ = REGLEIDO$("DESOPFAB", RENOTA&)
                  If CVI(Mid$(TTXX$, 3, 2)) = NOPE% Then
                     TTYY$ = Mid$(TTXX$, 5, 72)
                     If Text3.TEXT <> "" Then
                        Text3.TEXT = Text3.TEXT + Chr$(13) + Chr$(10)
                     End If
                     Text3.TEXT = Text3.TEXT + RTrim$(Left$(TTYY$, 72))
                  End If
               End If
            End If
         Next U&
         Screen.MousePointer = 1
      End If
   End If
   Text3.Refresh
   '
End Sub

Sub GRADESOPER()
   '
   COD$ = Label7.Caption
   CI% = CODINT%(COD$)
   If CI% > 0 Then
      If CI% <= NUMAS% Then
         NOPE% = Val(Label8.Caption)
         Screen.MousePointer = 11
         URN& = ULTREG&("DESOPFAB")
         RNC$ = RECFILT$("DESOPFAB", 1, MKI$(CI%))
         KK& = (-3)
         '
         TEXTO$ = Text3.TEXT + Chr$(13) + Chr$(10)
         Call FRATEXTO(TEXTO$, 72)
         For LL% = 1 To CARETEX%
            TTYY$ = MKI$(CI%) + MKI$(NOPE%) + RETEX$(LL%)
15          KK& = KK& + 4
            If KK& <= Len(RNC$) - 3 Then
                RENOTA& = CVS(Mid$(RNC$, KK&, 4))
              Else
                URN& = URN& + 1
                RENOTA& = URN&
            End If
            TTXX$ = REGLEIDO$("DESOPFAB", RENOTA&)
            NOPI% = CVI(Mid$(TTXX$, 3, 2))
            If NOPI% > 0 Then
               If NOPI% <> NOPE% Then
                  GoTo 15
               End If
            End If
            Call GRAREG("DESOPFAB", TTYY$, RENOTA&)
         Next LL%
         '
20       While KK& < Len(RNC$) - 3
            KK& = KK& + 4
            RENOTA& = CVS(Mid$(RNC$, KK&, 4))
            TTXX$ = REGLEIDO$("DESOPFAB", RENOTA&)
            If CVI(Mid$(TTXX$, 3, 2)) = NOPE% Then
               Call GRAREG("DESOPFAB", MKI$(0) + MKI$(0) + Space$(68), RENOTA&)
            End If
         Wend
         '
         Call CIERRARCH("DESOPFAB")
         Screen.MousePointer = 1
      End If
   End If

End Sub
'
Sub CARDATMECA()
   '
   COD$ = Label7.Caption
   CI% = CODINT%(Label7.Caption)
   Text1.TEXT = ""
   NOPE% = Val(Label8.Caption)
   '
   If CI% > 0 Then
      If CI% <= NUMAS% Then
         Screen.MousePointer = 11
         RNC$ = RECFILT$("DATMECAN", 1, MKI$(CI%))
         URN& = ULTREG&("DATMECAN")
         '
         For U& = 1 To Len(RNC$) Step 4
            RENOTA& = CVS(Mid$(RNC$, U&, 4))
            If RENOTA& > 0 Then
               If RENOTA& <= URN& Then
                  TTXX$ = REGLEIDO$("DATMECAN", RENOTA&)
                  If CVI(Mid$(TTXX$, 3, 2)) = NOPE% Then
                     TTYY$ = Mid$(TTXX$, 5, 50)
                     If Text1.TEXT <> "" Then
                        Text1.TEXT = Text1.TEXT + Chr$(13) + Chr$(10)
                     End If
                     Text1.TEXT = Text1.TEXT + RTrim$(Left$(TTYY$, 72))
                  End If
               End If
            End If
         Next U&
         Screen.MousePointer = 1
      End If
   End If
   Text1.Refresh
   '
End Sub

Sub GRADATMECA()
   '
   COD$ = Label7.Caption
   CI% = CODINT%(COD$)
   If CI% > 0 Then
      If CI% <= NUMAS% Then
         NOPE% = Val(Label8.Caption)
         Screen.MousePointer = 11
         URN& = ULTREG&("DATMECAN")
         RNC$ = RECFILT$("DATMECAN", 1, MKI$(CI%))
         KK& = (-3)
         '
         TEXTO$ = Text1.TEXT + Chr$(13) + Chr$(10)
         Call FRATEXTO(TEXTO$, 50)
         For LL% = 1 To CARETEX%
            TTYY$ = MKI$(CI%) + MKI$(NOPE%) + RETEX$(LL%)
15          KK& = KK& + 4
            If KK& <= Len(RNC$) - 3 Then
                RENOTA& = CVS(Mid$(RNC$, KK&, 4))
              Else
                URN& = URN& + 1
                RENOTA& = URN&
            End If
            TTXX$ = REGLEIDO$("DATMECAN", RENOTA&)
            NOPI% = CVI(Mid$(TTXX$, 3, 2))
            If NOPI% > 0 Then
               If NOPI% <> NOPE% Then
                  GoTo 15
               End If
            End If
            Call GRAREG("DATMECAN", TTYY$, RENOTA&)
         Next LL%
         '
20       While KK& < Len(RNC$) - 3
            KK& = KK& + 4
            RENOTA& = CVS(Mid$(RNC$, KK&, 4))
            TTXX$ = REGLEIDO$("DATMECAN", RENOTA&)
            If CVI(Mid$(TTXX$, 3, 2)) = NOPE% Then
               Call GRAREG("DATMECAN", MKI$(CI%) + MKI$(0) + Space$(124), RENOTA&)
            End If
         Wend
         '
         Call CIERRARCH("DATMECAN")
         Screen.MousePointer = 1
      End If
   End If

End Sub
'

Sub CARDATOPER()
   '
   COD$ = Label7.Caption
   CI% = CODINT%(Label7.Caption)
   NOPE% = Val(Label8.Caption)
   '
   If CI% > 0 Then
      If CI% <= NUMAS% Then
         Screen.MousePointer = 11
         RNC$ = RECFILT$("DATADOPS", 1, MKI$(CI%))
         URN& = ULTREG&("DATADOPS")
         '
         For U& = 1 To Len(RNC$) Step 4
            RENOTA& = CVS(Mid$(RNC$, U&, 4))
            If RENOTA& > 0 Then
               If RENOTA& <= URN& Then
                  TTXX$ = REGLEIDO$("DATADOPS", RENOTA&)
                  If CVI(Mid$(TTXX$, 3, 2)) = NOPE% Then
                     GoTo 29
                  End If
               End If
            End If
         Next U&
         TTXX$ = REGBLAN$("DATADOPS")
         '
29       NHPF$ = TRIM$(Mid$(TTXX$, 5, 16))
         Text8.TEXT = TRIM$(Mid$(TTXX$, 5, 16))
         Text2.TEXT = TRIM$(Mid$(TTXX$, 21, 4))
         FEX = CVI(Mid$(TTXX$, 25, 2)): If FEX < 1 Then FEX = HOY(HOS$)
         Text7.TEXT = FECHATEX$(FEX)
         Text5.TEXT = TRIM$(Mid$(TTXX$, 29, 8))
         USUA$ = TRIM$(Mid$(TTXX$, 37, 16)): If USUA$ = "" Then USUA$ = USERNAME$
         Text6.TEXT = USUA$
         Text4(0).TEXT = TRIM$(Mid$(TTXX$, 53, 48))
         Screen.MousePointer = 1
      End If
   End If
   '
End Sub

Sub GRADATOPER()
   '
   COD$ = Label7.Caption
   CI% = CODINT%(COD$)
   If CI% > 0 Then
      If CI% <= NUMAS% Then
         NOPE% = Val(Label8.Caption)
         Screen.MousePointer = 11
         URN& = ULTREG&("DATADOPS")
         RNC$ = RECFILT$("DATADOPS", 1, MKI$(CI%))
         KK& = (-3)
         '
         TTYY$ = REGBLAN$("DATADOPS")
         Call REPLA(TTYY$, MKI$(CI%), 1, 2)
         Call REPLA(TTYY$, MKI$(NOPE%), 3, 2)
         Call REPLA(TTYY$, Text8.TEXT, 5, 16)
         Call REPLA(TTYY$, Text2.TEXT, 21, 4)
         FEXI% = FECHANUM(Text7.TEXT): If FEXI% < 1 Then FEXI% = HOY(HOS$)
         Call REPLA(TTYY$, MKI$(FEXI%), 25, 2)
         Call REPLA(TTYY$, MKI$(FEXI%), 27, 2)
         Call REPLA(TTYY$, Text5.TEXT, 29, 8)
         USUA$ = TRIM$(Text6.TEXT): If USUA$ = "" Then USUA$ = USERNAME$
         Call REPLA(TTYY$, USUA$, 37, 16)
         Call REPLA(TTYY$, Text4(0).TEXT, 53, 48)
         '
15       KK& = KK& + 4
         If KK& <= Len(RNC$) - 3 Then
              RENOTA& = CVS(Mid$(RNC$, KK&, 4))
            Else
              URN& = URN& + 1
              RENOTA& = URN&
         End If
         TTXX$ = REGLEIDO$("DATADOPS", RENOTA&)
         NOPI% = CVI(Mid$(TTXX$, 3, 2))
         If NOPI% > 0 Then
            If NOPI% <> NOPE% Then
               GoTo 15
            End If
         End If
         Call GRAREG("DATADOPS", TTYY$, RENOTA&)
         '
20       While KK& < Len(RNC$) - 3
            KK& = KK& + 4
            RENOTA& = CVS(Mid$(RNC$, KK&, 4))
            TTXX$ = REGLEIDO$("DATADOPS", RENOTA&)
            If CVI(Mid$(TTXX$, 3, 2)) = NOPE% Then
               Call GRAREG("DATADOPS", MKI$(CI%) + MKI$(0) + Space$(124), RENOTA&)
            End If
         Wend
         '
         Call CIERRARCH("DATADOPS")
         Screen.MousePointer = 1
      End If
   End If
   '
End Sub
'
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
   Image1(Index).Picture = LoadPicture("")
   Picture1(Index).Cls
   COPLANO$ = RUTARGRA$ + TRIM$(Text4(Index).TEXT)
   If COPLANO$ <> "" Then
      If EXISTE%(COPLANO$) > 0 Then
         Picture1(Index).Visible = False
         Image1(Index).Picture = LoadPicture(COPLANO$)
         Picture1(Index).Top = TOPEINI(Index)
         Picture1(Index).Left = IZQUINI(Index)
         Picture1(Index).Width = ANCHINI(Index)
         Picture1(Index).Height = ALTUINI(Index)
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
   '
   Screen.MousePointer = 1
   '
99 End Sub

Private Sub Text5_Change()
   MODIFJ% = 1
End Sub

Private Sub Text6_Change()
   MODIFJ% = 1
End Sub

Private Sub Text7_Change()
   MODIFJ% = 1
End Sub

Private Sub Text8_Change()
   MODIFJ% = 1
End Sub

Private Sub VScroll1_Change()
     HCUADRO = Picture4.Width
     VCUADRO = Picture4.Height
     Picture4.PaintPicture Image1(IMACTI).Picture, (-1) * HScroll1.Value, (-1) * VScroll1.Value, 1 * HCUADRO, 1 * VCUADRO
End Sub
