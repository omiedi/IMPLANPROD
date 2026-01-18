VERSION 5.00
Begin VB.Form FORMUGANTOF 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFC0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Secuencia de Carga por Orden de Fabricación y Pieza"
   ClientHeight    =   8400
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   11910
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8400
   ScaleWidth      =   11910
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Imprimir"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   9870
      TabIndex        =   22
      Top             =   7410
      Width           =   1905
      Begin VB.OptionButton Option3 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Paisaje"
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
         TabIndex        =   25
         Top             =   550
         Width           =   960
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Retrato"
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
         TabIndex        =   24
         Top             =   300
         Value           =   -1  'True
         Width           =   960
      End
      Begin VB.CommandButton Command5 
         Height          =   540
         Left            =   1155
         Picture         =   "FORMUGANTOF.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   210
         Width           =   645
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Orden de Fabricación Activa"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2700
      Left            =   210
      TabIndex        =   7
      Top             =   5565
      Width           =   9465
      Begin VB.CommandButton BOTSEL 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   260
         Index           =   0
         Left            =   2310
         Picture         =   "FORMUGANTOF.frx":066A
         TabIndex        =   27
         TabStop         =   0   'False
         Top             =   380
         Width           =   160
      End
      Begin VB.ListBox List2 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1410
         Left            =   210
         TabIndex        =   20
         Top             =   1155
         Width           =   9045
      End
      Begin VB.Label Label15 
         BackColor       =   &H00FFFFFF&
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
         Height          =   255
         Left            =   8085
         TabIndex        =   19
         Top             =   735
         Width           =   1170
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Saldo:"
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
         Left            =   6510
         TabIndex        =   18
         Top             =   780
         Width           =   1485
      End
      Begin VB.Label Label13 
         BackColor       =   &H00FFFFFF&
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
         Height          =   255
         Left            =   8085
         TabIndex        =   17
         Top             =   380
         Width           =   1170
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Cantidad:"
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
         Left            =   6510
         TabIndex        =   16
         Top             =   420
         Width           =   1485
      End
      Begin VB.Label Label11 
         BackColor       =   &H00FFFFFF&
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
         Height          =   255
         Left            =   1155
         TabIndex        =   15
         Top             =   735
         Width           =   5580
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Producto:"
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
         TabIndex        =   14
         Top             =   780
         Width           =   960
      End
      Begin VB.Label Label8 
         BackColor       =   &H00FFFFFF&
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
         Height          =   255
         Left            =   5565
         TabIndex        =   13
         Top             =   380
         Width           =   1170
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha Solic:"
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
         Left            =   3990
         TabIndex        =   12
         Top             =   420
         Width           =   1485
      End
      Begin VB.Label Label4 
         BackColor       =   &H00FFFFFF&
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
         Height          =   255
         Left            =   3570
         TabIndex        =   11
         Top             =   380
         Width           =   330
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Status:"
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
         Left            =   2520
         TabIndex        =   10
         Top             =   420
         Width           =   960
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFFFFF&
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
         Height          =   260
         Left            =   1155
         TabIndex        =   9
         Top             =   380
         Width           =   1170
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Nro. O/F:"
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
         Width           =   960
      End
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   5220
      Left            =   210
      ScaleHeight     =   5160
      ScaleWidth      =   11505
      TabIndex        =   5
      Top             =   210
      Width           =   11565
      Begin VB.HScrollBar HScroll1 
         Height          =   225
         Left            =   2940
         TabIndex        =   26
         Top             =   4940
         Width           =   8360
      End
      Begin VB.VScrollBar VScroll1 
         Height          =   3615
         Left            =   11300
         TabIndex        =   21
         Top             =   1290
         Width           =   225
      End
      Begin VB.Line Line4 
         BorderWidth     =   2
         X1              =   0
         X2              =   11550
         Y1              =   525
         Y2              =   525
      End
      Begin VB.Line Line3 
         BorderWidth     =   2
         X1              =   -105
         X2              =   11550
         Y1              =   4920
         Y2              =   4920
      End
      Begin VB.Line Line1 
         BorderWidth     =   2
         X1              =   0
         X2              =   11550
         Y1              =   1260
         Y2              =   1260
      End
      Begin VB.Line Line2 
         BorderWidth     =   2
         X1              =   2940
         X2              =   2940
         Y1              =   0
         Y2              =   6930
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Rango"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1750
      Left            =   9870
      TabIndex        =   0
      Top             =   5565
      Width           =   1900
      Begin VB.OptionButton Option1 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Mensual"
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
         Left            =   315
         TabIndex        =   4
         Top             =   1365
         Width           =   1275
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Quincenal"
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
         TabIndex        =   3
         Top             =   1050
         Width           =   1275
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Semanal"
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
         Left            =   315
         TabIndex        =   2
         Top             =   735
         Value           =   -1  'True
         Width           =   1275
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Diario"
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
         Left            =   315
         TabIndex        =   1
         Top             =   420
         Width           =   1275
      End
   End
   Begin VB.Label TEPRUEBA 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "TEPRUEBA"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   10395
      TabIndex        =   6
      Top             =   5775
      Visible         =   0   'False
      Width           =   1020
   End
End
Attribute VB_Name = "FORMUGANTOF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim MATIEM$(1024), NORFBX$(1024), CIOFAX%(1024), PUNOFA$(1024)
Dim CAEQ%, CAOFA%, PRIOFVIS%, CAOFVIS%
Dim CODEQ$(1024), STATEQ%(1024), COLEQ$(1024)
Dim INCREY, X0MIN, X0MAX
Dim CODICON$, DENOCON$

Dim MAXSCROLL%

Dim INCREYP
Dim MARIZ%, MARSU%
Dim BIZQ%, BDER%, BSUP%, BINF%
Dim XINIP(512), YINIP(512)

Dim CAOTOT%, CAOTOT1%
Dim OTOT&(512), CIOTOT%(512), COMIE#(512), FINAL#(512), FEINOT(512), FEFIOT(512), YINIQ(512)


Private Sub BOTSEL_Click(Index As Integer)
   Call SELORFA
   If Val(Label2.Caption) < 1 Then Exit Sub
   Call CARLIORF
   Picture1.Cls
   MAXSCROLL% = 0
   Call DIBULINVER
   Call DIBUFECHA
   Call DIBUSUBOR
   Call PREGANTT
End Sub

Private Sub Command5_Click()
   PORIAN = Printer.Orientation
   If Option2.Value = True Then
        Printer.Orientation = 1
      Else
        Printer.Orientation = 2
   End If
   '
   Call PRILINVER
   Call PRIFECHA
   If TRIM$(Label2.Caption) <> "" Then
         Call PRISUBOR
         Call PRIGANTT
      Else
         Call PRIORTOT
         Call PRIGATOT
   End If
   Printer.EndDoc
   Printer.Orientation = PORIAN
End Sub

Private Sub Form_Load()
   EPAC$ = TRIM$(EMPREAC$)
   If EPAC$ <> "" Then
       Caption = Caption + "  -  " + EPAC$
   End If
   '
   FEIPER = HOY(HO$)
   While DIASEM%(FEIPER) <> 1
      FEIPER = DIANTE(FEIPER)
   Wend
   '
   HScroll1.Min = FEIPER - 32 * 13
   HScroll1.Max = FEIPER + 32 * 13
   HScroll1.Value = FEIPER
   '
   Picture1.Cls
   Call CALCUFECHA
   Call DIBULINVER
   Call DIBUFECHA
   'Call CARLIORF
   'Call DIBUSUBOR
   'Call PREGANTT
   '
   YACARGANTT% = 1
   '
End Sub

Private Sub HScroll1_Change()
   If YACARGANTT% = 1 Then
      FEIPERANT = FEIPER
      FEIPER = HScroll1.Value
      SENTI% = 1
      If FEIPER < FEIPERANT Then SENTI% = -1
10    FEX = FEIPER
      If Val(Left$(FECHATEX$(FEX), 2)) < 1 Then
         FEIPER = FEIPER + SENTI%
         GoTo 10
      End If
      '
      If Val(Mid$(FECHATEX$(FEX), 4, 2)) < 1 Then
         FEIPER = FEIPER + SENTI%
         GoTo 10
      End If
      FF1 = DIPOST(FEX)
      FF2 = DIANTE(FF1)
      If FF2 <> FEIPER Then
         FEIPER = FEIPER + SENTI%
         GoTo 10
      End If
      '
   End If
   '
   HScroll1.Value = FEIPER
   Picture1.Cls
   Call CALCUFECHA
   Call DIBULINVER
   Call DIBUFECHA
   Call DIBUSUBOR
   Call PREGANTT
End Sub

Private Sub List1_Click()
   If CONTROCLICK% = 0 Then
      Call PREGANTT
      NORFA$ = TRIM$(Left$(List1.TEXT, 24))
      Call MUEDATORF(NORFA$)
   End If
End Sub

Private Sub List3_Click()
   CIIX% = CODINT%(List3.TEXT)
   CONTROCLICK% = (-1)
   OFACT& = 0
   For U& = 1 To List1.ListCount
      List1.Selected(U& - 1) = False
      V& = Val(Mid$(List1.List(U& - 1), 25))
      If CIOFA%(V&) = CIIX% Then
         List1.Selected(U& - 1) = True
         OFACT& = U&
      End If
   Next U&
   CONTROCLICK% = 0
   List1.Refresh
   Call PREGANTT
   If OFACT& > 0 Then
      List1.ListIndex = OFACT& - 1
      NORFA$ = TRIM$(Left$(List1.TEXT, 24))
      Call MUEDATORF(NORFA$)
   End If
End Sub

Private Sub Option1_Click(Index As Integer)
   Picture1.Cls
   Call CALCUFECHA
   Call DIBULINVER
   Call DIBUFECHA
   Call DIBUSUBOR
   Call PREGANTT
End Sub

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   For UI% = 1 To cabarr%
      If X >= XINI1(UI%) Then
        If X <= XFIN1(UI%) Then
          If Y >= YINI1(UI%) Then
            If Y <= YINI1(UI%) + 145 Then
'              NORFAX$ = NUOFA$(UI%)
'              List2.Selected(UI% - 1) = True
'              For U& = 1 To List1.ListCount
'                 V& = Val(Mid$(List1.List(U& - 1), 25))
'                 If NORFB$(V&) = NORFAX$ Then
'                   CONTROCLICK% = -1
'                   List1.ListIndex = U& - 1
'                   CONTROCLICK% = 0
'                   OFACT& = U&
'                   Call PREGANTT
'                   Call MUEDATORF(NORFAX$)
'                   Exit Sub
'                 End If
 '             Next U&
            End If
          End If
        End If
      End If
   Next UI%
End Sub

Private Sub VScroll1_Change()
   PRIOFVIS% = VScroll1.Value + 1
   Picture1.Cls
   Call CALCUFECHA
   Call DIBULINVER
   Call DIBUFECHA
   Call DIBUSUBOR
   Call PREGANTT
End Sub
'

Sub DIBULINVER()
   CAPERIO% = 7: FRAHORA% = 6
   If Option1(0) = True Then
        CAPERIO% = 2: FRAHORA% = 2
      ElseIf Option1(2) = True Then
        CAPERIO% = 15: FRAHORA% = 6
      ElseIf Option1(3) = True Then
        CAPERIO% = 31: FRAHORA% = 12
   End If
   APERIO = (Picture1.Width - Line2.X1 - VScroll1.Width - 80) / (CAPERIO%)
   DrawWidth = 2
   X1 = Line2.X1
   Y2 = Picture1.Height
   YH1 = Line4.Y1
   For II% = 1 To CAPERIO% + 1
      DSXX% = DIASEM%(DESPERIO(II%))
      If II% <= CAPERIO% Then
         If DSXX% > 5 Or DSXX% < 1 Then
            Picture1.Line (X1 + 20, YH1)-(X1 + APERIO, Y2), RGB(220, 220, 220), BF
         End If
         XINI(II%) = X1
         For IJ% = FRAHORA% To 24 Step FRAHORA%
            XH1 = X1 + IJ% * APERIO / 24
            Picture1.DrawWidth = 1
            Picture1.Line (XH1, YH1)-(XH1, Y2), QBColor(8)
         Next IJ%
      End If
      Picture1.DrawWidth = 2
      X1 = X1 + APERIO
      If II% = CAPERIO% Then XMAXI = X1 - 20
      Picture1.Line (X1, 0)-(X1, Y2), QBColor(0)
   Next II%
   Picture1.Refresh
End Sub
'

Sub CALCUFECHA()
   FEBA = FEIPER
   For II% = 1 To 32   ' CAPERIO% + 1
      DESPERIO(II%) = FEBA
      FEBA = DIPOST(FEBA)
   Next II%
End Sub
'

Sub DIBUFECHA()
   Picture1.Font = TEPRUEBA.Font
   TEPRUEBA.FontBold = True
   If Option1(2).Value = True Then
      TEPRUEBA.FontBold = False
   End If
   Picture1.FontSize = TEPRUEBA.FontSize
   Picture1.FontBold = TEPRUEBA.FontBold
   FORMAFE% = 0: If CAPERIO% > 15 Then FORMAFE% = 1
   For II% = 1 To CAPERIO%
      TTXX$ = FECHATEX$(DESPERIO(II%))
      If CAPERIO% > 7 Then
         TTXX$ = Left$(TTXX$, 5)
      End If
      If CAPERIO% > 15 Then
         TTYY$ = Mid$(TTXX$, 4, 2)
         TTXX$ = Left$(TTXX$, 2)
         If Left$(TTYY$, 1) = "0" Then TTYY$ = Mid$(TTYY$, 2)
      End If
      TEPRUEBA.Caption = TTXX$
      Picture1.CurrentX = Line2.X1 + APERIO * II% - APERIO / 2 - TEPRUEBA.Width / 2
      Picture1.CurrentY = 120 - 60 * FORMAFE%
      Picture1.Print TTXX$
      If FORMAFE% = 1 Then
         TEPRUEBA.Caption = "---"
         Picture1.CurrentX = Line2.X1 + APERIO * II% - APERIO / 2 - TEPRUEBA.Width / 2
         Picture1.CurrentY = 165
         Picture1.Print "---"
         TEPRUEBA.Caption = TTYY$
         Picture1.CurrentX = Line2.X1 + APERIO * II% - APERIO / 2 - TEPRUEBA.Width / 2
         Picture1.CurrentY = 300
         Picture1.Print TTYY$
        Else
         TTYY$ = DSEM$(DESPERIO(II%))
         If CAPERIO% > 7 Then TTYY$ = Left$(TTYY$, 2)
         TTYY$ = Left$(TTYY$, 1) + LCase$(Mid$(TTYY$, 2))
         TTYY$ = "(" + TTYY$ + ")"
         TEPRUEBA.Caption = TTYY$
         Picture1.CurrentX = Line2.X1 + APERIO * II% - APERIO / 2 - TEPRUEBA.Width / 2
         Picture1.CurrentY = 300
         Picture1.Print TTYY$
      End If
   Next II%
   TEPRUEBA.FontBold = True
   Picture1.CurrentX = 80
   Picture1.CurrentY = 240
   Picture1.Print "Pieza / Referencia"
End Sub
'

Sub SELORFA()
   '
   Static INDI%
   If INDI% = 0 Then
      Screen.MousePointer = 11
      Call COPYSTRU("ORFANUE", "ORFANUL")
      J& = 0
      '
      For I& = 1 To ULTREG&("ORFANUE")
         XX$ = REGLEIDO$("ORFANUE", I&)
         ST% = Asc(Mid$(XX$, 112, 1))
         '
         If ST% <> 3 Then
            If CVI(Mid$(XX$, 119, 2)) > 0 Then
                Call REPLA(XX$, MKI$(0), 119, 2)
                Call GRAREG("ORFANUE", XX$, I&)
            End If
         End If
         '
         If ST% = 1 Then      ' solo lanzadas
            J& = J& + 1
            Call GRAREG("ORFANUL", XX$, J&)
         End If
       Next I&
       '
       Call SETULTREG("ORFANUL", J&)
       Call CIERRARCH("ORFANUL")
       Call FRESHECHO("ORFANUL")
       INDI% = 1
       '
       Screen.MousePointer = 1
   End If
   '
   'Call SELECHO("ORFANUL")
   XXX$ = REGSEL$("ORFANUL")
   NOFA$ = VALACT1$("ORFANUL")
   If TRIM$(NOFA$) = "" Then
      CODICON$ = ""
      DENOCON$ = ""
      Label2.Caption = ""
      Label11.Caption = ""
      Label4.Caption = ""
      Label8.Caption = ""
      Label13.Caption = ""
      Label15.Caption = ""
      List2.Clear
      Picture1.Cls
      Call CALCUFECHA
      Call DIBULINVER
      Call DIBUFECHA
      Exit Sub
   End If
   CODICON$ = CODEXT$(CVI(Mid$(XXX$, 71, 2)))
   DENOCON$ = DESCRIT0$(CVI(Mid$(XXX$, 71, 2)))
   '
   Label2.Caption = TRIM$(NOFA$)
   Label11.Caption = VALACT2$("ORFANUL")
   Label4.Caption = Str$(Asc(Mid$(XXX$, 112, 1)))
   FEX = CVI(Mid$(XXX$, 109, 2))
   Label8.Caption = FECHATEX$(FEX)
   Label13.Caption = Str$(CVD(Mid$(XXX$, 89, 8)))
   Label15.Caption = Str$(CVD(Mid$(XXX$, 89, 8)))
   '
End Sub
'
Sub CARLIORF()
   '
   NOFSEL& = Val(Label2.Caption)
   If Val(Label2.Caption) < 1 Then Exit Sub
   '
   Screen.MousePointer = 11
   List2.Clear
   '
   ROJO% = 128: VERDE% = 128: AZUL% = 128
   CAEQ% = 0
   If CAEQ% < 1 Then
      For U& = 1 To ULTREG&("PADMAQ")
         ZZ$ = REGLEIDO$("PADMAQ", U&)
         If TRIM$(Left$(ZZ$, 6)) <> "" Then
            CAEQ% = CAEQ% + 1
            CODEQ$(CAEQ%) = Left$(ZZ$, 6)
            STATEQ%(CAEQ%) = 0
            ROJO% = ROJO% + 150
            If ROJO% > 255 Then VERDE% = VERDE% + 150
            If VERDE% > 255 Then AZUL% = AZUL% + 150
            ROJO% = ROJO% Mod 256
            VERDE% = VERDE% Mod 256
            AZUL% = AZUL% Mod 256
            COLEQ$(CAEQ%) = Chr$(ROJO%) + Chr$(VERDE%) + Chr$(AZUL%)
            If Mid$(ZZ$, 97, 3) <> "   " Then
               If Mid$(ZZ$, 97, 3) <> String$(3, 0) Then
                  COLEQ$(CAEQ%) = Mid$(ZZ$, 97, 3)
               End If
            End If
         End If
      Next U&
   End If
   '
   CAOSEL% = 0
   For U& = 1 To ULTREG&("OPERFAP")
      '
      XX$ = REGLEIDO$("OPERFAP", U&)
      NOFA$ = TRIM$(Left$(XX$, 12))
      If Val(Mid$(NOFA$, 4, 6)) = NOFSEL& Then
         TTXX$ = "A" + Space$(23)
         Call REPLA(TTXX$, CSTRING$(Mid$(XX$, 135, 2), 1, 5, 0, 2), 2, 5)
         Call REPLA(TTXX$, CSTRING$(Mid$(XX$, 137, 2), 1, 4, 0, 2), 7, 4)
         Call REPLA(TTXX$, CSTRING$(MKS$(U&), 3, 6, 0, 2), 11, 6)
         CAOSEL% = CAOSEL% + 1
         PUNOFA$(CAOSEL%) = TTXX$
      End If
   Next U&
   '
   For K% = 1 To CAOSEL%
      For K1% = K% + 1 To CAOSEL%
         If PUNOFA$(K%) > PUNOFA$(K1%) Then
            TTXX$ = PUNOFA$(K%)
            PUNOFA$(K%) = PUNOFA$(K1%)
            PUNOFA$(K1%) = TTXX$
         End If
      Next K1%
   Next K%
   '
   CAOFA% = 0
   For KQ% = 1 To CAOSEL%
      U& = Val(Mid$(PUNOFA$(KQ%), 11, 6))
         '
         XX$ = REGLEIDO$("OPERFAP", U&)
         NOFA$ = TRIM$(Left$(XX$, 12))
         MAQUI$ = Mid$(XX$, 73, 6) + Mid$(XX$, 203, 18)
         ITIEM$ = Mid$(XX$, 135, 8)
         CIIX% = CVI(Mid$(XX$, 13, 2))
         DIVECTI& = CVS(Mid$(XX$, 253, 4))
         '
         For K% = 1 To CAOFA%
            If NORFBX$(K%) = NOFA$ Then GoTo 10
         Next K%
         CAOFA% = CAOFA% + 1
         K% = CAOFA%
         NORFBX$(K%) = NOFA$
         CIOFAX%(K%) = CIIX%
         MATIEM$(K%) = ""
         '
10       MATIEM$(K%) = MATIEM$(K%) + MAQUI$ + ITIEM$ + MKS$(DIVECTI&)
         '
         TTXX$ = Space$(128)
         DEOP$ = TRIM$(Mid$(XX$, 18, 47)) + " " + DESCRIT0$(CVI(Mid$(XX$, 13, 2)))
         Call REPLA(TTXX$, DEOP$, 1, 30)
         Call REPLA(TTXX$, Mid$(XX$, 73, 6), 33, 6)
         SALDO# = CVD(Mid$(XX$, 79, 8)) - CVD(Mid$(XX$, 87, 8))
         If SALDO# < 0 Then SALDO# = 0
         Call REPLA(TTXX$, CSTRING$(MKD$(SALDO#), 4, 10, 1, 2), 39, 10)
         Call REPLA(TTXX$, Left$(HORMINSE$(60 * CVS(Mid$(XX$, 131, 4))), 6), 50, 6)
         FEX1 = CVI(Mid$(XX$, 135, 2))
         HORA1% = CVI(Mid$(XX$, 137, 2))
         FEX2 = CVI(Mid$(XX$, 139, 2))
         HORA2% = CVI(Mid$(XX$, 141, 2))
         If FEX1 > FEX2 Then
            FEX1 = 0: FEX2 = 0
            HORA1% = 0: HORA2% = 0
            Call REPLA(XX$, String$(8, 0), 135, 8)
            Call GRAREG("OPERFAP", XX$, U&)
         End If
         Call REPLA(TTXX$, Left$(FECHATEX$(FEX1), 5), 58, 5)
         Call REPLA(TTXX$, Left$(HORATEX$(HORA1%), 5), 64, 5)
         Call REPLA(TTXX$, Left$(FECHATEX$(FEX2), 5), 72, 5)
         Call REPLA(TTXX$, Left$(HORATEX$(HORA2%), 5), 78, 5)
         List2.AddItem TTXX$
         '
      'End If
   Next KQ%
   List2.Refresh
   Screen.MousePointer = 1
   '
End Sub
'

Sub DIBUSUBOR()
   '
   ' Static MAXSCROLL%
   '
   If Val(Label2.Caption) < 1 Then Exit Sub
   '
   Picture1.Font = TEPRUEBA.Font
   Picture1.FontSize = TEPRUEBA.FontSize
   Picture1.FontBold = TEPRUEBA.FontBold
   INCREY = TEPRUEBA.Height * 1.05
   Picture1.CurrentY = Line1.Y1 - INCREY '/ 2
   '
   If PRIOFVIS% < 1 Then PRIOFVIS% = 1
   AMAXI = Line2.X1 - 100
   For U& = PRIOFVIS% To CAOFA%
      Picture1.CurrentY = Picture1.CurrentY + 1.4 * INCREY
      YINI(U&) = Picture1.CurrentY
      If YINI(U&) <= Picture1.Height - HScroll1.Height - 2.5 * INCREY Then
         TTXX$ = TRIM$(CODEXT$(CIOFAX%(U&))) + " (" + TRIM$(NORFBX$(U&)) + ")"
10       TEPRUEBA.Caption = TTXX$
         If TEPRUEBA.Width > AMAXI Then TTXX$ = Left$(TTXX$, Len(TTXX$) - 1): GoTo 10
         Picture1.CurrentX = 80
         Picture1.Print TTXX$;
         TTXX$ = TRIM$(DESCRIT0$(CIOFAX%(U&)))
20       TEPRUEBA.Caption = TTXX$
         If TEPRUEBA.Width > AMAXI Then TTXX$ = Left$(TTXX$, Len(TTXX$) - 1): GoTo 20
         Picture1.CurrentX = 80
         Picture1.CurrentY = Picture1.CurrentY + INCREY
         Picture1.Print TTXX$;
      End If
   Next U&
   '
   If TRIM$(CODICON$) <> "" Then
      TTXX$ = "Código: " + CODICON$
      Picture1.CurrentX = 80
      Picture1.CurrentY = (Line1.Y1 + Line4.Y1) / 2 - INCREY
      Picture1.Print TTXX$;
   End If
   '
   TTXX$ = TRIM$(DENOCON$)
25 TEPRUEBA.Caption = TTXX$
   If TEPRUEBA.Width > AMAXI Then TTXX$ = Left$(TTXX$, Len(TTXX$) - 1): GoTo 25
   Picture1.CurrentX = 80
   Picture1.CurrentY = (Line1.Y1 + Line4.Y1) / 2
   Picture1.Print TTXX$;
   '
30 If CAOFVIS% = 0 Then
      CAOFVIS% = (Picture1.Height - Line1.Y1 - HScroll1.Height - 2 * INCREY) / (2.4 * INCREY)
   End If
   '
   If MAXSCROLL% <= 0 Then
      VScroll1.Max = 0
      If CAOFVIS% < CAOFA% Then
         VScroll1.Max = CAOFA% - CAOFVIS%
      End If
      MAXSCROLL% = VScroll1.Max
   End If
   VScroll1.Max = MAXSCROLL%
   VScroll1.Value = PRIOFVIS% - 1
   '
End Sub
'

Sub PREGANTT()
   '
   CONTROCLICK = (-1)
   'TOPE = FORMUGANTT.List1.TopIndex
   'ACTI = FORMUGANTT.List1.ListIndex
   cabarr% = 0
   If Val(Label2.Caption) < 1 Then Exit Sub
   '
   X0MIN = 32767: X0MAX = 0
   For U& = 1 To CAOFA%
      NOFA$ = NORFBX$(U&)
      'NOFA$ = NORFB$(U&)
      COLO$ = Chr$(0) + Chr$(230) + Chr$(64)
      '
      'If FORMUGANTT.List1.Selected(U& - 1) = True Then
      '   COLO$ = Chr$(255) + Chr$(0) + Chr$(0)
      'End If
      '
      'If FORMUGANTT.List1.ListIndex = U& - 1 Then
      '   COLO$ = Chr$(128) + Chr$(0) + Chr$(128)
      'End If
      '
      V& = U&
      For J% = 1 To Len(MATIEM$(V&)) Step 36
         MAQUI$ = Mid$(MATIEM$(V&), J%, 24)
         FEHOR$ = Mid$(MATIEM$(V&), J% + 24, 8)
         DIVECTI& = CVS(Mid$(MATIEM$(V&), J% + 32, 4))
         FEINI% = CVI(Mid$(FEHOR$, 1, 2))
         HOINI% = CVI(Mid$(FEHOR$, 3, 2))
         FEFIN% = CVI(Mid$(FEHOR$, 5, 2))
         HOFIN% = CVI(Mid$(FEHOR$, 7, 2))
         If FEINI% > FEFIN% Then
            FEINI% = 0
            HOINI% = 0
            FEFIN% = 0
            HOFIN% = 0
         End If
         For K% = 1 To 24 Step 6
            EQUI$ = Mid$(MAQUI$, K%, 6)
            For KQ% = 1 To CAEQ%
               If CODEQ$(KQ%) = EQUI$ Then COLO$ = COLEQ$(KQ%)
            Next KQ%
            'If TRIM$(EQUI$) <> "" Then
               CBX% = cabarr%
               Call DIBUBARR(FEINI%, HOINI%, FEFIN%, HOFIN%, NOFA$, EQUI$, COLO$, DIVECTI&)
               If cabarr% > CBX% Then
                  NUOFA$(cabarr%) = NOFA$
               End If
            'End If
         Next K%
      Next J%
      '
   Next U&
   '
   HScroll1.LargeChange = CAPERIO%
   CONTROCLICK% = 0
   YACARGANTT% = 1
   '
End Sub
'

Sub DIBUBARR(FEINI%, HOINI%, FEFIN%, HOFIN%, NOFA$, EQUI$, COLO$, DIVECTI&)
   '
   Static DHR0%, TUNOR
   If Val(Label2.Caption) < 1 Then Exit Sub
   If TUNOR < 1 Then
      'DHR0% = 2400: TUNOR = 0
      'For U& = 1 To ULTREG&("TURGEN")
      '   X$ = REGLEIDO$("TURGEN", U&)
      '   DSX% = Asc(Left$(X$, 1))
      '   If DSX% > 0 Then
      '      DHR% = CVI(Mid$(X$, 2, 2))
      '      HHR% = CVI(Mid$(X$, 4, 2))
      '      If HHR% > DHR% Then
      '         If DHR% < DHR0% Then DHR0% = DHR%
      '         TT1 = 60 * Int(DHR% / 100) + (DHR% Mod 100)
      '         TT2 = 60 * Int(HHR% / 100) + (HHR% Mod 100)
      '         If (TT2 - TT1) > TUNOR Then TUNOR = TT2 - TT1
      '      End If
      '   End If
      'Next U&
      'DHR0% = 60 * Int(DHR0% / 100) + (DHR0% Mod 100)
   End If
   TUNOR = 24 * 60: DHR0% = 0
   '
   If FEINI% > DESPERIO(CAPERIO% + 1) Then
      Exit Sub
   End If
   '
   If FEFIN% < DESPERIO(1) Then
      Exit Sub
   End If
   '
   If FEINI% < DESPERIO(1) Then
      X1 = XINI(1)
      GoTo 10
   End If
   '
   For II% = 1 To CAPERIO%
      If FEINI% = DESPERIO(II%) Then
         MINU1% = 60 * Int(HOINI% / 100) + (HOINI% Mod 100)
         X1 = XINI(II%) + APERIO * (MINU1% - DHR0%) / TUNOR
         GoTo 10
      End If
   Next II%
   Exit Sub
   '
10 If FEFIN% >= DESPERIO(CAPERIO% + 1) Then
      X2 = XINI(CAPERIO%) + APERIO
      GoTo 20
   End If
   '
   For II% = 1 To CAPERIO%
      If FEFIN% = DESPERIO(II%) Then
         MINU2% = 60 * Int(HOFIN% / 100) + (HOFIN% Mod 100)
         X2 = XINI(II%) + APERIO * (MINU2% - DHR0%) / TUNOR
         GoTo 20
      End If
   Next II%
   '
20 If X1 < X0MIN Then X0MIN = X1
   If X2 > X0MAX Then X0MAX = X2
   Y0 = (Line1.Y1 + Line4.Y1) / 2
   Picture1.Line (X0MIN, Y0)-(X0MAX, Y0 + 145), RGB(128, 128, 128), BF
   Picture1.Line (X0MIN + 12, Y0 + 145)-(X0MAX, Y0 + 157), QBColor(0), BF
   Picture1.Line (X0MAX - 12, Y0 + 25)-(X0MAX, Y0 + 157), QBColor(0), BF
   '
   For II% = PRIOFVIS% To CAOFA%
      If NORFBX$(II%) = NOFA$ Then
         Y1 = YINI(II%) + INCREY
         GoTo 30
      End If
   Next II%
   'Exit Sub
   '
30 If Y1 < 100 Then Exit Sub
   'If X2 > XMAXI Then X2 = XMAXI
   ROJO% = Asc(Left$(COLO$, 1))
   VERDE% = Asc(Mid$(COLO$, 2, 1))
   AZUL% = Asc(Mid$(COLO$, 3, 1))
   Picture1.Line (X1, Y1 + 90)-(X2, Y1 + 145), RGB(ROJO%, VERDE%, AZUL%), BF
   Picture1.Line (X1 + 12, Y1 + 145)-(X2, Y1 + 157), QBColor(0), BF
   Picture1.Line (X2 - 12, Y1 + 25)-(X2, Y1 + 157), QBColor(0), BF
   X2MA = X2
   '
   cabarr% = cabarr% + 1
   XINI1(cabarr%) = X1
   XFIN1(cabarr%) = X2
   YINI1(cabarr%) = Y1
   '
   If DIVECTI& > 0 Then
      If DIVECTI& <= ULTREG&("VECTIPAR") Then
         VECTIPAR$ = REGLEIDO$("VECTIPAR", DIVECTI&)
         LOVECTI% = CVI(Left$(VECTIPAR$, 2))
         If LOVECTI% > 254 Then LOVECTI% = 254
         VECTIPAR$ = Mid$(VECTIPAR$, 3) + String$(6, 0)
         For KKI% = 1 To LOVECTI% Step 6
            FFII% = CVI(Mid$(VECTIPAR$, KKI%, 2))
            HHII% = CVI(Mid$(VECTIPAR$, KKI% + 2, 2))
            HHFF% = CVI(Mid$(VECTIPAR$, KKI% + 4, 2))
            If HHFF% < HHII% Then HHFF% = HHII%
            '
            If FFII% < DESPERIO(1) Or FFII% > DESPERIO(CAPERIO%) Then
               GoTo 190
            End If
            '
            For II% = 1 To CAPERIO%
               If FFII% = DESPERIO(II%) Then
                  MINU1% = 60 * Int(HHII% / 100) + (HHII% Mod 100)
                  X1 = XINI(II%) + APERIO * (MINU1% - DHR0%) / TUNOR
                  GoTo 110
               End If
            Next II%
            Exit Sub
            '
110         If FFII% >= DESPERIO(CAPERIO% + 1) Then
               X2 = XINI(CAPERIO%) + APERIO
               GoTo 130
            End If
            '
            For II% = 1 To CAPERIO%
               If FFII% = DESPERIO(II%) Then
                  MINU2% = 60 * Int(HHFF% / 100) + (HHFF% Mod 100)
                  X2 = XINI(II%) + APERIO * (MINU2% - DHR0%) / TUNOR
                  GoTo 130
               End If
            Next II%
            '
130         If Y1 < 100 Then Exit Sub
            If X2 > X2MA - 30 Then X2 = X2MA - 30
            ROJO% = Asc(Left$(COLO$, 1))
            VERDE% = Asc(Mid$(COLO$, 2, 1))
            AZUL% = Asc(Mid$(COLO$, 3, 1))
            Picture1.Line (X1, Y1)-(X2, Y1 + 100), RGB(ROJO%, VERDE%, AZUL%), BF
            '
190      Next KKI%
      End If
   End If
   '
End Sub
'
Sub PRILINVER()
   CAPERIO% = 7: FRAHORA% = 6
   If Option1(0) = True Then
        CAPERIO% = 2: FRAHORA% = 2
      ElseIf Option1(2) = True Then
        CAPERIO% = 15: FRAHORA% = 6
      ElseIf Option1(3) = True Then
        CAPERIO% = 31: FRAHORA% = 12
   End If
   BIZQ% = 1000: BDER% = 0.96 * Printer.ScaleWidth - 20
   BSUP% = 10: BINF% = 0.96 * Printer.ScaleHeight - 20
   If Printer.Orientation = 1 Then BSUP% = 500 Else BIZQ% = 500
   '
   DrawWidth = 20
   Printer.Line (BIZQ%, BSUP%)-(BDER%, BSUP% + 10), QBColor(0), BF
   Printer.Line (BDER%, BSUP%)-(BDER% + 10, BINF%), QBColor(0), BF
   Printer.Line (BIZQ%, BSUP%)-(BIZQ% + 10, BINF%), QBColor(0), BF
   Printer.Line (BIZQ%, BINF%)-(BDER%, BINF% + 10), QBColor(0), BF
   '
   MARSU% = Line4.Y1
   Printer.Line (BIZQ%, BSUP% + MARSU%)-(BDER%, BSUP% + MARSU% + 10), QBColor(0), BF
   '
   INCREYP = (BINF% - BSUP% - MARSU%) / (CAOFA% + 1) / 2.5
   If TEPRUEBA.Height * 1.05 < INCREYP Then INCREYP = TEPRUEBA.Height * 1.05
   If TRIM$(Label2.Caption) <> "" Then
      Printer.Line (BIZQ%, BSUP% + MARSU% + 3 * INCREYP)-(BDER%, BSUP% + MARSU% + 3 * INCREYP + 10), QBColor(0), BF
   End If
   MARIZ% = Line2.X1
   Printer.Line (BIZQ% + MARIZ%, BSUP%)-(BIZQ% + MARIZ% + 10, BINF%), QBColor(0), BF
   '
   APERIO = (BDER% - BIZQ% - MARIZ%) / (CAPERIO%)
   '
   DrawWidth = 2
   X1 = BIZQ% + MARIZ% 'Line2.X1
   Y2 = Printer.ScaleHeight   ' Picture1.Height
   For II% = 1 To CAPERIO% + 1
      XINIP(II%) = X1
      X1 = X1 + APERIO
      If II% = CAPERIO% Then XMAXI = X1 - 20
      Printer.Line (X1 - 3, BSUP%)-(X1 + 3, BINF%), QBColor(0), BF
      '
      If II% <= CAPERIO% Then    ' PARA GRISAR FERIADOS
         For IJ% = FRAHORA% To 23 Step FRAHORA%
            XH1 = XINIP(II%) + IJ% * APERIO / 24
            YH1 = BSUP% + MARSU%
            Printer.DrawWidth = 1
            Printer.Line (XH1, YH1)-(XH1, BINF%), QBColor(8)
         Next IJ%
         '
         FEX = DESPERIO(II%)
         DIA% = FEX: FERIA% = 0
         For UU& = 1 To ULTREG&("CALFERI")
            If CVI(Left$(REGLEIDO$("CALFERI", UU&), 2)) = DIA% Then
               FERIA% = 1
            End If
         Next UU&
         '
         If FERIA% = 1 Or DIASEM%(FEX) > 5 Or DIASEM%(FEX) < 1 Then
            XXA = XINIP(II%): YYA = BSUP%
            While YYA < BINF% + APERIO
               YYA = YYA + 300
               XA1 = XXA: YA1 = YYA
               XA2 = X1: YA2 = YYA - APERIO
               If YA2 < BSUP% Then
                  YA2 = BSUP%
                  XA2 = XA1 + (YA1 - YA2)
               End If
               If YA1 > BINF% Then
                  XA1 = XA1 + (YA1 - BINF%)
                  YA1 = BINF%
               End If
               If XA1 >= XINIP(II%) Then
                If XA1 <= XINIP(II%) + APERIO Then
                 If XA2 >= XINIP(II%) Then
                  If XA2 <= XINIP(II%) + APERIO Then
                   If YA1 >= BSUP% Then
                    If YA1 <= BINF% Then
                     If YA2 >= BSUP% Then
                      If YA2 <= BINF% Then
                       Printer.Line (XA1, YA1)-(XA2, YA2), QBColor(0)
                      End If
                     End If
                    End If
                   End If
                  End If
                 End If
                End If
               End If
            Wend
         End If
      End If
   Next II%
   'Picture1.Refresh
End Sub
'
Sub PRIFECHA()
   Printer.Font = TEPRUEBA.Font
   TEPRUEBA.FontBold = True
   If Option1(2).Value = True Then
      TEPRUEBA.FontBold = False
   End If
   Printer.FontSize = TEPRUEBA.FontSize
   Printer.FontBold = TEPRUEBA.FontBold
   FORMAFE% = 0: If CAPERIO% > 15 Then FORMAFE% = 1
   For II% = 1 To CAPERIO%
      TTXX$ = FECHATEX$(DESPERIO(II%))
      If CAPERIO% > 7 Then
         TTXX$ = Left$(TTXX$, 5)
      End If
      If CAPERIO% > 15 Then
         TTYY$ = Mid$(TTXX$, 4, 2)
         TTXX$ = Left$(TTXX$, 2)
         If Left$(TTYY$, 1) = "0" Then TTYY$ = Mid$(TTYY$, 2)
      End If
      TEPRUEBA.Caption = TTXX$
      Printer.CurrentX = BIZQ% + MARIZ% + APERIO * II% - APERIO / 2 - TEPRUEBA.Width / 2
      Printer.CurrentY = 120 - 60 * FORMAFE% + BSUP%
      Printer.Print TTXX$
      If FORMAFE% = 1 Then
         TEPRUEBA.Caption = "---"
         Printer.CurrentX = BIZQ% + MARIZ% + APERIO * II% - APERIO / 2 - TEPRUEBA.Width / 2
         Printer.CurrentY = 165 + BSUP%
         Printer.Print "---"
         TEPRUEBA.Caption = TTYY$
         Printer.CurrentX = BIZQ% + MARIZ% + APERIO * II% - APERIO / 2 - TEPRUEBA.Width / 2
         Printer.CurrentY = 300 + BSUP%
         Printer.Print TTYY$
        Else
         TTYY$ = DSEM$(DESPERIO(II%))
         If CAPERIO% > 7 Then TTYY$ = Left$(TTYY$, 2)
         TTYY$ = Left$(TTYY$, 1) + LCase$(Mid$(TTYY$, 2))
         TTYY$ = "(" + TTYY$ + ")"
         TEPRUEBA.Caption = TTYY$
         Printer.CurrentX = BIZQ% + MARIZ% + APERIO * II% - APERIO / 2 - TEPRUEBA.Width / 2
         Printer.CurrentY = 300 + BSUP%
         Printer.Print TTYY$
      End If
   Next II%
   TEPRUEBA.FontBold = True
   Printer.CurrentX = 80 + BIZQ%
   Printer.CurrentY = 240 + BSUP%
   Printer.Print "Pieza / Referencia"
End Sub
'
Sub PRISUBOR()
   '
   Static MAXSCROLL%
   '
   Printer.Font = TEPRUEBA.Font
   Printer.FontSize = TEPRUEBA.FontSize
   Printer.FontBold = TEPRUEBA.FontBold
   'INCREY = TEPRUEBA.Height * 1.05
   Printer.CurrentY = BSUP% + MARSU% + 2 * INCREYP 'Line1.Y1 - INCREY '/ 2
   '
   'If PRIOFVIS% < 1 Then PRIOFVIS% = 1
   AMAXI = Line2.X1 - 100
   For U& = 1 To CAOFA%
      Printer.CurrentY = Printer.CurrentY + 1.4 * INCREYP
      YINIP(U&) = Printer.CurrentY
      If YINIP(U&) <= BINF% - 2.5 * INCREYP Then
         TTXX$ = TRIM$(CODEXT$(CIOFAX%(U&))) + " (" + TRIM$(NORFBX$(U&)) + ")"
10       TEPRUEBA.Caption = TTXX$
         If TEPRUEBA.Width > AMAXI Then TTXX$ = Left$(TTXX$, Len(TTXX$) - 1): GoTo 10
         Printer.CurrentX = BIZQ% + 80
         Printer.Print TTXX$;
         TTXX$ = TRIM$(DESCRIT0$(CIOFAX%(U&)))
20       TEPRUEBA.Caption = TTXX$
         If TEPRUEBA.Width > AMAXI Then TTXX$ = Left$(TTXX$, Len(TTXX$) - 1): GoTo 20
         Printer.CurrentX = BIZQ% + 80
         Printer.CurrentY = Printer.CurrentY + INCREYP
         Printer.Print TTXX$;
         RECURY = Printer.CurrentY
         Printer.Line (BIZQ%, Printer.CurrentY + 1.25 * INCREYP)-(BDER%, Printer.CurrentY + 1.25 * INCREYP), QBColor(0)
         Printer.CurrentY = RECURY
      End If
   Next U&
   '
   If TRIM$(CODICON$) <> "" Then
      TTXX$ = "Código: " + CODICON$
      Printer.CurrentX = BIZQ% + 80
      Printer.CurrentY = BSUP% + MARSU% + 0.4 * INCREYP
      YINIP(0) = Printer.CurrentY
      Printer.Print TTXX$;
   End If
   '
   TTXX$ = TRIM$(DENOCON$)
25 TEPRUEBA.Caption = TTXX$
   If TEPRUEBA.Width > AMAXI Then TTXX$ = Left$(TTXX$, Len(TTXX$) - 1): GoTo 25
   Printer.CurrentX = BIZQ% + 80
   Printer.CurrentY = BSUP% + MARSU% + 1.4 * INCREYP
   Printer.Print TTXX$;
   '
End Sub
'
Sub PRIGANTT()
   '
   'CONTROCLICK = (-1)
   'TOPE = FORMUGANTT.List1.TopIndex
   'ACTI = FORMUGANTT.List1.ListIndex
   cabarr% = 0
   '
   X0MIN = 32767: X0MAX = 0
   For U& = 1 To CAOFA%
      NOFA$ = NORFBX$(U&)
      'NOFA$ = NORFB$(U&)
      COLO$ = Chr$(0) + Chr$(230) + Chr$(64)
      '
      'If FORMUGANTT.List1.Selected(U& - 1) = True Then
      '   COLO$ = Chr$(255) + Chr$(0) + Chr$(0)
      'End If
      '
      'If FORMUGANTT.List1.ListIndex = U& - 1 Then
      '   COLO$ = Chr$(128) + Chr$(0) + Chr$(128)
      'End If
      '
      V& = U&
      For J% = 1 To Len(MATIEM$(V&)) Step 36
         MAQUI$ = Mid$(MATIEM$(V&), J%, 24)
         FEHOR$ = Mid$(MATIEM$(V&), J% + 24, 8)
         DIVECTI& = CVS(Mid$(MATIEM$(V&), J% + 32, 4))
         FEINI% = CVI(Mid$(FEHOR$, 1, 2))
         HOINI% = CVI(Mid$(FEHOR$, 3, 2))
         FEFIN% = CVI(Mid$(FEHOR$, 5, 2))
         HOFIN% = CVI(Mid$(FEHOR$, 7, 2))
         If FEINI% > FEFIN% Then
            FEINI% = 0
            HOINI% = 0
            FEFIN% = 0
            HOFIN% = 0
         End If
         For K% = 1 To 24 Step 6
            EQUI$ = Mid$(MAQUI$, K%, 6)
            For KQ% = 1 To CAEQ%
               If CODEQ$(KQ%) = EQUI$ Then COLO$ = COLEQ$(KQ%)
            Next KQ%
            'If TRIM$(EQUI$) <> "" Then
               CBX% = cabarr%
               Call PRIBARR(FEINI%, HOINI%, FEFIN%, HOFIN%, NOFA$, EQUI$, COLO$, DIVECTI&)
               If cabarr% > CBX% Then
                  NUOFA$(cabarr%) = NOFA$
               End If
            'End If
         Next K%
      Next J%
      '
   Next U&
   '
   'HScroll1.LargeChange = CAPERIO%
   'CONTROCLICK% = 0
   'YACARGANTT% = 1
   '
End Sub

Sub PRIBARR(FEINI%, HOINI%, FEFIN%, HOFIN%, NOFA$, EQUI$, COLO$, DIVECTI&)
   '
   Static DHR0%, TUNOR
   If TUNOR < 1 Then
      'DHR0% = 2400: TUNOR = 0
      'For U& = 1 To ULTREG&("TURGEN")
      '   X$ = REGLEIDO$("TURGEN", U&)
      '   DSX% = Asc(Left$(X$, 1))
      '   If DSX% > 0 Then
      '      DHR% = CVI(Mid$(X$, 2, 2))
      '      HHR% = CVI(Mid$(X$, 4, 2))
      '      If HHR% > DHR% Then
      '         If DHR% < DHR0% Then DHR0% = DHR%
      '         TT1 = 60 * Int(DHR% / 100) + (DHR% Mod 100)
      '         TT2 = 60 * Int(HHR% / 100) + (HHR% Mod 100)
      '         If (TT2 - TT1) > TUNOR Then TUNOR = TT2 - TT1
      '      End If
      '   End If
      'Next U&
      'DHR0% = 60 * Int(DHR0% / 100) + (DHR0% Mod 100)
   End If
   TUNOR = 24 * 60: DHR0% = 0
   '
   If FEINI% > DESPERIO(CAPERIO% + 1) Then
      Exit Sub
   End If
   '
   If FEFIN% < DESPERIO(1) Then
      Exit Sub
   End If
   '
   If FEINI% < DESPERIO(1) Then
      X1 = XINIP(1)
      GoTo 10
   End If
   '
   For II% = 1 To CAPERIO%
      If FEINI% = DESPERIO(II%) Then
         MINU1% = 60 * Int(HOINI% / 100) + (HOINI% Mod 100)
         X1 = XINIP(II%) + APERIO * (MINU1% - DHR0%) / TUNOR
         GoTo 10
      End If
   Next II%
   Exit Sub
   '
10 If FEFIN% >= DESPERIO(CAPERIO% + 1) Then
      X2 = XINIP(CAPERIO%) + APERIO
      GoTo 20
   End If
   '
   For II% = 1 To CAPERIO%
      If FEFIN% = DESPERIO(II%) Then
         MINU2% = 60 * Int(HOFIN% / 100) + (HOFIN% Mod 100)
         X2 = XINIP(II%) + APERIO * (MINU2% - DHR0%) / TUNOR
         GoTo 20
      End If
   Next II%
   '
20 If X1 < X0MIN Then X0MIN = X1
   If X2 > X0MAX Then X0MAX = X2
   Y0 = YINIP(0) + INCREYP
   Printer.Line (X0MIN, Y0)-(X0MAX, Y0 + 145), RGB(128, 128, 128), BF
   Printer.Line (X0MIN + 12, Y0 + 145)-(X0MAX, Y0 + 157), QBColor(0), BF
   Printer.Line (X0MAX - 12, Y0 + 25)-(X0MAX, Y0 + 157), QBColor(0), BF
   '
   For II% = 1 To CAOFA%
      If NORFBX$(II%) = NOFA$ Then
         Y1 = YINIP(II%) + INCREYP
         GoTo 30
      End If
   Next II%
   'Exit Sub
   '
30 If Y1 < 100 Then Exit Sub
   'If X2 > XMAXI Then X2 = XMAXI
   ROJO% = Asc(Left$(COLO$, 1))
   VERDE% = Asc(Mid$(COLO$, 2, 1))
   AZUL% = Asc(Mid$(COLO$, 3, 1))
   Printer.Line (X1, Y1 + 90)-(X2, Y1 + 145), RGB(ROJO%, VERDE%, AZUL%), BF
   Printer.Line (X1 + 12, Y1 + 145)-(X2, Y1 + 157), QBColor(0), BF
   Printer.Line (X2 - 12, Y1 + 25)-(X2, Y1 + 157), QBColor(0), BF
   X2MA = X2
   '
   cabarr% = cabarr% + 1
   XINI1(cabarr%) = X1
   XFIN1(cabarr%) = X2
   YINI1(cabarr%) = Y1
   '
   If DIVECTI& > 0 Then
      If DIVECTI& <= ULTREG&("VECTIPAR") Then
         VECTIPAR$ = REGLEIDO$("VECTIPAR", DIVECTI&)
         LOVECTI% = CVI(Left$(VECTIPAR$, 2))
         If LOVECTI% > 254 Then LOVECTI% = 254
         VECTIPAR$ = Mid$(VECTIPAR$, 3) + String$(6, 0)
         For KKI% = 1 To LOVECTI% Step 6
            FFII% = CVI(Mid$(VECTIPAR$, KKI%, 2))
            HHII% = CVI(Mid$(VECTIPAR$, KKI% + 2, 2))
            HHFF% = CVI(Mid$(VECTIPAR$, KKI% + 4, 2))
            If HHFF% < HHII% Then HHFF% = HHII%
            '
            If FFII% < DESPERIO(1) Or FFII% > DESPERIO(CAPERIO%) Then
               GoTo 190
            End If
            '
            For II% = 1 To CAPERIO%
               If FFII% = DESPERIO(II%) Then
                  MINU1% = 60 * Int(HHII% / 100) + (HHII% Mod 100)
                  X1 = XINIP(II%) + APERIO * (MINU1% - DHR0%) / TUNOR
                  GoTo 110
               End If
            Next II%
            Exit Sub
            '
110         If FFII% >= DESPERIO(CAPERIO% + 1) Then
               X2 = XINI(CAPERIO%) + APERIO
               GoTo 130
            End If
            '
            For II% = 1 To CAPERIO%
               If FFII% = DESPERIO(II%) Then
                  MINU2% = 60 * Int(HHFF% / 100) + (HHFF% Mod 100)
                  X2 = XINIP(II%) + APERIO * (MINU2% - DHR0%) / TUNOR
                  GoTo 130
               End If
            Next II%
            '
130         If Y1 < 100 Then Exit Sub
            If X2 > X2MA - 15 Then X2 = X2MA - 15
            ROJO% = Asc(Left$(COLO$, 1))
            VERDE% = Asc(Mid$(COLO$, 2, 1))
            AZUL% = Asc(Mid$(COLO$, 3, 1))
            Printer.Line (X1, Y1)-(X2, Y1 + 100), RGB(ROJO%, VERDE%, AZUL%), BF
            '
190      Next KKI%
      End If
   End If
   '
End Sub
'
Sub PRIORTOT()
   '
   Static MAXSCROLL%
   '
   CAOTOT% = 0
   For I& = 1 To ULTREG&("ORFANUE")
      XX$ = REGLEIDO$("ORFANUE", I&)
      ST% = Asc(Mid$(XX$, 112, 1))
      '
      If ST% = 1 Then
         If CAOTOT% < 512 Then
            CAOTOT% = CAOTOT% + 1
            OTOT&(CAOTOT%) = CVS(Left$(XX$, 4))
            CIOTOT%(CAOTOT%) = CVI(Mid$(XX$, 71, 2))
         End If
      End If
   Next I&
   '
   For I& = 1 To ULTREG&("OPERFAP")
      XX$ = REGLEIDO$("OPERFAP", I&)
      NOFAPX& = Val(Mid$(XX$, 4, 6))
      For J& = 1 To CAOTOT%
         If OTOT&(J&) = NOFAPX& Then GoTo 10
      Next J&
      GoTo 19
      '
10    INIX# = 10000 * CDbl(CVI(Mid$(XX$, 135, 2))) + CDbl(CVI(Mid$(XX$, 137, 2)))
      FINX# = 10000 * CDbl(CVI(Mid$(XX$, 139, 2))) + CDbl(CVI(Mid$(XX$, 141, 2)))
      If FINX# < INIX# Then FINX# = INIX#
      '
      If COMIE#(J&) < 1 Then
         COMIE#(J&) = INIX#
         FINAL#(J&) = FINX#
      End If
      If INIX# < COMIE#(J&) Then COMIE#(J&) = INIX#
      If FINX# > FINAL#(J&) Then FINAL#(J&) = FINX#
      '
19 Next I&
   '
   CAOTOT1% = 0
   For I& = 1 To CAOTOT%
      If Int(COMIE#(I&) / 10000) > DESPERIO(CAPERIO%) Then GoTo 22
      If Int(FINAL#(I&) / 10000) < DESPERIO(1) Then GoTo 22
      CAOTOT1% = CAOTOT1% + 1
      OTOT&(CAOTOT1%) = OTOT&(I&)
      CIOTOT%(CAOTOT1%) = CIOTOT%(I&)
      COMIE#(CAOTOT1%) = COMIE#(I&)
      FINAL#(CAOTOT1%) = FINAL#(I&)
22 Next I&
   '
   Printer.Font = TEPRUEBA.Font
   Printer.FontSize = TEPRUEBA.FontSize
   Printer.FontBold = TEPRUEBA.FontBold
   Printer.CurrentY = BSUP% + MARSU% - 0.5 * INCREYP 'Line1.Y1 - INCREY '/ 2
   '
   'If PRIOFVIS% < 1 Then PRIOFVIS% = 1
   AMAXI = Line2.X1 - 100
   For U& = 1 To CAOTOT1%
      Printer.CurrentY = Printer.CurrentY + 1.4 * INCREYP
      YINIP(U&) = Printer.CurrentY
      If YINIP(U&) <= BINF% - 2.5 * INCREYP Then
         TTXX$ = "Código: " + TRIM$(CODEXT$(CIOTOT%(U&)))
30       TEPRUEBA.Caption = TTXX$
         If TEPRUEBA.Width > AMAXI Then TTXX$ = Left$(TTXX$, Len(TTXX$) - 1): GoTo 30
         Printer.CurrentX = BIZQ% + 80
         Printer.Print TTXX$;
         TTXX$ = TRIM$(DESCRIT0$(CIOTOT%(U&)))
40       TEPRUEBA.Caption = TTXX$
         If TEPRUEBA.Width > AMAXI Then TTXX$ = Left$(TTXX$, Len(TTXX$) - 1): GoTo 40
         Printer.CurrentX = BIZQ% + 80
         Printer.CurrentY = Printer.CurrentY + INCREYP
         YINIQ(U&) = Printer.CurrentY
         Printer.Print TTXX$;
         RECURY = Printer.CurrentY
         Printer.Line (BIZQ%, Printer.CurrentY + 1.25 * INCREYP)-(BDER%, Printer.CurrentY + 1.25 * INCREYP + 5), QBColor(0)
         Printer.CurrentY = RECURY
      End If
   Next U&
   '
End Sub

Sub PRIGATOT()
   '
   cabarr% = 0
   '
   X0MIN = 32767: X0MAX = 0
   For U& = 1 To CAOTOT1%
      COLO$ = Chr$(0) + Chr$(230) + Chr$(64)
      V& = U&
      FEINI% = Int(COMIE#(U&) / 10000)
      HOINI% = COMIE#(U&) - 10000# * CDbl(FEINI%)
      FEFIN% = Int(FINAL#(U&) / 10000)
      HOFIN% = FINAL#(U&) - 10000# * CDbl(FEFIN%)
      Y00 = YINIQ(U&)
      Call PRIBATOT(FEINI%, HOINI%, FEFIN%, HOFIN%, Y00, COLO$)
   Next U&
   '
End Sub
'
Sub PRIBATOT(FEINI%, HOINI%, FEFIN%, HOFIN%, Y00, COLO$)
   '
   Static DHR0%, TUNOR
   TUNOR = 24 * 60: DHR0% = 0
   '
   If FEINI% > DESPERIO(CAPERIO% + 1) Then
      Exit Sub
   End If
   '
   If FEFIN% < DESPERIO(1) Then
      Exit Sub
   End If
   '
   If FEINI% < DESPERIO(1) Then
      X1 = XINIP(1)
      GoTo 10
   End If
   '
   For II% = 1 To CAPERIO%
      If FEINI% = DESPERIO(II%) Then
         MINU1% = 60 * Int(HOINI% / 100) + (HOINI% Mod 100)
         X1 = XINIP(II%) + APERIO * (MINU1% - DHR0%) / TUNOR
         GoTo 10
      End If
   Next II%
   Exit Sub
   '
10 If FEFIN% >= DESPERIO(CAPERIO% + 1) Then
      X2 = XINIP(CAPERIO%) + APERIO
      GoTo 20
   End If
   '
   For II% = 1 To CAPERIO%
      If FEFIN% = DESPERIO(II%) Then
         MINU2% = 60 * Int(HOFIN% / 100) + (HOFIN% Mod 100)
         X2 = XINIP(II%) + APERIO * (MINU2% - DHR0%) / TUNOR
         GoTo 20
      End If
   Next II%
   '
20 X0MIN = X1
   X0MAX = X2
   Y0 = Y00
   Printer.Line (X0MIN, Y0)-(X0MAX, Y0 + 145), RGB(128, 128, 128), BF
   Printer.Line (X0MIN + 12, Y0 + 145)-(X0MAX, Y0 + 157), QBColor(0), BF
   Printer.Line (X0MAX - 12, Y0 + 25)-(X0MAX, Y0 + 157), QBColor(0), BF
   '
End Sub
'

