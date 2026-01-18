VERSION 5.00
Begin VB.Form FORGANTT 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Secuencia de Carga por Equipo Productivo"
   ClientHeight    =   8355
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   11910
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8355
   ScaleWidth      =   11910
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00000000&
      Height          =   850
      Left            =   0
      ScaleHeight     =   795
      ScaleWidth      =   11940
      TabIndex        =   4
      Top             =   7560
      Width           =   12000
      Begin VB.PictureBox Picture4 
         Height          =   570
         Left            =   50
         ScaleHeight     =   510
         ScaleWidth      =   1185
         TabIndex        =   20
         Top             =   105
         Width           =   1250
         Begin VB.Image Image2 
            Height          =   510
            Left            =   -380
            Picture         =   "FORGANTT.frx":0000
            Top             =   0
            Width           =   2010
         End
      End
      Begin VB.ListBox List3 
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   570
         IntegralHeight  =   0   'False
         Left            =   3300
         Sorted          =   -1  'True
         TabIndex        =   19
         Top             =   105
         Width           =   1905
      End
      Begin VB.PictureBox Picture3 
         Height          =   570
         Left            =   5250
         ScaleHeight     =   510
         ScaleWidth      =   1425
         TabIndex        =   14
         Top             =   105
         Width           =   1485
         Begin VB.OptionButton Option1 
            BackColor       =   &H00C0C0C0&
            Caption         =   "M"
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
            Left            =   840
            TabIndex        =   18
            Top             =   260
            Width           =   435
         End
         Begin VB.OptionButton Option1 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Q"
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
            Left            =   630
            TabIndex        =   17
            Top             =   50
            Width           =   435
         End
         Begin VB.OptionButton Option1 
            BackColor       =   &H00C0C0C0&
            Caption         =   "S"
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
            TabIndex        =   16
            Top             =   260
            Value           =   -1  'True
            Width           =   435
         End
         Begin VB.OptionButton Option1 
            BackColor       =   &H00C0C0C0&
            Caption         =   "D"
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
            Left            =   105
            TabIndex        =   15
            Top             =   50
            Width           =   435
         End
      End
      Begin VB.CommandButton Command8 
         Height          =   590
         Left            =   7455
         Picture         =   "FORGANTT.frx":35EA
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Consultas por Pantalla"
         Top             =   105
         Width           =   550
      End
      Begin VB.CommandButton Command7 
         Height          =   590
         Left            =   8085
         Picture         =   "FORGANTT.frx":3734
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Listados por Impresora"
         Top             =   105
         Width           =   550
      End
      Begin VB.CommandButton Command6 
         Height          =   590
         Left            =   8715
         Picture         =   "FORGANTT.frx":3A3E
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Acceso Directo a Aplicaciones Complementarias"
         Top             =   105
         Width           =   550
      End
      Begin VB.CommandButton Command4 
         Height          =   590
         Left            =   9345
         Picture         =   "FORGANTT.frx":3D48
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Acceso a Archivos Maestros"
         Top             =   105
         Width           =   550
      End
      Begin VB.CommandButton Command3 
         Height          =   590
         Left            =   9975
         Picture         =   "FORGANTT.frx":4052
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Configuración de Funcionamiento"
         Top             =   105
         Width           =   550
      End
      Begin VB.ListBox List1 
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   570
         Left            =   1350
         Sorted          =   -1  'True
         Style           =   1  'Checkbox
         TabIndex        =   8
         Top             =   105
         Width           =   1905
      End
      Begin VB.CommandButton Command5 
         Height          =   590
         Left            =   6825
         Picture         =   "FORGANTT.frx":4494
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Impresión Gráfico de Carga"
         Top             =   105
         Width           =   550
      End
      Begin VB.CommandButton Command2 
         Height          =   590
         Left            =   10605
         Picture         =   "FORGANTT.frx":4AFE
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Re-Cálcule General de Carga de Máquinas"
         Top             =   105
         Width           =   550
      End
      Begin VB.CommandButton Command1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   590
         Left            =   11235
         Picture         =   "FORGANTT.frx":4E08
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Terminar - Salir de la Aplicación"
         Top             =   105
         Width           =   550
      End
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00E0E0E0&
      Height          =   7600
      Left            =   0
      ScaleHeight     =   7545
      ScaleWidth      =   11895
      TabIndex        =   0
      Top             =   0
      Width           =   11950
      Begin VB.HScrollBar HScroll1 
         Height          =   225
         Left            =   840
         TabIndex        =   3
         Top             =   7300
         Width           =   10800
      End
      Begin VB.VScrollBar VScroll1 
         Height          =   6730
         Left            =   11655
         TabIndex        =   2
         Top             =   540
         Width           =   225
      End
      Begin VB.Line Line4 
         X1              =   11640
         X2              =   11640
         Y1              =   8085
         Y2              =   0
      End
      Begin VB.Line Line3 
         BorderWidth     =   2
         X1              =   0
         X2              =   12000
         Y1              =   7290
         Y2              =   7290
      End
      Begin VB.Line Line1 
         BorderWidth     =   2
         X1              =   0
         X2              =   12000
         Y1              =   525
         Y2              =   525
      End
      Begin VB.Line Line2 
         BorderWidth     =   2
         X1              =   840
         X2              =   840
         Y1              =   0
         Y2              =   8295
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
      Left            =   6090
      TabIndex        =   1
      Top             =   5355
      Visible         =   0   'False
      Width           =   1020
   End
End
Attribute VB_Name = "FORGANTT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim XMAXI
Dim MATIE$(1024)
Dim CAOFA%, CACODACT%
Dim TBAJACLI

Private Sub Command1_Click()
   Call CIERRARCH("*.*")
   Call FINAL("")
End Sub

Private Sub Command2_Click()
    '
    Command2.Enabled = False
    '
    COMPLECALCU% = 0
    CALCUCAR.Show 1
    '
    If COMPLECALCU% = 1 Then
      Call CIERRARCH("*.*")
      Call LIBARCH("*.*")
      Call FILESORT("OPERFAP", "", 20, 21)
      Call COMUNI("Proceso Completo")
      If EXISTE%(App.EXEName + ".EXE") > 0 Then
        Call BLOQEXE("*")
        Call CONECRUN(App.EXEName, "Programación de Carga de Máquinas")
        Call FINAL("")
      End If
    End If
    '
    Command2.Enabled = True
    '
End Sub

Private Sub Command3_Click()
   Command3.Enabled = False
   CONFICRP.Show 1
   Command3.Enabled = True
End Sub

Private Sub Command4_Click()
   Command4.Enabled = False
   Call CONECRUN("ARCHIMA", "Archivos Maestros")
   Command4.Enabled = True
End Sub

Private Sub Command5_Click()
   Command5.Enabled = False
   ANCHAN = Width
   Width = Printer.ScaleWidth
   PRINTFORM
   Width = ANCHAN
   Command5.Enabled = True
End Sub

Private Sub Command6_Click()
   Command6.Enabled = False
   ADIRCARG.Show 1
   '
   If ATENTI% > 0 Then ' SE INFORMARON OPERACIONES
     If COMALTER%("Debe Ahora Re-Calcular la Carga de Máquinas\\Si, Re-calcular\No, Continuar") = 1 Then
       '
       COMPLECALCU% = 0
       CALCUCAR.Show 1
       '
       If COMPLECALCU% = 1 Then
         Call CIERRARCH("*.*")
         Call LIBARCH("*.*")
         Call FILESORT("OPERFAP", "", 20, 21)
         Call COMUNI("Proceso Completo")
         If EXISTE%(App.EXEName + ".EXE") > 0 Then
           Call BLOQEXE("*")
           Call CONECRUN(App.EXEName, "Programación de Carga de Máquinas")
           Call FINAL("")
         End If
       End If
       '
     End If
   End If
   '
   Command6.Enabled = True
End Sub

Private Sub Command7_Click()
   Command7.Enabled = False
   LISCARMA.Show 1
   Command7.Enabled = True
End Sub

Private Sub Command8_Click()
   Command8.Enabled = False
   CONCARMA.Show 1
   Command8.Enabled = True
End Sub

Private Sub Form_Load()
   '
   VERANTER$ = "CARGMAQ"
   Call VERJOBID(OKEY%)
   If OKEY% < 1 Then Call FINAL("")
   '
   EPAC$ = TRIM$(EMPREAC$)
   If EPAC$ <> "" Then
     '
10   Caption = Caption + "  -  " + EPAC$
     'APPX$ = App.EXEName
     'If EXISTE%(APPX$ + ".KKK") > 0 Then GoTo 10
     'FMDX = FECHMODU%
     'BB1 = HOY(HOS$)
     'If DIENTRE%(FMDX, HOY(HOS$)) > 100 Then
     '  Call SAVEFILE(APPX$ + ".KKK", String$(128, 0))
     'End If
     '
   End If
   Call GRATITFOR(Caption)
   '
   FEIPER = HOY(HO$)
   Screen.MousePointer = 11
   For U& = 1 To ULTREG&("OPERFAP")
     XX$ = REGLEIDO$("OPERFAP", U&)
     FECOMI = CVI(Mid$(XX$, 135, 2))
     If FECOMI >= 1 Then
       If FECOMI < FEIPER Then
         FEIPER = FECOMI
       End If
     End If
   Next U&
   Screen.MousePointer = 1
   'While DIASEM%(FEIPER) <> 1
   '   FEIPER = DIANTE(FEIPER)
   'Wend
   '
   Show
   Refresh
   HScroll1.Min = FEIPER - 32 * 13
   HScroll1.Max = FEIPER + 32 * 13
   HScroll1.Value = FEIPER
   '
   Call DIBULINVER
   Call CALCUFECHA
   Call DIBUFECHA
   Call CARLIORF
   Call DIBUEQUIP
   Call PREGANTT
   'If List1.ListIndex >= 0 Then
   '   NORFA$ = TRIM$(Left$(List1.TEXT, 24))
   '   Call MUEDATORF(NORFA$)
   'End If
   '
   YACARGANTT% = 1
   '
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call CIERRARCH("*.*")
   Call FINAL("")
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
   Call DIBULINVER
   Call CALCUFECHA
   Call DIBUFECHA
   Call DIBUEQUIP
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

Private Sub List3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   If Button = 2 Then
     Call SELECHO("!CODIORFA")
     CDXX$ = TRIM$(VALACT1$("!CODIORFA"))
     If CDXX$ <> "" Then
       For UK& = 1 To List3.ListCount
         If TRIM$(List3.List(UK& - 1)) = CDXX$ Then
           List3.ListIndex = (UK& - 1)
           Exit For
         End If
       Next UK&
     End If
   End If
End Sub

Private Sub Option1_Click(Index As Integer)
   Picture1.Cls
   Call DIBULINVER
   Call CALCUFECHA
   Call DIBUFECHA
   Call DIBUEQUIP
   Call PREGANTT
End Sub

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   XMUOFA = MUORFACT.Left
   YMUOFA = MUORFACT.Top
   MUORFACT.Hide
   DoEvents
   '
   If X < Line2.X1 Then
     For UI% = PRIEQVIS% To CAEQ%
       If Y >= YEQU(UI%) Then
         If Y <= YEQU(UI%) + 250 Then
           COEQUIPX$ = CODEQ$(UI%)
           Picture1.ToolTipText = TRIM$(ECOARCH$("PADMAQ", COEQUIPX$))
           TBAJACLI = Timer
         End If
       End If
     Next UI%
     Exit Sub
   End If
   '
   For UI% = 1 To cabarr%
     If X >= XINI1(UI%) Then
       If X <= XFIN1(UI%) Then
         If Y >= YINI1(UI%) Then
           If Y <= YINI1(UI%) + 145 Then
             NORFAX$ = NUOFA$(UI%)
             For U& = 1 To List1.ListCount
               V& = Val(Mid$(List1.List(U& - 1), 25))
               If NORFB$(V&) = NORFAX$ Then
                 CONTROCLICK% = -1
                 List1.ListIndex = U& - 1
                 CONTROCLICK% = 0
                 OFACT& = U&
                 Call PREGANTT
                 Call MUEDATORF(NORFAX$)
                 Exit Sub
               End If
             Next U&
           End If
         End If
       End If
     End If
   Next UI%
End Sub


Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   If Timer > TBAJACLI + 1 Then
      Picture1.ToolTipText = ""
   End If
End Sub

Private Sub Picture1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
   If Timer > TBAJACLI + 1 Then
      Picture1.ToolTipText = ""
   End If
End Sub

Private Sub VScroll1_Change()
   PRIEQVIS% = VScroll1.Value + 1
   Picture1.Cls
   Call DIBULINVER
   Call CALCUFECHA
   Call DIBUFECHA
   Call DIBUEQUIP
   Call PREGANTT
End Sub

Sub DIBULINVER()
   CAPERIO% = 7
   CALINHOR% = 12: INTERLIN% = 3
   If Option1(0) = True Then
        CAPERIO% = 2
        CALINHOR% = 24: INTERLIN% = 4
      ElseIf Option1(2) = True Then
        CAPERIO% = 15
        CALINHOR% = 6: INTERLIN% = 3
      ElseIf Option1(3) = True Then
        CAPERIO% = 31
        CALINHOR% = 2: INTERLIN% = 2
   End If
   Call CALCUFECHA
   APERIO = (Picture1.Width - Line2.X1 - VScroll1.Width - 80) / (CAPERIO%)
   DrawWidth = 2
   X1 = Line2.X1
   Y2 = Picture1.Height
   Y3 = Line1.Y1
   For II% = 1 To CAPERIO% + 1
      DSXX% = DIASEM%(DESPERIO(II%))
      If II% <= CAPERIO% Then
        If DSXX% > 5 Or DSXX% < 1 Then
          Picture1.Line (X1 + 20, YH1)-(X1 + APERIO, Y2), RGB(160, 160, 160), BF
        End If
      End If
      XINI(II%) = X1
      X2 = X1
      X1 = X1 + APERIO
      If II% = CAPERIO% Then XMAXI = X1 - 20
      Picture1.Line (X1, 0)-(X1, Y2), QBColor(0)
      '
      If II% <= CAPERIO% Then
        For JJ% = 1 To CALINHOR% - 1
          X3 = X2 + JJ% * APERIO / CALINHOR%
          COLOLIN = RGB(200, 200, 200)
          If JJ% Mod INTERLIN% = 0 Then COLOLIN = RGB(140, 140, 140)
          Picture1.Line (X3, Y3)-(X3, Y2), COLOLIN
        Next JJ%
      End If
   Next II%
   Picture1.Refresh
End Sub
'
Sub DIBUFECHA()
   '
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
End Sub
'
Sub DIBUEQUIP()
   Static MAXSCROLL%
   '
   Picture1.Font = TEPRUEBA.Font
   Picture1.FontSize = TEPRUEBA.FontSize
   Picture1.FontBold = TEPRUEBA.FontBold
   INCREY = TEPRUEBA.Height * 1.2
   Picture1.CurrentY = Line1.Y1 - INCREY / 2
   '
   If PRIEQVIS% < 1 Then PRIEQVIS% = 1
   For U& = PRIEQVIS% To CAEQ%
      Picture1.CurrentY = Picture1.CurrentY + INCREY
      YINI(U&) = Picture1.CurrentY
      If YINI(U&) <= Picture1.Height - HScroll1.Height - INCREY Then
         Picture1.CurrentX = 80
         Picture1.Print CODEQ$(U&);
         YEQU(U&) = YINI(U&)
      End If
   Next U&
   '
20 If CAEQVIS% = 0 Then
      CAEQVIS% = (Picture1.Height - Line1.Y1 - HScroll1.Height - INCREY) / INCREY
   End If
   '
   If MAXSCROLL% <= 0 Then
      VScroll1.Max = 0
      If CAEQVIS% < CAEQ% Then
         VScroll1.Max = CAEQ% - CAEQVIS%
      End If
      MAXSCROLL% = VScroll1.Max
   End If
   VScroll1.Max = MAXSCROLL%
   VScroll1.Value = PRIEQVIS% - 1
   '
End Sub
'
Sub PREGANTT()
   '
   CONTROCLICK = (-1)
   'TOPE = List1.TopIndex
   'ACTI = List1.ListIndex
   cabarr% = 0
   NOFASE$ = TRIM$(Left$(List1.TEXT, 24))
   For U& = 1 To List1.ListCount
      NOFA$ = TRIM$(Left$(List1.List(U& - 1), 24))
      'NOFA$ = NORFB$(U&)
      COLO$ = Chr$(0) + Chr$(230) + Chr$(64)
      '
      If List1.Selected(U& - 1) = True Then
         COLO$ = Chr$(255) + Chr$(0) + Chr$(0)
      End If
      '
      If List1.ListIndex = U& - 1 Then
           COLO$ = Chr$(128) + Chr$(0) + Chr$(128)
        ElseIf Left$(NOFA$, 9) = Left$(NOFASE$, 9) Then
           COLO$ = Chr$(200) + Chr$(0) + Chr$(200)   ' PARA COLOREAR LAS ASOCIADAS
      End If
      '
      '
      COLOPRE$ = Chr$(220) + Chr$(220) + Chr$(0)
      V& = Val(Mid$(List1.List(U& - 1), 25))
      For J% = 1 To Len(MATIE$(V&)) Step 32
         MAQUI$ = Mid$(MATIE$(V&), J%, 24)
         FEHOR$ = Mid$(MATIE$(V&), J% + 24, 12)
         FEINI% = CVI(Mid$(FEHOR$, 1, 2))
         HOINI% = CVI(Mid$(FEHOR$, 3, 2))
         FEFIN% = CVI(Mid$(FEHOR$, 5, 2))
         HOFIN% = CVI(Mid$(FEHOR$, 7, 2))
         FEFINPREP% = CVI(Mid$(FEHOR$, 9, 2))
         HOFINPREP% = CVI(Mid$(FEHOR$, 11, 2))
         If FEINI% > FEFIN% Then
            FEINI% = 0
            HOINI% = 0
            FEFIN% = 0
            HOFIN% = 0
            FEFINPREP% = 0
            HOFINPREP% = 0
         End If
         For K% = 1 To 24 Step 6
            EQUI$ = Mid$(MAQUI$, K%, 6)
            If TRIM$(EQUI$) <> "" Then
               CBX% = cabarr%
               Call DIBUBARRA(FEINI%, HOINI%, FEFIN%, HOFIN%, EQUI$, COLO$, 0)
               Call DIBUBARRA(FEINI%, HOINI%, FEFINPREP%, HOFINPREP%, EQUI$, COLOPRE$, 1)
               If cabarr% > CBX% Then
                  NUOFA$(cabarr%) = NOFA$
               End If
            End If
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
Sub CARLIORF()
   '
   List1.Clear
   List3.Clear: JJ& = 0
   Call COPYSTRU("ECOMAST", "CODIORFA")
   CAOFA% = 0
   CAEQ% = 0
   For U& = 1 To ULTREG&("PADMAQ")
      ZZ$ = REGLEIDO$("PADMAQ", U&)
      If TRIM$(Left$(ZZ$, 6)) <> "" Then
         CAEQ% = CAEQ% + 1
         CODEQ$(CAEQ%) = Left$(ZZ$, 6)
         STATEQ%(CAEQ%) = 0
      End If
   Next U&
   '
   For U& = 1 To ULTREG&("OPERFAP")
      '
      XX$ = REGLEIDO$("OPERFAP", U&)
      NOFA$ = TRIM$(Left$(XX$, 12))
      MAQUI$ = Mid$(XX$, 73, 6) + Mid$(XX$, 203, 18)
      'ITIEM$ = Mid$(XX$, 135, 8)
      ITIEM$ = Mid$(XX$, 135, 12)
      CIIX% = CVI(Mid$(XX$, 13, 2))
      '
      YACODI% = 0
      For K% = 1 To CAOFA%
         If CIOFA%(K%) = CIIX% Then YACODI% = 1
         If NORFB$(K%) = NOFA$ Then GoTo 10
      Next K%
      CAOFA% = CAOFA% + 1
      K% = CAOFA%
      NORFB$(CAOFA%) = NOFA$
      CIOFA%(CAOFA%) = CIIX%
      MATIE$(CAOFA%) = ""
      TTXX$ = Space$(32)
      Call REPLA(TTXX$, NOFA$, 1, 24)
      Call REPLA(TTXX$, Str$(CAOFA%), 25, 8)
      List1.AddItem TTXX$
      If YACODI% = 0 Then
         List3.AddItem CODEXT$(CIIX%)
         TTYZ$ = Space$(64)
         Call REPLA(TTYZ$, CODEXT$(CIIX%), 1, 16)
         Call REPLA(TTYZ$, DESCRIT$(CIIX%), 17, 48)
         JJ& = JJ& + 1
         Call GRAREG("CODIORFA", TTYZ$, JJ&)
      End If
      '
10    MATIE$(K%) = MATIE$(K%) + MAQUI$ + ITIEM$
      '
      For K% = 1 To 24 Step 6
         EQUIPO$ = Mid$(MAQUI$, K%, 6)
         If TRIM$(EQUIPO$) <> "" Then
            For J% = 1 To CAEQ%
               If CODEQ$(J%) = EQUIPO$ Then
                  STATEQ%(J%) = 1
                  GoTo 20
               End If
            Next J%
            Call MENSERR(24, "Equipo '" + TRIM$(EQUIPO$) + "'no Figura en Tablas\en O/F Número " + TRIM$(NOFA$) + "para Código " + TRIM$(CODEXT$(CIIX%)) + ".")
         End If
20    Next K%
      '
   Next U&
   Call SETULTREG("CODIORFA", JJ&)
   Call FILESORT("CODIORFA", "", 1, 1)
   Call TRALOCAL("CODIORFA", "")
   '
   J% = 0
   For I% = 1 To CAEQ%
      If STATEQ%(I%) = 1 Then
         J% = J% + 1
         CODEQ$(J%) = CODEQ$(I%)
         STATEQ%(J%) = 1
      End If
   Next I%
   CAEQ% = J%
   '
End Sub
'
Sub DIBUBARRA(FEINI%, HOINI%, FEFIN%, HOFIN%, EQUI$, COLO$, PREPAR%)
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
      X2 = XINI(CAPERIO% + 1)
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
20 For II% = PRIEQVIS% To CAEQ%
      If CODEQ$(II%) = EQUI$ Then
         Y1 = YINI(II%)
         GoTo 30
      End If
   Next II%
   Exit Sub
   '
30 If Y1 < 100 Then Exit Sub
   If X2 > XMAXI Then X2 = XMAXI
   ROJO% = Asc(Left$(COLO$, 1))
   VERDE% = Asc(Mid$(COLO$, 2, 1))
   AZUL% = Asc(Mid$(COLO$, 3, 1))
   Picture1.Line (X1, Y1 + PREPAR% * 60)-(X2, Y1 + 145), RGB(ROJO%, VERDE%, AZUL%), BF
   Picture1.Line (X1 + 12, Y1 + 145)-(X2, Y1 + 157), QBColor(0), BF
   Picture1.Line (X2 - 12, Y1 + 25)-(X2, Y1 + 157), QBColor(0), BF
   '
   If PREPAR% = 0 Then
     cabarr% = cabarr% + 1
     XINI1(cabarr%) = X1
     XFIN1(cabarr%) = X2
     YINI1(cabarr%) = Y1
   End If
   '
End Sub
'

Sub MUEDATORF(NORFA$)
   '
   MUORFACT.Hide
   DoEvents
   MUORFACT.Label2 = NORFA$
   MUORFACT.Left = XMUOFA
   MUORFACT.Top = YMUOFA
   MUORFACT.Show ' 1
   '
End Sub






