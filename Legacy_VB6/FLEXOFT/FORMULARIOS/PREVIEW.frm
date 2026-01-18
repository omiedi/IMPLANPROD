VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form PREVIEW 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   8445
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11910
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   8445
   ScaleWidth      =   11910
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin ComctlLib.Toolbar Toolbar1 
      Height          =   495
      Left            =   9660
      TabIndex        =   10
      Top             =   80
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   873
      ButtonWidth     =   794
      ButtonHeight    =   714
      Appearance      =   1
      ImageList       =   "imlToolbarPics"
      _Version        =   327682
      BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
         NumButtons      =   5
         BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "fullscreen"
            Object.ToolTipText     =   "Pantalla Completa / Reducida"
            Object.Tag             =   ""
            ImageIndex      =   15
         EndProperty
         BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "Search"
            Object.ToolTipText     =   "Busqueda por Texto"
            Object.Tag             =   ""
            ImageIndex      =   14
         EndProperty
         BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "PRINT"
            Object.ToolTipText     =   "Imprimir"
            Object.Tag             =   ""
            ImageIndex      =   12
         EndProperty
         BeginProperty Button4 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Object.ToolTipText     =   "Ayuda en Línea"
            Object.Tag             =   ""
            ImageIndex      =   13
         EndProperty
         BeginProperty Button5 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "EXIT"
            Object.ToolTipText     =   "Cerrar Consulta"
            Object.Tag             =   ""
            ImageIndex      =   11
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin VB.PictureBox Picture9 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   450
      Left            =   100
      ScaleHeight     =   450
      ScaleWidth      =   2010
      TabIndex        =   16
      Top             =   50
      Width           =   2010
      Begin VB.Image Image1 
         Height          =   450
         Left            =   0
         Top             =   0
         Width           =   1380
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Font"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   580
      Left            =   6195
      TabIndex        =   11
      Top             =   -20
      Width           =   3375
      Begin VB.CommandButton Command1 
         Caption         =   "B"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   310
         Left            =   2940
         TabIndex        =   15
         Top             =   200
         Width           =   330
      End
      Begin VB.VScrollBar VScroll1 
         Height          =   310
         Left            =   2625
         Max             =   20
         Min             =   1
         TabIndex        =   14
         Top             =   200
         Value           =   20
         Width           =   175
      End
      Begin VB.TextBox Text1 
         Height          =   310
         Left            =   2310
         TabIndex        =   13
         Text            =   " "
         Top             =   200
         Width           =   330
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   525
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   200
         Width           =   1695
      End
   End
   Begin VB.VScrollBar VScroll2 
      Height          =   7080
      Left            =   11580
      TabIndex        =   6
      Top             =   1080
      Width           =   225
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   225
      Left            =   100
      Max             =   10000
      TabIndex        =   5
      Top             =   8120
      Width           =   11505
   End
   Begin VB.PictureBox Picture2 
      Height          =   7530
      Left            =   100
      ScaleHeight     =   7470
      ScaleWidth      =   11445
      TabIndex        =   0
      Top             =   620
      Width           =   11500
      Begin VB.ListBox List4 
         BeginProperty Font 
            Name            =   "Lucida Console"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   7100
         IntegralHeight  =   0   'False
         ItemData        =   "PREVIEW.frx":0000
         Left            =   -30
         List            =   "PREVIEW.frx":0002
         TabIndex        =   1
         Top             =   420
         Width           =   40000
      End
      Begin VB.PictureBox Picture7 
         Height          =   500
         Left            =   -30
         ScaleHeight     =   435
         ScaleWidth      =   89940
         TabIndex        =   2
         Top             =   -30
         Width           =   90000
         Begin VB.ListBox List1 
            BackColor       =   &H00C0C0C0&
            BeginProperty Font 
               Name            =   "Lucida Console"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   480
            IntegralHeight  =   0   'False
            ItemData        =   "PREVIEW.frx":0004
            Left            =   -30
            List            =   "PREVIEW.frx":0006
            TabIndex        =   8
            Top             =   -30
            Width           =   40000
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackColor       =   &H008080FF&
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "Fixedsys"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   0
            TabIndex        =   4
            Top             =   0
            Width           =   120
         End
         Begin VB.Label Label82 
            AutoSize        =   -1  'True
            BackColor       =   &H008080FF&
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "Fixedsys"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   105
            TabIndex        =   3
            Top             =   0
            Width           =   120
         End
      End
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "TEPRUE1"
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   3885
      TabIndex        =   9
      Top             =   105
      Visible         =   0   'False
      Width           =   840
   End
   Begin ComctlLib.ImageList imlToolbarPics 
      Left            =   5040
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483634
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   -2147483644
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   15
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "PREVIEW.frx":0008
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "PREVIEW.frx":0322
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "PREVIEW.frx":063C
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "PREVIEW.frx":0956
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "PREVIEW.frx":0C70
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "PREVIEW.frx":0F8A
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "PREVIEW.frx":12A4
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "PREVIEW.frx":15BE
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "PREVIEW.frx":18D8
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "PREVIEW.frx":1BF2
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "PREVIEW.frx":1F0C
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "PREVIEW.frx":20E6
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "PREVIEW.frx":21F8
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "PREVIEW.frx":2512
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "PREVIEW.frx":3888
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label TEPRUE 
      AutoSize        =   -1  'True
      Caption         =   "TEPRUE"
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   2940
      TabIndex        =   7
      Top             =   105
      Visible         =   0   'False
      Width           =   720
   End
End
Attribute VB_Name = "PREVIEW"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ANCHOMA, CAREMA&
Dim ARBUSQUE$

Private Sub BOTEXIT_Click()

End Sub

Private Sub Combo1_Change()
   If TRIM$(Combo1.TEXT) <> "" Then
     List4.Font = Combo1.TEXT
     List1.Font = Combo1.TEXT
     List4.Refresh
     DoEvents
     Call AJUSTANCH
   End If
End Sub

Private Sub Combo1_Click()
   List4.Font = Combo1.TEXT
   List1.Font = Combo1.TEXT
   List4.Refresh
   DoEvents
   Call AJUSTANCH
   Call GRACONTROL("FN-PREW", IDENTER$, Combo1.TEXT)
End Sub

Sub FULLSCREEN()
   Static ANCHNOR, ALTNOR, TOPNOR, LEFTNOR, ISMAXI%
   '
   If ANCHNOR < 1 Then
     ANCHNOR = Width
     ALTNOR = Height
     TOPNOR = Top
     LEFTNOR = Left
   End If
   '
   If ISMAXI% = 0 Then
       Top = 0
       Left = 0
       Width = Screen.Width
       Height = Screen.Height - 300
       Picture2.Width = Screen.Width - Picture2.Left - 450
       Picture2.Height = Height - Picture2.Top - 800
       List4.Height = Picture2.Height - 430
       VScroll2.Left = Picture2.Left + Picture2.Width
       VScroll2.Height = List4.Height
       HScroll1.Top = Picture2.Top + Picture2.Height
       HScroll1.Width = Picture2.Width
       Toolbar1.Left = VScroll2.Left + VScroll2.Width - Toolbar1.Width
       Frame1.Left = Toolbar1.Left - Frame1.Width - 120
       ISMAXI% = 1
       Call GRACONTROL("SCR-PREW", IDENTER$, "FULLSCREEN")
     Else
       Top = TOPNOR
       Left = LEFTNOR
       Width = ANCHNOR
       Height = ALTNOR
       Picture2.Width = Width - Picture2.Left - 450
       Picture2.Height = Height - Picture2.Top - 800
       List4.Height = Picture2.Height - 430
       VScroll2.Left = Picture2.Left + Picture2.Width
       VScroll2.Height = List4.Height
       HScroll1.Top = Picture2.Top + Picture2.Height
       HScroll1.Width = Picture2.Width
       Toolbar1.Left = VScroll2.Left + VScroll2.Width - Toolbar1.Width
       Frame1.Left = Toolbar1.Left - Frame1.Width - 120
       ISMAXI% = 0
       Call GRACONTROL("SCR-PREW", IDENTER$, "NORMAL")
   End If
   '
   Call AJUSTANCH
End Sub

Private Sub Command1_Click()
   List4.FontBold = Not (List4.FontBold)
   List1.FontBold = Not (List1.FontBold)
   List4.Refresh
   DoEvents
   Call AJUSTANCH
End Sub

Private Sub Form_Activate()
   MXX% = List4.ListCount - CAREMA&
   If MXX% < 0 Then MXX% = 0
   VScroll2.Value = 0
   VScroll2.Max = MXX%
   '
   On Error Resume Next
   Combo1.TEXT = List4.FontName
   On Error GoTo 0
   If List1.FontSize > 20 Then
      List1.FontSize = 20
   End If
   Text1 = TRIM$(Str$(List1.FontSize))
   VScroll1.Value = 21 - Val(Text1)
   Call AJUSTANCH
   '
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 114 Then
     If ARBUSQUE$ <> "" Then
       INI& = List4.ListIndex
10     For KU& = INI& + 1 To List4.ListCount - 1
         If InStr(UCase$(List4.List(KU&)), ARBUSQUE$) > 0 Then
           List4.ListIndex = KU&
           Exit Sub
         End If
       Next KU&
       If INI& > (-1) Then
         INI& = (-1)
         GoTo 10
       End If
       Call COMUNI("Argumento '" + ARBUSQUE$ + "' no Encontrado")
     End If
   End If
End Sub

Private Sub Form_Load()
   Call CENTRAFORM(Me)
   '
   Call PONELOGO
   '
   For I% = 1 To Screen.FontCount
     Label1.FontName = Screen.Fonts(I% - 1)
     A1 = 0: A2 = 0
     On Error Resume Next
     Label1 = String$(80, "l")
     A1 = Label1.Width
     Label1 = String$(80, "M")
     A2 = Label1.Width
     On Error GoTo 0
     If A1 > 0 And A2 > 0 Then
       If A1 = A2 Then
         Combo1.AddItem Label1.FontName
       End If
     End If
     On Error GoTo 0
   Next I%
   '
   FONA$ = CONTROL$("FN-PREW", IDENTER$)
   If FONA$ <> "" Then
      On Error Resume Next
      Combo1.TEXT = FONA$
      On Error GoTo 0
   End If
   '
   FOSI$ = CONTROL$("FS-PREW", IDENTER$)
   If FOSI$ <> "" Then
      On Error Resume Next
      Text1 = FOSI$
      On Error GoTo 0
   End If
   '
   FUSC$ = CONTROL$("SCR-PREW", IDENTER$)
   If FUSC$ = "FULLSCREEN" Then
      Call FULLSCREEN
   End If
   '
End Sub

Private Sub List1_Click()
   For U& = 1 To List1.ListCount:
     TTXX$ = List1.List(U& - 1)
     List1.RemoveItem (U& - 1)
     List1.AddItem TTXX$, U& - 1
   Next U&
End Sub

Private Sub List4_KeyDown(KeyCode As Integer, Shift As Integer)
'   If KeyCode = 114 Then
'     If ARBUSQUE$ <> "" Then
'       INI& = List4.ListIndex
'10     For KU& = INI& + 1 To List4.ListCount - 1
'         If InStr(UCase$(List4.List(KU&)), ARBUSQUE$) > 0 Then
'           List4.ListIndex = KU&
'           Exit Sub
'         End If
'       Next KU&
'       If INI& > (-1) Then
'         INI& = (-1)
'         GoTo 10
'       End If
'       Call COMUNI("Argumento '" + ARBUSQUE$ + "' no Encontrado")
'     End If
'   End If
End Sub

Private Sub Text1_Change()
   If Val(Text1) >= 1 Then
     If Val(Text1) < 100 Then
       List4.FontSize = Text1
       List1.FontSize = Text1
     End If
   End If
   List4.Refresh
   DoEvents
   Call AJUSTANCH
   Call GRACONTROL("FS-PREW", IDENTER$, Text1)
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As ComctlLib.Button)
   Select Case UCase$(Button.Key)
      Case "PRINT"
        Call WLITAB04.Image1_Click(0)
        Call WLITAB04.Command1_Click
      Case "SEARCH"
        Call Search
      Case "FULLSCREEN"
        Call FULLSCREEN
      Case "EXIT"
        Unload Me
   End Select
   '
End Sub

Private Sub VScroll1_Change()
  Text1 = 21 - TRIM$(VScroll1.Value)
End Sub

Private Sub VScroll2_Change()
   On Error Resume Next
   List4.TopIndex = VScroll2.Value
   On Error GoTo 0
End Sub

Private Sub VScroll2_Scroll()
   On Error Resume Next
   List4.TopIndex = VScroll2.Value
   On Error GoTo 0
End Sub

Private Sub HScroll1_Change()
  On Error Resume Next
  List4.Left = (-1) * 10 * HScroll1.Value - 40
  List1.Left = (-1) * 10 * HScroll1.Value - 40
  'Picture7.Left = (-1) * 10 * HScroll1.Value + 40
  On Error GoTo 0
End Sub

Private Sub HScroll1_Scroll()
  On Error Resume Next
  List4.Left = (-1) * 10 * HScroll1.Value - 40
  List1.Left = (-1) * 10 * HScroll1.Value - 40
  'Picture7.Left = (-1) * 10 * HScroll1.Value + 40
  On Error GoTo 0
End Sub

Sub AJUSTANCH()
   On Error Resume Next
   TEPRUE.Font = List4.Font
   TEPRUE.FontSize = List4.FontSize
   TEPRUE.FontBold = List4.FontBold
   '
   ANCHOMA = 0
   For KKI% = 1 To List4.ListCount
     TEPRUE = List4.List(KKI% - 1)
     If TEPRUE.Width > ANCHOMA Then
       ANCHOMA = TEPRUE.Width
     End If
   Next KKI%
   '
   HScroll1.Value = 0
   HScroll1.Max = 1
   If TEPRUE.Width > HScroll1.Width Then
     HScroll1.Max = (ANCHOMA + 100 - HScroll1.Width) / 10
   End If
   '
   CAREMA& = List4.Height / (TEPRUE.Height * 1)
   VScroll2.Max = VScroll2.Min
   If List4.ListCount > CAREMA& Then
     VScroll2.Max = List4.ListCount - CAREMA& + 1
   End If
   On Error GoTo 0
   '
End Sub

Sub Search()
   ABUS$ = InputBox$("Busqueda por Texto", , ARBUSQUE$)
   If ABUS$ = "" Then
     Exit Sub
   End If
   '
   ABUS$ = UCase$(ABUS$)
   If ABUS$ <> "" Then
     ARBUSQUE$ = ABUS$
     INI& = List4.ListIndex
10   For KU& = INI& + 1 To List4.ListCount - 1
       If InStr(UCase$(List4.List(KU&)), ABUS$) > 0 Then
         List4.ListIndex = KU&
         Exit Sub
       End If
     Next KU&
     If INI& > (-1) Then
       INI& = (-1)
       GoTo 10
     End If
     Call COMUNI("Argumento no Encontrado")
   End If
End Sub

Sub PONELOGO()
   '
   RUTARGRA$ = TRIM$(CONTROL$("", "RUTARGRA"))
   If RUTARGRA$ = "" Then RUTARGRA$ = TRIM$(LUDAT$)
   If Right$(RUTARGRA$, 1) <> "\" Then
      RUTARGRA$ = RUTARGRA$ + "\"
   End If
   '
   ISOLOGO$ = TRIM$(CONTROL$("", "ISOLOGO"))
   If ISOLOGO$ <> "" Then
     ARCHIMAG$ = RUTARGRA$ + ISOLOGO$
     If EXISTE%(ARCHIMAG$) = 1 Then
       '
       FORMALTER.Image3.Picture = LoadPicture("")
       FORMALTER.Image3.Picture = LoadPicture(ARCHIMAG$)
       '
12     HIMAGEN = FORMALTER.Image3.Picture.Width
       VIMAGEN = FORMALTER.Image3.Picture.Height
       '
       HCUADRO = Picture9.Width
       VCUADRO = Picture9.Height
       '
       REDUH = HIMAGEN / HCUADRO
       REDUV = VIMAGEN / VCUADRO
       REDUC = REDUH: If REDUV > REDUH Then REDUC = REDUV
       If REDUC < 1 Then REDUC = 1
       '
       With Picture9
         .Height = VCUADRO * REDUV / REDUC
         .Width = HCUADRO * REDUH / REDUC
         .Top = .Top + (VCUADRO - Picture9.Height) / 2
       End With
       '
       Image1.Width = Picture9.Width
       Image1.Height = Picture9.Height
       Image1.Stretch = True
       Picture9.Visible = True
       On Error Resume Next
       Image1.Picture = LoadPicture(ARCHIMAG$) ' FORMALTER.Image1.Picture ' , PH1 + BORIZQ, PV1 + BORSUP, ANCHCUA, ALTUCUA
       On Error GoTo 0
       '
       FORMALTER.Image3.Picture = LoadPicture("")
     End If
   End If
   '
End Sub

