VERSION 5.00
Begin VB.Form SECUDEMO 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Secuencias de Demostración"
   ClientHeight    =   8355
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   11910
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8355
   ScaleWidth      =   11910
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command6 
      Height          =   645
      Left            =   11235
      Picture         =   "SECUDEMO.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   29
      Top             =   945
      Width           =   540
   End
   Begin VB.CommandButton Command5 
      Height          =   645
      Left            =   11235
      Picture         =   "SECUDEMO.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   28
      Top             =   210
      Width           =   540
   End
   Begin VB.Frame Frame6 
      Caption         =   "Selección Imagen / Secuencia"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2115
      Left            =   6825
      TabIndex        =   23
      Top             =   105
      Width           =   4320
      Begin VB.Frame Frame5 
         Caption         =   "Secuencias"
         Height          =   1695
         Left            =   2205
         TabIndex        =   26
         Top             =   315
         Width           =   1905
         Begin VB.FileListBox File2 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1260
            Left            =   105
            Pattern         =   "*.SEC"
            TabIndex        =   27
            Top             =   315
            Width           =   1695
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Imágenes"
         Height          =   1695
         Left            =   210
         TabIndex        =   24
         Top             =   315
         Width           =   1905
         Begin VB.FileListBox File1 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1260
            Left            =   105
            Pattern         =   "*.JPG;*.BMP"
            TabIndex        =   25
            Top             =   315
            Width           =   1695
         End
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Instrucción Activa"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5800
      Left            =   5145
      TabIndex        =   8
      Top             =   2415
      Width           =   6630
      Begin VB.CommandButton Command4 
         Caption         =   "Ins"
         Height          =   285
         Left            =   5960
         TabIndex        =   22
         Top             =   735
         Width           =   540
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Add"
         Height          =   285
         Left            =   5985
         TabIndex        =   21
         Top             =   315
         Width           =   540
      End
      Begin VB.TextBox Text5 
         Height          =   285
         Left            =   5355
         MaxLength       =   8
         TabIndex        =   19
         Top             =   735
         Width           =   435
      End
      Begin VB.TextBox Text4 
         Height          =   285
         Left            =   4040
         MaxLength       =   8
         TabIndex        =   17
         Top             =   315
         Width           =   330
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   5355
         MaxLength       =   8
         TabIndex        =   15
         Top             =   315
         Width           =   435
      End
      Begin VB.TextBox INSTRUFI 
         Height          =   285
         Left            =   2100
         MaxLength       =   16
         TabIndex        =   13
         Top             =   735
         Width           =   2265
      End
      Begin VB.TextBox INSTRUTI 
         Height          =   285
         Left            =   105
         MaxLength       =   80
         TabIndex        =   12
         Top             =   1155
         Width           =   6390
      End
      Begin VB.TextBox INSTRUCO 
         Height          =   285
         Left            =   2100
         MaxLength       =   8
         TabIndex        =   10
         Top             =   315
         Width           =   1170
      End
      Begin VB.PictureBox Picture4 
         BackColor       =   &H00FF0000&
         Height          =   4070
         Left            =   100
         ScaleHeight     =   4005
         ScaleWidth      =   6345
         TabIndex        =   9
         Top             =   1575
         Width           =   6400
         Begin VB.Image Image1 
            Height          =   4500
            Left            =   0
            Stretch         =   -1  'True
            Top             =   0
            Width           =   6525
         End
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "Dem.Fin:"
         Height          =   225
         Left            =   4515
         TabIndex        =   20
         Top             =   840
         Width           =   750
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "Aper:"
         Height          =   225
         Left            =   3405
         TabIndex        =   18
         Top             =   420
         Width           =   540
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Dem.Ini:"
         Height          =   225
         Left            =   4515
         TabIndex        =   16
         Top             =   420
         Width           =   750
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Imagen / Secuencia:"
         Height          =   225
         Left            =   210
         TabIndex        =   14
         Top             =   840
         Width           =   1800
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Codigo:"
         Height          =   225
         Left            =   1260
         TabIndex        =   11
         Top             =   420
         Width           =   750
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Secuencia Activa"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5800
      Left            =   105
      TabIndex        =   6
      Top             =   2415
      Width           =   4950
      Begin VB.ListBox List1 
         BeginProperty Font 
            Name            =   "Courier"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5325
         Left            =   105
         TabIndex        =   7
         Top             =   315
         Width           =   4740
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Indice de Secuencias"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2115
      Left            =   105
      TabIndex        =   0
      Top             =   105
      Width           =   6630
      Begin VB.CommandButton Command7 
         Caption         =   "Kill"
         Height          =   360
         Left            =   5985
         TabIndex        =   30
         Top             =   1155
         Width           =   540
      End
      Begin VB.CommandButton Command2 
         Caption         =   "OK"
         Height          =   360
         Left            =   5985
         TabIndex        =   5
         Top             =   735
         Width           =   540
      End
      Begin VB.CommandButton Command1 
         Caption         =   "New"
         Height          =   360
         Left            =   5985
         TabIndex        =   4
         Top             =   310
         Width           =   540
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   1365
         TabIndex        =   3
         Top             =   1680
         Width           =   5160
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   105
         TabIndex        =   2
         Top             =   1680
         Width           =   1170
      End
      Begin VB.ListBox LISECUEN 
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1185
         ItemData        =   "SECUDEMO.frx":0454
         Left            =   105
         List            =   "SECUDEMO.frx":0456
         Sorted          =   -1  'True
         TabIndex        =   1
         Top             =   315
         Width           =   5790
      End
   End
End
Attribute VB_Name = "SECUDEMO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim EDITA%

Private Sub Command1_Click()
   Text1.Text = ""
   Text2.Text = ""
End Sub

Private Sub Command2_Click()
   If Trim$(Text1.Text) = "" Then Exit Sub
   TTXX$ = Space$(128)
   CODISECU$ = UCase$(Left$(Trim$(Text1.Text) + Space$(8), 8))
   Mid$(TTXX$, 1, 8) = CODISECU$
   Mid$(TTXX$, 11, 118) = Trim$(Text2.Text) + Space$(118)
   '
   LISECUEN.ListIndex = (-1)
10 For KKI% = 1 To LISECUEN.ListCount
     If Left$(LISECUEN.List(KKI% - 1), 8) = CODISECU$ Then
       LISECUEN.RemoveItem KKI% - 1
       GoTo 10
     End If
   Next KKI%
   '
   LISECUEN.AddItem TTXX$
   '
   N1% = FreeFile
   Dim BUFLE As String * 128
   Open DEMOPATH$ + "\SCRIPTS\SECUDEMO.DAT" For Random Access Read Write Shared As #N1% Len = 128
   For KKI% = 1 To LISECUEN.ListCount
     BUFLE$ = LISECUEN.List(KKI% - 1) + Space$(128)
     Put #N1%, KKI%, BUFLE$
   Next KKI%
   For KKI% = LISECUEN.ListCount + 1 To LOF(N1%) / 128
     BUFLE$ = Space$(128)
     Put #N1%, KKI%, BUFLE$
   Next KKI%
   Close #N1%
   '

End Sub

Private Sub Command3_Click()
   Command3.Enabled = False
   Call AGRELI("APP")
   Command3.Enabled = True
End Sub

Private Sub Command4_Click()
   Command4.Enabled = False
   Call AGRELI("INS")
   Command4.Enabled = True
End Sub

Private Sub Command5_Click()
   Hide
End Sub

Private Sub Command6_Click()
   Command6.Enabled = False
   Call GRALISECUEN
   EDITA% = 0
   Command6.Enabled = True
End Sub

Private Sub Command7_Click()
   '
   If LISECUEN.ListIndex >= 0 Then
      RESPU% = MsgBox("Realmente Desea Eliminar Esta Secuencia ?", 289)
      If RESPU% <> 1 Then Exit Sub
      '
      ILISTA& = LISECUEN.ListIndex
      LISECUEN.RemoveItem ILISTA&
      '
      Text1.Text = ""
      Text2.Text = ""
      '
      N1% = FreeFile
      Dim BUFLE As String * 128
      Open DEMOPATH$ + "\SCRIPTS\SECUDEMO.DAT" For Random Access Read Write Shared As #N1% Len = 128
      For KKI% = 1 To LISECUEN.ListCount
        BUFLE$ = LISECUEN.List(KKI% - 1)
        Put #N1%, KKI%, BUFLE$
      Next KKI%
      For KKI% = LISECUEN.ListCount + 1 To LOF(N1%) / 128
        BUFLE$ = Space$(128)
        Put #N1%, KKI%, BUFLE$
      Next KKI%
      Close #N1%
      '
   End If
End Sub

Private Sub File1_Click()
   '
   Image1.Stretch = False
   ADEMO$ = UCase$(File1.filename)
   Image1.Picture = LoadPicture("")
   Call MUESTRAIMAGEN(ADEMO$)
   '
End Sub

Private Sub File2_Click()
   ADEMO$ = UCase$(File2.filename)
   INSTRUFI.Text = ADEMO$
   Text3.Text = ""
   Text4.Text = ""
   Text5.Text = ""
End Sub

Private Sub Form_Load()
   '
   Call CENTRAFORM(Me)
   DoEvents
   '
   Screen.MousePointer = 11
   N1% = FreeFile
   Dim BUFLE As String * 128
   Open DEMOPATH$ + "\SCRIPTS\SECUDEMO.DAT" For Random Access Read Write Shared As #N1% Len = 128
   For KKI% = 1 To LOF(N1%) / 128
     Get #N1%, KKI%, BUFLE$
     If Trim$(BUFLE$) <> "" Then
       LISECUEN.AddItem BUFLE$
     End If
   Next KKI%
   Close #N1%
   '
   File1.Path = DEMOPATH$ + "\SCREENS"
   File2.Path = DEMOPATH$ + "\SCRIPTS"
   Screen.MousePointer = 1
   '
End Sub

Private Sub Form_Unload(Cancel As Integer)
   EDITA% = 0
End Sub

Private Sub INSTRUFI_GotFocus()
   INSTRUFI.SelStart = 0
   INSTRUFI.SelLength = Len(INSTRUFI.Text)
End Sub

Sub LISECUEN_Click()
   '
   ILISTA& = LISECUEN.ListIndex
   If ILISTA& >= 0 Then
     EDITA% = 0
     TTXX$ = LISECUEN.List(ILISTA&)
     Text1.Text = Trim$(Left$(TTXX$, 8))
     Text2.Text = Trim$(Mid$(TTXX$, 11))
     '
     Call CARLISECUEN
   End If
   '
End Sub

Sub CARLISECUEN()
   '
   SECUEN$ = Trim$(UCase$(Text1.Text))
   If SECUEN$ = "" Then Exit Sub
   ASECU$ = SECUEN$ + ".SEC"
   '
   List1.Clear
   '
   N1% = FreeFile
   Dim BUFLE1 As String * 256
   Open DEMOPATH$ + "\SCRIPTS\" + ASECU$ For Random Access Read Write As #N1% Len = 256
   '
   If LOF(N1%) < 256 Then
      '
      N2% = FreeFile
      Dim BUFLE2 As String * 256
      Open DEMOPATH$ + "\SCRIPTS\DEMOIMG.DAT" For Random Access Read As #N2% Len = 256
      JJ% = 0
      '
      For II% = 1 To LOF(N2%) / 256
        Get #N2%, II%, BUFLE2$
        SECULE$ = Trim$(UCase$(Left$(BUFLE2$, 8)))
        If SECULE$ = SECUEN$ Then
          ARIMA$ = Trim$(Mid$(BUFLE2$, 9, 16))
          PPXX% = InStr(ARIMA$, ".")
          If PPXX% < 1 Then PPXX% = 1 + Len(ARIMA$)
          ARIMA$ = Left$(ARIMA$, PPXX% - 1)
          ARIMA$ = Left$(ARIMA$ + Space$(8), 8)
          On Error Resume Next
          PAUSAI% = CVS(Mid$(BUFLE2$, 105, 4))
          PAUSAF% = CVS(Mid$(BUFLE2$, 109, 4))
          MABRE% = 0
          
          If Trim$(ARIMA$) <> "" Then
            TTXX$ = BUFLE2$
            Mid$(TTXX$, 1, 8) = Left$(ARIMA$ + Space$(8), 8)
            Mid$(TTXX$, 105, 152) = String$(152, 0)
            Mid$(TTXX$, 105, 2) = MKI$(PAUSAI%)
            Mid$(TTXX$, 107, 2) = MKI$(PAUSAF%)
            Mid$(TTXX$, 109, 2) = MKI$(MABRE%)
            JJ% = JJ% + 1
            BUFLE1$ = TTXX$
            Put #N1%, JJ%, BUFLE1$
          End If
        End If
      Next II%
      '
      Close #N2%
      '
   End If
   '
   For II% = 1 To LOF(N1%) / 256
      Get #N1%, II%, BUFLE1$
      INSTRU$ = Trim$(UCase$(Left$(BUFLE1$, 8)))
      ARIMA$ = Trim$(UCase$(Mid$(BUFLE1$, 9, 16)))
      LEYEN$ = Trim$(Mid$(BUFLE1$, 25, 80))
      PAUSI% = CVI(Mid$(BUFLE1$, 105, 2))
      PAUSF% = CVI(Mid$(BUFLE1$, 107, 2))
      MABRE% = CVI(Mid$(BUFLE1$, 109, 2))
      '
      If Trim$(INSTRU$) <> "" Then
        TTXX$ = Space$(256)
        Mid$(TTXX$, 1, 8) = INSTRU$
        Mid$(TTXX$, 11, 80) = LEYEN$
        Mid$(TTXX$, 91, 16) = ARIMA$
        Mid$(TTXX$, 107, 4) = Trim$(Str$(PAUSI%))
        Mid$(TTXX$, 111, 4) = Trim$(Str$(PAUSF%))
        Mid$(TTXX$, 115, 4) = Trim$(Str$(MABRE%))
        List1.AddItem TTXX$
      End If
   Next II%
   Close #N1%
   '
End Sub
'
Sub GRALISECUEN()
   '
   SECUEN$ = Trim$(UCase$(Text1.Text))
   If SECUEN$ = "" Then Exit Sub
   ASECU$ = SECUEN$ + ".SEC"
   '
   N1% = FreeFile
   Dim BUFLE1 As String * 256
   Open DEMOPATH + "\SCRIPTS\" + ASECU$ For Random Access Read Write As #N1% Len = 256
   JJ% = 0
   '
   For UI% = 1 To List1.ListCount
      TTYY$ = List1.List(UI% - 1)
      INSTRU$ = UCase$(Trim$(Mid$(TTYY$, 1, 8)))
      TITU$ = Trim$(Mid$(TTYY$, 11, 80))
      ARIMA$ = Trim$(Mid$(TTYY$, 91, 16))
      PAUSAI% = Val(Trim$(Mid$(TTYY$, 107, 4)))
      PAUSAF% = Val(Trim$(Mid$(TTYY$, 111, 4)))
      MABRE% = Val(Trim$(Mid$(TTYY$, 115, 4)))
      '
      TTXX$ = String$(256, 0)
      Mid$(TTXX$, 1, 8) = Left$(INSTRU$ + Space$(8), 8)
      Mid$(TTXX$, 9, 16) = Left$(ARIMA$ + Space$(16), 16)
      Mid$(TTXX$, 25, 80) = Left$(TITU$ + Space$(80), 80)
      Mid$(TTXX$, 105, 2) = MKI$(PAUSAI%)
      Mid$(TTXX$, 107, 2) = MKI$(PAUSAF%)
      Mid$(TTXX$, 109, 2) = MKI$(MABRE%)
      JJ% = JJ% + 1
      BUFLE1$ = TTXX$
      Put #N1%, JJ%, BUFLE1$
   Next UI%
   '
   For UI% = List1.ListCount + 1 To LOF(N1%) / 256
      BUFLE1$ = String$(256, 0)
      Put #N1%, UI%, BUFLE1$
   Next UI%
   '
   Close #N1%
   '
End Sub

Private Sub List1_Click()
   '
   If EDITA% = 0 Then
      TTXX$ = List1.Text
      INSTRUCO.Text = Trim$(Mid$(TTXX$, 1, 8))
      INSTRUTI.Text = Trim$(Mid$(TTXX$, 11, 80))
      INSTRUFI.Text = Trim$(Mid$(TTXX$, 91, 16))
      Text3.Text = Trim$(Mid$(TTXX$, 107, 4))
      Text5.Text = Trim$(Mid$(TTXX$, 111, 4))
      Text4.Text = Trim$(Mid$(TTXX$, 115, 4))
      '
      Image1.Stretch = False
      ADEMO$ = INSTRUFI.Text
      Image1.Picture = LoadPicture("")
      '
      Call MUESTRAIMAGEN(ADEMO$)
   End If
   '
99 On Error GoTo 0
   '
End Sub

Private Sub INSTRUCO_GotFocus()
   INSTRUCO.SelStart = 0
   INSTRUCO.SelLength = Len(INSTRUCO.Text)
End Sub

Sub MUESTRAIMAGEN(ADEMO$)
   '
   INSTRUFI.Text = ADEMO$
   '
   On Error GoTo 99
   Image1.Picture = LoadPicture(DEMOPATH$ + "\SCREENS\" + ADEMO$)
   On Error GoTo 0
   '
   REDUV = Picture4.Height / Image1.Height
   REDUH = Picture4.Width / Image1.Width
   REDUC = REDUH: If REDUV < REDUH Then REDUC = REDUV
   Image1.Stretch = True
   Image1.Width = Image1.Width * REDUC
   Image1.Height = Image1.Height * REDUC
   Image1.Top = (Picture4.Height - Image1.Height) / 2
   Image1.Left = (Picture4.Width - Image1.Width) / 2
   '
99 On Error GoTo 0
   '
End Sub

Sub AGRELI(MODOX$)
   '
   INSTRU$ = Trim$(UCase$(INSTRUCO.Text))
   If Len(INSTRU$) < 1 Or Len(INSTRU$) > 8 Then Exit Sub
   '
   For UI% = 1 To List1.ListCount
     If UCase$(Trim$(Left$(List1.List(UI% - 1), 8))) = INSTRU$ Then
       Exit Sub
     End If
   Next UI%
   '
   ARIMA$ = Trim$(UCase$(INSTRUFI.Text))
   If Len(ARIMA$) < 1 Or Len(ARIMA$) > 16 Then Exit Sub
   '
   LEYEN$ = Trim$(INSTRUTI.Text)
   If Len(LEYEN$) < 1 Or Len(LEYEN$) > 80 Then Exit Sub
   '
   PAUSI% = Val(Text3.Text)
   If PAUSI% < 0 Or PAUSI% > 15 Then PAUSI% = 3
   '
   PAUSF% = Val(Text5.Text)
   If PAUSF% < 0 Or PAUSF% > 15 Then PAUSF% = 3
   '
   MABRE% = Val(Text4.Text)
   If MABRE% < 0 Or MABRE% > 16 Then MABRE% = 0
   '
   TTXX$ = Space$(256)
   Mid$(TTXX$, 1, 8) = INSTRU$
   Mid$(TTXX$, 11, 80) = LEYEN$
   Mid$(TTXX$, 91, 16) = ARIMA$
   Mid$(TTXX$, 107, 4) = Trim$(Str$(PAUSI%))
   Mid$(TTXX$, 111, 4) = Trim$(Str$(PAUSF%))
   Mid$(TTXX$, 115, 4) = Trim$(Str$(MABRE%))
   '
   INSTRUCO.Text = "": EDITA% = 0
   If MODOX$ = "INS" Then
     If List1.ListIndex >= 0 Then
       ILISTA& = List1.ListIndex
       List1.AddItem TTXX$, ILISTA&
       List1.ListIndex = ILISTA&
       Exit Sub
     End If
   End If
   List1.AddItem TTXX$
   List1.ListIndex = List1.ListCount - 1
   '
End Sub

Private Sub List1_DblClick()
   '
   If EDITA% = 1 Then
      TTXX$ = List1.Text
      INSTRUCO.Text = Trim$(Mid$(TTXX$, 1, 8))
      INSTRUTI.Text = Trim$(Mid$(TTXX$, 11, 80))
      INSTRUFI.Text = Trim$(Mid$(TTXX$, 91, 16))
      Text3.Text = Trim$(Mid$(TTXX$, 107, 4))
      Text5.Text = Trim$(Mid$(TTXX$, 111, 4))
      Text4.Text = Trim$(Mid$(TTXX$, 115, 4))
      '
      Image1.Stretch = False
      ADEMO$ = INSTRUFI.Text
      Image1.Picture = LoadPicture("")
      '
      Call MUESTRAIMAGEN(ADEMO$)
   End If
   '
   EDITA% = 1
   ILISTA& = List1.ListIndex
   List1.RemoveItem ILISTA&
   If ILISTA& < List1.ListCount Then
     List1.ListIndex = ILISTA&
   End If
   '
End Sub

