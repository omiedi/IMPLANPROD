VERSION 5.00
Begin VB.Form FORM7 
   Caption         =   "PRUEBA6"
   ClientHeight    =   3795
   ClientLeft      =   60
   ClientTop       =   315
   ClientWidth     =   3060
   LinkTopic       =   "Form1"
   ScaleHeight     =   3795
   ScaleWidth      =   3060
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Control MOVISTO"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   210
      TabIndex        =   2
      Top             =   3150
      Width           =   2430
   End
   Begin VB.ListBox List1 
      Height          =   1230
      Left            =   210
      TabIndex        =   1
      Top             =   1785
      Width           =   2430
   End
   Begin VB.FileListBox File1 
      Height          =   1260
      Left            =   210
      TabIndex        =   0
      Top             =   210
      Width           =   2535
   End
End
Attribute VB_Name = "FORM7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
   Command1.Enabled = False
   HOII% = HOY(HOS$)
5  CLAX$ = ""
6  PPX% = Int(Rnd * 63)
   If PPX% < 48 Or PPX% > 57 Then GoTo 6
   CLAX$ = CLAX$ + Chr$(PPX%)
   If Len(CLAX$) < 6 Then GoTo 6
   IDENPRO$ = CLAX$
   '
   IDENPRO$ = "AAAAA"
   KK% = 0
10 Randomize Timer
   CII% = Rnd * NUMAS%
   If CII% > 0 Then
     If CII% <= NUMAS% Then
       If Left$(CODEXT$(CII%), 2) = "02" Then
         Call BLOQARCH("MOVISTO")
         Randomize Timer
         CANTUL = 20 * Rnd - 10
         Call MOVISTO(HOII%, CII%, IDENPRO$, "VS", IDENPRO$, IDENPRO$, "PRUEBA BLOQUEOS", 0, "$", 0, 1, CANTUL)
         KK% = KK% + 1
         If KK% >= 3 Then
           Call CIERRARCH("MOVISTO")
           TINI = Timer: While Timer < TINI + 3: Wend
           KK% = 0
         End If
       End If
     End If
   End If
   If ULTREG&("MOVISTO") < 1000 Then
     GoTo 10
   End If
   Call CIERRARCH("*.*")
   Command1.Enabled = True
   MsgBox "FIN"
End Sub

Private Sub File1_Click()
   AAA = File1.filename
   PPXX% = InStr(AAA, ".")
   If PPXX% < 1 Then PPXX% = 1 + Len(AAA)
   ARCHI$ = Left$(AAA, PPXX% - 1)
   '
   Call BLOQARCH(ARCHI$)
   '
   For KKI% = 1 To List1.ListCount
     If List1.List(KKI% - 1) = AAA Then
       Exit Sub
     End If
   Next KKI%
   List1.AddItem AAA
   List1.Refresh
   '
End Sub

Private Sub Form_Load()
   '
   Dim BFF As String * 128
   Call VERJOBID
   '
   Show
   Refresh
   '
   File1.Path = "C:\CLIENTES\DOSIVAC\DATOS"
   File1.Pattern = "*.DAT"
   '
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call CIERRARCH("*.*")
   Call FINAL("")
End Sub

Private Sub List1_Click()
   AAA = List1.TEXT
   PPXX% = InStr(AAA, ".")
   If PPXX% < 1 Then PPXX% = 1 + Len(AAA)
   ARCHI$ = Left$(AAA, PPXX% - 1)
   Call LIBARCH(ARCHI$)
   List1.RemoveItem List1.ListIndex
   List1.Refresh
End Sub
