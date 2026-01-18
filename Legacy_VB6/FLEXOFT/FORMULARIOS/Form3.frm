VERSION 5.00
Begin VB.Form Form3 
   Caption         =   "PRUEBA CON FILEMANAGER"
   ClientHeight    =   3960
   ClientLeft      =   60
   ClientTop       =   315
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3960
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   0
      Top             =   420
   End
   Begin VB.ListBox List1 
      Height          =   2595
      Left            =   315
      TabIndex        =   3
      Top             =   105
      Width           =   4005
   End
   Begin VB.TextBox Text1 
      Height          =   330
      Left            =   315
      TabIndex        =   2
      Text            =   " "
      Top             =   3045
      Width           =   4005
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Cerrar"
      Height          =   330
      Left            =   3150
      TabIndex        =   1
      Top             =   3465
      Width           =   1275
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Agregar"
      Height          =   330
      Left            =   315
      TabIndex        =   0
      Top             =   3465
      Width           =   1380
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim FFX%, JJ&
Dim ABC As String * 16

Private Sub Form_Load()
LUDAT$ = "\FLEXOWIN\"
VERDEMO% = (-1)
'If FFX% < 1 Then
'   FFX% = FreeFile
'   Open "\FLEXOWIN\PRUEBA.DAT" For Random Access Read Write Shared As #FFX% Len = 16
'End If
End Sub

Private Sub Command2_Click()
TIN0 = Timer
10 TINI = Timer
   Call BLOQARCH("PRUEBA")
   JJ& = 1 + ULTREG&("PRUEBA")
   TEPRU$ = Left$(TRIM$(Text1.Text) + Str$(JJ&) + "      ", 8)
   Randomize Timer
   For KI% = 1 To 13
     TEPRU$ = TEPRU$ + Chr$(48 + 9 * Rnd)
   Next KI%
   Call REGAPP("PRUEBA", TEPRU$)
   Call CIERRARCH("PRUEBA")
   Call MUELISTA
   DEMORA = Rnd / 30
   While Timer < TINI + DEMORA: Wend
   If JJ& Mod 10 = 0 Then
      While Timer < TINI + 0.3: Wend
   End If
If Timer > TIN0 + 40 Then Exit Sub
   If JJ& < 9999 Then GoTo 10
End Sub

Private Sub Timer1_Timer()
   Call MUELISTA
End Sub

Sub MUELISTA()
Static LONGIAN, FI1&
If ULTREG&("PRUEBA") <> LONGIAN Then
10 LONGIAN = ULTREG&("PRUEBA")
   NARX% = FILENBR%("PRUEBA")
   FI2& = ULTREG&("PRUEBA")
   FI3& = FileLen("\FLEXOWIN\PRUEBA.DAT") / 16
   FI4& = LOF(NARX%) / 16
   If FI3& <> FI4& Then
      'MsgBox FI2& & "    " & FI3& & "   " & FI4&
      Call CIERRARCH("*.*")
      GoTo 10
   End If
   For U& = FI1& + 1 To FI2&
      'Get #FFX%, U&, ABC$
20    UA& = U&
      ABC$ = REGLEIDO$("PRUEBA", U&)
      If U& <> UA& Then U& = UA&: GoTo 20
      If U& = FI2& Then ULTIMO$ = ABC$
      List1.AddItem ABC$
   Next U&
   FI1& = FI2&
   TOPE = List1.ListCount - 10
   If TOPE >= 0 Then List1.TopIndex = TOPE
   List1.Refresh
   If FI2& > 1 Then
      'Get #FFX%, FI2& - 1, ABC$
      ABC$ = REGLEIDO$("PRUEBA", FI2& - 1)
      ANTEUL$ = ABC$
      If Val(Mid$(ULTIMO$, 2, 5)) <> Val(Mid$(ANTEUL$, 2, 5)) + 1 Then
         MsgBox "FALLO " + Text1.Text + "   U:" + ULTIMO$ + "   A:" + ANTEUL$
      End If
   End If
End If

End Sub

