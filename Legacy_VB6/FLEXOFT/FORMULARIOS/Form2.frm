VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "PRUEBA CON CLOSE"
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
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim FFX%, JJ&
Dim ABC As String * 16

Private Sub Form_Load()
If FFX% < 1 Then
   FFX% = FreeFile
   Open "\FLEXOWIN\PRUEBA.DAT" For Random Access Read Write Shared As #FFX% Len = 16
End If
End Sub

Private Sub Command2_Click()
TIN0 = Timer
10 TINI = Timer
   If FFX% < 1 Then
      FFX% = FreeFile
      Open "\FLEXOWIN\PRUEBA.DAT" For Random Access Read Write Shared As #FFX% Len = 16
   End If
   '
   JJ& = 1 + LOF(FFX%) / 16 'FileLen("\FLEXOWIN\PRUEBA.DAT") / 16
   TEPRU$ = Left$(Trim$(Text1.Text) + Str$(JJ&) + "      ", 8)
   Randomize Timer
   For KI% = 1 To 13
     TEPRU$ = TEPRU$ + Chr$(48 + 9 * Rnd)
   Next KI%
   ABC$ = TEPRU$
   Put #FFX%, JJ&, ABC$
   Call MUELISTA
   Close #FFX%: FFX% = 0
   DEMORA = Rnd / 100
   While Timer < TINI + DEMORA: Wend
If Timer > TIN0 + 10 Then Exit Sub
   If JJ& < 9999 Then GoTo 10
End Sub

Private Sub Command3_Click()
Close #FFX%
   FFX% = FreeFile
   Open "\FLEXOWIN\PRUEBA.DAT" For Random Access Read Write Shared As #FFX% Len = 16
End Sub


Private Sub Timer1_Timer()
   If FFX% < 1 Then
      FFX% = FreeFile
      Open "\FLEXOWIN\PRUEBA.DAT" For Random Access Read Write Shared As #FFX% Len = 16
   End If
   '
   Call MUELISTA
End Sub

Sub MUELISTA()
Static LONGIAN, FI1&
If LOF(FFX%) <> LONGIAN Then
   LONGIAN = LOF(FFX%)
   FI2& = LOF(FFX%) / 16
   For U& = FI1& + 1 To FI2&
      Get #FFX%, U&, ABC$
      If U& = FI2& Then ULTIMO$ = ABC$
      List1.AddItem ABC$
   Next U&
   FI1& = FI2&
   TOPE = List1.ListCount - 10
   If TOPE >= 0 Then List1.TopIndex = TOPE
   List1.Refresh
   If FI2& > 1 Then
      Get #FFX%, FI2& - 1, ABC$
      ANTEUL$ = ABC$
      If Val(Mid$(ULTIMO$, 2, 5)) <> Val(Mid$(ANTEUL$, 2, 5)) + 1 Then
         MsgBox "FALLO " + Text1.Text + "   " + ULTIMO$ + "   " + ANTEUL$
      End If
   End If
End If

End Sub

