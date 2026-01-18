VERSION 5.00
Begin VB.Form COMPAFU 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Comparacion de Fuentes"
   ClientHeight    =   11625
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   14310
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   11625
   ScaleWidth      =   14310
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command4 
      Caption         =   "Next Dif."
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
      Left            =   11340
      TabIndex        =   19
      Top             =   105
      Width           =   645
   End
   Begin VB.CommandButton Command3 
      Caption         =   "End"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   13335
      TabIndex        =   16
      Top             =   670
      Width           =   690
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   960
      Left            =   7035
      MultiLine       =   -1  'True
      TabIndex        =   15
      Text            =   "COMPAFU.frx":0000
      Top             =   9240
      Width           =   6945
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   960
      Left            =   0
      MultiLine       =   -1  'True
      TabIndex        =   14
      Text            =   "COMPAFU.frx":0006
      Top             =   9240
      Width           =   6945
   End
   Begin VB.CommandButton Command2 
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   13335
      TabIndex        =   11
      Top             =   105
      Width           =   690
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Actualiza REMOTO"
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
      Left            =   12075
      TabIndex        =   10
      Top             =   105
      Width           =   1170
   End
   Begin VB.ListBox List3 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   840
      Left            =   4410
      TabIndex        =   9
      Top             =   105
      Width           =   6840
   End
   Begin VB.ListBox List2 
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7620
      Left            =   7035
      MultiSelect     =   2  'Extended
      TabIndex        =   1
      Top             =   1470
      Width           =   7000
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7620
      Left            =   0
      MultiSelect     =   2  'Extended
      TabIndex        =   0
      Top             =   1470
      Width           =   7000
   End
   Begin VB.Label Label11 
      BorderStyle     =   1  'Fixed Single
      Height          =   540
      Left            =   7035
      TabIndex        =   18
      Top             =   10395
      Width           =   6945
   End
   Begin VB.Label Label10 
      BorderStyle     =   1  'Fixed Single
      Height          =   540
      Left            =   0
      TabIndex        =   17
      Top             =   10395
      Width           =   6945
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   14280
      Y1              =   1050
      Y2              =   1050
   End
   Begin VB.Label Label9 
      Caption         =   "NEW"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   330
      Left            =   12600
      TabIndex        =   13
      Top             =   1050
      Visible         =   0   'False
      Width           =   645
   End
   Begin VB.Label Label8 
      Caption         =   "NEW"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   330
      Left            =   5880
      TabIndex        =   12
      Top             =   1050
      Visible         =   0   'False
      Width           =   645
   End
   Begin VB.Label Label7 
      Caption         =   "Rutina:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   3675
      TabIndex        =   8
      Top             =   105
      Width           =   1065
   End
   Begin VB.Label Label6 
      Caption         =   "Label6"
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
      Left            =   735
      TabIndex        =   7
      Top             =   420
      Width           =   5370
   End
   Begin VB.Label Label5 
      Caption         =   "Archivo:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   105
      TabIndex        =   6
      Top             =   105
      Width           =   1065
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Local"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   330
      Left            =   7245
      TabIndex        =   5
      Top             =   1050
      Width           =   5055
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Local"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   330
      Left            =   1050
      TabIndex        =   4
      Top             =   1050
      Width           =   5055
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   345
      Left            =   7035
      TabIndex        =   3
      Top             =   1050
      Width           =   75
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Local:"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   330
      Left            =   105
      TabIndex        =   2
      Top             =   1050
      Width           =   5055
   End
End
Attribute VB_Name = "COMPAFU"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
   ESTEAR$ = Label10
   ACOMPLE$ = Label11
   If Command1.Caption = "Actualiza LOCAL" Then
       FileCopy ACOMPLE$, ESTEAR$
       DoEvents
     ElseIf Command1.Caption = "Actualiza REMOTO" Then
       FileCopy ESTEAR$, ACOMPLE$
       DoEvents
   End If
   Hide
End Sub

Private Sub Command2_Click()
   Hide
End Sub

Private Sub Command3_Click()
   Call CIERRARCH("*.*")
   Close: End
End Sub

Private Sub Command4_Click()
   ini1% = List1.ListIndex
   If ini1% < 0 Then Exit Sub
   If List1.Selected(ini1%) = False Then Exit Sub
   '
   INI2% = List2.ListIndex
   If INI2% < 0 Then Exit Sub
   If List2.Selected(INI2%) = False Then Exit Sub
   '
   If UCase$(TRIM$(List1.TEXT)) <> UCase$(TRIM$(List2.TEXT)) Then
      Call MENSERR(24, "Debe Iniciar en Registros Coincidentes.")
      Exit Sub
   End If
   '
   For KI1% = ini1% To List1.ListCount - 1
     List1.Selected(KI1%) = False
   Next KI1%
   '
   For KI2% = INI2% To List2.ListCount - 1
     List2.Selected(KI2%) = False
   Next KI2%
   '
   JJ% = 0
5  JJ% = JJ% + 1
   If UCase$(LTrim(List1.List(ini1% + JJ%))) <> UCase$(LTrim(List2.List(INI2% + JJ%))) Then
      On Error Resume Next
      List1.Selected(ini1% + JJ%) = True
      On Error GoTo 0
      If List1.TopIndex < ini1% + JJ% - 20 Then List1.TopIndex = ini1% + JJ% - 20
      Text1 = List1.TEXT
      On Error Resume Next
      List2.Selected(INI2% + JJ%) = True
      On Error GoTo 0
      If List2.TopIndex < INI2% + JJ% - 20 Then List2.TopIndex = INI2% + JJ% - 20
      Text2 = List2.TEXT
      Exit Sub
   End If
   If InStr(UCase$(LTrim(List1.List(ini1% + JJ%))), "END SUB") > 0 Then GoTo 80
   If InStr(UCase$(LTrim(List1.List(ini1% + JJ%))), "END FUNCTION") > 0 Then GoTo 80
   If ini1% + JJ% >= List1.ListCount Then
     If INI2% + JJ% >= List2.ListCount Then
       GoTo 80
     End If
   End If
   GoTo 5
   '
80 MsgBox "No Más Diferencias"
   '
End Sub

Private Sub Label1_DblClick()
   If COMALTER%("Realmente Desea RESTAURAR la Versión Local ?\\No, Continuar\Si, Restaurar") = 2 Then
      ESTEAR$ = Label10
      ACOMPLE$ = Label11
      FileCopy ESTEAR$, ACOMPLE$
      DoEvents
      Hide
   End If
End Sub

Private Sub List1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
       Screen.MousePointer = 11
       TTXX$ = ""
       For U& = 1 To List1.ListCount
          If List1.Selected(U& - 1) = True Then
             TTXX$ = TTXX$ + List1.List(U& - 1) + vbCrLf
          End If
       Next U&
       Clipboard.Clear
       Clipboard.SetText TTXX$
       Screen.MousePointer = 1
    End If
End Sub

Private Sub List2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
       Screen.MousePointer = 11
       TTXX$ = ""
       For U& = 1 To List2.ListCount
          If List2.Selected(U& - 1) = True Then
             TTXX$ = TTXX$ + List2.List(U& - 1) + vbCrLf
          End If
       Next U&
       Clipboard.Clear
       Clipboard.SetText TTXX$
       Screen.MousePointer = 1
    End If
End Sub

Private Sub List3_DblClick()
   For KI1% = 1 To List1.ListCount
     List1.Selected(KI1% - 1) = False
   Next KI1%
   '
   For KI2% = 1 To List2.ListCount
     List2.Selected(KI2% - 1) = False
   Next KI2%
   '
   Text1 = "": Text2 = ""
   TBUS$ = List3.TEXT
   ini1% = 0: INI2% = 0
   For KI1% = 1 To List1.ListCount
     If LTrim$(List1.List(KI1% - 1)) = TBUS$ Then
       ini1% = KI1%
       GoTo 10
     End If
   Next KI1%
   '
10 For KI2% = 1 To List2.ListCount
     If LTrim$(List2.List(KI2% - 1)) = TBUS$ Then
       INI2% = KI2%
       GoTo 20
     End If
   Next KI2%
   '
20 KI1% = ini1%
   KI2% = INI2%
   '
   If ini1% > 0 And INI2% <= 0 Then
     List1.Selected(ini1% - 1) = True
     Text1 = List1.List(ini1% - 1)
     If List1.TopIndex < ini1% - 20 Then List1.TopIndex = ini1% - 20
     Exit Sub
   End If
   '
   If ini1% <= 0 And INI2% > 0 Then
     List2.Selected(INI2% - 1) = True
     If List2.TopIndex < INI2% - 20 Then List2.TopIndex = INI2% - 20
     Text2 = List2.List(INI2% - 1)
     Exit Sub
   End If
   '
25 If TRIM$(UCase$(List1.List(KI1% - 1))) <> TRIM$(UCase$(List2.List(KI2% - 1))) Then
     List1.Selected(KI1% - 1) = True
     If List1.TopIndex < KI1% - 20 Then List1.TopIndex = KI1% - 20
     Text1 = List1.List(KI1% - 1)
     List2.Selected(KI2% - 1) = True
     If List2.TopIndex < KI2% - 20 Then List2.TopIndex = KI2% - 20
     Text2 = List2.List(KI2% - 1)
     Exit Sub
   End If
   KI1% = KI1% + 1
   KI2% = KI2% + 1
   If KI1% <= List1.ListCount Then
     If KI2% <= List2.ListCount Then
       GoTo 25
     End If
   End If
   '
   List1.Selected(ini1% - 1) = True
   List1.TopIndex = ini1% - 1
   Text1 = List1.List(ini1% - 1)
   List2.Selected(INI2% - 1) = True
   List2.TopIndex = INI2% - 1
   Text2 = List2.List(INI2% - 1)
   MsgBox "Diferencia no Encontada"
   '
End Sub
