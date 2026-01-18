VERSION 5.00
Begin VB.Form FPRUWINZ 
   AutoRedraw      =   -1  'True
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List1 
      Height          =   645
      Left            =   240
      TabIndex        =   1
      Top             =   240
      Width           =   2295
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   3360
      TabIndex        =   0
      Top             =   240
      Width           =   1095
   End
End
Attribute VB_Name = "FPRUWINZ"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Command1.Enabled = False
     RETURNVALUE = Shell("C:\Archivos de programa\WinZip\WINZIP32.EXE", 4)
     AppActivate RETURNVALUE  ': AppActivate RETURNVALUE
     SendKeys "^N", True
     '
     AZIP$ = "C:\ANOVE\SABO\FLEX5.ZIP"
     On Error Resume Next
     Kill AZIP$
     On Error GoTo 0
     SendKeys AZIP$ + Chr$(13), True
     '
     DoEvents
     List1.AddItem "C:\FLEXOFT\WINDOWS\FLEXOFT.EXE"
     'List1.AddItem "C:\FLEXOFT\WINDOWS\FLFORMS.EXE"
     'List1.AddItem "C:\FLEXOFT\WINDOWS\FLEXPRFS.FLX"
     'List1.AddItem "C:\FLEXOFT\WINDOWS\FLEXRFS.ZIP"
     '
     TTXX$ = ""
     For U& = 1 To List1.ListCount
       TTXX$ = TTXX$ + Chr$(34) + ReplaSkey$(List1.List(U& - 1)) + Chr$(34) + Chr$(32)
     Next U&
     '
     SendKeys "%A", True
     SendKeys TRIM$(TTXX$) + Chr$(13) + Chr$(13), True
     SendKeys "%FX", True
Exit Sub
     '
     'TINI = Timer
5    DoEvents
     'If Timer < TINI + 0.5 Then GoTo 5
     'If EXISTE(AZIP$) < 1 Then
       'TINI = Timer
     '  GoTo 5
     'End If
     '
 GoTo 10
      'If Option3.Value = True Then
        SendKeys "+K", True
        SendKeys "%D" + "PEPE" + Chr$(13) ' Trim$(UCase$(Text3)) + Chr$(13), True
        TINI = Timer
6       DoEvents
        If Timer < TINI + 0.5 Then GoTo 6
        'If EXISTE%(AZIP$ + ".EXE") < 1 Then
        '  TINI = Timer
        '  GoTo 6
        'End If
        AppActivate "WinZip Self-Extractor"
        AppActivate "WinZip Self-Extractor"
        SendKeys "%N", True
        SendKeys "%C", True
      'End If
      '
10    AppActivate "WinZip - FLEX5" ' + AZIP$ ' + ".ZIP"
      AppActivate "WinZip - FLEX5" '+ AZIP$ ' + ".ZIP"
      'DoEvents
15 'On Error Resume Next
SendKeys "%FC", True '"%(F4))", True
'TINI = Timer: While Timer < TINI + 0.3: Wend
SendKeys "%F", True
    AppActivate "WinZip" ' + AZIP$ ' + ".ZIP"
      AppActivate "WinZip" '+ AZIP$ ' + ".ZIP"
SendKeys "%X", True
'TTYY$ = TTYY$ + "%(F4)"
'MsgBox TTYY$
'SendKeys TTYY$, True
'Exit Sub
      '
Visible = True
On Error GoTo 0
Exit Sub
      FIFUNA$ = AZIP$ + ".ZIP"
      'If Option3.Value = True Then
      '  FIFUNA$ = AZIP$ + ".EXE"
      'End If
      FISHNA$ = FIFUNA$
      While InStr(FISHNA$, "\") > 0
        FISHNA$ = Mid$(FISHNA$, 2)
      Wend
      If TRIM$(FISHNA$) <> "" Then
        'MAILMESS.List3.AddItem AJUSTI$(FISHNA$, 128) + FIFUNA$
      End If
      Command1.Enabled = True
      List1.Clear
      Hide
   'End If
   '

End Sub

Function ReplaSkey$(A$)
   '
   Static TCORCH$
   '
   If TCORCH$ = "" Then
     TCORCH$ = "+^%~()[]{}"
   End If
   '
   B$ = ""
   For UI% = 1 To Len(A$)
     C$ = Mid$(A$, UI%, 1)
     If InStr(TCORCH$, C$) > 0 Then
       C$ = Chr$(123) + C$ + Chr$(125) ' encierra entre llaves
     End If
     B$ = B$ + C$
   Next UI%
   '
   ReplaSkey$ = B$
   '
End Function

