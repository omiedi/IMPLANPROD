VERSION 5.00
Begin VB.Form FCORR_MO 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Libera Acceso a Mic_Outlk"
   ClientHeight    =   645
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   645
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   105
      Top             =   105
   End
   Begin VB.Label Label1 
      Height          =   435
      Left            =   3465
      TabIndex        =   0
      Top             =   210
      Width           =   750
   End
End
Attribute VB_Name = "FCORR_MO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
   If App.PrevInstance = True Then
      Close: End
   End If
End Sub

Private Sub Timer1_Timer()
   '
   Static SEGS%
   '
   SEGS% = SEGS% + 1
   ' Label1 = SEGS%
   If SEGS% Mod 120 = 0 Then
      Close
      Unload Me
      End
   End If
   '
   On Error Resume Next
   AppActivate "Microsoft Office Outlook", False
   If Err Then
      On Error GoTo 0
      Exit Sub
   End If
   '
   On Error GoTo 0
   SendKeys "%P", True
   SK$ = "{DOWN}"
   SendKeys SK$, True
   SendKeys SK$, True
   SendKeys SK$, True
   SendKeys SK$, True
   SendKeys Chr$(13), False
   Close
   End
   '
End Sub
