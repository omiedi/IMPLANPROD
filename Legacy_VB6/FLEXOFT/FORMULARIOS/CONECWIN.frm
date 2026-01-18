VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form CONECWIN 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Conecta con WINTABU"
   ClientHeight    =   1695
   ClientLeft      =   45
   ClientTop       =   10305
   ClientWidth     =   4455
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1695
   ScaleWidth      =   4455
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin VB.PictureBox Picture999 
      Height          =   330
      Index           =   0
      Left            =   1470
      ScaleHeight     =   270
      ScaleWidth      =   1110
      TabIndex        =   3
      Top             =   525
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   4080
      Top             =   0
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   0
      Top             =   0
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   2625
      OleObjectBlob   =   "CONECWIN.frx":0000
      Top             =   0
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   225
      Index           =   0
      Left            =   0
      OleObjectBlob   =   "CONECWIN.frx":0234
      TabIndex        =   4
      Top             =   0
      Visible         =   0   'False
      Width           =   1485
   End
   Begin VB.Label Label3 
      Height          =   615
      Left            =   2400
      TabIndex        =   2
      Top             =   1050
      Width           =   1095
   End
   Begin VB.Label Label2 
      Height          =   495
      Left            =   720
      TabIndex        =   1
      Top             =   960
      Width           =   735
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
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
      Left            =   525
      TabIndex        =   0
      Top             =   315
      Width           =   3690
   End
End
Attribute VB_Name = "CONECWIN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Veces%

Private Sub Form_Load()

   'Call APLICACIONSKINS(Me)
    
   If Top < Screen.Height Then
     Top = Screen.Height
     DoEvents
   End If
End Sub

Private Sub Label3_Change()
    If Label3 = "1" Then
      Timer2.Enabled = False
      Label2 = ""
      Label3 = ""
      Hide
    End If
End Sub

Private Sub Timer1_Timer()
   '
   If EXISTE%(LUCOM$ + "RCONTROL.DRV") < 1 Then
      Label1 = ""
      Veces% = 0
      Timer1.Enabled = False
      Hide
      Exit Sub
   End If
   '
   Veces% = Veces% + 1
   If Veces% >= 30 Then
      Veces% = 0
      On Error GoTo 20
      'AppActivate "X-Reports", True
      'DoEvents
      On Error GoTo 0
      Veces% = 0
   End If
   Exit Sub
   '
20 Resume 21
21 On Error GoTo 0
   Veces% = 0
   Timer1.Enabled = False
   Kill LUCOM$ + "RCONTROL.DRV"
   Hide
   '
End Sub

Private Sub label1_Change()
   If Label1.Caption = "" Then
      Exit Sub
   End If
   '
   On Error Resume Next
   AppActivate "X-Reports", True
   If Err > 0 Then
     INTEN% = 0
10   Err = 0
     On Error GoTo 0
     Shell LUBAS$ + "WINTAB02.EXE", 4
     TINI = Timer: While Timer < TINI + 1: Wend
     On Error Resume Next
     AppActivate "X-Reports", True
     If Err > 0 Then
       INTEN% = INTEN% + 1
       If INTEN% <= 3 Then GoTo 10
     End If
   End If
   '
   If Err > 0 Then
     Err = 0
     On Error GoTo 0
     Call MENSERR(24, "Error no Recuperable:\Imposible Activar Generador de Reportes.")
     Call CIERRARCH("*.*")
     Call FINAL("")
     Exit Sub
   End If
   '
   On Error GoTo 0
   If TRIM(Dir(LUCOM$ + "*.*")) <> "" Then
      Call SAVEFILE(LUCOM$ + "RCONTROL.DRV", Label1.Caption)
   End If
   Veces% = 0
   Timer1.Enabled = True
   '
End Sub


Private Sub Timer2_Timer()
   '
   Static WTT As String * 128
   '
   ACONE$ = "WINTABU"
   If XVALO(Right$(App.EXEName, 2)) >= 4 Then ACONE$ = "WINTAB04"
   If XVALO(Right$(App.EXEName, 2)) >= 7 Then ACONE$ = "WINTAB07"
   On Error GoTo 20
   NK% = FreeFile
   Open LUCOM$ + ACONE$ + ".EXK" For Random Access Read Write As #NK% Len = 128
   Get #NK%, 1, WTT$
   Put #NK%, 1, WTT$
   Close #NK%
   On Error GoTo 0
   '
   If XVALO(Label2) > 0 Then
     Label3 = "1"
   End If
   Exit Sub
   '
20 Resume 21
21 On Error GoTo 0
   If NK% > 0 Then
     Close #NK%
     NK% = 0
   End If
   If XVALO(Label2) < 1 Then
     Label2 = "1"
   End If
   '
End Sub
