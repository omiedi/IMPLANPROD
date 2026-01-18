VERSION 5.00
Begin VB.Form FORM6 
   Caption         =   "PRUEBA6"
   ClientHeight    =   3225
   ClientLeft      =   60
   ClientTop       =   315
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3225
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.FileListBox File1 
      Height          =   1260
      Left            =   210
      TabIndex        =   0
      Top             =   210
      Visible         =   0   'False
      Width           =   1065
   End
End
Attribute VB_Name = "FORM6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
   '
   Dim BFF As String * 128
   Call VERJOBID
   '
   Show
   Refresh
   '
   File1.Path = "C:\FLEXOWIN"
   File1.Pattern = "*.PRF"
   '
   Screen.MousePointer = 11
   JJ& = 0
   For KKI% = 1 To File1.ListCount
     ANAM$ = File1.List(KKI% - 1)
     PPXX% = InStr(ANAM$, "."): If PPXX% < 1 Then PPXX% = 1 + Len(ANAM$)
     ANAM$ = Left$(ANAM$, PPXX% - 1)
     ANAME$ = "C:\FLEXOWIN\" + File1.List(KKI% - 1)
     NX% = FreeFile
     Open ANAME$ For Random Access Read As #NX% Len = 128
     Get #NX%, 1, BFF$
     If CVI(Mid$(BFF$, 127, 2)) = 6 Then
       TTXX$ = Space$(64)
       TITU$ = Left$(BFF$, 47)
       Call REPLA(TTXX$, ANAM$, 1, 8)
       Call REPLA(TTXX$, TITU$, 9, 56)
       JJ& = JJ& + 1
       Call GRAREG("DIRFS", TTXX$, JJ&)
       Close #NX%: NX% = 0
       If InStr(UCase$(CurDir), "FLEXOWIN") < 1 Then
         FileCopy ANAME$, CurDir + "\" + ANAM$ + ".PRF"
       End If
     End If
     If NX% > 0 Then Close #NX%: NX% = 0
   Next KKI%
   Call SETULTREG("DIRFS", JJ&)
   Call CIERRARCH("DIRFS")
   Screen.MousePointer = 1
   '
10 Call SELECHO("DIRFS")
   ARCHU$ = VALACT1$("DIRFS")
   If ARCHU$ <> "" Then
     VTB% = VENTABU%(ARCHU$)
     GoTo 10
   End If
   '
   Call CIERRARCH("*.*")
   Call FINAL("")
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call CIERRARCH("*.*")
   Call FINAL("")
End Sub
