VERSION 5.00
Begin VB.Form ACTDRIVERS 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Genera Archivos de Extracción de Controladores"
   ClientHeight    =   6135
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   6930
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6135
   ScaleWidth      =   6930
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox List3 
      Height          =   1230
      Left            =   5565
      TabIndex        =   6
      Top             =   870
      Width           =   1170
   End
   Begin VB.ListBox List2 
      Height          =   2790
      Left            =   210
      TabIndex        =   4
      Top             =   3150
      Width           =   6525
   End
   Begin VB.CommandButton Command1 
      Height          =   435
      Left            =   5565
      Picture         =   "ACTDRIVERS.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Inicia Exploraciónde la Red"
      Top             =   315
      Width           =   1170
   End
   Begin VB.ListBox List1 
      Height          =   645
      Left            =   210
      TabIndex        =   2
      Top             =   2310
      Width           =   5160
   End
   Begin VB.FileListBox File1 
      Height          =   1845
      Left            =   2940
      TabIndex        =   1
      Top             =   315
      Width           =   2430
   End
   Begin VB.DirListBox Dir1 
      Height          =   1890
      Left            =   210
      TabIndex        =   0
      Top             =   315
      Width           =   2535
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   5565
      TabIndex        =   7
      Top             =   2310
      Width           =   1170
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   5565
      TabIndex        =   5
      Top             =   2670
      Width           =   1170
   End
End
Attribute VB_Name = "ACTDRIVERS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
   '
   Dim NOMARCH$(4096), FEARCH#(4096), RUTARCH$(4096)
   '
   DIEXPLO$ = ULODATOS$ + "\"
   CARCHI% = 0
   '
   DIRFSX$ = ULODATOS$ + "\RFSX"
   On Error Resume Next
   MkDir DIRFSX$
   On Error GoTo 0
   '
   Dir1.Path = DIRFSX$
   File1.Path = Dir1.Path
   File1.Pattern = "*.RFS"
   '
   For U& = 1 To File1.ListCount
     FINAX$ = UCase$(File1.List(U& - 1))
     If Asc(Left$(FINAX$, 1)) >= 65 Then
       If Asc(Left$(FINAX$, 1)) <= 90 Then
         LOARX = 0: FIDAT# = 0
         On Error Resume Next
         LOARX = FileLen(DIRFSX$ + "\" + FINAX$)
         FIDAT# = FileDateTime(DIRFSX$ + "\" + FINAX$)
         If LOARX > 0 Then
           If FIDAT# > 0 Then
             CARCHI% = CARCHI% + 1
             FEARCH#(CARCHI%) = FIDAT#
             NOMARCH$(CARCHI%) = FINAX$
           End If
         End If
         On Error GoTo 0
       End If
     End If
   Next U&
   '
5  List1.Clear
   ILISTA& = 0
   List1.AddItem DIEXPLO$
   '
10 Dir1.Path = List1.List(ILISTA&)
   Dir1.Refresh
   DoEvents
   For KKI% = 1 To Dir1.ListCount
     TTXX$ = Dir1.List(KKI% - 1) + "\"
     List1.AddItem TTXX$
     List1.ListIndex = List1.ListCount - 1
     List1.Refresh
     DoEvents
   Next KKI%
   '
   If ILISTA& < List1.ListCount - 1 Then
     ILISTA& = ILISTA& + 1
     GoTo 10
   End If
   '
20 For ILISTA& = 1 To List1.ListCount
     File1.Path = List1.List(ILISTA& - 1)
     If UCase$(File1.Path) = DIRFSX$ Then GoTo 291
     DoEvents
     File1.Pattern = "*.RFS"
     For KKI% = 1 To File1.ListCount
       FINAME$ = TRIM$(UCase$(File1.List(KKI% - 1)))
       If FINAME$ <> "" Then
         ACOMPLE$ = List1.List(ILISTA& - 1) + FINAME$
         If FileLen(ACOMPLE$) > 0 Then
           LARCHI& = FileLen(ACOMPLE$)
           Label1.Caption = Str$(Val(Label1.Caption) + LARCHI&)
           DoEvents
           FECHAC# = FileDateTime(ACOMPLE$)
           If FECHAC# <= CDbl(Now) Then
             For KKJ% = 1 To CARCHI%
               If FINAME$ = NOMARCH(KKJ%) Then
                 If FECHAC# > FEARCH#(KKJ%) Or (FECHAC# = FEARCH#(KKJ%) And InStr(ACOMPLE$, "\RFSX") > 0) Then
                   FEARCH#(KKJ%) = FECHAC#
                   RUTARCH$(KKJ%) = List1.List(ILISTA& - 1)
                 End If
                 GoTo 29
               End If
             Next KKJ%
             '
             If Asc(Left$(FINAME$, 1)) >= 65 Then
               If Asc(Left$(FINAME$, 1)) <= 90 Then
                 CARCHI% = CARCHI% + 1
                 If FECHAC# > CDbl(Now) Then FECHAC# = 0
                 NOMARCH$(CARCHI%) = FINAME$
                 FEARCH#(CARCHI%) = FECHAC#
                 RUTARCH$(CARCHI%) = List1.List(ILISTA& - 1)
               End If
             End If
           End If
         End If
       End If
29   Next KKI%
291 Next ILISTA&
   '
   If Left$(DIEXPLO$, 2) <> "C:" Then
     DIEXPLO$ = "C:\"
     GoTo 5
   End If
   '
   For KKI% = 1 To CARCHI%
     If TRIM$(RUTARCH$(KKI%)) <> "" Then
       FERFSX# = 0
       On Error Resume Next
       FERFSX# = CDbl(FileDateTime(ULODATOS$ + "\RFSX\" + NOMARCH(KKI%)))
       On Error GoTo 0
       If FEARCH#(KKI%) > FERFSX# Then
         DATU$ = Format$(FEARCH#(KKI%), "DD-MM-YY")
         DATR$ = Format$(FERFSX#, "DD-MM-YY")
         AAA = DATU$ + "   " + RUTARCH$(KKI%) + NOMARCH$(KKI%) + "   " + DATR$
         List2.AddItem AAA
         On Error Resume Next
         FINUE$ = LOADFILE$(RUTARCH$(KKI%) + NOMARCH$(KKI%))
         FIVIE$ = LOADFILE$(ULODATOS$ + "\RFSX\" + NOMARCH$(KKI%))
         If FINUE$ <> FIVIE$ Then
           If FINUE$ <> "" Then
             FileCopy RUTARCH$(KKI%) + NOMARCH$(KKI%), ULODATOS$ + "\RFSX\" + NOMARCH$(KKI%)
           End If
         End If
         On Error GoTo 0
       End If
     End If
   Next KKI%
   '
   '
   '
50 For KKI% = 1 To CARCHI%
     NOMARCH$(KKI%) = ""
     FEARCH#(KKI%) = 0
     RUTARCH$(KKI%) = ""
   Next KKI%
   '
   DIEXPLO$ = ULODATOS$ + "\"
   CARCHI% = 0
   '
   DIRFSX$ = ULODATOS$ + "\PRFX"
   On Error Resume Next
   MkDir DIRFSX$
   On Error GoTo 0
   '
   Dir1.Path = DIRFSX$
   File1.Path = Dir1.Path
   File1.Pattern = "*.PRF"
   '
   For U& = 1 To File1.ListCount
     FINAX$ = UCase$(File1.List(U& - 1))
     If Asc(Left$(FINAX$, 1)) >= 65 Then
       If Asc(Left$(FINAX$, 1)) <= 90 Then
         LOARX = 0: FIDAT# = 0
         On Error Resume Next
         LOARX = FileLen(DIRFSX$ + "\" + FINAX$)
         FIDAT# = FileDateTime(DIRFSX$ + "\" + FINAX$)
         If LOARX > 0 Then
           If FIDAT# > 0 Then
             CARCHI% = CARCHI% + 1
             FEARCH#(CARCHI%) = FIDAT#
             NOMARCH$(CARCHI%) = FINAX$
           End If
         End If
         On Error GoTo 0
       End If
     End If
   Next U&
   '
55 List1.Clear
   ILISTA& = 0
   List1.AddItem DIEXPLO$
   '
60 Dir1.Path = List1.List(ILISTA&)
   Dir1.Refresh
   DoEvents
   For KKI% = 1 To Dir1.ListCount
     TTXX$ = Dir1.List(KKI% - 1) + "\"
     List1.AddItem TTXX$
     List1.ListIndex = List1.ListCount - 1
     List1.Refresh
     DoEvents
   Next KKI%
   '
   If ILISTA& < List1.ListCount - 1 Then
     ILISTA& = ILISTA& + 1
     GoTo 60
   End If
   '
70 For ILISTA& = 1 To List1.ListCount
     File1.Path = List1.List(ILISTA& - 1)
     If UCase$(File1.Path) = DIRFSX$ Then GoTo 791
     DoEvents
     File1.Pattern = "*.PRF"
     For KKI% = 1 To File1.ListCount
       FINAME$ = TRIM$(UCase$(File1.List(KKI% - 1)))
       If FINAME$ <> "" Then
         ACOMPLE$ = List1.List(ILISTA& - 1) + FINAME$
         If FileLen(ACOMPLE$) > 0 Then
           LARCHI& = FileLen(ACOMPLE$)
           Label1.Caption = Str$(Val(Label1.Caption) + LARCHI&)
           DoEvents
           FECHAC# = FileDateTime(ACOMPLE$)
           If FECHAC# <= CDbl(Now) Then
             For KKJ% = 1 To CARCHI%
               If FINAME$ = NOMARCH(KKJ%) Then
                 If FECHAC# > FEARCH#(KKJ%) Or (FECHAC# = FEARCH#(KKJ%) And InStr(ACOMPLE$, "\PRFX") > 0) Then
                   FEARCH#(KKJ%) = FECHAC#
                   RUTARCH$(KKJ%) = List1.List(ILISTA& - 1)
                 End If
                 GoTo 79
               End If
             Next KKJ%
             '
             If Asc(Left$(FINAME$, 1)) >= 65 Then
               If Asc(Left$(FINAME$, 1)) <= 90 Then
                 CARCHI% = CARCHI% + 1
                 If FECHAC# > CDbl(Now) Then FECHAC# = 0
                 NOMARCH$(CARCHI%) = FINAME$
                 FEARCH#(CARCHI%) = FECHAC#
                 RUTARCH$(CARCHI%) = List1.List(ILISTA& - 1)
               End If
             End If
           End If
         End If
       End If
79   Next KKI%
791 Next ILISTA&
   '
   If Left$(DIEXPLO$, 2) <> "C:" Then
     DIEXPLO$ = "C:\"
     GoTo 55
   End If
   '
   For KKI% = 1 To CARCHI%
     If TRIM$(RUTARCH$(KKI%)) <> "" Then
       FERFSX# = 0
       On Error Resume Next
       FERFSX# = CDbl(FileDateTime(ULODATOS$ + "\PRFX\" + NOMARCH(KKI%)))
       On Error GoTo 0
       If FEARCH#(KKI%) > FERFSX# Then
         DATU$ = Format$(FEARCH#(KKI%), "DD-MM-YY")
         DATR$ = Format$(FERFSX#, "DD-MM-YY")
         AAA = DATU$ + "   " + RUTARCH$(KKI%) + NOMARCH$(KKI%) + "   " + DATR$
         List2.AddItem AAA
         On Error Resume Next
         FINUE$ = LOADFILE$(RUTARCH$(KKI%) + NOMARCH$(KKI%))
         FIVIE$ = LOADFILE$(ULODATOS$ + "\PRFX\" + NOMARCH$(KKI%))
         If FINUE$ <> FIVIE$ Then
           If FINUE$ <> "" Then
             FileCopy RUTARCH$(KKI%) + NOMARCH$(KKI%), ULODATOS$ + "\PRFX\" + NOMARCH$(KKI%)
           End If
         End If
         On Error GoTo 0
       End If
     End If
   Next KKI%
   '
   Screen.MousePointer = 11
   '
   NX% = FreeFile
   Open "C:\FLEXOFT\DOSCOMM.BAT" For Output As #NX%
   Print #NX%, "\FLEXOWIN\PKZIP F:\FLEXOWIN\FLEXRFS " + ULODATOS$ + "\RFSX\*.RFS"
   Print #NX%, "\FLEXOWIN\PKZIP F:\FLEXOWIN\FLEXPRF " + ULODATOS$ + "\PRFX\*.PRF"
   Close #NX%: NX% = 0
   Shell "C:\FLEXOFT\DOSCOMM.PIF"
   '
99 On Error Resume Next
   Kill ULODATOS$ + "\FLEXOWIN\FLEXPRFS.FLX"
   On Error GoTo 0
   '
   Dim BUFLE As String * 128
   Dim BUFGRA As String * 160
   NX% = FreeFile
   Open ULODATOS$ + "\FLEXOWIN\FLEXPRFS.FLX" For Random Access Read Write Shared As #NX% Len = 160
   JJ& = 0
   '
   List1.Clear
   File1.Path = ULODATOS$ + "\RFSX"
   DoEvents
   File1.Pattern = "*.RFS"
   '
   For KKI% = 1 To File1.ListCount
     FINAME$ = TRIM$(UCase$(File1.List(KKI% - 1)))
     If FINAME$ <> "" Then
       For JKI% = 1 To List3.ListCount
         If List3.List(JKI% - 1) > FINAME$ Then
           List3.AddItem FINAME$, JKI% - 1
           GoTo 14
         End If
       Next JKI%
       List3.AddItem FINAME$
     End If
14 Label2 = List3.ListCount
   Next KKI%
   '
   File1.Path = ULODATOS$ + "\PRFX"
   DoEvents
   File1.Pattern = "*.PRF"
   '
   For KKI% = 1 To File1.ListCount
     FINAME$ = TRIM$(UCase$(File1.List(KKI% - 1)))
     If FINAME$ <> "" Then
       For JKI% = 1 To List3.ListCount
         If List3.List(JKI% - 1) > FINAME$ Then
           List3.AddItem FINAME$, JKI% - 1
           GoTo 16
         End If
       Next JKI%
       List3.AddItem FINAME$
     End If
16 Label2 = List3.ListCount
   Label2.Refresh
   Next KKI%
   '
   Screen.MousePointer = 11
   For KKI% = 1 To List3.ListCount
     FINAME$ = TRIM$(UCase$(List3.List(KKI% - 1)))
     If FINAME$ <> "" Then
       If Left$(FINAME$, 1) <> "." Then
         CARPETRA$ = ULODATOS$ + "\RFSX"
         If InStr(FINAME$, ".PRF") > 0 Then
           CARPETRA$ = ULODATOS$ + "\PRFX"
         End If
         '
         ACOMPLE$ = CARPETRA$ + "\" + FINAME$
         LARCHI& = 0
         On Error Resume Next
         LARCHI& = FileLen(ACOMPLE$)
         On Error GoTo 0
         CARELE& = Int((LARCHI + 127) / 128)
         If LARCHI& > 0 Then
           N1% = FreeFile
           Open ACOMPLE$ For Random Access Read As #N1% Len = 128
           For II& = 1 To CARELE&
             Get #N1%, II&, BUFLE$
             BUFGRA$ = Space$(160)
             Mid$(BUFGRA$, 1, 32) = FINAME$
             Mid$(BUFGRA$, 33, 128) = BUFLE$
             JJ& = JJ& + 1
             Put #NX%, JJ&, BUFGRA$
           Next II&
           Close #N1%
         End If
       End If
     End If
   Next KKI%
   '
   Screen.MousePointer = 1
   List3.Visible = True
   Close #NX%
   '
   MsgBox "Proceso Completo"
   '
End Sub


