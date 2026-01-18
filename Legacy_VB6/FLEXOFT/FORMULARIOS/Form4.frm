VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6360
   ClientLeft      =   60
   ClientTop       =   315
   ClientWidth     =   7140
   LinkTopic       =   "Form1"
   ScaleHeight     =   6360
   ScaleWidth      =   7140
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      Height          =   225
      Left            =   5040
      TabIndex        =   9
      Top             =   0
      Width           =   1380
   End
   Begin VB.ListBox List3 
      Height          =   645
      Left            =   5565
      TabIndex        =   7
      Top             =   1575
      Width           =   1380
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Generar Disco   Instalacion"
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
      Left            =   5565
      TabIndex        =   6
      Top             =   945
      Width           =   1380
   End
   Begin VB.ListBox List2 
      Height          =   2595
      Left            =   210
      TabIndex        =   4
      Top             =   3360
      Width           =   6735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   435
      Left            =   5565
      TabIndex        =   3
      Top             =   315
      Width           =   1380
   End
   Begin VB.ListBox List1 
      Height          =   645
      Left            =   210
      TabIndex        =   2
      Top             =   2520
      Width           =   5160
   End
   Begin VB.FileListBox File1 
      Height          =   1845
      Left            =   3045
      TabIndex        =   1
      Top             =   315
      Width           =   2325
   End
   Begin VB.DirListBox Dir1 
      Height          =   1890
      Left            =   210
      TabIndex        =   0
      Top             =   315
      Width           =   2325
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   5565
      TabIndex        =   8
      Top             =   2415
      Width           =   1380
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   5565
      TabIndex        =   5
      Top             =   2835
      Width           =   1380
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
   Dim NOMARCH$(4096), FEARCH#(4096), RUTARCH$(4096)
   '
   DIEXPLO$ = "C:\"
   CARCHI% = 0
   '
   DIRFSX$ = "C:\RFSX"
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
   For KKI% = 1 To Dir1.ListCount
     TTXX$ = Dir1.List(KKI% - 1) + "\"
     List1.AddItem TTXX$
     List1.ListIndex = List1.ListCount - 1
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
       FINAME$ = Trim$(UCase$(File1.List(KKI% - 1)))
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
   If DIEXPLO$ = "C:\" Then
     DIEXPLO$ = "F:\"
     GoTo 5
   End If
   '
   For KKI% = 1 To CARCHI%
     If Trim$(RUTARCH$(KKI%)) <> "" Then
       FERFSX# = 0
       On Error Resume Next
       FERFSX# = CDbl(FileDateTime("C:\RFSX\" + NOMARCH(KKI%)))
       On Error GoTo 0
       If FEARCH#(KKI%) > FERFSX# Then
         DATU$ = Format$(FEARCH#(KKI%), "DD-MM-YY")
         DATR$ = Format$(FERFSX#, "DD-MM-YY")
         AAA = DATU$ + "   " + RUTARCH$(KKI%) + NOMARCH$(KKI%) + "   " + DATR$
         List2.AddItem AAA
         On Error Resume Next
         FINUE$ = LOADFILE$(RUTARCH$(KKI%) + NOMARCH$(KKI%))
         FIVIE$ = LOADFILE$("C:\RFSX\" + NOMARCH$(KKI%))
         If FINUE$ <> FIVIE$ Then
           If FINUE$ <> "" Then
             FileCopy RUTARCH$(KKI%) + NOMARCH$(KKI%), "C:\RFSX\" + NOMARCH$(KKI%)
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
   DIEXPLO$ = "C:\"
   CARCHI% = 0
   '
   DIRFSX$ = "C:\PRFX"
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
   For KKI% = 1 To Dir1.ListCount
     TTXX$ = Dir1.List(KKI% - 1) + "\"
     List1.AddItem TTXX$
     List1.ListIndex = List1.ListCount - 1
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
       FINAME$ = Trim$(UCase$(File1.List(KKI% - 1)))
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
   If DIEXPLO$ = "C:\" Then
     DIEXPLO$ = "F:\"
     GoTo 55
   End If
   '
   For KKI% = 1 To CARCHI%
     If Trim$(RUTARCH$(KKI%)) <> "" Then
       FERFSX# = 0
       On Error Resume Next
       FERFSX# = CDbl(FileDateTime("C:\PRFX\" + NOMARCH(KKI%)))
       On Error GoTo 0
       If FEARCH#(KKI%) > FERFSX# Then
         DATU$ = Format$(FEARCH#(KKI%), "DD-MM-YY")
         DATR$ = Format$(FERFSX#, "DD-MM-YY")
         AAA = DATU$ + "   " + RUTARCH$(KKI%) + NOMARCH$(KKI%) + "   " + DATR$
         List2.AddItem AAA
         On Error Resume Next
         FINUE$ = LOADFILE$(RUTARCH$(KKI%) + NOMARCH$(KKI%))
         FIVIE$ = LOADFILE$("C:\PRFX\" + NOMARCH$(KKI%))
         If FINUE$ <> FIVIE$ Then
           If FINUE$ <> "" Then
             FileCopy RUTARCH$(KKI%) + NOMARCH$(KKI%), "C:\PRFX\" + NOMARCH$(KKI%)
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
   Print #NX%, "\FLEXOWIN\PKZIP F:\FLEXOWIN\FLEXRFS C:\RFSX\*.RFS"
   Print #NX%, "\FLEXOWIN\PKZIP F:\FLEXOWIN\FLEXPRF C:\PRFX\*.PRF"
   Close #NX%: NX% = 0
   Shell "C:\FLEXOFT\DOSCOMM.PIF"
   '
99 On Error Resume Next
   Kill "F:\FLEXOWIN\FLEXPRFS.FLX"
   On Error GoTo 0
   '
   Dim BUFLE As String * 128
   Dim BUFGRA As String * 160
   NX% = FreeFile
   Open "F:\FLEXOWIN\FLEXPRFS.FLX" For Random Access Read Write Shared As #NX% Len = 160
   JJ& = 0
   '
   List1.Clear
   File1.Path = "C:\RFSX"
   DoEvents
   File1.Pattern = "*.RFS"
   '
   For KKI% = 1 To File1.ListCount
     FINAME$ = Trim$(UCase$(File1.List(KKI% - 1)))
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
   File1.Path = "C:\PRFX"
   DoEvents
   File1.Pattern = "*.PRF"
   '
   For KKI% = 1 To File1.ListCount
     FINAME$ = Trim$(UCase$(File1.List(KKI% - 1)))
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
     FINAME$ = Trim$(UCase$(List3.List(KKI% - 1)))
     If FINAME$ <> "" Then
       If Left$(FINAME$, 1) <> "." Then
         CARPETRA$ = "C:\RFSX"
         If InStr(FINAME$, ".PRF") > 0 Then
           CARPETRA$ = "C:\PRFX"
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
   
'   File1.Path = "C:\PRFX"
'   DoEvents
'   File1.Pattern = "*.PRF"
'   For KKI% = 1 To File1.ListCount
'     FINAME$ = Trim$(UCase$(File1.List(KKI% - 1)))
'     If FINAME$ <> "" Then
'       ACOMPLE$ = File1.Path + "\" + FINAME$
'       LARCHI& = FileLen(ACOMPLE$)
'       CARELE& = Int((LARCHI + 127) / 128)
'       If LARCHI& > 0 Then
'         N1% = FreeFile
'         Open ACOMPLE$ For Random Access Read As #N1% Len = 128
'         For II& = 1 To CARELE&
'           Get #N1%, II&, BUFLE$
'           BUFGRA$ = Space$(160)
'           Mid$(BUFGRA$, 1, 32) = FINAME$
'           Mid$(BUFGRA$, 33, 128) = BUFLE$
'           JJ& = JJ& + 1
'           Put #NX%, JJ&, BUFGRA$
'         Next II&
'         Close #N1%
'       End If
'     End If
'   Next KKI%
'   '
   Close #NX%
   '
   MsgBox "Proceso Completo"
   '
End Sub

Private Sub Command2_Click()
   Command2.Enabled = False
   Call GEDINSTA
   Command2.Enabled = True
End Sub

Sub GEDINSTA()    ' genera diskette de instalacion
    '
    If SERIVOL$("C:") <> "081D-0BD9" Then Exit Sub
    '
10  MsgBox "Coloque 'Diskette' de Instalación en Blanco"
    '
    On Error GoTo 15
    Dir1.Path = "A:\"
    Dir1.Refresh
    File1.Refresh
    On Error GoTo 0
    GoTo 20
    '
15  Resume 16
16  On Error GoTo 0
    MsgBox "Error de Acceso a la Disketera"
    GoTo 10
    '
20  If Dir1.ListCount > 0 Then
      MsgBox "Diskette no Válido - Tiene Información"
      GoTo 10
    End If
    '
    File1.Path = "A:\"
    File1.Refresh
    If File1.ListCount > 0 Then
      MsgBox "Diskette no Válido - Tiene Información"
      GoTo 10
    End If
    '
40  Dim BLEC As String * 127
    On Error GoTo 45
    N1% = FreeFile
    Open "A:FLEXKEY.KEY" For Random Access Read Write Shared As #N1% Len = 127
    On Error GoTo 0
    GoTo 50
    '
45  Resume 46
46  On Error GoTo 0
    MsgBox "Diskette Protegido - Quite Protección."
    GoTo 10
    '
50  Screen.MousePointer = 11
    SEVOLA$ = Left$(SERIVOL$("A:") + Space$(12), 12)
    If Trim$(SEVOLA$) = "" Then
       Close #N1%
       Screen.MousePointer = 1
       MsgBox "Diskette No Válido."
       GoTo 10
    End If
    '
    For UL& = 1 To 499
       TTXX$ = RANDSTRING$(125)
       If UL& = 1 Then
           Mid$(TTXX$, 17, 12) = ENCRIPTA$(SEVOLA$)
         ElseIf UL& = 2 Then
           Mid$(TTXX$, 29, 12) = ENCRIPTA$(String$(12, 19))
         ElseIf UL& = 3 Then
           Mid$(TTXX$, 71, 12) = ENCRIPTA$(String$(12, 23))
       End If
       Get #N1%, UL&, BLEC$
       SUCAR% = 0
       For KI% = 1 To 125
          SUCAR% = SUCAR% + Asc(Mid$(TTXX$, KI%, 1)) - 83
       Next KI%
       BLEC$ = TTXX$ + MKI$(SUCAR%)
       Put #N1%, UL&, BLEC$
    Next UL&
    '
    Close #N1%
    Screen.MousePointer = 1
    MsgBox "Diskette Generado Exitosamente"
    '
99 End Sub

Function SERIVOL$(ULOVOL$)
    '
    Dim AA$(32)
    SERIVOL$ = ""
    '
    NX% = FreeFile
    On Error GoTo 98
    Open "C:\FLEXOFT\DOSCOMM.BAT" For Output As #NX%
    Print #NX%, "VOL " + ULOVOL$ + " >C:\FLEXOFT\SERIVOL"
    Close #NX%: NX% = 0
    Shell "C:\FLEXOFT\DOSCOMM.PIF"
    TINI = Timer
    While Timer < TINI + 3: Wend
    '
    NX% = FreeFile
    Open "C:\FLEXOFT\SERIVOL" For Input As #NX%
    AX% = 0
    While Not EOF(NX%)
        Line Input #NX%, XX$
        If AX% < 32 Then
            AX% = AX% + 1
            AA$(AX%) = XX$
        End If
    Wend
    Close #NX%: NX% = 0
    '
    For AY% = 1 To AX%
        XX$ = Trim$(AA$(AY%))
        POSI% = InStr(XX$, "-")
        If POSI% > 5 Then
           If POSI% <= Len(XX$) - 4 Then
              If Mid$(XX$, POSI% - 5, 1) = " " Then
                 SERIVOL$ = Mid$(XX$, POSI% - 4, 9)
                 Exit Function
              End If
           End If
        End If
    Next AY%
    SERIVOL$ = "0000-0000"
    Exit Function
    '
98  Resume 99
99  SERIVOL$ = "1234-5678"
    Exit Function
    '
End Function

Function ENCRIPTA$(DE$)
    '
    DE1$ = ""
    For I% = 1 To Len(DE$)
        XL% = Asc(Mid$(DE$, I%))
        Y% = (XL% + 23 * I% + 17) Mod 256
        DE1$ = DE1$ + Chr$(Y%)
    Next I%
    ENCRIPTA$ = DE1$
    '
End Function
'

Function DESENCRI$(DE$)
    '
    DE1$ = ""
    For I% = 1 To Len(DE$)
        XL% = Asc(Mid$(DE$, I%))
        Y% = (XL% - 23 * I% - 17)
        While Y% < 0
            Y% = Y% + 256
        Wend
        DE1$ = DE1$ + Chr$(Y%)
    Next I%
    DESENCRI$ = DE1$
    '
End Function
'

Function RANDSTRING$(LON%)
    '
    RST$ = ""
    Randomize Timer
    For KL% = 1 To LON%
        PPX% = Int(Rnd * 255)
        RST$ = RST$ + Chr$(PPX%)
    Next KL%
    '
    RANDSTRING$ = RST$
    '
End Function
'
Function MKI$(A%)
    '
    AA& = A%
    If AA& < 0 Then AA& = AA& + 65536
    MKI$ = Chr$(AA& Mod 256) + Chr$(AA& \ 256)
    '
End Function
'

Private Sub Command3_Click()
Dim BLEC As String * 127
N1% = FreeFile
Open "A:FLEXKEY.KEY" For Random Access Read Write As #N1% Len = 127
Get #N1%, 2, BLEC
ABC$ = DESENCRI$(Mid$(BLEC$, 59, 38))
BBB = ABC
Close #N1%
End Sub
