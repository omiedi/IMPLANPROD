VERSION 5.00
Begin VB.Form APARFUEN 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Aparea Fuentes de Notebook y PC-Ernesto"
   ClientHeight    =   6135
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   9675
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6135
   ScaleWidth      =   9675
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command3 
      Caption         =   "BUSQUEDA"
      Height          =   255
      Left            =   3960
      TabIndex        =   7
      Top             =   0
      Width           =   2415
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   255
      Left            =   7920
      TabIndex        =   6
      Top             =   0
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2040
      TabIndex        =   5
      Top             =   300
      Width           =   3495
   End
   Begin VB.CommandButton Command1 
      Height          =   1830
      Left            =   8280
      Picture         =   "APARFUEN.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Inicia Exploraciónde la Red"
      Top             =   315
      Width           =   1170
   End
   Begin VB.ListBox List1 
      Height          =   3570
      Left            =   210
      TabIndex        =   2
      Top             =   2310
      Width           =   9240
   End
   Begin VB.FileListBox File1 
      Height          =   1845
      Left            =   5640
      TabIndex        =   1
      Top             =   300
      Width           =   2430
   End
   Begin VB.DirListBox Dir1 
      Height          =   1215
      Left            =   210
      TabIndex        =   0
      Top             =   900
      Width           =   5295
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Equipo Remoto:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   4
      Top             =   350
      Width           =   1935
   End
End
Attribute VB_Name = "APARFUEN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
   '
   CARCHI% = 0
   '
   OTROEQ$ = TRIM$(Combo1.TEXT)
   If OTROEQ$ = "" Then
      MsgBox "Identifique Equipo Remoto !!!"
      Exit Sub
   End If
   '
   ESTEQ$ = "C:"
   DIRFSX$ = OTROEQ$ + "\ARCHIVOS DE PROGRAMA\DEVSTUDIO\VB\FLEXOFT"
   DIEXPLO$ = DIRFSX$ + "\"
   '
   Screen.MousePointer = 11
   '
   Dir1.Path = DIRFSX$
   File1.Path = Dir1.Path
   File1.Pattern = "*.*"
   '
5  ' CARGA EN LIST1 LA LISTA DE LAS CARPETAS A EXPLORAR *************
   List1.Clear
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
20 PPXX% = Len(OTROEQ$)
   For ILISTA& = 1 To List1.ListCount
     List1.ListIndex = ILISTA& - 1
     File1.Path = List1.List(ILISTA& - 1)
     DoEvents
     File1.Pattern = "*.*"
     '
     For KKI% = 1 To File1.ListCount
       FINAME$ = TRIM$(UCase$(File1.List(KKI% - 1)))
       If FINAME$ <> "" Then
         COMPAFU.Label6 = FINAME$
         ACOMPLE$ = List1.List(ILISTA& - 1) + FINAME$
         ESTEAR$ = ESTEQ$ + Mid$(ACOMPLE$, PPXX% + 1)
         '
         If FINAME$ = "BADASTRU.MDB" Or InStr(FINAME$, ".LOG") > 0 Or InStr(FINAME$, ".PRF") > 0 Then
            On Error Resume Next
            Kill ACOMPLE$
            Kill ESTEAR$
            On Error GoTo 0
            GoTo 29
         End If
         '
         If FileLen(ACOMPLE$) > 0 Then
           LARCHI& = FileLen(ACOMPLE$)
           FECHOTRO# = FileDateTime(ACOMPLE$)
           FECHESTE# = 0
           On Error Resume Next
             FECHESTE# = FileDateTime(ESTEAR$)
           On Error GoTo 0
           If FECHOTRO# <> FECHESTE# Then
             ARZX1$ = LOADFILE$(ACOMPLE$)
             ARZX2$ = LOADFILE$(ESTEAR$)
             If ARZX1$ <> ARZX2$ Then
               If UCase$(Right$(ESTEAR$, 4)) = ".FRM" Or UCase$(Right$(ESTEAR$, 4)) = ".BAS" Then
                 Screen.MousePointer = 1
                 Call COMPARFUEN$(ESTEAR$, ACOMPLE$)
                 GoTo 29
                 Screen.MousePointer = 11
               End If
               If FECHOTRO# > FECHESTE# Then
                   FFAB$ = Format$(FECHESTE#, "DD/MM/yy  hh:mm:ss")
                   If MsgBox("Desea Actualizar Versión LOCAL (" + FFAB$ + ") de" + Chr$(10) + Chr$(13) + TRIM$(FINAME$) + " al " + Format$(FECHOTRO#, "DD/MM/YY  HH:MM:SS"), 4) = 6 Then
                     DoEvents
                     FileCopy ACOMPLE$, ESTEAR$
                   End If
                   DoEvents
                 ElseIf FECHOTRO# < FECHESTE# Then
                   FFAB$ = Format$(FECHOTRO#, "DD/MM/yy  hh:mm:ss")
                   If MsgBox("Desea Actualizar Versión REMOTA (" + FFAB$ + ") de" + Chr$(10) + Chr$(13) + TRIM$(FINAME$) + " al " + Format$(FECHESTE#, "DD/MM/YY  HH:MM:SS"), 4) = 6 Then
                     DoEvents
                     FileCopy ESTEAR$, ACOMPLE$
                   End If
                   DoEvents
               End If
             End If
           End If
         End If
       End If
29   Next KKI%
   Next ILISTA&
   '
   'ESTEQ$ = "C:"
   DIRFSX$ = "C:\ARCHIVOS DE PROGRAMA\DEVSTUDIO\VB\FLEXOFT"
   DIEXPLO$ = DIRFSX$ + "\"
   '
   Screen.MousePointer = 11
   '
   Dir1.Path = DIRFSX$
   File1.Path = Dir1.Path
   File1.Pattern = "*.*"
   '
35 ' CARGA EN LIST1 LA LISTA DE LAS CARPETAS A EXPLORAR *************
   List1.Clear
   ILISTA& = 0
   List1.AddItem DIEXPLO$
   '
40 Dir1.Path = List1.List(ILISTA&)
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
     GoTo 40
   End If
   '
50 PPXX% = 2
   For ILISTA& = 1 To List1.ListCount
     List1.ListIndex = ILISTA& - 1
     File1.Path = List1.List(ILISTA& - 1)
     DoEvents
     File1.Pattern = "*.*"
     '
     For KKI% = 1 To File1.ListCount
       FINAME$ = TRIM$(UCase$(File1.List(KKI% - 1)))
       If FINAME$ <> "" Then
         ACOMPLE$ = List1.List(ILISTA& - 1) + FINAME$
         OTRAR$ = OTROEQ$ + Mid$(ACOMPLE$, PPXX% + 1)
         '
         If FINAME$ = "BADASTRU.MDB" Or InStr(FINAME$, ".LOG") > 0 Or InStr(FINAME$, ".PRF") > 0 Or InStr(FINAME$, ".DRV") > 0 Or InStr(FINAME$, ".SEL") > 0 Then
            On Error Resume Next
            Kill ACOMPLE$
            Kill OTRAR$
            On Error GoTo 0
            GoTo 59
         End If
         '
         If FileLen(ACOMPLE$) > 0 Then
           LARCHI& = FileLen(ACOMPLE$)
           FECHESTE# = FileDateTime(ACOMPLE$)
           FECHOTRO# = 0
           On Error Resume Next
             FECHOTRO# = FileDateTime(OTRAR$)
           On Error GoTo 0
           If FECHOTRO# <> FECHESTE# Then
             ARZX1$ = LOADFILE$(ACOMPLE$)
             ARZX2$ = LOADFILE$(OTRAR$)
             If ARZX1$ <> ARZX2$ Then
               If FECHOTRO# > FECHESTE# Then
                   FFAB$ = Format$(FECHESTE#, "DD/MM/yy  hh:mm:ss")
                   If MsgBox("Desea Actualizar Versión LOCAL (" + FFAB$ + ") de" + Chr$(10) + Chr$(13) + TRIM$(FINAME$) + " al " + Format$(FECHOTRO#, "DD/MM/YY  HH:MM:SS"), 4) = 6 Then
                     DoEvents
                     FileCopy OTRAR$, ACOMPLE$
                   End If
                   DoEvents
                 ElseIf FECHOTRO# < FECHESTE# Then
                   FFAB$ = Format$(FECHOTRO#, "DD/MM/yy  hh:mm:ss")
                   If MsgBox("Desea Actualizar Versión REMOTA (" + FFAB$ + ") de" + Chr$(10) + Chr$(13) + TRIM$(FINAME$) + " al " + Format$(FECHESTE#, "DD/MM/YY  HH:MM:SS"), 4) = 6 Then
                     DoEvents
                     FileCopy ACOMPLE$, OTRAR$
                   End If
                   DoEvents
               End If
             End If
           End If
         End If
       End If
59   Next KKI%
   Next ILISTA&
   '
   Screen.MousePointer = 1
   MsgBox "Proceso Completo"
   '
End Sub


Private Sub Command2_Click()
   ESTEQ$ = "C:"
   OTROQ$ = "\\RECEPCION\D"
   AR1$ = OTROQ$ + "\ARCHIVOS DE PROGRAMA\DEVSTUDIO\VB\FLEXOFT\BIBLIOTE\RAM_FUNCTIONS.BAS"
   LON1 = FileLen(AR1$)
   N1% = FILEOPEN%(AR1$, 1, 128)
   AR2$ = ESTEQ$ + "\ARCHIVOS DE PROGRAMA\DEVSTUDIO\VB\FLEXOFT\BIBLIOTE\RAM_FUNCTIONS.BAS"
   N2% = FILEOPEN%(AR2$, 1, 128)
   '
10 Line Input #N1%, A1$
   Line Input #N2%, A2$
   List1.AddItem A1$
   List1.AddItem A2$
   List1.AddItem "--------------------------------------------"
   If A1$ = A2$ Then GoTo 10
   '
End Sub

Private Sub Command3_Click()
   CARCHI% = 0
   '
   OTROEQ$ = TRIM$(Combo1.TEXT)
   If OTROEQ$ = "" Then
      MsgBox "Identifique Equipo Remoto !!!"
      Exit Sub
   End If
   '
   ESTEQ$ = "C:"
   DIRFSX$ = OTROEQ$ + "\ARCHIVOS DE PROGRAMA\DEVSTUDIO\VB\FLEXOFT"
   DIEXPLO$ = DIRFSX$ + "\"
   '
   Screen.MousePointer = 11
   '
   Dir1.Path = DIRFSX$
   File1.Path = Dir1.Path
   File1.Pattern = "*.*"
   '
5  ' CARGA EN LIST1 LA LISTA DE LAS CARPETAS A EXPLORAR *************
   List1.Clear
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
20 PPXX% = Len(OTROEQ$)
   For ILISTA& = 1 To List1.ListCount
     List1.ListIndex = ILISTA& - 1
     File1.Path = List1.List(ILISTA& - 1)
     DoEvents
     File1.Pattern = "*.*"
     '
     For KKI% = 1 To File1.ListCount
       FINAME$ = TRIM$(UCase$(File1.List(KKI% - 1)))
       If FINAME$ <> "" Then
         ACOMPLE$ = List1.List(ILISTA& - 1) + FINAME$
         ESTEAR$ = ESTEQ$ + Mid$(ACOMPLE$, PPXX% + 1)
         If FileLen(ACOMPLE$) > 0 Then
           ABC1$ = UCase$(LOADFILE$(ACOMPLE$))
           If InStr(ABC1$, "PRESUP04.EXE") > 0 Then
             MsgBox ACOMPLE$
           End If
         End If
       End If
29   Next KKI%
   Next ILISTA&
   '
   Screen.MousePointer = 1
   MsgBox "Proceso Completo"
   '
End Sub

Private Sub Form_Load()
   Combo1.Clear
   Combo1.AddItem "\\NOTEBOOK\C_NOTEBOOK"
   Combo1.AddItem "\\Soporte\D"
   Combo1.AddItem "\\Desarrollo\Datos (D)"
   Combo1.AddItem "\\Ventas\DATOS (D)"
   Combo1.AddItem "\\Ernesto\sistema (C)"
End Sub

Sub COMPARFUEN(ARLOC$, AREMO$)
   '
   If EXISTE%(ARLOC$) < 1 Then Exit Sub
   If EXISTE%(AREMO$) < 1 Then Exit Sub
   '
   FECHESTE# = FileDateTime(ARLOC$)
   COMPAFU.Label3 = Format$(FECHESTE#, "DD/MM/yy  hh:mm:ss")
   FECHOTRO# = FileDateTime(AREMO$)
   COMPAFU.Label4 = Format$(FECHOTRO#, "DD/MM/yy  hh:mm:ss")
   '
   COMPAFU.Label8.Visible = False
   COMPAFU.Label9.Visible = False
   If FECHOTRO# > FECHESTE# Then
       COMPAFU.Command1.Caption = "Actualiza LOCAL"
       COMPAFU.Label9.Visible = True
     Else
       COMPAFU.Command1.Caption = "Actualiza REMOTO"
       COMPAFU.Label8.Visible = True
   End If
   '
   COMPAFU.Label10 = ARLOC$
   COMPAFU.Label11 = AREMO$
   N1% = FILEOPEN%(ARLOC$, 1, 128)
   N2% = FILEOPEN%(AREMO$, 1, 128)
   PRIDIF& = -1
   CONTA& = 0
   RUTIAC$ = "": RUTIANT$ = ""
   '
   COMPAFU.List1.Clear
   COMPAFU.List2.Clear
   COMPAFU.Text1 = ""
   COMPAFU.Text2 = ""
   '
   COMPAFU.Label2 = Combo1.TEXT + ":"
   COMPAFU.Label4.Left = COMPAFU.Label2.Left + COMPAFU.Label2.Width + 60
   COMPAFU.List3.Clear
   '
         Dim RUTNA1$(512), BLOQ1$(512), RUTNA2$(512), BLOQ2$(512)
         CARUT1% = 0
         Do While Not EOF(N1%)
           Line Input #N1%, TX1$
           TTXX$ = TRIM$(UCase$(TX1$))
           If Left$(TTXX$, 4) = "SUB " Or Left$(TTXX$, 12) = "PRIVATE SUB " Or Left$(TTXX$, 9) = "FUNCTION " Or Left$(TTXX$, 14) = "BEGIN VB.FORM " Then
             CARUT1% = CARUT1% + 1
             RUTNA1$(CARUT1%) = LTrim$(TX1$)
           End If
           BLOQ1$(CARUT1%) = BLOQ1$(CARUT1%) + TX1$ + CRLF$
         Loop
         Close #N1%: N1% = 0
         '
         CARUT2% = 0
         Do While Not EOF(N2%)
           Line Input #N2%, TX2$
           TTXX$ = TRIM$(UCase$(TX2$))
           If Left$(TTXX$, 4) = "SUB " Or Left$(TTXX$, 12) = "PRIVATE SUB " Or Left$(TTXX$, 9) = "FUNCTION " Or Left$(TTXX$, 14) = "BEGIN VB.FORM " Then
             CARUT2% = CARUT2% + 1
             RUTNA2$(CARUT2%) = LTrim$(TX2$)
           End If
           BLOQ2$(CARUT2%) = BLOQ2$(CARUT2%) + TX2$ + CRLF$
         Loop
         Close #N2%: N2% = 0
         '
         For KI1% = 1 To CARUT1%
           For KI2% = 1 To CARUT2%
             If UCase$(BLOQ2$(KI2%)) = UCase$(BLOQ1$(KI1%)) Then
               BLOQ1$(KI1%) = ""
               BLOQ2$(KI2%) = ""
               GoTo 19
             End If
           Next KI2%
19       Next KI1%
         '
         For KI1% = 1 To CARUT1%
           If BLOQ1$(KI1%) <> "" Then
             COMPAFU.List3.AddItem RUTNA1$(KI1%)
             TTXX$ = BLOQ1$(KI1%)
             Do While TTXX$ <> ""
               PPXX% = InStr(TTXX$, CRLF$)
               COMPAFU.List1.AddItem Left$(TTXX$, PPXX% - 1)
               TTXX$ = Mid$(TTXX$, PPXX% + 2)
             Loop
             '
             For KI2% = 1 To CARUT2%
               If UCase$(RUTNA2$(KI2%)) = UCase$(RUTNA1$(KI1%)) Then
                 TTXX$ = BLOQ2$(KI2%)
                 Do While TTXX$ <> ""
                   PPXX% = InStr(TTXX$, CRLF$)
                   COMPAFU.List2.AddItem Left$(TTXX$, PPXX% - 1)
                   TTXX$ = Mid$(TTXX$, PPXX% + 2)
                 Loop
                 BLOQ2$(KI2%) = ""
                 GoTo 29
               End If
             Next KI2%
             '
           End If
29       Next KI1%
         '
         For KI2% = 1 To CARUT2%
           If BLOQ2$(KI2%) <> "" Then
             COMPAFU.List3.AddItem RUTNA2$(KI2%)
             TTXX$ = BLOQ2$(KI2%)
             Do While TTXX$ <> ""
               PPXX% = InStr(TTXX$, CRLF$)
               COMPAFU.List2.AddItem Left$(TTXX$, PPXX% - 1)
               TTXX$ = Mid$(TTXX$, PPXX% + 2)
             Loop
           End If
         Next KI2%
         '
         COMPAFU.Show 1
         Screen.MousePointer = 11
         DoEvents
         Exit Sub
         Stop
               
'   Do While Not EOF(N1%)
'     Line Input #N1%, TX1$
'     COMPAFU.List1.AddItem TX1$
'     If InStr(UCase$(TX1$), "SUB ") > 0 Or InStr(UCase$(TX1$), "FUNCTION ") > 0 Or InStr(UCase$(TX1$), "BEGIN VB.FORM ") > 0 Then
'       RUTIAC$ = TRIM$(TX1$)
'     End If
'     TX2$ = ""
'     If Not EOF(N2%) Then
'       Line Input #N2%, TX2$
'       COMPAFU.List2.AddItem TX2$
'     End If
'     If UCase$(TX1$) <> UCase$(TX2$) Then
'       If RUTIAC$ <> RUTIANT$ Then
'         RUTIANT$ = RUTIAC$
'         COMPAFU.List3.AddItem RUTIAC$
'       End If
'     End If
'     If PRIDIF& < 0 Then
'       If UCase$(TX1$) <> UCase$(TX2$) Then
'         PRIDIF& = CONTA&
'       End If
'     End If
'     CONTA& = CONTA + 1
'   Loop
'   Close #N1%: N1% = 0
'   Close #N2%: N2% = 0
'   '
'   If PRIDIF& >= 0 Then
'     TOPIN& = PRIDIF& - 15
'       If TOPIN& < 0 Then TOPIN& = 0
'     On Error Resume Next
'     COMPAFU.List1.ListIndex = PRIDIF&
'     COMPAFU.List2.ListIndex = PRIDIF&
'     COMPAFU.List1.TopIndex = TOPIN&
'     COMPAFU.List2.TopIndex = TOPIN&
'     On Error GoTo 0
'   End If
'   '
   COMPAFU.Show 1
     
End Sub
