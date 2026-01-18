VERSION 5.00
Begin VB.Form INITAB02 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "X-Reports"
   ClientHeight    =   780
   ClientLeft      =   45
   ClientTop       =   10305
   ClientWidth     =   4455
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   780
   ScaleWidth      =   4455
   Begin VB.Timer Timer3 
      Interval        =   2000
      Left            =   945
      Top             =   0
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   525
      Top             =   0
   End
   Begin VB.Timer Timer1 
      Interval        =   200
      Left            =   0
      Top             =   0
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1260
      TabIndex        =   0
      Top             =   250
      Width           =   3060
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Comando:"
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
      Left            =   0
      TabIndex        =   1
      Top             =   315
      Width           =   1170
   End
End
Attribute VB_Name = "INITAB02"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim TOCERRADO%
Dim INIESPERA

Private Sub Form_Activate()
   '
   Static TIMA
   '
   While Timer < TIMA + 0.4: Wend
   TIMA = Timer
   '
End Sub

'
Private Sub Form_Load()
    '
    If Top < Screen.Height Then
      Top = Screen.Height
      DoEvents
    End If
    '
    RXSTRI$ = LOADFILE$("C:\FLEXOFT\FLEXRAM.DRV")
    LUBAS$ = TRIM$(Mid$(RXSTRI$, 145, 46))
      If Mid$(LUBAS$, 2, 1) <> ":" Then
        LUBAS$ = Left$(CurDir, 2) + LUBAS$
      End If
    ULOAC$ = Left$(CurDir, 2)
    USNBR% = CVI(Mid$(RXSTRI$, 9, 10))
    USERNAME$ = TRIM$(Mid$(RXSTRI$, 231, 30))
    MENUPRIN$ = TRIM$(Mid$(RXSTRI$, 13, 8))
    MODINVO$ = TRIM$(Mid$(RXSTRI$, 21, 8))
    DENOMOD$ = TRIM$(Mid$(RXSTRI$, 29, 32))
    MODMENU$ = TRIM$(Mid$(RXSTRI$, 61, 8))
    '
    EMPREAC$ = TRIM$(Mid$(RXSTRI$, 193, 30))
    LUDAT$ = TRIM$(Mid$(RXSTRI$, 99, 46))
    LUBAS$ = TRIM$(Mid$(RXSTRI$, 145, 46)): If Mid$(LUBAS$, 2, 1) <> ":" Then LUBAS$ = ULOAC$ + LUBAS$
    LUCOM$ = "C:\FLEXOFT\"
    LUBOOT$ = TRIM$(Mid$(RXSTRI$, 97, 1))
    '
    EJER% = Asc(Mid$(RXSTRI$, 69, 1))
      If EJER% <> 1 Then EJER% = (-1)
    COMEJ% = CVI(Mid$(RXSTRI$, 87, 2))
    FINEJ% = CVI(Mid$(RXSTRI$, 89, 2))
    EJU$ = EJACT$
    SBE$ = SUBEMP$
    '
    IDENTER$ = TRIM$(Mid$(RXSTRI$, 91, 6))
    NOMTER$ = TRIM$(Mid$(RXSTRI$, 261, 30))
    '
End Sub
'
Sub SELPROCE()
    '
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Cancel = 0
End Sub

Private Sub Text1_Change()
   '
   If Text1 = "" Then Exit Sub
   '
   Timer1.Enabled = False
   PROPRIN$ = ""
   ATRIBU$ = ""
   COMMA$ = TRIM$(Text1)
   '
   If UCase$(TRIM$(Text1)) = "CLOSETAB" Then
     Timer2.Enabled = False
     Call CIERRARCH("*.*")
     Kill "C:\FLEXOFT\RCONTROL.DRV"
     Close: End  ' Call FINAL("")
     Exit Sub
   End If
   '
   PPX% = InStr(COMMA$, "/")
   If PPX% < 1 Then PPX% = 1 + Len(COMMA$)
   PROPRIN$ = TRIM$(Left$(COMMA$, PPX% - 1))
   ATRIBU$ = TRIM$(Mid$(COMMA$, PPX% + 1))
   '
10 PROLLAMA$ = ""
   Timer2.Enabled = False
   If Left$(PROPRIN$, 1) = "*" Then
     PROPRIN$ = Mid$(PROPRIN$, 2)
   End If
   '
   FORMA$ = CurDir + "\" + PROPRIN$ + ".PRF"
   If EXISTE%(FORMA$) < 1 Then
      Call EXTRAPRF(FORMA$)
   End If
   '
   If EXISTE%(FORMA$) < 1 Then
      Call MENSERR(24, "Controlador no Encontrado:\" + PROPRIN$ + ".PRF")
      GoTo 99
   End If
   '
   XX$ = LOADFILE$(FORMA$)
   TIPRO% = CVI(Mid$(XX$, 127, 2))
   '
   Select Case TIPRO%
        '
        Case 1              ' SELECCIONA EDICION TABULADA
           XLISTAB.Label13 = ""
           XSORTAB.Label3 = ""
           XFORTAB.LITAB.Clear
           XFORTAB.Label1 = FORMA$
           YAFIL% = 0
           XFORTAB.ZOrder 0
           XFORTAB.Show 1
           GoTo 99
        Case 2              ' SELECCIONA LISTADO TABULADO
           XFORTAB.Label1 = ""
           XSORTAB.Label3 = ""
           XLISTAB.Label13 = ""
           XLISTAB.Label13 = FORMA$
           XLISTAB.ZOrder 0
           XLISTAB.Show 1
           GoTo 99
        Case 3              ' SELECCIONA FILTRO TABULADO
           XFORTAB.Label1 = ""
           XSORTAB.Label3 = ""
           XLISTAB.Label13 = ""
           'CCI$(0) = PROPRIN$
           'Call CONECRUN("#FILTABU\" + PROPRIN$, "")
           'Exit Sub
 '              WFILTAB.Show
 '              WFILTAB.Refresh
           GoTo 99
        Case 4              ' SELECCIONA ORDENAMIENTO
           XFORTAB.Label1 = ""
           XLISTAB.Label13 = ""
           XSORTAB.Label3 = FORMA$
           XSORTAB.ZOrder 0
           XSORTAB.Show 1
           GoTo 99
        Case 5              ' SELECCIONA PANTALLA PLANA
           XFORTAB.Label1 = ""
           XSORTAB.Label3 = ""
           XLISTAB.Label13 = ""
'              WPLATAB.Show
 '              WPLATAB.Refresh
           GoTo 99
   End Select
   '
   Call MENSERR(24, "No se Encuentra Controlador Tipo " + TRIM$(Str$(TIPRO%)) + "\Correspondiente a " + PROPRIN$ + ".PRF")
   '
99 Call CIERRARCH("*.*")
   Timer2.Enabled = False
   If PROLLAMA$ <> "" Then
     Text1 = ""
     Timer2.Enabled = True
     Exit Sub
   End If
   '
   If TRIM$(SALI$) <> "" Then
     If Left$(SALI$, 1) = "*" Then
       ATRIBU$ = ""
       PROPRIN$ = SALI$
       SALI$ = ""
       GoTo 10
     End If
   End If
   '
   Call CIERRARCH("*.*")
   Kill "C:\FLEXOFT\RCONTROL.DRV"
   Timer1.Enabled = True
   INIESPERA = Timer
   Text1 = ""
   DoEvents
   '
End Sub

Private Sub Timer1_Timer()
   Static FIDACTRL#
   '
   If TOCERRADO% <> 1 Then
     Call CIERRARCH("*.*")
     TOCERRADO% = 1
   End If
   '
   If INIESPERA > 0 Then
     If Timer - INIESPERA > 600 Then ' MANTIENE 10 MINUTOS
       Call CIERRARCH("*.*")
       Call FINAL("")
     End If
   End If
   '
   On Error Resume Next
   FDT# = FileDateTime("C:\FLEXOFT\RCONTROL.DRV")
   If FDT# > 0 Then
     If FDT# <> FIDACTRL# Then
       TOCERRADO% = 0
       FIDACTRL# = FDT#
       Text1 = LOADFILE$("C:\FLEXOFT\RCONTROL.DRV")
       On Error Resume Next
       LOCONTROL = FileLen("C:\FLEXOFT\RCONTROL.DRV")
       On Error GoTo 0
       If LOCONTROL > 0 Then
         Timer1.Enabled = False
       End If
     End If
   End If
   '
End Sub

Private Sub Timer2_Timer()
   '
   If PROLLAMA$ = "" Then GoTo 21
   '
   On Error GoTo 20
   AppActivate PROLLAMA$, True
   On Error GoTo 0
   Exit Sub
   '
20 Resume 21
21 On Error GoTo 0
   '
   PROLLAMA$ = ""
   If TRIM$(SALI$) <> "" Then
     If Left$(SALI$, 1) = "*" Then
       Timer2.Enabled = False
       Text1 = SALI$
       SALI$ = ""
       Exit Sub
     End If
   End If
   '
   Timer2.Enabled = False
   Call CIERRARCH("*.*")
   Kill "C:\FLEXOFT\RCONTROL.DRV"
   Timer1.Enabled = True
   INIESPERA = Timer
   Text1 = ""
   DoEvents
   '
End Sub

Private Sub Timer3_Timer()
   If PROLLAMA$ = "" Then
     If EXISTE%("C:\FLEXOFT\RCONTROL.DRV") > 0 Then
       On Error Resume Next
       AppActivate "X-REPORTS"
       DoEvents
       On Error GoTo 0
     End If
   End If
End Sub
