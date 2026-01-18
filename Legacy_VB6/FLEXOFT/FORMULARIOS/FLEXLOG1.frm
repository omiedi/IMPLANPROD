VERSION 5.00
Begin VB.Form FLEXLOG1 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0FFC0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Sistema Preconfigurado FLEXOFT (r)"
   ClientHeight    =   2910
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   4770
   Icon            =   "FLEXLOG1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2910
   ScaleWidth      =   4770
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.FileListBox File2 
      Height          =   870
      Left            =   105
      Pattern         =   "*.EXE"
      TabIndex        =   21
      Top             =   5565
      Visible         =   0   'False
      Width           =   4530
   End
   Begin VB.ListBox List1 
      Height          =   2205
      Left            =   105
      TabIndex        =   19
      Top             =   3150
      Width           =   4530
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Acceso al Sistema"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2640
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3585
      Begin VB.TextBox Text3 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2205
         TabIndex        =   2
         Top             =   2020
         Width           =   1170
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2205
         TabIndex        =   1
         Top             =   600
         Width           =   1170
      End
      Begin VB.Label Label1 
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
         Height          =   360
         Left            =   210
         TabIndex        =   6
         Top             =   1155
         Width           =   3165
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Clave de"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1350
         TabIndex        =   5
         Top             =   1995
         Width           =   765
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Usuario:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1395
         TabIndex        =   4
         Top             =   735
         Width           =   720
      End
      Begin VB.Image Image1 
         Height          =   585
         Left            =   315
         Picture         =   "FLEXLOG1.frx":030A
         Stretch         =   -1  'True
         Top             =   420
         Width           =   690
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Acceso:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1425
         TabIndex        =   3
         Top             =   2205
         Width           =   705
      End
      Begin VB.Image Image2 
         Height          =   525
         Left            =   210
         Picture         =   "FLEXLOG1.frx":0614
         Stretch         =   -1  'True
         Top             =   1890
         Width           =   585
      End
      Begin VB.Image Image4 
         Height          =   525
         Left            =   210
         Picture         =   "FLEXLOG1.frx":091E
         Stretch         =   -1  'True
         Top             =   1890
         Visible         =   0   'False
         Width           =   630
      End
      Begin VB.Image Image5 
         Height          =   540
         Left            =   210
         Picture         =   "FLEXLOG1.frx":0C28
         Stretch         =   -1  'True
         Top             =   1890
         Visible         =   0   'False
         Width           =   645
      End
   End
   Begin VB.FileListBox File1 
      Height          =   675
      Left            =   2520
      TabIndex        =   16
      Top             =   0
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00004000&
      Height          =   3200
      Left            =   3840
      ScaleHeight     =   3135
      ScaleWidth      =   1035
      TabIndex        =   7
      Top             =   -100
      Width           =   1095
      Begin VB.CommandButton Command4 
         Caption         =   "Command4"
         Height          =   225
         Left            =   210
         TabIndex        =   20
         Top             =   1680
         Visible         =   0   'False
         Width           =   435
      End
      Begin VB.PictureBox MARCOBARRA 
         BackColor       =   &H000000FF&
         Height          =   135
         Left            =   120
         ScaleHeight     =   75
         ScaleWidth      =   585
         TabIndex        =   14
         Top             =   2160
         Width           =   650
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H00FF0000&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   0
            TabIndex        =   15
            Top             =   0
            Width           =   12000
         End
      End
      Begin VB.CommandButton Command2 
         Caption         =   "O.K."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   650
         Left            =   120
         Picture         =   "FLEXLOG1.frx":0F32
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Ingresa al Sistema"
         Top             =   960
         Width           =   650
      End
      Begin VB.CommandButton command1 
         Caption         =   "Exit"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   650
         Left            =   120
         Picture         =   "FLEXLOG1.frx":123C
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Cierra - Abandona la Aplicación"
         Top             =   290
         Width           =   650
      End
      Begin VB.CommandButton BOTREC 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   650
         Left            =   120
         Picture         =   "FLEXLOG1.frx":1386
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Aprueba la Selección de Items Cumplidos"
         Top             =   4755
         Width           =   650
      End
      Begin VB.PictureBox Picture1 
         BackColor       =   &H000000FF&
         Height          =   135
         Left            =   120
         ScaleHeight     =   75
         ScaleWidth      =   585
         TabIndex        =   9
         Top             =   3720
         Width           =   650
         Begin VB.Frame Frame2 
            BackColor       =   &H00FF0000&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   0
            TabIndex        =   10
            Top             =   0
            Width           =   12000
         End
      End
      Begin VB.CommandButton Command3 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   650
         Left            =   120
         Picture         =   "FLEXLOG1.frx":1690
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Refresca - Invierte Selección"
         Top             =   4080
         Width           =   650
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   330
         Left            =   105
         TabIndex        =   17
         Top             =   1785
         Width           =   645
      End
      Begin VB.Image Image6 
         Height          =   390
         Left            =   -300
         Picture         =   "FLEXLOG1.frx":17DA
         ToolTipText     =   "Doble-Click para (Re)-Habilitar Estación de Trabajo"
         Top             =   2400
         Width           =   1515
      End
   End
   Begin VB.Label Label99 
      Caption         =   "Microsoft Vísual Basic 5.O ................................"
      Height          =   225
      Left            =   525
      TabIndex        =   18
      Top             =   315
      Visible         =   0   'False
      Width           =   2745
   End
End
Attribute VB_Name = "FLEXLOG1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim FF0 As String * 128
Dim CMD$

Dim CLA$, CLB$, CLAV$
Dim CLAVE$, CONTRA$
Dim QWERTY%
Dim USNBR%, NUSER%
Dim USERNA$
Dim EMPREAC$
Dim MULTIEM%
Dim YACAR%
'
Private Sub command1_Click()
   Close: End
End Sub

Private Sub Command3_Click()
TTXX$ = ""
ULO$ = Left$(CurDir, 2)
TTXX$ = Dir(ULO$ + "\*.", 16)
10 TTYY$ = Dir
If TTYY$ <> "" Then
  TTXX$ = TTXX$ + "; " + TTYY$
  GoTo 10
End If
MsgBox TTXX$
End Sub

Private Sub Command2_Click()
   If Len(Text3) > 0 Then
       Call Text3_KeyPress(13)
     Else
       Call Text1_KeyPress(13)
   End If
End Sub

Private Sub Command5_Click()
    AA123 = LOADFILE$("IDENSYS.DRV")
    For KU& = 961 To Len(AA123) Step 96
      TTXX$ = Mid$(AA123, KU&, 96)
      TTX1$ = DESENCRI$(Mid$(TTXX$, 8, 12))
      TTX2$ = DESENCRI$(Mid$(TTXX$, 38, 38))
      MsgBox TTX1$ + "   " + TTX2$
    Next KU&
    '
'
'    SUCAR% = 0
'    For KI% = 1 To 94
'       SUCAR% = SUCAR% + Asc(Mid$(TTXX$, KI%, 1)) - 83
'    Next KI%
'    IDS$ = TTXX$ + MKI$(SUCAR%)
'    Put #N2%, REI%, IDS$
'    Close #N2%
'    '
End Sub

'
Private Sub Form_Load()
    '
    Text1.Enabled = False
    Text3.Enabled = False
    Show
    Refresh
    DoEvents
    '
    On Error Resume Next
    Load CTRLSHOW
    If Err Then
      On Error GoTo 0
      Call INSTALACON
    End If
    On Error GoTo 0
    '
    Call ABRELOGON      ' DETERMINA SI ESUNDEMO% (1/0)
                        ' GENERA UN JOBID$
                        ' DETERMINA CARPETA LOCAL LUCOM$
    List1.AddItem "Es Un Demo = " & ESUNDEMO%
    List1.AddItem "JOBID = " + JOBID$
    List1.AddItem "C.Local = " + LUCOM$
    '
    Label2 = 10: DoEvents
    If ESUNDEMO% = 1 Then
      RUSYS$ = "C:\WINDOWS\SYSTEM"
      If EXISTE%(RUSYS$ + "\.") < 1 Then
        RUSYS$ = "C:\WINDOWS\SYSTEM32"
        If EXISTE%(RUSYS$ + "\.") < 1 Then
          RUSYS$ = "C:\WINNT\SYSTEM"
          If EXISTE%(RUSYS$ + "\.") < 1 Then
            RUSYS$ = "C:\WINNT\SYSTEM32"
            If EXISTE%(RUSYS$ + "\.") < 1 Then
              MsgBox "Imposible Ejecutar en este Equipo (1). Falta Carpeta Sistema Operativo Windows."
51            GoTo 51
            End If
          End If
        End If
      End If
      '
      Label2 = 20: DoEvents
      TTXX$ = RANDSTRING$(128)
      Call SAVEFILE(RUSYS$ + "\NULLPRUE.TMP", TTXX$)
      TTYY$ = LOADFILE$(RUSYS$ + "\NULLPRUE.TMP")
      If TTYY$ <> TTXX$ Then
         MsgBox "Imposible Ejecutar en este Equipo (2). Establezca Derechos de Crear Archivos, Leer y Escribir en Carpeta '" + RUSYS$ '."
52       GoTo 52
      End If
    End If
    '
    Label2 = 30: DoEvents
    CMD$ = Command$
    '
    Screen.MousePointer = 11
    '
'1- Identificar Unidad de Red Actual

    ChDir App.Path
' MsgBox App.Path & Chr$(13) + Chr$(10) & CurDir
    If Mid$(App.Path, 2, 1) <> ":" Or CurDir$ <> App.Path Then
      PPXX% = 3
53    PPYY% = InStr(PPXX% + 1, App.Path, "\")
      If PPYY% > 0 Then
        PPXX% = PPYY%
        GoTo 53
      End If
      RUTAX$ = Left$(App.Path, PPXX% - 1)
      MsgBox "Imposible Ejecutar - Asigne Unidad de Red a '" + RUTAX$ + "'."
      Close: End
    End If
    Label2 = 40: DoEvents
    '
'2- Disco Rigido Local 'C:'
    Screen.MousePointer = 11
    Call BUSCADISCO_C
    Label2 = 50: DoEvents
    '
'3- Establecer Version Demo SI=1 / NO=0
    Screen.MousePointer = 11
    VERDEFLOG% = 0
    If ESUNDEMO% = 1 Then
      VERDEFLOG% = 1   '
      TTXX$ = "El Presente es un Software de Demostración" + Chr$(13) + Chr$(10)
      TTXX$ = TTXX$ + "Plenamente Operativo por Tiempo Limitado" + Chr$(13) + Chr$(10)
      TTXX$ = TTXX$ + "Destinado Unicamente a Evaluar la Adquisi-" + Chr$(13) + Chr$(10)
      TTXX$ = TTXX$ + "ción y Licenciamiento Definitivos. Derechos" + Chr$(13) + Chr$(10)
      TTXX$ = TTXX$ + "Reservados Ley 11723."
      MsgBox TTXX$
    End If
    'Call VERSION_DEMO    ' por escritura / borrado en carpeta actual
    Label2 = 60: DoEvents
    '
'4- Carpeta Local 'C:\FLEXOFT'
    Screen.MousePointer = 11
    Call CARPETALOCAL    ' Verifica / Crea Carpeta C:\FLEXOFT
    Label2 = 70: DoEvents
    '
'5- Carpeta Sistema Basico
    Screen.MousePointer = 11
    Call CARPETASIBAS    ' Verifica / Crea Carpeta Sistema Basico
    Label2 = 80: DoEvents
    '
'6- Identificacion del Sistema
    Screen.MousePointer = 11
    Call CARIDENSYS
    Label2 = 90: DoEvents
    '
'7- Habilitacion de Terminal
    Screen.MousePointer = 11
    Call VERITER
    Label2 = 100: DoEvents
    '
'7- Unidad Logica de Datos
'   Call UNILOGIDATOS    ' Unidad Logica de Datos
    '
    If VERDEFLOG% > 0 Then
      XK1% = EXISTE%("C:\WINDOWS\SYSTEM\C0MCTL32.0CX")
      XK1% = XK1% + EXISTE%("C:\WINDOWS\SYSTEM32\C0MCTL32.0CX")
      XK1% = XK1% + EXISTE%("C:\WINNT\SYSTEM\C0MCTL32.0CX")
      XK1% = XK1% + EXISTE%("C:\WINNT\SYSTEM32\C0MCTL32.0CX")
      If XK1% > 0 Then
        TTXX$ = "Plazo de  Validez  de  Demostración Vencido" + Chr$(13) + Chr$(10)
        TTXX$ = TTXX$ + Chr$(13) + Chr$(10)
        TTXX$ = TTXX$ + "Para Obtener su Licencia de Uso, Contacte a" + Chr$(13) + Chr$(10)
        TTXX$ = TTXX$ + "FLEXOFT  Soluciones Informaticas Integrales" + Chr$(13) + Chr$(10)
        TTXX$ = TTXX$ + "info@flexoft.com.ar - Tel:  54-11-4524-1284" + Chr$(13) + Chr$(10)
        TTXX$ = TTXX$ + "              Buenos Aires - Argentina"
        MsgBox TTXX$
        Close: End
        Exit Sub
      End If
      Label2 = 120: DoEvents
    End If
    '
'8- Carpeta de Datos
260 Screen.MousePointer = 11
    Call CARPETADATOS    ' Verifica / Crea Carpeta de Datos
    Label2 = 150: DoEvents
'Call CARIDENTER
    '
    Screen.MousePointer = 11
    If Mid$(LUDAT$, 2, 1) <> ":" Then
      LUDAT$ = ULODA$ + LUDAT$
    End If
    LU$ = LUDAT$
    Label2 = 160: DoEvents
    '
'    VERDEFLOG% = 0
''GoTo 280
'    If VERDEFLOG% < 1 Then
'      If CONTROVENCIM%("FLSYSTEM") < 1 Then
'        Text1.Enabled = False
'        Text3.Enabled = False
'        command1.Enabled = False
'        Command2.Enabled = False
'        Cancel = 1
'        Exit Sub
'      End If
'    End If
280 Label2 = 170: DoEvents
    '
    If VERDEFLOG% < 1 Then
      Call UPDATES
    End If
    Label2 = 180: DoEvents
    '
    Screen.MousePointer = 1
    Text1.Enabled = True
    Text3.Enabled = True
    CLA$ = "": CLB$ = "": CLAVE$ = ""
    Label2 = 999: DoEvents
    Text1 = "": Text3 = ""
    Text1.SetFocus
    YACAR% = 1
    '
    Call VERIBAK
    '
    Exit Sub
    '
500 Dim RAMSTRI As String * 512
    RAMSTRI$ = String$(512, 0)
    '
    If VERDEFLOG% = 1 Then
       USNBR% = 101
       USERNA$ = "Usuario DEMO"
       MENUPRIN$ = ""
       MODINVO$ = CMD$
       EJER% = 1
       COMEJ% = 0
       FINEJ% = 0
       EMPREAC$ = "Su Empresa"
       LUBOOT$ = "C"
       IDENTER$ = "000000"
       NOMTER$ = "Estacion Local"
    End If
    '
End Sub
'
Sub BUSCADISCO_C()
   '
   ARXX$ = "C:\"
   BRXX$ = CurDir
   On Error GoTo 6
   ChDir ARXX$
   ChDir BRXX$
   On Error GoTo 0
   GoTo 99                ' exit sub
   '
6  Resume 7
7  On Error GoTo 0
   Screen.MousePointer = 1
   MsgBox "Imposible Ejecutar - No hay Disco Rígido Local 'C:'."
   Close: End
   '
99 End Sub

Sub VERSION_DEMO()
   '
10 APRUEBA$ = CurDir + "\NULLPRUE.$$$"
   LOPRUE& = 0
   On Error Resume Next
   LOPRUE& = FileLen(APRUEBA$)
   On Error GoTo 0
   'If LOPRUE& > 0 Then GoTo 10
   '
   ' Carpeta Actual - Crear archivo
   Dim RPRU As String * 128
   NX% = FreeFile
   On Error GoTo 21
   Open APRUEBA$ For Random Access Read Write Shared As #NX% Len = 128
   RPRU$ = String$(128, "*")
   Put #NX%, 1, RPRU$
   Close #NX%: NX% = 0
   On Error GoTo 0
   GoTo 30
   '
21 Resume 22
22 On Error GoTo 0
   Close #NX%: NX% = 0
   VERDEFLOG% = 1
   GoTo 99
   '
30 On Error GoTo 31
   Kill APRUEBA$
   On Error GoTo 0
   GoTo 99
   '
31 Resume 32
32 On Error GoTo 0
   VERDEFLOG% = 1
   '
99 End Sub

Sub CARPETALOCAL()     ' Carpeta de Trabajo Local
   '
   VUELTA% = 1
41 ARXX$ = LUCOM$
   If Right$(ARXX$, 1) = "\" Then ARXX$ = Left$(ARXX$, Len(ARXX$) - 1)
   BRXX$ = App.Path
   On Error GoTo 43
   ChDir ARXX$
   ChDir BRXX$
   On Error GoTo 0
   '
   ' VERIFICAR AHORA SI SE PUEDE ESCRIBIR EN LA CARPETA C:\FLEXOFT
   NX% = FreeFile
   On Error Resume Next
   Open ARXX$ + "\RCONTROL.DRV" For Output As #NX%
   If Err Then
     On Error GoTo 0
     Close #NX%
     On Error GoTo 0
     If ESUNDEMO% = 1 Then GoTo 44
   End If
   Close #NX%
   On Error GoTo 0
   GoTo 99
   '
43 Resume 44
44 On Error GoTo 0
   '
   If VUELTA% = 2 Then
     Screen.MousePointer = 1
     MsgBox "Imposible Crear Carpeta '" + ARXX$ + "'. Consulte."
     Close: End
   End If
   '
   TITUMESS$ = "Sistemas FLEXOFT"
   If VERDEFLOG% = 1 Then TITUMESS$ = "Demostración Activa FLEXOFT"
   If MsgBox("FLEXOFT Debe Crear una Carpeta 'C:\FLEXOFT' en su Disco Rígido Local 'C:'.", 1, TITUMESS$) <> 1 Then
     Close: End
   End If
   '
   Screen.MousePointer = 11
   MkDir ARXX$
   VUELTA% = 2
   GoTo 41
   '
99 'Screen.MousePointer = 1
   '
End Sub

Sub CARPETASIBAS()    ' Carpeta Sistema Basico
   '
   If VERDEFLOG% > 0 Then
      LUBAS$ = Trim$(CurDir) + "\"
      GoTo 99
   End If
   '
   On Error Resume Next
   LOAR& = FileLen("FLEXCFG.DAT")
   On Error GoTo 0
   LB$ = ""
   If LOAR& >= 384 Then
      N1% = FreeFile
      Open "FLEXCFG.DAT" For Random Access Read Shared As #N1% Len = 128
      '
      Get #N1%, 3, FF0$
      LB$ = Trim$(Left$(FF0$, 46))
      Close #N1%: N1% = 0
   End If
   '
   If LB$ = "" Or Len(LB$) < 3 Then
      LB$ = Trim$(Mid$(UCase$(CurDir), 3))
   End If
   '
   If Len(LB$) > 0 Then
      If Right$(LB$, 1) <> "\" Then
         LB$ = LB$ + "\"
      End If
   End If
   LUBAS$ = LB$
   If EXISTE%(LUBAS$ + ".") < 1 Then ' SI NO EXISTE LA CARPETA ASUME LA ACTUAL
     LUBAS$ = App.Path + "\"
   End If
   If Mid$(LUBAS$, 2, 1) <> ":" Then
      ULOAC$ = Left$(CurDir, 2)
      LUBAS$ = ULOAC$ + LUBAS$
   End If
   '
99 Call VERIEXISTE(LUBAS$ + "FLEXPRFS.FLX")
   Call VERIEXISTE(LUBAS$ + "FLEXRFS.ZIP")
   Call VERIEXISTE(LUBAS$ + "FLEXPRF.ZIP")
   Call VERIEXISTE(LUBAS$ + "OBJNAME.DRV")
   Call VERIEXISTE(LUBAS$ + "OBJDETA.DRV")
   '
   Call VERIEXISTE(LUBAS$ + "FLEXPIF.PIF")
   Call VERIEXISTE(LUBAS$ + "DOSCOMM.PIF")
   Call VERIEXISTE(LUBAS$ + "DOSCOMM.EXE")
   Call VERIEXISTE(LUBAS$ + "FLEXQB1.EXE")
   Call VERIEXISTE(LUBAS$ + "FLEXQB2.EXE")
   '
   Call VERIEXISTE(LUBAS$ + "PKZIP.EXE")
   Call VERIEXISTE(LUBAS$ + "PKUNZIP.EXE")
   Call VERIEXISTE(LUBAS$ + "ARJ.EXE")
   Call VERIEXISTE(LUBAS$ + "BRUN45.EXE")
   '
   Call TRAPIFLEX("FLEXPIF.PIF")
   Call TRAPIFLEX("DOSCOMM.PIF")
   Call TRAPIFLEX("DOSCOMM.EXE")
   Call TRAPIFLEX("FLEXQB1.EXE")
   Call TRAPIFLEX("FLEXQB2.EXE")
   Call TRAPIFLEX("BRUN45.EXE")
   '
End Sub

Sub VERIEXISTE(ARCHXX$)
   '
   If EXISTE%(ARCHXX$) < 1 Then
      MsgBox "Error: Falta '" + ARCHXX$ + "' en Carpeta " + LUBAS$
      Close: End
   End If
   '
End Sub
'
Sub TRAPIFLEX(ARCHXX$)
   '
   LCB$ = LUBAS$
   If EXISTE%(LUCOM$ + ARCHXX$) <> 1 Then
     On Error Resume Next
     FileCopy LCB$ + ARCHXX$, LUCOM$ + ARCHXX$
     On Error GoTo 0
     '
     If EXISTE%(LUCOM$ + ARCHXX$) < 1 Then
       MsgBox "Error: Imposible Instalar '" + ARCHXX$ + "' en Carpeta '" + LUCOM$ + "'"
      Close: End
     End If
     '
   End If
   '
End Sub
'
Sub CARPETADATOS()    ' Carpeta de datos
   '
   MULTIEM% = 0
   If VERDEFLOG% = 1 Then
     LD$ = "C:\FLEXOFT\DATOS"
     GoTo 60
   End If
   '
20 On Error GoTo 0
   If EXISTE%("FLEXMUL.DRV") = 1 Then
     SELUDAT1.List2.Clear
     SELUDAT1.List2.Visible = True
     '
     NX% = FILEOPEN("FLEXMUL.DRV", 0, 79)
     Dim FMX As String * 79
     '
     For I% = 1 To LOF(NX%) / 79
       Get #NX%, I%, FMX$
       SUCAR% = 0
       For J% = 1 To 77
         SUCAR% = SUCAR% + Asc(Mid$(FMX$, J%, 1))
       Next J%
       If CVI(Mid$(FMX$, 78, 2)) = SUCAR% Then
         RASO$ = DESENCRI$(Left$(FMX$, 30))
         PACCE$ = DESENCRI$(Mid$(FMX$, 32, 46))
         If EXISTE%(Trim$(PACCE$) + "\.") > 0 Then
           Z$ = Space$(96)
           Call REPLA(Z$, RASO$, 1, 30)
           Call REPLA(Z$, PACCE$, 51, 46)
           SELUDAT1.List2.AddItem Z$
         End If
       End If
     Next I%
     '
     Close #NX%
     If SELUDAT1.List2.ListCount > 0 Then
       SELUDAT1.List2.ListIndex = 0
       MULTIEM% = 1
       Screen.MousePointer = 1
       If SELUDAT1.List2.ListCount = 1 Then
         Call SELUDAT1.List2_DblClick
         Exit Sub
       End If
       SELUDAT1.Show 1
       Screen.MousePointer = 11
     End If
   End If
   '
   TCONFI$ = LOADFILE$(CurDir + "\" + "FLEXCFG.DAT")
   LD$ = Trim$(Mid$(TCONFI$, 51, 45))
' AQUI ESTA BIEN
   CUDI% = 0
   '
   If LD$ = "" Then
      LD$ = Mid$(CurDir, 3) + "\"
      Call SAVEFILE(LUCOM$ + "ULOGDAT.DRV", Left$(CurDir$, 2))
   End If
   '
   IDTX$ = Trim$(IDENTER$)
   If IDTX$ = "" Then
      NN11% = FILEOPEN%(LUCOM$ + "IDENTER.DRV", 0, 40)
      Dim X1 As String * 40
      Get #NN11%, 1, X1$
      J1$ = DESENCRI$(Left$(X1$, 38))
      IDENTER$ = Trim$(Mid$(J1$, 33, 6))
      IDTX$ = IDENTER$
      Close #NN11%
   End If
   '
' AQUI ESTA BIEN    MsgBox 45 & "   " & LD$
45 'DITRAX$ = Left$(App.Path, 2)
   'MsgBox DITRAX$
   'GoTo 46
If LD$ <> "" Then
If Mid$(LD$, 2, 1) = ":" Then
GoTo 60
End If
End If

   ULXX$ = LOADFILE$("ULODATOS.DRV")
   For KKI% = 1 To Len(ULXX$) Step 16
      ULYY$ = Mid$(ULXX$, KKI%, 16)
      If Trim$(Left$(ULYY$, 14)) = IDTX$ Then
        If Mid$(ULYY$, 16, 1) = ":" Then
          DITRAX$ = Mid$(ULYY$, 15, 2)
46        NN11% = FreeFile
          Open LUCOM$ + "ULOGDAT.DRV" For Output As #NN11%
          Print #NN11%, DITRAX$
          Close #NN11%: NN11% = 0
          GoTo 50
        End If
      End If
   Next KKI%
   ' si no hay puntero a U.L. de datos, poner por default directorio actual
   If VERDEFLOG% < 1 Then
     ULXX$ = ULXX$ + Left$(IDTX$ + Space$(14), 14) + UCase$(Left$(CurDir$, 2))
     Call SAVEFILE("ULODATOS.DRV", ULXX$)
     GoTo 45
   End If
   '
50 If Mid$(LD$, 2, 1) = ":" Then
      LD$ = Mid$(LD$, 3)
   End If
   LD$ = Trim$(Left$(LOADFILE$(LUCOM$ + "ULOGDAT.DRV"), 2)) + LD$
   '
60 CREACARPE% = 0: VUELTA% = 1
   If LD$ = "" Then
      CREACARPE% = 1
      LD$ = CurDir + "\DATOS"
      If VERDEFLOG% = 1 Then
        LD$ = "C:\FLEXOFT\DATOS"
      End If
   End If
   If Right$(LD$, 1) <> "\" Then
     LD$ = LD$ + "\"
   End If
   '
   BRXX$ = CurDir
   ARXX$ = Trim$(LD$)
   ARXX$ = Left$(ARXX$, Len(ARXX$) - 1)
   If Mid$(ARXX$, 2, 1) <> ":" Then
     ULOK$ = Trim$(Left$(LOADFILE$(LUCOM$ + "ULOGDAT.DRV"), 2))
     If Mid$(ULOK$ + "  ", 2, 1) <> ":" Then
       ULOK$ = Left$(CurDir$, 2)
     End If
     ARXX$ = ULOK$ + ARXX$
   End If
   '
63 On Error GoTo 61
   ChDir ARXX$
   ChDir BRXX$
   On Error GoTo 0
   '
   If CREACARPE% = 1 Then
     Screen.MousePointer = 1
     If VERDEFLOG% = 1 Then GoTo 631
     '
     If MsgBox("Instalar Set de Datos de Prueba", 36) = 6 Then
631    While GENDAPRUE.Visible = True: Wend
       GENDAPRUE.Label4 = ARXX$
       GENDAPRUE.Label18.Caption = "Datos de Prueba"
       Screen.MousePointer = 11
       GENDAPRUE.Show 1
     End If
   End If
   
   GoTo 70
   '
61 Resume 62
62 On Error GoTo 0
   ChDir BRXX$             ' VUELVE A DIRECTORIO ACTUAL
   'Screen.MousePointer = 1
   If VUELTA% > 1 Then GoTo 68
   '
   CURDIRX$ = UCase$(CurDir$)
   If UCase$(Mid$(CurDir$, 3)) = "\FLEXOFT\WINDOWS" Then
     CURDIRX$ = Left$(CurDir, 2) + "\FLEXOFT"
   End If
   CARPROPUES$ = CURDIRX$ + "\DATOS"
   '
   If VERDEFLOG% = 1 Then
     CARPROPUES$ = "C:\FLEXOFT\DATOS"
   End If
   '
   If CARPROPUES$ <> ARXX$ Then
     ARXX$ = CARPROPUES$
     CREACARPE% = 1
     GoTo 63
   End If
   '
   Screen.MousePointer = 1
   TITUMESS$ = "FLEXOFT - Carpeta de Datos"
   If VERDEFLOG% = 1 Then TITUMESS$ = "Demostración Activa FLEXOFT"
   ARXX$ = InputBox$("FLEXOFT Debe Ahora Crear una Carpeta de Datos en su Disco Rígido '" + Left$(CurDir, 2) + "'.", TITUMESS$, CARPROPUES$, 5000, 4700)
   If ARXX$ = "" Then
     Close: End
   End If
   If Len(ARXX$) > 45 Then GoTo 68
   '
   Screen.MousePointer = 11
66 On Error GoTo 67
   MkDir ARXX$
   CREACARPE% = 1: VUELTA% = 2
   On Error GoTo 0
   GoTo 63
    '
67 ERNU = Err
   Resume 68
68 On Error GoTo 0
   Screen.MousePointer = 1
   MsgBox "Imposible Crear la Carpeta '" + ARXX$ + "'."
   Close: End
   '
70 If EXISTE%("FLEXCFG.DAT") < 1 Then CREACARPE% = 1
   If CREACARPE% = 1 Then
     If VERDEFLOG% = 0 Then
       If Mid$(ARXX$, 2, 1) = ":" Then
          ARXX$ = Mid$(ARXX$, 3)
       End If
       If Right$(ARXX$, 1) <> "\" Then
          ARXX$ = ARXX$ + "\"
       End If
       LUDAT$ = ARXX$
       N1% = FreeFile
       Open "FLEXCFG.DAT" For Random Access Read Write Shared As #N1% Len = 128
       Get #N1%, 1, FF0$
       ZZ0$ = FF0$
       Mid$(ZZ0$, 51, 45) = Left$(ARXX$ + Space$(45), 45)
       FF0$ = ZZ0$
       Put #N1%, 1, FF0$
       Close #N1%: N1% = 0
     End If
   End If
   '
80 If VERDEFLOG% = 0 Then
       TTXX$ = LOADFILE$("FLEXCFG.DAT")
       LUDAT$ = Trim$(Mid$(TTXX$, 51, 45))
       If Right$(LUDAT$, 1) <> "\" Then
          LUDAT$ = LUDAT$ + "\"
          Mid$(TTXX$, 51, 45) = Left$(LUDAT$ + Space$(45), 45)
          Call SAVEFILE("FLEXCFG.DAT", TTXX$)
       End If
     Else
       LUDAT$ = "\FLEXOFT\DATOS\"
   End If
   If Mid$(LUDAT$, 2, 1) = ":" Then
     GoTo 99
     ' LUDAT$ = Mid$(LUDAT$, 3)
   End If
   Call UNILOGIDATOS
   LUDAT$ = ULODA$ + LUDAT$
   '
99 End Sub

Sub UNILOGIDATOS()
   '
   ULO$ = Left$(Trim$(LOADFILE$(LUCOM$ + "ULOGDAT.DRV")), 2)
   If VERDEFLOG% = 1 Then ULO$ = ""
   If ULO$ = "" Then
     ULO$ = Left$(CurDir$, 2)
     If VERDEFLOG% > 0 Then ULO$ = "C:"
     Call SAVEFILE(LUCOM$ + "ULOGDAT.DRV", ULO$)
   End If
   '
   ULODA$ = ULO$
   '
End Sub

Sub CARIDENSYS()
   '
   IDST$ = LOADFILE$("IDENSYS.DRV")
   If Len(IDST$) < 96 * 3 Then
     IDST$ = "": GoTo 50
   End If
   '
   For I% = 1 To 3
     OFDIVE% = 83: If I% = 1 Then OFDIVE% = 71
     IDS$ = Mid$(IDST$, 96 * (I% - 1) + 1, 96)
     SUCAR% = 0: For KI% = 1 To 94
                   SUCAR% = SUCAR% + Asc(Mid$(IDS$, KI%, 1)) - OFDIVE%
                 Next KI%
     If CVI(Mid$(IDS$, 95, 2)) <> SUCAR% Then
       Call CUELGAEJEC(I%)
       Exit Sub
     End If
   Next I%
   '
   IDS$ = Mid$(IDST$, 193, 94)
   CD$ = Left$(CurDir$, 2)
   DoEvents
   SEVOLF$ = Left$(SERIVOL$(CD$) + Space$(12), 12)
GoTo 45 ' queda para revisar mas adelante
'If TRIM$(SEVOLF$) = "0E44-11CF" Or TRIM$(SEVOLF$) = "081D-0BD9" Or TRIM$(SEVOLF$) = "0B40-19F9" Then GoTo 45
   If Left$(DESENCRI$(IDS$), 12) <> SEVOLF$ Then
     Call CUELGAEJEC(4)
   End If
   '
45 Exit Sub
   '
   '
50 IDS$ = ENCRIPTA$(Space$(94))
   SUCAR% = 0: For KI% = 1 To 94
                 SUCAR% = SUCAR% + Asc(Mid$(IDS$, KI%, 1)) - 71
               Next KI%
   IDT$ = IDS$ + MKI$(SUCAR%)
   '
   IDS$ = ENCRIPTA$(Space$(94))
   SUCAR% = 0: For KI% = 1 To 94
                 SUCAR% = SUCAR% + Asc(Mid$(IDS$, KI%, 1)) - 83
               Next KI%
   IDT$ = IDT$ + IDS$ + MKI$(SUCAR%)
   '
   IDX$ = RANDSTRING$(94)
   CD$ = Left$(CurDir$, 2)
   DoEvents
   SEVOLF$ = Left$(SERIVOL$(CD$) + Space$(12), 12)
   Call REPLA(IDX$, SEVOLF$, 1, 12)
   IDS$ = ENCRIPTA$(IDX$)
   SUCAR% = 0: For KI% = 1 To 94
                 SUCAR% = SUCAR% + Asc(Mid$(IDS$, KI%, 1)) - 83
               Next KI%
   IDT$ = IDT$ + IDS$ + MKI$(SUCAR%)
   '
   Call SAVEFILE("IDENSYS.DRV", IDT$)
   '
End Sub
'
Sub CUELGAEJEC(KKK%)
   '
10 MsgBox "Imposible Continuar Ejecución (" + Trim$(Str$(KKK%)) + ")."
   End
   '
   GoTo 10
   '
End Sub
'
Sub CARIDENTER()
   '
   DoEvents
   SEVOLC$ = Left$(SERIVOL$("C:") + Space$(12), 12)
   SEVOLX$ = ENCRIPTA$(SEVOLC$)
   '
   IDXX$ = LOADFILE$(LUCOM$ + "IDENTER.DRV")
   If Len(IDXX$) < 80 Then Call FLINSTALL
   '
   IDYY$ = Mid$(IDXX$, 41, 40)
   DIVERI% = CVI(Mid$(IDYY$, 39, 2))
   SUDI% = 0: For KKI% = 1 To 38
                SUDI% = SUDI% + Asc(Mid$(IDYY$, KKI%, 1))
              Next KKI%
   If SUDI% <> DIVERI% Then Call FLINSTALL
   If Mid$(IDYY$, 11, 12) <> SEVOLX$ Then Call FLINSTALL
   '
   IDYY$ = Left$(IDXX$ + String$(40, 0), 40)
   IDENCRIP$ = Left$(IDYY$, 38)
   DIVERI% = CVI(Mid$(IDYY$, 39, 2))
   SUDI% = 0: For KKI% = 1 To 38
                SUDI% = SUDI% + Asc(Mid$(IDENCRIP$, KKI%, 1))
              Next KKI%
   If SUDI% <> DIVERI% Then Call FLINSTALL
   '
   J1$ = DESENCRI$(IDENCRIP$)
   IDENTER$ = Trim$(Mid$(J1$, 33, 6))
   If IDENTER$ = "" Then Call FLINSTALL
   NOMTER$ = Trim$(Mid$(J1$, 1, 32))
   If NOMTER$ = "" Then Call FLINSTALL
   '
   IDST$ = LOADFILE$("IDENSYS.DRV")
   For I% = 11 To Len(IDST$) / 96
     IDS$ = Mid$(IDST$, 96 * (I% - 1) + 1, 96)
     SUCAR% = 0: For KI% = 1 To 94
                   SUCAR% = SUCAR% + Asc(Mid$(IDS$, KI%, 1)) - 83
                 Next KI%
     If CVI(Mid$(IDS$, 95, 2)) = SUCAR% Then
       If Mid$(IDS$, 8, 12) = SEVOLX$ Then
         If Mid$(IDS$, 38, 38) = IDENCRIP$ Then GoTo 99
       End If
     End If
   Next I%
   Call FLINSTALL
   '
99 End Sub
'
Sub FLINSTALL()
End
End Sub

Private Sub Image3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Static TANTE
   If Button = 2 Then
     If Timer - TANTE > 0.5 Then
         TANTE = Timer
         Exit Sub
       Else
         CLAX$ = RANDSTRING$(-6)
         CONTRASE$ = InputBox$("Ingrese Contraseña", "Santo y Seña (" + CLAX$ + ")")
         If CONTRASE$ <> CONTRACLA$(CLAX$) Then Exit Sub
         '
         DEF$ = Left$(LOADFILE$("FLXCTRL.EXE"), 129)
         If Len(DEF$) < 129 Then Exit Sub
         '
         FLUTIL% = CVI(Mid$(DEF$, 37, 2))
         CACCAD% = CVI(Mid$(DEF$, 39, 2))
         CULOG$ = Left$(CurDir, 2)
         SEVOLF$ = Left$(SERIVOL$(CULOG$) + Space$(12), 12)
         CLAX$ = RANDSTRING$(-6)
         TTXX$ = Trim$(SEVOLF$) + " / " + FETEX$(FLUTIL%) + "-" + Trim$(Str$(CACCAD%)) + " / " + CLAX$
         CONTRA2$ = InputBox$("Ingrese Contra-Clave", "Contraclave: " + TTXX$)
         '
         ' LA CONTRACLAVE SE ARMA ASí:
         ' SEIS CARACTERES PARA LA CONTRACLAVE DEL RANSTRING AL FINAL DEL ENCABEZADO
         ' UN GUION O ESPACIO
         ' LA FECHA LIMITE EN FORMATO AAMMDD
         ' UN GUION O ESPACIO
         ' LA CONTRACLAVE DE LA FECHA LIMITE EN FORMATO ANTERIOR
         '
         If CONTRA2$ = "" Then Exit Sub
         FLINUE$ = Mid$(CONTRA2$, 12, 2) + "/" + Mid$(CONTRA2$, 10, 2) + "/" + Mid$(CONTRA2$, 8, 2)
         If Left$(CONTRA2$, 6) = CONTRACLA$(CLAX$) Then
           If Mid$(CONTRA2$, 15, 6) = CONTRACLA$(Mid$(CONTRA2$, 8, 6)) Then
             FLINUX% = FENUM%(FLINUE$)
             Call REPLA(DEF$, MKI$(FLINUX%), 37, 2)
             Call REPLA(DEF$, MKI$(0), 39, 2)
             DIVERI% = 0
             For UI% = 1 To 127
               DIVERI% = DIVERI% + Asc(Mid$(DEF$, UI%, 1)) - 123
             Next UI%
             Call REPLA(DEF$, MKI$(DIVERI%), 128, 2)
             '
             ABC$ = LOADFILE$("FLXCTRL.EXE")
             Call REPLA(ABC$, DEF$, 1, 129)
             Call SAVEFILE("FLXCTRL.EXE", ABC$)
             If ABC$ <> LOADFILE$("FLXCTRL.EXE") Then
               MsgBox "Acceso Denegado a Controlador"
               Exit Sub
             End If
             MsgBox "Aceptado"
             Exit Sub
           End If
         End If
         '
         MsgBox "Rechazado - Consulte"
         '
     End If
   End If
End Sub

Private Sub Image1_DblClick()
   Call BLOQSYS
End Sub

Private Sub Image6_DblClick()
  If MsgBox("Realmente Desea Re-Habilitar Esta Estación de Trabajo ?", 3) = 6 Then
    Call HABILITER
    Close: End
  End If
End Sub

Private Sub Image6_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If Button = 2 Then
    Height = 5840
  End If
End Sub

Private Sub Image6_MouseUP(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If Button = 2 Then
    Height = 3270
  End If
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
   '
   If YACAR% < 1 Then Exit Sub
   '
   A% = KeyAscii: KeyAscii = 0
   If A% = 13 Then
        Call BUSER
     ElseIf A% = 27 Then
        Close: End
     ElseIf A% < 48 Then
        Exit Sub
     Else
        If A% > 96 Then A% = A% - 32
        If Len(Text1.Text) < 12 Then
           CLB$ = CLB$ + "*"
           Text1.Text = CLB$
           CLA$ = CLA$ + Chr$(A%)
           CLAVE$ = Left$(CLA$ + Space$(12), 12)
        End If
   End If
   '
End Sub

Sub BUSER()
    '
    Static CAIN%
    CLA$ = Left$(UCase$(CLAVE$) + Space$(12), 12)
    '
    If Trim$(CLA$) = "???" Then
      CLAX$ = RANDSTRING$(-6)
      CONTRASE$ = InputBox$("Ingrese Contraseña", "Santo y Seña (" + CLAX$ + ")")
      If CONTRASE$ <> CONTRACLA$(CLAX$) Then GoTo 40
      '
      HOII% = HOY(HOS$): FLAB% = HOII% + 33
      DEF$ = Left$(LOADFILE$("FLXCTRL.EXE"), 129)
      If Len(DEF$) < 129 Then GoTo 20
      Call REPLA(DEF$, MKI$(FLAB%), 37, 2)
      Call REPLA(DEF$, MKI$(0), 39, 2)
      DIVERI% = 0: For UI% = 1 To 127
                     DIVERI% = DIVERI% + Asc(Mid$(DEF$, UI%, 1)) - 123
                   Next UI%
                   Call REPLA(DEF$, MKI$(DIVERI%), 128, 2)
      ABC$ = LOADFILE$("FLXCTRL.EXE")
      Call REPLA(ABC$, DEF$, 1, 129)
      Call SAVEFILE("FLXCTRL.EXE", ABC$)
      If ABC$ <> LOADFILE$("FLXCTRL.EXE") Then
        MsgBox "Acceso Denegado a Controlador"
      End If
      '
20    QWERTY% = 101: USNBR% = 101
      UserName$ = "Supervisor General"
      Call GRADATRAM
      Exit Sub
    End If
    '
30  If EXISTE%("USER.DAT") < 1 Then
31      QWERTY% = 101
        USNBR% = 101
        USERNA$ = "Supervisor General"
        Call GRADATRAM
        Exit Sub
    End If
    '
    VUSER$ = LOADFILE$("USER.DAT")
    USNBR% = 0
    USERNA$ = ""
    CONTRA$ = ""
    '
    CAUS% = Len(VUSER$) / 64
    If CAUS% < 1 Then GoTo 31
    MNGR% = 0
    For JJ% = 1 To CAUS%
      F0$ = Mid$(VUSER$, 64 * JJ% - 63, 64)
      IDENT$ = Mid$(F0$, 31, 12)
        If Left$(IDENT$, 1) = "@" Then IDENT$ = DESENCRI$(Mid$(F0$, 32, 11)) + " "
        If LTrim$(RTrim$(IDENT$)) = "" Then GoTo 19
      DENO$ = Left$(F0$, 30)
        If Left$(DENO$, 1) = "@" Then DENO$ = DESENCRI$(Mid$(F0$, 2, 29)) + " "
      CLAV$ = Mid$(F0$, 43, 12)
        If Left$(CLAV$, 1) = "@" Then CLAV$ = DESENCRI$(Mid$(F0$, 44, 11)) + " "
        CLAV$ = REPLACAR$(CLAV$, Chr$(0), Chr$(32))
      CLAV$ = Trim$(CLAV$) + Space$(12)
      If CLA$ = IDENT$ Then
        NUSER% = JJ%
        Label1 = DENO$
        USERNA$ = DENO$
        Text3.Enabled = True
        CONTRA$ = ""
        Text3.SetFocus
        Exit Sub
      End If
19  Next JJ%
    '
40  MsgBox "Usuario no Definido"
    CAIN% = CAIN% + 1
    If CAIN% > 3 Then
        MsgBox "Superado Limite Intentos de Acceso !!!"
        Close: End
    End If
    Text1.Text = ""
    CLA$ = "": CLB$ = "": CLAVE$ = ""
    Text1.SetFocus
    '
End Sub

Sub GRADATRAM()
   '
   RSW% = 0
   RSWX$ = Trim$(UCase$(LOADFILE$(LUCOM$ + "RESIWIN.DRV")))
   If Left$(RSWX$, 7) = "RESIWIN" Then
     RSW% = 1
   End If
   '
   If JOBID$ = "" Then
     Randomize Timer
     JOBI$ = ""
51   PPX% = Int(Rnd * 90)
     If PPX% < 48 Then GoTo 51
     If PPX% > 57 And PPX% < 65 Then GoTo 51
     JOBI$ = JOBI$ + Chr$(PPX%)
     If Len(JOBI$) < 8 Then GoTo 51
     JOBID$ = JOBI$
   End If
   JOBIX$ = ENCRIPTA$(JOBID$)
   '
   CLARAN$ = RANDSTRING$(6)
   CONTRA$ = RANDSTRING$(6)
   If VERDEFLOG% < 1 Then
      CONTRA$ = ENCRIPTA$(CLARAN$)
   End If
   '
   IDXX$ = LOADFILE$(LUCOM$ + "IDENTER.DRV")
   IDXX$ = Left$(IDXX$ + String$(40, 0), 40)
   '
   IDENCRIP$ = Left$(IDXX$, 38)
   J1$ = DESENCRI$(Left$(IDXX$, 38))
   IDENTER$ = Trim$(Mid$(J1$, 33, 6))
   NOMTER$ = Trim$(Mid$(J1$, 1, 32))
   '
   DEXX$ = LOADFILE("IDENSYS.DRV")
   If Len(DEXX$) >= 96 Then
     DEYY$ = DESENCRI$(Left$(DEXX$, 94))
     EMPREAC$ = Trim$(Mid$(DEYY$, 13, 30))
   End If
   '
   If MULTIEM% > 0 Then
     EPXX$ = Trim$(Mid$(LOADFILE$(LUCOM$ + "FLEXRAM.DRV"), 193, 30))
     If EPXX$ <> "" Then EMPREAC$ = EPXX$
   End If
   '
   If VERDEFLOG% = 1 Then
      USNBR% = 101
      USERNA$ = "Usuario DEMO"
      MENUPRIN$ = ""
      MODINVO$ = Command$
      LUDAT$ = "C:\FLEXOFT\DATOS\"
      EJER% = 1
      COMEJ% = 0
      FINEJ% = 0
      EMPREAC$ = "Su Empresa"
      LUBOOT$ = "C"
      IDENTER$ = "000000"
      NOMTER$ = "Estacion Local"
   End If
   '
   If EXISTE%("LANGUAGE.DRV") > 0 Then
      On Error Resume Next
      LANGUA% = Val(LOADFILE$("LANGUAGE.DRV"))
      On Error GoTo 0
   End If
   If LANGUA% < 0 Or LANGUA% > 5 Then LANGUA% = 0
   '

   RAMSTRI$ = String$(512, 0)
   Mid$(RAMSTRI$, 9, 2) = MKI$(USNBR%)
   Mid$(RAMSTRI$, 91, 6) = IDENTER$
     LUBOOT$ = "C"
   Mid$(RAMSTRI$, 97, 1) = LUBOOT$ + Space$(1)
   Mid$(RAMSTRI$, 99, 46) = LUDAT$ + Space$(46)
   Mid$(RAMSTRI$, 145, 46) = LUBAS$ + Space$(46)
   Mid$(RAMSTRI$, 193, 30) = EMPREAC$ + Space$(30)
   Mid$(RAMSTRI$, 223, 8) = JOBIX$
   Mid$(RAMSTRI$, 231, 30) = USERNA$ + Space$(30)
   Mid$(RAMSTRI$, 261, 30) = NOMTER$ + Space$(30)
   Mid$(RAMSTRI$, 293, 6) = CLARAN$
   Mid$(RAMSTRI$, 299, 6) = CONTRA$
   Mid$(RAMSTRI$, 305, 1) = Trim$(Str$(LANGUA%))
   Mid$(RAMSTRI$, 306, 1) = Trim$(Str$(RSW%))
   '
   Call SAVEFILE(LUCOM$ + "FLEXRAM.DRV", RAMSTRI$)
   '
   If EXISTE%(LUDAT$ + "MASTER.DAT") < 1 Then
      MSY$ = String$(128, 0)
      FLIM% = 32767
      Mid$(MSY$, 124, 2) = MKI$(FLIM%)
      Mid$(MSY$, 126, 2) = MKI$(9876 - FLIM%)
      PPX% = 23
      For I% = 124 To 127
        PPX% = PPX% + (I% - 123) * Asc(Mid$(MSY$, I%, 1))
      Next I%
      PPX% = PPX% Mod 256
      Mid$(MSY$, 128, 1) = Chr$(PPX%)
      Call SAVEFILE(LUDAT$ + "MASTER.DAT", MSY$)
    End If
   On Error Resume Next
   Shell Trim$(CMD$) + ".EXE #" + JOBID$ + "#", 4
   On Error GoTo 0
   Close: End
   '
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
   '
   If YACAR% < 1 Then Exit Sub
   '
   If Text3.Text = "" Then
      CLA$ = "": CLB$ = ""
   End If
   A% = KeyAscii
   KeyAscii = 0
   Text3.Text = CLB$
   If A% = 13 Then
      CONTRA$ = Left$(CLA$ + Space$(12), 12)
      Call BUCLA
      Exit Sub
   End If
   '
   If A% = 27 Then Close: End
   If A% < 48 Then Exit Sub
   If A% > 96 Then A% = A% - 32
   If Len(Text3.Text) < 12 Then
      CLB$ = CLB$ + "*"
      Text3.Text = CLB$
      CLA$ = CLA$ + Chr$(A%)
      CONTRA$ = Left$(CLA$ + Space$(12), 12)
   End If
          '
End Sub

Sub BUCLA()
   '
   Static CAIN%
   '
'TT1$ = "": For U& = 1 To Len(CLAV$): TT1$ = TT1$ + Str$(Asc(Mid$(CLAV$, U&, 1))): Next U&
'TT2$ = "": For U& = 1 To Len(CONTRA$): TT2$ = TT2$ + Str$(Asc(Mid$(CONTRA$, U&, 1))): Next U&
'MsgBox "CLAV$" + TT1$
'MsgBox "CONTRA" + TT2$

   If Left$(CLAV$, 12) = Left$(CONTRA$, 12) Then
       USNBR% = NUSER%
       UserName$ = DENUS$
       If USNBR% = 1 Then USNBR% = 101
       Call GRADATRAM
     Else
       CAIN% = CAIN% + 1
       If CAIN% > 3 Then
          MsgBox "Superado Limite Permitido Intentos de Acceso !!!"
          Close: End
       End If
       '
       MsgBox "Contraseña Incorrecta"
       Text3 = ""
       Text3.SetFocus
   End If
   '
End Sub
'
Sub VERITER()
    '
    If VERDEFLOG% > 0 Then Exit Sub
    If Left$(JOBID$, 2) = "WF" Then Exit Sub ' por terminal server
    '
    DoEvents
    SVL$ = SERIVOL$("C:")
    SVL1$ = Left$(SVL$ + Space$(12), 12)
    SVLX$ = ENCRIPTA$(SVL1$)
    'LUCOM$ = "C:\FLEXOFT\"
    '
    ' verificar identificacion de terminal
    Dim IDXX As String * 40
    N1% = FILEOPEN%(LUCOM$ + "IDENTER.DRV", 0, 40)
    '
    If LOF(N1%) >= 80 Then
        '
        For REID% = 1 To 2
           Get #N1%, REID%, IDXX$
           SUDI% = 0
           For I% = 1 To 38
              SUDI% = SUDI% + Asc(Mid$(IDXX$, I%, 1))
           Next I%
           If SUDI% <> CVI(Mid$(IDXX$, 39, 2)) Then
              Close #N1%
              Call HABILITER
              Close: End
           End If
        Next REID%
        '
        Get #N1%, 1, IDXX$
        IDENCRIP$ = Left$(IDXX$, 38)
        J1$ = DESENCRI$(Left$(IDXX$, 38))
        IDX$ = Trim$(Mid$(J1$, 33, 6))
        NTX$ = Trim$(Mid$(J1$, 1, 32))
        If IDX$ = "" Or NTX$ = "" Then
           Close #N1%
           Call HABILITER
           Close: End
        End If
        '
        Get #N1%, 2, IDXX$
        J1$ = DESENCRI$(Mid$(IDXX$, 11, 12))
        If J1$ <> SVL1$ Then
           Close #N1%
           Call HABILITER
           Close: End
        End If
        '
        GoTo 20        ' aprobo controles
    End If
    '
    Close #N1%          ' no existe IDENTER.DRV
    Call HABILITER
    Close: End
    '
20  Close #N1%
    IDENTER$ = IDX$
    NOMTER$ = NTX$
    '
    If VERCD% = 1 Or VERDEFLOG% > 0 Then PUETRA% = 1: Exit Sub
    '
    ' verificar terminal en IDENSYS.DRV
    Dim IDS As String * 96
    N1% = FILEOPEN%("IDENSYS.DRV", 10, 96)
    '
    For I% = 11 To LOF(N1%) / 96
       Get #N1%, I%, IDS$
       SUCAR% = 0
       For KI% = 1 To 94
          SUCAR% = SUCAR% + Asc(Mid$(IDS$, KI%, 1)) - 83
       Next KI%
       If CVI(Mid$(IDS$, 95, 2)) = SUCAR% Then
          SVLL$ = Mid$(IDS$, 8, 12)
          If SVLL$ = SVLX$ Then
            If Mid$(IDS$, 38, 38) = IDENCRIP$ Then
              GoTo 99
            End If
          End If
       End If
    Next I%
    '
    Close #N1%
    Call HABILITER
    Close: End
    '
99  Close #N1%
    On Error Resume Next
    MkDir "\FLEXOFT\LC" + IDENTER$
    On Error GoTo 0
    '
End Sub
'
Sub HABILITER()    ' habilita puesto de trabajo
    '
    Dim BLEC As String * 127
    Dim IDS As String * 96
    Dim IDXX As String * 40
    '
    'LUCOM$ = "C:\FLEXOFT\"
    '
    PORDISKETTE% = 0
    SEVOLC$ = Left$(SERIVOL$("C:") + Space$(12), 12)
    SVLC$ = ENCRIPTA$(Trim$(SEVOLC$))
    TZXC$ = ""
    For KKI% = 1 To Len(SVLC$)
      TZXC$ = TZXC$ + Str$(Asc(Mid$(SVLC$, KKI%, 1)))
    Next KKI%
    'MsgBox SVLC$ & "    " & TZXC$
    If Trim$(TZXC$) = "109 119 155 158 177 224 235 13 17" Then GoTo 60
    If UCase$(CMD$) = "FLCACHE" Then GoTo 60
    '
    DEXX$ = LOADFILE("IDENSYS.DRV")
    If Len(DEXX$) >= 96 Then
      DEYY$ = DESENCRI$(Left$(DEXX$, 94))
      EMPREAC$ = UCase$(Trim$(Mid$(DEYY$, 13, 30)))
    End If
    If InStr(EMPREAC$, "BRASBOM") > 0 Then GoTo 60
    '
10  Screen.MousePointer = 1
    PORDISKETTE% = 1
    OPEX% = MsgBox(FORMESS$("Coloque 'Diskette' de Instalación\en la Disketera 'A:'."), 17, "Habilitación Estación de Trabajo")
    If OPEX% <> 1 Then
      CLAX$ = ""
      While Len(CLAX$) < 6
        PPX% = Int(Rnd * 63)
        If PPX% >= 48 Then
          If PPX% <= 57 Then
            CLAX$ = CLAX$ + Chr$(PPX%)
          End If
        End If
      Wend
      '
      CONTRASE$ = InputBox$("Ingrese Contraseña", "Santo y Seña (" + CLAX$ + ")")
      If CONTRASE$ <> CONTRACLA$(CLAX$) Then
        Close: End
        Exit Sub
      End If
      PORDISKETTE% = 0
      GoTo 60
    End If
    '
    FLEN& = 0
    On Error GoTo 25
    FLEN& = FileLen("A:FLEXKEY.KEY")
    On Error GoTo 0
    GoTo 30
    '
25  ERRX = Err.Number
    Resume 27
27  On Error GoTo 0
    If ERRX = 53 Then
         MsgBox ("Diskette Incorrecto.")
       Else
         MsgBox ("Error de Acceso a la Disketera.")
    End If
    GoTo 10
    '
30  If FLEN& < 127& * 499 Then
       MsgBox ("Diskette Incorrecto.")
       GoTo 10
    End If
    '
40  On Error GoTo 45
    N1% = FreeFile
    Open "A:FLEXKEY.KEY" For Random Access Read Write Shared As #N1% Len = 127
    Get #N1%, 1, BLEC$
    Put #N1%, 1, BLEC$
    On Error GoTo 0
    GoTo 50
    '
45  Resume 46
46  On Error GoTo 0
    MsgBox (FORMESS$("Diskette Protegido Contra Escritura.\Quite la Faja de Protección\e Intente de Nuevo."))
    GoTo 10
    '
50  Screen.MousePointer = 11
    SEVOLA$ = Left$(SERIVOL$("A:") + Space$(12), 12)
    SEVOLC$ = Left$(SERIVOL$("C:") + Space$(12), 12)
    CULOG$ = Left$(CurDir, 2)
    SEVOLF$ = Left$(SERIVOL$(CULOG$) + Space$(12), 12)
    If Trim$(SEVOLA$) = "" Then
       Close #N1%
       Screen.MousePointer = 1
       MsgBox ("Diskette No Válido (Err.50)")
       GoTo 10
    End If
    '
    If Trim$(SEVOLC$) = "" Then
       Close #N1%
       Screen.MousePointer = 1
       MsgBox (FORMESS$("Error de Acceso a Disco 'C:'.\Consulte a su Soporte de Sistemas."))
       GoTo 10
    End If
    '
    If Trim$(SEVOLF$) = "" Then
       Close #N1%
       Screen.MousePointer = 1
       MsgBox (FORMESS$("Error de Acceso a Disco '" + CULOG$ + "'.\Consulte a su Soporte de Sistemas."))
       Screen.MousePointer = 11
       GoTo 10
    End If
    '
    For UL& = 1 To FLEN& / 127
       Get #N1%, UL&, BLEC$
       SUCAR% = 0
       For KI% = 1 To 125
          SUCAR% = SUCAR% + Asc(Mid$(BLEC$, KI%, 1)) - 83
       Next KI%
       If CVI(Mid$(BLEC$, 126, 2)) <> SUCAR% Then
          Close #N1%
          Screen.MousePointer = 1
          MsgBox ("Diskette No Válido (Err.51)")
          GoTo 10
       End If
    Next UL&
    Screen.MousePointer = 1
    '
    Get #N1%, 1, BLEC$
    If DESENCRI$(Mid$(BLEC$, 17, 12)) <> SEVOLA$ Then
       Close #N1%
       MsgBox ("Diskette No Válido (Err.52)")
       GoTo 10
    End If
    '
    Get #N1%, 3, BLEC$
    SDES$ = DESENCRI$(Mid$(BLEC$, 71, 12))
    If SDES$ <> String$(12, 23) Then
      If SDES$ <> SEVOLF$ Then
        Close #N1%
        MsgBox (FORMESS("Diskette no Válido\Ya Utilizado en Otra Instalación."))
        GoTo 10
      End If
    End If
    '
    Get #N1%, 2, BLEC$
    SDES$ = DESENCRI$(Mid$(BLEC$, 29, 12))
    If SDES$ <> String$(12, 19) Then
       If SDES$ <> SEVOLC$ Then
          N2% = FILEOPEN%("IDENSYS.DRV", 0, 96)
          For I% = 11 To LOF(N2%) / 96
            Get #N2%, I%, IDS$
            SUCAR% = 0
            For KI% = 1 To 94
               SUCAR% = SUCAR% + Asc(Mid$(IDS$, KI%, 1)) - 83
            Next KI%
            If CVI(Mid$(IDS$, 95, 2)) = SUCAR% Then
               SVLX$ = DESENCRI$(Mid$(IDS$, 8, 12))
               If SVLX$ = SDES$ Then
                 IDENCRIP$ = Mid$(BLEC$, 59, 38)
                 NOTERVIE$ = Trim$(Left$(DESENCRI$(IDENCRIP$), 32))
                 MENSA$ = "El Diskette de Instalación Utilizado"
                 MENSA$ = MENSA$ + "\Corresponde a la Terminal Denominada"
                 MENSA$ = MENSA$ + "\'" + NOTERVIE$ + "'." + "\  "
                 MENSA$ = MENSA$ + "\Si Continúa con la Instalación,"
                 MENSA$ = MENSA$ + "ese Puesto de Trabajo Quedará DESHABILITADO.\   "
                 MENSA$ = MENSA$ + "\Desea Cancelar o Continuar ?"
                 If MsgBox(FORMESS$(MENSA$), 273, "Habilitación de Estación de Trabajo") <> 1 Then
                   Close: End
                 End If
                 '
                 IDS$ = String$(96, 0)
                 Put #N2%, I%, IDS$
                 '
                 Get #N1%, 2, BLEC$
                 TTXX$ = Left$(BLEC$, 125)
                 Mid$(TTXX$, 29, 12) = ENCRIPTA$(String$(12, 19))
                 SUCAR% = 0
                 For KI% = 1 To 125
                   SUCAR% = SUCAR% + Asc(Mid$(TTXX$, KI%, 1)) - 83
                 Next KI%
                 BLEC$ = TTXX$ + MKI$(SUCAR%)
                 Put #N1%, 2, BLEC$
                 Close #N2%
                 GoTo 60
               End If
            End If
          Next I%
          '
          Close #N1%
          Close #N2%
          MsgBox (FORMESS$("Diskette no Válido\Ya Utilizado en Otra Instalación."))
          GoTo 10
          '
       End If
    End If
    '
60  ID = Timer: While ID > 9999.99: ID = ID - 9000: Wend
    ID = Int(100 * ID)
    IDX$ = Mid$(Str$(ID), 2)
    While Len(IDX$) < 6: IDX$ = "0" + IDX$: Wend
    NTX$ = ""
    '
    If EXISTE%(LUCOM$ + "IDENTER.DRV") > 0 Then
      IDTX$ = Left$(LOADFILE$(LUCOM$ + "IDENTER.DRV"), 40)
      If Len(IDTX$) >= 40 Then
        SUDI% = 0: For I% = 1 To 38
                     SUDI% = SUDI% + Asc(Mid$(IDTX$, I%, 1))
                   Next I%
        If CVI(Mid$(IDTX$, 39, 2)) = SUDI% Then
          IDTY$ = DESENCRI$(Left$(IDTX$, 38))
          IDX$ = Mid$(IDTY$, 33, 6)
          NTX$ = Left$(IDTY$, 32)
        End If
      End If
    End If
    '
    VDEF$ = Space$(52)
    Call REPLA(VDEF$, IDX$, 1, 6)
    Call REPLA(VDEF$, SEVOLC$, 41, 12)
70  OBX$ = InputBox$("Nombre Identificador de Terminal", "Identificador Estación de Trabajo", NTX$)
    If OBX$ = "" Then Close: End
    '
    IDT$ = Trim$(OBX$)
    If Len(IDT$) < 4 Then
       MsgBox ("Ingrese un Nombre para la Terminal (min. 4 pos.)")
       GoTo 70
    End If
    '
    NTID$ = Space$(38)
    Call REPLA(NTID$, Left$(OBX$, 32), 1, 32)
    Call REPLA(NTID$, IDX$, 33, 6)
    IDENTER$ = Trim$(IDX$)
    NOMTER$ = Trim$(Left$(OBX$, 32))
    IDENCRIP$ = ENCRIPTA$(NTID$)
    '
    If PORDISKETTE% > 0 Then
      Get #N1%, 2, BLEC$
      TTXX$ = Left$(BLEC$, 125)
      Call REPLA(TTXX$, ENCRIPTA$(SEVOLC$), 29, 12)
      Call REPLA(TTXX$, IDENCRIP$, 59, 38)
      SUCAR% = 0
      For KI% = 1 To 125
         SUCAR% = SUCAR% + Asc(Mid$(TTXX$, KI%, 1)) - 83
      Next KI%
      BLEC$ = TTXX$ + MKI$(SUCAR%)
      Put #N1%, 2, BLEC$
      Close #N1%
    End If
    '
    N2% = FILEOPEN%("IDENSYS.DRV", 0, 96)
    REI% = 1 + LOF(N2%) / 96: If REI% < 11 Then REI% = 11
    TTXX$ = RANDSTRING$(94)
    Call REPLA(TTXX$, ENCRIPTA$(SEVOLC$), 8, 12)
    Call REPLA(TTXX$, IDENCRIP$, 38, 38)
    SUCAR% = 0
    For KI% = 1 To 94
       SUCAR% = SUCAR% + Asc(Mid$(TTXX$, KI%, 1)) - 83
    Next KI%
    IDS$ = TTXX$ + MKI$(SUCAR%)
    Put #N2%, REI%, IDS$
    Close #N2%
    '
    N3% = FILEOPEN%(LUCOM$ + "IDENTER.DRV", 0, 40)
    X$ = ENCRIPTA$(NTID$)
    SUDI% = 0: For I% = 1 To Len(X$)
                  SUDI% = SUDI% + Asc(Mid$(X$, I%, 1))
               Next I%
    Y$ = String$(40, 0)
    Call REPLA(Y$, X$, 1, 38)
    Call REPLA(Y$, MKI$(SUDI%), 39, 2)
    IDXX$ = Y$
    Put #N3%, 1, IDXX$
    '
    Y$ = String$(40, 0)
    Call REPLA(Y$, "B800", 1, 4)
    Call REPLA(Y$, MKS$(47104), 5, 4)
    Call REPLA(Y$, "SI", 9, 2)
    X$ = ENCRIPTA$(SEVOLC$)
    Call REPLA(Y$, X$, 11, 12)
    SUDI% = 0: For I% = 1 To 38
                 SUDI% = SUDI% + Asc(Mid$(Y$, I%, 1))
               Next I%
    Call REPLA(Y$, MKI$(SUDI%), 39, 2)
    IDXX$ = Y$
    Put #N3%, 2, IDXX$
    Close #N3%: N3% = 0
    '
    MsgBox (FORMESS$("Habilitación de Terminal '" + NOMTER$ + "' Completada Exitosamente.\  \Re-Inicie Ahora el Sistema."))
    '
99 End Sub

Sub UPDATES()
   '
   Screen.MousePointer = 11
   DIREPRO$ = App.Path
   If EXISTE%("F:\FLEXOFT\NOVEDAD\.") > 0 Then
       DIRENEW$ = "F:\FLEXOFT\NOVEDAD"
     Else
       DIRENEW$ = Left$(DIREPRO$, 2) + "\FLEXOFT\NOVEDAD"
   End If
   DIREVIE$ = DIREPRO$ + "\VIEJOS\"
   '
   YAMENSA% = 0
   INSTALAC% = 0
   '
   On Error Resume Next
   MkDir DIREVIE$
   On Error GoTo 0
   '
   On Error Resume Next
   File1.Path = DIRENEW$
   If Err < 1 Then
     On Error GoTo 0
     File1.Pattern = "*.*"
     '
     For KU& = 1 To File1.ListCount
       ANAME$ = Trim$(File1.List(KU& - 1))
       PPXX% = InStr(ANAME$, ".")
       If PPXX% < 1 Then PPXX% = 1 + Len(ANAME$)
       EXTEN$ = Trim$(UCase$(Mid$(ANAME$, PPXX% + 1)))
       '
       If Len(EXTEN$) = 3 Then
         '
         If EXTEN$ = "RFS" Or EXTEN$ = "FRM" Then
           FECHANUE# = FileDateTime(DIRENEW$ + "\" + ANAME$)
           FECHAVIE# = 0
           If EXISTE%(LUDAT$ + ANAME$) > 0 Then
             FECHAVIE# = FileDateTime(LUDAT$ + ANAME$)
           End If
           If FECHANUE# > FECHAVIE# + 0.0001 Then  ' ESO SON APROX 6 SEGUNDOS
             On Error Resume Next
             FileCopy DIRENEW$ + "\" + ANAME$, LUDAT$ + ANAME$
             If Err Then
               MsgBox "Imposible Instalar (3) " + ANAME$
             End If
             On Error GoTo 0
           End If
           GoTo 29         ' SALE DE INSTALACION DE DATOS
         End If
         '
         RUTACOM$ = DIREPRO$ + "\"
         If EXISTE%(RUTACOM$ + ANAME$) > 0 Or EXTEN$ = "PRF" Then
           FECHANUE# = FileDateTime(DIRENEW$ + "\" + ANAME$)
           FECHAVIE# = 0
           If EXISTE%(RUTACOM$ + ANAME$) > 0 Then
             FECHAVIE# = FileDateTime(RUTACOM$ + ANAME$)
           End If
           If FECHANUE# > FECHAVIE# + 0.0001 Then  ' ESO SON APROX 6 SEGUNDOS
             If YAMENSA% = 0 Then
               MESSA$ = "Hay Actualizaciones Listas para Instalarse." + Chr$(10) + Chr$(13) + Chr$(10) + Chr$(13)
               MESSA$ = MESSA$ + "Para Instalar las Actualizaciones Disponibles" + Chr$(10) + Chr$(13)
               MESSA$ = MESSA$ + "Cierre las Aplicaciones FLEXOFT en TODAS las" + Chr$(10) + Chr$(13)
               MESSA$ = MESSA$ + "Estaciones de Trabajo." + Chr$(10) + Chr$(13) + Chr$(10) + Chr$(13)
               MESSA$ = MESSA$ + "¿ Desea Ahora Actualizar su Sistema FLEXOFT ?"
               RESPU% = MsgBox(MESSA$, vbYesNoCancel)
               If RESPU% <> 6 Then
                 Screen.MousePointer = 1
                 Exit Sub     ' 6 es = SI
               End If
               YAMENSA% = 1
             End If
             '
             EXTECO$ = Left$(EXTE$, 2)
             ANAMCO$ = Left$(ANAME$, Len(ANAME$) - 1)
             On Error Resume Next
             Kill RUTACOM$ + ANAMCO$ + "0"
             Kill DIREVIE$ + ANAMCO$ + "9"
             On Error GoTo 0
             If EXISTE%(DIREVIE$ + ANAMCO$ + "9") > 0 Then
               MsgBox "Imposible Instalar (1) " + ANAME$
               On Error GoTo 0
               GoTo 29
             End If
             '
             On Error Resume Next
             For KI% = 8 To 0 Step -1
               Name DIREVIE$ + ANAMCO$ + Trim$(Str$(KI%)) As DIREVIE$ + ANAMCO$ + Trim$(Str$(KI% + 1))
             Next KI%
             On Error GoTo 0
             If EXISTE%(DIREVIE$ + ANAMCO$ + "0") > 0 Then
               MsgBox "Imposible Instalar (2) " + ANAME$
               On Error GoTo 0
               GoTo 29
             End If
             '
             If InStr(ANAME$, "MSTDRV04") < 1 Then  ' PARA QUE NO RESPALDE RECLAMO
               If EXISTE%(RUTACOM$ + ANAME$) < 1 Then GoTo 24
               On Error Resume Next
               Name RUTACOM$ + ANAME$ As RUTACOM$ + ANAMCO$ + "0"
               If Err Then
                 MsgBox "Imposible Instalar (4) " + ANAME$
                 On Error GoTo 0
                 GoTo 29
               End If
               On Error Resume Next
               Name RUTACOM$ + ANAMCO$ + "0" As DIREVIE$ + ANAMCO$ + "0"
               If Err Then
                 MsgBox "Imposible Instalar (5) " + ANAME$
                 On Error GoTo 0
                 GoTo 29
               End If
             End If
             '
24           On Error Resume Next
             FileCopy DIRENEW$ + "\" + ANAME$, RUTACOM$ + ANAME$
             If Err Then
               MsgBox "Imposible Instalar " + ANAME$
               On Error GoTo 0
               GoTo 29
             End If
             INSTALAC% = 1
             '
           End If
         End If
       End If
29   Next KU&
     '
     If INSTALAC% > 0 Then MsgBox "Actualizaciones Instaladas"
     '
   End If
   Screen.MousePointer = 1
   '
End Sub

Sub ABRELOGON()
   '
   ESUNDEMO% = 0
   On Error Resume Next
   IKXX$ = Label99.Caption
   If Err Then
     On Error GoTo 0
     MsgBox "Imposible Identificar versión Licencia / Demo - Consulte !"
     Close: End: Exit Sub
   End If
   '
   If Mid$(IKXX$, 51, 9) <> "........." Then
     IKYY$ = Mid$(IKXX$, 51, 9)
     '
     ' DIGITO VERIFICADOR DEL SECTOR DE CONTROL
     DIVERI& = 0
     For KKI% = 1 To 8
       KKJ% = KKI%: If KKI% > 4 Then KKJ% = KKI% + 1
       DIVERI& = Val(Mid$(IKYY$, KKJ%, 1)) + (7 * KKJ% ^ 2)
     Next KKI%
     DIVERI& = DIVERI& Mod 10
     If Val(Mid$(IKYY$, 5, 1)) <> DIVERI& Then
       MsgBox "Módulo '" + App.EXEName + "' Indebidamente Alterado (D.Ver)"
       Close: End: Exit Sub
     End If
     '
     ESUNDEMO% = 1
   End If
   '
   '
   If ESUNDEMO% = 1 Then
        LUCOM$ = "C:\FLEXOFT\"
        Randomize Timer
14      JOBI$ = ""
15      PPX% = Int(Rnd * 90)
        If PPX% < 48 Then GoTo 15
        If PPX% > 57 And PPX% < 65 Then GoTo 15
        JOBI$ = JOBI$ + Chr$(PPX%)
        If Len(JOBI$) < 8 Then GoTo 15
        JOBID$ = JOBI$
        If Left$(JOBID$, 2) = "WF" Then GoTo 14
      Else
        ' VERIFICAR ACCESO LECTURA ESCRITURA A C:\FLEXOFT
        On Error Resume Next
        MkDir "C:\FLEXOFT"
        On Error GoTo 0
        NX% = FreeFile
        On Error Resume Next
        Open "C:\FLEXOFT\RCONTROL.DRV" For Output As #NX%
        If Err Then
          On Error GoTo 0
          Close #NX%
          GoTo 50
        End If
        '
        ' AQUI RESULTA QUE HAY ACCESO LECTOESCRITURA A C:\FLEXOFT, ENTONCES
        LUCOM$ = "C:\FLEXOFT\"
        Randomize Timer
24      JOBI$ = ""
25      PPX% = Int(Rnd * 90)
        If PPX% < 48 Then GoTo 25
        If PPX% > 57 And PPX% < 65 Then GoTo 25
        JOBI$ = JOBI$ + Chr$(PPX%)
        If Len(JOBI$) < 8 Then GoTo 25
        JOBID$ = JOBI$
        If Left$(JOBID$, 2) = "WF" Then GoTo 24
   End If
   Exit Sub
   '
   '
50 ' LLEGA AQUI SI NO ES UN DEMO Y NO PUDO ACCEDER A C:\FLEXOFT
   '
   APPDIR$ = App.Path
   LICARTRA$ = LOADFILE$(APPDIR$ + "\LICARTRA.FLX")
   For KU& = 1 To Len(LICARTRA$) / 128
     XX$ = Mid$(LICARTRA$, 128 * KU& - 127, 128)
     If Trim$(Mid$(XX$, 9, 16)) = "" Then
       If Left$(XX$, 2) = "WF" Then
         JOBID$ = Left$(XX$, 8)
         LUCOM$ = APPDIR$ + "\" + JOBID$ + "\"
         If EXISTE%(LUCOM$ + ".") > 0 Then
           AHORA$ = Format$(Now, "DD/MM/YY   HH:MM")
           Mid$(XX$, 9, 16) = AHORA$
           Mid$(LICARTRA$, 128 * KU& - 127, 128) = XX$
           Call SAVEFILE(APPDIR$ + "\LICARTRA.FLX", LICARTRA$)
           Exit Sub
         End If
       End If
     End If
   Next KU&
   '
   Randomize Timer
   JOBI$ = ""
55 PPX% = Int(Rnd * 90)
   If PPX% < 48 Then GoTo 55
   If PPX% > 57 And PPX% < 65 Then GoTo 55
   JOBI$ = JOBI$ + Chr$(PPX%)
   If Len(JOBI$) < 6 Then GoTo 55
   JOBID$ = "WF" + JOBI$
   LUCOM$ = APPDIR$ + "\" + JOBID$ + "\"
   On Error Resume Next
   MkDir APPDIR$ + "\" + JOBID$
   If Err Then
     MsgBox "Imposible Crear Carpeta de Trabajo " + APPDIR$ + "\" + JOBID$
     Close: End: Exit Sub
   End If
   '
   While (Len(LICARTRA$) Mod 128) <> 0
     LICARTRA$ = LICARTRA$ + " "
   Wend
   XX$ = Space$(128)
   Mid$(XX$, 1, 8) = JOBID$
   AHORA$ = Format$(Now, "DD/MM/YY   HH:MM")
   Mid$(XX$, 9, 16) = AHORA$
   LICARTRA$ = LICARTRA$ + XX$
   Call SAVEFILE(APPDIR$ + "\LICARTRA.FLX", LICARTRA$)
   '
End Sub

Sub VERIBAK()
   '
   AA1$ = Trim$(LOADFILE$("RECUBAK"))
   If AA1$ <> "RECUBAK" Then
     Exit Sub
   End If
   '
   YAMENSA% = 0
   INSTALAC% = 0
   '
   MESSA$ = "Se ha Detectado una Reciente Restauración de Datos" + Chr$(10) + Chr$(13)
   MESSA$ = MESSA$ + "a Partir de una Copia de Respaldo Interna." + Chr$(10) + Chr$(13) + Chr$(10) + Chr$(13)
   MESSA$ = MESSA$ + "Para Completar Dicho Proceso de Restauración" + Chr$(10) + Chr$(13)
   MESSA$ = MESSA$ + "Cierre las Aplicaciones FLEXOFT en TODAS las" + Chr$(10) + Chr$(13)
   MESSA$ = MESSA$ + "Estaciones de Trabajo." + Chr$(10) + Chr$(13) + Chr$(10) + Chr$(13)
   MESSA$ = MESSA$ + "¿ Desea Ahora Completar el Proceso de Restauración ?"
   RESPU% = MsgBox(MESSA$, vbYesNoCancel)
   If RESPU% <> 6 Then
     Screen.MousePointer = 1
     Close: End
     Exit Sub     ' 6 es = SI
   End If
   '
   Screen.MousePointer = 11
   LUACC$ = LUDAT$ + "BdAccess\"
   '
   OPENFORMS = DoEvents
   HUBOERR% = 0
   AORI$ = Trim$(Dir(LUDAT$ + "*.MDB"))
20 If AORI$ <> "" Then
     ADESTI$ = AORI$
     If InStr(ADESTI$, "~") > 0 Then
       If Left$(ADESTI$, 6) = "ARTICU" Then
           ADESTI$ = "ARTICULOS.MDB"
         ElseIf Left$(ADESTI$, 6) = "BADACL" Then
           ADESTI$ = "BADACLIEN.MDB"
         ElseIf Left$(ADESTI$, 6) = "BADAPR" Then
           ADESTI$ = "BADAPROVE.MDB"
         Else
           MsgBox "Imposible Identificar '" + ADESTI$ + "'"
       End If
     End If
     '
25   ' AQUI COPIA Y RECUPERA
     On Error Resume Next
     FileCopy LUDAT$ + AORI$, LUACC$ + ADESTI$
     If Err Then
        MsgBox "Error " & Err & " - Imposible Restaurar '" + ADESTI$ + "'."
        HUBOERR% = 1
     End If
     On Error GoTo 0
     '
     On Error Resume Next
     Kill LUDAT$ + AORI$
     On Error GoTo 0
     '
     AORI$ = Trim$(Dir)
     GoTo 20
     '
   End If
   '
   Screen.MousePointer = 1
   If HUBOERR% < 1 Then
        Call SAVEFILE("RECUBAK", "")
        MsgBox "Restauracion Completa"
      Else
        MsgBox "Restauración con Errores - Consulte."
   End If
   '
End Sub

Sub BLOQSYS()
   '
   opx% = MsgBox("Realmente Desea Bloquear los Ejecutables en la Presente Instalación?", vbYesNo)
   If opx% <> 6 Then Exit Sub
   '
   KKII% = 0
   CUDI$ = Left$(CurDir$, 2)
   If Mid$(CUDI$, 2, 1) <> ":" Then
     MsgBox "Proceso Imposible en la Carpeta " + CurDir$
     Exit Sub
   End If
   SEVOL$ = SERIVOL$(CUDI$)
   If Len(SEVOL$) < 9 Then
     MsgBox "Proceso Imposible en Volumen " + SEVOL$
     Exit Sub
   End If
   '
   While Len(SEVOL$) < 13
     KKII% = KKII% + 1
     SEVOL$ = SEVOL$ + Chr$(18 + Asc(Mid$(SEVOL$, KKII%, 1)))
   Wend
   SEVOL$ = ENCRIPTA$(Left$(SEVOL$, 13))
   '
   For U& = 1 To File2.ListCount
     '
     File2.ListIndex = U& - 1
     '
     FINA$ = File2.List(U& - 1)
     LOFI$ = LOADFILE$(CurDir + "\" + FINA$)
     PQXX& = InStr(LOFI$, "Microsoft Vísual Basic 5.O")
     '
     If PQXX& > 0 Then
       '
       PPXX% = InStr(FINA$, ".")
       If PPXX% < 1 Then PPXX% = 1 + Len(FINA$)
       EXENA1$ = UCase$(Left$(FINA$, PPXX% - 1))
       EXENA$ = ENCRIPTA$(UCase$(Left$(FINA$, PPXX% - 1)))
       If Len(EXENA$) < 6 Then GoTo 19
       '
       While Len(EXENA$) < 13
         CARAVA% = (Asc(Mid$(EXENA$, 2, 1)) + Asc(Mid$(EXENA$, 4, 1))) Mod 256
         EXENA$ = Chr$(CARAVA%) + EXENA$
         CARAVA% = (Asc(Mid$(EXENA$, Len(EXENA$) - 3, 1)) + Asc(Mid$(EXENA$, Len(EXENA$) - 1, 1))) Mod 256
         EXENA$ = EXENA$ + Chr$(CARAVA%)
       Wend
       EXENA$ = Left$(EXENA$, 13)
       '
       CSTRI$ = Chr$(64)
       For KKI% = 2 To 13
         CARAVA% = (KKI% - 1) * Asc(Mid$(EXENA$, KKI%, 1)) + (KKI% + 3) * Asc(Mid$(SEVOL$, KKI%, 1))
         While CARAVA% < 48: CARAVA% = CARAVA% + 43: Wend
         While CARAVA% > 90: CARAVA% = CARAVA% - 43: Wend
         CSTRI$ = CSTRI$ + Chr$(CARAVA%)
       Next KKI%
       '
       If Mid$(LOFI$, PQXX& + 27, 13) = "............." Then
         Mid$(LOFI$, PQXX& + 27, 13) = CSTRI$
         DESTI$ = CurDir$ + "\" + FINA$
         If UCase$(EXENA1$) <> UCase$(App.EXEName) Then
           Call SAVEFILE(DESTI$, LOFI$)
         End If
       End If
       '
     End If
     '
19 Next U&
   '
30 MsgBox "Bloqueo Completo"

End Sub

'\\\OT-06 -0406-MH-21/09/06///
Sub INSTALACON()
     '
     MENSA$ = FORMESS$("Se han detectado controladores no instalados\en la presente estación de trabajo.\\Esto puede deberse a una reciente reinstalación\o reconfiguración del Sistema Operativo Windows")
     MsgBox (MENSA$)
     '
     DISCO$ = Left$(App.Path, 2)
     If Mid$(DISCO$, 2, 1) <> ":" Then GoTo 30
     '
     ARCHI1$ = DISCO$ + "\FLEXOFT\FX_SETUP\WS_SETUP.EX_"
     ARCHI2$ = DISCO$ + "\FLEXOFT\FX_SETUP\SETUP.EXE"
     '
     If EXISTE%(ARCHI1$) = 1 And EXISTE%(ARCHI2$) = 1 Then
          MENSA$ = FORMESS$("¿ Desea Intentar la Instalación Automática\de los Controladores Faltantes ?")
          If MsgBox(MENSA$, vbYesNo) = vbYes Then
            opx = Shell(ARCHI2$, vbNormalFocus)
          End If
       Else
30       MENSA$ = FORMESS$("No es Posible la Instalación Automática\de los Controladores Faltantes.\ \Consulte a su Soporte de Sistemas FLEXOFT.")
         MsgBox (MENSA$)
     End If
     Close: End
End Sub
'\\\OT-06 -0406-MH-FIN///
