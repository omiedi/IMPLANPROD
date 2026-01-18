VERSION 5.00
Object = "{20C62CAE-15DA-101B-B9A8-444553540000}#1.1#0"; "MSMAPI32.OCX"
Begin VB.Form FLEXLOG4 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00DDDDC0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Sistema Preconfigurado FLEXOFT (r)"
   ClientHeight    =   3225
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   4695
   Icon            =   "FLOGBRE.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3225
   ScaleWidth      =   4695
   StartUpPosition =   2  'CenterScreen
   Begin MSMAPI.MAPIMessages MAPIMessages1 
      Left            =   0
      Top             =   525
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   327680
      AddressEditFieldCount=   1
      AddressModifiable=   0   'False
      AddressResolveUI=   0   'False
      FetchSorted     =   0   'False
      FetchUnreadOnly =   0   'False
   End
   Begin MSMAPI.MAPISession MAPISession1 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   327680
      DownloadMail    =   -1  'True
      LogonUI         =   -1  'True
      NewSession      =   0   'False
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00DDDDC0&
      Caption         =   "Web Services"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   720
      Left            =   105
      TabIndex        =   20
      Top             =   2415
      Visible         =   0   'False
      Width           =   3585
      Begin VB.PictureBox Picture3 
         BackColor       =   &H00DDDDC0&
         BorderStyle     =   0  'None
         Height          =   400
         Left            =   50
         ScaleHeight     =   405
         ScaleWidth      =   3450
         TabIndex        =   21
         Top             =   210
         Width           =   3450
         Begin VB.Label Label3 
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   525
            TabIndex        =   22
            Top             =   105
            Width           =   2850
         End
         Begin VB.Image Image3 
            Height          =   240
            Left            =   105
            Picture         =   "FLOGBRE.frx":030A
            Stretch         =   -1  'True
            ToolTipText     =   "Doble-Click para Des-Habilitar Web Services en esta Estación de Trabajo"
            Top             =   105
            Width           =   240
         End
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   1995
      Top             =   0
   End
   Begin VB.FileListBox File2 
      Height          =   870
      Left            =   105
      Pattern         =   "*.EXE"
      TabIndex        =   18
      Top             =   5565
      Visible         =   0   'False
      Width           =   4530
   End
   Begin VB.ListBox List1 
      Height          =   2205
      Left            =   105
      TabIndex        =   17
      Top             =   3675
      Width           =   4530
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00DDDDC0&
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
      Height          =   2220
      Left            =   105
      TabIndex        =   0
      Top             =   105
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
         Top             =   1605
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
         Top             =   390
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
         Height          =   330
         Left            =   210
         TabIndex        =   6
         Top             =   1000
         Visible         =   0   'False
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
         Top             =   1575
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
         Top             =   525
         Width           =   720
      End
      Begin VB.Image Image1 
         Height          =   585
         Left            =   315
         Picture         =   "FLOGBRE.frx":114C
         Stretch         =   -1  'True
         Top             =   210
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
         Top             =   1785
         Width           =   705
      End
      Begin VB.Image Image2 
         Height          =   525
         Left            =   210
         Picture         =   "FLOGBRE.frx":1456
         Stretch         =   -1  'True
         Top             =   1470
         Width           =   585
      End
      Begin VB.Image Image4 
         Height          =   525
         Left            =   210
         Picture         =   "FLOGBRE.frx":1760
         Stretch         =   -1  'True
         Top             =   1470
         Visible         =   0   'False
         Width           =   630
      End
      Begin VB.Image Image5 
         Height          =   540
         Left            =   210
         Picture         =   "FLOGBRE.frx":1A6A
         Stretch         =   -1  'True
         Top             =   1470
         Visible         =   0   'False
         Width           =   645
      End
   End
   Begin VB.FileListBox File1 
      Height          =   675
      Left            =   2520
      TabIndex        =   14
      Top             =   0
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00004000&
      Height          =   3405
      Left            =   3840
      ScaleHeight     =   3345
      ScaleWidth      =   1035
      TabIndex        =   7
      Top             =   -100
      Width           =   1095
      Begin VB.CommandButton Command6 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   105
         Picture         =   "FLOGBRE.frx":1D74
         Style           =   1  'Graphical
         TabIndex        =   19
         ToolTipText     =   "Administrar Licencias"
         Top             =   2600
         Width           =   600
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
         Height          =   600
         Left            =   120
         Picture         =   "FLOGBRE.frx":21B6
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Ingresa al Sistema"
         Top             =   890
         Width           =   600
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
         Height          =   600
         Left            =   120
         Picture         =   "FLOGBRE.frx":24C0
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Cierra - Abandona la Aplicación"
         Top             =   290
         Width           =   600
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
         Picture         =   "FLOGBRE.frx":260A
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
         Picture         =   "FLOGBRE.frx":2914
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
         Left            =   90
         TabIndex        =   15
         ToolTipText     =   "Doble Click para Bajar Controladores (OCX's y DLL's)"
         Top             =   1520
         Width           =   645
      End
      Begin VB.Image Image6 
         Height          =   390
         Left            =   -315
         Picture         =   "FLOGBRE.frx":2A5E
         ToolTipText     =   "Doble-Click para Refrescar Opción de Bajada Automática de Novedades y para Actualizar FLOGON.EXE"
         Top             =   1995
         Width           =   1515
      End
   End
   Begin VB.Label Label99 
      Caption         =   "Microsoft Vísual Basic 5.O ................................"
      Height          =   225
      Left            =   525
      TabIndex        =   16
      Top             =   315
      Visible         =   0   'False
      Width           =   2745
   End
End
Attribute VB_Name = "FLEXLOG4"
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
Dim EMPREAC$, CodiEmp%, CodiSuc%
Dim MULTIEM%
Dim YACAR%
'
Dim WEBSE$
Dim SIEMPREADELANTE%

Private Sub Command1_Click()
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
'   If Len(Text3) > 0 Then
'       Call Text3_KeyPress(13)
'     Else
'       Call Text1_KeyPress(13)
'   End If
   Call Text1_KeyPress(13)
   If USNBR% > 0 Or NUSER% > 0 Then
      Call Text3_KeyPress(13)
   End If
End Sub

Private Sub Command6_Click()
    Command6.Enabled = False
    SELICEN2.Show 1
    Command6.Enabled = True
End Sub
'
Private Sub Form_Load()
    '
    If App.PrevInstance = True Then
       MsgBox "FLOGON ya en Ejecución en esta Terminal."
       On Error Resume Next
       AppActivate Caption
       SendKeys Chr$(31), True
       On Error GoTo 0
       OPENFORMS = DoEvents
       Close: End
       Exit Sub
    End If
    '
    On Error Resume Next
    Caption = "Logon FLEXOFT - " + Format(FileDateTime(App.Path + "\FLOGON.EXE"), "dd/mm/yy  hh:mm")
    On Error GoTo 0
    UL$ = Left$(App.Path, 2)
' ESTO PARA RESOLVER EL TEMA DE CIVAL DEL 11/02/2011
' LUCOM$ = App.Path: If Right$(LUCOM$, 1) <> "\" Then LUCOM$ = LUCOM$ + "\"
    '
    SVLX$ = Trim$(SERIVOL$(UL$))
    PVC$ = PROVENCI$
    If PVC$ <> "" Then SVLX$ = SVLX$ + " / " + PROVENCI$
    Frame3.Caption = Frame3.Caption + " (" + SVLX$ + ")"
    '
    Text1.Enabled = False
    Text3.Enabled = False
    Show
    '
    Call GRARECEL
    '
    Refresh
    DoEvents
    'Load CTRLSHOW
    '
''    FrmErrorControl.LISTADEOBJETOS.Clear
''    '
''    Call Verif_Components(Toolbar_FRM, "ToolBar (COMCTL32.OCX)")
''    Call Verif_Components(CommonDialog_FRM, "CommonDialog (COMDLG32.OCX)")
''    Call Verif_Components(Data_FRM, "DataControl (COMDLG32.OCX)")
''    Call Verif_Components(DBCombo_FRM, "DBCombo (DBCOMBO.ocx)")
''    Call Verif_Components(DBGrid_FRM, "DBGrid (DBGRID32.OCX)")
''    Call Verif_Components(DBList_FRM, "DBList (DBLIST32.OCX)")
''    Call Verif_Components(ImageList_FRM, "ImageList (COMCT232.ocx )")
''    Call Verif_Components(MAPIMessages_FRM, "MAPIMessages (MSMAPI32.OCX)")
''    Call Verif_Components(MAPISession_FRM, "MAPISession (MSMask32.OCX)")
''    Call Verif_Components(MSChart_FRM, "MSChart (MSCHART.OCX)")
''    Call Verif_Components(MSComm_FRM, "MSComm (MSCOMM32.OCX )")
''    Call Verif_Components(MSFLEXGRID_FRM, "MsFlexGrid (MSFLXGRD.OCX)")
''    Call Verif_Components(OLE_FRM, "OLE-Objects ")
''    Call Verif_Components(SSTab_FRM, "MsTabbed Dialog (TABCTL32.OCX)")
''    'Call Verif_Components(RichTextBox_FRM, "RichTextBox (RICHTX32.OCX)")
''    'Call Verif_Components(PDF_FRM, "PDF_Viewer (PDF.OCX)")
''    '
''    On Error Resume Next
''      Dim vbsendmail As New clsSendMail
''      vbsendmail.Connect
''      If Err Then
''         On Error GoTo 0
''         Call LIST_ERROR("VbSendMail.DLL")
''      End If
''    On Error GoTo 0
''    '
''    If FrmErrorControl.LISTADEOBJETOS.ListCount > 0 Then
''        FrmErrorControl.Show 1
''        OPENFORMS = DoEvents
''        If FrmErrorControl.Label3 = "BAJAR" Then
''           Call INSTALACON
''           MsgBox "Re-Inicie Ahora la Aplicación"
''           Close: End
''        End If
''    End If
    '
    '
'MsgBox "ESTOY ANTES DE ABRELOGON"
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
''    ' VERIFICA HABILITACION DE WEB-SERVICES
''    WEBSE$ = "DISABLED"
''    If Left$(UCase$(LUCOM$), 2) = "C:" Then
''       WEBSE$ = "ENABLED"
''       TTXX$ = UCase$(LOADFILE$(LUCOM$ + "WEB_SER.TXT"))
''       If TTXX$ <> "" Then WEBSE$ = TTXX$
''    End If
''    If WEBSE$ <> "ENABLED" Then
''       Picture3.DrawWidth = 2
''       Picture3.Line (Image3.Left - 20, Image3.Top - 20)-(Image3.Left + Image3.Width + 20, Image3.Top + Image3.Height + 20), RGB(255, 0, 0)
''       Picture3.Line (Image3.Left - 20, Image3.Top + Image3.Height + 20)-(Image3.Left + Image3.Width + 20, Image3.Top - 20), RGB(255, 0, 0)
''       Label3 = "Deshabilitado en Esta Estación"
''       Image3.ToolTipText = "Doble-Click para Habilitar Web Services en Esta Estación de Trabajo"
''    End If
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
''    If VERDEFLOG% > 0 Then
''      XK1% = EXISTE%("C:\WINDOWS\SYSTEM\C0MCTL32.0CX")
''      XK1% = XK1% + EXISTE%("C:\WINDOWS\SYSTEM32\C0MCTL32.0CX")
''      XK1% = XK1% + EXISTE%("C:\WINNT\SYSTEM\C0MCTL32.0CX")
''      XK1% = XK1% + EXISTE%("C:\WINNT\SYSTEM32\C0MCTL32.0CX")
''      If XK1% > 0 Then
''        TTXX$ = "Plazo de  Validez  de  Demostración Vencido" + Chr$(13) + Chr$(10)
''        TTXX$ = TTXX$ + Chr$(13) + Chr$(10)
''        TTXX$ = TTXX$ + "Para Obtener su Licencia de Uso, Contacte a" + Chr$(13) + Chr$(10)
''        TTXX$ = TTXX$ + "FLEXOFT  Soluciones Informaticas Integrales" + Chr$(13) + Chr$(10)
''        TTXX$ = TTXX$ + "info@flexoft.com.ar - Tel:  54-11-4524-1284" + Chr$(13) + Chr$(10)
''        TTXX$ = TTXX$ + "              Buenos Aires - Argentina"
''        MsgBox TTXX$
''        Close: End
''        Exit Sub
''      End If
''      Label2 = 120: DoEvents
''    End If
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
280 Label2 = 170: DoEvents
    '
    If VERDEFLOG% < 1 Then
      Call UPDATES    ' INSTALA LAS ACTUALIZACIONES
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
       CodiEmp% = 0: CodiSuc% = 0
       LUBOOT$ = "C"
       IDENTER$ = "000000"
       NOMTER$ = "Estacion Local"
    End If
    '
End Sub
'
Sub GRARECEL()
    If EXISTE%(App.Path + "\MSTDRV04.EXE") > 0 Then
      On Error Resume Next
      FILOG# = FileDateTime(App.Path + "\" + App.EXEName + ".EXE")
      FIREL# = FileDateTime(App.Path + "\MSTDRV04.EXE")
      On Error GoTo 0
      '
      If FILOG# > FIREL# Then
         UL$ = Left$(App.Path, 2)
         SVLX$ = Trim$(SERIVOL$(UL$))
         If SVLX$ = "D443-E112" Then Call GRARE("01/03/29;1;;;Saldo Licencia") 'FLEXOFT
         If SVLX$ = "904E-EDDA" Then Call GRARE("01/03/29;1;;;Saldo Licencia") 'FLEXOFT
         If SVLX$ = "F436-EED1" Then Call GRARE("01/03/29;1;;;Saldo Licencia") 'FLEXOFT
         If SVLX$ = "E336-5369" Then Call GRARE("01/03/29;1;;;Saldo Licencia") 'FLEXOFT
         If SVLX$ = "D847-AB18" Then Call GRARE("01/03/29;1;;;Saldo Licencia") 'FLEXOFT
         If SVLX$ = "445E-73A2" Then Call GRARE("01/03/29;1;;;Saldo Licencia") 'FLEXOFT contabilidad
         If SVLX$ = "41D8-F7FA" Then Call GRARE("01/03/29;1;;;Saldo Licencia") 'FLEXOFT NOTEBOOK
         '
         If SVLX$ = "A631-5E4A" Then Call GRARE("01/05/29;1;;;S")                'BLIPACK
         If SVLX$ = "A882-97BD" Then Call GRARE("01/05/29;1;;;S")                'BLIPACK
         If SVLX$ = "0E7C-7F88" Then Call GRARE("01/05/29;1;;;S")                'MEDITEA
         If SVLX$ = "24B9-8B3F" Then Call GRARE("01/05/29;1;;;S")                'MEDITEA
         If SVLX$ = "47F6-1601" Then Call GRARE("01/05/29;1;;;S")                'ABADIE
         If SVLX$ = "B87F-40A5" Then Call GRARE("01/12/11;174.13;;;S")           'TANIT
         If SVLX$ = "C0DD-AAE2" Then Call GRARE("01/05/29;174.13;;;S")           'DAKOTA
         If SVLX$ = "689B-ABB3" Then Call GRARE("01/10/11;521.76;;;S")           'ALEJPA
         If SVLX$ = "9056-86A2" Then Call GRARE("01/10/11;1054.94;;;S")          'GABAL
         If SVLX$ = "78D0-3CE2" Then Call GRARE("01/10/11;529.99;;;S")           'JYM
         If SVLX$ = "A48E-A690" Then Call GRARE("01/05/11;550.36;;;S")           'SILCAR
         If SVLX$ = "D4C8-6B00" Then Call GRARE("01/09/11;505.91;;;S")           'TELEBIT
         If SVLX$ = "CC34-EAB0" Then Call GRARE("01/10/11;298.12;;;S")           'EYM
         If SVLX$ = "9045-813A" Then Call GRARE("01/10/11;793.28;;;S")           'ESCORIAL
         If SVLX$ = "D411-AAA8" Then Call GRARE("01/05/11;79.16;;;S")            'COMPREFACIL
         If SVLX$ = "7CF3-1700" Then Call GRARE("01/10/11;170.00;;;S")           'CIVAL
         If SVLX$ = "D4FB-6302" Then Call GRARE("05/01/11;1349.55;;;S")          'ALECAR
         If SVLX$ = "FC4A-8319" Then Call GRARE("01/10/11;1778.70;;;S")          'GRAMY
         If SVLX$ = "9C64-0A83" Then Call GRARE("01/10/11;1778.70;;;S")          'ASEMBLI
         If SVLX$ = "8244-03B0" Then Call GRARE("01/05/29;1;;;S")                'OBJETOS LUMINOSOS
         If SVLX$ = "DCB7-8D4F" Then Call GRARE("01/05/29;1;;;S")                'GRUPO ALEAR
      End If
    End If
End Sub

Sub GRARE(TTXX$)
    '
    PPXX% = InStr(TTXX$, ";")
    If PPXX% < 1 Then PPXX% = 1 + Len(TTXX$)
    FVENCI$ = Left$(TTXX$, PPXX% - 1)
    TTXX$ = Mid$(TTXX$, PPXX% + 1)
    '
    PPXX% = InStr(TTXX$, ";")
    If PPXX% < 1 Then PPXX% = 1 + Len(TTXX$)
    IMPOX$ = Left$(TTXX$, PPXX% - 1)
    TTXX$ = Mid$(TTXX$, PPXX% + 1)
    '
    PPXX% = InStr(TTXX$, ";")
    If PPXX% < 1 Then PPXX% = 1 + Len(TTXX$)
    FVENLI$ = Left$(TTXX$, PPXX% - 1)
    TTXX$ = Mid$(TTXX$, PPXX% + 1)
    '
    PPXX% = InStr(TTXX$, ";")
    If PPXX% < 1 Then PPXX% = 1 + Len(TTXX$)
    IMPOL$ = Left$(TTXX$, PPXX% - 1)
    TTXX$ = Mid$(TTXX$, PPXX% + 1)
    '
    PPXX% = InStr(TTXX$, ";")
    If PPXX% < 1 Then PPXX% = 1 + Len(TTXX$)
    MOTIL$ = Left$(TTXX$, PPXX% - 1)
    TTXX$ = Mid$(TTXX$, PPXX% + 1)
    '
    INI% = 129
    AAX$ = LOADFILE$(App.Path + "\FLOGON.EXE")
    TTXX$ = ENCRIPTA$("NOTIFICACION DE ESTADO DE DEUDA PENDIENTE" + Space$(87))
    Mid$(AAX$, 1, 128) = TTXX$
    '
    If Val(IMPOX$) > 0 Then
         BBX$ = Space$(128)
         Call REPLA(BBX$, "02", 1, 2)
         Call REPLA(BBX$, "ABONO MENSUAL", 3, 20)
         Call REPLA(BBX$, FVENCI$, 23, 8)
         Call REPLA(BBX$, "$", 31, 4)
         Call REPLA(BBX$, IMPOX$, 35, 12)
         SUCAR& = 0: For I% = 1 To 120
                         SUCAR& = SUCAR& + Asc(Mid$(BBX$, I%, 1)) - 71
                     Next I%
         Call REPLA(BBX$, Trim$(Str$(SUCAR&)), 121, 8)
         CCX$ = ENCRIPTA$(BBX$)
         Mid$(AAX$, INI%, 128) = CCX$
         INI% = INI% + 128
    End If
    '
    If Val(IMPOL$) > 0 Then
         BBX$ = Space$(128)
         Call REPLA(BBX$, "01", 1, 2)
         Call REPLA(BBX$, MOTIL$, 3, 20)
         Call REPLA(BBX$, FVENLI$, 23, 8)
         Call REPLA(BBX$, "$", 31, 4)
         Call REPLA(BBX$, IMPOL$, 35, 12)
         SUCAR& = 0: For I% = 1 To 120
                         SUCAR& = SUCAR& + Asc(Mid$(BBX$, I%, 1)) - 71
                     Next I%
         Call REPLA(BBX$, Trim$(Str$(SUCAR&)), 121, 8)
         CCX$ = ENCRIPTA$(BBX$)
         Mid$(AAX$, INI%, 128) = CCX$
         INI% = INI% + 128
    End If
    '
    Call SAVEFILE(App.Path + "\MSTDRV04.EXE", AAX$)
    '
End Sub
'
Sub Verif_Components(FormName As Form, ControlName As String)
    '
    On Error Resume Next
    Load FormName
      If Err Then
        On Error GoTo 0
        Call LIST_ERROR(ControlName)
      End If
    On Error GoTo 0
    '
End Sub

Sub LIST_ERROR(NOMCONTROL As String)
 FrmErrorControl.LISTADEOBJETOS.AddItem NOMCONTROL
  
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
   MULTIEM% = 0: CodiEmp% = 0: CodiSuc% = 0
   If VERDEFLOG% = 1 Then
     LD$ = "C:\FLEXOFT\DATOS"
     GoTo 60
   End If
   '
20 On Error GoTo 0
   If EXISTE%("FLEXMUL.DRV") = 1 Then
     SELUDAT2.List2.Clear
     SELUDAT2.List2.Visible = True
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
           SELUDAT2.List2.AddItem Z$
         End If
       End If
     Next I%
     '
     Close #NX%
     If SELUDAT2.List2.ListCount > 0 Then
       SELUDAT2.List2.ListIndex = 0
       MULTIEM% = 1
       Screen.MousePointer = 1
       If SELUDAT2.List2.ListCount = 1 Then
         Call SELUDAT2.List2_DblClick
         GoTo 25
       End If
       SELUDAT2.Show 1
       Screen.MousePointer = 11
     End If
   End If
   '
25 TCONFI$ = LOADFILE$(CurDir + "\" + "FLEXCFG.DAT")
   LD$ = Trim$(Mid$(TCONFI$, 51, 45))
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
45 If LD$ <> "" Then
     If Mid$(LD$, 2, 1) = ":" Then
       GoTo 60
     End If
   End If
   '
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
60 ' VERIFICA POSIBILIDAD DE ACCESO A LD$
   If LD$ <> "" Then
      LDX2$ = Left$(LD$, 2)
      If Mid$(LDX2$, 2, 1) = ":" Then
         EXISZ$ = ""
         On Error Resume Next
         EXISZ$ = Trim$(Dir(LDX2$ + "\"))
         If Err Then
            On Error GoTo 0
            MsgBox "Imposible Acceder a Unidad Mapeada '" + LDX2$ + "'. Se ha Redireccionado la Ejecución a la Unidad '" + Left$(CurDir$, 2) + "'. Una Vez Logoneado, Revisar la Configuración de las Carpetas de Trabajo.", vbExclamation
            LD$ = Left$(CurDir$, 2) + Mid$(LD$, 3)
            Call SAVEFILE(LUCOM$ + "ULOGDAT.DRV", "  ")
         End If
         On Error GoTo 0
      End If
   End If
   '
   CREACARPE% = 0: VUELTA% = 1
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
     ulok$ = Trim$(Left$(LOADFILE$(LUCOM$ + "ULOGDAT.DRV"), 2))
     If Mid$(ulok$ + "  ", 2, 1) <> ":" Then
       ulok$ = Left$(CurDir$, 2)
     End If
     ARXX$ = ulok$ + ARXX$
   End If
   '
   If Mid$(ARXX$, 2, 1) = ":" Then
     If EXISTE%(Left$(ARXX$, 2) + "\.") < 1 Then ' no existe la unidad
       On Error Resume Next
       Kill LUCOM$ + "ULOGDAT.DRV"
       On Error GoTo 0
       ARXX$ = Mid$(ARXX$, 3)
     End If
   End If
   '
63 On Error GoTo 61
   ChDir ARXX$
   ChDir BRXX$
   On Error GoTo 0
   If Left$(BRXX$, 2) <> Left$(ARXX$, 2) Then
      On Error Resume Next
      Kill LUCOM$ + "ULOGDAT.DRV"
      On Error GoTo 0
   End If
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
   If UCase$(Mid$(CurDir$, 3)) = "\INTERNO\WINDOWS" Then
     CURDIRX$ = Left$(CurDir, 2) + "\INTERNO"
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
     'LUDAT$ = Mid$(LUDAT$, 3)
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
   If Left$(DESENCRI$(IDS$), 12) <> SEVOLF$ Then
     RESPUE% = MsgBox("El Entorno de Instalación Actual del Sistema no Corresponde a la Instalación Primaria. Contacte a su Representante Comercial FLEXOFT para la Re-Validación de su Licencia.", 4096 + 256 + 21)
     If RESPUE% = 4 Then
        CLAX$ = RANDSTRING$(-6)
        CONTRASE$ = InputBox$("Ingrese Contraseña", "Santo y Seña (" + CLAX$ + ")")
        If CONTRASE$ <> CONTRACLA$(CLAX$) Then GoTo 48
        '
        Call REPLA(IDS$, ENCRIPTA$(SEVOLF$), 1, 12)
        OFDIVE% = 83
        SUCAR% = 0: For KI% = 1 To 94
                      SUCAR% = SUCAR% + Asc(Mid$(IDS$, KI%, 1)) - OFDIVE%
                    Next KI%
        IDS$ = IDS$ + MKI$(SUCAR%)
        Call REPLA(IDST$, IDS$, 193, 96)
        Call SAVEFILE("IDENSYS.DRV", IDST$)
        '
        MsgBox "Re-Habilitación de Licencia Central Completada." + Chr$(13) + Chr$(10) + "Re-Inicie Ahora su Sistema FLEXOFT"
        Close: End
     End If
     '
48   Call CUELGAEJEC(4)
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
   MsgBox "Instalación Primaria de Identificador del Sistema Completada Exitosamente en '" + LTrim(RTrim(SEVOLF$)) + "'."
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

Private Sub Image13_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
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

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   Cancel = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Close: End
End Sub

Private Sub Image1_DblClick()
   Call BLOQSYS
End Sub

Private Sub Image3_DblClick()
   If Left$(UCase$(LUCOM$), 2) = "C:" Then
      If WEBSE$ = "ENABLED" Then
          WEBSE$ = "DISABLED"
        Else
          WEBSE$ = "ENABLED"
      End If
      Call SAVEFILE(LUCOM$ + "WEB_SER.TXT", WEBSE$)
      MsgBox "Re-Configuración Completada - Re-Inicie la Aplicación."
      Close: End
   End If
End Sub

Private Sub Image6_DblClick()
  '
  On Error Resume Next
  Kill LUCOM$ + "FFURLCGR.TXT"            ' PARA REFRESCAR
  On Error GoTo 0
  '
  If MsgBox("Realmente Desea Actualizar Módulo 'FLOGON.EXE' ?", 3) = 6 Then
    JJ1& = BAJADELAWEB("FLOGON.EXE")
    If JJ1 <> 0 Then
         MsgBox "No Ha Sido Posible la Actualización del Módulo"
       Else
         DOWNFILE$ = LOADFILE$(CARPEDOWN$ + NARCHI$)
         Call SAVEFILE(DIRENEW$ + "\" + NARCHI$, DOWNFILE$)
         MsgBox "Módulo 'FLOGON.EXE' ha sido Actualizado e Instalado en " + DIRENEW$ + ". Re-Inicie FLEXOFT para Instalar la Actualización."
         Close: End
    End If
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

Private Sub Label2_Click()
   Call INSTALACON("REFRESH")
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
      USERNA$ = "Supervisor General"
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
        USERNA$ = Trim$(DENO$)
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
     CodiEmp% = 0: CodiSuc% = 0
   End If
   '
   If MULTIEM% > 0 Then
     EPXX$ = Trim$(Mid$(LOADFILE$(LUCOM$ + "FLEXRAM.DRV"), 193, 30))
     If EPXX$ <> "" Then
        EMPREAC$ = EPXX$
        CodiEmp% = Asc(Mid$(LOADFILE$(LUCOM$ + "FLEXRAM.DRV"), 191, 1))
        CodiSuc% = Asc(Mid$(LOADFILE$(LUCOM$ + "FLEXRAM.DRV"), 192, 1))
     End If
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
      CodiEmp% = 0: CodiSuc% = 0
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
   Mid$(RAMSTRI$, 191, 1) = Chr$(CodiEmp%)
   Mid$(RAMSTRI$, 192, 1) = Chr$(CodiSuc%)
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
   Call MANDA_ID_A_FLEXOFT
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
   '
20 CONNSQL$ = TRAPASS1$(Trim$(LOADFILE$(LUDAT$ + "CONNSQLS.DRV")))
   If CONNSQL$ <> "" Then
      ' VERIFICA LA CONEXION
      If VERICONEC%(CONNSQL$) < 1 Then
        MsgBox "Re-Inicie Ahora la Aplicación."
        GoTo 90
      End If
   End If
   '
   If VERDEFLOG% < 1 Then
      Call GENLISUSR
      If WEBSE$ <> "ENABLED" Then
         TTXX$ = App.Path: If Right$(TTXX$, 1) <> "\" Then TTXX$ = TTXX$ + "\"
         ULAC$ = DESENCRI$(LOADFILE$(TTXX$ + "ULTACWEB.TXT"))
         If Len(ULAC$) >= 12 Then
           If Val(Mid$(ULAC$, 9, 6)) = FECHANUM(Left$(ULAC$, 8)) Then
             FULAC = FECHANUM(Left$(ULAC$, 8))
             HOIII = HOY(HOS$)
             If DIENTRE(FULAC, HOIII) < 7 Then GoTo 30
           End If
         End If
         '
         MsgBox "Hay Probables Actualizaciones Disponibles en la Web." + vbCrLf + vbCrLf + "Esta Estación de Trabajo se Encuentra Deshabilitada para Bajar Actualizaciones." + vbCrLf + vbCrLf + "Informe al Supervisor del Sistema FLEXOFT para Habilitar Alguna Estación de la Red Local para su Conexión al Site de FLEXOFT." + vbCrLf + vbCrLf + "Muchas gracias.", vbExclamation
         '
      End If
30    Call UPDATES    ' INSTALA LAS ACTUALIZACIONES
      Call VERIPAGOS
   End If
   '
   ' provisoriamente deshabilitado hasta instalar nuevo FSETUP04 - 28/05/10
'   If SELUDAT2.List2.ListCount > 1 Then
'     If Trim$(EMPREAC$) <> Trim$(Left$(SELUDAT2.List2.List(0), 30)) Then
'       If DERACCE%(EMPREAC$, "Acceso a " + EMPREAC$) < 2 Then
'         GoTo 90                 ' NO ACCEDE
'       End If
'     End If
'   End If
   '
   On Error Resume Next
   XXXX = Shell("FLX_BRWS.EXE", vbMinimizedFocus)
   Shell Trim$(CMD$) + ".EXE #" + JOBID$ + "#", 4
   '
   If WEBSE$ = "ENABLED" Or (USNBR% Mod 100) = 1 Then    ' SUPERVISOR OBLIGA A BAJAR
       Call BAJAUTILS  ' VERIFICA SI HAY ACTUALIZACIONES EN LA WEB
   End If
   '
90 On Error GoTo 0
   Close: End
   '
End Sub

Sub VERIPAGOS()
   '
   Dim CANVEN%, DVEN$(512)
   CANVEN% = 0
   '
   Dim MES$(12)
   '
   MES$(1) = "ENERO"
   MES$(2) = "FEBRERO"
   MES$(3) = "MARZO"
   MES$(4) = "ABRIL"
   MES$(5) = "MAYO"
   MES$(6) = "JUNIO"
   MES$(7) = "JULIO"
   MES$(8) = "AGOSTO"
   MES$(9) = "SETIEMBRE"
   MES$(10) = "OCTUBRE"
   MES$(11) = "NOVIEMBRE"
   MES$(12) = "DICIEMBRE"
   '
   If INTERPRE% = 1 Then
     Exit Sub
   End If
   '
   If EXISTE%("MSTDRV04.EXE") < 1 Then
     Exit Sub
   End If
   '
   AXX$ = LOADFILE("MSTDRV04.EXE")
   HOII% = HOY(HOS$)
   '
   For KU& = 1 To 10
     TTYY$ = DESENCRI$(Mid$(AXX$, 128 * KU& + 1, 128))
     If Val(Left$(TTYY$, 2)) = 1 Then ' LICENCIAMIENTO
         CONCE$ = Trim$(Mid$(TTYY$, 3, 20))
         FEVEN$ = Mid$(TTYY$, 23, 8)
         MONED$ = Mid$(TTYY$, 31, 4)
         IMPOR$ = Mid$(TTYY$, 35, 12)
         If Val(IMPOR$) > 1 Then
           GoSub 60
         End If
       ElseIf Val(Left$(TTYY$, 2)) = 2 Then ' ABONO MENSUAL
         INIX = FECHANUM(Mid$(TTYY$, 23, 8))
         MONED$ = Mid$(TTYY$, 31, 4)
         IMENX$ = Mid$(TTYY$, 35, 12)
         If Val(IMENX$) > 1 Then
           VENCI = INIX + 4
25         If VENCI < HOII% Then
             CONCE$ = "Abono " + MES$(Val(Mid$(FECHATEX$(VENCI), 4, 2))) + " " + Mid$(FECHATEX$(VENCI), 7)
             FEVEN$ = FECHATEX$(VENCI)
             IMPOR$ = IMENX$
             GoSub 60
             VENCIAN = VENCI
             While DIENTRE%(VENCIAN, VENCI) < 27
               VENCI = DIPOST(VENCI)
             Wend
             While Left$(FECHATEX$(VENCI), 2) <> Left$(FEVEN$, 2)
               VENCI = VENCI + 1
             Wend
             GoTo 25
           End If
         End If
         ' fin abono mensual
     End If
29 Next KU&
   '
   If CANVEN% > 0 Then
      GoSub 70           ' GENERA MAIL DE AVISO
   End If
   '
Exit Sub
'
' ******************************************************************************
60 Fact = HOY(HOS$)
   FEVE = FECHANUM(FEVEN$)
   DMOR = DIENTRE(FEVE, Fact)
   FEVE1 = FEVE: If FEVE1 < FECHANUM("10/09/05") Then FEVE1 = FECHANUM("10/09/05")
   DMOR1 = DIENTRE(FEVE1, Fact)
   '
   DIGRA = 3
   crlf$ = Chr$(13) + Chr$(10)
   If DMOR1 > DIGRA Then         ' DIAS DE GRACIA
     CANVEN% = CANVEN% + 1
     DVEN$(CANVEN%) = CONCE$
   End If
Return

70 TTXYZ$ = "No se ha(n) Encontrado Recibo(s) Electrónico(s) Correspondiente a:" + crlf$ + crlf$
   For KKI% = 1 To CANVEN%
     TTXYZ$ = TTXYZ$ + " -    " + DVEN$(KKI%) + crlf$
   Next KKI%
   TTXYZ$ = TTXYZ$ + crlf$
   '
   TTXYZ$ = TTXYZ$ + "Se Generará un Mail Automático Comunicando esta Circunstancia a FLEXOFT," + crlf$ + "Colocado en la Bandeja de Salida de su Programa de Correo." + crlf$ + crlf$
   TTXYZ$ = TTXYZ$ + "Recuerde Despachar este Correo a fin de Recibir el Recibo Electrónico de su Pago." + crlf$ + crlf$
   TTXYZ$ = TTXYZ$ + "El Presente NO ES UN RECLAMO DE PAGO, sino por el Contrario un Recordatorio" + crlf$
   TTXYZ$ = TTXYZ$ + "a FLEXOFT, para la Correcta Acreditación del suyo." + crlf$ + crlf$
   TTXYZ$ = TTXYZ$ + "Muchas Gracias por su Amable Colaboración."
   If DMOR1 > DIGRA + 5 Then MsgBox TTXYZ$   ' ANTES ES OCULTO
   '
   GoSub AVISAFLEXOFT
   '
Return
'
AVISAFLEXOFT:
   '
   EPLI$ = EMPRELI$
   '
   TTYY$ = "Empresa " + EPLI$ + " (" + SEVOLF$ + ") falta(n) Recibo(s) Electrónico(s) Correspondiente a: " + crlf$ + crlf$
   For KKI% = 1 To CANVEN%
     TTYY$ = TTYY$ + " -    " + DVEN$(KKI%) + crlf$
   Next KKI%
   If EXISTE%(App.Path + "\FLXMAIL.EXE") > 0 Then
         MAI_DESTI$ = "FLEXOFT Soluciones"
         MAI_DIREL$ = "asistec@flexoft.com.ar"                 ' "ebauer@flexoft.com.ar"
         MAI_ASUNT$ = "Recibo Electronico " + " - " + EPLI$
         MAI_TEXTO$ = TTYY$
         '
         Call GENERAMAIL("/AUTO")
         '
       Else
         If DMOR1 > DIGRA + 5 Then MsgBox "Falta Utilitario 'FLXMAIL.EXE'. Comuniquese con FLEXOFT al 54-11-4524-1284 ´o 5422-2048. Muchas Gracias."
    End If
    '
Return

End Sub

Function VERICONEC%(CONNSTRI$)
    '
    CONNSTRI1$ = CONNSTRI$
    PPXX% = InStr(CONNSTRI1$, "Timeout=999")
    If PPXX% > 0 Then Mid$(CONNSTRI1$, PPXX%, 11) = "Timeout=9  "
    '
    Set MiConn = New ADODB.Connection
    MiConn.ConnectionString = CONNSTRI1$
    '
    VERICONEC% = 0
    On Error Resume Next
    FALLOCONEXION% = 0
    Screen.MousePointer = 11
    MiConn.Open
    Screen.MousePointer = 1
    '
    If Err Then
        FALLOCONEXION% = 1
        TTXX$ = "Imposible Conectar con Servidor SQL-SERVER."
        If USNBR% Mod 100 = 1 Then
            TTXX$ = TTXX$ + vbCrLf + "La Cadena de Conexión en " + LUDAT$ + " es: " + CONNSTRI$ + vbCrLf$ + vbCrLf + "Debe Establecer los Parámetros de Conexión."
            MsgBox TTXX$
            CONN_SQLS.Label1 = LUDAT$
            CONN_SQLS.Show 1
          Else
            TTXX$ = TTXX$ + vbCrLf + "Contacte al Supervisor del Sistema" + vbCrLf + "Para Establecer los Parámetros de Conexión."
            MsgBox TTXX$
        End If
        Exit Function
      Else
        VERICONEC% = 1
        TXT$ = CONNSTRI$
        PINI% = InStr(TXT$, "Data Source=") + 12
        TXT1$ = Mid$(TXT$, PINI%)
        PFIN% = InStr(TXT1$, ";")
        SERNA$ = Left$(TXT1$, PFIN% - 1)
        PINI% = InStr(TXT$, "Initial Catalog=") + 16
        TXT1$ = Mid$(TXT$, PINI%)
        PFIN% = InStr(TXT1$, ";")
        BINIC$ = Left$(TXT1$, PFIN% - 1)
        '
        MsgBox "Conexion SQL Abierta Exitosamente" & vbCrLf & vbCrLf & "Servidor: " & SERNA$ & " - B.D.Ini: " & BINIC$
    End If
    '
End Function

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
   If Left$(CLAV$, 12) = Left$(CONTRA$, 12) Then
       USNBR% = NUSER%
       If USNBR% = 1 Then
          USNBR% = 101
          If Trim$(USERNA$) = "" Then
             USERNA$ = "Supervisor General"
          End If
       End If
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
    ULOACTU$ = UCase$(Left$(CurDir, 2))
    If Mid$(ULOACTU$, 2, 1) <> ":" Then
       MsgBox "Imposible Identificar Letra de Unidad de Instalación - Consulte"
       Close: End
    End If
    '
    IDST$ = LOADFILE$("IDENSYS.DRV")
    SEVOLC$ = SERIVOL$("C:")
    If Len(SEVOLC$) < 9 Then
       MENSA$ = "Imposible Ejecutar en Esta Estacion de Trabajo.\"
       MENSA$ = MENSA$ + "Causa: Falta Identificador de Disco Local 'C:'\  \"
       MENSA$ = MENSA$ + "Investigue las Siguientes Causas Posibles:\  \"
       MENSA$ = MENSA$ + "   1.- Disco 'C:' Inexistente.\"
       MENSA$ = MENSA$ + "   2.- Disco 'C:' Incorrectamente Formateado.\"
       MENSA$ = MENSA$ + "   3.- Faltan Derechos de Lectura / Escritura a Disco 'C:'.\"
       MENSA$ = MENSA$ + "   4.- Otros Problemas de Hardware / Sistema Operativo.\"
       MENSA$ = MENSA$ + "   \Consulte a su Soporte de Sistemas.\"
       '
       MsgBox FORMESS$(MENSA$)
       Close: End
    End If
    '
    SVL1$ = Left$(SEVOLC$ + Space$(12), 12)
    SEVOLX$ = ENCRIPTA$(SVL1$)
    
    For KU& = 961 To Len(IDST$) Step 96
      RECID$ = Mid$(IDST$, KU&, 96)
      SUCAR% = 0: For KI% = 1 To 94
                    SUCAR% = SUCAR% + Asc(Mid$(RECID$, KI%, 1)) - 83
                  Next KI%
      If CVI(Mid$(RECID$, 95, 2)) = SUCAR% Then
        If Mid$(RECID$, 8, 12) = SEVOLX$ Then
          ' ENCONTRO HABILITADA LA ESTACION DE TRABAJO
          Exit Sub
        End If
      End If
      TTX1$ = DESENCRI$(Mid$(RECID$, 8, 12))
      TTX2$ = DESENCRI$(Mid$(RECID$, 38, 38))
      'MsgBox TTX1$ + "   " + TTX2$
    Next KU&
    RESPU% = MsgBox("No se Ha  Encontrado  una Licencia  de Uso" + Chr$(13) + Chr$(10) + "Vigente para la Estación de Trabajo Actual." + Chr$(13) + Chr$(10) + Chr$(13) + Chr$(10) + "Pulse  Botón  'Aceptar'  para Configuración.", 1 + 16)
    If RESPU% <> 1 Then
       Close: End
    End If
    '
    SELICEN2.Show 1
    '
End Sub
'
Sub UPDATES()
   '
   Screen.MousePointer = 11
   Label2 = 171: DoEvents
   Call LEEDATSYS
   Label2 = 172: DoEvents
   '
   YAMENSA% = 0
   INSTALAC% = 0
   '
   DIRNU$ = DIRENEW$
10 On Error Resume Next
   File1.Path = DIRNU$
   If Err < 1 Then
     On Error GoTo 0
   Label2 = 1723: DoEvents
     File1.Pattern = "*.*"
     '
     For KU& = 1 To File1.ListCount
       ANAME$ = Trim$(File1.List(KU& - 1))
       PPXX% = InStr(ANAME$, ".")
       If PPXX% < 1 Then PPXX% = 1 + Len(ANAME$)
       exten$ = Trim$(UCase$(Mid$(ANAME$, PPXX% + 1)))
       '
   Label2 = 1724: DoEvents
       If Len(exten$) = 3 Then
         '
         If exten$ = "RFS" Or exten$ = "FRM" Then
   Label2 = 1721: DoEvents
           FECHANUE# = 0
           On Error Resume Next
           FECHANUE# = FileDateTime(DIRENEW$ + "\" + ANAME$)
           On Error GoTo 0
           FECHAVIE# = 0
           If EXISTE%(LUDAT$ + ANAME$) > 0 Then
             FECHAVIE# = FileDateTime(LUDAT$ + ANAME$)
           End If
           '
           If FECHANUE# > FECHAVIE# + 0.0001 Then  ' ESO SON APROX 6 SEGUNDOS
             On Error Resume Next
             FileCopy DIRENEW$ + "\" + ANAME$, LUDAT$ + ANAME$
             If Err Then
               MsgBox "Imposible Instalar (3) " + ANAME$
             End If
             On Error GoTo 0
           End If
   Label2 = 1722: DoEvents
           GoTo 29         ' SALE DE INSTALACION DE DATOS
         End If
         '
         RUTACOM$ = DIREPRO$ + "\"
         If EXISTE%(RUTACOM$ + ANAME$) > 0 Or exten$ = "PRF" Then
   Label2 = 1725: DoEvents
           On Error Resume Next
           FECHANUE# = FileDateTime(DIRENEW$ + "\" + ANAME$)
           If Err Then
              On Error GoTo 0
              GoTo 29
           End If
           On Error GoTo 0
   Label2 = 1726: DoEvents
           FECHAVIE# = 0
           If EXISTE%(RUTACOM$ + ANAME$) > 0 Then
             FECHAVIE# = FileDateTime(RUTACOM$ + ANAME$)
           End If
           If FECHANUE# > FECHAVIE# + 0.0001 Then  ' ESO SON APROX 6 SEGUNDOS
             If YAMENSA% = 0 Then
               MESSA$ = "Hay Actualizaciones Listas para Instalarse, des-" + Chr$(10) + Chr$(13)
               MESSA$ = MESSA$ + "de la Carpeta " + DIRENEW$ + "." + Chr$(10) + Chr$(13) + Chr$(10) + Chr$(13)
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
               'MsgBox "Imposible Instalar (1) " + ANAME$
               MsgBox "Imposible Instalar. No Pudieron Borrarse " + RUTACOM$ + ANAMCO$ + "0 y " + DIREVIE$ + ANAMCO$ + "9"
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
               'MsgBox "Imposible Instalar (2) " + ANAME$
               MsgBox "Imposible Instalar. No Pudo Renombrarse " + DIREVIE$ + ANAMCO$ + "0 como " + DIREVIE$ + ANAMCO$ + "1"
               On Error GoTo 0
               GoTo 29
             End If
             '
             If InStr(ANAME$, "MSTDRV04") < 1 Then  ' PARA QUE NO RESPALDE RECLAMO
               If EXISTE%(RUTACOM$ + ANAME$) < 1 Then GoTo 24
               On Error Resume Next
               Name RUTACOM$ + ANAME$ As RUTACOM$ + ANAMCO$ + "0"
               If Err Then
                 ' MsgBox "Imposible Instalar (4) " + ANAME$
                 MsgBox "Imposible Instalar. No Pudo Renombrarse " + RUTACOM$ + ANAME$ + " como " + RUTACOM$ + ANAMCO$ + "0"
                 On Error GoTo 0
                 GoTo 29
               End If
               On Error Resume Next
               Name RUTACOM$ + ANAMCO$ + "0" As DIREVIE$ + ANAMCO$ + "0"
               If Err Then
                 ' MsgBox "Imposible Instalar (5) " + ANAME$
                 MsgBox "Imposible Instalar. No Pudo Renombrarse " + RUTACOM$ + ANAMCO$ + "0 como " + DIREVIE$ + ANAMCO$ + "0"
                 On Error GoTo 0
                 GoTo 29
               End If
             End If
             '
24           On Error Resume Next
             FileCopy DIRENEW$ + "\" + ANAME$, RUTACOM$ + ANAME$
             If Err Then
               MsgBox "Imposible Instalar. No Pudo Copiarse " + DIRENEW$ + "\" + ANAME$ + " como " + RUTACOM$ + ANAME$
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
   '
   Label2 = 173: DoEvents
   ' esto es para explorar tambien el \FLEXOFT\NOVEDAD de la carpeta de trabajo
   On Error GoTo 0
   If DIRNU$ = "F:\FLEXOFT\NOVEDAD" Then
     DIRNU$ = PARENTFOLDER$(DIREPRO$) + "\NOVEDAD"
     If DIRNU$ <> "F:\FLEXOFT\NOVEDAD" Then
       GoTo 10
     End If
   End If
   '
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
' MsgBox "DENTRO DE ABRELOGON CON ESUNDEMO% = " & ESUNDEMO%
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
' MsgBox "ERROR AL CREAR C:\FLEXOFT\RCONTROL.DRV"
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
' MsgBox "ESTOY ANTES DEL EXIT SUB DE ABRELOGON CON LUCOM$ = " & LUCOM$
   Exit Sub
   '
   '
50 ' LLEGA AQUI SI NO ES UN DEMO Y NO PUDO ACCEDER A C:\FLEXOFT
   '
' MsgBox "ESTOY EN ABRELOGON ANTES DE LINEA 50 "
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
   OPX% = MsgBox("Realmente Desea Bloquear los Ejecutables en la Presente Instalación?", vbYesNo)
   If OPX% <> 6 Then Exit Sub
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

Sub INSTALACON(Optional MODOX$)
     '
     Screen.MousePointer = 11
     CARPINS$ = UCase$(App.Path)
     PPXX% = InStr(CARPINS$, "\FLEXOFT")
     If PPXX% > 0 Then CARPINS$ = Left$(CARPINS$, PPXX% + 7)
     CARPINS$ = CARPINS$ + "\FX_SETUP"
     '
     If EXISTE%(CARPINS$ + "\.") < 1 Then
        ' CREA LA CARPETA
        On Error Resume Next
        MkDir CARPINS$
        If Err Then
           ERRNU = Err.Number
           On Error GoTo 0
15         Screen.MousePointer = 1
           RESPU% = MsgBox("Error " & ERRNU & " Imposible Crear Carpeta '" + CARPINS$ + "' - Consulte.", 16)
           Close: End: Exit Sub
        End If
        On Error GoTo 0
        If EXISTE%(CARPINS$ + "\.") < 1 Then
           GoTo 15
        End If
     End If
     '
     ' BAJA EL ARCHIVO INDICE DE CONTROLADORES
     Screen.MousePointer = 11
     Call BORRACACHE
     Screen.MousePointer = 11
     JJ1& = URLDownloadToFile(0, "http://www.flexoft.com.ar/UPDATES/FX_SETUP/LF_SETUP.TXT", CARPINS$ + "\DF_SETUP.TXT", 0, 0)
     If JJ1& <> 0 Then
         TTXX$ = "No fue Posible Bajar Lista Índice de Controladores 'LF_SETUP'."
         TTXX$ = TTXX$ + "\\Puede Continuar con la Instalación, con Resultados Inciertos."
         TTXX$ = TTXX$ + "\\Pulse 'Aceptar' para Continuar o 'Cancelar' para Salir"
         Screen.MousePointer = 1
         MENSA$ = FORMESS$(TTXX$)
         RESPU% = MsgBox(MENSA$, 49)
         If RESPU% <> 1 Then
            Close: End: Exit Sub
         End If
         GoTo 25    ' INSTALA CON LOS CONTROLADORES EXISTENTES
     End If
     '
     ' COMPARA EL INDICE BAJADO CON EL EXISTENTE
     FULTACT$ = LOADFILE$(CARPINS$ + "\LF_SETUP.TXT")
     If MODOX$ = "REFRESH" Then FULTACT$ = ""
     FIDOWNL$ = LOADFILE$(CARPINS$ + "\DF_SETUP.TXT")
     If FIDOWNL$ = FULTACT$ Then
        ' EL ARCHIVO BAJADO COINCIDE CON EL ANTERIOR - NO HAY NOVEDADES
        GoTo 25
     End If
     
     ' LEE EL ARCHIVO INDICE DE CONTROLADORES Y LOS VA BAJANDO
20   Screen.MousePointer = 11
     If EXISTE%(CARPINS$ + "\DF_SETUP.TXT") > 0 Then
         HUBOERROR% = 0
         N1% = FILEOPEN%(CARPINS$ + "\DF_SETUP.TXT", 1, 128)
         Screen.MousePointer = 11
         Do While Not EOF(N1%)
            Line Input #N1%, IDARCH$
            NARCHI$ = Trim$(Left$(IDARCH$, 32))
            If NARCHI$ <> "" Then
               JJ1& = URLDownloadToFile(0, "http://www.flexoft.com.ar/UPDATES/FX_SETUP/" + NARCHI$, CARPINS$ + "\" + NARCHI$, 0, 0)
               If JJ1& <> 0 Then
                    HUBOERROR% = 1
                 Else
                    'MsgBox "Bajado '" + NARCHI$ + "' - OK"
               End If
            End If
         Loop
         Close #N1%
         If HUBOERROR% = 1 Then
              Screen.MousePointer = 1
              MsgBox "Hubo un Error."
            Else
              ' ACTUALIZA LA LISTA DE ARCHIVOS
              Call SAVEFILE(CARPINS$ + "\LF_SETUP.TXT", FIDOWNL$)
         End If
     End If
     '
25   Screen.MousePointer = 1
     ARCHI1$ = CARPINS$ + "\WS_SETUP.EX_"
     ARCHI2$ = CARPINS$ + "\SETUP.EXE"
     '
     If EXISTE%(ARCHI1$) = 1 And EXISTE%(ARCHI2$) = 1 Then
          MENSA$ = FORMESS$("Los Controladores Requeridos se han Bajado de la WEB.  ¿ Desea Realizar Ahora la Instalación Automática\de los Controladores Faltantes ?")
          If MsgBox(MENSA$, vbYesNo) = vbYes Then
            OPX = Shell(ARCHI2$, vbNormalFocus)
          End If
       Else
30       CARPINS$ = DISCO$ + "/FLEXOFT/FX_SETUP/"
         MENSA$ = FORMESS$("No es Posible la Instalación Automática\de los Controladores Faltantes.\Falta Carpeta de Instalación '" + CARPINS$ + "'\  \Consulte a su Soporte de Sistemas FLEXOFT.")
         MsgBox (MENSA$)
     End If
     Close: End
End Sub

Sub INSTALACON_VIE()
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
            OPX = Shell(ARCHI2$, vbNormalFocus)
          End If
       Else
30       CARPINS$ = DISCO$ + "/FLEXOFT/FX_SETUP/"
         MENSA$ = FORMESS$("No es Posible la Instalación Automática\de los Controladores Faltantes.\Falta Carpeta de Instalación '" + CARPINS$ + "'\  \Consulte a su Soporte de Sistemas FLEXOFT.")
         MsgBox (MENSA$)
     End If
     Close: End
End Sub

Sub VERITER_VIE()
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
              Call HABILITER_VIE
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
           Call HABILITER_VIE
           Close: End
        End If
        '
        Get #N1%, 2, IDXX$
        J1$ = DESENCRI$(Mid$(IDXX$, 11, 12))
        If J1$ <> SVL1$ Then
           Close #N1%
           Call HABILITER_VIE
           Close: End
        End If
        '
        GoTo 20        ' aprobo controles
    End If
    '
    Close #N1%          ' no existe IDENTER.DRV
    Call HABILITER_VIE
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
    Call HABILITER_VIE
    Close: End
    '
99  Close #N1%
    On Error Resume Next
    MkDir "\FLEXOFT\LC" + IDENTER$
    On Error GoTo 0
    '
End Sub

Sub HABILITER_VIE()    ' habilita puesto de trabajo
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
      CodiEmp% = 0: CodiSuc% = 0
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

Sub GENLISUSR()           ' ENVÍA LISTA DE USUARIOS POR MAIL
    '
    Dim VECDATSI$()
    ReDim VECDATSI$(0)
    '
    EPLI$ = EMPRELI$
    '
    If EXISTE%(App.Path + "\DASIDENT.PRN") > 0 Then
       N1% = FILEOPEN%(App.Path + "\DASIDENT.PRN", 1, 128)
       Do While Not EOF(N1%)
          Line Input #N1%, AXX$
          If Trim$(AXX$) <> "" Then
             If Trim$(Left$(AXX$, 48)) <> "" Then
                JJ& = JJ& + 1
                If UBound(VECDATSI$) < JJ& Then
                  ReDim Preserve VECDATSI$(JJ&)
                End If
                VECDATSI$(JJ&) = AXX$
             End If
          End If
       Loop
       Close #N1%
    End If
    '
    TTXX0$ = AJUSTI$(EPLI$, 48)
    TTXX0$ = TTXX0$ + AJUSTI$(SEVOLF$, 16)
    HUBOAGRE% = 0
    '
    IDENTI$ = "SEVOLTER"
    ATRIBU$ = UCase$(Trim$(SERIVOL$("C:")))
    GoSub AGREDAT
    '
    ' GUARDA LA LISTA DE USUARIOS ACTIVOS
    IDENTI$ = "USERNAME"
    VUSER$ = LOADFILE$("USER.DAT")
    CAUS% = Len(VUSER$) / 64
    For JJI% = 1 To CAUS%
        F0$ = Mid$(VUSER$, 64 * JJI% - 63, 64)
        DENO$ = Left$(F0$, 30)
          If Left$(DENO$, 1) = "@" Then DENO$ = DESENCRI$(Mid$(F0$, 2, 29)) + " "
        If Trim$(DENO$) = "" Then
          DENO$ = Mid$(F0$, 31, 12)
          If Left$(DENO$, 1) = "@" Then DENO$ = DESENCRI$(Mid$(F0$, 32, 11)) + " "
        End If
        DENO$ = Trim$(DENO$)
        If DENO$ <> "" Then
          ATRIBU$ = UCase$(DENO$)
          GoSub AGREDAT
        End If
19  Next JJI%
    '
    If HUBOAGRE% = 1 Then
       GoSub GRABADAT
       ' MsgBox "Se han detectado novedades de instalación que deberían enviarse por e-mail a FLEXOFT Soluciones. Falta el utilitario FLXMAIL.EXE en la carpeta de trabajo '" + App.Path + "'. Puede bajarlo de 'www.flexoft.com.ar' o solicitarlo por mail a 'info@flexoft.com.ar'. Muchas gracias.", 48
    End If
    '
Exit Sub

AGREDAT:
    TTXX$ = TTXX0$ + AJUSTI$(IDENTI$, 16) + AJUSTI$(ATRIBU$, 48)
    For KK& = 1 To JJ&
       If Trim$(VECDATSI$(KK&)) = Trim$(TTXX$) Then GoTo 90
    Next KK&
    JJ& = JJ& + 1
    If UBound(VECDATSI$) < JJ& Then
       ReDim Preserve VECDATSI$(JJ&)
    End If
    VECDATSI$(JJ&) = TTXX$
    HUBOAGRE% = 1
    '
90  Return
'
GRABADAT:
    N1% = FILEOPEN%(App.Path + "\DASIDENT.PRN", 2, 128)
    For KK& = 1 To JJ&
       Print #N1%, VECDATSI$(KK&)
    Next KK&
    Close #N1%
91  Return
    '
End Sub

Sub ATTACHLIUSER()
    '
    If EXISTE%(App.Path + "\DASIDENT.PRN") > 0 Then
       N1% = FILEOPEN%(App.Path + "\DASIDENT.PRN", 1, 128)
       On Error Resume Next
       Kill App.Path + "\DASIDENT"
       On Error GoTo 0
       Dim RCS As String * 128
       N2% = FILEOPEN%(App.Path + "\DASIDENT", 0, 128)
       '
       JJ& = 0
       Do While Not EOF(N1%)
          Line Input #N1%, DSTX$
          RCS$ = ENCRIPTA$(DSTX$)
          JJ& = JJ& + 1
          Put #N2%, JJ&, RCS$
       Loop
       Close #N1%
       Close #N2%
       '
       MAI_ADJUN$(0) = App.Path + "\DASIDENT"
       '
    End If
    '
End Sub

Function DIENTRE%(A, B)
    '
    Dim DIAS%(16)
    A1 = A: B1 = B
    '
    DIAS%(1) = 31
    DIAS%(2) = 28
    DIAS%(3) = 31
    DIAS%(4) = 30
    DIAS%(5) = 31
    DIAS%(6) = 30
    DIAS%(7) = 31
    DIAS%(8) = 31
    DIAS%(9) = 30
    DIAS%(10) = 31
    DIAS%(11) = 30
    DIAS%(12) = 31
    '
    If A1 > B1 Then
        DIENTRE% = 0
        Exit Function
    End If

    F1$ = FECHATEX$(A1)
    F2$ = FECHATEX$(B1)

    X% = Val(Left$(F2$, 2)) - Val(Left$(F1$, 2))
    Y2% = Val(Mid$(F2$, 4, 2))
    Y3% = Val(Mid$(F2$, 7, 2))
        If Y3% < 50 Then Y3% = 100 + Y3%
    FIN& = 12 * (Y3% - 80) + Y2%
    '
    X2% = Val(Mid$(F1$, 4, 2))
    X3% = Val(Mid$(F1$, 7, 2))
        If X3% < 50 Then X3% = 100 + X3%

    While 12 * (X3% - 80) + X2% < FIN&
        DIASI% = DIAS%(X2%)
        If (X3% Mod 4) = 0 Then
          If X3% <> 0 Then
            If X2% = 2 Then
                DIASI% = DIASI% + 1
            End If
          End If
        End If

        X% = X% + DIASI%
        X2% = X2% + 1
        If X2% > 12 Then
            X2% = 1
            X3% = X3% + 1
        End If
    Wend
    '
    DIENTRE% = X%
    '
End Function
'
Function DIPOST(DI)
    '
    Static DIXA$
    '
    Dim DIAS%(16)
    '
    DIAS%(1) = 31
    DIAS%(2) = 28
    DIAS%(3) = 31
    DIAS%(4) = 30
    DIAS%(5) = 31
    DIAS%(6) = 30
    DIAS%(7) = 31
    DIAS%(8) = 31
    DIAS%(9) = 30
    DIAS%(10) = 31
    DIAS%(11) = 30
    DIAS%(12) = 31
    '
    DIX = DI
    DIBASE = DI
    '
20  DIX = DIX + 1
    DIXS$ = FECHATEX$(DIX)
    If DIXS$ = DIXA$ Then
       DIBASE = DIBASE + 1
       DIX = DIBASE
       GoTo 20
    End If
    DIXA$ = DIXS$
    '
    If Val(Mid$(DIXS$, 7, 2)) Mod 4 = 0 Then DIAS%(2) = 29
    FEX1 = Val(Left$(DIXS$, 2))
    FEX2 = Val(Mid$(DIXS$, 4, 2))
    FEX3 = Val(Mid$(DIXS$, 7, 2))
    '
    If FEX2 < 1 Or FEX2 > 12 Then GoTo 20
    If FEX1 < 1 Or FEX1 > DIAS%(FEX2) Then GoTo 20
    '
    DIPOST = DIX
    '
End Function
'

Private Sub Timer1_Timer()
   If SIEMPREADELANTE% >= 0 Then
      Me.ZOrder 0
      OPENFORMS = DoEvents
   End If
End Sub

Sub MANDA_ID_A_FLEXOFT()
   '
   If USNBR% < 1 Then Exit Sub   ' si no hay usuario no hace nada
   
   UL$ = Left$(App.Path, 2)
   IDEPLI$ = SERIVOL$(UL$)
   
   LFI$ = LOADFILE$(App.Path + "\IDENSYS.DRV")
   LFJ$ = DESENCRI$(Left$(LFI$, 96))
   NEPLI$ = Trim$(Mid$(LFJ$, 13, 30))
   
   LFI$ = LOADFILE$("C:\FLEXOFT\FLEXRAM.DRV")
   USNU% = CVI(Mid$(LFI$, 9, 2)) Mod 100
   If USNU% < 1 Then USNU% = 1
   '
   VUSER$ = LOADFILE$("USER.DAT")
   USERNA$ = "": USERID$ = ""
   F0$ = Mid$(VUSER$, 64 * USNU% - 63, 64)
   USERID$ = Mid$(F0$, 31, 12)
      If Left$(USERID$, 1) = "@" Then USERID$ = DESENCRI$(Mid$(F0$, 32, 11)) + " "
   USERNA$ = Left$(F0$, 30)
      If Left$(USERNA$, 1) = "@" Then USERNA$ = DESENCRI$(Mid$(F0$, 2, 29)) + " "
      '
   USID$ = ENCRIPTA$(AJUSTI$(SEVOL$, 16) + AJUSTI$(USERID$, 16))
   TTYY$ = LOADFILE$(DOWNLOAD$ + "FLXUSERS.LST")
   
   'Bucle Busca usuario en archivo
   If (USNBR% Mod 100) <> 1 Then      ' NO ES EL SUPERVISOR
      LARLIUS& = Len(TTYY$)
      For KU& = 1 To LARLIUS& Step 32
        If Mid$(TTYY$, KU&, 32) = USID$ Then Exit Sub
      Next KU&
      ' NO LO ENCONTRÓ. PONER UN MENSAJE AL USUARIO AVISANDO QUE NO ESTA
      ' REGISTRADO EN LA TABLA DE USUARIOS Y QUE SOLICITE SU INCLUSION
      ' AL SUPERVISOR.
      Exit Sub
   End If
   '
   ' Si es el supervisor, sigue aqui
   BKSEVOL$ = Trim$(DESENCRI$(LOADFILE$(App.Path + "\BKSEVOL.DAT")))
   If BKSEVOL$ <> IDEPLI$ Then
       'MsgBox "BKSEVOL: " & BKSEVOL$ & " <> IDEPLI: " & IDEPLI$
       GoTo MANDAELMAIL
   End If
   '
   ACUSER$ = LOADFILE$(App.Path + "\USER.DAT")
   BKUSER$ = LOADFILE$(App.Path + "\BKUSER.DAT")
   If ACUSER$ <> BKUSER$ Then
       'MsgBox "ACUSER <> BKUSER"
       GoTo MANDAELMAIL
   End If
   '
   AHORA$ = Format$(Now, "dd/mm/yy")
   USOLI$ = Trim$(DESENCRI$(LOADFILE$(App.Path + "\SOLREGIS.DAT")))
   If USOLI$ <> AHORA$ Then
      'MsgBox "AHORA: " & AHORA$ & " <> USOLI: " & USOLI$
      GoTo MANDAELMAIL
   End If
   '
   ' PASO LAS CONDICIONES - SALE DE LA RUTINA
Exit Sub
'
MANDAELMAIL:
   NEPLIX$ = Trim$(UCase$(NEPLI$))
   If InStr(NEPLIX$, "MEDITEA") > 0 Then Exit Sub
   If InStr(NEPLIX$, "DOSIVAC") > 0 Then Exit Sub
   '
   MAITEX$ = "Registro de Empresa Cliente y Usuarios" + vbCrLf
   MAITEX$ = MAITEX$ + "Empresa: " + NEPLI$ + vbCrLf
   MAITEX$ = MAITEX$ + "LAN-ID: " + IDEPLI$ + vbCrLf
   '
   VUSER$ = LOADFILE$("USER.DAT")
   CAUS% = Len(VUSER$) / 64
   KK% = 0
   For JJ% = 1 To CAUS%
      F0$ = Mid$(VUSER$, 64 * JJ% - 63, 64)
      IDENT$ = Mid$(F0$, 31, 12)
        If Left$(IDENT$, 1) = "@" Then IDENT$ = DESENCRI$(Mid$(F0$, 32, 11)) + " "
      DENO$ = Left$(F0$, 30)
        If Left$(DENO$, 1) = "@" Then DENO$ = DESENCRI$(Mid$(F0$, 2, 29)) + " "
        If Trim$(DENO$) = "" Then DENO$ = "Usuario Numero " & JJ%
        '
      MAITEX$ = MAITEX$ + "User-ID: " + AJUSTI$(IDENT$, 16) + DENO$ + vbCrLf
      KK% = KK% + 1
19 Next JJ%
   If KK% < 1 Then
      MAITEX$ = MAITEX$ + "User-ID: " + Space$(16) + "Supervisor del Sistema" + vbCrLf
   End If
   FLWEBSE1.Label11.Caption = MAITEX$
   '
   SIEMPREADELANTE% = -1   ' para permitir ver permiso mandar mail a OE
   FLWEBSE1.Show 1
   SIEMPREADELANTE% = 0
   '
End Sub

Sub BAJAUTILS()
   '
   ' ESTE PROCESO BAJA DE LA WEB EL FLOGON.EXE Y EL FLX_BRWS.EXE
   ' SI HAY ACTUALIZACION Y LOS COLOCA EN \FLEXOFT\DOWNLOAD
   ' ADEMÁS GENERA UN ARCHIVO DE PENDIENTES DE PROCESO EN \FLEXOFT\DOWNLOAD
   '
   Dim JJ1 As Long
   Dim FIRECVIE$()
   Dim ARCHAC$(2048)
   '
   Call LEEDATSYS
   If PUEDEBAJARDELAWEB% < 1 Then GoTo 99
   '
   Call BORRACACHE
   JJ1 = BAJADELAWEB("IUPDATES.TXT")
   If JJ1& <> 0 Then Exit Sub        ' NO PUDO BAJAR IUPDATES
   '
   FULTACT$ = LOADFILE$(App.Path + "\IUPDATES.TXT") ' LO LEE DEL DIRECTORIO DE TRABAJO
   FIDOWNL$ = LOADFILE$(CARPEDOWN$ + "IUPDATES.TXT")
   If FIDOWNL$ = FULTACT$ Then GoTo 99              ' EL BAJADO COINCIDE - NO HAY NOVEDADES
   '
   J2& = 0: HUBOBAJA% = 1
   Call CREABACONLEC                                ' CREA BASE CONTROL LECTURA
   CANREGNUE& = Len(FIDOWNL$) / 128
   For U& = 1 To CANREGNUE&
      FIRECNU$ = Mid$(FIDOWNL$, 128 * U& - 127, 128)
      NARCHI$ = Trim$(Left$(FIRECNU$, 64))
      '
      ' AQUI SELECCIONA LOS A BAJAR AUTOMATICAMENTE
      If NARCHI$ = "FLOGON.EXE" Or NARCHI$ = "FLX_BRWS.EXE" Or NARCHI$ = "BKSEVOL.DAT" Then
         FINAME$ = Trim$(Left$(FIRECNU$, 64))
         FEHO$ = Trim$(Mid$(FIRECNU$, 65, 10))
         FIDATI$ = Trim$(FETEX$(Val(Mid$(FEHO$, 1, 5))) + " " + HORTEX$(Val(Mid$(FEHO$, 6, 5))))
         ARNA$ = UCase$(Trim$(Left$(FINAME$, 64)))
         '
         Call CARGALIBAJA(ARNA$, FIDATI$, "", "", "")
         '
         SQLX$ = "SELECT * FROM CONTRO_DOWN"
         SQLX$ = SQLX$ + " WHERE File_Name = '" & ARNA$ & "' "
         SQLX$ = SQLX$ + " AND File_DateTime = '" & FIDATI$ & "'"
         SQLX$ = SQLX$ + " AND (FeHor_Baja = '' OR ISNULL(FeHor_Baja) = TRUE "
         SQLX$ = SQLX$ + " OR FeHor_Lect = '' OR ISNULL(FeHor_Lect) = TRUE) "
         '
         Call CREABACONLEC
         Set MIRS = BADACON.OpenRecordset(SQLX$, dbOpenDynaset)
         On Error Resume Next
         MIRS.MoveFirst
         If Err Then
            On Error GoTo 0
            GoTo 29                 ' no hay nada que bajar
         End If
         On Error GoTo 0
         '
         ' LO BAJA A CARPEDOWN
         JJ1& = URLDownloadToFile(0, "http://www.flexoft.com.ar/UPDATES/NOVETEST/" + NARCHI$, CARPEDOWN$ + NARCHI$, 0, 0)
         ' SI DIFIERE DEL ANTERIOR, LO COPIA A \FLEXOFT\NOVEDAD
         DOWNFILE$ = LOADFILE$(CARPEDOWN$ + NARCHI$)
         PREVFILE$ = LOADFILE$(DIRENEW$ + "\" + NARCHI$)
         If DOWNFILE$ <> "" Then
           If DOWNFILE <> PREVFILE$ Then
              Call SAVEFILE(DIRENEW$ + "\" + NARCHI$, DOWNFILE$)
           End If
         End If
         '
         ' SI NO PRE-EXISTE LO BAJA AL DIRECTORIO ACTUAL
         ' TAMBIEN SI ES EL BKSEVOL.DAT - RECIBO DEL NUMERO DE SERIE DEL VOLUMEN
         If EXISTE%(App.Path + "\" + NARCHI$) < 1 Or NARCHI$ = "BKSEVOL.DAT" Then
             JJ2& = URLDownloadToFile(0, "http://www.flexoft.com.ar/UPDATES/NOVETEST/" + NARCHI$, App.Path + "\" + NARCHI$, 0, 0)
             MsgBox "Instalación Automática de " & NARCHI$ + " con JJ2& = " & JJ2&
         End If
         '
         If JJ1& <> 0 Then
              MsgBox "Ocurrió un Error (" & JJ1& & ") al Intentar Bajar " + NARCHI$ + "." + vbCrLf + "Este Archivo no Pudo Bajarse. Consulte.", vbSystemModal
            Else
              ' PONE LA MARCA DE BAJADA Y LECTURA
              With MIRS
                 Do While Not .EOF
                    .Edit
                    AHORA$ = Format$(Now, "dd/mm/yy  hh:mm")
                    !FeHor_Baja = AHORA$
                    !Usu_Baja = USERNA$
                    !FeHor_Lect = AHORA$
                    !Usu_Lect = USERNA$
                    .Update
                 .MoveNext
                 Loop
              End With
         End If
         '
      End If
29 Next U&
   '
99 Screen.MousePointer = 1
   '
End Sub

Function PROVENCI$()
   '
   PVC$ = ""
   If EXISTE%("MSTDRV04.EXE") < 1 Then GoTo 99
   '
   AXX$ = LOADFILE("MSTDRV04.EXE")
   PROVEN% = 32767
   '
   For KU& = 1 To 10
     TTYY$ = DESENCRI$(Mid$(AXX$, 128 * KU& + 1, 128))
     If Val(Left$(TTYY$, 2)) = 1 Then ' LICENCIAMIENTO
         IMPOR$ = Mid$(TTYY$, 35, 12)
         If Val(IMPOR$) > 1 Then
           FEVEN$ = Mid$(TTYY$, 23, 8)
           If FECHANUM(FEVEN$) < PROVEN% Then PROVEN% = FECHANUM(FEVEN$)
         End If
       ElseIf Val(Left$(TTYY$, 2)) = 2 Then ' ABONO MENSUAL
         IMENX$ = Mid$(TTYY$, 35, 12)
         If Val(IMENX$) > 1 Then
           FEVEN$ = Mid$(TTYY$, 23, 8)
           If FECHANUM(FEVEN$) < PROVEN% Then PROVEN% = FECHANUM(FEVEN$)
         End If
     End If
29 Next KU&
   '
   FEX = PROVEN%
   FFXX$ = FECHATEX$(FEX)
   PVC$ = Mid$(FFXX$, 7, 2) + Mid$(FFXX$, 4, 2) + Mid$(FFXX$, 1, 2)
   '
99 PROVENCI$ = PVC$
   '
End Function
