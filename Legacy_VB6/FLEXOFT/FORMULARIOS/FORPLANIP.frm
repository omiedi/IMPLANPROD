VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.1#0"; "COMDLG32.OCX"
Object = "{20C62CAE-15DA-101B-B9A8-444553540000}#1.1#0"; "MSMAPI32.OCX"
Begin VB.Form FORPLANIP 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFE0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Planilla de Producción"
   ClientHeight    =   4065
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   4830
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4065
   ScaleWidth      =   4830
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   540
      Left            =   3990
      TabIndex        =   18
      Top             =   2100
      Width           =   540
   End
   Begin VB.ListBox List2 
      Height          =   840
      Left            =   105
      TabIndex        =   17
      Top             =   3885
      Visible         =   0   'False
      Width           =   3585
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00FFFFE0&
      Caption         =   "Base de Datos Entregas"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2430
      Left            =   210
      TabIndex        =   0
      Top             =   105
      Width           =   3480
      Begin VB.PictureBox Picture1 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   840
         ScaleHeight     =   585
         ScaleWidth      =   2265
         TabIndex        =   9
         Top             =   1575
         Width           =   2325
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
            Height          =   600
            Left            =   0
            Picture         =   "FORPLANIP.frx":0000
            Style           =   1  'Graphical
            TabIndex        =   10
            ToolTipText     =   "Genera Plan Maestro de Entregas"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Búsqueda por Características"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   540
            Left            =   735
            TabIndex        =   11
            Top             =   105
            Width           =   1590
         End
      End
      Begin VB.PictureBox Picture8 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   630
         ScaleHeight     =   585
         ScaleWidth      =   2265
         TabIndex        =   1
         Top             =   1050
         Width           =   2325
         Begin VB.CommandButton Command14 
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
            Left            =   0
            Picture         =   "FORPLANIP.frx":1366
            Style           =   1  'Graphical
            TabIndex        =   2
            ToolTipText     =   "Genera Plan Maestro de Entregas"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label9 
            BackStyle       =   0  'Transparent
            Caption         =   "Generar Plan Ma- estro de Entregas"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   540
            Left            =   680
            TabIndex        =   3
            Top             =   60
            Width           =   1590
         End
      End
      Begin VB.PictureBox Picture14 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   420
         ScaleHeight     =   585
         ScaleWidth      =   2265
         TabIndex        =   4
         Top             =   525
         Width           =   2325
         Begin VB.CommandButton Command22 
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
            Left            =   0
            Picture         =   "FORPLANIP.frx":14B0
            Style           =   1  'Graphical
            TabIndex        =   5
            ToolTipText     =   "Abre Planilla de Producción (Excel)"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label15 
            BackStyle       =   0  'Transparent
            Caption         =   "Abrir (Editar) Pla- nilla Producción"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   540
            Left            =   735
            TabIndex        =   6
            Top             =   60
            Width           =   1590
         End
      End
   End
   Begin MSMAPI.MAPIMessages MAPIMessages1 
      Left            =   1260
      Top             =   0
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
      Left            =   525
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   327680
      DownloadMail    =   -1  'True
      LogonUI         =   -1  'True
      NewSession      =   0   'False
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   327680
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFE0&
      Caption         =   "Información de Compras"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1275
      Left            =   210
      TabIndex        =   13
      Top             =   2625
      Width           =   3480
      Begin VB.PictureBox Picture3 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   630
         ScaleHeight     =   585
         ScaleWidth      =   2265
         TabIndex        =   14
         Top             =   420
         Width           =   2325
         Begin VB.CommandButton Command4 
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
            Left            =   0
            Picture         =   "FORPLANIP.frx":15FA
            Style           =   1  'Graphical
            TabIndex        =   15
            ToolTipText     =   "Recibir Información Remota de Compras"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   "Recibir Informa- ción de Compras"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   540
            Left            =   735
            TabIndex        =   16
            Top             =   60
            Width           =   1590
         End
      End
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "FORPLANIP.frx":1904
      Height          =   49995
      Left            =   105
      OleObjectBlob   =   "FORPLANIP.frx":1914
      TabIndex        =   12
      Top             =   4515
      Visible         =   0   'False
      Width           =   7050
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Excel 5.0;"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   540
      Left            =   420
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   0  'Table
      RecordSource    =   ""
      Top             =   4725
      Visible         =   0   'False
      Width           =   1905
   End
   Begin VB.CommandButton Command29 
      Caption         =   "Setup"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   750
      Left            =   3885
      Picture         =   "FORPLANIP.frx":22C6
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Configuración de Funcionamiento"
      Top             =   1050
      Width           =   750
   End
   Begin VB.CommandButton Command1 
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
      Height          =   750
      Left            =   3885
      Picture         =   "FORPLANIP.frx":2708
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Terminar - Salir de la Aplicación"
      Top             =   210
      Width           =   750
   End
   Begin VB.Image Image2 
      Height          =   510
      Left            =   3280
      Picture         =   "FORPLANIP.frx":2852
      Top             =   3380
      Width           =   2010
   End
End
Attribute VB_Name = "FORPLANIP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
   Cancel = 0
   Call CIERRARCH("*.*")
   Call CONECRUN("PROMRP02", "")
   Call FINAL("")
End Sub

Private Sub Command14_Click()
   '
   Command14.Enabled = False
   Screen.MousePointer = 11
   APLAFA$ = TRIM$(LOADFILE$("C:\FLEXOFT\PLANIFAB.DRV"))
   DERENGLO = Val(TRIM$(CONTROL$("PLANIFAB", "DERENGLO")))
   HARENGLO = Val(TRIM$(CONTROL$("PLANIFAB", "HARENGLO")))
   COLCODIG = Val(TRIM$(CONTROL$("PLANIFAB", "COLCODIG")))
   COLFENTR = Val(TRIM$(CONTROL$("PLANIFAB", "COLFENTR")))
   '
   Data1.DatabaseName = APLAFA$   'ULODATOS$ + "\CLIENTES\ADELSUR\Planifab.xls"
   On Error Resume Next
   Data1.Refresh
   If Err Then
     ERNUI& = Err
     On Error GoTo 0
     If ERNUI& = 3170 Then
         Call MENSERR(24, "No se Encontró el ISAM Instalable")
       Else
         Call MENSERR(24, "Imposible Refrescar.\Planilla Abierta o en Uso Exclusivo\por otro Usuario.")
     End If
     GoTo 99
   End If
   On Error GoTo 0
   '
   Data1.RecordSource = "Hoja1$"
   Data1.Refresh
   '
   JJ& = 0: REBLA$ = REGBLAN$("PROGENT")
   For UROW = DERENGLO To HARENGLO
     DBGrid1.Row = UROW
     DBGrid1.Col = COLCODIG
     CODEX$ = TRIM$(DBGrid1.Text)
     DBGrid1.Col = COLFENTR
     FENTRE$ = TRIM$(DBGrid1.Text)
     If CODEX$ <> "" Or FENTRE$ <> "" Then
       If CODEX$ = "" Or CODINT%(CODEX$) < 1 Then
         Call MENSERR(24, "Codigo Inexistente o no Válido\en Registro " + TRIM$(Str$(UROW)) + ".")
         GoTo 19
       End If
       If FENTRE$ = "" Or FECHANUM(FENTRE$) < 1 Then
         Call MENSERR(24, "Fecha no Válida\en Registro " + TRIM$(Str$(UROW)) + ".")
         GoTo 19
       End If
       CIXI% = CODINT%(CODEX$)
       FEEN% = FECHANUM(FENTRE$)
       If FEEN% < COMESACT% Then
         FEX = COMESACT%
         Call MENSERR(24, "Fecha de Entrega no Válida\en Registro " + TRIM$(Str$(UROW)) + ".\Ajustada a " + FECHATEX$(FEX))
         FEEN% = COMESACT
       End If
       XX$ = REBLA$
       Call REPLA(XX$, MKI$(CIXI%), 1, 2)
       Call REPLA(XX$, MKI$(FEEN%), 3, 2)
       Call REPLA(XX$, MKS$(1), 7, 4)
       Call REPLA(XX$, MKS$(1), 15, 4)
       JJ& = JJ& + 1
       Call GRAREG("PROGENT", XX$, JJ&)
     End If
19 Next UROW
   '
   Call SETULTREG("PROGENT", JJ&)
   Call CIERRARCH("PROGENT")
   Call BAJALDISCO
   Screen.MousePointer = 1
   '
99 Command14.Enabled = True
   Call CONECRUN("PROMRP02", "")
   Call FINAL("")
   '
End Sub

Private Sub Command2_Click()
   FIN& = ULTREG&("MOVISTO")
   'For U& = 1 To FIN&
   '  XX$ = REGLEIDO$("MOVISTO", U&)
   '  Call REPLA(XX$, MKS$(U&), 112, 4)
   '  Call GRAREG("MOVISTO", XX$, U&)
   'Next U&
   'Call CIERRARCH("MOVISTO")
   'MsgBox "TERMINE"
   'Exit Sub
   '
   JJ& = 28377
   For U& = 35073 To FIN&
     Call TRACE(20, U&, FIN&)
     XX$ = REGLEIDO$("MOVISTO", U&)
     Call REPLA(XX$, MKS$(0), 112, 4)
     JJ& = JJ& + 1
     Call GRAREG("MOVISTO", XX$, JJ&)
   Next U&
   Call SETULTREG("MOVISTO", JJ&)
   Call CIERRARCH("MOVISTO")
   Call BAJALDISCO
End Sub

Private Sub Command22_Click()
   '
   Static VECES%
   If Timer < TANTU + 5 Then Exit Sub
   TANTU = Timer
   '
   Command22.Enabled = False
   APLAFA$ = TRIM$(LOADFILE$("C:\FLEXOFT\PLANIFAB.DRV"))
   Call MUESTRADOC(APLAFA$)
   '
99 Command22.Enabled = True
   If COMALTER%("Desea Ahora Refrescar el Plan de Entregas ? \\No, Continuar\Si, Refrescar") = 2 Then
     Call Command14_Click
     Call FINAL("")
   End If
   '
End Sub

Private Sub Command29_Click()
   Command29.Enabled = False
   STPLAFAB.Show 1
   Command29.Enabled = True
   Call FINAL("")
End Sub

Private Sub Command3_Click()
   '
   Command3.Enabled = False
   Call CONECRUN("BATRIB03.EXE", "Búsqueda por Atributos")
   Command3.Enabled = True
   '
End Sub

Private Sub Command4_Click()
   Command4.Enabled = False
   Call BAJMAICOMP
   Call CONECRUN("PROMRP02", "")
   Call FINAL("")
   Command4.Enabled = True
End Sub

Private Sub Form_Load()
   '
   Call VERJOBID(OKEY%)
   If OKEY% < 1 Then Call FINAL("")
   '
   EPAC$ = TRIM$(EMPREAC$)
   If EPAC$ <> "" Then
       Caption = Caption + " - " + EPAC$
   End If
   '
   If UCase$(APLINVO$) = "TRANECOMAI" Then
     Call TRANECOMAI
     Call FINAL("")
   End If
   '
10 If EXISTE%("C:\FLEXOFT\PLANIFAB.DRV") < 1 Then
11   Call COMUNI("Falta Indicar Ruta y Archivo\Correspondiente a Planilla de Fabricación.")
     STPLAFAB.Show 1
     GoTo 10
   End If
   '
   APLAFA$ = TRIM$(LOADFILE$("C:\FLEXOFT\PLANIFAB.DRV"))
   If APLAFA$ = "" Then GoTo 11
   If EXISTE%(APLAFA$) < 1 Then GoTo 11
   '
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Cancel = 0
   Call CIERRARCH("*.*")
   'Call CONECRUN("PROMRP02", "")
   Call FINAL("")
End Sub

Sub MUESTRADOC(DOCUACT$)
   '
   On Error Resume Next
   EXCLDOCS.OLE1.SourceDoc = DOCUACT$
   EXCLDOCS.OLE1.Refresh
   TIDOCU$ = "EXCL"
   '
   UNILOG$ = "": DOCUNE$ = EXCLDOCS.OLE1.SourceDoc
   '
   EXCLDOCS.OLE1.Action = 1
   EXCLDOCS.OLE1.Refresh
   '
   EXCLDOCS.VScroll1.Min = 0
   EXCLDOCS.VScroll1.Value = 0
    ABARRA = EXCLDOCS.OLE1.Height - EXCLDOCS.Picture1.Height
    If ABARRA < 1 Then ABARRA = 1
   EXCLDOCS.VScroll1.Max = ABARRA
   EXCLDOCS.VScroll1.SmallChange = ABARRA / 100
   EXCLDOCS.VScroll1.LargeChange = ABARRA / 10
   '
   EXCLDOCS.HScroll1.Min = 0
   EXCLDOCS.HScroll1.Value = 0
    ABARRA = EXCLDOCS.OLE1.Width - EXCLDOCS.Picture1.Width
    If ABARRA < 1 Then ABARRA = 1
   EXCLDOCS.HScroll1.Max = ABARRA
   EXCLDOCS.HScroll1.SmallChange = ABARRA / 100
   EXCLDOCS.HScroll1.LargeChange = ABARRA / 10
   '
   EXCLDOCS.OLE1.Top = 0
   '
90 EXCLDOCS.Show 1
   'Call CONECRUN("PROMRP02", "")
   Call FINAL("")
   '
99 On Error GoTo 0
   '
End Sub


Sub BAJMAICOMP()
   '
   OPX% = COMALTER%("Desea Refrescar Información\Remota de Compras ?\\Refrescar Compras\Cancelar")
   If OPX% <> 1 Then Exit Sub
   '
   MAPISession1.DownLoadMail = True
   MAPISession1.LogonUI = False
   MAPISession1.Action = 1
   MAPIMessages1.SessionID = MAPISession1.SessionID
   DoEvents
   '
   MAPIMessages1.Fetch
   CAMENSA& = MAPIMessages1.MsgCount
   '
   List2.Clear
   For U& = CAMENSA& To 1 Step -1
     Call TRACE(20, U&, CAMENSA&)
     MAPIMessages1.MsgIndex = U& - 1
     ASUNTO$ = MAPIMessages1.MsgSubject
     If InStr(ASUNTO$, "Informacion de Compras (") > 0 Then
       CLA1$ = Mid$(ASUNTO$, 25, 6)
       CLA2$ = Mid$(ASUNTO$, 32, 6)
       If CLA2$ = CONTRACLA$(CLA1$) Then
         TEXTO$ = MAPIMessages1.MsgNoteText
         List2.Clear
         While Len(TEXTO$) > 0
           PPXX% = InStr(TEXTO$, Chr$(13) + Chr$(10))
           If PPXX% < 1 Then PPXX% = 1 + Len(TEXTO$)
           TTXX$ = Left$(TEXTO$, PPXX% - 1)
           TEXTO$ = Mid$(TEXTO$, PPXX% + 2)
           If TRIM$(TTXX$) <> "" Then List2.AddItem TTXX$
         Wend
         MAPIMessages1.Delete
         GoTo 99
       End If
     End If
     '
89 Next U&
   '
   Call COMUNI("No se Encontró\Nueva Información de Compras")
   '
99 MAPISession1.Action = 2  ' !!!
   Refresh
   DoEvents
   '
   If List2.ListCount < 1 Then Exit Sub
   '
   HOII% = HOY(HOS$)
   For U& = 1 To List2.ListCount
     TTXX$ = List2.List(U& - 1)
     NROC& = Val(Mid$(TTXX$, 9, 8))
     FEMI% = FECHANUM(Mid$(TTXX$, 17, 8))
     NPROL& = Val(Mid$(TTXX$, 25, 8))
     If NPROL& > 0 Then
       If NPROL& < 32767 Then
         NPRO% = NPROL&
         DPRO$ = TRIM$(Mid$(TTXX$, 33, 48))
         CIXI% = 0
         On Error Resume Next
         CIXI% = Val(Mid$(TTXX$, 97, 8))
         On Error GoTo 0
         If CIXI% > 0 Then
           PRUNI# = Val(Mid$(TTXX$, 105, 16))
           FECAN$ = TRIM$(Mid$(TTXX$, 121))
           For KU% = 1 To Len(FECAN$) Step 20
             FESOL% = FECHANUM(Mid$(FECAN$, KU%, 8))
             CASOL = Val(Mid$(FECAN$, KU% + 8, 12))
             If FESOL% < HOII% Then FESOL% = HOII%
             If CASOL >= 0.05 Then
               ENCAO$ = FILTERGETRECORD$("OCOMENCA", 1, MKS$(NROC&))
               ' puede que sea nuevo registro o esté anulada
               If CVS(Left$(ENCAO$, 4)) < 1 Or Asc(Mid$(ENCAO$, 54, 1)) = 9 Then
                   Call REPLA(ENCAO$, MKS$(NROC&), 1, 4)
                     FEX = FEMI%
                     REFE$ = FECHATEX$(FEX) + " - " + DPRO$
                   Call REPLA(ENCAO$, REFE$, 5, 44)
                   Call REPLA(ENCAO$, MKI$(FEMI%), 49, 2)
                   Call REPLA(ENCAO$, MKI$(NPRO%), 51, 2)
                   Call REPLA(ENCAO$, Chr$(0), 53, 1)
                   Call REPLA(ENCAO$, Chr$(0), 54, 1)
                     MONECOI% = 1: MONEMI% = 1: TICAMBX = 1
                   Call REPLA(ENCAO$, Chr$(MONECOI%), 83, 1)
                   Call REPLA(ENCAO$, Chr$(MONEMI%), 84, 1)
                   Call REPLA(ENCAO$, MKS$(TICAMBX), 85, 4)
                     HP = 100 * Val(Left$(Time$, 2)) + Val(Mid$(Time$, 4, 2))
                   Call REPLA(ENCAO$, MKI$(HOII%), 55, 2)
                   Call REPLA(ENCAO$, MKI$(Int(HP)), 57, 2)
                   Call REPLA(ENCAO$, MKI$(USNBR%), 59, 2)
                   '
                   PRIDETA& = 1 + ULTREG&("OCOMDETA")
                   ULTDETA& = PRIDETA&
                   '
                   ' grabar encabezado
                   Call REPLA(ENCAO$, MKS$(PRIDETA&), 155, 4)
                   Call REPLA(ENCAO$, MKS$(ULTDETA&), 159, 4)
                   Call REGAPP("OCOMENCA", ENCAO$)
                   Call CIERRARCH("OCOMENCA")
                   '
                 Else
                   If CVI(Mid$(ENCAO$, 51, 2)) <> NPRO% Then ' no coincide Proveedor
                     Call MENSERR(24, "Imposible Procesar O/C " + TRIM$(Str$(NROC&)) + ".\No Coincide Proveedor.")
                     GoTo 90
                   End If
               End If
               '
               RFF$ = RECFILT$("OCOMDETA", 1, MKS$(NROC&))
               For U1& = 1 To Len(RFF$) Step 4
                 REOCO& = CVS(Mid$(RFF$, U1&, 4))
                 OCDT$ = REGLEIDO$("OCOMDETA", REOCO&)
                 If CVI(Mid$(OCDT$, 9, 2)) = CIXI% Then
                   If CVI(Mid$(OCDT$, 95, 2)) = FESOL% Then
                     If Asc(Mid$(OCDT$, 124, 1)) < 9 Then
                       'COINCIDE O/C, CODIGO Y FECHA
                       GoTo 90
                     End If
                   End If
                 End If
               Next U1&
               OCDT$ = REGBLAN$("OCOMDETA")
               Call REPLA(OCDT$, MKS$(NROC&), 1, 4)
               Call REPLA(OCDT$, MKI$(FEMI%), 5, 2)
               Call REPLA(OCDT$, MKI$(NPRO%), 7, 2)
               Call REPLA(OCDT$, MKI$(CIXI%), 9, 2)
               Call REPLA(OCDT$, UNIMED$(CIXI%), 11, 4)
               Call REPLA(OCDT$, MKS$(CASOL), 51, 4)
               Call REPLA(OCDT$, MKS$(CASOL), 55, 4)
               Call REPLA(OCDT$, UNIMED$(CIXI%), 59, 4)
               Call REPLA(OCDT$, MKS$(1), 63, 4)
               Call REPLA(OCDT$, MKD$(PRUNI#), 67, 8)
               Call REPLA(OCDT$, MKD$(PRUNI#), 79, 8)
               Call REPLA(OCDT$, MKD$(PRUNI# * CASOL), 87, 8)
               Call REPLA(OCDT$, MKI$(FESOL%), 95, 8)
               Call REGAPP("OCOMDETA", OCDT$)
             End If
           Next KU%
         End If
       End If
     End If
90 Next U&
   Call CIERRARCH("OCOMDETA")
   Call CIERRARCH("*.*")
   '
   Call COMUNI("Importación de Compras Completa")
   '
End Sub

Sub TRANECOMAI()
   '
   DESTIMAI$ = LCase$(CONTROL$("NECOMPRA", "DESTINEC"))
10 DESTIMAI$ = TRIM$(LCase$(InputBox("Direccion Electronica Destino", "Transferir Necesidades de Compras", DESTIMAI$)))
   If DESTIMAI$ = "" Then Exit Sub
   '
   If InStr(DESTIMAI$, "@") < 2 Then
     Call MENSERR(24, "Destino no Válido")
     GoTo 10
   End If
   '
   Call GRACONTROL("NECOMPRA", "DESTINEC", DESTIMAI$)
   '
   Screen.MousePointer = 11
   TTYY$ = "": TTXX$ = ""
   Call FILESORT("PLACOMP", "", 1, 3)
   TMT% = Asc(Left$(REGSEL$("TATIMAT"), 1))
   '
   Screen.MousePointer = 11
   CIANT% = 0
   For U& = 1 To ULTREG&("PLACOMP")
     XX$ = REGLEIDO$("PLACOMP", U&)
     CIXI% = CVI(Left$(XX$, 2))
     If TMT% < 1 Or TIMATE%(CIXI%) = TMT% Then
       If CIXI% <> CIANT% Then
         If TTXX$ <> "" Then
           TTYY$ = TTYY$ + TTXX$ + Chr$(13) + Chr$(10)
         End If
         CIANT% = CIXI%
         TTXX$ = AJUSTI$(CODEXT$(CIXI%), 16)
         TTXX$ = TTXX$ + AJUSTI$(DESCRIT0$(CIXI%), 64)
         TTXX$ = TTXX$ + AJUSTI$(UNIMED$(CIXI%), 4)
         TTXX$ = TTXX$ + Space$(10)
         TTXX$ = TTXX$ + FORMATNUM$(CIXI%, "F6.0")
       End If
       ' LONGITUD DE TEXTO QUEDO AJUSTADA A 100 CARACTERES
       FEX = CVI(Mid$(XX$, 3, 2))
       CANEC = CVS(Mid$(XX$, 21, 4))
       TTXX$ = TTXX$ + " " + FECHATEX$(FEX) + FORMATNUM$(CANEC, "F10.1") + " "
     End If
   Next U&
   TTYY$ = TTYY$ + TTXX$ + Chr$(13) + Chr$(10)
   '
   IDXS$ = RANDSTRING$(-6) ' NEGATIVO PARA FORZAR VALORES NUMERICOS
   IDENMESS$ = IDXS$ + "." + CONTRACLA$(IDXS$)
   '
   MAPISession1.DownLoadMail = False
   MAPISession1.LogonUI = False
   MAPISession1.Action = 1
   MAPIMessages1.SessionID = MAPISession1.SessionID
   '
   MAPIMessages1.MsgIndex = -1
   MAPIMessages1.MsgSubject = "Necesidades de Compras (" + IDENMESS$ + ")"
   MAPIMessages1.MsgNoteText = TTYY$
   MAPIMessages1.RecipIndex = 0
   MAPIMessages1.RecipType = 1
   'MAPIMessages1.RecipDisplayName = MUESTRA$
   MAPIMessages1.RecipAddress = DESTIMAI$
   '
   On Error Resume Next
   MAPIMessages1.Action = 3   ' vbMessageSend
   If Err Then
     MsgBox "Ocurrio Error " & Err
     Resume 90
   End If
   '
   Call COMUNI("Necesidades de Compras Transferidas.")
   '
90 On Error GoTo 0
   Screen.MousePointer = 1
   MAPISession1.Action = 2  ' !!!
   DoEvents
   '
End Sub


