VERSION 5.00
Begin VB.Form INIFLEX 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "FLEXOFT Industrial"
   ClientHeight    =   8295
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   11880
   Icon            =   "INIFLEX.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   8295
   ScaleWidth      =   11880
   Begin VB.Timer Timer2 
      Interval        =   10000
      Left            =   420
      Top             =   0
   End
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
      Left            =   8085
      Picture         =   "INIFLEX.frx":1D2A
      Style           =   1  'Graphical
      TabIndex        =   28
      ToolTipText     =   "Solicitud Interna de Materiales"
      Top             =   6050
      Visible         =   0   'False
      Width           =   600
   End
   Begin VB.PictureBox Picture9 
      Height          =   2010
      Left            =   315
      ScaleHeight     =   1950
      ScaleWidth      =   2055
      TabIndex        =   27
      Top             =   210
      Visible         =   0   'False
      Width           =   2115
      Begin VB.Image Image1 
         Height          =   1590
         Left            =   0
         Top             =   0
         Width           =   2010
      End
   End
   Begin VB.CommandButton Command10 
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
      Left            =   5775
      Picture         =   "INIFLEX.frx":2A24
      Style           =   1  'Graphical
      TabIndex        =   26
      ToolTipText     =   "Informes de Producción"
      Top             =   7200
      Visible         =   0   'False
      Width           =   600
   End
   Begin VB.CommandButton Command8 
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
      Left            =   9345
      Picture         =   "INIFLEX.frx":2E66
      Style           =   1  'Graphical
      TabIndex        =   24
      ToolTipText     =   "Programación por Pedidos de Clentes"
      Top             =   4410
      Visible         =   0   'False
      Width           =   600
   End
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   0
      Top             =   0
   End
   Begin VB.CommandButton Command18 
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
      Left            =   6510
      Picture         =   "INIFLEX.frx":3170
      Style           =   1  'Graphical
      TabIndex        =   12
      ToolTipText     =   "Gestión de Pedidos y Ordenes de Compra a Terceros."
      Top             =   5040
      Width           =   600
   End
   Begin VB.CommandButton Command15 
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
      Left            =   4150
      Picture         =   "INIFLEX.frx":347A
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "Ordenes de Trabajo y Fabricación"
      Top             =   6195
      Width           =   600
   End
   Begin VB.CommandButton Command12 
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
      Left            =   1680
      Picture         =   "INIFLEX.frx":3784
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "Control y Movimientos de Stocks"
      Top             =   5040
      Width           =   600
   End
   Begin VB.CommandButton Command13 
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
      Left            =   7875
      Picture         =   "INIFLEX.frx":3A8E
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Planificación de Necesidades de Materiales - Técnica MRP-II"
      Top             =   3360
      Width           =   600
   End
   Begin VB.CommandButton Command11 
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
      Left            =   525
      Picture         =   "INIFLEX.frx":3D98
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Actualización y Consulta de Archivos Maestros"
      Top             =   3360
      Width           =   600
   End
   Begin VB.PictureBox Picture8 
      BackColor       =   &H8000000A&
      BorderStyle     =   0  'None
      Height          =   1380
      Left            =   8190
      Picture         =   "INIFLEX.frx":40A2
      ScaleHeight     =   1380
      ScaleWidth      =   1905
      TabIndex        =   7
      Top             =   210
      Width           =   1905
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "A"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   3
         Left            =   160
         TabIndex        =   16
         Top             =   450
         Width           =   225
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "C"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   2
         Left            =   780
         TabIndex        =   15
         Top             =   900
         Width           =   225
      End
      Begin VB.Label Label1 
         BackColor       =   &H80000009&
         BackStyle       =   0  'Transparent
         Caption         =   "D"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   1
         Left            =   1460
         TabIndex        =   14
         Top             =   470
         Width           =   225
      End
      Begin VB.Label Label1 
         BackColor       =   &H80000009&
         BackStyle       =   0  'Transparent
         Caption         =   "P"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   0
         Left            =   800
         TabIndex        =   13
         Top             =   60
         Width           =   225
      End
   End
   Begin VB.PictureBox Picture7 
      BackColor       =   &H00800000&
      Height          =   8600
      Left            =   10650
      ScaleHeight     =   8535
      ScaleWidth      =   1215
      TabIndex        =   6
      Top             =   0
      Width           =   1275
      Begin VB.CommandButton Command9 
         Caption         =   "P-Cap"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   795
         Left            =   180
         Picture         =   "INIFLEX.frx":B834
         Style           =   1  'Graphical
         TabIndex        =   25
         ToolTipText     =   "Planeamiento de Capacidad de Fábrica"
         Top             =   4320
         Width           =   855
      End
      Begin VB.CommandButton Command7 
         Caption         =   "C.R.P."
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   795
         Left            =   180
         Picture         =   "INIFLEX.frx":C476
         Style           =   1  'Graphical
         TabIndex        =   23
         ToolTipText     =   "Programación de Carga de Máquinas"
         Top             =   3480
         Width           =   855
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Costos"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   795
         Left            =   180
         Picture         =   "INIFLEX.frx":C780
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   2640
         Width           =   855
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Util."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   795
         Left            =   180
         Picture         =   "INIFLEX.frx":CA8A
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   7350
         Width           =   855
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Direct"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   795
         Left            =   180
         Picture         =   "INIFLEX.frx":CD94
         Style           =   1  'Graphical
         TabIndex        =   20
         ToolTipText     =   "Acceso Directo a Otras Aplicaciones"
         Top             =   5670
         Width           =   855
      End
      Begin VB.CommandButton Command4 
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
         Height          =   795
         Left            =   180
         Picture         =   "INIFLEX.frx":D09E
         Style           =   1  'Graphical
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   6510
         Width           =   855
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Help"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   795
         Left            =   180
         Picture         =   "INIFLEX.frx":D4E0
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   1050
         Width           =   855
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
         Height          =   795
         Left            =   180
         Picture         =   "INIFLEX.frx":D7EA
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   210
         Width           =   855
      End
   End
   Begin VB.PictureBox Picture5 
      BorderStyle     =   0  'None
      Height          =   1170
      Left            =   6615
      Picture         =   "INIFLEX.frx":D934
      ScaleHeight     =   1170
      ScaleWidth      =   2115
      TabIndex        =   4
      Top             =   5250
      Width           =   2115
   End
   Begin VB.PictureBox Picture4 
      BorderStyle     =   0  'None
      Height          =   1170
      Left            =   7980
      Picture         =   "INIFLEX.frx":1542E
      ScaleHeight     =   1170
      ScaleWidth      =   2115
      TabIndex        =   3
      Top             =   3570
      Width           =   2115
   End
   Begin VB.PictureBox Picture3 
      BorderStyle     =   0  'None
      Height          =   1170
      Left            =   4250
      Picture         =   "INIFLEX.frx":1D0C0
      ScaleHeight     =   1170
      ScaleWidth      =   2010
      TabIndex        =   2
      Top             =   6405
      Width           =   2010
   End
   Begin VB.PictureBox Picture2 
      BorderStyle     =   0  'None
      Height          =   1245
      Left            =   1785
      Picture         =   "INIFLEX.frx":24E8A
      ScaleHeight     =   1245
      ScaleWidth      =   2115
      TabIndex        =   1
      Top             =   5250
      Width           =   2115
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   1170
      Left            =   630
      Picture         =   "INIFLEX.frx":2CAB8
      ScaleHeight     =   1170
      ScaleWidth      =   2010
      TabIndex        =   0
      Top             =   3570
      Width           =   2010
   End
   Begin VB.PictureBox Picture6 
      BorderStyle     =   0  'None
      Height          =   1485
      Left            =   2835
      Picture         =   "INIFLEX.frx":346E6
      ScaleHeight     =   1485
      ScaleWidth      =   5055
      TabIndex        =   5
      Top             =   1575
      Width           =   5055
   End
   Begin VB.Line Line7 
      X1              =   5145
      X2              =   9030
      Y1              =   2310
      Y2              =   4200
   End
   Begin VB.Line Line6 
      X1              =   5250
      X2              =   7665
      Y1              =   2310
      Y2              =   5565
   End
   Begin VB.Line Line5 
      X1              =   5250
      X2              =   5250
      Y1              =   2310
      Y2              =   6300
   End
   Begin VB.Line Line4 
      X1              =   2940
      X2              =   5250
      Y1              =   5565
      Y2              =   2310
   End
   Begin VB.Line Line1 
      X1              =   1575
      X2              =   5250
      Y1              =   4095
      Y2              =   2310
   End
End
Attribute VB_Name = "INIFLEX"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
  Call SAVEFILE(LUCOM$ + "RCONTROL.DRV", "CLOSETAB")
  Call CIERRARCH("*.*")
  Call BAJALDISCO
  Call GENBACKUP
  Call FINAL("")
End Sub

Private Sub Command10_Click()
   Command10.Enabled = False
   Call CONECRUN("IFABRI02", "Informes de Producción")
   Command10.Enabled = True
End Sub

Private Sub Command11_Click()
   Command11.Enabled = False
   ARCHIM00.Show 1
   Command11.Enabled = True
End Sub

Private Sub Command12_Click()
   Command12.Enabled = False
   If VERDEMO% < 1 And EXISTE%("CONSTO02.EXE") > 0 Then
       Call CONECRUN("CONSTO02", "Sistema de Gestión de Stocks")
     Else
       CONSTO00.Show 1
   End If
   Command12.Enabled = True
End Sub

Private Sub Command13_Click()
   Command13.Enabled = False
   If DERACCE%("PROGMRP", "Programación de Materiales MRP II") >= 2 Then
     If VERDEMO% < 1 And EXISTE%("PROMRP02.EXE") > 0 Then
         Call CONECRUN("PROMRP02", "Programación MRP-II")
       Else
         Screen.MousePointer = 11
         PROMAT00.Show 1
     End If
   End If
   Command13.Enabled = True
End Sub

Private Sub Command14_Click()
   Command14.Enabled = False
   Call CONECRUN("SOLIMA02", "Solicitud Interna de Materiales")
   Command14.Enabled = True
End Sub

Private Sub COMMAND15_CLICK()
   Command15.Enabled = False
   If EXISTE%("ORFABR02.EXE") > 0 Then
       Call CONECRUN("ORFABR02", "Ordenes de Fabricación")
     Else
       OFABRI00.Show 1
   End If
   Command15.Enabled = True
End Sub

Private Sub Command18_Click()
   Command18.Enabled = False
   If EXISTE%("ORCOMP03.EXE") > 0 Then
       Call CONECRUN("ORCOMP03", "Sistema de Ordenes de Compra")
     Else
       ORCOMP00.Show 1
   End If
   Command18.Enabled = True
End Sub

Private Sub Command3_Click()
   Command3.Enabled = False
   Call ACCDIR("FLEXO00")
   Command3.Enabled = True
End Sub

Private Sub Command4_Click()
   Command4.Enabled = False
   Call CONECRUN("FLSETUP", "Configuración de Funcionamiento")
   Call FINAL("")
   Command4.Enabled = True
End Sub

Private Sub Command5_Click()
   Command5.Enabled = False
   Call CONECRUN("ANACOS02", "Análisis de Costos")
   Command5.Enabled = True
End Sub

Private Sub Command7_Click()
   Command7.Enabled = False
   ACONEC$ = "CARMAQ02"
   If EXISTE%("CARMAQ03.EXE") > 0 Then ACONEC$ = "CARMAQ03"
   Call CONECRUN(ACONEC$, "Programación de Carga de Máquinas")
   Command7.Enabled = True
End Sub

Private Sub Command8_Click()
   Command8.Enabled = False
   Call CONECRUN("PEDIDO02", "Sistema de Pedidos de Clientes")
   Command8.Enabled = True
End Sub

Private Sub Command9_Click()
   Command9.Enabled = False
   ACONEC$ = "CALCAP04"
   Call CONECRUN(ACONEC$, "Planeamiento de Capacidad de Fabrica")
   Command9.Enabled = True
End Sub

Private Sub Form_Load()
   '
   CCOLO% = 193
   BackColor = RGB(CCOLO%, CCOLO%, CCOLO%)
   Picture1.BackColor = RGB(CCOLO%, CCOLO%, CCOLO%)
   Picture2.BackColor = RGB(CCOLO%, CCOLO%, CCOLO%)
   Picture3.BackColor = RGB(CCOLO%, CCOLO%, CCOLO%)
   Picture4.BackColor = RGB(CCOLO%, CCOLO%, CCOLO%)
   Picture5.BackColor = RGB(CCOLO%, CCOLO%, CCOLO%)
   Picture6.BackColor = RGB(CCOLO%, CCOLO%, CCOLO%)
   Picture8.BackColor = RGB(CCOLO%, CCOLO%, CCOLO%)
   '
   Call VERJOBID(OKEY%)
   If OKEY% < 1 Then Call FINAL("")
   '
   EPAC$ = TRIM$(EMPREAC$)
   If EPAC$ <> "" Then
10   Caption = Caption + "  -  " + EPAC$
     '  APPX$ = App.EXEName
     '  If EXISTE%(APPX$ + ".KKK") > 0 Then GoTo 10
     '  FMDX = FECHMODU%
     '  BB1 = HOY(HOS$)
     '  If DIENTRE%(FMDX, HOY(HOS$)) > 100 Then
     '    Call SAVEFILE(APPX$ + ".KKK", String$(128, 0))
     '  End If
     '
   End If
   Call GRATITFOR(Caption)
   '
   If EXISTE%("IFABRI02.EXE") > 0 Then
      Command10.Visible = True
   End If
   '
   If EXISTE%("SOLIMA02.EXE") > 0 Then
     Command14.Visible = True
   End If
   '
   If EXISTE%("ANACOS02.EXE") > 0 Then
     Command5.Enabled = True
   End If
   '
   If EXISTE%("CARMAQ02.EXE") > 0 Or EXISTE%("CARMAQ03.EXE") > 0 Then
     Command7.Enabled = True
   End If
   '
   If EXISTE%("CALCAP04.EXE") > 0 Then
     Command9.Enabled = True
   End If
   '
   Call PONELOGO
   '
   If VERDEMO% < 1 Then
     If EXISTE%("PEDIDO02.EXE") > 0 Then
       Command8.Visible = True
     End If
   End If
   '
End Sub

Sub PONELOGO()
   '
   RUTARGRA$ = TRIM$(CONTROL$("", "RUTARGRA"))
   If RUTARGRA$ = "" Then RUTARGRA$ = TRIM$(LUDAT$)
   If Right$(RUTARGRA$, 1) <> "\" Then
      RUTARGRA$ = RUTARGRA$ + "\"
   End If
   '
   ISOLOGO$ = TRIM$(CONTROL$("", "ISOLOGO"))
   If ISOLOGO$ <> "" Then
      ARCHIMAG$ = RUTARGRA$ + ISOLOGO$
      If EXISTE%(ARCHIMAG$) = 1 Then
        Image1.Picture = LoadPicture("")
        On Error Resume Next
        Image1.Picture = LoadPicture(ARCHIMAG$)
        If Err Then
          ARXX$ = REPLACAR$(ARCHIMAG$, "\", "/")
          Call MENSERR(24, "Archivo Imagen Isologo\" + ARXX$ + "\No Válido o Inexistente.")
        End If
        On Error GoTo 0
        ' Call PRINTIMAGE("ISOLOGO", ARCHIMAG$)
        If EXISTE%(LUDAT$ + "LOGOEMPR.BMP") < 1 Then ' PARA CASO GENTILI
           Picture9.Visible = True
           IMAGEFILE$ = ARCHIMAG$     ' LUDAT$ + "LOGOEMPR.BMP"
           '
           FORMALTER.Image3.Picture = LoadPicture("")
           On Error Resume Next
           FORMALTER.Image3.Picture = LoadPicture(IMAGEFILE$)
           On Error GoTo 0
           GoTo 12
        End If
        '
      End If
   End If
   '
   Picture9.Visible = False
   If EXISTE%(LUDAT$ + "LOGOEMPR.BMP") > 0 Then
     Picture9.Visible = True
     IMAGEFILE$ = LUDAT$ + "LOGOEMPR.BMP"
     '
     FORMALTER.Image3.Picture = LoadPicture("")
     FORMALTER.Image3.Picture = LoadPicture(IMAGEFILE$)
     '
12   HIMAGEN = FORMALTER.Image3.Picture.Width
     VIMAGEN = FORMALTER.Image3.Picture.Height
     '
     HCUADRO = Picture9.Width
     VCUADRO = Picture9.Height
     '
     REDUH = HIMAGEN / HCUADRO
     REDUV = VIMAGEN / VCUADRO
     REDUC = REDUH: If REDUV > REDUH Then REDUC = REDUV
     If REDUC < 1 Then REDUC = 1
     '
     With Picture9
       .Height = VCUADRO * REDUV / REDUC
       .Width = HCUADRO * REDUH / REDUC
       .Left = .Left + (HCUADRO - Picture9.Width) / 2
         If .Left < 300 Then Left = 300
       .Top = .Top - (VCUADRO - Picture9.Height) / 2
         If .Top < 400 Then .Top = 400
     End With
     '
     Image1.Width = Picture9.Width
     Image1.Height = Picture9.Height
     Image1.Stretch = True
     On Error Resume Next
     Image1.Picture = LoadPicture(IMAGEFILE) ' FORMALTER.Image1.Picture ' , PH1 + BORIZQ, PV1 + BORSUP, ANCHCUA, ALTUCUA
     On Error GoTo 0
     '
     FORMALTER.Image3.Picture = LoadPicture("")
   End If
   '
End Sub
'
Private Sub Form_Resize()
   If WindowState = 1 Then
     Call CIERRARCH("*.*")
     Call BAJALDISCO
   End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call SAVEFILE(LUCOM$ + "RCONTROL.DRV", "CLOSETAB")
   Call CIERRARCH("*.*")
   Cancel = 0
   Call FINAL("")
End Sub

Private Sub Timer1_Timer()
   Static LACTIVA%
   For KKI% = 0 To 3
     Label1(KKI%).FontSize = 12
     Label1(KKI%).FontBold = False
     Label1(KKI%).ForeColor = RGB(0, 0, 0)
     If KKI% = LACTIVA% Then
       Label1(KKI%).FontSize = 12
       Label1(KKI%).FontBold = True
       Label1(KKI%).ForeColor = RGB(255, 0, 0)
     End If
   Next KKI%
   DoEvents
   LACTIVA% = (LACTIVA% + 1) Mod 4
End Sub

Private Sub Timer2_Timer()
Exit Sub
   If EXISTE%("C:\FLEXOFT\PRUEBA.PDF") > 0 Then
     ACROL$ = FILEFIND$("C:", "ACRORD32.EXE", 0)
     If ACROL$ <> "" Then
       RUTOR$ = ACROL$ + " "
       TEXACT$ = "" '  "Adobe"
       RUTAR$ = "C:\FLEXOFT\"
       ANAME$ = RUTOR$ + RUTAR$ + "PRUEBA.PDF"
       RETURNVALUE = Shell(ANAME$, 1)
18     On Error Resume Next
       AppActivate TEXACT$ + " Reader - [PRUEBA.pdf]" ' Adobe Reader"
       'If Err Then
         'Resume 20
20       'On Error GoTo 0
         'TINI = Timer
         While Timer < TINI + 1: DoEvents: Wend
         'GoTo 18
       'End If
       On Error GoTo 0
       '
     End If
   End If
End Sub
