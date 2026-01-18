VERSION 5.00
Begin VB.Form INISIMPT 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Sistema Importación Temporaria"
   ClientHeight    =   8295
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   11880
   Icon            =   "INISIMPT.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   8295
   ScaleWidth      =   11880
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
      Left            =   6615
      Picture         =   "INISIMPT.frx":1D2A
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "Existencias Teóricas Materiales Importación Temporaria"
      Top             =   5210
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
      Left            =   4050
      Picture         =   "INISIMPT.frx":2034
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "B.Datos Materiales Re-Exportados y Prod.en Proceso"
      Top             =   6130
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
      Left            =   1575
      Picture         =   "INISIMPT.frx":233E
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Ingreso de Materiales por Importación Temporaria"
      Top             =   5190
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
      Picture         =   "INISIMPT.frx":2648
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Emisión de Reportes por Pantalla e Impresora"
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
      Picture         =   "INISIMPT.frx":2952
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Actualización y Consulta de Archivos Tipificados"
      Top             =   3300
      Width           =   600
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
      Begin VB.CommandButton Command7 
         Caption         =   "R.U.F."
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
         Picture         =   "INISIMPT.frx":2C5C
         Style           =   1  'Graphical
         TabIndex        =   18
         ToolTipText     =   "Reservado Uso Futuro"
         Top             =   4410
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Calc."
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
         Picture         =   "INISIMPT.frx":2F66
         Style           =   1  'Graphical
         TabIndex        =   17
         ToolTipText     =   "Re-Cálculo Consumos de I.T."
         Top             =   3570
         Width           =   855
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Queries"
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
         Picture         =   "INISIMPT.frx":3270
         Style           =   1  'Graphical
         TabIndex        =   16
         ToolTipText     =   "Consultas y Listados Varios"
         Top             =   7350
         Width           =   855
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Tablas"
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
         Picture         =   "INISIMPT.frx":357A
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "Acceso a Tablas de Validación"
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
         Picture         =   "INISIMPT.frx":3884
         Style           =   1  'Graphical
         TabIndex        =   14
         TabStop         =   0   'False
         ToolTipText     =   "Setup - Configuración de Funcionamiento"
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
         Picture         =   "INISIMPT.frx":3CC6
         Style           =   1  'Graphical
         TabIndex        =   13
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
         Picture         =   "INISIMPT.frx":3FD0
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Cierra y Abandona la Aplicacion"
         Top             =   210
         Width           =   855
      End
   End
   Begin VB.PictureBox Picture5 
      BorderStyle     =   0  'None
      Height          =   1170
      Left            =   6720
      Picture         =   "INISIMPT.frx":411A
      ScaleHeight     =   1170
      ScaleWidth      =   2115
      TabIndex        =   4
      Top             =   5460
      Width           =   2115
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Situación"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   210
         TabIndex        =   27
         Top             =   360
         Width           =   1590
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "de Stocks"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   210
         TabIndex        =   26
         Top             =   590
         Width           =   1590
      End
   End
   Begin VB.PictureBox Picture4 
      BorderStyle     =   0  'None
      Height          =   1170
      Left            =   7980
      Picture         =   "INISIMPT.frx":BAE0
      ScaleHeight     =   1170
      ScaleWidth      =   2115
      TabIndex        =   3
      Top             =   3570
      Width           =   2115
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Consultas"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   210
         TabIndex        =   29
         Top             =   365
         Width           =   1590
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "e Informes"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   210
         TabIndex        =   28
         Top             =   590
         Width           =   1590
      End
   End
   Begin VB.PictureBox Picture3 
      BorderStyle     =   0  'None
      Height          =   1170
      Left            =   4200
      Picture         =   "INISIMPT.frx":134A6
      ScaleHeight     =   1170
      ScaleWidth      =   2220
      TabIndex        =   2
      Top             =   6405
      Width           =   2220
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "y en Proceso"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   105
         TabIndex        =   25
         Top             =   555
         Width           =   1800
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Re-Exportados"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Left            =   105
         TabIndex        =   24
         Top             =   315
         Width           =   1800
      End
   End
   Begin VB.PictureBox Picture2 
      BorderStyle     =   0  'None
      Height          =   1245
      Left            =   1680
      Picture         =   "INISIMPT.frx":1AE6C
      ScaleHeight     =   1245
      ScaleWidth      =   2115
      TabIndex        =   1
      Top             =   5460
      Width           =   2115
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Mat.Primas"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   210
         TabIndex        =   23
         Top             =   580
         Width           =   1590
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Recepciones"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Left            =   210
         TabIndex        =   22
         Top             =   340
         Width           =   1590
      End
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   1170
      Left            =   630
      Picture         =   "INISIMPT.frx":22966
      ScaleHeight     =   1170
      ScaleWidth      =   2010
      TabIndex        =   0
      Top             =   3570
      Width           =   2010
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Tipificadas"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   210
         TabIndex        =   30
         Top             =   580
         Width           =   1590
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Fórmulas"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Left            =   210
         TabIndex        =   31
         Top             =   345
         Width           =   1590
      End
   End
   Begin VB.PictureBox Picture6 
      BorderStyle     =   0  'None
      Height          =   1485
      Left            =   1575
      Picture         =   "INISIMPT.frx":2A5F8
      ScaleHeight     =   1485
      ScaleWidth      =   7890
      TabIndex        =   5
      Top             =   1575
      Width           =   7890
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Temporaria"
         BeginProperty Font 
            Name            =   "Abadi MT Condensed"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   435
         Index           =   2
         Left            =   4515
         TabIndex        =   21
         Top             =   1005
         Width           =   2220
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Importación"
         BeginProperty Font 
            Name            =   "Abadi MT Condensed"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   435
         Index           =   1
         Left            =   4830
         TabIndex        =   20
         Top             =   750
         Width           =   2220
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Sistemas"
         BeginProperty Font 
            Name            =   "Abadi MT Condensed"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   435
         Index           =   0
         Left            =   5145
         TabIndex        =   19
         Top             =   495
         Width           =   2220
      End
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
Attribute VB_Name = "INISIMPT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
  Call SAVEFILE(LUCOM$ + "RCONTROL.DRV", "CLOSETAB")
  Call CIERRARCH("*.*")
  Call FINAL("")
End Sub

Private Sub Command11_Click()
   Command11.Enabled = False
   FORTIPIF.Show 1
   Command11.Enabled = True
End Sub

Private Sub Command12_Click()
   Command12.Enabled = False
   FORIRMP.Show 1
   Command12.Enabled = True
End Sub

Private Sub Command13_Click()
   Command13.Enabled = False
   COLISIMT.Show 1
   Command13.Enabled = True
End Sub

Private Sub COMMAND15_CLICK()
   Command15.Enabled = False
   'OFABRI00.Show 1
   Command15.Enabled = True
End Sub

Private Sub Command18_Click()
   Command18.Enabled = False
   COSTOIMT.Show 1
   Command18.Enabled = True
End Sub

Private Sub Command3_Click()
   Command3.Enabled = False
   Call CONECRUN("FLSETUP/OPMAST00", "Configuración de Funcionamiento")
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
   FORCALIT.Show 1
   Command5.Enabled = True
End Sub

Private Sub Command6_Click()
   Command6.Enabled = False
   Call GRAJOBID
   Call BAJRAM
   Shell LUBAS$ + "WINTABU " + "#" + JOBID$ + "#", 1
   Command6.Enabled = True
End Sub

Private Sub Command7_Click()
   Command7.Enabled = False
   ACONEC$ = "CARMAQ02"
   If EXISTE%("CARMAQ03.EXE") > 0 Then ACONEC$ = "CARMAQ03"
   Call CONECRUN(ACONEC$, "Programación de Carga de Máquinas")
   Command7.Enabled = True
End Sub


Private Sub Command8_Click()
   For U& = 1 To ULTREG&("TIPIFIC")
     XX$ = REGLEIDO$("TIPIFIC", U&)
     COBRU# = CVD(Mid$(XX$, 87, 8))
     If COBRU# >= 0.001 Then
       CONET# = (Int(1000 * CVD(Mid$(XX$, 55, 8)) + 0.5)) / 1000
       MENET# = (Int(1000 * CVD(Mid$(XX$, 63, 8)) + 0.5)) / 1000
       PENET# = (Int(1000 * CVD(Mid$(XX$, 75, 8)) + 0.5)) / 1000
       COBRU# = CONET# + MENET# + PENET#
       Call REPLA(XX$, MKD$(CONET#), 55, 8)
       Call REPLA(XX$, MKD$(MENET#), 63, 8)
       Call REPLA(XX$, MKD$(PENET#), 75, 8)
       Call REPLA(XX$, MKD$(COBRU#), 87, 8)
       Call GRAREG("TIPIFIC", XX$, U&)
     End If
   Next U&
   Call CIERRARCH("TIPIFIC")
   MsgBox "FIN"
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
   '
   Call VERJOBID(OKEY%)
   If OKEY% < 1 Then Call FINAL("")
   '
   EPAC$ = TRIM$(EMPREAC$)
   If EPAC$ <> "" Then
     '
10   Caption = Caption + "  -  " + EPAC$
     APPX$ = App.EXEName
     'If EXISTE%(APPX$ + ".KKK") > 0 Then GoTo 10
     FMDX = FECHMODU%
     BB1 = HOY(HOS$)
     If DIENTRE%(FMDX, HOY(HOS$)) > 100 Then
       ' Call SAVEFILE(APPX$ + ".KKK", String$(128, 0))
     End If
     '
   End If
   Call GRATITFOR(Caption)
   '
End Sub

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


