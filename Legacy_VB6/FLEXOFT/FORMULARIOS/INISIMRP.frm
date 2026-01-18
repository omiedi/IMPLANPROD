VERSION 5.00
Begin VB.Form INISIMRP 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Sistema de Programacion de Materiales (MRP)"
   ClientHeight    =   8295
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   11880
   Icon            =   "INISIMRP.frx":0000
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
      Left            =   5760
      Picture         =   "INISIMRP.frx":1D2A
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "Existencias Teóricas Materiales Importación Temporaria"
      Top             =   6060
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
      Left            =   8130
      Picture         =   "INISIMRP.frx":2034
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Actualización de Planes de Fabricación"
      Top             =   3375
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
      Left            =   5535
      Picture         =   "INISIMRP.frx":3666
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Acceso a B.Datos de Fórmulas Tipificadas"
      Top             =   630
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
      Left            =   1680
      Picture         =   "INISIMRP.frx":37B0
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "Emisión de Reportes por Pantalla e Impresora"
      Top             =   6300
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
      Left            =   1725
      Picture         =   "INISIMRP.frx":3ABA
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Actualización y Consulta de Archivos MaestrosTipificados"
      Top             =   300
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
      Begin VB.CommandButton Command8 
         Caption         =   "Compras"
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
         Picture         =   "INISIMRP.frx":3DC4
         Style           =   1  'Graphical
         TabIndex        =   31
         ToolTipText     =   "Gestión de Pedidos y Ordenes de Compra a Terceros."
         Top             =   3840
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
         Picture         =   "INISIMRP.frx":40CE
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Re-Cálculo Consumos de I.T."
         Top             =   2970
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
         Picture         =   "INISIMRP.frx":43D8
         Style           =   1  'Graphical
         TabIndex        =   17
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
         Picture         =   "INISIMRP.frx":46E2
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
         Picture         =   "INISIMRP.frx":49EC
         Style           =   1  'Graphical
         TabIndex        =   16
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
         Picture         =   "INISIMRP.frx":4E2E
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
         Picture         =   "INISIMRP.frx":5138
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
      Left            =   1800
      Picture         =   "INISIMRP.frx":5282
      ScaleHeight     =   1170
      ScaleWidth      =   2115
      TabIndex        =   4
      Top             =   6540
      Width           =   2115
      Begin VB.Label Label7 
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
         TabIndex        =   26
         Top             =   360
         Width           =   1590
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "y Listados"
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
         TabIndex        =   25
         Top             =   590
         Width           =   1590
      End
   End
   Begin VB.PictureBox Picture4 
      BorderStyle     =   0  'None
      Height          =   1170
      Left            =   5820
      Picture         =   "INISIMRP.frx":CC48
      ScaleHeight     =   1170
      ScaleWidth      =   2115
      TabIndex        =   3
      Top             =   6330
      Width           =   2115
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Importación"
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
         Top             =   365
         Width           =   1590
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Temporaria"
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
         Top             =   590
         Width           =   1590
      End
   End
   Begin VB.PictureBox Picture3 
      BorderStyle     =   0  'None
      Height          =   1170
      Left            =   8280
      Picture         =   "INISIMRP.frx":1460E
      ScaleHeight     =   1170
      ScaleWidth      =   2220
      TabIndex        =   2
      Top             =   3645
      Width           =   2220
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Fabricación"
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
         TabIndex        =   24
         Top             =   555
         Width           =   1800
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Planes de"
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
         Left            =   120
         TabIndex        =   23
         Top             =   315
         Width           =   1800
      End
   End
   Begin VB.PictureBox Picture2 
      BorderStyle     =   0  'None
      Height          =   1245
      Left            =   5640
      Picture         =   "INISIMRP.frx":1BFD4
      ScaleHeight     =   1245
      ScaleWidth      =   2115
      TabIndex        =   1
      Top             =   900
      Width           =   2115
      Begin VB.Label Label3 
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
         Left            =   220
         TabIndex        =   22
         Top             =   585
         Width           =   1590
      End
      Begin VB.Label Label1 
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
         TabIndex        =   21
         Top             =   340
         Width           =   1590
      End
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   1170
      Left            =   1830
      Picture         =   "INISIMRP.frx":23ACE
      ScaleHeight     =   1170
      ScaleWidth      =   2010
      TabIndex        =   0
      Top             =   570
      Width           =   2010
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Maestros"
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
         Left            =   240
         TabIndex        =   30
         Top             =   600
         Width           =   1590
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Archivos"
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
         TabIndex        =   29
         Top             =   345
         Width           =   1590
      End
   End
   Begin VB.PictureBox Picture6 
      BorderStyle     =   0  'None
      Height          =   1485
      Left            =   120
      Picture         =   "INISIMRP.frx":2B760
      ScaleHeight     =   1485
      ScaleWidth      =   7170
      TabIndex        =   5
      Top             =   3480
      Width           =   7170
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "de Materiales"
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
         TabIndex        =   20
         Top             =   1005
         Width           =   2220
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Programación"
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
         TabIndex        =   19
         Top             =   750
         Width           =   2220
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Sistemas  de"
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
         TabIndex        =   18
         Top             =   495
         Width           =   2220
      End
   End
   Begin VB.Line Line7 
      X1              =   3840
      X2              =   6720
      Y1              =   3960
      Y2              =   1440
   End
   Begin VB.Line Line6 
      X1              =   2760
      X2              =   3735
      Y1              =   1200
      Y2              =   3975
   End
   Begin VB.Line Line5 
      X1              =   3840
      X2              =   2760
      Y1              =   3960
      Y2              =   7200
   End
   Begin VB.Line Line4 
      X1              =   3840
      X2              =   9360
      Y1              =   4200
      Y2              =   4200
   End
   Begin VB.Line Line1 
      X1              =   3960
      X2              =   6960
      Y1              =   3945
      Y2              =   7080
   End
End
Attribute VB_Name = "INISIMRP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub command1_Click()
  Call SAVEFILE(LUCOM$ + "RCONTROL.DRV", "CLOSETAB")
  Call CIERRARCH("*.*")
  Call FINAL("")
End Sub

Private Sub Command11_Click()
   Command11.Enabled = False
   Call CONECRUN("ARCHIMA", "Archivos Maestros")
   Command11.Enabled = True
End Sub

Private Sub Command12_Click()
   Command12.Enabled = False
   FORTIPIF.Show 1
   Command12.Enabled = True
End Sub

Private Sub Command13_Click()
   Command13.Enabled = False
   COLISMRP.Show 1
   Command13.Enabled = True
End Sub

Private Sub COMMAND15_CLICK()
   Command15.Enabled = False
   APROGMRP.Show 1
   'OFABRI00.Show 1
   Command15.Enabled = True
End Sub

Private Sub Command18_Click()
   Command18.Enabled = False
   STIMTEMP.Show 1
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
   AVAMRPSB.Show 1
   Command5.Enabled = True
End Sub

Private Sub Command6_Click()
   Command6.Enabled = False
   Call GRAJOBID
   Call BAJRAM
   Shell LUBAS$ + "WINTABU " + "#" + JOBID$ + "#", 1
   Command6.Enabled = True
End Sub

Private Sub Command8_Click()
   Command8.Enabled = False
   Call CONECRUN("ORCOMP03", "Sub-Sistema de Gestion de Compras")
   Command8.Enabled = True
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
     If EXISTE%(APPX$ + ".KKK") > 0 Then GoTo 10
     FMDX = FECHMODU%
     BB1 = HOY(HOS$)
     If DIENTRE%(FMDX, HOY(HOS$)) > 100 Then
       Call SAVEFILE(APPX$ + ".KKK", String$(128, 0))
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


