VERSION 5.00
Begin VB.Form PROMAT00 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "MRP-II - Programación de Materiales"
   ClientHeight    =   8355
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   11880
   Icon            =   "PROMAT00.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8355
   ScaleWidth      =   11880
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command29 
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
      Height          =   750
      Left            =   10960
      Picture         =   "PROMAT00.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   38
      ToolTipText     =   "Ayuda Flexoft en Línea"
      Top             =   1010
      Width           =   750
   End
   Begin VB.Frame Frame2 
      Caption         =   "Opción"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2280
      Left            =   10920
      TabIndex        =   34
      Top             =   3465
      Width           =   825
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
         Left            =   105
         Picture         =   "PROMAT00.frx":0614
         Style           =   1  'Graphical
         TabIndex        =   37
         ToolTipText     =   "Gestión de Pedidos y Ordenes de Compra a Terceros."
         Top             =   1575
         Width           =   620
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
         Height          =   600
         Left            =   105
         Picture         =   "PROMAT00.frx":091E
         Style           =   1  'Graphical
         TabIndex        =   36
         ToolTipText     =   "Ordenes de Trabajo y Fabricación"
         Top             =   945
         Width           =   620
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
         Left            =   105
         Picture         =   "PROMAT00.frx":0C28
         Style           =   1  'Graphical
         TabIndex        =   35
         ToolTipText     =   "Control y Movimientos de Stocks"
         Top             =   315
         Width           =   620
      End
   End
   Begin VB.Frame Frame13 
      Caption         =   "Files"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1650
      Left            =   10920
      TabIndex        =   31
      Top             =   1785
      Width           =   825
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
         Left            =   105
         Picture         =   "PROMAT00.frx":0F32
         Style           =   1  'Graphical
         TabIndex        =   33
         ToolTipText     =   "Maestro de Clientes y Deudores"
         Top             =   945
         Width           =   620
      End
      Begin VB.CommandButton Command2 
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
         Picture         =   "PROMAT00.frx":1374
         Style           =   1  'Graphical
         TabIndex        =   32
         ToolTipText     =   "Maestro de Items de Stock"
         Top             =   315
         Width           =   620
      End
   End
   Begin VB.Frame Frame7 
      Caption         =   "Programa"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1640
      Left            =   2100
      TabIndex        =   22
      Top             =   105
      Width           =   1380
      Begin VB.CommandButton Command25 
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
         Left            =   630
         Picture         =   "PROMAT00.frx":167E
         Style           =   1  'Graphical
         TabIndex        =   23
         ToolTipText     =   "Listado General Programa de Entregas"
         Top             =   880
         Width           =   600
      End
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
         Left            =   210
         Picture         =   "PROMAT00.frx":1988
         Style           =   1  'Graphical
         TabIndex        =   24
         ToolTipText     =   "ConsultaPrograma por Artículo (Pantalla)"
         Top             =   420
         Width           =   600
      End
   End
   Begin VB.PictureBox MARCOBARRA 
      BackColor       =   &H000000FF&
      Height          =   135
      Left            =   10815
      ScaleHeight     =   75
      ScaleWidth      =   915
      TabIndex        =   16
      Top             =   7350
      Width           =   975
      Begin VB.Frame BARRATRAZA 
         BackColor       =   &H00FF0000&
         BorderStyle     =   0  'None
         Height          =   500
         Left            =   0
         TabIndex        =   17
         Top             =   0
         Width           =   12000
      End
   End
   Begin VB.Frame Frame6 
      Caption         =   "Cálculo de Necesidades"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1640
      Left            =   3570
      TabIndex        =   9
      Top             =   105
      Width           =   3585
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
         Height          =   650
         Left            =   2730
         Picture         =   "PROMAT00.frx":1C92
         Style           =   1  'Graphical
         TabIndex        =   26
         ToolTipText     =   "Configuración - Valores por Omisión"
         Top             =   420
         Width           =   660
      End
      Begin VB.PictureBox Picture2 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   210
         ScaleHeight     =   585
         ScaleWidth      =   2055
         TabIndex        =   10
         Top             =   420
         Width           =   2115
         Begin VB.CommandButton Command7 
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
            Picture         =   "PROMAT00.frx":20D4
            Style           =   1  'Graphical
            TabIndex        =   11
            ToolTipText     =   "Cálculo de Requerimientos de Materiales y Componentes"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   "Re-Cálculo de Necesidades"
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
            TabIndex        =   12
            Top             =   105
            Width           =   1380
         End
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BackColor       =   &H80000016&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
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
         Left            =   1080
         TabIndex        =   27
         Top             =   1200
         Width           =   1065
      End
      Begin VB.Label Label95 
         Alignment       =   2  'Center
         BackColor       =   &H80000016&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
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
         Left            =   2310
         TabIndex        =   21
         Top             =   1200
         Width           =   750
      End
      Begin VB.Label Label121 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "hs"
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
         Left            =   3045
         TabIndex        =   20
         Top             =   1275
         Width           =   330
      End
      Begin VB.Label Label122 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   2175
         TabIndex        =   19
         Top             =   1155
         Width           =   120
      End
      Begin VB.Label Label123 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Ult.Cálc:"
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
         TabIndex        =   18
         Top             =   1275
         Width           =   960
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Consultas y Listados"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1640
      Left            =   7245
      TabIndex        =   3
      Top             =   105
      Width           =   3480
      Begin VB.Frame Frame8 
         BorderStyle     =   0  'None
         Caption         =   "Frame8"
         Height          =   330
         Left            =   630
         TabIndex        =   28
         Top             =   1260
         Width           =   2745
         Begin VB.OptionButton Option2 
            Caption         =   "Impresión"
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
            Left            =   1260
            TabIndex        =   30
            Top             =   0
            Width           =   1275
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Pantalla"
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
            TabIndex        =   29
            Top             =   0
            Value           =   -1  'True
            Width           =   1170
         End
      End
      Begin VB.PictureBox Picture3 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   150
         ScaleHeight     =   585
         ScaleWidth      =   3105
         TabIndex        =   4
         Top             =   420
         Width           =   3165
         Begin VB.CommandButton Command9 
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
            Left            =   1995
            Picture         =   "PROMAT00.frx":23DE
            Style           =   1  'Graphical
            TabIndex        =   7
            ToolTipText     =   "Proyeccion de Compras"
            Top             =   0
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
            Left            =   2520
            Picture         =   "PROMAT00.frx":2820
            Style           =   1  'Graphical
            TabIndex        =   6
            ToolTipText     =   "Proyección de Fabricación"
            Top             =   0
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
            Left            =   0
            Picture         =   "PROMAT00.frx":2B2A
            Style           =   1  'Graphical
            TabIndex        =   5
            ToolTipText     =   "Análisis de Cobertura"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "Necesidades Programadas"
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
            TabIndex        =   8
            Top             =   105
            Width           =   1170
         End
      End
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
      Left            =   10960
      Picture         =   "PROMAT00.frx":2E34
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Cierra y Abandona la Aplicación"
      Top             =   210
      Width           =   750
   End
   Begin VB.Frame Frame1 
      Caption         =   "Entregas Proyectadas"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6420
      Left            =   105
      TabIndex        =   0
      Top             =   1785
      Width           =   10620
      Begin VB.ListBox List1 
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5460
         Left            =   105
         TabIndex        =   1
         Top             =   840
         Width           =   10410
      End
      Begin VB.PictureBox Picture1 
         Height          =   540
         Left            =   105
         ScaleHeight     =   480
         ScaleWidth      =   10350
         TabIndex        =   13
         Top             =   315
         Width           =   10410
         Begin VB.Frame Frame4 
            BorderStyle     =   0  'None
            Caption         =   "Frame4"
            Height          =   225
            Left            =   5490
            TabIndex        =   15
            Top             =   550
            Width           =   4425
         End
         Begin VB.Label Label1 
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   "Codigo          Descripción                      Cta  Cliente / Destino        Cant"
            BeginProperty Font 
               Name            =   "Fixedsys"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   40
            TabIndex        =   14
            Top             =   105
            Width           =   11460
         End
      End
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   840
      Picture         =   "PROMAT00.frx":2F7E
      Top             =   315
      Width           =   480
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "MRP-I"
      BeginProperty Font 
         Name            =   "Benguiat Frisky ATT"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   750
      Left            =   315
      TabIndex        =   25
      Top             =   945
      Width           =   1590
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H000000C0&
      FillColor       =   &H000000C0&
      FillStyle       =   7  'Diagonal Cross
      Height          =   960
      Left            =   210
      Shape           =   2  'Oval
      Top             =   480
      Width           =   1695
   End
   Begin VB.Image Image2 
      Height          =   510
      Left            =   10320
      Picture         =   "PROMAT00.frx":3288
      Top             =   7690
      Width           =   2010
   End
End
Attribute VB_Name = "PROMAT00"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim PROGENT(0 To 32767) As String * 212
Dim CODIPRO%(32767)
Dim CACODIP%
Dim LINVE$(0 To 32767)
Dim SEMI%(-1 To 7)
Dim CAPRO(-1 To 7)

Private Sub Command1_Click()
   '
   Call CIERRARCH("*.*")
   Call BAJALDISCO
   Hide
   '
End Sub

Private Sub Command10_Click()
   Command10.Enabled = False
   'OPPROMRP.Show 1
   Call CONECRUN("FLSETUP/OPPROMRP", "Configuración de Funcionamiento")
   Command10.Enabled = True
End Sub

Private Sub Command12_Click()
   Command12.Enabled = False
   CONSTO00.Show 1
   Command12.Enabled = True
End Sub

Private Sub COMMAND15_CLICK()
    Command15.Enabled = False
    If Option1.Value = True Then
          Call CONECRUN("*COBERFA", "")
        Else
          Call CONECRUN("*SOCOB01", "Listado de Cobertura de Stocks")
    End If
    Command15.Enabled = True
End Sub

Private Sub Command18_Click()
    Command18.Enabled = False
    If App.EXEName = "FLGEST00" Then
        Call CONECRUN("AMACLI02", "Padron Maestro de Clientes")
      Else
        AMACLI00.Show 1
    End If
    Command18.Enabled = True
End Sub

Private Sub Command2_Click()
   Command2.Enabled = False
   AMASTO00.Show 1
   Command2.Enabled = True
End Sub

Private Sub Command25_Click()
    Command25.Enabled = False
    Call CONECRUN("*LPROGEN", "")
    Command25.Enabled = True
End Sub

Private Sub Command29_Click()
    Command29.Enabled = False
    Call HELPONLINE("RESERVA")
    Command29.Enabled = True
End Sub

Private Sub Command3_Click()
   Command3.Enabled = False
   OFABRI00.Show 1
   Command3.Enabled = True
End Sub

Private Sub Command4_Click()
   Command4.Enabled = False
   ORCOMP00.Show 1
   Command4.Enabled = True
End Sub

Private Sub Command6_Click()
    Command6.Enabled = False
    Call FINAL("*CPROGEN")
    Command6.Enabled = True
End Sub

Private Sub Command7_Click()
   FORCANEC.Show 1
End Sub

Private Sub Command8_Click()
    Command8.Enabled = False
    If Option1.Value = True Then
          Call FINAL("*CPLAFAB")
        Else
          Call FINAL("*LPLAFAB")
    End If
    Command8.Enabled = True
End Sub

Private Sub Command9_Click()
    Command9.Enabled = False
    If Option1.Value = True Then
          Call FINAL("*CPLACOM")
        Else
          OPX% = COMALTER%("Elija Secuencia / Agrupamiento\\Por Proveedor\Por Artículo")
          If OPX% = 1 Then
              Call FINAL("*RPLACOM")
            ElseIf OPX% = 2 Then
              Call FINAL("*LPLACOM")
          End If
    End If
    Command9.Enabled = True
End Sub


Private Sub Form_Activate()
   Screen.MousePointer = 1
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   While WindowState <> 0
     WindowState = 0
     Refresh
     OPENFORMS = DoEvents
   Wend
End Sub

Private Sub Form_Load()
   '
   Call CENTRAFORM(Me)
   EPAC$ = TRIM$(EMPREAC$)
   If EPAC$ <> "" Then
      Caption = Caption + "  -  " + EPAC$
   End If
   Call GRATITFOR(Caption)
   '
   FECHOR$ = CONTROL$("CANECOM", "FECHOCAL")
   Label7.Caption = Left$(FECHOR$, 8): Label7.Refresh
   Label95.Caption = Mid$(FECHOR$, 11, 5): Label95.Refresh
   DoEvents
   '
   Call CARPROGENT
   '
   FECHAC% = HOY(HO$)
   'Label2.Caption = METEXTO$(FECHAC%)
   '
End Sub

Private Sub Form_Resize()
   If WindowState = 1 Then
      Call CIERRARCH("*.*")
      Call BAJALDISCO
   End If
   
End Sub

Private Sub Form_Unload(Cancel As Integer)
   '
   Call CIERRARCH("*.*")
   Call BAJALDISCO
   Hide
   '
End Sub

Sub CARPROGENT()
     '
     Screen.MousePointer = 11
     HOII% = HOY(HOS$)
     List1.Clear
125  FIN& = ULTREG&("PROGENT")
     CMA% = COMESACT%
     For I& = 1 To FIN&
       Call TRACE(20, I&, FIN&)
       P00$ = REGLEIDO$("PROGENT", I&)
       If CVI(Mid$(P00$, 3, 2)) < CMA% Then
         Call REPLA(P00$, MKI$(CMA%), 3, 2)
         Call GRAREG("PROGENT", P00$, I&)
       End If
       P1$ = Left$(P00$, 2)
       CI% = CVI(P1$)
       '
       If CI% > 0 Then
         If CI% <= NUMAS% Then
           If STATUIT%(CI%) >= 0 Then
             '
             NC% = CVI(Mid$(P00$, 5, 2))
             CANPRO = CVS(Mid$(P00$, 7, 4))
             If CANPRO >= 0.05 Then
               '
               TTXX$ = Space$(128)
               Call REPLA(TTXX$, CODEXT$(CI%), 1, 16)
               Call REPLA(TTXX$, DESCRIT0$(CI%), 17, 30)
               Call REPLA(TTXX$, CSTRING$(MKI$(NC%), 1, 6, 0, 2), 47, 6)
               Call REPLA(TTXX$, ECOARCH$("DEUDOR1", MKI$(NC%)), 55, 20)
               Call REPLA(TTXX$, CSTRING$(MKS$(CANPRO), 3, 9, 1, 2), 75, 9)
               List1.AddItem TTXX$
               '
             End If
           End If
         End If
       End If
129  Next I&
     '
     Screen.MousePointer = 1
End Sub

Private Sub Label1_DblClick()
   Call ACTPROGENT
End Sub

Private Sub List1_DblClick()
   Call ACTPROGENT
End Sub

Sub ACTPROGENT()
   '
   Call TRALOCAL("PROGENT", "")
   '
   ATRI$ = "/NC"
   ATRI$ = ATRI$ + "/TITUPAN=Actualización Programa de Entregas"
   ATRI$ = ATRI$ + "/BORDESUP=W" + TRIM$(Str$(Top + Frame1.Top + 150))
   ATRI$ = ATRI$ + "/BORDEIZQ=W" + TRIM$(Str$(Left + Frame1.Left))
   '
   VTB% = VENTABU%("CARPROEN" + ATRI$)
   If VTB% >= 0 Then
     Call TRACENTRAL("PROGENT", "PROGENT")
     Call CARPROGENT
   End If
End Sub
