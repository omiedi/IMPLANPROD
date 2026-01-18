VERSION 5.00
Begin VB.Form OPCOMPRA 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00E0FFC0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Gestión de Compras Industriales"
   ClientHeight    =   3420
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   5025
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3420
   ScaleWidth      =   5025
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame8 
      BackColor       =   &H00E0FFC0&
      Caption         =   "Procesos por Unica Vez"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1275
      Left            =   210
      TabIndex        =   11
      Top             =   1995
      Width           =   3585
      Begin VB.PictureBox Picture3 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   630
         ScaleHeight     =   585
         ScaleWidth      =   2265
         TabIndex        =   12
         Top             =   420
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
            Picture         =   "OPCOMPRA.frx":0000
            Style           =   1  'Graphical
            TabIndex        =   13
            ToolTipText     =   "Acceso al Control General de Gestión Comercial"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "Refrescar Moneda de Valuación"
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
            Left            =   630
            TabIndex        =   14
            Top             =   105
            Width           =   1620
         End
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00E0FFC0&
      Caption         =   "Consultas y Listados"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   210
      TabIndex        =   4
      Top             =   210
      Width           =   3585
      Begin VB.PictureBox Picture8 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   840
         ScaleHeight     =   585
         ScaleWidth      =   2265
         TabIndex        =   5
         Top             =   840
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
            Picture         =   "OPCOMPRA.frx":030A
            Style           =   1  'Graphical
            TabIndex        =   6
            ToolTipText     =   "Emitir o Registrar un Recibo de Cobranza"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label9 
            BackStyle       =   0  'Transparent
            Caption         =   "Recibo Cobranza en Cta. Corriente"
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
            TabIndex        =   7
            Top             =   105
            Width           =   1485
         End
      End
      Begin VB.PictureBox Picture14 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   420
         ScaleHeight     =   585
         ScaleWidth      =   2265
         TabIndex        =   8
         Top             =   420
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
            Picture         =   "OPCOMPRA.frx":0454
            Style           =   1  'Graphical
            TabIndex        =   9
            ToolTipText     =   "Generar una Factura en Cuenta Corriente"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label15 
            BackStyle       =   0  'Transparent
            Caption         =   "Facturación C.Cte"
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
            Left            =   630
            TabIndex        =   10
            Top             =   105
            Width           =   1590
         End
      End
   End
   Begin VB.PictureBox MARCOBARRA 
      BackColor       =   &H000000FF&
      Height          =   135
      Left            =   3990
      ScaleHeight     =   75
      ScaleWidth      =   795
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   2415
      Width           =   855
      Begin VB.Frame BARRATRAZA 
         BackColor       =   &H00FF0000&
         BorderStyle     =   0  'None
         Height          =   500
         Left            =   0
         TabIndex        =   3
         Top             =   0
         Width           =   12000
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
      Height          =   800
      Left            =   3990
      Picture         =   "OPCOMPRA.frx":0896
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Terminar - Salir de la Aplicación"
      Top             =   315
      Width           =   855
   End
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
      Height          =   800
      Left            =   3990
      Picture         =   "OPCOMPRA.frx":09E0
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Ayuda Flexoft en Línea"
      Top             =   1155
      Width           =   855
   End
   Begin VB.Image Image2 
      Height          =   510
      Left            =   3420
      Picture         =   "OPCOMPRA.frx":0CEA
      Top             =   2730
      Width           =   2010
   End
End
Attribute VB_Name = "OPCOMPRA"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
    Call CIERRARCH("*.*")
    Call BAJALDISCO
    Call FINAL("")
End Sub

Private Sub Command14_Click()
   Command14.Enabled = False
   'RECIBO00.Show 1
   Command14.Enabled = True
End Sub


Private Sub Command18_Click()
   Command18.Enabled = False
   Call CONECRUN("LISPRE02", "Listas de Precios de Venta")
   Command18.Enabled = True
End Sub


Private Sub Command22_Click()
    Command22.Enabled = False
    'FACTUR00.Show 1
    Command22.Enabled = True
End Sub

Private Sub Command3_Click()
   Command3.Enabled = False
   FIN& = ULTREG&("OCOMDETA")
   For U& = 1 To FIN&
     Call TRACE(20, U&, FIN&)
     XX$ = REGLEIDO$("OCOMDETA", U&)
     '
     STAT% = Asc(Mid$(XX$, 124, 1))
     If STAT% <> 9 Then           ' no esta anulada
       NUORSE& = CVS(Mid$(XX$, 117, 4))
       If NUORSE& < 1 Then ' excluye orden de servicios
         MONEMISION% = Asc(Mid$(XX$, 122, 1))
         If MONEMISION% < 1 Then MONEMISION% = 1
         MONEVALUA% = Asc(Mid$(XX$, 121, 1))
         If MONEVALUA% < 1 Then MONEVALUA% = 1
         VALOUNI$ = Mid$(XX$, 43, 12)
         NPRO% = CVI(Mid$(XX$, 7, 2))
         '
         MONEVAC% = MONEVACO%(NPRO%)
         If MONEVAC% <= 1 Then
           If MONEMISION% <= 1 Then
             COEFUNI = CVS(Mid$(XX$, 63, 4))
             If COEFUNI < 0.0001 Then COEFUNI = 1
             IUNI# = CVD(Mid$(XX$, 79, 8)) / COEFUNI
             Z1$ = Space$(12)
             Call REPLA(Z1$, "$", 1, 1)
             Call REPLA(Z1$, CSTRING$(MKD$(IUNI#), 3, 10, 4, 2), 3, 10)
             '
             Call REPLA(XX$, Z1$, 43, 12)
             Call REPLA(XX$, Chr$(MONEVAC%), 121, 1)
             Call GRAREG("OCOMDETA", XX$, U&)
             '
             CIXI% = CVI(Mid$(XX$, 9, 2))
             If NPRO% = PROVHABI%(CIXI%) Then
               VUNI = IUNI#
               FECOI% = CVI(Mid$(XX$, 5, 2))
               If CIXI% > 0 Then
                 If CIXI% <= NUMAS% Then
                   REMA& = CIXI% + 1
                   ZZ$ = REGLEIDO$("MASTER", REMA&)
                   Call REPLA(ZZ$, MKS$(VUNI), 111, 4)
                   Call REPLA(ZZ$, MKI$(FECOI%), 115, 2)
                   Call REPLA(ZZ$, Chr$(MONEVAC%), 95, 1)
                   Call GRAREG("MASTER", ZZ$, REMA&)
                 End If
               End If
             End If
           End If
         End If
       End If
     End If
   Next U&
   Call CIERRARCH("OCOMDETA")
   Command3.Enabled = True
   '
End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   While WindowState <> 0
     WindowState = 0
     OPENFORMS = DoEvents
   Wend
End Sub

Private Sub Form_Resize()
   If WindowState = 1 Then
     Call CIERRARCH("*.*")
     Call BAJALDISCO
   End If
End Sub


Private Sub Form_Unload(Cancel As Integer)
   Call CIERRARCH("*.*")
   Call BAJALDISCO
   Call FINAL("")
End Sub
