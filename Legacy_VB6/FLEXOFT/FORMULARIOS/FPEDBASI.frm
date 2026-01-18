VERSION 5.00
Begin VB.Form FPEDBASI 
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Administracion de Pedidos"
   ClientHeight    =   4635
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   9330
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   FillStyle       =   0  'Solid
   Icon            =   "FPEDBASI.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   ScaleHeight     =   4635
   ScaleWidth      =   9330
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox MARCOBARRA 
      BackColor       =   &H000000FF&
      Height          =   135
      Left            =   4515
      ScaleHeight     =   75
      ScaleWidth      =   1005
      TabIndex        =   38
      Top             =   1050
      Width           =   1065
      Begin VB.Frame BARRATRAZA 
         BackColor       =   &H00FF0000&
         BorderStyle     =   0  'None
         Height          =   500
         Left            =   0
         TabIndex        =   39
         Top             =   0
         Width           =   12000
      End
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Control"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   5775
      Picture         =   "FPEDBASI.frx":000C
      Style           =   1  'Graphical
      TabIndex        =   32
      ToolTipText     =   "Configuración de Procesos y Formularios"
      Top             =   315
      Width           =   750
   End
   Begin VB.Frame Frame7 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Archivos Maestros"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1590
      Left            =   210
      TabIndex        =   13
      ToolTipText     =   "Base de Datos de Clientes en Cuenta Corriente"
      Top             =   210
      Width           =   4005
      Begin VB.PictureBox Picture9 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   420
         ScaleHeight     =   585
         ScaleWidth      =   3105
         TabIndex        =   33
         Top             =   630
         Width           =   3165
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
            Left            =   2550
            Picture         =   "FPEDBASI.frx":0316
            Style           =   1  'Graphical
            TabIndex        =   34
            ToolTipText     =   "Base de Datos de Clientes en Cuenta Corriente"
            Top             =   0
            Width           =   560
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
            Left            =   0
            Picture         =   "FPEDBASI.frx":0758
            Style           =   1  'Graphical
            TabIndex        =   35
            ToolTipText     =   "A-B-M-L Base de Datos de Items de Stock"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label10 
            BackStyle       =   0  'Transparent
            Caption         =   "Archivos Maestros de Artículos y Clientes"
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
            TabIndex        =   36
            Top             =   105
            Width           =   1905
         End
      End
      Begin VB.Frame Frame13 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Listados de Control"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2640
         Left            =   2625
         TabIndex        =   14
         Top             =   4830
         Width           =   4005
         Begin VB.CommandButton Command35 
            Caption         =   "Inventarios"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   210
            TabIndex        =   20
            Top             =   420
            Width           =   1590
         End
         Begin VB.CommandButton Command34 
            Caption         =   "Negativos"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   2205
            TabIndex        =   19
            Top             =   420
            Width           =   1590
         End
         Begin VB.CommandButton Command33 
            Caption         =   "Consignación"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   210
            TabIndex        =   18
            Top             =   1155
            Width           =   1590
         End
         Begin VB.CommandButton Command30 
            Caption         =   "Críticos"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   2205
            TabIndex        =   17
            Top             =   1155
            Width           =   1590
         End
         Begin VB.CommandButton Command28 
            Caption         =   "Estadísticas"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   210
            TabIndex        =   16
            Top             =   1890
            Width           =   1590
         End
         Begin VB.CommandButton Command27 
            Caption         =   "Análisis A-B-C"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   2205
            TabIndex        =   15
            Top             =   1890
            Width           =   1590
         End
      End
   End
   Begin VB.CommandButton Command8 
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
      Height          =   855
      Left            =   6675
      Picture         =   "FPEDBASI.frx":0A62
      Style           =   1  'Graphical
      TabIndex        =   12
      ToolTipText     =   "Acceso Directo a Otras Aplicaciones"
      Top             =   315
      Width           =   750
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
      Height          =   855
      Left            =   7530
      Picture         =   "FPEDBASI.frx":0D6C
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "Ayuda FLEXOFT en Línea"
      Top             =   315
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
      Height          =   855
      Left            =   8400
      Picture         =   "FPEDBASI.frx":1076
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "Terminar - Salir de la Aplicación"
      Top             =   315
      Width           =   750
   End
   Begin VB.Frame Frame9 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Base de Datos de Pedidos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1590
      Left            =   210
      TabIndex        =   2
      Top             =   1995
      Width           =   4005
      Begin VB.PictureBox Picture1 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   420
         ScaleHeight     =   585
         ScaleWidth      =   3105
         TabIndex        =   21
         Top             =   630
         Width           =   3165
         Begin VB.CommandButton Command41 
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
            Picture         =   "FPEDBASI.frx":11C0
            Style           =   1  'Graphical
            TabIndex        =   23
            ToolTipText     =   "Registrar Pedido de Cliente"
            Top             =   0
            Width           =   560
         End
         Begin VB.CommandButton Command45 
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
            Picture         =   "FPEDBASI.frx":1602
            Style           =   1  'Graphical
            TabIndex        =   22
            ToolTipText     =   "Anulación de Pedidos"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Carga y Anulación de Pedidos de Clientes"
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
            TabIndex        =   37
            Top             =   105
            Width           =   1905
         End
      End
      Begin VB.Frame Frame11 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Listados de Control"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2640
         Left            =   2625
         TabIndex        =   3
         Top             =   4830
         Width           =   4005
         Begin VB.CommandButton Command23 
            Caption         =   "Inventarios"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   210
            TabIndex        =   9
            Top             =   420
            Width           =   1590
         End
         Begin VB.CommandButton Command22 
            Caption         =   "Negativos"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   2205
            TabIndex        =   8
            Top             =   420
            Width           =   1590
         End
         Begin VB.CommandButton Command21 
            Caption         =   "Consignación"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   210
            TabIndex        =   7
            Top             =   1155
            Width           =   1590
         End
         Begin VB.CommandButton Command20 
            Caption         =   "Críticos"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   2205
            TabIndex        =   6
            Top             =   1155
            Width           =   1590
         End
         Begin VB.CommandButton Command17 
            Caption         =   "Estadísticas"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   210
            TabIndex        =   5
            Top             =   1890
            Width           =   1590
         End
         Begin VB.CommandButton Command16 
            Caption         =   "Análisis A-B-C"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   2205
            TabIndex        =   4
            Top             =   1890
            Width           =   1590
         End
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Consultas y Listados"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2115
      Left            =   4515
      TabIndex        =   1
      Top             =   1470
      Width           =   4635
      Begin VB.Frame Frame5 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Dispositivo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   610
         Left            =   420
         TabIndex        =   29
         Top             =   1365
         Width           =   3795
         Begin VB.OptionButton Option2 
            BackColor       =   &H00C0FFFF&
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
            Left            =   2205
            TabIndex        =   31
            Top             =   270
            Width           =   1170
         End
         Begin VB.OptionButton Option1 
            BackColor       =   &H00C0FFFF&
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
            Left            =   525
            TabIndex        =   30
            Top             =   270
            Value           =   -1  'True
            Width           =   1065
         End
      End
      Begin VB.PictureBox Picture6 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   420
         ScaleHeight     =   585
         ScaleWidth      =   3735
         TabIndex        =   24
         Top             =   525
         Width           =   3795
         Begin VB.CommandButton Command63 
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
            Picture         =   "FPEDBASI.frx":190C
            Style           =   1  'Graphical
            TabIndex        =   28
            ToolTipText     =   "Detalle por Número de Pedido"
            Top             =   0
            Width           =   600
         End
         Begin VB.CommandButton Command62 
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
            Left            =   2550
            Picture         =   "FPEDBASI.frx":1C16
            Style           =   1  'Graphical
            TabIndex        =   26
            ToolTipText     =   "Pedidos Pendientes por Cliente"
            Top             =   0
            Width           =   600
         End
         Begin VB.CommandButton Command64 
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
            Left            =   3150
            Picture         =   "FPEDBASI.frx":2058
            Style           =   1  'Graphical
            TabIndex        =   25
            ToolTipText     =   "Pedidos Pendientes por Código de Producto"
            Top             =   0
            Width           =   600
         End
         Begin VB.Label Label6 
            BackStyle       =   0  'Transparent
            Caption         =   "Base de Datos de Pedidos"
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
            Left            =   840
            TabIndex        =   27
            Top             =   105
            Width           =   1485
         End
      End
   End
   Begin VB.Timer Timer1 
      Left            =   0
      Top             =   0
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "    Sistema Modular Integrado de Procesamiento de Datos                              Linea FLEXOFT (r) - Version 9.03"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Left            =   3045
      TabIndex        =   0
      Top             =   3990
      Width           =   6315
   End
   Begin VB.Image Image3 
      Height          =   885
      Left            =   -105
      Picture         =   "FPEDBASI.frx":2362
      Stretch         =   -1  'True
      Top             =   3780
      Width           =   10920
   End
End
Attribute VB_Name = "FPEDBASI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Call CIERRARCH("*.*")
    Call FINAL("")
End Sub

Private Sub Command14_Click()
    OPORDESPA.Show 1
End Sub

Private Sub Command18_Click()
    Command18.Enabled = False
    Call CONECRUN("AMACLI", "Padrón Maestro de Clientes")
    Command18.Enabled = True
End Sub

Private Sub Command19_Click()
    Command19.Enabled = False
    Call CONECRUN("AMAPRO", "Padrón de Proveedores")
    Command19.Enabled = True
End Sub

Private Sub Command2_Click()
    Command2.Enabled = False
    Call CONECRUN("AMASTO", "Maestro de Items de Stock")
    Command2.Enabled = True
End Sub

Private Sub Command32_Click()
    Command32.Enabled = False
    Call CONECRUN("HORUTAS", "Secuencia de Operaciones Standard")
    Command32.Enabled = True
End Sub

Private Sub Command4_Click()
    If InStr(UCase$(EMPRELI$), "SABO") < 1 Then
         Call OPNOIN("Análisis de Pedidos de Kits")
         Exit Sub
       Else
         FANAKITS.Show 1
    End If
End Sub

Private Sub Command43_Click()
    Command43.Enabled = False
    Call DESANUPED
    Command43.Enabled = True
End Sub

Private Sub Command25_Click()
    Command25.Enabled = False
    Call FINAL("*CPROGEN")
    Command25.Enabled = True
End Sub

Private Sub Command29_Click()
    Call HELPONLINE("ORDEFABR")
End Sub

Private Sub Command3_Click()
    Command3.Enabled = False
    Call CONECRUN("FORMULAS", "Fórmulas y Estructuras de Producto")
    Command3.Enabled = True
End Sub

Private Sub Command42_Click()
    Command42.Enabled = False
    ' Call PEDINTNUE
    Call OPNOIN("Registrar Pedido Interno")
    Command42.Enabled = True
End Sub

Private Sub Command41_Click()
    Command41.Enabled = False
    Call PEDCLINUE
    Command41.Enabled = True
End Sub


Private Sub Command5_Click()
   Command5.Enabled = False
   Call COPEDPE
   Command5.Enabled = True
End Sub

Private Sub Command64_Click()
    Command64.Enabled = False
    If Option1.Value = True Then
           SALI$ = "*COPEDMA"
           If FORMAPED% = 2 Then SALI$ = "*COQEDMA"
           Call FINAL(SALI$)
        Else
           Screen.MousePointer = 11
           JJ& = 0
           FIN& = ULTREG&("PEDDETP")
           MODO% = COMALTER%("Clasificación por Fecha de Entrega Solicitada:\\Mes en Curso\Posteriores")
           FLI1% = -1: FLI2% = FINMESACT%
           If MODO% = 2 Then
               FLI1% = FINMESACT% + 1: FLI2% = 32767
           End If
           For U& = 1 To FIN&
               Call TRACE(20, U&, FIN&)
               x$ = REGLEIDO$("PEDDETP", U&)
               FENTRESOL% = CVI(Mid$(x$, 25, 2))
               If FENTRESOL% >= FLI1% Then
                  If FENTRESOL% <= FLI2% Then
                     Y$ = String$(46, 0) + Left$(x$, 6) + Mid$(x$, 9, 24) + Mid$(x$, 75)
                     Call REPLA(Y$, Mid$(x$, 7, 2), 1, 2)
                     Call REPLA(Y$, Mid$(x$, 7, 2), 45, 2)
                     Call REPLA(Y$, Mid$(x$, 33, 42), 3, 42)
                     '
                     CAPEN = CVS(Mid$(x$, 79, 4)) - CVS(Mid$(x$, 95, 4))
                     If CAPEN < 0 Then CAPEN = 0
                     CANOASI = CAPEN - CVS(Mid$(x$, 75, 4))
                     If CANOASI < 0 Then CANOASI = 0
                     '
                     NC% = CVI(Mid$(x$, 7, 2))
                     PREUNI = CVS(Mid$(x$, 83, 4)) * (100 - PORDESCLI(NC%)) / 100
                     ITOTA# = CDbl(CAPEN * PREUNI)
                     INOAS# = CDbl(CANOASI * PREUNI)
                     
                     Call REPLA(Y$, MKS$(PREUNI), 85, 4)
                     Call REPLA(Y$, MKD$(ITOTA#), 89, 8)
                     Call REPLA(Y$, MKS$(CANOASI), 111, 4)
                     Call REPLA(Y$, MKD$(INOAS#), 119, 8)
                     '
                     JJ& = JJ& + 1
                     Call GRAREG("LIPEDPE", Y$, JJ&)
                  End If
               End If
           Next U&
           Call SETULTREG("LIPEDPE", JJ&)
           Call CIERRARCH("LIPEDPE")
           Screen.MousePointer = 1
           '
           SALI$ = "*SOPENMA"
           If FORMAPED% = 2 Then SALI$ = "*SOQENMA"
           Call FINAL(SALI$)
    End If
    Command64.Enabled = True
End Sub

Private Sub Command8_Click()
    Call ACCDIR("ADMIPED")
End Sub

Private Sub Command9_Click()
    Command9.Enabled = False
    ORDESPACH.Show 1
    Command9.Enabled = True
End Sub

Private Sub Command45_Click()
    Command45.Enabled = False
    Call ANULAPED
    Command45.Enabled = True
End Sub

Private Sub Command46_Click()
    Command46.Enabled = False
    Call CAMBIAPED
    Command46.Enabled = True
End Sub

Private Sub Command61_Click()
    Command61.Enabled = False
    If Option1.Value = True Then
           Call OPNOIN("")     '  Call FINAL("*COPEDNU")
        Else
           Call OPNOIN("")
    End If
    Command61.Enabled = True
End Sub

Private Sub Command62_Click()
    Command62.Enabled = False
    If Option1.Value = True Then
           SALI$ = "*COPEDCL"
           If FORMAPED% = 2 Then SALI$ = "*COQEDCL"
           Call FINAL(SALI$)
        Else
           Screen.MousePointer = 11
           JJ& = 0
           FIN& = ULTREG&("PEDDETP")
           '
           MODO% = COMALTER%("Clasificación por Fecha de Entrega Solicitada:\\Mes en Curso\Posteriores")
           FLI1% = -1: FLI2% = FINMESACT%
           If MODO% = 2 Then
               FLI1% = FINMESACT% + 1: FLI2% = 32767
           End If
           For U& = 1 To FIN&
                Call TRACE(20, U&, FIN&)
                x$ = REGLEIDO$("PEDDETP", U&)
                FENTRESOL% = CVI(Mid$(x$, 25, 2))
                If FENTRESOL% >= FLI1% Then
                   If FENTRESOL% <= FLI2% Then
                      Y$ = String$(46, 0) + Left$(x$, 6) + Mid$(x$, 9, 24) + Mid$(x$, 75)
                      Call REPLA(Y$, Mid$(x$, 7, 2), 1, 2)
                      Call REPLA(Y$, Mid$(x$, 7, 2), 45, 2)
                      Call REPLA(Y$, Mid$(x$, 33, 42), 3, 42)
                      '
                      CAPEN = CVS(Mid$(x$, 79, 4)) - CVS(Mid$(x$, 95, 4))
                      If CAPEN < 0 Then CAPEN = 0
                      CANOASI = CAPEN - CVS(Mid$(x$, 75, 4))
                      If CANOASI < 0 Then CANOASI = 0
                      '
                      NC% = CVI(Mid$(x$, 7, 2))
                      PREUNI = CVS(Mid$(x$, 83, 4)) * (100 - PORDESCLI(NC%)) / 100
                      ITOTA# = CDbl(CAPEN * PREUNI)
                      INOAS# = CDbl(CANOASI * PREUNI)
                      '
                      Call REPLA(Y$, MKS$(PREUNI), 85, 4)
                      Call REPLA(Y$, MKD$(ITOTA#), 89, 8)
                      Call REPLA(Y$, MKS$(CANOASI), 111, 4)
                      Call REPLA(Y$, MKD$(INOAS#), 119, 8)
                      '
                      JJ& = JJ& + 1
                      Call GRAREG("LIPEDPE", Y$, JJ&)
                   End If
                End If
           Next U&
           Call SETULTREG("LIPEDPE", JJ&)
           Call CIERRARCH("LIPEDPE")
           Screen.MousePointer = 1
           '
           SALI$ = "*SOPENCL"
           If FORMAPED% = 2 Then SALI$ = "*SOQENCL"
           Call FINAL(SALI$)
    End If
    Command62.Enabled = True
End Sub

Private Sub Command63_Click()
    Command63.Enabled = False
    If Option1.Value = True Then
           SALI$ = "*COPEDNO"
           If FORMAPED% = 2 Then SALI$ = "*COQEDNO"
           Call FINAL(SALI$)
        Else
           Call OPNOIN("")
    End If
    Command63.Enabled = True
End Sub

Private Sub Form_Load()
    '
    Call VERJOBID
    EPAC$ = TRIM$(EMPREAC$)
    If EPAC$ <> "" Then
        Caption = Caption + "  -  " + EPAC$
    End If
    '
    If EXISTE%("PEDNOVO.EXE") > 0 Then
        MODO% = ALTERNA%("Procesar Pedidos de Clientes\Control de la Base de Datos")
        If MODO% = 1 Then
           Call CONECRUN("#PEDNOVO", "Pedidos de Clientes - Novoespuma")
          ElseIf MODO% = 2 Then
            Call COPEDPE
        End If
        Call FINAL("")
    End If
    '
    COMMA$ = Command$
    PPX% = InStr(COMMA$, "#")
    If PPX% > 0 Then
        COMMA$ = Left$(COMMA$, PPX% - 1)
    End If
    '
    If TRIM$(COMMA$) = "REVIPEPE" Then
        REVIPEPE% = 1
        Show
        Refresh
        ORDESPACH.Show 1
    End If
    '
End Sub
