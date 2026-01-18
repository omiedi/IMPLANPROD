VERSION 5.00
Object = "{20C62CAE-15DA-101B-B9A8-444553540000}#1.1#0"; "MSMAPI32.OCX"
Begin VB.Form AFLEXINI 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Administración Técnica FLEXOFT"
   ClientHeight    =   8070
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   5130
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8070
   ScaleWidth      =   5130
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command131 
      Caption         =   "Command13"
      Height          =   225
      Left            =   3360
      TabIndex        =   42
      Top             =   0
      Width           =   750
   End
   Begin VB.CommandButton Command26 
      Caption         =   "Command13"
      Height          =   225
      Left            =   4095
      TabIndex        =   41
      Top             =   6825
      Width           =   645
   End
   Begin VB.CommandButton Command12 
      Caption         =   "Recla Pago"
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
      Left            =   4095
      TabIndex        =   40
      Top             =   5145
      Width           =   855
   End
   Begin VB.CommandButton Command10 
      Caption         =   "CD DEMO"
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
      Left            =   4095
      TabIndex        =   36
      Top             =   4305
      Width           =   855
   End
   Begin VB.FileListBox File3 
      Height          =   480
      Left            =   4080
      TabIndex        =   32
      Top             =   3720
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.FileListBox File2 
      Height          =   3210
      Left            =   5280
      TabIndex        =   31
      Top             =   240
      Width           =   2415
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Print"
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
      Left            =   4095
      Picture         =   "AFLEXINI.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   27
      ToolTipText     =   "Papelería Comercial"
      Top             =   2835
      Width           =   855
   End
   Begin MSMAPI.MAPIMessages MAPIMessages1 
      Left            =   840
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      AddressEditFieldCount=   1
      AddressModifiable=   0   'False
      AddressResolveUI=   0   'False
      FetchSorted     =   0   'False
      FetchUnreadOnly =   0   'False
   End
   Begin MSMAPI.MAPISession MAPISession1 
      Left            =   105
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DownloadMail    =   -1  'True
      LogonUI         =   -1  'True
      NewSession      =   0   'False
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Remote"
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
      Left            =   4095
      Picture         =   "AFLEXINI.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   26
      ToolTipText     =   "Habilitación Remota"
      Top             =   5985
      Width           =   855
   End
   Begin VB.CommandButton Command2 
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
      Height          =   800
      Left            =   4095
      Picture         =   "AFLEXINI.frx":0614
      Style           =   1  'Graphical
      TabIndex        =   23
      ToolTipText     =   "Configuración de Funcionamiento"
      Top             =   1995
      Width           =   855
   End
   Begin VB.PictureBox MARCOBARRA 
      BackColor       =   &H000000FF&
      Height          =   135
      Left            =   4095
      ScaleHeight     =   75
      ScaleWidth      =   795
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   7095
      Width           =   855
      Begin VB.Frame BARRATRAZA 
         BackColor       =   &H00FF0000&
         BorderStyle     =   0  'None
         Height          =   500
         Left            =   0
         TabIndex        =   7
         Top             =   0
         Width           =   12000
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0E0E0&
      Caption         =   "Sub-Sistema Técnico"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7710
      Left            =   210
      TabIndex        =   4
      Top             =   210
      Width           =   3690
      Begin VB.Frame Frame3 
         BackColor       =   &H00C0E0E0&
         Caption         =   "Otros Movimientos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1800
         Left            =   210
         TabIndex        =   16
         Top             =   4350
         Width           =   3270
         Begin VB.PictureBox Picture2 
            BackColor       =   &H80000016&
            Height          =   645
            Left            =   630
            ScaleHeight     =   585
            ScaleWidth      =   2265
            TabIndex        =   17
            Top             =   945
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
               Picture         =   "AFLEXINI.frx":0A56
               Style           =   1  'Graphical
               TabIndex        =   18
               ToolTipText     =   "Transferencia de Fondos"
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label2 
               BackStyle       =   0  'Transparent
               Caption         =   "Sobres Flexoft Clientes"
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
               TabIndex        =   19
               Top             =   105
               Width           =   1380
            End
         End
         Begin VB.PictureBox Picture4 
            BackColor       =   &H80000016&
            Height          =   645
            Left            =   315
            ScaleHeight     =   585
            ScaleWidth      =   2265
            TabIndex        =   20
            Top             =   420
            Width           =   2325
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
               Left            =   0
               Picture         =   "AFLEXINI.frx":0D60
               Style           =   1  'Graphical
               TabIndex        =   21
               ToolTipText     =   "Depósito de Valores"
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label4 
               BackStyle       =   0  'Transparent
               Caption         =   "Juego Disquettes NOWEWIN.ARJ"
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
               TabIndex        =   22
               Top             =   45
               Width           =   1590
            End
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00C0E0E0&
         Caption         =   "Operar con Diskettes"
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
         TabIndex        =   12
         Top             =   6270
         Width           =   3270
         Begin VB.PictureBox Picture8 
            BackColor       =   &H80000016&
            Height          =   645
            Left            =   525
            ScaleHeight     =   585
            ScaleWidth      =   2160
            TabIndex        =   13
            Top             =   420
            Width           =   2220
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
               Picture         =   "AFLEXINI.frx":106A
               Style           =   1  'Graphical
               TabIndex        =   14
               ToolTipText     =   "Controlar, Formatear, limpiar Diskettes"
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label9 
               BackStyle       =   0  'Transparent
               Caption         =   "Control, Limpie- za y Formatear"
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
               TabIndex        =   15
               Top             =   105
               Width           =   1485
            End
         End
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00C0E0E0&
         Caption         =   "Sistema FLEXOFT"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3855
         Left            =   240
         TabIndex        =   5
         Top             =   420
         Width           =   3270
         Begin VB.PictureBox Picture5 
            BackColor       =   &H80000016&
            Height          =   645
            Left            =   480
            ScaleHeight     =   585
            ScaleWidth      =   2265
            TabIndex        =   37
            Top             =   2205
            Width           =   2325
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
               Left            =   0
               Picture         =   "AFLEXINI.frx":1374
               Style           =   1  'Graphical
               TabIndex        =   38
               ToolTipText     =   "Genera Manuales de Utilizacion"
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label5 
               BackStyle       =   0  'Transparent
               Caption         =   "Generar Manua- les de Uso"
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
               Left            =   720
               TabIndex        =   39
               Top             =   60
               Width           =   1515
            End
         End
         Begin VB.PictureBox Picture3 
            BackColor       =   &H80000016&
            Height          =   645
            Left            =   480
            ScaleHeight     =   585
            ScaleWidth      =   2265
            TabIndex        =   33
            Top             =   3045
            Width           =   2325
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
               Left            =   0
               Picture         =   "AFLEXINI.frx":167E
               Style           =   1  'Graphical
               TabIndex        =   34
               ToolTipText     =   "Genera Diskettes 'Llave' de Habilitación Estación de Trabajo"
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label3 
               BackStyle       =   0  'Transparent
               Caption         =   "Aparear Fuentes entre Equipos"
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
               Left            =   720
               TabIndex        =   35
               Top             =   60
               Width           =   1635
            End
         End
         Begin VB.PictureBox Picture14 
            BackColor       =   &H80000016&
            Height          =   645
            Left            =   720
            ScaleHeight     =   585
            ScaleWidth      =   2265
            TabIndex        =   8
            Top             =   1440
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
               Picture         =   "AFLEXINI.frx":17C8
               Style           =   1  'Graphical
               TabIndex        =   1
               ToolTipText     =   "Genera Diskettes 'Llave' de Habilitación Estación de Trabajo"
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label15 
               BackStyle       =   0  'Transparent
               Caption         =   "Crear Diskettes Habilita-Terminal"
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
               Left            =   720
               TabIndex        =   9
               Top             =   105
               Width           =   1515
            End
         End
         Begin VB.PictureBox Picture1 
            BackColor       =   &H80000016&
            Height          =   645
            Left            =   480
            ScaleHeight     =   585
            ScaleWidth      =   2265
            TabIndex        =   28
            Top             =   930
            Width           =   2325
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
               Left            =   0
               Picture         =   "AFLEXINI.frx":1AD2
               Style           =   1  'Graphical
               TabIndex        =   29
               ToolTipText     =   "Genera Diskettes 'Llave' de Habilitación Estación de Trabajo"
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label1 
               BackStyle       =   0  'Transparent
               Caption         =   "Crear C.D.'s de Instalación"
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
               Left            =   720
               TabIndex        =   30
               Top             =   60
               Width           =   1515
            End
         End
         Begin VB.PictureBox Picture16 
            BackColor       =   &H80000016&
            Height          =   645
            Left            =   240
            ScaleHeight     =   585
            ScaleWidth      =   2265
            TabIndex        =   10
            Top             =   420
            Width           =   2325
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
               Picture         =   "AFLEXINI.frx":1DDC
               Style           =   1  'Graphical
               TabIndex        =   0
               ToolTipText     =   "Actualizar Archivos de Extracción de Controladores"
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label17 
               BackStyle       =   0  'Transparent
               Caption         =   "Compilado de  Controladores"
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
               TabIndex        =   11
               Top             =   60
               Width           =   1590
            End
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
      Height          =   800
      Left            =   4095
      Picture         =   "AFLEXINI.frx":1F26
      Style           =   1  'Graphical
      TabIndex        =   2
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
      Left            =   4095
      Picture         =   "AFLEXINI.frx":2070
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Ayuda Flexoft en Línea"
      Top             =   1155
      Width           =   855
   End
   Begin VB.DirListBox Dir1 
      Height          =   765
      Left            =   2520
      TabIndex        =   24
      Top             =   3360
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.FileListBox File1 
      Height          =   480
      Left            =   1995
      TabIndex        =   25
      Top             =   3255
      Visible         =   0   'False
      Width           =   1065
   End
   Begin VB.Image Image2 
      Height          =   510
      Left            =   3570
      Picture         =   "AFLEXINI.frx":237A
      Top             =   7395
      Width           =   2010
   End
End
Attribute VB_Name = "AFLEXINI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
    Call CIERRARCH("*.*")
    Call BAJALDISCO
    Call FINAL("")
End Sub

Private Sub Command10_Click()
   Command10.Enabled = False
   GENERDEMO.Show 1
   Command10.Enabled = True
End Sub

Private Sub Command11_Click()
   Command11.Enabled = False
   Screen.MousePointer = 11
   '
   SELEMANUA.Show 1
   Call FRESHALL
   DoEvents
   '
   For U& = 1 To SELEMANUA.List1.ListCount
     If SELEMANUA.List1.Selected(U& - 1) = True Then GoTo 15
   Next U&
12 Screen.MousePointer = 1
   MsgBox "No se Han Seleccionado Aplicaciones"
   GoTo 99
   '
15 If SELEMANUA.List2.ListCount < 1 Then GoTo 12
   '
   RUTOR$ = "C:\Archivos de programa\Adobe\Acrobat 6.0\Reader\AcroRd32.EXE "
   TEXACT$ = "Adobe"
   If EXISTE%(TRIM$(RUTOR$)) < 1 Then
     RUTOR$ = "C:\Archivos de programa\Adobe\Acrobat 5.0\Reader\AcroRd32.EXE "
     TEXACT$ = "Acrobat"
   End If
   RUTAR$ = "F:\SOLO_LEER\MANUALES FLEXOFT\MAN-PDF\"
   '
   'desde aqui GENERA LOS MANUALES
   For U& = 1 To SELEMANUA.List2.ListCount
     If SELEMANUA.List2.Selected(U& - 1) = True Then
       If EXISTE%(RUTAR$ + SELEMANUA.List2.List(U& - 1) + ".PDF") < 1 Then
         MsgBox "No Existe Archivo " + RUTAR$ + SELEMANUA.List2.List(U& - 1) + ".PDF"
         GoTo 29
       End If
       '
       ANAME$ = RUTOR$ + RUTAR$ + SELEMANUA.List2.List(U& - 1) + ".PDF"
       RETURNVALUE = Shell(ANAME$, 4)
18     On Error Resume Next
       AppActivate TEXACT$ + " Reader - [" + SELEMANUA.List2.List(U& - 1) + ".PDF]" ' Adobe Reader"
       If Err Then
         Resume 20
20       On Error GoTo 0
         TINI = Timer
         While Timer < TINI + 1: DoEvents: Wend
         GoTo 18
       End If
       On Error GoTo 0
       DoEvents
       SendKeys "%AI" + Chr$(13), True
21     On Error Resume Next
       AppActivate TEXACT$ + " Reader - [" + SELEMANUA.List2.List(U& - 1) + ".PDF]" ' Adobe Reader"
       If Err Then
          On Error GoTo 0
          GoTo 25
       End If
22     SendKeys "%AS", True ' "%(F4)", False
       GoTo 21
       '
25     TINI = Timer
       While Timer < TINI + 1: DoEvents: Wend
     End If
29 Next U&
   '
   '
   Screen.MousePointer = 1
   Call COMUNI("Se han Emitido los Manuales Pedidos")
   '
99 Command11.Enabled = True
End Sub

Private Sub Command12_Click()
   Command12.Enabled = False
   OPX% = ALTERNA%("METALMEDICA\Telebit\ESCORIAL\ALEAR\TANIT\ALEJPA\COMPREFACIL\SILCAR\ASEMBLI\J y M\LUJON")
   Select Case OPX%
     Case 1    ' METALMEDICA
       AAX$ = LOADFILE$("F:\FLEXOFT\NOVEDAD\FLOGON.EXE")
       TTXX$ = ENCRIPTA$("NOTIFICACION DE ESTADO DE DEUDA PENDIENTE" + Space$(87))
       Mid$(AAX$, 1, 128) = TTXX$
       '
       ' ABONO MENSUAL
       BBX$ = Space$(128)
       Call REPLA(BBX$, "01", 1, 2)
       Call REPLA(BBX$, "SIST.INDUSTRIAL", 3, 20)
       Call REPLA(BBX$, "06/03/08", 23, 8)
       Call REPLA(BBX$, "$", 31, 4)
       Call REPLA(BBX$, "        0.00", 35, 12)
         SUCAR& = 0: For I% = 1 To 120
                       SUCAR& = SUCAR& + Asc(Mid$(BBX$, I%, 1)) - 71
                     Next I%
       Call REPLA(BBX$, TRIM$(Str$(SUCAR&)), 121, 8)
       CCX$ = ENCRIPTA$(BBX$)
       Mid$(AAX$, 129, 128) = CCX$
       '
       RUTAX$ = "X:\METALMEDICA\FLEXOFT\WINDOWS\"
       Call SAVEFILE(RUTAX$ + "MSTDRV04.EXE", AAX$)
       '
     Case 2    ' TELEBIT
       AAX$ = LOADFILE$("F:\FLEXOFT\NOVEDAD\FLOGON.EXE")
       TTXX$ = ENCRIPTA$("NOTIFICACION DE ESTADO DE DEUDA PENDIENTE" + Space$(87))
       Mid$(AAX$, 1, 128) = TTXX$
       '
       ' ABONO MENSUAL
       BBX$ = Space$(128)
       Call REPLA(BBX$, "02", 1, 2)
       Call REPLA(BBX$, "ABONO MENSUAL", 3, 20)
       Call REPLA(BBX$, "01/02/08", 23, 8)
       Call REPLA(BBX$, "$", 31, 4)
       Call REPLA(BBX$, "      231.40", 35, 12)
         SUCAR& = 0: For I% = 1 To 120
                       SUCAR& = SUCAR& + Asc(Mid$(BBX$, I%, 1)) - 71
                     Next I%
       Call REPLA(BBX$, TRIM$(Str$(SUCAR&)), 121, 8)
       CCX$ = ENCRIPTA$(BBX$)
       Mid$(AAX$, 129, 128) = CCX$
       '
       ' SISTEMA SUELDOS
'       BBX$ = Space$(128)
'       Call REPLA(BBX$, "01", 1, 2)
'       Call REPLA(BBX$, "SISTEMA SUELDOS", 3, 20)
'       Call REPLA(BBX$, "05/10/05", 23, 8)
'       Call REPLA(BBX$, "$", 31, 4)
'       Call REPLA(BBX$, "      600.00", 35, 12)
'         SUCAR& = 0: For I% = 1 To 120
'                       SUCAR& = SUCAR& + Asc(Mid$(BBX$, I%, 1)) - 71
'                     Next I%
'       Call REPLA(BBX$, TRIM$(Str$(SUCAR&)), 121, 8)
'       CCX$ = ENCRIPTA$(BBX$)
'       Mid$(AAX$, 257, 128) = CCX$
       '
       RUTAX$ = "x:\TELEBIT\FLEXOFT\WINDOWS\"
       Call SAVEFILE(RUTAX$ + "MSTDRV04.EXE", AAX$)
3
     Case 3   ' ESCORIAL
       AAX$ = LOADFILE$("F:\FLEXOFT\NOVEDAD\FLOGON.EXE")
       TTXX$ = ENCRIPTA$("NOTIFICACION DE ESTADO DE DEUDA PENDIENTE" + Space$(87))
       Mid$(AAX$, 1, 128) = TTXX$
       '
       ' SISTEMA COSTO INDUSTRIAL
'       BBX$ = Space$(128)
'       Call REPLA(BBX$, "01", 1, 2)
'       Call REPLA(BBX$, "S.COSTO INDUSTRIAL", 3, 20)
'       Call REPLA(BBX$, "20/10/07", 23, 8)
'       Call REPLA(BBX$, "$", 31, 4)
'       Call REPLA(BBX$, "     1350.00", 35, 12)
'         SUCAR& = 0: For I% = 1 To 120
'                       SUCAR& = SUCAR& + Asc(Mid$(BBX$, I%, 1)) - 71
'                     Next I%
'       Call REPLA(BBX$, TRIM$(Str$(SUCAR&)), 121, 8)
'       CCX$ = ENCRIPTA$(BBX$)
'       Mid$(AAX$, 129, 128) = CCX$
       '
       ' ABONO MENSUAL
       BBX$ = Space$(128)
       Call REPLA(BBX$, "02", 1, 2)
       Call REPLA(BBX$, "ABONO MENSUAL", 3, 20)
       Call REPLA(BBX$, "01/07/09", 23, 8)
       Call REPLA(BBX$, "$", 31, 4)
       Call REPLA(BBX$, "      729.69", 35, 12)
         SUCAR& = 0: For I% = 1 To 120
                       SUCAR& = SUCAR& + Asc(Mid$(BBX$, I%, 1)) - 71
                     Next I%
       Call REPLA(BBX$, TRIM$(Str$(SUCAR&)), 121, 8)
       CCX$ = ENCRIPTA$(BBX$)
       Mid$(AAX$, 257, 128) = CCX$
       '
       RUTAX$ = "X:\ESCORIAL\FLEXOFT\WINDOWS\"
       Call SAVEFILE(RUTAX$ + "MSTDRV04.EXE", AAX$)
       '
     Case 4   ' ALONSO
       AAX$ = LOADFILE$("F:\FLEXOFT\NOVEDAD\FLOGON.EXE")
       TTXX$ = ENCRIPTA$("NOTIFICACION DE ESTADO DE DEUDA PENDIENTE" + Space$(87))
       Mid$(AAX$, 1, 128) = TTXX$
       '
       ' ABONO MENSUAL
       BBX$ = Space$(128)
       Call REPLA(BBX$, "02", 1, 2)
       Call REPLA(BBX$, "ABONO MENSUAL", 3, 20)
       Call REPLA(BBX$, "10/09/09", 23, 8)
       Call REPLA(BBX$, "$", 31, 4)
       Call REPLA(BBX$, "      313.71", 35, 12)
         SUCAR& = 0: For I% = 1 To 120
                       SUCAR& = SUCAR& + Asc(Mid$(BBX$, I%, 1)) - 71
                     Next I%
       Call REPLA(BBX$, TRIM$(Str$(SUCAR&)), 121, 8)
       CCX$ = ENCRIPTA$(BBX$)
       Mid$(AAX$, 257, 128) = CCX$
       '
       RUTAX$ = "X:\ALEAR\FLEXOFT\WINDOWS\"
       Call SAVEFILE(RUTAX$ + "MSTDRV04.EXE", AAX$)
       '
     Case 5    ' TANIT
       AAX$ = LOADFILE$("F:\FLEXOFT\NOVEDAD\FLOGON.EXE")
       TTXX$ = ENCRIPTA$("NOTIFICACION DE ESTADO DE DEUDA PENDIENTE" + Space$(87))
       Mid$(AAX$, 1, 128) = TTXX$
       '
       ' ABONO MENSUAL
       BBX$ = Space$(128)
       Call REPLA(BBX$, "01", 1, 2)
       Call REPLA(BBX$, "ABONO OCT/08-ABR/09", 3, 20)
       Call REPLA(BBX$, "01/11/08", 23, 8)
       Call REPLA(BBX$, "$", 31, 4)
       Call REPLA(BBX$, "     3144.00", 35, 12)
         SUCAR& = 0: For I% = 1 To 120
                       SUCAR& = SUCAR& + Asc(Mid$(BBX$, I%, 1)) - 71
                     Next I%
       Call REPLA(BBX$, TRIM$(Str$(SUCAR&)), 121, 8)
       CCX$ = ENCRIPTA$(BBX$)
       Mid$(AAX$, 129, 128) = CCX$
'       BBX$ = Space$(128)
'       Call REPLA(BBX$, "01", 1, 2)
'       Call REPLA(BBX$, "IMPORTAR INVENTARIO", 3, 20)
'       Call REPLA(BBX$, "05/04/07", 23, 8)
'       Call REPLA(BBX$, "$", 31, 4)
'       Call REPLA(BBX$, "      550.00", 35, 12)
'         SUCAR& = 0: For I% = 1 To 120
'                       SUCAR& = SUCAR& + Asc(Mid$(BBX$, I%, 1)) - 71
'                     Next I%
'       Call REPLA(BBX$, TRIM$(Str$(SUCAR&)), 121, 8)
'       CCX$ = ENCRIPTA$(BBX$)
'       Mid$(AAX$, 257, 128) = CCX$
       RUTAX$ = "X:\TANIT\FLEXOFT\WINDOWS\"
       Call SAVEFILE(RUTAX$ + "MSTDRV04.EXE", AAX$)
       '
     Case 6   ' ALEJPA
       AAX$ = LOADFILE$("F:\FLEXOFT\NOVEDAD\FLOGON.EXE")
       TTXX$ = ENCRIPTA$("NOTIFICACION DE ESTADO DE DEUDA PENDIENTE" + Space$(87))
       Mid$(AAX$, 1, 128) = TTXX$
       '
       BBX$ = Space$(128)
       Call REPLA(BBX$, "02", 1, 2)
       Call REPLA(BBX$, "ABONO MENSUAL", 3, 20)
       Call REPLA(BBX$, "10/09/09", 23, 8)
       Call REPLA(BBX$, "$", 31, 4)
       Call REPLA(BBX$, "      480.00", 35, 12)
         SUCAR& = 0: For I% = 1 To 120
                       SUCAR& = SUCAR& + Asc(Mid$(BBX$, I%, 1)) - 71
                     Next I%
       Call REPLA(BBX$, TRIM$(Str$(SUCAR&)), 121, 8)
       CCX$ = ENCRIPTA$(BBX$)
       Mid$(AAX$, 129, 128) = CCX$
       '
       RUTAX$ = "X:\ALEJPA\FLEXOFT\WINDOWS\"
       Call SAVEFILE(RUTAX$ + "MSTDRV04.EXE", AAX$)
       '
     Case 7   ' COMPREFACIL
       AAX$ = LOADFILE$("F:\FLEXOFT\NOVEDAD\FLOGON.EXE")
       TTXX$ = ENCRIPTA$("NOTIFICACION DE ESTADO DE DEUDA PENDIENTE" + Space$(87))
       Mid$(AAX$, 1, 128) = TTXX$
       '
       ' SISTEMA SUCURSALES
       BBX$ = Space$(128)
       Call REPLA(BBX$, "01", 1, 2)
       Call REPLA(BBX$, "TICKET CREDITO", 3, 20)
       Call REPLA(BBX$, "10/04/08", 23, 8)
       Call REPLA(BBX$, "$", 31, 4)
       Call REPLA(BBX$, "      450.00", 35, 12)
         SUCAR& = 0: For I% = 1 To 120
                       SUCAR& = SUCAR& + Asc(Mid$(BBX$, I%, 1)) - 71
                     Next I%
       Call REPLA(BBX$, TRIM$(Str$(SUCAR&)), 121, 8)
       CCX$ = ENCRIPTA$(BBX$)
       Mid$(AAX$, 129, 128) = CCX$
       '
       RUTAX$ = "X:\COMPREFACIL\FLEXOFT\WINDOWS\"
       Call SAVEFILE(RUTAX$ + "MSTDRV04.EXE", AAX$)
       '
     Case 8   ' SILCAR
       AAX$ = LOADFILE$("F:\FLEXOFT\NOVEDAD\FLOGON.EXE")
       TTXX$ = ENCRIPTA$("NOTIFICACION DE ESTADO DE DEUDA PENDIENTE" + Space$(87))
       Mid$(AAX$, 1, 128) = TTXX$
       '
      ' ABONO MENSUAL
       BBX$ = Space$(128)
       Call REPLA(BBX$, "02", 1, 2)
       Call REPLA(BBX$, "ABONO MENSUAL", 3, 20)
       Call REPLA(BBX$, "01/06/08", 23, 8)
       Call REPLA(BBX$, "$", 31, 4)
       Call REPLA(BBX$, "      324.39", 35, 12)
         SUCAR& = 0: For I% = 1 To 120
                       SUCAR& = SUCAR& + Asc(Mid$(BBX$, I%, 1)) - 71
                     Next I%
       Call REPLA(BBX$, TRIM$(Str$(SUCAR&)), 121, 8)
       CCX$ = ENCRIPTA$(BBX$)
       Mid$(AAX$, 257, 128) = CCX$
       '
       RUTAX$ = "X:\SILCAR\FLEXOFT\WINDOWS\"
       Call SAVEFILE(RUTAX$ + "MSTDRV04.EXE", AAX$)
       '
     Case 9    ' ASEMBLI
       AAX$ = LOADFILE$("F:\FLEXOFT\NOVEDAD\FLOGON.EXE")
       TTXX$ = ENCRIPTA$("NOTIFICACION DE ESTADO DE DEUDA PENDIENTE" + Space$(87))
       Mid$(AAX$, 1, 128) = TTXX$
       '
       ' ABONO MENSUAL
       BBX$ = Space$(128)
       Call REPLA(BBX$, "02", 1, 2)
       Call REPLA(BBX$, "ABONO MENSUAL", 3, 20)
       Call REPLA(BBX$, "15/05/08", 23, 8)
       Call REPLA(BBX$, "$", 31, 4)
       Call REPLA(BBX$, "     2019.90", 35, 12)
         SUCAR& = 0: For I% = 1 To 120
                       SUCAR& = SUCAR& + Asc(Mid$(BBX$, I%, 1)) - 71
                     Next I%
       Call REPLA(BBX$, TRIM$(Str$(SUCAR&)), 121, 8)
       CCX$ = ENCRIPTA$(BBX$)
       Mid$(AAX$, 129, 128) = CCX$
       '
       RUTAX$ = "x:\GRAMY\FLEXOFT\WINDOWS\"
       Call SAVEFILE(RUTAX$ + "MSTDRV04.EXE", AAX$)
       '
     Case 10    ' J Y M
       AAX$ = LOADFILE$("F:\FLEXOFT\NOVEDAD\FLOGON.EXE")
       TTXX$ = ENCRIPTA$("NOTIFICACION DE ESTADO DE DEUDA PENDIENTE" + Space$(87))
       Mid$(AAX$, 1, 128) = TTXX$
       '
       ' ABONO MENSUAL
       BBX$ = Space$(128)
       Call REPLA(BBX$, "02", 1, 2)
       Call REPLA(BBX$, "ABONO MENSUAL", 3, 20)
       Call REPLA(BBX$, "01/11/07", 23, 8)
       Call REPLA(BBX$, "$", 31, 4)
       Call REPLA(BBX$, "      320.00", 35, 12)
         SUCAR& = 0: For I% = 1 To 120
                       SUCAR& = SUCAR& + Asc(Mid$(BBX$, I%, 1)) - 71
                     Next I%
       Call REPLA(BBX$, TRIM$(Str$(SUCAR&)), 121, 8)
       CCX$ = ENCRIPTA$(BBX$)
       Mid$(AAX$, 129, 128) = CCX$
       '
       RUTAX$ = "X:\JYM\FLEXOFT\WINDOWS\"
       Call SAVEFILE(RUTAX$ + "MSTDRV04.EXE", AAX$)
       '
      Case 11    ' METALMEDICA
       AAX$ = LOADFILE$("F:\FLEXOFT\NOVEDAD\FLOGON.EXE")
       TTXX$ = ENCRIPTA$("NOTIFICACION DE ESTADO DE DEUDA PENDIENTE" + Space$(87))
       Mid$(AAX$, 1, 128) = TTXX$
       '
       ' ABONO MENSUAL
       BBX$ = Space$(128)
       Call REPLA(BBX$, "01", 1, 2)
       Call REPLA(BBX$, "SIST.GESTION", 3, 20)
       Call REPLA(BBX$, "25/06/08", 23, 8)
       Call REPLA(BBX$, "$", 31, 4)
       Call REPLA(BBX$, "     1900.00", 35, 12)
         SUCAR& = 0: For I% = 1 To 120
                       SUCAR& = SUCAR& + Asc(Mid$(BBX$, I%, 1)) - 71
                     Next I%
       Call REPLA(BBX$, TRIM$(Str$(SUCAR&)), 121, 8)
       CCX$ = ENCRIPTA$(BBX$)
       Mid$(AAX$, 129, 128) = CCX$
       '
       RUTAX$ = "X:\LUJON\FLEXOFT\WINDOWS\"
       Call SAVEFILE(RUTAX$ + "MSTDRV04.EXE", AAX$)
       '
  End Select
   '
   Command12.Enabled = True
End Sub

Private Sub Command13_Click()
   RUTOR$ = "C:\Archivos de programa\Adobe\Acrobat 6.0\Reader\AcroRd32.EXE "
   TEXACT$ = "Adobe"
   If EXISTE%(TRIM$(RUTOR$)) < 1 Then
     RUTOR$ = "C:\Archivos de programa\Adobe\Acrobat 5.0\Reader\AcroRd32.EXE "
     TEXACT$ = "Acrobat"
   End If
   RUTAR$ = "C:\MIS DOCUMENTOS\DOCUMENTO1"
   '
   ANAME$ = RUTOR$ + RUTAR$ + ".PDF"
   RETURNVALUE = Shell(ANAME$, 4)
'Shell "c:\MIS DOCUMENTOS\DOCUMENTO1.PDF"
End Sub

Private Sub Command131_Click()
   AA1$ = LOADFILE$("C:\FLEXOFT\ORI_ZEBRA.TXT")
10 PPXX% = InStr(AA1$, Chr$(13) + Chr$(10))
   If PPXX% > 0 Then
     AA2$ = Left$(AA1$, PPXX% - 1) + Chr$(10) + Mid$(AA1$, PPXX% + 2)
     AA1$ = AA2$
     GoTo 10
   End If
   If Right$(AA1$, 1) <> Chr$(10) Then AA1$ = AA1$ + Chr$(10)
   Call SAVEFILE("C:\FLEXOFT\TRA_ZEBRA.TXT", AA1$)
End Sub

Private Sub Command14_Click()
   Command14.Enabled = False
   CTRDISQS.Show 1
   Command14.Enabled = True
End Sub


Private Sub Command15_Click()
    RUTAX$ = "C:\Archivos de programa\DevStudio\VB\FLEXOFT"
    APROY$ = "CONSTO04.VBP"
    N1% = FILEOPEN%(RUTAX$ + "\PROY_ACC\" + APROY$, 1, 128)
10  If Not EOF(N1%) Then
      Line Input #N1%, AA1$
      If Left$(AA1$, 7) = "Form=.." Then
        AA2$ = Mid$(AA1$, 8)
        N2% = FILEOPEN%(RUTAX$ + AA2$, 1, 128)
20      If Not EOF(N2%) Then
          Line Input #N2%, AA3$
          PPXX% = InStr(UCase$(AA3$), "CALL SELECHO(")
          If PPXX% > 0 Then
            AA4$ = Mid$(AA3$, PPXX% + 13)
            MsgBox AA4$
          End If
          GoTo 20
        End If
        Close #N2%
      End If
      GoTo 10
    End If
    Stop
    
    
End Sub

Private Sub Command2_Click()
   Command2.Enabled = False
   Call CONECRUN("FLSETUP", "Configuración de Funcionamiento")
   Command2.Enabled = True
End Sub

Private Sub Command22_Click()
   Command22.Enabled = False
   Call GEDINSTA
   Command22.Enabled = True
End Sub

Private Sub Command26_Click()
   ORI$ = LOADFILE$("F:\FLEXOFT\NOVEDAD\FLOGON.EXE")
   CONVE$ = CONVERTEX$(ORI$)
   RECU$ = RECUTEX$(CONVE$)
   AA1 = Len(ORI$)
   AA2 = Len(RECU$)
   If RECU$ = ORI$ Then
        MsgBox "IGUALES"
     Else
        MsgBox "DISTINTOS"
   End If
End Sub

Private Sub Command3_Click()
   Command3.Enabled = False
   Call SELECHO("DEUDOR1")
   NC% = Val(VALACT1$("DEUDOR1"))
   If NC% > 0 Then
     DESTIDOC% = NC%
     Call PRINTDOC("@SOBR-FLX")
   End If
   Command3.Enabled = True
End Sub

Private Sub Command4_Click()
   Command4.Enabled = False
   VERDAFAL = False
   If COMALTER("Desea Refrescar B.Datos Mails Recibidos ?\\Si,Refrescar\No, Continuar") = 1 Then
     VERDAFAL = True
   End If
   '
   Screen.MousePointer = 11
   MAPISession1.DownLoadMail = VERDAFAL
   MAPISession1.LogonUI = False
   MAPISession1.Action = 1
   MAPIMessages1.SessionID = MAPISession1.SessionID
   '
   DoEvents
   MAPIMessages1.Fetch
   DoEvents
   CAMENSA& = MAPIMessages1.MsgCount
   DoEvents
   '
   For U& = 1 To CAMENSA&
     Call TRACE(20, U&, CAMENSA&)
10   If U& > MAPIMessages1.MsgCount Then GoTo 99
     MAPIMessages1.MsgIndex = U& - 1
 
            'ADJUN$ = MAPIMessages1.AttachmentName
            'If ADJUN$ <> "" Then
            '   adju2$ = ADJUN$
            '   While InStr(adju2$, "\") > 0
            '     PPXYZ% = InStr(adju2$, "\")
            '     adju2$ = Mid$(adju2$, PPXYZ% + 1)
            '   Wend
            '   FileCopy ADJUN$, "C:\PRUEBA\" + adju2$
            'End If
            'AAA = BBB
            '
     ASUNTO$ = TRIM$(MAPIMessages1.MsgSubject)
     If Left$(ASUNTO$, 3) = "PWX" Then
       TEMAI$ = DESENUM$(MAPIMessages1.MsgNoteText)
       CLACO$ = SUBTEX$(TEMAI$, "CLACONTRA=", Chr$(124))
       If Len(CLACO$) >= 12 Then
         CLA1$ = Left$(CLACO$, 6)
         CLA2$ = Mid$(CLACO$, 7, 6)
         If CLA2$ = CONTRACLA$(CLA1$) Then
           REMAIFLEX.Label2(0) = SUBTEX$(TEMAI$, "EMPRELICEN=", Chr$(124))
           REMAIFLEX.Label2(1) = MAPIMessages1.MsgOrigDisplayName
           REMAIFLEX.Label2(2) = MAPIMessages1.MsgOrigAddress
           REMAIFLEX.Label2(3) = SUBTEX$(TEMAI$, "MODULO=", Chr$(124))
           REMAIFLEX.Label2(4) = SUBTEX$(TEMAI$, "VENCIMIENTO=", Chr$(124))
           REMAIFLEX.Label2(5) = SUBTEX$(TEMAI$, "ACCESADIC=", Chr$(124))
           REMAIFLEX.Label2(6) = SUBTEX$(TEMAI$, "SERIVOL=", Chr$(124))
           REMAIFLEX.Label2(7) = SUBTEX$(TEMAI$, "ENVIADOEL=", Chr$(124))
           Screen.MousePointer = 1
           REMAIFLEX.Show 1
           Screen.MousePointer = 11
           If ATENTI% = 1 Then
             MAPIMessages1.Delete
             GoTo 10
           End If
         End If
       End If
     End If
   Next U&
   '
99 On Error Resume Next
   MAPISession1.Action = 2
   On Error GoTo 0
   '
   Command4.Enabled = True
   Screen.MousePointer = 1
End Sub

Private Sub Command5_Click()
   Command5.Enabled = False
   Call CONECRUN("FLFORMS", "Papelería Comercial")
   Command5.Enabled = True
End Sub

Sub LIMPIACARPE(RUTA$)
   On Error Resume Next
   Kill RUTA$ + "\*.*"
   On Error GoTo 0
End Sub
Private Sub Command6_Click()
   Command6.Enabled = False
   Screen.MousePointer = 11
   On Error Resume Next
   MkDir "D:\ARMA_CD"
   MkDir "D:\ARMA_CD\FX_SYSBAS"
   MkDir "D:\ARMA_CD\FX_SETUP"
   MkDir "D:\ARMA_CD\FX_DEMOS"
   MkDir "D:\ARMA_CD\FX_DEMOS\LABELS"
   MkDir "D:\ARMA_CD\FX_DEMOS\SCREENS"
   MkDir "D:\ARMA_CD\FX_DEMOS\SCRIPTS"
   MkDir "D:\ARMA_CD\FX_DEMOS\Demos Unitarios"
   MkDir "D:\ARMA_CD\FX_DEMOS\Demos Unitarios\Ejecutables"
   MkDir "D:\ARMA_CD\FX_DEMOS\Demos Unitarios\Zipeados"
   MkDir "D:\ARMA_CD\FX_PROGRAMS"
   MkDir "D:\ARMA_CD\FX_MANUALS"
   On Error GoTo 0
   '
   On Error Resume Next
   Call LIMPIACARPE("D:\ARMA_CD")
   Call LIMPIACARPE("D:\ARMA_CD\FX_SYSBAS")
   Call LIMPIACARPE("D:\ARMA_CD\FX_SETUP")
   Call LIMPIACARPE("D:\ARMA_CD\FX_DEMOS")
   Call LIMPIACARPE("D:\ARMA_CD\FX_DEMOS\LABELS")
   Call LIMPIACARPE("D:\ARMA_CD\FX_DEMOS\SCREENS")
   Call LIMPIACARPE("D:\ARMA_CD\FX_DEMOS\SCRIPTS")
   Call LIMPIACARPE("D:\ARMA_CD\FX_DEMOS\Demos Unitarios")
   Call LIMPIACARPE("D:\ARMA_CD\FX_DEMOS\Demos Unitarios\Ejecutables")
   Call LIMPIACARPE("D:\ARMA_CD\FX_DEMOS\Demos Unitarios\Zipeados")
   Call LIMPIACARPE("D:\ARMA_CD\FX_PROGRAMS")
   Call LIMPIACARPE("D:\ARMA_CD\FX_MANUALS")
   On Error GoTo 0
   '
   SELECAPLI.Show 1
   Call FRESHALL
   DoEvents
   '
   For U& = 1 To SELECAPLI.List1.ListCount
     If SELECAPLI.List1.Selected(U& - 1) = True Then GoTo 15
   Next U&
12 Screen.MousePointer = 1
   MsgBox "No se Han Seleccionado Aplicaciones"
   GoTo 99
   '
15 If SELECAPLI.List2.ListCount < 1 Then GoTo 12
   '
   RUTOR$ = "F:\FLEXOFT\NOVEDAD\"
   RUDES$ = "D:\ARMA_CD\"
   ARXX$ = "Instalar_FLEXOFT.EXE"
   FileCopy RUTOR$ + ARXX$, RUDES$ + ARXX$
   '
   RUTOR$ = "C:\WINDOWS\ICONS\USUALES\"
   RUDES$ = "D:\ARMA_CD\"
   ARXX$ = "AUTORUN.ICO"
   FileCopy RUTOR$ + ARXX$, RUDES$ + ARXX$
   '
   ARXX$ = "AUTORUN.INF"
   TTXX$ = "[autorun]" + Chr$(13) + Chr$(10)
   TTXX$ = TTXX$ + "open = Instalar_FLEXOFT.EXE" + Chr$(13) + Chr$(10)
   TTXX$ = TTXX$ + "icon = AUTORUN.ICO" + Chr$(13) + Chr$(10)
   Call SAVEFILE(RUDES$ + ARXX$, TTXX$)
   '
   Call TRACD("PRN2FILE.COM")
   Call TRACD("ANSI.SYS")
   '
   Call TRACD("FLEXPRFS.FLX")
   Call TRACD("FLEXPRF.ZIP")
   Call TRACD("FLEXRFS.ZIP")
   '
   Call TRACD("OBJDETA.DRV")
   Call TRACD("OBJNAME.DRV")
   Call TRACD("LANGUAGE.DRV")
   Call TRACD("TRALABS.DRV")
   '
   Call TRACD("FLEXPIF.PIF")
   Call TRACD("DOSCOMM.PIF")
   '
   Call TRACD("ARJ.EXE")
   Call TRACD("BASRUN.EXE")
   Call TRACD("BRUN45.EXE")
   Call TRACD("DOSCOMM.EXE")
   Call TRACD("PKZIP.EXE")
   Call TRACD("PKUNZIP.EXE")
   '
   Call TRACD("FLEXQB1.EXE")
   Call TRACD("FLEXQB2.EXE")
   Call TRACD("FLSETUP.EXE")
   Call TRACD("FLOGON.EXE")
   Call TRACD("FLFORMS.EXE")
   Call TRACD("FLXMAIL.EXE")
   Call TRACD("PARAOBJ.EXE")
   Call TRACD("PARASTR.EXE")
   Call TRACD("PARASOR.EXE")
   Call TRACD("PARALIS.EXE")
   Call TRACD("PRIFORMS.EXE")
   Call TRACD("PARAMED.EXE")
   Call TRACD("PARAVEN.EXE")
   Call TRACD("PROTABU.EXE")
   Call TRACD("WINTABU.EXE")
   Call TRACD("WINTAB02.EXE")
   Call TRACD("WINTAB04.EXE")
   Call TRACD("FLEXCEL.EXE")
   Call TRACD("FLEXMNU.EXE")
   '
   File2.Path = "F:\FLEXOFT\FX_SETUP"
   ABXX$ = LOADFILE$("F:\FLEXOFT\FX_SETUP\SETUP.LST")
   PPXX% = InStr(ABXX$, "$(ProgramFiles)\WS_SETUP")
   'If PPXX% > 0 Then
   '  ABXX$ = Left$(ABXX$, PPXX% - 1) + "\FLEXOFT\SETUP" + Mid$(ABXX$, PPXX% + 24)
   '  Call SAVEFILE("F:\FX_SETUP\SETUP.LST", ABXX$)
   'End If
   '
   For U& = 1 To File2.ListCount
     ANAME$ = File2.List(U& - 1)
     FileCopy "F:\FLEXOFT\FX_SETUP\" + ANAME$, "D:\ARMA_CD\FX_SETUP\" + ANAME$
   Next U&
   '
   'desde aqui instala programas aplicativos
   For U& = 1 To SELECAPLI.List2.ListCount
     ANAME$ = SELECAPLI.List2.List(U& - 1)
     RUTOR$ = "F:\FLEXOFT\NOVEDAD\"
     If InStr(UCase$(ANAME$), "SEGUIMIENTO COMERCI") > 0 Then
       RUTOR$ = "F:\SEGUICOM\"
     End If
     If EXISTE%(RUTOR$ + ANAME$ + ".EXE") > 0 Then
       On Error Resume Next
       FileCopy RUTOR$ + ANAME$ + ".EXE", "D:\ARMA_CD\FX_PROGRAMS\" + ANAME$ + ".EXE"
       If Err Then
         MsgBox "Imposible Instalar " + ANAME$
       End If
       On Error GoTo 0
     End If
   Next U&
   
   'desde aqui instala demostraciones
   '
' TODO ESTO ESTA PENDIENTE DE REVISIÓN
'   Call TRASLA("Demostracion_FLEXOFT.EXE", "DEMOFLEX", "FX_DEMOS")
'   '
'   File3.Path = "\\Flexserver\DiscoD\DOCU_COM\DEMOSTRACIONES\DEMOFLEX\LABELS"
'   File3.Pattern = "*.*"
'   For U& = 1 To File3.ListCount
'     ARCHXX$ = File3.List(U& - 1)
'     Call TRASLA(ARCHXX$, Mid$(File3.Path, 4), "FX_DEMOS\LABELS")
'   Next U&
'   '
'   File3.Path = "F:\DEMOFLEX\SCRIPTS"
'   File3.Pattern = "*.*"
'   For U& = 1 To File3.ListCount
'     ARCHXX$ = File3.List(U& - 1)
'     Call TRASLA(ARCHXX$, Mid$(File3.Path, 4), "FX_DEMOS\SCRIPTS")
'   Next U&
'   '
'   File3.Path = "F:\DEMOFLEX\SCREENS"
'   File3.Pattern = "*.*"
'   For U& = 1 To File3.ListCount
'     ARCHXX$ = File3.List(U& - 1)
'     Call TRASLA(ARCHXX$, Mid$(File3.Path, 4), "FX_DEMOS\SCREENS")
'   Next U&
'   '
'   File3.Path = "F:\SOLO_LEER\ARCHIVOSCIME-SOI (ACS)\Documentacion comercial (DC)\Demostraciones (EXE)"
'   File3.Pattern = "*.*"
'   For U& = 1 To File3.ListCount
'     ARCHXX$ = File3.List(U& - 1)
'     Call TRASLA(ARCHXX$, Mid$(File3.Path, 4), "FX_DEMOS\Demos Unitarios\Ejecutables")
'   Next U&
'   '
'   File3.Path = "F:\SOLO_LEER\ARCHIVOSCIME-SOI (ACS)\Documentacion comercial (DC)\Demostraciones (ZIP)"
'   File3.Pattern = "*.*"
'   For U& = 1 To File3.ListCount
'     ARCHXX$ = File3.List(U& - 1)
'     Call TRASLA(ARCHXX$, Mid$(File3.Path, 4), "FX_DEMOS\Demos Unitarios\Zipeados")
'   Next U&
   '
   '
   Screen.MousePointer = 1
   Call COMUNI("Carpetas y Archivos Creados\en 'D:/ARMA_CD'")
   '
99 Command6.Enabled = True
End Sub

Sub TRACD(ARXX$)
   '
   If EXISTE%("F:\FLEXOFT\NOVEDAD\" + ARXX$) > 0 Then
       RUTOR$ = "F:\FLEXOFT\NOVEDAD\"
     Else
       RUTOR$ = "F:\FLEXOWIN\"
   End If
   RUDES$ = "D:\ARMA_CD\FX_SYSBAS\"
   FileCopy RUTOR$ + ARXX$, RUDES$ + ARXX$
   '
End Sub

Sub TRASLA(ARCHXX$, RUTORI$, RUDESTI$)
   On Error Resume Next
   FileCopy "\\Flexserver\DiscoD\DOCU_COM\DEMOSTRACIONES\" + RUTORI$ + "\" + ARCHXX$, "D:\ARMA_CD\" + RUDESTI$ + "\" + ARCHXX$
   On Error GoTo 0
End Sub

Private Sub Command7_Click()
   Command7.Enabled = False
   'ACTDRIVERS.Show 1
   GEFLXPRFS.Show 1
   Command7.Enabled = True
End Sub

Private Sub Command8_Click()
   Command8.Enabled = False
   GENOVWIN.Show 1
   Command8.Enabled = True
End Sub

Private Sub Command9_Click()
   Command9.Enabled = False
   'APARFUEN.Show 1
   APARFUEN2.Show 1
   Command9.Enabled = True
End Sub

Private Sub Form_Load()
   Call VERJOBID(OKEY%)
   If OKEY% < 1 Then Call FINAL("")
   '
   EPAC$ = TRIM$(EMPREAC$)
   If EPAC$ <> "" Then
     Caption = Caption + "  -  " + EPAC$
   End If
   '
   If APLINVO$ = "COMPAFU" Then
      Call Command9_Click
      Call FINAL("")
   End If
   '
End Sub

Private Sub Form_Resize()
   If WindowState = 1 Then
     Call CIERRARCH("*.*")
     Call BAJALDISCO
   End If
End Sub


Private Sub Form_Unload(Cancel As Integer)
   Call CIERRARCH("*.*")
   Call FINAL("")
End Sub

Sub GEDINSTA()    ' genera diskette de instalacion
    '
    If SERIVOL$("C:") <> "081D-0BD9" Then
      If CLACONTRA% <> 1 Then
        Call COMUNI("Operación no Permitida\en esta Estación de Trabajo.")
        Exit Sub
      End If
    End If
    '
10  VDEVU% = MsgBox("Coloque 'Diskette' de Instalación en Blanco", 1)
    If VDEVU% <> 1 Then Exit Sub
    '
    On Error GoTo 15
    Dir1.Path = "A:\"
    Dir1.Refresh
    File1.Path = "A:\"
    File1.Refresh
    On Error GoTo 0
    GoTo 20
    '
15  Resume 16
16  On Error GoTo 0
    MsgBox "Error de Acceso a la Disketera"
    GoTo 10
    '
20  If Dir1.ListCount > 0 Then
      MsgBox "Diskette no Válido - Tiene Información"
      GoTo 10
    End If
    '
    File1.Path = "A:\"
    File1.Refresh
    If File1.ListCount > 0 Then
      MsgBox "Diskette no Válido - Tiene Información"
      GoTo 10
    End If
    '
40  Dim BLEC As String * 127
    On Error GoTo 45
    N1% = FreeFile
    Open "A:FLEXKEY.KEY" For Random Access Read Write Shared As #N1% Len = 127
    On Error GoTo 0
    GoTo 50
    '
45  Resume 46
46  On Error GoTo 0
    MsgBox "Diskette Protegido - Quite Protección."
    GoTo 10
    '
50  Screen.MousePointer = 11
    SEVOLA$ = Left$(SERIVOL$("A:") + Space$(12), 12)
    If TRIM$(SEVOLA$) = "" Then
       Close #N1%
       Screen.MousePointer = 1
       MsgBox "Diskette No Válido."
       GoTo 10
    End If
    '
    For UL& = 1 To 499
       TTXX$ = RANDSTRING$(125)
       If UL& = 1 Then
           Mid$(TTXX$, 17, 12) = ENCRIPTA$(SEVOLA$)
         ElseIf UL& = 2 Then
           Mid$(TTXX$, 29, 12) = ENCRIPTA$(String$(12, 19))
         ElseIf UL& = 3 Then
           Mid$(TTXX$, 71, 12) = ENCRIPTA$(String$(12, 23))
       End If
       Get #N1%, UL&, BLEC$
       SUCAR% = 0
       For KI% = 1 To 125
          SUCAR% = SUCAR% + Asc(Mid$(TTXX$, KI%, 1)) - 83
       Next KI%
       BLEC$ = TTXX$ + MKI$(SUCAR%)
       Put #N1%, UL&, BLEC$
    Next UL&
    '
    Close #N1%
    Screen.MousePointer = 1
    MsgBox "Diskette Generado Exitosamente"
    GoTo 10
    '
99 End Sub
