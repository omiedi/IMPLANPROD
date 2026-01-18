VERSION 5.00
Begin VB.Form MREPARA09 
   BackColor       =   &H00FFFFC0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Reparaciones - Actualización y Consultas"
   ClientHeight    =   4530
   ClientLeft      =   45
   ClientTop       =   585
   ClientWidth     =   7740
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "INWINFO012.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4530
   ScaleWidth      =   7740
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Reparaciones"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3345
      Left            =   3720
      TabIndex        =   0
      Top             =   1080
      Width           =   3945
      Begin VB.PictureBox Picture7 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   1400
         ScaleHeight     =   585
         ScaleWidth      =   2265
         TabIndex        =   25
         Top             =   2450
         Width           =   2325
         Begin VB.CommandButton Command23 
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
            Picture         =   "INWINFO012.frx":030A
            Style           =   1  'Graphical
            TabIndex        =   27
            ToolTipText     =   "Entrega de Reparación"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label6 
            BackStyle       =   0  'Transparent
            Caption         =   "Entrega de Reparación"
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
            TabIndex        =   26
            Top             =   60
            Width           =   1590
         End
      End
      Begin VB.PictureBox Picture2 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   1070
         ScaleHeight     =   585
         ScaleWidth      =   2265
         TabIndex        =   1
         Top             =   1950
         Width           =   2325
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
            Picture         =   "INWINFO012.frx":0614
            Style           =   1  'Graphical
            TabIndex        =   2
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Cierre de Orden de Reparación"
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
            TabIndex        =   3
            Top             =   60
            Width           =   1590
         End
      End
      Begin VB.PictureBox Picture9 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   800
         ScaleHeight     =   585
         ScaleWidth      =   2265
         TabIndex        =   32
         Top             =   1450
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
            Picture         =   "INWINFO012.frx":091E
            Style           =   1  'Graphical
            TabIndex        =   34
            ToolTipText     =   "Describir Falla y Reparación"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label8 
            BackStyle       =   0  'Transparent
            Caption         =   "Describir Falla y Reparación"
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
            TabIndex        =   33
            Top             =   60
            Width           =   1485
         End
      End
      Begin VB.PictureBox Picture8 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   500
         ScaleHeight     =   585
         ScaleWidth      =   2265
         TabIndex        =   29
         Top             =   950
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
            Picture         =   "INWINFO012.frx":0A68
            Style           =   1  'Graphical
            TabIndex        =   31
            ToolTipText     =   "Transferencia Materiales de Depósito "
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label7 
            BackStyle       =   0  'Transparent
            Caption         =   "Transferencia de Materiales"
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
            Width           =   1485
         End
      End
      Begin VB.PictureBox Picture4 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   240
         ScaleHeight     =   585
         ScaleWidth      =   2265
         TabIndex        =   4
         Top             =   450
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
            Picture         =   "INWINFO012.frx":0BB2
            Style           =   1  'Graphical
            TabIndex        =   5
            ToolTipText     =   "Generar una Orden de Reparación"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   "Generar Orden de Reparación"
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
            Top             =   70
            Width           =   1485
         End
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00400000&
      Height          =   1065
      Left            =   3700
      ScaleHeight     =   1005
      ScaleWidth      =   3975
      TabIndex        =   21
      Top             =   0
      Width           =   4035
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
         Height          =   650
         Left            =   120
         Picture         =   "INWINFO012.frx":0EBC
         Style           =   1  'Graphical
         TabIndex        =   38
         ToolTipText     =   "Control Final de Reparaciones"
         Top             =   180
         Width           =   650
      End
      Begin VB.CommandButton AGenRep 
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
         Left            =   1800
         Picture         =   "INWINFO012.frx":11C6
         Style           =   1  'Graphical
         TabIndex        =   37
         ToolTipText     =   "Acceso a Consultas Pre-Configuradas"
         Top             =   180
         Width           =   650
      End
      Begin VB.PictureBox MARCOBARRA 
         BackColor       =   &H000000FF&
         Height          =   135
         Left            =   940
         ScaleHeight     =   75
         ScaleWidth      =   585
         TabIndex        =   35
         Top             =   180
         Width           =   650
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H00FF0000&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   0
            TabIndex        =   36
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
         Height          =   650
         Left            =   3230
         Picture         =   "INWINFO012.frx":14D0
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   180
         Width           =   650
      End
      Begin VB.CommandButton Command6 
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
         Height          =   650
         Left            =   2500
         Picture         =   "INWINFO012.frx":161A
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   180
         Width           =   650
      End
      Begin VB.Image Image2 
         Height          =   390
         Left            =   560
         Picture         =   "INWINFO012.frx":1A5C
         Top             =   410
         Width           =   1515
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Consultas - Re-impresión"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1785
      Left            =   120
      TabIndex        =   15
      Top             =   2640
      Width           =   3480
      Begin VB.OptionButton Option6 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Estado de Reparaciones"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   380
         Left            =   525
         TabIndex        =   39
         Top             =   1050
         Width           =   2685
      End
      Begin VB.OptionButton Option5 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Notas de Entrega (Remito)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   380
         Left            =   480
         TabIndex        =   28
         Top             =   1360
         Width           =   2685
      End
      Begin VB.OptionButton Option4 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Orden de Reparación"
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
         Left            =   1950
         TabIndex        =   19
         Top             =   630
         Width           =   1380
      End
      Begin VB.OptionButton Option3 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Presupuesto al Cliente"
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
         Left            =   1950
         TabIndex        =   18
         Top             =   210
         Width           =   1380
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Presupuesto Interno "
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
         TabIndex        =   17
         Top             =   630
         Width           =   1380
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Planilla de Reparación"
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
         Top             =   210
         Value           =   -1  'True
         Width           =   1380
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Preparación"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1170
      Left            =   105
      TabIndex        =   11
      Top             =   105
      Width           =   3480
      Begin VB.PictureBox Picture5 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   240
         ScaleHeight     =   585
         ScaleWidth      =   2910
         TabIndex        =   12
         Top             =   360
         Width           =   2970
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
            Picture         =   "INWINFO012.frx":397E
            Style           =   1  'Graphical
            TabIndex        =   24
            ToolTipText     =   "Alta de Reparación"
            Top             =   0
            Width           =   560
         End
         Begin VB.CommandButton Command20 
            BackColor       =   &H8000000A&
            Height          =   600
            Left            =   2370
            Picture         =   "INWINFO012.frx":3C88
            Style           =   1  'Graphical
            TabIndex        =   20
            ToolTipText     =   "Búsqueda Cliente por CUIT, Domicilio, Nombre de Fantasía, etc"
            Top             =   0
            Width           =   560
         End
         Begin VB.CommandButton Command5 
            BackColor       =   &H8000000A&
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
            Left            =   1800
            Picture         =   "INWINFO012.frx":4FEE
            Style           =   1  'Graphical
            TabIndex        =   13
            ToolTipText     =   "Planilla de Presupuestaciones"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label5 
            BackStyle       =   0  'Transparent
            Caption         =   "Planilla de Presupuesto"
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
            Left            =   650
            TabIndex        =   14
            Top             =   120
            Width           =   1590
         End
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Presupuestos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1185
      Left            =   120
      TabIndex        =   7
      Top             =   1330
      Width           =   3480
      Begin VB.PictureBox Picture3 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   525
         ScaleHeight     =   585
         ScaleWidth      =   2370
         TabIndex        =   8
         Top             =   360
         Width           =   2430
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
            Picture         =   "INWINFO012.frx":52F8
            Style           =   1  'Graphical
            TabIndex        =   9
            ToolTipText     =   "Generar Presupuesto por una Reparación"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "Generar y Emitir un Presupuesto"
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
            TabIndex        =   10
            Top             =   105
            Width           =   1380
         End
      End
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   7680
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Menu Repara 
      Caption         =   "Reparaciones"
      Begin VB.Menu AltaRepa 
         Caption         =   "Alta de Reparación"
      End
      Begin VB.Menu Planipre 
         Caption         =   "Planilla de Presupuestación"
         Begin VB.Menu Pladir 
            Caption         =   "Directa"
         End
         Begin VB.Menu Plabuscli 
            Caption         =   "Búsqueda de Cliente"
         End
      End
      Begin VB.Menu plh3 
         Caption         =   "-"
      End
      Begin VB.Menu Genpresu 
         Caption         =   "Generar Presupuesto"
      End
      Begin VB.Menu Ordrep 
         Caption         =   "Orden de Reparación"
      End
      Begin VB.Menu plh1 
         Caption         =   "-"
      End
      Begin VB.Menu Cierrorep 
         Caption         =   "Cierre de O/Reparación"
      End
      Begin VB.Menu plh2 
         Caption         =   "-"
      End
      Begin VB.Menu Exit 
         Caption         =   "Terminar"
      End
   End
   Begin VB.Menu Consul 
      Caption         =   "Consultas y Re-Impresión"
      Begin VB.Menu Anota 
         Caption         =   "Anotador de Seguimiento"
      End
      Begin VB.Menu plh5 
         Caption         =   "-"
      End
      Begin VB.Menu Reimpre 
         Caption         =   "Re-Impresiones"
         Begin VB.Menu Plarepa 
            Caption         =   "Planilla de Reparación"
         End
         Begin VB.Menu Presuint 
            Caption         =   "Presupuesto Interno"
         End
         Begin VB.Menu Presucli 
            Caption         =   "Presupuesto al Cliente"
         End
         Begin VB.Menu Orderep 
            Caption         =   "Orden de Reparación"
         End
         Begin VB.Menu RtoRep 
            Caption         =   "Entregas de Reparación"
         End
      End
   End
   Begin VB.Menu Tablaux 
      Caption         =   "Tablas Auxiliares"
      Begin VB.Menu Master 
         Caption         =   "Maestro de Artículos"
      End
      Begin VB.Menu Clientes 
         Caption         =   "Padrón de Clientes"
      End
   End
   Begin VB.Menu Setup 
      Caption         =   "Configuración"
   End
   Begin VB.Menu Help 
      Caption         =   "Ayuda"
   End
End
Attribute VB_Name = "MREPARA09"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub AGenRep_Click()
    '
    AGenRep.Enabled = False
    Call ACONSUL_PRE(TRIM$(UCase$(Name)))
    AGenRep.Enabled = True
    '
End Sub

Private Sub AGenRep_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   '
   If Button = 2 Then
      Call CONECTA_CONSUL_PRE(TRIM$(UCase$(Name)))
   End If
   '
End Sub

Private Sub AltaRepa_Click()
   Call Command9_Click
End Sub

Private Sub Anota_Click()
   Call Command8_Click
End Sub

Private Sub Cierrorep_Click()
   Call Command2_Click
End Sub

Private Sub Clientes_Click()
   Call CIERRARCH("*.*")
   ACONEC$ = "ARCHIG07/AMACLI"
   Call CONECRUN(ACONEC$, "Padron Maestro de Clientes")

'   Call PLAPRERE07.Command11_Click
End Sub

Private Sub Command1_Click()
    Call CIERRARCH("*.*")
    Call FINAL("")
End Sub

Sub VALEMATPORREP()
   '
   Dim ORFASEL$(8192)
   Dim CANECE(32767)
   '
   Call ABREORFAB
   Call COPYSTRU("INDIPRES", "INDERTRA")
10 Call CARGALISEL("INDERTRA", "REPARA", "NUMREPA/F6.0;REFEREN/A56", "NUMREPA > 0 AND STATUS = 3", "DISTINCT")
   Call SELECHO("INDERTRA/Indice General de Reparaciónes Generadas")
   NROREP& = XVALO(TRIM$(VALACT1("INDERTRA")))
   If NROREP& < 1 Then Exit Sub
    '
   NUMEREPA$ = "OR-" & TRIM$(Str$(NROREP&))
    '
   For U& = 0 To 32767
      CANECE(U&) = 0
   Next U&
   '
   If CONTROL$("", "DEPRECEN") <> "" Then
     DEPOCENTRAL% = CONTROL$("", "DEPRECEN") ' TRAE SEGUN CONFIGURACION
   End If
   If CONTROL$("", "DEPREDES") <> "" Then
     DEPODESTINO% = CONTROL$("", "DEPREDES")
   End If
   '
   VDEF$ = String$(64, 0)
   Call REPLA(VDEF$, NUMEREPA$, 1, 12)
   Call REPLA(VDEF$, Chr$(DEPODESTINO%), 14, 1)
   Call REPLA(VDEF$, Chr$(DEPOCENTRAL%), 13, 1)
   '
12 OBX$ = OBJPLA$("SEREPDEP", VDEF$)
   If OBX$ = "" Then Exit Sub
   '
   DEPODEST% = Asc(Mid$(OBX$, 14, 1))
   DepOrig% = Asc(Mid$(OBX$, 13, 1))
   '
   'VALIDACION QUE TENGA DERECHOS PARA CAMBIAR DEPOSITOS
   If DERACCE%("CAMBIDEP/NOMESS", "CAMBIO DE DEPOSITOS") < 2 Then
     If DEPOCENTRAL% <> DepOrig% Or DEPODESTINO% <> DEPODEST% Then
        Call MENSERR(24, "Usuario no Habilitado para Modificar Depósitos\Módulo: CAMBIDEP")
        GoTo 12
     End If
   End If
   '
   'VALIDA DEPOSITOS
   If DEPODEST% < 1 Or ECOARCH$("TADEPOSI", Chr$(DEPODEST%)) = "" Then 'VALIDACION
     Call MENSERR(24, "Depósito Destino No Válido")
      GoTo 12
   End If
   
   If DepOrig% < 1 Or ECOARCH$("TADEPOSI", Chr$(DepOrig%)) = "" Then
      Call MENSERR(24, "Depósito Origen No Válido")
      GoTo 12
   End If
   
   Call ABRESTOCKS
   'RECORRE LA TABLA RESERVAS Y ARMA EL ARCHIVO RESERVAS
   SQLX$ = " SELECT * FROM RESERVA "
   SQLX$ = SQLX$ + " WHERE CantRes-CantCons > 0 "
   SQLX$ = SQLX$ + " AND IDSUBOR = '" & NUMEREPA$ & "'"
    '
   Dim RESERTMP As ADODB.Recordset
   Set RESERTMP = FLEXCONN.Execute(FILTSQL$(SQLX$))

   With RESERTMP
      On Error Resume Next
      .MoveFirst
      If Err Then
       On Error GoTo 0
       Call MENSERR(24, " No Existen Reservas efectuadas. ")
       GoTo 12
      End If
      On Error GoTo 0
      JJ& = 0
      CONTA% = 0
      FIN& = .RecordCount
      Call SETULTREG("!RESERVA", 0)
      Do While (.EOF = False)
        Y$ = REGBLAN("!RESERVA")
        '
        CII% = XVALO(!COD_INTE)
        If ESTRAZABLE(CII%) = 2 Then
           CONTA% = CONTA% + 1 ' valida si hay codigos trazables
           GoTo 13
        End If
            Call REPLA(Y$, MKI$(!COD_INTE), 1, 2)
            Call REPLA(Y$, !IDSUBOR, 3, 12)
            Call REPLA(Y$, MKI$(!COD_INTE), 15, 2)
            Call REPLA(Y$, MKI$(CVINT(!FechRes)), 17, 2)
            Call REPLA(Y$, MKD$(!CANTRES), 19, 8)
            Call REPLA(Y$, MKD$(!CantCons), 27, 8)
            Call REPLA(Y$, MKD$(!Saldo_Entre), 35, 8)
            Call REPLA(Y$, !IdenLote, 43, 16)
            Call REPLA(Y$, Chr$(!Deposito), 60, 1)
            Call REPLA(Y$, MKS$(!UsoUnit), 75, 4)
            '
            JJ& = JJ& + 1
            Call GRAREG("!RESERVA", Y$, JJ&)
        '
13    .MoveNext
      Loop
   End With
   Set RESERTMP = Nothing
   Call SETULTREG("!RESERVA", JJ&)
   Call CIERRARCH("!RESERVA")
   If CONTA% > 0 Then
      Call MENSERR(24, "Se Han Encontrado Códigos con Trazabilidad = 2\Los Cuales Seran Excluidos de la Lista de Componentes")
   End If
   '
'   Call ASIGNALOTE_RESERVA(NUMEREPA$, OK%)
'MsgBox OK%
   FIN& = ULTREG&("!RESERVA")
   For U& = 1 To FIN&
      Call TRACE(20, U&, FIN&)
      XX$ = REGLEIDO$("!RESERVA", U&)
      NUMEREPA$ = Mid$(XX$, 3, 12)
      CI% = CVI(Left$(XX$, 2))
      CANTI = CVD(Mid$(XX$, 19, 8)) '- CVD(Mid$(XX$, 27, 8))
      '
      If CANTI < 0 Then CANTI = 0
      '
      If TRIM$(UCase$(UNIMED$(CI%))) = "GK" Then
         CANTI = CANTI / 1000
      End If
      '
15    If CI% > 0 Then
         If CI% <= NUMAS% Then
            CANECE(CI%) = CANECE(CI%) + CANTI
         End If
      End If
      '
19 Next U&
   '
   Call SETULTREG("!VALEMAT", 0)
   FIN& = NUMAS%
   Screen.MousePointer = 11
   For U& = 1 To FIN&
      Call TRACE(20, U&, FIN&)
      CI% = U&
      If CANECE(CI%) > 0 Then
          II1$ = CODEXT$(CI%)
          CONDI$ = " DoSecLar = '" & TRIM$(NUMEREPA$) & "' AND Cod_Inte = " & CI% & " AND DepoMovi = " & DEPODEST% & ""
          CANTXX = XVALO(VALOBADA("MOVISTO", "SUM(CantIngre)", CONDI$, "NOMESS"))
          CANTSAL = XVALO(VALOBADA("MOVISTO", "SUM(CantSalid)", CONDI$, "NOMESS"))
          STODEST = CANTXX - CANTSAL 'STODEPOS(CI%, DEPODEST%)
          STOCENT = STODEPOS(CI%, DepOrig%)
          NECNETA = CANECE(CI%) - STODEST
          If NECNETA < 0 Then NECNETA = 0
          
          VALEMATX = NECNETA ' EL VALE SE PRESENTA POR LA NECESIDAD NETA
          If NECNETA > STOCENT Then ' SI NO HAY STOCK CONSUME HASTA EL STOCK
             VALEMATX = STOCENT
          End If
          '
          If VALEMATX < 0.05 Then VALEMATX = 0
           '
          ZZ$ = REGBLAN$("VALEMAT")
          Call REPLA(ZZ$, MKI$(CI%), 1, 2)
          Call REPLA(ZZ$, MKS$(CANECE(CI%)), 3, 4) ' NECESIDAD BRUTA
          Call REPLA(ZZ$, MKS$(STODEST), 7, 4)     ' STOCK DEPOSITO DESTINO
          Call REPLA(ZZ$, MKS$(NECNETA), 11, 4)    ' NECESIDAD NETA
          Call REPLA(ZZ$, MKS$(STOCENT), 15, 4)    ' STOCK DEP.CENTRAL
          Call REPLA(ZZ$, MKS$(VALEMATX), 19, 4)   ' VALE A EMITIR
          Call REPLA(ZZ$, MKS$(SECCI), 23, 4)
          Call REPLA(ZZ$, "......", 27, 6)
          Call REPLA(ZZ$, UBISTO$(CI%), 33, 16)
 '         Call REPLA(ZZ$, IDLO$, 49, 16)
          Call REGAPP("!VALEMAT", ZZ$)
          'End If
      End If
   Next U&
   Call CIERRARCH("!VALEMAT")
   '
   Screen.MousePointer = 1
   JJ& = ULTREG&("!VALEMAT")
25 TITUPA$ = "Vale de TRANSFERENCIA" + " - de " + TRIM$(ECOARCH("TADEPOSI", MKI$(DepOrig%))) + " a " + TRIM$(ECOARCH$("TADEPOSI", MKI$(DEPODEST%)))
   VTB% = VENTABU%("VALEMAT" + "/TITUPAN=" + TITUPA$)
   If VTB% < 0 Then GoTo 10
   '
   ' BLOQUE AGREGADO PARA CONTROL DE NEGATIVOS
   CONTA% = 0
   '
   For jj1& = 1 To ULTREG&("!VALEMAT") ' valida si hay codigos trazables
      XX$ = REGLEIDO$("!VALEMAT", jj1&)
      CI% = CVI(Mid$(YY$, 1, 2))
      If ESTRAZABLE(CI%) = 2 Then
        CONTA% = CONTA% + 1
      End If
   Next jj1&
   '
   If CONTA% > 0 Then
      Call MENSERR(24, "Se Han Encontrado Códigos con Trazabilidad = 2\Los Cuales Deben ser Excluidos de la Lista de Componentes")
      GoTo 25
   End If
   '
   If CONTROL("VANEDECE", "DEPOCEN") = "SI" Then ' SI CONFIGURACION DE CONTROL DE NEGATIVOS
    For YX& = 1 To ULTREG&("!VALEMAT")
       YY$ = REGLEIDO$("!VALEMAT", YX&)
       CI% = CVI(Mid$(YY$, 1, 2))
       CAN# = CVS(Mid$(YY$, 19, 4))
           If CAN# > 0.005 Then
             STOCXCENT# = STODEPOS(CI%, DepOrig%) 'CVS(Mid$(YY$, 15, 4))
             VALMAT# = CVS(Mid$(YY$, 19, 4))
             If STOCXCENT# < VALMAT# Or STOCXCENT# < 0 Then
               Call MENSERR(24, "Transferencia No Procesable\Código: " & CODEXT$(CI%) & "\Generará Un Stock Negativo en Depósito de Origen")
               GoTo 25
             End If
           End If
    Next YX&
   End If
   If PUEDEMOVI%(DepOrig%) < 0 Then GoTo 25 ' CANCELA SALIDA
   ' FIN CONTROL DE NEGATIVOS
   '
   If COMALTER%("Confirme la Transferencia a Realizar\\Cancelar\Confirmar Transferencia") <> 2 Then GoTo 10
   '
28 ' RECORRER LA TABLA Y CARGAR EL LIST DE FRM TRAZCAMPO
   TRAZCOMPO.List2.Clear
   For YX& = 1 To ULTREG&("!VALEMAT")
       YY$ = REGLEIDO$("!VALEMAT", YX&)
       CI% = CVI(Mid$(YY$, 1, 2))
       If ESTRAZABLE%(CI%) = 1 Then
         CAN# = CVS(Mid$(YY$, 19, 4))
         If CAN# > 0.005 Then
            Call REPLA(TX$, CODEXT$(CI%), 1, 16)
            Call REPLA(TX$, DESCRIT$(CI%), 17, 32)
            Call REPLA(TX$, FORMATNUM(CAN#, "F10.1"), 50, 11)
            TRAZCOMPO.List2.AddItem TX$
         End If
       End If
   Next YX&
   
   'Presentar formulario para asignar lotes
   If TRAZCOMPO.List2.ListCount > 0 Then
     TRAZCOMPO.Label7 = NUMEREPA$
     TRAZCOMPO.Show 1
     Call FRESHALL
     If TRAZCOMPO.List2.ListCount < 1 Then
          TRAZCOMPO.LIMPIAR
          GoTo 25
     End If
   End If
   '
   For JJI% = 1 To TRAZCOMPO.List2.ListCount
      XXI$ = TRAZCOMPO.List2.List(JJI% - 1)
      CODE$ = TRIM$(Left$(XXI$, 16))
      For JJJ% = JJI% + 1 To TRAZCOMPO.List2.ListCount
         XXJ$ = TRAZCOMPO.List2.List(JJJ% - 1)
         If TRIM$(Left$(XXJ$, 16)) = CODE$ Then
            Call MENSERR(24, "En Reparaciones no Válido Combinar Lotes")
            GoTo 28
         End If
      Next JJJ%
   Next JJI%
   '
   For U& = 1 To ULTREG&("!VALEMAT")
      XX$ = REGLEIDO$("!VALEMAT", U&)
      CAN = CVS(Mid$(XX$, 19, 4))
      LOTEX$ = ""
      If CAN >= 0.05 Then
         CI1% = CVI(Mid$(XX$, 1, 2))
         If ESTRAZABLE%(CI1%) > 0 Then
            CODE$ = TRIM$(CODEXT$(CI1%))
            For JJI% = 1 To TRAZCOMPO.List2.ListCount
              XYX$ = TRAZCOMPO.List2.List(JJI% - 1)
              If TRIM$(Left$(XYX$, 16)) = CODE$ Then
                 LOTEX$ = TRIM$(Mid$(XYX$, 62, 16))
                 GoTo 17
              End If
            Next JJI%
         End If
      End If
17    Call REPLA(XX$, LOTEX$, 49, 16)
      Call GRAREG("!VALEMAT", XX$, U&)
   Next U&
   '
   TRAZCOMPO.List2.Clear
   ' GRABAR TRANSFERENCIA
   '
   CMOV$ = TRIM$(CONTROL$("", "CODVALMA"))
   If CMOV$ = "" Then CMOV$ = "TR"
   '
   NRO& = ProNroComprobante(CMOV$)
   '
   Call BLOQARCH("MOVISTO")
   '
   Screen.MousePointer = 11
   NC% = 0
   VUNI# = 0
   MONE$ = ""
   'LOTE$ = ""
   FECHA% = HOY(HOS$)
   DOPRI$ = "TR." + TRIM$(Str$(NRO&))
   'DOSEC$ = "TR." + Mid$(HO$, 8, 1) + Mid$(HO$, 4, 2) + Left$(HO$, 2)
   DOSEC$ = TRIM$(NUMEREPA$) ' NUMERO DE ORDEN DE REPARACION
   While Len(DOSEC$) < 10
     DOSEC$ = DOSEC$ + " "
   Wend
                              '
   NORI% = 0
   For U& = 1 To ULTREG&("!VALEMAT")
      XX$ = REGLEIDO$("!VALEMAT", U&)
      If CVS(Mid$(XX$, 19, 4)) >= 0.05 Then
         '
         CI1% = CVI(Mid$(XX$, 1, 2))
         CAN = CVS(Mid$(XX$, 19, 4))
         REFE$ = "TR.a " + TRIM$(ECOARCH$("TADEPOSI", MKI$(DepOrig%)))
         '
         CANTI = (-1) * CAN
         NORFI% = NORI% + 1
         '
         LOTE$ = TRIM$(Mid$(XX$, 49, 16)) ' INICIALIZAMOS POR SI NO ENTRA NO QUEDE ANTERIOR
         '
         Call MOVISTOK(FECHA%, CI1%, LOTE$, CMOV$, DOPRI$, DOPRI$, DOSEC$, DOSEC$, REFE$, NORI%, VUNI#, 0, NC%, DepOrig%, CANTI)
         '
         REFE$ = "TR.de " + TRIM$(ECOARCH$("TADEPOSI", MKI$(DEPODEST%)))
         '
         CANTI = CAN
         Call MOVISTOK(FECHA%, CI1%, LOTE$, CMOV$, DOPRI$, DOPRI$, DOSEC$, DOSEC$, REFE$, NORI%, VUNI#, 0, NC%, DEPODEST%, CANTI)
         '
         Call ACTUREGISTRO("RESERVA", "IDENLOTE", "COD_INTE = " & CI1% & " AND IDSUBOR = '" & NUMEREPA$ & "'", LOTE$, RAFE&, "NOMESS")
      End If
   Next U&
   '
   Call BAJALDISCO
   '
   FORIPRE$ = CONTROL$("", "FORVALMA")
   If FORIPRE$ <> "" Then
      '
      FECHDOC$ = HOS$
      NUMEDOC$ = TRIM$(Str$(NRO&))
      While Len(NUMEDOC$) < 6
         NUMEDOC$ = "0" + NUMEDOC$
      Wend
      TIPODOC$ = "Vale de Materiales"
      '
      CODPARAM$(1) = "DEPOSITO": DOCPARAM$(1) = TRIM$(Str$(DepOrig%)) + " - " + TRIM$(ECOARCH$("TADEPOSI", MKI$(DepOrig%)))
      CODPARAM$(2) = "DESTINO": DOCPARAM$(2) = TRIM$(Str$(DEPODEST%)) + " - " + TRIM$(ECOARCH$("TADEPOSI", MKI$(DEPODEST%)))
      CODPARAM$(3) = "DOCUORIG": DOCPARAM$(3) = TRIM$(NUMEREPA$)
      CAPARDOC% = 3
      '
      CODTABU1$(1) = "COD-MAT"
      CODTABU1$(2) = "DES-MAT"
      CODTABU1$(3) = "UNIMED"
      CODTABU1$(4) = "CANTIDAD"
      CODTABU1$(5) = "CAKILOS"
      CODTABU1$(6) = "REF-MAT1"
      CODTABU1$(7) = "NRO-LOTE"
      CACOLTAB1% = 7
      '
      DESTIDOC% = 0
      CAITAB1% = 0
      CAITCRU% = 0
      '
      For U& = 1 To ULTREG&("!VALEMAT")
         XX$ = REGLEIDO$("!VALEMAT", U&)
         If CVS(Mid$(XX$, 19, 4)) >= 0.05 Then
            '
            CICRU% = CVI(Mid$(XX$, 1, 2))
            CACOL1$ = Str$(CVS(Mid$(XX$, 19, 4)))
            KICOL1$ = Str$(CVS(Mid$(XX$, 19, 4)) * PESUNI(CICRU%))
            LOTE$ = TRIM$(Mid$(XX$, 49, 16))
            '
            CAITAB1% = CAITAB1% + 1
            TETABU1$(1, CAITAB1%) = CODEXT$(CICRU%)
            TETABU1$(2, CAITAB1%) = DESCRIT0$(CICRU%)
            TETABU1$(3, CAITAB1%) = UNIMED$(CICRU%)
            TETABU1$(4, CAITAB1%) = CACOL1$
            TETABU1$(5, CAITAB1%) = KICOL1$
            TETABU1$(6, CAITAB1%) = UBISTO$(CICRU%)
            TETABU1$(7, CAITAB1%) = LOTE$
            '
         End If
      Next U&
      '
      If CAITAB1% < 1 Then
         Screen.MousePointer = 1
         GoTo 10
      End If
      '
      Call PRINTDOC("FORVALMA")   ' Vale de Materiales
      Call BAJALDISCO
      '
   End If
   '
   Screen.MousePointer = 1
   GoTo 10
   '
End Sub

Private Sub Command11_Click()
    Command11.Enabled = False
    OPX% = ALTERNA%("Vale de Transferencia\Devolución por O/R")
    If OPX% < 1 Or OPX% > 2 Then GoTo 99
    If OPX% = 1 Then
      If DERACCE%("VALREPA", "Vale de Transferencia de Repuestos") = 2 Then Call VALEMATPORREP: GoTo 99
    End If
    If OPX% = 2 Then
      If DERACCE%("DEVOREP", "Devolucion de Repuestos") = 2 Then Call DEVOPORREP
    End If
99  Command11.Enabled = True
End Sub


Private Sub Command2_Click()
    '
    Command2.Enabled = False
    '
    If DERACCE%("CIEREPA", "Cierre de Orden de Reparación") < 2 Then
      GoTo 99
    End If
    '
    Call ABREORFAB
    Call COPYSTRU("INDIPRES", "INDEREPR")
10  CONDI$ = "NUMREPA > 0 AND STATUS = 3"
    Call CARGALISEL("INDEREPR", "REPARA", "NUMREPA/F6.0;REFEREN/A56", CONDI$, "DISTINCT")
    Call SELECHO("INDEREPR/Indice General Planillas de Reparación")
    NROREP& = XVALO(TRIM$(VALACT1("INDEREPR")))
    If NROREP& < 1 Then GoTo 99
    '
    'VALIDAR QUE HAYA PASADO POR EL CONTROL FINAL
    '********************************************
    'TRAIGO FECHA DE RECEPCION Y NUMERO DE SERIE DE REPARA
    CONDI$ = "NUMREPA = " & NROREP&
    FECHARECEP% = CVINT(VALOBADA("REPARA", "fechrecep", CONDI$))
    FECHARECEPCION = FETEXCOR1$(FECHARECEP%)
    NUSERIE$ = VALOBADA("REPARA", "NUMEROSERIE", CONDI$)
    '
    'CONTROLO QUE HAYA UN CONTROL POSTERIOR A LA RECEPCION DE ESTE NUMERO DE SERIE
    CONDI$ = "FECH_ENSA >= '" & FECHARECEPCION & "' AND NUMEROSERIE = '" & NUSERIE$ & "'"
    If CantRegistros&("CONTROLFIN", CONDI$, "NOMESS") < 1 Then
       OPX% = COMALTER%("Reparación Número : " & NROREP& & " Falta Realizar  Control Final.\Solo Podra Continuar Si Posee Derecho Correspondiente\\Si, Continuar\No Cancelar")
'       Call COMUNI("Reparación Numero : " & NROREP& & " Falta Realizar Control Final")
       If OPX% = 1 Then 'SOLO CON DERECHOS PUEDE REALIZAR EL CIERRE SIN CONTROL FINAL
         If DERACCE%("CIERSINC", "CIERRE SIN CONTROL FINAL") > 1 Then
            CIERREP09.Label25 = NROREP&
            CIERREP09.Show 1
            GoTo 10
         End If
       End If
       GoTo 10
    Else
       CIERREP09.Label25 = NROREP&
       CIERREP09.Show 1
       GoTo 10
    End If
    '
99  Command2.Enabled = True
    '
End Sub

Private Sub Command20_Click()
     
    Command20.Enabled = False
    Call MENSERR(24, "No Disponible en la Presente Versión")
    'Call PLAPREPB
    Command20.Enabled = True
End Sub

Private Sub Command23_Click()
    Command23.Enabled = False
    '
    If DERACCE%("DESPARE", "Despacho de Reparación") < 2 Then
      GoTo 99
    End If
    '
    Call ABREORFAB
    '
    Call COPYSTRU("INDIPRES", "INDORREM")
15  Call CARGALISEL("INDORREM", "REPARA", "NUMREPA/F6.0;REFEREN/A56", "STATUS >= 4", "DISTINCT")
    Call SELECHO("INDORREM")
    NROREP& = XVALO(TRIM$(VALACT1("INDORREM")))
    If NROREP& < 1 Then GoTo 99
    '
    REMI$ = VALOBADA("REPARA", "NUMREMITO", "NUMREPA = " & NROREP&, "NOMESS")
    '
    CODXT$ = ""
    CI% = XVALO(VALOBADA("REPARA", "COD_INTE", "NUMREPA = " & NROREP&, "NOMESS"))
    If CI% > 0 Then
       CODXT$ = CODEXT$(CI%)
    End If
    '
    If TRIM$(REMI$) = "" Then
       PRONREM& = 1 + XVALO(VALOBADA("REPARA", "max(cast(substring(NUMREMITO, 9, 8)as int))", "NUMREPA > 0", "NOMESS"))
       REP$ = TRIM$(FORMATNUM$(PRONREM&, "F8.0"))
       While Len(REP$) < 8: REP$ = "0" + REP$: Wend
       REMI$ = "RP-0001-" & REP$
    End If
    '
    FREP = HOY(HO$)
    FILTX$ = TRIM$(Str$(NROREP&)) & ";" & TRIM$(REMI$) & ";" & TRIM$(Str$(FREP)) & ";" & TRIM$(CODXT$)
    Call STDFORM("REMX-MDT", FILTX$)
    If STDSCREEN.APRO <> "OK" Then
       GoTo 15
    End If
    '
    Call ACTUREGISTRO("REPARA", "STATUS", "NUMREPA = " & NROREP&, 5, REGAF&, "NOMESS")
    '
    'grabo fecha de reparacion y numero de reparacion en tanumse
    NUSERI$ = VALOBADA("REPARA", "NUMEROSERIE", "NUMREPA = " & NROREP&, "NOMESS")
    FE% = FREP
    FE1% = DIASPOST(FE%, 180)
    FEVEGREP = CVDAT(FE1%)
   
    Call ACTUREGISTRO("TANUMSE", "FevenGaRep", "NUMEROSERIE = '" & NUSERI$ & "'", FEVEGREP, REGAF&, "NOMESS")
    Call ACTUREGISTRO("TANUMSE", "UnumRep", "NUMEROSERIE = '" & NUSERI$ & "'", NROREP&, REGAF&, "NOMESS")
    Call ACTUREGISTRO("REPARA", "FechEnt", "NUMREPA = " & NROREP&, FETEXCOR1$(FE%), REGAF&, "NOMESS")
    '
    CUERPO$ = "Equipo Remitido" & vbCrLf
    CUERPO$ = CUERPO$ + "Número de Reparación: " & NROREP& & vbCrLf '
    CUERPO$ = CUERPO$ + "Número de Remito: " & REMI$ & vbCrLf
    CUERPO$ = CUERPO$ + "Cliente: " & RACLI$
    Call ANOTANUSE(NUMESER$, CUERPO$, , "NOMESS")

99  Command23.Enabled = True

End Sub

Private Sub Command3_Click()

    Command3.Enabled = False
    '
    If DERACCE%("PRESURE", "Presupuesto de Reparación") < 2 Then
      GoTo 99
    End If
    '
    Call ABREORFAB
    Call COPYSTRU("INDIPRES", "INDEREPR")
    Call CARGALISEL("INDEREPR", "REPARA", "NUMREPA/F6.0;REFEREN/A56", "NUMREPA > 0 AND STATUS < 4 ", "DISTINCT")
    If ULTREG&("INDEREPR") < 1 Then
       Call MENSERR(24, "No Hay Reparaciones Pendientes de Presupuestar.")
       GoTo 99
    End If
    '
10  Call SELECHO("INDEREPR/Indice General Planillas de Reparación")
    NROREP& = XVALO(TRIM$(VALACT1("INDEREPR")))
    If NROREP& < 1 Then GoTo 99
    PRESURE09.Label25 = NROREP&
    PRESURE09.Label2 = ""
    CONDI$ = "Status_Gar > 0 AND NUMREPA = " & NROREP&
    If CantRegistros&("REPARA", CONDI$) > 0 Then
       PRESURE09.Label2 = "EN GARANTÍA"
    End If
    '
    PRESURE09.Show 1
    GoTo 10
    '
99  Command3.Enabled = True
    '
End Sub

Private Sub Command4_Click()
    Command4.Enabled = False
    '
    If DERACCE%("GENREPA", "Generación de Orden de Reparación") < 2 Then
      GoTo 99
    End If
    '
    Call ABREORFAB

    Call COPYSTRU("INDIPRES", "INDORPAR")
    'CARGA INDICE CON LOS QUE ESTAN EN STATUS 2 O 3 (PRESUPUESTADOS O CON ORDEN DE REPARACION (PARA SU EDICION)) O EQUIPOS EN GARANTIA
15  Call CARGALISEL("INDORPAR", "REPARA", "NUMREPA/F6.0;REFEREN/A56", "(Status_Gar > 0 AND STATUS = 1) OR (STATUS = 2)", "DISTINCT")
    If ULTREG&("INDORPAR") < 1 Then
        Call MENSERR(24, "No Hay Ordenes de Reparacion por Emitir")
        GoTo 99
    End If
    '
    Call SELECHO("INDORPAR")
    NROREP& = XVALO(TRIM$(VALACT1("INDORPAR")))
    If NROREP& < 1 Then GoTo 99
    '
    ' CONTROL QUE NO HAYAN MATERIALES TRANSFERIDOS
    DepDest% = XVALO(CONTROL$("", "DEPREDES"))
    NUMEREPA$ = "OR-" & TRIM$(Str$(NROREP&))
    Set RESERMAT = FLEXCONN.Execute(FILTSQL$("SELECT * FROM RESERVA WHERE IDSUBOR='" & NUMEREPA$ & "'"))
    
    With RESERMAT
        Do While Not .EOF
            CI% = !COD_INTE
            CANTIRESER# = XVALO(!CANTRES)
            If CI% > 0 Then
              ' CANTIDAD INGRESADA EN MOVISTO
              CONDI$ = "Cod_Inte = " & CI% & " AND DOSECCOR = '" & TRIM$(NUMEREPA$) & "' And DepoMovi = " & DepDest%
              CANTING = XVALO(VALOBADA("MOVISTO", "SUM(CantIngre)", CONDI$, "NOMESS"))
              ' CANTIDAD SALIDA MOVISTO
              CONDI$ = "CodiMovi = 'TR' AND Cod_Inte = " & CI% & " AND DOPRICOR = '" & TRIM$(NUMEREPA$) & "' AND DepoMovi = " & DepDest%
              CANTISL = XVALO(VALOBADA("MOVISTO", "SUM(CantSalid)", CONDI$, "NOMESS"))
              CANTOTA# = XVALO(CANTING) - XVALO(CANTISL)
              '
              If CANTOTA# > 0.05 Then
                 Call COMUNI("Imposible Re-Generar Reparaciones\que Poseen Materiales Transferidos.\Realice una Devolucion y Reintentelo Nuevamente")
                 Command4.Enabled = True
                 Exit Sub
              End If
            End If
            .MoveNext
        Loop
    End With

20  FILTX$ = TRIM$(Str$(NROREP&)) ' & ";" & TRIM$(Str$(FREP))
    Call STDFORM("OREP-MDT", FILTX$)
    If STDSCREEN.APRO <> "OK" Then
       GoTo 15
    End If
    '
    FEORCOM% = XVALO(VALOFORM("Text-1"))
    ORCOMPRA$ = VALOFORM("Text-3")
    If TRIM$(ORCOMPRA$) = "" Or FEORCOM% < 1 Then
        Call COMUNI("No se han Ingresado Fecha y Nº de Orden de Compra")
        GoTo 20
    End If
    '
    'TRASLADO DE REPUREPA A RESERVA
    Call BORRAREGISTROS("RESERVA", "IDSUBOR = '" & NUMEREPA$ & "'", REGAF&, "NOMESS")

    SQLX$ = "SELECT * FROM REPUREPA WITH(NOLOCK)"
    SQLX$ = SQLX$ + " WHERE NUMREPA = " & NROREP&
    SQLX$ = SQLX$ + " ORDER BY NUORD_ITEM"
    '
    Dim REPUTMP As ADODB.Recordset
    Set REPUTMP = FLEXCONN.Execute(FILTSQL$(SQLX$))
    With REPUTMP
    Do While Not .EOF
         SQLX1$ = "SELECT * FROM RESERVA"
         SQLX1$ = SQLX1$ + " WHERE COD_INTE < 1 "
         Dim RESETMP As ADODB.Recordset
         Set RESETMP = New ADODB.Recordset
         RESETMP.Open FILTSQL$(SQLX1$), FLEXCONN, adOpenDynamic, adLockPessimistic, adCmdText
         With RESETMP
           .AddNew   ' YA CONVERTIDO
           On Error Resume Next
           !CodiEmpr = CodiEmp%
           On Error GoTo 0
           CI% = XVALO(REPUTMP!COD_INTE)
           CANTI = XVALO(REPUTMP!CANT_SOL)
           !COD_INTE = CI%
           !COD_EXTE = CODEXT$(CI%)
           !Descrip = Left$(DESCRIT$(CI%), 64) 'modificado por que devolvia un error de datos
           !IDSUBOR = NUMEREPA$
           !FechRes = Now
           !CANTRES = CANTI
           !CantCons = 0
           !Saldo_Entre = CANTI
           !IdenLote = ""
           !Deposito = 0
           !UsoUnit = CANTI
           !NuOrItem = XVALO(REPUTMP!NuOrd_Item)
           .Update
         End With
         RESETMP.Close
         Set RESETMP = Nothing

    .MoveNext
    Loop
    End With
    
    REPUTMP.Close
    Set REPUTMP = Nothing
    '
    Call ACTUREGISTRO("REPARA", "STATUS", "NUMREPA = " & NROREP&, 3, REGAF&, "NOMESS")
    '
    NUMESER$ = VALOFORM("PARAM-10")
    RACLI$ = VALOFORM("PARAM-4")
    CUERPO$ = "Generación de Orden de Reparación" & vbCrLf
    CUERPO$ = CUERPO$ + "Número de Reparación: " & NROREP& & vbCrLf
    CUERPO$ = CUERPO$ + "Cliente: " & RACLI$
    Call ANOTANUSE(NUMESER$, CUERPO$, , "NOMESS")

99  Command4.Enabled = True
End Sub

Private Sub Command5_Click()
    Command5.Enabled = False
    '
    If DERACCE%("PLAPRERE", "Planilla de Presupuestación de Reparaciones") < 2 Then
      GoTo 99
    End If
    '
    Call ABREORFAB
    Call COPYSTRU("INDIPRES", "INDEREPR")
    'SE CAMBIA STATUS < 4 PARA INCLUIR LAS O/REPARACIONES GENERADAS
    Call CARGALISEL("INDEREPR", "REPARA", "NUMREPA/F6.0;REFEREN/A56", "NUMREPA > 0 AND STATUS < 4", "DISTINCT")
10  Call SELECHO("INDEREPR/Indice General Planillas de Reparación")
    NROREP& = XVALO(TRIM$(VALACT1("INDEREPR")))
    If NROREP& < 1 Then GoTo 99
    '
    If HOY(HO$) = FECHANUM("19/09/09") Then
       FILTX$ = TRIM$(Str$(NROREP&))
       Call STDFORM("PREP-MDT", FILTX$)
       GoTo 10
    End If
    '
    PLAPRERE09.Label25 = NROREP&
    PLAPRERE09.Show 1
    GoTo 10
    '
99  Command5.Enabled = True
End Sub

Private Sub Command6_Click()
   Command6.Enabled = False
   '
   If DERACCE%("STPREPA", "Configuracion de Funcionamiento de Reparación") < 2 Then
     GoTo 99
   End If
   '
   CONFREPA09.Show 1 ' SETUP
   '
99 Command6.Enabled = True
End Sub

Private Sub Command7_Click()
    '
    Command7.Enabled = False
    If DERACCE%("CIEREPA", "Cierre de Orden de Reparación") < 2 Then
      GoTo 99
    End If
    '
    Call ABREORFAB
    Call COPYSTRU("INDIPRES", "INDEREPR")
10  Call CARGALISEL("INDEREPR", "REPARA", "NUMREPA/F6.0;REFEREN/A56", "NUMREPA > 0 AND STATUS = 3", "DISTINCT")
    Call SELECHO("INDEREPR/Indice General Planillas de Reparación")
    NROREP& = XVALO(TRIM$(VALACT1("INDEREPR")))
    If NROREP& < 1 Then GoTo 99
    '
    DESCRIREP.Label25 = NROREP&
    DESCRIREP.Show 1
    GoTo 10

99  Command7.Enabled = True
    '
End Sub

Sub MATEROREP()
    '
    Screen.MousePointer = 11
    Call COPYSTRU("PRESREPA", "OREPARPE")
    JJ& = 0
    For U& = 1 To ULTREG&("PRESREPA")
      XX$ = REGLEIDO$("PRESREPA", U&)
      If CVS(Mid$(XX$, 125, 4)) >= 0.5 Then
        JJ& = JJ& + 1
        Call GRAREG("OREPARPE", XX$, JJ&)
      End If
    Next U&
    Call SETULTREG("OREPARPE", JJ&)
    Call CIERRARCH("OREPARPE")
    Screen.MousePointer = 1
    '
    Call FINAL("*COREPUES")

99 End Sub
'
Sub OPEROREP()
   '
   Call COPYSTRU("INDIOTRA", "INDIOREP") ' INDICE DE TODAS LAS o/t'S
   '
   Screen.MousePointer = 11
   JJ& = 0
   REBLA$ = REGBLAN$("INDIOREP")
   FIN& = ULTREG&("PRESREPA")
   For U& = 1 To FIN&
     Call TRACE(20, U&, FIN&)
     XX$ = REGLEIDO$("PRESREPA", U&)
     If CVS(Mid$(XX$, 125, 4)) >= 0.5 Then
       NOREPARA& = CVS(Left$(XX$, 4))
       NORFA$ = "OR-" + TRIM$(Str$(NOREPARA&))
       DORFA$ = TRIM$(Mid$(XX$, 5, 60))
       FEMI% = CVI(Mid$(XX$, 121, 2))
       '
       ZZ$ = REBLA$
       Call REPLA(ZZ$, NORFA$, 1, 12)
       Call REPLA(ZZ$, DORFA$, 13, 48)
       Call REPLA(ZZ$, MKI$(FEMI%), 61, 2)
       JJ& = JJ& + 1
       Call GRAREG("INDIOREP", ZZ$, JJ&)
     End If
   Next U&
   Call SETULTREG("INDIOREP", JJ&)
   Call CIERRARCH("INDIOREP")
   Screen.MousePointer = 1
   '
   Call FINAL("*COMOBRAD")
   '
End Sub
'
Sub PRESEDOC()
    '
    Call ABREORFAB
3   If Option1.Value = True Then ' VER
          ASEL$ = "PRESREPA"
          TSEL$ = "Indice General Planillas de Reparación"
          TIDOCX$ = TRIM$("Planilla de Reparación")
        ElseIf Option2.Value = True Then ' MODIFICADA
          '
          If DERACCE%("VEPREPAR", "Visualizar Presupuestos de Reparación") < 2 Then
            Exit Sub
          End If
          '
          Screen.MousePointer = 11
          Call COPYSTRU("PRESREPA", "PREPAMI")
          Call CARGALISEL("PREPAMI", "REPARA", "NUMREPA/F6.0;REFEREN/A56", "NUMREPA > 0 AND STATUS > 1", "DISTINCT")
          Call CIERRARCH("PREPAMI")
          Screen.MousePointer = 1
          ASEL$ = "PREPAMI"
          TSEL$ = "Indice General Presupuestos de Reparación"
          TIDOCX$ = TRIM$("Presupuesto de Reparación")
        ElseIf Option3.Value = True Then ' MODIFICADA
          '
          If DERACCE%("VEPREPAR", "Visualizar Presupuestos de Reparación") < 2 Then
            Exit Sub
          End If
          '
          Screen.MousePointer = 11
          Call COPYSTRU("PRESREPA", "PREPAEMI")
          Call CARGALISEL("PREPAMI", "REPARA", "NUMREPA/F6.0;REFEREN/A56", "NUMREPA > 0 AND STATUS > 1", "DISTINCT")
          '
          Call CIERRARCH("PREPAMI")
          Screen.MousePointer = 1
          ASEL$ = "PREPAMI"
          TSEL$ = "Indice General Presupuestos de Reparación"
          'TIPODOC$ = "Presup.Reparación (Externo)" MODIFICADO POR QUE CON EL PARENTESIS NO ENVIABA EL MAIL
          TIDOCX$ = TRIM$("Presup.Reparación Externo")
        ElseIf Option4.Value = True Then
          Screen.MousePointer = 11
          Call COPYSTRU("PRESREPA", "OREPARPE")
          JJ& = 0
          For U& = 1 To ULTREG&("PRESREPA")
            XX$ = REGLEIDO$("PRESREPA", U&)
            If CVS(Mid$(XX$, 125, 4)) >= 0.5 Then
              JJ& = JJ& + 1
              Call GRAREG("OREPARPE", XX$, JJ&)
            End If
          Next U&
          Call SETULTREG("OREPARPE", JJ&)
          Call CIERRARCH("OREPARPE")
          If ULTREG&("OREPARPE") < 1 Then
             Call MENSERR(24, "No Disponible en la Presente Versión")
             Screen.MousePointer = 1
             GoTo 99
          End If
          Screen.MousePointer = 1
          ASEL$ = "OREPARPE"
          TSEL$ = "Indice General Ordenes de Reparación"
          TIDOCX$ = TRIM$("Orden de Reparación")
    End If
    '
4   Call FRESHECHO(ASEL$)
    If ULTREG&(ASEL$) < 1 Then
       Call MENSERR(24, "No Disponible en la Presente Versión")
       GoTo 99
    End If
    '
    Call SELECHO(ASEL$ + "/" + TSEL$)
    NPX& = Val(VALACT1$(ASEL$))
    If NPX& < 1 Then GoTo 99
    '
    Screen.MousePointer = 11
    NUDOCU$ = TRIM$(Str$(NPX&))
    TIDOCU$ = TIDOCX$ + " " + NUDOCU$ + " "
    '
    For UI& = ULTREG&("PDOCLST") To 1 Step -1
      XX$ = Left$(REGLEIDO$("PDOCLST", UI&), 64)
      If InStr(XX$, TIDOCU$) > 4 Then
        DOCUNU& = CVS(Left$(XX$, 4))
        GoTo 5
      End If
    Next UI&
    '
    Screen.MousePointer = 1
    Call COMUNI("No se Encontró el Documento Impreso\Correspondiente a esta Reparación.")
    On Error Resume Next
    GoTo 99
    '
5   Call MUESTRADOC(DOCUNU&) 'PERMITE VER LOS DOCUMENTOS ANTIGUOS POR PANTALLA
    GoTo 3
    '
99  End Sub

Sub PLAPREPA()
   '
10 Call SELECHO("DEUDOR1")
   NC% = Val(VALACT1$("DEUDOR1"))
   If NC% > 0 Then
     If RASOCLI$(NC%) <> "" Then
       PLAPRERE07.Label11 = TRIM$(Str$(NC%))
       PLAPRERE07.Label18 = TRIM$(Str$(NUPRESUP&))
       PLAPRERE07.Show 1
       GoTo 10
     End If
   End If
   '
End Sub

Sub PLAPREPB()
   Static ARGU$
   Command20.Enabled = False
10 ARGU$ = UCase$(TRIM$(InputBox$("Buscar", "Búsqueda Cliente por Texto", ARGU$)))
   If ARGU$ <> "" Then
     Screen.MousePointer = 11
     JJ& = 0
     Call COPYSTRU("DEUDOR1", "SELECLI")
     '
     Call APERBASDAT("CLIEN")
     SQLX$ = "SELECT * FROM DEUDOR12"
     Dim DEUDDDOR12 As ADODB.Recordset
     Set DEUDDDOR12 = New ADODB.Recordset
     DEUDDDOR12.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockReadOnly, adCmdText
     On Error Resume Next
     With DEUDDDOR12
       .MoveFirst
       If Err < 1 Then
         On Error GoTo 0
         Do While Not .EOF
           For KKI% = 1 To .Fields.Count
             On Error Resume Next
             VATEX$ = ""
             VATEX$ = .Fields(KKI% - 1).Value
             On Error GoTo 0
             If InStr(VATEX$, ARGU$) > 0 Then
               NCLIE% = !Nume_Cli
               RACLI$ = !Raso_Cli
               XX$ = MKI$(NCLIE%) + RACLI$
               JJ& = JJ& + 1
               Call GRAREG("!SELECLI", XX$, JJ&)
               Exit For
             End If
           Next KKI%
         .MoveNext
         Loop
       End If
     End With
     Set DEUDDDOR12 = Nothing
     Call SETULTREG("!SELECLI", JJ&)
     '
     Screen.MousePointer = 1
     If JJ& < 1 Then
         Call MENSERR(24, "No se Encontraron Coincidencias")
         GoTo 10
       ElseIf JJ& = 1 Then
         PLAPRERE07.Label11 = TRIM$(Str$(CVI(Left$(REGLEIDO$("!SELECLI", 1), 2))))
         PLAPRERE07.Label18 = TRIM$(Str$(NUPRESUP&))
         PLAPRERE07.Show 1
         GoTo 10
       Else
         Call FRESHECHO("!SELECLI")
         Call SELECHO("!SELECLI")
         NCLIX$ = VALACT1$("!SELECLI")
         If NCLIX$ <> "" Then
           PLAPRERE07.Label11 = TRIM$(NCLIX$)
           PLAPRERE07.Label18 = TRIM$(Str$(NUPRESUP&))
           PLAPRERE07.Show 1
           GoTo 10
         End If
     End If
   End If
   '
End Sub

Sub PRESUREPA()
   '
   
'   If DERACCE%("PRESREPA", "Generar Presupuestos Reparaciones") < 2 Then
'     Exit Sub
'   End If
   '
'3  Call FRESHECHO("PRESREPA")
'   PRESURE09.Label18 = ""
'   Call SELECHO("PRESREPA/Indice General Planillas de Reparación")
'   NPX& = Val(VALACT1$("PRESREPA"))
'   If NPX& < 1 Then Exit Sub
'   '
'   For UI& = ULTREG&("PRESREPA") To 1 Step -1
'     XX$ = REGLEIDO$("PRESREPA", UI&)
'     If CVS(Left$(XX$, 4)) = NPX& Then GoTo 10
'   Next UI&
'   Call COMUNI("Error de Indices - Consulte.")
'   Exit Sub
'   '
'10 REPRESUP = TRIM$(Str$(UI&))
'   XX$ = REGLEIDO$("PRESREPA", Val(REPRESUP))
'   '
'   If CVS(Mid$(XX$, 125, 4)) > 0.5 Then
'     Call COMUNI("Imposible Modificar este Presupuesto.\Corresponde O.Venta " + TRIM$(Str$(CVS(Mid$(XX$, 125, 4)))))
'     Exit Sub
'   End If
'   '
'   PRESURE07.Label18 = TRIM$(Str$(NPX&))
'   PRESURE07.Show 1
'   GoTo 3
   '
End Sub

Sub OREPARA()
   '
   If DERACCE%("OREPARA", "Generar / Cerrar Orden de Reparación") < 2 Then
     Exit Sub
   End If
   '
3  Screen.MousePointer = 11
   Call COPYSTRU("PRESREPA", "PREPAEMI")
   JJ& = 0
   For U& = 1 To ULTREG&("PRESREPA")
     XX$ = REGLEIDO$("PRESREPA", U&)
     If CVI(Mid$(XX$, 121, 2)) > 0 Then
       If CVS(Mid$(XX$, 125, 4)) < 0.5 Then
         JJ& = JJ& + 1
         Call GRAREG("PREPAEMI", XX$, JJ&)
       End If
     End If
   Next U&
   Call SETULTREG("PREPAEMI", JJ&)
   Call CIERRARCH("PREPAEMI")
   Screen.MousePointer = 1
   '
   If JJ& < 1 Then
     Call COMUNI("No Hay Presupuestos Pendientes\para Emitir Orden de Reparación.")
     Exit Sub
   End If
   '
   Call FRESHECHO("PREPAEMI")
   Call SELECHO("PREPAEMI/Indice Presupuestos de Reparación Pendientes")
   NPX& = Val(VALACT1$("PREPAEMI"))
   If NPX& < 1 Then Exit Sub
   '
   For UI& = ULTREG&("PRESREPA") To 1 Step -1
     XX$ = REGLEIDO$("PRESREPA", UI&)
     If CVS(Left$(XX$, 4)) = NPX& Then GoTo 10
   Next UI&
   Call COMUNI("Error de Indices - Consulte.")
   Exit Sub
   '
10 REPRESUP = TRIM$(Str$(UI&))
   XX$ = REGLEIDO$("PRESREPA", Val(REPRESUP))
   If CVI(Mid$(XX$, 121, 2)) < 1 Then
      Call COMUNI("Falta Generar y Emitir Presupuesto.")
      PRESURE07.Label18 = TRIM$(Str$(NPX&))
      PRESURE07.Show 1
      GoTo 3
   End If
   '
   If CVS(Mid$(XX$, 125, 4)) > 0.5 Then
      Call COMUNI("Orden de Reparación ya Emitida.")
      GoTo 3
   End If
   '
   ORDEREPA07.Label18 = ""
   ORDEREPA07.Label18 = TRIM$(Str$(NPX&))
   ORDEREPA07.Show 1
   GoTo 3
   '
End Sub

Sub CIERROREPA()
   '
   If DERACCE%("OREPARA", "Generar / Cerrar Orden de Reparación") < 2 Then
     Exit Sub
   End If
   '
3  Screen.MousePointer = 11
   Call COPYSTRU("PRESREPA", "OREPARPE")
   JJ& = 0
   For U& = 1 To ULTREG&("PRESREPA")
     XX$ = REGLEIDO$("PRESREPA", U&)
     If CVS(Mid$(XX$, 125, 4)) >= 0.5 Then
       If CVI(Mid$(XX$, 123, 2)) < 1 Then
         JJ& = JJ& + 1
         Call GRAREG("OREPARPE", XX$, JJ&)
       End If
     End If
   Next U&
   Call SETULTREG("OREPARPE", JJ&)
   Call CIERRARCH("OREPARPE")
   Screen.MousePointer = 1
   '
   If JJ& < 1 Then
     Call COMUNI("No hay Ordenes de Reparación Pendientes\que Puedan Cerrarse.")
     Exit Sub
   End If
   '
   Call SELECHO("OREPARPE/Indice de Ordenes de Reparación en Curso")
   NPX& = Val(VALACT1$("OREPARPE"))
   If NPX& < 1 Then Exit Sub
   '
   For UI& = ULTREG&("PRESREPA") To 1 Step -1
     XX$ = REGLEIDO$("PRESREPA", UI&)
     If CVS(Left$(XX$, 4)) = NPX& Then GoTo 10
   Next UI&
   Call COMUNI("Error de Indices - Consulte.")
   Exit Sub
   '
10 REPRE& = Str$(UI&)
   XX$ = REGLEIDO$("PRESREPA", REPRE&)
   If CVI(Mid$(XX$, 123, 2)) > 0 Then
      Call COMUNI("Orden de Reparación ya Cerrada")
      GoTo 3
   End If
   '
   If CVS(Mid$(XX$, 125, 4)) < 0.5 Then
      Call COMUNI("Falta Generar Orden de Reparación")
      ORDEREPA.Label18 = TRIM$(Str$(NPX&))
      ORDEREPA.Show 1
      GoTo 3
   End If
   '
   NSERIE$ = Mid$(XX$, 71, 15)
   NCLIE% = CVI(Mid$(XX$, 67, 2))
   TALOX$ = TRIM$(Mid$(XX$, 95, 6))
   NROPE& = CVS(Mid$(XX$, 125, 4))
   VDEF$ = XX$ + String$(128, 0)
   HOII% = HOY(HOS$)
   Call REPLA(VDEF$, MKI$(HOII%), 123, 2)
   DepOrig% = Val(CONTROL$("", "DEPOCEN"))
   Call REPLA(VDEF$, Chr$(DepOrig%), 129, 1)
20 OBX$ = OBJPLA$("CIERROREPA", VDEF$)
   If OBX$ = "" Then
     Exit Sub
   End If
   '
   DP2% = Asc(Mid$(OBX$, 129, 1))
   If DP2% < 1 Or TRIM$(ECOARCH$("ECODEP", Chr$(DP2%))) = "" Then
      Call MENSERR(24, "Falta Depósito de Consumo")
      GoTo 20
   End If
   FECIE% = CVI(Mid$(OBX$, 123, 2))
   '
   Screen.MousePointer = 11
   Call SETULTREG("!MATEROF", 0)
   NPI% = NPX& Mod 10000
   RFF$ = RECFILT$("REPUREPA", 2, MKI$(NPI%))
   '
   URE& = ULTREG&("REPUREPA")
30 For KK& = 1 To Len(RFF$) Step 4
     REG& = CVS(Mid$(RFF$, KK&, 4))
     If REG& > 0 Then
       If REG& <= URE& Then
         XX$ = REGLEIDO$("REPUREPA", REG&)
         If CVS(Left$(XX$, 4)) = NPX& Then
           CII% = CVI(Mid$(XX$, 7, 2))
           CANTI = CVS(Mid$(XX$, 9, 4))
           '
           ZZ$ = REGBLAN$("MATEROF")
           Call REPLA(ZZ$, MKI$(CII%), 1, 2)
           Call REPLA(ZZ$, UNIMED$(CII%), 3, 2)
           Call REPLA(ZZ$, MKS$(CANTI), 5, 4)
           Call REPLA(ZZ$, MKD$(CDbl(CANTI)), 9, 8)
           Call REPLA(ZZ$, MKS$(REG&), 17, 4)
           STDIS = STOCKDIS(CII%)
           Call REPLA(ZZ$, MKS$(STDIS), 21, 4)
           FALTA = CANTI - STDIS: If FALTA < 0 Then FALTA = 0
           Call REPLA(ZZ$, MKS$(FALTA), 25, 4)
           Call REGAPP("!MATEROF", ZZ$)
           '
         End If
       End If
     End If
   Next KK&
   Call CIERRARCH("!MATEROF")
   Screen.MousePointer = 1
   '
   TITUPA$ = "Consumo de Repuestos O.Rep." + TRIM$(Str$(NPX&))
   CIXI% = CVI(Mid$(OBX$, 69, 2))
   CODPRO$ = CODEXT$(CIXI%)
   TITUPA$ = TITUPA$ + " - " + TRIM$(CODPRO$) + " - " + TRIM$(DESCRIT$(CIXI%)) + " - Cant.=" + Str$(Int(CACONJU# + 0.5))
   If NCLIE% > 0 Then TITUPA$ = TITUPA$ + " - " + RASOCLI$(NCLIE%)
   '
   VTB% = VENTABU%("MATEROF/TITUPAN=" + TITUPA$)
   If VTB% < 0 Then
      Exit Sub
   End If
   '
   Call CIERRARCH("MOVISTO")
   Call BLOQARCH("MOVISTO")
   Call BLOQARCH("REPUREPA")
   Call BLOQARCH("PRESREPA")
   '
   XX$ = REGLEIDO$("PRESREPA", REPRE&)
   Call REPLA(XX$, MKI$(FECIE%), 123, 2)
   Call GRAREG("PRESREPA", XX$, REPRE&)
   Call CIERRARCH("PRESREPA")
   '
   DOPRI$ = "OR-" + TRIM$(Str$(NPX&))
   DOSEC$ = "TL-" + TRIM$(TALOX$)
   REFE$ = "Reparación " + CODPRO$
   '
   For U& = 1 To ULTREG&("!MATEROF")
     ZZ$ = REGLEIDO$("!MATEROF", U&)
     '
     CI2% = CVI(Mid$(ZZ$, 1, 2))
     CAN = (-1) * CVS(Mid$(ZZ$, 5, 4))
     '
     USOUNI# = CAN
     RERES& = CVS(Mid$(ZZ$, 17, 4))
     If CI2% > 0 Then
       If CI2% <= NUMAS% Then
         If Abs(CAN) >= 0.05 Then
           LOTE$ = ""
           CMOV$ = "RP"
           '
           If Abs(CAN) >= 0.05 Then
              DEPX% = DP2%: If DP2% = 99 Then DEPX% = DEPOPRE%(CI2%)
              'Call MOVISTOK(FECIE%, CI2%, LOTE$, CMOV$, DOPRI$, DOSEC$, REFE$, VUNI#, MONE$, NC%, DEPX%, CAN)
              Call MOVISTOK(FECIE%, CI2%, LOTE$, CMOV$, DOPRI$, DOPRI$, DOSEC$, DOSEC$, REFE$, NUORIT%, VUNI#, 1, NC%, DEPX%, CAN)
           End If
           REGMOVI& = ULTREG&("MOVISTO")
           '
           If RERES& > 0 Then
             If RERES& <= ULTREG&("REPUREPA") Then
               YY$ = REGLEIDO$("REPUREPA", RERES&)
               CARESE = CVS(Mid$(YY$, 21, 4))
               Call REPLA(YY$, MKS$((-1) * CAN), 25, 4)
               Call GRAREG("REPUREPA", YY$, RERES&)
               GoTo 99
             End If
           End If
           '
           YY$ = REGBLAN$("REPUREPA")
           Call REPLA(YY$, MKS$(NPX&), 1, 4)
             NPI% = NPX& Mod 10000
           Call REPLA(YY$, MKI$(NPI%), 5, 2)
           Call REPLA(YY$, MKI$(CI2%), 7, 2)
           Call REPLA(YY$, MKS$((-1) * CAN), 25, 4)
           Call REPLA(YY$, "OR-" + TRIM$(Str$(NPX&)), 33, 12)
           Call REGAPP("REPUREPA", YY$)
         End If
       End If
     End If
99 Next U&
   Call CIERRARCH("MOVISTO")
   Call CIERRARCH("REPUREPA")
   '
   ' CANCELAR TODAS LAS RESERVAS
   Call ABRESERVA
   SQLX$ = "SELECT * FROM RESERVA"
   SQLX$ = SQLX$ + " WHERE IdSubOr = '" & DOPRI$ & "' "
   'Set RESELE = Stocks.OpenRecordset(SQLX$, dbOpenDynaset)
   Dim RESELE As ADODB.Recordset
   Set RESELE = New ADODB.Recordset
   RESELE.Open FILTSQL$(SQLX$), FLEXCONN, adOpenDynamic, adLockPessimistic, adCmdText

   With RESELE
101   On Error Resume Next
      .MoveFirst
      If Err < 1 Then
         On Error Resume Next
         .Delete
         If Err Then GoTo 102
         GoTo 101
      End If
      On Error GoTo 0
   End With
102 RESELE.Close
   '
   HOII% = HOY(HOS$)
   TTXX$ = Space$(96)
   Call REPLA(TTXX$, HOS$ + "  " + "Cierre Orden de Reparación " + TRIM$(Str$(NPX&)), 1, 56)
   Call REPLA(TTXX$, USERTER$, 69, 24)
   Call REGAPP("ANOTAREP", Left$(MKS$(NPX&) + TTXX$, 96))
   Call REGAPP("ANOTAREP", MKS$(NPX&) + String$(88, "="))
   Call CIERRARCH("ANOTAREP")
   '
   TTXX$ = Space$(96)
   Call REPLA(TTXX$, HOS$ + "  " + "Cierre O/Rep " + TRIM$(Str$(NPX&)) + " - Talón " + TRIM$(TALOX$), 1, 56)
   Call REPLA(TTXX$, USERTER$, 69, 24)

   Call AGREGA_ANOTAPED(NROPE&, TTXX$)
   Call AGREGA_ANOTAPED(NROPE&, String$(88, "="))
   '
   'GRABACION EN HISTORIAL DE NUMERO DE SERIE
   '
   If NSERIE$ <> "" Then
    CUERPO$ = "Cierre de Orden de Reparación"
    NSERIE$ = TRIM$(Text1.TEXT)
    Call ANOTANUSE(NSERIE$, CUERPO$, , "NOMESS")
   End If
   '
   SQLX$ = "UPDATE PEDDETA SET StaDespa = 1, AcuPrepa = 1 "
   SQLX$ = SQLX$ + " WHERE NroPed = " & NROPE& & " "   ' COINCIDE NÚMERO DE PEDIDO
   SQLX$ = SQLX$ + " AND CodiInt = " & CIXI% & " AND NroClie = " & NCLIE% & " "
   '
   On Error Resume Next
   FLEXCONN.Execute SQLX$
   If Err Then
      On Error GoTo 0
      Call MENSERR(24, "No se Encontró la Nota de Venta\Correspondiente a esta Reparación.\  \Consulte a su Soporte de Sistemas.")
   End If
   On Error GoTo 0
   '
90 Call CIERRARCH("DESPAPED")
   Call CIERRARCH("*.*")
   Call BAJALDISCO
   '
   GoTo 3
   '
End Sub
'
Function PRONUOCO&()
   '
   PRONU& = 1
   REGXX& = ULTREG&("OCOMENCA")
   If REGXX& > 0 Then
     PRONU& = 1 + CVS(Left$(REGLEIDO$("OCOMENCA", REGXX&), 4))
   End If
   If PRONU& <= REGXX& Then
     Screen.MousePointer = 11
     For U& = 1 To REGXX&
       X0$ = REGLEIDO$("OCOMENCA", U&)
       If CVS(Left$(X0$, 4)) >= PRONU& Then
         PRONU& = 1 + CVS(Left$(X0$, 4))
       End If
     Next U&
     Screen.MousePointer = 1
   End If
   '
   PRONUOCO& = PRONU&
   '
End Function

Function PRONUOTER&()    ' numero orden de servicio terceros
   '
   PRONU& = 1
   REGXX& = ULTREG&("OPERTER")
   If REGXX& > 0 Then
     PRONU& = 1 + CVS(Mid$(REGLEIDO$("OPERTER", REGXX&), 75, 4))
   End If
   If PRONU& <= REGXX& Then
     Screen.MousePointer = 11
     For U& = 1 To REGXX&
       X0$ = REGLEIDO$("OPERTER", U&)
       If CVS(Mid$(X0$, 75, 4)) >= PRONU& Then
         PRONU& = 1 + CVS(Mid$(X0$, 75, 4))
       End If
     Next U&
     Screen.MousePointer = 1
   End If
   '
   PRONUOTER& = PRONU&
   '
End Function



Private Sub Command8_Click()
   Command8.Enabled = False
   Call ABREORFAB
   '
   Call COPYSTRU("INDIPRES", "INDRECFI")
   
10 Call CARGALISEL("INDRECFI", "REPARA", "NUMREPA/F6.0;REFEREN/A56", "NUMREPA > 0 AND STATUS >= 3", "DISTINCT")
   Call SELECHO("INDRECFI/Indice General de Reparaciónes ")
   NROREP& = XVALO(TRIM$(VALACT1("INDRECFI")))
   '
   If NROREP& < 1 Then GoTo 99
   '
   NUSERIE$ = VALOBADA("REPARA", "NUMEROSERIE", "NUMREPA = '" & NROREP& & "'", "NOMESS")
   
   RECONRE07.Label24 = NUSERIE$
   RECONRE07.Show 1
   GoTo 10
99 Command8.Enabled = True
End Sub


Private Sub Command9_Click_VIE()
    '
    If DERACCE%("NOTAREP", "Nota de Reparación") < 2 Then
      Exit Sub
    End If

    Call ABREORFAB
    Call CREATABLA_EQUIPOS

10  OPX% = ALTERNA%("Nueva Reparación\Corregir Datos")
    If OPX% < 1 Or OPX% > 2 Then Exit Sub
    '
    If OPX% = 1 Then
        'NUEVO
        Call SELECHO("DEUDOR1")
        NC% = XVALO(TRIM$(VALACT1$("DEUDOR1")))
        If NC% < 1 Then GoTo 10 ' ESTO SE ESTA VIENDO EL FRM COMO MODIFICAR PARA
        '                       ' QUE PRESENTE LA INFORMACION DEL DEUDOR12 Y GRABE EN REPARA
        CONDI$ = "NUMREPA > 0"
        NROREP& = 1 + XVALO(VALOBADA("REPARA", "MAX(NUMREPA)", CONDI$, "NOMESS"))
        FREP = HOY(HO$)
      ElseIf OPX% = 2 Then
        'CORREGIR DATOS
        Call COPYSTRU("INDIPRES", "INDERPCI")
15      Call CARGALISEL("INDERPCI", "REPARA", "NUMREPA/F6.0;REFEREN/A56", "NUMREPA > 0  AND STATUS < 4", "DISTINCT")
        Call SELECHO("INDERPCI/Indice General de Reparación")
        NROREP& = XVALO(TRIM$(VALACT1("INDERPCI")))
        If NROREP& < 1 Then GoTo 10
        NC% = XVALO(VALOBADA("REPARA", "NUME_CLI", "NUMREPA = " & NROREP& & ""))
        FREP = CVINT(VALOBADA("REPARA", "FechRecep", "NUMREPA = " & NROREP& & ""))
        If FREP < 1 Then FREP = HOY(HO$)
    End If
    '
    FILTX$ = TRIM$(Str$(NROREP&)) & ";" & TRIM$(Str$(FREP)) & ";" & TRIM$(Str$(NC%))
    Call STDFORM("NREP-MDT", FILTX$)
    If STDSCREEN.APRO <> "OK" Then
       If OPX% = 2 Then GoTo 15
       Exit Sub
    End If
    '
'AQUI INTRODUCIR EL CODIGO INTERNO Y ARMAR EL TEXTO DE LA REFERENCIA
    DESKRIPCION$ = VALOFORM("Text-4")
    CI% = 0
    NUMESER$ = VALOFORM("Text-7")
    '
    CONDI$ = "NUMEROSERIE = '" & NUMESER$ & "'"
    CI% = XVALO(VALOBADA("TANUMSE", "COD_INTE", CONDI$, "NOMESS"))
    If CI% > 0 Then
       DESKRIPCION$ = DESCRIT$(CI%)
    End If
    '
    'VERIFICAR SI EL EQUIPO SE ENCUENTRA EN GARANTIA
    FE% = HOY(HO$)
    ESTREP% = 0
    '
    FECHCARANT% = FeVenGarant(NUMESER$)
    If FECHCARANT% > 0 Then
      If FE% < FECHCARANT% Then
         ESTREP% = 1
         Call COMUNI("Equipo en Garantía")
         GoTo 20
      End If
    End If
    '
    FECHAGARAREP% = FeVenGarantRep(NUMESER$) ' GARANTIA POR REPARACION
    If FECHAGARAREP% > 0 Then
       If FE% < FECHAGARAREP% Then
         ESTREP% = 2
         Call COMUNI("Reparación en Garantía")
       End If
    End If
    '
20  Call ACTUREGISTRO("REPARA", "COD_INTE", "NUMREPA = " & NROREP& & "", CI%, REGAF&, "NOMESS")
    Call ACTUREGISTRO("REPARA", "DESCRIPCION", "NUMREPA = " & NROREP& & "", DESKRIPCION$, REGAF&, "NOMESS")
    Call ACTUREGISTRO("REPARA", "Status_Gar", "NUMREPA = '" & NROREP& & "'", ESTREP%, REGAF&, "NOMESS")
    '
    RACLI$ = VALOFORM("Param-5")
    REFER$ = TRIM$(Left$(DESKRIPCION$, 22)) & " - " & RACLI$
    '
    Call ACTUREGISTRO("REPARA", "REFEREN", "NUMREPA = " & NROREP& & "", REFER$, REGAF&, "NOMESS")
    Call ACTUREGISTRO("REPARA", "STATUS", "NUMREPA = " & NROREP& & "", "0", REGAF&, "NOMESS")
    '
    CUERPO$ = "Recepción del Equipo para Reparación" & vbCrLf
    If OPX% = 2 Then ' CORREGIR DATOS
       CUERPO$ = "Modificacion de Recepción de Equipo para Reparación" & vbCrLf
    End If
    CUERPO$ = CUERPO$ + "Número de Reparación: " & NROREP& & vbCrLf
    CUERPO$ = CUERPO$ + "Cliente: " & RACLI$
    Call ANOTANUSE(NUMESER$, CUERPO$, , "NOMESS")

    
    If OPX% = 2 Then GoTo 15
    '
    Call COMUNI("Equipo Ingresado Para Presupuestación")
    '
End Sub

Private Sub Command9_Click()
    '
    Command9.Enabled = False
    Call ABREORFAB
    Call CREATABLA_EQUIPOS
    '
    If DERACCE%("NOTAREP", "Nota de Reparación") > 1 Then
       Call RECEP_EQUIP
    End If
    '
    Command9.Enabled = True
    '
End Sub

Private Sub Exit_Click()
   Call Command1_Click
End Sub

Private Sub Form_Load()
    '
    Show
    Refresh
    '
    Call VERJOBID(OKEY%)
    If OKEY% < 1 Then Close: End
    '
    EPAC$ = TRIM$(EMPREAC$)
    If EPAC$ <> "" Then
        Caption = Caption + "  -  " + EPAC$
    End If
    Call GRATITFOR(Caption)
    '
    Call ABRECONEXION
    Call CREATABLA_OFABRIC
    '
    Call CREACAMPO("OFABRIC", "REPARA", "SecRepar", 10, 48)
    Call CREACAMPO("OFABRIC", "REPARA", "Usu_Recep", 10, 48)
    Call CREACAMPO("OFABRIC", "REPARA", "Ctac_Cli", 10, 48)
    Call CREACAMPO("OFABRIC", "REPARA", "Reacondi", 12, 0)
    Call CREACAMPO("OFABRIC", "REPARA", "Tec_Repara", 10, 64)
    Call CREACAMPO("OFABRIC", "REPARA", "DeTipFalla", 12, 0)
    Call CREACAMPO("OFABRIC", "REPARA", "DeCauFalla", 12, 0)
    '
    Call CREACAMPO("OFABRIC", "TIPOFALL", "NUMREPA", 4, 0, 1)
    Call CREACAMPO("OFABRIC", "TIPOFALL", "ORIGEN", 3, 0, 1)     ' 1 CLIENTE; 2 TECNICO
    Call CREACAMPO("OFABRIC", "TIPOFALL", "CODIFALL", 4, 0, 1)
    Call CREACAMPO("OFABRIC", "TIPOFALL", "DESCFALL", 10, 128, 0)
    '
    Call CREACAMPO("OFABRIC", "CAUSFALL", "NUMREPA", 4, 0, 1)
    Call CREACAMPO("OFABRIC", "CAUSFALL", "CODIFALL", 4, 0, 1)
    Call CREACAMPO("OFABRIC", "CAUSFALL", "DESCFALL", 10, 128, 0)
    '
    Call CREACAMPO("OFABRIC", "ACCEREP", "NUMREPA", 4, 0, 1)
    Call CREACAMPO("OFABRIC", "ACCEREP", "DESCRIACC", 10, 128, 0)
    Call CREACAMPO("OFABRIC", "ACCEREP", "CANTIDAD", 4, 0, 1)
    Call CREACAMPO("OFABRIC", "ACCEREP", "OBSERVA", 10, 128)
    '
End Sub

Function NUPRESUP&()
   '
   NPP& = 1
   '
   FIN& = ULTREG&("PRESREPA")
   If FIN& > 0 Then
      For U& = FIN& To FIN& - 128 Step -1
         If U& < 1 Then GoTo 99
         XX$ = REGLEIDO$("PRESREPA", U&)
         NPPX& = CVS(Mid$(XX$, 1, 4))
         If NPPX& >= NPP& Then
            NPP& = 1 + NPPX&
         End If
      Next U&
   End If
   '
99 NUPRESUP& = NPP&
   '
End Function
'

Private Sub Genpresu_Click()
   Call Command3_Click
End Sub

Private Sub Lispre_Click()
   Call CONFREPA07.Command5_Click
End Sub

Private Sub Master_Click()
    Call CIERRARCH("*.*")
    ACONEC$ = "ARCHIG07/AMASTO"
    Call CONECRUN(ACONEC$, "Maestro de Artículos de Stock")

'   Call PLAPRERE07.Command8_Click
End Sub

Private Sub Mater_Orep_Click()
   Option5.Value = True
   Call Command7_Click
End Sub

Private Sub Oper_Orep_Click()
   Option6.Value = True
   Call Command7_Click
End Sub

Private Sub Option1_DblClick()
   Call PRESEDOC
End Sub

Private Sub Option2_DBLClick()
   Call PRESEDOC
End Sub

Private Sub Option3_DblClick()
   Call PRESEDOC
End Sub

Private Sub Option4_DblClick()
   Call PRESEDOC
End Sub

Private Sub Option5_DblClick()
    Call ABREORFAB
    '
    Call COPYSTRU("INDIPRES", "INDIRREM")
15  Call CARGALISEL("INDIRREM", "REPARA", "NUMREPA/F6.0;REFEREN/A56", "STATUS = 5", "DISTINCT")
    Call SELECHO("INDIRREM")
    NROREP& = XVALO(TRIM$(VALACT1("INDIRREM")))
    If NROREP& < 1 Then Exit Sub
    '
    REMI$ = VALOBADA("REPARA", "NUMREMITO", "NUMREPA = " & NROREP&, "NOMESS")
    CODXT$ = ""
    CI% = XVALO(VALOBADA("REPARA", "COD_INTE", "NUMREPA = " & NROREP&, "NOMESS"))
    If CI% > 0 Then
       CODXT$ = CODEXT$(CI%)
    End If
    '
    FREP = CVINT(VALOBADA("REPARA", "FECHENT", "NUMREPA = " & NROREP&, "NOMESS"))
    '
    FILTX$ = TRIM$(Str$(NROREP&)) & ";" & TRIM$(REMI$) & ";" & TRIM$(Str$(FREP)) & ";" & TRIM$(CODXT$)
    Call STDFORM("REMX-MDT", FILTX$, "PRINT")
    '
End Sub


Private Sub Option6_DblClick()
   'Presentar Cargalisel
   Call CARGALISEL("ESTAREPA", "REPARA", "NUMREPA/F6.0;REFEREN/A32;NUMEROSERIE/A14", "NUMREPA > 0 ", "DISTINCT")
10 Call SELECHO("ESTAREPA/Indice General de Reparaciónes Generadas")
   NROREP& = XVALO(TRIM$(VALACT1("ESTAREPA")))
   
   If NROREP& < 1 Then Exit Sub
   '
   'ESTADO
   '*********
   EST% = XVALO(VALOBADA("REPARA", "STATUS", "NUMREPA = " & NROREP& & "", "NOMESS"))
   Select Case EST%
      Case 0
        ESTADO$ = "Ingresado"
      Case 1
        ESTADO$ = "Planilla de Presupuestación"
      Case 2
        ESTADO$ = "Generación de Presupuesto"
      Case 3
        ESTADO$ = "Generación de O.Reparación"
      Case 4
        ESTADO$ = "Orden de Reparación Cerrada"
      Case 5
        ESTADO$ = "Generación de Remito"
   End Select
   '
   'PRESENTAR FRM DE ESTADO
   '**********************
   FILTX$ = TRIM$(Str$(NROREP&)) & ";" & TRIM$(ESTADO$)
   Call STDFORM("NSTA-MDT", FILTX$)
   
   '
   GoTo 10
   'presentar standscreen
   
End Sub

Private Sub Orderep_Click()
   Option4.Value = True
'   Call Command7_Click
End Sub

Private Sub Ordrep_Click()
   Call Command4_Click
End Sub

Private Sub Plabuscli_Click()
   Call Command20_Click
End Sub

Private Sub Pladir_Click()
   Call Command5_Click
End Sub

Private Sub Plarepa_Click()
   Option1.Value = True
'   Call Command7_Click
End Sub

Private Sub Presucli_Click()
   Option3.Value = True
   Option3_DblClick
End Sub

Private Sub Presuint_Click()
   Option2.Value = True
   Option2_DBLClick
End Sub

Private Sub RtoRep_Click()
   Call Option5_DblClick
End Sub

Private Sub Setup_Click()
   Call Command6_Click
End Sub

Private Sub ESTAREP()
  '
  Call DESHASFECHA(DES%, HAS%, PERIO$)
  '
  If TRIM$(PERIO$) = "" Then Exit Sub 'SI CANCELO
  '
  Screen.MousePointer = 11
  FIN& = ULTREG&("PRESREPA")
  JJ& = 0
  For U& = 1 To FIN&
    Call TRACE(20, U&, FIN&)
    XX$ = REGLEIDO$("PRESREPA", U&)
    FEIN1% = CVI(Mid$(XX$, 65, 2))   ' FECHA DE RECEPCION
    '
    If FEIN1% >= DES% Then
      If FEIN1% <= HAS% Then
        JJ& = JJ& + 1
        NO1% = JJ&
        NUREP1& = CVS(Mid$(XX$, 1, 4))
        CLIE1% = CVI(Mid$(XX$, 67, 2))
        CI1% = CVI(Mid$(XX$, 69, 2))
        TALON$ = TRIM$(Mid$(XX$, 95, 6))
        FEPRE1% = CVI(Mid$(XX$, 121, 2))   ' FECHA PRESUPUESTO
        FECIE1% = CVI(Mid$(XX$, 123, 2))   ' FECHA DE CIERRE
        NUOVEN1 = CVS(Mid$(XX$, 125, 4))
        '
        NUOVE& = XVALO(NUOVEN1)
        'OBTENGO LA FECHA DE LA O/VEN
'        XZ1$ = FILTERGETRECORD$("PEDENCA", 1, MKS$(NUOVEN1))
'        FEOREP1% = CVI(Mid$(XZ1$, 53, 2))  ' FECHA DE APROBACION
        '
        '\\\OT-08-1033-OD-11/11/08
        Call ABREPEDCLI
        FEOREP1% = CVINT(VALOBADA("PEDENCA", "FechEmis", "NroPed = " & NUOVE& & ""))
        '\\\OT-08-1033-OD-FIN
        '
        YY$ = REGBLAN$("ESTAREP")
        '
        Call REPLA(YY$, MKI$(NO1%), 1, 2)
        Call REPLA(YY$, MKS$(NUREP1&), 3, 4)
        Call REPLA(YY$, MKI$(CLIE1%), 7, 2)
        Call REPLA(YY$, MKI$(CI1%), 9, 2)
        Call REPLA(YY$, MKI$(FEIN1%), 11, 2)
        Call REPLA(YY$, MKI$(FEPRE1%), 13, 2)
        Call REPLA(YY$, MKI$(FEOREP1%), 15, 2)
        Call REPLA(YY$, MKS$(NUOVEN1), 17, 4)
        Call REPLA(YY$, MKI$(FECIE1%), 21, 2)
        Call REPLA(YY$, TALON$, 23, 6)
        '
        Call GRAREG("ESTAREP", YY$, JJ&)
      End If
    End If
  Next U&
  '
  Call SETULTREG("ESTAREP", JJ&)
  Call CIERRARCH("ESTAREP")
  Screen.MousePointer = 1
  '
  If JJ& > 0 Then
    Call HEADERS("ESTAREP", "")
    Call HEADERS("ESTAREP", "Período: " & PERIO$)
    Call FINAL("*ESTAREP")
  End If
  '
End Sub
Private Sub OREPABI()
   '
   FIN& = ULTREG&("PRESREPA")
   For U& = 1 To FIN&
     Call TRACE(20, UI&, FIN&)
     XX$ = REGLEIDO$("PRESREPA", U&)
     If CVS(Mid$(XX$, 125, 4)) >= 0.5 Then
        If CVI(Mid$(XX$, 123, 2)) < 1 Then
           JJ& = JJ& + 1
           Call GRAREG("OREPARPE", XX$, JJ&)
        End If
     End If
   Next U&
   Call SETULTREG("OREPARPE", JJ&)
   Screen.MousePointer = 1
   '
   Call FINAL("*LIREPAPE")
   '
End Sub

Sub DEVOPORREP()
   '
   Dim MATRANSFE#(32767)
   '
   Call ABREORFAB
   Call COPYSTRU("INDIPRES", "INDERTRA")
10 Call CARGALISEL("INDERTRA", "REPARA", "NUMREPA/F6.0;REFEREN/A56", "NUMREPA > 0 AND STATUS = 3", "DISTINCT")
   Call SELECHO("INDERTRA/Indice General de Reparaciónes Generadas")
   NROREP& = XVALO(TRIM$(VALACT1("INDERTRA")))
   If NROREP& < 1 Then Exit Sub
   '
   NUMEREPA$ = "OR-" & TRIM$(Str$(NROREP&))
   '
   For t& = 1 To NUMAS%
      MATRANSFE#(t&) = 0
   Next t&
   '
   If CONTROL$("", "DEPREDES") <> "" Then
     DEPORIGINAL% = CONTROL$("", "DEPREDES")
   End If
   '
   Call REPLA(VDEF$, NUMEREPA$, 1, 12)
   Call REPLA(VDEF$, Chr$(DEPORIGINAL%), 13, 1) ' DEPOSITO TALLER
   '
12 OBX$ = OBJPLA$("SEREPDEP", VDEF$)
   If OBX$ = "" Then GoTo 10
   '
   DepOrig% = Asc(Mid$(OBX$, 13, 1))
   DEPODEST% = Asc(Mid$(OBX$, 14, 1))
   '
'   'VALIDACION QUE TENGA DERECHOS PARA CAMBIAR DEPOSITOS
   If DERACCE%("CAMBIDEP/NOMESS", "CAMBIO DE DEPOSITOS") < 2 Then
     If DEPORIGINAL% <> DEPODEST% Then
        Call MENSERR(24, "Usuario no Habilitado para Modificar Depósito\Módulo: CAMBIDEP")
        GoTo 12
     End If
   End If
   '
   'VALIDA DEPOSITOS
   If DepOrig% < 1 Or ECOARCH$("TADEPOSI", Chr$(DepOrig%)) = "" Then 'VALIDACION
     Call MENSERR(24, "Depósito Origen No Válido")
     VDEF$ = OBX$
     GoTo 12
   End If
   '
   If DEPODEST% < 1 Or ECOARCH$("TADEPOSI", Chr$(DEPODEST%)) = "" Then
      Call MENSERR(24, "Depósito Destino No Válido")
      VDEF$ = OBX$
      GoTo 12
   End If
   '
   Screen.MousePointer = 11
   '
   ' RECORRO MOVISTO Y CARGO EL VECTOR CON LA CANTIDAD INGRESADA POR ARTICULO SEGUN LA O.F.
   SQLX$ = "SELECT * FROM MOVISTO  WITH(NOLOCK)"
   SQLX$ = SQLX$ + " WHERE DOSECCOR = '" & NUMEREPA$ & "' "
   SQLX$ = SQLX$ + " AND CODIMOVI = 'TR' "
   SQLX$ = SQLX$ + " AND DEPOMOVI = " & DepOrig% & ""
   Set MoviTemp = FLEXCONN.Execute(FILTSQL$(SQLX$))
   With MoviTemp
     Do While Not .EOF
       CI% = XVALO(!COD_INTE)
       CAN = XVALO(!CantIngre) - XVALO(!Cantsalid)
       '
       MATRANSFE#(CI%) = MATRANSFE#(CI%) + CAN
       .MoveNext
     Loop
   End With
   MoviTemp.Close
   Set MoviTemp = Nothing
   '
   JJ& = 0
   Call SETULTREG("!MATEROF", 0)
   For I% = 1 To NUMAS%
     CI% = I%
     UM1$ = UNIMED$(CI%)
     CAN1 = MATRANSFE#(CI%)
     If MATRANSFE#(CI%) > 0.000005 Then
        JJ& = JJ& + 1
        Y$ = REGBLAN$("!MATEROF")
        Call REPLA(Y$, MKI$(CI%), 1, 2)
        Call REPLA(Y$, UM1$, 3, 2)
        Call REPLA(Y$, MKS$(CAN1), 5, 4)
        Call GRAREG("!MATEROF", Y$, JJ&)
     End If
   Next I%
   If JJ& < 1 Then
      Call COMUNI("No se Encontró Material sin Consumir Para Esta Orden de Fabricación")
      Screen.MousePointer = 1
      GoTo 10
   End If
   
   Screen.MousePointer = 1
   VTB% = VENTABU%("MATEROF1")
   
   If VTB% < 0 Then GoTo 10
   
   OPPX% = COMALTER%("Confirmación de Devolución\\Realizar Devolución\Cancelar")
   If OPPX% < 1 Or OPPX% > 1 Then
       GoTo 10
   End If
   '
   'Devolucion de Materiales
   CMOV$ = TRIM$(CONTROL$("", "CODVALMA"))
   If CMOV$ = "" Then CMOV$ = "TR"
   HOII% = HOY(HOS$)
   FECHA% = HOII%
   LOTE$ = NUMEREPA$
'   CMOV$ = "TR"
   DOSEC$ = "TR-" + TRIM$(Mid$(NUMEREPA$, 4))
   NRO& = ProNroComprobante(CMOV$)
   DOPRI$ = "TR." + TRIM$(Str$(NRO&))
   VUNI# = 0
   MONE$ = ""
   DOSEC$ = TRIM$(NUMEREPA$) ' NUMERO DE ORDEN DE FABRICACION
   While Len(DOSEC$) < 10
      DOSEC$ = DOSEC$ + " "
   Wend

   'Recorro el Materof

   FIN& = ULTREG&("!MATEROF")
   For H& = 1 To FIN&
      YY$ = REGLEIDO$("!MATEROF", H&)
      CI% = CVI(Mid$(YY$, 1, 2))
      CAN = CVS(Mid$(YY$, 5, 4))
      If CAN > 0.05 Then
         LOTE$ = ""
         If ESTRAZABLE(CI%) = 1 Then
            CONDIX$ = "COD_INTE = " & CI% & " AND DOSECCOR = '" & NUMEREPA$ & "' AND CODIMOVI = 'TR' AND DEPOMOVI = " & DepOrig% & ""
            LOTE$ = VALOBADA("MOVISTO", "IdenLote", CONDIX$, "NOMESS")
         End If
         '
         REFE$ = "TR.de " + TRIM$(ECOARCH$("TADEPOSI", MKI$(DepOrig%)))
          '
         Call MOVISTOK(FECHA%, CI%, LOTE$, CMOV$, DOPRI$, DOPRI$, DOSEC$, DOSEC$, REFE$, NORI%, VUNI#, 0, NC%, DEPODEST%, CAN)
          '
         CAN = (-1) * CAN
         REFE$ = "TR.a " + TRIM$(ECOARCH$("TADEPOSI", MKI$(DEPODEST%)))
         Call MOVISTOK(FECHA%, CI%, LOTE$, CMOV$, DOPRI$, DOPRI$, DOSEC$, DOSEC$, REFE$, NORI%, VUNI#, 0, NC%, DepOrig%, CAN)
         
      End If
   Next H&
   '
   Screen.MousePointer = 1
   Call COMUNI("Proceso Terminado")
   
End Sub

Sub RECEP_EQUIP()             ' RECEPCION EQUIPO PARA REPARACION
   '
5  RECEPEQ09.Text1 = "."
   RECEPEQ09.Text1 = ""
   RECEPEQ09.Text2 = ""
   RECACCE09.Text1 = ""
   RECACCE09.Text2 = ""
   Call RECACCE09.Form_Load
   '
   OPX% = ALTERNA%("Nueva Reparación\Corregir Datos")
   If OPX% < 1 Or OPX% > 2 Then Exit Sub
   '
   If OPX% = 1 Then
        'NUEVO
        CONDI$ = "NUMREPA > 0"
        NROREP& = 1 + XVALO(VALOBADA("REPARA", "MAX(NUMREPA)", CONDI$, "NOMESS"))
        RECEPEQ09.Label15 = NROREP&
        RECEPEQ09.Label16 = Format(Now, "dd/mm/yy - hh:mm")
        RECEPEQ09.Label17 = USERNAME$
      ElseIf OPX% = 2 Then
        'CORREGIR DATOS
        Call COPYSTRU("INDIPRES", "INDERPCI")
15      Call CARGALISEL("INDERPCI", "REPARA", "NUMREPA/F6.0;REFEREN/A56", "NUMREPA > 0  AND STATUS < 4", "DISTINCT")
        Call SELECHO("INDERPCI/Indice General de Reparación")
        NROREP& = XVALO(TRIM$(VALACT1("INDERPCI")))
        If NROREP& < 1 Then GoTo 5
        RECEPEQ09.Label15 = NROREP&
           FREP = VALOBADA("REPARA", "FechRecep", "NUMREPA = " & NROREP& & "", "NOMESS")
           If FREP = "" Then FREP = Now
        RECEPEQ09.Label16 = Format(FREP, "dd/mm/yy - hh:mm")
           USRECEP$ = VALOBADA("REPARA", "USU_RECEP", "NUMREPA = " & NROREP& & "", "NOMESS")
           If USRECEP$ = "" Then USRECEP$ = USERNAME$
        RECEPEQ09.Label17 = USRECEP$      ' TOMAR DE LA BASE DE DATOS
        RECEPEQ09.Text2 = TRIM$(VALOBADA("REPARA", "NUMEROSERIE", "NUMREPA = " & NROREP& & "", "NOMESS"))
        RECEPEQ09.Text1 = XVALO(VALOBADA("REPARA", "NUME_CLI", "NUMREPA = " & NROREP& & "", "NOMESS"))
        RECEPEQ09.Label14 = FECHATEX$(CVINT(VALOBADA("REPARA", "FechEnSol", "NUMREPA = " & NROREP& & "", "NOMESS")))
        RECEPEQ09.Text9 = VALOBADA("REPARA", "Ctac_Cli", "NUMREPA = " & NROREP& & "", "NOMESS")
        RECEPEQ09.Text11 = VALOBADA("REPARA", "SecRepar", "NUMREPA = " & NROREP& & "", "NOMESS")
        RECEPEQ09.Label7 = ""
           If VALOBADA("REPARA", "Status_Gar", "NUMREPA = " & NROREP& & "", "NOMESS") > 0 Then
              RECEPEQ09.Label7 = "SI"
           End If
   End If
   '
10 RECEPEQ09.APRO = ""
   On Error Resume Next
   RECEPEQ09.Text2.SetFocus
   On Error GoTo 0
12 RECEPEQ09.Show 1
   If RECEPEQ09.APRO <> "NEXT" Then Exit Sub
   '
   If OPX% = 1 Then
      ' validacion de fecha de entrega solicitada solo en primera entrada.
      FESOLENTRE = FECHANUM(RECEPEQ09.Label14.Caption)
      If FESOLENTRE < HOY(HOS$) Then
         Call MENSERR(24, "La Fecha Solicitada de Entrega\Debe ser POSTERIOR a la Fecha Actual.")
         GoTo 12
      End If
   End If
   '
   RECACCE09.Label14 = RECEPEQ09.Text2
   RECACCE09.Label15 = RECEPEQ09.Label15
   RECACCE09.Label9 = RECEPEQ09.Label9
   RECACCE09.Label7 = RECEPEQ09.Label7
   RECACCE09.Label10 = RECEPEQ09.Label10
   RECACCE09.APRO = ""
   '
   If OPX% = 2 Then   ' MODIFICACION DE DATOS
       '
       RECACCE09.Text1 = VALOBADA("REPARA", "OBSERVA", "NUMREPA = " & NROREP& & "", "NOMESS")
       '
       ' tipo de falla
       SQLX$ = "SELECT * FROM TIPOFALL WHERE NUMREPA = " & NROREP& & " AND ORIGEN = 1"
       Dim FASEL As ADODB.Recordset
       Set FASEL = New ADODB.Recordset
       Set FASEL = FLEXCONN.Execute(SQLX$)
       With FASEL
          Do While Not .EOF
             NUFAL& = XVALO(!CODIFALL)
             DEFAL$ = SAFETEXT$(!DESCFALL)
             For KKI% = 1 To RECACCE09.List1.ListCount
                If XVALO(Left$(RECACCE09.List1.List(KKI% - 1), 4)) = NUFAL Then
                   RECACCE09.List1.Selected(KKI% - 1) = True
                   GoTo 14
                End If
             Next KKI%
             '
             If RECACCE09.Text2 <> "" Then RECACCE09.Text2 = RECACCE09.Text2 + vbCrLf
             RECACCE09.Text2 = RECACCE09.Text2 + DEFAL$
14        .MoveNext
          Loop
       End With
       FASEL.Close
       Set FASEL = Nothing
       '
       ' accesorios
       SQLX$ = "SELECT * FROM ACCEREP WHERE NUMREPA = " & NROREP&
       Dim FACCE As ADODB.Recordset
       Set FACCE = New ADODB.Recordset
       Set FACCE = FLEXCONN.Execute(SQLX$)
       With FACCE
          Do While Not .EOF
             CANACC& = XVALO(!Cantidad)
             If CANACC& > 0 Then
                DEACC$ = TRIM$(SAFETEXT$(!DESCRIACC))
                OBSEX$ = TRIM$(SAFETEXT$(!Observa))
                For KKI% = 2 To RECACCE09.GRID.Rows
                   TGRID$ = TRIM$(RECACCE09.GRID.TextMatrix(KKI% - 1, 0))
                   If UCase$(TGRID$) = UCase$(DEACC$) Or TGRID$ = "" Then
                      RECACCE09.GRID.TextMatrix(KKI% - 1, 0) = DEACC$
                      RECACCE09.GRID.TextMatrix(KKI% - 1, 1) = CANACC&
                      RECACCE09.GRID.TextMatrix(KKI% - 1, 2) = OBSEX$
                      GoTo 24
                   End If
                Next KKI%
                '
                RECACCE09.GRID.Rows = RECACCE09.GRID.Rows + 1
                KKI% = RECACCE09.GRID.Rows
                RECACCE09.GRID.TextMatrix(KKI% - 1, 0) = DEACC$
                RECACCE09.GRID.TextMatrix(KKI% - 1, 1) = CANACC&
                RECACCE09.GRID.TextMatrix(KKI% - 1, 2) = OBSEX$
             End If
24        .MoveNext
          Loop
       End With
       FACCE.Close
       Set FACCE = Nothing
   End If
   '
   RECACCE09.Show 1
   If RECACCE09.APRO = "PREV" Then GoTo 10
   If RECACCE09.APRO <> "NEXT" Then Exit Sub
   '
   NROREP& = XVALO(RECEPEQ09.Label15)
' AQUI VERIFICAR QUE ES MODIFICACION O NUMERO NUEVO
   CI% = CODINT%(RECEPEQ09.Label9)
   NUMESER$ = TRIM$(RECEPEQ09.Text2)
   NCXX& = XVALO(RECEPEQ09.Text1)
   TTXX$ = "Confirma Recepción de Equipo\'" + TRIM$(CODEXT$(CI%)) + " - " + DESCRIT0$(CI%) + "'\Serie Numero '" + NUMESER$ + "'\para Presupuesto de Reparación ?\\Si, Recibir\No, Cancelar"
   If COMALTER%(TTXX$) <> 1 Then GoTo 10
   '
   Screen.MousePointer = 11
   DESKRIPCION$ = RECEPEQ09.Label10
   Call CREATABLA_OFABRIC
   CONDI$ = "NUMREPA = " & NROREP&
   If CantRegistros("REPARA", CONDI$) < 1 Then
      Call CreaRegistro("REPARA", "NUMREPA", NROREP&, "NOMESS")
   End If
   '
20 RACLI$ = TRIM$(RECEPEQ09.Text3)
   REFER$ = TRIM$(Left$(DESKRIPCION$, 22)) & " - " & RACLI$
   '
   Screen.MousePointer = 11
   Call ACTUREGISTRO("REPARA", "COD_INTE", CONDI$, CI%, REGAF&, "NOMESS")
   Call ACTUREGISTRO("REPARA", "DESCRIPCION", CONDI$, DESKRIPCION$, REGAF&, "NOMESS")
   Call ACTUREGISTRO("REPARA", "NumeroSerie", CONDI$, NUMESER$, REGAF&, "NOMESS")
      ESTREP% = 0: If RECEPEQ09.Label7 = "SI" Then ESTREP% = 1
   Call ACTUREGISTRO("REPARA", "Status_Gar", CONDI$, ESTREP%, REGAF&, "NOMESS")
   '
   Call ACTUREGISTRO("REPARA", "REFEREN", CONDI$, REFER$, REGAF&, "NOMESS")
   Call ACTUREGISTRO("REPARA", "STATUS", CONDI$, "0", REGAF&, "NOMESS")
   Call ACTUREGISTRO("REPARA", "FechRecep", CONDI$, FETEXCOR1$(FECHANUM(Left$(RECEPEQ09.Label16, 8))), RAFE&, "NOMESS")
   Call ACTUREGISTRO("REPARA", "SecRepar", CONDI$, RECEPEQ09.Text11, RAFE&, "NOMESS")
   Call ACTUREGISTRO("REPARA", "Usu_Recep", CONDI$, RECEPEQ09.Label17, RAFE&, "NOMESS")
   Call ACTUREGISTRO("REPARA", "FechEnSol", CONDI$, FETEXCOR1$(FECHANUM(RECEPEQ09.Label14.Caption)), RAFE&, "NOMESS")
   Call ACTUREGISTRO("REPARA", "Nume_Cli", CONDI$, NCXX&, REGAF&, "NOMESS")
   Call ACTUREGISTRO("REPARA", "Raso_Cli", CONDI$, RECEPEQ09.Text3, REGAF&, "NOMESS")
   Call ACTUREGISTRO("REPARA", "Domi_Cli", CONDI$, RECEPEQ09.Text4, REGAF&, "NOMESS")
   Call ACTUREGISTRO("REPARA", "Loca_Cli", CONDI$, RECEPEQ09.Text6, REGAF&, "NOMESS")
'     Call ACTUREGISTRO("REPARA", "Pais_Cli", CONDI$, RECEPEQ09.Text3, REGAF&, "NOMESS")
   Call ACTUREGISTRO("REPARA", "Tele_Cli", CONDI$, RECEPEQ09.Text7, REGAF&, "NOMESS")
   Call ACTUREGISTRO("REPARA", "Mail_Cli", CONDI$, RECEPEQ09.Text10, REGAF&, "NOMESS")
   Call ACTUREGISTRO("REPARA", "Ctac_Cli", CONDI$, RECEPEQ09.Text9, REGAF&, "NOMESS")
   
'     Call ACTUREGISTRO("REPARA", "Accesorios", 12, 0)
   Call ACTUREGISTRO("REPARA", "Observa", CONDI$, RECACCE09.Text1.TEXT, REGAF&, "NOMESS")
   Call ACTUREGISTRO("REPARA", "Cuit_Cli", CONDI$, CUITCLI$(NCXX&), REGAF&, "NOMESS")
   Call ACTUREGISTRO("REPARA", "Piva_Cli", CONDI$, PIVACLI%(NCXX&), REGAF&, "NOMESS")
'      Call ACTUREGISTRO("REPARA", "Status_Gar", 3, 0)
'      Call ACTUREGISTRO("REPARA", "TipoFalla", 3, 0)
'      Call ACTUREGISTRO("REPARA", "DeTipFalla", 10, 128)
'      Call ACTUREGISTRO("REPARA", "Referen2", 10, 64)
   '
   Screen.MousePointer = 11
   '
   ' DESCRIPCION DE LOS FALLOS
   DELARFA$ = ""
   Call BORRAREGISTROS("TIPOFALL", CONDI$, RAFE&, "NOMESS")
   For KKI% = 1 To RECACCE09.List1.ListCount
      If RECACCE09.List1.Selected(KKI% - 1) = True Then
         If DELARFA$ <> "" Then DELARFA$ = DELARFA$ + vbCrLf
         DELARFA$ = DELARFA$ + TRIM$(RECACCE09.List1.List(KKI% - 1))
         '
28       RNDX& = XVALO(RANDSTRING$(-6))
         CONDI1$ = "NUMREPA = " & RNDX&
         If CantRegistros("TIPOFALL", CONDI1$, "NOMESS") > 0 Then GoTo 28
         Call CreaRegistro("TIPOFALL", "NUMREPA", RNDX&, "NOMESS")
         Call ACTUREGISTRO("TIPOFALL", "ORIGEN", CONDI1$, 1, RAFE&, "NOMESS")  ' 1 CLIENTE; 2 TECNICO
         Call ACTUREGISTRO("TIPOFALL", "CODIFALL", CONDI1$, XVALO(Left$(RECACCE09.List1.List(KKI% - 1), 4)), RAFE&, "NOMESS")
         Call ACTUREGISTRO("TIPOFALL", "DESCFALL", CONDI1$, TRIM$(Mid$(RECACCE09.List1.List(KKI% - 1), 6)), RAFE&, "NOMESS")
         Call ACTUREGISTRO("TIPOFALL", "NUMREPA", CONDI1$, NROREP&, RAFE&, "NOMESS")
      End If
   Next KKI%
   '
   Screen.MousePointer = 11
   '
   If TRIM$(RECACCE09.Text2) <> "" Then
       If DELARFA$ <> "" Then DELARFA$ = DELARFA$ + vbCrLf
       DELARFA$ = DELARFA$ + TRIM$(RECACCE09.Text2)
       '
38     RNDX& = XVALO(RANDSTRING$(-6))
       CONDI1$ = "NUMREPA = " & RNDX&
       If CantRegistros("TIPOFALL", CONDI1$, "NOMESS") > 0 Then GoTo 38
       Call CreaRegistro("TIPOFALL", "NUMREPA", RNDX&, "NOMESS")
       Call ACTUREGISTRO("TIPOFALL", "ORIGEN", CONDI1$, 1, RAFE&, "NOMESS")  ' 1 CLIENTE; 2 TECNICO
       Call ACTUREGISTRO("TIPOFALL", "CODIFALL", CONDI1$, 999, RAFE&, "NOMESS")
       Call ACTUREGISTRO("TIPOFALL", "DESCFALL", CONDI1$, RECACCE09.Text2, RAFE&, "NOMESS")
       Call ACTUREGISTRO("TIPOFALL", "NUMREPA", CONDI1$, NROREP&, RAFE&, "NOMESS")
   End If
   Call ACTUREGISTRO("REPARA", "Defec_Infor", CONDI$, DELARFA$, RAFE&, "NOMESS")
   '
   Screen.MousePointer = 11
   '
   ' ACCESORIOS
   ACCESOR$ = ""
   Call BORRAREGISTROS("ACCEREP", CONDI$, RAFE&, "NOMESS")
   For KKI% = 2 To RECACCE09.GRID.Rows
      CANTU& = XVALO(RECACCE09.GRID.TextMatrix(KKI% - 1, 1))
      If CANTU& > 0 Then
         If ACCESOR$ <> "" Then ACCESOR$ = ACCESOR$ + vbCrLf
         ACCESOR$ = ACCESOR$ + TRIM$(RECACCE09.GRID.TextMatrix(KKI% - 1, 0)) + " (" + TRIM$(Str$(CANTU&)) + ")"
         If TRIM$(RECACCE09.GRID.TextMatrix(KKI% - 1, 2)) <> "" Then
            ACCESOR$ = ACCESOR$ + " - " + TRIM$(RECACCE09.GRID.TextMatrix(KKI% - 1, 2))
         End If
         '
48       RNDX& = XVALO(RANDSTRING$(-6))
         CONDI1$ = "NUMREPA = " & RNDX&
         If CantRegistros("ACCEREP", CONDI1$, "NOMESS") > 0 Then GoTo 48
         Call CreaRegistro("ACCEREP", "NUMREPA", RNDX&, "NOMESS")
            DESCRIACC$ = RECACCE09.GRID.TextMatrix(KKI% - 1, 0)
         Call ACTUREGISTRO("ACCEREP", "DESCRIACC", CONDI1$, DESCRIACC$, RAFE&, "NOMESS")
         Call ACTUREGISTRO("ACCEREP", "CANTIDAD", CONDI1$, CANTU&, RAFE&, "NOMESS")
            OBSEX$ = RECACCE09.GRID.TextMatrix(KKI% - 1, 2)
         Call ACTUREGISTRO("ACCEREP", "OBSERVA", CONDI1$, OBSEX$, RAFE&, "NOMESS")
         Call ACTUREGISTRO("ACCEREP", "NUMREPA", CONDI1$, NROREP&, RAFE&, "NOMESS")
      End If
   Next KKI%
   Call ACTUREGISTRO("REPARA", "Accesorios", CONDI$, ACCESOR$, RAFE&, "NOMESS")
   '
   '
   Screen.MousePointer = 11
   '
   CUERPO$ = "Recepción del Equipo para Reparación" & vbCrLf
   If OPX% = 2 Then ' CORREGIR DATOS
      CUERPO$ = "Modificacion de Recepción de Equipo para Reparación" & vbCrLf
   End If
   CUERPO$ = CUERPO$ + vbCrLf + "N-Rep: " & NROREP& & " - Cliente: " & RACLI$ + vbCrLf
   '
   If DELARFA$ <> "" Then
      CUERPO$ = CUERPO$ + vbCrLf
      CUERPO$ = CUERPO$ + "Fallas Informadas por Cliente:" + vbCrLf
      CUERPO$ = CUERPO$ + DELARFA$ + vbCrLf
   End If
   '
   If ACCESOR$ <> "" Then
      CUERPO$ = CUERPO$ + vbCrLf
      CUERPO$ = CUERPO$ + "Accesorios Recibidos:" + vbCrLf
      CUERPO$ = CUERPO$ + ACCESOR$ + vbCrLf
   End If
   '
   If TRIM$(RECACCE09.Text1) <> "" Then
      CUERPO$ = CUERPO$ + vbCrLf
      CUERPO$ = CUERPO$ + "Observaciones:" + vbCrLf
      CUERPO$ = CUERPO$ + RECACCE09.Text1 + vbCrLf
   End If
   '
   Call ANOTANUSE(NUMESER$, CUERPO$, , "NOMESS")
   '
   Screen.MousePointer = 1
   If OPX% = 2 Then GoTo 15
   '
   Call COMUNI("Equipo Ingresado Para Presupuestación")
   '
   FREP = HOY(HOS$)
   '
   FILTX$ = TRIM$(Str$(NROREP&)) & ";" & TRIM$(Str$(FREP)) & ";" & TRIM$(Str$(NC%))
   Call STDFORM("NREP-MDT", FILTX$, "PRINT")
    '
End Sub

