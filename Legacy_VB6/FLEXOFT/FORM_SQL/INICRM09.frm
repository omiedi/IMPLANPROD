VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form INICRM09 
   BackColor       =   &H00C0C0C0&
   Caption         =   "CRM - FLEXOFT de Administracion Comercial"
   ClientHeight    =   8445
   ClientLeft      =   60
   ClientTop       =   315
   ClientWidth     =   11880
   Icon            =   "INICRM09.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8445
   ScaleWidth      =   11880
   Begin VB.Timer Timer5 
      Interval        =   10000
      Left            =   4800
      Top             =   0
   End
   Begin VB.Timer Timer4 
      Interval        =   10000
      Left            =   4410
      Top             =   0
   End
   Begin VB.Timer Timer2 
      Interval        =   6000
      Left            =   3570
      Top             =   0
   End
   Begin VB.PictureBox Picture7 
      BackColor       =   &H00000040&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   8600
      Left            =   11025
      ScaleHeight     =   8535
      ScaleWidth      =   1215
      TabIndex        =   0
      Top             =   0
      Width           =   1275
      Begin VB.CommandButton Command9 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   105
         Picture         =   "INICRM09.frx":1D2A
         Style           =   1  'Graphical
         TabIndex        =   35
         ToolTipText     =   "Diseñador de Formularios"
         Top             =   5670
         Width           =   600
      End
      Begin VB.CommandButton Command8 
         Caption         =   "Command8"
         Height          =   255
         Left            =   120
         TabIndex        =   34
         Top             =   1320
         Visible         =   0   'False
         Width           =   495
      End
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
         Left            =   105
         Picture         =   "INICRM09.frx":216C
         Style           =   1  'Graphical
         TabIndex        =   33
         ToolTipText     =   "Mails Pendientes de Envio"
         Top             =   4410
         Width           =   600
      End
      Begin VB.CommandButton Command29 
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
         Picture         =   "INICRM09.frx":25AE
         Style           =   1  'Graphical
         TabIndex        =   32
         ToolTipText     =   "Contactos Programados"
         Top             =   3780
         Width           =   600
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Clear"
         Height          =   255
         Left            =   105
         TabIndex        =   29
         Top             =   3480
         Visible         =   0   'False
         Width           =   600
      End
      Begin VB.PictureBox Picture4 
         BackColor       =   &H00000040&
         BorderStyle     =   0  'None
         Height          =   2115
         Left            =   0
         ScaleHeight     =   2115
         ScaleWidth      =   1275
         TabIndex        =   21
         Top             =   6195
         Width           =   1275
         Begin VB.PictureBox MARCOBARRA 
            BackColor       =   &H000000FF&
            Height          =   135
            Left            =   90
            ScaleHeight     =   75
            ScaleWidth      =   585
            TabIndex        =   25
            Top             =   945
            Width           =   650
            Begin VB.Frame BARRATRAZA 
               BackColor       =   &H00FF0000&
               BorderStyle     =   0  'None
               Height          =   500
               Left            =   0
               TabIndex        =   26
               Top             =   0
               Width           =   12000
            End
         End
         Begin VB.Image Image4 
            Height          =   390
            Left            =   -315
            Picture         =   "INICRM09.frx":2CF0
            Top             =   1185
            Width           =   1515
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            Caption         =   "00:00:00"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   35
            TabIndex        =   24
            Top             =   1680
            Width           =   750
         End
         Begin VB.Label Label4 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            Caption         =   "00/00/00"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   35
            TabIndex        =   23
            Top             =   525
            Width           =   750
         End
         Begin VB.Label Label5 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Miercoles"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0080FFFF&
            Height          =   255
            Left            =   0
            TabIndex        =   22
            Top             =   210
            Width           =   795
         End
      End
      Begin VB.CommandButton Command6 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   105
         Picture         =   "INICRM09.frx":4C12
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "Generador de Reportes - Consultas SQL"
         Top             =   5040
         Width           =   600
      End
      Begin VB.CommandButton Command16 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   105
         Picture         =   "INICRM09.frx":4F1C
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Cerrar Sesión - Cambiar Usuario / Entorno"
         Top             =   700
         Width           =   600
      End
      Begin VB.CommandButton Command4 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   105
         Picture         =   "INICRM09.frx":5066
         Style           =   1  'Graphical
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   2810
         Width           =   600
      End
      Begin VB.CommandButton Command3 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   105
         Picture         =   "INICRM09.frx":54A8
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Cambio de Clave Personal de Acceso"
         Top             =   2205
         Width           =   600
      End
      Begin VB.CommandButton Command2 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   105
         Picture         =   "INICRM09.frx":57B2
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Manuales de Uso y Operacion"
         Top             =   1600
         Width           =   600
      End
      Begin VB.CommandButton Command1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   105
         Picture         =   "INICRM09.frx":5ABC
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Terminar - Cerrar Sesion de Trabajo"
         Top             =   105
         Width           =   600
      End
   End
   Begin VB.PictureBox Picture10 
      BackColor       =   &H80000001&
      Height          =   8500
      Left            =   -10
      ScaleHeight     =   8445
      ScaleWidth      =   3090
      TabIndex        =   17
      Top             =   -50
      Width           =   3150
      Begin VB.PictureBox Picture19 
         Height          =   855
         Left            =   -240
         ScaleHeight     =   795
         ScaleWidth      =   3000
         TabIndex        =   27
         Top             =   0
         Visible         =   0   'False
         Width           =   3060
         Begin VB.Image Image11 
            Height          =   1590
            Left            =   0
            Top             =   0
            Width           =   2010
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00404040&
         Caption         =   "Frame1"
         Height          =   8415
         Left            =   3045
         TabIndex        =   19
         Top             =   -100
         Width           =   75
      End
      Begin ComctlLib.TreeView Treeview1 
         Height          =   2400
         Left            =   0
         TabIndex        =   18
         Top             =   2040
         Width           =   3000
         _ExtentX        =   5292
         _ExtentY        =   4233
         _Version        =   327682
         Indentation     =   441
         LabelEdit       =   1
         LineStyle       =   1
         Style           =   7
         ImageList       =   "imlSmallIcons"
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label6 
         BorderStyle     =   1  'Fixed Single
         Caption         =   " Menu Principal"
         Height          =   260
         Left            =   0
         TabIndex        =   28
         Top             =   840
         Width           =   2850
      End
      Begin ComctlLib.ImageList imlIcons 
         Left            =   600
         Top             =   4185
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   32
         ImageHeight     =   32
         MaskColor       =   12632256
         _Version        =   327682
         BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
            NumListImages   =   1
            BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
               Picture         =   "INICRM09.frx":5C06
               Key             =   "book"
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Frame1"
      ForeColor       =   &H00404040&
      Height          =   75
      Left            =   0
      TabIndex        =   31
      Top             =   7140
      Width           =   15000
   End
   Begin VB.PictureBox Picture3 
      BackColor       =   &H0080C0FF&
      Height          =   2000
      Left            =   3120
      ScaleHeight     =   1935
      ScaleWidth      =   15945
      TabIndex        =   30
      Top             =   5880
      Width           =   16000
      Begin VB.Image Image3 
         Height          =   480
         Index           =   0
         Left            =   210
         Picture         =   "INICRM09.frx":5F20
         Top             =   210
         Width           =   480
      End
   End
   Begin VB.Timer Timer3 
      Interval        =   1000
      Left            =   3990
      Top             =   0
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   4620
      Top             =   1785
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox Picture8 
      BackColor       =   &H00C0C0C0&
      Height          =   1300
      Left            =   8925
      Picture         =   "INICRM09.frx":6362
      ScaleHeight     =   1245
      ScaleWidth      =   1770
      TabIndex        =   1
      Top             =   105
      Visible         =   0   'False
      Width           =   1830
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
         Left            =   165
         TabIndex        =   7
         Top             =   420
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
         TabIndex        =   6
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
         TabIndex        =   5
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
         TabIndex        =   4
         Top             =   60
         Width           =   225
      End
   End
   Begin VB.PictureBox Picture2 
      AutoRedraw      =   -1  'True
      Height          =   1380
      Left            =   3990
      ScaleHeight     =   1320
      ScaleWidth      =   2370
      TabIndex        =   20
      ToolTipText     =   "Click con Boton Derecho para Cambiar la Imagen"
      Top             =   3255
      Visible         =   0   'False
      Width           =   2430
   End
   Begin VB.CommandButton Command17 
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
      Picture         =   "INICRM09.frx":DAF4
      Style           =   1  'Graphical
      TabIndex        =   16
      ToolTipText     =   "Depósito Caotico SAPORITI"
      Top             =   6195
      Visible         =   0   'False
      Width           =   600
   End
   Begin VB.FileListBox File1 
      Height          =   1065
      Left            =   2520
      Pattern         =   "*.EXK"
      TabIndex        =   14
      Top             =   735
      Visible         =   0   'False
      Width           =   1590
   End
   Begin VB.PictureBox Picture9 
      Height          =   2010
      Left            =   420
      ScaleHeight     =   1950
      ScaleWidth      =   2055
      TabIndex        =   12
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
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   3150
      Top             =   0
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
      Left            =   1785
      Picture         =   "INICRM09.frx":DC3E
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Control y Movimientos de Stocks"
      Top             =   5040
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
      Left            =   630
      Picture         =   "INICRM09.frx":DF48
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Actualización y Consulta de Archivos Maestros"
      Top             =   3360
      Width           =   600
   End
   Begin VB.Image Image2 
      Height          =   5865
      Left            =   6120
      Picture         =   "INICRM09.frx":E252
      Top             =   105
      Visible         =   0   'False
      Width           =   7485
   End
   Begin ComctlLib.ImageList imlSmallIcons 
      Left            =   5000
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   13
      ImageHeight     =   13
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   6
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "INICRM09.frx":16229
            Key             =   "closed"
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "INICRM09.frx":1674B
            Key             =   "cylinder"
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "INICRM09.frx":16C6D
            Key             =   "leaf"
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "INICRM09.frx":1718F
            Key             =   "open"
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "INICRM09.frx":176B1
            Key             =   "smlBook"
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "INICRM09.frx":17D13
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "INICRM09"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'agregar referencia microsoft outlook 11.0 object library

Dim olApp As Outlook.Application
Dim Inbox As Outlook.MAPIFolder
Dim oOMail As Outlook.MailItem
Dim InboxItems As Outlook.Items
Dim Mailobject As Object        ' Atenuado como objeto Mailobject
Dim Subject As String           ' Sin perjuicio de As String Dim
Dim Tomails As String

Dim ARRABOR%
Dim ESTAENVIANDO%
Dim POROUTLOOK%
'
Sub INGREID()
     OPCOTEL.Show 1
End Sub
'
Private Sub Command1_Click()
   '
   Call GRAVALCON(USNBR%, "POSIFORPRIN", "TOP", FORMATNUM$(Top, "F10.0"))
   Call COMPACTABD
   '
   Call SAVEFILE(LUCOM$ + "RCONTROL.DRV", "CLOSETAB")
   Call CIERRARCH("*.*")
   Call BAJALDISCO
   If PROCACTI% < 1 Then
      Call GENBACKUP
      Call CANCELUCOM
      Call FINAL("")
   End If
   '
End Sub

Private Sub Command11_Click()
   Command11.Enabled = False
   If EXISTE%("ARCHIG07.EXE") > 0 Then
     Call CONECRUN("ARCHIG07", "Archivos Maestros")
   Else
     Call COMUNI("Falta módulo ARCHIG07")
   End If
   Command11.Enabled = True
End Sub

Private Sub Command12_Click()
   Command12.Enabled = False
   If EXISTE%("CONSTO07.EXE") > 0 Then
     Call CONECRUN("CONSTO07", "Control de Stocks")
   Else
     Call COMUNI("Falta módulo CONSTO07")
   End If
   Command12.Enabled = True
End Sub

Private Sub Command16_Click()
   COMMAND16.Enabled = False
   Call CIERRARCH("*.*")
   Call BAJALDISCO
   '
   If PROCACTI% < 1 Then
      Call CANCELUCOM
      ' verifica si es multiempresa
      Call RECONN_EMPRE(OKXX%)
      If OKXX% < 1 Then GoTo 80
      TCONFI$ = LOADFILE$(CurDir + "\" + "FLEXCFG.DAT")
      If TCONFI$ = "" Then Exit Sub
      '
      LUDAT$ = TRIM$(Mid$(TCONFI$, 51, 45))
      If Right$(LUDAT$, 1) <> "\" Then
         LUDAT$ = LUDAT$ + "\"
         Mid$(TCONFI$, 51, 45) = Left$(LUDAT$ + Space$(45), 45)
         Call SAVEFILE("FLEXCFG.DAT", TCONFI$)
      End If
      If Mid$(LUDAT$, 2, 1) <> ":" Then
         ULO$ = Left$(TRIM$(LOADFILE$(LUCOM$ + "ULOGDAT.DRV")), 2)
         If ULO$ = "" Then ULO$ = Left$(CurDir$, 2)
         LUDAT$ = ULO$ + LUDAT$
      End If
      '
      RSTRI$ = LOADFILE$(LUCOM$ + "FLEXRAM.DRV") + String$(512, 0)
      EMPREAC$ = TRIM$(Mid$(RSTRI$, 193, 30))
         Call GRARAM(99, AJUSTI$(LUDAT$, 46))
         Call GRARAM(193, AJUSTI$(EMPREAC$, 30))
         Call GRARAM(191, Chr$(CodiEmp%))
         Call BAJRAM
      EMPREAC$ = EMPREAC$ + " (" + TRIM$(Str$(CodiEmp%)) + ") "
      EMPREAC$ = EMPREAC$ + VERNUM$ + " - " + USERNAME$
      Screen.MousePointer = 1
      If OKEY% < 1 Then Call FINAL("")
      EPAC$ = TRIM$(EMPREAC$)
      If EPAC$ <> "" Then
         Caption = "Sistema Preconfigurado FLEXOFT" + " - " + EPAC$
         Me.Refresh
      End If
      Call GRATITFOR(Caption)
      Call PONELOGO
      Call UPDATES
GoTo 99

80    ' conecta con FLOGON.EXE
      On Error Resume Next
      RETOR$ = TRIM$(UCase$(App.EXEName))
      Shell "FLOGON.EXE " + RETOR$, vbNormalFocus
      If Err Then
        Call MERRFINAL("Falta Utilitario 'FLOGON.EXE'.")
      End If
      Call FINAL("")
   End If
   '
99 COMMAND16.Enabled = True
End Sub

Sub UPDATES()
   '
   Screen.MousePointer = 11
   Label2 = 171: DoEvents
'   Call LEEDATSYS
   '
   YAMENSA% = 0
   INSTALAC% = 0
   '
   DIREPRO$ = App.Path
   If EXISTE%("F:\FLEXOFT\NOVEDAD\.") > 0 Then
       DIRENEW$ = "F:\FLEXOFT\NOVEDAD"
     ElseIf EXISTE%("C:\FLEXOFT\NOVEDAD\.") > 0 Then
       DIRENEW$ = "C:\FLEXOFT\NOVEDAD"
     Else
       DIRENEW$ = PARENTFOLDER$(DIREPRO$) + "\NOVEDAD"
   End If
   DIRNU$ = DIRENEW$
10 On Error Resume Next
   File1.Path = DIRNU$
   If Err < 1 Then
     On Error GoTo 0
     File1.Pattern = "*.*"
     '
     For KU& = 1 To File1.ListCount
       ANAME$ = TRIM$(File1.List(KU& - 1))
       PPXX% = InStr(ANAME$, ".")
       If PPXX% < 1 Then PPXX% = 1 + Len(ANAME$)
       exten$ = TRIM$(UCase$(Mid$(ANAME$, PPXX% + 1)))
       '
       If Len(exten$) = 3 Then
         '
         If exten$ = "RFS" Or exten$ = "FRM" Then
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
           GoTo 29         ' SALE DE INSTALACION DE DATOS
         End If
         '
         RUTACOM$ = App.Path + "\"
         If EXISTE%(RUTACOM$ + ANAME$) > 0 Or exten$ = "PRF" Then
           On Error Resume Next
           FECHANUE# = FileDateTime(DIRENEW$ + "\" + ANAME$)
           If Err Then
              On Error GoTo 0
              GoTo 29
           End If
           On Error GoTo 0
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
               Name DIREVIE$ + ANAMCO$ + TRIM$(Str$(KI%)) As DIREVIE$ + ANAMCO$ + TRIM$(Str$(KI% + 1))
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
   End If
   '
   ' esto es para explorar tambien el \FLEXOFT\NOVEDAD de la carpeta de trabajo
   On Error GoTo 0
   If DIRNU$ = "F:\FLEXOFT\NOVEDAD" Then
     DIRNU$ = PARENTFOLDER$(DIREPRO$) + "\NOVEDAD"
     If DIRNU$ <> "F:\FLEXOFT\NOVEDAD" Then
       GoTo 10
     End If
   End If
   '
   If INSTALAC% > 0 Then MsgBox "Actualizaciones Instaladas"
   Screen.MousePointer = 1
   '
End Sub

Sub RECONN_EMPRE(OKXX%)
   '
   OKXX% = 0
   If EXISTE%("FLEXMUL.DRV") = 1 Then
      SELUDAT2.List2.Clear
      SELUDAT2.List2.Visible = True
      '
      NX% = FILEOPEN("FLEXMUL.DRV", 0, 79)
      Dim FMX As String * 79
      '
      For i% = 1 To LOF(NX%) / 79
         Get #NX%, i%, FMX$
         SUCAR% = 0
         For J% = 1 To 77
            SUCAR% = SUCAR% + Asc(Mid$(FMX$, J%, 1))
         Next J%
         If CVI(Mid$(FMX$, 78, 2)) = SUCAR% Then
            RASO$ = DESENCRI$(Left$(FMX$, 30))
            PACCE$ = DESENCRI$(Mid$(FMX$, 32, 46))
            If EXISTE%(TRIM$(PACCE$) + "\.") > 0 Then
               Z$ = Space$(96)
               Call REPLA(Z$, RASO$, 1, 30)
               Call REPLA(Z$, PACCE$, 51, 46)
               SELUDAT2.List2.AddItem Z$
            End If
         End If
      Next i%
      Close #NX%
      '
      If SELUDAT2.List2.ListCount > 0 Then
         OKXX% = 1
         SELUDAT2.List2.ListIndex = 0
         Screen.MousePointer = 1
         If SELUDAT2.List2.ListCount = 1 Then
            Call SELUDAT2.List2_DblClick
            GoTo 25
         End If
         SELUDAT2.Show 1
         Screen.MousePointer = 11
         '
         If SELUDAT2.List2.ListCount > 1 Then
           If TRIM$(EMPREAC$) <> TRIM$(Left$(SELUDAT2.List2.List(0), 30)) Then
             If DERACCE%(EMPREAC$, "Acceso a " + EMPREAC$) < 2 Then
               Call FINAL("")                 ' NO ACCEDE
             End If
           End If
         End If
         '
      End If
   End If
   '
25
End Sub

Private Sub Command17_Click()
   Command17.Visible = False
   ACONEC$ = "TRANSA07"
   Call CONECRUN(ACONEC$, "Deposito Caotico SAPORITI")
   Command17.Visible = True
End Sub

Private Sub Command2_Click()
   Command2.Enabled = False
   Call MENSERR(24, "Ayuda en Linea no Disponible\en Esta Localización.")
   Command2.Enabled = True
End Sub

Private Sub Command29_Click()
   Command29.Enabled = False
   OPX% = COMALTER%("Opciones de Trabajo:\\Actividades Programadas\Responder Mails Recibidos")
   If OPX% = 1 Then
        Call CONACT09.Command1_Click
        CONACT09.Show 1
      ElseIf OPX% = 2 Then
        USNU% = USNBR% Mod 100
        CONDI$ = "USUA_DESTI = " & USNU%
           If USNU% = 1 Then CONDI$ = "( " & CONDI$ & " OR USUA_DESTI IS NULL )"
        CONDI$ = CONDI$ & " AND (CONF_RECEP < '01/01/01' OR NUMAIRESPU = 0 OR NUMAIRESPU IS NULL) "
        CONDI$ = CONDI$ & " ORDER BY FHOR_RECEP ASC "
        Call COPYSTRU("IBOLPEN", "MAIPENRE")
10      Call CARGALISEL("MAIPENRE", "BANDENTRA", "NUME_MAIL/F5; FHOR_RECEP/D17;ORIG_MAIL/A48;", CONDI$, "NOMESS")
        Call SELECHO("MAIPENRE/Indice de Mails Pendientes de Envio")
        NROMAI& = XVALO(VALACT1$("MAIPENRE"))
        If NROMAI& > 0 Then
           SHINMAIL09.Label13 = TRIM$(Str$(NROMAI&))
           SHINMAIL09.Show 1
           GoTo 10
        End If
   End If
   Command29.Enabled = True
End Sub

Private Sub Command3_Click()
   Command3.Enabled = False
   Call CONECRUN("FSETUP04/CLAVUSER", "Cambio de Clave Personal de Acceso")
99 Command3.Enabled = True
End Sub

Private Sub Command4_Click()
   If USNBR% <> 101 Then GoTo 20
   Command4.Enabled = False
   OPX% = COMALTER%("Opciones de Configuración:\\Conf.General\Particular\Clasificaciones")
   If OPX% = 1 Then
        Call CONECRUN("FSETUP04", "")
      ElseIf OPX% = 2 Then
20      STPCRM09.Show 1
      ElseIf OPX% = 3 Then
        Call FISECOM09.ACTAGRUP
   End If
   Command4.Enabled = True
End Sub

Private Sub Command5_Click()
     NOMECOR$ = "SQLCTRL"
     On Error Resume Next
     BaDaConsul.Close
     On Error GoTo 0
     On Error Resume Next
     Kill App.Path + "\" + NOMECOR$ + "1.MDB"
     On Error GoTo 0
60   On Error Resume Next
     CompactDatabase App.Path + "\" + NOMECOR$ + ".MDB", App.Path + "\" + NOMECOR$ + "1.MDB"
     If Err Then GoTo 75
     FileCopy App.Path + "\" + NOMECOR$ + ".MDB", App.Path + "\" + NOMECOR$ + ".OLD"
     If Err Then GoTo 75
     FileCopy App.Path + "\" + NOMECOR$ + "1.MDB", App.Path + "\" + NOMECOR$ + ".MDB"
     If Err Then
75     TTXX$ = "Error (" & Err.Number & ") al Compactar " + NOMECOR$
       If Err.Number = 3343 Then
          On Error GoTo 0
          TTXX$ = TTXX$ + "\  \Esta Base de Datos está Dañada."
          TTXX$ = TTXX$ + "\¿ Desea Que FLEXOFT Intente Repararla ?"
          TTXX$ = TTXX$ + "\\Si, Reparar\No, Cancelar"
          If COMALTER%(TTXX$) = 1 Then
            On Error Resume Next
            DBEngine.RepairDatabase LUACC$ + NOMECOR$ + ".MDB"
            GoTo 60
          End If
          GoTo 80
       End If
       MsgBox TTXX$
     End If
80   On Error GoTo 0
     On Error Resume Next
     Kill App.Path + "\" + NOMECOR$ + "1.MDB"
     On Error GoTo 0
     MsgBox "FIN"
End Sub

Private Sub Command6_Click()
   Command6.Enabled = False
   Call CONECRUN("GENREP07", "Generador de Reportes SQL")
   Command6.Enabled = True
End Sub
 
Private Sub Command7_Click()
   Command7.Enabled = False
   If DERACCE%("REMAIPEN", "Ver y Remitir Mails Pendientes de Envío") >= 2 Then
      FECHAC = FETEXCOR1$(HOY(HOS$)) + " 00:00:01"
      CONDI$ = "(CANT_ENVI < 1 OR CANT_ENVI IS NULL)"
      CONDI$ = CONDI$ + " OR (FREC_REPI > 0 AND FEFIN_REPI >= '" & FECHAC & "')"
      CONDI$ = CONDI$ + " AND FHOR_SOLENVI > '01/01/01' "
      CONDI$ = CONDI$ + " order by FHOR_SOLENVI ASC"
      Call COPYSTRU("IBOLPEN", "MAIPENVI")
10    Call CARGALISEL("MAIPENVI", "BANDENVI", "NUME_MAIL/F5; FHOR_SOLENVI/D17;PARA_MAIL/A48;", CONDI$, "NOMESS")
      Call SELECHO("MAIPENVI/Indice de Mails Pendientes de Envio")
      NROMAI& = XVALO(VALACT1$("MAIPENVI"))
      If NROMAI& > 0 Then
         Call MAISEND(NROMAI&, "SHOW")
         GoTo 10
      End If
   End If
   Command7.Enabled = True
End Sub

Private Sub Command8_Click()
    '
    Dim BSEGSEL As ADODB.Recordset
    Set BSEGSEL = New ADODB.Recordset
    Set BSEGSEL = FLEXCONN.Execute("SELECT CODI_CLI, MAIL_CLI,SECTOR1,SECTOR2,SECTOR3 FROM BSEGCOM")
    With BSEGSEL
       Do While Not .EOF
          COCLI$ = TRIM$(SAFETEXT$(!CODI_CLI))
          MACLI$ = TRIM$(SAFETEXT$(!Mail_Cli))
          If MACLI$ <> "" Then
            If InStr(MACLI$, "@") < 2 Then
              CONDI$ = "CODI_CLI = '" & COCLI$ & "'"
              PAWEB$ = TRIM$(VALOBADA("BSEGCOM", "PAGWEB", CONDI$))
              PPXX% = InStr(LCase$(PAWEB$), "www.")
              If PPXX% > 0 Then
                 DOMIN$ = Mid$(PAWEB$, PPXX% + 4)
                 If DOMIN$ <> "" Then
                    PPYY% = InStr(LCase$(MACLI$), LCase$(DOMIN$))
                    If PPYY% > 0 Then
                       MACLU$ = TRIM$(Left$(MACLI$, PPYY% - 1) + "@" + LCase$(DOMIN$))
                       Call ACTUREGISTRO("BSEGCOM", "MAIL_CLI", CONDI$, MACLU$, RAFE&)
                    End If
                 End If
              End If
            End If
          End If
          '
          MACLI$ = TRIM$(SAFETEXT$(!SECTOR1))
          If MACLI$ <> "" Then
            If InStr(MACLI$, "@") < 2 Then
              CONDI$ = "CODI_CLI = '" & COCLI$ & "'"
              PAWEB$ = TRIM$(VALOBADA("BSEGCOM", "PAGWEB", CONDI$))
              PPXX% = InStr(LCase$(PAWEB$), "www.")
              If PPXX% > 0 Then
                 DOMIN$ = Mid$(PAWEB$, PPXX% + 4)
                 If DOMIN$ <> "" Then
                    PPYY% = InStr(LCase$(MACLI$), LCase$(DOMIN$))
                    If PPYY% > 0 Then
                       MACLU$ = TRIM$(Left$(MACLI$, PPYY% - 1) + "@" + LCase$(DOMIN$))
                       Call ACTUREGISTRO("BSEGCOM", "SECTOR1", CONDI$, MACLU$, RAFE&)
                    End If
                 End If
              End If
            End If
          End If
          '
          MACLI$ = TRIM$(SAFETEXT$(!SECTOR2))
          If MACLI$ <> "" Then
            If InStr(MACLI$, "@") < 2 Then
              CONDI$ = "CODI_CLI = '" & COCLI$ & "'"
              PAWEB$ = TRIM$(VALOBADA("BSEGCOM", "PAGWEB", CONDI$))
              PPXX% = InStr(LCase$(PAWEB$), "www.")
              If PPXX% > 0 Then
                 DOMIN$ = Mid$(PAWEB$, PPXX% + 4)
                 If DOMIN$ <> "" Then
                    PPYY% = InStr(LCase$(MACLI$), LCase$(DOMIN$))
                    If PPYY% > 0 Then
                       MACLU$ = TRIM$(Left$(MACLI$, PPYY% - 1) + "@" + LCase$(DOMIN$))
                       Call ACTUREGISTRO("BSEGCOM", "SECTOR2", CONDI$, MACLU$, RAFE&)
                    End If
                 End If
              End If
            End If
          End If
          '
          MACLI$ = TRIM$(SAFETEXT$(!SECTOR3))
          If MACLI$ <> "" Then
            If InStr(MACLI$, "@") < 2 Then
              CONDI$ = "CODI_CLI = '" & COCLI$ & "'"
              PAWEB$ = TRIM$(VALOBADA("BSEGCOM", "PAGWEB", CONDI$))
              PPXX% = InStr(LCase$(PAWEB$), "www.")
              If PPXX% > 0 Then
                 DOMIN$ = Mid$(PAWEB$, PPXX% + 4)
                 If DOMIN$ <> "" Then
                    PPYY% = InStr(LCase$(MACLI$), LCase$(DOMIN$))
                    If PPYY% > 0 Then
                       MACLU$ = TRIM$(Left$(MACLI$, PPYY% - 1) + "@" + LCase$(DOMIN$))
                       Call ACTUREGISTRO("BSEGCOM", "SECTOR3", CONDI$, MACLU$, RAFE&)
                    End If
                 End If
              End If
            End If
          End If
          '
       .MoveNext
       Loop
    End With
    BSEGSEL.Close
    Set BSEGSEL = Nothing
    '
    Call PoneEnCeroNull("CONTACSCOM", "ORDI_CTAC", RAFE&)
    Dim CTCSEL As ADODB.Recordset
    Set CTCSEL = New ADODB.Recordset
    Set CTCSEL = FLEXCONN.Execute("SELECT CODICLI, MAIL_CTAC ,ORDI_CTAC FROM CONTACSCOM")
    With CTCSEL
       Do While Not .EOF
          COCLI$ = TRIM$(SAFETEXT$(!CODICLI))
          MACLI$ = TRIM$(SAFETEXT$(!MAIL_CTAC))
          ORCON% = XVALO(!ORDI_CTAC)
          If MACLI$ <> "" Then
            If InStr(MACLI$, "@") < 2 Then
              CONDI$ = "CODI_CLI = '" & COCLI$ & "'"
              PAWEB$ = TRIM$(VALOBADA("BSEGCOM", "PAGWEB", CONDI$))
              PPXX% = InStr(LCase$(PAWEB$), "www.")
              If PPXX% > 0 Then
                 DOMIN$ = Mid$(PAWEB$, PPXX% + 4)
                 If DOMIN$ <> "" Then
                    PPYY% = InStr(LCase$(MACLI$), LCase$(DOMIN$))
                    If PPYY% > 0 Then
                       MACLU$ = TRIM$(Left$(MACLI$, PPYY% - 1) + "@" + LCase$(DOMIN$))
                       CONDI2$ = "CODICLI = '" & COCLI$ & "' AND ORDI_CTAC = " & ORCON%
                       Call ACTUREGISTRO("CONTACSCOM", "MAIL_CTAC", CONDI2$, MACLU$, RAFE&)
                    End If
                 End If
              End If
            End If
          End If
          '
       .MoveNext
       Loop
    End With
    CTCSEL.Close
    Set CTCSEL = Nothing
    '
    Call COMUNI("FIN DE PROCESO")
    '
End Sub

Private Sub Command9_Click()
   Command9.Enabled = False
   Call CONECRUN("DSGFRM09", "Diseñador de Formularios")
   Command9.Enabled = True
End Sub

Private Sub Form_Load()
   '
   Screen.MousePointer = 11
   Visible = False
   '
   Call VERJOBID(OKEY%)
   If OKEY% < 1 Then Call FINAL("")
   '
POROUTLOOK% = 1
   '
   Screen.MousePointer = 11
   EPAC$ = TRIM$(EMPREAC$)
   If EPAC$ <> "" Then
     Caption = Caption + "  -  " + EPAC$
   End If
   Call GRATITFOR(Caption)
   '
   Picture4.BackColor = Picture7.BackColor
   '
   If DINFE& < 1 Or DINFE& > Height Then DINFE& = 1500
   Frame3.Top = Height - DINFE&
   '
   On Error Resume Next
   Picture3.Top = Frame3.Top + Frame3.Height
   Picture3.Height = Height - Picture3.Top
   On Error GoTo 0
   Screen.MousePointer = 11
   Call ACOMODATODO
   '
   Call ABREPEDCLI
   Call CREACAMPO("VENTAS", "OPORTUNI", "CodOport", 10, 12, 1)
   Call CREACAMPO("VENTAS", "OPORTUNI", "Codi_Cli", 10, 12, 1)
   Call CREACAMPO("VENTAS", "OPORTUNI", "Raso_Cli", 10, 64)
   Call CREACAMPO("VENTAS", "OPORTUNI", "RefeOport", 10, 256)
   Call CREACAMPO("VENTAS", "OPORTUNI", "Descri_Oport", 12, 0)
   Call CREACAMPO("VENTAS", "OPORTUNI", "His_Oportu", 12, 0)
   Call CREACAMPO("VENTAS", "OPORTUNI", "RepComer", 3, 0)
   Call CREACAMPO("VENTAS", "OPORTUNI", "Aprobador", 3, 0)
   Call CREACAMPO("VENTAS", "OPORTUNI", "TipoOportu", 3, 0)
   Call CREACAMPO("VENTAS", "OPORTUNI", "FunilVta", 3, 0)
   Call CREACAMPO("VENTAS", "OPORTUNI", "FechAlta", 8, 0)
   Call CREACAMPO("VENTAS", "OPORTUNI", "ValorEstimado", 7, 0)
   Call CREACAMPO("VENTAS", "OPORTUNI", "PersCtacto", 10, 64)
   Call CREACAMPO("VENTAS", "OPORTUNI", "PersDecisor", 10, 64)
   '
   Call CREACAMPO("", "BSEGCOM", "TCELU1", 10, 64)
   Call CREACAMPO("", "BSEGCOM", "TCELU2", 10, 64)
   Call CREACAMPO("", "BSEGCOM", "TCELU3", 10, 64)
   Call CREACAMPO("", "CONTACSCOM", "CELU_CTAC", 10, 64)
   '
   ' PONE LOS ACCESOS DIRECTOS
   Screen.MousePointer = 11
   Call CARGACCDIR
   '
   Treeview1.Style = tvwTreelinesPlusMinusPictureText
   On Error Resume Next
   imlSmallIcons.ListImages.Add , "nueva", LoadPicture("C:\WINDOWS\ICONS\BOOKS05.ICO")
   On Error GoTo 0
   '
   Screen.MousePointer = 11
   Call CARGAMENU
   '
   Visible = True
   Screen.MousePointer = 1
   '
End Sub

Sub AGREGANODO(CODPA$, CODIJ$, DESIJ$, Optional MODOP$)
    '
    On Error Resume Next
    If CODIJ$ = "ROOT" Then
        Set nodX = Treeview1.Nodes.Add(, , "ROOT", DESIJ$, "closed", "open")
      Else
        Set nodX = Treeview1.Nodes.Add(CODPA$, tvwChild, CODIJ$, DESIJ$, "closed", "open")
    End If
    If UCase$(MODOP$) = "OPEN" Or CODPA$ = "ROOT" Or CODIJ$ = "ROOT" Then
       nodX.EnsureVisible       ' Muestra todos los nodos.
    End If
    On Error GoTo 0
End Sub
'
Sub CARGAMENU()
    '
    Treeview1.Nodes.Clear
    '
    Call AGREGANODO("", "ROOT", "Sistema FLEXOFT de Administracion Comercial")
        Call AGREGANODO("ROOT", "ARCHIVOS", "Archivos Maestros")
            Call AGREGANODO("ARCHIVOS", "AMAPROS", "Maestro de Contactos y Prospectos", "OPEN")
               Call AGREGANODO("AMAPROS", "CONPROGRA", "Contactos Programados")
               Call AGREGANODO("AMAPROS", "BUPORTEX", "Busca por Textos")
               Call AGREGANODO("AMAPROS", "PROSNUE", "Nuevo Contacto o Prospecto")
               Call AGREGANODO("AMAPROS", "IMPORDAT", "Importacion de Datos")
               '
            Call AGREGANODO("ARCHIVOS", "OPORTUN", "Oportunidades de Negocios")
               Call AGREGANODO("OPORTUN", "OPORTUVIE", "Abrir Oportunidad Existente")
               Call AGREGANODO("OPORTUN", "OPORTUNUE", "Nueva Oportunidad")
               '
        Call AGREGANODO("ROOT", "PROGACTI", "Planeamiento", "OPEN")
            Call AGREGANODO("PROGACTI", "NUEVACT", "Nueva Actividad Programada", "OPEN")
            Call AGREGANODO("PROGACTI", "CONACTI", "Control de Actividades Programadas")
            Call AGREGANODO("PROGACTI", "MENSAIN", "Mensajes Internos")
            '
        Call AGREGANODO("ROOT", "CORREOEL", "Correo Electronico")
            Call AGREGANODO("CORREOEL", "SENDMAIL", "Redactar y Enviar Correo", "OPEN")
            Call AGREGANODO("CORREOEL", "ANSWERMA", "Responder Mails Pendientes")
            Call AGREGANODO("CORREOEL", "BANDENVI", "Enviados por Rango de Fechas")
            Call AGREGANODO("CORREOEL", "MAINOPRO", "Correo Entrante Pendiente de Proceso")
            Call AGREGANODO("CORREOEL", "ARMAICIRC", "Newsletter - Mail Circular")
            '
        Call AGREGANODO("ROOT", "INFORMES", "Informes de Novedades")
            '
        Call AGREGANODO("ROOT", "LISTADOS", "Reportes, Consultas y Listados")
            '
        Call AGREGANODO("ROOT", "CONFIGUR", "Herramientas de Configuracion")
            Call AGREGANODO("CONFIGUR", "ARCONFIG", "Tablas de Configuracion")
               Call AGREGANODO("ARCONFIG", "TAREPCOM", "Representantes Comerciales")
               Call AGREGANODO("ARCONFIG", "DIRECOM", "Aprobadores / Directores Comerciales")
               Call AGREGANODO("ARCONFIG", "FUNILVE", "Tabla Funil de Ventas")
               Call AGREGANODO("ARCONFIG", "TATIPOP", "Tabla de Tipos de Oportunidades")
               Call AGREGANODO("ARCONFIG", "INGRESOID", "Carga de Identificador de Usuario")
            Call AGREGANODO("CONFIGUR", "CONFIGEN", "Configuracion General")
               Call AGREGANODO("CONFIGEN", "FSETUP04", "Configuracion de Funcionamiento")
               Call AGREGANODO("CONFIGEN", "CLAVEACC", "Cambio Clave Personal de Acceso")
               Call AGREGANODO("CONFIGEN", "COMPACBD", "Mantenimiento Base de Control")
               Call AGREGANODO("CONFIGEN", "GENREP07", "Diseñador de Reportes")
               Call AGREGANODO("CONFIGEN", "DSGFRM09", "Diseñador de Formularios")
            Call AGREGANODO("CONFIGUR", "CORRENTRA", "Procesamiento de Correo Entrante")
               Call AGREGANODO("CORRENTRA", "MAINOPRO", "Correo Entrante Pendiente de Proceso")
               Call AGREGANODO("CORRENTRA", "CRITBORR", "Criterios de Rechazo y Borrado")
                  Call AGREGANODO("CRITBORR", "ADIRCRIT", "Actualizacion Directa de Tabla")
                  Call AGREGANODO("CRITBORR", "POREMITE", "Actualizacion por Remitente")
                  Call AGREGANODO("CRITBORR", "PORDOMIN", "Actualizacion por Dominio del Remitente")
                  Call AGREGANODO("CRITBORR", "PORASUNT", "Actualizacion por Asunto")
                  '
    Treeview1.Nodes.Item(1).Selected = True
    Call ACOMODATODO
    '
End Sub

Sub CARGACCDIR()
   CANADIR = XVALO(VALCONTROL(USNBR%, "ADIRECT", "CANADIR"))
   If CANADIR > 255 Then CANADIR = 255
   '
   For i& = 1 To CANADIR
      If i& > Image3.Count Then
         Load Image3(i& - 1)
         Image3(i& - 1).Top = Image3(i& - 2).Top
         Image3(i& - 1).Left = Image3(i& - 2).Left + Image3(i& - 2).Width + 100
         Image3(i& - 1).Visible = True
      End If
      '
      ADIR$ = "ACCDIR-" + TRIM$(Str$(i& - 1))
      ARCHICON$ = VALCONTROL(USNBR%, ADIR$, "ARCHICON")
      TOOTIPTE$ = VALCONTROL(USNBR%, ADIR$, "TOOTIPTE")
      On Error Resume Next
      Image3(i& - 1).Picture = LoadPicture(ARCHICON$)
      If Err Then
         On Error GoTo 0
         Call MENSERR(24, "Imagen no Válida")
      End If
      On Error GoTo 0
      Image3(i& - 1).Visible = True
      Image3(i& - 1).ToolTipText = TOOTIPTE$
      Image3(i& - 1).Refresh
   Next i&
End Sub

Sub PONELOGO()
   '
   RUTARGRA$ = TRIM$(Control$("", "RUTARGRA"))
   If RUTARGRA$ = "" Then RUTARGRA$ = TRIM$(LUDAT$)
   If Right$(RUTARGRA$, 1) <> "\" Then
      RUTARGRA$ = RUTARGRA$ + "\"
   End If
   '
   ISOLOGO$ = TRIM$(Control$("", "ISOLOGO"))
   If ISOLOGO$ <> "" Then
      ARCHIMAG$ = RUTARGRA$ + ISOLOGO$
      If EXISTE%(ARCHIMAG$) = 1 Then
        Image11.Picture = LoadPicture("")
        On Error Resume Next
        Image11.Picture = LoadPicture(ARCHIMAG$)
        If Err Then
          On Error GoTo 0
          ARXX$ = REPLACAR$(ARCHIMAG$, "\", "/")
          Call MENSERR(24, "Archivo Imagen Isologo\" + ARXX$ + "\No Válido o Inexistente.")
        End If
        On Error GoTo 0
        Picture19.Visible = True
        IMAGEFILE$ = ARCHIMAG$     ' LUDAT$ + "LOGOEMPR.BMP"
        '
        FORMALTER.Image3.Picture = LoadPicture("")
        On Error Resume Next
        FORMALTER.Image3.Picture = LoadPicture(IMAGEFILE$)
        On Error GoTo 0
        '
12      HIMAGEN = FORMALTER.Image3.Picture.Width
        VIMAGEN = FORMALTER.Image3.Picture.Height
        '
        'Picture19.Width = Frame1.Left
        'Picture19.Height = 1000
        'Picture19.Left = 0
        HCUADRO = Frame1.Left   'Picture19.Width
        VCUADRO = 1000 ' Picture19.Height
        '
        REDUH = HIMAGEN / HCUADRO
        REDUV = VIMAGEN / VCUADRO
        REDUC = REDUH: If REDUV > REDUH Then REDUC = REDUV
        If REDUC < 1 Then REDUC = 1
        '
        With Picture19
          .Height = VCUADRO * REDUV / REDUC
          .Width = HCUADRO * REDUH / REDUC
          '.Left = .Left + (HCUADRO - Picture19.Width) / 2
          .Left = (HCUADRO - Picture19.Width) / 2
            If .Left < 0 Then .Left = 0
          .Top = .Top - (VCUADRO - Picture19.Height) / 2
            If .Top < 0 Then .Top = 0
        End With
        '
        Image11.Width = Picture19.Width
        Image11.Height = Picture19.Height
        Image11.Stretch = True
        On Error Resume Next
        Image11.Picture = LoadPicture(IMAGEFILE) ' FORMALTER.IMAGE11.Picture ' , PH1 + BORIZQ, PV1 + BORSUP, ANCHCUA, ALTUCUA
        On Error GoTo 0
        '
        FORMALTER.Image3.Picture = LoadPicture("")
        '
      End If
   End If
   '
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   '
   Cancel = False
   If UnloadMode = 0 Then     ' desde el botón x
      Cancel = True           ' cancela la descarga
      MsgBox "Debe Utilizar el Botón de Salida para Cerrar Correctamente la Sesión de Trabajo", vbExclamation
   End If
   '
End Sub
'
Private Sub Form_Resize()
   If WindowState = 1 Then
     Call CIERRARCH("*.*")
     Call BAJALDISCO
   End If
   Picture10.Height = ScaleHeight + 50
   Picture4.Top = ScaleHeight - Picture4.Height
   Command9.Top = Picture4.Top - Command9.Height
   Command9.Visible = True
   Command6.Top = Command9.Top - Command6.Height - 30
   Command6.Visible = True
   Picture4.Visible = True
   If Picture4.Top < Command9.Top + Command9.Height Then Command9.Top = Picture4.Top - Command9.Height
   If Command6.Top < Command4.Top + Command4.Height + 200 Then Command6.Visible = False
   If Command9.Top < Command4.Top + Command4.Height + 200 Then Command9.Visible = False
   If Picture4.Top < Command4.Top + Command4.Height Then Picture4.Visible = False
   Frame1.Height = ScaleHeight + 250
   Picture7.Height = ScaleHeight
   APIC7 = 2 * Command1.Left + Command1.Width
   Picture7.Left = ScaleWidth - APIC7 - 20
      On Error Resume Next
      DINFE& = XVALO(VALCONTROL$(USNBR%, "POSIFORPRIN", "FRAME3_TOP"))
      On Error GoTo 0
      If DINFE& < 1 Or DINFE& > Height Then DINFE& = 1500
   Frame3.Top = Height - DINFE&
   Picture3.Top = Frame3.Top + Frame3.Height
   Picture3.Height = Height - Picture3.Top
   ' Se encierra provisoriamente en On Error, ya que presenta un error al minimizar.
   On Error Resume Next
   Treeview1.Height = ScaleHeight - Treeview1.Top
   On Error GoTo 0
   '
   Call CARGAIMAGEN
   Picture7.Refresh
   Picture8.Refresh
   Picture10.Refresh
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call FINAL("")
End Sub


Private Sub frame1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   ARRABOR% = 1
End Sub

Private Sub frame1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Frame1.MousePointer = 9
   If ARRABOR% = 1 Then
      If Frame1.Left + X > 20 Then
          Frame1.Left = Frame1.Left + X
          Call ACOMODATODO
      End If
   End If
End Sub

Private Sub frame1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
   ARRABOR% = 0
End Sub

Sub ACOMODATODO()
    On Error Resume Next
    Picture10.Width = Frame1.Left + Frame1.Width + 30
    Label6.Width = Frame1.Left + 50
    Label7.Width = Frame1.Left + 50
    On Error GoTo 0
    Call CARGAIMAGEN
    Call PONELOGO
    Label6.Top = Picture19.Height
    Treeview1.Top = Label6.Top + Label6.Height - 10
    Treeview1.Height = Me.ScaleHeight - Treeview1.Top
    Picture7.Refresh
    Picture8.Refresh
    Picture10.Refresh
    '
    On Error Resume Next
    Label7.Top = Frame2.Top + Frame2.Height - 30
    Treeview1.Width = 12000
    Treeview1.Width = Frame1.Left
    On Error GoTo 0
    Picture10.Refresh
End Sub

Private Sub frame3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   ARRABOR% = 1
End Sub

Private Sub frame3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Frame3.MousePointer = 7
   If ARRABOR% = 1 Then
      Frame3.Top = Frame3.Top + Y
      On Error Resume Next
      Picture2.Height = Frame3.Top
      Picture3.Top = Frame3.Top + Frame3.Height
      Picture3.Height = Height - Picture3.Top
      On Error GoTo 0
      Picture3.Refresh
      Call CARGAIMAGEN
   End If
End Sub

Private Sub frame3_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
   ARRABOR% = 0
   DINFE& = Height - Frame3.Top
   Call GRAVALCON(USNBR%, "POSIFORPRIN", "FRAME3_TOP", FORMATNUM$(DINFE&, "F10.0"))
End Sub

Private Sub Image3_DblClick(Index As Integer)
   AREXE$ = VALCONTROL$(USNBR%, "ACCDIR-" + TRIM$(Str$(Index)), "DESTAPLI")
   LINCOMMA$ = VALCONTROL$(USNBR%, "ACCDIR-" + TRIM$(Str$(Index)), "LINCOMMA")
   On Error Resume Next
   NX1 = Shell(AREXE$ + " " + LINCOMMA$, vbNormalFocus)
   If Err Then
      MsgBox "Imposible Ejecutar '" + AREXE$ + " " + LINCOMMA$ + "'"
   End If
   On Error GoTo 0
End Sub

Private Sub Image3_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   If Button = 2 Then
      FORCONSHK.Label7 = ""
      FORCONSHK.Label6 = TRIM$(Str$(Index))
      ADIR$ = "ACCDIR-" + TRIM$(Str$(Index))
      FORCONSHK.Text1 = VALCONTROL(USNBR%, ADIR$, "DESTAPLI")
      FORCONSHK.Text2 = VALCONTROL(USNBR%, ADIR$, "LINCOMMA")
      FORCONSHK.Text3 = VALCONTROL(USNBR%, ADIR$, "ARCHICON")
      FORCONSHK.Text4 = VALCONTROL(USNBR%, ADIR$, "TOOTIPTE")
      '
      FORCONSHK.Show 1
      If FORCONSHK.Label7 = "OK" Then
         DESTAPLI$ = FORCONSHK.Text1
         LINCOMMA$ = FORCONSHK.Text2
         ARCHICON$ = FORCONSHK.Text3
         TOOTIPTE$ = FORCONSHK.Text4
         '
         If FORCONSHK.Label6 = "" Then
             INDIA& = Image3.Count
           Else
             INDIA& = XVALO(FORCONSHK.Label6)
         End If
         Call GRAVALCON(USNBR%, "ADIRECT", "CANADIR", TRIM$(Str$(1 + INDIA&)))
         '
         ADIR$ = "ACCDIR-" + TRIM$(Str$(INDIA&))
         Call GRAVALCON(USNBR%, ADIR$, "DESTAPLI", DESTAPLI$)
         Call GRAVALCON(USNBR%, ADIR$, "LINCOMMA", LINCOMMA$)
         Call GRAVALCON(USNBR%, ADIR$, "ARCHICON", ARCHICON$)
         Call GRAVALCON(USNBR%, ADIR$, "TOOTIPTE", TOOTIPTE$)
      End If
      '
      Call CARGACCDIR
   End If
End Sub

Private Sub Image4_DblClick()
    Call CONECRUN("CORREO09", "Procesar Correo Entrante")
End Sub

Private Sub Picture2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    '
    If Button = 2 Then
10      CommonDialog1.CancelError = True
        CommonDialog1.DialogTitle = "Archivo Origen de Imagen"
            ' cancelado, si no da error - CommonDialog1.Flags = &H2& + &H4& + &H8 + &H800 + &H2000 + &H8000 'cdlOFNCreatePrompt
        CommonDialog1.FILTER = "Todos los archivos (*.*)|*.*|Mapas de Bit" & _
        "(*.BMP)|*.BMP|Archivos Graficos JPEG (*.JPG)|*.JPG"
        ' Especificar el filtro predeterminado
        CommonDialog1.FilterIndex = 2
        On Error GoTo 99
        CommonDialog1.ShowOpen
        On Error GoTo 0
        '
        DoEvents
        ARIMAG$ = UCase$(CommonDialog1.filename)
        '
        If ARIMAG$ <> "" Then
           On Error Resume Next
           Image2.Picture = LoadPicture(ARIMAG$)
           If Not (Err) Then
              On Error GoTo 0
              Call GRAVALCON(USNBR%, "IMAGEN", "ARCHIMAGEN", ARIMAG$)
              Call CARGAIMAGEN
              Exit Sub
           End If
           On Error GoTo 0
           Call MENSERR(24, "Imposible Aplicar la Imagen Elegida.")
           GoTo 10
        End If
    End If
99  On Error GoTo 0
    '
End Sub

Private Sub Picture3_DblClick()
'URLopen "http://www.smn.gov.ar/"
'Inet1.Password = "I(3Lei#4"
'Inet1.USERNAME = "Juan Soto"
'Text1 = Inet1.OpenURL
End Sub

Private Sub Picture3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   If Button = 2 Then
      FORCONSHK.Label7 = ""
      FORCONSHK.Show 1
      If FORCONSHK.Label7 = "OK" Then
         DESTAPLI$ = FORCONSHK.Text1
         LINCOMMA$ = FORCONSHK.Text2
         ARCHICON$ = FORCONSHK.Text3
         TOOTIPTE$ = FORCONSHK.Text4
         '
         If FORCONSHK.Label6 = "" Then
             INDIA& = Image3.Count
             Call GRAVALCON(USNBR%, "ADIRECT", "CANADIR", TRIM$(Str$(1 + INDIA&)))
           Else
             INDIA& = XVALO(FORCONSHK.Label6)
         End If
         '
         ADIR$ = "ACCDIR-" + TRIM$(Str$(INDIA&))
         Call GRAVALCON(USNBR%, ADIR$, "DESTAPLI", DESTAPLI$)
         Call GRAVALCON(USNBR%, ADIR$, "LINCOMMA", LINCOMMA$)
         Call GRAVALCON(USNBR%, ADIR$, "ARCHICON", ARCHICON$)
         Call GRAVALCON(USNBR%, ADIR$, "TOOTIPTE", TOOTIPTE$)
      End If
      '
      Call CARGACCDIR
   End If
End Sub

Private Sub Timer1_Timer()
   ' hace girar el circulo PDCA
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
   '
End Sub

Private Sub Timer2_Timer()
   '
Exit Sub
   ' controla el envio de mails
   Call ENVIAAUTO
   '
End Sub

Function PROCACTI%()
   '
   File1.Path = LUCOM$
   File1.Pattern = "*.EXK"
   File1.Refresh
   '
   For U& = 1 To File1.ListCount
      X$ = TRIM$(UCase$(File1.List(U& - 1)))
      If X$ = "ADFLESQL.EXK" Then GoTo 9         ' EXCEPCION
      If Left$(X$, 4) = "FLEX" Then GoTo 9        ' EXCEPCION
      If Left$(X$, 6) = "SEGCOM" Then GoTo 9        ' EXCEPCION
      If X$ <> "" Then
        If X$ <> App.EXEName + ".EXK" Then
          If FileLen(LUCOM$ + X$) > 0 Then
            On Error GoTo 10
            Kill LUCOM$ + X$
            On Error GoTo 0
          End If
        End If
      End If
9  Next U&
   PROCACTI% = 0
   Exit Function
   '
10 Resume 11
11 On Error GoTo 0
   PROCACTI% = 1
   X1$ = Left$(X$, Len(X$) - 1) + "E"
   TTXX$ = "Imposible Cerrar Normalmente la Sesión de Trabajo.\   \"
   TTXX$ = TTXX$ + "Se Encuentra Abierto y Ejecutando el Proceso\"
   TTXX$ = TTXX$ + "'" + X1$ + "'\  \"
   TTXX$ = TTXX$ + "Cierre este Proceso Antes de Cerrar Sesión."
   Call COMUNI(TTXX$)
   '
End Function

Private Sub Timer3_Timer()
    ' muestra la fecha y hora actual
    HOII = HOY(HOS$)
    Label4 = HOS$
    DIASE$ = DSEM$(HOII)
    Label5 = UCase$(Left$(DIASE$, 1)) + LCase$(Mid$(DIASE$, 2))
    Label3 = Left$(Time, 8)
End Sub

Private Sub Timer4_Timer()
   '
   Static ESTAMOSTRANDO%
   Static VECES%
   '
   If ESTAMOSTRANDO% > 0 Then Exit Sub
   If Timer < UTIMETEC + 30 Then Exit Sub
   '
   ESTAMOSTRANDO% = 1
   Call ABRECONEXION
   '
   ' revisa la bandeja de mensajes
   VECES% = (VECES% + 1) Mod 60        ' PARA QUE LO HAGA CADA 10 MINUTOS
   USNU% = USNBR% Mod 100
   If USNU% < 1 Then Exit Sub
   '
   CONDI$ = "FHOR_SOLENVI < '" & AHORA$ & "' "
     If VECES% <> 0 Then               ' SOLO MUESTRA LOS DE ENVIO INMEDIATO
        CONDI$ = CONDI$ + " AND ENVIA_YA = 1"
     End If
   CONDI$ = CONDI$ + " AND ((PARA_MENSA = " & USNU% & " AND (CONF_RECEP < '01/01/01' OR CONF_RECEP IS NULL))"
   CONDI$ = CONDI$ + " OR (CCOP_MENSA = " & USNU% & " AND (CONF_RECOP < '01/01/01' OR CONF_RECOP IS NULL)))"
   '
   CAMENSA& = CantRegistros("MENSAINT", CONDI$, "NOMESS")
   If CAMENSA& > 0 Then
      Dim MENSELE As ADODB.Recordset
      Set MENSELE = New ADODB.Recordset
      Set MENSELE = FLEXCONN.Execute(FILTSQL$("SELECT NUME_MENSA FROM MENSAINT WITH(NOLOCK) WHERE " + CONDI$ + " ORDER BY FHOR_PREPA ASC"))
      With MENSELE
         Do While Not .EOF
           NUMENSAJ& = XVALO(!NUME_MENSA)
           If NUMENSAJ& > 0 Then
              SHOWMESS09.Label13 = TRIM$(Str$(NUMENSAJ&))
              SHOWMESS09.Show 1
           End If
         .MoveNext
         Loop
      End With
   End If
   ESTAMOSTRANDO% = 0
   '
End Sub

Private Sub Timer5_Timer()
    '
    If SHOWMAIL09.Visible = True Then Exit Sub
    '
    TTXX$ = LOADFILE(LUCOM$ + "STPFMAIL.CFG")
    If TRIM$(Mid$(TTXX$, 553, 8)) = "CLIENT" Then
        Exit Sub    ' ES CLIENTE, NO SERVIDOR DE CORREP
    End If
    '
    Call ABRECONEXION
    CONDI$ = "cant_envi < 1 OR PENDIENV > 0"
    QTR& = CantRegistros("BANDENVI", CONDI$)
    If QTR& > 0 Then
       MAIPE& = XVALO(VALOBADA("BANDENVI", "NUME_MAIL", CONDI$))
       If MAIPE& > 0 Then
          SHOWMAIL09.Label13 = MAIPE&
          ' SHOWMAIL09.Show 1
          Call SHOWMAIL09.Command1_Click
          CONDIX$ = "NUME_MAIL = " & MAIPE&
          Call ACTUREGISTRO("BANDENVI", "PENDIENV", CONDIX$, 0, RAFE&, "NOMESS")
          Exit Sub
       End If
    End If
    '
End Sub

Private Sub Treeview1_NodeClick(ByVal Node As ComctlLib.Node)
   '
   CLANODO$ = TRIM$(Left$(Node.Key, 64))
   PPXX% = InStr(CLANODO$, "/"): If PPXX% < 1 Then PPXX% = 1 + Len(CLANODO$)
   EXECONN$ = Left$(CLANODO$, PPXX% - 1)
   '
   If EXECONN$ = "AMAPROS" Then
         Call FISECOM09.Command16_Click
       ElseIf EXECONN$ = "CONPROGRA" Then
         Call FISECOM09.Command13_Click
       ElseIf EXECONN$ = "BUPORTEX" Then
         Call FISECOM09.Command2_Click
       ElseIf EXECONN$ = "PROSNUE" Then
         Call FISECOM09.Command21_Click
       ElseIf EXECONN$ = "IMPORDAT" Then
         If DERACCE%("IMPORCRM", "Importacion de Datos al CRM") > 1 Then
            IMPORCRM.Show 1
         End If
       ElseIf EXECONN$ = "FSETUP04" Then
         Call Command4_Click
       ElseIf EXECONN$ = "CLAVEACC" Then
         Call Command3_Click
       ElseIf EXECONN$ = "COMPACBD" Then
         Call Command5_Click
       ElseIf EXECONN$ = "GENREP07" Then
         Call Command6_Click
       ElseIf EXECONN$ = "OPORTUNUE" Then
         Call NUEVAOPORTUN
       ElseIf EXECONN$ = "OPORTUN" Then
         Call ABREOPORTUN
       ElseIf EXECONN$ = "OPORTUVIE" Then
         Call ABREOPORTUN
       ElseIf EXECONN$ = "NUEVACT" Then
         Call CONACT09.Command8_Click
       ElseIf EXECONN$ = "CONACTI" Then
         Call CONACT09.Command1_Click
         CONACT09.Show 1
       ElseIf EXECONN$ = "ARMAICIRC" Then
         Call FISECOM09.Command20_Click
       ElseIf EXECONN$ = "MENSAIN" Then
         MENSAIN09.Show 1
       ElseIf EXECONN$ = "INFORMES" Then
         Call CARGANOVED
       ElseIf EXECONN$ = "TAREPCOM" Then
        Call CARGTAREPCOM
       ElseIf EXECONN$ = "DIRECOM" Then
         Call CARGDIRECOM
       ElseIf EXECONN$ = "FUNILVE" Then
        Call CARGFUNILVE
       ElseIf EXECONN$ = "TATIPOP" Then
         Call CARGTATIPOP
       ElseIf EXECONN$ = "INGRESOID" Then
         Call INGREID
       ElseIf EXECONN$ = "BANDENVI" Then
         Call BANDENVI
       ElseIf EXECONN$ = "MAINOPRO" Then
         Call PROMAIPEN
       ElseIf EXECONN$ = "ADIRCRIT" Then
         CRITBORR09.Show 1
       ElseIf EXECONN$ = "POREMITE" Then
         Call ACPOREMITE
       ElseIf EXECONN$ = "PORDOMIN" Then
         Call ACPORDOMIN
       ElseIf EXECONN$ = "PORASUNT" Then
         Call ACPORASUNT
       ElseIf EXISTE%(App.Path + "\" + EXECONN$ + ".EXE") > 0 Then
         Call CONECRUN(EXECONN$, "")
   End If
   '
End Sub
'
Sub BANDENVI()
   '
   Call DESHASFECHA(DES%, HAS%, PERIO$)
   DESDEX$ = FETEXCOR1$(DES%)
   HASTAX$ = FETEXCOR1$(DIPOST(HAS%))
   CONDI$ = "CANT_ENVI >0 AND FHOR_ULTENVI>= '" & DESDEX$ & "' AND FHOR_ULTENVI < '" & HASTAX$ & "'"
   CONDI$ = CONDI$ + " order by FHOR_ULTENVI ASC"
   Call COPYSTRU("IBOLPEN", "MAIPENVI")
10 Call CARGALISEL("MAIPENVI", "BANDENVI", "NUME_MAIL/F5; FHOR_ULTENVI/D17;PARA_MAIL/A30;ASUN_MAIL/A48", CONDI$, "NOMESS")
11 Call SELECHO("MAIPENVI/Indice de Mails Enviados del " + FECHATEX$(DES%) + " al " + FECHATEX$(HAS%))
   NROMAI& = XVALO(VALACT1$("MAIPENVI"))
   If NROMAI& > 0 Then
     SHOWMAIL09.Label13 = NROMAI&
      SHOWMAIL09.Show 1
      GoTo 11
   End If
   '
End Sub

Sub ACPOREMITE()

End Sub

Sub ACPORDOMIN()
   '
   If POROUTLOOK% > 0 Then
      Call ACPORDOMIN_OUTL       ' VA POR EL MICROSOFT OUTLOOK
      Exit Sub
   End If
   '
End Sub

Sub ACPORDOMIN_OUTL()
   '
   Dim DOMII$(), CANVE&()
   CADOMI& = 0
   '
   Set olApp = CreateObject("Outlook.Application")
   Set Inbox = olApp.GetNamespace("Mapi").GetDefaultFolder(olFolderInbox) ' Establecer Bandeja de entrada = OlApp.GetNamespace ("Mapi"). GetDefaultFolder (olFolderInbox)
   Set InboxItems = Inbox.Items
   FIN& = InboxItems.Count
   '
   ' BARRE LA BANDEJA DE ENTRADA
   SCREENMOUSEPOINTER = 11
   U& = 0
   For Each Mailobject In InboxItems 'Para cada Mailobject En InboxItems
       U& = U& + 1
       Call TRACE(20, U&, FIN&)
       '
       ' lee los datos del mensaje
       DIELREM$ = TRIM$(Mailobject.SenderEmailAddress)
       PPXX% = InStr(DIELREM$, "@")
       If PPXX% > 0 Then
          DOMIX$ = TRIM$(LCase$(Mid$(DIELREM$, PPXX%)))
          For JJ& = 1 To CADOMI&
             If DOMII$(JJ&) = DOMIX$ Then
                CANVE&(JJ&) = CANVE&(JJ&) + 1
                GoTo 19
             End If
          Next JJ&
          '
          CADOMI& = CADOMI& + 1
          ReDim Preserve DOMII$(CADOMI&), CANVE&(CADOMI&)
          DOMII$(CADOMI&) = DOMIX$
          CANVE&(CADOMI&) = 1
       End If
19  Next
    '
    ' ORDENAR POR FRECUENCIA DESCENDENTE
    For i& = 1 To CADOMI&
      For J& = i& + 1 To CADOMI&
        If CANVE&(i&) < CANVE&(J&) Then
          Call SWAP(DOMII$(i&), DOMII$(J&))
          Call SWAP(CANVE&(i&), CANVE&(J&))
        End If
      Next J&
    Next i&
    '
    ' PRESENTA EN LISTA DE SELECCION
    For i& = 1 To CADOMI&
       Call GRAREG("!INDIBOL", MKS$(CANVE&(i&)) + DOMII$(i&), i&)
    Next i&
    Call SETULTREG("!INDIBOL", CADOMI&)
    '
    Call SELECHO("!INDIBOL/Frecuencia por Dominio")
          
End Sub

Sub ACPORASUNT()

End Sub

Sub CARGANOVED()
   '
21 Call SELECHO("INDIPROS/Indice Prospect")
   VDEVU$ = VALACT1$("INDIPROS")
   If VDEVU$ <> "" Then
      ANOTASEG09.Label26 = TRIM$(VDEVU$)
      ANOTASEG09.Show 1
   End If
   '
End Sub

Sub NUEVAOPORTUN(Optional CODICLIX$)
    '
    If CODICLIX$ = "" Then
       Call SELECHO("INDIPROS")
       CODICLIX$ = VALACT1$("INDIPROS")
    End If
    '
    If CODICLIX$ <> "" Then
       NUOPO$ = PROXIOPORT$
       Call STDFORM("NUOPOR", NUOPO$ & ";" & CODICLIX$)
    End If
    '
End Sub

Sub ABREOPORTUN()
    '
    CONDI$ = "CodOport <>'' ORDER BY CodOport"
    Call CARGALISEL("INDOPORT", "OPORTUNI", "CodOport/A12;Raso_Cli/A16;RefeOport/A36", CONDI$, "NOMESS")
    Call SELECHO("INDOPORT/INDICE DE OPORTUNIDADES")
    
    CODIOPO$ = TRIM$(VALACT1$("INDOPORT"))
    If CODIOPO$ <> "" Then
       Call STDFORM("NUOPOR", CODIOPO$)
    End If
    '
End Sub

Sub CARGAIMAGEN()
    Picture2.Top = 0
    Picture2.Left = Picture10.Width
    On Error Resume Next
    Picture2.Width = Picture7.Left - Picture10.Width
    Picture2.Height = Frame3.Top
    On Error GoTo 0
    Picture2.Refresh
    Picture3.Top = Frame3.Top + Frame3.Height
    Picture3.Left = Frame1.Left + Frame1.Width
    '
    HIMAGEN = Image2.Picture.Width
    VIMAGEN = Image2.Picture.Height
    '
    HCUADRO = Picture2.Width
    VCUADRO = Picture2.Height
    '
    REDUH = HIMAGEN / HCUADRO
    REDUV = VIMAGEN / VCUADRO
    REDUC = REDUH: If REDUV < REDUH Then REDUC = REDUV
    '
    HH1& = HIMAGEN / REDUC: VV1& = VIMAGEN / REDUC
    Picture2.PaintPicture Image2.Picture, 0, 0, HH1&, VV1& ' Picture2.Height, Picture2.Height
    Picture2.Visible = True
    Picture2.Refresh

End Sub

Function PROXIOPORT$()
   '
   If COTERMI$ = "" Then
      ITERMI$ = IDENTER$
10    COTERMI$ = Control$(ITERMI$, "COTERMI")
      If Len(COTERMI$) <> 2 Then
        OPCOTEL.Show 1
        GoTo 10
      End If
   End If
   '
   Call ABRECLIEN
   OTX$ = "OC-______-" & COTERMI$
   CAMPOSE$ = "MAX(CAST(SUBSTRING(CODOPORT,4,6) AS Integer))"
   CONDI$ = "CODOPORT LIKE '" & OTX$ & "'"
   NUCONSU& = 1 + XVALO(VALOBADA("OPORTUNI", CAMPOSE$, CONDI$))
   '
20 NEXTNU$ = TRIM$(Str$(NUCONSU&))
   While Len(NEXTNU$) < 6: NEXTNU$ = "0" + NEXTNU$: Wend
   PROXIOPORT$ = "OC-" + NEXTNU$ + "-" + COTERMI$
   '
End Function
'
Sub ENVIAAUTO()
   '
   If ESTAENVIANDO% > 0 Then
      Exit Sub
   End If
   '
   ESTAENVIANDO% = 1
   FECHAC = FETEXCOR1$(HOY(HOS$)) + " 00:00:01"
   SQLX$ = "SELECT * FROM BANDENVI WITH(NOLOCK) "
   SQLX$ = SQLX$ + " WHERE (CANT_ENVI < 1 OR CANT_ENVI IS NULL)"
   SQLX$ = SQLX$ + " OR (FREC_REPI > 0 AND FEFIN_REPI >= '" & FECHAC & "')"
   Dim MAIPEN As ADODB.Recordset
   Set MAIPEN = New ADODB.Recordset
   Set MAIPEN = FLEXCONN.Execute(FILTSQL$(SQLX$))
   '
   With MAIPEN
      Do While Not .EOF
         FHORSOLENVI = 0
         On Error Resume Next
         FHORSOLENVI = CDbl(!FHOR_SOLENVI)
         On Error GoTo 0
         FHORACTUAL = CDbl(Now)
         '
         If FHORACTUAL < FHORSOLENVI Then GoTo 19  ' NO ENVIA PORQUE NO LLEGO LA HORA PROGRAMADA
         '
         NROMAI& = XVALO(!NUME_MAIL)
         VECESENVIADO& = XVALO(!CANT_ENVI)
         If VECESENVIADO& < 1 Then Call MAISEND(NROMAI&): GoTo 19
         '
         ' AQUI CAE SI YA SE ENVIO AL MENOS UNA VEZ
         FRECUREP% = XVALO(!FREC_REPI)
         If FRECUREP% < 1 Then GoTo 19             ' SI NO SE REPITE, NO SE ENVIA
         '
         FEFINREP% = CVINT(!FEFIN_REPI)
         HOII% = HOY(HOS$)
         If HOII% > FEFINREP% Then GoTo 19         ' YA PASO LA FECHA FIN DE REPETICION
         '
         FULTENVI = CVINT(!FHOR_ULTENVI)
         FEPROEN% = DIASPOST(FULTENVI, FRECUREP%)
         If HOII% >= FEPROEN% Then Call MAISEND(NROMAI&): GoTo 19
         '
19    .MoveNext
      Loop
   End With
   ESTAENVIANDO% = 0
   '
End Sub

Sub MAISEND(NROMAI&, Optional MODO$)
    SHOWMAIL09.Label13 = NROMAI&
    If MODO$ = "SHOW" Then
         SHOWMAIL09.Show 1
       Else
         Call SHOWMAIL09.Command1_Click
    End If
End Sub

Sub PROMAIPEN()
    '
    Static YACRE%
    If YACRE% < 1 Then
       Call CREACAMPO("", "BANPENPRO", "NUME_MAIL", 4, 4, 1)
       Call CREACAMPO("", "BANPENPRO", "IDEN_MAIL", 10, 128)
       Call CREACAMPO("", "BANPENPRO", "ORIG_MAIL", 10, 256)
       Call CREACAMPO("", "BANPENPRO", "DREL_MAIL", 10, 256)
       Call CREACAMPO("", "BANPENPRO", "ASUN_MAIL", 10, 256)
       Call CREACAMPO("", "BANPENPRO", "PARA_MAIL", 10, 256)
       Call CREACAMPO("", "BANPENPRO", "TEXT_MAIL", 12, 0)
       Call CREACAMPO("", "BANPENPRO", "FHOR_RECEP", 8, 0, 1)       ' FECHA Y HORA QUE FUE PREPARADO
       Call CREACAMPO("", "BANPENPRO", "PROS_DEST", 10, 32, 1)
       Call CREACAMPO("", "BANPENPRO", "BORR_MAIL", 3, 2, 1)
       '
       Call CREACAMPO("", "ADJPENPRO", "NUME_MAIL", 4, 4, 1)
       Call CREACAMPO("", "ADJPENPRO", "NORD_MAIL", 3, 2, 1)
       Call CREACAMPO("", "ADJPENPRO", "FILE_NAME", 10, 256)
       YACRE% = 1
    End If
    '
    If COMALTER%("Desea Revisar la Bandeja de Entrada ?\\Si, Revisarla\No, Saltear Revisión") <> 2 Then
5      Screen.MousePointer = 11
       On Error Resume Next
       FORSERMAI09.Visible = True
       On Error GoTo 0
       Call FORSERMAI09.REVIBANDE               ' revisa bandeja de entrada
       FORSERMAI09.Timer1.Enabled = False
       FORSERMAI09.Hide
       Screen.MousePointer = 1
    End If
    '
10  CONDIX$ = "NUME_MAIL > 0 "
    CONDIX$ = CONDIX$ + " AND (PROS_DEST = '' OR PROS_DEST IS NULL) "
    CONDIX$ = CONDIX$ + " AND (BORR_MAIL = 0 OR BORR_MAIL IS NULL) "
    CONDIX$ = CONDIX$ + " ORDER BY FHOR_RECEP ASC"
    ' Call CARGALISEL("!INDIBOL", "BANPENPRO", "NUME_MAIL/F6; FHOR_RECEP/D17; ASUN_MAIL/A60", CONDIX$)
    Call CARGALISTA(PROMAIPE09.List1, "BANPENPRO", "NUME_MAIL/F6; FHOR_RECEP/D17; ASUN_MAIL/A60", CONDIX$)
    On Error Resume Next
    PROMAIPE09.List1.ListIndex = POLISTA
    On Error GoTo 0
    PROMAIPE09.Label9 = PROMAIPE09.List1.ListCount
    ' Call FRESHECHO("!INDOBOL")
    PROMAIPE09.Show 1
    Call FRESHALL
    If PROMAIPE09.HUBORRA = "1" Then
          GoTo 5
        ElseIf PROMAIPE09.Label1 = "2" Then
          POLISTA = PROMAIPE09.List1.ListIndex
          Call PROMAIPE09.BORRAPENDIENTES(OKX%)
          GoTo 10
    End If
    '
End Sub

Function PROXIACTI$()
    '
    Static YACREADACT%, COTERMI$
    '
    SCMA = Screen.MousePointer
    Screen.MousePointer = 11
    '
    If COTERMI$ = "" Then
       ITERMI$ = IDENTER$
10     COTERMI$ = Control$(ITERMI$, "COTERMI")
       If Len(COTERMI$) <> 2 Then
         Call MENSERR(24, "Falta Identificador de Estacion de Trabajo")
         OPCOTEL.Show 1
         GoTo 10
       End If
    End If
    '
    If YACREADACT% < 1 Then
       ' crea la tabla de contactos programados
       Call CREACAMPO("", "ACTIPROY", "NUME_ACTI", 10, 12, 1)     ' NUMERO DE ACTIVIDAD
       Call CREACAMPO("", "ACTIPROY", "Codi_Cli", 10, 12, 1)      ' CODIGO DE PROSPECTO
       Call CREACAMPO("", "ACTIPROY", "CODOPORT", 10, 12, 1)      ' CODIGO DE OPORTUNIDAD
       Call CREACAMPO("", "ACTIPROY", "CONTACTO", 10, 64, 1)      ' PERSONA DE CONTACTO
       Call CREACAMPO("", "ACTIPROY", "REFE_ACTI", 10, 128, 1)    ' REFERENCIA DE ACTIVIDAD
       Call CREACAMPO("", "ACTIPROY", "DESC_ACTI", 12, 0)         ' DESCRIPCION LARGA DE ACTIVIDAD
       Call CREACAMPO("", "ACTIPROY", "USUA_PROG", 3, 1, 1)       ' USUARIO PROGRAMADO
       Call CREACAMPO("", "ACTIPROY", "USUA_ASIG", 3, 1, 1)       ' USUARIO ASIGNADO
       Call CREACAMPO("", "ACTIPROY", "FHOR_PROG", 8, 0, 1)       ' FECHA Y HORA PROGRAMADO
       Call CREACAMPO("", "ACTIPROY", "USUA_REAL", 3, 1, 1)       ' USUARIO QUE REALIZO
       Call CREACAMPO("", "ACTIPROY", "FHOR_REAL", 8, 0, 1)       ' FECHA FIN DE REITERACION
       Call CREACAMPO("", "ACTIPROY", "COMENTAR", 12, 0)          ' COMENTARIO DE REALIZACION
       Call CREACAMPO("", "ACTIPROY", "USUA_CTRL", 3, 1, 1)       ' USUARIO DE CONTROL Y APROBACION
       Call CREACAMPO("", "ACTIPROY", "FHOR_CTRL", 8, 0, 1)       ' FECHA FIN DE REITERACION
       Call CREACAMPO("", "ACTIPROY", "STAT_ACTI", 3, 1, 1)       ' STATUS (0 = ACTIVA ; 3 = CERRADA)
       YACREADACT% = 1
    End If
    '
    SQLX$ = "SELECT * FROM ACTIPROY WITH(NOLOCK) WHERE ISNUMERIC(NUME_ACTI) = 1"
    Dim APROYE As ADODB.Recordset
    Set APROYE = New ADODB.Recordset
    APROYE.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockReadOnly, adCmdText
    '
    With APROYE
       Do While Not .EOF
          NCX$ = SAFETEXT$(!Nume_ACTI)
          If IsNumeric(NCX$) = True Then
             NCY$ = NCX$: While Len(NCY$) < 5: NCY$ = "0" + NCY$: Wend
             NCY$ = COTERMI$ + "-" + NCY$
             Call ACTUREGISTRO("ACTIPROY", "NUME_ACTI", "NUME_ACTI = '" & NCX$ & "'", NCY$, RAFE&)
          End If
       .MoveNext
       Loop
    End With
    '
    CONDI$ = "NUME_ACTI LIKE '" + COTERMI$ + "%'"
    NTAR& = 1 + XVALO(Mid$(VALOBADA("ACTIPROY", "MAX(NUME_ACTI)", CONDI$, "NOMESS"), 4, 5))
    NCY$ = TRIM$(Str$(NTAR&)): While Len(NCY$) < 5: NCY$ = "0" + NCY$: Wend
    PROXIACTI$ = COTERMI$ + "-" + NCY$
    '
    Screen.MousePointer = SCMA
    '
End Function
