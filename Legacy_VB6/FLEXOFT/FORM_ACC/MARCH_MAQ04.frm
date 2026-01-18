VERSION 5.00
Begin VB.Form MARCH_MAQ04 
   BackColor       =   &H00D0F0FF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Archivos Maestros - Actualización y Consulta"
   ClientHeight    =   4560
   ClientLeft      =   45
   ClientTop       =   585
   ClientWidth     =   5430
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4560
   ScaleWidth      =   5430
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00404080&
      Height          =   5055
      Left            =   4515
      ScaleHeight     =   4995
      ScaleWidth      =   1125
      TabIndex        =   15
      Top             =   0
      Width           =   1185
      Begin VB.PictureBox MARCOBARRA 
         BackColor       =   &H000000FF&
         Height          =   135
         Left            =   100
         ScaleHeight     =   75
         ScaleWidth      =   540
         TabIndex        =   18
         Top             =   3645
         Width           =   600
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H00FF0000&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   0
            TabIndex        =   19
            Top             =   105
            Width           =   12000
         End
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
         Height          =   630
         Left            =   100
         Picture         =   "MARCH_MAQ04.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   17
         ToolTipText     =   "Ayuda FLEXOFT en Línea"
         Top             =   960
         Width           =   650
      End
      Begin VB.CommandButton Command24 
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
         Height          =   640
         Left            =   100
         Picture         =   "MARCH_MAQ04.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   16
         ToolTipText     =   "Salir - Abandonar Aplicación"
         Top             =   315
         Width           =   650
      End
      Begin VB.Image Image2 
         Height          =   390
         Left            =   -315
         Picture         =   "MARCH_MAQ04.frx":0454
         Top             =   3945
         Width           =   1515
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00D0F0FF&
      Caption         =   "Actualización Bases de Datos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4215
      Left            =   210
      TabIndex        =   0
      Top             =   210
      Width           =   4110
      Begin VB.Frame Frame2 
         BackColor       =   &H00D0F0FF&
         Caption         =   "Máquinas y Equipos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1800
         Left            =   210
         TabIndex        =   20
         Top             =   315
         Width           =   3690
         Begin VB.PictureBox Picture9 
            BackColor       =   &H80000016&
            Height          =   645
            Left            =   945
            ScaleHeight     =   585
            ScaleWidth      =   2055
            TabIndex        =   23
            Top             =   945
            Width           =   2115
            Begin VB.CommandButton Command1 
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
               Picture         =   "MARCH_MAQ04.frx":2376
               Style           =   1  'Graphical
               TabIndex        =   24
               ToolTipText     =   "Repuestos e Insumos"
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label10 
               BackStyle       =   0  'Transparent
               Caption         =   "Repuestos e Insumos"
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
               TabIndex        =   25
               Top             =   105
               Width           =   1005
            End
         End
         Begin VB.PictureBox Picture3 
            BackColor       =   &H80000016&
            Height          =   645
            Left            =   630
            ScaleHeight     =   585
            ScaleWidth      =   2055
            TabIndex        =   21
            Top             =   420
            Width           =   2115
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
               Picture         =   "MARCH_MAQ04.frx":2680
               Style           =   1  'Graphical
               TabIndex        =   26
               ToolTipText     =   "Máquinas y Equipos"
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label1 
               BackStyle       =   0  'Transparent
               Caption         =   "Máquinas y Equipos"
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
               TabIndex        =   22
               Top             =   60
               Width           =   1590
            End
         End
      End
      Begin VB.Frame Frame5 
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
         TabIndex        =   2
         Top             =   4830
         Width           =   4005
         Begin VB.CommandButton Command4 
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
            TabIndex        =   8
            Top             =   420
            Width           =   1590
         End
         Begin VB.CommandButton Command5 
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
            TabIndex        =   7
            Top             =   420
            Width           =   1590
         End
         Begin VB.CommandButton Command10 
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
            TabIndex        =   6
            Top             =   1155
            Width           =   1590
         End
         Begin VB.CommandButton Command11 
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
            TabIndex        =   5
            Top             =   1155
            Width           =   1590
         End
         Begin VB.CommandButton Command12 
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
            TabIndex        =   4
            Top             =   1890
            Width           =   1590
         End
         Begin VB.CommandButton Command13 
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
            TabIndex        =   3
            Top             =   1890
            Width           =   1590
         End
      End
      Begin VB.Frame Frame8 
         BackColor       =   &H00D0F0FF&
         Caption         =   "Proveedores y Operarios"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1800
         Left            =   210
         TabIndex        =   1
         Top             =   2205
         Width           =   3690
         Begin VB.PictureBox Picture2 
            BackColor       =   &H80000016&
            Height          =   645
            Left            =   945
            ScaleHeight     =   585
            ScaleWidth      =   2055
            TabIndex        =   12
            Top             =   945
            Width           =   2115
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
               Left            =   -20
               Picture         =   "MARCH_MAQ04.frx":298A
               Style           =   1  'Graphical
               TabIndex        =   13
               ToolTipText     =   "Operarios y Especialistas"
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label3 
               BackStyle       =   0  'Transparent
               Caption         =   "Operarios y Especialistas"
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
               TabIndex        =   14
               Top             =   105
               Width           =   1590
            End
         End
         Begin VB.PictureBox Picture10 
            BackColor       =   &H80000016&
            Height          =   645
            Left            =   630
            ScaleHeight     =   585
            ScaleWidth      =   2055
            TabIndex        =   9
            Top             =   420
            Width           =   2115
            Begin VB.CommandButton Command19 
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
               Left            =   -20
               Picture         =   "MARCH_MAQ04.frx":2DCC
               Style           =   1  'Graphical
               TabIndex        =   10
               ToolTipText     =   "Proveedores y Subcontratistas"
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label11 
               BackStyle       =   0  'Transparent
               Caption         =   "Proveedores y Subcontratistas"
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
               TabIndex        =   11
               Top             =   50
               Width           =   1590
            End
         End
      End
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   5880
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Menu Archi 
      Caption         =   "Archivos Maestros"
      Begin VB.Menu MnuMaq 
         Caption         =   "Máquinas y Equipos Productivos"
      End
      Begin VB.Menu MnuRep 
         Caption         =   "Repuestos e Insumos"
      End
      Begin VB.Menu php15 
         Caption         =   "-"
      End
      Begin VB.Menu Proveed 
         Caption         =   "Proveedores y Subcontratistas"
      End
      Begin VB.Menu MnuOpe 
         Caption         =   "Operarios y Especialistas"
      End
      Begin VB.Menu PLH2 
         Caption         =   "-"
      End
      Begin VB.Menu Exit 
         Caption         =   "Terminar"
      End
   End
   Begin VB.Menu Help 
      Caption         =   "Ayuda"
   End
End
Attribute VB_Name = "MARCH_MAQ04"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Amacli_Click()
   Call Command18_Click
End Sub

Private Sub Command1_Click()
   Command1.Enabled = False
   If EXISTE%("AMASTO04.EXE") > 0 Then
       Call CONECRUN("AMASTO04", "Maestro de Items de Stock")
     Else
       AMAS_MAQ04.Show 1
   End If
   Command1.Enabled = True
End Sub

Private Sub Command18_Click()
    Command18.Enabled = False
    PAD_MAQ04.Show 1
    '
    'Call FINAL("*PADRON")
    'AppActivate Caption
    'SendKeys Chr$(1)
    Command18.Enabled = True
End Sub

  Sub Command19_Click()
    Command19.Enabled = False
    ACRE_MAQ04.Show 1
'    Call CIERRARCH("*.*")
'    ACONEC$ = "ARCHIG04/AMAPRO"
'    Call CONECRUN(ACONEC$, "Padron Maestro de Proveedores")
    Command19.Enabled = True
End Sub
Private Sub Command24_Click()
    Call CIERRARCH("*.*")
    Call BAJALDISCO
    Unload Me
End Sub

Private Sub Command29_Click()
  '
  Call HELPONLINE("MARMAQ")
  '
End Sub

Private Sub Command3_Click()
  '
  Command3.Enabled = False
    AMAEQUI04.Show 1
  Command3.Enabled = True
  '
End Sub

Private Sub Exit_Click()
   Call Command24_Click
End Sub

Private Sub Form_Load()
   '
   Call VERJOBID(OKEY%)
   If OKEY% < 1 Then Call FINAL("")
   '
   EJER% = 1
   '
   EPAC$ = TRIM$(EMPREAC$)
   If EPAC$ <> "" Then
      Caption = Caption + "  -  " + EPAC$
   End If
   Call GRATITFOR(Caption)
   '
   If Left$(APLINVO$, 6) = "AMACLI" Then
       Call Command18_Click  ' Actualización de Base de Datos de Clientes
       Call FINAL("")
     ElseIf Left$(APLINVO$, 6) = "AMASTO" Then
       Call Command1_Click   ' Actualización Maestro de Items de Stock
       Call FINAL("")
     ElseIf Left$(APLINVO$, 6) = "AMAPRO" Then
       Call Command19_Click  ' Actualización de Base de Datos de Proveedores
       Call FINAL("")
     ElseIf Left$(APLINVO$, 8) = "OPMASPRO" Then
       Call ACRE_GES04.Command14_Click  ' Valores por Omisión Base de Datos de Proveedores
       Call FINAL("")
      ElseIf Left$(APLINVO$, 8) = "LIMASTER" Then
       Call OPLIMAS04.LIMASTER(0)
       Call FINAL("")
      ElseIf Left$(APLINVO$, 8) = "TABLES" Then
       OPMAST04.Show 1
       Call FINAL("")
   End If
   '
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call CIERRARCH("*.*")
   Call BAJALDISCO
End Sub

Private Sub Help_Click()
   Call Command29_Click
End Sub
Private Sub Mastok_Click()
   Call Command1_Click
End Sub

Private Sub MnuMaq_Click()
  '
  Call Command3_Click
  '
End Sub

Private Sub MnuOpe_Click()
  '
  Call Command18_Click
  '
End Sub

Private Sub MnuRep_Click()
  '
  Call Command1_Click
  '
End Sub

Private Sub Proveed_Click()
   Call Command19_Click
End Sub
