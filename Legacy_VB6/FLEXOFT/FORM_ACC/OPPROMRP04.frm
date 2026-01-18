VERSION 5.00
Begin VB.Form OPPROMRP04 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0FFC0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Configuración de Funcionamiento - MRP-II"
   ClientHeight    =   4575
   ClientLeft      =   45
   ClientTop       =   585
   ClientWidth     =   7935
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4575
   ScaleWidth      =   7935
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame7 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Filtro por Clientes"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   750
      Left            =   3675
      TabIndex        =   32
      Top             =   3675
      Width           =   3165
      Begin VB.CheckBox Check1 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Según Pedidos en Curso"
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
         Left            =   315
         TabIndex        =   33
         Top             =   360
         Width           =   2640
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00006000&
      Height          =   5000
      Left            =   7035
      ScaleHeight     =   4935
      ScaleWidth      =   1005
      TabIndex        =   27
      Top             =   0
      Width           =   1065
      Begin VB.PictureBox Picture2 
         Height          =   4650
         Left            =   0
         ScaleHeight     =   4590
         ScaleWidth      =   1005
         TabIndex        =   34
         Top             =   0
         Visible         =   0   'False
         Width           =   1065
         Begin VB.Image Image3 
            Height          =   4650
            Left            =   0
            Picture         =   "OPPROMRP04.frx":0000
            Stretch         =   -1  'True
            Top             =   0
            Width           =   1065
         End
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Save"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   700
         Left            =   105
         Picture         =   "OPPROMRP04.frx":2398
         Style           =   1  'Graphical
         TabIndex        =   31
         ToolTipText     =   "Guarda Parámetros"
         Top             =   1050
         Width           =   650
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
         Height          =   700
         Left            =   105
         Picture         =   "OPPROMRP04.frx":26A2
         Style           =   1  'Graphical
         TabIndex        =   30
         ToolTipText     =   "Abandona Configuración"
         Top             =   350
         Width           =   650
      End
      Begin VB.PictureBox MARCOBARRA 
         BackColor       =   &H000000FF&
         Height          =   135
         Left            =   105
         ScaleHeight     =   75
         ScaleWidth      =   585
         TabIndex        =   28
         Top             =   3780
         Width           =   650
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H00FF0000&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   0
            TabIndex        =   29
            Top             =   0
            Width           =   12000
         End
      End
      Begin VB.Image Image2 
         Height          =   390
         Left            =   -315
         Picture         =   "OPPROMRP04.frx":27EC
         ToolTipText     =   "Acceso Directo a Otras Aplicaciones"
         Top             =   4080
         Width           =   1515
      End
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Programación Semanal"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1065
      Left            =   3675
      TabIndex        =   24
      Top             =   2520
      Width           =   3165
      Begin VB.OptionButton Option7 
         BackColor       =   &H00C0FFC0&
         Caption         =   "1ra Semana Incompleta = 1"
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
         TabIndex        =   26
         Top             =   600
         Width           =   2850
      End
      Begin VB.OptionButton Option6 
         BackColor       =   &H00C0FFC0&
         Caption         =   "1ra.Semana Incompleta = 53"
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
         TabIndex        =   25
         Top             =   300
         Value           =   -1  'True
         Width           =   2850
      End
   End
   Begin VB.Frame Frame8 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Deposito Stocks"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   900
      Left            =   3675
      TabIndex        =   20
      Top             =   1575
      Width           =   3165
      Begin VB.CommandButton Command9 
         Caption         =   "."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   630
         TabIndex        =   22
         Top             =   370
         Width           =   175
      End
      Begin VB.TextBox Text7 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   210
         TabIndex        =   21
         Text            =   " "
         Top             =   370
         Width           =   435
      End
      Begin VB.Label Label1 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   945
         TabIndex        =   23
         Top             =   370
         Width           =   2115
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Lotes por Omisión"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2220
      Left            =   210
      TabIndex        =   13
      Top             =   2205
      Width           =   3270
      Begin VB.Frame Frame4 
         BackColor       =   &H00C0FFC0&
         BorderStyle     =   0  'None
         Caption         =   "Frame4"
         Height          =   1905
         Left            =   2310
         TabIndex        =   18
         Top             =   160
         Width           =   850
         Begin VB.TextBox Text2 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   0
            TabIndex        =   3
            Text            =   " "
            Top             =   315
            Width           =   750
         End
         Begin VB.TextBox Text3 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   0
            TabIndex        =   4
            Text            =   " "
            Top             =   735
            Width           =   750
         End
         Begin VB.TextBox Text4 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   0
            TabIndex        =   5
            Text            =   " "
            Top             =   1155
            Width           =   750
         End
         Begin VB.TextBox Text5 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   0
            TabIndex        =   6
            Text            =   " "
            Top             =   1575
            Width           =   750
         End
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Productos Terminados"
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
         Left            =   -105
         TabIndex        =   17
         Top             =   525
         Width           =   2325
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Subconjuntos"
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
         Left            =   -105
         TabIndex        =   16
         Top             =   945
         Width           =   2325
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Mecanizados"
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
         Left            =   -105
         TabIndex        =   15
         Top             =   1365
         Width           =   2325
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Elementos Comprados"
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
         Left            =   -105
         TabIndex        =   14
         Top             =   1785
         Width           =   2325
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Frecuencia - Plazo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1905
      Left            =   210
      TabIndex        =   11
      Top             =   210
      Width           =   3270
      Begin VB.Frame Frame5 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Período"
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
         Left            =   210
         TabIndex        =   19
         Top             =   945
         Width           =   2850
         Begin VB.OptionButton Option5 
            BackColor       =   &H00C0FFC0&
            Caption         =   "Mensual"
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
            Left            =   1575
            TabIndex        =   2
            Top             =   315
            Width           =   1170
         End
         Begin VB.OptionButton Option4 
            BackColor       =   &H00C0FFC0&
            Caption         =   "Semanal"
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
            Left            =   210
            TabIndex        =   1
            Top             =   315
            Width           =   1695
         End
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2520
         TabIndex        =   0
         Text            =   " "
         Top             =   525
         Width           =   540
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Plazo de Disposición:"
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
         Left            =   -210
         TabIndex        =   12
         Top             =   525
         Width           =   2640
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Modalidad"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1290
      Left            =   3675
      TabIndex        =   10
      Top             =   210
      Width           =   3165
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Cero Stocks y Mínimos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   525
         TabIndex        =   8
         Top             =   600
         Width           =   2430
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00C0FFC0&
         Caption         =   "No Considera Negativos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   525
         TabIndex        =   9
         Top             =   900
         Width           =   2535
      End
      Begin VB.OptionButton Option3 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Proceso Normal"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   525
         TabIndex        =   7
         Top             =   300
         Value           =   -1  'True
         Width           =   1905
      End
   End
   Begin VB.Menu archivo 
      Caption         =   "Archivo"
      Begin VB.Menu grabar 
         Caption         =   "Grabar"
      End
      Begin VB.Menu s 
         Caption         =   "-"
      End
      Begin VB.Menu salir 
         Caption         =   "Salir"
      End
   End
End
Attribute VB_Name = "OPPROMRP04"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Call CIERRARCH("*.*")
    Hide
End Sub

Private Sub Command2_Click()
    '
    Dim LOSTN(4), LOST$(4)
    '
    SEMANAL% = 0: PERIO$ = "MENSUAL"
    If Option4.Value = True Then
        SEMANAL% = 1: PERIO$ = "SEMANAL"
    End If
    Call GRACONTROL("CANECOM", "PERIODO", PERIO$)
    '
    PLADISTA% = Val(Text1.TEXT)
    Call GRACONTROL("CANECOM", "PLADISTA", TRIM$(Str$(PLADISTA%)))
    '
    LOST$(3) = TRIM$(Text2.TEXT): LOSTN(3) = Val(LOST$(3))
    LOST$(2) = TRIM$(Text3.TEXT): LOSTN(2) = Val(LOST$(2))
    LOST$(1) = TRIM$(Text4.TEXT): LOSTN(1) = Val(LOST$(1))
    LOST$(0) = TRIM$(Text5.TEXT): LOSTN(0) = Val(LOST$(0))
    REPOMINS$ = "10": REPOMIN = 10
    '
    Call GRACONTROL("CANECOM", "LOST-TER", LOST$(3))
    Call GRACONTROL("CANECOM", "LOST-MON", LOST$(2))
    Call GRACONTROL("CANECOM", "LOST-MEC", LOST$(1))
    Call GRACONTROL("CANECOM", "LOST-COM", LOST$(0))
    Call GRACONTROL("CANECOM", "REPOMIN", REPOMINS$)
    '
    FILPECLI$ = "NO": If Check1.Value = 1 Then FILPECLI$ = "SI"
    Call GRACONTROL("CANECOM", "FILPECLI", FILPECLI$)
    '
    modo% = 0
    If Option1.Value = True Then modo% = 1
    If Option2.Value = True Then modo% = 2
    Call GRACONTROL("CANECOM", "MODOCALC", TRIM$(Str$(modo%)))
    '
    Call GRACONTROL$("PROMATER", "DEPOSTOK", TRIM$(Text7))
    '
    PRISEMIN% = 53: If Option7.Value = True Then PRISEMIN% = 1
    Call GRACONTROL("", "PRISEMIN", TRIM$(Str$(PRISEMIN%)))
    '
End Sub

Private Sub Command9_Click()
    Call SELECHO("TADEPOSI")
    VDEVU$ = VALACT1$("TADEPOSI")
    If VDEVU$ <> "" Then
      Text7.TEXT = VDEVU$
    End If
End Sub

Private Sub Form_Activate()
    '
    Dim LOSTN(4), LOST$(4)
    '
    DoEvents
    CT$ = UCase$(Control$("CANECOM", "PERIODO"))
    SEMANAL% = 0: Option5.Value = True
    If InStr(CT$, "SEMAN") > 0 Then
        SEMANAL% = 1
        Option4.Value = True
    End If
    '
    CT$ = UCase$(Control$("CANECOM", "PLADISTA"))
    PLADISTA% = Val(CT$)
    If PLADISTA% < 0 Then PLADISTA% = 0
    If PLADISTA% > 6 Then PLADISTA% = 6
    Text1.TEXT = TRIM$(Str$(PLADISTA%))
    '
    CT$ = UCase$(Control$("CANECOM", "LOST-TER"))
    LOSTN(3) = Val(CT$): If LOSTN(3) < 0 Then LOSTN(3) = 0
    LOST$(3) = Mid$(Str$(LOSTN(3)), 2)
    Text2.TEXT = LOST$(3)
    '
    CT$ = UCase$(Control$("CANECOM", "LOST-MON"))
    LOSTN(2) = Val(CT$): If LOSTN(2) < 0 Then LOSTN(2) = 0
    LOST$(2) = Mid$(Str$(LOSTN(2)), 2)
    Text3.TEXT = LOST$(2)
    '
    CT$ = UCase$(Control$("CANECOM", "LOST-MEC"))
    LOSTN(1) = Val(CT$): If LOSTN(1) < 0 Then LOSTN(1) = 0
    LOST$(1) = Mid$(Str$(LOSTN(1)), 2)
    Text4.TEXT = LOST$(1)
    '
    CT$ = UCase$(Control$("CANECOM", "LOST-COM"))
    LOSTN(0) = Val(CT$): If LOSTN(0) < 0 Then LOSTN(0) = 0
    LOST$(0) = Mid$(Str$(LOSTN(0)), 2)
    Text5.TEXT = LOST$(0)
    '
    CT$ = UCase$(Control$("CANECOM", "REPOMIN"))
    REPOMIN = Val(CT$): If REPOMIN < 10 Then REPOMIN = 10
    REPOMINS$ = Mid$(Str$(REPOMIN), 2)
    '
    If UCase$(Control$("CANECOM", "FILPECLI")) = "SI" Then Check1.Value = 1
    '
    modo% = Val(Control$("CANECOM", "MODOCALC"))
    If modo% < 1 Or modo% > 2 Then modo% = 0
    Option3.Value = True
    If modo% = 1 Then Option1.Value = True
    If modo% = 2 Then Option2.Value = True
    '
    Text7 = TRIM$(Control$("PROMATER", "DEPOSTOK"))
    Label1 = ECOARCH$("TADEPOSI", Chr$(Val(Text7.TEXT)))
    '
    PRISEMIN% = 53
    If Val(Control$("", "PRISEMIN")) = 1 Then PRISEMIN% = 1
    If PRISEMIN% = 1 Then Option7.Value = True
    
    Refresh
    '
End Sub

Private Sub Form_Load()
    If Control("NEWASPEC", "CAMBIOAS") = "SI" Then
      Me.Picture2.Visible = True
      Me.BackColor = &H8000000F
      Me.Label1.BackColor = &H8000000F
      Me.Label10.BackColor = &H8000000F
      Me.Label5.BackColor = &H8000000F
      Me.Label6.BackColor = &H8000000F
      Me.Label8.BackColor = &H8000000F
      Me.Label9.BackColor = &H8000000F
      Me.Option1.BackColor = &H8000000F
      Me.Option2.BackColor = &H8000000F
      Me.Option3.BackColor = &H8000000F
      Me.Option4.BackColor = &H8000000F
      Me.Option5.BackColor = &H8000000F
      Me.Option6.BackColor = &H8000000F
      Me.Option7.BackColor = &H8000000F
      
      Me.Frame1.BackColor = &H8000000F
      Me.Frame2.BackColor = &H8000000F
      Me.Frame3.BackColor = &H8000000F
      Me.Frame4.BackColor = &H8000000F
      Me.Frame5.BackColor = &H8000000F
      Me.Frame6.BackColor = &H8000000F
      Me.Frame7.BackColor = &H8000000F
      Me.Frame8.BackColor = &H8000000F
      Me.Check1.BackColor = &H8000000F
      
      
      Me.MARCOBARRA.Top = 4250
      Set MARCOBARRA.Container = Me.Picture2
      MARCOBARRA.ZOrder 0
    End If
    

End Sub

Private Sub grabar_Click()
  Call Command2_Click
End Sub

Private Sub salir_Click()
    Call Command1_Click
End Sub

Private Sub Text7_Change()
    If Val(Text7) < 1 Then
        Label1 = "Todos los Depósitos"
      Else
        Label1 = ECOARCH$("TADEPOSI", Chr$(Val(Text7)))
    End If
End Sub

Private Sub Text7_DblClick()
    Call SELECHO("TADEPOSI")
    VDEVU$ = VALACT1$("TADEPOSI")
    If VDEVU$ <> "" Then
      Text7.TEXT = VDEVU$
    End If
End Sub
