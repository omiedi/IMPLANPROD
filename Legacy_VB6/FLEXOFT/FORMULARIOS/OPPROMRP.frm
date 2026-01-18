VERSION 5.00
Begin VB.Form OPPROMRP 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0FFC0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Configuración de Funcionamiento - MRP-II"
   ClientHeight    =   4035
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   8055
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4035
   ScaleWidth      =   8055
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox MARCOBARRA 
      BackColor       =   &H000000FF&
      Height          =   135
      Left            =   7000
      ScaleHeight     =   75
      ScaleWidth      =   795
      TabIndex        =   29
      Top             =   3045
      Width           =   855
      Begin VB.Frame BARRATRAZA 
         BackColor       =   &H00FF0000&
         BorderStyle     =   0  'None
         Height          =   500
         Left            =   0
         TabIndex        =   30
         Top             =   0
         Width           =   12000
      End
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Programación Semanal"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1170
      Left            =   3675
      TabIndex        =   26
      Top             =   2730
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
         TabIndex        =   28
         Top             =   630
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
         TabIndex        =   27
         Top             =   315
         Value           =   -1  'True
         Width           =   2850
      End
   End
   Begin VB.Frame Frame8 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Deposito Stocks"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   900
      Left            =   3675
      TabIndex        =   22
      Top             =   1785
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
         TabIndex        =   24
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
         TabIndex        =   23
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
         TabIndex        =   25
         Top             =   370
         Width           =   2115
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Lotes por Omisión"
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
      TabIndex        =   15
      Top             =   2100
      Width           =   3270
      Begin VB.Frame Frame4 
         BackColor       =   &H00C0FFC0&
         BorderStyle     =   0  'None
         Caption         =   "Frame4"
         Height          =   1380
         Left            =   2310
         TabIndex        =   20
         Top             =   350
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
            Top             =   0
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
            Top             =   315
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
            Top             =   630
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
            Top             =   945
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
         TabIndex        =   19
         Top             =   420
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
         TabIndex        =   18
         Top             =   735
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
         TabIndex        =   17
         Top             =   1050
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
         TabIndex        =   16
         Top             =   1365
         Width           =   2325
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
      Height          =   855
      Left            =   7000
      Picture         =   "OPPROMRP.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "Abandona Configuración"
      Top             =   315
      Width           =   855
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Frecuencia - Plazo"
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
      TabIndex        =   13
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
         TabIndex        =   21
         Top             =   840
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
         Top             =   420
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
         TabIndex        =   14
         Top             =   420
         Width           =   2640
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Modalidad"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1500
      Left            =   3675
      TabIndex        =   12
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
         Height          =   435
         Left            =   525
         TabIndex        =   8
         Top             =   630
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
         Height          =   435
         Left            =   525
         TabIndex        =   9
         Top             =   945
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
         Height          =   435
         Left            =   525
         TabIndex        =   7
         Top             =   315
         Value           =   -1  'True
         Width           =   1905
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
      Height          =   855
      Left            =   7000
      Picture         =   "OPPROMRP.frx":014A
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "Guarda Parámetros"
      Top             =   1260
      Width           =   855
   End
   Begin VB.Image Image2 
      Height          =   510
      Left            =   6450
      Picture         =   "OPPROMRP.frx":0454
      Top             =   3370
      Width           =   2010
   End
End
Attribute VB_Name = "OPPROMRP"
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
    MODO% = 0
    If Option1.Value = True Then MODO% = 1
    If Option2.Value = True Then MODO% = 2
    Call GRACONTROL("CANECOM", "MODOCALC", TRIM$(Str$(MODO%)))
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
    CT$ = UCase$(CONTROL$("CANECOM", "PERIODO"))
    SEMANAL% = 0: Option5.Value = True
    If InStr(CT$, "SEMAN") > 0 Then
        SEMANAL% = 1
        Option4.Value = True
    End If
    '
    CT$ = UCase$(CONTROL$("CANECOM", "PLADISTA"))
    PLADISTA% = Val(CT$)
    If PLADISTA% < 0 Then PLADISTA% = 0
    If PLADISTA% > 6 Then PLADISTA% = 6
    Text1.TEXT = TRIM$(Str$(PLADISTA%))
    '
    CT$ = UCase$(CONTROL$("CANECOM", "LOST-TER"))
    LOSTN(3) = Val(CT$): If LOSTN(3) < 0 Then LOSTN(3) = 0
    LOST$(3) = Mid$(Str$(LOSTN(3)), 2)
    Text2.TEXT = LOST$(3)
    '
    CT$ = UCase$(CONTROL$("CANECOM", "LOST-MON"))
    LOSTN(2) = Val(CT$): If LOSTN(2) < 0 Then LOSTN(2) = 0
    LOST$(2) = Mid$(Str$(LOSTN(2)), 2)
    Text3.TEXT = LOST$(2)
    '
    CT$ = UCase$(CONTROL$("CANECOM", "LOST-MEC"))
    LOSTN(1) = Val(CT$): If LOSTN(1) < 0 Then LOSTN(1) = 0
    LOST$(1) = Mid$(Str$(LOSTN(1)), 2)
    Text4.TEXT = LOST$(1)
    '
    CT$ = UCase$(CONTROL$("CANECOM", "LOST-COM"))
    LOSTN(0) = Val(CT$): If LOSTN(0) < 0 Then LOSTN(0) = 0
    LOST$(0) = Mid$(Str$(LOSTN(0)), 2)
    Text5.TEXT = LOST$(0)
    '
    CT$ = UCase$(CONTROL$("CANECOM", "REPOMIN"))
    REPOMIN = Val(CT$): If REPOMIN < 10 Then REPOMIN = 10
    REPOMINS$ = Mid$(Str$(REPOMIN), 2)
    '
    MODO% = Val(CONTROL$("CANECOM", "MODOCALC"))
    If MODO% < 1 Or MODO% > 2 Then MODO% = 0
    Option3.Value = True
    If MODO% = 1 Then Option1.Value = True
    If MODO% = 2 Then Option2.Value = True
    '
    Text7 = TRIM$(CONTROL$("PROMATER", "DEPOSTOK"))
    Label1 = ECOARCH$("TADEPOSI", Chr$(Val(Text7.TEXT)))
    '
    PRISEMIN% = 53
    If Val(CONTROL$("", "PRISEMIN")) = 1 Then PRISEMIN% = 1
    If PRISEMIN% = 1 Then Option7.Value = True
    
    Refresh
    '
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
