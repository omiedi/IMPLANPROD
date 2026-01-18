VERSION 5.00
Begin VB.Form FORLIPROVE 
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Listados de Proveedores y Cuentas de Gastos"
   ClientHeight    =   3735
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   9630
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3735
   ScaleWidth      =   9630
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Avance de Proceso"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   4440
      TabIndex        =   17
      Top             =   2730
      Width           =   4005
      Begin VB.PictureBox MARCOBARRA 
         BackColor       =   &H000000FF&
         Height          =   135
         Left            =   315
         ScaleHeight     =   75
         ScaleWidth      =   3270
         TabIndex        =   18
         Top             =   450
         Width           =   3330
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H00FF0000&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   0
            TabIndex        =   19
            Top             =   0
            Width           =   12000
         End
      End
   End
   Begin VB.ListBox List1 
      Height          =   540
      Left            =   8610
      Sorted          =   -1  'True
      TabIndex        =   16
      Top             =   2205
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Rango de Cuentas"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2220
      Left            =   4410
      TabIndex        =   7
      Top             =   315
      Width           =   4005
      Begin VB.TextBox Text10 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2625
         TabIndex        =   13
         Top             =   420
         Width           =   855
      End
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
         Height          =   360
         Left            =   3465
         TabIndex        =   12
         Top             =   420
         Width           =   175
      End
      Begin VB.TextBox Text9 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   315
         TabIndex        =   11
         TabStop         =   0   'False
         Text            =   " "
         Top             =   840
         Width           =   3330
      End
      Begin VB.TextBox Text8 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2625
         TabIndex        =   10
         Text            =   " "
         Top             =   1260
         Width           =   855
      End
      Begin VB.CommandButton Command8 
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
         Height          =   360
         Left            =   3465
         TabIndex        =   9
         Top             =   1260
         Width           =   175
      End
      Begin VB.TextBox Text7 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   315
         TabIndex        =   8
         TabStop         =   0   'False
         Text            =   " "
         Top             =   1680
         Width           =   3330
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Desde:"
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
         TabIndex        =   15
         Top             =   525
         Width           =   2325
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Hasta :"
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
         TabIndex        =   14
         Top             =   1365
         Width           =   2325
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Selección"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2220
      Left            =   210
      TabIndex        =   6
      Top             =   315
      Width           =   4005
      Begin VB.OptionButton OPTION2 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Todas las Cuentas (Corrientes y Gastos)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Index           =   2
         Left            =   1050
         TabIndex        =   23
         Top             =   1470
         Value           =   -1  'True
         Width           =   2115
      End
      Begin VB.OptionButton OPTION2 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Cuentas de Gastos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Index           =   1
         Left            =   1050
         TabIndex        =   22
         Top             =   945
         Width           =   2115
      End
      Begin VB.OptionButton OPTION2 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Cuentas Corrientes de Proveedores"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Index           =   0
         Left            =   1050
         TabIndex        =   21
         Top             =   420
         Width           =   2115
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Ordenamiento"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   210
      TabIndex        =   3
      Top             =   2730
      Width           =   4005
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0FFFF&
         Caption         =   "C.Postal"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   3
         Left            =   2835
         TabIndex        =   20
         Top             =   420
         Width           =   1065
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Alfabético"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   2
         Left            =   1470
         TabIndex        =   5
         Top             =   420
         Width           =   1380
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Numérico"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   1
         Left            =   210
         TabIndex        =   4
         Top             =   420
         Value           =   -1  'True
         Width           =   1380
      End
   End
   Begin VB.CommandButton Command7 
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
      Left            =   8610
      Picture         =   "FORLIPROVE.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Ayuda FLEXOFT en Línea"
      Top             =   1365
      Width           =   855
   End
   Begin VB.CommandButton Command2 
      Caption         =   "O.K."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   730
      Left            =   8610
      Picture         =   "FORLIPROVE.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Emite Reporte Solicitado"
      Top             =   2835
      Width           =   855
   End
   Begin VB.CommandButton command1 
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
      Left            =   8610
      Picture         =   "FORLIPROVE.frx":0614
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Abandona la Aplicación"
      Top             =   420
      Width           =   855
   End
End
Attribute VB_Name = "FORLIPROVE"
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
    Command2.Enabled = False
    Call GELISPRO
    Command2.Enabled = True
    '
End Sub

Private Sub Command2_GotFocus()
    '
    If CARFAPEN% = 0 Then
        Refresh
        Screen.MousePointer = 11
        Call SELRAPRO
        CARFAPEN% = 1
    End If
    '
End Sub

Private Sub Command8_Click()
    ARCHECO$ = "PROVED1"
    If OPTION2(0).Value = True Then
         ARCHECO$ = "ACREDOR"
       ElseIf OPTION2(1).Value = True Then
         ARCHECO$ = "CUEGAST"
    End If
    Call SELECHO(ARCHECO$)
    NCMAX% = Val(VALACT1$(ARCHECO$))
    If NCMAX% > 0 Then
        Text8.TEXT = TRIM$(Str$(NCMAX%))
        Text7.TEXT = RASOCLI$(NCMAX%)
    End If
End Sub

Private Sub Command9_Click()
    ARCHECO$ = "PROVED1"
    If OPTION2(0).Value = True Then
         ARCHECO$ = "ACREDOR"
       ElseIf OPTION2(1).Value = True Then
         ARCHECO$ = "CUEGAST"
    End If
    Call SELECHO(ARCHECO$)
    NCMIN% = Val(VALACT1$(ARCHECO$))
    If NCMIN% > 0 Then
        Text10.TEXT = TRIM$(Str$(NCMIN%))
        Text9.TEXT = RASOCLI$(NCMIN%)
    End If
End Sub

Private Sub Form_Load()
    '
    EPAC$ = TRIM$(EMPREAC$)
    If EPAC$ <> "" Then
        Caption = Caption + "  -  " + EPAC$
    End If
    '
End Sub

Private Sub Option1_Click(Index As Integer)
    If Index = 0 Then
        Option1(0) = False
        Option1(1) = True
    End If
End Sub

Private Sub Text10_Change()
    NCMIN% = Val(Text10.TEXT)
    Text9.TEXT = RASOCLI$((-1) * NCMIN%)
End Sub

Private Sub Text2_CLICK()
    Text1.SetFocus
End Sub

Private Sub Text8_Change()
    NCMAX% = Val(Text8.TEXT)
    Text7.TEXT = RASOCLI$((-1) * NCMAX%)
End Sub

Private Sub Text10_DblClick()
    ARCHECO$ = "PROVED1"
    If OPTION2(0).Value = True Then
         ARCHECO$ = "ACREDOR"
       ElseIf OPTION2(1).Value = True Then
         ARCHECO$ = "CUEGAST"
    End If
    Call SELECHO(ARCHECO$)
    NCMIN% = Val(VALACT1$(ARCHECO$))
    If NCMIN% > 0 Then
        Text10.TEXT = TRIM$(Str$(NCMIN%))
        Text9.TEXT = RASOCLI$((-1) * NCMIN%)
    End If
End Sub

Private Sub Text7_CLICK()
    Text8.SetFocus
End Sub

Private Sub Text8_DblClick()
    ARCHECO$ = "PROVED1"
    If OPTION2(0).Value = True Then
         ARCHECO$ = "ACREDOR"
       ElseIf OPTION2(1).Value = True Then
         ARCHECO$ = "CUEGAST"
    End If
    Call SELECHO(ARCHECO$)
    NCMAX% = Val(VALACT1$(ARCHECO$))
    If NCMAX% > 0 Then
        Text8.TEXT = TRIM$(Str$(NCMAX%))
        Text7.TEXT = RASOCLI$((-1) * NCMAX%)
    End If
End Sub

Private Sub Text9_CLICK()
    Text10.SetFocus
End Sub
