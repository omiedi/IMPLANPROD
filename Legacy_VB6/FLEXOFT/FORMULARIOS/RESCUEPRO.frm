VERSION 5.00
Begin VB.Form RESCUEPRO 
   BackColor       =   &H00C0E0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Resúmenes de Cuenta Corriente - Proveedores"
   ClientHeight    =   2985
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   9300
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
   ScaleHeight     =   2985
   ScaleWidth      =   9300
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00004000&
      Height          =   3480
      Left            =   8400
      ScaleHeight     =   3420
      ScaleWidth      =   915
      TabIndex        =   24
      Top             =   0
      Width           =   975
      Begin VB.PictureBox MARCOBARRA 
         BackColor       =   &H000000FF&
         Height          =   135
         Left            =   105
         ScaleHeight     =   75
         ScaleWidth      =   585
         TabIndex        =   27
         Top             =   1260
         Width           =   645
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H00FF0000&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   0
            TabIndex        =   28
            Top             =   0
            Width           =   12000
         End
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
         Height          =   640
         Left            =   105
         Picture         =   "RESCUEPRO.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   26
         ToolTipText     =   "Emite Reporte Solicitado"
         Top             =   1575
         Width           =   650
      End
      Begin VB.CommandButton command1 
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
         Left            =   105
         Picture         =   "RESCUEPRO.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   25
         ToolTipText     =   "Abandona la Aplicación"
         Top             =   315
         Width           =   650
      End
      Begin VB.Image Image1 
         Height          =   390
         Left            =   -315
         Picture         =   "RESCUEPRO.frx":0454
         Top             =   2385
         Width           =   1515
      End
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H00C0E0C0&
      Caption         =   "Rango de Fechas"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1230
      Left            =   4305
      TabIndex        =   11
      Top             =   210
      Width           =   4005
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
         Height          =   360
         Left            =   2520
         TabIndex        =   16
         Top             =   630
         Width           =   1065
      End
      Begin VB.CommandButton Command10 
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
         Left            =   3570
         TabIndex        =   15
         Top             =   630
         Width           =   175
      End
      Begin VB.TextBox Text6 
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
         Left            =   735
         TabIndex        =   13
         Top             =   630
         Width           =   1065
      End
      Begin VB.CommandButton Command6 
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
         Left            =   1785
         TabIndex        =   12
         Top             =   630
         Width           =   175
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Hasta:"
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
         Left            =   1995
         TabIndex        =   17
         Top             =   380
         Width           =   855
      End
      Begin VB.Label Label3 
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
         Left            =   -1365
         TabIndex        =   14
         Top             =   380
         Width           =   2325
      End
   End
   Begin VB.ListBox List1 
      Height          =   540
      Left            =   8400
      Sorted          =   -1  'True
      TabIndex        =   10
      Top             =   2100
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00C0E0C0&
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
      Height          =   2640
      Left            =   105
      TabIndex        =   1
      Top             =   210
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
         TabIndex        =   7
         Top             =   525
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
         TabIndex        =   6
         Top             =   525
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
         TabIndex        =   5
         TabStop         =   0   'False
         Text            =   " "
         Top             =   945
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
         TabIndex        =   4
         Text            =   " "
         Top             =   1680
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
         TabIndex        =   3
         Top             =   1680
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
         TabIndex        =   2
         TabStop         =   0   'False
         Text            =   " "
         Top             =   2100
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
         TabIndex        =   9
         Top             =   630
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
         TabIndex        =   8
         Top             =   1785
         Width           =   2325
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00C0E0C0&
      Caption         =   "Selección y Ordenamiento"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1380
      Left            =   4305
      TabIndex        =   0
      Top             =   1470
      Width           =   4005
      Begin VB.Frame Frame2 
         BackColor       =   &H00C0E0C0&
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   435
         Left            =   210
         TabIndex        =   21
         Top             =   350
         Width           =   3690
         Begin VB.OptionButton Option1 
            BackColor       =   &H00C0E0C0&
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
            Index           =   3
            Left            =   2100
            TabIndex        =   23
            Top             =   0
            Width           =   1275
         End
         Begin VB.OptionButton Option1 
            BackColor       =   &H00C0E0C0&
            Caption         =   "Cuentas Corrientes"
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
            Left            =   210
            TabIndex        =   22
            Top             =   0
            Value           =   -1  'True
            Width           =   1380
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00C0E0C0&
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   300
         Left            =   210
         TabIndex        =   18
         Top             =   945
         Width           =   3690
         Begin VB.OptionButton Option1 
            BackColor       =   &H00C0E0C0&
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
            TabIndex        =   20
            Top             =   0
            Value           =   -1  'True
            Width           =   1380
         End
         Begin VB.OptionButton Option1 
            BackColor       =   &H00C0E0C0&
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
            Left            =   2100
            TabIndex        =   19
            Top             =   0
            Width           =   1380
         End
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00C0C0C0&
         X1              =   0
         X2              =   3990
         Y1              =   840
         Y2              =   840
      End
   End
End
Attribute VB_Name = "RESCUEPRO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
    Call CIERRARCH("*.*")
    Hide
End Sub
'

Private Sub Command10_Click()
    Call SELECHO("SELFECHA")
    If FECHANUM(VALACT1$("SELFECHA")) > 0 Then
        Text5.TEXT = VALACT1$("SELFECHA")
    End If
End Sub

Private Sub Command2_Click()
    Hide
    Call FRESHALL
    Call GERESCUEPRO
End Sub
'

Private Sub Command6_Click()
    Call SELECHO("SELFECHA")
    If FECHANUM(VALACT1$("SELFECHA")) > 0 Then
        Text6.TEXT = VALACT1$("SELFECHA")
    End If
End Sub

Private Sub Command7_Click()
   Call HELPONLINE("COGECOM")
End Sub

Private Sub Command8_Click()
    Call SELECHO("PROVED1")
    NCMAX% = Val(VALACT1$("PROVED1"))
    If NCMAX% > 0 Then
        Text8.TEXT = TRIM$(Str$(NCMAX%))
        Text7.TEXT = RASOCLI$((-1) * NCMAX%)
    End If
End Sub

Private Sub Command9_Click()
    Call SELECHO("PROVED1")
    NCMIN% = Val(VALACT1$("PROVED1"))
    If NCMIN% > 0 Then
        Text10.TEXT = TRIM$(Str$(NCMIN%))
        Text9.TEXT = RASOCLI$((-1) * NCMIN%)
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


Private Sub Text5_DblClick()
    Call SELECHO("SELFECHA")
    If FECHANUM(VALACT1$("SELFECHA")) > 0 Then
        Text5.TEXT = VALACT1$("SELFECHA")
    End If
End Sub

Private Sub Text8_Change()
    NCMAX% = Val(Text8.TEXT)
    Text7.TEXT = RASOCLI$((-1) * NCMAX%)
End Sub

Private Sub Text10_DblClick()
    Call SELECHO("PROVED1")
    NCMIN% = Val(VALACT1$("PROVED1"))
    If NCMIN% > 0 Then
        Text10.TEXT = TRIM$(Str$(NCMIN%))
        Text9.TEXT = RASOCLI$((-1) * NCMIN%)
    End If
End Sub

Private Sub Text6_DblClick()
    Call SELECHO("SELFECHA")
    If FECHANUM(VALACT1$("SELFECHA")) > 0 Then
        Text6.TEXT = VALACT1$("SELFECHA")
    End If
End Sub

Private Sub Text7_CLICK()
    Text8.SetFocus
End Sub

Private Sub Text8_DblClick()
    Call SELECHO("PROVED1")
    NCMAX% = Val(VALACT1$("PROVED1"))
    If NCMAX% > 0 Then
        Text8.TEXT = TRIM$(Str$(NCMAX%))
        Text7.TEXT = RASOCLI$((-1) * NCMAX%)
    End If
End Sub

Private Sub Text9_CLICK()
    Text10.SetFocus
End Sub



