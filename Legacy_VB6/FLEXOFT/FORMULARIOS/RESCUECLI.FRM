VERSION 5.00
Begin VB.Form RESCUECLI 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0FFE0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Resúmenes de Cuenta Corriente - Clientes"
   ClientHeight    =   3840
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   9540
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
   ScaleHeight     =   3840
   ScaleWidth      =   9540
   StartUpPosition =   1  'CenterOwner
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
      Height          =   840
      Left            =   8535
      Picture         =   "RESCUECLI.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Emite Reporte Solicitado"
      Top             =   2835
      Width           =   855
   End
   Begin VB.PictureBox MARCOBARRA 
      BackColor       =   &H000000FF&
      Height          =   135
      Left            =   8535
      ScaleHeight     =   75
      ScaleWidth      =   810
      TabIndex        =   32
      Top             =   2415
      Width           =   865
      Begin VB.Frame BARRATRAZA 
         BackColor       =   &H00FF0000&
         BorderStyle     =   0  'None
         Height          =   500
         Left            =   0
         TabIndex        =   33
         Top             =   0
         Width           =   12000
      End
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H00C0FFE0&
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
      Height          =   960
      Left            =   210
      TabIndex        =   25
      Top             =   2730
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
         Left            =   2415
         TabIndex        =   30
         Top             =   420
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
         Left            =   3465
         TabIndex        =   29
         Top             =   420
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
         Left            =   525
         TabIndex        =   27
         Top             =   420
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
         Left            =   1575
         TabIndex        =   26
         Top             =   420
         Width           =   175
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "H:"
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
         Left            =   1470
         TabIndex        =   31
         Top             =   525
         Width           =   855
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "D:"
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
         Left            =   -1890
         TabIndex        =   28
         Top             =   525
         Width           =   2325
      End
   End
   Begin VB.ListBox List1 
      Height          =   540
      Left            =   8505
      Sorted          =   -1  'True
      TabIndex        =   24
      Top             =   2730
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00C0FFE0&
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
      Left            =   210
      TabIndex        =   15
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
         TabIndex        =   21
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
         TabIndex        =   20
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
         TabIndex        =   19
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
         TabIndex        =   18
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
         TabIndex        =   17
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
         TabIndex        =   16
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
         TabIndex        =   23
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
         TabIndex        =   22
         Top             =   1365
         Width           =   2325
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFE0&
      Caption         =   "Filtros"
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
      TabIndex        =   6
      Top             =   315
      Width           =   4005
      Begin VB.TextBox Text4 
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
         TabIndex        =   13
         TabStop         =   0   'False
         Text            =   "Todos - Sin Clasificación"
         Top             =   1680
         Width           =   3330
      End
      Begin VB.CommandButton Command4 
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
         Top             =   1260
         Width           =   175
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
         Height          =   360
         Left            =   3045
         TabIndex        =   11
         Text            =   " "
         Top             =   1260
         Width           =   435
      End
      Begin VB.TextBox Text2 
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
         TabIndex        =   9
         TabStop         =   0   'False
         Text            =   " Todos - Sin Clasificación"
         Top             =   840
         Width           =   3330
      End
      Begin VB.CommandButton Command3 
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
         TabIndex        =   8
         Top             =   420
         Width           =   175
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
         Height          =   360
         Left            =   3045
         TabIndex        =   7
         Text            =   " "
         Top             =   420
         Width           =   435
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Vendedor:"
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
         Left            =   630
         TabIndex        =   14
         Top             =   1365
         Width           =   2325
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Grupo de Cliente:"
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
         Left            =   630
         TabIndex        =   10
         Top             =   525
         Width           =   2325
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00C0FFE0&
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
      Height          =   960
      Left            =   4410
      TabIndex        =   3
      Top             =   2730
      Width           =   4005
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0FFE0&
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
         Left            =   2205
         TabIndex        =   5
         Top             =   420
         Width           =   1380
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0FFE0&
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
         Left            =   735
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
      Left            =   8535
      Picture         =   "RESCUECLI.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Ayuda FLEXOFT en Línea"
      Top             =   1365
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
      Left            =   8535
      Picture         =   "RESCUECLI.frx":0614
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Abandona la Aplicación"
      Top             =   420
      Width           =   855
   End
End
Attribute VB_Name = "RESCUECLI"
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
    Call GERESCUECLI
End Sub
'

Private Sub Command3_Click()
    Call SELECHO("GRUCOCLI")
    Text1.TEXT = VALACT1$("GRUCOCLI")
    If TRIM$(Text1.TEXT) = "" Then
        Text2.TEXT = "Todos - Sin Clasificación"
    End If
End Sub
'

Private Sub Command4_Click()
    Call SELECHO("VENDEDOR")
    Text3.TEXT = VALACT1$("VENDEDOR")
    If TRIM$(Text3.TEXT) = "" Then
        Text4.TEXT = "Todos - Sin Clasificación"
    End If
End Sub


Private Sub Command6_Click()
    Call SELECHO("SELFECHA")
    If FECHANUM(VALACT1$("SELFECHA")) > 0 Then
        Text6.TEXT = VALACT1$("SELFECHA")
    End If
End Sub

Private Sub Command8_Click()
    Call SELECHO("DEUDOR1")
    NCMAX% = Val(VALACT1$("DEUDOR1"))
    If NCMAX% > 0 Then
        Text8.TEXT = TRIM$(Str$(NCMAX%))
        Text7.TEXT = RASOCLI$(NCMAX%)
    End If
End Sub

Private Sub Command9_Click()
    Call SELECHO("DEUDOR1")
    NCMIN% = Val(VALACT1$("DEUDOR1"))
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

Private Sub Text1_DblClick()
    Call SELECHO("GRUCOCLI")
    Text1.TEXT = VALACT1$("GRUCOCLI")
    If TRIM$(Text1.TEXT) = "" Then
        Text2.TEXT = "Todos - Sin Clasificación"
    End If
End Sub

Private Sub Text10_Change()
    NCMIN% = Val(Text10.TEXT)
    Text9.TEXT = RASOCLI$(NCMIN%)
End Sub

Private Sub Text2_CLICK()
    Text1.SetFocus
End Sub

Private Sub Text3_DblClick()
    Call SELECHO("VENDEDOR")
    Text3.TEXT = VALACT1$("VENDEDOR")
    If TRIM$(Text3.TEXT) = "" Then
        Text4.TEXT = "Todos - Sin Clasificación"
    End If
End Sub

Private Sub Text1_LostFocus()
    If Val(Text1.TEXT) < 256 Then
        GRUCLI% = Val(Text1.TEXT)
        GRUX$ = ECOARCH$("GRUCOCLI", MKI$(GRUCLI%))
        If GRUX$ <> "" Then
            Text2.TEXT = GRUX$
            Exit Sub
        End If
    End If
    Text1.TEXT = ""
    Text2.TEXT = "Todos - Sin Clasificación"
End Sub

Private Sub Text3_LostFocus()
    If Val(Text3.TEXT) < 256 Then
        VENDE% = Val(Text3.TEXT)
        VENDX$ = ECOARCH$("VENDEDOR", Chr$(VENDE%))
        If VENDX <> "" Then
            Text4.TEXT = VENDX$
            Exit Sub
        End If
    End If
    Text3.TEXT = ""
    Text4.TEXT = "Todos - Sin Clasificación"
End Sub

Private Sub Text4_CLICK()
    Text3.SetFocus
End Sub

Private Sub Text1_Change()
    If Val(Text1.TEXT) < 256 Then
        GRUCLI% = Val(Text1.TEXT)
        GRUX$ = ECOARCH$("GRUCOCLI", MKI$(GRUCLI%))
        If GRUX$ <> "" Then
            Text2.TEXT = GRUX$
            Exit Sub
        End If
    End If
End Sub

Private Sub Text3_Change()
    If Val(Text3.TEXT) < 256 Then
        VENDE% = Val(Text3.TEXT)
        VENDX$ = ECOARCH$("VENDEDOR", Chr$(VENDE%))
        If VENDX <> "" Then
            Text4.TEXT = VENDX$
            Exit Sub
        End If
    End If
End Sub

Private Sub Text5_DblClick()
    Call SELECHO("SELFECHA")
    If FECHANUM(VALACT1$("SELFECHA")) > 0 Then
        Text5.TEXT = VALACT1$("SELFECHA")
    End If
End Sub

Private Sub Text8_Change()
    NCMAX% = Val(Text8.TEXT)
    Text7.TEXT = RASOCLI$(NCMAX%)
End Sub

Private Sub Text10_DblClick()
    Call SELECHO("DEUDOR1")
    NCMIN% = Val(VALACT1$("DEUDOR1"))
    If NCMIN% > 0 Then
        Text10.TEXT = TRIM$(Str$(NCMIN%))
        Text9.TEXT = RASOCLI$(NCMIN%)
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
    Call SELECHO("DEUDOR1")
    NCMAX% = Val(VALACT1$("DEUDOR1"))
    If NCMAX% > 0 Then
        Text8.TEXT = TRIM$(Str$(NCMAX%))
        Text7.TEXT = RASOCLI$(NCMAX%)
    End If
End Sub

Private Sub Text9_Click()
    Text10.SetFocus
End Sub

Sub GERESCUECLI()
    '
    Screen.MousePointer = 11
    GRUCLI% = Val(RESCUECLI.Text1.TEXT)
    GRUX$ = TRIM$(RESCUECLI.Text2.TEXT)
    VENDE% = Val(RESCUECLI.Text3.TEXT)
    VENDX$ = TRIM$(RESCUECLI.Text4.TEXT)
    DES% = FECHANUM(RESCUECLI.Text6.TEXT)
    HAS% = FECHANUM(RESCUECLI.Text5.TEXT)
    NCMIN% = Val(RESCUECLI.Text10.TEXT)
    NCMAX% = Val(RESCUECLI.Text8.TEXT)
    '
    MODORDE% = 1
    If RESCUECLI.Option1(2) = True Then MODORDE% = 2
    '
    FIN& = ULTREG&("DEUDOR1")
    UJ& = 0
    RESCUECLI.List1.Clear
    Call BLANARCH("RESCUCLI")
    '
    FIN& = ULTREG&("DEUDOR1")
    For U& = 1 To FIN&
        '
        Call TRACE(20, U&, FIN&)
        X$ = REGLEIDO$("DEUDOR1", U&)
        NC% = CVI(Left$(X$, 2))
        '
        If NC% >= NCMIN% Then
          If NC% <= NCMAX% Then
            If GRUCLI% = 0 Or GRUCOCLI%(NC%) = GRUCLI% Then
              If VENDE% = 0 Or VENDCLI%(NC%) = VENDE% Then
                  '
                  TLISTA$ = Space$(36)
                  If MODORDE% = 1 Then
                       Call REPLA(TLISTA$, CSTRING$(MKI$(NC%), 1, 6, 0, 2), 1, 6)
                    ElseIf MODORDE% = 2 Then
                       Call REPLA(TLISTA$, RASOCLI$(NC%), 1, 30)
                  End If
                  Call REPLA(TLISTA$, CSTRING$(MKI$(NC%), 1, 6, 0, 2), 31, 6)
                  RESCUECLI.List1.AddItem TLISTA$
                  '
              End If
            End If
          End If
        End If
        '
    Next U&
    '
    FIN& = RESCUECLI.List1.ListCount
    '
    For U& = 1 To FIN&
        '
        Call TRACE(20, U&, FIN&)
        RESCUECLI.List1.ListIndex = U& - 1
        NC% = Val(Mid$(RESCUECLI.List1.TEXT, 31, 6))
        On Error Resume Next
        RESCUDEU.Text1.TEXT = TRIM$(Str$(NC%))
        RECUEC00.Text1.TEXT = TRIM$(Str$(NC%))
        Call FRESHALL
        '
        REBLA$ = REGBLAN$("RESCUCLI")
        CLIE$ = TRIM$(RESCUDEU.Text2.TEXT) + " (" + TRIM$(Str$(NC%)) + ")"
        CLIE$ = TRIM$(RECUEC00.Text2.TEXT) + " (" + TRIM$(Str$(NC%)) + ")"
        '
        DOMIE$ = TRIM$(RESCUDEU.Text3.TEXT)
        DOMIE$ = TRIM$(RECUEC00.Text3.TEXT)
        '
        LOCAE$ = TRIM$(RESCUDEU.Text5.TEXT)
        LOCAE$ = TRIM$(RECUEC00.Text5.TEXT)
        '
        LOCAF$ = RESCUDEU.Text4.TEXT
        LOCAF$ = RECUEC00.Text4.TEXT
        '
        If TRIM$(LOCAF$) <> "" Then
            LOCAE$ = TRIM$(LOCAF$) + " - " + LOCAE$
        End If
        '
        TELEE$ = "Telefax: " + TRIM$(RESCUDEU.Text7.TEXT)
        TELEE$ = "Telefax: " + TRIM$(RECUEC00.Text7.TEXT)
        '
        Call REPLA(REBLA$, MKI$(NC%), 1, 2)
        Call REPLA(REBLA$, "Sres.", 97, 6)
        Call REPLA(REBLA$, CLIE$, 103, 48)
        Call REPLA(REBLA$, DOMIE$, 151, 48)
        Call REPLA(REBLA$, LOCAE$, 199, 32)
        Call REPLA(REBLA$, TELEE$, 231, 24)
        '
        YATRANSPO% = 0: ITRANSPO# = 0
        FIN& = RECUEC00.List1.ListCount
        FIN& = RESCUDEU.List1.ListCount
        '
        For V& = 1 To FIN&
            '
            REGI$ = RESCUDEU.List1.List(V& - 1)
            REGI$ = RECUEC00.List1.List(V& - 1)
            '
            FECHA$ = Left$(REGI$, 8)
            FF% = FECHANUM(FECHA$)
            If FF% < DES% Then
                ITRANSPO# = Val(Mid$(REGI$, 63, 14))
            End If
            If FF% >= DES% Then
                If FF% <= HAS% Then
                    If YATRANSPO% = 0 Then
                        REGJ$ = REBLA$
                        FEX = DES%
                        Call REPLA(REGJ$, FECHATEX$(FEX), 3, 8)
                        Call REPLA(REGJ$, "Saldo Anterior", 11, 32)
                        Call REPLA(REGJ$, MKD$(ITRANSPO#), 63, 8)
                        Call REGAPP("RESCUCLI", REGJ$)
                        Call REGAPP("RESCUCLI", REBLA$)
                        YATRANSPO% = 1
                    End If
                    '
                    REGJ$ = REBLA$
                    DOCU$ = Mid$(REGI$, 11, 16)
                    NUDO& = Val(Mid$(REGI$, 27, 8))
                    DEBE# = Val(Mid$(REGI$, 35, 13))
                    HABER# = Val(Mid$(REGI$, 50, 13))
                    SALDO# = Val(Mid$(REGI$, 63, 14))
                    '
                    Call REPLA(REGJ$, FECHA$, 3, 8)
                    Call REPLA(REGJ$, DOCU$, 11, 32)
                    Call REPLA(REGJ$, MKS$(NUDO&), 43, 4)
                    Call REPLA(REGJ$, MKD$(DEBE#), 47, 8)
                    Call REPLA(REGJ$, MKD$(HABER#), 55, 8)
                    Call REPLA(REGJ$, MKD$(SALDO#), 63, 8)
                    Call REGAPP("RESCUCLI", REGJ$)
                    '
                End If
            End If
        Next V&
        '
        If YATRANSPO% = 1 Then
            Call REGAPP("RESCUCLI", String$(256, "="))
        End If
        '
    Next U&
    On Error GoTo 0
    '
    Call CONECRUN("*RESCCLI", "Listado Resumen de Cuenta Deudores")
    '
End Sub
'


