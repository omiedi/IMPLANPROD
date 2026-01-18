VERSION 5.00
Begin VB.Form RESCUC04 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0FFE0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Resúmenes de Cuenta Corriente - Clientes"
   ClientHeight    =   3645
   ClientLeft      =   45
   ClientTop       =   585
   ClientWidth     =   9510
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
   ScaleHeight     =   3645
   ScaleWidth      =   9510
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Picture4 
      Height          =   3900
      Left            =   8610
      ScaleHeight     =   3840
      ScaleWidth      =   1110
      TabIndex        =   36
      Top             =   0
      Visible         =   0   'False
      Width           =   1170
      Begin VB.Image Image3 
         Height          =   3900
         Left            =   0
         Picture         =   "RESCUC04.frx":0000
         Stretch         =   -1  'True
         Top             =   -120
         Width           =   975
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0FFE0&
      Caption         =   "Composición de Deuda"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   800
      Left            =   4410
      TabIndex        =   34
      Top             =   2680
      Width           =   4005
      Begin VB.CheckBox Check1 
         BackColor       =   &H00C0FFE0&
         Caption         =   "Reconstrucción Automática"
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
         Left            =   630
         TabIndex        =   35
         Top             =   400
         Width           =   2955
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00004000&
      Height          =   3900
      Left            =   8610
      ScaleHeight     =   3840
      ScaleWidth      =   1545
      TabIndex        =   29
      Top             =   0
      Width           =   1605
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
         Height          =   640
         Left            =   105
         Picture         =   "RESCUC04.frx":2398
         Style           =   1  'Graphical
         TabIndex        =   33
         ToolTipText     =   "Abandona la Aplicación"
         Top             =   260
         Width           =   650
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
         Height          =   855
         Left            =   105
         Picture         =   "RESCUC04.frx":24E2
         Style           =   1  'Graphical
         TabIndex        =   32
         ToolTipText     =   "Emite Reporte Solicitado"
         Top             =   2610
         Width           =   650
      End
      Begin VB.PictureBox MARCOBARRA 
         BackColor       =   &H000000FF&
         Height          =   135
         Left            =   105
         ScaleHeight     =   75
         ScaleWidth      =   585
         TabIndex        =   30
         Top             =   1680
         Width           =   645
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H00FF0000&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   0
            TabIndex        =   31
            Top             =   0
            Width           =   12000
         End
      End
      Begin VB.Image Image1 
         Height          =   390
         Left            =   -315
         Picture         =   "RESCUC04.frx":27EC
         Top             =   1995
         Width           =   1515
      End
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H00C0FFE0&
      Caption         =   "Rango de Fechas"
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
      Left            =   210
      TabIndex        =   22
      Top             =   2310
      Width           =   2115
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
         Left            =   630
         TabIndex        =   27
         Top             =   735
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
         Height          =   285
         Left            =   1680
         TabIndex        =   26
         Top             =   735
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
         Height          =   285
         Left            =   630
         TabIndex        =   24
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
         Height          =   285
         Left            =   1680
         TabIndex        =   23
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
         Left            =   -315
         TabIndex        =   28
         Top             =   790
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
         Left            =   -1785
         TabIndex        =   25
         Top             =   480
         Width           =   2325
      End
   End
   Begin VB.ListBox List1 
      Height          =   540
      Left            =   8505
      Sorted          =   -1  'True
      TabIndex        =   21
      Top             =   2730
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00C0FFE0&
      Caption         =   "Rango de Cuentas"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2010
      Left            =   210
      TabIndex        =   12
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
         TabIndex        =   18
         Top             =   315
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
         TabIndex        =   17
         Top             =   315
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
         TabIndex        =   16
         TabStop         =   0   'False
         Text            =   " "
         Top             =   735
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
         TabIndex        =   15
         Text            =   " "
         Top             =   1155
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
         TabIndex        =   14
         Top             =   1155
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
         TabIndex        =   13
         TabStop         =   0   'False
         Text            =   " "
         Top             =   1575
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
         TabIndex        =   20
         Top             =   420
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
         TabIndex        =   19
         Top             =   1260
         Width           =   2325
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFE0&
      Caption         =   "Filtros"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2380
      Left            =   4410
      TabIndex        =   3
      Top             =   210
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
         TabIndex        =   10
         TabStop         =   0   'False
         Text            =   "Todos - Sin Clasificación"
         Top             =   1890
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
         TabIndex        =   9
         Top             =   1470
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
         TabIndex        =   8
         Text            =   " "
         Top             =   1470
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
         TabIndex        =   6
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
         TabIndex        =   5
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
         TabIndex        =   4
         Text            =   " "
         Top             =   420
         Width           =   435
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00C0C0C0&
         X1              =   0
         X2              =   3990
         Y1              =   1260
         Y2              =   1260
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
         TabIndex        =   11
         Top             =   1575
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
         TabIndex        =   7
         Top             =   525
         Width           =   2325
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00C0FFE0&
      Caption         =   "Ordenamiento"
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
      Left            =   2415
      TabIndex        =   0
      Top             =   2310
      Width           =   1800
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
         Left            =   315
         TabIndex        =   2
         Top             =   735
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
         Left            =   315
         TabIndex        =   1
         Top             =   420
         Value           =   -1  'True
         Width           =   1380
      End
   End
   Begin VB.Menu menu 
      Caption         =   "Menú"
      Begin VB.Menu imprimir 
         Caption         =   "Imprimir"
      End
      Begin VB.Menu s 
         Caption         =   "-"
      End
      Begin VB.Menu salier 
         Caption         =   "Salir"
      End
   End
End
Attribute VB_Name = "RESCUC04"
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
    CONTROLGRABA% = 0
    If Check1.Value = 1 Then CONTROLGRABA% = 1
    '
    Hide
    Call FRESHALL
    Call GERESCUECLI
    CONTROLGRABA% = 0
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
    NCMAX% = XVALO(VALACT1$("DEUDOR1"))
    If NCMAX% > 0 Then
        Text8.TEXT = TRIM$(Str$(NCMAX%))
        Text7.TEXT = RASOCLI$(NCMAX%)
    End If
End Sub

Private Sub Command9_Click()
    Call SELECHO("DEUDOR1")
    NCMIN% = XVALO(VALACT1$("DEUDOR1"))
    If NCMIN% > 0 Then
        Text10.TEXT = TRIM$(Str$(NCMIN%))
        Text9.TEXT = RASOCLI$(NCMIN%)
    End If
End Sub

Private Sub Form_Load()
    '
    If CONTROL("NEWASPEC", "CAMBIOAS") = "SI" Then
      Picture4.Visible = True
      Me.BackColor = &HFCD7B6
      Me.Frame1.BackColor = &HFCD7B6
      Me.Frame3.BackColor = &HFCD7B6
      Me.Frame4.BackColor = &HFCD7B6
      Me.Frame5.BackColor = &HFCD7B6
      Me.Frame6.BackColor = &HFCD7B6
      Me.Option1(1).BackColor = &HFCD7B6
      Me.Option1(2).BackColor = &HFCD7B6
      Check1.BackColor = &HFCD7B6
    End If

    EPAC$ = TRIM$(EMPREAC$)
    If EPAC$ <> "" Then
        Caption = Caption + "  -  " + EPAC$
    End If
    '
End Sub

Private Sub imprimir_Click()
    Call Command2_Click
End Sub

Private Sub Option1_Click(Index As Integer)
    If Index = 0 Then
        Option1(0) = False
        Option1(1) = True
    End If
End Sub

Private Sub salier_Click()
    Command1_Click
End Sub

Private Sub Text1_DblClick()
    Call SELECHO("GRUCOCLI")
    Text1.TEXT = VALACT1$("GRUCOCLI")
    If TRIM$(Text1.TEXT) = "" Then
        Text2.TEXT = "Todos - Sin Clasificación"
    End If
End Sub

Private Sub Text10_Change()
    NCMIN% = XVALO(Text10.TEXT)
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
    If XVALO(Text1.TEXT) < 256 Then
        GruCli% = XVALO(Text1.TEXT)
        GRUX$ = ECOARCH$("GRUCOCLI", MKI$(GruCli%))
        If GRUX$ <> "" Then
            Text2.TEXT = GRUX$
            Exit Sub
        End If
    End If
    Text1.TEXT = ""
    Text2.TEXT = "Todos - Sin Clasificación"
End Sub

Private Sub Text3_LostFocus()
    If XVALO(Text3.TEXT) < 256 Then
        VENDE% = XVALO(Text3.TEXT)
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
    If XVALO(Text1.TEXT) < 256 Then
        GruCli% = XVALO(Text1.TEXT)
        GRUX$ = ECOARCH$("GRUCOCLI", MKI$(GruCli%))
        If GRUX$ <> "" Then
            Text2.TEXT = GRUX$
            Exit Sub
        End If
    End If
End Sub

Private Sub Text3_Change()
    If XVALO(Text3.TEXT) < 256 Then
        VENDE% = XVALO(Text3.TEXT)
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
    NCMAX% = XVALO(Text8.TEXT)
    Text7.TEXT = RASOCLI$(NCMAX%)
End Sub

Private Sub Text10_DblClick()
    Call SELECHO("DEUDOR1")
    NCMIN% = XVALO(VALACT1$("DEUDOR1"))
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
    NCMAX% = XVALO(VALACT1$("DEUDOR1"))
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
    GruCli% = XVALO(Text1): GRUX$ = TRIM$(Text2)
    VENDE% = XVALO(Text3): VENDX$ = TRIM$(Text4)
    DES% = FECHANUM(Text6): HAS% = FECHANUM(Text5)
    NCMIN% = XVALO(Text10): NCMAX% = XVALO(Text8)
    '
    MODORDE% = 1
    If Option1(2) = True Then MODORDE% = 2
    '
    FIN& = ULTREG&("DEUDOR1")
    UJ& = 0
    List1.Clear
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
           If GruCli% = 0 Or GRUCOCLI%(NC%) = GruCli% Then
             If VENDE% = 0 Or VENDCLI%(NC%) = VENDE% Then
               '
               TLISTA$ = Space$(36)
               If MODORDE% = 1 Then
                    Call REPLA(TLISTA$, CSTRING$(MKI$(NC%), 1, 6, 0, 2), 1, 6)
                 ElseIf MODORDE% = 2 Then
                    Call REPLA(TLISTA$, RASOCLI$(NC%), 1, 30)
               End If
               Call REPLA(TLISTA$, CSTRING$(MKI$(NC%), 1, 6, 0, 2), 31, 6)
               List1.AddItem TLISTA$
               '
             End If
           End If
         End If
       End If
       '
    Next U&
    '
    FIN& = List1.ListCount
    For U& = 1 To FIN&
      '
      Call TRACE(20, U&, FIN&)
      List1.ListIndex = U& - 1
      NC% = XVALO(Mid$(List1.TEXT, 31, 6))
      On Error Resume Next
      RECUEC04.Text1 = TRIM$(Str$(NC%))
      Call FRESHALL
      '
      If CONTROLGRABA% = 1 Then
        Call RECUEC04.Command4_Click
      End If
      
      REBLA$ = REGBLAN$("RESCUCLI")
      CLIE$ = TRIM$(RECUEC04.Text2) + " (" + TRIM$(Str$(NC%)) + ")"
      DOMIE$ = TRIM$(RECUEC04.Text3)
      LOCAE$ = TRIM$(RECUEC04.Text5)
      LOCAF$ = RECUEC04.Text4
      If TRIM$(LOCAF$) <> "" Then
        LOCAE$ = TRIM$(LOCAF$) + " - " + LOCAE$
      End If
      TELEE$ = "Telefax: " + TRIM$(RECUEC04.Text7)
      '
      Call REPLA(REBLA$, MKI$(NC%), 1, 2)
      Call REPLA(REBLA$, "Sres.", 97, 6)
      Call REPLA(REBLA$, CLIE$, 103, 48)
      Call REPLA(REBLA$, DOMIE$, 151, 48)
      Call REPLA(REBLA$, LOCAE$, 199, 32)
      Call REPLA(REBLA$, TELEE$, 231, 24)
      '
      YATRANSPO% = 0: ITRANSPO# = 0
      FIN1& = RECUEC04.List1.ListCount
      '
      For V& = 1 To FIN1&
        '
        REGI$ = RECUEC04.List1.List(V& - 1)
        '
        FECHA$ = Left$(REGI$, 8)
        FF% = FECHANUM(FECHA$)
        If FF% < DES% Then
            ITRANSPO# = XVALO(Mid$(REGI$, 63, 14))
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
            NUDO& = XVALO(Mid$(REGI$, 27, 8))
            DEBE# = XVALO(Mid$(REGI$, 35, 13))
            HABER# = XVALO(Mid$(REGI$, 50, 13))
            SALDO# = XVALO(Mid$(REGI$, 63, 14))
            SIMPES$ = TRIM$(Mid$(REGI$, 79, 2))
            IDOLA# = 0
            TCXX = XVALO(Mid$(REGI$, 81, 9))
            If TCXX >= 0.01 Then
              IDOLA# = ROUND#((Abs(DEBE# - HABER#)) / TCXX, 2)
            End If
            '
            Call REPLA(REGJ$, FECHA$, 3, 8)
            Call REPLA(REGJ$, DOCU$, 11, 32)
            Call REPLA(REGJ$, MKS$(NUDO&), 43, 4)
            Call REPLA(REGJ$, MKD$(DEBE#), 47, 8)
            Call REPLA(REGJ$, MKD$(HABER#), 55, 8)
            Call REPLA(REGJ$, MKD$(SALDO#), 63, 8)
            Call REPLA(REGJ$, SIMPES$, 71, 4)
            Call REPLA(REGJ$, MKD$(IDOLA#), 75, 8)
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
    Call CIERRARCH("RESCUCLI")
    Screen.MousePointer = 1
    '
End Sub
'


