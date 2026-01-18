VERSION 5.00
Begin VB.Form FAPEPAGO 
   BackColor       =   &H00FFFFC0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Obligaciones Pendientes de Pago"
   ClientHeight    =   5175
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   9630
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
   ScaleHeight     =   5175
   ScaleWidth      =   9630
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
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
      Height          =   750
      Left            =   8610
      Picture         =   "FAPEPAGO.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Abandona la Aplicación"
      Top             =   420
      Width           =   855
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
      Height          =   750
      Left            =   8610
      Picture         =   "FAPEPAGO.frx":014A
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Ayuda FLEXOFT en Línea"
      Top             =   1155
      Width           =   855
   End
   Begin VB.CommandButton Command5 
      DownPicture     =   "FAPEPAGO.frx":0454
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
      Left            =   8610
      Picture         =   "FAPEPAGO.frx":059E
      Style           =   1  'Graphical
      TabIndex        =   26
      ToolTipText     =   "Re-calcula y Refresca Opciones"
      Top             =   1890
      Width           =   855
   End
   Begin VB.CommandButton Command10 
      Caption         =   "Deuda"
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
      Left            =   8610
      Picture         =   "FAPEPAGO.frx":06E8
      Style           =   1  'Graphical
      TabIndex        =   33
      ToolTipText     =   "Composición de Deuda a Fecha"
      Top             =   2625
      Width           =   855
   End
   Begin VB.ListBox List1 
      Height          =   540
      Left            =   8610
      Sorted          =   -1  'True
      TabIndex        =   27
      Top             =   105
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Frame Frame7 
      BackColor       =   &H00FFFFC0&
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
      Height          =   810
      Left            =   210
      TabIndex        =   25
      Top             =   4200
      Width           =   4005
      Begin VB.PictureBox MARCOBARRA 
         BackColor       =   &H000000FF&
         Height          =   135
         Left            =   315
         ScaleHeight     =   75
         ScaleWidth      =   3315
         TabIndex        =   28
         Top             =   420
         Width           =   3375
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
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00FFFFC0&
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
      Height          =   2430
      Left            =   210
      TabIndex        =   16
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
         TabIndex        =   22
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
         TabIndex        =   21
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
         TabIndex        =   20
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
         TabIndex        =   19
         Text            =   " "
         Top             =   1470
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
         TabIndex        =   18
         Top             =   1470
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
         TabIndex        =   17
         TabStop         =   0   'False
         Text            =   " "
         Top             =   1890
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
         TabIndex        =   24
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
         TabIndex        =   23
         Top             =   1575
         Width           =   2325
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Salida"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   810
      Left            =   4410
      TabIndex        =   13
      Top             =   4200
      Width           =   4005
      Begin VB.OptionButton Option1 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Pantalla"
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
         Index           =   4
         Left            =   525
         TabIndex        =   15
         Top             =   420
         Width           =   1380
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Impresión"
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
         Index           =   8
         Left            =   2310
         TabIndex        =   14
         Top             =   420
         Value           =   -1  'True
         Width           =   1275
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Formato de Salida"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   4410
      TabIndex        =   9
      Top             =   2310
      Width           =   4005
      Begin VB.OptionButton Option1 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Listado Tabulado"
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
         Index           =   7
         Left            =   420
         TabIndex        =   12
         Top             =   420
         Value           =   -1  'True
         Width           =   2535
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Plan Financiero de Pagos"
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
         Index           =   6
         Left            =   420
         TabIndex        =   11
         Top             =   840
         Width           =   3270
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Gráfico"
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
         Index           =   5
         Left            =   420
         TabIndex        =   10
         Top             =   1260
         Width           =   1380
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFC0&
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
      Height          =   1065
      Left            =   210
      TabIndex        =   8
      Top             =   2940
      Width           =   4005
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
         Left            =   2520
         TabIndex        =   31
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
         Left            =   3570
         TabIndex        =   30
         Top             =   420
         Width           =   175
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Límite de Vencimiento:"
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
         Left            =   105
         TabIndex        =   32
         Top             =   525
         Width           =   2325
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00FFFFC0&
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
      Height          =   1800
      Left            =   4410
      TabIndex        =   3
      Top             =   315
      Width           =   4005
      Begin VB.OptionButton Option1 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Por Fecha de Vencimiento"
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
         Left            =   420
         TabIndex        =   7
         Top             =   1365
         Width           =   2745
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00FFFFC0&
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
         TabIndex        =   6
         Top             =   840
         Width           =   1380
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00FFFFC0&
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
         TabIndex        =   5
         Top             =   840
         Value           =   -1  'True
         Width           =   1380
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Por Proveedor"
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
         Index           =   0
         Left            =   420
         TabIndex        =   4
         Top             =   420
         Width           =   1905
      End
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
      Height          =   700
      Left            =   8610
      Picture         =   "FAPEPAGO.frx":0832
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Emite Reporte Solicitado"
      Top             =   4305
      Width           =   855
   End
   Begin VB.Image Image2 
      Height          =   510
      Left            =   8040
      Picture         =   "FAPEPAGO.frx":0B3C
      Top             =   3480
      Width           =   2010
   End
End
Attribute VB_Name = "FAPEPAGO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim INIRECON%

Private Sub Command1_Click()
    Call CIERRARCH("*.*")
    Hide
End Sub

Private Sub Command10_Click()
   Command10.Enabled = False
   Command2.Enabled = False
   'Call COMPODEU
   Call COMPODEUNUE    ' POR REAPLICACION DE PAGOS, A PARTIR DEL 15/03/05
   Command2.Enabled = True
   Command10.Enabled = True
End Sub

Private Sub Command2_Click()
    Command10.Enabled = False
    Command2.Enabled = False
    Call FAPEPAG2
    If Option1(4).Value = True Then  ' pantalla
         If Option1(3).Value = False Then
            Call FINAL("*COPAPE1")
            GoTo 99
          Else
            Call FINAL("*COPAPE2")
            GoTo 99
         End If
       Else                          ' impresora
         If Option1(3).Value = False Then
            Call FINAL("*LIPAPE1")
            GoTo 99
          Else
            Call FINAL("*LIPAPE2")
            GoTo 99
         End If
    End If
99  Command2.Enabled = True
    Command10.Enabled = True
End Sub

Private Sub Command2_GotFocus()
    '
    If CARFAPEN% = 0 Then
        Refresh
        Screen.MousePointer = 11
        Call FAPEPAG1
        CARFAPEN% = 1
    End If
    '
End Sub

Private Sub Command3_Click()
    Call SELECHO("GRUCOCLI")
    Text1.TEXT = VALACT1$("GRUCOCLI")
    If TRIM$(Text1.TEXT) = "" Then
        Text2.TEXT = "Todos - Sin Clasificación"
    End If
End Sub

Private Sub Command4_Click()
    Call SELECHO("VENDEDOR")
    Text3.TEXT = VALACT1$("VENDEDOR")
    If TRIM$(Text3.TEXT) = "" Then
        Text4.TEXT = "Todos - Sin Clasificación"
    End If
End Sub

Private Sub Command5_Click()
    Call FAPEPAG1
End Sub

Private Sub Command6_Click()
    Call SELECHO("SELFECHA")
    If FECHANUM(VALACT1$("SELFECHA")) > 0 Then
        Text6.TEXT = VALACT1$("SELFECHA")
    End If
End Sub

Private Sub Command8_Click()
    Call SELECHO("ACREDOR")
    NCMAX% = Val(VALACT1$("ACREDOR"))
    If NCMAX% > 0 Then
        Text8.TEXT = TRIM$(Str$(NCMAX%))
        Text7.TEXT = RASOCLI$(NCMAX%)
    End If
End Sub

Private Sub Command9_Click()
    Call SELECHO("ACREDOR")
    NCMIN% = Val(VALACT1$("ACREDOR"))
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

Private Sub Text1_DblClick()
    Call SELECHO("GRUCOCLI")
    Text1.TEXT = VALACT1$("GRUCOCLI")
    If TRIM$(Text1.TEXT) = "" Then
        Text2.TEXT = "Todos - Sin Clasificación"
    End If
End Sub

Private Sub Text10_Change()
    NCMIN% = Val(Text10.TEXT)
    Text9.TEXT = RASOCLI$((-1) * NCMIN%)
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

Private Sub Text8_Change()
    NCMAX% = Val(Text8.TEXT)
    Text7.TEXT = RASOCLI$((-1) * NCMAX%)
End Sub

Private Sub Text10_DblClick()
    Call SELECHO("ACREDOR")
    NCMIN% = Val(VALACT1$("ACREDOR"))
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
    Call SELECHO("ACREDOR")
    NCMAX% = Val(VALACT1$("ACREDOR"))
    If NCMAX% > 0 Then
        Text8.TEXT = TRIM$(Str$(NCMAX%))
        Text7.TEXT = RASOCLI$((-1) * NCMAX%)
    End If
End Sub

Private Sub Text9_CLICK()
    Text10.SetFocus
End Sub

Sub COMPODEU()
    '
    Dim NROPO$(8196), CLIPO%(8196)
    Dim TICUE%(32767), SARRACLI#(32767), SUDEUCLI#(32767)
    'Dim TIS$(10)
    'TDAT$ = "AS,VC,FC,ND,RB,DC,NC,AS"
    'For I% = 0 To 7
    '    TIS$(I%) = Mid$(TDAT$, 3 * I% + 1, 2)
    'Next I%
    '
    Call SELECHO("SELFECHA")
    If FECHANUM(VALACT1$("SELFECHA")) < 1 Then
       Exit Sub
    End If
    '
    FLIM% = FECHANUM(VALACT1$("SELFECHA"))
    '
    FIN& = ULTREG&("PROVED1")
    For U& = 1 To FIN&
       Call TRACE(20, U&, FIN&)
       PP1$ = REGLEIDO$("PROVED1", U&)
       NC% = CVI(Left$(PP1$, 2))
       If NC% > 0 Then
          PP2$ = REGLEIDO$("PROVED2", U&)
          TICUE%(NC%) = Asc(Mid$(PP2$, 119, 1))
          If TICUE%(NC%) < 0 Or TICUE%(NC%) > 1 Then
             Call MENSERR(24, "Tipo de Cuenta no Válido\para Cuenta " + TRIM$(Str$(NC%)) + " " + TRIM$(Mid$(PP2$, 3, 30)) + ".")
             TICUE%(NC%) = 0
          End If
       End If
    Next U&
    '
    FIN& = ULTREG&("CUECORR")
    For U& = 1 To FIN&
      Call TRACE(20, U&, FIN&)
      X$ = REGLEIDO$("CUECORR", U&)
      NCLIE% = CVI(Mid$(X$, 5, 2))
      If NCLIE% > 0 Then
        If TICUE%(NCLIE%) = 0 Then
          FEDOC% = CVI(Left$(X$, 2))
          If FEDOC% <= FLIM% Then
            IMPO# = CVD(Mid$(X$, 123, 8))
            IMPO# = IMPO# - CVD(Mid$(X$, 115, 8))
            SARRACLI#(NCLIE%) = SARRACLI#(NCLIE%) + IMPO#
            SARRATO# = SARRATO# + IMPO#
            '
            SALDEX$ = Mid$(X$, 139, 60)
            For KKI% = 1 To 6
              FVI% = CVI(Mid$(SALDEX$, 10 * KKI% - 9, 2))
              SALDO# = CVD(Mid$(SALDEX$, 10 * KKI% - 7, 8))
              If FVI% > 0 Then
                If Abs(SALDO#) >= 0.005 Then
                  FF% = CVI(Left$(X$, 2))
                  NC% = CVI(Mid$(X$, 5, 2))
                  NFS$ = Mid$(X$, 7, 18)
                  IMORIG# = Abs(CVD(Mid$(X$, 115, 8)) - CVD(Mid$(X$, 123, 8)))
                  ES$ = "    "
                  If SALDO# < 0 Then
                      ES$ = " CR "
                    Else
                      If FVI% <= HOO% + 8 Then ES$ = " ** "
                      If FVI% <= HOO% Then ES$ = "!!!!"
                  End If
                  '
                  Z$ = REGBLAN$("FAPEPAG")
                  Call REPLA(Z$, MKI$(NC%), 1, 2)
                  Call REPLA(Z$, NFS$, 3, 18)
                  Call REPLA(Z$, MKI$(FF%), 21, 2)
                  Call REPLA(Z$, MKI$(FVI%), 23, 2)
                  Call REPLA(Z$, MKD$(SALDO#), 25, 8)
                  Call REPLA(Z$, MKD$(IMORIG#), 33, 8)
                  Call REPLA(Z$, ES$, 41, 4)
                  '
                  UJ& = UJ& + 1
                  Call GRAREG("FAPEPAG", Z$, UJ&)
                  SUDEUCLI#(NC%) = SUDEUCLI#(NC%) + SALDO#
                  SUDEUTO# = SUDEUTO# + SALDO#
                  '
                End If
              End If
            Next KKI%
          End If
          If FEDOC% > FLIM% Then
            CAMOPO% = CAMOPO% + 1
            If CAMOPO% > 8196 Then
              Call MENSERR(24, "Imposible Re-Calcular Composición de Deuda\en la Fecha Elegida.")
              Call FINAL("")
            End If
            NROPO$(CAMOPO%) = Mid$(X$, 7, 18)
            CLIPO%(CAMOPO%) = NCLIE%
          End If
        End If
      End If
    Next U&
    '
    '
    ' LEER APLIPAG - DESCARTAR MOVIMIENTOS ANTERIORES A FLIM%
    '              - DESCARTAR MOVIMIENTOS QUE FIGURAN EN MOVIMIENTOS POSTERIORES
    '
    FIN& = ULTREG&("APLIPAG")
    For UI& = 1 To FIN&
      X$ = REGLEIDO$("APLIPAG", UI&)
      FAPLI% = CVI(Left$(X$, 2))
      If FAPLI% > FLIM% Then
        CLIAPLI% = CVI(Mid$(X$, 3, 2))
        If CLIAPLI% > 0 Then
          If Abs(SARRACLI#(CLIAPLI%) - SUDEUCLI#(CLIAPLI%)) >= 0.005 Then
            OPOSTE$ = Mid$(X$, 5, 18)
            DOCUORI$ = Mid$(X$, 55, 18)
            'If Mid$(DOCUORI$, 3, 1) <> "-" Or Mid$(DOCUORI$, 8, 1) <> "-" Or Mid$(DOCUORI$, 17, 1) <> "-" Then
   'AAA = BBB
            'End If
            For KKI% = 1 To CAMOPO%
              If NROPO$(KKI%) = DOCUORI$ Then
                If CLIPO%(KKI%) = CLIAPLI% Then GoTo 29
              End If
            Next KKI%
            'GoTo 29
            '
24          Z$ = REGBLAN$("FAPEPAG")
            Call REPLA(Z$, MKI$(CLIAPLI%), 1, 2)
            NFS$ = DOCUORI$
            Call REPLA(Z$, NFS$, 3, 18)
            FF% = FLIM%
            Call REPLA(Z$, MKI$(FF%), 21, 2)
            FVI% = FAPLI%
            Call REPLA(Z$, MKI$(FVI%), 23, 2)
            SALDEU# = CVD(Mid$(X$, 23, 8))
            Call REPLA(Z$, MKD$(SALDEU#), 25, 8)
            IMORIG# = CVD(Mid$(X$, 73, 8))
            Call REPLA(Z$, MKD$(IMORIG#), 33, 8)
            '
            UJ& = UJ& + 1
            Call GRAREG("FAPEPAG", Z$, UJ&)
            SUDEUCLI#(CLIAPLI%) = SUDEUCLI#(CLIAPLI%) + SALDEU#
            SUDEUTO# = SUDEUTO# + SALDEU#
            '
          End If
        End If
      End If
29  Next UI&
    '
   '
    FIN& = 32767
    For UI& = 1 To FIN&
      CLIAPLI% = UI&
      DIFESAL# = SARRACLI#(CLIAPLI%) - SUDEUCLI#(CLIAPLI%)
      If Abs(DIFESAL#) >= 0.005 Then
          '
          RFF$ = RECFILT$("CUECORR", 4, MKI$(CLIAPLI%))
          For U& = Len(RFF$) To 1 Step -4
            REG& = CVS(Mid$(RFF$, U& - 3, 4))
            If REG& > 0 Then
              If REG& <= ULTREG&("CUECORR") Then
                X$ = REGLEIDO$("CUECORR", REG&)
                If CVI(Left$(X$, 2)) <= FLIM% Then
                  FF% = CVI(Left$(X$, 2))
                  FVI% = FLIM%
                  NFS$ = Mid$(X$, 7, 18)
                  IMORIG# = CVD(Mid$(X$, 123, 8)) - CVD(Mid$(X$, 115, 8))
                  ES$ = "    "
                  If Abs(IMORIG# - DIFESAL#) < 0.005 Then
                    SALDO# = IMORIG#
                    '
                    Z$ = REGBLAN$("FAPEPAG")
                    Call REPLA(Z$, MKI$(CLIAPLI%), 1, 2)
                    Call REPLA(Z$, NFS$, 3, 18)
                    Call REPLA(Z$, MKI$(FF%), 21, 2)
                    Call REPLA(Z$, MKI$(FVI%), 23, 2)
                    Call REPLA(Z$, MKD$(SALDO#), 25, 8)
                    Call REPLA(Z$, MKD$(Abs(IMORIG#)), 33, 8)
                    Call REPLA(Z$, ES$, 41, 4)
                    '
                    UJ& = UJ& + 1
                    Call GRAREG("FAPEPAG", Z$, UJ&)
                    SUDEUCLI#(CLIAPLI%) = SUDEUCLI#(CLIAPLI%) + SALDO#
                    SUDEUTO# = SUDEUTO# + SALDO#
                    GoTo 39
                    '
                  End If
                End If
              End If
            End If
          Next U&
      End If
39  Next UI&
          
              
     For UI& = 1 To FIN&
      CLIAPLI% = UI&
      DIFESAL# = SARRACLI#(CLIAPLI%) - SUDEUCLI#(CLIAPLI%)
      If Abs(DIFESAL#) >= 0.005 Then
           
          Z$ = REGBLAN$("FAPEPAG")
          Call REPLA(Z$, MKI$(CLIAPLI%), 1, 2)
          NFS$ = "Dif.Saldo"
          Call REPLA(Z$, NFS$, 3, 18)
          FF% = FLIM%
          Call REPLA(Z$, MKI$(FF%), 21, 2)
          FVI% = FLIM%
          Call REPLA(Z$, MKI$(FVI%), 23, 2)
          SALDEU# = DIFESAL#
          Call REPLA(Z$, MKD$(SALDEU#), 25, 8)
          IMORIG# = Abs(DIFESAL#)
          Call REPLA(Z$, MKD$(IMORIG#), 33, 8)
          '
          UJ& = UJ& + 1
          Call GRAREG("FAPEPAG", Z$, UJ&)
          SUDEUCLI#(CLIAPLI%) = SUDEUCLI#(CLIAPLI%) + SALDEU#
          SUDEUTO# = SUDEUTO# + SALDEU#
          '
      End If
49  Next UI&
    '
50  Call SETULTREG("FAPEPAG", UJ&)
    Call CIERRARCH("FAPEPAG")
    '              - AGREGAR A LISTA DE SVD202
    ' ORDENAR POR NUMERO O ALFABETICO Y POR FECHA EMISION
    '
    Screen.MousePointer = 11
    NCMIN% = Val(FAPEPAGO.Text10.TEXT)
    NCMAX% = Val(FAPEPAGO.Text8.TEXT)
    '
    MODORDE% = 1
    If FAPEPAGO.Option1(2) = True Then MODORDE% = 2
    If FAPEPAGO.Option1(3) = True Then MODORDE% = 3
    '
    FIN& = ULTREG&("FAPEPAG")
    UJ& = 0
    FAPEPAGO.List1.Clear
    '
    For U& = 1 To FIN&
        '
        Call TRACE(20, U&, FIN&)
        X$ = REGLEIDO$("FAPEPAG", U&)
        NC% = CVI(Left$(X$, 2))
        '
        If NC% >= NCMIN% Then
          If NC% <= NCMAX% Then
            FVI% = CVI(Mid$(X$, 23, 2))
            '
            TLISTA$ = Space$(36)
            If MODORDE% = 1 Then
                 Call REPLA(TLISTA$, CSTRING$(MKI$(NC%), 1, 6, 0, 2), 1, 6)
              ElseIf MODORDE% = 2 Then
                 Call REPLA(TLISTA$, RASOCLI$((-1) * NC%), 1, 30)
              ElseIf MODORDE% = 3 Then
                 Call REPLA(TLISTA$, CSTRING$(MKI$(FVI%), 1, 6, 0, 2), 1, 6)
            End If
            Call REPLA(TLISTA$, CSTRING$(MKS$(U&), 3, 6, 0, 2), 31, 6)
            FAPEPAGO.List1.AddItem TLISTA$
            '
          End If
        End If
        '
    Next U&
    '
    FIN& = FAPEPAGO.List1.ListCount
    '
    For U& = 1 To FIN&
        '
        Call TRACE(20, U&, FIN&)
        FAPEPAGO.List1.ListIndex = U& - 1
        REG& = Val(Mid$(FAPEPAGO.List1.TEXT, 31, 8))
        X$ = REGLEIDO$("FAPEPAG", REG&)
        '
        NC% = CVI(Left$(X$, 2))
        NFS$ = Mid$(X$, 3, 18)
        FF% = CVI(Mid$(X$, 21, 2))
        FVI% = CVI(Mid$(X$, 23, 2))
        SALDEU# = CVD(Mid$(X$, 25, 8))
        IMORIG# = CVD(Mid$(X$, 33, 8))
        ES$ = Mid$(X$, 41, 4)
        '
        Y$ = REGBLAN$("FAPEPAGX")
        Call REPLA(Y$, MKI$(NC%), 1, 2)
        Call REPLA(Y$, RASOCLI$((-1) * NC%), 3, 30)
        Call REPLA(Y$, TEL1CLI$((-1) * NC%), 33, 16)
        Call REPLA(Y$, NFS$, 49, 18)
        Call REPLA(Y$, MKI$(FF%), 67, 2)
        Call REPLA(Y$, MKD$(IMORIG#), 69, 8)
        Call REPLA(Y$, MKI$(FVI%), 77, 2)
        Call REPLA(Y$, MKD$(SALDEU#), 79, 8)
        Call REPLA(Y$, ES$, 95, 4)
        '
        UJ& = UJ& + 1
        Call GRAREG("FAPEPAGX", Y$, UJ&)
        '          '
    Next U&
    '
    Call SETULTREG("FAPEPAGX", UJ&)
    Call CIERRARCH("FAPEPAGX")
    '
    '
    FEX = FLIM%
    Call HEADERS("FAPEPAGX", "")
    Call HEADERS("FAPEPAGX", "Composición de Deuda al: " + FECHATEX$(FEX))
    '
    Call FINAL("*LIPAPE1")
    '
    Screen.MousePointer = 1

End Sub

Sub COMPODEUNUE()
    '
    Dim NROPO$(8196), CLIPO%(8196)
    Dim TICUE%(32767), SARRACLI#(32767), SUDEUCLI#(32767)
    '
    Call SELECHO("SELFECHA")
    If FECHANUM(VALACT1$("SELFECHA")) < 1 Then
       Exit Sub
    End If
    '
    FLIM% = FECHANUM(VALACT1$("SELFECHA"))
    '
    FIN& = ULTREG&("PROVED1")
    For U& = 1 To FIN&
       Call TRACE(20, U&, FIN&)
       PP1$ = REGLEIDO$("PROVED1", U&)
       NC% = CVI(Left$(PP1$, 2))
       If NC% > 0 Then
          PP2$ = REGLEIDO$("PROVED2", U&)
          TICUE%(NC%) = Asc(Mid$(PP2$, 119, 1))
          If TICUE%(NC%) < 0 Or TICUE%(NC%) > 1 Then
             Call MENSERR(24, "Tipo de Cuenta no Válido\para Cuenta " + TRIM$(Str$(NC%)) + " " + TRIM$(Mid$(PP2$, 3, 30)) + ".")
             TICUE%(NC%) = 0
          End If
       End If
    Next U&
    '
    NCLIMAX& = 0
    FIN& = ULTREG&("CUECORR")
    For U& = 1 To FIN&
      Call TRACE(20, U&, FIN&)
      X$ = REGLEIDO$("CUECORR", U&)
      NCLIE% = CVI(Mid$(X$, 5, 2))
      If NCLIE% > 0 Then
        If TICUE%(NCLIE%) = 0 Then
          FEDOC% = CVI(Left$(X$, 2))
          If FEDOC% <= FLIM% Then
            IMPO# = CVD(Mid$(X$, 123, 8))
            IMPO# = IMPO# - CVD(Mid$(X$, 115, 8))
            SARRACLI#(NCLIE%) = SARRACLI#(NCLIE%) + IMPO#
            SARRATO# = SARRATO# + IMPO#
            '
            If Abs(IMPO#) >= 0.005 Then
              If NCLIE% > NCLIMAX& Then NCLIMAX& = NCLIE%
            End If
            '
          End If
        End If
      End If
    Next U&
    '
For U& = 1 To ULTREG&("FAPEPAGX")
  Y$ = REGLEIDO$("FAPEPAGX", U&)
  NCLIE% = CVI(Left$(Y$, 2))
  SUDEUCLI#(NCLIE%) = SUDEUCLI#(NCLIE%) + CVD(Mid$(Y$, 79, 8))
  SUDEUTO# = SUDEUTO# + CVD(Mid$(Y$, 79, 8))
Next U&
For U& = 1 To 32766
  If Abs(SARRACLI#(U&) - SUDEUCLI#(U&)) >= 0.005 Then
    AAA = BBB
  End If
Next U&

    
    If NCLIMAX& > 32766 Then NCLIMAX& = 32766 ' PARA EVITAR OVERFLOW
    INIRECON% = 0
    For NCLIE% = 0 To NCLIMAX&
      If TICUE%(NCLIE%) = 0 Then  ' excluye cuentas de gastos
        NCX& = NCLIE%
        Call TRACE(20, NCX&, NCLIMAX&)
        If Abs(SARRACLI#(NCLIE%)) >= 0.005 Then
          SUMARRAS# = SARRACLI#(NCLIE%)
          Call RECONDEUPRO(NCLIE%, FLIM%, SUMARRAS#)
        End If
      End If
    Next NCLIE%
    '
    INIRECON% = 2   ' PARA PONER MARCA DE FIN DE ARCHIVO
    Call RECONDEUPRO(NCLIE%, FLIM%, SUMARRAS#)
    '
    Call CIERRARCH("FAPEPAG")
    '              - AGREGAR A LISTA DE SVD202
    ' ORDENAR POR NUMERO O ALFABETICO Y POR FECHA EMISION
    '
    Screen.MousePointer = 11
    NCMIN% = Val(FAPEPAGO.Text10.TEXT)
    NCMAX% = Val(FAPEPAGO.Text8.TEXT)
    '
    MODORDE% = 1
    If FAPEPAGO.Option1(2) = True Then MODORDE% = 2
    If FAPEPAGO.Option1(3) = True Then MODORDE% = 3
    '
    FIN& = ULTREG&("FAPEPAG")
    UJ& = 0
    FAPEPAGO.List1.Clear
    '
    For U& = 1 To FIN&
        '
        Call TRACE(20, U&, FIN&)
        X$ = REGLEIDO$("FAPEPAG", U&)
        NC% = CVI(Left$(X$, 2))
        '
        If NC% >= NCMIN% Then
          If NC% <= NCMAX% Then
            FVI% = CVI(Mid$(X$, 23, 2))
            '
            TLISTA$ = Space$(36)
            If MODORDE% = 1 Then
                 Call REPLA(TLISTA$, CSTRING$(MKI$(NC%), 1, 6, 0, 2), 1, 6)
              ElseIf MODORDE% = 2 Then
                 Call REPLA(TLISTA$, RASOCLI$((-1) * NC%), 1, 30)
              ElseIf MODORDE% = 3 Then
                 Call REPLA(TLISTA$, CSTRING$(MKI$(FVI%), 1, 6, 0, 2), 1, 6)
            End If
            Call REPLA(TLISTA$, CSTRING$(MKS$(U&), 3, 6, 0, 2), 31, 6)
            FAPEPAGO.List1.AddItem TLISTA$
            '
          End If
        End If
        '
    Next U&
    '
    FIN& = FAPEPAGO.List1.ListCount
    '
    For U& = 1 To FIN&
        '
        Call TRACE(20, U&, FIN&)
        FAPEPAGO.List1.ListIndex = U& - 1
        REG& = Val(Mid$(FAPEPAGO.List1.TEXT, 31, 8))
        X$ = REGLEIDO$("FAPEPAG", REG&)
        '
        NC% = CVI(Left$(X$, 2))
        NFS$ = Mid$(X$, 3, 18)
        FF% = CVI(Mid$(X$, 21, 2))
        FVI% = CVI(Mid$(X$, 23, 2))
        SALDEU# = CVD(Mid$(X$, 25, 8))
        IMORIG# = CVD(Mid$(X$, 33, 8))
        ES$ = Mid$(X$, 41, 4)
        '
        Y$ = REGBLAN$("FAPEPAGX")
        Call REPLA(Y$, MKI$(NC%), 1, 2)
        Call REPLA(Y$, RASOCLI$((-1) * NC%), 3, 30)
        Call REPLA(Y$, TEL1CLI$((-1) * NC%), 33, 16)
        Call REPLA(Y$, NFS$, 49, 18)
        Call REPLA(Y$, MKI$(FF%), 67, 2)
        Call REPLA(Y$, MKD$(IMORIG#), 69, 8)
        Call REPLA(Y$, MKI$(FVI%), 77, 2)
        Call REPLA(Y$, MKD$(SALDEU#), 79, 8)
        Call REPLA(Y$, ES$, 95, 4)
        '
        UJ& = UJ& + 1
        Call GRAREG("FAPEPAGX", Y$, UJ&)
        '          '
    Next U&
    '
    Call SETULTREG("FAPEPAGX", UJ&)
    Call CIERRARCH("FAPEPAGX")
    '
    '
    FEX = FLIM%
    HOII% = HOY(HOS$)
    Call HEADERS("FAPEPAGX", "")
    Call HEADERS("FAPEPAGX", "Composición de Deuda al: " + FECHATEX$(FEX) + " (Reconstruida al " + HOS$ + ")")
    '
    Call FINAL("*LIPAPE1")
    '
    Screen.MousePointer = 1

End Sub

Sub RECONDEUPRO(NCX%, FLIM%, SUMARRAS#)
    '
    Static UJ&
    Dim DOCU$(8192), IMPO#(8192), REGDO&(8192), FFDOC%(8192)
    '
    If INIRECON% = 0 Then
      INIRECON% = 1
      UJ& = 0
    End If
    '
    If INIRECON% = 2 Then
      Call SETULTREG("FAPEPAG", UJ&)
      Exit Sub
    End If
    '
    ATENTI% = 1
    '
    Call BLOQARCH("CUECORR")
    '
100 Screen.MousePointer = 11
    '
    MAXDIAL& = ULTREG&("CUECORR")
    RFF$ = RECFILT$("CUECORR", 4, MKI$(NCX%))
    '
    NUDOCU% = 0
    FIN& = Len(RFF$)
    For U& = 1 To Len(RFF$) Step 4
      'Call TRACE(20, U&, FIN&)
      REG& = CVS(Mid$(RFF$, U&, 4))
      X$ = REGLEIDO$("CUECORR", REG&)
      '
      If CVI(Mid$(X$, 1, 2)) <= FLIM% Then ' filtra por fecha limite
        NUDOCU% = NUDOCU% + 1
        If NUDOCU% > 8192 Then
           Call MENSERR(24, "Imposible Continuar Proceso.\Demasiados Movimientos en la Cuenta Elegida.\Consulte a su Soporte de Sistemas FLEXOFT.")
           Call CIERRARCH("*.*")
           Call FINAL("")
        End If
        IMPO#(NUDOCU%) = CVD(Mid$(X$, 123, 8)) - CVD(Mid$(X$, 115, 8))
        DOCU$(NUDOCU%) = Mid$(X$, 7, 2) + Mid$(X$, 12, 13)
        REGDO&(NUDOCU%) = REG&
        FFDOC%(NUDOCU%) = CVI(Mid$(X$, 1, 2))
      End If
    Next U&
    '
    FIN& = NUDOCU%
    For KKI% = 1 To NUDOCU%
       U& = KKI%
       'Call TRACE(20, U&, FIN&)
       For KKJ% = KKI% + 1 To NUDOCU%
         If FFDOC%(KKI%) > FFDOC%(KKJ%) Then
           DOCX$ = DOCU$(KKI%)
           FFX% = FFDOC%(KKI%)
           IMX# = IMPO#(KKI%)
           RGX& = REGDO&(KKI%)
           DOCU$(KKI%) = DOCU$(KKJ%)
           FFDOC%(KKI%) = FFDOC%(KKJ%)
           IMPO#(KKI%) = IMPO#(KKJ%)
           REGDO&(KKI%) = REGDO&(KKJ%)
           DOCU$(KKJ%) = DOCX$
           FFDOC%(KKJ%) = FFX%
           IMPO#(KKJ%) = IMX#
           REGDO&(KKJ%) = RGX&
         End If
       Next KKJ%
    Next KKI%
    '
    SUMAPEN# = 0
    For KKI% = 1 To NUDOCU%
       SUMAPEN# = SUMAPEN# + IMPO#(KKI%)
       If Abs(SUMAPEN#) < 0.005 Then
          For KKJ% = 1 To KKI%
             IMPO#(KKJ%) = 0
          Next KKJ%
       End If
    Next KKI%
    '
    FIN& = ULTREG&("APLIPAG")
    For U& = 1 To FIN&
       'Call TRACE(20, U&, FIN&)
       X$ = REGLEIDO$("APLIPAG", U&)
       If CVI(Mid$(X$, 3, 2)) = NCX% Then
          IAPLI# = CVD(Mid$(X$, 23, 8))
          RECI$ = Mid$(X$, 5, 18)
          RECI$ = Left$(RECI$, 2) + Mid$(RECI$, 6)
          FACU$ = Mid$(X$, 55, 18)
          FACU$ = Left$(FACU$, 2) + Mid$(FACU$, 6)
          '
          IRECIMA# = 0: IFACMAX# = 0
          For KKI% = 1 To NUDOCU%
             If DOCU$(KKI%) = RECI$ Then
               If IMPO#(KKI%) < 0 Then
                 IRECIMA# = IRECIMA# - IMPO#(KKI%)
               End If
             End If
             If DOCU$(KKI%) = FACU$ Then
               If IMPO#(KKI%) > 0 Then
                 IFACMAX# = IFACMAX# + IMPO#(KKI%)
               End If
             End If
          Next KKI%
          '
          If IAPLI# > IRECIMA# Then IAPLI# = IRECIMA#
          If IAPLI# > IFACMAX# Then IAPLI# = IFACMAX#
          '
          IRECIMA# = IAPLI#
          IFACMAX# = IAPLI#
          '
          If IAPLI# > 0 Then
             For KKI% = 1 To NUDOCU%
                If DOCU$(KKI%) = RECI$ Then
                   If IMPO#(KKI%) < 0 Then
                      IAP1# = IRECIMA#
                      If IAP1# > Abs(IMPO#(KKI%)) Then IAP1# = Abs(IMPO#(KKI%))
                      IMPO#(KKI%) = IMPO#(KKI%) + IAP1#
                      IRECIMA# = IRECIMA# - IAP1#
                   End If
                End If
                If DOCU$(KKI%) = FACU$ Then
                   If IMPO#(KKI%) > 0 Then
                      IAP2# = IFACMAX#
                      If IAP2# > Abs(IMPO#(KKI%)) Then IAP2# = Abs(IMPO#(KKI%))
                      IMPO#(KKI%) = IMPO#(KKI%) - IAP2#
                      IFACMAX# = IFACMAX# - IAP2#
                   End If
                End If
             Next KKI%
          End If
          '
        End If
    Next U&
    '
    FIN& = NUDOCU%
    For KKI% = 1 To NUDOCU%
       U& = KKI%
       'Call TRACE(20, U&, FIN&)
       For KKJ% = KKI% + 1 To NUDOCU%
         If FFDOC%(KKI%) > FFDOC%(KKJ%) Then
           DOCX$ = DOCU$(KKI%)
           FFX% = FFDOC%(KKI%)
           IMX# = IMPO#(KKI%)
           RGX& = REGDO&(KKI%)
           DOCU$(KKI%) = DOCU$(KKJ%)
           FFDOC%(KKI%) = FFDOC%(KKJ%)
           IMPO#(KKI%) = IMPO#(KKJ%)
           REGDO&(KKI%) = REGDO&(KKJ%)
           DOCU$(KKJ%) = DOCX$
           FFDOC%(KKJ%) = FFX%
           IMPO#(KKJ%) = IMX#
           REGDO&(KKJ%) = RGX&
         End If
       Next KKJ%
    Next KKI%
    '
    SUMAPEN# = 0
    FIN& = NUDOCU%
    For KKI% = 1 To NUDOCU%
       U& = KKI%: 'Call TRACE(20, U&, FIN&)
       SUMAPEN# = SUMAPEN# + IMPO#(KKI%)
       If Abs(SUMAPEN#) < 0.005 Then
          For KKJ% = 1 To KKI%
             IMPO#(KKJ%) = 0
          Next KKJ%
       End If
    Next KKI%
    '
    SUGRA# = 0
    HOII% = HOY(HO$)
    FIN& = NUDOCU%
    For KKI% = 1 To NUDOCU%
       U& = KKI%
       'Call TRACE(20, U&, FIN&)
       If REGDO&(KKI%) < 1 Then
       AAA = BBB
       End If
       If REGDO&(KKI%) > 0 Then
          X$ = REGLEIDO$("CUECORR", REGDO&(KKI%))
          If Abs(IMPO#(KKI%)) >= 0.005 Then
               FEMI% = CVI(Left$(X$, 2))
               SIG% = Sgn(IMPO#(KKI%))
               ITOT# = Abs(IMPO#(KKI%))
               '
               VENUE$ = ""
               For KKK% = 139 To 198 Step 10
                  FEVE% = CVI(Mid$(X$, KKK%, 2))
                  IPAX# = Abs(CVD(Mid$(X$, KKK% + 2, 8)))
                  If FEVE% < 1 Then FEVE% = 0: IPAX# = 0
                  If IPAX# > ITOT# Then IPAX# = ITOT#
                  If FEVE% > 0 Then
                     If IPAX# >= 0.005 Then
                        VENUE$ = VENUE$ + MKI$(FEVE%) + MKD$(IPAX# * SIG%)
                        ITOT# = ITOT# - IPAX#
                     End If
                  End If
               Next KKK%
               If ITOT# >= 0.005 Then
                  If Len(VENUE$) > 50 Then
                     IPAX# = ITOT# * SIG% + CVD(Mid$(VENUE$, 3, 8))
                     Call REPLA(VENUE$, MKD$(IPAX#), 3, 8)
                    Else
                     FEVE% = HOII%: If SIG% < 0 Then FEVE% = FEMI%
                     VENUE$ = VENUE$ + MKI$(FEVE) + MKD$(ITOT# * SIG%)
                  End If
               End If
               VENUE$ = VENUE$
               '
               Z$ = REGBLAN$("FAPEPAG")
               Call REPLA(Z$, MKI$(NCX%), 1, 2)
               Call REPLA(Z$, DOCU$(KKI%), 3, 18)
               Call REPLA(Z$, MKI$(FEMI%), 21, 2)
                 IMORIG# = CVD(Mid$(X$, 123, 8)) - CVD(Mid$(X$, 115, 8))
               Call REPLA(Z$, MKD$(IMORIG#), 33, 8)
               '
               For XU% = 1 To Len(VENUE$) Step 10
                 FVI% = CVI(Mid$(VENUE$, XU%, 2))
                 SALDO# = CVD(Mid$(VENUE$, XU% + 2, 8))
                 Call REPLA(Z$, MKI$(FVI%), 23, 2)
                 Call REPLA(Z$, MKD$(SALDO#), 25, 8)
                 Call REPLA(Z$, ES$, 41, 4)
                 '
                 UJ& = UJ& + 1
                 Call GRAREG("FAPEPAG", Z$, UJ&)
                 SUGRA# = SUGRA# + SALDO#
               Next XU%
               '
          End If
       End If
    Next KKI%
    '
    If Abs(SUMARRAS# - SUMAPEN#) >= 0.005 Then
      Call MENSERR(24, "Diferencia de Saldo Compensada en Composición de Deuda\de la Cuenta '" + TRIM$(RASOCLI$((-1) * NCX%)) + "'.")
      '
      Z$ = REGBLAN$("FAPEPAG")
      Call REPLA(Z$, MKI$(NCX%), 1, 2)
      Call REPLA(Z$, "Dif.Saldo", 3, 18)
      Call REPLA(Z$, MKI$(FLIM%), 21, 2)
      Call REPLA(Z$, MKI$(FLIM%), 23, 2)
      Call REPLA(Z$, MKD$(SUMARRAS# - SUMAPEN#), 25, 8)
      Call REPLA(Z$, MKD$(SUMARRAS# - SUMAPEN#), 33, 8)
      '
      UJ& = UJ& + 1
      Call GRAREG("FAPEPAG", Z$, UJ&)
    End If
    '
999 Screen.MousePointer = 1
    ATENTI% = 0
    '
End Sub


