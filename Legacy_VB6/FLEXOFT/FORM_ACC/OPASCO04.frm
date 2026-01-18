VERSION 5.00
Begin VB.Form OPASCO04 
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Formulario de Impresión de Asiento Contable"
   ClientHeight    =   3525
   ClientLeft      =   45
   ClientTop       =   585
   ClientWidth     =   5730
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
   ScaleHeight     =   3525
   ScaleWidth      =   5730
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture16 
      Height          =   3600
      Left            =   4800
      ScaleHeight     =   3540
      ScaleWidth      =   915
      TabIndex        =   17
      Top             =   0
      Visible         =   0   'False
      Width           =   975
      Begin VB.Image Image3 
         Height          =   3600
         Left            =   0
         Picture         =   "OPASCO04.frx":0000
         Stretch         =   -1  'True
         Top             =   0
         Width           =   1095
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Tablas"
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
      Left            =   3675
      TabIndex        =   15
      Top             =   2205
      Width           =   990
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
         Height          =   495
         Index           =   1
         Left            =   210
         Picture         =   "OPASCO04.frx":2398
         Style           =   1  'Graphical
         TabIndex        =   16
         ToolTipText     =   "Referencias Normalizadas"
         Top             =   420
         Width           =   570
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Detalle por Pase"
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
      TabIndex        =   12
      Top             =   2205
      Width           =   3510
      Begin VB.OptionButton Option2 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Referencia del Asiento"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   315
         TabIndex        =   14
         Top             =   735
         Width           =   3060
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Denominación Cuenta Contable"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   315
         TabIndex        =   13
         Top             =   420
         Value           =   -1  'True
         Width           =   3060
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Formulario de Impresión"
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
      Left            =   90
      TabIndex        =   5
      Top             =   180
      Width           =   4560
      Begin VB.CheckBox Check1 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Suprime Impresión del Comprobante"
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
         Left            =   630
         TabIndex        =   11
         Top             =   1440
         Width           =   3615
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Abrir"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   810
         Index           =   0
         Left            =   3690
         Picture         =   "OPASCO04.frx":24E2
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   420
         Width           =   735
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
         Height          =   345
         Left            =   2085
         TabIndex        =   9
         Top             =   450
         Width           =   1320
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
         Height          =   330
         Left            =   3390
         TabIndex        =   6
         Top             =   450
         Width           =   175
      End
      Begin VB.Label Label3 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   180
         TabIndex        =   8
         Top             =   900
         Width           =   3405
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Asiento Contable:"
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
         Left            =   -420
         TabIndex        =   7
         Top             =   550
         Width           =   2430
      End
   End
   Begin VB.PictureBox Picture7 
      BackColor       =   &H00404000&
      Height          =   3645
      Left            =   4830
      ScaleHeight     =   3585
      ScaleWidth      =   1140
      TabIndex        =   0
      Top             =   0
      Width           =   1200
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
         Height          =   650
         Left            =   110
         Picture         =   "OPASCO04.frx":27EC
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   890
         Width           =   650
      End
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
         Height          =   650
         Left            =   110
         Picture         =   "OPASCO04.frx":2AF6
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Terminar - Salir de la Aplicación"
         Top             =   245
         Width           =   650
      End
      Begin VB.PictureBox MARCOBARRA 
         BackColor       =   &H000000FF&
         Height          =   135
         Left            =   110
         ScaleHeight     =   75
         ScaleWidth      =   585
         TabIndex        =   1
         Top             =   2600
         Width           =   640
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H00FF0000&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   0
            TabIndex        =   2
            Top             =   0
            Width           =   12000
         End
      End
      Begin VB.Image Image1 
         Height          =   390
         Left            =   -310
         Picture         =   "OPASCO04.frx":2C40
         Top             =   2930
         Width           =   1515
      End
   End
   Begin VB.Menu menu 
      Caption         =   "Menú"
      Begin VB.Menu grabar 
         Caption         =   "Grabar"
      End
      Begin VB.Menu salir 
         Caption         =   "Salir"
      End
   End
End
Attribute VB_Name = "OPASCO04"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
'\\\OT-5-572-FG-05/08/05///
    Call CIERRARCH("*.*")
    Hide
'\\\OT-5-572-FG-FIN//
End Sub

Private Sub Command2_Click()
'\\\OT-5-572-FG-05/08/05///
    Call GRACONTROL("", "ASICONTA", TRIM$(Text1))
    IMPRECO$ = "SI"
    If Check1.Value = 1 Then
      IMPRECO$ = "NO"
    End If
    Call GRACONTROL("ASICONTA", "IMPRECOM", IMPRECO$)
    DETAPASE$ = "DENOCUEN"
    If Option2.Value = True Then
      DETAPASE$ = "REFASIEN"
    End If
    Call GRACONTROL("ASICONTA", "DETAPASE", DETAPASE$)
'\\\OT-5-572-FG-FIN///
End Sub

Private Sub Command3_Click()
'\\\OT-5-572-FG-05/08/05///
    Call SELECHO("INDIFRM")
    VDEVU$ = VALACT1$("INDIFRM")
    If VDEVU$ <> "" Then
        Text1 = VDEVU$
        Label3 = VALACT2$("INDIFRM")
    End If
'\\\OT-5-572-FG-FIN///
End Sub

Private Sub Command4_Click(Index As Integer)
   Command4(Index).Enabled = False
   Call TRALOCAL("TAREFASI", "")
   VTB% = VENTABU%("TAREFASI/NC/TITUPAN=Referencias Normalizadas")
   If VTB% >= 0 Then
     Call TRACENTRAL("TAREFASI", "")
   End If
   Call CIERRARCH("TAREFASI")
   Command4(Index).Enabled = True
End Sub

Private Sub Command6_Click(Index As Integer)
'\\\OT-5-572-FG-05/08/05///
    Call CONECRUN("FLFORMS/" + TRIM$(Text1.TEXT), "Formularios de Impresion")
'\\\OT-5-572-FG-FIN///
End Sub

Private Sub Form_Load()
'\\\OT-5-572-FG-05/08/05///
    If CONTROL("NEWASPEC", "CAMBIOAS") = "SI" Then
      Picture16.Visible = True
      Picture16.ZOrder 0
      Me.BackColor = &HFCD7B6
      Me.Frame1.BackColor = &HFCD7B6
      Me.Frame2.BackColor = &HFCD7B6
      Me.Frame4.BackColor = &HFCD7B6
      Me.Option1.BackColor = &HFCD7B6
      Me.Option2.BackColor = &HFCD7B6
      Me.Check1.BackColor = &HFCD7B6
      '
      Set MARCOBARRA.Container = Me.Picture16
      MARCOBARRA.Visible = True
      MARCOBARRA.Top = 3300
      MARCOBARRA.Left = 100
      MARCOBARRA.ZOrder 0

    End If

    Text1 = CONTROL$("", "ASICONTA")
    Label3 = ECOARCH$("INDIFRM", Text1)
    IMPRECO$ = TRIM$(CONTROL$("ASICONTA", "IMPRECOM"))
    If IMPRECO$ = "NO" Then
      Check1.Value = 1
    End If
    If CONTROL("ASICONTA", "DETAPASE") = "REFASIEN" Then
      Option2.Value = True
    End If
'\\\OT-5-572-FG-FIN///
End Sub

Private Sub grabar_Click()
   Call Command2_Click
End Sub

Private Sub salir_Click()
   Call Command1_Click
End Sub

Private Sub Text1_Change()
'\\\OT-5-572-FG-05/08/05///
    DECUE$ = ECOARCH("INDIFRM", TRIM$(Text1))
    If DECUE$ <> "" Then
        Label3 = DECUE$
      Else
        Label3 = ""
    End If
'\\\OT-5-572-FG-FIN///
End Sub

Private Sub Text1_DblClick()
'\\\OT-5-572-FG-05/08/05///
    Call SELECHO("INDIFRM")
    VDEVU$ = VALACT1$("INDIFRM")
    If VDEVU$ <> "" Then
        Text1.TEXT = VDEVU$
        Text2.TEXT = VALACT2$("INDIFRM")
    End If
'\\\OT-5-572-FG-FIN///
End Sub
