VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form OPASCO07 
   BackColor       =   &H0095C5E1&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Formulario de Impresión de Asiento Contable"
   ClientHeight    =   3525
   ClientLeft      =   45
   ClientTop       =   300
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
   Begin VB.PictureBox Picture999 
      Height          =   540
      Index           =   0
      Left            =   2280
      ScaleHeight     =   480
      ScaleWidth      =   1110
      TabIndex        =   16
      Top             =   420
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H0095C5E1&
      Caption         =   "TABLAS"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1170
      Left            =   3675
      TabIndex        =   13
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
         Picture         =   "OPASCO07.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Referencias Normalizadas"
         Top             =   420
         Width           =   570
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H0095C5E1&
      Caption         =   "DETALLE POR PASE"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1170
      Left            =   90
      TabIndex        =   10
      Top             =   2205
      Width           =   3510
      Begin VB.OptionButton Option2 
         Appearance      =   0  'Flat
         BackColor       =   &H0095C5E1&
         Caption         =   "Referencia del Asiento"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   315
         TabIndex        =   12
         Top             =   735
         Width           =   3060
      End
      Begin VB.OptionButton Option1 
         Appearance      =   0  'Flat
         BackColor       =   &H0095C5E1&
         Caption         =   "Denominación Cuenta Contable"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   315
         TabIndex        =   11
         Top             =   420
         Value           =   -1  'True
         Width           =   3060
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H0095C5E1&
      Caption         =   "FORMULARIO DE IMPRESIÓN"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1905
      Left            =   90
      TabIndex        =   3
      Top             =   180
      Width           =   4560
      Begin VB.CheckBox Check1 
         Appearance      =   0  'Flat
         BackColor       =   &H0095C5E1&
         Caption         =   "Suprime Impresión del Comprobante"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   630
         TabIndex        =   9
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
         Picture         =   "OPASCO07.frx":014A
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   420
         Width           =   735
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   2085
         TabIndex        =   7
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
         TabIndex        =   4
         Top             =   450
         Width           =   175
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   180
         TabIndex        =   6
         Top             =   850
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
         TabIndex        =   5
         Top             =   550
         Width           =   2430
      End
   End
   Begin VB.PictureBox Picture7 
      Appearance      =   0  'Flat
      BackColor       =   &H002C3D46&
      ForeColor       =   &H80000008&
      Height          =   3645
      Left            =   4830
      ScaleHeight     =   3615
      ScaleWidth      =   1170
      TabIndex        =   0
      Top             =   0
      Width           =   1200
      Begin VB.PictureBox MARCOBARRA 
         Height          =   195
         Left            =   120
         ScaleHeight     =   135
         ScaleWidth      =   585
         TabIndex        =   18
         Top             =   3120
         Width           =   650
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H0048CC9D&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   -120
            TabIndex        =   19
            Top             =   0
            Width           =   12000
         End
      End
      Begin VB.CommandButton Command14 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   690
         Left            =   110
         Picture         =   "OPASCO07.frx":0454
         Style           =   1  'Graphical
         TabIndex        =   15
         ToolTipText     =   "Tabla de Datos Adicionales de Usuarios"
         Top             =   1680
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
         Height          =   690
         Left            =   110
         Picture         =   "OPASCO07.frx":059E
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   960
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
         Height          =   690
         Left            =   110
         Picture         =   "OPASCO07.frx":08A8
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Terminar - Salir de la Aplicación"
         Top             =   245
         Width           =   650
      End
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   0
      OleObjectBlob   =   "OPASCO07.frx":09F2
      Top             =   1680
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   435
      Index           =   0
      Left            =   60
      OleObjectBlob   =   "OPASCO07.frx":0C26
      TabIndex        =   17
      Top             =   0
      Visible         =   0   'False
      Width           =   1485
   End
End
Attribute VB_Name = "OPASCO07"
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

Private Sub Command14_Click()
    '\\\OT-08-0991-OD-30/10/08///
    If DERACCE%("CONFICON", "Configuración Confidencial") < 2 Then Exit Sub
    '
    Command14.Enabled = False
    Call TRALOCAL("DATUSERS", "")
    '
    N1% = FILEOPEN%("USER.DAT", 9, 64)
    Dim F1 As String * 64
    '
    Y& = 0
    For i& = 1 To LOF(N1%) / 64
       '
       X$ = REGLEIDO$("!DATUSERS", i&)
       If i& > ULTREG&("!DATUSERS") Then X$ = REGBLAN$("DATUSERS")
       I1% = i&
       Call REPLA(X$, Chr$(I1%), 1, 1)
       NOM$ = USERNAM(I1%)
       Call REPLA(X$, UCase$(NOM$), 2, 30)
       If i& > ULTREG&("!DATUSERS") Then
         Call REPLA(X$, Chr$(0), 32, 1)
         Call REPLA(X$, MKI$(0), 33, 2)
       End If
       Call GRAREG("!DATUSERS", X$, i&)
    Next i&
    '
305 Close #N1%: N1% = 0
    Call CIERRARCH("!DATUSERS")
    '
    
10  VTB% = VENTABU%("DATUSERS")
    If VTB% < 0 Then GoTo 99
    J& = 0
    'VALIDACIONES
    X$ = ""
    For U& = 1 To ULTREG&("!DATUSERS")
        X$ = REGLEIDO$("!DATUSERS", U&)
        NIVEL% = Asc(Mid$(X$, 32, 1))
        If NIVEL% < 0 Or NIVEL% > 5 Then
          Call MENSERR(24, "Número de Nivel no Admitido (Máximo 5 )")
          GoTo 10
        End If
    Next U&
    ' tranfiere la inofrmacion de
    Call TRACENTRAL("DATUSERS", "")
    '\\\OT-08-0991-OD-FIN///
99  Command14.Enabled = True
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

Sub Command4_Click(Index As Integer)
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
    Text1 = Control$("", "ASICONTA")
    Label3 = ECOARCH$("INDIFRM", Text1)
    UTILIZA_SKIN% = 1
    Call APLICACIONSKINS(Me)
    IMPRECO$ = TRIM$(Control$("ASICONTA", "IMPRECOM"))
    If IMPRECO$ = "NO" Then
      Check1.Value = 1
    End If
    If Control("ASICONTA", "DETAPASE") = "REFASIEN" Then
      Option2.Value = True
    End If
'\\\OT-5-572-FG-FIN///
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
