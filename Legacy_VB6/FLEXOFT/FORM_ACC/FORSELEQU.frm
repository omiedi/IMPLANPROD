VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form FORSELEQU 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0E0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Fecha y Equipos para Planilla de Carga"
   ClientHeight    =   4515
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   6570
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4515
   ScaleWidth      =   6570
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture999 
      Height          =   435
      Index           =   1
      Left            =   0
      ScaleHeight     =   375
      ScaleWidth      =   1215
      TabIndex        =   13
      Top             =   0
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00004000&
      Height          =   4575
      Left            =   5670
      ScaleHeight     =   4515
      ScaleWidth      =   1035
      TabIndex        =   3
      Top             =   0
      Width           =   1095
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
         Left            =   120
         Picture         =   "FORSELEQU.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Refresca - Invierte Selección"
         Top             =   1480
         Width           =   650
      End
      Begin VB.CommandButton BOTREC 
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
         Left            =   120
         Picture         =   "FORSELEQU.frx":014A
         Style           =   1  'Graphical
         TabIndex        =   0
         ToolTipText     =   "Aprueba la Pre-Selección"
         Top             =   800
         Width           =   650
      End
      Begin VB.CommandButton BOTEXIT 
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
         Left            =   120
         Picture         =   "FORSELEQU.frx":0454
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Cancela la Pre-Selección"
         Top             =   120
         Width           =   650
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00E0E0E0&
      Height          =   615
      Left            =   105
      ScaleHeight     =   555
      ScaleWidth      =   5625
      TabIndex        =   2
      Top             =   3720
      Width           =   5685
      Begin VB.CommandButton Command3 
         Caption         =   "."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   4620
         TabIndex        =   11
         Top             =   105
         Width           =   180
      End
      Begin VB.CommandButton Command2 
         Caption         =   "."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   2205
         TabIndex        =   10
         Top             =   105
         Width           =   180
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
         Height          =   300
         Left            =   3675
         TabIndex        =   8
         Text            =   " "
         Top             =   105
         Width           =   960
      End
      Begin VB.TextBox Text1 
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
         Left            =   1260
         TabIndex        =   6
         Text            =   " "
         Top             =   105
         Width           =   960
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Hasta"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   3045
         TabIndex        =   9
         Top             =   210
         Width           =   540
      End
      Begin VB.Label Label33 
         BackStyle       =   0  'Transparent
         Caption         =   "Desde"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   630
         TabIndex        =   7
         Top             =   210
         Width           =   645
      End
   End
   Begin VB.ListBox LIMATSEL 
      BeginProperty Font 
         Name            =   "Fixedsys"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3630
      Left            =   105
      Style           =   1  'Checkbox
      TabIndex        =   1
      Top             =   105
      Width           =   5580
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   4020
      OleObjectBlob   =   "FORSELEQU.frx":075E
      Top             =   2100
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   435
      Index           =   0
      Left            =   4080
      OleObjectBlob   =   "FORSELEQU.frx":0992
      TabIndex        =   14
      Top             =   420
      Visible         =   0   'False
      Width           =   1485
   End
   Begin VB.Label Label2 
      Height          =   435
      Left            =   210
      TabIndex        =   12
      Top             =   4515
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   5775
      Y1              =   4335
      Y2              =   4335
   End
End
Attribute VB_Name = "FORSELEQU"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub BOTEXIT_Click()
    LIMATSEL.Clear
    Label2 = ""
    Unload Me
End Sub
'

Private Sub BOTREC_Click()
    FEC1& = FECHANUM(Text1.TEXT)
    FEC2& = FECHANUM(Text2.TEXT)
    '
    If (FEC1& = 0) Or (FEC2& = 0) Then
       Call MENSERR(24, "Fecha(s) no Válidas.")
       Exit Sub
    End If
    '
    If (FEC2& < FEC1&) Then
       Call MENSERR(24, "Fecha Final Anterior a Inicial.")
       Exit Sub
    End If
    '
    Label2 = "OK"
    Hide
    '
End Sub

Private Sub Command1_Click()
   Command1.Enabled = False
   TPI = LIMATSEL.TopIndex
   For U& = 1 To LIMATSEL.ListCount
     If LIMATSEL.Selected(U& - 1) = True Then
         LIMATSEL.Selected(U& - 1) = False
       Else
         LIMATSEL.Selected(U& - 1) = True
     End If
   Next U&
   On Error Resume Next
   LIMATSEL.TopIndex = TPI
   On Error GoTo 0
   DoEvents
   Command1.Enabled = True
End Sub

Private Sub Command2_Click()
   '\\\OT-06-0369-MH-05/09/06///
   Call SELECHO("SELFECHA")
   VDEVU$ = TRIM$(VALACT1$("SELFECHA"))
   If VDEVU$ <> "" Then
      Text1.TEXT = VALACT1$("SELFECHA")
   End If
End Sub

Private Sub Command3_Click()
   Call SELECHO("SELFECHA")
   VDEVU$ = TRIM$(VALACT1$("SELFECHA"))
   If VDEVU$ <> "" Then
      Text2.TEXT = VALACT1$("SELFECHA")
   End If
'\\\OT-06-0369-MH-FIN///
End Sub

Private Sub Form_Load()
    '\\\OT-06-0369-MH-05/09/06///
    For I& = 1 To ULTREG&("PADMAQ")
      F0$ = REGLEIDO$("PADMAQ", I&)
      CODIGO$ = Mid$(F0$, 1, 6)
      DESCRI$ = Mid$(F0$, 7, 30)
      LIMATSEL.AddItem (CODIGO$ + "  " + DESCRI$)
    Next I&
    '\\\OT-06-0369-MH-06/09/06///
    FHOY% = HOY(HO$)
    Text1.TEXT = FECHATEX(FHOY%)
    CADIAS% = Val(CONTROL$("", "PERDIA"))
    FECHA2% = FHOY%
    '
    For I& = 1 To CADIAS%
        FECHA2% = DIPOST(FECHA2%)
    Next I&
    Text2.TEXT = FECHATEX$(FECHA2%)
    Call APLICACIONSKINS(Me, Picture2)
    '\\\OT-06-0369-MH-FIN///
End Sub

Private Sub Text1_Change()
     '\\\OT-06-0369-MH-06/09/06///
    FECHA1% = FECHANUM(Text1.TEXT)
    CADIAS% = Val(CONTROL$("", "PERDIA"))
    
    'If FECHANUM(Text1.TEXT) <> 0 Then
           FECHA2% = FECHA1%
              For I& = 1 To CADIAS%
                 FECHA2% = DIPOST(FECHA2%)
              Next I&
           Text2.TEXT = FECHATEX$(FECHA2%)
        'Else
        '    Call MENSERR(24, "Por favor, ingrese una fecha válida.")
    'End If
    '\\\OT-06-0369-MH-FIN///
End Sub






