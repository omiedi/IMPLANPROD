VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form CALENDAR 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   " Vista 9 días a Partir de Fecha Activa"
   ClientHeight    =   8085
   ClientLeft      =   45
   ClientTop       =   600
   ClientWidth     =   15135
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8085
   ScaleWidth      =   15135
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "< Retroceder de a un día"
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
      Left            =   7800
      TabIndex        =   17
      Top             =   7680
      Width           =   2500
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Avanzar de a un día >"
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
      Left            =   12600
      TabIndex        =   16
      Top             =   7680
      Width           =   2500
   End
   Begin VB.TextBox Text18 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
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
      Left            =   5040
      Locked          =   -1  'True
      TabIndex        =   15
      TabStop         =   0   'False
      Text            =   "Desde:"
      Top             =   7680
      Width           =   675
   End
   Begin VB.TextBox LABEL1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0,00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   11274
         SubFormatType   =   1
      EndProperty
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
      Left            =   5895
      Locked          =   -1  'True
      TabIndex        =   14
      Top             =   7680
      Width           =   1290
   End
   Begin VB.CommandButton Command8 
      Caption         =   "."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   5760
      TabIndex        =   13
      ToolTipText     =   "Click Para Seleccionar Sección"
      Top             =   7680
      Width           =   175
   End
   Begin VB.PictureBox Picture999 
      Height          =   540
      Index           =   0
      Left            =   2400
      ScaleHeight     =   480
      ScaleWidth      =   1110
      TabIndex        =   3
      Top             =   720
      Visible         =   0   'False
      Width           =   1170
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   0
      OleObjectBlob   =   "CALENDAR.frx":0000
      Top             =   1680
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   435
      Index           =   0
      Left            =   2280
      OleObjectBlob   =   "CALENDAR.frx":0234
      TabIndex        =   4
      Top             =   -240
      Visible         =   0   'False
      Width           =   1485
   End
   Begin MSFlexGridLib.MSFlexGrid MSF1 
      Height          =   2295
      Index           =   0
      Left            =   0
      TabIndex        =   5
      Top             =   240
      Width           =   4995
      _ExtentX        =   8811
      _ExtentY        =   4048
      _Version        =   393216
      Cols            =   4
      BackColor       =   16777215
      BackColorFixed  =   16501405
      BackColorBkg    =   16777215
      WordWrap        =   -1  'True
      FocusRect       =   2
      AllowUserResizing=   1
      FormatString    =   "|>CAMPO1|>CAMPO2|>CAMPO3"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSFlexGridLib.MSFlexGrid MSF1 
      Height          =   2295
      Index           =   1
      Left            =   5025
      TabIndex        =   6
      Top             =   240
      Width           =   4995
      _ExtentX        =   8811
      _ExtentY        =   4048
      _Version        =   393216
      Cols            =   4
      BackColor       =   16777215
      BackColorFixed  =   16501405
      BackColorBkg    =   16777215
      WordWrap        =   -1  'True
      FocusRect       =   2
      AllowUserResizing=   1
      FormatString    =   "|>CAMPO1|>CAMPO2|>CAMPO3"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSFlexGridLib.MSFlexGrid MSF1 
      Height          =   2295
      Index           =   2
      Left            =   10080
      TabIndex        =   7
      Top             =   240
      Width           =   4995
      _ExtentX        =   8811
      _ExtentY        =   4048
      _Version        =   393216
      Cols            =   4
      BackColor       =   16777215
      BackColorFixed  =   16501405
      BackColorBkg    =   16777215
      WordWrap        =   -1  'True
      FocusRect       =   2
      AllowUserResizing=   1
      FormatString    =   "|>CAMPO1|>CAMPO2|>CAMPO3"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSFlexGridLib.MSFlexGrid MSF1 
      Height          =   2295
      Index           =   3
      Left            =   0
      TabIndex        =   8
      Top             =   2760
      Width           =   4995
      _ExtentX        =   8811
      _ExtentY        =   4048
      _Version        =   393216
      Cols            =   4
      BackColor       =   16777215
      BackColorFixed  =   16501405
      BackColorBkg    =   16777215
      WordWrap        =   -1  'True
      FocusRect       =   2
      AllowUserResizing=   1
      FormatString    =   "|>CAMPO1|>CAMPO2|>CAMPO3"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSFlexGridLib.MSFlexGrid MSF1 
      Height          =   2295
      Index           =   4
      Left            =   5025
      TabIndex        =   9
      Top             =   2760
      Width           =   4995
      _ExtentX        =   8811
      _ExtentY        =   4048
      _Version        =   393216
      Cols            =   4
      BackColor       =   16777215
      BackColorFixed  =   16501405
      BackColorBkg    =   16777215
      WordWrap        =   -1  'True
      FocusRect       =   2
      AllowUserResizing=   1
      FormatString    =   "|>CAMPO1|>CAMPO2|>CAMPO3"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSFlexGridLib.MSFlexGrid MSF1 
      Height          =   2295
      Index           =   5
      Left            =   10080
      TabIndex        =   10
      Top             =   2760
      Width           =   4995
      _ExtentX        =   8811
      _ExtentY        =   4048
      _Version        =   393216
      Cols            =   4
      BackColor       =   16777215
      BackColorFixed  =   16501405
      BackColorBkg    =   16777215
      WordWrap        =   -1  'True
      FocusRect       =   2
      AllowUserResizing=   1
      FormatString    =   "|>CAMPO1|>CAMPO2|>CAMPO3"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   3375
      Left            =   15240
      TabIndex        =   11
      ToolTipText     =   "Doble Click Editar"
      Top             =   3480
      Width           =   4995
      _ExtentX        =   8811
      _ExtentY        =   5953
      _Version        =   393216
      Cols            =   4
      BackColor       =   16777215
      BackColorFixed  =   16501405
      BackColorBkg    =   16777215
      WordWrap        =   -1  'True
      FocusRect       =   2
      AllowUserResizing=   1
      FormatString    =   "|>CAMPO1|>CAMPO2|>CAMPO3"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSFlexGridLib.MSFlexGrid MSF1 
      Height          =   2295
      Index           =   6
      Left            =   0
      TabIndex        =   24
      Top             =   5280
      Width           =   4995
      _ExtentX        =   8811
      _ExtentY        =   4048
      _Version        =   393216
      Cols            =   4
      BackColor       =   16777215
      BackColorFixed  =   16501405
      BackColorBkg    =   16777215
      WordWrap        =   -1  'True
      FocusRect       =   2
      AllowUserResizing=   1
      FormatString    =   "|>CAMPO1|>CAMPO2|>CAMPO3"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSFlexGridLib.MSFlexGrid MSF1 
      Height          =   2295
      Index           =   7
      Left            =   5040
      TabIndex        =   26
      Top             =   5280
      Width           =   4995
      _ExtentX        =   8811
      _ExtentY        =   4048
      _Version        =   393216
      Cols            =   4
      BackColor       =   16777215
      BackColorFixed  =   16501405
      BackColorBkg    =   16777215
      WordWrap        =   -1  'True
      FocusRect       =   2
      AllowUserResizing=   1
      FormatString    =   "|>CAMPO1|>CAMPO2|>CAMPO3"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSFlexGridLib.MSFlexGrid MSF1 
      Height          =   2295
      Index           =   8
      Left            =   10080
      TabIndex        =   27
      Top             =   5280
      Width           =   4995
      _ExtentX        =   8811
      _ExtentY        =   4048
      _Version        =   393216
      Cols            =   4
      BackColor       =   16777215
      BackColorFixed  =   16501405
      BackColorBkg    =   16777215
      WordWrap        =   -1  'True
      FocusRect       =   2
      AllowUserResizing=   1
      FormatString    =   "|>CAMPO1|>CAMPO2|>CAMPO3"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label LBLDIA 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   8
      Left            =   10080
      TabIndex        =   29
      Top             =   5080
      Width           =   585
   End
   Begin VB.Label LBLDIA 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   7
      Left            =   5040
      TabIndex        =   28
      Top             =   5080
      Width           =   585
   End
   Begin VB.Label LBLDIA 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   6
      Left            =   0
      TabIndex        =   25
      Top             =   5080
      Width           =   585
   End
   Begin VB.Label LBLDIA 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   5
      Left            =   10080
      TabIndex        =   23
      Top             =   2550
      Width           =   585
   End
   Begin VB.Label LBLDIA 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   4
      Left            =   5040
      TabIndex        =   22
      Top             =   2550
      Width           =   585
   End
   Begin VB.Label LBLDIA 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   3
      Left            =   120
      TabIndex        =   21
      Top             =   2550
      Width           =   585
   End
   Begin VB.Label LBLDIA 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   2
      Left            =   10080
      TabIndex        =   20
      Top             =   30
      Width           =   585
   End
   Begin VB.Label LBLDIA 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   1
      Left            =   5040
      TabIndex        =   19
      Top             =   30
      Width           =   585
   End
   Begin VB.Label LBLDIA 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   18
      Top             =   30
      Width           =   585
   End
   Begin VB.Label TEPRUEBA 
      AutoSize        =   -1  'True
      Caption         =   "TEPRUEBA"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   0
      TabIndex        =   12
      Top             =   0
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Label Label5 
      Height          =   435
      Left            =   4200
      TabIndex        =   2
      Top             =   0
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.Label Label2 
      Height          =   435
      Left            =   4830
      TabIndex        =   1
      Top             =   0
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.Label Label17 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404000&
      Height          =   330
      Left            =   7440
      TabIndex        =   0
      Top             =   1140
      Width           =   1695
   End
   Begin VB.Menu archivo 
      Caption         =   "Archivo"
      Begin VB.Menu salir 
         Caption         =   "Salir"
      End
   End
End
Attribute VB_Name = "CALENDAR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub BOTEXIT_Click()
   Unload Me
End Sub
'

Sub BOTREC_Click()
   '
End Sub

Private Sub Command1_Click()
   Command1.Enabled = False
   FE% = FECHANUM(LABEL1)
   If FE% > 0 Then
     FE% = DIANTE(FE%, 1)
     LABEL1 = FECHATEX$(FE%)
   End If
99 Command1.Enabled = True

End Sub

Private Sub Command3_Click()

End Sub

Private Sub Command5_Click()
   Command5.Enabled = False
   FE% = FECHANUM(LABEL1)
   If FE% > 0 Then
     FE% = DIASPOST(FE%, 1)
     LABEL1 = FECHATEX$(FE%)
   End If
99 Command5.Enabled = True

End Sub

Private Sub Command8_Click()
   Command8.Enabled = False
   Call SELECHO("SELFECHA")
   VDEV$ = VALACT1$("SELFECHA")
   If VDEV$ <> "" Then
    Me.LABEL1 = VDEV$
    
   End If
99 Command8.Enabled = True

End Sub


'LA RUTINA SE MOVIO AL FORFASQL
'Sub ASIGNALOTES()
'   'SELECCION DE LOTES A ASIGNAR AL PEDIDO
'   REGIS% = 0
'   ASIGLOT07.List1.Clear
'   ASIGLOT07.List2.Clear
'   For U& = 1 To ULTREG&("!CARDEFAC")
'       XX$ = REGLEIDO$("!CARDEFAC", U&)
'       CCII% = CVI(Mid$(XX$, 5, 2))
'       If CCII% > 0 And ESTRAZABLE%(CCII%) = 1 Then
'         TXTX$ = ""
'         Call REPLA(TXTX$, CODEXT$(CCII%), 1, 16)
'         Call REPLA(TXTX$, DESCRIT$(CCII%), 18, 33)
'         SALPEND = CVS(Mid$(XX$, 1, 4))
'         Call REPLA(TXTX$, FORMATNUM$(SALPEND, "F10.1"), 54, 12)
'         Call REPLA(TXTX$, TRIM$(FORMATNUM$(U&, "I5")), 80, 5)
'         ASIGLOT07.List1.AddItem (TXTX$)
'         REGIS% = REGIS% + 1
'       End If
'   Next U&
'   If REGIS% > 0 Then
'      ASIGLOT07.Show 1
'   End If
'   '
'End Sub

Private Sub Form_Load()
   For i& = 0 To MSF1.Count - 1
      Call CARGA_ENCABEZADO(MSF1(i&), "REFERENCIA/A18;Trabajo/A12;F.Horig/A6;H.Orig/A8", Me)
   Next i&
End Sub

Private Sub Form_Unload(Cancel As Integer)
   For i& = 0 To MSF1.Count - 1
      MSF1(i&).Clear
   Next i&

End Sub

Private Sub label1_Change()
   'COLOCAR LA FECHA INICIAL DONDE SE REALIZA EL DOBLE CLICK
   Call LIMPIARMSFCALENDARIO
   Dim MSF As MSFlexGrid
   FE% = FECHANUM(LABEL1)
   Call ACTULABELDIA(FE%)

   FEACUMU% = FE%
   If FE% < 1 Then Exit Sub
   FECHADES = FETEXCOR1$(FE%)
   FEHAS% = DIASPOST(FE%, 9)
   FECHAHAS = FETEXCOR1$(FEHAS%)
   FEREALIMIN% = FECHANUM("01/02/80")
   FEREALMINTEX$ = FETEXCOR1$(FEREALIMIN%)
'   DSEM$ (FECHA)
   J& = 0
   SQLX$ = "SELECT * FROM AGENDA WITH(NOLOCK) WHERE Fecha_Original >= '" & FECHADES & "' AND  Fecha_Original <= '" & FECHAHAS & "'"
   SQLX$ = SQLX$ + " AND DIA_REALIZADO <= '" & FEREALMINTEX$ & "' ORDER BY Fecha_Original,HORARIO_ORIG"
   Set RS = READSET(SQLX$)
   With RS
   Do While Not .EOF
      FELEIDO% = CVINT(!Fecha_Original)
20    If FELEIDO% = FEACUMU% Then
            If J& >= 9 Then Exit Do
            MSF1(J&).Rows = MSF1(J&).Rows + 1
            NC = XVALO(!nume_cli)
            ORD% = ORD% + 1
            MSF1(J&).TextMatrix(ORD%, 1) = TRIM$(Left$(rasocli$(NC), 24)) + "(" + TRIM$(Left$(DOMICLI$(NC), 24)) + ")"
            MSF1(J&).TextMatrix(ORD%, 2) = SAFETEXT$(!DescTrab)
            MSF1(J&).TextMatrix(ORD%, 3) = FECHATEX$(CVINT(!Fecha_Original))
            HORORIG$ = TRIM$(!Horario_Orig)
            If HORORIG$ = "12:00:00 p.m." Then HORORIG$ = ""
            MSF1(J&).TextMatrix(ORD%, 4) = HORORIG$
      Else
         J& = J& + 1
         FEACUMU% = DIASPOST(FEACUMU%, 1)
         ORD% = 0
         If FEACUMU% > FEHAS% Then GoTo 80
         GoTo 20
      End If
60   .MoveNext
   Loop
   End With
80 RS.Close
   Set RS = Nothing
   'REALIZAR UNA SELECT QUE MUESTRE LA SEMANA A PARTIR DE DICHA FECHA ORDENADA POR FECHA
   'IR CARGANDO EN CADA GRILLA SEGUN CORRESPONDA PASANDO DE UNA A OTRA. CADA VEZ QUE CAMBIE LA FECHA.
   '
End Sub
Sub LIMPIARMSFCALENDARIO()
   For i& = 0 To MSF1.Count - 1
      MSF1(i&).Rows = 1
   Next i&
   
End Sub
Private Sub salir_Click()
   Unload Me
End Sub
Sub ACTULABELDIA(FE1%)
    FE% = FE1%
    FEHOY% = HOY(HO$)
    For i& = 0 To MSF1.Count - 1
       DIANAME$ = DSEM$(FE%)
       'If DIANAME$ = "DOMINGO" Then FE% = DIASPOST(FE%, 1): DIANAME$ = DSEM$(FE%)
       LBLDIA(i&) = DIANAME$ & "  " & FECHATEX$(FE%)
       If FE% = FEHOY% Then
          LBLDIA(i&).ForeColor = vbRed
          MSF1(i&).BackColorFixed = &H8080FF
       Else
          LBLDIA(i&).ForeColor = vbBlack
          MSF1(i&).BackColorFixed = &HFBCA9D
          
       End If
       FE% = DIASPOST(FE%, 1)
    Next i&
End Sub


