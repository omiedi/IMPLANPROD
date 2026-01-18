VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form ESTOSETU 
   Appearance      =   0  'Flat
   BackColor       =   &H00E2D3C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Configuración Para Cálculo Automático"
   ClientHeight    =   6195
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   5400
   ClipControls    =   0   'False
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   13.5
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6195
   ScaleWidth      =   5400
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      BackColor       =   &H00E2D3C0&
      Caption         =   " Mail de Destinatarios de Alerta Stocks Sub-Mínimos"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2160
      Left            =   0
      TabIndex        =   21
      Top             =   3960
      Width           =   4485
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   4
         Left            =   120
         TabIndex        =   26
         Top             =   1680
         Width           =   4215
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   3
         Left            =   120
         TabIndex        =   25
         Top             =   1320
         Width           =   4215
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   2
         Left            =   120
         TabIndex        =   24
         Top             =   960
         Width           =   4215
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   1
         Left            =   120
         TabIndex        =   23
         Top             =   600
         Width           =   4215
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   0
         Left            =   120
         TabIndex        =   22
         Top             =   240
         Width           =   4215
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E2D3C0&
      Caption         =   "Considera Hasta"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   960
      Left            =   0
      TabIndex        =   15
      Top             =   120
      Width           =   4485
      Begin VB.CommandButton Command5 
         Caption         =   "."
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
         Left            =   3260
         TabIndex        =   16
         Top             =   420
         Width           =   175
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Mes y Año:"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   -15
         TabIndex        =   18
         Top             =   480
         Width           =   1065
      End
      Begin VB.Label MESAN 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
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
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   1260
         TabIndex        =   17
         Top             =   420
         Width           =   1995
      End
   End
   Begin VB.TextBox TXTRUTA 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   13
      Top             =   2280
      Width           =   4035
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   4800
      Top             =   3120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00E2D3C0&
      Caption         =   "Generar Archivo de Sub-Minimos en:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1005
      Left            =   0
      TabIndex        =   11
      Top             =   2040
      Width           =   4485
      Begin VB.CommandButton Command1 
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
         Height          =   660
         Left            =   4170
         TabIndex        =   12
         Top             =   240
         Width           =   250
      End
   End
   Begin VB.PictureBox Picture999 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Index           =   0
      Left            =   2415
      ScaleHeight     =   375
      ScaleWidth      =   1110
      TabIndex        =   10
      Top             =   105
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.PictureBox Picture14 
      BackColor       =   &H00808080&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6240
      Left            =   4560
      ScaleHeight     =   6180
      ScaleWidth      =   870
      TabIndex        =   7
      Top             =   0
      Width           =   930
      Begin VB.PictureBox MARCOBARRA 
         Appearance      =   0  'Flat
         BackColor       =   &H002B9973&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   200
         Left            =   70
         ScaleHeight     =   165
         ScaleWidth      =   615
         TabIndex        =   19
         Top             =   5280
         Width           =   650
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H0048CC9D&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   500
            Left            =   0
            TabIndex        =   20
            Top             =   0
            Width           =   12000
         End
      End
      Begin VB.CommandButton Command3 
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
         Height          =   650
         Left            =   90
         Picture         =   "ESTOSETU.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Inicia Proceso de Cálculo Estadístico"
         Top             =   4560
         Width           =   650
      End
      Begin VB.CommandButton Command2 
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
         Height          =   650
         Left            =   90
         Picture         =   "ESTOSETU.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   8
         ToolTipText     =   "Cierra y Abandona la Aplicación"
         Top             =   120
         Width           =   650
      End
      Begin VB.Image Image2 
         Height          =   390
         Left            =   -330
         Picture         =   "ESTOSETU.frx":0454
         ToolTipText     =   "Acceso Directo a Otras Aplicaciones"
         Top             =   5670
         Width           =   1515
      End
   End
   Begin VB.Frame Frame9 
      BackColor       =   &H00E2D3C0&
      Caption         =   "Meses"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   0
      TabIndex        =   1
      Top             =   1125
      Width           =   4485
      Begin VB.OptionButton Option3 
         BackColor       =   &H00E2D3C0&
         Caption         =   "24"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   3
         Left            =   2520
         TabIndex        =   5
         Top             =   400
         Width           =   600
      End
      Begin VB.OptionButton Option3 
         BackColor       =   &H00E2D3C0&
         Caption         =   "12"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   2
         Left            =   1680
         TabIndex        =   4
         Top             =   400
         Width           =   600
      End
      Begin VB.OptionButton Option3 
         BackColor       =   &H00E2D3C0&
         Caption         =   "8"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   1
         Left            =   945
         TabIndex        =   3
         Top             =   400
         Value           =   -1  'True
         Width           =   600
      End
      Begin VB.OptionButton Option3 
         BackColor       =   &H00E2D3C0&
         Caption         =   "6"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   0
         Left            =   210
         TabIndex        =   2
         Top             =   400
         Width           =   600
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   450
         Left            =   3855
         TabIndex        =   6
         Top             =   255
         Width           =   540
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E2D3C0&
      Caption         =   "Seleccionar Que Días Enviar Mail de Alerta"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   840
      Left            =   0
      TabIndex        =   0
      Top             =   3120
      Width           =   4485
      Begin VB.ComboBox Combo1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         ItemData        =   "ESTOSETU.frx":2376
         Left            =   240
         List            =   "ESTOSETU.frx":2378
         Sorted          =   -1  'True
         TabIndex        =   14
         Text            =   "Combo1"
         Top             =   360
         Width           =   3015
      End
   End
End
Attribute VB_Name = "ESTOSETU"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
   Command1.Enabled = False
   RESUL$ = TRIM$(SELECCION_Archivo$("XLS", CommonDialog1))
   If RESUL$ <> "" Then TXTRUTA = RESUL$
   Screen.MousePointer = 1
   Command1.Enabled = True
End Sub

Private Sub Command2_Click()
   Label1 = ""
   Unload Me
End Sub

Private Sub Command3_Click()
   Command3.Enabled = False
   
   If TRIM$(MESAN) = "" Then
      Call COMUNI("Falta Ingresar Mes Y Año")
      GoTo 99
   End If
   If XVALO(Label1) < 1 Then
      Call COMUNI("Falta Seleccionar Cantidad de Meses")
      GoTo 99
   End If
   '
      '
   MAIL1$ = Me.Text1(0): If InStr(MAIL1$, "@") < 1 And TRIM(Text1(0)) <> "" Then Call COMUNI("Mail No Válido"): Call PINTATEXT(Text1(0)): GoTo 99
   MAIL2$ = Me.Text1(1): If InStr(MAIL2$, "@") < 1 And TRIM(Text1(1)) <> "" Then Call COMUNI("Mail No Válido"): Call PINTATEXT(Text1(1)): GoTo 99
   MAIL3$ = Me.Text1(2): If InStr(MAIL3$, "@") < 1 And TRIM(Text1(2)) <> "" Then Call COMUNI("Mail No Válido"): Call PINTATEXT(Text1(2)): GoTo 99
   MAIL4$ = Me.Text1(3): If InStr(MAIL4$, "@") < 1 And TRIM(Text1(3)) <> "" Then Call COMUNI("Mail No Válido"): Call PINTATEXT(Text1(3)): GoTo 99
   MAIL5$ = Me.Text1(4): If InStr(MAIL5$, "@") < 1 And TRIM(Text1(4)) <> "" Then Call COMUNI("Mail No Válido"): Call PINTATEXT(Text1(4)): GoTo 99
   

   Call GRACONTROL("ESTOCK07", "MESANOSE", MESAN.Caption)
   Call GRACONTROL("ESTOCK07", "CANTMESE", Label1.Caption)
   Call GRACONTROL("ESTOCK07", "RUTASUMI", TXTRUTA)
   Call GRACONTROL("ESTOCK07", "DIAPROXE", Combo1.TEXT)
   '
   Call GRACONTROL("ESTOCK07", "MAIL1", Me.Text1(0))
   Call GRACONTROL("ESTOCK07", "MAIL2", Me.Text1(1))
   Call GRACONTROL("ESTOCK07", "MAIL3", Me.Text1(2))
   Call GRACONTROL("ESTOCK07", "MAIL4", Me.Text1(3))
   Call GRACONTROL("ESTOCK07", "MAIL5", Me.Text1(4))
   Call generarArchivotxt

   Hide
99 Command3.Enabled = True
End Sub

Private Sub Command4_Click()

End Sub

Private Sub Command5_Click()
   Command5.Enabled = False
   Call CARSELMESA
   Call SELECHO("SELMESA")
   MESASEL$ = VALACT1$("SELMESA")
   If TRIM$(MESASEL$) <> "" Then
      MESAN = MESASEL$
      MESAN.Refresh
   End If
   Command5.Enabled = True
End Sub

Private Sub Form_Load()
   Call APLICACIONSKINS(Me)
  
   Call CARSELMESA
   MESAN.Caption = CONTROL("ESTOCK07", "MESANOSE")
   Label1.Caption = CONTROL("ESTOCK07", "CANTMESE")
   Select Case XVALO(Label1.Caption)
    Case 6
      Option3(0).Value = True
    Case 8
      Option3(1).Value = True
    Case 12
      Option3(2).Value = True
    Case 24
      Option3(3).Value = True
   End Select
   If XVALO(Label1.Caption) < 1 Then Option3(1).Value = True
   TXTRUTA = CONTROL("ESTOCK07", "RUTASUMI")
   Combo1.Clear
   Combo1.AddItem "DOMINGO"
   Combo1.AddItem "LUNES"
   Combo1.AddItem "MARTES"
   Combo1.AddItem "MIERCOLES"
   Combo1.AddItem "JUEVES"
   Combo1.AddItem "VIERNES"
   Combo1.AddItem "SABADO"
   Combo1.TEXT = CONTROL("ESTOCK07", "DIAPROXE")
   
   Me.Text1(0) = CONTROL("ESTOCK07", "MAIL1")
   Me.Text1(1) = CONTROL("ESTOCK07", "MAIL2")
   Me.Text1(2) = CONTROL("ESTOCK07", "MAIL3")
   Me.Text1(3) = CONTROL("ESTOCK07", "MAIL4")
   Me.Text1(4) = CONTROL("ESTOCK07", "MAIL5")
   '

End Sub
Sub generarArchivotxt()
    '
    RUTAX$ = TRIM$(CONTROL("RETRASLA", "RCENTRAL")) 'configurado en setup de RETRASLA09
    On Error Resume Next
    Kill RUTAX$ + "\DATOENVMAIL.TXT"
    NARCHI2% = FILEOPEN%(RUTAX$ + "\DATOENVMAIL.TXT", 2, 128) ' ABRE EL ARCHIVO DE ESCRITURA
    Print #NARCHI2%, TXTRUTA.TEXT
    For i& = 0 To Text1.UBound
      If Text1(i&) <> "" Then
         If InStr(Text1(i&), "@") > 0 Then
            Print #NARCHI2%, Text1(i&)
         End If
      End If
    Next i&
    '
    Close #NARCHI2%
    On Error GoTo 0


End Sub

Private Sub Form_Unload(Cancel As Integer)
   Label1 = ""
End Sub

Private Sub GRID_Click()

End Sub

Private Sub GRID_DblClick()
   
End Sub

Private Sub Label4_Click()

End Sub

Private Sub Label5_Click()

End Sub

Private Sub Label6_Click()

End Sub

Private Sub Label7_Click()

End Sub

Private Sub Label3_Click(Index As Integer)

End Sub

Private Sub MESAN_DblClick()
   Call CARSELMESA
   Call SELECHO("SELMESA")
   MESASEL$ = VALACT1$("SELMESA")
   If TRIM$(MESASEL$) <> "" Then
      MESAN = MESASEL$
      MESAN.Refresh
   End If
End Sub

Private Sub Option3_Click(Index As Integer)
   For KI% = 0 To 3
     If Option3(KI%).Value = True Then
       Label1 = Option3(Index).Caption
     End If
   Next KI%
End Sub

Private Sub Text2_Change()

End Sub

