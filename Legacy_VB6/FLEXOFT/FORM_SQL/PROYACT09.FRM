VERSION 5.00
Begin VB.Form PROYACT09 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Programacion de Actividades"
   ClientHeight    =   4590
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8445
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4590
   ScaleWidth      =   8445
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00004000&
      Height          =   5300
      Left            =   7140
      ScaleHeight     =   5235
      ScaleWidth      =   1890
      TabIndex        =   28
      Top             =   3120
      Width           =   1950
      Begin VB.CommandButton Command2 
         Height          =   700
         Left            =   80
         Picture         =   "PROYACT09.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   32
         ToolTipText     =   "Guardar con Modificaciones"
         Top             =   105
         Width           =   550
      End
      Begin VB.CommandButton Command7 
         Height          =   700
         Left            =   635
         Picture         =   "PROYACT09.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   31
         ToolTipText     =   "Cierra y Abandona la Aplicación"
         Top             =   105
         Width           =   550
      End
      Begin VB.PictureBox MARCOBARRA 
         BackColor       =   &H000000FF&
         Height          =   135
         Left            =   90
         ScaleHeight     =   75
         ScaleWidth      =   585
         TabIndex        =   29
         Top             =   4200
         Width           =   640
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H00FF0000&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   0
            TabIndex        =   30
            Top             =   105
            Width           =   12000
         End
      End
      Begin VB.Image Image2 
         Height          =   390
         Left            =   -100
         Picture         =   "PROYACT09.frx":058C
         Top             =   900
         Width           =   1515
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Programación"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1440
      Left            =   105
      TabIndex        =   13
      Top             =   3045
      Width           =   6945
      Begin VB.CheckBox Check1 
         Caption         =   "Enviar Confirmacion"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   5355
         TabIndex        =   27
         Top             =   735
         Width           =   1500
      End
      Begin VB.TextBox Text2 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   4305
         TabIndex        =   25
         Top             =   945
         Width           =   825
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   4305
         TabIndex        =   23
         Top             =   630
         Width           =   825
      End
      Begin VB.ComboBox Combo3 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   735
         Style           =   2  'Dropdown List
         TabIndex        =   17
         Top             =   960
         Width           =   2730
      End
      Begin VB.CommandButton Command15 
         Caption         =   "."
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   3.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   5340
         TabIndex        =   16
         Top             =   315
         Width           =   150
      End
      Begin VB.TextBox Text12 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   4305
         TabIndex        =   15
         Top             =   315
         Width           =   1030
      End
      Begin VB.ComboBox Combo2 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   735
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   630
         Width           =   2730
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
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
         Left            =   3570
         TabIndex        =   26
         Top             =   990
         Width           =   735
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
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
         Left            =   3570
         TabIndex        =   24
         Top             =   675
         Width           =   735
      End
      Begin VB.Label Label14 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   735
         TabIndex        =   22
         Top             =   315
         Width           =   2730
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         Caption         =   " F.Proy:"
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
         Left            =   3570
         TabIndex        =   21
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label20 
         Alignment       =   1  'Right Justify
         Caption         =   "Ctrl:"
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
         Left            =   45
         TabIndex        =   20
         Top             =   990
         Width           =   645
      End
      Begin VB.Label Label24 
         Alignment       =   1  'Right Justify
         Caption         =   "Resp:"
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
         Left            =   45
         TabIndex        =   19
         Top             =   675
         Width           =   645
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "Gen:"
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
         Left            =   45
         TabIndex        =   18
         Top             =   360
         Width           =   645
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Identificación"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   2870
      Left            =   105
      TabIndex        =   0
      Top             =   105
      Width           =   8235
      Begin VB.CommandButton COMMAND16 
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
         Height          =   260
         Left            =   5355
         TabIndex        =   34
         ToolTipText     =   "Lista Completa de Contactos"
         Top             =   315
         Width           =   150
      End
      Begin VB.CommandButton Command1 
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
         Height          =   260
         Left            =   7875
         TabIndex        =   33
         ToolTipText     =   "Lista Completa de Contactos"
         Top             =   315
         Width           =   150
      End
      Begin VB.TextBox Text9 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1470
         Left            =   735
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   3
         Top             =   1290
         Width           =   7365
      End
      Begin VB.TextBox Text8 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   280
         Left            =   735
         TabIndex        =   2
         Text            =   " "
         Top             =   945
         Width           =   7365
      End
      Begin VB.ComboBox Combo1 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   5670
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   600
         Width           =   2415
      End
      Begin VB.Label LABEL52 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   735
         TabIndex        =   12
         Top             =   600
         Width           =   4755
      End
      Begin VB.Label Label36 
         Alignment       =   1  'Right Justify
         Caption         =   "Ref:"
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
         Left            =   150
         TabIndex        =   11
         Top             =   990
         Width           =   555
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Caption         =   "Desc:"
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
         Left            =   150
         TabIndex        =   10
         Top             =   1320
         Width           =   555
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         Caption         =   "Prospecto:"
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
         Left            =   2730
         TabIndex        =   9
         Top             =   345
         Width           =   960
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "Nro:"
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
         Left            =   150
         TabIndex        =   8
         Top             =   345
         Width           =   555
      End
      Begin VB.Label Label41 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   735
         TabIndex        =   7
         Top             =   315
         Width           =   990
      End
      Begin VB.Label Label26 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3780
         TabIndex        =   6
         Top             =   315
         Width           =   1575
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         Caption         =   "Oport:"
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
         Left            =   5460
         TabIndex        =   5
         Top             =   345
         Width           =   750
      End
      Begin VB.Label Label10 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   6300
         TabIndex        =   4
         Top             =   315
         Width           =   1575
      End
   End
End
Attribute VB_Name = "PROYACT09"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command5_Click()

End Sub

Private Sub Command1_Click()
    Call SELECHO("!INDOPORT/INDICE DE OPORTUNIDADES")
    CODIOPO$ = TRIM$(VALACT1$("!INDOPORT"))
    If CODIOPO$ <> "" Then
       Label10 = CODIOPO$
    End If
End Sub

Private Sub Command15_Click()
    Call SELECHO("SELFECHA")
    VDEV$ = VALACT1$("SELFECHA")
    If VDEV$ <> "" Then
        Text12.TEXT = VDEV$
    End If
End Sub

Private Sub COMMAND16_Click()
   ABUSX$ = "INDIPROS" ' : If USNBR% < 100 Then ABUSX$ = "BSEGALE"
21 Call SELECHO(ABUSX$ & "/Indice Prospect")
   VDEVU$ = VALACT1$(ABUSX$)
   If VDEVU$ <> "" Then
      Label26 = TRIM$(VDEVU$)
   End If
End Sub

Private Sub Command2_Click()
   '
   Static COTERMI$
   '
   If COTERMI$ = "" Then
      ITERMI$ = IDENTER$
10    COTERMI$ = Control$(ITERMI$, "COTERMI")
      If Len(COTERMI$) <> 2 Then
        Call MENSERR(24, "Falta Identificador de Estacion de Trabajo")
        OPCOTEL.Show 1
        GoTo 10
      End If
   End If
   '
   COCLI$ = TRIM$(Label26)
   CODOP$ = TRIM$(Label10)
   PERSCON$ = Combo1.TEXT
   REFACT$ = TRIM$(Text8)
   FEPROG% = FECHANUM(Text12)
   '
   If COCLI$ = "" Then                    ' importante: dejar esta condicion que si no no aparece en listas
      Call MENSERR(24, "Falta Codigo de Carpeta o Prospecto"): Exit Sub
   End If
   '
   If REFACT$ = "" Then
      Call MENSERR(24, "Falta Referencia"): Exit Sub
   End If
   '
   If FEPROG% < 1 Then
      Call MENSERR(24, "Falta Fecha Proyectada"): Exit Sub
   End If
   '
   If Combo2 = "" Then
      Call MENSERR(24, "Falta Responsable Asignado"): Exit Sub
   End If
   '
   Call FISECOM09.PROGRACTI(COCLI$, CODOP$, PERSCON$, REFACT$, FEPROG%)
   '
   NUACT$ = TRIM$(VALOBADA("ACTIPROY", "MAX(NUME_ACTI)", "NUME_ACTI LIKE '" + COTERMI$ + "%'", "NOMESS"))
   CONDI$ = "NUME_ACTI = '" & NUACT$ & "'"
   Call ACTUREGISTRO("ACTIPROY", "DESC_ACTI", CONDI$, Text9, RAFE&, "NOMESS")         ' DESCRIPCION LARGA DE ACTIVIDAD
     USUASI% = XVALO(Left$(Combo2.TEXT, 3))
   Call ACTUREGISTRO("ACTIPROY", "USUA_ASIG", CONDI$, USUASI%, RAFE&, "NOMESS")   ' USUARIO ASIGNADO
     USUCRL% = XVALO(Left$(Combo3.TEXT, 3))
   Call ACTUREGISTRO("ACTIPROY", "USUA_CTRL", CONDI$, USUCRL%, RAFE&, "NOMESS")   ' USUARIO ASIGNADO
   '
   Call CONACT09.Command1_Click
   Unload Me
   '
End Sub

Private Sub Command7_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   '
   Call CARLIUS(Combo2)
   Call CARLIUS(Combo3)
   '
End Sub

Private Sub Label26_Change()
    '
    Label10 = ""
    LABEL52 = VALOBADA("BSEGCOM", "Raso_Cli", "CODI_CLI = '" & Label26 & "'")
    CONDI$ = "CodOport <>'' AND CODI_CLI = '" & Label26 & "'"
    Call CARGALISEL("!INDOPORT", "OPORTUNI", "CodOport/A12;RefeOport/A48", CONDI$, "NOMESS")
    '
    Call FISECOM09.CARLICONTAC(Combo1, Label26)
    '
End Sub

