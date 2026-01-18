VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form FSTPRCF_EYM 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Configuración de Formulario Registro Control Final"
   ClientHeight    =   6660
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   6375
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
   ScaleHeight     =   6660
   ScaleWidth      =   6375
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00404040&
      Height          =   6690
      Left            =   5460
      ScaleHeight     =   6630
      ScaleWidth      =   1110
      TabIndex        =   4
      Top             =   0
      Width           =   1170
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
         Height          =   500
         Left            =   105
         Picture         =   "FSTPRCF_EYM.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   630
         Width           =   645
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
         Height          =   500
         Left            =   105
         Picture         =   "FSTPRCF_EYM.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   105
         Width           =   645
      End
      Begin VB.Image Image1 
         Height          =   390
         Left            =   -300
         Picture         =   "FSTPRCF_EYM.frx":0454
         ToolTipText     =   "Acceso Directo a Otras Aplicaciones"
         Top             =   6120
         Width           =   1515
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Presentación de Costo y Precios"
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
      TabIndex        =   0
      Top             =   7770
      Width           =   5160
      Begin VB.OptionButton Option3 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Cos.Repos. Calculado"
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
         Left            =   1920
         TabIndex        =   3
         Top             =   360
         Width           =   1305
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Precio Base (C.Rep. + 15 %)"
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
         Left            =   3360
         TabIndex        =   2
         Top             =   360
         Width           =   1695
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Cos.Repos. Informado"
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
         Left            =   360
         TabIndex        =   1
         Top             =   360
         Value           =   -1  'True
         Width           =   1905
      End
   End
   Begin MSComDlg.CommonDialog cmd 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame4 
      Height          =   6495
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   5175
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   765
         Left            =   240
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   18
         Text            =   "FSTPRCF_EYM.frx":2376
         Top             =   5520
         Width           =   4680
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   765
         Left            =   240
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   16
         Text            =   "FSTPRCF_EYM.frx":2378
         Top             =   4440
         Width           =   4680
      End
      Begin VB.TextBox Text5 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   765
         Left            =   240
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   14
         Text            =   "FSTPRCF_EYM.frx":237A
         Top             =   3360
         Width           =   4680
      End
      Begin VB.TextBox Text4 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   765
         Left            =   240
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   12
         Text            =   "FSTPRCF_EYM.frx":237C
         Top             =   2280
         Width           =   4680
      End
      Begin VB.TextBox Text7 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   765
         Left            =   240
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   9
         Text            =   "FSTPRCF_EYM.frx":237E
         Top             =   1200
         Width           =   4680
      End
      Begin VB.ComboBox Combo1 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   2040
         TabIndex        =   8
         Text            =   "Combo1"
         Top             =   360
         Width           =   2865
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Ruta Archivo Excele Equipo C15"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   5
         Left            =   -120
         TabIndex        =   19
         Top             =   5280
         Width           =   5085
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Ruta Archivo Excel Etiqueta  De Embalaje C15"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   4
         Left            =   -120
         TabIndex        =   17
         Top             =   4200
         Width           =   5085
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Ruta de Archivo Excel de Control Final C15"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   2
         Left            =   -120
         TabIndex        =   15
         Top             =   3120
         Width           =   5085
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Ruta de Arch. Excel Control Final Asist.Técnica"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   1
         Left            =   -120
         TabIndex        =   13
         Top             =   2040
         Width           =   5085
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Ruta de Archivo Excel de Control Final"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   3
         Left            =   -480
         TabIndex        =   11
         Top             =   960
         Width           =   5085
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Impresora de Registro de Datos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   0
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   1845
      End
   End
End
Attribute VB_Name = "FSTPRCF_EYM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
    Call CIERRARCH("*.*")
    Hide
End Sub

Private Sub Command2_Click()
    '
    '
    Call GRACONTROL(IDENTER$, "PRIREGCF", Combo1.TEXT)
    Call GRACONTROL("*", "PRIREGCV", Combo1.TEXT)
    

    Call GRACONTROL("REGICFIN", "ARCHCFIN", Text7.TEXT)
    Call GRACONTROL$("REPARACI", "CTFIREPA", Text4.TEXT)
    Call GRACONTROL$("REGICFIN", "CF_EXC15", Text5.TEXT)
    
    Call GRACONTROL$("ETIQUC15", "ETEMBC15", Text1.TEXT)
    Call GRACONTROL$("ETIQUC15", "ETEQUC15", Text2.TEXT)
    
    Call CIERRARCH("*.*")
    Call COMUNI("Re-Configuración Completada.\Re-Inicie Ahora la Aplicación '" + UCase$(App.EXEName) + "'.")
    Call FINAL("")
    '
End Sub



Private Sub Form_Load()
    '
    Call TRANSLABELS(Name)
    '
    Text7.TEXT = CONTROL("REGICFIN", "ARCHCFIN")
    Text4.TEXT = CONTROL$("REPARACI", "CTFIREPA")
    Text5.TEXT = CONTROL$("REGICFIN", "CF_EXC15")
    
    Text1.TEXT = CONTROL$("ETIQUC15", "ETEMBC15")
    Text2.TEXT = CONTROL$("ETIQUC15", "ETEQUC15")

    PDES$ = CONTROL$(IDENTER$, "PRIREGCF")
    If PDES$ = "" Then PDES$ = CONTROL$("*", "PRIREGCF")

    Combo1.Clear
    Dim PRX As Printer
    On Error Resume Next
    Combo1.TEXT = Printer.DeviceName
    On Error GoTo 0
    For Each PRX In Printers
      Combo1.AddItem PRX.DeviceName
      If PRX.DeviceName = PDES$ Then
        Combo1.TEXT = PDES$
      End If
    Next

    '
End Sub

Sub MEDICIONES()
     '
     If DERACCE%("MEDICION", "Carga Valores Admisibles de Medicion") < 2 Then Exit Sub
     '
10   SQLX$ = "SELECT * FROM CF_MEDICION WITH(NOLOCK)"
     SQLX$ = SQLX$ + " WHERE  MARBORRA = 0 "
     SQLX$ = SQLX$ + " ORDER BY CODEXTE"
     Dim MEDTMP As ADODB.Recordset
     Set MEDTMP = READSET(SQLX$)
     '
     Call SETULTREG("!MEDICION", 0)
     JJ& = 0
     With MEDTMP
        Do While (.EOF = False)
           X$ = REGBLAN$("MEDICION")
           Call REPLA(X$, MKI(!COD_INTE), 1, 2)
           Call REPLA(X$, MKD(!MINIMO_1), 3, 8)
           Call REPLA(X$, MKD(!MAXIMO_1), 11, 8)
           Call REPLA(X$, MKD(!MINIMO_2), 19, 8)
           Call REPLA(X$, MKD(!MAXIMO_2), 27, 8)
           Call REPLA(X$, MKD(!MINIMO_3), 35, 8)
           Call REPLA(X$, MKD(!MAXIMO_3), 43, 8)
           JJ& = JJ& + 1
           Call GRAREG("!MEDICION", X$, JJ&)
            .MoveNext
         Loop
      End With
      MEDTMP.Close
      Set MEDTMP = Nothing

20    VTB% = VENTABU%("MEDICION/NC/TITUPAN=Tabla de Valores Admisibles de Medicion")

      If VTB% < 0 Then Exit Sub
      '
      CONDI$ = "COD_INTE > 0 "
      Call ACTUREGISTRO("CF_MEDICION", "MARBORRA", CONDI$, 1, REG&)
         
      SQLX$ = "SELECT * FROM CF_MEDICION "
      SQLX$ = SQLX$ + " WHERE COD_INTE < 1 "
      Set MEDITMP = New ADODB.Recordset
      MEDITMP.Open FILTSQL$(SQLX$), FLEXCONN, adOpenKeyset, adLockPessimistic, adCmdText
           With MEDITMP
             For i& = 1 To ULTREG&("!MEDICION")
               XX$ = REGLEIDO$("!MEDICION", i&)
               CI% = CVI(Mid$(XX$, 1, 2))
               .AddNew   ' YA CONVERTIDO
               On Error Resume Next
               !CodiEmpr = CodiEmp%
               On Error GoTo 0
               !COD_INTE = CI%
               !CODEXTE = CODEXT$(CI%)
               !MINIMO_1 = CVD(Mid$(XX$, 3, 8))
               !MAXIMO_1 = CVD(Mid$(XX$, 11, 8))
               !MINIMO_2 = CVD(Mid$(XX$, 19, 8))
               !MAXIMO_2 = CVD(Mid$(XX$, 27, 8))
               !MINIMO_3 = CVD(Mid$(XX$, 35, 8))
               !MAXIMO_3 = CVD(Mid$(XX$, 43, 8))
               !MARBORRA = 0
               .Update
             Next i&
           End With
         MEDITMP.Close
         Set MEDITMP = Nothing
         '
         CONDI$ = "MARBORRA = 1"
         Call BORRAREGISTROS("CF_MEDICION", CONDI$, REG&)
         '
         SQLX$ = "SELECT COUNT(*) , COD_INTE FROM CF_MEDICION WITH(NOLOCK) WHERE MARBORRA = 0 GROUP BY COD_INTE HAVING COUNT(*) > 1"
         Set MEDTMP = READSET(SQLX$)
         CI% = XVALO(MEDTMP!COD_INTE)
         If CI% > 0 Then
           Call COMUNI("Código " & CODEXT$(CI%) & " Repetido.")
           GoTo 10
         End If
         '
         Call COMUNI("Valores Admisibles Guardados")
         '
End Sub


Private Sub Text1_DblClick()
   RUTAX$ = TRIM$(SELECCION_Archivo$("*.*", cmd))
   If RUTAX$ <> "" Then Text1 = RUTAX$
   Screen.MousePointer = 1
End Sub


Private Sub Text2_DblClick()
   RUTAX$ = TRIM$(SELECCION_Archivo$("*.*", cmd))
   If RUTAX$ <> "" Then Text2 = RUTAX$
   Screen.MousePointer = 1
End Sub

Private Sub Text4_DblClick()
   RUTAX$ = TRIM$(SELECCION_Archivo$("*.*", cmd))
   If RUTAX$ <> "" Then Text4 = RUTAX$
   Screen.MousePointer = 1

End Sub

Private Sub Text5_DblClick()
   RUTAX$ = TRIM$(SELECCION_Archivo$("*.*", cmd))
   If RUTAX$ <> "" Then Text5 = RUTAX$
   Screen.MousePointer = 1
End Sub

Private Sub Text7_DblClick()
   RUTAX$ = TRIM$(SELECCION_Archivo$("*.*", cmd))
   If RUTAX$ <> "" Then Text7 = RUTAX$
   Screen.MousePointer = 1

End Sub
