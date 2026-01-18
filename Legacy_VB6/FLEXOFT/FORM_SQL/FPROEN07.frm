VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form FPROEN07 
   BackColor       =   &H00C0D0D0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Fechas Proyectadas de Entrega"
   ClientHeight    =   2670
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   4125
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2670
   ScaleWidth      =   4125
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Picture999 
      Height          =   435
      Index           =   1
      Left            =   1440
      ScaleHeight     =   375
      ScaleWidth      =   1215
      TabIndex        =   12
      Top             =   0
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      ForeColor       =   &H80000008&
      Height          =   3585
      Left            =   3240
      ScaleHeight     =   3555
      ScaleWidth      =   1035
      TabIndex        =   9
      Top             =   0
      Width           =   1065
      Begin VB.PictureBox MARCOBARRA 
         Appearance      =   0  'Flat
         BackColor       =   &H002B9973&
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   105
         ScaleHeight     =   165
         ScaleWidth      =   615
         TabIndex        =   14
         Top             =   2280
         Width           =   650
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H0048CC9D&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   0
            TabIndex        =   15
            Top             =   0
            Width           =   12000
         End
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
         Height          =   635
         Left            =   105
         Picture         =   "FPROEN07.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Genera y Emite Reporte"
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
         Height          =   635
         Left            =   105
         Picture         =   "FPROEN07.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Terminar - Salir de la Aplicación"
         Top             =   270
         Width           =   650
      End
   End
   Begin VB.CommandButton Command29 
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
      Left            =   630
      Picture         =   "FPROEN07.frx":0454
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Ayuda Flexoft en Línea"
      Top             =   3990
      Width           =   750
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0D0D0&
      Caption         =   "SALIDA"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   750
      Left            =   120
      TabIndex        =   5
      Top             =   1800
      Width           =   2955
      Begin VB.OptionButton Option2 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0D0D0&
         Caption         =   "Impresora"
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
         Height          =   330
         Index           =   1
         Left            =   1575
         TabIndex        =   7
         Top             =   315
         Width           =   1245
      End
      Begin VB.OptionButton Option2 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0D0D0&
         Caption         =   "Pantalla"
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
         Height          =   330
         Index           =   0
         Left            =   420
         TabIndex        =   6
         Top             =   315
         Value           =   -1  'True
         Width           =   2500
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0D0D0&
      Caption         =   "OPCIONES"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1590
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2955
      Begin VB.OptionButton Option1 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0D0D0&
         Caption         =   "Por Número de O/F"
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
         Height          =   330
         Index           =   3
         Left            =   420
         TabIndex        =   4
         Top             =   1170
         Width           =   2500
      End
      Begin VB.OptionButton Option1 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0D0D0&
         Caption         =   "Por Producto / Artículo"
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
         Height          =   330
         Index           =   2
         Left            =   420
         TabIndex        =   3
         Top             =   885
         Width           =   2500
      End
      Begin VB.OptionButton Option1 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0D0D0&
         Caption         =   "Por Destino / Cliente"
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
         Height          =   330
         Index           =   1
         Left            =   420
         TabIndex        =   2
         Top             =   600
         Width           =   2500
      End
      Begin VB.OptionButton Option1 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0D0D0&
         Caption         =   "General"
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
         Height          =   330
         Index           =   0
         Left            =   420
         TabIndex        =   1
         Top             =   315
         Value           =   -1  'True
         Width           =   2500
      End
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   4020
      OleObjectBlob   =   "FPROEN07.frx":075E
      Top             =   2100
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   435
      Index           =   0
      Left            =   4080
      OleObjectBlob   =   "FPROEN07.frx":0992
      TabIndex        =   13
      Top             =   420
      Visible         =   0   'False
      Width           =   1485
   End
End
Attribute VB_Name = "FPROEN07"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Hide
End Sub

Private Sub Command2_Click()
    '
    CONSIMULACION% = 0
    If EXISTE_TABLA("SIMU_ORDEFABR") > 0 Then
       If CantRegistros("SIMU_ORDEFABR", , "NOMESS") > 0 Then
         OPX% = COMALTER%("Opciones de Trabajo\\Listado Cálculo Normal\Listado Considerando Simulación")
         If OPX% > 1 Then
           CONSIMULACION% = XVALO(CONTROL("CARMAQUI", "ULTISIMU"))
           If CONSIMULACION% < 1 Then
              Call COMUNI("Atención!! No Hay Cálculos con Simulación de Fabricación\Se Visualizará Cálculo Normal")
           End If
         End If
       End If
    End If
    ARCHORDEFABR$ = "ORDEFABR"
    If CONSIMULACION% < 1 Then
       Call GENORDEFDAT(1)
    Else
       ARCHORDEFABR$ = "SRDEFABP"
       SIMUCARM.CONSIMULACION = 1
       Call SIMUCARM.GENORDEFDAT_SIMU(1)
    End If

    
    Screen.MousePointer = 11
    Call BLANARCH("ORDESOR")
    For U& = 1 To ULTREG&(ARCHORDEFABR$)
       XX$ = REGLEIDO$(ARCHORDEFABR$, U&)
       ST% = Asc(Mid$(XX$, 108, 1))
       If ST% > 0 Then
          If ST% < 3 Then
             Call REGAPP("ORDESOR", XX$)
             'Debug.Print XX$
          End If
       End If
    Next U&
    Screen.MousePointer = 1
        '
    If Option2(1).Value = True Then
'        If Option1(0).Value = True Then Call FINAL("*SFEPRO0")
'        If Option1(1).Value = True Then Call FINAL("*SFEPRO1")
'        If Option1(2).Value = True Then Call FINAL("*SFEPRO2")
'        If Option1(3).Value = True Then Call FINAL("*SFEPRO3")
        If Option1(0).Value = True Then Call FINAL("?FEPROTODOS")
        If Option1(1).Value = True Then Call FINAL("?FEPRODESTINO")
        If Option1(2).Value = True Then Call FINAL("?FEPROCODIGO")
        If Option1(3).Value = True Then Call FINAL("?FEPROORDEN")
        GoTo 99
        
    End If
    '
    If CONSIMULACION% > 0 Then
'        If Option1(0).Value = True Then Call FINAL("*SEPROE0")
'        If Option1(1).Value = True Then Call FINAL("*SEPROE1")
'        If Option1(2).Value = True Then Call FINAL("*SEPROE2")
'        If Option1(3).Value = True Then Call FINAL("*SEPROE3")
        If Option1(0).Value = True Then Call FINAL("?SIFEPTODOS")
        If Option1(1).Value = True Then Call FINAL("?SIFEPDESTINO")
        If Option1(2).Value = True Then Call FINAL("?SIFEPCODIGO")
        If Option1(3).Value = True Then Call FINAL("?SIFEPORDEN")

    Else
'        If Option1(0).Value = True Then Call FINAL("*FEPROE0")
'        If Option1(1).Value = True Then Call FINAL("*FEPROE1")
'        If Option1(2).Value = True Then Call FINAL("*FEPROE2")
'        If Option1(3).Value = True Then Call FINAL("*FEPROE3")
        If Option1(0).Value = True Then Call FINAL("?FEPROTODOS")
        If Option1(1).Value = True Then Call FINAL("?FEPRODESTINO")
        If Option1(2).Value = True Then Call FINAL("?FEPROCODIGO")
        If Option1(3).Value = True Then Call FINAL("?FEPROORDEN")
    End If
99  CONSIMULACION = 0
End Sub
Private Sub Command2ORIGINAL_Click()
    '
    CONSIMULACION% = 0
    If EXISTE_TABLA("SIMU_ORDEFABR") > 0 Then
       If CantRegistros("SIMU_ORDEFABR", , "NOMESS") > 0 Then
         OPX% = COMALTER%("Opciones de Trabajo\\Listado Cálculo Normal\Listado Considerando Simulación")
         If OPX% > 1 Then
           CONSIMULACION% = XVALO(CONTROL("CARMAQUI", "ULTISIMU"))
           If CONSIMULACION% < 1 Then
              Call COMUNI("Atención!! No Hay Cálculos con Simulación de Fabricación\Se Visualizará Cálculo Normal")
           End If
         End If
       End If
    End If
    ARCHORDEFABR$ = "ORDEFABR"
    If CONSIMULACION% < 1 Then
       Call GENORDEFDAT(1)
    Else
       ARCHORDEFABR$ = "SRDEFABR"
       SIMUCARM.CONSIMULACION = 1
       Call SIMUCARM.GENORDEFDAT_SIMU(1)
       SIMUCARM.CONSIMULACION = 0
    End If

    
    If Option2(1).Value = True Then
        Screen.MousePointer = 11
'        Call COPYSTRU(ARCHORDEFABR$, "ORDESOR")
        Call BLANARCH("ORDESOR")
        For U& = 1 To ULTREG&(ARCHORDEFABR$)
           XX$ = REGLEIDO$(ARCHORDEFABR$, U&)
           ST% = Asc(Mid$(XX$, 108, 1))
           If ST% > 0 Then
              If ST% < 3 Then
                 Call REGAPP("ORDESOR", XX$)
              End If
           End If
        Next U&
        Screen.MousePointer = 1
        '
        If Option1(0).Value = True Then Call FINAL("*SFEPRO0")
        If Option1(1).Value = True Then Call FINAL("*SFEPRO1")
        If Option1(2).Value = True Then Call FINAL("*SFEPRO2")
        If Option1(3).Value = True Then Call FINAL("*SFEPRO3")
        Exit Sub
    End If
    '
    If Option1(0).Value = True Then Call FINAL("*FEPROE0")
    If Option1(1).Value = True Then Call FINAL("*FEPROE1")
    If Option1(2).Value = True Then Call FINAL("*FEPROE2")
    If Option1(3).Value = True Then Call FINAL("*FEPROE3")
    '
End Sub

Private Sub Command29_Click()
    '
    Call HELPONLINE("FEPROEN")
    '
End Sub

Private Sub Form_Load()
  Call APLICACIONSKINS(Me, Picture1)
End Sub

