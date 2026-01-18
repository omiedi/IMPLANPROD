VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form FSTF054 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Configuración de Impresoras de Formularios F054 y Reg. de Armado"
   ClientHeight    =   5745
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   6360
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
   ScaleHeight     =   5745
   ScaleWidth      =   6360
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame3 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Puerto de Impresion Etiqueta Grande"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   810
      Left            =   120
      TabIndex        =   21
      Top             =   2160
      Width           =   5175
      Begin VB.ComboBox Combo4 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   975
         TabIndex        =   22
         Text            =   "Combo4"
         Top             =   360
         Width           =   4035
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Puerto:"
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
         Left            =   -1440
         TabIndex        =   23
         Top             =   420
         Width           =   2325
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2715
      Left            =   120
      TabIndex        =   14
      Top             =   3000
      Width           =   5160
      Begin VB.TextBox Text3 
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
         Height          =   525
         Left            =   135
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   24
         Text            =   "FSTF054.frx":0000
         Top             =   2040
         Width           =   4920
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
         Height          =   525
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   17
         Text            =   "FSTF054.frx":0002
         Top             =   1200
         Width           =   4920
      End
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
         Height          =   525
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   16
         Text            =   "FSTF054.frx":0004
         Top             =   360
         Width           =   4920
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Ruta Archivo Excel Registro de Armado C15"
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
         Left            =   120
         TabIndex        =   25
         Top             =   1800
         Width           =   5085
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Ruta Archivo Excel Registro de Armado "
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
         Left            =   100
         TabIndex        =   18
         Top             =   960
         Width           =   5085
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Ruta Archivo Excel Etiquetas Aut. y Hoja Normal:"
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
         Left            =   -360
         TabIndex        =   15
         Top             =   120
         Width           =   5085
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1995
      Left            =   120
      TabIndex        =   9
      Top             =   120
      Width           =   5160
      Begin VB.ComboBox Combo3 
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
         Left            =   2160
         TabIndex        =   19
         Text            =   "Combo3"
         Top             =   1440
         Width           =   2865
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
         Left            =   2160
         TabIndex        =   12
         Text            =   "Combo1"
         Top             =   240
         Width           =   2865
      End
      Begin VB.ComboBox Combo2 
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
         Left            =   2160
         TabIndex        =   10
         Text            =   "Combo2"
         Top             =   840
         Width           =   2865
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Impresora Reg. Armado:"
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
         Left            =   0
         TabIndex        =   20
         Top             =   1440
         Width           =   2085
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Impresora Hoja Normal:"
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
         Left            =   0
         TabIndex        =   13
         Top             =   840
         Width           =   2085
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Impresora de Etiquetas Autoadhesiva Datos"
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
         Left            =   0
         TabIndex        =   11
         Top             =   240
         Width           =   2085
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00404040&
      Height          =   5730
      Left            =   5460
      ScaleHeight     =   5670
      ScaleWidth      =   1110
      TabIndex        =   6
      Top             =   0
      Width           =   1170
      Begin VB.CommandButton command1 
         Appearance      =   0  'Flat
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
         Picture         =   "FSTF054.frx":0006
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   105
         Width           =   645
      End
      Begin VB.CommandButton Command2 
         Appearance      =   0  'Flat
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
         Picture         =   "FSTF054.frx":0150
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   630
         Width           =   645
      End
      Begin VB.Image Image1 
         Height          =   390
         Left            =   -315
         Picture         =   "FSTF054.frx":045A
         ToolTipText     =   "Acceso Directo a Otras Aplicaciones"
         Top             =   5280
         Width           =   1515
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Impresión de Informe de Ensayo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1275
      Left            =   360
      TabIndex        =   0
      Top             =   5880
      Width           =   5160
      Begin VB.CommandButton Command5 
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
         Height          =   780
         Left            =   4080
         Picture         =   "FSTF054.frx":237C
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   315
         Width           =   750
      End
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
         Height          =   360
         Left            =   210
         TabIndex        =   3
         TabStop         =   0   'False
         Text            =   " "
         Top             =   735
         Width           =   3750
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
         Left            =   3780
         TabIndex        =   2
         Top             =   315
         Width           =   175
      End
      Begin VB.TextBox Text4 
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
         Left            =   2415
         TabIndex        =   1
         Text            =   " "
         Top             =   315
         Width           =   1380
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Formulario:"
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
         Left            =   0
         TabIndex        =   5
         Top             =   420
         Width           =   2325
      End
   End
   Begin MSComDlg.CommonDialog cmd 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "FSTF054"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
    Call CIERRARCH("*.*")
    Unload Me
End Sub

Private Sub Command2_Click()
    '
    Call GRACONTROL(IDENTER$, "PRTQEYM1", Combo1.TEXT)
    Call GRACONTROL("*", "PRTQEYM1", Combo1.TEXT)
    
    Call GRACONTROL(IDENTER$, "PRTQEYM2", Combo2.TEXT)
    Call GRACONTROL("*", "PRTQEYM2", Combo2.TEXT)

    Call GRACONTROL(IDENTER$, "PRTQEYM3", Combo3.TEXT)
    Call GRACONTROL("*", "PRTQEYM3", Combo3.TEXT)
    
    Call GRACONTROL(IDENTER$, "PORTSTIK", Combo4.TEXT)
    Call GRACONTROL("*", "PORTSTIK", Combo4.TEXT)

    Call GRACONTROL("TRAZABLE", "EXCEAUTO", Text2.TEXT)

    Call GRACONTROL("TRAZABLE", "EXCEARMA", Text1.TEXT)
    
    Call GRACONTROL("TRAZABLE", "EXARMA15", Text3.TEXT)
    '
    '
    Call CIERRARCH("*.*")
    '
End Sub

Private Sub Command3_Click()
    Call SELECHO("INDIFRM")
    VDEVU$ = VALACT1$("INDIFRM")
    If VDEVU$ <> "" Then
        Text4.TEXT = VDEVU$
        Text5.TEXT = VALACT2$("INDIFRM")
    End If
End Sub


Private Sub Command5_Click()
    Call CONECRUN("FLFORMS/" + TRIM$(Text4.TEXT), "Formularios de Impresion")
End Sub

Private Sub Form_Load()
    '
    PDES$ = CONTROL$(IDENTER$, "PRTQEYM1")
    If PDES$ = "" Then PDES$ = CONTROL$("*", "PRTQEYM1")
    
    Text2.TEXT = CONTROL("TRAZABLE", "EXCEAUTO")

    Text1.TEXT = CONTROL("TRAZABLE", "EXCEARMA")
    
    Text3.TEXT = CONTROL("TRAZABLE", "EXARMA15")

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

    PDES$ = CONTROL$(IDENTER$, "PRTQEYM2")
    If PDES$ = "" Then PDES$ = CONTROL$("*", "PRTQEYM2")

    Combo2.Clear
    On Error Resume Next
    Combo2.TEXT = Printer.DeviceName
    On Error GoTo 0
    For Each PRX In Printers
      Combo2.AddItem PRX.DeviceName
      If PRX.DeviceName = PDES$ Then
        Combo2.TEXT = PDES$
      End If
    Next

    PDES$ = CONTROL$(IDENTER$, "PRTQEYM3")
    If PDES$ = "" Then PDES$ = CONTROL$("*", "PRTQEYM3")

    Combo3.Clear
    On Error Resume Next
    Combo2.TEXT = Printer.DeviceName
    On Error GoTo 0
    For Each PRX In Printers
      Combo3.AddItem PRX.DeviceName
      If PRX.DeviceName = PDES$ Then
        Combo3.TEXT = PDES$
      End If
    Next

    '

    PFAC$ = CONTROL$(IDENTER$, "PORTSTIK")
    If PFAC$ = "" Then PFAC$ = CONTROL$("*", "PORTSTIK")
    Combo4.Clear
    Dim PRX1 As Printer
    On Error Resume Next
    Combo4.TEXT = Printer.DeviceName
    On Error GoTo 0
    For Each PRX1 In Printers
      Combo4.AddItem PRX1.DeviceName
      If PRX1.DeviceName = PFAC$ Then
        Combo4.TEXT = PFAC$
      End If
    Next
    PFAC$ = TRIM$(CONTROL$(IDENTER$, "PORTSTIK"))
    If PFAC$ <> "" Then Combo4.TEXT = PFAC$
    
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

Private Sub Text3_DblClick()
   RUTAX$ = TRIM$(SELECCION_Archivo$("*.*", cmd))
   If RUTAX$ <> "" Then Text3 = RUTAX$
   Screen.MousePointer = 1

End Sub
