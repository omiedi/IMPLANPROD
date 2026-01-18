VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form SETGEMA09 
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Gestión de Mantenimiento Industrial - Tablas y Configuracion"
   ClientHeight    =   6810
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   5550
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6810
   ScaleWidth      =   5550
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture999 
      Height          =   330
      Index           =   0
      Left            =   0
      ScaleHeight     =   270
      ScaleWidth      =   1110
      TabIndex        =   36
      Top             =   0
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Formulario de Impresión"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2610
      Left            =   120
      TabIndex        =   18
      Top             =   3960
      Width           =   5320
      Begin VB.CommandButton BOTOPENFORM 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   620
         Index           =   2
         Left            =   4485
         Picture         =   "SETGEMA09.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   33
         Top             =   1800
         Width           =   645
      End
      Begin VB.CommandButton BOTSELFORM 
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
         Height          =   285
         Index           =   2
         Left            =   4170
         TabIndex        =   32
         Top             =   1800
         Width           =   175
      End
      Begin VB.TextBox FORMUNAME 
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
         Index           =   2
         Left            =   2910
         TabIndex        =   31
         Text            =   " "
         Top             =   1800
         Width           =   1275
      End
      Begin VB.TextBox FORMUNAME 
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
         Index           =   1
         Left            =   2910
         TabIndex        =   27
         Text            =   " "
         Top             =   1080
         Width           =   1275
      End
      Begin VB.CommandButton BOTSELFORM 
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
         Height          =   285
         Index           =   1
         Left            =   4170
         TabIndex        =   26
         Top             =   1080
         Width           =   175
      End
      Begin VB.CommandButton BOTOPENFORM 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   620
         Index           =   1
         Left            =   4485
         Picture         =   "SETGEMA09.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   1080
         Width           =   645
      End
      Begin VB.CommandButton BOTOPENFORM 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   620
         Index           =   0
         Left            =   4485
         Picture         =   "SETGEMA09.frx":0614
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   315
         Width           =   645
      End
      Begin VB.CommandButton BOTSELFORM 
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
         Height          =   285
         Index           =   0
         Left            =   4170
         TabIndex        =   20
         Top             =   315
         Width           =   175
      End
      Begin VB.TextBox FORMUNAME 
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
         Index           =   0
         Left            =   2910
         TabIndex        =   19
         Text            =   " "
         Top             =   315
         Width           =   1275
      End
      Begin VB.Label FORMUECO 
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
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
         Index           =   2
         Left            =   180
         TabIndex        =   35
         Top             =   2130
         Width           =   4170
      End
      Begin VB.Label FORMUDES 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Cronograma de Mantenimiento:"
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
         Index           =   2
         Left            =   135
         TabIndex        =   34
         Top             =   1860
         Width           =   2685
      End
      Begin VB.Label FORMUDES 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Plan de Mantenimiento:"
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
         Index           =   1
         Left            =   495
         TabIndex        =   29
         Top             =   1140
         Width           =   2325
      End
      Begin VB.Label FORMUECO 
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
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
         Index           =   1
         Left            =   180
         TabIndex        =   28
         Top             =   1410
         Width           =   4170
      End
      Begin VB.Label FORMUECO 
         BackColor       =   &H80000009&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " "
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
         Index           =   0
         Left            =   180
         TabIndex        =   23
         Top             =   645
         Width           =   4170
      End
      Begin VB.Label FORMUDES 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Imagenes de Equipos:"
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
         Index           =   0
         Left            =   495
         TabIndex        =   22
         Top             =   375
         Width           =   2325
      End
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Height          =   1215
      Left            =   5400
      ScaleHeight     =   1215
      ScaleWidth      =   300
      TabIndex        =   24
      Top             =   3960
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00004040&
      Height          =   4215
      Left            =   4680
      ScaleHeight     =   4155
      ScaleWidth      =   1320
      TabIndex        =   11
      Top             =   -240
      Width           =   1380
      Begin VB.PictureBox MARCOBARRA 
         BackColor       =   &H000000FF&
         Height          =   135
         Left            =   105
         ScaleHeight     =   75
         ScaleWidth      =   585
         TabIndex        =   15
         Top             =   3300
         Width           =   650
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H00FF0000&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   0
            TabIndex        =   16
            Top             =   105
            Width           =   12000
         End
      End
      Begin VB.CommandButton Command3 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   105
         Picture         =   "SETGEMA09.frx":091E
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Ayuda en Línea"
         Top             =   1620
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
         Height          =   600
         Left            =   105
         Picture         =   "SETGEMA09.frx":0C28
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Guarda los Cambios Realizados"
         Top             =   1020
         Width           =   650
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
         Height          =   600
         Left            =   105
         Picture         =   "SETGEMA09.frx":0F32
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Guarda - Cierra - Abandona la Aplicación"
         Top             =   420
         Width           =   650
      End
      Begin VB.Image Image2 
         Height          =   390
         Left            =   -300
         Picture         =   "SETGEMA09.frx":107C
         Top             =   3570
         Width           =   1515
      End
   End
   Begin VB.Frame Frame9 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Tablas de Validación"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3690
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4425
      Begin VB.Frame Frame6 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Criterios y Tablas Generales"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   945
         Left            =   210
         TabIndex        =   2
         Top             =   360
         Width           =   4005
         Begin VB.CommandButton Command5 
            Height          =   330
            Left            =   3360
            Picture         =   "SETGEMA09.frx":2F9E
            Style           =   1  'Graphical
            TabIndex        =   9
            Top             =   315
            Width           =   435
         End
         Begin VB.OptionButton Option1 
            BackColor       =   &H00C0FFFF&
            Caption         =   "Tabla de Unidades de Medida"
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
            Index           =   6
            Left            =   315
            TabIndex        =   8
            Top             =   565
            Width           =   3375
         End
         Begin VB.OptionButton Option1 
            BackColor       =   &H00C0FFFF&
            Caption         =   "Tabla de Tipos de Material"
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
            Index           =   0
            Left            =   315
            TabIndex        =   3
            Top             =   315
            Width           =   3480
         End
      End
      Begin VB.Frame Frame11 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Configuración Industrial"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2085
         Left            =   210
         TabIndex        =   1
         Top             =   1440
         Width           =   4005
         Begin VB.OptionButton Option1 
            BackColor       =   &H00C0FFFF&
            Caption         =   "Lineas / Sistemas"
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
            Index           =   7
            Left            =   315
            TabIndex        =   30
            Top             =   1650
            Width           =   3500
         End
         Begin VB.CommandButton Command6 
            Height          =   330
            Left            =   3360
            Picture         =   "SETGEMA09.frx":30E8
            Style           =   1  'Graphical
            TabIndex        =   10
            Top             =   510
            Width           =   435
         End
         Begin VB.OptionButton Option1 
            BackColor       =   &H00C0FFFF&
            Caption         =   "Plantas "
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
            Index           =   1
            Left            =   315
            TabIndex        =   17
            Top             =   360
            Width           =   3500
         End
         Begin VB.OptionButton Option1 
            BackColor       =   &H00C0FFFF&
            Caption         =   "Maquinas y Equipos"
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
            Index           =   5
            Left            =   315
            TabIndex        =   7
            Top             =   1395
            Width           =   3500
         End
         Begin VB.OptionButton Option1 
            BackColor       =   &H00C0FFFF&
            Caption         =   "Categorias de Operarios"
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
            Index           =   4
            Left            =   315
            TabIndex        =   6
            Top             =   1140
            Width           =   3500
         End
         Begin VB.OptionButton Option1 
            BackColor       =   &H00C0FFFF&
            Caption         =   "Secciones y Centros de Costos"
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
            Index           =   2
            Left            =   315
            TabIndex        =   5
            Top             =   630
            Width           =   3500
         End
         Begin VB.OptionButton Option1 
            BackColor       =   &H00C0FFFF&
            Caption         =   "Depósitos y Almacenes"
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
            Index           =   3
            Left            =   315
            TabIndex        =   4
            Top             =   885
            Width           =   3500
         End
      End
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   225
      Index           =   0
      Left            =   1155
      OleObjectBlob   =   "SETGEMA09.frx":3232
      TabIndex        =   37
      Top             =   0
      Visible         =   0   'False
      Width           =   1590
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   2730
      OleObjectBlob   =   "SETGEMA09.frx":32AC
      Top             =   0
   End
End
Attribute VB_Name = "SETGEMA09"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim TABLA$(7)
Private Sub Check1_Click()
   Option1(11).Enabled = False
   If Check1.Value = 1 Then
     Option1(11).Enabled = True
   End If
End Sub

Private Sub BOTOPENFORM_Click(Index As Integer)
    '
    Call CONECRUN("FLFORMS/" + TRIM$(FORMUNAME(Index).TEXT), "Formularios de Impresion")
    If EXISTE%("FLFORMS.EXE") < 1 Then
        Call FINAL("")
    End If
    '
End Sub

Private Sub BOTSELFORM_Click(Index As Integer)
    '
    Call SELECHO("INDIFRM")
    VDEVU$ = VALACT1$("INDIFRM")
    If VDEVU$ <> "" Then
        FORMUNAME(Index).TEXT = VDEVU$
        FORMUECO(Index).Caption = VALACT2$("INDIFRM")
    End If
    '
End Sub

Private Sub Command1_Click()
   '
   Call CIERRARCH("*.*")
   Unload Me
   Call FRESHALL
   '
End Sub
Private Sub Command4_Click()
   Command4.Enabled = False
   Call Command5_Click
   Command4.Enabled = True
End Sub

Private Sub Command2_Click()
   '
   Call GRACONTROL("", "FORGRAFI", FORMUNAME(0))
   Call GRACONTROL("", "FOR-PMC", FORMUNAME(1))
   Call GRACONTROL("", "FOR-CMC", FORMUNAME(2))
   '
End Sub

Private Sub Command3_Click()
  '
  Call HELPONLINE("SETGEMA")
  '
End Sub

 Sub Command5_Click()
    '
    Command5.Enabled = False
    For Index = 0 To 12
      If Option1(Index) = True Then
        Call CARTABLA(Index)
        GoTo 99
      End If
   Next Index
99 Command5.Enabled = True
   '
End Sub

 Sub Command6_Click()
   Command6.Enabled = False
   Call Command5_Click
   Command6.Enabled = True
End Sub


'
Private Sub Form_Load()
  '
  Call APLICACIONSKINS(Me)
  
  TABLA$(0) = "TATIMAT"       ' tabla de tipos de material
  TABLA$(1) = "PLANTPRO"
  TABLA$(2) = "SECPROD"
  TABLA$(3) = "TADEPOSI"
  TABLA$(4) = "CATOPER"
  TABLA$(5) = "PADRON"
  TABLA$(6) = "UMEDIDA"
  '
  FORMUNAME(0) = TRIM$(CONTROL$("", "FORGRAFI"))
  FORMUECO(0) = ECOARCH$("INDIFRM", FORMUNAME(0))
  '
  FORMUNAME(1) = TRIM$(CONTROL$("", "FOR-PMC"))
  FORMUECO(1) = ECOARCH$("INDIFRM", FORMUNAME(1))
  '
  FORMUNAME(2) = TRIM$(CONTROL$("", "FOR-CMC"))
  FORMUECO(2) = ECOARCH$("INDIFRM", FORMUNAME(2))
  '
End Sub
Private Sub FORMUNAME_DblClick(Index As Integer)
    '
    Call SELECHO("INDIFRM")
    VDEVU$ = VALACT1$("INDIFRM")
    If VDEVU$ <> "" Then
        FORMUNAME(Index).TEXT = VDEVU$
        FORMUECO(Index).Caption = VALACT2$("INDIFRM")
    End If
    '
End Sub

'
 Sub Option1_Click(Index As Integer)
    '
    For KKI% = 0 To 6
      If KKI% <> Index Then
        Option1(KKI%).Value = False
      End If
    Next KKI%
    '
End Sub

Sub CARTABLA(Index)
   '
   If Index = 0 Then
       Call CARTATIMA
    ElseIf Index = 1 Then
       Call CARPLANTA
    ElseIf Index = 2 Then
       Call CARSECPROD
    ElseIf Index = 3 Then
       Call CARTADEPOSI
    ElseIf Index = 4 Then
       Call CARCATOPER
    ElseIf Index = 5 Then
       Call CARPADRON
    ElseIf Index = 6 Then
       Call CARUMEDIDA
    ElseIf Index = 7 Then
       Call CARABMSIS
   End If
   '
End Sub
'
Sub Option1_DblClick(Index As Integer)
   Call CARTABLA(Index)
End Sub
Sub CARPADRON()
  '
  On Error Resume Next
  AMAEQUI04.Show 1
  If Err.Number = 400 Then
    On Error GoTo 0
    Unload Me
  End If
  On Error GoTo 0
  '
End Sub
Sub CARPLANTA()
  '
  Call TRALOCAL("PLANTPRO", "PLANTPRO")
  '
'  ATRI$ = "BORDEIZQ=W" + TRIM$(Str$(Left - 1900))
'  ATRI$ = ATRI$ + "/BORDESUP=W" + TRIM$(Str$(Top + 850))
  VTB% = VENTABU%("PLANTPRO/" + ATRI$)
  If VTB% >= 0 Then
     JJ& = 0
     For U& = 1 To ULTREG&("!PLANTPRO")
        X$ = REGLEIDO$("!PLANTPRO", U&)
        JJ& = JJ& + 1
        Call GRAREG("PLANTPRO", X$, JJ&)
     Next U&
     Call SETULTREG("PLANTPRO", JJ&)
     Call CIERRARCH("PLANTPRO")
  End If
  '
End Sub
Sub CARTATIMA()
    '
    Call BLANARCH("!TABTIMA")
    For I& = 1 To ULTREG&("TABTIMA")
       X$ = REGLEIDO$("TABTIMA", I&)
       If Asc(Left$(X$, 1)) > 0 Then
           Call REGAPP("!TABTIMA", X$)
       End If
    Next I&
    Call CIERRARCH("!TABTIMA")
    '
10  VTB% = VENTABU%("TABTIMA")
    If VTB% < 0 Then Exit Sub
    J& = 0
    For I& = 1 To ULTREG&("!TABTIMA")
       X$ = REGLEIDO$("!TABTIMA", I&)
       If Asc(Left$(X$, 1)) > 0 Then
          DPP$ = Mid$(X$, 17, 1)
          If ECOARCH$("ECODEP", DPP$) = "" Then
             Call MENSERR(24, "Depósito no Definido")
             GoTo 10
          End If
          If Asc(DPP$) = 99 Then
             Call MENSERR(24, "Depósito 99 no Válido")
             GoTo 10
          End If
          J& = J& + 1
          Call GRAREG("TABTIMA", X$, J&)
          Call GRAREG("TATIMAT", X$, J&)
       End If
    Next I&
    Call SETULTREG("TABTIMA", J&)
    Call CIERRARCH("TABTIMA")
    Call SETULTREG("TATIMAT", J&)
    Call CIERRARCH("TATIMAT")
    Call CIERRARCH("!TABTIMA")
    '
End Sub
'
Sub CARTADEPOSI()
    '
    Call BLANARCH("!TADEPOSI")
    J& = 0
    For U& = 1 To ULTREG&("SUCURSAL")
        X$ = REGLEIDO$("SUCURSAL", U&)
        CODEPO% = Asc(Mid$(X$, 64, 1))
        If CODEPO% > 0 Then
           If CODEPO% <> 99 Then
              If J& < 62 Then
                 J& = J& + 1
                 Y$ = REGBLAN$("TADEPOSI")
                 Call REPLA(Y$, Chr$(CODEPO%), 1, 1)
                 Call REPLA(Y$, Left$(X$, 63), 2, 63)
                 Call REGAPP("!TADEPOSI", Y$)
              End If
           End If
        End If
    Next U&
    Call CIERRARCH("SUCURSAL")
    Call CIERRARCH("!TADEPOSI")
    '
    VTB% = VENTABU%("TADEPOSI")
    If VTB% < 0 Then Exit Sub
    '
    J& = 0
    For U& = 1 To ULTREG&("!TADEPOSI")
        X$ = REGLEIDO$("!TADEPOSI", U&)
        CODEPO% = Asc(Left$(X$, 1))
        If CODEPO% > 0 Then
           If CODEPO% < 99 Then
             If J& < 62 Then
               J& = J& + 1
               Y$ = REGBLAN$("SUCURSAL")
               Call REPLA(Y$, Mid$(X$, 2, 63), 1, 63)
               Call REPLA(Y$, Chr$(CODEPO%), 64, 1)
               Call GRAREG("SUCURSAL", Y$, J&)
               Y$ = REGBLAN$("ECODEP")
               Call REPLA(Y$, Left$(X$, 21), 1, 21)
               Call GRAREG("ECODEP", Y$, J&)
               Call GRAREG("TADEPOSI", X$, J&)
             End If
           End If
        End If
    Next U&
    Call SETULTREG("SUCURSAL", J&)
    Call CIERRARCH("SUCURSAL")
    Call SETULTREG("TADEPOSI", J&)
    Call CIERRARCH("TADEPOSI")
    '
    Y$ = REGBLAN$("ECODEP")
    Call REPLA(Y$, Chr$(99) + "Predeterminado      ", 1, 21)
    J& = J& + 1
    Call GRAREG("ECODEP", Y$, J&)
    Call SETULTREG("ECODEP", J&)
    Call CIERRARCH("ECODEP")
    Call CIERRARCH("!TADEPOSI")
    '
End Sub
'
Sub CARSECPROD()
    '
    Call BLANARCH("!SECPROD")
    J& = 0
    For U& = 1 To ULTREG&("SECPROD")
        X$ = REGLEIDO$("SECPROD", U&)
        COSECC = CVS(Left$(X$, 4))
        If COSECC >= 1 Then
            Call REGAPP("!SECPROD", X$)
        End If
    Next U&
    Call CIERRARCH("SECPROD")
    Call CIERRARCH("!SECPROD")
    '
    VTB% = VENTABU%("SECPROD")
    If VTB% < 0 Then Exit Sub
    '
    J& = 0
    For U& = 1 To ULTREG&("!SECPROD")
        X$ = REGLEIDO$("!SECPROD", U&)
        COSECC = CVS(Left$(X$, 4))
        If COSECC >= 1 Then
            J& = J& + 1
            Call GRAREG("SECPROD", X$, J&)
        End If
    Next U&
    '
    Call SETULTREG("SECPROD", J&)
    Call CIERRARCH("SECPROD")
    Call CIERRARCH("!SECPROD")
    '
End Sub

Sub CARCATOPER()
   '
   With OPADMAQ09
     .Option1(3).Value = True
     Call .Command16_Click
   End With
   '
End Sub

Sub CARUMEDIDA()
   '
   Call TRALOCAL("UNIMED", "")
10 VTB% = VENTABU%("TUMEDIDA/NC/TITUPAN=Unidades de Medida")
   If VTB% >= 0 Then
      JJ& = 0
      For U& = 1 To ULTREG&("!UNIMED")
         X$ = REGLEIDO$("!UNIMED", U&)
         If TRIM$(Left$(X$, 4)) <> "" Then
           If Len(TRIM$(Left$(X$, 4))) > 2 Then
             Call MENSERR(24, "Codigo de Unidad no Válido.\Máximo 2 Posiciones")
             GoTo 10
           End If
           Y$ = Left$(X$, 2) + Mid$(X$, 5, 28)
           JJ& = JJ& + 1
           Call GRAREG("UNIMED", X$, JJ&)
           Call GRAREG("UMEDIDA", Y$, JJ&)
         End If
      Next U&
      Call SETULTREG("UNIMED", JJ&)
      Call SETULTREG("UMEDIDA", JJ&)
   End If
   Call CIERRARCH("UNIMED")
   Call CIERRARCH("UMEDIDA")
   '
End Sub

Sub CARABMSIS()
    '
'    Call BLANARCH("!ABMSIS")
'    For U& = 1 To ULTREG&("ABMSIS")
'        X$ = REGLEIDO$("ABMSIS", U&)
'        Call REGAPP("!ABMSIS", X$)
'    Next U&
    Call TRALOCAL("ABMSIS", "")
'    Call CIERRARCH("ABMSIS")
'    Call CIERRARCH("!ABMSIS")
    '
    VTB% = VENTABU%("ABMSIS")
    If VTB% < 0 Then Exit Sub
    '
    Call TRACENTRAL("ABMSIS", "")
'    JJ& = 0
'    For U& = 1 To ULTREG&("!ABMSIS")
'        X$ = REGLEIDO$("!ABMSIS", U&)
'        Y$ = REGBLAN$("ABMSIS")
'        Call REPLA(Y$, Mid$(X$, 1, 64), 1, 64)
'        Call REPLA(Y$, Mid$(X$, 65, 32), 65, 32)
'        Call REPLA(Y$, Mid$(X$, 97, 16), 97, 16)
'        Call REPLA(Y$, Mid$(X$, 113, 32), 113, 32)
'        JJ& = JJ& + 1
'        Call GRAREG("ABMSIS", Y$, JJ&)
'    Next U&
    '
'    Call SETULTREG("ABMSIS", JJ&)
'    Call CIERRARCH("ABMSIS")
'    Call CIERRARCH("!ABMSIS")
    '
End Sub

