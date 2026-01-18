VERSION 5.00
Begin VB.Form OPADMAQ07 
   BackColor       =   &H00C0FFC0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Empleados, Operarios y Especialistas   - Configuración"
   ClientHeight    =   3240
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   5970
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3240
   ScaleWidth      =   5970
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00006000&
      Height          =   8100
      Left            =   5085
      ScaleHeight     =   8040
      ScaleWidth      =   900
      TabIndex        =   12
      Top             =   0
      Width           =   960
      Begin VB.PictureBox MARCOBARRA 
         BackColor       =   &H000000FF&
         Height          =   135
         Left            =   120
         ScaleHeight     =   75
         ScaleWidth      =   585
         TabIndex        =   15
         Top             =   2310
         Width           =   640
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
      Begin VB.CommandButton command1 
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
         Height          =   690
         Left            =   105
         Picture         =   "OPADMAQ07.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   14
         ToolTipText     =   "Cierra - Abandona la Aplicación"
         Top             =   300
         Width           =   640
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Save"
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
         Left            =   105
         Picture         =   "OPADMAQ07.frx":014A
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Guardar "
         Top             =   990
         Width           =   640
      End
      Begin VB.Image Image2 
         Height          =   390
         Left            =   -315
         Picture         =   "OPADMAQ07.frx":0454
         Top             =   2680
         Width           =   1515
      End
   End
   Begin VB.Frame Frame9 
      BackColor       =   &H00C0FFC0&
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
      Height          =   2910
      Left            =   210
      TabIndex        =   0
      Top             =   210
      Width           =   4635
      Begin VB.Frame Frame6 
         BackColor       =   &H00C0FFC0&
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
         Height          =   1080
         Left            =   210
         TabIndex        =   6
         Top             =   1680
         Width           =   4215
         Begin VB.OptionButton Option1 
            BackColor       =   &H00C0FFC0&
            Caption         =   "Categorias de Empleados"
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
            Left            =   525
            TabIndex        =   7
            Top             =   690
            Width           =   3375
         End
         Begin VB.OptionButton Option1 
            BackColor       =   &H00C0FFC0&
            Caption         =   "Asignaciones y Subsidios Familiares"
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
            Index           =   7
            Left            =   315
            TabIndex        =   20
            Top             =   2565
            Visible         =   0   'False
            Width           =   3855
         End
         Begin VB.OptionButton Option1 
            BackColor       =   &H00C0FFC0&
            Caption         =   "Retenciones Legales y Aportes"
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
            TabIndex        =   19
            Top             =   2295
            Visible         =   0   'False
            Width           =   3855
         End
         Begin VB.OptionButton Option1 
            BackColor       =   &H00C0FFC0&
            Caption         =   "Remuneraciones Ordinarias"
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
            TabIndex        =   18
            Top             =   2025
            Visible         =   0   'False
            Width           =   3735
         End
         Begin VB.OptionButton Option1 
            BackColor       =   &H00C0FFC0&
            Caption         =   "Salarios Nominales Basicos"
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
            TabIndex        =   17
            Top             =   1470
            Visible         =   0   'False
            Width           =   3615
         End
         Begin VB.OptionButton Option1 
            BackColor       =   &H00C0FFC0&
            Caption         =   "Parametros Individuales"
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
            Index           =   8
            Left            =   315
            TabIndex        =   9
            Top             =   1755
            Visible         =   0   'False
            Width           =   3375
         End
         Begin VB.OptionButton Option1 
            BackColor       =   &H00C0FFC0&
            Caption         =   "Grupos de Empleados"
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
            Left            =   525
            TabIndex        =   8
            Top             =   420
            Width           =   3480
         End
      End
      Begin VB.CommandButton Command20 
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
         Left            =   3255
         TabIndex        =   1
         Top             =   525
         Width           =   175
      End
      Begin VB.CommandButton Command16 
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
         Height          =   880
         Left            =   3570
         Picture         =   "OPADMAQ07.frx":2376
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   525
         Width           =   855
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H00E0E0E0&
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
         Top             =   1050
         Width           =   3225
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H00E0E0E0&
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
         Left            =   1680
         TabIndex        =   2
         Text            =   " "
         Top             =   525
         Width           =   1735
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Tabla:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   -735
         TabIndex        =   5
         Top             =   630
         Width           =   2325
      End
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Criterios Técnicos de Clasificación"
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
      Left            =   630
      TabIndex        =   10
      Top             =   2940
      Visible         =   0   'False
      Width           =   3795
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Tablas de Clasificación Técnica"
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
      Left            =   630
      TabIndex        =   11
      Top             =   2940
      Visible         =   0   'False
      Width           =   3690
   End
End
Attribute VB_Name = "OPADMAQ07"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim TABLA$(16)
Dim MODUCONE$

Sub Command1_Click()
   '
   Call CIERRARCH("*.*")
   Hide
   '
End Sub

Sub Command16_Click()
   '
   Command16.Enabled = False
   For Index = 0 To 8
     If Option1(Index) = True Then
        Call CARTABLA(Index)
        GoTo 99
     End If
   Next Index
99 Command16.Enabled = True
   '
End Sub

Private Sub Form_Load()
  '
  TABLA$(0) = "DATGREMP"       'GRUPOS DE EMPLEADOS
  TABLA$(1) = ""
  TABLA$(2) = ""
  TABLA$(3) = "CATEGO"         'CATEGORIA DE EMPLEADOS
  TABLA$(4) = "NOMREM"         'SALARIOS BASICOS
  TABLA$(5) = "CONCEP"         'REMUNERACIONES ORDINARIAS
  TABLA$(6) = "RETLEG"         'RETENCIONES LEGALES Y APORTES
  TABLA$(7) = "ASIGNA"         'ASIGNACIONES FLIARES Y SUBSIDIOS
  TABLA$(8) = "PARAME"         'PARAMETROS INDIVIDUALES
  '
  If CONTROL$("", "MOREPOSI") = "PROHABIT" Then
    Option3.Value = True
  End If
  '
End Sub
'
 Sub Option1_Click(Index As Integer)
    '
    For KKI% = 0 To 8
      If KKI% <> Index Then
        Option1(KKI%).Value = False
      End If
    Next KKI%
    '
    Text1.TEXT = TABLA(Index)
    Option1(Index).Value = True
    Text2.TEXT = Option1(Index).Caption
    Command16.Enabled = True
    'Command16.SetFocus 'No se puede asignar el foco a
                        'un objeto no visible.
                        '(cuando lo llamo desde la barra
                        ' de menu)
    '
End Sub

Sub CARTABLA(Index)
   '
   If Index = 0 Then
       Call CARGRUEMP 'GRUPO DE EMPLEADOS
    ElseIf Index = 3 Then
       Call CARCATOPE 'CATEGORIA DE EMPLEADOS
    ElseIf Index = 4 Then
       Call CARSALBAS 'SALARIOS BASICOS
    ElseIf Index = 5 Then
       REMORD05.Show 1 'REMUNERACIONES ORDINARIAS
    ElseIf Index = 6 Then
       RETAPO05.Show 1 'RETENCIONES LEGALES Y APORTES
    ElseIf Index = 7 Then
       Call CARASISUB 'ASIGNACIONES FLIARES Y SUBSIDIOS
    ElseIf Index = 8 Then
       Call CARPARAINDI 'PARAMETROS INDIVIDUALES
   End If
   '
End Sub
'
Private Sub Option1_DblClick(Index As Integer)
   Call CARTABLA(Index)
End Sub

Private Sub Text1_GotFocus()
   Command16.SetFocus
End Sub

Private Sub Text2_GotFocus()
   Command16.SetFocus
End Sub
Sub CARGRUEMP()
  '
  Call TRALOCAL("DATGREMP", "DATGREMP")
  Call TRALOCAL("DATGREM1", "DATGREM1")
  '
  ATRI$ = "BORDEIZQ=W" + TRIM$(Str$(Left - 1800))
  ATRI$ = ATRI$ + "/BORDESUP=W" + TRIM$(Str$(Top + 850))
  VTB% = VENTABU%("DATGREMP/" + ATRI)
  If VTB% >= 0 Then
     JJ& = 0
     For u& = 1 To ULTREG&("!DATGREMP")
        X$ = REGLEIDO$("!DATGREMP", u&)
         JJ& = JJ& + 1
        Call GRAREG("DATGREMP", X$, JJ&)
        '
         Z$ = Space(27)
         Call REPLA(Z$, Chr$(JJ&), 1, 1)
         Call REPLA(Z$, TRIM$(Mid$(X$, 1, 25)), 2, 25)
        Call GRAREG("DATGREM1", Z$, JJ&) 'INIDICE/ECO
        '
     Next u&
     Call SETULTREG("DATGREMP", JJ&)
     Call CIERRARCH("DATGREMP")
     Call SETULTREG("DATGREM1", JJ&) 'INDICE/ECO
     Call CIERRARCH("DATGREM1") 'INDICE/ECO
  End If
  '
End Sub
'
Sub CARCATOPE()
  '
  Call TRALOCAL("CATEGO", "CATEGO")
  '
  ATRI$ = ""
  VTB% = VENTABU%("CATEGO/" + ATRI$)
  If VTB% >= 0 Then
     PUCONTROL$ = String$(255, 0)
     JJ& = 0
     For u& = 1 To ULTREG&("!CATEGO")
        X$ = REGLEIDO$("!CATEGO", u&)
        COCAT% = Asc(Left$(X$, 1))
        If COCAT% > 0 Then
          If COCAT% < 256 Then
            If Mid$(PUCONTROL$, COCAT%, 1) = Chr$(0) Then
              Mid$(PUCONTROL$, COCAT%, 1) = Chr$(1)
              JJ& = JJ& + 1
              Call GRAREG("CATEGO", X$, JJ&)
            End If
          End If
        End If
     Next u&
     Call SETULTREG("CATEGO", JJ&)
     Call CIERRARCH("CATEGO")
     Call FILESORT("CATEGO", "", 1, 1)
  End If
  '
End Sub
Sub CARSALBAS()
  '
  Call TRALOCAL("NOMREM", "NOMREM")
  ATRI$ = "BORDEIZQ=W" + TRIM$(Str$(Left - 2600))
  ATRI$ = ATRI$ + "/BORDESUP=W" + TRIM$(Str$(Top + 850))
  VTB% = VENTABU%("NOMREM/" + ATRI$)
  If VTB% >= 0 Then
     JJ& = 0
     For u& = 1 To ULTREG&("!NOMREM")
        X$ = REGLEIDO$("!NOMREM", u&)
        JJ& = JJ& + 1
        Call GRAREG("NOMREM", X$, JJ&)
     Next u&
     Call SETULTREG("NOMREM", JJ&)
     Call CIERRARCH("NOMREM")
  End If
  '
End Sub
Sub CARPARAINDI()
  '
  Call TRALOCAL("PARAME", "PARAME")
  '
  ATRI$ = "BORDEIZQ=W" + TRIM$(Str$(Left - 1900))
  ATRI$ = ATRI$ + "/BORDESUP=W" + TRIM$(Str$(Top + 850))
  VTB% = VENTABU%("PARAME/" + ATRI$)
  If VTB% >= 0 Then
     JJ& = 0
     For u& = 1 To ULTREG&("!PARAME")
        X$ = REGLEIDO$("!PARAME", u&)
        JJ& = JJ& + 1
        Call GRAREG("PARAME", X$, JJ&)
     Next u&
     Call SETULTREG("PARAME", JJ&)
     Call CIERRARCH("PARAME")
  End If
  '
End Sub
Sub CARASISUB()
  '
  Call TRALOCAL("ASIGNA", "ASIGNA")
  '
  ATRI$ = "BORDEIZQ=W" + TRIM$(Str$(Left - 1400))
  ATRI$ = ATRI$ + "/BORDESUP=W" + TRIM$(Str$(Top + 850))
  VTB% = VENTABU%("ASIGNA/" + ATRI$)
  If VTB% >= 0 Then
     JJ& = 0
     For u& = 1 To ULTREG&("!ASIGNA")
        X$ = REGLEIDO$("!ASIGNA", u&)
        JJ& = JJ& + 1
        Call GRAREG("ASIGNA", X$, JJ&)
     Next u&
     Call SETULTREG("ASIGNA", JJ&)
     Call CIERRARCH("ASIGNA")
  End If
  '
End Sub
