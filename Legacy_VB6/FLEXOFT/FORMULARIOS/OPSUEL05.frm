VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Begin VB.Form OPSUEL05 
   BackColor       =   &H00C0FFC0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Sueldos y Jornales - Configuración"
   ClientHeight    =   5052
   ClientLeft      =   48
   ClientTop       =   300
   ClientWidth     =   5964
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5052
   ScaleWidth      =   5964
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture999 
      Height          =   435
      Index           =   0
      Left            =   105
      ScaleHeight     =   384
      ScaleWidth      =   1128
      TabIndex        =   22
      Top             =   735
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00006000&
      Height          =   8100
      Left            =   5085
      ScaleHeight     =   8052
      ScaleWidth      =   912
      TabIndex        =   12
      Top             =   0
      Width           =   960
      Begin VB.PictureBox MARCOBARRA 
         BackColor       =   &H000000FF&
         Height          =   135
         Left            =   120
         ScaleHeight     =   84
         ScaleWidth      =   588
         TabIndex        =   15
         Top             =   3945
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
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   690
         Left            =   105
         Picture         =   "OPSUEL05.frx":0000
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
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   690
         Left            =   105
         Picture         =   "OPSUEL05.frx":014A
         Style           =   1  'Graphical
         TabIndex        =   13
         ToolTipText     =   "Guardar "
         Top             =   990
         Width           =   640
      End
      Begin VB.Image Image2 
         Height          =   312
         Left            =   -312
         Picture         =   "OPSUEL05.frx":0454
         Top             =   4188
         Width           =   1212
      End
   End
   Begin VB.Frame Frame9 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Tablas de Validación"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4650
      Left            =   210
      TabIndex        =   0
      Top             =   210
      Width           =   4635
      Begin VB.Frame Frame6 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Criterios y Tablas Generales"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2730
         Left            =   210
         TabIndex        =   6
         Top             =   1680
         Width           =   4215
         Begin VB.OptionButton Option1 
            BackColor       =   &H00C0FFC0&
            Caption         =   "Sectores Productivos"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   7.8
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   9
            Left            =   315
            TabIndex        =   21
            Top             =   355
            Width           =   3480
         End
         Begin VB.OptionButton Option1 
            BackColor       =   &H00C0FFC0&
            Caption         =   "Asignaciones y Subsidios Familiares"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   7.8
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
            Top             =   2250
            Width           =   3855
         End
         Begin VB.OptionButton Option1 
            BackColor       =   &H00C0FFC0&
            Caption         =   "Retenciones Legales y Aportes"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   7.8
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
            Top             =   1980
            Width           =   3855
         End
         Begin VB.OptionButton Option1 
            BackColor       =   &H00C0FFC0&
            Caption         =   "Remuneraciones Ordinarias"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   7.8
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
            Top             =   1710
            Width           =   3735
         End
         Begin VB.OptionButton Option1 
            BackColor       =   &H00C0FFC0&
            Caption         =   "Salarios Nominales Basicos"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   7.8
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
            Top             =   1170
            Width           =   3615
         End
         Begin VB.OptionButton Option1 
            BackColor       =   &H00C0FFC0&
            Caption         =   "Parámetros Individuales"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   7.8
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
            Top             =   1440
            Width           =   3375
         End
         Begin VB.OptionButton Option1 
            BackColor       =   &H00C0FFC0&
            Caption         =   "Grupos de Empleados"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   7.8
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   0
            Left            =   315
            TabIndex        =   8
            Top             =   630
            Width           =   3480
         End
         Begin VB.OptionButton Option1 
            BackColor       =   &H00C0FFC0&
            Caption         =   "Categorías de Empleados"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   7.8
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   3
            Left            =   315
            TabIndex        =   7
            Top             =   900
            Width           =   3375
         End
      End
      Begin VB.CommandButton Command20 
         Caption         =   "."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.2
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
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   880
         Left            =   3570
         Picture         =   "OPSUEL05.frx":2376
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   525
         Width           =   855
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
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
            Size            =   9.6
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
         Width           =   1695
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Tabla:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.6
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
         Size            =   7.8
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
         Size            =   7.8
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
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   420
      OleObjectBlob   =   "OPSUEL05.frx":2680
      Top             =   0
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel999 
      Height          =   435
      Index           =   0
      Left            =   0
      OleObjectBlob   =   "OPSUEL05.frx":28B4
      TabIndex        =   23
      Top             =   315
      Visible         =   0   'False
      Width           =   1590
   End
End
Attribute VB_Name = "OPSUEL05"
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
  TABLA$(9) = "SECPROD"         'SECTORES PRODUCTIVOS
  '
  Call APLICACIONSKINS(Me, Picture1)

  If Control$("", "MOREPOSI") = "PROHABIT" Then
    Option3.Value = True
  End If
  '
End Sub
'
 Sub Option1_Click(Index As Integer)
    '
    For KKI% = 0 To 9
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
       REMDET05.Show 1        'REMUNERACIONES ORDINARIAS
       Call FRESHALL
       Call COSUJO05.HACECONCEP
    ElseIf Index = 6 Then
       RETDET05.Show 1         'RETENCIONES LEGALES Y APORTES
       Call FRESHALL
       Call COSUJO05.HACECONCEP
    ElseIf Index = 7 Then
       Call CARASISUB 'ASIGNACIONES FLIARES Y SUBSIDIOS
       Call FRESHALL
       Call COSUJO05.HACECONCEP
    ElseIf Index = 8 Then
       Call CARPARAINDI 'PARAMETROS INDIVIDUALES
    ElseIf Index = 9 Then
       Call CARSECPROD  'SECTORES PRODUCTIVOS
   End If
   '
End Sub
'
Private Sub Option1_DBLClick(Index As Integer)
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
     For U& = 1 To ULTREG&("!DATGREMP")
        X$ = REGLEIDO$("!DATGREMP", U&)
         JJ& = JJ& + 1
        Call GRAREG("DATGREMP", X$, JJ&)
        '
         Z$ = Space(27)
         Call REPLA(Z$, Chr$(JJ&), 1, 1)
         Call REPLA(Z$, TRIM$(Mid$(X$, 1, 25)), 2, 25)
        Call GRAREG("DATGREM1", Z$, JJ&) 'INIDICE/ECO
        '
     Next U&
     Call SETULTREG("DATGREMP", JJ&)
     Call CIERRARCH("DATGREMP")
     Call SETULTREG("DATGREM1", JJ&) 'INDICE/ECO
     Call CIERRARCH("DATGREM1") 'INDICE/ECO
  End If
  '
End Sub
Sub CARCATOPE()
  '
  Call TRALOCAL("CATEGO", "CATEGO")
  '
  ATRI$ = "BORDEIZQ=W" + TRIM$(Str$(Left - 1400))
  ATRI$ = ATRI$ + "/BORDESUP=W" + TRIM$(Str$(Top + 850))
  VTB% = VENTABU%("CATEGO/" + ATRI$)
  If VTB% >= 0 Then
     PUCONTROL$ = String$(255, 0)
     JJ& = 0
     For U& = 1 To ULTREG&("!CATEGO")
        X$ = REGLEIDO$("!CATEGO", U&)
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
     Next U&
     Call SETULTREG("CATEGO", JJ&)
     Call CIERRARCH("CATEGO")
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
     For U& = 1 To ULTREG&("!NOMREM")
        X$ = REGLEIDO$("!NOMREM", U&)
        JJ& = JJ& + 1
        Call GRAREG("NOMREM", X$, JJ&)
     Next U&
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
     For U& = 1 To ULTREG&("!PARAME")
        X$ = REGLEIDO$("!PARAME", U&)
        JJ& = JJ& + 1
        Call GRAREG("PARAME", X$, JJ&)
     Next U&
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
     For U& = 1 To ULTREG&("!ASIGNA")
        X$ = REGLEIDO$("!ASIGNA", U&)
        JJ& = JJ& + 1
        Call GRAREG("ASIGNA", X$, JJ&)
     Next U&
     Call SETULTREG("ASIGNA", JJ&)
     Call CIERRARCH("ASIGNA")
  End If
  '
End Sub
Sub CARSECPROD() 'SECTORES PRODUCTIVOS
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
    ATRI$ = "BORDEIZQ=W" + TRIM$(Str$(Left - 390))
    ATRI$ = ATRI$ + "/BORDESUP=W" + TRIM$(Str$(Top + 50))
    VTB% = VENTABU%("SECPRO2/" + ATRI$) 'XQ SECPROD.PRF ME MUESTRA LOS DEPOSITOS
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
