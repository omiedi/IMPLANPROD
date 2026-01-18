VERSION 5.00
Begin VB.Form OPCHEQ04 
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Configuración Gestión de Valores"
   ClientHeight    =   7995
   ClientLeft      =   45
   ClientTop       =   585
   ClientWidth     =   5790
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
   ScaleHeight     =   7995
   ScaleWidth      =   5790
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture16 
      Height          =   4100
      Left            =   4800
      ScaleHeight     =   4035
      ScaleWidth      =   1140
      TabIndex        =   42
      Top             =   0
      Visible         =   0   'False
      Width           =   1200
      Begin VB.Image Image3 
         Height          =   4150
         Left            =   0
         Picture         =   "OPCHEQ04.frx":0000
         Stretch         =   -1  'True
         Top             =   -120
         Width           =   1095
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Formularios de Impresión"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3675
      Left            =   240
      TabIndex        =   21
      Top             =   4200
      Width           =   5370
      Begin VB.CommandButton Command8 
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
         Left            =   4200
         TabIndex        =   41
         Top             =   2835
         Width           =   175
      End
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
         Height          =   630
         Left            =   4515
         Picture         =   "OPCHEQ04.frx":2398
         Style           =   1  'Graphical
         TabIndex        =   39
         Top             =   2835
         Width           =   650
      End
      Begin VB.TextBox Text4 
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
         Left            =   3015
         TabIndex        =   38
         Text            =   " "
         Top             =   2835
         Width           =   1215
      End
      Begin VB.TextBox Text3 
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
         Height          =   255
         Left            =   180
         TabIndex        =   37
         TabStop         =   0   'False
         Text            =   " "
         Top             =   3225
         Width           =   4215
      End
      Begin VB.TextBox Text7 
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
         Height          =   255
         Left            =   180
         TabIndex        =   35
         TabStop         =   0   'False
         Text            =   " "
         Top             =   2385
         Width           =   4215
      End
      Begin VB.TextBox Text10 
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
         Left            =   3015
         TabIndex        =   34
         Text            =   " "
         Top             =   1995
         Width           =   1215
      End
      Begin VB.CommandButton Command16 
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
         Left            =   4215
         TabIndex        =   33
         Top             =   1995
         Width           =   175
      End
      Begin VB.CommandButton Command17 
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
         Height          =   630
         Left            =   4515
         Picture         =   "OPCHEQ04.frx":26A2
         Style           =   1  'Graphical
         TabIndex        =   32
         Top             =   1995
         Width           =   650
      End
      Begin VB.CommandButton Command6 
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
         Height          =   630
         Index           =   2
         Left            =   4515
         Picture         =   "OPCHEQ04.frx":29AC
         Style           =   1  'Graphical
         TabIndex        =   31
         Top             =   1220
         Width           =   650
      End
      Begin VB.CommandButton Command6 
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
         Height          =   630
         Index           =   0
         Left            =   4515
         Picture         =   "OPCHEQ04.frx":2CB6
         Style           =   1  'Graphical
         TabIndex        =   30
         Top             =   420
         Width           =   650
      End
      Begin VB.CommandButton Command4 
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
         Left            =   4215
         TabIndex        =   27
         Top             =   1215
         Width           =   175
      End
      Begin VB.TextBox Text2 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3015
         TabIndex        =   26
         Top             =   1200
         Width           =   1215
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3000
         TabIndex        =   25
         Top             =   400
         Width           =   1215
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
         Height          =   285
         Left            =   4200
         TabIndex        =   22
         Top             =   420
         Width           =   175
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Movimientos de Cheques:"
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
         Left            =   0
         TabIndex        =   40
         Top             =   2880
         Width           =   2955
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Compensación de Cuentas:"
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
         Left            =   0
         TabIndex        =   36
         Top             =   2040
         Width           =   2955
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Déposito Bancario:"
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
         Left            =   525
         TabIndex        =   29
         Top             =   1275
         Width           =   2325
      End
      Begin VB.Label Label2 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   180
         TabIndex        =   28
         Top             =   1590
         Width           =   4215
      End
      Begin VB.Label Label3 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   165
         TabIndex        =   24
         Top             =   795
         Width           =   4215
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Transferencia de Fondos:"
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
         Left            =   420
         TabIndex        =   23
         Top             =   480
         Width           =   2430
      End
   End
   Begin VB.PictureBox Picture7 
      BackColor       =   &H00404000&
      Height          =   4080
      Left            =   4830
      ScaleHeight     =   4020
      ScaleWidth      =   1140
      TabIndex        =   13
      Top             =   0
      Width           =   1200
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
         Height          =   560
         Left            =   140
         Picture         =   "OPCHEQ04.frx":2FC0
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   1575
         Width           =   650
      End
      Begin VB.CommandButton Command6 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   560
         Index           =   1
         Left            =   140
         Picture         =   "OPCHEQ04.frx":32CA
         Style           =   1  'Graphical
         TabIndex        =   19
         ToolTipText     =   "Acceso a Tabla de Cuentas Bancarias"
         Top             =   2685
         Width           =   650
      End
      Begin VB.CommandButton Command11 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   560
         Left            =   140
         Picture         =   "OPCHEQ04.frx":35D4
         Style           =   1  'Graphical
         TabIndex        =   18
         ToolTipText     =   "Archivo Maestro de Cuentas Contables"
         Top             =   2130
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
         Height          =   560
         Left            =   140
         Picture         =   "OPCHEQ04.frx":371E
         Style           =   1  'Graphical
         TabIndex        =   17
         ToolTipText     =   "Terminar - Salir de la Aplicación"
         Top             =   195
         Width           =   650
      End
      Begin VB.CommandButton Command7 
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
         Height          =   560
         Left            =   140
         MaskColor       =   &H00404000&
         Picture         =   "OPCHEQ04.frx":3868
         Style           =   1  'Graphical
         TabIndex        =   16
         ToolTipText     =   "Ayuda FLEXOFT en Línea"
         Top             =   750
         Width           =   650
      End
      Begin VB.PictureBox MARCOBARRA 
         BackColor       =   &H000000FF&
         Height          =   135
         Left            =   140
         ScaleHeight     =   75
         ScaleWidth      =   585
         TabIndex        =   14
         Top             =   3375
         Width           =   640
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H00FF0000&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   0
            TabIndex        =   15
            Top             =   0
            Width           =   12000
         End
      End
      Begin VB.Image Image1 
         Height          =   390
         Left            =   -300
         Picture         =   "OPCHEQ04.frx":3B72
         Top             =   3630
         Width           =   1515
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Cheques Rechazados"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2640
      Left            =   240
      TabIndex        =   1
      Top             =   1440
      Width           =   4425
      Begin VB.CommandButton Command14 
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
         Left            =   3990
         TabIndex        =   5
         Top             =   525
         Width           =   175
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Procesamiento"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1065
         Left            =   210
         TabIndex        =   2
         Top             =   1365
         Width           =   4005
         Begin VB.OptionButton Option3 
            BackColor       =   &H00C0FFFF&
            Caption         =   "Por Ajuste de Saldo Deudor"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   450
            Left            =   210
            TabIndex        =   4
            Top             =   420
            Value           =   -1  'True
            Width           =   1800
         End
         Begin VB.OptionButton Option4 
            BackColor       =   &H00C0FFFF&
            Caption         =   "Por Nota de Debito al Cliente"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   450
            Left            =   1995
            TabIndex        =   3
            Top             =   420
            Width           =   1905
         End
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Cuenta Cheques Rechazados:"
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
         Left            =   105
         TabIndex        =   8
         Top             =   420
         Width           =   2325
      End
      Begin VB.Label Label11 
         BackColor       =   &H00E0E0E0&
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
         Height          =   285
         Left            =   2520
         TabIndex        =   7
         Top             =   525
         Width           =   1485
      End
      Begin VB.Label Label12 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   375
         TabIndex        =   6
         Top             =   900
         Width           =   3795
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Gestion de Valores"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1275
      Left            =   210
      TabIndex        =   0
      Top             =   120
      Width           =   4425
      Begin VB.CommandButton Command13 
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
         Left            =   3990
         TabIndex        =   9
         Top             =   420
         Width           =   175
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Cartera de Valores:"
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
         Left            =   105
         TabIndex        =   12
         Top             =   480
         Width           =   2325
      End
      Begin VB.Label Label8 
         BackColor       =   &H00E0E0E0&
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
         Height          =   285
         Left            =   2520
         TabIndex        =   11
         Top             =   420
         Width           =   1485
      End
      Begin VB.Label Label9 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   375
         TabIndex        =   10
         Top             =   795
         Width           =   3795
      End
   End
   Begin VB.Menu menu 
      Caption         =   "Menú"
      Begin VB.Menu grabar 
         Caption         =   "Grabar"
      End
      Begin VB.Menu s 
         Caption         =   "-"
      End
      Begin VB.Menu salir 
         Caption         =   "Salir"
      End
   End
   Begin VB.Menu tabla 
      Caption         =   "Tablas"
      Begin VB.Menu cuentascontables 
         Caption         =   "Acceso a Tabla de Cuentas Contables"
      End
      Begin VB.Menu cuentasbancarias 
         Caption         =   "Acceso a Tabla de Cuentas Bancarias"
      End
   End
End
Attribute VB_Name = "OPCHEQ04"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Call CIERRARCH("*.*")
    Hide
End Sub

Private Sub Command11_Click()
   Command11.Enabled = False
   Call CARCUEINGAS
   Command11.Enabled = True
End Sub

Private Sub Command14_Click()
    Call SELECHO("ECUECON")
    VDEVU$ = VALACT1$("ECUECON")
    If VDEVU$ <> "" Then
        Label11 = VDEVU$
        Label12 = VALACT2$("ECUECON")
    End If
End Sub
'
Private Sub Command13_Click()
    Call SELECHO("ECUECON")
    VDEVU$ = VALACT1$("ECUECON")
    If VDEVU$ <> "" Then
        Label8 = VDEVU$
        Label9 = VALACT2$("ECUECON")
    End If
End Sub


Private Sub Command17_Click()
    ACONEC$ = "FLFORMS": If EXISTE%(ACONEC$ + ".EXE") < 1 Then ACONEC$ = "PRIFORMS"
    Call CONECRUN(ACONEC$ + "/" + TRIM$(Text10.TEXT), "Formularios de Impresion")
    If EXISTE%(ACONEC$ + ".EXE") < 1 Then
        Call FINAL("")
    End If
End Sub

'
Private Sub Command2_Click()
    '
    Call GRACONTROL("", "CUECHERE", Label11)
    '\\\OT-05-0440-WAR-29/06/05///
    Call GRACONTROL("", "FORTRAFO", Text1)
    '\\\OT-05-0440-WAR-FIN///
    '\\\OT-05-0475-WAR-07/07/05///
    Call GRACONTROL("", "FORDEPBA", Text2)
    '\\\OT-05-0475-WAR-FIN///
    MONUME$ = "AJUSALDO"
    If Option4.Value = True Then MONUME$ = "NOTADEBI"
    Call GRACONTROL("", "MOVICHEQ", Text4)
    '
    Call CIERRARCH("*.*")
    '\\\OT-08-0633-FL-14/08/08///
    'Call GRACONTROL("", "CUECHECA", Label8)
    Call VERICAMCAR(Label8.Caption)
    '\\\OT-08-0633-FL-FIN///
    Call FINAL("")
    '
End Sub

Sub VERICAMCAR(CTANUE$)
   '
   CTAANT$ = CONTROL$("", "CUECHECA")
   If CTAANT$ <> CTANUE$ Then
     If DERACCE%("CAMCUECA", "Modificar Cuenta Cartera de Valores") > 1 Then
       OPX% = COMALTER%("El Cambio de la Cuenta Asociada a la Cartera\Puede Alterar el Funcionamiento del Subsistema.\   \¿ Está Seguro de Cambiarla ?\\No, Cancelar el Cambio\Si, Registrar Nueva Cuenta")
       If OPX% = 2 Then
         Call GRACONTROL("", "CUECHECA", CTANUE$)
         Call MENSERR(24, "Se Ha Modificado la Cuenta 'Cartera de Valores'\de la Cuenta Nro " + TRIM$(CTAANT$) + " " + TRIM$(ECOARCH$("ECUECON", CTAANT$)) + "\por la Cuenta Nro " + TRIM$(CTANUE$) + " " + TRIM$(ECOARCH$("ECUECON", CTANUE$)) + "\  \Re-Inicie Ahora la Aplicación.")
       End If
     End If
   End If
   '
End Sub

Private Sub Command3_Click()
    '\\\OT-05-0440-WAR-29/06/05///
    Call SELECHO("INDIFRM")
    VDEVU$ = VALACT1$("INDIFRM")
    If VDEVU$ <> "" Then
        Text1 = VDEVU$
        Label3 = VALACT2$("INDIFRM")
    End If
    '\\\OT-05-0440-WAR-FIN///
End Sub
'

Private Sub Command4_Click()
    '\\\OT-05-0475-WAR-07/07/05///
    Call SELECHO("INDIFRM")
    VDEVU$ = VALACT1$("INDIFRM")
    If VDEVU$ <> "" Then
        Text2 = VDEVU$
        Label2 = VALACT2$("INDIFRM")
    End If
    '\\\OT-05-0475-WAR-FIN///
End Sub

Private Sub Command5_Click()
    ACONEC$ = "FLFORMS": If EXISTE%(ACONEC$ + ".EXE") < 1 Then ACONEC$ = "PRIFORMS"
    Call CONECRUN(ACONEC$ + "/" + TRIM$(Text4.TEXT), "Formularios de Impresion")
    If EXISTE%(ACONEC$ + ".EXE") < 1 Then
        Call FINAL("")
    End If

End Sub

Private Sub Command6_Click(Index As Integer)
   '
   Command6(Index).Enabled = False
   Select Case Index
     Case 0
       Call CONECRUN("FLFORMS/" + TRIM$(Text1.TEXT), "Formularios de Impresion")
     Case 1
       '\\\OT-5-661-FG-06/10/05///
       X% = ALTERNA%("Tabla de Cuentas Bancarias\Tabla de Cuentas de Tesorería")
       If X% = 1 Then
           Call CARCUEBANC
         '\\\OT-06-0570-RG-28/12/06///
         ElseIf X% = 2 Then
         '\\\OT-06-0570-RG-FIN///
           Call CARCUETES
       End If
       '\\\OT-5-661-FG-FIN///
     Case 2
       Call CONECRUN("FLFORMS/" + TRIM$(Text2.TEXT), "Formularios de Impresion")
   End Select
   Command6(Index).Enabled = True
   '
End Sub

Private Sub Command7_Click()
   Command7.Enabled = False
   Call HELPONLINE("CUECONTA")
   Command7.Enabled = True
End Sub

Private Sub Command8_Click()
    Call SELECHO("INDIFRM")
    VDEVU$ = VALACT1$("INDIFRM")
    If VDEVU$ <> "" Then
        Text4.TEXT = VDEVU$
        Text3.TEXT = VALACT2$("INDIFRM")
    End If

End Sub

Private Sub cuentasbancarias_Click()
    Call Command6_Click(1)
End Sub

Private Sub cuentascontables_Click()
   Call Command11_Click
End Sub

Private Sub Form_Load()
    If CONTROL("NEWASPEC", "CAMBIOAS") = "SI" Then
      Picture16.Visible = True
      Picture16.ZOrder 0
      Me.BackColor = &H8000000F
      Me.Frame1.BackColor = &H8000000F
      Me.Frame2.BackColor = &H8000000F
      Me.Frame3.BackColor = &H8000000F
      Me.Frame4.BackColor = &H8000000F
      Me.Option3.BackColor = &H8000000F
      Me.Option4.BackColor = &H8000000F
      MARCOBARRA.Top = 4800
      Set MARCOBARRA.Container = Me.Picture16
      MARCOBARRA.ZOrder 0
    End If
    '
    Label8 = CONTROL$("", "CUECHECA")
    Label9 = ECOARCH$("ECUECON", Label8)
    Label11 = CONTROL$("", "CUECHERE")
    Label12 = ECOARCH$("ECUECON", Label11)
    '\\\OT-05-0440-WAR-29/06/05///
    Text1 = CONTROL$("", "FORTRAFO")
    Label3 = ECOARCH$("INDIFRM", Text1)
    '\\\OT-05-0440-WAR-FIN///
    '\\\OT-05-0475-WAR-07/07/05///
    Text2 = CONTROL$("", "FORDEPBA")
    Label2 = ECOARCH$("INDIFRM", Text2)
    '\\\OT-05-0475-WAR-07/07/05///
    '\\\OT-06-0144-WAR-17/04/06///
    Text10.TEXT = CONTROL$("", "FORCOMPE")
    Text7.TEXT = ECOARCH$("INDIFRM", Text10.TEXT)
    Text4.TEXT = CONTROL$("", "MOVICHEQ")
    Text3.TEXT = ECOARCH$("INDIFRM", Text4.TEXT)
    
    '\\\OT-06-0144-WAR-FIN///
    '
    If CONTROL$("", "CHERECHA") = "NOTADEBI" Then Option4.Value = True
    '
End Sub



Private Sub grabar_Click()
    Command2_Click
End Sub

Private Sub salir_Click()
   Call Command1_Click
End Sub

Private Sub Text1_Change()
    DECUE$ = ECOARCH("INDIFRM", TRIM$(Text1))
    If DECUE$ <> "" Then
        Label3 = DECUE$
      Else
        Label3 = ""
    End If
End Sub
Private Sub Text1_DblClick()
    Call SELECHO("INDIFRM")
    VDEVU$ = VALACT1$("INDIFRM")
    If VDEVU$ <> "" Then
        Text1.TEXT = VDEVU$
        Text2.TEXT = VALACT2$("INDIFRM")
    End If
End Sub

Private Sub Text2_DblClick()
    '\\\OT-06-0570-RG-28/12/06///
    Call Command4_Click
    '\\\OT-06-0570-RG-FIN///
End Sub

Private Sub Text7_CLICK()
  Text10.SetFocus
End Sub

Private Sub Text10_DblClick()
    Call SELECHO("INDIFRM")
    VDEVU$ = VALACT1$("INDIFRM")
    If VDEVU$ <> "" Then
        Text10.TEXT = VDEVU$
        Text7.TEXT = VALACT2$("INDIFRM")
    End If
End Sub

Private Sub Command16_Click()
    Call SELECHO("INDIFRM")
    VDEVU$ = VALACT1$("INDIFRM")
    If VDEVU$ <> "" Then
        Text10.TEXT = VDEVU$
        Text7.TEXT = VALACT2$("INDIFRM")
    End If
End Sub

Private Sub Text2_Change()
    '\\\OT-05-0475-WAR-07/07/05///
    DECUE$ = ECOARCH("INDIFRM", TRIM$(Text2))
    If DECUE$ <> "" Then
        Label2 = DECUE$
      Else
        Label2 = ""
    End If
    '\\\OT-05-0475-WAR-FIN///
End Sub

Private Sub Text3_DBLCLICK()
    Call SELECHO("ECUECON")
    VDEVU$ = VALACT1$("ECUECON")
    If VDEVU$ <> "" Then
        Text3.TEXT = VDEVU$
        Text4.TEXT = VALACT2$("ECUECON")
    End If
End Sub

Private Sub Label11_DblClick()
    Call SELECHO("ECUECON")
    VDEVU$ = VALACT1$("ECUECON")
    If VDEVU$ <> "" Then
        Label11 = VDEVU$
        Label12 = VALACT2$("ECUECON")
    End If
End Sub

Private Sub Label8_DblClick()
    Call SELECHO("ECUECON")
    VDEVU$ = VALACT1$("ECUECON")
    If VDEVU$ <> "" Then
        Label8 = VDEVU$
        Label9 = VALACT2$("ECUECON")
    End If
End Sub
'\\\OT-5-661-FG-06/10/05///
Sub CARCUETES()
   '
   JJ& = 0
   NUCUEN% = ULTREG&("CUENTAS")
   Screen.MousePointer = 11
   For U& = 1 To ULTREG&("TACUECAJ")
      XX$ = REGLEIDO$("TACUECAJ", U&)
      CXI% = CUEINT%(TRIM$(Left$(XX$, 12)))
      If CXI% > 0 Then
         If CXI% <= NUCUEN% Then
            JJ& = JJ& + 1
            Call GRAREG("!TACUECAJ", XX$, JJ&)
         End If
      End If
   Next U&
   Call SETULTREG("!TACUECAJ", JJ&)
   Call CIERRARCH("!TACUECAJ")
   Call CIERRARCH("TACUECAJ")
   Screen.MousePointer = 1
   '
   VTB% = VENTABU%("TACUECAJ")
   If VTB% >= 0 Then
      Screen.MousePointer = 11
      JJ& = 0
      For U& = 1 To ULTREG&("!TACUECAJ")
         XX$ = REGLEIDO$("!TACUECAJ", U&)
         CXI% = CUEINT%(TRIM$(Left$(XX$, 12)))
         If CXI% > 0 Then
            If CXI% <= NUCUEN% Then
               JJ& = JJ& + 1
               Call GRAREG("TACUECAJ", XX$, JJ&)
            End If
         End If
      Next U&
      Call SETULTREG("TACUECAJ", JJ&)
      Call CIERRARCH("TACUECAJ")
      Call FRESHECHO("TACUECAJ")
      Call CIERRARCH("!TACUECAJ")
      Screen.MousePointer = 1
   End If
   '
End Sub
'\\\OT-5-661-FG-FIN///
