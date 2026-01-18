VERSION 5.00
Begin VB.Form CALIPROV09 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0FFE0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Sistema de Calificación de Proveedores"
   ClientHeight    =   4065
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   7350
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
   ScaleHeight     =   4065
   ScaleWidth      =   7350
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00004000&
      Height          =   950
      Left            =   4170
      ScaleHeight     =   885
      ScaleWidth      =   2865
      TabIndex        =   34
      Top             =   450
      Width           =   2925
      Begin VB.CommandButton Command2 
         Caption         =   "Setup"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   900
         Left            =   900
         Picture         =   "CALIPROV09.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   39
         ToolTipText     =   "Configuración de Funcionamiento"
         Top             =   0
         Width           =   645
      End
      Begin VB.CommandButton Command7 
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
         Height          =   900
         Left            =   1545
         Picture         =   "CALIPROV09.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   38
         ToolTipText     =   "Ayuda FLEXOFT en Línea"
         Top             =   0
         Width           =   645
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
         Height          =   900
         Left            =   2190
         Picture         =   "CALIPROV09.frx":074C
         Style           =   1  'Graphical
         TabIndex        =   37
         ToolTipText     =   "Abandona la Aplicación"
         Top             =   0
         Width           =   645
      End
      Begin VB.PictureBox MARCOBARRA 
         BackColor       =   &H000000FF&
         Height          =   135
         Left            =   150
         ScaleHeight     =   75
         ScaleWidth      =   540
         TabIndex        =   35
         Top             =   200
         Width           =   600
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H00FF0000&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   0
            TabIndex        =   36
            Top             =   0
            Width           =   12000
         End
      End
      Begin VB.Image Image1 
         Height          =   390
         Left            =   -280
         Picture         =   "CALIPROV09.frx":0896
         Top             =   500
         Width           =   1515
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0FFE0&
      Caption         =   "Ponderación"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1065
      Left            =   510
      TabIndex        =   27
      Top             =   5670
      Width           =   2115
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
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
         Index           =   0
         Left            =   1575
         TabIndex        =   32
         Text            =   "20"
         Top             =   630
         Width           =   435
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
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
         Index           =   1
         Left            =   105
         MaxLength       =   2
         TabIndex        =   31
         Text            =   "60"
         ToolTipText     =   "% Por Rechazos"
         Top             =   315
         Width           =   330
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
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
         Index           =   2
         Left            =   630
         MaxLength       =   2
         TabIndex        =   30
         Text            =   "20"
         ToolTipText     =   "% Por Entrega de Certificados"
         Top             =   315
         Width           =   330
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
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
         Index           =   3
         Left            =   1155
         MaxLength       =   2
         TabIndex        =   29
         Text            =   "5"
         ToolTipText     =   "% Por Estado"
         Top             =   315
         Width           =   330
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
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
         Index           =   4
         Left            =   1680
         MaxLength       =   2
         TabIndex        =   28
         Text            =   "15"
         ToolTipText     =   "% Por Atraso / Cumplimiento de Entregas"
         Top             =   315
         Width           =   330
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "% Límite Rechazos:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   105
         TabIndex        =   33
         Top             =   660
         Width           =   1380
      End
   End
   Begin VB.OptionButton Option31 
      BackColor       =   &H00C0FFE0&
      Caption         =   "Analítico de Calidad "
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
      Left            =   2895
      TabIndex        =   26
      Top             =   5805
      Width           =   2115
   End
   Begin VB.OptionButton Option41 
      BackColor       =   &H00C0FFE0&
      Caption         =   "Promedio de Calidad"
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
      Left            =   2895
      TabIndex        =   25
      Top             =   6060
      Width           =   2115
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0FFE0&
      Caption         =   "Consultas y Listados"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3885
      Left            =   120
      TabIndex        =   1
      Top             =   45
      Width           =   7110
      Begin VB.OptionButton Option4 
         BackColor       =   &H00C0FFE0&
         Caption         =   "Calificacion Integral"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4695
         TabIndex        =   42
         Top             =   3360
         Value           =   -1  'True
         Width           =   2340
      End
      Begin VB.OptionButton Option3 
         BackColor       =   &H00C0FFE0&
         Caption         =   "Calif. Calidad"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   4710
         TabIndex        =   41
         Top             =   2340
         Width           =   2115
      End
      Begin VB.OptionButton Option6 
         BackColor       =   &H00C0FFE0&
         Caption         =   "Calif. Cumplimiento"
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
         Left            =   4710
         TabIndex        =   40
         Top             =   2040
         Width           =   2115
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Emitir Listado"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   4605
         Picture         =   "CALIPROV09.frx":27B8
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   2670
         Width           =   2115
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00C0FFE0&
         Caption         =   "Indice de Calidad"
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
         Left            =   4695
         TabIndex        =   22
         Top             =   1740
         Width           =   2115
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00C0FFE0&
         Caption         =   "Indice Cumplimiento"
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
         Left            =   4695
         TabIndex        =   21
         Top             =   1440
         Width           =   2115
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00C0FFE0&
         Caption         =   "Período"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1185
         Left            =   210
         TabIndex        =   14
         Top             =   2520
         Width           =   3900
         Begin VB.CommandButton Command5 
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
            Left            =   1680
            TabIndex        =   16
            Top             =   615
            Width           =   175
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
            Left            =   3465
            TabIndex        =   15
            Top             =   615
            Width           =   175
         End
         Begin VB.Label Label9 
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " "
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
            TabIndex        =   20
            ToolTipText     =   "Doble Click Para Editar Manualmente"
            Top             =   615
            Width           =   975
         End
         Begin VB.Label Label4 
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
            Left            =   735
            TabIndex        =   19
            ToolTipText     =   "Doble Click Para Editar Manualmente"
            Top             =   615
            Width           =   975
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Desde:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   -1260
            TabIndex        =   18
            Top             =   315
            Width           =   2325
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Hasta:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   420
            TabIndex        =   17
            Top             =   315
            Width           =   2325
         End
      End
      Begin VB.Frame Frame5 
         BackColor       =   &H00C0FFE0&
         Caption         =   "Rango de Proveedores"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2115
         Left            =   210
         TabIndex        =   5
         Top             =   315
         Width           =   3900
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
            Left            =   3465
            TabIndex        =   9
            Top             =   1260
            Width           =   175
         End
         Begin VB.TextBox Text8 
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
            Left            =   2625
            TabIndex        =   8
            Text            =   " "
            Top             =   1260
            Width           =   855
         End
         Begin VB.CommandButton Command9 
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
            Left            =   3465
            TabIndex        =   7
            Top             =   360
            Width           =   175
         End
         Begin VB.TextBox Text10 
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
            Left            =   2625
            TabIndex        =   6
            Top             =   360
            Width           =   855
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
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
            Left            =   210
            TabIndex        =   13
            Top             =   1305
            Width           =   2325
         End
         Begin VB.Label Label6 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
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
            Left            =   210
            TabIndex        =   12
            Top             =   405
            Width           =   2325
         End
         Begin VB.Label Label7 
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " "
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
            Left            =   210
            TabIndex        =   11
            Top             =   720
            Width           =   3450
         End
         Begin VB.Label Label8 
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   1  'Fixed Single
            Caption         =   " "
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
            Left            =   210
            TabIndex        =   10
            Top             =   1620
            Width           =   3450
         End
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00C0FFE0&
      Caption         =   "Informes"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1060
      Left            =   510
      TabIndex        =   0
      Top             =   4605
      Width           =   3690
      Begin VB.PictureBox Picture1 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   840
         ScaleHeight     =   585
         ScaleWidth      =   1950
         TabIndex        =   2
         Top             =   285
         Width           =   2010
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
            Left            =   0
            Picture         =   "CALIPROV09.frx":2E22
            Style           =   1  'Graphical
            TabIndex        =   3
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Informes de Recepción"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   540
            Left            =   735
            TabIndex        =   4
            Top             =   105
            Width           =   1590
         End
      End
   End
   Begin VB.Label Label13 
      Caption         =   "5"
      Height          =   330
      Left            =   0
      TabIndex        =   24
      Top             =   0
      Visible         =   0   'False
      Width           =   540
   End
End
Attribute VB_Name = "CALIPROV09"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim DESFECHA%, HASFECHA%
Dim PERIO$
Dim REOCO&(32767), PUBOL$(32767), CAPEN(32767)

Sub CALICUMPLI()
   '
   Call ICUMPLI
   '
   'Comienza calculo de cumplimiento
   '
   Dim QtEnPlazo()
   Dim QtAtras()
   ReDim QtEnPlazo(1)
   ReDim QtAtras(1)
   '
   Ultimo& = 0
   ANT% = 0
   For U& = 1 To ULTREG("COCOBREC")
     x$ = REGLEIDO("COCOBREC", U&)
     PROV% = CVI(Mid(x$, 1, 2))
     CANTREC# = CVS(Mid(x$, 15, 4))
     FECHRECI = CVI(Mid(x$, 19, 2))
     FECHPROG = CVI(Mid(x$, 13, 2))
     '
     If PROV% > ANT% Then
        ReDim Preserve QtEnPlazo(PROV%)
        ReDim Preserve QtAtras(PROV%)
        ANT% = PROV%
     End If
     '
     If FECHRECI <= FECHPROG Then
        QtEnPlazo(PROV%) = QtEnPlazo(PROV%) + CANTREC#
       Else
        QtAtras(PROV%) = QtAtras(PROV%) + CANTREC#
     End If
     '
     If PROV% > Ultimo& Then Ultimo& = PROV%
   Next U&
    
    'AGREGAR AL ARCHIVO LOS RESULTADOS
    REGIGRA& = 0
    Call SETULTREG("CALIPROV", REGIGRA&)
    Call CIERRARCH("CALIPROV")
    
    For U& = 1 To Ultimo&
      TOTAL# = QtEnPlazo(U&) + QtAtras(U&)
      If TOTAL# > 0 Then
         PORC# = (QtEnPlazo(U&) / TOTAL#) * 100
         '
         REGIGRA& = REGIGRA& + 1
         Call REPLA(x$, MKI(U&), 1, 2)
         Call REPLA(x$, MKD(TOTAL#), 3, 8)
         Call REPLA(x$, MKD(XVALO(QtEnPlazo(U&))), 11, 8)
         Call REPLA(x$, MKD(XVALO(QtAtras(U&))), 19, 8)
         Call REPLA(x$, MKS(PORC#), 27, 4)
         Call GRAREG("CALIPROV", x$, REGIGRA&)
         '
      End If
    Next U&
   
   Call SETULTREG("CALIPROV", REGIGRA&)
   Call CIERRARCH("CALIPROV")
   '
   If Option4.Value = True Then GoTo 99 'Calificacion integral
   '
   Call FILESORT("CALIPROV", "", 1, 2, "ASC")
   Call FINAL("*LICALPRO")
   '
99 Screen.MousePointer = 1
End Sub

Sub CALINTEG()
   '
   Dim CALIDAD()
   ReDim CALIDAD(1)
   '
   Dim CUMPLI()
   ReDim CUMPLI(1)
   
   'Se lee PRIMERO la calidad ya que todos tienen calidad
   'se asigna por defecto 100% de cumplimiento
   'luego se lee el cumplimiento, en caso de tener cumplimiento se asigna
   'el valor leido,
   'Esto se hace para evitar que aparezcan sin cumplimiento los registros sin OC
   '
   Call CALICUMPLI
   Call CALICALIDAD
   '
   'OBTENGO PORC. CALIDAD
   MAYOR% = 0
   For U& = 1 To ULTREG("CALICALI")
     x$ = REGLEIDO("CALICALI", U&)
     PROV% = CVI(Mid(x$, 1, 2))
     PORC# = XVALO(CVS(Mid(x$, 27, 4)))
     If PROV% > MAYOR% Then
        ReDim Preserve CALIDAD(PROV%)
        ReDim Preserve CUMPLI(PROV%)
        MAYOR% = PROV%
     End If
     CALIDAD(PROV%) = PORC#
     CUMPLI(PROV%) = 100
   Next U&
   '
   ReDim Preserve CUMPLI(MAYOR%)
   ReDim Preserve CALIDAD(MAYOR%)
   '
   'OBTENGO PORC. CUMPLIMIENTO
  
   For U& = 1 To ULTREG("CALIPROV")
     x$ = REGLEIDO("CALIPROV", U&)
     PROV% = CVI(Mid(x$, 1, 2))
     PORC# = XVALO(CVS(Mid(x$, 27, 4)))
     If PROV% > MAYOR% Then
        ReDim Preserve CUMPLI(PROV%)
        MAYOR% = PROV%
     End If
     '
     If PORC# > 0 Then
        CUMPLI(PROV%) = PORC#
     End If
     
   Next U&
   '
   'ATAJAR ERROR DE INDICE FUERA DE RANGO
   ReDim Preserve CUMPLI(MAYOR%)
   ReDim Preserve CALIDAD(MAYOR%)
   '
   'GRABAR LOS PORC. EN ARCHIVO CALINTE
   PondeCali# = XVALO(CONTROL("CALIPROV", "PONDECAL"))
   PondeCumpli# = XVALO(CONTROL("CALIPROV", "PONDECUM"))

   For AX& = 1 To MAYOR%
     If CALIDAD(AX&) > 0 Or CUMPLI(AX&) > 0 Then
        INDICE# = ((CALIDAD(AX&) * PondeCali#) / 100) + ((CUMPLI(AX&) * PondeCumpli#) / 100)
        '
        INDICE# = INDICE# / 100
        CUM# = XVALO(CUMPLI(AX&)) / 100
        CAL# = XVALO(CALIDAD(AX&)) / 100
        '
        REGIGRA& = REGIGRA& + 1
        Call REPLA(x$, MKI(AX&), 1, 2)
        Call REPLA(x$, MKS(CUM#), 3, 4)
        Call REPLA(x$, MKS(CAL#), 7, 4)
        Call REPLA(x$, MKS(INDICE#), 11, 4)
        Clse$ = CLASE$(INDICE#)
        '
        Call REPLA(x$, Clse$, 15, 1)
        Call GRAREG("CALINTE", x$, REGIGRA&)
        Call ACTUREGISTRO("PROVED12", "CALINTEG", "Nume_Cli = " & AX&, INDICE#, REG&)
        '
     End If
   Next AX&
   '
   Call SETULTREG("CALINTE", REGIGRA&)
   Call CIERRARCH("CALINTE")
   '
   Call FILESORT("CALINTE", "", 1, 2, "ASC")
   Call FINAL("*LICALINT")
   '
99 Screen.MousePointer = 1
   
End Sub
'
Private Sub Command1_Click()
    Call CIERRARCH("*.*")
    Call FINAL("")
End Sub
'
Private Sub Command2_Click()
   Command2.Enabled = False
   If DERACCE%("ACTICAL", "Actualización Indices de Calificación") >= 2 Then
     STPCALPR09.Show 1
   End If
   Command2.Enabled = True
End Sub
'

Private Sub Command3_Click()
   Command3.Enabled = False
   '
10 Call SELECHO("INDIBOL")
   NBR$ = VALACT1$("INDIBOL")   ' OBJPLA$("SELBOLET", "")
   If NBR$ <> "" Then
     NROBOL& = Val(NBR$)
     If NROBOL& > 0 Then
       ICALREC07.Label5 = TRIM$(Str$(NROBOL&))
       ICALREC07.Show 1
       GoTo 10
     End If
   End If
   '
   Command3.Enabled = True
End Sub

Private Sub Command4_Click()
   Call SELECHO("SELFECHA")
   If VALACT1$("SELFECHA") <> "" Then
     Label9 = VALACT1$("SELFECHA")
   End If
End Sub

Private Sub Command5_Click()
   Call SELECHO("SELFECHA")
   If VALACT1$("SELFECHA") <> "" Then
     Label4 = VALACT1$("SELFECHA")
   End If
End Sub

Private Sub Command6_Click()
   Command6.Enabled = False
   '
   DESFECHA% = FECHANUM(Label4)
   HASFECHA% = FECHANUM(Label9)
   If DESFECHA% < 1 Or HASFECHA% < 1 Or HASFECHA% < DESFECHA% Then
     Call MENSERR(24, "Rango de Fechas Incorrecto")
     GoTo 99
   End If
   '
   PERIO$ = "Periodo: " + Label4 + " - " + Label9
   For KKI% = 0 To 32767
     REOCO&(KKI%) = 0
     PUBOL$(KKI%) = ""
     CAPEN(KKI%) = 0
     If KKI% >= 32767 Then Exit For
   Next KKI%
   '
   If Option1.Value = True Then
       Call ICUMPLI
     ElseIf Option2.Value = True Then
       Call ICALID
     ElseIf Option3.Value = True Then
       Call CALICALIDAD
     ElseIf Option4.Value = True Then
        Call CALINTEG
     ElseIf Option6.Value = True Then
       Call CALICUMPLI
   End If
   '
99 Command6.Enabled = True
End Sub

Private Sub Command8_Click()
    Call SELECHO("PROVED1")
    NCMAX% = Val(VALACT1$("PROVED1"))
    If NCMAX% > 0 Then
      Text8.TEXT = TRIM$(Str$(NCMAX%))
    End If
End Sub

Private Sub Command9_Click()
    Call SELECHO("PROVED1")
    NCMIN% = Val(VALACT1$("PROVED1"))
    If NCMIN% > 0 Then
       Text10.TEXT = TRIM$(Str$(NCMIN%))
    End If
End Sub

Private Sub Form_Load()
    '
    Call VERJOBID(OKEY%)
    If OKEY% < 1 Then Call FINAL("")
    '
    EPAC$ = TRIM$(EMPREAC$)
    If TRIM$(APLINVO$) <> "" Then
      ICALREC0707.Label5 = TRIM$(APLINVO$)
      ICALREC07.Show 1
      Call CIERRARCH("*.*")
      Call BAJALDISCO
      Call FINAL("")
      Exit Sub
    End If
    
    If EPAC$ <> "" Then
       Caption = Caption + " - " + EPAC$
    End If
    '
    Show
    Refresh
    If ULTREG&("ACREDOR") > 0 Then
      Screen.MousePointer = 11
      NCMIN% = 32767: NCMAX% = 0
      FIN& = ULTREG&("ACREDOR")
      For U& = 1 To FIN&
        NCX% = CVI(Left$(REGLEIDO$("ACREDOR", U&), 2))
        If NCX% > 0 Then
          If NCX% > NCMAX% Then NCMAX% = NCX%
          If NCX% < NCMIN% Then NCMIN% = NCX%
        End If
      Next U&
      '
      Screen.MousePointer = 1
      Text10 = TRIM$(Str$(NCMIN%))
      Text8 = TRIM$(Str$(NCMAX%))
    End If
    '
    HAS% = HOY(HOS$)
    FEX = HAS%
    Label9 = FECHATEX$(FEX)
    Label4 = "01" + Mid$(Label9, 3)
    '
    If Val(CONTROL$("CALIPROV", "ESCACALI")) = 10 Then
      Label13 = "10"
    End If
    '
    PONDERA$ = TRIM$(CONTROL$("CALIPROV", "PONDERA"))
    If PONDERA$ <> "" Then
      For KKI% = 0 To 4
        Text1(KKI%) = Mid$(PONDERA$, 3 * KKI% + 1, 2)
      Next KKI%
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call CIERRARCH("*.*")
   Call FINAL("")
End Sub

Private Sub Label4_DblClick()
'codigo anterior
'   Call SELECHO("SELFECHA")
'   If VALACT1$("SELFECHA") <> "" Then
'     Label9 = VALACT1$("SELFECHA")
'   End If
'------------------------------------
     Call DESHASFECHA(DESDEX%, HASTAX%, PERIO$)
     If PERIO$ = "" Then Exit Sub
     '
     Label4 = FECHATEX$(DESDEX%)
     Label9 = FECHATEX$(HASTAX%)
End Sub

Private Sub Label9_DblClick()
'CODIGO ANTERIOR
'   Call SELECHO("SELFECHA")
'   If VALACT1$("SELFECHA") <> "" Then
'     Label9 = VALACT1$("SELFECHA")
'   End If

     Call DESHASFECHA(DESDEX%, HASTAX%, PERIO$)
     If PERIO$ = "" Then Exit Sub
     '
     Label4 = FECHATEX$(DESDEX%)
     Label9 = FECHATEX$(HASTAX%)

End Sub

Private Sub Text10_Change()
   Label7 = ""
   NCMIN% = Val(Text10.TEXT)
   If Val(NCMIN%) > 0 Then
     Label7 = RASOCLI$((-1) * NCMIN%)
   End If
End Sub

Private Sub Text10_DblClick()
    Call SELECHO("PROVED1")
    NCMIN% = Val(VALACT1$("PROVED1"))
    If NCMIN% > 0 Then
      Text10.TEXT = TRIM$(Str$(NCMIN%))
    End If
End Sub

Private Sub Text8_Change()
   Label8 = ""
   NCMAX% = Val(Text8.TEXT)
   If Val(NCMAX%) > 0 Then
     Label8 = RASOCLI$((-1) * NCMAX%)
   End If
End Sub

Private Sub Text8_DblClick()
    Call SELECHO("PROVED1")
    NCMAX% = Val(VALACT1$("PROVED1"))
    If NCMAX% > 0 Then
      Text8.TEXT = TRIM$(Str$(NCMAX%))
    End If
End Sub

Sub ICUMPLI()
   '
   ' INDICE DE CUMPLIMIENTO DE PROVEEDORES
   '
   Screen.MousePointer = 11
   '
10 J& = 0: PUCON$ = "": U& = 0
   'FIN& = ULTREG&("OCOMDETA")
   '\\\OT-08-0641-OD-25/08/08///
   DES = FECHATEX(DESFECHA%)
   HAS = FECHATEX(HASFECHA%)
   SQLX$ = "SELECT * FROM OCOMDETA WITH(NOLOCK)"
   SQLX$ = SQLX$ + " WHERE Stat_Ocom <> 9 "
   SQLX$ = SQLX$ + " AND Femi_Ocom >= '" & DES & "'"
   SQLX$ = SQLX$ + " AND Femi_Ocom <= '" & HAS & "'"
   Dim OrdenCompraTemp As ADODB.Recordset
   Set OrdenCompraTemp = FLEXCONN.Execute(FILTSQL$(SQLX$))
   With OrdenCompraTemp
      Do While Not (.EOF)
        U& = U& + 1
        J& = J& + 1
        REOCO&(J&) = U&
        CAPEN(J&) = XVALO(!Cant_Ocom) 'CVS(Mid$(X$, 55, 4))
        CLACO$ = MKI$(!Npro_Ocom) + MKS$(!Nume_Ocom) + MKI$(!COD_INTE)   'Mid$(X$, 7, 2) + Left$(X$, 4) + Mid$(X$, 9, 2)
        PUCON$ = PUCON$ + CLACO$
        .MoveNext
      Loop
   End With
   OrdenCompraTemp.Close
   Set OrdenCompraTemp = Nothing
   CAOCO& = J&
   SQLX$ = "SELECT * FROM BOLRECEP WITH(NOLOCK)"
   Dim BolRecepTmp As ADODB.Recordset
   Set BolRecepTmp = FLEXCONN.Execute(FILTSQL$(SQLX$))
   '
   With BolRecepTmp
      Do While Not (.EOF)
        U& = U& + 1
        J& = J& + 1
        CAREC = XVALO(!CantNom) '(Mid$(X$, 89, 8))
        CLACO$ = MKI$(!NumProv) + MKS$(!NumOC) + MKI$(!COD_INTE)
        For K& = 1 To CAOCO&
            If Mid$(PUCON$, 8 * K& - 7, 8) = CLACO$ Then
              UKA& = K&
              If CAPEN(K&) > 0 Then
                PUBOL$(K&) = PUBOL$(K&) + MKS$(U&)
                CAPEN(K&) = CAPEN(K&) - CAREC
                GoTo 19
              End If
            End If
        Next K&
        '
        If UKA& > 0 Then
          PUBOL$(UKA&) = PUBOL$(UKA&) + MKS$(U&)
          CAPEN(UKA&) = CAPEN(UKA&) - CAREC
        End If

19      .MoveNext
      Loop
   End With
   BolRecepTmp.Close
   Set BolRecepTmp = Nothing
   '
   If UKA& > 0 Then
     PUBOL$(UKA&) = PUBOL$(UKA&) + MKS$(U&)
     CAPEN(UKA&) = CAPEN(UKA&) - CAREC
   End If
   '
   HOII = HOY(HOS$)
   SQLX$ = "SELECT * FROM OCOMDETA WITH(NOLOCK)"
   SQLX$ = SQLX$ + " WHERE Stat_Ocom <> 9 "
   SQLX$ = SQLX$ + " AND Femi_Ocom >= '" & DES & "'"
   SQLX$ = SQLX$ + " AND Femi_Ocom <= '" & HAS & "'"
   '
   Z0$ = REGBLAN$("COCOBREC")
   '
   Dim OrdenCompTemp As ADODB.Recordset
   Set OrdenCompTemp = FLEXCONN.Execute(FILTSQL$(SQLX$))
   With OrdenCompTemp
      Do While Not (.EOF)
      
          Call REPLA(Z0$, MKI$(!Npro_Ocom), 1, 2)
          Call REPLA(Z0$, MKS$(!Nume_Ocom), 3, 4)
          Call REPLA(Z0$, MKI$(!COD_INTE), 7, 2)
          Call REPLA(Z0$, MKS$(!Cant_Ocom), 9, 4)
          FEXI% = CVINT(!Fentre_Sol)
          Call REPLA(Z0$, MKI$(FEXI%), 13, 2)
          FENTREPRO = CVINT(!Fentre_Sol)
          '
          For K& = 1 To CAOCO&
            CLACO$ = MKI$(!Npro_Ocom) + MKS$(!Nume_Ocom) + MKI$(!COD_INTE)
            If Mid$(PUCON$, 8 * K& - 7, 8) = CLACO$ Then Exit For
          Next K&
          '
          If Len(PUBOL$(K&)) < 4 Then
               DIATRA% = 0
               If FENTREPRO < HOII Then
                 DIATRA% = DIENTRE%(FENTREPRO, HOII)
               End If
               Call REPLA(Z0$, MKI$(DIATRA%), 37, 2)
               NCOCO& = NCOCO& + 1
               Call GRAREG("COCOBREC", Z0$, NCOCO&)
             Else
               'For II& = 1 To Len(PUBOL$(K&)) Step 4
                 Z$ = Z0$
                 'REBREC& = CVS(Mid$(PUBOL$(K&), II&, 4))
                 'Y$ = REGLEIDO$("BOLRECEP", REBREC&)
                 NumProveedor% = XVALO(!Npro_Ocom)
                 NumOrdCompra& = XVALO(!Nume_Ocom)
                 CodMaterial% = XVALO(!COD_INTE)
                 SQLX$ = "SELECT * FROM BOLRECEP WITH(NOLOCK)"
                 SQLX$ = SQLX$ + " WHERE NumProv = " & NumProveedor% & " "
                 SQLX$ = SQLX$ + " AND NumOC = " & NumOrdCompra& & " "
                 SQLX$ = SQLX$ + " AND Cod_Inte = " & CodMaterial% & ""
                 Dim BolRecTemp As ADODB.Recordset
                 Set BolRecTemp = FLEXCONN.Execute(FILTSQL$(SQLX$))
                 With BolRecTemp
                     Do While Not (.EOF)
                       Call REPLA(Z$, MKS$(XVALO(!CantNom)), 15, 4)
                       Call REPLA(Z$, MKI$(CVINT(!FechRecep)), 19, 2)
                       Call REPLA(Z$, SAFETEXT$(!NumRem_Prov), 21, 16)
                       FERECI = CVINT(!FechRecep)
                       DIATRA% = 0
                       If FENTREPRO < FERECI Then
                         DIATRA% = DIENTRE%(FENTREPRO, FERECI)
                       End If
                       Call REPLA(Z$, MKI$(DIATRA%), 37, 2)
                       NCOCO& = NCOCO& + 1
                       Call GRAREG("COCOBREC", Z$, NCOCO&)
                       .MoveNext
                     Loop
                 End With
                 BolRecTemp.Close
                 Set BolRecTemp = Nothing
    
          End If
          .MoveNext
      Loop
   End With
   OrdenCompTemp.Close
   Set OrdenCompTemp = Nothing
   '
   Call SETULTREG("COCOBREC", NCOCO&)
   Call CIERRARCH("COCOBREC")
   '
   If Option6.Value = True Then GoTo 99 ' CALIFICACION PROVEEDORES
   If Option4.Value = True Then GoTo 99 ' Calificacion de Calidad
   '
   Call HEADERS("COCOBREC", "")
   Call HEADERS("COCOBREC", PERIO$)
   '
   Call FILESORT("COCOBREC", "", 1, 4, "ASC")
   Call FINAL("*LICUPRO")
   '
99 Screen.MousePointer = 1
   '
End Sub
'
Sub ICALID()

50 ' INDICE DE CALIDAD DE PROVEEDORES
   '
   Screen.MousePointer = 11
   VECLOT$ = ""
  '\\\OT-08-0641-OD-25/08/08///

   DES = FETEXCOR1$(DESFECHA%)
   HAS = FETEXCOR1$(HASFECHA%)

   SQLX$ = "SELECT * FROM BOLRECEP"
   SQLX$ = SQLX$ + " WHERE FechRecep >= '" & DES & "'"
   SQLX$ = SQLX$ + " AND FechRecep <= '" & HAS & "'"
   SQLX$ = SQLX$ + " AND CantNom > 0 "
   '
   Dim BolRecepTemp As ADODB.Recordset
   Set BolRecepTemp = FLEXCONN.Execute(FILTSQL$(SQLX$))
   With BolRecepTemp
    On Error Resume Next
    .MoveFirst
    If Err < 1 Then
      On Error GoTo 0
      Do While Not (.EOF)
        NP% = XVALO(!NumProv)
        NRO& = XVALO(!NUMEBORE) 'CVS(Left$(X$, 4))
        CI% = XVALO(!COD_INTE) 'CVI(Mid$(X$, 83, 2))
        REMI$ = SAFETEXT(!NumRem_Prov) 'Mid$(X$, 119, 16)
        CANOM = XVALO(!CantNom) ' CANTIDAD NOMINAL
        CANAP = XVALO(!CanToApro) 'CVD(Mid$(X$, 163, 8))
        LOTE$ = SAFETEXT$(!IdenLote) 'Mid$(X$, 69, 12)
        '
        Z$ = REGBLAN$("COCAREC")
        Call REPLA(Z$, MKI$(NP%), 1, 2)
        Call REPLA(Z$, MKS$(NRO&), 3, 4)
        Call REPLA(Z$, MKI$(FF%), 7, 2)
        Call REPLA(Z$, REMI$, 9, 16)
        Call REPLA(Z$, LOTE$, 25, 12)
        Call REPLA(Z$, MKI$(CI%), 37, 2)
        Call REPLA(Z$, MKS$(CANOM), 39, 4)
        Call REPLA(Z$, MKI$(1), 55, 2)
        JJ& = JJ& + 1
        Call GRAREG("COCAREC", Z$, JJ&)
        VECLOT$ = VECLOT$ + AJUSTI$(LOTE$, 12)
        .MoveNext
      Loop
    End If
    On Error GoTo 0
   End With
   BolRecepTemp.Close
   Set BolRecepTemp = Nothing
   '
   Call ABRESTOCKS
   SQLX$ = "SELECT * FROM MOVISTO"
   SQLX$ = SQLX$ + " WHERE CodiMovi = 'BR' OR CODIMOVI = 'RD' "
   SQLX$ = SQLX$ + "OR ReferCor LIKE 'RECHAZO%'"
   SQLX$ = SQLX$ + "OR ReferCor LIKE 'FALTANTE%'"
   Dim MovistoTemp As ADODB.Recordset
   Set MovistoTemp = FLEXCONN.Execute(FILTSQL$(SQLX$))
   With MovistoTemp
    On Error Resume Next
    .MoveFirst
    If Err < 1 Then
      On Error GoTo 0
      Do While Not (.EOF)
        CANTI = XVALO(!CANTINGRE) - XVALO(!CANTSALID) 'CVD(Mid$(Y$, 88, 8)) - CVD(Mid$(Y$, 96, 8))
         If Abs(CANTI) >= 0.05 Then
           LOTE$ = SAFETEXT(!IdenLote)
            If TRIM$(LOTE$) <> "" Then
              If InStr(VECLOT$, TRIM$(LOTE$)) > 0 Then ' SE INCLUYE TRIM PORQUE SINO ALGUNAS BR NO LAS ENCONTRABA AUNQUE ESTABAN EN VECLOT
              For KK& = 1 To Len(VECLOT$) Step 12
               If TRIM$(Mid$(VECLOT$, KK&, 12)) = TRIM$(LOTE$) Then
                 REG& = (KK& + 11) / 12
                 If REG& > 0 Then
                   If REG& <= JJ& Then
                     Z$ = REGLEIDO$("COCAREC", REG&)
                     REFE$ = TRIM$(UCase$(!ReferCor))
                     If Left$(REFE$, 7) = "RECHAZO" Then
                          Call REPLA(Z$, MKS$(Abs(CANTI)), 47, 4)
                          Call REPLA(Z$, MKI$(1), 57, 2)
                       ElseIf Left$(REFE$, 8) = "FALTANTE" Then
                          Call REPLA(Z$, MKS$(Abs(CANTI)), 51, 4)
                       Else
                          Call REPLA(Z$, MKS$(Abs(CANTI)), 43, 4)
                     End If
                     Call GRAREG("COCAREC", Z$, REG&)
                     GoTo 69
                   End If
                 End If
               End If
              Next KK&
              End If
           End If
         End If
69    .MoveNext
      Loop
    End If
    On Error GoTo 0
   End With
   MovistoTemp.Close
   Set MovistoTemp = Nothing
   '
   Call SETULTREG("COCAREC", JJ&)
   Call CIERRARCH("COCAREC")
   Call CIERRARCH("*.*")
   '
   If Option3.Value = True Then GoTo 99 ' Calificacion de Calidad
   If Option4.Value = True Then GoTo 99 ' Calificacion de Calidad
   '
   Call HEADERS("COCAREC", "")
   Call HEADERS("COCAREC", PERIO$)
   '
   Call FILESORT("COCAREC", "", 1, 4, "ASC")
   Call FINAL("*LICAREC")
   '
99 Screen.MousePointer = 1
   '
End Sub
'
Sub ANACALI()
   '
   Call ANACALP
   '
   Call HEADERS("ANACALP", "")
   Call HEADERS("ANACALP", PERIO$)
   '
   Call FILESORT("ANACALP", "", 1, 3, "ASC")
   Call FINAL("*ANACALP")
   '
End Sub
'
Sub PROCALI()
   '
   Call ANACALP
   '
   For U& = 1 To ULTREG&("RESCALP")
     XX$ = REGLEIDO$("RESCALP", U&)
     CACAS% = CVI(Mid$(XX$, 101, 2))
     If CACAS% < 1 Then CACAS% = 1
     For KKI% = 1 To 8
       Call REPLA(XX$, MKS$(CVS(Mid$(XX$, 24 + 4 * KKI% - 3, 4)) / CACAS%), 24 + 4 * KKI% - 3, 4)
     Next KKI%
     Call REPLA(XX$, MKS$(CVS(Mid$(XX$, 97, 4)) / CACAS%), 97, 4)
     Call GRAREG("RESCALP", XX$, U&)
   Next U&
   Call CIERRARCH("RESCALP")
   '
   Call HEADERS("RESCALP", "")
   Call HEADERS("RESCALP", PERIO$)
   '
   Call FILESORT("RESCALP", "", 1, 1, "ASC")
   Call FINAL("*RESCALP")
   '
End Sub
'
Sub ANACALP()
   '
   Dim PUNPAR(8)
   ESCALI = Val(Label13)
   '
   Screen.MousePointer = 11
   REBLA$ = REGBLAN$("ANACALP")
   J1& = 0: J2& = 0: PUNPRO$ = ""
   FIN& = ULTREG&("CALIREC")
   For U& = 1 To FIN&
     Call TRACE(20, U&, FIN&)
     XX$ = REGLEIDO$("CALIREC", U&)
     FF% = CVI(Mid$(XX$, 5, 2))
     If FF% >= DESFECHA% Then
       If FF% <= HASFECHA% Then
         NP% = CVI(Mid$(XX$, 7, 2))
             NBR& = CVS(Left$(XX$, 4))
             REMI$ = Mid$(XX$, 9, 16)
             INSPE$ = Mid$(XX$, 27, 22)
             PUNTOT = CVS(Mid$(XX$, 49, 4))
             REPUN$ = Mid$(XX$, 65, 192)
             For KKI% = 1 To 8
               PUNPAR(KKI%) = ESCALI * CVS(Mid$(REPUN$, 4 * KKI% - 3, 4))
             Next KKI%
             '
             YY$ = REBLA$
             Call REPLA(YY$, MKI$(NP%), 1, 2)
             Call REPLA(YY$, MKS$(NBR&), 3, 4)
             Call REPLA(YY$, MKI$(FF%), 7, 2)
             Call REPLA(YY$, REMI$, 9, 16)
             For KKI% = 1 To 8
               Call REPLA(YY$, MKS$(PUNPAR(KKI%)), 24 + 4 * KKI% - 3, 4)
             Next KKI%
             Call REPLA(YY$, MKS$(PUNTOT), 97, 4)
             Call REPLA(YY$, MKI$(1), 101, 2)
             Call REPLA(YY$, INSPE$, 103, 22)
             Call REPLA(YY$, MKS$(0), 125, 4)
             J1& = J1& + 1
             Call GRAREG("ANACALP", YY$, J1&)
             '
             For J3& = 1 To J2&
               ZZ$ = REGLEIDO$("RESCALP", J3&)
               If CVI(Left$(ZZ$, 2)) = NP% Then
                 Call REPLA(ZZ$, MKS$(0), 3, 4)
                 Call REPLA(ZZ$, MKI$(0), 7, 2)
                 Call REPLA(ZZ$, "", 9, 16)
                 For KKI% = 1 To 8
                   Call REPLA(ZZ$, MKS$(CVS(Mid$(ZZ$, 24 + 4 * KKI% - 3, 4)) + PUNPAR(KKI%)), 24 + 4 * KKI% - 3, 4)
                 Next KKI%
                 Call REPLA(ZZ$, MKS$(CVS(Mid$(ZZ$, 97, 4)) + PUNTOT), 97, 4)
                 Call REPLA(ZZ$, MKI$(CVI(Mid$(ZZ$, 101, 2)) + 1), 101, 2)
                 Call REPLA(ZZ$, "", 103, 22)
                 Call GRAREG("RESCALP", ZZ$, J3&)
                 GoTo 49
               End If
             Next J3&
             J2& = J2& + 1
             Call GRAREG("RESCALP", YY$, J2&)
             '
       End If
     End If
49 Next U&
   '
   Call SETULTREG("ANACALP", J1&)
   Call SETULTREG("RESCALP", J2&)
   Screen.MousePointer = 1
   '
End Sub

Sub CALICALIDAD()
   '
   Call ICALID
   'Comienza calculo de Calidad
   '
   Dim QtTotal()
   Dim QtRech()
   ReDim QtTotal(1)
   ReDim QtRech(1)
   '
   Ultimo& = 0
   ANT% = 0
   For U& = 1 To ULTREG("COCAREC")
     x$ = REGLEIDO("COCAREC", U&)
     PROV% = CVI(Mid(x$, 1, 2))
     CANTOT# = CVS(Mid(x$, 39, 4))
     CANTRECH# = CVS(Mid(x$, 47, 4))
     '
     If PROV% > ANT% Then
        ReDim Preserve QtTotal(PROV%)
        ReDim Preserve QtRech(PROV%)
        ANT% = PROV%
     End If
    
     QtTotal(PROV%) = QtTotal(PROV%) + CANTOT#
     QtRech(PROV%) = QtRech(PROV%) + CANTRECH#

     '
     If PROV% > Ultimo& Then Ultimo& = PROV%
   Next U&
    
    'AGREGAR AL ARCHIVO LOS RESULTADOS
    REGIGRA& = 0
    Call SETULTREG("CALICALI", REGIGRA&)
    Call CIERRARCH("CALICALI")
    
    For U& = 1 To Ultimo&
      TOTAL# = QtTotal(U&)
      If TOTAL# > 0 Then
         APRO# = (QtTotal(U&) - QtRech(U&))
         PORC# = (APRO# / QtTotal(U&)) * 100
         '
         REGIGRA& = REGIGRA& + 1
         Call REPLA(x$, MKI(U&), 1, 2)
         Call REPLA(x$, MKD(XVALO(QtTotal(U&))), 3, 8)
         Call REPLA(x$, MKD(APRO#), 11, 8)
         Call REPLA(x$, MKD(XVALO(QtRech(U&))), 19, 8)
         Call REPLA(x$, MKS(PORC#), 27, 4)
         Call GRAREG("CALICALI", x$, REGIGRA&)
         '
      End If
    Next U&
   
   Call SETULTREG("CALICALI", REGIGRA&)
   Call CIERRARCH("CALICALI")
   '
   If Option4.Value = True Then GoTo 99 'Calificacion integral
   '
   Call FILESORT("CALICALI", "", 1, 2, "ASC")
   Call FINAL("*LICALCAL")
   '
99 Screen.MousePointer = 1
End Sub
    

