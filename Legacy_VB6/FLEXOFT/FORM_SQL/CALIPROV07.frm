VERSION 5.00
Begin VB.Form CALIPROV07 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0FFE0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Sistema de Calificación de Proveedores"
   ClientHeight    =   5235
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   6900
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
   ScaleHeight     =   5235
   ScaleWidth      =   6900
   StartUpPosition =   1  'CenterOwner
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
      Height          =   3750
      Left            =   140
      TabIndex        =   1
      Top             =   1365
      Width           =   6630
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
         Left            =   4305
         TabIndex        =   28
         Top             =   2520
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
            Index           =   4
            Left            =   1680
            MaxLength       =   2
            TabIndex        =   34
            Text            =   "15"
            ToolTipText     =   "% Por Atraso / Cumplimiento de Entregas"
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
            TabIndex        =   33
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
            Index           =   2
            Left            =   630
            MaxLength       =   2
            TabIndex        =   32
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
            Index           =   0
            Left            =   1575
            TabIndex        =   30
            Text            =   "20"
            Top             =   630
            Width           =   435
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
            TabIndex        =   29
            Top             =   660
            Width           =   1380
         End
      End
      Begin VB.OptionButton Option5 
         BackColor       =   &H00C0FFE0&
         Caption         =   "Calificación Integral"
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
         Left            =   4305
         TabIndex        =   27
         Top             =   2100
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
         Left            =   4305
         Picture         =   "CALIPROV07.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   1470
         Width           =   2115
      End
      Begin VB.OptionButton Option4 
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
         Left            =   4305
         TabIndex        =   24
         Top             =   1155
         Width           =   2115
      End
      Begin VB.OptionButton Option3 
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
         Left            =   4305
         TabIndex        =   23
         Top             =   900
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
         Left            =   4305
         TabIndex        =   22
         Top             =   660
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
         Left            =   4305
         TabIndex        =   21
         Top             =   405
         Value           =   -1  'True
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
         Height          =   1065
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
      Left            =   140
      TabIndex        =   0
      Top             =   210
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
            Picture         =   "CALIPROV07.frx":066A
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
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00004000&
      Height          =   950
      Left            =   3990
      ScaleHeight     =   885
      ScaleWidth      =   3345
      TabIndex        =   35
      Top             =   320
      Width           =   3405
      Begin VB.PictureBox MARCOBARRA 
         BackColor       =   &H000000FF&
         Height          =   135
         Left            =   150
         ScaleHeight     =   75
         ScaleWidth      =   540
         TabIndex        =   39
         Top             =   200
         Width           =   600
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H00FF0000&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   0
            TabIndex        =   40
            Top             =   0
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
         Height          =   900
         Left            =   2235
         Picture         =   "CALIPROV07.frx":0974
         Style           =   1  'Graphical
         TabIndex        =   38
         ToolTipText     =   "Abandona la Aplicación"
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
         Left            =   1590
         Picture         =   "CALIPROV07.frx":0ABE
         Style           =   1  'Graphical
         TabIndex        =   37
         ToolTipText     =   "Ayuda FLEXOFT en Línea"
         Top             =   0
         Width           =   645
      End
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
         Left            =   945
         Picture         =   "CALIPROV07.frx":0DC8
         Style           =   1  'Graphical
         TabIndex        =   36
         ToolTipText     =   "Configuración de Funcionamiento"
         Top             =   0
         Width           =   645
      End
      Begin VB.Image Image1 
         Height          =   390
         Left            =   -280
         Picture         =   "CALIPROV07.frx":120A
         Top             =   500
         Width           =   1515
      End
   End
   Begin VB.Label Label13 
      Caption         =   "5"
      Height          =   330
      Left            =   0
      TabIndex        =   26
      Top             =   0
      Visible         =   0   'False
      Width           =   540
   End
End
Attribute VB_Name = "CALIPROV07"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim DESFECHA%, HASFECHA%
Dim DESPROVE%, HASPROVE%
Dim PERIO$
Dim REOCO&(8192), PUBOL$(8192), CAPEN(8192)
'
Private Sub Command1_Click()
    Call CIERRARCH("*.*")
    Call FINAL("")
End Sub
'

Private Sub Command2_Click()
   Command2.Enabled = False
   If DERACCE%("ACTICAL", "Actualización Indices de Calificación") >= 2 Then
     STPCALPR07.Show 1
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
   DESPROVE% = Val(Text10)
   If DESPROVE% < 1 Or ECOARCH$("ACREDOR", MKI$(DESPROVE%)) = "" Then
     Call MENSERR(24, "Proveedor Inexistente.")
     GoTo 99
   End If
   '
   HASPROVE% = Val(Text8)
   If HASPROVE% < 1 Or ECOARCH$("ACREDOR", MKI$(HASPROVE%)) = "" Then
     Call MENSERR(24, "Proveedor Inexistente.")
     GoTo 99
   End If
   '
   If HASPROVE% < DESPROVE% Then
     Call MENSERR(24, "Rango de Proveedores Incorrecto.")
     GoTo 99
   End If
   '
   PERIO$ = "Periodo: " + Label4 + " - " + Label9
   For KKI% = 0 To 8192
     REOCO&(KKI%) = 0
     PUBOL$(KKI%) = ""
     CAPEN(KKI%) = 0
   Next KKI%
   '
   If Option1.Value = True Then
       Call ICUMPLI
     ElseIf Option2.Value = True Then
       Call ICALID
     ElseIf Option3.Value = True Then
       Call ANACALI
     ElseIf Option4.Value = True Then
       Call PROCALI
     ElseIf Option5.Value = True Then
       Call CALINTEG
   End If
   '
99 Command6.Enabled = True
End Sub

Sub CALINTEG()
   '
   Call ICALID
   Call ICUMPLI
   Call ANACALP
   '
   Screen.MousePointer = 11
   '
   PONDERA$ = Space$(15)
   For KKI% = 0 To 4
     Mid$(PONDERA$, 3 * KKI% + 1, 2) = Text1(KKI%)
   Next KKI%
   Call GRACONTROL("CALIPROV", "PONDERA", PONDERA$)
   '
   Dim TOPIRE(32767), TOPIFA(32767)
   Dim TOPICOM(32767), TODIATRA#(32767)
   Dim TOPICH(32767)
   Dim TOPUNCERT(32767), TOPUNSTAT(32767), MAXIPOS(32767)
   Dim ACTIPRO%(32767)
   '
   FIN& = ULTREG&("COCAREC")
   For U& = 1 To FIN&
     Call TRACE(20, U&, FIN&)
     XX$ = REGLEIDO$("COCAREC", U&)
     NPRO% = CVI(Left$(XX$, 2))
     CAPIRE = CVS(Mid$(XX$, 39, 4))
     CAPIFA = CVS(Mid$(XX$, 51, 4))
     If NPRO% >= 0 Then
       ACTIPRO%(NPRO%) = 1
       TOPIRE(NPRO%) = TOPIRE(NPRO%) + CAPIRE
       TOPIFA(NPRO%) = TOPIFA(NPRO%) + CAPIFA
     End If
   Next U&
   '
   FIN& = ULTREG&("ANACALP")
   For U& = 1 To FIN&
     Call TRACE(20, U&, FIN&)
     XX$ = REGLEIDO$("ANACALP", U&)
     NPRO% = CVI(Left$(XX$, 2))
     PUNCERT = CVS(Mid$(XX$, 29, 4))
     PUNSTAT = CVS(Mid$(XX$, 33, 4))
     If NPRO% >= 0 Then
       ACTIPRO%(NPRO%) = 1
       TOPUNCERT(NPRO%) = TOPUNCERT(NPRO%) + PUNCERT
       TOPUNSTAT(NPRO%) = TOPUNSTAT(NPRO%) + PUNSTAT
       MAXIPOS(NPRO%) = MAXIPOS(NPRO%) + 5
     End If
   Next U&
   '
   FIN& = ULTREG&("COCOBREC")
   For U& = 1 To FIN&
     Call TRACE(20, U&, FIN&)
     XX$ = REGLEIDO$("COCOBREC", U&)
     NPRO% = CVI(Left$(XX$, 2))
     CAPIOR = CVS(Mid$(XX$, 9, 4))
     DIATRA = CVI(Mid$(XX$, 37, 2))
     If NPRO% >= 0 Then
       If DIATRA > 0 Then
         ACTIPRO%(NPRO%) = 1
         TOPICOM(NPRO%) = TOPICOM(NPRO%) + CAPIOR
         TODIATRA#(NPRO%) = TODIATRA#(NPRO%) + CDbl(CAPIOR * DIATRA)
       End If
     End If
   Next U&
   '
   DES% = FECHANUM(Label4)
   Has% = FECHANUM(Label9)
   FIN& = ULTREG&("RECHLIN")
   For J& = 1 To ULTREG&("RECHLIN")
     Call TRACE(20, J&, FIN&)
     X$ = REGLEIDO$("RECHLIN", J&)
     FEREC3% = CVI(Mid$(X$, 7, 2))
     If FEREC3% >= DES% Then
       If FEREC3% <= Has% Then
         NPRO% = CVI(Mid$(X$, 9, 2))
         CANTI3 = CVS(Mid$(X$, 11, 4))
         If NPRO% >= 0 Then
           If CANTI3 > 0.5 Then
             ACTIPRO%(NPRO%) = 1
             TOPICH(NPRO%) = TOPICH(NPRO%) + CANTI3
           End If
         End If
       End If
     End If
   Next J&
   '
   REBLA$ = REGBLAN$("CALINTEG")
   JJ& = 0
   For U& = 1 To 32767
     If ACTIPRO%(U&) > 0 Then
       ZZ$ = REBLA$
       NPRX% = U&
       Call REPLA(ZZ$, MKI$(NPRX%), 1, 2)
       Call REPLA(ZZ$, MKS$(TOPIRE(NPRX%)), 3, 4)
       Call REPLA(ZZ$, MKS$(TOPIFA(NPRX%)), 7, 4)
       PUNTA0 = 100    ' FALTANTES
       If TOPIRE(NPRX%) > 0.5 Then
         Call REPLA(ZZ$, MKS$(100 * TOPIFA(NPRX%) / TOPIRE(NPRX%)), 11, 4)
         PUNTA0 = 100 - 100 * TOPIFA(NPRX%) / TOPIRE(NPRX%)
       End If
       '
       Call REPLA(ZZ$, MKS$(TOPICH(NPRX%)), 15, 4)
       PUNTA1 = 100    ' RECHAZOS
       If TOPIRE(NPRX%) > 0.5 Then
         Call REPLA(ZZ$, MKS$(100 * TOPICH(NPRX%) / TOPIRE(NPRX%)), 19, 4)
         PONDERECH = Val(Text1(0)): If PONDERECH < 1 Then PONDERECH = 100
         PUNPERDI = (10000 / PONDERECH) * TOPICH(NPRX%) / TOPIRE(NPRX%)
'         (10000 / VALOR DE PANTALLA PORCENTAJE LIMITE) * CANTIDAD RECHAZADA /TOTAL RECIBIDA
         PUNTA1 = 100 - PUNPERDI: If PUNTA1 < 0 Then PUNTA1 = 0
       End If
       '
       ATRAPROM = 0
       PUNTA4 = 100
       If TOPICOM(NPRX%) > 0.5 Then
         ATRAPROM = TODIATRA#(NPRX%) / TOPICOM(NPRX%)
       End If
       Call REPLA(ZZ$, MKS$(ATRAPROM), 23, 4)
       If ATRAPROM >= 7 Then PUNTA4 = 67
       If ATRAPROM >= 14 Then PUNTA4 = 33
       If ATRAPROM > 21 Then PUNTA4 = 0
       '
       PUNTA2 = 100    ' CERTIFICADOS
       PUNTA3 = 100    ' ESTADO
       If MAXIPOS(NPRX%) > 0 Then
         Call REPLA(ZZ$, MKS$(100 * TOPUNCERT(NPRX%) / MAXIPOS(NPRX%)), 27, 4)
         PUNTA2 = 100 * TOPUNCERT(NPRX%) / MAXIPOS(NPRX%)
         Call REPLA(ZZ$, MKS$(100 * TOPUNSTAT(NPRX%) / MAXIPOS(NPRX%)), 31, 4)
         PUNTA3 = 100 * TOPUNSTAT(NPRX%) / MAXIPOS(NPRX%)
       End If
       '
'AA1 = RASOCLI$((-1) * NPRX%)
       'PUNTOTA = 0.7 * PUNTA1 + 0.1 * PUNTA2 + 0.05 * PUNTA3 + 0.15 * PUNTA4
       PUNTOTA = (Val(Text1(1)) * PUNTA1 + Val(Text1(2)) * PUNTA2 + Val(Text1(3)) * PUNTA3 + Val(Text1(4)) * PUNTA4) / 100
       Call REPLA(ZZ$, MKS$(PUNTOTA / 10), 35, 4)
       '
       JJ& = JJ& + 1
       Call GRAREG("CALINTEG", ZZ$, JJ&)
       '
     End If
   Next U&
   Call SETULTREG("CALINTEG", JJ&)
   Call HEADERS("CALINTEG", "")
   Call HEADERS("CALINTEG", PERIO$)
   '
   Call FINAL("*CALINTEG")
   '
99 Screen.MousePointer = 1
   '
End Sub
'
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
    Has% = HOY(HOS$)
    FEX = Has%
    Label9 = FECHATEX$(FEX)
    Label4 = "01" + Mid$(Label9, 3)
    '
    If Val(Control$("CALIPROV", "ESCACALI")) = 10 Then
      Label13 = "10"
    End If
    '
    PONDERA$ = TRIM$(Control$("CALIPROV", "PONDERA"))
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
   Call SELECHO("SELFECHA")
   If VALACT1$("SELFECHA") <> "" Then
     Label9 = VALACT1$("SELFECHA")
   End If
End Sub

Private Sub Label9_DblClick()
   Call SELECHO("SELFECHA")
   If VALACT1$("SELFECHA") <> "" Then
     Label9 = VALACT1$("SELFECHA")
   End If
End Sub

Private Sub Text10_Change()
   Label7 = ""
   NCMIN% = Val(Text10.TEXT)
   If Val(NCMIN%) > 0 Then
     Label7 = rasocli$((-1) * NCMIN%)
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
     Label8 = rasocli$((-1) * NCMAX%)
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
10 J& = 0: PUCON$ = "": U& = 0
   'FIN& = ULTREG&("OCOMDETA")
   '\\\OT-08-0641-OD-25/08/08///
   DESTINO% = 0
   If ULTREG&("ECUECON") > 0 Then
        Call SELECHO("ECUECON")
        CCO$ = TRIM$(VALACT1$("ECUECON"))
        If CCO$ <> "" Then DESTINO% = CUEINT%(CCO$)
   End If
   '
   Screen.MousePointer = 11
   '
   DES = FECHATEX(DESFECHA%)
   Has = FECHATEX(HASFECHA%)
   SQLX$ = "SELECT * FROM OCOMDETA WITH(NOLOCK)"
   SQLX$ = SQLX$ + " WHERE Stat_Ocom <> 9 "
   SQLX$ = SQLX$ + " AND Femi_Ocom >= '" & DES & "'"
   SQLX$ = SQLX$ + " AND Femi_Ocom <= '" & Has & "'"
   SQLX$ = SQLX$ + " AND Npro_Ocom >= " & DESPROVE% & ""
   SQLX$ = SQLX$ + " AND Npro_Ocom <= " & HASPROVE% & ""
   If DESTINO% > 0 Then SQLX$ = SQLX$ + " AND icontab = " & DESTINO% & ""
   Dim OrdenCompraTemp As ADODB.Recordset
   Set OrdenCompraTemp = FLEXCONN.Execute(FILTSQL$(SQLX$))
   With OrdenCompraTemp
      Do While Not (.EOF)
        U& = U& + 1
        J& = J& + 1
        REOCO&(J&) = U&
        CAPEN(J&) = XVALO(!cant_ocom) 'CVS(Mid$(X$, 55, 4))
        CLACO$ = MKI$(XVALO(!Npro_Ocom)) + MKS$(XVALO(!NUME_OCOM)) + MKI$(XVALO(!COD_INTE))   'Mid$(X$, 7, 2) + Left$(X$, 4) + Mid$(X$, 9, 2)
        PUCON$ = PUCON$ + CLACO$
        .MoveNext
      Loop
   End With
   OrdenCompraTemp.Close
   Set OrdenCompraTemp = Nothing
   CAOCO& = J&
   
'   For U& = 1 To FIN&
'      Call TRACE(20, U&, FIN&)
'      X$ = REGLEIDO$("OCOMDETA", U&)
'      FF% = CVI(Mid$(X$, 5, 2))        ' fecha de emision
'      NP% = CVI(Mid$(X$, 7, 2))        ' proveedor
'      STAT% = Asc(Mid$(X$, 124, 1))
'      '
'      If STAT% <> 9 Then               ' excluye anuladas
'        If NP% >= DESPROVE% Then
'          If NP% <= HASPROVE% Then           ' excluye pedidos suministro
'            If FF% >= DESFECHA% Then
'              If FF% <= HASFECHA% Then
'                J& = J& + 1
'                If J& > 8192 Then
'                  Call MENSERR(24, "Demasiados Movimientos\en el Rango de Fechas Elegido.")
'                  GoTo 10
'                End If
'                '
'                REOCO&(J&) = U&
'                CAPEN(J&) = CVS(Mid$(X$, 55, 4))
'                CLACO$ = Mid$(X$, 7, 2) + Left$(X$, 4) + Mid$(X$, 9, 2)
'                PUCON$ = PUCON$ + CLACO$
'              End If
'            End If
'          End If
'        End If
'      End If
'   Next U&
'   '
   SQLX$ = "SELECT * FROM BOLRECEP WITH(NOLOCK)"
   Dim BolRecepTmp As ADODB.Recordset
   Set BolRecepTmp = FLEXCONN.Execute(FILTSQL$(SQLX$))
   '
   With BolRecepTmp
      Do While Not (.EOF)
        U& = U& + 1
        J& = J& + 1
        CAREC = XVALO(!CANTNOM) '(Mid$(X$, 89, 8))
        CLACO$ = MKI$(XVALO(!NUMPROV)) + MKS$(XVALO(!NUMOC)) + MKI$(XVALO(!COD_INTE))
        For k& = 1 To CAOCO&
            If Mid$(PUCON$, 8 * k& - 7, 8) = CLACO$ Then
              UKA& = k&
              If CAPEN(k&) > 0 Then
                PUBOL$(k&) = PUBOL$(k&) + MKS$(U&)
                CAPEN(k&) = CAPEN(k&) - CAREC
                GoTo 19
              End If
            End If
        Next k&
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


'   FIN& = ULTREG&("BOLRECEP")
'   For U& = 1 To FIN&
'      Call TRACE(20, U&, FIN&)
'      X$ = REGLEIDO$("BOLRECEP", U&)
'      CAREC = CVD '(Mid$(X$, 89, 8))
'      CLACO$ = Mid$(X$, 105, 2) + Mid$(X$, 151, 4) + Mid$(X$, 83, 2)
'      UKA& = 0
'      For K& = 1 To CAOCO&
'        If Mid$(PUCON$, 8 * K& - 7, 8) = CLACO$ Then
'          UKA& = K&
'          If CAPEN(K&) > 0 Then
'            PUBOL$(K&) = PUBOL$(K&) + MKS$(U&)
'            CAPEN(K&) = CAPEN(K&) - CAREC
'            GoTo 19
'          End If
'        End If
'      Next K&
      '
      If UKA& > 0 Then
        PUBOL$(UKA&) = PUBOL$(UKA&) + MKS$(U&)
        CAPEN(UKA&) = CAPEN(UKA&) - CAREC
      End If
      '
'19 Next U&
   '
   HOII = HOY(HOS$)
   SQLX$ = "SELECT * FROM OCOMDETA WITH(NOLOCK)"
   SQLX$ = SQLX$ + " WHERE Stat_Ocom <> 9 "
   SQLX$ = SQLX$ + " AND Femi_Ocom >= '" & DES & "'"
   SQLX$ = SQLX$ + " AND Femi_Ocom <= '" & Has & "'"
   SQLX$ = SQLX$ + " AND Npro_Ocom >= " & DESPROVE% & ""
   SQLX$ = SQLX$ + " AND Npro_Ocom <= " & HASPROVE% & ""
   If DESTINO% > 0 Then SQLX$ = SQLX$ + " AND icontab = " & DESTINO% & ""
   '
   Z0$ = REGBLAN$("COCOBRE1")
   '
   Dim OrdenCompTemp As ADODB.Recordset
   Set OrdenCompTemp = FLEXCONN.Execute(FILTSQL$(SQLX$))
   With OrdenCompTemp
      Do While Not (.EOF)
          Call REPLA(Z0$, MKI$(XVALO(!Npro_Ocom)), 1, 2)
          Call REPLA(Z0$, MKS$(XVALO(!NUME_OCOM)), 3, 4)
          Call REPLA(Z0$, MKI$(XVALO(!COD_INTE)), 7, 2)
          Call REPLA(Z0$, MKS$(XVALO(!cant_ocom)), 9, 4)
          FEXI% = CVINT(!Fentre_Sol)
          Call REPLA(Z0$, MKI$(FEXI%), 13, 2)
          FENTREPRO = CVINT(!Fentre_Sol)
            CCON% = XVALO(!icontab)
          Call REPLA(Z0$, CODCUE(CCON%), 37, 12)
          '
          For k& = 1 To CAOCO&
            CLACO$ = MKI$(XVALO(!Npro_Ocom)) + MKS$(XVALO(!NUME_OCOM)) + MKI$(XVALO(!COD_INTE))
            If Mid$(PUCON$, 8 * k& - 7, 8) = CLACO$ Then Exit For
          Next k&
          '
          If Len(PUBOL$(k&)) < 4 Then
               DIATRA% = 0
               If FENTREPRO < HOII Then
                 DIATRA% = DIENTRE%(FENTREPRO, HOII)
               End If
               Call REPLA(Z0$, MKI$(DIATRA%), 49, 2)
               NCOCO& = NCOCO& + 1
               Call GRAREG("COCOBRE1", Z0$, NCOCO&)
             Else
               'For II& = 1 To Len(PUBOL$(K&)) Step 4
                 Z$ = Z0$
                 'REBREC& = CVS(Mid$(PUBOL$(K&), II&, 4))
                 'Y$ = REGLEIDO$("BOLRECEP", REBREC&)
                 NumProveedor% = XVALO(!Npro_Ocom)
                 NumOrdCompra& = XVALO(!NUME_OCOM)
                 CodMaterial% = XVALO(!COD_INTE)
                 SQLX$ = "SELECT * FROM BOLRECEP WITH(NOLOCK)"
                 SQLX$ = SQLX$ + " WHERE NumProv = " & NumProveedor% & " "
                 SQLX$ = SQLX$ + " AND NumOC = " & NumOrdCompra& & " "
                 SQLX$ = SQLX$ + " AND Cod_Inte = " & CodMaterial% & ""
                 Dim BolRecTemp As ADODB.Recordset
                 Set BolRecTemp = FLEXCONN.Execute(FILTSQL$(SQLX$))
                 With BolRecTemp
                     Do While Not (.EOF)
                       Call REPLA(Z$, MKS$(XVALO(!CANTNOM)), 15, 4)
                       Call REPLA(Z$, MKI$(CVINT(!FechRecep)), 19, 2)
                       Call REPLA(Z$, SAFETEXT$(!NumRem_Prov), 21, 16)
                       FERECI = CVINT(!FechRecep)
                       DIATRA% = 0
                       If FENTREPRO < FERECI Then
                         DIATRA% = DIENTRE%(FENTREPRO, FERECI)
                       End If
                       Call REPLA(Z$, MKI$(DIATRA%), 49, 2)
                       NCOCO& = NCOCO& + 1
                       Call GRAREG("COCOBRE1", Z$, NCOCO&)
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
      
'      'X$ = REGLEIDO$("OCOMDETA", REOCO&(K&))
'      Z0$ = REGBLAN$("COCOBREC")
'
'      Call REPLA(Z0$, Mid$(X$, 7, 2), 1, 2)
'      Call REPLA(Z0$, Left$(X$, 4), 3, 4)
'      Call REPLA(Z0$, Mid$(X$, 9, 2), 7, 2)
'      Call REPLA(Z0$, Mid$(X$, 55, 4), 9, 4)
'      Call REPLA(Z0$, Mid$(X$, 95, 2), 13, 2)
'      FENTREPRO = CVI(Mid$(X$, 95, 2))
'      '
'      If Len(PUBOL$(K&)) < 4 Then
'           DIATRA% = 0
'           If FENTREPRO < HOII Then
'             DIATRA% = DIENTRE%(FENTREPRO, HOII)
'           End If
'           Call REPLA(Z0$, MKI$(DIATRA%), 37, 2)
'           NCOCO& = NCOCO& + 1
'           Call GRAREG("COCOBREC", Z0$, NCOCO&)
'         Else
'           For II& = 1 To Len(PUBOL$(K&)) Step 4
'             Z$ = Z0$
'             REBREC& = CVS(Mid$(PUBOL$(K&), II&, 4))
'             Y$ = REGLEIDO$("BOLRECEP", REBREC&)
'             Call REPLA(Z$, MKS$(CVD(Mid$(Y$, 89, 8))), 15, 4)
'             Call REPLA(Z$, Mid$(Y$, 81, 2), 19, 2)
'             Call REPLA(Z$, Mid$(Y$, 119, 16), 21, 16)
'             FERECI = CVI(Mid$(Y$, 81, 2))
'             DIATRA% = 0
'             If FENTREPRO < FERECI Then
'               DIATRA% = DIENTRE%(FENTREPRO, FERECI)
'             End If
'             Call REPLA(Z$, MKI$(DIATRA%), 37, 2)
'             NCOCO& = NCOCO& + 1
'             Call GRAREG("COCOBREC", Z$, NCOCO&)
'           Next II&
'      End If
      '
   'Next K&
   '\\\OT-08-0641-OD-FIN//
   '
   Call SETULTREG("COCOBRE1", NCOCO&)
   Call CIERRARCH("COCOBRE1")
   '
   If Option5.Value = True Then GoTo 99 ' CALIFICACION INTEGRAL
   '
   Call HEADERS("COCOBRE1", "")
   Call HEADERS("COCOBRE1", PERIO$)
   If DESTINO% > 0 Then Call HEADERS("COCOBRE1", "Destino: " & DENOCUE$(DESTINO%))
   '
   Call FILESORT("COCOBRE1", "", 1, 4, "ASC")
   Call FINAL("*LICUPR1")
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
   Has = FETEXCOR1$(HASFECHA%)

   SQLX$ = "SELECT * FROM BOLRECEP"
   SQLX$ = SQLX$ + " WHERE FechRecep >= '" & DES & "'"
   SQLX$ = SQLX$ + " AND FechRecep <= '" & Has & "'"
   SQLX$ = SQLX$ + " AND NumProv >= " & DESPROVE% & ""
   SQLX$ = SQLX$ + " AND NumProv <= " & HASPROVE% & ""
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
        NP% = XVALO(!NUMPROV)
        NRO& = XVALO(!NUMEBORE) 'CVS(Left$(X$, 4))
        CI% = XVALO(!COD_INTE) 'CVI(Mid$(X$, 83, 2))
        REMI$ = SAFETEXT(!NumRem_Prov) 'Mid$(X$, 119, 16)
        CANOM = XVALO(!CANTNOM) ' CANTIDAD NOMINAL
        CANAP = XVALO(!CanToApro) 'CVD(Mid$(X$, 163, 8))
        LOTE$ = SAFETEXT$(!IDENLOTE) 'Mid$(X$, 69, 12)
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
           LOTE$ = SAFETEXT(!IDENLOTE)
            If TRIM$(LOTE$) <> "" Then
              If InStr(VECLOT$, TRIM$(LOTE$)) > 0 Then ' SE INCLUYE TRIM PORQUE SINO ALGUNAS BR NO LAS ENCONTRABA AUNQUE ESTABAN EN VECLOT
              For KK& = 1 To Len(VECLOT$) Step 12
               If TRIM$(Mid$(VECLOT$, KK&, 12)) = TRIM$(LOTE$) Then
                 REG& = (KK& + 11) / 12
                 If REG& > 0 Then
                   If REG& <= JJ& Then
                     Z$ = REGLEIDO$("COCAREC", REG&)
                     REFE$ = TRIM$(UCase$(SAFETEXT$(!ReferCor)))
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
   If Option5.Value = True Then GoTo 99 ' CALIFICACION INTEGRAL
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
         If NP% >= DESPROVE% Then
           If NP% <= HASPROVE% Then
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
       End If
     End If
49 Next U&
   '
   Call SETULTREG("ANACALP", J1&)
   Call SETULTREG("RESCALP", J2&)
   Screen.MousePointer = 1
   '
End Sub
