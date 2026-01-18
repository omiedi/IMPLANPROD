VERSION 5.00
Begin VB.Form RETDET09 
   BackColor       =   &H00C0E0FF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Detalle - Retenciones legales y Aportes"
   ClientHeight    =   5865
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7410
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5865
   ScaleWidth      =   7410
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command10 
      Caption         =   "."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1930
      TabIndex        =   42
      Top             =   105
      Width           =   175
   End
   Begin VB.TextBox Text9 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1320
      TabIndex        =   36
      Top             =   480
      Width           =   4455
   End
   Begin VB.TextBox Text8 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3795
      MaxLength       =   8
      TabIndex        =   34
      Top             =   120
      Width           =   1980
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0E0FF&
      Height          =   4860
      Left            =   140
      TabIndex        =   3
      Top             =   840
      Width           =   6180
      Begin VB.Frame Frame2 
         BackColor       =   &H00C0E0FF&
         BorderStyle     =   0  'None
         Caption         =   "Frame2"
         Height          =   705
         Left            =   100
         TabIndex        =   37
         Top             =   2760
         Visible         =   0   'False
         Width           =   5955
         Begin VB.TextBox Text11 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   1365
            TabIndex        =   39
            Top             =   210
            Width           =   1335
         End
         Begin VB.TextBox Text10 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   4515
            TabIndex        =   38
            Top             =   210
            Width           =   1335
         End
         Begin VB.Label Label11 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFF80&
            BackStyle       =   0  'Transparent
            Caption         =   "Importe Retención:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   420
            Left            =   -315
            TabIndex        =   41
            Top             =   105
            Width           =   1575
         End
         Begin VB.Label Label9 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFF80&
            BackStyle       =   0  'Transparent
            Caption         =   "Importe Aportes:"
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
            Left            =   3255
            TabIndex        =   40
            Top             =   105
            Width           =   1155
         End
      End
      Begin VB.TextBox Text7 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   1455
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   28
         Text            =   "RETDET09.frx":0000
         Top             =   3570
         Width           =   4500
      End
      Begin VB.TextBox Text6 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   4935
         TabIndex        =   26
         Top             =   2970
         Width           =   1020
      End
      Begin VB.TextBox Text5 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1470
         TabIndex        =   24
         Top             =   2970
         Width           =   975
      End
      Begin VB.TextBox Text4 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1455
         MaxLength       =   3
         TabIndex        =   21
         Top             =   2280
         Width           =   405
      End
      Begin VB.CommandButton Command6 
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
         Height          =   300
         Left            =   1860
         Picture         =   "RETDET09.frx":0002
         TabIndex        =   20
         Top             =   2280
         Width           =   175
      End
      Begin VB.TextBox Text3 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1455
         MaxLength       =   3
         TabIndex        =   17
         Top             =   1800
         Width           =   405
      End
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
         Height          =   300
         Left            =   1860
         Picture         =   "RETDET09.frx":030C
         TabIndex        =   16
         Top             =   1800
         Width           =   175
      End
      Begin VB.TextBox Text2 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1455
         MaxLength       =   3
         TabIndex        =   13
         Top             =   1320
         Width           =   405
      End
      Begin VB.CommandButton Command2 
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
         Height          =   300
         Left            =   1860
         Picture         =   "RETDET09.frx":0616
         TabIndex        =   12
         Top             =   1320
         Width           =   175
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1455
         MaxLength       =   3
         TabIndex        =   9
         Top             =   840
         Width           =   405
      End
      Begin VB.CommandButton Command1 
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
         Height          =   300
         Left            =   1860
         Picture         =   "RETDET09.frx":0920
         TabIndex        =   8
         Top             =   840
         Width           =   175
      End
      Begin VB.TextBox Text16 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1455
         MaxLength       =   3
         TabIndex        =   5
         Top             =   360
         Width           =   405
      End
      Begin VB.CommandButton Command19 
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
         Height          =   300
         Left            =   1860
         Picture         =   "RETDET09.frx":0C2A
         TabIndex        =   4
         Top             =   360
         Width           =   175
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Formula:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   -540
         TabIndex        =   29
         Top             =   3570
         Width           =   1905
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFF80&
         BackStyle       =   0  'Transparent
         Caption         =   "Recalculo:"
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
         Left            =   300
         TabIndex        =   27
         Top             =   2385
         Width           =   1110
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFF80&
         BackStyle       =   0  'Transparent
         Caption         =   "Quincena:"
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
         Left            =   300
         TabIndex        =   25
         Top             =   1905
         Width           =   1110
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFF80&
         BackStyle       =   0  'Transparent
         Caption         =   "% Aportes:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   3675
         TabIndex        =   23
         Top             =   3045
         Width           =   1110
      End
      Begin VB.Label Label7 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2100
         TabIndex        =   22
         Top             =   2280
         Width           =   3855
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFF80&
         BackStyle       =   0  'Transparent
         Caption         =   "Carácter:"
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
         Left            =   300
         TabIndex        =   19
         Top             =   960
         Width           =   1110
      End
      Begin VB.Label Label5 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2100
         TabIndex        =   18
         Top             =   1800
         Width           =   3855
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFF80&
         BackStyle       =   0  'Transparent
         Caption         =   "Modo:"
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
         Left            =   -420
         TabIndex        =   15
         Top             =   1440
         Width           =   1815
      End
      Begin VB.Label Label3 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2100
         TabIndex        =   14
         Top             =   1320
         Width           =   3855
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFF80&
         BackStyle       =   0  'Transparent
         Caption         =   "% Retencion:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   -240
         TabIndex        =   11
         Top             =   3045
         Width           =   1575
      End
      Begin VB.Label Label1 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2100
         TabIndex        =   10
         Top             =   840
         Width           =   3855
      End
      Begin VB.Label Label16 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFF80&
         BackStyle       =   0  'Transparent
         Caption         =   "Estado:"
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
         Left            =   300
         TabIndex        =   7
         Top             =   465
         Width           =   1110
      End
      Begin VB.Label Label23 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2100
         TabIndex        =   6
         Top             =   360
         Width           =   3855
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00004080&
      Height          =   6135
      Left            =   6510
      ScaleHeight     =   6075
      ScaleWidth      =   1335
      TabIndex        =   0
      Top             =   0
      Width           =   1395
      Begin VB.CommandButton Command21 
         Caption         =   "New"
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
         Left            =   90
         Picture         =   "RETDET09.frx":0F34
         Style           =   1  'Graphical
         TabIndex        =   43
         ToolTipText     =   "Alta de Nuevo Item de Retenciones legales y Aportes"
         Top             =   1785
         Width           =   650
      End
      Begin VB.CommandButton Command9 
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
         Left            =   90
         Picture         =   "RETDET09.frx":107E
         Style           =   1  'Graphical
         TabIndex        =   33
         ToolTipText     =   "Guardar "
         Top             =   2625
         Width           =   640
      End
      Begin VB.CommandButton Command3 
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
         Left            =   90
         Picture         =   "RETDET09.frx":1388
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Exit"
         Top             =   120
         Width           =   650
      End
      Begin VB.CommandButton Command4 
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
         Left            =   90
         Picture         =   "RETDET09.frx":14D2
         Style           =   1  'Graphical
         TabIndex        =   1
         ToolTipText     =   "Ayuda en Línea"
         Top             =   810
         Width           =   650
      End
      Begin VB.Image Image2 
         Height          =   390
         Left            =   -330
         Picture         =   "RETDET09.frx":17DC
         Top             =   5250
         Width           =   1515
      End
   End
   Begin VB.Label Label13 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Abreviatura:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1800
      TabIndex        =   35
      Top             =   225
      Width           =   1905
   End
   Begin VB.Label Label19 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFF80&
      BackStyle       =   0  'Transparent
      Caption         =   "Descripción:"
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
      Left            =   120
      TabIndex        =   32
      Top             =   585
      Width           =   1110
   End
   Begin VB.Label Label17 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFF80&
      BackStyle       =   0  'Transparent
      Caption         =   "Código:"
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
      Left            =   120
      TabIndex        =   31
      Top             =   240
      Width           =   1110
   End
   Begin VB.Label Label15 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1320
      TabIndex        =   30
      Top             =   120
      Width           =   615
   End
End
Attribute VB_Name = "RETDET09"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim MODIFIC%

Private Sub Command1_Click()
    '
    Call SELECHO("TACARAC")
    VDEVU$ = TRIM$(VALACT1$("TACARAC"))
    If VDEVU$ <> "" Then
       Text1.TEXT = VALACT1$("TACARAC")
    End If
    '
End Sub

Private Sub Command10_Click()
  '
  Call SELECHO("!INDIRET")
  COD$ = TRIM$(VALACT1$("!INDIRET"))
  If COD$ = "" Then Exit Sub
  '
  JJ& = Val(COD$)
  ZZ$ = REGLEIDO$("RETLEG", JJ&)
  DESCRI$ = TRIM$(Mid$(ZZ$, 1, 30))
  ABRE$ = TRIM$(Mid$(ZZ$, 121, 8))
  PROP$ = Mid$(ZZ$, 31, 8)
  FORMU$ = Mid$(ZZ$, 49, 72)
  JJ1 = CVS(Mid$(ZZ$, 39, 4)) '%RETENCION
  JJ2 = CVS(Mid$(ZZ$, 43, 4)) '%APORTES
  With RETDET05
    .Text8 = ABRE$ 'ABREVIATURA
    .Text7 = FORMU$ 'FORMULA
       EST% = Asc(Mid$(PROP$, 1, 1))
       ET$ = ESTADO(EST%) 'DEVUELVE EL VALOR EN $ DEL ESTADO
    .Text16 = ET$ 'ESTADO
    .Text1 = Asc(Mid$(PROP$, 3, 1)) 'CARACTER
    .Text2 = Asc(Mid$(PROP$, 4, 1)) 'MODO DE CALCULO
    .Text3 = Asc(Mid$(PROP$, 6, 1)) 'QUINCENA
    .Text4 = Asc(Mid$(PROP$, 7, 1)) 'RECALCULO
    If Val(.Text2) <> 3 Then 'SI MODO <> SUMAS FIJAS
      .Text5 = FORMATNUM(JJ1, "F5.2") '%RET
      .Text6 = FORMATNUM(JJ2, "F5.2") '%APORTES
    Else
      .Text11 = FORMATNUM(JJ1, "F9.2") 'IMP. RET
      .Text10 = FORMATNUM(JJ2, "F9.2") 'IMP. APORTES
    End If
    .Label15 = COD$
    .Text9 = DESCRI$
    MODIFIC% = 0
  End With
  '
End Sub
'
Sub INDIRET()
  '
  ZBLA$ = REGBLAN$("INDIRET")
  JJ& = 0
  For J& = 1 To ULTREG&("RETLEG")
    X$ = REGLEIDO$("RETLEG", J&)
    DESCRI$ = Mid$(X$, 1, 30)
    ZZ$ = Space(32)
    '
    JX% = J&
    ZZ$ = ZBLA$
    Call REPLA(ZZ$, MKI$(JX%), 1, 2)
    Call REPLA(ZZ$, DESCRI$, 3, 30)
    JJ& = JJ& + 1
    Call GRAREG("!INDIRET", ZZ$, JJ&)
  Next J&
  Call SETULTREG("!INDIRET", JJ&)
  '
End Sub
Private Sub Command19_Click()
    '
    Call SELECHO("TASTATUS")
    VDEVU$ = TRIM$(VALACT1$("TASTATUS"))
    If VDEVU$ <> "" Then
       Text16.TEXT = VALACT1$("TASTATUS")
    End If
    '
End Sub

Private Sub Command2_Click()
    '
    Call SELECHO("TAMOCRET")
    VDEVU$ = TRIM$(VALACT1$("TAMOCRET"))
    If VDEVU$ <> "" Then
       Text2.TEXT = VALACT1$("TAMOCRET")
    End If
    '
End Sub

Private Sub Command21_Click()
  '
  '
  Command21.Enabled = False
  'VDEF$ TIENE QUE DAR EL PROX. Nº DE REMUNERACIÓN
  VDEF$ = Space(32)
  PRONU% = Int(ULTREG&("RETLEG"))
  Call REPLA(VDEF$, MKI$(PRONU% + 1), 1, 2)
9 OBJ$ = OBJPLA$("NUEREM", VDEF$)
  If OBJ$ = "" Then GoTo 99
  VDEF$ = OBJ$
  '
  DESCRI$ = TRIM$(Mid$(OBJ$, 3, 30))
  If TRIM$(DESCRI$) = "" Then
    Call MENSERR(24, "Descripción no Válida.")
    GoTo 9
  End If
  '
  ZZ$ = Space(64)
  NUME = CVI(Mid$(OBJ$, 1, 2))
  NUMER$ = FORMATNUM$(NUME, "F5")
  NUME1$ = AJUSTD$(Val(NUMER$), 7)
  Call REPLA(ZZ$, NUME1$, 1, 7)
  Call REPLA(ZZ$, DESCRI$, 10, 30)
  'List1.AddItem ZZ$
  '
  ZZ1$ = Space(128)
  Call REPLA(ZZ1$, DESCRI$, 1, 30) 'DESCRIPCION
  Call REPLA(ZZ1$, Chr$(1), 31, 1) 'ESTADO ACTIVO
  Call REPLA(ZZ1$, Chr$(1), 32, 1) 'TIPO NORMAL
  Call REPLA(ZZ1$, Chr$(0), 33, 1) 'MODO CALCULO INTERNO
  Call REPLA(ZZ1$, Chr$(0), 34, 1) 'CARACTER OPCIONAL
  Call REPLA(ZZ1$, Chr$(1), 35, 1) 'RETENCION GRAVADO
  Call REPLA(ZZ1$, Chr$(0), 36, 1) 'QUINCENA TODAS
  Call REPLA(ZZ1$, Chr$(0), 37, 1) 'RECALCULO NO
  Call REPLA(ZZ1$, Chr$(0), 38, 1) 'FILLER
  Call REPLA(ZZ1$, " ", 39, 10) 'FILLER 1
  Call REPLA(ZZ1$, " ", 49, 72) 'FORMULA
  Call REPLA(ZZ1$, " ", 121, 8) 'ABREVIATURA
  Call REGAPP("RETLEG", ZZ1$)
  '
  Call INDIRET
  Call FRESHECHO("!INDIRET")
  '
  COD$ = TRIM$(NUMER$)
  JJ& = Val(COD$)
  ZZ$ = REGLEIDO$("RETLEG", JJ&)
  DESCRI$ = TRIM$(Mid$(ZZ$, 1, 30))
  ABRE$ = TRIM$(Mid$(ZZ$, 121, 8))
  PROP$ = Mid$(ZZ$, 31, 8)
  FORMU$ = Mid$(ZZ$, 49, 72)
  JJ1 = CVS(Mid$(ZZ$, 39, 4)) '%RETENCION
  JJ2 = CVS(Mid$(ZZ$, 43, 4)) '%APORTES
  With RETDET05
    .Text8 = ABRE$ 'ABREVIATURA
    .Text7 = FORMU$ 'FORMULA
       EST% = Asc(Mid$(PROP$, 1, 1))
       ET$ = ESTADO(EST%) 'DEVUELVE EL VALOR EN $ DEL ESTADO
    .Text16 = ET$ 'ESTADO
    .Text1 = Asc(Mid$(PROP$, 3, 1)) 'CARACTER
    .Text2 = Asc(Mid$(PROP$, 4, 1)) 'MODO DE CALCULO
    .Text3 = Asc(Mid$(PROP$, 6, 1)) 'QUINCENA
    .Text4 = Asc(Mid$(PROP$, 7, 1)) 'RECALCULO
    If Val(.Text2) <> 3 Then 'SI MODO <> SUMAS FIJAS
      .Text5 = FORMATNUM(JJ1, "F5.2") '%RET
      .Text6 = FORMATNUM(JJ2, "F5.2") '%APORTES
    Else
      .Text11 = FORMATNUM(JJ1, "F9.2") 'IMP. RET
      .Text10 = FORMATNUM(JJ2, "F9.2") 'IMP. APORTES
    End If
    .Label15 = COD$
    .Text9 = DESCRI$
    MODIFIC% = 0
  End With
  '
99 Command21.Enabled = True
  '
End Sub

Private Sub Command3_Click()
   '
   Call CIERRARCH("*.*")
   Unload Me
   '
End Sub

Private Sub Command5_Click()
    '
    Call SELECHO("TAQUINLI")
    VDEVU$ = TRIM$(VALACT1$("TAQUINLI"))
    If VDEVU$ <> "" Then
       Text3.TEXT = VALACT1$("TAQUINLI")
    End If
    '
End Sub

Private Sub Command6_Click()
    '
    Call SELECHO("TARECALC")
    VDEVU$ = TRIM$(VALACT1$("TARECALC"))
    If VDEVU$ <> "" Then
       Text4.TEXT = VALACT1$("TARECALC")
    End If
    '
End Sub

Private Sub Command7_Click()
    '
Call SETULTREG("RETLEG", 16)
Exit Sub
    URE& = ULTREG&("RETLEG")
    For U& = URE& - 1 To 1 Step -1
      XX$ = REGLEIDO$("RETLEG", U&)
      Call GRAREG("RETLEG", XX$, U& + 1)
    Next U&
    Call CIERRARCH("RETLEG")
    Exit Sub
    '
    Call SELECHO("TAQUINLI")
    VDEVU$ = TRIM$(VALACT1$("TAQUINLI"))
    If VDEVU$ <> "" Then
       Text5.TEXT = VALACT1$("TAQUINLI")
    End If
    '
End Sub

Private Sub Command8_Click()
    '
    Call SELECHO("TARECALC")
    VDEVU$ = TRIM$(VALACT1$("TARECALC"))
    If VDEVU$ <> "" Then
       Text6.TEXT = VALACT1$("TARECALC")
    End If
    '
End Sub

Private Sub Command9_Click()
  Command9.Enabled = False
  Screen.MousePointer = 11
  'VALIDACIONES
  If TRIM$(Text9) = "" Then
    Call MENSERR(24, "Descripción no Válida.")
    GoTo 99
  End If
  '
  If TRIM(RECFILT("TAMOCRET", 1, MKI$(Val(Text2)))) = "" Then
    Call MENSERR(24, "Modo de Calculo no Válido.")
    GoTo 99
  End If
  '
  If TRIM(RECFILT("TACARAC", 1, MKI$(Val(Text1)))) = "" Then
    Call MENSERR(24, "Carácter no Válido.")
    GoTo 99
  End If
  '
  If TRIM(RECFILT("TAQUINLI", 1, MKI$(Val(Text3)))) = "" Then
    Call MENSERR(24, "Quincena no Válida.")
    GoTo 99
  End If
  '
  If TRIM(RECFILT("TARECALC", 1, MKI$(Val(Text4)))) = "" Then
    Call MENSERR(24, "Modo de Recalculo no Válido.")
    GoTo 99
  End If
  '
  If TRIM(RECFILT("TASTATUS", 1, Text16)) = "" Then
    Call MENSERR(24, "Estado no Válido.")
    GoTo 99
  End If
  REGI& = Val(Label15)
  ZZ1$ = REGLEIDO("RETLEG", REGI&)
  '
  DESCRI$ = TRIM$(Text9)
  ABREV$ = TRIM$(Text8)
  FORMU$ = Text7: If Val(Text2) = 1 Then FORMU$ = "" ' INFORME IMPORTE
  ESTA1% = ESTADO1%(TRIM(Text16))
  RETE = Val(Text5)
  APO = Val(Text6)
  IMPRETE = Val(Text11)
  IMPAPO = Val(Text10)
  '
  Call REPLA(ZZ1$, DESCRI$, 1, 30) 'DESCRIPCION
  Call REPLA(ZZ1$, Chr(ESTA1%), 31, 1) 'ESTADO ACTIVO
  Call REPLA(ZZ1$, Chr$(Val(Text1)), 33, 1) 'CARACTER
  Call REPLA(ZZ1$, Chr$(Val(Text2)), 34, 1) 'MODo DE CALCULO
  Call REPLA(ZZ1$, Chr$(Val(Text3)), 36, 1) 'QUINCENA TODAS
  Call REPLA(ZZ1$, Chr$(Val(Text4)), 37, 1) 'RECALCULO NO
  Call REPLA(ZZ1$, Chr$(0), 38, 1) 'FILLER
  If Val(Text2) <> 3 Then 'si MODO <> DE SUMAS FIJAS
    Call REPLA(ZZ1$, MKS$(RETE), 39, 4) '% RETENCION
    Call REPLA(ZZ1$, MKS$(APO), 43, 4) '% APORTES
  Else
    Call REPLA(ZZ1$, MKS$(IMPRETE), 39, 4) 'IMP. RETENCION
    Call REPLA(ZZ1$, MKS$(IMPAPO), 43, 4) 'IMP. APORTES
  End If
  Call REPLA(ZZ1$, " ", 47, 2) 'FILLER 1
  Call REPLA(ZZ1$, FORMU$, 49, 72) 'FORMULA
  Call REPLA(ZZ1$, ABREV$, 121, 8) 'ABREVIATURA
  Call GRAREG("RETLEG", ZZ1$, REGI&)
  Call SETULTREG("RETLEG", ULTREG&("RETLEG"))
  '
  Call INDIRET
  Call FRESHECHO("!INDIRET") 'PARA REFRESCAR EL LIST DE RET Y APORTES
  '
99 Screen.MousePointer = 1
  Command9.Enabled = True
  '
End Sub

Private Sub Form_Load()
  '
  Call CENTRAFORM(Me)
  Call INDIRET
  '
End Sub

Private Sub Text1_Change()
  '
  Label1 = ECOARCH$("TACARAC", MKI$(Val(Text1)))
  '
End Sub

Private Sub Text11_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
     Text10.SetFocus
   End If
End Sub

Private Sub Text11_LostFocus()
   Text11 = TRIM$(FORMATNUM$(Val(Text11), "F11.2"))
   If Val(Text2) = 3 Then
     FORMUX$ = TRIM$(Text11)
     Text7 = FORMUX$
   End If
End Sub

Private Sub Text10_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
     Command9.SetFocus
   End If
End Sub

Private Sub Text10_LostFocus()
   Text10 = TRIM$(FORMATNUM$(Val(Text10), "F11.2"))
End Sub

Private Sub Text16_Change()
  '
  Text16 = UCase(Text16)
  Label23 = ECOARCH$("TASTATUS", Text16)
  '
End Sub

Private Sub Text2_Change()
  '
  Label3 = ECOARCH$("TAMOCRET", MKI(Val(Text2)))
  Call STATEXT 'TEXT EDITABLES SEGUN MODO SELECCIONADO
  '
End Sub

Private Sub Text3_Change()
  '
  Label5 = ECOARCH$("TAQUINLI", MKI$(Val(Text3)))
  '
End Sub

Private Sub Text4_Change()
  Label7 = ECOARCH$("TARECALC", MKI$(Val(Text4)))
End Sub
Function ESTADO1%(ESTA$) 'DEVUELVE EL VALOR EN $ DEL ESTADO
  Select Case ESTA$
    Case "S"
      ESTADO1% = 0 'SUSPENDIDO
    Case "A"
      ESTADO1% = 1 'ACTIVO
    Case "F"
      ESTADO1% = 2 'FANTASMA
  End Select
End Function
'
Sub STATEXT() 'TEXT BOX'S EDITABLES SEGUN MODO SELECCIONADO
  '
  MO1% = Val(Text2)
  Select Case MO1%
    Case 0 'NI RET NI A.PATR.
      Frame2.Visible = False
      Text6.Enabled = False
      Text5.Enabled = False
      Text7.Enabled = True
    Case 1 'NI RET NI A.PATR
      Frame2.Visible = False
      Text6.Enabled = False
      Text5.Enabled = False
      Text7.Enabled = False
    Case 2 'SOLO RET
      Frame2.Visible = False
      Text6.Enabled = True '%APTR
      Text5.Enabled = True '%RET
      Text7.Enabled = False
    Case 3
      Frame2.Visible = True
      Text7.Enabled = False
  End Select
  '
End Sub

Private Sub Text5_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
     Text6.SetFocus
   End If
End Sub

Private Sub Text6_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
     Command9.SetFocus
   End If
End Sub

Private Sub Text5_LostFocus()
   Text5 = FORMATNUM$(Val(Text5), "F6.2")
   If Val(Text2) = 2 Then
     FORMUX$ = TRIM$(Text5) + " x RM0 /100"
     Text7 = FORMUX$
   End If
End Sub

Private Sub Text6_LostFocus()
   Text6 = FORMATNUM$(Val(Text6), "F6.2")
End Sub

Private Sub Text9_Change()
   MODIFIC% = 1
End Sub

Function ESTADO$(EST%) 'DEVUELVE EL VALOR EN $ DEL ESTADO
  Select Case EST%
    Case 0
      ESTADO$ = "S" 'SUSPENDIDO
    Case 1
      ESTADO$ = "A" 'ACTIVO
    Case 2
      ESTADO$ = "F" 'FANTASMA
  End Select
End Function

