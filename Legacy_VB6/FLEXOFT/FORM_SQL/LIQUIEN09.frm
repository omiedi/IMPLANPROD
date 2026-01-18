VERSION 5.00
Begin VB.Form LIQUIEN09 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0C0A0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Liquidacion de Entregas de Sucursal"
   ClientHeight    =   8325
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   11910
   FillStyle       =   0  'Solid
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8325
   ScaleWidth      =   11910
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0A0&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   8520
      Left            =   0
      TabIndex        =   0
      Top             =   -120
      Width           =   12000
      Begin VB.PictureBox Picture14 
         BackColor       =   &H00400000&
         Height          =   8220
         Left            =   10900
         ScaleHeight     =   8160
         ScaleWidth      =   1245
         TabIndex        =   22
         Top             =   120
         Width           =   1305
         Begin VB.CommandButton Command1 
            Caption         =   "Liq."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   660
            Left            =   120
            Picture         =   "LIQUIEN09.frx":0000
            Style           =   1  'Graphical
            TabIndex        =   28
            ToolTipText     =   "Consulta por Liquidacion"
            Top             =   2040
            Width           =   750
         End
         Begin VB.CommandButton Command11 
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
            Height          =   660
            Left            =   120
            Picture         =   "LIQUIEN09.frx":014A
            Style           =   1  'Graphical
            TabIndex        =   27
            ToolTipText     =   "Apertura de Nueva Cuenta de Cliente"
            Top             =   1320
            Width           =   750
         End
         Begin VB.CommandButton BOTNEXT 
            Caption         =   "Next"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   660
            Left            =   120
            Picture         =   "LIQUIEN09.frx":058C
            Style           =   1  'Graphical
            TabIndex        =   26
            ToolTipText     =   "Continúa Carga de Recibo."
            Top             =   600
            Width           =   750
         End
         Begin VB.CommandButton Command2 
            Caption         =   "Cancel"
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
            Left            =   120
            Picture         =   "LIQUIEN09.frx":0896
            TabIndex        =   25
            ToolTipText     =   "Abandona la Aplicación"
            Top             =   120
            Width           =   750
         End
         Begin VB.PictureBox Picture1 
            BackColor       =   &H000000FF&
            Height          =   135
            Left            =   120
            ScaleHeight     =   75
            ScaleWidth      =   645
            TabIndex        =   23
            Top             =   7320
            Width           =   705
            Begin VB.Frame Frame3 
               BackColor       =   &H00FF0000&
               BorderStyle     =   0  'None
               Height          =   500
               Left            =   0
               TabIndex        =   24
               Top             =   0
               Width           =   12000
            End
         End
         Begin VB.Image Image1 
            Height          =   390
            Left            =   -240
            Picture         =   "LIQUIEN09.frx":0BA0
            Top             =   7560
            Width           =   1515
         End
      End
      Begin VB.TextBox Text6 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   6195
         TabIndex        =   14
         TabStop         =   0   'False
         Text            =   "X"
         ToolTipText     =   "Configuracion de Funcionamiento - Doble-Click Aquí"
         Top             =   315
         Width           =   435
      End
      Begin VB.PictureBox Picture2 
         BackColor       =   &H00E0E0E0&
         Height          =   6570
         Left            =   200
         ScaleHeight     =   6510
         ScaleWidth      =   10395
         TabIndex        =   6
         Top             =   1680
         Width           =   10450
         Begin VB.Frame Frame2 
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   0  'None
            Caption         =   "Frame4"
            Height          =   330
            Left            =   0
            TabIndex        =   9
            Top             =   0
            Width           =   10515
            Begin VB.Line Line1 
               X1              =   10080
               X2              =   10080
               Y1              =   0
               Y2              =   1050
            End
            Begin VB.Line Line2 
               X1              =   6720
               X2              =   6720
               Y1              =   0
               Y2              =   1050
            End
            Begin VB.Label Label9 
               BackStyle       =   0  'Transparent
               Caption         =   $"LIQUIEN09.frx":2AC2
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Left            =   0
               TabIndex        =   10
               Top             =   45
               Width           =   10410
            End
            Begin VB.Line Line8 
               X1              =   1320
               X2              =   1320
               Y1              =   0
               Y2              =   1050
            End
            Begin VB.Line Line9 
               X1              =   3000
               X2              =   3000
               Y1              =   0
               Y2              =   1050
            End
            Begin VB.Line Line10 
               X1              =   3840
               X2              =   3840
               Y1              =   0
               Y2              =   1050
            End
            Begin VB.Line Line11 
               X1              =   8880
               X2              =   8880
               Y1              =   0
               Y2              =   1170
            End
         End
         Begin VB.Frame Frame5 
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   0  'None
            Caption         =   "Frame3"
            Height          =   6300
            Left            =   0
            TabIndex        =   7
            Top             =   315
            Width           =   10515
            Begin VB.ListBox List2 
               BeginProperty Font 
                  Name            =   "Fixedsys"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   6180
               Left            =   -30
               Style           =   1  'Checkbox
               TabIndex        =   8
               Top             =   0
               Width           =   10455
            End
            Begin VB.ListBox List3 
               BeginProperty Font 
                  Name            =   "Fixedsys"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   2085
               Left            =   -30
               TabIndex        =   12
               Top             =   0
               Width           =   10450
            End
            Begin VB.Line Line5 
               X1              =   0
               X2              =   10440
               Y1              =   6180
               Y2              =   6180
            End
            Begin VB.Label Label11 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               BeginProperty Font 
                  Name            =   "Fixedsys"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Left            =   0
               TabIndex        =   11
               Top             =   0
               Width           =   1695
            End
            Begin VB.Line Line14 
               BorderWidth     =   2
               X1              =   0
               X2              =   10395
               Y1              =   2085
               Y2              =   2085
            End
            Begin VB.Line Line15 
               X1              =   0
               X2              =   10450
               Y1              =   6180
               Y2              =   6180
            End
         End
         Begin VB.Line Line4 
            X1              =   0
            X2              =   6240
            Y1              =   6480
            Y2              =   6480
         End
      End
      Begin VB.Frame Frame8 
         BackColor       =   &H00C0C0A0&
         BorderStyle     =   0  'None
         Caption         =   "Frame7"
         Height          =   1335
         Left            =   2640
         TabIndex        =   13
         Top             =   7030
         Visible         =   0   'False
         Width           =   2600
      End
      Begin VB.Label Label14 
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
         Height          =   300
         Left            =   9600
         TabIndex        =   21
         Top             =   315
         Width           =   1095
      End
      Begin VB.Label Label13 
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
         Height          =   300
         Left            =   8280
         TabIndex        =   20
         Top             =   1155
         Width           =   2415
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
         Height          =   300
         Left            =   2400
         TabIndex        =   19
         Top             =   1155
         Width           =   2775
      End
      Begin VB.Label Label10 
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
         Height          =   300
         Left            =   1440
         TabIndex        =   18
         Top             =   1155
         Width           =   855
      End
      Begin VB.Label Label8 
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
         Height          =   300
         Left            =   1440
         TabIndex        =   17
         Top             =   735
         Width           =   3735
      End
      Begin VB.Label Label7 
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
         Height          =   300
         Left            =   1440
         TabIndex        =   16
         Top             =   315
         Width           =   3735
      End
      Begin VB.Label Label4 
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
         Height          =   300
         Left            =   240
         TabIndex        =   15
         Top             =   1140
         Width           =   855
      End
      Begin VB.Line Line27 
         X1              =   10665
         X2              =   10665
         Y1              =   1680
         Y2              =   8230
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Documento no Válido como Factura"
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
         Left            =   5355
         TabIndex        =   5
         ToolTipText     =   "Configuracion de Funcionamiento - Doble-Click Aquí"
         Top             =   1050
         Width           =   2115
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Referencia:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   7560
         TabIndex        =   4
         Top             =   870
         Width           =   1380
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   8760
         TabIndex        =   3
         Top             =   330
         Width           =   750
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Cuenta:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   315
         TabIndex        =   1
         Top             =   735
         Width           =   960
      End
   End
   Begin VB.Line Line3 
      X1              =   0
      X2              =   0
      Y1              =   0
      Y2              =   1050
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Pos.IVA:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   960
   End
End
Attribute VB_Name = "LIQUIEN09"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim LIST2ANT$, AUTONUME%, TICAMBX#
Dim ACCECALC$

Private Sub BOTNEXT_Click()
   BOTNEXT.Enabled = False
   Call APROLIQ
   BOTNEXT.Enabled = True
End Sub


Private Sub Command1_Click()
    Command1.Enabled = False
    Call FINAL("*CODELIQ")
    Command1.Enabled = True
End Sub

Private Sub Command11_Click()
    Command11.Enabled = False
    Call CIERRARCH("*.*")
    ACONEC$ = "AMACLI02"
    If EXISTE%(ACONEC$ + ".EXE") < 1 Then ACONEC$ = "AMACLI01"
    If EXISTE%(ACONEC$ + ".EXE") < 1 Then ACONEC$ = "AMACLI/NEW"
    Call CONECRUN(ACONEC$, "Alta de Nueva Cuenta de Cliente")
    Call FRESHECHO("DEUDOR1")
    Command11.Enabled = True
End Sub

Private Sub Command2_Click()
   Call CIERRARCH("*.*")
   List2.Clear
   Hide
End Sub

Private Sub Form_Load()
    '
    Call CENTRAFORM(Me)
    EPAC$ = TRIM$(EMPREAC$)
    If EPAC$ <> "" Then
        Caption = Caption + "  -  " + EPAC$
    End If
    '
End Sub

Private Sub Form_Resize()
   If WindowState = 1 Then
     Call CIERRARCH("*.*")
     Call BLOQEXE("*")
     Call LIBARCH("*.*")
     Reset
     Call BLOQEXE(App.EXEName)
   End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
   List2.Clear
   Cancel = 0
   Call CIERRARCH("*.*")
   Call FINAL("")
End Sub
'

Private Sub APROLIQ()
   '
   JJ& = 0
   FF% = FECHANUM(Label14)
   For U& = 1 To List2.ListCount
     If List2.Selected(U& - 1) = True Then
       RELE& = Val(Mid$(List2.List(U& - 1), 121, 8))
       CALI = Val(Mid$(List2.List(U& - 1), 72, 10))
       x$ = REGLEIDO$("ENTRECOM", RELE&)
       Call REPLA(x$, Label13, 195, 30)
       Call REPLA(x$, MKI$(FF%), 225, 2)
       Call REPLA(x$, MKS$(CALI), 227, 4)
       Call REPLA(x$, MKS$(RELE&), 231, 4)
       JJ& = JJ& + 1
       Call GRAREG("!ENTRECOM", x$, JJ&)
     End If
   Next U&
   Call SETULTREG("!ENTRECOM", JJ&)
   If JJ& < 1 Then
     Call MENSERR(24, "Falta Seleccionar Items Liquidados")
     Exit Sub
   End If
   '
   HAYLI% = 0
   ILIQUI% = 1 + ULTREG&("INDILIQ")
   VTB% = VENTABU%("VENTRECO/BORDEIZQ=W" + TRIM$(Str$(Left + Frame1.Left + Picture2.Left + List2.Left + 70)) + "/BORDESUP=W" + TRIM$(Str$(Top + Frame1.Top + Picture2.Top + Frame2.Top)))
   If VTB% >= 0 Then
     For U& = 1 To ULTREG&("!ENTRECOM")
       x$ = REGLEIDO$("!ENTRECOM", U&)
       RELE& = CVS(Mid$(x$, 231, 4))
       CALI = CVS(Mid$(x$, 227, 4))
       If CALI > CVD(Mid$(x$, 183, 8)) Then
         CALI = CVD(Mid$(x$, 183, 8))
         Call REPLA(x$, MKS$(CALI), 227, 4)
       End If
       If CALI >= 0.05 Then
         SALPE# = CVD(Mid$(x$, 183, 8)) - CALI
         If SALPE# >= 0.05 Then
           y$ = REGLEIDO$("ENTRECOM", RELE&)
           Call REPLA(y$, MKD$(SALPE#), 183, 8)
           Call REGAPP("ENTRECOM", y$)
           Call CIERRARCH("ENTRECOM")
           CALID# = CALI
           Call REPLA(x$, MKD$(CALID#), 183, 8)
         End If
         Call REPLA(x$, MKS$(0), 231, 4)
         Call REPLA(x$, MKI$(ILIQUI%), 235, 2)
         Call GRAREG("ENTRECOM", x$, RELE&)
         HAYLI% = 1
       End If
     Next U&
     Call CIERRARCH("ENTRECOM")
     If HAYLI% = 1 Then
       Z$ = MKI$(ILIQUI%) + Label13
       Call REGAPP("INDILIQ", Z$)
     End If
     Hide
   End If
   '
End Sub

Private Sub Label9_DblClick()
   Call APROLIQ
End Sub
