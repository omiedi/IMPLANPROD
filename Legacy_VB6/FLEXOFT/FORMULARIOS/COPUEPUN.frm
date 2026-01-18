VERSION 5.00
Begin VB.Form COPUEPUN 
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Complementos ISO 9000"
   ClientHeight    =   5325
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   6660
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
   ScaleHeight     =   5325
   ScaleWidth      =   6660
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox MARCOBARRA 
      BackColor       =   &H000000FF&
      Height          =   135
      Left            =   5600
      ScaleHeight     =   75
      ScaleWidth      =   795
      TabIndex        =   28
      Top             =   3780
      Width           =   855
      Begin VB.Frame BARRATRAZA 
         BackColor       =   &H00FF0000&
         BorderStyle     =   0  'None
         Height          =   500
         Left            =   0
         TabIndex        =   29
         Top             =   0
         Width           =   12000
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Listados por Impresora"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1485
      Left            =   210
      TabIndex        =   14
      Top             =   3675
      Width           =   5160
      Begin VB.PictureBox Picture2 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   525
         ScaleHeight     =   585
         ScaleWidth      =   1845
         TabIndex        =   25
         Top             =   525
         Width           =   1905
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
            Height          =   600
            Left            =   0
            Picture         =   "COPUEPUN.frx":0000
            Style           =   1  'Graphical
            TabIndex        =   26
            ToolTipText     =   "A-B-M-L Base de Datos de Items de Stock"
            Top             =   0
            Width           =   560
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   "Plan de Control"
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
            Left            =   840
            TabIndex        =   27
            Top             =   105
            Width           =   960
         End
      End
      Begin VB.PictureBox Picture1 
         BackColor       =   &H80000016&
         Height          =   645
         Left            =   2730
         ScaleHeight     =   585
         ScaleWidth      =   1845
         TabIndex        =   22
         Top             =   525
         Width           =   1905
         Begin VB.CommandButton Command5 
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
            Picture         =   "COPUEPUN.frx":030A
            Style           =   1  'Graphical
            TabIndex        =   23
            ToolTipText     =   "Fórmulas y Estructuras de Conjuntos"
            Top             =   0
            Width           =   600
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Puesta a Punto"
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
            Left            =   840
            TabIndex        =   24
            Top             =   105
            Width           =   750
         End
      End
      Begin VB.Frame Frame5 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Listados de Control"
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
         Left            =   2625
         TabIndex        =   15
         Top             =   4830
         Width           =   4005
         Begin VB.CommandButton Command9 
            Caption         =   "Inventarios"
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
            Left            =   210
            TabIndex        =   21
            Top             =   420
            Width           =   1590
         End
         Begin VB.CommandButton Command8 
            Caption         =   "Negativos"
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
            Left            =   2205
            TabIndex        =   20
            Top             =   420
            Width           =   1590
         End
         Begin VB.CommandButton Command10 
            Caption         =   "Consignación"
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
            Left            =   210
            TabIndex        =   19
            Top             =   1155
            Width           =   1590
         End
         Begin VB.CommandButton Command11 
            Caption         =   "Críticos"
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
            Left            =   2205
            TabIndex        =   18
            Top             =   1155
            Width           =   1590
         End
         Begin VB.CommandButton Command12 
            Caption         =   "Estadísticas"
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
            Left            =   210
            TabIndex        =   17
            Top             =   1890
            Width           =   1590
         End
         Begin VB.CommandButton Command13 
            Caption         =   "Análisis A-B-C"
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
            Left            =   2205
            TabIndex        =   16
            Top             =   1890
            Width           =   1590
         End
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Control y Puesta a Punto"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3180
      Left            =   210
      TabIndex        =   2
      Top             =   315
      Width           =   5160
      Begin VB.Frame Frame1 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Actualizacion y Consulta"
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
         Left            =   210
         TabIndex        =   7
         Top             =   1680
         Width           =   4700
         Begin VB.PictureBox Picture9 
            BackColor       =   &H80000016&
            Height          =   645
            Left            =   315
            ScaleHeight     =   585
            ScaleWidth      =   1845
            TabIndex        =   11
            Top             =   420
            Width           =   1905
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
               Height          =   600
               Left            =   0
               Picture         =   "COPUEPUN.frx":0614
               Style           =   1  'Graphical
               TabIndex        =   12
               ToolTipText     =   "A-B-M-L Base de Datos de Items de Stock"
               Top             =   0
               Width           =   560
            End
            Begin VB.Label Label10 
               BackStyle       =   0  'Transparent
               Caption         =   "Plan de Control"
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
               Left            =   840
               TabIndex        =   13
               Top             =   105
               Width           =   960
            End
         End
         Begin VB.PictureBox Picture10 
            BackColor       =   &H80000016&
            Height          =   645
            Left            =   2520
            ScaleHeight     =   585
            ScaleWidth      =   1845
            TabIndex        =   8
            Top             =   420
            Width           =   1905
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
               Picture         =   "COPUEPUN.frx":091E
               Style           =   1  'Graphical
               TabIndex        =   9
               ToolTipText     =   "Fórmulas y Estructuras de Conjuntos"
               Top             =   0
               Width           =   600
            End
            Begin VB.Label Label11 
               BackStyle       =   0  'Transparent
               Caption         =   "Puesta a Punto"
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
               Left            =   840
               TabIndex        =   10
               Top             =   105
               Width           =   750
            End
         End
      End
      Begin VB.TextBox Text2 
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
         Left            =   210
         MultiLine       =   -1  'True
         TabIndex        =   5
         Top             =   945
         Width           =   4700
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
         Height          =   315
         Left            =   2520
         TabIndex        =   4
         Top             =   495
         Width           =   2220
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
         Height          =   315
         Left            =   4725
         TabIndex        =   3
         Top             =   495
         Width           =   175
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Código: "
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
         Left            =   1260
         TabIndex        =   6
         Top             =   525
         Width           =   1275
      End
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
      Height          =   855
      Left            =   5600
      Picture         =   "COPUEPUN.frx":0C28
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1365
      Width           =   855
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
      Height          =   855
      Left            =   5600
      Picture         =   "COPUEPUN.frx":0F32
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   420
      Width           =   855
   End
   Begin VB.Image Image2 
      Height          =   555
      Left            =   5355
      Picture         =   "COPUEPUN.frx":107C
      Stretch         =   -1  'True
      Top             =   4590
      Width           =   5250
   End
End
Attribute VB_Name = "COPUEPUN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
    Call CIERRARCH("*.*")
    Hide
End Sub

Private Sub Command2_Click()
    Call SELECHO("MASTER")
    VDEVU$ = VALACT1$("MASTER")
    If VDEVU$ <> "" Then
        COD$ = VDEVU$
        Text1.Text = VDEVU$
        Text2.Text = DESCRIT0$(CODINT%(COD$))
      Else
        COD$ = ""
        Text1.Text = ""
        Text2.Text = ""
    End If
End Sub

Private Sub Command3_Click()
   Command3.Enabled = False
   COD$ = TRIM$(Text1.Text)
   CII% = CODINT%(COD$)
   If CII% < 1 Then
      Call MENSERR(24, "Falta Ingresar Código de Artículo")
      GoTo 99
   End If
   '
   TITUPA$ = "PUESTA A PUNTO - " + TRIM$(COD$) + " - " + TRIM$(DESCRIT$(CII%))
   ' LEER DATOR REGISTRADOS
10 RFF$ = RECFILT$("PUEPUNTO", 1, MKI$(CII%))
   JJ& = 0
   For UK& = 1 To Len(RFF$) Step 4
      U& = CVS(Mid$(RFF$, UK&, 4))
      If U& > 0 Then
         If U& <= ULTREG&("PUEPUNTO") Then
            X$ = REGLEIDO$("PUEPUNTO", U&)
            If CVI(Left$(X$, 2)) <> CII% Then
               X$ = REGLEIDO$("PUEPUNTO", 1)
               Call GRAREG("PUEPUNTO", X$, 1)
               Call CIERRARCH("PUEPUNTO")
               GoTo 10
            End If
            JJ& = JJ& + 1
            Call GRAREG("!PUEPUNTO", X$, JJ&)
         End If
      End If
   Next UK&
   Call SETULTREG("!PUEPUNTO", JJ&)
   Call CIERRARCH("!PUEPUNTO")
   '
   VTB% = VENTABU%("PUEPUNTO/TITUPAN=" + TITUPA$)
   If VTB% < 0 Then GoTo 99
   '
   'GRABAR NUEVA INFORMACION
   KK& = 1
   For JJ& = 1 To ULTREG&("!PUEPUNTO")
      X$ = REGLEIDO$("!PUEPUNTO", JJ&)
      Call REPLA(X$, MKI$(CII%), 1, 2)
      REDESTI& = 0
      If KK& < Len(RFF$) Then
         REDESTI& = CVS(Mid$(RFF$, KK&, 4))
      End If
      If REDESTI& > 0 Then
         If REDESTI& <= ULTREG&("PUEPUNTO") Then
            Call GRAREG("PUEPUNTO", X$, REDESTI&)
            GoTo 19
         End If
      End If
      Call REGAPP("PUEPUNTO", X$)
      '
19    KK& = KK& + 4
   Next JJ&
   '
   For KK1& = KK& To Len(RFF$) Step 4
      REDESTI& = CVS(Mid$(RFF$, KK1&, 4))
      Call GRAREG("PUEPUNTO", String$(256, 0), REDESTI&)
   Next KK1&
   '
   Call CIERRARCH("PUEPUNTO")
   '
99 Command3.Enabled = True
End Sub

Private Sub Command4_Click()
   Command4.Enabled = False
   COD$ = TRIM$(Text1.Text)
   CII% = CODINT%(COD$)
   If CII% < 1 Then
      Call MENSERR(24, "Falta Ingresar Código de Artículo")
      GoTo 99
   End If
   '
   TITUPA$ = "PLAN DE CONTROL - " + TRIM$(COD$) + " - " + TRIM$(DESCRIT$(CII%))
   ' LEER DATOR REGISTRADOS
10 RFF$ = RECFILT$("PLACONTR", 1, MKI$(CII%))
   JJ& = 0
   For UK& = 1 To Len(RFF$) Step 4
      U& = CVS(Mid$(RFF$, UK&, 4))
      If U& > 0 Then
         If U& <= ULTREG&("PLACONTR") Then
            X$ = REGLEIDO$("PLACONTR", U&)
            If CVI(Left$(X$, 2)) <> CII% Then
               X$ = REGLEIDO$("PLACONTR", 1)
               Call GRAREG("PLACONTR", X$, 1)
               Call CIERRARCH("PLACONTR")
               GoTo 10
            End If
            JJ& = JJ& + 1
            Call GRAREG("!PLACONTR", X$, JJ&)
         End If
      End If
   Next UK&
   Call SETULTREG("!PLACONTR", JJ&)
   Call CIERRARCH("!PLACONTR")
   '
   VTB% = VENTABU%("PLACON-1/TITUPAN=" + TITUPA$)
   If VTB% < 0 Then GoTo 99
   '
   VTB% = VENTABU%("PLACON-2/TITUPAN=" + TITUPA$)
   If VTB% < 0 Then GoTo 99
   '
   'GRABAR NUEVA INFORMACION
   KK& = 1
   For JJ& = 1 To ULTREG&("!PLACONTR")
      X$ = REGLEIDO$("!PLACONTR", JJ&)
      Call REPLA(X$, MKI$(CII%), 1, 2)
      REDESTI& = 0
      If KK& < Len(RFF$) Then
         REDESTI& = CVS(Mid$(RFF$, KK&, 4))
      End If
      If REDESTI& > 0 Then
         If REDESTI& <= ULTREG&("PLACONTR") Then
            Call GRAREG("PLACONTR", X$, REDESTI&)
            GoTo 19
         End If
      End If
      Call REGAPP("PLACONTR", X$)
      '
19    KK& = KK& + 4
   Next JJ&
   '
   For KK1& = KK& To Len(RFF$) Step 4
      REDESTI& = CVS(Mid$(RFF$, KK1&, 4))
      Call GRAREG("PLACONTR", String$(256, 0), REDESTI&)
   Next KK1&
   '
   Call CIERRARCH("PLACONTR")
   '
99 Command4.Enabled = True
End Sub

Private Sub Command5_Click()
   'Call OPNOIN("Listado de Puesta a Punto")
   Command5.Enabled = False
   Call CONECRUN("*LIPUEPU", "Listado de Puesta a Punto")
   Command5.Enabled = True
End Sub

Private Sub Command6_Click()
   'Call OPNOIN("Listado de Plan de Control")
   Command6.Enabled = False
   Call CONECRUN("*LIPLACO", "Listado de Plan de Control")
   Command6.Enabled = True
End Sub

Private Sub Text1_DBLCLICK()
    Call SELECHO("MASTER")
    VDEVU$ = VALACT1$("MASTER")
    If VDEVU$ <> "" Then
        COD$ = VDEVU$
        Text1.Text = VDEVU$
        Text2.Text = DESCRIT0$(CODINT%(COD$))
      Else
        COD$ = ""
        Text1.Text = ""
        Text2.Text = ""
    End If
End Sub

Private Sub Text2_CLICK()
    Text1.SetFocus
End Sub


