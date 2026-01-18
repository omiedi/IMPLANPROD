VERSION 5.00
Begin VB.Form OPORFA09 
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Configuración y Formulario de Orden de Fabricación"
   ClientHeight    =   4575
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   6375
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
   ScaleHeight     =   4575
   ScaleWidth      =   6375
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Central de Pesadas"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1230
      Left            =   210
      TabIndex        =   19
      Top             =   210
      Width           =   5010
      Begin VB.TextBox Text3 
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
         Height          =   285
         Left            =   3780
         MaxLength       =   3
         TabIndex        =   21
         Top             =   420
         Width           =   810
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
         Height          =   285
         Left            =   4590
         TabIndex        =   20
         Top             =   420
         Width           =   175
      End
      Begin VB.Label Label24 
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
         Height          =   285
         Left            =   165
         TabIndex        =   23
         Top             =   765
         Width           =   4605
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Depósito Asociado:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1965
         TabIndex        =   22
         Top             =   495
         Width           =   1680
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Nota de Pedido:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1350
      Left            =   210
      TabIndex        =   12
      Top             =   1575
      Width           =   5010
      Begin VB.CommandButton Command7 
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
         Height          =   730
         Left            =   3975
         Picture         =   "OPORFA09.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   390
         Width           =   855
      End
      Begin VB.TextBox Text7 
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
         Left            =   195
         TabIndex        =   15
         TabStop         =   0   'False
         Text            =   " "
         Top             =   840
         Width           =   3705
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
         Height          =   300
         Left            =   3720
         TabIndex        =   14
         Top             =   420
         Width           =   175
      End
      Begin VB.TextBox Text6 
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
         Left            =   2355
         TabIndex        =   13
         Text            =   " "
         Top             =   420
         Width           =   1380
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Formulario Principal:"
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
         Left            =   -30
         TabIndex        =   17
         Top             =   480
         Width           =   2325
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00004040&
      Height          =   3060
      Left            =   5355
      ScaleHeight     =   3000
      ScaleWidth      =   1050
      TabIndex        =   5
      Top             =   -60
      Width           =   1110
      Begin VB.PictureBox MARCOBARRA 
         BackColor       =   &H000000FF&
         Height          =   135
         Left            =   165
         ScaleHeight     =   75
         ScaleWidth      =   585
         TabIndex        =   8
         Top             =   2175
         Width           =   650
         Begin VB.Frame BARRATRAZA 
            BackColor       =   &H00FF0000&
            BorderStyle     =   0  'None
            Height          =   500
            Left            =   0
            TabIndex        =   9
            Top             =   -180
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
         Left            =   150
         Picture         =   "OPORFA09.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   120
         Width           =   650
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
         Left            =   150
         Picture         =   "OPORFA09.frx":0454
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   900
         Width           =   650
      End
      Begin VB.Image Image2 
         Height          =   390
         Left            =   -255
         Picture         =   "OPORFA09.frx":075E
         ToolTipText     =   "Acceso Directo a Otras Aplicaciones"
         Top             =   2385
         Width           =   1515
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Orden de Fabricación"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1410
      Left            =   210
      TabIndex        =   0
      Top             =   3045
      Width           =   5970
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
         Height          =   735
         Left            =   3970
         Picture         =   "OPORFA09.frx":2680
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   480
         Width           =   840
      End
      Begin VB.Frame Frame5 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Código"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   825
         Left            =   4935
         TabIndex        =   10
         Top             =   400
         Width           =   855
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
            Left            =   195
            TabIndex        =   11
            Text            =   "OF"
            Top             =   285
            Width           =   450
         End
      End
      Begin VB.TextBox Text1 
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
         Left            =   2340
         TabIndex        =   4
         Text            =   " "
         Top             =   510
         Width           =   1380
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
         Left            =   3705
         TabIndex        =   3
         Top             =   510
         Width           =   175
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
         Height          =   300
         Left            =   195
         TabIndex        =   2
         TabStop         =   0   'False
         Text            =   " "
         Top             =   930
         Width           =   3705
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Formulario Principal:"
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
         Left            =   -105
         TabIndex        =   1
         Top             =   615
         Width           =   2325
      End
   End
End
Attribute VB_Name = "OPORFA09"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub command1_Click()
    Call CIERRARCH("*.*")
    Hide
End Sub

Private Sub Command10_Click()
  'OT-05-0728-WAR-08/11/05///
  Call TRALOCAL("SECPROD", "SECPROD")
  '
  ATRI$ = "BORDEIZQ=W" + TRIM$(Str$(Left + 200))
  ATRI$ = ATRI$ + "/BORDESUP=W" + TRIM$(Str$(Top + 850))
  VTB% = VENTABU%("SECPROD1/" + ATRI$)
  If VTB% >= 0 Then
     JJ& = 0
    For U& = 1 To ULTREG&("!SECPROD")
       X$ = REGLEIDO$("!SECPROD", U&)
       JJ& = JJ& + 1
       Call GRAREG("SECPROD", X$, JJ&)
    Next U&
    Call SETULTREG("SECPROD", JJ&)
    Call CIERRARCH("SECPROD")
  End If
  'OT-05-0728-WAR-FIN///
End Sub

Private Sub Command2_Click()
    '
    Call GRACONTROL("", "FORMORFA", Text1.TEXT)
    Call GRACONTROL("", "CODMORFA", Text5.TEXT)
    Call GRACONTROL("", "FORNOPE", Text6.TEXT)
    Call GRACONTROL("STOCKS", "DEPCENPE", Text3.TEXT)
    '
    Call CIERRARCH("*.*")
    GENORFA09.MostrarNumOF
    '
    Hide
    '
End Sub

Private Sub Command3_Click()
    Call SELECHO("INDIFRM")
    vdevu$ = VALACT1$("INDIFRM")
    If vdevu$ <> "" Then
        Text1.TEXT = vdevu$
        Text2.TEXT = VALACT2$("INDIFRM")
    End If
End Sub
'


Private Sub Command4_Click()
    Call SELECHO("INDIFRM")
    vdevu$ = VALACT1$("INDIFRM")
    If vdevu$ <> "" Then
        Text6.TEXT = vdevu$
        Text7.TEXT = VALACT2$("INDIFRM")
    End If
End Sub

Private Sub Command5_Click()
    ACONEC$ = "FLFORMS": If EXISTE%(ACONEC$ + ".EXE") < 1 Then ACONEC$ = "PRIFORMS"
    Call CONECRUN(ACONEC$ + "/" + TRIM$(Text1.TEXT), "Formularios de Impresion")
End Sub

Private Sub Command6_Click()
   '
   Command6.Enabled = False
   Call SELECHO("TADEPOSI")
   NPROV$ = VALACT1$("TADEPOSI")
   If NPROV$ <> "" Then
      Text3.TEXT = NPROV$
   End If
   Command6.Enabled = True
   '
End Sub

Private Sub Command8_Click()
   PRINTFORM
End Sub

Private Sub Command9_Click()
   Command9.Enabled = False
   OPVALEMAT.Show 1
   Command9.Enabled = True
End Sub

Private Sub Command7_Click()
    Command7.Enabled = False
    Call CONECRUN("FLFORMS/" + TRIM$(Text6.TEXT), "Formularios de Impresion")
    Command7.Enabled = True
End Sub

Private Sub Form_Load()
    '
    Text1.TEXT = CONTROL$("", "FORMORFA")
    Text2.TEXT = ECOARCH$("INDIFRM", Text1.TEXT)
    Text5.TEXT = CONTROL("", "CODMORFA")
    Text3.TEXT = CONTROL$("STOCKS", "DEPCENPE")
    '
    Text6.TEXT = CONTROL$("", "FORNOPE")
    Text7.TEXT = ECOARCH$("INDIFRM", Text6.TEXT)
    '
End Sub

Private Sub Text1_DblClick()
    Call SELECHO("INDIFRM")
    vdevu$ = VALACT1$("INDIFRM")
    If vdevu$ <> "" Then
        Text1.TEXT = vdevu$
        Text2.TEXT = VALACT2$("INDIFRM")
    End If
End Sub

Private Sub Text2_CLICK()
    Text1.SetFocus
End Sub

Private Sub Text4_CLICK()
    Text2.SetFocus
End Sub

Private Sub Text3_Change()
    If XVALO(Text3) < 256 Then
       Label24 = ECOARCH("TADEPOSI", Chr$(XVALO(Text3)))
    End If
End Sub

Private Sub Text3_DBLCLICK()
    Call SELECHO("INDIFRM")
    vdevu$ = VALACT1$("INDIFRM")
    If vdevu$ <> "" Then
        Text3.TEXT = vdevu$
        text4.TEXT = VALACT2$("INDIFRM")
    End If
End Sub

Private Sub Text6_DblClick()
    Call SELECHO("INDIFRM")
    vdevu$ = VALACT1$("INDIFRM")
    If vdevu$ <> "" Then
        Text6.TEXT = vdevu$
        Text7.TEXT = VALACT2$("INDIFRM")
    End If
End Sub
