VERSION 5.00
Begin VB.Form COENTRE 
   BackColor       =   &H00EEEEC0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Control de Entregas por Articulo"
   ClientHeight    =   4200
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   4965
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
   ScaleHeight     =   4200
   ScaleWidth      =   4965
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox MARCOBARRA 
      BackColor       =   &H000000FF&
      Height          =   135
      Left            =   3950
      ScaleHeight     =   75
      ScaleWidth      =   795
      TabIndex        =   13
      Top             =   3255
      Width           =   855
      Begin VB.Frame BARRATRAZA 
         BackColor       =   &H00FF0000&
         BorderStyle     =   0  'None
         Height          =   500
         Left            =   0
         TabIndex        =   14
         Top             =   0
         Width           =   12000
      End
   End
   Begin VB.ListBox List1 
      Height          =   540
      Left            =   8610
      Sorted          =   -1  'True
      TabIndex        =   9
      Top             =   105
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Frame Frame7 
      BackColor       =   &H00EEEEC0&
      Caption         =   "Rango de Fechas"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1200
      Left            =   210
      TabIndex        =   8
      Top             =   1890
      Width           =   3585
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
         Height          =   300
         Left            =   3150
         TabIndex        =   19
         Top             =   735
         Width           =   175
      End
      Begin VB.TextBox Text3 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   2100
         TabIndex        =   18
         Text            =   " "
         Top             =   735
         Width           =   1065
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
         Left            =   1470
         TabIndex        =   16
         Top             =   735
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
         Height          =   300
         Left            =   420
         TabIndex        =   15
         Text            =   " "
         Top             =   735
         Width           =   1065
      End
      Begin VB.Label Label2 
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
         Left            =   1470
         TabIndex        =   20
         Top             =   420
         Width           =   1065
      End
      Begin VB.Label Label5 
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
         Left            =   -1470
         TabIndex        =   17
         Top             =   420
         Width           =   2325
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00EEEEC0&
      Caption         =   "Artículo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1590
      Left            =   210
      TabIndex        =   4
      Top             =   210
      Width           =   3585
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
         Left            =   3150
         TabIndex        =   5
         Top             =   460
         Width           =   175
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
         Height          =   300
         Left            =   1260
         TabIndex        =   6
         Top             =   460
         Width           =   2010
      End
      Begin VB.Label Label1 
         BackColor       =   &H00E0E0E0&
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
         Height          =   495
         Left            =   105
         TabIndex        =   10
         Top             =   900
         Width           =   3255
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
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
         Height          =   225
         Left            =   -1155
         TabIndex        =   7
         Top             =   525
         Width           =   2325
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00EEEEC0&
      Caption         =   "Salida"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   915
      Left            =   210
      TabIndex        =   3
      Top             =   3150
      Width           =   3585
      Begin VB.OptionButton Option12 
         BackColor       =   &H00EEEEC0&
         Caption         =   "Impresión"
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
         Left            =   1890
         TabIndex        =   12
         Top             =   420
         Value           =   -1  'True
         Width           =   1170
      End
      Begin VB.OptionButton Option11 
         BackColor       =   &H00EEEEC0&
         Caption         =   "Pantalla"
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
         Left            =   525
         TabIndex        =   11
         Top             =   420
         Width           =   1170
      End
   End
   Begin VB.CommandButton Command11 
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
      Height          =   800
      Left            =   3950
      Picture         =   "COENTRE.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Ayuda FLEXOFT en Línea"
      Top             =   1155
      Width           =   855
   End
   Begin VB.CommandButton Command12 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   585
      Left            =   3950
      Picture         =   "COENTRE.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Emite Reporte Solicitado"
      Top             =   2520
      Width           =   855
   End
   Begin VB.CommandButton command10 
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
      Height          =   800
      Left            =   3950
      Picture         =   "COENTRE.frx":0614
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Abandona la Aplicación"
      Top             =   315
      Width           =   855
   End
   Begin VB.Image Image2 
      Height          =   555
      Left            =   3750
      Picture         =   "COENTRE.frx":075E
      Stretch         =   -1  'True
      Top             =   3500
      Width           =   4725
   End
End
Attribute VB_Name = "COENTRE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
    Call SELECHO("MASTER")
    cod1$ = TRIM$(VALACT1$("MASTER"))
    If cod1$ <> "" Then
       Text1.Text = cod1$
    End If
End Sub

Private Sub Command10_Click()
   Call CIERRARCH("*.*")
   Call FINAL("")
End Sub

Private Sub Command12_Click()
   Command12.Enabled = False
   Call COENTREGA
   Command12.Enabled = True
End Sub

Private Sub Command2_Click()
    Call SELECHO("SELFECHA")
    DESDE$ = TRIM$(VALACT1$("SELFECHA"))
    If DESDE$ <> "" Then
       Text2.Text = DESDE$
    End If
End Sub

Private Sub Command3_Click()
    Call SELECHO("SELFECHA")
    HASTA$ = TRIM$(VALACT1$("SELFECHA"))
    If HASTA$ <> "" Then
       Text3.Text = HASTA$
    End If
End Sub

Private Sub Form_Load()
    '
    Call VERJOBID
    Call TRANSLABELS(Name)
    '
    EPAC$ = TRIM$(EMPREAC$)
    If EPAC$ <> "" Then
        Caption = Caption + "  -  " + EPAC$
    End If
    '
    FEX = HOY(HO$)
    Text3 = FECHATEX$(FEX)
    Text2 = "01" + Mid$(Text3, 3)
    '
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Call CIERRARCH("*.*")
   Call FINAL("")
End Sub

Private Sub Text1_DblClick()
    Call SELECHO("MASTER")
    cod1$ = TRIM$(VALACT1$("MASTER"))
    If cod1$ <> "" Then
       Text1.Text = cod1$
    End If
End Sub

Private Sub Text2_DblClick()
    Call SELECHO("SELFECHA")
    DESDE$ = TRIM$(VALACT1$("SELFECHA"))
    If DESDE$ <> "" Then
       Text2.Text = DESDE$
    End If
End Sub

Private Sub Text3_DBLCLICK()
    Call SELECHO("SELFECHA")
    HASTA$ = TRIM$(VALACT1$("SELFECHA"))
    If HASTA$ <> "" Then
       Text3.Text = HASTA$
    End If
End Sub

Private Sub Text1_Change()
    CI1% = CODINT%(Text1.Text)
    If CI1% > 0 Then
         Label1.Caption = DESCRIT$(CI1%)
       Else
         Label1.Caption = ""
    End If
End Sub

Sub COENTREGA()
   CIXI% = CODINT%(Text1.Text)
   If CIXI% > 0 Then
     If CIXI% <= NUMAS% Then
       DES% = FECHANUM(Text2)
       If DES% > 0 Then
         HAS% = FECHANUM(Text3)
         If HAS% >= DES% Then
           '
           Screen.MousePointer = 11
           FIN& = ULTREG&("MOVISTO")
           J& = 0
           RFF$ = RECFILT$("MOVISTO", 2, MKI$(CIXI%))
           If Len(RFF$) < 4 Then
               Call MENSERR(24, "Artículo sin Movimiento")
             Else
               For UJ& = 1 To Len(RFF$) Step 4
                 U& = CVS(Mid$(RFF$, UJ&, 4))
                 If U& > 0 Then
                   If U& <= FIN& Then
                     X$ = REGLEIDO$("MOVISTO", U&)
                     CLI% = CVI(Mid$(X$, 84, 2))
                     If CLI% > 0 Then
                       If ECOARCH$("DEUDOR1", MKI$(CLI%)) <> "" Then
                         FF% = CVI(Left$(X$, 2))
                         If FF% >= DES% Then
                           If FF% <= HAS% Then
                             REMIFAC$ = Mid$(X$, 23, 20)
                             CANTU# = CVD(Mid$(X$, 96, 8)) - CVD(Mid$(X$, 88, 8))
                             '
                             Z$ = REGBLAN$("FICHENT")
                             Call REPLA(Z$, MKI$(CLI%), 1, 2)
                             Call REPLA(Z$, PERIO$, 3, 20)
                             Call REPLA(Z$, MKI$(FF%), 23, 2)
                             Call REPLA(Z$, MKI$(CIXI%), 25, 2)
                             Call REPLA(Z$, MKD$(CANTU#), 27, 8)
                             Call REPLA(Z$, REMIFAC$, 35, 20)
                             '
                             J& = J& + 1
                             Call GRAREG("FICHENT", Z$, J&)
                             '
                           End If
                         End If
                       End If
                     End If
                   End If
                 End If
               Next UJ&
               '
               Call SETULTREG("FICHENT", J&)
               Call CIERRARCH("FICHENT")
               Screen.MousePointer = 1
               '
               If J& < 1 Then
                 Call MENSERR(24, "Artículo sin Movimiento\en el Período Elegido.")
                 GoTo 99
               End If
               '
               Call HEADERS("FICHENT", "")
               Call HEADERS("FICHENT", "Artículo: " + TRIM$(CODEXT$(CIXI%)) + " - " + DESCRIT0$(CIXI%))
               Call CIERRARCH("*.*")
               '
               If Option11.Value = True Then
                   Call FINAL("*COENTRE")
                 Else
                   Call FINAL("*LIENTRE")
               End If
               '
           End If
         End If
       End If
     End If
   End If

99 Call CIERRARCH("*.*")
   Screen.MousePointer = 1
   '
End Sub
