VERSION 5.00
Begin VB.Form FORCONFABO 
   BackColor       =   &H00C0FFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Configuración de Abonos Mensuales"
   ClientHeight    =   5130
   ClientLeft      =   60
   ClientTop       =   315
   ClientWidth     =   11910
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5130
   ScaleWidth      =   11910
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
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
      Height          =   8200
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   12000
      Begin VB.TextBox Eco21 
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
         Index           =   2
         Left            =   8190
         TabIndex        =   36
         Text            =   " "
         Top             =   1420
         Width           =   2490
      End
      Begin VB.TextBox TEXT21 
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
         Index           =   2
         Left            =   10080
         TabIndex        =   35
         Text            =   " "
         Top             =   1050
         Width           =   435
      End
      Begin VB.CommandButton Command21 
         Caption         =   "."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   2
         Left            =   10500
         TabIndex        =   34
         Top             =   1050
         Width           =   180
      End
      Begin VB.TextBox Eco21 
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
         Index           =   1
         Left            =   8190
         TabIndex        =   32
         Text            =   " "
         Top             =   580
         Width           =   2490
      End
      Begin VB.TextBox TEXT21 
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
         Index           =   1
         Left            =   10080
         TabIndex        =   31
         Text            =   " "
         Top             =   210
         Width           =   435
      End
      Begin VB.CommandButton Command21 
         Caption         =   "."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   1
         Left            =   10500
         TabIndex        =   30
         Top             =   210
         Width           =   180
      End
      Begin VB.PictureBox Picture2 
         BackColor       =   &H00C0C0C0&
         Height          =   3060
         Left            =   7410
         ScaleHeight     =   3000
         ScaleWidth      =   3210
         TabIndex        =   18
         Top             =   1890
         Width           =   3270
         Begin VB.Frame Frame2 
            BackColor       =   &H00E0E0E0&
            BorderStyle     =   0  'None
            Caption         =   "Frame2"
            Height          =   2640
            Left            =   0
            TabIndex        =   19
            Top             =   350
            Width           =   3585
            Begin VB.TextBox Text12 
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
               Left            =   105
               MaxLength       =   4
               TabIndex        =   29
               Top             =   735
               Width           =   540
            End
            Begin VB.TextBox Text10 
               BeginProperty Font 
                  Name            =   "Courier"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Left            =   1470
               TabIndex        =   27
               Text            =   " "
               Top             =   210
               Width           =   1590
            End
            Begin VB.Label Label23 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   1  'Fixed Single
               BeginProperty Font 
                  Name            =   "Courier"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Left            =   1470
               TabIndex        =   28
               Top             =   1260
               Width           =   1590
            End
            Begin VB.Line Line5 
               X1              =   0
               X2              =   3200
               Y1              =   2630
               Y2              =   2630
            End
            Begin VB.Line Line4 
               X1              =   3200
               X2              =   3200
               Y1              =   0
               Y2              =   2625
            End
            Begin VB.Line Line1 
               X1              =   0
               X2              =   3200
               Y1              =   0
               Y2              =   0
            End
            Begin VB.Label Label19 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Tot.Bruto:"
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
               Left            =   -840
               TabIndex        =   25
               Top             =   2205
               Width           =   2220
            End
            Begin VB.Label Label20 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   1  'Fixed Single
               BeginProperty Font 
                  Name            =   "Courier"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Left            =   1470
               TabIndex        =   24
               Top             =   2160
               Width           =   1590
            End
            Begin VB.Label Label18 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   1  'Fixed Single
               BeginProperty Font 
                  Name            =   "Courier"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Left            =   1470
               TabIndex        =   23
               Top             =   735
               Width           =   1590
            End
            Begin VB.Label Label17 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "% IVA:"
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
               Left            =   210
               TabIndex        =   22
               Top             =   840
               Width           =   1170
            End
            Begin VB.Label Label22 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "IVA-RNI:"
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
               Left            =   -840
               TabIndex        =   21
               Top             =   1365
               Width           =   2220
            End
            Begin VB.Label Label13 
               Alignment       =   1  'Right Justify
               BackStyle       =   0  'Transparent
               Caption         =   "Tot,Neto:"
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
               Left            =   -840
               TabIndex        =   20
               Top             =   315
               Width           =   2220
            End
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   $"FORCONFABO.frx":0000
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
            TabIndex        =   26
            Top             =   60
            Width           =   10410
         End
         Begin VB.Line Line2 
            X1              =   3200
            X2              =   3200
            Y1              =   0
            Y2              =   2940
         End
      End
      Begin VB.PictureBox Picture1 
         BackColor       =   &H00C0C0C0&
         Height          =   3060
         Left            =   105
         ScaleHeight     =   3000
         ScaleWidth      =   7095
         TabIndex        =   15
         Top             =   1890
         Width           =   7155
         Begin VB.TextBox Text8 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2640
            Left            =   -30
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   17
            Top             =   330
            Width           =   7140
         End
         Begin VB.Line Line6 
            X1              =   0
            X2              =   7140
            Y1              =   2970
            Y2              =   2970
         End
         Begin VB.Label Label12 
            BackStyle       =   0  'Transparent
            Caption         =   $"FORCONFABO.frx":008A
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
            TabIndex        =   16
            Top             =   60
            Width           =   10410
         End
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
         Left            =   5250
         TabIndex        =   14
         TabStop         =   0   'False
         Text            =   " "
         Top             =   1155
         Width           =   2640
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
         Left            =   6405
         MaxLength       =   1
         TabIndex        =   13
         TabStop         =   0   'False
         Text            =   "X"
         ToolTipText     =   "Configuracion de Funcionamiento - Doble-Click Aquí"
         Top             =   315
         Width           =   500
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
         Height          =   780
         Left            =   10920
         Picture         =   "FORCONFABO.frx":0114
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "Apertura de Nueva Cuenta de Cliente"
         Top             =   1890
         Width           =   855
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Notas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   780
         Left            =   10920
         Picture         =   "FORCONFABO.frx":0556
         Style           =   1  'Graphical
         TabIndex        =   11
         ToolTipText     =   "Notas de Clientes"
         Top             =   2730
         Width           =   855
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
         Height          =   855
         Left            =   10920
         Picture         =   "FORCONFABO.frx":0860
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Abandona la Aplicación"
         Top             =   210
         Width           =   855
      End
      Begin VB.CommandButton BOTNEXT 
         Caption         =   "O.K."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1005
         Left            =   10920
         Picture         =   "FORCONFABO.frx":0B6A
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Aprueba - Registra - Imprime Nota de Débito"
         Top             =   3900
         Width           =   855
      End
      Begin VB.TextBox Text5 
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
         Left            =   2310
         TabIndex        =   8
         TabStop         =   0   'False
         Text            =   " "
         Top             =   1155
         Width           =   2745
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
         Height          =   300
         Left            =   1470
         TabIndex        =   7
         TabStop         =   0   'False
         Text            =   " "
         Top             =   1155
         Width           =   750
      End
      Begin VB.TextBox Text3 
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
         Left            =   1470
         TabIndex        =   6
         TabStop         =   0   'False
         Text            =   " "
         Top             =   735
         Width           =   3585
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
         Left            =   1470
         TabIndex        =   5
         TabStop         =   0   'False
         Text            =   " "
         Top             =   315
         Width           =   3585
      End
      Begin VB.CommandButton Command1 
         Caption         =   "."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1050
         TabIndex        =   1
         Top             =   1155
         Width           =   180
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
         Height          =   300
         Left            =   210
         TabIndex        =   0
         Text            =   " "
         Top             =   1155
         Width           =   855
      End
      Begin VB.Label Label21 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Vended:"
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
         Index           =   2
         Left            =   9030
         TabIndex        =   37
         Top             =   1050
         Width           =   960
      End
      Begin VB.Label Label21 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "C.Pago:"
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
         Index           =   1
         Left            =   9030
         TabIndex        =   33
         Top             =   210
         Width           =   960
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
         TabIndex        =   4
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
      TabIndex        =   9
      Top             =   0
      Width           =   960
   End
End
Attribute VB_Name = "FORCONFABO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim LIST2ANT$

Private Sub BOTNEXT_Click()
   '
   OKEY% = 0
   NC% = Val(Text1.TEXT)
   If ECOARCH$("DEUDOR1", MKI$(NC%)) = "" Then
      Call MENSERR(24, "Numero de Cuenta no Válido")
      Text1.SetFocus
      Exit Sub
   End If
   '
   CPAG% = Val(TEXT21(1).TEXT)
   If ECOARCH$("CONPAG", MKI$(CPAG%)) = "" Then
      Call MENSERR(24, "Condición de Pago no Válida")
      TEXT21(1).SetFocus
      Exit Sub
   End If
   '
   VEND% = Val(TEXT21(2).TEXT)
   '
   If TRIM$(Text8.TEXT) = "" Then
      Call MENSERR(24, "Falta Leyenda")
      Text8.SetFocus
      Exit Sub
   End If
   '
   ITOTAB# = Val(Text10.TEXT)
   RFF$ = RECFILT$("ABONOS", 1, MKI$(NC%))
   TTXX$ = String$(1024, 0)
   Call REPLA(TTXX$, MKI$(NC%), 1, 2)
   Call REPLA(TTXX$, Chr$(CPAG%), 3, 1)
   Call REPLA(TTXX$, Chr$(VEND%), 4, 1)
   Call REPLA(TTXX$, MKD$(ITOTAB#), 5, 8)
   Call REPLA(TTXX$, Text8.TEXT, 25, 1000)
   REABO& = 0
   If Len(RFF$) >= 4 Then
      REABO& = CVS(Left$(RFF$, 4))
   End If
   If REABO& < 1 Or REABO& > ULTREG&("ABONOS") Then
      REABO& = 1 + ULTREG&("ABONOS")
   End If
   '
   Call GRAREG("ABONOS", TTXX$, REABO&)
   Call CIERRARCH("ABONOS")
   '
   Call CLEARCONFABO
   Text1.SetFocus
   '
End Sub

Private Sub Command1_Click()
    '
    Call SELECHO("DEUDOR1")
    NCXI% = Val(TRIM$(VALACT1$("DEUDOR1")))
    If NCXI% > 0 Then
        Text1.TEXT = TRIM$(Str$(NCXI%))
        NC% = NCXI%
        OKEY% = 0
        Call CONFABO1(OKEY%)
        If OKEY% <> 1 Then
           Call CLEARCONFABO
           Exit Sub
        End If
    End If
    '
End Sub

Private Sub Command10_Click()
    Call SELECHO("ECODEP")
    DPX1% = Val(VALACT1$("ECODEP"))
    If DPX1% > 0 Then
        DP1% = DPX1%
        Text10.TEXT = TRIM$(Str$(DP1%))
        Text7.TEXT = ECOARCH$("ECODEP", Chr$(DP1%))
    End If
End Sub

Private Sub Command11_Click()
    Command11.Enabled = False
    Call CIERRARCH("*.*")
    Call CONECRUN("AMACLI/NEW", "Alta de Nueva Cuenta de Cliente")
    Call FRESHECHO("DEUDOR1")
    Command11.Enabled = True
End Sub

Private Sub Command12_Click()
   Call MENSERR(24, "Opción Inserción no Disponible.\Debe Agregarse al Final de la Lista.")
End Sub

Private Sub Command2_Click()
    If Text8.TEXT <> "" Then
          Text8.TEXT = ""
          Text1.TEXT = ""
          Text2.TEXT = ""
          Text3.TEXT = ""
          Text4.TEXT = ""
          Text5.TEXT = ""
          '
          Text10.TEXT = ""
          Label18.Caption = ""
          Label20.Caption = ""
          Label23.Caption = ""
          '
          Text1.SetFocus
        Else
          OPFACDIP.Option1.Value = True
          '
          VACON% = (-1)
          Unload FORCONFABO
          FORFACDIP.Show
          OPFACDIP.Show 1
    End If
End Sub

Private Sub Command3_Click()
    Command3.Enabled = False
    NCLI$ = TRIM$(Text1.TEXT)
    Call CONECRUN("*NOTACLI/" + NCLI$, "Notas de Clientes")
    Command3.Enabled = True
End Sub

Private Sub Command4_Click()
    Command4.Enabled = False
    RESCUDEU.Text1.TEXT = Text1.TEXT
    Hide
    PROLLAMA$ = "FORCONFABO"
    RESCUDEU.Show
    Command4.Enabled = True
End Sub

Private Sub Command9_Click()
    Call SELECHO("SELFECHA")
    VDEVV$ = VALACT1$("SELFECHA")
    If VDEVV$ <> "" Then
        Text9.TEXT = VDEVV$
        FF% = FECHANUM(Text9.TEXT)
        FEFAC% = FF%
    End If
End Sub

Private Sub Form_Load()
    '
    EPAC$ = TRIM$(EMPREAC$)
    If EPAC$ <> "" Then
        Caption = Caption + "  -  " + EPAC$
    End If
    '
    Refresh
    VACON% = 0
    '
End Sub

Private Sub Label2_DblClick()
    OPREMIFAC.Show 1
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Cancel = 1
   If VACON% = (-1) Then
      Cancel = 0
   End If
End Sub

Private Sub Text1_Change()
    Text8.TEXT = ""
    Text10.TEXT = ""
    Label18.Caption = ""
    Label20.Caption = ""
    Label23.Caption = ""
    BOTNEXT.Enabled = True
End Sub

Private Sub Text1_DBLCLICK()
    '
    Call SELECHO("DEUDOR1")
    NCXI% = Val(TRIM$(VALACT1$("DEUDOR1")))
    If NCXI% > 0 Then
        Text1.TEXT = TRIM$(Str$(NCXI%))
        NC% = NCXI%
        OKEY% = 0
        Call CONFABO1(OKEY%)
        If OKEY% <> 1 Then
           Call CLEARCONFABO
           Exit Sub
        End If
    End If
    '
End Sub

Private Sub Text1_GotFocus()
    Static CTXX%
    If CTXX% = 0 Then
        Refresh
        CTXX% = 1
    End If
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       OKEY% = 0
       Call CONFABO1(OKEY%)
       If OKEY% <> 1 Then
           Call CLEARCONFABO
           Exit Sub
       End If
    End If
    '
End Sub

Private Sub Text10_Change()
    '
   ALIVABA = Val(Text12.TEXT)
   TOPEDIDO# = Val(Text10.TEXT)
   If PIVAFAC% <> 5 Then Label18.Caption = CSTRING$(MKD$(TOPEDIDO# * ALIVABA / 100), 3, 10, 2, 2)
   If RENOIN% = 1 Then
      Label23.Caption = CSTRING$(MKD$(TOPEDIDO# * (ALIVABA / ALIVA) * ALIVARNI / 100), 3, 10, 2, 2)
   End If
   Label20.Caption = CSTRING$(MKD$(TOPEDIDO# + Val(Label18.Caption) + Val(Label23.Caption)), 3, 10, 2, 2)
   '
End Sub

Private Sub Text10_KeyPress(KeyAscii As Integer)
   If KeyAscii = 13 Then
      TOPEDIDO# = Val(Text10.TEXT)
      Text10.TEXT = CSTRING$(MKD$(TOPEDIDO#), 3, 10, 2, 2)
   End If
End Sub

Private Sub Text10_LostFocus()
    TOPEDIDO# = Val(Text10.TEXT)
    Text10.TEXT = CSTRING$(MKD$(TOPEDIDO#), 3, 10, 2, 2)
End Sub

Private Sub Text11_CLICK()
   Text1.SetFocus
End Sub

Private Sub Text12_Change()
   ALIVABA = Val(Text12.TEXT)
   TOPEDIDO# = Val(Text10.TEXT)
   If PIVAFAC% <> 5 Then Label18.Caption = CSTRING$(MKD$(TOPEDIDO# * ALIVABA / 100), 3, 10, 2, 2)
   If RENOIN% = 1 Then
      Label23.Caption = CSTRING$(MKD$(TOPEDIDO# * (ALIVABA / ALIVA) * ALIVARNI / 100), 3, 10, 2, 2)
   End If
   Label20.Caption = CSTRING$(MKD$(TOPEDIDO# + Val(Label18.Caption) + Val(Label23.Caption)), 3, 10, 2, 2)
End Sub

Private Sub Text2_GotFocus()
    Text1.SetFocus
End Sub

Private Sub Text3_GotFocus()
    Text1.SetFocus
End Sub

Private Sub Text4_GotFocus()
    Text1.SetFocus
End Sub

Private Sub Text5_GotFocus()
    Text1.SetFocus
End Sub

Private Sub Text6_Click()
    Text1.SetFocus
End Sub

Private Sub Text6_DblClick()
    OPFACTUR.Show 1
End Sub

Private Sub Text7_GotFocus()
    Text1.SetFocus
End Sub

Private Sub Text9_DblClick()
    Call SELECHO("SELFECHA")
    VDEVV$ = VALACT1$("SELFECHA")
    If VDEVV$ <> "" Then
        Text9.TEXT = VDEVV$
        FEFAC% = FECHANUM(Text9.TEXT)
    End If
End Sub

Private Sub Text9_LostFocus()
    VDEVV$ = TRIM$(Text9.TEXT)
    If FECHANUM(VDEVV$) > 0 Then
          FEFAC% = FECHANUM(VDEVV$)
        Else
          Call MENSERR(24, "Fecha no Válida")
          FEX = HOY(HO$)
          FEFAC% = FEX
          Text9.TEXT = HO$
          Text9.SetFocus
    End If
End Sub

Private Sub TEXT21_Change(Index As Integer)
    '
    If Index = 1 Then
        Eco21(Index).TEXT = ECOARCH$("CONPAG", MKI$(Val(TEXT21(Index).TEXT)))
      ElseIf Index = 2 Then
        Eco21(Index).TEXT = ECOARCH$("VENDEDOR", Chr$(Val(TEXT21(Index).TEXT)))
    End If
    '
End Sub

Private Sub Command21_Click(Index As Integer)
    '
    If Index = 1 Then
        ARCHEC$ = "CONPAG": LOCA% = 2
      ElseIf Index = 2 Then
        ARCHEC$ = "VENDEDOR": LOCA% = 1
    End If
    ARGU$ = CBIN$(TRIM$(TEXT21(Index).TEXT), 1, LOCA%)
    Call SELECHO(ARCHEC$)
    TEXT21(Index).TEXT = VALACT1$(ARCHEC$)
    '
End Sub


